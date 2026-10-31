<#
.SYNOPSIS
    Get-DiskUsage の回帰テストです。

.DESCRIPTION
    一時フォルダにダミー環境（日本語のパス、260 文字を超えるパス、ジャンクションのループ、
    読み取り禁止のフォルダ、古いファイル、文字コードや書き方の違う設定ファイルなど）を作り、
    Get-DiskUsage を実行して期待値と照合します。ダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダと設定ファイルにだけアクセス拒否を設定します（終了時に戻します）。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Get-DiskUsage を実行する PowerShell（powershell.exe または pwsh）。

.PARAMETER KeepTemp
    ダミー環境と結果を削除せずに残します（調査用）。

.NOTES
    Copyright (c) 2026 結進舎
    Released under the MIT License.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Exe,

    [switch]$KeepTemp
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Continue'

$tool     = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Get-DiskUsage\Get-DiskUsage.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Get-DiskUsage_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tree     = Join-Path $workRoot 'tree'
$script:failCount = 0

#region 補助関数

function Assert-Equal([string]$Name, $Actual, $Expected) {
    if ("$Actual" -ceq "$Expected") {
        '  PASS  {0} = {1}' -f $Name, $Actual
    } else {
        '  FAIL  {0} : 実際={1} 期待={2}' -f $Name, $Actual, $Expected
        $script:failCount++
    }
}

function Write-Skip([string]$Name, [string]$Reason) { '  SKIP  {0}（{1}）' -f $Name, $Reason }

function Write-Section([string]$Title) { ''; ('■ ' + $Title) }

# CSV の数値は 3 桁区切りで出力されるため、カンマを外して数値にする
function ConvertTo-Long($Text) { [long]("$Text" -replace ',', '') }

function Get-ByteSum($Rows) {
    $sum = 0L
    foreach ($row in $Rows) { $sum += ConvertTo-Long $row.'サイズ(バイト)' }
    $sum
}

# 260 文字を超えるパスも作れるように、\\?\ を付けて作る
function New-SizedFile([string]$FilePath, [long]$Size, [int]$DaysAgo = 0) {
    $parent = $FilePath.Substring(0, $FilePath.LastIndexOf('\'))
    [void][IO.Directory]::CreateDirectory('\\?\' + $parent)
    $stream = [IO.File]::Create('\\?\' + $FilePath)
    $stream.SetLength($Size)
    $stream.Close()
    if ($DaysAgo -gt 0) { [IO.File]::SetLastWriteTime('\\?\' + $FilePath, (Get-Date).AddDays(-$DaysAgo)) }
}

# コンソールで実行すると、長い行は画面の幅で折り返され、出力に改行が入る。
# 画面の幅によって結果が変わらないように、文字を探す前に改行を取り除く
function ConvertTo-FlatText([string]$Text) { $Text -replace '\r?\n', '' }

# Log には、画面への出力から改行を取り除いたものを入れる
function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out  = Join-Path $workRoot "out_$Label"
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $dir  = $null
    if (Test-Path -LiteralPath $out) {
        $found = @(Get-ChildItem -LiteralPath $out -Directory)
        if ($found.Count -gt 0) { $dir = $found[0].FullName }
    }
    [pscustomobject]@{ Code = $code; Log = $log; Dir = $dir }
}

# 結果のファイルがないとき（ツールが保存に失敗したときなど）も止まらないように、空として読む
function Read-Result([string]$Dir) {
    function Read-CsvOrEmpty([string]$Name) {
        if (-not $Dir) { return @() }
        $file = Join-Path $Dir $Name
        if (-not (Test-Path -LiteralPath $file)) { return @() }
        @(Import-Csv -LiteralPath $file -Encoding UTF8)
    }
    $summary = ''
    if ($Dir -and (Test-Path -LiteralPath (Join-Path $Dir 'summary.txt'))) { $summary = [IO.File]::ReadAllText((Join-Path $Dir 'summary.txt')) }
    [pscustomobject]@{
        Folders = @(Read-CsvOrEmpty 'folders.csv')
        Files   = @(Read-CsvOrEmpty 'files.csv')
        Exts    = @(Read-CsvOrEmpty 'extensions.csv')
        Skipped = @(Read-CsvOrEmpty 'skipped.csv')
        Summary = $summary
    }
}

# 5.1 と 7 で並び順が変わらないように、文字コード順に並べてカンマでつなぐ
function Join-Sorted($Items) {
    $array = [string[]]@($Items)
    [Array]::Sort($array, [StringComparer]::Ordinal)
    $array -join ','
}

# 保存されたファイルの名前の一覧
function Get-ResultFileNames([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { $_.Name })
}

function Get-LineCount([string]$Dir, [string]$FileName) {
    if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir $FileName))) { return -1 }
    @(Get-Content -LiteralPath (Join-Path $Dir $FileName) -Encoding UTF8).Count
}

# 除外指定でスキップされたフォルダの、Root から見た相対パスの一覧
function Get-ExcludedNames($Result, [string]$Root) {
    Join-Sorted (@($Result.Skipped | Where-Object { $_.'理由' -eq '除外指定' }) | ForEach-Object { $_.'パス'.Substring($Root.Length + 1) })
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $log = ConvertTo-FlatText (& cmd.exe /d /c $batch 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# ダミー環境の既定の結果（オプションなし）を確かめる。ローカルと共有フォルダ経由で共通
function Test-DefaultResult($Run, [string]$Root) {
    Assert-Equal '終了コード' $Run.Code 0
    $x = Read-Result $Run.Dir
    Assert-Equal 'ファイル数' $x.Files.Count 9
    Assert-Equal 'ファイルの合計バイト' (Get-ByteSum $x.Files) 6317234
    Assert-Equal '拡張子の合計バイト' (Get-ByteSum $x.Exts) 6317234
    Assert-Equal '拡張子の種類' $x.Exts.Count 7
    $log = @($x.Exts | Where-Object { $_.'拡張子' -eq '.log' })
    Assert-Equal '.log（大文字の .LOG も合算）' ('{0}/{1}' -f ($log | ForEach-Object { $_.'サイズ(バイト)' }), ($log | ForEach-Object { $_.'ファイル数' })) '3,500,000/2'
    Assert-Equal '拡張子なし' (@($x.Exts | Where-Object { $_.'拡張子' -eq '（拡張子なし）' }) | ForEach-Object { $_.'サイズ(バイト)' }) '5,000'
    Assert-Equal 'a.log の MB（小数 2 桁）' (@($x.Files | Where-Object { $_.'ファイル' -like '*\a.log' }) | ForEach-Object { $_.'サイズ(MB)' }) '2.86'
    Assert-Equal 'フォルダの行数（3 階層まで）' $x.Folders.Count 10
    $jp = @($x.Folders | Where-Object { $_.'フォルダ' -eq "$Root\日本語 フォルダ" })
    Assert-Equal '日本語のフォルダ' ('{0}/{1}' -f ($jp | ForEach-Object { $_.'サイズ(バイト)' }), ($jp | ForEach-Object { $_.'ファイル数' })) '2,500,000/2'
    Assert-Equal 'lvl3 は載る' @($x.Folders | Where-Object { $_.'フォルダ' -like '*\lvl3' }).Count 1
    Assert-Equal 'lvl4 は載らない' @($x.Folders | Where-Object { $_.'フォルダ' -like '*\lvl4' }).Count 0
    $longDir = "$Root\" + ('L' * 120) + '\' + ('M' * 120)
    $lp = @($x.Folders | Where-Object { $_.'フォルダ' -eq $longDir })
    Assert-Equal '260 文字を超えるフォルダ' ($lp | ForEach-Object { $_.'サイズ(バイト)' }) '1,234'
    $allText = ($x.Folders + $x.Files + $x.Skipped | ForEach-Object { ($_.PSObject.Properties | ForEach-Object { $_.Value }) -join '|' }) -join '|'
    Assert-Equal '\\?\ が残っていない' $allText.Contains('\\?\') 'False'
    Assert-Equal '1 位のファイル' (@($x.Files) | Select-Object -First 1 | ForEach-Object { $_.'ファイル' }) "$Root\a.log"
    Assert-Equal 'deep.LOG の拡張子' (@($x.Files | Where-Object { $_.'ファイル' -like '*\deep.LOG' }) | ForEach-Object { $_.'拡張子' }) '.log'
    $link = @($x.Skipped | Where-Object { $_.'理由' -eq 'リンク' })
    Assert-Equal 'リンク（ループするジャンクション）' ('{0}/{1}' -f ($link | ForEach-Object { $_.'パス' }), ($link | ForEach-Object { $_.'詳細' })) "$Root\jloop/ジャンクション"
    if ($script:denyApplied) {
        $denied = @($x.Skipped | Where-Object { $_.'理由' -eq '権限なし' })
        Assert-Equal '権限なし' ('{0}/{1}' -f ($denied | ForEach-Object { $_.'パス' }), ($denied | ForEach-Object { $_.'詳細' })) "$Root\denied/アクセスが拒否されました"
        Assert-Equal 'summary のスキップ行' ([regex]::Match($x.Summary, 'スキップ.*').Value.Trim()) 'スキップ 1 件（権限なし）  たどらなかったリンク 1 件'
        Assert-Equal 'summary の件数' ([regex]::Match($x.Summary, 'ファイル \S+ 件  フォルダ \S+ 件').Value) 'ファイル 9 件  フォルダ 11 件'
    } else {
        Write-Skip '権限なし' 'アクセス拒否を設定できなかったため'
    }
}

# 設定ファイルを書く。文字コードの判定を確かめるため、文字コードを選べるようにする
function New-ConfigFile([string]$FilePath, [string[]]$Lines, [Text.Encoding]$Encoding = (New-Object Text.UTF8Encoding $false)) {
    $parent = $FilePath.Substring(0, $FilePath.LastIndexOf('\'))
    [void][IO.Directory]::CreateDirectory('\\?\' + $parent)
    [IO.File]::WriteAllText('\\?\' + $FilePath, ($Lines -join "`r`n"), $Encoding)
    $FilePath
}

# 設定ファイルの OutputFolder や相対パスの基準を確かめるため、Invoke-Tool と違って -OutputFolder を足さず、
# 指定したフォルダから起動する
function Invoke-From([string]$WorkDir, [string[]]$ToolArgs) {
    Push-Location -LiteralPath $WorkDir
    try {
        $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1 | Out-String)
        $code = $LASTEXITCODE
    } finally {
        Pop-Location
    }
    [pscustomobject]@{ Code = $code; Log = $log }
}

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'DiskUsage_*' -ErrorAction SilentlyContinue)
    if ($found.Count -gt 0) { return $found[0].FullName }
    $null
}

# 画面の「対象」の行に出たフォルダ
function Get-TargetText([string]$Log) { [regex]::Match($Log, '対象: (.+?)   \d{4}-').Groups[1].Value }

function Get-FileCountText([string]$Log) { [regex]::Match($Log, 'ファイル \S+ 件').Value }

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$script:denyApplied = $false
$deniedDir = Join-Path $tree 'denied'
$userSid   = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$cfgDir    = Join-Path $workRoot 'cfg\設定 フォルダ'
$cfgLocked = Join-Path $cfgDir 'locked.ini'
$script:cfgDenyApplied = $false

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    New-SizedFile "$tree\a.log" 3000000 200
    New-SizedFile "$tree\b.txt" 1000
    New-SizedFile "$tree\noext" 5000
    New-SizedFile "$tree\日本語 フォルダ\データ.bak" 2000000 150
    New-SizedFile "$tree\日本語 フォルダ\sub\deep.LOG" 500000
    New-SizedFile "$tree\node_modules\x.js" 700000
    New-SizedFile "$tree\pkg\node_modules\y.js" 100000
    New-SizedFile "$deniedDir\secret.dat" 400000
    New-SizedFile ("$tree\" + ('L' * 120) + '\' + ('M' * 120) + '\deep.bin') 1234
    New-SizedFile "$tree\lvl1\lvl2\lvl3\lvl4\f.dat" 10000
    $null = New-Item -ItemType Junction -Path "$tree\jloop" -Target $tree
    # 自分自身に「フォルダの一覧表示」を拒否する
    $null = & icacls.exe $deniedDir /deny "*${userSid}:(RD)" 2>&1
    $script:denyApplied = ($LASTEXITCODE -eq 0)
    Assert-Equal 'アクセス拒否の設定' $script:denyApplied 'True'

    Write-Section '既定（オプションなし）'
    Test-DefaultResult (Invoke-Tool 'default' @('-Path', $tree)) $tree

    Write-Section '-OlderThan 100'
    $run = Invoke-Tool 'older' @('-Path', $tree, '-OlderThan', '100')
    Assert-Equal '終了コード' $run.Code 0
    $x = Read-Result $run.Dir
    Assert-Equal 'ファイル' (($x.Files | ForEach-Object { Split-Path $_.'ファイル' -Leaf }) -join ',') 'a.log,データ.bak'
    Assert-Equal 'ファイルの合計バイト' (Get-ByteSum $x.Files) 5000000
    Assert-Equal '拡張子の種類' $x.Exts.Count 2
    $jp = @($x.Folders | Where-Object { $_.'フォルダ' -eq "$tree\日本語 フォルダ" })
    Assert-Equal '日本語のフォルダ（古いファイルだけ）' ('{0}/{1}' -f ($jp | ForEach-Object { $_.'サイズ(バイト)' }), ($jp | ForEach-Object { $_.'ファイル数' })) '2,000,000/1'
    Assert-Equal 'summary の条件行' ([regex]::IsMatch($x.Summary, '条件: 最終更新が 100 日以上前のファイルのみ')) 'True'

    Write-Section '-Exclude（カンマ区切りの 1 つの文字列・大文字小文字の違い）'
    $run = Invoke-Tool 'exclude' @('-Path', $tree, '-Exclude', '*\node_modules,*\LVL2')
    Assert-Equal '終了コード' $run.Code 0
    $x = Read-Result $run.Dir
    Assert-Equal 'ファイル数' $x.Files.Count 6
    Assert-Equal 'ファイルの合計バイト' (Get-ByteSum $x.Files) 5507234
    Assert-Equal '除外指定の件数' @($x.Skipped | Where-Object { $_.'理由' -eq '除外指定' }).Count 3
    Assert-Equal 'summary の除外行' ([regex]::IsMatch($x.Summary, '除外したフォルダ 3 件')) 'True'

    Write-Section 'コマンドプロンプトから起動'
    # -Exclude のカンマ区切りと、最後の引数が "パス\"（\" が引用符のまま届く）の場合
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -NoCsv -Exclude "*\node_modules,*\lvl2" -Path "{2}\"' -f $Exe, $tool, $tree)
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '件数' ([regex]::Match($run.Log, 'ファイル \S+ 件').Value) 'ファイル 6 件'
    Assert-Equal '除外' ([regex]::IsMatch($run.Log, '除外したフォルダ 3 件')) 'True'
    # "パス\" の後ろに引数が続くと、後ろの引数までパスとして届く。ヒントを出して終了する
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -Path "{2}\" -NoCsv' -f $Exe, $tool, $tree)
    Assert-Equal '"パス\" の後ろに引数: 終了コード' $run.Code 1
    Assert-Equal '"パス\" の後ろに引数: ヒント' ([regex]::IsMatch($run.Log, '最後に \\ を付けないでください')) 'True'

    Write-Section '異常系'
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path (Join-Path $workRoot 'no\such\folder') 2>&1 | Out-String)
    Assert-Equal '存在しないパス: 終了コード' $LASTEXITCODE 1
    Assert-Equal '存在しないパス: メッセージ' ([regex]::IsMatch($log, '指定したフォルダが見つかりません')) 'True'
    $usedLetters = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    $freeLetter  = @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $usedLetters -notcontains $_ } | Select-Object -First 1
    if ($freeLetter) {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $tree -OutputFolder "${freeLetter}:\nope" 2>&1 | Out-String)
        Assert-Equal '書き込めない保存先: 終了コード' $LASTEXITCODE 0
        Assert-Equal '書き込めない保存先: 警告' ([regex]::IsMatch($log, '保存先に書き込めないため')) 'True'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }
    $noCsvOut = Join-Path $workRoot 'out_nocsv'
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $tree -NoCsv -OutputFolder $noCsvOut 2>&1
    Assert-Equal '-NoCsv: 終了コード' $LASTEXITCODE 0
    Assert-Equal '-NoCsv: 保存しない' (Test-Path -LiteralPath $noCsvOut) 'False'

    Write-Section 'Ctrl+C で中断（C:\ を調べている途中で停止）'
    # Ctrl+C と同じ「パイプラインの停止」を、対象の PowerShell の中から送る
    $stopScript = Join-Path $workRoot 'stop-during-scan.ps1'
    $stopOut    = Join-Path $workRoot 'out_ctrlc'
    [IO.File]::WriteAllLines($stopScript, [string[]]@(
        'param($Tool, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $env:SystemDrive + ''\'').AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'Start-Sleep -Seconds 3'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + (($ps.Streams.Information | ForEach-Object { $_.MessageData.ToString() }) -join '' / '')'
    ), (New-Object Text.UTF8Encoding $true))
    $result = @{}
    & $Exe -NoProfile -ExecutionPolicy Bypass -File $stopScript -Tool $tool -OutRoot $stopOut 2>&1 | ForEach-Object {
        $line = "$_"
        $eq = $line.IndexOf('=')
        if ($eq -gt 0) { $result[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
    }
    if ($result['RUNNING'] -eq 'True') {
        Assert-Equal '止める前は保存先を作ってある' $result['BEFORE'] '1'
        Assert-Equal '止めた後は保存先が残らない' $result['AFTER'] '0'
        Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' '3 秒以内に調べ終わったため、途中で止められなかった'
    }

    Write-Section 'レビューで見つかった問題の再発防止'
    $extra    = Join-Path $workRoot 'extra'
    $allFiles = 'extensions.csv,files.csv,folders.csv,skipped.csv,summary.txt'
    New-SizedFile "$extra\flat\a.log" 100
    foreach ($name in '[a]', 'a', "back``tick", 'a.b', 'aXb') { New-SizedFile "$extra\br\$name\f.bin" 10 }
    foreach ($name in 'app.log.1', 'app.log.10', 'part.001', 'x.log') { New-SizedFile "$extra\ext\$name" 5 }
    New-SizedFile ("$extra\ext\s." + [char]0xD800) 7

    # 行が 0 件のとき、CSV の保存が途中で止まっていた
    $run = Invoke-Tool 'r-flat' @('-Path', "$extra\flat")
    Assert-Equal 'サブフォルダなし: 終了コード' $run.Code 0
    Assert-Equal 'サブフォルダなし: 5 つとも保存' (Get-ResultFileNames $run.Dir) $allFiles
    Assert-Equal 'サブフォルダなし: folders.csv は見出しだけ' (Get-LineCount $run.Dir 'folders.csv') 1
    $run = Invoke-Tool 'r-older0' @('-Path', "$extra\flat", '-OlderThan', '90')
    Assert-Equal '-OlderThan で 0 件: 5 つとも保存' (Get-ResultFileNames $run.Dir) $allFiles
    Assert-Equal '-OlderThan で 0 件: files.csv は見出しだけ' (Get-LineCount $run.Dir 'files.csv') 1

    # -Exclude の [ ] やバッククォートが、ワイルドカードや 5.1 のエスケープとして解釈されていた
    $cases = @(@('*\[a]', '[a]'), @('*\[a', ''), @("*\BACK``TICK", "back``tick"), @('*\a.b', 'a.b'))
    for ($i = 0; $i -lt $cases.Count; $i++) {
        $run = Invoke-Tool "r-exclude$i" @('-Path', "$extra\br", '-Exclude', $cases[$i][0])
        Assert-Equal ('-Exclude {0}: 終了コード' -f $cases[$i][0]) $run.Code 0
        Assert-Equal ('-Exclude {0}: 除外されたフォルダ' -f $cases[$i][0]) (Get-ExcludedNames (Read-Result $run.Dir) "$extra\br") $cases[$i][1]
    }

    # 数字だけの拡張子が Excel で数値に変わっていた。壊れた文字の拡張子があると、CSV がすべて消えていた
    $run = Invoke-Tool 'r-ext' @('-Path', "$extra\ext")
    Assert-Equal '拡張子: 終了コード' $run.Code 0
    Assert-Equal '拡張子: 壊れた文字があっても 5 つとも保存' (Get-ResultFileNames $run.Dir) $allFiles
    $x = Read-Result $run.Dir
    foreach ($ext in '=".1"', '=".10"', '=".001"', '.log') {
        Assert-Equal ('拡張子: extensions.csv に {0}' -f $ext) @($x.Exts | Where-Object { $_.'拡張子' -eq $ext }).Count 1
    }
    Assert-Equal '拡張子: files.csv の app.log.1' (@($x.Files | Where-Object { $_.'ファイル' -like '*\app.log.1' }) | ForEach-Object { $_.'拡張子' }) '=".1"'

    # 260 文字を超える -Path、\\?\ 付きの -Path、末尾にピリオドが付いた -Path を扱えなかった
    $longDir = "$tree\" + ('L' * 120) + '\' + ('M' * 120)
    $run = Invoke-Tool 'r-long' @('-Path', $longDir, '-NoCsv')
    Assert-Equal '260 文字を超える -Path' ('{0}/{1}' -f $run.Code, [regex]::Match($run.Log, 'ファイル \S+ 件').Value) '0/ファイル 1 件'
    $run = Invoke-Tool 'r-prefix' @('-Path', ('\\?\' + $tree), '-NoCsv')
    Assert-Equal '\\?\ 付きの -Path' ('{0}/{1}' -f $run.Code, [regex]::Match($run.Log, 'ファイル \S+ 件').Value) '0/ファイル 9 件'
    Assert-Equal '\\?\ 付きの -Path: 表示に \\?\ が残らない' $run.Log.Contains('\\?\') 'False'
    $run = Invoke-Tool 'r-period' @('-Path', "$extra\flat.", '-NoCsv')
    Assert-Equal '末尾にピリオドが付いた -Path' ('{0}/{1}' -f $run.Code, [regex]::Match($run.Log, 'ファイル \S+ 件').Value) '0/ファイル 1 件'

    # コマンドプロンプトで -Exclude / -OutputFolder の最後が \" だと、別のフォルダを黙って調べていた
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -NoCsv -Exclude "{2}\a\" -Path "{2}"' -f $Exe, $tool, "$extra\br")
    Assert-Equal 'cmd の -Exclude "…\" の後ろに引数' ('{0}/{1}' -f $run.Code, [regex]::IsMatch($run.Log, '最後に \\ を付けないでください')) '1/True'
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -OutputFolder "{2}\" -NoCsv -Path "{3}"' -f $Exe, $tool, (Join-Path $workRoot 'out_q'), "$extra\br")
    Assert-Equal 'cmd の -OutputFolder "…\" の後ろに引数' ('{0}/{1}' -f $run.Code, [regex]::IsMatch($run.Log, '最後に \\ を付けないでください')) '1/True'
    # コマンドプロンプトで、カンマ区切りの全体を ' で囲むと効かなかった
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -NoCsv -Path "{2}" -Exclude ''*\[a],*\aXb''' -f $Exe, $tool, "$extra\br")
    Assert-Equal 'cmd の ''…,…''' ([regex]::Match($run.Log, '除外したフォルダ \S+ 件').Value) '除外したフォルダ 2 件'

    # 保存先を作れないときの警告に、改行が混ざっていた。
    # 画面に出た文字は、コンソールの幅で折り返されて改行が入るため使わない。対象の PowerShell の中でツールを動かし、
    # 警告のメッセージそのものをファイルに書き出して確かめる
    $warnScript = Join-Path $workRoot 'collect-warnings.ps1'
    $warnFile   = Join-Path $workRoot 'warnings.txt'
    [IO.File]::WriteAllLines($warnScript, [string[]]@(
        'param($Tool, $Path, $OutputFolder, $ResultFile)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $Path).AddParameter(''OutputFolder'', $OutputFolder)'
        '[void]$ps.Invoke()'
        '$lines = @($ps.Streams.Warning | ForEach-Object { $_.Message.Replace("`r", ''<CR>'').Replace("`n", ''<LF>'') })'
        '[IO.File]::WriteAllLines($ResultFile, [string[]]$lines, (New-Object Text.UTF8Encoding $true))'
    ), (New-Object Text.UTF8Encoding $true))
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $warnScript -Tool $tool -Path "$extra\flat" -OutputFolder '\\localhost\workbench_no_such_share\out' -ResultFile $warnFile 2>&1
    $saveWarnings = @()
    if (Test-Path -LiteralPath $warnFile) { $saveWarnings = @([IO.File]::ReadAllLines($warnFile) | Where-Object { $_ -like '*保存先に書き込めないため*' }) }
    Assert-Equal '保存先を作れないときの警告が出る' ($saveWarnings.Count -gt 0) 'True'
    Assert-Equal '保存先を作れないときの警告に改行が入らない' (@($saveWarnings | Where-Object { $_.Contains('<CR>') -or $_.Contains('<LF>') }).Count) 0

    # 日付の書式が、OS の地域設定によって変わっていた
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00') {
        $cultureOut = Join-Path $workRoot 'out_culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), "$extra\flat".Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $found = @(Get-ChildItem -LiteralPath $cultureOut -Directory -ErrorAction SilentlyContinue)
        $x = Read-Result $(if ($found.Count -gt 0) { $found[0].FullName } else { $null })
        Assert-Equal 'fi-FI でも summary の日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($x.Summary, '対象: .*   \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   ')) 'True'
        Assert-Equal 'fi-FI でも files.csv の日時は yyyy/MM/dd HH:mm:ss' ([regex]::IsMatch("$(@($x.Files | ForEach-Object { $_.'最終更新日時' }) | Select-Object -First 1)", '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$')) 'True'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }

    # 調べ終わった後（並べ替えや保存の途中）に Ctrl+C すると、途中までの結果が残っていた
    $many = Join-Path $workRoot 'many'
    for ($i = 0; $i -lt 150; $i++) { for ($j = 0; $j -lt 100; $j++) { [void][IO.Directory]::CreateDirectory("$many\d$i\e$j") } }
    $saveStopScript = Join-Path $workRoot 'stop-during-save.ps1'
    $saveStopOut    = Join-Path $workRoot 'out_ctrlc_save'
    [IO.File]::WriteAllLines($saveStopScript, [string[]]@(
        'param($Tool, $Target, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $Target).AddParameter(''OutputFolder'', $OutRoot)'
        '$handle  = $ps.BeginInvoke()'
        '$sw      = [Diagnostics.Stopwatch]::StartNew()'
        '$stopped = $false'
        'while (-not $handle.IsCompleted -and $sw.Elapsed.TotalSeconds -lt 120) {'
        '    if (@(Get-ChildItem -LiteralPath $OutRoot -Recurse -Filter ''summary.txt'' -ErrorAction SilentlyContinue).Count -gt 0) { $ps.Stop(); $stopped = $true; break }'
        '    Start-Sleep -Milliseconds 20'
        '}'
        '''STOPPED='' + $stopped'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + (($ps.Streams.Information | ForEach-Object { $_.MessageData.ToString() } | Where-Object { $_ -like ''*中断*'' }) -join '' / '')'
    ), (New-Object Text.UTF8Encoding $true))
    [void][IO.Directory]::CreateDirectory($saveStopOut)
    $result = @{}
    & $Exe -NoProfile -ExecutionPolicy Bypass -File $saveStopScript -Tool $tool -Target $many -OutRoot $saveStopOut 2>&1 | ForEach-Object {
        $line = "$_"
        $eq = $line.IndexOf('=')
        if ($eq -gt 0) { $result[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
    }
    if ($result['STOPPED'] -eq 'True') {
        Assert-Equal '保存の途中で止めた後は保存先が残らない' $result['AFTER'] '0'
        Assert-Equal '保存の途中で止めたときのメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip '保存の途中で Ctrl+C' 'summary.txt ができる前に保存まで終わったため、途中で止められなかった'
    }

    Write-Section '設定ファイル（-ConfigFile）'
    $cwd = Join-Path $workRoot 'cwd'
    [void][IO.Directory]::CreateDirectory($cwd)
    $basic = New-ConfigFile "$cfgDir\basic.ini" @(
        '# 設定ファイルのテスト（日本語のコメント）'
        '; セミコロンのコメント'
        'path = ..\..\tree'
        'DEPTH=2'
        'Top = 5'
        'OlderThan ='
        'Exclude = *\node_modules'
        'Exclude = "*\LVL2"'
        'OutputFolder = .\結果'
        'NoCsv = false'
    )
    # 別のフォルダから起動して、相対パスが設定ファイルのフォルダを基準にすることを確かめる
    $run = Invoke-From $cwd @('-ConfigFile', $basic)
    Assert-Equal '終了コード' $run.Code 0
    $runDir = Get-RunDir "$cfgDir\結果"
    Assert-Equal '保存先は設定ファイルのフォルダが基準' (Get-ResultFileNames $runDir) $allFiles
    Assert-Equal '起動したフォルダには保存しない' (Test-Path -LiteralPath "$cwd\結果") 'False'
    $x = Read-Result $runDir
    Assert-Equal 'ファイル数（Exclude を 2 行に分けて書く）' $x.Files.Count 6
    Assert-Equal 'ファイルの合計バイト' (Get-ByteSum $x.Files) 5507234
    Assert-Equal '除外指定の件数' @($x.Skipped | Where-Object { $_.'理由' -eq '除外指定' }).Count 3
    Assert-Equal 'Depth と Top' ([regex]::IsMatch($x.Summary, '■ 大きいフォルダ（2 階層まで・上位 5）')) 'True'
    Assert-Equal 'フォルダの階層は 2 まで' (($x.Folders | ForEach-Object { [int]$_.'階層' } | Measure-Object -Maximum).Maximum) 2
    Assert-Equal '値が空の OlderThan は指定なし' $x.Summary.Contains('条件:') 'False'
    Assert-Equal 'summary の 3 行目に設定ファイル' (($x.Summary -split '\r?\n')[2]) "設定ファイル: $basic"

    # Windows 11 のメモ帳は BOM なしの UTF-8、Server 2016 / 2019 のメモ帳は Shift_JIS で保存する
    $encodings = [ordered]@{
        'UTF-8（BOM なし）' = New-Object Text.UTF8Encoding $false
        'UTF-8（BOM 付き）' = New-Object Text.UTF8Encoding $true
        'Shift_JIS'         = [Text.Encoding]::GetEncoding(932)
        'UTF-16 LE'         = [Text.Encoding]::Unicode
        'UTF-16 BE'         = [Text.Encoding]::BigEndianUnicode
    }
    $i = 0
    foreach ($label in $encodings.Keys) {
        $file = New-ConfigFile "$cfgDir\enc$i.ini" @('Path = ..\..\tree', 'NoCsv = true', 'Exclude = *\日本語 フォルダ') $encodings[$label]
        $run  = Invoke-From $cwd @('-ConfigFile', $file)
        Assert-Equal ('文字コード {0}: 日本語の Exclude が効く' -f $label) ('{0}/{1}/{2}' -f $run.Code, (Get-FileCountText $run.Log), [regex]::Match($run.Log, '除外したフォルダ \S+ 件').Value) '0/ファイル 7 件/除外したフォルダ 1 件'
        $i++
    }

    New-SizedFile "$workRoot\cfg\t2\a#b;c\f.bin" 10
    $variants = @(
        @('項目名の大文字と小文字・= の前後に空白なし', @('PATH=..\..\tree\日本語 フォルダ', 'nocsv=TRUE'), 'ファイル 2 件'),
        @('"…" で囲んだ、スペースを含むパス', @('Path = "..\..\tree\日本語 フォルダ"', 'NoCsv = true'), 'ファイル 2 件'),
        @("'…' で囲んだパス", @("Path = '..\..\tree\日本語 フォルダ'", 'NoCsv = true'), 'ファイル 2 件'),
        @('# と ; を含むフォルダ名', @('Path = ..\t2\a#b;c', 'NoCsv = true'), 'ファイル 1 件'),
        @('Exclude の全体を "…" で囲む', @('Path = ..\..\tree', 'NoCsv = true', 'Exclude = "*\node_modules,*\LVL2"'), '除外したフォルダ 3 件'),
        @('Exclude を 1 つずつ "…" で囲む', @('Path = ..\..\tree', 'NoCsv = true', 'Exclude = "*\node_modules", "*\LVL2"'), '除外したフォルダ 3 件'),
        @('OlderThan', @('Path = ..\..\tree', 'NoCsv = true', 'OlderThan = 100'), '条件: 最終更新が 100 日以上前のファイルのみ')
    )
    for ($i = 0; $i -lt $variants.Count; $i++) {
        $file = New-ConfigFile "$cfgDir\var$i.ini" $variants[$i][1]
        $run  = Invoke-From $cwd @('-ConfigFile', $file)
        Assert-Equal ('書き方: {0}' -f $variants[$i][0]) ('{0}/{1}' -f $run.Code, $run.Log.Contains($variants[$i][2])) '0/True'
    }

    # コマンドラインの値が優先する。-Exclude は置き換え（両方を合わせない）、-NoCsv は NoCsv = false より優先
    $prio = New-ConfigFile "$cfgDir\prio.ini" @('Path = ..\..\tree\日本語 フォルダ', 'Top = 5', 'Exclude = *\node_modules', 'OutputFolder = .\prio-out', 'NoCsv = false')
    $run  = Invoke-From $cwd @('-ConfigFile', $prio, '-Path', $tree, '-Top', '1', '-Exclude', '*\pkg', '-NoCsv')
    Assert-Equal 'コマンドラインが優先: 終了コード' $run.Code 0
    Assert-Equal 'コマンドラインが優先: -Path' (Get-TargetText $run.Log) $tree
    Assert-Equal 'コマンドラインが優先: -Top' $run.Log.Contains('（3 階層まで・上位 1）') 'True'
    Assert-Equal 'コマンドラインが優先: -Exclude は置き換え' ('{0}/{1}' -f (Get-FileCountText $run.Log), $run.Log.Contains('除外: *\pkg')) 'ファイル 8 件/True'
    Assert-Equal 'コマンドラインが優先: -NoCsv' (Test-Path -LiteralPath "$cfgDir\prio-out") 'False'

    # -ConfigFile の相対パスは、起動したフォルダが基準
    $min = New-ConfigFile "$cfgDir\min.ini" @('Path = ..\..\tree', 'NoCsv = true')
    $run = Invoke-From (Join-Path $workRoot 'cfg') @('-ConfigFile', '設定 フォルダ\min.ini')
    Assert-Equal '相対パスの -ConfigFile' ('{0}/{1}/{2}' -f $run.Code, (Get-FileCountText $run.Log), $run.Log.Contains("設定ファイル: $min")) '0/ファイル 9 件/True'

    $cmdIni = New-ConfigFile "$workRoot\cfg\cmd.ini" @('Path = ..\tree', 'NoCsv = true', 'Exclude = *\日本語 フォルダ')
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -ConfigFile "{2}"' -f $Exe, $tool, $cmdIni)
    Assert-Equal 'コマンドプロンプトから -ConfigFile' ('{0}/{1}' -f $run.Code, (Get-FileCountText $run.Log)) '0/ファイル 7 件'
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -ConfigFile "{2}\" -NoCsv' -f $Exe, $tool, "$workRoot\cfg")
    Assert-Equal 'cmd の -ConfigFile "…\" の後ろに引数' ('{0}/{1}' -f $run.Code, [regex]::IsMatch($run.Log, '最後に \\ を付けないでください')) '1/True'

    $longIni = New-ConfigFile ("$workRoot\cfg\" + ('L' * 120) + '\' + ('M' * 120) + '\long.ini') @('Path = .', 'NoCsv = true')
    $run = Invoke-From $cwd @('-ConfigFile', $longIni)
    Assert-Equal '260 文字を超えるパスの設定ファイル' ('{0}/{1}/{2}' -f $run.Code, (Get-FileCountText $run.Log), $run.Log.Contains("設定ファイル: $longIni")) '0/ファイル 1 件/True'

    Write-Section '設定ファイル: 書き間違い（調べ始める前に止まる）'
    $mistakes = @(
        @('知らない項目', @('Path = ..\..\tree', 'Exlude = *\a'), @('設定ファイルの 2 行目: 知らない項目です: Exlude', '使える項目: Path, Depth, Top, OlderThan, Exclude, OutputFolder, NoCsv')),
        @('ConfigFile は書けない', @('ConfigFile = other.ini'), @('設定ファイルの 1 行目: 知らない項目です: ConfigFile')),
        @('同じ項目を 2 回', @('Depth = 2', 'Top = 5', 'depth = 3'), @('設定ファイルの 3 行目: Depth が 2 回書かれています（1 行目と 3 行目）')),
        @('値が空の行を挟んで 2 回', @('Depth = 2', 'Depth =', 'Depth = 3'), @('設定ファイルの 3 行目: Depth が 2 回書かれています（1 行目と 3 行目）')),
        @('Depth が範囲外', @('Depth = 0'), @('設定ファイルの 1 行目: Depth は 1～100 の数値で指定してください: 0')),
        @('Top が範囲外', @('Top = 1001'), @('Top は 1～1000 の数値で指定してください: 1001')),
        @('OlderThan が数値でない', @('OlderThan = abc'), @('OlderThan は 1～36500 の数値で指定してください: abc')),
        @('符号の付いた数値', @('Depth = +3'), @('Depth は 1～100 の数値で指定してください: +3')),
        @('NoCsv が true / false 以外', @('NoCsv = yes'), @('NoCsv は true か false で指定してください: yes')),
        @('= がない行', @('Depth 4'), @('「項目 = 値」の形になっていません: Depth 4')),
        @('項目名がない行', @('= 4'), @('「項目 = 値」の形になっていません: = 4')),
        @('セクションの見出し', @('[Get-DiskUsage]', 'Depth = 2'), @('設定ファイルの 1 行目: 「項目 = 値」の形になっていません: [Get-DiskUsage]')),
        @('Path の引用符が片方だけ', @('Path = "D:\data'), @('Path の値に引用符（"）が残っています: "D:\data')),
        @('Exclude の引用符が片方だけ', @('Exclude = "*\a'), @('Exclude の値に引用符（"）が残っています: "*\a')),
        @('定義されていない環境変数（OutputFolder）', @('OutputFolder = \\server\share\%WORKBENCH_NO_SUCH_VAR%'), @('定義されていない環境変数です: %WORKBENCH_NO_SUCH_VAR%')),
        @('定義されていない環境変数（Exclude）', @('Exclude = %WORKBENCH_NO_SUCH_VAR%\*'), @('定義されていない環境変数です: %WORKBENCH_NO_SUCH_VAR%'))
    )
    for ($i = 0; $i -lt $mistakes.Count; $i++) {
        $file    = New-ConfigFile "$cfgDir\err$i.ini" $mistakes[$i][1]
        $run     = Invoke-From $cwd @('-ConfigFile', $file)
        $missing = @(@($mistakes[$i][2]) + "ファイル: $file" | Where-Object { -not $run.Log.Contains($_) })
        # 終了コード / 表示に足りない文 / 調べ始めたか
        Assert-Equal ('書き間違い: {0}' -f $mistakes[$i][0]) ('{0}/{1}/{2}' -f $run.Code, ($missing -join ' | '), $run.Log.Contains('■ 大きいフォルダ')) '1//False'
    }
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\nothing.ini")
    Assert-Equal '設定ファイルが見つからない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("設定ファイルが見つかりません: $cfgDir\nothing.ini")) '1/True'
    $null = New-ConfigFile "$cfgDir\hidden.ini.txt" @('Depth = 2')
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\hidden.ini")
    Assert-Equal '拡張子が .txt になっている: ヒント' ('{0}/{1}' -f $run.Code, $run.Log.Contains("拡張子が .txt になっていませんか: $cfgDir\hidden.ini.txt")) '1/True'
    $null = New-ConfigFile $cfgLocked @('Depth = 2')
    $null = & icacls.exe $cfgLocked /deny "*${userSid}:(R)" 2>&1
    $script:cfgDenyApplied = ($LASTEXITCODE -eq 0)
    if ($script:cfgDenyApplied) {
        $run = Invoke-From $cwd @('-ConfigFile', $cfgLocked)
        Assert-Equal '設定ファイルを読めない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("設定ファイルを読めません: $cfgLocked")) '1/True'
        $null = & icacls.exe $cfgLocked /remove:d "*$userSid" 2>&1
        $script:cfgDenyApplied = $false
    } else {
        Write-Skip '設定ファイルを読めない' 'アクセス拒否を設定できなかったため'
    }

    Write-Section '設定ファイル: レビューで見つかった問題の再発防止'
    # %TEMP% などの環境変数が展開されず、%TEMP% という名前のフォルダに保存していた
    $env:WORKBENCH_TEST_ROOT = $workRoot
    $envIni = New-ConfigFile "$cfgDir\env.ini" @('Path = %WORKBENCH_TEST_ROOT%\tree', 'OutputFolder = %WORKBENCH_TEST_ROOT%\cfg-env-out', 'Exclude = %WORKBENCH_TEST_ROOT%\tree\node_modules')
    $run = Invoke-From $cwd @('-ConfigFile', $envIni)
    Assert-Equal '環境変数: 終了コード' $run.Code 0
    Assert-Equal '環境変数: OutputFolder' (Get-ResultFileNames (Get-RunDir "$workRoot\cfg-env-out")) $allFiles
    Assert-Equal '環境変数: Exclude' ([regex]::Match($run.Log, '除外したフォルダ \S+ 件').Value) '除外したフォルダ 1 件'
    Assert-Equal '環境変数: %名前% のフォルダを作らない' (Test-Path -LiteralPath "$cfgDir\%WORKBENCH_TEST_ROOT%") 'False'

    # URL エンコードされたフォルダ名（%E3%81%82）を、定義されていない環境変数と誤って止めていた
    New-SizedFile "$workRoot\cfg\pe\%E3%81%82\f.bin" 10
    $urlIni = New-ConfigFile "$cfgDir\urlenc.ini" @('Path = ..\pe\%E3%81%82', 'Exclude = *\%E3%81%82\none', 'NoCsv = true')
    $run = Invoke-From $cwd @('-ConfigFile', $urlIni)
    Assert-Equal 'URL エンコードされたフォルダ名' ('{0}/{1}' -f $run.Code, (Get-FileCountText $run.Log)) '0/ファイル 1 件'

    # 値が空の行も「2 回書いた」に数え、ひな形の空の行の下に書き足すと止まっていた
    $file = New-ConfigFile "$cfgDir\empty-then.ini" @('Path = ..\..\tree', 'NoCsv = true', 'Depth =', 'Depth = 1')
    $run  = Invoke-From $cwd @('-ConfigFile', $file)
    Assert-Equal '値が空の行の下に同じ項目' ('{0}/{1}' -f $run.Code, $run.Log.Contains('■ 大きいフォルダ（1 階層まで')) '0/True'

    # 設定ファイルの Path が見つからない・OutputFolder に書き込めないとき、設定ファイルの値だと分からなかった
    $noPath = New-ConfigFile "$cfgDir\nopath.ini" @('NoCsv = true', 'Path = .\nosuch')
    $run = Invoke-From $cwd @('-ConfigFile', $noPath)
    Assert-Equal '設定ファイルの Path が見つからない: 行番号' ('{0}/{1}' -f $run.Code, $run.Log.Contains("指定したフォルダが見つかりません: $cfgDir\nosuch（設定ファイルの 2 行目）")) '1/True'
    $run = Invoke-From $cwd @('-ConfigFile', $noPath, '-Path', (Join-Path $workRoot 'no\such'))
    Assert-Equal 'コマンドラインの -Path が見つからない: 行番号を付けない' ('{0}/{1}' -f $run.Code, $run.Log.Contains('設定ファイルの')) '1/False'
    New-SizedFile "$workRoot\cfg\blocker.txt" 1
    $badOut = New-ConfigFile "$cfgDir\badout.ini" @('Path = ..\..\tree\日本語 フォルダ', 'OutputFolder = ..\blocker.txt\sub')
    $run = Invoke-From $cwd @('-ConfigFile', $badOut)
    Assert-Equal '設定ファイルの OutputFolder に書き込めない: 行番号' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため'), $run.Log.Contains('（設定ファイルの 2 行目）')) '0/True/True'

    # 5.1 は、名前に [ ] を含むフォルダで起動すると、名前をワイルドカードとして扱い、別の場所（$PSHOME、
    # 1 つだけ一致した隣のフォルダ、ドライブのルート）から始まる。-Path を省略すると、そこを黙って調べていた
    $start = Join-Path $workRoot 'start'
    $startCases = @(
        @('一致するフォルダがない（$PSHOME から始まる）', 'br[1] x', @()),
        @('隣に 1 つ一致するフォルダがある', 'sib\[a]', @('sib\a')),
        @('[ が閉じていない（ドライブのルートから始まる）', 'ub\a[b', @('ub\ab')),
        @('隣に 2 つ一致する（ドライブのルートから始まる）', 'mm\[ab]', @('mm\a', 'mm\b'))
    )
    foreach ($case in $startCases) {
        New-SizedFile "$start\$($case[1])\f.bin" 11
        foreach ($decoy in $case[2]) { New-SizedFile "$start\$decoy\decoy.bin" 22 }
    }
    foreach ($case in $startCases) {
        $dir = "$start\$($case[1])"
        $run = Invoke-From $dir @('-NoCsv')
        Assert-Equal ('[ ] を含むフォルダから起動: {0}' -f $case[0]) ('{0}/{1}' -f (Get-TargetText $run.Log), (Get-FileCountText $run.Log)) "$dir/ファイル 1 件"
    }
    # 相対の -ConfigFile・-OutputFolder・-Path も、隣のフォルダではなく起動したフォルダを基準にする
    $sibDir = "$start\sib\[a]"
    $null = New-ConfigFile "$sibDir\rel.ini" @('NoCsv = false')
    $run = Invoke-From $sibDir @('-ConfigFile', 'rel.ini', '-OutputFolder', 'out', '-Path', '.')
    Assert-Equal '[ ] を含むフォルダから相対パス: 終了コード' $run.Code 0
    Assert-Equal '[ ] を含むフォルダから相対パス: 設定ファイル' $run.Log.Contains("設定ファイル: $sibDir\rel.ini") 'True'
    Assert-Equal '[ ] を含むフォルダから相対パス: 保存先' (Get-ResultFileNames (Get-RunDir "$sibDir\out")) $allFiles
    Assert-Equal '[ ] を含むフォルダから相対パス: 隣のフォルダに保存しない' (Test-Path -LiteralPath "$start\sib\a\out") 'False'

    Write-Section '共有フォルダ経由（\\localhost\<ドライブ>$）'
    $uncTree = '\\localhost\{0}${1}' -f $tree.Substring(0, 1), $tree.Substring(2)
    if (Test-Path -LiteralPath $uncTree) {
        $run = Invoke-Tool 'unc' @('-Path', $uncTree)
        Test-DefaultResult $run $uncTree
        Assert-Equal 'ドライブの行（共有）' ([regex]::IsMatch((Read-Result $run.Dir).Summary, 'ドライブ \\\\localhost\\')) 'True'
    } else {
        Write-Skip '共有フォルダ経由' '管理共有を読めないため'
    }

    Write-Section '実機: Windows の互換用リンク（All Users）をたどらない'
    $usersDir = Split-Path $env:PUBLIC -Parent
    $allUsers = Join-Path $usersDir 'All Users'
    if (Test-Path -LiteralPath $allUsers) {
        $run = Invoke-Tool 'allusers' @('-Path', $usersDir, '-Depth', '1')
        Assert-Equal '終了コード' $run.Code 0
        $x = Read-Result $run.Dir
        Assert-Equal 'All Users を順位に載せない' @($x.Folders | Where-Object { $_.'フォルダ' -eq $allUsers }).Count 0
        Assert-Equal 'All Users はリンクとして飛ばす' (@($x.Skipped | Where-Object { $_.'パス' -eq $allUsers }) | ForEach-Object { $_.'理由' }) 'リンク'
        # 5.1 は読み取りが禁止されたリンクの種類を判定できない。7 は判定できるが、5.1 と同じ結果になるよう互換用リンクとする
        Assert-Equal 'All Users の詳細（5.1 と 7 で同じ）' (@($x.Skipped | Where-Object { $_.'パス' -eq $allUsers }) | ForEach-Object { $_.'詳細' }) 'Windows の互換用リンク'
    } else {
        Write-Skip 'All Users' 'このPCにないため'
    }
} finally {
    # アクセス拒否を戻してから消す
    if ($script:denyApplied) { $null = & icacls.exe $deniedDir /remove:d "*$userSid" 2>&1 }
    if ($script:cfgDenyApplied) { $null = & icacls.exe $cfgLocked /remove:d "*$userSid" 2>&1 }
    Remove-Item Env:\WORKBENCH_TEST_ROOT -ErrorAction SilentlyContinue
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        # ループするジャンクションは、リンク先をたどらないように先にリンクだけ消す
        try {
            $junction = Join-Path $tree 'jloop'
            if (Test-Path -LiteralPath $junction) { [IO.Directory]::Delete($junction) }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

exit $script:failCount
