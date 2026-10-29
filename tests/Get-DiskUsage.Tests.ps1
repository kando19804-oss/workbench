<#
.SYNOPSIS
    Get-DiskUsage の回帰テストです。

.DESCRIPTION
    一時フォルダにダミー環境（日本語のパス、260 文字を超えるパス、ジャンクションのループ、
    読み取り禁止のフォルダ、古いファイルなど）を作り、Get-DiskUsage を実行して期待値と照合します。
    ダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダにだけアクセス拒否を設定します（終了時に戻します）。

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

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$script:denyApplied = $false
$deniedDir = Join-Path $tree 'denied'
$userSid   = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value

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
    } else {
        Write-Skip 'All Users' 'このPCにないため'
    }
} finally {
    # アクセス拒否を戻してから消す
    if ($script:denyApplied) { $null = & icacls.exe $deniedDir /remove:d "*$userSid" 2>&1 }
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
