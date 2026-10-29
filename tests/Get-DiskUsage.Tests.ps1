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

function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out  = Join-Path $workRoot "out_$Label"
    $log  = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String
    $code = $LASTEXITCODE
    $dir  = $null
    if (Test-Path -LiteralPath $out) {
        $found = @(Get-ChildItem -LiteralPath $out -Directory)
        if ($found.Count -gt 0) { $dir = $found[0].FullName }
    }
    [pscustomobject]@{ Code = $code; Log = $log; Dir = $dir }
}

function Read-Result([string]$Dir) {
    if (-not $Dir) {
        return [pscustomobject]@{ Folders = @(); Files = @(); Exts = @(); Skipped = @(); Summary = '' }
    }
    [pscustomobject]@{
        Folders = @(Import-Csv -LiteralPath (Join-Path $Dir 'folders.csv') -Encoding UTF8)
        Files   = @(Import-Csv -LiteralPath (Join-Path $Dir 'files.csv') -Encoding UTF8)
        Exts    = @(Import-Csv -LiteralPath (Join-Path $Dir 'extensions.csv') -Encoding UTF8)
        Skipped = @(Import-Csv -LiteralPath (Join-Path $Dir 'skipped.csv') -Encoding UTF8)
        Summary = [IO.File]::ReadAllText((Join-Path $Dir 'summary.txt'))
    }
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $log = & cmd.exe /d /c $batch 2>&1 | Out-String
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
    $log = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path (Join-Path $workRoot 'no\such\folder') 2>&1 | Out-String
    Assert-Equal '存在しないパス: 終了コード' $LASTEXITCODE 1
    Assert-Equal '存在しないパス: メッセージ' ([regex]::IsMatch($log, '指定したフォルダが見つかりません')) 'True'
    $usedLetters = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    $freeLetter  = @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $usedLetters -notcontains $_ } | Select-Object -First 1
    if ($freeLetter) {
        $log = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $tree -OutputFolder "${freeLetter}:\nope" 2>&1 | Out-String
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
