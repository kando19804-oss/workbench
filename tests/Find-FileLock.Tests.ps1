<#
.SYNOPSIS
    Find-FileLock の回帰テストです。

.DESCRIPTION
    一時フォルダのファイルをこのテスト自身やテストが起動したプロセスで掴み、Find-FileLock で実際に検出できるかを確かめます。
    共有フォルダの接続と、API が返す特殊なエラーは本物を作れないので、Get-SmbOpenFile と同じ名前の偽の関数や、
    同じ名前の偽の型を読み込んだラッパーから実行します。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Find-FileLock を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Find-FileLock\Find-FileLock.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Find-FileLock_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
# 結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーしたスクリプトを動かす
$tool     = Join-Path $workRoot 'tool\Find-FileLock.ps1'
$wrapper  = Join-Path $workRoot 'wrapper.ps1'
$data     = Join-Path $workRoot 'data'
$script:failCount = 0
$script:runNumber = 0
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

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

# コンソールで実行すると、長い行は画面の幅で折り返され、出力に改行が入る。
# 画面の幅によって結果が変わらないように、文字を探す前に改行を取り除く
function ConvertTo-FlatText([string]$Text) { $Text -replace '\r?\n', '' }

function Test-Contains([string]$Text, [string]$Part) { $Text.IndexOf($Part, [StringComparison]::Ordinal) -ge 0 }

function Write-Text([string]$FilePath, [string]$Text) {
    [void][IO.Directory]::CreateDirectory((Split-Path $FilePath))
    [IO.File]::WriteAllText($FilePath, $Text, (New-Object Text.UTF8Encoding $false))
}

function Read-Run([string]$Out, [string]$Log, [int]$Code) {
    $dir = $null
    if (Test-Path -LiteralPath $Out) { $dir = @(Get-ChildItem -LiteralPath $Out -Directory | Where-Object { $_.Name -like 'FileLock_*' } | ForEach-Object { $_.FullName })[0] }
    $rows = @()
    $skips = @()
    if ($dir) {
        $rows  = @(Import-Csv -LiteralPath (Join-Path $dir 'locks.csv') -Encoding UTF8)
        $skips = @(Import-Csv -LiteralPath (Join-Path $dir 'skipped.csv') -Encoding UTF8)
    }
    [pscustomobject]@{ Code = $Code; Log = (ConvertTo-FlatText $Log); Dir = $dir; Rows = $rows; Skips = $skips }
}

# ツールをそのまま動かす。結果は毎回別のフォルダに保存する
function Invoke-Tool([string[]]$ToolArgs) {
    $script:runNumber++
    $out  = Join-Path $workRoot ('out_{0}' -f $script:runNumber)
    $log  = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String
    $code = $LASTEXITCODE
    Read-Run $out $log $code
}

# 偽の型や関数を読み込んだラッパーから動かす。Mode は偽の Query の動き、Smb は Get-SmbOpenFile が返す値
function Invoke-Fake([string]$Target, [string]$Mode, [string]$SmbJson) {
    $script:runNumber++
    $out = Join-Path $workRoot ('out_{0}' -f $script:runNumber)
    $env:FAKE_QUERY = $Mode
    $env:FAKE_SMB   = $SmbJson
    $log  = & $Exe -NoProfile -ExecutionPolicy Bypass -File $wrapper -Tool $tool -Target $Target -Out $out 2>&1 | Out-String
    $code = $LASTEXITCODE
    $env:FAKE_QUERY = $null
    $env:FAKE_SMB   = $null
    Read-Run $out $log $code
}

function Get-Clean($Value) { ([string]$Value) -replace '^="', '' -replace '"$', '' }

# 実際に掴むファイルを開く
function Open-Held([string]$FilePath, [IO.FileShare]$Share) {
    return [IO.File]::Open(('\\?\' + $FilePath), [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, $Share)
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

$held      = New-Object 'System.Collections.Generic.List[object]'
$processes = New-Object 'System.Collections.Generic.List[object]'
$aclFolder = $null
$script:substUsed = $null
try {
    #region ダミー環境

    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    $wrapperCode = @'
param([string]$Tool, [string]$Target, [string]$Out)
Add-Type -TypeDefinition @"
namespace Workbench.FindFileLock { public static class Native {
    public static int Query(string path, bool openLink, out int[] ids) {
        ids = new int[0];
        string mode = System.Environment.GetEnvironmentVariable("FAKE_QUERY");
        if (mode == "empty") { return 0; }
        if (mode == "denied") { return 5; }
        if (mode == "gone") { return 2; }
        if (mode == "sharing") { return 32; }
        if (mode == "error") { return 87; }
        if (mode == "many") { ids = new int[300]; for (int i = 0; i < 300; i++) { ids[i] = 100000 + i; } return 0; }
        if (mode == "link") { ids = new int[] { openLink ? 1 : 4 }; return 0; }
        return 0;
    }
} }
"@
function global:Get-SmbOpenFile {
    if ($env:FAKE_SMB -eq 'throw') { throw (New-Object UnauthorizedAccessException 'denied') }
    if ($env:FAKE_SMB) { @($env:FAKE_SMB | ConvertFrom-Json) | ForEach-Object { $_ } }
}
& $Tool -Path $Target -OutputFolder $Out
exit $LASTEXITCODE
'@
    Write-Text $wrapper $wrapperCode

    Write-Text "$data\free.txt" 'x'
    Write-Text "$data\held_none.txt" 'x'
    Write-Text "$data\held_share.txt" 'x'
    Write-Text "$data\日本語 と スペース\掴み中 ファイル.txt" 'x'
    Write-Text "$data\[bracket]\held.txt" 'x'
    Write-Text "$data\tree\a\b\deep.txt" 'x'
    Write-Text "$data\tree\c\other.txt" 'x'
    Write-Text "$data\readonly.txt" 'x'
    [IO.File]::SetAttributes("$data\readonly.txt", [IO.FileAttributes]::ReadOnly)
    Write-Text "$data\-minus.txt" 'x'

    #endregion

    #region 正常系

    Write-Section '使われていないファイル'
    $free = Invoke-Tool @('-Path', "$data\free.txt")
    Assert-Equal '終了コード' $free.Code '0'
    Assert-Equal '結果の行' (Test-Contains $free.Log '結果: 使用中のものは見つかりません') 'True'
    Assert-Equal 'プロセスの欄は（なし）' (Test-Contains $free.Log '（なし）') 'True'
    Assert-Equal '開けるかの確認' (Test-Contains $free.Log '可') 'True'
    Assert-Equal 'locks.csv は見出しだけ' $free.Rows.Count '0'

    $stampBefore = [IO.File]::GetLastWriteTimeUtc("$data\free.txt")
    Start-Sleep -Milliseconds 50
    $bytesBefore = [IO.File]::ReadAllText("$data\free.txt")
    [void](Invoke-Tool @('-Path', "$data\free.txt"))
    Assert-Equal '書き込みのモードで開いても更新日時が変わらない' ([IO.File]::GetLastWriteTimeUtc("$data\free.txt") -eq $stampBefore) 'True'
    Assert-Equal '内容も変わらない' ([IO.File]::ReadAllText("$data\free.txt")) $bytesBefore

    Write-Section '排他で掴んでいるファイル'
    $held.Add((Open-Held "$data\held_none.txt" ([IO.FileShare]::None)))
    $lock = Invoke-Tool @('-Path', "$data\held_none.txt")
    Assert-Equal '終了コード' $lock.Code '2'
    Assert-Equal '結果の行' (Test-Contains $lock.Log '結果: 使用中') 'True'
    $mine = @($lock.Rows | Where-Object { $_.'PID' -eq "$PID" })
    Assert-Equal 'テストのプロセスが出る' $mine.Count '1'
    if ($mine.Count -gt 0) {
        Assert-Equal '種類' $mine[0].'種類' 'プロセス'
        Assert-Equal 'ファイル' (Get-Clean $mine[0].'ファイル') "$data\held_none.txt"
        Assert-Equal '読み取り' $mine[0].'読み取り' '不可（ほかのプログラムが使用中）'
        Assert-Equal '書き込み' $mine[0].'書き込み' '不可（ほかのプログラムが使用中）'
        Assert-Equal '実行ファイルが入る' ([bool]$mine[0].'実行ファイル') 'True'
    }

    Write-Section '共有ありで掴んでいるファイル（ロックではないが使用中）'
    $held.Add((Open-Held "$data\held_share.txt" ([IO.FileShare]'ReadWrite, Delete')))
    $share = Invoke-Tool @('-Path', "$data\held_share.txt")
    Assert-Equal '終了コード' $share.Code '2'
    $mine = @($share.Rows | Where-Object { $_.'PID' -eq "$PID" })
    Assert-Equal 'テストのプロセスが出る' $mine.Count '1'
    if ($mine.Count -gt 0) { Assert-Equal '書き込みは可' $mine[0].'書き込み' '可' }

    Write-Section '実行中の EXE・読み込み中の DLL・現在のフォルダ'
    $exeCopy = "$data\running\ping_copy.exe"
    [void][IO.Directory]::CreateDirectory((Split-Path $exeCopy))
    Copy-Item -LiteralPath "$env:SystemRoot\System32\ping.exe" -Destination $exeCopy
    $running = Start-Process -FilePath $exeCopy -ArgumentList '-n 60 127.0.0.1' -PassThru -WindowStyle Hidden
    $processes.Add($running)
    Start-Sleep -Milliseconds 500
    if ($running.HasExited) {
        Write-Skip '実行中の EXE' 'コピーした ping.exe が起動しませんでした'
    } else {
        $exeRun = Invoke-Tool @('-Path', $exeCopy)
        Assert-Equal '実行中の EXE が出る' (@($exeRun.Rows | Where-Object { $_.'PID' -eq "$($running.Id)" }).Count) '1'
    }

    $dllRun = Invoke-Tool @('-Path', "$env:SystemRoot\System32\kernel32.dll")
    Assert-Equal '読み込み中の DLL で使用中になる' $dllRun.Code '2'
    Assert-Equal '読み込み中の DLL に複数のプロセスが出る' (@($dllRun.Rows).Count -gt 1) 'True'

    [void][IO.Directory]::CreateDirectory("$data\cwd")
    $cwdProcess = Start-Process -FilePath "$env:SystemRoot\System32\ping.exe" -ArgumentList '-n 60 127.0.0.1' -WorkingDirectory "$data\cwd" -PassThru -WindowStyle Hidden
    $processes.Add($cwdProcess)
    Start-Sleep -Milliseconds 500
    $cwdRun = Invoke-Tool @('-Path', "$data\cwd")
    Assert-Equal '現在のフォルダとして使っているプロセスが出る' (@($cwdRun.Rows | Where-Object { $_.'PID' -eq "$($cwdProcess.Id)" }).Count) '1'
    Assert-Equal 'フォルダの行は読み取りが空欄' (@($cwdRun.Rows | Where-Object { $_.'PID' -eq "$($cwdProcess.Id)" } | ForEach-Object { $_.'読み取り' }) -join ',') ''

    Write-Section 'フォルダの中'
    $held.Add((Open-Held "$data\tree\a\b\deep.txt" ([IO.FileShare]::None)))
    $tree = Invoke-Tool @('-Path', "$data\tree")
    Assert-Equal '終了コード' $tree.Code '2'
    Assert-Equal '調べた数' (Test-Contains $tree.Log '調べた数: ファイル 2 件  フォルダ 4 件') 'True'
    Assert-Equal '深いファイルが出る' (@($tree.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq "$data\tree\a\b\deep.txt" }).Count -ge 1) 'True'
    Assert-Equal '使っていないファイルは出ない' (@($tree.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq "$data\tree\c\other.txt" }).Count) '0'

    Write-Section '名前と属性'
    $held.Add((Open-Held "$data\日本語 と スペース\掴み中 ファイル.txt" ([IO.FileShare]::None)))
    $jp = Invoke-Tool @('-Path', "$data\日本語 と スペース\掴み中 ファイル.txt")
    Assert-Equal '日本語とスペース' $jp.Code '2'
    $held.Add((Open-Held "$data\[bracket]\held.txt" ([IO.FileShare]::None)))
    $bracket = Invoke-Tool @('-Path', "$data\[bracket]")
    Assert-Equal '[ ] を含むフォルダ' $bracket.Code '2'
    $readonly = Invoke-Tool @('-Path', "$data\readonly.txt")
    Assert-Equal '読み取り専用は使用中ではない' $readonly.Code '0'
    Assert-Equal '読み取り専用の表示' (Test-Contains $readonly.Log '不可（読み取り専用）') 'True'
    $held.Add((Open-Held "$data\-minus.txt" ([IO.FileShare]::None)))
    $minus = Invoke-Tool @('-Path', "$data\-minus.txt")
    Assert-Equal '- で始まる名前も調べる' $minus.Code '2'

    # 260 文字を超えるパス
    $longDir = "$data\long\" + (('d' * 50 + '\') * 6)
    $longFile = $longDir + 'held.txt'
    [void][IO.Directory]::CreateDirectory('\\?\' + $longDir)
    [IO.File]::WriteAllText('\\?\' + $longFile, 'x')
    $held.Add((Open-Held $longFile ([IO.FileShare]::None)))
    $long = Invoke-Tool @('-Path', "$data\long")
    Assert-Equal '260 文字を超えるパス' $long.Code '2'
    Assert-Equal '長いパスは \\?\ なしで書く' (@($long.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq $longFile.TrimEnd('\') }).Count -ge 1) 'True'

    Write-Section 'リンク'
    [void][IO.Directory]::CreateDirectory("$data\junction\real")
    Write-Text "$data\jtarget\inside.txt" 'x'
    [void](& cmd.exe /c mklink /J "$data\junction\link" "$data\jtarget" 2>&1)
    if (Test-Path -LiteralPath "$data\junction\link") {
        $held.Add((Open-Held "$data\jtarget\inside.txt" ([IO.FileShare]::None)))
        $junction = Invoke-Tool @('-Path', "$data\junction")
        Assert-Equal 'リンクの先はたどらない' $junction.Code '0'
        Assert-Equal 'たどらなかったリンクの件数' (Test-Contains $junction.Log 'たどらなかったリンク 1 件') 'True'
        Assert-Equal 'skipped.csv に残す' (@($junction.Skips | Where-Object { $_.'理由' -eq 'たどらなかったリンク' }).Count) '1'
        # -Path にリンクそのものを指定したときは、リンク先を調べる
        $direct = Invoke-Tool @('-Path', "$data\junction\link")
        Assert-Equal '-Path のリンクの先を調べる' $direct.Code '2'
    } else {
        Write-Skip 'リンク' 'ジャンクションを作れません'
    }

    #endregion

    #region 異常系

    Write-Section '止まるとき'
    $none = Invoke-Tool @('-Path', "$data\nothing.txt")
    Assert-Equal '見つからない: 終了コード' $none.Code '1'
    Assert-Equal '見つからない: メッセージ' (Test-Contains $none.Log 'ファイルもフォルダも見つかりません') 'True'
    $noPath = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool 2>&1 | Out-String
    Assert-Equal '-Path なし: 終了コード' $LASTEXITCODE '1'
    Assert-Equal '-Path なし: メッセージ' (Test-Contains (ConvertTo-FlatText $noPath) '-Path で指定してください') 'True'

    $cmdFile = "$workRoot\quote.cmd"
    Write-Text $cmdFile ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -Path ""{2}\"" -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $Exe, $tool, $data)
    $quote = ConvertTo-FlatText (& cmd.exe /c $cmdFile 2>&1 | Out-String)
    Assert-Equal '\" で引数を巻き込む: 終了コード' $LASTEXITCODE '1'
    Assert-Equal '\" で引数を巻き込む: ヒント' (Test-Contains $quote '引用符') 'True'

    # コマンドプロンプトの "Z:\" が Z: で届いても、現在のフォルダではなくルートを指す（レビューの指摘）。
    # ドライブ全体を調べないよう、小さなフォルダを subst でドライブにして試す
    $substDrive = @(90..68 | ForEach-Object { [string][char]$_ } | Where-Object { -not (Test-Path -LiteralPath ($_ + ':\')) })[0]
    [void](& subst.exe ($substDrive + ':') "$data\tree" 2>&1)
    if (Test-Path -LiteralPath ($substDrive + ':\c')) {
        $script:substUsed = $substDrive
        $rootBat = "$workRoot\root.cmd"
        Write-Text $rootBat ("@echo off`r`nchcp 65001 > nul`r`ncd /d {0}:\c`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -NoSave -Path ""{0}:\""`r`nexit /b %ERRORLEVEL%`r`n" -f $substDrive, $Exe, $tool)
        $rootLog = ConvertTo-FlatText (& cmd.exe /c $rootBat 2>&1 | Out-String)
        Assert-Equal 'ドライブ名だけの値はルートを指す' (Test-Contains $rootLog ('対象: {0}:\' -f $substDrive)) 'True'
        Assert-Equal '現在のフォルダを調べない' (Test-Contains $rootLog ('対象: {0}:\c' -f $substDrive)) 'False'
    } else {
        Write-Skip 'ドライブ名だけの値' 'subst でドライブを作れません'
    }

    if ($hostTag -eq '51') {
        # 制約付き言語モードでは、最初に止まる（レビューの指摘）
        $env:__PSLockdownPolicy = '4'
        $clm = ConvertTo-FlatText (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$data\free.txt" -NoSave 2>&1 | Out-String)
        $clmCode = $LASTEXITCODE
        $env:__PSLockdownPolicy = $null
        Assert-Equal '制約付き言語モード: 終了コード' $clmCode '1'
        Assert-Equal '制約付き言語モード: メッセージ' (Test-Contains $clm '制約付き言語モード') 'True'
    } else {
        Write-Skip '制約付き言語モード' '5.1 だけで確かめます'
    }

    Write-Section '保存'
    $saved = Invoke-Tool @('-Path', "$data\free.txt")
    foreach ($name in 'summary.txt', 'locks.csv', 'skipped.csv') {
        $bytes = [IO.File]::ReadAllBytes((Join-Path $saved.Dir $name))
        Assert-Equal ("$name は BOM 付き") ('{0}-{1}-{2}' -f $bytes[0], $bytes[1], $bytes[2]) '239-187-191'
    }
    $header = [IO.File]::ReadAllLines((Join-Path $saved.Dir 'locks.csv'))[0]
    Assert-Equal 'locks.csv の列' $header '"種類","ファイル","PID","プロセス","サービス","実行ユーザー","起動日時","実行ファイル","コマンドライン","クライアント","クライアントのユーザー","権限","読み取り","書き込み"'
    $noSave = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$data\free.txt" -NoSave 2>&1 | Out-String
    Assert-Equal '-NoSave の終了コード' $LASTEXITCODE '0'
    Assert-Equal '-NoSave は保存先を出さない' (Test-Contains (ConvertTo-FlatText $noSave) '保存先:') 'False'
    Assert-Equal '-NoSave は結果を作らない' (@(Get-ChildItem -LiteralPath (Split-Path $tool) -Directory -Filter 'FileLock_*').Count) '0'
    Write-Text "$workRoot\notdir" 'x'
    $unwritable = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$data\free.txt" -OutputFolder "$workRoot\notdir" 2>&1 | Out-String
    Assert-Equal '書き込めない保存先: 終了コード' $LASTEXITCODE '0'
    Assert-Equal '書き込めない保存先: 警告' (Test-Contains (ConvertTo-FlatText $unwritable) '保存先に書き込めないため') 'True'

    Write-Section '読めないフォルダ'
    $aclFolder = "$data\denied"
    Write-Text "$aclFolder\secret.txt" 'x'
    $me = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $denied = & icacls.exe $aclFolder /deny ('{0}:(OI)(CI)(RD)' -f $me) 2>&1
    if ($LASTEXITCODE -eq 0) {
        $deny = Invoke-Tool @('-Path', $aclFolder)
        Assert-Equal '権限なし: 止まらず終了コード 1' $deny.Code '1'
        Assert-Equal '権限なし: skipped.csv に 1 件（2 件に数えない）' (@($deny.Skips | Where-Object { $_.'理由' -eq '権限なし' }).Count) '1'
        Assert-Equal '権限なし: 件数の表示' (Test-Contains $deny.Log 'スキップ 1 件（権限なし 1）') 'True'
    } else {
        Write-Skip '読めないフォルダ' "icacls に失敗: $denied"
    }
    [void](& icacls.exe $aclFolder /remove:d $me 2>&1)
    $aclFolder = $null

    #endregion

    #region 偽の API と偽の共有

    Write-Section 'API の戻り値'
    $empty = Invoke-Fake "$data\held_none.txt" 'empty' ''
    Assert-Equal '見つからないが共有違反: 終了コード 2' $empty.Code '2'
    Assert-Equal '見つからないが共有違反: 結果の行' (Test-Contains $empty.Log '結果: 使用中（使っているものは特定できません）') 'True'
    $gone = Invoke-Fake "$data\free.txt" 'gone' ''
    Assert-Equal '調べている間に消えた: スキップに数えない' $gone.Code '0'
    $apiDenied = Invoke-Fake "$data\free.txt" 'denied' ''
    Assert-Equal '権限なし: 終了コード 1' $apiDenied.Code '1'
    Assert-Equal '権限なし: 理由' (@($apiDenied.Skips | ForEach-Object { $_.'理由' }) -join ',') '権限なし'
    $apiError = Invoke-Fake "$data\free.txt" 'error' ''
    Assert-Equal 'その他のエラー: 理由' (@($apiError.Skips | ForEach-Object { $_.'理由' }) -join ',') 'その他のエラー'
    $sharing = Invoke-Fake "$data\held_none.txt" 'sharing' ''
    Assert-Equal '問い合わせが共有違反: 開けるかの確認で使用中' $sharing.Code '2'
    $linkMode = Invoke-Fake "$data\free.txt" 'link' ''
    Assert-Equal '-Path はリンクをたどって開く' (@($linkMode.Rows | Where-Object { $_.'PID' -eq '4' }).Count) '1'
    Assert-Equal 'PID 4 の注記' (Test-Contains $linkMode.Log '（カーネル：共有フォルダ経由の接続など）') 'True'
    $many = Invoke-Fake "$data\free.txt" 'many' ''
    Assert-Equal '画面は 200 行まで' (Test-Contains $many.Log '（ほか 100 行は locks.csv を参照してください）') 'True'
    Assert-Equal 'CSV はすべて' $many.Rows.Count '300'
    Assert-Equal '終了したプロセスの表示' (Test-Contains $many.Log '（終了したプロセス）') 'True'

    Write-Section '共有フォルダの接続'
    $d = $data.Replace('\', '\\')
    $smbJson = '[{"Path":"' + $d + '\\free.txt","ClientComputerName":"10.0.0.25","ClientUserName":"CONTOSO\\suzuki","Permissions":1179785},' +
        '{"Path":"' + $d + '\\free.txt","ClientComputerName":"10.0.0.25","ClientUserName":"CONTOSO\\suzuki","Permissions":1180063},' +
        '{"Path":"' + $d + '\\free.txt2","ClientComputerName":"10.0.0.99","ClientUserName":"x","Permissions":1}]'
    $smb = Invoke-Fake "$data\free.txt" 'empty' $smbJson
    Assert-Equal '接続があれば使用中' $smb.Code '2'
    $clientRows = @($smb.Rows | Where-Object { $_.'種類' -eq 'クライアント' })
    Assert-Equal '同じ接続は 1 行にまとめ、似た名前は含めない' $clientRows.Count '1'
    if ($clientRows.Count -gt 0) {
        Assert-Equal 'クライアント' (Get-Clean $clientRows[0].'クライアント') '10.0.0.25'
        Assert-Equal 'ユーザー' $clientRows[0].'クライアントのユーザー' 'CONTOSO\suzuki'
        Assert-Equal '権限は強いほう' $clientRows[0].'権限' '読み書き'
    }
    $smbTree = '[{"Path":"' + $d + '\\tree\\c\\other.txt","ClientComputerName":"pc1","ClientUserName":"u","Permissions":1},{"Path":"' + $d + '\\tree2\\x.txt","ClientComputerName":"pc2","ClientUserName":"u","Permissions":1}]'
    $smbFolder = Invoke-Fake "$data\tree" 'empty' $smbTree
    Assert-Equal 'フォルダの中の接続だけを出す' (@($smbFolder.Rows | Where-Object { $_.'種類' -eq 'クライアント' }).Count) '1'
    $smbThrow = Invoke-Fake "$data\free.txt" 'empty' 'throw'
    Assert-Equal '調べられなくても止まらない' $smbThrow.Code '0'
    $noteExpected = if ($isAdmin) { '共有フォルダの接続元を調べられませんでした' } else { '共有フォルダの接続元は、管理者として実行すると調べられます' }
    Assert-Equal '調べられなかった案内' (Test-Contains $smbThrow.Log $noteExpected) 'True'

    Write-Section '共有フォルダのパス'
    if ($isAdmin) {
        $drive = $data.Substring(0, 1)
        $unc = '\\localhost\' + $drive + '$' + $data.Substring(2) + '\free.txt'
        if (Test-Path -LiteralPath $unc) {
            $uncRun = Invoke-Tool @('-Path', $unc)
            Assert-Equal 'ネットワーク上のパスで何も見つからなければ 0 にしない' $uncRun.Code '1'
            Assert-Equal 'ファイルサーバーで実行する案内' (Test-Contains $uncRun.Log '別の PC からは、この PC のプロセスしか分かりません') 'True'
        } else {
            Write-Skip '共有フォルダのパス' '管理共有を開けません'
        }
    } else {
        Write-Skip '共有フォルダのパス' '管理者のときだけ確かめます'
    }

    #endregion

    #region 起動

    Write-Section '起動'
    # 名前に [ ] を含むフォルダから、相対パスで起動する
    Write-Text "$workRoot\[br]\a.txt" 'x'
    $brBat = "$workRoot\bracket.cmd"
    Write-Text $brBat ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\[br]""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -Path a.txt -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $workRoot, $Exe, $tool)
    $brLog = ConvertTo-FlatText (& cmd.exe /c $brBat 2>&1 | Out-String)
    Assert-Equal '[ ] を含むフォルダから起動する' (Test-Contains $brLog ('{0}\[br]\a.txt' -f $workRoot)) 'True'

    # 調べるフォルダの中で実行すると、案内を出す
    Write-Text "$data\self\x.txt" 'x'
    $selfBat = "$workRoot\self.cmd"
    Write-Text $selfBat ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\self""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -Path . -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $data, $Exe, $tool)
    $selfLog = ConvertTo-FlatText (& cmd.exe /c $selfBat 2>&1 | Out-String)
    Assert-Equal '調べるフォルダの中で実行した案内' (Test-Contains $selfLog 'このツールを、調べるフォルダの中で実行しています') 'True'

    #endregion
} finally {
    foreach ($stream in $held) { try { $stream.Dispose() } catch { } }
    if ($script:substUsed) { [void](& subst.exe ($script:substUsed + ':') /d 2>&1) }
    foreach ($process in $processes) { try { if (-not $process.HasExited) { $process.Kill(); $process.WaitForExit(5000) | Out-Null } } catch { } }
    if ($aclFolder) {
        try { [void](& icacls.exe $aclFolder /remove:d ([Security.Principal.WindowsIdentity]::GetCurrent().Name) 2>&1) } catch { }
    }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            if (Test-Path -LiteralPath "$data\readonly.txt") { [IO.File]::SetAttributes("$data\readonly.txt", [IO.FileAttributes]::Normal) }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

exit $script:failCount
