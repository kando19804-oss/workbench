<#
.SYNOPSIS
    Find-FileLock の回帰テストです。

.DESCRIPTION
    一時フォルダのファイルをこのテスト自身やテストが起動したプロセスで掴み、Find-FileLock で実際に検出できるかを確かめます。
    共有フォルダの接続と、API が返す特殊なエラーは本物を作れないので、Get-SmbOpenFile と同じ名前の偽の関数や、
    同じ名前の偽の型を読み込んだラッパーから実行します。保存の失敗などの本物では起きない条件は、
    ツールの処理を 1 か所だけ書き換えたコピーで確かめます。ダミー環境は終了時に削除します。

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
$utf8Bom  = New-Object Text.UTF8Encoding $true
$invariant = [Globalization.CultureInfo]::InvariantCulture
$script:failCount = 0
$script:runNumber = 0
$me      = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$missingNote = '※ 管理者として実行すると、ほかのユーザーのプロセスの実行ユーザー・実行ファイル・コマンドラインも取れます'

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

# 文字列が何回出てくるか
function Get-Occurrence([string]$Text, [string]$Part) { [regex]::Matches($Text, [regex]::Escape($Part)).Count }

# 画面の幅（全角は 2 文字分、半角カタカナは 1 文字分）
function Get-Width([string]$Text) {
    $width = 0
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int]$ch
        if ($code -lt 0x100 -or ($code -ge 0xFF61 -and $code -le 0xFF9F)) { $width += 1 }
        elseif ($code -ge 0xDC00 -and $code -le 0xDFFF) { }
        else { $width += 2 }
    }
    $width
}

function Write-Text([string]$FilePath, [string]$Text) {
    [void][IO.Directory]::CreateDirectory((Split-Path $FilePath))
    [IO.File]::WriteAllText($FilePath, $Text, (New-Object Text.UTF8Encoding $false))
}

# 結果のファイルがないとき（保存の途中で失敗したときなど）も止まらないように、空として読む
function Read-CsvRows([string]$Dir, [string]$Name) {
    $file = Join-Path $Dir $Name
    if (-not [IO.File]::Exists($file)) { return @() }
    return @(Import-Csv -LiteralPath $file -Encoding UTF8)
}

function Read-Run([string]$Out, [string]$Log, [int]$Code) {
    $dir = $null
    if (Test-Path -LiteralPath $Out) { $dir = @(Get-ChildItem -LiteralPath $Out -Directory | Where-Object { $_.Name -like 'FileLock_*' } | ForEach-Object { $_.FullName })[0] }
    $rows = @()
    $skips = @()
    $summary = ''
    if ($dir) {
        $rows  = @(Read-CsvRows $dir 'locks.csv')
        $skips = @(Read-CsvRows $dir 'skipped.csv')
        if ([IO.File]::Exists((Join-Path $dir 'summary.txt'))) { $summary = [IO.File]::ReadAllText((Join-Path $dir 'summary.txt')) }
    }
    [pscustomobject]@{ Code = $Code; Log = (ConvertTo-FlatText $Log); Dir = $dir; Rows = $rows; Skips = $skips; Summary = $summary }
}

# ツールをそのまま動かす。結果は毎回別のフォルダに保存する。ToolPath には書き換えたコピーを渡せる
function Invoke-Tool([string[]]$ToolArgs, [string]$ToolPath) {
    if (-not $ToolPath) { $ToolPath = $tool }
    $script:runNumber++
    $out  = Join-Path $workRoot ('out_{0}' -f $script:runNumber)
    $log  = & $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1 | Out-String
    $code = $LASTEXITCODE
    Read-Run $out $log $code
}

# 保存先を付けずに動かす（-NoSave や、調べる前に止まるとき）
function Invoke-Plain([string[]]$ToolArgs, [string]$ToolPath) {
    if (-not $ToolPath) { $ToolPath = $tool }
    $log  = & $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1 | Out-String
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = (ConvertTo-FlatText $log) }
}

# 偽の型や関数を読み込んだラッパーから動かす。Mode は偽の Query の動き、Smb は Get-SmbOpenFile が返す値
function Invoke-Fake([string]$Target, [string]$Mode, [string]$SmbJson, [switch]$NoSave, [string]$ToolPath) {
    if (-not $ToolPath) { $ToolPath = $tool }
    $script:runNumber++
    $out = Join-Path $workRoot ('out_{0}' -f $script:runNumber)
    $env:FAKE_QUERY  = $Mode
    $env:FAKE_SMB    = $SmbJson
    $env:FAKE_NOSAVE = $(if ($NoSave) { '1' } else { $null })
    $log  = & $Exe -NoProfile -ExecutionPolicy Bypass -File $wrapper -Tool $ToolPath -Target $Target -Out $out 2>&1 | Out-String
    $code = $LASTEXITCODE
    $env:FAKE_QUERY  = $null
    $env:FAKE_SMB    = $null
    $env:FAKE_NOSAVE = $null
    Read-Run $out $log $code
}

# ツールの処理を 1 か所だけ書き換えたコピーを作る（本物では起こせない条件を確かめるため）
function New-PatchedTool([string]$Name, [string]$From, [string]$To) {
    $text = [IO.File]::ReadAllText($repoTool)
    if (-not $text.Contains($From)) { throw "書き換える場所が見つかりません: $From" }
    $file = Join-Path $workRoot "patched\$Name\Find-FileLock.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text.Replace($From, $To), $utf8Bom)
    $file
}

function Get-Clean($Value) { ([string]$Value) -replace '^="', '' -replace '"$', '' }

# 実際に掴むファイルを開く
function Open-Held([string]$FilePath, [IO.FileShare]$Share) {
    return [IO.File]::Open(('\\?\' + $FilePath), [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, $Share)
}

# アクセスの拒否を付ける。終わりに必ず外すため、付けた場所を覚えておく
function Add-Deny([string]$Target, [string]$Rights) {
    $result = & icacls.exe $Target /deny ('{0}:{1}' -f $me, $Rights) 2>&1
    if ($LASTEXITCODE -ne 0) { return "$result" }
    $aclPaths.Add($Target)
    return $null
}

function Remove-Deny([string]$Target) {
    [void](& icacls.exe $Target /remove:d $me 2>&1)
    [void]$aclPaths.Remove($Target)
}

# リンクを作る。消すときにリンクの先まで消さないよう（Directory.Delete の再帰は、中のジャンクションで例外を出す）、作った場所を覚えておく
function New-Link([string]$Kind, [string]$LinkPath, [string]$TargetPath) {
    # ファイルのシンボリックリンクは種類を付けない（空の引数を渡すと mklink の書式の誤りになる）
    $mklinkArgs = @('/d', '/c', 'mklink')
    if ($Kind) { $mklinkArgs += $Kind }
    $mklinkArgs += @($LinkPath, $TargetPath)
    [void](& cmd.exe @mklinkArgs 2>&1)
    $made = $false
    try { $made = ([IO.File]::GetAttributes($LinkPath) -band [IO.FileAttributes]::ReparsePoint) -ne 0 } catch { }
    if ($made) { $linkPaths.Add($LinkPath) }
    return $made
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

$held      = New-Object 'System.Collections.Generic.List[object]'
$processes = New-Object 'System.Collections.Generic.List[object]'
$aclPaths  = New-Object 'System.Collections.Generic.List[string]'
$linkPaths = New-Object 'System.Collections.Generic.List[string]'
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
if ($env:FAKE_NOSAVE) { & $Tool -Path $Target -NoSave } else { & $Tool -Path $Target -OutputFolder $Out }
exit $LASTEXITCODE
'@
    Write-Text $wrapper $wrapperCode

    Write-Text "$data\free.txt" 'x'
    Write-Text "$data\held_none.txt" 'x'
    Write-Text "$data\held_share.txt" 'x'
    Write-Text "$data\held_read.txt" 'x'
    Write-Text "$data\日本語 と スペース\掴み中 ファイル.txt" 'x'
    Write-Text "$data\[bracket]\held.txt" 'x'
    Write-Text "$data\tree\a\b\deep.txt" 'x'
    Write-Text "$data\tree\c\other.txt" 'x'
    Write-Text "$data\readonly.txt" 'x'
    [IO.File]::SetAttributes("$data\readonly.txt", [IO.FileAttributes]::ReadOnly)
    Write-Text "$data\-minus.txt" 'x'
    Write-Text "$data\offline.txt" 'x'
    Write-Text "$data\noaccess.txt" 'x'
    foreach ($i in 1..201) { Write-Text ("$data\many\f{0:000}.txt" -f $i) 'x' }

    #endregion

    #region 正常系

    Write-Section '使われていないファイル'
    $free = Invoke-Tool @('-Path', "$data\free.txt")
    Assert-Equal '終了コード' $free.Code '0'
    Assert-Equal '結果の行' (Test-Contains $free.Log '結果: 使用中のものは見つかりません') 'True'
    Assert-Equal '見つからなかった 2 つの欄は（なし）' (Get-Occurrence $free.Log '（なし）') '2'
    Assert-Equal '開けるかの確認は読み取り・書き込みとも可' ($free.Log -match ('可\s+可\s+' + [regex]::Escape("$data\free.txt"))) 'True'
    Assert-Equal 'locks.csv は見出しだけ' $free.Rows.Count '0'
    Assert-Equal 'summary.txt に開けるかの確認' ($free.Summary -match ('可\s+可\s+' + [regex]::Escape("$data\free.txt"))) 'True'

    $stampBefore = [IO.File]::GetLastWriteTimeUtc("$data\free.txt")
    Start-Sleep -Milliseconds 50
    $bytesBefore = [IO.File]::ReadAllText("$data\free.txt")
    [void](Invoke-Tool @('-Path', "$data\free.txt"))
    Assert-Equal '書き込みのモードで開いても更新日時が変わらない' ([IO.File]::GetLastWriteTimeUtc("$data\free.txt") -eq $stampBefore) 'True'
    Assert-Equal '内容も変わらない' ([IO.File]::ReadAllText("$data\free.txt")) $bytesBefore

    # -Path は名前を省略して書ける
    $positional = Invoke-Plain @("$data\free.txt", '-NoSave')
    Assert-Equal '-Path の名前を省略: 終了コード' $positional.Code '0'
    Assert-Equal '-Path の名前を省略: 対象' (Test-Contains $positional.Log ('対象: {0}\free.txt' -f $data)) 'True'

    Write-Section '排他で掴んでいるファイル'
    $held.Add((Open-Held "$data\held_none.txt" ([IO.FileShare]::None)))
    $lock = Invoke-Tool @('-Path', "$data\held_none.txt")
    Assert-Equal '終了コード' $lock.Code '2'
    Assert-Equal '結果の行' (Test-Contains $lock.Log '結果: 使用中   プロセス 1 個  クライアント 0 件') 'True'
    $mine = @($lock.Rows | Where-Object { $_.'PID' -eq "$PID" })
    Assert-Equal 'テストのプロセスが出る' $mine.Count '1'
    Assert-Equal 'ほかの行はない' $lock.Rows.Count '1'
    $myStart = (Get-Process -Id $PID).StartTime
    if ($mine.Count -gt 0) {
        Assert-Equal '種類' $mine[0].'種類' 'プロセス'
        Assert-Equal 'ファイル' (Get-Clean $mine[0].'ファイル') "$data\held_none.txt"
        Assert-Equal '読み取り' $mine[0].'読み取り' '不可（ほかのプログラムが使用中）'
        Assert-Equal '書き込み' $mine[0].'書き込み' '不可（ほかのプログラムが使用中）'
        Assert-Equal '実行ファイルが入る' ([bool]$mine[0].'実行ファイル') 'True'
        Assert-Equal '実行ユーザーが入る' ([bool]$mine[0].'実行ユーザー') 'True'
        Assert-Equal '起動日時の形（CSV）' $mine[0].'起動日時' $myStart.ToString('yyyy/MM/dd HH:mm:ss', $invariant)
        Assert-Equal 'クライアントの列は空欄' ($mine[0].'クライアント' + $mine[0].'クライアントのユーザー' + $mine[0].'権限') ''
    }
    Assert-Equal '起動日時の形（画面）' ($lock.Log -match ('(?<!\d)' + $PID + '\s+\S+\s+\S+\s+' + [regex]::Escape($myStart.ToString('yyyy-MM-dd HH:mm', $invariant)) + '\s')) 'True'
    Assert-Equal '自分のプロセスだけなら、管理者の案内は出さない' (Test-Contains $lock.Log $missingNote) 'False'

    # 表の列をそろえる（全角は 2 文字分の幅）。画面と同じ行を書く summary.txt で確かめる
    $sumLines  = @($lock.Summary -split '\r?\n')
    $headIndex = @(for ($i = 0; $i -lt $sumLines.Count; $i++) { if ($sumLines[$i] -match '^\s*PID\s') { $i } })[0]
    if ($null -ne $headIndex) {
        $headLine = $sumLines[$headIndex]
        $rowLine  = $sumLines[$headIndex + 1]
        Assert-Equal '表の列をそろえる（ファイルの列）' (Get-Width $headLine.Substring(0, $headLine.IndexOf('ファイル'))) (Get-Width $rowLine.Substring(0, $rowLine.IndexOf($data)))
        Assert-Equal '表の列をそろえる（PID は右寄せ）' ($headLine.IndexOf('PID') + 3) ($rowLine.IndexOf("$PID") + "$PID".Length)
    } else {
        Assert-Equal 'summary.txt にプロセスの表がある' 'なし' 'あり'
    }
    $checkIndex = @(for ($i = 0; $i -lt $sumLines.Count; $i++) { if ($sumLines[$i] -match '^\s*読み取り\s+書き込み\s+ファイル') { $i } })[0]
    if ($null -ne $checkIndex) {
        $headLine = $sumLines[$checkIndex]
        $rowLine  = $sumLines[$checkIndex + 1]
        Assert-Equal '表の列をそろえる（全角の値）' (Get-Width $headLine.Substring(0, $headLine.IndexOf('ファイル'))) (Get-Width $rowLine.Substring(0, $rowLine.IndexOf($data)))
    } else {
        Assert-Equal 'summary.txt に開けるかの確認の表がある' 'なし' 'あり'
    }

    Write-Section '共有ありで掴んでいるファイル（ロックではないが使用中）'
    $held.Add((Open-Held "$data\held_share.txt" ([IO.FileShare]'ReadWrite, Delete')))
    $share = Invoke-Tool @('-Path', "$data\held_share.txt")
    Assert-Equal '終了コード' $share.Code '2'
    $mine = @($share.Rows | Where-Object { $_.'PID' -eq "$PID" })
    Assert-Equal 'テストのプロセスが出る' $mine.Count '1'
    if ($mine.Count -gt 0) {
        Assert-Equal '読み取りは可' $mine[0].'読み取り' '可'
        Assert-Equal '書き込みは可' $mine[0].'書き込み' '可'
    }

    # 読み取りだけを許して掴んでいるファイル。書き込みだけが使用中になる
    $held.Add((Open-Held "$data\held_read.txt" ([IO.FileShare]::Read)))
    $readShare = Invoke-Tool @('-Path', "$data\held_read.txt")
    Assert-Equal '読み取りだけ共有: 終了コード' $readShare.Code '2'
    Assert-Equal '読み取りだけ共有: 開けるかの確認' (@($readShare.Rows | Where-Object { $_.'PID' -eq "$PID" } | ForEach-Object { $_.'読み取り' + '/' + $_.'書き込み' }) -join ',') '可/不可（ほかのプログラムが使用中）'
    $readUnknown = Invoke-Fake "$data\held_read.txt" 'empty' ''
    Assert-Equal '書き込みだけ使用中で、使っているものは不明: 終了コード 2' $readUnknown.Code '2'
    Assert-Equal '書き込みだけ使用中で、使っているものは不明: 結果の行' (Test-Contains $readUnknown.Log '結果: 使用中（使っているものは特定できません）') 'True'

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
        $exeRows = @($exeRun.Rows | Where-Object { $_.'PID' -eq "$($running.Id)" })
        Assert-Equal '実行中の EXE が出る' $exeRows.Count '1'
        if ($exeRows.Count -gt 0) {
            Assert-Equal 'プロセス' (Get-Clean $exeRows[0].'プロセス') 'ping_copy.exe'
            Assert-Equal '実行ファイル' (Get-Clean $exeRows[0].'実行ファイル') $exeCopy
            Assert-Equal 'コマンドラインは CSV に書く' ($exeRows[0].'コマンドライン' -like '*-n 60 127.0.0.1*') 'True'
        }
        Assert-Equal 'コマンドラインは画面に出さない' (Test-Contains $exeRun.Log '127.0.0.1') 'False'
        Assert-Equal 'コマンドラインは summary.txt に書かない' (Test-Contains $exeRun.Summary '127.0.0.1') 'False'
    }

    $dllRun = Invoke-Tool @('-Path', "$env:SystemRoot\System32\kernel32.dll")
    Assert-Equal '読み込み中の DLL で使用中になる' $dllRun.Code '2'
    Assert-Equal '読み込み中の DLL に複数のプロセスが出る' (@($dllRun.Rows).Count -gt 1) 'True'
    Assert-Equal '1 つのファイルでもプロセスごとに 1 行' (@($dllRun.Rows | ForEach-Object { $_.'PID' } | Sort-Object -Unique).Count) (@($dllRun.Rows).Count)
    Assert-Equal 'サービスの列が入る' (@($dllRun.Rows | Where-Object { $_.'サービス' -and $_.'PID' -ne '4' }).Count -gt 0) 'True'
    $multi = @($dllRun.Rows | Where-Object { $_.'サービス' -like '*,*' })
    if ($multi.Count -gt 0) {
        Assert-Equal '複数のサービスは「, 」で区切る' (@($multi | Where-Object { $_.'サービス' -notmatch '^[^,]+(, [^,]+)+$' }).Count) '0'
        $names = @($multi[0].'サービス' -split ', ')
        Assert-Equal '複数のサービスは名前の順' ($names -join '|') (@($names | Sort-Object) -join '|')
    } else {
        Write-Skip '複数のサービス' '1 つのプロセスで複数のサービスが動いていません'
    }
    # このツール自身は出さない（ツールのプロセスも kernel32.dll を読み込んでいる）
    Assert-Equal 'このツール自身は出さない' (@($dllRun.Rows | Where-Object { ([string]$_.'コマンドライン').Contains($tool) }).Count) '0'
    # ほかのユーザー（SYSTEM など）のプロセスの情報は、管理者でないと取れない
    Assert-Equal '管理者でないときだけ、取れなかった情報の案内' (Test-Contains $dllRun.Log $missingNote) "$(-not $isAdmin)"

    [void][IO.Directory]::CreateDirectory("$data\cwd")
    $cwdProcess = Start-Process -FilePath "$env:SystemRoot\System32\ping.exe" -ArgumentList '-n 60 127.0.0.1' -WorkingDirectory "$data\cwd" -PassThru -WindowStyle Hidden
    $processes.Add($cwdProcess)
    Start-Sleep -Milliseconds 500
    $cwdRun = Invoke-Tool @('-Path', "$data\cwd")
    Assert-Equal '現在のフォルダとして使っているプロセスが出る' (@($cwdRun.Rows | Where-Object { $_.'PID' -eq "$($cwdProcess.Id)" }).Count) '1'
    Assert-Equal 'フォルダの行は読み取りが空欄' (@($cwdRun.Rows | Where-Object { $_.'PID' -eq "$($cwdProcess.Id)" } | ForEach-Object { $_.'読み取り' + $_.'書き込み' }) -join ',') ''
    Assert-Equal 'フォルダは開けるかを確かめない' (Get-Occurrence $cwdRun.Log '（なし）') '2'

    Write-Section 'フォルダの中'
    $held.Add((Open-Held "$data\tree\a\b\deep.txt" ([IO.FileShare]::None)))
    $tree = Invoke-Tool @('-Path', "$data\tree")
    Assert-Equal '終了コード' $tree.Code '2'
    Assert-Equal '調べた数' (Test-Contains $tree.Log '調べた数: ファイル 2 件  フォルダ 4 件') 'True'
    Assert-Equal '深いファイルが出る' (@($tree.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq "$data\tree\a\b\deep.txt" }).Count -ge 1) 'True'
    Assert-Equal '使っていないファイルは出ない' (@($tree.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq "$data\tree\c\other.txt" }).Count) '0'
    Assert-Equal '使っているファイルだけ開けるかを確かめる' (Test-Contains $tree.Summary 'other.txt') 'False'
    Assert-Equal '使っているファイルの開けるかの確認' (@($tree.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq "$data\tree\a\b\deep.txt" } | ForEach-Object { $_.'書き込み' }) -join ',') '不可（ほかのプログラムが使用中）'

    Write-Section '名前と属性'
    $held.Add((Open-Held "$data\日本語 と スペース\掴み中 ファイル.txt" ([IO.FileShare]::None)))
    $jp = Invoke-Tool @('-Path', "$data\日本語 と スペース\掴み中 ファイル.txt")
    Assert-Equal '日本語とスペース' $jp.Code '2'
    Assert-Equal '日本語とスペース: ファイル' (@($jp.Rows | ForEach-Object { Get-Clean $_.'ファイル' }) -join ',') "$data\日本語 と スペース\掴み中 ファイル.txt"
    $held.Add((Open-Held "$data\[bracket]\held.txt" ([IO.FileShare]::None)))
    $bracket = Invoke-Tool @('-Path', "$data\[bracket]")
    Assert-Equal '[ ] を含むフォルダ' $bracket.Code '2'
    Assert-Equal '[ ] を含むフォルダ: ファイル' (@($bracket.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq "$data\[bracket]\held.txt" }).Count) '1'
    $readonly = Invoke-Tool @('-Path', "$data\readonly.txt")
    Assert-Equal '読み取り専用は使用中ではない' $readonly.Code '0'
    Assert-Equal '読み取り専用の表示' ($readonly.Log -match ('可\s+不可（読み取り専用）\s+' + [regex]::Escape("$data\readonly.txt"))) 'True'
    $held.Add((Open-Held "$data\-minus.txt" ([IO.FileShare]::None)))
    $minus = Invoke-Tool @('-Path', "$data\-minus.txt")
    Assert-Equal '- で始まる名前も調べる' $minus.Code '2'

    # オンライン専用ファイル（オフライン属性）は、開くとダウンロードが始まるので開かない
    [IO.File]::SetAttributes("$data\offline.txt", [IO.FileAttributes]::Offline)
    $held.Add((Open-Held "$data\offline.txt" ([IO.FileShare]::None)))
    $offline = Invoke-Tool @('-Path', "$data\offline.txt")
    Assert-Equal 'オンライン専用: 終了コード' $offline.Code '2'
    Assert-Equal 'オンライン専用: 読み取り・書き込み' (@($offline.Rows | ForEach-Object { $_.'読み取り' + '/' + $_.'書き込み' }) -join ',') '確かめていない/確かめていない'
    # 問い合わせが共有違反になったオンライン専用ファイルは、開いて確かめられないので、調べられなかったものとして残す
    $offlineSharing = Invoke-Fake "$data\offline.txt" 'sharing' ''
    Assert-Equal 'オンライン専用で共有違反: 終了コード 1' $offlineSharing.Code '1'
    Assert-Equal 'オンライン専用で共有違反: skipped.csv' (@($offlineSharing.Skips | ForEach-Object { $_.'詳細' }) -join ',') '問い合わせが共有違反になり、開けるかも確かめられません（オンライン専用ファイル）'

    # 開けるかの確認の「権限なし」
    $denyText = Add-Deny "$data\noaccess.txt" '(RD,WD)'
    if (-not $denyText) {
        $noAccess = Invoke-Tool @('-Path', "$data\noaccess.txt")
        Assert-Equal '開けない権限: 使用中ではない' $noAccess.Code '0'
        Assert-Equal '開けない権限: 表示' ($noAccess.Log -match ('不可（権限なし）\s+不可（権限なし）\s+' + [regex]::Escape("$data\noaccess.txt"))) 'True'
        # 問い合わせが共有違反になったのに、開けるかも確かめられなければ、調べられなかったものとして残す
        $sharingDenied = Invoke-Fake "$data\noaccess.txt" 'sharing' ''
        Assert-Equal '共有違反で確かめられない: 終了コード 1' $sharingDenied.Code '1'
        Assert-Equal '共有違反で確かめられない: skipped.csv' (@($sharingDenied.Skips | ForEach-Object { $_.'理由' + ':' + $_.'詳細' }) -join ',') 'その他のエラー:問い合わせが共有違反になり、開けるかも確かめられません（権限なし）'
        Remove-Deny "$data\noaccess.txt"
    } else {
        Write-Skip '開けない権限' "icacls に失敗: $denyText"
    }

    # 260 文字を超えるパス
    $longDir = "$data\long\" + (('d' * 50 + '\') * 6)
    $longFile = $longDir + 'held.txt'
    [void][IO.Directory]::CreateDirectory('\\?\' + $longDir)
    [IO.File]::WriteAllText('\\?\' + $longFile, 'x')
    $held.Add((Open-Held $longFile ([IO.FileShare]::None)))
    $long = Invoke-Tool @('-Path', "$data\long")
    Assert-Equal '260 文字を超えるパス' $long.Code '2'
    Assert-Equal '長いパスは \\?\ なしで書く' (@($long.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq $longFile }).Count -ge 1) 'True'
    Assert-Equal '長いパスの開けるかの確認' (@($long.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq $longFile } | ForEach-Object { $_.'書き込み' }) -join ',') '不可（ほかのプログラムが使用中）'
    # 長いファイルを直接指定し、長い保存先に保存する
    $longOut = $longDir + 'out'
    $longDirect = Invoke-Plain @('-Path', $longFile, '-OutputFolder', $longOut)
    Assert-Equal '長いファイルを直接指定: 終了コード' $longDirect.Code '2'
    Assert-Equal '長いファイルを直接指定: 対象' (Test-Contains $longDirect.Log ('対象: ' + $longFile)) 'True'
    $longRuns = @()
    if ([IO.Directory]::Exists('\\?\' + $longOut)) { $longRuns = @([IO.Directory]::GetDirectories('\\?\' + $longOut, 'FileLock_*')) }
    Assert-Equal '長い保存先に保存する' $longRuns.Count '1'
    if ($longRuns.Count -gt 0) {
        Assert-Equal '長い保存先: 3 つのファイル' (@('summary.txt', 'locks.csv', 'skipped.csv' | Where-Object { [IO.File]::Exists($longRuns[0] + '\' + $_) }).Count) '3'
        Assert-Equal '長い保存先: locks.csv の行' ([IO.File]::ReadAllLines($longRuns[0] + '\locks.csv').Count) '2'
        Assert-Equal '長い保存先: 保存先の表示は \\?\ なし' (Test-Contains $longDirect.Log ('保存先: ' + $longOut + '\FileLock_')) 'True'
    }

    Write-Section 'リンク'
    [void][IO.Directory]::CreateDirectory("$data\junction\real")
    Write-Text "$data\jtarget\inside.txt" 'x'
    if (New-Link '/J' "$data\junction\link" "$data\jtarget") {
        $held.Add((Open-Held "$data\jtarget\inside.txt" ([IO.FileShare]::None)))
        $junction = Invoke-Tool @('-Path', "$data\junction")
        Assert-Equal 'リンクの先はたどらない' $junction.Code '0'
        Assert-Equal 'たどらなかったリンクの件数' (Test-Contains $junction.Log 'たどらなかったリンク 1 件') 'True'
        Assert-Equal '調べた数にリンクを数えない' (Test-Contains $junction.Log '調べた数: ファイル 0 件  フォルダ 2 件') 'True'
        Assert-Equal 'skipped.csv に残す' (@($junction.Skips | ForEach-Object { $_.'理由' + ':' + (Split-Path -Leaf (Get-Clean $_.'パス')) + ':' + $_.'詳細' }) -join ',') 'たどらなかったリンク:link:ジャンクション'
        Assert-Equal 'リンクはスキップの件数に入れない' (Test-Contains $junction.Log 'スキップ') 'False'
        # -Path にリンクそのものを指定したときは、リンク先を調べる
        $direct = Invoke-Tool @('-Path', "$data\junction\link")
        Assert-Equal '-Path のリンクの先を調べる' $direct.Code '2'
    } else {
        Write-Skip 'リンク' 'ジャンクションを作れません'
    }

    # Windows の互換用リンク（隠し・システム属性のジャンクション）と、シンボリックリンク
    [void][IO.Directory]::CreateDirectory("$data\links")
    [void][IO.Directory]::CreateDirectory("$data\ltarget")
    Write-Text "$data\ltarget\file.txt" 'x'
    $expectedLinks = New-Object 'System.Collections.Generic.List[string]'
    if (New-Link '/J' "$data\links\compat" "$data\ltarget") {
        [void](& attrib.exe +h +s /L "$data\links\compat" 2>&1)
        $expectedLinks.Add('compat:Windows の互換用リンク')
    }
    $symDir  = New-Link '/D' "$data\links\symdir" "$data\ltarget"
    $symFile = New-Link '' "$data\links\symfile.txt" "$data\ltarget\file.txt"
    if ($symDir) { $expectedLinks.Add('symdir:シンボリックリンク') }
    if ($symFile) { $expectedLinks.Add('symfile.txt:シンボリックリンク') }
    if ($expectedLinks.Count -gt 0) {
        $linkRun = Invoke-Tool @('-Path', "$data\links")
        $linkText = @($linkRun.Skips | ForEach-Object { (Split-Path -Leaf (Get-Clean $_.'パス')) + ':' + $_.'詳細' } | Sort-Object) -join ','
        Assert-Equal 'リンクの種類' $linkText (@($expectedLinks | Sort-Object) -join ',')
        Assert-Equal 'リンクは調べた数に入れない' (Test-Contains $linkRun.Log '調べた数: ファイル 0 件  フォルダ 1 件') 'True'
    }
    if (-not $symDir -or -not $symFile) { Write-Skip 'シンボリックリンク' '作る権限がありません（管理者か開発者モードで確かめます）' }

    #endregion

    #region 異常系

    Write-Section '止まるとき'
    $none = Invoke-Tool @('-Path', "$data\nothing.txt")
    Assert-Equal '見つからない: 終了コード' $none.Code '1'
    Assert-Equal '見つからない: メッセージ' (Test-Contains $none.Log ('ファイルもフォルダも見つかりません（または、アクセスする権限がありません）: {0}\nothing.txt' -f $data)) 'True'
    Assert-Equal '見つからない: 結果を作らない' $none.Dir ''
    $noPath = Invoke-Plain @()
    Assert-Equal '-Path なし: 終了コード' $noPath.Code '1'
    Assert-Equal '-Path なし: メッセージ' (Test-Contains $noPath.Log '調べるファイルかフォルダを -Path で指定してください') 'True'
    $emptyPath = Invoke-Plain @('-Path', '""')
    Assert-Equal '-Path が引用符だけ: 終了コード' $emptyPath.Code '1'

    # 共有フォルダを開けないときは、確かめる順番を案内する（ループバックの名前なので、外へは通信しない）
    $noShare = Invoke-Plain @('-Path', '\\localhost\wbtest-noshare\x.txt', '-NoSave')
    Assert-Equal '開けない共有フォルダ: 終了コード' $noShare.Code '1'
    Assert-Equal '開けない共有フォルダ: 案内' (Test-Contains $noShare.Log '共有フォルダのときは、次の順に確かめてください。') 'True'
    Assert-Equal '開けない共有フォルダ: 管理者のときの案内' (Test-Contains $noShare.Log '※ 管理者として実行しています。') "$isAdmin"

    $cmdFile = "$workRoot\quote.cmd"
    Write-Text $cmdFile ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -Path ""{2}\"" -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $Exe, $tool, $data)
    $quote = ConvertTo-FlatText (& cmd.exe /c $cmdFile 2>&1 | Out-String)
    Assert-Equal '\" で引数を巻き込む: 終了コード' $LASTEXITCODE '1'
    Assert-Equal '\" で引数を巻き込む: ヒント' (Test-Contains $quote '-Path の値に引用符（"）が残っています') 'True'

    # コマンドプロンプトでは ' がそのまま届く
    $singleFile = "$workRoot\single.cmd"
    Write-Text $singleFile ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -Path '{2}\free.txt' -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $Exe, $tool, $data)
    $single = ConvertTo-FlatText (& cmd.exe /c $singleFile 2>&1 | Out-String)
    Assert-Equal 'cmd の '' で囲んだパス: 終了コード' $LASTEXITCODE '0'
    Assert-Equal 'cmd の '' で囲んだパス: 対象' (Test-Contains $single ('対象: {0}\free.txt' -f $data)) 'True'

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
        Assert-Equal 'ルートの中をすべて調べる' (Test-Contains $rootLog '調べた数: ファイル 2 件  フォルダ 4 件') 'True'
    } else {
        Write-Skip 'ドライブ名だけの値' 'subst でドライブを作れません'
    }

    # 制約付き言語モードでは、最初に止まる（レビューの指摘）。
    # __PSLockdownPolicy は 5.1 だとコンピューター全体の環境変数しか効かないので、起動した PowerShell の中で切り替えてから呼ぶ
    $clmCommand = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '{0}' -Path '{1}' -NoSave; exit `$LASTEXITCODE" -f $tool, "$data\free.txt"
    $clm = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -Command $clmCommand 2>&1 | Out-String)
    Assert-Equal '制約付き言語モード: 終了コード' $LASTEXITCODE '1'
    Assert-Equal '制約付き言語モード: メッセージ' (Test-Contains $clm 'Windows の API を呼ぶ準備ができません。制約付き言語モードか、コンパイラーの実行が禁止されている環境では使えません') 'True'
    Assert-Equal '制約付き言語モード: 最初の確認で止まる' (Test-Contains $clm '使えません（') 'False'

    # Add-Type が使えないとき（コンパイラーの実行が禁止されている環境）。.NET のエラーの文は途中や末尾に改行が入ることがあるので、1 行にまとめる
    $noCompile = New-PatchedTool 'nocompile' 'Add-Type -TypeDefinition $nativeSource -IgnoreWarnings -WarningAction SilentlyContinue -ErrorAction Stop' "throw ""compile`r`n  failed`r`n"""
    $compile = Invoke-Plain @('-Path', "$data\free.txt", '-NoSave') $noCompile
    Assert-Equal 'Add-Type の失敗: 終了コード' $compile.Code '1'
    Assert-Equal 'Add-Type の失敗: 元のエラーを 1 行で出す' (Test-Contains $compile.Log 'コンパイラーの実行が禁止されている環境では使えません（compile failed）') 'True'

    Write-Section '保存'
    $saved = Invoke-Tool @('-Path', "$data\free.txt")
    foreach ($name in 'summary.txt', 'locks.csv', 'skipped.csv') {
        $bytes = [IO.File]::ReadAllBytes((Join-Path $saved.Dir $name))
        Assert-Equal ("$name は BOM 付き") ('{0}-{1}-{2}' -f $bytes[0], $bytes[1], $bytes[2]) '239-187-191'
    }
    Assert-Equal '保存先のフォルダ名' ((Split-Path -Leaf $saved.Dir) -match ('^FileLock_' + [regex]::Escape($env:COMPUTERNAME) + '_\d{8}_\d{6}$')) 'True'
    Assert-Equal '保存先の表示' (Test-Contains $saved.Log ('保存先: ' + $saved.Dir)) 'True'
    $header = [IO.File]::ReadAllLines((Join-Path $saved.Dir 'locks.csv'))[0]
    Assert-Equal 'locks.csv の列' $header '"種類","ファイル","PID","プロセス","サービス","実行ユーザー","起動日時","実行ファイル","コマンドライン","クライアント","クライアントのユーザー","権限","読み取り","書き込み"'
    Assert-Equal 'skipped.csv の列' ([IO.File]::ReadAllLines((Join-Path $saved.Dir 'skipped.csv'))[0]) '"理由","パス","詳細"'
    Assert-Equal 'summary.txt の最後は保存先' (@([IO.File]::ReadAllLines((Join-Path $saved.Dir 'summary.txt')))[-1]) ('保存先: ' + $saved.Dir)
    $noSave = Invoke-Plain @('-Path', "$data\free.txt", '-NoSave')
    Assert-Equal '-NoSave の終了コード' $noSave.Code '0'
    Assert-Equal '-NoSave は保存先を出さない' (Test-Contains $noSave.Log '保存先:') 'False'
    Assert-Equal '-NoSave は結果を作らない' (@(Get-ChildItem -LiteralPath (Split-Path $tool) -Directory -Filter 'FileLock_*').Count) '0'
    Write-Text "$workRoot\notdir" 'x'
    $unwritable = Invoke-Plain @('-Path', "$data\free.txt", '-OutputFolder', "$workRoot\notdir")
    Assert-Equal '書き込めない保存先: 終了コード' $unwritable.Code '0'
    Assert-Equal '書き込めない保存先: 警告' (Test-Contains $unwritable.Log ('保存先に書き込めないため、結果は保存しません: {0}\notdir' -f $workRoot)) 'True'
    Assert-Equal '書き込めない保存先: 画面には結果を出す' (Test-Contains $unwritable.Log '結果: 使用中のものは見つかりません') 'True'
    # -OutputFolder を省略すると、スクリプトと同じフォルダに保存する
    $defaultOut = Invoke-Plain @('-Path', "$data\free.txt")
    Assert-Equal '既定の保存先はスクリプトのフォルダ' (@(Get-ChildItem -LiteralPath (Split-Path $tool) -Directory -Filter 'FileLock_*').Count) '1'

    # 保存の途中で失敗したときは、作ったフォルダを消し、画面にはすべて出す
    $saveFail = New-PatchedTool 'savefail' "Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'skipped.csv')) `$SkipColumns `$skipRows" "throw 'disk full'"
    $failRun = Invoke-Fake "$data\free.txt" 'many' '' -ToolPath $saveFail
    Assert-Equal '保存の失敗: 警告' (Test-Contains $failRun.Log '結果を保存できませんでした: disk full') 'True'
    Assert-Equal '保存の失敗: 作ったフォルダを消す' $failRun.Dir ''
    Assert-Equal '保存の失敗: 画面にはすべて出す' (Get-Occurrence $failRun.Log '（終了したプロセス）') '300'
    Assert-Equal '保存の失敗: 保存先を出さない' (Test-Contains $failRun.Log '保存先:') 'False'
    Assert-Equal '保存の失敗: 終了コードは結果のとおり' $failRun.Code '2'

    # 想定外のエラーでは、作ったフォルダを消し、中断とは別の言葉で知らせる
    $crash = New-PatchedTool 'crash' "Write-Progress -Activity `$ToolTitle -Status '結果を保存しています'" "throw 'unexpected'"
    $crashRun = Invoke-Tool @('-Path', "$data\free.txt") $crash
    Assert-Equal '想定外のエラー: 終了コード' ($crashRun.Code -ne 0) 'True'
    Assert-Equal '想定外のエラー: 表示' (Test-Contains $crashRun.Log 'エラーで止まりました。結果は保存していません。') 'True'
    Assert-Equal '想定外のエラー: 作ったフォルダを消す' $crashRun.Dir ''

    Write-Section '読めないフォルダ'
    # 中の一覧だけを拒否したフォルダ（問い合わせは通る）と、読み取りをすべて拒否したフォルダ（問い合わせも中の一覧も拒否される）
    $aclList = "$data\aclmix\denied_list"
    $aclAll  = "$data\aclmix\denied_all"
    Write-Text "$aclList\secret.txt" 'x'
    Write-Text "$aclAll\secret.txt" 'x'
    Write-Text "$data\aclmix\held.txt" 'x'
    $denied = Add-Deny $aclList '(OI)(CI)(RD)'
    if (-not $denied) { $denied = Add-Deny $aclAll '(OI)(CI)(R)' }
    if (-not $denied) {
        $denyList = Invoke-Tool @('-Path', $aclList)
        Assert-Equal '一覧の権限なし: 止まらず終了コード 1' $denyList.Code '1'
        Assert-Equal '一覧の権限なし: skipped.csv' (@($denyList.Skips | ForEach-Object { $_.'理由' + ':' + (Get-Clean $_.'パス') }) -join ',') ('権限なし:' + $aclList)
        Assert-Equal '一覧の権限なし: 件数の表示' (Test-Contains $denyList.Log 'スキップ 1 件（権限なし 1）') 'True'
        $denyAll = Invoke-Tool @('-Path', $aclAll)
        Assert-Equal '読み取りの権限なし: 止まらず終了コード 1' $denyAll.Code '1'
        Assert-Equal '読み取りの権限なし: 問い合わせと一覧の両方で拒否されても 1 件' (@($denyAll.Skips | ForEach-Object { $_.'理由' + ':' + (Get-Clean $_.'パス') }) -join ',') ('権限なし:' + $aclAll)
        Assert-Equal '読み取りの権限なし: 件数の表示' (Test-Contains $denyAll.Log 'スキップ 1 件（権限なし 1）') 'True'
        # 使っているものが見つかれば、読めない場所があっても 2
        $held.Add((Open-Held "$data\aclmix\held.txt" ([IO.FileShare]::None)))
        $mixed = Invoke-Tool @('-Path', "$data\aclmix")
        Assert-Equal '見つかったうえに読めない場所: 終了コード 2' $mixed.Code '2'
        Assert-Equal '見つかったうえに読めない場所: 件数の表示' (Test-Contains $mixed.Log 'スキップ 2 件（権限なし 2）') 'True'
        Assert-Equal '見つかったうえに読めない場所: skipped.csv' (@($mixed.Skips | ForEach-Object { Split-Path -Leaf (Get-Clean $_.'パス') } | Sort-Object) -join ',') 'denied_all,denied_list'
        Remove-Deny $aclList
        Remove-Deny $aclAll
    } else {
        Write-Skip '読めないフォルダ' "icacls に失敗: $denied"
    }

    #endregion

    #region 偽の API と偽の共有

    Write-Section 'API の戻り値'
    $empty = Invoke-Fake "$data\held_none.txt" 'empty' ''
    Assert-Equal '見つからないが共有違反: 終了コード 2' $empty.Code '2'
    Assert-Equal '見つからないが共有違反: 結果の行' (Test-Contains $empty.Log '結果: 使用中（使っているものは特定できません）') 'True'
    Assert-Equal '見つからないが共有違反: 案内' (Test-Contains $empty.Log '※ 使っているものは見つかりませんが、書き込みで開けないファイルがあります。OS やドライバーが使っているか（pagefile.sys など）、別の PC から共有フォルダ経由で使われているのかもしれません。共有フォルダのファイルなら、ファイルサーバーの上で実行してください') 'True'
    Assert-Equal '見つかったときは、見つからないときの案内を出さない' (Test-Contains $lock.Log '使っているものは見つかりませんが') 'False'
    $gone = Invoke-Fake "$data\free.txt" 'gone' ''
    Assert-Equal '調べている間に消えた: スキップに数えない' $gone.Code '0'
    Assert-Equal '調べている間に消えた: skipped.csv にも書かない' $gone.Skips.Count '0'
    $apiDenied = Invoke-Fake "$data\free.txt" 'denied' ''
    Assert-Equal '権限なし: 終了コード 1' $apiDenied.Code '1'
    Assert-Equal '権限なし: 理由' (@($apiDenied.Skips | ForEach-Object { $_.'理由' }) -join ',') '権限なし'
    $apiError = Invoke-Fake "$data\free.txt" 'error' ''
    Assert-Equal 'その他のエラー: 終了コード 1' $apiError.Code '1'
    Assert-Equal 'その他のエラー: 理由' (@($apiError.Skips | ForEach-Object { $_.'理由' }) -join ',') 'その他のエラー'
    Assert-Equal 'その他のエラー: 件数の表示' (Test-Contains $apiError.Log 'スキップ 1 件（その他のエラー 1）') 'True'
    Assert-Equal 'その他のエラー: 詳細は 1 行' (@($apiError.Skips | Where-Object { $_.'詳細' -and $_.'詳細' -notmatch '[\r\n]' }).Count) '1'
    $sharing = Invoke-Fake "$data\held_none.txt" 'sharing' ''
    Assert-Equal '問い合わせが共有違反: 開けるかの確認で使用中' $sharing.Code '2'
    Assert-Equal '問い合わせが共有違反: スキップに数えない' $sharing.Skips.Count '0'
    $sharingFree = Invoke-Fake "$data\free.txt" 'sharing' ''
    Assert-Equal '問い合わせが共有違反でも開ければ使用中ではない' $sharingFree.Code '0'
    $linkMode = Invoke-Fake "$data\free.txt" 'link' ''
    Assert-Equal '-Path はリンクをたどって開く' (@($linkMode.Rows | Where-Object { $_.'PID' -eq '4' }).Count) '1'
    Assert-Equal 'PID 4 の注記' (Test-Contains $linkMode.Log '（カーネル：共有フォルダ経由の接続など）') 'True'
    Assert-Equal 'PID 4 の注記は CSV のサービスの列' (@($linkMode.Rows | ForEach-Object { $_.'サービス' }) -join ',') '（カーネル：共有フォルダ経由の接続など）'
    Assert-Equal 'PID 4 だけなら、管理者の案内は出さない' (Test-Contains $linkMode.Log $missingNote) 'False'
    $many = Invoke-Fake "$data\free.txt" 'many' ''
    Assert-Equal '画面は 200 行まで' (Test-Contains $many.Log '（ほか 100 行は locks.csv を参照してください）') 'True'
    Assert-Equal '画面に出す行の数' (Get-Occurrence $many.Log '（終了したプロセス）') '200'
    Assert-Equal 'CSV はすべて' $many.Rows.Count '300'
    Assert-Equal 'summary.txt はすべて' (Get-Occurrence $many.Summary '（終了したプロセス）') '300'
    Assert-Equal '終了したプロセスの表示' (Test-Contains $many.Log '（終了したプロセス）') 'True'
    Assert-Equal 'プロセスの数' (Test-Contains $many.Log '結果: 使用中   プロセス 300 個  クライアント 0 件') 'True'
    Assert-Equal '終了したプロセスだけなら、管理者の案内は出さない' (Test-Contains $many.Log $missingNote) 'False'
    Assert-Equal 'PID の順に並べる' (@($many.Rows | Select-Object -First 3 | ForEach-Object { $_.'PID' }) -join ',') '100000,100001,100002'
    $manyNoSave = Invoke-Fake "$data\free.txt" 'many' '' -NoSave
    Assert-Equal '-NoSave なら画面にすべて出す' (Get-Occurrence $manyNoSave.Log '（終了したプロセス）') '300'
    Assert-Equal '-NoSave なら省略の案内を出さない' (Test-Contains $manyNoSave.Log '（ほか') 'False'

    # 開けるかの確認の欄も 200 行まで。超えた分は summary.txt を見てもらう
    $manyFiles = Invoke-Fake "$data\many" 'link' ''
    Assert-Equal 'フォルダの中の 201 ファイル: 調べた数' (Test-Contains $manyFiles.Log '調べた数: ファイル 201 件  フォルダ 1 件') 'True'
    Assert-Equal 'プロセスの欄は 200 行まで' (Test-Contains $manyFiles.Log '（ほか 2 行は locks.csv を参照してください）') 'True'
    Assert-Equal '開けるかの確認の欄は 200 行まで' (Test-Contains $manyFiles.Log '（ほか 1 行は summary.txt を参照してください）') 'True'
    Assert-Equal 'summary.txt は開けるかの確認もすべて' ([regex]::Matches($manyFiles.Summary, ('可\s+可\s+' + [regex]::Escape("$data\many\f"))).Count) '201'
    Assert-Equal 'locks.csv はすべて（フォルダの行を含む）' $manyFiles.Rows.Count '202'
    Assert-Equal 'フォルダの行は開けるかが空欄' (@($manyFiles.Rows | Where-Object { (Get-Clean $_.'ファイル') -eq "$data\many" } | ForEach-Object { $_.'読み取り' + $_.'書き込み' }) -join ',') ''

    Write-Section '共有フォルダの接続'
    $d = $data.Replace('\', '\\')
    $smbJson = '[{"Path":"' + $d + '\\free.txt","ClientComputerName":"10.0.0.25","ClientUserName":"CONTOSO\\suzuki","Permissions":1179785},' +
        '{"Path":"' + $d + '\\free.txt","ClientComputerName":"10.0.0.25","ClientUserName":"CONTOSO\\suzuki","Permissions":1180063},' +
        '{"Path":"' + $d + '\\free.txt2","ClientComputerName":"10.0.0.99","ClientUserName":"x","Permissions":1}]'
    $smb = Invoke-Fake "$data\free.txt" 'empty' $smbJson
    Assert-Equal '接続があれば使用中' $smb.Code '2'
    Assert-Equal '接続の数' (Test-Contains $smb.Log '結果: 使用中   プロセス 0 個  クライアント 1 件') 'True'
    $clientRows = @($smb.Rows | Where-Object { $_.'種類' -eq 'クライアント' })
    Assert-Equal '同じ接続は 1 行にまとめ、似た名前は含めない' $clientRows.Count '1'
    if ($clientRows.Count -gt 0) {
        Assert-Equal 'クライアント（Excel で数値にならない形）' $clientRows[0].'クライアント' '="10.0.0.25"'
        Assert-Equal 'ユーザー' $clientRows[0].'クライアントのユーザー' 'CONTOSO\suzuki'
        Assert-Equal '権限は強いほう' $clientRows[0].'権限' '読み書き'
        Assert-Equal 'クライアントの行にも開けるかの確認' ($clientRows[0].'読み取り' + '/' + $clientRows[0].'書き込み') '可/可'
        Assert-Equal 'プロセスの列は空欄' ($clientRows[0].'PID' + $clientRows[0].'プロセス' + $clientRows[0].'実行ファイル' + $clientRows[0].'コマンドライン') ''
    }
    Assert-Equal '画面のクライアントの行' ($smb.Log -match ('10\.0\.0\.25\s+CONTOSO\\suzuki\s+読み書き\s+' + [regex]::Escape("$data\free.txt"))) 'True'

    # 権限の判定（読み取りだけ・書き込みだけ・どちらでもない）と、弱い権限から強い権限へのまとめ
    $smbAccess = '[{"Path":"' + $d + '\\free.txt","ClientComputerName":"pc1","ClientUserName":"u1","Permissions":128},' +
        '{"Path":"' + $d + '\\free.txt","ClientComputerName":"pc1","ClientUserName":"u1","Permissions":1},' +
        '{"Path":"' + $d + '\\free.txt","ClientComputerName":"pc2","ClientUserName":"u2","Permissions":2},' +
        '{"Path":"' + $d + '\\free.txt","ClientComputerName":"pc3","ClientUserName":"u3","Permissions":1048576},' +
        '{"Path":"' + $d + '\\free.txt","ClientComputerName":"pc4","ClientUserName":"u4","Permissions":1179785},' +
        '{"Path":"' + $d + '\\free.txt","ClientComputerName":"pc4","ClientUserName":"u4","Permissions":128}]'
    $access = Invoke-Fake "$data\free.txt" 'empty' $smbAccess
    Assert-Equal '権限の判定' (@($access.Rows | ForEach-Object { $_.'クライアント' + ':' + $_.'権限' }) -join ',') 'pc1:読み取り,pc2:読み書き,pc3:その他,pc4:読み取り'

    $smbTree = '[{"Path":"' + $d + '\\tree\\c\\other.txt","ClientComputerName":"pc1","ClientUserName":"u","Permissions":1},{"Path":"' + $d + '\\tree2\\x.txt","ClientComputerName":"pc2","ClientUserName":"u","Permissions":1}]'
    $smbFolder = Invoke-Fake "$data\tree" 'empty' $smbTree
    Assert-Equal 'フォルダの中の接続だけを出す' (@($smbFolder.Rows | Where-Object { $_.'種類' -eq 'クライアント' } | ForEach-Object { Get-Clean $_.'ファイル' }) -join ',') "$data\tree\c\other.txt"
    Assert-Equal '接続だけで見つかったファイルも開けるかを確かめる' (@($smbFolder.Rows | Where-Object { $_.'種類' -eq 'クライアント' } | ForEach-Object { $_.'読み取り' + '/' + $_.'書き込み' }) -join ',') '可/可'
    if ($script:substUsed) {
        # ドライブのルートの接続は、\ を外すと D: になるので元に戻す
        $smbRoot = '[{"Path":"' + $script:substUsed + ':\\","ClientComputerName":"pc9","ClientUserName":"u","Permissions":1}]'
        $rootRun = Invoke-Fake ($script:substUsed + ':\') 'empty' $smbRoot
        Assert-Equal 'ドライブのルートの接続' (@($rootRun.Rows | Where-Object { $_.'種類' -eq 'クライアント' } | ForEach-Object { Get-Clean $_.'ファイル' }) -join ',') ($script:substUsed + ':\')
    }
    $smbThrow = Invoke-Fake "$data\free.txt" 'empty' 'throw'
    Assert-Equal '調べられなくても止まらない' $smbThrow.Code '0'
    $noteExpected = if ($isAdmin) { '※ 共有フォルダの接続元を調べられませんでした（denied）' } else { '※ 共有フォルダの接続元は、管理者として実行すると調べられます' }
    Assert-Equal '調べられなかった案内' (Test-Contains $smbThrow.Log $noteExpected) 'True'
    $noSmb = New-PatchedTool 'nosmb' '(Get-Command Get-SmbOpenFile -ErrorAction SilentlyContinue)' '(Get-Command Get-NoSuchSmbOpenFile -ErrorAction SilentlyContinue)'
    $noSmbRun = Invoke-Plain @('-Path', "$data\free.txt", '-NoSave') $noSmb
    Assert-Equal 'Get-SmbOpenFile がない: 終了コード' $noSmbRun.Code '0'
    Assert-Equal 'Get-SmbOpenFile がない: 案内' (Test-Contains $noSmbRun.Log '※ 共有フォルダの接続元を調べられませんでした（Get-SmbOpenFile が見つかりません）') 'True'

    Write-Section '共有フォルダのパス'
    # ネットワーク上のパスと判定されたとき（本物は管理者のときだけ下で確かめる）
    $network = New-PatchedTool 'network' 'function Test-NetworkPath([string]$FullPath) {' 'function Test-NetworkPath([string]$FullPath) { return $true'
    $netFree = Invoke-Fake "$data\free.txt" 'empty' 'throw' -ToolPath $network
    Assert-Equal 'ネットワーク上のパスで何も見つからなければ 1' $netFree.Code '1'
    Assert-Equal 'ネットワーク上のパス: 案内' (Test-Contains $netFree.Log '※ 別の PC からは、この PC のプロセスしか分かりません。') 'True'
    Assert-Equal 'ネットワーク上のパス: 共有フォルダの接続は調べない' (Test-Contains $netFree.Log '共有フォルダの接続元') 'False'
    $netHeld = Invoke-Tool @('-Path', "$data\held_none.txt") $network
    Assert-Equal 'ネットワーク上のパスでも、見つかれば 2' $netHeld.Code '2'
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
        Write-Skip '本物の共有フォルダのパス' '管理者のときだけ確かめます'
    }

    #endregion

    #region 実機の条件

    Write-Section '実機の条件'
    # ページファイルは、属性の読み取りでも共有違反になる。使っているプロセスは分からないが、使用中とする
    $pagefile = "$env:SystemDrive\pagefile.sys"
    if ([IO.File]::Exists($pagefile)) {
        $pf = Invoke-Plain @('-Path', $pagefile, '-NoSave')
        Assert-Equal 'pagefile.sys: 終了コード 2' $pf.Code '2'
        Assert-Equal 'pagefile.sys: 結果の行' (Test-Contains $pf.Log '結果: 使用中（使っているものは特定できません）') 'True'
        Assert-Equal 'pagefile.sys: 開けるかの確認' ($pf.Log -match ('不可（ほかのプログラムが使用中）\s+不可（ほかのプログラムが使用中）\s+' + [regex]::Escape($pagefile))) 'True'
    } else {
        Write-Skip 'pagefile.sys' 'ページファイルがありません'
    }

    #endregion

    #region 起動

    Write-Section '起動'
    # 名前に [ ] を含むフォルダから、相対パスで起動する
    Write-Text "$workRoot\[br]\a.txt" 'x'
    $brBat = "$workRoot\bracket.cmd"
    Write-Text $brBat ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\[br]""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -Path a.txt -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $workRoot, $Exe, $tool)
    $brLog = ConvertTo-FlatText (& cmd.exe /c $brBat 2>&1 | Out-String)
    Assert-Equal '[ ] を含むフォルダから起動する: 終了コード' $LASTEXITCODE '0'
    Assert-Equal '[ ] を含むフォルダから起動する' (Test-Contains $brLog ('対象: {0}\[br]\a.txt' -f $workRoot)) 'True'

    # 調べるフォルダの中で実行すると、案内を出す
    Write-Text "$data\self\x.txt" 'x'
    $selfBat = "$workRoot\self.cmd"
    Write-Text $selfBat ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\self""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -Path . -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $data, $Exe, $tool)
    $selfLog = ConvertTo-FlatText (& cmd.exe /c $selfBat 2>&1 | Out-String)
    Assert-Equal '調べるフォルダの中で実行した案内' (Test-Contains $selfLog 'このツールを、調べるフォルダの中で実行しています') 'True'
    Assert-Equal '外で実行すれば案内しない' (Test-Contains $tree.Log 'このツールを、調べるフォルダの中で実行しています') 'False'

    #endregion
} catch {
    # 途中で止まったときも失敗に数える。コマンドレットのエラーは、catch がないと try の後ろの END まで進み、合格の数が減るだけで気づけない
    '  FAIL  テストが途中で止まりました: {0}' -f ($_.Exception.Message -replace '\s*[\r\n]+\s*', ' ')
    $script:failCount++
} finally {
    foreach ($stream in $held) { try { $stream.Dispose() } catch { } }
    if ($script:substUsed) { [void](& subst.exe ($script:substUsed + ':') /d 2>&1) }
    foreach ($process in $processes) { try { if (-not $process.HasExited) { $process.Kill(); $process.WaitForExit(5000) | Out-Null } } catch { } }
    foreach ($aclPath in @($aclPaths)) { try { [void](& icacls.exe $aclPath /remove:d $me 2>&1) } catch { } }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            # リンクは、リンクだけを先に消す（Directory.Delete の再帰は、中のジャンクションを消したうえで例外を出す）
            foreach ($linkPath in $linkPaths) {
                $attributes = $null
                try { $attributes = [IO.File]::GetAttributes($linkPath) } catch { continue }
                if (($attributes -band [IO.FileAttributes]::Directory) -ne 0) { [IO.Directory]::Delete($linkPath) } else { [IO.File]::Delete($linkPath) }
            }
            foreach ($name in 'readonly.txt', 'offline.txt') {
                if (Test-Path -LiteralPath "$data\$name") { [IO.File]::SetAttributes("$data\$name", [IO.FileAttributes]::Normal) }
            }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

'  END'
exit $script:failCount
