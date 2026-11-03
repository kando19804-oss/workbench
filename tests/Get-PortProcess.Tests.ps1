<#
.SYNOPSIS
    Get-PortProcess の回帰テストです。

.DESCRIPTION
    このテストの中で、待ち受けるだけのプロセス（powershell.exe）を別に起動し、ループバック（127.0.0.1 と ::1）で
    TCP と UDP のポートを開きます。そのプロセスとの間に、接続中の通信と、切断の後始末をしている通信（TimeWait）も作り、
    Get-PortProcess を実行して期待値と照合します。外部のサーバーには接続しません。
    サービスと PID 4（System）は、この PC で待ち受けているものがあるときだけ確かめます。

    一時フォルダに、設定ファイル（日本語のパス、260 文字を超えるパス、文字コードや書き方の違うもの）を作ります。
    「書き込めない保存先」を再現するため、テストが作ったフォルダにだけアクセス拒否を設定します（終了時に戻します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境と待ち受けのプロセスは、終了時に片付けます。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Get-PortProcess を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Get-PortProcess\Get-PortProcess.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Get-PortProcess_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool     = Join-Path $workRoot 'tool\Get-PortProcess.ps1'
$utf8Bom  = New-Object Text.UTF8Encoding $true
$csvHead  = '"プロトコル","ローカルアドレス","ローカルポート","リモートアドレス","リモートポート","状態","PID","プロセス","サービス","実行ユーザー","起動日時","実行ファイル","コマンドライン"'
$marker   = 'wbtest-listener-cmdline'
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

# コンソールで実行すると、長い行は画面の幅で折り返され、出力に改行が入る。
# 画面の幅によって結果が変わらないように、文字を探す前に改行を取り除く
function ConvertTo-FlatText([string]$Text) { $Text -replace '\r?\n', '' }

# Log には、画面への出力から改行を取り除いたものを入れる
function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out  = Join-Path $workRoot "out\$Label"
    $sw   = [Diagnostics.Stopwatch]::StartNew()
    $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = $log; Dir = (Get-RunDir $out); Seconds = $sw.Elapsed.TotalSeconds }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行
function Invoke-Plain([string[]]$ToolArgs) {
    $sw  = [Diagnostics.Stopwatch]::StartNew()
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log; Seconds = $sw.Elapsed.TotalSeconds }
}

# 設定ファイルの OutputFolder や相対パスの基準を確かめるため、-OutputFolder を足さず、指定したフォルダから起動する
function Invoke-From([string]$WorkDir, [string[]]$ToolArgs) {
    Push-Location -LiteralPath $WorkDir
    try {
        $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
        $code = $LASTEXITCODE
    } finally {
        Pop-Location
    }
    [pscustomobject]@{ Code = $code; Log = $log }
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $log = ConvertTo-FlatText ((@(& cmd.exe /d /c $batch 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'PortProcess_*' -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($found.Count -gt 0) { return $found[-1].FullName }
    $null
}

# 結果のファイルがないとき（ツールが保存に失敗したときなど）も止まらないように、空として読む
function Read-Result([string]$Dir) {
    $rows    = @()
    $summary = ''
    if ($Dir -and (Test-Path -LiteralPath (Join-Path $Dir 'ports.csv'))) { $rows = @(Import-Csv -LiteralPath (Join-Path $Dir 'ports.csv') -Encoding UTF8) }
    if ($Dir -and (Test-Path -LiteralPath (Join-Path $Dir 'summary.txt'))) { $summary = [IO.File]::ReadAllText((Join-Path $Dir 'summary.txt')) }
    [pscustomobject]@{ Rows = $rows; Summary = $summary; Lines = @($summary -split '\r?\n') }
}

# ports.csv の行を「プロトコル ローカルアドレス:ポート 状態 PID」の形にする（並び順をまとめて確かめるため）
function Get-RowKeys($Result) {
    (@($Result.Rows | ForEach-Object { '{0} {1}:{2} {3} {4}' -f $_.'プロトコル', $_.'ローカルアドレス', $_.'ローカルポート', $_.'状態', $_.'PID' })) -join ' | '
}

# 条件に合う ports.csv の行。ないとき（ツールが保存に失敗したときなど）も止まらないように、空の行を返す
function Find-Row($Result, [scriptblock]$Where) {
    $found = @($Result.Rows | Where-Object $Where) | Select-Object -First 1
    if ($found) { return $found }
    $empty = [ordered]@{}
    foreach ($name in ($csvHead -replace '"', '') -split ',') { $empty[$name] = '(なし)' }
    [pscustomobject]$empty
}

function Join-Sorted($Items) {
    $array = [string[]]@($Items | Where-Object { $null -ne $_ })
    [Array]::Sort($array, [StringComparer]::Ordinal)
    $array -join ','
}

function Get-ResultFileNames([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { $_.Name })
}

function Get-BomFlags([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir -File | ForEach-Object {
        $bytes = [IO.File]::ReadAllBytes($_.FullName)
        '{0}={1}' -f $_.Name, ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    })
}

function Get-FirstLine([string]$Dir, [string]$FileName) {
    if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir $FileName))) { return '(なし)' }
    @([IO.File]::ReadAllLines((Join-Path $Dir $FileName)))[0]
}

# 画面の「条件」の行（例: ポート 443, 8080   TCP は待ち受けのみ）
function Get-ConditionText([string]$Log) { [regex]::Match($Log, '条件: (.+?(待ち受けのみ|接続中など))').Groups[1].Value }

function Get-CountText([string]$Log) { [regex]::Match($Log, '該当 [^\n]*?個）|該当なし').Value }

# 260 文字を超えるパスにも書けるように、\\?\ を付けて書く
function New-TextFile([string]$FilePath, [string[]]$Lines, [Text.Encoding]$Encoding = $utf8Bom) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllText('\\?\' + $FilePath, ($Lines -join "`r`n") + "`r`n", $Encoding)
    $FilePath
}

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $used -notcontains $_ -and -not [IO.Directory]::Exists("${_}:\") } | Select-Object -First 1
}

# 自分自身にアクセス拒否を設定する。終了時に戻すため、設定したパスを覚えておく
function Add-Deny([string]$Path, [string]$Rights) {
    $null = & icacls.exe $Path /deny "*${userSid}:$Rights" 2>&1
    if ($LASTEXITCODE -ne 0) { return $false }
    $denied.Add($Path)
    $true
}

# 画面の表の 1 行（全角は 2 文字分の幅として、列をそろえる）。$Cells は @(値, 幅, 右寄せ) の並び
function Format-Line([object[]]$Cells) {
    $parts = foreach ($cell in $Cells) {
        $text  = [string]$cell[0]
        $width = 0
        foreach ($ch in $text.ToCharArray()) { $width += $(if ([int]$ch -lt 0x100) { 1 } else { 2 }) }
        $pad = ' ' * [math]::Max(0, [int]$cell[1] - $width)
        if ($cell[2]) { $pad + $text } else { $text + $pad }
    }
    (' ' + ($parts -join '  ')).TrimEnd()
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$userSid  = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$userName = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$isAdmin  = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$denied   = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したパス（終了時に戻す）
$listener = $null
$clients  = New-Object 'System.Collections.Generic.List[object]'   # このテストから張った接続（終了時に閉じる）

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    # 待ち受けるだけのプロセス。同じ番号 P で、TCP の IPv4・IPv6 と UDP を待ち受ける。
    # Q は P より桁の少ない番号（番号を文字として並べると順番が逆になる）。CLOSE は受け付けたらすぐ閉じ（TimeWait を作る）、
    # HOLD は受け付けた接続を持ち続ける（接続中を作る）
    $portFile = Join-Path $workRoot 'listener-ports.txt'
    $stopFile = Join-Path $workRoot 'listener-stop.txt'
    $listenerScript = Join-Path $workRoot 'listener.ps1'
    [IO.File]::WriteAllLines($listenerScript, [string[]]@(
        'param($PortFile, $StopFile, $Tag)'
        '$t4 = $null; $t6 = $null; $u4 = $null'
        'for ($i = 0; $i -lt 30 -and -not $t4; $i++) {'
        '    $a = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, 0); $a.Start(); $p = $a.LocalEndpoint.Port; $b = $null'
        '    if ($p -lt 10000) { $a.Stop(); continue }'
        '    try { $b = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::IPv6Loopback, $p); $b.Start(); $u4 = New-Object Net.Sockets.UdpClient (New-Object Net.IPEndPoint ([Net.IPAddress]::Loopback, $p)); $t4 = $a; $t6 = $b }'
        '    catch { if ($b) { $b.Stop() }; $a.Stop() }'
        '}'
        '$q = $null'
        'foreach ($n in (Get-Random -InputObject (9000..9899) -Count 100)) { try { $q = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, $n); $q.Start(); break } catch { $q = $null } }'
        '$close = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, 0); $close.Start()'
        '$hold  = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, 0); $hold.Start()'
        '$held  = New-Object System.Collections.ArrayList'
        '$lines = @(("P={0}" -f $(if ($t4) { $t4.LocalEndpoint.Port } else { 0 })), ("Q={0}" -f $(if ($q) { $q.LocalEndpoint.Port } else { 0 })), ("CLOSE={0}" -f $close.LocalEndpoint.Port), ("HOLD={0}" -f $hold.LocalEndpoint.Port))'
        '[IO.File]::WriteAllLines($PortFile + ".tmp", [string[]]$lines); [IO.File]::Move($PortFile + ".tmp", $PortFile)'
        'while (-not [IO.File]::Exists($StopFile)) {'
        '    while ($close.Pending()) { $close.AcceptTcpClient().Close() }'
        '    while ($hold.Pending()) { [void]$held.Add($hold.AcceptTcpClient()) }'
        '    Start-Sleep -Milliseconds 50'
        '}'
    ), $utf8Bom)
    $listener = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$listenerScript`"",
        '-PortFile', "`"$portFile`"", '-StopFile', "`"$stopFile`"", '-Tag', $marker) -WindowStyle Hidden -PassThru
    $deadline = (Get-Date).AddSeconds(30)
    while (-not [IO.File]::Exists($portFile) -and (Get-Date) -lt $deadline -and -not $listener.HasExited) { Start-Sleep -Milliseconds 100 }
    $ports = @{}
    if ([IO.File]::Exists($portFile)) { foreach ($line in [IO.File]::ReadAllLines($portFile)) { $kv = $line -split '=', 2; $ports[$kv[0]] = [int]$kv[1] } }
    $P = $(if ($ports.ContainsKey('P')) { $ports['P'] } else { 0 })
    $Q = $(if ($ports.ContainsKey('Q')) { $ports['Q'] } else { 0 })
    $closePort = $(if ($ports.ContainsKey('CLOSE')) { $ports['CLOSE'] } else { 0 })
    $holdPort  = $(if ($ports.ContainsKey('HOLD')) { $ports['HOLD'] } else { 0 })
    $lpid = $listener.Id
    Assert-Equal '待ち受けのプロセスの準備（P・Q・CLOSE・HOLD）' ('{0}/{1}/{2}/{3}' -f ($P -gt 0), ($Q -gt 0), ($closePort -gt 0), ($holdPort -gt 0)) 'True/True/True/True'
    $lproc     = Get-Process -Id $lpid
    $lpath     = $lproc.Path
    $lstart    = $lproc.StartTime
    $hostName  = (Get-Process -Id $PID).ProcessName + '.exe'

    # 接続中: このテストのプロセスから HOLD につなぎ、相手が受け付けたまま持ち続ける
    $holdClient = New-Object Net.Sockets.TcpClient
    $holdClient.Connect('127.0.0.1', $holdPort)
    $clients.Add($holdClient)
    $holdLocal = $holdClient.Client.LocalEndPoint.Port
    # TimeWait: 相手（待ち受けのプロセス）が先に閉じるので、相手の側に TimeWait が残る。PID は 0 になる
    $closeClient = New-Object Net.Sockets.TcpClient
    $closeClient.Connect('127.0.0.1', $closePort)
    $closeClient.GetStream().ReadTimeout = 5000
    try { [void]$closeClient.GetStream().Read((New-Object byte[] 1), 0, 1) } catch { }
    $closeClient.Close()
    Start-Sleep -Milliseconds 300
    $timeWait = @(Get-NetTCPConnection -LocalPort $closePort -State TimeWait -ErrorAction SilentlyContinue).Count -gt 0

    Write-Section 'ポートを指定（-Port）'
    $run = Invoke-Tool 'basic' @('-Port', "$P")
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '保存するファイル' (Get-ResultFileNames $run.Dir) 'ports.csv,summary.txt'
    Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'ports.csv=True,summary.txt=True'
    Assert-Equal 'ports.csv の見出し' (Get-FirstLine $run.Dir 'ports.csv') $csvHead
    Assert-Equal '並び順（TCP の IPv4・IPv6、UDP の順）' (Get-RowKeys $x) "TCP 127.0.0.1:$P 待ち受け $lpid | TCP ::1:$P 待ち受け $lpid | UDP 127.0.0.1:$P - $lpid"
    $row = Find-Row $x { $_.'プロトコル' -eq 'TCP' -and $_.'ローカルアドレス' -eq '127.0.0.1' }
    Assert-Equal 'リモートは空・プロセス・サービスは空' ('{0}|{1}|{2}|{3}' -f $row.'リモートアドレス', $row.'リモートポート', $row.'プロセス', $row.'サービス') '||powershell.exe|'
    Assert-Equal '実行ユーザー' $row.'実行ユーザー' $userName
    Assert-Equal '実行ファイル（大文字と小文字は問わない）' ($row.'実行ファイル' -eq $lpath) 'True'
    Assert-Equal '起動日時（yyyy/MM/dd HH:mm:ss）' $row.'起動日時' $lstart.ToString('yyyy/MM/dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    Assert-Equal 'コマンドラインは CSV にだけ書く' ('{0}/{1}/{2}' -f $row.'コマンドライン'.Contains($marker), $x.Summary.Contains($marker), $run.Log.Contains($marker)) 'True/False/False'
    Assert-Equal 'summary の 1 行目' $x.Lines[0] 'ポート使用プロセス特定 (Get-PortProcess 1.0.0)'
    Assert-Equal 'summary の 2 行目' ([regex]::IsMatch($x.Lines[1], ('^対象: {0}   \d{{4}}-\d{{2}}-\d{{2}} \d{{2}}:\d{{2}}:\d{{2}}   管理者: {1}$' -f [regex]::Escape([Environment]::MachineName), $(if ($isAdmin) { 'はい' } else { 'いいえ' })))) 'True'
    Assert-Equal '条件の行' (Get-ConditionText $run.Log) "ポート $P   TCP は待ち受けのみ"
    Assert-Equal '該当の行' (Get-CountText $run.Log) '該当 3 件（プロセス 1 個）'
    Assert-Equal '自分のプロセスだけなら、管理者の注意は出ない' $run.Log.Contains('※ 管理者として実行すると') 'False'
    Assert-Equal '画面に保存先' $run.Log.Contains(('保存先: {0}' -f $run.Dir)) 'True'

    # 画面の表（全角は 2 文字分の幅として列をそろえる）
    $wLocal = ("127.0.0.1:$P").Length
    $wPid   = [math]::Max(3, "$lpid".Length)
    $expected = @(
        (Format-Line @(@('プロトコル', 10, $false), @('ローカル', $wLocal, $false), @('状態', 8, $false), @('PID', $wPid, $true), @('プロセス', 14, $false), @('サービス', 8, $false))),
        (Format-Line @(@('TCP', 10, $false), @("127.0.0.1:$P", $wLocal, $false), @('待ち受け', 8, $false), @("$lpid", $wPid, $true), @('powershell.exe', 14, $false), @('', 8, $false))),
        (Format-Line @(@('TCP', 10, $false), @("[::1]:$P", $wLocal, $false), @('待ち受け', 8, $false), @("$lpid", $wPid, $true), @('powershell.exe', 14, $false), @('', 8, $false))),
        (Format-Line @(@('UDP', 10, $false), @("127.0.0.1:$P", $wLocal, $false), @('-', 8, $false), @("$lpid", $wPid, $true), @('powershell.exe', 14, $false), @('', 8, $false)))
    )
    $at = [array]::IndexOf($x.Lines, '■ ポート')
    Assert-Equal 'ポートの表' $(if ($at -ge 0) { ($x.Lines[($at + 1)..($at + 4)]) -join ' | ' } else { '(なし)' }) ($expected -join ' | ')
    $wOwner = [math]::Max(12, $userName.Length)
    $expected = @(
        (Format-Line @(@('PID', $wPid, $true), @('プロセス', 14, $false), @('実行ユーザー', $wOwner, $false), @('起動日時', 16, $false), @('実行ファイル', 12, $false))),
        (Format-Line @(@("$lpid", $wPid, $true), @('powershell.exe', 14, $false), @($userName, $wOwner, $false), @($lstart.ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture), 16, $false), @($row.'実行ファイル', 12, $false)))
    )
    $at = [array]::IndexOf($x.Lines, '■ プロセス')
    Assert-Equal 'プロセスの表' $(if ($at -ge 0) { ($x.Lines[($at + 1)..($at + 2)]) -join ' | ' } else { '(なし)' }) ($expected -join ' | ')

    # 番号の順に並べる（文字として並べると 9xxx が 1xxxx より後になる）。範囲とカンマの組み合わせ
    $run = Invoke-Tool 'order' @('-Port', "$P, $Q-$Q")
    Assert-Equal '番号の順（桁の違う番号）' (Get-RowKeys (Read-Result $run.Dir)) "TCP 127.0.0.1:$Q 待ち受け $lpid | TCP 127.0.0.1:$P 待ち受け $lpid | TCP ::1:$P 待ち受け $lpid | UDP 127.0.0.1:$P - $lpid"
    Assert-Equal '条件の行（始まりと終わりが同じ範囲は 1 つの番号）' (Get-ConditionText $run.Log) "ポート $P, $Q   TCP は待ち受けのみ"
    $run = Invoke-Plain @('-Port', ('{0} - {1}' -f ($P - 1), ($P + 1)), '-NoSave')
    Assert-Equal '範囲の中のポート' ('{0}/{1}/{2}' -f $run.Code, (Get-ConditionText $run.Log), $run.Log.Contains("127.0.0.1:$P")) ('0/ポート {0}-{1}   TCP は待ち受けのみ/True' -f ($P - 1), ($P + 1))

    # 待ち受けていないポートは「該当なし」。空いていることの確かめにも使える
    $free = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, 0)
    $free.Start(); $freePort = $free.LocalEndpoint.Port; $free.Stop()
    $run = Invoke-Tool 'none' @('-Port', "$freePort")
    $x = Read-Result $run.Dir
    Assert-Equal '該当なし: 終了コード 0' ('{0}/{1}' -f $run.Code, (Get-CountText $run.Log)) '0/該当なし'
    Assert-Equal '該当なし: 表を出さない' ('{0}/{1}' -f $run.Log.Contains('■ ポート'), $run.Log.Contains('■ プロセス')) 'False/False'
    Assert-Equal '該当なし: ports.csv は見出しだけ' ('{0}/{1}' -f (Get-FirstLine $run.Dir 'ports.csv'), @($x.Rows).Count) "$csvHead/0"

    Write-Section 'ポートを指定しない（待ち受けているすべてのポート）'
    $run = Invoke-Tool 'all' @()
    $x = Read-Result $run.Dir
    $keys = @($x.Rows | ForEach-Object { '{0}|{1}|{2}' -f $_.'プロトコル', $_.'ローカルアドレス', $_.'ローカルポート' })
    Assert-Equal 'このテストの待ち受けが入っている' ('{0}/{1}/{2}' -f ($keys -contains "TCP|127.0.0.1|$P"), ($keys -contains "TCP|::1|$P"), ($keys -contains "UDP|127.0.0.1|$P")) 'True/True/True'
    Assert-Equal 'TCP は待ち受けだけ' @($x.Rows | Where-Object { $_.'プロトコル' -eq 'TCP' -and $_.'状態' -ne '待ち受け' }).Count 0
    Assert-Equal '条件の行' (Get-ConditionText $run.Log) 'すべてのポート   TCP は待ち受けのみ'
    # TCP、UDP の順。同じプロトコルの中は、ローカルポートの番号の順
    $sortedOk = $true
    for ($i = 1; $i -lt $x.Rows.Count; $i++) {
        $a = $x.Rows[$i - 1]; $b = $x.Rows[$i]
        if ($a.'プロトコル' -eq 'UDP' -and $b.'プロトコル' -eq 'TCP') { $sortedOk = $false }
        if ($a.'プロトコル' -eq $b.'プロトコル' -and [int]$a.'ローカルポート' -gt [int]$b.'ローカルポート') { $sortedOk = $false }
    }
    Assert-Equal '並び順（プロトコル、ポートの番号）' $sortedOk 'True'
    $pidCount = @($x.Rows | ForEach-Object { $_.'PID' } | Sort-Object -Unique).Count
    Assert-Equal '該当の行（件数とプロセスの数）' (Get-CountText $run.Log) ('該当 {0:N0} 件（プロセス {1:N0} 個）' -f $x.Rows.Count, $pidCount)

    Write-Section '接続中の通信も調べる（-Established）'
    $run = Invoke-Tool 'established' @('-Port', "$holdPort", '-Established')
    $x = Read-Result $run.Dir
    Assert-Equal '条件の行' (Get-ConditionText $run.Log) "ポート $holdPort   TCP は待ち受けと接続中など"
    $row = Find-Row $x { $_.'ローカルポート' -eq "$holdPort" -and $_.'状態' -eq '待ち受け' }
    Assert-Equal '待ち受け: リモートは空' ('{0}|{1}|{2}' -f $row.'PID', $row.'リモートアドレス', $row.'リモートポート') "$lpid||"
    $row = Find-Row $x { $_.'ローカルポート' -eq "$holdPort" -and $_.'状態' -eq '接続中' }
    Assert-Equal '接続中（受け付けた側）' ('{0}|{1}|{2}' -f $row.'PID', $row.'リモートアドレス', $row.'リモートポート') "$lpid|127.0.0.1|$holdLocal"
    # リモートのポートが一致する行も対象にする（このテストのプロセスから HOLD につないでいる）
    $row = Find-Row $x { $_.'ローカルポート' -eq "$holdLocal" }
    Assert-Equal '接続中（つないだ側）: リモートのポートで見つかる' ('{0}|{1}|{2}|{3}|{4}' -f $row.'状態', $row.'PID', $row.'プロセス', $row.'リモートアドレス', $row.'リモートポート') "接続中|$PID|$hostName|127.0.0.1|$holdPort"
    Assert-Equal '画面にリモートの列' ($run.Log.Contains('リモート') -and $run.Log.Contains("127.0.0.1:$holdPort")) 'True'
    $run = Invoke-Tool 'listen-only' @('-Port', "$holdPort")
    Assert-Equal '-Established なし: 待ち受けだけ' (Get-RowKeys (Read-Result $run.Dir)) "TCP 127.0.0.1:$holdPort 待ち受け $lpid"
    # UDP は -Established に関係なく出す
    $run = Invoke-Tool 'udp-est' @('-Port', "$P", '-Established')
    Assert-Equal 'UDP は -Established でも出す・状態は -' ((Find-Row (Read-Result $run.Dir) { $_.'プロトコル' -eq 'UDP' }).'状態') '-'

    if ($timeWait) {
        $run = Invoke-Tool 'timewait' @('-Port', "$closePort", '-Established')
        $x = Read-Result $run.Dir
        $row = Find-Row $x { $_.'状態' -eq 'TimeWait' -and $_.'ローカルポート' -eq "$closePort" }
        Assert-Equal 'TimeWait: PID 0 と注記' ('{0}|{1}|{2}' -f $row.'PID', $row.'プロセス', $row.'サービス') '0|System Idle Process|（切断の後始末：プロセスなし）'
        Assert-Equal 'TimeWait: 画面にも注記' $run.Log.Contains('（切断の後始末：プロセスなし）') 'True'
        Assert-Equal 'TimeWait: PID 0 では管理者の注意を出さない' $run.Log.Contains('※ 管理者として実行すると') 'False'
    } else {
        Write-Skip 'TimeWait（PID 0）' 'TimeWait の通信を作れなかったため'
    }

    Write-Section 'サービスと PID 4（この PC で待ち受けているものがあるとき）'
    $listen   = @(Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue)
    $services = @(Get-CimInstance -ClassName Win32_Service -ErrorAction SilentlyContinue | Where-Object { $_.ProcessId -gt 4 })
    $svcByPid = @{}
    foreach ($s in $services) { $k = [int]$s.ProcessId; if (-not $svcByPid.ContainsKey($k)) { $svcByPid[$k] = @() }; $svcByPid[$k] += [string]$s.Name }
    $svcConn  = @($listen | Where-Object { $svcByPid.ContainsKey([int]$_.OwningProcess) }) | Select-Object -First 1
    if ($svcConn) {
        $svcPid   = [int]$svcConn.OwningProcess
        $svcNames = @($svcByPid[$svcPid] | Sort-Object)
        $run = Invoke-Tool 'service' @('-Port', "$($svcConn.LocalPort)")
        $x = Read-Result $run.Dir
        $row = Find-Row $x { $_.'PID' -eq "$svcPid" }
        Assert-Equal ('サービス（CSV にはすべて）: ポート {0}' -f $svcConn.LocalPort) $row.'サービス' ($svcNames -join ', ')
        $shown = if ($svcNames.Count -le 3) { $svcNames -join ', ' } else { '{0} ほか {1} 個' -f ($svcNames[0..2] -join ', '), ($svcNames.Count - 3) }
        Assert-Equal '画面のサービス（3 つまで）' $run.Log.Contains($shown) 'True'
        # 管理者でないと、サービスのアカウントのプロセスの情報は取れない。取れないときだけ注意を出す
        $needNote = (-not $isAdmin) -and @($x.Rows | Where-Object { [int]$_.'PID' -gt 4 -and (-not $_.'実行ファイル' -or -not $_.'実行ユーザー') }).Count -gt 0
        Assert-Equal '管理者の注意（取れなかった情報があるときだけ）' $run.Log.Contains('※ 管理者として実行すると') $needNote
    } else {
        Write-Skip 'サービス' 'サービスで待ち受けているポートがないため'
    }
    $manyConn = @($listen | Where-Object { $svcByPid.ContainsKey([int]$_.OwningProcess) -and @($svcByPid[[int]$_.OwningProcess]).Count -gt 3 }) | Select-Object -First 1
    if ($manyConn) {
        $names = @($svcByPid[[int]$manyConn.OwningProcess] | Sort-Object)
        $run = Invoke-Plain @('-Port', "$($manyConn.LocalPort)", '-NoSave')
        Assert-Equal '4 つ以上のサービス: ほか N 個' $run.Log.Contains(('{0} ほか {1} 個' -f ($names[0..2] -join ', '), ($names.Count - 3))) 'True'
    } else {
        Write-Skip '4 つ以上のサービス' '4 つ以上のサービスが動くプロセスで待ち受けているポートがないため'
    }
    $sysConn = @($listen | Where-Object { [int]$_.OwningProcess -eq 4 }) | Select-Object -First 1
    if ($sysConn) {
        $run = Invoke-Tool 'system' @('-Port', "$($sysConn.LocalPort)")
        $row = Find-Row (Read-Result $run.Dir) { $_.'PID' -eq '4' }
        Assert-Equal ('PID 4: System と注記（ポート {0}）' -f $sysConn.LocalPort) ('{0}|{1}' -f $row.'プロセス', $row.'サービス') 'System|（カーネル：HTTP.sys や SMB など）'
    } else {
        Write-Skip 'PID 4（System）' 'PID 4 で待ち受けているポートがないため'
    }

    Write-Section '-Port の書き方と誤り（調べ始める前に止まる）'
    $portErrors = @(
        @('0', '-Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: 0'),
        @('65536', '-Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: 65536'),
        @('http', '-Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: http'),
        @('9000-8000', '-Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: 9000-8000'),
        @('80-90-100', '-Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: 80-90-100'),
        @('99999999999', '-Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: 99999999999'),
        @("$P,abc", '-Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: abc'),
        @(',', '-Port にポートがありません: ,')
    )
    foreach ($case in $portErrors) {
        $run = Invoke-Plain @('-Port', $case[0], '-NoSave')
        Assert-Equal ('-Port {0}' -f $case[0]) ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains($case[1]), $run.Log.Contains('ポート使用プロセス特定 (')) '1/True/False'
    }
    $run = Invoke-Plain @('-Port', " $P , ", '-NoSave')
    Assert-Equal '前後の空白と末尾のカンマ' ('{0}/{1}' -f $run.Code, (Get-ConditionText $run.Log)) "0/ポート $P   TCP は待ち受けのみ"

    Write-Section '保存'
    $noSaveOut = Join-Path $workRoot 'out\nosave'
    $run = Invoke-Plain @('-Port', "$P", '-NoSave', '-OutputFolder', $noSaveOut)
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}/{2}' -f $run.Code, (Test-Path -LiteralPath $noSaveOut), $run.Log.Contains('保存先')) '0/False/False'
    $jpOut = Join-Path $workRoot 'out\保存 先'
    $run = Invoke-Plain @('-Port', "$P", '-OutputFolder', $jpOut)
    Assert-Equal '日本語とスペースを含む保存先' ('{0}/{1}' -f $run.Code, (Get-ResultFileNames (Get-RunDir $jpOut))) '0/ports.csv,summary.txt'
    # 通常の書き方では末尾のピリオドは無視される。\\?\ 付きで作ると、ピリオドの付いたフォルダができてしまう
    $run = Invoke-Plain @('-Port', "$P", '-OutputFolder', (Join-Path $workRoot 'out\trail.'))
    Assert-Equal '末尾にピリオドが付いた保存先' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultFileNames (Get-RunDir (Join-Path $workRoot 'out\trail'))), [IO.Directory]::Exists('\\?\' + (Join-Path $workRoot 'out\trail.'))) '0/ports.csv,summary.txt/False'
    # 同じ秒に 2 回実行しても、前の結果を上書きしない
    $twiceOut = Join-Path $workRoot 'out\twice'
    $null = Invoke-Plain @('-Port', "$P", '-OutputFolder', $twiceOut)
    $null = Invoke-Plain @('-Port', "$P", '-OutputFolder', $twiceOut)
    Assert-Equal '2 回実行すると 2 つのフォルダ' @(Get-ChildItem -LiteralPath $twiceOut -Directory -Filter 'PortProcess_*').Count 2
    $freeLetter = Get-FreeLetter
    if ($freeLetter) {
        $run = Invoke-Plain @('-Port', "$P", '-OutputFolder', "${freeLetter}:\nope")
        Assert-Equal '書き込めない保存先: 警告して画面表示だけ' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("保存先に書き込めないため、結果は保存しません: ${freeLetter}:\nope"), $run.Log.Contains('■ ポート')) '0/True/True'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }
    # 保存先のフォルダは作れるが、その中にファイルを作れないようにして、保存に失敗させる
    $guardOut = Join-Path $workRoot 'out\save-fail'
    [void][IO.Directory]::CreateDirectory($guardOut)
    if (Add-Deny $guardOut '(OI)(CI)(WD)') {
        $run = Invoke-Plain @('-Port', "$P", '-OutputFolder', $guardOut)
        Assert-Equal '保存に失敗: 警告する' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした'), $run.Log.Contains('保存先:')) '0/True/False'
        Assert-Equal '保存に失敗: 書きかけの保存先を残さない' @(Get-ChildItem -LiteralPath $guardOut -Force).Count 0
    } else {
        Write-Skip '保存に失敗したとき' 'アクセス拒否を設定できなかったため'
    }

    Write-Section 'コマンドプロンプトから起動'
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $cmdOut  = Join-Path $workRoot 'out\cmd'
    $run = Invoke-ViaCmd ('{0} -Port {1},{2} -OutputFolder "{3}\"' -f $cmdBase, $P, $Q, $cmdOut)
    Assert-Equal '最後の引数が "パス\" でも保存できる・カンマ区切りのポート' ('{0}/{1}/{2}' -f $run.Code, (Get-ConditionText $run.Log), (Get-ResultFileNames (Get-RunDir $cmdOut))) "0/ポート $P, $Q   TCP は待ち受けのみ/ports.csv,summary.txt"
    $run = Invoke-ViaCmd ('{0} -Port 0' -f $cmdBase)
    Assert-Equal '引数の誤りの終了コード（%ERRORLEVEL%）' $run.Code 1
    # 後ろの -Port まで保存先の値になる。すべてのポートを調べ始めずに、引用符のヒントを出して止まる
    $run = Invoke-ViaCmd ('{0} -OutputFolder "{1}\" -Port {2}' -f $cmdBase, $cmdOut, $P)
    Assert-Equal '"パス\" の後ろに引数: ヒントを出して止まる' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('-OutputFolder の値に引用符'), $run.Log.Contains('最後に \ を付けないでください'), $run.Log.Contains('ポート使用プロセス特定 (')) '1/True/True/False'
    $run = Invoke-ViaCmd ('{0} -NoSave -Port ''{1},{2}''' -f $cmdBase, $P, $Q)
    Assert-Equal 'cmd で ''…,…'' と囲んだ -Port' ('{0}/{1}' -f $run.Code, (Get-ConditionText $run.Log)) "0/ポート $P, $Q   TCP は待ち受けのみ"
    $run = Invoke-ViaCmd ('{0} -NoSave -Port "{1}-{2}" -Established' -f $cmdBase, ($P - 1), ($P + 1))
    Assert-Equal 'cmd から範囲と -Established' ('{0}/{1}' -f $run.Code, (Get-ConditionText $run.Log)) ('0/ポート {0}-{1}   TCP は待ち受けと接続中など' -f ($P - 1), ($P + 1))

    Write-Section '名前に [ ] を含むフォルダから起動（-ConfigFile と -OutputFolder を相対パスで指定）'
    # 5.1 は、名前に [ ] を含むフォルダで起動すると、名前をワイルドカードとして扱い、別の場所から始まる
    $start = Join-Path $workRoot 'start'
    $startCases = @(
        @('一致するフォルダがない（$PSHOME から始まる）', 'br[1] x', @()),
        @('隣に 1 つ一致するフォルダがある', 'sib\[a]', @('sib\a')),
        @('[ が閉じていない（ドライブのルートから始まる）', 'ub\a[b', @('ub\ab')),
        @('隣に 2 つ一致する（ドライブのルートから始まる）', 'mm\[ab]', @('mm\a', 'mm\b'))
    )
    foreach ($case in $startCases) {
        $null = New-TextFile "$start\$($case[1])\port.ini" @("Port = $P", 'NoSave = true')
        # 間違ったフォルダの設定ファイルを読んだら条件で分かるように、隣のフォルダには別のポートを書く
        foreach ($decoy in $case[2]) { $null = New-TextFile "$start\$decoy\port.ini" @("Port = $Q", 'NoSave = true') }
    }
    foreach ($case in $startCases) {
        $dir = "$start\$($case[1])"
        $run = Invoke-From $dir @('-ConfigFile', 'port.ini')
        Assert-Equal ('[ ] を含むフォルダから -ConfigFile: {0}' -f $case[0]) ('{0}/{1}' -f $run.Code, (Get-ConditionText $run.Log)) "0/ポート $P   TCP は待ち受けのみ"
        $run = Invoke-From $dir @('-Port', "$P", '-OutputFolder', 'out')
        Assert-Equal ('[ ] を含むフォルダから -OutputFolder: {0}' -f $case[0]) ('{0}/{1}' -f $run.Code, (Get-ResultFileNames (Get-RunDir "$dir\out"))) '0/ports.csv,summary.txt'
    }

    Write-Section '260 文字を超えるパス'
    $longDir = "$workRoot\" + ('L' * 150) + '\' + ('M' * 150)
    $run = Invoke-Plain @('-Port', "$P", '-OutputFolder', "$longDir\out")
    $longRun = @(Get-ChildItem -LiteralPath ('\\?\' + "$longDir\out") -Directory -ErrorAction SilentlyContinue)
    $longFiles = if ($longRun.Count -gt 0) { Join-Sorted (Get-ChildItem -LiteralPath $longRun[0].FullName -File | ForEach-Object { $_.Name }) } else { '(なし)' }
    Assert-Equal '-OutputFolder' ('{0}/{1}/{2}' -f $run.Code, $longFiles, $run.Log.Contains('保存先に書き込めない')) '0/ports.csv,summary.txt/False'
    $longIni = New-TextFile "$longDir\port.ini" @("Port = $P", 'OutputFolder = 結果')
    $run = Invoke-Plain @('-ConfigFile', $longIni)
    $longRun = @(Get-ChildItem -LiteralPath ('\\?\' + "$longDir\結果") -Directory -ErrorAction SilentlyContinue)
    Assert-Equal '-ConfigFile（OutputFolder は設定ファイルのフォルダが基準）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("設定ファイル: $longIni"), $longRun.Count) '0/True/1'

    Write-Section '設定ファイル（-ConfigFile）'
    $cfgDir = Join-Path $workRoot 'cfg\設定 フォルダ'
    $cwd    = Join-Path $workRoot 'cwd'
    [void][IO.Directory]::CreateDirectory($cwd)
    $basic = New-TextFile "$cfgDir\basic.ini" @(
        '# 設定ファイルのテスト（日本語のコメント）'
        '; セミコロンのコメント'
        "Port = $P"
        "port = `"$Q`""
        'Established = true'
        'OutputFolder = .\結果'
        'NoSave = false'
        'Port ='
    )
    # 別のフォルダから起動して、相対パスが設定ファイルのフォルダを基準にすることを確かめる
    $run = Invoke-From $cwd @('-ConfigFile', $basic)
    Assert-Equal '終了コード' $run.Code 0
    $runDir = Get-RunDir "$cfgDir\結果"
    Assert-Equal '保存先は設定ファイルのフォルダが基準' (Get-ResultFileNames $runDir) 'ports.csv,summary.txt'
    Assert-Equal '起動したフォルダには保存しない' (Test-Path -LiteralPath "$cwd\結果") 'False'
    $x = Read-Result $runDir
    $condLine = [array]::IndexOf($x.Lines, @($x.Lines | Where-Object { $_.StartsWith('条件: ') })[0])
    Assert-Equal '「条件」の前の行に設定ファイル' $(if ($condLine -gt 0) { $x.Lines[$condLine - 1] } else { '(なし)' }) "設定ファイル: $basic"
    Assert-Equal 'Port（2 行）・Established' (Get-ConditionText $run.Log) "ポート $P, $Q   TCP は待ち受けと接続中など"

    # コマンドラインの値が優先する。-Port は置き換え（両方を合わせない）、-NoSave は NoSave = false より優先
    $run = Invoke-From $cwd @('-ConfigFile', $basic, '-Port', "$Q", '-NoSave')
    Assert-Equal 'コマンドラインが優先' ('{0}/{1}/{2}' -f $run.Code, (Get-ConditionText $run.Log), @(Get-ChildItem -LiteralPath "$cfgDir\結果" -Directory).Count) "0/ポート $Q   TCP は待ち受けと接続中など/1"

    $encodings = [ordered]@{
        'UTF-8（BOM あり）'       = $utf8Bom
        'UTF-8（BOM なし）'       = New-Object Text.UTF8Encoding $false
        'Shift_JIS'               = [Text.Encoding]::GetEncoding(932)
        'UTF-16 LE（メモ帳の Unicode）' = [Text.Encoding]::Unicode
        'UTF-16 BE'               = [Text.Encoding]::BigEndianUnicode
    }
    $i = 0
    foreach ($label in $encodings.Keys) {
        $file = New-TextFile "$cfgDir\enc$i.ini" @("Port = $P", 'OutputFolder = ..\設定 フォルダ\文字コード') $encodings[$label]
        $run  = Invoke-From $cwd @('-ConfigFile', $file)
        Assert-Equal ('文字コード {0}: 日本語のパス' -f $label) ('{0}/{1}' -f $run.Code, $run.Log.Contains("保存先: $cfgDir\文字コード\PortProcess_")) '0/True'
        $i++
    }

    # %TEMP% などの環境変数を展開する
    $env:WORKBENCH_TEST_ROOT = $workRoot
    $envIni = New-TextFile "$cfgDir\env.ini" @("Port = $P", 'OutputFolder = %WORKBENCH_TEST_ROOT%\out\cfg-env')
    $run = Invoke-From $cwd @('-ConfigFile', $envIni)
    Assert-Equal '環境変数' ('{0}/{1}' -f $run.Code, (Get-ResultFileNames (Get-RunDir "$workRoot\out\cfg-env"))) '0/ports.csv,summary.txt'

    Write-Section '設定ファイル: 書き間違い（調べ始める前に止まる）'
    $mistakes = @(
        @('知らない項目', @("Port = $P", 'Establised = true'), @('設定ファイルの 2 行目: 知らない項目です: Establised', '使える項目: Port, Established, OutputFolder, NoSave')),
        @('ConfigFile は書けない', @('ConfigFile = other.ini'), @('設定ファイルの 1 行目: 知らない項目です: ConfigFile')),
        @('同じ項目を 2 回（Port 以外）', @('NoSave = true', "Port = $P", 'nosave = false'), @('設定ファイルの 3 行目: NoSave が 2 回書かれています（1 行目と 3 行目）')),
        @('Port の誤り（何行目か）', @("Port = $P", 'Port = 443, 70000'), @('設定ファイルの 2 行目: Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: 70000')),
        @('Established が true / false 以外', @('Established = yes'), @('設定ファイルの 1 行目: Established は true か false で指定してください: yes')),
        @('NoSave が true / false 以外', @('NoSave = 1'), @('設定ファイルの 1 行目: NoSave は true か false で指定してください: 1')),
        @('= がない行', @('Port 443'), @('「項目 = 値」の形になっていません: Port 443')),
        @('セクションの見出し', @('[Get-PortProcess]', "Port = $P"), @('設定ファイルの 1 行目: 「項目 = 値」の形になっていません: [Get-PortProcess]')),
        @('OutputFolder の引用符が片方だけ', @('OutputFolder = "C:\report'), @('OutputFolder の値に引用符（"）が残っています: "C:\report')),
        @('定義されていない環境変数', @('OutputFolder = %WORKBENCH_NO_SUCH_VAR%\x'), @('定義されていない環境変数です: %WORKBENCH_NO_SUCH_VAR%'))
    )
    for ($i = 0; $i -lt $mistakes.Count; $i++) {
        $file    = New-TextFile "$cfgDir\err$i.ini" $mistakes[$i][1]
        $run     = Invoke-From $cwd @('-ConfigFile', $file)
        $missing = @(@($mistakes[$i][2]) + "ファイル: $file" | Where-Object { -not $run.Log.Contains($_) })
        Assert-Equal ('書き間違い: {0}' -f $mistakes[$i][0]) ('{0}/{1}/{2}' -f $run.Code, ($missing -join ' | '), $run.Log.Contains('ポート使用プロセス特定 (')) '1//False'
    }
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\nothing.ini")
    Assert-Equal '設定ファイルが見つからない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("設定ファイルが見つかりません: $cfgDir\nothing.ini")) '1/True'
    $null = New-TextFile "$cfgDir\memo.ini.txt" @("Port = $P")
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\memo.ini")
    Assert-Equal '設定ファイルの拡張子が .txt' ('{0}/{1}' -f $run.Code, $run.Log.Contains("拡張子が .txt になっていませんか: $cfgDir\memo.ini.txt")) '1/True'
    if ($freeLetter) {
        $badOut = New-TextFile "$cfgDir\badout.ini" @("Port = $P", "OutputFolder = ${freeLetter}:\nope")
        $run = Invoke-From $cwd @('-ConfigFile', $badOut)
        Assert-Equal '設定ファイルの OutputFolder に書き込めない: 行番号' ('{0}/{1}' -f $run.Code, $run.Log.Contains('（設定ファイルの 2 行目）')) '0/True'
    }
    $cmdIni = New-TextFile "$workRoot\cmdcfg\cmd.ini" @("Port = $holdPort", 'Established = true', 'NoSave = true')
    $run = Invoke-ViaCmd ('{0} -ConfigFile "{1}"' -f $cmdBase, $cmdIni)
    Assert-Equal 'コマンドプロンプトから -ConfigFile（スイッチの項目）' ('{0}/{1}/{2}' -f $run.Code, (Get-ConditionText $run.Log), $run.Log.Contains('保存先')) "0/ポート $holdPort   TCP は待ち受けと接続中など/False"

    Write-Section 'レビューとテストで見つかった問題の再発防止'
    # 1. 実行ユーザーを Win32_Process の GetOwner で 1 つずつ取っていた。GetOwner は 1 回ごとにすべてのプロセスを調べ直すため、
    #    待ち受けているプロセスが多いと 20 秒以上かかっていた。実行ファイルが取れない（GetOwner も失敗する）プロセスが多い PC で、
    #    その数だけ GetOwner を呼ぶより短い時間で終わることを確かめる
    $listenPids = @(@($listen | ForEach-Object { [int]$_.OwningProcess }) + @(Get-NetUDPEndpoint -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.OwningProcess }) | Sort-Object -Unique)
    $procs   = @(Get-CimInstance -ClassName Win32_Process -ErrorAction SilentlyContinue | Where-Object { $listenPids -contains [int]$_.ProcessId -and [int]$_.ProcessId -gt 4 })
    $noPath  = @($procs | Where-Object { -not $_.ExecutablePath })
    $sw = [Diagnostics.Stopwatch]::StartNew()
    if ($noPath.Count -gt 0) { $null = Invoke-CimMethod -InputObject $noPath[0] -MethodName GetOwner -ErrorAction SilentlyContinue }
    $wasted = $noPath.Count * $sw.Elapsed.TotalSeconds
    if ($wasted -ge 6) {
        $run = Invoke-Plain @('-NoSave')
        Assert-Equal ('すべてのポート: 失敗する GetOwner の分（{0:N1} 秒）より短い' -f $wasted) ('{0}/{1}' -f $run.Code, ($run.Seconds -lt $wasted)) '0/True'
    } else {
        Write-Skip '実行ユーザーを取る時間' ('この PC では差が出ないため（実行ファイルが取れないプロセス {0} 個）' -f $noPath.Count)
    }
    # 2. 1 の直しで、管理者のときに、Get-Process -IncludeUserName で取れない保護されたプロセス（lsass.exe など）の
    #    実行ユーザーが空になっていた（GetOwner なら取れる）
    $protectedRow = $null
    if ($isAdmin) {
        foreach ($proc in $noPath) {
            $r = Invoke-CimMethod -InputObject $proc -MethodName GetOwner -ErrorAction SilentlyContinue
            if (-not $r -or $r.ReturnValue -ne 0) { continue }
            $port = @(@($listen | Where-Object { [int]$_.OwningProcess -eq [int]$proc.ProcessId } | ForEach-Object { $_.LocalPort }) +
                @(Get-NetUDPEndpoint -OwningProcess $proc.ProcessId -ErrorAction SilentlyContinue | ForEach-Object { $_.LocalPort })) | Select-Object -First 1
            if ($port) { $protectedRow = @($proc.ProcessId, $port, $(if ($r.Domain) { '{0}\{1}' -f $r.Domain, $r.User } else { $r.User })); break }
        }
    }
    if ($protectedRow) {
        $run = Invoke-Tool 'protected' @('-Port', "$($protectedRow[1])")
        Assert-Equal ('管理者: 保護されたプロセスの実行ユーザー（PID {0}）' -f $protectedRow[0]) (Find-Row (Read-Result $run.Dir) { $_.'PID' -eq "$($protectedRow[0])" }).'実行ユーザー' $protectedRow[2]
    } elseif ($isAdmin) {
        Write-Skip '保護されたプロセスの実行ユーザー' '実行ファイルが取れないプロセスで待ち受けているポートがないため'
    } else {
        Write-Skip '保護されたプロセスの実行ユーザー' '管理者でないため（管理者のときだけ取れる）'
    }
    # 3. 260 文字を超える -OutputFolder と -ConfigFile を 5.1 で使えなかった（上の「260 文字を超えるパス」で確かめている）
    # 4. 末尾にピリオドが付いた保存先（上の「保存」で確かめている）

    Write-Section '地域設定'
    # 日付の書式が、OS の地域設定によって変わらないこと
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Port '{1}' -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $P, $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        Assert-Equal 'fi-FI でも対象の行の日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($x.Summary, '対象: .*   \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   管理者')) 'True'
        Assert-Equal 'fi-FI でもプロセスの表の起動日時は yyyy-MM-dd HH:mm' $x.Summary.Contains($lstart.ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)) 'True'
        Assert-Equal 'fi-FI でも ports.csv の起動日時は yyyy/MM/dd HH:mm:ss' ((Find-Row $x { $_.'PID' -eq "$lpid" }).'起動日時' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }
} finally {
    # 接続を閉じ、待ち受けのプロセスを止め、アクセス拒否を戻してから消す
    foreach ($c in $clients) { try { $c.Close() } catch { } }
    if ($listener) {
        try { [IO.File]::WriteAllText($stopFile, 'stop') } catch { }
        if (-not $listener.WaitForExit(5000)) { try { $listener.Kill() } catch { } }
    }
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    Remove-Item Env:\WORKBENCH_TEST_ROOT -ErrorAction SilentlyContinue
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

'  END'
exit $script:failCount
