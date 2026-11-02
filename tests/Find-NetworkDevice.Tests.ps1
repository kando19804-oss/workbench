<#
.SYNOPSIS
    Find-NetworkDevice の回帰テストです。

.DESCRIPTION
    本物のネットワークの範囲には ping を送りません。送るのは次の宛先だけです。
    - ループバック（127.x.x.x。Windows ではすべてが ping に応答し、この PC の外に出ません）
    - 文書用のアドレス 192.0.2.x（数件。どこにもつながらず、応答なしの確認に使います）
    - この PC の既定のゲートウェイ（1 件）

    「この PC のネットワーク」と ARP の表を確かめるときは、アダプターの情報を差し替えたツールのコピーと、
    PATH の先頭に置いた偽の arp.cmd を使います（127.x.x.x をこの PC のネットワークに見立てます）。
    NetBIOS の名前は、127.0.1.x で答える補助スクリプトと、この PC 自身の NetBIOS を相手に確かめます（補助スクリプトは終了時に止めます）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Find-NetworkDevice を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Find-NetworkDevice\Find-NetworkDevice.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Find-NetworkDevice_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Find-NetworkDevice.ps1'
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
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
function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $out  = Join-Path $workRoot "out\$Label"
    $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = $log; Dir = (Get-RunDir $out) }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行
function Invoke-Plain([string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 相対パスの基準を確かめるため、指定したフォルダから起動する
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
    $encoding = if ($codePage -eq 65001) { $utf8NoBom } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $log = ConvertTo-FlatText ((@(& cmd.exe /d /c $batch 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# [powershell]::Create() などで動かすスクリプトを、一時フォルダに書く
function New-HelperScript([string]$Name, [string[]]$Lines) {
    $file = Join-Path $workRoot $Name
    [IO.File]::WriteAllLines($file, $Lines, $utf8Bom)
    $file
}

# 補助スクリプトの「名前=値」の出力を受け取る
function Invoke-Helper([string]$Script, [string[]]$HelperArgs) {
    $result = @{}
    & $Exe -NoProfile -ExecutionPolicy Bypass -File $Script @HelperArgs 2>&1 | ForEach-Object {
        $line = "$_"
        $eq = $line.IndexOf('=')
        if ($eq -gt 0) { $result[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
    }
    $result
}

# 名前の問い合わせを省いたツールの中身。ループバックの名前は引けないので、そのままだと実行のたびに待ち時間いっぱい
# （DNS と NetBIOS で 6 秒）待ち、残った DNS の問い合わせ（1 回で 30 秒以上かかる）のためにプロセスの終わりも遅れて、テストに時間がかかる
function Get-FastToolText {
    [IO.File]::ReadAllText($repoTool).Replace('$NameWaitMs     = 3000', '$NameWaitMs     = 200').Replace('$dnsTargets = @($devices.Values | Where-Object { $_.Kind -ne ''この PC'' })', '$dnsTargets = @()')
}

# 定数や処理を 1 か所だけ書き換えたツールのコピーを作る（本物の環境では起きない条件を確かめるため）
function New-PatchedTool([string]$Name, [string]$From, [string]$To) {
    $text = Get-FastToolText
    if (-not $text.Contains($From)) { throw "書き換える場所が見つかりません: $From" }
    $file = Join-Path $workRoot "patched\$Name\Find-NetworkDevice.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text.Replace($From, $To), $utf8Bom)
    $file
}

# この PC のアダプターを差し替えたツールのコピーを作る。Adapters は @(名前, IP, プレフィックス, MAC, ゲートウェイ（カンマ区切り）) の並び
function New-AdapterTool([string]$Name, [object[]]$Adapters) {
    $code = New-Object 'System.Collections.Generic.List[string]'
    $code.Add('$adapters = New-Object ''System.Collections.Generic.List[object]''')
    foreach ($a in $Adapters) {
        $gateways = (@(([string]$a[4]) -split ',' | Where-Object { $_ } | ForEach-Object { "'$_'" })) -join ','
        $code.Add(('$adapters.Add([pscustomobject]@{{ Name = ''{0}''; Address = (ConvertTo-IpNumber ''{1}''); Ip = ''{1}''; Prefix = {2}; Mac = ''{3}''; Gateways = @({4}) }})' -f $a[0], $a[1], $a[2], $a[3], $gateways))
    }
    New-PatchedTool $Name '$adapters  = Get-LocalAdapter' ($code -join "`r`n")
}

# 偽の arp コマンドが返す表を書く（表示の言語に頼っていないことを確かめるため、見出しは日本語にする）
function Set-FakeArp([string[]]$Rows) {
    $lines = @('', 'インターフェイス: 127.0.1.5 --- 0x5', '  インターネット アドレス 物理アドレス     種類') + $Rows
    [IO.File]::WriteAllLines((Join-Path $fakeBin 'arp.txt'), [string[]]$lines, [Text.Encoding]::GetEncoding(932))
}

function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'NetworkDevice_*' -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($found.Count -gt 0) { return $found[-1].FullName }
    $null
}

function Read-Devices([string]$Dir) {
    if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir 'devices.csv'))) { return @() }
    @(Import-Csv -LiteralPath (Join-Path $Dir 'devices.csv') -Encoding UTF8)
}

# 結果の IP アドレスを、ファイルの順に並べる
function Get-Ips($Rows) { (@($Rows) | ForEach-Object { $_.'IP アドレス' }) -join ',' }

function Get-CodeAndIps($Run) { '{0}/{1}' -f $Run.Code, (Get-Ips (Read-Devices $Run.Dir)) }

# 指定した列の値を、IP アドレスの行ごとに並べる
function Get-Cells($Rows, [string]$Column) { (@($Rows) | ForEach-Object { $_.$Column }) -join ',' }

function Get-Row($Rows, [string]$Ip) { @(@($Rows) | Where-Object { $_.'IP アドレス' -eq $Ip }) }

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $used -notcontains $_ -and -not [IO.Directory]::Exists("${_}:\") } | Select-Object -First 1
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
[void][IO.Directory]::CreateDirectory((Split-Path -Parent $tool))
[IO.File]::WriteAllText($tool, (Get-FastToolText), $utf8Bom)
# 名前を確かめるときは、待ち時間を変えていないものを使う。Windows はループバックの UDP 137 番を通さないので、宛先のポートを変える
$toolNames = Join-Path $workRoot 'toolnames\Find-NetworkDevice.ps1'
[void][IO.Directory]::CreateDirectory((Split-Path -Parent $toolNames))
[IO.File]::WriteAllText($toolNames, [IO.File]::ReadAllText($repoTool).Replace('$packet.Length, $ipText, 137)', '$packet.Length, $ipText, 13700)'), $utf8Bom)
$savedPath = $env:PATH
$fakeBin   = Join-Path $workRoot 'fakebin'
[void][IO.Directory]::CreateDirectory($fakeBin)
[IO.File]::WriteAllLines((Join-Path $fakeBin 'arp.cmd'), [string[]]@('@echo off', 'type "%~dp0arp.txt"'), [Text.Encoding]::ASCII)
$responder = $null
$responderStop = Join-Path $workRoot 'responder.stop'

# この PC の、既定のゲートウェイのあるアダプター（テストの期待値を作るため）
$adapter = $null
foreach ($nic in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
    if ($nic.OperationalStatus -ne 'Up' -or $nic.NetworkInterfaceType -eq 'Loopback' -or $nic.NetworkInterfaceType -eq 'Tunnel') { continue }
    $props = $nic.GetIPProperties()
    $gateway = @($props.GatewayAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -ne '0.0.0.0' } | ForEach-Object { $_.Address.ToString() })
    $unicast = @($props.UnicastAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' })
    if ($gateway.Count -gt 0 -and $unicast.Count -gt 0) {
        $adapter = [pscustomobject]@{
            Name = $nic.Name; Ip = $unicast[0].Address.ToString(); Mask = $unicast[0].IPv4Mask.ToString(); Gateway = $gateway[0]
            Mac = (($nic.GetPhysicalAddress().GetAddressBytes() | ForEach-Object { $_.ToString('X2') }) -join '-')
        }
        break
    }
}

# 192.0.2.x が応答しないことを確かめる（応答するネットワークでは、応答なしのテストを飛ばす）
$silentReady = $true
$probe = New-Object Net.NetworkInformation.Ping
try { foreach ($ip in '192.0.2.10', '192.0.2.11') { if ($probe.Send($ip, 500).Status -eq 'Success') { $silentReady = $false } } } catch { $silentReady = $false } finally { $probe.Dispose() }

try {
    Write-Section 'ループバック（ネットワークの外の範囲として、ping だけで探す）'
    $run = Invoke-Tool 'loop' @('-Range', '127.0.0.9-127.0.0.11')
    $rows = Read-Devices $run.Dir
    Assert-Equal '終了コード' $run.Code 0
    # IP アドレスの数値の順（文字の順だと 127.0.0.10 が 127.0.0.9 より前になる）
    Assert-Equal 'IP アドレスの小さい順' (Get-Ips $rows) '127.0.0.9,127.0.0.10,127.0.0.11'
    Assert-Equal 'ping に応答' (Get-Cells $rows 'ping') '応答,応答,応答'
    Assert-Equal 'ネットワークの外なので MAC アドレスは空欄' (Get-Cells $rows 'MAC アドレス') ',,'
    Assert-Equal 'ネットワークの列' (Get-Row $rows '127.0.0.9').'ネットワーク' '127.0.0.9-127.0.0.11'
    Assert-Equal '応答時間は数値' (@(@($rows) | Where-Object { $_.'応答時間(ms)' -notmatch '^\d+$' }).Count) 0
    Assert-Equal '種類は空欄' (Get-Cells $rows '種類') ',,'
    Assert-Equal 'ネットワークの外の注意' $run.Log.Contains('※ 127.0.0.9-127.0.0.11 は、この PC のネットワークの外です。MAC アドレスは分かりません') 'True'
    Assert-Equal '件数の行' ([regex]::Match($run.Log, '機器 \S+ 台（[^）]*）').Value) '機器 3 台（ping に応答 3・応答なし 0）'
    Assert-Equal '対象の行' $run.Log.Contains('対象: 127.0.0.9-127.0.0.11（127.0.0.9 ～ 127.0.0.11・3 件）') 'True'
    Assert-Equal '対象の行は 1 回だけ' ([regex]::Matches($run.Log, '対象: ').Count) 1
    Assert-Equal '先頭の行' $run.Log.StartsWith('ネットワークの機器一覧 (Find-NetworkDevice 1.0.0)') 'True'
    Assert-Equal '保存先の行' $run.Log.Contains(('保存先: {0}' -f $run.Dir)) 'True'

    Write-Section 'CSV とテキストの形'
    Assert-Equal '保存先のフォルダの名前' ((Split-Path -Leaf $run.Dir) -replace '\d{8}_\d{6}$', 'DATE') ('NetworkDevice_{0}_DATE' -f $env:COMPUTERNAME)
    Assert-Equal '保存するファイル' ((@(Get-ChildItem -LiteralPath $run.Dir -File | ForEach-Object { $_.Name } | Sort-Object)) -join ',') 'devices.csv,summary.txt'
    Assert-Equal 'CSV の見出し' ([IO.File]::ReadAllLines((Join-Path $run.Dir 'devices.csv'))[0]) '"IP アドレス","名前","名前の調べ方","MAC アドレス","ping","応答時間(ms)","種類","ネットワーク"'
    foreach ($name in 'devices.csv', 'summary.txt') {
        $head = New-Object byte[] 3
        $stream = [IO.File]::OpenRead((Join-Path $run.Dir $name))
        try { [void]$stream.Read($head, 0, 3) } finally { $stream.Dispose() }
        Assert-Equal ('{0} は BOM 付き UTF-8' -f $name) ($head -join ',') '239,187,191'
    }
    $summary = @([IO.File]::ReadAllLines((Join-Path $run.Dir 'summary.txt')))
    Assert-Equal 'summary.txt の先頭' $summary[0] 'ネットワークの機器一覧 (Find-NetworkDevice 1.0.0)'
    Assert-Equal 'summary.txt に対象の行' (@($summary | Where-Object { $_ -like '対象: *' }).Count) 1
    Assert-Equal 'summary.txt に表の行' (@($summary | Where-Object { $_ -like ' 127.0.0.*' }).Count) 3
    Assert-Equal 'summary.txt の最後は保存先' $summary[-1] ('保存先: {0}' -f $run.Dir)

    Write-Section '範囲の書き方'
    $rangeCases = @(
        @('CIDR（ネットワークとブロードキャストを除く）', '127.0.0.8/30', '127.0.0.9,127.0.0.10'),
        @('/31 はすべて使う', '127.0.0.8/31', '127.0.0.8,127.0.0.9'),
        @('/32 は 1 つ', '127.0.0.8/32', '127.0.0.8'),
        @('CIDR の途中のアドレスから（ネットワークの先頭にそろえる）', '127.0.0.10/30', '127.0.0.9,127.0.0.10'),
        @('1 つだけの IP', '127.0.0.8', '127.0.0.8'),
        @('始まりと終わりが同じ', '127.0.0.8-127.0.0.8', '127.0.0.8'),
        @('- の前後の空白', '127.0.0.8 - 127.0.0.9', '127.0.0.8,127.0.0.9'),
        @('カンマ区切りと小さい順', '127.0.0.20,127.0.0.5-127.0.0.6', '127.0.0.5,127.0.0.6,127.0.0.20'),
        @('カンマの前後の空白', ' 127.0.0.8 , 127.0.0.9 ', '127.0.0.8,127.0.0.9'),
        @('重なる範囲は 1 行', '127.0.0.8-127.0.0.10,127.0.0.9-127.0.0.11', '127.0.0.8,127.0.0.9,127.0.0.10,127.0.0.11'),
        @('オクテットをまたいでも数値の順', '127.1.0.1,127.0.1.0,127.0.0.255', '127.0.0.255,127.0.1.0,127.1.0.1'),
        @('空の要素は無視', '127.0.0.8,,', '127.0.0.8')
    )
    foreach ($case in $rangeCases) {
        $run = Invoke-Tool ('range' + [Array]::IndexOf($rangeCases, $case)) @('-Range', $case[1])
        Assert-Equal ('-Range {0}: {1}' -f $case[1], $case[0]) (Get-CodeAndIps $run) "0/$($case[2])"
    }
    # レビューで見つかった問題: PowerShell から -Range a,b と配列で渡すと、空白でつながってエラーになっていた
    $out = Join-Path $workRoot 'out\array'
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command ("& '{0}' -Range 127.0.0.30,127.0.0.31 -OutputFolder '{1}'" -f $tool, $out) 2>&1
    Assert-Equal 'PowerShell から配列で渡す' (Get-Ips (Read-Devices (Get-RunDir $out))) '127.0.0.30,127.0.0.31'
    # ping を分けて送る境目（同時に 64 件）で、取りこぼさない
    $run = Invoke-Tool 'batch' @('-Range', '127.0.3.1-127.0.3.200')
    $rows = @(Read-Devices $run.Dir)
    Assert-Equal '200 件でも取りこぼさない' ('{0}/{1}/{2}' -f $rows.Count, $rows[0].'IP アドレス', $rows[-1].'IP アドレス') '200/127.0.3.1/127.0.3.200'
    Assert-Equal '200 件すべて応答' (@($rows | Where-Object { $_.'ping' -ne '応答' }).Count) 0

    Write-Section '止まるとき'
    $badSyntax = '-Range の書き方が正しくありません'
    $stopCases = @(
        @('文字', @('-Range', 'abc'), $badSyntax),
        @('オクテットが足りない', @('-Range', '1.2.3'), $badSyntax),
        @('オクテットが多い', @('-Range', '1.2.3.4.5'), $badSyntax),
        @('256 以上', @('-Range', '1.2.3.256'), $badSyntax),
        @('プレフィックス 33', @('-Range', '1.2.3.4/33'), $badSyntax),
        @('プレフィックス 0', @('-Range', '1.2.3.4/0'), $badSyntax),
        @('CIDR の IP の誤り', @('-Range', '1.2.3.999/24'), $badSyntax),
        @('始まりが終わりより後', @('-Range', '127.0.0.9-127.0.0.1'), $badSyntax),
        @('終わりの誤り', @('-Range', '1.2.3.4-abc'), $badSyntax),
        @('終わりが短い', @('-Range', '1.2.3.4-50'), $badSyntax),
        @('空の範囲', @('-Range', ' , '), '-Range に範囲を書いてください'),
        @('誤りの後ろに正しい範囲', @('-Range', 'abc,127.0.0.1'), $badSyntax),
        @('4,097 件', @('-Range', '127.0.0.1-127.0.16.1'), '-Range の範囲が大きすぎます（4,097 件）。合わせて 4,096 件まで'),
        @('合わせて 4,097 件', @('-Range', '127.0.0.1-127.0.16.0,127.0.20.1'), '-Range の範囲が大きすぎます（4,097 件）'),
        @('/19 は大きすぎる', @('-Range', '10.0.0.0/19'), '-Range の範囲が大きすぎます（8,190 件）')
    )
    foreach ($case in $stopCases) {
        $run = Invoke-Plain ($case[1] + @('-NoSave'))
        Assert-Equal $case[0] ('{0}/{1}' -f $run.Code, $run.Log.Contains($case[2])) '1/True'
    }
    foreach ($bad in '99', '10001', 'abc') {
        $run = Invoke-Plain @('-Range', '127.0.0.1', '-Timeout', $bad, '-NoSave')
        Assert-Equal ('-Timeout {0} は止まる' -f $bad) ('{0}/{1}' -f $run.Code, $run.Log.Contains('Timeout')) '1/True'
    }
    foreach ($good in '100', '10000') {
        $run = Invoke-Plain @('-Range', '127.0.0.1', '-Timeout', $good, '-NoSave')
        Assert-Equal ('-Timeout {0} は使える' -f $good) ('{0}/{1}' -f $run.Code, $run.Log.Contains('機器 1 台')) '0/True'
    }
    $run = Invoke-Plain @('positional-value')
    Assert-Equal '名前のない引数は受け付けない' $run.Code 1

    Write-Section 'この PC のネットワークと ARP の表（差し替え）'
    $env:PATH = "$fakeBin;$savedPath"
    $lanTool = New-AdapterTool 'lan' @(, @('テスト LAN', '127.0.1.5', 24, 'AA-BB-CC-00-00-05', '127.0.1.1'))
    Set-FakeArp @(
        '  127.0.1.1             aa-bb-cc-00-00-01     動的        ',
        '  127.0.1.5             de-ad-be-ef-00-05     動的        ',
        '  127.0.1.7             0a-0b-0c-0d-0e-0f     動的        ',
        '  127.0.1.255           ff-ff-ff-ff-ff-ff     静的        ',
        '  127.0.1.9             00-00-00-00-00-00     無効        ',
        '  127.0.1.10            01-00-5e-00-00-16     静的        ',
        '  127.0.0.20            aa-bb-cc-00-00-20     動的        ',
        '  127.0.1.99            aa-bb-cc-00-00-99     動的        '
    )
    $run = Invoke-Tool 'arp' @('-Range', '127.0.1.1-127.0.1.10') $lanTool
    $rows = Read-Devices $run.Dir
    Assert-Equal '範囲の中はすべて載る（すべて ping に応答する）' (Get-Ips $rows) '127.0.1.1,127.0.1.2,127.0.1.3,127.0.1.4,127.0.1.5,127.0.1.6,127.0.1.7,127.0.1.8,127.0.1.9,127.0.1.10'
    Assert-Equal 'ARP の表の MAC アドレスを大文字で書く' (Get-Row $rows '127.0.1.7').'MAC アドレス' '0A-0B-0C-0D-0E-0F'
    Assert-Equal 'ゼロの MAC アドレス（未解決）は使わない' (Get-Row $rows '127.0.1.9').'MAC アドレス' ''
    Assert-Equal 'マルチキャストの MAC アドレスは使わない' (Get-Row $rows '127.0.1.10').'MAC アドレス' ''
    Assert-Equal '範囲の外の ARP の行は載せない' (@(Get-Row $rows '127.0.1.99').Count) 0
    $row = Get-Row $rows '127.0.1.5'
    Assert-Equal 'この PC の行（名前・MAC・種類）' ('{0}/{1}/{2}' -f $row.'名前', $row.'MAC アドレス', $row.'種類') ('{0}/AA-BB-CC-00-00-05/この PC' -f $env:COMPUTERNAME)
    Assert-Equal 'この PC の行（名前の調べ方・ping・応答時間は空欄）' ('{0}/{1}/{2}' -f $row.'名前の調べ方', $row.'ping', $row.'応答時間(ms)') '//'
    $row = Get-Row $rows '127.0.1.1'
    Assert-Equal 'ゲートウェイの行' ('{0}/{1}/{2}' -f $row.'種類', $row.'MAC アドレス', $row.'ping') 'ゲートウェイ/AA-BB-CC-00-00-01/応答'
    Assert-Equal '件数の行にこの PC' $run.Log.Contains('機器 10 台（ping に応答 9・応答なし 0・この PC 1）') 'True'
    Assert-Equal '対象の行にアダプターの名前' $run.Log.Contains('対象: 127.0.1.1-127.0.1.10（127.0.1.1 ～ 127.0.1.10・10 件）   テスト LAN') 'True'
    Assert-Equal 'この PC の IP アドレスの行' ([regex]::IsMatch($run.Log, '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   ' + [regex]::Escape($env:COMPUTERNAME) + '（127\.0\.1\.5）')) 'True'
    Assert-Equal 'ネットワークの中なら外の注意を出さない' $run.Log.Contains('※') 'False'
    Assert-Equal '画面の種類の列' ('{0}/{1}' -f $run.Log.Contains('この PC'), $run.Log.Contains('ゲートウェイ')) 'True/True'
    $run = Invoke-Tool 'arpbroadcast' @('-Range', '127.0.1.255') $lanTool
    Assert-Equal 'ブロードキャストの MAC アドレスは使わない' (Get-Row (Read-Devices $run.Dir) '127.0.1.255').'MAC アドレス' ''
    # この PC のネットワークの外は、ARP の表に載っていても使わない（ルーターの MAC アドレスになっているため）
    $run = Invoke-Tool 'arpoutside' @('-Range', '127.0.0.20') $lanTool
    Assert-Equal '外のアドレスの ARP の行は使わない' ('{0}/{1}' -f (Get-Ips (Read-Devices $run.Dir)), @(Read-Devices $run.Dir)[0].'MAC アドレス') '127.0.0.20/'
    # レビューで見つかった問題: 範囲の一部だけがこの PC のネットワークに入るときに、全体を外として扱っていた
    $run = Invoke-Tool 'partial' @('-Range', '127.0.0.20-127.0.1.1') $lanTool
    $rows = Read-Devices $run.Dir
    Assert-Equal '一部だけ外の注意' $run.Log.Contains('※ 127.0.0.20-127.0.1.1 の一部は、この PC のネットワークの外です。その部分は MAC アドレスが分かりません') 'True'
    Assert-Equal '一部だけ外: 中の部分は MAC アドレスが分かる' (Get-Row $rows '127.0.1.1').'MAC アドレス' 'AA-BB-CC-00-00-01'
    Assert-Equal '一部だけ外: 外の部分は MAC アドレスを使わない' (Get-Row $rows '127.0.0.20').'MAC アドレス' ''
    Assert-Equal '一部だけ外: 対象の行にアダプターの名前' $run.Log.Contains('   テスト LAN') 'True'
    $run = Invoke-Tool 'macform' @('-Range', '127.0.1.20') (New-AdapterTool 'gwself' @(, @('LAN', '127.0.1.20', 24, 'AA-BB-CC-00-00-20', '127.0.1.20')))
    Assert-Equal 'この PC がゲートウェイのときは「この PC」' (Get-Row (Read-Devices $run.Dir) '127.0.1.20').'種類' 'この PC'

    Write-Section 'ping に応答しないが ARP に応答した機器'
    if ($silentReady) {
        $silentTool = New-AdapterTool 'silent' @(, @('文書用', '192.0.2.5', 24, 'AA-BB-CC-00-02-05', '192.0.2.10'))
        Set-FakeArp @('  192.0.2.10            aa-bb-cc-00-02-10     動的        ')
        $run = Invoke-Tool 'silent' @('-Range', '192.0.2.10-192.0.2.11', '-Timeout', '200') $silentTool
        $rows = Read-Devices $run.Dir
        Assert-Equal '応答なしでも MAC アドレスが分かれば載せる' (Get-Ips $rows) '192.0.2.10'
        $row = Get-Row $rows '192.0.2.10'
        Assert-Equal '応答なしの行' ('{0}/{1}/{2}/{3}' -f $row.'ping', $row.'応答時間(ms)', $row.'MAC アドレス', $row.'種類') '応答なし//AA-BB-CC-00-02-10/ゲートウェイ'
        Assert-Equal '応答なしの件数' $run.Log.Contains('機器 1 台（ping に応答 0・応答なし 1）') 'True'
        Assert-Equal '画面の ping の欄' $run.Log.Contains('応答なし  ゲートウェイ') 'True'

        Write-Section '機器が見つからないとき'
        $run = Invoke-Tool 'none' @('-Range', '192.0.2.11', '-Timeout', '200')
        Assert-Equal '終了コードは 0' $run.Code 0
        Assert-Equal '機器は見つかりませんでした' ('{0}/{1}' -f $run.Log.Contains('機器は見つかりませんでした'), $run.Log.Contains('IP アドレス  ')) 'True/False'
        Assert-Equal '件数の行' $run.Log.Contains('機器 0 台（ping に応答 0・応答なし 0）') 'True'
        Assert-Equal '見出しだけの CSV を保存する' ('{0}/{1}' -f @(Read-Devices $run.Dir).Count, [IO.File]::Exists((Join-Path $run.Dir 'devices.csv'))) '0/True'

        Write-Section '-Timeout'
        foreach ($case in @(@('200', 0.2, 2.5), @('2500', 2.5, 5.0))) {
            $run = Invoke-Tool ('timeout' + $case[0]) @('-Range', '192.0.2.11', '-Timeout', $case[0])
            $seconds = [double]::Parse([regex]::Match($run.Log, '所要 ([\d.]+)秒').Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
            Assert-Equal ('-Timeout {0} の所要時間（{1} 秒）' -f $case[0], $seconds) ($seconds -ge $case[1] -and $seconds -lt $case[2]) 'True'
        }
    } else {
        Write-Skip 'ping に応答しない機器' '文書用のアドレス 192.0.2.x が応答するネットワークのため'
    }

    Write-Section '範囲を自動で決める（アダプターを差し替えて、ループバックに ping を送る）'
    Set-FakeArp @('  127.0.1.1             aa-bb-cc-00-00-01     動的        ')
    $autoTool = New-AdapterTool 'auto' @(
        @('ゲートウェイなし', '127.0.2.5', 24, 'AA-BB-CC-00-02-05', ''),
        @('テスト LAN', '127.0.1.5', 30, 'AA-BB-CC-00-00-05', '127.0.1.6')
    )
    $run = Invoke-Tool 'auto' @() $autoTool
    $rows = Read-Devices $run.Dir
    Assert-Equal '既定のゲートウェイのあるネットワークだけを探す' (Get-CodeAndIps $run) '0/127.0.1.5,127.0.1.6'
    Assert-Equal '/30 のラベル' (Get-Row $rows '127.0.1.5').'ネットワーク' '127.0.1.4/30'
    Assert-Equal '自動の対象の行' $run.Log.Contains('対象: 127.0.1.4/30（127.0.1.5 ～ 127.0.1.6・2 件）   テスト LAN') 'True'
    Assert-Equal '自動はネットワークの外の注意を出さない' $run.Log.Contains('※') 'False'
    Assert-Equal 'この PC の IP アドレスは、ゲートウェイのあるアダプターのものを出す' $run.Log.Contains('（127.0.1.5）') 'True'
    $run = Invoke-Tool 'nogw' @('-Range', '127.0.0.8') (New-AdapterTool 'nogw' @(, @('ゲートウェイなし', '127.0.2.5', 24, 'AA-BB-CC-00-02-05', '')))
    Assert-Equal 'ゲートウェイのあるアダプターがないときは、あるものを出す' $run.Log.Contains('（127.0.2.5）') 'True'
    $run = Invoke-Tool 'autotwo' @() (New-AdapterTool 'auto2' @(
        @('有線', '127.0.1.5', 30, 'AA-BB-CC-00-00-05', '127.0.1.6'),
        @('無線', '127.0.2.5', 30, 'AA-BB-CC-00-02-05', '127.0.2.6')
    ))
    Assert-Equal '複数のネットワークを探す' (Get-CodeAndIps $run) '0/127.0.1.5,127.0.1.6,127.0.2.5,127.0.2.6'
    Assert-Equal 'ネットワークごとの対象の行' ('{0}/{1}' -f $run.Log.Contains('   有線'), $run.Log.Contains('   無線')) 'True/True'
    Assert-Equal '両方の IP アドレスを出す' $run.Log.Contains('（127.0.1.5, 127.0.2.5）') 'True'
    $run = Invoke-Tool 'autosame' @() (New-AdapterTool 'auto3' @(
        @('有線', '127.0.1.5', 30, 'AA-BB-CC-00-00-05', '127.0.1.6'),
        @('無線', '127.0.1.6', 30, 'AA-BB-CC-00-00-06', '127.0.1.5')
    ))
    Assert-Equal '同じネットワークは 1 回だけ探す' ('{0}/{1}' -f (Get-Ips (Read-Devices $run.Dir)), [regex]::Matches($run.Log, '対象: ').Count) '127.0.1.5,127.0.1.6/1'
    # 1,024 件の境目（/22 は 1,022 件で探す、/21 は 2,046 件で探さない）
    $run = Invoke-Tool 'auto22' @() (New-AdapterTool 'auto22' @(, @('LAN', '127.0.4.5', 22, 'AA-BB-CC-00-04-05', '127.0.4.1')))
    $rows = @(Read-Devices $run.Dir)
    Assert-Equal '/22 は探す' ('{0}/{1}/{2}/{3}' -f $run.Code, $rows.Count, $rows[0].'IP アドレス', $rows[-1].'IP アドレス') '0/1022/127.0.4.1/127.0.7.254'
    Assert-Equal '件数の 3 桁区切り（対象の行と件数の行）' ('{0}/{1}' -f $run.Log.Contains('・1,022 件）'), $run.Log.Contains('機器 1,022 台（ping に応答 1,021・応答なし 0・この PC 1）')) 'True/True'
    # 大きすぎるネットワーク（VPN の /16 など）は飛ばして、ほかを探す
    $mixTool = New-AdapterTool 'automix' @(
        @('VPN', '10.1.2.3', 21, 'AA-BB-CC-00-10-03', '10.1.0.1'),
        @('VPN 2', '10.1.2.4', 21, 'AA-BB-CC-00-10-04', '10.1.0.1'),
        @('テスト LAN', '127.0.1.5', 30, 'AA-BB-CC-00-00-05', '127.0.1.6')
    )
    $run = Invoke-Tool 'automix' @() $mixTool
    Assert-Equal '大きすぎるネットワークは飛ばして、ほかを探す' (Get-CodeAndIps $run) '0/127.0.1.5,127.0.1.6'
    Assert-Equal '大きすぎるネットワークの知らせ' $run.Log.Contains('※ 10.1.0.0/21（VPN）は大きすぎるため、探していません（2,046 件。1,024 件までです）') 'True'
    Assert-Equal '同じネットワークの知らせは 1 回だけ' ([regex]::Matches($run.Log, '大きすぎるため').Count) 1
    Assert-Equal '狭め方の例' $run.Log.Contains('-Range "10.1.2.0/24"') 'True'
    Assert-Equal 'summary.txt にも知らせを残す' (@([IO.File]::ReadAllLines((Join-Path $run.Dir 'summary.txt')) | Where-Object { $_ -like '※ 10.1.0.0/21*' }).Count) 1
    $run = Invoke-Tool 'autolarge' @() (New-AdapterTool 'autolarge' @(, @('VPN', '10.1.2.3', 16, 'AA-BB-CC-00-10-03', '10.1.0.1')))
    Assert-Equal 'すべて大きすぎるときは止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('10.1.0.0/16（VPN）は大きすぎるため、探しません'), [bool]$run.Dir) '1/True/False'
    $run = Invoke-Tool 'autonone' @() (New-AdapterTool 'autonone' @(, @('ゲートウェイなし', '127.0.2.5', 24, 'AA-BB-CC-00-02-05', '')))
    Assert-Equal 'ゲートウェイがないときは止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('探すネットワークを決められません'), $run.Log.Contains('-Range で範囲を指定してください')) '1/True/True'
    $run = Invoke-Tool 'autorange' @('-Range', '127.0.0.8') $mixTool
    Assert-Equal '-Range があれば大きすぎる知らせは出さない' $run.Log.Contains('大きすぎる') 'False'

    # サブネット マスクのビット数（この PC のアダプターにないマスクも確かめるため、計算だけするコピーを動かす）
    $maskTool = New-PatchedTool 'mask' '$adapters  = Get-LocalAdapter' ("foreach (`$m in '255.0.0.0', '255.255.240.0', '255.255.255.0', '255.255.255.252', '255.255.255.255') { Write-Host (Get-PrefixLength (ConvertTo-IpNumber `$m)) }; exit 0")
    $run = Invoke-Plain @() $maskTool
    Assert-Equal 'サブネット マスクのビット数' $run.Log '820243032'

    Write-Section '表示の言語が英語の ARP の表'
    [IO.File]::WriteAllLines((Join-Path $fakeBin 'arp.txt'), [string[]]@('', 'Interface: 127.0.1.5 --- 0x5', '  Internet Address      Physical Address      Type', '  127.0.1.1             aa-bb-cc-00-00-01     dynamic   '), [Text.Encoding]::ASCII)
    $run = Invoke-Tool 'arpen' @('-Range', '127.0.1.1') $lanTool
    Assert-Equal '英語の見出しでも読む' (Get-Row (Read-Devices $run.Dir) '127.0.1.1').'MAC アドレス' 'AA-BB-CC-00-00-01'
    $env:PATH = $savedPath

    Write-Section 'この PC と既定のゲートウェイ（本物のアダプター。ゲートウェイに ping を 1 回送る）'
    if ($adapter) {
        $run = Invoke-Tool 'self' @('-Range', ('{0},{1}' -f $adapter.Ip, $adapter.Gateway))
        $rows = Read-Devices $run.Dir
        $row = Get-Row $rows $adapter.Ip
        Assert-Equal 'この PC' ('{0}/{1}/{2}/{3}' -f $row.'種類', $row.'MAC アドレス', $row.'ping', $row.'名前') ('この PC/{0}//{1}' -f $adapter.Mac, $env:COMPUTERNAME)
        $row = Get-Row $rows $adapter.Gateway
        Assert-Equal 'ゲートウェイ' $row.'種類' 'ゲートウェイ'
        Assert-Equal 'ゲートウェイの MAC アドレス（本物の ARP の表から読む）' ($row.'MAC アドレス' -match '^[0-9A-F]{2}(-[0-9A-F]{2}){5}$') 'True'
        Assert-Equal '対象の行に本物のアダプターの名前' $run.Log.Contains(('   {0}' -f $adapter.Name)) 'True'
        Assert-Equal '本物のネットワークの中なら外の注意を出さない' $run.Log.Contains('※') 'False'
    } else {
        Write-Skip 'この PC と既定のゲートウェイ' '既定のゲートウェイのあるネットワーク アダプターがないため'
    }

    Write-Section '名前（DNS）'
    # Windows は hosts や自分の名前から 127.0.0.1 の名前を返す。テストでは同じ問い合わせの結果と比べる
    $expected = $null
    try { $expected = [Net.Dns]::GetHostEntry('127.0.0.1').HostName } catch { }
    if ($expected -and $expected -ne '127.0.0.1') {
        $run = Invoke-Tool 'dns' @('-Range', '127.0.0.1') $toolNames
        $row = Get-Row (Read-Devices $run.Dir) '127.0.0.1'
        Assert-Equal 'DNS の逆引きの名前' ('{0}/{1}' -f $row.'名前', $row.'名前の調べ方') "$expected/DNS"
    } else {
        Write-Skip 'DNS の逆引きの名前' '127.0.0.1 の名前を引けないため'
    }

    # Windows でのテストで見つかった問題: 答えの返らない DNS の問い合わせを積み上げて、1,000 台で表示まで 52 秒、プロセスの終わりまで 5 分近くかかっていた
    $run = Invoke-Tool 'dnspile' @('-Range', '127.0.6.1-127.0.6.130') $toolNames
    $seconds = [double]::Parse([regex]::Match($run.Log, '所要 ([\d.]+)秒').Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
    Assert-Equal ('答えの返らない問い合わせを積み上げない（130 台で {0} 秒）' -f $seconds) ($seconds -lt 11) 'True'
    Assert-Equal '問い合わせをやめても、すべての機器を載せる' (@(Read-Devices $run.Dir).Count) 130

    Write-Section '名前（NetBIOS。補助スクリプトが 127.0.1.x で答える）'
    # 本物の NetBIOS は、ループバックでも、この PC 自身の IP アドレス宛てでも、UDP 137 番に届かないので確かめられない（実機確認で見る）
    $responderLines = @'
param($ReadyFile, $StopFile)
$enc = [Text.Encoding]::GetEncoding(932)
function New-Entry([string]$Name, [int]$Suffix, [bool]$Group) {
    $bytes = New-Object byte[] 18
    for ($i = 0; $i -lt 15; $i++) { $bytes[$i] = 32 }
    $raw = $enc.GetBytes($Name)
    [Array]::Copy($raw, $bytes, [math]::Min(15, $raw.Length))
    $bytes[15] = $Suffix
    if ($Group) { $bytes[16] = 0x80 }
    , $bytes
}
$answers = @{
    '127.0.1.20' = @{ Count = 3; Entries = @((New-Entry 'WORKGROUP' 0 $true), (New-Entry 'PC-NB' 0 $false), (New-Entry 'PC-NB' 0x20 $false)) }
    '127.0.1.21' = @{ Count = 1; Entries = @(, (New-Entry 'ONLY-SERVICE' 0x20 $false)) }
    '127.0.1.22' = @{ Count = 1; Entries = @(, (New-Entry 'サーバー１' 0 $false)) }
    '127.0.1.23' = @{ Count = 1; Entries = @(, (New-Entry '=1+1' 0 $false)) }
    '127.0.1.24' = @{ Count = 1; Entries = @(, (New-Entry '12345' 0 $false)) }
    '127.0.1.25' = @{ Short = $true }
    '127.0.1.26' = @{ Count = 3; Entries = @(, (New-Entry 'WRONG-COUNT' 0x20 $false)) }
    '127.0.1.27' = @{ Count = 1; Entries = @(, (New-Entry 'WORKGROUP' 0 $true)) }
    '127.0.1.30' = @{ Count = 1; Entries = @(, (New-Entry '.5' 0 $false)) }
    '127.0.1.31' = @{ Count = 1; Entries = @(, (New-Entry '1e5' 0 $false)) }
    '127.0.1.32' = @{ Count = 1; Entries = @(, (New-Entry '１．５' 0 $false)) }
    '127.0.1.33' = @{ Count = 1; Entries = @(, (New-Entry '2026-01-01' 0 $false)) }
    '127.0.1.34' = @{ Count = 1; Entries = @(, (New-Entry '=a"b' 0 $false)) }
    '127.0.1.35' = @{ Count = 1; Entries = @(, (New-Entry 'a"b' 0 $false)) }
    '127.0.1.36' = @{ Count = 1; Entries = @(, (New-Entry '@host' 0 $false)) }
}
$sockets = @{}
foreach ($ip in $answers.Keys) { $sockets[$ip] = New-Object Net.Sockets.UdpClient (New-Object Net.IPEndPoint ([Net.IPAddress]::Parse($ip)), 13700) }
[IO.File]::WriteAllText($ReadyFile, 'ready')
$end = (Get-Date).AddMinutes(10)
while ((Get-Date) -lt $end -and -not (Test-Path -LiteralPath $StopFile)) {
    foreach ($ip in $answers.Keys) {
        $socket = $sockets[$ip]
        if ($socket.Available -le 0) { continue }
        $remote = New-Object Net.IPEndPoint ([Net.IPAddress]::Any), 0
        $query = $socket.Receive([ref]$remote)
        $reply = New-Object 'System.Collections.Generic.List[byte]'
        $reply.AddRange([byte[]]($query[0], $query[1], 0x84, 0x00, 0, 0, 0, 1, 0, 0, 0, 0))
        if ($answers[$ip].ContainsKey('Short')) {
            $reply.AddRange([byte[]](1, 2, 3))
        } else {
            $reply.Add([byte]0x20)
            $reply.AddRange([Text.Encoding]::ASCII.GetBytes('CK' + ('A' * 30)))
            $reply.AddRange([byte[]](0, 0, 0x21, 0, 1, 0, 0, 0, 0, 0, 0))
            $reply.Add([byte]$answers[$ip].Count)
            foreach ($entry in $answers[$ip].Entries) { $reply.AddRange([byte[]]$entry) }
        }
        $data = $reply.ToArray()
        [void]$socket.Send($data, $data.Length, $remote)
    }
    Start-Sleep -Milliseconds 10
}
foreach ($socket in $sockets.Values) { $socket.Close() }
'@ -split "`r?`n"
    $responderFile = New-HelperScript 'netbios-responder.ps1' $responderLines
    $responderReady = Join-Path $workRoot 'responder.ready'
    $responder = Start-Process -FilePath $Exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$responderFile`"", '-ReadyFile', "`"$responderReady`"", '-StopFile', "`"$responderStop`"") -PassThru -WindowStyle Hidden
    foreach ($i in 1..150) { if (Test-Path -LiteralPath $responderReady) { break }; Start-Sleep -Milliseconds 100 }
    if (Test-Path -LiteralPath $responderReady) {
        $run = Invoke-Tool 'netbios' @('-Range', '127.0.1.20-127.0.1.36') $toolNames
        $rows = Read-Devices $run.Dir
        $row = Get-Row $rows '127.0.1.20'
        Assert-Equal 'グループでなく末尾が 00 の名前を使う' ('{0}/{1}' -f $row.'名前', $row.'名前の調べ方') 'PC-NB/NetBIOS'
        $row = Get-Row $rows '127.0.1.21'
        Assert-Equal '末尾が 00 でない名前は使わない' ('{0}/{1}' -f $row.'名前', $row.'名前の調べ方') '/'
        Assert-Equal '日本語の名前（Shift_JIS）' (Get-Row $rows '127.0.1.22').'名前' 'サーバー１'
        Assert-Equal '数式に見える名前は Excel 対策' (Get-Row $rows '127.0.1.23').'名前' '="=1+1"'
        Assert-Equal '数字だけの名前は Excel 対策' (Get-Row $rows '127.0.1.24').'名前' '="12345"'
        Assert-Equal '短すぎる応答は無視する' ('{0}/{1}' -f (Get-Row $rows '127.0.1.25').'ping', (Get-Row $rows '127.0.1.25').'名前') '応答/'
        Assert-Equal '名前の数より短い応答でも止まらない' ('{0}/{1}' -f (Get-Row $rows '127.0.1.26').'ping', (Get-Row $rows '127.0.1.26').'名前') '応答/'
        Assert-Equal 'グループの名前しかないときは空欄' (Get-Row $rows '127.0.1.27').'名前' ''
        Assert-Equal '答えない機器は空欄' (Get-Row $rows '127.0.1.28').'名前' ''
        Assert-Equal '画面にも名前' ('{0}/{1}' -f $run.Log.Contains('PC-NB'), $run.Log.Contains('=1+1')) 'True/True'
        Assert-Equal '列をそろえる（全角の名前は 2 文字分）' ([regex]::IsMatch($run.Log, '127\.0\.1\.22 +サーバー１ {1,}')) 'True'
        $excelCases = @(
            @('小数', '127.0.1.30', '=".5"'), @('指数', '127.0.1.31', '="1e5"'), @('全角の小数', '127.0.1.32', '="１．５"'), @('日付', '127.0.1.33', '="2026-01-01"'),
            @('" を含む数式', '127.0.1.34', '="=a""b"'), @('" を含むふつうの名前', '127.0.1.35', 'a"b'), @('@', '127.0.1.36', '="@host"')
        )
        foreach ($case in $excelCases) { Assert-Equal ('Excel 対策: {0}' -f $case[0]) (Get-Row $rows $case[1]).'名前' $case[2] }
    } else {
        Write-Skip 'NetBIOS の名前' '補助スクリプトが待ち受けられないため'
    }

    Write-Section '保存'
    $run = Invoke-Plain @('-Range', '127.0.0.1', '-NoSave', '-OutputFolder', "$workRoot\out\nosave")
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}/{2}' -f $run.Code, (Test-Path -LiteralPath "$workRoot\out\nosave"), $run.Log.Contains('127.0.0.1')) '0/False/True'
    Assert-Equal '-NoSave: 保存先の行を出さない' $run.Log.Contains('保存先') 'False'
    $run = Invoke-Plain @('-Range', '127.0.0.1')
    Assert-Equal '既定の保存先はスクリプトのフォルダ' ('{0}/{1}' -f $run.Code, [bool](Get-RunDir (Split-Path -Parent $tool))) '0/True'
    # 同じ秒に実行しても、前の結果を消さない（Ctrl+C や保存の失敗のときはフォルダごと消すため）
    $sameDir = Join-Path $workRoot 'out\same'
    [void][IO.Directory]::CreateDirectory($sameDir)
    $stampNow = Get-Date
    foreach ($s in 0..40) {
        $older = Join-Path $sameDir ('NetworkDevice_{0}_{1}' -f $env:COMPUTERNAME, $stampNow.AddSeconds($s).ToString('yyyyMMdd_HHmmss'))
        [void][IO.Directory]::CreateDirectory($older)
        [IO.File]::WriteAllText("$older\keep.txt", 'keep')
    }
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Range '127.0.0.1' -OutputFolder $sameDir 2>&1
    Assert-Equal '同じ名前のフォルダがあれば _2 を付ける' (@(Get-ChildItem -LiteralPath $sameDir -Directory -Filter 'NetworkDevice_*_2').Count) 1
    Assert-Equal '前の結果を残す' (@(Get-ChildItem -LiteralPath $sameDir -Recurse -Filter 'keep.txt').Count) 41
    # Windows でのテストで見つかった問題: 5.1 で 260 文字を超える保存先に保存できなかった
    $longOut = "$workRoot\" + ('O' * 120) + '\' + ('P' * 120)
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Range '127.0.0.1' -OutputFolder $longOut 2>&1
    $longRuns = @(); try { $longRuns = @([IO.Directory]::GetDirectories('\\?\' + $longOut)) } catch { }
    Assert-Equal '-OutputFolder が長くても保存する' ('{0}/{1}' -f $LASTEXITCODE, $longRuns.Count) '0/1'
    if ($longRuns.Count -eq 1) { Assert-Equal '長い保存先のファイル' (@([IO.Directory]::GetFiles($longRuns[0])).Count) 2 }
    $spaceOut = Join-Path $workRoot 'out\日本語 と スペース [x]'
    $run = Invoke-Plain @('-Range', '127.0.0.1', '-OutputFolder', $spaceOut)
    Assert-Equal '保存先に日本語・スペース・[ ]' ('{0}/{1}' -f $run.Code, (@(Read-Devices (Get-RunDir $spaceOut)).Count)) '0/1'
    $fileOut = Join-Path $workRoot 'out\afile'
    [IO.File]::WriteAllText($fileOut, 'x')
    $run = Invoke-Plain @('-Range', '127.0.0.1', '-OutputFolder', $fileOut)
    Assert-Equal '保存先がファイル: 警告を出して画面に表示' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため、結果は保存しません'), $run.Log.Contains('機器 1 台')) '0/True/True'
    $letter = Get-FreeLetter
    if ($letter) {
        $run = Invoke-Plain @('-Range', '127.0.0.1', '-OutputFolder', "${letter}:\nope")
        Assert-Equal 'ないドライブ: 警告を出して画面に表示' ('{0}/{1}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため')) '0/True'
    } else {
        Write-Skip 'ないドライブ' '空いているドライブ文字がないため'
    }
    $patched = New-PatchedTool 'savefail' "Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'devices.csv'))" "throw 'disk full'; Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'devices.csv'))"
    $run = Invoke-Tool 'savefail' @('-Range', '127.0.0.1') $patched
    Assert-Equal '保存に失敗: 警告を出して書きかけを残さない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした: disk full'), [bool]$run.Dir) '0/True/False'
    Assert-Equal '保存に失敗: 画面には表を出す' ('{0}/{1}' -f $run.Log.Contains('127.0.0.1'), $run.Log.Contains('保存先:')) 'True/False'
    $patched = New-PatchedTool 'crash' '$sorted = @($devices.Values | Sort-Object -Property Ip)' "throw 'boom'"
    $run = Invoke-Tool 'crash' @('-Range', '127.0.0.1') $patched
    Assert-Equal '想定外のエラー: 結果を残さず、中断とは区別する' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('エラーで止まりました。結果は保存していません。'), $run.Log.Contains('中断しました'), [bool]$run.Dir) '1/True/False/False'

    Write-Section '[ ] を含むフォルダからの起動'
    $start = Join-Path $workRoot 'start'
    $startCases = @(@('隣に一致するフォルダ', '[a]', @('a')), @('一致するフォルダなし', 'b[1]', @()), @('パターンとして正しくない', 'c[d', @()), @('2 つ以上一致', '[ef]', @('e', 'f')))
    foreach ($case in $startCases) {
        $dir = "$start\$($case[1])"
        [void][IO.Directory]::CreateDirectory($dir)
        foreach ($decoy in $case[2]) { [void][IO.Directory]::CreateDirectory("$start\$decoy") }
        $run = Invoke-From $dir @('-Range', '127.0.0.1', '-OutputFolder', 'out')
        Assert-Equal ('[ ] を含むフォルダから -OutputFolder: {0}' -f $case[0]) ('{0}/{1}' -f $run.Code, [bool](Get-RunDir "$dir\out")) '0/True'
    }

    Write-Section 'コマンドプロンプトから起動'
    $cmdHead = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $run = Invoke-ViaCmd ('{0} -NoSave -Range "127.0.0.1-127.0.0.2,127.0.0.4"' -f $cmdHead)
    Assert-Equal 'カンマ区切りの -Range' ('{0}/{1}' -f $run.Code, ([regex]::Match($run.Log, '機器 \S+ 台').Value)) '0/機器 3 台'
    $run = Invoke-ViaCmd ("{0} -NoSave -Range '127.0.0.1'" -f $cmdHead)
    Assert-Equal "cmd の ' はそのまま届く" ('{0}/{1}' -f $run.Code, $run.Log.Contains('機器 1 台')) '0/True'
    $run = Invoke-ViaCmd ('{0} -Range 127.0.0.1 -OutputFolder "{1}\" -NoSave' -f $cmdHead, $workRoot)
    Assert-Equal '"パス\" の後ろに引数（止まってヒントを出す）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('最後に \ を付けないでください')) '1/True'
    $cmdOut = Join-Path $workRoot 'out\cmd'
    [void][IO.Directory]::CreateDirectory($cmdOut)
    $run = Invoke-ViaCmd ('{0} -Range 127.0.0.1 -OutputFolder "{1}\"' -f $cmdHead, $cmdOut)
    Assert-Equal '最後の引数が \"（引用符を外して保存する）' ('{0}/{1}' -f $run.Code, [bool](Get-RunDir $cmdOut)) '0/True'
    $run = Invoke-ViaCmd ('{0} -NoSave -Range abc' -f $cmdHead)
    Assert-Equal 'cmd で受け取る終了コード（止まると 1）' $run.Code 1

    Write-Section '画面の色と、ping を送る前の表示'
    $colorScript = New-HelperScript 'colors.ps1' @(
        'param($Tool)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Range'', ''127.0.0.1'').AddParameter(''NoSave'', $true)'
        '[void]$ps.Invoke()'
        'foreach ($record in $ps.Streams.Information) { $m = $record.MessageData; if ($m.Message -like ''※*'') { ''NOTE='' + $m.ForegroundColor } elseif ($m.Message -like ''機器 *'') { ''COUNT='' + $m.ForegroundColor } }'
    )
    $colors = Invoke-Helper $colorScript @('-Tool', $tool)
    Assert-Equal '※ の行は黄色' $colors['NOTE'] 'Yellow'
    Assert-Equal '件数の行は色なし' $colors['COUNT'] ''
    # 名前の問い合わせを待つ時間を延ばしたコピーで、ping の後の待ちの途中に止める
    $slowTool = New-PatchedTool 'slowname' '$NameWaitMs     = 200' '$NameWaitMs     = 60000'
    $stopScript = New-HelperScript 'stop-during-search.ps1' @(
        'param($Tool, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Range'', ''127.0.5.1'').AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'Start-Sleep -Milliseconds 4000'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''HEAD='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''対象: *'' }).Count'
        '''TABLE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''機器 *'' }).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
    )
    $result = Invoke-Helper $stopScript @('-Tool', $slowTool, '-OutRoot', (Join-Path $workRoot 'out\ctrlc'))
    if ($result['RUNNING'] -eq 'True') {
        Assert-Equal '「対象」の行は探し終わる前に出す' ('{0}/{1}' -f $result['HEAD'], $result['TABLE']) '1/0'
        Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $result['BEFORE'], $result['AFTER']) '1/0'
        Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' '4 秒より前に探し終わったため'
    }

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $cultureProbe = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy-MM-dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $cultureProbe)"
    if ($control -and $control -ne '2026-10-29 08:05:00|1.5') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Range '127.0.0.1' -OutputFolder '{1}' | Out-Null" -f $tool.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $cultureDir = Get-RunDir $cultureOut
        $summaryText = if ($cultureDir) { [IO.File]::ReadAllText((Join-Path $cultureDir 'summary.txt')) } else { '' }
        Assert-Equal 'fi-FI でも日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($summaryText, '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   ')) 'True'
        Assert-Equal 'fi-FI でも所要時間の小数点は .' ([regex]::IsMatch($summaryText, '所要 \d+\.\d秒')) 'True'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }
} finally {
    $env:PATH = $savedPath
    if ($responder) {
        try { [IO.File]::WriteAllText($responderStop, 'stop') } catch { }
        try { if (-not $responder.WaitForExit(3000)) { $responder.Kill() } } catch { }
    }
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

exit $script:failCount
