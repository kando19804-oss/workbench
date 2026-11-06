<#
.SYNOPSIS
    Test-NameResolution の回帰テストです。

.DESCRIPTION
    本物の DNS サーバーには問い合わせません。外へ通信しないように、次の 3 つを使い分けます。
    - 偽の環境: hosts の場所と OS の結果を差し替えたツールのコピーを、Resolve-DnsName などと同じ名前の偽の関数を定義した
      ラッパーから動かします。判定の組み合わせは、ほとんどここで確かめます（OS の結果を自由に作れて、問い合わせの記録も見られます）
    - ループバックの DNS サーバー: 127.x.x.x で答える補助スクリプトを立て、本物の Resolve-DnsName で動かします。
      偽の関数の形が本物とずれていないかを、ここで確かめます（補助スクリプトは終了時に止めます）
    - 本物の OS の結果: localhost とこの PC の名前だけを使います（どちらもネットワークに問い合わせません）

    環境変数 SystemRoot を差し替えて hosts を読ませる方法は使えません（Winsock を初期化できなくなり、OS の名前解決がすべて失敗します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーしたスクリプトを動かします。
    ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Test-NameResolution を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Test-NameResolution\Test-NameResolution.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Test-NameResolution_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Test-NameResolution.ps1'   # 本物のコピー
$fakeTool  = Join-Path $workRoot 'fake\Test-NameResolution.ps1'   # hosts の場所と OS の結果を差し替えたコピー
$wrapper   = Join-Path $workRoot 'wrapper.ps1'
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
$selfName  = [Net.Dns]::GetHostName()
# ループバックの DNS サーバーの番地。わざとバグのテストを並べて流しても重ならないよう、実行のたびに変える
$dnsNet    = '127.{0}.{1}' -f (Get-Random -Minimum 64 -Maximum 250), (Get-Random -Minimum 1 -Maximum 250)
$closedDns = "$dnsNet.59"   # 誰も待ち受けない番地
$script:failCount = 0
$script:runNumber = 0

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

function Write-Utf8([string]$FilePath, [string]$Text) {
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($FilePath))
    [IO.File]::WriteAllText($FilePath, $Text, $utf8NoBom)
}

# hosts の場所と OS の結果を差し替える。OS の結果は、ラッパーが定義する Get-FakeAddresses で作る
function Get-FakeToolText {
    $text = [IO.File]::ReadAllText($repoTool)
    foreach ($patch in @(
        @('$hostsPath  = Get-HostsPath', '$hostsPath  = $global:fake.HostsFile'),
        @('[Net.Dns]::GetHostAddresses($Name)', '(Get-FakeAddresses $Name)')
    )) {
        if (-not $text.Contains($patch[0])) { throw "書き換える場所が見つかりません: $($patch[0])" }
        $text = $text.Replace($patch[0], $patch[1])
    }
    $text
}

# 本物の環境では起きない条件を作るため、偽の環境のコピーの 1 か所を書き換える
function New-PatchedTool([string]$Name, [string]$From, [string]$To) {
    $text = Get-FakeToolText
    if (-not $text.Contains($From)) { throw "書き換える場所が見つかりません: $From" }
    $file = Join-Path $workRoot "patched\$Name\Test-NameResolution.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text.Replace($From, $To), $utf8Bom)
    $file
}

# 名前 → IP アドレス の表から、どのサーバーでも同じ答えを返すゾーンを作る
function New-Zone([hashtable]$Map) {
    $zone = @{}
    foreach ($key in $Map.Keys) { $zone[$key] = @{ '*' = @($Map[$key]) } }
    $zone
}

# 偽の環境の設定を書く。Spec のキー:
#   Args（ツールの引数）、Hosts（hosts の中身。文字列かバイト列。$null なら hosts なし）、Os（名前 → OS の結果）、
#   Zone（名前 → サーバー（* はどれでも）→ IP）、Cname（名前 → たどる名前）、Ptr（IP → 名前）、Ttl（名前 → 秒）、
#   Errors（サーバー → 種類（A、AAAA、PTR）→ エラーの番号か ID。Except に書いた名前は除く）、
#   Servers、SearchList、ConnSuffix、NoData9501、ServerThrow、InterfaceThrow、SuffixThrow、Slow（ミリ秒）、Culture、
#   RealDns（本物の Resolve-DnsName）、RealServers（本物の DNS サーバーの設定）、Tool（動かすコピー）、NoOut（-OutputFolder を足さない）
function New-FakeRun([hashtable]$Spec) {
    $script:runNumber++
    $root = Join-Path $workRoot ('r{0}' -f $script:runNumber)
    [void][IO.Directory]::CreateDirectory($root)
    $hostsFile = Join-Path $root 'hosts'
    if (-not $Spec.ContainsKey('Hosts')) { Write-Utf8 $hostsFile '' }
    elseif ($Spec.Hosts -is [byte[]]) { [IO.File]::WriteAllBytes($hostsFile, $Spec.Hosts) }
    elseif ($null -ne $Spec.Hosts) { Write-Utf8 $hostsFile $Spec.Hosts }
    $out = Join-Path $root 'out'
    $toolArgs = @{}
    foreach ($key in $Spec.Args.Keys) { $toolArgs[$key] = $Spec.Args[$key] }
    if (-not $toolArgs.ContainsKey('NoSave') -and -not $toolArgs.ContainsKey('OutputFolder') -and -not $Spec.ContainsKey('NoOut')) { $toolArgs['OutputFolder'] = $out }
    $full = @{ HostsFile = $hostsFile; QueryLog = (Join-Path $root 'queries.log'); Args = $toolArgs; ConnSuffix = ''; Tool = $fakeTool }
    foreach ($key in 'Os', 'Zone', 'Ptr', 'Cname', 'Ttl', 'Errors') { if ($Spec.ContainsKey($key)) { $full[$key] = $Spec[$key] } else { $full[$key] = @{} } }
    $full['Servers'] = [string[]]@('10.0.0.53')
    if ($Spec.ContainsKey('Servers')) { $full['Servers'] = [string[]]@($Spec.Servers) }
    $full['SearchList'] = [string[]]@()
    if ($Spec.ContainsKey('SearchList')) { $full['SearchList'] = [string[]]@($Spec.SearchList) }
    foreach ($key in 'ConnSuffix', 'NoData9501', 'ServerThrow', 'InterfaceThrow', 'SuffixThrow', 'Slow', 'Culture', 'RealDns', 'RealServers', 'Tool') {
        if ($Spec.ContainsKey($key)) { $full[$key] = $Spec[$key] }
    }
    $specFile = Join-Path $root 'spec.json'
    [IO.File]::WriteAllText($specFile, ($full | ConvertTo-Json -Depth 8), $utf8Bom)
    [pscustomobject]@{ SpecFile = $specFile; Out = $out; QueryLog = $full.QueryLog; HostsFile = $hostsFile }
}

function Read-FakeRun($Prepared, [string]$Log, $Code) {
    $dir = $null
    if (Test-Path -LiteralPath $Prepared.Out) { $dir = @(Get-ChildItem -LiteralPath $Prepared.Out -Directory | Where-Object { $_.Name -like 'NameResolution_*' } | ForEach-Object { $_.FullName })[0] }
    $rows = @()
    if ($dir -and (Test-Path -LiteralPath (Join-Path $dir 'results.csv'))) { $rows = @(Import-Csv -LiteralPath (Join-Path $dir 'results.csv') -Encoding UTF8) }
    $queries = @()
    if ([IO.File]::Exists($Prepared.QueryLog)) { $queries = @([IO.File]::ReadAllLines($Prepared.QueryLog)) }
    [pscustomobject]@{ Code = $Code; Log = $Log; Dir = $dir; Rows = $rows; Queries = $queries; Out = $Prepared.Out; HostsFile = $Prepared.HostsFile }
}

function Invoke-Fake([hashtable]$Spec) {
    $prepared = New-FakeRun $Spec
    $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $wrapper -SpecFile $prepared.SpecFile 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $code = $LASTEXITCODE
    Read-FakeRun $prepared $log $code
}

# 本物のコピーを、そのまま動かす（使う名前は localhost とこの PC の名前だけ。DNS サーバーはループバックだけ）
function Invoke-Real([string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string[]]$Lines) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { $utf8NoBom } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]](@('@echo off') + $Lines), $encoding)
    $log = ConvertTo-FlatText ((@(& cmd.exe /d /c $batch 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
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

function Get-Row($Run, [string]$Name) {
    @($Run.Rows | Where-Object { ($_.'名前' -replace '^="', '' -replace '"$', '') -eq $Name })[0]
}

function Get-Cell($Run, [string]$Name, [string]$Column) {
    $row = Get-Row $Run $Name
    if ($null -eq $row) { return '（行なし）' }
    ([string]$row.$Column) -replace '^="', '' -replace '"$', ''
}

function Get-Verdict($Run, [string]$Name) { Get-Cell $Run $Name '判定' }

function Test-Note($Run, [string]$Name, [string]$Part) { Test-Contains (Get-Cell $Run $Name '備考') $Part }

function Get-QueryCount($Run, [string]$Pattern) { @($Run.Queries | Where-Object { $_ -like $Pattern }).Count }

# 長いパスのファイルを作る（5.1 は \\?\ を付けないと書けない）
function Write-LongFile([string]$FilePath, [string]$Text) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllText('\\?\' + $FilePath, $Text, $utf8NoBom)
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$dnsServer = $null
$dnsStop   = Join-Path $workRoot 'dns-stop'

try {
    #region ダミー環境

    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool
    [void][IO.Directory]::CreateDirectory((Split-Path $fakeTool))
    [IO.File]::WriteAllText($fakeTool, (Get-FakeToolText), $utf8Bom)

    $wrapperCode = @'
param([string]$SpecFile)
$spec = [IO.File]::ReadAllText($SpecFile) | ConvertFrom-Json
$global:fake = $spec
if ($spec.PSObject.Properties['Culture']) { [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($spec.Culture) }

function global:Get-FakeAddresses([string]$Name) {
    $entry = $global:fake.Os.PSObject.Properties[$Name]
    if (-not $entry) { throw (New-Object Net.Sockets.SocketException 11001) }
    $list = New-Object 'System.Collections.Generic.List[Net.IPAddress]'
    foreach ($ip in @($entry.Value)) { $list.Add([Net.IPAddress]::Parse($ip)) }
    return , $list.ToArray()
}

# 本物の Resolve-DnsName を使うときは、先に読み込む（あとから自動で読み込まれると、同じモジュールの偽の関数が本物に置き換わる）
if ($spec.PSObject.Properties['RealDns']) { Import-Module DnsClient }

if (-not $spec.PSObject.Properties['RealServers']) {
    function global:Get-NetIPInterface {
        if ($global:fake.PSObject.Properties['InterfaceThrow']) { throw 'インターフェイスを読めません' }
        [pscustomobject]@{ InterfaceIndex = 1 }
    }
    function global:Get-DnsClientServerAddress {
        if ($global:fake.PSObject.Properties['ServerThrow']) { throw (New-Object InvalidOperationException "サーバーの設定を読めません`r`n2 行目") }
        [pscustomobject]@{ InterfaceIndex = 1; ServerAddresses = @(@($global:fake.Servers) + 'fec0:0:0:ffff::1') }
        [pscustomobject]@{ InterfaceIndex = 2; ServerAddresses = @('10.8.8.8') }
    }
    function global:Get-DnsClientGlobalSetting {
        if ($global:fake.PSObject.Properties['SuffixThrow']) { throw '検索一覧を読めません' }
        [pscustomobject]@{ SuffixSearchList = @($global:fake.SearchList) }
    }
    function global:Get-DnsClient {
        if ($global:fake.PSObject.Properties['SuffixThrow']) { throw '接続ごとのサフィックスを読めません' }
        [pscustomobject]@{ ConnectionSpecificSuffix = $global:fake.ConnSuffix }
    }
}

# 本物のエラーは Win32Exception で、文は「名前 : 内容」の形。文字列を渡すと、ID だけで見分ける形のエラーにする
function global:Stop-FakeQuery([string]$Name, $Code) {
    if ($Code -is [string]) { throw (New-Object Management.Automation.ErrorRecord (New-Object Exception "$Name : 別の形のエラー"), "$Code,Microsoft.DnsClient.Commands.ResolveDnsName", ([Management.Automation.ErrorCategory]::NotSpecified), $null) }
    $text = if ($Code -eq 9998) { "$Name : 制御$([char]27)文字" } else { "$Name : 何かのエラー`r`n" }
    throw (New-Object ComponentModel.Win32Exception ([int]$Code), $text)
}

if (-not $spec.PSObject.Properties['RealDns']) {
    function global:Resolve-DnsName {
        param($Name, $Type, [switch]$DnsOnly, [switch]$NoHostsFile, $Server, $ErrorAction)
        [IO.File]::AppendAllText($global:fake.QueryLog, ('{0} {1} {2} DnsOnly={3} NoHostsFile={4}' -f $Server, $Type, $Name, $DnsOnly.IsPresent, $NoHostsFile.IsPresent) + "`r`n")
        if ($global:fake.PSObject.Properties['Slow']) { Start-Sleep -Milliseconds $global:fake.Slow }
        $conf = $global:fake.Errors.PSObject.Properties[[string]$Server]
        if ($conf) {
            $except = $conf.Value.PSObject.Properties['Except']
            if (-not ($except -and @($except.Value) -contains $Name)) {
                $code = $conf.Value.PSObject.Properties[$Type]
                if ($code) { Stop-FakeQuery $Name $code.Value }
            }
        }
        if ($Type -eq 'PTR') {
            $ptr = $global:fake.Ptr.PSObject.Properties[[string]$Name]
            if (-not $ptr) { Stop-FakeQuery $Name 9003 }
            foreach ($n in @($ptr.Value)) { [pscustomobject]@{ Section = 'Answer'; Type = 'PTR'; Name = "$Name.in-addr.arpa"; NameHost = $n; TTL = 60 } }
            return
        }
        $entry = $global:fake.Zone.PSObject.Properties[[string]$Name]
        if (-not $entry) { Stop-FakeQuery $Name 9003 }
        $perServer = $entry.Value.PSObject.Properties[[string]$Server]
        if (-not $perServer) { $perServer = $entry.Value.PSObject.Properties['*'] }
        if (-not $perServer) { Stop-FakeQuery $Name 9003 }
        $wantV6 = $Type -eq 'AAAA'
        $ips = @(@($perServer.Value) | Where-Object { ([string]$_).Contains(':') -eq $wantV6 })
        if ($ips.Count -eq 0) {
            if ($global:fake.PSObject.Properties['NoData9501']) { Stop-FakeQuery $Name 9501 }
            # 本物は、登録がないとエラーにせず、権威の SOA を Authority に入れて返す
            [pscustomobject]@{ Section = 'Authority'; Type = 'SOA'; Name = 'test'; TTL = 60 }
            return
        }
        $owner = $Name
        $chain = $global:fake.Cname.PSObject.Properties[[string]$Name]
        if ($chain) {
            foreach ($next in @($chain.Value)) {
                [pscustomobject]@{ Section = 'Answer'; Type = 'CNAME'; Name = $owner; NameHost = $next; TTL = 300 }
                $owner = $next
            }
        }
        $ttlEntry = $global:fake.Ttl.PSObject.Properties[[string]$Name]
        $ttl = if ($ttlEntry) { $ttlEntry.Value } else { 120 }
        foreach ($ip in $ips) { [pscustomobject]@{ Section = 'Answer'; Type = $Type; Name = $owner; IPAddress = [string]$ip; TTL = $ttl } }
    }
}

$named = @{}
foreach ($p in $spec.Args.PSObject.Properties) { $named[$p.Name] = $p.Value }
& $spec.Tool @named
exit $LASTEXITCODE
'@
    [IO.File]::WriteAllText($wrapper, $wrapperCode, $utf8Bom)

    #endregion

    Write-Section 'OK と、保存するファイル'
    $run = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1') }) }
    Assert-Equal '判定と終了コード' ('{0}/{1}' -f (Get-Verdict $run 'web.test'), $run.Code) 'OK/0'
    Assert-Equal '引けた元' (Get-Cell $run 'web.test' '引けた元') 'DNS'
    Assert-Equal 'CSV の列' (($run.Rows[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',') '判定,名前,引けた元,OS の結果,hosts の値,hosts の行番号,DNS の結果,解決した名前,CNAME,TTL,期待する IP,逆引き,メモ,備考'
    Assert-Equal '保存するファイル' ((Get-ChildItem -LiteralPath $run.Dir | ForEach-Object { $_.Name } | Sort-Object) -join ',') 'results.csv,summary.txt'
    Assert-Equal 'CSV の BOM' ([IO.File]::ReadAllBytes((Join-Path $run.Dir 'results.csv'))[0..2] -join ',') '239,187,191'
    Assert-Equal 'summary.txt の BOM' ([IO.File]::ReadAllBytes((Join-Path $run.Dir 'summary.txt'))[0..2] -join ',') '239,187,191'
    Assert-Equal 'TTL' (Get-Cell $run 'web.test' 'TTL') '120'
    Assert-Equal '解決した名前' (Get-Cell $run 'web.test' '解決した名前') 'web.test'
    Assert-Equal 'DNS の結果（サーバーごと）' (Get-Cell $run 'web.test' 'DNS の結果') '10.0.0.53: 10.0.0.1'
    Assert-Equal 'IP アドレスは Excel が数値にしない形' (Get-Row $run 'web.test').'OS の結果' '="10.0.0.1"'
    Assert-Equal 'フォルダの名前' ((Split-Path -Leaf $run.Dir) -match ('^NameResolution_{0}_\d{{8}}_\d{{6}}$' -f [regex]::Escape($env:COMPUTERNAME))) 'True'
    Assert-Equal '聞いた問い合わせ（A と AAAA を 1 回ずつ、この DNS サーバーだけ）' (($run.Queries | ForEach-Object { ($_ -split ' ')[0..2] -join ' ' }) -join ',') '10.0.0.53 A web.test,10.0.0.53 AAAA web.test'
    Assert-Equal 'hosts とキャッシュを使わない指定' (Get-QueryCount $run '* DnsOnly=True NoHostsFile=True') 2
    foreach ($part in @('名前解決確認 (Test-NameResolution 1.0.0)', '（管理者: ', '対象 1 件   DNS サーバー: 10.0.0.53（OS の設定）', ('hosts: {0}（有効な行 0 件）' -f $run.HostsFile), '結果: OK 1 件  注意 0 件  引けない 0 件  所要 ', ('保存先: {0}' -f $run.Dir))) {
        Assert-Equal "画面の表示: $part" (Test-Contains $run.Log $part) 'True'
    }
    $summary = ConvertTo-FlatText ([IO.File]::ReadAllText((Join-Path $run.Dir 'summary.txt')))
    foreach ($part in @('名前解決確認 (Test-NameResolution', ' OK        web.test', '結果: OK 1 件', ('保存先: {0}' -f $run.Dir))) {
        Assert-Equal "summary.txt: $part" (Test-Contains $summary $part) 'True'
    }

    $multi = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('2001:db8::1', '10.0.0.2', '10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.2', '10.0.0.1', '2001:db8::1') }); Ttl = @{ 'web.test' = 3600 } }
    Assert-Equal '複数の IP アドレス（順番をそろえる）' (Get-Cell $multi 'web.test' 'OS の結果') '10.0.0.1, 10.0.0.2, 2001:db8::1'
    Assert-Equal '複数の IP アドレス（DNS）' (Get-Cell $multi 'web.test' 'DNS の結果') '10.0.0.53: 10.0.0.1, 10.0.0.2, 2001:db8::1'
    Assert-Equal '複数の IP アドレスでも OK' (Get-Verdict $multi 'web.test') 'OK'
    Assert-Equal 'TTL は 3 桁区切り' (Get-Cell $multi 'web.test' 'TTL') '3,600'

    $cname = Invoke-Fake @{ Args = @{ Target = 'alias.test' }; Os = @{ 'alias.test' = @('10.0.0.4') }; Zone = (New-Zone @{ 'alias.test' = @('10.0.0.4') }); Cname = @{ 'alias.test' = @('mid.test', 'end.test') }; Ttl = @{ 'alias.test' = 4000 } }
    Assert-Equal 'CNAME を順につなぐ' (Get-Cell $cname 'alias.test' 'CNAME') 'alias.test → mid.test → end.test'
    Assert-Equal '解決した名前は最後の名前' (Get-Cell $cname 'alias.test' '解決した名前') 'end.test'
    Assert-Equal 'TTL は一番短いもの（CNAME の 300）' (Get-Cell $cname 'alias.test' 'TTL') '300'

    Write-Section '表の IP アドレスの幅'
    # この PC の名前のように IP アドレスが多いと、列の幅が広がってすべての行が折り返される。表では半角 40 文字分を超える分を「ほか N 件」にする
    $manyIps = @(11..22 | ForEach-Object { "10.0.0.$_" })
    $wide = Invoke-Fake @{ Args = @{ Target = 'many.test' }; Os = @{ 'many.test' = $manyIps }; Zone = (New-Zone @{ 'many.test' = $manyIps }) }
    $wideLine = ' OK        many.test  DNS       10.0.0.11, 10.0.0.12 ほか 10 件  10.0.0.11, 10.0.0.12 ほか 10 件'
    Assert-Equal '幅を超える分は「ほか N 件」（summary.txt。列の幅も省いた値に合わせる）' (@([IO.File]::ReadAllLines((Join-Path $wide.Dir 'summary.txt')) | Where-Object { $_.Contains('many.test') }) -join '|') $wideLine
    Assert-Equal '幅を超える分は「ほか N 件」（画面）' (Test-Contains $wide.Log $wideLine) 'True'
    Assert-Equal 'CSV にはすべて書く' ('{0}/{1}' -f (Get-Cell $wide 'many.test' 'OS の結果'), (Get-Cell $wide 'many.test' 'DNS の結果')) ('{0}/10.0.0.53: {0}' -f ($manyIps -join ', '))
    $edgeIps = @{
        'e40.test'   = @('10.0.0.101', '2001:db8::aaaa:bbbb:cccc:ddd')    # ちょうど 40 文字
        'e41.test'   = @('10.0.0.101', '2001:db8::aaaa:bbbb:cccc:dddd')   # 41 文字
        'v6.test'    = @('2001:db8:aaaa:bbbb:cccc:dddd:eeee:1', '2001:db8:aaaa:bbbb:cccc:dddd:eeee:2', '2001:db8:aaaa:bbbb:cccc:dddd:eeee:3')
        'three.test' = @('10.0.0.11', '10.0.0.12', '10.0.0.13')           # 「ほか」を付けると 40 文字を超えるが、付けなければ入る
    }
    $edge = Invoke-Fake @{ Args = @{ Target = 'e40.test,e41.test,v6.test,three.test' }; Os = $edgeIps; Zone = (New-Zone $edgeIps) }
    $edgeLines = [IO.File]::ReadAllLines((Join-Path $edge.Dir 'summary.txt'))
    $edgeLine = @{}
    foreach ($name in $edgeIps.Keys) { $edgeLine[$name] = [string](@($edgeLines | Where-Object { $_.Contains(" $name ") })[0]) }
    Assert-Equal 'ちょうど 40 文字は省かない' (Test-Contains $edgeLine['e40.test'] '10.0.0.101, 2001:db8::aaaa:bbbb:cccc:ddd  ') 'True'
    Assert-Equal '41 文字は省く' (Test-Contains $edgeLine['e41.test'] '10.0.0.101 ほか 1 件  ') 'True'
    Assert-Equal '最初の 1 つは 40 文字を超えても出す' (Test-Contains $edgeLine['v6.test'] '2001:db8:aaaa:bbbb:cccc:dddd:eeee:1 ほか 2 件  ') 'True'
    Assert-Equal '最後の 1 つは「ほか」なしで入れば出す' ('{0}/{1}' -f (Test-Contains $edgeLine['three.test'] '10.0.0.11, 10.0.0.12, 10.0.0.13  '), (Test-Contains $edgeLine['three.test'] 'ほか')) 'True/False'

    Write-Section '引けない'
    $fail = Invoke-Fake @{ Args = @{ Target = 'none.test' } }
    Assert-Equal '引けない' ('{0}/{1}/{2}' -f (Get-Verdict $fail 'none.test'), (Get-Cell $fail 'none.test' '引けた元'), $fail.Code) '引けない/-/2'
    Assert-Equal '備考' (Test-Note $fail 'none.test' 'DNS: 名前が存在しない') 'True'
    Assert-Equal 'OS の結果は空欄' (Get-Cell $fail 'none.test' 'OS の結果') ''
    Assert-Equal '画面では -' (Test-Contains $fail.Log ' 引けない  none.test  -         -          -') 'True'
    Assert-Equal 'DNS の結果' (Get-Cell $fail 'none.test' 'DNS の結果') '10.0.0.53: 名前が存在しない'
    Assert-Equal '結果の行' (Test-Contains $fail.Log '結果: OK 0 件  注意 0 件  引けない 1 件') 'True'
    $noData = Invoke-Fake @{ Args = @{ Target = 'none.test' }; Zone = (New-Zone @{ 'none.test' = @() }) }
    Assert-Equal 'A / AAAA の登録がない（本物の形: SOA だけが返る）' ('{0}/{1}' -f (Get-Cell $noData 'none.test' 'DNS の結果'), (Test-Note $noData 'none.test' 'DNS: A / AAAA の登録がない')) '10.0.0.53: A / AAAA の登録がない/True'
    $noData2 = Invoke-Fake @{ Args = @{ Target = 'none.test' }; Zone = (New-Zone @{ 'none.test' = @() }); NoData9501 = $true }
    Assert-Equal 'A / AAAA の登録がない（エラー 9501 の形）' (Get-Cell $noData2 'none.test' 'DNS の結果') '10.0.0.53: A / AAAA の登録がない'
    $reg = Invoke-Fake @{ Args = @{ Target = 'dns-only.test' }; Zone = (New-Zone @{ 'dns-only.test' = @('10.0.0.1') }) }
    Assert-Equal 'DNS にあっても OS で引けないと「引けない」' ('{0}/{1}/{2}' -f (Get-Verdict $reg 'dns-only.test'), (Test-Note $reg 'dns-only.test' 'DNS には登録があるが、OS では引けない'), (Get-Cell $reg 'dns-only.test' '引けた元')) '引けない/True/DNS'
    $noServerFail = Invoke-Fake @{ Args = @{ Target = 'none.test' }; Servers = @() }
    Assert-Equal 'DNS サーバーがなく OS でも引けないと、OS の理由を書く' ('{0}/{1}' -f (Get-Verdict $noServerFail 'none.test'), ((Get-Cell $noServerFail 'none.test' '備考') -match 'DNS サーバーに聞けなかった、.+')) '引けない/True'

    Write-Section 'キャッシュと、OS の結果'
    $stale = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.1.1.1') }) }
    Assert-Equal 'OS の結果が DNS と違うとキャッシュの案内' ('{0}/{1}/{2}' -f (Get-Verdict $stale 'web.test'), (Test-Note $stale 'web.test' 'キャッシュが古い可能性（ipconfig /flushdns で確かめてください）'), $stale.Code) '注意/True/2'
    Assert-Equal '結果の行（注意）' (Test-Contains $stale.Log '結果: OK 0 件  注意 1 件  引けない 0 件') 'True'
    $staleServer = Invoke-Fake @{ Args = @{ Target = 'web.test'; Server = '10.0.0.54' }; Os = @{ 'web.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.1.1.1') }) }
    Assert-Equal '-Server のときはキャッシュのせいにしない' ('{0}/{1}' -f (Test-Note $staleServer 'web.test' 'OS の結果と、指定した DNS サーバーの答えが違う'), (Test-Note $staleServer 'web.test' 'キャッシュ')) 'True/False'
    $subset = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1', '10.0.0.2') }) }
    Assert-Equal 'OS の値が DNS の答えに含まれれば違わない' (Get-Verdict $subset 'web.test') 'OK'
    $extraOs = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('10.0.0.1', '10.0.0.9') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1', '10.0.0.2') }) }
    Assert-Equal 'DNS にない値が OS にあれば違う' (Test-Note $extraOs 'web.test' 'キャッシュが古い可能性') 'True'
    $v6Stale = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('2001:db8::2') }; Zone = (New-Zone @{ 'web.test' = @('2001:db8::1') }) }
    Assert-Equal 'IPv6 も比べる' (Test-Note $v6Stale 'web.test' 'キャッシュが古い可能性') 'True'
    $otherKind = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('2001:db8::2') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1') }) }
    Assert-Equal '種類（IPv4 か IPv6）が違う答えは比べない' (Get-Verdict $otherKind 'web.test') 'OK'
    $linkLocal = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('fe80::1%3', '10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1', '2001:db8::1') }) }
    Assert-Equal 'リンクローカルは比べない' (Get-Verdict $linkLocal 'web.test') 'OK'
    Assert-Equal 'スコープ（%3）は外す' (Get-Cell $linkLocal 'web.test' 'OS の結果') '10.0.0.1, fe80::1'
    $self = Invoke-Fake @{ Args = @{ Target = $selfName }; Os = @{ $selfName = @('10.0.0.1') }; Zone = (New-Zone @{ $selfName = @('10.1.1.1') }) }
    Assert-Equal 'この PC 自身の名前は OS の結果と比べない' ('{0}/{1}' -f (Get-Verdict $self $selfName), (Test-Note $self $selfName 'キャッシュ')) 'OK/False'
    $selfUpper = Invoke-Fake @{ Args = @{ Target = $selfName.ToLowerInvariant() }; Os = @{ $selfName = @('10.0.0.1') }; Zone = (New-Zone @{ $selfName = @('10.1.1.1') }) }
    Assert-Equal 'この PC の名前は大文字と小文字を区別しない' (Get-Verdict $selfUpper $selfName.ToLowerInvariant()) 'OK'
    $selfDomainTool = New-PatchedTool 'selfdomain' '$selfDomain = [Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName' '$selfDomain = ''corp.test'''
    $selfFqdn = "$selfName.corp.test"
    $selfDomain = Invoke-Fake @{ Args = @{ Target = $selfFqdn }; Os = @{ $selfFqdn = @('10.0.0.1') }; Zone = (New-Zone @{ $selfFqdn = @('10.1.1.1') }); Tool = $selfDomainTool }
    Assert-Equal 'この PC の FQDN も比べない' (Get-Verdict $selfDomain $selfFqdn) 'OK'
    $other = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('10.0.0.1') } }
    Assert-Equal 'DNS になく OS だけが解決' ('{0}/{1}/{2}' -f (Get-Verdict $other 'web.test'), (Get-Cell $other 'web.test' '引けた元'), (Test-Note $other 'web.test' 'DNS にはない（キャッシュ、NetBIOS、LLMNR の可能性）')) 'OK/その他/True'

    Write-Section 'hosts'
    $diff = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = "# 1 行目`r`n10.9.9.9 web.test # コメント`r`n"; Os = @{ 'web.test' = @('10.9.9.9') }; Zone = (New-Zone @{ 'web.test' = @('10.1.1.1') }) }
    Assert-Equal 'hosts と DNS が違う' ('{0}/{1}/{2}' -f (Get-Verdict $diff 'web.test'), (Get-Cell $diff 'web.test' '引けた元'), (Test-Note $diff 'web.test' 'hosts と DNS が違う（hosts の 2 行目）')) '注意/hosts/True'
    Assert-Equal 'hosts の値と行番号' ('{0}/{1}' -f (Get-Cell $diff 'web.test' 'hosts の値'), (Get-Cell $diff 'web.test' 'hosts の行番号')) '10.9.9.9/2'
    Assert-Equal 'hosts の行番号は Excel が数値にしない形' (Get-Row $diff 'web.test').'hosts の行番号' '="2"'
    Assert-Equal '有効な行の数' (Test-Contains $diff.Log '（有効な行 1 件）') 'True'
    $inc = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = '10.1.1.1 web.test'; Os = @{ 'web.test' = @('10.1.1.1') }; Zone = (New-Zone @{ 'web.test' = @('10.1.1.1', '10.1.1.2') }) }
    Assert-Equal 'hosts の値が DNS の答えに含まれれば違わない' (Get-Verdict $inc 'web.test') 'OK'
    $v6 = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = '10.1.1.1 web.test'; Os = @{ 'web.test' = @('10.1.1.1') }; Zone = (New-Zone @{ 'web.test' = @('10.1.1.1', '2001:db8::1') }) }
    Assert-Equal 'IPv4 だけの hosts と両方の DNS は違わない' (Get-Verdict $v6 'web.test') 'OK'
    $v6Only = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = '2001:db8::9 web.test'; Os = @{ 'web.test' = @('2001:db8::9') }; Zone = (New-Zone @{ 'web.test' = @('10.1.1.1') }) }
    Assert-Equal 'IPv6 だけの hosts と IPv4 だけの DNS は比べない' (Get-Verdict $v6Only 'web.test') 'OK'
    $v6Form = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = '2001:DB8:0:0::1 web.test'; Os = @{ 'web.test' = @('2001:db8::1') }; Zone = (New-Zone @{ 'web.test' = @('2001:db8::1') }) }
    Assert-Equal 'IPv6 の書き方の違いは同じ' ('{0}/{1}' -f (Get-Verdict $v6Form 'web.test'), (Get-Cell $v6Form 'web.test' 'hosts の値')) 'OK/2001:db8::1'
    $multiHosts = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = "10.1.1.1 web.test`r`n10.1.1.2 web.test`r`n"; Os = @{ 'web.test' = @('10.1.1.1', '10.1.1.2') }; Zone = (New-Zone @{ 'web.test' = @('10.1.1.1', '10.1.1.2') }) }
    Assert-Equal 'hosts に同じ名前が複数' ('{0}/{1}' -f (Get-Verdict $multiHosts 'web.test'), (Test-Note $multiHosts 'web.test' 'hosts に同じ名前が複数ある（1 行目と 2 行目）')) '注意/True'
    Assert-Equal '複数の行番号' (Get-Cell $multiHosts 'web.test' 'hosts の行番号') '1, 2'
    $mixed = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = "10.1.1.1 web.test`r`n2001:db8::1 web.test`r`n"; Os = @{ 'web.test' = @('10.1.1.1') }; Zone = (New-Zone @{ 'web.test' = @('10.1.1.1') }) }
    Assert-Equal 'IPv4 と IPv6 が 1 つずつなら複数としない' (Test-Note $mixed 'web.test' '複数') 'False'
    $sameIp = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = "10.1.1.1 web.test`r`n10.1.1.1 web.test`r`n"; Os = @{ 'web.test' = @('10.1.1.1') }; Zone = (New-Zone @{ 'web.test' = @('10.1.1.1') }) }
    Assert-Equal '同じ IP アドレスの重複は複数としない' (Test-Note $sameIp 'web.test' '複数') 'False'
    $hostsOnly = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = '10.1.1.1 web.test'; Os = @{ 'web.test' = @('10.1.1.1') } }
    Assert-Equal 'hosts にだけある名前は OK' ('{0}/{1}/{2}' -f (Get-Verdict $hostsOnly 'web.test'), (Get-Cell $hostsOnly 'web.test' '備考'), (Get-Cell $hostsOnly 'web.test' '引けた元')) 'OK/DNS にはない/hosts'
    $hostsDead = Invoke-Fake @{ Args = @{ Target = 'web.test'; Server = '10.0.0.99' }; Hosts = '10.1.1.1 web.test'; Os = @{ 'web.test' = @('10.1.1.1') }; Errors = @{ '10.0.0.99' = @{ A = 1460; AAAA = 1460 } } }
    Assert-Equal 'hosts にあっても DNS に聞けなければ注意' ('{0}/{1}' -f (Get-Verdict $hostsDead 'web.test'), (Get-Cell $hostsDead 'web.test' '引けた元')) '注意/hosts'
    $format = Invoke-Fake @{ Args = @{ Target = 'a.test,b.test,c.test,d.test' }; Hosts = "`t 10.1.1.1`ta.test`t b.test  # 後ろ 10.2.2.2 d.test`r`n# 10.3.3.3 d.test`r`n１０.０.０.１ d.test`r`nnot-ip d.test`r`n10.4.4.4`r`n10.5.5.5 C.TEST`n10.6.6.6 x.test"; Os = @{ 'a.test' = @('10.1.1.1'); 'b.test' = @('10.1.1.1'); 'c.test' = @('10.5.5.5'); 'd.test' = @('10.0.0.1') } }
    Assert-Equal 'タブと空白で区切る・1 行に複数の名前' ('{0}/{1}' -f (Get-Cell $format 'a.test' 'hosts の値'), (Get-Cell $format 'b.test' 'hosts の値')) '10.1.1.1/10.1.1.1'
    Assert-Equal '# より後ろと、IP アドレスでない行は読まない' (Get-Cell $format 'd.test' 'hosts の値') ''
    Assert-Equal '名前の大文字と小文字は区別しない' (Get-Cell $format 'c.test' 'hosts の値') '10.5.5.5'
    Assert-Equal '止まらない' (Test-Contains $format.Log 'エラーで止まりました') 'False'
    Assert-Equal '有効な行は IP アドレスと名前のある行だけ' (Test-Contains $format.Log '（有効な行 3 件）') 'True'
    $sjisHosts = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = [byte[]][Text.Encoding]::GetEncoding(932).GetBytes("# 日本語のコメント`r`n10.1.1.1 web.test`r`n"); Os = @{ 'web.test' = @('10.1.1.1') } }
    Assert-Equal 'Shift_JIS の hosts を読む' (Get-Cell $sjisHosts 'web.test' 'hosts の行番号') '2'
    $bomHosts = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = [byte[]]($utf8Bom.GetPreamble() + $utf8NoBom.GetBytes("10.1.1.1 web.test`r`n")); Os = @{ 'web.test' = @('10.1.1.1') } }
    Assert-Equal 'BOM 付きの hosts の 1 行目を読む' (Get-Cell $bomHosts 'web.test' 'hosts の値') '10.1.1.1'
    $noHosts = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Hosts = $null; Os = @{ 'web.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1') }) }
    Assert-Equal 'hosts を読めなくても続ける' ('{0}/{1}/{2}' -f (Test-Contains $noHosts.Log '※ hosts を読めなかったため、hosts なしとして調べました（'), (Get-Verdict $noHosts 'web.test'), $noHosts.Code) 'True/OK/0'
    $shortHosts = Invoke-Fake @{ Args = @{ Target = 'web' }; Hosts = '10.1.1.1 web.test'; SearchList = @('test'); Os = @{ 'web' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1') }) }
    Assert-Equal 'hosts は名前が完全に一致するものだけ（短い名前にサフィックスを付けない）' (Get-Cell $shortHosts 'web' 'hosts の値') ''

    Write-Section 'DNS サーバー'
    $two = @('10.0.0.53', '10.0.0.54')
    $split = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Servers = $two; Os = @{ 'web.test' = @('10.5.5.1') }; Zone = @{ 'web.test' = @{ '10.0.0.53' = @('10.5.5.1'); '10.0.0.54' = @('10.5.5.2') } } }
    Assert-Equal 'サーバーによって答えが違う' ('{0}/{1}/{2}' -f (Get-Verdict $split 'web.test'), (Test-Note $split 'web.test' 'DNS サーバーによって答えが違う'), (Test-Contains $split.Log 'サーバーで違う')) '注意/True/True'
    Assert-Equal 'サーバーごとに CSV に書く' (Get-Cell $split 'web.test' 'DNS の結果') '10.0.0.53: 10.5.5.1 / 10.0.0.54: 10.5.5.2'
    Assert-Equal '対象の行に 2 台' (Test-Contains $split.Log 'DNS サーバー: 10.0.0.53, 10.0.0.54（OS の設定）') 'True'
    $half = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Servers = $two; Os = @{ 'web.test' = @('10.5.5.1') }; Zone = @{ 'web.test' = @{ '10.0.0.53' = @('10.5.5.1') } } }
    Assert-Equal '片方だけが答える' ('{0}/{1}' -f (Test-Note $half 'web.test' 'DNS サーバーによって答えが違う'), (Get-Cell $half 'web.test' 'DNS の結果')) 'True/10.0.0.53: 10.5.5.1 / 10.0.0.54: 名前が存在しない'
    $roundRobin = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Servers = $two; Os = @{ 'web.test' = @('10.1.1.1', '10.1.1.2') }; Zone = @{ 'web.test' = @{ '10.0.0.53' = @('10.1.1.1', '10.1.1.2'); '10.0.0.54' = @('10.1.1.2', '10.1.1.1') } } }
    Assert-Equal '順番だけの違いは同じ' ('{0}/{1}' -f (Get-Verdict $roundRobin 'web.test'), (Test-Contains $roundRobin.Log 'サーバーで違う')) 'OK/False'

    $deadErrors = @{ '10.0.0.99' = @{ A = 1460; AAAA = 1460; PTR = 1460 } }
    $deadOs = @{ 'a.test' = @('10.0.0.1'); 'b.test' = @('10.0.0.1'); 'c.test' = @('10.0.0.1'); 'd.test' = @('10.0.0.1'); 'e.test' = @('10.0.0.1') }
    $dead = Invoke-Fake @{ Args = @{ Target = 'a.test,b.test,c.test,d.test'; Server = '10.0.0.99,10.0.0.53' }; Os = $deadOs; Errors = $deadErrors; Zone = (New-Zone @{ 'a.test' = @('10.0.0.1'); 'd.test' = @('10.0.0.1') }) }
    Assert-Equal '応答しないサーバーは注意' ('{0}/{1}' -f (Get-Verdict $dead 'a.test'), (Test-Note $dead 'a.test' 'DNS サーバー 10.0.0.99 が応答しない')) '注意/True'
    Assert-Equal '応答なしと CSV に書く' (Get-Cell $dead 'a.test' 'DNS の結果') '10.0.0.99: 応答なし / 10.0.0.53: 10.0.0.1'
    Assert-Equal '3 回続くと聞かない（問い合わせの数）' ('{0}/{1}' -f (Get-QueryCount $dead '10.0.0.99 * c.test *'), (Get-QueryCount $dead '10.0.0.99 * d.test *')) '2/0'
    Assert-Equal '4 件目は聞いていないと書く' ('{0}/{1}' -f (Get-Cell $dead 'd.test' 'DNS の結果'), (Test-Note $dead 'd.test' 'DNS サーバー 10.0.0.99 は応答しないので聞いていない')) '10.0.0.99: 応答しないので聞いていない / 10.0.0.53: 10.0.0.1/True'
    Assert-Equal '聞いていない 4 件目も注意' (Get-Verdict $dead 'd.test') '注意'
    $streakErrors = @{ '10.0.0.99' = @{ A = 1460; AAAA = 1460; Except = @('c.test') } }
    $streak = Invoke-Fake @{ Args = @{ Target = 'a.test,b.test,c.test,d.test,e.test'; Server = '10.0.0.99' }; Os = $deadOs; Errors = $streakErrors; Zone = (New-Zone @{ 'c.test' = @('10.0.0.1') }) }
    Assert-Equal '答えがあると回数を数え直す' (Get-QueryCount $streak '10.0.0.99 * e.test *') 2
    $aaaaDead = Invoke-Fake @{ Args = @{ Target = 'a.test,b.test,c.test,d.test'; Server = '10.0.0.77' }; Os = $deadOs; Errors = @{ '10.0.0.77' = @{ AAAA = 1460 } }; Zone = (New-Zone @{ 'a.test' = @('10.0.0.1'); 'b.test' = @('10.0.0.1'); 'c.test' = @('10.0.0.1'); 'd.test' = @('10.0.0.1') }) }
    Assert-Equal 'AAAA だけ応答しないサーバーの答えを捨てない' ('{0}/{1}/{2}' -f (Get-Cell $aaaaDead 'a.test' '引けた元'), (Test-Note $aaaaDead 'a.test' 'DNS サーバー 10.0.0.77: AAAA は応答なし'), (Get-Verdict $aaaaDead 'a.test')) 'DNS/True/注意'
    Assert-Equal 'DNS の結果に片方の失敗を添える' (Get-Cell $aaaaDead 'a.test' 'DNS の結果') '10.0.0.77: 10.0.0.1（AAAA は応答なし）'
    Assert-Equal 'AAAA だけ応答しないサーバーは聞き続ける' (Get-QueryCount $aaaaDead '10.0.0.77 AAAA d.test *') 1
    $aDeadOnly = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = '10.0.0.77' }; Os = $deadOs; Errors = @{ '10.0.0.77' = @{ A = 1460 } } }
    Assert-Equal 'A が応答せず AAAA に答えがないと、「ない」と言い切らない' ('{0}/{1}' -f (Get-Cell $aDeadOnly 'a.test' 'DNS の結果'), (Get-Cell $aDeadOnly 'a.test' '引けた元')) '10.0.0.77: A は応答なし/OS のみ'

    # DNS が動いていないサーバーや経路のないサーバーも、タイムアウトと同じく「応答しない」
    $noResponseCodes = 1231, 1232, 1460, 10050, 10051, 10054, 10060, 10061, 10064, 10065
    $codeErrors = @{}
    $codeServers = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt $noResponseCodes.Count; $i++) {
        $codeErrors["10.1.0.$($i + 1)"] = @{ A = $noResponseCodes[$i]; AAAA = $noResponseCodes[$i] }
        $codeServers.Add("10.1.0.$($i + 1)")
    }
    $codes = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = ($codeServers -join ',') }; Os = $deadOs; Errors = $codeErrors }
    for ($i = 0; $i -lt $noResponseCodes.Count; $i++) {
        Assert-Equal "エラー $($noResponseCodes[$i]) は応答しない" (Test-Contains (Get-Cell $codes 'a.test' 'DNS の結果') "10.1.0.$($i + 1): 応答なし") 'True'
    }
    Assert-Equal '応答しないサーバーは 1 つの備考にまとめる' (Get-Cell $codes 'a.test' '備考') ('DNS サーバー {0} が応答しない' -f ($codeServers -join '、'))
    $deadTwo = Invoke-Fake @{ Args = @{ Target = 'a.test,b.test,c.test,d.test'; Server = '10.0.0.98,10.0.0.99' }; Os = $deadOs; Errors = @{ '10.0.0.98' = @{ A = 1460; AAAA = 1460 }; '10.0.0.99' = @{ A = 1460; AAAA = 1460 } } }
    Assert-Equal '聞いていないサーバーも 1 つの備考にまとめる' (Get-Cell $deadTwo 'd.test' '備考') 'DNS サーバー 10.0.0.98、10.0.0.99 は応答しないので聞いていない'
    $idErrors = @{
        '10.2.0.1' = @{ A = 'ERROR_TIMEOUT'; AAAA = 'ERROR_TIMEOUT' }
        '10.2.0.2' = @{ A = 'WSAETIMEDOUT'; AAAA = 'WSAETIMEDOUT' }
        '10.2.0.3' = @{ A = 'DNS_ERROR_RCODE_NAME_ERROR'; AAAA = 'DNS_ERROR_RCODE_NAME_ERROR' }
        '10.2.0.4' = @{ A = 'DNS_INFO_NO_RECORDS'; AAAA = 'DNS_INFO_NO_RECORDS' }
        '10.2.0.5' = @{ A = '10054'; AAAA = '10054' }
    }
    $ids = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = '10.2.0.1,10.2.0.2,10.2.0.3,10.2.0.4,10.2.0.5' }; Os = $deadOs; Errors = $idErrors }
    Assert-Equal 'エラーの ID でも見分ける' (Get-Cell $ids 'a.test' 'DNS の結果') '10.2.0.1: 応答なし / 10.2.0.2: 応答なし / 10.2.0.3: 名前が存在しない / 10.2.0.4: A / AAAA の登録がない / 10.2.0.5: 応答なし'

    $servfail = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = '10.0.0.78' }; Os = $deadOs; Errors = @{ '10.0.0.78' = @{ A = 9002; AAAA = 9002 } } }
    Assert-Equal 'A と AAAA が同じエラーなら 1 つにまとめる' ('{0}/{1}' -f (Get-Cell $servfail 'a.test' '備考'), (Get-Cell $servfail 'a.test' 'DNS の結果')) 'DNS サーバー 10.0.0.78: SERVFAIL/10.0.0.78: SERVFAIL'
    Assert-Equal 'エラーを返すサーバーは注意' ('{0}/{1}' -f (Get-Verdict $servfail 'a.test'), (Get-Cell $servfail 'a.test' '引けた元')) '注意/OS のみ'
    $refused = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = '10.0.0.78' }; Os = $deadOs; Errors = @{ '10.0.0.78' = @{ A = 9002; AAAA = 9005 } } }
    Assert-Equal 'A と AAAA が違うエラーなら分けて書く' (Get-Cell $refused 'a.test' '備考') 'DNS サーバー 10.0.0.78: A: SERVFAIL、AAAA: REFUSED'
    $partial = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = '10.0.0.78' }; Os = @{ 'a.test' = @('2001:db8::1') }; Errors = @{ '10.0.0.78' = @{ A = 9002 } }; Zone = (New-Zone @{ 'a.test' = @('2001:db8::1') }) }
    Assert-Equal '片方がエラーなら注意（答えは残す）' ('{0}/{1}/{2}' -f (Get-Verdict $partial 'a.test'), (Get-Cell $partial 'a.test' '備考'), (Get-Cell $partial 'a.test' '引けた元')) '注意/DNS サーバー 10.0.0.78: A: SERVFAIL/DNS'
    $otherError = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = '10.0.0.78' }; Os = $deadOs; Errors = @{ '10.0.0.78' = @{ A = 9999; AAAA = 9999 } } }
    Assert-Equal 'そのほかのエラーは、名前を除いて 1 行で書く' (Get-Cell $otherError 'a.test' '備考') 'DNS サーバー 10.0.0.78: 何かのエラー'
    $errorThree = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = '10.0.0.78,10.0.0.79,10.0.0.80' }; Os = $deadOs; Errors = @{ '10.0.0.78' = @{ A = 9002; AAAA = 9002 }; '10.0.0.79' = @{ A = 9005; AAAA = 9005 }; '10.0.0.80' = @{ A = 9002; AAAA = 9002 } } }
    Assert-Equal '同じエラーのサーバーをまとめる（サーバーの順）' (Get-Cell $errorThree 'a.test' '備考') 'DNS サーバー 10.0.0.78、10.0.0.80: SERVFAIL、DNS サーバー 10.0.0.79: REFUSED'
    $partialTwo = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = '10.0.0.78,10.0.0.79' }; Os = @{ 'a.test' = @('2001:db8::1') }; Errors = @{ '10.0.0.78' = @{ A = 9002 }; '10.0.0.79' = @{ A = 9002 } }; Zone = (New-Zone @{ 'a.test' = @('2001:db8::1') }) }
    Assert-Equal '片方だけのエラーもまとめる' (Get-Cell $partialTwo 'a.test' '備考') 'DNS サーバー 10.0.0.78、10.0.0.79: A: SERVFAIL'

    $ctrlError = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = '10.0.0.78' }; Os = $deadOs; Errors = @{ '10.0.0.78' = @{ A = 9998; AAAA = 9998 } } }
    Assert-Equal 'エラーの文の制御文字も置き換える' ('{0}/{1}' -f (Get-Cell $ctrlError 'a.test' 'DNS の結果'), (Get-Cell $ctrlError 'a.test' '備考')) '10.0.0.78: 制御?文字/DNS サーバー 10.0.0.78: 制御?文字'
    $dupServer = Invoke-Fake @{ Args = @{ Target = 'a.test' }; Servers = @('10.0.0.53', '10.0.0.53'); Os = $deadOs }
    Assert-Equal 'OS の設定で同じ DNS サーバーが重なっても 1 回だけ聞く' ('{0}/{1}' -f (Test-Contains $dupServer.Log 'DNS サーバー: 10.0.0.53（OS の設定）'), (Get-QueryCount $dupServer '10.0.0.53 A a.test *')) 'True/1'
    $noServer = Invoke-Fake @{ Args = @{ Target = 'a.test' }; Servers = @(); Os = $deadOs }
    Assert-Equal 'DNS サーバーがないと注意' ('{0}/{1}/{2}/{3}' -f (Get-Verdict $noServer 'a.test'), (Get-Cell $noServer 'a.test' '引けた元'), (Test-Note $noServer 'a.test' 'DNS サーバーに聞けなかった'), $noServer.Code) '注意/OS のみ/True/2'
    Assert-Equal 'DNS サーバーがないときの表示' ('{0}/{1}' -f (Test-Contains $noServer.Log 'DNS サーバー: なし'), (Test-Contains $noServer.Log '※ DNS サーバーが設定されていません')) 'True/True'
    $serverThrow = Invoke-Fake @{ Args = @{ Target = 'a.test' }; ServerThrow = $true; Os = $deadOs }
    Assert-Equal '設定を読めないときの表示（理由は 1 行で）' ('{0}/{1}' -f (Test-Contains $serverThrow.Log '※ DNS サーバーの設定を読めませんでした（サーバーの設定を読めません 2 行目）'), (Test-Note $serverThrow 'a.test' 'DNS サーバーに聞けなかった')) 'True/True'
    $ifThrow = Invoke-Fake @{ Args = @{ Target = 'a.test' }; InterfaceThrow = $true; Os = $deadOs }
    Assert-Equal 'インターフェイスを読めなくても、すべての設定を使う' (Test-Contains $ifThrow.Log 'DNS サーバー: 10.0.0.53, 10.8.8.8（OS の設定）') 'True'
    $given = Invoke-Fake @{ Args = @{ Target = 'a.test'; Server = "10.0.0.60, 10.0.0.60,'10.0.0.61',2001:DB8::53" }; Os = $deadOs }
    Assert-Equal '-Server の表示（重複なし、IPv6 はそろえる）' (Test-Contains $given.Log 'DNS サーバー: 10.0.0.60, 10.0.0.61, 2001:db8::53（指定）') 'True'
    Assert-Equal '-Server のときは OS の設定のサーバーに聞かない' (Get-QueryCount $given '10.0.0.53 *') 0

    Write-Section '短い名前'
    $short = Invoke-Fake @{ Args = @{ Target = 'web' }; SearchList = @('first.test', 'second.test'); ConnSuffix = 'conn.test'; Os = @{ 'web' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.second.test' = @('10.0.0.1'); 'web.conn.test' = @('10.0.0.9') }) }
    Assert-Equal '検索一覧を順に付ける（接続ごとのサフィックスは使わない）' (Get-Cell $short 'web' '解決した名前') 'web.second.test'
    Assert-Equal '答えがあれば次を聞かない' (($short.Queries | ForEach-Object { ($_ -split ' ')[2] } | Select-Object -Unique) -join ',') 'web.first.test,web.second.test'
    $searchOnly = Invoke-Fake @{ Args = @{ Target = 'web' }; SearchList = @('first.test'); ConnSuffix = 'conn.test'; Os = @{ 'web' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.conn.test' = @('10.0.0.1') }) }
    Assert-Equal '検索一覧があれば、接続ごとのサフィックスには聞かない' (Get-QueryCount $searchOnly '* web.conn.test *') 0
    $conn = Invoke-Fake @{ Args = @{ Target = 'web' }; ConnSuffix = 'conn.test'; Os = @{ 'web' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.conn.test' = @('10.0.0.1') }) }
    Assert-Equal '検索一覧がなければ接続ごとのサフィックス' (Get-Cell $conn 'web' '解決した名前') 'web.conn.test'
    $bare = Invoke-Fake @{ Args = @{ Target = 'web' }; SearchList = @('first.test'); Os = @{ 'web' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web' = @('10.0.0.1') }) }
    Assert-Equal 'どれでも答えがなければサフィックスなしも聞く' ('{0}/{1}' -f (Get-Cell $bare 'web' '解決した名前'), (Get-Cell $bare 'web' '引けた元')) 'web/DNS'
    $dotted = Invoke-Fake @{ Args = @{ Target = 'a.test' }; SearchList = @('first.test'); Os = @{ 'a.test' = @('10.0.0.1') } }
    Assert-Equal 'ドットのある名前にはサフィックスを付けない' (Get-QueryCount $dotted '* a.test.first.test *') 0
    $primaryTool = New-PatchedTool 'primary' '$primary = [Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName' '$primary = ''primary.test'''
    $primary = Invoke-Fake @{ Args = @{ Target = 'web' }; ConnSuffix = 'conn.test'; Os = @{ 'web' = @('10.0.0.1') }; Tool = $primaryTool }
    Assert-Equal 'プライマリ、接続ごと、サフィックスなしの順' (($primary.Queries | ForEach-Object { ($_ -split ' ')[2] } | Select-Object -Unique) -join ',') 'web.primary.test,web.conn.test,web'
    $suffixThrow = Invoke-Fake @{ Args = @{ Target = 'web' }; SuffixThrow = $true; Os = @{ 'web' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web' = @('10.0.0.1') }) }
    Assert-Equal 'サフィックスを読めなくても続ける' (Get-Verdict $suffixThrow 'web') 'OK'
    $shortDead = Invoke-Fake @{ Args = @{ Target = 'web'; Server = '10.0.0.99' }; SearchList = @('first.test', 'second.test'); Os = @{ 'web' = @('10.0.0.1') }; Errors = $deadErrors }
    Assert-Equal '応答しないサーバーには、次のサフィックスを聞かない' (Get-QueryCount $shortDead '10.0.0.99 * web.second.test *') 0

    Write-Section '期待する IP アドレス'
    $webZone = New-Zone @{ 'web.test' = @('10.0.2.10', '2001:db8::10') }
    $webOs = @{ 'web.test' = @('10.0.2.10', '2001:db8::10') }
    $match = Invoke-Fake @{ Args = @{ Target = 'web.test=10.0.2.10' }; Os = $webOs; Zone = $webZone }
    Assert-Equal '一致（種類が違う IP アドレスは比べない）' ('{0}/{1}/{2}' -f (Get-Verdict $match 'web.test'), (Get-Cell $match 'web.test' '期待する IP'), $match.Code) 'OK/10.0.2.10/0'
    $wrong = Invoke-Fake @{ Args = @{ Target = 'web.test=10.2.2.2' }; Os = $webOs; Zone = $webZone }
    Assert-Equal '違う（OS と DNS の両方を比べ、1 つの備考にまとめる）' ('{0}/{1}/{2}' -f (Get-Verdict $wrong 'web.test'), (Get-Cell $wrong 'web.test' '備考'), $wrong.Code) '注意/期待する IP アドレスと違う（OS の結果、DNS サーバー 10.0.0.53）/2'
    $wrongTwo = Invoke-Fake @{ Args = @{ Target = 'web.test=10.2.2.2' }; Servers = $two; Os = $webOs; Zone = $webZone }
    Assert-Equal '違う DNS サーバーもまとめる' (Get-Cell $wrongTwo 'web.test' '備考') '期待する IP アドレスと違う（OS の結果、DNS サーバー 10.0.0.53、10.0.0.54）'
    $extra = Invoke-Fake @{ Args = @{ Target = 'web.test=10.0.2.10' }; Os = @{ 'web.test' = @('10.0.2.10') }; Zone = (New-Zone @{ 'web.test' = @('10.0.2.10', '10.2.2.2') }) }
    Assert-Equal '期待していない同じ種類の IP アドレスがあれば違う' (Get-Cell $extra 'web.test' '備考') '期待する IP アドレスと違う（DNS サーバー 10.0.0.53）'
    $missing = Invoke-Fake @{ Args = @{ Target = 'web.test=10.0.2.10;10.0.2.11' }; Os = @{ 'web.test' = @('10.0.2.10') }; Zone = (New-Zone @{ 'web.test' = @('10.0.2.10', '10.0.2.11') }) }
    Assert-Equal '期待がすべて含まれていないと違う（DNS は OS の値を含むだけでは OK にしない）' (Get-Cell $missing 'web.test' '備考') '期待する IP アドレスと違う（OS の結果）'
    $twoIps = Invoke-Fake @{ Args = @{ Target = 'web.test=10.0.2.11;10.0.2.10' }; Os = @{ 'web.test' = @('10.0.2.10', '10.0.2.11') }; Zone = (New-Zone @{ 'web.test' = @('10.0.2.10', '10.0.2.11') }) }
    Assert-Equal '複数の期待（順番は比べない）' ('{0}/{1}' -f (Get-Verdict $twoIps 'web.test'), (Get-Cell $twoIps 'web.test' '期待する IP')) 'OK/10.0.2.10;10.0.2.11'
    $v6Expected = Invoke-Fake @{ Args = @{ Target = 'web.test=2001:DB8:0::10' }; Os = $webOs; Zone = $webZone }
    Assert-Equal 'IPv6 の期待は書き方をそろえて比べる' ('{0}/{1}' -f (Get-Verdict $v6Expected 'web.test'), (Get-Cell $v6Expected 'web.test' '期待する IP')) 'OK/2001:db8::10'
    $hostsOkDnsOld = Invoke-Fake @{ Args = @{ Target = 'web.test=10.0.2.10' }; Hosts = '10.0.2.10 web.test'; Os = @{ 'web.test' = @('10.0.2.10') }; Zone = (New-Zone @{ 'web.test' = @('10.2.2.2') }) }
    Assert-Equal 'hosts が期待どおりでも DNS が古ければ注意' ('{0}/{1}' -f (Test-Note $hostsOkDnsOld 'web.test' '期待する IP アドレスと違う（DNS サーバー'), (Test-Note $hostsOkDnsOld 'web.test' 'OS の結果')) 'True/False'
    $expectNoDns = Invoke-Fake @{ Args = @{ Target = 'web.test=10.0.2.10' }; Os = @{ 'web.test' = @('10.0.2.10') } }
    Assert-Equal 'DNS に名前がなければ期待と違う' (Test-Note $expectNoDns 'web.test' '期待する IP アドレスと違う（DNS サーバー') 'True'
    $expectDead = Invoke-Fake @{ Args = @{ Target = 'web.test=10.0.2.10'; Server = '10.0.0.99' }; Os = @{ 'web.test' = @('10.0.2.10') }; Errors = $deadErrors }
    Assert-Equal '応答しないサーバーは期待と比べない' (Test-Note $expectDead 'web.test' '期待する IP') 'False'
    $expectFail = Invoke-Fake @{ Args = @{ Target = 'web.test=10.0.2.10' } }
    Assert-Equal 'OS で引けなければ「引けない」' (Get-Verdict $expectFail 'web.test') '引けない'

    Write-Section '逆引き'
    $revOs = @{ 'web.test' = @('10.0.0.1') }
    $revZone = New-Zone @{ 'web.test' = @('10.0.0.1') }
    $revOk = Invoke-Fake @{ Args = @{ Target = 'web.test'; Reverse = $true }; Os = $revOs; Zone = $revZone; Ptr = @{ '10.0.0.1' = @('WEB.TEST') } }
    Assert-Equal '一致（大文字と小文字は区別しない）' ('{0}/{1}' -f (Get-Cell $revOk 'web.test' '逆引き'), (Get-Verdict $revOk 'web.test')) '10.0.0.1: 一致/OK'
    $revDot = Invoke-Fake @{ Args = @{ Target = 'web.test'; Reverse = $true }; Os = $revOs; Zone = $revZone; Ptr = @{ '10.0.0.1' = @('web.test.') } }
    Assert-Equal '最後のピリオドは比べない' (Get-Cell $revDot 'web.test' '逆引き') '10.0.0.1: 一致'
    $revBad = Invoke-Fake @{ Args = @{ Target = 'web.test'; Reverse = $true }; Os = $revOs; Zone = $revZone; Ptr = @{ '10.0.0.1' = @('other.test') } }
    Assert-Equal '違うと注意' ('{0}/{1}/{2}' -f (Get-Verdict $revBad 'web.test'), (Test-Note $revBad 'web.test' '逆引きの名前が違う（10.0.0.1 → other.test）'), (Get-Cell $revBad 'web.test' '逆引き')) '注意/True/10.0.0.1: 違う（other.test）'
    $revAny = Invoke-Fake @{ Args = @{ Target = 'web.test'; Reverse = $true }; Os = $revOs; Zone = $revZone; Ptr = @{ '10.0.0.1' = @('other.test', 'web.test') } }
    Assert-Equal '複数の名前のどれかが同じなら一致' (Get-Cell $revAny 'web.test' '逆引き') '10.0.0.1: 一致'
    $revCname = Invoke-Fake @{ Args = @{ Target = 'alias.test'; Reverse = $true }; Os = @{ 'alias.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'alias.test' = @('10.0.0.1') }); Cname = @{ 'alias.test' = @('end.test') }; Ptr = @{ '10.0.0.1' = @('end.test') } }
    Assert-Equal '解決した名前と同じなら一致' (Get-Cell $revCname 'alias.test' '逆引き') '10.0.0.1: 一致'
    $revShort = Invoke-Fake @{ Args = @{ Target = 'web'; Reverse = $true }; SearchList = @('corp.test'); Os = @{ 'web' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.corp.test' = @('10.0.0.1') }); Ptr = @{ '10.0.0.1' = @('web.corp.test') } }
    Assert-Equal '短い名前は解決した名前と比べる' (Get-Cell $revShort 'web' '逆引き') '10.0.0.1: 一致'
    $revNone = Invoke-Fake @{ Args = @{ Target = 'web.test'; Reverse = $true }; Os = $revOs; Zone = $revZone }
    Assert-Equal '登録がなければ判定を変えない' ('{0}/{1}/{2}' -f (Get-Verdict $revNone 'web.test'), (Test-Note $revNone 'web.test' '逆引きの登録がない（10.0.0.1）'), (Get-Cell $revNone 'web.test' '逆引き')) 'OK/True/10.0.0.1: なし'
    $revMulti = Invoke-Fake @{ Args = @{ Target = 'web.test'; Reverse = $true }; Os = @{ 'web.test' = @('10.0.0.1', '10.0.0.2') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1', '10.0.0.2') }); Ptr = @{ '10.0.0.1' = @('web.test') } }
    Assert-Equal 'IP アドレスごとに書く' (Get-Cell $revMulti 'web.test' '逆引き') '10.0.0.1: 一致 / 10.0.0.2: なし'
    $revManyIps = @('10.0.0.1', '10.0.0.2', '10.0.0.3', '10.0.0.4', '10.0.0.5', '10.0.0.6')
    $revMany = Invoke-Fake @{ Args = @{ Target = 'web.test'; Reverse = $true; Server = '10.0.0.66' }; Os = @{ 'web.test' = $revManyIps }; Zone = (New-Zone @{ 'web.test' = $revManyIps }); Errors = @{ '10.0.0.66' = @{ PTR = 1460; Except = @('10.0.0.3', '10.0.0.4', '10.0.0.5', '10.0.0.6') } }; Ptr = @{ '10.0.0.3' = @('other.test'); '10.0.0.4' = @('x.test', 'y.test') } }
    Assert-Equal '逆引きの備考は、同じ言葉の IP アドレスをまとめる' (Get-Cell $revMany 'web.test' '備考') '逆引きの名前が違う（10.0.0.3 → other.test、10.0.0.4 → x.test, y.test）、逆引きを聞けなかった（10.0.0.1、10.0.0.2）、逆引きの登録がない（10.0.0.5、10.0.0.6）'
    $revOff = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = $revOs; Zone = $revZone; Ptr = @{ '10.0.0.1' = @('other.test') } }
    Assert-Equal '-Reverse がなければ聞かない' ('{0}/{1}/{2}' -f (Get-Cell $revOff 'web.test' '逆引き'), (Get-QueryCount $revOff '* PTR *'), (Get-Verdict $revOff 'web.test')) '/0/OK'
    $revFail = Invoke-Fake @{ Args = @{ Target = 'none.test'; Reverse = $true } }
    Assert-Equal 'OS で引けなければ逆引きしない' (Get-QueryCount $revFail '* PTR *') 0
    $ipRun = Invoke-Fake @{ Args = @{ Target = '192.0.2.10' }; Ptr = @{ '192.0.2.10' = @('web01.test') } }
    Assert-Equal 'IP アドレスを書くと逆引きの名前を備考に書く' ('{0}/{1}/{2}/{3}' -f (Get-Cell $ipRun '192.0.2.10' '引けた元'), (Get-Cell $ipRun '192.0.2.10' '備考'), (Get-Cell $ipRun '192.0.2.10' '逆引き'), (Get-Cell $ipRun '192.0.2.10' 'OS の結果')) 'IP アドレス/逆引き: web01.test/192.0.2.10: web01.test/192.0.2.10'
    Assert-Equal 'IP アドレスは A と AAAA を聞かない' ('{0}/{1}' -f (Get-QueryCount $ipRun '* A *'), (Get-QueryCount $ipRun '* AAAA *')) '0/0'
    $ipV6 = Invoke-Fake @{ Args = @{ Target = '2001:DB8::10' }; Ptr = @{ '2001:db8::10' = @('v6.test') } }
    Assert-Equal 'IPv6 アドレスも逆引きする' (Test-Note $ipV6 '2001:DB8::10' '逆引き: v6.test') 'True'
    $ipNone = Invoke-Fake @{ Args = @{ Target = '192.0.2.10' } }
    Assert-Equal 'IP アドレスに登録がなければ判定を変えない' ('{0}/{1}' -f (Get-Verdict $ipNone '192.0.2.10'), (Test-Note $ipNone '192.0.2.10' '逆引きの登録がない')) 'OK/True'
    $revDead = Invoke-Fake @{ Args = @{ Target = '192.0.2.10'; Server = '10.0.0.99' }; Errors = $deadErrors }
    Assert-Equal '聞けなければ注意' ('{0}/{1}/{2}' -f (Get-Verdict $revDead '192.0.2.10'), (Test-Note $revDead '192.0.2.10' '逆引きを聞けなかった（192.0.2.10）'), (Get-Cell $revDead '192.0.2.10' '逆引き')) '注意/True/192.0.2.10: 聞けなかった'
    $revNext = Invoke-Fake @{ Args = @{ Target = '192.0.2.10'; Server = '10.0.0.99,10.0.0.53' }; Errors = $deadErrors; Ptr = @{ '192.0.2.10' = @('web01.test') } }
    Assert-Equal '応答しなければ次のサーバーに聞く' (Test-Note $revNext '192.0.2.10' '逆引き: web01.test') 'True'
    $revOther = Invoke-Fake @{ Args = @{ Target = '192.0.2.10'; Server = '10.0.0.78,10.0.0.53' }; Errors = @{ '10.0.0.78' = @{ PTR = 9002 } }; Ptr = @{ '192.0.2.10' = @('web01.test') } }
    Assert-Equal 'エラーを返せば次のサーバーに聞く' (Test-Note $revOther '192.0.2.10' '逆引き: web01.test') 'True'
    $ptrDead = Invoke-Fake @{ Args = @{ Target = '192.0.2.1,192.0.2.2,192.0.2.3,192.0.2.4,web.test'; Server = '10.0.0.66' }; Os = $revOs; Zone = $revZone; Errors = @{ '10.0.0.66' = @{ PTR = 1460 } } }
    Assert-Equal '逆引きだけ応答しないサーバーは 3 回で聞かなくなる' (Get-QueryCount $ptrDead '10.0.0.66 PTR *') 3
    Assert-Equal '逆引きだけ応答しないサーバーにも正引きは聞く' (Get-Cell $ptrDead 'web.test' '引けた元') 'DNS'
    $ptrAfterDead = Invoke-Fake @{ Args = @{ Target = 'a.test,b.test,c.test,192.0.2.1'; Server = '10.0.0.99' }; Os = $deadOs; Errors = $deadErrors }
    Assert-Equal '正引きで応答しないと判断したサーバーには逆引きも聞かない' (Get-QueryCount $ptrAfterDead '10.0.0.99 PTR *') 0

    Write-Section '-Hosts'
    $allHosts = "127.0.0.1 localhost`r`n::1 localhost`r`n127.0.0.1 loop.test`r`n::1 loop6.test`r`n10.3.3.3 forgot.test other.test`r`n127.0.0.1 both.test`r`n10.4.4.4 both.test`r`n"
    $allOs = @{ 'forgot.test' = @('10.3.3.3'); 'other.test' = @('10.3.3.3'); 'both.test' = @('10.4.4.4', '127.0.0.1') }
    $all = Invoke-Fake @{ Args = @{ Hosts = $true }; Hosts = $allHosts; Os = $allOs; Zone = (New-Zone @{ 'forgot.test' = @('10.9.9.9') }) }
    Assert-Equal 'hosts の名前をすべて調べる（localhost と 127.0.0.1・::1 だけの名前を除く）' (($all.Rows | ForEach-Object { $_.'名前' }) -join ',') 'forgot.test,other.test,both.test'
    Assert-Equal 'DNS と違うものは注意' (Test-Note $all 'forgot.test' 'hosts と DNS が違う') 'True'
    Assert-Equal 'DNS にないものは OK で備考' ('{0}/{1}' -f (Get-Verdict $all 'other.test'), (Get-Cell $all 'other.test' '備考')) 'OK/DNS にはない'
    $upper = Invoke-Fake @{ Args = @{ Hosts = $true; NoSave = $true }; Hosts = "10.0.0.1 LOCALHOST`r`n" }
    Assert-Equal 'localhost は大文字でも対象にしない' ('{0}/{1}' -f $upper.Code, (Test-Contains $upper.Log '調べる対象がありません')) '1/True'
    $noHostsAll = Invoke-Fake @{ Args = @{ Hosts = $true }; Hosts = $null }
    Assert-Equal '-Hosts で hosts を読めないと止まる' ('{0}/{1}/{2}' -f $noHostsAll.Code, (Test-Contains $noHostsAll.Log 'hosts を読めません: '), (Test-Path -LiteralPath $noHostsAll.Out)) '1/True/False'
    $noName = Invoke-Fake @{ Args = @{ Hosts = $true }; Hosts = '# 空' }
    Assert-Equal 'hosts に名前がないと止まる' ('{0}/{1}' -f $noName.Code, (Test-Contains $noName.Log '調べる対象がありません')) '1/True'
    Write-Utf8 "$workRoot\hl.csv" "名前`r`nlist.test`r`n"
    $together = Invoke-Fake @{ Args = @{ List = "$workRoot\hl.csv"; Target = 'target.test,forgot.test'; Hosts = $true }; Hosts = '10.3.3.3 forgot.test hosts.test'; Os = @{ 'list.test' = @('10.0.0.1'); 'target.test' = @('10.0.0.2'); 'forgot.test' = @('10.3.3.3'); 'hosts.test' = @('10.3.3.3') } }
    Assert-Equal '-List、-Target、-Hosts を一緒に（同じ名前は 1 回）' (($together.Rows | ForEach-Object { $_.'名前' }) -join ',') 'list.test,target.test,forgot.test,hosts.test'
    $ctrlHosts = Invoke-Fake @{ Args = @{ Hosts = $true }; Hosts = "10.3.3.3 bad$([char]27)name.test`r`n"; Os = @{ "bad$([char]27)name.test" = @('10.3.3.3') } }
    Assert-Equal 'hosts の名前の制御文字も置き換える' ('{0}/{1}' -f (Get-Verdict $ctrlHosts 'bad?name.test'), $ctrlHosts.Log.Contains([string][char]27)) 'OK/False'

    Write-Section '対象の CSV'
    $newList = "$workRoot\names.csv"
    $made = Invoke-Fake @{ Args = @{ NewList = $newList; NoSave = $true } }
    Assert-Equal 'ひな形を作る' ('{0}/{1}' -f $made.Code, (Test-Path -LiteralPath $newList)) '0/True'
    Assert-Equal 'ひな形の案内' ('{0}/{1}' -f (Test-Contains $made.Log "ひな形を作りました: $newList"), (Test-Contains $made.Log '-List で指定して')) 'True/True'
    Assert-Equal 'ひな形の BOM' ([IO.File]::ReadAllBytes($newList)[0..2] -join ',') '239,187,191'
    Assert-Equal 'ひな形の見出し' ([IO.File]::ReadAllLines($newList)[0]) '名前,期待する IP,メモ'
    $again = Invoke-Fake @{ Args = @{ NewList = $newList; NoSave = $true } }
    Assert-Equal 'ひな形を上書きしない' ('{0}/{1}' -f $again.Code, (Test-Contains $again.Log "同じ名前のファイルがあります: $newList（上書きしません）")) '1/True'
    $sameFolder = Invoke-Fake @{ Args = @{ NewList = $workRoot; NoSave = $true } }
    Assert-Equal 'フォルダと同じ名前でも作らない' ('{0}/{1}' -f $sameFolder.Code, (Test-Contains $sameFolder.Log '同じ名前のファイルがあります')) '1/True'
    $noParent = Invoke-Fake @{ Args = @{ NewList = "$workRoot\no_such_dir\a.csv"; NoSave = $true } }
    Assert-Equal 'ひな形を作れない' ('{0}/{1}' -f $noParent.Code, (Test-Contains $noParent.Log 'ひな形を作れません: ')) '1/True'
    $fromTemplate = Invoke-Fake @{ Args = @{ List = $newList } }
    Assert-Equal 'ひな形をそのまま読める（# の行を飛ばす）' (($fromTemplate.Rows | ForEach-Object { $_.'名前' }) -join ',') 'web01.example.local,db01.example.local'
    Assert-Equal 'ひな形の期待する IP とメモ' ('{0}/{1}' -f (Get-Cell $fromTemplate 'web01.example.local' '期待する IP'), (Get-Cell $fromTemplate 'web01.example.local' 'メモ')) '10.0.2.10/切り替え後の Web サーバー'

    $listOs = @{ 'web.test' = @('10.0.2.10'); 'db.test' = @('10.0.2.20') }
    Write-Utf8 "$workRoot\list.csv" "名前,期待するIP,メモ`r`n# コメント,,`r`nweb.test,10.0.2.10,3月1日の切り替え`r`n,,`r`n`r`ndb.test,,`r`n"
    $listRun = Invoke-Fake @{ Args = @{ List = "$workRoot\list.csv" }; Os = $listOs }
    Assert-Equal '見出しの空白が違っても期待する IP を読む' (Get-Cell $listRun 'web.test' '期待する IP') '10.0.2.10'
    Assert-Equal 'メモはいつも ="..."' (Get-Row $listRun 'web.test').'メモ' '="3月1日の切り替え"'
    Assert-Equal '空の行と # の行を飛ばす' (@($listRun.Rows).Count) 2
    Assert-Equal 'メモが空なら空欄' (Get-Row $listRun 'db.test').'メモ' ''
    Assert-Equal '-Target の期待する IP を優先する' (Get-Cell (Invoke-Fake @{ Args = @{ List = "$workRoot\list.csv"; Target = 'web.test=10.7.7.7' }; Os = $listOs }) 'web.test' '期待する IP') '10.7.7.7'
    $listKeep = Invoke-Fake @{ Args = @{ List = "$workRoot\list.csv"; Target = 'WEB.TEST' }; Os = $listOs }
    Assert-Equal '-Target に期待がなければ CSV の期待とメモを残す' ('{0}/{1}' -f (Get-Cell $listKeep 'web.test' '期待する IP'), (Get-Cell $listKeep 'web.test' 'メモ')) '10.0.2.10/3月1日の切り替え'
    Write-Utf8 "$workRoot\memo.csv" "名前,メモ`r`nweb.test,`r`nweb.test,あとのメモ`r`nweb.test,さらにあとのメモ`r`n"
    Assert-Equal '同じ名前の行は、足りないメモを補う（最初のメモを残す）' (Get-Cell (Invoke-Fake @{ Args = @{ List = "$workRoot\memo.csv" }; Os = $listOs }) 'web.test' 'メモ') 'あとのメモ'
    Write-Utf8 "$workRoot\en.csv" "MEMO,EXPECTED,NAME,Other`r`nメモ1,10.0.2.10,web.test,x`r`n"
    $english = Invoke-Fake @{ Args = @{ List = "$workRoot\en.csv" }; Os = $listOs }
    Assert-Equal '英語の見出し・列の順番が違う・知らない列' ('{0}/{1}' -f (Get-Cell $english 'web.test' '期待する IP'), (Get-Cell $english 'web.test' 'メモ')) '10.0.2.10/メモ1'
    Write-Utf8 "$workRoot\quoted.csv" "名前,メモ`r`n""web.test"",""カンマ, と """"引用符"""" と`r`n改行""`r`n"
    $quoted = Invoke-Fake @{ Args = @{ List = "$workRoot\quoted.csv" }; Os = $listOs }
    Assert-Equal '引用符とカンマとセル内の改行' (Get-Cell $quoted 'web.test' 'メモ') 'カンマ, と ""引用符"" と 改行'
    Write-Utf8 "$workRoot\short.csv" "名前,期待する IP,メモ`r`nweb.test`r`n"
    Assert-Equal '行の列が足りなくても読む' (Get-Verdict (Invoke-Fake @{ Args = @{ List = "$workRoot\short.csv" }; Os = $listOs }) 'web.test') 'OK'
    Write-Utf8 "$workRoot\lf.csv" "名前`nweb.test`ndb.test`n"
    Assert-Equal '改行が LF だけでも読む' (@((Invoke-Fake @{ Args = @{ List = "$workRoot\lf.csv" }; Os = $listOs }).Rows).Count) 2
    $encodings = @(
        @('Shift_JIS', [byte[]][Text.Encoding]::GetEncoding(932).GetBytes("名前,メモ`r`nweb.test,日本語のメモ`r`n")),
        @('UTF-8 BOM', [byte[]]($utf8Bom.GetPreamble() + $utf8NoBom.GetBytes("名前,メモ`r`nweb.test,日本語のメモ`r`n"))),
        @('UTF-16', [byte[]]([Text.Encoding]::Unicode.GetPreamble() + [Text.Encoding]::Unicode.GetBytes("名前,メモ`r`nweb.test,日本語のメモ`r`n")))
    )
    foreach ($case in $encodings) {
        [IO.File]::WriteAllBytes("$workRoot\enc.csv", $case[1])
        Assert-Equal "CSV の文字コード: $($case[0])" (Get-Cell (Invoke-Fake @{ Args = @{ List = "$workRoot\enc.csv" }; Os = $listOs }) 'web.test' 'メモ') '日本語のメモ'
    }

    $csvStops = @(
        @("名前,期待の IP`r`nweb.test,1.2.3.4`r`n", '対象の CSV の見出しを読めません: 期待の IP'),
        @("Name,ExpectIP`r`nweb.test,1.2.3.4`r`n", '対象の CSV の見出しを読めません: ExpectIP'),
        @("ホスト`r`nweb.test`r`n", '対象の CSV に「名前」の列がありません'),
        @("名前,期待する IP`r`nweb.test,abc`r`n", '対象の CSV の 2 行目: 期待する IP アドレスが正しくありません: abc'),
        @("名前,期待する IP`r`nweb.test,１０.０.０.１`r`n", '対象の CSV の 2 行目: 期待する IP アドレスが正しくありません'),
        @("名前,期待する IP`r`nweb.test,192.168.1`r`n", '対象の CSV の 2 行目: 期待する IP アドレスが正しくありません: 192.168.1'),
        @("名前,期待する IP`r`nweb.test,256.1.1.1`r`n", '対象の CSV の 2 行目: 期待する IP アドレスが正しくありません: 256.1.1.1'),
        @("名前,期待する IP`r`na.test,10.0.0.1`r`nA.TEST,10.0.0.2`r`n", '対象の CSV の 3 行目: 同じ名前に別の期待する IP アドレスが書かれています: A.TEST'),
        @("名前,期待する IP,メモ`r`na.test,,""x`r`ny""`r`nb.test,abc,`r`n", '対象の CSV の 4 行目: 期待する IP アドレスが正しくありません: abc'),
        @("名前`r`n""abc`r`n", '対象の CSV の 2 行目: 引用符（"）が閉じていません')
    )
    foreach ($case in $csvStops) {
        Write-Utf8 "$workRoot\stop.csv" $case[0]
        $result = Invoke-Fake @{ Args = @{ List = "$workRoot\stop.csv" } }
        Assert-Equal "CSV の誤り: $($case[1])" ('{0}/{1}/{2}' -f $result.Code, (Test-Contains $result.Log $case[1]), (Test-Path -LiteralPath $result.Out)) '1/True/False'
    }
    Write-Utf8 "$workRoot\order.csv" "名前,期待する IP`r`nweb.test,10.0.0.1;10.0.0.2`r`nweb.test,10.0.0.2;10.0.0.1`r`n"
    Assert-Equal '順番だけ違う期待する IP は止めない' (@((Invoke-Fake @{ Args = @{ List = "$workRoot\order.csv" }; Os = $listOs }).Rows).Count) 1
    Write-Utf8 "$workRoot\empty.csv" "名前,期待する IP`r`n"
    $emptyCsv = Invoke-Fake @{ Args = @{ List = "$workRoot\empty.csv" } }
    Assert-Equal '見出しだけの CSV' ('{0}/{1}' -f $emptyCsv.Code, (Test-Contains $emptyCsv.Log '調べる対象がありません')) '1/True'
    Write-Utf8 "$workRoot\hint.csv.txt" "名前`r`nweb.test`r`n"
    $hint = Invoke-Fake @{ Args = @{ List = "$workRoot\hint.csv" } }
    Assert-Equal '.txt のヒント' ('{0}/{1}/{2}' -f $hint.Code, (Test-Contains $hint.Log "対象の CSV が見つかりません: $workRoot\hint.csv"), (Test-Contains $hint.Log "拡張子が .txt になっていませんか: $workRoot\hint.csv.txt")) '1/True/True'
    Write-Utf8 "$workRoot\日本語 と スペース\対象 [1].csv" "名前`r`nweb.test`r`n"
    Assert-Equal 'CSV のパスに日本語・スペース・[ ]' (Get-Verdict (Invoke-Fake @{ Args = @{ List = "$workRoot\日本語 と スペース\対象 [1].csv" }; Os = $listOs }) 'web.test') 'OK'

    Write-Section 'CSV の Excel 対策'
    $excelOs = @{ '001' = @('10.0.0.1'); '2026-11-05' = @('10.0.0.1'); '1e5' = @('10.0.0.1'); 'web.test' = @('10.0.0.1') }
    $excel = Invoke-Fake @{ Args = @{ Target = '001,2026-11-05,1e5,web.test' }; Os = $excelOs }
    Assert-Equal '数字だけの名前' (@($excel.Rows)[0].'名前') '="001"'
    Assert-Equal '日付に見える名前' (@($excel.Rows)[1].'名前') '="2026-11-05"'
    Assert-Equal '指数に見える名前' (@($excel.Rows)[2].'名前') '="1e5"'
    Assert-Equal '普通の名前はそのまま' (@($excel.Rows)[3].'名前') 'web.test'
    Write-Utf8 "$workRoot\excel.csv" "名前,メモ`r`nweb.test,""a""""b""`r`n"
    Assert-Equal 'メモの引用符は式の中で 2 つ重ねる' (Get-Row (Invoke-Fake @{ Args = @{ List = "$workRoot\excel.csv" }; Os = $excelOs }) 'web.test').'メモ' '="a""b"'
    $longMemo = 'あ' * 256
    Write-Utf8 "$workRoot\excel2.csv" "名前,メモ`r`nweb.test,$longMemo`r`n"
    Assert-Equal '式に書けない長いメモは先頭に空白' (Get-Row (Invoke-Fake @{ Args = @{ List = "$workRoot\excel2.csv" }; Os = $excelOs }) 'web.test').'メモ' (' ' + $longMemo)


    $eqPtr = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1') }); Cname = @{ 'web.test' = @('=cmd|calc.test') } }
    Assert-Equal 'DNS が答えた = で始まる名前は数式にしない' (Get-Row $eqPtr 'web.test').'解決した名前' '="=cmd|calc.test"'

    Write-Section '制御文字（DNS サーバーが決める値）'
    $esc = [string][char]27
    $evil = Invoke-Fake @{ Args = @{ Target = 'evil.test,192.0.2.1,192.0.2.2'; Reverse = $true }; Os = @{ 'evil.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'evil.test' = @('10.0.0.1') }); Cname = @{ 'evil.test' = @("x${esc}[31m.test", "y`nz.test") }; Ptr = @{ '10.0.0.1' = @("a${esc}[31m.test"); '192.0.2.1' = @("b`r`nc.test"); '192.0.2.2' = @("d$([char]0)e.test") } }
    Assert-Equal 'CNAME の制御文字' (Get-Cell $evil 'evil.test' 'CNAME') 'evil.test → x?[31m.test → y?z.test'
    Assert-Equal '解決した名前の制御文字' (Get-Cell $evil 'evil.test' '解決した名前') 'y?z.test'
    Assert-Equal '逆引きの制御文字（違う名前）' (Get-Cell $evil 'evil.test' '逆引き') '10.0.0.1: 違う（a?[31m.test）'
    Assert-Equal '備考の制御文字' (Test-Note $evil 'evil.test' '逆引きの名前が違う（10.0.0.1 → a?[31m.test）') 'True'
    Assert-Equal '改行' (Get-Cell $evil '192.0.2.1' '備考') '逆引き: b??c.test'
    Assert-Equal 'NUL' (Get-Cell $evil '192.0.2.2' '備考') '逆引き: d?e.test'
    Assert-Equal '画面にエスケープ文字を出さない' $evil.Log.Contains($esc) 'False'
    $evilCsv = [IO.File]::ReadAllText((Join-Path $evil.Dir 'results.csv'))
    $evilSummary = [IO.File]::ReadAllText((Join-Path $evil.Dir 'summary.txt'))
    Assert-Equal 'CSV にエスケープ文字を書かない' $evilCsv.Contains($esc) 'False'
    Assert-Equal 'summary.txt にエスケープ文字を書かない' $evilSummary.Contains($esc) 'False'
    Assert-Equal 'CSV の行の数（名前の改行で行が増えない）' ($evilCsv.TrimEnd() -split '\r?\n').Count 4
    Assert-Equal '比べるときは元の名前を使う' (Get-Cell (Invoke-Fake @{ Args = @{ Target = "a${esc}b.test"; Reverse = $true }; Os = @{ "a${esc}b.test" = @('10.0.0.1') }; Ptr = @{ '10.0.0.1' = @("a${esc}b.test") } }) 'a?b.test' '逆引き') '10.0.0.1: 一致'

    Write-Section '設定ファイル'
    $confOs = @{ 'web.test' = @('10.0.2.10'); 'db.test' = @('10.0.2.20') }
    Write-Utf8 "$workRoot\conf.ini" "# コメント`r`n; これもコメント`r`n`r`nTarget = web.test=10.0.2.10`r`nTarget = 192.0.2.10`r`nServer = 10.0.0.53`r`nServer = 10.0.0.54`r`nReverse = false`r`nNoSave = false`r`nHosts = false`r`n"
    $confRun = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\conf.ini" }; Os = $confOs }
    Assert-Equal '複数の行の Target' (@($confRun.Rows).Count) 2
    Assert-Equal 'Target の = の後ろを期待する IP として読む' (Get-Cell $confRun 'web.test' '期待する IP') '10.0.2.10'
    Assert-Equal '複数の行の Server' (Test-Contains $confRun.Log '10.0.0.53, 10.0.0.54（指定）') 'True'
    Assert-Equal '設定ファイルの表示' (Test-Contains $confRun.Log "設定ファイル: $workRoot\conf.ini") 'True'
    Assert-Equal 'false は無効' (Get-Cell $confRun 'web.test' '逆引き') ''
    $confOver = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\conf.ini"; Target = 'db.test'; Server = '10.0.0.60' }; Os = $confOs }
    Assert-Equal 'コマンドラインの -Target を優先する（合わせない）' (($confOver.Rows | ForEach-Object { $_.'名前' }) -join ',') 'db.test'
    Assert-Equal 'コマンドラインの -Server を優先する（合わせない）' ('{0}/{1}' -f (Test-Contains $confOver.Log '10.0.0.60（指定）'), (Get-QueryCount $confOver '10.0.0.53 *')) 'True/0'
    Write-Utf8 "$workRoot\conf2\sub.ini" "Target = 192.0.2.10`r`nList = names.csv`r`nOutputFolder = 記録`r`nHosts = true`r`nReverse = true`r`n"
    Write-Utf8 "$workRoot\conf2\names.csv" "名前`r`nweb.test`r`n"
    $confRel = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\conf2\sub.ini" }; NoOut = $true; Hosts = '10.3.3.3 forgot.test'; Os = @{ 'web.test' = @('10.0.2.10'); 'forgot.test' = @('10.3.3.3') }; Zone = (New-Zone @{ 'web.test' = @('10.0.2.10') }); Ptr = @{ '10.0.2.10' = @('web.test') } }
    Assert-Equal '相対パスは設定ファイルのあるフォルダが基準（List）' (Test-Contains $confRel.Log '対象 3 件') 'True'
    Assert-Equal '相対パスは設定ファイルのあるフォルダが基準（OutputFolder）' (@(Get-ChildItem -LiteralPath "$workRoot\conf2\記録" -Directory -Filter 'NameResolution_*').Count) 1
    $confRelRows = @(Import-Csv -LiteralPath (Join-Path @(Get-ChildItem -LiteralPath "$workRoot\conf2\記録" -Directory)[0].FullName 'results.csv') -Encoding UTF8)
    Assert-Equal 'Hosts = true' (@($confRelRows | Where-Object { $_.'名前' -eq 'forgot.test' }).Count) 1
    Assert-Equal 'Reverse = true' (@($confRelRows | Where-Object { $_.'名前' -eq 'web.test' })[0].'逆引き') '10.0.2.10: 一致'
    $env:WB_TNR_OUT = "$workRoot\envout"
    Write-Utf8 "$workRoot\conf3.ini" "Target = web.test`r`nNoSave = true`r`nOutputFolder = %WB_TNR_OUT%\nosave`r`n"
    $confSave = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\conf3.ini" }; NoOut = $true; Os = $confOs }
    Assert-Equal 'NoSave = true' ('{0}/{1}' -f (Test-Contains $confSave.Log '保存先:'), (Test-Path -LiteralPath "$workRoot\envout\nosave")) 'False/False'

    Write-Utf8 "$workRoot\conf5.ini" "Target = web.test`r`nOutputFolder = %WB_TNR_OUT%\saved`r`n"
    $confEnv = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\conf5.ini" }; NoOut = $true; Os = $confOs }

    Assert-Equal '環境変数を展開する' (@(Get-ChildItem -LiteralPath "$workRoot\envout\saved" -Directory -Filter 'NameResolution_*' -ErrorAction SilentlyContinue).Count) 1

    Write-Utf8 "$workRoot\conf6.ini" "List = no_such.csv`r`n"
    $confNote = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\conf6.ini" }; NoOut = $true }
    Assert-Equal '見つからない値に行番号を添える' ('{0}/{1}' -f $confNote.Code, (Test-Contains $confNote.Log '（設定ファイルの 1 行目）')) '1/True'
    Write-Utf8 "$workRoot\conf7.ini" "Target = web.test`r`nOutputFolder = $workRoot\conf.ini`r`n"
    $confWarn = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\conf7.ini" }; NoOut = $true; Os = $confOs }
    Assert-Equal '書き込めない保存先に行番号を添える' ('{0}/{1}' -f (Test-Contains $confWarn.Log '保存先に書き込めないため'), (Test-Contains $confWarn.Log '（設定ファイルの 2 行目）')) 'True/True'
    [IO.File]::WriteAllBytes("$workRoot\sjis.ini", [Text.Encoding]::GetEncoding(932).GetBytes("# 日本語のコメント`r`nTarget = web.test`r`n"))
    Assert-Equal 'Shift_JIS の設定ファイル' (Get-Verdict (Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\sjis.ini" }; Os = $confOs }) 'web.test') 'OK'
    Write-Utf8 "$workRoot\quote.ini" "Target = ""web.test"",""db.test""`r`n"
    Assert-Equal '1 つずつ引用符で囲む' (@((Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\quote.ini" }; Os = $confOs }).Rows).Count) 2
    $confStops = @(
        @('Exlude = a', '設定ファイルの 1 行目: 知らない項目です: Exlude'),
        @('NewList = a.csv', 'NewList は設定ファイルには書けません'),
        @('Server = dns01', 'Server は IP アドレスで指定してください: dns01'),
        @('Target = a=abc', 'Target の期待する IP アドレスが正しくありません: abc'),
        @('Target = a=', 'Target の期待する IP アドレスが空です'),
        @('NoSave = yes', 'NoSave は true か false で指定してください: yes'),
        @('Hosts = 1', 'Hosts は true か false で指定してください: 1'),
        @("# c`r`nList = a.csv`r`nList = b.csv", 'List が 2 回書かれています（2 行目と 3 行目）'),
        @('Target = %NO_SUCH_VARIABLE_X%', '定義されていない環境変数です'),
        @('List = %NO_SUCH_VARIABLE_X%\a.csv', '定義されていない環境変数です'),
        @('これは値がない行', '「項目 = 値」の形になっていません'),
        @('List = C:\a"b.csv', 'List の値に引用符（"）が残っています')
    )
    foreach ($case in $confStops) {
        Write-Utf8 "$workRoot\bad.ini" $case[0]
        $result = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\bad.ini" }; NoOut = $true }
        Assert-Equal "書き間違い: $($case[1])" ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $case[1])) '1/True'
        Assert-Equal '書き間違い: パスを添える' (Test-Contains $result.Log "ファイル: $workRoot\bad.ini") 'True'
    }
    $confMissing = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\no_such.ini" } }
    Assert-Equal '設定ファイルが見つからない' ('{0}/{1}' -f $confMissing.Code, (Test-Contains $confMissing.Log "設定ファイルが見つかりません: $workRoot\no_such.ini")) '1/True'
    Write-Utf8 "$workRoot\txt.ini.txt" 'Target = a'
    Assert-Equal '設定ファイルの .txt のヒント' (Test-Contains (Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\txt.ini" } }).Log '拡張子が .txt になっていませんか') 'True'
    Write-Utf8 "$workRoot\emptyvalue.ini" "Target = `r`nList = `r`nTarget = web.test`r`n"
    Assert-Equal '値が空の行は書かなかったのと同じ' (@((Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\emptyvalue.ini" }; Os = $confOs }).Rows).Count) 1

    Write-Section '止まるとき'
    $stops = @(
        @('対象なし', @{ NoSave = $true }, '調べる名前を指定してください'),
        @('-Target が空', @{ Target = ' , ' }, '調べる対象がありません'),
        @('-Server の誤り', @{ Target = 'a'; Server = 'dns01' }, '-Server は IP アドレスで指定してください: dns01'),
        @('-Server の数字が足りない', @{ Target = 'a'; Server = '1.2.3' }, '-Server は IP アドレスで指定してください: 1.2.3'),
        @('期待する IP の誤り', @{ Target = 'a=1.2.3' }, '-Target の期待する IP アドレスが正しくありません: 1.2.3'),
        @('= の後ろが空', @{ Target = 'a=' }, '-Target の期待する IP アドレスが空です'),
        @('= の前が空', @{ Target = '=10.0.0.1' }, '-Target の名前がありません'),
        @('-Target の中の食い違い', @{ Target = 'a=10.0.0.1,A=10.0.0.2' }, '-Target で、同じ名前に別の期待する IP アドレス'),
        @('見つからない CSV', @{ List = "$workRoot\no_such.csv" }, '対象の CSV が見つかりません'),
        @('-List の引用符', @{ List = 'C:\data" -NoSave' }, '-List の値に引用符（"）が残っています'),
        @('-OutputFolder の引用符', @{ Target = 'a'; OutputFolder = 'C:\data" -NoSave' }, '-OutputFolder の値に引用符（"）が残っています'),
        @('-ConfigFile の引用符', @{ ConfigFile = 'C:\data" -NoSave' }, '-ConfigFile の値に引用符（"）が残っています'),
        @('-NewList の引用符', @{ NewList = 'C:\data" -NoSave' }, '-NewList の値に引用符（"）が残っています'),
        @('-Target の引用符', @{ Target = 'a"b' }, '-Target の値に引用符（"）が残っています'),
        @('-Server の引用符', @{ Target = 'a'; Server = '10.0.0.1"x' }, '-Server の値に引用符（"）が残っています')
    )
    foreach ($stop in $stops) {
        $result = Invoke-Fake @{ Args = $stop[1]; NoOut = $true }
        Assert-Equal $stop[0] ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $stop[2])) '1/True'
    }
    $sameTarget = Invoke-Fake @{ Args = @{ Target = 'web.test=10.0.2.10,web.test=10.0.2.10' }; Os = @{ 'web.test' = @('10.0.2.10') } }
    Assert-Equal '同じ期待なら止めない' (@($sameTarget.Rows).Count) 1
    $caseTarget = Invoke-Fake @{ Args = @{ Target = 'WEB.test,web.TEST' }; Os = @{ 'web.test' = @('10.0.2.10') } }
    Assert-Equal '同じ名前は大文字と小文字によらず 1 回' (($caseTarget.Rows | ForEach-Object { $_.'名前' }) -join ',') 'WEB.test'
    $leadingZero = Invoke-Fake @{ Args = @{ Target = 'web.test=010.000.002.010' }; Os = @{ 'web.test' = @('10.0.2.10') } }
    Assert-Equal '0 で始まる IP アドレスは 10 進数で読む（8 進数にしない）' (Get-Cell $leadingZero 'web.test' '期待する IP') '10.0.2.10'
    $noCmdlet = New-PatchedTool 'nocmdlet' 'Get-Command Resolve-DnsName -ErrorAction SilentlyContinue' 'Get-Command No-Such-ResolveDnsName -ErrorAction SilentlyContinue'
    $result = Invoke-Fake @{ Args = @{ Target = 'a' }; Tool = $noCmdlet }
    Assert-Equal 'Resolve-DnsName を使えない' ('{0}/{1}/{2}' -f $result.Code, (Test-Contains $result.Log 'Resolve-DnsName を使えないため、調べられません'), (Test-Path -LiteralPath $result.Out)) '1/True/False'

    Write-Section '保存と表示'
    $hundred = New-Object 'System.Collections.Generic.List[string]'
    $bulkOs = @{}
    for ($i = 1; $i -le 205; $i++) { $n = 'n{0:000}.test' -f $i; $hundred.Add($n); $bulkOs[$n] = @('10.0.0.1') }
    $bulk = Invoke-Fake @{ Args = @{ Target = ($hundred -join ',') }; Os = $bulkOs; Zone = (New-Zone @{ 'n205.test' = @('10.0.0.1') }) }
    Assert-Equal '200 件まで表示する' ('{0}/{1}' -f (Test-Contains $bulk.Log 'n200.test'), (Test-Contains $bulk.Log 'n201.test')) 'True/False'
    Assert-Equal '超えた分の案内' (Test-Contains $bulk.Log '（ほか 5 件は results.csv を参照してください）') 'True'
    Assert-Equal '件数の表示' (Test-Contains $bulk.Log '対象 205 件') 'True'
    Assert-Equal 'summary.txt は省略しない' ([IO.File]::ReadAllText((Join-Path $bulk.Dir 'summary.txt')).Contains('n205.test')) 'True'
    Assert-Equal 'results.csv はすべて' (@($bulk.Rows).Count) 205

    $bulkNoSave = Invoke-Fake @{ Args = @{ Target = ($hundred -join ','); NoSave = $true }; Os = $bulkOs }
    Assert-Equal '-NoSave は案内を出さない' (Test-Contains $bulkNoSave.Log '（ほか') 'False'
    Assert-Equal '-NoSave はすべて表示する' (Test-Contains $bulkNoSave.Log 'n205.test') 'True'
    Assert-Equal '-NoSave はフォルダを作らない' ('{0}/{1}' -f (Test-Path -LiteralPath $bulkNoSave.Out), (Test-Contains $bulkNoSave.Log '保存先:')) 'False/False'

    Write-Utf8 "$workRoot\notdir" 'x'
    $unwritable = Invoke-Fake @{ Args = @{ Target = 'web.test'; OutputFolder = "$workRoot\notdir" }; Os = @{ 'web.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1') }) }
    Assert-Equal '書き込めない保存先は警告して画面に表示' ('{0}/{1}/{2}' -f $unwritable.Code, (Test-Contains $unwritable.Log '保存先に書き込めないため、結果は保存しません'), (Test-Contains $unwritable.Log ' OK        web.test')) '0/True/True'
    $unwritableBulk = Invoke-Fake @{ Args = @{ Target = ($hundred -join ','); OutputFolder = "$workRoot\notdir" }; Os = $bulkOs }
    Assert-Equal '保存できないときはすべて表示する' (Test-Contains $unwritableBulk.Log 'n205.test') 'True'
    $nested = "$workRoot\deep\a\b"
    $deep = Invoke-Fake @{ Args = @{ Target = 'web.test'; OutputFolder = $nested }; Os = @{ 'web.test' = @('10.0.0.1') } }
    Assert-Equal '保存先のフォルダがなければ作る' (@(Get-ChildItem -LiteralPath $nested -Directory -Filter 'NameResolution_*').Count) 1
    $spaceOut = "$workRoot\日本語 と スペース [x]"
    $spaced = Invoke-Fake @{ Args = @{ Target = 'web.test'; OutputFolder = $spaceOut }; Os = @{ 'web.test' = @('10.0.0.1') } }
    Assert-Equal '保存先に日本語・スペース・[ ]' ('{0}/{1}' -f $spaced.Code, (@([IO.Directory]::GetDirectories($spaceOut)).Count)) '0/1'

    $saveFail = New-PatchedTool 'savefail' "Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'results.csv'))" "throw 'disk full'; Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'results.csv'))"
    $failRun = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1') }); Tool = $saveFail }
    Assert-Equal '保存の失敗は警告して、書きかけを残さない' ('{0}/{1}/{2}' -f (Test-Contains $failRun.Log '結果を保存できませんでした: disk full'), [bool]$failRun.Dir, (Test-Contains $failRun.Log '保存先:')) 'True/False/False'
    Assert-Equal '保存の失敗は終了コードに影響しない' ('{0}/{1}' -f $failRun.Code, (Test-Contains $failRun.Log ' OK        web.test')) '0/True'
    $crash = New-PatchedTool 'crash' '$countOk   = @($results | Where-Object { $_.Verdict -eq ''OK'' }).Count' 'throw ''boom'''
    $crashRun = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Os = @{ 'web.test' = @('10.0.0.1') }; Tool = $crash }
    Assert-Equal '想定外のエラーは知らせて、書きかけを残さない' ('{0}/{1}/{2}' -f $crashRun.Code, (Test-Contains $crashRun.Log 'エラーで止まりました。結果は保存していません。'), [bool]$crashRun.Dir) '1/True/False'

    $colorScript = Join-Path $workRoot 'colors.ps1'
    [IO.File]::WriteAllLines($colorScript, [string[]]@(
        'param($Wrapper, $SpecFile)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Wrapper).AddParameter(''SpecFile'', $SpecFile)'
        '[void]$ps.Invoke()'
        'foreach ($record in $ps.Streams.Information) { $m = $record.MessageData; if ($m.Message -match ''^ (OK|注意|引けない) '') { $Matches[1] + ''='' + $m.ForegroundColor } }'
    ), $utf8Bom)
    $colorSpec = New-FakeRun @{ Args = @{ Target = 'ok.test,warn.test,none.test'; NoSave = $true }; Os = @{ 'ok.test' = @('10.0.0.1'); 'warn.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'ok.test' = @('10.0.0.1'); 'warn.test' = @('10.9.9.9') }) }
    $colors = Invoke-Helper $colorScript @('-Wrapper', $wrapper, '-SpecFile', $colorSpec.SpecFile)
    Assert-Equal '「注意」は黄色、「引けない」は赤、OK は色なし' ('{0}/{1}/{2}' -f $colors['注意'], $colors['引けない'], $colors['OK']) 'Yellow/Red/'

    # Ctrl+C の代わりに Stop() で止める
    $stopScript = Join-Path $workRoot 'stop.ps1'
    [IO.File]::WriteAllLines($stopScript, [string[]]@(
        'param($Wrapper, $SpecFile, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Wrapper).AddParameter(''SpecFile'', $SpecFile)'
        '$handle = $ps.BeginInvoke()'
        'for ($i = 0; $i -lt 300 -and @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count -eq 0; $i++) { Start-Sleep -Milliseconds 100 }'
        'Start-Sleep -Milliseconds 800'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
    ), $utf8Bom)
    $stopOs = @{}
    $stopNames = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 1; $i -le 60; $i++) { $n = 's{0}.test' -f $i; $stopNames.Add($n); $stopOs[$n] = @('10.0.0.1') }
    $stopSpec = New-FakeRun @{ Args = @{ Target = ($stopNames -join ',') }; Os = $stopOs; Slow = 200 }
    $stopped = Invoke-Helper $stopScript @('-Wrapper', $wrapper, '-SpecFile', $stopSpec.SpecFile, '-OutRoot', $stopSpec.Out)
    if ($stopped['RUNNING'] -eq 'True') {
        Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $stopped['BEFORE'], $stopped['AFTER']) '1/0'
        Assert-Equal '中断のメッセージ' $stopped['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' '止める前に調べ終わったため'
    }

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $culture = Invoke-Fake @{ Args = @{ Target = 'web.test' }; Culture = 'fi-FI'; Os = @{ 'web.test' = @('10.0.0.1') }; Zone = (New-Zone @{ 'web.test' = @('10.0.0.1') }); Ttl = @{ 'web.test' = 3600 } }
    Assert-Equal 'TTL の区切りは , のまま' (Get-Cell $culture 'web.test' 'TTL') '3,600'
    Assert-Equal '日時の区切りは : のまま' ($culture.Log -match '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   ') 'True'
    Assert-Equal '所要の小数点は . のまま' ($culture.Log -match '所要 \d+\.\d秒') 'True'

    Write-Section '起動のしかた'
    $closed = @('-Server', $closedDns)
    $viaFile = Invoke-Real (@('-Target', 'localhost,127.0.0.1', '-NoSave') + $closed)
    Assert-Equal '-File で -Target にカンマ区切りを渡す' (Test-Contains $viaFile.Log '対象 2 件') 'True'
    $viaServers = Invoke-Real @('-Target', 'localhost', '-Server', "$closedDns,$dnsNet.58", '-NoSave')
    Assert-Equal '-File で -Server にカンマ区切りを渡す' (Test-Contains $viaServers.Log "DNS サーバー: $closedDns, $dnsNet.58（指定）") 'True'
    Assert-Equal '位置で渡す引数は受け付けない' (Invoke-Real @('localhost')).Code 1
    $viaNone = Invoke-Real @()
    Assert-Equal '引数なしは止まる' ('{0}/{1}' -f $viaNone.Code, (Test-Contains $viaNone.Log '調べる名前を指定してください')) '1/True'

    # コマンドプロンプトで "C:\data\" のように書くと、\" が引用符のまま届く

    $quoteCmd = Invoke-ViaCmd @(('{0} -NoProfile -ExecutionPolicy Bypass -File "{1}" -List "{2}\" -NoSave' -f $Exe, $tool, $workRoot), 'exit /b %errorlevel%')
    Assert-Equal 'コマンドプロンプト: 最後が \" の引数は止まる' ('{0}/{1}' -f $quoteCmd.Code, (Test-Contains $quoteCmd.Log '-List の値に引用符（"）が残っています')) '1/True'
    Assert-Equal 'コマンドプロンプト: 正しい書き方を案内する' (Test-Contains $quoteCmd.Log '最後に \ を付けないでください') 'True'
    $cmdExit = Invoke-ViaCmd @(('{0} -NoProfile -ExecutionPolicy Bypass -File "{1}" -Target "localhost" -Server {2} -NoSave' -f $Exe, $tool, $closedDns), 'exit /b %errorlevel%')
    Assert-Equal 'コマンドプロンプト: 終了コード（注意があると 2）' $cmdExit.Code 2

    # 名前に [ ] を含むフォルダから、相対パスで起動する
    Write-Utf8 "$workRoot\[br]\a.csv" "名前`r`nlocalhost`r`n"
    $brCmd = Invoke-ViaCmd @(('cd /d "{0}\[br]"' -f $workRoot), ('{0} -NoProfile -ExecutionPolicy Bypass -File "{1}" -List a.csv -OutputFolder 結果 -Server {2}' -f $Exe, $tool, $closedDns), 'exit /b %errorlevel%')
    Assert-Equal '[ ] を含むフォルダから相対パスの -List' (Test-Contains $brCmd.Log '対象 1 件') 'True'
    Assert-Equal '[ ] を含むフォルダから相対パスの -OutputFolder' (@(Get-ChildItem -LiteralPath "$workRoot\[br]\結果" -Directory -Filter 'NameResolution_*' -ErrorAction SilentlyContinue).Count) 1
    $brNew = Invoke-ViaCmd @(('cd /d "{0}\[br]"' -f $workRoot), ('{0} -NoProfile -ExecutionPolicy Bypass -File "{1}" -NewList new.csv' -f $Exe, $tool), 'exit /b %errorlevel%')
    Assert-Equal '[ ] を含むフォルダから相対パスの -NewList' ('{0}/{1}' -f $brNew.Code, [IO.File]::Exists("$workRoot\[br]\new.csv")) '0/True'
    Write-Utf8 "$workRoot\[br]\c.ini" "Target = localhost`r`nNoSave = true`r`nServer = $closedDns`r`n"
    $brConf = Invoke-ViaCmd @(('cd /d "{0}\[br]"' -f $workRoot), ('{0} -NoProfile -ExecutionPolicy Bypass -File "{1}" -ConfigFile c.ini' -f $Exe, $tool), 'exit /b %errorlevel%')
    Assert-Equal '[ ] を含むフォルダから相対パスの -ConfigFile' (Test-Contains $brConf.Log "設定ファイル: $workRoot\[br]\c.ini") 'True'

    # 既定の保存先はスクリプトのフォルダ（テストは一時フォルダのコピーを動かしている）
    $defaultRun = Invoke-Real (@('-Target', 'localhost') + $closed)
    Assert-Equal '既定の保存先はスクリプトのフォルダ' (@(Get-ChildItem -LiteralPath (Split-Path $tool) -Directory -Filter 'NameResolution_*').Count) 1

    # Windows でのテストで見つかった問題（他のツール）: 5.1 で 260 文字を超えるパスを扱えなかった
    $longDir = "$workRoot\" + ('L' * 120) + '\' + ('M' * 120)
    Write-LongFile "$longDir\names.csv" "名前`r`nlocalhost`r`n"
    $longList = Invoke-Real (@('-List', "$longDir\names.csv", '-NoSave') + $closed)
    Assert-Equal '-List が 260 文字を超えても読む' (Test-Contains $longList.Log '対象 1 件') 'True'
    Write-LongFile "$longDir\c.ini" "Target = localhost`r`nNoSave = true`r`nServer = $closedDns`r`n"
    $longConf = Invoke-Real @('-ConfigFile', "$longDir\c.ini")
    Assert-Equal '-ConfigFile が 260 文字を超えても読む' (Test-Contains $longConf.Log '対象 1 件') 'True'
    $longNew = Invoke-Real @('-NewList', "$longDir\new.csv")
    Assert-Equal '-NewList が 260 文字を超えても作る' ('{0}/{1}' -f $longNew.Code, [IO.File]::Exists("\\?\$longDir\new.csv")) '0/True'
    $longOut = "$longDir\out"
    $longSave = Invoke-Real (@('-Target', 'localhost', '-OutputFolder', $longOut) + $closed)
    $longRuns = @(); try { $longRuns = @([IO.Directory]::GetDirectories("\\?\$longOut")) } catch { }
    Assert-Equal '-OutputFolder が 260 文字を超えても保存する' ('{0}/{1}' -f $longSave.Code, $longRuns.Count) '2/1'
    if ($longRuns.Count -eq 1) { Assert-Equal '長い保存先のファイル' (@([IO.Directory]::GetFiles($longRuns[0])).Count) 2 }

    Write-Section '本物の環境（OS の結果と hosts。問い合わせはループバックだけ）'
    $regKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SYSTEM\CurrentControlSet\Services\Tcpip\Parameters')
    $realHostsPath = Join-Path ([string]$regKey.GetValue('DataBasePath')) 'hosts'
    $regKey.Close()
    $real = Invoke-Real @('-Target', 'localhost', '-Server', $closedDns, '-OutputFolder', "$workRoot\real1")
    Assert-Equal '本物の hosts の場所（OS の設定に従う）' (Test-Contains $real.Log "hosts: $realHostsPath（有効な行 ") 'True'
    Assert-Equal '止まらずに動く' ('{0}/{1}' -f $real.Code, (Test-Contains $real.Log 'エラーで止まりました')) '2/False'
    Assert-Equal 'localhost は OS で引ける' ($real.Log -match ' 注意 +localhost +(hosts|OS のみ) ') 'True'
    Assert-Equal 'DNS が動いていないサーバーは応答しない（本物のエラー）' (Test-Contains $real.Log "DNS サーバー $closedDns が応答しない") 'True'
    $otherPathTool = Join-Path $workRoot 'patched\otherpath\Test-NameResolution.ps1'
    [void][IO.Directory]::CreateDirectory((Split-Path $otherPathTool))
    [IO.File]::WriteAllText($otherPathTool, [IO.File]::ReadAllText($repoTool).Replace('$folder = $key.GetValue(''DataBasePath'')', '$folder = ''C:\wbtest-other-etc'''), $utf8Bom)
    $otherPath = Invoke-Real @('-Target', 'localhost', '-Server', $closedDns, '-NoSave') $otherPathTool
    Assert-Equal 'OS の設定が既定と違えば、その場所を読む' (Test-Contains $otherPath.Log '※ hosts を読めなかったため') 'True'
    $noPathTool = Join-Path $workRoot 'patched\nopath\Test-NameResolution.ps1'
    [void][IO.Directory]::CreateDirectory((Split-Path $noPathTool))
    [IO.File]::WriteAllText($noPathTool, [IO.File]::ReadAllText($repoTool).Replace("GetValue('DataBasePath')", "GetValue('NoSuchValue')"), $utf8Bom)
    $noPath = Invoke-Real @('-Target', 'localhost', '-Server', $closedDns, '-NoSave') $noPathTool
    Assert-Equal 'OS の設定を読めなければ既定の場所' (Test-Contains $noPath.Log ("hosts: {0}（有効な行 " -f (Join-Path $env:windir 'System32\drivers\etc\hosts'))) 'True'

    # 本物の DNS サーバーの設定を読む（問い合わせは偽の関数が受けるので、外へは出ない）
    if (Get-Command Get-DnsClientServerAddress -ErrorAction SilentlyContinue) {
        $connected = @(Get-NetIPInterface -ConnectionState Connected | ForEach-Object { $_.InterfaceIndex })
        $realServers = New-Object 'System.Collections.Generic.List[string]'
        foreach ($entry in @(Get-DnsClientServerAddress)) {
            if ($connected -notcontains $entry.InterfaceIndex) { continue }
            foreach ($s in @($entry.ServerAddresses)) { if ($s -and $s -notlike 'fec0:0:0:ffff::*' -and $realServers -notcontains $s) { $realServers.Add($s) } }
        }
        $realConf = Invoke-Fake @{ Args = @{ Target = 'web'; NoSave = $true }; RealServers = $true; Os = @{ 'web' = @('10.0.0.1') } }
        if ($realServers.Count -gt 0) {
            Assert-Equal '本物の設定の DNS サーバーを読む' (Test-Contains $realConf.Log ('DNS サーバー: {0}（OS の設定）' -f ($realServers -join ', '))) 'True'
            Assert-Equal 'fec0 は聞かない' (Get-QueryCount $realConf 'fec0*') 0
            Assert-Equal '短い名前は最後にサフィックスなしで聞く（本物のサフィックスの設定）' ((@($realConf.Queries)[-1] -split ' ')[2]) 'web'
        } else {
            Assert-Equal '本物の設定に DNS サーバーがない' (Test-Contains $realConf.Log '※ DNS サーバーが設定されていません') 'True'
        }
    } else {
        Write-Skip '本物の DNS サーバーの設定' 'このPCにないため'
    }

    Write-Section 'ループバックの DNS サーバーと、本物の Resolve-DnsName'
    $esc = [string][char]27
    $dnsNormal = "$dnsNet.53"; $dnsSilent = "$dnsNet.54"; $dnsServfail = "$dnsNet.55"; $dnsRefused = "$dnsNet.56"; $dnsSilentAaaa = "$dnsNet.57"; $dnsOther = "$dnsNet.60"
    $dnsZone = @{
        'web.test'        = @{ A = @('10.0.0.1', '10.0.0.2'); AAAA = @('2001:db8::1'); Ttl = 3600 }
        'v4.test'         = @{ A = @('10.0.0.3'); Ttl = 300 }
        'alias.test'      = @{ Cname = @('mid.test', 'end.test'); A = @('10.0.0.4'); AAAA = @('2001:db8::4'); Ttl = 120 }
        'nodata.test'     = @{ Ttl = 60 }
        'evil.test'       = @{ Cname = @("x${esc}[31m.test", "y`nz.test"); A = @('10.0.0.6'); Ttl = 60 }
        'web.conn.test'   = @{ A = @('10.0.0.5'); Ttl = 60 }
        'bare'            = @{ A = @('10.0.0.7'); Ttl = 60 }
        'localhost'       = @{ A = @('127.0.0.1'); AAAA = @('::1'); Ttl = 60 }
        $selfName         = @{ A = @('10.9.9.9'); Ttl = 60 }
        '1.0.0.10.in-addr.arpa' = @{ Ptr = @('web.test'); Ttl = 60 }
        '3.0.0.10.in-addr.arpa' = @{ Ptr = @("evil${esc}[31m.test"); Ttl = 60 }
        '4.0.0.10.in-addr.arpa' = @{ Ptr = @("line`nbreak.test"); Ttl = 60 }
        '4.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.0.8.b.d.0.1.0.0.2.ip6.arpa' = @{ Ptr = @('alias.test'); Ttl = 60 }
    }
    $dnsSpec = @{
        $dnsNormal     = @{ Mode = 'normal'; Zone = $dnsZone }
        $dnsOther      = @{ Mode = 'normal'; Zone = @{ 'web.test' = @{ A = @('10.0.9.9'); Ttl = 60 } } }
        $dnsSilent     = @{ Mode = 'silent'; Zone = @{} }
        $dnsServfail   = @{ Mode = 'servfail'; Zone = @{} }
        $dnsRefused    = @{ Mode = 'refused'; Zone = @{} }
        $dnsSilentAaaa = @{ Mode = 'silentaaaa'; Zone = $dnsZone }
    }
    $dnsSpecFile = Join-Path $workRoot 'dns-spec.json'
    $dnsFile     = Join-Path $workRoot 'dns-server.ps1'
    $dnsReady    = Join-Path $workRoot 'dns-ready'
    $dnsErrFile  = Join-Path $workRoot 'dns-error.txt'
    [IO.File]::WriteAllText($dnsSpecFile, ($dnsSpec | ConvertTo-Json -Depth 8), $utf8Bom)
    $dnsCode = @'
param([string]$SpecFile, [string]$StopFile, [string]$ReadyFile, [string]$ErrorFile)
# ループバックで答える小さな DNS サーバー（UDP）。Mode: normal、silent（答えない）、silentaaaa（AAAA にだけ答えない）、servfail、refused
function ConvertTo-DnsName([string]$Name) {
    $out = New-Object 'System.Collections.Generic.List[byte]'
    foreach ($label in $Name.TrimEnd('.').Split('.')) {
        $bytes = [Text.Encoding]::UTF8.GetBytes($label)
        $out.Add([byte]$bytes.Length)
        $out.AddRange([byte[]]$bytes)
    }
    $out.Add(0)
    return , $out.ToArray()
}
function Add-Record($List, [byte[]]$Owner, [int]$Type, [int]$Ttl, [byte[]]$Data) {
    $List.AddRange($Owner)
    $List.Add([byte]($Type -shr 8)); $List.Add([byte]($Type -band 255))
    $List.Add(0); $List.Add(1)
    foreach ($shift in 24, 16, 8, 0) { $List.Add([byte](($Ttl -shr $shift) -band 255)) }
    $List.Add([byte]($Data.Length -shr 8)); $List.Add([byte]($Data.Length -band 255))
    $List.AddRange($Data)
}
function New-Reply([byte[]]$Query, $Conf) {
    $at = 12
    $labels = @()
    while ($Query[$at] -ne 0) { $size = $Query[$at]; $labels += [Text.Encoding]::UTF8.GetString($Query, $at + 1, $size); $at += $size + 1 }
    $name = $labels -join '.'
    $at++
    $type = ($Query[$at] * 256) + $Query[$at + 1]
    $at += 4
    $mode = [string]$Conf.Mode
    if ($mode -eq 'silent' -or ($mode -eq 'silentaaaa' -and $type -eq 28)) { return $null }
    $rcode = 0
    if ($mode -eq 'servfail') { $rcode = 2 }
    if ($mode -eq 'refused') { $rcode = 5 }
    $answers = New-Object 'System.Collections.Generic.List[byte]'
    $authority = New-Object 'System.Collections.Generic.List[byte]'
    $answerCount = 0
    $authorityCount = 0
    if ($rcode -eq 0) {
        $entry = $Conf.Zone.PSObject.Properties[$name]
        if (-not $entry) {
            $rcode = 3
        } else {
            $record = $entry.Value
            $owner = ConvertTo-DnsName $name
            $found = 0
            if ($record.PSObject.Properties['Cname']) {
                foreach ($next in @($record.Cname)) {
                    $target = ConvertTo-DnsName $next
                    Add-Record $answers $owner 5 ([int]$record.Ttl) $target
                    $answerCount++
                    $owner = $target
                }
            }
            if ($type -eq 1 -and $record.PSObject.Properties['A']) {
                foreach ($ip in @($record.A)) { Add-Record $answers $owner 1 ([int]$record.Ttl) ([Net.IPAddress]::Parse($ip).GetAddressBytes()); $answerCount++; $found++ }
            } elseif ($type -eq 28 -and $record.PSObject.Properties['AAAA']) {
                foreach ($ip in @($record.AAAA)) { Add-Record $answers $owner 28 ([int]$record.Ttl) ([Net.IPAddress]::Parse($ip).GetAddressBytes()); $answerCount++; $found++ }
            } elseif ($type -eq 12 -and $record.PSObject.Properties['Ptr']) {
                foreach ($target in @($record.Ptr)) { Add-Record $answers $owner 12 ([int]$record.Ttl) (ConvertTo-DnsName $target); $answerCount++; $found++ }
            }
            if ($found -eq 0) {
                # 名前はあるが、その種類の登録がない（権威の SOA を返す）
                $soa = New-Object 'System.Collections.Generic.List[byte]'
                $soa.AddRange([byte[]](ConvertTo-DnsName 'ns.test'))
                $soa.AddRange([byte[]](ConvertTo-DnsName 'admin.test'))
                foreach ($value in 1, 3600, 600, 86400, 60) { foreach ($shift in 24, 16, 8, 0) { $soa.Add([byte](($value -shr $shift) -band 255)) } }
                Add-Record $authority (ConvertTo-DnsName 'test') 6 60 $soa.ToArray()
                $authorityCount++
            }
        }
    }
    $reply = New-Object 'System.Collections.Generic.List[byte]'
    $reply.Add($Query[0]); $reply.Add($Query[1])
    $reply.Add(0x81); $reply.Add([byte](0x80 -bor $rcode))
    $reply.Add(0); $reply.Add(1)
    $reply.Add([byte]($answerCount -shr 8)); $reply.Add([byte]($answerCount -band 255))
    $reply.Add(0); $reply.Add([byte]$authorityCount)
    $reply.Add(0); $reply.Add(0)
    for ($k = 12; $k -lt $at; $k++) { $reply.Add($Query[$k]) }
    $reply.AddRange($answers)
    $reply.AddRange($authority)
    return , $reply.ToArray()
}
try {
    $spec = [IO.File]::ReadAllText($SpecFile) | ConvertFrom-Json
    $sockets = New-Object System.Collections.ArrayList
    foreach ($p in $spec.PSObject.Properties) {
        $socket = New-Object Net.Sockets.Socket ([Net.Sockets.AddressFamily]::InterNetwork), ([Net.Sockets.SocketType]::Dgram), ([Net.Sockets.ProtocolType]::Udp)
        $socket.Bind((New-Object Net.IPEndPoint ([Net.IPAddress]::Parse($p.Name)), 53))
        [void]$sockets.Add(@{ Socket = $socket; Conf = $p.Value })
    }
    [IO.File]::WriteAllText($ReadyFile, 'ready')
    $buffer = New-Object byte[] 4096
    # テストが途中で止まっても残らないよう、時間で終わる
    $deadline = (Get-Date).AddMinutes(15)
    while (-not [IO.File]::Exists($StopFile) -and (Get-Date) -lt $deadline) {
        foreach ($entry in $sockets) {
            if ($entry.Socket.Poll(5000, [Net.Sockets.SelectMode]::SelectRead)) {
                $from = [Net.EndPoint](New-Object Net.IPEndPoint ([Net.IPAddress]::Any), 0)
                try { [void]$entry.Socket.ReceiveFrom($buffer, [ref]$from) } catch { continue }
                try {
                    $reply = New-Reply $buffer $entry.Conf
                    if ($reply) { [void]$entry.Socket.SendTo($reply, $from) }
                } catch { [IO.File]::AppendAllText($ErrorFile, $_.Exception.Message + "`r`n") }
            }
        }
    }
    foreach ($entry in $sockets) { $entry.Socket.Close() }
} catch {
    [IO.File]::AppendAllText($ErrorFile, $_.Exception.Message + "`r`n")
}
'@
    [IO.File]::WriteAllText($dnsFile, $dnsCode, $utf8Bom)
    $dnsServer = Start-Process -FilePath $Exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$dnsFile`"", '-SpecFile', "`"$dnsSpecFile`"", '-StopFile', "`"$dnsStop`"", '-ReadyFile', "`"$dnsReady`"", '-ErrorFile', "`"$dnsErrFile`"") -PassThru -WindowStyle Hidden
    for ($i = 0; $i -lt 200 -and -not [IO.File]::Exists($dnsReady) -and -not $dnsServer.HasExited; $i++) { [Threading.Thread]::Sleep(100) }

    if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) {
        Write-Skip 'ループバックの DNS サーバー' 'Resolve-DnsName がないため'
    } elseif (-not [IO.File]::Exists($dnsReady)) {
        $why = if ([IO.File]::Exists($dnsErrFile)) { [IO.File]::ReadAllText($dnsErrFile).Trim() } else { '準備が終わりませんでした' }
        Write-Skip 'ループバックの DNS サーバー' "UDP 53 番を待ち受けられないため: $why"
    } else {
        $liveOs = @{ 'web.test' = @('10.0.0.1', '10.0.0.2', '2001:db8::1'); 'v4.test' = @('10.0.0.3'); 'alias.test' = @('10.0.0.4', '2001:db8::4'); 'nodata.test' = @('10.0.0.8'); 'evil.test' = @('10.0.0.6') }
        $live = Invoke-Fake @{ Args = @{ Target = 'web.test,v4.test,alias.test,nodata.test,none.test'; Server = $dnsNormal }; RealDns = $true; Os = $liveOs }
        Assert-Equal '本物: A と AAAA' ('{0}/{1}' -f (Get-Verdict $live 'web.test'), (Get-Cell $live 'web.test' 'DNS の結果')) "OK/${dnsNormal}: 10.0.0.1, 10.0.0.2, 2001:db8::1"
        Assert-Equal '本物: TTL' (Get-Cell $live 'web.test' 'TTL') '3,600'
        Assert-Equal '本物: 解決した名前' (Get-Cell $live 'web.test' '解決した名前') 'web.test'
        Assert-Equal '本物: AAAA の登録がない（SOA が返る）' ('{0}/{1}' -f (Get-Verdict $live 'v4.test'), (Get-Cell $live 'v4.test' 'DNS の結果')) "OK/${dnsNormal}: 10.0.0.3"
        Assert-Equal '本物: CNAME' ('{0}/{1}/{2}' -f (Get-Cell $live 'alias.test' 'CNAME'), (Get-Cell $live 'alias.test' '解決した名前'), (Get-Cell $live 'alias.test' 'TTL')) 'alias.test → mid.test → end.test/end.test/120'
        Assert-Equal '本物: A / AAAA の登録がない' ('{0}/{1}/{2}' -f (Get-Verdict $live 'nodata.test'), (Get-Cell $live 'nodata.test' '引けた元'), (Get-Cell $live 'nodata.test' 'DNS の結果')) "OK/その他/${dnsNormal}: A / AAAA の登録がない"
        Assert-Equal '本物: 名前が存在しない' ('{0}/{1}' -f (Get-Verdict $live 'none.test'), (Get-Cell $live 'none.test' 'DNS の結果')) "引けない/${dnsNormal}: 名前が存在しない"
        $liveExpect = Invoke-Fake @{ Args = @{ Target = 'v4.test=10.0.0.3,web.test=10.0.0.9'; Server = $dnsNormal }; RealDns = $true; Os = $liveOs }
        Assert-Equal '本物: 期待する IP アドレス' ('{0}/{1}' -f (Get-Verdict $liveExpect 'v4.test'), (Get-Cell $liveExpect 'web.test' '備考')) "OK/期待する IP アドレスと違う（OS の結果、DNS サーバー $dnsNormal）"
        $liveDiff = Invoke-Fake @{ Args = @{ Target = 'web.test'; Server = "$dnsNormal,$dnsOther" }; RealDns = $true; Os = $liveOs }
        Assert-Equal '本物: サーバーによって答えが違う' (Test-Note $liveDiff 'web.test' 'DNS サーバーによって答えが違う') 'True'
        $liveErr = Invoke-Fake @{ Args = @{ Target = 'web.test'; Server = "$dnsServfail,$dnsRefused,$closedDns,$dnsNormal" }; RealDns = $true; Os = $liveOs }
        Assert-Equal '本物: SERVFAIL' (Test-Note $liveErr 'web.test' "DNS サーバー ${dnsServfail}: SERVFAIL") 'True'
        Assert-Equal '本物: REFUSED' (Test-Note $liveErr 'web.test' "DNS サーバー ${dnsRefused}: REFUSED") 'True'
        Assert-Equal '本物: DNS が動いていないサーバーは応答しない' (Test-Note $liveErr 'web.test' "DNS サーバー $closedDns が応答しない") 'True'
        Assert-Equal '本物: エラーのサーバーがあっても答えは使う' ('{0}/{1}' -f (Get-Cell $liveErr 'web.test' '引けた元'), (Get-Verdict $liveErr 'web.test')) 'DNS/注意'
        $liveRev = Invoke-Fake @{ Args = @{ Target = '10.0.0.1,10.0.0.99,2001:db8::4'; Server = $dnsNormal }; RealDns = $true }
        Assert-Equal '本物: 逆引き（IPv4）' (Get-Cell $liveRev '10.0.0.1' '備考') '逆引き: web.test'
        Assert-Equal '本物: 逆引きの登録がない' (Get-Cell $liveRev '10.0.0.99' '備考') '逆引きの登録がない（10.0.0.99）'
        Assert-Equal '本物: 逆引き（IPv6）' (Get-Cell $liveRev '2001:db8::4' '備考') '逆引き: alias.test'
        $liveRevName = Invoke-Fake @{ Args = @{ Target = 'web.test'; Reverse = $true; Server = $dnsNormal }; RealDns = $true; Os = $liveOs }
        Assert-Equal '本物: -Reverse' (Get-Cell $liveRevName 'web.test' '逆引き') '10.0.0.1: 一致 / 10.0.0.2: なし / 2001:db8::1: なし'
        $liveEvil = Invoke-Fake @{ Args = @{ Target = 'evil.test,10.0.0.3,10.0.0.4'; Server = $dnsNormal }; RealDns = $true; Os = $liveOs }
        Assert-Equal '本物: CNAME の制御文字' (Get-Cell $liveEvil 'evil.test' 'CNAME') 'evil.test → x?[31m.test → y?z.test'
        Assert-Equal '本物: 逆引きのエスケープ文字' (Get-Cell $liveEvil '10.0.0.3' '備考') '逆引き: evil?[31m.test'
        Assert-Equal '本物: 逆引きの改行' (Get-Cell $liveEvil '10.0.0.4' '備考') '逆引き: line?break.test'
        Assert-Equal '本物: 画面にエスケープ文字を出さない' $liveEvil.Log.Contains($esc) 'False'
        $liveShort = Invoke-Fake @{ Args = @{ Target = 'web'; Server = $dnsNormal }; RealDns = $true; ConnSuffix = 'conn.test'; Os = @{ 'web' = @('10.0.0.5') } }
        Assert-Equal '本物: 短い名前' (Get-Cell $liveShort 'web' '解決した名前') 'web.conn.test'
        $liveBare = Invoke-Fake @{ Args = @{ Target = 'bare'; Server = $dnsNormal }; RealDns = $true; Os = @{ 'bare' = @('10.0.0.7') } }
        Assert-Equal '本物: サフィックスなしの名前も DNS サーバーに届く' (Get-Cell $liveBare 'bare' 'DNS の結果') "${dnsNormal}: 10.0.0.7"
        # -Server で聞いた答えは DNS のキャッシュに入る（ipconfig /displaydns に出る）が、アプリの名前解決が使うキャッシュとは別に持つ。
        # Get-DnsClientCache は管理者でないとエラーも出さずに 0 件を返すので使わない。-CacheOnly は管理者でなくても、アプリが使うキャッシュを読める
        $appCache = 'なし'
        try {
            $cached = @(Resolve-DnsName 'web.test' -Type A -CacheOnly -ErrorAction Stop | Where-Object { $_.Section -eq 'Answer' } | ForEach-Object { $_.IPAddress })
            if ($cached.Count -gt 0) { $appCache = $cached -join ', ' }
        } catch { }
        Assert-Equal '本物: アプリが使う DNS のキャッシュに入れない' $appCache 'なし'
        # 答えない問い合わせは 1 回で約 11 秒かかる
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $liveSilent = Invoke-Fake @{ Args = @{ Target = 'web.test'; Server = "$dnsSilent,$dnsSilentAaaa" }; RealDns = $true; Os = $liveOs }
        Assert-Equal '本物: 答えないサーバーは応答しない（タイムアウト）' (Test-Note $liveSilent 'web.test' "DNS サーバー $dnsSilent が応答しない") 'True'
        Assert-Equal '本物: AAAA にだけ答えないサーバーの答えを使う' ('{0}/{1}' -f (Test-Note $liveSilent 'web.test' "DNS サーバー ${dnsSilentAaaa}: AAAA は応答なし"), (Get-Cell $liveSilent 'web.test' '引けた元')) 'True/DNS'

        # 偽の環境を使わず、本物の OS の結果とループバックの DNS サーバーで動かす
        $liveTool = Invoke-Real @('-Target', 'localhost', '-Server', $dnsNormal, '-OutputFolder', "$workRoot\real2")
        $liveCsv = @(Import-Csv -LiteralPath (Join-Path @(Get-ChildItem -LiteralPath "$workRoot\real2" -Directory)[0].FullName 'results.csv') -Encoding UTF8)
        Assert-Equal '本物だけ: localhost' ('{0}/{1}' -f $liveTool.Code, $liveCsv[0].'判定') '0/OK'
        $liveSelf = Invoke-Real @('-Target', $selfName, '-Server', $dnsNormal, '-NoSave')
        Assert-Equal '本物だけ: この PC の名前は OS の結果と比べない' ('{0}/{1}' -f $liveSelf.Code, (Test-Contains $liveSelf.Log 'キャッシュ')) '0/False'
    }
} finally {
    if ($dnsServer) {
        try {
            [IO.File]::WriteAllText($dnsStop, 'stop')
            [void]$dnsServer.WaitForExit(5000)
            if (-not $dnsServer.HasExited) { $dnsServer.Kill() }
        } catch { }
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

'  END'
exit $script:failCount
