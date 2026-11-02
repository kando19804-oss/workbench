<#
.SYNOPSIS
    Test-NameResolution の回帰テストです。

.DESCRIPTION
    DNS サーバーの答えを自由に作れないため、Resolve-DnsName などと同じ名前の偽の関数を定義したラッパーから
    Test-NameResolution を実行し、判定を確かめます。hosts は環境変数 SystemRoot を一時フォルダに向けて読ませます。
    OS の結果（[Net.Dns]::GetHostAddresses）は偽にできないので、必ず解決できる localhost と、解決できない .invalid の名前を使います。
    最後に、偽の関数を使わずに本物の Resolve-DnsName でも一度動かします。ダミー環境は終了時に削除します。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Test-NameResolution\Test-NameResolution.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Test-NameResolution_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
# 結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーしたスクリプトを動かす
$tool     = Join-Path $workRoot 'tool\Test-NameResolution.ps1'
$wrapper  = Join-Path $workRoot 'wrapper.ps1'
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
    [void][IO.Directory]::CreateDirectory((Split-Path $FilePath))
    [IO.File]::WriteAllText($FilePath, $Text, (New-Object Text.UTF8Encoding $false))
}

# 偽の環境で動かす。Spec: Args（ツールの引数）、Zone、Ptr、Servers、SearchList、ConnSuffix、Hosts（hosts の中身。$null なら hosts なし）
function Invoke-Fake([hashtable]$Spec) {
    $script:runNumber++
    $number = $script:runNumber
    $root = Join-Path $workRoot "root_$number"
    $etc  = Join-Path $root 'System32\drivers\etc'
    [void][IO.Directory]::CreateDirectory($etc)
    if ($Spec.ContainsKey('Hosts') -and $null -ne $Spec.Hosts) { Write-Utf8 (Join-Path $etc 'hosts') $Spec.Hosts }
    $out = Join-Path $workRoot "out_$number"
    $toolArgs = @{}
    foreach ($key in $Spec.Args.Keys) { $toolArgs[$key] = $Spec.Args[$key] }
    if (-not $toolArgs.ContainsKey('NoSave')) { $toolArgs['OutputFolder'] = $out }
    $full = @{
        Tool = $tool; Root = $root; Args = $toolArgs
        Zone = $(if ($Spec.ContainsKey('Zone')) { $Spec.Zone } else { @{} })
        Ptr = $(if ($Spec.ContainsKey('Ptr')) { $Spec.Ptr } else { @{} })
        Servers = $(if ($Spec.ContainsKey('Servers')) { $Spec.Servers } else { @('10.0.0.53') })
        SearchList = $(if ($Spec.ContainsKey('SearchList')) { $Spec.SearchList } else { @() })
        ConnSuffix = $(if ($Spec.ContainsKey('ConnSuffix')) { $Spec.ConnSuffix } else { '' })
    }
    $specFile = Join-Path $workRoot "spec_$number.json"
    [IO.File]::WriteAllText($specFile, ($full | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding $true))
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $wrapper -SpecFile $specFile 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $dir = $null
    if (Test-Path -LiteralPath $out) { $dir = @(Get-ChildItem -LiteralPath $out -Directory | Where-Object { $_.Name -like 'NameResolution_*' } | ForEach-Object { $_.FullName })[0] }
    $rows = @()
    if ($dir -and (Test-Path -LiteralPath (Join-Path $dir 'results.csv'))) { $rows = @(Import-Csv -LiteralPath (Join-Path $dir 'results.csv') -Encoding UTF8) }
    [pscustomobject]@{ Code = $code; Log = $log; Dir = $dir; Rows = $rows }
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

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

try {
    #region ダミー環境

    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    $wrapperCode = @'
param([string]$SpecFile)
$spec = [IO.File]::ReadAllText($SpecFile) | ConvertFrom-Json
# hosts の場所は OS の設定（%SystemRoot%\System32\drivers\etc）に従うので、SystemRoot を向け替える
$env:SystemRoot = $spec.Root
$env:windir = $spec.Root
$global:fake = $spec
function global:Get-NetIPInterface { [pscustomobject]@{ InterfaceIndex = 1 } }
function global:Get-DnsClientServerAddress { [pscustomobject]@{ InterfaceIndex = 1; ServerAddresses = @($global:fake.Servers + 'fec0:0:0:ffff::1') }; [pscustomobject]@{ InterfaceIndex = 2; ServerAddresses = @('10.8.8.8') } }
function global:Get-DnsClientGlobalSetting { [pscustomobject]@{ SuffixSearchList = @($global:fake.SearchList) } }
function global:Get-DnsClient { [pscustomobject]@{ ConnectionSpecificSuffix = $global:fake.ConnSuffix } }
function global:Resolve-DnsName {
    param($Name, $Type, [switch]$DnsOnly, [switch]$NoHostsFile, $Server, $ErrorAction)
    if ($Server -eq '10.0.0.99') { throw [ComponentModel.Win32Exception]::new(1460) }                                 # 応答しない
    if ($Server -eq '10.0.0.77' -and $Type -eq 'AAAA') { throw [ComponentModel.Win32Exception]::new(1460) }          # AAAA だけ応答しない
    if ($Server -eq '10.0.0.78' -and $Type -eq 'A') { throw [ComponentModel.Win32Exception]::new(9002) }             # A はエラー
    if ($Server -eq '10.0.0.66' -and $Type -eq 'PTR') { throw [ComponentModel.Win32Exception]::new(1460) }           # 逆引きだけ応答しない
    if ($Type -eq 'PTR') {
        $ptr = $global:fake.Ptr.PSObject.Properties[$Name]
        if (-not $ptr) { throw [ComponentModel.Win32Exception]::new(9003) }
        foreach ($n in @($ptr.Value)) { [pscustomobject]@{ Section = 'Answer'; Type = 'PTR'; Name = $Name; NameHost = "$n."; TTL = 60 } }
        return
    }
    $entry = $global:fake.Zone.PSObject.Properties[$Name]
    if (-not $entry) { throw [ComponentModel.Win32Exception]::new(9003) }
    $perServer = $entry.Value.PSObject.Properties[$Server]
    if (-not $perServer) { $perServer = $entry.Value.PSObject.Properties['*'] }
    if (-not $perServer) { throw [ComponentModel.Win32Exception]::new(9003) }
    $wantV6 = $Type -eq 'AAAA'
    $ips = @($perServer.Value | Where-Object { ([string]$_).Contains(':') -eq $wantV6 })
    if ($ips.Count -eq 0) { throw [ComponentModel.Win32Exception]::new(9501) }
    if ($Name -like 'cname.*') { [pscustomobject]@{ Section = 'Answer'; Type = 'CNAME'; Name = $Name; NameHost = 'lb.test'; TTL = 300 } }
    foreach ($ip in $ips) { [pscustomobject]@{ Section = 'Answer'; Type = $Type; Name = $Name; IPAddress = [string]$ip; TTL = 120 } }
}
$named = @{}
foreach ($p in $spec.Args.PSObject.Properties) { $named[$p.Name] = $p.Value }
& $spec.Tool @named
exit $LASTEXITCODE
'@
    [IO.File]::WriteAllText($wrapper, $wrapperCode, (New-Object Text.UTF8Encoding $true))

    # OS の結果が localhost で何を返すかは環境によるので、先に調べて答えに使う
    $localIps = @([Net.Dns]::GetHostAddresses('localhost') | ForEach-Object { $_.ToString() })
    $bad = 'no-such-name.invalid'

    #endregion

    Write-Section 'OK と、保存するファイル'
    $run = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = '# 空'; Zone = @{ localhost = @{ '*' = $localIps } } }
    Assert-Equal '判定と終了コード' ('{0}/{1}' -f (Get-Verdict $run 'localhost'), $run.Code) 'OK/0'
    Assert-Equal '引けた元' (Get-Cell $run 'localhost' '引けた元') 'DNS'
    Assert-Equal 'CSV の列' (($run.Rows[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',') '判定,名前,引けた元,OS の結果,hosts の値,hosts の行番号,DNS の結果,解決した名前,CNAME,TTL,期待する IP,逆引き,メモ,備考'
    Assert-Equal '保存するファイル' ((Get-ChildItem -LiteralPath $run.Dir | ForEach-Object { $_.Name } | Sort-Object) -join ',') 'results.csv,summary.txt'
    Assert-Equal 'CSV の BOM' ([IO.File]::ReadAllBytes((Join-Path $run.Dir 'results.csv'))[0..2] -join ',') '239,187,191'
    Assert-Equal 'TTL' (Get-Cell $run 'localhost' 'TTL') '120'
    Assert-Equal '使わない値（fec0、切断した NIC）のサーバーに聞かない' (Test-Contains $run.Log 'fec0') 'False'
    Assert-Equal 'つながっていない NIC の DNS サーバーに聞かない' (Test-Contains $run.Log '10.8.8.8') 'False'
    Assert-Equal 'hosts の表示' (Test-Contains $run.Log 'hosts: ') 'True'

    $cname = Invoke-Fake @{ Args = @{ Target = 'cname.test' }; Hosts = ''; Zone = @{ 'cname.test' = @{ '*' = @('10.0.0.1') } } }
    Assert-Equal 'CNAME' (Get-Cell $cname 'cname.test' 'CNAME') 'cname.test → lb.test'

    Write-Section '引けない'
    $fail = Invoke-Fake @{ Args = @{ Target = $bad }; Hosts = '' }
    Assert-Equal '引けない' ('{0}/{1}/{2}' -f (Get-Verdict $fail $bad), (Get-Cell $fail $bad '引けた元'), $fail.Code) '引けない/-/2'
    Assert-Equal '備考' (Test-Note $fail $bad 'DNS: 名前が存在しない') 'True'
    $noData = Invoke-Fake @{ Args = @{ Target = $bad }; Hosts = ''; Zone = @{ $bad = @{ 'none' = @('10.0.0.1') } } }
    Assert-Equal 'サーバーごとに答えがない' (Test-Note $noData $bad '名前が存在しない') 'True'
    $reg = Invoke-Fake @{ Args = @{ Target = $bad }; Hosts = ''; Zone = @{ $bad = @{ '*' = @('10.0.0.1') } } }
    Assert-Equal 'DNS にあっても OS で引けないと「引けない」' ('{0}/{1}' -f (Get-Verdict $reg $bad), (Test-Note $reg $bad 'OS では引けない')) '引けない/True'

    Write-Section 'キャッシュと hosts'
    $stale = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = ''; Zone = @{ localhost = @{ '*' = @('10.1.1.1') } } }
    Assert-Equal 'OS の結果が DNS と違うとキャッシュの案内' ('{0}/{1}' -f (Get-Verdict $stale 'localhost'), (Test-Note $stale 'localhost' 'キャッシュが古い可能性')) '注意/True'
    $staleServer = Invoke-Fake @{ Args = @{ Target = 'localhost'; Server = '10.0.0.54' }; Hosts = ''; Zone = @{ localhost = @{ '*' = @('10.1.1.1') } } }
    Assert-Equal '-Server のときはキャッシュのせいにしない' ('{0}/{1}' -f (Test-Note $staleServer 'localhost' '指定した DNS サーバーの答えが違う'), (Test-Note $staleServer 'localhost' 'キャッシュ')) 'True/False'
    $selfName = [Net.Dns]::GetHostName()
    $self = Invoke-Fake @{ Args = @{ Target = $selfName }; Hosts = ''; Zone = @{ $selfName = @{ '*' = @('10.1.1.1') } } }
    Assert-Equal 'この PC 自身の名前は OS の結果と比べない' (Test-Note $self $selfName 'キャッシュ') 'False'

    $diff = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = "# 1`r`n10.9.9.9 localhost # コメント`r`n"; Zone = @{ localhost = @{ '*' = @('10.1.1.1') } } }
    Assert-Equal 'hosts と DNS が違う' ('{0}/{1}/{2}' -f (Get-Verdict $diff 'localhost'), (Get-Cell $diff 'localhost' '引けた元'), (Test-Note $diff 'localhost' 'hosts と DNS が違う（hosts の 2 行目）')) '注意/hosts/True'
    Assert-Equal 'hosts の値と行番号' ('{0}/{1}' -f (Get-Cell $diff 'localhost' 'hosts の値'), (Get-Cell $diff 'localhost' 'hosts の行番号')) '10.9.9.9/2'
    $inc = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = '10.1.1.1 localhost'; Zone = @{ localhost = @{ '*' = @('10.1.1.1', '10.1.1.2') } } }
    Assert-Equal 'hosts の値が DNS の答えに含まれれば違わない' (Get-Verdict $inc 'localhost') 'OK'
    $v6 = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = '10.1.1.1 localhost'; Zone = @{ localhost = @{ '*' = @('10.1.1.1', '2001:db8::1') } } }
    Assert-Equal 'IPv4 だけの hosts と両方の DNS は違わない' (Get-Verdict $v6 'localhost') 'OK'
    $multi = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = "10.1.1.1 localhost`r`n10.1.1.2 localhost`r`n"; Zone = @{ localhost = @{ '*' = @('10.1.1.1') } } }
    Assert-Equal 'hosts に同じ名前が複数' (Test-Note $multi 'localhost' 'hosts に同じ名前が複数ある（1 行目と 2 行目）') 'True'
    $mixed = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = "10.1.1.1 localhost`r`n::1 localhost`r`n"; Zone = @{ localhost = @{ '*' = @('10.1.1.1') } } }
    Assert-Equal 'IPv4 と IPv6 が 1 つずつなら複数としない' (Test-Note $mixed 'localhost' '複数') 'False'
    $hostsOnly = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = '10.1.1.1 localhost' }
    Assert-Equal 'hosts にだけある名前は OK' ('{0}/{1}' -f (Get-Verdict $hostsOnly 'localhost'), (Test-Note $hostsOnly 'localhost' 'DNS にはない')) 'OK/True'
    $other = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = '' }
    Assert-Equal 'DNS になく OS だけが解決' ('{0}/{1}/{2}' -f (Get-Verdict $other 'localhost'), (Get-Cell $other 'localhost' '引けた元'), (Test-Note $other 'localhost' 'キャッシュ、NetBIOS、LLMNR')) 'OK/その他/True'
    $zen = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = "１０.０.０.１ zenkaku.test`r`nnot-ip x.test`r`n10.1.1.1 localhost`r`n" }
    Assert-Equal '全角の数字や IP でない行は読み飛ばす' ('{0}/{1}' -f (Get-Cell $zen 'localhost' 'hosts の値'), (Test-Contains $zen.Log 'エラーで止まりました')) '10.1.1.1/False'

    $noHosts = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = $null }
    Assert-Equal 'hosts を読めなくても続ける' ('{0}/{1}' -f (Test-Contains $noHosts.Log 'hosts を読めなかったため'), (Get-Verdict $noHosts 'localhost')) 'True/OK'
    $noHostsAll = Invoke-Fake @{ Args = @{ Hosts = $true }; Hosts = $null }
    Assert-Equal '-Hosts で hosts を読めないと止まる' ('{0}/{1}' -f $noHostsAll.Code, (Test-Contains $noHostsAll.Log 'hosts を読めません')) '1/True'

    Write-Section 'DNS サーバー'
    $split = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = ''; Servers = @('10.0.0.53', '10.0.0.54'); Zone = @{ localhost = @{ '10.0.0.53' = @('10.5.5.1'); '10.0.0.54' = @('10.5.5.2') } } }
    Assert-Equal 'サーバーによって答えが違う' ('{0}/{1}' -f (Test-Note $split 'localhost' 'DNS サーバーによって答えが違う'), (Test-Contains $split.Log 'サーバーで違う')) 'True/True'
    Assert-Equal 'サーバーごとに CSV に書く' (Get-Cell $split 'localhost' 'DNS の結果') '10.0.0.53: 10.5.5.1 / 10.0.0.54: 10.5.5.2'
    $half = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = ''; Servers = @('10.0.0.53', '10.0.0.54'); Zone = @{ localhost = @{ '10.0.0.53' = $localIps } } }
    Assert-Equal '片方だけが答える' (Test-Note $half 'localhost' 'DNS サーバーによって答えが違う') 'True'
    $roundRobin = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = ''; Servers = @('10.0.0.53', '10.0.0.54'); Zone = @{ localhost = @{ '10.0.0.53' = @('10.1.1.1', '10.1.1.2'); '10.0.0.54' = @('10.1.1.2', '10.1.1.1') } } }
    Assert-Equal '順番だけの違いは同じ' (Test-Note $roundRobin 'localhost' 'サーバーによって') 'False'
    $dead = Invoke-Fake @{ Args = @{ Target = "localhost,a.invalid,b.invalid,c.invalid"; Server = '10.0.0.99,10.0.0.53' }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }
    Assert-Equal '応答しないサーバーは注意' ('{0}/{1}' -f (Get-Verdict $dead 'localhost'), (Test-Note $dead 'localhost' 'DNS サーバー 10.0.0.99 が応答しない')) '注意/True'
    Assert-Equal '3 回続くと聞かない' (Test-Note $dead 'c.invalid' 'DNS サーバー 10.0.0.99') 'True'
    Assert-Equal '4 件目は聞いていないと書く' (Get-Cell $dead 'c.invalid' 'DNS の結果') '10.0.0.99: 応答しないので聞いていない / 10.0.0.53: 名前が存在しない'
    $aaaaDead = Invoke-Fake @{ Args = @{ Target = 'localhost,a.invalid,b.invalid,c.invalid'; Server = '10.0.0.77' }; Hosts = ''; Zone = @{ localhost = @{ '*' = @('127.0.0.1') }; 'a.invalid' = @{ '*' = @('10.0.0.1') }; 'b.invalid' = @{ '*' = @('10.0.0.1') }; 'c.invalid' = @{ '*' = @('10.0.0.1') } } }
    Assert-Equal 'AAAA だけ応答しないサーバーの答えを捨てない' ('{0}/{1}' -f (Get-Cell $aaaaDead 'localhost' '引けた元'), (Test-Note $aaaaDead 'localhost' 'AAAA は応答なし')) 'DNS/True'
    Assert-Equal 'AAAA だけ応答しないサーバーは聞き続ける' (Test-Note $aaaaDead 'c.invalid' '聞いていない') 'False'
    $aError = Invoke-Fake @{ Args = @{ Target = 'localhost'; Server = '10.0.0.78' }; Hosts = ''; Zone = @{ localhost = @{ '*' = @('::1') } } }
    Assert-Equal '片方がエラーなら注意' (Get-Verdict $aError 'localhost') '注意'
    $noServer = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = ''; Servers = @() }
    Assert-Equal 'DNS サーバーがないと注意' ('{0}/{1}/{2}/{3}' -f (Get-Verdict $noServer 'localhost'), (Get-Cell $noServer 'localhost' '引けた元'), (Test-Note $noServer 'localhost' 'DNS サーバーに聞けなかった'), $noServer.Code) '注意/OS のみ/True/2'

    Write-Section '短い名前'
    $short = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = ''; SearchList = @('first.test', 'second.test'); ConnSuffix = 'conn.test'; Zone = @{ 'localhost.second.test' = @{ '*' = $localIps }; 'localhost.conn.test' = @{ '*' = @('10.0.0.9') } } }
    Assert-Equal '検索一覧を順に付ける（接続ごとのサフィックスは使わない）' (Get-Cell $short 'localhost' '解決した名前') 'localhost.second.test'
    $conn = Invoke-Fake @{ Args = @{ Target = 'localhost' }; Hosts = ''; ConnSuffix = 'conn.test'; Zone = @{ 'localhost.conn.test' = @{ '*' = $localIps } } }
    Assert-Equal '検索一覧がなければ接続ごとのサフィックス' (Get-Cell $conn 'localhost' '解決した名前') 'localhost.conn.test'

    Write-Section '期待する IP アドレス'
    $v4 = @($localIps | Where-Object { -not $_.Contains(':') })
    if ($v4.Count -eq 0) { $v4 = @('127.0.0.1') }
    $match = Invoke-Fake @{ Args = @{ Target = "localhost=$($v4[0])" }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }
    Assert-Equal '一致（種類が違う IP アドレスは比べない）' ('{0}/{1}' -f (Get-Verdict $match 'localhost'), (Get-Cell $match 'localhost' '期待する IP')) ('OK/{0}' -f $v4[0])
    $wrong = Invoke-Fake @{ Args = @{ Target = 'localhost=10.2.2.2' }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }
    Assert-Equal '違う（OS と DNS の両方を比べる）' ('{0}/{1}/{2}' -f (Get-Verdict $wrong 'localhost'), (Test-Note $wrong 'localhost' '期待する IP アドレスと違う（OS の結果）'), (Test-Note $wrong 'localhost' '期待する IP アドレスと違う（DNS サーバー 10.0.0.53）')) '注意/True/True'
    $extra = Invoke-Fake @{ Args = @{ Target = "localhost=$($v4[0])" }; Hosts = ''; Zone = @{ localhost = @{ '*' = @($v4[0], '10.2.2.2') } } }
    Assert-Equal '期待していない同じ種類の IP アドレスがあれば違う' (Test-Note $extra 'localhost' '期待する IP アドレスと違う（DNS サーバー') 'True'
    $hostsOkDnsOld = Invoke-Fake @{ Args = @{ Target = "localhost=$($v4[0])" }; Hosts = "$($v4[0]) localhost"; Zone = @{ localhost = @{ '*' = @('10.2.2.2') } } }
    Assert-Equal 'hosts が期待どおりでも DNS が古ければ注意' (Test-Note $hostsOkDnsOld 'localhost' '期待する IP アドレスと違う（DNS サーバー') 'True'

    Write-Section '逆引き'
    $revOk = Invoke-Fake @{ Args = @{ Target = 'localhost'; Reverse = $true; Server = '10.0.0.53' }; Hosts = ''; Zone = @{ localhost = @{ '*' = @('127.0.0.1') } }; Ptr = @{ '127.0.0.1' = @('LOCALHOST'); '::1' = @('localhost') } }
    Assert-Equal '一致（大文字と小文字は区別しない）' (Test-Contains (Get-Cell $revOk 'localhost' '逆引き') '127.0.0.1: 一致') 'True'
    $revBad = Invoke-Fake @{ Args = @{ Target = 'localhost'; Reverse = $true; Server = '10.0.0.53' }; Hosts = ''; Zone = @{ localhost = @{ '*' = @('127.0.0.1') } }; Ptr = @{ '127.0.0.1' = @('other.test'); '::1' = @('other.test') } }
    Assert-Equal '違うと注意' ('{0}/{1}' -f (Get-Verdict $revBad 'localhost'), (Test-Note $revBad 'localhost' '逆引きの名前が違う')) '注意/True'
    $revNone = Invoke-Fake @{ Args = @{ Target = '192.0.2.10'; Server = '10.0.0.53' }; Hosts = '' }
    Assert-Equal '登録がなければ判定を変えない' ('{0}/{1}' -f (Get-Verdict $revNone '192.0.2.10'), (Test-Note $revNone '192.0.2.10' '逆引きの登録がない')) 'OK/True'
    $ipRun = Invoke-Fake @{ Args = @{ Target = '192.0.2.10'; Server = '10.0.0.53' }; Hosts = ''; Ptr = @{ '192.0.2.10' = @('web01.test') } }
    Assert-Equal 'IP アドレスを書くと逆引きの名前を備考に書く' ('{0}/{1}' -f (Get-Cell $ipRun '192.0.2.10' '引けた元'), (Test-Note $ipRun '192.0.2.10' '逆引き: web01.test')) 'IP アドレス/True'
    $revDead = Invoke-Fake @{ Args = @{ Target = '192.0.2.10'; Server = '10.0.0.99' }; Hosts = '' }
    Assert-Equal '聞けなければ注意' ('{0}/{1}' -f (Get-Verdict $revDead '192.0.2.10'), (Test-Note $revDead '192.0.2.10' '逆引きを聞けなかった')) '注意/True'
    $revNext = Invoke-Fake @{ Args = @{ Target = '192.0.2.10'; Server = '10.0.0.99,10.0.0.53' }; Hosts = ''; Ptr = @{ '192.0.2.10' = @('web01.test') } }
    Assert-Equal '応答しなければ次のサーバーに聞く' (Test-Note $revNext '192.0.2.10' '逆引き: web01.test') 'True'
    $ptrDead = Invoke-Fake @{ Args = @{ Target = '192.0.2.1,192.0.2.2,192.0.2.3,192.0.2.4,localhost'; Server = '10.0.0.66' }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }
    Assert-Equal '逆引きだけ応答しないサーバーにも正引きは聞く' (Get-Cell $ptrDead 'localhost' '引けた元') 'DNS'

    Write-Section '-Hosts'
    $all = Invoke-Fake @{ Args = @{ Hosts = $true }; Hosts = "127.0.0.1 localhost`r`n::1 localhost`r`n127.0.0.1 loop.test`r`n10.3.3.3 forgot.test other.test`r`n" }
    Assert-Equal 'hosts の名前をすべて調べる（localhost と 127.0.0.1 の行を除く）' (($all.Rows | ForEach-Object { $_.'名前' }) -join ',') 'forgot.test,other.test'

    Write-Section '対象の CSV'
    $newList = "$workRoot\names.csv"
    $made = Invoke-Fake @{ Args = @{ NewList = $newList; NoSave = $true }; Hosts = '' }
    Assert-Equal 'ひな形を作る' ('{0}/{1}' -f $made.Code, (Test-Path -LiteralPath $newList)) '0/True'
    $again = Invoke-Fake @{ Args = @{ NewList = $newList; NoSave = $true }; Hosts = '' }
    Assert-Equal 'ひな形を上書きしない' ('{0}/{1}' -f $again.Code, (Test-Contains $again.Log '同じ名前のファイルがあります')) '1/True'
    $fromTemplate = Invoke-Fake @{ Args = @{ List = $newList }; Hosts = '' }
    Assert-Equal 'ひな形をそのまま読める' (@($fromTemplate.Rows).Count) 2

    Write-Utf8 "$workRoot\list.csv" "名前,期待するIP,メモ`r`n# コメント,,`r`nlocalhost,$($v4[0]),3月1日の切り替え`r`n,,`r`n"
    $listRun = Invoke-Fake @{ Args = @{ List = "$workRoot\list.csv" }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }
    Assert-Equal '見出しの空白が違っても期待する IP を読む' (Get-Cell $listRun 'localhost' '期待する IP') $v4[0]
    Assert-Equal 'メモはいつも ="..."' (Get-Row $listRun 'localhost').'メモ' '="3月1日の切り替え"'
    Assert-Equal '-Target の期待する IP を優先する' (Get-Cell (Invoke-Fake @{ Args = @{ List = "$workRoot\list.csv"; Target = 'localhost=10.7.7.7' }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }) 'localhost' '期待する IP') '10.7.7.7'

    $csvStops = @(
        @("名前,期待の IP`r`nlocalhost,1.2.3.4`r`n", '見出しを読めません'),
        @("ホスト`r`nlocalhost`r`n", '「名前」の列がありません'),
        @("名前,期待する IP`r`nlocalhost,abc`r`n", '2 行目: 期待する IP アドレスが正しくありません'),
        @("名前,期待する IP`r`nlocalhost,１０.０.０.１`r`n", '2 行目: 期待する IP アドレスが正しくありません'),
        @("名前,期待する IP`r`na,10.0.0.1`r`na,10.0.0.2`r`n", '3 行目: 同じ名前に別の期待する IP アドレス'),
        @("名前`r`n""abc`r`n", '引用符（"）が閉じていません')
    )
    foreach ($case in $csvStops) {
        Write-Utf8 "$workRoot\stop.csv" $case[0]
        $result = Invoke-Fake @{ Args = @{ List = "$workRoot\stop.csv"; NoSave = $true }; Hosts = '' }
        Assert-Equal "CSV の誤り: $($case[1])" ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $case[1])) '1/True'
    }
    Write-Utf8 "$workRoot\order.csv" "名前,期待する IP`r`nlocalhost,10.0.0.1;10.0.0.2`r`nlocalhost,10.0.0.2;10.0.0.1`r`n"
    Assert-Equal '順番だけ違う期待する IP は止めない' (@((Invoke-Fake @{ Args = @{ List = "$workRoot\order.csv" }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }).Rows).Count) 1

    Write-Section '設定ファイル'
    Write-Utf8 "$workRoot\conf.ini" "# コメント`r`nTarget = localhost=$($v4[0])`r`nTarget = 192.0.2.10`r`nServer = 10.0.0.53`r`nReverse = false`r`nNoSave = false`r`n"
    $confRun = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\conf.ini" }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }
    Assert-Equal '複数の行の Target' (@($confRun.Rows).Count) 2
    Assert-Equal 'Target の = の後ろを期待する IP として読む' (Get-Cell $confRun 'localhost' '期待する IP') $v4[0]
    Assert-Equal 'Server を読む' (Test-Contains $confRun.Log '10.0.0.53（指定）') 'True'
    Assert-Equal '設定ファイルの表示' (Test-Contains $confRun.Log "設定ファイル: $workRoot\conf.ini") 'True'
    Assert-Equal 'コマンドラインの -Target を優先する' (@((Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\conf.ini"; Target = 'localhost' }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }).Rows).Count) 1
    $confStops = @(
        @('Exlude = a', '知らない項目です: Exlude'),
        @('NewList = a.csv', 'NewList は設定ファイルには書けません'),
        @('Server = dns01', 'Server は IP アドレスで指定してください: dns01'),
        @('Target = a=abc', 'Target の期待する IP アドレスが正しくありません: abc'),
        @('Target = a=', 'Target の期待する IP アドレスが空です'),
        @('NoSave = yes', 'NoSave は true か false で指定してください: yes'),
        @("List = a.csv`r`nList = b.csv", 'List が 2 回書かれています'),
        @('Target = %NO_SUCH_VARIABLE_X%', '定義されていない環境変数です')
    )
    foreach ($case in $confStops) {
        Write-Utf8 "$workRoot\bad.ini" $case[0]
        $result = Invoke-Fake @{ Args = @{ ConfigFile = "$workRoot\bad.ini"; NoSave = $true }; Hosts = '' }
        Assert-Equal "書き間違い: $($case[1])" ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $case[1])) '1/True'
    }

    Write-Section '止まるとき'
    $stops = @(
        @('対象なし', @{ NoSave = $true }, '調べる名前を指定してください'),
        @('-Server の誤り', @{ Target = 'a'; Server = 'dns01'; NoSave = $true }, '-Server は IP アドレスで指定してください: dns01'),
        @('期待する IP の誤り', @{ Target = 'a=1.2.3'; NoSave = $true }, '-Target の期待する IP アドレスが正しくありません: 1.2.3'),
        @('= の後ろが空', @{ Target = 'a='; NoSave = $true }, '-Target の期待する IP アドレスが空です'),
        @('-Target の中の食い違い', @{ Target = 'a=10.0.0.1,a=10.0.0.2'; NoSave = $true }, '-Target で、同じ名前に別の期待する IP アドレス'),
        @('見つからない CSV', @{ List = "$workRoot\no_such.csv"; NoSave = $true }, '対象の CSV が見つかりません'),
        @('hosts に名前がない -Hosts', @{ Hosts = $true; NoSave = $true }, '調べる対象がありません')
    )
    foreach ($stop in $stops) {
        $result = Invoke-Fake @{ Args = $stop[1]; Hosts = '# 空' }
        Assert-Equal $stop[0] ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $stop[2])) '1/True'
    }

    Write-Section '本物の Resolve-DnsName'
    if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
        # 127.0.0.2 で DNS は動いていないので、応答しないサーバーとして分類できるかも確かめる
        $real = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Target "localhost,$bad" -Server 127.0.0.2 -NoSave 2>&1 | Out-String)
        Assert-Equal '止まらずに動く' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $real 'エラーで止まりました')) '2/False'
        Assert-Equal '応答しないサーバーを見分ける' (Test-Contains $real 'DNS サーバー 127.0.0.2') 'True'
        $realOs = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Target "localhost,$selfName" -NoSave 2>&1 | Out-String)
        Assert-Equal 'OS の設定の DNS サーバーで動く' ('{0}/{1}' -f (Test-Contains $realOs 'エラーで止まりました'), (Test-Contains $realOs '結果: OK')) 'False/True'
    } else {
        Write-Skip '本物の Resolve-DnsName' 'このPCにないため'
    }

    Write-Section '起動のしかた'
    # コマンドプロンプトで "C:\data\" のように書くと、\" が引用符のまま届く
    $bat = "$workRoot\quote.cmd"
    Write-Utf8 $bat ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -List ""{2}\"" -NoSave`r`nexit /b %errorlevel%`r`n" -f $Exe, $tool, $workRoot)
    $quoteLog = ConvertTo-FlatText (& cmd.exe /c $bat 2>&1 | Out-String)
    Assert-Equal '最後が \" の引数' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $quoteLog '引用符（"）が残っています')) '1/True'

    # 名前に [ ] を含むフォルダから、相対パスで起動する
    Write-Utf8 "$workRoot\[br]\a.csv" "名前`r`nlocalhost`r`n"
    $brBat = "$workRoot\bracket.cmd"
    Write-Utf8 $brBat ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\[br]""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -List a.csv -NoSave`r`nexit /b %errorlevel%`r`n" -f $workRoot, $Exe, $tool)
    $brLog = ConvertTo-FlatText (& cmd.exe /c $brBat 2>&1 | Out-String)
    Assert-Equal '[ ] を含むフォルダから起動する' (Test-Contains $brLog '対象 1 件') 'True'

    # 保存先に書き込めないときは、画面表示だけで終わる（保存先がファイル）
    Write-Utf8 "$workRoot\notdir" 'x'
    $unwritable = Invoke-Fake @{ Args = @{ Target = 'localhost'; OutputFolder = "$workRoot\notdir"; NoSave = $false }; Hosts = ''; Zone = @{ localhost = @{ '*' = $localIps } } }
    Assert-Equal '書き込めない保存先' ('{0}/{1}' -f $unwritable.Code, (Test-Contains $unwritable.Log '保存先に書き込めないため')) '0/True'
} finally {
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
