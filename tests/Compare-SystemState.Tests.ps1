<#
.SYNOPSIS
    Compare-SystemState の回帰テストです。

.DESCRIPTION
    サーバーの状態を自由に変えられないため、Get-CimInstance などと同じ名前の偽の関数を定義したラッパーから
    Compare-SystemState を実行し、比べ方を確かめます。hosts は環境変数 SystemRoot を一時フォルダに向けて読ませます。
    最後に、偽の関数を使わずに本物の状態でも保存と比較を行い、続けて 2 回取っても誤検知が出ないことを確かめます。
    ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Compare-SystemState を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Compare-SystemState\Compare-SystemState.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Compare-SystemState_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
# 結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーしたスクリプトを動かす
$tool     = Join-Path $workRoot 'tool\Compare-SystemState.ps1'
$wrapper  = Join-Path $workRoot 'wrapper.ps1'
$fakeRoot = Join-Path $workRoot 'root'
$fakeItems = 'Service,Port,HotFix,Network,Route,Share,LocalUser,LocalGroup,Hosts,Firewall'
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

function Get-Clean($Value) { ([string]$Value) -replace '^="', '' -replace '"$', '' }

# 偽の環境で動かす。Phase は before か after。Fail はわざと失敗させるクラスやコマンドの名前。Args はツールの引数
function Invoke-Fake([hashtable]$ToolArgs, [string]$Phase, [string]$Fail, [string]$Hosts) {
    $script:runNumber++
    $out = Join-Path $workRoot ('out_{0}' -f $script:runNumber)
    if (-not $ToolArgs.ContainsKey('NoSave') -and -not $ToolArgs.ContainsKey('OutputFolder')) { $ToolArgs['OutputFolder'] = $out }
    if ($Hosts) { Write-Text (Join-Path $fakeRoot 'System32\drivers\etc\hosts') $Hosts }
    $argFile = Join-Path $workRoot ('args_{0}.json' -f $script:runNumber)
    [IO.File]::WriteAllText($argFile, ($ToolArgs | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding $true))
    $env:FAKE_PHASE = $Phase
    $env:FAKE_FAIL  = $Fail
    $log  = & $Exe -NoProfile -ExecutionPolicy Bypass -File $wrapper -Tool $tool -ArgFile $argFile -Root $fakeRoot 2>&1 | Out-String
    $code = $LASTEXITCODE
    $env:FAKE_PHASE = $null
    $env:FAKE_FAIL  = $null
    $dir = $null
    if (Test-Path -LiteralPath $out) { $dir = @(Get-ChildItem -LiteralPath $out -Directory | ForEach-Object { $_.FullName })[0] }
    $rows = @()
    $skips = @()
    if ($dir -and (Test-Path -LiteralPath (Join-Path $dir 'differences.csv'))) { $rows = @(Import-Csv -LiteralPath (Join-Path $dir 'differences.csv') -Encoding UTF8) }
    if ($dir -and (Test-Path -LiteralPath (Join-Path $dir 'skipped.csv'))) { $skips = @(Import-Csv -LiteralPath (Join-Path $dir 'skipped.csv') -Encoding UTF8) }
    [pscustomobject]@{ Code = $code; Log = (ConvertTo-FlatText $log); Dir = $dir; Rows = $rows; Skips = $skips }
}

function Get-DiffRow($Run, [string]$Label, [string]$Target) {
    @($Run.Rows | Where-Object { $_.'項目' -eq $Label -and (Get-Clean $_.'対象') -eq $Target })
}

# 行がなくても、例外で Assert-Equal を飛ばさず、FAIL として数えられるようにする
function Get-DiffCell($Run, [string]$Label, [string]$Target, [string]$Column) {
    $found = @(Get-DiffRow $Run $Label $Target)
    if ($found.Count -eq 0) { return '（行なし）' }
    return [string]$found[0].$Column
}

function Get-SkipCell($Run, [string]$Column) {
    if (@($Run.Skips).Count -eq 0) { return '（行なし）' }
    return [string]@($Run.Skips)[0].$Column
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

try {
    #region ダミー環境

    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool
    [void][IO.Directory]::CreateDirectory((Join-Path $fakeRoot 'System32\drivers\etc'))

    $wrapperCode = @'
param([string]$Tool, [string]$ArgFile, [string]$Root)
$toolArgs = @{}
foreach ($property in ([IO.File]::ReadAllText($ArgFile) | ConvertFrom-Json).PSObject.Properties) { $toolArgs[$property.Name] = $property.Value }
$env:SystemRoot = $Root
function global:Test-After { $env:FAKE_PHASE -eq 'after' }
function global:Assert-Fake([string]$Name) { if ($env:FAKE_FAIL -and ($env:FAKE_FAIL -split ',') -contains $Name) { throw "偽の失敗: $Name" } }
function global:Get-CimInstance {
    param($ClassName, $Filter, $ErrorAction)
    Assert-Fake $ClassName
    $after = Test-After
    switch ($ClassName) {
        'Win32_Service' {
            [pscustomobject]@{ Name = 'W3SVC'; DisplayName = 'WWW'; State = $(if ($after) { 'Stopped' } else { 'Running' }); StartMode = 'Auto'; StartName = 'LocalSystem'; PathName = 'C:\w3.exe'; ProcessId = 100; DelayedAutoStart = $false }
            [pscustomobject]@{ Name = 'BITS'; DisplayName = 'BITS'; State = $(if ($after) { 'Running' } else { 'Stopped' }); StartMode = 'Manual'; StartName = 'LocalSystem'; PathName = 'C:\b.exe'; ProcessId = 0 }
            [pscustomobject]@{ Name = 'Old'; DisplayName = 'Old'; State = 'Running'; StartMode = 'Auto'; StartName = 'x'; PathName = $(if ($after) { 'C:\new.exe -password SECRETA' } else { 'C:\old.exe -password SECRETB' }); ProcessId = 0 }
            [pscustomobject]@{ Name = 'Gone'; DisplayName = 'Gone'; State = 'Running'; StartMode = 'Auto'; StartName = 'x'; PathName = 'C:\g.exe'; ProcessId = 0 } | Where-Object { -not $after }
            if ($after) { [pscustomobject]@{ Name = 'Tomcat9'; DisplayName = 'Tomcat'; State = 'Running'; StartMode = 'Auto'; StartName = 'LocalService'; PathName = 'C:\t.exe'; ProcessId = 200 } }
        }
        'Win32_Process' { [pscustomobject]@{ ProcessId = 100; Name = 'svchost.exe' }; [pscustomobject]@{ ProcessId = 200; Name = 'java.exe' } }
        'Win32_QuickFixEngineering' { [pscustomobject]@{ HotFixID = 'KB1'; Description = 'Update'; InstalledOn = '11/1/2026' }; if ($after) { [pscustomobject]@{ HotFixID = 'KB2'; Description = 'Update'; InstalledOn = '11/3/2026' } } }
        'Win32_OperatingSystem' { [pscustomobject]@{ Version = '10.0.20348'; LastBootUpTime = $(if ($after) { [datetime]'2026-11-03 23:05:00' } else { [datetime]'2026-10-28 08:01:00' }) } }
        'Win32_Share' { [pscustomobject]@{ Name = 'Keys'; Path = 'C:\k'; Description = '' }; [pscustomobject]@{ Name = 'Count'; Path = 'C:\c'; Description = '' }; [pscustomobject]@{ Name = '2026-10-31'; Path = 'C:\d'; Description = '' } }
        'Win32_ComputerSystem' { [pscustomobject]@{ DomainRole = $(if ($env:FAKE_FAIL -eq 'DC') { 5 } else { 3 }) } }
        'Win32_UserAccount' { [pscustomobject]@{ Name = 'admin'; Disabled = $after; PasswordExpires = $false; Description = 'd' } }
        'Win32_Group' { [pscustomobject]@{ Name = $(if ($env:FAKE_FAIL -eq 'BadGroup') { 'NoSuchGroupForTest' } else { 'Users' }) } }
        'Win32_NetworkAdapter' { [pscustomobject]@{ Index = 1; NetConnectionID = 'イーサネット' } }
        'Win32_NetworkAdapterConfiguration' {
            [pscustomobject]@{ Index = 1; Description = 'Intel'; MACAddress = 'AA:BB'; DHCPEnabled = $false; IPAddress = $(if ($after) { @('10.0.0.5', 'fe80::1') } else { @('fe80::1', '10.0.0.5') }); IPSubnet = $(if ($after) { @('255.255.255.0', '64') } else { @('64', '255.255.255.0') })
                DefaultIPGateway = @('10.0.0.1'); DNSServerSearchOrder = $(if ($after) { @('10.0.0.3', '10.0.0.2') } else { @('10.0.0.2', '10.0.0.3') }) }
        }
        default { throw "知らないクラス: $ClassName" }
    }
}
function global:Get-NetTCPConnection {
    param($ErrorAction)
    Assert-Fake 'Tcp'
    [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 443; OwningProcess = 100; State = 'Listen' }
    [pscustomobject]@{ LocalAddress = '10.0.0.5'; LocalPort = 50000; OwningProcess = 100; State = 'Established' }
    if (Test-After) {
        [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 8080; OwningProcess = 200; State = 'Listen' }
        [pscustomobject]@{ LocalAddress = '::'; LocalPort = 8080; OwningProcess = 200; State = 'Listen' }
        [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 49700; OwningProcess = 100; State = 'Listen' }
    }
}
function global:Get-NetUDPEndpoint { param($ErrorAction) [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 123; OwningProcess = 100 }; [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 123; OwningProcess = 200 } }
function global:Get-NetRoute {
    param($PolicyStore, $ErrorAction)
    foreach ($prefix in '0.0.0.0/0', '224.0.0.0/4', 'ff00::/8', 'fe80::/64', '127.0.0.0/8', '169.254.0.0/16', '::1/128', '255.255.255.255/32', '10.0.0.0/24') {
        [pscustomobject]@{ DestinationPrefix = $prefix; NextHop = '10.0.0.1'; InterfaceAlias = 'イーサネット'; RouteMetric = $(if (Test-After) { 25 } else { 0 }) }
    }
}
function global:Get-NetFirewallProfile { param($PolicyStore, $ErrorAction) [pscustomobject]@{ Name = 'Domain'; Enabled = 'True' } }
function global:Get-NetFirewallPortFilter { param($PolicyStore, $ErrorAction) [pscustomobject]@{ InstanceID = 'R1'; Protocol = 'TCP'; LocalPort = @('443', '80'); RemotePort = 'Any' } }
function global:Get-NetFirewallApplicationFilter { param($PolicyStore, $ErrorAction) [pscustomobject]@{ InstanceID = 'R1'; Program = 'Any' } }
function global:Get-NetFirewallAddressFilter { param($PolicyStore, $ErrorAction) [pscustomobject]@{ InstanceID = 'R1'; RemoteAddress = 'Any' } }
function global:Get-NetFirewallRule { param($PolicyStore, $ErrorAction) [pscustomobject]@{ InstanceID = 'R1'; DisplayName = 'Web'; Enabled = $(if (Test-After) { 'False' } else { 'True' }); Direction = 'Inbound'; Action = 'Allow'; Profile = 'Any' } }
& $Tool @toolArgs
exit $LASTEXITCODE
'@
    Write-Text $wrapper $wrapperCode

    #endregion

    #region 保存

    Write-Section 'スナップショットの保存'
    $before = Invoke-Fake @{ Snapshot = $true; Item = $fakeItems } 'before' '' "127.0.0.1 localhost # コメント`r`n# 行のコメント`r`n10.0.0.9 db db.local`r`n"
    Assert-Equal '終了コード' $before.Code '0'
    Assert-Equal 'フォルダ名' ((Split-Path $before.Dir -Leaf) -like 'SystemSnapshot_*') 'True'
    Assert-Equal '比べるコマンドを表示' (Test-Contains $before.Log '-Reference') 'True'
    foreach ($name in 'info.txt', 'services.csv', 'ports.csv', 'hotfixes.csv', 'network.csv', 'routes.csv', 'shares.csv', 'localusers.csv', 'localgroups.csv', 'hosts.csv', 'firewall.csv', 'skipped.csv') {
        $path = Join-Path $before.Dir $name
        if (Test-Path -LiteralPath $path) {
            $bytes = [IO.File]::ReadAllBytes($path)
            Assert-Equal ("$name は BOM 付き") ('{0}-{1}-{2}' -f $bytes[0], $bytes[1], $bytes[2]) '239-187-191'
        } else {
            Assert-Equal ("$name がある") 'False' 'True'
        }
    }
    Assert-Equal '指定しなかった項目の CSV は作らない' (Test-Path -LiteralPath (Join-Path $before.Dir 'tasks.csv')) 'False'
    Assert-Equal 'services.csv の列' ([IO.File]::ReadAllLines((Join-Path $before.Dir 'services.csv'))[0]) '"サービス名","表示名","状態","スタートアップ","実行ユーザー","実行ファイル"'
    $routes = @(Import-Csv -LiteralPath (Join-Path $before.Dir 'routes.csv') -Encoding UTF8)
    Assert-Equal 'マルチキャストなどのルートは取らない' $routes.Count '2'
    $hostsRows = @(Import-Csv -LiteralPath (Join-Path $before.Dir 'hosts.csv') -Encoding UTF8)
    Assert-Equal 'hosts はコメントを除き、組ごとに 1 行' $hostsRows.Count '3'
    $ports = @(Import-Csv -LiteralPath (Join-Path $before.Dir 'ports.csv') -Encoding UTF8)
    Assert-Equal '接続中の TCP は取らない' (@($ports | Where-Object { (Get-Clean $_.'ポート') -eq '50000' }).Count) '0'
    $udp = @($ports | Where-Object { (Get-Clean $_.'ポート') -eq '123' })
    Assert-Equal '同じポートのプロセスは 1 行にまとめる' (($udp | ForEach-Object { Get-Clean $_.'プロセス名' }) -join '/') 'java.exe, svchost.exe'
    $shareRows = @(Import-Csv -LiteralPath (Join-Path $before.Dir 'shares.csv') -Encoding UTF8)
    Assert-Equal 'Excel で変わる名前は ="..."' (@($shareRows | Where-Object { $_.'共有名' -eq '="2026-10-31"' }).Count) '1'

    #endregion

    #region 比較

    Write-Section '同じ状態を比べる'
    $same = Invoke-Fake @{ Reference = $before.Dir; Item = 'Service,Port,HotFix,Share,Hosts,Route,Network,Firewall,LocalUser' } 'before' ''
    Assert-Equal '終了コード 0' $same.Code '0'
    Assert-Equal '結果の行' (Test-Contains $same.Log '結果: 違いなし') 'True'
    Assert-Equal '再起動なし' (Test-Contains $same.Log '再起動: なし') 'True'
    Assert-Equal 'Keys という名前の共有があっても数えられる' (Test-Contains $same.Log '共有フォルダ') 'True'

    Write-Section '作業の後と比べる'
    $after = Invoke-Fake @{ Reference = $before.Dir; Item = $fakeItems } 'after' ''
    Assert-Equal '終了コード 2' $after.Code '2'
    Assert-Equal '結果の行' (Test-Contains $after.Log '結果: 違いあり') 'True'
    Assert-Equal '再起動あり' (Test-Contains $after.Log '再起動: あり（基準 2026-10-28 08:01:00 → 比較 2026-11-03 23:05:00）') 'True'
    foreach ($name in 'summary.txt', 'differences.csv', 'skipped.csv', 'snapshot\info.txt') {
        Assert-Equal ("$name がある") (Test-Path -LiteralPath (Join-Path $after.Dir $name)) 'True'
    }
    Assert-Equal 'differences.csv の列' ([IO.File]::ReadAllLines((Join-Path $after.Dir 'differences.csv'))[0]) '"結果","項目","対象","属性","基準の値","比較の値","備考"'
    Assert-Equal 'サービスの追加' (Get-DiffCell $after 'サービス' 'Tomcat9' '結果') '追加'
    Assert-Equal '追加の値' (Get-DiffCell $after 'サービス' 'Tomcat9' '比較の値') 'スタートアップ=自動 / 状態=実行中 / 実行ユーザー=LocalService / 実行ファイル=C:\t.exe'
    Assert-Equal 'サービスの削除' (Get-DiffCell $after 'サービス' 'Gone' '結果') '削除'
    Assert-Equal '自動のサービスの停止は変更' ('{0}/{1}/{2}/{3}' -f (Get-DiffCell $after 'サービス' 'W3SVC' '結果'), (Get-DiffCell $after 'サービス' 'W3SVC' '属性'), (Get-DiffCell $after 'サービス' 'W3SVC' '基準の値'), (Get-DiffCell $after 'サービス' 'W3SVC' '比較の値')) '変更/状態/実行中/停止'
    Assert-Equal '手動のサービスの状態だけは参考' ('{0}/{1}' -f (Get-DiffCell $after 'サービス' 'BITS' '結果'), (Get-DiffCell $after 'サービス' 'BITS' '備考')) '参考/手動開始のサービスの状態の変化'
    Assert-Equal '参考は画面に出さない' (Test-Contains $after.Log '手動開始のサービスの状態の変化') 'False'
    Assert-Equal '動的ポートは参考' (Get-DiffCell $after '待ち受けポート' 'TCP 0.0.0.0:49700' '結果') '参考'
    Assert-Equal 'ポートの追加' (Get-DiffCell $after '待ち受けポート' 'TCP 0.0.0.0:8080' '結果') '追加'
    Assert-Equal 'IPv6 のアドレスは [ ] で囲む' (Get-DiffCell $after '待ち受けポート' 'TCP [::]:8080' '結果') '追加'
    Assert-Equal '同じプロセスでサービスの同居が変わっても違いにしない' @(Get-DiffRow $after '待ち受けポート' 'UDP 0.0.0.0:123').Count '0'
    Assert-Equal '更新プログラムの追加' (Get-DiffCell $after '更新プログラム' 'KB2' '結果') '追加'
    Assert-Equal 'メトリックの変更' (Get-DiffCell $after 'ルーティング' '0.0.0.0/0 → 10.0.0.1（イーサネット）' '属性') 'メトリック'
    Assert-Equal 'ユーザーの無効化' (Get-DiffCell $after 'ローカルユーザー' 'admin' '比較の値') '無効'
    Assert-Equal 'ファイアウォールの対象は表示名（ID）' (Get-DiffCell $after 'ファイアウォール' 'Web（R1）' '属性') '有効'
    $net = @(Get-DiffRow $after 'ネットワーク設定' 'イーサネット')
    Assert-Equal 'IP の順序の違いは違いにせず、DNS の順序の違いは違いにする' (($net | ForEach-Object { $_.'属性' }) -join ',') 'DNS サーバー'

    Write-Section 'パスワードが含まれることのある値'
    Assert-Equal '画面に値を出さない' ((Test-Contains $after.Log 'SECRETA') -or (Test-Contains $after.Log 'SECRETB')) 'False'
    Assert-Equal '変わったことは出す' (Test-Contains $after.Log '実行ファイルが変わりました') 'True'
    $summary = [IO.File]::ReadAllText((Join-Path $after.Dir 'summary.txt'))
    Assert-Equal 'summary.txt にも値を書かない' (Test-Contains $summary 'SECRETA') 'False'
    Assert-Equal 'CSV には書く' ((Get-DiffCell $after 'サービス' 'Old' '比較の値')) 'C:\new.exe -password SECRETA'

    Write-Section '保存したもの同士'
    $two = Invoke-Fake @{ Reference = $before.Dir; Difference = (Join-Path $after.Dir 'snapshot'); Item = $fakeItems } 'before' ''
    Assert-Equal '終了コード 2' $two.Code '2'
    Assert-Equal '追加が同じ' (Get-DiffCell $two 'サービス' 'Tomcat9' '結果') '追加'
    Assert-Equal 'snapshot フォルダを作らない' (Test-Path -LiteralPath (Join-Path $two.Dir 'snapshot')) 'False'
    Assert-Equal 'Excel で変わる名前を元に戻して突き合わせる' @(Get-DiffRow $two '共有フォルダ' '2026-10-31').Count '0'

    Write-Section '-Item の絞り込み'
    $only = Invoke-Fake @{ Reference = $before.Dir; Item = 'hotfix' } 'after' ''
    Assert-Equal '大文字と小文字を区別しない' $only.Rows.Count '1'
    $all = Invoke-Fake @{ Reference = (Invoke-Fake @{ Snapshot = $true; Item = 'HotFix,Share' } 'before' '').Dir } 'after' ''
    Assert-Equal '省略すると基準の項目を比べる' $all.Rows.Count '1'

    #endregion

    #region 比べられないとき

    Write-Section '項目が取れないとき（レビューの指摘）'
    $tcp = Invoke-Fake @{ Reference = $before.Dir; Item = 'Port,HotFix' } 'before' 'Tcp'
    Assert-Equal 'TCP が取れないと、すべて削除に見せない' @(Get-DiffRow $tcp '待ち受けポート' 'TCP 0.0.0.0:443').Count '0'
    Assert-Equal '比べられないので終了コード 1' $tcp.Code '1'
    Assert-Equal '結果の行に書く' (Test-Contains $tcp.Log '比べられなかった項目が 1 件') 'True'
    Assert-Equal 'skipped.csv に残す' (Get-SkipCell $tcp '項目') '待ち受けポート'

    $noSave = Invoke-Fake @{ Reference = $before.Dir; Item = 'Port,HotFix'; NoSave = $true } 'before' 'Tcp'
    Assert-Equal '-NoSave のときは理由を画面に出す' (Test-Contains $noSave.Log '■ 比べられなかったもの') 'True'
    Assert-Equal '-NoSave のときは skipped.csv を案内しない' (Test-Contains $noSave.Log 'skipped.csv') 'False'

    Write-Section '一部を読めないとき（レビューの指摘）'
    $partSnap = Invoke-Fake @{ Snapshot = $true; Item = 'LocalGroup' } 'before' 'BadGroup'
    Assert-Equal '止まらない' $partSnap.Code '0'
    Add-Content -LiteralPath (Join-Path $partSnap.Dir 'localgroups.csv') '"Users","HOST\gone"' -Encoding UTF8
    $part = Invoke-Fake @{ Reference = $partSnap.Dir } 'before' 'BadGroup'
    Assert-Equal '読めなかった側の削除は参考' (Get-DiffCell $part 'ローカルグループのメンバー' 'Users: HOST\gone' '結果') '参考'
    Assert-Equal '違いに数えず終了コード 1' $part.Code '1'
    Assert-Equal '読めなかった対象の件数' (Test-Contains $part.Log '読めなかった対象が 2 件') 'True'

    Write-Section 'ドメインコントローラー'
    $dcSnap = Invoke-Fake @{ Snapshot = $true; Item = 'LocalUser,HotFix' } 'before' 'DC'
    Assert-Equal '対象外と表示' (Test-Contains $dcSnap.Log '対象外') 'True'
    $dc = Invoke-Fake @{ Reference = $dcSnap.Dir } 'before' 'DC'
    Assert-Equal '両方とも対象外なら、比べられないに数えない' $dc.Code '0'

    Write-Section '基準と比較の違い'
    $copy = Join-Path $workRoot 'renamed'
    Copy-Item -LiteralPath (Join-Path $after.Dir 'snapshot') -Destination $copy -Recurse
    $infoPath = Join-Path $copy 'info.txt'
    $infoText = [IO.File]::ReadAllText($infoPath)
    $infoOther = [regex]::Replace($infoText, 'ホスト名 = .*', 'ホスト名 = OTHERHOST')
    $infoOther = [regex]::Replace($infoOther, '管理者 = .*', $(if ($isAdmin) { '管理者 = いいえ' } else { '管理者 = はい' }))
    [IO.File]::WriteAllText($infoPath, $infoOther, (New-Object Text.UTF8Encoding $true))
    $other = Invoke-Fake @{ Reference = $before.Dir; Difference = $copy; Item = 'HotFix' } 'before' ''
    Assert-Equal '違うコンピューターの案内' (Test-Contains $other.Log '違うコンピューターの状態を比べています') 'True'
    Assert-Equal '管理者の違いの案内' (Test-Contains $other.Log '管理者かどうかが違います') 'True'
    Assert-Equal '違うコンピューターでは再起動の行を出さない' (Test-Contains $other.Log '再起動:') 'False'

    $scope = [regex]::Replace($infoOther, 'タスクの範囲 = .*', 'タスクの範囲 = すべて')
    $scope += "項目.Task = 取得`r`n"
    [IO.File]::WriteAllText($infoPath, $scope, (New-Object Text.UTF8Encoding $true))
    Write-Text (Join-Path $copy 'tasks.csv') '"パス","有効","実行ユーザー","トリガー","操作"'
    $refScope = Join-Path $workRoot 'refscope'
    Copy-Item -LiteralPath $before.Dir -Destination $refScope -Recurse
    Add-Content -LiteralPath (Join-Path $refScope 'info.txt') "項目.Task = 取得" -Encoding UTF8
    Write-Text (Join-Path $refScope 'tasks.csv') '"パス","有効","実行ユーザー","トリガー","操作"'
    $scopeRun = Invoke-Fake @{ Reference = $refScope; Difference = $copy; Item = 'Task' } 'before' ''
    Assert-Equal 'タスクの範囲が違うと比べない' (Test-Contains $scopeRun.Log '比べられなかった項目が 1 件') 'True'

    Write-Section '壊れたスナップショット'
    $broken = Join-Path $workRoot 'broken'
    Copy-Item -LiteralPath $before.Dir -Destination $broken -Recurse
    [IO.File]::WriteAllText((Join-Path $broken 'services.csv'), 'a,b,c')
    $badCsv = Invoke-Fake @{ Reference = $broken; Item = 'Service,HotFix' } 'before' ''
    Assert-Equal '1 つの CSV が読めなくても止まらない' (@($badCsv.Skips | Where-Object { $_.'理由' -eq '読めない' }).Count) '1'
    Assert-Equal '読めない項目は比べない' (@($badCsv.Rows | Where-Object { $_.'項目' -eq 'サービス' }).Count) '0'
    $noItems = Join-Path $workRoot 'noitems'
    [void][IO.Directory]::CreateDirectory($noItems)
    Write-Text (Join-Path $noItems 'info.txt') "ツール = Compare-SystemState`r`n保存日時 = 2026/11/03 22:00:00`r`n"
    $empty = Invoke-Fake @{ Reference = $noItems; NoSave = $true } 'before' ''
    Assert-Equal '項目のないスナップショットは止める' ('{0}/{1}' -f $empty.Code, (Test-Contains $empty.Log 'スナップショットのフォルダとして読めません')) '1/True'

    #endregion

    #region 止まるとき

    Write-Section '止まるとき'
    $cases = @(
        @{ Name = '何も指定しない'; Args = @{}; Text = '作業の前は -Snapshot を' }
        @{ Name = '-Snapshot と -Reference'; Args = @{ Snapshot = $true; Reference = 'x' }; Text = '-Snapshot のときは、-Reference は指定できません' }
        @{ Name = '-Snapshot と -Difference'; Args = @{ Snapshot = $true; Difference = 'x' }; Text = '-Snapshot のときは、-Difference は指定できません' }
        @{ Name = '-Snapshot と -NoSave'; Args = @{ Snapshot = $true; NoSave = $true }; Text = '-Snapshot のときは、-NoSave は指定できません' }
        @{ Name = '-AllTasks だけ'; Args = @{ AllTasks = $true; Reference = $before.Dir }; Text = '-AllTasks は、-Snapshot のときだけ指定できます' }
        @{ Name = '-Difference だけ'; Args = @{ Difference = $before.Dir }; Text = '-Difference を使うときは、-Reference も指定してください' }
        @{ Name = '見つからないフォルダ'; Args = @{ Reference = "$workRoot\nothing" }; Text = '基準のフォルダを開けません' }
        @{ Name = 'スナップショットでないフォルダ'; Args = @{ Reference = $workRoot }; Text = 'スナップショットのフォルダとして読めません' }
        @{ Name = '基準と比較が同じ'; Args = @{ Reference = $before.Dir; Difference = $before.Dir }; Text = '基準と比較が同じフォルダです' }
        @{ Name = '-Item の誤り'; Args = @{ Reference = $before.Dir; Item = 'Servce' }; Text = '-Item の値が正しくありません: Servce' }
        @{ Name = '基準にない項目'; Args = @{ Reference = $before.Dir; Item = 'Task' }; Text = '基準のスナップショットに、次の項目がありません: Task' }
    )
    foreach ($case in $cases) {
        $case.Args['NoSave'] = $case.Args.ContainsKey('NoSave') -and $case.Args['NoSave']
        if (-not $case.Args['NoSave']) { $case.Args.Remove('NoSave'); $case.Args['OutputFolder'] = "$workRoot\unused" }
        $run = Invoke-Fake $case.Args 'before' ''
        Assert-Equal ($case.Name + ': 止まる') ('{0}/{1}' -f $run.Code, (Test-Contains $run.Log $case.Text)) '1/True'
    }

    $cmdFile = "$workRoot\quote.cmd"
    Write-Text $cmdFile ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -Reference ""{2}\"" -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $Exe, $tool, $before.Dir)
    $quote = ConvertTo-FlatText (& cmd.exe /c $cmdFile 2>&1 | Out-String)
    Assert-Equal '\" で引数を巻き込む: 終了コード' $LASTEXITCODE '1'
    Assert-Equal '\" で引数を巻き込む: ヒント' (Test-Contains $quote '引用符') 'True'

    if ($hostTag -eq '51') {
        $env:__PSLockdownPolicy = '4'
        $clm = ConvertTo-FlatText (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool -Snapshot 2>&1 | Out-String)
        $clmCode = $LASTEXITCODE
        $env:__PSLockdownPolicy = $null
        Assert-Equal '制約付き言語モード: 止まる' ('{0}/{1}' -f $clmCode, (Test-Contains $clm '制約付き言語モード')) '1/True'
    } else {
        Write-Skip '制約付き言語モード' '5.1 だけで確かめます'
    }

    Write-Text "$workRoot\notdir" 'x'
    $unwritable = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Snapshot -Item Hosts -OutputFolder "$workRoot\notdir" 2>&1 | Out-String
    Assert-Equal '保存できないスナップショットは止める' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains (ConvertTo-FlatText $unwritable) 'スナップショットを保存できません')) '1/True'
    $writeDiff = Invoke-Fake @{ Reference = $before.Dir; Item = 'HotFix'; OutputFolder = "$workRoot\notdir" } 'before' ''
    Assert-Equal '比べるときは警告して画面表示だけ' ('{0}/{1}' -f $writeDiff.Code, (Test-Contains $writeDiff.Log '保存先に書き込めないため')) '0/True'

    #endregion

    #region 本物の状態

    Write-Section '本物の状態で、続けて 2 回取る'
    $realOut = Join-Path $workRoot 'real'
    $realLog = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Snapshot -OutputFolder $realOut 2>&1 | Out-String
    Assert-Equal '保存の終了コード' $LASTEXITCODE '0'
    $realSnap = @(Get-ChildItem -LiteralPath $realOut -Directory | ForEach-Object { $_.FullName })[0]
    $realInfo = [IO.File]::ReadAllText((Join-Path $realSnap 'info.txt'))
    foreach ($name in 'Service', 'Port', 'Task', 'Software', 'HotFix', 'Network', 'Route', 'Share', 'Hosts', 'Environment', 'Firewall') {
        Assert-Equal ("$name を取れる") (Test-Contains $realInfo ('項目.{0} = 取得' -f $name)) 'True'
    }
    foreach ($name in 'LocalUser', 'LocalGroup') {
        $got = (Test-Contains $realInfo ('項目.{0} = 取得' -f $name)) -or (Test-Contains $realInfo ('項目.{0} = 対象外' -f $name))
        Assert-Equal ("$name を取れる") $got 'True'
    }
    Assert-Equal 'サービスが 1 件以上' (@(Import-Csv -LiteralPath (Join-Path $realSnap 'services.csv') -Encoding UTF8).Count -gt 0) 'True'
    Assert-Equal 'ファイアウォールのフィルターが規則と結び付く' (@(Import-Csv -LiteralPath (Join-Path $realSnap 'firewall.csv') -Encoding UTF8 | Where-Object { $_.'種類' -eq '規則' -and $_.'プロトコル' }).Count -gt 0) 'True'
    $realDiffOut = Join-Path $workRoot 'realdiff'
    $realDiffLog = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Reference $realSnap -OutputFolder $realDiffOut 2>&1 | Out-String
    $realDir = @(Get-ChildItem -LiteralPath $realDiffOut -Directory | ForEach-Object { $_.FullName })[0]
    $realRows = @(Import-Csv -LiteralPath (Join-Path $realDir 'differences.csv') -Encoding UTF8 | Where-Object { $_.'結果' -ne '参考' })
    foreach ($row in $realRows) { Write-Section ('本物の状態の違い: {0} / {1} / {2} / {3} → {4}' -f $row.'項目', (Get-Clean $row.'対象'), $row.'属性', $row.'基準の値', $row.'比較の値') }
    $stable = @($realRows | Where-Object { $_.'項目' -notin 'サービス', '待ち受けポート' })
    Assert-Equal '設定の項目に誤検知が出ない（サービスとポートを除く）' $stable.Count '0'
    Assert-Equal '再起動なし' (Test-Contains (ConvertTo-FlatText $realDiffLog) '再起動: なし') 'True'

    # 名前に [ ] を含むフォルダから、相対パスで起動する
    [void][IO.Directory]::CreateDirectory("$workRoot\[br]")
    $brBat = "$workRoot\bracket.cmd"
    Write-Text $brBat ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\[br]""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -Snapshot -Item Hosts -OutputFolder res`r`nexit /b %ERRORLEVEL%`r`n" -f $workRoot, $Exe, $tool)
    [void](& cmd.exe /c $brBat 2>&1 | Out-String)
    Assert-Equal '[ ] を含むフォルダから起動する' (Test-Path -LiteralPath "$workRoot\[br]\res") 'True'

    if (-not $isAdmin) { Write-Skip '管理者でのテスト' '管理者として実行すると、取れる情報が増えます' }

    #endregion
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
