<#
.SYNOPSIS
    Compare-SystemState の回帰テストです。

.DESCRIPTION
    サーバーの状態を自由に変えられないため、Get-CimInstance などと同じ名前の偽の関数を定義したラッパーから
    Compare-SystemState のコピーを実行し、比べ方を確かめます。タスクは New-Object -ComObject Schedule.Service を偽の
    スケジューラーに差し替えます。hosts の場所と、インストール済みソフトと環境変数のレジストリは、ツールのコピーを
    書き換えて、一時フォルダと、テスト専用のレジストリのキー（HKCU\Software\wbtest-css-<番号>）に向けます。
    保存の失敗などの本物では起きない条件も、ツールのコピーを書き換えて確かめます。
    最後に、偽の関数を使わずに本物の状態でも保存と比較を行い、続けて 2 回取っても誤検知が出ないことを確かめます。
    ダミー環境とテスト専用のレジストリのキーは、終了時に削除します。

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
$regBase  = 'Software\wbtest-css-{0}' -f (Get-Random -Minimum 100000 -Maximum 999999)
$allItems = 'Service,Port,Task,Software,HotFix,Network,Route,Share,LocalUser,LocalGroup,Hosts,Environment,Firewall'
$script:failCount = 0
$script:runNumber = 0
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$adminNote = '※ 管理者として実行すると、取れる情報が増えます'

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

function Write-Text([string]$FilePath, [string]$Text) {
    [void][IO.Directory]::CreateDirectory((Split-Path $FilePath))
    [IO.File]::WriteAllText($FilePath, $Text, (New-Object Text.UTF8Encoding $false))
}

# .ps1 は BOM 付きで書く（BOM がないと、5.1 は Shift_JIS として読み、日本語が壊れて文法エラーになる）
function Write-Script([string]$FilePath, [string]$Text) {
    [void][IO.Directory]::CreateDirectory((Split-Path $FilePath))
    [IO.File]::WriteAllText($FilePath, $Text, (New-Object Text.UTF8Encoding $true))
}

function Get-Clean($Value) { ([string]$Value) -replace '^="', '' -replace '"$', '' }

# ツールの処理を書き換えたコピーを作る。Pairs は @(元, 置き換え) の配列
function New-PatchedTool([string]$Name, [object[]]$Pairs) {
    $text = [IO.File]::ReadAllText($repoTool)
    foreach ($pair in $Pairs) {
        if (-not $text.Contains($pair[0])) { throw "書き換える場所が見つかりません: $($pair[0])" }
        $text = $text.Replace($pair[0], $pair[1])
    }
    $file = Join-Path $workRoot "patched\$Name\Compare-SystemState.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text, (New-Object Text.UTF8Encoding $true))
    $file
}

# 結果のファイルがないとき（保存の途中で失敗したときなど）も止まらないように、空として読む
function Read-CsvRows([string]$FilePath) {
    if (-not [IO.File]::Exists($FilePath)) { return @() }
    return @(Import-Csv -LiteralPath $FilePath -Encoding UTF8)
}

# 偽の環境で動かす。Phase は before か after。Fail はわざと失敗させるクラスやコマンドの名前（カンマ区切り）。
# Hosts を渡すと、偽の hosts を書き換える。Extra は、ほかの環境変数（FAKE_MANY など）
function Invoke-Fake([hashtable]$ToolArgs, [string]$Phase, [string]$Fail, [string]$Hosts, [hashtable]$Extra, [string]$ToolPath) {
    if (-not $ToolPath) { $ToolPath = $fakeTool }
    $script:runNumber++
    $out = Join-Path $workRoot ('out_{0}' -f $script:runNumber)
    if (-not $ToolArgs.ContainsKey('NoSave') -and -not $ToolArgs.ContainsKey('OutputFolder')) { $ToolArgs['OutputFolder'] = $out }
    if ($Hosts) { Write-Text (Join-Path $fakeRoot 'System32\drivers\etc\hosts') $Hosts }
    $argFile = Join-Path $workRoot ('args_{0}.json' -f $script:runNumber)
    [IO.File]::WriteAllText($argFile, ($ToolArgs | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding $true))
    $env:FAKE_PHASE = $Phase
    $env:FAKE_FAIL  = $Fail
    if ($Extra) { foreach ($key in $Extra.Keys) { [Environment]::SetEnvironmentVariable($key, [string]$Extra[$key]) } }
    $log  = & $Exe -NoProfile -ExecutionPolicy Bypass -File $wrapper -Tool $ToolPath -ArgFile $argFile -Root $fakeRoot 2>&1 | Out-String
    $code = $LASTEXITCODE
    $env:FAKE_PHASE = $null
    $env:FAKE_FAIL  = $null
    if ($Extra) { foreach ($key in $Extra.Keys) { [Environment]::SetEnvironmentVariable($key, $null) } }
    $dir = $null
    if (Test-Path -LiteralPath $out) { $dir = @(Get-ChildItem -LiteralPath $out -Directory | ForEach-Object { $_.FullName })[0] }
    $rows = @()
    $skips = @()
    $summary = ''
    if ($dir) {
        $rows  = @(Read-CsvRows (Join-Path $dir 'differences.csv'))
        $skips = @(Read-CsvRows (Join-Path $dir 'skipped.csv'))
        if ([IO.File]::Exists((Join-Path $dir 'summary.txt'))) { $summary = [IO.File]::ReadAllText((Join-Path $dir 'summary.txt')) }
    }
    [pscustomobject]@{ Code = $code; Log = (ConvertTo-FlatText $log); Dir = $dir; Rows = $rows; Skips = $skips; Summary = $summary }
}

function Get-DiffRow($Run, [string]$Label, [string]$Target) {
    @($Run.Rows | Where-Object { $_.'項目' -eq $Label -and (Get-Clean $_.'対象') -eq $Target })
}

# 行がなくても、例外で Assert-Equal を飛ばさず、FAIL として数えられるようにする
function Get-DiffCell($Run, [string]$Label, [string]$Target, [string]$Column) {
    $found = @(Get-DiffRow $Run $Label $Target)
    if ($found.Count -eq 0) { return '（行なし）' }
    return (@($found | ForEach-Object { [string]$_.$Column }) -join ',')
}

function Get-SkipCell($Run, [string]$Column) {
    if (@($Run.Skips).Count -eq 0) { return '（行なし）' }
    return (@($Run.Skips | ForEach-Object { [string]$_.$Column }) -join ',')
}

# スナップショットの CSV を読む
function Read-SnapshotCsv([string]$Dir, [string]$Name) { @(Read-CsvRows (Join-Path $Dir $Name)) }

# テスト専用のレジストリのキーに値を書く
function Set-FakeRegistry([string]$SubKey, [hashtable]$Values) {
    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey(($regBase + '\' + $SubKey))
    try {
        foreach ($name in $Values.Keys) {
            $value = $Values[$name]
            if ($value -is [int]) { $key.SetValue($name, $value, [Microsoft.Win32.RegistryValueKind]::DWord) }
            elseif ([string]$value -match '%') { $key.SetValue($name, $value, [Microsoft.Win32.RegistryValueKind]::ExpandString) }
            else { $key.SetValue($name, $value, [Microsoft.Win32.RegistryValueKind]::String) }
        }
    } finally {
        $key.Dispose()
    }
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

$fakeTool = $null
try {
    #region ダミー環境

    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool
    [void][IO.Directory]::CreateDirectory((Join-Path $fakeRoot 'System32\drivers\etc'))

    # 偽の環境で動かすコピー。hosts の場所と、ソフト・環境変数のレジストリを、テスト用の場所に向ける
    # （環境変数 SystemRoot を差し替えると、Winsock などの本物の部品が動かなくなる）。
    # ユーザーごとのサービスの種類は、名前が FakeUserSvc で始まるものだけ、テンプレートと実体として扱う
    $fakePairs = @(
        @("`$file = Join-Path `$env:SystemRoot 'System32\drivers\etc\hosts'", "`$file = Join-Path `$env:FAKE_ROOT 'System32\drivers\etc\hosts'"),
        @("`$root = Open-MachineKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' `$view[0]", "`$root = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey(`$env:FAKE_REG + '\' + `$view[0])"),
        @("`$key = Open-MachineKey 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment' 'Default'", "`$key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey(`$env:FAKE_REG + '\Environment')"),
        @("        `$key = Open-MachineKey ('SYSTEM\CurrentControlSet\Services\' + `$Name) 'Default'", "        if (`$Name -like 'FakeUserSvc_*') { `$info.Bits = 0xE0; return `$info } elseif (`$Name -eq 'FakeUserSvc') { `$info.Bits = 0x60; return `$info } elseif (`$Name -like 'Trig*') { `$info.Trigger = `$true; return `$info }`r`n        `$key = Open-MachineKey ('SYSTEM\CurrentControlSet\Services\' + `$Name) 'Default'")
    )
    $fakeTool = New-PatchedTool 'fake' $fakePairs

    $wrapperCode = @'
param([string]$Tool, [string]$ArgFile, [string]$Root)
$toolArgs = @{}
foreach ($property in ([IO.File]::ReadAllText($ArgFile) | ConvertFrom-Json).PSObject.Properties) { $toolArgs[$property.Name] = $property.Value }
$env:FAKE_ROOT = $Root
$env:FAKE_REG  = $env:FAKE_REG_BASE + '\' + $env:FAKE_PHASE
function global:Test-After { $env:FAKE_PHASE -eq 'after' }
function global:Test-Fail([string]$Name) { $env:FAKE_FAIL -and ($env:FAKE_FAIL -split ',') -contains $Name }
function global:Assert-Fake([string]$Name) { if (Test-Fail $Name) { throw "偽の失敗: $Name" } }
function global:Get-CimInstance {
    param($ClassName, $Filter, $ErrorAction)
    Assert-Fake $ClassName
    $after = Test-After
    switch ($ClassName) {
        'Win32_Service' {
            # 前後の空白と改行は無視する（W3SVC の実行ファイル）。突き合わせは大文字と小文字を区別せず、値は区別する（CaseSvc）
            [pscustomobject]@{ Name = 'W3SVC'; DisplayName = 'WWW'; State = $(if ($after) { 'Stopped' } else { 'Running' }); StartMode = 'Auto'; StartName = 'LocalSystem'; PathName = $(if ($after) { "  C:\w3.exe`r`n" } else { 'C:\w3.exe' }); ProcessId = 100; DelayedAutoStart = $false }
            [pscustomobject]@{ Name = 'BITS'; DisplayName = 'BITS'; State = $(if ($after) { 'Running' } else { 'Stopped' }); StartMode = 'Manual'; StartName = 'LocalSystem'; PathName = 'C:\b.exe'; ProcessId = 0 }
            [pscustomobject]@{ Name = 'Old'; DisplayName = 'Old'; State = 'Running'; StartMode = 'Auto'; StartName = 'x'; PathName = $(if ($after) { 'C:\new.exe -password SECRETA' } else { 'C:\old.exe -password SECRETB' }); ProcessId = 0 }
            if (-not $after) { [pscustomobject]@{ Name = 'Gone'; DisplayName = 'Gone'; State = 'Running'; StartMode = 'Auto'; StartName = 'x'; PathName = 'C:\g.exe'; ProcessId = 0 } }
            if ($after) { [pscustomobject]@{ Name = 'Tomcat9'; DisplayName = 'Tomcat'; State = 'Running'; StartMode = 'Auto'; StartName = 'LocalService'; PathName = 'C:\t.exe'; ProcessId = 200 } }
            [pscustomobject]@{ Name = $(if ($after) { 'casesvc' } else { 'CaseSvc' }); DisplayName = 'Case'; State = 'Running'; StartMode = 'Auto'; StartName = $(if ($after) { 'localsystem' } else { 'LocalSystem' }); PathName = 'C:\case.exe'; ProcessId = 0 }
            # ユーザーごとのサービス。ログオンのたびに実体の名前の末尾と状態が変わる
            [pscustomobject]@{ Name = 'FakeUserSvc'; DisplayName = 'User'; State = 'Stopped'; StartMode = 'Manual'; StartName = 'LocalSystem'; PathName = 'C:\u.exe'; ProcessId = 0 }
            [pscustomobject]@{ Name = $(if ($after) { 'FakeUserSvc_bbb2' } else { 'FakeUserSvc_aaa1' }); DisplayName = 'User_x'; State = $(if ($after) { 'Stopped' } else { 'Running' }); StartMode = 'Manual'; StartName = 'HOST\u'; PathName = 'C:\u.exe'; ProcessId = 0 }
            # 遅延開始とトリガー開始（名前が Trig で始まるものはトリガーあり）。状態の変化が参考になるのは、遅延＋トリガーと、手動だけ
            [pscustomobject]@{ Name = 'TrigDelayed'; DisplayName = 'TD'; State = $(if ($after) { 'Running' } else { 'Stopped' }); StartMode = 'Auto'; DelayedAutoStart = $true; StartName = 'LocalSystem'; PathName = 'C:\td.exe'; ProcessId = 0 }
            [pscustomobject]@{ Name = 'TrigManual'; DisplayName = 'TM'; State = $(if ($after) { 'Running' } else { 'Stopped' }); StartMode = 'Manual'; StartName = 'LocalSystem'; PathName = 'C:\tm.exe'; ProcessId = 0 }
            [pscustomobject]@{ Name = 'TrigAuto'; DisplayName = 'TA'; State = $(if ($after) { 'Stopped' } else { 'Running' }); StartMode = 'Auto'; DelayedAutoStart = $false; StartName = 'LocalSystem'; PathName = 'C:\ta.exe'; ProcessId = 0 }
            [pscustomobject]@{ Name = 'TrigAutoStart'; DisplayName = 'TS'; State = $(if ($after) { 'Running' } else { 'Stopped' }); StartMode = 'Auto'; DelayedAutoStart = $false; StartName = 'LocalSystem'; PathName = 'C:\ts.exe'; ProcessId = 0 }
            [pscustomobject]@{ Name = 'DelayedOnly'; DisplayName = 'DO'; State = $(if ($after) { 'Stopped' } else { 'Running' }); StartMode = 'Auto'; DelayedAutoStart = $true; StartName = 'LocalSystem'; PathName = 'C:\do.exe'; ProcessId = 0 }
            [pscustomobject]@{ Name = 'TrigDisabled'; DisplayName = 'TX'; State = 'Stopped'; StartMode = 'Disabled'; StartName = 'LocalSystem'; PathName = 'C:\tx.exe'; ProcessId = 0 }
            if ($after -and $env:FAKE_MANY) {
                foreach ($i in 1..[int]$env:FAKE_MANY) { [pscustomobject]@{ Name = ('Many{0:000}' -f $i); DisplayName = 'Many'; State = 'Running'; StartMode = 'Auto'; StartName = 'x'; PathName = 'C:\m.exe'; ProcessId = 0 } }
            }
        }
        'Win32_Process' { [pscustomobject]@{ ProcessId = 100; Name = 'svchost.exe' }; [pscustomobject]@{ ProcessId = 200; Name = 'java.exe' } }
        'Win32_QuickFixEngineering' { [pscustomobject]@{ HotFixID = 'KB1'; Description = 'Update'; InstalledOn = '11/1/2026' }; if ($after) { [pscustomobject]@{ HotFixID = 'KB2'; Description = 'Update'; InstalledOn = '11/3/2026' } } }
        'Win32_OperatingSystem' { [pscustomobject]@{ Version = $(if ($after) { '10.0.20349' } else { '10.0.20348' }); LastBootUpTime = $(if ($after) { [datetime]'2026-11-03 23:05:00' } else { [datetime]'2026-10-28 08:01:00' }) } }
        'Win32_Share' { [pscustomobject]@{ Name = 'Keys'; Path = 'C:\k'; Description = '' }; [pscustomobject]@{ Name = 'Count'; Path = 'C:\c'; Description = '' }; [pscustomobject]@{ Name = '2026-10-31'; Path = 'C:\d'; Description = '' } }
        'Win32_ComputerSystem' { [pscustomobject]@{ DomainRole = $(if (Test-Fail 'DC') { 5 } else { 3 }) } }
        'Win32_UserAccount' { [pscustomobject]@{ Name = 'admin'; Disabled = $after; PasswordExpires = $false; Description = 'd' } }
        'Win32_Group' { [pscustomobject]@{ Name = $(if (Test-Fail 'BadGroup') { 'NoSuchGroupForTest' } else { 'Users' }) } }
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
    # TCP の待ち受けのプロセスが変わったら、違いに数える
    [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 3389; OwningProcess = $(if (Test-After) { 200 } else { 100 }); State = 'Listen' }
    if (Test-After) {
        [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 8080; OwningProcess = 200; State = 'Listen' }
        [pscustomobject]@{ LocalAddress = '::'; LocalPort = 8080; OwningProcess = 200; State = 'Listen' }
        [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 49700; OwningProcess = 100; State = 'Listen' }
    }
}
function global:Get-NetUDPEndpoint {
    param($ErrorAction)
    [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 123; OwningProcess = 100 }
    [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 123; OwningProcess = 200 }
    # mDNS のように、同じ番号を使うプロセスが入れ替わる
    [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 5353; OwningProcess = $(if (Test-After) { 200 } else { 100 }) }
    if (Test-After) { [pscustomobject]@{ LocalAddress = '0.0.0.0'; LocalPort = 514; OwningProcess = 200 } }
}
function global:Get-NetRoute {
    param($PolicyStore, $ErrorAction)
    foreach ($prefix in '0.0.0.0/0', '224.0.0.0/4', 'ff00::/8', 'fe80::/64', '127.0.0.0/8', '169.254.0.0/16', '::1/128', '255.255.255.255/32', '10.0.0.0/24') {
        [pscustomobject]@{ DestinationPrefix = $prefix; NextHop = '10.0.0.1'; InterfaceAlias = 'イーサネット'; RouteMetric = $(if (Test-After) { 25 } else { 0 }) }
    }
}
function global:Get-NetFirewallProfile { param($PolicyStore, $ErrorAction) [pscustomobject]@{ Name = 'Domain'; Enabled = $(if (Test-After) { 'False' } else { 'True' }) } }
function global:Get-NetFirewallPortFilter { param($PolicyStore, $ErrorAction) [pscustomobject]@{ InstanceID = 'R1'; Protocol = 'TCP'; LocalPort = @('443', '80'); RemotePort = 'Any' } }
function global:Get-NetFirewallApplicationFilter { param($PolicyStore, $ErrorAction) [pscustomobject]@{ InstanceID = 'R1'; Program = 'Any' } }
function global:Get-NetFirewallAddressFilter { param($PolicyStore, $ErrorAction) [pscustomobject]@{ InstanceID = 'R1'; RemoteAddress = 'Any' } }
function global:Get-NetFirewallRule { param($PolicyStore, $ErrorAction) [pscustomobject]@{ InstanceID = 'R1'; DisplayName = 'Web'; Enabled = $(if (Test-After) { 'False' } else { 'True' }); Direction = 'Inbound'; Action = 'Allow'; Profile = 'Any' } }

# 偽のタスクスケジューラー（New-Object -ComObject Schedule.Service の代わり）
function global:New-FakeTrigger([int]$Days) {
    [pscustomobject]@{ Type = 2; StartBoundary = '2026-11-01T03:00:00'; DaysInterval = $Days; Enabled = $true; Repetition = [pscustomobject]@{ Interval = ''; Duration = '' } }
}
# フォルダとタスクは .NET のクラスで作る。本物の COM と同じく、メソッドやプロパティの例外が
# MethodInvocationException や GetValueInvocationException に包まれて届くようにするため（ScriptMethod では包まれ方が違う）
$global:fakeTaskSource = @"
using System;
using System.Runtime.InteropServices;
namespace WbTest.Css {
    public class FakeTask {
        public string Name; public bool Enabled; public object Body; public int Error;
        public object Definition { get { if (Error != 0) { throw new COMException("壊れたタスク", Error); } return Body; } }
    }
    public class FakeFolder {
        public string Path; public object[] TaskList; public object[] FolderList; public bool Locked;
        public object[] GetTasks(int flags) { if (Locked) { throw new UnauthorizedAccessException("アクセスが拒否されました"); } return TaskList; }
        public object[] GetFolders(int flags) { return FolderList; }
    }
}
"@
function global:New-FakeTask([string]$Name, [bool]$Enabled, [string]$UserId, [object[]]$Triggers, [string]$ActionPath, [string]$Arguments, [bool]$Broken) {
    $task = Microsoft.PowerShell.Utility\New-Object WbTest.Css.FakeTask
    $task.Name = $Name
    $task.Enabled = $Enabled
    # 定義を読むと例外になるタスク（0x80070002）
    if ($Broken) { $task.Error = -2147024894 }
    $action = [pscustomobject]@{ Type = 0; Path = $ActionPath; Arguments = $Arguments }
    $task.Body = [pscustomobject]@{ Principal = [pscustomobject]@{ UserId = $UserId; GroupId = '' }; Triggers = $Triggers; Actions = @($action) }
    $task
}
function global:New-FakeFolder([string]$Path, [object[]]$Tasks, [object[]]$Folders, [bool]$Locked) {
    $folder = Microsoft.PowerShell.Utility\New-Object WbTest.Css.FakeFolder
    $folder.Path = $Path
    $folder.TaskList = $Tasks
    $folder.FolderList = $Folders
    $folder.Locked = $Locked
    $folder
}
function global:New-FakeScheduler {
    # タスクを読むときだけコンパイルする
    if (-not ('WbTest.Css.FakeFolder' -as [type])) { Add-Type -TypeDefinition $global:fakeTaskSource }
    $after = Test-After
    $daily = New-FakeTrigger 1
    $rootTasks = @(
        (New-FakeTask 'Backup' (-not $after) 'S-1-5-18' @($daily) 'C:\backup.cmd' '' $false),
        (New-FakeTask 'Report' $true 'S-1-5-18' @($daily) 'C:\report.exe' $(if ($after) { "/pass:SECRETTASK2`r`n/x" } else { "/pass:SECRETTASK1`r`n/x" }) $false)
    )
    if (-not $after) { $rootTasks += New-FakeTask 'Gone' $true 'S-1-5-18' @($daily) 'C:\gone.cmd' '' $false }
    if ($after) { $rootTasks += New-FakeTask 'NewTask' $true 'S-1-5-18' @($daily) 'C:\new.cmd' '' $false }
    $job    = New-FakeTask 'Job' $true 'S-1-5-18' @(New-FakeTrigger $(if ($after) { 2 } else { 1 })) 'C:\job.cmd' '' (Test-Fail 'BrokenTask')
    $app    = New-FakeFolder '\App' @($job) @() $false
    $win    = New-FakeFolder '\Microsoft\Windows' @(New-FakeTask 'WinTask' $true 'S-1-5-18' @($daily) 'C:\win.cmd' '' $false) @() $false
    $ms     = New-FakeFolder '\Microsoft' @() @($win) $false
    $locked = New-FakeFolder '\Locked' @() @() (Test-Fail 'LockedFolder')
    $root   = New-FakeFolder '\' $rootTasks @($ms, $app, $locked) $false
    $service = [pscustomobject]@{ Root = $root }
    $service | Add-Member -MemberType ScriptMethod -Name Connect -Value { }
    $service | Add-Member -MemberType ScriptMethod -Name GetFolder -Value { param($Path) $this.Root }
    $service
}
# New-Object は、Schedule.Service のときだけ偽物を返し、ほかは本物に渡す（コレクションを並べて出さないよう , で包む）
function global:New-Object {
    param([Parameter(Position = 0)][string]$TypeName, [Parameter(Position = 1)][object[]]$ArgumentList, [string]$ComObject)
    if ($ComObject) {
        Assert-Fake 'Schedule'
        if ($ComObject -ne 'Schedule.Service') { throw "知らない COM: $ComObject" }
        return , (New-FakeScheduler)
    }
    if ($PSBoundParameters.ContainsKey('ArgumentList')) { return , (Microsoft.PowerShell.Utility\New-Object -TypeName $TypeName -ArgumentList $ArgumentList) }
    return , (Microsoft.PowerShell.Utility\New-Object -TypeName $TypeName)
}
& $Tool @toolArgs
exit $LASTEXITCODE
'@
    Write-Script $wrapper $wrapperCode

    # インストール済みソフト（64 ビットと 32 ビット）と、システムの環境変数の偽物
    $env:FAKE_REG_BASE = $regBase
    foreach ($phase in 'before', 'after') {
        $after = $phase -eq 'after'
        Set-FakeRegistry "$phase\Registry64\AppA" @{ DisplayName = 'App A'; DisplayVersion = $(if ($after) { '1.1' } else { '1.0' }); Publisher = 'Pub' }
        Set-FakeRegistry "$phase\Registry64\AppA-old" @{ DisplayName = 'App A'; DisplayVersion = '0.9' }
        # 同じ名前・同じバージョンが 2 つ登録されていても、バージョンは 1 つにまとめる
        Set-FakeRegistry "$phase\Registry64\AppA-dup" @{ DisplayName = 'App A'; DisplayVersion = $(if ($after) { '1.1' } else { '1.0' }) }
        Set-FakeRegistry "$phase\Registry64\Sys" @{ DisplayName = 'System Part'; DisplayVersion = '1'; SystemComponent = 1 }
        Set-FakeRegistry "$phase\Registry64\Upd" @{ DisplayName = 'Update for App A'; DisplayVersion = '1'; ParentKeyName = 'AppA' }
        Set-FakeRegistry "$phase\Registry64\NoName" @{ DisplayVersion = '1' }
        Set-FakeRegistry "$phase\Registry64\Num" @{ DisplayName = '2026'; DisplayVersion = '1.5' }
        Set-FakeRegistry "$phase\Registry32\AppA32" @{ DisplayName = 'App A'; DisplayVersion = '1.0' }
        if ($after) { Set-FakeRegistry "$phase\Registry32\New" @{ DisplayName = 'New App'; DisplayVersion = '1.0' } }
        else { Set-FakeRegistry "$phase\Registry32\Old" @{ DisplayName = 'Old App'; DisplayVersion = '3.0' } }
        $envValues = @{ Path = '%SystemRoot%\system32;C:\x'; SECRET_KEY = $(if ($after) { 'pass-NEWSECRET' } else { 'pass-OLDSECRET' }); Keep = 'v' }
        if ($after) { $envValues['Added'] = 'pass-ADDEDSECRET' }
        Set-FakeRegistry "$phase\Environment" $envValues
    }

    $hostsBefore = "127.0.0.1 localhost # コメント`r`n# 行のコメント`r`n`r`n10.0.0.9 db db.local`r`n"
    $hostsAfter  = "127.0.0.1 localhost`r`n10.0.0.9 db`r`n10.0.0.10 web`r`n"

    #endregion

    #region 保存

    Write-Section 'スナップショットの保存'
    $before = Invoke-Fake @{ Snapshot = $true } 'before' '' $hostsBefore
    Assert-Equal '終了コード' $before.Code '0'
    Assert-Equal 'フォルダ名' ((Split-Path $before.Dir -Leaf) -match ('^SystemSnapshot_' + [regex]::Escape($env:COMPUTERNAME) + '_\d{8}_\d{6}$')) 'True'
    Assert-Equal '比べるコマンドを表示' (Test-Contains $before.Log ('作業の後で、次のコマンドで比べられます:powershell -ExecutionPolicy Bypass -File {0} -Reference {1}' -f $fakeTool, $before.Dir)) 'True'
    Assert-Equal '件数の表（サービス）' ($before.Log -match ' サービス\s+12 ') 'True'
    Assert-Equal '件数の表（ソフト）' ($before.Log -match ' インストール済みソフト\s+4 ') 'True'
    Assert-Equal '管理者でないときだけ案内' (Test-Contains $before.Log $adminNote) "$(-not $isAdmin)"
    Assert-Equal 'すべて取れたら、取れなかった案内は出さない' (Test-Contains $before.Log '取れなかったもの') 'False'
    foreach ($name in 'info.txt', 'services.csv', 'ports.csv', 'tasks.csv', 'software.csv', 'hotfixes.csv', 'network.csv', 'routes.csv', 'shares.csv', 'localusers.csv', 'localgroups.csv', 'hosts.csv', 'environment.csv', 'firewall.csv', 'skipped.csv') {
        $path = Join-Path $before.Dir $name
        if (Test-Path -LiteralPath $path) {
            $bytes = [IO.File]::ReadAllBytes($path)
            Assert-Equal ("$name は BOM 付き") ('{0}-{1}-{2}' -f $bytes[0], $bytes[1], $bytes[2]) '239-187-191'
        } else {
            Assert-Equal ("$name がある") 'False' 'True'
        }
    }
    $info = [IO.File]::ReadAllText((Join-Path $before.Dir 'info.txt'))
    Assert-Equal 'info.txt のタスクの範囲' (Test-Contains $info 'タスクの範囲 = \Microsoft\ の下を除く') 'True'
    Assert-Equal 'info.txt の項目' (@([regex]::Matches($info, '項目\.\w+ = 取得')).Count) '13'
    Assert-Equal 'services.csv の列' ([IO.File]::ReadAllLines((Join-Path $before.Dir 'services.csv'))[0]) '"サービス名","表示名","状態","スタートアップ","実行ユーザー","実行ファイル"'
    $services = Read-SnapshotCsv $before.Dir 'services.csv'
    Assert-Equal 'ユーザーごとのサービスは末尾を除いて 1 つにまとめる' (@($services | Where-Object { $_.'サービス名' -like 'FakeUserSvc*' } | ForEach-Object { $_.'サービス名' + '/' + $_.'状態' + '/' + $_.'実行ユーザー' }) -join ',') 'FakeUserSvc//'
    Assert-Equal 'スタートアップに遅延とトリガーを書き添える' (@($services | Where-Object { $_.'サービス名' -in 'W3SVC', 'DelayedOnly', 'TrigDelayed', 'TrigAuto', 'TrigManual', 'TrigDisabled' } | Sort-Object { $_.'サービス名' } | ForEach-Object { $_.'サービス名' + '=' + $_.'スタートアップ' }) -join ',') 'DelayedOnly=自動（遅延）,TrigAuto=自動（トリガー）,TrigDelayed=自動（遅延・トリガー）,TrigDisabled=無効,TrigManual=手動（トリガー）,W3SVC=自動'
    $routes = Read-SnapshotCsv $before.Dir 'routes.csv'
    Assert-Equal 'マルチキャストなどのルートは取らない' (@($routes | ForEach-Object { Get-Clean $_.'宛先' } | Sort-Object) -join ',') '0.0.0.0/0,10.0.0.0/24'
    $hostsRows = Read-SnapshotCsv $before.Dir 'hosts.csv'
    Assert-Equal 'hosts はコメントを除き、組ごとに 1 行' (@($hostsRows | ForEach-Object { (Get-Clean $_.'アドレス') + ' ' + $_.'ホスト名' } | Sort-Object) -join ',') '10.0.0.9 db,10.0.0.9 db.local,127.0.0.1 localhost'
    $ports = Read-SnapshotCsv $before.Dir 'ports.csv'
    Assert-Equal '接続中の TCP は取らない' (@($ports | Where-Object { (Get-Clean $_.'ポート') -eq '50000' }).Count) '0'
    $udp = @($ports | Where-Object { (Get-Clean $_.'ポート') -eq '123' })
    Assert-Equal '同じポートのプロセスは 1 行にまとめる' (($udp | ForEach-Object { Get-Clean $_.'プロセス名' }) -join '/') 'java.exe, svchost.exe'
    $shareRows = Read-SnapshotCsv $before.Dir 'shares.csv'
    Assert-Equal 'Excel で変わる名前は ="..."' (@($shareRows | Where-Object { $_.'共有名' -eq '="2026-10-31"' }).Count) '1'
    $tasks = Read-SnapshotCsv $before.Dir 'tasks.csv'
    Assert-Equal 'タスクは \Microsoft\ の下を除く' (@($tasks | ForEach-Object { $_.'パス' } | Sort-Object) -join ',') '\App\Job,\Backup,\Gone,\Report'
    $report = @($tasks | Where-Object { $_.'パス' -eq '\Report' })
    if ($report.Count -gt 0) {
        Assert-Equal 'タスクの実行ユーザーは SID を名前に直す' ($report[0].'実行ユーザー' -match '\\SYSTEM$') 'True'
        Assert-Equal 'タスクのトリガーの文章' $report[0].'トリガー' '毎日 03:00'
        Assert-Equal 'タスクの操作の途中の改行は空白 1 つ' $report[0].'操作' 'C:\report.exe /pass:SECRETTASK1 /x'
    } else {
        Assert-Equal 'tasks.csv に \Report がある' 'なし' 'あり'
    }
    $software = Read-SnapshotCsv $before.Dir 'software.csv'
    Assert-Equal 'プログラムと機能に出るソフトだけ、名前とビットごと' (@($software | ForEach-Object { (Get-Clean $_.'名前') + '（' + $_.'ビット' + '）=' + (Get-Clean $_.'バージョン') } | Sort-Object) -join ',') '2026（64 ビット）=1.5,App A（32 ビット）=1.0,App A（64 ビット）=0.9, 1.0,Old App（32 ビット）=3.0'
    Assert-Equal 'Excel で変わるソフトの名前は ="..."' (@($software | Where-Object { $_.'名前' -eq '="2026"' }).Count) '1'
    $firewall = Read-SnapshotCsv $before.Dir 'firewall.csv'
    Assert-Equal 'firewall.csv の規則' (@($firewall | Where-Object { $_.'種類' -eq '規則' } | ForEach-Object { @($_.'ID', $_.'表示名', $_.'有効', $_.'方向', $_.'操作', $_.'プロファイル', $_.'プロトコル', (Get-Clean $_.'ローカルポート'), $_.'リモートポート', $_.'プログラム', $_.'リモートアドレス') -join '|' }) -join ',') 'R1|Web|有効|受信|許可|Any|TCP|443, 80|Any|Any|Any'
    Assert-Equal 'firewall.csv のプロファイル' (@($firewall | Where-Object { $_.'種類' -eq 'プロファイル' } | ForEach-Object { $_.'ID' + '|' + $_.'有効' }) -join ',') 'Domain|有効'
    $environment = Read-SnapshotCsv $before.Dir 'environment.csv'
    Assert-Equal '環境変数は %SystemRoot% を展開しない' (@($environment | Where-Object { $_.'変数名' -eq 'Path' } | ForEach-Object { $_.'値' }) -join ',') '%SystemRoot%\system32;C:\x'

    # -AllTasks で保存すると、\Microsoft\ の下も取る
    $allTasksSnap = Invoke-Fake @{ Snapshot = $true; Item = 'Task'; AllTasks = $true } 'before' ''
    Assert-Equal '-AllTasks: \Microsoft\ の下も取る' (@(Read-SnapshotCsv $allTasksSnap.Dir 'tasks.csv' | Where-Object { $_.'パス' -eq '\Microsoft\Windows\WinTask' }).Count) '1'
    Assert-Equal '-AllTasks: info.txt のタスクの範囲' (Test-Contains ([IO.File]::ReadAllText((Join-Path $allTasksSnap.Dir 'info.txt'))) 'タスクの範囲 = すべて') 'True'
    $partialItems = Invoke-Fake @{ Snapshot = $true; Item = 'HotFix,Share' } 'before' ''
    Assert-Equal '指定しなかった項目の CSV は作らない' (@(Get-ChildItem -LiteralPath $partialItems.Dir -Filter '*.csv' | ForEach-Object { $_.Name } | Sort-Object) -join ',') 'hotfixes.csv,shares.csv,skipped.csv'
    Assert-Equal '件数の表は指定した項目だけ' (Test-Contains $partialItems.Log 'サービス') 'False'

    #endregion

    #region 比較

    Write-Section '同じ状態を比べる'
    $same = Invoke-Fake @{ Reference = $before.Dir } 'before' ''
    Assert-Equal '終了コード 0' $same.Code '0'
    Assert-Equal '結果の行' (Test-Contains $same.Log '結果: 違いなし   追加 0 件  削除 0 件  変更 0 件  参考 0 件') 'True'
    Assert-Equal '違いの欄' (Test-Contains $same.Log '（違いはありません）') 'True'
    Assert-Equal '再起動なし' (Test-Contains $same.Log '再起動: なし') 'True'
    Assert-Equal 'OS が同じなら OS の行を出さない' (Test-Contains $same.Log 'OS:') 'False'
    Assert-Equal 'Keys や Count という名前の共有があっても数えられる' ($same.Log -match ' 共有フォルダ\s+3\s+3\s+0\s+0\s+0\s+0') 'True'
    Assert-Equal '両方とも管理者でないときだけ案内' (Test-Contains $same.Log $adminNote) "$(-not $isAdmin)"

    Write-Section '作業の後と比べる'
    $after = Invoke-Fake @{ Reference = $before.Dir } 'after' '' $hostsAfter
    Assert-Equal '終了コード 2' $after.Code '2'
    Assert-Equal '結果の行と件数' (Test-Contains $after.Log '結果: 違いあり   追加 9 件  削除 4 件  変更 17 件  参考 6 件') 'True'
    Assert-Equal '再起動あり' (Test-Contains $after.Log '再起動: あり（基準 2026-10-28 08:01:00 → 比較 2026-11-03 23:05:00）') 'True'
    Assert-Equal 'OS のビルドが変わったら OS の行' ($after.Log -match 'OS: 10\.0\.20348\.\d+ → 10\.0\.20349\.\d+') 'True'
    foreach ($name in 'summary.txt', 'differences.csv', 'skipped.csv', 'snapshot\info.txt') {
        Assert-Equal ("$name がある") (Test-Path -LiteralPath (Join-Path $after.Dir $name)) 'True'
    }
    Assert-Equal 'フォルダ名' ((Split-Path $after.Dir -Leaf) -match ('^SystemDiff_' + [regex]::Escape($env:COMPUTERNAME) + '_\d{8}_\d{6}$')) 'True'
    Assert-Equal 'differences.csv の列' ([IO.File]::ReadAllLines((Join-Path $after.Dir 'differences.csv'))[0]) '"結果","項目","対象","属性","基準の値","比較の値","備考"'
    Assert-Equal 'differences.csv の行の数（参考を含む）' $after.Rows.Count '36'
    Assert-Equal '項目ごとの件数の表' ($after.Log -match ' サービス\s+12\s+12\s+1\s+1\s+5\s+4') 'True'
    Assert-Equal 'トリガーのサービスの起動は参考（gpsvc など）' ('{0}/{1}' -f (Get-DiffCell $after 'サービス' 'TrigAutoStart' '結果'), (Get-DiffCell $after 'サービス' 'TrigAutoStart' '備考')) '参考/トリガー開始のサービスの起動'
    Assert-Equal '遅延＋トリガーのサービスの状態だけは参考' ('{0}/{1}' -f (Get-DiffCell $after 'サービス' 'TrigDelayed' '結果'), (Get-DiffCell $after 'サービス' 'TrigDelayed' '備考')) '参考/遅延開始・トリガー開始のサービスの状態の変化'
    Assert-Equal '手動（トリガー）のサービスの状態は参考' ('{0}/{1}' -f (Get-DiffCell $after 'サービス' 'TrigManual' '結果'), (Get-DiffCell $after 'サービス' 'TrigManual' '備考')) '参考/手動開始のサービスの状態の変化'
    Assert-Equal '遅延でないトリガーのサービスの停止は変更（LanmanServer など）' ('{0}/{1}→{2}' -f (Get-DiffCell $after 'サービス' 'TrigAuto' '結果'), (Get-DiffCell $after 'サービス' 'TrigAuto' '基準の値'), (Get-DiffCell $after 'サービス' 'TrigAuto' '比較の値')) '変更/実行中→停止'
    Assert-Equal 'トリガーのない遅延のサービスの停止は変更' (Get-DiffCell $after 'サービス' 'DelayedOnly' '結果') '変更'
    Assert-Equal 'UDP のプロセス名の変化は参考' ('{0}/{1}/{2}→{3}' -f (Get-DiffCell $after '待ち受けポート' 'UDP 0.0.0.0:5353' '結果'), (Get-DiffCell $after '待ち受けポート' 'UDP 0.0.0.0:5353' '備考'), (Get-DiffCell $after '待ち受けポート' 'UDP 0.0.0.0:5353' '基準の値'), (Get-DiffCell $after '待ち受けポート' 'UDP 0.0.0.0:5353' '比較の値')) '参考/UDP の番号を使うプロセスの変化/svchost.exe→java.exe'
    Assert-Equal 'UDP の番号の追加は違い' (Get-DiffCell $after '待ち受けポート' 'UDP 0.0.0.0:514' '結果') '追加'
    Assert-Equal 'サービスの追加' (Get-DiffCell $after 'サービス' 'Tomcat9' '結果') '追加'
    Assert-Equal '追加の値' (Get-DiffCell $after 'サービス' 'Tomcat9' '比較の値') 'スタートアップ=自動 / 状態=実行中 / 実行ユーザー=LocalService / 実行ファイル=C:\t.exe'
    Assert-Equal '追加の基準の値と属性は空欄' ((Get-DiffCell $after 'サービス' 'Tomcat9' '基準の値') + (Get-DiffCell $after 'サービス' 'Tomcat9' '属性')) ''
    Assert-Equal 'サービスの削除' (Get-DiffCell $after 'サービス' 'Gone' '結果') '削除'
    Assert-Equal '削除の値は基準の値' ((Get-DiffCell $after 'サービス' 'Gone' '基準の値') + '|' + (Get-DiffCell $after 'サービス' 'Gone' '比較の値')) 'スタートアップ=自動 / 状態=実行中 / 実行ユーザー=x / 実行ファイル=C:\g.exe|'
    Assert-Equal '自動のサービスの停止は変更（実行ファイルの前後の空白と改行は無視）' ('{0}/{1}/{2}/{3}' -f (Get-DiffCell $after 'サービス' 'W3SVC' '結果'), (Get-DiffCell $after 'サービス' 'W3SVC' '属性'), (Get-DiffCell $after 'サービス' 'W3SVC' '基準の値'), (Get-DiffCell $after 'サービス' 'W3SVC' '比較の値')) '変更/状態/実行中/停止'
    Assert-Equal '手動のサービスの状態だけは参考' ('{0}/{1}' -f (Get-DiffCell $after 'サービス' 'BITS' '結果'), (Get-DiffCell $after 'サービス' 'BITS' '備考')) '参考/手動開始のサービスの状態の変化'
    Assert-Equal '突き合わせは大文字と小文字を区別せず、値は区別する' ('{0}/{1}/{2}→{3}' -f (Get-DiffCell $after 'サービス' 'CaseSvc' '結果'), (Get-DiffCell $after 'サービス' 'CaseSvc' '属性'), (Get-DiffCell $after 'サービス' 'CaseSvc' '基準の値'), (Get-DiffCell $after 'サービス' 'CaseSvc' '比較の値')) '変更/実行ユーザー/LocalSystem→localsystem'
    Assert-Equal 'ユーザーごとのサービスは、末尾と状態が変わっても違いにしない' (@($after.Rows | Where-Object { (Get-Clean $_.'対象') -like 'FakeUserSvc*' }).Count) '0'
    Assert-Equal '参考は画面に出さない' (Test-Contains $after.Log '手動開始のサービスの状態の変化') 'False'
    Assert-Equal '動的ポートは参考' ('{0}/{1}' -f (Get-DiffCell $after '待ち受けポート' 'TCP 0.0.0.0:49700' '結果'), (Get-DiffCell $after '待ち受けポート' 'TCP 0.0.0.0:49700' '備考')) '参考/動的に割り当てられるポート（追加）'
    Assert-Equal 'ポートの追加' (Get-DiffCell $after '待ち受けポート' 'TCP 0.0.0.0:8080' '結果') '追加'
    Assert-Equal 'IPv6 のアドレスは [ ] で囲む' (Get-DiffCell $after '待ち受けポート' 'TCP [::]:8080' '結果') '追加'
    Assert-Equal '同じプロセスでサービスの同居が変わっても違いにしない' @(Get-DiffRow $after '待ち受けポート' 'UDP 0.0.0.0:123').Count '0'
    Assert-Equal 'タスクの有効・無効' ('{0}/{1}/{2}→{3}' -f (Get-DiffCell $after 'タスク' '\Backup' '結果'), (Get-DiffCell $after 'タスク' '\Backup' '属性'), (Get-DiffCell $after 'タスク' '\Backup' '基準の値'), (Get-DiffCell $after 'タスク' '\Backup' '比較の値')) '変更/有効/有効→無効'
    Assert-Equal 'タスクのトリガー' ('{0}→{1}' -f (Get-DiffCell $after 'タスク' '\App\Job' '基準の値'), (Get-DiffCell $after 'タスク' '\App\Job' '比較の値')) '毎日 03:00→2 日ごと 03:00'
    Assert-Equal 'タスクの追加と削除' ('{0}/{1}' -f (Get-DiffCell $after 'タスク' '\NewTask' '結果'), (Get-DiffCell $after 'タスク' '\Gone' '結果')) '追加/削除'
    Assert-Equal '\Microsoft\ の下のタスクは比べない' @($after.Rows | Where-Object { (Get-Clean $_.'対象') -like '\Microsoft\*' }).Count '0'
    # バージョンは Excel で数値に変わる形なので ="..." で書く
    Assert-Equal 'ソフトのバージョン（並べ替えてまとめる）' ('{0}→{1}' -f (Get-DiffCell $after 'インストール済みソフト' 'App A（64 ビット）' '基準の値'), (Get-DiffCell $after 'インストール済みソフト' 'App A（64 ビット）' '比較の値')) '="0.9, 1.0"→="0.9, 1.1"'
    Assert-Equal 'ソフトの追加と削除' ('{0}/{1}' -f (Get-DiffCell $after 'インストール済みソフト' 'New App（32 ビット）' '結果'), (Get-DiffCell $after 'インストール済みソフト' 'Old App（32 ビット）' '結果')) '追加/削除'
    Assert-Equal 'ビットが違えば別のソフト' @(Get-DiffRow $after 'インストール済みソフト' 'App A（32 ビット）').Count '0'
    Assert-Equal '更新プログラムの追加' (Get-DiffCell $after '更新プログラム' 'KB2' '結果') '追加'
    Assert-Equal 'メトリックの変更' (Get-DiffCell $after 'ルーティング' '0.0.0.0/0 → 10.0.0.1（イーサネット）' '属性') 'メトリック'
    Assert-Equal 'ユーザーの無効化' (Get-DiffCell $after 'ローカルユーザー' 'admin' '比較の値') '無効'
    Assert-Equal 'hosts の追加と削除' ('{0}/{1}' -f (Get-DiffCell $after 'hosts ファイル' '10.0.0.10 web' '結果'), (Get-DiffCell $after 'hosts ファイル' '10.0.0.9 db.local' '結果')) '追加/削除'
    Assert-Equal '環境変数の追加' ('{0}/{1}' -f (Get-DiffCell $after '環境変数' 'Added' '結果'), (Get-DiffCell $after '環境変数' 'Added' '比較の値')) '追加/値=pass-ADDEDSECRET'
    Assert-Equal 'TCP の待ち受けのプロセスの変化は違い' ('{0}/{1}→{2}' -f (Get-DiffCell $after '待ち受けポート' 'TCP 0.0.0.0:3389' '結果'), (Get-DiffCell $after '待ち受けポート' 'TCP 0.0.0.0:3389' '基準の値'), (Get-DiffCell $after '待ち受けポート' 'TCP 0.0.0.0:3389' '比較の値')) '変更/svchost.exe→java.exe'
    Assert-Equal 'ファイアウォールの対象は表示名（ID）' (Get-DiffCell $after 'ファイアウォール' 'Web（R1）' '属性') '有効'
    Assert-Equal 'ファイアウォールのプロファイル' ('{0}/{1}→{2}' -f (Get-DiffCell $after 'ファイアウォール' 'プロファイル Domain' '属性'), (Get-DiffCell $after 'ファイアウォール' 'プロファイル Domain' '基準の値'), (Get-DiffCell $after 'ファイアウォール' 'プロファイル Domain' '比較の値')) '有効/有効→無効'
    $net = @(Get-DiffRow $after 'ネットワーク設定' 'イーサネット')
    Assert-Equal 'IP の順序の違いは違いにせず、DNS の順序の違いは違いにする' (($net | ForEach-Object { $_.'属性' }) -join ',') 'DNS サーバー'
    Assert-Equal '画面の違いの行' ($after.Log -match '変更\s+サービス\s+W3SVC\s+状態: 実行中 → 停止') 'True'
    Assert-Equal '画面の追加の行は値を「・」でつなぐ' ($after.Log -match '追加\s+サービス\s+Tomcat9\s+自動・実行中・LocalService') 'True'
    Assert-Equal 'snapshot は今の状態' (Test-Contains ([IO.File]::ReadAllText((Join-Path $after.Dir 'snapshot\info.txt'))) 'OS = 10.0.20349.') 'True'

    Write-Section 'パスワードが含まれることのある値'
    foreach ($secret in 'SECRETA', 'SECRETB', 'SECRETTASK1', 'SECRETTASK2', 'NEWSECRET', 'OLDSECRET', 'ADDEDSECRET') {
        Assert-Equal "画面に値を出さない（$secret）" (Test-Contains $after.Log $secret) 'False'
        Assert-Equal "summary.txt にも書かない（$secret）" (Test-Contains $after.Summary $secret) 'False'
    }
    Assert-Equal 'サービスの実行ファイルが変わったことは出す' (Test-Contains $after.Log '実行ファイルが変わりました') 'True'
    Assert-Equal 'タスクの操作が変わったことは出す' (Test-Contains $after.Log '操作が変わりました') 'True'
    Assert-Equal '環境変数の値が変わったことは出す' (Test-Contains $after.Log '値が変わりました') 'True'
    Assert-Equal 'CSV には書く（サービス）' ((Get-DiffCell $after 'サービス' 'Old' '比較の値')) 'C:\new.exe -password SECRETA'
    Assert-Equal 'CSV には書く（タスク）' ((Get-DiffCell $after 'タスク' '\Report' '比較の値')) 'C:\report.exe /pass:SECRETTASK2 /x'
    Assert-Equal 'CSV には書く（環境変数）' ((Get-DiffCell $after '環境変数' 'SECRET_KEY' '基準の値')) 'pass-OLDSECRET'

    Write-Section '保存したもの同士'
    $two = Invoke-Fake @{ Reference = $before.Dir; Difference = (Join-Path $after.Dir 'snapshot') } 'before' ''
    Assert-Equal '終了コード 2' $two.Code '2'
    Assert-Equal '今の状態と比べたときと同じ件数（読み戻しで値が変わらない）' (Test-Contains $two.Log '結果: 違いあり   追加 9 件  削除 4 件  変更 17 件  参考 6 件') 'True'
    Assert-Equal '比較の行は保存したフォルダ' (Test-Contains $two.Log ('比較: {0} ' -f $env:COMPUTERNAME)) 'True'
    Assert-Equal 'snapshot フォルダを作らない' (Test-Path -LiteralPath (Join-Path $two.Dir 'snapshot')) 'False'
    Assert-Equal 'Excel で変わる名前を元に戻して突き合わせる' (@(Get-DiffRow $two '共有フォルダ' '2026-10-31').Count + @(Get-DiffRow $two 'インストール済みソフト' '2026（64 ビット）').Count) '0'

    Write-Section '-Item の絞り込みとタスクの範囲'
    $only = Invoke-Fake @{ Reference = $before.Dir; Item = 'hotfix' } 'after' ''
    Assert-Equal '大文字と小文字を区別しない' $only.Rows.Count '1'
    $quoted = Invoke-Fake @{ Reference = $before.Dir; Item = "'HotFix', 'Share'"; NoSave = $true } 'after' ''
    Assert-Equal '-Item の値の '' と空白を除く' ('{0}/{1}' -f $quoted.Code, ($quoted.Log -match ' 共有フォルダ\s+3\s+3')) '2/True'
    $all = Invoke-Fake @{ Reference = $partialItems.Dir } 'after' ''
    Assert-Equal '省略すると基準の項目を比べる' ('{0}/{1}' -f $all.Rows.Count, (Test-Contains $all.Log 'サービス')) '1/False'
    $allTasksCmp = Invoke-Fake @{ Reference = $allTasksSnap.Dir } 'before' ''
    Assert-Equal '-AllTasks の基準と比べるときは、今の状態も \Microsoft\ の下まで取る' ('{0}/{1}' -f $allTasksCmp.Code, (Test-Contains ([IO.File]::ReadAllText((Join-Path $allTasksCmp.Dir 'snapshot\info.txt'))) 'タスクの範囲 = すべて')) '0/True'

    Write-Section '画面に出す違いの数'
    $many = Invoke-Fake @{ Reference = $before.Dir; Item = 'Service' } 'after' '' '' @{ FAKE_MANY = 250 }
    Assert-Equal '件数' (Test-Contains $many.Log '追加 251 件  削除 1 件  変更 5 件  参考 4 件') 'True'
    Assert-Equal '画面は 200 件まで' (Test-Contains $many.Log '（ほか 57 件は differences.csv を参照してください）') 'True'
    # 対象の順に並ぶので、画面の 200 件は CaseSvc・DelayedOnly・Gone と Many001～Many197
    Assert-Equal '画面に出した行' (Get-Occurrence $many.Log 'Many') '197'
    Assert-Equal 'summary.txt はすべて' (Get-Occurrence $many.Summary 'Many') '250'
    $manyNoSave = Invoke-Fake @{ Reference = $before.Dir; Item = 'Service'; NoSave = $true } 'after' '' '' @{ FAKE_MANY = 250 }
    Assert-Equal '-NoSave ならすべて出す' ('{0}/{1}' -f (Get-Occurrence $manyNoSave.Log 'Many'), (Test-Contains $manyNoSave.Log '（ほか')) '250/False'
    Assert-Equal '-NoSave なら保存先を出さない' (Test-Contains $manyNoSave.Log '保存先:') 'False'

    #endregion

    #region 比べられないとき

    Write-Section '項目が取れないとき（レビューの指摘）'
    $tcp = Invoke-Fake @{ Reference = $before.Dir; Item = 'Port,HotFix' } 'before' 'Tcp'
    Assert-Equal 'TCP が取れないと、すべて削除に見せない' @(Get-DiffRow $tcp '待ち受けポート' 'TCP 0.0.0.0:443').Count '0'
    Assert-Equal '比べられないので終了コード 1' $tcp.Code '1'
    Assert-Equal '結果の行に書く' (Test-Contains $tcp.Log '結果: 違いなし（ただし、比べられなかった項目が 1 件あります。skipped.csv を確認してください）') 'True'
    Assert-Equal 'skipped.csv に残す' ('{0}/{1}/{2}' -f (Get-SkipCell $tcp '理由'), (Get-SkipCell $tcp '項目'), (Get-SkipCell $tcp '詳細')) 'その他のエラー/待ち受けポート/比較: 偽の失敗: Tcp'
    Assert-Equal '件数の表は理由と -' ($tcp.Log -match ' 待ち受けポート\s+\d+\s+その他のエラー\s+-\s+-\s+-\s+-') 'True'

    $noSave = Invoke-Fake @{ Reference = $before.Dir; Item = 'Port,HotFix'; NoSave = $true } 'before' 'Tcp'
    Assert-Equal '-NoSave のときは理由を画面に出す' (Test-Contains $noSave.Log '■ 比べられなかったもの') 'True'
    Assert-Equal '-NoSave のときは下の一覧を案内する' (Test-Contains $noSave.Log '比べられなかった項目が 1 件あります。下の一覧を確認してください') 'True'
    Assert-Equal '-NoSave のときは skipped.csv を案内しない' (Test-Contains $noSave.Log 'skipped.csv') 'False'

    $tcpDiff = Invoke-Fake @{ Reference = $before.Dir; Item = 'Port,HotFix' } 'after' 'Tcp'
    Assert-Equal '違いがあって比べられない項目もあれば 2' $tcpDiff.Code '2'
    Assert-Equal '違いがあって比べられない項目もあるときの案内' (Test-Contains $tcpDiff.Log '※ 比べられなかった項目が 1 件あります（skipped.csv を確認してください）') 'True'

    $snapFail = Invoke-Fake @{ Snapshot = $true; Item = 'Port,HotFix' } 'before' 'Tcp'
    Assert-Equal '保存で取れなかった項目: 止まらない' $snapFail.Code '0'
    Assert-Equal '保存で取れなかった項目: 案内' (Test-Contains $snapFail.Log '※ 取れなかったものが 1 件あります（skipped.csv を確認してください）') 'True'
    Assert-Equal '保存で取れなかった項目: 件数の表' ($snapFail.Log -match ' 待ち受けポート\s+その他のエラー') 'True'
    Assert-Equal '保存で取れなかった項目: CSV を作らない' (Test-Path -LiteralPath (Join-Path $snapFail.Dir 'ports.csv')) 'False'
    Assert-Equal '保存で取れなかった項目: info.txt' (Test-Contains ([IO.File]::ReadAllText((Join-Path $snapFail.Dir 'info.txt'))) '項目.Port = その他のエラー') 'True'
    $fromFailed = Invoke-Fake @{ Reference = $snapFail.Dir } 'before' ''
    Assert-Equal '基準で取れなかった項目は比べない' ('{0}/{1}' -f $fromFailed.Code, (Get-SkipCell $fromFailed '詳細')) '1/基準: 偽の失敗: Tcp'

    Write-Section '一部を読めないとき（レビューの指摘）'
    $partSnap = Invoke-Fake @{ Snapshot = $true; Item = 'LocalGroup' } 'before' 'BadGroup'
    Assert-Equal '止まらない' $partSnap.Code '0'
    Add-Content -LiteralPath (Join-Path $partSnap.Dir 'localgroups.csv') '"Users","HOST\gone"' -Encoding UTF8
    $part = Invoke-Fake @{ Reference = $partSnap.Dir } 'before' 'BadGroup'
    Assert-Equal '読めなかった側の削除は参考' ('{0}/{1}' -f (Get-DiffCell $part 'ローカルグループのメンバー' 'Users: HOST\gone' '結果'), (Get-DiffCell $part 'ローカルグループのメンバー' 'Users: HOST\gone' '備考')) '参考/削除に見えるが、一部を読めなかったため確かではない'
    Assert-Equal '違いに数えず終了コード 1' $part.Code '1'
    Assert-Equal '読めなかった対象の件数' (Test-Contains $part.Log '読めなかった対象が 2 件') 'True'

    # タスクのフォルダを読めない（権限なし）、タスクの定義を読めない（その他のエラー）
    $lockedSnap = Invoke-Fake @{ Snapshot = $true; Item = 'Task' } 'before' 'LockedFolder,BrokenTask'
    Assert-Equal 'タスクの一部を読めない: 止まらない' $lockedSnap.Code '0'
    Assert-Equal 'タスクの一部を読めない: skipped.csv' (@(Read-SnapshotCsv $lockedSnap.Dir 'skipped.csv' | ForEach-Object { $_.'理由' + ':' + ($_.'詳細' -replace ':.*$', '') } | Sort-Object) -join ',') 'その他のエラー:\App\Job,権限なし:\Locked'
    Assert-Equal 'タスクの一部を読めない: 読めた分は保存する' (@(Read-SnapshotCsv $lockedSnap.Dir 'tasks.csv').Count) '3'
    $locked = Invoke-Fake @{ Reference = $before.Dir; Item = 'Task' } 'after' 'LockedFolder,BrokenTask'
    Assert-Equal '今の状態で読めなかったタスクの削除は参考' ('{0}/{1}' -f (Get-DiffCell $locked 'タスク' '\App\Job' '結果'), (Get-DiffCell $locked 'タスク' '\Gone' '結果')) '参考/参考'
    Assert-Equal '基準で読めた側の追加は違いのまま' (Get-DiffCell $locked 'タスク' '\NewTask' '結果') '追加'
    Assert-Equal '読めなかったタスクの理由' (@($locked.Skips | ForEach-Object { $_.'理由' } | Sort-Object) -join ',') 'その他のエラー,権限なし'
    $noScheduler = Invoke-Fake @{ Reference = $before.Dir; Item = 'Task,HotFix' } 'before' 'Schedule'
    Assert-Equal 'タスクスケジューラーにつながらない' ('{0}/{1}' -f $noScheduler.Code, (Get-SkipCell $noScheduler '項目')) '1/タスク'

    Write-Section 'ドメインコントローラー'
    $dcSnap = Invoke-Fake @{ Snapshot = $true; Item = 'LocalUser,HotFix' } 'before' 'DC'
    Assert-Equal '対象外と表示' ($dcSnap.Log -match ' ローカルユーザー\s+対象外') 'True'
    Assert-Equal '対象外は取れなかったものに数えない' (Test-Contains $dcSnap.Log '取れなかったもの') 'False'
    $dc = Invoke-Fake @{ Reference = $dcSnap.Dir } 'before' 'DC'
    Assert-Equal '両方とも対象外なら、比べられないに数えない' $dc.Code '0'
    $dcOther = Invoke-Fake @{ Reference = $dcSnap.Dir } 'before' ''
    Assert-Equal '片方だけ対象外なら、比べられない' ('{0}/{1}' -f $dcOther.Code, (Test-Contains $dcOther.Log '比べられなかった項目が 1 件')) '1/True'

    Write-Section '基準と比較の違い'
    $copy = Join-Path $workRoot 'renamed'
    Copy-Item -LiteralPath (Join-Path $after.Dir 'snapshot') -Destination $copy -Recurse
    $infoPath = Join-Path $copy 'info.txt'
    $infoText = [IO.File]::ReadAllText($infoPath)
    $infoOther = [regex]::Replace($infoText, 'ホスト名 = .*', 'ホスト名 = OTHERHOST')
    $infoOther = [regex]::Replace($infoOther, '管理者 = .*', $(if ($isAdmin) { '管理者 = いいえ' } else { '管理者 = はい' }))
    [IO.File]::WriteAllText($infoPath, $infoOther, (New-Object Text.UTF8Encoding $true))
    $other = Invoke-Fake @{ Reference = $before.Dir; Difference = $copy; Item = 'HotFix' } 'before' ''
    Assert-Equal '違うコンピューターの案内' (Test-Contains $other.Log '※ 違うコンピューターの状態を比べています。名前や IP アドレスの違いも出ます') 'True'
    Assert-Equal '管理者の違いの案内' (Test-Contains $other.Log '※ 基準と比較で、管理者かどうかが違います。') 'True'
    Assert-Equal '違うコンピューターでは再起動の行を出さない' (Test-Contains $other.Log '再起動:') 'False'
    Assert-Equal '片方が管理者なら、管理者の案内は出さない' (Test-Contains $other.Log $adminNote) 'False'

    # 保存したもの同士で、タスクの範囲が違うと比べない
    $scopeCopy = Join-Path $workRoot 'scope'
    Copy-Item -LiteralPath (Join-Path $after.Dir 'snapshot') -Destination $scopeCopy -Recurse
    $scopeInfo = Join-Path $scopeCopy 'info.txt'
    [IO.File]::WriteAllText($scopeInfo, [regex]::Replace([IO.File]::ReadAllText($scopeInfo), 'タスクの範囲 = .*', 'タスクの範囲 = すべて'), (New-Object Text.UTF8Encoding $true))
    $scopeRun = Invoke-Fake @{ Reference = $before.Dir; Difference = $scopeCopy; Item = 'Task,HotFix' } 'before' ''
    Assert-Equal 'タスクの範囲が違うと比べない' ('{0}/{1}/{2}' -f $scopeRun.Code, (Test-Contains $scopeRun.Log '比べられなかった項目が 1 件'), @($scopeRun.Rows | Where-Object { $_.'項目' -eq 'タスク' }).Count) '2/True/0'
    Assert-Equal 'タスクの範囲の理由' (@($scopeRun.Skips | ForEach-Object { $_.'詳細' }) -join ',') '基準と比較で、タスクの範囲（-AllTasks）が違います'

    # 保存したもの同士で、片方にない項目
    $fewer = Invoke-Fake @{ Snapshot = $true; Item = 'HotFix' } 'before' ''
    $absentRun = Invoke-Fake @{ Reference = $partialItems.Dir; Difference = $fewer.Dir } 'before' ''
    Assert-Equal '比較のスナップショットにない項目' ('{0}/{1}/{2}' -f $absentRun.Code, (Get-SkipCell $absentRun '理由'), (Get-SkipCell $absentRun '項目')) '1/スナップショットにない/共有フォルダ'

    Write-Section '壊れたスナップショット'
    $broken = Join-Path $workRoot 'broken'
    Copy-Item -LiteralPath $before.Dir -Destination $broken -Recurse
    [IO.File]::WriteAllText((Join-Path $broken 'services.csv'), 'a,b,c')
    $badCsv = Invoke-Fake @{ Reference = $broken; Item = 'Service,HotFix' } 'before' ''
    Assert-Equal '1 つの CSV が読めなくても止まらない' (@($badCsv.Skips | ForEach-Object { $_.'理由' + ':' + $_.'詳細' }) -join ',') '読めない:基準: services.csv: 見出しの行が違います'
    Assert-Equal '読めない項目は比べない' (@($badCsv.Rows | Where-Object { $_.'項目' -eq 'サービス' }).Count) '0'
    [IO.File]::Delete((Join-Path $broken 'services.csv'))
    $missingCsv = Invoke-Fake @{ Reference = $broken; Item = 'Service,HotFix'; NoSave = $true } 'before' ''
    Assert-Equal 'CSV がなくても止まらない' ('{0}/{1}' -f $missingCsv.Code, (Test-Contains $missingCsv.Log '読めない')) '1/True'
    $noItems = Join-Path $workRoot 'noitems'
    [void][IO.Directory]::CreateDirectory($noItems)
    Write-Text (Join-Path $noItems 'info.txt') "ツール = Compare-SystemState`r`n保存日時 = 2026/11/03 22:00:00`r`n"
    $empty = Invoke-Fake @{ Reference = $noItems; NoSave = $true } 'before' ''
    Assert-Equal '項目のないスナップショットは止める' ('{0}/{1}' -f $empty.Code, (Test-Contains $empty.Log 'スナップショットのフォルダとして読めません')) '1/True'
    $badTime = Join-Path $workRoot 'badtime'
    [void][IO.Directory]::CreateDirectory($badTime)
    Write-Text (Join-Path $badTime 'info.txt') "ツール = Compare-SystemState`r`n保存日時 = きのう`r`n項目.HotFix = 取得`r`n"
    $badTimeRun = Invoke-Fake @{ Reference = $badTime; NoSave = $true } 'before' ''
    Assert-Equal '保存日時が読めないスナップショットは止める' ('{0}/{1}' -f $badTimeRun.Code, (Test-Contains $badTimeRun.Log 'スナップショットのフォルダとして読めません')) '1/True'

    #endregion

    #region 止まるとき

    Write-Section '止まるとき'
    $cases = @(
        @{ Name = '何も指定しない'; Args = @{}; Text = '作業の前は -Snapshot を、作業の後は -Reference を指定してください' }
        @{ Name = '-Snapshot と -Reference'; Args = @{ Snapshot = $true; Reference = 'x' }; Text = '-Snapshot のときは、-Reference は指定できません' }
        @{ Name = '-Snapshot と -Difference'; Args = @{ Snapshot = $true; Difference = 'x' }; Text = '-Snapshot のときは、-Difference は指定できません' }
        @{ Name = '-Snapshot と -NoSave'; Args = @{ Snapshot = $true; NoSave = $true }; Text = '-Snapshot のときは、-NoSave は指定できません' }
        @{ Name = '-AllTasks だけ'; Args = @{ AllTasks = $true; Reference = $before.Dir }; Text = '-AllTasks は、-Snapshot のときだけ指定できます' }
        @{ Name = '-Difference だけ'; Args = @{ Difference = $before.Dir }; Text = '-Difference を使うときは、-Reference も指定してください' }
        @{ Name = '見つからないフォルダ'; Args = @{ Reference = "$workRoot\nothing" }; Text = ('基準のフォルダを開けません: {0}\nothing' -f $workRoot) }
        @{ Name = '見つからない比較のフォルダ'; Args = @{ Reference = $before.Dir; Difference = "$workRoot\nothing2" }; Text = '比較のフォルダを開けません' }
        @{ Name = 'スナップショットでないフォルダ'; Args = @{ Reference = $workRoot }; Text = ('基準のスナップショットのフォルダとして読めません: {0}（Compare-SystemState の -Snapshot で保存したフォルダを指定してください）' -f $workRoot) }
        @{ Name = '基準と比較が同じ'; Args = @{ Reference = $before.Dir; Difference = ($before.Dir + '\') }; Text = '基準と比較が同じフォルダです' }
        @{ Name = '-Item の誤り'; Args = @{ Reference = $before.Dir; Item = 'Servce' }; Text = '-Item の値が正しくありません: Servce（使える値: Service, Port, Task, Software, HotFix, Network, Route, Share, LocalUser, LocalGroup, Hosts, Environment, Firewall）' }
        @{ Name = '-Item が空'; Args = @{ Reference = $before.Dir; Item = ', ,' }; Text = '-Item に項目がありません' }
        @{ Name = '基準にない項目'; Args = @{ Reference = $partialItems.Dir; Item = 'HotFix,Task,Firewall' }; Text = '基準のスナップショットに、次の項目がありません: Task, Firewall' }
    )
    foreach ($case in $cases) {
        $case.Args['NoSave'] = $case.Args.ContainsKey('NoSave') -and $case.Args['NoSave']
        if (-not $case.Args['NoSave']) { $case.Args.Remove('NoSave'); $case.Args['OutputFolder'] = "$workRoot\unused" }
        $run = Invoke-Fake $case.Args 'before' ''
        Assert-Equal ($case.Name + ': 止まる') ('{0}/{1}' -f $run.Code, (Test-Contains $run.Log $case.Text)) '1/True'
    }
    Assert-Equal '止まったときは保存先を作らない' (Test-Path -LiteralPath "$workRoot\unused") 'False'

    $cmdFile = "$workRoot\quote.cmd"
    Write-Text $cmdFile ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -Reference ""{2}\"" -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $Exe, $tool, $before.Dir)
    $quote = ConvertTo-FlatText (& cmd.exe /c $cmdFile 2>&1 | Out-String)
    Assert-Equal '\" で引数を巻き込む: 終了コード' $LASTEXITCODE '1'
    Assert-Equal '\" で引数を巻き込む: ヒント' (Test-Contains $quote '-Reference の値に引用符（"）が残っています') 'True'

    # 制約付き言語モードでは、最初に止まる。
    # __PSLockdownPolicy は 5.1 だとコンピューター全体の環境変数しか効かないので、起動した PowerShell の中で切り替えてから呼ぶ
    $clmCommand = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '{0}' -Snapshot -Item Hosts -OutputFolder '{1}'; exit `$LASTEXITCODE" -f $tool, "$workRoot\clm"
    $clm = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -Command $clmCommand 2>&1 | Out-String)
    $clmCode = $LASTEXITCODE
    Assert-Equal '制約付き言語モード: 止まる' ('{0}/{1}' -f $clmCode, (Test-Contains $clm '制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）')) '1/True'
    Assert-Equal '制約付き言語モード: 保存先を作らない' (Test-Path -LiteralPath "$workRoot\clm") 'False'

    Write-Text "$workRoot\notdir" 'x'
    $unwritable = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Snapshot -Item Hosts -OutputFolder "$workRoot\notdir" 2>&1 | Out-String
    Assert-Equal '保存できないスナップショットは止める' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains (ConvertTo-FlatText $unwritable) ('スナップショットを保存できません: {0}\notdir' -f $workRoot))) '1/True'
    $writeDiff = Invoke-Fake @{ Reference = $before.Dir; Item = 'HotFix'; OutputFolder = "$workRoot\notdir" } 'after' ''
    Assert-Equal '比べるときは警告して画面表示だけ' ('{0}/{1}/{2}' -f $writeDiff.Code, (Test-Contains $writeDiff.Log ('保存先に書き込めないため、結果は保存しません: {0}\notdir' -f $workRoot)), (Test-Contains $writeDiff.Log 'KB2')) '2/True/True'

    #endregion

    #region 保存の失敗と想定外のエラー

    Write-Section '保存の失敗と想定外のエラー'
    $saveFailTool = New-PatchedTool 'savefail' ($fakePairs + , @("Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'skipped.csv')) `$SkipColumns `$skipCsv", "throw 'disk full'"))
    $saveFail = Invoke-Fake @{ Reference = $before.Dir; Item = 'Port,HotFix' } 'after' 'Tcp' '' $null $saveFailTool
    Assert-Equal '比べた結果を保存できない: 警告' (Test-Contains $saveFail.Log '結果を保存できませんでした: disk full') 'True'
    Assert-Equal '比べた結果を保存できない: 作ったフォルダを消す' $saveFail.Dir ''
    Assert-Equal '比べた結果を保存できない: 理由を画面に出す' ('{0}/{1}' -f (Test-Contains $saveFail.Log '■ 比べられなかったもの'), (Test-Contains $saveFail.Log '下の一覧を確認してください')) 'True/True'
    Assert-Equal '比べた結果を保存できない: 終了コードは結果のとおり' $saveFail.Code '2'

    $snapSaveTool = New-PatchedTool 'snapsavefail' ($fakePairs + , @("[IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path `$Folder 'info.txt')), [string[]]`$info, `$Utf8Bom)", "throw 'disk full'"))
    $snapSave = Invoke-Fake @{ Snapshot = $true; Item = 'HotFix' } 'before' '' '' $null $snapSaveTool
    Assert-Equal 'スナップショットを保存できない: 止まる' ('{0}/{1}' -f $snapSave.Code, (Test-Contains $snapSave.Log 'スナップショットを保存できませんでした: disk full')) '1/True'
    Assert-Equal 'スナップショットを保存できない: 作ったフォルダを消す' $snapSave.Dir ''

    $crashTool = New-PatchedTool 'crash' ($fakePairs + , @('        $diffTotal = $added + $removed + $changed', "        throw 'unexpected'"))
    $crash = Invoke-Fake @{ Reference = $before.Dir; Item = 'HotFix' } 'after' '' '' $null $crashTool
    Assert-Equal '想定外のエラー: 表示' (Test-Contains $crash.Log 'エラーで止まりました。結果は保存していません。') 'True'
    Assert-Equal '想定外のエラー: 作ったフォルダを消す' $crash.Dir ''
    Assert-Equal '想定外のエラー: 終了コード' ($crash.Code -ne 0) 'True'

    #endregion

    #region パスと起動

    Write-Section 'パスと起動'
    # 日本語・スペース・[ ] を含むパス
    $jpDir = Join-Path $workRoot '日本語 と スペース [b]'
    [void][IO.Directory]::CreateDirectory($jpDir)
    Copy-Item -LiteralPath $before.Dir -Destination (Join-Path $jpDir 'snap') -Recurse
    $jp = Invoke-Fake @{ Reference = (Join-Path $jpDir 'snap'); Item = 'HotFix,Share'; OutputFolder = (Join-Path $jpDir 'out') } 'before' ''
    Assert-Equal '日本語・スペース・[ ] を含むパス' ('{0}/{1}' -f $jp.Code, (@(Get-ChildItem -LiteralPath (Join-Path $jpDir 'out') -Directory -Filter 'SystemDiff_*').Count)) '0/1'

    # 表示するコマンドは、必要なときだけ引用符で囲む（スペースは "..."、$ は '...'）
    $spaceOut = Join-Path $workRoot 'out dir'
    $spaceSnap = Invoke-Fake @{ Snapshot = $true; Item = 'HotFix'; OutputFolder = $spaceOut } 'before' ''
    $spaceDir = @(Get-ChildItem -LiteralPath $spaceOut -Directory | ForEach-Object { $_.FullName })[0]
    Assert-Equal 'スペースを含む保存先は "..." で表示' (Test-Contains $spaceSnap.Log ('-Reference "{0}"' -f $spaceDir)) 'True'
    $dollarOut = Join-Path $workRoot 'pay$dir'
    $dollarSnap = Invoke-Fake @{ Snapshot = $true; Item = 'HotFix'; OutputFolder = $dollarOut } 'before' ''
    $dollarDir = @(Get-ChildItem -LiteralPath $dollarOut -Directory | ForEach-Object { $_.FullName })[0]
    Assert-Equal '$ を含む保存先は ''...'' で表示' (Test-Contains $dollarSnap.Log ("-Reference '{0}'" -f $dollarDir)) 'True'

    # 260 文字を超える保存先と基準
    $longOut = "$workRoot\long\" + (('d' * 50 + '\') * 5) + 'out'
    $longSnap = Invoke-Fake @{ Snapshot = $true; Item = 'HotFix,Share'; OutputFolder = $longOut } 'before' ''
    $longDirs = @()
    if ([IO.Directory]::Exists('\\?\' + $longOut)) { $longDirs = @([IO.Directory]::GetDirectories('\\?\' + $longOut, 'SystemSnapshot_*')) }
    Assert-Equal '260 文字を超える保存先' ('{0}/{1}' -f $longSnap.Code, $longDirs.Count) '0/1'
    if ($longDirs.Count -gt 0) {
        $longRef = $longDirs[0].Substring(4)
        Assert-Equal '260 文字を超える保存先: CSV' ([IO.File]::Exists('\\?\' + $longRef + '\shares.csv')) 'True'
        $longCmp = Invoke-Fake @{ Reference = $longRef; NoSave = $true } 'before' ''
        Assert-Equal '260 文字を超える基準' ('{0}/{1}' -f $longCmp.Code, (Test-Contains $longCmp.Log '結果: 違いなし')) '0/True'
    }

    # 名前に [ ] を含むフォルダから、相対パスで起動する
    [void][IO.Directory]::CreateDirectory("$workRoot\[br]")
    $brBat = "$workRoot\bracket.cmd"
    Write-Text $brBat ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\[br]""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -Snapshot -Item Hosts -OutputFolder res`r`nexit /b %ERRORLEVEL%`r`n" -f $workRoot, $Exe, $tool)
    [void](& cmd.exe /c $brBat 2>&1 | Out-String)
    Assert-Equal '[ ] を含むフォルダから起動する: 終了コード' $LASTEXITCODE '0'
    Assert-Equal '[ ] を含むフォルダから起動する' (@(Get-ChildItem -LiteralPath "$workRoot\[br]\res" -Directory -ErrorAction SilentlyContinue).Count) '1'

    #endregion

    #region 本物の状態

    Write-Section '本物の状態で、続けて 2 回取る'
    $realOut = Join-Path $workRoot 'real'
    $realLog = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Snapshot -OutputFolder $realOut 2>&1 | Out-String
    Assert-Equal '保存の終了コード' $LASTEXITCODE '0'
    $realSnap = @(Get-ChildItem -LiteralPath $realOut -Directory | ForEach-Object { $_.FullName })[0]
    $realInfo = [IO.File]::ReadAllText((Join-Path $realSnap 'info.txt'))
    foreach ($name in 'Service', 'Port', 'Task', 'Software', 'HotFix', 'Network', 'Route', 'Share', 'Hosts', 'Environment') {
        Assert-Equal ("$name を取れる") (Test-Contains $realInfo ('項目.{0} = 取得' -f $name)) 'True'
    }
    foreach ($name in 'LocalUser', 'LocalGroup') {
        $got = (Test-Contains $realInfo ('項目.{0} = 取得' -f $name)) -or (Test-Contains $realInfo ('項目.{0} = 対象外' -f $name))
        Assert-Equal ("$name を取れる") $got 'True'
    }
    # ファイアウォールの実際に効いている規則（ActiveStore）は、管理者でないと読めない
    $firewallExpected = if ($isAdmin) { '取得' } else { '権限なし' }
    Assert-Equal 'Firewall は管理者なら取れる' (Test-Contains $realInfo ('項目.Firewall = {0}' -f $firewallExpected)) 'True'
    $realServices = @(Read-SnapshotCsv $realSnap 'services.csv')
    Assert-Equal 'サービスが 1 件以上' ($realServices.Count -gt 0) 'True'
    # スタートアップの「トリガー」は、レジストリの TriggerInfo があるサービスに付く（自動と手動のとき）。
    # ユーザーごとのサービス（状態が空欄）は、名前の末尾を除いているので、レジストリの名前と合わないことがあり、除く
    $triggerMismatch = @($realServices | Where-Object { $_.'スタートアップ' -match '^(自動|手動)' -and $_.'状態' } | Where-Object {
        $hasTrigger = Test-Path -LiteralPath ('HKLM:\SYSTEM\CurrentControlSet\Services\{0}\TriggerInfo' -f (Get-Clean $_.'サービス名'))
        $hasTrigger -ne ($_.'スタートアップ' -match 'トリガー')
    } | ForEach-Object { Get-Clean $_.'サービス名' })
    $triggerCount = @($realServices | Where-Object { $_.'スタートアップ' -match 'トリガー' }).Count
    Assert-Equal '本物のサービスのトリガー' ('{0}/{1}' -f ($triggerCount -gt 0), ($triggerMismatch -join ',')) 'True/'
    Assert-Equal '本物の hosts と環境変数を読む' ('{0}/{1}' -f (@(Read-SnapshotCsv $realSnap 'environment.csv' | Where-Object { $_.'変数名' -eq 'windir' }).Count), [IO.File]::Exists((Join-Path $realSnap 'hosts.csv'))) '1/True'
    if ($isAdmin) {
        Assert-Equal 'ファイアウォールのフィルターが規則と結び付く' (@(Read-SnapshotCsv $realSnap 'firewall.csv' | Where-Object { $_.'種類' -eq '規則' -and $_.'プロトコル' }).Count -gt 0) 'True'
    } else {
        Assert-Equal 'ファイアウォールを取れなければ CSV を作らない' (Test-Path -LiteralPath (Join-Path $realSnap 'firewall.csv')) 'False'
        Assert-Equal 'ファイアウォールを取れなければ skipped.csv に残す' (@(Read-SnapshotCsv $realSnap 'skipped.csv' | Where-Object { $_.'項目' -eq 'ファイアウォール' } | ForEach-Object { $_.'理由' }) -join ',') '権限なし'
    }
    $realDiffOut = Join-Path $workRoot 'realdiff'
    $realDiffLog = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Reference $realSnap -OutputFolder $realDiffOut 2>&1 | Out-String
    $realDir = @(Get-ChildItem -LiteralPath $realDiffOut -Directory | ForEach-Object { $_.FullName })[0]
    $realRows = @(Read-CsvRows (Join-Path $realDir 'differences.csv') | Where-Object { $_.'結果' -ne '参考' })
    foreach ($row in $realRows) { Write-Section ('本物の状態の違い: {0} / {1} / {2} / {3} → {4}' -f $row.'項目', (Get-Clean $row.'対象'), $row.'属性', $row.'基準の値', $row.'比較の値') }
    $stable = @($realRows | Where-Object { $_.'項目' -notin 'サービス', '待ち受けポート' })
    Assert-Equal '設定の項目に誤検知が出ない（サービスとポートを除く）' $stable.Count '0'
    Assert-Equal 'UDP のプロセス名の変化を違いに数えない' (@($realRows | Where-Object { $_.'項目' -eq '待ち受けポート' -and (Get-Clean $_.'対象') -like 'UDP *' -and $_.'属性' -eq 'プロセス名' }).Count) '0'
    Assert-Equal '再起動なし' (Test-Contains (ConvertTo-FlatText $realDiffLog) '再起動: なし') 'True'

    #endregion
} catch {
    # 途中で止まったときも失敗に数える。コマンドレットのエラーは、catch がないと try の後ろの END まで進み、合格の数が減るだけで気づけない
    '  FAIL  テストが途中で止まりました: {0}' -f ($_.Exception.Message -replace '\s*[\r\n]+\s*', ' ')
    $script:failCount++
} finally {
    try { [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($regBase, $false) } catch { }
    $env:FAKE_REG_BASE = $null
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
