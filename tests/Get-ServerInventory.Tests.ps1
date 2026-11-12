<#
.SYNOPSIS
    Get-ServerInventory の回帰テストです。

.DESCRIPTION
    サーバーの構成を自由に変えられないため、Get-CimInstance や Get-NetAdapter などと同じ名前の偽の関数を定義したラッパーから
    Get-ServerInventory のコピーを実行し、決まった値で確かめます。レジストリは、ツールのコピーを書き換えて、テスト専用の
    レジストリのキー（HKCU\Software\wbtest-gsi-<番号>）に向けます。ファームウェアの種類、プライマリ DNS サフィックス、
    powercfg、グループのメンバーも、コピーを書き換えて偽の値にします。保存の失敗などの本物では起きない条件も、
    ツールのコピーを書き換えて確かめます。
    最後に、偽の関数を使わずに本物の状態でも実行し、別の方法で数えた件数と突き合わせます（読むだけです）。
    ダミー環境とテスト専用のレジストリのキーは、終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Get-ServerInventory を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Get-ServerInventory\Get-ServerInventory.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Get-ServerInventory_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
# 結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーしたスクリプトを動かす
$tool     = Join-Path $workRoot 'tool\Get-ServerInventory.ps1'
$wrapper  = Join-Path $workRoot 'wrapper.ps1'
$regBase  = 'Software\wbtest-gsi-{0}' -f (Get-Random -Minimum 100000 -Maximum 999999)
$script:failCount = 0
$script:runNumber = 0
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$adminNote = '※ 管理者として実行すると、取れる情報が増えます'

# README の「CSV の列」
$Columns = [ordered]@{
    'parameters.csv'  = '分類', '項目', '値', '備考'
    'disks.csv'       = 'ディスク番号', 'モデル', '容量(GB)', '容量(バイト)', 'パーティションの形式', '接続', '状態'
    'volumes.csv'     = 'ディスク番号', 'パーティション番号', '種類', 'ドライブ', 'ラベル', 'ファイルシステム', '容量(GB)', '容量(バイト)', '空き(GB)', '空き(バイト)', '使用率', 'クラスターサイズ'
    'network.csv'     = '接続の名前', 'アダプター', '状態', '速度', 'MAC アドレス', 'DHCP', 'IPv4 アドレス', 'サブネットマスク', 'ゲートウェイ', 'DNS サーバー', 'IPv6', 'IPv6 アドレス', 'DNS サフィックス', 'DNS に登録', 'NetBIOS over TCP/IP', 'ネットワークの種類', 'MTU'
    'routes.csv'      = '宛先', 'ゲートウェイ', '接続の名前', 'メトリック'
    'features.csv'    = '名前', '表示名', '種類', '親'
    'software.csv'    = '名前', 'ビット', 'バージョン', '発行元', 'インストール日'
    'hotfixes.csv'    = 'KB 番号', '種類', 'インストール日', 'インストールした人'
    'services.csv'    = 'サービス名', '表示名', 'スタートアップ', '状態', '実行ユーザー'
    'localusers.csv'  = 'ユーザー名', '有効', 'パスワードの期限', '説明'
    'localgroups.csv' = 'グループ名', 'メンバー名'
    'shares.csv'      = '共有名', 'パス', '種類', '説明'
    'skipped.csv'     = '理由', '項目', '詳細'
}

# README の parameters.csv の分類と項目（決まったもの。スロットとページファイルは、分類の最後に足される）
$FixedItems = [ordered]@{
    'コンピューター'       = 'ホスト名', '完全修飾ドメイン名', 'ドメイン・ワークグループ', 'ドメインの役割', 'メーカー', 'モデル', 'シリアル番号', 'BIOS のバージョン', 'BIOS の日付', 'ファームウェア', 'セキュアブート'
    'OS'                   = '名前', 'エディション', 'バージョン', 'ビルド', 'アーキテクチャ', 'インストールの種類', 'インストール日', '最終起動', 'Windows フォルダ', 'OS の言語', 'システムロケール', 'タイムゾーン', 'ライセンス認証'
    'CPU'                  = '名前', 'ソケット数', 'コア数', '論理プロセッサ数', '最大クロック'
    'メモリ'               = '実装メモリ', '使えるメモリ'
    '仮想メモリ'           = , '管理'
    'メモリダンプ'         = '種類', '保存先', '自動的に再起動'
    'リモートデスクトップ' = '接続', 'ネットワークレベル認証', 'ポート'
    'ファイアウォール'     = 'ドメイン', 'プライベート', 'パブリック'
    'UAC'                  = 'UAC', '通知のレベル'
    '時刻同期'             = '種類', 'NTP サーバー'
    'Windows Update'       = 'WSUS サーバー', '自動更新'
    '電源'                 = , '電源プラン'
    'ネットワーク'         = , 'DNS サフィックスの検索一覧'
}

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

# .ps1 は BOM 付きで書く（BOM がないと、5.1 は Shift_JIS として読み、日本語が壊れて文法エラーになる）
function Write-Script([string]$FilePath, [string]$Text) {
    [void][IO.Directory]::CreateDirectory((Split-Path $FilePath))
    [IO.File]::WriteAllText($FilePath, $Text, (New-Object Text.UTF8Encoding $true))
}

# Excel 向けの ="..." を外す
function Get-Clean($Value) {
    $text = [string]$Value
    if ($text.StartsWith('="') -and $text.EndsWith('"') -and $text.Length -ge 3) { return $text.Substring(2, $text.Length - 3).Replace('""', '"') }
    return $text
}

# ツールの処理を書き換えたコピーを作る。Pairs は @(元, 置き換え) の配列。元の文字列は、ちょうど 1 か所だけあること
function New-PatchedTool([string]$Name, [object[]]$Pairs) {
    $text = [IO.File]::ReadAllText($repoTool)
    foreach ($pair in $Pairs) {
        $count = [regex]::Matches($text, [regex]::Escape($pair[0])).Count
        if ($count -ne 1) { throw ('書き換える場所が {0} か所あります: {1}' -f $count, $pair[0]) }
        $text = $text.Replace($pair[0], $pair[1])
    }
    $file = Join-Path $workRoot "patched\$Name\Get-ServerInventory.ps1"
    Write-Script $file $text
    $file
}

# 結果のファイルがないとき（保存の途中で失敗したときなど）も止まらないように、空として読む
function Read-CsvRows([string]$FilePath) {
    if (-not [IO.File]::Exists($FilePath)) { return @() }
    return @(Import-Csv -LiteralPath $FilePath -Encoding UTF8)
}

function Get-RunDir([string]$OutRoot) {
    if (-not [IO.Directory]::Exists($OutRoot)) { return $null }
    return @(Get-ChildItem -LiteralPath $OutRoot -Directory | ForEach-Object { $_.FullName })[0]
}

# 偽の環境で動かす。Scenario は a（クライアント）、b（サーバー）、c（設定の端のケース）、d（レジストリが空）。
# Env は、ほかの環境変数（FAKE_FAIL、FAKE_DENY、FAKE_DC など）
function Invoke-Fake([hashtable]$ToolArgs, [string]$Scenario, [hashtable]$Env, [string]$ToolPath) {
    if (-not $ToolPath) { $ToolPath = $fakeTool }
    if (-not $Env) { $Env = @{} }
    $script:runNumber++
    $out = Join-Path $workRoot ('out_{0}' -f $script:runNumber)
    if (-not $ToolArgs.ContainsKey('NoSave') -and -not $ToolArgs.ContainsKey('OutputFolder')) { $ToolArgs['OutputFolder'] = $out }
    $argFile = Join-Path $workRoot ('args_{0}.json' -f $script:runNumber)
    [IO.File]::WriteAllText($argFile, ($ToolArgs | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding $true))
    $vars = @{
        FAKE_SCENARIO = $Scenario; FAKE_REG = ($regBase + '\' + $Scenario); FAKE_POWERCFG = $powercfg[$Scenario]
        FAKE_FIRMWARE = $(if ($Scenario -eq 'b') { '1' } elseif ($Scenario -eq 'c') { '3' } else { '2' })
        FAKE_SUFFIX = $(if ($Scenario -eq 'b') { 'corp.example.local' } else { '' })
    }
    foreach ($key in $Env.Keys) { $vars[$key] = $Env[$key] }
    foreach ($key in $vars.Keys) { [Environment]::SetEnvironmentVariable($key, [string]$vars[$key]) }
    $log  = & $Exe -NoProfile -ExecutionPolicy Bypass -File $wrapper -Tool $ToolPath -ArgFile $argFile 2>&1 | Out-String
    $code = $LASTEXITCODE
    foreach ($key in $vars.Keys) { [Environment]::SetEnvironmentVariable($key, $null) }
    $dir = $null
    if ($ToolArgs.ContainsKey('OutputFolder')) { $dir = Get-RunDir ([string]$ToolArgs['OutputFolder']) }
    $summary = ''
    if ($dir -and [IO.File]::Exists((Join-Path $dir 'summary.txt'))) { $summary = [IO.File]::ReadAllText((Join-Path $dir 'summary.txt')) }
    [pscustomobject]@{ Code = $code; Log = (ConvertTo-FlatText $log); RawLog = $log; Dir = $dir; Summary = $summary }
}

function Read-Result($Run, [string]$File) {
    if (-not $Run.Dir) { return @() }
    return @(Read-CsvRows (Join-Path $Run.Dir $File))
}

# parameters.csv の値（Excel 向けの ="..." を外したもの）。行がなければ「（行なし）」
function Get-Param($Run, [string]$Category, [string]$Name, [string]$Column = '値') {
    $found = @(Read-Result $Run 'parameters.csv' | Where-Object { $_.'分類' -eq $Category -and (Get-Clean $_.'項目') -eq $Name })
    if ($found.Count -eq 0) { return '（行なし）' }
    return (@($found | ForEach-Object { Get-Clean $_.$Column }) -join ',')
}

# 一覧の CSV の 1 行を「列=値」ではなく、値だけを | でつないだ文字列にする（="..." は外す）
function Get-RowText($Row, [string[]]$Names) {
    return (@(foreach ($name in $Names) { Get-Clean $Row.$name }) -join '|')
}

function Get-ListRow($Run, [string]$File, [string]$Column, [string]$Key) {
    $found = @(Read-Result $Run $File | Where-Object { (Get-Clean $_.$Column) -eq $Key })
    if ($found.Count -eq 0) { return $null }
    return $found[0]
}

# CSV の生の行（ファイルの文字そのまま）。="..." になっているかを確かめるのに使う
function Get-RawLines($Run, [string]$File) {
    if (-not $Run.Dir) { return @() }
    return @([IO.File]::ReadAllLines((Join-Path $Run.Dir $File), [Text.Encoding]::UTF8))
}

# テスト専用のレジストリのキーに値を書く。int は DWORD、配列は複数行の文字列、% を含む文字列は展開できる文字列
function Set-FakeRegistry([string]$SubKey, [hashtable]$Values) {
    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey(($regBase + '\' + $SubKey))
    try {
        foreach ($name in $Values.Keys) {
            $value = $Values[$name]
            if ($value -is [int]) { $key.SetValue($name, $value, [Microsoft.Win32.RegistryValueKind]::DWord) }
            elseif ($value -is [array]) { $key.SetValue($name, [string[]]$value, [Microsoft.Win32.RegistryValueKind]::MultiString) }
            elseif ([string]$value -match '%') { $key.SetValue($name, $value, [Microsoft.Win32.RegistryValueKind]::ExpandString) }
            else { $key.SetValue($name, $value, [Microsoft.Win32.RegistryValueKind]::String) }
        }
    } finally {
        $key.Dispose()
    }
}

# 全角を 2 文字分として数えた幅
function Get-Width([string]$Text) {
    $width = 0
    foreach ($ch in $Text.ToCharArray()) { if ([int]$ch -lt 0x100 -or ([int]$ch -ge 0xFF61 -and [int]$ch -le 0xFF9F)) { $width++ } else { $width += 2 } }
    return $width
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

$fakeTool = $null
$powercfg = @{}
$denyDir  = $null
$denyRule = $null
try {
    #region ダミー環境

    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    # 偽の環境で動かすコピー。レジストリ、ファームウェアの種類、DNS サフィックス、powercfg、グループのメンバーを差し替える
    $fakePairs = @(
        , @("    `$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::`$View)",
            "    if (`$env:FAKE_REG) { return [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey(`$env:FAKE_REG + '\' + `$View + '\' + `$SubKey) }`r`n    `$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::`$View)")
        , @("    `$type = Get-FirmwareType", "    `$type = if (`$env:FAKE_FIRMWARE) { [int]`$env:FAKE_FIRMWARE } else { Get-FirmwareType }")
        , @("    try { `$suffix = [string][Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName } catch { }", "    `$suffix = [string]`$env:FAKE_SUFFIX")
        , @("    `$exe = Join-Path `$env:SystemRoot 'System32\powercfg.exe'", "    `$exe = `$env:FAKE_POWERCFG")
        , @("        Import-Module ServerManager -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:`$false", "        # 偽の Get-WindowsFeature を使うので、モジュールは読み込まない")
        , @("    `$group = [ADSI]('WinNT://./{0},group' -f `$GroupName)", "    if (`$env:FAKE_SCENARIO) { return (Get-FakeGroupMember `$GroupName) }`r`n    `$group = [ADSI]('WinNT://./{0},group' -f `$GroupName)")
    )
    $fakeTool = New-PatchedTool 'fake' $fakePairs

    # 偽の powercfg（名前あり、名前なし、失敗）
    $powercfg['a'] = Join-Path $workRoot 'powercfg_a.cmd'
    $powercfg['b'] = Join-Path $workRoot 'powercfg_b.cmd'
    $powercfg['c'] = Join-Path $workRoot 'powercfg_c.cmd'
    $powercfg['d'] = $powercfg['a']
    Write-Text $powercfg['a'] "@echo off`r`necho Power Scheme GUID: 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c  ^(High performance^)`r`n"
    Write-Text $powercfg['b'] "@echo off`r`necho Power Scheme GUID: 381b4222-f694-41f0-9685-ff5bb260df2e`r`n"
    Write-Text $powercfg['c'] "@echo off`r`necho Access is denied.`r`nexit /b 5`r`n"

    # レジストリ。a はクライアント、b はサーバー（グループポリシーあり）、c は端のケース、d は空
    $cv = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $un = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    Set-FakeRegistry "a\Registry64\$cv" @{ EditionID = 'Professional'; DisplayVersion = '24H2'; ReleaseId = '2009'; CurrentBuild = '26100'; UBR = 2033; InstallationType = 'Client' }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Control\Nls\Locale' @{ '' = '00000411' }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Control\SecureBoot\State' @{ UEFISecureBootEnabled = 1 }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Control\CrashControl' @{ CrashDumpEnabled = 7; DumpFile = '%SystemRoot%\MEMORY.DMP'; AutoReboot = 1 }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Control\Terminal Server' @{ fDenyTSConnections = 0 }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' @{ UserAuthentication = 0; PortNumber = 3390 }
    Set-FakeRegistry 'a\Registry64\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' @{ EnableLUA = 1; ConsentPromptBehaviorAdmin = 5; PromptOnSecureDesktop = 0 }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Services\W32Time\Parameters' @{ Type = 'NTP'; NtpServer = 'ntp.nict.jp,0x9' }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' @{ PagingFiles = @('?:\pagefile.sys') }
    Set-FakeRegistry "a\Registry64\$un\App1" @{ DisplayName = 'App One'; DisplayVersion = '16.0'; Publisher = 'Pub'; InstallDate = '20261015' }
    Set-FakeRegistry "a\Registry64\$un\App1b" @{ DisplayName = 'App One'; DisplayVersion = '16.1'; Publisher = 'Pub'; InstallDate = '20261020' }
    Set-FakeRegistry "a\Registry64\$un\Hidden" @{ DisplayName = 'Hidden Part'; SystemComponent = 1 }
    Set-FakeRegistry "a\Registry64\$un\Upd" @{ DisplayName = 'Update For App'; ParentKeyName = 'App1' }
    Set-FakeRegistry "a\Registry64\$un\NoName" @{ Publisher = 'x' }
    Set-FakeRegistry "a\Registry32\$un\App2" @{ DisplayName = '=Calc App'; DisplayVersion = '1.2.3' }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Services\CDPUserSvc_ab12' @{ Type = 224 }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Services\CDPUserSvc' @{ Type = 96 }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Services\TrigSvc' @{ Type = 16 }
    Set-FakeRegistry 'a\Registry64\SYSTEM\CurrentControlSet\Services\TrigSvc\TriggerInfo' @{ x = 1 }

    Set-FakeRegistry "b\Registry64\$cv" @{ EditionID = 'ServerStandard'; ReleaseId = '1607'; CurrentBuild = '14393'; UBR = 7428; InstallationType = 'Server Core' }
    Set-FakeRegistry 'b\Registry64\SYSTEM\CurrentControlSet\Control\Nls\Locale' @{ '' = '00000409' }
    Set-FakeRegistry 'b\Registry64\SYSTEM\CurrentControlSet\Control\CrashControl' @{ CrashDumpEnabled = 1; FilterPages = 1; DumpFile = '%SystemRoot%\MEMORY.DMP'; AutoReboot = 0 }
    Set-FakeRegistry 'b\Registry64\SYSTEM\CurrentControlSet\Control\Terminal Server' @{ fDenyTSConnections = 0 }
    Set-FakeRegistry 'b\Registry64\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services' @{ fDenyTSConnections = 1 }
    Set-FakeRegistry 'b\Registry64\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' @{ UserAuthentication = 1; PortNumber = 3389 }
    Set-FakeRegistry 'b\Registry64\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' @{ EnableLUA = 1; ConsentPromptBehaviorAdmin = 2; PromptOnSecureDesktop = 0 }
    Set-FakeRegistry 'b\Registry64\SYSTEM\CurrentControlSet\Services\W32Time\Parameters' @{ Type = 'NT5DS'; NtpServer = 'time.windows.com,0x9' }
    Set-FakeRegistry 'b\Registry64\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' @{ WUServer = 'http://wsus01:8530' }
    Set-FakeRegistry 'b\Registry64\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' @{ UseWUServer = 1; AUOptions = 4; ScheduledInstallDay = 0; ScheduledInstallTime = 3 }
    Set-FakeRegistry 'b\Registry64\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' @{ PagingFiles = @('C:\pagefile.sys 4096 8192', 'D:\pagefile.sys 0 0', 'E:\pagefile.sys', '?:\pagefile.sys') }

    Set-FakeRegistry "c\Registry64\$cv" @{ EditionID = 'Professional' }
    Set-FakeRegistry 'c\Registry64\SYSTEM\CurrentControlSet\Services\W32Time\Parameters' @{ Type = 'NoSync'; NtpServer = 'local.ntp' }
    Set-FakeRegistry 'c\Registry64\SOFTWARE\Policies\Microsoft\W32Time\Parameters' @{ NtpServer = 'gpo.ntp' }
    Set-FakeRegistry 'c\Registry64\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' @{ WUServer = 'http://old:8530' }
    Set-FakeRegistry 'c\Registry64\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' @{ NoAutoUpdate = 1 }
    Set-FakeRegistry 'c\Registry64\SYSTEM\CurrentControlSet\Control\CrashControl' @{ CrashDumpEnabled = 0 }
    Set-FakeRegistry 'c\Registry64\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' @{ EnableLUA = 0 }

    [void][Microsoft.Win32.Registry]::CurrentUser.CreateSubKey(($regBase + '\d\Registry64'))

    $wrapperCode = @'
param([string]$Tool, [string]$ArgFile)
$toolArgs = @{}
foreach ($property in ([IO.File]::ReadAllText($ArgFile) | ConvertFrom-Json).PSObject.Properties) { $toolArgs[$property.Name] = $property.Value }
if ($env:FAKE_CULTURE) { [Threading.Thread]::CurrentThread.CurrentCulture = $env:FAKE_CULTURE }
# 偽の関数より先に本物のモジュールを読み込んでおく（後から自動で読み込まれると、偽の関数が本物に置き換わる）
foreach ($module in 'Storage', 'NetAdapter', 'NetTCPIP', 'DnsClient', 'NetConnection', 'NetSecurity') { Import-Module $module -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }
function global:Test-Fail([string]$Name) { $env:FAKE_FAIL -and ($env:FAKE_FAIL -split ',') -contains $Name }
function global:Assert-Fake([string]$Name) {
    if ($env:FAKE_DENY -and ($env:FAKE_DENY -split ',') -contains $Name) { throw (New-Object UnauthorizedAccessException ('偽の拒否: ' + $Name)) }
    if (Test-Fail $Name) { throw ('偽の失敗: ' + $Name) }
}
function global:Test-Scenario([string]$Name) { $env:FAKE_SCENARIO -eq $Name }
function global:New-Fake([hashtable]$Values) { [pscustomobject]$Values }
function global:Get-CimInstance {
    param($ClassName, $Filter, $ErrorAction)
    Assert-Fake $ClassName
    $server = Test-Scenario 'b'
    switch ($ClassName) {
        'Win32_ComputerSystem' {
            if ($server) {
                New-Fake @{ Name = 'SRV01'; DNSHostName = 'srv01'; Domain = 'corp.example.local'; PartOfDomain = $true; Workgroup = $null; DomainRole = $(if ($env:FAKE_DC) { 5 } else { 3 }); Manufacturer = 'VMware, Inc.'; Model = 'VMware20,1'; TotalPhysicalMemory = [uint64]17072494002; AutomaticManagedPagefile = $false }
            } else {
                New-Fake @{ Name = 'WBTEST-PC'; DNSHostName = 'wbtest-pc'; Domain = 'WORKGROUP'; PartOfDomain = $false; Workgroup = 'WORKGROUP'; DomainRole = 0; Manufacturer = 'FakeMaker'; Model = 'Model 1'; TotalPhysicalMemory = [uint64]21367462298; AutomaticManagedPagefile = -not (Test-Scenario 'c') }
            }
        }
        'Win32_BIOS' {
            if ($server) { New-Fake @{ SerialNumber = 'VMware-56 4d'; SMBIOSBIOSVersion = 'VMW201.00V'; ReleaseDate = [datetime]'2023-05-22' } }
            else { New-Fake @{ SerialNumber = 'SN-0001'; SMBIOSBIOSVersion = '2212'; ReleaseDate = [datetime]'2022-12-13' } }
        }
        'Win32_OperatingSystem' {
            if ($server) { New-Fake @{ Caption = 'Microsoft Windows Server 2016 Standard'; BuildNumber = '14393'; OSArchitecture = '64 ビット'; ProductType = $(if ($env:FAKE_DC) { 2 } else { 3 }); InstallDate = [datetime]'2024-04-01 10:00:00'; LastBootUpTime = [datetime]'2026-11-03 23:05:00'; WindowsDirectory = 'C:\Windows'; OSLanguage = 1033 } }
            else { New-Fake @{ Caption = 'Microsoft Windows 11 Pro '; BuildNumber = '26100'; OSArchitecture = '64-bit'; ProductType = 1; InstallDate = [datetime]'2025-08-23 11:34:07'; LastBootUpTime = [datetime]'2026-11-11 20:43:19'; WindowsDirectory = 'C:\WINDOWS'; OSLanguage = 1041 } }
        }
        'Win32_TimeZone' { New-Fake @{ Caption = '(UTC+09:00) 大阪、札幌、東京' } }
        'SoftwareLicensingProduct' {
            if ($server) {
                # アドオン（ESU など）は、本体より先に並んでいても除く
                New-Fake @{ LicenseStatus = 1; Description = 'Windows(R) Operating System, VOLUME_MAK channel'; GracePeriodRemaining = 0; LicenseIsAddon = $true }
                New-Fake @{ LicenseStatus = 1; Description = 'Windows(R) Operating System, TIMEBASED_EVAL channel'; GracePeriodRemaining = 43200; LicenseIsAddon = $false }
            } elseif (-not (Test-Scenario 'c')) {
                New-Fake @{ LicenseStatus = 1; Description = 'Windows(R) Operating System, RETAIL channel'; GracePeriodRemaining = 0; LicenseIsAddon = $false }
            }
        }
        'Win32_Processor' {
            if ($server) { New-Fake @{ Name = 'Intel(R) Xeon(R) CPU E5-2680 v4 @ 2.40GHz'; NumberOfCores = 2; NumberOfLogicalProcessors = 2; MaxClockSpeed = 2400 } }
            else {
                New-Fake @{ Name = 'Intel(R) Xeon(R) Gold 6338 CPU @ 2.00GHz'; NumberOfCores = 4; NumberOfLogicalProcessors = 8; MaxClockSpeed = 2000 }
                New-Fake @{ Name = 'Intel(R) Xeon(R) Gold 6338 CPU @ 2.00GHz'; NumberOfCores = 4; NumberOfLogicalProcessors = 8; MaxClockSpeed = 2000 }
            }
        }
        'Win32_PhysicalMemory' {
            if (-not $server) {
                New-Fake @{ DeviceLocator = 'DIMM0'; BankLabel = 'BANK 0'; Capacity = [uint64]8589934592; Speed = 3200; ConfiguredClockSpeed = 3200; Manufacturer = 'Crucial'; PartNumber = 'CT8G  ' }
                New-Fake @{ DeviceLocator = 'DIMM1'; BankLabel = 'BANK 1'; Capacity = [uint64]8589934592; Speed = 3200; ConfiguredClockSpeed = 3200; Manufacturer = 'Crucial'; PartNumber = 'CT8G' }
                New-Fake @{ DeviceLocator = 'DIMM1'; BankLabel = 'BANK 2'; Capacity = [uint64]4294967296; Speed = 2666; ConfiguredClockSpeed = 0; Manufacturer = ''; PartNumber = '' }
            }
        }
        'Win32_PageFileUsage' {
            if ($server) { New-Fake @{ Name = 'c:\pagefile.sys'; AllocatedBaseSize = 4096 }; New-Fake @{ Name = 'F:\pagefile.sys'; AllocatedBaseSize = 1024 } }
            elseif (-not (Test-Scenario 'c')) { New-Fake @{ Name = 'C:\pagefile.sys'; AllocatedBaseSize = 4096 } }
        }
        'Win32_Volume' {
            New-Fake @{ DeviceID = '\\?\Volume{efi}\'; Label = ''; FileSystem = 'FAT32'; Capacity = [uint64]104857600; FreeSpace = [uint64]67108864; BlockSize = [uint64]1024 }
            New-Fake @{ DeviceID = '\\?\Volume{c}\'; Label = $null; FileSystem = 'NTFS'; Capacity = [uint64]107374182400; FreeSpace = [uint64]53687091200; BlockSize = [uint64]4096 }
            New-Fake @{ DeviceID = '\\?\Volume{rec}\'; Label = ''; FileSystem = 'NTFS'; Capacity = [uint64]1073741824; FreeSpace = [uint64]536870912; BlockSize = [uint64]4096 }
            New-Fake @{ DeviceID = '\\?\Volume{d}\'; Label = 'DATA'; FileSystem = 'NTFS'; Capacity = [uint64]214748364800; FreeSpace = [uint64]161061273600; BlockSize = [uint64]65536 }
            New-Fake @{ DeviceID = '\\?\Volume{cd}\'; Label = 'DVD'; FileSystem = 'CDFS'; Capacity = $null; FreeSpace = $null; BlockSize = $null }
        }
        'Win32_NetworkAdapterConfiguration' {
            New-Fake @{ InterfaceIndex = 5; TcpipNetbiosOptions = [uint32]2 }
            New-Fake @{ InterfaceIndex = 6; TcpipNetbiosOptions = $null }
            New-Fake @{ InterfaceIndex = 7; TcpipNetbiosOptions = [uint32]0 }
        }
        'Win32_OptionalFeature' {
            if ($server) { New-Fake @{ Name = 'ClientOnly'; Caption = 'Client Only' } }
            else { New-Fake @{ Name = 'Feat-B'; Caption = 'Feature B' }; New-Fake @{ Name = 'Feat-A'; Caption = 'Feature A' } }
        }
        'Win32_QuickFixEngineering' {
            New-Fake @{ HotFixID = 'KB3'; Description = 'Update'; InstalledOn = [datetime]'2026-10-17'; InstalledBy = 'NT AUTHORITY\SYSTEM' }
            New-Fake @{ HotFixID = 'KB0'; Description = 'Update'; InstalledOn = $null; InstalledBy = '' }
            New-Fake @{ HotFixID = 'KB2'; Description = 'Update'; InstalledOn = [datetime]'2026-01-15'; InstalledBy = 'NT AUTHORITY\SYSTEM' }
            New-Fake @{ HotFixID = 'KB1'; Description = 'Security Update'; InstalledOn = [datetime]'2026-10-17'; InstalledBy = 'CORP\admin' }
        }
        'Win32_Service' {
            New-Fake @{ Name = 'Svc1'; DisplayName = 'Service One'; State = 'Running'; StartMode = 'Auto'; DelayedAutoStart = $true; StartName = 'LocalSystem'; PathName = 'C:\svc1.exe -password SECRETPW' }
            New-Fake @{ Name = 'TrigSvc'; DisplayName = 'Trig'; State = 'Stopped'; StartMode = 'Manual'; DelayedAutoStart = $false; StartName = 'NT AUTHORITY\LocalService'; PathName = 'C:\t.exe' }
            New-Fake @{ Name = 'CDPUserSvc_ab12'; DisplayName = 'Connected Devices_ab12'; State = 'Running'; StartMode = 'Auto'; DelayedAutoStart = $false; StartName = ''; PathName = 'C:\c.exe' }
            New-Fake @{ Name = 'CDPUserSvc'; DisplayName = 'Connected Devices'; State = 'Stopped'; StartMode = 'Auto'; DelayedAutoStart = $false; StartName = ''; PathName = 'C:\c.exe' }
            New-Fake @{ Name = 'Zeta'; DisplayName = 'Z'; State = 'Stopped'; StartMode = 'Disabled'; DelayedAutoStart = $false; StartName = 'LocalSystem'; PathName = 'C:\z.exe' }
            if (Test-Scenario 'c') {
                # 管理者でないときに設定を読めないサービス。本物（LSM など）と同じく、表示名はサービス名、スタートアップは Unknown、ほかは空で返る
                New-Fake @{ Name = 'ZUnread'; DisplayName = 'ZUnread'; State = 'Stopped'; StartMode = 'Unknown'; DelayedAutoStart = $null; StartName = $null; PathName = $null }
                New-Fake @{ Name = 'LSM'; DisplayName = 'LSM'; State = 'Running'; StartMode = 'Unknown'; DelayedAutoStart = $null; StartName = $null; PathName = $null }
            }
        }
        'Win32_UserAccount' {
            New-Fake @{ Name = 'admin'; Disabled = $false; PasswordExpires = $false; Description = '管理者' }
            # 説明の改行は空白 1 つに、ほかの制御文字は ? にする
            New-Fake @{ Name = 'Guest'; Disabled = $true; PasswordExpires = $true; Description = "ゲスト`r`n説明`t付き" }
        }
        'Win32_Group' {
            New-Fake @{ Name = 'Administrators' }
            New-Fake @{ Name = 'Empty' }
            if (Test-Fail 'BadGroup') { New-Fake @{ Name = 'Bad' } }
        }
        'Win32_Share' {
            New-Fake @{ Name = 'DATA'; Path = 'D:\Data'; Type = [uint32]0; Description = 'データ' }
            New-Fake @{ Name = 'PRN'; Path = ''; Type = [uint32]1; Description = '' }
            New-Fake @{ Name = 'C$'; Path = 'C:\'; Type = [uint32]2147483648; Description = 'Default share' }
            New-Fake @{ Name = 'IPC$'; Path = ''; Type = [uint32]2147483651; Description = 'Remote IPC' }
            New-Fake @{ Name = '2026-10-31'; Path = 'D:\x'; Type = [uint32]0; Description = '' }
            New-Fake @{ Name = 'TMP'; Path = 'D:\t'; Type = [uint32]1073741824; Description = '' }
        }
        default { throw "知らないクラス: $ClassName" }
    }
}
function global:Get-FakeGroupMember([string]$GroupName) {
    if ($GroupName -eq 'Administrators') { 'WBTEST-PC\admin'; 'CORP\Domain Admins' }
    if ($GroupName -eq 'Bad') { throw (New-Object UnauthorizedAccessException '偽の拒否: Bad') }
}
function global:Get-Disk {
    param($ErrorAction)
    Assert-Fake 'Get-Disk'
    New-Fake @{ Number = 1; FriendlyName = 'Disk One'; Size = [uint64]214748364800; PartitionStyle = 'MBR'; BusType = 'SATA'; OperationalStatus = 'Offline' }
    New-Fake @{ Number = 0; FriendlyName = 'Disk Zero'; Size = [uint64]107374182400; PartitionStyle = 'GPT'; BusType = 'NVMe'; OperationalStatus = 'Online' }
}
function global:Get-Partition {
    param($ErrorAction)
    Assert-Fake 'Get-Partition'
    New-Fake @{ DiskNumber = 1; PartitionNumber = 1; Type = 'IFS'; IsActive = $true; Size = [uint64]214748364800; AccessPaths = @('D:\', 'C:\mnt\data\', '\\?\Volume{d}\') }
    # パーティションは、ボリュームの容量より少し大きい（容量はボリュームの値を書く）
    New-Fake @{ DiskNumber = 0; PartitionNumber = 3; Type = 'Basic'; IsActive = $false; Size = [uint64]107375230976; AccessPaths = @('C:\', '\\?\Volume{c}\') }
    New-Fake @{ DiskNumber = 0; PartitionNumber = 1; Type = 'System'; IsActive = $false; Size = [uint64]104857600; AccessPaths = @('\\?\Volume{efi}\') }
    New-Fake @{ DiskNumber = 0; PartitionNumber = 2; Type = 'Reserved'; IsActive = $false; Size = [uint64]16777216; AccessPaths = $null }
    New-Fake @{ DiskNumber = 0; PartitionNumber = 4; Type = 'Recovery'; IsActive = $false; Size = [uint64]1073741824; AccessPaths = @('\\?\Volume{rec}\') }
}
function global:Get-NetAdapter {
    param($ErrorAction)
    Assert-Fake 'Get-NetAdapter'
    New-Fake @{ Name = 'Team'; InterfaceDescription = 'Multiplexor'; Status = 'Up'; LinkSpeed = '2 Gbps'; MacAddress = 'AA-BB-CC-00-11-33'; InterfaceIndex = 7 }
    New-Fake @{ Name = 'Ethernet1'; InterfaceDescription = 'Intel #2'; Status = 'Disconnected'; LinkSpeed = '0 bps'; MacAddress = '00-11-22-33-44-55'; InterfaceIndex = 6 }
    New-Fake @{ Name = 'Ethernet0'; InterfaceDescription = 'Intel(R) Ethernet'; Status = 'Up'; LinkSpeed = '1 Gbps'; MacAddress = 'AA-BB-CC-00-11-22'; InterfaceIndex = 5 }
}
function global:Get-NetIPAddress {
    param($PolicyStore, $ErrorAction)
    Assert-Fake 'Get-NetIPAddress'
    New-Fake @{ InterfaceIndex = 5; AddressFamily = 'IPv4'; IPAddress = '192.168.10.21'; PrefixLength = 24; SuffixOrigin = 'Manual' }
    New-Fake @{ InterfaceIndex = 5; AddressFamily = 'IPv4'; IPAddress = '192.168.9.5'; PrefixLength = 25; SuffixOrigin = 'Manual' }
    New-Fake @{ InterfaceIndex = 5; AddressFamily = 'IPv6'; IPAddress = 'fe80::1%5'; PrefixLength = 64; SuffixOrigin = 'Link' }
    New-Fake @{ InterfaceIndex = 5; AddressFamily = 'IPv6'; IPAddress = '2001:db8::10'; PrefixLength = 64; SuffixOrigin = 'Link' }
    New-Fake @{ InterfaceIndex = 5; AddressFamily = 'IPv6'; IPAddress = '2001:db8::abcd'; PrefixLength = 128; SuffixOrigin = 'Random' }
    foreach ($i in 1..5) { New-Fake @{ InterfaceIndex = 7; AddressFamily = 'IPv4'; IPAddress = ('172.16.0.{0}' -f $i); PrefixLength = 16; SuffixOrigin = 'Manual' } }
    New-Fake @{ InterfaceIndex = 1; AddressFamily = 'IPv4'; IPAddress = '127.0.0.1'; PrefixLength = 8; SuffixOrigin = 'WellKnown' }
}
function global:Get-NetIPInterface {
    param($ErrorAction)
    Assert-Fake 'Get-NetIPInterface'
    New-Fake @{ InterfaceIndex = 5; AddressFamily = 'IPv4'; Dhcp = 'Disabled'; NlMtu = [uint32]1500 }
    New-Fake @{ InterfaceIndex = 6; AddressFamily = 'IPv4'; Dhcp = 'Enabled'; NlMtu = [uint32]1500 }
    New-Fake @{ InterfaceIndex = 7; AddressFamily = 'IPv4'; Dhcp = 'Disabled'; NlMtu = [uint32]9000 }
    # IPv6 の設定は、IPv4 の DHCP と MTU に混ざらない
    New-Fake @{ InterfaceIndex = 5; AddressFamily = 'IPv6'; Dhcp = 'Enabled'; NlMtu = [uint32]1280 }
}
function global:Get-NetRoute {
    param($PolicyStore, $ErrorAction)
    Assert-Fake 'Get-NetRoute'
    if ($PolicyStore -eq 'PersistentStore') {
        New-Fake @{ DestinationPrefix = '10.20.0.0/16'; NextHop = '192.168.10.254'; InterfaceAlias = 'Ethernet0'; InterfaceIndex = 5; RouteMetric = [uint16]1000 }
        New-Fake @{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.10.1'; InterfaceAlias = 'Ethernet0'; InterfaceIndex = 5; RouteMetric = [uint16]0 }
        return
    }
    New-Fake @{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.10.254'; InterfaceAlias = 'Ethernet0'; InterfaceIndex = 5; RouteMetric = [uint16]0 }
    New-Fake @{ DestinationPrefix = '192.168.10.0/24'; NextHop = '0.0.0.0'; InterfaceAlias = 'Ethernet0'; InterfaceIndex = 5; RouteMetric = [uint16]256 }
    New-Fake @{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.10.1'; InterfaceAlias = 'Ethernet0'; InterfaceIndex = 5; RouteMetric = [uint16]0 }
    New-Fake @{ DestinationPrefix = '0.0.0.0/0'; NextHop = '0.0.0.0'; InterfaceAlias = 'Team'; InterfaceIndex = 7; RouteMetric = [uint16]0 }
}
function global:Get-DnsClientServerAddress {
    param($ErrorAction)
    Assert-Fake 'Get-DnsClientServerAddress'
    New-Fake @{ InterfaceIndex = 5; AddressFamily = [uint16]2; ServerAddresses = @('192.168.10.12', '192.168.10.11') }
    New-Fake @{ InterfaceIndex = 5; AddressFamily = [uint16]23; ServerAddresses = @('fec0:0:0:ffff::1') }
    New-Fake @{ InterfaceIndex = 7; AddressFamily = [uint16]2; ServerAddresses = @('10.0.0.4', '10.0.0.3', '10.0.0.2', '10.0.0.1') }
}
function global:Get-DnsClient {
    param($ErrorAction)
    Assert-Fake 'Get-DnsClient'
    New-Fake @{ InterfaceIndex = 5; ConnectionSpecificSuffix = 'corp.example.local'; RegisterThisConnectionsAddress = $true }
    New-Fake @{ InterfaceIndex = 6; ConnectionSpecificSuffix = ''; RegisterThisConnectionsAddress = $false }
}
function global:Get-NetAdapterBinding {
    param($ErrorAction)
    Assert-Fake 'Get-NetAdapterBinding'
    New-Fake @{ Name = 'Ethernet0'; ComponentID = 'ms_tcpip'; Enabled = $true }
    New-Fake @{ Name = 'Ethernet0'; ComponentID = 'ms_tcpip6'; Enabled = $true }
    New-Fake @{ Name = 'Ethernet1'; ComponentID = 'ms_tcpip6'; Enabled = $false }
}
function global:Get-NetConnectionProfile {
    param($ErrorAction)
    Assert-Fake 'Get-NetConnectionProfile'
    New-Fake @{ InterfaceIndex = 5; NetworkCategory = 'DomainAuthenticated' }
    New-Fake @{ InterfaceIndex = 7; NetworkCategory = 'Private' }
}
function global:Get-DnsClientGlobalSetting {
    param($ErrorAction)
    Assert-Fake 'Get-DnsClientGlobalSetting'
    if (Test-Scenario 'b') { New-Fake @{ SuffixSearchList = @() } } else { New-Fake @{ SuffixSearchList = @('corp.example.local', 'example.local') } }
}
# 管理者のときは、実際に効いている設定（ActiveStore）を読む。どちらを読んだかが分かるように、ドメインの値を変えておく
function global:Get-NetFirewallProfile {
    param($PolicyStore, $ErrorAction)
    Assert-Fake 'Get-NetFirewallProfile'
    New-Fake @{ Name = 'Domain'; Enabled = $(if ($PolicyStore -eq 'ActiveStore') { 'False' } else { 'True' }) }
    New-Fake @{ Name = 'Private'; Enabled = 'False' }
    New-Fake @{ Name = 'Public'; Enabled = 'NotConfigured' }
}
function global:Get-WindowsFeature {
    param($ErrorAction, $WarningAction)
    Assert-Fake 'Get-WindowsFeature'
    New-Fake @{ Name = 'Web-Server'; DisplayName = 'Web サーバー (IIS)'; FeatureType = 'Role'; Parent = $null; Installed = $true }
    New-Fake @{ Name = 'Web-WebServer'; DisplayName = 'Web サーバー'; FeatureType = 'Role Service'; Parent = 'Web-Server'; Installed = $true }
    New-Fake @{ Name = 'RSAT'; DisplayName = 'リモート サーバー管理ツール'; FeatureType = 'Feature'; Parent = $null; Installed = $false }
    New-Fake @{ Name = 'FS-SMB1'; DisplayName = 'SMB 1.0/CIFS ファイル共有のサポート'; FeatureType = 'Feature'; Parent = $null; Installed = $true }
}
& $Tool @toolArgs
exit $LASTEXITCODE
'@
    Write-Script $wrapper $wrapperCode

    #endregion

    #region クライアント（場面 a）

    Write-Section 'クライアント・ワークグループ（偽の環境）'
    $a = Invoke-Fake @{} 'a'
    Assert-Equal '終了コード' $a.Code '0'
    Assert-Equal '保存先のフォルダの名前' ([regex]::IsMatch((Split-Path -Leaf ([string]$a.Dir)), '^ServerInventory_[^_]+_\d{8}_\d{6}$')) 'True'
    $files = @(if ($a.Dir) { Get-ChildItem -LiteralPath $a.Dir -File | ForEach-Object { $_.Name } | Sort-Object })
    Assert-Equal '保存するファイル' ($files -join ',') ((@($Columns.Keys) + 'summary.txt' | Sort-Object) -join ',')
    foreach ($file in $Columns.Keys) {
        $path = [IO.Path]::Combine([string]$a.Dir, $file)
        $bytes = if ([IO.File]::Exists($path)) { [IO.File]::ReadAllBytes($path) } else { [byte[]]@() }
        $bom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $head = @(Get-RawLines $a $file)
        Assert-Equal ("$file の BOM と見出し") ('{0}/{1}' -f $bom, $(if ($head.Count -gt 0) { $head[0] } else { '' })) ('True/"' + ($Columns[$file] -join '","') + '"')
    }

    # parameters.csv の分類と項目の順（決まった項目の後ろに、スロットとページファイル）
    $expectedKeys = New-Object 'System.Collections.Generic.List[string]'
    foreach ($category in $FixedItems.Keys) {
        foreach ($name in $FixedItems[$category]) { $expectedKeys.Add($category + '|' + $name) }
        if ($category -eq 'メモリ') { foreach ($slot in 'スロット DIMM0', 'スロット DIMM1', 'スロット DIMM1 (2)') { $expectedKeys.Add('メモリ|' + $slot) } }
        if ($category -eq '仮想メモリ') { $expectedKeys.Add('仮想メモリ|ページファイル C:\pagefile.sys') }
    }
    $actualKeys = @(Read-Result $a 'parameters.csv' | ForEach-Object { $_.'分類' + '|' + (Get-Clean $_.'項目') })
    Assert-Equal 'parameters.csv の分類と項目の順' ($actualKeys -join ';') ($expectedKeys -join ';')

    $expectA = [ordered]@{
        'コンピューター|ホスト名' = 'WBTEST-PC'; 'コンピューター|完全修飾ドメイン名' = 'wbtest-pc'; 'コンピューター|ドメイン・ワークグループ' = 'WORKGROUP'
        'コンピューター|ドメインの役割' = 'スタンドアロンのワークステーション'; 'コンピューター|メーカー' = 'FakeMaker'; 'コンピューター|モデル' = 'Model 1'
        'コンピューター|シリアル番号' = 'SN-0001'; 'コンピューター|BIOS のバージョン' = '2212'; 'コンピューター|BIOS の日付' = '2022-12-13'
        'コンピューター|ファームウェア' = 'UEFI'; 'コンピューター|セキュアブート' = '有効'
        'OS|名前' = 'Microsoft Windows 11 Pro'; 'OS|エディション' = 'Professional'; 'OS|バージョン' = '24H2'; 'OS|ビルド' = '26100.2033'
        'OS|アーキテクチャ' = '64 ビット'; 'OS|インストールの種類' = 'クライアント'; 'OS|インストール日' = '2025-08-23 11:34:07'; 'OS|最終起動' = '2026-11-11 20:43:19'
        'OS|Windows フォルダ' = 'C:\WINDOWS'; 'OS|OS の言語' = 'ja-JP'; 'OS|システムロケール' = 'ja-JP'; 'OS|タイムゾーン' = '(UTC+09:00) 大阪、札幌、東京'
        'OS|ライセンス認証' = '認証済み（RETAIL）'
        'CPU|名前' = 'Intel(R) Xeon(R) Gold 6338 CPU @ 2.00GHz'; 'CPU|ソケット数' = '2'; 'CPU|コア数' = '8'; 'CPU|論理プロセッサ数' = '16'; 'CPU|最大クロック' = '2,000 MHz'
        'メモリ|実装メモリ' = '20.0 GB'; 'メモリ|使えるメモリ' = '19.9 GB'; 'メモリ|スロット DIMM0' = '8.0 GB、3,200 MHz、Crucial CT8G'
        'メモリ|スロット DIMM1 (2)' = '4.0 GB、2,666 MHz'
        '仮想メモリ|管理' = '自動'; '仮想メモリ|ページファイル C:\pagefile.sys' = '自動（現在 4,096 MB）'
        'メモリダンプ|種類' = '自動メモリダンプ'; 'メモリダンプ|保存先' = '%SystemRoot%\MEMORY.DMP'; 'メモリダンプ|自動的に再起動' = 'する'
        'リモートデスクトップ|接続' = '許可'; 'リモートデスクトップ|ネットワークレベル認証' = '不要'; 'リモートデスクトップ|ポート' = '3390'
        'ファイアウォール|プライベート' = '無効'; 'ファイアウォール|パブリック' = '未構成'
        'UAC|UAC' = '有効'; 'UAC|通知のレベル' = '既定（デスクトップを暗転しない）'
        '時刻同期|種類' = 'NTP'; '時刻同期|NTP サーバー' = 'ntp.nict.jp,0x9'
        'Windows Update|WSUS サーバー' = 'なし'; 'Windows Update|自動更新' = 'ポリシーなし'
        '電源|電源プラン' = 'High performance'; 'ネットワーク|DNS サフィックスの検索一覧' = 'corp.example.local, example.local'
    }
    foreach ($key in $expectA.Keys) {
        $parts = $key.Split('|')
        Assert-Equal ('a: ' + $key) (Get-Param $a $parts[0] $parts[1]) $expectA[$key]
    }
    # 管理者のときは実際に効いている設定（ActiveStore）、管理者でないときはこの PC の設定を読み、備考に書く
    if ($isAdmin) {
        Assert-Equal 'a: ファイアウォールは ActiveStore（管理者）' ('{0}/{1}' -f (Get-Param $a 'ファイアウォール' 'ドメイン'), (Get-Param $a 'ファイアウォール' 'ドメイン' '備考')) '無効/'
    } else {
        Assert-Equal 'a: ファイアウォールはこの PC の設定（管理者でない）' ('{0}/{1}' -f (Get-Param $a 'ファイアウォール' 'ドメイン'), (Get-Param $a 'ファイアウォール' 'ドメイン' '備考')) '有効/この PC の設定。グループポリシーの分を含みません'
    }
    Assert-Equal 'a: 時刻同期の備考' (Get-Param $a '時刻同期' 'NTP サーバー' '備考') 'この PC の設定'
    Assert-Equal 'a: 取れたものには備考を書かない' (@(Read-Result $a 'parameters.csv' | Where-Object { $_.'備考' -and $_.'分類' -notin 'ファイアウォール', '時刻同期' }).Count) '0'

    # Excel で数値や日付に変わる値だけを ="..." にする
    $rawParams = @(Get-RawLines $a 'parameters.csv') -join "`n"
    Assert-Equal 'a: BIOS のバージョン 2212 は ="..."' (Test-Contains $rawParams '"BIOS のバージョン","=""2212"""') 'True'
    Assert-Equal 'a: ソケット数は ="..."' (Test-Contains $rawParams '"ソケット数","=""2"""') 'True'
    Assert-Equal 'a: 文字の値は ="..." にしない' (Test-Contains $rawParams '"メーカー","FakeMaker"') 'True'

    Write-Section '一覧（偽の環境）'
    $disks = @(Read-Result $a 'disks.csv')
    Assert-Equal 'ディスク: 番号の順' (($disks | ForEach-Object { Get-RowText $_ $Columns['disks.csv'] }) -join ' / ') '0|Disk Zero|100.0|107,374,182,400|GPT|NVMe|オンライン / 1|Disk One|200.0|214,748,364,800|MBR|SATA|オフライン'
    $rawDisks = @(Get-RawLines $a 'disks.csv') -join "`n"
    Assert-Equal 'ディスク: 数値の列は ="..." にしない' (Test-Contains $rawDisks '"0","Disk Zero","100.0","107,374,182,400"') 'True'

    $volumes = @(Read-Result $a 'volumes.csv')
    Assert-Equal 'ボリューム: 件数' $volumes.Count '5'
    Assert-Equal 'ボリューム: EFI' (Get-RowText $volumes[0] $Columns['volumes.csv']) '0|1|EFI システム|||FAT32|0.1|104,857,600|0.1|67,108,864|36.0%|1,024'
    Assert-Equal 'ボリューム: MSR（ボリュームなし）' (Get-RowText $volumes[1] $Columns['volumes.csv']) '0|2|予約（MSR）||||0.0|16,777,216||||'
    Assert-Equal 'ボリューム: C:' (Get-RowText $volumes[2] $Columns['volumes.csv']) '0|3|基本|C:||NTFS|100.0|107,374,182,400|50.0|53,687,091,200|50.0%|4,096'
    Assert-Equal 'ボリューム: 回復' (Get-RowText $volumes[3] $Columns['volumes.csv']) '0|4|回復|||NTFS|1.0|1,073,741,824|0.5|536,870,912|50.0%|4,096'
    Assert-Equal 'ボリューム: MBR のアクティブとマウント先' (Get-RowText $volumes[4] $Columns['volumes.csv']) '1|1|基本（アクティブ）|D:, C:\mnt\data|DATA|NTFS|200.0|214,748,364,800|150.0|161,061,273,600|25.0%|65,536'

    $net0 = Get-ListRow $a 'network.csv' '接続の名前' 'Ethernet0'
    $net1 = Get-ListRow $a 'network.csv' '接続の名前' 'Ethernet1'
    $netT = Get-ListRow $a 'network.csv' '接続の名前' 'Team'
    Assert-Equal 'ネットワーク: 名前の順' ((@(Read-Result $a 'network.csv') | ForEach-Object { $_.'接続の名前' }) -join ',') 'Ethernet0,Ethernet1,Team'
    Assert-Equal 'ネットワーク: Ethernet0' (Get-RowText $net0 $Columns['network.csv']) 'Ethernet0|Intel(R) Ethernet|接続|1 Gbps|AA-BB-CC-00-11-22|無効|192.168.9.5, 192.168.10.21|255.255.255.128, 255.255.255.0|192.168.10.1, 192.168.10.254|192.168.10.12, 192.168.10.11|有効|2001:db8::10|corp.example.local|する|無効|ドメイン|1,500'
    Assert-Equal 'ネットワーク: 切断中' (Get-RowText $net1 $Columns['network.csv']) 'Ethernet1|Intel #2|切断||00-11-22-33-44-55|有効|||||無効|||しない|||1,500'
    Assert-Equal 'ネットワーク: Team' (Get-RowText $netT $Columns['network.csv']) 'Team|Multiplexor|接続|2 Gbps|AA-BB-CC-00-11-33|無効|172.16.0.1, 172.16.0.2, 172.16.0.3, 172.16.0.4, 172.16.0.5|255.255.0.0, 255.255.0.0, 255.255.0.0, 255.255.0.0, 255.255.0.0||10.0.0.4, 10.0.0.3, 10.0.0.2, 10.0.0.1|||||既定（DHCP サーバーの設定に従う）|プライベート|9,000'
    $rawNet = @(Get-RawLines $a 'network.csv') -join "`n"
    Assert-Equal 'ネットワーク: 数字だけの MAC アドレスは ="..."' (Test-Contains $rawNet '"=""00-11-22-33-44-55"""') 'True'
    Assert-Equal 'ネットワーク: IPv6 の一時アドレスとリンクローカルを除く' ('{0}/{1}' -f (Test-Contains $rawNet '2001:db8::abcd'), (Test-Contains $rawNet 'fe80::')) 'False/False'

    $routes = @(Read-Result $a 'routes.csv')
    Assert-Equal '静的ルート' (($routes | ForEach-Object { Get-RowText $_ $Columns['routes.csv'] }) -join ' / ') '0.0.0.0/0|192.168.10.1|Ethernet0|0 / 10.20.0.0/16|192.168.10.254|Ethernet0|1,000'

    $features = @(Read-Result $a 'features.csv')
    Assert-Equal '機能（クライアント）' (($features | ForEach-Object { Get-RowText $_ $Columns['features.csv'] }) -join ' / ') 'Feat-A|Feature A|Windows の機能| / Feat-B|Feature B|Windows の機能|'

    $software = @(Read-Result $a 'software.csv')
    Assert-Equal 'ソフト: 件数（部品と更新と名前なしを除く）' $software.Count '2'
    Assert-Equal 'ソフト: 同じ名前のバージョンとインストール日をまとめる' (Get-RowText (Get-ListRow $a 'software.csv' '名前' 'App One') $Columns['software.csv']) 'App One|64 ビット|16.0, 16.1|Pub|2026-10-15, 2026-10-20'
    Assert-Equal 'ソフト: 32 ビット' (Get-RowText (Get-ListRow $a 'software.csv' '名前' '=Calc App') $Columns['software.csv']) '=Calc App|32 ビット|1.2.3||'
    Assert-Equal 'ソフト: = で始まる名前は ="..."' (Test-Contains (@(Get-RawLines $a 'software.csv') -join "`n") '"=""=Calc App"""') 'True'

    $hotfixes = @(Read-Result $a 'hotfixes.csv')
    Assert-Equal '更新プログラム: すべて・新しい順' (($hotfixes | ForEach-Object { Get-RowText $_ $Columns['hotfixes.csv'] }) -join ' / ') 'KB1|Security Update|2026-10-17|CORP\admin / KB3|Update|2026-10-17|NT AUTHORITY\SYSTEM / KB2|Update|2026-01-15|NT AUTHORITY\SYSTEM / KB0|Update||'

    $services = @(Read-Result $a 'services.csv')
    Assert-Equal 'サービス' (($services | ForEach-Object { Get-RowText $_ $Columns['services.csv'] }) -join ' / ') 'CDPUserSvc|Connected Devices|自動|| / Svc1|Service One|自動（遅延）|実行中|LocalSystem / TrigSvc|Trig|手動（トリガー）|停止|NT AUTHORITY\LocalService / Zeta|Z|無効|停止|LocalSystem'

    Assert-Equal 'ローカルユーザー' ((@(Read-Result $a 'localusers.csv') | ForEach-Object { Get-RowText $_ $Columns['localusers.csv'] }) -join ' / ') 'admin|有効|なし|管理者 / Guest|無効|あり|ゲスト 説明?付き'
    Assert-Equal 'ローカルグループ' ((@(Read-Result $a 'localgroups.csv') | ForEach-Object { Get-RowText $_ $Columns['localgroups.csv'] }) -join ' / ') 'Administrators|CORP\Domain Admins / Administrators|WBTEST-PC\admin'

    # 共有の種類（レビューの指摘: UInt32 の 0 と 1 がハッシュテーブルのキーに当たらず、数値のまま出ていた）
    Assert-Equal '共有フォルダ' ((@(Read-Result $a 'shares.csv') | ForEach-Object { Get-RowText $_ $Columns['shares.csv'] }) -join ' / ') '2026-10-31|D:\x|ディスク| / C$|C:\|管理共有|Default share / DATA|D:\Data|ディスク|データ / IPC$||IPC|Remote IPC / PRN||プリンター| / TMP|D:\t|1073741824|'

    Assert-Equal '取れなかったものはない' (@(Read-Result $a 'skipped.csv').Count) '0'
    $allText = (@(if ($a.Dir) { Get-ChildItem -LiteralPath $a.Dir -File | ForEach-Object { [IO.File]::ReadAllText($_.FullName) } }) -join "`n") + $a.Log
    Assert-Equal 'サービスの引数（パスワード）を書かない' (Test-Contains $allText 'SECRETPW') 'False'

    Write-Section '画面（偽の環境）'
    $screenExpect = @(
        @('ホスト名（FQDN が同じなら名前だけ）', 'ホスト名          WBTEST-PC')
        @('ワークグループの表示', 'ワークグループ    WORKGROUP（スタンドアロンのワークステーション）')
        @('メーカー・モデル', 'FakeMaker  Model 1')
        @('OS', 'Microsoft Windows 11 Pro（24H2、ビルド 26100.2033、64 ビット）')
        @('CPU', 'Intel(R) Xeon(R) Gold 6338 CPU @ 2.00GHz（2 ソケット・8 コア・16 論理プロセッサ）')
        @('メモリ', '20.0 GB（8.0 GB × 2、4.0 GB × 1）')
        @('ディスク', '0  100.0 GB  GPT   Disk Zero')
        @('ボリューム（マウント先も）', 'D:, C:\mnt\data  DATA    NTFS')
        @('ネットワーク: IPv4 は数の順・30 文字を超える分は「ほか N 件」', '192.168.9.5/25 ほか 1 件')
        @('ネットワーク: 多い値は「ほか N 件」', '172.16.0.1/16 ほか 4 件')
        @('ネットワーク: 30 文字に収まる値は並べる', '192.168.10.1, 192.168.10.254')
        @('リモートデスクトップ', '許可（ネットワークレベル認証: 不要、ポート 3390）')
        @('時刻同期', 'NTP（ntp.nict.jp,0x9）')
        @('Windows Update', 'WSUS: なし  自動更新: ポリシーなし')
        @('電源プラン', 'High performance')
        @('ライセンス認証', '認証済み（RETAIL）')
        @('更新プログラムの最新', '4（最新 KB1  2026-10-17）')
        @('保存先', ('保存先: ' + $a.Dir))
    )
    foreach ($case in $screenExpect) { Assert-Equal ('画面: ' + $case[0]) (Test-Contains $a.Log $case[1]) 'True' }
    Assert-Equal '画面: 取れなかったものがなければ注意を出さない' ('{0}/{1}' -f (Test-Contains $a.Log '※ 取れなかった項目'), (Test-Contains $a.Log $adminNote)) 'False/False'
    if ($isAdmin) {
        Assert-Equal '画面: ファイアウォール（管理者）' (Test-Contains $a.Log 'ドメイン: 無効  プライベート: 無効  パブリック: 未構成') 'True'
    } else {
        Assert-Equal '画面: ファイアウォール（管理者でない）' (Test-Contains $a.Log 'ドメイン: 有効  プライベート: 無効  パブリック: 未構成（グループポリシーの分を含みません）') 'True'
    }
    Assert-Equal 'summary.txt は画面と同じ内容' (Test-Contains (ConvertTo-FlatText $a.Summary) 'ホスト名          WBTEST-PC') 'True'
    # 120 桁の画面で折り返さない（保存先の行を除く。管理者でないときのファイアウォールの行を含む）
    $wide = @($a.Summary -split '\r?\n' | Where-Object { -not $_.StartsWith('保存先') -and (Get-Width $_) -gt 120 })
    Assert-Equal '画面の行は 120 桁以内' ($wide -join ' / ') ''

    #endregion

    #region サーバー（場面 b）

    Write-Section 'サーバー・ドメイン（偽の環境）'
    $b = Invoke-Fake @{} 'b'
    Assert-Equal '終了コード' $b.Code '0'
    $evalDate = (Get-Date).AddMinutes(43200).ToString('yyyy-MM-dd')
    $expectB = [ordered]@{
        'コンピューター|ホスト名' = 'SRV01'; 'コンピューター|完全修飾ドメイン名' = 'srv01.corp.example.local'; 'コンピューター|ドメイン・ワークグループ' = 'corp.example.local'
        'コンピューター|ドメインの役割' = 'メンバーサーバー'; 'コンピューター|ファームウェア' = 'BIOS'; 'コンピューター|セキュアブート' = '対象外（BIOS）'
        'OS|名前' = 'Microsoft Windows Server 2016 Standard'; 'OS|エディション' = 'ServerStandard'; 'OS|バージョン' = '1607'; 'OS|ビルド' = '14393.7428'
        'OS|アーキテクチャ' = '64 ビット'; 'OS|インストールの種類' = 'Server Core'; 'OS|OS の言語' = 'en-US'; 'OS|システムロケール' = 'en-US'
        'OS|ライセンス認証' = ('認証済み（TIMEBASED_EVAL、期限 {0}）' -f $evalDate)
        'CPU|ソケット数' = '1'; 'CPU|最大クロック' = '2,400 MHz'
        'メモリ|実装メモリ' = ''; 'メモリ|使えるメモリ' = '15.9 GB'
        '仮想メモリ|管理' = '手動'; '仮想メモリ|ページファイル C:\pagefile.sys' = '初期 4,096 MB、最大 8,192 MB（現在 4,096 MB）'
        '仮想メモリ|ページファイル D:\pagefile.sys' = 'システム管理サイズ'; '仮想メモリ|ページファイル E:\pagefile.sys' = 'システム管理サイズ'
        '仮想メモリ|ページファイル F:\pagefile.sys' = '設定なし（現在 1,024 MB）'
        'メモリダンプ|種類' = 'アクティブメモリダンプ'; 'メモリダンプ|自動的に再起動' = 'しない'
        'リモートデスクトップ|接続' = '許可しない'; 'リモートデスクトップ|ネットワークレベル認証' = '必要'; 'リモートデスクトップ|ポート' = '3389'
        'UAC|通知のレベル' = 'カスタム（ConsentPromptBehaviorAdmin=2、PromptOnSecureDesktop=0）'
        '時刻同期|種類' = 'ドメインの階層（NT5DS）'; '時刻同期|NTP サーバー' = 'time.windows.com,0x9'
        'Windows Update|WSUS サーバー' = 'http://wsus01:8530'; 'Windows Update|自動更新' = '自動ダウンロードしてインストールを予約（毎日 3:00）'
        '電源|電源プラン' = '381b4222-f694-41f0-9685-ff5bb260df2e'; 'ネットワーク|DNS サフィックスの検索一覧' = 'なし'
    }
    foreach ($key in $expectB.Keys) {
        $parts = $key.Split('|')
        Assert-Equal ('b: ' + $key) (Get-Param $b $parts[0] $parts[1]) $expectB[$key]
    }
    # 自動の印（?:\pagefile.sys）は行にしない。設定とページファイルの使用状況から、過不足なく 1 行ずつ
    Assert-Equal 'b: 仮想メモリの行' ((@(Read-Result $b 'parameters.csv' | Where-Object { $_.'分類' -eq '仮想メモリ' } | ForEach-Object { Get-Clean $_.'項目' })) -join ',') '管理,ページファイル C:\pagefile.sys,ページファイル D:\pagefile.sys,ページファイル E:\pagefile.sys,ページファイル F:\pagefile.sys'
    $notesB = [ordered]@{
        'メモリ|実装メモリ' = 'メモリのスロットの情報がありません'
        'リモートデスクトップ|接続' = 'グループポリシー'
        '時刻同期|NTP サーバー' = 'この PC の設定。種類が NT5DS のため使われません'
        'Windows Update|WSUS サーバー' = 'グループポリシー'; 'Windows Update|自動更新' = 'グループポリシー'
        '電源|電源プラン' = '電源プランの名前が分かりません'
    }
    foreach ($key in $notesB.Keys) {
        $parts = $key.Split('|')
        Assert-Equal ('b の備考: ' + $key) (Get-Param $b $parts[0] $parts[1] '備考') $notesB[$key]
    }
    Assert-Equal 'b: 役割と機能（インストール済みだけ・階層の順）' ((@(Read-Result $b 'features.csv') | ForEach-Object { Get-RowText $_ $Columns['features.csv'] }) -join ' / ') 'Web-Server|Web サーバー (IIS)|役割| / Web-WebServer|Web サーバー|役割サービス|Web-Server / FS-SMB1|SMB 1.0/CIFS ファイル共有のサポート|機能|'
    $bScreen = @(
        @('ホスト名と FQDN', 'SRV01（srv01.corp.example.local）')
        @('ドメイン', 'ドメイン          corp.example.local（メンバーサーバー）')
        @('OS（ReleaseId）', 'Microsoft Windows Server 2016 Standard（1607、ビルド 14393.7428、64 ビット）')
        @('メモリ（スロットなし）', 'メモリ            15.9 GB')
        @('リモートデスクトップ', 'リモートデスクトップ  許可しない')
        @('時刻同期（NT5DS は NTP サーバーを出さない）', '時刻同期              ドメインの階層（NT5DS）')
        @('Windows Update', 'WSUS: http://wsus01:8530  自動更新: 自動ダウンロードしてインストールを予約（毎日 3:00）')
        @('ライセンス認証', ('認証済み（TIMEBASED_EVAL、期限 {0}）' -f $evalDate))
    )
    foreach ($case in $bScreen) { Assert-Equal ('b の画面: ' + $case[0]) (Test-Contains $b.Log $case[1]) 'True' }
    Assert-Equal 'b の画面: NT5DS の NTP サーバーは出さない' (Test-Contains $b.Log 'time.windows.com') 'False'

    Write-Section 'ドメインコントローラー（偽の環境）'
    $dc = Invoke-Fake @{} 'b' @{ FAKE_DC = '1' }
    Assert-Equal 'DC: 終了コード' $dc.Code '0'
    Assert-Equal 'DC: ドメインの役割' (Get-Param $dc 'コンピューター' 'ドメインの役割') 'ドメインコントローラー'
    Assert-Equal 'DC: ローカルユーザーとグループは見出しだけ' ('{0}/{1}' -f @(Read-Result $dc 'localusers.csv').Count, @(Read-Result $dc 'localgroups.csv').Count) '0/0'
    Assert-Equal 'DC: skipped.csv に対象外' ((@(Read-Result $dc 'skipped.csv') | ForEach-Object { '{0}|{1}' -f $_.'理由', $_.'項目' }) -join ',') '対象外|ローカルユーザー,対象外|ローカルグループ'
    Assert-Equal 'DC: 画面の件数は「対象外」' (Test-Contains $dc.Log 'ローカルユーザー        対象外') 'True'
    Assert-Equal 'DC: 対象外は取れなかった項目に数えない' (Test-Contains $dc.Log '※ 取れなかった項目') 'False'
    Assert-Equal 'DC: 役割と機能は Server として取る' (@(Read-Result $dc 'features.csv').Count) '3'

    #endregion

    #region 端のケース（場面 c と d）

    Write-Section '設定の端のケース（偽の環境）'
    $c = Invoke-Fake @{} 'c'
    Assert-Equal 'c: 終了コード' $c.Code '0'
    $expectC = [ordered]@{
        'コンピューター|ファームウェア' = '|ファームウェアの種類が分かりません（3）'
        'コンピューター|セキュアブート' = '|ファームウェアの種類が分かりません'
        'OS|ビルド' = '26100|'
        'OS|ライセンス認証' = '|Windows のライセンスの情報が見つかりません'
        '仮想メモリ|管理' = '手動|'
        '仮想メモリ|ページファイル' = 'なし|'
        'メモリダンプ|種類' = 'なし|'
        'メモリダンプ|保存先' = '|'
        'メモリダンプ|自動的に再起動' = 'する|'
        'リモートデスクトップ|接続' = '許可しない|'
        'リモートデスクトップ|ネットワークレベル認証' = '必要|'
        'リモートデスクトップ|ポート' = '|ポートの設定がありません'
        'UAC|UAC' = '無効|'
        'UAC|通知のレベル' = '|UAC が無効です'
        '時刻同期|種類' = '同期しない（NoSync）|この PC の設定'
        '時刻同期|NTP サーバー' = 'gpo.ntp|グループポリシー。種類が NoSync のため使われません'
        'Windows Update|WSUS サーバー' = 'なし|WUServer（http://old:8530）はありますが、UseWUServer が 1 でないため使われません'
        'Windows Update|自動更新' = '無効|グループポリシー'
        '電源|電源プラン' = '|取れませんでした（その他のエラー）'
    }
    foreach ($key in $expectC.Keys) {
        $parts = $key.Split('|')
        Assert-Equal ('c: ' + $key) ('{0}|{1}' -f (Get-Param $c $parts[0] $parts[1]), (Get-Param $c $parts[0] $parts[1] '備考')) $expectC[$key]
    }
    $cSkips = @(Read-Result $c 'skipped.csv')
    Assert-Equal 'c: powercfg の失敗を skipped.csv に残す' ((@($cSkips | Where-Object { $_.'項目' -ne 'サービス' } | ForEach-Object { '{0}|{1}|{2}' -f $_.'理由', $_.'項目', (Test-Contains $_.'詳細' 'powercfg が失敗しました（終了コード 5）') })) -join ',') 'その他のエラー|電源プラン|True'
    # 実機確認の指摘: 管理者でないと、Windows が設定を見せないサービスがあり、スタートアップの Unknown をそのまま書いていた
    $cServices = @(Read-Result $c 'services.csv' | Where-Object { 'LSM', 'ZUnread' -contains (Get-Clean $_.'サービス名') } | ForEach-Object { Get-RowText $_ $Columns['services.csv'] })
    Assert-Equal 'c: 設定を読めないサービスは「不明」、表示名と実行ユーザーは空欄' ($cServices -join ' / ') 'LSM||不明|実行中| / ZUnread||不明|停止|'
    Assert-Equal 'c: 設定を読めないサービスを skipped.csv に残す（名前の順）' ((@($cSkips | Where-Object { $_.'項目' -eq 'サービス' } | ForEach-Object { '{0}|{1}' -f $_.'理由', (Get-Clean $_.'詳細') })) -join ' / ') '権限なし|LSM: スタートアップ・表示名・実行ユーザーを読めません（サービスの設定を読む権限がありません） / 権限なし|ZUnread: スタートアップ・表示名・実行ユーザーを読めません（サービスの設定を読む権限がありません）'
    Assert-Equal 'c: 画面の注意' (Test-Contains $c.Log '※ 取れなかった項目が 3 件あります（skipped.csv を確認してください）') 'True'
    Assert-Equal 'c: 画面: 管理者の注意（管理者でないときだけ）' (Test-Contains $c.Log $adminNote) ([string](-not $isAdmin))
    $cServiceCount = [regex]::Match([string]$c.RawLog, '(?m)^ サービス +([0-9,]+)\r?$')
    Assert-Equal 'c: 画面: サービスの件数（読めないものも数える）' $(if ($cServiceCount.Success) { $cServiceCount.Groups[1].Value } else { '（行がない）' }) @(Read-Result $c 'services.csv').Count

    $d = Invoke-Fake @{} 'd'
    Assert-Equal 'd: 終了コード' $d.Code '0'
    $expectD = [ordered]@{
        'コンピューター|セキュアブート' = '|セキュアブートの情報がありません'
        'OS|システムロケール' = '|システムロケールの情報がありません'
        'メモリダンプ|種類' = '|メモリダンプの設定がありません'
        '時刻同期|種類' = '|時刻同期の設定がありません'
        '時刻同期|NTP サーバー' = '|時刻同期の設定がありません'
        'Windows Update|自動更新' = 'ポリシーなし|'
    }
    foreach ($key in $expectD.Keys) {
        $parts = $key.Split('|')
        Assert-Equal ('d: ' + $key) ('{0}|{1}' -f (Get-Param $d $parts[0] $parts[1]), (Get-Param $d $parts[0] $parts[1] '備考')) $expectD[$key]
    }
    Assert-Equal 'd: ソフトがなくても見出しの CSV を作る' ('{0}/{1}' -f @(Read-Result $d 'software.csv').Count, [IO.File]::Exists([IO.Path]::Combine([string]$d.Dir, 'software.csv'))) '0/True'

    #endregion

    #region 取れないとき

    Write-Section '取れないとき（偽の環境）'
    $failEnv = @{ FAKE_DENY = 'Win32_BIOS'; FAKE_FAIL = 'Get-Disk,Get-DnsClient,BadGroup' }
    $f = Invoke-Fake @{} 'a' $failEnv
    Assert-Equal '取れないものがあっても終了コードは 0' $f.Code '0'
    $fSkips = @(Read-Result $f 'skipped.csv' | ForEach-Object { '{0}|{1}|{2}' -f $_.'理由', $_.'項目', (Get-Clean $_.'詳細') })
    Assert-Equal 'skipped.csv' ($fSkips -join ' / ') '権限なし|BIOS|偽の拒否: Win32_BIOS / その他のエラー|ディスク|偽の失敗: Get-Disk / その他のエラー|ネットワーク|Get-DnsClient: 偽の失敗: Get-DnsClient / 権限なし|ローカルグループ|Bad: 偽の拒否: Bad'
    Assert-Equal '取れなかった値は空欄にして備考に理由' ('{0}|{1}' -f (Get-Param $f 'コンピューター' 'シリアル番号'), (Get-Param $f 'コンピューター' 'シリアル番号' '備考')) '|取れませんでした（権限なし）'
    # レビューの指摘: 1 つ目の単位の項目の一覧が配列の配列になり、取れた値（ホスト名）を空欄で上書きしていた
    Assert-Equal '取れた値は上書きしない' ('{0}|{1}' -f (Get-Param $f 'コンピューター' 'ホスト名'), (Get-Param $f 'メモリ' '使えるメモリ')) 'WBTEST-PC|19.9 GB'
    Assert-Equal '取れなかった一覧は見出しだけ' ('{0}/{1}' -f @(Read-Result $f 'disks.csv').Count, @(Get-RawLines $f 'disks.csv').Count) '0/1'
    Assert-Equal '一部を読めなくても、ほかは出す（ネットワーク）' ('{0}/{1}' -f @(Read-Result $f 'network.csv').Count, (Get-ListRow $f 'network.csv' '接続の名前' 'Ethernet0').'DNS サフィックス') '3/'
    Assert-Equal '一部を読めなくても、ほかは出す（グループ）' @(Read-Result $f 'localgroups.csv').Count '2'
    Assert-Equal '画面: 取れなかった一覧' (Test-Contains $f.Log '（取れませんでした: その他のエラー）') 'True'
    Assert-Equal '画面: 取れなかった値' (Test-Contains $f.Log '■ ディスク    （取れませんでした: その他のエラー）') 'True'
    Assert-Equal '画面: 件数の注意' (Test-Contains $f.Log '※ 取れなかった項目が 4 件あります（skipped.csv を確認してください）') 'True'
    Assert-Equal '画面: 管理者の注意（管理者でないときだけ）' (Test-Contains $f.Log $adminNote) ([string](-not $isAdmin))

    $fNoSave = Invoke-Fake @{ NoSave = $true } 'a' $failEnv
    Assert-Equal '-NoSave: 終了コード' $fNoSave.Code '0'
    Assert-Equal '-NoSave: 理由の一覧を画面に出す' ('{0}/{1}/{2}' -f (Test-Contains $fNoSave.Log '■ 取れなかったもの'), (Test-Contains $fNoSave.Log '下の一覧を確認してください'), (Test-Contains $fNoSave.Log 'Get-DnsClient: 偽の失敗')) 'True/True/True'
    Assert-Equal '-NoSave: 保存先の行を出さない' (Test-Contains $fNoSave.Log '保存先:') 'False'

    $all = Invoke-Fake @{} 'a' @{ FAKE_FAIL = 'Win32_ComputerSystem,Win32_OperatingSystem,Get-NetAdapter' }
    Assert-Equal 'コンピューターと OS が取れない: 終了コード' $all.Code '0'
    Assert-Equal 'コンピューターと OS が取れない: 画面' ('{0}/{1}' -f (Test-Contains $all.Log 'ホスト名          （取れませんでした）'), (Test-Contains $all.Log 'OS                （取れませんでした）')) 'True/True'
    Assert-Equal 'コンピューターと OS が取れない: 役割と機能も取れない' (@(Read-Result $all 'skipped.csv' | Where-Object { $_.'項目' -eq '役割と機能' }).Count) '1'

    #endregion

    #region 保存の失敗と想定外のエラー

    Write-Section '保存の失敗と想定外のエラー'
    $saveFailTool = New-PatchedTool 'savefail' ($fakePairs + , @("            Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'skipped.csv')) `$SkipColumns `$skipRows", "            throw 'disk full'"))
    $saveFail = Invoke-Fake @{} 'a' @{ FAKE_FAIL = 'Get-Disk' } $saveFailTool
    Assert-Equal '保存できない: 警告' (Test-Contains $saveFail.Log '結果を保存できませんでした: disk full') 'True'
    Assert-Equal '保存できない: 作ったフォルダを消す' ([string]$saveFail.Dir) ''
    Assert-Equal '保存できない: 理由を画面に出す' ('{0}/{1}' -f (Test-Contains $saveFail.Log '■ 取れなかったもの'), (Test-Contains $saveFail.Log '下の一覧を確認してください')) 'True/True'
    Assert-Equal '保存できない: 終了コード' $saveFail.Code '1'

    $crashTool = New-PatchedTool 'crash' ($fakePairs + , @('    $screen = Get-ScreenLines', "    throw 'unexpected'"))
    $crash = Invoke-Fake @{} 'a' @{} $crashTool
    Assert-Equal '想定外のエラー: 表示' (Test-Contains $crash.Log 'エラーで止まりました。結果は保存していません。') 'True'
    Assert-Equal '想定外のエラー: 作ったフォルダを消す' ([string]$crash.Dir) ''
    Assert-Equal '想定外のエラー: 終了コード' ($crash.Code -ne 0) 'True'

    #endregion

    #region 止まるとき

    Write-Section '止まるとき'
    $both = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -NoSave -OutputFolder "$workRoot\unused" 2>&1 | Out-String
    Assert-Equal '-NoSave と -OutputFolder' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains (ConvertTo-FlatText $both) '-NoSave のときは、-OutputFolder は指定できません')) '1/True'
    Assert-Equal '止まったときは保存先を作らない' (Test-Path -LiteralPath "$workRoot\unused") 'False'

    # コマンドプロンプトで、途中の引数が "...\" だと、後ろの引数まで巻き込む
    $cmdFile = "$workRoot\quote.cmd"
    Write-Text $cmdFile ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -OutputFolder ""{2}\out_q\"" -NoSave`r`nexit /b %ERRORLEVEL%`r`n" -f $Exe, $tool, $workRoot)
    $quote = ConvertTo-FlatText (& cmd.exe /c $cmdFile 2>&1 | Out-String)
    Assert-Equal '\" で引数を巻き込む: 終了コード' $LASTEXITCODE '1'
    Assert-Equal '\" で引数を巻き込む: ヒント' (Test-Contains $quote '-OutputFolder の値に引用符（"）が残っています') 'True'
    Assert-Equal '\" で引数を巻き込む: 保存先を作らない' (Test-Path -LiteralPath "$workRoot\out_q") 'False'

    # 制約付き言語モードでは、最初に止まる
    $clmCommand = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '{0}' -OutputFolder '{1}'; exit `$LASTEXITCODE" -f $tool, "$workRoot\clm"
    $clm = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -Command $clmCommand 2>&1 | Out-String)
    $clmCode = $LASTEXITCODE
    Assert-Equal '制約付き言語モード: 止まる' ('{0}/{1}' -f $clmCode, (Test-Contains $clm '制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）')) '1/True'
    Assert-Equal '制約付き言語モード: 保存先を作らない' (Test-Path -LiteralPath "$workRoot\clm") 'False'

    # 保存先に書き込めないときは、集める前に止まる（フォルダの代わりにファイルがある、アクセス拒否、ドライブがない）
    Write-Text "$workRoot\notdir" 'x'
    $notDir = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -OutputFolder "$workRoot\notdir" 2>&1 | Out-String)
    Assert-Equal '保存先がファイル' ('{0}/{1}/{2}' -f $LASTEXITCODE, (Test-Contains $notDir ('保存先に書き込めません: {0}\notdir' -f $workRoot)), (Test-Contains $notDir '-NoSave を付けると、画面に表示だけします')) '1/True/True'
    Assert-Equal '保存先がファイル: 集め始めない' (Test-Contains $notDir '■ コンピューター') 'False'

    $denyDir = Join-Path $workRoot 'denied'
    [void][IO.Directory]::CreateDirectory($denyDir)
    $me = [Security.Principal.WindowsIdentity]::GetCurrent().User
    $denyRule = New-Object Security.AccessControl.FileSystemAccessRule($me, 'CreateDirectories,CreateFiles', 'ContainerInherit,ObjectInherit', 'None', 'Deny')
    $acl = Get-Acl -LiteralPath $denyDir
    $acl.AddAccessRule($denyRule)
    Set-Acl -LiteralPath $denyDir -AclObject $acl
    $denied = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -OutputFolder $denyDir 2>&1 | Out-String)
    Assert-Equal '保存先がアクセス拒否' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $denied ('保存先に書き込めません: {0}' -f $denyDir))) '1/True'
    $acl = Get-Acl -LiteralPath $denyDir
    [void]$acl.RemoveAccessRule($denyRule)
    Set-Acl -LiteralPath $denyDir -AclObject $acl
    $denyRule = $null

    $usedLetters = @([IO.DriveInfo]::GetDrives() | ForEach-Object { $_.Name.Substring(0, 1).ToUpperInvariant() })
    $freeLetter = @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' | Where-Object { $usedLetters -notcontains $_ -and -not (Test-Path ($_ + ':\')) })[0]
    if ($freeLetter) {
        $noDrive = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -OutputFolder ('{0}:\inventory' -f $freeLetter) 2>&1 | Out-String)
        Assert-Equal 'ドライブがない' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $noDrive ('保存先に書き込めません: {0}:\inventory' -f $freeLetter))) '1/True'
        # 管理者のときだけ、ネットワークドライブが見えない理由を案内する
        Assert-Equal 'ドライブがない: 管理者のときは UNC の案内' (Test-Contains $noDrive '\\server\share の形で指定してください') ([string]$isAdmin)
    } else {
        Write-Skip 'ドライブがない' '空いているドライブ文字がないため'
    }

    #endregion

    #region パスと起動

    Write-Section 'パスと起動'
    # 保存先を指定しないときは、スクリプトと同じフォルダ
    $defaultRun = Invoke-Fake @{ OutputFolder = $null } 'a'
    $fakeDirs = @(Get-ChildItem -LiteralPath (Split-Path $fakeTool) -Directory -Filter 'ServerInventory_*')
    Assert-Equal '既定の保存先はスクリプトのフォルダ' ('{0}/{1}' -f $defaultRun.Code, $fakeDirs.Count) '0/1'

    $jpDir = Join-Path $workRoot '日本語 と スペース [b]'
    $jp = Invoke-Fake @{ OutputFolder = $jpDir } 'a'
    Assert-Equal '日本語・スペース・[ ] を含む保存先' ('{0}/{1}/{2}' -f $jp.Code, (@(Get-ChildItem -LiteralPath $jpDir -Directory -Filter 'ServerInventory_*').Count), (Test-Contains $jp.Log ('保存先: ' + $jpDir))) '0/1/True'

    $longOut = "$workRoot\long\" + (('d' * 50 + '\') * 5) + 'out'
    $long = Invoke-Fake @{ OutputFolder = $longOut } 'a'
    $longDirs = @()
    if ([IO.Directory]::Exists('\\?\' + $longOut)) { $longDirs = @([IO.Directory]::GetDirectories('\\?\' + $longOut, 'ServerInventory_*')) }
    Assert-Equal '260 文字を超える保存先' ('{0}/{1}' -f $long.Code, $longDirs.Count) '0/1'
    if ($longDirs.Count -gt 0) { Assert-Equal '260 文字を超える保存先: 14 個のファイル' (@([IO.Directory]::GetFiles($longDirs[0])).Count) '14' }

    # 名前に [ ] を含むフォルダから、相対パスで起動する（偽の環境の変数はコマンドプロンプトに引き継ぐ）
    [void][IO.Directory]::CreateDirectory("$workRoot\[br]")
    $brArgs = Join-Path $workRoot 'args_bracket.json'
    [IO.File]::WriteAllText($brArgs, (@{ OutputFolder = 'res' } | ConvertTo-Json), (New-Object Text.UTF8Encoding $true))
    $brBat = "$workRoot\bracket.cmd"
    Write-Text $brBat ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\[br]""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -Tool ""{3}"" -ArgFile ""{4}""`r`nexit /b %ERRORLEVEL%`r`n" -f $workRoot, $Exe, $wrapper, $fakeTool, $brArgs)
    foreach ($pair in @{ FAKE_SCENARIO = 'a'; FAKE_REG = ($regBase + '\a'); FAKE_POWERCFG = $powercfg['a']; FAKE_FIRMWARE = '2' }.GetEnumerator()) { [Environment]::SetEnvironmentVariable($pair.Key, $pair.Value) }
    [void](& cmd.exe /c $brBat 2>&1 | Out-String)
    $brCode = $LASTEXITCODE
    foreach ($name in 'FAKE_SCENARIO', 'FAKE_REG', 'FAKE_POWERCFG', 'FAKE_FIRMWARE') { [Environment]::SetEnvironmentVariable($name, $null) }
    Assert-Equal '[ ] を含むフォルダから相対パスで保存する' ('{0}/{1}' -f $brCode, @(Get-ChildItem -LiteralPath "$workRoot\[br]\res" -Directory -ErrorAction SilentlyContinue).Count) '0/1'

    # 地域設定で、数値と日時の区切り文字が変わらない
    $cultureProbe = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $cultureProbe)"
    if ($control -and $control -ne '1.5') {
        $fi = Invoke-Fake @{} 'a' @{ FAKE_CULTURE = 'fi-FI' }
        Assert-Equal 'fi-FI でも容量の小数点は .' ('{0}/{1}' -f (Get-Param $fi 'メモリ' '使えるメモリ'), (Get-Param $fi 'OS' 'インストール日')) '19.9 GB/2025-08-23 11:34:07'
        Assert-Equal 'fi-FI でもバイトの区切りは ,' (Get-RowText @(Read-Result $fi 'disks.csv')[0] $Columns['disks.csv']) '0|Disk Zero|100.0|107,374,182,400|GPT|NVMe|オンライン'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }

    #endregion

    #region 小さな処理

    Write-Section '小さな処理（ツールの関数をそのまま呼ぶ）'
    # ツールの関数を取り出して、テストの PowerShell（5.1 か 7）で呼ぶ
    $unitFile = Join-Path $workRoot 'unit.ps1'
    Write-Script $unitFile @'
param([string]$Tool)
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$Invariant  = [Globalization.CultureInfo]::InvariantCulture
$DateFormat = 'yyyy-MM-dd'
$CellWidth  = 30
$excelRisk  = New-Object Text.RegularExpressions.Regex ('^(?:[=+\-@]|[\d\s.,/:+\-]+$)')
$ast = [Management.Automation.Language.Parser]::ParseFile($Tool, [ref]$null, [ref]$null)
foreach ($fn in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($fn.Name -in 'ConvertTo-SubnetMask', 'Get-LicenseText', 'Get-OptionalValue', 'Format-ListCell', 'Get-DisplayWidth', 'Format-Invariant', 'Protect-ExcelText', 'Get-Ipv4SortKey') {
        . ([scriptblock]::Create($fn.Extent.Text))
    }
}
foreach ($prefix in 0, 1, 8, 16, 20, 24, 25, 31, 32) { 'mask{0}={1}' -f $prefix, (ConvertTo-SubnetMask $prefix) }
$now = [datetime]'2026-11-11 12:00:00'
function New-Product($Status, $Channel, $Grace) { [pscustomobject]@{ LicenseStatus = $Status; Description = ('Windows(R) Operating System, {0} channel' -f $Channel); GracePeriodRemaining = $Grace } }
'lic_kms=' + (Get-LicenseText (New-Product 1 'VOLUME_KMSCLIENT' 259200) $now)
'lic_eval=' + (Get-LicenseText (New-Product 1 'TIMEBASED_EVAL' 1440) $now)
'lic_evalnotice=' + (Get-LicenseText (New-Product 5 'TIMEBASED_EVAL' 0) $now)
'lic_grace=' + (Get-LicenseText (New-Product 2 'RETAIL' 2880) $now)
'lic_unknown=' + (Get-LicenseText ([pscustomobject]@{ LicenseStatus = 9; Description = 'other' }) $now)
'cell_fit=' + (Format-ListCell @('10.0.0.1', '10.0.0.2'))
'cell_more=' + (Format-ListCell @('192.168.100.101/24', '192.168.100.102/24', '192.168.100.103/24'))
'cell_first=' + (Format-ListCell @(('a' * 50), 'b'))
'cell_empty=' + (Format-ListCell @('', $null))
'excel=' + ((@('2212', '16.0', 'abc', '=x', '-1', '1.2.3', '2026-10-31', 'ＡＢ', '１２') | ForEach-Object { Protect-ExcelText $_ }) -join '|')
'sort=' + ((@('192.168.10.1', '192.168.9.1', '10.0.0.1', 'bad') | Sort-Object { Get-Ipv4SortKey $_ }) -join ',')
'END'
'@
    $unit = @{}
    & $Exe -NoProfile -ExecutionPolicy Bypass -File $unitFile -Tool $repoTool 2>&1 | ForEach-Object {
        $line = "$_"
        $eq = $line.IndexOf('=')
        if ($eq -gt 0) { $unit[$line.Substring(0, $eq)] = $line.Substring($eq + 1) } elseif ($line -eq 'END') { $unit['END'] = 'END' }
    }
    Assert-Equal '小さな処理: 最後まで動いた' $unit['END'] 'END'
    $masks = @{ 0 = '0.0.0.0'; 1 = '128.0.0.0'; 8 = '255.0.0.0'; 16 = '255.255.0.0'; 20 = '255.255.240.0'; 24 = '255.255.255.0'; 25 = '255.255.255.128'; 31 = '255.255.255.254'; 32 = '255.255.255.255' }
    foreach ($prefix in 0, 1, 8, 16, 20, 24, 25, 31, 32) { Assert-Equal ('サブネットマスク /{0}' -f $prefix) $unit[('mask{0}' -f $prefix)] $masks[$prefix] }
    Assert-Equal 'ライセンス: KMS は期限を書かない' $unit['lic_kms'] '認証済み（VOLUME_KMSCLIENT）'
    Assert-Equal 'ライセンス: 評価版は期限を書く' $unit['lic_eval'] '認証済み（TIMEBASED_EVAL、期限 2026-11-12）'
    Assert-Equal 'ライセンス: 期限切れの評価版' $unit['lic_evalnotice'] '通知（TIMEBASED_EVAL）'
    Assert-Equal 'ライセンス: 猶予期間は期限を書く' $unit['lic_grace'] '猶予期間（RETAIL、期限 2026-11-13）'
    Assert-Equal 'ライセンス: 知らない状態' $unit['lic_unknown'] '状態 9'
    Assert-Equal '画面の値: 収まる' $unit['cell_fit'] '10.0.0.1, 10.0.0.2'
    Assert-Equal '画面の値: 「ほか N 件」' $unit['cell_more'] '192.168.100.101/24 ほか 2 件'
    Assert-Equal '画面の値: 最初の 1 つは必ず出す' $unit['cell_first'] (('a' * 50) + ' ほか 1 件')
    Assert-Equal '画面の値: 空' $unit['cell_empty'] ''
    Assert-Equal 'Excel 向けの ="..."' $unit['excel'] '="2212"|="16.0"|abc|="=x"|="-1"|="1.2.3"|="2026-10-31"|ＡＢ|="１２"'
    Assert-Equal 'IPv4 は数の順' $unit['sort'] '10.0.0.1,192.168.9.1,192.168.10.1,bad'

    #endregion

    #region Ctrl+C

    Write-Section 'Ctrl+C で中断'
    # 集め始める前に長く待つコピーで止める（読むだけなので、途中で止めても何も壊れない）
    $slowTool = New-PatchedTool 'slow' @(, @('    foreach ($unit in $ParamUnits) { Invoke-ParamUnit $unit }', "    Start-Sleep -Seconds 60`r`n    foreach (`$unit in `$ParamUnits) { Invoke-ParamUnit `$unit }"))
    $stopFile = "$workRoot\stop.ps1"
    Write-Script $stopFile (@(
        'param($Tool, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'Start-Sleep -Milliseconds 3000'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
    ) -join "`r`n")
    # Stop() は、動いている最中に呼んでも取りこぼして最後まで動き続けることがあるので、何度か試す
    $stopped = $false
    for ($try = 1; $try -le 5 -and -not $stopped; $try++) {
        $stopResult = @{}
        & $Exe -NoProfile -ExecutionPolicy Bypass -File $stopFile -Tool $slowTool -OutRoot ("$workRoot\out_ctrlc_{0}" -f $try) 2>&1 | ForEach-Object {
            $line = "$_"
            $eq = $line.IndexOf('=')
            if ($eq -gt 0) { $stopResult[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
        }
        if ($stopResult['RUNNING'] -eq 'True' -and $stopResult['AFTER'] -eq '0') { $stopped = $true }
    }
    if ($stopped) {
        Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $stopResult['BEFORE'], $stopResult['AFTER']) '1/0'
        Assert-Equal '中断のメッセージ' $stopResult['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' 'Stop() が効かなかったため（5 回試した）'
    }

    #endregion

    #region 本物の状態

    Write-Section '本物の状態（読むだけ）'
    # コマンドプロンプトから、最後の引数を "...\" で終えて起動する（引用符を外して普通に動く）
    $realOut = Join-Path $workRoot 'real'
    $realCmd = "$workRoot\real.cmd"
    Write-Text $realCmd ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -OutputFolder ""{2}\""`r`nexit /b %ERRORLEVEL%`r`n" -f $Exe, $tool, $realOut)
    $realLog = ConvertTo-FlatText (& cmd.exe /c $realCmd 2>&1 | Out-String)
    Assert-Equal '本物: 終了コード（最後の \" は外す）' $LASTEXITCODE '0'
    $realDir = Get-RunDir $realOut
    $real = [pscustomobject]@{ Dir = $realDir; Log = $realLog }
    Assert-Equal '本物: 14 個のファイル' $(if ($realDir) { @(Get-ChildItem -LiteralPath $realDir -File).Count } else { 0 }) '14'
    $realKeys = @(Read-Result $real 'parameters.csv' | Where-Object { (Get-Clean $_.'項目') -notmatch '^(スロット|ページファイル)' } | ForEach-Object { $_.'分類' + '|' + (Get-Clean $_.'項目') })
    $fixedKeys = @(foreach ($category in $FixedItems.Keys) { foreach ($name in $FixedItems[$category]) { $category + '|' + $name } })
    Assert-Equal '本物: parameters.csv の決まった項目' ($realKeys -join ';') ($fixedKeys -join ';')
    Assert-Equal '本物: ホスト名' (Get-Param $real 'コンピューター' 'ホスト名') $env:COMPUTERNAME
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    Assert-Equal '本物: OS の名前' (Get-Param $real 'OS' '名前') ([string]$os.Caption).Trim()
    Assert-Equal '本物: ソケット数' (Get-Param $real 'CPU' 'ソケット数') ('{0:N0}' -f @(Get-CimInstance -ClassName Win32_Processor).Count)
    Assert-Equal '本物: その他のエラーはない' ((@(Read-Result $real 'skipped.csv' | Where-Object { $_.'理由' -eq 'その他のエラー' }) | ForEach-Object { '{0}: {1}' -f $_.'項目', $_.'詳細' }) -join ' / ') ''
    # 管理者でないときは、設定を読めないサービス（LSM など）がある。管理者なら 0 件のことが多いが、どちらでも件数が合うことを見る
    $unreadableServices = @(Get-CimInstance -ClassName Win32_Service | Where-Object { [string]$_.StartMode -eq 'Unknown' -or -not $_.StartMode } | ForEach-Object { [string]$_.Name } | Sort-Object)
    Assert-Equal '本物: 設定を読めないサービスを skipped.csv に残す' ((@(Read-Result $real 'skipped.csv' | Where-Object { $_.'項目' -eq 'サービス' }) | ForEach-Object { ([string](Get-Clean $_.'詳細')).Split(':')[0] }) -join ',') ($unreadableServices -join ',')
    Assert-Equal '本物: services.csv に Unknown を書かない' (@(Read-Result $real 'services.csv' | Where-Object { (Get-Clean $_.'スタートアップ') -eq 'Unknown' }).Count) '0'
    # レビューの指摘: CIM の一覧をそのままパイプに流すと、一覧全体が 1 件になっていた
    Assert-Equal '本物: 更新プログラムの件数' @(Read-Result $real 'hotfixes.csv').Count @(Get-CimInstance -ClassName Win32_QuickFixEngineering).Count
    $realShares = @(Get-CimInstance -ClassName Win32_Share)
    Assert-Equal '本物: 共有フォルダの件数' @(Read-Result $real 'shares.csv').Count $realShares.Count
    $ipcRow = Get-ListRow $real 'shares.csv' '共有名' 'IPC$'
    if ($ipcRow) { Assert-Equal '本物: IPC$ の種類' $ipcRow.'種類' 'IPC' } else { Write-Skip '本物: IPC$ の種類' 'IPC$ がないため' }
    $adminShare = @($realShares | Where-Object { [long]$_.Type -eq 2147483648 } | Select-Object -First 1)
    if ($adminShare.Count -gt 0) { Assert-Equal '本物: 管理共有の種類' (Get-ListRow $real 'shares.csv' '共有名' $adminShare[0].Name).'種類' '管理共有' } else { Write-Skip '本物: 管理共有の種類' '管理共有がないため' }
    $diskShare = @($realShares | Where-Object { [long]$_.Type -eq 0 } | Select-Object -First 1)
    if ($diskShare.Count -gt 0) { Assert-Equal '本物: 普通の共有の種類' (Get-ListRow $real 'shares.csv' '共有名' $diskShare[0].Name).'種類' 'ディスク' } else { Write-Skip '本物: 普通の共有の種類' '普通の共有がないため（偽の環境で確かめた）' }
    Assert-Equal '本物: ディスクの件数' @(Read-Result $real 'disks.csv').Count @(Get-Disk).Count
    Assert-Equal '本物: アダプターの件数' @(Read-Result $real 'network.csv').Count @(Get-NetAdapter).Count
    $role = [int](Get-CimInstance -ClassName Win32_ComputerSystem).DomainRole
    if ($role -lt 4) {
        Assert-Equal '本物: ローカルユーザーの件数' @(Read-Result $real 'localusers.csv').Count @(Get-CimInstance -ClassName Win32_UserAccount -Filter 'LocalAccount=TRUE').Count
    } else {
        Write-Skip '本物: ローカルユーザーの件数' 'ドメインコントローラーのため'
    }
    if ([int]$os.ProductType -eq 1) {
        Assert-Equal '本物: Windows の機能の件数' @(Read-Result $real 'features.csv').Count @(Get-CimInstance -ClassName Win32_OptionalFeature -Filter 'InstallState=1').Count
    } else {
        Write-Skip '本物: Windows の機能の件数' 'Windows Server のため（役割と機能は実機で確かめる）'
    }
    # IPv6 の一時アドレスは書かない
    $tempV6 = @(Get-NetIPAddress -AddressFamily IPv6 -ErrorAction SilentlyContinue | Where-Object { "$($_.SuffixOrigin)" -eq 'Random' } | ForEach-Object { [string]$_.IPAddress })
    if ($tempV6.Count -gt 0) {
        $netText = (@(Get-RawLines $real 'network.csv') -join "`n")
        Assert-Equal '本物: IPv6 の一時アドレスを書かない' (@($tempV6 | Where-Object { Test-Contains $netText $_ }).Count) '0'
    } else {
        Write-Skip '本物: IPv6 の一時アドレスを書かない' '一時アドレスがないため'
    }
    # ユーザーごとのサービスの実体（名前の末尾がログオンごとに変わる）は、末尾を除いた名前で書く
    $instances = @(Get-CimInstance -ClassName Win32_Service | Where-Object {
        $type = (Get-ItemProperty -LiteralPath ('HKLM:\SYSTEM\CurrentControlSet\Services\{0}' -f $_.Name) -Name Type -ErrorAction SilentlyContinue)
        $type -and (([int]$type.Type -band 0x80) -ne 0)
    } | ForEach-Object { [string]$_.Name })
    if ($instances.Count -gt 0) {
        $serviceNames = @(Read-Result $real 'services.csv' | ForEach-Object { Get-Clean $_.'サービス名' })
        Assert-Equal '本物: ユーザーごとのサービスの末尾を除く' (@($instances | Where-Object { $serviceNames -contains $_ }) -join ',') ''
    } else {
        Write-Skip '本物: ユーザーごとのサービスの末尾を除く' 'ユーザーごとのサービスがないため'
    }
    Assert-Equal '本物: 画面' ('{0}/{1}/{2}' -f (Test-Contains $realLog '■ コンピューター'), (Test-Contains $realLog '■ 一覧の件数'), (Test-Contains $realLog ('保存先: ' + $realDir))) 'True/True/True'

    # 保存しないときは、スクリプトのフォルダに何も作らない
    $before = @(Get-ChildItem -LiteralPath (Split-Path $tool) -Directory).Count
    $noSave = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -NoSave 2>&1 | Out-String)
    Assert-Equal '本物: -NoSave' ('{0}/{1}/{2}' -f $LASTEXITCODE, (@(Get-ChildItem -LiteralPath (Split-Path $tool) -Directory).Count - $before), (Test-Contains $noSave '■ コンピューター')) '0/0/True'

    #endregion
} catch {
    # 途中で止まったときも失敗に数える。コマンドレットのエラーは、catch がないと try の後ろの END まで進み、合格の数が減るだけで気づけない
    '  FAIL  テストが途中で止まりました: {0}' -f ($_.Exception.Message -replace '\s*[\r\n]+\s*', ' ')
    $script:failCount++
} finally {
    # アクセス拒否を残したまま止まったときは、先に戻す
    if ($denyRule -and $denyDir) {
        try { $acl = Get-Acl -LiteralPath $denyDir; [void]$acl.RemoveAccessRule($denyRule); Set-Acl -LiteralPath $denyDir -AclObject $acl } catch { }
    }
    try { [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($regBase, $false) } catch { }
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
