<#
.SYNOPSIS
    Get-CertificateExpiry の回帰テストです。

.DESCRIPTION
    実行したユーザーの CurrentUser に、テスト専用のストア（wbtest-<番号>）を作り、一時的な自己署名の証明書
    （期限切れ・期限間近・有効・更新済みなど）を入れて、Get-CertificateExpiry を実行して期待値と照合します。
    本物のストア（個人など）には書き込みません。読むだけです。秘密鍵は、ストアに「秘密鍵あり」の印だけを残し、
    鍵そのものは残しません。テスト専用のストアは、終了時に削除します。

    証明書ファイル（DER・PEM・途中で切れた PEM・秘密鍵・PFX など）、ループバックに立てた TLS サーバーへの接続、
    偽の applicationHost.config を使った IIS の紐付け、設定ファイル、止まる場合も確かめます。
    ダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダとファイルにだけアクセス拒否を設定します（終了時に戻します）。
    ドライブのルートの確認のため、空いているドライブ文字を subst で一時的に割り当てます（終了時に外します）。
    C:\Users\Default の互換用リンクと、管理共有（\\localhost\C$）を読み取ります。読めない環境では、そのテストを飛ばします。
    本物の環境では起きない条件（ストアの権限なし、HTTP.sys の登録、保存の失敗など）は、ツールの処理を書き換えたコピーで確かめます。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Get-CertificateExpiry を実行する PowerShell（powershell.exe または pwsh）。

.PARAMETER KeepTemp
    ダミー環境と結果を削除せずに残します（調査用）。テスト専用のストアは残しません。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Get-CertificateExpiry\Get-CertificateExpiry.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Get-CertificateExpiry_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
# 結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーしたスクリプトを動かす
$tool     = Join-Path $workRoot 'tool\Get-CertificateExpiry.ps1'
$runId    = [Guid]::NewGuid().ToString('N').Substring(0, 8)
$userSid  = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$isAdmin  = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$invariant = [Globalization.CultureInfo]::InvariantCulture
$denied   = New-Object 'System.Collections.Generic.List[string]'
$junctions = New-Object 'System.Collections.Generic.List[string]'
$substDrive = $null
$thumbs   = @{}
$rawOf    = @{}
$pfxOf    = @{}
$timeOf   = @{}
$fromOf   = @{}
$x509     ='Security.Cryptography.X509Certificates'
$keyFlags = [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet
# テスト専用のストア。本物のストアに書き込まないよう、名前に番号を付けて新しく作る
$storeName  = "wbtest-$runId"
$store2Name = "wbtest2-$runId"
$storeText  = "CurrentUser\$storeName"
$store2Text = "CurrentUser\$store2Name"
$store3Name = "wbtest3-$runId"
$userStore  = $null
$userStore2 = $null
$userStore3 = $null
$userStore3m = $null
$certKey    = $null
$server     = $null
$serverStop = Join-Path $workRoot 'server-stop.txt'
$savedWindir = $env:windir
$utf8NoBom  = New-Object Text.UTF8Encoding $false
$utf8Bom    = New-Object Text.UTF8Encoding $true
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

function Test-Contains([string]$Text, [string]$Part) { $Text.IndexOf($Part, [StringComparison]::Ordinal) -ge 0 }

function Get-RunDir([string]$Parent) {
    if (-not (Test-Path -LiteralPath $Parent)) { return $null }
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory | Where-Object { $_.Name -like 'CertificateExpiry_*' })
    if ($found.Count -gt 0) { return $found[0].FullName }
    $null
}

$script:runNumber = 0
# Log には、画面への出力から改行を取り除いたものを入れる
function Invoke-Tool([string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $script:runNumber++
    $out  = Join-Path $workRoot ("out_{0}" -f $script:runNumber)
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $dir  = Get-RunDir $out
    $certRows = @()
    $skipRows = @()
    $summary  = @()
    if ($dir) {
        $certRows = @(Import-Csv -LiteralPath (Join-Path $dir 'certificates.csv') -Encoding UTF8)
        $skipRows = @(Import-Csv -LiteralPath (Join-Path $dir 'skipped.csv') -Encoding UTF8)
        $summary  = @([IO.File]::ReadAllLines((Join-Path $dir 'summary.txt')))
    }
    [pscustomobject]@{ Code = $code; Log = $log; Dir = $dir; Certs = $certRows; Skipped = $skipRows; Summary = $summary }
}

# 保存しないで動かす（画面と終了コードだけを見る）
function Invoke-NoSave([string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -NoSave 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# -NoSave も -OutputFolder も足さずに動かす
function Invoke-Plain([string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 相対パスの基準を確かめるため、指定したフォルダから起動する
function Invoke-From([string]$WorkDir, [string[]]$ToolArgs) {
    Push-Location -LiteralPath $WorkDir
    try {
        $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1 | Out-String)
        $code = $LASTEXITCODE
    } finally {
        Pop-Location
    }
    [pscustomobject]@{ Code = $code; Log = $log }
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch = Join-Path $workRoot 'run.cmd'
    [IO.File]::WriteAllText($batch, ("@echo off`r`nchcp 65001 > nul`r`n{0}`r`nexit /b %errorlevel%`r`n" -f $CommandLine), $utf8NoBom)
    $log = ConvertTo-FlatText (& cmd.exe /d /c $batch 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 定数や処理を 1 か所だけ書き換えたツールのコピーを作る（本物の環境では起きない条件を確かめるため）
function New-PatchedTool([string]$Name, [string]$From, [string]$To) {
    $text = [IO.File]::ReadAllText($repoTool)
    if (-not $text.Contains($From)) { throw "書き換える場所が見つかりません: $From" }
    $file = Join-Path $workRoot "patched\$Name\Get-CertificateExpiry.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text.Replace($From, $To), $utf8Bom)
    $file
}

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $used -notcontains $_ -and -not [IO.Directory]::Exists("${_}:\") } | Select-Object -First 1
}

function Get-ThumbText($Row) { $Row.'拇印' -replace '^="', '' -replace '"$', '' }

# 作った証明書の行（名前で探す）
function Get-CertRow($Run, [string]$Name) { @($Run.Certs | Where-Object { (Get-ThumbText $_) -eq $thumbs[$Name] }) }

function Get-Status($Run, [string]$Name) { (@(Get-CertRow $Run $Name) | ForEach-Object { $_.'状態' }) -join '|' }

function Get-Cell($Run, [string]$Name, [string]$Column) {
    $rows = @(Get-CertRow $Run $Name)
    if ($rows.Count -eq 0) { return '(なし)' }
    $rows[0].$Column
}

function Get-PlaceRows($Run, [string]$Place) { @($Run.Certs | Where-Object { $_.'場所' -eq $Place }) }

function Get-PlaceStatus($Run, [string]$Place) { (@(Get-PlaceRows $Run $Place) | ForEach-Object { $_.'状態' }) -join '|' }

function Get-SkipRow($Run, [string]$Subject) { @($Run.Skipped | Where-Object { $_.'対象' -eq $Subject }) }

function Get-SkipReason($Run, [string]$Subject) { (@(Get-SkipRow $Run $Subject) | ForEach-Object { $_.'理由' }) -join '|' }

function Get-SkipDetail($Run, [string]$Subject) { (@(Get-SkipRow $Run $Subject) | ForEach-Object { $_.'詳細' }) -join '|' }

# summary.txt の、指定した文字で始まる行
function Get-SummaryLine($Run, [string]$Start) { @($Run.Summary | Where-Object { $_.StartsWith($Start, [StringComparison]::Ordinal) }) -join '|' }

function Get-CountLine($Run) { (Get-SummaryLine $Run '証明書 ') -replace '  所要 .*$', '' }

function Deny-Access([string]$TargetPath, [string]$Rights) {
    $null = & icacls.exe $TargetPath /deny "*${userSid}:$Rights" 2>&1
    $denied.Add($TargetPath)
}

function New-TestFile([string]$FilePath, [byte[]]$Bytes) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllBytes('\\?\' + $FilePath, $Bytes)
}

function New-TextFile([string]$FilePath, [string]$Text) { New-TestFile $FilePath ([Text.Encoding]::ASCII.GetBytes($Text)) }

function ConvertTo-Pem([byte[]]$Bytes, [string]$Label = 'CERTIFICATE') {
    "-----BEGIN $Label-----`r`n" + [Convert]::ToBase64String($Bytes, [Base64FormattingOptions]::InsertLineBreaks) + "`r`n-----END $Label-----`r`n"
}

function Test-FileBom([string]$FilePath) {
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
}

# 画面の表の幅（全角は 2 文字分）
function Get-Width([string]$Text) {
    $width = 0
    foreach ($ch in $Text.ToCharArray()) { if ([int]$ch -lt 0x100) { $width += 1 } else { $width += 2 } }
    $width
}

function Get-Pad([string]$Text, [int]$Width, [bool]$Right = $false) {
    $pad = ' ' * [math]::Max(0, $Width - (Get-Width $Text))
    if ($Right) { return $pad + $Text }
    $Text + $pad
}

# 表の 1 行の期待値
function Get-TableLine([string]$Status, [string]$Days, [datetime]$Time, [string]$Place, [int]$PlaceWidth, [string]$Text) {
    ' ' + (Get-Pad $Status 8) + '  ' + (Get-Pad $Days 8 $true) + '  ' + $Time.ToString('yyyy-MM-dd HH:mm', $invariant) + '  ' + (Get-Pad $Place $PlaceWidth) + '  ' + $Text
}

function Format-CsvTime([datetime]$Time) { $Time.ToString('yyyy/MM/dd HH:mm:ss', $invariant) }

# 自己署名の証明書を作る。$Before と $After は今からの日数
function New-TestCert([string]$Subject, [string[]]$Dns, [double]$Before, [double]$After, [string[]]$Ip = @(), [string[]]$Usage = @(), [string[]]$RawDns = @()) {
    $request = New-Object "$x509.CertificateRequest" ($Subject, $certKey, [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    if ($RawDns) {
        # SubjectAlternativeNameBuilder は名前を小文字にするので、大文字の名前は DER を自分で組み立てる（長さはどれも 127 バイト以下）
        $body = New-Object 'System.Collections.Generic.List[byte]'
        foreach ($name in $RawDns) {
            $body.Add(0x82)
            $body.Add([byte]$name.Length)
            $body.AddRange([Text.Encoding]::ASCII.GetBytes($name))
        }
        $der = [byte[]](@(0x30, $body.Count) + $body.ToArray())
        $request.CertificateExtensions.Add((New-Object "$x509.X509Extension" ('2.5.29.17', $der, $false)))
    } elseif ($Dns -or $Ip) {
        $san = New-Object "$x509.SubjectAlternativeNameBuilder"
        foreach ($name in $Dns) { $san.AddDnsName($name) }
        foreach ($address in $Ip) { $san.AddIpAddress([Net.IPAddress]::Parse($address)) }
        $request.CertificateExtensions.Add($san.Build())
    }
    if ($Usage) {
        $oids = New-Object Security.Cryptography.OidCollection
        foreach ($oid in $Usage) { [void]$oids.Add((New-Object Security.Cryptography.Oid $oid)) }
        $request.CertificateExtensions.Add((New-Object "$x509.X509EnhancedKeyUsageExtension" ($oids, $false)))
    }
    $request.CreateSelfSigned([DateTimeOffset]::Now.AddDays($Before), [DateTimeOffset]::Now.AddDays($After))
}

# 証明書を作ってストアに入れる。秘密鍵ありの証明書は、一時的な鍵として読み込んでストアに入れ、読み込んだ分を破棄する
# （ストアには「秘密鍵あり」の印だけが残り、鍵のファイルは残らない）
function Add-TestCert([string]$Name, [string]$Subject, [string[]]$Dns, [double]$Before, [double]$After, [hashtable]$Option = @{}) {
    $ipList = @()
    $usageList = @()
    $rawList = @()
    if ($Option.ContainsKey('Ip')) { $ipList = $Option.Ip }
    if ($Option.ContainsKey('Usage')) { $usageList = $Option.Usage }
    if ($Option.ContainsKey('RawDns')) { $rawList = $Option.RawDns }
    $made = New-TestCert $Subject $Dns $Before $After $ipList $usageList $rawList
    $thumbs[$Name] = $made.Thumbprint
    $rawOf[$Name]  = $made.RawData
    $timeOf[$Name] = $made.NotAfter
    $fromOf[$Name] = $made.NotBefore
    $pfxOf[$Name]  = $made.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pfx)
    if ($Option.ContainsKey('NoStore')) { return }
    $target = if ($Option.ContainsKey('Store2')) { $userStore2 } elseif ($Option.ContainsKey('Into')) { $Option.Into } else { $userStore }
    if ($Option.ContainsKey('Also')) { $Option.Also.Add((New-Object "$x509.X509Certificate2" (, [byte[]]$made.RawData))) }
    if ($Option.ContainsKey('Key')) {
        $entry = New-Object "$x509.X509Certificate2" ([byte[]]$pfxOf[$Name], [string]$null, $keyFlags)
    } else {
        $entry = New-Object "$x509.X509Certificate2" (, [byte[]]$made.RawData)
    }
    if ($Option.ContainsKey('Friendly')) { $entry.FriendlyName = $Option.Friendly }
    $target.Add($entry)
    $entry.Reset()
    $made.Reset()
}

# 偽の %windir% を作り、applicationHost.config を置く
function New-FakeWindir([string]$Name, [string]$HostConfig, [string]$RedirectConfig, [string]$SystemFolder = 'System32') {
    $fake = Join-Path $workRoot "windir_$Name"
    $config = Join-Path $fake "$SystemFolder\inetsrv\config"
    [void][IO.Directory]::CreateDirectory($config)
    if ($HostConfig) { [IO.File]::WriteAllText((Join-Path $config 'applicationHost.config'), $HostConfig) }
    if ($RedirectConfig) { [IO.File]::WriteAllText((Join-Path $config 'redirection.config'), $RedirectConfig) }
    $fake
}

function New-HostConfig([string]$SitesXml) {
    '<?xml version="1.0"?><configuration><system.applicationHost><sites>' + $SitesXml + '</sites></system.applicationHost></configuration>'
}

function Invoke-WithWindir([string]$Fake, [string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $env:windir = $Fake
    try { Invoke-Tool $ToolArgs $ToolPath } finally { $env:windir = $savedWindir }
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

try {
    #region ダミー環境

    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    $certKey = [Security.Cryptography.RSA]::Create(2048)
    $userStore = New-Object "$x509.X509Store" ($storeName, [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $userStore.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $userStore2 = New-Object "$x509.X509Store" ($store2Name, [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $userStore2.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    # LocalMachine に書き込めないので、$store3Name-m を LocalMachine の同じ名前のストアに見立てる
    $userStore3 = New-Object "$x509.X509Store" ($store3Name, [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $userStore3.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $userStore3m = New-Object "$x509.X509Store" ("$store3Name-m", [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $userStore3m.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)

    $serverAuth = '1.3.6.1.5.5.7.3.1'
    # 状態と残り日数（日数に .5 を足して、実行にかかる時間で日数の表示が変わらないようにする）
    Add-TestCert 'expired' 'CN=wb-expired' @('wb-expired.example.jp') -400 -9.5
    Add-TestCert 'near' 'CN=wb-near' @('wb-near.example.jp') -300 20.5 @{ Key = $true; Usage = @($serverAuth) }
    Add-TestCert 'valid' 'CN=wb-valid' @('wb-valid.example.jp', 'www.wb-valid.example.jp') -10 300.5 @{ Key = $true; Friendly = '-test"x' }
    Add-TestCert 'future' 'CN=wb-future' @('wb-future.example.jp') 10 400
    Add-TestCert 'at60' 'CN=wb-at60' @('wb-at60.example.jp') -10 60.5
    Add-TestCert 'over60' 'CN=wb-over60' @('wb-over60.example.jp') -10 61.5
    Add-TestCert 'within' 'CN=wb-within' @('wb-within.example.jp') -10 0.5
    Add-TestCert 'justExpired' 'CN=wb-just' @('wb-just.example.jp') -10 -0.5
    Add-TestCert 'long' 'CN=wb-long' @('wb-long.example.jp') -10 2000.5
    Add-TestCert 'oldLong' 'CN=wb-oldlong' @('wb-oldlong.example.jp') -3000 -1499.5
    # サブジェクトと別名の形
    Add-TestCert 'comma' 'CN=wb-comma, O="Example, Inc.", C=JP' @('wb-comma.example.jp') -10 100
    Add-TestCert 'japanese' 'CN=日本語の証明書' @('wb-jp.example.jp') -10 100
    Add-TestCert 'nosubject' '' @('wb-nosubject.example.jp') -300 15.5
    Add-TestCert 'ip' 'CN=wb-ip' @('wb-ip.example.jp') -10 100 @{ Ip = @('192.0.2.7', '2001:db8::7') }
    Add-TestCert 'ipOnly' 'CN=wb-iponly' @() -10 100 @{ Ip = @('192.0.2.8') }
    Add-TestCert 'plain' 'CN=wb-plain' @() -10 100
    Add-TestCert 'usage' 'CN=wb-usage' @('wb-usage.example.jp') -10 100 @{ Usage = @('1.3.6.1.5.5.7.3.1', '1.3.6.1.5.5.7.3.2', '1.3.6.1.5.5.7.3.3', '1.3.6.1.5.5.7.3.4', '1.3.6.1.5.5.7.3.8', '1.3.6.1.5.5.7.3.9', '1.3.6.1.4.1.311.10.3.4', '1.3.6.1.4.1.311.20.2.2', '1.2.3.4.5') }
    Add-TestCert 'bracket' 'CN=wb-[b]' @('wb-bracket.example.jp') -10 100
    Add-TestCert 'bracketPlain' 'CN=wb-b' @('wb-bracketplain.example.jp') -10 100
    # 更新済みの決まり
    Add-TestCert 'renewOld' 'CN=wb-renew' @('wb-renew.example.jp') -400 -5
    Add-TestCert 'renewNew' 'CN=wb-renew' @('wb-renew.example.jp') -5 360
    Add-TestCert 'sanOther' 'CN=wb-renew' @('wb-other.example.jp') -400 -3
    Add-TestCert 'nearOld' 'CN=wb-nearrenew' @('wb-nearrenew.example.jp') -300 10
    Add-TestCert 'nearNew' 'CN=wb-nearrenew' @('wb-nearrenew.example.jp') -5 360
    Add-TestCert 'keyOld' 'CN=wb-key1' @('wb-key1.example.jp') -400 -5 @{ Key = $true }
    Add-TestCert 'keyNew' 'CN=wb-key1' @('wb-key1.example.jp') -5 360
    Add-TestCert 'key2Old' 'CN=wb-key2' @('wb-key2.example.jp') -400 -5 @{ Key = $true }
    Add-TestCert 'key2New' 'CN=wb-key2' @('wb-key2.example.jp') -5 360 @{ Key = $true }
    Add-TestCert 'key3Old' 'CN=wb-key3' @('wb-key3.example.jp') -400 -5
    Add-TestCert 'key3New' 'CN=wb-key3' @('wb-key3.example.jp') -5 360 @{ Key = $true }
    Add-TestCert 'weakOld' 'CN=wb-weak' @('wb-weak.example.jp') -400 -5
    Add-TestCert 'weakNew' 'CN=wb-weak' @('wb-weak.example.jp') -5 30
    Add-TestCert 'futOld' 'CN=wb-fut' @('wb-fut.example.jp') -400 -5
    Add-TestCert 'futNew' 'CN=wb-fut' @('wb-fut.example.jp') 10 400
    Add-TestCert 'sanOld' 'CN=wb-sanorder' @('a.wb-sanorder.example.jp', 'b.wb-sanorder.example.jp') -400 -5
    Add-TestCert 'sanNew' 'CN=wb-sanorder' @() -5 360 @{ RawDns = @('B.wb-sanorder.example.jp', 'a.wb-sanorder.example.jp') }
    Add-TestCert 'subjA' 'CN=wb-subj-a' @('wb-subj.example.jp') -400 -5
    Add-TestCert 'subjB' 'CN=wb-subj-b' @('wb-subj.example.jp') -5 360
    Add-TestCert 'earlyOld' 'CN=wb-early' @('wb-early.example.jp') -5 360
    Add-TestCert 'earlyNew' 'CN=wb-early' @('wb-early.example.jp') -400 -5
    Add-TestCert 'iisOld' 'CN=wb-iis' @('wb-iis.example.jp') -400 -5
    Add-TestCert 'iisNew' 'CN=wb-iis' @('wb-iis.example.jp') -5 360
    Add-TestCert 'crossOld' 'CN=wb-cross' @('wb-cross.example.jp') -400 -5
    Add-TestCert 'crossNew' 'CN=wb-cross' @('wb-cross.example.jp') -5 360 @{ Store2 = $true }
    # LocalMachine と重なる証明書（both は両方に、userOnly は CurrentUser にだけある）
    Add-TestCert 'bothExpired' 'CN=wb-both-expired' @('wb-both-expired.example.jp') -400 -5 @{ Into = $userStore3; Also = $userStore3m }
    Add-TestCert 'userOnly' 'CN=wb-overlap' @('wb-overlap.example.jp') -400 -5 @{ Into = $userStore3 }
    Add-TestCert 'bothNew' 'CN=wb-overlap' @('wb-overlap.example.jp') -5 360 @{ Into = $userStore3; Also = $userStore3m }
    Add-TestCert 'machineOnly' 'CN=wb-machine' @('wb-machine.example.jp') -10 100 @{ Into = $userStore3m }
    Add-TestCert 'otherName' 'CN=wb-othername' @('wb-othername.example.jp') -10 100 @{ Store2 = $true; Also = $userStore3m }
    # ストアに入れず、ファイルと TLS サーバーで使う
    Add-TestCert 'local' 'CN=localhost' @('localhost') -10 100 @{ NoStore = $true }
    Add-TestCert 'fileOk' 'CN=wb-file' @('wb-file.example.jp') -10 300 @{ NoStore = $true }
    $userStore.Close()
    $userStore2.Close()
    $userStore3.Close()
    $userStore3m.Close()

    $pfxBytes = $pfxOf['valid']
    $allPfx = New-Object "$x509.X509Certificate2Collection"
    [void]$allPfx.Add((New-Object "$x509.X509Certificate2" (, [byte[]]$rawOf['valid'])))
    $p7Bytes = $allPfx.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pkcs7)

    #endregion

    Write-Section 'ストア: 状態と残り日数'
    $run = Invoke-Tool @('-Store', $storeText)
    Assert-Equal '終了コード（期限切れあり）' $run.Code 2
    Assert-Equal '保存したファイル' ((Get-ChildItem -LiteralPath $run.Dir | ForEach-Object { $_.Name } | Sort-Object) -join ',') 'certificates.csv,skipped.csv,summary.txt'
    Assert-Equal '保存先のフォルダの名前' ((Split-Path -Leaf $run.Dir) -replace '\d{8}_\d{6}$', 'DATE') ('CertificateExpiry_{0}_DATE' -f $env:COMPUTERNAME)
    foreach ($case in @(
        @('期限切れ', 'expired', '期限切れ'), @('期限間近', 'near', '期限間近'), @('有効', 'valid', '有効'), @('有効期間前', 'future', '有効期間前'),
        @('残りが基準ちょうどなら期限間近', 'at60', '期限間近'), @('残りが基準より多いと有効', 'over60', '有効'),
        @('残り 1 日未満は期限間近', 'within', '期限間近'), @('過ぎたばかりは期限切れ', 'justExpired', '期限切れ')
    )) { Assert-Equal $case[0] (Get-Status $run $case[1]) $case[2] }
    foreach ($case in @(
        @('残り日数', 'near', '20'), @('残り日数（基準ちょうど）', 'at60', '60'), @('残り日数（基準 + 1）', 'over60', '61'), @('残り 24 時間未満は 0', 'within', '0'),
        @('過ぎると負の数', 'justExpired', '-1'), @('期限切れの日数', 'expired', '-10'), @('3 桁区切り', 'long', '2,000'), @('負の数も 3 桁区切り', 'oldLong', '-1,500')
    )) { Assert-Equal $case[0] (Get-Cell $run $case[1] '残り日数') $case[2] }

    Write-Section 'ストア: 更新済みの決まり'
    foreach ($case in @(
        @('同じサブジェクトと SAN の新しい証明書があれば更新済み', 'renewOld', '更新済み'), @('新しい証明書は有効', 'renewNew', '有効'),
        @('SAN が違うと更新済みにしない', 'sanOther', '期限切れ'), @('期限間近でも新しい証明書があれば更新済み', 'nearOld', '更新済み'),
        @('古い方に秘密鍵があり、新しい方になければ更新済みにしない', 'keyOld', '期限切れ'), @('どちらにも秘密鍵があれば更新済み', 'key2Old', '更新済み'),
        @('古い方に秘密鍵がなければ、新しい方はどちらでもよい', 'key3Old', '更新済み'), @('新しい証明書が期限間近なら更新済みにしない', 'weakOld', '期限切れ'),
        @('その新しい証明書は期限間近のまま', 'weakNew', '期限間近'), @('新しい証明書が有効期間前なら更新済みにしない', 'futOld', '期限切れ'),
        @('SAN の順番と大文字小文字は問わない', 'sanOld', '更新済み'), @('サブジェクトが違うと更新済みにしない', 'subjA', '期限切れ'),
        @('有効期限が早い方は、遅い方を置き換えない', 'earlyOld', '有効'), @('別のストアの証明書では更新済みにしない', 'crossOld', '期限切れ')
    )) { Assert-Equal $case[0] (Get-Status $run $case[1]) $case[2] }
    Assert-Equal '更新済みの備考に新しい証明書の拇印' (Get-Cell $run 'renewOld' '備考') ('新しい証明書: {0}' -f $thumbs['renewNew'])
    Assert-Equal '更新しないものの備考は空' (Get-Cell $run 'expired' '備考') ''

    Write-Section 'ストア: CSV の形'
    $certFile = Join-Path $run.Dir 'certificates.csv'
    Assert-Equal 'certificates.csv の見出し' ([IO.File]::ReadAllLines($certFile)[0]) '"状態","残り日数","有効期限","有効開始日","種類","場所","サブジェクト","別名(SAN)","発行者","用途","秘密鍵","IIS のサイト","フレンドリ名","拇印","備考"'
    Assert-Equal 'skipped.csv の見出し' ([IO.File]::ReadAllLines((Join-Path $run.Dir 'skipped.csv'))[0]) '"理由","対象","詳細"'
    Assert-Equal 'BOM（3 つのファイル）' ('{0}/{1}/{2}' -f (Test-FileBom $certFile), (Test-FileBom (Join-Path $run.Dir 'skipped.csv')), (Test-FileBom (Join-Path $run.Dir 'summary.txt'))) 'True/True/True'
    $row = @(Get-CertRow $run 'near')[0]
    Assert-Equal '行の全部の列' (($row.PSObject.Properties | Where-Object { $_.Name -ne '有効開始日' } | ForEach-Object { $_.Value }) -join '|') ('期限間近|20|{0}|ストア|{1}|CN=wb-near|wb-near.example.jp|CN=wb-near|サーバー認証|あり|||{2}|' -f (Format-CsvTime $timeOf['near']), $storeText, $thumbs['near'])
    Assert-Equal '有効開始日' $row.'有効開始日' (Format-CsvTime $fromOf['near'])
    Assert-Equal 'IP アドレスだけの別名は Excel で変わらないように守る' (Get-Cell $run 'ipOnly' '別名(SAN)') '="192.0.2.8"'
    Assert-Equal '秘密鍵なし' (Get-Cell $run 'expired' '秘密鍵') 'なし'
    Assert-Equal '別名(SAN) は , でつなぐ' (Get-Cell $run 'valid' '別名(SAN)') 'wb-valid.example.jp, www.wb-valid.example.jp'
    Assert-Equal '別名(SAN) の IP アドレス' (Get-Cell $run 'ip' '別名(SAN)') 'wb-ip.example.jp, 192.0.2.7, 2001:db8::7'
    Assert-Equal '別名がない証明書' (Get-Cell $run 'plain' '別名(SAN)') ''
    Assert-Equal 'Excel で変わるフレンドリ名は ="..." で守る' (Get-Cell $run 'valid' 'フレンドリ名') '="-test""x"'
    Assert-Equal 'カンマと引用符を含むサブジェクト' (Get-Cell $run 'comma' 'サブジェクト') 'CN=wb-comma, O="Example, Inc.", C=JP'
    Assert-Equal '日本語のサブジェクト' (Get-Cell $run 'japanese' 'サブジェクト') 'CN=日本語の証明書'
    Assert-Equal '用途は OS の言語によらない' (Get-Cell $run 'usage' '用途') 'サーバー認証, クライアント認証, コード署名, 電子メールの保護, タイムスタンプ, OCSP 署名, 暗号化ファイルシステム, スマートカード ログオン, 1.2.3.4.5'
    Assert-Equal '用途がない証明書' (Get-Cell $run 'plain' '用途') ''
    $times = [string[]]@($run.Certs | ForEach-Object { $_.'有効期限' })
    $sortedTimes = [string[]]$times.Clone()
    [Array]::Sort($sortedTimes, [StringComparer]::Ordinal)
    Assert-Equal '有効期限の早い順' ($times -join ',') ($sortedTimes -join ',')
    Assert-Equal '有効期限の早い順（最初は最も古いもの）' (Get-ThumbText @($run.Certs)[0]) $thumbs['oldLong']
    # 数字だけの拇印は Excel で変わる（本物では作れないので書き換えたコピーで確かめる）
    $patched = New-PatchedTool 'digits' 'Thumbprint = $Cert.Thumbprint;' "Thumbprint = '1234567890123456789012345678901234567890';"
    New-TextFile "$workRoot\digits\a.crt" (ConvertTo-Pem $rawOf['valid'])
    $digits = Invoke-Tool @('-Path', "$workRoot\digits") $patched
    Assert-Equal '数字だけの拇印は ="..." で守る' (@($digits.Certs)[0].'拇印') '="1234567890123456789012345678901234567890"'

    Write-Section 'ストア: 画面の形'
    Assert-Equal '1 行目' $run.Summary[0] '証明書期限一覧 (Get-CertificateExpiry 1.0.0)'
    Assert-Equal '2 行目（日時・ホスト名・管理者）' ($run.Summary[1] -replace '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', 'DATE') ('DATE   {0}（管理者: {1}）' -f $env:COMPUTERNAME, $(if ($isAdmin) { 'はい' } else { 'いいえ' }))
    Assert-Equal '対象の行' $run.Summary[2] "対象: $storeText"
    Assert-Equal '基準の行' $run.Summary[3] '期限間近の基準: 残り 60 日以内'
    Assert-Equal '基準の次は空行' $run.Summary[4] ''
    Assert-Equal '所要の形' ((Get-SummaryLine $run '証明書 ') -match '  所要 \d+\.\d秒$') 'True'
    $tableStart = [array]::IndexOf([string[]]$run.Summary, '■ 対応が必要な証明書（期限の近い順）')
    Assert-Equal '表の前に空行' $run.Summary[$tableStart - 1] ''
    $placeWidth = Get-Width $storeText
    Assert-Equal '表の見出し' $run.Summary[$tableStart + 1] (' 状態      残り日数  有効期限          ' + (Get-Pad '場所' $placeWidth) + '  サブジェクト')
    Assert-Equal '表の行（期限の近い順の最初）' $run.Summary[$tableStart + 2] (Get-TableLine '期限切れ' '-1,500' $timeOf['oldLong'] $storeText $placeWidth 'CN=wb-oldlong')
    Assert-Equal '表の行（期限間近）' (@($run.Summary | Where-Object { $_ -like '* CN=wb-near' }) -join '|') (Get-TableLine '期限間近' '20' $timeOf['near'] $storeText $placeWidth 'CN=wb-near')
    Assert-Equal 'サブジェクトが空なら別名で見分ける' (@($run.Summary | Where-Object { $_ -like '*（サブジェクトなし）*' }) -join '|') (Get-TableLine '期限間近' '15' $timeOf['nosubject'] $storeText $placeWidth '（サブジェクトなし）wb-nosubject.example.jp')
    Assert-Equal '表に更新済みを出さない' (@($run.Summary | Where-Object { $_ -like '* CN=wb-renew' -and $_ -like ' 更新済み*' }).Count) 0
    Assert-Equal '表に有効を出さない' (@($run.Summary | Where-Object { $_ -like '* CN=wb-valid' }).Count) 0
    Assert-Equal '表に有効期間前を出さない' (@($run.Summary | Where-Object { $_ -like '* CN=wb-future' }).Count) 0
    $tableTimes = [string[]]@($run.Summary | Where-Object { $_ -like ' 期限*' } | ForEach-Object { $_.Substring(17, 16) })
    $sortedTable = [string[]]$tableTimes.Clone()
    [Array]::Sort($sortedTable, [StringComparer]::Ordinal)
    Assert-Equal '表の行は期限の近い順' ($tableTimes -join ',') ($sortedTable -join ',')
    Assert-Equal '表の行数（期限切れと期限間近だけ）' $tableTimes.Count @($run.Certs | Where-Object { $_.'状態' -eq '期限切れ' -or $_.'状態' -eq '期限間近' }).Count
    Assert-Equal '保存先の行' (Get-SummaryLine $run '保存先: ') "保存先: $($run.Dir)"
    # 画面にも同じ内容を出す（折り返しの影響がないよう、空白を除いて比べる）
    $flatLog = $run.Log -replace '\s', ''
    $missing = @($run.Summary | Where-Object { $_ -ne '' -and -not (Test-Contains $flatLog ($_ -replace '\s', '')) })
    Assert-Equal '画面に summary.txt のすべての行' ($missing -join '|') ''
    Assert-Equal 'ストアの読み取りで \\?\ を出さない' (Test-Contains $run.Log '\\?\') 'False'

    Write-Section '-Days'
    $r = Invoke-Tool @('-Store', $storeText, '-Days', '400')
    Assert-Equal '-Days で基準が変わる' (Get-Status $r 'valid') '期限間近'
    Assert-Equal '新しい証明書も期限間近なら更新済みにしない' (Get-Status $r 'renewOld') '期限切れ'
    Assert-Equal '基準の行' $r.Summary[3] '期限間近の基準: 残り 400 日以内'
    $r = Invoke-NoSave @('-Store', $storeText, '-Days', '3650')
    Assert-Equal '基準の日数は指定のまま' (Test-Contains $r.Log '期限間近の基準: 残り 3650 日以内') 'True'
    $r = Invoke-Tool @('-Store', $storeText, '-Days', '20')
    Assert-Equal '残り 20 日は -Days 20 で期限間近' (Get-Status $r 'near') '期限間近'
    $r = Invoke-Tool @('-Store', $storeText, '-Days', '19')
    Assert-Equal '残り 20 日は -Days 19 で有効' (Get-Status $r 'near') '有効'
    foreach ($case in @(@('-Days', '0'), @('-Days', '3651'), @('-Days', 'abc'), @('-Timeout', '0'), @('-Timeout', '61'))) {
        $out = "$workRoot\out_bad_$($case[0])_$($case[1])"
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Store $storeText $case[0] $case[1] -OutputFolder $out 2>&1
        Assert-Equal ('{0} {1} は調べる前に止まる' -f $case[0], $case[1]) ('{0}/{1}' -f $LASTEXITCODE, [bool](Get-RunDir $out)) '1/False'
    }

    Write-Section '-Exclude'
    $r = Invoke-Tool @('-Store', $storeText, '-Exclude', 'CN=wb-near,*wb-expired*')
    Assert-Equal '-Exclude で除く' ('{0}/{1}' -f @(Get-CertRow $r 'near').Count, @(Get-CertRow $r 'expired').Count) '0/0'
    Assert-Equal '除外の行' (Get-SummaryLine $r '除外: ') '除外: CN=wb-near, *wb-expired*'
    Assert-Equal '件数の行に除外' ((Get-SummaryLine $r '証明書 ') -match '  除外 2 件  所要') 'True'
    Assert-Equal '除外の行は基準の次' $r.Summary[4] '除外: CN=wb-near, *wb-expired*'
    Assert-Equal 'skipped.csv に除外指定' ('{0}/{1}' -f (Get-SkipReason $r "${storeText}: CN=wb-near"), (Get-SkipDetail $r "${storeText}: CN=wb-near")) '除外指定/CN=wb-near'
    $r = Invoke-Tool @('-Store', $storeText, '-Exclude', 'cn=WB-NEAR')
    Assert-Equal '大文字と小文字は区別しない' @(Get-CertRow $r 'near').Count 0
    $r = Invoke-Tool @('-Store', $storeText, '-Exclude', 'wb-near')
    Assert-Equal '* がなければ全体が一致するものだけ' @(Get-CertRow $r 'near').Count 1
    $r = Invoke-Tool @('-Store', $storeText, '-Exclude', '*Example, Inc.*,CN=wb-near')
    Assert-Equal 'カンマとスペースを含むサブジェクトも除ける' ('{0}/{1}' -f @(Get-CertRow $r 'comma').Count, @(Get-CertRow $r 'near').Count) '0/0'
    $r = Invoke-Tool @('-Store', $storeText, '-Exclude', 'CN=wb-[b]')
    Assert-Equal '[ ] はそのままの文字' ('{0}/{1}' -f @(Get-CertRow $r 'bracket').Count, @(Get-CertRow $r 'bracketPlain').Count) '0/1'
    $r = Invoke-Tool @('-Store', $storeText, '-Exclude', 'CN=wb-?')
    Assert-Equal '? はそのままの文字' @(Get-CertRow $r 'bracketPlain').Count 1
    $r = Invoke-Tool @('-Store', $storeText, '-Exclude', 'CN=wb-`b')
    Assert-Equal '` はそのままの文字' @(Get-CertRow $r 'bracketPlain').Count 1
    $r = Invoke-Tool @('-Store', $storeText, '-Exclude', '*日本語*')
    Assert-Equal '日本語の -Exclude' @(Get-CertRow $r 'japanese').Count 0

    Write-Section 'ストアの指定'
    # 本物の既定のストアは読むだけ。書き込みは専用のストアにしかしない
    $r = Invoke-NoSave @()
    Assert-Equal '既定のストア' (Test-Contains $r.Log '対象: LocalMachine\My, LocalMachine\WebHosting, LocalMachine\Remote Desktop, CurrentUser\My') 'True'
    $realMy = New-Object "$x509.X509Store" ('My', [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $realMy.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    $realCount = @($realMy.Certificates).Count
    $realMy.Close()
    $r = Invoke-Tool @('-Store', 'CurrentUser\My')
    Assert-Equal '本物の CurrentUser\My を読む' @(Get-PlaceRows $r 'CurrentUser\My').Count $realCount
    $patched = New-PatchedTool 'defaults' "'LocalMachine\My', 'LocalMachine\WebHosting', 'LocalMachine\Remote Desktop', 'CurrentUser\My'" "'CurrentUser\wbnone-$runId', '$storeText'"
    $r = Invoke-Tool @() $patched
    Assert-Equal '既定のストアの実行でも証明書を見つける' (Get-Status $r 'expired') '期限切れ'
    Assert-Equal 'ないストアは何も言わずに飛ばす（既定のとき）' ('{0}/{1}' -f @($r.Skipped).Count, (Test-Contains $r.Log 'ストアが見つかりません')) '0/False'
    $r = Invoke-Tool @('-Store', "$store2Text,$storeText")
    Assert-Equal '複数のストア' ('{0}/{1}' -f @(Get-PlaceRows $r $store2Text).Count, (@(Get-PlaceRows $r $storeText).Count -gt 0)) '2/True'
    Assert-Equal '対象の行に複数のストア' $r.Summary[2] "対象: $store2Text, $storeText"
    $r = Invoke-Tool @('-Store', "$storeText,$storeText")
    Assert-Equal '同じストアを 2 回指定しても 1 回だけ' @(Get-CertRow $r 'near').Count 1
    $r = Invoke-Tool @('-Store', "currentuser\$($storeName.ToUpperInvariant())")
    Assert-Equal 'ストアの指定は大文字小文字を問わない' ('{0}/{1}' -f (Get-Status $r 'near'), $r.Summary[2]) ('期限間近/対象: CurrentUser\{0}' -f $storeName.ToUpperInvariant())
    $r = Invoke-Tool @('-Store', $store2Text)
    Assert-Equal '有効だけのストアは 0' $r.Code 0
    $r = Invoke-Tool @('-AllStores')
    Assert-Equal '-AllStores の対象の行' $r.Summary[2] '対象: すべてのストア'
    Assert-Equal '-AllStores は Root も調べる' (@($r.Certs | Where-Object { $_.'場所' -eq 'LocalMachine\Root' }).Count -gt 0) 'True'
    Assert-Equal '-AllStores はテスト専用のストアも調べる' (@(Get-PlaceRows $r $storeText).Count -gt 0) 'True'
    Assert-Equal '-AllStores は UserDS を調べない' (@($r.Certs | Where-Object { $_.'場所' -like '*\UserDS' }).Count) 0
    if (@(Get-ChildItem -LiteralPath 'Cert:\CurrentUser' | Where-Object { $_.Name -eq 'UserDS' }).Count -gt 0) {
        Assert-Equal 'UserDS は除外指定に残す' ('{0}/{1}' -f (Get-SkipReason $r 'CurrentUser\UserDS'), (Get-SkipDetail $r 'CurrentUser\UserDS')) '除外指定/Active Directory に問い合わせるストアのため'
    } else {
        Write-Skip 'UserDS' 'このPCにないため'
    }

    Write-Section 'LocalMachine と重なる証明書'
    # 本物の Root は、Windows が CurrentUser にも見せる（読むだけ）
    $userRoot = Invoke-Tool @('-Store', 'CurrentUser\Root')
    $machineRoot = Invoke-Tool @('-Store', 'LocalMachine\Root')
    $bothRoot = Invoke-Tool @('-Store', 'LocalMachine\Root,CurrentUser\Root')
    $overlapRows = @($bothRoot.Skipped | Where-Object { $_.'理由' -eq '除外指定' -and $_.'対象' -like 'CurrentUser\Root: *' })
    Assert-Equal 'CurrentUser\Root だけなら、見えているものをすべて調べる' ('{0}/{1}' -f (@($userRoot.Certs).Count -gt 0), (Test-Contains $userRoot.Log 'LocalMachine と重なる')) 'True/False'
    Assert-Equal '本物の Root で重なりを省く' (@($overlapRows).Count -gt 0) 'True'
    Assert-Equal '省いた分だけ行が減る' @($bothRoot.Certs).Count (@($machineRoot.Certs).Count + @($userRoot.Certs).Count - $overlapRows.Count)
    Assert-Equal '省いた件数の行' (@($bothRoot.Summary) -contains ('LocalMachine と重なる証明書 {0} 件は省きました' -f $overlapRows.Count)) 'True'
    Assert-Equal '省いた理由' @($overlapRows)[0].'詳細' '同じ証明書が LocalMachine\Root にあります'
    $machineThumbs = @{}
    foreach ($row in (Get-PlaceRows $bothRoot 'LocalMachine\Root')) { $machineThumbs[(Get-ThumbText $row)] = $true }
    Assert-Equal 'CurrentUser\Root に LocalMachine と同じ行が残らない' (@(Get-PlaceRows $bothRoot 'CurrentUser\Root' | Where-Object { $machineThumbs.ContainsKey((Get-ThumbText $_)) }).Count) 0
    Assert-Equal 'LocalMachine の行はすべて残す' @(Get-PlaceRows $bothRoot 'LocalMachine\Root').Count @($machineRoot.Certs).Count
    $reversed = Invoke-Tool @('-Store', 'CurrentUser\Root,LocalMachine\Root')
    Assert-Equal '指定の順番によらず同じ' @($reversed.Certs).Count @($bothRoot.Certs).Count
    $otherName = Invoke-Tool @('-Store', 'LocalMachine\CA,CurrentUser\Root')
    Assert-Equal '名前の違うストアとは重ねない' @(Get-PlaceRows $otherName 'CurrentUser\Root').Count @($userRoot.Certs).Count
    $all = Invoke-Tool @('-AllStores')
    Assert-Equal '-AllStores でも省く' (@(Get-PlaceRows $all 'CurrentUser\Root' | Where-Object { $machineThumbs.ContainsKey((Get-ThumbText $_)) }).Count) 0
    Assert-Equal '-AllStores の専用のストアは省かない' (@(Get-PlaceRows $all $storeText).Count -gt 0) 'True'
    # 中身を決めたいときは、LocalMachine\<名前> で <名前>-m を開くコピーを使う
    $machineMap = 'New-Object Security.Cryptography.X509Certificates.X509Store($(if ($Item.Place -eq ''LocalMachine'') { $Item.Name + ''-m'' } else { $Item.Name }), [Security.Cryptography.X509Certificates.StoreLocation]''CurrentUser'')'
    $patched = New-PatchedTool 'machinemap' 'New-Object Security.Cryptography.X509Certificates.X509Store($Item.Name, [Security.Cryptography.X509Certificates.StoreLocation]$Item.Place)' $machineMap
    $r = Invoke-Tool @('-Store', "LocalMachine\$store3Name,CurrentUser\$store3Name") $patched
    $userPlace = "CurrentUser\$store3Name"
    $machinePlace = "LocalMachine\$store3Name"
    Assert-Equal '重なる証明書は LocalMachine の行だけ' ('{0}/{1}' -f @(Get-CertRow $r 'bothExpired').Count, (Get-Cell $r 'bothExpired' '場所')) "1/$machinePlace"
    Assert-Equal 'CurrentUser にだけある証明書は残す' (Get-Cell $r 'userOnly' '場所') $userPlace
    Assert-Equal 'LocalMachine にだけある証明書は残す' (Get-Cell $r 'machineOnly' '場所') $machinePlace
    Assert-Equal '更新済みの判定は、省く前の中身で行う' (Get-Status $r 'userOnly') '更新済み'
    Assert-Equal '省かれた新しい証明書は LocalMachine の行に残る' ('{0}/{1}' -f (Get-Status $r 'bothNew'), (Get-Cell $r 'bothNew' '場所')) "有効/$machinePlace"
    Assert-Equal '省いた分は件数に数えない' (Get-CountLine $r) '証明書 5 件  期限切れ 1 件  期限間近 0 件  更新済み 1 件  有効 3 件'
    Assert-Equal '省いた件数' (@($r.Summary) -contains 'LocalMachine と重なる証明書 2 件は省きました') 'True'
    Assert-Equal '省いた行の対象' (Get-SkipReason $r "${userPlace}: CN=wb-both-expired") '除外指定'
    Assert-Equal '省いただけなら読めないものに数えない（期限切れがあるので 2）' $r.Code 2
    Assert-Equal '省いた件数の行は件数の行の次' ([array]::IndexOf([string[]]$r.Summary, 'LocalMachine と重なる証明書 2 件は省きました')) ([array]::IndexOf([string[]]$r.Summary, (Get-SummaryLine $r '証明書 ')) + 1)
    $r = Invoke-Tool @('-Store', "LocalMachine\$store3Name,CurrentUser\$($store3Name.ToUpperInvariant())") $patched
    Assert-Equal 'ストアの名前の大文字小文字は問わない' (Test-Contains $r.Log 'LocalMachine と重なる証明書 2 件は省きました') 'True'
    $r = Invoke-Tool @('-Store', "LocalMachine\$store3Name,$store2Text") $patched
    Assert-Equal '名前の違うストアの同じ証明書は省かない' ('{0}/{1}' -f @(Get-CertRow $r 'otherName').Count, (Test-Contains $r.Log 'LocalMachine と重なる')) '2/False'

    # ストアを開けないとき（本物では作れないので書き換えたコピーで確かめる）
    $openLine ='$certStore.Open([Security.Cryptography.X509Certificates.OpenFlags]''ReadOnly, OpenExistingOnly'')'
    $denyText = 'if ($Item.Name -eq ''{0}'') {{ throw (New-Object UnauthorizedAccessException ''拒否' + "`r`n" + 'しました'') }}; {1}'
    $patched = New-PatchedTool 'storedeny' $openLine ($denyText -f $store2Name, $openLine)
    $r = Invoke-Tool @('-Store', "$store2Text,$storeText") $patched
    Assert-Equal 'ストアの権限なしは記録して続ける' ('{0}/{1}/{2}' -f (Get-SkipReason $r $store2Text), (Get-SkipDetail $r $store2Text), (@(Get-PlaceRows $r $storeText).Count -gt 0)) '権限なし/拒否 しました/True'
    Assert-Equal 'スキップの行' (Get-SummaryLine $r 'スキップ ') 'スキップ 1 件（権限なし）'
    Assert-Equal '権限なしのときの案内' ((@($r.Summary) -contains '※ 管理者として実行すると、読める範囲が広がります')) $(if ($isAdmin) { 'False' } else { 'True' })
    $r = Invoke-Tool @('-Store', $store2Text) $patched
    Assert-Equal '権限なしだけなら 1' $r.Code 1
    $patched = New-PatchedTool 'storehr' $openLine ('if ($Item.Name -eq ''{0}'') {{ throw (New-Object Security.Cryptography.CryptographicException (-2147024891)) }}; {1}' -f $store2Name, $openLine)
    Assert-Equal 'HRESULT のアクセス拒否も権限なし' (Get-SkipReason (Invoke-Tool @('-Store', $store2Text) $patched) $store2Text) '権限なし'
    $patched = New-PatchedTool 'storesec' $openLine ('if ($Item.Name -eq ''{0}'') {{ throw (New-Object Security.SecurityException ''no'') }}; {1}' -f $store2Name, $openLine)
    Assert-Equal 'SecurityException も権限なし' (Get-SkipReason (Invoke-Tool @('-Store', $store2Text) $patched) $store2Text) '権限なし'
    $patched = New-PatchedTool 'storeother' $openLine ('if ($Item.Name -eq ''{0}'') {{ throw (New-Object InvalidOperationException ''oops'') }}; {1}' -f $store2Name, $openLine)
    $r = Invoke-Tool @('-Store', $store2Text) $patched
    Assert-Equal 'ほかのエラーはその他のエラー' ('{0}/{1}/{2}' -f (Get-SkipReason $r $store2Text), (Get-SummaryLine $r 'スキップ '), $r.Code) 'その他のエラー/スキップ 1 件（その他のエラー）/1'
    Assert-Equal 'その他のエラーでは管理者の案内を出さない' (@($r.Summary) -contains '※ 管理者として実行すると、読める範囲が広がります') 'False'
    # 管理者の表示は、管理者かどうかによらず両方を確かめる
    foreach ($case in @(@('$true', 'はい', 'False'), @('$false', 'いいえ', 'True'))) {
        $patched = New-PatchedTool "admin$($case[1])" '$outText = $null' ("`$isAdmin = {0}; `$outText = `$null; `$certStoreDeny = '{1}'" -f $case[0], $store2Name)
        $text = [IO.File]::ReadAllText($patched).Replace($openLine, ('if ($Item.Name -eq $certStoreDeny) {{ throw (New-Object UnauthorizedAccessException ''x'') }}; {0}' -f $openLine))
        [IO.File]::WriteAllText($patched, $text, $utf8Bom)
        $r = Invoke-Tool @('-Store', $store2Text) $patched
        Assert-Equal ('管理者の表示: {0}' -f $case[1]) ('{0}/{1}' -f (Test-Contains $r.Summary[1] "（管理者: $($case[1])）"), (@($r.Summary) -contains '※ 管理者として実行すると、読める範囲が広がります')) ('True/{0}' -f $case[2])
    }

    Write-Section 'ファイル'
    $files = "$workRoot\files"
    New-TestFile "$files\der.cer" $rawOf['near']
    New-TextFile "$files\pem.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$files\bundle.pem" ((ConvertTo-Pem $rawOf['near']) + (ConvertTo-Pem $rawOf['valid']))
    New-TextFile "$files\old.cer" (ConvertTo-Pem $rawOf['renewOld'])
    New-TextFile "$files\new.cer" (ConvertTo-Pem $rawOf['renewNew'])
    New-TestFile "$files\upper.CER" $rawOf['at60']
    New-TestFile "$files\mixed.Der" $rawOf['over60']
    New-TextFile "$files\trusted.pem" (ConvertTo-Pem $rawOf['future'] 'TRUSTED CERTIFICATE')
    New-TextFile "$files\x509.pem" (ConvertTo-Pem $rawOf['long'] 'X509 CERTIFICATE')
    New-TextFile "$files\attrs.pem" ("Bag Attributes`r`n    friendlyName: test`r`nsubject=CN=wb`r`n" + (ConvertTo-Pem $rawOf['oldLong']))
    New-TextFile "$files\withkey.pem" ((ConvertTo-Pem $rawOf['plain']) + "-----BEGIN PRIVATE KEY-----`r`nAAAA`r`n-----END PRIVATE KEY-----`r`n")
    New-TestFile "$files\日本語 フォルダ\[x]\expired.der" $rawOf['expired']
    $longDir = "$files\" + (('y' * 60 + '\') * 5)
    New-TestFile ($longDir + 'long.cer') $rawOf['future']
    New-TextFile "$files\truncated.pem" ((ConvertTo-Pem $rawOf['comma']) + "-----BEGIN CERTIFICATE-----`r`nMIIB")
    New-TextFile "$files\cut.pem" "-----BEGIN CERTIFICATE-----`r`nMIIB"
    New-TextFile "$files\badbase64.pem" "-----BEGIN CERTIFICATE-----`r`n!!!notbase64!!!`r`n-----END CERTIFICATE-----`r`n"
    New-TextFile "$files\notcert.pem" "-----BEGIN CERTIFICATE-----`r`nAAAA`r`n-----END CERTIFICATE-----`r`n"
    New-TextFile "$files\privkey.pem" "-----BEGIN PRIVATE KEY-----`r`nAAAA`r`n-----END PRIVATE KEY-----`r`n"
    New-TextFile "$files\rsakey.pem" "-----BEGIN RSA PRIVATE KEY-----`r`nAAAA`r`n-----END RSA PRIVATE KEY-----`r`n"
    New-TextFile "$files\dhparam.pem" "-----BEGIN DH PARAMETERS-----`r`nAAAA`r`n-----END DH PARAMETERS-----`r`n"
    New-TextFile "$files\csr.pem" "-----BEGIN CERTIFICATE REQUEST-----`r`nAAAA`r`n-----END CERTIFICATE REQUEST-----`r`n"
    New-TestFile "$files\secret.pfx" $pfxBytes
    New-TestFile "$files\secret.p12" $pfxBytes
    New-TestFile "$files\nopass.crt" $pfxBytes
    New-TestFile "$files\p7.crt" $p7Bytes
    New-TextFile "$files\broken.crt" 'not a certificate'
    New-TestFile "$files\zero.cer" ([byte[]]@())
    New-TextFile "$files\memo.txt" 'not a certificate'
    New-TestFile "$files\noext" $rawOf['valid']
    $stream = [IO.File]::Create("$files\huge.cer")
    $stream.SetLength(11MB)
    $stream.Close()
    $exact = [IO.File]::Create("$files\limit.cer")
    $exact.SetLength(10MB)
    $exact.Close()
    $null = New-Item -ItemType Junction -Path "$files\jloop" -Target $files
    $junctions.Add("$files\jloop")

    $fileRun = Invoke-Tool @('-Path', $files)
    Assert-Equal '-Path だけならストアは調べない' (@($fileRun.Certs | Where-Object { $_.'種類' -ne 'ファイル' }).Count) 0
    Assert-Equal '対象の行' $fileRun.Summary[2] "対象: $files"
    foreach ($case in @(
        @('DER', 'der.cer', '期限間近'), @('PEM', 'pem.crt', '有効'), @('同じフォルダの新しい証明書で更新済み', 'old.cer', '更新済み'),
        @('大文字の拡張子', 'upper.CER', '期限間近'), @('大文字小文字の混ざった拡張子', 'mixed.Der', '有効'),
        @('BEGIN TRUSTED CERTIFICATE', 'trusted.pem', '有効期間前'), @('BEGIN X509 CERTIFICATE', 'x509.pem', '有効'), @('証明書の前に文字がある PEM', 'attrs.pem', '期限切れ'),
        @('証明書と秘密鍵が入った PEM は証明書だけ読む', 'withkey.pem', '有効'), @('日本語・スペース・[ ] を含むパス', '日本語 フォルダ\[x]\expired.der', '期限切れ')
    )) { Assert-Equal $case[0] (Get-PlaceStatus $fileRun "$files\$($case[1])") $case[2] }
    Assert-Equal '複数の証明書が入った PEM' (@(Get-PlaceRows $fileRun "$files\bundle.pem").Count) 2
    Assert-Equal '証明書と秘密鍵が入った PEM は記録しない' (Get-SkipReason $fileRun "$files\withkey.pem") ''
    Assert-Equal '260 文字を超えるパス' (@(Get-PlaceRows $fileRun ($longDir + 'long.cer')).Count) 1
    Assert-Equal '\\?\ が残っていない' ((Get-ChildItem -LiteralPath $fileRun.Dir | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join '').Contains('\\?\') 'False'
    $fileRow = @(Get-PlaceRows $fileRun "$files\der.cer")[0]
    Assert-Equal '種類はファイル、秘密鍵と IIS は空欄' ('{0}/{1}/{2}' -f $fileRow.'種類', $fileRow.'秘密鍵', $fileRow.'IIS のサイト') 'ファイル//'
    foreach ($case in @(
        @('途中で切れた PEM は「読めない」', 'truncated.pem', '読めない'), @('証明書が 1 枚もない切れた PEM', 'cut.pem', '読めない'), @('Base64 が壊れた PEM', 'badbase64.pem', '読めない'),
        @('中身が証明書でない PEM', 'notcert.pem', '読めない'), @('秘密鍵だけの PEM は読まない', 'privkey.pem', '読まない形式'), @('RSA の秘密鍵だけの PEM', 'rsakey.pem', '読まない形式'),
        @('証明書以外の PEM は読まない', 'dhparam.pem', '読まない形式'), @('証明書要求は読まない', 'csr.pem', '読まない形式'), @('.pfx は開かない', 'secret.pfx', '読まない形式'),
        @('.p12 は開かない', 'secret.p12', '読まない形式'), @('中身が PFX の .crt は開かない', 'nopass.crt', '読まない形式'), @('中身が PKCS#7 の .crt は開かない', 'p7.crt', '読まない形式'),
        @('証明書でないファイルは「読めない」', 'broken.crt', '読めない'), @('0 バイトのファイル', 'zero.cer', '読めない'), @('10 MB を超えるファイルは読まない', 'huge.cer', '読めない'),
        @('関係のない拡張子は記録しない', 'memo.txt', ''), @('拡張子がないファイルは記録しない', 'noext', ''), @('ジャンクションはたどらない', 'jloop', 'リンク')
    )) { Assert-Equal $case[0] (Get-SkipReason $fileRun "$files\$($case[1])") $case[2] }
    Assert-Equal '途中で切れた PEM は読めた分を調べる' (@(Get-PlaceRows $fileRun "$files\truncated.pem").Count) 1
    Assert-Equal '証明書が 1 枚もない切れた PEM は行を作らない' (@(Get-PlaceRows $fileRun "$files\cut.pem").Count) 0
    Assert-Equal '中身が PFX の .crt を行にしない' (@(Get-PlaceRows $fileRun "$files\nopass.crt").Count) 0
    Assert-Equal 'ちょうど 10 MB は読む（中身が証明書でないので読めない）' (Test-Contains (Get-SkipDetail $fileRun "$files\limit.cer") '大きすぎます') 'False'
    Assert-Equal '大きすぎるときの詳細は 3 桁区切り' (Get-SkipDetail $fileRun "$files\huge.cer") '11,534,336 バイトあり、証明書のファイルとしては大きすぎます'
    Assert-Equal '秘密鍵だけの PEM の詳細' (Get-SkipDetail $fileRun "$files\rsakey.pem") '秘密鍵のファイルのため、読んでいません'
    Assert-Equal '証明書要求の詳細' (Get-SkipDetail $fileRun "$files\csr.pem") '証明書以外の PEM ファイルのため、読んでいません'
    Assert-Equal '.pfx の詳細' (Get-SkipDetail $fileRun "$files\secret.pfx") 'パスワードが必要な形式のため、開いていません'
    Assert-Equal '中身が PKCS#7 の詳細' (Get-SkipDetail $fileRun "$files\p7.crt") '証明書 1 枚の形式ではないため、開いていません（Pkcs7）'
    Assert-Equal '途中で切れた PEM の詳細' (Get-SkipDetail $fileRun "$files\truncated.pem") '証明書の途中でファイルが切れています'
    Assert-Equal '0 バイトのファイルの詳細' (Get-SkipDetail $fileRun "$files\zero.cer") '空のファイルです'
    Assert-Equal 'skipped の詳細は 1 行' (@($fileRun.Skipped | Where-Object { $_.'詳細' -match '[\r\n]' }).Count) 0
    Assert-Equal 'スキップの行（読めない・リンク・読まない形式）' (Get-SummaryLine $fileRun 'スキップ ') 'スキップ 8 件（読めない）  読まない形式 8 件  たどらなかったリンク 1 件'

    New-TextFile "$workRoot\memo-cert.txt" (ConvertTo-Pem $rawOf['valid'])
    $direct = Invoke-Tool @('-Path', "$workRoot\memo-cert.txt")
    Assert-Equal '直接指定は拡張子によらず読む' ('{0}/{1}' -f $direct.Code, @($direct.Certs).Count) '0/1'
    $r = Invoke-Tool @('-Path', "$files\secret.pfx")
    Assert-Equal '直接指定の .pfx も開かない' ('{0}/{1}/{2}' -f $r.Code, (Get-SkipReason $r "$files\secret.pfx"), @($r.Certs).Count) '0/読まない形式/0'
    $r = Invoke-Tool @('-Path', "$files\memo.txt")
    Assert-Equal '直接指定で読めないものは「読めない」' ('{0}/{1}' -f $r.Code, (Get-SkipReason $r "$files\memo.txt")) '1/読めない'
    $r = Invoke-Tool @('-Path', ($longDir + 'long.cer'))
    Assert-Equal '直接指定の 260 文字を超えるパス' (@($r.Certs).Count) 1
    $r = Invoke-Tool @('-Path', "$files\日本語 フォルダ\[x]\expired.der")
    Assert-Equal '直接指定の [ ] を含むパス' (@($r.Certs).Count) 1
    $overlap = Invoke-Tool @('-Path', "$files,$files\日本語 フォルダ")
    Assert-Equal '指定が重なっても 1 回だけ調べる' (@(Get-PlaceRows $overlap "$files\日本語 フォルダ\[x]\expired.der").Count) 1
    $r = Invoke-Tool @('-Path', "$files\der.cer,$files\DER.CER")
    Assert-Equal '大文字小文字だけ違う同じファイルも 1 回だけ' (@($r.Certs).Count) 1
    New-TextFile "$workRoot\only\a,b\x.crt" (ConvertTo-Pem $rawOf['valid'])
    Assert-Equal 'カンマを含むフォルダ' (@((Invoke-Tool @('-Path', "$workRoot\only\a,b")).Certs).Count) 1
    New-TestFile "$workRoot\scope1\old.cer" $rawOf['renewOld']
    New-TestFile "$workRoot\scope2\new.cer" $rawOf['renewNew']
    $r = Invoke-Tool @('-Path', "$workRoot\scope1,$workRoot\scope2")
    Assert-Equal '別のフォルダの証明書では更新済みにしない' (Get-PlaceStatus $r "$workRoot\scope1\old.cer") '期限切れ'
    # 表の列の幅は、全角を 2 文字分として数える
    New-TestFile "$workRoot\幅\長い名前.cer" $rawOf['near']
    New-TestFile "$workRoot\幅\ab.cer" $rawOf['within']
    $r = Invoke-Tool @('-Path', "$workRoot\幅")
    $placeA = "$workRoot\幅\ab.cer"
    $placeB = "$workRoot\幅\長い名前.cer"
    $placeWidth = Get-Width $placeB
    $tableStart = [array]::IndexOf([string[]]$r.Summary, '■ 対応が必要な証明書（期限の近い順）')
    Assert-Equal '全角を含む場所の見出し' $r.Summary[$tableStart + 1] (' 状態      残り日数  有効期限          ' + (Get-Pad '場所' $placeWidth) + '  サブジェクト')
    Assert-Equal '全角を含む場所の行' ('{0}|{1}' -f $r.Summary[$tableStart + 2], $r.Summary[$tableStart + 3]) ('{0}|{1}' -f (Get-TableLine '期限間近' '0' $timeOf['within'] $placeA $placeWidth 'CN=wb-within'), (Get-TableLine '期限間近' '20' $timeOf['near'] $placeB $placeWidth 'CN=wb-near'))

    Write-Section '件数の行と、表の上限'
    $countDir = "$workRoot\counts"
    foreach ($name in 'expired', 'near', 'valid', 'future', 'renewOld', 'renewNew') { New-TestFile "$countDir\$name.cer" $rawOf[$name] }
    $r = Invoke-Tool @('-Path', $countDir)
    Assert-Equal '件数の行' (Get-CountLine $r) '証明書 6 件  期限切れ 1 件  期限間近 1 件  更新済み 1 件  有効 2 件  有効期間前 1 件'
    $r = Invoke-Tool @('-Path', $countDir, '-Exclude', 'CN=wb-future,CN=wb-valid')
    Assert-Equal '有効期間前がなければ出さない・除外の件数' (Get-CountLine $r) '証明書 4 件  期限切れ 1 件  期限間近 1 件  更新済み 1 件  有効 1 件  除外 2 件'
    $manyDir = "$workRoot\many"
    New-TextFile "$manyDir\many.pem" ((ConvertTo-Pem $rawOf['near']) * 1234)
    $r = Invoke-Tool @('-Path', $manyDir)
    Assert-Equal '件数の 3 桁区切り' (Get-CountLine $r) '証明書 1,234 件  期限切れ 0 件  期限間近 1,234 件  更新済み 0 件  有効 0 件'
    Assert-Equal '画面の表は 200 件まで' ([regex]::Matches($r.Log, [regex]::Escape(' CN=wb-near')).Count) 200
    Assert-Equal '超えた分の案内' (Test-Contains $r.Log '（ほか 1,034 件は certificates.csv を参照してください）') 'True'
    Assert-Equal 'summary.txt はすべて書く' (@($r.Summary | Where-Object { $_ -like '* CN=wb-near' }).Count) 1234
    Assert-Equal 'summary.txt に超えた分の案内を書かない' (@($r.Summary | Where-Object { $_ -like '（ほか*' }).Count) 0
    $r = Invoke-NoSave @('-Path', $manyDir)
    Assert-Equal '-NoSave のときはすべて表示する' ('{0}/{1}' -f [regex]::Matches($r.Log, [regex]::Escape(' CN=wb-near')).Count, (Test-Contains $r.Log '（ほか')) '1234/False'

    Write-Section '読めない場所と、終了コード'
    New-TextFile "$workRoot\mix\ok.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\mix\bad.crt" 'broken'
    Assert-Equal '対応不要で読めないものがあると 1' (Invoke-Tool @('-Path', "$workRoot\mix")).Code 1
    New-TextFile "$workRoot\mix2\near.crt" (ConvertTo-Pem $rawOf['near'])
    New-TextFile "$workRoot\mix2\bad.crt" 'broken'
    Assert-Equal '対応が必要で読めないものもあると 2' (Invoke-Tool @('-Path', "$workRoot\mix2")).Code 2
    New-TextFile "$workRoot\clean\ok.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\clean\key.pem" "-----BEGIN PRIVATE KEY-----`r`nAAAA`r`n-----END PRIVATE KEY-----`r`n"
    $clean = Invoke-Tool @('-Path', "$workRoot\clean")
    Assert-Equal '「読まない形式」だけなら 0' ('{0}/{1}' -f $clean.Code, (Test-Contains $clean.Log '（対応が必要な証明書はありません）')) '0/True'
    Assert-Equal '対応が必要な証明書がないときの行' (@($clean.Summary) -contains '    （対応が必要な証明書はありません）') 'True'
    Assert-Equal 'ないときは表の見出しを出さない' (Test-Contains $clean.Log 'サブジェクト') 'False'
    [void][IO.Directory]::CreateDirectory("$workRoot\empty")
    $r = Invoke-Tool @('-Path', "$workRoot\empty")
    Assert-Equal '証明書がないフォルダは 0' ('{0}/{1}' -f $r.Code, (Get-CountLine $r)) '0/証明書 0 件  期限切れ 0 件  期限間近 0 件  更新済み 0 件  有効 0 件'
    New-TestFile "$workRoot\onlyrenew\old.cer" $rawOf['renewOld']
    New-TestFile "$workRoot\onlyrenew\new.cer" $rawOf['renewNew']
    Assert-Equal '更新済みだけなら 0' (Invoke-Tool @('-Path', "$workRoot\onlyrenew")).Code 0
    New-TestFile "$workRoot\onlyfuture\f.cer" $rawOf['future']
    Assert-Equal '有効期間前だけなら 0' (Invoke-Tool @('-Path', "$workRoot\onlyfuture")).Code 0
    New-TestFile "$workRoot\onlyex\near.cer" $rawOf['near']
    Assert-Equal '期限間近を除外したら 0' (Invoke-Tool @('-Path', "$workRoot\onlyex", '-Exclude', 'CN=wb-near')).Code 0
    $null = New-Item -ItemType Junction -Path "$workRoot\onlylink\j" -Target "$workRoot\clean" -Force
    $junctions.Add("$workRoot\onlylink\j")
    Assert-Equal 'リンクだけなら 0' (Invoke-Tool @('-Path', "$workRoot\onlylink")).Code 0

    New-TextFile "$workRoot\deny\ok.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\deny\secret\x.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\deny\noread.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\deny\bad.crt" 'broken'
    Deny-Access "$workRoot\deny\secret" '(RD)'
    Deny-Access "$workRoot\deny\noread.crt" '(R)'
    $canList = $true
    try { $null = [IO.Directory]::GetFileSystemEntries("$workRoot\deny\secret") } catch { $canList = $false }
    if ($canList) {
        Write-Skip 'アクセス拒否' '設定できなかったため'
    } else {
        $r = Invoke-Tool @('-Path', "$workRoot\deny")
        Assert-Equal '読めないフォルダ' (Get-SkipReason $r "$workRoot\deny\secret") '権限なし'
        Assert-Equal '読めないファイル' (Get-SkipReason $r "$workRoot\deny\noread.crt") '権限なし'
        Assert-Equal '読めない場所があっても続ける' ('{0}/{1}' -f $r.Code, @($r.Certs).Count) '1/1'
        Assert-Equal 'スキップの行（種類が 2 つ）' (Get-SummaryLine $r 'スキップ ') 'スキップ 3 件（権限なし 2・読めない 1）'
        Assert-Equal '権限なしのときの案内' (@($r.Summary) -contains '※ 管理者として実行すると、読める範囲が広がります') $(if ($isAdmin) { 'False' } else { 'True' })
        $r = Invoke-Tool @('-Path', "$workRoot\deny\secret")
        Assert-Equal '指定したフォルダ自体が読めない' ('{0}/{1}' -f $r.Code, (Get-SkipReason $r "$workRoot\deny\secret")) '1/権限なし'
        # 種類を判定できないリンクは、読み取りが禁止されていると普通のフォルダに見える（判定できない形を書き換えたコピーで作る）
        $linkDeny = "$workRoot\linkdeny"
        New-TextFile "$linkDeny\real\x.crt" (ConvertTo-Pem $rawOf['valid'])
        $null = New-Item -ItemType Junction -Path "$linkDeny\j" -Target "$linkDeny\real"
        $junctions.Add("$linkDeny\j")
        Deny-Access "$linkDeny\real" '(RD)'
        $patched = New-PatchedTool 'unknownlink' 'if (($Entry.Attributes -band $rpFlag) -ne $rpFlag) { return $null }' 'return $null'
        $r = Invoke-Tool @('-Path', $linkDeny) $patched
        Assert-Equal '読み取りが禁止されたリンクは権限なしに数えない' ('{0}/{1}/{2}' -f (Get-SkipReason $r "$linkDeny\j"), (Get-SkipDetail $r "$linkDeny\j"), (Get-SkipReason $r "$linkDeny\real")) 'リンク/読み取りが禁止されたリンク/権限なし'
        $r = Invoke-Tool @('-Path', "$linkDeny\j") $patched
        Assert-Equal '指定したリンク自体が読めないときは権限なし' (Get-SkipReason $r "$linkDeny\j") '権限なし'
    }

    Write-Section 'リンク'
    $real = 'C:\Users\Default'
    if ([IO.Directory]::Exists("$real\Application Data") -or [IO.Directory]::Exists("$real\AppData")) {
        $r = Invoke-Tool @('-Path', $real)
        $links = @($r.Skipped | Where-Object { $_.'詳細' -eq 'Windows の互換用リンク' })
        Assert-Equal 'Windows の互換用リンクをたどらない' ($links.Count -gt 0) 'True'
        Assert-Equal '互換用リンクの理由はリンク' (@($links | Where-Object { $_.'理由' -ne 'リンク' }).Count) 0
    } else {
        Write-Skip 'Windows の互換用リンク' 'C:\Users\Default がないため'
    }
    $symDir = "$workRoot\sym"
    New-TextFile "$symDir\real\a.crt" (ConvertTo-Pem $rawOf['valid'])
    $symMade = $false
    try {
        $null = New-Item -ItemType SymbolicLink -Path "$symDir\link" -Target "$symDir\real" -ErrorAction Stop
        $symMade = $true
    } catch { }
    if ($symMade) {
        $r = Invoke-Tool @('-Path', $symDir)
        Assert-Equal 'シンボリックリンクをたどらない' ('{0}/{1}/{2}' -f (Get-SkipReason $r "$symDir\link"), (Get-SkipDetail $r "$symDir\link"), @($r.Certs).Count) 'リンク/シンボリックリンク/1'
        [IO.Directory]::Delete("$symDir\link")
    } else {
        Write-Skip 'シンボリックリンク' '作れないため（管理者か開発者モードが必要）'
    }
    $r = Invoke-Tool @('-Path', "$files\jloop\pem.crt")
    Assert-Equal '指定したファイルの途中のリンクは、指定なのでたどる' (@($r.Certs).Count) 1

    Write-Section '接続先'
    $serverDir = "$workRoot\server"
    [void][IO.Directory]::CreateDirectory($serverDir)
    foreach ($name in 'near', 'local', 'renewOld', 'renewNew', 'expired') { [IO.File]::WriteAllBytes("$serverDir\$name.pfx", $pfxOf[$name]) }
    $serverConf = "$serverDir\listeners.txt"
    [IO.File]::WriteAllLines($serverConf, [string[]]@("near=$serverDir\near.pfx", "local=$serverDir\local.pfx", "old=$serverDir\renewOld.pfx", "new=$serverDir\renewNew.pfx", "expired=$serverDir\expired.pfx", 'silent=silent', 'close=close'))
    $portFile = "$serverDir\ports.txt"
    $serverFile = "$serverDir\server.ps1"
    $serverCode = @'
param([string]$ConfPath, [string]$PortFile, [string]$StopFile)
$entries = New-Object System.Collections.ArrayList
foreach ($line in [IO.File]::ReadAllLines($ConfPath)) {
    $kv = $line -split '=', 2
    $listener = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::IPv6Any), 0
    $listener.Server.DualMode = $true
    $listener.Start()
    $cert = $null
    if ($kv[1] -ne 'silent' -and $kv[1] -ne 'close') {
        $cert = New-Object Security.Cryptography.X509Certificates.X509Certificate2 ([byte[]][IO.File]::ReadAllBytes($kv[1]), [string]$null, [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)
    }
    [void]$entries.Add([pscustomobject]@{ Name = $kv[0]; Mode = $kv[1]; Listener = $listener; Cert = $cert })
}
$lines = @($entries | ForEach-Object { '{0}={1}' -f $_.Name, $_.Listener.LocalEndpoint.Port })
[IO.File]::WriteAllLines($PortFile + '.tmp', [string[]]$lines)
[IO.File]::Move($PortFile + '.tmp', $PortFile)
$held = New-Object System.Collections.ArrayList
while (-not [IO.File]::Exists($StopFile)) {
    foreach ($entry in $entries) {
        while ($entry.Listener.Pending()) {
            $client = $entry.Listener.AcceptTcpClient()
            if ($entry.Mode -eq 'silent') { [void]$held.Add($client); continue }
            if ($entry.Mode -eq 'close') { $client.Close(); continue }
            try {
                $client.ReceiveTimeout = 5000
                $client.SendTimeout = 5000
                $tls = New-Object Net.Security.SslStream($client.GetStream())
                $tls.AuthenticateAsServer($entry.Cert)
                $tls.Dispose()
            } catch { }
            $client.Close()
        }
    }
    Start-Sleep -Milliseconds 20
}
foreach ($entry in $entries) { if ($entry.Cert) { $entry.Cert.Reset() }; $entry.Listener.Stop() }
foreach ($client in $held) { $client.Close() }
'@
    [IO.File]::WriteAllText($serverFile, $serverCode, $utf8Bom)
    $server = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$serverFile`"", "`"$serverConf`"", "`"$portFile`"", "`"$serverStop`"") -WindowStyle Hidden -PassThru
    $deadline = (Get-Date).AddSeconds(30)
    while (-not (Test-Path -LiteralPath $portFile) -and (Get-Date) -lt $deadline -and -not $server.HasExited) { Start-Sleep -Milliseconds 100 }
    $ports = @{}
    if (Test-Path -LiteralPath $portFile) { foreach ($line in [IO.File]::ReadAllLines($portFile)) { $kv = $line -split '=', 2; $ports[$kv[0]] = $kv[1] } }
    Assert-Equal 'テスト用の TLS サーバーの準備' $ports.Count 7
    # 閉じているポートを探す（一度開いて閉じる）
    $probe = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback), 0
    $probe.Start()
    $closedPort = $probe.LocalEndpoint.Port
    $probe.Stop()
    $pNear = $ports['near']

    $tlsRun = Invoke-Tool @('-Target', "localhost:$pNear,127.0.0.1:$pNear,[::1]:$pNear,127.0.0.1:$closedPort", '-Timeout', '3')
    Assert-Equal '-Target だけならストアは調べない' (@($tlsRun.Certs | Where-Object { $_.'種類' -ne '接続先' }).Count) 0
    $localRow = @(Get-PlaceRows $tlsRun "localhost:$pNear")
    Assert-Equal '証明書を受け取る' ('{0}/{1}' -f $localRow.Count, $(if ($localRow.Count) { Get-ThumbText $localRow[0] })) ('1/{0}' -f $thumbs['near'])
    Assert-Equal '状態と種類' $(if ($localRow.Count) { '{0}/{1}/{2}' -f $localRow[0].'状態', $localRow[0].'種類', $localRow[0].'秘密鍵' }) '期限間近/接続先/'
    Assert-Equal '信頼できない発行者と、名前の不一致を備考に書く' $(if ($localRow.Count) { $localRow[0].'備考' }) '名前が一致しない、信頼できない発行者'
    Assert-Equal '画面に備考を添える' (Test-Contains $tlsRun.Log 'CN=wb-near（名前が一致しない、信頼できない発行者）') 'True'
    Assert-Equal 'IPv4 で接続する' (@(Get-PlaceRows $tlsRun "127.0.0.1:$pNear").Count) 1
    Assert-Equal 'IPv6 で接続する' (@(Get-PlaceRows $tlsRun "[::1]:$pNear").Count) 1
    Assert-Equal '接続できないものは記録して続ける' (Get-SkipReason $tlsRun "127.0.0.1:$closedPort") '接続できない'
    Assert-Equal '期限間近があるので 2' $tlsRun.Code 2
    Assert-Equal '対象の行' $tlsRun.Summary[2] "対象: localhost:$pNear, 127.0.0.1:$pNear, [::1]:$pNear, 127.0.0.1:$closedPort"
    Assert-Equal 'スキップの行' (Get-SummaryLine $tlsRun 'スキップ ') 'スキップ 1 件（接続できない）'
    $r = Invoke-Tool @('-Target', "localhost:$($ports['local'])")
    Assert-Equal '名前が一致すれば「名前が一致しない」を書かない' ('{0}/{1}' -f $r.Code, @($r.Certs)[0].'備考') '0/信頼できない発行者'
    $r = Invoke-Tool @('-Target', "127.0.0.1:$($ports['old']),127.0.0.1:$($ports['new'])")
    Assert-Equal '接続先の証明書は更新済みにならない' (Get-PlaceStatus $r "127.0.0.1:$($ports['old'])") '期限切れ'
    $r = Invoke-Tool @('-Target', "127.0.0.1:$($ports['expired'])")
    Assert-Equal '期限切れの証明書も受け取る' ('{0}/{1}' -f $r.Code, (Get-PlaceStatus $r "127.0.0.1:$($ports['expired'])")) '2/期限切れ'
    $r = Invoke-Tool @('-Target', "127.0.0.1:$pNear", '-Exclude', 'CN=wb-near')
    Assert-Equal '接続先の証明書も除外できる' ('{0}/{1}' -f $r.Code, (Get-SkipReason $r "127.0.0.1:${pNear}: CN=wb-near")) '0/除外指定'
    $r = Invoke-Tool @('-Target', "127.0.0.1:$pNear,127.0.0.1:$pNear,LOCALHOST:$pNear,localhost:$pNear")
    Assert-Equal '同じ接続先は 1 回だけ' @($r.Certs).Count 2
    $r = Invoke-Tool @('-Target', 'localhost,localhost:443', '-Timeout', '1')
    Assert-Equal 'ポートの書き方が違う同じ接続先も 1 回だけ' (@($r.Certs).Count + @($r.Skipped).Count) 1
    $r = Invoke-Tool @('-Target', 'LOCALHOST,localhost:443', '-Timeout', '1')
    Assert-Equal '大文字小文字だけ違う同じ接続先も 1 回だけ' (@($r.Certs).Count + @($r.Skipped).Count) 1
    $r = Invoke-NoSave @('-Target', 'localhost', '-Timeout', '1')
    Assert-Equal 'ポートを省くと 443' (Test-Contains $r.Log '対象: localhost:443') 'True'
    $r = Invoke-NoSave @('-Target', '::1', '-Timeout', '1')
    Assert-Equal '括弧のない IPv6 アドレス' (Test-Contains $r.Log '対象: [::1]:443') 'True'
    $r = Invoke-NoSave @('-Target', '[::1]', '-Timeout', '1')
    Assert-Equal '括弧だけの IPv6 アドレス' (Test-Contains $r.Log '対象: [::1]:443') 'True'
    $closedRun = Invoke-Tool @('-Target', "127.0.0.1:$closedPort", '-Timeout', '3')
    Assert-Equal '接続できないだけなら 1' ('{0}/{1}' -f $closedRun.Code, (Test-Contains (Get-SkipDetail $closedRun "127.0.0.1:$closedPort") "`n")) '1/False'
    $bothRun = Invoke-Tool @('-Target', "127.0.0.1:$pNear", '-Store', $store2Text, '-Path', "$workRoot\clean")
    Assert-Equal '-Store と一緒なら両方調べる' ('{0}/{1}/{2}' -f @(Get-PlaceRows $bothRun "127.0.0.1:$pNear").Count, @(Get-CertRow $bothRun 'crossNew').Count, @(Get-PlaceRows $bothRun "$workRoot\clean\ok.crt").Count) '1/1/1'
    Assert-Equal '対象の行はストア・接続先・ファイルの順' $bothRun.Summary[2] "対象: $store2Text, 127.0.0.1:$pNear, $workRoot\clean"
    # TLS で応答しないサーバー
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = Invoke-Tool @('-Target', "127.0.0.1:$($ports['silent'])", '-Timeout', '1')
    $sw.Stop()
    Assert-Equal '応答しないサーバーは -Timeout で諦める' ('{0}/{1}/{2}' -f $r.Code, (Get-SkipReason $r "127.0.0.1:$($ports['silent'])"), ($sw.Elapsed.TotalSeconds -lt 20)) '1/接続できない/True'
    $r = Invoke-Tool @('-Target', "127.0.0.1:$($ports['close'])", '-Timeout', '3')
    Assert-Equal 'すぐ切断するサーバー' ('{0}/{1}' -f $r.Code, (Get-SkipReason $r "127.0.0.1:$($ports['close'])")) '1/接続できない'
    # 本物の環境では作りにくい条件は、書き換えたコピーで確かめる
    $patched = New-PatchedTool 'connecttimeout' 'if (-not $pending.AsyncWaitHandle.WaitOne($Timeout * 1000))' 'if ($true)'
    $r = Invoke-Tool @('-Target', "127.0.0.1:$pNear", '-Timeout', '7') $patched
    Assert-Equal '接続のタイムアウト' (Get-SkipDetail $r "127.0.0.1:$pNear") '7 秒以内に接続できませんでした'
    $patched = New-PatchedTool 'nodns' '$pending = $client.BeginConnect($Item.Host, $Item.Port, $null, $null)' 'throw (New-Object Net.Sockets.SocketException 11001)'
    $r = Invoke-Tool @('-Target', 'no-such-host:443') $patched
    Assert-Equal '名前を解決できない' ('{0}/{1}' -f $r.Code, (Get-SkipReason $r 'no-such-host:443')) '1/接続できない'
    $patched = New-PatchedTool 'noipv6' '$client = New-Object Net.Sockets.TcpClient ([Net.Sockets.AddressFamily]::InterNetworkV6)' 'throw ''no IPv6'''
    Assert-Equal 'IPv6 を使えない PC では IPv4 で接続する' (@((Invoke-Tool @('-Target', "127.0.0.1:$pNear") $patched).Certs).Count) 1
    $chainLine = 'if ($Chain) { $script:tlsChain = @($Chain.ChainStatus | ForEach-Object { [string]$_.Status }) }'
    $patched = New-PatchedTool 'chainok' $chainLine '$script:tlsChain = @(''NoError'', ''NotTimeValid'', ''RevocationStatusUnknown'', ''OfflineRevocation'')'
    Assert-Equal '期限切れと失効の確認不可だけなら信頼できない発行者としない' (@((Invoke-Tool @('-Target', "localhost:$($ports['local'])") $patched).Certs)[0].'備考') ''
    $patched = New-PatchedTool 'chainbad' $chainLine '$script:tlsChain = @(''NotTimeValid'', ''PartialChain'')'
    Assert-Equal 'そのほかの問題があれば信頼できない発行者' (@((Invoke-Tool @('-Target', "localhost:$($ports['local'])") $patched).Certs)[0].'備考') '信頼できない発行者'

    Write-Section 'IIS の紐付け'
    $iisXml = '<site name="Portal"><bindings><binding protocol="http" bindingInformation="*:80:" /><binding protocol="https" bindingInformation="*:443:portal.example" certificateHash="{0}" /></bindings></site>' -f $thumbs['iisOld'].ToLowerInvariant()
    $ftpXml = '<site name="Files"><ftpServer><security><ssl serverCertHash="{0}" /></security></ftpServer></site>' -f $thumbs['near']
    $unknownXml = '<site name="Unknown"><bindings><binding protocol="https" bindingInformation="192.0.2.77:64999:" /></bindings></site>'
    $centralXml = '<site name="Central"><bindings><binding protocol="https" bindingInformation="192.0.2.77:64999:" sslFlags="3" /></bindings></site>'
    $centralWordXml = '<site name="Central2"><bindings><binding protocol="https" bindingInformation="192.0.2.77:64999:" sslFlags="Sni,CentralCert" /></bindings></site>'
    $sniOnlyXml = '<site name="Sni"><bindings><binding protocol="https" bindingInformation="192.0.2.77:64999:" sslFlags="1" /></bindings></site>'

    $iis = Invoke-WithWindir (New-FakeWindir 'ok' (New-HostConfig ($iisXml + $ftpXml)) '') @('-Store', $storeText)
    Assert-Equal 'サイトで使っている古い証明書は更新済みにしない' (Get-Status $iis 'iisOld') '期限切れ'
    Assert-Equal 'サイトとバインド' (Get-Cell $iis 'iisOld' 'IIS のサイト') 'Portal (*:443:portal.example)'
    Assert-Equal '画面にサイトの名前を出す' (Test-Contains $iis.Log 'CN=wb-iis（IIS: Portal）') 'True'
    Assert-Equal 'FTP サイトの証明書' (Get-Cell $iis 'near' 'IIS のサイト') 'Files (FTP)'
    Assert-Equal '使っていない古い証明書は更新済み' (Get-Status $iis 'renewOld') '更新済み'
    Assert-Equal '読めなかったと出ない' (Test-Contains $iis.Log 'IIS の設定を読めなかった') 'False'
    # 1 つの証明書を複数のサイトで使う
    $multiXml = '<site name="B"><bindings><binding protocol="https" bindingInformation="*:443:b.example" certificateHash="{0}" /><binding protocol="https" bindingInformation="*:8443:b.example" certificateHash="{0}" /></bindings></site><site name="A"><bindings><binding protocol="https" bindingInformation="*:443:a.example" certificateHash="{0}" /></bindings></site>' -f $thumbs['near']
    $r = Invoke-WithWindir (New-FakeWindir 'multi' (New-HostConfig $multiXml) '') @('-Store', $storeText)
    Assert-Equal '複数のバインドは , でつなぐ' (Get-Cell $r 'near' 'IIS のサイト') 'B (*:443:b.example), B (*:8443:b.example), A (*:443:a.example)'
    Assert-Equal '画面のサイトの名前は重ねない' (Test-Contains $r.Log 'CN=wb-near（IIS: B, A）') 'True'
    $dashXml = '<site name="-site"><bindings><binding protocol="https" bindingInformation="*:443:" certificateHash="{0}" /></bindings></site>' -f $thumbs['near']
    $r = Invoke-WithWindir (New-FakeWindir 'dash' (New-HostConfig $dashXml) '') @('-Store', $storeText)
    Assert-Equal 'Excel で変わるサイトの名前は ="..." で守る' (Get-Cell $r 'near' 'IIS のサイト') '="-site (*:443:)"'
    # ファイルの証明書には書かない
    $r = Invoke-WithWindir (New-FakeWindir 'okfile' (New-HostConfig $ftpXml) '') @('-Store', $storeText, '-Path', "$files\der.cer")
    $nearInStore = @(Get-PlaceRows $r $storeText | Where-Object { (Get-ThumbText $_) -eq $thumbs['near'] })[0]
    Assert-Equal 'IIS のサイトはストアの証明書だけに書く' ('{0}/{1}' -f $nearInStore.'IIS のサイト', @(Get-PlaceRows $r "$files\der.cer")[0].'IIS のサイト') 'Files (FTP)/'

    $unknown = Invoke-WithWindir (New-FakeWindir 'unknown' (New-HostConfig $unknownXml) '') @('-Store', $storeText)
    Assert-Equal '証明書が分からないバインドがあると更新済みにしない' ('{0}/{1}' -f (Get-Status $unknown 'renewOld'), (Test-Contains $unknown.Log 'IIS の設定を読めなかった')) '期限切れ/True'
    foreach ($case in @(@('管理者でないときは案内を添える', '$false', 'True'), @('管理者のときは案内を添えない', '$true', 'False'))) {
        $patched = New-PatchedTool "iisadmin$($case[1])" '$outText = $null' ("`$isAdmin = {0}; `$outText = `$null" -f $case[1])
        $r = Invoke-WithWindir (New-FakeWindir 'unknown' (New-HostConfig $unknownXml) '') @('-Store', $store2Text) $patched
        Assert-Equal $case[0] (Test-Contains $r.Log 'ストアの証明書は更新済みにしていません（管理者として実行すると読めることがあります）') $case[2]
    }
    $central = Invoke-WithWindir (New-FakeWindir 'central' (New-HostConfig $centralXml) '') @('-Store', $storeText)
    Assert-Equal '中央証明書ストアのバインドは気にしない' ('{0}/{1}' -f (Get-Status $central 'renewOld'), (Test-Contains $central.Log 'IIS の設定を読めなかった')) '更新済み/False'
    $central = Invoke-WithWindir (New-FakeWindir 'central2' (New-HostConfig $centralWordXml) '') @('-Store', $storeText)
    Assert-Equal '中央証明書ストアのバインド（文字の書き方）' (Test-Contains $central.Log 'IIS の設定を読めなかった') 'False'
    $r = Invoke-WithWindir (New-FakeWindir 'snionly' (New-HostConfig $sniOnlyXml) '') @('-Store', $storeText)
    Assert-Equal 'SNI だけのバインドは中央証明書ストアではない' (Test-Contains $r.Log 'IIS の設定を読めなかった') 'True'
    $shared = Invoke-WithWindir (New-FakeWindir 'shared' (New-HostConfig $iisXml) '<configuration><configurationRedirection enabled="true" path="\\fs\iis" /></configuration>') @('-Store', $storeText)
    Assert-Equal '共有構成は読めなかった扱い' ('{0}/{1}' -f (Get-Status $shared 'renewOld'), (Test-Contains $shared.Log 'IIS の設定を読めなかった')) '期限切れ/True'
    $sharedOff = Invoke-WithWindir (New-FakeWindir 'sharedoff' (New-HostConfig $iisXml) '<configuration><configurationRedirection enabled="false" /></configuration>') @('-Store', $storeText)
    Assert-Equal '共有構成が無効なら読む' ('{0}/{1}' -f (Get-Status $sharedOff 'renewOld'), (Test-Contains $sharedOff.Log 'IIS の設定を読めなかった')) '更新済み/False'
    $brokenXml = Invoke-WithWindir (New-FakeWindir 'broken' '<configuration>' '') @('-Store', $storeText)
    Assert-Equal '壊れた設定でも止まらない' ('{0}/{1}' -f (Get-Status $brokenXml 'renewOld'), (Test-Contains $brokenXml.Log 'エラーで止まりました')) '期限切れ/False'
    $noFile = Invoke-WithWindir (New-FakeWindir 'nofile' '' '') @('-Store', $storeText)
    Assert-Equal '設定ファイルがなければ IIS なし' ('{0}/{1}' -f (Get-Status $noFile 'renewOld'), (Test-Contains $noFile.Log 'IIS の設定を読めなかった')) '更新済み/False'
    $noIisDir = Join-Path $workRoot 'windir_none'
    [void][IO.Directory]::CreateDirectory($noIisDir)
    $noIis = Invoke-WithWindir $noIisDir @('-Store', $storeText)
    Assert-Equal 'inetsrv がなければ IIS なし' (Test-Contains $noIis.Log 'IIS の設定を読めなかった') 'False'
    # 読めないときと、-Path だけのとき
    $brokenDir = New-FakeWindir 'broken' '<configuration>' ''
    $r = Invoke-WithWindir $brokenDir @('-Path', "$workRoot\onlyrenew")
    Assert-Equal '-Path だけなら IIS を読まない' ('{0}/{1}' -f (Test-Contains $r.Log 'IIS の設定を読めなかった'), (Get-PlaceStatus $r "$workRoot\onlyrenew\old.cer")) 'False/更新済み'
    $r = Invoke-WithWindir $brokenDir @('-Store', $storeText, '-Path', "$workRoot\onlyrenew")
    $renewInStore = @(Get-PlaceRows $r $storeText | Where-Object { (Get-ThumbText $_) -eq $thumbs['renewOld'] })[0]
    Assert-Equal 'IIS を読めなくても、ファイルは更新済みにする' ('{0}/{1}' -f $renewInStore.'状態', (Get-PlaceStatus $r "$workRoot\onlyrenew\old.cer")) '期限切れ/更新済み'
    $denyDir = New-FakeWindir 'denied' (New-HostConfig $iisXml) ''
    $denyFile = Join-Path $denyDir 'System32\inetsrv\config\applicationHost.config'
    Deny-Access $denyFile '(R)'
    $r = Invoke-WithWindir $denyDir @('-Store', $storeText)
    Assert-Equal '読めない applicationHost.config' ('{0}/{1}' -f (Get-Status $r 'renewOld'), (Test-Contains $r.Log 'IIS の設定を読めなかった')) '期限切れ/True'
    # 64 ビットの OS の 32 ビットの PowerShell では Sysnative を見る
    $patched = New-PatchedTool 'wow64' '-not [Environment]::Is64BitProcess' '$true'
    $r = Invoke-WithWindir (New-FakeWindir 'wow' (New-HostConfig $iisXml) '' 'Sysnative') @('-Store', $storeText) $patched
    Assert-Equal '32 ビットの PowerShell でも IIS の設定を読む' (Get-Cell $r 'iisOld' 'IIS のサイト') 'Portal (*:443:portal.example)'
    # HTTP.sys の登録（レジストリは本物を変えず、読む処理を差し替える）
    $hexOf = { param($Name) (($thumbs[$Name] -split '(..)' | Where-Object { $_ }) | ForEach-Object { '0x' + $_ }) -join ',' }
    $fakeRegistry = '$fakeBytes = @{{ ''SslBindingInfo\0.0.0.0:443'' = [byte[]]@({0}); ''SslBindingInfo\192.0.2.5:8443'' = [byte[]]@({1}); ''SslSniBindingInfo\sni.example:443'' = [byte[]]@({2}) }}[$key]; $registry = if ($fakeBytes) {{ $fakeEntry = [pscustomobject]@{{ Bytes = $fakeBytes }}; $fakeEntry | Add-Member ScriptMethod GetValue {{ param($Name) if ($Name -eq ''SslCertHash'') {{ , $this.Bytes }} }}; $fakeEntry | Add-Member ScriptMethod Close {{ }}; $fakeEntry }} else {{ $null }}' -f (& $hexOf 'iisOld'), (& $hexOf 'near'), (& $hexOf 'expired')
    $patched = New-PatchedTool 'httpsys' '$registry = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(''SYSTEM\CurrentControlSet\Services\HTTP\Parameters\'' + $key)' $fakeRegistry
    $httpXml = '<site name="AnyIp"><bindings><binding protocol="https" bindingInformation="*:443:" /></bindings></site><site name="OneIp"><bindings><binding protocol="https" bindingInformation="192.0.2.5:8443:" /></bindings></site><site name="WithSni"><bindings><binding protocol="https" bindingInformation="*:443:sni.example" /></bindings></site>'
    $r = Invoke-WithWindir (New-FakeWindir 'httpsys' (New-HostConfig $httpXml) '') @('-Store', $storeText) $patched
    Assert-Equal 'HTTP.sys の登録: * は 0.0.0.0' (Get-Cell $r 'iisOld' 'IIS のサイト') 'AnyIp (*:443:)'
    Assert-Equal 'HTTP.sys の登録: IP アドレスを指定' (Get-Cell $r 'near' 'IIS のサイト') 'OneIp (192.0.2.5:8443:)'
    Assert-Equal 'HTTP.sys の登録: SNI はホスト名で探す' (Get-Cell $r 'expired' 'IIS のサイト') 'WithSni (*:443:sni.example)'
    Assert-Equal 'HTTP.sys で分かれば読めなかったと出ない' (Test-Contains $r.Log 'IIS の設定を読めなかった') 'False'

    Write-Section '設定ファイル'
    $conf = "$workRoot\conf.ini"
    [IO.File]::WriteAllText($conf, ("# コメント`r`n; コメント`r`n`r`nDays = 400`r`nPath = clean`r`nPath = only\a,b`r`nExclude = CN=nobody`r`nNoSave = false`r`nTimeout = 3`r`n"), $utf8NoBom)
    $confRun = Invoke-Tool @('-ConfigFile', $conf)
    Assert-Equal '複数の行の Path と、カンマを含む Path' (@($confRun.Certs).Count) 2
    Assert-Equal 'Days を設定ファイルから読む' (@($confRun.Certs)[0].'状態') '期限間近'
    Assert-Equal '設定ファイルの表示' (Get-SummaryLine $confRun '設定ファイル: ') "設定ファイル: $conf"
    Assert-Equal '設定ファイルの行は対象の次' $confRun.Summary[3] "設定ファイル: $conf"
    Assert-Equal 'コマンドラインを優先する' (Invoke-Tool @('-ConfigFile', $conf, '-Days', '1')).Code 0
    $r = Invoke-Tool @('-ConfigFile', $conf, '-Path', "$workRoot\digits")
    Assert-Equal 'コマンドラインの -Path は設定ファイルと合わせない' @($r.Certs).Count 1

    $env:WB_TEST_SUBJECT = 'wb-valid'
    [IO.File]::WriteAllText("$workRoot\envconf.ini", "Path = clean`r`nExclude = CN=%WB_TEST_SUBJECT%`r`n", $utf8NoBom)
    Assert-Equal 'Exclude の環境変数' (@((Invoke-Tool @('-ConfigFile', "$workRoot\envconf.ini")).Certs).Count) 0
    $env:WB_TEST_SUBJECT = "$workRoot\clean"
    [IO.File]::WriteAllText("$workRoot\envpath.ini", "Path = %WB_TEST_SUBJECT%`r`nNoSave = true`r`n", $utf8NoBom)
    Assert-Equal 'Path の環境変数' (Test-Contains (Invoke-Plain @('-ConfigFile', "$workRoot\envpath.ini")).Log "対象: $workRoot\clean") 'True'
    $env:WB_TEST_SUBJECT = "127.0.0.1:$closedPort"
    [IO.File]::WriteAllText("$workRoot\envtarget.ini", "Target = %WB_TEST_SUBJECT%`r`nTimeout = 1`r`nNoSave = true`r`n", $utf8NoBom)
    Assert-Equal 'Target の環境変数' (Test-Contains (Invoke-Plain @('-ConfigFile', "$workRoot\envtarget.ini")).Log "対象: 127.0.0.1:$closedPort") 'True'
    Remove-Item Env:\WB_TEST_SUBJECT

    [IO.File]::WriteAllText("$workRoot\store.ini", "Store = $storeText`r`nStore = $store2Text`r`nNoSave = true`r`n", $utf8NoBom)
    Assert-Equal 'Store を複数の行に書ける' (Test-Contains (Invoke-Plain @('-ConfigFile', "$workRoot\store.ini")).Log "対象: $storeText, $store2Text") 'True'
    [IO.File]::WriteAllText("$workRoot\storecomma.ini", "Store = `"$storeText`",`"$store2Text`"`r`nNoSave = true`r`n", $utf8NoBom)
    Assert-Equal '1 つずつ囲んだ書き方' (Test-Contains (Invoke-Plain @('-ConfigFile', "$workRoot\storecomma.ini")).Log "対象: $storeText, $store2Text") 'True'
    Assert-Equal '-AllStores は設定ファイルの Store より優先する' (Test-Contains (Invoke-NoSave @('-ConfigFile', "$workRoot\store.ini", '-AllStores')).Log '対象: すべてのストア') 'True'
    [IO.File]::WriteAllText("$workRoot\all.ini", "AllStores = true`r`n", $utf8NoBom)
    Assert-Equal 'AllStores を設定ファイルに書ける' (Test-Contains (Invoke-NoSave @('-ConfigFile', "$workRoot\all.ini")).Log '対象: すべてのストア') 'True'
    Assert-Equal '-Store は設定ファイルの AllStores より優先する' (Test-Contains (Invoke-NoSave @('-ConfigFile', "$workRoot\all.ini", '-Store', $store2Text)).Log "対象: $store2Text") 'True'
    [IO.File]::WriteAllText("$workRoot\excomma.ini", "Store = $storeText`r`nExclude = *Example, Inc.*`r`nExclude = *wb-near*,*wb-expired*`r`n", $utf8NoBom)
    $r = Invoke-Tool @('-ConfigFile', "$workRoot\excomma.ini")
    Assert-Equal 'Exclude は , の後にスペースがあれば区切らない・複数の行' ('{0}/{1}/{2}/{3}' -f @(Get-CertRow $r 'comma').Count, @(Get-CertRow $r 'near').Count, @(Get-CertRow $r 'expired').Count, @(Get-CertRow $r 'valid').Count) '0/0/0/1'
    [IO.File]::WriteAllText("$workRoot\exquote.ini", "Store = $storeText`r`nExclude = `"*wb-near*`",`"*wb-expired*`"`r`n", $utf8NoBom)
    $r = Invoke-Tool @('-ConfigFile', "$workRoot\exquote.ini")
    Assert-Equal 'Exclude を 1 つずつ引用符で囲む' ('{0}/{1}/{2}' -f @(Get-CertRow $r 'near').Count, @(Get-CertRow $r 'expired').Count, @(Get-CertRow $r 'valid').Count) '0/0/1'
    [IO.File]::WriteAllText("$workRoot\outrel.ini", "Path = clean`r`nOutputFolder = 記録\%COMPUTERNAME%`r`n", $utf8NoBom)
    $r = Invoke-From $workRoot @('-ConfigFile', "$workRoot\outrel.ini")
    Assert-Equal '相対の OutputFolder は設定ファイルのあるフォルダが基準・環境変数' ('{0}/{1}' -f $r.Code, [bool](Get-RunDir "$workRoot\記録\$env:COMPUTERNAME")) '0/True'
    [IO.File]::WriteAllText("$workRoot\notime.ini", "Target = 127.0.0.1:$closedPort`r`nTimeout = 1`r`nNoSave = true`r`n", $utf8NoBom)
    Assert-Equal 'Target と Timeout と NoSave = true' ('{0}/{1}' -f (Invoke-Plain @('-ConfigFile', "$workRoot\notime.ini")).Code, (Test-Contains (Invoke-Plain @('-ConfigFile', "$workRoot\notime.ini")).Log '保存先')) '1/False'
    # 文字コードを間違えて読むと Path が読めず、既定のストアを調べてしまうので、終了コードだけでなく対象の行を見る
    [IO.File]::WriteAllText("$workRoot\enc-sjis.ini", "# 日本語のコメント`r`nPath = 幅`r`nNoSave = true`r`n", [Text.Encoding]::GetEncoding(932))
    [IO.File]::WriteAllText("$workRoot\enc-le.ini", "Path = 幅`r`nNoSave = true`r`n", [Text.Encoding]::Unicode)
    [IO.File]::WriteAllText("$workRoot\enc-be.ini", "Path = 幅`r`nNoSave = true`r`n", [Text.Encoding]::BigEndianUnicode)
    [IO.File]::WriteAllText("$workRoot\enc-bom.ini", "Path = `"幅`"`r`nNoSave = 'true'`r`nExclude =`r`n", $utf8Bom)
    foreach ($case in @(@('Shift_JIS の設定ファイル', 'enc-sjis.ini'), @('UTF-16 LE の設定ファイル', 'enc-le.ini'), @('UTF-16 BE の設定ファイル', 'enc-be.ini'), @('BOM 付きと引用符、空の値', 'enc-bom.ini'))) {
        $r = Invoke-Plain @('-ConfigFile', "$workRoot\$($case[1])")
        Assert-Equal $case[0] ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $workRoot\幅")) '2/True'
    }
    $r = Invoke-NoSave @('-ConfigFile', "$workRoot\nothing.ini")
    Assert-Equal 'ない設定ファイル' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "設定ファイルが見つかりません: $workRoot\nothing.ini")) '1/True'
    [IO.File]::WriteAllText("$workRoot\memo.ini.txt", 'Days = 1', $utf8NoBom)
    $r = Invoke-NoSave @('-ConfigFile', "$workRoot\memo.ini")
    Assert-Equal '拡張子が .txt のときの案内' (Test-Contains $r.Log '拡張子が .txt になっていませんか') 'True'
    [IO.File]::WriteAllText("$workRoot\badpath.ini", "# 1`r`nPath = no_such_folder`r`n", $utf8NoBom)
    $r = Invoke-NoSave @('-ConfigFile', "$workRoot\badpath.ini")
    Assert-Equal 'Path が見つからないときは行番号を添える' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '（設定ファイルの 2 行目）')) '1/True'
    [IO.File]::WriteAllText("$workRoot\badout.ini", "Path = clean`r`nOutputFolder = $workRoot\notdir2`r`n", $utf8NoBom)
    New-TextFile "$workRoot\notdir2" 'x'
    $r = Invoke-Plain @('-ConfigFile', "$workRoot\badout.ini")
    Assert-Equal 'OutputFolder に書き込めないときは行番号を添える' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '（設定ファイルの 2 行目）')) '0/True'

    $badCases = @(
        @('Exlude = a', '知らない項目です: Exlude'),
        @('Exlude = a', '使える項目: Days, Store, AllStores, Target, Timeout, Path, Exclude, OutputFolder, NoSave'),
        @('Days = 0', 'Days は 1～3650 の数値で指定してください: 0'),
        @('Days = 3651', 'Days は 1～3650 の数値で指定してください: 3651'),
        @('Days = abc', 'Days は 1～3650 の数値で指定してください: abc'),
        @('Timeout = 0', 'Timeout は 1～60 の数値で指定してください: 0'),
        @('Timeout = 99', 'Timeout は 1～60 の数値で指定してください: 99'),
        @('NoSave = yes', 'NoSave は true か false で指定してください: yes'),
        @('AllStores = 1', 'AllStores は true か false で指定してください: 1'),
        @('Store = My', 'Store は LocalMachine\My の形で指定してください: My'),
        @('Target = a:0', 'Target は ホスト か ホスト:ポート'),
        @('Days 3', '「項目 = 値」の形になっていません'),
        @("Days = 1`r`nDays = 2", 'Days が 2 回書かれています（1 行目と 2 行目）'),
        @("Timeout = 1`r`nTimeout = 2", 'Timeout が 2 回書かれています'),
        @('Path = %NO_SUCH_VARIABLE_X%', '定義されていない環境変数です: %NO_SUCH_VARIABLE_X%'),
        @('Exclude = %NO_SUCH_VARIABLE_X%', '定義されていない環境変数です: %NO_SUCH_VARIABLE_X%'),
        @('Target = %NO_SUCH_VARIABLE_X%', '定義されていない環境変数です: %NO_SUCH_VARIABLE_X%'),
        @('Exclude = a"b', 'Exclude の値に引用符（"）が残っています: a"b'),
        @('OutputFolder = a"b', 'OutputFolder の値に引用符（"）が残っています: a"b'),
        @("# 1`r`nStore = $storeText`r`nStore = LocalMachine\NoSuchStoreX", '（設定ファイルの 3 行目）')
    )
    foreach ($case in $badCases) {
        [IO.File]::WriteAllText("$workRoot\bad.ini", $case[0], $utf8NoBom)
        $result = Invoke-NoSave @('-ConfigFile', "$workRoot\bad.ini")
        Assert-Equal "書き間違い: $($case[1])" ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $case[1])) '1/True'
    }
    [IO.File]::WriteAllText("$workRoot\bad.ini", "`r`nDays = 0", $utf8NoBom)
    $r = Invoke-NoSave @('-ConfigFile', "$workRoot\bad.ini")
    Assert-Equal '書き間違いは行番号とファイルを添える' ('{0}/{1}' -f (Test-Contains $r.Log '設定ファイルの 2 行目: '), (Test-Contains $r.Log "ファイル: $workRoot\bad.ini")) 'True/True'

    Write-Section '止まるとき'
    $stops = @(
        @('-Store の形', @('-Store', 'My'), '-Store は LocalMachine\My の形で指定してください: My'),
        @('-Store の場所', @('-Store', 'Other\My'), '-Store は LocalMachine\My の形で指定してください: Other\My'),
        @('-Store の名前がない', @('-Store', 'CurrentUser\'), '-Store は LocalMachine\My の形で指定してください: CurrentUser\'),
        @('ないストア', @('-Store', 'LocalMachine\NoSuchStoreX'), 'ストアが見つかりません: LocalMachine\NoSuchStoreX'),
        @('ないストア（CurrentUser）', @('-Store', "$storeText,CurrentUser\NoSuchStoreX"), 'ストアが見つかりません: CurrentUser\NoSuchStoreX'),
        @('-Store と -AllStores', @('-Store', $storeText, '-AllStores'), '-Store と -AllStores は一緒に指定できません'),
        @('-Target のポートが大きい', @('-Target', 'web01:99999'), '-Target は ホスト か ホスト:ポート（ポートは 1～65535）の形で指定してください: web01:99999'),
        @('-Target のポートが 0', @('-Target', 'web01:0'), '-Target は ホスト か ホスト:ポート'),
        @('-Target のポートが空', @('-Target', 'web01:'), '-Target は ホスト か ホスト:ポート'),
        @('-Target のポートが文字', @('-Target', 'web01:abc'), '-Target は ホスト か ホスト:ポート'),
        @('-Target の URL', @('-Target', 'https://www.example.jp/'), '-Target は ホスト か ホスト:ポート'),
        @('-Target のコロンが多い', @('-Target', 'web01:443:x'), '-Target は ホスト か ホスト:ポート'),
        @('-Target の括弧のポートが空', @('-Target', '[::1]:'), '-Target は ホスト か ホスト:ポート'),
        @('-Target のホストが空', @('-Target', ':443'), '-Target は ホスト か ホスト:ポート'),
        @('見つからない -Path', @('-Path', "$workRoot\no_such"), "ファイルかフォルダが見つかりません: $workRoot\no_such")
    )
    foreach ($stop in $stops) {
        $result = Invoke-NoSave $stop[1]
        Assert-Equal $stop[0] ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $stop[2])) '1/True'
    }
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Store 'LocalMachine\NoSuchStoreX' -OutputFolder "$workRoot\out_nostore" 2>&1
    Assert-Equal '止まったときは結果を残さない' ('{0}/{1}' -f $LASTEXITCODE, [bool](Get-RunDir "$workRoot\out_nostore")) '1/False'
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$workRoot\no_such" -OutputFolder "$workRoot\out_nopath" 2>&1
    Assert-Equal '-Path が見つからないときはフォルダを作らない' (Test-Path -LiteralPath "$workRoot\out_nopath") 'False'

    Write-Section 'パスの指定のしかた'
    New-TextFile "$workRoot\m1\a.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\m2\b.crt" (ConvertTo-Pem $rawOf['near'])
    $r = Invoke-From $workRoot @('-Path', 'm1,m2', '-NoSave')
    Assert-Equal '-Path に相対パス' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $workRoot\m1, $workRoot\m2")) '2/True'
    $r = Invoke-From $workRoot @('-Path', '.\m1\..\m2', '-NoSave')
    Assert-Equal '-Path の .. を解決する' (Test-Contains $r.Log "対象: $workRoot\m2") 'True'
    $r = Invoke-NoSave @('-Path', ('\\?\' + "$workRoot\m1"))
    Assert-Equal '-Path に \\?\ を付けてもよい' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $workRoot\m1")) '0/True'
    $r = Invoke-NoSave @('-Path', "$workRoot\m1\")
    Assert-Equal '-Path の最後に \' (Test-Contains $r.Log "対象: $workRoot\m1") 'True'
    $letter = Get-FreeLetter
    if ($letter) {
        $null = & subst.exe "${letter}:" "$workRoot\m2" 2>&1
        if ($LASTEXITCODE -eq 0) {
            $substDrive = "${letter}:"
            $r = Invoke-Tool @('-Path', "$substDrive\")
            Assert-Equal 'ドライブのルート（\ あり）' ('{0}/{1}' -f @(Get-PlaceRows $r "$substDrive\b.crt").Count, $r.Summary[2]) "1/対象: $substDrive\"
            $r = Invoke-Tool @('-Path', $substDrive)
            Assert-Equal 'ドライブのルート（ドライブ文字だけ）' ('{0}/{1}' -f @(Get-PlaceRows $r "$substDrive\b.crt").Count, $r.Summary[2]) "1/対象: $substDrive\"
            $r = Invoke-Tool @('-Path', "$substDrive\b.crt")
            Assert-Equal 'ドライブのルートのファイル' (Get-PlaceStatus $r "$substDrive\b.crt") '期限間近'
        } else {
            Write-Skip 'ドライブのルート' 'subst を使えなかったため'
        }
    } else {
        Write-Skip 'ドライブのルート' '空いているドライブ文字がないため'
    }
    $uncRoot = '\\localhost\{0}$\{1}' -f $workRoot.Substring(0, 1), $workRoot.Substring(3)
    if ([IO.Directory]::Exists("$uncRoot\m1")) {
        $r = Invoke-Tool @('-Path', "$uncRoot\m1,$uncRoot\m2")
        Assert-Equal 'UNC のパス' ('{0}/{1}/{2}' -f $r.Code, (Get-PlaceStatus $r "$uncRoot\m1\a.crt"), (Get-PlaceStatus $r "$uncRoot\m2\b.crt")) '2/有効/期限間近'
        Assert-Equal 'UNC のパスに \\?\ が残らない' ([IO.File]::ReadAllText((Join-Path $r.Dir 'certificates.csv')).Contains('\\?\')) 'False'
        $r = Invoke-Tool @('-Path', ("$uncRoot\files\" + (('y' * 60 + '\') * 5) + 'long.cer'))
        Assert-Equal 'UNC の 260 文字を超えるパス' @($r.Certs).Count 1
    } else {
        Write-Skip 'UNC のパス' '管理共有を読めないため'
    }
    $start = "$workRoot\start"
    $startCases = @(@('隣に一致するフォルダ', '[a]', @('a')), @('一致するフォルダなし', 'b[1]', @()), @('パターンとして正しくない', 'c[d', @()), @('2 つ以上一致', '[ef]', @('e', 'f')))
    foreach ($case in $startCases) {
        $dir = "$start\$($case[1])"
        New-TextFile "$dir\sub\p.crt" (ConvertTo-Pem $rawOf['valid'])
        [IO.File]::WriteAllText("$dir\c.ini", "Path = sub`r`nNoSave = true`r`n", $utf8NoBom)
        # 間違ったフォルダを調べたら分かるように、隣のフォルダには期限間近の証明書を置く
        foreach ($decoy in $case[2]) { New-TextFile "$start\$decoy\sub\wrong.crt" (ConvertTo-Pem $rawOf['near']) }
        $r = Invoke-From $dir @('-Path', 'sub', '-OutputFolder', 'out')
        $x = Get-RunDir "$dir\out"
        Assert-Equal ('[ ] を含むフォルダから -Path と -OutputFolder: {0}' -f $case[0]) ('{0}/{1}' -f $r.Code, [bool]$x) '0/True'
        $r = Invoke-From $dir @('-ConfigFile', 'c.ini')
        Assert-Equal ('[ ] を含むフォルダから -ConfigFile: {0}' -f $case[0]) ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $dir\sub")) '0/True'
    }

    Write-Section '保存'
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$workRoot\m1" 2>&1
    Assert-Equal '既定の保存先はスクリプトのフォルダ' ([bool](Get-RunDir (Split-Path -Parent $tool))) 'True'
    $sameDir = "$workRoot\out_same"
    [void][IO.Directory]::CreateDirectory($sameDir)
    $stampNow = Get-Date
    foreach ($s in 0..40) {
        $older = Join-Path $sameDir ('CertificateExpiry_{0}_{1}' -f $env:COMPUTERNAME, $stampNow.AddSeconds($s).ToString('yyyyMMdd_HHmmss'))
        [void][IO.Directory]::CreateDirectory($older)
        [IO.File]::WriteAllText("$older\keep.txt", 'keep')
    }
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$workRoot\m1" -OutputFolder $sameDir 2>&1
    Assert-Equal '同じ名前のフォルダがあれば _2 を付ける' (@(Get-ChildItem -LiteralPath $sameDir -Directory -Filter 'CertificateExpiry_*_2').Count) 1
    Assert-Equal '前の結果を残す' (@(Get-ChildItem -LiteralPath $sameDir -Recurse -Filter 'keep.txt').Count) 41
    $longOut = "$workRoot\" + ('O' * 120) + '\' + ('P' * 120)
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$workRoot\m1" -OutputFolder $longOut 2>&1
    $longRuns = @(); try { $longRuns = @([IO.Directory]::GetDirectories('\\?\' + $longOut)) } catch { }
    Assert-Equal '-OutputFolder が長くても保存する' ('{0}/{1}' -f $LASTEXITCODE, $longRuns.Count) '0/1'
    if ($longRuns.Count -eq 1) { Assert-Equal '長い保存先のファイル' (@([IO.Directory]::GetFiles($longRuns[0])).Count) 3 }
    $spaceOut = "$workRoot\out 日本語 [x]"
    $r = Invoke-Plain @('-Path', "$workRoot\m1", '-OutputFolder', $spaceOut)
    Assert-Equal '保存先に日本語・スペース・[ ]' ('{0}/{1}' -f $r.Code, [bool](Get-RunDir $spaceOut)) '0/True'
    $r = Invoke-Plain @('-Path', "$workRoot\m1", '-OutputFolder', "$workRoot\out\trail.")
    Assert-Equal '末尾にピリオドが付いた保存先' ('{0}/{1}/{2}' -f $r.Code, [bool](Get-RunDir "$workRoot\out\trail"), (Test-Path -LiteralPath "\\?\$workRoot\out\trail.")) '0/True/False'
    # 保存先に書き込めないときは、画面表示だけで終わる（保存先がファイル）
    New-TextFile "$workRoot\notdir" 'x'
    New-TextFile "$workRoot\m3\near.crt" (ConvertTo-Pem $rawOf['near'])
    $r = Invoke-Plain @('-Path', "$workRoot\m3", '-OutputFolder', "$workRoot\notdir")
    Assert-Equal '書き込めない保存先: 警告して画面に表示' ('{0}/{1}/{2}/{3}' -f $r.Code, (Test-Contains $r.Log '保存先に書き込めないため'), (Test-Contains $r.Log 'CN=wb-near'), (Test-Contains $r.Log '保存先: ')) '2/True/True/False'
    $letter = Get-FreeLetter
    if ($letter) {
        $r = Invoke-Plain @('-Path', "$workRoot\m1", '-OutputFolder', "${letter}:\nope")
        Assert-Equal 'ないドライブ: 警告を出して続ける' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '保存先に書き込めないため')) '0/True'
    } else {
        Write-Skip 'ないドライブ' '空いているドライブ文字がないため'
    }
    $r = Invoke-Plain @('-Path', $manyDir, '-OutputFolder', "$workRoot\notdir")
    Assert-Equal '保存できないときは表をすべて出す' ([regex]::Matches($r.Log, [regex]::Escape(' CN=wb-near')).Count) 1234
    $patched = New-PatchedTool 'savefail' "Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'certificates.csv'))" "throw 'disk full'; Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'certificates.csv'))"
    $r = Invoke-Tool @('-Path', "$workRoot\m3") $patched
    Assert-Equal '保存に失敗: 警告を出して書きかけを残さない' ('{0}/{1}/{2}' -f $r.Code, (Test-Contains $r.Log '結果を保存できませんでした: disk full'), [bool]$r.Dir) '2/True/False'
    Assert-Equal '保存に失敗: 画面には表を出し、保存先は出さない' ('{0}/{1}' -f (Test-Contains $r.Log 'CN=wb-near'), (Test-Contains $r.Log '保存先:')) 'True/False'
    $r = Invoke-Tool @('-Path', $manyDir) $patched
    Assert-Equal '保存に失敗: 表をすべて出す' ([regex]::Matches($r.Log, [regex]::Escape(' CN=wb-near')).Count) 1234
    $patched = New-PatchedTool 'crash' '$sortedCerts = @($certs | Sort-Object' "throw 'boom'; `$sortedCerts = @(`$certs | Sort-Object"
    $r = Invoke-Tool @('-Path', "$workRoot\m3") $patched
    Assert-Equal '想定外のエラー: 結果を残さず、中断とは区別する' ('{0}/{1}/{2}/{3}' -f $r.Code, (Test-Contains $r.Log 'エラーで止まりました。結果は保存していません。'), (Test-Contains $r.Log '中断しました'), [bool]$r.Dir) '1/True/False/False'

    Write-Section 'コマンドプロンプトから起動'
    $cmdHead = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $r = Invoke-ViaCmd ('{0} -NoSave -Path "{1},{2}"' -f $cmdHead, "$workRoot\m1", "$workRoot\m2")
    Assert-Equal 'カンマ区切りの -Path と終了コード 2' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $workRoot\m1, $workRoot\m2")) '2/True'
    $r = Invoke-ViaCmd ("{0} -NoSave -Store '{1},{2}' -Exclude 'CN=wb-near,*wb-expired*' -Days 90" -f $cmdHead, $storeText, $store2Text)
    Assert-Equal "cmd の ' はそのまま届く" ('{0}/{1}/{2}/{3}' -f $r.Code, (Test-Contains $r.Log "対象: $storeText, $store2Text"), (Test-Contains $r.Log '除外: CN=wb-near, *wb-expired*'), (Test-Contains $r.Log '残り 90 日以内')) '2/True/True/True'
    $r = Invoke-ViaCmd ("{0} -NoSave -Target '127.0.0.1:{1},[::1]:{1}' -Timeout 3" -f $cmdHead, $pNear)
    Assert-Equal "cmd の ' で囲んだ -Target" ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: 127.0.0.1:$pNear, [::1]:$pNear")) '2/True'
    $r = Invoke-ViaCmd ('{0} -NoSave -Path "{1}\" -Days 90' -f $cmdHead, "$workRoot\m1")
    Assert-Equal '"パス\" の後ろに引数（止まってヒントを出す）' ('{0}/{1}/{2}' -f $r.Code, (Test-Contains $r.Log '-Path の値に引用符'), (Test-Contains $r.Log '最後に \ を付けないでください')) '1/True/True'
    $r = Invoke-ViaCmd ('{0} -NoSave -Path "{1}\"' -f $cmdHead, "$workRoot\m1")
    Assert-Equal '最後の引数が \"（引用符を外して調べる）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $workRoot\m1")) '0/True'
    $r = Invoke-ViaCmd ('{0} -Path "{1}" -OutputFolder "{2}\" -Days 90' -f $cmdHead, "$workRoot\m1", "$workRoot\out_q")
    Assert-Equal '-OutputFolder の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '-OutputFolder の値に引用符')) '1/True'
    $r = Invoke-ViaCmd ('{0} -Path "{1}" -NoSave -OutputFolder "{2}\"' -f $cmdHead, "$workRoot\m1", "$workRoot\out_q2")
    Assert-Equal '最後の引数の -OutputFolder が \"' $r.Code 0
    $r = Invoke-ViaCmd ('{0} -NoSave -ConfigFile "{1}\" -Days 90' -f $cmdHead, "$workRoot\conf.ini")
    Assert-Equal '-ConfigFile の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '-ConfigFile の値に引用符')) '1/True'
    $r = Invoke-ViaCmd ('{0} -NoSave -Path "{1}" -Exclude "{2}\" -Days 90' -f $cmdHead, "$workRoot\m1", 'CN=x')
    Assert-Equal '-Exclude の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '-Exclude の値に引用符')) '1/True'
    $r = Invoke-ViaCmd ('{0} -NoSave -Store "{1}\" -Days 90' -f $cmdHead, $storeText)
    Assert-Equal '-Store の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '-Store の値に引用符')) '1/True'
    $r = Invoke-ViaCmd ('{0} -NoSave -Target "{1}\" -Days 90' -f $cmdHead, 'web01')
    Assert-Equal '-Target の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '-Target の値に引用符')) '1/True'

    Write-Section 'Ctrl+C で中断'
    # 調べたあとに長く待つコピーで止める（読むだけなので、途中で止めても何も壊れない）
    $slowTool = New-PatchedTool 'slow' '$now = Get-Date' 'Start-Sleep -Seconds 60; $now = Get-Date'
    $stopFile = "$workRoot\stop-during-scan.ps1"
    [IO.File]::WriteAllLines($stopFile, [string[]]@(
        'param($Tool, $Target, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $Target).AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'Start-Sleep -Milliseconds 3000'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
    ), $utf8Bom)
    $stopResult = @{}
    & $Exe -NoProfile -ExecutionPolicy Bypass -File $stopFile -Tool $slowTool -Target "$workRoot\m1" -OutRoot "$workRoot\out_ctrlc" 2>&1 | ForEach-Object {
        $line = "$_"
        $eq = $line.IndexOf('=')
        if ($eq -gt 0) { $stopResult[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
    }
    if ($stopResult['RUNNING'] -eq 'True') {
        Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $stopResult['BEFORE'], $stopResult['AFTER']) '1/0'
        Assert-Equal '中断のメッセージ' $stopResult['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' '3 秒より前に終わったため'
    }

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $cultureProbe = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $cultureProbe)"
    if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
        $cultureOut = "$workRoot\out_culture"
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $manyDir.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $cultureDir = Get-RunDir $cultureOut
        $summaryText = if ($cultureDir) { [IO.File]::ReadAllText((Join-Path $cultureDir 'summary.txt')) } else { '' }
        Assert-Equal 'fi-FI でも件数の区切りは ,' (Test-Contains $summaryText '証明書 1,234 件') 'True'
        Assert-Equal 'fi-FI でも日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($summaryText, '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   ')) 'True'
        Assert-Equal 'fi-FI でも所要時間の小数点は .' ([regex]::IsMatch($summaryText, '所要 \d+\.\d秒')) 'True'
        Assert-Equal 'fi-FI でも表の日時は yyyy-MM-dd HH:mm' ([regex]::IsMatch($summaryText, '  \d{4}-\d{2}-\d{2} \d{2}:\d{2}  ')) 'True'
        $row = if ($cultureDir) { @(Import-Csv -LiteralPath (Join-Path $cultureDir 'certificates.csv') -Encoding UTF8)[0] } else { $null }
        Assert-Equal 'fi-FI でも CSV の日時' $(if ($row) { $row.'有効期限' }) (Format-CsvTime $timeOf['near'])
        $negativeOut = "$workRoot\out_culture2"
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), "$files\attrs.pem".Replace("'", "''"), $negativeOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $negativeDir = Get-RunDir $negativeOut
        $row = if ($negativeDir) { @(Import-Csv -LiteralPath (Join-Path $negativeDir 'certificates.csv') -Encoding UTF8)[0] } else { $null }
        Assert-Equal 'fi-FI でも負の数と 3 桁区切り' $(if ($row) { $row.'残り日数' }) '-1,500'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }
} finally {
    $env:windir = $savedWindir
    # TLS サーバーは自分で終わらせる（読み込んだ一時的な鍵を消すため）
    if ($server) {
        try { [IO.File]::WriteAllText($serverStop, 'stop') } catch { }
        if (-not $server.WaitForExit(10000)) { try { $server.Kill() } catch { } }
    }
    if ($certKey) { $certKey.Dispose() }
    # テスト専用のストアを消す。消す対象は、このテストが作った名前のストアだけ
    foreach ($name in $storeName, $store2Name, $store3Name, "$store3Name-m") {
        try { [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree("Software\Microsoft\SystemCertificates\$name", $false) } catch { Write-Section ('テスト専用のストアを削除できませんでした: {0}（{1}）' -f $name, $_.Exception.Message) }
    }
    # アクセス拒否を戻してから消す
    foreach ($target in $denied) { $null = & icacls.exe $target /remove:d "*$userSid" 2>&1 }
    if ($substDrive) { $null = & subst.exe $substDrive /D 2>&1 }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        # ループするジャンクションは、リンク先をたどらないように先にリンクだけ消す
        try {
            foreach ($junction in $junctions) { if (Test-Path -LiteralPath $junction) { [IO.Directory]::Delete($junction) } }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

'  END'
exit $script:failCount
