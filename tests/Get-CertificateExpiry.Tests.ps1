<#
.SYNOPSIS
    Get-CertificateExpiry の回帰テストです。

.DESCRIPTION
    実行したユーザーの CurrentUser\My に、一時的な自己署名の証明書（期限切れ・期限間近・有効・更新済みなど）を作り、
    Get-CertificateExpiry を実行して期待値と照合します。証明書ファイル（DER・PEM・途中で切れた PEM・秘密鍵・PFX など）、
    ローカルに立てた TLS サーバーへの接続、偽の applicationHost.config を使った IIS の紐付け、設定ファイル、止まる場合も確かめます。
    作った証明書とダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダとファイルにだけアクセス拒否を設定します（終了時に戻します）。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Get-CertificateExpiry を実行する PowerShell（powershell.exe または pwsh）。

.PARAMETER KeepTemp
    ダミー環境と結果を削除せずに残します（調査用）。作った証明書は残しません。

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
$denied   = New-Object 'System.Collections.Generic.List[string]'
$junctions = New-Object 'System.Collections.Generic.List[string]'
$thumbs   = @{}
$rawOf    = @{}
$server   = $null
$savedWindir = $env:windir
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
function Invoke-Tool([string[]]$ToolArgs) {
    $script:runNumber++
    $out  = Join-Path $workRoot ("out_{0}" -f $script:runNumber)
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $dir  = Get-RunDir $out
    $certRows = @()
    $skipRows = @()
    if ($dir) {
        $certRows = @(Import-Csv -LiteralPath (Join-Path $dir 'certificates.csv') -Encoding UTF8)
        $skipRows = @(Import-Csv -LiteralPath (Join-Path $dir 'skipped.csv') -Encoding UTF8)
    }
    [pscustomobject]@{ Code = $code; Log = $log; Dir = $dir; Certs = $certRows; Skipped = $skipRows }
}

function Invoke-NoSave([string[]]$ToolArgs) {
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -NoSave 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

function Get-ThumbText($Row) { $Row.'拇印' -replace '^="', '' -replace '"$', '' }

# 作った証明書の行（名前で探す）
function Get-CertRow($Run, [string]$Name) { @($Run.Certs | Where-Object { (Get-ThumbText $_) -eq $thumbs[$Name] }) }

function Get-Status($Run, [string]$Name) { (@(Get-CertRow $Run $Name) | ForEach-Object { $_.'状態' }) -join '|' }

function Get-PlaceRows($Run, [string]$Place) { @($Run.Certs | Where-Object { $_.'場所' -eq $Place }) }

function Get-SkipReason($Run, [string]$Subject) { (@($Run.Skipped | Where-Object { $_.'対象' -eq $Subject }) | ForEach-Object { $_.'理由' }) -join '|' }

function Deny-Access([string]$TargetPath, [string]$Rights) {
    $null = & icacls.exe $TargetPath /deny "*${userSid}:$Rights" 2>&1
    $denied.Add($TargetPath)
}

function New-TestFile([string]$FilePath, [byte[]]$Bytes) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllBytes('\\?\' + $FilePath, $Bytes)
}

function New-TextFile([string]$FilePath, [string]$Text) { New-TestFile $FilePath ([Text.Encoding]::ASCII.GetBytes($Text)) }

function ConvertTo-Pem([byte[]]$Bytes) {
    "-----BEGIN CERTIFICATE-----`r`n" + [Convert]::ToBase64String($Bytes, [Base64FormattingOptions]::InsertLineBreaks) + "`r`n-----END CERTIFICATE-----`r`n"
}

# 偽の %windir% を作り、applicationHost.config を置く
function New-FakeWindir([string]$Name, [string]$HostConfig, [string]$RedirectConfig) {
    $fake = Join-Path $workRoot "windir_$Name"
    $config = Join-Path $fake 'System32\inetsrv\config'
    [void][IO.Directory]::CreateDirectory($config)
    if ($HostConfig) { [IO.File]::WriteAllText((Join-Path $config 'applicationHost.config'), $HostConfig) }
    if ($RedirectConfig) { [IO.File]::WriteAllText((Join-Path $config 'redirection.config'), $RedirectConfig) }
    $fake
}

function New-HostConfig([string]$SitesXml) {
    '<?xml version="1.0"?><configuration><system.applicationHost><sites>' + $SitesXml + '</sites></system.applicationHost></configuration>'
}

function Invoke-WithWindir([string]$Fake, [string[]]$ToolArgs) {
    $env:windir = $Fake
    try { Invoke-Tool $ToolArgs } finally { $env:windir = $savedWindir }
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

try {
    #region ダミー環境

    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    # New-SelfSignedCertificate は 7 では動かないことがあるので、いつも 5.1 で作る
    $certSpec = @(
        @('expired', "CN=wb-expired-$runId", "wb-expired-$runId.example.jp", -400, -10, ''),
        @('near', "CN=wb-near-$runId", "wb-near-$runId.example.jp", -300, 20, ''),
        @('valid', "CN=wb-valid-$runId", "wb-valid-$runId.example.jp,www.wb-valid-$runId.example.jp", -10, 300, '-test"x'),
        @('renewOld', "CN=wb-renew-$runId", "wb-renew-$runId.example.jp", -400, -5, ''),
        @('renewNew', "CN=wb-renew-$runId", "wb-renew-$runId.example.jp", -5, 360, ''),
        @('sanOther', "CN=wb-renew-$runId", "wb-other-$runId.example.jp", -400, -3, ''),
        @('future', "CN=wb-future-$runId", "wb-future-$runId.example.jp", 10, 400, ''),
        @('comma', "CN=wb-comma-$runId, O=Example, C=JP", "wb-comma-$runId.example.jp", -10, 100, ''),
        @('iisOld', "CN=wb-iis-$runId", "wb-iis-$runId.example.jp", -400, -5, ''),
        @('iisNew', "CN=wb-iis-$runId", "wb-iis-$runId.example.jp", -5, 360, '')
    )
    $maker = New-Object Text.StringBuilder
    [void]$maker.AppendLine('$ErrorActionPreference = "Stop"')
    foreach ($spec in $certSpec) {
        $dnsList = ($spec[2] -split ',' | ForEach-Object { "'$_'" }) -join ','
        $friendly = if ($spec[5]) { " -FriendlyName '{0}'" -f $spec[5] } else { '' }
        [void]$maker.AppendLine(("`$c = New-SelfSignedCertificate -Subject '{0}' -DnsName {1} -NotBefore (Get-Date).AddDays({2}) -NotAfter (Get-Date).AddDays({3}) -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable{4}; '{5}=' + `$c.Thumbprint" -f $spec[1], $dnsList, $spec[3], $spec[4], $friendly, $spec[0]))
    }
    $makerFile = Join-Path $workRoot 'make-certs.ps1'
    [IO.File]::WriteAllText($makerFile, $maker.ToString(), (New-Object Text.UTF8Encoding $true))
    foreach ($line in (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $makerFile 2>&1)) {
        $parts = ([string]$line) -split '='
        if ($parts.Count -eq 2 -and $parts[1] -match '^[0-9A-F]{40}$') { $thumbs[$parts[0]] = $parts[1] }
    }
    if ($thumbs.Count -ne $certSpec.Count) { throw ('テスト用の証明書を作れませんでした（{0} / {1} 件）' -f $thumbs.Count, $certSpec.Count) }

    $userStore = New-Object Security.Cryptography.X509Certificates.X509Store('My', [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $userStore.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    foreach ($key in @($thumbs.Keys)) {
        foreach ($c in $userStore.Certificates) { if ($c.Thumbprint -eq $thumbs[$key]) { $rawOf[$key] = $c.RawData } }
    }
    $pfxBytes = $null
    foreach ($c in $userStore.Certificates) { if ($c.Thumbprint -eq $thumbs['valid']) { $pfxBytes = $c.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pfx) } }
    $userStore.Close()

    #endregion

    Write-Section 'ストア'
    $run = Invoke-Tool @('-Store', 'CurrentUser\My')
    Assert-Equal '終了コード（期限切れあり）' $run.Code 2
    Assert-Equal '保存したファイル' ((Get-ChildItem -LiteralPath $run.Dir | ForEach-Object { $_.Name } | Sort-Object) -join ',') 'certificates.csv,skipped.csv,summary.txt'
    Assert-Equal '期限切れ' (Get-Status $run 'expired') '期限切れ'
    Assert-Equal '期限間近' (Get-Status $run 'near') '期限間近'
    Assert-Equal '有効' (Get-Status $run 'valid') '有効'
    Assert-Equal '有効期間前' (Get-Status $run 'future') '有効期間前'
    Assert-Equal '同じサブジェクトと SAN の新しい証明書があれば更新済み' (Get-Status $run 'renewOld') '更新済み'
    Assert-Equal '新しい証明書は有効' (Get-Status $run 'renewNew') '有効'
    Assert-Equal 'SAN が違うと更新済みにしない' (Get-Status $run 'sanOther') '期限切れ'
    $nearRow = @(Get-CertRow $run 'near')[0]
    Assert-Equal '残り日数' ($nearRow.'残り日数' -in '19', '20') 'True'
    Assert-Equal '種類と場所' ('{0}/{1}' -f $nearRow.'種類', $nearRow.'場所') 'ストア/CurrentUser\My'
    Assert-Equal '有効期限の形' ($nearRow.'有効期限' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
    Assert-Equal '秘密鍵' $nearRow.'秘密鍵' 'あり'
    Assert-Equal '用途は OS の言語によらない' (Test-Contains $nearRow.'用途' 'サーバー認証') 'True'
    $validRow = @(Get-CertRow $run 'valid')[0]
    Assert-Equal '別名(SAN)' $validRow.'別名(SAN)' "wb-valid-$runId.example.jp, www.wb-valid-$runId.example.jp"
    Assert-Equal 'Excel で変わるフレンドリ名は ="..." で守る' $validRow.'フレンドリ名' '="-test""x"'
    Assert-Equal '更新済みの備考に新しい証明書' (@(Get-CertRow $run 'renewOld')[0].'備考') ('新しい証明書: {0}' -f $thumbs['renewNew'])
    $times = [string[]]@($run.Certs | ForEach-Object { $_.'有効期限' })
    $sortedTimes = [string[]]$times.Clone()
    [Array]::Sort($sortedTimes, [StringComparer]::Ordinal)
    Assert-Equal '有効期限の早い順' ($times -join ',') ($sortedTimes -join ',')
    Assert-Equal '表に更新済みを出さない' (Test-Contains $run.Log $thumbs['renewNew']) 'False'
    Assert-Equal '表に期限間近を出す' (Test-Contains $run.Log "CN=wb-near-$runId") 'True'
    Assert-Equal 'CSV の BOM' (([IO.File]::ReadAllBytes((Join-Path $run.Dir 'certificates.csv'))[0..2] -join ',')) '239,187,191'

    $longDays = Invoke-Tool @('-Store', 'CurrentUser\My', '-Days', '400')
    Assert-Equal '-Days で基準が変わる' (Get-Status $longDays 'valid') '期限間近'
    Assert-Equal '新しい証明書も期限間近なら更新済みにしない' (Get-Status $longDays 'renewOld') '期限切れ'

    $ex = Invoke-Tool @('-Store', 'CurrentUser\My', '-Exclude', "CN=wb-near-$runId,*wb-expired-$runId*")
    Assert-Equal '-Exclude で除く' ('{0}/{1}' -f @(Get-CertRow $ex 'near').Count, @(Get-CertRow $ex 'expired').Count) '0/0'
    Assert-Equal '除外の件数' (Test-Contains $ex.Log '除外 2 件') 'True'
    $exComma = Invoke-Tool @('-Store', 'CurrentUser\My', '-Exclude', "CN=wb-comma-$runId, O=Example, C=JP,CN=wb-near-$runId")
    Assert-Equal 'カンマとスペースを含むサブジェクトも除ける' ('{0}/{1}' -f @(Get-CertRow $exComma 'comma').Count, @(Get-CertRow $exComma 'near').Count) '0/0'

    $default = Invoke-NoSave @()
    Assert-Equal '既定のストア' (Test-Contains $default.Log 'LocalMachine\My, LocalMachine\WebHosting, LocalMachine\Remote Desktop, CurrentUser\My') 'True'
    Assert-Equal '既定の実行でも作った証明書を見つける' (Test-Contains $default.Log "CN=wb-expired-$runId") 'True'
    $all = Invoke-Tool @('-AllStores')
    Assert-Equal '-AllStores は Root も調べる' (@($all.Certs | Where-Object { $_.'場所' -like '*\Root' }).Count -gt 0) 'True'
    Assert-Equal '-AllStores は UserDS を調べない' (@($all.Certs | Where-Object { $_.'場所' -like '*\UserDS' }).Count) 0

    Write-Section 'ファイル'
    $files = "$workRoot\files"
    New-TestFile "$files\der.cer" $rawOf['near']
    New-TextFile "$files\pem.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$files\bundle.pem" ((ConvertTo-Pem $rawOf['near']) + (ConvertTo-Pem $rawOf['valid']))
    New-TextFile "$files\old.cer" (ConvertTo-Pem $rawOf['renewOld'])
    New-TextFile "$files\new.cer" (ConvertTo-Pem $rawOf['renewNew'])
    New-TextFile "$files\日本語 フォルダ\[x]\expired.der" 'placeholder'
    New-TestFile "$files\日本語 フォルダ\[x]\expired.der" $rawOf['expired']
    $longDir = "$files\" + (('y' * 60 + '\') * 5)
    New-TestFile ($longDir + 'long.cer') $rawOf['future']
    New-TextFile "$files\truncated.pem" ((ConvertTo-Pem $rawOf['comma']) + "-----BEGIN CERTIFICATE-----`r`nMIIB")
    New-TextFile "$files\privkey.pem" "-----BEGIN PRIVATE KEY-----`r`nAAAA`r`n-----END PRIVATE KEY-----`r`n"
    New-TextFile "$files\dhparam.pem" "-----BEGIN DH PARAMETERS-----`r`nAAAA`r`n-----END DH PARAMETERS-----`r`n"
    New-TestFile "$files\secret.pfx" $pfxBytes
    New-TestFile "$files\nopass.crt" $pfxBytes
    New-TextFile "$files\broken.crt" 'not a certificate'
    New-TextFile "$files\memo.txt" 'not a certificate'
    $stream = [IO.File]::Create("$files\huge.cer")
    $stream.SetLength(11MB)
    $stream.Close()
    $null = New-Item -ItemType Junction -Path "$files\jloop" -Target $files
    $junctions.Add("$files\jloop")

    $fileRun = Invoke-Tool @('-Path', $files)
    Assert-Equal '-Path だけならストアは調べない' (@($fileRun.Certs | Where-Object { $_.'種類' -ne 'ファイル' }).Count) 0
    Assert-Equal 'DER' (@(Get-PlaceRows $fileRun "$files\der.cer")[0].'状態') '期限間近'
    Assert-Equal 'PEM' (@(Get-PlaceRows $fileRun "$files\pem.crt")[0].'状態') '有効'
    Assert-Equal '複数の証明書が入った PEM' (@(Get-PlaceRows $fileRun "$files\bundle.pem").Count) 2
    Assert-Equal '同じフォルダの新しい証明書で更新済み' (@(Get-PlaceRows $fileRun "$files\old.cer")[0].'状態') '更新済み'
    Assert-Equal '日本語・スペース・[ ] を含むパス' (@(Get-PlaceRows $fileRun "$files\日本語 フォルダ\[x]\expired.der")[0].'状態') '期限切れ'
    Assert-Equal '260 文字を超えるパス' (@(Get-PlaceRows $fileRun ($longDir + 'long.cer')).Count) 1
    Assert-Equal '\\?\ が残っていない' ((Get-ChildItem -LiteralPath $fileRun.Dir | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join '').Contains('\\?\') 'False'
    Assert-Equal '途中で切れた PEM は読めた分を調べる' (@(Get-PlaceRows $fileRun "$files\truncated.pem").Count) 1
    Assert-Equal '途中で切れた PEM は「読めない」' (Get-SkipReason $fileRun "$files\truncated.pem") '読めない'
    Assert-Equal '秘密鍵だけの PEM は読まない' (Get-SkipReason $fileRun "$files\privkey.pem") '読まない形式'
    Assert-Equal '証明書以外の PEM は読まない' (Get-SkipReason $fileRun "$files\dhparam.pem") '読まない形式'
    Assert-Equal '.pfx は開かない' (Get-SkipReason $fileRun "$files\secret.pfx") '読まない形式'
    Assert-Equal '中身が PFX の .crt は開かない' ('{0}/{1}' -f (Get-SkipReason $fileRun "$files\nopass.crt"), @(Get-PlaceRows $fileRun "$files\nopass.crt").Count) '読まない形式/0'
    Assert-Equal '証明書でないファイルは「読めない」' (Get-SkipReason $fileRun "$files\broken.crt") '読めない'
    Assert-Equal '10 MB を超えるファイルは読まない' (Get-SkipReason $fileRun "$files\huge.cer") '読めない'
    Assert-Equal '関係のない拡張子は記録しない' (Get-SkipReason $fileRun "$files\memo.txt") ''
    Assert-Equal 'ジャンクションはたどらない' (Get-SkipReason $fileRun "$files\jloop") 'リンク'

    New-TextFile "$workRoot\memo-cert.txt" (ConvertTo-Pem $rawOf['valid'])
    $direct = Invoke-Tool @('-Path', "$workRoot\memo-cert.txt")
    Assert-Equal '直接指定は拡張子によらず読む' ('{0}/{1}' -f $direct.Code, @($direct.Certs).Count) '0/1'
    $overlap = Invoke-Tool @('-Path', "$files,$files\日本語 フォルダ")
    Assert-Equal '指定が重なっても 1 回だけ調べる' (@(Get-PlaceRows $overlap "$files\日本語 フォルダ\[x]\expired.der").Count) 1
    New-TextFile "$workRoot\only\a,b\x.crt" (ConvertTo-Pem $rawOf['valid'])
    Assert-Equal 'カンマを含むフォルダ' (@((Invoke-Tool @('-Path', "$workRoot\only\a,b")).Certs).Count) 1

    Write-Section '読めない場所と、終了コード'
    New-TextFile "$workRoot\mix\ok.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\mix\bad.crt" 'broken'
    Assert-Equal '対応不要で読めないものがあると 1' (Invoke-Tool @('-Path', "$workRoot\mix")).Code 1
    New-TextFile "$workRoot\clean\ok.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\clean\key.pem" "-----BEGIN PRIVATE KEY-----`r`nAAAA`r`n-----END PRIVATE KEY-----`r`n"
    $clean = Invoke-Tool @('-Path', "$workRoot\clean")
    Assert-Equal '「読まない形式」だけなら 0' ('{0}/{1}' -f $clean.Code, (Test-Contains $clean.Log '（対応が必要な証明書はありません）')) '0/True'
    [void][IO.Directory]::CreateDirectory("$workRoot\empty")
    Assert-Equal '証明書がないフォルダは 0' (Invoke-Tool @('-Path', "$workRoot\empty")).Code 0

    New-TextFile "$workRoot\deny\ok.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\deny\secret\x.crt" (ConvertTo-Pem $rawOf['valid'])
    New-TextFile "$workRoot\deny\noread.crt" (ConvertTo-Pem $rawOf['valid'])
    Deny-Access "$workRoot\deny\secret" '(RD)'
    Deny-Access "$workRoot\deny\noread.crt" '(R)'
    if (Test-Path -LiteralPath "$workRoot\deny\secret\x.crt" -ErrorAction SilentlyContinue) {
        Write-Skip 'アクセス拒否' '設定できなかったため'
    } else {
        $deny = Invoke-Tool @('-Path', "$workRoot\deny")
        Assert-Equal '読めないフォルダ' (Get-SkipReason $deny "$workRoot\deny\secret") '権限なし'
        Assert-Equal '読めないファイル' (Get-SkipReason $deny "$workRoot\deny\noread.crt") '権限なし'
        Assert-Equal '読めない場所があっても続ける' ('{0}/{1}' -f $deny.Code, @($deny.Certs).Count) '1/1'
    }

    Write-Section '接続先'
    $portFile = Join-Path $workRoot 'server.port'
    $serverFile = Join-Path $workRoot 'server.ps1'
    $serverCode = @'
param([string]$Thumb, [string]$PortFile)
$store = New-Object Security.Cryptography.X509Certificates.X509Store('My', [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
$store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
$cert = @($store.Certificates | Where-Object { $_.Thumbprint -eq $Thumb })[0]
$listener = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::IPv6Any), 0
$listener.Server.DualMode = $true
$listener.Start()
[IO.File]::WriteAllText($PortFile, [string]$listener.LocalEndpoint.Port)
while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
        $tls = New-Object Net.Security.SslStream($client.GetStream())
        $tls.AuthenticateAsServer($cert)
        $tls.Dispose()
    } catch { }
    $client.Close()
}
'@
    [IO.File]::WriteAllText($serverFile, $serverCode, (New-Object Text.UTF8Encoding $true))
    $server = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$serverFile`"", $thumbs['near'], "`"$portFile`"") -WindowStyle Hidden -PassThru
    $waited = 0
    while (-not (Test-Path -LiteralPath $portFile) -and $waited -lt 100) { Start-Sleep -Milliseconds 200; $waited++ }
    if (-not (Test-Path -LiteralPath $portFile)) {
        Write-Skip '接続先' 'テスト用の TLS サーバーを起動できなかったため'
    } else {
        Start-Sleep -Milliseconds 500
        $port = ([IO.File]::ReadAllText($portFile)).Trim()
        # 閉じているポートを探す（一度開いて閉じる）
        $probe = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback), 0
        $probe.Start()
        $closedPort = $probe.LocalEndpoint.Port
        $probe.Stop()

        $tlsRun = Invoke-Tool @('-Target', "localhost:$port,127.0.0.1:$port,[::1]:$port,127.0.0.1:$closedPort", '-Timeout', '3')
        Assert-Equal '-Target だけならストアは調べない' (@($tlsRun.Certs | Where-Object { $_.'種類' -ne '接続先' }).Count) 0
        $localRow = @(Get-PlaceRows $tlsRun "localhost:$port")
        Assert-Equal '証明書を受け取る' ('{0}/{1}' -f $localRow.Count, $(if ($localRow.Count) { Get-ThumbText $localRow[0] })) ('1/{0}' -f $thumbs['near'])
        Assert-Equal '状態' $localRow[0].'状態' '期限間近'
        Assert-Equal '信頼できない発行者と、名前の不一致を備考に書く' ('{0}/{1}' -f (Test-Contains $localRow[0].'備考' '信頼できない発行者'), (Test-Contains $localRow[0].'備考' '名前が一致しない')) 'True/True'
        Assert-Equal 'IPv4 で接続する' (@(Get-PlaceRows $tlsRun "127.0.0.1:$port").Count) 1
        Assert-Equal 'IPv6 で接続する' (@(Get-PlaceRows $tlsRun "[::1]:$port").Count) 1
        Assert-Equal '接続できないものは記録して続ける' (Get-SkipReason $tlsRun "127.0.0.1:$closedPort") '接続できない'
        Assert-Equal '期限間近があるので 2' $tlsRun.Code 2
        $closedRun = Invoke-Tool @('-Target', "127.0.0.1:$closedPort", '-Timeout', '2')
        Assert-Equal '接続できないだけなら 1' $closedRun.Code 1
        $bothRun = Invoke-Tool @('-Target', "127.0.0.1:$port", '-Store', 'CurrentUser\My')
        Assert-Equal '-Store と一緒なら両方調べる' ('{0}/{1}' -f (@($bothRun.Certs | Where-Object { $_.'種類' -eq '接続先' }).Count), (@(Get-CertRow $bothRun 'expired').Count)) '1/1'
    }

    Write-Section 'IIS の紐付け'
    $iisXml = '<site name="Portal"><bindings><binding protocol="http" bindingInformation="*:80:" /><binding protocol="https" bindingInformation="*:443:portal.example" certificateHash="{0}" /></bindings></site>' -f $thumbs['iisOld'].ToLowerInvariant()
    $ftpXml = '<site name="Files"><ftpServer><security><ssl serverCertHash="{0}" /></security></ftpServer></site>' -f $thumbs['near']
    $unknownXml = '<site name="Unknown"><bindings><binding protocol="https" bindingInformation="192.0.2.77:64999:" /></bindings></site>'
    $centralXml = '<site name="Central"><bindings><binding protocol="https" bindingInformation="192.0.2.77:64999:" sslFlags="3" /></bindings></site>'

    $iis = Invoke-WithWindir (New-FakeWindir 'ok' (New-HostConfig ($iisXml + $ftpXml)) '') @('-Store', 'CurrentUser\My')
    Assert-Equal 'サイトで使っている古い証明書は更新済みにしない' (Get-Status $iis 'iisOld') '期限切れ'
    Assert-Equal 'サイトとバインド' (@(Get-CertRow $iis 'iisOld')[0].'IIS のサイト') 'Portal (*:443:portal.example)'
    Assert-Equal '画面にサイトの名前を出す' (Test-Contains $iis.Log 'IIS: Portal') 'True'
    Assert-Equal 'FTP サイトの証明書' (@(Get-CertRow $iis 'near')[0].'IIS のサイト') 'Files (FTP)'
    Assert-Equal '使っていない古い証明書は更新済み' (Get-Status $iis 'renewOld') '更新済み'
    Assert-Equal '読めなかったと出ない' (Test-Contains $iis.Log 'IIS の設定を読めなかった') 'False'

    $unknown = Invoke-WithWindir (New-FakeWindir 'unknown' (New-HostConfig $unknownXml) '') @('-Store', 'CurrentUser\My')
    Assert-Equal '証明書が分からないバインドがあると更新済みにしない' ('{0}/{1}' -f (Get-Status $unknown 'renewOld'), (Test-Contains $unknown.Log 'IIS の設定を読めなかった')) '期限切れ/True'
    $central = Invoke-WithWindir (New-FakeWindir 'central' (New-HostConfig $centralXml) '') @('-Store', 'CurrentUser\My')
    Assert-Equal '中央証明書ストアのバインドは気にしない' ('{0}/{1}' -f (Get-Status $central 'renewOld'), (Test-Contains $central.Log 'IIS の設定を読めなかった')) '更新済み/False'
    $shared = Invoke-WithWindir (New-FakeWindir 'shared' (New-HostConfig $iisXml) '<configuration><configurationRedirection enabled="true" path="\\fs\iis" /></configuration>') @('-Store', 'CurrentUser\My')
    Assert-Equal '共有構成は読めなかった扱い' ('{0}/{1}' -f (Get-Status $shared 'renewOld'), (Test-Contains $shared.Log 'IIS の設定を読めなかった')) '期限切れ/True'
    $brokenXml = Invoke-WithWindir (New-FakeWindir 'broken' '<configuration>' '') @('-Store', 'CurrentUser\My')
    Assert-Equal '壊れた設定でも止まらない' ('{0}/{1}' -f (Get-Status $brokenXml 'renewOld'), (Test-Contains $brokenXml.Log 'エラーで止まりました')) '期限切れ/False'
    $noFile = Invoke-WithWindir (New-FakeWindir 'nofile' '' '') @('-Store', 'CurrentUser\My')
    Assert-Equal '設定ファイルがなければ IIS なし' ('{0}/{1}' -f (Get-Status $noFile 'renewOld'), (Test-Contains $noFile.Log 'IIS の設定を読めなかった')) '更新済み/False'
    $noIisDir = Join-Path $workRoot 'windir_none'
    [void][IO.Directory]::CreateDirectory($noIisDir)
    $noIis = Invoke-WithWindir $noIisDir @('-Store', 'CurrentUser\My')
    Assert-Equal 'inetsrv がなければ IIS なし' (Test-Contains $noIis.Log 'IIS の設定を読めなかった') 'False'

    Write-Section '設定ファイル'
    $conf = "$workRoot\conf.ini"
    [IO.File]::WriteAllText($conf, ("# コメント`r`nDays = 400`r`nPath = clean`r`nPath = only\a,b`r`nExclude = CN=nobody`r`nNoSave = false`r`n"), (New-Object Text.UTF8Encoding $false))
    $confRun = Invoke-Tool @('-ConfigFile', $conf)
    Assert-Equal '複数の行の Path と、カンマを含む Path' (@($confRun.Certs).Count) 2
    Assert-Equal 'Days を設定ファイルから読む' (@($confRun.Certs)[0].'状態') '期限間近'
    Assert-Equal '設定ファイルの表示' (Test-Contains $confRun.Log "設定ファイル: $conf") 'True'
    Assert-Equal 'コマンドラインを優先する' (Invoke-Tool @('-ConfigFile', $conf, '-Days', '1')).Code 0

    $env:WB_TEST_SUBJECT = "wb-valid-$runId"
    [IO.File]::WriteAllText("$workRoot\envconf.ini", "Path = clean`r`nExclude = CN=%WB_TEST_SUBJECT%`r`n", (New-Object Text.UTF8Encoding $false))
    Assert-Equal 'Exclude の環境変数' (@((Invoke-Tool @('-ConfigFile', "$workRoot\envconf.ini")).Certs).Count) 0
    Remove-Item Env:\WB_TEST_SUBJECT

    [IO.File]::WriteAllText("$workRoot\store.ini", "Store = CurrentUser\My`r`n", (New-Object Text.UTF8Encoding $false))
    Assert-Equal '-AllStores は設定ファイルの Store より優先する' (Test-Contains (Invoke-NoSave @('-ConfigFile', "$workRoot\store.ini", '-AllStores')).Log 'すべてのストア') 'True'
    [IO.File]::WriteAllText("$workRoot\all.ini", "AllStores = true`r`n", (New-Object Text.UTF8Encoding $false))
    Assert-Equal '-Store は設定ファイルの AllStores より優先する' (Test-Contains (Invoke-NoSave @('-ConfigFile', "$workRoot\all.ini", '-Store', 'CurrentUser\My')).Log '対象: CurrentUser\My') 'True'

    $badCases = @(
        @('Exlude = a', '知らない項目です: Exlude'),
        @('Days = 0', 'Days は 1～3650 の数値で指定してください: 0'),
        @('Timeout = 99', 'Timeout は 1～60 の数値で指定してください: 99'),
        @('NoSave = yes', 'NoSave は true か false で指定してください: yes'),
        @('Store = My', 'Store は LocalMachine\My の形で指定してください: My'),
        @('Target = a:0', 'Target は ホスト か ホスト:ポート'),
        @('Days 3', '「項目 = 値」の形になっていません'),
        @("Days = 1`r`nDays = 2", 'Days が 2 回書かれています'),
        @('Path = %NO_SUCH_VARIABLE_X%', '定義されていない環境変数です'),
        @("# 1`r`nStore = CurrentUser\My`r`nStore = LocalMachine\NoSuchStoreX", '（設定ファイルの 3 行目）')
    )
    foreach ($case in $badCases) {
        [IO.File]::WriteAllText("$workRoot\bad.ini", $case[0], (New-Object Text.UTF8Encoding $false))
        $result = Invoke-NoSave @('-ConfigFile', "$workRoot\bad.ini")
        Assert-Equal "書き間違い: $($case[1])" ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $case[1])) '1/True'
    }

    Write-Section '止まるとき'
    $stops = @(
        @('-Store の形', @('-Store', 'My'), '-Store は LocalMachine\My の形で指定してください: My'),
        @('ないストア', @('-Store', 'LocalMachine\NoSuchStoreX'), 'ストアが見つかりません: LocalMachine\NoSuchStoreX'),
        @('-Store と -AllStores', @('-Store', 'CurrentUser\My', '-AllStores'), '-Store と -AllStores は一緒に指定できません'),
        @('-Target のポート', @('-Target', 'web01:99999'), '-Target は ホスト か ホスト:ポート'),
        @('-Target の URL', @('-Target', 'https://www.example.jp/'), '-Target は ホスト か ホスト:ポート'),
        @('-Target のコロンが多い', @('-Target', 'web01:443:x'), '-Target は ホスト か ホスト:ポート'),
        @('見つからない -Path', @('-Path', "$workRoot\no_such"), 'ファイルかフォルダが見つかりません')
    )
    foreach ($stop in $stops) {
        $result = Invoke-NoSave $stop[1]
        Assert-Equal $stop[0] ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $stop[2])) '1/True'
    }
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Store 'LocalMachine\NoSuchStoreX' -OutputFolder "$workRoot\out_nostore" 2>&1
    Assert-Equal '止まったときは結果を残さない' ([bool](Get-RunDir "$workRoot\out_nostore")) 'False'

    Write-Section '起動のしかた'
    # コマンドプロンプトで "C:\data\" のように書くと、\" が引用符のまま届く
    $bat = "$workRoot\quote.cmd"
    [IO.File]::WriteAllText($bat, ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -Path ""{2}\clean\"" -NoSave`r`nexit /b %errorlevel%`r`n" -f $Exe, $tool, $workRoot), (New-Object Text.UTF8Encoding $false))
    $quoteLog = ConvertTo-FlatText (& cmd.exe /c $bat 2>&1 | Out-String)
    Assert-Equal '最後が \" の引数' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $quoteLog '引用符（"）が残っています')) '1/True'

    # 名前に [ ] を含むフォルダから、相対パスで起動する
    New-TextFile "$workRoot\[br]\sub\a.crt" (ConvertTo-Pem $rawOf['valid'])
    $brBat = "$workRoot\bracket.cmd"
    [IO.File]::WriteAllText($brBat, ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\[br]""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -Path sub -NoSave`r`nexit /b %errorlevel%`r`n" -f $workRoot, $Exe, $tool), (New-Object Text.UTF8Encoding $false))
    $brLog = ConvertTo-FlatText (& cmd.exe /c $brBat 2>&1 | Out-String)
    Assert-Equal '[ ] を含むフォルダから起動する' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $brLog '証明書 1 件')) '0/True'

    # 保存先に書き込めないときは、画面表示だけで終わる（保存先がファイル）
    New-TextFile "$workRoot\notdir" 'x'
    $unwritable = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$workRoot\clean" -OutputFolder "$workRoot\notdir" 2>&1 | Out-String)
    Assert-Equal '書き込めない保存先' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $unwritable '保存先に書き込めないため')) '0/True'
} finally {
    $env:windir = $savedWindir
    if ($server) { try { Stop-Process -Id $server.Id -Force -ErrorAction SilentlyContinue } catch { } }
    # 作った証明書と、その秘密鍵を消す。消す対象は、このテストが作った拇印だけ
    if ($thumbs.Count -gt 0) {
        $remover = ($thumbs.Values | ForEach-Object { "Remove-Item -LiteralPath 'Cert:\CurrentUser\My\$_' -DeleteKey -ErrorAction SilentlyContinue" }) -join "`r`n"
        $removerFile = Join-Path $workRoot 'remove-certs.ps1'
        [IO.File]::WriteAllText($removerFile, $remover, (New-Object Text.UTF8Encoding $true))
        $null = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $removerFile 2>&1
    }
    # アクセス拒否を戻してから消す
    foreach ($target in $denied) { $null = & icacls.exe $target /remove:d "*$userSid" 2>&1 }
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

exit $script:failCount
