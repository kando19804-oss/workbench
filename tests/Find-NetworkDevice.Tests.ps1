<#
.SYNOPSIS
    Find-NetworkDevice の回帰テストです。

.DESCRIPTION
    ループバック（127.0.0.x。Windows では 127.0.0.0/8 のすべてが ping に応答する）と、この PC 自身・既定のゲートウェイを
    相手に Find-NetworkDevice を実行し、範囲の読み方、並び順、この PC とゲートウェイの見分け、MAC アドレスの読み取り、
    CSV の形、止まる場合を確かめます。
    この PC のネットワークが 256 件以下のときだけ、範囲を自動で決める動き（そのネットワーク全体への ping）も確かめます。
    結果は一時フォルダに保存し、終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Find-NetworkDevice を実行する PowerShell（powershell.exe または pwsh）。

.PARAMETER KeepTemp
    結果を削除せずに残します（調査用）。

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

$tool     = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Find-NetworkDevice\Find-NetworkDevice.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Find-NetworkDevice_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
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

function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out  = Join-Path $workRoot "out_$Label"
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $dir  = $null
    if (Test-Path -LiteralPath $out) {
        $found = @(Get-ChildItem -LiteralPath $out -Directory -Filter 'NetworkDevice_*')
        if ($found.Count -gt 0) { $dir = $found[0].FullName }
    }
    [pscustomobject]@{ Code = $code; Log = $log; Dir = $dir }
}

function Read-Devices([string]$Dir) {
    if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir 'devices.csv'))) { return @() }
    @(Import-Csv -LiteralPath (Join-Path $Dir 'devices.csv') -Encoding UTF8)
}

function Get-Row($Rows, [string]$Ip) { @($Rows | Where-Object { $_.'IP アドレス' -eq $Ip }) }

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $log = ConvertTo-FlatText (& cmd.exe /d /c $batch 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
[void][IO.Directory]::CreateDirectory($workRoot)

# この PC の、既定のゲートウェイのあるアダプター（テストの期待値を作るため）
$adapter = $null
foreach ($nic in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
    if ($nic.OperationalStatus -ne 'Up' -or $nic.NetworkInterfaceType -eq 'Loopback') { continue }
    $props = $nic.GetIPProperties()
    $gateway = @($props.GatewayAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -ne '0.0.0.0' } | ForEach-Object { $_.Address.ToString() }) | Select-Object -First 1
    $unicast = @($props.UnicastAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' }) | Select-Object -First 1
    if ($gateway -and $unicast) {
        $adapter = [pscustomobject]@{
            Ip = $unicast.Address.ToString(); Mask = $unicast.IPv4Mask.ToString(); Gateway = $gateway
            Mac = (($nic.GetPhysicalAddress().GetAddressBytes() | ForEach-Object { $_.ToString('X2') }) -join '-')
        }
        break
    }
}

try {
    Write-Section 'ループバック（ネットワークの外の範囲として、ping だけで探す）'
    $run = Invoke-Tool 'loop' @('-Range', '127.0.0.9-127.0.0.11')
    Assert-Equal '終了コード' $run.Code 0
    $rows = Read-Devices $run.Dir
    # IP アドレスの数値の順（文字の順だと 127.0.0.10 が 127.0.0.9 より前になる）
    Assert-Equal 'IP アドレスの小さい順' (($rows | ForEach-Object { $_.'IP アドレス' }) -join ',') '127.0.0.9,127.0.0.10,127.0.0.11'
    Assert-Equal 'ping に応答' (($rows | ForEach-Object { $_.'ping' }) -join ',') '応答,応答,応答'
    Assert-Equal 'ネットワークの外なので MAC アドレスは空欄' (($rows | ForEach-Object { $_.'MAC アドレス' }) -join ',') ',,'
    Assert-Equal 'ネットワークの列' (@($rows | Select-Object -First 1) | ForEach-Object { $_.'ネットワーク' }) '127.0.0.9-127.0.0.11'
    Assert-Equal 'ネットワークの外の注意' $run.Log.Contains('※ 127.0.0.9-127.0.0.11 は、この PC のネットワークの外です') 'True'
    Assert-Equal '件数の行' ([regex]::Match($run.Log, '機器 \S+ 台（[^）]*）').Value) '機器 3 台（ping に応答 3・応答なし 0）'
    $columns = (Get-Content -LiteralPath (Join-Path $run.Dir 'devices.csv') -Encoding UTF8 -TotalCount 1)
    Assert-Equal 'CSV の見出し' $columns '"IP アドレス","名前","名前の調べ方","MAC アドレス","ping","応答時間(ms)","種類","ネットワーク"'
    $bytes = [IO.File]::ReadAllBytes((Join-Path $run.Dir 'devices.csv'))
    Assert-Equal 'BOM 付き UTF-8' ('{0:X2}{1:X2}{2:X2}' -f $bytes[0], $bytes[1], $bytes[2]) 'EFBBBF'
    Assert-Equal 'summary.txt' (Test-Path -LiteralPath (Join-Path $run.Dir 'summary.txt')) 'True'

    Write-Section '範囲の書き方'
    $run = Invoke-Tool 'cidr' @('-Range', '127.0.0.8/30')
    Assert-Equal 'CIDR（ネットワークとブロードキャストを除く）' ((@(Read-Devices $run.Dir) | ForEach-Object { $_.'IP アドレス' }) -join ',') '127.0.0.9,127.0.0.10'
    $run = Invoke-Tool 'multi' @('-Range', '127.0.0.20,127.0.0.5-127.0.0.6')
    Assert-Equal 'カンマ区切り・1 つだけの IP' ((@(Read-Devices $run.Dir) | ForEach-Object { $_.'IP アドレス' }) -join ',') '127.0.0.5,127.0.0.6,127.0.0.20'
    # レビューで見つかった問題: PowerShell から -Range a,b と配列で渡すと、空白でつながってエラーになっていた
    $out = Join-Path $workRoot 'out_array'
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -Command ("& '{0}' -Range 127.0.0.30,127.0.0.31 -OutputFolder '{1}'" -f $tool, $out) 2>&1 | Out-String)
    $dir = @(Get-ChildItem -LiteralPath $out -Directory -ErrorAction SilentlyContinue) | Select-Object -First 1
    Assert-Equal 'PowerShell から配列で渡す' ((@(Read-Devices $(if ($dir) { $dir.FullName } else { $null })) | ForEach-Object { $_.'IP アドレス' }) -join ',') '127.0.0.30,127.0.0.31'

    Write-Section 'この PC と既定のゲートウェイ'
    if ($adapter) {
        $run = Invoke-Tool 'self' @('-Range', ('{0},{1}' -f $adapter.Ip, $adapter.Gateway))
        $rows = Read-Devices $run.Dir
        Assert-Equal 'この PC' ('{0}/{1}/{2}' -f (Get-Row $rows $adapter.Ip | ForEach-Object { $_.'種類' }), (Get-Row $rows $adapter.Ip | ForEach-Object { $_.'MAC アドレス' }), (Get-Row $rows $adapter.Ip | ForEach-Object { $_.'ping' })) ('この PC/{0}/' -f $adapter.Mac)
        $gatewayRow = Get-Row $rows $adapter.Gateway
        Assert-Equal 'ゲートウェイ' (($gatewayRow | ForEach-Object { $_.'種類' }) -join ',') 'ゲートウェイ'
        Assert-Equal 'ゲートウェイの MAC アドレス（ARP の表から読む）' ((($gatewayRow | ForEach-Object { $_.'MAC アドレス' }) -join ',') -match '^[0-9A-F]{2}(-[0-9A-F]{2}){5}$') 'True'

        # レビューで見つかった問題: 範囲の一部だけがこの PC のネットワークに入るときに、全体を外として扱っていた
        $ipNumber = 0L; foreach ($o in $adapter.Ip.Split('.')) { $ipNumber = $ipNumber * 256 + [int]$o }
        $maskNumber = 0L; foreach ($o in $adapter.Mask.Split('.')) { $maskNumber = $maskNumber * 256 + [int]$o }
        $network = $ipNumber -band $maskNumber
        $toText = { param([long]$v) '{0}.{1}.{2}.{3}' -f (($v -shr 24) -band 255), (($v -shr 16) -band 255), (($v -shr 8) -band 255), ($v -band 255) }
        $partial = '{0}-{1}' -f (& $toText ($network - 1)), (& $toText ($network + 1))
        $run = Invoke-Tool 'partial' @('-Range', $partial)
        Assert-Equal '一部だけ外の注意' $run.Log.Contains(('※ {0} の一部は、この PC のネットワークの外です' -f $partial)) 'True'

        $size = [long](4294967296 - $maskNumber)
        if ($size -le 256) {
            Write-Section '範囲を自動で決める（この PC のネットワーク全体に ping を送る）'
            $run = Invoke-Tool 'auto' @()
            $rows = Read-Devices $run.Dir
            Assert-Equal '終了コード' $run.Code 0
            Assert-Equal 'この PC を含む' ((Get-Row $rows $adapter.Ip | ForEach-Object { $_.'種類' }) -join ',') 'この PC'
            Assert-Equal 'ゲートウェイを含む' ((Get-Row $rows $adapter.Gateway | ForEach-Object { $_.'種類' }) -join ',') 'ゲートウェイ'
            Assert-Equal '同じ IP アドレスが 2 行ない' (@($rows | Group-Object -Property 'IP アドレス' | Where-Object { $_.Count -gt 1 }).Count) 0
        } else {
            Write-Skip '範囲を自動で決める' ('この PC のネットワークが {0:N0} 件あり、テストで全体に ping を送るには大きいため' -f $size)
        }
    } else {
        Write-Skip 'この PC と既定のゲートウェイ' '既定のゲートウェイのあるネットワーク アダプターがないため'
    }

    Write-Section '止まるとき・保存できないとき'
    $cases = @(
        @('書き方の誤り', @('-Range', 'abc'), '-Range の書き方が正しくありません'),
        @('始まりが終わりより後', @('-Range', '127.0.0.9-127.0.0.1'), '-Range の書き方が正しくありません'),
        @('範囲が大きすぎる', @('-Range', '10.0.0.0/19'), '-Range の範囲が大きすぎます'),
        @('-Timeout の範囲外', @('-Timeout', '50'), 'Timeout')
    )
    foreach ($case in $cases) {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @($case[1]) -NoSave 2>&1 | Out-String)
        Assert-Equal $case[0] ('{0}/{1}' -f $LASTEXITCODE, $log.Contains($case[2])) '1/True'
    }
    $noSaveOut = Join-Path $workRoot 'out_nosave'
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Range '127.0.0.1' -NoSave -OutputFolder $noSaveOut 2>&1 | Out-String)
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}/{2}' -f $LASTEXITCODE, (Test-Path -LiteralPath $noSaveOut), $log.Contains('127.0.0.1')) '0/False/True'
    $usedLetters = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    $freeLetter  = @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $usedLetters -notcontains $_ } | Select-Object -First 1
    if ($freeLetter) {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Range '127.0.0.1' -OutputFolder "${freeLetter}:\nope" 2>&1 | Out-String)
        Assert-Equal '書き込めない保存先' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('保存先に書き込めないため')) '0/True'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }

    Write-Section 'コマンドプロンプトから起動'
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -NoSave -Range "127.0.0.1-127.0.0.2,127.0.0.4"' -f $Exe, $tool)
    Assert-Equal 'カンマ区切りの -Range' ('{0}/{1}' -f $run.Code, ([regex]::Match($run.Log, '機器 \S+ 台').Value)) '0/機器 3 台'
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -Range 127.0.0.1 -OutputFolder "{2}\" -NoSave' -f $Exe, $tool, $workRoot)
    Assert-Equal '"パス\" の後ろに引数: ヒント' ('{0}/{1}' -f $run.Code, ([regex]::IsMatch($run.Log, '最後に \\ を付けないでください'))) '1/True'
} finally {
    if ($KeepTemp) {
        Write-Section ('結果を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try { [IO.Directory]::Delete($workRoot, $true) } catch { Write-Section ('結果を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message) }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

exit $script:failCount
