<#
.SYNOPSIS
    同じネットワークにある機器を探して、IP アドレス・名前・MAC アドレスを一覧にします。

.DESCRIPTION
    範囲内のすべての IP アドレスに ping を送り、Windows の ARP の表から MAC アドレスを読みます。
    見つかった機器の名前を DNS の逆引きと NetBIOS で調べ、画面に表示して devices.csv（BOM 付き UTF-8）に保存します。

    範囲を指定しないときは、この PC のネットワーク（既定のゲートウェイのあるネットワーク）を探します。
    機器の設定を変えることはありません。自分が管理するネットワークで使ってください。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Find-NetworkDevice.ps1

.PARAMETER Range
    探す IP アドレスの範囲。"192.168.1.0/24"（CIDR）、"192.168.1.10-192.168.1.50"（始まりと終わり）、"192.168.1.1"（1 つだけ）で書きます。
    カンマで区切って複数書けます。合わせて 4,096 件までです。省略すると、この PC のネットワークを探します。

.PARAMETER Timeout
    ping の待ち時間（ミリ秒、100～10000）。既定値は 500 です。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Find-NetworkDevice.ps1

    この PC がつながっているネットワークの機器を探します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Find-NetworkDevice.ps1 -Range "192.168.1.100-192.168.1.150"

    指定した範囲の機器を探します。

.NOTES
    Find-NetworkDevice 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [string[]]$Range = @(),

    [ValidateRange(100, 10000)]
    [int]$Timeout = 500,

    [string]$OutputFolder,

    [switch]$NoSave
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion  = '1.0.0'
$MaxAuto        = 1024    # 自動で決めた範囲の上限（1 つのネットワークあたり）
$MaxManual      = 4096    # -Range で指定した範囲の上限（合計）
$Parallel       = 64      # 同時に送る ping の数
$NameWaitMs     = 3000    # 名前の問い合わせ（DNS の逆引き・NetBIOS）を待つ時間
$DeviceColumns  = 'IP アドレス', '名前', '名前の調べ方', 'MAC アドレス', 'ping', '応答時間(ms)', '種類', 'ネットワーク'
$Invariant      = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom        = New-Object Text.UTF8Encoding $true
$OnWindows      = [IO.Path]::DirectorySeparatorChar -eq '\'   # 開発用に Mac の pwsh でも動かせるようにするため
$machineName    = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

# .NET Framework の名前解決はスレッドプールの上で 1 件ずつ止まって待つため、引けない名前が多いと
# スレッドが足りなくなり、後ろの引ける名前まで待ち時間を超えてしまう。同時に問い合わせる数だけ、スレッドを先に用意しておく
try {
    $minWorker = 0; $minIo = 0
    [Threading.ThreadPool]::GetMinThreads([ref]$minWorker, [ref]$minIo)
    if ($minWorker -lt $Parallel * 2) { [void][Threading.ThreadPool]::SetMinThreads($Parallel * 2, $minIo) }
} catch { }

#region 補助関数

# コマンドプロンプトから渡された引数に残る引用符を外す
# （"C:\" が C:" として届く問題や、cmd では ' がそのまま届く問題への対策）
function Remove-Quote([string]$Text) {
    $t = $Text.Trim().Trim('"')
    if ($t.Length -ge 2 -and $t.StartsWith("'") -and $t.EndsWith("'")) { $t = $t.Substring(1, $t.Length - 2) }
    return $t
}

# 日付や数値の区切り文字が、OS の地域設定によって変わらないようにする
function Format-Invariant([string]$Format) {
    return [string]::Format($Invariant, $Format, [object[]]$args)
}

function Format-Duration([TimeSpan]$Span) {
    if ($Span.TotalSeconds -lt 60) { return [string]::Format($Invariant, '{0:N1}秒', $Span.TotalSeconds) }
    if ($Span.TotalHours -lt 1) { return '{0}分{1}秒' -f $Span.Minutes, $Span.Seconds }
    return '{0}時間{1}分{2}秒' -f [math]::Floor($Span.TotalHours), $Span.Minutes, $Span.Seconds
}

# .NET のメソッドで起きた例外は MethodInvocationException に包まれて届くため、元の例外を取り出す
function Get-InnerException([Exception]$Exception) {
    while ($Exception -is [Management.Automation.MethodInvocationException] -and $Exception.InnerException) {
        $Exception = $Exception.InnerException
    }
    return $Exception
}

# .NET のメッセージは途中や末尾に改行が入ることがあるため、1 行にまとめる
function Get-ErrorText([Exception]$Exception) {
    return ($Exception.Message -replace '\s*[\r\n]+\s*', ' ').Trim()
}

# CSV の 1 行
function ConvertTo-CsvLine([object[]]$Fields) {
    $parts = New-Object 'string[]' $Fields.Count
    for ($i = 0; $i -lt $Fields.Count; $i++) { $parts[$i] = ([string]$Fields[$i]).Replace('"', '""') }
    return '"' + ($parts -join '","') + '"'
}

# 5.1 と 7 のどちらでも BOM 付き UTF-8 で保存する（Export-Csv の -Encoding UTF8 は 7 だと BOM なしになる）
function Save-Csv([string]$FilePath, [string[]]$Columns, $Rows) {
    $writer = New-Object IO.StreamWriter($FilePath, $false, $Utf8Bom)
    try {
        $writer.WriteLine((ConvertTo-CsvLine $Columns))
        foreach ($row in $Rows) { $writer.WriteLine((ConvertTo-CsvLine $row)) }
    } finally {
        $writer.Dispose()
    }
}

# Excel は、= + - @ で始まる値を数式に、数字と区切りの記号だけの値を数値や日付に変えてしまうため、その形の値は ="..." で囲む
$excelRisk = New-Object Text.RegularExpressions.Regex ('^(?:[=+\-@]|[\d\s.,/:+\-]+$)')
function Protect-ExcelText([string]$Text) {
    if ($Text -eq '' -or -not $excelRisk.IsMatch($Text)) { return $Text }
    return '="' + $Text.Replace('"', '""') + '"'
}

# 画面の表で列をそろえるため、全角文字を 2 文字分の幅として数える（半角カタカナは 1 文字分）
function Get-DisplayWidth([string]$Text) {
    $width = 0
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int]$ch
        if ($code -lt 0x100 -or ($code -ge 0xFF61 -and $code -le 0xFF9F)) { $width += 1 }
        elseif ($code -ge 0xDC00 -and $code -le 0xDFFF) { }   # サロゲートペアの後半は、前半で 2 文字分を数えた
        else { $width += 2 }
    }
    return $width
}

function Format-Column([string]$Text, [int]$Width, [bool]$AlignRight) {
    $pad = ' ' * [math]::Max(0, $Width - (Get-DisplayWidth $Text))
    if ($AlignRight) { return $pad + $Text }
    return $Text + $pad
}

function Stop-WithError([string]$Message) {
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value) {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -OutputFolder "D:\調査"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
}

# .NET の IsPathRooted は | などの文字で例外を出すため、先頭の形で判定する
function Test-RootedPath([string]$Text) {
    return $Text -match '^([A-Za-z]:|[\\/])'
}

# 5.1 は、名前に [ や ] を含むフォルダで起動すると、その名前をワイルドカードとして PowerShell の現在の場所を決める
# （プロセスの現在のフォルダは正しい）。移る先は、一致するフォルダがなければ $PSHOME、1 つだけ一致すればそのフォルダ
# （[a] なら隣の a）、パターンとして正しくない（a[b）か 2 つ以上一致する（[ab]）とドライブのルートになる。
# そのときだけ、相対パスをプロセスの現在のフォルダを基準に解決する。対話の PowerShell で cd したあとは
# プロセスの現在のフォルダが古いままなので、「違っていたら常にプロセスの現在のフォルダを使う」とはしない
$startDir   = $null
$processDir = [Environment]::CurrentDirectory.TrimEnd('\')
$location   = (Get-Location).ProviderPath.TrimEnd('\')
if ($processDir -match '[\[\]]' -and $location -ne $processDir) {
    $misplaced = $location -eq $PSHOME.TrimEnd('\') -or $location -match '^[A-Za-z]:$'
    if (-not $misplaced) {
        try { $misplaced = [Management.Automation.WildcardPattern]::new($processDir, [Management.Automation.WildcardOptions]::IgnoreCase).IsMatch($location) } catch { }
    }
    if ($misplaced) { $startDir = $processDir }
}

function Resolve-StartPath([string]$Text) {
    if ($startDir -and $Text -and -not (Test-RootedPath $Text)) { return $startDir + '\' + $Text }
    return $Text
}

# IPv4 アドレスと数値（0～4294967295）の変換。符号の問題を避けるため long で扱う
function ConvertTo-IpNumber([string]$Text) {
    if ($Text -notmatch '^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$') { return $null }
    $value = 0L
    for ($i = 1; $i -le 4; $i++) {
        $octet = [int]$Matches[$i]
        if ($octet -gt 255) { return $null }
        $value = $value * 256 + $octet
    }
    return $value
}

function ConvertTo-IpText([long]$Value) {
    return '{0}.{1}.{2}.{3}' -f (($Value -shr 24) -band 255), (($Value -shr 16) -band 255), (($Value -shr 8) -band 255), ($Value -band 255)
}

# サブネット マスクのビット数（255.255.255.0 → 24）
function Get-PrefixLength([long]$Mask) {
    $bits = 0
    for ($i = 31; $i -ge 0; $i--) { if (($Mask -band ([long]1 -shl $i)) -ne 0) { $bits++ } else { break } }
    return $bits
}

# プレフィックスから、範囲（ネットワーク アドレスとブロードキャスト アドレスを除く）を作る
function New-CidrRange([long]$Address, [int]$Prefix) {
    $size    = [long][math]::Pow(2, 32 - $Prefix)
    $network = [long]([math]::Floor($Address / $size) * $size)
    $first   = $network
    $last    = $network + $size - 1
    # /31 と /32 は、ネットワーク アドレスとブロードキャスト アドレスがないので、すべてを使う
    if ($Prefix -le 30) { $first++; $last-- }
    return @{ Label = '{0}/{1}' -f (ConvertTo-IpText $network), $Prefix; First = $first; Last = $last; Network = $network; Size = $size }
}

# MAC アドレスを、Windows の arp -a と同じ 00-1A-2B-3C-4D-5E の形にする
function Format-Mac([string]$Text) {
    $parts = @($Text -split '[-:]' | ForEach-Object { $_.PadLeft(2, '0').ToUpperInvariant() })
    return $parts -join '-'
}

# この PC のネットワーク アダプター（有効で、IPv4 のアドレスがあるもの）
function Get-LocalAdapter {
    $list = New-Object 'System.Collections.Generic.List[object]'
    foreach ($nic in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
        if ($nic.OperationalStatus -ne 'Up') { continue }
        if ($nic.NetworkInterfaceType -eq 'Loopback' -or $nic.NetworkInterfaceType -eq 'Tunnel') { continue }
        $props = $null
        try { $props = $nic.GetIPProperties() } catch { continue }
        $gateways = @(foreach ($g in $props.GatewayAddresses) { if ($g.Address.AddressFamily -eq 'InterNetwork' -and $g.Address.ToString() -ne '0.0.0.0') { $g.Address.ToString() } })
        $macBytes = $nic.GetPhysicalAddress().GetAddressBytes()
        $mac = if ($macBytes.Length -eq 6) { ($macBytes | ForEach-Object { $_.ToString('X2') }) -join '-' } else { '' }
        foreach ($unicast in $props.UnicastAddresses) {
            if ($unicast.Address.AddressFamily -ne 'InterNetwork') { continue }
            $address = ConvertTo-IpNumber $unicast.Address.ToString()
            $mask = $null
            try { $mask = ConvertTo-IpNumber $unicast.IPv4Mask.ToString() } catch { }
            if ($null -eq $address -or $null -eq $mask -or $mask -eq 0) { continue }
            $list.Add([pscustomobject]@{ Name = $nic.Name; Address = $address; Ip = $unicast.Address.ToString(); Prefix = (Get-PrefixLength $mask); Mac = $mac; Gateways = $gateways })
        }
    }
    return , $list
}

# この PC のネットワーク（アダプターの IP アドレスとサブネット マスクで決まる範囲）に入る IP アドレスか
function Test-LocalAddress([long]$Value) {
    foreach ($adapter in $adapters) {
        $local = New-CidrRange $adapter.Address $adapter.Prefix
        if ($Value -ge $local.Network -and $Value -le $local.Network + $local.Size - 1) { return $true }
    }
    return $false
}

#endregion

#region 引数の確認

$startTime = Get-Date
$adapters  = Get-LocalAdapter
$networks  = New-Object 'System.Collections.Generic.List[object]'

if ($Range.Count -gt 0) {
    # -File で起動すると配列を渡せないため、カンマ区切りの 1 つの文字列も受け付ける
    $rangeParts = @(foreach ($item in $Range) { (Remove-Quote $item) -split ',' })
    foreach ($part in $rangeParts) {
        $item = $part.Trim().Trim("'")
        if (-not $item) { continue }
        if ($item -match '^([\d.]+)/(\d{1,2})$') {
            $address = ConvertTo-IpNumber $Matches[1]
            $prefix  = [int]$Matches[2]
            if ($null -eq $address -or $prefix -lt 1 -or $prefix -gt 32) { Stop-WithError ('-Range の書き方が正しくありません: {0}（"192.168.1.0/24"、"192.168.1.10-192.168.1.50"、"192.168.1.1" のどれかの形で書きます）' -f $item) }
            $net = New-CidrRange $address $prefix
        } elseif ($item -match '^([\d.]+)\s*-\s*([\d.]+)$') {
            $first = ConvertTo-IpNumber $Matches[1]
            $last  = ConvertTo-IpNumber $Matches[2]
            if ($null -eq $first -or $null -eq $last -or $first -gt $last) { Stop-WithError ('-Range の書き方が正しくありません: {0}（"192.168.1.0/24"、"192.168.1.10-192.168.1.50"、"192.168.1.1" のどれかの形で書きます）' -f $item) }
            $net = @{ Label = '{0}-{1}' -f (ConvertTo-IpText $first), (ConvertTo-IpText $last); First = $first; Last = $last }
        } elseif ($item -match '^[\d.]+$' -and $null -ne (ConvertTo-IpNumber $item)) {
            # 1 つだけの IP アドレス
            $single = ConvertTo-IpNumber $item
            $net = @{ Label = $item; First = $single; Last = $single }
        } else {
            Stop-WithError ('-Range の書き方が正しくありません: {0}（"192.168.1.0/24"、"192.168.1.10-192.168.1.50"、"192.168.1.1" のどれかの形で書きます）' -f $item)
        }
        # この PC のネットワークに、全部が入るか、一部が入るか（入る部分は ARP で MAC アドレスが分かる）
        $net.Adapter = $null
        $net.Coverage = 'outside'
        foreach ($adapter in $adapters) {
            $local = New-CidrRange $adapter.Address $adapter.Prefix
            $localLast = $local.Network + $local.Size - 1
            if ($net.First -ge $local.Network -and $net.Last -le $localLast) { $net.Adapter = $adapter; $net.Coverage = 'inside'; break }
            if ($net.First -le $localLast -and $net.Last -ge $local.Network) { $net.Adapter = $adapter; $net.Coverage = 'partial' }
        }
        $networks.Add($net)
    }
    if ($networks.Count -eq 0) { Stop-WithError '-Range に範囲を書いてください（例: "192.168.1.0/24"）' }
    $total = 0L
    foreach ($net in $networks) { $total += $net.Last - $net.First + 1 }
    if ($total -gt $MaxManual) {
        Stop-WithError (Format-Invariant '-Range の範囲が大きすぎます（{0:N0} 件）。合わせて {1:N0} 件までにして、分けて実行してください' $total $MaxManual)
    }
} else {
    # 既定のゲートウェイのあるアダプターのネットワークを探す
    foreach ($adapter in $adapters) {
        if ($adapter.Gateways.Count -eq 0) { continue }
        $net = New-CidrRange $adapter.Address $adapter.Prefix
        $count = $net.Last - $net.First + 1
        if ($count -gt $MaxAuto) {
            Stop-WithError ((Format-Invariant 'この PC のネットワーク {0}（{1}）は大きすぎるため、探しません（{2:N0} 件。{3:N0} 件までです）。' $net.Label $adapter.Name $count $MaxAuto) + "`r`n" +
                '-Range で、探す範囲を狭めて指定してください（例: -Range "' + (ConvertTo-IpText ($adapter.Address - ($adapter.Address % 256))) + '/24"）')
        }
        if (@($networks | Where-Object { $_.Label -eq $net.Label }).Count -gt 0) { continue }
        $net.Adapter = $adapter
        $net.Coverage = 'inside'
        $networks.Add($net)
    }
    if ($networks.Count -eq 0) {
        Stop-WithError ('探すネットワークを決められません（既定のゲートウェイが設定されたネットワーク アダプターがありません）。' + "`r`n" + '-Range で範囲を指定してください（例: -Range "192.168.1.0/24"）')
    }
}

$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    $outText = Resolve-StartPath $outText
}

#endregion

#region 保存先の準備

# 探し終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir      = $null
$saveWarning = $null
if (-not $NoSave) {
    $baseDir = if ($outText) { $outText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
        $runName = Format-Invariant 'NetworkDevice_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while (Test-Path -LiteralPath $runDir) { $suffix++; $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix) }
        [void][IO.Directory]::CreateDirectory($runDir)
    } catch {
        $saveWarning = '保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, (Get-ErrorText (Get-InnerException $_.Exception))
        $runDir = $null
        Write-Warning $saveWarning
    }
}

#endregion

# 探している途中や保存の途中で Ctrl+C されたときや、保存に失敗したときは、書きかけの結果を残さない
$finished   = $false
$unexpected = $false
$saveError  = $null
$sw         = [Diagnostics.Stopwatch]::StartNew()
$lines      = New-Object 'System.Collections.Generic.List[string]'
$devices    = @{}   # IP の数値 → 機器
try {
    #region ping

    # この PC 自身のアドレスには ping を送らない（アダプターの設定から書く）
    $selfIps = @{}
    foreach ($adapter in $adapters) { $selfIps[$adapter.Address] = $adapter }
    $targets = New-Object 'System.Collections.Generic.List[object]'
    $seen = @{}
    foreach ($net in $networks) {
        for ($value = $net.First; $value -le $net.Last; $value++) {
            if ($seen.ContainsKey($value)) { continue }
            $seen[$value] = $true
            if ($selfIps.ContainsKey($value)) {
                $adapter = $selfIps[$value]
                $devices[$value] = [pscustomobject]@{ Ip = $value; Name = $machineName; NameBy = ''; Mac = $adapter.Mac; Ping = ''; Ms = $null; Kind = 'この PC'; Network = $net.Label }
                continue
            }
            $targets.Add(@{ Value = $value; Network = $net })
        }
    }

    $done = 0
    for ($offset = 0; $offset -lt $targets.Count; $offset += $Parallel) {
        Write-Progress -Activity 'ネットワークの機器一覧' -Status (Format-Invariant 'ping を送っています: {0:N0} / {1:N0} 件' $done $targets.Count) -PercentComplete ([int](100 * $done / [math]::Max(1, $targets.Count)))
        $batch = New-Object 'System.Collections.Generic.List[object]'
        $tasks = New-Object 'System.Collections.Generic.List[System.Threading.Tasks.Task]'
        for ($i = $offset; $i -lt [math]::Min($offset + $Parallel, $targets.Count); $i++) {
            $pinger = New-Object Net.NetworkInformation.Ping
            $job = @{ Target = $targets[$i]; Pinger = $pinger; Task = $null }
            try {
                $job.Task = $pinger.SendPingAsync((ConvertTo-IpText $targets[$i].Value), $Timeout)
                $tasks.Add($job.Task)
            } catch { }
            $batch.Add($job)
        }
        # 待ち時間を過ぎても終わらないものは、応答なしとみなす
        try { [void][Threading.Tasks.Task]::WaitAll($tasks.ToArray(), $Timeout + 1000) } catch { }
        foreach ($job in $batch) {
            $task = $job.Task
            if ($task -and $task.Status -eq 'RanToCompletion' -and "$($task.Result.Status)" -eq 'Success') {
                $value = $job.Target.Value
                $devices[$value] = [pscustomobject]@{ Ip = $value; Name = ''; NameBy = ''; Mac = ''; Ping = '応答'; Ms = [long]$task.Result.RoundtripTime; Kind = ''; Network = $job.Target.Network.Label }
            }
            try { $job.Pinger.Dispose() } catch { }
        }
        $done += $batch.Count
    }

    #endregion

    #region ARP

    # ping を送ると Windows が ARP で MAC アドレスを調べるので、少し待ってから ARP の表を読む。
    # 表示の言語（見出しは日本語のことがある）に頼らないよう、IP アドレスと MAC アドレスの形の行だけを読む
    Write-Progress -Activity 'ネットワークの機器一覧' -Status 'ARP の表を読んでいます'
    Start-Sleep -Milliseconds 500
    $arpLines = @()
    try { $arpLines = @(& arp -a 2>$null) } catch { }
    $arpTable = @{}
    foreach ($line in $arpLines) {
        $ipText = $null; $macText = $null
        if ($line -match '^\s*(\d{1,3}(?:\.\d{1,3}){3})\s+([0-9a-fA-F]{2}(?:-[0-9a-fA-F]{2}){5})\s') {
            $ipText = $Matches[1]; $macText = $Matches[2]
        } elseif (-not $OnWindows -and $line -match '\((\d{1,3}(?:\.\d{1,3}){3})\) at ([0-9a-fA-F]{1,2}(?::[0-9a-fA-F]{1,2}){5})') {
            $ipText = $Matches[1]; $macText = $Matches[2]
        }
        if (-not $ipText) { continue }
        $mac = Format-Mac $macText
        # ブロードキャスト・マルチキャスト・未解決の行は機器ではない
        if ($mac -eq 'FF-FF-FF-FF-FF-FF' -or $mac -eq '00-00-00-00-00-00' -or $mac.StartsWith('01-00-5E')) { continue }
        $value = ConvertTo-IpNumber $ipText
        if ($null -ne $value) { $arpTable[$value] = $mac }
    }
    foreach ($net in $networks) {
        foreach ($value in @($arpTable.Keys)) {
            if ($value -lt $net.First -or $value -gt $net.Last) { continue }
            # この PC のネットワークの外のアドレスは、ARP の表に載っていても使わない（ルーターの MAC アドレスになっているため）
            if (-not (Test-LocalAddress $value)) { continue }
            if ($devices.ContainsKey($value)) {
                if (-not $devices[$value].Mac) { $devices[$value].Mac = $arpTable[$value] }
            } else {
                $devices[$value] = [pscustomobject]@{ Ip = $value; Name = ''; NameBy = ''; Mac = $arpTable[$value]; Ping = '応答なし'; Ms = $null; Kind = ''; Network = $net.Label }
            }
        }
    }
    # 既定のゲートウェイ
    foreach ($adapter in $adapters) {
        foreach ($gateway in $adapter.Gateways) {
            $value = ConvertTo-IpNumber $gateway
            if ($null -ne $value -and $devices.ContainsKey($value) -and -not $devices[$value].Kind) { $devices[$value].Kind = 'ゲートウェイ' }
        }
    }

    #endregion

    #region 名前

    Write-Progress -Activity 'ネットワークの機器一覧' -Status (Format-Invariant '名前を調べています: {0:N0} 台' $devices.Count)
    # DNS の逆引きを、ping と同じ数ずつまとめて問い合わせる（一度に始めると、待ち時間を数え始めた後に始まる問い合わせが出るため）
    $dnsTargets = @($devices.Values | Where-Object { $_.Kind -ne 'この PC' })
    for ($offset = 0; $offset -lt $dnsTargets.Count; $offset += $Parallel) {
        $dnsJobs = New-Object 'System.Collections.Generic.List[object]'
        $dnsTasks = New-Object 'System.Collections.Generic.List[System.Threading.Tasks.Task]'
        for ($i = $offset; $i -lt [math]::Min($offset + $Parallel, $dnsTargets.Count); $i++) {
            try {
                $task = [Net.Dns]::GetHostEntryAsync([Net.IPAddress]::Parse((ConvertTo-IpText $dnsTargets[$i].Ip)))
                $dnsJobs.Add(@{ Device = $dnsTargets[$i]; Task = $task })
                $dnsTasks.Add($task)
            } catch { }
        }
        if ($dnsTasks.Count -gt 0) { try { [void][Threading.Tasks.Task]::WaitAll($dnsTasks.ToArray(), $NameWaitMs) } catch { } }
        foreach ($job in $dnsJobs) {
            $task = $job.Task
            if ($task.Status -ne 'RanToCompletion') { continue }
            $hostName = [string]$task.Result.HostName
            # 名前が分からないときは、IP アドレスがそのまま返ることがある
            if ($hostName -and $hostName -ne (ConvertTo-IpText $job.Device.Ip)) { $job.Device.Name = $hostName; $job.Device.NameBy = 'DNS' }
        }
    }

    # DNS で分からなかった機器は、NetBIOS の名前（ノード状態）を UDP 137 番に問い合わせる。1 つのソケットからまとめて送り、まとめて受け取る
    $netbiosTargets = @($devices.Values | Where-Object { -not $_.Name -and $_.Kind -ne 'この PC' })
    if ($netbiosTargets.Count -gt 0) {
        # 問い合わせの名前は「*」を NetBIOS の形（1 文字を 2 文字に分ける）にしたもの
        $query = New-Object 'System.Collections.Generic.List[byte]'
        $query.AddRange([byte[]](0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x20))
        $query.AddRange([Text.Encoding]::ASCII.GetBytes('CK' + ('A' * 30)))
        $query.AddRange([byte[]](0x00, 0x00, 0x21, 0x00, 0x01))
        $packet = $query.ToArray()
        $byIp = @{}
        foreach ($device in $netbiosTargets) { $byIp[(ConvertTo-IpText $device.Ip)] = $device }
        $oem = [Text.Encoding]::ASCII
        try { $oem = [Text.Encoding]::GetEncoding(932) } catch { }
        $udp = $null
        try {
            $udp = New-Object Net.Sockets.UdpClient 0
            # NetBIOS に答えない機器から ICMP のポート到達不能が返ると、Windows は次の Receive を例外にする。それを止める
            try { [void]$udp.Client.IOControl(-1744830452, [byte[]](0, 0, 0, 0), $null) } catch { }
            $index = 0
            foreach ($ipText in $byIp.Keys) {
                $index++
                $packet[0] = [byte](($index -shr 8) -band 255)
                $packet[1] = [byte]($index -band 255)
                try { [void]$udp.Send($packet, $packet.Length, $ipText, 137) } catch { }
            }
            $deadline = $sw.ElapsedMilliseconds + $NameWaitMs
            $remote = New-Object Net.IPEndPoint ([Net.IPAddress]::Any), 0
            while ($sw.ElapsedMilliseconds -lt $deadline) {
                if ($udp.Available -le 0) { Start-Sleep -Milliseconds 20; continue }
                try { $response = $udp.Receive([ref]$remote) } catch { continue }
                $from = $remote.Address.ToString()
                if (-not $byIp.ContainsKey($from)) { continue }
                # 見出し 12 バイト + 名前 34 バイト + 種類・クラス・TTL・長さ 10 バイトの後に、名前の数と名前（1 つ 18 バイト）が並ぶ
                if ($response.Length -lt 57) { continue }
                $count = [int]$response[56]
                for ($n = 0; $n -lt $count; $n++) {
                    $at = 57 + 18 * $n
                    if ($at + 18 -gt $response.Length) { break }
                    $suffix = $response[$at + 15]
                    $isGroup = ($response[$at + 16] -band 0x80) -ne 0
                    # 末尾が 0x00 で、グループでない名前が、その機器の名前
                    if ($suffix -eq 0 -and -not $isGroup) {
                        $name = $oem.GetString($response, $at, 15).Trim()
                        if ($name) { $byIp[$from].Name = $name; $byIp[$from].NameBy = 'NetBIOS' }
                        break
                    }
                }
                $byIp.Remove($from)
                if ($byIp.Count -eq 0) { break }
            }
        } catch {
        } finally {
            if ($udp) { $udp.Close() }
        }
    }
    $elapsed = $sw.Elapsed
    Write-Progress -Activity 'ネットワークの機器一覧' -Completed

    #endregion

    #region 画面と summary.txt の内容

    $sorted = @($devices.Values | Sort-Object -Property Ip)
    $lines.Add((Format-Invariant 'ネットワークの機器一覧 (Find-NetworkDevice {0})' $ScriptVersion))
    $outside = @()
    foreach ($net in $networks) {
        $count = $net.Last - $net.First + 1
        $where = if ($net.Adapter) { '   ' + $net.Adapter.Name } else { '' }
        $lines.Add((Format-Invariant '対象: {0}（{1} ～ {2}・{3:N0} 件）{4}' $net.Label (ConvertTo-IpText $net.First) (ConvertTo-IpText $net.Last) $count $where))
        if ($net.Coverage -ne 'inside') { $outside += $net }
    }
    $selfText = @($adapters | Where-Object { $_.Gateways.Count -gt 0 } | ForEach-Object { $_.Ip }) -join ', '
    if (-not $selfText) { $selfText = @($adapters | ForEach-Object { $_.Ip }) -join ', ' }
    $lines.Add((Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（{2}）' $startTime $machineName $selfText))
    foreach ($net in $outside) {
        if ($net.Coverage -eq 'partial') {
            $lines.Add(('※ {0} の一部は、この PC のネットワークの外です。その部分は MAC アドレスが分かりません（ping に応答した機器だけを載せます）' -f $net.Label))
        } else {
            $lines.Add(('※ {0} は、この PC のネットワークの外です。MAC アドレスは分かりません（ping に応答した機器だけを載せます）' -f $net.Label))
        }
    }
    $lines.Add('')

    $widthName = 8
    foreach ($device in $sorted) { $widthName = [math]::Max($widthName, (Get-DisplayWidth $device.Name)) }
    if ($sorted.Count -eq 0) {
        $lines.Add('機器は見つかりませんでした')
    } else {
        $lines.Add(' ' + (Format-Column 'IP アドレス' 15 $false) + '  ' + (Format-Column '名前' $widthName $false) + '  ' + (Format-Column 'MAC アドレス' 17 $false) + '  ' + (Format-Column 'ping' 8 $false) + '  種類')
        foreach ($device in $sorted) {
            $pingText = if ($device.Ping -eq '応答') { Format-Invariant '{0,4} ms' $device.Ms } elseif ($device.Ping) { $device.Ping } else { '   -' }
            $lines.Add(' ' + (Format-Column (ConvertTo-IpText $device.Ip) 15 $false) + '  ' + (Format-Column $device.Name $widthName $false) + '  ' + (Format-Column $device.Mac 17 $false) + '  ' + (Format-Column $pingText 8 $false) + '  ' + $device.Kind)
        }
    }
    $answered = @($sorted | Where-Object { $_.Ping -eq '応答' }).Count
    $silent   = @($sorted | Where-Object { $_.Ping -eq '応答なし' }).Count
    $selfCount = @($sorted | Where-Object { $_.Kind -eq 'この PC' }).Count
    $parts = @((Format-Invariant 'ping に応答 {0:N0}' $answered), (Format-Invariant '応答なし {0:N0}' $silent))
    if ($selfCount -gt 0) { $parts += Format-Invariant 'この PC {0:N0}' $selfCount }
    $lines.Add('')
    $lines.Add((Format-Invariant '機器 {0:N0} 台（{1}）  所要 {2}' $sorted.Count ($parts -join '・') (Format-Duration $elapsed)))

    #endregion

    #region 保存

    if ($runDir) {
        try {
            $lines.Add(('保存先: {0}' -f $runDir))
            [IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), [string[]]$lines, $Utf8Bom)
            # 1 台だけのときに配列がほどけないよう、List に入れる
            $rows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($device in $sorted) {
                $rows.Add([object[]]@((ConvertTo-IpText $device.Ip), (Protect-ExcelText $device.Name), $device.NameBy, $device.Mac, $device.Ping,
                    $(if ($null -ne $device.Ms) { $device.Ms.ToString('N0', $Invariant) } else { '' }), $device.Kind, $device.Network))
            }
            Save-Csv (Join-Path $runDir 'devices.csv') $DeviceColumns $rows
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
        }
    }

    #endregion

    $finished = $true
} catch {
    # 想定外のエラーは、Ctrl+C の「中断しました」と区別して知らせる
    $unexpected = $true
    throw
} finally {
    Write-Progress -Activity 'ネットワークの機器一覧' -Completed
    if (-not $finished -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete($runDir, $true) } catch { } }
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

#region 画面表示

foreach ($line in $lines) {
    if ($line.StartsWith('保存先: ')) { continue }
    if ($line.StartsWith('※')) { Write-Host $line -ForegroundColor Yellow } else { Write-Host $line }
}
if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
} elseif ($saveWarning) {
    Write-Warning $saveWarning
}

#endregion

exit 0
