<#
.SYNOPSIS
    複数のサーバーについて、名前解決・ping・TCP ポートへの接続をまとめて確かめます。

.DESCRIPTION
    CSV ファイルまたはコマンドラインで指定した対象について、名前解決・ping・TCP の接続を同時に確かめ、
    結果を画面に表示します。結果は CSV（BOM 付き UTF-8）でも保存します。
    ポートを指定した対象は TCP で接続できれば OK、ポートを指定しない対象は ping に応答があれば OK です。

    調べる相手に変更を加えることはありません。TCP は接続できたらすぐに切断し、データは送りません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Test-Connectivity.ps1 -List .\targets.csv

.PARAMETER List
    対象を書いた CSV ファイル。列は「ホスト,ポート,メモ」です（ポートとメモは省略できます）。

.PARAMETER Target
    対象をコマンドラインで指定します。「ホスト」か「ホスト:ポート」を、カンマで区切って 1 つの文字列にします。
    IPv6 アドレスにポートを付けるときは、[fe80::1]:443 のように [ ] で囲みます。

.PARAMETER Timeout
    1 件あたりの待ち時間（秒、1～60）。名前解決・ping・TCP の接続それぞれに適用します。既定値は 2 です。

.PARAMETER Repeat
    指定した秒数ごと（1～3600）に、Ctrl+C で止めるまで確かめ続けます。
    2 回目からは、判定が変わった対象だけを表示します。

.PARAMETER NewList
    対象のひな形の CSV を、指定したパスに作って終わります。同じ名前のファイルがあるときは上書きしません。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.PARAMETER ConfigFile
    設定ファイルのパス。List、Target、Timeout、Repeat、OutputFolder、NoSave を「項目 = 値」の形で書いておけます。
    コマンドラインで指定した値は、設定ファイルより優先します。書き方は README を参照してください。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Test-Connectivity.ps1 -NewList .\targets.csv

    対象のひな形 targets.csv を作ります。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Test-Connectivity.ps1 -List .\targets.csv

    targets.csv に書いた対象を確かめます。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Test-Connectivity.ps1 -Target "web01:443,db01:1433,192.168.1.1"

    web01 の 443 番と db01 の 1433 番に TCP で接続できるか、192.168.1.1 が ping に応答するかを確かめます。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Test-Connectivity.ps1 -List .\targets.csv -Repeat 5

    5 秒ごとに確かめ続け、判定が変わったときだけ表示します。Ctrl+C で止めます。

.NOTES
    Test-Connectivity 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$List,

    [string[]]$Target = @(),

    [ValidateRange(1, 60)]
    [int]$Timeout = 2,

    [ValidateRange(1, 3600)]
    [int]$Repeat,

    [string]$NewList,

    [string]$OutputFolder,

    [switch]$NoSave,

    [string]$ConfigFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$BatchSize     = 50    # 同時に確かめる対象の数（相手やネットワークに負荷をかけすぎないため）
$PollMs        = 10    # 非同期の処理が終わったかを見回る間隔（ミリ秒）
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$Utf8NoBom     = New-Object Text.UTF8Encoding $false
$Columns       = @('No', '判定', 'ホスト', 'IP アドレス', 'ping', '応答時間(ms)', 'ポート', 'TCP', '接続時間(ms)', 'メモ', '確認日時', '詳細')
$ChangeColumns = @('日時', 'No', 'ホスト', 'ポート', '変化前', '変化後', '詳細')
$clock         = [Diagnostics.Stopwatch]::StartNew()   # 接続時間などを測る共通の時計

# .NET Framework の名前解決はスレッドプールの上で 1 件ずつ止まって待つため、応答の遅い名前が多いと
# スレッドが足りなくなり、後ろの正常な名前まで待ち時間を超えてしまう。同時に確かめる数だけ、スレッドを先に用意しておく
try {
    $minWorker = 0; $minIo = 0
    [Threading.ThreadPool]::GetMinThreads([ref]$minWorker, [ref]$minIo)
    if ($minWorker -lt $BatchSize * 2) { [void][Threading.ThreadPool]::SetMinThreads($BatchSize * 2, $minIo) }
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

# .NET のメソッドで起きた例外は MethodInvocationException に、Task の例外は AggregateException に包まれて届くため、元の例外を取り出す
function Get-InnerException([Exception]$Exception) {
    while (($Exception -is [Management.Automation.MethodInvocationException] -or $Exception -is [AggregateException]) -and $Exception.InnerException) {
        $Exception = $Exception.InnerException
    }
    return $Exception
}

# .NET のメッセージは途中や末尾に改行が入ることがあるため、1 行にまとめる
function Get-ErrorText([Exception]$Exception) {
    return ($Exception.Message -replace '\s*[\r\n]+\s*', ' ').Trim()
}

# Excel は「001」「2026-10-31」のような値を数値や日付に変えてしまうため、文字のまま見せたい値は ="..." の形で出す
# 値の中の " は "" に重ねる。Excel の式に書ける文字列は 255 文字までなので、それより長い値は囲まずにそのまま出す
function ConvertTo-ExcelText([string]$Text) {
    if ($Text -eq '' -or $Text.Length -gt 255) { return $Text }
    return '="' + $Text.Replace('"', '""') + '"'
}

# 5.1 と 7 のどちらでも BOM 付き UTF-8 で保存する（Export-Csv の -Encoding UTF8 は 7 だと BOM なしになる）。
# 行が 0 件のときは見出しだけを書く
function Save-Csv($Rows, [string[]]$Columns, [string]$FilePath) {
    if ($null -eq $Rows) { $Rows = @() }
    $lines = @($Rows | ConvertTo-Csv -NoTypeInformation)
    if ($lines.Count -eq 0) { $lines = @(($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') }
    [IO.File]::WriteAllLines($FilePath, [string[]]$lines, $Utf8Bom)
}

function Stop-WithError([string]$Message) {
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value) {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -List "C:\tools\targets.csv"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
}

# 5.1 の Get-Content は BOM のない UTF-8 を Shift_JIS として読んでしまうため、文字コードを自分で判定する。
# BOM があればそれに従い、なければ UTF-8 として読めるかを確かめ、読めなければ Shift_JIS とみなす
function Read-ConfigText([string]$FilePath) {
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return (New-Object Text.UTF8Encoding $false).GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2) }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return [Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2) }
    try { return (New-Object Text.UTF8Encoding $false, $true).GetString($bytes) } catch { }
    return [Text.Encoding]::GetEncoding(932).GetString($bytes)
}

# 設定ファイルの値は、前後を同じ引用符で囲んだときだけ外す（片方だけの引用符は書き間違いとして残す）
function Remove-ConfigQuote([string]$Text) {
    $t = $Text.Trim()
    if ($t.Length -ge 2 -and (($t.StartsWith('"') -and $t.EndsWith('"')) -or ($t.StartsWith("'") -and $t.EndsWith("'")))) {
        $t = $t.Substring(1, $t.Length - 2)
    }
    return $t
}

# 設定ファイルの書き間違いは、何行目かと設定ファイルのパスを添えて止める
function Stop-WithConfigError([int]$LineNumber, [string]$Message, [string]$Hint) {
    $text = '設定ファイルの {0} 行目: {1}' -f $LineNumber, $Message
    if ($Hint) { $text += "`r`n" + $Hint }
    Stop-WithError ($text + "`r`n" + ('ファイル: {0}' -f $configPath))
}

# 設定ファイルから来た値なら、何行目かを添える（見つからない・書き込めないときに原因を探しやすくするため）
function Get-ConfigNote([string]$Name) {
    if ($configApplied.ContainsKey($Name)) { return '（設定ファイルの {0} 行目）' -f $configApplied[$Name] }
    return ''
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

# 相対パスを、PowerShell の現在の場所を基準にしたフルパスにする（ファイルがなくてもよい）
function Get-FullPath([string]$Text) {
    $full = $Text
    try {
        $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Text)
        $full = [IO.Path]::GetFullPath($full)
    } catch { }
    return $full
}

# %TEMP% などの環境変数を展開する。名前を打ち間違えると %名前% のまま残り、その名前のフォルダに保存してしまうため止める。
# URL エンコードされたフォルダ名（%E3%81%82 など）を誤って止めないよう、16 進 2 桁の %XX% は対象にしない
function Expand-ConfigVariable([int]$LineNumber, [string]$Text) {
    $expanded = [Environment]::ExpandEnvironmentVariables($Text)
    if ($expanded -match '%(?![0-9A-Fa-f]{2}%)[A-Za-z_][^%\\=\s]*%') { Stop-WithConfigError $LineNumber ('定義されていない環境変数です: {0}' -f $Matches[0]) }
    return $expanded
}

# CSV の全体を、レコード（Line = 始まりの行番号、Fields = 列の値）に分ける。
# "" で囲んだ値の中のカンマ・""・改行（Excel のセル内改行）にも対応するため、行に分けてからではなく 1 文字ずつ読む。
# ConvertFrom-Csv は、Excel が付ける空の見出しが 2 つ以上あると止まり、何行目かも分からないため使わない
function ConvertFrom-CsvText([string]$Text) {
    $records    = New-Object 'System.Collections.Generic.List[object]'
    $fields     = New-Object 'System.Collections.Generic.List[string]'
    $current    = New-Object Text.StringBuilder
    $quoted     = $false
    $lineNumber = 1
    $startLine  = 1
    for ($i = 0; $i -le $Text.Length; $i++) {
        $atEnd = $i -eq $Text.Length
        $ch    = if ($atEnd) { [char]0 } else { $Text[$i] }
        if (-not $atEnd -and ($ch -eq "`n" -or ($ch -eq "`r" -and -not ($i + 1 -lt $Text.Length -and $Text[$i + 1] -eq "`n")))) { $lineNumber++ }
        if ($quoted -and -not $atEnd) {
            if ($ch -eq '"') {
                if ($i + 1 -lt $Text.Length -and $Text[$i + 1] -eq '"') { [void]$current.Append('"'); $i++ } else { $quoted = $false }
            } else {
                [void]$current.Append($ch)
            }
        } elseif ($ch -eq '"') {
            $quoted = $true
        } elseif ($ch -eq ',') {
            $fields.Add($current.ToString())
            [void]$current.Clear()
        } elseif ($atEnd -or $ch -eq "`r" -or $ch -eq "`n") {
            # CRLF は、CR のところでレコードを閉じ、LF は空のレコードとして読み飛ばす
            $fields.Add($current.ToString())
            [void]$current.Clear()
            $isBlank = $fields.Count -eq 1 -and $fields[0].Trim() -eq ''
            if (-not $isBlank) { $records.Add([pscustomobject]@{ Line = $startLine; Fields = $fields.ToArray() }) }
            $fields.Clear()
            $startLine = $lineNumber
        } else {
            [void]$current.Append($ch)
        }
    }
    return , $records.ToArray()
}

# IP アドレスとして書かれた値なら IPAddress を、そうでなければ $null を返す（$null のときは名前解決する）。
# IPAddress.TryParse は「001」を 0.0.0.1 と、「192.168.001.010」を 8 進数として 192.168.1.8 と読むため、
# IPv4 は 4 つの数字を 10 進数として自分で読む。IPv6 は : を含むものだけを TryParse に任せる
function ConvertTo-IPAddress([string]$Text) {
    if ($Text -match '^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$') {
        $bytes = New-Object 'byte[]' 4
        for ($i = 0; $i -lt 4; $i++) {
            $part = [int]::Parse($Matches[$i + 1], $Invariant)
            if ($part -gt 255) { return $null }
            $bytes[$i] = [byte]$part
        }
        return New-Object Net.IPAddress (, $bytes)
    }
    $parsed = $null
    if ($Text.Contains(':') -and [Net.IPAddress]::TryParse($Text, [ref]$parsed)) { return $parsed }
    return $null
}

# ポートは 1～65535 の数値だけを受け付ける（+80 や 80.0 のような書き方は誤りとして扱う）
function ConvertTo-Port([string]$Text) {
    $number = 0
    if ([int]::TryParse($Text, [Globalization.NumberStyles]::None, $Invariant, [ref]$number) -and $number -ge 1 -and $number -le 65535) { return $number }
    return $null
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

function Format-Cell([string]$Text, [int]$Width, [bool]$AlignRight) {
    $pad = ' ' * [math]::Max(0, $Width - (Get-DisplayWidth $Text))
    if ($AlignRight) { return $pad + $Text }
    return $Text + $pad
}

#endregion

#region 確かめる処理

# 非同期の処理（Task）を、待ち時間まで見回り、始めてから終わるまでの時間（Started からの差）を記録する。
# Task.Wait は例外の起きた Task で止まるうえ、終わった時刻も分からないため、短い間隔で IsCompleted を見る（誤差は PollMs ほど）
function Wait-AsyncTask($Jobs, [int]$WaitMs) {
    $deadline = $clock.ElapsedMilliseconds + $WaitMs
    $pending  = @($Jobs | Where-Object { $_.Task })
    while ($pending.Count -gt 0) {
        $rest = @()
        foreach ($job in $pending) {
            if ($job.Task.IsCompleted) { $job.Elapsed = $clock.ElapsedMilliseconds - $job.Started } else { $rest += $job }
        }
        $pending = $rest
        if ($pending.Count -eq 0 -or $clock.ElapsedMilliseconds -ge $deadline) { break }
        [Threading.Thread]::Sleep($PollMs)
    }
}

# 1 回分の確認。対象を BatchSize 件ずつ、名前解決 → ping と TCP の順に、同時に確かめる
function Invoke-CheckRound($Items) {
    $waitMs  = $Timeout * 1000
    $results = New-Object 'System.Collections.Generic.List[object]'
    for ($offset = 0; $offset -lt $Items.Count; $offset += $BatchSize) {
        $last  = [math]::Min($offset + $BatchSize, $Items.Count) - 1
        $batch = @($Items[$offset..$last])
        $checkedAt = Get-Date

        # 1. 名前解決（IP アドレスが書かれていれば、そのまま使う）
        $jobs = @(
            foreach ($item in $batch) {
                $job = [pscustomobject]@{
                    Item = $item; Task = $null; Started = 0L; Elapsed = $null; Address = $null; Detail = ''
                    Ping = '-'; PingMs = $null; Tcp = '-'; TcpMs = $null
                }
                $parsed = ConvertTo-IPAddress $item.Host
                if ($parsed) {
                    $job.Address = $parsed
                } else {
                    $job.Started = $clock.ElapsedMilliseconds
                    try { $job.Task = [Net.Dns]::GetHostAddressesAsync($item.Host) } catch { $job.Detail = '名前解決: ' + (Get-ErrorText (Get-InnerException $_.Exception)) }
                }
                $job
            }
        )
        Wait-AsyncTask $jobs $waitMs
        foreach ($job in $jobs) {
            if (-not $job.Task) { continue }
            if ($job.Task.Status -eq 'RanToCompletion') {
                $addresses = @($job.Task.Result)
                $ipv4      = @($addresses | Where-Object { $_.AddressFamily -eq 'InterNetwork' })
                if ($ipv4.Count -gt 0) { $job.Address = $ipv4[0] } elseif ($addresses.Count -gt 0) { $job.Address = $addresses[0] }
                if (-not $job.Address) { $job.Detail = '名前解決: アドレスが見つかりません' }
            } elseif ($job.Task.IsFaulted) {
                $job.Detail = '名前解決: ' + (Get-ErrorText (Get-InnerException $job.Task.Exception))
            } else {
                $job.Detail = '名前解決: 待ち時間内に応答がありません'
            }
            $job.Task = $null
        }

        # 2. ping と TCP を同時に始める
        $netJobs = New-Object 'System.Collections.Generic.List[object]'
        foreach ($job in $jobs) {
            if (-not $job.Address) { continue }
            $pinger = New-Object Net.NetworkInformation.Ping
            $pingJob = [pscustomobject]@{ Kind = 'ping'; Owner = $job; Handle = $pinger; Task = $null; Started = $clock.ElapsedMilliseconds; Elapsed = $null; Error = $null }
            try { $pingJob.Task = $pinger.SendPingAsync($job.Address, $waitMs) } catch { $pingJob.Error = Get-InnerException $_.Exception }
            $netJobs.Add($pingJob)
            if ($job.Item.Port) {
                $client = New-Object Net.Sockets.TcpClient ($job.Address.AddressFamily)
                $tcpJob = [pscustomobject]@{ Kind = 'tcp'; Owner = $job; Handle = $client; Task = $null; Started = $clock.ElapsedMilliseconds; Elapsed = $null; Error = $null }
                try { $tcpJob.Task = $client.ConnectAsync($job.Address, $job.Item.Port) } catch { $tcpJob.Error = Get-InnerException $_.Exception }
                $netJobs.Add($tcpJob)
            }
        }
        Wait-AsyncTask $netJobs $waitMs

        foreach ($netJob in $netJobs) {
            $job   = $netJob.Owner
            $task  = $netJob.Task
            # $error は自動変数（$Error）と同じ名前になるため使わない
            $failure = $netJob.Error
            if ($task -and $task.IsFaulted) { $failure = Get-InnerException $task.Exception }
            if ($netJob.Kind -eq 'ping') {
                if ($task -and $task.Status -eq 'RanToCompletion' -and "$($task.Result.Status)" -eq 'Success') {
                    $job.Ping   = '応答あり'
                    $job.PingMs = $task.Result.RoundtripTime
                } else {
                    $job.Ping = '応答なし'
                    if ($failure) { $job.Detail = (@($job.Detail, ('ping: ' + (Get-ErrorText $failure))) | Where-Object { $_ }) -join ' / ' }
                }
            } else {
                if ($task -and $task.Status -eq 'RanToCompletion') {
                    $job.Tcp   = '接続できた'
                    $job.TcpMs = $netJob.Elapsed
                } elseif ($failure -is [Net.Sockets.SocketException] -and $failure.SocketErrorCode -eq 'ConnectionRefused') {
                    $job.Tcp = '拒否された'
                } elseif (-not $failure -or ($failure -is [Net.Sockets.SocketException] -and $failure.SocketErrorCode -eq 'TimedOut')) {
                    $job.Tcp = 'タイムアウト'
                } else {
                    $job.Tcp = 'エラー'
                    $job.Detail = (@($job.Detail, ('TCP: ' + (Get-ErrorText $failure))) | Where-Object { $_ }) -join ' / '
                }
            }
            # 終わっていない接続や ping も、ここで打ち切る
            try { if ($netJob.Kind -eq 'tcp') { $netJob.Handle.Close() } else { $netJob.Handle.Dispose() } } catch { }
        }

        foreach ($job in $jobs) {
            $item = $job.Item
            if (-not $job.Address) {
                $ok = $false
            } elseif ($item.Port) {
                $ok = $job.Tcp -eq '接続できた'
            } else {
                $ok = $job.Ping -eq '応答あり'
            }
            $results.Add([pscustomobject]@{
                No        = $item.No
                Ok        = $ok
                Host      = $item.Host
                Address   = $(if ($job.Address) { $job.Address.ToString() } else { '' })
                Ping      = $job.Ping
                PingMs    = $job.PingMs
                Port      = $item.Port
                Tcp       = $(if ($item.Port -and $job.Address) { $job.Tcp } else { '-' })
                TcpMs     = $job.TcpMs
                Memo      = $item.Memo
                CheckedAt = $checkedAt
                Detail    = $job.Detail
            })
        }
    }
    return , $results.ToArray()
}

# 変化を比べるときの状態（ポートのある対象は TCP、ない対象は ping）
function Get-StateText($Result) {
    if (-not $Result.Address) { return '名前解決失敗' }
    if ($Result.Port) { return $Result.Tcp }
    return $Result.Ping
}

function Get-TargetLabel($Result) {
    if (-not $Result.Port) { return $Result.Host }
    if ($Result.Host.Contains(':')) { return '[{0}]:{1}' -f $Result.Host, $Result.Port }
    return '{0}:{1}' -f $Result.Host, $Result.Port
}

#endregion

#region 表示と保存

# 画面と summary.txt に出す表。NG の行は赤で表示するため、行ごとに判定を添えて返す
function Get-ResultTable($Results) {
    $head  = @('No', '判定', 'ホスト', 'IP アドレス', 'ping', 'ポート', 'TCP', 'メモ')
    $right = @($true, $false, $false, $false, $false, $true, $false, $false)
    $rows  = @(
        foreach ($r in $Results) {
            $pingText = if ($r.Ping -eq '応答あり') { Format-Invariant '{0} ms' $r.PingMs } else { $r.Ping }
            , @([string]$r.No, $(if ($r.Ok) { 'OK' } else { 'NG' }), $r.Host,
                $(if ($r.Address) { $r.Address } else { '名前解決失敗' }), $pingText,
                $(if ($r.Port) { [string]$r.Port } else { '' }), $r.Tcp, $r.Memo)
        }
    )
    $widths = @(foreach ($i in 0..($head.Count - 1)) {
        $w = Get-DisplayWidth $head[$i]
        foreach ($row in $rows) { $w = [math]::Max($w, (Get-DisplayWidth $row[$i])) }
        $w
    })
    $lines = New-Object 'System.Collections.Generic.List[object]'
    $cells = @(foreach ($i in 0..($head.Count - 1)) { Format-Cell $head[$i] $widths[$i] $right[$i] })
    $lines.Add([pscustomobject]@{ Text = (' ' + ($cells -join '  ')).TrimEnd(); Ok = $true })
    for ($n = 0; $n -lt $rows.Count; $n++) {
        $cells = @(foreach ($i in 0..($head.Count - 1)) { Format-Cell $rows[$n][$i] $widths[$i] $right[$i] })
        $lines.Add([pscustomobject]@{ Text = (' ' + ($cells -join '  ')).TrimEnd(); Ok = $Results[$n].Ok })
    }
    return , $lines.ToArray()
}

function Get-HeaderLines($Results, [TimeSpan]$Elapsed, [int]$Round) {
    $okCount = @($Results | Where-Object { $_.Ok }).Count
    $ngCount = $Results.Count - $okCount
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add((Format-Invariant '疎通一括確認 (Test-Connectivity {0})' $ScriptVersion))
    $lines.Add((Format-Invariant '実行元: {0}（{1}）   {2:yyyy-MM-dd HH:mm:ss}' $machineName $(if ($localAddresses) { $localAddresses -join ', ' } else { '不明' }) $startTime))
    if ($dnsServers -or $gateways) {
        $lines.Add(('DNS サーバー: {0}   デフォルトゲートウェイ: {1}' -f $(if ($dnsServers) { $dnsServers -join ', ' } else { '不明' }), $(if ($gateways) { $gateways -join ', ' } else { '不明' })))
    }
    if ($configPath) { $lines.Add(('設定ファイル: {0}' -f $configPath)) }
    $lines.Add(('対象: {0}   タイムアウト {1} 秒' -f ($sourceTexts -join '、'), $Timeout))
    if ($Repeat) {
        $lines.Add((Format-Invariant '繰り返し: {0} 秒ごと   結果は {1:N0} 回目の確認（{2:HH:mm:ss}）' $Repeat $Round $Results[0].CheckedAt))
    }
    $lines.Add((Format-Invariant '結果: OK {0:N0} 件  NG {1:N0} 件   所要 {2}' $okCount $ngCount (Format-Duration $Elapsed)))
    return , $lines.ToArray()
}

function Show-Report($Header, $Table) {
    foreach ($line in $Header) { Write-Host $line }
    Write-Host ''
    foreach ($row in $Table) {
        if ($row.Ok) { Write-Host $row.Text } else { Write-Host $row.Text -ForegroundColor Red }
    }
}

function ConvertTo-ResultRow($Result) {
    return [pscustomobject]@{
        'No'           = $Result.No
        '判定'         = $(if ($Result.Ok) { 'OK' } else { 'NG' })
        'ホスト'       = (ConvertTo-ExcelText $Result.Host)
        'IP アドレス'  = $Result.Address
        'ping'         = $Result.Ping
        '応答時間(ms)' = $(if ($null -ne $Result.PingMs) { ([long]$Result.PingMs).ToString('N0', $Invariant) } else { '' })
        'ポート'       = $(if ($Result.Port) { $Result.Port } else { '' })
        'TCP'          = $Result.Tcp
        '接続時間(ms)' = $(if ($null -ne $Result.TcpMs) { ([long]$Result.TcpMs).ToString('N0', $Invariant) } else { '' })
        'メモ'         = (ConvertTo-ExcelText $Result.Memo)
        '確認日時'     = $Result.CheckedAt.ToString('yyyy/MM/dd HH:mm:ss', $Invariant)
        '詳細'         = $Result.Detail
    }
}

# summary.txt と results.csv を書く。-Repeat のときは、変化の一覧も summary.txt に入れる
function Save-Report($Header, $Table, $Results, $ChangeLines) {
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.AddRange([string[]]$Header)
    $lines.Add('')
    foreach ($row in $Table) { $lines.Add($row.Text) }
    if ($null -ne $ChangeLines) {
        $lines.Add('')
        $lines.Add('■ 判定が変わった記録')
        if ($ChangeLines.Count -eq 0) { $lines.Add('（変化なし）') } else { $lines.AddRange([string[]]$ChangeLines.ToArray()) }
    }
    [IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), [string[]]$lines, $Utf8Bom)
    Save-Csv @($Results | ForEach-Object { ConvertTo-ResultRow $_ }) $Columns (Join-Path $runDir 'results.csv')
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの。数値の範囲は param の ValidateRange と同じにする
$configItems   = 'List', 'Target', 'Timeout', 'Repeat', 'OutputFolder', 'NoSave'
$configRanges  = @{ Timeout = 1, 60; Repeat = 1, 3600 }
$configPath    = $null
$configApplied = @{}   # 設定ファイルの値を使った項目 → 行番号

if ($ConfigFile) {
    $configText = Remove-Quote $ConfigFile
    if ($configText.Contains('"')) { Stop-WithQuoteError '-ConfigFile' $configText }
    $configPath = Get-FullPath (Resolve-StartPath $configText)
    if (-not [IO.File]::Exists($configPath)) {
        # エクスプローラーで拡張子を表示していないと、メモ帳で check.ini.txt として保存してしまうことがある
        $message = '設定ファイルが見つかりません: {0}' -f $configPath
        if ([IO.File]::Exists($configPath + '.txt')) { $message += "`r`n" + ('拡張子が .txt になっていませんか: {0}.txt' -f $configPath) }
        Stop-WithError $message
    }
    try {
        $configLines = (Read-ConfigText $configPath) -split '\r\n|\n|\r'
    } catch {
        Stop-WithError ('設定ファイルを読めません: {0}（{1}）' -f $configPath, (Get-ErrorText (Get-InnerException $_.Exception)))
    }

    $configDir    = Split-Path -LiteralPath $configPath   # -LiteralPath には -Parent を付けられない（親フォルダを返すのが既定）
    $configValues = @{}
    $configLineOf = @{}
    $targetItems  = New-Object 'System.Collections.Generic.List[string]'
    $lineNumber   = 0
    foreach ($rawLine in $configLines) {
        $lineNumber++
        $lineText = $rawLine.Trim()
        # # と ; はパスに使える文字なので、行の先頭にあるときだけコメントとみなす
        if ($lineText -eq '' -or $lineText.StartsWith('#') -or $lineText.StartsWith(';')) { continue }
        $separator = $lineText.IndexOf('=')
        if ($separator -le 0) { Stop-WithConfigError $lineNumber ('「項目 = 値」の形になっていません: {0}' -f $lineText) }

        $itemKey  = $lineText.Substring(0, $separator).Trim()
        $itemName = $configItems | Where-Object { $_ -eq $itemKey }
        if (-not $itemName) { Stop-WithConfigError $lineNumber ('知らない項目です: {0}' -f $itemKey) ('使える項目: {0}' -f ($configItems -join ', ')) }

        $rawValue  = $lineText.Substring($separator + 1).Trim()
        $itemValue = Remove-ConfigQuote $rawValue
        # 値が空の行は、書かなかったのと同じ扱い（ひな形の空の行の下に書き足しても、2 回書いたことにしない）
        if ($itemValue -eq '') { continue }

        # Target は長くなりやすいので複数の行に分けて書ける。ほかの項目は、どちらが効いているか分からなくなるため 1 回だけ
        if ($itemName -ne 'Target' -and $configLineOf.ContainsKey($itemName)) {
            Stop-WithConfigError $lineNumber ('{0} が 2 回書かれています（{1} 行目と {2} 行目）' -f $itemName, $configLineOf[$itemName], $lineNumber)
        }
        $configLineOf[$itemName] = $lineNumber

        if ($itemName -eq 'Target') {
            if ($itemValue.Contains('"')) { Stop-WithConfigError $lineNumber ('Target の値に引用符（"）が残っています: {0}' -f $itemValue) }
            $targetItems.Add($itemValue)
        } elseif ($itemName -eq 'NoSave') {
            if ($itemValue -eq 'true') { $configValues['NoSave'] = $true }
            elseif ($itemValue -eq 'false') { $configValues['NoSave'] = $false }
            else { Stop-WithConfigError $lineNumber ('NoSave は true か false で指定してください: {0}' -f $itemValue) }
        } elseif ($configRanges.ContainsKey($itemName)) {
            $range  = $configRanges[$itemName]
            $number = 0
            if (-not [int]::TryParse($itemValue, [Globalization.NumberStyles]::None, $Invariant, [ref]$number) -or
                $number -lt $range[0] -or $number -gt $range[1]) {
                Stop-WithConfigError $lineNumber ('{0} は {1}～{2} の数値で指定してください: {3}' -f $itemName, $range[0], $range[1], $itemValue)
            }
            $configValues[$itemName] = $number
        } else {
            # パスの項目は、コマンドプロンプトで書くときと同じように %TEMP% などの環境変数を展開する
            if ($itemValue.Contains('"')) { Stop-WithConfigError $lineNumber ('{0} の値に引用符（"）が残っています: {1}' -f $itemName, $itemValue) }
            $itemValue = Expand-ConfigVariable $lineNumber $itemValue
            # 相対パスは、設定ファイルのあるフォルダを基準にする（タスクスケジューラでは現在のフォルダが System32 になるため）
            if (-not (Test-RootedPath $itemValue)) {
                $itemValue = $configDir.TrimEnd('\') + '\' + $itemValue
                try { $itemValue = [IO.Path]::GetFullPath($itemValue) } catch { }
            }
            $configValues[$itemName] = $itemValue
        }
    }
    if ($targetItems.Count -gt 0) { $configValues['Target'] = $targetItems.ToArray() }

    # コマンドラインで指定した項目は、設定ファイルより優先する
    foreach ($itemName in @($configValues.Keys)) {
        if (-not $PSBoundParameters.ContainsKey($itemName)) {
            Set-Variable -Name $itemName -Value $configValues[$itemName]
            $configApplied[$itemName] = $configLineOf[$itemName]
        }
    }
}

#endregion

#region ひな形を作る（-NewList）

if ($NewList) {
    $newText = Remove-Quote $NewList
    if ($newText.Contains('"')) { Stop-WithQuoteError '-NewList' $newText }
    $newPath = Get-FullPath (Resolve-StartPath $newText)
    if ([IO.File]::Exists($newPath) -or [IO.Directory]::Exists($newPath)) {
        Stop-WithError ('同じ名前のファイルがあるため、ひな形を作りません（上書きしません）: {0}' -f $newPath)
    }
    $template = @(
        'ホスト,ポート,メモ'
        '# 先頭が # の行は確かめません。ポートを空欄にすると ping だけを確かめます,,'
        'web01,443,Web サーバー（HTTPS）'
        'db01,1433,SQL Server'
        '192.168.1.1,,ゲートウェイ（ping だけ）'
    )
    try {
        # Excel でダブルクリックして開けるよう、BOM 付き UTF-8 で書く。既存のファイルは上書きしない
        $stream = New-Object IO.FileStream ($newPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
        $writer = New-Object IO.StreamWriter ($stream, $Utf8Bom)
        try { foreach ($line in $template) { $writer.WriteLine($line) } } finally { $writer.Dispose() }
    } catch {
        Stop-WithError ('ひな形を作れません: {0}（{1}）' -f $newPath, (Get-ErrorText (Get-InnerException $_.Exception)))
    }
    Write-Host ('ひな形を作りました: {0}' -f $newPath)
    Write-Host 'Excel やメモ帳で対象を書き換えてから、-List で指定して実行してください。'
    exit 0
}

#endregion

#region 対象を読む

$items       = New-Object 'System.Collections.Generic.List[object]'
$sourceTexts = @()

if ($List) {
    $listText = Remove-Quote $List
    if ($listText.Contains('"')) { Stop-WithQuoteError '-List' $listText }
    $listPath = Get-FullPath (Resolve-StartPath $listText)
    if (-not [IO.File]::Exists($listPath)) {
        $message = ('対象の CSV が見つかりません: {0}' -f $listPath) + (Get-ConfigNote 'List')
        if ([IO.File]::Exists($listPath + '.txt')) { $message += "`r`n" + ('拡張子が .txt になっていませんか: {0}.txt' -f $listPath) }
        Stop-WithError $message
    }
    try {
        $listRecords = ConvertFrom-CsvText (Read-ConfigText $listPath)
    } catch {
        Stop-WithError ('対象の CSV を読めません: {0}（{1}）' -f $listPath, (Get-ErrorText (Get-InnerException $_.Exception)))
    }

    $hostCol  = -1
    $portCol  = -1
    $memoCol  = -1
    $fromList = 0
    foreach ($record in $listRecords) {
        $lineNumber = $record.Line
        $fields     = $record.Fields
        if ($hostCol -lt 0) {
            # 最初の空でない行は見出し。日本語と英語のどちらの見出しも受け付ける
            for ($i = 0; $i -lt $fields.Count; $i++) {
                $name = $fields[$i].Trim()
                if ($name -eq 'ホスト' -or $name -eq 'Host') { $hostCol = $i }
                elseif ($name -eq 'ポート' -or $name -eq 'Port') { $portCol = $i }
                elseif ($name -eq 'メモ' -or $name -eq 'Memo') { $memoCol = $i }
            }
            if ($hostCol -lt 0) {
                Stop-WithError (('対象の CSV に「ホスト」の列がありません。1 行目を見出し（ホスト,ポート,メモ）にしてください') + "`r`n" + ('ファイル: {0}' -f $listPath))
            }
            continue
        }
        $hostText = if ($hostCol -lt $fields.Count) { $fields[$hostCol].Trim() } else { '' }
        if ($hostText -eq '' -or $hostText.StartsWith('#')) { continue }
        $portText = if ($portCol -ge 0 -and $portCol -lt $fields.Count) { $fields[$portCol].Trim() } else { '' }
        # セル内改行は、画面の表がくずれないよう空白にする
        $memoText = if ($memoCol -ge 0 -and $memoCol -lt $fields.Count) { ($fields[$memoCol] -replace '\s*[\r\n]+\s*', ' ').Trim() } else { '' }
        $port = 0
        if ($portText -ne '') {
            $port = ConvertTo-Port $portText
            if (-not $port) {
                Stop-WithError (('対象の CSV の {0} 行目: ポートは 1～65535 の数値で指定してください: {1}' -f $lineNumber, $portText) + "`r`n" + ('ファイル: {0}' -f $listPath))
            }
        }
        $items.Add([pscustomobject]@{ No = $items.Count + 1; Host = $hostText; Port = $port; Memo = $memoText })
        $fromList++
    }
    $sourceTexts += Format-Invariant '{0}（{1:N0} 件）' $listPath $fromList
}

# -File で起動すると配列を渡せないため、カンマ区切りの文字列で受ける
$fromTarget = 0
$position   = 0
foreach ($entry in $Target) {
    foreach ($part in ($entry -split ',')) {
        $text = (Remove-Quote $part).Trim("'").Trim()
        if ($text -eq '') { continue }
        $position++
        if ($text.Contains('"')) { Stop-WithQuoteError '-Target' $text }
        $hostText = $text
        $portText = ''
        if ($text -match '^\[(.+)\](?::(.*))?$') {
            # [IPv6]:ポート
            $hostText = $Matches[1]
            if ($Matches.ContainsKey(2)) { $portText = $Matches[2] }
        } elseif (@($text.ToCharArray() | Where-Object { $_ -eq ':' }).Count -eq 1) {
            $hostText = $text.Substring(0, $text.IndexOf(':'))
            $portText = $text.Substring($text.IndexOf(':') + 1)
            if ($hostText -eq '') { Stop-WithError ('-Target の {0} つ目: ホストがありません: {1}' -f $position, $text) }
        }
        $port = 0
        if ($text.EndsWith(':') -or $portText -ne '') {
            $port = ConvertTo-Port $portText
            if (-not $port) { Stop-WithError ('-Target の {0} つ目: ポートは 1～65535 の数値で指定してください: {1}' -f $position, $text) }
        }
        $items.Add([pscustomobject]@{ No = $items.Count + 1; Host = $hostText; Port = $port; Memo = '' })
        $fromTarget++
    }
}
if ($fromTarget -gt 0) { $sourceTexts += Format-Invariant '-Target（{0:N0} 件）' $fromTarget }

if (-not $List -and $Target.Count -eq 0) { Stop-WithError '確かめる対象を、-List（CSV ファイル）か -Target で指定してください。ひな形は -NewList で作れます。' }
if ($items.Count -eq 0) { Stop-WithError '確かめる対象がありません' }

$outputText = $null
if ($OutputFolder) {
    $outputText = Remove-Quote $OutputFolder
    if ($outputText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outputText }
    $outputText = Resolve-StartPath $outputText
}

#endregion

#region 実行元の情報

$startTime      = Get-Date
$machineName    = [Environment]::MachineName
$localAddresses = @()
$dnsServers     = @()
$gateways       = @()
try {
    foreach ($nic in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
        if ("$($nic.OperationalStatus)" -ne 'Up' -or "$($nic.NetworkInterfaceType)" -eq 'Loopback') { continue }
        $props = $nic.GetIPProperties()
        foreach ($u in $props.UnicastAddresses) {
            if ($u.Address.AddressFamily -eq 'InterNetwork') { $localAddresses += $u.Address.ToString() }
        }
        try { foreach ($d in $props.DnsAddresses) { if ($d.AddressFamily -eq 'InterNetwork') { $dnsServers += $d.ToString() } } } catch { }
        try { foreach ($g in $props.GatewayAddresses) { if ($g.Address.AddressFamily -eq 'InterNetwork' -and $g.Address.ToString() -ne '0.0.0.0') { $gateways += $g.Address.ToString() } } } catch { }
    }
} catch { }
$localAddresses = @($localAddresses | Select-Object -Unique)
$dnsServers     = @($dnsServers | Select-Object -Unique)
$gateways       = @($gateways | Select-Object -Unique)

#endregion

#region 保存先の準備

# 確かめ終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir      = $null
$saveWarning = $null
if (-not $NoSave) {
    $baseDir = if ($outputText) { $outputText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        $runName = Format-Invariant 'Connectivity_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while (Test-Path -LiteralPath $runDir) { $suffix++; $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix) }
        [void][IO.Directory]::CreateDirectory($runDir)
    } catch {
        $saveWarning = ('保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, (Get-ErrorText (Get-InnerException $_.Exception))) + (Get-ConfigNote 'OutputFolder')
        $runDir = $null
        Write-Warning $saveWarning
    }
}

#endregion

#region 1 回目の確認

$itemArray = $items.ToArray()
$finished  = $false
$saveError = $null
$pathShown = $false
try {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity '疎通一括確認' -Status ('{0:N0} 件を確かめています' -f $itemArray.Count)
    $results = Invoke-CheckRound $itemArray
    Write-Progress -Activity '疎通一括確認' -Completed
    $header = Get-HeaderLines $results $sw.Elapsed 1
    $table  = Get-ResultTable $results
    Show-Report $header $table
    if ($runDir -and -not $Repeat) {
        try { Save-Report $header $table $results $null } catch { $saveError = Get-ErrorText (Get-InnerException $_.Exception) }
    }
    $finished = $true
} finally {
    Write-Progress -Activity '疎通一括確認' -Completed
    if (-not $finished -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete($runDir, $true) } catch { } }
        if (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

#endregion

#region 繰り返し確かめる（-Repeat）

if ($Repeat) {
    $round       = 1
    $roundTime   = $sw.Elapsed
    $changeLines = New-Object 'System.Collections.Generic.List[string]'
    $changesPath = $null
    if ($runDir) {
        $changesPath = Join-Path $runDir 'changes.csv'
        try { Save-Csv @() $ChangeColumns $changesPath } catch { $saveError = Get-ErrorText (Get-InnerException $_.Exception); $changesPath = $null }
    }

    # Ctrl+C をキー入力として受け取り、止めたあとに保存と終了コードの処理をできるようにする。
    # コンソールがない（入力をリダイレクトした）ときは使えないので、そのときは finally で保存する
    $keyInput  = $false
    $oldCtrlC  = $false
    try {
        if (-not [Console]::IsInputRedirected) {
            $oldCtrlC = [Console]::TreatControlCAsInput
            [Console]::TreatControlCAsInput = $true
            $keyInput = $true
        }
    } catch { }

    $statusWidth = 0
    Write-Host ''
    Write-Host ('{0} 秒ごとに確かめています。止めるときは Ctrl+C を押してください。' -f $Repeat)
    try {
        $nextStart = $startTime.AddSeconds($Repeat)
        $stopRequested = $false
        while (-not $stopRequested) {
            while ((Get-Date) -lt $nextStart) {
                if ($keyInput) {
                    try {
                        while ([Console]::KeyAvailable) {
                            $key = [Console]::ReadKey($true)
                            if ($key.Key -eq 'C' -and ($key.Modifiers -band [ConsoleModifiers]::Control)) { $stopRequested = $true }
                        }
                    } catch { }
                }
                if ($stopRequested) { break }
                [Threading.Thread]::Sleep(100)
            }
            if ($stopRequested) { break }

            # 1 回の確認が -Repeat より長くかかったときは、終わってすぐに次を始める
            $nextStart = (Get-Date).AddSeconds($Repeat)
            $round++
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $newResults = Invoke-CheckRound $itemArray
            $roundTime  = $sw.Elapsed

            for ($i = 0; $i -lt $newResults.Count; $i++) {
                $before = $results[$i]
                $after  = $newResults[$i]
                if ($before.Ok -eq $after.Ok) { continue }
                $fromText = if ($before.Ok) { 'OK' } else { 'NG' }
                $toText   = if ($after.Ok) { 'OK' } else { 'NG' }
                $detail   = '{0} → {1}' -f (Get-StateText $before), (Get-StateText $after)
                $line     = Format-Invariant '{0:HH:mm:ss}  {1} → {2}  {3}  {4}  {5}' $after.CheckedAt $fromText $toText $after.No (Get-TargetLabel $after) $detail
                $changeLines.Add($line)
                if ($statusWidth -gt 0) { Write-Host ("`r" + (' ' * $statusWidth) + "`r") -NoNewline }
                Write-Host $line -ForegroundColor $(if ($after.Ok) { 'Green' } else { 'Red' })
                if ($changesPath) {
                    try {
                        $row = [pscustomobject]@{
                            '日時' = $after.CheckedAt.ToString('yyyy/MM/dd HH:mm:ss', $Invariant); 'No' = $after.No
                            'ホスト' = (ConvertTo-ExcelText $after.Host); 'ポート' = $(if ($after.Port) { $after.Port } else { '' })
                            '変化前' = $fromText; '変化後' = $toText; '詳細' = $detail
                        }
                        # 見出しは作ったときに書いたので、2 行目（値の行）だけを書き足す。BOM は先頭に 1 つだけにする
                        $csvLine = @($row | ConvertTo-Csv -NoTypeInformation)[1]
                        [IO.File]::AppendAllText($changesPath, $csvLine + "`r`n", $Utf8NoBom)
                    } catch {
                        $saveError = Get-ErrorText (Get-InnerException $_.Exception)
                    }
                }
            }
            $results = $newResults

            $okNow  = @($results | Where-Object { $_.Ok }).Count
            $status = Format-Invariant '（確認 {0:N0} 回目  OK {1:N0} 件  NG {2:N0} 件）' $round $okNow ($results.Count - $okNow)
            $pad    = [math]::Max(0, $statusWidth - (Get-DisplayWidth $status))
            Write-Host ("`r" + $status + (' ' * $pad)) -NoNewline
            $statusWidth = Get-DisplayWidth $status
        }
    } finally {
        if ($keyInput) { try { [Console]::TreatControlCAsInput = $oldCtrlC } catch { } }
        Write-Host ''
        # -Repeat では Ctrl+C で止めるのが通常の終わり方なので、止めたときも保存する
        if ($runDir) {
            try {
                $header = Get-HeaderLines $results $roundTime $round
                $table  = Get-ResultTable $results
                Save-Report $header $table $results $changeLines
            } catch {
                $saveError = Get-ErrorText (Get-InnerException $_.Exception)
            }
            # 入力をリダイレクトしていると、Ctrl+C でここから先に進まずに終わるため、保存先はここで表示する
            if (-not $saveError) {
                Write-Host ''
                Write-Host ('保存先: {0}' -f $runDir)
                $pathShown = $true
            }
        }
    }
}

#endregion

if ($saveError) {
    Write-Host ''
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($runDir -and -not $pathShown) {
    Write-Host ''
    Write-Host ('保存先: {0}' -f $runDir)
} elseif ($saveWarning) {
    Write-Host ''
    Write-Warning $saveWarning
}

if (@($results | Where-Object { -not $_.Ok }).Count -gt 0) { exit 2 }
exit 0
