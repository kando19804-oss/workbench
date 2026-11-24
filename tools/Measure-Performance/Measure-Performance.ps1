<#
.SYNOPSIS
    CPU・メモリ・ディスク・ネットワークの使われ方と、重いプロセスを一定の間隔で記録します。

.DESCRIPTION
    このコンピューターの CPU・メモリ・ディスク・ネットワークの使われ方と、CPU やメモリを多く使っているプロセスを、
    指定した間隔で CSV（BOM 付き UTF-8）に記録します。止めると、期間中の平均と最大、目安を超えた時間帯、
    重かったプロセスをまとめて表示します。

    性能カウンターは英語の名前で読むため、Windows の表示の言語によらず動きます。
    設定を変更したり、プロセスを止めたりすることはありません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Measure-Performance.ps1

.PARAMETER Interval
    記録する間隔（秒。1～3600）。既定値は 5 です。

.PARAMETER Duration
    記録する時間（分。1～10080）。時間がたつと自動で止まります。
    省略すると、Ctrl+C で止めるまで記録します。

.PARAMETER Top
    毎回記録するプロセスの数（0～50）。CPU を多く使っている上位と、メモリを多く使っている上位を、
    それぞれこの件数まで記録します。既定値は 5 です。0 にすると、プロセスを毎回は記録しません。

.PARAMETER Process
    上位に入らなくても毎回記録するプロセスの名前（w3wp か w3wp.exe の形）。* を使えます。
    * 以外の文字は、そのままの文字として扱います。大文字と小文字は区別しません。
    複数指定するときは、カンマで区切ります。

.PARAMETER Counter
    一緒に記録するカウンター。\Web Service(_Total)\Current Connections のように英語の名前で書きます
    （この PC のパフォーマンス モニターに表示される名前でもかまいません）。インスタンスに * を使えます。
    複数指定するときは、カンマで区切ります（かっこの中のカンマは区切りにしません）。

.PARAMETER OutputFolder
    記録の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    記録を保存せず、画面表示だけにします。

.PARAMETER ConfigFile
    設定ファイルのパス。Interval～NoSave の値を「項目 = 値」の形で書いておけます。
    コマンドラインで指定した値は、設定ファイルより優先します。書き方は README を参照してください。

.PARAMETER _UnknownArgument
    指定しません。どのパラメーターにも当たらない引数（パラメーターの名前の書き間違いなど）を受け取り、
    案内を出して止まるためのものです。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Measure-Performance.ps1

    Ctrl+C で止めるまで、5 秒ごとに記録します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Measure-Performance.ps1 -Interval 30 -Duration 480

    30 秒ごとに 8 時間記録します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Measure-Performance.ps1 -Process "w3wp,java"

    w3wp.exe と java.exe は、上位に入らなくても毎回記録します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Measure-Performance.ps1 -Counter "\Web Service(_Total)\Current Connections"

    IIS の接続数も一緒に記録します。

.NOTES
    Measure-Performance 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [ValidateRange(1, 3600)]
    [int]$Interval = 5,

    [ValidateRange(1, 10080)]
    [int]$Duration,

    [ValidateRange(0, 50)]
    [int]$Top = 5,

    [string[]]$Process = @(),

    [string[]]$Counter = @(),

    [string]$OutputFolder,

    [switch]$NoSave,

    [string]$ConfigFile,

    # どのパラメーターにも当たらない引数を受け取る（案内を出して止めるため。指定しない）
    [Parameter(ValueFromRemainingArguments = $true, DontShow = $true)]
    [string[]]$_UnknownArgument
)

# 制約付き言語モード（AppLocker や WDAC の環境）では、このツールが使う .NET の型も Add-Type も使えないので、最初に止める
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host '制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）'
    exit 1
}

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion  = '1.0.0'
$ToolName       = 'Measure-Performance'
$ToolTitle      = '性能の記録'
$MaxCsvRows     = 1000000   # 1 つの CSV に書く行の上限（Excel で開ける行数は 1,048,576 行）
$SummaryRows    = 20        # まとめの「指定したプロセス」「追加のカウンター」に出す行数
$NameWidth      = 24        # まとめの表のプロセスの名前の幅（半角の文字数）。超える分は … にする
$LiveNameWidth  = 20        # 記録中の行のプロセスの名前の幅
$ServiceWidth   = 30        # まとめの表のサービスの幅
$CounterWidth   = 50        # まとめの表のカウンターの幅
$DriveWidth     = 6         # 記録中の行のドライブの幅
$CpuLimit       = 90.0      # 目安: CPU 使用率(%) がこれ以上
$QueuePerCpu    = 2         # 目安: プロセッサ待ち行列が、論理プロセッサの数のこの倍以上
$MemoryLimit    = 90.0      # 目安: メモリ使用率(%) がこれ以上
$CommitPctLimit = 90.0      # 目安: コミット率(%) がこれ以上
$LatencyLimit   = 20.0      # 目安: ディスクの応答時間(ms) がこれ以上
$LatencyMinIops = 10.0      # 応答時間は、読み書きが 1 秒にこの回数以上あったときだけ判定する
$FreeLimit      = 10.0      # 目安: ディスクの空き(%) がこれ未満
$NetLimit       = 80.0      # 目安: ネットワークの使用率(%) がこれ以上
$GapFactor      = 2         # 前の回から、間隔のこの倍を超えて間があいたら、その回は記録しない（スリープから戻ったときなど）
$AdapterRefreshSeconds = 300   # 名前の分からないアダプターが出たとき、接続の名前を調べ直す間隔（秒）
$TimeFormat     = 'yyyy/MM/dd HH:mm:ss'
$Invariant      = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom        = New-Object Text.UTF8Encoding $true
$Utf8NoBom      = New-Object Text.UTF8Encoding $false
$machineName    = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

$SystemColumns  = '日時', 'CPU 使用率(%)', '最も忙しいコア(%)', 'プロセッサ待ち行列', 'メモリ使用率(%)', '空きメモリ(MB)', 'コミット(MB)', 'コミットの上限(MB)', 'コミット率(%)', 'ページ読み込み(回/秒)', 'プロセス数'
$DiskColumns    = '日時', 'ドライブ', '使用時間(%)', '応答時間(ms)', '読み取りの応答時間(ms)', '書き込みの応答時間(ms)', '回数(回/秒)', '読み取り(MB/秒)', '書き込み(MB/秒)', '待ち行列', '空き(GB)', '空き(%)'
$NetColumns     = '日時', 'アダプター', '説明', '受信(Mbps)', '送信(Mbps)', '使用率(%)', '速度(Mbps)'
$ProcessColumns = '日時', 'プロセス', 'PID', 'CPU(%)', 'CPU の順位', 'メモリ(MB)', 'メモリの順位', 'コミット(MB)', 'I/O(MB/秒)', 'ハンドル数', 'スレッド数', '指定', 'サービス'
$ListColumns    = 'プロセス', 'PID', 'サービス', '実行ユーザー', '起動日時', '実行ファイル', 'コマンドライン', '最初に見た日時', '最後に見た日時', 'CPU の平均(%)', 'CPU の最大(%)', 'メモリの最大(MB)', 'コミットの最大(MB)', 'コミットの増減(MB)'
$CounterColumns = '日時', 'カウンター', 'インスタンス', '値'

# 読むカウンター（英語の名前）。CPU とメモリが読めなければ止まる。ディスクとネットワークは、読めなければ飛ばす
$CoreCounters = [ordered]@{
    CpuTotal    = '\Processor Information(_Total)\% Processor Time'
    CpuCores    = '\Processor Information(*)\% Processor Time'
    Queue       = '\System\Processor Queue Length'
    Processes   = '\System\Processes'
    Available   = '\Memory\Available Bytes'
    Committed   = '\Memory\Committed Bytes'
    CommitLimit = '\Memory\Commit Limit'
    CommitPct   = '\Memory\% Committed Bytes In Use'
    PageReads   = '\Memory\Page Reads/sec'
}
$DiskCounters = [ordered]@{
    Idle     = '\LogicalDisk(*)\% Idle Time'
    Latency  = '\LogicalDisk(*)\Avg. Disk sec/Transfer'
    LatRead  = '\LogicalDisk(*)\Avg. Disk sec/Read'
    LatWrite = '\LogicalDisk(*)\Avg. Disk sec/Write'
    Iops     = '\LogicalDisk(*)\Disk Transfers/sec'
    ReadB    = '\LogicalDisk(*)\Disk Read Bytes/sec'
    WriteB   = '\LogicalDisk(*)\Disk Write Bytes/sec'
    Queue    = '\LogicalDisk(*)\Current Disk Queue Length'
    FreeMB   = '\LogicalDisk(*)\Free Megabytes'
    FreePct  = '\LogicalDisk(*)\% Free Space'
}
$NetCounters = [ordered]@{
    Received  = '\Network Interface(*)\Bytes Received/sec'
    Sent      = '\Network Interface(*)\Bytes Sent/sec'
    Bandwidth = '\Network Interface(*)\Current Bandwidth'
}

#region 補助関数

# コマンドプロンプトから渡された引数に残る引用符を外す
# （"C:\" が C:" として届く問題や、cmd では ' がそのまま届く問題への対策）
function Remove-Quote([string]$Text) {
    $t = $Text.Trim().Trim('"')
    if ($t.Length -ge 2 -and $t.StartsWith("'") -and $t.EndsWith("'")) { $t = $t.Substring(1, $t.Length - 2) }
    return $t
}

# 5.1 は 260 文字を超えるパスを通常の書き方では読めないため、内部では \\?\ 付きのパスを使う
function ConvertTo-LongPath([string]$FullPath) {
    if ($FullPath.StartsWith('\\?\')) { return $FullPath }
    if ($FullPath.StartsWith('\\')) { return '\\?\UNC\' + $FullPath.Substring(2) }
    return '\\?\' + $FullPath
}

# 表示・保存用に \\?\ を外す
function ConvertTo-DisplayPath([string]$FullPath) {
    if ($FullPath.StartsWith('\\?\UNC\')) { return '\\' + $FullPath.Substring(8) }
    if ($FullPath.StartsWith('\\?\')) { return $FullPath.Substring(4) }
    return $FullPath
}

# 日付や数値の区切り文字が、OS の地域設定によって変わらないようにする
function Format-Invariant([string]$Format) {
    return [string]::Format($Invariant, $Format, [object[]]$args)
}

function Format-Duration([TimeSpan]$Span) {
    if ($Span.TotalSeconds -lt 60) { return '{0}秒' -f [math]::Floor($Span.TotalSeconds) }
    if ($Span.TotalHours -lt 1) { return '{0}分{1}秒' -f $Span.Minutes, $Span.Seconds }
    return '{0}時間{1}分{2}秒' -f [math]::Floor($Span.TotalHours), $Span.Minutes, $Span.Seconds
}

# .NET のメソッドで起きた例外は MethodInvocationException に包まれて届くため、元の例外を取り出す
function Get-InnerException([Exception]$Exception) {
    while (($Exception -is [Management.Automation.MethodInvocationException] -or $Exception -is [Reflection.TargetInvocationException]) -and $Exception.InnerException) {
        $Exception = $Exception.InnerException
    }
    return $Exception
}

# 例外メッセージに含まれる \\?\ 付きのパスを、表示用のパスに戻す。
# .NET のメッセージは途中や末尾に改行が入ることがあるため、1 行にまとめる
function Get-ErrorText([Exception]$Exception) {
    $text = $Exception.Message.Replace('\\?\UNC\', '\\').Replace('\\?\', '')
    return ($text -replace '\s*[\r\n]+\s*', ' ').Trim()
}

# CSV の 1 行
function ConvertTo-CsvLine([object[]]$Fields) {
    $parts = [string[]]::new($Fields.Count)
    for ($i = 0; $i -lt $Fields.Count; $i++) { $parts[$i] = ([string]$Fields[$i]).Replace('"', '""') }
    return '"' + ($parts -join '","') + '"'
}

# Excel は、= + - @ で始まる値を数式に、数字と区切りの記号だけの値（2026-10-31、001）や「.1」「1e5」を
# 数値や日付に変えてしまうため、その形の値は ="..." で囲む。式に書けない 255 文字を超える値は、先頭に空白を足す
$excelRisk = New-Object Text.RegularExpressions.Regex ('^(?:[=+\-@]|[\d\s.,/:+\-]+$)')
function Protect-ExcelText([string]$Text) {
    if ($Text -eq '') { return $Text }
    $risky = $excelRisk.IsMatch($Text)
    if (-not $risky) {
        # 全角の数字も Excel は数値にするため、半角にそろえてから判定する。
        # 壊れた文字（対になっていないサロゲート）を含むと Normalize が例外を出すので、そのときはそのまま判定する
        $normalized = $Text
        try { $normalized = $Text.Normalize([Text.NormalizationForm]::FormKC) } catch { }
        $number = 0.0
        $risky = [double]::TryParse($normalized.TrimEnd('%'), [Globalization.NumberStyles]::Float, $Invariant, [ref]$number)
    }
    if (-not $risky) { return $Text }
    if ($Text.Length -le 255) { return '="' + $Text.Replace('"', '""') + '"' }
    return ' ' + $Text
}

# 値を 1 行の文字列にそろえる。途中の改行は空白 1 つに、ほかの制御文字は ? にする（画面や CSV を崩さないため）
function ConvertTo-CellText($Value) {
    if ($null -eq $Value) { return '' }
    $text = ([string]$Value) -replace '\s*[\r\n]+\s*', ' '
    return ($text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '?').Trim()
}

# 画面の表で列をそろえるため、全角文字を 2 文字分の幅として数える（半角カタカナは 1 文字分）。
# 表の値が半角だけのときは、1 文字ずつ数えない
$wideChar = New-Object Text.RegularExpressions.Regex ('[^\u0000-ÿ｡-ﾟ]')
function Get-DisplayWidth([string]$Text) {
    if (-not $wideChar.IsMatch($Text)) { return $Text.Length }
    $width = 0
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int]$ch
        if ($code -lt 0x100 -or ($code -ge 0xFF61 -and $code -le 0xFF9F)) { $width += 1 }
        elseif ($code -ge 0xDC00 -and $code -le 0xDFFF) { }   # サロゲートペアの後半は、前半で 2 文字分を数えた
        else { $width += 2 }
    }
    return $width
}

# 画面の表の値を、半角 $Width 文字分までにする（超える分は … にする）
function Limit-Width([string]$Text, [int]$Width) {
    if ((Get-DisplayWidth $Text) -le $Width) { return $Text }
    $result = ''
    foreach ($ch in $Text.ToCharArray()) {
        if ((Get-DisplayWidth ($result + $ch)) -gt $Width - 2) { break }
        $result += $ch
    }
    return $result + '…'
}

function Format-PadLeft([string]$Text, [int]$Width) {
    $w = Get-DisplayWidth $Text
    if ($w -ge $Width) { return $Text }
    return (' ' * ($Width - $w)) + $Text
}

function Format-PadRight([string]$Text, [int]$Width) {
    $w = Get-DisplayWidth $Text
    if ($w -ge $Width) { return $Text }
    return $Text + (' ' * ($Width - $w))
}

# 見出しと行（文字列の配列）から、列をそろえた表の行を作る。Groups を渡すと、見出しの上に
# 「受信(Mbps)」のような、いくつかの列をまとめた見出しの行を足す（要素は @(見出し, まとめる列の数)）
function Format-TextTable([string[]]$Head, [bool[]]$Right, $Rows, $Groups = $null) {
    $all = @(, $Head) + @($Rows)
    $columns = $Head.Count
    $widths = [int[]]::new($columns)
    $cellWidths = [object[]]::new($all.Count)
    for ($r = 0; $r -lt $all.Count; $r++) {
        $rowWidths = [int[]]::new($columns)
        for ($i = 0; $i -lt $columns; $i++) {
            $w = Get-DisplayWidth ([string]$all[$r][$i])
            $rowWidths[$i] = $w
            if ($w -gt $widths[$i]) { $widths[$i] = $w }
        }
        $cellWidths[$r] = $rowWidths
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    if ($Groups) {
        # まとめた見出しが、まとめる列の幅より広いときは、最後の列を広げる
        $col = 0
        foreach ($g in $Groups) {
            $total = 2 * ($g[1] - 1)
            for ($i = $col; $i -lt $col + $g[1]; $i++) { $total += $widths[$i] }
            $labelWidth = Get-DisplayWidth ([string]$g[0])
            if ($labelWidth -gt $total) { $widths[$col + $g[1] - 1] += $labelWidth - $total }
            $col += $g[1]
        }
        $parts = [System.Collections.Generic.List[string]]::new()
        $col = 0
        foreach ($g in $Groups) {
            $total = 2 * ($g[1] - 1)
            for ($i = $col; $i -lt $col + $g[1]; $i++) { $total += $widths[$i] }
            $parts.Add((Format-PadRight ([string]$g[0]) $total))
            $col += $g[1]
        }
        $lines.Add((' ' + ($parts -join '  ')).TrimEnd())
    }
    $cells = [string[]]::new($columns)
    for ($r = 0; $r -lt $all.Count; $r++) {
        for ($i = 0; $i -lt $columns; $i++) {
            $pad = ' ' * ($widths[$i] - $cellWidths[$r][$i])
            $cells[$i] = if ($Right[$i]) { $pad + [string]$all[$r][$i] } else { [string]$all[$r][$i] + $pad }
        }
        $lines.Add((' ' + ($cells -join '  ')).TrimEnd())
    }
    return , $lines.ToArray()
}

# 数値の書式。値がないときは空の文字列
function Format-Number($Value, [string]$Format) {
    if ($null -eq $Value) { return '' }
    return ([double]$Value).ToString($Format, $Invariant)
}

# 増減（+12、-4、0）
function Format-Signed([double]$Value) {
    $rounded = [math]::Round($Value)
    if ($rounded -gt 0) { return '+' + $rounded.ToString('N0', $Invariant) }
    if ($rounded -lt 0) { return '-' + (-$rounded).ToString('N0', $Invariant) }
    return '0'
}

# 速度（1 Gbps、2.5 Gbps、100 Mbps）
function Format-Speed([double]$BitsPerSecond) {
    if ($BitsPerSecond -ge 1e9) { return ($BitsPerSecond / 1e9).ToString('0.#', $Invariant) + ' Gbps' }
    if ($BitsPerSecond -ge 1e6) { return ($BitsPerSecond / 1e6).ToString('0.#', $Invariant) + ' Mbps' }
    return ($BitsPerSecond / 1e3).ToString('0.#', $Invariant) + ' Kbps'
}

# まとめの時刻。記録を始めた日と違う日なら、月日も付ける
function Format-When($Time) {
    if ($null -eq $Time) { return '' }
    if ($Time.Date -eq $recordStart.Date) { return $Time.ToString('HH:mm:ss', $Invariant) }
    return $Time.ToString('MM-dd HH:mm:ss', $Invariant)
}

function Stop-WithError([string]$Message) {
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value, [string]$Example = 'D:\perf') {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        ('値を引用符で囲むときは、最後に \ を付けないでください（例: {0} "{1}"）。' -f $Name, $Example) + '最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
}

# 引数の値に引用符（"）が残っていないかを、ほかの確認より先にまとめて確かめ、そのあと、どのパラメーターにも当たらない引数
# （-_UnknownArgument に集まる）があれば止める。コマンドプロンプトで "C:\data\" のように最後を \" にすると、引用符が閉じずに
# 後ろの引数までその値に入り、空白を含む値の残りが余る（-OutputFolder "D:\perf\" -Process "a b" なら b）。
# その前に、引用符の書き方を案内する。QuoteExamples は、パラメーターの名前 → 正しい書き方の例
function Assert-CommandArgument($Bound, $QuoteExamples) {
    foreach ($name in @($QuoteExamples.Keys)) {
        if (-not $Bound.ContainsKey($name)) { continue }
        foreach ($value in @($Bound[$name])) {
            foreach ($part in ([string]$value -split ',')) {
                if ((Remove-Quote $part).Trim("'").Contains('"')) { Stop-WithQuoteError ('-' + $name) (Remove-Quote ([string]$value)) $QuoteExamples[$name] }
            }
        }
    }
    if (-not $Bound.ContainsKey('_UnknownArgument')) { return }
    $shown = @(foreach ($arg in @($Bound['_UnknownArgument'])) { if ([string]$arg -match '^$|\s') { '"{0}"' -f $arg } else { [string]$arg } }) -join ' '
    Stop-WithError (('どのパラメーターにも当たらない引数があります: {0}' -f $shown) + "`r`n" +
        'パラメーターの名前の書き間違いや、空白を含む値の引用符（"）の付け忘れがないかを確かめてください。コマンドプロンプトで値の最後を \" にしたとき（"C:\data\" など）も、後ろの引数がずれてこうなります。')
}
Assert-CommandArgument $PSBoundParameters ([ordered]@{ OutputFolder = 'D:\perf'; Process = 'w3wp,java'; Counter = '\Web Service(_Total)\Current Connections'; ConfigFile = 'D:\tools\nightly.ini' })

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

# %TEMP% などの環境変数を展開する。名前を打ち間違えると %名前% のまま残り、その名前のフォルダに保存してしまうため止める。
# URL エンコードされたフォルダ名（%E3%81%82 など）を誤って止めないよう、16 進 2 桁の %XX% は対象にしない
function Expand-ConfigVariable([int]$LineNumber, [string]$Text) {
    $expanded = [Environment]::ExpandEnvironmentVariables($Text)
    if ($expanded -match '%(?![0-9A-Fa-f]{2}%)[A-Za-z_][^%\\=\s]*%') { Stop-WithConfigError $LineNumber ('定義されていない環境変数です: {0}' -f $Matches[0]) }
    return $expanded
}

# プロセスの名前の一覧を、カンマで区切って読む。各要素の前後の引用符と空白も外す
function Split-List([string]$Text) {
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($part in ($Text -split ',')) {
        $item = (Remove-Quote $part).Trim("'").Trim()
        if ($item) { $list.Add($item) }
    }
    return , $list
}

# カウンターの一覧を、かっこの外のカンマで区切って読む（\Processor Information(0,1)\... のように、インスタンスの名前にカンマが入るため）
function Split-CounterList([string]$Text) {
    $list = [System.Collections.Generic.List[string]]::new()
    $depth = 0
    $builder = New-Object Text.StringBuilder
    foreach ($ch in ($Text + ',').ToCharArray()) {
        if ($ch -eq '(') { $depth++ } elseif ($ch -eq ')' -and $depth -gt 0) { $depth-- }
        if ($ch -eq ',' -and $depth -eq 0) {
            $item = (Remove-Quote $builder.ToString()).Trim("'").Trim()
            if ($item) { $list.Add($item) }
            [void]$builder.Clear()
        } else {
            [void]$builder.Append($ch)
        }
    }
    return , $list
}

# 名前で並べるときの順を、5.1 と 7 でそろえるためのキー。Sort-Object の文字列の比べ方は 5.1（.NET Framework）と 7 で違い、
# ハイフンなどの記号の扱いが変わる。大文字にそろえた文字の番号を 16 進数で並べた文字列にすると、どちらでも文字の番号の順になる
$sortKeys = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
function ConvertTo-SortKey([string]$Text) {
    $key = $null
    if ($sortKeys.TryGetValue($Text, [ref]$key)) { return $key }
    $builder = New-Object Text.StringBuilder ($Text.Length * 4)
    foreach ($ch in $Text.ToUpperInvariant().ToCharArray()) { [void]$builder.Append(([int]$ch).ToString('X4', $Invariant)) }
    $key = $builder.ToString()
    $sortKeys[$Text] = $key
    return $key
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの。数値の範囲は param の ValidateRange と同じにする
$configItems   = 'Interval', 'Duration', 'Top', 'Process', 'Counter', 'OutputFolder', 'NoSave'
$configRanges  = @{ Interval = 1, 3600; Duration = 1, 10080; Top = 0, 50 }
$configLists   = 'Process', 'Counter'   # 複数の行に分けて書ける項目
$configPath    = $null
$configApplied = @{}   # 設定ファイルの値を使った項目 → 行番号

if ($ConfigFile) {
    $configText = Remove-Quote $ConfigFile
    if ($configText.Contains('"')) { Stop-WithQuoteError '-ConfigFile' $configText 'D:\tools\nightly.ini' }
    $configText = Resolve-StartPath $configText
    $configPath = $configText
    try {
        $configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((ConvertTo-DisplayPath $configText))
        $configPath = [IO.Path]::GetFullPath($configPath)
    } catch { }
    if (-not [IO.File]::Exists((ConvertTo-LongPath $configPath))) {
        # エクスプローラーで拡張子を表示していないと、メモ帳で nightly.ini.txt として保存してしまうことがある
        $message = '設定ファイルが見つかりません: {0}' -f $configPath
        if ([IO.File]::Exists((ConvertTo-LongPath ($configPath + '.txt')))) {
            $message += "`r`n" + ('拡張子が .txt になっていませんか: {0}.txt' -f $configPath)
        }
        Stop-WithError $message
    }
    try {
        $configLines = (Read-ConfigText (ConvertTo-LongPath $configPath)) -split '\r\n|\n|\r'
    } catch {
        Stop-WithError ('設定ファイルを読めません: {0}（{1}）' -f $configPath, (Get-ErrorText (Get-InnerException $_.Exception)))
    }

    $configDir    = Split-Path -LiteralPath $configPath   # -LiteralPath には -Parent を付けられない（親フォルダを返すのが既定）
    $configValues = @{}
    $configLineOf = @{}
    $listItems    = @{ Process = [System.Collections.Generic.List[string]]::new(); Counter = [System.Collections.Generic.List[string]]::new() }
    $lineNumber   = 0
    foreach ($rawLine in $configLines) {
        $lineNumber++
        $lineText = $rawLine.Trim()
        # # と ; は値に使える文字なので、行の先頭にあるときだけコメントとみなす
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

        # Process と Counter は長くなりやすいので複数の行に分けて書ける。ほかの項目は、どちらが効いているか分からなくなるため 1 回だけ
        if ($configLists -notcontains $itemName -and $configLineOf.ContainsKey($itemName)) {
            Stop-WithConfigError $lineNumber ('{0} が 2 回書かれています（{1} 行目と {2} 行目）' -f $itemName, $configLineOf[$itemName], $lineNumber)
        }
        if (-not $configLineOf.ContainsKey($itemName)) { $configLineOf[$itemName] = $lineNumber }

        if ($configLists -contains $itemName) {
            # "a,b" のように全体を囲む書き方と、"a","b" のように 1 つずつ囲む書き方のどちらも受け付ける
            $source = if ($itemValue.Contains('"')) { $rawValue } else { $itemValue }
            $parts = if ($itemName -eq 'Counter') { Split-CounterList $source } else { Split-List $source }
            foreach ($part in $parts) {
                $partText = Remove-ConfigQuote $part
                if ($partText.Contains('"')) { Stop-WithConfigError $lineNumber ('{0} の値に引用符（"）が残っています: {1}' -f $itemName, $partText) }
                if ($partText) { $listItems[$itemName].Add($partText) }
            }
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
    foreach ($itemName in $configLists) {
        if ($listItems[$itemName].Count -gt 0) { $configValues[$itemName] = $listItems[$itemName].ToArray() }
    }

    # コマンドラインで指定した項目は、設定ファイルより優先する
    foreach ($itemName in @($configValues.Keys)) {
        if (-not $PSBoundParameters.ContainsKey($itemName)) {
            Set-Variable -Name $itemName -Value $configValues[$itemName]
            $configApplied[$itemName] = $configLineOf[$itemName]
        }
    }
}

#endregion

#region 引数の確認

if ($NoSave -and $OutputFolder) { Stop-WithError '-NoSave のときは、-OutputFolder は指定できません' }

if ($Duration -and $Duration * 60 -lt $Interval) {
    Stop-WithError ('-Duration（{0} 分）が -Interval（{1} 秒）より短いため、1 回も記録できません' -f $Duration, $Interval)
}

# 入力をリダイレクトして起動すると、Ctrl+C をキーとして受け取れず、止めると PowerShell ごと終わる
$inputRedirected = $true
try { $inputRedirected = [Console]::IsInputRedirected } catch { }
if ($inputRedirected -and -not $Duration) {
    Stop-WithError ('Ctrl+C で止められない起動のしかたです。-Duration で記録する時間を指定してください' + "`r`n" +
        '（入力をリダイレクトして起動したときは、Ctrl+C をキーとして受け取れないためです）')
}

$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    if ($outText -match '^(\\\\\?\\)?[A-Za-z]:$') { $outText += '\' }
    $outText = Resolve-StartPath $outText
}

# -Process。.exe を付けて書いたときは名前全体と、付けずに書いたときは .exe を除いた名前と比べる。
# ワイルドカードは * だけにして、それ以外の文字は文字どおりに扱う。5.1 の WildcardPattern.Escape は
# バッククォートをエスケープしないため、自分でエスケープしてから照合用のパターンを作る
$wildcardOptions = [Management.Automation.WildcardOptions]'IgnoreCase, CultureInvariant'
$processNames    = [System.Collections.Generic.List[string]]::new()
$processPatterns = [System.Collections.Generic.List[Management.Automation.WildcardPattern]]::new()
$processFullName = [System.Collections.Generic.List[bool]]::new()
foreach ($item in @($Process)) {
    foreach ($name in (Split-List ([string]$item))) {
        if ($name.Contains('"')) { Stop-WithQuoteError '-Process' $name 'w3wp,java' }
        if ($processNames -contains $name) { continue }
        $escaped = $name.Replace('`', '``').Replace('[', '`[').Replace(']', '`]').Replace('?', '`?')
        $processNames.Add($name)
        $processPatterns.Add([Management.Automation.WildcardPattern]::new($escaped, $wildcardOptions))
        $processFullName.Add($name.EndsWith('.exe', [StringComparison]::OrdinalIgnoreCase))
    }
}

# -Counter。\オブジェクト(インスタンス)\カウンター の形。インスタンスの名前には \ が入ることがある（\Paging File(\??\C:\pagefile.sys)\% Usage）
$counterPaths = [System.Collections.Generic.List[string]]::new()
foreach ($item in @($Counter)) {
    foreach ($path in (Split-CounterList ([string]$item))) {
        if ($path.Contains('"')) { Stop-WithQuoteError '-Counter' $path '\Web Service(_Total)\Current Connections' }
        if ($path.StartsWith('\\')) { Stop-WithError ('-Counter には、ほかのコンピューターのカウンターは書けません: {0}' -f $path) }
        if ($path -notmatch '^\\[^\\].*\\[^\\]+$') {
            Stop-WithError ('-Counter の書き方が違います: {0}（\オブジェクト(インスタンス)\カウンター の形で書いてください）' -f $path)
        }
        if (-not $counterPaths.Contains($path)) { $counterPaths.Add($path) }
    }
}

$startTime = Get-Date
$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

#endregion

#region Windows の API

$nativeSource = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Text;

namespace Workbench.MeasurePerformance
{
    public static class SystemInfo
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct MEMORYSTATUSEX
        {
            public uint Length; public uint MemoryLoad; public ulong TotalPhys; public ulong AvailPhys;
            public ulong TotalPageFile; public ulong AvailPageFile; public ulong TotalVirtual; public ulong AvailVirtual; public ulong AvailExtendedVirtual;
        }

        [DllImport("kernel32.dll")]
        private static extern uint GetActiveProcessorCount(ushort groupNumber);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GlobalMemoryStatusEx(ref MEMORYSTATUSEX status);

        // プロセッサ グループ（論理プロセッサが 64 を超える PC）をまたいだ数。Environment.ProcessorCount は、今のグループの数しか返さないことがある
        public static int ProcessorCount()
        {
            uint count = 0;
            try { count = GetActiveProcessorCount(0xFFFF); } catch (EntryPointNotFoundException) { }
            return count > 0 ? (int)count : Environment.ProcessorCount;
        }

        // 物理メモリの全体。Hyper-V の動的メモリの仮想マシンでは、その時点で割り当てられている量になる
        public static double TotalPhysicalMemory()
        {
            MEMORYSTATUSEX status = new MEMORYSTATUSEX();
            status.Length = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX));
            return GlobalMemoryStatusEx(ref status) ? (double)status.TotalPhys : 0;
        }
    }

    public sealed class CounterData
    {
        public uint Status;
        public string[] Names = new string[0];
        public double[] Values = new double[0];
    }

    // 性能カウンター（PDH）。カウンターは英語の名前で足す（PdhAddEnglishCounter）。表示の言語によらず同じ名前で読める
    public sealed class PdhQuery
    {
        private const uint PdhMoreData = 0x800007D2;
        private const uint FormatDouble = 0x00000200 | 0x00008000;   // PDH_FMT_DOUBLE | PDH_FMT_NOCAP100

        [StructLayout(LayoutKind.Sequential)]
        private struct FMT_VALUE { public uint CStatus; public double Value; }

        [StructLayout(LayoutKind.Sequential)]
        private struct FMT_ITEM { public IntPtr Name; public FMT_VALUE Value; }

        [StructLayout(LayoutKind.Sequential)]
        private struct RAW_COUNTER { public uint CStatus; public uint TimeLow; public uint TimeHigh; public long FirstValue; public long SecondValue; public uint MultiCount; }

        [DllImport("pdh.dll", CharSet = CharSet.Unicode)]
        private static extern uint PdhOpenQueryW(string dataSource, IntPtr userData, out IntPtr query);

        [DllImport("pdh.dll", CharSet = CharSet.Unicode)]
        private static extern uint PdhAddEnglishCounterW(IntPtr query, string path, IntPtr userData, out IntPtr counter);

        [DllImport("pdh.dll", CharSet = CharSet.Unicode)]
        private static extern uint PdhAddCounterW(IntPtr query, string path, IntPtr userData, out IntPtr counter);

        [DllImport("pdh.dll")]
        private static extern uint PdhCollectQueryData(IntPtr query);

        [DllImport("pdh.dll", CharSet = CharSet.Unicode)]
        private static extern uint PdhGetFormattedCounterArrayW(IntPtr counter, uint format, ref uint bufferSize, out uint itemCount, IntPtr items);

        [DllImport("pdh.dll")]
        private static extern uint PdhGetRawCounterValue(IntPtr counter, out uint counterType, out RAW_COUNTER value);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr GetModuleHandleW(string moduleName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern int FormatMessageW(uint flags, IntPtr source, uint messageId, uint languageId, StringBuilder buffer, int size, IntPtr arguments);

        private IntPtr query;

        public PdhQuery()
        {
            uint result = PdhOpenQueryW(null, IntPtr.Zero, out query);
            if (result != 0) { throw new InvalidOperationException(ErrorText(result)); }
        }

        // 英語の名前で見つからないときは、AllowLocalized なら表示の言語の名前として探す
        public uint Add(string path, bool allowLocalized, out IntPtr counter)
        {
            uint result = PdhAddEnglishCounterW(query, path, IntPtr.Zero, out counter);
            if (result != 0 && allowLocalized)
            {
                IntPtr localized;
                if (PdhAddCounterW(query, path, IntPtr.Zero, out localized) == 0) { counter = localized; return 0; }
            }
            return result;
        }

        public uint Collect()
        {
            return PdhCollectQueryData(query);
        }

        // 1 回読み取った後の、カウンターの状態（0 と 1 は値がある）
        public uint RawStatus(IntPtr counter)
        {
            uint counterType;
            RAW_COUNTER raw;
            uint result = PdhGetRawCounterValue(counter, out counterType, out raw);
            return result != 0 ? result : raw.CStatus;
        }

        // インスタンスごとの値。値のないインスタンス（読み取りの間に増えたものなど）は含めない
        public CounterData Read(IntPtr counter)
        {
            CounterData data = new CounterData();
            uint size = 0;
            uint count;
            uint result = PdhGetFormattedCounterArrayW(counter, FormatDouble, ref size, out count, IntPtr.Zero);
            if (result != PdhMoreData) { data.Status = result == 0 ? 0x800007D5 : result; return data; }
            IntPtr buffer = Marshal.AllocHGlobal((int)size);
            try
            {
                result = PdhGetFormattedCounterArrayW(counter, FormatDouble, ref size, out count, buffer);
                data.Status = result;
                if (result != 0) { return data; }
                List<string> names = new List<string>();
                List<double> values = new List<double>();
                int itemSize = Marshal.SizeOf(typeof(FMT_ITEM));
                for (int i = 0; i < count; i++)
                {
                    FMT_ITEM item = (FMT_ITEM)Marshal.PtrToStructure(new IntPtr(buffer.ToInt64() + (long)i * itemSize), typeof(FMT_ITEM));
                    if (item.Value.CStatus != 0 && item.Value.CStatus != 1) { continue; }
                    if (double.IsNaN(item.Value.Value) || double.IsInfinity(item.Value.Value)) { continue; }
                    string name = item.Name == IntPtr.Zero ? "" : Marshal.PtrToStringUni(item.Name);
                    names.Add(name ?? "");
                    values.Add(item.Value.Value);
                }
                data.Names = names.ToArray();
                data.Values = values.ToArray();
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
            return data;
        }

        public static string ErrorText(uint code)
        {
            StringBuilder buffer = new StringBuilder(1024);
            IntPtr module = GetModuleHandleW("pdh.dll");
            // FORMAT_MESSAGE_FROM_HMODULE | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS
            int length = FormatMessageW(0x00000800 | 0x00001000 | 0x00000200, module, code, 0, buffer, buffer.Capacity, IntPtr.Zero);
            string text = length > 0 ? buffer.ToString().Trim() : "";
            string hex = "0x" + code.ToString("X8");
            return text.Length > 0 ? text + " (" + hex + ")" : hex;
        }
    }

    public sealed class ProcStat
    {
        public string Key;
        public string Name;
        public int Pid;
        public DateTime Started;      // 分からないときは DateTime.MinValue
        public string Services = "";
        public string User = "";
        public string Path = "";
        public string CommandLine = "";
        public DateTime FirstSeen;
        public DateTime LastSeen;
        public double CpuSum;         // 記録した回の CPU(%) の合計。期間全体の平均は、記録した回の数で割る
        public double CpuMax;
        public double MemoryMax;
        public double CommitMax;
        public double CommitFirst;
        public double CommitLast;
        public int Samples;
        public bool Specified;
        internal List<string> ServiceList = new List<string>();
    }

    public sealed class ProcRow
    {
        public ProcStat Stat;
        public double Cpu;
        public double Memory;
        public double Commit;
        public double Io;
        public long Handles;
        public long Threads;
        public int CpuRank;
        public int MemoryRank;
    }

    // プロセスの値は、Windows のプロセスの一覧（NtQuerySystemInformation）から取る。性能カウンターの Process と同じ値で、
    // PID と起動時刻で見分けられる（カウンターは svchost#3 のような名前で、プロセスが終わると番号がずれる）。
    // 一覧は、管理者でなくても、すべてのプロセスの分を読める
    public sealed class ProcessSampler
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct SYSTEM_PROCESS_INFORMATION
        {
            public uint NextEntryOffset; public uint NumberOfThreads; public long WorkingSetPrivateSize;
            public uint HardFaultCount; public uint NumberOfThreadsHighWatermark; public ulong CycleTime;
            public long CreateTime; public long UserTime; public long KernelTime;
            public ushort NameLength; public ushort NameMaximumLength; public IntPtr NameBuffer;
            public int BasePriority; public IntPtr UniqueProcessId; public IntPtr InheritedFromUniqueProcessId;
            public uint HandleCount; public uint SessionId; public UIntPtr UniqueProcessKey;
            public UIntPtr PeakVirtualSize; public UIntPtr VirtualSize; public uint PageFaultCount;
            public UIntPtr PeakWorkingSetSize; public UIntPtr WorkingSetSize;
            public UIntPtr QuotaPeakPagedPoolUsage; public UIntPtr QuotaPagedPoolUsage;
            public UIntPtr QuotaPeakNonPagedPoolUsage; public UIntPtr QuotaNonPagedPoolUsage;
            public UIntPtr PagefileUsage; public UIntPtr PeakPagefileUsage; public UIntPtr PrivatePageCount;
            public long ReadOperationCount; public long WriteOperationCount; public long OtherOperationCount;
            public long ReadTransferCount; public long WriteTransferCount; public long OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ENUM_SERVICE_STATUS_PROCESS
        {
            public IntPtr ServiceName; public IntPtr DisplayName;
            public uint ServiceType; public uint CurrentState; public uint ControlsAccepted; public uint Win32ExitCode;
            public uint ServiceSpecificExitCode; public uint CheckPoint; public uint WaitHint; public uint ProcessId; public uint ServiceFlags;
        }

        private struct Snap
        {
            public int Pid; public long Create; public long CpuTime; public long IoBytes; public string Name;
            public long WorkingSetPrivate; public ulong PrivateBytes; public uint Handles; public uint Threads;
        }

        [DllImport("ntdll.dll")]
        private static extern int NtQuerySystemInformation(int infoClass, IntPtr buffer, int length, out int returnLength);

        [DllImport("ntdll.dll")]
        private static extern int NtQueryInformationProcess(IntPtr process, int infoClass, IntPtr buffer, int length, out int returnLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint access, [MarshalAs(UnmanagedType.Bool)] bool inherit, int processId);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool QueryFullProcessImageNameW(IntPtr process, uint flags, StringBuilder name, ref uint size);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetTokenInformation(IntPtr token, int infoClass, IntPtr buffer, int length, out int returnLength);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool LookupAccountSidW(string systemName, IntPtr sid, StringBuilder name, ref uint nameLength, StringBuilder domain, ref uint domainLength, out int use);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ConvertSidToStringSidW(IntPtr sid, out IntPtr text);

        [DllImport("kernel32.dll")]
        private static extern IntPtr LocalFree(IntPtr memory);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr OpenSCManagerW(string machineName, string databaseName, uint access);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumServicesStatusExW(IntPtr manager, int infoLevel, uint serviceType, uint serviceState, IntPtr buffer, uint size, out uint needed, out uint returned, ref uint resumeHandle, string groupName);

        [DllImport("advapi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseServiceHandle(IntPtr handle);

        private const int StatusInfoLengthMismatch = unchecked((int)0xC0000004);
        private const uint ProcessQueryLimitedInformation = 0x1000;

        private readonly int ownPid;
        private readonly int cpuCount;
        private readonly WildcardPattern[] patterns;
        private readonly bool[] fullName;
        private IntPtr buffer = IntPtr.Zero;
        private int bufferSize;
        private Dictionary<string, Snap> previous;
        private long previousTimestamp;
        private readonly Dictionary<string, ProcStat> stats = new Dictionary<string, ProcStat>();
        private readonly Dictionary<string, string> accounts = new Dictionary<string, string>();
        private DateTime serviceRefreshed = DateTime.MinValue;

        public readonly List<ProcStat> StatList = new List<ProcStat>();
        public readonly ProcStat Self = new ProcStat();
        public readonly bool[] PatternMatched;
        public int Recorded;

        public ProcessSampler(int ownPid, int cpuCount, WildcardPattern[] patterns, bool[] fullName)
        {
            this.ownPid = ownPid;
            this.cpuCount = Math.Max(1, cpuCount);
            this.patterns = patterns ?? new WildcardPattern[0];
            this.fullName = fullName ?? new bool[0];
            PatternMatched = new bool[this.patterns.Length];
        }

        // 平均の起点にする一覧を取る
        public void Baseline()
        {
            previous = TakeSnapshot();
            previousTimestamp = Stopwatch.GetTimestamp();
        }

        // 前の回からの値を計算する。Record が false のとき（間があいた回）は、起点だけを取り直す。
        // 返す行は、このツール自身とアイドルを除いたすべてのプロセスで、CPU の順位の順
        public ProcRow[] Sample(DateTime now, bool record)
        {
            Dictionary<string, Snap> current = TakeSnapshot();
            long timestamp = Stopwatch.GetTimestamp();
            double seconds = (timestamp - previousTimestamp) / (double)Stopwatch.Frequency;
            Dictionary<string, Snap> before = previous;
            previous = current;
            previousTimestamp = timestamp;
            if (!record || before == null) { return new ProcRow[0]; }

            List<ProcRow> rows = new List<ProcRow>();
            bool found = false;
            foreach (KeyValuePair<string, Snap> pair in current)
            {
                Snap snap = pair.Value;
                if (snap.Pid == 0) { continue; }
                Snap old;
                bool existed = before.TryGetValue(pair.Key, out old);
                // 前の回の後に起動したプロセスは、起動してからの値がそのまま、この間に使った分になる
                long cpuTime = existed ? snap.CpuTime - old.CpuTime : snap.CpuTime;
                long ioBytes = existed ? snap.IoBytes - old.IoBytes : snap.IoBytes;
                double cpu = seconds > 0 ? Math.Max(0, cpuTime) / 1e7 / seconds / cpuCount * 100 : 0;
                if (cpu > 100) { cpu = 100; }
                double io = seconds > 0 ? Math.Max(0, ioBytes) / seconds / 1048576.0 : 0;
                double memory = snap.WorkingSetPrivate / 1048576.0;
                double commit = snap.PrivateBytes / 1048576.0;

                ProcStat stat;
                if (snap.Pid == ownPid)
                {
                    stat = Self;
                }
                else if (!stats.TryGetValue(pair.Key, out stat))
                {
                    stat = NewStat(pair.Key, snap, now);
                    found = true;
                }
                if (stat.Samples == 0) { stat.CommitFirst = commit; stat.FirstSeen = now; }
                stat.Samples++;
                stat.CpuSum += cpu;
                if (cpu > stat.CpuMax) { stat.CpuMax = cpu; }
                if (memory > stat.MemoryMax) { stat.MemoryMax = memory; }
                if (commit > stat.CommitMax) { stat.CommitMax = commit; }
                stat.CommitLast = commit;
                stat.LastSeen = now;
                if (snap.Pid == ownPid) { continue; }

                ProcRow row = new ProcRow();
                row.Stat = stat; row.Cpu = cpu; row.Memory = memory; row.Commit = commit; row.Io = io;
                row.Handles = snap.Handles; row.Threads = snap.Threads;
                rows.Add(row);
            }
            Recorded++;

            // サービスは、新しいプロセスが出たときと、1 分ごとに調べ直す（svchost.exe には、後からサービスが加わることがある）
            if (found || (now - serviceRefreshed).TotalSeconds >= 60) { RefreshServices(current, now); }

            rows.Sort(delegate (ProcRow a, ProcRow b) { int c = b.Memory.CompareTo(a.Memory); return c != 0 ? c : a.Stat.Pid.CompareTo(b.Stat.Pid); });
            for (int i = 0; i < rows.Count; i++) { rows[i].MemoryRank = i + 1; }
            rows.Sort(delegate (ProcRow a, ProcRow b) { int c = b.Cpu.CompareTo(a.Cpu); return c != 0 ? c : a.Stat.Pid.CompareTo(b.Stat.Pid); });
            for (int i = 0; i < rows.Count; i++) { rows[i].CpuRank = i + 1; }
            return rows.ToArray();
        }

        // 記録する行: CPU の上位、メモリの上位（それぞれ Top 件まで）と、-Process で指定したプロセス。CPU の順位の順
        public ProcRow[] Select(ProcRow[] rows, int top)
        {
            List<ProcRow> selected = new List<ProcRow>();
            foreach (ProcRow row in rows)
            {
                if (row.CpuRank <= top || row.MemoryRank <= top || row.Stat.Specified) { selected.Add(row); }
            }
            return selected.ToArray();
        }

        private ProcStat NewStat(string key, Snap snap, DateTime now)
        {
            ProcStat stat = new ProcStat();
            stat.Key = key;
            stat.Name = snap.Name;
            stat.Pid = snap.Pid;
            stat.Started = snap.Create > 0 ? DateTime.FromFileTime(snap.Create) : DateTime.MinValue;
            stat.FirstSeen = now;
            string shortName = snap.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) ? snap.Name.Substring(0, snap.Name.Length - 4) : snap.Name;
            for (int i = 0; i < patterns.Length; i++)
            {
                if (patterns[i].IsMatch(fullName[i] ? snap.Name : shortName)) { stat.Specified = true; PatternMatched[i] = true; }
            }
            ReadDetails(stat);
            stats[key] = stat;
            StatList.Add(stat);
            return stat;
        }

        // 実行ユーザー・実行ファイル・コマンドライン。管理者でないと、ほかのユーザーのプロセスは開けないので空のまま
        private void ReadDetails(ProcStat stat)
        {
            if (stat.Pid <= 4) { return; }
            IntPtr process = OpenProcess(ProcessQueryLimitedInformation, false, stat.Pid);
            if (process == IntPtr.Zero) { return; }
            try
            {
                StringBuilder path = new StringBuilder(32768);
                uint size = (uint)path.Capacity;
                if (QueryFullProcessImageNameW(process, 0, path, ref size)) { stat.Path = path.ToString(); }
                stat.User = ReadUser(process);
                stat.CommandLine = ReadCommandLine(process);
            }
            catch (Exception) { }
            finally
            {
                CloseHandle(process);
            }
        }

        private string ReadUser(IntPtr process)
        {
            IntPtr token;
            if (!OpenProcessToken(process, 0x0008, out token)) { return ""; }   // TOKEN_QUERY
            IntPtr info = IntPtr.Zero;
            try
            {
                int length;
                GetTokenInformation(token, 1, IntPtr.Zero, 0, out length);   // TokenUser
                if (length <= 0) { return ""; }
                info = Marshal.AllocHGlobal(length);
                if (!GetTokenInformation(token, 1, info, length, out length)) { return ""; }
                IntPtr sid = Marshal.ReadIntPtr(info);
                IntPtr sidText;
                if (!ConvertSidToStringSidW(sid, out sidText)) { return ""; }
                string sidString = Marshal.PtrToStringUni(sidText);
                LocalFree(sidText);
                string account;
                if (accounts.TryGetValue(sidString, out account)) { return account; }
                StringBuilder name = new StringBuilder(256);
                StringBuilder domain = new StringBuilder(256);
                uint nameLength = 256;
                uint domainLength = 256;
                int use;
                if (LookupAccountSidW(null, sid, name, ref nameLength, domain, ref domainLength, out use))
                {
                    account = domain.Length > 0 ? domain.ToString() + "\\" + name.ToString() : name.ToString();
                }
                else
                {
                    account = sidString;
                }
                accounts[sidString] = account;
                return account;
            }
            finally
            {
                if (info != IntPtr.Zero) { Marshal.FreeHGlobal(info); }
                CloseHandle(token);
            }
        }

        // ProcessCommandLineInformation（60）。Windows 8.1 / Server 2012 R2 から、属性の読み取りの権限だけで読める
        private static string ReadCommandLine(IntPtr process)
        {
            int length;
            int status = NtQueryInformationProcess(process, 60, IntPtr.Zero, 0, out length);
            if (length <= 0) { return ""; }
            IntPtr info = Marshal.AllocHGlobal(length);
            try
            {
                status = NtQueryInformationProcess(process, 60, info, length, out length);
                if (status != 0) { return ""; }
                int textLength = Marshal.ReadInt16(info) & 0xFFFF;
                IntPtr text = Marshal.ReadIntPtr(info, IntPtr.Size);   // UNICODE_STRING の Buffer
                if (text == IntPtr.Zero || textLength == 0) { return ""; }
                return Marshal.PtrToStringUni(text, textLength / 2);
            }
            finally
            {
                Marshal.FreeHGlobal(info);
            }
        }

        // PID → サービスの名前。管理者でなくても読める
        private void RefreshServices(Dictionary<string, Snap> current, DateTime now)
        {
            serviceRefreshed = now;
            Dictionary<int, List<string>> map = EnumServices();
            if (map == null) { return; }
            foreach (KeyValuePair<string, Snap> pair in current)
            {
                List<string> names;
                ProcStat stat;
                if (!map.TryGetValue(pair.Value.Pid, out names) || !stats.TryGetValue(pair.Key, out stat)) { continue; }
                bool changed = false;
                foreach (string name in names)
                {
                    if (!stat.ServiceList.Contains(name)) { stat.ServiceList.Add(name); changed = true; }
                }
                if (changed)
                {
                    stat.ServiceList.Sort(StringComparer.OrdinalIgnoreCase);
                    stat.Services = string.Join(", ", stat.ServiceList.ToArray());
                }
            }
        }

        private static Dictionary<int, List<string>> EnumServices()
        {
            IntPtr manager = OpenSCManagerW(null, null, 0x0004);   // SC_MANAGER_ENUMERATE_SERVICE
            if (manager == IntPtr.Zero) { return null; }
            Dictionary<int, List<string>> map = new Dictionary<int, List<string>>();
            try
            {
                uint resume = 0;
                int size = 256 * 1024;
                int itemSize = Marshal.SizeOf(typeof(ENUM_SERVICE_STATUS_PROCESS));
                while (true)
                {
                    IntPtr buffer = Marshal.AllocHGlobal(size);
                    try
                    {
                        uint needed;
                        uint returned;
                        // SC_ENUM_PROCESS_INFO、SERVICE_WIN32、SERVICE_ACTIVE
                        bool ok = EnumServicesStatusExW(manager, 0, 0x30, 1, buffer, (uint)size, out needed, out returned, ref resume, null);
                        int error = ok ? 0 : Marshal.GetLastWin32Error();
                        if (!ok && error != 234) { return map; }   // ERROR_MORE_DATA 以外は、取れた分だけ
                        for (int i = 0; i < returned; i++)
                        {
                            ENUM_SERVICE_STATUS_PROCESS item = (ENUM_SERVICE_STATUS_PROCESS)Marshal.PtrToStructure(new IntPtr(buffer.ToInt64() + (long)i * itemSize), typeof(ENUM_SERVICE_STATUS_PROCESS));
                            if (item.ProcessId == 0 || item.ServiceName == IntPtr.Zero) { continue; }
                            List<string> names;
                            if (!map.TryGetValue((int)item.ProcessId, out names)) { names = new List<string>(); map[(int)item.ProcessId] = names; }
                            names.Add(Marshal.PtrToStringUni(item.ServiceName));
                        }
                        if (ok) { break; }
                        if (returned == 0) { size = (int)Math.Max(needed + 4096, (uint)size * 2); }
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(buffer);
                    }
                }
            }
            finally
            {
                CloseServiceHandle(manager);
            }
            return map;
        }

        private Dictionary<string, Snap> TakeSnapshot()
        {
            if (buffer == IntPtr.Zero) { bufferSize = 1024 * 1024; buffer = Marshal.AllocHGlobal(bufferSize); }
            int status;
            while (true)
            {
                int needed;
                status = NtQuerySystemInformation(5, buffer, bufferSize, out needed);   // SystemProcessInformation
                if (status != StatusInfoLengthMismatch) { break; }
                Marshal.FreeHGlobal(buffer);
                buffer = IntPtr.Zero;
                bufferSize = Math.Max(bufferSize * 2, needed + 65536);
                buffer = Marshal.AllocHGlobal(bufferSize);
            }
            if (status != 0) { throw new InvalidOperationException("NtQuerySystemInformation: 0x" + status.ToString("X8")); }

            Dictionary<string, Snap> result = new Dictionary<string, Snap>(StringComparer.Ordinal);
            long offset = 0;
            while (true)
            {
                SYSTEM_PROCESS_INFORMATION info = (SYSTEM_PROCESS_INFORMATION)Marshal.PtrToStructure(new IntPtr(buffer.ToInt64() + offset), typeof(SYSTEM_PROCESS_INFORMATION));
                Snap snap = new Snap();
                snap.Pid = (int)info.UniqueProcessId.ToInt64();
                snap.Create = info.CreateTime;
                snap.CpuTime = info.UserTime + info.KernelTime;
                snap.IoBytes = info.ReadTransferCount + info.WriteTransferCount;
                snap.Name = info.NameBuffer == IntPtr.Zero ? (snap.Pid == 0 ? "System Idle Process" : "") : Marshal.PtrToStringUni(info.NameBuffer, info.NameLength / 2);
                snap.WorkingSetPrivate = info.WorkingSetPrivateSize;
                snap.PrivateBytes = info.PrivatePageCount.ToUInt64();
                snap.Handles = info.HandleCount;
                snap.Threads = info.NumberOfThreads;
                result[snap.Pid.ToString() + ":" + snap.Create.ToString()] = snap;
                if (info.NextEntryOffset == 0) { break; }
                offset += info.NextEntryOffset;
            }
            return result;
        }
    }
}
'@

# 同じ PowerShell でもう一度実行したときは、読み込み済みの型を使う（同じ名前の型は読み込み直せない）
if (-not ('Workbench.MeasurePerformance.ProcessSampler' -as [type])) {
    try {
        # 5.1 の Add-Type は、コンパイラーの警告もエラーとして扱う
        Add-Type -TypeDefinition $nativeSource -IgnoreWarnings -WarningAction SilentlyContinue -ErrorAction Stop
    } catch {
        Stop-WithError ('Windows の API を呼ぶ準備ができません。コンパイラーの実行が禁止されている環境では使えません（{0}）' -f (Get-ErrorText (Get-InnerException $_.Exception)))
    }
}

#endregion

#region この PC の情報とカウンターの準備

$cpuCount = [Workbench.MeasurePerformance.SystemInfo]::ProcessorCount()
$queueLimit = $cpuCount * $QueuePerCpu

$osName = ''
try { $osName = [string](Get-CimInstance -ClassName Win32_OperatingSystem -Property Caption -ErrorAction Stop).Caption } catch { }
if (-not $osName) {
    try { $osName = [string](Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name ProductName -ErrorAction Stop).ProductName } catch { }
}
$osName = ($osName -replace '^Microsoft\s+', '').Trim()

$cpuName = ''
try { $cpuName = [string](Get-ItemProperty -LiteralPath 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0' -Name ProcessorNameString -ErrorAction Stop).ProcessorNameString } catch { }
$cpuName = ($cpuName -replace '\s+', ' ').Trim()

$pdh = $null
try { $pdh = New-Object Workbench.MeasurePerformance.PdhQuery } catch {
    Stop-WithError ('性能カウンターを読めません（{0}）' -f (Get-ErrorText (Get-InnerException $_.Exception)))
}
$counterHint = '性能カウンターが無効になっているか、壊れています。パフォーマンス モニターで表示できるかを確かめてください'

# CPU とメモリ。足せなければ止まる
$coreHandles = @{}
foreach ($key in @($CoreCounters.Keys)) {
    $handle = [IntPtr]::Zero
    $result = $pdh.Add($CoreCounters[$key], $false, [ref]$handle)
    if ($result -ne 0) {
        Stop-WithError (('性能カウンターを読めません: {0}（{1}）' -f $CoreCounters[$key], [Workbench.MeasurePerformance.PdhQuery]::ErrorText($result)) + "`r`n" + $counterHint)
    }
    $coreHandles[$key] = $handle
}

# ディスクとネットワーク。足せなければ、その部分を記録しない
$groupNotes = [System.Collections.Generic.List[string]]::new()
function Add-CounterGroup($Specs, [string]$Label) {
    $handles = @{}
    foreach ($key in @($Specs.Keys)) {
        $handle = [IntPtr]::Zero
        $result = $pdh.Add($Specs[$key], $false, [ref]$handle)
        if ($result -ne 0) {
            $groupNotes.Add(('※ {0}のカウンターを読めないため、記録しません（{1}: {2}）' -f $Label, $Specs[$key], [Workbench.MeasurePerformance.PdhQuery]::ErrorText($result)))
            return $null
        }
        $handles[$key] = $handle
    }
    return $handles
}
$diskHandles = Add-CounterGroup $DiskCounters 'ディスク'
$netHandles  = Add-CounterGroup $NetCounters 'ネットワーク'

# 追加のカウンター（-Counter）。英語の名前で見つからなければ、表示の言語の名前として探す
$extraHandles = [System.Collections.Generic.List[object]]::new()
foreach ($path in $counterPaths) {
    $handle = [IntPtr]::Zero
    $result = $pdh.Add($path, $true, [ref]$handle)
    if ($result -ne 0) {
        Stop-WithError (('-Counter のカウンターが見つかりません: {0}（{1}）' -f $path, [Workbench.MeasurePerformance.PdhQuery]::ErrorText($result)) + "`r`n" +
            '英語の名前か、この PC のパフォーマンス モニターに表示される名前で書いてください')
    }
    $extraHandles.Add(@($path, $handle))
}

$sampler = New-Object Workbench.MeasurePerformance.ProcessSampler ($PID, $cpuCount, $processPatterns.ToArray(), $processFullName.ToArray())

# 1 回読み取って、カウンターが値を返すかを確かめる。インスタンスを指定した -Counter は、そのインスタンスがあるかも確かめる
[void]$pdh.Collect()
foreach ($key in 'CpuTotal', 'Available') {
    $status = $pdh.RawStatus($coreHandles[$key])
    if ($status -ne 0 -and $status -ne 1) {
        Stop-WithError (('性能カウンターを読めません: {0}（{1}）' -f $CoreCounters[$key], [Workbench.MeasurePerformance.PdhQuery]::ErrorText($status)) + "`r`n" + $counterHint)
    }
}
foreach ($extra in $extraHandles) {
    if ($extra[0].Contains('*')) { continue }
    $status = $pdh.RawStatus($extra[1])
    if ($status -ne 0 -and $status -ne 1) {
        Stop-WithError (('-Counter のカウンターが見つかりません: {0}（{1}）' -f $extra[0], [Workbench.MeasurePerformance.PdhQuery]::ErrorText($status)) + "`r`n" +
            '英語の名前か、この PC のパフォーマンス モニターに表示される名前で書いてください')
    }
}

# ネットワーク接続の名前（「イーサネット」など）。カウンターのインスタンスの名前は、アダプターの説明の ( ) を [ ] に、# / \ を _ に変えたもの
$adapterNames = @{}
$adapterRefreshed = [DateTime]::MinValue
function Update-AdapterName {
    $script:adapterRefreshed = [DateTime]::UtcNow
    try {
        foreach ($a in @(Get-CimInstance -ClassName Win32_NetworkAdapter -Property Name, NetConnectionID -ErrorAction Stop)) {
            if (-not $a.Name -or -not $a.NetConnectionID) { continue }
            $instance = ([string]$a.Name).Replace('(', '[').Replace(')', ']').Replace('#', '_').Replace('/', '_').Replace('\', '_')
            $adapterNames[$instance] = [string]$a.NetConnectionID
        }
    } catch { }
}
function Get-AdapterName([string]$Instance) {
    if (-not $adapterNames.ContainsKey($Instance) -and ([DateTime]::UtcNow - $adapterRefreshed).TotalSeconds -ge $AdapterRefreshSeconds) { Update-AdapterName }
    if ($adapterNames.ContainsKey($Instance)) { return $adapterNames[$Instance] }
    return $Instance
}

# 1 つの値のカウンター
function Get-SingleValue([IntPtr]$Handle) {
    $data = $pdh.Read($Handle)
    if ($data.Status -ne 0 -or $data.Values.Count -eq 0) { return $null }
    return $data.Values[0]
}

# 記録を始めるときのドライブとアダプター（空き容量と速度は、1 回の読み取りで値が出る）
$driveList = @()
if ($diskHandles) {
    $data = $pdh.Read($diskHandles['FreeMB'])
    $driveList = @($data.Names | Where-Object { $_ -ne '_Total' } | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } })
}
$adapterList = @()
if ($netHandles) {
    Update-AdapterName
    $data = $pdh.Read($netHandles['Bandwidth'])
    $adapterList = @(for ($i = 0; $i -lt $data.Names.Count; $i++) {
        $label = Get-AdapterName $data.Names[$i]
        if ($data.Values[$i] -gt 0) { $label += '（{0}）' -f (Format-Speed $data.Values[$i]) }
        $label
    })
}

#endregion

#region 保存先の準備

# 記録を始めてから書き込めないと分かるのを避けるため、先に保存先のフォルダと CSV の見出しを作っておく
$runDir = $null
$csvFiles = [ordered]@{}
function New-CsvFile([string]$Key, [string]$BaseName, [string[]]$Columns) {
    $file = @{ Key = $Key; Base = $BaseName; Columns = $Columns; Part = 1; Rows = 0; Name = ($BaseName + '.csv'); Path = (Join-Path $runDir ($BaseName + '.csv')) }
    [IO.File]::WriteAllText((ConvertTo-LongPath $file.Path), (ConvertTo-CsvLine $Columns) + "`r`n", $Utf8Bom)
    $csvFiles[$Key] = $file
}
if (-not $NoSave) {
    $baseDir = if ($outText) { $outText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        # \\?\ 付きのパスでは末尾のピリオドや空白が残るため、先に通常の書き方で整える（260 文字を超えると例外になるので、そのときはそのまま）
        try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
        $runName = Format-Invariant 'Performance_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        $suffix = 1
        while ([IO.Directory]::Exists((ConvertTo-LongPath $runDir)) -or [IO.File]::Exists((ConvertTo-LongPath $runDir))) {
            $suffix++
            $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix)
        }
        [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $runDir))
        New-CsvFile 'system' 'system' $SystemColumns
        New-CsvFile 'disks' 'disks' $DiskColumns
        New-CsvFile 'network' 'network' $NetColumns
        New-CsvFile 'processes' 'processes' $ProcessColumns
        if ($extraHandles.Count -gt 0) { New-CsvFile 'counters' 'counters' $CounterColumns }
    } catch {
        $ex = Get-InnerException $_.Exception
        $message = ('保存先に書き込めません: {0}（{1}）' -f $baseDir, (Get-ErrorText $ex)) + (Get-ConfigNote 'OutputFolder')
        # 管理者として開いた PowerShell からは、通常の権限で割り当てたネットワークドライブが見えない
        if ($isAdmin -and $ex -is [Management.Automation.DriveNotFoundException]) {
            $message += "`r`n" + '管理者として実行した PowerShell からは、通常の権限で割り当てたネットワークドライブは見えません。\\server\share の形で指定してください'
        }
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        Stop-WithError ($message + "`r`n" + '-NoSave を付けると、画面に表示だけします')
    }
}

# 記録の途中で書き込めなくなったときは、保存をやめて画面の表示だけで続ける
$saving       = [bool]$runDir
$saveError    = $null
$saveErrorAt  = $null
$splitNotes   = [System.Collections.Generic.List[string]]::new()

# CSV を次のファイル（processes_2.csv など）に分ける
function Start-CsvPart($File, [string]$Reason) {
    $File.Part++
    $File.Name = '{0}_{1}.csv' -f $File.Base, $File.Part
    $File.Path = Join-Path $runDir $File.Name
    $File.Rows = 0
    [IO.File]::WriteAllText((ConvertTo-LongPath $File.Path), (ConvertTo-CsvLine $File.Columns) + "`r`n", $Utf8Bom)
    $splitNotes.Add(('※ {0} に分けました（{1}）' -f $File.Name, $Reason))
}

# CSV に行を書き足す。Excel などで開かれていて書き込めない（共有違反）ときは、次のファイルに分けて続ける
function Add-CsvLine([string]$Key, $Lines) {
    if (-not $script:saving -or $Lines.Count -eq 0) { return }
    $file = $csvFiles[$Key]
    if ($null -eq $file) { return }
    try {
        if ($file.Rows + $Lines.Count -gt $MaxCsvRows) { Start-CsvPart $file ('行が {0:N0} 行に達したため' -f $MaxCsvRows) }
        $text = ($Lines -join "`r`n") + "`r`n"
        try {
            [IO.File]::AppendAllText((ConvertTo-LongPath $file.Path), $text, $Utf8NoBom)
        } catch {
            $ex = Get-InnerException $_.Exception
            # ERROR_SHARING_VIOLATION（32）・ERROR_LOCK_VIOLATION（33）
            if ($ex -isnot [IO.IOException] -or (($ex.HResult -band 0xFFFF) -ne 32 -and ($ex.HResult -band 0xFFFF) -ne 33)) { throw }
            $openedName = $file.Name
            Start-CsvPart $file ('{0} を開いているほかのプログラムがあり、書き足せなかったため' -f $openedName)
            Write-Host $splitNotes[$splitNotes.Count - 1] -ForegroundColor Yellow
            [IO.File]::AppendAllText((ConvertTo-LongPath $file.Path), $text, $Utf8NoBom)
        }
        $file.Rows += $Lines.Count
    } catch {
        $script:saving = $false
        $script:saveError = Get-ErrorText (Get-InnerException $_.Exception)
        $script:saveErrorAt = Get-Date
        Write-Warning ('記録を保存できなくなりました。画面の表示だけで続けます: {0}' -f $script:saveError)
    }
}

#endregion

#region 記録

# 平均・最大を集める
function New-Stat { return @{ Sum = 0.0; Count = 0; Max = $null; MaxTime = $null } }
function Add-Stat($Stat, $Value, $Time) {
    if ($null -eq $Value) { return }
    $Stat.Sum += [double]$Value
    $Stat.Count++
    if ($null -eq $Stat.Max -or $Value -gt $Stat.Max) { $Stat.Max = [double]$Value; $Stat.MaxTime = $Time }
}
function Get-Average($Stat) {
    if ($Stat.Count -eq 0) { return $null }
    return $Stat.Sum / $Stat.Count
}

$systemStats = [ordered]@{
    CpuTotal  = New-Stat; CpuCore = New-Stat; Queue = New-Stat; Memory = New-Stat; CommitPct = New-Stat; PageReads = New-Stat
}
$diskStats    = @{}   # ドライブ → 集計
$netStats     = @{}   # インスタンス → 集計
$counterStats = [ordered]@{}   # カウンターとインスタンス → 集計

# 目安を超えた回数と、続けて超えた最も長い区間
$limits = @{}
function Update-Limit([string]$Key, [string]$Label, [int]$Order, [string]$Name, [bool]$Hit, [datetime]$Time) {
    $item = $limits[$Key]
    if ($null -eq $item) {
        $item = @{ Label = $Label; Order = $Order; Sort = (ConvertTo-SortKey $Name); Count = 0; RunLength = 0; RunStart = $null; BestLength = 0; BestStart = $null; BestEnd = $null; FirstTime = $null; Touched = $false }
        $limits[$Key] = $item
    }
    $item.Touched = $true
    if (-not $Hit) { $item.RunLength = 0; return }
    $item.Count++
    if ($item.Count -eq 1) { $item.FirstTime = $Time }
    if ($item.RunLength -eq 0) { $item.RunStart = $Time }
    $item.RunLength++
    if ($item.RunLength -gt $item.BestLength) { $item.BestLength = $item.RunLength; $item.BestStart = $item.RunStart; $item.BestEnd = $Time }
}
# その回に判定しなかったもの（なくなったドライブなど）は、続いた区間を切る
function Close-LimitRound([bool]$Reset) {
    foreach ($item in $limits.Values) {
        if ($Reset -or -not $item.Touched) { $item.RunLength = 0 }
        $item.Touched = $false
    }
}

function Get-Mark([bool]$Hit) { if ($Hit) { return ' ※' } return '   ' }

# まとめの表に出しきれなかった件数。保存しないときは、見る CSV がないので件数だけ
function Get-MoreText([int]$Count, [string]$FileName) {
    if ($runDir -and $saving) { return Format-Invariant ' （ほか {0:N0} 件は {1} を参照してください）' $Count $FileName }
    return Format-Invariant ' （ほか {0:N0} 件）' $Count
}

# 記録中の行。列の幅は見出しに合わせる
$liveHead = ' ' + (Format-PadRight '時刻' 8) + '  ' + (Format-PadLeft 'CPU' 5) + '   ' + '  ' + (Format-PadLeft '最大のコア' 10) + '  ' +
    (Format-PadLeft 'メモリ' 6) + '   ' + '  ' + (Format-PadRight 'ディスク（最も忙しい）' 23) + '  ' + (Format-PadLeft 'ネットワーク' 13) + '   ' + '  ' + 'CPU の 1 位'

$recordStart   = $null
$recorded      = 0
$skippedRounds = 0
$gapRounds     = 0
$endReason     = $null
$unexpected    = $null
$keyInput      = $false
$oldCtrlC      = $false

# 記録を始めるときの表示
$durationText = ''
$headLines = [System.Collections.Generic.List[string]]::new()
$headLines.Add((Format-Invariant '{0} ({1} {2})' $ToolTitle $ToolName $ScriptVersion))
$adminText = if ($isAdmin) { 'はい' } else { 'いいえ' }
if ($osName) {
    $headLines.Add((Format-Invariant '対象: {0}（管理者: {1}）  {2}   {3:yyyy-MM-dd HH:mm:ss}' $machineName $adminText $osName $startTime))
} else {
    $headLines.Add((Format-Invariant '対象: {0}（管理者: {1}）   {2:yyyy-MM-dd HH:mm:ss}' $machineName $adminText $startTime))
}
if ($configPath) { $headLines.Add(('設定ファイル: {0}' -f $configPath)) }
$memoryTotal = [Workbench.MeasurePerformance.SystemInfo]::TotalPhysicalMemory()
$cpuLine = if ($cpuName) { 'CPU: {0}（論理プロセッサ {1}）' -f $cpuName, $cpuCount } else { 'CPU: 論理プロセッサ {0}' -f $cpuCount }
$headLines.Add((Format-Invariant '{0}   メモリ {1:N1} GB' $cpuLine ($memoryTotal / 1GB)))
function Join-Limited($Items, [int]$Limit) {
    $list = @($Items)
    if ($list.Count -eq 0) { return 'なし' }
    if ($list.Count -le $Limit) { return $list -join ', ' }
    return (($list | Select-Object -First $Limit) -join ', ') + (' ほか {0} 個' -f ($list.Count - $Limit))
}
$diskText = if ($diskHandles) { Join-Limited $driveList 10 } else { '記録しない' }
$netText  = if ($netHandles) { Join-Limited $adapterList 4 } else { '記録しない' }
$headLines.Add(('ディスク: {0}   ネットワーク: {1}' -f $diskText, $netText))
foreach ($note in $groupNotes) { $headLines.Add($note) }
$processText = @()
if ($Top -gt 0) { $processText += 'CPU とメモリの上位 {0} 件' -f $Top }
if ($processNames.Count -gt 0) { $processText += '指定 {0}' -f ($processNames -join ', ') }
if ($processText.Count -eq 0) { $processText = @('毎回は記録しない') }
$recordStart = Get-Date
if ($Duration) {
    $durationText = if ($Duration -lt 60) { '{0} 分間' -f $Duration } elseif ($Duration % 60 -eq 0) { '{0} 時間' -f ($Duration / 60) } else { '{0} 時間 {1} 分' -f [math]::Floor($Duration / 60), ($Duration % 60) }
    $endPlan = $recordStart.AddMinutes($Duration)
    $durationText += '（{0} まで）' -f $(if ($endPlan.Date -eq $recordStart.Date) { $endPlan.ToString('HH:mm:ss', $Invariant) } else { $endPlan.ToString('MM-dd HH:mm:ss', $Invariant) })
} else {
    $durationText = 'Ctrl+C で止めるまで'
}
$headLines.Add(('記録: {0} 秒ごと・{1}   プロセス: {2}' -f $Interval, $durationText, ($processText -join '、')))
if ($extraHandles.Count -gt 0) { $headLines.Add(('追加のカウンター: {0}' -f (Join-Limited @($extraHandles | ForEach-Object { $_[0] }) 3))) }
if ($runDir) { $headLines.Add(('保存先: {0}' -f $runDir)) }

foreach ($line in $headLines) {
    if ($line.StartsWith('※', [StringComparison]::Ordinal)) { Write-Host $line -ForegroundColor Yellow } else { Write-Host $line }
}
if (-not $inputRedirected) { Write-Host '止めるときは Ctrl+C を押してください。それまでの記録を保存します。' }
Write-Host ''
Write-Host $liveHead

# Ctrl+C をキー入力として受け取り、止めたあとにまとめと保存をできるようにする。
# コンソールがない（入力をリダイレクトした）ときは使えないので、そのときは finally でまとめを保存する
try {
    if (-not $inputRedirected) {
        $oldCtrlC = [Console]::TreatControlCAsInput
        [Console]::TreatControlCAsInput = $true
        $keyInput = $true
    }
} catch { }

$finalLines = $null
try {
    # 準備に時間がかかっても、最初の回が 1 間隔の平均になるよう、記録の直前にもう一度読み取って起点にする
    [void]$pdh.Collect()
    $sampler.Baseline()
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $previousWall = [DateTime]::UtcNow
    $nextAt = [double]$Interval
    $endAt  = if ($Duration) { [double]$Duration * 60 } else { [double]::MaxValue }

    while ($true) {
        # 次の予定の時刻まで待つ。待っている間に Ctrl+C を受け取る
        $stopRequested = $false
        while ($sw.Elapsed.TotalSeconds -lt $nextAt) {
            if ($keyInput) {
                try {
                    while ([Console]::KeyAvailable) {
                        $keyInfo = [Console]::ReadKey($true)
                        if ($keyInfo.Key -eq 'C' -and ($keyInfo.Modifiers -band [ConsoleModifiers]::Control)) { $stopRequested = $true }
                    }
                } catch { }
            }
            if ($stopRequested) { break }
            $remain = ($nextAt - $sw.Elapsed.TotalSeconds) * 1000
            if ($remain -gt 0) { [Threading.Thread]::Sleep([int][math]::Min(100, [math]::Ceiling($remain))) }
        }
        if ($stopRequested) { $endReason = 'Ctrl+C'; break }

        [void]$pdh.Collect()
        $now  = Get-Date
        $wall = [DateTime]::UtcNow
        # 前の回から間があきすぎたとき（スリープから戻ったときなど）は、あいだの長い時間の平均になるので記録しない
        $gap  = ($wall - $previousWall).TotalSeconds -gt $GapFactor * $Interval
        $previousWall = $wall
        $procRows = $sampler.Sample($now, -not $gap)

        if ($gap) {
            $gapRounds++
            Close-LimitRound $true
        } else {
            $recorded++
            $timeText = $now.ToString($TimeFormat, $Invariant)

            # CPU とメモリ
            $cpuTotal = Get-SingleValue $coreHandles['CpuTotal']
            if ($null -ne $cpuTotal) { $cpuTotal = [math]::Max(0.0, [math]::Min(100.0, $cpuTotal)) }
            $coreMax = $null
            $data = $pdh.Read($coreHandles['CpuCores'])
            for ($i = 0; $i -lt $data.Names.Count; $i++) {
                if ($data.Names[$i].Contains('_Total')) { continue }
                $v = [math]::Max(0.0, [math]::Min(100.0, $data.Values[$i]))
                if ($null -eq $coreMax -or $v -gt $coreMax) { $coreMax = $v }
            }
            $queue      = Get-SingleValue $coreHandles['Queue']
            $procCount  = Get-SingleValue $coreHandles['Processes']
            $available  = Get-SingleValue $coreHandles['Available']
            $committed  = Get-SingleValue $coreHandles['Committed']
            $commitMax  = Get-SingleValue $coreHandles['CommitLimit']
            $commitPct  = Get-SingleValue $coreHandles['CommitPct']
            $pageReads  = Get-SingleValue $coreHandles['PageReads']
            $memoryTotal = [Workbench.MeasurePerformance.SystemInfo]::TotalPhysicalMemory()
            $memoryPct  = $null
            if ($null -ne $available -and $memoryTotal -gt 0) { $memoryPct = [math]::Max(0.0, [math]::Min(100.0, ($memoryTotal - $available) / $memoryTotal * 100)) }

            Add-CsvLine 'system' @(ConvertTo-CsvLine @($timeText, (Format-Number $cpuTotal 'N1'), (Format-Number $coreMax 'N1'), (Format-Number $queue 'N0'),
                (Format-Number $memoryPct 'N1'), (Format-Number $(if ($null -ne $available) { $available / 1MB }) 'N1'),
                (Format-Number $(if ($null -ne $committed) { $committed / 1MB }) 'N1'), (Format-Number $(if ($null -ne $commitMax) { $commitMax / 1MB }) 'N1'),
                (Format-Number $commitPct 'N1'), (Format-Number $pageReads 'N1'), (Format-Number $procCount 'N0')))

            Add-Stat $systemStats.CpuTotal $cpuTotal $now
            Add-Stat $systemStats.CpuCore $coreMax $now
            Add-Stat $systemStats.Queue $queue $now
            Add-Stat $systemStats.Memory $memoryPct $now
            Add-Stat $systemStats.CommitPct $commitPct $now
            Add-Stat $systemStats.PageReads $pageReads $now

            $cpuHit    = $null -ne $cpuTotal -and $cpuTotal -ge $CpuLimit
            $queueHit  = $null -ne $queue -and $queue -ge $queueLimit
            $memoryHit = $null -ne $memoryPct -and $memoryPct -ge $MemoryLimit
            $commitHit = $null -ne $commitPct -and $commitPct -ge $CommitPctLimit
            if ($null -ne $cpuTotal) { Update-Limit 'cpu' ('CPU 使用率 {0}% 以上' -f $CpuLimit) 1 '' $cpuHit $now }
            if ($null -ne $queue) { Update-Limit 'queue' ('プロセッサ待ち行列 {0} 以上（論理プロセッサの数の {1} 倍）' -f $queueLimit, $QueuePerCpu) 2 '' $queueHit $now }
            if ($null -ne $memoryPct) { Update-Limit 'memory' ('メモリ使用率 {0}% 以上' -f $MemoryLimit) 3 '' $memoryHit $now }
            if ($null -ne $commitPct) { Update-Limit 'commit' ('コミット率 {0}% 以上' -f $CommitPctLimit) 4 '' $commitHit $now }

            # ディスク
            $liveDisk = $null
            $liveDiskHit = $false
            if ($diskHandles) {
                $drives = @{}
                foreach ($key in @($DiskCounters.Keys)) {
                    $data = $pdh.Read($diskHandles[$key])
                    for ($i = 0; $i -lt $data.Names.Count; $i++) {
                        $name = $data.Names[$i]
                        if ($name -eq '_Total') { continue }
                        if (-not $drives.ContainsKey($name)) { $drives[$name] = @{ Idle = $null; Latency = $null; LatRead = $null; LatWrite = $null; Iops = $null; ReadB = $null; WriteB = $null; Queue = $null; FreeMB = $null; FreePct = $null } }
                        $drives[$name][$key] = $data.Values[$i]
                    }
                }
                $diskLines = [System.Collections.Generic.List[string]]::new()
                foreach ($name in @($drives.Keys | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } })) {
                    $d = $drives[$name]
                    $busy    = if ($null -ne $d.Idle) { [math]::Max(0.0, [math]::Min(100.0, 100 - $d.Idle)) } else { $null }
                    $latency = if ($null -ne $d.Latency) { $d.Latency * 1000 } else { $null }
                    $readMB  = if ($null -ne $d.ReadB) { $d.ReadB / 1MB } else { $null }
                    $writeMB = if ($null -ne $d.WriteB) { $d.WriteB / 1MB } else { $null }
                    $freeGB  = if ($null -ne $d.FreeMB) { $d.FreeMB / 1024 } else { $null }
                    $diskLines.Add((ConvertTo-CsvLine @($timeText, (Protect-ExcelText $name), (Format-Number $busy 'N1'), (Format-Number $latency 'N2'),
                        (Format-Number $(if ($null -ne $d.LatRead) { $d.LatRead * 1000 }) 'N2'), (Format-Number $(if ($null -ne $d.LatWrite) { $d.LatWrite * 1000 }) 'N2'),
                        (Format-Number $d.Iops 'N1'), (Format-Number $readMB 'N1'), (Format-Number $writeMB 'N1'), (Format-Number $d.Queue 'N0'),
                        (Format-Number $freeGB 'N1'), (Format-Number $d.FreePct 'N1'))))

                    $s = $diskStats[$name]
                    if ($null -eq $s) { $s = @{ Busy = (New-Stat); Latency = (New-Stat); Read = (New-Stat); Write = (New-Stat); FreeGB = $null; FreePct = $null }; $diskStats[$name] = $s }
                    Add-Stat $s.Busy $busy $now
                    # 応答時間の平均は、読み書きのあった回だけで出す（読み書きのない回の 0 で、平均が小さく見えないように）
                    if ($null -ne $latency -and $null -ne $d.Iops -and $d.Iops -gt 0) { Add-Stat $s.Latency $latency $now }
                    Add-Stat $s.Read $readMB $now
                    Add-Stat $s.Write $writeMB $now
                    if ($null -ne $freeGB) { $s.FreeGB = $freeGB }
                    if ($null -ne $d.FreePct) { $s.FreePct = $d.FreePct }

                    $latencyHit = $null -ne $latency -and $null -ne $d.Iops -and $d.Iops -ge $LatencyMinIops -and $latency -ge $LatencyLimit
                    if ($null -ne $latency) { Update-Limit ('latency:' + $name) ('{0} の応答時間 {1} ms 以上' -f $name, $LatencyLimit) 5 $name $latencyHit $now }
                    if ($null -ne $d.FreePct) { Update-Limit ('free:' + $name) ('{0} の空き {1}% 未満' -f $name, $FreeLimit) 6 $name ($d.FreePct -lt $FreeLimit) $now }

                    # 記録中の行には、応答時間が目安を超えたドライブがあれば、そのうち最も遅いもの、なければ使用時間の最も高いものを出す
                    $candidate = @{ Name = $name; Busy = $busy; Latency = $latency; Hit = $latencyHit }
                    if ($null -eq $liveDisk) {
                        $liveDisk = $candidate
                    } elseif ($latencyHit -and (-not $liveDisk.Hit -or $latency -gt $liveDisk.Latency)) {
                        $liveDisk = $candidate
                    } elseif (-not $liveDisk.Hit -and -not $latencyHit -and $null -ne $busy -and ($null -eq $liveDisk.Busy -or $busy -gt $liveDisk.Busy)) {
                        $liveDisk = $candidate
                    }
                }
                if ($null -ne $liveDisk) { $liveDiskHit = $liveDisk.Hit }
                Add-CsvLine 'disks' $diskLines
            }

            # ネットワーク
            $netTotal = $null
            $netHit = $false
            if ($netHandles) {
                $adapters = @{}
                foreach ($key in @($NetCounters.Keys)) {
                    $data = $pdh.Read($netHandles[$key])
                    for ($i = 0; $i -lt $data.Names.Count; $i++) {
                        $name = $data.Names[$i]
                        if (-not $adapters.ContainsKey($name)) { $adapters[$name] = @{ Received = $null; Sent = $null; Bandwidth = $null } }
                        $adapters[$name][$key] = $data.Values[$i]
                    }
                }
                $netLines = [System.Collections.Generic.List[string]]::new()
                foreach ($name in @($adapters.Keys | Sort-Object -Property @{ Expression = { ConvertTo-SortKey (Get-AdapterName $_) } }, @{ Expression = { ConvertTo-SortKey $_ } })) {
                    $a = $adapters[$name]
                    $rx = if ($null -ne $a.Received) { $a.Received * 8 / 1e6 } else { $null }
                    $tx = if ($null -ne $a.Sent) { $a.Sent * 8 / 1e6 } else { $null }
                    $speed = if ($null -ne $a.Bandwidth) { $a.Bandwidth / 1e6 } else { $null }
                    # 受信と送信は同時に速度いっぱいまで使えるため、大きいほうを速度で割る
                    $util = $null
                    if ($null -ne $speed -and $speed -gt 0 -and ($null -ne $rx -or $null -ne $tx)) { $util = [math]::Max([double]$rx, [double]$tx) / $speed * 100 }
                    $label = Get-AdapterName $name
                    $netLines.Add((ConvertTo-CsvLine @($timeText, (Protect-ExcelText $label), (Protect-ExcelText $name), (Format-Number $rx 'N1'), (Format-Number $tx 'N1'), (Format-Number $util 'N1'), (Format-Number $speed 'N1'))))
                    if ($null -ne $rx -or $null -ne $tx) { $netTotal = [double]$netTotal + [double]$rx + [double]$tx }

                    $s = $netStats[$name]
                    if ($null -eq $s) { $s = @{ Received = (New-Stat); Sent = (New-Stat); Util = (New-Stat); Bandwidth = $null; Traffic = $false }; $netStats[$name] = $s }
                    Add-Stat $s.Received $rx $now
                    Add-Stat $s.Sent $tx $now
                    Add-Stat $s.Util $util $now
                    if ($null -ne $a.Bandwidth -and $a.Bandwidth -gt 0) { $s.Bandwidth = $a.Bandwidth }
                    if (([double]$rx + [double]$tx) -gt 0) { $s.Traffic = $true }
                    if ($null -ne $util) {
                        $utilHit = $util -ge $NetLimit
                        if ($utilHit) { $netHit = $true }
                        Update-Limit ('net:' + $name) ('{0} の使用率 {1}% 以上' -f $label, $NetLimit) 7 $label $utilHit $now
                    }
                }
                Add-CsvLine 'network' $netLines
            }

            # プロセス
            $topRow = if ($procRows.Count -gt 0) { $procRows[0] } else { $null }
            $selected = $sampler.Select($procRows, $Top)
            if ($selected.Count -gt 0) {
                $procLines = [System.Collections.Generic.List[string]]::new()
                foreach ($row in $selected) {
                    $procLines.Add((ConvertTo-CsvLine @($timeText, (Protect-ExcelText (ConvertTo-CellText $row.Stat.Name)), $row.Stat.Pid, (Format-Number $row.Cpu 'N1'), $row.CpuRank,
                        (Format-Number $row.Memory 'N1'), $row.MemoryRank, (Format-Number $row.Commit 'N1'), (Format-Number $row.Io 'N1'),
                        (Format-Number $row.Handles 'N0'), (Format-Number $row.Threads 'N0'), $(if ($row.Stat.Specified) { 'はい' } else { '' }), (Protect-ExcelText $row.Stat.Services))))
                }
                Add-CsvLine 'processes' $procLines
            }

            # 追加のカウンター
            if ($extraHandles.Count -gt 0) {
                $extraLines = [System.Collections.Generic.List[string]]::new()
                foreach ($extra in $extraHandles) {
                    $data = $pdh.Read($extra[1])
                    for ($i = 0; $i -lt $data.Names.Count; $i++) {
                        $extraLines.Add((ConvertTo-CsvLine @($timeText, (Protect-ExcelText $extra[0]), (Protect-ExcelText (ConvertTo-CellText $data.Names[$i])), (Format-Number $data.Values[$i] '#,##0.###'))))
                        $statKey = $extra[0] + [char]1 + $data.Names[$i]
                        $s = $counterStats[$statKey]
                        if ($null -eq $s) { $s = @{ Path = $extra[0]; Instance = $data.Names[$i]; Stat = (New-Stat) }; $counterStats[$statKey] = $s }
                        Add-Stat $s.Stat $data.Values[$i] $now
                    }
                }
                Add-CsvLine 'counters' $extraLines
            }

            Close-LimitRound $false

            # 記録中の行
            $cpuField  = (Format-PadLeft $(if ($null -ne $cpuTotal) { $cpuTotal.ToString('N0', $Invariant) + '%' } else { '-' }) 5) + (Get-Mark ($cpuHit -or $queueHit))
            $coreField = Format-PadLeft $(if ($null -ne $coreMax) { $coreMax.ToString('N0', $Invariant) + '%' } else { '-' }) 10
            $memField  = (Format-PadLeft $(if ($null -ne $memoryPct) { $memoryPct.ToString('N0', $Invariant) + '%' } else { '-' }) 6) + (Get-Mark ($memoryHit -or $commitHit))
            if ($null -ne $liveDisk) {
                $diskField = (Format-PadRight (Limit-Width $liveDisk.Name $DriveWidth) $DriveWidth) +
                    (Format-PadLeft $(if ($null -ne $liveDisk.Busy) { $liveDisk.Busy.ToString('N0', $Invariant) + '%' } else { '-' }) 4) + ' ' +
                    (Format-PadLeft $(if ($null -ne $liveDisk.Latency) { $liveDisk.Latency.ToString('N1', $Invariant) } else { '-' }) 6) + ' ms' + (Get-Mark $liveDiskHit)
            } else {
                $diskField = Format-PadRight '-' 23
            }
            $netField  = (Format-PadLeft $(if ($null -ne $netTotal) { $netTotal.ToString('N1', $Invariant) } else { '-' }) 8) + ' Mbps' + (Get-Mark $netHit)
            $procField = if ($topRow) { '{0} {1}%' -f (Limit-Width (ConvertTo-CellText $topRow.Stat.Name) $LiveNameWidth), $topRow.Cpu.ToString('N1', $Invariant) } else { '' }
            $liveLine  = (' ' + $now.ToString('HH:mm:ss', $Invariant) + '  ' + $cpuField + '  ' + $coreField + '  ' + $memField + '  ' + $diskField + '  ' + $netField + '  ' + $procField).TrimEnd()
            if ($cpuHit -or $queueHit -or $memoryHit -or $commitHit -or $liveDiskHit -or $netHit) { Write-Host $liveLine -ForegroundColor Yellow } else { Write-Host $liveLine }
        }

        # 次の予定の時刻。間に合わなかった回は飛ばす
        $elapsed = $sw.Elapsed.TotalSeconds
        $nextAt += $Interval
        while ($nextAt -le $elapsed) { $nextAt += $Interval; $skippedRounds++ }
        if ($nextAt -gt $endAt + 0.0001) { $endReason = '指定の時間'; break }
    }
} catch {
    $unexpected = $_
} finally {
    if ($keyInput) { try { [Console]::TreatControlCAsInput = $oldCtrlC } catch { } }
    # 入力をリダイレクトしているときの Ctrl+C は、ここだけが動く
    if (-not $endReason) { $endReason = if ($unexpected) { 'エラー' } else { 'Ctrl+C' } }
    $recordEnd = Get-Date

    #region まとめ

    $sumLines = [System.Collections.Generic.List[string]]::new()
    $sumLines.Add(('記録を終えました（{0}）  {1} ～ {2}（{3}・{4:N0} 回）' -f $endReason, (Format-When $recordStart), (Format-When $recordEnd), (Format-Duration ($recordEnd - $recordStart)), $recorded))
    if ($unexpected) { $sumLines.Add(('※ エラーで止まりました: {0}' -f (Get-ErrorText (Get-InnerException $unexpected.Exception)))) }

    $allStats = @()
    if ($recorded -eq 0) {
        $sumLines.Add('1 回も記録しないうちに止めました')
    } else {
        # CPU とメモリ
        $sumLines.Add('')
        $sumLines.Add('■ CPU とメモリ')
        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($item in @(@('CPU 使用率(%)', 'CpuTotal'), @('最も忙しいコア(%)', 'CpuCore'), @('プロセッサ待ち行列', 'Queue'), @('メモリ使用率(%)', 'Memory'), @('コミット率(%)', 'CommitPct'), @('ページ読み込み(回/秒)', 'PageReads'))) {
            $s = $systemStats[$item[1]]
            if ($s.Count -eq 0) { continue }
            $rows.Add([string[]]@($item[0], (Format-Number (Get-Average $s) 'N1'), ('{0}（{1}）' -f (Format-Number $s.Max 'N1'), (Format-When $s.MaxTime))))
        }
        foreach ($line in (Format-TextTable @('', '平均', '最大（時刻）') @($false, $true, $true) $rows.ToArray())) { $sumLines.Add($line) }

        # ディスク
        if ($diskStats.Count -gt 0) {
            $sumLines.Add('')
            $sumLines.Add('■ ディスク')
            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($name in @($diskStats.Keys | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } })) {
                $s = $diskStats[$name]
                $free = ''
                if ($null -ne $s.FreeGB) { $free = Format-Invariant '{0:N1} GB' $s.FreeGB }
                if ($null -ne $s.FreePct) { $free += Format-Invariant '（{0:N0}%）' $s.FreePct }
                $rows.Add([string[]]@($name, (Format-Number (Get-Average $s.Busy) 'N0'), (Format-Number $s.Busy.Max 'N0'),
                    (Format-Number $(if ($s.Latency.Count -gt 0) { Get-Average $s.Latency } else { 0 }) 'N1'), (Format-Number $(if ($s.Latency.Count -gt 0) { $s.Latency.Max } else { 0 }) 'N1'),
                    (Format-Number (Get-Average $s.Read) 'N1'), (Format-Number $s.Read.Max 'N1'), (Format-Number (Get-Average $s.Write) 'N1'), (Format-Number $s.Write.Max 'N1'), $free))
            }
            $groups = @(@('', 1), @('使用時間(%)', 2), @('応答時間(ms)', 2), @('読み取り(MB/秒)', 2), @('書き込み(MB/秒)', 2), @('', 1))
            foreach ($line in (Format-TextTable @('ドライブ', '平均', '最大', '平均', '最大', '平均', '最大', '平均', '最大', '空き') @($false, $true, $true, $true, $true, $true, $true, $true, $true, $false) $rows.ToArray() $groups)) { $sumLines.Add($line) }
        }

        # ネットワーク（通信のあったアダプターだけ）
        $netNames = @($netStats.Keys | Where-Object { $netStats[$_].Traffic } | Sort-Object -Property @{ Expression = { ConvertTo-SortKey (Get-AdapterName $_) } }, @{ Expression = { ConvertTo-SortKey $_ } })
        if ($netNames.Count -gt 0) {
            $sumLines.Add('')
            $sumLines.Add('■ ネットワーク')
            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($name in $netNames) {
                $s = $netStats[$name]
                $label = Get-AdapterName $name
                if ($s.Bandwidth) { $label += '（{0}）' -f (Format-Speed $s.Bandwidth) }
                $rows.Add([string[]]@($label, (Format-Number (Get-Average $s.Received) 'N1'), (Format-Number $s.Received.Max 'N1'),
                    (Format-Number (Get-Average $s.Sent) 'N1'), (Format-Number $s.Sent.Max 'N1'), $(if ($s.Util.Count -gt 0) { Format-Number $s.Util.Max 'N0' } else { '-' })))
            }
            $groups = @(@('', 1), @('受信(Mbps)', 2), @('送信(Mbps)', 2), @('使用率(%)', 1))
            foreach ($line in (Format-TextTable @('アダプター', '平均', '最大', '平均', '最大', '最大') @($false, $true, $true, $true, $true, $true) $rows.ToArray() $groups)) { $sumLines.Add($line) }
        }

        # プロセス。期間全体の平均は、記録した回の数で割る（動いていなかった回は 0% として数える）
        $allStats = @($sampler.StatList.ToArray())
        $summaryTop = if ($Top -gt 0) { $Top } else { 5 }
        $byCpu = @($allStats | Sort-Object -Property @{ Expression = { $_.CpuSum }; Descending = $true }, @{ Expression = { ConvertTo-SortKey $_.Name } }, @{ Expression = { $_.Pid } })
        if ($byCpu.Count -gt 0) {
            $sumLines.Add('')
            $sumLines.Add('■ CPU を多く使ったプロセス（期間の平均）')
            $rows = [System.Collections.Generic.List[object]]::new()
            $rank = 0
            foreach ($st in ($byCpu | Select-Object -First $summaryTop)) {
                $rank++
                $rows.Add([string[]]@($rank, (Format-Number ($st.CpuSum / $recorded) 'N1'), (Format-Number $st.CpuMax 'N1'), (Limit-Width (ConvertTo-CellText $st.Name) $NameWidth), $st.Pid, (Limit-Width $st.Services $ServiceWidth)))
            }
            foreach ($line in (Format-TextTable @('順位', '平均(%)', '最大(%)', 'プロセス', 'PID', 'サービス') @($true, $true, $true, $false, $true, $false) $rows.ToArray())) { $sumLines.Add($line) }

            $sumLines.Add('')
            $sumLines.Add('■ メモリを多く使ったプロセス（期間の最大）')
            $byMemory = @($allStats | Sort-Object -Property @{ Expression = { $_.MemoryMax }; Descending = $true }, @{ Expression = { ConvertTo-SortKey $_.Name } }, @{ Expression = { $_.Pid } })
            $rows = [System.Collections.Generic.List[object]]::new()
            $rank = 0
            foreach ($st in ($byMemory | Select-Object -First $summaryTop)) {
                $rank++
                $rows.Add([string[]]@($rank, (Format-Number $st.MemoryMax 'N0'), (Format-Signed ($st.CommitLast - $st.CommitFirst)), (Limit-Width (ConvertTo-CellText $st.Name) $NameWidth), $st.Pid, (Limit-Width $st.Services $ServiceWidth)))
            }
            foreach ($line in (Format-TextTable @('順位', 'メモリ(MB)', 'コミットの増減(MB)', 'プロセス', 'PID', 'サービス') @($true, $true, $true, $false, $true, $false) $rows.ToArray())) { $sumLines.Add($line) }
        }

        # 指定したプロセス（-Process）
        if ($processNames.Count -gt 0) {
            $specified = @($allStats | Where-Object { $_.Specified } | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_.Name } }, @{ Expression = { $_.Pid } })
            if ($specified.Count -gt 0) {
                $sumLines.Add('')
                $sumLines.Add('■ 指定したプロセス（-Process）')
                $rows = [System.Collections.Generic.List[object]]::new()
                foreach ($st in ($specified | Select-Object -First $SummaryRows)) {
                    $rows.Add([string[]]@((Limit-Width (ConvertTo-CellText $st.Name) $NameWidth), $st.Pid, (Format-Number ($st.CpuSum / $recorded) 'N1'), (Format-Number $st.CpuMax 'N1'), (Format-Number $st.MemoryMax 'N0'), (Format-Signed ($st.CommitLast - $st.CommitFirst))))
                }
                foreach ($line in (Format-TextTable @('プロセス', 'PID', 'CPU の平均(%)', 'CPU の最大(%)', 'メモリの最大(MB)', 'コミットの増減(MB)') @($false, $true, $true, $true, $true, $true) $rows.ToArray())) { $sumLines.Add($line) }
                if ($specified.Count -gt $SummaryRows) { $sumLines.Add((Get-MoreText ($specified.Count - $SummaryRows) 'process-list.csv')) }
            }
        }

        # 追加のカウンター（-Counter）
        if ($counterStats.Count -gt 0) {
            $sumLines.Add('')
            $sumLines.Add('■ 追加のカウンター')
            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($c in (@($counterStats.Values) | Select-Object -First $SummaryRows)) {
                $rows.Add([string[]]@((Limit-Width $c.Path $CounterWidth), (Limit-Width (ConvertTo-CellText $c.Instance) $NameWidth), (Format-Number (Get-Average $c.Stat) '#,##0.###'), ('{0}（{1}）' -f (Format-Number $c.Stat.Max '#,##0.###'), (Format-When $c.Stat.MaxTime))))
            }
            foreach ($line in (Format-TextTable @('カウンター', 'インスタンス', '平均', '最大（時刻）') @($false, $false, $true, $true) $rows.ToArray())) { $sumLines.Add($line) }
            if ($counterStats.Count -gt $SummaryRows) { $sumLines.Add((Get-MoreText ($counterStats.Count - $SummaryRows) 'counters.csv')) }
        }

        # 目安を超えたところ
        $sumLines.Add('')
        $sumLines.Add('■ 目安を超えたところ')
        $hits = @($limits.Values | Where-Object { $_.Count -gt 0 } | Sort-Object -Property @{ Expression = { $_.Order } }, @{ Expression = { $_.Sort } })
        if ($hits.Count -eq 0) {
            $sumLines.Add(' 目安を超えたところはありません')
        } else {
            foreach ($item in $hits) {
                if ($item.Count -eq 1) {
                    $sumLines.Add((' ※ {0}: 1 回（{1}）' -f $item.Label, (Format-When $item.FirstTime)))
                } elseif ($item.BestLength -le 1) {
                    $sumLines.Add((Format-Invariant ' ※ {0}: {1:N0} 回（続けて超えたことはありません）' $item.Label $item.Count))
                } else {
                    $sumLines.Add((Format-Invariant ' ※ {0}: {1:N0} 回（最も長く続いたのは {2} ～ {3} の {4:N0} 回）' $item.Label $item.Count (Format-When $item.BestStart) (Format-When $item.BestEnd) $item.BestLength))
                }
            }
        }

        $sumLines.Add('')
        $selfCpu = if ($sampler.Self.Samples -gt 0) { $sampler.Self.CpuSum / $sampler.Self.Samples } else { 0 }
        $sumLines.Add((Format-Invariant 'このツールの CPU 使用率: 平均 {0:N1}%' $selfCpu))
    }

    # 注意
    $notes = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $processNames.Count; $i++) {
        if ($recorded -gt 0 -and -not $sampler.PatternMatched[$i]) { $notes.Add(('※ -Process に当たるプロセスは、期間中に見つかりませんでした: {0}' -f $processNames[$i])) }
    }
    if ($skippedRounds -gt 0 -or $gapRounds -gt 0) {
        $notes.Add((Format-Invariant '※ 間に合わずに飛ばした回 {0:N0} 回・間があいて記録しなかった回 {1:N0} 回（スリープから戻ったときなど）' $skippedRounds $gapRounds))
    }
    foreach ($note in $splitNotes) { $notes.Add($note) }
    if ($saveError) { $notes.Add(('※ {0} から記録を保存できていません: {1}' -f (Format-When $saveErrorAt), $saveError)) }
    if (-not $isAdmin -and $recorded -gt 0) { $notes.Add('※ 管理者として実行すると、ほかのユーザーのプロセスの実行ユーザー・実行ファイル・コマンドラインも記録できます') }
    if ($notes.Count -gt 0) {
        $sumLines.Add('')
        foreach ($note in $notes) { $sumLines.Add($note) }
    }

    Write-Host ''
    foreach ($line in $sumLines) {
        if ($line.StartsWith('■', [StringComparison]::Ordinal)) { Write-Host $line -ForegroundColor Cyan }
        elseif ($line.TrimStart().StartsWith('※', [StringComparison]::Ordinal)) { Write-Host $line -ForegroundColor Yellow }
        else { Write-Host $line }
    }

    #endregion

    #region 保存（summary.txt と process-list.csv）

    if ($saving) {
        try {
            $listLines = [System.Collections.Generic.List[string]]::new()
            $listLines.Add((ConvertTo-CsvLine $ListColumns))
            $ordered = @($allStats | Sort-Object -Property @{ Expression = { $_.CpuSum }; Descending = $true }, @{ Expression = { ConvertTo-SortKey $_.Name } }, @{ Expression = { $_.Pid } })
            foreach ($st in $ordered) {
                $listLines.Add((ConvertTo-CsvLine @((Protect-ExcelText (ConvertTo-CellText $st.Name)), $st.Pid, (Protect-ExcelText $st.Services), (Protect-ExcelText (ConvertTo-CellText $st.User)),
                    $(if ($st.Started -gt [DateTime]::MinValue) { $st.Started.ToString($TimeFormat, $Invariant) } else { '' }),
                    (Protect-ExcelText (ConvertTo-CellText $st.Path)), (Protect-ExcelText (ConvertTo-CellText $st.CommandLine)),
                    $st.FirstSeen.ToString($TimeFormat, $Invariant), $st.LastSeen.ToString($TimeFormat, $Invariant),
                    (Format-Number ($st.CpuSum / [math]::Max(1, $recorded)) 'N1'), (Format-Number $st.CpuMax 'N1'), (Format-Number $st.MemoryMax 'N1'),
                    (Format-Number $st.CommitMax 'N1'), (Format-Number ($st.CommitLast - $st.CommitFirst) 'N1'))))
            }
            [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'process-list.csv')), [string[]]$listLines.ToArray(), $Utf8Bom)
            $textLines = @($headLines) + @('') + @($sumLines)
            [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$textLines, $Utf8Bom)
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
            $saving = $false
        }
    }
    if ($runDir) {
        Write-Host ''
        if ($saving) {
            Write-Host ('保存先: {0}' -f $runDir)
        } else {
            Write-Warning ('記録の一部を保存できませんでした: {0}' -f $saveError)
            Write-Host ('保存先: {0}' -f $runDir)
        }
    }

    #endregion
}

#endregion

if ($unexpected -or ($runDir -and -not $saving)) { exit 1 }
exit 0
