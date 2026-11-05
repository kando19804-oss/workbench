<#
.SYNOPSIS
    指定した名前について、hosts・DNS サーバー・OS の結果を突き合わせ、どこで引けているか、食い違いがないかを一覧にします。

.DESCRIPTION
    名前ごとに、hosts の値、DNS サーバーの答え（サーバーごと）、OS の結果を別々に調べて突き合わせ、
    引けた元と判定（OK・注意・引けない）を画面に表示し、CSV（BOM 付き UTF-8）にも保存します。
    期待する IP アドレスとの照合、hosts の全行の確認、逆引きの確認もできます。

    hosts や DNS のキャッシュを変更することはありません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Test-NameResolution.ps1 -Target "web01.example.local,db01.example.local"

.PARAMETER Target
    調べる名前。カンマで区切って複数指定できます。名前=IP アドレス の形で、期待する IP アドレスを書けます。
    IP アドレスを書くと、逆引きを調べます。

.PARAMETER List
    対象を書いた CSV ファイル（名前,期待する IP,メモ）。ひな形は -NewList で作れます。

.PARAMETER Hosts
    hosts に書かれた名前をすべて調べます。

.PARAMETER NewList
    対象のひな形の CSV を作って終わります。同じ名前のファイルがあるときは、上書きせずに止まります。

.PARAMETER Server
    聞く DNS サーバーの IP アドレス。カンマで区切って複数指定できます。
    指定すると、OS の設定の DNS サーバーには聞きません。

.PARAMETER Reverse
    得られた IP アドレスの逆引きも調べます。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.PARAMETER ConfigFile
    設定ファイルのパス。Target、List、Hosts、Server、Reverse、OutputFolder、NoSave の値を
    「項目 = 値」の形で書いておけます。コマンドラインで指定した値は、設定ファイルより優先します。
    書き方は README を参照してください。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Test-NameResolution.ps1 -Target "web01.example.local=10.0.2.10"

    切り替え後の IP アドレスになっているかを確かめます。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Test-NameResolution.ps1 -Hosts

    hosts に書かれた名前をすべて調べて、DNS と比べます。

.NOTES
    Test-NameResolution 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [string[]]$Target = @(),

    [string]$List,

    [switch]$Hosts,

    [string]$NewList,

    [string[]]$Server = @(),

    [switch]$Reverse,

    [string]$OutputFolder,

    [switch]$NoSave,

    [string]$ConfigFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion  = '1.0.0'
$ScreenRows     = 200       # 画面に表示する件数（保存するときだけ。超えた分はファイルを見てもらう）
$ProgressMs     = 250       # 進捗表示の間隔（ミリ秒）
$DeadServerLimit = 3        # この回数続けて応答しない DNS サーバーには、その後は聞かない
$OutputPrefix   = 'NameResolution_'
$ResultColumns  = '判定', '名前', '引けた元', 'OS の結果', 'hosts の値', 'hosts の行番号', 'DNS の結果', '解決した名前', 'CNAME', 'TTL', '期待する IP', '逆引き', 'メモ', '備考'
$Invariant      = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom        = New-Object Text.UTF8Encoding $true
$OnWindows      = [IO.Path]::DirectorySeparatorChar -eq '\'   # 開発用に Mac の pwsh でも動かせるようにするため
$machineName    = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

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
    if (-not $OnWindows) { return $FullPath }
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
    if ($Span.TotalSeconds -lt 60) { return [string]::Format($Invariant, '{0:N1}秒', $Span.TotalSeconds) }
    if ($Span.TotalHours -lt 1) { return '{0}分{1}秒' -f $Span.Minutes, $Span.Seconds }
    return '{0}時間{1}分{2}秒' -f [math]::Floor($Span.TotalHours), $Span.Minutes, $Span.Seconds
}

# .NET のメソッドで起きた例外は MethodInvocationException に包まれて届くため、元の例外を取り出す
function Get-InnerException([Exception]$Exception) {
    while (($Exception -is [Management.Automation.MethodInvocationException] -or $Exception -is [Management.Automation.GetValueInvocationException]) -and $Exception.InnerException) {
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

# CSV の 1 行。件数が多くなるため ConvertTo-Csv を使わずに組み立てる
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
    # 式の中の " は 2 つ重ねる
    $quotedText = $Text.Replace('"', '""')
    if ($quotedText.Length -le 255) { return '="' + $quotedText + '"' }
    return ' ' + $Text
}

# メモのように自由に書く値は、Excel が日付などに変える形を見分けきれないので、いつも ="..." で囲む
function ConvertTo-ExcelLiteral([string]$Text) {
    if ($Text -eq '') { return $Text }
    $quotedText = $Text.Replace('"', '""')
    if ($quotedText.Length -le 255) { return '="' + $quotedText + '"' }
    return ' ' + $Text
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

function Format-Column([string]$Text, [int]$Width) {
    return $Text + (' ' * [math]::Max(0, $Width - (Get-DisplayWidth $Text)))
}

function Stop-WithError([string]$Message) {
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value) {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -List "C:\data\names.csv"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# %TEMP% などの環境変数を展開する。名前を打ち間違えると %名前% のまま残り、その名前のフォルダに保存してしまうため止める。
# URL エンコードされたフォルダ名（%E3%81%82 など）を誤って止めないよう、16 進 2 桁の %XX% は対象にしない
function Expand-ConfigVariable([int]$LineNumber, [string]$Text) {
    $expanded = [Environment]::ExpandEnvironmentVariables($Text)
    if ($expanded -match '%(?![0-9A-Fa-f]{2}%)[A-Za-z_][^%\\=\s]*%') { Stop-WithConfigError $LineNumber ('定義されていない環境変数です: {0}' -f $Matches[0]) }
    return $expanded
}

# 入力されたパスをフルパスにする。末尾の \ は、ドライブのルート以外は外す
function Resolve-FullPath([string]$Text) {
    $full = ConvertTo-DisplayPath $Text
    try {
        $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($full)
        $full = [IO.Path]::GetFullPath($full)
    } catch { }
    if ($full -notmatch '^[A-Za-z]:\\$' -and $full -ne '/') { $full = $full.TrimEnd('\') }
    return $full
}

# CSV の全体を、レコード（Line = 始まりの行番号、Fields = 列の値）に分ける。
# "" で囲んだ値の中のカンマ・""・改行（Excel のセル内改行）にも対応するため、行に分けてからではなく 1 文字ずつ読む。
# ConvertFrom-Csv は、Excel が付ける空の見出しが 2 つ以上あると止まり、何行目かも分からないため使わない
function ConvertFrom-CsvText([string]$Text) {
    $records    = New-Object 'System.Collections.Generic.List[object]'
    $fields     = New-Object 'System.Collections.Generic.List[string]'
    $current    = New-Object Text.StringBuilder
    $quoted     = $false   # いま "" で囲んだ値の中にいるか
    $wasQuoted  = $false   # この値が "" で始まったか（閉じたあとの " は、ふつうの文字として扱う）
    $lineNumber = 1
    $startLine  = 1
    for ($i = 0; $i -le $Text.Length; $i++) {
        $atEnd = $i -eq $Text.Length
        $ch    = if ($atEnd) { [char]0 } else { $Text[$i] }
        if (-not $atEnd -and ($ch -eq "`n" -or ($ch -eq "`r" -and -not ($i + 1 -lt $Text.Length -and $Text[$i + 1] -eq "`n")))) { $lineNumber++ }
        if ($atEnd -and $quoted) {
            # 閉じていない " があると、後ろの行がすべて 1 つの値として読まれ、対象が黙って消えるため止める
            throw (New-Object FormatException ('{0} 行目: 引用符（"）が閉じていません' -f $startLine))
        }
        if ($quoted) {
            if ($ch -eq '"') {
                if ($i + 1 -lt $Text.Length -and $Text[$i + 1] -eq '"') { [void]$current.Append('"'); $i++ } else { $quoted = $false }
            } else {
                [void]$current.Append($ch)
            }
        } elseif ($ch -eq '"' -and $current.Length -eq 0 -and -not $wasQuoted) {
            # Excel と同じく、値の先頭の " だけを引用符の始まりとみなす（19" ラック の " は文字のまま）
            $quoted    = $true
            $wasQuoted = $true
        } elseif ($ch -eq ',') {
            $fields.Add($current.ToString())
            [void]$current.Clear()
            $wasQuoted = $false
        } elseif ($atEnd -or $ch -eq "`r" -or $ch -eq "`n") {
            # CRLF は、CR のところでレコードを閉じ、LF は空のレコードとして読み飛ばす
            $fields.Add($current.ToString())
            [void]$current.Clear()
            $wasQuoted = $false
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

# IP アドレスとして書かれた値なら IPAddress を、そうでなければ $null を返す。
# IPAddress.TryParse は「001」を 0.0.0.1 と、「192.168.001.010」を 8 進数として 192.168.1.8 と読むため、
# IPv4 は 4 つの数字を 10 進数として自分で読む。IPv6 は : を含むものだけを TryParse に任せる
function ConvertTo-IPAddress([string]$Text) {
    # \d は全角の数字にも当てはまり、[int]::Parse で例外になるので、[0-9] で書く
    if ($Text -match '^([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})$') {
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

# IP アドレスの表記をそろえる（IPv6 の省略の違いや、大文字と小文字の違いで「違う」と出ないように）。スコープ（%3 など）は外す
function ConvertTo-IpText($Address) {
    $text = $Address.ToString()
    $percent = $text.IndexOf([char]'%')
    if ($percent -gt 0) { $text = $text.Substring(0, $percent) }
    return $text
}

function Test-IPv6Text([string]$IpText) { return $IpText.Contains(':') }

# DNS の名前（CNAME、逆引きなど）は、答えるサーバーが決める値で、エスケープ文字や改行も入れられる。
# そのまま出すと画面とファイルが崩れるので、制御文字は ? に置き換える
function ConvertTo-PlainName([string]$Text) { return ($Text -replace '\p{Cc}', '?') }

# 期待する IP アドレスの文字列（; で区切る）を、そろえた表記の配列にする。正しくない値があれば、その値を Bad に入れる
function ConvertFrom-ExpectedText([string]$Text) {
    $values = New-Object 'System.Collections.Generic.List[string]'
    foreach ($part in ($Text -split ';')) {
        $trimmed = $part.Trim()
        if ($trimmed -eq '') { continue }
        $address = ConvertTo-IPAddress $trimmed
        if (-not $address) { return @{ Bad = $trimmed; Ips = @() } }
        $values.Add((ConvertTo-IpText $address))
    }
    $values.Sort([StringComparer]::Ordinal)   # 順番は比べない
    return @{ Bad = $null; Ips = $values.ToArray() }
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの。NewList は、同じ設定ファイルで調べ続けられるよう書けない
$configItems   = 'Target', 'List', 'Hosts', 'Server', 'Reverse', 'OutputFolder', 'NoSave'
$configFlags   = 'Hosts', 'Reverse', 'NoSave'
$configLists   = 'Target', 'Server'
$configPath    = $null
$configApplied = @{}   # 設定ファイルの値を使った項目 → 行番号

if ($ConfigFile) {
    $configText = Remove-Quote $ConfigFile
    if ($configText.Contains('"')) { Stop-WithQuoteError '-ConfigFile' $configText }
    $configText = Resolve-StartPath $configText
    $configPath = $configText
    try {
        $configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((ConvertTo-DisplayPath $configText))
        $configPath = [IO.Path]::GetFullPath($configPath)
    } catch { }
    if (-not [IO.File]::Exists((ConvertTo-LongPath $configPath))) {
        # エクスプローラーで拡張子を表示していないと、メモ帳で monthly.ini.txt として保存してしまうことがある
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
    $listItems    = @{}
    foreach ($itemName in $configLists) { $listItems[$itemName] = New-Object 'System.Collections.Generic.List[string]' }
    $lineNumber   = 0
    foreach ($rawLine in $configLines) {
        $lineNumber++
        $lineText = $rawLine.Trim()
        # # と ; はフォルダ名に使える文字なので、行の先頭にあるときだけコメントとみなす
        if ($lineText -eq '' -or $lineText.StartsWith('#') -or $lineText.StartsWith(';')) { continue }
        $separator = $lineText.IndexOf('=')
        if ($separator -le 0) { Stop-WithConfigError $lineNumber ('「項目 = 値」の形になっていません: {0}' -f $lineText) }

        $itemKey  = $lineText.Substring(0, $separator).Trim()
        if ($itemKey -eq 'NewList') { Stop-WithConfigError $lineNumber 'NewList は設定ファイルには書けません。コマンドラインで -NewList を指定してください' }
        $itemName = $configItems | Where-Object { $_ -eq $itemKey }
        if (-not $itemName) { Stop-WithConfigError $lineNumber ('知らない項目です: {0}' -f $itemKey) ('使える項目: {0}' -f ($configItems -join ', ')) }

        $rawValue  = $lineText.Substring($separator + 1).Trim()
        $itemValue = Remove-ConfigQuote $rawValue
        # 値が空の行は、書かなかったのと同じ扱い（ひな形の空の行の下に書き足しても、2 回書いたことにしない）
        if ($itemValue -eq '') { continue }

        # Target と Server は複数の行に分けて書ける。ほかの項目は、どちらが効いているか分からなくなるため 1 回だけ
        if ($configLists -notcontains $itemName -and $configLineOf.ContainsKey($itemName)) {
            Stop-WithConfigError $lineNumber ('{0} が 2 回書かれています（{1} 行目と {2} 行目）' -f $itemName, $configLineOf[$itemName], $lineNumber)
        }
        $configLineOf[$itemName] = $lineNumber

        if ($configLists -contains $itemName) {
            # "a,b" のように全体を囲む書き方と、"a","b" のように 1 つずつ囲む書き方のどちらも受け付ける
            $source = if ($itemValue.Contains('"')) { $rawValue } else { $itemValue }
            foreach ($part in ($source -split ',')) {
                $pattern = Remove-ConfigQuote $part
                if ($pattern.Contains('"')) { Stop-WithConfigError $lineNumber ('{0} の値に引用符（"）が残っています: {1}' -f $itemName, $pattern) }
                if (-not $pattern) { continue }
                $pattern = Expand-ConfigVariable $lineNumber $pattern
                if ($itemName -eq 'Server' -and -not (ConvertTo-IPAddress $pattern)) { Stop-WithConfigError $lineNumber ('Server は IP アドレスで指定してください: {0}' -f $pattern) }
                if ($itemName -eq 'Target' -and $pattern.Contains('=')) {
                    $checked = ConvertFrom-ExpectedText ($pattern.Substring($pattern.IndexOf('=') + 1))
                    if ($checked.Bad) { Stop-WithConfigError $lineNumber ('Target の期待する IP アドレスが正しくありません: {0}' -f $checked.Bad) }
                    # = を書いたのに値がないのは、環境変数が空のときなど。照合しないまま OK にしないよう止める
                    if ($checked.Ips.Count -eq 0) { Stop-WithConfigError $lineNumber ('Target の期待する IP アドレスが空です: {0}' -f $pattern) }
                }
                $listItems[$itemName].Add($pattern)
            }
        } elseif ($configFlags -contains $itemName) {
            if ($itemValue -eq 'true') { $configValues[$itemName] = $true }
            elseif ($itemValue -eq 'false') { $configValues[$itemName] = $false }
            else { Stop-WithConfigError $lineNumber ('{0} は true か false で指定してください: {1}' -f $itemName, $itemValue) }
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

#region ひな形を作る（-NewList）

if ($NewList) {
    $newText = Remove-Quote $NewList
    if ($newText.Contains('"')) { Stop-WithQuoteError '-NewList' $newText }
    $newPath = Resolve-FullPath (Resolve-StartPath $newText)
    if ([IO.File]::Exists((ConvertTo-LongPath $newPath)) -or [IO.Directory]::Exists((ConvertTo-LongPath $newPath))) {
        Stop-WithError ('同じ名前のファイルがあります: {0}（上書きしません）' -f $newPath)
    }
    $template = @(
        '名前,期待する IP,メモ'
        '# 先頭が # の行は調べません。期待する IP は空欄でかまいません。複数のときは ; で区切ります,,'
        'web01.example.local,10.0.2.10,切り替え後の Web サーバー'
        'db01.example.local,,データベース'
    )
    try {
        # Excel でダブルクリックして開けるよう、BOM 付き UTF-8 で書く。既存のファイルは上書きしない
        $stream = New-Object IO.FileStream ((ConvertTo-LongPath $newPath), [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
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

#region 引数の確認

$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    $outText = Resolve-StartPath $outText
}

if (-not $List -and $Target.Count -eq 0 -and -not $Hosts) { Stop-WithError '調べる名前を指定してください（-Target、-List、-Hosts のどれかで指定します）。ひな形は -NewList で作れます。' }

$serverTexts = New-Object 'System.Collections.Generic.List[string]'
foreach ($entry in $Server) {
    foreach ($part in ($entry -split ',')) {
        $text = (Remove-Quote $part).Trim("'").Trim()
        if ($text -eq '') { continue }
        if ($text.Contains('"')) { Stop-WithQuoteError '-Server' $text }
        $address = ConvertTo-IPAddress $text
        if (-not $address) { Stop-WithError ('-Server は IP アドレスで指定してください: {0}' -f $text) }
        $normalized = ConvertTo-IpText $address
        if ($serverTexts -notcontains $normalized) { $serverTexts.Add($normalized) }
    }
}

if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) {
    Stop-WithError 'Resolve-DnsName を使えないため、調べられません（DnsClient モジュールが見つかりません）'
}

# hosts の場所は OS の設定（DataBasePath）に従う。読めなければ、Windows の既定の場所とする
function Get-HostsPath {
    try {
        $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SYSTEM\CurrentControlSet\Services\Tcpip\Parameters')
        if ($key) {
            try {
                $folder = $key.GetValue('DataBasePath')
                if ($folder) { return (Join-Path ([string]$folder) 'hosts') }
            } finally {
                $key.Close()
            }
        }
    } catch { }
    return (Join-Path ([string]$env:windir) 'System32\drivers\etc\hosts')
}

# hosts を読み、名前 → 項目（IP、行番号）の一覧 の表にする。# より後ろはコメント
$hostsMap   = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
$hostsOrder = New-Object 'System.Collections.Generic.List[string]'
$hostsError = $null
$hostsLines = 0
$hostsPath  = Get-HostsPath
try {
    $hostsText = Read-ConfigText $hostsPath
    $hostsLineNumber = 0
    foreach ($rawLine in ($hostsText -split '\r\n|\n|\r')) {
        $hostsLineNumber++
        $body = $rawLine
        $hash = $body.IndexOf([char]'#')
        if ($hash -ge 0) { $body = $body.Substring(0, $hash) }
        $tokens = $body.Split([char[]]@(' ', "`t"), [StringSplitOptions]::RemoveEmptyEntries)
        if ($tokens.Count -lt 2) { continue }
        $address = ConvertTo-IPAddress $tokens[0]
        if (-not $address) { continue }
        $hostsLines++
        for ($i = 1; $i -lt $tokens.Count; $i++) {
            if (-not $hostsMap.ContainsKey($tokens[$i])) {
                $hostsMap[$tokens[$i]] = New-Object 'System.Collections.Generic.List[object]'
                $hostsOrder.Add($tokens[$i])
            }
            $hostsMap[$tokens[$i]].Add(@{ Ip = (ConvertTo-IpText $address); Line = $hostsLineNumber })
        }
    }
} catch {
    $hostsError = Get-ErrorText (Get-InnerException $_.Exception)
}

# 対象を読む。同じ名前は 1 回だけ調べる。後から書いたほうに期待する IP やメモがあれば、足りない分を補う
$items     = New-Object 'System.Collections.Generic.List[object]'
$itemIndex = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
function Add-Item([string]$Name, [string[]]$ExpectedValues, [string]$Memo) {
    $existing = $null
    if ($itemIndex.TryGetValue($Name, [ref]$existing)) {
        # 後から読むのは -Target なので、期待する IP アドレスはコマンドラインの値を優先する
        if ($ExpectedValues.Count -gt 0) { $existing.Expected = $ExpectedValues }
        if ($Memo -and -not $existing.Memo) { $existing.Memo = $Memo }
        return
    }
    $record = [pscustomobject]@{ Name = $Name; Expected = $ExpectedValues; Memo = $Memo }
    $itemIndex[$Name] = $record
    $items.Add($record)
}

if ($List) {
    $listText = Remove-Quote $List
    if ($listText.Contains('"')) { Stop-WithQuoteError '-List' $listText }
    $listPath = Resolve-FullPath (Resolve-StartPath $listText)
    if (-not [IO.File]::Exists((ConvertTo-LongPath $listPath))) {
        $message = ('対象の CSV が見つかりません: {0}' -f $listPath) + (Get-ConfigNote 'List')
        if ([IO.File]::Exists((ConvertTo-LongPath ($listPath + '.txt')))) { $message += "`r`n" + ('拡張子が .txt になっていませんか: {0}.txt' -f $listPath) }
        Stop-WithError $message
    }
    try {
        $listRecords = ConvertFrom-CsvText (Read-ConfigText (ConvertTo-LongPath $listPath))
    } catch {
        $readError = Get-InnerException $_.Exception
        if ($readError -is [FormatException]) {
            Stop-WithError (('対象の CSV の {0}' -f $readError.Message) + "`r`n" + ('ファイル: {0}' -f $listPath))
        }
        Stop-WithError ('対象の CSV を読めません: {0}（{1}）' -f $listPath, (Get-ErrorText $readError))
    }
    $nameCol = -1
    $expectedCol = -1
    $memoCol = -1
    foreach ($record in $listRecords) {
        $fields = $record.Fields
        if ($nameCol -lt 0) {
            # 最初の空でない行は見出し。日本語と英語のどちらの見出しも受け付ける
            for ($i = 0; $i -lt $fields.Count; $i++) {
                # 手で書くと空白が揺れやすいので、空白を除いて比べる
                $heading = ($fields[$i] -replace '\s', '').ToLowerInvariant()
                if ($heading -eq '名前' -or $heading -eq 'name') { $nameCol = $i }
                elseif ($heading -eq '期待するip' -or $heading -eq 'expected' -or $heading -eq 'expectedip') { $expectedCol = $i }
                elseif ($heading -eq 'メモ' -or $heading -eq 'memo') { $memoCol = $i }
                elseif ($heading.Contains('期待') -or $heading.Contains('expect')) {
                    # 読めない見出しを無視すると、照合しないまま OK になってしまうので止める
                    Stop-WithError (('対象の CSV の見出しを読めません: {0}（「期待する IP」と書いてください）' -f $fields[$i].Trim()) + "`r`n" + ('ファイル: {0}' -f $listPath))
                }
            }
            if ($nameCol -lt 0) {
                Stop-WithError ('対象の CSV に「名前」の列がありません。1 行目を見出し（名前,期待する IP,メモ）にしてください' + "`r`n" + ('ファイル: {0}' -f $listPath))
            }
            continue
        }
        $nameText = if ($nameCol -lt $fields.Count) { $fields[$nameCol].Trim() } else { '' }
        if ($nameText -eq '' -or $nameText.StartsWith('#')) { continue }
        $expectedText = if ($expectedCol -ge 0 -and $expectedCol -lt $fields.Count) { $fields[$expectedCol] } else { '' }
        # セル内改行は、画面の表がくずれないよう空白にする
        $memoText = if ($memoCol -ge 0 -and $memoCol -lt $fields.Count) { ($fields[$memoCol] -replace '\s*[\r\n]+\s*', ' ').Trim() } else { '' }
        $checked = ConvertFrom-ExpectedText $expectedText
        if ($checked.Bad) {
            Stop-WithError (('対象の CSV の {0} 行目: 期待する IP アドレスが正しくありません: {1}' -f $record.Line, $checked.Bad) + "`r`n" + ('ファイル: {0}' -f $listPath))
        }
        # 2 台構成を行に分けて書くと、片方だけを期待することになるので止める
        $listed = $null
        if ($checked.Ips.Count -gt 0 -and $itemIndex.TryGetValue($nameText, [ref]$listed) -and $listed.Expected.Count -gt 0 -and
            -not [string]::Equals(($listed.Expected -join ';'), ($checked.Ips -join ';'), [StringComparison]::Ordinal)) {
            Stop-WithError (('対象の CSV の {0} 行目: 同じ名前に別の期待する IP アドレスが書かれています: {1}（複数のときは ; で区切って 1 行に書きます）' -f $record.Line, $nameText) + "`r`n" + ('ファイル: {0}' -f $listPath))
        }
        Add-Item $nameText $checked.Ips $memoText
    }
}

# -File で起動すると配列を渡せないため、カンマ区切りの文字列で受ける
$targetExpected = @{}   # -Target の中で、同じ名前に別の期待する IP アドレスを書いていないかを見る
foreach ($entry in $Target) {
    foreach ($part in ($entry -split ',')) {
        $text = (Remove-Quote $part).Trim("'").Trim()
        if ($text -eq '') { continue }
        if ($text.Contains('"')) { Stop-WithQuoteError '-Target' $text }
        $nameText = $text
        $expectedValues = @()
        $equal = $text.IndexOf([char]'=')
        if ($equal -ge 0) {
            $nameText = $text.Substring(0, $equal).Trim()
            $checked = ConvertFrom-ExpectedText ($text.Substring($equal + 1))
            if ($checked.Bad) { Stop-WithError ('-Target の期待する IP アドレスが正しくありません: {0}' -f $checked.Bad) }
            $expectedValues = $checked.Ips
            if ($nameText -eq '') { Stop-WithError ('-Target の名前がありません: {0}' -f $text) }
            # = を書いたのに値がないのは、バッチの変数が空のときなど。照合しないまま OK にしないよう止める
            if ($expectedValues.Count -eq 0) { Stop-WithError ('-Target の期待する IP アドレスが空です: {0}' -f $text) }
            $expectedKey = $nameText.ToLowerInvariant()
            $expectedJoined = $expectedValues -join ';'
            if ($targetExpected.ContainsKey($expectedKey) -and -not [string]::Equals($targetExpected[$expectedKey], $expectedJoined, [StringComparison]::Ordinal)) {
                Stop-WithError ('-Target で、同じ名前に別の期待する IP アドレスが書かれています: {0}（複数のときは ; で区切ります）' -f $nameText)
            }
            $targetExpected[$expectedKey] = $expectedJoined
        }
        Add-Item $nameText $expectedValues ''
    }
}

if ($Hosts) {
    # hosts を調べるように指定されて読めないのは、続けても意味がないので止める
    if ($hostsError) { Stop-WithError ('hosts を読めません: {0}（{1}）' -f $hostsPath, $hostsError) }
    foreach ($nameText in $hostsOrder) {
        if ($nameText -ieq 'localhost') { continue }
        $hostsFound = $null
        [void]$hostsMap.TryGetValue($nameText, [ref]$hostsFound)
        $hostsIps = @($hostsFound | ForEach-Object { $_.Ip })
        if (@($hostsIps | Where-Object { $_ -ne '127.0.0.1' -and $_ -ne '::1' }).Count -eq 0) { continue }
        Add-Item $nameText @() ''
    }
}

if ($items.Count -eq 0) { Stop-WithError '調べる対象がありません' }

#endregion

#region 判定の部品

$startTime = Get-Date
$progress  = @{ Next = 0L }
$sw        = [Diagnostics.Stopwatch]::StartNew()
$failStreak = @{}   # DNS サーバー → 正引きに続けて応答しなかった回数
$ptrStreak  = @{}   # DNS サーバー → 逆引きに続けて応答しなかった回数（逆引きだけ答えないサーバーを、正引きでも聞かなくならないよう分ける）
# この PC 自身の名前。OS がすべての NIC のアドレスを返すので、OS の結果と DNS の答えを比べない
$selfNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
[void]$selfNames.Add($machineName)
try {
    $selfHost = [Net.Dns]::GetHostName()
    [void]$selfNames.Add($selfHost)
    $selfDomain = [Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
    if ($selfDomain) { [void]$selfNames.Add($selfHost + '.' + $selfDomain) }
} catch { }

function Show-Progress([string]$Status, [string]$Operation) {
    if ($sw.ElapsedMilliseconds -lt $progress.Next) { return }
    Write-Progress -Activity '名前解決確認' -Status $Status -CurrentOperation $Operation
    $progress.Next = $sw.ElapsedMilliseconds + $ProgressMs
}

# OS の設定の DNS サーバーを読む。つながっているインターフェイスのものだけ。fec0:0:0:ffff:: は、Windows が初めから持っている使われない値
function Get-OsDnsServer {
    $found = New-Object 'System.Collections.Generic.List[string]'
    $connected = $null
    try { $connected = @(Get-NetIPInterface -ConnectionState Connected -ErrorAction Stop | ForEach-Object { $_.InterfaceIndex }) } catch { }
    foreach ($entry in @(Get-DnsClientServerAddress -ErrorAction Stop)) {
        if ($connected -and $connected -notcontains $entry.InterfaceIndex) { continue }
        foreach ($serverText in @($entry.ServerAddresses)) {
            if (-not $serverText -or $serverText -like 'fec0:0:0:ffff::*') { continue }
            if ($found -notcontains $serverText) { $found.Add($serverText) }
        }
    }
    return , $found
}

# 短い名前に付ける DNS サフィックスを読む。Windows と同じく、検索一覧が設定されていればそれだけを使い、
# なければプライマリ DNS サフィックスと、接続ごとのサフィックスを使う
function Get-DnsSuffix {
    $found = New-Object 'System.Collections.Generic.List[string]'
    try {
        foreach ($suffix in @((Get-DnsClientGlobalSetting -ErrorAction Stop).SuffixSearchList)) {
            if ($suffix -and $found -notcontains $suffix) { $found.Add($suffix) }
        }
    } catch { }
    if ($found.Count -gt 0) { return , $found }
    try {
        $primary = [Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
        if ($primary) { $found.Add($primary) }
    } catch { }
    try {
        foreach ($entry in @(Get-DnsClient -ErrorAction Stop)) {
            if ($entry.ConnectionSpecificSuffix -and $found -notcontains $entry.ConnectionSpecificSuffix) { $found.Add($entry.ConnectionSpecificSuffix) }
        }
    } catch { }
    return , $found
}

# Resolve-DnsName のエラーを分類する。文の言語によらないよう、番号と ID で見分ける。
# DNS が動いていないサーバー（ポートが閉じている）や経路のないサーバーは、タイムアウトを待たずに接続のエラーが返るので、これも「応答しない」とする
# （ネットワークのエラー: 1231、1232、10050、10051、10054、10061、10064、10065。タイムアウト: 1460、10060）
$NoResponseCodes = 1231, 1232, 1460, 10050, 10051, 10054, 10060, 10061, 10064, 10065
function Get-DnsErrorKind($ErrorRecord) {
    $code = 0
    $ex = $ErrorRecord.Exception
    if ($ex -is [ComponentModel.Win32Exception]) { $code = $ex.NativeErrorCode }
    $id = [string]$ErrorRecord.FullyQualifiedErrorId
    if ($code -eq 9003 -or $id -match 'DNS_ERROR_RCODE_NAME_ERROR|\b9003\b') { return 'NoName' }
    if ($code -eq 9501 -or $id -match 'DNS_INFO_NO_RECORDS|\b9501\b') { return 'NoData' }
    if ($NoResponseCodes -contains $code -or $id -match 'ERROR_TIMEOUT|WSAETIMEDOUT|\b(?:1231|1232|1460|10050|10051|10054|10060|10061|10064|10065)\b') { return 'NoResponse' }
    return 'Other'
}

# エラーの文は「名前 : 内容」の形で、内容は OS の言語で変わる。DNS サーバーが返したエラーは RCODE の名前で書き、それ以外は名前を除いて書く
function Get-DnsErrorText($ErrorRecord) {
    $ex = $ErrorRecord.Exception
    if ($ex -is [ComponentModel.Win32Exception]) {
        switch ($ex.NativeErrorCode) {
            9001 { return 'FORMERR' }
            9002 { return 'SERVFAIL' }
            9004 { return 'NOTIMP' }
            9005 { return 'REFUSED' }
        }
    }
    return ((Get-ErrorText (Get-InnerException $ex)) -replace '^\S+ : ', '')
}

function Get-RecordValue($Record, [string]$Property) {
    $member = $Record.PSObject.Properties[$Property]
    if ($member) { return $member.Value }
    return $null
}

# 1 つの DNS サーバーに、1 つの名前の 1 つの種類を聞く
function Invoke-DnsQuery([string]$Name, [string]$Type, [string]$DnsServer) {
    try {
        # -DnsOnly と -NoHostsFile で hosts を使わず、-Server を指定してキャッシュも使わない
        $records = @(Resolve-DnsName -Name $Name -Type $Type -DnsOnly -NoHostsFile -Server $DnsServer -ErrorAction Stop)
        return @{ Kind = 'Ok'; Records = @($records | Where-Object { [string](Get-RecordValue $_ 'Section') -eq 'Answer' }); Detail = '' }
    } catch {
        return @{ Kind = (Get-DnsErrorKind $_); Records = @(); Detail = (Get-DnsErrorText $_) }
    }
}

# 1 つの DNS サーバーの、1 つの名前の答えをまとめる。短い名前はサフィックスを順に付けて聞く
function Get-ServerAnswer([string]$Name, [string]$DnsServer, $Suffixes) {
    $answer = [pscustomobject]@{ Server = $DnsServer; State = 'None'; Ips = @(); Resolved = ''; Chain = @(); Ttl = -1; Detail = '名前が存在しない'; Warning = '' }
    if ($failStreak.ContainsKey($DnsServer) -and $failStreak[$DnsServer] -ge $DeadServerLimit) {
        $answer.State = 'Skipped'
        $answer.Detail = '応答しないので聞いていない'
        return $answer
    }
    $candidates = New-Object 'System.Collections.Generic.List[string]'
    if ($Name.Contains('.')) {
        $candidates.Add($Name)
    } else {
        foreach ($suffix in $Suffixes) { $candidates.Add($Name + '.' + $suffix) }
        $candidates.Add($Name)
    }
    foreach ($candidate in $candidates) {
        $ips = New-Object 'System.Collections.Generic.List[string]'
        $cnames = New-Object 'System.Collections.Generic.List[string]'
        $resolvedName = ''
        $ttl = -1
        $noResponseCount = 0
        $otherDetails = New-Object 'System.Collections.Generic.List[string]'
        $typeWarnings = New-Object 'System.Collections.Generic.List[string]'
        foreach ($type in 'A', 'AAAA') {
            $query = Invoke-DnsQuery $candidate $type $DnsServer
            if ($query.Kind -eq 'NoResponse') {
                $noResponseCount++
                $typeWarnings.Add(('{0} は応答なし' -f $type))
                continue
            }
            if ($query.Kind -eq 'Other') {
                $otherDetails.Add($query.Detail)
                $typeWarnings.Add(('{0}: {1}' -f $type, $query.Detail))
                continue
            }
            # 名前はあるが、A と AAAA の登録がない
            if (($query.Kind -eq 'NoData' -or $query.Kind -eq 'Ok') -and $answer.State -eq 'None') { $answer.Detail = 'A / AAAA の登録がない' }
            foreach ($record in $query.Records) {
                $recordType = [string](Get-RecordValue $record 'Type')
                $recordTtl  = Get-RecordValue $record 'TTL'
                if ($null -ne $recordTtl -and ($ttl -lt 0 -or [long]$recordTtl -lt $ttl)) { $ttl = [long]$recordTtl }
                if ($recordType -eq 'CNAME' -or $recordType -eq '5') {
                    if ($cnames.Count -eq 0) { $cnames.Add([string](Get-RecordValue $record 'Name')) }
                    $cnameTarget = [string](Get-RecordValue $record 'NameHost')
                    if ($cnames -notcontains $cnameTarget) { $cnames.Add($cnameTarget) }
                } elseif ($recordType -eq 'A' -or $recordType -eq 'AAAA' -or $recordType -eq '1' -or $recordType -eq '28') {
                    $address = ConvertTo-IPAddress ([string](Get-RecordValue $record 'IPAddress'))
                    if ($address) {
                        $ipText = ConvertTo-IpText $address
                        if ($ips -notcontains $ipText) { $ips.Add($ipText) }
                        $resolvedName = [string](Get-RecordValue $record 'Name')
                    }
                }
            }
        }
        # A と AAAA が同じエラーなら、1 つにまとめる
        if ($otherDetails.Count -eq 2 -and [string]::Equals($otherDetails[0], $otherDetails[1], [StringComparison]::Ordinal)) {
            $typeWarnings.Clear()
            $typeWarnings.Add($otherDetails[0])
        }
        # A と AAAA のどちらも応答しないときだけ、応答しないサーバーとして数える
        if ($noResponseCount -eq 2) {
            $failStreak[$DnsServer] = [int]$(if ($failStreak.ContainsKey($DnsServer)) { $failStreak[$DnsServer] } else { 0 }) + 1
            $answer.State = 'NoResponse'
            $answer.Detail = '応答なし'
            return $answer
        }
        $failStreak[$DnsServer] = 0
        $answer.Warning = $typeWarnings -join '、'
        if ($ips.Count -gt 0) {
            $sorted = New-Object 'System.Collections.Generic.List[string]' (, [string[]]$ips.ToArray())
            $sorted.Sort([StringComparer]::Ordinal)
            $answer.State = 'Ok'
            $answer.Ips = $sorted.ToArray()
            $answer.Resolved = $resolvedName.TrimEnd('.')
            $answer.Chain = $cnames.ToArray()
            $answer.Ttl = $ttl
            $answer.Detail = ''
            return $answer
        }
        # 片方が答えられず、もう片方に IP アドレスがないときは、「ない」とは言い切れない
        if ($otherDetails.Count -gt 0 -or $noResponseCount -gt 0) {
            $answer.State = 'Error'
            $answer.Detail = $typeWarnings -join '、'
            $answer.Warning = ''
        }
    }
    return $answer
}

# OS の結果。アプリが受け取るのと同じ呼び出しで、hosts とキャッシュを含む
function Get-OsResult([string]$Name) {
    try {
        $texts = New-Object 'System.Collections.Generic.List[string]'
        foreach ($address in [Net.Dns]::GetHostAddresses($Name)) {
            $ipText = ConvertTo-IpText $address
            if ($texts -notcontains $ipText) { $texts.Add($ipText) }
        }
        $texts.Sort([StringComparer]::Ordinal)
        return @{ Ips = $texts.ToArray(); Detail = '' }
    } catch {
        return @{ Ips = @(); Detail = (Get-ErrorText (Get-InnerException $_.Exception)) }
    }
}

# 逆引き。登録された名前の配列を返す。登録がなければ空の配列、どのサーバーにも聞けなければ $null。
# 応答しないサーバーは、次のサーバーに移る
function Get-ReverseName([string]$IpText) {
    foreach ($dnsServer in $dnsServers) {
        if ($failStreak.ContainsKey($dnsServer) -and $failStreak[$dnsServer] -ge $DeadServerLimit) { continue }
        if ($ptrStreak.ContainsKey($dnsServer) -and $ptrStreak[$dnsServer] -ge $DeadServerLimit) { continue }
        $query = Invoke-DnsQuery $IpText 'PTR' $dnsServer
        if ($query.Kind -eq 'NoResponse') {
            $ptrStreak[$dnsServer] = [int]$(if ($ptrStreak.ContainsKey($dnsServer)) { $ptrStreak[$dnsServer] } else { 0 }) + 1
            continue
        }
        $ptrStreak[$dnsServer] = 0
        if ($query.Kind -eq 'Other') { continue }
        $names = New-Object 'System.Collections.Generic.List[string]'
        foreach ($record in $query.Records) {
            $hostName = Get-RecordValue $record 'NameHost'
            if ($hostName) { $names.Add(([string]$hostName).TrimEnd('.')) }
        }
        return , $names
    }
    return $null
}

# 期待する IP アドレスと一致するか。すべて含まれていて、期待していない同じ種類の IP アドレスがなければ一致
function Test-ExpectedMatch([string[]]$Actual, [string[]]$Expected) {
    foreach ($value in $Expected) { if ($Actual -notcontains $value) { return $false } }
    $expectV4 = @($Expected | Where-Object { -not (Test-IPv6Text $_) }).Count -gt 0
    $expectV6 = @($Expected | Where-Object { Test-IPv6Text $_ }).Count -gt 0
    foreach ($value in $Actual) {
        $isV6 = Test-IPv6Text $value
        if ($Expected -notcontains $value -and (($isV6 -and $expectV6) -or (-not $isV6 -and $expectV4))) { return $false }
    }
    return $true
}

function Join-Ip($Ips) { return (@($Ips) -join ', ') }

#endregion

#region 保存先の準備

# 調べ終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir      = $null
$saveWarning = $null
if (-not $NoSave) {
    $baseDir = if ($outText) { $outText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        # \\?\ 付きのパスでは末尾のピリオドや空白が残るため、先に通常の書き方で整える（260 文字を超えると例外になるので、そのときはそのまま）
        try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
        $runName = Format-Invariant '{0}{1}_{2:yyyyMMdd_HHmmss}' $OutputPrefix $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while ([IO.Directory]::Exists((ConvertTo-LongPath $runDir)) -or [IO.File]::Exists((ConvertTo-LongPath $runDir))) {
            $suffix++
            $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix)
        }
        [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $runDir))
    } catch {
        $reason = Get-ErrorText (Get-InnerException $_.Exception)
        $runDir = $null
        $saveWarning = ('保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, $reason) + (Get-ConfigNote 'OutputFolder')
        Write-Warning $saveWarning
    }
}

#endregion

# 調べている途中や保存の途中で Ctrl+C されたときや、保存に失敗したときは、書きかけの結果を残さない
$finished   = $false
$unexpected = $false
$saveError  = $null
$exitCode   = 0
$lines      = New-Object 'System.Collections.Generic.List[string]'   # summary.txt に書く行
try {
    #region 調べる

    $serverListText = '（OS の設定）'
    $dnsServers = $serverTexts
    $serverError = $null
    if ($dnsServers.Count -gt 0) {
        $serverListText = '（指定）'
    } else {
        try {
            $dnsServers = Get-OsDnsServer
        } catch {
            $serverError = Get-ErrorText (Get-InnerException $_.Exception)
            $dnsServers = New-Object 'System.Collections.Generic.List[string]'
        }
    }
    $suffixList = Get-DnsSuffix

    $results = New-Object 'System.Collections.Generic.List[object]'
    $done = 0
    foreach ($item in $items) {
        $done++
        Show-Progress (Format-Invariant '調べています: {0:N0} / {1:N0} 件' $done $items.Count) $item.Name
        $notes = New-Object 'System.Collections.Generic.List[string]'
        $attention = $false
        $inputAddress = ConvertTo-IPAddress $item.Name

        # OS の結果
        $osResult = if ($inputAddress) { @{ Ips = @((ConvertTo-IpText $inputAddress)); Detail = '' } } else { Get-OsResult $item.Name }
        $osIps = @($osResult.Ips)
        $osFound = $osIps.Count -gt 0

        # hosts
        $hostsEntries = @()
        # List を @( ) で囲むと、7.6 で「Argument types do not match」になることがあるので、ToArray を使う
        $hostsFound = $null
        if (-not $inputAddress -and $hostsMap.TryGetValue($item.Name, [ref]$hostsFound)) { $hostsEntries = $hostsFound.ToArray() }
        $hostsIps = @($hostsEntries | ForEach-Object { $_.Ip } | Select-Object -Unique)
        $hostsLineText = (@($hostsEntries | ForEach-Object { $_.Line } | Select-Object -Unique) -join ', ')

        # DNS サーバー
        $answers = New-Object 'System.Collections.Generic.List[object]'
        if (-not $inputAddress) {
            foreach ($dnsServer in $dnsServers) { $answers.Add((Get-ServerAnswer $item.Name $dnsServer $suffixList)) }
        }
        $okAnswers = @($answers | Where-Object { $_.State -eq 'Ok' })
        $dnsIps = @($okAnswers | ForEach-Object { $_.Ips } | Select-Object -Unique)

        # 引けた元
        $sourceText = '-'
        if ($inputAddress) { $sourceText = 'IP アドレス' }
        elseif ($hostsEntries.Count -gt 0) { $sourceText = 'hosts' }
        elseif ($okAnswers.Count -gt 0) { $sourceText = 'DNS' }
        elseif ($osFound -and @($answers | Where-Object { $_.State -eq 'None' }).Count -gt 0) { $sourceText = 'その他' }
        elseif ($osFound) { $sourceText = 'OS のみ' }

        # DNS サーバーに聞けなかったときは、hosts との食い違いやキャッシュの古さを確かめられていない
        $dnsUnknown = -not $inputAddress -and $okAnswers.Count -eq 0 -and @($answers | Where-Object { $_.State -eq 'None' }).Count -eq 0
        if ($dnsUnknown) {
            $attention = $true
            if ($answers.Count -eq 0) { $notes.Add('DNS サーバーに聞けなかった') }
        }

        # hosts と DNS を比べる。hosts に書かれた種類（IPv4 か IPv6）だけを比べる
        if ($hostsEntries.Count -gt 0) {
            $hostsV4 = @($hostsIps | Where-Object { -not (Test-IPv6Text $_) })
            $hostsV6 = @($hostsIps | Where-Object { Test-IPv6Text $_ })
            if ($hostsV4.Count -gt 1 -or $hostsV6.Count -gt 1) {
                $attention = $true
                $notes.Add(('hosts に同じ名前が複数ある（{0} 行目）' -f (@($hostsEntries | ForEach-Object { $_.Line } | Select-Object -Unique) -join ' 行目と ')))
            }
            if ($okAnswers.Count -gt 0) {
                $differs = $false
                foreach ($ipText in $hostsIps) {
                    $sameKind = @($dnsIps | Where-Object { (Test-IPv6Text $_) -eq (Test-IPv6Text $ipText) })
                    if ($sameKind.Count -gt 0 -and $sameKind -notcontains $ipText) { $differs = $true }
                }
                if ($differs) {
                    $attention = $true
                    $notes.Add(('hosts と DNS が違う（hosts の {0} 行目）' -f $hostsLineText))
                }
            } elseif (@($answers | Where-Object { $_.State -eq 'None' }).Count -gt 0) {
                $notes.Add('DNS にはない')
            }
        } elseif ($osFound -and -not $inputAddress -and $okAnswers.Count -eq 0 -and @($answers | Where-Object { $_.State -eq 'None' }).Count -gt 0) {
            $notes.Add('DNS にはない（キャッシュ、NetBIOS、LLMNR の可能性）')
        } elseif ($osFound -and $okAnswers.Count -gt 0) {
            # hosts になく、OS の結果が DNS の答えにない。同じ種類の答えがあるものだけを比べる。
            # この PC 自身の名前は、OS がすべての NIC のアドレスを返すので比べない。リンクローカルのアドレスも DNS には載らないので比べない
            $isSelf = $selfNames.Contains($item.Name)
            $stale = $false
            if (-not $isSelf) {
                foreach ($ipText in $osIps) {
                    if ($ipText.StartsWith('fe80:', [StringComparison]::OrdinalIgnoreCase)) { continue }
                    $sameKind = @($dnsIps | Where-Object { (Test-IPv6Text $_) -eq (Test-IPv6Text $ipText) })
                    if ($sameKind.Count -gt 0 -and $sameKind -notcontains $ipText) { $stale = $true }
                }
            }
            if ($stale) {
                $attention = $true
                # -Server で聞いた先は、OS が使う DNS サーバーとは別のことがあるので、キャッシュが原因とは言えない
                if ($serverTexts.Count -gt 0) { $notes.Add('OS の結果と、指定した DNS サーバーの答えが違う') }
                else { $notes.Add('キャッシュが古い可能性（ipconfig /flushdns で確かめてください）') }
            }
        }

        # DNS サーバーごとの答えを比べる
        $answerKeys = @($answers | Where-Object { $_.State -eq 'Ok' -or $_.State -eq 'None' } | ForEach-Object { $_.Ips -join ',' } | Select-Object -Unique)
        if ($answerKeys.Count -gt 1) {
            $attention = $true
            $notes.Add('DNS サーバーによって答えが違う')
        }
        foreach ($answer in $answers) {
            if ($answer.State -eq 'NoResponse') {
                $attention = $true
                $notes.Add(('DNS サーバー {0} が応答しない' -f $answer.Server))
            } elseif ($answer.State -eq 'Skipped') {
                $attention = $true
                $notes.Add(('DNS サーバー {0} は応答しないので聞いていない' -f $answer.Server))
            } elseif ($answer.State -eq 'Error') {
                $attention = $true
                $notes.Add(('DNS サーバー {0}: {1}' -f $answer.Server, $answer.Detail))
            } elseif ($answer.State -eq 'Ok' -and $answer.Warning) {
                # A だけが答えて AAAA が失敗したときなど
                $attention = $true
                $notes.Add(('DNS サーバー {0}: {1}' -f $answer.Server, $answer.Warning))
            }
        }

        # 期待する IP アドレス。OS の結果と、答えたすべての DNS サーバーの両方と比べる
        $expected = @($item.Expected)
        if ($expected.Count -gt 0) {
            if ($osFound -and -not (Test-ExpectedMatch $osIps $expected)) {
                $attention = $true
                $notes.Add('期待する IP アドレスと違う（OS の結果）')
            }
            foreach ($answer in $answers) {
                if (($answer.State -eq 'Ok' -or $answer.State -eq 'None') -and -not (Test-ExpectedMatch @($answer.Ips) $expected)) {
                    $attention = $true
                    $notes.Add(('期待する IP アドレスと違う（DNS サーバー {0}）' -f $answer.Server))
                }
            }
        }

        # 逆引き
        $reverseParts = New-Object 'System.Collections.Generic.List[string]'
        if (($Reverse -or $inputAddress) -and $osFound) {
            $resolvedList = New-Object 'System.Collections.Generic.List[string]'
            $resolvedList.Add($item.Name)
            foreach ($answer in $okAnswers) { if ($answer.Resolved) { $resolvedList.Add($answer.Resolved) } }
            foreach ($ipText in $osIps) {
                $reverseNames = Get-ReverseName $ipText
                if ($null -eq $reverseNames) {
                    $attention = $true
                    $reverseParts.Add(('{0}: 聞けなかった' -f $ipText))
                    $notes.Add(('逆引きを聞けなかった（{0}）' -f $ipText))
                } elseif ($reverseNames.Count -eq 0) {
                    $reverseParts.Add(('{0}: なし' -f $ipText))
                    $notes.Add(('逆引きの登録がない（{0}）' -f $ipText))
                } elseif ($inputAddress) {
                    # IP アドレスを調べたときは、逆引きの名前がいちばん知りたい結果なので、画面にも出す
                    $reverseParts.Add(('{0}: {1}' -f $ipText, ($reverseNames -join ', ')))
                    $notes.Add(('逆引き: {0}' -f ($reverseNames -join ', ')))
                } else {
                    $isMatch = $false
                    foreach ($reverseName in $reverseNames) {
                        foreach ($resolvedText in $resolvedList) { if ($reverseName -ieq $resolvedText) { $isMatch = $true } }
                    }
                    if ($isMatch) {
                        $reverseParts.Add(('{0}: 一致' -f $ipText))
                    } else {
                        $attention = $true
                        $reverseParts.Add(('{0}: 違う（{1}）' -f $ipText, ($reverseNames -join ', ')))
                        $notes.Add(('逆引きの名前が違う（{0} → {1}）' -f $ipText, ($reverseNames -join ', ')))
                    }
                }
            }
        }

        # 判定
        $verdict = 'OK'
        if (-not $osFound) {
            $verdict = '引けない'
            if ($okAnswers.Count -gt 0) { $notes.Add('DNS には登録があるが、OS では引けない') }
            elseif (@($answers | Where-Object { $_.State -eq 'None' }).Count -gt 0 -and $hostsEntries.Count -eq 0) { $notes.Add(('DNS: {0}' -f @($answers | Where-Object { $_.State -eq 'None' })[0].Detail)) }
            elseif ($answers.Count -eq 0 -and $osResult.Detail) { $notes.Add($osResult.Detail) }
        } elseif ($attention) {
            $verdict = '注意'
        }

        # 表示の文字
        $dnsDisplay = '-'
        if ($okAnswers.Count -gt 0) {
            $dnsDisplay = if ($answerKeys.Count -gt 1) { 'サーバーで違う' } else { Join-Ip $okAnswers[0].Ips }
        }
        $dnsCsvParts = @($answers | ForEach-Object {
            $valueText = if ($_.State -eq 'Ok') { Join-Ip $_.Ips } else { $_.Detail }
            if ($_.State -eq 'Ok' -and $_.Warning) { $valueText += '（{0}）' -f $_.Warning }
            '{0}: {1}' -f $_.Server, $valueText
        })
        $firstOk = if ($okAnswers.Count -gt 0) { $okAnswers[0] } else { $null }
        $ttlValues = @($okAnswers | Where-Object { $_.Ttl -ge 0 } | ForEach-Object { $_.Ttl } | Sort-Object)

        $results.Add([pscustomobject]@{
            Verdict = $verdict; Name = (ConvertTo-PlainName $item.Name); Source = $sourceText
            OsText = $(if ($osFound) { Join-Ip $osIps } else { '-' })
            HostsText = (Join-Ip $hostsIps); HostsLine = $hostsLineText
            DnsDisplay = $dnsDisplay; DnsCsv = (ConvertTo-PlainName ($dnsCsvParts -join ' / '))
            Resolved = $(if ($firstOk) { ConvertTo-PlainName $firstOk.Resolved } else { '' })
            Cname = $(if ($firstOk) { ConvertTo-PlainName ($firstOk.Chain -join ' → ') } else { '' })
            Ttl = $(if ($ttlValues.Count -gt 0) { ([long]$ttlValues[0]).ToString('N0', $Invariant) } else { '' })
            Expected = ($expected -join ';'); ReverseText = (ConvertTo-PlainName ($reverseParts -join ' / '))
            Memo = $item.Memo; Note = (ConvertTo-PlainName ($notes -join '、'))
        })
    }

    #endregion

    #region 結果を組み立てる

    $countOk   = @($results | Where-Object { $_.Verdict -eq 'OK' }).Count
    $countWarn = @($results | Where-Object { $_.Verdict -eq '注意' }).Count
    $countFail = @($results | Where-Object { $_.Verdict -eq '引けない' }).Count
    if ($countWarn + $countFail -gt 0) { $exitCode = 2 }

    $lines.Add(('名前解決確認 (Test-NameResolution {0})' -f $ScriptVersion))
    $lines.Add((Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（管理者: {2}）' $startTime $machineName $(if ($isAdmin) { 'はい' } else { 'いいえ' })))
    if ($configPath) { $lines.Add(('設定ファイル: {0}' -f $configPath)) }
    $serverLineText = if ($dnsServers.Count -gt 0) { '{0}{1}' -f ($dnsServers -join ', '), $serverListText } else { 'なし' }
    $lines.Add((Format-Invariant '対象 {0:N0} 件   DNS サーバー: {1}' $items.Count $serverLineText))
    if ($hostsError) {
        $lines.Add(('※ hosts を読めなかったため、hosts なしとして調べました（{0}）' -f $hostsError))
    } else {
        $lines.Add((Format-Invariant 'hosts: {0}（有効な行 {1:N0} 件）' $hostsPath $hostsLines))
    }
    if ($serverError) { $lines.Add(('※ DNS サーバーの設定を読めませんでした（{0}）' -f $serverError)) }
    elseif ($dnsServers.Count -eq 0) { $lines.Add('※ DNS サーバーが設定されていません') }
    $lines.Add('')
    $headCount = $lines.Count

    # 表の列の幅は、すべての行に合わせる（summary.txt と画面で同じ表を使うため）
    $widths = @{ Verdict = 8; Name = 4; Source = 8; Os = 9; Dns = 10 }
    foreach ($row in $results) {
        $widths.Name   = [math]::Max($widths.Name, (Get-DisplayWidth $row.Name))
        $widths.Source = [math]::Max($widths.Source, (Get-DisplayWidth $row.Source))
        $widths.Os     = [math]::Max($widths.Os, (Get-DisplayWidth $row.OsText))
        $widths.Dns    = [math]::Max($widths.Dns, (Get-DisplayWidth $row.DnsDisplay))
    }
    $tableHeader = (' ' + (Format-Column '判定' $widths.Verdict) + '  ' + (Format-Column '名前' $widths.Name) + '  ' + (Format-Column '引けた元' $widths.Source) + '  ' +
        (Format-Column 'OS の結果' $widths.Os) + '  ' + (Format-Column 'DNS の結果' $widths.Dns) + '  備考').TrimEnd()
    $tableLines = New-Object 'System.Collections.Generic.List[object]'
    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($row in $results) {
        $text = (' ' + (Format-Column $row.Verdict $widths.Verdict) + '  ' + (Format-Column $row.Name $widths.Name) + '  ' + (Format-Column $row.Source $widths.Source) + '  ' +
            (Format-Column $row.OsText $widths.Os) + '  ' + (Format-Column $row.DnsDisplay $widths.Dns) + '  ' + $row.Note).TrimEnd()
        $color = switch ($row.Verdict) { '注意' { 'Yellow' } '引けない' { 'Red' } default { '' } }
        $tableLines.Add(@($text, $color))
        $rows.Add([object[]]@(
            $row.Verdict, (Protect-ExcelText $row.Name), $row.Source, (Protect-ExcelText $(if ($row.OsText -eq '-') { '' } else { $row.OsText })), (Protect-ExcelText $row.HostsText), (Protect-ExcelText $row.HostsLine),
            (Protect-ExcelText $row.DnsCsv), (Protect-ExcelText $row.Resolved), (Protect-ExcelText $row.Cname), $row.Ttl, (Protect-ExcelText $row.Expected),
            (Protect-ExcelText $row.ReverseText), (ConvertTo-ExcelLiteral $row.Memo), (Protect-ExcelText $row.Note)))
    }
    $elapsed = (Get-Date) - $startTime
    $resultLine = Format-Invariant '結果: OK {0:N0} 件  注意 {1:N0} 件  引けない {2:N0} 件  所要 {3}' $countOk $countWarn $countFail (Format-Duration $elapsed)

    # summary.txt には、すべて書く
    $lines.Add($tableHeader)
    foreach ($entry in $tableLines) { $lines.Add($entry[0]) }
    $lines.Add('')
    $lines.Add($resultLine)
    if ($runDir) { $lines.Add(('保存先: {0}' -f $runDir)) }

    #endregion

    #region 保存

    if ($runDir) {
        Write-Progress -Activity '名前解決確認' -Status '結果を保存しています'
        try {
            [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'results.csv')) $ResultColumns $rows
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
    Write-Progress -Activity '名前解決確認' -Completed
    if (-not $finished -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

#region 画面表示

for ($i = 0; $i -lt $headCount; $i++) { Write-Host $lines[$i] }
Write-Host $tableHeader
# 保存できたときは、画面には先頭の分だけを出す。保存しなかったときは、ほかに見る場所がないためすべて出す
$limit = if ($runDir -and -not $saveError) { $ScreenRows } else { $tableLines.Count }
$shown = 0
foreach ($entry in $tableLines) {
    if ($shown -ge $limit) { break }
    if ($entry[1]) { Write-Host $entry[0] -ForegroundColor $entry[1] } else { Write-Host $entry[0] }
    $shown++
}
if ($tableLines.Count -gt $shown) {
    Write-Host (Format-Invariant '（ほか {0:N0} 件は results.csv を参照してください）' ($tableLines.Count - $shown))
}
Write-Host ''
Write-Host $resultLine

if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
} elseif ($saveWarning) {
    Write-Warning $saveWarning
}

#endregion

exit $exitCode
