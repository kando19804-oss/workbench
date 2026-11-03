<#
.SYNOPSIS
    証明書の有効期限を調べて、期限切れや期限の近いものを洗い出します。

.DESCRIPTION
    Windows の証明書ストア、接続先のサーバーが使っている証明書、フォルダの中の証明書ファイルの有効期限を調べ、
    期限切れと期限間近の証明書を画面に表示し、すべての証明書を CSV（BOM 付き UTF-8）に保存します。
    この PC に IIS があれば、証明書を使っているサイトも調べます。

    証明書を追加・更新・削除することはありません。
    権限がなくて読めないストアやフォルダはスキップし、件数を表示します。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Get-CertificateExpiry.ps1

.PARAMETER Days
    「期限間近」とする残り日数（1～3650）。既定は 60 です。

.PARAMETER Store
    調べるストア。LocalMachine\My の形で、カンマで区切って書きます。
    省略すると、LocalMachine の My・WebHosting・Remote Desktop と CurrentUser の My を調べます。

.PARAMETER AllStores
    この PC と実行したユーザーの、すべてのストアを調べます。-Store と一緒には使えません。

.PARAMETER Target
    接続して証明書を調べるサーバー。ホスト か ホスト:ポート を、カンマで区切って書きます。ポートを省くと 443 です。
    -Target か -Path だけを指定すると、ストアは調べません。

.PARAMETER Timeout
    -Target の 1 件あたりの待ち時間（秒、1～60）。既定は 5 です。

.PARAMETER Path
    調べる証明書ファイル、またはフォルダ。カンマで区切って複数指定できます。フォルダは下の階層まで調べます。

.PARAMETER Exclude
    調べない証明書。サブジェクトに対して、ワイルドカード（*）で書きます。カンマで区切って複数指定できます。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.PARAMETER ConfigFile
    設定ファイルのパス。Days、Store、AllStores、Target、Timeout、Path、Exclude、OutputFolder、NoSave の値を
    「項目 = 値」の形で書いておけます。コマンドラインで指定した値は、設定ファイルより優先します。
    書き方は README を参照してください。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-CertificateExpiry.ps1 -Days 90

    運用で使うストアについて、残り 90 日以内の証明書を洗い出します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-CertificateExpiry.ps1 -Target "web01,lb01.example.co.jp:8443"

    接続先のサーバーが使っている証明書の期限を調べます。

.NOTES
    Get-CertificateExpiry 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [ValidateRange(1, 3650)]
    [int]$Days = 60,

    [string[]]$Store = @(),

    [switch]$AllStores,

    [string[]]$Target = @(),

    [ValidateRange(1, 60)]
    [int]$Timeout = 5,

    [string[]]$Path = @(),

    [string[]]$Exclude = @(),

    [string]$OutputFolder,

    [switch]$NoSave,

    [string]$ConfigFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion  = '1.0.0'
$ScreenRows     = 200       # 画面に表示する証明書の件数（保存するときだけ。超えた分はファイルを見てもらう）
$ProgressMs     = 250       # 進捗表示の間隔（ミリ秒）
$MaxCertFile    = 10MB      # 証明書として読むファイルの大きさの上限（間違えて巨大なファイルを指定したときに備える）
$OutputPrefix   = 'CertificateExpiry_'
$DefaultStores  = 'LocalMachine\My', 'LocalMachine\WebHosting', 'LocalMachine\Remote Desktop', 'CurrentUser\My'
$CertExtensions = '.cer', '.crt', '.pem', '.der'
$SecretExtensions = '.pfx', '.p12'
$ExcludeSeparator = ',(?! )'  # -Exclude の区切り。サブジェクトの「, 」では区切らない
$CertColumns    = '状態', '残り日数', '有効期限', '有効開始日', '種類', '場所', 'サブジェクト', '別名(SAN)', '発行者', '用途', '秘密鍵', 'IIS のサイト', 'フレンドリ名', '拇印', '備考'
$SkipColumns    = '理由', '対象', '詳細'
$TimeFormat     = 'yyyy/MM/dd HH:mm:ss'
# 拡張キー使用法の表示名。Oid.FriendlyName は OS の言語で変わるため、自分で持つ
$UsageNames     = @{
    '1.3.6.1.5.5.7.3.1' = 'サーバー認証'; '1.3.6.1.5.5.7.3.2' = 'クライアント認証'; '1.3.6.1.5.5.7.3.3' = 'コード署名'
    '1.3.6.1.5.5.7.3.4' = '電子メールの保護'; '1.3.6.1.5.5.7.3.8' = 'タイムスタンプ'; '1.3.6.1.5.5.7.3.9' = 'OCSP 署名'
    '1.3.6.1.4.1.311.10.3.4' = '暗号化ファイルシステム'; '1.3.6.1.4.1.311.20.2.2' = 'スマートカード ログオン'
}
$HrFileNotFound  = -2147024894   # 0x80070002（ストアがない）
$HrAccessDenied  = -2147024891   # 0x80070005
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
    $quoted = $Text.Replace('"', '""')
    if ($quoted.Length -le 255) { return '="' + $quoted + '"' }
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
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -Path "C:\data"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# ストアの指定を確かめて、場所と名前に分ける。正しくなければ $null を返す
function ConvertFrom-StoreText([string]$Text) {
    $found = [regex]::Match($Text, '^(LocalMachine|CurrentUser)\\(.+)$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $found.Success) { return $null }
    $place = if ($found.Groups[1].Value -ieq 'LocalMachine') { 'LocalMachine' } else { 'CurrentUser' }
    return @{ Place = $place; Name = $found.Groups[2].Value; Text = $place + '\' + $found.Groups[2].Value }
}

# ホスト、ホスト:ポート、[IPv6]:ポート の形を確かめて、ホストとポートに分ける。ポートが正しくなければ $null を返す
function ConvertFrom-TargetText([string]$Text) {
    $hostText = $Text
    $portText = ''
    $bracket = [regex]::Match($Text, '^\[([^\]]+)\](?::(.*))?$')
    if ($bracket.Success) {
        $hostText = $bracket.Groups[1].Value
        if ($bracket.Groups[2].Success) { $portText = $bracket.Groups[2].Value; if ($portText -eq '') { return $null } }
    } elseif (([regex]::Matches($Text, ':')).Count -eq 1) {
        $hostText = $Text.Substring(0, $Text.IndexOf([char]':'))
        $portText = $Text.Substring($Text.IndexOf([char]':') + 1)
        if ($portText -eq '') { return $null }
    } elseif ($Text.Contains(':')) {
        # コロンが 2 つ以上あるのは、括弧なしの IPv6 アドレスだけ
        $parsedAddress = $null
        if (-not [Net.IPAddress]::TryParse($Text, [ref]$parsedAddress)) { return $null }
    }
    if ($hostText -eq '') { return $null }
    $port = 443
    if ($portText -and (-not [int]::TryParse($portText, [Globalization.NumberStyles]::None, $Invariant, [ref]$port) -or $port -lt 1 -or $port -gt 65535)) { return $null }
    $display = if ($hostText.Contains(':')) { '[{0}]:{1}' -f $hostText, $port } else { '{0}:{1}' -f $hostText, $port }
    return @{ Host = $hostText; Port = $port; Text = $display }
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの。数値の範囲は param の ValidateRange と同じにする
$configItems   = 'Days', 'Store', 'AllStores', 'Target', 'Timeout', 'Path', 'Exclude', 'OutputFolder', 'NoSave'
$configRanges  = @{ Days = 1, 3650; Timeout = 1, 60 }
$configFlags   = 'AllStores', 'NoSave'
$configLists   = 'Store', 'Target', 'Exclude'
$configPath    = $null
$configApplied = @{}   # 設定ファイルの値を使った項目 → 行番号
$configPathLines = @{} # 設定ファイルの Path の値 → 行番号（Path は複数の行に書けるため）
$configStoreLines = @{} # 設定ファイルの Store の値 → 行番号

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
    $pathItems    = New-Object 'System.Collections.Generic.List[string]'
    $lineNumber   = 0
    foreach ($rawLine in $configLines) {
        $lineNumber++
        $lineText = $rawLine.Trim()
        # # と ; はフォルダ名に使える文字なので、行の先頭にあるときだけコメントとみなす
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

        # Store・Target・Path・Exclude は複数の行に分けて書ける。ほかの項目は、どちらが効いているか分からなくなるため 1 回だけ
        if ($itemName -ne 'Path' -and $configLists -notcontains $itemName -and $configLineOf.ContainsKey($itemName)) {
            Stop-WithConfigError $lineNumber ('{0} が 2 回書かれています（{1} 行目と {2} 行目）' -f $itemName, $configLineOf[$itemName], $lineNumber)
        }
        $configLineOf[$itemName] = $lineNumber

        if ($configLists -contains $itemName) {
            # "a,b" のように全体を囲む書き方と、"a","b" のように 1 つずつ囲む書き方のどちらも受け付ける
            $source = if ($itemValue.Contains('"')) { $rawValue } else { $itemValue }
            # サブジェクトは「CN=a, O=b」のようにカンマとスペースを含むので、Exclude ではスペースが続くカンマを区切りにしない
            $separatorRule = if ($itemName -eq 'Exclude') { $ExcludeSeparator } else { ',' }
            foreach ($part in ($source -split $separatorRule)) {
                $pattern = Remove-ConfigQuote $part
                if ($pattern.Contains('"')) { Stop-WithConfigError $lineNumber ('{0} の値に引用符（"）が残っています: {1}' -f $itemName, $pattern) }
                if (-not $pattern) { continue }
                if ($itemName -ne 'Store') { $pattern = Expand-ConfigVariable $lineNumber $pattern }
                if ($itemName -eq 'Store') {
                    $parsedStore = ConvertFrom-StoreText $pattern
                    if (-not $parsedStore) { Stop-WithConfigError $lineNumber ('Store は LocalMachine\My の形で指定してください: {0}' -f $pattern) }
                    $configStoreLines[$parsedStore.Text] = $lineNumber
                }
                if ($itemName -eq 'Target' -and -not (ConvertFrom-TargetText $pattern)) { Stop-WithConfigError $lineNumber ('Target は ホスト か ホスト:ポート（ポートは 1～65535）の形で指定してください: {0}' -f $pattern) }
                $listItems[$itemName].Add($pattern)
            }
        } elseif ($configFlags -contains $itemName) {
            if ($itemValue -eq 'true') { $configValues[$itemName] = $true }
            elseif ($itemValue -eq 'false') { $configValues[$itemName] = $false }
            else { Stop-WithConfigError $lineNumber ('{0} は true か false で指定してください: {1}' -f $itemName, $itemValue) }
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
            if ($itemName -eq 'Path') {
                # カンマを含む名前のフォルダも書けるよう、1 行を 1 つの場所として扱う
                $pathItems.Add($itemValue)
                $configPathLines[$itemValue] = $lineNumber
            } else {
                $configValues[$itemName] = $itemValue
            }
        }
    }
    foreach ($itemName in $configLists) {
        if ($listItems[$itemName].Count -gt 0) { $configValues[$itemName] = $listItems[$itemName].ToArray() }
    }
    if ($pathItems.Count -gt 0) { $configValues['Path'] = $pathItems.ToArray() }

    # コマンドラインで指定した項目は、設定ファイルより優先する
    foreach ($itemName in @($configValues.Keys)) {
        if (-not $PSBoundParameters.ContainsKey($itemName)) {
            # Store と AllStores は一緒に使えないので、片方をコマンドラインで指定したら、設定ファイルのもう片方は使わない
            if ($itemName -eq 'Store' -and $PSBoundParameters.ContainsKey('AllStores')) { continue }
            if ($itemName -eq 'AllStores' -and $PSBoundParameters.ContainsKey('Store')) { continue }
            Set-Variable -Name $itemName -Value $configValues[$itemName]
            $configApplied[$itemName] = $configLineOf[$itemName]
        }
    }
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

# -Store・-Target・-Exclude は、-File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# コマンドプロンプトでは ' がそのまま届くため、'a,b' のように全体を囲んだ場合も、各要素の前後の ' を外す
function Get-ListText([string[]]$Values, [string]$Name, [string]$Rule = ',') {
    $texts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($value in $Values) {
        foreach ($part in ($value -split $Rule)) {
            $text = (Remove-Quote $part).Trim("'")
            if ($text.Contains('"')) { Stop-WithQuoteError $Name $text }
            if ($text -and -not ($texts -contains $text)) { $texts.Add($text) }
        }
    }
    return , $texts
}

$storeItems = New-Object 'System.Collections.Generic.List[object]'
$storeNamed = $false
foreach ($text in (Get-ListText $Store '-Store')) {
    $parsed = ConvertFrom-StoreText $text
    if (-not $parsed) { Stop-WithError ('-Store は LocalMachine\My の形で指定してください: {0}' -f $text) }
    $storeItems.Add($parsed)
    $storeNamed = $true
}
if ($storeNamed -and $AllStores) { Stop-WithError '-Store と -AllStores は一緒に指定できません' }

$targetItems = New-Object 'System.Collections.Generic.List[object]'
foreach ($text in (Get-ListText $Target '-Target')) {
    $parsed = ConvertFrom-TargetText $text
    if (-not $parsed) { Stop-WithError ('-Target は ホスト か ホスト:ポート（ポートは 1～65535）の形で指定してください: {0}' -f $text) }
    if (@($targetItems | Where-Object { $_.Text -ieq $parsed.Text }).Count -eq 0) { $targetItems.Add($parsed) }
}

# -Path は、-File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# カンマを含む名前のフォルダのため、全体が 1 つの場所として存在するときは区切らない。設定ファイルの値は 1 行が 1 つの場所
$pathTexts = New-Object 'System.Collections.Generic.List[string]'
foreach ($item in $Path) {
    if ($configApplied.ContainsKey('Path')) { $pathTexts.Add($item); continue }
    $whole = Remove-Quote $item
    if ($whole.Contains('"')) { Stop-WithQuoteError '-Path' $whole }
    $isWhole = $false
    try {
        if ($whole.Contains(',')) {
            $wholeLong = ConvertTo-LongPath (Resolve-FullPath (Resolve-StartPath $whole))
            $isWhole = [IO.Directory]::Exists($wholeLong) -or [IO.File]::Exists($wholeLong)
        }
    } catch { }
    if ($isWhole) { $pathTexts.Add($whole); continue }
    foreach ($part in ($whole -split ',')) {
        $text = (Remove-Quote $part).Trim("'")
        if ($text) { $pathTexts.Add($text) }
    }
}
$pathTargets = New-Object 'System.Collections.Generic.List[object]'
foreach ($text in $pathTexts) {
    $note = ''
    if ($configPathLines.ContainsKey($text) -and $configApplied.ContainsKey('Path')) { $note = '（設定ファイルの {0} 行目）' -f $configPathLines[$text] }
    $pathText = $text
    if ($pathText -match '^(\\\\\?\\)?[A-Za-z]:$') { $pathText += '\' }
    $pathFull = Resolve-FullPath (Resolve-StartPath $pathText)
    $isFolder = $false
    $isFile   = $false
    try {
        $isFolder = [IO.Directory]::Exists((ConvertTo-LongPath $pathFull))
        $isFile   = [IO.File]::Exists((ConvertTo-LongPath $pathFull))
    } catch { }
    if (-not $isFolder -and -not $isFile) { Stop-WithError (('ファイルかフォルダが見つかりません: {0}' -f $pathFull) + $note) }
    $pathTargets.Add(@{ Full = $pathFull; IsFolder = $isFolder })
}

# -Target か -Path だけを指定したときは、この PC のストアを調べない（手元の証明書が混ざらないように）
$useDefaultStores = -not $storeNamed -and -not $AllStores -and $targetItems.Count -eq 0 -and $pathTargets.Count -eq 0
$useStores = $storeNamed -or $AllStores -or $useDefaultStores
if ($useDefaultStores) { foreach ($text in $DefaultStores) { $storeItems.Add((ConvertFrom-StoreText $text)) } }

# ワイルドカードは * だけにして、それ以外の文字は文字どおりに扱う。5.1 の WildcardPattern.Escape は
# バッククォートをエスケープしないため、自分でエスケープしてから照合用のパターンを作る
$wildcardOptions = [Management.Automation.WildcardOptions]'IgnoreCase, CultureInvariant'
$excludeTexts = Get-ListText $Exclude '-Exclude' $ExcludeSeparator
$excludeRules = @(
    foreach ($text in $excludeTexts) {
        $escaped = $text.Replace('`', '``').Replace('[', '`[').Replace(']', '`]').Replace('?', '`?')
        [pscustomobject]@{ Text = $text; Matcher = [Management.Automation.WildcardPattern]::new($escaped, $wildcardOptions) }
    }
)

#endregion

#region 判定の部品

$certType   = [Security.Cryptography.X509Certificates.X509Certificate2]
$dirType    = [IO.DirectoryInfo]
$rpFlag     = [IO.FileAttributes]::ReparsePoint
$hsFlag     = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System

$startTime  = Get-Date
$certs      = New-Object 'System.Collections.Generic.List[object]'
$skipped    = New-Object 'System.Collections.Generic.List[object]'
$skipKinds  = [ordered]@{ '権限なし' = 0; '接続できない' = 0; '読めない' = 0; 'その他のエラー' = 0 }
$excludedCount = 0
$linkCount     = 0
$secretCount   = 0
$readFiles  = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)   # -Path の指定が重なっても、同じファイルを二重に数えない
$progress   = @{ Next = 0L }
$sw         = [Diagnostics.Stopwatch]::StartNew()

function Add-Skip([string]$Reason, [string]$Subject, [string]$Detail) {
    $skipped.Add([object[]]@($Reason, $Subject, $Detail))
    if ($skipKinds.Contains($Reason)) { $skipKinds[$Reason]++ }
}

function Show-Progress([string]$Status) {
    if ($sw.ElapsedMilliseconds -lt $progress.Next) { return }
    Write-Progress -Activity '証明書期限一覧' -Status $Status
    $progress.Next = $sw.ElapsedMilliseconds + $ProgressMs
}

# 別名(SAN) の拡張を読む。Format() の文字は OS の言語で変わるため、DER を自分で読む。
# 中身は GeneralName の並びで、DNS 名は [2]（0x82）、IP アドレスは [7]（0x87）
function Get-SanNames($Cert) {
    $names = New-Object 'System.Collections.Generic.List[string]'
    foreach ($extension in $Cert.Extensions) {
        if ($extension.Oid.Value -ne '2.5.29.17') { continue }
        $raw = $extension.RawData
        # 長さを読み、$script:derIndex を中身の先頭に進める
        $readLength = {
            $first = [int]$raw[$script:derIndex]
            $script:derIndex++
            if ($first -lt 0x80) { return $first }
            $total = 0
            for ($n = 0; $n -lt ($first -band 0x7F); $n++) { $total = $total * 256 + [int]$raw[$script:derIndex]; $script:derIndex++ }
            return $total
        }
        try {
            $script:derIndex = 1
            [void](& $readLength)
            while ($script:derIndex -lt $raw.Length) {
                $tag = [int]$raw[$script:derIndex]
                $script:derIndex++
                $length = & $readLength
                $i = $script:derIndex
                if ($tag -eq 0x82) {
                    $names.Add([Text.Encoding]::ASCII.GetString($raw, $i, $length))
                } elseif ($tag -eq 0x87 -and ($length -eq 4 -or $length -eq 16)) {
                    $address = New-Object 'byte[]' $length
                    [Array]::Copy($raw, $i, $address, 0, $length)
                    $names.Add(([Net.IPAddress]::new($address)).ToString())
                }
                $script:derIndex = $i + $length
            }
        } catch { }
    }
    return , $names
}

function Get-UsageText($Cert) {
    $usages = @()
    foreach ($extension in $Cert.Extensions) {
        if ($extension -isnot [Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]) { continue }
        foreach ($oid in $extension.EnhancedKeyUsages) {
            $usages += if ($UsageNames.ContainsKey($oid.Value)) { $UsageNames[$oid.Value] } else { $oid.Value }
        }
    }
    return $usages -join ', '
}

# 証明書を一覧に加える。$Scope は「更新済み」を比べる範囲（同じストア、同じフォルダ）で、接続先は $null
function Add-CertRecord($Cert, [string]$Kind, [string]$Place, $Scope, [string]$KeyText, [string]$Note) {
    $subject = $Cert.Subject
    foreach ($rule in $excludeRules) {
        if ($rule.Matcher.IsMatch($subject)) {
            $script:excludedCount++
            Add-Skip '除外指定' ('{0}: {1}' -f $Place, $subject) (Protect-ExcelText $rule.Text)
            return
        }
    }
    $sanNames = Get-SanNames $Cert
    $sanSorted = New-Object 'System.Collections.Generic.List[string]'
    foreach ($sanName in $sanNames) { $sanSorted.Add($sanName.ToLowerInvariant()) }
    $sanSorted.Sort([StringComparer]::Ordinal)
    $friendly = ''
    try { $friendly = $Cert.FriendlyName } catch { }
    $certs.Add([pscustomobject]@{
        Status = ''; DaysLeft = 0; NotAfter = $Cert.NotAfter; NotBefore = $Cert.NotBefore; Kind = $Kind; Place = $Place; Scope = $Scope
        Subject = $subject; San = ($sanNames -join ', '); SanKey = ($sanSorted -join ','); Issuer = $Cert.Issuer; Usage = (Get-UsageText $Cert)
        KeyText = $KeyText; Iis = ''; IisNames = ''; Friendly = $friendly; Thumbprint = $Cert.Thumbprint; Note = $Note
    })
}

# ストアを開く。開けたらストアを返す。ないときは $null、ほかの理由で開けないときは理由の文字列を返す
function Open-CertStore($Item) {
    $certStore = New-Object Security.Cryptography.X509Certificates.X509Store($Item.Name, [Security.Cryptography.X509Certificates.StoreLocation]$Item.Place)
    try {
        $certStore.Open([Security.Cryptography.X509Certificates.OpenFlags]'ReadOnly, OpenExistingOnly')
        return $certStore
    } catch {
        $ex = Get-InnerException $_.Exception
        if ($ex.HResult -eq $HrFileNotFound -or -not $OnWindows) { return $null }
        return $ex
    }
}

function Read-CertStore($Item) {
    Show-Progress ('ストアを調べています: {0}' -f $Item.Text)
    $opened = Open-CertStore $Item
    if ($null -eq $opened) { return }
    if ($opened -is [Exception]) {
        $reason = if ($opened.HResult -eq $HrAccessDenied -or $opened -is [UnauthorizedAccessException] -or $opened -is [Security.SecurityException]) { '権限なし' } else { 'その他のエラー' }
        Add-Skip $reason $Item.Text (Get-ErrorText $opened)
        return
    }
    try {
        foreach ($cert in $opened.Certificates) {
            try {
                $keyText = if ($cert.HasPrivateKey) { 'あり' } else { 'なし' }
                Add-CertRecord $cert 'ストア' $Item.Text $Item.Text $keyText ''
            } catch {
                Add-Skip 'その他のエラー' $Item.Text (Get-ErrorText (Get-InnerException $_.Exception))
            }
        }
    } finally {
        $opened.Close()
    }
}

# 接続先の証明書を受け取る。TLS の検証の結果は備考に書くだけで、接続は止めない
function Read-TargetCert($Item) {
    Show-Progress ('接続しています: {0}' -f $Item.Text)
    $client = $null
    $tls    = $null
    $script:tlsErrors = [Net.Security.SslPolicyErrors]::None
    $script:tlsChain  = @()
    try {
        # .NET Framework の引数なしの TcpClient は IPv4 専用なので、IPv4 と IPv6 の両方に接続できるソケットを作る。IPv6 を使えない PC では IPv4 だけにする
        try {
            $client = New-Object Net.Sockets.TcpClient ([Net.Sockets.AddressFamily]::InterNetworkV6)
            $client.Client.DualMode = $true
        } catch {
            if ($client) { $client.Close() }
            $client = New-Object Net.Sockets.TcpClient
        }
        $client.ReceiveTimeout = $Timeout * 1000
        $client.SendTimeout    = $Timeout * 1000
        $pending = $client.BeginConnect($Item.Host, $Item.Port, $null, $null)
        if (-not $pending.AsyncWaitHandle.WaitOne($Timeout * 1000)) { throw (New-Object TimeoutException ('{0} 秒以内に接続できませんでした' -f $Timeout)) }
        $client.EndConnect($pending)
        $callback = [Net.Security.RemoteCertificateValidationCallback] {
            param($Sender2, $Certificate, $Chain, $PolicyErrors)
            $script:tlsErrors = $PolicyErrors
            if ($Chain) { $script:tlsChain = @($Chain.ChainStatus | ForEach-Object { [string]$_.Status }) }
            return $true
        }
        $tls = New-Object Net.Security.SslStream($client.GetStream(), $false, $callback)
        # 5.1（.NET Framework）は、4.6.2 でも TLS 1.2 まで使えるよう明示する。TLS 1.3 は、.NET Framework 4.8 に名前はあっても、
        # OS が対応していないと接続に失敗することがあるので指定しない。7 は OS に任せる（None が OS の既定の意味）
        $protocols = if ($PSVersionTable.PSEdition -eq 'Core') { [Security.Authentication.SslProtocols]::None } else { [Security.Authentication.SslProtocols]'Tls, Tls11, Tls12' }
        $tls.AuthenticateAsClient($Item.Host, $null, $protocols, $false)
        $remote = New-Object Security.Cryptography.X509Certificates.X509Certificate2 ($tls.RemoteCertificate)

        # 期限切れは状態で分かるので、備考には書かない。失効を確かめられないだけのものも書かない
        $notes = @()
        if (($script:tlsErrors -band [Net.Security.SslPolicyErrors]::RemoteCertificateNameMismatch) -ne 0) { $notes += '名前が一致しない' }
        $otherChain = @($script:tlsChain | Where-Object { $_ -notin 'NoError', 'NotTimeValid', 'RevocationStatusUnknown', 'OfflineRevocation' })
        if ($otherChain.Count -gt 0) { $notes += '信頼できない発行者' }
        Add-CertRecord $remote '接続先' $Item.Text $null '' ($notes -join '、')
    } catch {
        Add-Skip '接続できない' $Item.Text (Get-ErrorText (Get-InnerException $_.Exception))
    } finally {
        if ($tls) { $tls.Dispose() }
        if ($client) { $client.Close() }
    }
}

# リパースポイントのうち、リンクと判定できたものの種類を返す（OneDrive のファイルやフォルダもリパースポイントなので、それ以外は $null）
function Get-LinkName($Entry) {
    if (($Entry.Attributes -band $rpFlag) -ne $rpFlag) { return $null }
    # Windows が互換用に作ったリンク（All Users、Application Data など）は、5.1 では種類を判定できない。どれも隠し・システム属性付き。
    # 7 は種類を判定できるが、5.1 と同じ結果になるよう、この目印を先に見る
    if ($Entry -is $dirType -and ($Entry.Attributes -band $hsFlag) -eq $hsFlag) { return 'Windows の互換用リンク' }
    $linkType = $null
    try { $linkType = $Entry.LinkType } catch { }
    if ($linkType -eq 'Junction') { return 'ジャンクション' }
    if ($linkType -eq 'SymbolicLink') { return 'シンボリックリンク' }
    return $null
}

# 証明書ファイルを読む。PEM は複数の証明書を入れられるので、すべて読む。秘密鍵の部分は読まない
function Read-CertFile([string]$LongPath) {
    $display = ConvertTo-DisplayPath $LongPath
    if (-not $readFiles.Add($display)) { return }
    Show-Progress ('ファイルを調べています: {0}' -f $display)
    $extension = ''
    $dot = $LongPath.LastIndexOf([char]'.')
    if ($dot -gt 0) { $extension = $LongPath.Substring($dot).ToLowerInvariant() }
    # パスワードが必要な形式は開かない
    if ($SecretExtensions -contains $extension) {
        $script:secretCount++
        Add-Skip '読まない形式' $display 'パスワードが必要な形式のため、開いていません'
        return
    }
    $parentEnd = $display.LastIndexOfAny([char[]]@('\', '/'))
    $scope = if ($parentEnd -gt 0) { $display.Substring(0, $parentEnd) } else { $display }
    try {
        $info = New-Object IO.FileInfo $LongPath
        if ($info.Length -gt $MaxCertFile) {
            Add-Skip '読めない' $display (Format-Invariant '{0:N0} バイトあり、証明書のファイルとしては大きすぎます' $info.Length)
            return
        }
        $bytes = [IO.File]::ReadAllBytes($LongPath)
    } catch {
        $ex = Get-InnerException $_.Exception
        $reason = if ($ex -is [UnauthorizedAccessException]) { '権限なし' } else { '読めない' }
        Add-Skip $reason $display (Get-ErrorText $ex)
        return
    }
    $text = [Text.Encoding]::GetEncoding(28591).GetString($bytes)   # PEM の区切りは ASCII なので、どの文字コードでも探せるように 1 バイトずつ文字にする
    $blocks = [regex]::Matches($text, '-----BEGIN (?:TRUSTED |X509 )?CERTIFICATE-----(.*?)-----END', [Text.RegularExpressions.RegexOptions]::Singleline)
    # 証明書の始まりだけがあって終わりがないのは、途中で切れたファイル。読めた分は調べて、切れていることを残す
    $beginCount = ([regex]::Matches($text, '-----BEGIN (?:TRUSTED |X509 )?CERTIFICATE-----')).Count
    $truncated  = $beginCount -gt $blocks.Count
    if ($truncated) { Add-Skip '読めない' $display '証明書の途中でファイルが切れています' }
    if ($truncated -and $blocks.Count -eq 0) { return }
    # 秘密鍵だけのファイル（Let's Encrypt の privkey.pem など）や、dhparam.pem、証明書要求（CSR）は、証明書ではないので読まない
    if ($blocks.Count -eq 0 -and $text -match '-----BEGIN ') {
        $script:secretCount++
        $reasonText = if ($text -match '-----BEGIN [A-Z ]*PRIVATE KEY-----') { '秘密鍵のファイルのため、読んでいません' } else { '証明書以外の PEM ファイルのため、読んでいません' }
        Add-Skip '読まない形式' $display $reasonText
        return
    }
    $sources = New-Object 'System.Collections.Generic.List[object]'
    try {
        if ($blocks.Count -gt 0) {
            foreach ($block in $blocks) { $sources.Add([Convert]::FromBase64String(($block.Groups[1].Value -replace '\s', ''))) }
        } else {
            $sources.Add($bytes)
        }
        foreach ($source in $sources) {
            # X509Certificate2 は、パスワードなしの PFX（秘密鍵を保存場所に書き込むことがある）や署名付きの実行ファイルも開いてしまうので、中身の形を先に確かめる
            $contentType = $certType::GetCertContentType([byte[]]$source)
            if ($contentType -ne [Security.Cryptography.X509Certificates.X509ContentType]::Cert) {
                $script:secretCount++
                Add-Skip '読まない形式' $display ('証明書 1 枚の形式ではないため、開いていません（{0}）' -f $contentType)
                continue
            }
            Add-CertRecord ($certType::new([byte[]]$source)) 'ファイル' $display $scope '' ''
        }
    } catch {
        Add-Skip '読めない' $display (Get-ErrorText (Get-InnerException $_.Exception))
    }
}

function Read-CertFolder([string]$FolderPath) {
    $stack = New-Object 'System.Collections.Generic.Stack[object]'
    $stack.Push((New-Object IO.DirectoryInfo (ConvertTo-LongPath $FolderPath)))
    $atRoot = $true
    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        Show-Progress ('フォルダを調べています: {0}' -f (ConvertTo-DisplayPath $dir.FullName))
        try {
            $entries = $dir.GetFileSystemInfos()
        } catch {
            $ex = Get-InnerException $_.Exception
            # 種類を判定できないリンクは、読み取りが禁止されていると普通のフォルダに見える。権限なしには数えない
            if (-not $atRoot -and $ex -is [UnauthorizedAccessException] -and ($dir.Attributes -band $rpFlag) -eq $rpFlag) {
                $script:linkCount++
                Add-Skip 'リンク' (ConvertTo-DisplayPath $dir.FullName) '読み取りが禁止されたリンク'
                continue
            }
            $reason = if ($ex -is [UnauthorizedAccessException]) { '権限なし' } else { 'その他のエラー' }
            Add-Skip $reason (ConvertTo-DisplayPath $dir.FullName) (Get-ErrorText $ex)
            continue
        }
        $atRoot = $false
        foreach ($entry in $entries) {
            if (($entry.Attributes -band $rpFlag) -eq $rpFlag) {
                $linkName = Get-LinkName $entry
                if ($linkName) {
                    $script:linkCount++
                    Add-Skip 'リンク' (ConvertTo-DisplayPath $entry.FullName) $linkName
                    continue
                }
            }
            if ($entry -is $dirType) { $stack.Push($entry); continue }
            $extension = $entry.Extension.ToLowerInvariant()
            if ($CertExtensions -contains $extension -or $SecretExtensions -contains $extension) { Read-CertFile $entry.FullName }
        }
    }
}

# IIS のサイトと証明書の紐付け（拇印 → サイトの一覧）を読む。IIS がなければ空のまま。読めなければ $false を返す
$iisSites = @{}
function Add-IisSite([string]$Hash, [string]$Name, [string]$Text) {
    $key = $Hash.ToUpperInvariant()
    if (-not $iisSites.ContainsKey($key)) { $iisSites[$key] = New-Object 'System.Collections.Generic.List[object]' }
    $iisSites[$key].Add(@{ Name = $Name; Text = $Text })
}

function Read-IisBinding {
    if (-not $env:windir) { return $true }
    # inetsrv のフォルダがなければ IIS はない。config のフォルダは管理者でないと読めないことがあるので、ここでは確かめず、開いてみて判断する
    # 64 ビットの OS の 32 ビットの PowerShell では、System32 が SysWOW64 に読み替えられて、IIS の設定が見えない
    $systemFolder = if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) { 'Sysnative' } else { 'System32' }
    $inetsrv = Join-Path $env:windir ($systemFolder + '\inetsrv')
    if (-not [IO.Directory]::Exists($inetsrv)) { return $true }
    $iisConfigPath = Join-Path $inetsrv 'config\applicationHost.config'
    $unknownBinding = $false
    try {
        # IIS が使っている最中でも読めるように開く
        try {
            $stream = New-Object IO.FileStream($iisConfigPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]'ReadWrite, Delete')
        } catch {
            $ex = Get-InnerException $_.Exception
            if ($ex -is [IO.FileNotFoundException] -or $ex -is [IO.DirectoryNotFoundException]) { return $true }
            return $false
        }
        try {
            $document = New-Object Xml.XmlDocument
            $document.XmlResolver = $null
            $document.Load($stream)
        } finally {
            $stream.Dispose()
        }
        # 共有構成では、本当の設定は別の場所にあり、ローカルのファイルは当てにならない
        $redirectPath = Join-Path $inetsrv 'config\redirection.config'
        if ([IO.File]::Exists($redirectPath)) {
            $redirect = New-Object Xml.XmlDocument
            $redirect.XmlResolver = $null
            $redirect.Load($redirectPath)
            $redirectNode = $redirect.SelectSingleNode('/configuration/configurationRedirection')
            if ($redirectNode -and $redirectNode.GetAttribute('enabled') -ieq 'true') { return $false }
        }
        foreach ($site in $document.SelectNodes('/configuration/system.applicationHost/sites/site')) {
            $siteName = $site.GetAttribute('name')
            foreach ($binding in $site.SelectNodes('bindings/binding')) {
                if ($binding.GetAttribute('protocol') -ine 'https') { continue }
                $info = $binding.GetAttribute('bindingInformation')
                $hash = $binding.GetAttribute('certificateHash')
                if (-not $hash) { $hash = Get-HttpSysHash $info }
                if (-not $hash) {
                    # 中央証明書ストア（sslFlags の 2 ビット）のバインドは、ストアの証明書を使わない。それ以外は、どの証明書を使っているか分からない
                    $flags = $binding.GetAttribute('sslFlags')
                    $central = $flags -match 'Central' -or ($flags -match '^[0-9]+$' -and ([int]$flags -band 2) -ne 0)
                    if (-not $central) { $unknownBinding = $true }
                    continue
                }
                Add-IisSite $hash $siteName ('{0} ({1})' -f $siteName, $info)
            }
            # FTP サイトの証明書
            $ftpSsl = $site.SelectSingleNode('ftpServer/security/ssl')
            if ($ftpSsl -and $ftpSsl.GetAttribute('serverCertHash')) {
                Add-IisSite $ftpSsl.GetAttribute('serverCertHash') $siteName ('{0} (FTP)' -f $siteName)
            }
        }
        return -not $unknownBinding
    } catch {
        return $false
    }
}

# バインドに証明書の情報がないときは、HTTP.sys の設定（レジストリ）から探す。
# バインドの形は <IP>:<ポート>:<ホスト名> で、SNI のものは <ホスト名>:<ポート>、それ以外は <IP>:<ポート> に登録されている
function Get-HttpSysHash([string]$Info) {
    $found = [regex]::Match($Info, '^(.*):([0-9]+):(.*)$')
    if (-not $found.Success) { return $null }
    $address = $found.Groups[1].Value
    $port    = $found.Groups[2].Value
    $hostText = $found.Groups[3].Value
    $keys = @()
    if ($hostText) { $keys += 'SslSniBindingInfo\{0}:{1}' -f $hostText, $port }
    $keys += 'SslBindingInfo\{0}:{1}' -f $(if ($address -eq '*' -or $address -eq '') { '0.0.0.0' } else { $address }), $port
    foreach ($key in $keys) {
        $registry = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SYSTEM\CurrentControlSet\Services\HTTP\Parameters\' + $key)
        if (-not $registry) { continue }
        try {
            $value = $registry.GetValue('SslCertHash')
            if ($value) { return ([BitConverter]::ToString([byte[]]$value)).Replace('-', '') }
        } finally {
            $registry.Close()
        }
    }
    return $null
}

function Format-CertTime($Time) { return ([datetime]$Time).ToString($TimeFormat, $Invariant) }

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

    if ($AllStores) {
        # X509Store にはストアの一覧を得る手段がないため、Cert: ドライブから名前だけを読む
        foreach ($place in 'LocalMachine', 'CurrentUser') {
            try {
                foreach ($entry in @(Get-ChildItem -LiteralPath ('Cert:\' + $place))) {
                    # UserDS は Active Directory に問い合わせるストアなので、ネットワークに接続しないよう調べない
                    if ($entry.Name -ieq 'UserDS') {
                        Add-Skip '除外指定' ($place + '\' + $entry.Name) 'Active Directory に問い合わせるストアのため'
                        continue
                    }
                    $storeItems.Add((ConvertFrom-StoreText ($place + '\' + $entry.Name)))
                }
            } catch {
                Add-Skip 'その他のエラー' $place (Get-ErrorText (Get-InnerException $_.Exception))
            }
        }
    }
    # 打ち間違いに気づかないまま「問題なし」と出ないよう、-Store で指定したストアがないときは調べる前に止める
    if ($storeNamed) {
        foreach ($item in $storeItems) {
            $opened = Open-CertStore $item
            if ($null -eq $opened) {
                $note = if ($configApplied.ContainsKey('Store')) { '（設定ファイルの {0} 行目）' -f $configStoreLines[$item.Text] } else { '' }
                if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
                $runDir = $null
                $finished = $true
                Stop-WithError (('ストアが見つかりません: {0}' -f $item.Text) + $note)
            }
            if ($opened -isnot [Exception]) { $opened.Close() }
        }
    }

    foreach ($item in $storeItems) { Read-CertStore $item }
    foreach ($item in $targetItems) { Read-TargetCert $item }
    foreach ($item in $pathTargets) {
        if ($item.IsFolder) { Read-CertFolder $item.Full } else { Read-CertFile (ConvertTo-LongPath $item.Full) }
    }

    $iisReadable = $true
    if ($useStores) { $iisReadable = Read-IisBinding }

    #endregion

    #region 状態を決める

    $now = Get-Date
    foreach ($cert in $certs) {
        $cert.DaysLeft = [int][math]::Floor(($cert.NotAfter - $now).TotalDays)
        if ($cert.Kind -eq 'ストア' -and $iisSites.ContainsKey($cert.Thumbprint)) {
            $cert.Iis = (@($iisSites[$cert.Thumbprint] | ForEach-Object { $_.Text }) -join ', ')
            $cert.IisNames = (@($iisSites[$cert.Thumbprint] | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', ')
        }
        if ($cert.NotAfter -lt $now) { $cert.Status = '期限切れ' }
        elseif ($cert.NotBefore -gt $now) { $cert.Status = '有効期間前' }
        elseif ($cert.DaysLeft -le $Days) { $cert.Status = '期限間近' }
        else { $cert.Status = '有効' }
    }

    # 同じ場所に置き換わる証明書があれば「更新済み」とする。接続先は実際に使われている証明書なので、対象にしない。
    # IIS のサイトで使っている古い証明書は、新しい証明書があってもまだ止まるおそれがあるので、対象にしない。
    # IIS の設定を読めなかったときは、使っているかどうかが分からないので、ストアの証明書はすべて対象にしない
    $byIdentity = @{}
    foreach ($cert in $certs) {
        if (-not $cert.Scope) { continue }
        $key = '{0}|{1}|{2}' -f $cert.Scope, $cert.Subject, $cert.SanKey
        if (-not $byIdentity.ContainsKey($key)) { $byIdentity[$key] = New-Object 'System.Collections.Generic.List[object]' }
        $byIdentity[$key].Add($cert)
    }
    foreach ($cert in $certs) {
        if ($cert.Status -ne '期限切れ' -and $cert.Status -ne '期限間近') { continue }
        if (-not $cert.Scope -or $cert.Iis) { continue }
        if ($cert.Kind -eq 'ストア' -and -not $iisReadable) { continue }
        $key = '{0}|{1}|{2}' -f $cert.Scope, $cert.Subject, $cert.SanKey
        foreach ($newer in $byIdentity[$key]) {
            if ($newer.Status -ne '有効' -or $newer.NotAfter -le $cert.NotAfter) { continue }
            if ($cert.KeyText -eq 'あり' -and $newer.KeyText -ne 'あり') { continue }
            $cert.Status = '更新済み'
            $cert.Note = '新しい証明書: {0}' -f $newer.Thumbprint
            break
        }
    }

    # CurrentUser の Root や CA には、Windows が LocalMachine の証明書も見せる。同じ証明書が 2 行並ばないよう、
    # 同じ名前の LocalMachine のストアにも読んだ証明書は、CurrentUser の行を省く。
    # 更新済みの判定は Windows が見せるとおりの中身で行いたいので、判定の後で省く
    $machineCerts = @{}
    foreach ($cert in $certs) {
        if ($cert.Kind -eq 'ストア' -and $cert.Place.StartsWith('LocalMachine\', [StringComparison]::Ordinal)) {
            $machineCerts[$cert.Place.Substring(13) + '|' + $cert.Thumbprint] = $true
        }
    }
    $overlapCount = 0
    $keptCerts = New-Object 'System.Collections.Generic.List[object]'
    foreach ($cert in $certs) {
        if ($cert.Kind -eq 'ストア' -and $cert.Place.StartsWith('CurrentUser\', [StringComparison]::Ordinal) -and $machineCerts.ContainsKey($cert.Place.Substring(12) + '|' + $cert.Thumbprint)) {
            $overlapCount++
            Add-Skip '除外指定' ('{0}: {1}' -f $cert.Place, $cert.Subject) ('同じ証明書が LocalMachine\{0} にあります' -f $cert.Place.Substring(12))
            continue
        }
        $keptCerts.Add($cert)
    }
    $certs = $keptCerts

    $sortedCerts = @($certs | Sort-Object -Property NotAfter, Place, Thumbprint)
    $counts = @{ '期限切れ' = 0; '期限間近' = 0; '更新済み' = 0; '有効' = 0; '有効期間前' = 0 }
    foreach ($cert in $sortedCerts) { $counts[$cert.Status]++ }
    $actionCerts = @($sortedCerts | Where-Object { $_.Status -eq '期限切れ' -or $_.Status -eq '期限間近' })

    #endregion

    #region 結果を組み立てる

    $unreadCount = 0
    foreach ($value in $skipKinds.Values) { $unreadCount += $value }
    if ($actionCerts.Count -gt 0) { $exitCode = 2 } elseif ($unreadCount -gt 0) { $exitCode = 1 }

    $targetParts = @()
    if ($useStores) {
        if ($AllStores) { $targetParts += 'すべてのストア' } else { $targetParts += @($storeItems | ForEach-Object { $_.Text }) }
    }
    $targetParts += @($targetItems | ForEach-Object { $_.Text })
    $targetParts += @($pathTargets | ForEach-Object { $_.Full })

    $lines.Add(('証明書期限一覧 (Get-CertificateExpiry {0})' -f $ScriptVersion))
    $lines.Add((Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（管理者: {2}）' $startTime $machineName $(if ($isAdmin) { 'はい' } else { 'いいえ' })))
    $lines.Add(('対象: {0}' -f ($targetParts -join ', ')))
    if ($configPath) { $lines.Add(('設定ファイル: {0}' -f $configPath)) }
    $lines.Add((Format-Invariant '期限間近の基準: 残り {0} 日以内' $Days))
    if ($excludeRules.Count -gt 0) { $lines.Add(('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', '))) }
    $lines.Add('')
    $countParts = @(Format-Invariant '証明書 {0:N0} 件' $sortedCerts.Count)
    foreach ($name in '期限切れ', '期限間近', '更新済み', '有効') { $countParts += Format-Invariant '{0} {1:N0} 件' $name $counts[$name] }
    if ($counts['有効期間前'] -gt 0) { $countParts += Format-Invariant '有効期間前 {0:N0} 件' $counts['有効期間前'] }
    if ($excludedCount -gt 0) { $countParts += Format-Invariant '除外 {0:N0} 件' $excludedCount }
    $countParts += '所要 {0}' -f (Format-Duration ((Get-Date) - $startTime))
    $lines.Add(($countParts -join '  '))

    $notes = @()
    $kinds = @(foreach ($name in $skipKinds.Keys) { if ($skipKinds[$name] -gt 0) { , @($name, $skipKinds[$name]) } })
    if ($kinds.Count -eq 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $kinds[0][1] $kinds[0][0]
    } elseif ($kinds.Count -gt 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $unreadCount (($kinds | ForEach-Object { Format-Invariant '{0} {1:N0}' $_[0] $_[1] }) -join '・')
    }
    if ($secretCount -gt 0) { $notes += Format-Invariant '読まない形式 {0:N0} 件' $secretCount }
    if ($linkCount -gt 0) { $notes += Format-Invariant 'たどらなかったリンク {0:N0} 件' $linkCount }
    if ($notes.Count -gt 0) { $lines.Add(($notes -join '  ')) }
    if ($overlapCount -gt 0) { $lines.Add((Format-Invariant 'LocalMachine と重なる証明書 {0:N0} 件は省きました' $overlapCount)) }
    if (-not $iisReadable) {
        $iisNote = '※ IIS の設定を読めなかった部分があるため、サイトとの紐付けが足りないことがあります。ストアの証明書は更新済みにしていません'
        if (-not $isAdmin) { $iisNote += '（管理者として実行すると読めることがあります）' }
        $lines.Add($iisNote)
    }
    if ($skipKinds['権限なし'] -gt 0 -and -not $isAdmin) { $lines.Add('※ 管理者として実行すると、読める範囲が広がります') }
    $lines.Add('')
    $headCount = $lines.Count

    # 表の列の幅は、表に出す行に合わせる
    $placeWidth = 4
    foreach ($cert in $actionCerts) { $placeWidth = [math]::Max($placeWidth, (Get-DisplayWidth $cert.Place)) }
    $tableHeader = ' ' + (Format-Column '状態' 8 $false) + '  ' + (Format-Column '残り日数' 8 $true) + '  ' + (Format-Column '有効期限' 16 $false) + '  ' + (Format-Column '場所' $placeWidth $false) + '  サブジェクト'
    $tableLines = New-Object 'System.Collections.Generic.List[object]'
    foreach ($cert in $actionCerts) {
        # サブジェクトが空の証明書は、別名で見分ける
        $subjectText = if ($cert.Subject) { $cert.Subject } else { '（サブジェクトなし）' + $cert.San }
        if ($cert.IisNames) { $subjectText += '（IIS: {0}）' -f $cert.IisNames }
        if ($cert.Note) { $subjectText += '（{0}）' -f $cert.Note }
        $text = ' ' + (Format-Column $cert.Status 8 $false) + '  ' + (Format-Column ($cert.DaysLeft.ToString('N0', $Invariant)) 8 $true) + '  ' +
            $cert.NotAfter.ToString('yyyy-MM-dd HH:mm', $Invariant) + '  ' + (Format-Column $cert.Place $placeWidth $false) + '  ' + $subjectText
        $color = if ($cert.Status -eq '期限切れ') { 'Red' } else { 'Yellow' }
        $tableLines.Add(@($text, $color))
    }

    # summary.txt には、すべて書く
    $lines.Add('■ 対応が必要な証明書（期限の近い順）')
    if ($tableLines.Count -eq 0) {
        $lines.Add('    （対応が必要な証明書はありません）')
    } else {
        $lines.Add($tableHeader)
        foreach ($entry in $tableLines) { $lines.Add($entry[0]) }
    }
    $lines.Add('')
    if ($runDir) { $lines.Add(('保存先: {0}' -f $runDir)) }

    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($cert in $sortedCerts) {
        $rows.Add([object[]]@(
            $cert.Status, $cert.DaysLeft.ToString('N0', $Invariant), (Format-CertTime $cert.NotAfter), (Format-CertTime $cert.NotBefore), $cert.Kind, $cert.Place,
            (Protect-ExcelText $cert.Subject), (Protect-ExcelText $cert.San), (Protect-ExcelText $cert.Issuer), $cert.Usage, $cert.KeyText, (Protect-ExcelText $cert.Iis), (Protect-ExcelText $cert.Friendly),
            (Protect-ExcelText $cert.Thumbprint), $cert.Note))
    }

    #endregion

    #region 保存

    if ($runDir) {
        Write-Progress -Activity '証明書期限一覧' -Status '結果を保存しています'
        try {
            [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'certificates.csv')) $CertColumns $rows
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $skipped
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
    Write-Progress -Activity '証明書期限一覧' -Completed
    if (-not $finished -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

#region 画面表示

for ($i = 0; $i -lt $headCount; $i++) { Write-Host $lines[$i] }
Write-Host '■ 対応が必要な証明書（期限の近い順）' -ForegroundColor Cyan
if ($tableLines.Count -eq 0) {
    Write-Host '    （対応が必要な証明書はありません）'
} else {
    Write-Host $tableHeader
    # 保存できたときは、画面には先頭の分だけを出す。保存しなかったときは、ほかに見る場所がないためすべて出す
    $limit = if ($runDir -and -not $saveError) { $ScreenRows } else { $tableLines.Count }
    $shown = 0
    foreach ($entry in $tableLines) {
        if ($shown -ge $limit) { break }
        Write-Host $entry[0] -ForegroundColor $entry[1]
        $shown++
    }
    if ($tableLines.Count -gt $shown) {
        Write-Host (Format-Invariant '（ほか {0:N0} 件は certificates.csv を参照してください）' ($tableLines.Count - $shown))
    }
}
Write-Host ''

if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
} elseif ($saveWarning) {
    Write-Warning $saveWarning
}

#endregion

exit $exitCode
