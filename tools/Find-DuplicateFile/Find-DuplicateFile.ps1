<#
.SYNOPSIS
    内容が同じファイルと、コピーらしい名前のファイルを洗い出します。

.DESCRIPTION
    指定したフォルダの中のファイルを、サイズ、先頭部分のハッシュ、全体の SHA-256 の順に絞り込んで比べ、
    内容が同じファイルのグループを、減らせる容量の大きい順に画面に表示し、CSV（BOM 付き UTF-8）にも保存します。
    「報告書 - コピー.xlsx」のようにコピーらしい名前のファイルも、元と思われるファイルと比べて一覧にします。

    調べる場所のファイルを変更・削除することはありません。
    権限がなくて読めないフォルダや、開けないファイルはスキップし、件数を表示します。
    ジャンクションやシンボリックリンクはたどりません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Find-DuplicateFile.ps1 -Path D:\share

.PARAMETER Path
    調べるフォルダ。ドライブ（D:\）や共有フォルダ（\\server\share）も指定できます。省略すると現在のフォルダです。
    複数指定するときは、"D:\share,E:\archive" のようにカンマで区切って 1 つの文字列にします。

.PARAMETER MinSize
    内容を比べるファイルの最小サイズ。500KB、10MB、1.5GB のように書きます。単位を書かないとバイトです。既定は 1MB です。
    コピーらしい名前の判定には使いません。

.PARAMETER Include
    調べるファイルの名前。ワイルドカード（*）で、"*.xlsx,*.pdf" のようにカンマで区切って書きます。
    * 以外の文字（[ ] ? など）は、そのままの文字として扱います。大文字と小文字は区別しません。

.PARAMETER Exclude
    調べないフォルダやファイル。ワイルドカード（*）で、カンマで区切って書きます。
    \ を含まない書き方は名前に対して（どの階層でも）、D:\ や \\ で始まる書き方はフルパスに対して、
    それ以外の \ を含む書き方は -Path の各フォルダから見た相対パスに対して当てはめます。

.PARAMETER OlderThan
    最終更新から指定した日数以上たったファイルだけを調べます（1～36500）。

.PARAMETER Top
    画面に表示する件数。重複のグループと、コピーらしい名前のファイルのそれぞれに適用します（1～1000）。既定は 20 です。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.PARAMETER ConfigFile
    設定ファイルのパス。Path、MinSize、Include、Exclude、OlderThan、Top、OutputFolder、NoSave の値を
    「項目 = 値」の形で書いておけます。コマンドラインで指定した値は、設定ファイルより優先します。
    書き方は README を参照してください。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Find-DuplicateFile.ps1 -Path "D:\share,E:\archive"

    2 つのフォルダをまたいで、内容が同じファイルを探します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Find-DuplicateFile.ps1 -Path D:\share -MinSize 1 -Include "*.xlsx,*.pdf"

    Excel と PDF について、小さなファイルも含めて比べます。

.NOTES
    Find-DuplicateFile 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Position = 0)]
    [string[]]$Path = @('.'),

    [string]$MinSize = '1MB',

    [string[]]$Include = @(),

    [string[]]$Exclude = @(),

    [ValidateRange(1, 36500)]
    [int]$OlderThan,

    [ValidateRange(1, 1000)]
    [int]$Top = 20,

    [string]$OutputFolder,

    [switch]$NoSave,

    [string]$ConfigFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion  = '1.0.0'
$ProgressMs     = 250       # 進捗表示の間隔（ミリ秒）
$HashBufferSize = 1MB       # ハッシュを計算するときに 1 回で読む長さ
$HeadSize       = 64KB      # 全体を読む前に比べる、先頭部分の長さ
$OutputPrefix   = 'DuplicateFile_'
# このツールの結果のフォルダの名前（DuplicateFile_<ホスト名>_<日時>）。同じ文字で始まるだけの利用者のフォルダは調べる
$OutputNameRule = New-Object Text.RegularExpressions.Regex ('^DuplicateFile_.+_[0-9]{8}_[0-9]{6}(?:_[0-9]+)?$'), ([Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant')
$SystemFolders  = '$RECYCLE.BIN', 'System Volume Information'
$GroupColumns   = 'グループ', '件数', '実体の数', '1 件のサイズ(MB)', '1 件のサイズ(バイト)', '減らせる容量(MB)', '減らせる容量(バイト)', 'ハッシュ'
$FileColumns    = 'グループ', '実体', 'ファイル', 'ファイル名', 'サイズ(MB)', 'サイズ(バイト)', '最終更新日時', 'ハッシュ'
$CopyColumns    = '名前の形', '元との比較', 'ファイル', '元と思われるファイル', 'サイズ(MB)', 'サイズ(バイト)', '元のサイズ(バイト)', '最終更新日時', '元の最終更新日時', '備考'
$SkipColumns    = '理由', 'パス', '詳細'
$TimeFormat     = 'yyyy/MM/dd HH:mm:ss'
# オンライン専用ファイルの属性（OFFLINE・RECALL_ON_OPEN・RECALL_ON_DATA_ACCESS）。読むとダウンロードが始まる
$OnlineMask     = 0x00441000
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

function Format-Size([long]$Bytes) {
    $units = 'B', 'KB', 'MB', 'GB', 'TB', 'PB'
    $value = [double]$Bytes
    $i = 0
    while ($value -ge 1024 -and $i -lt $units.Count - 1) { $value /= 1024; $i++ }
    if ($i -eq 0) { return '{0} B' -f $Bytes }
    return [string]::Format($Invariant, '{0:N1} {1}', $value, $units[$i])
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
    if ($Text.Length -le 255) { return '="' + $Text + '"' }
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

# 共有フォルダやネットワークドライブを開けないときの案内。管理者として実行しているときは、そのことも添える
function Get-NetworkHint([string]$FullPath) {
    $isNetwork = $FullPath.StartsWith('\\')
    if (-not $isNetwork -and $FullPath -match '^[A-Za-z]:') {
        try { $isNetwork = (New-Object IO.DriveInfo ($FullPath.Substring(0, 2))).DriveType -eq 'Network' } catch { }
        # 管理者として実行していると、通常のユーザーで割り当てたネットワークドライブはドライブごと見えない
        if (-not [IO.Directory]::Exists($FullPath.Substring(0, 2) + '\')) { $isNetwork = $isNetwork -or $isAdmin }
    }
    $hint = @()
    if ($isNetwork) {
        $hint += '共有フォルダのときは、次の順に確かめてください。'
        $hint += '  1. エクスプローラーのアドレス欄にパスを貼り付けて、開けるか試す'
        $hint += '  2. ユーザー名とパスワードを聞かれたら入力する（「資格情報を記憶する」にチェックを入れると、次からは聞かれません）'
        $hint += '  3. 開けたら、もう一度実行する'
        if ($isAdmin) {
            $hint += '※ 管理者として実行しています。通常のユーザーで接続した共有フォルダやネットワークドライブは、管理者として開いた PowerShell からは見えないことがあります。管理者としてではなく実行してください。'
        }
    }
    return $hint
}

# フォルダを中まで読めるかを確かめる（共有フォルダは、接続していないと読めない）。読めなければ止める
function Assert-Folder([string]$FullPath, [string]$ConfigNote) {
    $reason = $null
    $ok     = $false
    try { $ok = [IO.Directory]::Exists((ConvertTo-LongPath $FullPath)) } catch { }
    if ($ok) {
        try {
            [void][IO.Directory]::EnumerateFileSystemEntries((ConvertTo-LongPath $FullPath)).GetEnumerator().MoveNext()
        } catch {
            $ok     = $false
            $reason = Get-ErrorText (Get-InnerException $_.Exception)
        }
    }
    if (-not $ok) {
        $message = 'フォルダを開けません: {0}' -f $FullPath
        if ($reason) { $message += '（{0}）' -f $reason }
        $message += $ConfigNote
        Stop-WithError ((@($message) + (Get-NetworkHint $FullPath)) -join "`r`n")
    }
}

# 500KB、10MB、1.5GB の形の文字列をバイト数にする。単位がなければバイト。正しくなければ -1 を返す
function ConvertFrom-SizeText([string]$Text) {
    $found = [regex]::Match($Text.Trim(), '^([0-9]+(?:\.[0-9]+)?)(KB|MB|GB|TB)?$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $found.Success) { return -1L }
    $scale = switch ($found.Groups[2].Value.ToUpperInvariant()) { 'KB' { 1KB } 'MB' { 1MB } 'GB' { 1GB } 'TB' { 1TB } default { 1 } }
    $value = [double]::Parse($found.Groups[1].Value, $Invariant) * $scale
    if ($value -lt 1 -or $value -ge [long]::MaxValue) { return -1L }
    return [long][math]::Floor($value)
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの。数値の範囲は param の ValidateRange と同じにする
$configItems   = 'Path', 'MinSize', 'Include', 'Exclude', 'OlderThan', 'Top', 'OutputFolder', 'NoSave'
$configRanges  = @{ OlderThan = 1, 36500; Top = 1, 1000 }
$configLists   = 'Include', 'Exclude'
$configPath    = $null
$configApplied = @{}   # 設定ファイルの値を使った項目 → 行番号
$configPathLines = @{} # 設定ファイルの Path の値 → 行番号（Path は複数の行に書けるため）

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
    $listItems    = @{ Include = (New-Object 'System.Collections.Generic.List[string]'); Exclude = (New-Object 'System.Collections.Generic.List[string]') }
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

        # Path・Include・Exclude は複数の行に分けて書ける。ほかの項目は、どちらが効いているか分からなくなるため 1 回だけ
        if ($itemName -ne 'Path' -and $configLists -notcontains $itemName -and $configLineOf.ContainsKey($itemName)) {
            Stop-WithConfigError $lineNumber ('{0} が 2 回書かれています（{1} 行目と {2} 行目）' -f $itemName, $configLineOf[$itemName], $lineNumber)
        }
        $configLineOf[$itemName] = $lineNumber

        if ($configLists -contains $itemName) {
            # "a,b" のように全体を囲む書き方と、"a","b" のように 1 つずつ囲む書き方のどちらも受け付ける
            $source = if ($itemValue.Contains('"')) { $rawValue } else { $itemValue }
            foreach ($part in ($source -split ',')) {
                $pattern = Remove-ConfigQuote $part
                if ($pattern.Contains('"')) { Stop-WithConfigError $lineNumber ('{0} の値に引用符（"）が残っています: {1}' -f $itemName, $pattern) }
                if ($pattern) { $listItems[$itemName].Add((Expand-ConfigVariable $lineNumber $pattern)) }
            }
        } elseif ($itemName -eq 'NoSave') {
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
        } elseif ($itemName -eq 'MinSize') {
            if ((ConvertFrom-SizeText $itemValue) -lt 1) { Stop-WithConfigError $lineNumber ('MinSize は 500KB、10MB のように指定してください: {0}' -f $itemValue) }
            $configValues[$itemName] = $itemValue
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
                # カンマを含む名前のフォルダも書けるよう、1 行を 1 つのフォルダとして扱う
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

$minBytes = ConvertFrom-SizeText $MinSize
if ($minBytes -lt 1) { Stop-WithError ('-MinSize は 500KB、10MB のように指定してください: {0}' -f $MinSize) }

# -Path は、-File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# カンマを含む名前のフォルダのため、全体が 1 つのフォルダとして存在するときは区切らない。設定ファイルの値は 1 行が 1 つのフォルダ
$rootTexts = New-Object 'System.Collections.Generic.List[string]'
foreach ($item in $Path) {
    if ($configApplied.ContainsKey('Path')) { $rootTexts.Add($item); continue }
    $whole = Remove-Quote $item
    if ($whole.Contains('"')) { Stop-WithQuoteError '-Path' $whole }
    $isWhole = $false
    try { $isWhole = $whole.Contains(',') -and [IO.Directory]::Exists((ConvertTo-LongPath (Resolve-FullPath (Resolve-StartPath $whole)))) } catch { }
    if ($isWhole) { $rootTexts.Add($whole); continue }
    foreach ($part in ($whole -split ',')) {
        $text = (Remove-Quote $part).Trim("'")
        if ($text) { $rootTexts.Add($text) }
    }
}
if ($rootTexts.Count -eq 0) { Stop-WithError '調べるフォルダを指定してください（-Path で指定するか、設定ファイルの Path に書きます）' }

$rootPaths = New-Object 'System.Collections.Generic.List[string]'
foreach ($text in $rootTexts) {
    $note = ''
    if ($configPathLines.ContainsKey($text) -and $configApplied.ContainsKey('Path')) { $note = '（設定ファイルの {0} 行目）' -f $configPathLines[$text] }
    $rootText = $text
    if ($rootText -match '^(\\\\\?\\)?[A-Za-z]:$') { $rootText += '\' }
    $rootFull = Resolve-FullPath (Resolve-StartPath $rootText)
    Assert-Folder $rootFull $note
    $rootPaths.Add($rootFull)
}

# 同じフォルダや、片方の中にあるフォルダを指定すると、同じファイルを重複として数えてしまう
# （ネットワークドライブと UNC パスのように、同じ場所を別の名前で指定した場合は判定できない）
$sepText = [string][IO.Path]::DirectorySeparatorChar
for ($i = 0; $i -lt $rootPaths.Count; $i++) {
    for ($j = $i + 1; $j -lt $rootPaths.Count; $j++) {
        $firstPrefix  = $rootPaths[$i].TrimEnd([char]$sepText) + $sepText
        $secondPrefix = $rootPaths[$j].TrimEnd([char]$sepText) + $sepText
        if ([string]::Equals($firstPrefix, $secondPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            Stop-WithError ('同じフォルダが 2 回指定されています: {0}' -f $rootPaths[$i])
        }
        if ($firstPrefix.StartsWith($secondPrefix, [StringComparison]::OrdinalIgnoreCase) -or $secondPrefix.StartsWith($firstPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            Stop-WithError ('指定したフォルダの片方が、もう片方の中にあります。同じファイルを二重に数えてしまいます。' + "`r`n" + $rootPaths[$i] + "`r`n" + $rootPaths[$j])
        }
    }
}

# -Include と -Exclude は、-File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# コマンドプロンプトでは ' がそのまま届くため、'*.log,logs' のように全体を囲んだ場合も、各要素の前後の ' を外す
function Get-PatternText([string[]]$Values, [string]$Name) {
    $texts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($value in $Values) {
        foreach ($part in ($value -split ',')) {
            $text = (Remove-Quote $part).Trim("'")
            if ($text.Contains('"')) { Stop-WithQuoteError $Name $text }
            $text = $text.TrimEnd('\')
            if ($text -and -not ($texts -contains $text)) { $texts.Add($text) }
        }
    }
    return , $texts
}

# ワイルドカードは * だけにして、それ以外の文字は文字どおりに扱う。5.1 の WildcardPattern.Escape は
# バッククォートをエスケープしないため、自分でエスケープしてから照合用のパターンを作る
$wildcardOptions = [Management.Automation.WildcardOptions]'IgnoreCase, CultureInvariant'
function New-NameMatcher([string]$Text) {
    $escaped = $Text.Replace('`', '``').Replace('[', '`[').Replace(']', '`]').Replace('?', '`?')
    return [Management.Automation.WildcardPattern]::new($escaped, $wildcardOptions)
}

# \ を含まない書き方は名前に、ドライブや \\ で始まる書き方はフルパスに、それ以外の \ を含む書き方は相対パスに当てはめる。
# 相対パスの先頭の \ と .\ は「直下から」の意味なので外す
$excludeTexts = Get-PatternText $Exclude '-Exclude'
$excludeRules = @(
    foreach ($text in $excludeTexts) {
        if ($text -match '^([A-Za-z]:\\|\\\\)') {
            [pscustomobject]@{ Text = $text; Scope = 'Full'; Matcher = (New-NameMatcher $text) }
        } elseif ($text.Contains('\')) {
            [pscustomobject]@{ Text = $text; Scope = 'Rel'; Matcher = (New-NameMatcher ($text -replace '^\.\\', '').TrimStart('\')) }
        } else {
            [pscustomobject]@{ Text = $text; Scope = 'Name'; Matcher = (New-NameMatcher $text) }
        }
    }
)
$hasExclude  = $excludeRules.Count -gt 0
$hasFullRule = @($excludeRules | Where-Object { $_.Scope -eq 'Full' }).Count -gt 0

$includeTexts    = Get-PatternText $Include '-Include'
$includeMatchers = @(foreach ($text in $includeTexts) { New-NameMatcher $text })
$hasInclude      = $includeMatchers.Count -gt 0

$startTime = Get-Date
$useCutoff = $OlderThan -gt 0   # 指定しなかったときは 0（設定ファイルで指定した場合も含めて判定する）
$cutoff    = $startTime.AddDays(-$OlderThan)
$cutoffUtc = $cutoff.ToUniversalTime()

#endregion

#region 判定の部品

$dirType  = [IO.DirectoryInfo]
$rpFlag   = [IO.FileAttributes]::ReparsePoint
$hsFlag   = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System

$skipped   = New-Object 'System.Collections.Generic.List[object]'
$skipKinds = [ordered]@{ '権限なし' = 0; '読めない' = 0; 'オンライン専用' = 0; '調査中に変更' = 0; 'その他のエラー' = 0 }
$excludedCount = 0
$linkCount     = 0
$progress  = @{ Next = 0L; Done = 0L; Total = 0L; Files = 0; FileTotal = 0 }
$sw        = [Diagnostics.Stopwatch]::StartNew()

function Add-Skip([string]$Reason, [string]$ItemPath, [string]$Detail) {
    $skipped.Add([object[]]@($Reason, $ItemPath, $Detail))
    if ($skipKinds.Contains($Reason)) { $skipKinds[$Reason]++ }
}

# 除外に当てはまれば、当てはまった書き方を返す
function Get-ExcludeMatch([string]$Rel, [string]$Name, [string]$Full) {
    foreach ($rule in $excludeRules) {
        $target = switch ($rule.Scope) { 'Name' { $Name } 'Rel' { $Rel } default { $Full } }
        if ($rule.Matcher.IsMatch($target)) { return $rule.Text }
    }
    return $null
}

function Test-IncludeMatch([string]$Name) {
    foreach ($matcher in $includeMatchers) { if ($matcher.IsMatch($Name)) { return $true } }
    return $false
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

# 表示の間隔が来たかどうかは、呼ぶ側で確かめる（毎回、表示の文字列を組み立てないため）
function Show-Progress([string]$Status, [string]$Operation) {
    $percent = -1
    if ($progress.Total -gt 0) { $percent = [int][math]::Min(100, 100 * $progress.Done / $progress.Total) }
    Write-Progress -Activity '重複ファイル検出' -Status $Status -CurrentOperation $Operation -PercentComplete $percent
    $progress.Next = $sw.ElapsedMilliseconds + $ProgressMs
}

# ファイルの情報は、同じファイルを内容の比較とコピーらしい名前の比較で読み直さないよう、1 つの記録にまとめる
$itemByPath = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
function Get-FileItem($Entry) {
    $full = $Entry.FullName
    $item = $null
    if ($itemByPath.TryGetValue($full, [ref]$item)) { return $item }
    # 更新日時が範囲外のファイル（日時が壊れたものなど）は、日時を読むと例外になる。比べずに記録だけ残す
    $time = $null
    $timeError = $null
    try { $time = $Entry.LastWriteTimeUtc } catch { $timeError = Get-ErrorText (Get-InnerException $_.Exception) }
    $item = [pscustomobject]@{
        Full = $full; Name = $Entry.Name; Size = $Entry.Length; Time = $time
        Online = (([int]$Entry.Attributes -band $OnlineMask) -ne 0); Hash = $null; Fail = $null; Entity = 0
    }
    if ($timeError) { Add-ItemFailure $item 'その他のエラー' $timeError }
    $itemByPath[$full] = $item
    return $item
}

# 読めなかったファイルは、何度比べようとしても skipped.csv に 1 回だけ残す
function Add-ItemFailure($Item, [string]$Reason, [string]$Detail) {
    if ($Item.Fail) { return }
    $Item.Fail = $Reason
    Add-Skip $Reason (ConvertTo-DisplayPath $Item.Full) $Detail
}

# ファイルを読めなかった例外を、skipped.csv の理由にする
function Add-ReadFailure($Item, [Exception]$Exception) {
    $ex = Get-InnerException $Exception
    $reason = if ($ex -is [UnauthorizedAccessException]) { '権限なし' }
              elseif ($ex -is [IO.FileNotFoundException] -or $ex -is [IO.DirectoryNotFoundException]) { '調査中に変更' }
              else { '読めない' }
    Add-ItemFailure $Item $reason (Get-ErrorText $ex)
}

# NTFS のフォルダの一覧のサイズと更新日時は、ハードリンクの別名や書き込み中のファイルで古いことがあるため、ファイルそのものの値に更新する。
# 読めなければ $false を返し、skipped に残す
function Update-ItemInfo($Item) {
    if ($Item.Fail) { return $false }
    if ($Item.Hash) { return $true }
    try {
        $info = New-Object IO.FileInfo $Item.Full
        if (-not $info.Exists) {
            Add-ItemFailure $Item '調査中に変更' '調べている間にファイルがなくなりました'
            return $false
        }
        $Item.Size = $info.Length
        $Item.Time = $info.LastWriteTimeUtc
        return $true
    } catch {
        Add-ReadFailure $Item $_.Exception
        return $false
    }
}

# ほかのプログラムが書き込み中のファイルも読めるように開く。FIPS モードの PC でも使えるよう、CryptoServiceProvider を使う
$hashBuffer = New-Object 'byte[]' $HashBufferSize
$sha256     = New-Object Security.Cryptography.SHA256CryptoServiceProvider   # ファイルごとに作ると遅いため、Initialize で使い回す

# 先頭部分のハッシュを返す。読めないときは $null を返し、skipped に残す
function Get-HeadHash($Item) {
    if ($Item.Fail) { return $null }
    $stream = $null
    try {
        $stream = New-Object IO.FileStream($Item.Full, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]'ReadWrite, Delete'), 65536)
        $total = 0
        while ($total -lt $HeadSize) {
            $read = $stream.Read($hashBuffer, $total, $HeadSize - $total)
            if ($read -le 0) { break }
            $total += $read
        }
        $sha256.Initialize()
        return [BitConverter]::ToString($sha256.ComputeHash($hashBuffer, 0, $total)).Replace('-', '')
    } catch {
        Add-ReadFailure $Item $_.Exception
        return $null
    } finally {
        if ($stream) { $stream.Dispose() }
    }
}

# 全体の SHA-256 を、Get-FileHash と同じ大文字の 16 進数で返す。読めないときは $null を返し、skipped に残す
function Get-FullHash($Item) {
    if ($Item.Hash) { return $Item.Hash }
    if ($Item.Fail) { return $null }
    $display = ConvertTo-DisplayPath $Item.Full
    if ($Item.Online) {
        Add-ItemFailure $Item 'オンライン専用' '読むとダウンロードが始まるため、内容を比べていません'
        return $null
    }
    $stream = $null
    try {
        if (-not (Update-ItemInfo $Item)) { return $null }
        $sizeBefore = $Item.Size
        $timeBefore = $Item.Time
        $sha256.Initialize()
        $stream = New-Object IO.FileStream($Item.Full, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]'ReadWrite, Delete'), 65536)
        $total = 0L
        while ($true) {
            $read = $stream.Read($hashBuffer, 0, $hashBuffer.Length)
            if ($read -le 0) { break }
            [void]$sha256.TransformBlock($hashBuffer, 0, $read, $null, 0)
            $total += $read
            $progress.Done += $read
            if ($sw.ElapsedMilliseconds -ge $progress.Next) {
                Show-Progress (Format-Invariant 'ハッシュを計算しています: {0:N0} / {1:N0} 件（{2} / {3}）' $progress.Files $progress.FileTotal (Format-Size $progress.Done) (Format-Size $progress.Total)) $display
            }
        }
        [void]$sha256.TransformFinalBlock($hashBuffer, 0, 0)
        $stream.Dispose()
        $stream = $null
        # 読んでいる間に書き換えられたファイルは、ハッシュが中途半端な内容のものになる
        $unchanged = $total -eq $sizeBefore -and (Update-ItemInfo $Item) -and $Item.Size -eq $sizeBefore -and $Item.Time -eq $timeBefore
        if (-not $unchanged) {
            Add-ItemFailure $Item '調査中に変更' '読んでいる間にサイズか更新日時が変わりました'
            return $null
        }
        $Item.Hash = [BitConverter]::ToString($sha256.Hash).Replace('-', '')
        return $Item.Hash
    } catch {
        Add-ReadFailure $Item $_.Exception
        return $null
    } finally {
        if ($stream) { $stream.Dispose() }
        $progress.Files++
    }
}

# コピーらしい名前の判定。\d は全角の数字にも当てはまるため、[0-9] で書く
$regexOptions = [Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant'
$copyHint     = New-Object Text.RegularExpressions.Regex ('コピー|copy|[(（][0-9]|old|bak|backup|旧|古|バックアップ|控え|[0-9]{4}'), ([Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant, Compiled')
$copySuffix   = New-Object Text.RegularExpressions.Regex ('^(.+?)(?: - (?:コピー|copy)|のコピー)(?: ?[(（][0-9]+[)）])?$'), $regexOptions
$copyPrefix   = New-Object Text.RegularExpressions.Regex ('^(?:コピー(?: ?[(（][0-9]+[)）])? [～〜] |copy(?: [(（][0-9]+[)）])? of )(.+)$'), $regexOptions
$numberSuffix = New-Object Text.RegularExpressions.Regex ('^(.+?) ?[(（][0-9]{1,3}[)）]$'), $regexOptions
$oldSuffix    = New-Object Text.RegularExpressions.Regex ('^(.+?)[_\- ](?:old|bak|backup|旧|古|バックアップ|控え)$'), $regexOptions
$dateSuffix   = New-Object Text.RegularExpressions.Regex ('^(.+?)[_\- ]([0-9]{8}|[0-9]{4}-[0-9]{2}-[0-9]{2}|[0-9]{6}|[0-9]{4})$'), $regexOptions
$FormOrder    = @{ 'コピー' = 1; '番号付き' = 2; '古い版' = 3; '日付付き' = 4 }

# 日付として正しい数字かどうか（20261102、2026-11-02、261102、1102）
function Test-DateText([string]$Text) {
    $digits = $Text.Replace('-', '')
    $format = 'yyyyMMdd'
    if ($digits.Length -eq 6) { $format = 'yyMMdd' }
    elseif ($digits.Length -eq 4) { $digits = '2000' + $digits }   # 2 月 29 日も通るよう、うるう年で確かめる
    $parsed = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($digits, $format, $Invariant, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $false }
    return $parsed.Year -ge 1900 -and $parsed.Year -le 2099
}

# 今調べているフォルダの中で、元と思われるファイルを探すための表。コピーらしい名前があったフォルダでだけ作る
$dirIndex = @{ Files = $null; Names = $null; Stems = $null }
function Get-DirFile([string]$Name) {
    if ($null -eq $dirIndex.Names) {
        $dirIndex.Names = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
        $dirIndex.Stems = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($file in $dirIndex.Files) {
            $dirIndex.Names[$file.Name] = $file
            $dot = $file.Name.LastIndexOf([char]'.')
            if ($dot -le 0) { continue }
            $fileExt = $file.Name.Substring($dot)
            if ($fileExt -ieq '.bak' -or $fileExt -ieq '.old') { continue }
            $fileStem = $file.Name.Substring(0, $dot)
            # 拡張子だけが違うファイルが 2 つ以上あると、どれが元か分からない
            if ($dirIndex.Stems.ContainsKey($fileStem)) { $dirIndex.Stems[$fileStem] = $false } else { $dirIndex.Stems[$fileStem] = $file }
        }
    }
    $file = $null
    if ($dirIndex.Names.TryGetValue($Name, [ref]$file)) { return $file }
    return $null
}

# 報告書.bak の元は、報告書.xlsx のように拡張子だけが違うファイル。1 つに決まるときだけ元とする
function Get-DirFileByStem([string]$Stem) {
    $found = Get-DirFile $Stem
    if ($found) { return $found }
    $file = $null
    if ($dirIndex.Stems.TryGetValue($Stem, [ref]$file) -and $file) { return $file }
    return $null
}

# コピーらしい名前なら、名前の形と元と思われるファイルを返す。番号付きと日付付きは、元があるものだけ
function Get-CopyMatch($Entry) {
    $name = $Entry.Name
    $dot  = $name.LastIndexOf([char]'.')
    $stem = if ($dot -gt 0) { $name.Substring(0, $dot) } else { $name }
    $ext  = if ($dot -gt 0) { $name.Substring($dot) } else { '' }

    # コピーの形は重なることがあるので、すべて取り除く
    $base = $stem
    while ($true) {
        $found = $copySuffix.Match($base)
        if (-not $found.Success) { $found = $copyPrefix.Match($base) }
        if (-not $found.Success) { break }
        $base = $found.Groups[1].Value
    }
    if ($base.Length -ne $stem.Length) { return @{ Form = 'コピー'; Original = (Get-DirFile ($base + $ext)) } }

    $found = $numberSuffix.Match($stem)
    if ($found.Success) {
        $original = Get-DirFile ($found.Groups[1].Value + $ext)
        if ($original) { return @{ Form = '番号付き'; Original = $original } }
    }

    $found = $oldSuffix.Match($stem)
    if ($found.Success) { return @{ Form = '古い版'; Original = (Get-DirFile ($found.Groups[1].Value + $ext)) } }
    if ($ext -ieq '.bak' -or $ext -ieq '.old') { return @{ Form = '古い版'; Original = (Get-DirFileByStem $stem) } }

    $found = $dateSuffix.Match($stem)
    if ($found.Success -and (Test-DateText $found.Groups[2].Value)) {
        $original = Get-DirFile ($found.Groups[1].Value + $ext)
        if ($original) { return @{ Form = '日付付き'; Original = $original } }
    }
    return $null
}

# ハードリンクの判定に使う Windows の API。.NET Framework には同じ機能がない。
# 実体を共有している（リンクが 2 つ以上ある）ときだけ実体の ID を返し、1 つなら空文字、判定できなければ $null を返す
$NativeSource = @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

public static class WorkbenchFileIdentity {
    [StructLayout(LayoutKind.Sequential)]
    private struct ByHandleFileInformation {
        public uint FileAttributes;
        public uint CreationTimeLow, CreationTimeHigh;
        public uint LastAccessTimeLow, LastAccessTimeHigh;
        public uint LastWriteTimeLow, LastWriteTimeHigh;
        public uint VolumeSerialNumber;
        public uint FileSizeHigh, FileSizeLow;
        public uint NumberOfLinks;
        public uint FileIndexHigh, FileIndexLow;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFileW(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetFileInformationByHandle(SafeFileHandle handle, out ByHandleFileInformation info);

    public static string GetId(string path) {
        // FILE_READ_ATTRIBUTES、すべての共有、OPEN_EXISTING
        using (SafeFileHandle handle = CreateFileW(path, 0x80, 7, IntPtr.Zero, 3, 0, IntPtr.Zero)) {
            if (handle.IsInvalid) { return null; }
            ByHandleFileInformation info;
            if (!GetFileInformationByHandle(handle, out info)) { return null; }
            if (info.NumberOfLinks < 2) { return ""; }
            if (info.FileIndexHigh == 0 && info.FileIndexLow == 0) { return null; }
            return info.VolumeSerialNumber + ":" + info.FileIndexHigh + ":" + info.FileIndexLow;
        }
    }
}
'@

function Format-Time($TimeUtc, [string]$Format = $TimeFormat) {
    if ($null -eq $TimeUtc) { return '' }
    return $TimeUtc.ToLocalTime().ToString($Format, $Invariant)
}

function Format-Megabytes([long]$Bytes) {
    return ($Bytes / 1MB).ToString('N2', $Invariant)
}

function Format-Bytes([long]$Bytes) {
    return $Bytes.ToString('N0', $Invariant)
}

# スキップ・リンク・除外の件数の行と、管理者として実行するとよいことの案内
function Get-NoteLine {
    $result = @()
    $notes  = @()
    $total  = 0
    foreach ($value in $skipKinds.Values) { $total += $value }
    $kinds = @(foreach ($name in $skipKinds.Keys) { if ($skipKinds[$name] -gt 0) { , @($name, $skipKinds[$name]) } })
    if ($kinds.Count -eq 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $kinds[0][1] $kinds[0][0]
    } elseif ($kinds.Count -gt 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $total (($kinds | ForEach-Object { Format-Invariant '{0} {1:N0}' $_[0] $_[1] }) -join '・')
    }
    if ($linkCount -gt 0) { $notes += Format-Invariant 'たどらなかったリンク {0:N0} 件' $linkCount }
    if ($excludedCount -gt 0) { $notes += Format-Invariant '除外 {0:N0} 件' $excludedCount }
    if ($notes.Count -gt 0) { $result += $notes -join '  ' }
    return $result
}

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
$fatal      = $null
$saveError  = $null
$exitCode   = 0
$lines      = New-Object 'System.Collections.Generic.List[string]'   # summary.txt に書く行
$totalFiles = 0L
$totalDirs  = 0L
$totalSize  = 0L
$bySize     = New-Object 'System.Collections.Generic.Dictionary[long,object]'
$copies     = New-Object 'System.Collections.Generic.List[object]'
try {
    #region フォルダをたどる

    # 再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる
    foreach ($rootPath in $rootPaths) {
        $rootLong  = ConvertTo-LongPath $rootPath
        $rootTrim  = $rootLong.TrimEnd([char[]]@('\', '/'))
        $prefixLen = $rootTrim.Length + 1
        $stack = New-Object 'System.Collections.Generic.Stack[object]'
        $stack.Push((New-Object IO.DirectoryInfo $rootLong))
        $atRoot = $true

        while ($stack.Count -gt 0) {
            $dir = $stack.Pop()
            if ($sw.ElapsedMilliseconds -ge $progress.Next) { Show-Progress (Format-Invariant '調べています: フォルダ {0:N0} 件 / ファイル {1:N0} 件' $totalDirs $totalFiles) (ConvertTo-DisplayPath $dir.FullName) }

            try {
                $entries = $dir.GetFileSystemInfos()
            } catch {
                $ex = Get-InnerException $_.Exception
                if ($atRoot) {
                    $fatal = 'フォルダを読めません: {0}（{1}）' -f $rootPath, (Get-ErrorText $ex)
                    break
                }
                # 種類を判定できないリンクは、読み取りが禁止されていると普通のフォルダに見える。権限なしには数えない
                if ($ex -is [UnauthorizedAccessException] -and ($dir.Attributes -band $rpFlag) -eq $rpFlag) {
                    $linkCount++
                    Add-Skip 'リンク' (ConvertTo-DisplayPath $dir.FullName) '読み取りが禁止されたリンク'
                    continue
                }
                $reason = if ($ex -is [UnauthorizedAccessException]) { '権限なし' } else { 'その他のエラー' }
                $detail = if ($ex -is [UnauthorizedAccessException]) { 'アクセスが拒否されました' } else { Get-ErrorText $ex }
                Add-Skip $reason (ConvertTo-DisplayPath $dir.FullName) $detail
                continue
            }
            $atRoot = $false

            $dirFiles = New-Object 'System.Collections.Generic.List[object]'
            foreach ($entry in $entries) {
                $isDir = $entry -is $dirType
                # 前回の結果を比べないよう、このツールの結果のフォルダは中を見ない
                if ($isDir -and $OutputNameRule.IsMatch($entry.Name)) {
                    $outputDisplay = ConvertTo-DisplayPath $entry.FullName
                    if (-not [string]::Equals($outputDisplay, $runDir, [StringComparison]::OrdinalIgnoreCase)) { Add-Skip '除外指定' $outputDisplay 'このツールの結果のフォルダ' }
                    continue
                }
                # ごみ箱の中身は重複として見ても意味がなく、System Volume Information はふつう読めないため、調べない
                if ($isDir -and $SystemFolders -contains $entry.Name) {
                    $excludedCount++
                    Add-Skip '除外指定' (ConvertTo-DisplayPath $entry.FullName) 'ごみ箱・システムの管理用のフォルダ'
                    continue
                }
                if ($hasExclude) {
                    $rel = $entry.FullName.Substring($prefixLen)
                    if (-not $OnWindows) { $rel = $rel.Replace('/', '\') }
                    $fullText = if ($hasFullRule) { ConvertTo-DisplayPath $entry.FullName } else { '' }
                    $matched = Get-ExcludeMatch $rel $entry.Name $fullText
                    if ($matched) {
                        $excludedCount++
                        Add-Skip '除外指定' (ConvertTo-DisplayPath $entry.FullName) (Protect-ExcelText $matched)
                        continue
                    }
                }
                # リンクの判定は、リパースポイントのときだけ行う（件数が多いと、関数を呼ぶだけで時間がかかるため）
                if (($entry.Attributes -band $rpFlag) -eq $rpFlag) {
                    $linkName = Get-LinkName $entry
                    if ($linkName) {
                        $linkCount++
                        Add-Skip 'リンク' (ConvertTo-DisplayPath $entry.FullName) $linkName
                        continue
                    }
                }
                if ($isDir) {
                    $totalDirs++
                    $stack.Push($entry)
                } else {
                    $dirFiles.Add($entry)
                }
            }

            # 元と思われるファイルは -Include と -OlderThan に当てはまらなくても探すため、フォルダごとにまとめて判定する
            $dirIndex.Files = $dirFiles
            $dirIndex.Names = $null
            $dirIndex.Stems = $null
            foreach ($entry in $dirFiles) {
                # -OlderThan で更新日時を読むときの例外（日時が壊れたファイル）は、そのファイルだけ飛ばして続ける
                try {
                    if ($hasInclude -and -not (Test-IncludeMatch $entry.Name)) { continue }
                    if ($useCutoff -and $entry.LastWriteTimeUtc -ge $cutoffUtc) { continue }
                    $length = $entry.Length
                    $entryItem = $null
                    if ($length -ge $minBytes) { $entryItem = Get-FileItem $entry }
                    $copyMatch = $null
                    $originalItem = $null
                    if ($copyHint.IsMatch($entry.Name)) {
                        $copyMatch = Get-CopyMatch $entry
                        if ($copyMatch) {
                            if (-not $entryItem) { $entryItem = Get-FileItem $entry }
                            if ($copyMatch.Original) { $originalItem = Get-FileItem $copyMatch.Original }
                        }
                    }
                    $totalFiles++
                    $totalSize += $length
                    if ($length -ge $minBytes) {
                        $sameSize = $null
                        if (-not $bySize.TryGetValue($length, [ref]$sameSize)) {
                            $sameSize = New-Object 'System.Collections.Generic.List[object]'
                            $bySize[$length] = $sameSize
                        }
                        $sameSize.Add($entryItem)
                    }
                    if ($copyMatch) {
                        $copies.Add([pscustomobject]@{ Form = $copyMatch.Form; Order = $FormOrder[$copyMatch.Form]; Item = $entryItem; Original = $originalItem; Result = ''; Note = ''; Size = $length; Full = $entry.FullName })
                    }
                } catch {
                    Add-Skip 'その他のエラー' (ConvertTo-DisplayPath $entry.FullName) (Get-ErrorText (Get-InnerException $_.Exception))
                }
            }
        }
        if ($fatal) { break }
    }

    #endregion

    if (-not $fatal) {
        #region 内容を比べる

        # サイズが同じものを、先頭部分のハッシュで絞り込む（全体を読むのは時間がかかるため）
        $fullGroups = New-Object 'System.Collections.Generic.List[object]'
        $headDone = 0
        foreach ($sameSize in $bySize.Values) {
            if ($sameSize.Count -lt 2) { continue }
            $readable = New-Object 'System.Collections.Generic.List[object]'
            foreach ($item in $sameSize) {
                if ($item.Online) { [void](Get-FullHash $item) } else { $readable.Add($item) }
            }
            if ($readable.Count -lt 2) { continue }
            if ($readable[0].Size -le $HeadSize) { $fullGroups.Add($readable); continue }

            $byHead = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
            foreach ($item in $readable) {
                $headDone++
                if ($sw.ElapsedMilliseconds -ge $progress.Next) { Show-Progress (Format-Invariant '先頭部分を比べています: {0:N0} 件' $headDone) (ConvertTo-DisplayPath $item.Full) }
                $headHash = Get-HeadHash $item
                if (-not $headHash) { continue }
                $sameHead = $null
                if (-not $byHead.TryGetValue($headHash, [ref]$sameHead)) {
                    $sameHead = New-Object 'System.Collections.Generic.List[object]'
                    $byHead[$headHash] = $sameHead
                }
                $sameHead.Add($item)
            }
            foreach ($sameHead in $byHead.Values) { if ($sameHead.Count -ge 2) { $fullGroups.Add($sameHead) } }
        }

        foreach ($group in $fullGroups) {
            $progress.FileTotal += $group.Count
            $progress.Total     += $group.Count * $group[0].Size
        }
        $comparedFiles = 0L
        $comparedSize  = 0L
        $hashGroups = New-Object 'System.Collections.Generic.List[object]'
        foreach ($group in $fullGroups) {
            $byHash = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
            foreach ($item in $group) {
                $fullHash = Get-FullHash $item
                if (-not $fullHash) { continue }
                $comparedFiles++
                $comparedSize += $item.Size
                $sameHash = $null
                if (-not $byHash.TryGetValue($fullHash, [ref]$sameHash)) {
                    $sameHash = New-Object 'System.Collections.Generic.List[object]'
                    $byHash[$fullHash] = $sameHash
                }
                $sameHash.Add($item)
            }
            foreach ($key in $byHash.Keys) { if ($byHash[$key].Count -ge 2) { $hashGroups.Add(@($key, $byHash[$key])) } }
        }

        # ハードリンクを判定する。消しても容量が空かないため、同じ実体は 1 つとして数える
        $identityReady = $false
        if ($hashGroups.Count -gt 0 -and $OnWindows) {
            try {
                if (-not ('WorkbenchFileIdentity' -as [type])) { Add-Type -TypeDefinition $NativeSource -ErrorAction Stop }
                $identityReady = $true
            } catch { }
        }
        $identityFailed = $hashGroups.Count -gt 0 -and -not $identityReady
        $hardLinkFiles  = 0L
        $groupsRaw = New-Object 'System.Collections.Generic.List[object]'
        foreach ($pair in $hashGroups) {
            $entityOf   = @{}
            $nextEntity = 0
            # 表示と同じ順に実体の番号を振る
            $sortedItems = @($pair[1] | Sort-Object -Property Time, Full)
            foreach ($item in $sortedItems) {
                $identity = ''
                if ($identityReady) {
                    $identity = [WorkbenchFileIdentity]::GetId($item.Full)
                    if ($null -eq $identity) { $identityFailed = $true; $identity = '' }
                }
                if ($identity -and $entityOf.ContainsKey($identity)) {
                    $item.Entity = $entityOf[$identity]
                } else {
                    $nextEntity++
                    $item.Entity = $nextEntity
                    if ($identity) { $entityOf[$identity] = $nextEntity }
                }
            }
            $hardLinkFiles += $sortedItems.Count - $nextEntity
            # すべて同じ実体なら、消しても容量は空かないので、重複には出さない
            if ($nextEntity -lt 2) { continue }
            $groupsRaw.Add([pscustomobject]@{
                Hash = $pair[0]; Size = $sortedItems[0].Size; Count = $sortedItems.Count; Entities = $nextEntity
                Waste = [long]($sortedItems[0].Size * ($nextEntity - 1))
                Items = $sortedItems
            })
        }
        $groups = @($groupsRaw | Sort-Object -Property @{ Expression = 'Waste'; Descending = $true }, @{ Expression = 'Size'; Descending = $true }, Hash)

        #endregion

        #region コピーらしい名前を元と比べる

        $progress.Total = 0L; $progress.Done = 0L; $progress.Files = 0; $progress.FileTotal = 0
        $counted = New-Object 'System.Collections.Generic.HashSet[string]'   # 同じ元を持つコピーが複数あっても、読むのは 1 回
        foreach ($copy in $copies) {
            if ($copy.Original -and $copy.Original.Size -eq $copy.Item.Size) {
                foreach ($target in $copy.Item, $copy.Original) {
                    if (-not $target.Hash -and $counted.Add($target.Full)) { $progress.FileTotal++; $progress.Total += $target.Size }
                }
            }
        }
        foreach ($copy in $copies) {
            if (-not $copy.Original) {
                [void](Update-ItemInfo $copy.Item)
                $copy.Size = $copy.Item.Size
                $copy.Result = '元が見つからない'
                continue
            }
            $copyHash = $null
            $origHash = $null
            [void](Update-ItemInfo $copy.Item)
            [void](Update-ItemInfo $copy.Original)
            $copy.Size = $copy.Item.Size
            # 調べている間になくなったファイルは、サイズを比べても意味がない
            $vanished = $copy.Item.Fail -eq '調査中に変更' -or $copy.Original.Fail -eq '調査中に変更'
            if (-not $vanished) {
                if ($copy.Original.Size -ne $copy.Item.Size) { $copy.Result = '違う'; continue }
                if ($copy.Item.Size -eq 0) { $copy.Result = '同じ'; continue }
                # オンライン専用ファイルは、読むとダウンロードが始まるので読まない（読まずに記録だけ残る）。片方でも比べられないと分かっていれば、もう片方も読まない
                foreach ($target in $copy.Item, $copy.Original) { if ($target.Online) { [void](Get-FullHash $target) } }
                if (-not $copy.Item.Fail -and -not $copy.Original.Fail) {
                    # 大きなファイルは、先頭部分が違えば全体を読まない
                    if ($copy.Item.Size -gt $HeadSize -and -not ($copy.Item.Hash -and $copy.Original.Hash)) {
                        $copyHead = Get-HeadHash $copy.Item
                        $origHead = Get-HeadHash $copy.Original
                        if ($copyHead -and $origHead -and -not [string]::Equals($copyHead, $origHead, [StringComparison]::Ordinal)) { $copy.Result = '違う'; continue }
                    }
                    $copyHash = Get-FullHash $copy.Item
                    if ($copyHash) { $origHash = Get-FullHash $copy.Original }
                }
            }
            if ($copyHash -and $origHash) {
                $copy.Result = if ([string]::Equals($copyHash, $origHash, [StringComparison]::Ordinal)) { '同じ' } else { '違う' }
            } else {
                $copy.Result = '比べられなかった'
                $failed = if ($copy.Item.Fail) { $copy.Item } else { $copy.Original }
                $copy.Note = if ($failed.Fail) { '{0}: {1}' -f $failed.Fail, (ConvertTo-DisplayPath $failed.Full) } else { '' }
            }
        }
        $copyRows = @($copies | Sort-Object -Property Order, @{ Expression = 'Size'; Descending = $true }, Full)
        $copyCounts = @{ '同じ' = 0; '違う' = 0; '元が見つからない' = 0; '比べられなかった' = 0 }
        foreach ($copy in $copyRows) { $copyCounts[$copy.Result]++ }

        #endregion

        #region 結果を組み立てる

        $elapsed = (Get-Date) - $startTime
        $wasteTotal = 0L
        $dupFiles   = 0L
        foreach ($group in $groups) { $wasteTotal += $group.Waste; $dupFiles += $group.Count }
        $unreadCount = 0
        foreach ($value in $skipKinds.Values) { $unreadCount += $value }
        if ($groups.Count -gt 0) { $exitCode = 2 } elseif ($unreadCount -gt 0) { $exitCode = 1 }

        $lines.Add(('重複ファイル検出 (Find-DuplicateFile {0})' -f $ScriptVersion))
        $lines.Add(('対象: {0}' -f ($rootPaths -join ', ')))
        if ($configPath) { $lines.Add(('設定ファイル: {0}' -f $configPath)) }
        $lines.Add((Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（管理者: {2}）' $startTime $machineName $(if ($isAdmin) { 'はい' } else { 'いいえ' })))
        $lines.Add(('条件: 内容を比べるのは {0} 以上のファイル' -f ((Format-Size $minBytes) -replace '\.0 ', ' ')))
        if ($hasInclude) { $lines.Add(('絞り込み: {0}' -f ($includeTexts -join ', '))) }
        if ($hasExclude) { $lines.Add(('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', '))) }
        if ($useCutoff) { $lines.Add((Format-Invariant '条件: 最終更新が {0} 日以上前のファイルのみ（{1:yyyy-MM-dd HH:mm} より前）' $OlderThan $cutoff)) }
        $lines.Add('')
        $lines.Add((Format-Invariant '調べたファイル {0:N0} 件（{1}）  フォルダ {2:N0} 件  所要 {3}' $totalFiles (Format-Size $totalSize) $totalDirs (Format-Duration $elapsed)))
        $lines.Add((Format-Invariant '内容まで比べたファイル {0:N0} 件（{1}）' $comparedFiles (Format-Size $comparedSize)))
        # 保存できなかったときは skipped.csv を見るよう案内できないため、保存の後で作り直せるようにしておく
        $dupLine = {
            param([bool]$HasFile)
            if ($groups.Count -gt 0) {
                Format-Invariant '重複 {0:N0} グループ・{1:N0} 件  減らせる容量 {2}' $groups.Count $dupFiles (Format-Size $wasteTotal)
            } elseif ($unreadCount -gt 0) {
                Format-Invariant '重複 0 グループ（ただし、比べられなかった場所が {0:N0} 件あります{1}）' $unreadCount $(if ($HasFile) { '。skipped.csv を確認してください' } else { '' })
            } else {
                '重複 0 グループ'
            }
        }
        $dupLineIndex = $lines.Count
        $lines.Add((& $dupLine ([bool]$runDir)))
        $copyParts = @(
            (Format-Invariant '元と同じ {0:N0}' $copyCounts['同じ']),
            (Format-Invariant '元と違う {0:N0}' $copyCounts['違う']),
            (Format-Invariant '元が見つからない {0:N0}' $copyCounts['元が見つからない'])
        )
        if ($copyCounts['比べられなかった'] -gt 0) { $copyParts += Format-Invariant '比べられなかった {0:N0}' $copyCounts['比べられなかった'] }
        if ($copyRows.Count -eq 0) {
            $lines.Add('コピーらしい名前 0 件')
        } else {
            $lines.Add((Format-Invariant 'コピーらしい名前 {0:N0} 件（{1}）' $copyRows.Count ($copyParts -join '・')))
        }
        foreach ($line in (Get-NoteLine)) { $lines.Add($line) }
        if ($hardLinkFiles -gt 0) { $lines.Add((Format-Invariant 'ハードリンク {0:N0} 件は、減らせる容量に数えていません' $hardLinkFiles)) }
        if ($identityFailed) { $lines.Add('※ ハードリンクを判定できませんでした。減らせる容量は実際より多く出ることがあります') }
        if ($skipKinds['権限なし'] -gt 0 -and -not $isAdmin) { $lines.Add('※ 管理者として実行すると、読める範囲が広がります') }
        $lines.Add('')
        $headCount = $lines.Count

        # 重複のグループは複数の行になるため、グループごとに行をまとめておく
        $groupBlocks = New-Object 'System.Collections.Generic.List[object]'
        $groupRows   = New-Object 'System.Collections.Generic.List[object]'
        $fileRows    = New-Object 'System.Collections.Generic.List[object]'
        $rank = 0
        foreach ($group in $groups) {
            $rank++
            $countText = Format-Invariant '{0} × {1:N0} 件' (Format-Size $group.Size) $group.Count
            if ($group.Entities -ne $group.Count) { $countText += Format-Invariant '、実体 {0:N0} つ' $group.Entities }
            $block = New-Object 'System.Collections.Generic.List[string]'
            $block.Add((Format-Invariant ' {0}. 減らせる容量 {1}（{2}）' $rank (Format-Size $group.Waste) $countText))
            foreach ($item in $group.Items) {
                $display = ConvertTo-DisplayPath $item.Full
                $block.Add((Format-Invariant '    {0:yyyy-MM-dd HH:mm}  {1}' $item.Time.ToLocalTime() $display))
                $fileRows.Add([object[]]@($rank, $item.Entity, $display, (Protect-ExcelText $item.Name), (Format-Megabytes $item.Size), (Format-Bytes $item.Size), (Format-Time $item.Time), $group.Hash))
            }
            $groupBlocks.Add($block)
            $groupRows.Add([object[]]@($rank, $group.Count.ToString('N0', $Invariant), $group.Entities.ToString('N0', $Invariant), (Format-Megabytes $group.Size), (Format-Bytes $group.Size), (Format-Megabytes $group.Waste), (Format-Bytes $group.Waste), $group.Hash))
        }

        $copyHeader = ' ' + (Format-Column '名前の形' 8 $false) + '  ' + (Format-Column '元との比較' 16 $false) + '  ' + (Format-Column 'サイズ' 10 $true) + '  ' + (Format-Column '最終更新' 16 $false) + '  ファイル'
        $copyLines = New-Object 'System.Collections.Generic.List[string]'
        $copyCsv   = New-Object 'System.Collections.Generic.List[object]'
        foreach ($copy in $copyRows) {
            $display = ConvertTo-DisplayPath $copy.Item.Full
            $copyLines.Add(' ' + (Format-Column $copy.Form 8 $false) + '  ' + (Format-Column $copy.Result 16 $false) + '  ' + (Format-Column (Format-Size $copy.Size) 10 $true) + '  ' +
                (Format-Column (Format-Time $copy.Item.Time 'yyyy-MM-dd HH:mm') 16 $false) + '  ' + $display)
            $original = $copy.Original
            $copyCsv.Add([object[]]@($copy.Form, $copy.Result, $display, $(if ($original) { ConvertTo-DisplayPath $original.Full } else { '' }),
                (Format-Megabytes $copy.Size), (Format-Bytes $copy.Size), $(if ($original) { Format-Bytes $original.Size } else { '' }),
                (Format-Time $copy.Item.Time), $(if ($original) { Format-Time $original.Time } else { '' }), $copy.Note))
        }

        # summary.txt には、すべて書く
        $lines.Add('■ 減らせる容量の大きいグループ')
        if ($groupBlocks.Count -eq 0) { $lines.Add('    （重複はありません）') }
        foreach ($block in $groupBlocks) { foreach ($line in $block) { $lines.Add($line) } }
        $lines.Add('')
        $lines.Add('■ コピーらしい名前のファイル')
        if ($copyLines.Count -eq 0) {
            $lines.Add('    （ありません）')
        } else {
            $lines.Add($copyHeader)
            foreach ($line in $copyLines) { $lines.Add($line) }
        }
        $lines.Add('')
        if ($runDir) { $lines.Add(('保存先: {0}' -f $runDir)) }

        #endregion

        #region 保存

        if ($runDir) {
            Write-Progress -Activity '重複ファイル検出' -Status '結果を保存しています'
            try {
                [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'groups.csv')) $GroupColumns $groupRows
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'duplicates.csv')) $FileColumns $fileRows
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'copynames.csv')) $CopyColumns $copyCsv
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $skipped
            } catch {
                $saveError = Get-ErrorText (Get-InnerException $_.Exception)
            }
            Write-Progress -Activity '重複ファイル検出' -Completed
        }

        #endregion
    }

    $finished = $true
} catch {
    # 想定外のエラーは、Ctrl+C の「中断しました」と区別して知らせる
    $unexpected = $true
    throw
} finally {
    Write-Progress -Activity '重複ファイル検出' -Completed
    if (-not $finished -or $fatal -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}
if ($fatal) { Stop-WithError $fatal }

#region 画面表示

if ($saveError) { $lines[$dupLineIndex] = & $dupLine $false }
for ($i = 0; $i -lt $headCount; $i++) { Write-Host $lines[$i] }

Write-Host (Format-Invariant '■ 減らせる容量の大きいグループ（上位 {0}）' $Top) -ForegroundColor Cyan
if ($groupBlocks.Count -eq 0) { Write-Host '    （重複はありません）' }
$shown = 0
foreach ($block in $groupBlocks) {
    if ($shown -ge $Top) { break }
    foreach ($line in $block) { Write-Host $line }
    $shown++
}
if ($groupBlocks.Count -gt $shown) {
    $where = if ($runDir -and -not $saveError) { 'は groups.csv を参照してください' } else { '' }
    Write-Host (Format-Invariant '（ほか {0:N0} グループ{1}）' ($groupBlocks.Count - $shown) $where)
}
Write-Host ''

Write-Host (Format-Invariant '■ コピーらしい名前のファイル（上位 {0}）' $Top) -ForegroundColor Cyan
if ($copyLines.Count -eq 0) {
    Write-Host '    （ありません）'
} else {
    Write-Host $copyHeader
    $shown = 0
    foreach ($line in $copyLines) {
        if ($shown -ge $Top) { break }
        Write-Host $line
        $shown++
    }
    if ($copyLines.Count -gt $shown) {
        $where = if ($runDir -and -not $saveError) { 'は copynames.csv を参照してください' } else { '' }
        Write-Host (Format-Invariant '（ほか {0:N0} 件{1}）' ($copyLines.Count - $shown) $where)
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
