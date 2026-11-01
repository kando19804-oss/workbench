<#
.SYNOPSIS
    2 つのフォルダの中身を比べて、追加・削除・変更されたファイルとフォルダを一覧にします。

.DESCRIPTION
    基準（Reference）と比較（Difference）のフォルダを、相対パスで突き合わせて比べます。
    比較にだけあるものを「追加」、基準にだけあるものを「削除」、サイズや更新日時（-Hash のときは内容）が
    違うファイルを「変更」として画面に表示し、CSV（BOM 付き UTF-8）にも保存します。

    作業の前にフォルダの状態を保存しておき（-Snapshot）、作業の後に今のフォルダと比べることもできます。

    比べるフォルダのファイルを変更・削除することはありません。
    権限がなくて読めないフォルダはスキップし、件数を表示します。
    ジャンクションやシンボリックリンクはたどりません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Compare-Folder.ps1 -Reference D:\data -Difference E:\backup\data

.PARAMETER Reference
    基準のフォルダ。ドライブ（D:\）や共有フォルダ（\\server\share）も指定できます。
    -Snapshot で保存したスナップショットのファイル（snapshot.csv）も指定できます。

.PARAMETER Difference
    比較するフォルダ。-Snapshot のとき以外は必須です。

.PARAMETER Hash
    ファイルの内容を SHA-256 のハッシュで比べます。ファイルを最後まで読むため、時間がかかります。
    指定しないときは、サイズと更新日時で比べます。

.PARAMETER Exclude
    比べないフォルダやファイル。ワイルドカード（*）で指定します。
    \ を含まない書き方は名前に対して（どの階層でも）、\ を含む書き方は相対パスに対して当てはめます。
    * 以外の文字（[ ] ? など）は、そのままの文字として扱います。大文字と小文字は区別しません。
    複数指定するときは、カンマで区切ります。

.PARAMETER IncludeSame
    違いのなかったファイルとフォルダも、differences.csv に書きます。

.PARAMETER Snapshot
    Reference のフォルダの状態を snapshot.csv に保存するだけで、比べません。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。-Snapshot と一緒には使えません。

.PARAMETER ConfigFile
    設定ファイルのパス。Reference、Difference、Hash、Exclude、IncludeSame、OutputFolder、NoSave の値を
    「項目 = 値」の形で書いておけます。コマンドラインで指定した値は、設定ファイルより優先します。
    書き方は README を参照してください。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Compare-Folder.ps1 -Reference D:\data -Difference E:\backup\data -Hash

    2 つのフォルダを、ファイルの内容まで比べます。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Compare-Folder.ps1 -Reference D:\app -Snapshot

    D:\app の状態を保存します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Compare-Folder.ps1 -Reference .\FolderSnapshot_SERVER01_20261031_220000\snapshot.csv -Difference D:\app

    保存しておいた状態と、今の D:\app を比べます。

.NOTES
    Compare-Folder 1.0.0
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
    [string]$Reference,

    [Parameter(Position = 1)]
    [string]$Difference,

    [switch]$Hash,

    [string[]]$Exclude = @(),

    [switch]$IncludeSame,

    [switch]$Snapshot,

    [string]$OutputFolder,

    [switch]$NoSave,

    [string]$ConfigFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion  = '1.0.0'
$ScreenRows     = 200       # 画面に表示する違いの件数（保存するときだけ。超えた分はファイルを見てもらう）
$TimeTolerance  = 2         # 更新日時の違いを同じとみなす秒数（FAT32 や exFAT は 2 秒単位で記録するため）
$ProgressMs     = 250       # 進捗表示の間隔（ミリ秒）
$HashBufferSize = 1MB       # ハッシュを計算するときに 1 回で読む長さ
$OutputPrefixes = 'FolderDiff_', 'FolderSnapshot_'
$SnapshotTool   = 'Compare-Folder'
$SnapshotColumns = '種類', '相対パス', 'サイズ(バイト)', '最終更新日時', 'ハッシュ', '状態'
$DiffColumns    = '結果', '種類', '相対パス', '違い', '基準のサイズ(バイト)', '比較のサイズ(バイト)',
                  '基準の最終更新日時', '比較の最終更新日時', '基準のハッシュ', '比較のハッシュ', '備考'
$SkipColumns    = '理由', '側', 'パス', '詳細'
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
# 数値や日付に変えてしまうため、その形の値は ="..." で囲む。式に書けない 255 文字を超える値は、先頭に空白を足す。
# スナップショットは読み戻すため、長さにかかわらず ="..." にする（$Always）
$excelRisk = New-Object Text.RegularExpressions.Regex ('^(?:[=+\-@]|[\d\s.,/:+\-]+$)')
function Protect-ExcelText([string]$Text, [bool]$Always) {
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
    if ($Always -or $Text.Length -le 255) { return '="' + $Text + '"' }
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
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -Reference "C:\data"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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
function Assert-Folder([string]$FullPath, [string]$Label, [string]$ConfigName) {
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
        $message = '{0}のフォルダを開けません: {1}' -f $Label, $FullPath
        if ($reason) { $message += '（{0}）' -f $reason }
        $message += Get-ConfigNote $ConfigName
        Stop-WithError ((@($message) + (Get-NetworkHint $FullPath)) -join "`r`n")
    }
}

# 結果の表示用に、引数を必要なときだけ引用符で囲む。"D:\data\" のように最後が \" になると
# 引用符が閉じなくなるため、最後の \ は 2 つにする
function Format-Argument([string]$Text) {
    if ($Text -notmatch '[\s&()^,;=!%]') { return $Text }
    if ($Text.EndsWith('\')) { $Text += '\' }
    return '"' + $Text + '"'
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの。Snapshot は、同じ設定ファイルを保存と比較の両方に使えるよう書けない
$configItems  = 'Reference', 'Difference', 'Hash', 'Exclude', 'IncludeSame', 'OutputFolder', 'NoSave'
$configFlags  = 'Hash', 'IncludeSame', 'NoSave'
$configPath   = $null
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
        # エクスプローラーで拡張子を表示していないと、メモ帳で migrate.ini.txt として保存してしまうことがある
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
    $excludeItems = New-Object 'System.Collections.Generic.List[string]'
    $lineNumber   = 0
    foreach ($rawLine in $configLines) {
        $lineNumber++
        $lineText = $rawLine.Trim()
        # # と ; はフォルダ名に使える文字なので、行の先頭にあるときだけコメントとみなす
        if ($lineText -eq '' -or $lineText.StartsWith('#') -or $lineText.StartsWith(';')) { continue }
        $separator = $lineText.IndexOf('=')
        if ($separator -le 0) { Stop-WithConfigError $lineNumber ('「項目 = 値」の形になっていません: {0}' -f $lineText) }

        $itemKey  = $lineText.Substring(0, $separator).Trim()
        if ($itemKey -eq 'Snapshot') {
            Stop-WithConfigError $lineNumber 'Snapshot は設定ファイルには書けません。コマンドラインで -Snapshot を指定してください'
        }
        $itemName = $configItems | Where-Object { $_ -eq $itemKey }
        if (-not $itemName) { Stop-WithConfigError $lineNumber ('知らない項目です: {0}' -f $itemKey) ('使える項目: {0}' -f ($configItems -join ', ')) }

        $rawValue  = $lineText.Substring($separator + 1).Trim()
        $itemValue = Remove-ConfigQuote $rawValue
        # 値が空の行は、書かなかったのと同じ扱い（ひな形の空の行の下に書き足しても、2 回書いたことにしない）
        if ($itemValue -eq '') { continue }

        # Exclude は長くなりやすいので複数の行に分けて書ける。ほかの項目は、どちらが効いているか分からなくなるため 1 回だけ
        if ($itemName -ne 'Exclude' -and $configLineOf.ContainsKey($itemName)) {
            Stop-WithConfigError $lineNumber ('{0} が 2 回書かれています（{1} 行目と {2} 行目）' -f $itemName, $configLineOf[$itemName], $lineNumber)
        }
        $configLineOf[$itemName] = $lineNumber

        if ($itemName -eq 'Exclude') {
            # "a,b" のように全体を囲む書き方と、"a","b" のように 1 つずつ囲む書き方のどちらも受け付ける
            $source = if ($itemValue.Contains('"')) { $rawValue } else { $itemValue }
            foreach ($part in ($source -split ',')) {
                $pattern = Remove-ConfigQuote $part
                if ($pattern.Contains('"')) { Stop-WithConfigError $lineNumber ('Exclude の値に引用符（"）が残っています: {0}' -f $pattern) }
                if ($pattern) { $excludeItems.Add((Expand-ConfigVariable $lineNumber $pattern)) }
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
    if ($excludeItems.Count -gt 0) { $configValues['Exclude'] = $excludeItems.ToArray() }

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

# コマンドプロンプトで -Reference "C:\data\" と書くと、後ろの引数まで -Reference の値に入り、-Difference が空になる。
# 本当の原因（最後の \"）を知らせるため、空かどうかより先に引用符を確かめる
$refText = Remove-Quote $Reference
if ($refText.Contains('"')) { Stop-WithQuoteError '-Reference' $refText }
$difText = Remove-Quote $Difference
if ($difText.Contains('"')) { Stop-WithQuoteError '-Difference' $difText }
$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    $outText = Resolve-StartPath $outText
}

# 設定ファイルにも書けるため、どちらで指定するかを案内する
if (-not $refText) { Stop-WithError '基準のフォルダを指定してください（-Reference で指定するか、設定ファイルの Reference に書きます）' }
if ($Snapshot) {
    # 設定ファイルの Difference は、同じ設定ファイルを比較にも使えるよう、-Snapshot のときは使わない
    if ($PSBoundParameters.ContainsKey('Difference')) { Stop-WithError '-Snapshot のときは、-Difference は指定できません' }
    if ($NoSave) {
        $message = '-Snapshot のときは、-NoSave は指定できません'
        if ($configApplied.ContainsKey('NoSave')) { $message += '（設定ファイルの {0} 行目に NoSave = true があります）' -f $configApplied['NoSave'] }
        Stop-WithError $message
    }
    $difText = ''
} elseif (-not $difText) {
    Stop-WithError '比較のフォルダを指定してください（-Difference で指定するか、設定ファイルの Difference に書きます）'
}

if ($refText -match '^(\\\\\?\\)?[A-Za-z]:$') { $refText += '\' }
if ($difText -match '^(\\\\\?\\)?[A-Za-z]:$') { $difText += '\' }
$refPath = Resolve-FullPath (Resolve-StartPath $refText)
$difPath = if ($difText) { Resolve-FullPath (Resolve-StartPath $difText) } else { $null }

# -Exclude は、-File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# コマンドプロンプトでは ' がそのまま届くため、'*.log,logs' のように全体を囲んだ場合も、各要素の前後の ' を外す
$excludeTexts = New-Object 'System.Collections.Generic.List[string]'
foreach ($item in $Exclude) {
    foreach ($part in ($item -split ',')) {
        $text = (Remove-Quote $part).Trim("'")
        if ($text.Contains('"')) { Stop-WithQuoteError '-Exclude' $text }
        $text = $text.TrimEnd('\')
        if ($text -and -not ($excludeTexts -contains $text)) { $excludeTexts.Add($text) }
    }
}

# 基準がファイルなら、スナップショットとして読む
$refIsSnapshot = $false
try { $refIsSnapshot = [IO.File]::Exists((ConvertTo-LongPath $refPath)) } catch { }
if ($Snapshot -and $refIsSnapshot) { Stop-WithError '-Snapshot のときは、-Reference にフォルダを指定してください' }
if (-not $refIsSnapshot) { Assert-Folder $refPath '基準' 'Reference' }
if ($difPath) { Assert-Folder $difPath '比較' 'Difference' }

# ネットワークドライブと UNC パスのように、同じ場所を別の名前で指定した場合は判定できない
if ($difPath -and -not $refIsSnapshot -and $refPath.TrimEnd('\') -eq $difPath.TrimEnd('\')) {
    Stop-WithError ('基準と比較が同じフォルダです: {0}' -f $refPath)
}

#endregion

#region スナップショットの読み込み

# スナップショットの先頭には、種類が「情報」の行で、保存したときのフォルダ・日時・ハッシュの有無・除外を書いてある
$snapInfo = $null
$snapRows = $null
if ($refIsSnapshot) {
    $snapError = $null
    $records   = $null
    try {
        $snapText = (Read-ConfigText (ConvertTo-LongPath $refPath)) -split '\r\n|\n|\r' | Where-Object { $_ -ne '' }
        $records  = @($snapText | ConvertFrom-Csv)
    } catch {
        $snapError = Get-ErrorText (Get-InnerException $_.Exception)
    }
    if ($snapError) { Stop-WithError ('スナップショットのファイルを読めません: {0}（{1}）' -f $refPath, $snapError) }

    $columnsOk = $records.Count -gt 0 -and
        (@($records[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',') -eq ($SnapshotColumns -join ',')
    if (-not $columnsOk -or $records[0].'種類' -ne '情報' -or $records[0].'相対パス' -ne 'ツール' -or -not $records[0].'状態'.StartsWith($SnapshotTool)) {
        Stop-WithError ('スナップショットのファイルとして読めません: {0}（Compare-Folder の -Snapshot で保存した snapshot.csv を指定してください）' -f $refPath)
    }

    $snapInfo = @{ Folder = ''; Saved = ''; HasHash = $false; Exclude = @() }
    $snapRows = New-Object 'System.Collections.Generic.List[object]'
    $rowNumber = 1   # 見出しが 1 行目
    foreach ($record in $records) {
        $rowNumber++
        $kind = $record.'種類'
        if ($kind -eq '情報') {
            switch ($record.'相対パス') {
                'フォルダ' { $snapInfo.Folder = $record.'状態' }
                '保存日時' { $snapInfo.Saved = $record.'状態' }
                'ハッシュ' { $snapInfo.HasHash = $record.'状態' -eq 'SHA-256' }
                '除外'     { $snapInfo.Exclude += $record.'状態' }
            }
            continue
        }
        $rel = $record.'相対パス'
        if ($rel.StartsWith('="') -and $rel.EndsWith('"')) { $rel = $rel.Substring(2, $rel.Length - 3) }
        $size   = 0L
        $timeUtc = $null
        $valid  = ($kind -eq 'ファイル' -or $kind -eq 'フォルダ' -or $kind -eq 'リンク') -and $rel -ne ''
        if ($valid -and $kind -eq 'ファイル') {
            $parsedTime = [datetime]::MinValue
            $valid = [long]::TryParse($record.'サイズ(バイト)'.Replace(',', ''), [Globalization.NumberStyles]::None, $Invariant, [ref]$size) -and
                [datetime]::TryParseExact($record.'最終更新日時', $TimeFormat, $Invariant, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsedTime)
            if ($valid) { $timeUtc = $parsedTime.ToUniversalTime() }
        }
        if (-not $valid) {
            Stop-WithError ('スナップショットのファイルとして読めません: {0}（{1} 行目が正しくありません。Excel で上書き保存すると、形が変わることがあります）' -f $refPath, $rowNumber)
        }
        $snapRows.Add([pscustomobject]@{ Kind = $kind; Rel = $rel; Size = $size; Time = $timeUtc; Hash = $record.'ハッシュ'; State = $record.'状態' })
    }

    if ($Hash -and -not $snapInfo.HasHash) {
        Stop-WithError 'このスナップショットにはハッシュがありません。-Hash を付けずに実行してください'
    }
    # スナップショットを保存したときの除外は、比べるときにも使う（付け忘れると、除外したものがすべて「追加」に出るため）
    foreach ($text in $snapInfo.Exclude) {
        if (-not ($excludeTexts -contains $text)) { $excludeTexts.Add($text) }
    }
}

$useHash = [bool]$Hash -or ($refIsSnapshot -and $snapInfo.HasHash)

# ワイルドカードは * だけにして、それ以外の文字は文字どおりに扱う。5.1 の WildcardPattern.Escape は
# バッククォートをエスケープしないため、自分でエスケープしてから照合用のパターンを作る。
# \ を含む書き方は相対パスに、含まない書き方は名前に当てはめる。先頭の \ は「直下から」の意味なので外す
$wildcardOptions = [Management.Automation.WildcardOptions]'IgnoreCase, CultureInvariant'
$excludeRules = @(
    foreach ($text in $excludeTexts) {
        $body    = $text.TrimStart('\')
        $escaped = $body.Replace('`', '``').Replace('[', '`[').Replace(']', '`]').Replace('?', '`?')
        [pscustomobject]@{ Text = $text; ByPath = $text.Contains('\'); Matcher = [Management.Automation.WildcardPattern]::new($escaped, $wildcardOptions) }
    }
)
$hasExclude = $excludeRules.Count -gt 0

$startTime = Get-Date

#endregion

#region 判定の部品

$dirType  = [IO.DirectoryInfo]
$rpFlag   = [IO.FileAttributes]::ReparsePoint
$hsFlag   = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System

$skipped   = New-Object 'System.Collections.Generic.List[object]'
$skipKinds = [ordered]@{ '権限なし' = 0; 'その他のエラー' = 0; '読めない' = 0 }
$excludedCount = 0
$progress  = @{ Next = 0L; Done = 0L; Total = 0L; Files = 0; FileTotal = 0 }
$sw        = [Diagnostics.Stopwatch]::StartNew()

function Add-Skip([string]$Reason, [string]$Side, [string]$ItemPath, [string]$Detail) {
    $skipped.Add([object[]]@($Reason, $Side, $ItemPath, $Detail))
    if ($skipKinds.Contains($Reason)) { $skipKinds[$Reason]++ }
}

# 除外に当てはまれば、当てはまった書き方を返す
function Get-ExcludeMatch([string]$Rel, [string]$Name) {
    foreach ($rule in $excludeRules) {
        if ($rule.ByPath) { if ($rule.Matcher.IsMatch($Rel)) { return $rule.Text } }
        elseif ($rule.Matcher.IsMatch($Name)) { return $rule.Text }
    }
    return $null
}

# スナップショットの行は、親のフォルダが除外に当てはまるときも除く（フォルダをたどるときは、親で止まるため要らない）
function Get-ExcludeMatchWithParent([string]$Rel) {
    $parts = $Rel.Split('\')
    $path  = ''
    foreach ($part in $parts) {
        $path = if ($path) { $path + '\' + $part } else { $part }
        $matched = Get-ExcludeMatch $path $part
        if ($matched) { return $matched }
    }
    return $null
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

function Show-Progress([string]$Status, [string]$Operation) {
    if ($sw.ElapsedMilliseconds -lt $progress.Next) { return }
    $percent = -1
    if ($progress.Total -gt 0) { $percent = [int][math]::Min(100, 100 * $progress.Done / $progress.Total) }
    Write-Progress -Activity 'フォルダ差分' -Status $Status -CurrentOperation $Operation -PercentComplete $percent
    $progress.Next = $sw.ElapsedMilliseconds + $ProgressMs
}

# フォルダの中をすべて読み、相対パス → 項目 の表を返す。中を読めなかったフォルダは State に理由を入れる。
# 再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる
function Read-FolderState([string]$RootPath, [string]$Side) {
    $items     = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
    $totals    = @{ Files = 0L; Dirs = 0L; Size = 0L; Fatal = $null }
    $rootLong  = ConvertTo-LongPath $RootPath
    $rootTrim  = $rootLong.TrimEnd([char[]]@('\', '/'))
    $prefixLen = $rootTrim.Length + 1
    $stack = New-Object 'System.Collections.Generic.Stack[object]'
    $stack.Push(@((New-Object IO.DirectoryInfo $rootLong), ''))

    while ($stack.Count -gt 0) {
        $node    = $stack.Pop()
        $dir     = $node[0]
        $dirRel  = $node[1]
        Show-Progress (Format-Invariant '{0}を調べています: フォルダ {1:N0} 件 / ファイル {2:N0} 件' $Side $totals.Dirs $totals.Files) (ConvertTo-DisplayPath $dir.FullName)

        try {
            $entries = $dir.GetFileSystemInfos()
        } catch {
            $ex = Get-InnerException $_.Exception
            if ($dirRel -eq '') {
                $totals.Fatal = '{0}のフォルダを読めません: {1}（{2}）' -f $Side, $RootPath, (Get-ErrorText $ex)
                break
            }
            $reason = if ($ex -is [UnauthorizedAccessException]) { '権限なし' } else { 'その他のエラー' }
            $detail = if ($ex -is [UnauthorizedAccessException]) { 'アクセスが拒否されました' } else { Get-ErrorText $ex }
            $items[$dirRel].State = $reason
            Add-Skip $reason $Side (ConvertTo-DisplayPath $dir.FullName) $detail
            continue
        }

        foreach ($entry in $entries) {
            $rel = $entry.FullName.Substring($prefixLen)
            if (-not $OnWindows) { $rel = $rel.Replace('/', '\') }
            $isDir = $entry -is $dirType
            # 前回の結果やスナップショットを比べないよう、このツールの結果のフォルダは中を見ない
            if ($isDir) {
                $isOutput = $false
                foreach ($prefix in $OutputPrefixes) { if ($entry.Name.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { $isOutput = $true } }
                if ($isOutput) { continue }
            }
            if ($hasExclude) {
                $matched = Get-ExcludeMatch $rel $entry.Name
                if ($matched) {
                    $script:excludedCount++
                    Add-Skip '除外指定' $Side (ConvertTo-DisplayPath $entry.FullName) $matched
                    continue
                }
            }
            $linkName = Get-LinkName $entry
            if ($linkName) {
                $items[$rel] = [pscustomobject]@{ Kind = 'リンク'; Rel = $rel; Size = 0L; Time = $null; Full = $entry.FullName; Hash = ''; State = ''; Note = $linkName; Online = $false }
            } elseif ($isDir) {
                $items[$rel] = [pscustomobject]@{ Kind = 'フォルダ'; Rel = $rel; Size = 0L; Time = $null; Full = $entry.FullName; Hash = ''; State = ''; Note = ''; Online = $false }
                $totals.Dirs++
                $stack.Push(@($entry, $rel))
            } else {
                $length = $entry.Length
                $online = ([int]$entry.Attributes -band $OnlineMask) -ne 0
                $items[$rel] = [pscustomobject]@{ Kind = 'ファイル'; Rel = $rel; Size = $length; Time = $entry.LastWriteTimeUtc; Full = $entry.FullName; Hash = ''; State = ''; Note = ''; Online = $online }
                $totals.Files++
                $totals.Size += $length
            }
        }
    }
    return @{ Items = $items; Totals = $totals }
}

# SHA-256 のハッシュを、Get-FileHash と同じ大文字の 16 進数で返す。開けないときは $null を返し、skipped に残す。
# ほかのプログラムが書き込み中のファイルも読めるように開く。FIPS モードの PC でも使えるよう、CryptoServiceProvider を使う
$hashBuffer = New-Object 'byte[]' $HashBufferSize
function Get-HashText($Item, [string]$Side) {
    $display = ConvertTo-DisplayPath $Item.Full
    $sha     = New-Object Security.Cryptography.SHA256CryptoServiceProvider
    $stream  = $null
    try {
        $stream = New-Object IO.FileStream($Item.Full, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]'ReadWrite, Delete'), 65536)
        while ($true) {
            $read = $stream.Read($hashBuffer, 0, $hashBuffer.Length)
            if ($read -le 0) { break }
            [void]$sha.TransformBlock($hashBuffer, 0, $read, $null, 0)
            $progress.Done += $read
            Show-Progress (Format-Invariant 'ハッシュを計算しています: {0:N0} / {1:N0} 件（{2} / {3}）' $progress.Files $progress.FileTotal (Format-Size $progress.Done) (Format-Size $progress.Total)) $display
        }
        [void]$sha.TransformFinalBlock($hashBuffer, 0, 0)
        return [BitConverter]::ToString($sha.Hash).Replace('-', '')
    } catch {
        $ex = Get-InnerException $_.Exception
        Add-Skip '読めない' $Side $display (Get-ErrorText $ex)
        return $null
    } finally {
        if ($stream) { $stream.Dispose() }
        $sha.Dispose()
        $progress.Files++
    }
}

function Format-Time($TimeUtc) {
    if ($null -eq $TimeUtc) { return '' }
    return $TimeUtc.ToLocalTime().ToString($TimeFormat, $Invariant)
}

function Format-Bytes($Item) {
    if ($null -eq $Item -or $Item.Kind -ne 'ファイル') { return '' }
    return $Item.Size.ToString('N0', $Invariant)
}

# スキップ・リンク・除外の件数の行と、管理者として実行するとよいことの案内
function Get-NoteLine([int]$LinkCount) {
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
    if ($LinkCount -gt 0) { $notes += Format-Invariant 'たどらなかったリンク {0:N0} 件' $LinkCount }
    if ($excludedCount -gt 0) { $notes += Format-Invariant '除外 {0:N0} 件' $excludedCount }
    if ($notes.Count -gt 0) { $result += $notes -join '  ' }
    if ($skipKinds['権限なし'] -gt 0 -and -not $isAdmin) { $result += '※ 管理者として実行すると、読める範囲が広がります' }
    return $result
}

#endregion

#region 保存先の準備

# 比べ終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir      = $null
$saveWarning = $null
if (-not $NoSave) {
    $baseDir = if ($outText) { $outText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    $runPrefix = if ($Snapshot) { 'FolderSnapshot_' } else { 'FolderDiff_' }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        $runName = Format-Invariant '{0}{1}_{2:yyyyMMdd_HHmmss}' $runPrefix $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while (Test-Path -LiteralPath $runDir) { $suffix++; $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix) }
        [void][IO.Directory]::CreateDirectory($runDir)
    } catch {
        $reason = Get-ErrorText (Get-InnerException $_.Exception)
        $runDir = $null
        # スナップショットは、保存できなければ意味がないため止める
        if ($Snapshot) { Stop-WithError (('保存先に書き込めないため、スナップショットを保存できません: {0}（{1}）' -f $baseDir, $reason) + (Get-ConfigNote 'OutputFolder')) }
        $saveWarning = ('保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, $reason) + (Get-ConfigNote 'OutputFolder')
        Write-Warning $saveWarning
    }
}

#endregion

# 調べている途中や保存の途中で Ctrl+C されたときや、保存に失敗したときは、書きかけの結果を残さない
$finished  = $false
$fatal     = $null
$saveError = $null
$exitCode  = 0
$screen    = New-Object 'System.Collections.Generic.List[object]'   # 画面に出す行（文字, 色）
$lines     = New-Object 'System.Collections.Generic.List[string]'   # summary.txt に書く行
try {
    #region 調べる

    $refState = $null
    if (-not $refIsSnapshot) {
        $refState = Read-FolderState $refPath '基準'
        if ($refState.Totals.Fatal) { $fatal = $refState.Totals.Fatal }
    }
    $difState = $null
    if (-not $fatal -and $difPath) {
        $difState = Read-FolderState $difPath '比較'
        if ($difState.Totals.Fatal) { $fatal = $difState.Totals.Fatal }
    }

    #endregion

    if (-not $fatal -and $Snapshot) {
        #region スナップショットの保存

        $items = @($refState.Items.Values)
        if ($useHash) {
            foreach ($item in $items) {
                if ($item.Kind -eq 'ファイル' -and -not $item.Online) { $progress.Total += $item.Size; $progress.FileTotal++ }
            }
            foreach ($item in $items) {
                if ($item.Kind -ne 'ファイル' -or $item.Online) { continue }
                $value = Get-HashText $item '基準'
                if ($null -eq $value) { $item.State = '読めない' } else { $item.Hash = $value }
            }
        }
        $elapsed = $sw.Elapsed
        Write-Progress -Activity 'フォルダ差分' -Completed

        # オブジェクトの配列を [Array]::Sort に渡すと、PowerShell が別の配列に変えて渡すため並べ替わらない。相対パスの配列を並べ替えて引き直す
        $rels = New-Object 'string[]' $refState.Items.Count
        $refState.Items.Keys.CopyTo($rels, 0)
        $sortKeys = New-Object 'string[]' $rels.Count
        for ($i = 0; $i -lt $rels.Count; $i++) { $sortKeys[$i] = $rels[$i].Replace('\', [char]1) }
        [Array]::Sort($sortKeys, $rels, [StringComparer]::OrdinalIgnoreCase)
        $items = @(foreach ($rel in $rels) { $refState.Items[$rel] })

        $totals = $refState.Totals
        $lines.Add((Format-Invariant 'フォルダ差分 (Compare-Folder {0})' $ScriptVersion))
        $lines.Add(('スナップショットを保存します: {0}' -f $refPath))
        if ($configPath) { $lines.Add(('設定ファイル: {0}' -f $configPath)) }
        $lines.Add((Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（管理者: {2}）' $startTime $machineName $(if ($isAdmin) { 'はい' } else { 'いいえ' })))
        $lines.Add($(if ($useHash) { '保存する内容: サイズと更新日時とハッシュ（SHA-256）' } else { '保存する内容: サイズと更新日時' }))
        if ($hasExclude) { $lines.Add(('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', '))) }
        $lines.Add('')
        $lines.Add((Format-Invariant 'ファイル {0:N0} 件  フォルダ {1:N0} 件  合計 {2}  所要 {3}' $totals.Files $totals.Dirs (Format-Size $totals.Size) (Format-Duration $elapsed)))
        $linkItems = @($items | Where-Object { $_.Kind -eq 'リンク' }).Count
        foreach ($line in (Get-NoteLine $linkItems)) { $lines.Add($line) }

        $rows = New-Object 'System.Collections.Generic.List[object]'
        $rows.Add([object[]]@('情報', 'ツール', '', '', '', ('{0} {1}' -f $SnapshotTool, $ScriptVersion)))
        $rows.Add([object[]]@('情報', 'フォルダ', '', '', '', $refPath))
        $rows.Add([object[]]@('情報', '保存日時', '', '', '', $startTime.ToString($TimeFormat, $Invariant)))
        $rows.Add([object[]]@('情報', 'ハッシュ', '', '', '', $(if ($useHash) { 'SHA-256' } else { 'なし' })))
        foreach ($rule in $excludeRules) { $rows.Add([object[]]@('情報', '除外', '', '', '', $rule.Text)) }
        foreach ($item in $items) {
            $state = $item.State
            if (-not $state -and $item.Online) { $state = 'オンライン専用' }
            $rows.Add([object[]]@($item.Kind, (Protect-ExcelText $item.Rel $true), (Format-Bytes $item), (Format-Time $item.Time), $item.Hash, $state))
        }

        #endregion
    } elseif (-not $fatal) {
        #region 突き合わせ

        if ($refIsSnapshot) {
            # スナップショットの行を、フォルダを読んだときと同じ形の表にする
            $refItems = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
            $refTotals = @{ Files = 0L; Dirs = 0L; Size = 0L }
            foreach ($row in $snapRows) {
                if ($hasExclude) {
                    $matched = Get-ExcludeMatchWithParent $row.Rel
                    if ($matched) { continue }
                }
                $state = if ($row.State -eq '権限なし' -or $row.State -eq 'その他のエラー') { $row.State } else { '' }
                $refItems[$row.Rel] = [pscustomobject]@{ Kind = $row.Kind; Rel = $row.Rel; Size = $row.Size; Time = $row.Time; Full = ''; Hash = $row.Hash; State = $state; Note = ''; Online = $row.State -eq 'オンライン専用' }
                if ($row.Kind -eq 'ファイル') { $refTotals.Files++; $refTotals.Size += $row.Size }
                elseif ($row.Kind -eq 'フォルダ') { $refTotals.Dirs++ }
                if ($state) {
                    Add-Skip $state '基準' ($snapInfo.Folder.TrimEnd('\') + '\' + $row.Rel) 'スナップショットを保存したときに読めなかったフォルダ'
                }
            }
        } else {
            $refItems  = $refState.Items
            $refTotals = $refState.Totals
        }
        $difItems  = $difState.Items
        $difTotals = $difState.Totals

        # 片方でも中を読めなかったフォルダは、その中を比べない（読めなかった中身が、すべて追加や削除に見えるため）
        $blocked = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($table in $refItems, $difItems) {
            foreach ($item in $table.Values) { if ($item.State) { [void]$blocked.Add($item.Rel) } }
        }

        $keySet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($key in $refItems.Keys) { [void]$keySet.Add($key) }
        foreach ($key in $difItems.Keys) { [void]$keySet.Add($key) }
        $keys = New-Object 'string[]' $keySet.Count
        $keySet.CopyTo($keys)
        if ($blocked.Count -gt 0) {
            $keys = @(foreach ($key in $keys) {
                $isBlocked = $false
                $cut = $key.LastIndexOf('\')
                while ($cut -gt 0) {
                    $parent = $key.Substring(0, $cut)
                    if ($blocked.Contains($parent)) { $isBlocked = $true; break }
                    $cut = $parent.LastIndexOf('\')
                }
                if (-not $isBlocked) { $key }
            })
        }
        # フォルダの中身がフォルダのすぐ後に並ぶよう、\ をどの文字よりも前に並ぶ文字に置き換えて並べる
        $sortKeys = New-Object 'string[]' $keys.Count
        for ($i = 0; $i -lt $keys.Count; $i++) { $sortKeys[$i] = $keys[$i].Replace('\', [char]1) }
        [Array]::Sort($sortKeys, $keys, [StringComparer]::OrdinalIgnoreCase)

        # ハッシュを計算するファイルの件数と容量を、進捗の表示のために先に数える
        if ($useHash) {
            foreach ($key in $keys) {
                $r = $null; $d = $null
                [void]$refItems.TryGetValue($key, [ref]$r)
                [void]$difItems.TryGetValue($key, [ref]$d)
                if ($r -and $d -and $r.Kind -eq 'ファイル' -and $d.Kind -eq 'ファイル' -and $r.Size -eq $d.Size -and -not $r.Online -and -not $d.Online) {
                    if (-not $refIsSnapshot) { $progress.Total += $r.Size; $progress.FileTotal++ }
                    if (-not $refIsSnapshot -or $r.Hash) { $progress.Total += $d.Size; $progress.FileTotal++ }
                }
            }
        }

        $counts = [ordered]@{ '追加' = 0; '削除' = 0; '変更' = 0; '種類が違う' = 0; '日時のみ' = 0; '同じ' = 0 }
        $results = New-Object 'System.Collections.Generic.List[object]'
        $linkKeys = 0
        foreach ($key in $keys) {
            $r = $null; $d = $null
            [void]$refItems.TryGetValue($key, [ref]$r)
            [void]$difItems.TryGetValue($key, [ref]$d)
            if (($r -and $r.Kind -eq 'リンク') -or ($d -and $d.Kind -eq 'リンク')) { $linkKeys++ }

            $kindText = ''
            $parts    = @()
            $note     = ''
            if (-not $d) {
                $result = '削除'; $kindText = $r.Kind
            } elseif (-not $r) {
                $result = '追加'; $kindText = $d.Kind
            } elseif ($r.Kind -ne $d.Kind) {
                $result = '種類が違う'; $kindText = '{0}→{1}' -f $r.Kind, $d.Kind
            } elseif ($r.Kind -ne 'ファイル') {
                $result = '同じ'; $kindText = $r.Kind
            } else {
                $kindText = 'ファイル'
                $sizeDiff = $r.Size -ne $d.Size
                $timeDiff = [math]::Abs(($r.Time - $d.Time).TotalSeconds) -gt $TimeTolerance
                $contentChecked = $false
                $contentDiff    = $false
                if ($useHash -and -not $sizeDiff) {
                    if ($r.Online -or $d.Online) {
                        $note = 'オンライン専用のため、ハッシュを計算していない'
                    } elseif ($refIsSnapshot -and -not $r.Hash) {
                        $note = 'スナップショットにハッシュがないため、サイズと更新日時で比べた'
                    } else {
                        if (-not $refIsSnapshot) { $r.Hash = Get-HashText $r '基準' }
                        if ($null -ne $r.Hash) { $d.Hash = Get-HashText $d '比較' }
                        # 開けなかったファイルは比べられないので、結果に入れない（skipped.csv に残してある）
                        if ($null -eq $r.Hash -or $null -eq $d.Hash) {
                            if ($null -eq $r.Hash) { $r.Hash = '' }
                            if ($null -eq $d.Hash) { $d.Hash = '' }
                            continue
                        }
                        $contentChecked = $true
                        $contentDiff    = $r.Hash -ne $d.Hash
                    }
                }
                if ($contentChecked) {
                    if ($contentDiff) {
                        $result = '変更'
                        $parts = @('内容')
                        if ($timeDiff) { $parts += '更新日時' }
                    } elseif ($timeDiff) {
                        $result = '日時のみ'; $parts = @('更新日時')
                    } else {
                        $result = '同じ'
                    }
                } else {
                    if ($sizeDiff) { $parts += 'サイズ' }
                    if ($timeDiff) { $parts += '更新日時' }
                    $result = if ($parts.Count -gt 0) { '変更' } else { '同じ' }
                }
            }
            if (-not $note) {
                if ($r -and $r.Kind -eq 'リンク') { $note = $r.Note }
                elseif ($d -and $d.Kind -eq 'リンク') { $note = $d.Note }
            }
            $counts[$result]++
            if ($result -eq '同じ' -and -not $IncludeSame) { continue }
            $results.Add([pscustomobject]@{ Result = $result; Kind = $kindText; Rel = $(if ($r) { $r.Rel } else { $d.Rel }); Diff = ($parts -join '・'); Ref = $r; Dif = $d; Note = $note })
        }
        $elapsed = $sw.Elapsed
        Write-Progress -Activity 'フォルダ差分' -Completed

        #endregion

        #region 画面と summary.txt の内容

        $diffCount = $counts['追加'] + $counts['削除'] + $counts['変更'] + $counts['種類が違う']
        $skipCount = 0
        foreach ($value in $skipKinds.Values) { $skipCount += $value }

        $lines.Add((Format-Invariant 'フォルダ差分 (Compare-Folder {0})' $ScriptVersion))
        if ($refIsSnapshot) {
            $savedText = $snapInfo.Saved.Replace('/', '-')
            $lines.Add(('基準: スナップショット {0}' -f $refPath))
            $lines.Add(('      （{0} を {1} に保存）' -f $snapInfo.Folder, $savedText))
        } else {
            $lines.Add(('基準: {0}' -f $refPath))
        }
        $lines.Add(('比較: {0}' -f $difPath))
        if ($configPath) { $lines.Add(('設定ファイル: {0}' -f $configPath)) }
        $lines.Add((Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（管理者: {2}）' $startTime $machineName $(if ($isAdmin) { 'はい' } else { 'いいえ' })))
        $lines.Add($(if ($useHash) { '比べ方: サイズと更新日時とハッシュ（SHA-256）' } else { '比べ方: サイズと更新日時' }))
        if ($hasExclude) { $lines.Add(('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', '))) }
        $lines.Add('')
        $lines.Add((Format-Invariant '基準  ファイル {0:N0} 件  フォルダ {1:N0} 件  合計 {2}' $refTotals.Files $refTotals.Dirs (Format-Size $refTotals.Size)))
        $lines.Add((Format-Invariant '比較  ファイル {0:N0} 件  フォルダ {1:N0} 件  合計 {2}' $difTotals.Files $difTotals.Dirs (Format-Size $difTotals.Size)))
        $lines.Add(('所要 {0}' -f (Format-Duration $elapsed)))
        $lines.Add('')
        $headIndex = $lines.Count

        # 違いの表。列の幅は、表に出す行に合わせる
        $tableRows = @($results | Where-Object { $_.Result -ne '同じ' })
        $widthResult = 4; $widthKind = 4; $widthDiff = 4
        foreach ($row in $tableRows) {
            $widthResult = [math]::Max($widthResult, (Get-DisplayWidth $row.Result))
            $widthKind   = [math]::Max($widthKind, (Get-DisplayWidth $row.Kind))
            $widthDiff   = [math]::Max($widthDiff, (Get-DisplayWidth $row.Diff))
        }
        $tableLines = New-Object 'System.Collections.Generic.List[object]'
        foreach ($row in $tableRows) {
            $refSize = if ($row.Ref -and $row.Ref.Kind -eq 'ファイル') { Format-Size $row.Ref.Size } else { '-' }
            $difSize = if ($row.Dif -and $row.Dif.Kind -eq 'ファイル') { Format-Size $row.Dif.Size } else { '-' }
            $text = ' ' + (Format-Column $row.Result $widthResult $false) + '  ' + (Format-Column $row.Kind $widthKind $false) + '  ' +
                (Format-Column $refSize 12 $true) + '  ' + (Format-Column $difSize 12 $true) + '  ' + (Format-Column $row.Diff $widthDiff $false) + '  ' + $row.Rel
            $color = switch ($row.Result) { '追加' { 'Green' } '削除' { 'Red' } '変更' { 'Yellow' } '種類が違う' { 'Magenta' } default { '' } }
            $tableLines.Add(@($text, $color))
        }
        $tableHeader = ' ' + (Format-Column '結果' $widthResult $false) + '  ' + (Format-Column '種類' $widthKind $false) + '  ' +
            (Format-Column '基準のサイズ' 12 $true) + '  ' + (Format-Column '比較のサイズ' 12 $true) + '  ' + (Format-Column '違い' $widthDiff $false) + '  相対パス'

        $resultParts = @()
        foreach ($name in '追加', '削除', '変更', '種類が違う', '日時のみ', '同じ') {
            if ($counts[$name] -gt 0 -or $name -eq '追加' -or $name -eq '削除' -or $name -eq '変更' -or $name -eq '同じ') {
                $resultParts += Format-Invariant '{0} {1:N0} 件' $name $counts[$name]
            }
        }
        if ($diffCount -gt 0) {
            $status = '違いあり'
            $exitCode = 2
        } elseif ($skipCount -gt 0) {
            $status = Format-Invariant '違いなし（ただし、比べられなかった場所が {0:N0} 件あります{1}）' $skipCount $(if ($runDir) { '。skipped.csv を確認してください' } else { '' })
            $exitCode = 1
        } else {
            $status = '違いなし'
        }
        $tail = New-Object 'System.Collections.Generic.List[string]'
        $tail.Add('')
        $tail.Add(('結果: {0}   {1}' -f $status, ($resultParts -join '  ')))
        foreach ($line in (Get-NoteLine $linkKeys)) { $tail.Add($line) }

        # summary.txt には、違いをすべて書く
        $lines.Add('■ 違い（相対パスの順）')
        if ($tableLines.Count -eq 0) {
            $lines.Add('    （違いはありません）')
        } else {
            $lines.Add($tableHeader)
            foreach ($entry in $tableLines) { $lines.Add($entry[0]) }
        }
        foreach ($line in $tail) { $lines.Add($line) }

        $rows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($row in $results) {
            $r = $row.Ref; $d = $row.Dif
            $rows.Add([object[]]@(
                $row.Result, $row.Kind, (Protect-ExcelText $row.Rel $false), $row.Diff,
                (Format-Bytes $r), (Format-Bytes $d),
                $(if ($r) { Format-Time $r.Time } else { '' }), $(if ($d) { Format-Time $d.Time } else { '' }),
                $(if ($r) { $r.Hash } else { '' }), $(if ($d) { $d.Hash } else { '' }), $row.Note))
        }

        #endregion
    }

    #region 保存

    if (-not $fatal -and $runDir) {
        Write-Progress -Activity 'フォルダ差分' -Status '結果を保存しています'
        try {
            [IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), [string[]]$lines, $Utf8Bom)
            if ($Snapshot) {
                Save-Csv (Join-Path $runDir 'snapshot.csv') $SnapshotColumns $rows
            } else {
                Save-Csv (Join-Path $runDir 'differences.csv') $DiffColumns $rows
            }
            Save-Csv (Join-Path $runDir 'skipped.csv') $SkipColumns $skipped
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
        }
        Write-Progress -Activity 'フォルダ差分' -Completed
    }

    #endregion

    $finished = $true
} finally {
    Write-Progress -Activity 'フォルダ差分' -Completed
    if (-not $finished -or $fatal -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete($runDir, $true) } catch { } }
        if (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}
if ($fatal) { Stop-WithError $fatal }
if ($Snapshot -and $saveError) { Stop-WithError ('スナップショットを保存できませんでした: {0}' -f $saveError) }

#region 画面表示

$saved = $runDir -and -not $saveError
if ($Snapshot) {
    foreach ($line in $lines) { Write-Host $line }
    Write-Host ('保存先: {0}' -f $runDir)
    Write-Host ''
    Write-Host '作業の後で、次のコマンドで比べられます:'
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { '.\Compare-Folder.ps1' }
    Write-Host ('powershell -ExecutionPolicy Bypass -File {0} -Reference {1} -Difference {2}' -f
        (Format-Argument $scriptPath), (Format-Argument (Join-Path $runDir 'snapshot.csv')), (Format-Argument $refPath))
    exit 0
}

for ($i = 0; $i -lt $headIndex; $i++) { Write-Host $lines[$i] }
Write-Host '■ 違い（相対パスの順）' -ForegroundColor Cyan
if ($tableLines.Count -eq 0) {
    Write-Host '    （違いはありません）'
} else {
    Write-Host $tableHeader
    # 保存できたときは、画面には先頭の分だけを出す。保存しなかったときは、ほかに見る場所がないためすべて出す
    $limit = if ($saved) { $ScreenRows } else { $tableLines.Count }
    $shown = 0
    foreach ($entry in $tableLines) {
        if ($shown -ge $limit) { break }
        if ($entry[1]) { Write-Host $entry[0] -ForegroundColor $entry[1] } else { Write-Host $entry[0] }
        $shown++
    }
    if ($tableLines.Count -gt $shown) {
        Write-Host (Format-Invariant '（ほか {0:N0} 件は differences.csv を参照してください）' ($tableLines.Count - $shown))
    }
}
foreach ($line in $tail) { Write-Host $line }

if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
} elseif ($saveWarning) {
    Write-Warning $saveWarning
}

#endregion

exit $exitCode
