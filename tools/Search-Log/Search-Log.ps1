<#
.SYNOPSIS
    複数のログファイルから、指定した文字列を含む行を、前後の行と一緒に探し出します。

.DESCRIPTION
    指定したフォルダ（サブフォルダを含む）やファイルから、文字列を含む行を探し、前後の行と一緒に画面に表示します。
    結果はテキストと CSV（BOM 付き UTF-8）でも保存します。
    ログの文字コード（UTF-8・Shift_JIS・UTF-16）は、ファイルごとに自動で判定します。
    書き込み中のログも、書き込みを妨げずに読みます。ログを変更することはありません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Search-Log.ps1 -Path D:\logs -Pattern "timeout"

.PARAMETER Pattern
    探す文字列。"timeout|refused" のように | で区切ると、どれかを含む行を探します。
    既定では書いた文字をそのまま探します（. や ( もただの文字）。

.PARAMETER Path
    探すフォルダかファイル。フォルダはサブフォルダも含めて探します。複数指定するときは、カンマで区切ります。
    省略すると、現在のフォルダを探します。

.PARAMETER Regex
    Pattern と IgnorePattern を正規表現として扱います。このときは | で語に分けません。

.PARAMETER CaseSensitive
    大文字と小文字を区別します。

.PARAMETER IgnorePattern
    Pattern に一致しても、この文字列を含む行は結果から除きます。書き方は Pattern と同じです。

.PARAMETER Filter
    探すファイルの名前の条件（* と ?）。カンマで区切って複数書けます。既定値は "*.log,*.txt" です。

.PARAMETER Exclude
    探さないフォルダやファイル。フルパスに対して、ワイルドカード（*）で指定します。カンマで区切って複数書けます。

.PARAMETER Since
    最終更新がこの日時以降のファイルだけを探します。"2026-10-31 09:00" のような日時か、2h・30m・3d のような今からの時間で書きます。

.PARAMETER Context
    一致した行の前後に表示する行数（0～50）。3 なら前後 3 行ずつ、"2,10" なら前 2 行・後ろ 10 行です。既定値は 2 です。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.PARAMETER ConfigFile
    設定ファイルのパス。Path～NoSave の値を「項目 = 値」の形で書いておけます。
    コマンドラインで指定した値は、設定ファイルより優先します。書き方は README を参照してください。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Search-Log.ps1 -Path D:\logs -Pattern "timeout|refused"

    D:\logs の下の *.log と *.txt から、timeout か refused を含む行を探します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Search-Log.ps1 -Path D:\logs -Pattern "ERROR" -Since 2h -IgnorePattern "healthcheck"

    2 時間以内に更新されたログから、ERROR を含み healthcheck を含まない行を探します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Search-Log.ps1 -Path D:\logs -Pattern "Exception" -Context "2,15"

    Exception を含む行を、前 2 行・後ろ 15 行と一緒に表示します。

.NOTES
    Search-Log 1.0.0
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
    [string]$Pattern,

    [string[]]$Path = @('.'),

    [switch]$Regex,

    [switch]$CaseSensitive,

    [string]$IgnorePattern,

    [string[]]$Filter = @('*.log,*.txt'),

    [string[]]$Exclude = @(),

    [string]$Since,

    [string[]]$Context = @('2'),

    [string]$OutputFolder,

    [switch]$NoSave,

    [string]$ConfigFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$ScreenMatches = 200      # 画面に表示する一致の行数（保存するときだけ。超えた分はファイルを見てもらう）
$ScreenWidth   = 500      # 画面に表示する 1 行の長さ
$CellLimit     = 32000    # CSV の 1 つの値の長さ（Excel のセルの上限 32,767 文字より少し短く）
$SniffBytes    = 1MB      # 文字コードの判定に読む長さ
$BinaryBytes   = 8KB      # バイナリかどうかの判定に見る長さ
$ProgressMs    = 250      # 進捗表示の間隔（ミリ秒）
$OutputPrefix  = 'SearchLog_'
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$OnWindows     = [IO.Path]::DirectorySeparatorChar -eq '\'   # 開発用に Mac の pwsh でも動かせるようにするため
$SniffBuffer   = New-Object 'byte[]' $SniffBytes   # 文字コードの判定に使う。ファイルごとに作ると、小さなファイルが多いときに遅いため使い回す

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

# CSV の 1 つの値。件数が多くなるため ConvertTo-Csv を使わずに組み立てる。
# Excel は = + - @ で始まる値を数式として扱うため、その形の値は ="..." で囲む（式に書けない 255 文字を超える値は、先頭に空白を足す）
function ConvertTo-CsvField([string]$Text, [bool]$Guard) {
    if ($Guard -and $Text -match '^[=+\-@]') {
        if ($Text.Length -le 255) { $Text = '="' + $Text.Replace('"', '""') + '"' } else { $Text = ' ' + $Text }
    }
    return '"' + $Text.Replace('"', '""') + '"'
}

function Stop-WithError([string]$Message) {
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value) {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -Path "C:\logs"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# ワイルドカードは * と ? だけにして、それ以外の文字は文字どおりに扱う。5.1 の WildcardPattern.Escape は
# バッククォートをエスケープしないため、自分でエスケープしてから照合用のパターンを作る
function New-Wildcard([string]$Text, [bool]$AllowQuestion) {
    $escaped = $Text.Replace('`', '``').Replace('[', '`[').Replace(']', '`]')
    if (-not $AllowQuestion) { $escaped = $escaped.Replace('?', '`?') }
    return [Management.Automation.WildcardPattern]::new($escaped, [Management.Automation.WildcardOptions]'IgnoreCase, CultureInvariant')
}

# 文字コードを判定する。BOM があればそれに従い、先頭に NUL があればバイナリ（$null）、
# 先頭の SniffBytes が UTF-8 として正しければ UTF-8、そうでなければ Shift_JIS とみなす
function Get-LogEncoding([IO.Stream]$Stream) {
    $buffer = $SniffBuffer
    $read   = 0
    while ($read -lt $buffer.Length) {
        $n = $Stream.Read($buffer, $read, $buffer.Length - $read)
        if ($n -le 0) { break }
        $read += $n
    }
    if ($read -ge 3 -and $buffer[0] -eq 0xEF -and $buffer[1] -eq 0xBB -and $buffer[2] -eq 0xBF) { return , @((New-Object Text.UTF8Encoding $false), 'UTF-8') }
    if ($read -ge 2 -and $buffer[0] -eq 0xFF -and $buffer[1] -eq 0xFE) { return , @([Text.Encoding]::Unicode, 'UTF-16') }
    if ($read -ge 2 -and $buffer[0] -eq 0xFE -and $buffer[1] -eq 0xFF) { return , @([Text.Encoding]::BigEndianUnicode, 'UTF-16') }
    if ([Array]::IndexOf($buffer, [byte]0, 0, [math]::Min($read, $BinaryBytes)) -ge 0) { return $null }

    # 読んだ範囲の終わりで文字が途中で切れていることがあるため、ファイルの続きがあるときは、末尾の 3 バイトまでを外して確かめる
    $strict  = New-Object Text.UTF8Encoding $false, $true
    $maxCut  = if ($read -eq $buffer.Length) { 3 } else { 0 }
    for ($cut = 0; $cut -le $maxCut -and $cut -le $read; $cut++) {
        try {
            [void]$strict.GetCharCount($buffer, 0, $read - $cut)
            return , @((New-Object Text.UTF8Encoding $false), 'UTF-8')
        } catch { }
    }
    return , @([Text.Encoding]::GetEncoding(932), 'Shift_JIS')
}

# 探す文字列を、語ごとの正規表現にする。-Regex でなければ、書いた文字のまま探すようにエスケープする
function New-TermRegex([string]$Text, [string]$Name) {
    $source = if ($Regex) { $Text } else { [Text.RegularExpressions.Regex]::Escape($Text) }
    try {
        return New-Object Text.RegularExpressions.Regex ($source, $regexOptions)
    } catch {
        Stop-WithError ('{0} の正規表現の書き方が正しくありません: {1}（{2}）' -f $Name, $Text, (Get-ErrorText (Get-InnerException $_.Exception)))
    }
}

# コマンドラインの値を語に分ける。-Regex のときは分けない（正規表現の | を壊さないため）
function Split-Term([string]$Text) {
    if ($Regex) { return , @($Text) }
    return , @($Text -split '\|' | Where-Object { $_ -ne '' })
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの
$configItems   = 'Path', 'Pattern', 'Regex', 'CaseSensitive', 'IgnorePattern', 'Filter', 'Exclude', 'Since', 'Context', 'OutputFolder', 'NoSave'
$configMulti   = 'Path', 'Pattern', 'IgnorePattern', 'Filter', 'Exclude'   # 複数の行に分けて書ける項目
$configSwitch  = 'Regex', 'CaseSensitive', 'NoSave'
$configPath    = $null
$configApplied = @{}   # 設定ファイルの値を使った項目 → 行番号
$configTerms   = @{}   # 設定ファイルの Pattern と IgnorePattern（1 行が 1 語）

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
        # エクスプローラーで拡張子を表示していないと、メモ帳で search.ini.txt として保存してしまうことがある
        $message = '設定ファイルが見つかりません: {0}' -f $configPath
        if ([IO.File]::Exists((ConvertTo-LongPath ($configPath + '.txt')))) { $message += "`r`n" + ('拡張子が .txt になっていませんか: {0}.txt' -f $configPath) }
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
    $multiValues  = @{}
    $lineNumber   = 0
    foreach ($rawLine in $configLines) {
        $lineNumber++
        $lineText = $rawLine.Trim()
        # # と ; はパスや探す文字列に使える文字なので、行の先頭にあるときだけコメントとみなす
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

        # 複数の行に書ける項目のほかは、どちらが効いているか分からなくなるため 1 回だけ
        if ($configMulti -notcontains $itemName -and $configLineOf.ContainsKey($itemName)) {
            Stop-WithConfigError $lineNumber ('{0} が 2 回書かれています（{1} 行目と {2} 行目）' -f $itemName, $configLineOf[$itemName], $lineNumber)
        }
        if (-not $configLineOf.ContainsKey($itemName)) { $configLineOf[$itemName] = $lineNumber }

        if ($configSwitch -contains $itemName) {
            if ($itemValue -eq 'true') { $configValues[$itemName] = $true }
            elseif ($itemValue -eq 'false') { $configValues[$itemName] = $false }
            else { Stop-WithConfigError $lineNumber ('{0} は true か false で指定してください: {1}' -f $itemName, $itemValue) }
            continue
        }
        if ($itemName -eq 'Pattern' -or $itemName -eq 'IgnorePattern') {
            # 1 行が 1 語。| では分けない（| そのものを探せるようにするため）
            if (-not $multiValues.ContainsKey($itemName)) { $multiValues[$itemName] = New-Object 'System.Collections.Generic.List[string]' }
            $multiValues[$itemName].Add($itemValue)
            continue
        }
        if ($itemName -eq 'Path' -or $itemName -eq 'Exclude' -or $itemName -eq 'OutputFolder') {
            # パスの項目は、コマンドプロンプトで書くときと同じように %TEMP% などの環境変数を展開する
            if ($itemValue.Contains('"')) { Stop-WithConfigError $lineNumber ('{0} の値に引用符（"）が残っています: {1}' -f $itemName, $itemValue) }
            $itemValue = Expand-ConfigVariable $lineNumber $itemValue
            # 相対パスは、設定ファイルのあるフォルダを基準にする（タスクスケジューラでは現在のフォルダが System32 になるため）
            if ($itemName -ne 'Exclude' -and -not (Test-RootedPath $itemValue)) {
                $itemValue = $configDir.TrimEnd('\') + '\' + $itemValue
                try { $itemValue = [IO.Path]::GetFullPath($itemValue) } catch { }
            }
        }
        if ($configMulti -contains $itemName) {
            if (-not $multiValues.ContainsKey($itemName)) { $multiValues[$itemName] = New-Object 'System.Collections.Generic.List[string]' }
            $multiValues[$itemName].Add($itemValue)
        } else {
            $configValues[$itemName] = $itemValue
        }
    }
    foreach ($itemName in @($multiValues.Keys)) {
        if ($itemName -eq 'Pattern' -or $itemName -eq 'IgnorePattern') { $configTerms[$itemName] = $multiValues[$itemName].ToArray() }
        else { $configValues[$itemName] = $multiValues[$itemName].ToArray() }
    }

    # コマンドラインで指定した項目は、設定ファイルより優先する
    foreach ($itemName in @($configValues.Keys)) {
        if (-not $PSBoundParameters.ContainsKey($itemName)) {
            Set-Variable -Name $itemName -Value $configValues[$itemName]
            $configApplied[$itemName] = $configLineOf[$itemName]
        }
    }
    foreach ($itemName in @($configTerms.Keys)) {
        if ($PSBoundParameters.ContainsKey($itemName)) { $configTerms.Remove($itemName) } else { $configApplied[$itemName] = $configLineOf[$itemName] }
    }
}

#endregion

#region 引数の確認

# 前後の行数。"3" なら前後 3 行ずつ、"2,10" なら前 2 行・後ろ 10 行
$contextText = (Remove-Quote ($Context -join ',')).Trim("'")
if ($contextText -notmatch '^\s*(\d{1,2})\s*(?:,\s*(\d{1,2})\s*)?$') {
    Stop-WithError (('-Context は 0～50 の数値か、"2,10" のように前と後ろの行数で指定してください: {0}' -f $contextText) + (Get-ConfigNote 'Context'))
}
$beforeCount = [int]$Matches[1]
$afterCount  = if ($Matches.ContainsKey(2)) { [int]$Matches[2] } else { $beforeCount }
if ($beforeCount -gt 50 -or $afterCount -gt 50) {
    Stop-WithError (('-Context は 0～50 の数値か、"2,10" のように前と後ろの行数で指定してください: {0}' -f $contextText) + (Get-ConfigNote 'Context'))
}

# 最終更新の条件。日時か、2h・30m・3d のような今からの時間
$startTime = Get-Date
$sinceTime = $null
if ($Since) {
    $sinceText = (Remove-Quote $Since).Trim("'")
    if ($sinceText -match '^(\d{1,5})\s*([mhd])$') {
        $amount = [int]$Matches[1]
        switch ($Matches[2]) {
            'm' { $sinceTime = $startTime.AddMinutes(-$amount) }
            'h' { $sinceTime = $startTime.AddHours(-$amount) }
            'd' { $sinceTime = $startTime.AddDays(-$amount) }
        }
    } else {
        $formats = 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd', 'yyyy/MM/dd HH:mm:ss', 'yyyy/MM/dd HH:mm', 'yyyy/MM/dd',
                   'yyyy-M-d H:mm:ss', 'yyyy-M-d H:mm', 'yyyy-M-d', 'yyyy/M/d H:mm:ss', 'yyyy/M/d H:mm', 'yyyy/M/d'
        $parsed = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($sinceText, [string[]]$formats, $Invariant, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
            Stop-WithError (('-Since は "2026-10-31 09:00" のような日時か、2h・30m・3d のような今からの時間で指定してください: {0}' -f $sinceText) + (Get-ConfigNote 'Since'))
        }
        $sinceTime = $parsed
    }
}

# ファイル名の条件と除外。-File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける
$filterRules = @(
    foreach ($item in $Filter) {
        foreach ($part in ($item -split ',')) {
            $text = (Remove-Quote $part).Trim("'").Trim()
            if ($text) { [pscustomobject]@{ Text = $text; Matcher = (New-Wildcard $text $true) } }
        }
    }
)
if ($filterRules.Count -eq 0) { $filterRules = @([pscustomobject]@{ Text = '*'; Matcher = (New-Wildcard '*' $true) }) }

$excludeRules = @(
    foreach ($item in $Exclude) {
        foreach ($part in ($item -split ',')) {
            $text = (Remove-Quote $part).Trim("'")
            if ($text.Contains('"')) { Stop-WithQuoteError '-Exclude' $text }
            if ($text -notmatch '^[A-Za-z]:\\$') { $text = $text.TrimEnd('\') }
            if ($text) { [pscustomobject]@{ Text = $text; Matcher = (New-Wildcard $text $false) } }
        }
    }
)
$hasExclude = $excludeRules.Count -gt 0

# 探す場所。設定ファイルの Path は 1 行が 1 つ、コマンドラインの -Path はカンマで区切る
$pathTexts = @(
    if ($configApplied.ContainsKey('Path')) { $Path }
    else { foreach ($item in $Path) { foreach ($part in ($item -split ',')) { $part } } }
)
$roots        = New-Object 'System.Collections.Generic.List[object]'
$missingPaths = New-Object 'System.Collections.Generic.List[string]'
foreach ($rawPath in $pathTexts) {
    $text = (Remove-Quote $rawPath).Trim("'")
    if ($text -eq '') { continue }
    if ($text.Contains('"')) { Stop-WithQuoteError '-Path' $text }
    if ($text -match '^(\\\\\?\\)?[A-Za-z]:$') { $text += '\' }
    $text = Resolve-StartPath $text
    $full = $null
    try {
        $full = ConvertTo-DisplayPath $text
        $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($full)
        $full = [IO.Path]::GetFullPath($full)
    } catch { }
    if ($full -and [IO.Directory]::Exists((ConvertTo-LongPath $full))) {
        if ($full -notmatch '^[A-Za-z]:\\$') { $full = $full.TrimEnd('\') }
        $roots.Add([pscustomobject]@{ Path = $full; IsFile = $false })
    } elseif ($full -and [IO.File]::Exists((ConvertTo-LongPath $full))) {
        $roots.Add([pscustomobject]@{ Path = $full; IsFile = $true })
    } else {
        $missingPaths.Add($text)
    }
}
if ($roots.Count -eq 0 -and $missingPaths.Count -gt 0) {
    Stop-WithError (('指定したフォルダかファイルが見つかりません: {0}' -f ($missingPaths -join ', ')) + (Get-ConfigNote 'Path'))
}
if ($roots.Count -eq 0) { Stop-WithError '探すフォルダかファイルを -Path で指定してください' }
foreach ($missing in $missingPaths) { Write-Warning ('見つからないため、探しません: {0}' -f $missing) }

$outputText = $null
if ($OutputFolder) {
    $outputText = Remove-Quote $OutputFolder
    if ($outputText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outputText }
    $outputText = Resolve-StartPath $outputText
}

$regexOptions = [Text.RegularExpressions.RegexOptions]'CultureInvariant, Compiled'
if (-not $CaseSensitive) { $regexOptions = $regexOptions -bor [Text.RegularExpressions.RegexOptions]::IgnoreCase }

# 探す語。設定ファイルの Pattern は 1 行が 1 語、コマンドラインの -Pattern は | で分ける（-Regex のときは分けない）
$terms = @()
if ($configTerms.ContainsKey('Pattern')) { $terms = @($configTerms['Pattern'] | Where-Object { $_ -ne '' }) }
elseif ($Pattern) { $terms = Split-Term $Pattern }
if ($terms.Count -eq 0) { Stop-WithError '探す文字列を -Pattern で指定してください（例: -Pattern "timeout|refused"）' }
$termRegexes = @(foreach ($term in $terms) { New-TermRegex $term 'Pattern' })
# 正規表現の語が 2 つ以上あるときは、まとめると後方参照（\1 など）の番号がずれるため、1 つずつ試す
$checkEach = $Regex -and $terms.Count -gt 1
$matcherSource = (@(foreach ($r in $termRegexes) { '(?:' + $r.ToString() + ')' })) -join '|'
$matcher = $null
try { $matcher = New-Object Text.RegularExpressions.Regex ($matcherSource, $regexOptions) } catch { $checkEach = $true }

$ignoreTerms = @()
if ($configTerms.ContainsKey('IgnorePattern')) { $ignoreTerms = @($configTerms['IgnorePattern'] | Where-Object { $_ -ne '' }) }
elseif ($IgnorePattern) { $ignoreTerms = Split-Term $IgnorePattern }
$hasIgnore = $ignoreTerms.Count -gt 0
$ignorer   = $null
$ignoreRegexes = @()
$ignoreEach    = $false
if ($hasIgnore) {
    $ignoreRegexes = @(foreach ($term in $ignoreTerms) { New-TermRegex $term 'IgnorePattern' })
    $ignoreSource  = (@(foreach ($r in $ignoreRegexes) { '(?:' + $r.ToString() + ')' })) -join '|'
    $ignoreEach = $Regex -and $ignoreTerms.Count -gt 1
    try { $ignorer = New-Object Text.RegularExpressions.Regex ($ignoreSource, $regexOptions) } catch { $ignoreEach = $true }
}

#endregion

#region 保存先の準備

# 探し終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作り、結果を書きながら探す
$runDir      = $null
$saveWarning = $null
$saveError   = $null
$machineName = [Environment]::MachineName
if (-not $NoSave) {
    $baseDir = if ($outputText) { $outputText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        $runName = Format-Invariant '{0}{1}_{2:yyyyMMdd_HHmmss}' $OutputPrefix $machineName $startTime
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

#region 探す

$skipped       = New-Object 'System.Collections.Generic.List[object]'
$skipKinds     = [ordered]@{ '権限なし' = 0; 'バイナリ' = 0; '見つからない' = 0; 'その他のエラー' = 0 }
$linkCount     = 0
$excludedCount = 0
$sinceSkipped  = 0
$fileCount     = 0L    # 調べたファイル
$matchTotal    = 0L    # 一致した行
$matchFiles    = 0L    # 一致した行のあるファイル
$termLines     = New-Object 'long[]' $terms.Count
$termFiles     = New-Object 'long[]' $terms.Count
$fileSummary   = New-Object 'System.Collections.Generic.List[object]'   # 一致のあったファイルごとの件数

# 画面には、保存するときは ScreenMatches 行まで出す。超えたら、その一致の後ろの行までは出して止める（screenTail）
$screenLimit  = if ($runDir) { [long]$ScreenMatches } else { [long]::MaxValue }
$screenOn     = $true
$screenTail   = $false
$shownMatches = 0L
$needTitle    = $false   # 保存できなくなって画面に切り替えたあと、いまのファイルの見出しを出し直すか

$dirType = [IO.DirectoryInfo]
$rpFlag  = [IO.FileAttributes]::ReparsePoint
$hsFlag  = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System
$seen    = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
# Excel が数式・数値・日付に変える値（= + - @ で始まる、または数字と区切りの記号だけ）
$excelRisk = New-Object Text.RegularExpressions.Regex ('^(?:[=+\-@]|[\d\s.,/:+\-]+$)')
# 語が 1 つのときの「一致した語」の列（CSV の値として "" を重ねたもの）は、毎回同じなので先に作っておく
$singleTermCell = $terms[0]
if ($excelRisk.IsMatch($singleTermCell)) { $singleTermCell = '="' + $singleTermCell.Replace('"', '""') + '"' }
$singleTermCell = $singleTermCell.Replace('"', '""')

function Add-Skip([string]$Reason, [string]$ItemPath, [string]$Detail) {
    $skipped.Add([pscustomobject]@{ '理由' = $Reason; 'パス' = $ItemPath; '詳細' = $Detail })
    if ($skipKinds.Contains($Reason)) { $skipKinds[$Reason]++ }
}

function Add-SkipError([Exception]$Exception, [string]$ItemPath) {
    if ($Exception -is [UnauthorizedAccessException]) { Add-Skip '権限なし' $ItemPath 'アクセスが拒否されました' }
    else { Add-Skip 'その他のエラー' $ItemPath (Get-ErrorText $Exception) }
}

# リパースポイントのうち、リンクと判定できたものの種類を返す（OneDrive のファイルやフォルダもリパースポイントなので、それ以外は $null）
function Get-LinkName($Entry) {
    if (($Entry.Attributes -band $rpFlag) -ne $rpFlag) { return $null }
    $linkType = $null
    try { $linkType = $Entry.LinkType } catch { }
    if ($linkType -eq 'Junction') { return 'ジャンクション' }
    if ($linkType -eq 'SymbolicLink') { return 'シンボリックリンク' }
    # Windows が互換用に作ったリンク（All Users、Application Data など）は、5.1 では種類を判定できない。どれも隠し・システム属性付き
    if ($Entry -is $dirType -and ($Entry.Attributes -band $hsFlag) -eq $hsFlag) { return 'Windows の互換用リンク' }
    return $null
}

# 画面に 1 行出す。一致した行は、一致した部分を色で目立たせる（画面に出す行は限られるので、ここは関数でよい）
function Show-Entry([string]$Kind, [long]$Number, [string]$Text) {
    if ($Kind -eq 's') { Write-Host '  --' -ForegroundColor DarkGray; return }
    if ($Text.Length -gt $ScreenWidth) { $Text = $Text.Substring(0, $ScreenWidth) + '…' }
    if ($Kind -eq 'c') { Write-Host (' ' + $Number.ToString($Invariant).PadLeft(6) + '  ' + $Text); return }
    Write-Host ('>' + $Number.ToString($Invariant).PadLeft(6) + '  ') -NoNewline -ForegroundColor Yellow
    # 語ごとに一致した範囲を集め、重なりを除いて左から色を付ける
    $ranges = @(foreach ($r in $termRegexes) { foreach ($m in $r.Matches($Text)) { if ($m.Length -gt 0) { , @($m.Index, $m.Length) } } })
    $ranges = @($ranges | Sort-Object -Property @{ Expression = { $_[0] } }, @{ Expression = { $_[1] }; Descending = $true })
    $pos = 0
    foreach ($range in $ranges) {
        if ($range[0] -lt $pos) { continue }
        if ($range[0] -gt $pos) { Write-Host $Text.Substring($pos, $range[0] - $pos) -NoNewline }
        Write-Host $Text.Substring($range[0], $range[1]) -NoNewline -ForegroundColor Yellow
        $pos = $range[0] + $range[1]
    }
    Write-Host $Text.Substring($pos)
}

# 一致のあったファイルの件数をまとめる（ファイルの終わりで、または途中で読めなくなったときに 1 回だけ呼ぶ）
function Add-FileResult([string]$ItemPath, [string]$EncodingName, [long]$Hits, [bool[]]$TermHit) {
    if ($script:textWriter) { try { $script:textWriter.WriteLine('') } catch { Stop-Saving $_.Exception } }
    $script:matchTotal += $Hits
    $script:matchFiles++
    for ($t = 0; $t -lt $TermHit.Count; $t++) { if ($TermHit[$t]) { $termFiles[$t]++ } }
    $fileSummary.Add([pscustomobject]@{ Path = $ItemPath; Encoding = $EncodingName; Hits = $Hits })
}

$textWriter = $null
$csvWriter  = $null

# 保存の途中で書けなくなったら、書きかけの結果は残さない（最後にフォルダごと消す）。
# 画面の上限は外し、ここから先の一致はすべて画面に出す
function Stop-Saving([Exception]$Exception) {
    $script:saveError = Get-ErrorText (Get-InnerException $Exception)
    if ($script:textWriter) { try { $script:textWriter.Dispose() } catch { } }
    if ($script:csvWriter) { try { $script:csvWriter.Dispose() } catch { } }
    $script:textWriter  = $null
    $script:csvWriter   = $null
    $script:screenLimit = [long]::MaxValue
    if (-not $script:screenOn -or $script:screenTail) {
        $script:screenOn   = $true
        $script:screenTail = $false
        $script:needTitle  = $true
        Write-Host ''
        Write-Host '（保存できなくなったため、ここから先の一致はすべて画面に表示します）' -ForegroundColor Yellow
    }
}

# 条件は先に表示し、結果の件数は探し終わってから表示する（一致は見つけしだい表示する）
$header = New-Object 'System.Collections.Generic.List[string]'
$header.Add((Format-Invariant 'ログ横断検索 (Search-Log {0})' $ScriptVersion))
$targetText = ($roots | ForEach-Object { $_.Path }) -join ', '
if (@($roots | Where-Object { -not $_.IsFile }).Count -gt 0) { $targetText += '（{0}）' -f (($filterRules | ForEach-Object { $_.Text }) -join ', ') }
$header.Add((Format-Invariant '対象: {0}   {1:yyyy-MM-dd HH:mm:ss}   {2}' $targetText $startTime $machineName))
if ($configPath) { $header.Add(('設定ファイル: {0}' -f $configPath)) }
$modeText = $(if ($CaseSensitive) { '大文字と小文字を区別する' } else { '大文字と小文字を区別しない' })
if ($Regex) { $modeText = '正規表現・' + $modeText }
$header.Add(('検索: {0}（{1}）   前 {2} 行・後ろ {3} 行' -f ($terms -join ' | '), $modeText, $beforeCount, $afterCount))
if ($hasIgnore) { $header.Add(('除く: {0}' -f ($ignoreTerms -join ' | '))) }
if ($sinceTime) { $header.Add((Format-Invariant '条件: 最終更新が {0:yyyy-MM-dd HH:mm} 以降のファイル' $sinceTime)) }
if ($hasExclude) { $header.Add(('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', '))) }
foreach ($line in $header) { Write-Host $line }
foreach ($missing in $missingPaths) { Add-Skip '見つからない' $missing '指定したフォルダかファイルが見つかりません' }

$sw           = [Diagnostics.Stopwatch]::StartNew()
$nextProgress = 0L
$finished     = $false
try {
    if ($runDir) {
        try {
            $textWriter = New-Object IO.StreamWriter ((Join-Path $runDir 'matches.txt'), $false, $Utf8Bom)
            $csvWriter  = New-Object IO.StreamWriter ((Join-Path $runDir 'matches.csv'), $false, $Utf8Bom)
            $csvWriter.WriteLine('"ファイル","行番号","一致した語","一致した行","文字コード","最終更新日時"')
        } catch {
            Stop-Saving $_.Exception
        }
    }

    foreach ($root in $roots) {
        # 1. 探すファイルを集める。再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる
        $files = New-Object 'System.Collections.Generic.List[object]'
        if ($root.IsFile) {
            $files.Add((New-Object IO.FileInfo (ConvertTo-LongPath $root.Path)))
        } else {
            $stack = New-Object 'System.Collections.Generic.Stack[object]'
            $stack.Push((New-Object IO.DirectoryInfo (ConvertTo-LongPath $root.Path)))
            while ($stack.Count -gt 0) {
                $dir = $stack.Pop()
                try {
                    $entries = $dir.GetFileSystemInfos()
                } catch {
                    Add-SkipError (Get-InnerException $_.Exception) (ConvertTo-DisplayPath $dir.FullName)
                    continue
                }
                foreach ($entry in $entries) {
                    $entryPath = ConvertTo-DisplayPath $entry.FullName
                    $isDir     = $entry -is $dirType
                    # このツールの結果のフォルダは探さない（前回の結果に一致して、結果が増えていくのを防ぐため）
                    if ($isDir -and $entry.Name.StartsWith($OutputPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
                    if ($hasExclude) {
                        $matched = $null
                        foreach ($rule in $excludeRules) { if ($rule.Matcher.IsMatch($entryPath)) { $matched = $rule.Text; break } }
                        if ($matched) { $excludedCount++; Add-Skip '除外指定' $entryPath $matched; continue }
                    }
                    if (-not $isDir) {
                        $nameOk = $false
                        foreach ($rule in $filterRules) { if ($rule.Matcher.IsMatch($entry.Name)) { $nameOk = $true; break } }
                        if (-not $nameOk) { continue }
                    }
                    # フォルダもファイルも、リンクはたどらない（同じ内容を二重に数えたり、探す範囲の外を読んだりしないため）
                    $linkName = Get-LinkName $entry
                    if ($linkName) { $linkCount++; Add-Skip 'リンク' $entryPath $linkName; continue }
                    if ($isDir) { $stack.Push($entry) } else { $files.Add($entry) }
                }
            }
        }
        # パスの順に探す。Sort-Object はカルチャで比べ、5.1 と 7 で順が変わることがあるため、大文字にそろえたパスを序数で比べる。
        # [Array]::Sort に項目の配列を渡すと、並べ替わるのがコピーのほうになることがあるため、文字列だけを並べ、末尾の番号で元の項目を引く
        $keys = New-Object 'string[]' $files.Count
        for ($i = 0; $i -lt $files.Count; $i++) { $keys[$i] = $files[$i].FullName.ToUpperInvariant() + [char]0 + $i }
        [Array]::Sort($keys, [StringComparer]::Ordinal)
        $ordered = @(foreach ($key in $keys) { $files[[int]$key.Substring($key.LastIndexOf([char]0) + 1)] })

        # 2. 1 つずつ探す
        foreach ($file in $ordered) {
            $displayPath = ConvertTo-DisplayPath $file.FullName
            if (-not $seen.Add($displayPath)) { continue }   # -Path が重なっていても、同じファイルは 1 回だけ
            if ($sinceTime) {
                $modified = $null
                try { $file.Refresh(); $modified = $file.LastWriteTime } catch { }
                if ($modified -and $modified -lt $sinceTime) { $sinceSkipped++; continue }
            }
            if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                Write-Progress -Activity 'ログ横断検索' -Status ('ファイル {0:N0} 件 / 一致 {1:N0} 行' -f $fileCount, $matchTotal) -CurrentOperation $displayPath
                $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
            }

            $stream      = $null
            $reader      = $null
            $fileHits    = 0L
            $fileTermHit = New-Object 'bool[]' $terms.Count
            try {
                # 書き込み中のログも読めるよう、書き込みと削除を妨げない共有モードで開く
                $stream = New-Object IO.FileStream ($file.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read,
                    ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete), 65536)
                $encodingInfo = Get-LogEncoding $stream
                if (-not $encodingInfo) {
                    Add-Skip 'バイナリ' $displayPath '先頭に NUL を含む（バイナリか、BOM のない UTF-16）'
                    continue
                }
                $encodingName = $encodingInfo[1]
                [void]$stream.Seek(0, [IO.SeekOrigin]::Begin)
                $reader = New-Object IO.StreamReader ($stream, $encodingInfo[0], $true, 65536)
                $fileCount++

                # CSV の行のうち、ファイルごとに同じ部分は先に作っておく
                $titleLine = '■ ' + $displayPath + '（' + $encodingName + '）'
                $csvHead   = '"' + $displayPath.Replace('"', '""') + '","'
                $csvTail   = '","' + $encodingName + '","' + $file.LastWriteTime.ToString('yyyy/MM/dd HH:mm:ss', $Invariant) + '"'

                # 1 行ずつ読む。行の数が多いため、この中では関数を呼ばず、一致したらすぐに書き出す（画面に出す行だけ関数を呼ぶ）
                $prevNos     = New-Object 'System.Collections.Generic.Queue[long]'
                $prevTexts   = New-Object 'System.Collections.Generic.Queue[string]'
                $lineNo      = 0L
                $lastShown   = 0L
                $afterLeft   = 0
                while ($null -ne ($line = $reader.ReadLine())) {
                    $lineNo++
                    if (($lineNo -band 0xFFFF) -eq 0 -and $sw.ElapsedMilliseconds -ge $nextProgress) {
                        # 大きなファイルでも止まって見えないよう、読んでいる途中でも進捗を出す
                        Write-Progress -Activity 'ログ横断検索' -Status ('ファイル {0:N0} 件 / 一致 {1:N0} 行' -f $fileCount, ($matchTotal + $fileHits)) -CurrentOperation ('{0}（{1:N0} 行目）' -f $displayPath, $lineNo)
                        $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
                    }
                    if ($checkEach) {
                        $hit = $false
                        foreach ($r in $termRegexes) { if ($r.IsMatch($line)) { $hit = $true; break } }
                    } else {
                        $hit = $matcher.IsMatch($line)
                    }
                    if ($hit -and $hasIgnore) {
                        if ($ignoreEach) { foreach ($r in $ignoreRegexes) { if ($r.IsMatch($line)) { $hit = $false; break } } }
                        elseif ($ignorer.IsMatch($line)) { $hit = $false }
                    }

                    if ($hit) {
                        # 画面の上限に達したあとは、その一致の後ろの行まで出したら止める
                        if ($screenTail) { $screenOn = $false; $screenTail = $false }
                        if ($fileHits -eq 0) {
                            if ($textWriter) { try { $textWriter.WriteLine($titleLine) } catch { Stop-Saving $_.Exception } }
                            if ($screenOn) { Write-Host ''; Write-Host $titleLine -ForegroundColor Cyan; $needTitle = $false }
                        }
                        if ($needTitle -and $screenOn) { Write-Host ''; Write-Host $titleLine -ForegroundColor Cyan; $needTitle = $false }
                        $firstNo = if ($prevNos.Count -gt 0) { $prevNos.Peek() } else { $lineNo }
                        if ($lastShown -gt 0 -and $firstNo -gt $lastShown + 1) {
                            if ($textWriter) { try { $textWriter.WriteLine('  --') } catch { Stop-Saving $_.Exception } }
                            if ($screenOn) { Show-Entry 's' 0 '' }
                        }
                        while ($prevNos.Count -gt 0) {
                            $no   = $prevNos.Dequeue()
                            $text = $prevTexts.Dequeue()
                            if ($textWriter) { try { $textWriter.WriteLine(' ' + $no.ToString($Invariant).PadLeft(6) + '  ' + $text) } catch { Stop-Saving $_.Exception } }
                            if ($screenOn) { Show-Entry 'c' $no $text }
                        }

                        # 一致した語。語が 1 つなら数えるだけにして、一致の多いログでも遅くならないようにする
                        if ($terms.Count -eq 1) {
                            $termLines[0]++
                            $fileTermHit[0] = $true
                            $termCell = $singleTermCell
                        } else {
                            $termText = ''
                            for ($t = 0; $t -lt $termRegexes.Count; $t++) {
                                if ($termRegexes[$t].IsMatch($line)) {
                                    $termText = if ($termText) { $termText + ', ' + $terms[$t] } else { $terms[$t] }
                                    $termLines[$t]++
                                    $fileTermHit[$t] = $true
                                }
                            }
                            if ($excelRisk.IsMatch($termText)) { $termText = '="' + $termText.Replace('"', '""') + '"' }
                            $termCell = $termText.Replace('"', '""')
                        }
                        if ($textWriter) {
                            try {
                                $textWriter.WriteLine('>' + $lineNo.ToString($Invariant).PadLeft(6) + '  ' + $line)
                                # Excel のセルに入る長さにし、数式・数値・日付に変わる値は ="..." で囲む（式に書けない 255 文字を超える値は、先頭に空白を足す）
                                $cell = if ($line.Length -gt $CellLimit) { $line.Substring(0, $CellLimit) } else { $line }
                                if ($excelRisk.IsMatch($cell)) { if ($cell.Length -le 255) { $cell = '="' + $cell.Replace('"', '""') + '"' } else { $cell = ' ' + $cell } }
                                $csvWriter.WriteLine($csvHead + $lineNo.ToString($Invariant) + '","' + $termCell + '","' + $cell.Replace('"', '""') + $csvTail)
                            } catch {
                                Stop-Saving $_.Exception
                            }
                        }
                        if ($screenOn) {
                            if ($needTitle) { Write-Host ''; Write-Host $titleLine -ForegroundColor Cyan; $needTitle = $false }
                            Show-Entry 'm' $lineNo $line
                            $shownMatches++
                            if ($shownMatches -ge $screenLimit) { $screenTail = $true }
                        }
                        $fileHits++
                        $lastShown = $lineNo
                        $afterLeft = $afterCount
                    } elseif ($afterLeft -gt 0) {
                        if ($textWriter) { try { $textWriter.WriteLine(' ' + $lineNo.ToString($Invariant).PadLeft(6) + '  ' + $line) } catch { Stop-Saving $_.Exception } }
                        if ($screenOn) {
                            if ($needTitle) { Write-Host ''; Write-Host $titleLine -ForegroundColor Cyan; $needTitle = $false }
                            Show-Entry 'c' $lineNo $line
                        }
                        $lastShown = $lineNo
                        $afterLeft--
                    } elseif ($beforeCount -gt 0) {
                        $prevNos.Enqueue($lineNo)
                        $prevTexts.Enqueue($line)
                        if ($prevNos.Count -gt $beforeCount) { [void]$prevNos.Dequeue(); [void]$prevTexts.Dequeue() }
                    }
                }
                if ($fileHits -gt 0) { Add-FileResult $displayPath $encodingName $fileHits $fileTermHit }
            } catch {
                $failure = Get-InnerException $_.Exception
                if ($fileHits -gt 0) {
                    # 一致は見つけしだい書き出しているので、途中で読めなくなっても、それまでの一致は件数に入れる
                    Add-FileResult $displayPath $encodingName $fileHits $fileTermHit
                    Add-Skip 'その他のエラー' $displayPath (Format-Invariant '途中まで読んだ（一致 {0:N0} 行は結果に含む）: {1}' $fileHits (Get-ErrorText $failure))
                } else {
                    Add-SkipError $failure $displayPath
                }
            } finally {
                if ($reader) { $reader.Dispose() } elseif ($stream) { $stream.Dispose() }
            }
        }
    }
    Write-Progress -Activity 'ログ横断検索' -Completed
    $elapsed = $sw.Elapsed

    #region 結果のまとめ

    $footer = New-Object 'System.Collections.Generic.List[string]'
    if ($matchTotal -gt 0) {
        $footer.Add((Format-Invariant '結果: 一致 {0:N0} 行（{1:N0} ファイル）   調べたファイル {2:N0} 件   所要 {3}' $matchTotal $matchFiles $fileCount (Format-Duration $elapsed)))
    } else {
        $footer.Add((Format-Invariant '結果: 一致なし   調べたファイル {0:N0} 件   所要 {1}' $fileCount (Format-Duration $elapsed)))
    }
    $notes = @()
    $kinds = @(foreach ($k in $skipKinds.Keys) { if ($skipKinds[$k] -gt 0) { , @($k, $skipKinds[$k]) } })
    if ($kinds.Count -gt 0) {
        $total = 0
        foreach ($k in $kinds) { $total += $k[1] }
        if ($kinds.Count -eq 1) { $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $total $kinds[0][0] }
        else { $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $total (($kinds | ForEach-Object { Format-Invariant '{0} {1:N0}' $_[0] $_[1] }) -join '・') }
    }
    if ($linkCount -gt 0) { $notes += Format-Invariant 'たどらなかったリンク {0:N0} 件' $linkCount }
    if ($excludedCount -gt 0) { $notes += Format-Invariant '除外したフォルダ・ファイル {0:N0} 件' $excludedCount }
    if ($sinceTime) { $notes += Format-Invariant '最終更新で除いたファイル {0:N0} 件' $sinceSkipped }
    if ($notes.Count -gt 0) { $footer.Add($notes -join '  ') }

    # 語ごとの件数（語が 2 つ以上で、一致があったとき）
    $termTable = New-Object 'System.Collections.Generic.List[string]'
    if ($terms.Count -gt 1 -and $matchTotal -gt 0) {
        $width = 0
        foreach ($term in $terms) { $width = [math]::Max($width, (Get-DisplayWidth $term)) }
        $termTable.Add('■ 語ごとの件数')
        for ($t = 0; $t -lt $terms.Count; $t++) {
            $pad = ' ' * ($width - (Get-DisplayWidth $terms[$t]))
            $termTable.Add((Format-Invariant '  {0}{1}  {2:N0} 行（{3:N0} ファイル）' $terms[$t] $pad $termLines[$t] $termFiles[$t]))
        }
    }

    #endregion

    #region 保存

    if ($runDir -and -not $saveError) {
        try {
            $textWriter.Dispose(); $textWriter = $null
            $csvWriter.Dispose(); $csvWriter = $null

            $summary = New-Object 'System.Collections.Generic.List[string]'
            $summary.AddRange([string[]]$header)
            $summary.AddRange([string[]]$footer)
            if ($termTable.Count -gt 0) { $summary.Add(''); $summary.AddRange([string[]]$termTable) }
            $summary.Add('')
            $summary.Add('■ ファイルごとの件数')
            if ($fileSummary.Count -eq 0) { $summary.Add('  （一致なし）') }
            foreach ($f in $fileSummary) { $summary.Add((Format-Invariant '  {0,8:N0} 行  {1,-9}  {2}' $f.Hits $f.Encoding $f.Path)) }
            [IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), [string[]]$summary, $Utf8Bom)

            $lines = @('"理由","パス","詳細"') + @(foreach ($s in $skipped) {
                '"' + $s.'理由' + '","' + $s.'パス'.Replace('"', '""') + '","' + $s.'詳細'.Replace('"', '""') + '"'
            })
            [IO.File]::WriteAllLines((Join-Path $runDir 'skipped.csv'), [string[]]$lines, $Utf8Bom)
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
        }
    }
    if ($runDir -and $saveError) {
        # 書きかけの結果は残さない。消すのは、このツールが作った保存先のフォルダだけ
        if ($textWriter) { try { $textWriter.Dispose() } catch { } }
        if ($csvWriter) { try { $csvWriter.Dispose() } catch { } }
        $textWriter = $null
        $csvWriter  = $null
        try { [IO.Directory]::Delete($runDir, $true) } catch { }
    }

    #endregion

    #region 結果の表示

    $hiddenMatches = $matchTotal - $shownMatches
    if ($hiddenMatches -gt 0) {
        Write-Host ''
        if ($saveError) {
            Write-Host (Format-Invariant '（保存できなかったため、ほか {0:N0} 行は表示していません）' $hiddenMatches) -ForegroundColor Yellow
        } else {
            Write-Host (Format-Invariant '（ほか {0:N0} 行は matches.txt を参照してください）' $hiddenMatches) -ForegroundColor Yellow
        }
    }
    Write-Host ''
    foreach ($line in $footer) { Write-Host $line }
    if ($termTable.Count -gt 0) {
        Write-Host ''
        foreach ($line in $termTable) { if ($line.StartsWith('■')) { Write-Host $line -ForegroundColor Cyan } else { Write-Host $line } }
    }

    #endregion

    $finished = $true
} finally {
    Write-Progress -Activity 'ログ横断検索' -Completed
    if ($textWriter) { try { $textWriter.Dispose() } catch { } }
    if ($csvWriter) { try { $csvWriter.Dispose() } catch { } }
    if (-not $finished) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete($runDir, $true) } catch { } }
        Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow
    }
}

#endregion

if ($saveError) {
    Write-Host ''
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($runDir) {
    Write-Host ''
    Write-Host ('保存先: {0}' -f $runDir)
} elseif ($saveWarning) {
    Write-Host ''
    Write-Warning $saveWarning
}

if ($matchTotal -gt 0) { exit 0 }
exit 2
