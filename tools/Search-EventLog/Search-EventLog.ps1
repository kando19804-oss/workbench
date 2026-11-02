<#
.SYNOPSIS
    指定した期間のイベントログから、エラーなどのイベントを取り出して集計し、CSV に保存します。

.DESCRIPTION
    System や Application などのイベントログ（または .evtx ファイル）から、期間・レベル・イベント ID・ソースで
    絞り込んだイベントを取り出します。同じイベントが何回出ているかを集計して画面に表示し、
    events.csv と summary.csv（BOM 付き UTF-8）に保存します。

    イベントログを変更・削除することはありません。
    読めないログ（管理者でないときの Security など）は飛ばして、skipped.csv に残します。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Search-EventLog.ps1

.PARAMETER LogName
    読むイベントログの名前。カンマで区切って複数書けます。既定値は System,Application です。

.PARAMETER Path
    読む .evtx ファイル、または .evtx ファイルのあるフォルダ（サブフォルダも探します）。カンマで区切って複数書けます。
    指定すると、LogName は使いません。

.PARAMETER Since
    この日時以降のイベントを取り出します。"2026-10-31 09:00" のような日時か、2h・30m・3d のような今からの時間で書きます。
    既定値は 24h（直近 24 時間）です。

.PARAMETER Until
    この日時までのイベントを取り出します。書き方は Since と同じです。

.PARAMETER Level
    取り出すレベル。Critical、Error、Warning、Information、Verbose（または 重大、エラー、警告、情報、詳細）を
    カンマで区切って書きます。既定値は Critical,Error です。

.PARAMETER Id
    取り出すイベント ID。カンマで区切って複数書けます。

.PARAMETER ExcludeId
    除くイベント ID。カンマで区切って複数書けます。

.PARAMETER Source
    取り出すソース（イベント ビューアーの「ソース」）。カンマで区切って複数書けます。

.PARAMETER Pattern
    メッセージにこの文字を含むイベントだけを取り出します。| で区切ると、どれかを含むものを取り出します。
    書いた文字のまま探し、大文字と小文字は区別しません。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Search-EventLog.ps1

    直近 24 時間の、System と Application の重大・エラーを取り出します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Search-EventLog.ps1 -Since "2026-10-31 22:00" -Until "2026-11-01 02:00"

    指定した時間帯の重大・エラーを取り出します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Search-EventLog.ps1 -Path D:\collected -Since 3d

    D:\collected の下の .evtx ファイルから、直近 3 日の重大・エラーを取り出します。

.NOTES
    Search-EventLog 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$LogName = 'System,Application',

    [string]$Path,

    [string]$Since = '24h',

    [string]$Until,

    [string]$Level = 'Critical,Error',

    [string]$Id,

    [string]$ExcludeId,

    [string]$Source,

    [string]$Pattern,

    [string]$OutputFolder,

    [switch]$NoSave
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$ScreenGroups  = 30       # 画面に表示する集計の件数（保存するときだけ）
$ScreenEvents  = 20       # 画面に表示する新しい順の件数（保存するときだけ）
$ScreenMessage = 80       # 画面に表示するメッセージの長さ
$CellLimit     = 32000    # CSV の 1 つの値の長さ（Excel のセルの上限 32,767 文字より少し短く）
$ProgressMs    = 250      # 進捗表示の間隔（ミリ秒）
$LongPathLength = 250     # この長さ以上の .evtx のパスは、\\?\ を付けて読む
$ProviderPrefix = 'Microsoft-Windows-'
$NoMessage     = '（メッセージを取得できません）'
$EventColumns  = '日時', 'ログ', 'レベル', 'イベントID', 'ソース', 'タスクのカテゴリ', 'コンピューター', 'ユーザー', 'メッセージ'
$GroupColumns  = '件数', 'レベル', 'ログ', 'ソース', 'イベントID', '最初に発生', '最後に発生', 'メッセージ（1 行目）'
$SkipColumns   = '理由', '対象', '詳細'
$TimeFormat    = 'yyyy/MM/dd HH:mm:ss'
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$machineName   = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

# レベルの名前と、イベントログの値（「情報」には、Security の監査イベントなどのレベル 0 も含める）
$LevelNames = @{ 0 = '情報'; 1 = '重大'; 2 = 'エラー'; 3 = '警告'; 4 = '情報'; 5 = '詳細' }
$LevelKeys  = @{
    'critical' = @(1); '重大' = @(1); 'error' = @(2); 'エラー' = @(2); 'warning' = @(3); '警告' = @(3)
    'information' = @(4, 0); '情報' = @(4, 0); 'verbose' = @(5); '詳細' = @(5)
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
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -Path "D:\collected"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# -File で起動すると配列を渡せないため、カンマ区切りの文字列で受ける。各要素の前後の ' と空白も外す
function Split-List([string]$Text) {
    $list = New-Object 'System.Collections.Generic.List[string]'
    foreach ($part in ($Text -split ',')) {
        $item = (Remove-Quote $part).Trim("'").Trim()
        if ($item) { $list.Add($item) }
    }
    return , $list
}

# 日時か、2h・30m・3d のような今からの時間を読む。Relative は「（24 時間）」のように添える文字
function ConvertTo-PointInTime([string]$Text, [string]$Name) {
    $value = (Remove-Quote $Text).Trim("'")
    if ($value -match '^(\d{1,5})\s*([mhd])$') {
        $amount = [int]$Matches[1]
        switch ($Matches[2]) {
            'm' { return @{ Time = $startTime.AddMinutes(-$amount); Relative = '{0} 分' -f $amount } }
            'h' { return @{ Time = $startTime.AddHours(-$amount); Relative = '{0} 時間' -f $amount } }
            'd' { return @{ Time = $startTime.AddDays(-$amount); Relative = '{0} 日' -f $amount } }
        }
    }
    $formats = 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd', 'yyyy/MM/dd HH:mm:ss', 'yyyy/MM/dd HH:mm', 'yyyy/MM/dd',
               'yyyy-M-d H:mm:ss', 'yyyy-M-d H:mm', 'yyyy-M-d', 'yyyy/M/d H:mm:ss', 'yyyy/M/d H:mm', 'yyyy/M/d'
    $parsed = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($value, [string[]]$formats, $Invariant, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
        Stop-WithError ('{0} は "2026-10-31 09:00" のような日時か、2h・30m・3d のような今からの時間で指定してください: {1}' -f $Name, $value)
    }
    return @{ Time = $parsed; Relative = '' }
}

# 例外の連なり（InnerException）のどこかに、指定した種類の例外があるか
function Test-InnerException([Exception]$Exception, [type]$Type) {
    while ($Exception) {
        if ($Exception -is $Type) { return $true }
        $Exception = $Exception.InnerException
    }
    return $false
}

# Get-WinEvent の Path はワイルドカードとして解決されるため、[ ] * ? と ` を文字どおりに扱うようエスケープする。
# 260 文字以上のパスは \\?\ を付けないと読めない（? もエスケープするので、7 でも読める）
function ConvertTo-EvtxPath([string]$FullPath) {
    $text = if ($FullPath.Length -ge $LongPathLength) { ConvertTo-LongPath $FullPath } else { $FullPath }
    return $text.Replace('`', '``').Replace('[', '`[').Replace(']', '`]').Replace('*', '`*').Replace('?', '`?')
}

# イベント ビューアーの「ソース」と同じように、先頭の Microsoft-Windows- を省く
function ConvertTo-SourceName([string]$Provider) {
    if ($Provider.StartsWith($ProviderPrefix, [StringComparison]::OrdinalIgnoreCase)) { return $Provider.Substring($ProviderPrefix.Length) }
    return $Provider
}

# イベントログの側で絞るための式（XPath）に書く日時
function ConvertTo-QueryTime([datetime]$Time) {
    return $Time.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ', $Invariant)
}

# リパースポイントのうち、リンクと判定できたものの種類を返す（OneDrive のフォルダもリパースポイントなので、それ以外は $null）
function Get-LinkName($Entry) {
    if (($Entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) { return $null }
    # Windows が互換用に作ったリンク（All Users、Application Data など）は、5.1 では種類を判定できない。どれも隠し・システム属性付き
    $hiddenSystem = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System
    if ($Entry -is [IO.DirectoryInfo] -and ($Entry.Attributes -band $hiddenSystem) -eq $hiddenSystem) { return 'Windows の互換用リンク' }
    $linkType = $null
    try { $linkType = $Entry.LinkType } catch { }
    if ($linkType -eq 'Junction') { return 'ジャンクション' }
    if ($linkType -eq 'SymbolicLink') { return 'シンボリックリンク' }
    return $null
}

# 新しい順の表の見出しと行（行は、文字とレベルの組）
function Get-EventTable($List) {
    $width = @{ Level = 6; Log = 4; Source = 6 }
    foreach ($ev in $List) {
        $width.Level  = [math]::Max($width.Level, (Get-DisplayWidth $ev.Level))
        $width.Log    = [math]::Max($width.Log, (Get-DisplayWidth $ev.Log))
        $width.Source = [math]::Max($width.Source, (Get-DisplayWidth $ev.Source))
    }
    $header = ' ' + (Format-Column '日時' 19 $false) + '  ' + (Format-Column 'レベル' $width.Level $false) + '  ' + (Format-Column 'ログ' $width.Log $false) + '  ' +
        (Format-Column 'ソース' $width.Source $false) + '  ' + (Format-Column 'ID' 5 $false) + '  メッセージ（1 行目）'
    $rows = @(foreach ($ev in $List) {
        $text = ' ' + $ev.Time.ToString('yyyy-MM-dd HH:mm:ss', $Invariant) + '  ' + (Format-Column $ev.Level $width.Level $false) + '  ' + (Format-Column $ev.Log $width.Log $false) + '  ' +
            (Format-Column $ev.Source $width.Source $false) + '  ' + (Format-Column ([string]$ev.Id) 5 $false) + '  ' + (Limit-Text $ev.FirstLine $ScreenMessage)
        , @($text, $ev.LevelValue)
    })
    return @{ Header = $header; Lines = $rows }
}

# メッセージの 1 行目（空の行は飛ばす）
function Get-FirstLine([string]$Text) {
    foreach ($line in ($Text -split '\r?\n')) {
        $trimmed = $line.Trim()
        if ($trimmed) { return $trimmed }
    }
    return ''
}

function Limit-Text([string]$Text, [int]$Length) {
    if ($Text.Length -le $Length) { return $Text }
    return $Text.Substring(0, $Length) + '…'
}

# イベントに記録されたユーザー（SID）を名前に直す。同じ SID が何度も出るため、覚えておく
$accountCache = @{}
function Get-AccountName($Sid) {
    if ($null -eq $Sid) { return '' }
    $key = $Sid.Value
    if (-not $accountCache.ContainsKey($key)) {
        $name = $key
        try { $name = $Sid.Translate([Security.Principal.NTAccount]).Value } catch { }
        $accountCache[$key] = $name
    }
    return $accountCache[$key]
}

#endregion

#region 引数の確認

$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

$startTime = Get-Date
$sinceInfo = ConvertTo-PointInTime $Since '-Since'
$untilInfo = $null
if ($Until) { $untilInfo = ConvertTo-PointInTime $Until '-Until' }
$endTime = if ($untilInfo) { $untilInfo.Time } else { $startTime }
if ($sinceInfo.Time -ge $endTime) {
    $endLabel = if ($untilInfo) { '-Until の日時' } else { '今' }
    Stop-WithError (Format-Invariant '-Since の日時（{0:yyyy-MM-dd HH:mm}）が、{1}（{2:yyyy-MM-dd HH:mm}）と同じか、それより後になっています' $sinceInfo.Time $endLabel $endTime)
}

# レベル
$levelValues = New-Object 'System.Collections.Generic.List[int]'
$levelLabels = New-Object 'System.Collections.Generic.List[string]'
foreach ($item in (Split-List $Level)) {
    $key = $item.ToLowerInvariant()
    if (-not $LevelKeys.ContainsKey($key)) {
        Stop-WithError ('-Level は Critical、Error、Warning、Information、Verbose（または 重大、エラー、警告、情報、詳細）を、カンマで区切って指定してください: {0}' -f $item)
    }
    foreach ($value in $LevelKeys[$key]) { if (-not $levelValues.Contains($value)) { $levelValues.Add($value) } }
    $label = $LevelNames[$LevelKeys[$key][0]]
    if (-not $levelLabels.Contains($label)) { $levelLabels.Add($label) }
}
if ($levelValues.Count -eq 0) { Stop-WithError '-Level を指定してください' }

# イベント ID
function ConvertTo-IdList([string]$Text, [string]$Name) {
    $list = New-Object 'System.Collections.Generic.List[int]'
    foreach ($item in (Split-List $Text)) {
        $number = 0
        if (-not [int]::TryParse($item, [Globalization.NumberStyles]::None, $Invariant, [ref]$number) -or $number -gt 65535) {
            Stop-WithError ('{0} は 0～65535 の数値を、カンマで区切って指定してください: {1}' -f $Name, $item)
        }
        if (-not $list.Contains($number)) { $list.Add($number) }
    }
    return , $list
}
$idList      = ConvertTo-IdList $Id '-Id'
# 検索の式に書ける条件の数には上限がある（レベル 6 つと期間に、ID 22 個までは通り、23 個でエラーになった）。
# レベルと合わせても余裕を持って収まるよう、ID が多いときは取り出した後に絞る
$MaxQueryIds = 15
$idSet = New-Object 'System.Collections.Generic.HashSet[int]'
foreach ($value in $idList) { [void]$idSet.Add($value) }
$idInQuery = $idList.Count -gt 0 -and $idList.Count -le $MaxQueryIds
$excludeList = ConvertTo-IdList $ExcludeId '-ExcludeId'
$excludeSet  = New-Object 'System.Collections.Generic.HashSet[int]'
foreach ($value in $excludeList) { [void]$excludeSet.Add($value) }
$sourceList  = Split-List $Source
$sourceKeys  = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($value in $sourceList) { [void]$sourceKeys.Add((ConvertTo-SourceName $value)) }
$patternTerms = @(foreach ($term in ($(if ($Pattern) { $Pattern } else { '' }) -split '\|')) { if ($term.Trim()) { $term.Trim() } })

$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    $outText = Resolve-StartPath $outText
}

# 読む対象（ログ名か、.evtx ファイル）
$skipped   = New-Object 'System.Collections.Generic.List[object]'
$skipKinds = [ordered]@{ '権限なし' = 0; 'ログがない' = 0; 'ファイルを読めない' = 0; 'リンク' = 0; 'その他のエラー' = 0 }
function Add-Skip([string]$Reason, [string]$Target, [string]$Detail) {
    $skipped.Add([object[]]@($Reason, $Target, $Detail))
    $skipKinds[$Reason]++
}

$targets = New-Object 'System.Collections.Generic.List[object]'
$useFiles = [bool]$Path
if ($useFiles) {
    foreach ($item in (Split-List $Path)) {
        if ($item.Contains('"')) { Stop-WithQuoteError '-Path' $item }
        $full = ConvertTo-DisplayPath (Resolve-StartPath $item)
        try {
            $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($full)
            $full = [IO.Path]::GetFullPath($full)
        } catch { }
        if ([IO.File]::Exists((ConvertTo-LongPath $full))) {
            $targets.Add(@{ Label = $full; Filter = @{ Path = (ConvertTo-EvtxPath $full) } })
        } elseif ([IO.Directory]::Exists((ConvertTo-LongPath $full))) {
            # フォルダの下の .evtx を探す。リンクはたどらず、読めないフォルダは飛ばす
            # （ドライブのルートは、最後の \ を外すと \\?\D: になって読めない）
            $stack = New-Object 'System.Collections.Generic.Stack[string]'
            $top = $full.TrimEnd('\')
            if ($top -match '^[A-Za-z]:$') { $top += '\' }
            $stack.Push((ConvertTo-LongPath $top))
            while ($stack.Count -gt 0) {
                $dir = $stack.Pop()
                try { $entries = (New-Object IO.DirectoryInfo $dir).GetFileSystemInfos() } catch {
                    $ex = Get-InnerException $_.Exception
                    $reason = if ($ex -is [UnauthorizedAccessException]) { '権限なし' } else { 'その他のエラー' }
                    Add-Skip $reason (ConvertTo-DisplayPath $dir) (Get-ErrorText $ex)
                    continue
                }
                # ファイルは名前の順に加え、サブフォルダは名前の順にたどれるよう、逆の順に積む
                $subDirs = New-Object 'System.Collections.Generic.List[string]'
                foreach ($entry in ($entries | Sort-Object -Property Name)) {
                    # リンクはたどらない（OneDrive のフォルダもリパースポイントなので、リンクと判定できたものだけを飛ばす）
                    $linkName = Get-LinkName $entry
                    if ($linkName -and $entry -is [IO.DirectoryInfo]) {
                        Add-Skip 'リンク' (ConvertTo-DisplayPath $entry.FullName) $linkName
                        continue
                    }
                    if ($entry -is [IO.DirectoryInfo]) { $subDirs.Add($entry.FullName) }
                    elseif ($entry.Extension -ieq '.evtx') {
                        $file = ConvertTo-DisplayPath $entry.FullName
                        $targets.Add(@{ Label = $file; Filter = @{ Path = (ConvertTo-EvtxPath $file) } })
                    }
                }
                for ($i = $subDirs.Count - 1; $i -ge 0; $i--) { $stack.Push($subDirs[$i]) }
            }
        } else {
            Add-Skip 'ファイルを読めない' $full '見つかりません'
        }
    }
    if ($targets.Count -eq 0) { Stop-WithError ('-Path に、読める .evtx ファイルがありません: {0}' -f $Path) }
} else {
    foreach ($item in (Split-List $LogName)) { $targets.Add(@{ Label = $item; Filter = @{ LogName = $item } }) }
    if ($targets.Count -eq 0) { Stop-WithError '-LogName を指定してください' }
}

#endregion

#region 保存先の準備

# 読み終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir      = $null
$saveWarning = $null
if (-not $NoSave) {
    $baseDir = if ($outText) { $outText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
        $runName = Format-Invariant 'EventLog_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while ([IO.Directory]::Exists((ConvertTo-LongPath $runDir)) -or [IO.File]::Exists((ConvertTo-LongPath $runDir))) { $suffix++; $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix) }
        [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $runDir))
    } catch {
        $saveWarning = '保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, (Get-ErrorText (Get-InnerException $_.Exception))
        $runDir = $null
        Write-Warning $saveWarning
    }
}

#endregion

# 読んでいる途中や保存の途中で Ctrl+C されたときや、保存に失敗したときは、書きかけの結果を残さない
$finished   = $false
$unexpected = $false
$fatal      = $null
$saveError  = $null
$sw         = [Diagnostics.Stopwatch]::StartNew()
$events     = New-Object 'System.Collections.Generic.List[object]'
$noMessageCount = 0
$lines      = New-Object 'System.Collections.Generic.List[string]'
try {
    #region 読み取り

    # -FilterHashtable は使わない。権限のないログや壊れた .evtx でも「該当なし」と同じエラーを返して区別できず、
    # 名前に & を含む .evtx も読めないため
    $query = "*[System[({0}) and TimeCreated[@SystemTime>='{1}' and @SystemTime<='{2}']" -f (@($levelValues | ForEach-Object { "Level=$_" }) -join ' or '),
        (ConvertTo-QueryTime $sinceInfo.Time), (ConvertTo-QueryTime $endTime)
    if ($idInQuery) { $query += ' and ({0})' -f (@($idList | ForEach-Object { "EventID=$_" }) -join ' or ') }
    $query += ']]'

    $nextProgress = 0L
    foreach ($target in $targets) {
        $filter = $target.Filter
        $countBefore = $events.Count
        Write-Progress -Activity 'イベントログ抽出' -Status (Format-Invariant 'イベント {0:N0} 件' $events.Count) -CurrentOperation $target.Label
        try {
            # 1 件ずつ受け取って、読みながら絞り込む（すべてを先にメモリへ読み込まないため）。
            # -Oldest は、壊れた .evtx でも壊れた場所の手前まで読むため（新しい順だと、最後のほうが壊れていると 1 件も読めない）。
            # 分析とデバッグのログも、-Oldest でないと読めない。並べ替えは後で行う
            Get-WinEvent @filter -FilterXPath $query -Oldest -ErrorAction Stop | ForEach-Object {
                $record = $_
                if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                    Write-Progress -Activity 'イベントログ抽出' -Status (Format-Invariant 'イベント {0:N0} 件' $events.Count) -CurrentOperation $target.Label
                    $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
                }
                if ($excludeSet.Count -gt 0 -and $excludeSet.Contains([int]$record.Id)) { return }
                if ($idList.Count -gt 0 -and -not $idInQuery -and -not $idSet.Contains([int]$record.Id)) { return }
                # ソースは、Microsoft-Windows- を省いた名前でも探せるように、ここで絞る
                $sourceName = ConvertTo-SourceName ([string]$record.ProviderName)
                if ($sourceKeys.Count -gt 0 -and -not $sourceKeys.Contains($sourceName)) { return }
                # メッセージは、ソースの情報がないと取得できない（例外になることもある）
                $message = $null
                try { $message = $record.Message } catch { }
                if ($patternTerms.Count -gt 0) {
                    if (-not $message) { return }
                    $hit = $false
                    foreach ($term in $patternTerms) { if ($message.IndexOf($term, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $hit = $true; break } }
                    if (-not $hit) { return }
                }
                if (-not $message) { $message = $NoMessage; $script:noMessageCount++ }
                $task = ''
                try { $task = [string]$record.TaskDisplayName } catch { }
                $levelValue = [int]$record.Level
                $levelName  = if ($LevelNames.ContainsKey($levelValue)) { $LevelNames[$levelValue] } else { [string]$levelValue }
                $events.Add([pscustomobject]@{
                    Time = $record.TimeCreated; Log = [string]$record.LogName; LevelValue = $levelValue; Level = $levelName; Id = [int]$record.Id
                    Source = $sourceName; Task = $task; Machine = [string]$record.MachineName; User = (Get-AccountName $record.UserId)
                    Message = $message; FirstLine = (Get-FirstLine $message)
                })
            }
        } catch {
            $errorId = [string]$_.FullyQualifiedErrorId
            # 該当するイベントがなかっただけのときは、読めたものとして扱う
            if ($errorId -notlike 'NoMatchingEventsFound*') {
                $ex = Get-InnerException $_.Exception
                if ($errorId -like 'NoMatchingLogsFound*') {
                    Add-Skip 'ログがない' $target.Label 'この PC にはないログです'
                } elseif (Test-InnerException $_.Exception ([UnauthorizedAccessException])) {
                    Add-Skip '権限なし' $target.Label 'アクセスが拒否されました'
                } else {
                    $detail = Get-ErrorText $ex
                    if ($events.Count -gt $countBefore) { $detail = Format-Invariant '途中まで読んだ（イベント {0:N0} 件は結果に含む）: {1}' ($events.Count - $countBefore) $detail }
                    Add-Skip $(if ($useFiles) { 'ファイルを読めない' } else { 'その他のエラー' }) $target.Label $detail
                }
                $target['Failed'] = $true
            }
        }
    }
    $elapsed = $sw.Elapsed
    Write-Progress -Activity 'イベントログ抽出' -Completed

    # すべての対象を読めなかったときは、「該当なし」と区別して止める（途中まで読めたイベントがあれば、それを出す）
    $failedTargets = @($targets | Where-Object { $_.ContainsKey('Failed') }).Count
    if ($failedTargets -eq $targets.Count -and $events.Count -eq 0) {
        $fatal = '読めるログがありませんでした。' + (@($skipped | ForEach-Object { '{0}: {1}（{2}）' -f $_[1], $_[0], $_[2] }) -join ' / ')
        if (-not $isAdmin -and $skipKinds['権限なし'] -gt 0) { $fatal += "`r`n" + '※ 管理者として実行すると、Security などのログも読めます' }
    }

    #endregion

    if (-not $fatal) {
        #region 集計

        $sorted = @($events | Sort-Object -Property Time -Descending)
        $groups = @{}
        $levelCounts = @{}
        foreach ($ev in $sorted) {
            if ($levelCounts.ContainsKey($ev.Level)) { $levelCounts[$ev.Level]++ } else { $levelCounts[$ev.Level] = 1 }
            $key = '{0}|{1}|{2}|{3}' -f $ev.Log, $ev.Level, $ev.Source, $ev.Id
            $group = $groups[$key]
            if ($null -eq $group) {
                # 新しい順に見ているので、最初に見たものが最後に発生したもの
                $group = [pscustomobject]@{ Count = 0; Level = $ev.Level; LevelValue = $ev.LevelValue; Log = $ev.Log; Source = $ev.Source; Id = $ev.Id; First = $ev.Time; Last = $ev.Time; FirstLine = $ev.FirstLine }
                $groups[$key] = $group
            }
            $group.Count++
            $group.First = $ev.Time
        }
        $groupList = @($groups.Values | Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Last'; Descending = $true })

        #endregion

        #region 画面と summary.txt の内容

        $lines.Add((Format-Invariant 'イベントログ抽出 (Search-EventLog {0})' $ScriptVersion))
        if ($useFiles) {
            $more = if ($targets.Count -gt 1) { Format-Invariant ' ほか {0:N0} ファイル' ($targets.Count - 1) } else { '' }
            $lines.Add(('対象: {0}{1}' -f $targets[0].Label, $more))
        } else {
            $lines.Add(('対象: {0}（管理者: {1}）   ログ: {2}' -f $machineName, $(if ($isAdmin) { 'はい' } else { 'いいえ' }), (($targets | ForEach-Object { $_.Label }) -join ', ')))
        }
        $span = if ($sinceInfo.Relative -and -not $untilInfo) { '（{0}）' -f $sinceInfo.Relative } else { '' }
        $lines.Add((Format-Invariant '期間: {0:yyyy-MM-dd HH:mm} ～ {1:yyyy-MM-dd HH:mm}{2}   レベル: {3}' $sinceInfo.Time $endTime $span ($levelLabels -join ', ')))
        $conditions = @()
        if ($idList.Count -gt 0) { $conditions += 'ID ' + ($idList -join ', ') }
        if ($excludeList.Count -gt 0) { $conditions += '除く ID ' + ($excludeList -join ', ') }
        if ($sourceList.Count -gt 0) { $conditions += 'ソース ' + ($sourceList -join ', ') }
        if ($patternTerms.Count -gt 0) { $conditions += 'メッセージ ' + ($patternTerms -join ' | ') }
        if ($conditions.Count -gt 0) { $lines.Add(('条件: {0}' -f ($conditions -join '・'))) }
        $lines.Add('')
        $headIndex = $lines.Count

        # 集計の表
        $groupWidth = @{ Count = 4; Level = 6; Log = 4; Source = 6 }
        if ($groupList.Count -gt 0) { $groupWidth.Count = [math]::Max(4, $groupList[0].Count.ToString('N0', $Invariant).Length) }
        foreach ($group in $groupList) {
            $groupWidth.Level  = [math]::Max($groupWidth.Level, (Get-DisplayWidth $group.Level))
            $groupWidth.Log    = [math]::Max($groupWidth.Log, (Get-DisplayWidth $group.Log))
            $groupWidth.Source = [math]::Max($groupWidth.Source, (Get-DisplayWidth $group.Source))
        }
        $groupHeader = ' ' + (Format-Column '件数' $groupWidth.Count $true) + '  ' + (Format-Column 'レベル' $groupWidth.Level $false) + '  ' + (Format-Column 'ログ' $groupWidth.Log $false) + '  ' +
            (Format-Column 'ソース' $groupWidth.Source $false) + '  ' + (Format-Column 'ID' 5 $false) + '  ' + (Format-Column '最後に発生' 16 $false) + '  メッセージ（1 行目）'
        $groupLines = @(foreach ($group in $groupList) {
            $text = ' ' + (Format-Column ($group.Count.ToString('N0', $Invariant)) $groupWidth.Count $true) + '  ' + (Format-Column $group.Level $groupWidth.Level $false) + '  ' + (Format-Column $group.Log $groupWidth.Log $false) + '  ' +
                (Format-Column $group.Source $groupWidth.Source $false) + '  ' + (Format-Column ([string]$group.Id) 5 $false) + '  ' + $group.Last.ToString('yyyy-MM-dd HH:mm', $Invariant) + '  ' + (Limit-Text $group.FirstLine $ScreenMessage)
            , @($text, $group.LevelValue)
        })
        # 新しい順の表。全件は events.csv に書くので、画面の行は画面に出す分だけ作る（イベントが多いと時間がかかるため）。
        # 保存に失敗したときは、画面の後でもう一度すべての行を作る
        $eventTable = Get-EventTable $(if ($runDir) { @($sorted | Select-Object -First $ScreenEvents) } else { $sorted })

        $countParts = @()
        foreach ($value in 1, 2, 3, 4, 5) {
            $name = $LevelNames[$value]
            if ($levelCounts.ContainsKey($name)) { $countParts += Format-Invariant '{0} {1:N0}' $name $levelCounts[$name] }
        }
        $tail = New-Object 'System.Collections.Generic.List[string]'
        $tail.Add('')
        if ($sorted.Count -eq 0) {
            $tail.Add('該当するイベントはありません')
        } else {
            $tail.Add((Format-Invariant 'イベント {0:N0} 件（{1}）  種類 {2:N0}  所要 {3}' $sorted.Count ($countParts -join '・') $groupList.Count (Format-Duration $elapsed)))
        }
        $kinds = @(foreach ($name in $skipKinds.Keys) { if ($skipKinds[$name] -gt 0) { , @($name, $skipKinds[$name]) } })
        $skipTotal = 0
        foreach ($kind in $kinds) { $skipTotal += $kind[1] }
        if ($kinds.Count -eq 1) { $tail.Add((Format-Invariant 'スキップ {0:N0} 件（{1}）' $skipTotal $kinds[0][0])) }
        elseif ($kinds.Count -gt 1) { $tail.Add((Format-Invariant 'スキップ {0:N0} 件（{1}）' $skipTotal (($kinds | ForEach-Object { Format-Invariant '{0} {1:N0}' $_[0] $_[1] }) -join '・'))) }
        if ($noMessageCount -gt 0) { $tail.Add((Format-Invariant 'メッセージを取得できなかったイベント {0:N0} 件' $noMessageCount)) }
        if (-not $isAdmin -and $skipKinds['権限なし'] -gt 0) { $tail.Add('※ 管理者として実行すると、Security などのログも読めます') }

        # summary.txt には、集計をすべて書く
        $lines.Add(('■ 集計（件数の多い順）'))
        if ($groupLines.Count -eq 0) { $lines.Add('    （該当なし）') } else { $lines.Add($groupHeader); foreach ($entry in $groupLines) { $lines.Add($entry[0]) } }
        foreach ($line in $tail) { $lines.Add($line) }

        #endregion

        #region 保存

        if ($runDir) {
            Write-Progress -Activity 'イベントログ抽出' -Status '結果を保存しています'
            try {
                $lines.Add(('保存先: {0}' -f $runDir))
                [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
                $rows = New-Object 'System.Collections.Generic.List[object]'
                foreach ($ev in $sorted) {
                    $cell = if ($ev.Message.Length -gt $CellLimit) { $ev.Message.Substring(0, $CellLimit) } else { $ev.Message }
                    # ソースの名前は、管理者でなくても好きな名前で書き込めるので、ログの中の値はどれも数式の対策をする
                    $rows.Add([object[]]@($ev.Time.ToString($TimeFormat, $Invariant), (Protect-ExcelText $ev.Log), $ev.Level, $ev.Id, (Protect-ExcelText $ev.Source), (Protect-ExcelText $ev.Task),
                        (Protect-ExcelText $ev.Machine), (Protect-ExcelText $ev.User), (Protect-ExcelText $cell)))
                }
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'events.csv')) $EventColumns $rows
                $rows = New-Object 'System.Collections.Generic.List[object]'
                foreach ($group in $groupList) {
                    $rows.Add([object[]]@($group.Count.ToString('N0', $Invariant), $group.Level, (Protect-ExcelText $group.Log), (Protect-ExcelText $group.Source), $group.Id,
                        $group.First.ToString($TimeFormat, $Invariant), $group.Last.ToString($TimeFormat, $Invariant), (Protect-ExcelText $group.FirstLine)))
                }
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'summary.csv')) $GroupColumns $rows
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $skipped
            } catch {
                $saveError = Get-ErrorText (Get-InnerException $_.Exception)
            }
            Write-Progress -Activity 'イベントログ抽出' -Completed
        }

        #endregion
    }

    $finished = $true
} catch {
    # 想定外のエラーは、Ctrl+C の「中断しました」と区別して知らせる
    $unexpected = $true
    throw
} finally {
    Write-Progress -Activity 'イベントログ抽出' -Completed
    if (-not $finished -or $fatal -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}
if ($fatal) { Stop-WithError $fatal }

#region 画面表示

# 保存できたときは、画面には先頭の分だけを出す。保存しなかったときは、ほかに見る場所がないためすべて出す
$saved = $runDir -and -not $saveError
function Write-Table([string]$Title, [string]$Header, $Rows, [int]$Limit, [string]$MoreFile, [int]$Total = -1) {
    if ($Total -lt 0) { $Total = $Rows.Count }
    Write-Host $Title -ForegroundColor Cyan
    if ($Rows.Count -eq 0) { Write-Host '    （該当なし）'; return }
    Write-Host $Header
    $max = if ($saved) { $Limit } else { $Rows.Count }
    $shown = 0
    foreach ($entry in $Rows) {
        if ($shown -ge $max) { break }
        switch ($entry[1]) {
            1 { Write-Host $entry[0] -ForegroundColor Red }
            2 { Write-Host $entry[0] -ForegroundColor Yellow }
            default { Write-Host $entry[0] }
        }
        $shown++
    }
    if ($Total -gt $shown) { Write-Host (Format-Invariant '（ほか {0:N0} 件は {1} を参照してください）' ($Total - $shown) $MoreFile) }
}

for ($i = 0; $i -lt $headIndex; $i++) { Write-Host $lines[$i] }
if ($sorted.Count -gt 0) {
    $groupTitle = if ($saved) { '■ 集計（件数の多い順・上位 {0}）' -f $ScreenGroups } else { '■ 集計（件数の多い順）' }
    Write-Table $groupTitle $groupHeader $groupLines $ScreenGroups 'summary.csv'
    Write-Host ''
    $eventTitle = if ($saved) { '■ 新しい順（{0} 件まで）' -f $ScreenEvents } else { '■ 新しい順' }
    # 保存に失敗したときは、ほかに見る場所がないため、すべての行を作り直して出す
    if (-not $saved -and $eventTable.Lines.Count -lt $sorted.Count) { $eventTable = Get-EventTable $sorted }
    Write-Table $eventTitle $eventTable.Header $eventTable.Lines $ScreenEvents 'events.csv' $sorted.Count
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

if ($sorted.Count -gt 0) { exit 0 }
exit 2
