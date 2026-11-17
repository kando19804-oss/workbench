<#
.SYNOPSIS
    このサーバーが、いつ止まり・いつ起動したか、誰が・何の理由で再起動したか、予期しない停止がなかったかを、System ログから一覧にします。

.DESCRIPTION
    System ログの停止の指示（1074）、OS の停止と起動（13・12）、予期しない停止（6008・41）、ブルースクリーン（1001）などを
    時刻の順に組み合わせて、停止 1 回を 1 行にした一覧を作ります。停止していた時間と、その前に稼働していた時間も出します。
    結果は画面に表示し、summary.txt・reboots.csv・events.csv・skipped.csv（BOM 付き UTF-8）に保存します。

    イベントログを変更・削除することはありません。管理者でなくても使えます。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Get-RebootHistory.ps1

.PARAMETER Since
    期間の始め。停止か起動のどちらかが期間に入るものを取り出します。
    "2026-11-10 09:00" のような日時か、2h・30m・3d のような今からの時間で書きます。省略すると、System ログに残っている、すべての記録です。

.PARAMETER Until
    期間の終わり。書き方は Since と同じです。省略すると、今までです。

.PARAMETER Path
    読む .evtx ファイル、または .evtx ファイルのあるフォルダ（サブフォルダも探します）。カンマで区切って複数書けます。
    System ログの .evtx を読みます。指定すると、この PC のイベントログは読みません。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。-OutputFolder と一緒には使えません。

.PARAMETER _UnknownArgument
    指定しません。どのパラメーターにも当たらない引数（パラメーターの名前の書き間違いなど）を受け取り、
    案内を出して止まるためのものです。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-RebootHistory.ps1

    System ログに残っている、すべての停止と起動を一覧にします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-RebootHistory.ps1 -Since 30d

    直近 30 日の停止と起動を一覧にします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-RebootHistory.ps1 -Path D:\collected\ap01

    ほかのサーバーから集めた .evtx ファイルを読みます。

.NOTES
    Get-RebootHistory 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$Since,

    [string]$Until,

    [string]$Path,

    [string]$OutputFolder,

    [switch]$NoSave,

    # どのパラメーターにも当たらない引数を受け取る（案内を出して止めるため。指定しない）
    [Parameter(ValueFromRemainingArguments = $true, DontShow = $true)]
    [string[]]$_UnknownArgument
)

# 制約付き言語モード（AppLocker や WDAC の環境）では、このツールが使う .NET の型を使えないので、最初に止める
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host '制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）'
    exit 1
}

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion   = '1.0.0'
$ToolName        = 'Get-RebootHistory'
$ToolTitle       = '再起動の履歴'
$ScreenRows      = 20       # 画面に表示する行数（保存するときだけ）
$NameWidth       = 20       # 画面の表のユーザーの幅（半角の文字数）。超える分は … にする
$ReasonWidth     = 35       # 画面の表のきっかけ・理由の幅。停止時間が 10 日を超え、ユーザーが 20 桁でも、表が 120 桁に収まる
$BootMergeSeconds = 600     # OS の起動（12）と、この秒数の内のイベントログの起動（6005）・起動の種類（27）は、同じ起動の記録
$ServiceRestartSeconds = 60 # イベントログのサービスだけの再起動は、6006 からこの秒数の内に 6005 が書かれる
$EventLogDelaySeconds = 60  # イベントログの起動（6005）は、OS の起動より遅れて書かれる。12 のない起動に予期しない停止の記録を割り当てるときは、この秒数だけ前を起動とみなす
$LastRecordScan  = 200      # 停止の前の最後の記録を探すとき、新しい順に見る記録の数（停止のときの記録は、次の起動のあとに書かれるため）
$ProgressMs      = 250      # 進捗表示の間隔（ミリ秒）
$LongPathLength  = 250      # この長さ以上の .evtx のパスは、\\?\ を付けて読む
$SystemLog       = 'System'
$RebootColumns   = '停止', '起動', '停止していた時間', 'その前の稼働時間', '種類', 'Windows Update', 'きっかけ', 'ユーザー', '理由', '理由のコード', 'コメント', '原因', '停止コード', 'ダンプ', 'あとで入力した理由', '停止の指示', 'プロセス', '起動の種類', 'OS のバージョン', '停止の前の最後の記録', '備考', 'コンピューター'
$EventColumns    = '日時', '操作', '詳細', 'ユーザー', 'プロセス', 'イベント ID', 'ソース', 'コンピューター'
$SkipColumns     = '理由', '対象', '詳細'
$TimeFormat      = 'yyyy/MM/dd HH:mm:ss'
$ScreenTime      = 'yyyy-MM-dd HH:mm'      # 画面の表は秒を省く（表を 120 桁に収めるため。CSV には秒まで書く）
$Invariant       = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom         = New-Object Text.UTF8Encoding $true
$machineName     = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

# 読む記録（ソースとイベント ID）。イベントログの側で絞り込む
$RecordQuery = "*[System[(Provider[@Name='User32'] and (EventID=1074 or EventID=1075 or EventID=1076))" +
    " or (Provider[@Name='EventLog'] and (EventID=6005 or EventID=6006 or EventID=6008))" +
    " or (Provider[@Name='Microsoft-Windows-Kernel-General'] and (EventID=12 or EventID=13))" +
    " or (Provider[@Name='Microsoft-Windows-Kernel-Power'] and (EventID=41 or EventID=109))" +
    " or (Provider[@Name='Microsoft-Windows-Kernel-Boot'] and EventID=27)" +
    " or (Provider[@Name='Microsoft-Windows-WER-SystemErrorReporting'] and EventID=1001)" +
    " or (Provider[@Name='Microsoft-Windows-Eventlog'] and EventID=104)]]"

# events.csv の操作の名前
$Operations = @{
    1074 = '停止の指示'; 1075 = '停止の取り消し'; 1076 = '理由の入力'; 109 = '停止の開始'; 13 = 'OS の停止'; 6006 = 'イベントログの停止'
    12 = 'OS の起動'; 27 = '起動の種類'; 6005 = 'イベントログの起動'; 6008 = '予期しない停止'; 41 = '正常に停止しないまま起動'; 1001 = 'ブルースクリーン'; 104 = 'ログの消去'
}

# 停止を指示したプロセス（小文字のファイル名）→ きっかけ
$TriggerLabels = @{
    'startmenuexperiencehost.exe' = 'スタート メニュー'
    'shutdown.exe'                = 'shutdown コマンド'
    'trustedinstaller.exe'        = 'Windows Update'
    'mousocoreworker.exe'         = 'Windows Update'
    'usoclient.exe'               = 'Windows Update'
    'wuauclt.exe'                 = 'Windows Update'
    'wlms.exe'                    = '評価版の期限切れ'                 # 評価版のライセンスの期限が切れると、1 時間ごとにシャットダウンする
    'wmiprvse.exe'                = 'WMI（Restart-Computer など）'    # Restart-Computer、Rename-Computer -Restart など。理由は「この理由のタイトルが見つかりません」になる
}

# 理由のコードの「オペレーティング システム」の、更新プログラムに関わる小分類（アップグレード・サービス パック・修正プログラム・セキュリティ修正）
$UpdateMinorCodes = @(0x03, 0x10, 0x11, 0x12)

# よく出る停止コードの名前。キーは 16 進数 8 桁の文字（数値のキーは、値の型が違うと当たらないため）
$BugcheckNames = @{
    '0000000A' = 'IRQL_NOT_LESS_OR_EQUAL'
    '00000019' = 'BAD_POOL_HEADER'
    '0000001A' = 'MEMORY_MANAGEMENT'
    '0000001E' = 'KMODE_EXCEPTION_NOT_HANDLED'
    '00000024' = 'NTFS_FILE_SYSTEM'
    '0000003B' = 'SYSTEM_SERVICE_EXCEPTION'
    '00000050' = 'PAGE_FAULT_IN_NONPAGED_AREA'
    '0000007A' = 'KERNEL_DATA_INPAGE_ERROR'
    '0000007B' = 'INACCESSIBLE_BOOT_DEVICE'
    '0000007E' = 'SYSTEM_THREAD_EXCEPTION_NOT_HANDLED'
    '0000007F' = 'UNEXPECTED_KERNEL_MODE_TRAP'
    '0000009F' = 'DRIVER_POWER_STATE_FAILURE'
    '000000BE' = 'ATTEMPTED_WRITE_TO_READONLY_MEMORY'
    '000000C2' = 'BAD_POOL_CALLER'
    '000000C5' = 'DRIVER_CORRUPTED_EXPOOL'
    '000000D1' = 'DRIVER_IRQL_NOT_LESS_OR_EQUAL'
    '000000E2' = 'MANUALLY_INITIATED_CRASH'
    '000000EA' = 'THREAD_STUCK_IN_DEVICE_DRIVER'
    '000000EF' = 'CRITICAL_PROCESS_DIED'
    '000000F4' = 'CRITICAL_OBJECT_TERMINATION'
    '00000101' = 'CLOCK_WATCHDOG_TIMEOUT'
    '00000109' = 'CRITICAL_STRUCTURE_CORRUPTION'
    '00000116' = 'VIDEO_TDR_FAILURE'
    '00000124' = 'WHEA_UNCORRECTABLE_ERROR'
    '0000012B' = 'FAULTY_HARDWARE_CORRUPTED_PAGE'
    '00000133' = 'DPC_WATCHDOG_VIOLATION'
    '00000139' = 'KERNEL_SECURITY_CHECK_FAILURE'
    '00000154' = 'UNEXPECTED_STORE_EXCEPTION'
    'DEADDEAD' = 'MANUALLY_INITIATED_CRASH1'
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

# 停止していた時間や稼働時間（0:06:40、1 日を超えると 1日 2:03:00）
function Format-Span([TimeSpan]$Span) {
    if ($Span.Ticks -lt 0) { $Span = [TimeSpan]::Zero }
    if ($Span.TotalDays -ge 1) { return Format-Invariant '{0}日 {1}:{2:00}:{3:00}' ([math]::Floor($Span.TotalDays)) $Span.Hours $Span.Minutes $Span.Seconds }
    return Format-Invariant '{0}:{1:00}:{2:00}' $Span.Hours $Span.Minutes $Span.Seconds
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

# 例外の連なり（InnerException）のどこかに、指定した種類の例外があるか
function Test-InnerException([Exception]$Exception, [type]$Type) {
    while ($Exception) {
        if ($Exception -is $Type) { return $true }
        $Exception = $Exception.InnerException
    }
    return $false
}

# CSV の 1 行
function ConvertTo-CsvLine([object[]]$Fields) {
    $parts = [string[]]::new($Fields.Count)
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

# 見出しと行（文字列の配列）から、列をそろえた表の行を作る
function Format-TextTable([string[]]$Head, [bool[]]$Right, $Rows) {
    $all = @(, $Head) + @($Rows)
    $columns = $Head.Count
    $widths = [int[]]::new($columns)
    $cellWidths = [object[]]::new($all.Count)
    for ($r = 0; $r -lt $all.Count; $r++) {
        $rowWidths = [int[]]::new($columns)
        for ($i = 0; $i -lt $columns; $i++) {
            $text = [string]$all[$r][$i]
            $w = if ($wideChar.IsMatch($text)) { Get-DisplayWidth $text } else { $text.Length }
            $rowWidths[$i] = $w
            if ($w -gt $widths[$i]) { $widths[$i] = $w }
        }
        $cellWidths[$r] = $rowWidths
    }
    $lines = [System.Collections.Generic.List[string]]::new()
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

function Stop-WithError([string]$Message) {
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value, [string]$Example = 'D:\collected') {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        ('値を引用符で囲むときは、最後に \ を付けないでください（例: {0} "{1}"）。' -f $Name, $Example) + '最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
}

# 引数の値に引用符（"）が残っていないかを、ほかの確認より先にまとめて確かめ、そのあと、どのパラメーターにも当たらない引数
# （-_UnknownArgument に集まる）があれば止める。コマンドプロンプトで "C:\data\" のように最後を \" にすると、引用符が閉じずに
# 後ろの引数までその値に入り、空白を含む値の残りが余る（-Path "C:\data\" -Since "2026-11-12 09:00" なら 09:00）。
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
Assert-CommandArgument $PSBoundParameters ([ordered]@{ Since = '2026-11-10 09:00'; Until = '2026-11-10 18:00'; Path = 'D:\collected'; OutputFolder = 'D:\collected' })

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
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($part in ($Text -split ',')) {
        $item = (Remove-Quote $part).Trim("'").Trim()
        if ($item) { $list.Add($item) }
    }
    return , $list
}

# 日時か、2h・30m・3d のような今からの時間を読む。Relative は「（30 日）」のように添える文字
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
        Stop-WithError ('{0} の日時を読めません: {1}（"2026-11-10 09:00" のような日時か、2h・30m・3d のような今からの時間で指定してください）' -f $Name, $value)
    }
    return @{ Time = $parsed; Relative = '' }
}

# イベントログ（または .evtx）を読む。Get-WinEvent は使わない（使わないメッセージを 1 件ごとに組み立てるため遅い。
# Path をワイルドカードとして読む問題もない）。Source の Name はログの名前かファイルのパス
function New-EventReader($Source, [string]$Query, [bool]$Reverse = $false) {
    $pathType = if ($Source.IsFile) { [Diagnostics.Eventing.Reader.PathType]::FilePath } else { [Diagnostics.Eventing.Reader.PathType]::LogName }
    $eventQuery = New-Object Diagnostics.Eventing.Reader.EventLogQuery($Source.Name, $pathType, $Query)
    $eventQuery.ReverseDirection = $Reverse
    $reader = New-Object Diagnostics.Eventing.Reader.EventLogReader($eventQuery)
    # 既定では 64 件ずつまとめて取り出すため、途中が壊れた .evtx では、壊れた場所の手前の記録までまとめて失う。1 件ずつ取り出す
    # （条件に合う記録が少ないので、遅くはならない）
    $reader.BatchSize = 1
    return $reader
}

# 条件に合う最初の記録（$Reverse なら最後の記録）の、日時・ログの名前・コンピューター。ないときは $null
function Read-FirstEvent($Source, [string]$Query, [bool]$Reverse = $false) {
    $reader = New-EventReader $Source $Query $Reverse
    try {
        $record = $reader.ReadEvent()
        if ($null -eq $record) { return $null }
        try { return @{ Time = $record.TimeCreated; Log = [string]$record.LogName; Machine = [string]$record.MachineName } } finally { $record.Dispose() }
    } finally {
        $reader.Dispose()
    }
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

# コンピューターの名前（server01.corp.example.local）を、組にするためのキー（SERVER01）にする
function Get-ComputerKey([string]$Name) {
    return $Name.Split('.')[0].ToUpperInvariant()
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

# 記録の数値。数値でも "0x80020003" のような文字でも届くので、どちらも読む
function ConvertTo-UInt64Value($Value) {
    if ($null -eq $Value) { return [uint64]0 }
    if ($Value -is [bool]) { if ($Value) { return [uint64]1 }; return [uint64]0 }
    if ($Value -is [string]) {
        $text = $Value.Trim()
        if ($text -match '^0[xX]([0-9A-Fa-f]{1,16})$') { return [Convert]::ToUInt64($Matches[1], 16) }
        $number = [uint64]0
        if ([uint64]::TryParse($text, [ref]$number)) { return $number }
        return [uint64]0
    }
    # 負の Int32 は、ビットのまま符号なしにする
    if ($Value -is [int]) { return [uint64][BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$Value), 0) }
    if ($Value -is [long]) { return [BitConverter]::ToUInt64([BitConverter]::GetBytes([long]$Value), 0) }
    return [uint64]$Value
}

# 記録の日時の値。ローカルの時刻にそろえる（記録の値は、ふつうはローカルの時刻で届く）
function ConvertTo-LocalTime($Value) {
    if ($Value -isnot [datetime]) { return $null }
    $time = [datetime]$Value
    if ($time.Kind -eq [DateTimeKind]::Utc) { return $time.ToLocalTime() }
    return $time
}

# バイト列の中の SYSTEMTIME（16 バイト）を日時にする。正しくない値なら $null
function ConvertFrom-SystemTime([byte[]]$Bytes, [int]$Offset, [DateTimeKind]$Kind) {
    if ($null -eq $Bytes -or $Bytes.Length -lt $Offset + 16) { return $null }
    $year   = [BitConverter]::ToUInt16($Bytes, $Offset)
    $month  = [BitConverter]::ToUInt16($Bytes, $Offset + 2)
    $day    = [BitConverter]::ToUInt16($Bytes, $Offset + 6)
    $hour   = [BitConverter]::ToUInt16($Bytes, $Offset + 8)
    $minute = [BitConverter]::ToUInt16($Bytes, $Offset + 10)
    $second = [BitConverter]::ToUInt16($Bytes, $Offset + 12)
    $msec   = [BitConverter]::ToUInt16($Bytes, $Offset + 14)
    if ($year -lt 1601) { return $null }
    try {
        $time = [datetime]::new([int]$year, [int]$month, [int]$day, [int]$hour, [int]$minute, [int]$second, [int]$msec, $Kind)
    } catch {
        return $null
    }
    if ($Kind -eq [DateTimeKind]::Utc) { return $time.ToLocalTime() }
    return $time
}

# 停止の指示（1074）のプロセス。「C:\Windows\system32\shutdown.exe (SERVER01)」の最後のコンピューターの名前を外す
function Get-ProcessPath([string]$Text) {
    return (ConvertTo-CellText $Text) -replace '\s+\([^()]*\)$', ''
}

# プロセスのファイル名（小文字ではない、記録のまま）
function Get-ProcessFileName([string]$ProcessPath) {
    $index = $ProcessPath.LastIndexOf('\')
    if ($index -ge 0) { return $ProcessPath.Substring($index + 1) }
    return $ProcessPath
}

# きっかけ（よく出るプロセスは分かる名前に、ほかはファイル名）
function Get-TriggerLabel([string]$ProcessPath) {
    $file = Get-ProcessFileName $ProcessPath
    $key = $file.ToLowerInvariant()
    if ($TriggerLabels.ContainsKey($key)) { return $TriggerLabels[$key] }
    return $file
}

# 停止の指示（1074）の種類の文字（OS の言語で書かれる。日本語と英語を読む）→ Restart / Shutdown。分からなければ $null
function Get-ActionFromText([string]$Text) {
    $t = (ConvertTo-CellText $Text).ToLowerInvariant()
    if ($t -eq '再起動' -or $t -eq 'restart') { return 'Restart' }
    if ($t -eq '電源を切る' -or $t -eq 'シャットダウン' -or $t -eq 'power off' -or $t -eq 'shutdown') { return 'Shutdown' }
    return $null
}

# 停止の開始（109）の ShutdownActionType → Restart / Shutdown（4 はシャットダウン、5 は再起動、6 は電源を切る）
function Get-ActionFromType([uint64]$Type) {
    if ($Type -eq 5) { return 'Restart' }
    if ($Type -eq 4 -or $Type -eq 6) { return 'Shutdown' }
    return $null
}

function Get-ActionLabel($Action) {
    if ($Action -eq 'Restart') { return '再起動' }
    if ($Action -eq 'Shutdown') { return 'シャットダウン' }
    return ''
}

# 停止コードの表示（ブルースクリーン 0x0000003B（SYSTEM_SERVICE_EXCEPTION））
function Format-Bugcheck([uint64]$Code) {
    $hex = ([uint32]($Code -band 0xFFFFFFFFL)).ToString('X8', $Invariant)
    $text = '0x' + $hex
    if ($BugcheckNames.ContainsKey($hex)) { $text += '（{0}）' -f $BugcheckNames[$hex] }
    return $text
}

function Format-Time($Time) {
    if ($null -eq $Time) { return '' }
    return ([datetime]$Time).ToString($TimeFormat, $Invariant)
}

function Format-ScreenTime($Time) {
    if ($null -eq $Time) { return '' }
    return ([datetime]$Time).ToString($ScreenTime, $Invariant)
}

# 画面の起動の時刻。停止と同じ日なら時刻だけ、違う日なら月日と時刻（表を 120 桁に収めるため。年は付けない。
# 年をまたいでも、停止の列の年と停止時間で分かる。年を付けると、その列の分だけ表のすべての行が広がる）
function Format-BootTime($Stop, $Boot) {
    if ($null -eq $Boot) { return '（記録なし）' }
    $b = [datetime]$Boot
    if ($null -ne $Stop -and ([datetime]$Stop).Date -eq $b.Date) { return $b.ToString('HH:mm', $Invariant) }
    return $b.ToString('MM-dd HH:mm', $Invariant)
}

#endregion

#region 引数の確認

if ($NoSave -and $OutputFolder) { Stop-WithError '-NoSave のときは、-OutputFolder は指定できません' }

$startTime = Get-Date
$sinceInfo = $null
$untilInfo = $null
if ($Since) { $sinceInfo = ConvertTo-PointInTime $Since '-Since' }
if ($Until) { $untilInfo = ConvertTo-PointInTime $Until '-Until' }
if ($sinceInfo) {
    $endTime = if ($untilInfo) { $untilInfo.Time } else { $startTime }
    if ($sinceInfo.Time -ge $endTime) {
        $endLabel = if ($untilInfo) { '-Until の日時' } else { '今' }
        Stop-WithError (Format-Invariant '-Since の日時（{0:yyyy-MM-dd HH:mm}）が、{1}（{2:yyyy-MM-dd HH:mm}）と同じか、それより後になっています' $sinceInfo.Time $endLabel $endTime)
    }
}

$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    if ($outText -match '^(\\\\\?\\)?[A-Za-z]:$') { $outText += '\' }
    $outText = Resolve-StartPath $outText
}

$skipped   = [System.Collections.Generic.List[object]]::new()
$skipKinds = [ordered]@{ '権限なし' = 0; 'ログがない' = 0; 'ファイルを読めない' = 0; '対象外のログ' = 0; 'リンク' = 0; 'その他のエラー' = 0 }
function Add-Skip([string]$Reason, [string]$Target, [string]$Detail) {
    $skipped.Add([object[]]@($Reason, $Target, $Detail))
    $skipKinds[$Reason]++
}

# 読む対象。Oldest はログに残っているいちばん古い記録の日時
$sources = [System.Collections.Generic.List[object]]::new()
$useFiles = [bool]$Path
if ($useFiles) {
    $files = [System.Collections.Generic.List[string]]::new()
    foreach ($item in (Split-List $Path)) {
        if ($item.Contains('"')) { Stop-WithQuoteError '-Path' $item }
        $full = ConvertTo-DisplayPath (Resolve-StartPath $item)
        try {
            $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($full)
            $full = [IO.Path]::GetFullPath($full)
        } catch { }
        if ([IO.File]::Exists((ConvertTo-LongPath $full))) {
            $files.Add($full)
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
                $subDirs = [System.Collections.Generic.List[string]]::new()
                foreach ($entry in ($entries | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_.Name } })) {
                    $linkName = Get-LinkName $entry
                    if ($linkName -and $entry -is [IO.DirectoryInfo]) {
                        Add-Skip 'リンク' (ConvertTo-DisplayPath $entry.FullName) $linkName
                        continue
                    }
                    if ($entry -is [IO.DirectoryInfo]) { $subDirs.Add($entry.FullName) }
                    elseif ($entry.Extension -ieq '.evtx') { $files.Add((ConvertTo-DisplayPath $entry.FullName)) }
                }
                for ($i = $subDirs.Count - 1; $i -ge 0; $i--) { $stack.Push($subDirs[$i]) }
            }
        } else {
            Add-Skip 'ファイルを読めない' $full '見つかりません'
        }
    }
    # 同じファイルを重ねて指定したとき（ファイルと、それを含むフォルダなど）は 1 回だけ読む
    $uniqueFiles = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    # 何のログかは、ファイルの最初の記録で見分ける
    foreach ($file in $files) {
        if (-not $uniqueFiles.Add($file)) { continue }
        Write-Progress -Activity $ToolTitle -Status 'ファイルを調べています' -CurrentOperation $file
        # 260 文字以上のパスは \\?\ を付けないと読めない
        $source = @{ Label = $file; Name = $(if ($file.Length -ge $LongPathLength) { ConvertTo-LongPath $file } else { $file }); IsFile = $true; Oldest = $null; Computer = ''; Failed = $false }
        try {
            $first = Read-FirstEvent $source '*'
        } catch {
            if (Test-InnerException $_.Exception ([UnauthorizedAccessException])) { Add-Skip '権限なし' $file 'アクセスが拒否されました'; continue }
            Add-Skip 'ファイルを読めない' $file (Get-ErrorText (Get-InnerException $_.Exception))
            continue
        }
        if ($null -eq $first) { Add-Skip '対象外のログ' $file '記録がありません'; continue }
        if ($first.Log -ine $SystemLog) { Add-Skip '対象外のログ' $file ('System ログではありません（{0}）' -f $first.Log); continue }
        $source.Oldest = $first.Time
        # コンピューターの名前は、いちばん新しい記録のもの（構築の後に名前を変えたサーバーは、古い記録が前の名前になっているため）。
        # 最後のほうが壊れていて新しい順に読めないときは、いちばん古い記録の名前を使う
        $machine = $first.Machine
        try {
            $last = Read-FirstEvent $source '*' $true
            if ($last) { $machine = $last.Machine }
        } catch { }
        $source.Computer = Get-ComputerKey $machine
        $sources.Add($source)
    }
    Write-Progress -Activity $ToolTitle -Completed
    if ($sources.Count -eq 0) {
        $systemFound = @($skipped | Where-Object { $_[0] -eq '対象外のログ' }).Count -gt 0 -and @($skipped | Where-Object { $_[0] -ne '対象外のログ' -and $_[0] -ne 'リンク' }).Count -eq 0
        if ($systemFound) {
            # 読めるファイルはあるが、System ログではない（Application.evtx だけを指定したなど）
            Stop-WithError ('-Path に、System ログの .evtx がありません: {0}' -f $Path + "`r`n" + '停止と起動の記録は System ログ（System.evtx）にあります')
        }
        # 打ち間違いか、権限か、形式の違うファイルかが分かるように、理由を添える
        $reasons = @($skipped | Select-Object -First 5 | ForEach-Object { '  {0}: {1}（{2}）' -f $_[1], $_[0], $_[2] })
        $message = '読める .evtx ファイルがありません: {0}' -f $Path
        if ($reasons.Count -gt 0) { $message += "`r`n" + ($reasons -join "`r`n") }
        if ($skipped.Count -gt $reasons.Count) { $message += "`r`n" + (Format-Invariant '  ほか {0:N0} 件' ($skipped.Count - $reasons.Count)) }
        Stop-WithError $message
    }
} else {
    # この PC の System ログ。読めなければ、集め始める前に止める
    $source = @{ Label = $SystemLog; Name = $SystemLog; IsFile = $false; Oldest = $null; Computer = (Get-ComputerKey $machineName); Failed = $false }
    try {
        $first = Read-FirstEvent $source '*'
    } catch {
        if (Test-InnerException $_.Exception ([UnauthorizedAccessException])) {
            Stop-WithError ('System ログを読む権限がありません' + "`r`n" + '管理者として実行してください。ほかの PC から集めた .evtx ファイルは、-Path で読めます')
        }
        Stop-WithError ('System ログを読めません: {0}' -f (Get-ErrorText (Get-InnerException $_.Exception)))
    }
    if ($first) { $source.Oldest = $first.Time }
    $sources.Add($source)
}

#endregion

#region 保存先の準備

# 読み終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir = $null
if (-not $NoSave) {
    $baseDir = if ($outText) { $outText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        # \\?\ 付きのパスでは末尾のピリオドや空白が残るため、先に通常の書き方で整える（260 文字を超えると例外になるので、そのときはそのまま）
        try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
        $runName = Format-Invariant 'RebootHistory_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while ([IO.Directory]::Exists((ConvertTo-LongPath $runDir)) -or [IO.File]::Exists((ConvertTo-LongPath $runDir))) {
            $suffix++
            $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix)
        }
        [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $runDir))
    } catch {
        $ex = Get-InnerException $_.Exception
        $message = '保存先に書き込めません: {0}（{1}）' -f $baseDir, (Get-ErrorText $ex)
        # 管理者として開いた PowerShell からは、通常の権限で割り当てたネットワークドライブが見えない
        $isAdmin = $false
        try {
            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        } catch { }
        if ($isAdmin -and $ex -is [Management.Automation.DriveNotFoundException]) {
            $message += "`r`n" + '管理者として実行した PowerShell からは、通常の権限で割り当てたネットワークドライブは見えません。\\server\share の形で指定してください'
        }
        Stop-WithError ($message + "`r`n" + '-NoSave を付けると、画面に表示だけします')
    }
}

#endregion

$finished   = $false
$fatal      = $false
$unexpected = $false
$saveError  = $null
$sw         = [Diagnostics.Stopwatch]::StartNew()
try {
    #region 読み取り

    # 記録の値は、名前で取り出す（記録の版によって値の並びが違うため）。名前のない古い形の記録（6008）は Properties で読む
    function New-Selector([string[]]$Names) {
        return New-Object Diagnostics.Eventing.Reader.EventLogPropertySelector (, [string[]]@($Names | ForEach-Object { "Event/EventData/Data[@Name='$_']" }))
    }
    $select12   = New-Selector @('StartTime', 'MajorVersion', 'MinorVersion', 'BuildVersion', 'QfeVersion')
    $select13   = New-Selector @('StopTime')
    $select27   = New-Selector @('BootType')
    $select109  = New-Selector @('ShutdownActionType')
    $select41   = New-Selector @('BugcheckCode', 'PowerButtonTimestamp', 'SleepInProgress', 'LongPowerButtonPressDetected')
    $select1074 = New-Selector @('param1', 'param2', 'param3', 'param4', 'param5', 'param6', 'param7')
    $select1075 = New-Selector @('param1', 'param2')
    $select1076 = New-Selector @('param1', 'param2', 'param3', 'param4', 'param5', 'param6')
    $select1001 = New-Selector @('param1', 'param2')
    # ログの消去は、値が UserData（LogFileCleared）に入る。EventData に入る版があっても読めるよう、両方を見る
    $select104  = New-Object Diagnostics.Eventing.Reader.EventLogPropertySelector (, [string[]]@('Event/UserData/*/SubjectDomainName', 'Event/UserData/*/SubjectUserName', 'Event/UserData/*/Channel',
        "Event/EventData/Data[@Name='SubjectDomainName']", "Event/EventData/Data[@Name='SubjectUserName']", "Event/EventData/Data[@Name='Channel']"))

    $records = [System.Collections.Generic.List[object]]::new()
    $seen    = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    # 記録を 1 件、一覧に入れる。組にするときのコンピューターは、読んだログ（ファイル）ごとに決める（記録に書かれた名前は、
    # 名前を変える前の記録だと前の名前なので）。Time は組み合わせに使う時刻（OS の起動・停止は、記録の値の時刻）
    function Add-EventRecord($Record, $Source) {
        $computerText = [string]$Record.MachineName
        $created = $Record.TimeCreated
        # 同じ記録を、重なった .evtx から 2 回読まないようにする。ログを消去すると記録の番号が振り直されることがあるので、日時も含める
        if (-not $seen.Add(('{0}|{1}|{2}|{3}' -f $computerText, $Record.LogName, $Record.RecordId, $created.Ticks))) { return }
        $id = [int]$Record.Id
        # 項目は先に全部用意する（StrictMode では、ないプロパティを読むとエラーになるため）
        $rec = [pscustomobject]@{
            Computer = $Source.Computer; ComputerText = $computerText; Time = $created; Created = $created; RecordId = [long]$Record.RecordId; Seq = $records.Count
            EventId = $id; Provider = [string]$Record.ProviderName; User = ''; Process = ''; Reason = ''; Code = ''; CodeValue = [uint64]0; Action = $null; ActionText = ''
            Comment = ''; Version = ''; BootType = $null; Bugcheck = [uint64]0; PowerButton = $false; Sleep = $false; BugText = ''; Dump = ''; Channel = ''; StopTime = $null; Detail = ''
            ServiceRestart = $false
        }
        switch ($id) {
            12 {
                $v = $Record.GetPropertyValues($select12)
                $start = ConvertTo-LocalTime $v[0]
                if ($null -ne $start) { $rec.Time = $start }
                if ($null -ne $v[3]) { $rec.Version = Format-Invariant '{0}.{1}.{2}.{3}' (ConvertTo-UInt64Value $v[1]) (ConvertTo-UInt64Value $v[2]) (ConvertTo-UInt64Value $v[3]) (ConvertTo-UInt64Value $v[4]) }
                if ($rec.Version) { $rec.Detail = 'OS のバージョン: ' + $rec.Version }
            }
            13 {
                $v = $Record.GetPropertyValues($select13)
                $stop = ConvertTo-LocalTime $v[0]
                if ($null -ne $stop) { $rec.Time = $stop }
            }
            27 {
                $v = $Record.GetPropertyValues($select27)
                if ($null -ne $v[0]) { $rec.BootType = [int](ConvertTo-UInt64Value $v[0]) }
                $rec.Detail = switch ($rec.BootType) { 0 { '通常の起動' } 1 { '高速スタートアップ' } 2 { '休止状態からの再開' } default { if ($null -ne $rec.BootType) { '種類 {0}' -f $rec.BootType } else { '' } } }
            }
            109 {
                $v = $Record.GetPropertyValues($select109)
                $type = ConvertTo-UInt64Value $v[0]
                $rec.Action = Get-ActionFromType $type
                $rec.Detail = switch ($type) { 4 { 'シャットダウン' } 5 { '再起動' } 6 { '電源を切る' } default { '種類 {0}' -f $type } }
            }
            41 {
                $v = $Record.GetPropertyValues($select41)
                $rec.Bugcheck = ConvertTo-UInt64Value $v[0]
                $rec.PowerButton = (ConvertTo-UInt64Value $v[1]) -ne 0 -or (ConvertTo-UInt64Value $v[3]) -ne 0
                $rec.Sleep = (ConvertTo-UInt64Value $v[2]) -ne 0
                if ($rec.Bugcheck -ne 0) { $rec.Detail = '停止コード ' + (Format-Bugcheck $rec.Bugcheck) }
                elseif ($rec.PowerButton) { $rec.Detail = '電源ボタンを押した記録があります' }
                elseif ($rec.Sleep) { $rec.Detail = 'スリープの途中でした' }
            }
            1074 {
                $v = $Record.GetPropertyValues($select1074)
                $rec.Process = Get-ProcessPath ([string]$v[0])
                $rec.Reason = ConvertTo-CellText $v[2]
                $rec.Code = ConvertTo-CellText $v[3]
                $rec.CodeValue = ConvertTo-UInt64Value $rec.Code
                $rec.ActionText = ConvertTo-CellText $v[4]
                $rec.Action = Get-ActionFromText $rec.ActionText
                $rec.Comment = ConvertTo-CellText $v[5]
                $rec.User = ConvertTo-CellText $v[6]
                # events.csv には、記録の文字のまま書く（「電源を切る」など）
                $rec.Detail = '{0}: {1}（{2}）' -f $rec.ActionText, $rec.Reason, $rec.Code
                if ($rec.Comment) { $rec.Detail += ' コメント: ' + $rec.Comment }
            }
            1075 {
                $v = $Record.GetPropertyValues($select1075)
                $rec.User = ConvertTo-CellText $v[1]
            }
            1076 {
                $v = $Record.GetPropertyValues($select1076)
                $rec.Reason = ConvertTo-CellText $v[0]
                $rec.Code = ConvertTo-CellText $v[1]
                $rec.Comment = ConvertTo-CellText $v[4]
                $rec.User = ConvertTo-CellText $v[5]
                $rec.Detail = '{0}（{1}）' -f $rec.Reason, $rec.Code
                if ($rec.Comment) { $rec.Detail += ' コメント: ' + $rec.Comment }
            }
            6008 {
                # 停止の時刻は、記録の中のバイナリの UTC の SYSTEMTIME（説明の文の日時は、地域の設定で書式が変わり、見えない文字も入る）。
                # UTC が読めなければ、その前のローカルの SYSTEMTIME を使う
                $bytes = $null
                foreach ($property in $Record.Properties) { if ($property.Value -is [byte[]]) { $bytes = [byte[]]$property.Value } }
                $rec.StopTime = ConvertFrom-SystemTime $bytes 16 ([DateTimeKind]::Utc)
                if ($null -eq $rec.StopTime) { $rec.StopTime = ConvertFrom-SystemTime $bytes 0 ([DateTimeKind]::Local) }
                if ($null -ne $rec.StopTime) { $rec.Detail = '停止の時刻: ' + (Format-Time $rec.StopTime) }
            }
            1001 {
                $v = $Record.GetPropertyValues($select1001)
                $rec.BugText = ConvertTo-CellText $v[0]
                $rec.Dump = ConvertTo-CellText $v[1]
                if ($rec.BugText -match '^0[xX]([0-9A-Fa-f]{1,8})\b') { $rec.Bugcheck = [Convert]::ToUInt64($Matches[1], 16) }
                $rec.Detail = $rec.BugText
                if ($rec.Dump) { $rec.Detail += ' ダンプ: ' + $rec.Dump }
            }
            104 {
                $v = $Record.GetPropertyValues($select104)
                $domain = ConvertTo-CellText $(if ($null -ne $v[0]) { $v[0] } else { $v[3] })
                $name = ConvertTo-CellText $(if ($null -ne $v[1]) { $v[1] } else { $v[4] })
                $rec.User = if ($domain -and $domain -ne '-') { $domain + '\' + $name } else { $name }
                $rec.Channel = ConvertTo-CellText $(if ($null -ne $v[2]) { $v[2] } else { $v[5] })
                # System ログのほかのログの消去は、停止と起動に関わらないので読まない
                if ($rec.Channel -and $rec.Channel -ine $SystemLog) { return }
                $rec.Detail = '{0} ログを消去' -f $(if ($rec.Channel) { $rec.Channel } else { $SystemLog })
            }
        }
        $records.Add($rec)
    }

    $readCount = 0
    $nextProgress = 0L
    foreach ($source in $sources) {
        $status = 'System ログを読んでいます'
        Write-Progress -Activity $ToolTitle -Status $status -CurrentOperation $source.Label
        $countBefore = $readCount
        $reader = $null
        try {
            # 古い順に読む（壊れた .evtx でも、壊れた場所の手前まで読めるため）
            $reader = New-EventReader $source $RecordQuery
            while ($true) {
                $record = $reader.ReadEvent()
                if ($null -eq $record) { break }
                try { Add-EventRecord $record $source } finally { $record.Dispose() }
                $readCount++
                if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                    Write-Progress -Activity $ToolTitle -Status (Format-Invariant '{0}（記録 {1:N0} 件）' $status $readCount) -CurrentOperation $source.Label
                    $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
                }
            }
        } catch {
            $ex = Get-InnerException $_.Exception
            # 1 件も読めなかったログは、後で「すべて読めなかった」かを判断するために印を付ける
            if ($readCount -eq $countBefore) { $source.Failed = $true }
            if (Test-InnerException $_.Exception ([UnauthorizedAccessException])) {
                Add-Skip '権限なし' $source.Label 'アクセスが拒否されました'
            } elseif (Test-InnerException $_.Exception ([Diagnostics.Eventing.Reader.EventLogNotFoundException])) {
                Add-Skip 'ログがない' $source.Label (Get-ErrorText $ex)
            } else {
                $detail = Get-ErrorText $ex
                if ($readCount -gt $countBefore) { $detail = Format-Invariant '途中まで読んだ（記録 {0:N0} 件は結果に含む）: {1}' ($readCount - $countBefore) $detail }
                Add-Skip $(if ($useFiles) { 'ファイルを読めない' } else { 'その他のエラー' }) $source.Label $detail
            }
        } finally {
            if ($reader) { $reader.Dispose() }
        }
    }

    # System ログを 1 件も読めなかったときは、「該当なし」と区別して止める（作った保存先のフォルダは finally で消す）
    if (@($sources | Where-Object { -not $_.Failed }).Count -eq 0) {
        $reasons = @($skipped | Select-Object -First 5 | ForEach-Object { '  {0}: {1}（{2}）' -f $_[1], $_[0], $_[2] })
        $fatal = $true
        $finished = $true
        Stop-WithError (('System ログを読めませんでした' + "`r`n") + ($reasons -join "`r`n"))
    }

    #endregion

    #region 停止と起動を組み合わせる

    Write-Progress -Activity $ToolTitle -Status '停止と起動を組み合わせています'

    # 停止の前の最後の記録（停止した時刻が分からないときの目安）。停止のときの記録は次の起動のあとに書かれ、記録の番号の順と
    # 時刻の順が一致しないので、起動より前の時刻の記録を新しい順にいくつか見て、いちばん新しい時刻を取る
    function Get-LastRecordBefore([string]$Computer, [datetime]$Time) {
        $latest = $null
        foreach ($src in $sources) {
            if ($src.Computer -ne $Computer -or $src.Failed) { continue }
            $reader = $null
            try {
                $reader = New-EventReader $src ("*[System[TimeCreated[@SystemTime<'{0}']]]" -f (ConvertTo-QueryTime $Time)) $true
                for ($n = 0; $n -lt $LastRecordScan; $n++) {
                    $record = $reader.ReadEvent()
                    if ($null -eq $record) { break }
                    try {
                        $t = $record.TimeCreated
                        if ($null -ne $t -and ($null -eq $latest -or $t -gt $latest)) { $latest = $t }
                    } finally { $record.Dispose() }
                }
            } catch {
                # 最後のほうが壊れた .evtx では新しい順に読めない。目安なので、分からなければ空のままにする
            } finally {
                if ($reader) { $reader.Dispose() }
            }
        }
        return $latest
    }

    # 起動の時刻（昇順）の配列から、Time 以前の起動の数を数える（二分探索）
    function Get-BootCount([datetime[]]$Times, [datetime]$Time) {
        $low = 0
        $high = $Times.Length
        while ($low -lt $high) {
            $mid = [int][math]::Floor(($low + $high) / 2)
            if ($Times[$mid] -le $Time) { $low = $mid + 1 } else { $high = $mid }
        }
        return $low
    }

    # 時刻の順に並んだ起動の一覧から、Time の前後 Seconds 秒の内で、いちばん近い起動（二分探索なので、起動が多くても遅くならない）
    function Find-NearBoot($Boots, [datetime[]]$Times, [datetime]$Time, [int]$Seconds) {
        $n = Get-BootCount $Times $Time
        $best = $null
        $bestDiff = 0.0
        foreach ($i in @(($n - 1), $n)) {
            if ($i -lt 0 -or $i -ge $Times.Length) { continue }
            $diff = [math]::Abs(($Time - $Times[$i]).TotalSeconds)
            if ($diff -le $Seconds -and ($null -eq $best -or $diff -lt $bestDiff)) { $best = $Boots[$i]; $bestDiff = $diff }
        }
        return $best
    }

    # 起動。Assign は、予期しない停止の記録（起動のあとに書かれる）を割り当てるときの、起動の時刻。
    # Source は、どの記録から分かった起動か（12・6005・27、OS の情報なら Os）
    function New-Boot([datetime]$Time, [string]$Version, $BootType, [string]$ComputerText, [datetime]$Assign, [string]$Source) {
        return [pscustomobject]@{ Time = $Time; Version = $Version; BootType = $BootType; ComputerText = $ComputerText; Assign = $Assign; Source = $Source }
    }

    # 時刻の秒の端数を切り捨ててから引く（CSV の秒までの時刻から計算した長さと、1 秒ずれないように）
    function Measure-Span([datetime]$From, [datetime]$To) {
        $f = $From.AddTicks(-($From.Ticks % [TimeSpan]::TicksPerSecond))
        $t = $To.AddTicks(-($To.Ticks % [TimeSpan]::TicksPerSecond))
        return $t - $f
    }

    # 今の起動（この PC を調べたときだけ）。OS の情報の最後の起動の時刻と、記録のいちばん新しい起動の、新しいほうにする（後で決める）
    $cimBoot = $null
    if (-not $useFiles) {
        try { $cimBoot = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime } catch { }
    }

    $rows     = [System.Collections.Generic.List[object]]::new()
    $pending  = [System.Collections.Generic.List[object]]::new()   # 最後の起動のあとの、まだ止まっていない停止の指示
    $lastBoots = @{}                                              # コンピューター → いちばん新しい起動の時刻

    $computers = @($records | ForEach-Object { $_.Computer } | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } } -Unique)
    foreach ($computer in $computers) {
        # Sort-Object は同じキーのものの並びが 5.1 と 7 で違うので、記録の番号と読んだ順まで並べて、結果をそろえる
        $list = @($records | Where-Object { $_.Computer -eq $computer } | Sort-Object -Property Time, RecordId, Seq)

        # 起動。OS の起動（12）を区切りにする。12 の記録が 1 件もないコンピューターは、イベントログの起動（6005）で補う。
        # 起動の種類（27）だけの起動（高速スタートアップ）も補う。$list は時刻の順なので、12 の起動も時刻の順に並ぶ
        $kernelBoots = [System.Collections.Generic.List[object]]::new()
        foreach ($rec in $list) {
            if ($rec.EventId -eq 12) { $kernelBoots.Add((New-Boot $rec.Time $rec.Version $null $rec.ComputerText $rec.Time '12')) }
        }
        $kernelArray = $kernelBoots.ToArray()
        $kernelTimes = [datetime[]]::new($kernelArray.Length)
        for ($i = 0; $i -lt $kernelArray.Length; $i++) { $kernelTimes[$i] = $kernelArray[$i].Time }
        $bootList = [System.Collections.Generic.List[object]]::new()
        $bootList.AddRange([object[]]$kernelArray)
        foreach ($rec in $list) {
            if ($rec.EventId -ne 6005) { continue }
            # OS の起動のあと 10 分の内のイベントログの起動は、同じ起動の記録
            $n = Get-BootCount $kernelTimes $rec.Time
            if ($n -gt 0 -and ($rec.Time - $kernelTimes[$n - 1]).TotalSeconds -le $BootMergeSeconds) { continue }
            if ($kernelArray.Length -gt 0) {
                # OS の起動を記録するコンピューターで、OS の起動のない 6005 は、ふつうはイベントログのサービスだけを再起動したもの
                # （OS は再起動していない）。ただ、壊れた .evtx などで OS の起動の記録だけが読めていないこともあるので、
                # 次の形のときだけサービスの再起動とみなし、直前の 6006 も停止の記録として使わない
                # ・6006 があれば、その後 60 秒の内の 6005 で、間に OS の停止（13）と停止の開始（109）がない（OS を再起動すると、間にこの 2 つが入る）
                # ・6006 がなければ（サービスが落ちたとき）、前の起動の後に停止の記録（1074・109・13）がない
                # ・どちらも、6005 の前後に予期しない停止の記録（6008・41）がない
                $from = if ($n -gt 0) { $kernelTimes[$n - 1] } else { [datetime]::MinValue }
                $stopRecord = $null
                $kernelStops = [System.Collections.Generic.List[datetime]]::new()
                $orders = 0
                $crashed = $false
                foreach ($other in $list) {
                    $created = $other.Created
                    if (($other.EventId -eq 6008 -or $other.EventId -eq 41) -and [math]::Abs(($created - $rec.Time).TotalSeconds) -le $BootMergeSeconds) { $crashed = $true }
                    if ($other.Time -le $from -or $other.Time -gt $rec.Time) { continue }
                    switch ($other.EventId) {
                        6006 { if (-not $other.ServiceRestart) { $stopRecord = $other } }
                        13 { $kernelStops.Add($other.Time) }
                        109 { $kernelStops.Add($other.Time) }
                        1074 { $orders++ }
                    }
                }
                $isService = $false
                if (-not $crashed) {
                    if ($stopRecord) {
                        $between = @($kernelStops | Where-Object { $_ -ge $stopRecord.Time }).Count
                        $isService = ($rec.Time - $stopRecord.Time).TotalSeconds -le $ServiceRestartSeconds -and $between -eq 0
                    } else {
                        $isService = $kernelStops.Count -eq 0 -and $orders -eq 0
                    }
                }
                if ($isService) {
                    $rec.ServiceRestart = $true
                    $rec.Detail = 'イベントログのサービスの再起動（OS は再起動していません）'
                    if ($stopRecord) { $stopRecord.ServiceRestart = $true; $stopRecord.Detail = $rec.Detail }
                    continue
                }
            }
            # 6005 は OS の起動より遅れて書かれ、予期しない停止の記録（41 など）のほうが先に書かれることがあるので、割り当ての時刻を前にずらす
            $bootList.Add((New-Boot $rec.Time '' $null $rec.ComputerText $rec.Time.AddSeconds(-$EventLogDelaySeconds) '6005'))
        }
        $mergedArray = @($bootList.ToArray() | Sort-Object -Property Time)
        $mergedTimes = [datetime[]]::new($mergedArray.Length)
        for ($i = 0; $i -lt $mergedArray.Length; $i++) { $mergedTimes[$i] = $mergedArray[$i].Time }
        foreach ($rec in $list) {
            if ($rec.EventId -ne 27 -or $null -eq $rec.BootType) { continue }
            $near = Find-NearBoot $mergedArray $mergedTimes $rec.Time $BootMergeSeconds
            if ($near) {
                if ($null -eq $near.BootType) { $near.BootType = $rec.BootType }
                # 6005 で補った起動は、起動の種類（27）のほうが OS の起動に近い時刻なので、その時刻にする
                if ($near.Source -eq '6005' -and $rec.Time -lt $near.Time) { $near.Time = $rec.Time; $near.Assign = $rec.Time; $near.Source = '27' }
                continue
            }
            # 休止状態からの再開（2）は、起動し直していないので数えない
            if ($rec.BootType -ne 2) { $bootList.Add((New-Boot $rec.Time '' $rec.BootType $rec.ComputerText $rec.Time '27')) }
        }
        $boots = @($bootList.ToArray() | Sort-Object -Property Time)
        $bootTimes = [datetime[]]::new($boots.Count)
        $assignTimes = [datetime[]]::new($boots.Count)
        for ($i = 0; $i -lt $boots.Count; $i++) {
            $bootTimes[$i] = $boots[$i].Time
            # 二分探索のため、割り当ての時刻も順に並ぶようにする
            $assignTimes[$i] = if ($i -gt 0 -and $boots[$i].Assign -lt $assignTimes[$i - 1]) { $assignTimes[$i - 1] } else { $boots[$i].Assign }
        }
        # 今の起動に使う、いちばん新しい起動（OS の起動か起動の種類の記録から分かったもの。6005 は OS の起動より遅れて書かれるため使わない）
        for ($i = $boots.Count - 1; $i -ge 0; $i--) {
            if ($boots[$i].Source -ne '6005') { $lastBoots[$computer] = $boots[$i].Time; break }
        }

        # 記録を、起動と起動の間（区切り）に分ける。区切り k は、起動 k の前（k = 起動の数なら、最後の起動の後）。
        # 停止の記録は、時刻で区切りに入れる。予期しない停止の記録（6008・41・1001・1076）は起動のあとに書かれるので、
        # 書かれた時刻の直前の起動の前の区切りに入れる（余裕を足すと、起動の間隔が短いときに次の停止に付いてしまう）
        $segmentCount = $boots.Count + 1
        $pre  = [object[]]::new($segmentCount)
        $post = [object[]]::new($segmentCount)
        for ($k = 0; $k -lt $segmentCount; $k++) { $pre[$k] = [System.Collections.Generic.List[object]]::new(); $post[$k] = [System.Collections.Generic.List[object]]::new() }
        foreach ($rec in $list) {
            switch ($rec.EventId) {
                { $_ -eq 1074 -or $_ -eq 1075 -or $_ -eq 109 -or $_ -eq 13 -or $_ -eq 6006 } {
                    # イベントログのサービスの再起動の 6006 は、OS の停止ではない
                    if (-not $rec.ServiceRestart) { $pre[(Get-BootCount $bootTimes $rec.Time)].Add($rec) }
                }
                { $_ -eq 6008 -or $_ -eq 41 -or $_ -eq 1001 -or $_ -eq 1076 } {
                    $k = (Get-BootCount $assignTimes $rec.Created) - 1
                    # 残っている最初の起動より前の記録は、その起動が分からないので組み合わせない（events.csv には書く）
                    if ($k -ge 0) { $post[$k].Add($rec) }
                }
            }
        }

        for ($k = 0; $k -lt $segmentCount; $k++) {
            $boot = if ($k -lt $boots.Count) { $boots[$k] } else { $null }
            $prevBoot = if ($k -gt 0) { $boots[$k - 1] } else { $null }
            # 停止の指示は、取り消し（1075）のない、最後のもの
            $order = $null
            $last13 = $null; $last6006 = $null; $last109 = $null
            foreach ($rec in $pre[$k]) {
                switch ($rec.EventId) {
                    1074 { $order = $rec }
                    1075 { $order = $null }
                    13 { $last13 = $rec }
                    6006 { $last6006 = $rec }
                    109 { $last109 = $rec }
                }
            }
            $r6008 = $null; $r41 = $null; $r1001 = $null; $r1076 = $null
            foreach ($rec in $post[$k]) {
                switch ($rec.EventId) {
                    6008 { if ($null -eq $r6008) { $r6008 = $rec } }
                    41 { if ($null -eq $r41) { $r41 = $rec } }
                    1001 { if ($null -eq $r1001) { $r1001 = $rec } }
                    1076 { if ($null -eq $r1076) { $r1076 = $rec } }
                }
            }
            $hasStop = $null -ne $last13 -or $null -ne $last6006 -or $null -ne $last109
            $isUnexpected = $null -ne $r6008 -or $null -ne $r41 -or $null -ne $r1001 -or $null -ne $r1076

            # 最後の起動の後：止まっていなければ、まだ止まっていない停止の指示として注意に出す
            $bootFromOs = $false
            if ($null -eq $boot) {
                if (-not $hasStop) {
                    if ($order) { $pending.Add($order) }
                    continue
                }
                # この PC は今動いているので、最後の起動の後に停止の記録があれば、そのあとの起動の記録が残っていない。
                # OS の情報の最後の起動が停止より後なら、それを起動とする。そうでなければ（休止状態など）、行にしない
                if (-not $useFiles) {
                    $lastStop = @($pre[$k] | Where-Object { $_.EventId -ne 1074 -and $_.EventId -ne 1075 } | ForEach-Object { $_.Time } | Sort-Object)[-1]
                    if ($null -eq $cimBoot -or $cimBoot -le $lastStop) { continue }
                    $boot = New-Boot $cimBoot '' $null $pre[$k][0].ComputerText $cimBoot 'Os'
                    $bootFromOs = $true
                }
            } elseif (-not $hasStop -and -not $isUnexpected -and $null -eq $order) {
                # 停止の記録がまったくない。残っている最初の起動の前は、ログの始まりなので出さない
                if ($null -eq $prevBoot) { continue }
            }

            $row = [pscustomobject]@{
                Computer = $computer; ComputerText = $(if ($boot) { $boot.ComputerText } else { $pre[$k][0].ComputerText }); Stop = $null; Boot = $null; Kind = '不明'
                WindowsUpdate = $false; Trigger = ''; User = ''; Reason = ''; Code = ''; Comment = ''; Cause = ''; Bugcheck = ''; Dump = ''; Later = ''; Ordered = $null; Process = ''
                BootType = ''; Version = ''; LastRecord = $null; Notes = [System.Collections.Generic.List[string]]::new(); Screen = ''; SortTime = $null; Seq = $rows.Count
                PrevBoot = $(if ($prevBoot) { $prevBoot.Time } else { $null })
            }
            if ($boot) {
                $row.Boot = $boot.Time
                $row.Version = $boot.Version
                if ($prevBoot -and $prevBoot.Version -and $boot.Version -and $prevBoot.Version -ne $boot.Version) { $row.Version = '{0} → {1}' -f $prevBoot.Version, $boot.Version }
                $row.BootType = switch ($boot.BootType) { 0 { '通常' } 1 { '高速スタートアップ' } default { if ($null -ne $boot.BootType) { '種類 {0}' -f $boot.BootType } else { '' } } }
            }
            if ($order) {
                $row.Trigger = Get-TriggerLabel $order.Process
                $row.User = $order.User
                $row.Reason = $order.Reason
                $row.Code = $order.Code
                $row.Comment = $order.Comment
                $row.Ordered = $order.Time
                $row.Process = $order.Process
                # Windows Update による停止：プロセスか、理由のコード（オペレーティング システムの、アップグレード・サービス パック・修正プログラム・セキュリティ修正）
                $code = $order.CodeValue
                $row.WindowsUpdate = $row.Trigger -eq 'Windows Update' -or ((($code -band 0x00FF0000) -eq 0x00020000) -and $UpdateMinorCodes -contains [int]($code -band 0xFFFF))
            }
            # 再起動かシャットダウンか（停止の開始の値を先に見る。停止の指示の文字は、記録を書いた OS の言語で入っているため）
            $action = $null
            if ($last109) { $action = $last109.Action }
            if ($null -eq $action -and $order) { $action = $order.Action }

            if ($isUnexpected) {
                $row.Kind = '予期しない停止'
                # 停止の時刻：予期しない停止の記録の時刻 → なければ、OS の停止 → イベントログの停止 → 停止の開始
                # （停止の指示の時刻は使わない。時間をおいた停止の指示のあと、ずっと後で止まったこともあるため）
                if ($r6008) { $row.Stop = $r6008.StopTime }
                if ($null -eq $row.Stop) {
                    if ($last13) { $row.Stop = $last13.Time } elseif ($last6006) { $row.Stop = $last6006.Time } elseif ($last109) { $row.Stop = $last109.Time }
                }
                # 原因
                $code = [uint64]0
                if ($r41 -and $r41.Bugcheck -ne 0) { $code = $r41.Bugcheck } elseif ($r1001 -and $r1001.Bugcheck -ne 0) { $code = $r1001.Bugcheck }
                if ($code -ne 0) {
                    $row.Bugcheck = '0x' + ([uint32]($code -band 0xFFFFFFFFL)).ToString('X8', $Invariant)
                    $row.Cause = 'ブルースクリーン ' + (Format-Bugcheck $code)
                } elseif ($r1001) {
                    $row.Cause = 'ブルースクリーン'
                } elseif ($r41 -and $r41.PowerButton) {
                    $row.Cause = '電源ボタンの長押し'
                } elseif ($r41 -and $r41.Sleep) {
                    $row.Cause = 'スリープの途中で停止'
                } else {
                    $row.Cause = '不明（電源断・フリーズ・強制的な電源オフなど）'
                }
                if ($r1001) { $row.Dump = $r1001.Dump }
                if ($r1076) {
                    $row.Later = $r1076.Reason
                    if ($r1076.Comment) { $row.Later += '「{0}」' -f $r1076.Comment }
                    if ($r1076.User) { $row.Later += '（{0}）' -f $r1076.User }
                }
                # 停止の途中で止まった（停止の指示や停止の記録のあとに、予期しない停止が記録された）
                $during = ''
                if ($order -or $hasStop) {
                    $label = Get-ActionLabel $action
                    $during = if ($label) { $label } else { 'シャットダウンか再起動' }
                    $row.Notes.Add(('{0}の途中で止まりました' -f $during))
                }
                $screen = $row.Cause
                if ($during) {
                    $who = if ($row.Trigger) { '（{0}）' -f $row.Trigger } else { '' }
                    $screen = '{0}の途中で停止{1}: {2}' -f $during, $who, $row.Cause
                }
                # 管理者が入力した理由は、画面の幅で切られないよう先に書く
                if ($r1076) { $screen = '入力された理由: {0} ／ {1}' -f $(if ($r1076.Comment) { $r1076.Comment } else { $r1076.Reason }), $screen }
                $row.Screen = $screen
            } elseif ($order -or $hasStop) {
                # 停止の時刻：OS の停止 → イベントログの停止 → 停止の開始 → 停止の指示
                if ($last13) { $row.Stop = $last13.Time } elseif ($last6006) { $row.Stop = $last6006.Time } elseif ($last109) { $row.Stop = $last109.Time } else { $row.Stop = $order.Time }
                if ($action) { $row.Kind = Get-ActionLabel $action } else { $row.Notes.Add('停止の記録はありますが、再起動かシャットダウンかが分かりません') }
                # 高速スタートアップでシャットダウンすると、OS の停止とイベントログの停止は記録されない（OS を完全には止めないため）
                if ($null -eq $last13 -and $null -eq $last6006 -and -not ($boot -and $boot.BootType -eq 1)) { $row.Notes.Add('停止の完了の記録がありません') }
                if ($order) {
                    $why = if ($order.Comment) { $order.Comment } else { $order.Reason }
                    $row.Screen = if ($why) { '{0}: {1}' -f $row.Trigger, $why } else { $row.Trigger }
                } elseif ($null -eq $action) {
                    $row.Screen = '再起動かシャットダウンかが分かりません'
                }
            } else {
                $row.Notes.Add('停止の記録がありません')
                $row.Screen = '停止の記録がありません'
            }
            if ($null -eq $boot) { $row.Notes.Add('このあとの起動の記録がありません') }
            if ($bootFromOs) { $row.Notes.Add('起動の記録が残っていないため、起動の時刻は OS の情報（最後の起動の時刻）から取りました') }

            # 停止していた時間とは別に、停止した時刻が分からないときは、停止の前の最後の記録を目安として書く
            if ($null -eq $row.Stop -and $boot) { $row.LastRecord = Get-LastRecordBefore $computer $boot.Time }
            $row.SortTime = if ($null -ne $row.Stop) { $row.Stop } else { $row.Boot }
            $rows.Add($row)
        }
    }

    #endregion

    #region 一覧にする

    function Test-InWindow($Time) {
        if ($null -eq $Time) { return $false }
        if ($sinceInfo -and $Time -lt $sinceInfo.Time) { return $false }
        if ($untilInfo -and $Time -gt $untilInfo.Time) { return $false }
        return $true
    }

    # 停止か起動のどちらかが期間に入るもの
    $periodRows = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $rows) {
        if ((Test-InWindow $row.Stop) -or (Test-InWindow $row.Boot)) { $periodRows.Add($row) }
    }
    $stops = @($periodRows.ToArray() | Sort-Object -Property @{ Expression = 'SortTime'; Descending = $true }, @{ Expression = { ConvertTo-SortKey $_.Computer }; Descending = $false }, @{ Expression = 'Seq'; Descending = $false })

    # events.csv には、期間の中の記録を書く（新しい順）
    $periodRecords = @($records | Where-Object { Test-InWindow $_.Time } | Sort-Object -Property @{ Expression = 'Time'; Descending = $true }, @{ Expression = { ConvertTo-SortKey $_.Computer }; Descending = $false }, @{ Expression = 'Seq'; Descending = $false })

    # 今の起動（この PC を調べたときだけ）。OS の情報の最後の起動の時刻と、記録のいちばん新しい起動の、新しいほう
    # （高速スタートアップの起動は、OS の情報の最後の起動の時刻を変えないため）
    $currentBoot = $null
    $currentInLog = $false
    if (-not $useFiles) {
        $currentBoot = $cimBoot
        $localKey = Get-ComputerKey $machineName
        if ($lastBoots.ContainsKey($localKey)) {
            $currentInLog = $true
            if ($null -eq $currentBoot -or $lastBoots[$localKey] -gt $currentBoot) { $currentBoot = $lastBoots[$localKey] }
        }
    }

    #endregion

    #region 画面と summary.txt の内容

    $readComputers = @($sources | ForEach-Object { $_.Computer } | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } } -Unique)
    $multiComputer = $readComputers.Count -gt 1
    function Get-ComputerPrefix([string]$Computer) { if ($multiComputer) { return '{0}: ' -f $Computer }; return '' }

    $head = [System.Collections.Generic.List[string]]::new()
    $head.Add((Format-Invariant '{0} ({1} {2})' $ToolTitle $ToolName $ScriptVersion))
    $oldestByComputer = @{}
    foreach ($source in $sources) {
        if ($null -eq $source.Oldest -or $source.Failed) { continue }
        if (-not $oldestByComputer.ContainsKey($source.Computer) -or $source.Oldest -lt $oldestByComputer[$source.Computer]) { $oldestByComputer[$source.Computer] = $source.Oldest }
    }
    $oldestText = ''
    if ($oldestByComputer.Count -eq 1) { $oldestText = Format-Invariant '   System ログ: {0:yyyy-MM-dd HH:mm} から残っています' (@($oldestByComputer.Values)[0]) }
    if ($useFiles) {
        # 読めたファイルの数（記録を 1 件も読めなかったファイルは数えない）
        $readSources = @($sources | Where-Object { -not $_.Failed })
        $more = if ($readSources.Count -gt 1) { Format-Invariant ' ほか {0:N0} ファイル' ($readSources.Count - 1) } else { '' }
        $head.Add(('対象: {0}{1}{2}' -f $readSources[0].Label, $more, $oldestText))
    } else {
        $head.Add(('対象: {0}{1}' -f $machineName, $oldestText))
    }
    if ($sinceInfo) {
        $endText = if ($untilInfo) { $untilInfo.Time } else { $startTime }
        $span = if ($sinceInfo.Relative -and -not $untilInfo) { '（{0}）' -f $sinceInfo.Relative } else { '' }
        $head.Add((Format-Invariant '期間: {0:yyyy-MM-dd HH:mm} ～ {1:yyyy-MM-dd HH:mm}{2}' $sinceInfo.Time $endText $span))
    } elseif ($untilInfo) {
        $head.Add((Format-Invariant '期間: ～ {0:yyyy-MM-dd HH:mm}' $untilInfo.Time))
    } else {
        $head.Add('期間: すべて（System ログに残っている記録）')
    }
    if ($null -ne $currentBoot) {
        $head.Add((Format-Invariant '今の起動: {0:yyyy-MM-dd HH:mm}（稼働 {1}）' $currentBoot (Format-Span (Measure-Span $currentBoot $startTime))))
    }
    $head.Add('')

    # 表の行
    $stopHead = [string[]]@('停止', '起動', '停止時間', '種類', 'ユーザー', 'きっかけ・理由')
    $stopRight = [bool[]]@($false, $false, $true, $false, $false, $false)
    if ($multiComputer) { $stopHead += 'コンピューター'; $stopRight += $false }
    $stopRows = @(foreach ($row in $stops) {
        $stopText = if ($null -ne $row.Stop) { Format-ScreenTime $row.Stop } else { '（不明）' }
        $downText = if ($null -ne $row.Stop -and $null -ne $row.Boot) { Format-Span (Measure-Span $row.Stop $row.Boot) } else { '' }
        $cells = [string[]]@($stopText, (Format-BootTime $row.Stop $row.Boot), $downText, $row.Kind, (Limit-Width $row.User $NameWidth), (Limit-Width $row.Screen $ReasonWidth))
        if ($multiComputer) { $cells += $row.Computer }
        , $cells
    })

    $tail = [System.Collections.Generic.List[string]]::new()
    $total = $stops.Count
    # 見出しの最後が空行なので、表があるときだけ、表の後に空行を入れる
    if ($total -gt 0) { $tail.Add('') }
    $timeText = '所要 {0}' -f (Format-Duration $sw.Elapsed)
    if ($total -eq 0) {
        $tail.Add(('該当する停止はありません  {0}' -f $timeText))
    } else {
        $parts = @()
        foreach ($kind in @('再起動', 'シャットダウン', '予期しない停止', '不明')) {
            $count = @($stops | Where-Object { $_.Kind -eq $kind }).Count
            if ($count -eq 0) { continue }
            $text = Format-Invariant '{0} {1:N0} 回' $kind $count
            $updates = @($stops | Where-Object { $_.Kind -eq $kind -and $_.WindowsUpdate }).Count
            if (($kind -eq '再起動' -or $kind -eq 'シャットダウン') -and $updates -gt 0) { $text += Format-Invariant '（うち Windows Update {0:N0} 回）' $updates }
            $parts += $text
        }
        $tail.Add(('{0}   {1}' -f ($parts -join '・'), $timeText))
    }
    # 注意
    foreach ($computer in @($oldestByComputer.Keys | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } })) {
        if ($sinceInfo -and $oldestByComputer[$computer] -gt $sinceInfo.Time) {
            $tail.Add((Format-Invariant '※ {1}{0:yyyy-MM-dd HH:mm} より前の記録は、System ログに残っていません（上書きされています）' $oldestByComputer[$computer] (Get-ComputerPrefix $computer)))
        }
    }
    if (-not $useFiles -and $null -ne $currentBoot -and -not $currentInLog -and $oldestByComputer.Count -gt 0 -and $currentBoot -lt @($oldestByComputer.Values)[0]) {
        $tail.Add((Format-Invariant '※ 今の起動（{0:yyyy-MM-dd HH:mm}）の記録は、System ログに残っていません。それより前の停止は分かりません' $currentBoot))
    }
    foreach ($order in @($pending | Sort-Object -Property Time, Seq)) {
        $label = Get-ActionLabel $order.Action
        if (-not $label) { $label = '停止' }
        $when = if ($useFiles) { 'ログを書き出した時点で、まだ' } else { 'まだ' }
        $who = if ($order.User) { '（{0}）' -f $order.User } else { '' }
        $tail.Add((Format-Invariant '※ {4}{0:yyyy-MM-dd HH:mm} に{1}の指示{2}がありますが、{3}{1}していません' $order.Time $label $who $when (Get-ComputerPrefix $order.Computer)))
    }
    foreach ($clear in @($records | Where-Object { $_.EventId -eq 104 -and (Test-InWindow $_.Time) } | Sort-Object -Property Time, Seq)) {
        $who = if ($clear.User) { '（{0}）' -f $clear.User } else { '' }
        $tail.Add((Format-Invariant '※ {2}{0:yyyy-MM-dd HH:mm} に System ログが消去されています{1}' $clear.Time $who (Get-ComputerPrefix $clear.Computer)))
    }
    $kinds = @(foreach ($name in $skipKinds.Keys) { if ($skipKinds[$name] -gt 0) { , @($name, $skipKinds[$name]) } })
    $skipTotal = 0
    foreach ($kind in $kinds) { $skipTotal += $kind[1] }
    if ($kinds.Count -eq 1) { $tail.Add((Format-Invariant 'スキップ {0:N0} 件（{1}）' $skipTotal $kinds[0][0])) }
    elseif ($kinds.Count -gt 1) { $tail.Add((Format-Invariant 'スキップ {0:N0} 件（{1}）' $skipTotal (($kinds | ForEach-Object { Format-Invariant '{0} {1:N0}' $_[0] $_[1] }) -join '・'))) }

    # 表（Limit が 0 なら、すべての行）
    function Get-Body([int]$Limit) {
        $lines = [System.Collections.Generic.List[string]]::new()
        $shown = @($stopRows)
        if ($Limit -gt 0 -and $shown.Count -gt $Limit) { $shown = @($shown[0..($Limit - 1)]) }
        $lines.Add(('■ 停止と起動（新しい順{0}）' -f $(if ($Limit -gt 0) { Format-Invariant '・{0} 件まで' $Limit } else { '' })))
        $lines.AddRange([string[]](Format-TextTable $stopHead $stopRight $shown))
        $rest = $stopRows.Count - $shown.Count
        if ($rest -gt 0) { $lines.Add((Format-Invariant '（ほか {0:N0} 件は reboots.csv を参照してください）' $rest)) }
        return , $lines
    }

    #endregion

    #region 保存

    if ($runDir) {
        Write-Progress -Activity $ToolTitle -Status '結果を保存しています'
        try {
            $summary = [System.Collections.Generic.List[string]]::new()
            $summary.AddRange([string[]]$head.ToArray())
            if ($total -gt 0) { $summary.AddRange([string[]](Get-Body 0).ToArray()) }
            $summary.AddRange([string[]]$tail.ToArray())
            $summary.Add(('保存先: {0}' -f $runDir))
            [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$summary, $Utf8Bom)

            $csvRows = [System.Collections.Generic.List[object]]::new()
            foreach ($row in $stops) {
                $downText = if ($null -ne $row.Stop -and $null -ne $row.Boot) { Format-Span (Measure-Span $row.Stop $row.Boot) } else { '' }
                $upText = if ($null -ne $row.Stop -and $null -ne $row.PrevBoot) { Format-Span (Measure-Span $row.PrevBoot $row.Stop) } else { '' }
                $csvRows.Add([object[]]@((Format-Time $row.Stop), (Format-Time $row.Boot), (Protect-ExcelText $downText), (Protect-ExcelText $upText), $row.Kind, $(if ($row.WindowsUpdate) { 'はい' } else { '' }),
                    (Protect-ExcelText $row.Trigger), (Protect-ExcelText $row.User), (Protect-ExcelText $row.Reason), (Protect-ExcelText $row.Code), (Protect-ExcelText $row.Comment),
                    $row.Cause, $row.Bugcheck, (Protect-ExcelText $row.Dump), (Protect-ExcelText $row.Later), (Format-Time $row.Ordered), (Protect-ExcelText $row.Process),
                    $row.BootType, (Protect-ExcelText $row.Version), (Format-Time $row.LastRecord), ($row.Notes -join '。'), (Protect-ExcelText $row.ComputerText)))
            }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'reboots.csv')) $RebootColumns $csvRows

            $csvRows = [System.Collections.Generic.List[object]]::new()
            foreach ($rec in $periodRecords) {
                $csvRows.Add([object[]]@((Format-Time $rec.Time), $Operations[$rec.EventId], (Protect-ExcelText $rec.Detail), (Protect-ExcelText $rec.User), (Protect-ExcelText $rec.Process),
                    $rec.EventId, $rec.Provider, (Protect-ExcelText $rec.ComputerText)))
            }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'events.csv')) $EventColumns $csvRows

            $csvRows = [System.Collections.Generic.List[object]]::new()
            foreach ($skip in $skipped) { $csvRows.Add([object[]]@($skip[0], (Protect-ExcelText $skip[1]), (Protect-ExcelText $skip[2]))) }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $csvRows
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
    Write-Progress -Activity $ToolTitle -Completed
    if (-not $finished -or $saveError -or $fatal) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        $runDir = $null
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

#region 画面表示

foreach ($line in $head) { Write-Host $line }
if ($total -gt 0) {
    # 保存できたときは、画面には先頭の分だけを出す。保存しなかったときは、ほかに見る場所がないためすべて出す
    $limit = if ($runDir) { $ScreenRows } else { 0 }
    foreach ($line in (Get-Body $limit)) {
        if ($line.StartsWith('■', [StringComparison]::Ordinal)) { Write-Host $line -ForegroundColor Cyan } else { Write-Host $line }
    }
}
foreach ($line in $tail) {
    if ($line.StartsWith('※', [StringComparison]::Ordinal)) { Write-Host $line -ForegroundColor Yellow } else { Write-Host $line }
}

$exitCode = if ($total -gt 0) { 0 } else { 2 }
if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
    $exitCode = 1
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
}

#endregion

exit $exitCode
