<#
.SYNOPSIS
    タスクスケジューラに登録されているタスクを一覧にして、CSV と XML に保存します。

.DESCRIPTION
    タスクスケジューラのタスクを、フォルダをたどってすべて読み、いつ（トリガー）・誰の権限で（実行ユーザー）・
    何を動かすか（操作）・前回の結果・次回の実行日時を画面に表示し、tasks.csv（BOM 付き UTF-8）に保存します。
    あわせて、各タスクを取り込み直せる XML を xml フォルダに保存します。

    タスクを変更・削除・実行することはありません。
    既定では、Windows 標準のタスク（\Microsoft\ の下）は出力しません。-All で含めます。
    権限がなくて読めないフォルダやタスクはスキップし、件数を表示します。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Export-TaskList.ps1

.PARAMETER Path
    出力するタスクのフォルダ（\集計 のように \ で始めて書きます）。下のフォルダもすべて出力します。
    省略すると、すべてのフォルダ（\）を出力します。

.PARAMETER All
    Windows 標準のタスク（\Microsoft\ の下）も出力します。

.PARAMETER NoXml
    各タスクの XML を保存しません。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Export-TaskList.ps1

    Windows 標準以外のタスクを出力します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Export-TaskList.ps1 -All

    Windows 標準のタスクも含めて、すべて出力します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Export-TaskList.ps1 -Path \集計 -NoXml

    \集計 フォルダの下のタスクの一覧だけを保存します。

.NOTES
    Export-TaskList 1.0.0
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
    [string]$Path = '\',

    [switch]$All,

    [switch]$NoXml,

    [string]$OutputFolder,

    [switch]$NoSave
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$ScreenRows    = 200      # 画面に表示するタスクの件数（保存するときだけ。超えた分はファイルを見てもらう）
$ProgressMs    = 250      # 進捗表示の間隔（ミリ秒）
$TaskColumns   = 'フォルダ', 'タスク名', '状態', '実行ユーザー', '実行の条件', '最上位の特権', 'トリガー', '操作', '作業フォルダ',
                 '前回の実行日時', '前回の結果', '結果のコード', '次回の実行日時', '作成者', '作成日', '説明', '隠し', 'XML ファイル'
$SkipColumns   = '理由', 'パス', '詳細'
$TimeFormat    = 'yyyy/MM/dd HH:mm:ss'
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$machineName   = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

# タスクスケジューラの値（Task Scheduler 2.0 の列挙値）
$StateNames   = @{ 0 = '不明'; 1 = '無効'; 2 = 'キューに登録'; 3 = '準備完了'; 4 = '実行中' }
$LogonNames   = @{ 1 = 'ログオンしているかどうかにかかわらず'; 2 = 'パスワードを保存しない'; 3 = 'ログオンしているときのみ';
                   4 = 'ログオンしているときのみ'; 5 = 'サービス アカウント'; 6 = 'ログオンしているときのみ' }
$DayNames     = '日', '月', '火', '水', '木', '金', '土'   # DaysOfWeek のビットの順（1 = 日曜）
$ResultRunning  = 0x00041301   # SCHED_S_TASK_RUNNING
$ResultReady    = 0x00041300   # SCHED_S_TASK_READY（まだ一度も実行していない）
$ResultNotYet   = 0x00041303   # SCHED_S_TASK_HAS_NOT_RUN

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

# COM の「アクセスが拒否されました」は、UnauthorizedAccessException か、HRESULT 0x80070005 の COMException で届く
function Test-AccessDenied([Exception]$Exception) {
    return $Exception -is [UnauthorizedAccessException] -or $Exception.HResult -eq -2147024891
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
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -OutputFolder "D:\引き継ぎ"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# タスクスケジューラの日時は、一度も実行していないときなどに 1899 年や 1999 年の値になるため、そのときは $null を返す
function Get-ValidTime($Value) {
    if ($null -eq $Value) { return $null }
    try { $time = [datetime]$Value } catch { return $null }
    if ($time.Year -lt 2000) { return $null }
    return $time
}

# トリガーの開始日時（StartBoundary、"2026-11-02T03:00:00" や "2026-11-02T03:00:00+09:00" の形）から、日付と時刻を取り出す
function Get-BoundaryPart([string]$Boundary, [bool]$WithDate) {
    $m = [regex]::Match($Boundary, '^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})')
    if (-not $m.Success) { return '' }
    if ($WithDate) { return $m.Groups[1].Value + ' ' + $m.Groups[2].Value }
    return $m.Groups[2].Value
}

# ISO 8601 の期間（PT1H、PT30M、P1D など）を「1 時間」「30 分」「1 日」の形にする
function Format-IsoDuration([string]$Text) {
    $m = [regex]::Match($Text, '^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$')
    if (-not $m.Success) { return $Text }
    $parts = @()
    if ($m.Groups[1].Success) { $parts += '{0} 日' -f $m.Groups[1].Value }
    if ($m.Groups[2].Success) { $parts += '{0} 時間' -f $m.Groups[2].Value }
    if ($m.Groups[3].Success) { $parts += '{0} 分' -f $m.Groups[3].Value }
    if ($m.Groups[4].Success) { $parts += '{0} 秒' -f $m.Groups[4].Value }
    return $parts -join ' '
}

# ビットの立っている位置（1 から数える）の一覧
function Get-BitList([long]$Mask, [int]$Count) {
    $list = @()
    for ($i = 0; $i -lt $Count; $i++) { if (($Mask -band ([long]1 -shl $i)) -ne 0) { $list += $i + 1 } }
    return $list
}

function Get-DayText([long]$Mask) {
    return ((Get-BitList $Mask 7) | ForEach-Object { $DayNames[$_ - 1] }) -join '・'
}

# 月の指定が「すべての月」でなければ、「（1・4・7・10 月）」の形で返す
function Get-MonthText([long]$Mask) {
    $months = @(Get-BitList $Mask 12)
    if ($months.Count -eq 0 -or $months.Count -eq 12) { return '' }
    return '（{0} 月）' -f ($months -join '・')
}

# SID で書かれた実行ユーザーは、名前に直す（直せなければそのまま）
function Get-AccountName([string]$Id) {
    if ($Id -notmatch '^S-1-') { return $Id }
    try { return (New-Object Security.Principal.SecurityIdentifier $Id).Translate([Security.Principal.NTAccount]).Value } catch { return $Id }
}

# トリガー 1 つを、「毎日 03:00」のような文章にする
function Get-TriggerText($Trigger) {
    $time = Get-BoundaryPart ([string]$Trigger.StartBoundary) $false
    switch ([int]$Trigger.Type) {
        0 { $text = 'イベント発生時' }
        1 { $text = '1 回 ' + (Get-BoundaryPart ([string]$Trigger.StartBoundary) $true) }
        2 {
            $interval = [int]$Trigger.DaysInterval
            $text = if ($interval -le 1) { '毎日 ' + $time } else { '{0} 日ごと {1}' -f $interval, $time }
        }
        3 {
            $interval = [int]$Trigger.WeeksInterval
            $head = if ($interval -le 1) { '毎週' } else { '{0} 週ごと' -f $interval }
            $text = '{0} {1} {2}' -f $head, (Get-DayText $Trigger.DaysOfWeek), $time
        }
        4 {
            $days = @(Get-BitList $Trigger.DaysOfMonth 31)
            $dayParts = @()
            if ($days.Count -gt 0) { $dayParts += ($days -join '・') + ' 日' }
            if ($Trigger.RunOnLastDayOfMonth) { $dayParts += '最終日' }
            $month = Get-MonthText $Trigger.MonthsOfYear
            $text = '毎月{0}{1} {2}' -f $(if ($month) { $month } else { ' ' }), ($dayParts -join '・'), $time
        }
        5 {
            $weeks = @(Get-BitList $Trigger.WeeksOfMonth 4 | ForEach-Object { '第 {0}' -f $_ })
            if ($Trigger.RunOnLastWeekOfMonth) { $weeks += '最終' }
            $month = Get-MonthText $Trigger.MonthsOfYear
            $text = '毎月{0}{1} {2}曜 {3}' -f $(if ($month) { $month } else { ' ' }), ($weeks -join '・'), (Get-DayText $Trigger.DaysOfWeek), $time
        }
        6 { $text = 'アイドル時' }
        7 { $text = 'タスクの作成・変更時' }
        8 { $text = '起動時' }
        9 {
            $user = [string]$Trigger.UserId
            $text = if ($user) { 'ログオン時（{0}）' -f (Get-AccountName $user) } else { 'ログオン時（すべてのユーザー）' }
        }
        11 { $text = 'セッションの状態の変更時' }
        default { $text = '種類 {0} のトリガー' -f $Trigger.Type }
    }
    $interval = [string]$Trigger.Repetition.Interval
    if ($interval) {
        $duration = [string]$Trigger.Repetition.Duration
        $span = if ($duration) { ' {0}繰り返す' -f (Format-IsoDuration $duration) } else { '無期限に繰り返す' }
        $text += '（{0}ごとに{1}）' -f (Format-IsoDuration $interval), $span
    }
    if (-not $Trigger.Enabled) { $text += '（無効）' }
    return $text
}

function Get-ActionText($Action) {
    switch ([int]$Action.Type) {
        0 {
            $text = [string]$Action.Path
            if ([string]$Action.Arguments) { $text += ' ' + [string]$Action.Arguments }
            return $text
        }
        5 { return 'COM ハンドラー' }
        6 { return 'メールの送信' }
        7 { return 'メッセージの表示' }
    }
    return '種類 {0} の操作' -f $Action.Type
}

# 前回の結果を「成功」「失敗」「実行中」「まだ実行していない」に分ける
function Get-ResultKind([long]$Code, $LastRun) {
    if ($Code -eq $ResultRunning) { return '実行中' }
    if ($Code -eq $ResultReady -or $Code -eq $ResultNotYet -or $null -eq $LastRun) { return 'まだ実行していない' }
    if ($Code -eq 0) { return '成功' }
    return '失敗'
}

# ファイル名に使えない文字を _ にする（タスク名やフォルダ名を、XML の保存先に使うため）
$invalidNameChars = [IO.Path]::GetInvalidFileNameChars()
function ConvertTo-SafeName([string]$Name) {
    $chars = $Name.ToCharArray()
    for ($i = 0; $i -lt $chars.Length; $i++) { if ($invalidNameChars -contains $chars[$i]) { $chars[$i] = '_' } }
    $safe = (New-Object string (, $chars)).TrimEnd('.', ' ')
    if (-not $safe) { $safe = '_' }
    return $safe
}

#endregion

#region 引数の確認

$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

# タスクのフォルダは \集計 の形。/ や末尾の \ も受け付ける
$taskRoot = (Remove-Quote $Path).Replace('/', '\')
if ($taskRoot.Contains('"')) { Stop-WithQuoteError '-Path' $taskRoot }
if (-not $taskRoot.StartsWith('\')) { $taskRoot = '\' + $taskRoot }
if ($taskRoot -ne '\') { $taskRoot = $taskRoot.TrimEnd('\') }
if (-not $taskRoot) { $taskRoot = '\' }

$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    $outText = Resolve-StartPath $outText
}

# \Microsoft\ の下は、-All のときと、-Path で \Microsoft\ の下を指定したときだけ出力する
$isMicrosoftPath = { param([string]$FolderPath) $FolderPath -eq '\Microsoft' -or $FolderPath.StartsWith('\Microsoft\', [StringComparison]::OrdinalIgnoreCase) }
$includeMicrosoft = [bool]$All -or (& $isMicrosoftPath $taskRoot)

try {
    $service = New-Object -ComObject Schedule.Service
    $service.Connect()
} catch {
    Stop-WithError ('タスクスケジューラに接続できません（{0}）' -f (Get-ErrorText (Get-InnerException $_.Exception)))
}
try {
    $rootFolder = $service.GetFolder($taskRoot)
} catch {
    Stop-WithError ('タスクのフォルダが見つかりません: {0}（{1}）' -f $taskRoot, (Get-ErrorText (Get-InnerException $_.Exception)))
}

$startTime = Get-Date

#endregion

#region 保存先の準備

# 読み終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir      = $null
$saveWarning = $null
if (-not $NoSave) {
    $baseDir = if ($outText) { $outText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        $runName = Format-Invariant 'TaskList_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
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

$tasks     = New-Object 'System.Collections.Generic.List[object]'
$skipped   = New-Object 'System.Collections.Generic.List[object]'
$skipKinds = [ordered]@{ '権限なし' = 0; 'その他のエラー' = 0 }

function Add-Skip([string]$ItemPath, [Exception]$Exception) {
    $ex = Get-InnerException $Exception
    if (Test-AccessDenied $ex) {
        $skipped.Add([object[]]@('権限なし', $ItemPath, 'アクセスが拒否されました'))
        $skipKinds['権限なし']++
    } else {
        $skipped.Add([object[]]@('その他のエラー', $ItemPath, (Get-ErrorText $ex)))
        $skipKinds['その他のエラー']++
    }
}

# 読んでいる途中や保存の途中で Ctrl+C されたときや、保存に失敗したときは、書きかけの結果を残さない
$finished   = $false
$unexpected = $false
$saveError  = $null
$sw         = [Diagnostics.Stopwatch]::StartNew()
$lines      = New-Object 'System.Collections.Generic.List[string]'
try {
    #region 読み取り

    # 再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる
    $stack = New-Object 'System.Collections.Generic.Stack[object]'
    $stack.Push($rootFolder)
    $nextProgress = 0L
    while ($stack.Count -gt 0) {
        $folder     = $stack.Pop()
        $folderPath = [string]$folder.Path
        if ($sw.ElapsedMilliseconds -ge $nextProgress) {
            Write-Progress -Activity 'タスクスケジューラ出力' -Status (Format-Invariant 'タスク {0:N0} 件' $tasks.Count) -CurrentOperation $folderPath
            $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
        }

        # 隠しタスクも含める（TASK_ENUM_HIDDEN = 1）
        $folderTasks = $null
        try { $folderTasks = @($folder.GetTasks(1)) } catch { Add-Skip $folderPath $_.Exception }
        if ($null -ne $folderTasks) {
            foreach ($task in $folderTasks) {
                $taskPath = $folderPath.TrimEnd('\') + '\' + [string]$task.Name
                try {
                    $definition = $task.Definition
                    $principal  = $definition.Principal
                    $userId     = [string]$principal.UserId
                    if (-not $userId) { $userId = [string]$principal.GroupId }
                    $lastRun    = Get-ValidTime $task.LastRunTime
                    $code       = [long]$task.LastTaskResult -band 0xFFFFFFFFL
                    $tasks.Add([pscustomobject]@{
                        Folder    = $folderPath
                        Name      = [string]$task.Name
                        Path      = $taskPath
                        State     = [int]$task.State
                        Enabled   = [bool]$task.Enabled
                        User      = Get-AccountName $userId
                        Logon     = $(if ($LogonNames.ContainsKey([int]$principal.LogonType)) { $LogonNames[[int]$principal.LogonType] } else { '' })
                        Highest   = [int]$principal.RunLevel -eq 1
                        Triggers  = (@(foreach ($trigger in $definition.Triggers) { Get-TriggerText $trigger }) -join ' / ')
                        Actions   = (@(foreach ($action in $definition.Actions) { Get-ActionText $action }) -join ' / ')
                        WorkDir   = (@(foreach ($action in $definition.Actions) { if ([int]$action.Type -eq 0 -and [string]$action.WorkingDirectory) { [string]$action.WorkingDirectory } }) -join ' / ')
                        LastRun   = $lastRun
                        Code      = $code
                        Result    = Get-ResultKind $code $lastRun
                        NextRun   = Get-ValidTime $task.NextRunTime
                        Author    = [string]$definition.RegistrationInfo.Author
                        Created   = ([string]$definition.RegistrationInfo.Date -replace '^(\d{4})-(\d{2})-(\d{2})T(\d{2}:\d{2}:\d{2}).*$', '$1/$2/$3 $4')
                        Comment   = [string]$definition.RegistrationInfo.Description
                        Hidden    = [bool]$definition.Settings.Hidden
                        Xml       = [string]$task.Xml
                        XmlFile   = ''
                    })
                } catch {
                    Add-Skip $taskPath $_.Exception
                }
            }
        }

        $subFolders = $null
        try { $subFolders = @($folder.GetFolders(0)) } catch { Add-Skip $folderPath $_.Exception }
        if ($null -ne $subFolders) {
            # スタックから取り出す順が名前の順になるよう、逆の順に積む
            for ($i = $subFolders.Count - 1; $i -ge 0; $i--) {
                $sub = $subFolders[$i]
                if (-not $includeMicrosoft -and (& $isMicrosoftPath ([string]$sub.Path))) { continue }
                $stack.Push($sub)
            }
        }
    }
    $elapsed = $sw.Elapsed
    Write-Progress -Activity 'タスクスケジューラ出力' -Completed

    # フォルダとタスク名の順に並べる。オブジェクトの配列を [Array]::Sort に渡すと、PowerShell が別の配列に変えて渡すため
    # 並べ替わらない。パスの配列を並べ替えて引き直す（\ をどの文字よりも前に並べ、フォルダの中身がすぐ後に来るようにする）
    $byKey = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($task in $tasks) { $byKey[$task.Path.Replace('\', [char]1)] = $task }
    $keys = New-Object 'string[]' $byKey.Count
    $byKey.Keys.CopyTo($keys, 0)
    [Array]::Sort($keys, [StringComparer]::OrdinalIgnoreCase)
    $sorted = @(foreach ($key in $keys) { $byKey[$key] })

    #endregion

    #region 画面と summary.txt の内容

    $enabledCount = 0
    $failCount    = 0
    foreach ($task in $sorted) {
        if ($task.Enabled) { $enabledCount++ }
        # 無効のタスクの、昔の失敗は数えない
        if ($task.Enabled -and $task.Result -eq '失敗') { $failCount++ }
    }

    $lines.Add((Format-Invariant 'タスクスケジューラ出力 (Export-TaskList {0})' $ScriptVersion))
    $lines.Add((Format-Invariant '対象: {0}（管理者: {1}）   {2:yyyy-MM-dd HH:mm:ss}' $machineName $(if ($isAdmin) { 'はい' } else { 'いいえ' }) $startTime))
    if ($includeMicrosoft -and -not (& $isMicrosoftPath $taskRoot)) {
        $lines.Add(('範囲: {0} の下（Windows 標準のタスクを含む）' -f $taskRoot))
    } elseif ($includeMicrosoft) {
        $lines.Add(('範囲: {0} の下' -f $taskRoot))
    } else {
        $lines.Add(('範囲: {0} の下（\Microsoft\ の下を除く）' -f $taskRoot))
    }
    if ($NoXml) { $lines.Add('XML: 保存しない') }
    $lines.Add('')
    $headIndex = $lines.Count

    $table = New-Object 'System.Collections.Generic.List[object]'   # (文字, 色)
    $widthState = 4; $widthResult = 10; $widthUser = 12
    $cells = @(foreach ($task in $sorted) {
        $resultText = if (-not $task.Enabled) { '-' } elseif ($task.Result -eq '成功' -or $task.Result -eq '失敗') { '{0} (0x{1:X})' -f $task.Result, $task.Code } else { $task.Result }
        $nextText   = if ($task.NextRun -and $task.Enabled) { $task.NextRun.ToString('yyyy-MM-dd HH:mm', $Invariant) } else { '-' }
        , @($StateNames[$task.State], $resultText, $nextText, $task.User, $task.Path, $task)
    })
    foreach ($cell in $cells) {
        $widthState  = [math]::Max($widthState, (Get-DisplayWidth $cell[0]))
        $widthResult = [math]::Max($widthResult, (Get-DisplayWidth $cell[1]))
        $widthUser   = [math]::Max($widthUser, (Get-DisplayWidth $cell[3]))
    }
    $tableHeader = ' ' + (Format-Column '状態' $widthState) + '  ' + (Format-Column '前回の結果' $widthResult) + '  ' + (Format-Column '次回の実行' 16) + '  ' + (Format-Column '実行ユーザー' $widthUser) + '  タスク'
    foreach ($cell in $cells) {
        $text  = ' ' + (Format-Column $cell[0] $widthState) + '  ' + (Format-Column $cell[1] $widthResult) + '  ' + (Format-Column $cell[2] 16) + '  ' + (Format-Column $cell[3] $widthUser) + '  ' + $cell[4]
        $color = if (-not $cell[5].Enabled) { 'DarkGray' } elseif ($cell[5].Result -eq '失敗') { 'Yellow' } else { '' }
        $table.Add(@($text, $color))
    }

    if ($table.Count -eq 0) {
        $lines.Add('    （タスクはありません）')
    } else {
        $lines.Add($tableHeader)
        foreach ($entry in $table) { $lines.Add($entry[0]) }
    }
    $tail = New-Object 'System.Collections.Generic.List[string]'
    $tail.Add('')
    $countLine = Format-Invariant 'タスク {0:N0} 件（有効 {1:N0}・無効 {2:N0}）' $sorted.Count $enabledCount ($sorted.Count - $enabledCount)
    if ($failCount -gt 0) { $countLine += Format-Invariant '  前回失敗 {0:N0} 件' $failCount }
    $countLine += '  所要 ' + (Format-Duration $elapsed)
    $tail.Add($countLine)
    $kinds = @(foreach ($name in $skipKinds.Keys) { if ($skipKinds[$name] -gt 0) { , @($name, $skipKinds[$name]) } })
    if ($kinds.Count -eq 1) {
        $tail.Add((Format-Invariant 'スキップ {0:N0} 件（{1}）' $kinds[0][1] $kinds[0][0]))
    } elseif ($kinds.Count -gt 1) {
        $tail.Add((Format-Invariant 'スキップ {0:N0} 件（権限なし {1:N0}・その他のエラー {2:N0}）' ($skipKinds['権限なし'] + $skipKinds['その他のエラー']) $skipKinds['権限なし'] $skipKinds['その他のエラー']))
    }
    if (-not $isAdmin) { $tail.Add('※ 管理者として実行していないため、ほかのユーザーや SYSTEM のタスクが出ないことがあります') }
    foreach ($line in $tail) { $lines.Add($line) }

    #endregion

    #region 保存

    if ($runDir) {
        Write-Progress -Activity 'タスクスケジューラ出力' -Status '結果を保存しています'
        try {
            if (-not $NoXml) {
                # タスクスケジューラと同じフォルダの構成で、タスクスケジューラが書き出すときと同じ UTF-16（BOM 付き）で保存する
                foreach ($task in $sorted) {
                    $parts = @($task.Folder.Split([char[]]@('\'), [StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { ConvertTo-SafeName $_ })
                    $relative = (@($parts) + (ConvertTo-SafeName $task.Name)) -join '\'
                    $relative += '.xml'
                    $xmlPath = Join-Path (Join-Path $runDir 'xml') $relative
                    [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath (Split-Path -LiteralPath $xmlPath)))
                    [IO.File]::WriteAllText((ConvertTo-LongPath $xmlPath), $task.Xml, [Text.Encoding]::Unicode)
                    $task.XmlFile = $relative
                }
            }
            $rows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($task in $sorted) {
                $rows.Add([object[]]@(
                    $task.Folder, (Protect-ExcelText $task.Name), $StateNames[$task.State], (Protect-ExcelText $task.User), $task.Logon,
                    $(if ($task.Highest) { 'はい' } else { 'いいえ' }), $task.Triggers, (Protect-ExcelText $task.Actions), (Protect-ExcelText $task.WorkDir),
                    $(if ($task.LastRun) { $task.LastRun.ToString($TimeFormat, $Invariant) } else { '' }), $task.Result, ('0x{0:X8}' -f $task.Code),
                    $(if ($task.NextRun) { $task.NextRun.ToString($TimeFormat, $Invariant) } else { '' }),
                    (Protect-ExcelText $task.Author), $task.Created, (Protect-ExcelText $task.Comment), $(if ($task.Hidden) { 'はい' } else { '' }), $task.XmlFile))
            }
            $lines.Add(('保存先: {0}' -f $runDir))
            [IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), [string[]]$lines, $Utf8Bom)
            Save-Csv (Join-Path $runDir 'tasks.csv') $TaskColumns $rows
            Save-Csv (Join-Path $runDir 'skipped.csv') $SkipColumns $skipped
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
        }
        Write-Progress -Activity 'タスクスケジューラ出力' -Completed
    }

    #endregion

    $finished = $true
} catch {
    # 想定外のエラーは、Ctrl+C の「中断しました」と区別して知らせる
    $unexpected = $true
    throw
} finally {
    Write-Progress -Activity 'タスクスケジューラ出力' -Completed
    if (-not $finished -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

#region 画面表示

$saved = $runDir -and -not $saveError
for ($i = 0; $i -lt $headIndex; $i++) { Write-Host $lines[$i] }
if ($table.Count -eq 0) {
    Write-Host '    （タスクはありません）'
} else {
    Write-Host $tableHeader
    # 保存できたときは、画面には先頭の分だけを出す。保存しなかったときは、ほかに見る場所がないためすべて出す
    $limit = if ($saved) { $ScreenRows } else { $table.Count }
    $shown = 0
    foreach ($entry in $table) {
        if ($shown -ge $limit) { break }
        if ($entry[1]) { Write-Host $entry[0] -ForegroundColor $entry[1] } else { Write-Host $entry[0] }
        $shown++
    }
    if ($table.Count -gt $shown) {
        Write-Host (Format-Invariant '（ほか {0:N0} 件は tasks.csv を参照してください）' ($table.Count - $shown))
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

exit 0
