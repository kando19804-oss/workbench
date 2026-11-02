<#
.SYNOPSIS
    サーバーの状態を保存しておき、後で比べて、何が変わったかを一覧にします。

.DESCRIPTION
    サービス、待ち受けポート、タスク、インストール済みソフト、更新プログラム、ネットワーク設定、ルーティング、
    共有フォルダ、ローカルユーザー、ローカルグループのメンバー、hosts ファイル、システムの環境変数、ファイアウォールの
    状態を保存します（スナップショット）。保存したものと今の状態、または保存したもの同士を比べ、追加・削除・変更を
    画面に表示し、CSV（BOM 付き UTF-8）にも保存します。

    読み取りだけを行い、サーバーの設定を変更することはありません。
    タスクの操作、サービスの実行ファイル、環境変数の値はパスワードなどが含まれることがあるため、
    画面には表示せず CSV にだけ書きます。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Compare-SystemState.ps1 -Snapshot

.PARAMETER Snapshot
    今の状態を保存するだけで、比べません。

.PARAMETER Reference
    基準のスナップショットのフォルダ。-Snapshot のとき以外は必須です。

.PARAMETER Difference
    比較のスナップショットのフォルダ。省略すると、今の状態と比べます。

.PARAMETER Item
    調べる項目。"Service,Port" のようにカンマで区切って書きます。省略すると、すべての項目です（比べるときは、基準のスナップショットにある項目すべて）。
    Service、Port、Task、Software、HotFix、Network、Route、Share、LocalUser、LocalGroup、Hosts、Environment、Firewall

.PARAMETER AllTasks
    Windows 標準のタスク（\Microsoft\ の下）も保存します。-Snapshot のときだけ指定できます。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。-Snapshot と一緒には使えません。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Compare-SystemState.ps1 -Snapshot

    作業の前に、今の状態を保存します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Compare-SystemState.ps1 -Reference .\SystemSnapshot_SERVER01_20261103_220000

    作業の後に、保存しておいた状態と今の状態を比べます。

.NOTES
    Compare-SystemState 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [switch]$Snapshot,

    [string]$Reference,

    [string]$Difference,

    [string]$Item,

    [switch]$AllTasks,

    [string]$OutputFolder,

    [switch]$NoSave
)

# 制約付き言語モード（AppLocker や WDAC の環境）では、このツールが使う .NET の型や COM を使えないので、最初に止める
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host '制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）'
    exit 1
}

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$ToolName      = 'Compare-SystemState'
$ToolTitle     = '作業前後スナップショット比較'
$ScreenRows    = 200      # 画面に表示する違いの件数（保存するときだけ。超えた分はファイルを見てもらう）
$DynamicPort   = 49152    # Windows が動的に割り当てるポートの範囲の始まり
$DiffColumns   = '結果', '項目', '対象', '属性', '基準の値', '比較の値', '備考'
$SkipColumns   = '理由', '項目', '詳細'
$TimeFormat    = 'yyyy/MM/dd HH:mm:ss'
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$OnWindows     = [IO.Path]::DirectorySeparatorChar -eq '\'   # 開発用に Mac の pwsh でも動かせるようにするため
$machineName   = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

# 項目の定義。Keys は突き合わせる列、Compare は比べる列（空ならあるかないかだけ）、Hidden は画面に値を出さない列
$ItemDefs = [ordered]@{
    Service     = @{ Label = 'サービス'; File = 'services.csv'; Columns = 'サービス名', '表示名', '状態', 'スタートアップ', '実行ユーザー', '実行ファイル'
                     Keys = @('サービス名'); Compare = @('スタートアップ', '状態', '実行ユーザー', '実行ファイル'); Hidden = @('実行ファイル') }
    Port        = @{ Label = '待ち受けポート'; File = 'ports.csv'; Columns = 'プロトコル', 'ローカルアドレス', 'ポート', 'プロセス名', 'サービス'
                     Keys = @('プロトコル', 'ローカルアドレス', 'ポート'); Compare = @('プロセス名', 'サービス'); Hidden = @() }
    Task        = @{ Label = 'タスク'; File = 'tasks.csv'; Columns = 'パス', '有効', '実行ユーザー', 'トリガー', '操作'
                     Keys = @('パス'); Compare = @('有効', '実行ユーザー', 'トリガー', '操作'); Hidden = @('操作') }
    Software    = @{ Label = 'インストール済みソフト'; File = 'software.csv'; Columns = '名前', 'ビット', 'バージョン', '発行元'
                     Keys = @('名前', 'ビット'); Compare = @('バージョン'); Hidden = @() }
    HotFix      = @{ Label = '更新プログラム'; File = 'hotfixes.csv'; Columns = 'KB 番号', '種類', 'インストール日'
                     Keys = @('KB 番号'); Compare = @(); Hidden = @() }
    Network     = @{ Label = 'ネットワーク設定'; File = 'network.csv'; Columns = '接続の名前', 'アダプター', 'MAC アドレス', 'DHCP', 'IP アドレス', 'サブネット', 'ゲートウェイ', 'DNS サーバー'
                     Keys = @('接続の名前'); Compare = @('IP アドレス', 'サブネット', 'ゲートウェイ', 'DNS サーバー', 'DHCP', 'MAC アドレス'); Hidden = @() }
    Route       = @{ Label = 'ルーティング'; File = 'routes.csv'; Columns = '宛先', 'ゲートウェイ', '接続の名前', 'メトリック'
                     Keys = @('宛先', 'ゲートウェイ', '接続の名前'); Compare = @('メトリック'); Hidden = @() }
    Share       = @{ Label = '共有フォルダ'; File = 'shares.csv'; Columns = '共有名', 'パス', '説明'
                     Keys = @('共有名'); Compare = @('パス'); Hidden = @() }
    LocalUser   = @{ Label = 'ローカルユーザー'; File = 'localusers.csv'; Columns = 'ユーザー名', '有効', 'パスワードの期限', '説明'
                     Keys = @('ユーザー名'); Compare = @('有効', 'パスワードの期限'); Hidden = @() }
    LocalGroup  = @{ Label = 'ローカルグループのメンバー'; File = 'localgroups.csv'; Columns = 'グループ名', 'メンバー名'
                     Keys = @('グループ名', 'メンバー名'); Compare = @(); Hidden = @() }
    Hosts       = @{ Label = 'hosts ファイル'; File = 'hosts.csv'; Columns = 'アドレス', 'ホスト名'
                     Keys = @('アドレス', 'ホスト名'); Compare = @(); Hidden = @() }
    Environment = @{ Label = '環境変数'; File = 'environment.csv'; Columns = '変数名', '値'
                     Keys = @('変数名'); Compare = @('値'); Hidden = @('値') }
    Firewall    = @{ Label = 'ファイアウォール'; File = 'firewall.csv'; Columns = '種類', 'ID', '表示名', '有効', '方向', '操作', 'プロファイル', 'プロトコル', 'ローカルポート', 'リモートポート', 'プログラム', 'リモートアドレス'
                     Keys = @('種類', 'ID'); Compare = @('有効', '方向', '操作', 'プロファイル', 'プロトコル', 'ローカルポート', 'リモートポート', 'プログラム', 'リモートアドレス'); Hidden = @() }
}
# 列の位置を引けるようにしておく
foreach ($def in $ItemDefs.Values) {
    $def.Index = @{}
    for ($i = 0; $i -lt $def.Columns.Count; $i++) { $def.Index[$def.Columns[$i]] = $i }
    $def.KeyIndex = @(foreach ($column in $def.Keys) { $def.Index[$column] })
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

# .NET や COM のメソッドで起きた例外は MethodInvocationException に、プロパティ（タスクの Definition など）を読むときに
# 起きた例外は GetValueInvocationException に包まれて届くため、元の例外を取り出す（どちらも ExtendedTypeSystemException の仲間）
function Get-InnerException([Exception]$Exception) {
    while ($Exception -is [Management.Automation.ExtendedTypeSystemException] -and $Exception.InnerException) {
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

# 「アクセスが拒否されました」は、UnauthorizedAccessException か、HRESULT 0x80070005 の COMException か、CimException で届く
function Test-AccessDenied([Exception]$Exception) {
    if ($Exception -is [UnauthorizedAccessException] -or $Exception.HResult -eq -2147024891) { return $true }
    if ($Exception.GetType().Name -eq 'CimException') { return "$($Exception.NativeErrorCode)" -eq 'AccessDenied' }
    return $false
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
# 数値や日付に変えてしまうため、その形の値は ="..." で囲む。スナップショットは読み戻すため、長さにかかわらず ="..." にする
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
    return '="' + $Text.Replace('"', '""') + '"'
}

# Protect-ExcelText の逆。比べるときに元の値に戻す
function Restore-ExcelText([string]$Text) {
    if ($Text.StartsWith('="', [StringComparison]::Ordinal) -and $Text.EndsWith('"', [StringComparison]::Ordinal) -and $Text.Length -ge 3) {
        return $Text.Substring(2, $Text.Length - 3).Replace('""', '"')
    }
    return $Text
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

# 見出しと行（文字列の配列）から、列をそろえた表の行を作る
function Format-Table([string[]]$Head, [bool[]]$Right, $Rows) {
    $widths = @(foreach ($i in 0..($Head.Count - 1)) {
        $w = Get-DisplayWidth $Head[$i]
        foreach ($row in $Rows) { $w = [math]::Max($w, (Get-DisplayWidth $row[$i])) }
        $w
    })
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($row in @(, $Head) + @($Rows)) {
        $cells = @(foreach ($i in 0..($Head.Count - 1)) { Format-Cell $row[$i] $widths[$i] $Right[$i] })
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
function Stop-WithQuoteError([string]$Name, [string]$Value) {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -Reference "C:\tools\SystemSnapshot_1"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# 結果の表示用に、引数を必要なときだけ引用符で囲む。"D:\data\" のように最後が \" になると
# 引用符が閉じなくなるため、最後の \ は 2 つにする
function Format-Argument([string]$Text) {
    if ($Text -notmatch "[\s&()^,;=!%'`$]") { return $Text }
    # "..." の中の $名前 は PowerShell では変数として展開されるため、' を含まなければ '...' で囲む（コマンドプロンプトでは ' が残るので、$ があるときだけ）
    if ($Text.Contains('$') -and -not $Text.Contains("'")) { return "'" + $Text + "'" }
    if ($Text.EndsWith('\')) { $Text += '\' }
    return '"' + $Text + '"'
}

# 5.1 の Get-Content は BOM のない UTF-8 を Shift_JIS として読んでしまうため、文字コードを自分で判定する。
# BOM があればそれに従い、なければ UTF-8 として読めるかを確かめ、読めなければ Shift_JIS とみなす
function Read-TextAuto([string]$FilePath) {
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return (New-Object Text.UTF8Encoding $false).GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2) }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return [Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2) }
    try { return (New-Object Text.UTF8Encoding $false, $true).GetString($bytes) } catch { }
    return [Text.Encoding]::GetEncoding(932).GetString($bytes)
}

# 値を 1 行の文字列にそろえる（CSV を 1 行ずつ読み戻すため、途中の改行を残さない）
function ConvertTo-CellText($Value) {
    if ($null -eq $Value) { return '' }
    return (([string]$Value) -replace '\s*[\r\n]+\s*', ' ').Trim()
}

# 配列の値を、並べ替えて 1 つの文字列にする。順序に意味がないものに使う（空の値は除く）
function Join-Sorted($Values) {
    $list = New-Object 'System.Collections.Generic.List[string]'
    foreach ($value in @($Values)) {
        $text = ConvertTo-CellText $value
        if ($text -ne '' -and -not $list.Contains($text)) { $list.Add($text) }
    }
    $names = $list.ToArray()
    [Array]::Sort($names, [StringComparer]::OrdinalIgnoreCase)
    return $names -join ', '
}

# 順序に意味があるものは、そのままつなぐ
function Join-Ordered($Values) {
    return (@(foreach ($value in @($Values)) { $text = ConvertTo-CellText $value; if ($text -ne '') { $text } }) -join ', ')
}

# 列の順に値を並べた行を作って加える
function Add-Row($Rows) {
    $Rows.Add([string[]]@(foreach ($value in $args) { ConvertTo-CellText $value }))
}

# プロパティがないかもしれない値を読む（OS や PowerShell の版で、ないことがある）
function Get-OptionalValue($Object, [string]$Name) {
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

#endregion

#region タスクの文章（Export-TaskList と同じ書き方）

$DayNames = '日', '月', '火', '水', '木', '金', '土'   # DaysOfWeek のビットの順（1 = 日曜）

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

#endregion

#region 項目ごとの取得

# 取得の途中で一部だけ読めなかったものを残す。$Skips は項目を読む間だけ差し替える
$currentSkips = $null
function Add-PartialSkip([string]$ItemPath, [Exception]$Exception) {
    $ex = Get-InnerException $Exception
    $reason = if (Test-AccessDenied $ex) { '権限なし' } else { 'その他のエラー' }
    $currentSkips.Add([object[]]@($reason, ('{0}: {1}' -f $ItemPath, (Get-ErrorText $ex))))
}

# Windows の上でだけ開けるレジストリ。ビューを指定して HKLM を開く
function Open-MachineKey([string]$SubKey, [string]$View) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::$View)
    return $base.OpenSubKey($SubKey)
}

# ユーザーごとのサービスの種類を、レジストリの Type から調べる。
# SERVICE_USER_SERVICE（0x40）はユーザーごとのサービス、SERVICE_USERSERVICE_INSTANCE（0x80）はログオンごとの実体（末尾が変わる）
function Get-ServiceTypeBits([string]$Name) {
    try {
        $key = Open-MachineKey ('SYSTEM\CurrentControlSet\Services\' + $Name) 'Default'
        if ($key) {
            try { return [int]$key.GetValue('Type', 0) } finally { $key.Dispose() }
        }
    } catch { }
    return 0
}

function Read-ItemService($Rows) {
    $stateNames = @{ 'Running' = '実行中'; 'Stopped' = '停止'; 'Paused' = '一時停止'; 'Start Pending' = '開始中'; 'Stop Pending' = '停止中'; 'Pause Pending' = '一時停止中'; 'Continue Pending' = '再開中' }
    $modeNames  = @{ 'Auto' = '自動'; 'Manual' = '手動'; 'Disabled' = '無効'; 'Boot' = 'ブート'; 'System' = 'システム' }
    foreach ($s in @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop)) {
        $name  = [string]$s.Name
        $state = [string]$s.State
        if ($stateNames.ContainsKey($state)) { $state = $stateNames[$state] }
        $mode = [string]$s.StartMode
        if ($modeNames.ContainsKey($mode)) { $mode = $modeNames[$mode] }
        if ($mode -eq '自動' -and (Get-OptionalValue $s 'DelayedAutoStart')) { $mode = '自動（遅延）' }
        $user = [string]$s.StartName
        # ユーザーごとのサービスは、ログオンのたびに名前の末尾と状態が変わるので、末尾を除いた名前にまとめ、状態と実行ユーザーは比べない
        $bits = Get-ServiceTypeBits $name
        if (($bits -band 0x40) -ne 0) {
            if (($bits -band 0x80) -ne 0) { $name = $name -replace '_[0-9a-fA-F]+$', '' }
            $state = ''
            $user  = ''
        }
        Add-Row $Rows $name $s.DisplayName $state $mode $user $s.PathName
    }
}

function Read-ItemPort($Rows) {
    $processNames = @{}
    foreach ($p in @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop)) { $processNames[[int]$p.ProcessId] = [string]$p.Name }
    $serviceNames = @{}
    foreach ($s in @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop)) {
        $serviceId = [int]$s.ProcessId
        if ($serviceId -eq 0) { continue }
        if (-not $serviceNames.ContainsKey($serviceId)) { $serviceNames[$serviceId] = New-Object 'System.Collections.Generic.List[string]' }
        $serviceNames[$serviceId].Add([string]$s.Name)
    }
    # 同じアドレスとポートを複数のプロセスが使うことがあるので、1 行にまとめる
    $merged = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
    $addEntry = {
        param([string]$Protocol, [string]$Address, [int]$Number, [int]$OwnerId)
        $key = '{0}|{1}|{2}' -f $Protocol, $Address, $Number
        if (-not $merged.ContainsKey($key)) {
            $merged[$key] = [pscustomobject]@{
                Protocol = $Protocol; Address = $Address; Number = $Number
                Names = New-Object 'System.Collections.Generic.List[string]'
                Services = New-Object 'System.Collections.Generic.List[string]'
            }
        }
        $entry = $merged[$key]
        if ($processNames.ContainsKey($OwnerId)) { $entry.Names.Add($processNames[$OwnerId]) }
        if ($serviceNames.ContainsKey($OwnerId)) { $entry.Services.AddRange($serviceNames[$OwnerId]) }
    }
    foreach ($c in @(Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue)) { & $addEntry 'TCP' ([string]$c.LocalAddress) ([int]$c.LocalPort) ([int]$c.OwningProcess) }
    foreach ($u in @(Get-NetUDPEndpoint -ErrorAction Stop)) { & $addEntry 'UDP' ([string]$u.LocalAddress) ([int]$u.LocalPort) ([int]$u.OwningProcess) }
    foreach ($entry in $merged.Values) {
        Add-Row $Rows $entry.Protocol $entry.Address $entry.Number (Join-Sorted $entry.Names.ToArray()) (Join-Sorted $entry.Services.ToArray())
    }
}

function Read-ItemTask($Rows) {
    $service = New-Object -ComObject Schedule.Service
    $service.Connect()
    $stack = New-Object 'System.Collections.Generic.Stack[object]'
    $stack.Push($service.GetFolder('\'))
    while ($stack.Count -gt 0) {
        $folder     = $stack.Pop()
        $folderPath = [string]$folder.Path
        # 隠しタスクも含める（TASK_ENUM_HIDDEN = 1）
        $folderTasks = $null
        try { $folderTasks = @($folder.GetTasks(1)) } catch { Add-PartialSkip $folderPath $_.Exception }
        if ($null -ne $folderTasks) {
            foreach ($task in $folderTasks) {
                $taskPath = $folderPath.TrimEnd('\') + '\' + [string]$task.Name
                try {
                    $definition = $task.Definition
                    $userId = [string]$definition.Principal.UserId
                    if (-not $userId) { $userId = [string]$definition.Principal.GroupId }
                    Add-Row $Rows $taskPath $(if ($task.Enabled) { '有効' } else { '無効' }) (Get-AccountName $userId) `
                        (@(foreach ($trigger in $definition.Triggers) { Get-TriggerText $trigger }) -join ' / ') `
                        (@(foreach ($action in $definition.Actions) { Get-ActionText $action }) -join ' / ')
                } catch {
                    Add-PartialSkip $taskPath $_.Exception
                }
            }
        }
        $subFolders = $null
        try { $subFolders = @($folder.GetFolders(0)) } catch { Add-PartialSkip $folderPath $_.Exception }
        if ($null -ne $subFolders) {
            foreach ($sub in $subFolders) {
                $subPath = [string]$sub.Path
                if (-not $script:taskScopeAll -and ($subPath -eq '\Microsoft' -or $subPath.StartsWith('\Microsoft\', [StringComparison]::OrdinalIgnoreCase))) { continue }
                $stack.Push($sub)
            }
        }
    }
}

function Read-ItemSoftware($Rows) {
    $views = if ([Environment]::Is64BitOperatingSystem) { @(@('Registry64', '64 ビット'), @('Registry32', '32 ビット')) } else { @(, @('Registry32', '32 ビット')) }
    foreach ($view in $views) {
        $root = Open-MachineKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' $view[0]
        if (-not $root) { continue }
        $found = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
        try {
            foreach ($subName in $root.GetSubKeyNames()) {
                $key = $null
                try {
                    $key = $root.OpenSubKey($subName)
                    if (-not $key) { continue }
                    $displayName = [string]$key.GetValue('DisplayName', '')
                    # 「プログラムと機能」に出ないもの（システムの部品と、別のソフトの更新）は除く
                    if (-not $displayName -or [int]$key.GetValue('SystemComponent', 0) -eq 1 -or [string]$key.GetValue('ParentKeyName', '')) { continue }
                    if (-not $found.ContainsKey($displayName)) {
                        $found[$displayName] = [pscustomobject]@{ Versions = New-Object 'System.Collections.Generic.List[string]'; Publisher = [string]$key.GetValue('Publisher', '') }
                    }
                    $found[$displayName].Versions.Add([string]$key.GetValue('DisplayVersion', ''))
                } catch {
                    Add-PartialSkip $subName $_.Exception
                } finally {
                    if ($key) { $key.Dispose() }
                }
            }
        } finally {
            $root.Dispose()
        }
        foreach ($displayName in $found.Keys) {
            Add-Row $Rows $displayName $view[1] (Join-Sorted $found[$displayName].Versions.ToArray()) $found[$displayName].Publisher
        }
    }
}

function Read-ItemHotFix($Rows) {
    foreach ($h in @(Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop)) {
        Add-Row $Rows $h.HotFixID $h.Description $h.InstalledOn
    }
}

function Read-ItemNetwork($Rows) {
    $connections = @{}
    foreach ($a in @(Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction Stop)) { $connections[[int]$a.Index] = [string]$a.NetConnectionID }
    foreach ($n in @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE' -ErrorAction Stop)) {
        $name = [string]$n.Description
        if ($connections.ContainsKey([int]$n.Index) -and $connections[[int]$n.Index]) { $name = $connections[[int]$n.Index] }
        # IP アドレスとサブネットは同じ順に並んでいるので、組にして IP アドレスの順に並べる
        $pairs = New-Object 'System.Collections.Generic.List[string]'
        $addresses = @($n.IPAddress)
        $subnets   = @($n.IPSubnet)
        for ($i = 0; $i -lt $addresses.Count; $i++) {
            if ($null -eq $addresses[$i]) { continue }
            $pairs.Add(([string]$addresses[$i]) + '|' + $(if ($i -lt $subnets.Count) { [string]$subnets[$i] } else { '' }))
        }
        $pairArray = $pairs.ToArray()
        [Array]::Sort($pairArray, [StringComparer]::OrdinalIgnoreCase)
        $ipText     = (@(foreach ($pair in $pairArray) { $pair.Split('|')[0] }) -join ', ')
        $subnetText = (@(foreach ($pair in $pairArray) { $pair.Split('|')[1] }) -join ', ')
        Add-Row $Rows $name $n.Description $n.MACAddress $(if ($n.DHCPEnabled) { '有効' } else { '無効' }) $ipText $subnetText (Join-Sorted $n.DefaultIPGateway) (Join-Ordered $n.DNSServerSearchOrder)
    }
}

function Read-ItemRoute($Rows) {
    foreach ($r in @(Get-NetRoute -PolicyStore ActiveStore -ErrorAction Stop)) {
        $prefix = [string]$r.DestinationPrefix
        # マルチキャスト、リンクローカル、ループバック、ブロードキャストは、接続するたびに変わるので取らない
        if ($prefix -match '^(22[4-9]|23\d|127|169\.254)\.' -or $prefix -match '^(ff|fe80:|::1/)' -or $prefix -eq '255.255.255.255/32') { continue }
        Add-Row $Rows $prefix $r.NextHop $r.InterfaceAlias $r.RouteMetric
    }
}

function Read-ItemShare($Rows) {
    foreach ($s in @(Get-CimInstance -ClassName Win32_Share -ErrorAction Stop)) { Add-Row $Rows $s.Name $s.Path $s.Description }
}

# ドメインコントローラーには、ローカルユーザーがない
function Assert-NotDomainController {
    $role = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).DomainRole
    if ([int]$role -ge 4) { throw (New-Object InvalidOperationException 'NotApplicable') }
}

function Read-ItemLocalUser($Rows) {
    Assert-NotDomainController
    foreach ($u in @(Get-CimInstance -ClassName Win32_UserAccount -Filter 'LocalAccount=TRUE' -ErrorAction Stop)) {
        Add-Row $Rows $u.Name $(if ($u.Disabled) { '無効' } else { '有効' }) $(if ($u.PasswordExpires) { 'あり' } else { 'なし' }) $u.Description
    }
}

# グループのメンバーを、WinNT プロバイダーで読む（Get-LocalGroupMember は、名前を解決できない SID があると失敗する）
function Get-GroupMemberNames([string]$GroupName) {
    $group = [ADSI]('WinNT://./{0},group' -f $GroupName)
    foreach ($member in @($group.Invoke('Members'))) {
        $adsPath = [string]$member.GetType().InvokeMember('AdsPath', 'GetProperty', $null, $member, $null)
        # WinNT://DOMAIN/user、WinNT://WORKGROUP/HOST/user、解決できないものは WinNT://S-1-5-...
        $parts = @($adsPath -replace '^WinNT://', '' -split '/')
        if ($parts.Count -ge 2) { $parts[$parts.Count - 2] + '\' + $parts[$parts.Count - 1] } else { $parts[0] }
    }
}

function Read-ItemLocalGroup($Rows) {
    Assert-NotDomainController
    foreach ($g in @(Get-CimInstance -ClassName Win32_Group -Filter 'LocalAccount=TRUE' -ErrorAction Stop)) {
        $groupName = [string]$g.Name
        try {
            foreach ($memberName in @(Get-GroupMemberNames $groupName)) { Add-Row $Rows $groupName $memberName }
        } catch {
            Add-PartialSkip $groupName $_.Exception
        }
    }
}

function Read-ItemHosts($Rows) {
    $file = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
    if (-not [IO.File]::Exists($file)) { return }
    foreach ($line in ((Read-TextAuto $file) -split '\r\n|\n|\r')) {
        $body = ($line -replace '#.*$', '').Trim()
        if ($body -eq '') { continue }
        $parts = @($body -split '\s+')
        for ($i = 1; $i -lt $parts.Count; $i++) { Add-Row $Rows $parts[0] $parts[$i] }
    }
}

function Read-ItemEnvironment($Rows) {
    $key = Open-MachineKey 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment' 'Default'
    if (-not $key) { throw (New-Object InvalidOperationException 'システムの環境変数のレジストリを開けません') }
    try {
        foreach ($valueName in $key.GetValueNames()) {
            # %SystemRoot% などは展開せず、書かれたとおりに比べる
            Add-Row $Rows $valueName $key.GetValue($valueName, '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        }
    } finally {
        $key.Dispose()
    }
}

function Read-ItemFirewall($Rows) {
    foreach ($p in @(Get-NetFirewallProfile -PolicyStore ActiveStore -ErrorAction Stop)) {
        Add-Row $Rows 'プロファイル' $p.Name $p.Name $(if ("$($p.Enabled)" -eq 'True') { '有効' } else { '無効' }) '' '' '' '' '' '' '' ''
    }
    # 規則ごとに問い合わせると時間がかかるので、フィルターはまとめて取る
    $portFilters = @{}
    foreach ($f in @(Get-NetFirewallPortFilter -PolicyStore ActiveStore -ErrorAction Stop)) { $portFilters[[string]$f.InstanceID] = $f }
    $appFilters = @{}
    foreach ($f in @(Get-NetFirewallApplicationFilter -PolicyStore ActiveStore -ErrorAction Stop)) { $appFilters[[string]$f.InstanceID] = $f }
    $addressFilters = @{}
    foreach ($f in @(Get-NetFirewallAddressFilter -PolicyStore ActiveStore -ErrorAction Stop)) { $addressFilters[[string]$f.InstanceID] = $f }
    foreach ($r in @(Get-NetFirewallRule -PolicyStore ActiveStore -ErrorAction Stop)) {
        $id = [string]$r.InstanceID
        $protocol = ''; $localPort = ''; $remotePort = ''; $program = ''; $remoteAddress = ''
        if ($portFilters.ContainsKey($id)) {
            $protocol   = [string]$portFilters[$id].Protocol
            $localPort  = Join-Sorted $portFilters[$id].LocalPort
            $remotePort = Join-Sorted $portFilters[$id].RemotePort
        }
        if ($appFilters.ContainsKey($id)) { $program = [string]$appFilters[$id].Program }
        if ($addressFilters.ContainsKey($id)) { $remoteAddress = Join-Sorted $addressFilters[$id].RemoteAddress }
        $direction = switch ("$($r.Direction)") { 'Inbound' { '受信' } 'Outbound' { '送信' } default { "$($r.Direction)" } }
        $operation = switch ("$($r.Action)") { 'Allow' { '許可' } 'Block' { '拒否' } default { "$($r.Action)" } }
        Add-Row $Rows '規則' $id $r.DisplayName $(if ("$($r.Enabled)" -eq 'True') { '有効' } else { '無効' }) $direction $operation "$($r.Profile)" $protocol $localPort $remotePort $program $remoteAddress
    }
}

# 行の並びと重なりをそろえる。Dictionary は突き合わせにもそのまま使う
function ConvertTo-RowTable($Rows, $Def) {
    $table = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $Rows) {
        $key = (@(foreach ($index in $Def.KeyIndex) { $row[$index] }) -join '|')
        # 同じ対象が 2 回出たときは、最初のものを使う（ユーザーごとのサービスの実体など）
        if (-not $table.ContainsKey($key)) { $table[$key] = $row }
    }
    return , $table
}

function Get-SortedKeys($Table) {
    $keys = New-Object 'string[]' $Table.Count
    $Table.Keys.CopyTo($keys, 0)
    [Array]::Sort($keys, [StringComparer]::OrdinalIgnoreCase)
    return , $keys
}

# 今の状態を取る。Items は 項目名 → @{ Table; Status（空なら取れた）; Detail }
function Read-CurrentState([string[]]$Names) {
    $state = @{
        Host = $machineName; Time = (Get-Date); Admin = $isAdmin; AllTasks = $script:taskScopeAll
        Os = ''; Boot = ''; Items = @{}; Skips = New-Object 'System.Collections.Generic.List[object]'
    }
    try {
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $state.Os = [string]$osInfo.Version
        if ($osInfo.LastBootUpTime) { $state.Boot = ([datetime]$osInfo.LastBootUpTime).ToString($TimeFormat, $Invariant) }
        $build = Open-MachineKey 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' 'Default'
        if ($build) {
            try { $state.Os += '.' + [string]$build.GetValue('UBR', 0) } finally { $build.Dispose() }
        }
    } catch { }

    foreach ($name in $Names) {
        $def = $ItemDefs[$name]
        Write-Progress -Activity $ToolTitle -Status ('{0}を読んでいます' -f $def.Label)
        $rows = New-Object 'System.Collections.Generic.List[object]'
        $script:currentSkips = New-Object 'System.Collections.Generic.List[object]'
        $status = ''
        $detail = ''
        try {
            & "Read-Item$name" $rows
        } catch {
            $ex = Get-InnerException $_.Exception
            if ($ex -is [InvalidOperationException] -and $ex.Message -eq 'NotApplicable') {
                $status = '対象外'
                $detail = 'ドメインコントローラーには、ローカルユーザーとローカルグループがありません'
            } elseif (Test-AccessDenied $ex) {
                $status = '権限なし'
                $detail = 'アクセスが拒否されました'
            } else {
                $status = 'その他のエラー'
                $detail = Get-ErrorText $ex
            }
            $state.Skips.Add([object[]]@($status, $name, $detail))
        }
        foreach ($skip in $script:currentSkips) { $state.Skips.Add([object[]]@($skip[0], $name, $skip[1])) }
        $state.Items[$name] = @{ Table = (ConvertTo-RowTable $rows $def); Status = $status }
    }
    return $state
}

#endregion

#region スナップショットの保存と読み込み

function Save-State($State, [string]$Folder) {
    [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $Folder))
    $info = New-Object 'System.Collections.Generic.List[string]'
    $info.Add(('ツール = {0}' -f $ToolName))
    $info.Add(('バージョン = {0}' -f $ScriptVersion))
    $info.Add(('ホスト名 = {0}' -f $State.Host))
    $info.Add(('保存日時 = {0}' -f ([datetime]$State.Time).ToString($TimeFormat, $Invariant)))
    $info.Add(('管理者 = {0}' -f $(if ($State.Admin) { 'はい' } else { 'いいえ' })))
    $info.Add(('OS = {0}' -f $State.Os))
    $info.Add(('最終起動日時 = {0}' -f $State.Boot))
    $info.Add(('タスクの範囲 = {0}' -f $(if ($State.AllTasks) { 'すべて' } else { '\Microsoft\ の下を除く' })))
    foreach ($name in $ItemDefs.Keys) {
        if (-not $State.Items.ContainsKey($name)) { continue }
        $itemState = $State.Items[$name]
        $info.Add(('項目.{0} = {1}' -f $name, $(if ($itemState.Status) { $itemState.Status } else { '取得' })))
        if ($itemState.Status) { continue }
        $def   = $ItemDefs[$name]
        $table = $itemState.Table
        $rows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($key in (Get-SortedKeys $table)) {
            $rows.Add([string[]]@(foreach ($value in $table[$key]) { Protect-ExcelText $value }))
        }
        Save-Csv (ConvertTo-LongPath (Join-Path $Folder $def.File)) $def.Columns $rows
    }
    [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $Folder 'info.txt')), [string[]]$info, $Utf8Bom)
    $skipRows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($skip in $State.Skips) { $skipRows.Add([object[]]@($skip[0], $ItemDefs[$skip[1]].Label, (Protect-ExcelText $skip[2]))) }
    Save-Csv (ConvertTo-LongPath (Join-Path $Folder 'skipped.csv')) $SkipColumns $skipRows
}

# スナップショットの 1 つの CSV を読む。形が違うときは例外
function Read-ItemCsv([string]$FilePath, $Def) {
    $lines = [IO.File]::ReadAllLines((ConvertTo-LongPath $FilePath), [Text.Encoding]::UTF8)
    if ($lines.Count -eq 0) { throw '空のファイルです' }
    $quoted = ConvertTo-CsvLine $Def.Columns
    $plain  = $Def.Columns -join ','
    if (-not [string]::Equals($lines[0], $quoted, [StringComparison]::Ordinal) -and -not [string]::Equals($lines[0], $plain, [StringComparison]::Ordinal)) {
        throw '見出しの行が違います'
    }
    $rows = New-Object 'System.Collections.Generic.List[object]'
    if ($lines.Count -gt 1) {
        foreach ($record in @($lines[1..($lines.Count - 1)] | ConvertFrom-Csv -Header $Def.Columns)) {
            $rows.Add([string[]]@(foreach ($column in $Def.Columns) { Restore-ExcelText ([string]$record.$column) }))
        }
    }
    return , $rows
}

# スナップショットのフォルダを読む。読めないフォルダは止める。1 つの CSV が読めないだけなら、その項目を「読めない」にする
function Read-StateFolder([string]$Folder, [string]$Label, [string[]]$Names) {
    $folderLong = ConvertTo-LongPath $Folder
    $exists = $false
    try { $exists = [IO.Directory]::Exists($folderLong) } catch { }
    if (-not $exists) { Stop-WithError ('{0}のフォルダを開けません: {1}' -f $Label, $Folder) }
    $notSnapshot = '{0}のスナップショットのフォルダとして読めません: {1}（{2} の -Snapshot で保存したフォルダを指定してください）' -f $Label, $Folder, $ToolName
    $infoPath = ConvertTo-LongPath (Join-Path $Folder 'info.txt')
    if (-not [IO.File]::Exists($infoPath)) { Stop-WithError $notSnapshot }
    $info = @{}
    try {
        foreach ($line in ((Read-TextAuto $infoPath) -split '\r\n|\n|\r')) {
            $separator = $line.IndexOf(' = ')
            if ($separator -gt 0) { $info[$line.Substring(0, $separator)] = $line.Substring($separator + 3) }
        }
    } catch {
        Stop-WithError $notSnapshot
    }
    if (-not $info.ContainsKey('ツール') -or $info['ツール'] -ne $ToolName -or -not $info.ContainsKey('保存日時')) { Stop-WithError $notSnapshot }

    $time = $null
    try { $time = [datetime]::ParseExact($info['保存日時'], $TimeFormat, $Invariant) } catch { Stop-WithError $notSnapshot }
    $state = @{
        Host = [string]$info['ホスト名']; Time = $time; Admin = $info['管理者'] -eq 'はい'; AllTasks = $info['タスクの範囲'] -eq 'すべて'
        Os = [string]$info['OS']; Boot = [string]$info['最終起動日時']; Items = @{}; Skips = New-Object 'System.Collections.Generic.List[object]'; Saved = @()
    }
    $state.Saved = @(foreach ($name in $ItemDefs.Keys) { if ($info.ContainsKey('項目.' + $name)) { $name } })

    # 保存したときに一部を読めなかった項目は、そのことを引き継ぐ。項目の見出しから名前に戻す
    $skipFile = Join-Path $Folder 'skipped.csv'
    if ([IO.File]::Exists((ConvertTo-LongPath $skipFile))) {
        try {
            $skipLines = [IO.File]::ReadAllLines((ConvertTo-LongPath $skipFile), [Text.Encoding]::UTF8)
            if ($skipLines.Count -gt 1) {
                foreach ($record in @($skipLines[1..($skipLines.Count - 1)] | ConvertFrom-Csv -Header $SkipColumns)) {
                    foreach ($name in $ItemDefs.Keys) {
                        if ($ItemDefs[$name].Label -eq [string]$record.'項目') { $state.Skips.Add([object[]]@([string]$record.'理由', $name, (Restore-ExcelText ([string]$record.'詳細')))) }
                    }
                }
            }
        } catch { }
    }

    foreach ($name in $Names) {
        $def = $ItemDefs[$name]
        if (-not $info.ContainsKey('項目.' + $name)) {
            $state.Items[$name] = @{ Table = $null; Status = 'スナップショットにない' }
            $state.Skips.Add([object[]]@('スナップショットにない', $name, '保存したときに、この項目を指定していません'))
            continue
        }
        $saved = $info['項目.' + $name]
        if ($saved -ne '取得') {
            $state.Items[$name] = @{ Table = $null; Status = $saved }
            continue
        }
        try {
            $rows = Read-ItemCsv (Join-Path $Folder $def.File) $def
            $state.Items[$name] = @{ Table = (ConvertTo-RowTable $rows $def); Status = '' }
        } catch {
            $state.Items[$name] = @{ Table = $null; Status = '読めない' }
            $state.Skips.Add([object[]]@('読めない', $name, ('{0}: {1}' -f $def.File, (Get-ErrorText (Get-InnerException $_.Exception)))))
        }
    }
    return $state
}

#endregion

#region 比較

# 画面と CSV の「対象」の文章
function Get-TargetText([string]$Name, [string[]]$Row) {
    switch ($Name) {
        'Port' {
            $address = $Row[1]
            if ($address.Contains(':')) { $address = '[{0}]' -f $address }
            return '{0} {1}:{2}' -f $Row[0], $address, $Row[2]
        }
        'Software' { return '{0}（{1}）' -f $Row[0], $Row[1] }
        'Route' { return '{0} → {1}（{2}）' -f $Row[0], $Row[1], $Row[2] }
        'LocalGroup' { return '{0}: {1}' -f $Row[0], $Row[1] }
        'Hosts' { return '{0} {1}' -f $Row[0], $Row[1] }
        'Firewall' {
            if ($Row[0] -eq 'プロファイル') { return 'プロファイル ' + $Row[1] }
            return $(if ($Row[2]) { $Row[2] } else { $Row[1] })
        }
    }
    return $Row[0]
}

# 追加と削除の行の値（CSV 用）。比べる列を「列=値」の形でつなぐ
function Get-AllValuesText($Def, [string[]]$Row) {
    return (@(foreach ($column in $Def.Compare) { if ($Row[$Def.Index[$column]] -ne '') { '{0}={1}' -f $column, $Row[$Def.Index[$column]] } }) -join ' / ')
}

# 追加と削除の行の値（画面用）。パスワードが含まれることのある列は出さない
function Get-ScreenValuesText($Def, [string[]]$Row) {
    return (@(foreach ($column in $Def.Compare) { if ($Def.Hidden -notcontains $column -and $Row[$Def.Index[$column]] -ne '') { $Row[$Def.Index[$column]] } }) -join '・')
}

# 49152 番以上は、Windows が動的に割り当てる範囲で、再起動のたびに変わる
function Test-DynamicPort([string[]]$Row) {
    $number = 0
    return [int]::TryParse($Row[2], [ref]$number) -and $number -ge $DynamicPort
}

# 1 つの項目を比べて、結果の行を Results に加え、件数を返す
function Compare-Item([string]$Name, $RefTable, $DifTable, $Results) {
    $def = $ItemDefs[$Name]
    $counts = @{ Added = 0; Removed = 0; Changed = 0; Minor = 0 }
    $allKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($key in $RefTable.Keys) { [void]$allKeys.Add($key) }
    foreach ($key in $DifTable.Keys) { [void]$allKeys.Add($key) }
    $keys = New-Object 'string[]' $allKeys.Count
    $allKeys.CopyTo($keys)
    [Array]::Sort($keys, [StringComparer]::OrdinalIgnoreCase)

    foreach ($key in $keys) {
        $refRow = if ($RefTable.ContainsKey($key)) { $RefTable[$key] } else { $null }
        $difRow = if ($DifTable.ContainsKey($key)) { $DifTable[$key] } else { $null }
        $anyRow = if ($null -ne $refRow) { $refRow } else { $difRow }
        $target = Get-TargetText $Name $anyRow
        $isDynamic = $Name -eq 'Port' -and (Test-DynamicPort $anyRow)
        $note = if ($isDynamic) { '動的に割り当てられるポート' } else { '' }

        if ($null -eq $refRow -or $null -eq $difRow) {
            $isAdded = $null -eq $refRow
            $result  = if ($isDynamic) { '参考' } elseif ($isAdded) { '追加' } else { '削除' }
            if ($isDynamic) { $counts.Minor++ } elseif ($isAdded) { $counts.Added++ } else { $counts.Removed++ }
            $allText = Get-AllValuesText $def $anyRow
            $Results.Add([pscustomobject]@{
                Result = $result; Item = $Name; Target = $target; Attribute = ''
                RefValue = $(if ($isAdded) { '' } else { $allText }); DifValue = $(if ($isAdded) { $allText } else { '' })
                Note = $(if ($isDynamic) { '{0}（{1}）' -f $note, $(if ($isAdded) { '追加' } else { '削除' }) } else { '' })
                Screen = Get-ScreenValuesText $def $anyRow
            })
            continue
        }

        foreach ($column in $def.Compare) {
            $index = $def.Index[$column]
            if ([string]::Equals($refRow[$index], $difRow[$index], [StringComparison]::Ordinal)) { continue }
            $isMinor = $isDynamic
            $rowNote = $note
            # 手動開始のサービスは、必要なときだけ起動して止まるので、状態の変化だけなら通常の動き
            if ($Name -eq 'Service' -and $column -eq '状態') {
                $modeIndex = $def.Index['スタートアップ']
                if ($refRow[$modeIndex] -eq '手動' -and $difRow[$modeIndex] -eq '手動') {
                    $isMinor = $true
                    $rowNote = '手動開始のサービスの状態の変化'
                }
            }
            if ($isMinor) { $counts.Minor++ } else { $counts.Changed++ }
            $screen = if ($def.Hidden -contains $column) { '{0}が変わりました' -f $column } else { '{0}: {1} → {2}' -f $column, $refRow[$index], $difRow[$index] }
            $Results.Add([pscustomobject]@{
                Result = $(if ($isMinor) { '参考' } else { '変更' }); Item = $Name; Target = $target; Attribute = $column
                RefValue = $refRow[$index]; DifValue = $difRow[$index]; Note = $rowNote; Screen = $screen
            })
        }
    }
    return $counts
}

#endregion

#region 引数の確認

$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

$hasReference  = [bool]$Reference
$hasDifference = [bool]$Difference
if ($Snapshot) {
    if ($hasReference) { Stop-WithError '-Snapshot のときは、-Reference は指定できません' }
    if ($hasDifference) { Stop-WithError '-Snapshot のときは、-Difference は指定できません' }
    if ($NoSave) { Stop-WithError '-Snapshot のときは、-NoSave は指定できません' }
} else {
    if ($AllTasks) { Stop-WithError '-AllTasks は、-Snapshot のときだけ指定できます' }
    if ($hasDifference -and -not $hasReference) { Stop-WithError '-Difference を使うときは、-Reference も指定してください' }
    if (-not $hasReference) { Stop-WithError '作業の前は -Snapshot を、作業の後は -Reference を指定してください' }
}

# -File で起動すると配列を渡せないため、カンマ区切りの文字列で受ける
$itemNames = @()
if ($Item) {
    $validNames = @($ItemDefs.Keys)
    foreach ($part in ((Remove-Quote $Item) -split ',')) {
        $text = $part.Trim().Trim("'")
        if ($text -eq '') { continue }
        $matched = @($validNames | Where-Object { [string]::Equals($_, $text, [StringComparison]::OrdinalIgnoreCase) })
        if ($matched.Count -eq 0) { Stop-WithError ('-Item の値が正しくありません: {0}（使える値: {1}）' -f $text, ($validNames -join ', ')) }
        if ($itemNames -notcontains $matched[0]) { $itemNames += $matched[0] }
    }
    if ($itemNames.Count -eq 0) { Stop-WithError ('-Item に項目がありません: {0}' -f $Item) }
    # 表の順にそろえる
    $itemNames = @($validNames | Where-Object { $itemNames -contains $_ })
}

function Get-FolderArgument([string]$Name, [string]$Value) {
    $text = Remove-Quote $Value
    if ($text.Contains('"')) { Stop-WithQuoteError $Name $text }
    return Resolve-FullPath (Resolve-StartPath $text)
}

$refFolder = $null
$difFolder = $null
if ($hasReference) { $refFolder = Get-FolderArgument '-Reference' $Reference }
if ($hasDifference) { $difFolder = Get-FolderArgument '-Difference' $Difference }
if ($refFolder -and $difFolder -and $refFolder.Equals($difFolder, [StringComparison]::OrdinalIgnoreCase)) {
    Stop-WithError ('基準と比較が同じフォルダです: {0}' -f $refFolder)
}

$outputText = $null
if ($OutputFolder) {
    $outputText = Remove-Quote $OutputFolder
    if ($outputText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outputText }
    if ($outputText -match '^(\\\\\?\\)?[A-Za-z]:$') { $outputText += '\' }
    $outputText = Resolve-StartPath $outputText
}

$startTime = Get-Date
$taskScopeAll = [bool]$AllTasks

#endregion

#region 基準のスナップショットを読む

$refState = $null
$difState = $null
$compareNames = $itemNames
if (-not $Snapshot) {
    # 項目の指定がなければ、基準にある項目すべて。指定があれば、基準にあることを確かめる
    $refProbe = Read-StateFolder $refFolder '基準' @()
    if ($itemNames.Count -eq 0) {
        $compareNames = @($refProbe.Saved)
    } else {
        $absent = @($itemNames | Where-Object { $refProbe.Saved -notcontains $_ })
        if ($absent.Count -gt 0) { Stop-WithError ('基準のスナップショットに、次の項目がありません: {0}' -f ($absent -join ', ')) }
    }
    $refState = Read-StateFolder $refFolder '基準' $compareNames
    # 今の状態を取るときのタスクの範囲は、基準に合わせる（範囲が違うと、すべて「追加」に見えてしまう）
    $taskScopeAll = $refState.AllTasks
    if ($difFolder) { $difState = Read-StateFolder $difFolder '比較' $compareNames }
} elseif ($itemNames.Count -eq 0) {
    $compareNames = @($ItemDefs.Keys)
}

#endregion

#region 保存先の準備

# 読み終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir      = $null
$saveWarning = $null
if (-not $NoSave) {
    $baseDir = if ($outputText) { $outputText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        # \\?\ 付きのパスでは末尾のピリオドや空白が残るため、先に通常の書き方で整える（260 文字を超えると例外になるので、そのときはそのまま）
        try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
        $runPrefix = if ($Snapshot) { 'SystemSnapshot' } else { 'SystemDiff' }
        $runName = Format-Invariant '{0}_{1}_{2:yyyyMMdd_HHmmss}' $runPrefix $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while ([IO.Directory]::Exists((ConvertTo-LongPath $runDir)) -or [IO.File]::Exists((ConvertTo-LongPath $runDir))) {
            $suffix++
            $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix)
        }
        [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $runDir))
    } catch {
        $saveWarning = '保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, (Get-ErrorText (Get-InnerException $_.Exception))
        $runDir = $null
        if ($Snapshot) { Stop-WithError ('スナップショットを保存できません: {0}（{1}）' -f $baseDir, (Get-ErrorText (Get-InnerException $_.Exception))) }
    }
}

#endregion

$finished   = $false
$unexpected = $false
$saveError  = $null
$sw         = [Diagnostics.Stopwatch]::StartNew()
$exitCode   = 0
$header     = New-Object 'System.Collections.Generic.List[string]'
$tail       = New-Object 'System.Collections.Generic.List[string]'
$countLines = @()
$diffLines  = @()
$diffShown  = 0
try {
    if ($Snapshot) {
        #region スナップショットの保存

        $state = Read-CurrentState $compareNames
        Write-Progress -Activity $ToolTitle -Status '結果を保存しています'
        try {
            Save-State $state $runDir
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
        }

        $header.Add((Format-Invariant '{0} ({1} {2})' $ToolTitle $ToolName $ScriptVersion))
        $header.Add('スナップショットを保存します')
        $header.Add((Format-Invariant '{0}   {1:yyyy-MM-dd HH:mm:ss}   管理者: {2}' $machineName $startTime $(if ($isAdmin) { 'はい' } else { 'いいえ' })))
        $rows = @(foreach ($name in $compareNames) {
            $itemState = $state.Items[$name]
            , [string[]]@($ItemDefs[$name].Label, $(if ($itemState.Status) { $itemState.Status } else { Format-Invariant '{0:N0}' $itemState.Table.Count }))
        })
        $countLines = Format-Table ([string[]]@('項目', '件数')) ([bool[]]@($false, $true)) $rows
        $skipNotApplicable = @($state.Skips | Where-Object { $_[0] -ne '対象外' }).Count
        if ($skipNotApplicable -gt 0) { $tail.Add((Format-Invariant '※ 取れなかったものが {0:N0} 件あります（skipped.csv を確認してください）' $skipNotApplicable)) }
        if (-not $isAdmin) { $tail.Add('※ 管理者として実行すると、取れる情報が増えます') }
        $tail.Add(('所要 {0}' -f (Format-Duration $sw.Elapsed)))

        #endregion
    } else {
        #region 比較

        $isCurrent = -not $difState
        if ($isCurrent) { $difState = Read-CurrentState $compareNames }

        # 比べられない項目は、中身を比べず skipped.csv に残す。基準・比較のどちらの理由かを添える
        $skipRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($skip in $refState.Skips) { if ($compareNames -contains $skip[1]) { $skipRows.Add([object[]]@($skip[0], $skip[1], ('基準: ' + $skip[2]))) } }
        foreach ($skip in $difState.Skips) { if ($compareNames -contains $skip[1]) { $skipRows.Add([object[]]@($skip[0], $skip[1], ('比較: ' + $skip[2]))) } }

        $results   = New-Object 'System.Collections.Generic.List[object]'
        $itemCount = @{}
        $notCompared = New-Object 'System.Collections.Generic.List[string]'
        foreach ($name in $compareNames) {
            $refItem = $refState.Items[$name]
            $difItem = $difState.Items[$name]
            if ($refItem.Status -or $difItem.Status) {
                # どちらも「対象外」なら、比べるものがないだけ
                if (-not ($refItem.Status -eq '対象外' -and $difItem.Status -eq '対象外')) { $notCompared.Add($name) }
                continue
            }
            # 保存したもの同士で、タスクの範囲が違うと、すべて「追加」に見えてしまう
            if ($name -eq 'Task' -and $refState.AllTasks -ne $difState.AllTasks) {
                $skipRows.Add([object[]]@('その他のエラー', $name, '基準と比較で、タスクの範囲（-AllTasks）が違います'))
                $notCompared.Add($name)
                continue
            }
            $itemCount[$name] = Compare-Item $name $refItem.Table $difItem.Table $results
        }

        $added = 0; $removed = 0; $changed = 0; $minor = 0
        foreach ($counts in $itemCount.Values) { $added += $counts.Added; $removed += $counts.Removed; $changed += $counts.Changed; $minor += $counts.Minor }
        $diffTotal = $added + $removed + $changed
        # 「対象外」は比べられなかったことに数えない
        $partialCount = @($skipRows | Where-Object { $_[0] -ne '対象外' -and $notCompared -notcontains $_[1] }).Count
        $problemCount = $notCompared.Count + $partialCount
        if ($diffTotal -gt 0) { $exitCode = 2 } elseif ($problemCount -gt 0) { $exitCode = 1 }

        # 見出し
        $adminText = { param($Value) if ($Value) { 'はい' } else { 'いいえ' } }
        $header.Add((Format-Invariant '{0} ({1} {2})' $ToolTitle $ToolName $ScriptVersion))
        $header.Add((Format-Invariant '基準: {0} {1:yyyy-MM-dd HH:mm:ss}（管理者: {2}）  {3}' $refState.Host $refState.Time (& $adminText $refState.Admin) $refFolder))
        if ($isCurrent) {
            $header.Add((Format-Invariant '比較: 今の状態 {0} {1:yyyy-MM-dd HH:mm:ss}（管理者: {2}）' $difState.Host $difState.Time (& $adminText $difState.Admin)))
        } else {
            $header.Add((Format-Invariant '比較: {0} {1:yyyy-MM-dd HH:mm:ss}（管理者: {2}）  {3}' $difState.Host $difState.Time (& $adminText $difState.Admin) $difFolder))
        }
        if ($refState.Boot -and $difState.Boot) {
            if ($refState.Boot -eq $difState.Boot) {
                $header.Add('再起動: なし')
            } else {
                $header.Add(('再起動: あり（基準 {0} → 比較 {1}）' -f $refState.Boot.Replace('/', '-'), $difState.Boot.Replace('/', '-')))
            }
        }
        if ($refState.Os -and $difState.Os -and $refState.Os -ne $difState.Os) { $header.Add(('OS: {0} → {1}' -f $refState.Os, $difState.Os)) }

        # 項目ごとの件数
        $rows = @(foreach ($name in $compareNames) {
            $refItem = $refState.Items[$name]
            $difItem = $difState.Items[$name]
            $refText = if ($refItem.Status) { $refItem.Status } else { Format-Invariant '{0:N0}' $refItem.Table.Count }
            $difText = if ($difItem.Status) { $difItem.Status } else { Format-Invariant '{0:N0}' $difItem.Table.Count }
            if ($itemCount.ContainsKey($name)) {
                $counts = $itemCount[$name]
                , [string[]]@($ItemDefs[$name].Label, $refText, $difText, (Format-Invariant '{0:N0}' $counts.Added), (Format-Invariant '{0:N0}' $counts.Removed), (Format-Invariant '{0:N0}' $counts.Changed), (Format-Invariant '{0:N0}' $counts.Minor))
            } else {
                , [string[]]@($ItemDefs[$name].Label, $refText, $difText, '-', '-', '-', '-')
            }
        })
        $countLines = Format-Table ([string[]]@('項目', '基準', '比較', '追加', '削除', '変更', '参考')) ([bool[]]@($false, $true, $true, $true, $true, $true, $true)) $rows

        # 違いの一覧。「参考」は画面に出さない
        $mainRows = @(foreach ($row in $results) {
            if ($row.Result -ne '参考') { , [string[]]@($row.Result, $ItemDefs[$row.Item].Label, $row.Target, $row.Screen) }
        })
        if ($mainRows.Count -gt 0) {
            $diffLines = Format-Table ([string[]]@('結果', '項目', '対象', '変わった内容')) ([bool[]]@($false, $false, $false, $false)) $mainRows
        }

        if ($diffTotal -gt 0) {
            $resultText = '結果: 違いあり'
        } elseif ($problemCount -gt 0) {
            $resultText = Format-Invariant '結果: 違いなし（ただし、比べられなかった項目が {0:N0} 件あります。skipped.csv を確認してください）' $problemCount
        } else {
            $resultText = '結果: 違いなし'
        }
        $tail.Add((Format-Invariant '{0}   追加 {1:N0} 件  削除 {2:N0} 件  変更 {3:N0} 件  参考 {4:N0} 件   所要 {5}' $resultText $added $removed $changed $minor (Format-Duration $sw.Elapsed)))
        if ($problemCount -gt 0 -and $diffTotal -gt 0) { $tail.Add((Format-Invariant '比べられない項目 {0:N0} 件（skipped.csv を確認してください）' $problemCount)) }
        if ($refState.Host -ne $difState.Host) { $tail.Add('※ 違うコンピューターの状態を比べています。名前や IP アドレスの違いも出ます') }
        if ($refState.Admin -ne $difState.Admin) { $tail.Add('※ 基準と比較で、管理者かどうかが違います。取れる情報の違いが、違いとして出ることがあります') }
        if (-not $refState.Admin -and -not $difState.Admin) { $tail.Add('※ 管理者として実行すると、取れる情報が増えます') }

        #endregion

        #region 保存

        if ($runDir) {
            $lines = New-Object 'System.Collections.Generic.List[string]'
            $lines.AddRange([string[]]$header)
            $lines.Add('')
            $lines.Add('■ 項目ごとの件数')
            $lines.AddRange([string[]]$countLines)
            $lines.Add('')
            $lines.Add('■ 違い（項目の順）')
            if ($diffLines.Count -gt 0) { $lines.AddRange([string[]]$diffLines) } else { $lines.Add('    （違いはありません）') }
            $lines.Add('')
            $lines.AddRange([string[]]$tail.ToArray())
            $lines.Add(('保存先: {0}' -f $runDir))

            $csvRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($row in $results) {
                $csvRows.Add([object[]]@($row.Result, $ItemDefs[$row.Item].Label, (Protect-ExcelText $row.Target), $row.Attribute,
                    (Protect-ExcelText $row.RefValue), (Protect-ExcelText $row.DifValue), $row.Note))
            }
            $skipCsv = New-Object 'System.Collections.Generic.List[object]'
            foreach ($skip in $skipRows) { $skipCsv.Add([object[]]@($skip[0], $ItemDefs[$skip[1]].Label, (Protect-ExcelText $skip[2]))) }

            Write-Progress -Activity $ToolTitle -Status '結果を保存しています'
            try {
                [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'differences.csv')) $DiffColumns $csvRows
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $skipCsv
                if ($isCurrent) { Save-State $difState (Join-Path $runDir 'snapshot') }
            } catch {
                $saveError = Get-ErrorText (Get-InnerException $_.Exception)
            }
        }

        #endregion
    }

    $finished = $true
} catch {
    # 想定外のエラーは、Ctrl+C の「中断しました」と区別して知らせる
    $unexpected = $true
    throw
} finally {
    Write-Progress -Activity $ToolTitle -Completed
    if (-not $finished -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        $runDir = $null
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}
if ($Snapshot -and $saveError) { Stop-WithError ('スナップショットを保存できませんでした: {0}' -f $saveError) }

#region 画面表示

foreach ($line in $header) { Write-Host $line }
Write-Host ''
if (-not $Snapshot) { Write-Host '■ 項目ごとの件数' -ForegroundColor Cyan }
foreach ($line in $countLines) { Write-Host $line }
Write-Host ''
if (-not $Snapshot) {
    Write-Host '■ 違い（項目の順）' -ForegroundColor Cyan
    if ($diffLines.Count -eq 0) {
        Write-Host '    （違いはありません）'
    } else {
        # 保存できたときは、画面には先頭の分だけを出す。保存しなかったときは、ほかに見る場所がないためすべて出す（表の 1 行目は見出し）
        $limit = if ($runDir) { $ScreenRows + 1 } else { $diffLines.Count }
        $diffShown = [math]::Min($limit, $diffLines.Count)
        for ($i = 0; $i -lt $diffShown; $i++) { Write-Host $diffLines[$i] }
        if ($diffLines.Count -gt $diffShown) {
            Write-Host (Format-Invariant '（ほか {0:N0} 件は differences.csv を参照してください）' ($diffLines.Count - $diffShown))
        }
    }
    Write-Host ''
}
foreach ($line in $tail) { Write-Host $line }

if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($saveWarning) {
    Write-Warning $saveWarning
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
    if ($Snapshot) {
        $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { '.\Compare-SystemState.ps1' }
        Write-Host ''
        Write-Host '作業の後で、次のコマンドで比べられます:'
        Write-Host ('powershell -ExecutionPolicy Bypass -File {0} -Reference {1}' -f (Format-Argument $scriptPath), (Format-Argument $runDir))
    }
}

#endregion

exit $exitCode
