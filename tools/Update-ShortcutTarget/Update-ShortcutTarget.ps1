<#
.SYNOPSIS
    ショートカット（.lnk）のリンク先にある共有フォルダのサーバーを、確認したうえでまとめて置き換えます。

.DESCRIPTION
    指定したフォルダの中のショートカットを調べ、リンク先・引数・作業フォルダ・アイコンの場所・コメントにある
    共有フォルダのパス（\\サーバー\…）を一覧にします。

    -From を指定しないときは、棚卸しとして、共有フォルダを指すショートカットをすべて一覧にし、サーバーごとの件数を表示します。
    -From だけを指定したときは、そのサーバーを指すショートカットを一覧にします。
    -From と -To を指定したときは、置き換え先があるかを確かめてから置き換えの予定を表示し、Y を入力すると置き換えます。
    元のファイルは、結果のフォルダの backup に控えます。-To を指定しないときは、ファイルを変更しません。

    権限がなくて読めないフォルダはスキップし、件数を表示します。
    ジャンクションやシンボリックリンクはたどりません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Update-ShortcutTarget.ps1 -Path C:\Users\yamada\Desktop

.PARAMETER Path
    調べるフォルダか、ショートカットのファイル。フォルダを指定すると、サブフォルダも含めて調べます。
    複数指定するときは、カンマで区切って 1 つの文字列にします。必ず指定します。

.PARAMETER From
    置き換える前のサーバー。IP アドレスかサーバー名（192.168.1.10、oldfs）、または共有フォルダのパス（\\192.168.1.10\営業）で書きます。
    カンマで区切って複数書けます。省略すると、共有フォルダを指すショートカットをすべて一覧にします（棚卸し）。

.PARAMETER To
    置き換えた後のサーバー。-From と同じ形で 1 つだけ書きます。省略すると、置き換えません。

.PARAMETER Exclude
    調べないフォルダやファイル。ワイルドカード（*）で、カンマで区切って書きます。
    \ を含まない書き方は名前に、ドライブや \\ で始まる書き方はフルパスに、それ以外は -Path から見た相対パスに当てはめます。

.PARAMETER Force
    確認せずに置き換えます。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。置き換えるときは使えません。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Update-ShortcutTarget.ps1 -Path C:\Users\yamada\Desktop

    デスクトップのショートカットのうち、共有フォルダを指すものを、サーバーごとに一覧にします（ファイルは変更しません）。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Update-ShortcutTarget.ps1 -Path C:\Users\yamada\Desktop -From 192.168.1.10 -To 192.168.2.20

    192.168.1.10 を指すショートカットを、192.168.2.20 に置き換えます（確認してから置き換えます）。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Update-ShortcutTarget.ps1 -Path C:\Users\yamada\Desktop -From \\192.168.1.10\営業 -To \\192.168.2.20\sales

    共有名も変わったときに、\\192.168.1.10\営業 の下を \\192.168.2.20\sales の下に置き換えます。

.NOTES
    Update-ShortcutTarget 1.0.0
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
    [string[]]$Path = @(),

    [string[]]$From = @(),

    [string[]]$To = @(),

    [string[]]$Exclude = @(),

    [switch]$Force,

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

$ScriptVersion  = '1.0.0'
$ToolTitle      = 'ショートカットのリンク先置き換え'
$ScreenRows     = 200       # 画面に表示するショートカットの件数（保存するときだけ。超えた分はファイルを見てもらう）
$ProgressMs     = 250       # 進捗表示の間隔（ミリ秒）
$MaxLinkBytes   = 1MB       # これより大きいファイルは、ショートカットとして読まない
$PathWidth      = 50        # 画面の表で、パスの列を広げる上限（半角の文字数）
$ServerWidth    = 30        # 画面の表で、サーバーの列を広げる上限
$OutputPrefix   = 'ShortcutTarget_'
# このツールの結果のフォルダの名前（ShortcutTarget_<ホスト名>_<日時>）。同じ文字で始まるだけの利用者のフォルダは調べる
$OutputNameRule = New-Object Text.RegularExpressions.Regex ('^ShortcutTarget_.+_[0-9]{8}_[0-9]{6}(?:_[0-9]+)?$'), ([Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant')
$SystemFolders  = '$RECYCLE.BIN', 'System Volume Information'
# オンライン専用ファイルの属性（OFFLINE・RECALL_ON_OPEN・RECALL_ON_DATA_ACCESS）。読むとダウンロードが始まる
$OnlineMask     = 0x00441000
$ShortcutColumns = 'ショートカット', 'サーバー', 'リンク先', '引数', '作業フォルダ', 'アイコンの場所', 'コメント', '最終更新日時', '結果', '備考', '控え'
$ChangeColumns  = 'ショートカット', '項目', '置き換える前', '置き換えた後', '結果'
$SkipColumns    = '理由', 'パス', '詳細'
$TimeFormat     = 'yyyy/MM/dd HH:mm:ss'
$Invariant      = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom        = New-Object Text.UTF8Encoding $true
$OnWindows      = [IO.Path]::DirectorySeparatorChar -eq '\'
$machineName    = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

# ショートカットのファイルの先頭（ヘッダーの大きさ 0x4C と、ShellLink の CLSID {00021401-0000-0000-C000-000000000046}）
$LinkHeader = [byte[]](0x4C, 0, 0, 0, 0x01, 0x14, 0x02, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46)

# 置き換えの予定と結果
$StPlan     = '置き換える'
$StNotFound = '置き換え先が見つからない'
$StUnreach  = '置き換え先に接続できない'
$StCannot   = '置き換えられない'
$StDone     = '置き換えた'
$StFail     = '失敗'
$StNo       = '置き換えなかった'
$StatusOrder = @{ $StPlan = 0; $StNotFound = 1; $StUnreach = 2; $StCannot = 3 }

# ショートカットの 5 つの項目。リンク先・作業フォルダ・アイコンの場所は値全体が 1 つのパス、引数とコメントは文
$FieldKeys  = 'Target', 'Arguments', 'WorkDir', 'IconPath', 'Comment'
$FieldNames = @{ Target = 'リンク先'; Arguments = '引数'; WorkDir = '作業フォルダ'; IconPath = 'アイコンの場所'; Comment = 'コメント' }
$PathFields = @{ Target = $true; WorkDir = $true; IconPath = $true }

# WScript.Shell が正しく扱える長さ（文字数）。コメントと作業フォルダはこれを超えると、読んだときにリンク先が空になり、
# ほかの項目にも別の項目の文字が混ざる。アイコンの場所と引数は、これを超えると書くときに黙って切られる。
# リンク先は、MAX_PATH（260）未満でないと設定できない
$MaxLength = @{ Target = 259; Arguments = 1023; WorkDir = 260; IconPath = 259; Comment = 260 }

# ショートカットのファイルの、文字の項目（StringData）の順と目印（LinkFlags のビット）。この順に並んでいる
$StringFields = @(, @(0x4, 'Comment')) + @(, @(0x8, 'Relative')) + @(, @(0x10, 'WorkDir')) + @(, @(0x20, 'Arguments')) + @(, @(0x40, 'IconPath'))
# Unicode でない（古い形式の）ショートカットの文字は、OS の既定の文字コードで書かれている
$ansiEncoding = [Text.Encoding]::Default
try { $ansiEncoding = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage) } catch { }

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

# Windows のエラー番号（HRESULT の下位 16 ビット）
function Get-ErrorCode([Exception]$Exception) {
    return $Exception.HResult -band 0xFFFF
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

# Excel は、= + - @ で始まる値を数式に、数字と区切りの記号だけの値（192.168.1.10、001）や「.1」「1e5」を
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

# 調べている途中で止めるときも、finally が「中断しました」と出さないよう、止めたことを記録しておく
$stopping = $false
function Stop-WithError([string]$Message) {
    $script:stopping = $true
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value) {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -Path "C:\data"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# %USERPROFILE% のような環境変数を展開する。ログオンスクリプトやタスクスケジューラでは、
# コマンドプロンプトを通らずに起動されて、展開されないまま届くことがあるため
function Expand-PathText([string]$Text) {
    if (-not $Text.Contains('%')) { return $Text }
    return [Environment]::ExpandEnvironmentVariables($Text)
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

# ワイルドカードは * だけにして、それ以外の文字は文字どおりに扱う。5.1 の WildcardPattern.Escape は
# バッククォートをエスケープしないため、自分でエスケープしてから照合用のパターンを作る
$wildcardOptions = [Management.Automation.WildcardOptions]'IgnoreCase, CultureInvariant'
function New-NameMatcher([string]$Text) {
    $escaped = $Text.Replace('`', '``').Replace('[', '`[').Replace(']', '`]').Replace('?', '`?')
    return [Management.Automation.WildcardPattern]::new($escaped, $wildcardOptions)
}

# -File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# コマンドプロンプトでは ' がそのまま届くため、各要素の前後の ' も外す
function Get-ListText([string[]]$Values, [string]$Name) {
    $texts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($value in $Values) {
        foreach ($part in ($value -split ',')) {
            $text = (Remove-Quote $part).Trim("'")
            if ($text.Contains('"')) { Stop-WithQuoteError $Name $text }
            if ($text) { $texts.Add($text) }
        }
    }
    return , $texts
}

# 文字の番号の順に並べるための鍵（Sort-Object の文字列の比べ方は 5.1 と 7 で違い、記号の扱いで順が入れ替わるため）
function Get-SortText([string]$Text) {
    return $Text.ToUpperInvariant()
}

#endregion

#region 引数の確認

$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

# -Path は、-File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# カンマを含む名前のフォルダのため、全体が 1 つのフォルダかファイルとして存在するときは区切らない
$rootTexts = New-Object 'System.Collections.Generic.List[string]'
foreach ($item in $Path) {
    $whole = Remove-Quote $item
    if ($whole.Contains('"')) { Stop-WithQuoteError '-Path' $whole }
    $whole = Expand-PathText $whole
    $isWhole = $false
    if ($whole.Contains(',')) {
        try {
            $wholeFull = ConvertTo-LongPath (Resolve-FullPath (Resolve-StartPath $whole))
            $isWhole = [IO.Directory]::Exists($wholeFull) -or [IO.File]::Exists($wholeFull)
        } catch { }
    }
    if ($isWhole) { $rootTexts.Add($whole); continue }
    foreach ($part in ($whole -split ',')) {
        $text = (Remove-Quote $part).Trim("'")
        if ($text) { $rootTexts.Add($text) }
    }
}
if ($rootTexts.Count -eq 0) { Stop-WithError '調べるフォルダかショートカットを -Path で指定してください' }

$roots = New-Object 'System.Collections.Generic.List[object]'
foreach ($text in $rootTexts) {
    $rootText = $text
    if ($rootText -match '^(\\\\\?\\)?[A-Za-z]:$') { $rootText += '\' }
    $rootFull = Resolve-FullPath (Resolve-StartPath $rootText)
    $rootLong = ConvertTo-LongPath $rootFull
    $isFile = $false
    $isDir  = $false
    try { $isFile = [IO.File]::Exists($rootLong) } catch { }
    if (-not $isFile) { try { $isDir = [IO.Directory]::Exists($rootLong) } catch { } }
    if (-not $isFile -and -not $isDir) {
        Stop-WithError ((@(('フォルダかファイルが見つかりません: {0}' -f $rootFull)) + (Get-NetworkHint $rootFull)) -join "`r`n")
    }
    if ($isFile -and -not $rootFull.EndsWith('.lnk', [StringComparison]::OrdinalIgnoreCase)) {
        Stop-WithError ('ショートカット（.lnk）ではありません: {0}' -f $rootFull)
    }
    if ($isDir) {
        # フォルダを中まで読めるかを確かめる（共有フォルダは、接続していないと読めない）
        try {
            [void][IO.Directory]::EnumerateFileSystemEntries($rootLong).GetEnumerator().MoveNext()
        } catch {
            $reason = Get-ErrorText (Get-InnerException $_.Exception)
            Stop-WithError ((@(('フォルダを開けません: {0}（{1}）' -f $rootFull, $reason)) + (Get-NetworkHint $rootFull)) -join "`r`n")
        }
    }
    $roots.Add([pscustomobject]@{ Full = $rootFull; Long = $rootLong; IsFile = $isFile })
}

# 同じフォルダや、片方の中にあるフォルダを指定すると、同じショートカットを二重に数えてしまう
for ($i = 0; $i -lt $roots.Count; $i++) {
    for ($j = $i + 1; $j -lt $roots.Count; $j++) {
        $firstPrefix  = $roots[$i].Full.TrimEnd('\') + '\'
        $secondPrefix = $roots[$j].Full.TrimEnd('\') + '\'
        if ($firstPrefix.StartsWith($secondPrefix, [StringComparison]::OrdinalIgnoreCase) -or $secondPrefix.StartsWith($firstPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            Stop-WithError ('指定したフォルダの片方が、もう片方の中にあります: {0} と {1}' -f $roots[$i].Full, $roots[$j].Full)
        }
    }
}

# サーバー名に使えない文字（\ / : * ? " < > | と空白、カンマ）。共有名とフォルダ名には空白を使える
$ServerChars = '[^\\/:*?"<>|\s,]'
$serverRule  = New-Object Text.RegularExpressions.Regex ('^' + $ServerChars + '+$')
$folderRule  = New-Object Text.RegularExpressions.Regex ('^[^\\/:*?"<>|\x00-\x1F]+$')

# -From と -To の 1 つを読む。\\ の後ろ（Body）が 1 つの名前ならサーバー、\ で区切られていれば共有フォルダのパス
function ConvertTo-ServerSpec([string]$Text, [string]$Name) {
    $body = $Text
    if ($body.StartsWith('\\')) { $body = $body.Substring(2) }
    $body = $body.TrimEnd('\')
    $parts = $body.Split('\')
    $valid = $body.Length -gt 0 -and $serverRule.IsMatch($parts[0]) -and $parts[0] -ne '.'
    for ($k = 1; $k -lt $parts.Count -and $valid; $k++) {
        $valid = $folderRule.IsMatch($parts[$k]) -and $parts[$k].Trim().Length -gt 0
    }
    if (-not $valid) {
        Stop-WithError ('{0} は IP アドレスかサーバー名、または \\サーバー\共有 の形で指定してください: {1}' -f $Name, $Text)
    }
    $isShare = $parts.Count -ge 2
    return [pscustomobject]@{
        Body    = $body
        IsShare = $isShare
        Server  = $parts[0]
        # 画面の見出しと、エラーの表示に使う（サーバーは書いたとおり、共有フォルダのパスは \\ を付ける）
        Label   = $(if ($isShare) { '\\' + $body } else { $body })
    }
}

$fromSpecs = New-Object 'System.Collections.Generic.List[object]'
foreach ($text in (Get-ListText $From '-From')) {
    $spec = ConvertTo-ServerSpec $text '-From'
    $dup = $false
    foreach ($other in $fromSpecs) { if ([string]::Equals($other.Body, $spec.Body, [StringComparison]::OrdinalIgnoreCase)) { $dup = $true } }
    if (-not $dup) { $fromSpecs.Add($spec) }
}
$toTexts = Get-ListText $To '-To'
if ($toTexts.Count -gt 0 -and $fromSpecs.Count -eq 0) { Stop-WithError '-To を使うときは、置き換える前のサーバーを -From で指定してください' }
if ($toTexts.Count -gt 1) { Stop-WithError '-To は 1 つだけ指定してください' }
$toSpec = $null
if ($toTexts.Count -eq 1) { $toSpec = ConvertTo-ServerSpec $toTexts[0] '-To' }

$replacing = $null -ne $toSpec
$finding   = $fromSpecs.Count -gt 0
$toInsideFrom = $null   # -To が -From の中にあるときの、2 回実行したときのパスの例
if ($replacing) {
    foreach ($spec in $fromSpecs) {
        if ($spec.IsShare -ne $toSpec.IsShare) { Stop-WithError '-From と -To は、どちらもサーバー、またはどちらも \\サーバー\共有 の形で指定してください' }
        if ([string]::Equals($spec.Body, $toSpec.Body, [StringComparison]::OrdinalIgnoreCase)) { Stop-WithError ('-From と -To が同じです: {0}' -f $spec.Label) }
        if ($toSpec.Body.StartsWith($spec.Body + '\', [StringComparison]::OrdinalIgnoreCase)) {
            $toInsideFrom = '\\' + $toSpec.Body + $toSpec.Body.Substring($spec.Body.Length)
        }
    }
    if ($NoSave) { Stop-WithError '置き換えるときは -NoSave を使えません（元のファイルを控えられないため）' }
}

# 共有フォルダのパスを見つける規則。先頭の \\（または \\?\UNC\）を 1 番目のグループにする。
# パスの項目（リンク先・作業フォルダ・アイコンの場所）は値の先頭だけ、文（引数・コメント）は文の中のどこでも探す。
# 文の中では、前が \ のもの（\\\\ の途中など）は見ない
$UncLead      = '(\\\\\?\\UNC\\|\\\\)'
$regexOptions = [Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant'
# \\.\pipe\… のような、サーバーが . だけのものはデバイスの名前なので見ない
$NotDevice = '(?!\.(?!' + $ServerChars + '))'
$pathServerRule = [Text.RegularExpressions.Regex]::new('^' + $UncLead + $NotDevice + '(' + $ServerChars + '+)(?=\\|$)', $regexOptions)
$textServerRule = [Text.RegularExpressions.Regex]::new('(?<!\\)' + $UncLead + $NotDevice + '(' + $ServerChars + '+)', $regexOptions)

# -From に当たる規則。長いものを先に試し、1 か所を 1 回だけ置き換えるよう、1 つの正規表現にまとめる。
# サーバーは、文の中ではサーバー名が続かないところまで、共有フォルダのパスは、文の中では \ か " か文の終わりまでが一致したものだけ
$fromPathRule = $null
$fromTextRule = $null
if ($finding) {
    $pathAlts = New-Object 'System.Collections.Generic.List[string]'
    $textAlts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($spec in @($fromSpecs | Sort-Object -Property @{ Expression = { $_.Body.Length }; Descending = $true })) {
        $escaped = [Text.RegularExpressions.Regex]::Escape($spec.Body)
        $pathAlts.Add($escaped + '(?=\\|$)')
        if ($spec.IsShare) { $textAlts.Add($escaped + '(?=[\\"]|$)') } else { $textAlts.Add($escaped + '(?!' + $ServerChars + ')') }
    }
    $fromPathRule = [Text.RegularExpressions.Regex]::new('^' + $UncLead + '(?:' + ($pathAlts -join '|') + ')', $regexOptions)
    $fromTextRule = [Text.RegularExpressions.Regex]::new('(?<!\\)' + $UncLead + '(?:' + ($textAlts -join '|') + ')', $regexOptions)
}
# 置き換えた後の文字。-To に $ があっても置き換えの記号として扱われないよう、$$ にする
$replacement = $null
if ($replacing) { $replacement = '${1}' + $toSpec.Body.Replace('$', '$$') }

# -Exclude は、\ を含まない書き方は名前に、ドライブや \\ で始まる書き方はフルパスに、それ以外の \ を含む書き方は相対パスに当てはめる。
# 相対パスの先頭の \ と .\ は「直下から」の意味なので外す
$excludeRules = @(
    foreach ($text in (Get-ListText $Exclude '-Exclude')) {
        $text = $text.TrimEnd('\')
        if (-not $text) { continue }
        if ($text -match '^([A-Za-z]:\\|\\\\)') {
            [pscustomobject]@{ Text = $text; Scope = 'Full'; Matcher = (New-NameMatcher $text) }
        } elseif ($text.Contains('\')) {
            [pscustomobject]@{ Text = $text; Scope = 'Rel'; Matcher = (New-NameMatcher ($text -replace '^\.\\', '').TrimStart('\')) }
        } else {
            [pscustomobject]@{ Text = $text; Scope = 'Name'; Matcher = (New-NameMatcher $text) }
        }
    }
)
$hasExclude = $excludeRules.Count -gt 0

$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    $outText = Resolve-StartPath (Expand-PathText $outText)
}

# ショートカットのリンク先を読むには、Windows に入っている WScript.Shell を使う
$shell = $null
try {
    $shell = New-Object -ComObject WScript.Shell
} catch {
    Stop-WithError ('ショートカットを読み書きする準備ができません（WScript.Shell を使えません）。制約付き言語モードなど、制限された環境では使えません: {0}' -f (Get-ErrorText (Get-InnerException $_.Exception)))
}
# 書き換えるには Shell.Application（エクスプローラーと同じ部品）を使う。エクスプローラーで作ったショートカットは、
# リンク先の表示名やパスの控え（PropertyStoreDataBlock）を持っていて、WScript.Shell でリンク先を変えると、控えに古いサーバーが残るため。
# Shell.Application は、リンク先の情報と一緒に控えも作り直す。ファイルの場所を基準にした相対パスも作り直す
$shellApp = $null
if ($replacing) {
    try {
        $shellApp = New-Object -ComObject Shell.Application
    } catch {
        Stop-WithError ('ショートカットを書き換える準備ができません（Shell.Application を使えません）。Server Core など、エクスプローラーのない環境では置き換えられません: {0}' -f (Get-ErrorText (Get-InnerException $_.Exception)))
    }
}

$startTime = Get-Date

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
        # 置き換えるときは、元のファイルを控えられないため止める
        if ($replacing) { Stop-WithError ('保存先に書き込めないため、置き換えずに止めました: {0}（{1}）' -f $baseDir, $reason) }
        $saveWarning = '保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, $reason
        Write-Warning $saveWarning
    }
}

#endregion

#region 判定の部品

$dirType  = [IO.DirectoryInfo]
$rpFlag   = [IO.FileAttributes]::ReparsePoint
$hsFlag   = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System

$skipped   = New-Object 'System.Collections.Generic.List[object]'
$skipKinds = [ordered]@{ '権限なし' = 0; '読めないショートカット' = 0; 'クラウドのみ' = 0; 'その他のエラー' = 0 }
$excludedCount = 0
$linkCount     = 0
$linkTotal     = 0         # 読めたショートカットの数
$items         = New-Object 'System.Collections.Generic.List[object]'   # 一覧に出すショートカット
$sha256        = [Security.Cryptography.SHA256]::Create()
$sw            = [Diagnostics.Stopwatch]::StartNew()
$nextProgress  = 0L
# 260 文字以上のパスのショートカットは WScript.Shell で開けないため、短いパスの一時ファイルに写して読む
$tempLink      = Join-Path ([IO.Path]::GetTempPath()) ('UpdateShortcutTarget_{0}.lnk' -f $PID)

function Add-Skip([string]$Reason, [string]$ItemPath, [string]$Detail) {
    $skipped.Add([object[]]@($Reason, $ItemPath, $Detail))
    if ($skipKinds.Contains($Reason)) { $skipKinds[$Reason]++ }
}

function Add-SkipError([Exception]$Exception, [string]$ItemPath) {
    if ($Exception -is [UnauthorizedAccessException]) { Add-Skip '権限なし' $ItemPath 'アクセスが拒否されました' }
    else { Add-Skip 'その他のエラー' $ItemPath (Get-ErrorText $Exception) }
}

# 除外に当てはまれば、当てはまった書き方を返す
function Get-ExcludeMatch([string]$Rel, [string]$Name, [string]$Full) {
    foreach ($rule in $excludeRules) {
        $target = switch ($rule.Scope) { 'Name' { $Name } 'Rel' { $Rel } default { $Full } }
        if ($rule.Matcher.IsMatch($target)) { return $rule.Text }
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

# WScript.Shell が開けるパスか（260 文字未満）
function Test-ShortEnough([string]$Display) {
    return $Display.Length -lt 260
}

# アイコンの場所（パス,番号）を、パスと番号に分ける
function Split-IconLocation([string]$Raw) {
    $comma = $Raw.LastIndexOf(',')
    if ($comma -lt 0) { return , @($Raw, '0') }
    return , @($Raw.Substring(0, $comma), $Raw.Substring($comma + 1))
}

# ショートカットのファイルから、コメント・作業フォルダ・引数・アイコンの場所（StringData）と、アイコンの番号を読む。
# WScript.Shell は、コメントの「@shell32.dll,-22579」のような値を訳した文に変えて返し、長い値では項目がずれるため、
# ファイルにある値そのものを自分で読む。形式が正しくなければ $null
function Read-LinkStrings([byte[]]$Bytes) {
    $flags = [BitConverter]::ToUInt32($Bytes, 0x14)
    $pos = [long]0x4C
    if ($flags -band 0x1) {
        if ($pos + 2 -gt $Bytes.Length) { return $null }
        $pos += 2 + [BitConverter]::ToUInt16($Bytes, [int]$pos)
    }
    if ($flags -band 0x2) {
        if ($pos + 4 -gt $Bytes.Length) { return $null }
        $pos += [BitConverter]::ToUInt32($Bytes, [int]$pos)
    }
    $unicode = ($flags -band 0x80) -ne 0
    $values = @{ Comment = ''; Relative = ''; WorkDir = ''; Arguments = ''; IconPath = ''; IconIndex = ([string][BitConverter]::ToInt32($Bytes, 0x38)); EnvTarget = '' }
    foreach ($field in $StringFields) {
        if (($flags -band $field[0]) -eq 0) { continue }
        if ($pos + 2 -gt $Bytes.Length) { return $null }
        $count = [BitConverter]::ToUInt16($Bytes, [int]$pos)
        $pos += 2
        $size = if ($unicode) { 2 * $count } else { $count }
        if ($pos + $size -gt $Bytes.Length) { return $null }
        if ($unicode) { $values[$field[1]] = [Text.Encoding]::Unicode.GetString($Bytes, [int]$pos, $size) }
        else { $values[$field[1]] = $ansiEncoding.GetString($Bytes, [int]$pos, $size) }
        $pos += $size
    }
    # 後ろの拡張データ（ExtraData）から、環境変数の形のリンク先（EnvironmentVariableDataBlock、0xA0000001）を読む。
    # サーバーそのもの（\\192.168.1.10）を指すショートカットは、WScript.Shell ではリンク先が空になり、ここにだけ名前がある。
    # 中の Unicode の欄は空のことがあるため、そのときは ANSI の欄を読む。拡張データが壊れていても、ほかの値は使う
    while ($pos + 8 -le $Bytes.Length) {
        $blockSize = [BitConverter]::ToUInt32($Bytes, [int]$pos)
        if ($blockSize -lt 8 -or $pos + $blockSize -gt $Bytes.Length) { break }
        # 16 進数のリテラル 0xA0000001 は Int32（負の数）になるため、Int32 で読んで比べる
        if ([BitConverter]::ToInt32($Bytes, [int]$pos + 4) -eq 0xA0000001 -and $blockSize -ge 0x314) {
            $text = [Text.Encoding]::Unicode.GetString($Bytes, [int]$pos + 268, 520)
            if ($text.IndexOf([char]0) -ge 0) { $text = $text.Substring(0, $text.IndexOf([char]0)) }
            if (-not $text) {
                $text = $ansiEncoding.GetString($Bytes, [int]$pos + 8, 260)
                if ($text.IndexOf([char]0) -ge 0) { $text = $text.Substring(0, $text.IndexOf([char]0)) }
            }
            $values.EnvTarget = $text
            break
        }
        $pos += $blockSize
    }
    return $values
}

# サーバーそのもの（\\サーバー）の形か
$serverOnlyRule = [Text.RegularExpressions.Regex]::new('^\\\\' + '(?!\.(?![^\\/:*?"<>|\s,]))' + '[^\\/:*?"<>|\s,]+\\?$', [Text.RegularExpressions.RegexOptions]'CultureInvariant')

# 値の中の共有フォルダのサーバーを、書かれた順に集める（大文字と小文字だけが違うものは 1 つにまとめる）
function Get-ServerList($Item) {
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $FieldKeys) {
        $value = [string]$Item.($key)
        if (-not $value) { continue }
        $rule = if ($PathFields.ContainsKey($key)) { $pathServerRule } else { $textServerRule }
        foreach ($found in $rule.Matches($value)) {
            $name = $found.Groups[2].Value
            $dup = $false
            foreach ($known in $list) { if ([string]::Equals($known, $name, [StringComparison]::OrdinalIgnoreCase)) { $dup = $true; break } }
            if (-not $dup) { $list.Add($name) }
        }
    }
    return , $list
}

# 規則に最初に当たる項目と、文の中の位置を返す
function Find-FirstMatch($Item, $PathRule, $TextRule) {
    foreach ($key in $FieldKeys) {
        $value = [string]$Item.($key)
        if (-not $value) { continue }
        $rule = if ($PathFields.ContainsKey($key)) { $PathRule } else { $TextRule }
        $found = $rule.Match($value)
        if ($found.Success) { return @{ Key = $key; Index = $found.Index } }
    }
    return $null
}

# 画面に出す共有フォルダのパス。パスの項目は値全体、文の中は、" で囲まれていれば閉じる " まで、なければ空白か " の前まで。
# リンク先以外で見つかったものは、「（引数）」のように項目の名前を付ける
function Get-ShowText([string]$Key, [string]$Text, [int]$Index) {
    $shown = $Text
    if (-not $PathFields.ContainsKey($Key)) {
        $end = $Text.Length
        if ($Index -gt 0 -and $Text[$Index - 1] -eq '"') {
            $quote = $Text.IndexOf('"', $Index)
            if ($quote -ge 0) { $end = $quote }
        } else {
            for ($k = $Index; $k -lt $Text.Length; $k++) {
                if ([char]::IsWhiteSpace($Text[$k]) -or $Text[$k] -eq '"') { $end = $k; break }
            }
        }
        $shown = $Text.Substring($Index, $end - $Index)
    }
    if ($Key -eq 'Target') { return $shown }
    return '（' + $FieldNames[$Key] + '）' + $shown
}

# ショートカットを読み、共有フォルダを指していれば一覧に加える。読めないときは skipped に残す
function Add-Shortcut($Entry) {
    $display = ConvertTo-DisplayPath $Entry.FullName
    if (([int]$Entry.Attributes -band $OnlineMask) -ne 0) {
        Add-Skip 'クラウドのみ' $display 'クラウドのみのファイル（読むとダウンロードが始まるため、調べていません）'
        return
    }
    if ($Entry.Length -gt $MaxLinkBytes) {
        Add-Skip '読めないショートカット' $display '1 MB を超えるため、ショートカットとして読んでいません'
        return
    }
    try {
        $bytes = Read-FileBytes (ConvertTo-LongPath $display)
    } catch {
        Add-SkipError (Get-InnerException $_.Exception) $display
        return
    }
    # WScript.Shell は、形式の違うファイルもエラーにせず空の値を返すため、先頭を自分で確かめる
    $valid = $bytes.Length -ge 0x4C
    for ($k = 0; $k -lt $LinkHeader.Length -and $valid; $k++) { $valid = $bytes[$k] -eq $LinkHeader[$k] }
    if (-not $valid) {
        Add-Skip '読めないショートカット' $display 'ショートカットの形式ではありません'
        return
    }
    $strings = Read-LinkStrings $bytes
    if ($null -eq $strings) {
        Add-Skip '読めないショートカット' $display 'ショートカットの中身を読めません（形式が正しくありません）'
        return
    }
    $short = Test-ShortEnough $display
    try {
        if ($short) {
            $link = $shell.CreateShortcut($display)
        } else {
            [IO.File]::WriteAllBytes($tempLink, $bytes)
            $link = $shell.CreateShortcut($tempLink)
        }
        $target = [string]$link.TargetPath
        $wshArguments = [string]$link.Arguments
        $wshWorkDir = [string]$link.WorkingDirectory
        $wshIcon = Split-IconLocation ([string]$link.IconLocation)
        $wshComment = [string]$link.Description
    } catch {
        Add-Skip '読めないショートカット' $display (Get-ErrorText (Get-InnerException $_.Exception))
        return
    }
    # WScript.Shell が読んだ値がファイルの中身と合わないときは、リンク先も正しく読めていないおそれがあるため、調べない。
    # コメントの「@」で始まる値は、WScript.Shell が訳した文に変えて返すので比べない。
    # リンク先がファイルのパスでないもの（リンク先が空）は、WScript.Shell がほかの項目も空で返すことがあるので比べない
    # （項目はファイルから読んだ値を使う）
    $wrongReason = $null
    if ($strings.Comment.Length -gt $MaxLength['Comment']) { $wrongReason = 'コメント' }
    elseif ($strings.WorkDir.Length -gt $MaxLength['WorkDir']) { $wrongReason = '作業フォルダ' }
    elseif ($target -and (-not [string]::Equals($wshArguments, $strings.Arguments, [StringComparison]::Ordinal) -or
        -not [string]::Equals($wshWorkDir, $strings.WorkDir, [StringComparison]::Ordinal) -or
        -not [string]::Equals($wshIcon[0], $strings.IconPath, [StringComparison]::Ordinal) -or $wshIcon[1] -ne $strings.IconIndex -or
        (-not $strings.Comment.StartsWith('@') -and -not [string]::Equals($wshComment, $strings.Comment, [StringComparison]::Ordinal)))) {
        $wrongReason = ''
    }
    if ($null -ne $wrongReason) {
        $detail = if ($wrongReason) { '{0}が {1} 文字を超えるため、Windows の部品（WScript.Shell）で正しく読めません' -f $wrongReason, $MaxLength[$(if ($wrongReason -eq 'コメント') { 'Comment' } else { 'WorkDir' })] }
                  else { 'Windows の部品（WScript.Shell）で読んだ値が、ファイルの中身と合いません（ほかのツールで作った、長すぎる項目があるものなど）' }
        Add-Skip '読めないショートカット' $display $detail
        return
    }
    # サーバーそのもの（\\192.168.1.10）を指すショートカットは、WScript.Shell ではリンク先が空になるため、ファイルから読んだ名前を使う
    $targetFromEnv = $false
    if (-not $target -and $serverOnlyRule.IsMatch($strings.EnvTarget)) {
        $target = $strings.EnvTarget.TrimEnd('\')
        $targetFromEnv = $true
    }
    $script:linkTotal++
    $values = @{ Target = $target; Arguments = $strings.Arguments; WorkDir = $strings.WorkDir; IconPath = $strings.IconPath; Comment = $strings.Comment }
    $servers = Get-ServerList $values
    if ($servers.Count -eq 0) { return }
    # 一覧に出すものだけ、記録を作る（何万件も回るため、共有フォルダを指さないものには作らない）
    $item = [pscustomobject]@{
        Path = $display; Long = (ConvertTo-LongPath $display); Short = $short; TargetFromEnv = $targetFromEnv
        Time = $Entry.LastWriteTime; ReadOnly = (($Entry.Attributes -band [IO.FileAttributes]::ReadOnly) -ne 0)
        Hash = [BitConverter]::ToString($sha256.ComputeHash($bytes))
        Target = $target; Arguments = $strings.Arguments; WorkDir = $strings.WorkDir; IconPath = $strings.IconPath; IconIndex = $strings.IconIndex; Comment = $strings.Comment
        IconText = $(if ($strings.IconPath) { $strings.IconPath + ',' + $strings.IconIndex } else { '' })
        Servers = $null; ServerText = ''; FirstServer = ''; Show = ''; ShowNew = ''
        New = @{}; Changes = [System.Collections.Generic.List[object]]::new()
        Status = ''; Result = ''; Note = ''; Backup = ''; Roots = [System.Collections.Generic.List[string]]::new()
    }
    $item.Servers = $servers
    $item.ServerText = $servers -join ', '
    $item.FirstServer = $servers[0]
    if ($finding) {
        $hit = Find-FirstMatch $item $fromPathRule $fromTextRule
        if ($null -eq $hit) { return }
        $item.Show = Get-ShowText $hit.Key ([string]$item.($hit.Key)) $hit.Index
        if ($replacing) {
            foreach ($key in $FieldKeys) {
                $old = [string]$item.($key)
                $rule = if ($PathFields.ContainsKey($key)) { $fromPathRule } else { $fromTextRule }
                $new = $old
                if ($old) { $new = $rule.Replace($old, $replacement) }
                $item.New[$key] = $new
                if (-not [string]::Equals($old, $new, [StringComparison]::Ordinal)) {
                    $before = $old
                    $after  = $new
                    if ($key -eq 'IconPath') { $before = $old + ',' + $item.IconIndex; $after = $new + ',' + $item.IconIndex }
                    $item.Changes.Add([pscustomobject]@{ Key = $key; Before = $before; After = $after })
                }
            }
            # 置き換えた後の表示は、置き換える前に見つけた位置から取る（それより前の文字は変わらないため）
            $item.ShowNew = Get-ShowText $hit.Key ([string]$item.New[$hit.Key]) $hit.Index
        }
    } else {
        $hit = Find-FirstMatch $item $pathServerRule $textServerRule
        $item.Show = Get-ShowText $hit.Key ([string]$item.($hit.Key)) $hit.Index
    }
    $items.Add($item)
}

# 共有フォルダのパスから、\\サーバー\共有 の部分を取り出す（共有の部分がなければ $null）
function Get-ShareRoot([string]$Unc) {
    $body = $Unc
    if ($body.StartsWith('\\?\UNC\', [StringComparison]::OrdinalIgnoreCase)) { $body = $body.Substring(8) }
    elseif ($body.StartsWith('\\')) { $body = $body.Substring(2) }
    else { return $null }
    $parts = $body.Split('\')
    if ($parts.Count -lt 2 -or -not $parts[0] -or -not $parts[1]) { return $null }
    return '\\' + $parts[0] + '\' + $parts[1]
}

# つながらないときのエラー番号（ネットワーク パスが見つからない・届かない・切断された・時間切れなど）
$UnreachableCodes = 53, 64, 121, 1231, 1232, 1311
# 見つからないときのエラー番号（ファイル・パスが見つからない、共有名が見つからない）
$NotFoundCodes = 2, 3, 67

# 置き換え先の共有フォルダにつながるかを、共有ごとに 1 回だけ確かめる。
# つながらないサーバーは 1 回に 20 秒ほどかかるため、ネットワークのエラーが出たサーバーは、ほかの共有も確かめない
$rootStates  = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
$pathStates  = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
$serverDown  = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::OrdinalIgnoreCase)
function Test-UncPath([string]$UncPath, [bool]$IsRoot) {
    $cache = if ($IsRoot) { $rootStates } else { $pathStates }
    if ($cache.ContainsKey($UncPath)) { return $cache[$UncPath] }
    $server = $UncPath.Substring(2).Split('\')[0]
    if ($serverDown.ContainsKey($server)) {
        $state = [pscustomobject]@{ State = 'Unreachable'; Reason = $serverDown[$server] }
        $cache[$UncPath] = $state
        return $state
    }
    Write-Progress -Activity $ToolTitle -Status '置き換え先を確かめています' -CurrentOperation $UncPath
    # 共有のルートは \\?\UNC\ を付けると「パスが正しくない」になるため、通常の書き方で確かめる
    $checkPath = if ($IsRoot -or $UncPath.Length -lt 260) { $UncPath } else { ConvertTo-LongPath $UncPath }
    $state = $null
    try {
        [void][IO.File]::GetAttributes($checkPath)
        $state = [pscustomobject]@{ State = 'OK'; Reason = '' }
    } catch {
        $ex = Get-InnerException $_.Exception
        $code = Get-ErrorCode $ex
        $reason = (Get-ErrorText $ex).TrimEnd('。', '.')
        if ($ex -is [UnauthorizedAccessException] -or $code -eq 5) {
            $state = [pscustomobject]@{ State = 'Denied'; Reason = $reason }
        } elseif ($ex -is [IO.FileNotFoundException] -or $ex -is [IO.DirectoryNotFoundException] -or $NotFoundCodes -contains $code) {
            $state = [pscustomobject]@{ State = 'NotFound'; Reason = $reason }
        } elseif ($UnreachableCodes -contains $code) {
            $serverDown[$server] = $reason
            $state = [pscustomobject]@{ State = 'Unreachable'; Reason = $reason }
        } elseif ($IsRoot) {
            # ログオンの失敗などは、サーバーにはつながっているので、ほかの共有は確かめる
            $state = [pscustomobject]@{ State = 'Unreachable'; Reason = $reason }
        } else {
            $state = [pscustomobject]@{ State = 'NotFound'; Reason = $reason }
        }
    }
    $cache[$UncPath] = $state
    return $state
}

# サーバーそのもの（\\サーバー）につながるか。どのサーバーにもある IPC$ を調べ、サーバーが答えれば（見つからない・権限なしを含む）つながるとする
function Test-ServerReach([string]$Server) {
    $key = '\\' + $Server
    if ($rootStates.ContainsKey($key)) { return $rootStates[$key] }
    if ($serverDown.ContainsKey($Server)) {
        $state = [pscustomobject]@{ State = 'Unreachable'; Reason = $serverDown[$Server] }
        $rootStates[$key] = $state
        return $state
    }
    Write-Progress -Activity $ToolTitle -Status '置き換え先を確かめています' -CurrentOperation $key
    $state = [pscustomobject]@{ State = 'OK'; Reason = '' }
    try {
        [void][IO.File]::GetAttributes($key + '\IPC$')
    } catch {
        $ex = Get-InnerException $_.Exception
        $code = Get-ErrorCode $ex
        $reason = (Get-ErrorText $ex).TrimEnd('。', '.')
        if ($UnreachableCodes -contains $code) {
            $serverDown[$Server] = $reason
            $state = [pscustomobject]@{ State = 'Unreachable'; Reason = $reason }
        } elseif (-not ($ex -is [UnauthorizedAccessException] -or $ex -is [IO.FileNotFoundException] -or $ex -is [IO.DirectoryNotFoundException] -or (@(2, 3, 5, 67) -contains $code))) {
            $state = [pscustomobject]@{ State = 'Unreachable'; Reason = $reason }
        }
    }
    $rootStates[$key] = $state
    return $state
}

# 置き換えた後の値にある、-To のサーバーの共有フォルダ（\\サーバー\共有）を集める
function Get-NewRoots($Item) {
    $list = New-Object 'System.Collections.Generic.List[string]'
    foreach ($change in $Item.Changes) {
        $value = [string]$Item.New[$change.Key]
        $rule = if ($PathFields.ContainsKey($change.Key)) { $pathServerRule } else { $textServerRule }
        foreach ($found in $rule.Matches($value)) {
            if (-not [string]::Equals($found.Groups[2].Value, $toSpec.Server, [StringComparison]::OrdinalIgnoreCase)) { continue }
            $text = if ($PathFields.ContainsKey($change.Key)) { $value } else { (Get-ShowText $change.Key $value $found.Index).Substring($FieldNames[$change.Key].Length + 2) }
            $root = Get-ShareRoot $text
            if ($root -and -not $list.Contains($root)) { $list.Add($root) }
        }
    }
    return , $list
}

# 5.1 は、Ctrl+C（パイプラインの停止）を取りこぼして動き続けることがある。
# 置き換えるときは、コンソールで動かしている間、Ctrl+C をキーの入力として受け取り、ショートカットの区切りごとに確かめて止める。
# 止めたあとの片付けと結果の保存は、finally で行う
$ctrlCSaved = $null   # 変える前の TreatControlCAsInput（変えていないときは $null）
function Enable-CtrlCInput {
    if (-not $replacing -or $null -ne $script:ctrlCSaved) { return }
    try {
        if ([Console]::IsInputRedirected) { return }
        $saved = [Console]::TreatControlCAsInput
        [Console]::TreatControlCAsInput = $true
        $script:ctrlCSaved = $saved
    } catch { }   # コンソールがないとき（タスクスケジューラなど）は、通常の Ctrl+C のまま
}

function Disable-CtrlCInput {
    if ($null -eq $script:ctrlCSaved) { return }
    try { [Console]::TreatControlCAsInput = $script:ctrlCSaved } catch { }
    $script:ctrlCSaved = $null
}

function Stop-IfCtrlC {
    if ($null -eq $script:ctrlCSaved) { return }
    $pressed = $false
    try {
        while ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::C -and ($key.Modifiers -band [ConsoleModifiers]::Control)) { $pressed = $true }
        }
    } catch { }
    if ($pressed) { exit 1 }
}

# 控えるときの、結果のフォルダから見たパス（backup\0001_営業共有.lnk）。元のフォルダの構成をまねると、
# 結果のフォルダの名前と合わせて 260 文字を超え、5.1 の Copy-Item で戻せなくなるため、番号と名前だけにする
function Get-BackupRelative([int]$Number, [string]$Display) {
    return 'backup\' + $Number.ToString('D4', $Invariant) + '_' + $Display.Substring($Display.LastIndexOf('\') + 1)
}

# 書き換えるショートカットを、Shell.Application で開く
function Get-ShellLink([string]$Display) {
    $folder = $shellApp.NameSpace([IO.Path]::GetDirectoryName($Display))
    if ($null -eq $folder) { throw 'ショートカットのあるフォルダを開けません（Shell.Application）' }
    $item = $folder.ParseName([IO.Path]::GetFileName($Display))
    if ($null -eq $item -or -not $item.IsLink) { throw 'ショートカットとして開けません（Shell.Application）' }
    return $item.GetLink
}

# 書き換えた後に読み直した値が、予定どおりか。パスは大文字と小文字を区別しない（Windows が共有名の大文字・小文字を変えて返すことがあるため）
function Test-SameValue([string]$Expected, [string]$Actual, [bool]$IsPath) {
    if ($IsPath) { return [string]::Equals($Expected.TrimEnd('\'), $Actual.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) }
    return [string]::Equals($Expected, $Actual, [StringComparison]::Ordinal)
}

# ファイルを読む。書き換えた直後のショートカットは、ウイルス対策ソフトなどが削除の権限付きで開いていることがあり、
# ReadAllBytes（読み取りだけを共有）では「ほかのプロセスが使っている」になるため、書き込みと削除も共有して開く
function Read-FileBytes([string]$LongPath) {
    $stream = New-Object IO.FileStream($LongPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]'ReadWrite, Delete'))
    try {
        $buffer = New-Object 'byte[]' $stream.Length
        $count = 0
        while ($count -lt $buffer.Length) {
            $n = $stream.Read($buffer, $count, $buffer.Length - $count)
            if ($n -le 0) { break }
            $count += $n
        }
        if ($count -ne $buffer.Length) { throw '読んでいる途中でファイルの大きさが変わりました' }
        return , $buffer
    } finally {
        $stream.Dispose()
    }
}

function Test-SameBytes([byte[]]$A, [byte[]]$B) {
    if ($A.Length -ne $B.Length) { return $false }
    for ($k = 0; $k -lt $A.Length; $k++) { if ($A[$k] -ne $B[$k]) { return $false } }
    return $true
}

# ファイルを元の内容に戻す（属性とアクセス権を変えないよう、開いて書き直す）
function Restore-Bytes([string]$LongPath, [byte[]]$Bytes) {
    $stream = New-Object IO.FileStream($LongPath, [IO.FileMode]::Open, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $stream.Write($Bytes, 0, $Bytes.Length)
        $stream.SetLength($Bytes.Length)
        $stream.Flush()
    } finally {
        $stream.Dispose()
    }
}

# スキップ・リンク・除外の件数の行
function Get-NoteLine {
    $notes = @()
    $total = 0
    foreach ($value in $skipKinds.Values) { $total += $value }
    $kinds = @(foreach ($name in $skipKinds.Keys) { if ($skipKinds[$name] -gt 0) { , @($name, $skipKinds[$name]) } })
    if ($kinds.Count -eq 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $kinds[0][1] $kinds[0][0]
    } elseif ($kinds.Count -gt 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $total (($kinds | ForEach-Object { Format-Invariant '{0} {1:N0}' $_[0] $_[1] }) -join '・')
    }
    if ($linkCount -gt 0) { $notes += Format-Invariant 'たどらなかったリンク {0:N0} 件' $linkCount }
    if ($excludedCount -gt 0) { $notes += Format-Invariant '除外 {0:N0} 件' $excludedCount }
    return ($notes -join '  ')
}

# 文字の番号の順に並べる（$Keys と同じ順に $Values を並べ替える）。
# [Array] に変えて渡さないと、PowerShell が型を合わせた写しを作って並べ替え、$Values が並ばないことがある
function Get-SortedList([string[]]$Keys, [object[]]$Values) {
    if ($Values.Count -le 1) { return , $Values }
    [Array]::Sort([Array]$Keys, [Array]$Values, [Collections.IComparer][StringComparer]::Ordinal)
    return , $Values
}

#endregion

# 調べている途中や置き換えの途中で Ctrl+C されたときに、どこまで進んだかを残せるようにする
$finished      = $false
$unexpected    = $false
$saveError     = $null
$replaceStarted = $false
$answeredNo    = $false
$currentItem   = $null
$sortedItems   = @()
$lines         = New-Object 'System.Collections.Generic.List[string]'   # summary.txt に書く行

function Save-Result {
    if (-not $runDir) { return }
    try {
        # 保存先は \\?\ 付きで作っているため、長いパスでも書けるよう、ここも \\?\ を付ける
        [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
        $rows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($entry in $sortedItems) {
            $rows.Add([object[]]@($entry.Path, (Protect-ExcelText $entry.ServerText), (Protect-ExcelText $entry.Target), (Protect-ExcelText $entry.Arguments),
                (Protect-ExcelText $entry.WorkDir), (Protect-ExcelText $entry.IconText), (Protect-ExcelText $entry.Comment),
                $entry.Time.ToString($TimeFormat, $Invariant), $entry.Result, $entry.Note, $entry.Backup))
        }
        Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'shortcuts.csv')) $ShortcutColumns $rows
        if ($replacing) {
            $changeRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($entry in $sortedItems) {
                foreach ($change in $entry.Changes) {
                    $changeRows.Add([object[]]@($entry.Path, $FieldNames[$change.Key], (Protect-ExcelText $change.Before), (Protect-ExcelText $change.After), $entry.Result))
                }
            }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'changes.csv')) $ChangeColumns $changeRows
        }
        Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $skipped
    } catch {
        $script:saveError = Get-ErrorText (Get-InnerException $_.Exception)
    }
}

# 行を summary.txt に足し、画面にも出す（■ は見出しの色、※ は注意の色）
function Write-Line([string]$Text) {
    $lines.Add($Text)
    if ($Text.StartsWith('■')) { Write-Host $Text -ForegroundColor Cyan }
    elseif ($Text.StartsWith('※')) { Write-Host $Text -ForegroundColor Yellow }
    else { Write-Host $Text }
}

# 一覧の表を足す。保存するときは、画面には先頭の分だけを出す（保存しないときは、ほかに見る場所がないためすべて）
function Write-Table([string]$Header, [string[]]$Rows) {
    Write-Line $Header
    for ($k = 0; $k -lt $Rows.Count; $k++) {
        $lines.Add($Rows[$k])
        if ($k -lt $ScreenRows -or -not $runDir) { Write-Host $Rows[$k] }
        elseif ($k -eq $ScreenRows) { Write-Host (Format-Invariant '（ほか {0:N0} 件は shortcuts.csv を参照してください）' ($Rows.Count - $ScreenRows)) }
    }
}

# 列の幅（値の最大。上限を超えるものは、その列からはみ出させる）
function Get-ColumnWidth([string[]]$Values, [int]$Min, [int]$Max) {
    $width = $Min
    foreach ($value in $Values) { $width = [math]::Max($width, (Get-DisplayWidth $value)) }
    return [math]::Min($width, $Max)
}

try {
    #region 調べる

    Enable-CtrlCInput
    foreach ($root in $roots) {
        if ($root.IsFile) {
            Stop-IfCtrlC
            Add-Shortcut (New-Object IO.FileInfo $root.Long)
            continue
        }
        $prefixLen = $root.Long.TrimEnd('\').Length + 1
        # 再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる
        $stack = New-Object 'System.Collections.Generic.Stack[object]'
        $stack.Push((New-Object IO.DirectoryInfo $root.Long))
        while ($stack.Count -gt 0) {
            Stop-IfCtrlC
            $dir = $stack.Pop()
            if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                Write-Progress -Activity $ToolTitle -Status (Format-Invariant '調べています: ショートカット {0:N0} 件' $linkTotal) -CurrentOperation (ConvertTo-DisplayPath $dir.FullName)
                $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
            }
            try {
                $entries = $dir.GetFileSystemInfos()
            } catch {
                $ex = Get-InnerException $_.Exception
                # 種類を判定できないリンクは、読み取りが禁止されていると普通のフォルダに見える。権限なしには数えない
                if ($ex -is [UnauthorizedAccessException] -and ($dir.Attributes -band $rpFlag) -eq $rpFlag) {
                    $linkCount++
                    Add-Skip 'リンク' (ConvertTo-DisplayPath $dir.FullName) '読み取りが禁止されたリンク'
                    continue
                }
                Add-SkipError $ex (ConvertTo-DisplayPath $dir.FullName)
                continue
            }
            foreach ($entry in $entries) {
                $isDir = $entry -is $dirType
                if (-not $isDir -and -not $entry.Name.EndsWith('.lnk', [StringComparison]::OrdinalIgnoreCase)) { continue }
                # 前回の結果（控えたショートカット）を置き換えないよう、このツールの結果のフォルダは中を見ない
                if ($isDir -and $OutputNameRule.IsMatch($entry.Name)) {
                    $outputDisplay = ConvertTo-DisplayPath $entry.FullName
                    if (-not [string]::Equals($outputDisplay, $runDir, [StringComparison]::OrdinalIgnoreCase)) {
                        $excludedCount++
                        Add-Skip '除外指定' $outputDisplay 'このツールの結果のフォルダ'
                    }
                    continue
                }
                # ごみ箱の中身は置き換えても意味がなく、System Volume Information はふつう読めないため、調べない
                if ($isDir -and $SystemFolders -contains $entry.Name) {
                    $excludedCount++
                    Add-Skip '除外指定' (ConvertTo-DisplayPath $entry.FullName) 'ごみ箱・システムの管理用のフォルダ'
                    continue
                }
                if ($hasExclude) {
                    $matched = Get-ExcludeMatch $entry.FullName.Substring($prefixLen) $entry.Name (ConvertTo-DisplayPath $entry.FullName)
                    if ($matched) {
                        $excludedCount++
                        Add-Skip '除外指定' (ConvertTo-DisplayPath $entry.FullName) (Protect-ExcelText $matched)
                        continue
                    }
                }
                if (($entry.Attributes -band $rpFlag) -eq $rpFlag) {
                    $linkName = Get-LinkName $entry
                    if ($linkName) {
                        $linkCount++
                        Add-Skip 'リンク' (ConvertTo-DisplayPath $entry.FullName) $linkName
                        continue
                    }
                }
                if ($isDir) {
                    $stack.Push($entry)
                } else {
                    Stop-IfCtrlC
                    Add-Shortcut $entry
                }
            }
        }
    }

    #endregion

    #region 置き換えの予定

    if ($replacing) {
        foreach ($entry in $items) {
            Stop-IfCtrlC
            # 書き込めないものは、置き換え先を確かめる前に分ける（つながらないサーバーを待たないため）
            if (-not $entry.Short) {
                $entry.Status = $StCannot
                $entry.Note = 'パスが長すぎるため、置き換えられません（ショートカットを書き換える Windows の部品が、260 文字以上のパスを扱えないため）'
                continue
            }
            if ($entry.ReadOnly) {
                $entry.Status = $StCannot
                $entry.Note = '読み取り専用のファイルです'
                continue
            }
            # 置き換えると WScript.Shell が扱える長さを超える項目（書くと切られる、読むとずれる）
            $tooLong = @(foreach ($change in $entry.Changes) {
                if (([string]$entry.New[$change.Key]).Length -gt $MaxLength[$change.Key]) { $FieldNames[$change.Key] }
            })
            if ($tooLong.Count -gt 0) {
                $entry.Status = $StCannot
                $entry.Note = '置き換えると{0}が長くなりすぎるため、置き換えられません（Windows の部品（WScript.Shell）が扱える長さを超えるため）' -f ($tooLong -join '・')
                continue
            }
            # WScript.Shell は、\\?\UNC\ の形のリンク先を受け付けない
            if (([string]$entry.New['Target']).StartsWith('\\?\') -and -not [string]::Equals([string]$entry.New['Target'], $entry.Target, [StringComparison]::Ordinal)) {
                $entry.Status = $StCannot
                $entry.Note = 'リンク先が \\?\UNC\ の形のため、置き換えられません（Windows の部品（WScript.Shell）が受け付けないため）'
                continue
            }
            try {
                $probe = New-Object IO.FileStream($entry.Long, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, ([IO.FileShare]'ReadWrite, Delete'))
                $probe.Dispose()
            } catch {
                $ex = Get-InnerException $_.Exception
                $entry.Status = $StCannot
                $entry.Note = if ($ex -is [UnauthorizedAccessException]) { '書き込む権限がありません' } else { 'ファイルを書き込み用に開けません: ' + (Get-ErrorText $ex) }
                continue
            }
            # 置き換え先の共有フォルダにつながるか
            $status = $StPlan
            $note = ''
            $denied = $false
            foreach ($root in (Get-NewRoots $entry)) {
                $entry.Roots.Add($root)
                $state = Test-UncPath $root $true
                if ($state.State -eq 'Unreachable') { $status = $StUnreach; $note = '{0} に接続できません（{1}）' -f $root, $state.Reason; break }
                if ($state.State -eq 'NotFound' -and $status -eq $StPlan) { $status = $StNotFound; $note = '共有フォルダ {0} が見つかりません' -f $root }
                if ($state.State -eq 'Denied') { $denied = $true }
            }
            # リンク先が共有フォルダのパスなら、そのフォルダかファイルがあるか
            $newTarget = [string]$entry.New['Target']
            $targetChanged = -not [string]::Equals($newTarget, $entry.Target, [StringComparison]::Ordinal)
            if ($status -eq $StPlan -and $targetChanged -and -not $denied -and (Get-ShareRoot $newTarget)) {
                $uncTarget = $newTarget
                if ($uncTarget.StartsWith('\\?\UNC\', [StringComparison]::OrdinalIgnoreCase)) { $uncTarget = '\\' + $uncTarget.Substring(8) }
                $state = Test-UncPath $uncTarget $false
                switch ($state.State) {
                    'NotFound' { $status = $StNotFound; $note = '{0} が見つかりません' -f $uncTarget }
                    'Unreachable' { $status = $StUnreach; $note = '{0} に接続できません（{1}）' -f $uncTarget, $state.Reason }
                    'Denied' { $denied = $true }
                }
            } elseif ($status -eq $StPlan -and $targetChanged -and $serverOnlyRule.IsMatch($newTarget)) {
                # サーバーそのものを指すショートカットは、サーバーにつながるかだけを確かめる
                $server = $newTarget.Substring(2).TrimEnd('\')
                $entry.Roots.Add('\\' + $server)
                $state = Test-ServerReach $server
                if ($state.State -eq 'Unreachable') { $status = $StUnreach; $note = '\\{0} に接続できません（{1}）' -f $server, $state.Reason }
            }
            if ($status -eq $StPlan -and $denied) { $note = '権限がないため、置き換え先を確かめられませんでした' }
            $entry.Status = $status
            $entry.Note = $note
        }
        Write-Progress -Activity $ToolTitle -Completed
    }

    #endregion

    #region 一覧の表示

    $scanElapsed = $sw.Elapsed
    Write-Progress -Activity $ToolTitle -Completed

    # 並べる。棚卸しと -From だけのときはサーバー・パスの順、置き換えるときは予定・パスの順
    $itemArray = $items.ToArray()
    $keys = New-Object 'string[]' $itemArray.Count
    for ($k = 0; $k -lt $itemArray.Count; $k++) {
        $entry = $itemArray[$k]
        if ($replacing) { $keys[$k] = [string]$StatusOrder[$entry.Status] + "`0" + (Get-SortText $entry.Path) }
        else { $keys[$k] = (Get-SortText $entry.FirstServer) + "`0" + (Get-SortText $entry.Path) }
    }
    $sortedItems = Get-SortedList $keys $itemArray

    Write-Line (Format-Invariant '{0} (Update-ShortcutTarget {1})' $ToolTitle $ScriptVersion)
    Write-Line ('対象: {0}' -f (($roots | ForEach-Object { $_.Full }) -join ', '))
    Write-Line (Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（管理者: {2}）' $startTime $machineName $(if ($isAdmin) { 'はい' } else { 'いいえ' }))
    if ($replacing) { Write-Line ('置き換え: {0} → \\{1}' -f (($fromSpecs | ForEach-Object { '\\' + $_.Body }) -join ', '), $toSpec.Body) }
    if ($hasExclude) { Write-Line ('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', ')) }
    $fromLabel = ($fromSpecs | ForEach-Object { $_.Label }) -join ', '
    if ($replacing) {
        Write-Line (Format-Invariant 'ショートカット {0:N0} 件  置き換える前のサーバーを指すもの {1:N0} 件  所要 {2}' $linkTotal $sortedItems.Count (Format-Duration $scanElapsed))
    } elseif ($finding) {
        Write-Line (Format-Invariant 'ショートカット {0:N0} 件  {1} を指すもの {2:N0} 件  所要 {3}' $linkTotal $fromLabel $sortedItems.Count (Format-Duration $scanElapsed))
    } else {
        Write-Line (Format-Invariant 'ショートカット {0:N0} 件  共有フォルダを指すもの {1:N0} 件  所要 {2}' $linkTotal $sortedItems.Count (Format-Duration $scanElapsed))
    }
    $noteLine = Get-NoteLine
    if ($noteLine) { Write-Line $noteLine }
    Write-Line ''

    if (-not $finding) {
        # 棚卸し：サーバーごとの件数（件数の多い順、同じ件数はサーバー名の順）
        $counts = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $sortedItems) {
            foreach ($name in $entry.Servers) {
                if ($counts.ContainsKey($name)) { $counts[$name].Count++ } else { $counts[$name] = [pscustomobject]@{ Name = $name; Count = 1 } }
            }
        }
        Write-Line '■ サーバーごとの件数'
        if ($counts.Count -eq 0) {
            Write-Line '    （共有フォルダを指すショートカットはありません）'
        } else {
            $countArray = @($counts.Values)
            $countKeys = @(foreach ($entry in $countArray) { (999999999 - $entry.Count).ToString('D9', $Invariant) + "`0" + (Get-SortText $entry.Name) })
            $countArray = Get-SortedList $countKeys $countArray
            $widthServer = Get-ColumnWidth @($countArray | ForEach-Object { $_.Name }) 8 $ServerWidth
            Write-Line (' ' + (Format-Column 'サーバー' $widthServer $false) + '  ' + (Format-Column '件数' 6 $true))
            foreach ($entry in $countArray) { Write-Line (' ' + (Format-Column $entry.Name $widthServer $false) + '  ' + (Format-Column ($entry.Count.ToString('N0', $Invariant)) 6 $true)) }
            Write-Line ''
            $widthServer = Get-ColumnWidth @($sortedItems | ForEach-Object { $_.FirstServer }) 8 $ServerWidth
            $widthShow = Get-ColumnWidth @($sortedItems | ForEach-Object { $_.Show }) 18 $PathWidth
            $rows = @(foreach ($entry in $sortedItems) { ' ' + (Format-Column $entry.FirstServer $widthServer $false) + '  ' + (Format-Column $entry.Show $widthShow $false) + '  ' + $entry.Path })
            Write-Line '■ 共有フォルダを指すショートカット（サーバーの順）'
            Write-Table (' ' + (Format-Column 'サーバー' $widthServer $false) + '  ' + (Format-Column '共有フォルダのパス' $widthShow $false) + '  ショートカット') $rows
        }
    } elseif (-not $replacing) {
        Write-Line ('■ {0} を指すショートカット' -f $fromLabel)
        if ($sortedItems.Count -eq 0) {
            Write-Line ('    （{0} を指すショートカットはありません）' -f $fromLabel)
        } else {
            $widthShow = Get-ColumnWidth @($sortedItems | ForEach-Object { $_.Show }) 8 $PathWidth
            $rows = @(foreach ($entry in $sortedItems) { ' ' + (Format-Column $entry.Show $widthShow $false) + '  ' + $entry.Path })
            Write-Table (' ' + (Format-Column 'リンク先' $widthShow $false) + '  ショートカット') $rows
        }
    } else {
        Write-Line '■ 置き換えの予定'
        if ($sortedItems.Count -eq 0) {
            Write-Line '    （置き換える前のサーバーを指すショートカットはありません）'
        } else {
            $widthStatus = Get-ColumnWidth @($sortedItems | ForEach-Object { $_.Status }) 4 30
            $widthShow = Get-ColumnWidth @($sortedItems | ForEach-Object { $_.ShowNew }) 22 $PathWidth
            $rows = @(foreach ($entry in $sortedItems) { ' ' + (Format-Column $entry.Status $widthStatus $false) + '  ' + (Format-Column $entry.ShowNew $widthShow $false) + '  ' + $entry.Path })
            Write-Table (' ' + (Format-Column '予定' $widthStatus $false) + '  ' + (Format-Column '置き換えた後のリンク先' $widthShow $false) + '  ショートカット') $rows
        }
    }

    #endregion

    #region 置き換え

    if ($replacing) {
        $plan = @($sortedItems | Where-Object { $_.Status -eq $StPlan })
        $restKinds = [ordered]@{ $StNotFound = 0; $StUnreach = 0; $StCannot = 0 }
        foreach ($entry in $sortedItems) { if ($restKinds.Contains($entry.Status)) { $restKinds[$entry.Status]++ } }
        $restCount = $sortedItems.Count - $plan.Count
        foreach ($entry in $sortedItems) { if ($entry.Status -ne $StPlan) { $entry.Result = $entry.Status } }

        Write-Line ''
        if ($plan.Count -gt 0) { Write-Line (Format-Invariant '置き換えるショートカット {0:N0} 件' $plan.Count) }
        if ($restCount -gt 0) {
            $restParts = @(foreach ($name in $restKinds.Keys) { if ($restKinds[$name] -gt 0) { Format-Invariant '{0} {1:N0} 件' $name $restKinds[$name] } })
            Write-Line (Format-Invariant '置き換えないショートカット {0:N0} 件（{1}）' $restCount ($restParts -join '・'))
        }
        # つながらなかった共有フォルダ（同じサーバーは、最初に見つかった 1 つだけ）
        $shownServers = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $sortedItems) {
            if ($entry.Status -ne $StUnreach) { continue }
            foreach ($root in $entry.Roots) {
                $state = $rootStates[$root]
                if ($state.State -eq 'Unreachable' -and $shownServers.Add($root.Substring(2).Split('\')[0])) {
                    Write-Line ('※ {0} に接続できません（{1}）。新しいサーバーが動いているか、サーバーの名前や IP アドレスが正しいかを確かめてください' -f $root, $state.Reason)
                }
            }
        }
        # 共有フォルダそのものが見つからないときは、共有名が変わったことが多い
        $shareMissing = $false
        foreach ($state in $rootStates.Values) { if ($state.State -eq 'NotFound') { $shareMissing = $true } }
        if (-not $toSpec.IsShare -and $shareMissing) { Write-Line '※ 共有名が変わったときは、-From と -To を \\サーバー\共有 の形で指定してください' }
        if (-not $isAdmin -and @($sortedItems | Where-Object { $_.Note -eq '書き込む権限がありません' }).Count -gt 0) {
            Write-Line '※ 管理者として実行すると、置き換えられるものがあります（パブリックのデスクトップなど）'
        }
        if ($toInsideFrom -and $plan.Count -gt 0) { Write-Line ('※ -To が -From の中にあります。もう一度実行すると、二重に置き換わります（{0}）' -f $toInsideFrom) }

        if ($plan.Count -eq 0) {
            Write-Line '置き換えるショートカットはありません'
        } else {
            Write-Line '元のファイルは、保存先の backup フォルダに控えます'
            $answer = 'Y'
            if (-not $Force) {
                # 画面のない実行（タスクスケジューラの「ログオンしているかどうかにかかわらず実行する」など）で、確認待ちのまま止まり続けないようにする
                $cannotAsk = '確認できないため、置き換えずに止めました。確認せずに置き換えるときは -Force を付けてください'
                if (-not [Environment]::UserInteractive) { Stop-WithError $cannotAsk }
                # 確認を待つ間は、通常の Ctrl+C に戻す（キーの入力として受け取ると、Read-Host に Ctrl+C の文字が入るため）
                Stop-IfCtrlC
                Disable-CtrlCInput
                $answer = ''
                while ($answer -notmatch '^[YyNnＹｙＮｎ]') {
                    # -NonInteractive で起動されていると、Read-Host は英語のエラーで止まる
                    try { $answer = Read-Host '置き換えますか？ [Y] はい  [N] いいえ' } catch { Stop-WithError $cannotAsk }
                    if ($null -eq $answer) { Stop-WithError $cannotAsk }   # 入力が終わっている（リダイレクトされた入力の終わり）
                    $answer = $answer.Trim()
                }
            }
            if ($answer -match '^[NnＮｎ]') {
                $answeredNo = $true
                foreach ($entry in $plan) { $entry.Result = $StNo }
                Write-Line '置き換えませんでした'
            } else {
                Enable-CtrlCInput
                $replaceStarted = $true
                $replaceWatch = [Diagnostics.Stopwatch]::StartNew()
                $doneCount = 0
                $failCount = 0
                $backupNumber = 0
                foreach ($entry in $plan) {
                    # 押されていたら、ここ（前のショートカットを書き終えた後）で止める
                    Stop-IfCtrlC
                    if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                        Write-Progress -Activity $ToolTitle -Status (Format-Invariant '置き換えています: {0:N0} / {1:N0} 件' ($doneCount + $failCount) $plan.Count) -CurrentOperation $entry.Path
                        $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
                    }
                    try {
                        $bytes = Read-FileBytes $entry.Long
                    } catch {
                        $entry.Result = $StFail
                        $entry.Note = '読めないため、置き換えていません: ' + (Get-ErrorText (Get-InnerException $_.Exception))
                        $failCount++
                        continue
                    }
                    # 予定を立てた後（確認を待っている間など）に変わったショートカットは、置き換えない
                    if ([BitConverter]::ToString($sha256.ComputeHash($bytes)) -ne $entry.Hash) {
                        $entry.Result = $StFail
                        $entry.Note = '調べた後に変更されたため、置き換えていません'
                        $failCount++
                        continue
                    }
                    # 元のファイルを控える。控えられなければ置き換えない
                    $backupNumber++
                    $backupRelative = Get-BackupRelative $backupNumber $entry.Path
                    $backupPath = Join-Path $runDir $backupRelative
                    try {
                        [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath (Join-Path $runDir 'backup')))
                        [IO.File]::WriteAllBytes((ConvertTo-LongPath $backupPath), $bytes)
                    } catch {
                        $entry.Result = $StFail
                        $entry.Note = '元のファイルを控えられないため、置き換えていません: ' + (Get-ErrorText (Get-InnerException $_.Exception))
                        $failCount++
                        continue
                    }
                    $entry.Backup = $backupRelative
                    # 書き込みの途中で止まったときに、そのショートカットを「失敗」と記録できるようにしておく
                    $currentItem = $entry
                    try {
                        $link = Get-ShellLink $entry.Path
                        foreach ($change in $entry.Changes) {
                            switch ($change.Key) {
                                'Target' { $link.Path = [string]$entry.New['Target'] }
                                'Arguments' { $link.Arguments = [string]$entry.New['Arguments'] }
                                'WorkDir' { $link.WorkingDirectory = [string]$entry.New['WorkDir'] }
                                'IconPath' { $link.SetIconLocation([string]$entry.New['IconPath'], [int]$entry.IconIndex) }
                                'Comment' { $link.Description = [string]$entry.New['Comment'] }
                            }
                        }
                        $link.Save()
                        # 読み直して、5 つの項目が予定どおりかを確かめる。リンク先は WScript.Shell で、
                        # ほかの項目はファイルから読む（WScript.Shell はコメントを訳した文に変えて返し、長い値では項目がずれるため）
                        $checkStrings = Read-LinkStrings (Read-FileBytes $entry.Long)
                        if ($null -eq $checkStrings) { throw '置き換えた後のショートカットを読めません（形式が正しくありません）' }
                        $check = $shell.CreateShortcut($entry.Path)
                        $actualTarget = if ($entry.TargetFromEnv) { $checkStrings.EnvTarget.TrimEnd('\') } else { [string]$check.TargetPath }
                        $actual = @{ Target = $actualTarget; Arguments = $checkStrings.Arguments; WorkDir = $checkStrings.WorkDir; IconPath = $checkStrings.IconPath; Comment = $checkStrings.Comment }
                        $wrong = @(foreach ($key in $FieldKeys) {
                            if (-not (Test-SameValue ([string]$entry.New[$key]) $actual[$key] $PathFields.ContainsKey($key))) { $FieldNames[$key] }
                        })
                        if ($checkStrings.IconIndex -ne $entry.IconIndex -and $wrong -notcontains $FieldNames['IconPath']) { $wrong += $FieldNames['IconPath'] }
                        if ($wrong.Count -gt 0) { throw ('読み直した{0}が、置き換えた内容と一致しません' -f ($wrong -join '・')) }
                        # 備考は予定のまま残す（「権限がないため、置き換え先を確かめられませんでした」など）
                        $entry.Result = $StDone
                        $doneCount++
                    } catch {
                        $reason = if ($_.Exception -is [Management.Automation.RuntimeException] -and -not $_.Exception.InnerException) { $_.Exception.Message } else { Get-ErrorText (Get-InnerException $_.Exception) }
                        # 元の内容に戻す。書く前に失敗して、ファイルが元のままなら戻さない
                        # （ほかのプログラムが開いていて書けなかったときに、「戻せませんでした」と出さないため）
                        $unchanged = $false
                        try { $unchanged = Test-SameBytes (Read-FileBytes $entry.Long) $bytes } catch { }
                        if ($unchanged) {
                            $restored = '（ファイルは変わっていません）'
                        } else {
                            $restored = '（元に戻しました）'
                            try { Restore-Bytes $entry.Long $bytes } catch { $restored = '（元に戻せませんでした。backup のファイルから戻してください）' }
                        }
                        $entry.Result = $StFail
                        $entry.Note = $reason + $restored
                        $failCount++
                    }
                    $currentItem = $null
                }
                Write-Progress -Activity $ToolTitle -Completed
                $resultLine = if ($failCount -eq 0) { '結果: 置き換えました  {0:N0} 件  失敗 0 件  所要 {1}' } else { '結果: 置き換えられなかったショートカットがあります  置き換え {0:N0} 件  失敗 {2:N0} 件  所要 {1}（shortcuts.csv を確認してください）' }
                Write-Line ''
                Write-Line (Format-Invariant $resultLine $doneCount (Format-Duration $replaceWatch.Elapsed) $failCount)
            }
        }
    }

    #endregion

    if ($runDir) { $lines.Add(('保存先: {0}' -f $runDir)) }
    Save-Result
    $finished = $true
} catch {
    # 想定外のエラーは、Ctrl+C の「中断しました」と区別して知らせる
    $unexpected = $true
    throw
} finally {
    Disable-CtrlCInput
    Write-Progress -Activity $ToolTitle -Completed
    try { if ([IO.File]::Exists($tempLink)) { [IO.File]::Delete($tempLink) } } catch { }
    if (-not $finished -and $stopping) {
        # 止めた理由は、Stop-WithError が表示している
        if ($runDir -and -not $replaceStarted) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
    } elseif (-not $finished) {
        if ($replaceStarted) {
            # 置き換えの途中で止まったときは、置き換えたショートカットと控えが残っているため、結果も残す
            # 書き込みの途中だったショートカットは、壊れているおそれがあるため「失敗」にする
            if ($currentItem) {
                $currentItem.Result = $StFail
                $currentItem.Note = '書き込みの途中で止まりました。backup のファイルから戻してください'
            }
            foreach ($entry in $sortedItems) {
                if ($entry.Status -eq $StPlan -and $entry.Result -eq '') { $entry.Result = $StNo; $entry.Note = '中断したため、置き換えていません' }
            }
            $stopText = if ($unexpected) { 'エラーで止まりました' } else { '中断しました' }
            $lines.Add(('{0}。それまでに置き換えたショートカットと、控えたファイルは残っています' -f $stopText))
            Save-Result
            Write-Host ('{0}。それまでに置き換えたショートカットと、控えたファイルは残っています: {1}' -f $stopText, $runDir) -ForegroundColor Yellow
        } else {
            if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
            if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
            else { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
        }
    }
}

if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
} elseif ($saveWarning) {
    Write-Warning $saveWarning
}

# 確認で N を選んだときは、置き換えられないショートカットがあっても 0
if ($replacing -and -not $answeredNo) {
    $failed = @($sortedItems | Where-Object { $_.Result -eq $StNotFound -or $_.Result -eq $StUnreach -or $_.Result -eq $StCannot -or $_.Result -eq $StFail }).Count
    if ($failed -gt 0) { exit 1 }
}
exit 0
