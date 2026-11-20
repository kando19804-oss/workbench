<#
.SYNOPSIS
    指定した日数より古いファイルを一覧にし、確認したうえで、まとめて削除するか、別の場所へ移します。

.DESCRIPTION
    指定したフォルダの中（サブフォルダを含む）から、作成日時と最終更新日時がどちらも指定した日数より前のファイルを探し、
    フォルダごとの件数と容量と一緒に一覧にします。

    -Delete を指定すると、削除の予定を表示して確認を求め、Y を入力すると削除します。削除したファイルは、ごみ箱に入りません。
    -MoveTo を指定すると、指定したフォルダへ、同じフォルダの構成で移します。
    どちらも指定しないときは、一覧にするだけで、ファイルを変更しません。何を削除したか（移したか）は、結果のフォルダに残します。

    読み取り専用・システム属性・クラウドのみ（OneDrive など）のファイルは、古くても削除しません。
    ジャンクションやシンボリックリンクはたどりません。権限がなくて読めないフォルダはスキップし、件数を表示します。
    ドライブのルートや Windows のフォルダなど、OS やアプリが動くのに必要な場所は指定できません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Remove-OldFile.ps1 -Path D:\logs -OlderThan 90

.PARAMETER Path
    調べるフォルダ。サブフォルダも含めて調べます。複数指定するときは、カンマで区切って 1 つの文字列にします。必ず指定します。

.PARAMETER OlderThan
    何日より古いファイルを対象にするか（1～36500）。作成日時と最終更新日時が、どちらも実行した日時のこの日数前より前のファイルを対象にします。
    必ず指定します。

.PARAMETER Include
    対象にするファイルの名前。ワイルドカード（*）で、カンマで区切って書きます（例: "*.log,*.txt"）。省略すると、すべてのファイルです。

.PARAMETER Exclude
    対象にしないフォルダやファイル。ワイルドカード（*）で、カンマで区切って書きます。
    \ を含まない書き方は名前に、ドライブや \\ で始まる書き方はフルパスに、それ以外は -Path から見た相対パスに当てはめます。

.PARAMETER Delete
    対象のファイルを、確認してから削除します。削除したファイルは、ごみ箱に入りません。

.PARAMETER MoveTo
    対象のファイルを、確認してから指定したフォルダへ移します。フォルダは先に作っておきます。-Delete とは一緒に使えません。

.PARAMETER RemoveEmptyFolder
    対象のファイルを削除した（移した）ことで空になったフォルダも削除します。-Path のフォルダそのものは削除しません。

.PARAMETER Force
    確認せずに削除します（移します）。-Delete か -MoveTo と一緒に使います。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。削除するとき・移すときは使えません。

.PARAMETER _UnknownArgument
    指定しません。どのパラメーターにも当たらない引数（パラメーターの名前の書き間違いなど）を受け取り、
    案内を出して止まるためのものです。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Remove-OldFile.ps1 -Path D:\logs -OlderThan 90

    D:\logs の中で、90 日より古いファイルを一覧にします（ファイルは変更しません）。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Remove-OldFile.ps1 -Path D:\logs -OlderThan 90 -Include "*.log" -Delete

    90 日より古い .log ファイルを、確認してから削除します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Remove-OldFile.ps1 -Path D:\logs -OlderThan 365 -MoveTo E:\archive\logs

    365 日より古いファイルを、確認してから E:\archive\logs へ移します（フォルダの構成はそのまま）。

.NOTES
    Remove-OldFile 1.0.0
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

    [string]$OlderThan,

    [string[]]$Include = @(),

    [string[]]$Exclude = @(),

    [switch]$Delete,

    [string]$MoveTo,

    [switch]$RemoveEmptyFolder,

    [switch]$Force,

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

$ScriptVersion  = '1.0.0'
$ToolTitle      = '古いファイルの整理'
$FolderRows     = 20        # 画面に表示するフォルダの数（保存するときだけ。超えた分はファイルを見てもらう）
$FailRows       = 20        # 画面に表示する、削除できなかったファイルの数
$ProgressMs     = 250       # 進捗表示の間隔（ミリ秒）
$MaxDays        = 36500
$OutputPrefix   = 'OldFile_'
# このツールの結果のフォルダの名前（OldFile_<ホスト名>_<日時>）。同じ文字で始まるだけの利用者のフォルダは調べる
$OutputNameRule = [Text.RegularExpressions.Regex]::new('^OldFile_.+_[0-9]{8}_[0-9]{6}(?:_[0-9]+)?$', [Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant')
$SystemFolders  = '$RECYCLE.BIN', 'System Volume Information'
# オンライン専用ファイルの属性（OFFLINE・RECALL_ON_OPEN・RECALL_ON_DATA_ACCESS）。読むとダウンロードが始まる
$OnlineMask     = 0x00441000
$ReadOnlyFlag   = [int][IO.FileAttributes]::ReadOnly
$SystemFlag     = [int][IO.FileAttributes]::System
$EncryptedFlag  = [int][IO.FileAttributes]::Encrypted
# 別のドライブへ移すときに写す属性（隠し・アーカイブ・一時・インデックスを作らない）
$CopyAttrMask   = [int]([IO.FileAttributes]'Hidden, Archive, Temporary, NotContentIndexed')
$FileColumns    = 'ファイル', 'サイズ(バイト)', '作成日時', '最終更新日時', '経過日数', '結果', '備考', '移した先'
$FolderColumns  = 'フォルダ', '件数', 'サイズ(バイト)', 'いちばん古い日時', 'いちばん新しい日時'
$EmptyColumns   = 'フォルダ', '結果', '備考'
$SkipColumns    = '理由', 'パス', '詳細'
$TimeFormat     = 'yyyy/MM/dd HH:mm:ss'
$Invariant      = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom        = [Text.UTF8Encoding]::new($true)
$OnWindows      = [IO.Path]::DirectorySeparatorChar -eq '\'
$machineName    = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

# 結果の言葉
$StOut      = '対象外'
$StCannot   = '移せない'
$StFail     = '失敗'
$StDone     = '削除した'
$StEmptyNo  = '削除しなかった'

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

function Format-Size([long]$Bytes) {
    $units = 'B', 'KB', 'MB', 'GB', 'TB', 'PB'
    $value = [double]$Bytes
    $i = 0
    while ($value -ge 1024 -and $i -lt $units.Count - 1) { $value /= 1024; $i++ }
    if ($i -eq 0) { return '{0} B' -f $Bytes }
    return [string]::Format($Invariant, '{0:N1} {1}', $value, $units[$i])
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

# Windows のエラー番号（API を直接呼んだときは NativeErrorCode、.NET の例外は HRESULT の下位 16 ビット）
function Get-ErrorCode([Exception]$Exception) {
    if ($Exception -is [ComponentModel.Win32Exception]) { return $Exception.NativeErrorCode }
    return $Exception.HResult -band 0xFFFF
}

# CSV の 1 行
function ConvertTo-CsvLine([object[]]$Fields) {
    $parts = [string[]]::new($Fields.Count)
    for ($i = 0; $i -lt $Fields.Count; $i++) { $parts[$i] = ([string]$Fields[$i]).Replace('"', '""') }
    return '"' + ($parts -join '","') + '"'
}

# 5.1 と 7 のどちらでも BOM 付き UTF-8 で保存する（Export-Csv の -Encoding UTF8 は 7 だと BOM なしになる）
function Save-Csv([string]$FilePath, [string[]]$Columns, $Rows) {
    $writer = [IO.StreamWriter]::new($FilePath, $false, $Utf8Bom)
    try {
        $writer.WriteLine((ConvertTo-CsvLine $Columns))
        foreach ($row in $Rows) { $writer.WriteLine((ConvertTo-CsvLine $row)) }
    } finally {
        $writer.Dispose()
    }
}

# Excel は、= + - @ で始まる値を数式に、数字と区切りの記号だけの値（001、1.5）や「.1」「1e5」を
# 数値や日付に変えてしまうため、その形の値は ="..." で囲む。式に書けない 255 文字を超える値は、先頭に空白を足す
$excelRisk = [Text.RegularExpressions.Regex]::new('^(?:[=+\-@]|[\d\s.,/:+\-]+$)')
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

# 引数の値に引用符（"）が残っていないかを、ほかの確認より先にまとめて確かめ、そのあと、どのパラメーターにも当たらない引数
# （-_UnknownArgument に集まる）があれば止める。コマンドプロンプトで "C:\data\" のように最後を \" にすると、引用符が閉じずに
# 後ろの引数までその値に入り、空白を含む値の残りが余るか（-Path "C:\data\" -Since "2026-11-12 09:00" なら 09:00）、位置で受ける
# パラメーターに入る。どちらでも、引用符の書き方を先に案内する。カンマ区切りの値は、要素ごとに前後の引用符を外してから確かめる
function Assert-CommandArgument($Bound, [string[]]$QuoteNames) {
    foreach ($name in $QuoteNames) {
        if (-not $Bound.ContainsKey($name)) { continue }
        foreach ($value in @($Bound[$name])) {
            foreach ($part in ([string]$value -split ',')) {
                if ((Remove-Quote $part).Trim("'").Contains('"')) { Stop-WithQuoteError ('-' + $name) (Remove-Quote ([string]$value)) }
            }
        }
    }
    if (-not $Bound.ContainsKey('_UnknownArgument')) { return }
    $shown = @(foreach ($arg in @($Bound['_UnknownArgument'])) { if ([string]$arg -match '^$|\s') { '"{0}"' -f $arg } else { [string]$arg } }) -join ' '
    Stop-WithError (('どのパラメーターにも当たらない引数があります: {0}' -f $shown) + "`r`n" +
        'パラメーターの名前の書き間違いや、空白を含む値の引用符（"）の付け忘れがないかを確かめてください。コマンドプロンプトで値の最後を \" にしたとき（"C:\data\" など）も、後ろの引数がずれてこうなります。')
}
Assert-CommandArgument $PSBoundParameters @('Path', 'OlderThan', 'Include', 'Exclude', 'MoveTo', 'OutputFolder')

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

# %TEMP% のような環境変数を展開する。タスクスケジューラなどでは、
# コマンドプロンプトを通らずに起動されて、展開されないまま届くことがあるため
function Expand-PathText([string]$Text) {
    if (-not $Text.Contains('%')) { return $Text }
    return [Environment]::ExpandEnvironmentVariables($Text)
}

# 共有フォルダやネットワークドライブを開けないときの案内。管理者として実行しているときは、そのことも添える
function Get-NetworkHint([string]$FullPath) {
    $isNetwork = $FullPath.StartsWith('\\')
    if (-not $isNetwork -and $FullPath -match '^[A-Za-z]:') {
        try { $isNetwork = ([IO.DriveInfo]::new($FullPath.Substring(0, 2))).DriveType -eq 'Network' } catch { }
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
            $hint += '※ 管理者として実行しています。通常のユーザーで接続した共有フォルダやネットワークドライブは、管理者として開いた PowerShell からは見えないことがあります。\\サーバー名\共有名 の形で指定するか、管理者としてではなく実行してください。'
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
    $texts = [System.Collections.Generic.List[string]]::new()
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

# 文字の番号の順に並べる（$Keys と同じ順に $Values を並べ替える）。
# [Array] に変えて渡さないと、PowerShell が型を合わせた写しを作って並べ替え、$Values が並ばないことがある
function Get-SortedList([string[]]$Keys, [object[]]$Values) {
    if ($Values.Count -le 1) { return , $Values }
    [Array]::Sort([Array]$Keys, [Array]$Values, [Collections.IComparer][StringComparer]::Ordinal)
    return , $Values
}

$dirType  = [IO.DirectoryInfo]
$rpFlag   = [IO.FileAttributes]::ReparsePoint
$hsFlag   = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System

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

# 移す先のドライブの空き容量。共有フォルダは DriveInfo が対応していないため、FileSystemObject で取得する。
# 取得できないときは -1 を返す
function Get-FreeSpace([string]$FullPath) {
    try {
        if ($FullPath.StartsWith('\\')) {
            $fso = New-Object -ComObject Scripting.FileSystemObject
            return [long]$fso.GetDrive($fso.GetDriveName($FullPath)).FreeSpace
        }
        $drive = [IO.DriveInfo]::new([IO.Path]::GetPathRoot($FullPath))
        if ($drive.IsReady) { return [long]$drive.AvailableFreeSpace }
    } catch { }
    return -1L
}

#endregion

#region 引数の確認

$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

# -Path は、-File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# カンマを含む名前のフォルダのため、全体が 1 つのフォルダとして存在するときは区切らない
$rootTexts = [System.Collections.Generic.List[string]]::new()
foreach ($item in $Path) {
    $whole = Remove-Quote $item
    if ($whole.Contains('"')) { Stop-WithQuoteError '-Path' $whole }
    $whole = Expand-PathText $whole
    $isWhole = $false
    if ($whole.Contains(',')) {
        try { $isWhole = [IO.Directory]::Exists((ConvertTo-LongPath (Resolve-FullPath (Resolve-StartPath $whole)))) } catch { }
    }
    if ($isWhole) { $rootTexts.Add($whole); continue }
    foreach ($part in ($whole -split ',')) {
        $text = (Remove-Quote $part).Trim("'")
        if ($text) { $rootTexts.Add($text) }
    }
}
if ($rootTexts.Count -eq 0) { Stop-WithError '調べるフォルダを -Path で指定してください' }

# -OlderThan は、範囲外や数値でない値を日本語で案内するため、文字で受け取って確かめる
$olderText = ''
if ($null -ne $OlderThan) { $olderText = (Remove-Quote $OlderThan).Trim("'") }
if (-not $olderText) { Stop-WithError '何日より古いファイルを対象にするかを、-OlderThan で指定してください（例: -OlderThan 90）' }
$days = 0
$normalizedDays = $olderText
try { $normalizedDays = $olderText.Normalize([Text.NormalizationForm]::FormKC) } catch { }
if ($normalizedDays -notmatch '^[0-9]{1,6}$' -or -not [int]::TryParse($normalizedDays, [Globalization.NumberStyles]::None, $Invariant, [ref]$days) -or $days -lt 1 -or $days -gt $MaxDays) {
    Stop-WithError (Format-Invariant '-OlderThan は 1～{0} の数値で指定してください: {1}' $MaxDays $olderText)
}

# 何をするか（一覧だけ・削除・移動）
$moveText = $null
if ($PSBoundParameters.ContainsKey('MoveTo')) {
    $moveText = Remove-Quote $MoveTo
    if ($moveText.Contains('"')) { Stop-WithQuoteError '-MoveTo' $moveText }
    if (-not $moveText) { Stop-WithError '移す先のフォルダを -MoveTo で指定してください' }
}
if ($Delete -and $null -ne $moveText) { Stop-WithError '-Delete と -MoveTo は一緒に使えません' }
$mode = if ($Delete) { 'Delete' } elseif ($null -ne $moveText) { 'Move' } else { 'List' }
$acting = $mode -ne 'List'
if ($Force -and -not $acting) { Stop-WithError '-Force は -Delete か -MoveTo と一緒に指定してください' }
if ($NoSave -and $acting) { Stop-WithError '削除するとき・移すときは -NoSave を使えません（記録を残すため）' }
if ($mode -eq 'Move' -and $rootTexts.Count -gt 1) { Stop-WithError '-MoveTo を使うときは、-Path を 1 つだけ指定してください' }

# 削除するとき・移すときの言葉
$Words = if ($mode -eq 'Move') {
    @{ Plan = '移す予定'; Done = '移した'; No = '移さなかった'; Ask = '移しますか？ [Y] はい  [N] いいえ'; DidAll = '移しました'; DidNot = '移しませんでした'
       FailTitle = '■ 移せなかったファイル'; Files = '移すファイル'; Rest = '移さないファイル'; Interrupted = '中断したため、移していません'
       Changed = '予定を立てた後に変更されたため、移していません'; Gone = '予定を立てた後になくなっていました'; Stopped = '移した'
       CannotAsk = '確認できないため、移さずに止めました。確認せずに移すときは -Force を付けてください'
       AdminAfter = '※ 管理者として実行すると、移せるものがあります' }
} else {
    @{ Plan = '削除する予定'; Done = '削除した'; No = '削除しなかった'; Ask = '削除しますか？ [Y] はい  [N] いいえ'; DidAll = '削除しました'; DidNot = '削除しませんでした'
       FailTitle = '■ 削除できなかったファイル'; Files = '削除するファイル'; Rest = '削除しないファイル'; Interrupted = '中断したため、削除していません'
       Changed = '予定を立てた後に変更されたため、削除していません'; Gone = '予定を立てた後になくなっていました'; Stopped = '削除した'
       CannotAsk = '確認できないため、削除せずに止めました。確認せずに削除するときは -Force を付けてください'
       AdminAfter = '※ 管理者として実行すると、削除できるものがあります' }
}

# 指定できない場所。誤って指定すると、OS やアプリが動かなくなるところ
# ドライブ（C: や管理共有の \\server\C$）ごとに、ルートから見たこれらのフォルダと、Users の直下の各ユーザーのフォルダ
$ProtectedNames = 'Windows', 'Windows\System32', 'Windows\SysWOW64', 'Windows\WinSxS', 'Windows\Sysnative', 'Program Files', 'Program Files (x86)', 'ProgramData', 'Users'
$driveRootRule  = [Text.RegularExpressions.Regex]::new('^(?:[A-Za-z]:|\\\\[^\\]+\\[A-Za-z]\$)(?=\\|$)', [Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant')
# 管理共有 \\server\ADMIN$ は、その PC の Windows のフォルダそのもの
$adminShareRule = [Text.RegularExpressions.Regex]::new('^\\\\[^\\]+\\ADMIN\$(?=\\|$)', [Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant')
$WindowsInner   = 'System32', 'SysWOW64', 'WinSxS', 'Sysnative'
# この PC の Windows のフォルダ、Program Files、各ユーザーのプロファイルのフォルダ（ほかのドライブに置いている場合に備える）
$protectedFixed = [System.Collections.Generic.List[string]]::new()
function Add-ProtectedPath([string]$Text) {
    if (-not $Text) { return }
    try {
        $full = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($Text)).TrimEnd('\')
        if ($full) { $protectedFixed.Add($full) }
    } catch { }
}
foreach ($name in 'SystemRoot', 'ProgramFiles', 'ProgramFiles(x86)', 'ProgramW6432', 'ProgramData', 'PUBLIC') {
    Add-ProtectedPath ([Environment]::GetEnvironmentVariable($name))
}
# Sysnative は、32 ビットの PowerShell から 64 ビットの System32 を指す名前
if ($env:SystemRoot) { foreach ($name in $WindowsInner) { Add-ProtectedPath (Join-Path $env:SystemRoot $name) } }
try {
    $profileKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList')
    if ($profileKey) {
        try {
            foreach ($name in 'ProfilesDirectory', 'Public', 'Default') { Add-ProtectedPath ([string]$profileKey.GetValue($name)) }
            foreach ($sid in $profileKey.GetSubKeyNames()) {
                $sub = $profileKey.OpenSubKey($sid)
                if ($sub) { try { Add-ProtectedPath ([string]$sub.GetValue('ProfileImagePath')) } finally { $sub.Dispose() } }
            }
        } finally { $profileKey.Dispose() }
    }
} catch { }

# 指定できない場所なら、その理由の文を返す
function Get-ProtectedText([string]$Full) {
    $norm = $Full.TrimEnd('\')
    $match = $driveRootRule.Match($norm)
    $checks = [System.Collections.Generic.List[string]]::new($protectedFixed)
    $adminMatch = $adminShareRule.Match($norm)
    if ($adminMatch.Success) {
        $checks.Add($adminMatch.Value)
        foreach ($name in $WindowsInner) { $checks.Add($adminMatch.Value + '\' + $name) }
    }
    if ($match.Success) {
        if ($match.Length -eq $norm.Length) { return 'ドライブのルートは指定できません: {0}' -f $Full }
        foreach ($name in $ProtectedNames) { $checks.Add($match.Value + '\' + $name) }
        # Users の直下の各ユーザーのフォルダ（C:\Users\yamada、C:\Users\Public など）
        $parent = $norm.Substring(0, $norm.LastIndexOf('\'))
        if ([string]::Equals($parent, $match.Value + '\Users', [StringComparison]::OrdinalIgnoreCase)) {
            return 'この場所は指定できません（ユーザーのフォルダには、設定やデータのファイルがあるため）: {0}' -f $Full
        }
    }
    # そのものが指定できない場所なら、中にあるものより先に、その理由を返す
    foreach ($check in $checks) {
        if ([string]::Equals($norm, $check, [StringComparison]::OrdinalIgnoreCase)) {
            return 'この場所は指定できません（OS やアプリが動くのに必要なファイルがあるため）: {0}' -f $Full
        }
    }
    foreach ($check in $checks) {
        if (($check + '\').StartsWith($norm + '\', [StringComparison]::OrdinalIgnoreCase)) {
            return 'この場所は指定できません（中に {0} があるため）: {1}' -f $check, $Full
        }
    }
    return $null
}

# Windows の API を呼ぶ部品（Add-Type で C# をコンパイルする）。
# - 削除・移動するファイル（空になったフォルダ）を 1 件ずつ、ほかのプロセスに削除・移動させない形で開き、開いたものの
#   実際の場所が調べたときと同じかを確かめてから、開いたまま削除（移動）する。パスで削除すると、確かめた後で途中のフォルダを
#   ジャンクションに差し替えられたとき（その場での変換を含む）、-Path の外のファイルを消してしまうため
# - -Path の実際の場所（SUBST・ネットワークドライブ・リンクの先）を確かめる
# - 古いファイルの記録を、軽い入れ物（C# のクラス）で持つ。[pscustomobject] は 1 件 3 KB ほど使うため
$NativeSource = @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

public class WbRemoveOldFileItem {
    public string Display;
    public long Size;
    public DateTime CreatedUtc;
    public DateTime WrittenUtc;
    public int Attrs;
    public object Node;
    public bool IsTarget;
    public string OutKind;
    public string Result;
    public string Note;
    public string Reason = "";
    public string MovedTo = "";
    public string DestDisplay = "";
    public string DestLong = "";
    // PowerShell は、文字の引数に渡した $null を空の文字にするため、空の outKind も「対象」とみなす
    public WbRemoveOldFileItem(string display, long size, DateTime createdUtc, DateTime writtenUtc, int attrs, object node, string outKind, string result, string note) {
        Display = display; Size = size; CreatedUtc = createdUtc; WrittenUtc = writtenUtc; Attrs = attrs; Node = node;
        IsTarget = string.IsNullOrEmpty(outKind); OutKind = IsTarget ? null : outKind; Result = result; Note = note;
    }
}

// 開いたファイルやフォルダの状態（開いたハンドルから読んだもの）
public class WbRemoveOldFileState {
    public uint Attributes;
    public uint ReparseTag;
    public bool HasTag;
    public long Size;
    public DateTime CreatedUtc;
    public DateTime WrittenUtc;
    public string FinalPath;
    // リンク（ジャンクション・シンボリックリンクなど、別の場所を指すリパースポイント）か。
    // OneDrive などのリパースポイントは、別の場所を指さないのでリンクではない。種類（タグ）を返さない共有フォルダの機器では、リパースポイントならリンクとみなす
    public bool IsLink { get { return (Attributes & 0x400) != 0 && (!HasTag || (ReparseTag & 0x20000000) != 0); } }
}

public static class WbRemoveOldFileNative {
    const uint FileReadAttributes = 0x80;
    const uint Delete = 0x10000;
    const uint Synchronize = 0x100000;
    const uint GenericRead = 0x80000000;
    const uint GenericWrite = 0x40000000;
    const uint ShareReadWrite = 0x3;
    const uint ShareAll = 0x7;
    const uint CreateNew = 1;
    const uint OpenExisting = 3;
    const uint BackupSemantics = 0x02000000;
    const uint OpenReparsePoint = 0x00200000;
    const int FileRenameInfo = 3;
    const int FileDispositionInfo = 4;
    const int FileAttributeTagInfo = 9;

    [StructLayout(LayoutKind.Sequential)]
    struct AttributeTagInfo { public uint FileAttributes; public uint ReparseTag; }

    [StructLayout(LayoutKind.Sequential)]
    struct ByHandleFileInformation {
        public uint FileAttributes;
        public uint CreationTimeLow, CreationTimeHigh, LastAccessTimeLow, LastAccessTimeHigh, LastWriteTimeLow, LastWriteTimeHigh;
        public uint VolumeSerialNumber, FileSizeHigh, FileSizeLow, NumberOfLinks, FileIndexHigh, FileIndexLow;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern SafeFileHandle CreateFileW(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool GetFileInformationByHandleEx(SafeFileHandle handle, int infoClass, out AttributeTagInfo info, uint size);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool GetFileInformationByHandle(SafeFileHandle handle, out ByHandleFileInformation info);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern uint GetFinalPathNameByHandleW(SafeFileHandle handle, StringBuilder buffer, uint length, uint flags);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool SetFileInformationByHandle(SafeFileHandle handle, int infoClass, IntPtr info, uint size);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool ReadFile(SafeFileHandle handle, byte[] buffer, int count, out int read, IntPtr overlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool WriteFile(SafeFileHandle handle, byte[] buffer, int count, out int written, IntPtr overlapped);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool DeleteFileW(string name);

    // 開いたハンドルから中身を読んで、新しいファイルに書く（同じ名前があれば ERROR_FILE_EXISTS（80））。
    // パスで読むと、確かめた後で途中のフォルダを差し替えられたとき、別のファイルの中身を写してしまうため。
    // 途中で失敗したら、書きかけのファイルを消してから例外を出す
    public static long CopyFromHandle(SafeFileHandle source, string destPath) {
        SafeFileHandle dest = CreateFileW(destPath, GenericWrite, 0, IntPtr.Zero, CreateNew, 0x80, IntPtr.Zero);
        if (dest.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
        long total = 0;
        try {
            using (dest) {
                byte[] buffer = new byte[1024 * 1024];
                while (true) {
                    int read;
                    if (!ReadFile(source, buffer, buffer.Length, out read, IntPtr.Zero)) throw new Win32Exception(Marshal.GetLastWin32Error());
                    if (read == 0) break;
                    int written;
                    if (!WriteFile(dest, buffer, read, out written, IntPtr.Zero)) throw new Win32Exception(Marshal.GetLastWin32Error());
                    if (written != read) throw new Win32Exception(29);
                    total += read;
                }
            }
        } catch {
            DeleteFileW(destPath);
            throw;
        }
        return total;
    }

    static string GetFinalPathOf(SafeFileHandle handle) {
        StringBuilder buffer = new StringBuilder(512);
        uint length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 0);
        if (length >= buffer.Capacity) {
            buffer = new StringBuilder((int)length + 1);
            length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 0);
        }
        if (length == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
        return buffer.ToString();
    }

    // 実際の場所（\\?\C:\... か \\?\UNC\server\share\...）。SUBST・ネットワークドライブ・リンクの先をたどった後のパス
    public static string GetFinalPath(string path) {
        using (SafeFileHandle handle = CreateFileW(path, FileReadAttributes, ShareAll, IntPtr.Zero, OpenExisting, BackupSemantics, IntPtr.Zero)) {
            if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
            return GetFinalPathOf(handle);
        }
    }

    // 削除・移動のために、ファイル（かフォルダ）を開く。リンクはたどらない。
    // 削除の権限で、削除の共有を許さずに開くため、閉じるまでは、ほかのプロセスはこれを削除・移動できない。
    // 開いた後は、パス（親のフォルダの名前など）が変わっても、このハンドルは同じファイルを指したまま。
    // 中身を写すとき（別のドライブへ移すとき）は、読む権限も付ける
    public static SafeFileHandle OpenForChange(string path, bool isDirectory, bool forRead, out WbRemoveOldFileState state) {
        uint flags = OpenReparsePoint | (isDirectory ? BackupSemantics : 0);
        uint access = Delete | FileReadAttributes | Synchronize | (forRead ? GenericRead : 0);
        SafeFileHandle handle = CreateFileW(path, access, ShareReadWrite, IntPtr.Zero, OpenExisting, flags, IntPtr.Zero);
        if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
        try {
            state = Query(handle);
        } catch {
            handle.Dispose();
            throw;
        }
        return handle;
    }

    // 開いたハンドルから、属性・サイズ・日時・実際の場所を読む
    public static WbRemoveOldFileState Query(SafeFileHandle handle) {
        ByHandleFileInformation info;
        if (!GetFileInformationByHandle(handle, out info)) throw new Win32Exception(Marshal.GetLastWin32Error());
        WbRemoveOldFileState state = new WbRemoveOldFileState();
        state.Attributes = info.FileAttributes;
        state.Size = ((long)info.FileSizeHigh << 32) | info.FileSizeLow;
        state.CreatedUtc = DateTime.FromFileTimeUtc(((long)info.CreationTimeHigh << 32) | info.CreationTimeLow);
        state.WrittenUtc = DateTime.FromFileTimeUtc(((long)info.LastWriteTimeHigh << 32) | info.LastWriteTimeLow);
        AttributeTagInfo tag;
        if (GetFileInformationByHandleEx(handle, FileAttributeTagInfo, out tag, (uint)Marshal.SizeOf(typeof(AttributeTagInfo)))) {
            state.ReparseTag = tag.ReparseTag;
            state.HasTag = true;
        }
        state.FinalPath = GetFinalPathOf(handle);
        return state;
    }

    // 開いたハンドルに削除の印を付ける（閉じたときに削除される）。中にものがあるフォルダは、ERROR_DIR_NOT_EMPTY（145）になる
    public static void MarkForDelete(SafeFileHandle handle) {
        IntPtr buffer = Marshal.AllocHGlobal(4);
        try {
            Marshal.WriteInt32(buffer, 1);
            if (!SetFileInformationByHandle(handle, FileDispositionInfo, buffer, 1)) throw new Win32Exception(Marshal.GetLastWin32Error());
        } finally {
            Marshal.FreeHGlobal(buffer);
        }
    }

    // 開いたハンドルのファイルを、同じドライブの中の別のパスへ移す（上書きはしない）。
    // 移す先は、\\?\ 付きのパス。Windows の内部の書き方（\??\）に直して渡す（MoveFileEx と同じ）。
    // 別のドライブだと ERROR_NOT_SAME_DEVICE（17）、同じ名前があると ERROR_ALREADY_EXISTS（183）になる
    public static void RenameTo(SafeFileHandle handle, string targetLongPath) {
        string target = targetLongPath.StartsWith(@"\\?\") ? @"\??\" + targetLongPath.Substring(4) : targetLongPath;
        int nameOffset = IntPtr.Size == 8 ? 20 : 12;
        int nameBytes = target.Length * 2;
        int size = nameOffset + nameBytes + 2;
        IntPtr buffer = Marshal.AllocHGlobal(size);
        try {
            for (int i = 0; i < size; i++) Marshal.WriteByte(buffer, i, 0);
            Marshal.WriteIntPtr(buffer, IntPtr.Size, IntPtr.Zero);
            Marshal.WriteInt32(buffer, IntPtr.Size * 2, nameBytes);
            char[] chars = target.ToCharArray();
            Marshal.Copy(chars, 0, IntPtr.Add(buffer, nameOffset), chars.Length);
            if (!SetFileInformationByHandle(handle, FileRenameInfo, buffer, (uint)size)) throw new Win32Exception(Marshal.GetLastWin32Error());
        } finally {
            Marshal.FreeHGlobal(buffer);
        }
    }
}
'@
$nativeReady = $false
$nativeError = $null
try {
    # 5.1 の Add-Type は、コンパイラーの警告もエラーとして扱う
    if (-not ('WbRemoveOldFileNative' -as [type])) { Add-Type -TypeDefinition $NativeSource -IgnoreWarnings -WarningAction SilentlyContinue -ErrorAction Stop }
    $nativeReady = $true
} catch {
    $nativeError = Get-ErrorText (Get-InnerException $_.Exception)
}
# 削除・移動は、フォルダを開いたままにできないと安全に行えないため止める（一覧だけなら続ける）
if ($acting -and -not $nativeReady) {
    Stop-WithError ('Windows の API を呼ぶ準備ができません（Add-Type を使えません）。この環境では、削除・移動はできません。一覧だけなら、-Delete と -MoveTo を付けずに実行できます: {0}' -f $nativeError)
}

$roots = [System.Collections.Generic.List[object]]::new()
foreach ($text in $rootTexts) {
    $rootText = $text
    if ($rootText -match '^(\\\\\?\\)?[A-Za-z]:$') { $rootText += '\' }
    $rootFull = Resolve-FullPath (Resolve-StartPath $rootText)
    $rootLong = ConvertTo-LongPath $rootFull
    $isFile = $false
    $isDir  = $false
    try { $isDir = [IO.Directory]::Exists($rootLong) } catch { }
    if (-not $isDir) { try { $isFile = [IO.File]::Exists($rootLong) } catch { } }
    if ($isFile) { Stop-WithError ('フォルダを指定してください（ファイルは指定できません）: {0}' -f $rootFull) }
    if (-not $isDir) {
        Stop-WithError ((@(('フォルダが見つかりません: {0}' -f $rootFull)) + (Get-NetworkHint $rootFull)) -join "`r`n")
    }
    $protectedText = Get-ProtectedText $rootFull
    if ($protectedText) {
        Stop-WithError ($protectedText + "`r`n" + 'この中の古いファイルを整理するときは、もっと下のフォルダを指定してください（例: C:\Windows\Temp、C:\Users\yamada\Downloads）')
    }
    # -Path やその途中にリンクがあると、リンク先の実際の場所が分からず、指定できない場所かどうかを確かめられない。
    # ドライブのルート（C:）と共有のルート（\\server\share）の手前まで、1 つずつ親をたどって確かめる
    $pathRoot = if ($rootFull -match '^(\\\\[^\\]+\\[^\\]+)') { $Matches[1] } else { $rootFull.Substring(0, [math]::Min(2, $rootFull.Length)) }
    $checkPath = $rootFull.TrimEnd('\')
    while ($checkPath.Length -gt $pathRoot.Length) {
        $linkName = $null
        try {
            $checkInfo = [IO.DirectoryInfo]::new((ConvertTo-LongPath $checkPath))
            if ([int]$checkInfo.Attributes -ne -1) { $linkName = Get-LinkName $checkInfo }
        } catch { }
        if ($linkName) {
            Stop-WithError ('-Path の途中にリンク（{0}）があります。リンク先の実際のフォルダを指定してください: {1}' -f $linkName, $checkPath)
        }
        $cut = $checkPath.LastIndexOf('\')
        if ($cut -le 0) { break }
        $checkPath = $checkPath.Substring(0, $cut)
    }
    # フォルダを中まで読めるかを確かめる（共有フォルダは、接続していないと読めない）
    try {
        [void][IO.Directory]::EnumerateFileSystemEntries($rootLong).GetEnumerator().MoveNext()
    } catch {
        $reason = Get-ErrorText (Get-InnerException $_.Exception)
        Stop-WithError ((@(('フォルダを開けません: {0}（{1}）' -f $rootFull, $reason)) + (Get-NetworkHint $rootFull)) -join "`r`n")
    }
    # SUBST やネットワークドライブを通すと、名前だけでは指定できない場所か分からないため、実際の場所でも確かめる。
    # 実際の場所は、削除するときに、ファイルが -Path の中にあるかを確かめるのにも使う
    $finalLong = $null
    if ($nativeReady) {
        $finalError = $null
        try { $finalLong = ([WbRemoveOldFileNative]::GetFinalPath($rootLong)).TrimEnd('\') } catch { $finalError = Get-ErrorText (Get-InnerException $_.Exception) }
        if ($finalLong) {
            $finalFull = ConvertTo-DisplayPath $finalLong
            if (-not [string]::Equals($finalFull, $rootFull.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
                $protectedText = Get-ProtectedText $finalFull
                if ($protectedText) {
                    Stop-WithError ($protectedText + "`r`n" + ('指定したパス {0} の実際の場所です（SUBST・ネットワークドライブ・リンクの先も確かめています）。もっと下のフォルダを指定してください' -f $rootFull))
                }
            }
        } elseif ($acting) {
            Stop-WithError ('フォルダの実際の場所を確かめられないため、{0}ずに止めました: {1}（{2}）' -f $(if ($mode -eq 'Move') { '移さ' } else { '削除せ' }), $rootFull, $finalError)
        }
    }
    $roots.Add([pscustomobject]@{ Full = $rootFull; Long = $rootLong; FinalLong = $finalLong })
}

# 同じフォルダや、片方の中にあるフォルダを指定すると、同じファイルを二重に数えてしまう
for ($i = 0; $i -lt $roots.Count; $i++) {
    for ($j = $i + 1; $j -lt $roots.Count; $j++) {
        $firstPrefix  = $roots[$i].Full.TrimEnd('\') + '\'
        $secondPrefix = $roots[$j].Full.TrimEnd('\') + '\'
        if ($firstPrefix.StartsWith($secondPrefix, [StringComparison]::OrdinalIgnoreCase) -or $secondPrefix.StartsWith($firstPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            Stop-WithError ('指定したフォルダの片方が、もう片方の中にあります: {0} と {1}' -f $roots[$i].Full, $roots[$j].Full)
        }
    }
}

# 移す先。打ち間違いで別の場所にフォルダを作らないよう、先に作ってあることを求める
$moveFull = $null
$moveLong = $null
$sameVolume = $false
if ($mode -eq 'Move') {
    $moveText = Expand-PathText $moveText
    if ($moveText -match '^(\\\\\?\\)?[A-Za-z]:$') { $moveText += '\' }
    $moveFull = Resolve-FullPath (Resolve-StartPath $moveText)
    $moveLong = ConvertTo-LongPath $moveFull
    $moveExists = $false
    try { $moveExists = [IO.Directory]::Exists($moveLong) } catch { }
    if (-not $moveExists) { Stop-WithError ('移す先のフォルダが見つかりません: {0}。先に作ってから実行してください' -f $moveFull) }
    $rootPrefix = $roots[0].Full.TrimEnd('\') + '\'
    $movePrefix = $moveFull.TrimEnd('\') + '\'
    if ($rootPrefix.StartsWith($movePrefix, [StringComparison]::OrdinalIgnoreCase) -or $movePrefix.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        Stop-WithError ('移す先が、調べるフォルダと重なっています: {0} と {1}' -f $roots[0].Full, $moveFull)
    }
    try { $sameVolume = [string]::Equals([IO.Path]::GetPathRoot($roots[0].Full), [IO.Path]::GetPathRoot($moveFull), [StringComparison]::OrdinalIgnoreCase) } catch { }
}

# -Include は名前に当てはめる
$includeRules = @(foreach ($text in (Get-ListText $Include '-Include')) { [pscustomobject]@{ Text = $text; Matcher = (New-NameMatcher $text) } })
$hasInclude = $includeRules.Count -gt 0
function Test-IncludeMatch([string]$Name) {
    foreach ($rule in $includeRules) { if ($rule.Matcher.IsMatch($Name)) { return $true } }
    return $false
}

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

# 実行しているスクリプトのファイルは、古くても対象にしない（調べるときは \\?\ 付きのパスで比べる）
$selfLong = $null
if ($PSCommandPath) {
    $selfFull = $PSCommandPath
    try { $selfFull = [IO.Path]::GetFullPath($PSCommandPath) } catch { }
    $selfLong = ConvertTo-LongPath $selfFull
}

$startTime = Get-Date
$startUtc  = $startTime.ToUniversalTime()
# 基準の日時は、実行を始めた日時から N×24 時間前
$cutoffUtc = $startUtc.AddDays(-$days)
$cutoff    = $cutoffUtc.ToLocalTime()

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
        # 削除するとき・移すときは、何を削除したかの記録を残せないため止める
        if ($acting) { Stop-WithError ('保存先に書き込めないため、{0}ずに止めました: {1}（{2}）' -f $(if ($mode -eq 'Move') { '移さ' } else { '削除せ' }), $baseDir, $reason) }
        $saveWarning = '保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, $reason
        Write-Warning $saveWarning
    }
}

#endregion

#region 判定の部品

$skipped   = [System.Collections.Generic.List[object]]::new()
$skipKinds = [ordered]@{ '権限なし' = 0; 'その他のエラー' = 0 }
$excludedCount = 0
$linkCount     = 0
$fileCount     = 0         # 調べたファイル（-Include に当てはまり、除外されなかったもの）
$oldItems      = [System.Collections.Generic.List[object]]::new()   # 古いファイル
$dirNodes      = [System.Collections.Generic.List[object]]::new()   # 調べたフォルダ（たどった順）
$folderStats   = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
$sw            = [Diagnostics.Stopwatch]::StartNew()
$nextProgress  = 0L

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

# 削除・移動に失敗した理由（画面の短い言葉と、files.csv の備考）
function Get-FailReason([Exception]$Exception) {
    $code = Get-ErrorCode $Exception
    if ($Exception -is [UnauthorizedAccessException] -or $code -eq 5) { return @('アクセス拒否', 'アクセスが拒否されました') }
    if ($code -eq 32 -or $code -eq 33) { return @('使用中', 'ほかのプロセスが使用中です') }
    if ($code -eq 112 -or $code -eq 39) { return @('容量不足', '移す先の空き容量が足りません') }
    if ($code -eq 80 -or $code -eq 183) { return @('同じ名前がある', '移す先に同じ名前のファイルがあります') }
    return @('その他', (Get-ErrorText $Exception))
}

# 5.1 は、Ctrl+C（パイプラインの停止）を取りこぼして動き続けることがある。
# 削除するとき・移すときは、コンソールで動かしている間、Ctrl+C をキーの入力として受け取り、ファイルの区切りごとに確かめて止める。
# 止めたあとの片付けと結果の保存は、finally で行う
$ctrlCSaved = $null   # 変える前の TreatControlCAsInput（変えていないときは $null）
function Enable-CtrlCInput {
    if (-not $acting -or $null -ne $script:ctrlCSaved) { return }
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

# 種類ごとの件数を「読み取り専用 2・システムファイル 1」の形にする
function Get-KindText($Kinds) {
    return (@(foreach ($name in $Kinds.Keys) { if ($Kinds[$name] -gt 0) { Format-Invariant '{0} {1:N0}' $name $Kinds[$name] } }) -join '・')
}

function Format-LocalTime([DateTime]$Utc, [string]$Format) {
    return $Utc.ToLocalTime().ToString($Format, $Invariant)
}

# 経過日数（作成日時と最終更新日時の新しいほうから、実行した日時まで。1 日に満たない分は切り捨てる）
function Get-AgeText($Item) {
    $newerUtc = if ($Item.WrittenUtc -gt $Item.CreatedUtc) { $Item.WrittenUtc } else { $Item.CreatedUtc }
    return ([long][math]::Floor(($startUtc - $newerUtc).TotalDays)).ToString('N0', $Invariant)
}

# 調べたときのパスの、実際の場所（-Path の実際の場所 + -Path から見た相対パス）
function Get-ExpectedFinal($Node, [string]$Display) {
    $root = $Node.Root
    return $root.FinalLong + $Display.Substring($root.Full.TrimEnd('\').Length)
}

# 予定を立てた後に変わったか（更新された、サイズが変わった、読み取り専用などになった、リンクに差し替えられた）
function Test-StateChanged($Entry, $State) {
    $attrs = [int]$State.Attributes
    return $State.Size -ne $Entry.Size -or $State.CreatedUtc -ne $Entry.CreatedUtc -or $State.WrittenUtc -ne $Entry.WrittenUtc -or
        (($attrs -band ($ReadOnlyFlag -bor $SystemFlag -bor $OnlineMask)) -ne 0) -or
        (($attrs -band [int]$rpFlag) -ne ($Entry.Attrs -band [int]$rpFlag))
}

function Set-EntryFail($Entry, [string]$Reason, [string]$Note) {
    $Entry.Result = $StFail
    $Entry.Reason = $Reason
    $Entry.Note = $Note
}

# 1 件を削除するか、移す。結果を $Entry に書き、'Done'・'Fail'・'Gone'（なくなっていた）のどれかを返す。
# ファイルを、ほかのプロセスに削除・移動させない形で開き、開いたファイルの実際の場所が調べたときと同じか、
# 変わっていないかを確かめてから、開いたまま削除（移動）する。パスで削除すると、確かめた後で途中のフォルダを
# ジャンクションに差し替えられたとき、-Path の外のファイルを消してしまうため
function Invoke-FileAction($Entry, [string]$Long) {
    $state = $null
    try {
        $handle = [WbRemoveOldFileNative]::OpenForChange($Long, $false, ($mode -eq 'Move'), [ref]$state)
    } catch {
        $ex = Get-InnerException $_.Exception
        $code = Get-ErrorCode $ex
        if ($code -eq 2 -or $code -eq 3) {
            $Entry.Result = $Words.No
            $Entry.Note = $Words.Gone
            return 'Gone'
        }
        $reason = Get-FailReason $ex
        Set-EntryFail $Entry $reason[0] $reason[1]
        return 'Fail'
    }
    try {
        $expected = Get-ExpectedFinal $Entry.Node $Entry.Display
        if (-not [string]::Equals($state.FinalPath, $expected, [StringComparison]::OrdinalIgnoreCase)) {
            Set-EntryFail $Entry '予定の後に変更' ('ファイルの実際の場所が、調べたときと違います（途中のフォルダがリンクに変わったなど）: {0}' -f (ConvertTo-DisplayPath $state.FinalPath))
            return 'Fail'
        }
        if (Test-StateChanged $Entry $state) {
            Set-EntryFail $Entry '予定の後に変更' $Words.Changed
            return 'Fail'
        }
        if ($mode -eq 'Delete') {
            [WbRemoveOldFileNative]::MarkForDelete($handle)
            $Entry.Result = $Words.Done
            return 'Done'
        }

        # 移す。足りないフォルダは作る
        $destDir = $Entry.DestLong.Substring(0, $Entry.DestLong.LastIndexOf('\'))
        if (-not $createdDirs.Contains($destDir)) {
            try {
                [void][IO.Directory]::CreateDirectory($destDir)
            } catch {
                Set-EntryFail $Entry (Get-FailReason (Get-InnerException $_.Exception))[0] ('移す先のフォルダを作れません: ' + (Get-ErrorText (Get-InnerException $_.Exception)))
                return 'Fail'
            }
            [void]$createdDirs.Add($destDir)
        }
        # 同じドライブの中は、開いたファイルの名前を変えるだけ（日時・属性・アクセス権はそのまま）
        $renamed = $false
        if ($sameVolume) {
            try {
                [WbRemoveOldFileNative]::RenameTo($handle, $Entry.DestLong)
                $renamed = $true
            } catch {
                $ex = Get-InnerException $_.Exception
                # 別のボリューム（マウントしたフォルダなど）だったときは、コピーして移す
                if ((Get-ErrorCode $ex) -ne 17) {
                    $reason = Get-FailReason $ex
                    Set-EntryFail $Entry $reason[0] $reason[1]
                    return 'Fail'
                }
            }
        }
        if (-not $renamed -and ($Entry.Attrs -band $EncryptedFlag) -ne 0) {
            # 同じドライブのつもりが別のボリュームだったとき（マウントしたフォルダなど）も、暗号化されたファイルは写さない
            Set-EntryFail $Entry '暗号化されている' '暗号化（EFS）されたファイルは、別のドライブへ移せません（暗号化が外れて写るため）'
            return 'Fail'
        }
        if (-not $renamed) {
            # 別のドライブへは、開いたハンドルから中身を読んで写し、確かめてから、元のファイルに削除の印を付ける
            # （パスで読むと、確かめた後で途中のフォルダを差し替えられたとき、別のファイルの中身を写してしまうため）
            $copied = $false
            try {
                [void][WbRemoveOldFileNative]::CopyFromHandle($handle, $Entry.DestLong)
                $copied = $true
                $destInfo = [IO.FileInfo]::new($Entry.DestLong)
                if ($destInfo.Length -ne $Entry.Size) { throw 'コピーした後のサイズが元と違います' }
                # 属性（隠しなど）と日時を元に合わせる（移した先でも古さを判定できるように）
                $copyAttrs = $Entry.Attrs -band $CopyAttrMask
                if ($copyAttrs -eq 0) { $copyAttrs = [int][IO.FileAttributes]::Normal }
                [IO.File]::SetAttributes($Entry.DestLong, [IO.FileAttributes]$copyAttrs)
                [IO.File]::SetCreationTimeUtc($Entry.DestLong, $Entry.CreatedUtc)
                [IO.File]::SetLastWriteTimeUtc($Entry.DestLong, $Entry.WrittenUtc)
                # コピーしている間に変わったファイルは、元を残す
                if (Test-StateChanged $Entry ([WbRemoveOldFileNative]::Query($handle))) { throw '移している途中で、元のファイルが変更されました' }
                [WbRemoveOldFileNative]::MarkForDelete($handle)
            } catch {
                $ex = $_.Exception
                if ($ex -is [Management.Automation.RuntimeException] -and -not $ex.InnerException) { $reason = @('その他', $ex.Message) }
                else { $reason = Get-FailReason (Get-InnerException $ex) }
                $cleanup = ''
                if ($copied) {
                    $cleanup = '（移した先のコピーは削除しました）'
                    try { [IO.File]::Delete($Entry.DestLong) } catch { $cleanup = '（移した先のコピーを削除できませんでした: {0}）' -f $Entry.DestDisplay }
                }
                Set-EntryFail $Entry $reason[0] ($reason[1] + $cleanup)
                return 'Fail'
            }
        }
        $Entry.Result = $Words.Done
        $Entry.MovedTo = $Entry.DestDisplay
        return 'Done'
    } catch {
        $reason = Get-FailReason (Get-InnerException $_.Exception)
        Set-EntryFail $Entry $reason[0] $reason[1]
        return 'Fail'
    } finally {
        # 削除の印を付けたものは、閉じたときに削除される
        $handle.Dispose()
    }
}

# 空になったフォルダを 1 つ削除する。ファイルと同じく、開いて実際の場所を確かめてから、開いたまま削除の印を付ける。
# 中にものが残っていれば、Windows が削除を断る（ERROR_DIR_NOT_EMPTY）。'Done'・'No'・'Fail' のどれかを返す
function Remove-EmptyFolder($Node) {
    $state = $null
    try {
        $handle = [WbRemoveOldFileNative]::OpenForChange($Node.Long, $true, $false, [ref]$state)
    } catch {
        $ex = Get-InnerException $_.Exception
        $code = Get-ErrorCode $ex
        if ($code -eq 2 -or $code -eq 3) {
            $Node.Result = $StEmptyNo
            $Node.Note = '予定を立てた後になくなっていました'
            return 'No'
        }
        $Node.Result = $StFail
        $Node.Note = (Get-FailReason $ex)[1]
        return 'Fail'
    }
    try {
        if ($state.IsLink -or -not [string]::Equals($state.FinalPath, (Get-ExpectedFinal $Node $Node.Display), [StringComparison]::OrdinalIgnoreCase)) {
            $Node.Result = $StEmptyNo
            $Node.Note = '調べた後で、フォルダがリンクに変わっていました'
            return 'No'
        }
        try {
            [WbRemoveOldFileNative]::MarkForDelete($handle)
        } catch {
            $ex = Get-InnerException $_.Exception
            if ((Get-ErrorCode $ex) -eq 145) {
                $Node.Result = $StEmptyNo
                $Node.Note = '中にファイルかフォルダが残っています'
                return 'No'
            }
            $Node.Result = $StFail
            $Node.Note = (Get-FailReason $ex)[1]
            return 'Fail'
        }
        $Node.Result = $StDone
        return 'Done'
    } finally {
        $handle.Dispose()
    }
}

#endregion

# 調べている途中や削除の途中で Ctrl+C されたときに、どこまで進んだかを残せるようにする
$finished      = $false
$unexpected    = $false
$saveError     = $null
$actStarted    = $false
$answeredNo    = $false
$asking        = $false   # 確認の入力を待っている間（Ctrl+C で止まったときに、確認の行の後ろで改行するため）
$sortedItems   = @()
$sortedFolders = @()
$emptyPlan     = @()
$lines         = [System.Collections.Generic.List[string]]::new()   # summary.txt に書く行

function Save-Result {
    if (-not $runDir) { return }
    try {
        # 保存先は \\?\ 付きで作っているため、長いパスでも書けるよう、ここも \\?\ を付ける
        [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in $sortedItems) {
            $rows.Add([object[]]@($entry.Display, $entry.Size.ToString('N0', $Invariant), (Format-LocalTime $entry.CreatedUtc $TimeFormat),
                (Format-LocalTime $entry.WrittenUtc $TimeFormat), (Get-AgeText $entry), $entry.Result, $entry.Note, $entry.MovedTo))
        }
        Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'files.csv')) $FileColumns $rows
        $folderRows = [System.Collections.Generic.List[object]]::new()
        foreach ($stat in $sortedFolders) {
            $folderRows.Add([object[]]@($stat.Display, $stat.Count.ToString('N0', $Invariant), $stat.Size.ToString('N0', $Invariant),
                (Format-LocalTime $stat.OldestUtc $TimeFormat), (Format-LocalTime $stat.NewestUtc $TimeFormat)))
        }
        Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'folders.csv')) $FolderColumns $folderRows
        if ($RemoveEmptyFolder) {
            $emptyRows = [System.Collections.Generic.List[object]]::new()
            foreach ($node in $emptyPlan) { $emptyRows.Add([object[]]@($node.Display, $node.Result, $node.Note)) }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'emptyfolders.csv')) $EmptyColumns $emptyRows
        }
        Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $skipped
        $script:saveError = $null
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
function Write-Table([string]$Header, [string[]]$Rows, [int]$Limit, [string]$MoreFormat) {
    Write-Line $Header
    for ($k = 0; $k -lt $Rows.Count; $k++) {
        $lines.Add($Rows[$k])
        if ($k -lt $Limit -or -not $runDir) { Write-Host $Rows[$k] }
        elseif ($k -eq $Limit) { Write-Host (Format-Invariant $MoreFormat ($Rows.Count - $Limit)) }
    }
}

try {
    #region 調べる

    Enable-CtrlCInput
    foreach ($root in $roots) {
        $prefixLen = $root.Long.TrimEnd('\').Length + 1
        # 再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる。フォルダは、親より後に、たどった順に並ぶ
        $stack = [System.Collections.Generic.Stack[object]]::new()
        $stack.Push([pscustomobject]@{ Info = [IO.DirectoryInfo]::new($root.Long); Parent = $null })
        while ($stack.Count -gt 0) {
            Stop-IfCtrlC
            $frame = $stack.Pop()
            $dir = $frame.Info
            $dirAttrs = 0
            try { $dirAttrs = [int]$dir.Attributes } catch { }
            # Root は、このフォルダを含む -Path（削除するときに、実際の場所を確かめるため）
            $node = [pscustomobject]@{
                Display = (ConvertTo-DisplayPath $dir.FullName); Long = $dir.FullName; Parent = $frame.Parent; IsRoot = ($null -eq $frame.Parent); Root = $root
                Attrs = $dirAttrs; Entries = 0; Targets = 0; EmptyChildren = 0; Readable = $false; Empty = $false; Result = ''; Note = ''
            }
            $dirNodes.Add($node)
            if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                Write-Progress -Activity $ToolTitle -Status (Format-Invariant '調べています: ファイル {0:N0} 件  古いファイル {1:N0} 件' $fileCount $oldItems.Count) -CurrentOperation $node.Display
                $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
            }
            try {
                $entries = $dir.GetFileSystemInfos()
            } catch {
                $ex = Get-InnerException $_.Exception
                # 種類を判定できないリンクは、読み取りが禁止されていると普通のフォルダに見える。権限なしには数えない
                if ($ex -is [UnauthorizedAccessException] -and ($dir.Attributes -band $rpFlag) -eq $rpFlag) {
                    $linkCount++
                    Add-Skip 'リンク' $node.Display '読み取りが禁止されたリンク'
                    continue
                }
                Add-SkipError $ex $node.Display
                continue
            }
            $node.Readable = $true
            # 空になったフォルダを判定するため、中にあるものはすべて数える（除外したもの・リンク・古くないファイルも）
            $node.Entries = $entries.Count
            foreach ($entry in $entries) {
                $isDir = $entry -is $dirType
                # このツールの結果のフォルダは、中を見ない（前の実行の記録を消さないように）
                if ($isDir -and $OutputNameRule.IsMatch($entry.Name)) {
                    $outputDisplay = ConvertTo-DisplayPath $entry.FullName
                    if (-not [string]::Equals($outputDisplay, $runDir, [StringComparison]::OrdinalIgnoreCase)) {
                        $excludedCount++
                        Add-Skip '除外指定' $outputDisplay 'このツールの結果のフォルダ'
                    }
                    continue
                }
                # ごみ箱と System Volume Information は、Windows が管理するため調べない
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
                    $stack.Push([pscustomobject]@{ Info = $entry; Parent = $node })
                    continue
                }
                if ($hasInclude -and -not (Test-IncludeMatch $entry.Name)) { continue }
                if ($selfLong -and [string]::Equals($entry.FullName, $selfLong, [StringComparison]::OrdinalIgnoreCase)) {
                    $excludedCount++
                    Add-Skip '除外指定' (ConvertTo-DisplayPath $entry.FullName) 'このツールのスクリプト'
                    continue
                }
                $fileCount++
                # 日時が壊れたファイル（範囲外）は、日時を読むと $null が返るか、例外になる
                $createdUtc = $null
                $writtenUtc = $null
                $timeError  = $null
                try { $createdUtc = $entry.CreationTimeUtc; $writtenUtc = $entry.LastWriteTimeUtc } catch { $timeError = Get-ErrorText (Get-InnerException $_.Exception) }
                if ($timeError -or $null -eq $createdUtc -or $null -eq $writtenUtc) {
                    if (-not $timeError) { $timeError = '作成日時か最終更新日時が範囲外です' }
                    Add-Skip 'その他のエラー' (ConvertTo-DisplayPath $entry.FullName) $timeError
                    continue
                }
                # 作成日時と最終更新日時の新しいほうで判定する（コピーして置いた古いファイルは、作成日時が置いた日時になる）
                $newerUtc = if ($writtenUtc -gt $createdUtc) { $writtenUtc } else { $createdUtc }
                if ($newerUtc -ge $cutoffUtc) { continue }
                $display = ConvertTo-DisplayPath $entry.FullName

                $attrs = [int]$entry.Attributes
                $outKind = $null
                $note = ''
                if (($attrs -band $ReadOnlyFlag) -ne 0) { $outKind = '読み取り専用'; $note = '読み取り専用のファイルです' }
                elseif (($attrs -band $SystemFlag) -ne 0) { $outKind = 'システムファイル'; $note = 'システムファイルです' }
                elseif (($attrs -band $OnlineMask) -ne 0) { $outKind = 'クラウドのみ'; $note = 'クラウドのみのファイルです' }
                $size = $entry.Length
                $result = if ($outKind) { $StOut } else { '' }
                if ($nativeReady) {
                    $oldItems.Add([WbRemoveOldFileItem]::new($display, $size, $createdUtc, $writtenUtc, $attrs, $node, $outKind, $result, $note))
                } else {
                    # Add-Type を使えないとき（一覧だけ）は、同じ名前の項目を [pscustomobject] で持つ
                    $oldItems.Add([pscustomobject]@{
                        Display = $display; Size = $size; CreatedUtc = $createdUtc; WrittenUtc = $writtenUtc; Attrs = $attrs; Node = $node
                        IsTarget = ($null -eq $outKind); OutKind = $outKind; Result = $result; Note = $note; Reason = ''; MovedTo = ''; DestDisplay = ''; DestLong = ''
                    })
                }
                $stat = $null
                if ($folderStats.TryGetValue($node.Display, [ref]$stat)) {
                    $stat.Count++
                    $stat.Size += $size
                    if ($newerUtc -lt $stat.OldestUtc) { $stat.OldestUtc = $newerUtc }
                    if ($newerUtc -gt $stat.NewestUtc) { $stat.NewestUtc = $newerUtc }
                } else {
                    $folderStats[$node.Display] = [pscustomobject]@{ Display = $node.Display; Count = 1; Size = [long]$size; OldestUtc = $newerUtc; NewestUtc = $newerUtc }
                }
            }
        }
    }
    Write-Progress -Activity $ToolTitle -Completed
    $scanElapsed = $sw.Elapsed

    #endregion

    #region 予定を立てる

    # ファイルはパスの順に並べる
    $itemArray = $oldItems.ToArray()
    $keys = [string[]]::new($itemArray.Count)
    for ($k = 0; $k -lt $itemArray.Count; $k++) { $keys[$k] = Get-SortText $itemArray[$k].Display }
    $sortedItems = Get-SortedList $keys $itemArray

    # 移すときは、移す先のパスを決め、同じ名前のファイルがあるものは移さない（上書きしない）
    if ($mode -eq 'Move') {
        $rootPrefixLen = $roots[0].Full.TrimEnd('\').Length + 1
        foreach ($entry in $sortedItems) {
            if (-not $entry.IsTarget) { continue }
            Stop-IfCtrlC
            $entry.DestDisplay = $moveFull.TrimEnd('\') + '\' + $entry.Display.Substring($rootPrefixLen)
            $entry.DestLong = ConvertTo-LongPath $entry.DestDisplay
            $exists = $false
            try { $exists = [IO.File]::Exists($entry.DestLong) -or [IO.Directory]::Exists($entry.DestLong) } catch { }
            if ($exists) {
                $entry.IsTarget = $false
                $entry.OutKind = '移す先に同じ名前がある'
                $entry.Result = $StCannot
                $entry.Reason = '同じ名前がある'
                $entry.Note = '移す先に同じ名前のファイルがあります: {0}' -f $entry.DestDisplay
            } elseif (-not $sameVolume -and ($entry.Attrs -band $EncryptedFlag) -ne 0) {
                # 別のドライブへは中身を読んで写すため、暗号化（EFS）されたファイルは、暗号化が外れて写ってしまう
                $entry.IsTarget = $false
                $entry.OutKind = '暗号化されている'
                $entry.Result = $StCannot
                $entry.Reason = '暗号化されている'
                $entry.Note = '暗号化（EFS）されたファイルは、別のドライブへ移せません（暗号化が外れて写るため）'
            }
        }
    }

    $targetCount = 0
    $targetBytes = 0L
    $oldBytes    = 0L
    $outKinds = [ordered]@{}
    if ($mode -eq 'Move') { $outKinds['移す先に同じ名前がある'] = 0; $outKinds['暗号化されている'] = 0 }
    foreach ($name in '読み取り専用', 'システムファイル', 'クラウドのみ') { $outKinds[$name] = 0 }
    foreach ($entry in $sortedItems) {
        $oldBytes += $entry.Size
        if ($entry.IsTarget) {
            $targetCount++
            $targetBytes += $entry.Size
            $entry.Node.Targets++
        } else {
            $outKinds[$entry.OutKind]++
        }
    }
    $restCount = $sortedItems.Count - $targetCount

    # 空になるフォルダ。中にあるものが、対象のファイルと、空になるフォルダだけのフォルダ（子は親より後に並んでいるので、後ろから決める）。
    # 実行する前から空だったフォルダは数えない。読み取り専用の属性のフォルダは、ファイルと同じく削除しない（対象外）
    $emptyCount = 0
    if ($RemoveEmptyFolder) {
        for ($k = $dirNodes.Count - 1; $k -ge 0; $k--) {
            $node = $dirNodes[$k]
            if ($node.IsRoot -or -not $node.Readable -or $node.Entries -eq 0 -or $node.Entries -ne ($node.Targets + $node.EmptyChildren)) { continue }
            if (($node.Attrs -band $ReadOnlyFlag) -ne 0) {
                $node.Result = $StOut
                $node.Note = '読み取り専用のフォルダです'
                continue
            }
            $node.Empty = $true
            $emptyCount++
            $node.Parent.EmptyChildren++
        }
        $emptyArray = @(foreach ($node in $dirNodes) { if ($node.Empty -or $node.Result -eq $StOut) { $node } })
        $emptyKeys = @(foreach ($node in $emptyArray) { Get-SortText $node.Display })
        $emptyPlan = Get-SortedList $emptyKeys $emptyArray
    }

    # フォルダは、古いファイルの合計サイズの大きい順（同じならパスの順）
    $folderArray = @($folderStats.Values)
    $folderKeys = @(foreach ($stat in $folderArray) { ([long]::MaxValue - $stat.Size).ToString('D19', $Invariant) + "`0" + (Get-SortText $stat.Display) })
    $sortedFolders = Get-SortedList $folderKeys $folderArray

    #endregion

    #region 一覧の表示

    Write-Line (Format-Invariant '{0} (Remove-OldFile {1})' $ToolTitle $ScriptVersion)
    Write-Line ('対象: {0}' -f (($roots | ForEach-Object { $_.Full }) -join ', '))
    if ($mode -eq 'Move') { Write-Line ('移す先: {0}' -f $moveFull) }
    Write-Line (Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（管理者: {2}）' $startTime $machineName $(if ($isAdmin) { 'はい' } else { 'いいえ' }))
    $condition = Format-Invariant '条件: 作成・最終更新とも {0} 日以上前（{1:yyyy-MM-dd HH:mm} より前）のファイル' $days $cutoff
    if ($hasInclude) { $condition += '  名前: ' + (($includeRules | ForEach-Object { $_.Text }) -join ', ') }
    Write-Line $condition
    if ($hasExclude) { Write-Line ('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', ')) }
    Write-Line (Format-Invariant 'ファイル {0:N0} 件  古いファイル {1:N0} 件（{2}）  所要 {3}' $fileCount $sortedItems.Count (Format-Size $oldBytes) (Format-Duration $scanElapsed))
    $noteLine = Get-NoteLine
    if ($noteLine) { Write-Line $noteLine }
    Write-Line ''

    Write-Line '■ 古いファイルのあるフォルダ（サイズの大きい順）'
    if ($sortedFolders.Count -eq 0) {
        Write-Line '    （古いファイルはありません）'
    } else {
        $header = (Format-Column 'サイズ' 12 $true) + '  ' + (Format-Column '件数' 8 $true) + '  ' + (Format-Column 'いちばん古い' 12 $false) + '  フォルダ'
        $rows = @(foreach ($stat in $sortedFolders) {
            (Format-Column (Format-Size $stat.Size) 12 $true) + '  ' + (Format-Column ($stat.Count.ToString('N0', $Invariant)) 8 $true) + '  ' +
                (Format-Column (Format-LocalTime $stat.OldestUtc 'yyyy-MM-dd') 12 $false) + '  ' + $stat.Display
        })
        Write-Table $header $rows $FolderRows '（ほか {0:N0} フォルダは folders.csv を参照してください）'
    }
    Write-Line ''

    $restText = Get-KindText $outKinds
    $deniedSeen = $skipKinds['権限なし'] -gt 0
    if (-not $acting) {
        if ($targetCount -eq 0) {
            Write-Line '対象のファイルはありません'
            if ($restCount -gt 0) { Write-Line (Format-Invariant '対象外 {0:N0} 件（{1}）' $restCount $restText) }
        } else {
            $countLine = Format-Invariant '対象 {0:N0} 件（{1}）' $targetCount (Format-Size $targetBytes)
            if ($restCount -gt 0) { $countLine += Format-Invariant '  対象外 {0:N0} 件（{1}）' $restCount $restText }
            Write-Line $countLine
        }
        if ($RemoveEmptyFolder) { Write-Line (Format-Invariant '空になる予定のフォルダ {0:N0} 件' $emptyCount) }
        if (-not $isAdmin -and $deniedSeen) { Write-Line '※ 管理者として実行すると、対象にできる範囲が広がります' }
        if ($targetCount -gt 0) { Write-Line '削除するときは -Delete を、移すときは -MoveTo を付けて、もう一度実行してください' }
    }

    #endregion

    #region 削除・移動

    if ($acting) {
        if ($targetCount -gt 0) { Write-Line (Format-Invariant '{0} {1:N0} 件（{2}）' $Words.Files $targetCount (Format-Size $targetBytes)) }
        if ($restCount -gt 0) { Write-Line (Format-Invariant '{0} {1:N0} 件（{2}）' $Words.Rest $restCount $restText) }
        if (-not $isAdmin -and $deniedSeen) { Write-Line '※ 管理者として実行すると、対象にできる範囲が広がります' }

        if ($targetCount -eq 0) {
            Write-Line '対象のファイルはありません'
        } else {
            # 別のドライブへ移すときは、移す先の空き容量が足りるかを、確認の前に確かめる
            if ($mode -eq 'Move' -and -not $sameVolume) {
                $free = Get-FreeSpace $moveFull
                if ($free -ge 0 -and $free -lt $targetBytes) {
                    $driveText = $moveFull
                    try { $driveText = [IO.Path]::GetPathRoot($moveFull) } catch { }
                    Stop-WithError ('移す先の空き容量が足りないため、移さずに止めました（必要 {0}・空き {1}）: {2}' -f (Format-Size $targetBytes), (Format-Size $free), $driveText)
                }
            }
            if ($RemoveEmptyFolder -and $emptyCount -gt 0) { Write-Line (Format-Invariant '空になるフォルダ {0:N0} 件も削除します' $emptyCount) }
            if ($mode -eq 'Delete') { Write-Line '※ 削除したファイルは元に戻せません（ごみ箱には入りません）' }

            $answer = 'Y'
            if (-not $Force) {
                # 画面のない実行（タスクスケジューラの「ログオンしているかどうかにかかわらず実行する」など）で、確認待ちのまま止まり続けないようにする
                if (-not [Environment]::UserInteractive) { Stop-WithError $Words.CannotAsk }
                # 確認を待つ間は、通常の Ctrl+C に戻す（キーの入力として受け取ると、Read-Host に Ctrl+C の文字が入るため）
                Stop-IfCtrlC
                Disable-CtrlCInput
                $answer = ''
                $asking = $true
                while ($answer -notmatch '^[YyNnＹｙＮｎ]') {
                    # -NonInteractive で起動されていると、Read-Host は英語のエラーで止まる
                    try { $answer = Read-Host $Words.Ask } catch { Stop-WithError $Words.CannotAsk }
                    if ($null -eq $answer) { Stop-WithError $Words.CannotAsk }   # 入力が終わっている（リダイレクトされた入力の終わり）
                    $answer = $answer.Trim()
                }
                $asking = $false
                $lines.Add(('{0}: {1}' -f $Words.Ask, $answer))
            } else {
                $lines.Add('-Force を指定したため、確認していません')
            }

            if ($answer -match '^[NnＮｎ]') {
                $answeredNo = $true
                foreach ($entry in $sortedItems) { if ($entry.IsTarget) { $entry.Result = $Words.No } }
                foreach ($node in $emptyPlan) { if ($node.Empty) { $node.Result = $StEmptyNo } }
                Write-Line $Words.DidNot
            } else {
                # 何を削除しようとしているかを、削除する前に保存しておく（途中で止まっても残るように）
                foreach ($entry in $sortedItems) { if ($entry.IsTarget) { $entry.Result = $Words.Plan } }
                foreach ($node in $emptyPlan) { if ($node.Empty) { $node.Result = '削除する予定' } }
                $actStarted = $true
                Save-Result
                if ($saveError) {
                    $actStarted = $false
                    Stop-WithError ('記録を保存できないため、{0}ずに止めました: {1}' -f $(if ($mode -eq 'Move') { '移さ' } else { '削除せ' }), $saveError)
                }
                Enable-CtrlCInput
                $actWatch = [Diagnostics.Stopwatch]::StartNew()
                $doneCount = 0
                $doneBytes = 0L
                $failCount = 0
                $goneCount = 0
                $createdDirs = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($entry in $sortedItems) {
                    if (-not $entry.IsTarget) { continue }
                    # 押されていたら、ここ（前のファイルを終えた後）で止める
                    Stop-IfCtrlC
                    if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                        Write-Progress -Activity $ToolTitle -Status (Format-Invariant '{0}: {1:N0} / {2:N0} 件' $(if ($mode -eq 'Move') { '移しています' } else { '削除しています' }) ($doneCount + $failCount + $goneCount) $targetCount) -CurrentOperation $entry.Display
                        $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
                    }
                    # 予定を立てた後（確認を待っている間など）に、なくなったファイル・変わったファイル・場所が変わったファイルは、削除しない
                    $long = ConvertTo-LongPath $entry.Display
                    switch (Invoke-FileAction $entry $long) {
                        'Done' { $doneCount++; $doneBytes += $entry.Size }
                        'Gone' { $goneCount++ }
                        default { $failCount++ }
                    }
                }

                # 空になったフォルダを、深いほうから削除する。削除する直前に、空であることと、実際の場所をもう一度確かめる
                $emptyDone = 0
                $emptyNo   = 0
                $emptyFail = 0
                if ($RemoveEmptyFolder) {
                    for ($k = $dirNodes.Count - 1; $k -ge 0; $k--) {
                        $node = $dirNodes[$k]
                        if (-not $node.Empty) { continue }
                        Stop-IfCtrlC
                        switch (Remove-EmptyFolder $node) {
                            'Done' { $emptyDone++ }
                            'No' { $emptyNo++ }
                            default { $emptyFail++ }
                        }
                    }
                }
                Write-Progress -Activity $ToolTitle -Completed

                Write-Line ''
                $resultLine = Format-Invariant '結果: {0}  {1:N0} 件（{2}）  失敗 {3:N0} 件  所要 {4}' $Words.DidAll $doneCount (Format-Size $doneBytes) $failCount (Format-Duration $actWatch.Elapsed)
                if ($goneCount -gt 0) { $resultLine += Format-Invariant '  なくなっていた {0:N0} 件' $goneCount }
                Write-Line $resultLine
                if ($RemoveEmptyFolder -and $emptyCount -gt 0) {
                    $emptyLine = Format-Invariant '空になったフォルダ {0:N0} 件を削除しました' $emptyDone
                    $emptyRest = @()
                    if ($emptyNo -gt 0) { $emptyRest += Format-Invariant '削除しなかった {0:N0} 件' $emptyNo }
                    if ($emptyFail -gt 0) { $emptyRest += Format-Invariant '失敗 {0:N0} 件' $emptyFail }
                    if ($emptyRest.Count -gt 0) { $emptyLine += '（' + ($emptyRest -join '・') + '）' }
                    Write-Line $emptyLine
                }
                $deniedNote = (Get-FailReason ([UnauthorizedAccessException]::new()))[1]
                if (-not $isAdmin -and (@($sortedItems | Where-Object { $_.Reason -eq 'アクセス拒否' }).Count -gt 0 -or @($emptyPlan | Where-Object { $_.Note -eq $deniedNote }).Count -gt 0)) {
                    Write-Line $Words.AdminAfter
                }
            }
        }

        # 削除できなかった（移せなかった）ファイル
        $failed = @($sortedItems | Where-Object { $_.Result -eq $StFail -or $_.Result -eq $StCannot })
        if ($failed.Count -gt 0 -and -not $answeredNo) {
            Write-Line ''
            $widthReason = 4
            foreach ($entry in $failed) { $widthReason = [math]::Max($widthReason, (Get-DisplayWidth $entry.Reason)) }
            $rows = @(foreach ($entry in $failed) { ' ' + (Format-Column $entry.Reason $widthReason $false) + '  ' + $entry.Display })
            Write-Line $Words.FailTitle
            Write-Table (' ' + (Format-Column '理由' $widthReason $false) + '  ファイル') $rows $FailRows '（ほか {0:N0} 件は files.csv を参照してください）'
            Write-Line ''
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
    if (-not $finished -and $stopping) {
        # 止めた理由は、Stop-WithError が表示している
        if ($runDir -and -not $actStarted) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
    } elseif (-not $finished) {
        if ($actStarted) {
            # 削除（移動）の途中で止まったときは、削除したファイルの記録を残す
            foreach ($entry in $sortedItems) {
                if ($entry.IsTarget -and $entry.Result -eq $Words.Plan) { $entry.Result = $Words.No; $entry.Note = $Words.Interrupted }
            }
            foreach ($node in $emptyPlan) {
                if ($node.Result -eq '削除する予定') { $node.Result = $StEmptyNo; $node.Note = '中断したため、削除していません' }
            }
            $stopText = if ($unexpected) { 'エラーで止まりました' } else { '中断しました' }
            $lines.Add(('{0}。それまでに{1}ファイルの記録は残っています' -f $stopText, $Words.Stopped))
            Save-Result
            Write-Host ('{0}。それまでに{1}ファイルの記録は残っています: {2}' -f $stopText, $Words.Stopped, $runDir) -ForegroundColor Yellow
        } else {
            if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
            # 5.1 は、確認の入力を Ctrl+C でやめたときに改行しない。確認の行（「…[N] いいえ:」）に続けて表示しないようにする
            if ($asking) { Write-Host '' }
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

# 削除するとき・移すときに記録を保存できなかったときは、気づけるよう 1 にする
if ($acting -and $saveError) { exit 1 }
# 確認で N を選んだときは、移せないファイルがあっても 0
if ($acting -and -not $answeredNo) {
    $failedCount = @($sortedItems | Where-Object { $_.Result -eq $StFail -or $_.Result -eq $StCannot }).Count
    $failedCount += @($emptyPlan | Where-Object { $_.Result -eq $StFail }).Count
    if ($failedCount -gt 0) { exit 1 }
}
exit 0
