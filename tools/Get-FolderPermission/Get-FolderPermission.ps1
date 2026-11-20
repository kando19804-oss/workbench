<#
.SYNOPSIS
    共有フォルダとその中のフォルダのアクセス権を、共有のアクセス許可と NTFS のアクセス権の両方から一覧にします。

.DESCRIPTION
    共有フォルダとその中のフォルダについて、誰に・どんな権限が付いているかを一覧にし、CSV に残します。
    親のアクセス権を継承しているだけのフォルダは出さず、継承を切っているフォルダ、アクセス権を追加しているフォルダ、継承が親と
    合わないフォルダ（同じドライブの中で移動したものなど）だけを出します（-All を付けると、すべてのフォルダを出します）。
    継承を切っているところ、継承が親と合わないところ、個人を直接指定しているところ、名前が見つからない（削除された）アカウント、
    広い範囲に書き込みを許可しているところ、拒否の設定、並び順が正しくないアクセス権に印を付けます。

    読むだけで、アクセス権を変更することはありません。
    ジャンクションやシンボリックリンクはたどりません。権限がなくて読めないフォルダはスキップし、件数を表示します。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Get-FolderPermission.ps1

.PARAMETER Path
    調べるフォルダ。この PC のフォルダ、共有フォルダ（\\fs01\営業）、ネットワークドライブを指定できます。
    \\fs01 のようにサーバー名だけを書くと、そのサーバーの共有フォルダをすべて調べます。
    複数指定するときは、カンマで区切って 1 つの文字列にします。省略すると、この PC の共有フォルダをすべて調べます。

.PARAMETER Depth
    何階層下まで調べるか（1～100）。-Path の直下が 1 です。省略すると、すべての階層を調べます。

.PARAMETER All
    アクセス権を変えていないフォルダも含めて、すべてのフォルダのアクセス権を出します。

.PARAMETER IncludeFile
    ファイルも調べて、アクセス権を変えてあるファイル（継承を切っている・追加している・継承が親と合わない）を出します。

.PARAMETER ExpandGroup
    アクセス権に出てくるグループのメンバーを調べて、groups.csv に出します。

.PARAMETER Exclude
    調べないフォルダやファイル。ワイルドカード（*）で、カンマで区切って書きます。
    \ を含まない書き方は名前に、ドライブや \\ で始まる書き方はフルパスに、それ以外は -Path から見た相対パスに当てはめます。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.PARAMETER _UnknownArgument
    指定しません。どのパラメーターにも当たらない引数（パラメーターの名前の書き間違いなど）を受け取り、
    案内を出して止まるためのものです。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-FolderPermission.ps1

    この PC の共有フォルダをすべて調べます（ファイルサーバーの上で実行します）。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-FolderPermission.ps1 -Path \\fs01

    サーバー fs01 の共有フォルダをすべて調べます。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-FolderPermission.ps1 -Path D:\share\営業 -IncludeFile -ExpandGroup

    D:\share\営業 から下のフォルダとファイルを調べ、アクセス権に出てくるグループのメンバーも調べます。

.NOTES
    Get-FolderPermission 1.0.0
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

    [string]$Depth,

    [switch]$All,

    [switch]$IncludeFile,

    [switch]$ExpandGroup,

    [string[]]$Exclude = @(),

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
$ToolTitle      = 'アクセス権一覧'
$ItemRows       = 20        # 画面に表示するフォルダの数（保存するときだけ。超えた分はファイルを見てもらう）
$ExampleCount   = 3         # 「確かめたほうがよいところ」に出す例の数
$ProgressMs     = 250       # 進捗表示の間隔（ミリ秒）
$MaxDepth       = 100
$OutputPrefix   = 'FolderPermission_'
$SystemFolders  = '$RECYCLE.BIN', 'System Volume Information'
# 「すべての共有」に含めない、Windows が自動で作る共有（名前の最後が $ の管理用の共有は、種類の印で見分ける）
$AutoShares     = @('print$')
$ShareColumns   = 'サーバー', '共有名', 'フォルダ', '説明', '種類', 'アカウント', 'アカウントの種類', 'アクセス許可', '値', 'SID', '備考'
$ItemColumns    = 'パス', '種類', '階層', '継承', '直接の設定', '継承した設定', '所有者', '共有名', '注意'
$AceColumns     = 'パス', '種類', '階層', '継承', '許可・拒否', 'アカウント', 'アカウントの種類', 'アクセス権', '適用先', '設定した場所', '詳細', '値', '注意', 'SID'
$GroupColumns   = 'グループ', 'メンバー', 'メンバーの種類', '経由', '状態', 'SID', '備考'
$SkipColumns    = '理由', 'パス', '詳細'
$Invariant      = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom        = [Text.UTF8Encoding]::new($true)
$OnWindows      = [IO.Path]::DirectorySeparatorChar -eq '\'
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

function Format-Count([long]$Value) {
    return $Value.ToString('N0', $Invariant)
}

# .NET や COM のメソッドで起きた例外は MethodInvocationException などに包まれて届くため、元の例外を取り出す。
# InvokeMember（グループのメンバーを読むとき）の例外は TargetInvocationException に包まれる
function Get-InnerException([Exception]$Exception) {
    while (($Exception -is [Management.Automation.ExtendedTypeSystemException] -or $Exception -is [Reflection.TargetInvocationException]) -and $Exception.InnerException) {
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

# Windows のエラー番号の説明（OS の言語の文）
function Get-Win32Text([int]$Code) {
    return Get-ErrorText ([ComponentModel.Win32Exception]::new($Code))
}

# CSV の 1 行（組み立ては、数が多いと速いよう C# で行う。この関数を呼ぶのは、Add-Type の後だけ）
function ConvertTo-CsvLine([object[]]$Fields) {
    $parts = [string[]]::new($Fields.Count)
    for ($i = 0; $i -lt $Fields.Count; $i++) { $parts[$i] = [string]$Fields[$i] }
    return [WbFolderPermNative]::CsvLine($parts)
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

# Excel が数式や数値に変えてしまう値を ="..." で囲む（中身は C# の ProtectExcel。この関数を呼ぶのは、Add-Type の後だけ）
function Protect-ExcelText([string]$Text) {
    return [WbFolderPermNative]::ProtectExcel($Text)
}

# 画面の表で列をそろえるため、全角文字を 2 文字分の幅として数える（中身は C#。この関数を呼ぶのは、Add-Type の後だけ）
function Get-DisplayWidth([string]$Text) {
    return [WbFolderPermNative]::DisplayWidth($Text)
}

function Format-Column([string]$Text, [int]$Width, [bool]$AlignRight) {
    return [WbFolderPermNative]::Pad($Text, $Width, $AlignRight)
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
# 後ろの引数までその値に入り、空白を含む値の残りが余るか、位置で受けるパラメーターに入る。どちらでも、引用符の書き方を先に案内する。
# カンマ区切りの値は、要素ごとに前後の引用符を外してから確かめる
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
Assert-CommandArgument $PSBoundParameters @('Path', 'Depth', 'Exclude', 'OutputFolder')

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

# 文字の番号の順に並べるための鍵（Sort-Object の文字列の比べ方は 5.1 と 7 で違い、記号の扱いで順が入れ替わるため）。
# パスは、\ を最も小さい文字に置き換えて、フォルダの中身が、名前の続く隣のフォルダ（「営業 2」など）より先に並ぶようにする
function Get-SortText([string]$Text) {
    return $Text.ToUpperInvariant()
}

function Get-PathSortText([string]$Text) {
    return $Text.ToUpperInvariant().Replace('\', [string][char]1)
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

# -Depth は、範囲外や数値でない値を日本語で案内するため、文字で受け取って確かめる（0 はすべての階層）
$depthLimit = 0
if ($PSBoundParameters.ContainsKey('Depth')) {
    $depthText = (Remove-Quote $Depth).Trim("'")
    $normalizedDepth = $depthText
    try { $normalizedDepth = $depthText.Normalize([Text.NormalizationForm]::FormKC) } catch { }
    if ($normalizedDepth -notmatch '^[0-9]{1,4}$' -or -not [int]::TryParse($normalizedDepth, [Globalization.NumberStyles]::None, $Invariant, [ref]$depthLimit) -or $depthLimit -lt 1 -or $depthLimit -gt $MaxDepth) {
        Stop-WithError (Format-Invariant '-Depth は 1～{0} の数値で指定してください: {1}' $MaxDepth $depthText)
    }
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

#endregion

#region Windows の API

# Windows の API を呼ぶ部品（Add-Type で C# をコンパイルする）。
# - アクセス権（所有者と DACL）を、260 文字を超えるパスでも読み、継承を切っているか・追加しているかを確かめる
# - アクセス権の中身を、エクスプローラーと同じ区切りの名前に直し、アカウントの名前と種類を調べる（SID ごとに 1 回だけ問い合わせる）
# - 共有フォルダの一覧と、共有のアクセス許可を読む
# - フォルダの数が多いときに 5.1 で遅くならないよう、記録は軽い入れ物（C# のクラス）で持つ
$NativeSource = @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

public class WbFolderPermAccount {
    public string Sid = "";
    public string Name = "";
    public string Domain = "";
    public string ShortName = "";
    public string Kind = "";
    public int Use;
    public int Error;
    public bool IsWide;
}

public class WbFolderPermAce {
    public bool Deny;
    public bool Inherited;
    public int Flags;
    public uint Mask;
    public WbFolderPermAccount Account;
    public string Rights = "";
    public string Detail = "";
    public string AppliesTo = "";
    public int Notes;
}

public class WbFolderPermItem {
    public string Display = "";
    // 名前を問い合わせる PC（この PC は空）。グループのメンバーを調べるときに使う
    public string LookupSystem = "";
    public bool IsFile;
    public int Depth;
    public bool IsRoot;
    public string ShareName = "";
    public bool Protected;
    public bool Stale;
    public bool NullDacl;
    public bool Canonical = true;
    public int ExplicitCount;
    public int InheritedCount;
    public string Owner = "";
    public WbFolderPermAce[] Aces = new WbFolderPermAce[0];
    public int Notes;
}

public class WbFolderPermSecurity {
    public int Error;
    public byte[] Descriptor;
    public bool Protected;
    public bool HasExplicit;
    // 継承の印の付いた設定がある・DACL に自動継承の印がある（共有のルートで、継承の印のない古い形の ACL かを見るため）
    public bool HasInherited;
    public bool AutoInherited;
    public bool NullDacl;
    // 継承した設定が、親から継承するはずの設定と合わない（同じドライブの中で移動したフォルダなど）
    public bool Stale;
    public bool Changed { get { return Error == 0 && (Protected || HasExplicit || NullDacl || Stale); } }
}

public class WbFolderPermShare {
    public string Name = "";
    public uint Type;
    public string Remark = "";
}

public class WbFolderPermShareInfo {
    public int Error;
    public string Path = "";
    public string Remark = "";
    public byte[] Descriptor;
}

// 「確かめたほうがよいところ」の集計（数が多いと 5.1 のループでは遅いため、C# で数える）
public class WbFolderPermFindings {
    public int ProtectedFolders;
    public int ProtectedFiles;
    public List<string> Stale = new List<string>();
    public List<string> Order = new List<string>();
    // 個人を直接指定・名前が見つからない・広い範囲に書き込みを許可・拒否の件数と、例（重ならないもの）
    public int[] Counts = new int[4];
    public List<string>[] Examples = { new List<string>(), new List<string>(), new List<string>(), new List<string>() };
    public bool LookupFailed;
}

// 子が親から継承するはずのアクセス権。Windows は、同じ SID に同じ権限が効いているときに設定を省いたり、「下へ渡すだけ」の形に
// 変えたりするため、1 つずつの形ではなく、意味（子に効く権限と、下へ渡す権限）で比べる
public class WbFolderPermExpected {
    // 子に効く権限：「種類|SID」ごとの権限の合計
    public Dictionary<string, uint> Effective = new Dictionary<string, uint>(StringComparer.Ordinal);
    // CREATOR OWNER・CREATOR GROUP から子の所有者（グループ）に効く分。子を作ったときにだけでき、後から親の変更が伝わったときには
    // できないことがあるため、あってもなくてもよい。SID の部分は @OWNER・@GROUP
    public Dictionary<string, uint> Optional = new Dictionary<string, uint>(StringComparer.Ordinal);
    // 下へ渡す権限：「種類|SID|C（サブフォルダへ）か O（ファイルへ）」ごとの権限の合計。Windows は、ほかの設定に含まれる
    // 渡し方（CI＋OI の設定に含まれる CI だけの設定など）を省くため、渡し方の組み合わせではなく、渡る先ごとに足し合わせる
    public Dictionary<string, uint> Propagate = new Dictionary<string, uint>(StringComparer.Ordinal);
}

public static class WbFolderPermNative {
    public const int NoteProtected = 1;
    public const int NoteUser = 2;
    public const int NoteUnknown = 4;
    public const int NoteWide = 8;
    public const int NoteDeny = 16;
    public const int NoteOrder = 32;
    public const int NoteStale = 64;
    public const string KindUser = "ユーザー";
    public const string KindGroup = "グループ";
    public const string KindBuiltin = "組み込み";
    public const string KindComputer = "コンピューター";
    public const string KindApp = "Windows のアプリ";
    public const string KindUnknown = "名前が見つからない";
    public const string KindLookupFailed = "名前を調べられない";
    // 書き込み・削除・アクセス権の変更・所有権の取得（データの書き込み、フォルダの作成、サブフォルダとファイルの削除、削除、WRITE_DAC、WRITE_OWNER）
    const uint WriteBits = 0x2 | 0x4 | 0x40 | 0x10000 | 0x40000 | 0x80000;

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
    static extern int GetNamedSecurityInfoW(string name, int objectType, uint info, IntPtr owner, IntPtr group, IntPtr dacl, IntPtr sacl, out IntPtr descriptor);

    [DllImport("kernel32.dll")]
    static extern IntPtr LocalFree(IntPtr memory);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool LookupAccountSidW(string system, byte[] sid, StringBuilder name, ref uint nameLength, StringBuilder domain, ref uint domainLength, out int use);

    [DllImport("netapi32.dll", CharSet = CharSet.Unicode)]
    static extern int NetShareEnum(string server, int level, out IntPtr buffer, int preferredMax, out int read, out int total, ref int resume);

    [DllImport("netapi32.dll", CharSet = CharSet.Unicode)]
    static extern int NetShareGetInfo(string server, string share, int level, out IntPtr buffer);

    [DllImport("netapi32.dll", CharSet = CharSet.Unicode)]
    static extern int NetWkstaGetInfo(string server, int level, out IntPtr buffer);

    [DllImport("netapi32.dll")]
    static extern int NetApiBufferFree(IntPtr buffer);

    [DllImport("advapi32.dll")]
    static extern uint GetSecurityDescriptorLength(IntPtr descriptor);

    [DllImport("advapi32.dll", SetLastError = true)]
    static extern bool MakeSelfRelativeSD(IntPtr absolute, byte[] relative, ref uint length);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool GetVolumeInformationW(string root, StringBuilder volumeName, int volumeNameSize, out uint serial, out uint maxComponent, out uint flags, StringBuilder fileSystemName, int fileSystemNameSize);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern SafeFileHandle CreateFileW(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern uint GetFinalPathNameByHandleW(SafeFileHandle handle, StringBuilder buffer, uint length, uint flags);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct ShareInfo1 { public string Name; public uint Type; public string Remark; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct ShareInfo2 { public string Name; public uint Type; public string Remark; public uint Permissions; public uint MaxUses; public uint CurrentUses; public string Path; public string Password; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct ShareInfo502 { public string Name; public uint Type; public string Remark; public uint Permissions; public uint MaxUses; public uint CurrentUses; public string Path; public string Password; public uint Reserved; public IntPtr SecurityDescriptor; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct WkstaInfo100 { public uint PlatformId; public string ComputerName; public string LanGroup; public uint VersionMajor; public uint VersionMinor; }

    static Dictionary<string, WbFolderPermAccount> accounts = new Dictionary<string, WbFolderPermAccount>(StringComparer.OrdinalIgnoreCase);

    // 所有者と DACL を読む。260 文字を超えるパスは \\?\ 付きで渡す。
    // GetFileSecurity はディスクの上の中身をそのまま返すため、継承の印のない古い形の ACL（自動継承の印のないもの）では、
    // 親から継承した設定まで「このフォルダで設定した」に見える。GetNamedSecurityInfo は、そのときに親と見比べて継承の印を
    // 付け直す（エクスプローラーや Get-Acl と同じ見え方）ので、こちらを使う。
    // ただし共有のルート（\\fs01\営業）は、共有越しには親をたどれないため付け直せない（呼ぶ側で管理共有越しに読み直す）
    public static WbFolderPermSecurity ReadSecurity(string path) {
        WbFolderPermSecurity result = new WbFolderPermSecurity();
        IntPtr pointer;
        // SE_FILE_OBJECT (1)、OWNER (1) | GROUP (2) | DACL (4)。グループは、CREATOR GROUP の継承を確かめるため
        int error = GetNamedSecurityInfoW(path, 1, 7, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, out pointer);
        if (error != 0) { result.Error = error; return result; }
        byte[] descriptor;
        try {
            int size = (int)GetSecurityDescriptorLength(pointer);
            descriptor = new byte[size];
            Marshal.Copy(pointer, descriptor, 0, size);
        } finally {
            LocalFree(pointer);
        }
        int length = descriptor.Length;
        result.Descriptor = descriptor;
        if (length < 20) { result.NullDacl = true; return result; }
        int control = BitConverter.ToUInt16(descriptor, 2);
        int daclOffset = BitConverter.ToInt32(descriptor, 16);
        result.Protected = (control & 0x1000) != 0;
        result.AutoInherited = (control & 0x400) != 0;
        if ((control & 0x4) == 0 || daclOffset <= 0 || daclOffset + 8 > length) { result.NullDacl = true; return result; }
        int count = BitConverter.ToUInt16(descriptor, daclOffset + 4);
        int pos = daclOffset + 8;
        for (int i = 0; i < count && pos + 4 <= length; i++) {
            int size = BitConverter.ToUInt16(descriptor, pos + 2);
            if (size < 4 || pos + size > length) break;
            if ((descriptor[pos + 1] & 0x10) == 0) result.HasExplicit = true;
            else result.HasInherited = true;
            if (result.HasExplicit && result.HasInherited) break;
            pos += size;
        }
        return result;
    }

    // アクセス権の中身を、1 つの設定ごとに名前を付けて $item に入れる。mode は 0 フォルダ、1 ファイル、2 共有のアクセス許可。
    // descriptor が null のとき（共有にアクセス許可の設定がないとき）は、全員がフル コントロールとして扱う
    public static void Analyze(WbFolderPermItem item, byte[] descriptor, int mode, string system) {
        List<WbFolderPermAce> list = new List<WbFolderPermAce>();
        item.Canonical = true;
        item.NullDacl = false;
        item.ExplicitCount = 0;
        item.InheritedCount = 0;
        if (descriptor == null || descriptor.Length < 20) {
            item.NullDacl = true;
        } else {
            int control = BitConverter.ToUInt16(descriptor, 2);
            if (mode != 2) item.Protected = (control & 0x1000) != 0;
            int ownerOffset = BitConverter.ToInt32(descriptor, 4);
            int daclOffset = BitConverter.ToInt32(descriptor, 16);
            if (ownerOffset > 0 && ownerOffset + 8 <= descriptor.Length) {
                WbFolderPermAccount owner = ReadSidAccount(descriptor, ownerOffset, system);
                if (owner != null) item.Owner = owner.Name;
            }
            if ((control & 0x4) == 0 || daclOffset <= 0 || daclOffset + 8 > descriptor.Length) {
                item.NullDacl = true;
            } else {
                int count = BitConverter.ToUInt16(descriptor, daclOffset + 4);
                int pos = daclOffset + 8;
                // 並び順の決まり：このフォルダで設定した拒否 → このフォルダで設定した許可 → 継承した設定
                int stage = 0;
                for (int i = 0; i < count && pos + 4 <= descriptor.Length; i++) {
                    int type = descriptor[pos];
                    int flags = descriptor[pos + 1];
                    int size = BitConverter.ToUInt16(descriptor, pos + 2);
                    if (size < 4 || pos + size > descriptor.Length) break;
                    // 許可（0）・拒否（1）と、その条件付きのもの（9・10）。オブジェクトの種類を指定した設定などは、ファイルとフォルダでは使われない
                    if ((type == 0 || type == 1 || type == 9 || type == 10) && size >= 16) {
                        WbFolderPermAccount account = ReadSidAccount(descriptor, pos + 8, system);
                        if (account != null) {
                            WbFolderPermAce ace = new WbFolderPermAce();
                            ace.Deny = type == 1 || type == 10;
                            ace.Inherited = (flags & 0x10) != 0;
                            ace.Flags = flags;
                            ace.Mask = MapGeneric(BitConverter.ToUInt32(descriptor, pos + 4));
                            ace.Account = account;
                            string detail;
                            ace.Rights = NameRights(ace.Mask, flags, mode, out detail);
                            if (type == 9 || type == 10) detail = "条件付きの設定" + (detail.Length > 0 ? "。" + detail : "");
                            ace.Detail = detail;
                            ace.AppliesTo = NameAppliesTo(flags, mode);
                            if (mode != 2) {
                                if (account.Kind == KindUser) ace.Notes |= NoteUser;
                                if (account.Kind == KindUnknown) ace.Notes |= NoteUnknown;
                                if (ace.Deny) ace.Notes |= NoteDeny;
                                else if (account.IsWide && (ace.Mask & WriteBits) != 0) ace.Notes |= NoteWide;
                            }
                            if (ace.Inherited) { stage = 2; item.InheritedCount++; }
                            else {
                                item.ExplicitCount++;
                                if (stage == 2 || (ace.Deny && stage == 1)) item.Canonical = false;
                                else if (!ace.Deny) stage = 1;
                            }
                            list.Add(ace);
                        }
                    }
                    pos += size;
                }
            }
        }
        if (item.NullDacl) {
            WbFolderPermAce ace = new WbFolderPermAce();
            ace.Account = GetAccountBySid("S-1-1-0", system);
            ace.Mask = 0x1F01FF;
            ace.Rights = "フル コントロール";
            ace.Detail = mode == 2 ? "共有のアクセス許可の設定がないため、全員が操作できます" : "アクセス権の設定（DACL）がないため、全員が操作できます";
            ace.AppliesTo = mode == 0 ? "このフォルダ・サブフォルダ・ファイル" : (mode == 1 ? "このファイル" : "");
            if (mode != 2) ace.Notes = NoteWide;
            item.ExplicitCount++;
            list.Add(ace);
        }
        item.Aces = list.ToArray();
        int notes = 0;
        if (mode != 2) {
            if (item.Protected) notes |= NoteProtected;
            if (item.Stale) notes |= NoteStale;
            if (!item.Canonical) notes |= NoteOrder;
            // 継承した設定は、-Path のフォルダと、継承が親と合わないフォルダでだけ数える（そこが設定の出どころになるため）
            foreach (WbFolderPermAce ace in item.Aces) {
                if (!ace.Inherited || item.IsRoot || item.Stale) notes |= ace.Notes;
            }
        }
        item.Notes = notes;
    }

    // 親のアクセス権から、子（フォルダかファイル）が継承するはずの権限を作る（Windows の継承の決まり）。
    // 親の DACL がないときは、比べられないので null を返す
    public static WbFolderPermExpected ExpectedFor(byte[] parent, bool childIsDir) {
        if (parent == null || parent.Length < 20) return null;
        int control = BitConverter.ToUInt16(parent, 2);
        int daclOffset = BitConverter.ToInt32(parent, 16);
        if ((control & 0x4) == 0 || daclOffset <= 0 || daclOffset + 8 > parent.Length) return null;
        WbFolderPermExpected keys = new WbFolderPermExpected();
        int count = BitConverter.ToUInt16(parent, daclOffset + 4);
        int pos = daclOffset + 8;
        for (int i = 0; i < count && pos + 4 <= parent.Length; i++) {
            int type = parent[pos];
            int flags = parent[pos + 1];
            int size = BitConverter.ToUInt16(parent, pos + 2);
            if (size < 4 || pos + size > parent.Length) break;
            if ((type == 0 || type == 1 || type == 9 || type == 10) && size >= 16) {
                string sid = SidAt(parent, pos + 8);
                if (sid != null) AddExpected(keys, type, flags, BitConverter.ToUInt32(parent, pos + 4), sid, childIsDir);
            }
            pos += size;
        }
        return keys;
    }

    static void AddMask(Dictionary<string, uint> map, string key, uint mask) {
        uint current;
        map.TryGetValue(key, out current);
        map[key] = current | mask;
    }

    // 1 つの設定が子に渡る形。フォルダには CI の、ファイルには OI の付いた設定が効く（汎用のアクセス権は直したもの）。
    // CREATOR OWNER・CREATOR GROUP は、子の所有者（グループ）に効く。
    // 子のフォルダからさらに下へは、CI か OI の付いた設定が、元の権限のまま渡る（NP なら子で止まる）
    static void AddExpected(WbFolderPermExpected expected, int type, int flags, uint mask, string sid, bool childIsDir) {
        // ストアのアプリなどの SID（S-1-15-）は、アプリが作ったフォルダで Windows が特別な継承をするため、比べない
        if (sid.StartsWith("S-1-15-")) return;
        bool ci = (flags & 2) != 0, oi = (flags & 1) != 0, np = (flags & 4) != 0;
        uint norm = NormalizeMask(mask);
        if (childIsDir ? ci : oi) {
            if (sid == "S-1-3-0") AddMask(expected.Optional, type + "|@OWNER", norm);
            else if (sid == "S-1-3-1") AddMask(expected.Optional, type + "|@GROUP", norm);
            else AddMask(expected.Effective, type + "|" + sid, norm);
        }
        if (childIsDir && !np) {
            if (ci) AddMask(expected.Propagate, type + "|" + sid + "|C", norm);
            if (oi) AddMask(expected.Propagate, type + "|" + sid + "|O", norm);
        }
    }

    // 比べるときの権限。汎用のアクセス権（GENERIC_ALL など）とファイルのアクセス権のどちらの書き方にもなり、子に効く分は
    // ファイルのアクセス権の範囲（0x1F01FF）に切りそろえられるため、両方をそろえてから比べる
    static uint NormalizeMask(uint mask) {
        return MapGeneric(mask) & 0x001F01FF;
    }

    static bool SameMap(Dictionary<string, uint> a, Dictionary<string, uint> b) {
        if (a.Count != b.Count) return false;
        foreach (KeyValuePair<string, uint> pair in a) {
            uint value;
            if (!b.TryGetValue(pair.Key, out value) || value != pair.Value) return false;
        }
        return true;
    }

    // 子の継承した設定（継承の印の付いたもの）が、親から継承するはずの権限と、意味として同じか。
    // 子に効く権限（IO の付いていないもの）と、下へ渡す権限（CI か OI の付いたもの）を、それぞれ比べる。
    // CREATOR OWNER・CREATOR GROUP から子の所有者（グループ）に効く分は、あってもなくてもよい。
    // 継承を切っているときと、比べられないときは true
    public static bool InheritanceMatches(WbFolderPermExpected expected, byte[] child) {
        if (expected == null || child == null || child.Length < 20) return true;
        int control = BitConverter.ToUInt16(child, 2);
        int daclOffset = BitConverter.ToInt32(child, 16);
        if ((control & 0x1000) != 0 || (control & 0x4) == 0 || daclOffset <= 0 || daclOffset + 8 > child.Length) return true;
        Dictionary<string, uint> effective = new Dictionary<string, uint>(StringComparer.Ordinal);
        Dictionary<string, uint> propagate = new Dictionary<string, uint>(StringComparer.Ordinal);
        int count = BitConverter.ToUInt16(child, daclOffset + 4);
        int pos = daclOffset + 8;
        for (int i = 0; i < count && pos + 4 <= child.Length; i++) {
            int type = child[pos];
            int flags = child[pos + 1];
            int size = BitConverter.ToUInt16(child, pos + 2);
            if (size < 4 || pos + size > child.Length) break;
            if ((flags & 0x10) != 0 && (type == 0 || type == 1 || type == 9 || type == 10) && size >= 16) {
                string sid = SidAt(child, pos + 8);
                uint mask = NormalizeMask(BitConverter.ToUInt32(child, pos + 4));
                if (sid != null && !sid.StartsWith("S-1-15-")) {
                    if ((flags & 8) == 0) AddMask(effective, type + "|" + sid, mask);
                    // NP の付いた継承は、親からは来ないはずの形なので、別の渡し方として数え、食い違いとして見つける
                    string np = (flags & 4) != 0 ? "N" : "";
                    if ((flags & 2) != 0) AddMask(propagate, type + "|" + sid + "|C" + np, mask);
                    if ((flags & 1) != 0) AddMask(propagate, type + "|" + sid + "|O" + np, mask);
                }
            }
            pos += size;
        }
        if (!SameMap(propagate, expected.Propagate)) return false;
        if (SameMap(effective, expected.Effective)) return true;
        if (expected.Optional.Count == 0) return false;
        // CREATOR OWNER・CREATOR GROUP の分は、子を作った人（のグループ）の SID で作られる。所有者を後から変えても（所有者の置き換え・
        // takeown など）その設定は残るため、今の所有者だけでなく、子に効く設定にある SID のどれか 1 つに当てはめても試す。
        // それぞれ「なし」「今の所有者（グループ）」「子に効く設定にある SID」の組み合わせを試す（CREATOR OWNER の分は、SID を問わず 1 つまで許す）
        string owner = SidAt(child, BitConverter.ToInt32(child, 4)) ?? "";
        string group = SidAt(child, BitConverter.ToInt32(child, 8)) ?? "";
        List<KeyValuePair<string, uint>> optional = new List<KeyValuePair<string, uint>>(expected.Optional);
        if (optional.Count > 2) optional = optional.GetRange(0, 2);
        List<string>[] choices = new List<string>[optional.Count];
        for (int k = 0; k < optional.Count; k++) {
            string key = optional[k].Key;
            string prefix = key.Substring(0, key.IndexOf('|') + 1);
            choices[k] = new List<string>();
            choices[k].Add(null);
            choices[k].Add(prefix + (key.EndsWith("@OWNER") ? owner : group));
            foreach (string actualKey in effective.Keys) { if (actualKey.StartsWith(prefix)) choices[k].Add(actualKey); }
        }
        return TryOptional(effective, expected.Effective, optional, choices, 0);
    }

    static bool TryOptional(Dictionary<string, uint> actual, Dictionary<string, uint> current, List<KeyValuePair<string, uint>> optional, List<string>[] choices, int index) {
        if (index == optional.Count) return SameMap(actual, current);
        foreach (string choice in choices[index]) {
            Dictionary<string, uint> next = current;
            if (choice != null) {
                next = new Dictionary<string, uint>(current, StringComparer.Ordinal);
                AddMask(next, choice, optional[index].Value);
            }
            if (TryOptional(actual, next, optional, choices, index + 1)) return true;
        }
        return false;
    }

    // 古い形の ACL（自動継承の印がない）では、CREATOR OWNER・CREATOR GROUP から作られた「所有者（グループ）に効く設定」を、
    // Windows（GetNamedSecurityInfo）が継承として付け直さないことがある（親から継承した同じアカウントの同じ権限と重なるとき。
    // 管理者が作ったフォルダの「Administrators フル コントロール このフォルダだけ」など。エクスプローラーでも「継承元: なし」に見える）。
    // だれかが足した設定ではないので、このフォルダで設定したものがそれだけのときは、継承した設定として扱う（継承の印を付ける）。
    // ファイルでは、Windows が作るときに同じ設定を 1 つにまとめるため、この形は残らない（フォルダだけで使う）
    public static void MarkCreatorAces(WbFolderPermSecurity security, WbFolderPermExpected expected) {
        if (security == null || expected == null || security.Error != 0 || security.NullDacl || security.Protected || security.AutoInherited || !security.HasExplicit) return;
        uint ownerMask = 0, groupMask = 0;
        bool needOwner = expected.Optional.TryGetValue("0|@OWNER", out ownerMask);
        bool needGroup = expected.Optional.TryGetValue("0|@GROUP", out groupMask);
        if (!needOwner && !needGroup) return;
        byte[] d = security.Descriptor;
        int ownerOffset = BitConverter.ToInt32(d, 4);
        int groupOffset = BitConverter.ToInt32(d, 8);
        string owner = ownerOffset > 0 ? SidAt(d, ownerOffset) : null;
        string group = groupOffset > 0 ? SidAt(d, groupOffset) : null;
        int daclOffset = BitConverter.ToInt32(d, 16);
        int count = BitConverter.ToUInt16(d, daclOffset + 4);
        int pos = daclOffset + 8;
        List<int> marks = new List<int>();
        for (int i = 0; i < count && pos + 4 <= d.Length; i++) {
            int type = d[pos];
            int flags = d[pos + 1];
            int size = BitConverter.ToUInt16(d, pos + 2);
            if (size < 4 || pos + size > d.Length) break;
            if ((flags & 0x10) == 0) {
                // 許可で、このフォルダ（ファイル）だけに効き（OI・CI・NP・IO なし）、所有者（グループ）に、CREATOR OWNER（GROUP）と同じ権限
                string sid = (type == 0 && (flags & 0x0F) == 0 && size >= 16) ? SidAt(d, pos + 8) : null;
                uint norm = sid != null ? NormalizeMask(BitConverter.ToUInt32(d, pos + 4)) : 0;
                if (needOwner && sid != null && sid == owner && norm == ownerMask) { marks.Add(pos); needOwner = false; }
                else if (needGroup && sid != null && sid == group && norm == groupMask) { marks.Add(pos); needGroup = false; }
                else return;   // ほかにもこのフォルダで設定したものがある（そのまま「このフォルダ」の設定として出す）
            }
            pos += size;
        }
        if (marks.Count == 0) return;
        foreach (int at in marks) d[at + 1] = (byte)(d[at + 1] | 0x10);
        security.HasExplicit = false;
        security.HasInherited = true;
    }

    // ---- 出力（数が多いと 5.1 の関数呼び出しでは遅いため、行を組み立てる処理は C# で行う） ----

    static readonly System.Text.RegularExpressions.Regex ExcelRisk = new System.Text.RegularExpressions.Regex(@"^(?:[=+\-@]|[\d\s.,/:+\-]+$)");
    static readonly System.Globalization.CultureInfo Invariant = System.Globalization.CultureInfo.InvariantCulture;

    // Excel は、= + - @ で始まる値を数式に、数字と区切りの記号だけの値（001、1.5）や「.1」「1e5」を数値や日付に変えてしまうため、
    // その形の値は ="..." で囲む。式に書けない 255 文字を超える値は、先頭に空白を足す。全角の数字も数値になるため、半角にそろえて判定する
    public static string ProtectExcel(string text) {
        if (string.IsNullOrEmpty(text)) return text ?? "";
        bool risky = ExcelRisk.IsMatch(text);
        if (!risky) {
            string normalized = text;
            // 壊れた文字（対になっていないサロゲート）を含むと Normalize が例外を出すので、そのときはそのまま判定する
            try { normalized = text.Normalize(NormalizationForm.FormKC); } catch (Exception) { }
            double number;
            risky = double.TryParse(normalized.TrimEnd('%'), System.Globalization.NumberStyles.Float, Invariant, out number);
        }
        if (!risky) return text;
        if (text.Length <= 255) return "=\"" + text + "\"";
        return " " + text;
    }

    // 画面の表で列をそろえるため、全角文字を 2 文字分の幅として数える（半角カタカナは 1 文字分）
    public static int DisplayWidth(string text) {
        int width = 0;
        foreach (char ch in text ?? "") {
            int code = ch;
            if (code < 0x100 || (code >= 0xFF61 && code <= 0xFF9F)) width += 1;
            else if (code >= 0xDC00 && code <= 0xDFFF) { }   // サロゲートペアの後半は、前半で 2 文字分を数えた
            else width += 2;
        }
        return width;
    }

    public static string Pad(string text, int width, bool alignRight) {
        text = text ?? "";
        string pad = new string(' ', Math.Max(0, width - DisplayWidth(text)));
        return alignRight ? pad + text : text + pad;
    }

    public static string CsvLine(string[] fields) {
        StringBuilder line = new StringBuilder();
        for (int i = 0; i < fields.Length; i++) {
            if (i > 0) line.Append(',');
            line.Append('"').Append((fields[i] ?? "").Replace("\"", "\"\"")).Append('"');
        }
        return line.ToString();
    }

    static string Count(long value) { return value.ToString("N0", Invariant); }

    // folders.csv と permissions.csv の行を書く
    public static void WriteItemRows(System.IO.TextWriter itemWriter, System.IO.TextWriter aceWriter, WbFolderPermItem[] items) {
        foreach (WbFolderPermItem item in items) {
            string kind = item.IsFile ? "ファイル" : "フォルダ";
            string inherit = item.Protected ? "切っている" : "あり";
            string depth = Count(item.Depth);
            itemWriter.WriteLine(CsvLine(new string[] { item.Display, kind, depth, inherit, Count(item.ExplicitCount), Count(item.InheritedCount),
                ProtectExcel(item.Owner), ProtectExcel(item.ShareName), NoteText(item.Notes) }));
            foreach (WbFolderPermAce ace in item.Aces) {
                string place = ace.Inherited ? "親から継承" : (item.IsFile ? "このファイル" : "このフォルダ");
                aceWriter.WriteLine(CsvLine(new string[] { item.Display, kind, depth, inherit, ace.Deny ? "拒否" : "許可",
                    ProtectExcel(ace.Account.Name), ace.Account.Kind, ace.Rights, ace.AppliesTo, place, ace.Detail,
                    "0x" + ace.Mask.ToString("X8"), NoteText(ace.Notes), ace.Account.Sid }));
            }
        }
    }

    // 一覧の行（フォルダごとに、状態の行と、アクセス権の行）。-Path と共有のフォルダ、継承が親と合わないフォルダは、継承した分も出す。
    // 列は値に合わせてそろえるが、長いパスやアカウントが 1 つあっても全体がずれないよう、パスは 50、アカウントは 40 文字分までにする
    public static string[][] FormatItems(WbFolderPermItem[] items) {
        int kindWidth = 4, accountWidth = 10, rightsWidth = 8, appliesWidth = 4, pathWidth = 0;
        foreach (WbFolderPermItem item in items) {
            pathWidth = Math.Max(pathWidth, Math.Min(50, DisplayWidth(item.Display)));
            bool showAll = item.IsRoot || item.ShareName.Length > 0 || item.Stale;
            foreach (WbFolderPermAce ace in item.Aces) {
                if (ace.Inherited && !showAll) continue;
                kindWidth = Math.Max(kindWidth, DisplayWidth(AceKind(ace)));
                accountWidth = Math.Max(accountWidth, Math.Min(40, DisplayWidth(ace.Account.Name)));
                rightsWidth = Math.Max(rightsWidth, DisplayWidth(ace.Rights));
                appliesWidth = Math.Max(appliesWidth, DisplayWidth(ace.AppliesTo));
            }
        }
        string[][] result = new string[items.Length][];
        for (int i = 0; i < items.Length; i++) {
            WbFolderPermItem item = items[i];
            bool showAll = item.IsRoot || item.ShareName.Length > 0 || item.Stale;
            List<string> lines = new List<string>();
            lines.Add(" " + Pad(item.Display, pathWidth, false) + "  " + ItemStatus(item));
            foreach (WbFolderPermAce ace in item.Aces) {
                if (ace.Inherited && !showAll) continue;
                string text = "     " + Pad(AceKind(ace), kindWidth, false) + "  " + Pad(ace.Account.Name, accountWidth, false) + "  " +
                    Pad(ace.Rights, rightsWidth, false) + "  " + Pad(ace.AppliesTo, appliesWidth, false);
                if (ace.Notes != 0 && (!ace.Inherited || item.IsRoot || item.Stale)) text += "  ★" + NoteText(ace.Notes);
                lines.Add(text.TrimEnd());
            }
            // 空の DACL（だれにも許可していない）は、行がないと中身が分からないため、そう書く
            if (item.Aces.Length == 0) {
                lines.Add(("     " + Pad("", kindWidth, false) + "  " + "（だれにも許可していません）").TrimEnd());
            }
            result[i] = lines.ToArray();
        }
        return result;
    }

    static string AceKind(WbFolderPermAce ace) {
        return (ace.Deny ? "拒否" : "許可") + (ace.Inherited ? "（継承）" : "");
    }

    // 継承の状態（継承を切っている・継承が親と合わない・追加している・継承だけ）と、並び順・共有名
    public static string ItemStatus(WbFolderPermItem item) {
        List<string> states = new List<string>();
        if (item.NullDacl) states.Add("アクセス権の設定がない");
        else if (item.Protected) states.Add("継承を切っている");
        else {
            if (item.Stale) states.Add("継承が親と合わない");
            if (item.ExplicitCount > 0) states.Add("追加している");
        }
        if (!item.Canonical) states.Add("並び順が正しくない");
        string status = states.Count > 0 ? string.Join("・", states.ToArray()) : "継承だけ";
        if (item.ShareName.Length > 0) status += "（共有: " + item.ShareName + "）";
        return status;
    }

    // 「確かめたほうがよいところ」を数える。アクセス権の行を、設定した場所で数える（継承した設定は、-Path のフォルダと、
    // 継承が親と合わないフォルダの行でだけ数える）。個人と名前が見つからないものはアカウントを、広い範囲と拒否はアカウントとフォルダを例にする
    public static WbFolderPermFindings Summarize(WbFolderPermItem[] items) {
        WbFolderPermFindings findings = new WbFolderPermFindings();
        int[] bits = { NoteUser, NoteUnknown, NoteWide, NoteDeny };
        HashSet<string>[] seen = { new HashSet<string>(), new HashSet<string>(), new HashSet<string>(), new HashSet<string>() };
        foreach (WbFolderPermItem item in items) {
            if (item.Protected) { if (item.IsFile) findings.ProtectedFiles++; else findings.ProtectedFolders++; }
            if (item.Stale) findings.Stale.Add(item.Display);
            if (!item.Canonical) findings.Order.Add(item.Display);
            foreach (WbFolderPermAce ace in item.Aces) {
                if (ace.Account.Kind == KindLookupFailed) findings.LookupFailed = true;
                if (ace.Inherited && !(item.IsRoot || item.Stale)) continue;
                for (int k = 0; k < bits.Length; k++) {
                    if ((ace.Notes & bits[k]) == 0) continue;
                    findings.Counts[k]++;
                    string example = (k <= 1) ? ace.Account.Name : ace.Account.Name + ": " + item.Display;
                    if (seen[k].Add(example)) findings.Examples[k].Add(example);
                }
            }
        }
        return findings;
    }

    // アクセス権に出てくるグループと組み込みのアカウント（問い合わせる PC ごとに 1 回だけ。出てきた順）
    public static object[] CollectGroups(WbFolderPermItem[] items) {
        List<object> result = new List<object>();
        HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (WbFolderPermItem item in items) {
            foreach (WbFolderPermAce ace in item.Aces) {
                if (ace.Account.Kind != KindGroup && ace.Account.Kind != KindBuiltin) continue;
                if (seen.Add(item.LookupSystem + "|" + ace.Account.Sid)) result.Add(new object[] { ace.Account, item.LookupSystem });
            }
        }
        return result.ToArray();
    }

    static string SidAt(byte[] data, int offset) {
        if (offset <= 0 || offset + 8 > data.Length) return null;
        int sidLength = 8 + 4 * data[offset + 1];
        if (offset + sidLength > data.Length) return null;
        byte[] sid = new byte[sidLength];
        Buffer.BlockCopy(data, offset, sid, 0, sidLength);
        return SidToText(sid);
    }

    // 汎用のアクセス権（GENERIC_READ など。CREATOR OWNER などに付く）を、ファイルとフォルダのアクセス権に直す
    static uint MapGeneric(uint mask) {
        uint m = mask;
        if ((m & 0x80000000) != 0) m |= 0x120089;
        if ((m & 0x40000000) != 0) m |= 0x120116;
        if ((m & 0x20000000) != 0) m |= 0x1200A0;
        if ((m & 0x10000000) != 0) m |= 0x1F01FF;
        return m & 0x0FFFFFFF;
    }

    static readonly uint[] RightBits = { 0x20, 0x1, 0x80, 0x8, 0x2, 0x4, 0x100, 0x10, 0x40, 0x10000, 0x20000, 0x40000, 0x80000 };
    static readonly string[] RightNames = {
        "フォルダのスキャン/ファイルの実行", "フォルダの一覧/データの読み取り", "属性の読み取り", "拡張属性の読み取り",
        "ファイルの作成/データの書き込み", "フォルダの作成/データの追加", "属性の書き込み", "拡張属性の書き込み",
        "サブフォルダとファイルの削除", "削除", "アクセス許可の読み取り", "アクセス許可の変更", "所有権の取得" };

    // エクスプローラーの「セキュリティ」タブと同じ区切りの名前。どれにも当たらないものは「特殊」とし、中身を detail に並べる。
    // 同期（SYNCHRONIZE）は、付いていてもいなくても同じ名前にする
    public static string NameRights(uint mask, int flags, int mode, out string detail) {
        detail = "";
        uint core = mask & 0x000F01FF;
        uint other = mask & ~(uint)0x001F01FF;
        if (other == 0) {
            if (mode == 2) {
                if (core == 0x0F01FF) return "フル コントロール";
                if (core == 0x0301BF) return "変更";
                if (core == 0x0200A9) return "読み取り";
            } else {
                if (core == 0x0F01FF) return "フル コントロール";
                if (core == 0x0301BF) return "変更";
                if (core == 0x0200A9) return (mode == 0 && (flags & 3) == 2) ? "フォルダの内容の一覧表示" : "読み取りと実行";
                if (core == 0x020089) return "読み取り";
                if (core == 0x000116) return "書き込み";
                if (core == 0x02019F) return "読み取りと書き込み";
                if (core == 0x0201BF) return "読み取りと実行・書き込み";
            }
        }
        detail = DescribeRights(mask);
        return "特殊";
    }

    public static string DescribeRights(uint mask) {
        List<string> names = new List<string>();
        for (int i = 0; i < RightBits.Length; i++) { if ((mask & RightBits[i]) != 0) names.Add(RightNames[i]); }
        uint other = mask & ~(uint)0x001F01FF;
        if (other != 0) names.Add("その他（0x" + other.ToString("X8") + "）");
        if (names.Count == 0) return "（なし）";
        return string.Join("、", names.ToArray());
    }

    // 適用先（エクスプローラーの 7 種類）。CI = 2、OI = 1、NP（すぐ下だけ）= 4、IO（このフォルダには効かない）= 8
    public static string NameAppliesTo(int flags, int mode) {
        if (mode == 1) return "このファイル";
        if (mode == 2) return "";
        bool ci = (flags & 2) != 0, oi = (flags & 1) != 0, np = (flags & 4) != 0, io = (flags & 8) != 0;
        string text;
        if (io) {
            if (ci && oi) text = "サブフォルダ・ファイルだけ";
            else if (ci) text = "サブフォルダだけ";
            else if (oi) text = "ファイルだけ";
            else text = "（どこにも適用されない）";
        } else {
            if (ci && oi) text = "このフォルダ・サブフォルダ・ファイル";
            else if (ci) text = "このフォルダ・サブフォルダ";
            else if (oi) text = "このフォルダ・ファイル";
            else text = "このフォルダだけ";
        }
        if (np && (ci || oi)) text += "（すぐ下だけ）";
        return text;
    }

    static readonly string[] NoteNames = { "継承を切っている", "個人を直接指定", "名前が見つからない", "広い範囲に書き込みを許可", "拒否", "並び順が正しくない", "継承が親と合わない" };

    // 印の組み合わせを「継承を切っている・個人を直接指定」の形にする
    public static string NoteText(int notes) {
        List<string> names = new List<string>();
        for (int i = 0; i < NoteNames.Length; i++) { if ((notes & (1 << i)) != 0) names.Add(NoteNames[i]); }
        return string.Join("・", names.ToArray());
    }

    public static string NoteName(int bit) {
        for (int i = 0; i < NoteNames.Length; i++) { if (bit == (1 << i)) return NoteNames[i]; }
        return "";
    }

    static WbFolderPermAccount ReadSidAccount(byte[] data, int offset, string system) {
        if (offset + 8 > data.Length) return null;
        int subCount = data[offset + 1];
        int sidLength = 8 + 4 * subCount;
        if (offset + sidLength > data.Length) return null;
        byte[] sid = new byte[sidLength];
        Buffer.BlockCopy(data, offset, sid, 0, sidLength);
        return GetAccount(sid, SidToText(sid), system);
    }

    public static string SidToText(byte[] sid) {
        ulong authority = 0;
        for (int k = 0; k < 6; k++) authority = (authority << 8) | sid[2 + k];
        StringBuilder text = new StringBuilder("S-" + sid[0] + "-" + authority);
        int subCount = sid[1];
        for (int k = 0; k < subCount; k++) text.Append("-" + BitConverter.ToUInt32(sid, 8 + 4 * k));
        return text.ToString();
    }

    static byte[] TextToSid(string text) {
        string[] parts = text.Split('-');
        if (parts.Length < 3 || parts[0] != "S") return null;
        int subCount = parts.Length - 3;
        byte[] sid = new byte[8 + 4 * subCount];
        sid[0] = byte.Parse(parts[1]);
        sid[1] = (byte)subCount;
        ulong authority = ulong.Parse(parts[2]);
        for (int k = 5; k >= 0; k--) { sid[2 + k] = (byte)(authority & 0xFF); authority >>= 8; }
        for (int k = 0; k < subCount; k++) Buffer.BlockCopy(BitConverter.GetBytes(uint.Parse(parts[3 + k])), 0, sid, 8 + 4 * k, 4);
        return sid;
    }

    public static WbFolderPermAccount GetAccountBySid(string text, string system) {
        byte[] sid = TextToSid(text);
        if (sid == null) {
            WbFolderPermAccount bad = new WbFolderPermAccount();
            bad.Sid = text; bad.Name = text; bad.Kind = KindUnknown; bad.Error = 1337;
            return bad;
        }
        return GetAccount(sid, text, system);
    }

    static int Lookup(string system, byte[] sid, out string name, out string domain, out int use) {
        uint nameLength = 256, domainLength = 256;
        StringBuilder nameBuffer = new StringBuilder((int)nameLength), domainBuffer = new StringBuilder((int)domainLength);
        name = null; domain = null;
        if (!LookupAccountSidW(system, sid, nameBuffer, ref nameLength, domainBuffer, ref domainLength, out use)) {
            int error = Marshal.GetLastWin32Error();
            if (error != 122) return error;
            nameBuffer = new StringBuilder((int)nameLength + 1);
            domainBuffer = new StringBuilder((int)domainLength + 1);
            nameLength = (uint)nameBuffer.Capacity;
            domainLength = (uint)domainBuffer.Capacity;
            if (!LookupAccountSidW(system, sid, nameBuffer, ref nameLength, domainBuffer, ref domainLength, out use)) return Marshal.GetLastWin32Error();
        }
        name = nameBuffer.ToString();
        domain = domainBuffer.ToString();
        return 0;
    }

    // 広い範囲（ほぼ全員）：Everyone、ANONYMOUS LOGON、Authenticated Users、Users、Guests、Domain Users、Domain Guests
    static bool IsWideSid(string text) {
        if (text == "S-1-1-0" || text == "S-1-5-7" || text == "S-1-5-11" || text == "S-1-5-32-545" || text == "S-1-5-32-546") return true;
        if (text.StartsWith("S-1-5-21-") && (text.EndsWith("-513") || text.EndsWith("-514"))) return text.Split('-').Length == 8;
        return false;
    }

    // SID の名前と種類。system は問い合わせる PC（null はこの PC）。共有フォルダ越しのときは、サーバーのローカルのアカウントを
    // サーバーに問い合わせる。サーバーにつながらないときは、この PC で調べる
    static WbFolderPermAccount GetAccount(byte[] sid, string text, string system) {
        if (string.IsNullOrEmpty(system)) system = null;
        string key = (system ?? "") + "|" + text;
        WbFolderPermAccount account;
        if (accounts.TryGetValue(key, out account)) return account;
        account = new WbFolderPermAccount();
        account.Sid = text;
        string name, domain;
        int use;
        int error = Lookup(system, sid, out name, out domain, out use);
        if (error != 0 && error != 1332 && system != null) error = Lookup(null, sid, out name, out domain, out use);
        account.Error = error;
        if (error == 0) {
            account.Use = use;
            account.ShortName = name;
            account.Domain = domain;
            account.Name = string.IsNullOrEmpty(domain) ? name : domain + "\\" + name;
            switch (use) {
                case 1: account.Kind = KindUser; break;
                case 2: case 4: account.Kind = KindGroup; break;
                case 9: account.Kind = KindComputer; break;
                case 3: case 5: case 10: case 11: account.Kind = KindBuiltin; break;
                default: account.Kind = KindUnknown; break;
            }
        } else {
            account.Name = text;
            // ERROR_NONE_MAPPED（1332）は、名前が見つからない（削除されたアカウントなど）。
            // それ以外（ドメインコントローラーにつながらないなど）は、削除されたとは限らないので分ける
            account.Kind = error == 1332 ? KindUnknown : KindLookupFailed;
        }
        // S-1-15- はストアのアプリなどの権限。S-1-5-32- の後ろに 2 つ以上の数が続くものは、アプリの権限から作られるグループで、名前がないのが普通
        if (text.StartsWith("S-1-15-") || (text.StartsWith("S-1-5-32-") && text.Split('-').Length > 5)) account.Kind = KindApp;
        account.IsWide = IsWideSid(text);
        accounts[key] = account;
        return account;
    }

    // 共有フォルダの一覧（名前・種類・説明）。管理者でなくても読める
    public static WbFolderPermShare[] EnumShares(string server, out int error) {
        List<WbFolderPermShare> list = new List<WbFolderPermShare>();
        IntPtr buffer;
        int read, total, resume = 0;
        error = NetShareEnum(string.IsNullOrEmpty(server) ? null : server, 1, out buffer, -1, out read, out total, ref resume);
        try {
            if (error == 0 || error == 234) {
                int size = Marshal.SizeOf(typeof(ShareInfo1));
                for (int i = 0; i < read; i++) {
                    ShareInfo1 info = (ShareInfo1)Marshal.PtrToStructure(new IntPtr(buffer.ToInt64() + (long)i * size), typeof(ShareInfo1));
                    WbFolderPermShare share = new WbFolderPermShare();
                    share.Name = info.Name ?? "";
                    share.Type = info.Type;
                    share.Remark = info.Remark ?? "";
                    list.Add(share);
                }
                error = 0;
            }
        } finally {
            if (buffer != IntPtr.Zero) NetApiBufferFree(buffer);
        }
        return list.ToArray();
    }

    // 共有のフォルダ（サーバーの上のパス）と、共有のアクセス許可。アクセス許可を読めないときは Error に理由の番号を入れ、
    // フォルダだけでも読めるか試す
    public static WbFolderPermShareInfo GetShareInfo(string server, string name) {
        WbFolderPermShareInfo result = new WbFolderPermShareInfo();
        if (string.IsNullOrEmpty(server)) server = null;
        IntPtr buffer;
        int error = NetShareGetInfo(server, name, 502, out buffer);
        if (error == 0) {
            try {
                ShareInfo502 info = (ShareInfo502)Marshal.PtrToStructure(buffer, typeof(ShareInfo502));
                result.Path = info.Path ?? "";
                result.Remark = info.Remark ?? "";
                if (info.SecurityDescriptor != IntPtr.Zero) {
                    int length = (int)GetSecurityDescriptorLength(info.SecurityDescriptor);
                    int control = Marshal.ReadInt16(info.SecurityDescriptor, 2);
                    if ((control & 0x8000) != 0) {
                        byte[] descriptor = new byte[length];
                        Marshal.Copy(info.SecurityDescriptor, descriptor, 0, length);
                        result.Descriptor = descriptor;
                    } else {
                        uint relativeLength = 0;
                        MakeSelfRelativeSD(info.SecurityDescriptor, null, ref relativeLength);
                        byte[] descriptor = new byte[relativeLength];
                        if (MakeSelfRelativeSD(info.SecurityDescriptor, descriptor, ref relativeLength)) result.Descriptor = descriptor;
                        else result.Error = Marshal.GetLastWin32Error();
                    }
                }
            } finally {
                NetApiBufferFree(buffer);
            }
            return result;
        }
        result.Error = error;
        if (NetShareGetInfo(server, name, 2, out buffer) == 0) {
            try {
                ShareInfo2 info = (ShareInfo2)Marshal.PtrToStructure(buffer, typeof(ShareInfo2));
                result.Path = info.Path ?? "";
                result.Remark = info.Remark ?? "";
            } finally {
                NetApiBufferFree(buffer);
            }
        }
        return result;
    }

    // サーバーのコンピューター名（ローカルのアカウントのドメインの部分と比べるため）。読めないときは null
    public static string GetComputerName(string server) {
        IntPtr buffer;
        if (NetWkstaGetInfo(server, 100, out buffer) != 0) return null;
        try {
            WkstaInfo100 info = (WkstaInfo100)Marshal.PtrToStructure(buffer, typeof(WkstaInfo100));
            return info.ComputerName;
        } finally {
            NetApiBufferFree(buffer);
        }
    }

    // ドライブのファイルシステムの名前と、アクセス権を持てるか（FILE_PERSISTENT_ACLS）。読めないときは 0 以外のエラー番号を返す
    public static int GetVolumeAcl(string root, out string fileSystem, out bool hasAcl) {
        StringBuilder volumeName = new StringBuilder(261), fileSystemName = new StringBuilder(261);
        uint serial, maxComponent, flags;
        fileSystem = "";
        hasAcl = true;
        if (!GetVolumeInformationW(root, volumeName, volumeName.Capacity, out serial, out maxComponent, out flags, fileSystemName, fileSystemName.Capacity)) {
            return Marshal.GetLastWin32Error();
        }
        fileSystem = fileSystemName.ToString();
        hasAcl = (flags & 0x8) != 0;
        return 0;
    }

    // 実際の場所（\\?\C:\... か \\?\UNC\server\share\...）。SUBST・ネットワークドライブ・リンクの先をたどった後のパス
    public static string GetFinalPath(string path) {
        // FILE_READ_ATTRIBUTES、共有は読み取り・書き込み・削除、OPEN_EXISTING、FILE_FLAG_BACKUP_SEMANTICS（フォルダを開くため）
        using (SafeFileHandle handle = CreateFileW(path, 0x80, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero)) {
            if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
            StringBuilder buffer = new StringBuilder(512);
            uint length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 0);
            if (length >= buffer.Capacity) {
                buffer = new StringBuilder((int)length + 1);
                length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 0);
            }
            if (length == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
            return buffer.ToString();
        }
    }
}
'@

try {
    # 5.1 の Add-Type は、コンパイラーの警告もエラーとして扱う
    if (-not ('WbFolderPermNative' -as [type])) { Add-Type -TypeDefinition $NativeSource -IgnoreWarnings -WarningAction SilentlyContinue -ErrorAction Stop }
} catch {
    Stop-WithError ('Windows の API を呼ぶ準備ができません（Add-Type を使えません）: {0}' -f (Get-ErrorText (Get-InnerException $_.Exception)))
}

$NoteProtected = [WbFolderPermNative]::NoteProtected
$NoteUser      = [WbFolderPermNative]::NoteUser
$NoteUnknown   = [WbFolderPermNative]::NoteUnknown
$NoteWide      = [WbFolderPermNative]::NoteWide
$NoteDeny      = [WbFolderPermNative]::NoteDeny
$NoteOrder     = [WbFolderPermNative]::NoteOrder
$KindGroup     = [WbFolderPermNative]::KindGroup
$KindBuiltin   = [WbFolderPermNative]::KindBuiltin
$KindUser      = [WbFolderPermNative]::KindUser

#endregion

#region 調べる場所を決める

# 共有フォルダ。Server は表示用のサーバー名、System は名前を問い合わせる PC（この PC は $null）、Path はサーバーの上のフォルダ
$shares = [System.Collections.Generic.List[object]]::new()
$shareByKey = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
# ほかのサーバーの共有の、サーバーの上のパス（「サーバー|パス」）から共有名を引く表（中の共有のフォルダに名前を付けるため）
$serverShareByPath = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)

# 共有フォルダを 1 つ足す（同じものは 1 回だけ）。アクセス許可を読み、名前を付けておく
function Add-Share([string]$System, [string]$Name, [string]$Remark, $Info) {
    $key = '{0}|{1}' -f $System, $Name
    if ($shareByKey.ContainsKey($key)) { return }
    if ($null -eq $Info) { $Info = [WbFolderPermNative]::GetShareInfo($System, $Name) }
    if (-not $Remark) { $Remark = $Info.Remark }
    $server = if ($System) { $System } else { $machineName }
    $display = if ($System) { '\\{0}\{1}' -f $System, $Name } else { $Name }
    $item = $null
    $errorText = $null
    $noteText = $null
    if ($Info.Error -ne 0) {
        $errorText = Get-Win32Text $Info.Error
    } elseif ($null -ne $Info.Descriptor) {
        $item = [WbFolderPermItem]::new()
        $item.Display = $display
        $item.LookupSystem = $System
        [WbFolderPermNative]::Analyze($item, $Info.Descriptor, 2, $System)
    } else {
        # アクセス許可の設定が保存されていない共有（C$ などの管理用の共有）は、Windows が既定の設定を使う。
        # この PC の共有は、実際に使われる設定を Get-SmbShareAccess で読む。ほかのサーバーは、推測せずにそう書く
        $item = Get-DefaultShareItem $System $Name $display
        if ($null -eq $item) { $noteText = '設定なし（Windows の既定のアクセス許可が使われます）' }
    }
    $record = [pscustomobject]@{ Server = $server; System = $System; Name = $Name; Display = $display; Path = $Info.Path; Remark = $Remark; Item = $item; Error = $errorText; Note = $noteText }
    $shares.Add($record)
    $shareByKey[$key] = $record
    if ($System -and $Info.Path) {
        $pathKey = $System + '|' + $Info.Path.TrimEnd('\')
        $existing = $null
        if ($serverShareByPath.TryGetValue($pathKey, [ref]$existing)) { $serverShareByPath[$pathKey] = $existing + '、' + $Name }
        else { $serverShareByPath[$pathKey] = $Name }
    }
}

# 足した共有フォルダの記録（なければ $null）
function Get-ShareRecord([string]$System, [string]$Name) {
    $record = $null
    if ($shareByKey.TryGetValue(('{0}|{1}' -f $System, $Name), [ref]$record)) { return $record }
    return $null
}

# この PC の共有で、アクセス許可の設定が保存されていないとき、Windows が使う既定の設定を Get-SmbShareAccess で読む。読めなければ $null
$ShareRights = @{ 'Full' = @('フル コントロール', 0x1F01FF); 'Change' = @('変更', 0x1301BF); 'Read' = @('読み取り', 0x1200A9) }
function Get-DefaultShareItem([string]$System, [string]$Name, [string]$Display) {
    # この PC を \\127.0.0.1 や \\localhost で指したときも、この PC の共有として読む
    if ($System -and @('127.0.0.1', '::1', 'localhost', '.', $machineName) -notcontains $System) { return $null }
    try { $entries = @(Get-SmbShareAccess -Name $Name -ErrorAction Stop) } catch { return $null }
    if ($entries.Count -eq 0) { return $null }
    $aceList = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $entries) {
        $ace = [WbFolderPermAce]::new()
        $ace.Deny = [string]$entry.AccessControlType -eq 'Deny'
        $accountName = [string]$entry.AccountName
        $sid = $null
        try { $sid = ([Security.Principal.NTAccount]::new($accountName)).Translate([Security.Principal.SecurityIdentifier]).Value } catch { }
        if ($sid) {
            $ace.Account = [WbFolderPermNative]::GetAccountBySid($sid, $null)
        } else {
            $account = [WbFolderPermAccount]::new()
            $account.Name = $accountName
            $account.Kind = [WbFolderPermNative]::KindUnknown
            $ace.Account = $account
        }
        $right = [string]$entry.AccessRight
        if ($ShareRights.ContainsKey($right)) {
            $ace.Rights = $ShareRights[$right][0]
            $ace.Mask = [uint32]$ShareRights[$right][1]
        } else {
            $ace.Rights = '特殊'
        }
        $ace.Detail = '設定が保存されていないため、Windows の既定のアクセス許可を表示しています'
        $aceList.Add($ace)
    }
    $item = [WbFolderPermItem]::new()
    $item.Display = $Display
    $item.LookupSystem = $System
    $item.Aces = [WbFolderPermAce[]]$aceList.ToArray()
    $item.ExplicitCount = $aceList.Count
    return $item
}

# 「すべての共有」に含める共有か。ディスクの共有（種類の下位 8 ビットが 0）で、管理用の印（STYPE_SPECIAL、0x80000000）がなく、
# Windows が自動で作るもの（print$）でもないもの。0x80000000 は 5.1 では Int32 の負の数になるため、[long] で比べる
function Test-UserShare($Share) {
    $type = [long]$Share.Type
    return ($type -band 0xFF) -eq 0 -and ($type -band 2147483648L) -eq 0 -and -not ($AutoShares -contains $Share.Name)
}

# この PC の共有フォルダ（管理用の共有と、プリンターなどディスク以外の共有は除く）。フォルダは、アクセス許可を読めないときも
# 一覧から分かるよう、Win32_Share でも調べる
$localShareCache = $null
function Get-LocalShareList {
    if ($null -ne $script:localShareCache) { return , $script:localShareCache }
    $list = [System.Collections.Generic.List[object]]::new()
    $enumError = 0
    $found = [WbFolderPermNative]::EnumShares($null, [ref]$enumError)
    $cimPaths = $null
    foreach ($share in $found) {
        if (-not (Test-UserShare $share)) { continue }
        $info = [WbFolderPermNative]::GetShareInfo($null, $share.Name)
        if (-not $info.Path) {
            if ($null -eq $cimPaths) {
                $cimPaths = @{}
                try { foreach ($s in @(Get-CimInstance -ClassName Win32_Share -ErrorAction Stop)) { $cimPaths[[string]$s.Name] = [string]$s.Path } } catch { }
            }
            if ($cimPaths.ContainsKey($share.Name)) { $info.Path = $cimPaths[$share.Name] }
        }
        $list.Add([pscustomobject]@{ Name = $share.Name; Remark = $share.Remark; Info = $info; Real = $info.Path.TrimEnd('\') })
    }
    $script:localShareCache = $list
    return , $list
}

# 調べるフォルダ。Full は表示用のパス、Long は \\?\ 付きのパス、Real は実際の場所（SUBST・ネットワークドライブを通したときに、
# この PC の共有フォルダと突き合わせるため）、System は名前を問い合わせる PC、ShareName はこのフォルダそのものが共有のときの名前
$targets = [System.Collections.Generic.List[object]]::new()
$targetTexts = [System.Collections.Generic.List[string]]::new()
$localTargets = $false

# ServerPath は、ほかのサーバーのフォルダのときの、サーバーの上のパス（共有のフォルダを読めたときだけ。重なりの判定と共有名に使う）
function New-Target([string]$Full, [string]$Real, [string]$System, [string]$ShareName) {
    return [pscustomobject]@{ Full = $Full; Long = (ConvertTo-LongPath $Full); Real = $Real; System = $System; ShareName = $ShareName; Strict = $true; ServerPath = '' }
}

# アクセス権のないファイルシステム（FAT32・exFAT）なら止める
function Assert-AclVolume([string]$Real, [string]$Shown) {
    $root = if ($Real -match '^(\\\\[^\\]+\\[^\\]+)') { $Matches[1] + '\' } elseif ($Real -match '^([A-Za-z]:)') { $Matches[1] + '\' } else { $null }
    if (-not $root) { return }
    $fileSystem = ''
    $hasAcl = $true
    if ([WbFolderPermNative]::GetVolumeAcl($root, [ref]$fileSystem, [ref]$hasAcl) -ne 0) { return }
    if (-not $hasAcl) { Stop-WithError ('このドライブ（{0}）には、アクセス権がありません: {1}' -f $fileSystem, $Shown) }
}

if ($rootTexts.Count -eq 0) {
    # -Path を省略したときは、この PC の共有フォルダすべて。フォルダが重なるときは、外側だけを調べる
    $localTargets = $true
    $localList = Get-LocalShareList
    foreach ($share in $localList) { Add-Share $null $share.Name $share.Remark $share.Info }
    $withPath = @(foreach ($share in $localList) { if ($share.Real) { $share } })
    if ($withPath.Count -eq 0) { Stop-WithError 'この PC には共有フォルダがありません。調べるフォルダを -Path で指定してください' }
    foreach ($share in $withPath) {
        $full = $share.Real
        if ($full -match '^[A-Za-z]:$') { $full += '\' }
        $target = New-Target $full $share.Real $null $share.Name
        $target.Strict = $false
        $targets.Add($target)
    }
    $targetTexts.Add(('この PC の共有フォルダ {0} 件' -f $localList.Count))
} else {
    foreach ($text in $rootTexts) {
        # \\サーバー名 だけのときは、そのサーバーの共有フォルダすべて
        if ($text -match '^\\\\([^\\?]+)\\?$') {
            $server = $Matches[1]
            $enumError = 0
            $found = [WbFolderPermNative]::EnumShares($server, [ref]$enumError)
            if ($enumError -ne 0) {
                Stop-WithError ((@(('サーバーの共有フォルダの一覧を読めません: \\{0}（{1}）' -f $server, (Get-Win32Text $enumError))) + (Get-NetworkHint ('\\' + $server))) -join "`r`n")
            }
            $count = 0
            foreach ($share in $found) {
                if (-not (Test-UserShare $share)) { continue }
                $count++
                Add-Share $server $share.Name $share.Remark $null
                $target = New-Target ('\\{0}\{1}' -f $server, $share.Name) ('\\{0}\{1}' -f $server, $share.Name) $server $share.Name
                $target.Strict = $false
                $target.ServerPath = [string](Get-ShareRecord $server $share.Name).Path
                $targets.Add($target)
            }
            if ($count -eq 0) {
                Stop-WithError ('\\{0} には共有フォルダがありません（C$ や ADMIN$ などの管理用の共有は含めません。調べるときは \\{0}\C$\フォルダ のように指定してください）' -f $server)
            }
            $targetTexts.Add(('\\{0}（共有フォルダ {1} 件）' -f $server, $count))
            continue
        }
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
        # フォルダを中まで読めるかを確かめる（共有フォルダは、接続していないと読めない）
        try {
            [void][IO.Directory]::EnumerateFileSystemEntries($rootLong).GetEnumerator().MoveNext()
        } catch {
            $reason = Get-ErrorText (Get-InnerException $_.Exception)
            Stop-WithError ((@(('フォルダを開けません: {0}（{1}）' -f $rootFull, $reason)) + (Get-NetworkHint $rootFull)) -join "`r`n")
        }
        # SUBST やネットワークドライブを通したときは、実際の場所で、共有フォルダとサーバーを決める
        $real = $rootFull
        try { $real = ConvertTo-DisplayPath ([WbFolderPermNative]::GetFinalPath($rootLong)).TrimEnd('\') } catch { }
        if ($real -match '^[A-Za-z]:$') { $real += '\' }
        Assert-AclVolume $real $rootFull
        if ($real -match '^\\\\([^\\]+)\\([^\\]+)(\\.*)?$') {
            # 共有フォルダ（ネットワークドライブを含む）。名前は、サーバーに問い合わせる
            $server = $Matches[1]
            $shareName = $Matches[2]
            $restPath = [string]$Matches[3]
            Add-Share $server $shareName '' $null
            $target = New-Target $rootFull $real $server $(if ($restPath) { '' } else { $shareName })
            $sharePath = [string](Get-ShareRecord $server $shareName).Path
            if ($sharePath) { $target.ServerPath = $sharePath.TrimEnd('\') + $restPath }
            $targets.Add($target)
        } else {
            $localTargets = $true
            $targets.Add((New-Target $rootFull $real $null ''))
        }
        $targetTexts.Add($rootFull)
    }
}

# この PC のフォルダを調べるときは、そのフォルダを含む共有と、その中にある共有も出す（実際の場所で比べる）
$localShareByReal = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
if ($localTargets) {
    $localList = Get-LocalShareList
    foreach ($share in $localList) {
        if (-not $share.Real) { continue }
        $sharePrefix = $share.Real.TrimEnd('\') + '\'
        foreach ($target in $targets) {
            if ($target.System) { continue }
            $targetPrefix = $target.Real.TrimEnd('\') + '\'
            if ($sharePrefix.StartsWith($targetPrefix, [StringComparison]::OrdinalIgnoreCase) -or $targetPrefix.StartsWith($sharePrefix, [StringComparison]::OrdinalIgnoreCase)) {
                Add-Share $null $share.Name $share.Remark $share.Info
                break
            }
        }
        $key = $share.Real.TrimEnd('\')
        $existing = $null
        if ($localShareByReal.TryGetValue($key, [ref]$existing)) { $localShareByReal[$key] = $existing + '、' + $share.Name }
        else { $localShareByReal[$key] = $share.Name }
    }
}

# 同じフォルダや、片方の中にあるフォルダは、外側だけを調べる（同じフォルダを二重に数えないため）。
# ほかのサーバーの共有は、サーバーの上のパスが分かれば、そのパスで比べる（\\fs01\share と \\fs01\営業 が D:\share と D:\share\営業 のとき）。
# この PC のフォルダは、実際の場所で比べる
function Get-OverlapKey($Target) {
    if ($Target.System -and $Target.ServerPath) { return $Target.System + '|' + $Target.ServerPath.TrimEnd('\') }
    if ($Target.System) { return $Target.Full.TrimEnd('\') }
    return $Target.Real.TrimEnd('\')
}
$targetArray = $targets.ToArray()
$targetKeys = [string[]]@(foreach ($target in $targetArray) { Get-PathSortText (Get-OverlapKey $target) })
$targetArray = Get-SortedList $targetKeys $targetArray
$targets = [System.Collections.Generic.List[object]]::new()
$keptKeys = [System.Collections.Generic.List[string]]::new()
# 外側に含まれるとして外した内側のフォルダ。外側の途中が読めないと届かないため、届かなかったものは後で別に調べる
$droppedTargets = [System.Collections.Generic.List[object]]::new()
$pendingKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($target in $targetArray) {
    $key = Get-OverlapKey $target
    $prefix = $key + '\'
    $inside = $false
    foreach ($keptKey in $keptKeys) {
        if ($prefix.StartsWith($keptKey + '\', [StringComparison]::OrdinalIgnoreCase)) { $inside = $true; break }
    }
    if (-not $inside) {
        $targets.Add($target)
        $keptKeys.Add($key)
    } elseif ($pendingKeys.Add($key)) {
        $droppedTargets.Add([pscustomobject]@{ Key = $key; Target = $target })
    }
}

# 調べているフォルダの、重なりの判定と同じ形の鍵（外した内側のフォルダに着いたかを確かめるため）
function Get-FolderOverlapKey($Target, [string]$Display) {
    $rest = $Display.Substring($Target.Full.TrimEnd('\').Length)
    if ($Target.System -and $Target.ServerPath) { return ($Target.System + '|' + $Target.ServerPath.TrimEnd('\') + $rest).TrimEnd('\') }
    if ($Target.System) { return $Display.TrimEnd('\') }
    return ($Target.Real.TrimEnd('\') + $rest).TrimEnd('\')
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
        $saveWarning = '保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, $reason
        Write-Warning $saveWarning
    }
}

#endregion

#region 判定の部品

$skipped   = [System.Collections.Generic.List[object]]::new()
$skipKinds = [ordered]@{ '権限なし' = 0; 'その他のエラー' = 0 }
$excludedCount  = 0
$linkCount      = 0
$folderCount    = 0
$fileCount      = 0
$changedFolders = 0
$changedFiles   = 0
$items          = [System.Collections.Generic.List[object]]::new()   # 一覧に出すフォルダとファイル
$unsureRoots    = [System.Collections.Generic.List[string]]::new()   # 継承した設定を見分けられなかった共有のルート
$sw             = [Diagnostics.Stopwatch]::StartNew()
$nextProgress   = 0L

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

# 読めなかった理由（skipped.csv の理由と詳細）
function Get-ReadFailure([int]$Code, [string]$What) {
    if ($Code -eq 5) { return @('権限なし', ('{0}（アクセスが拒否されました）' -f $What)) }
    return @('その他のエラー', ('{0}（{1}）' -f $What, (Get-Win32Text $Code)))
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

# フォルダ（ファイル）のアクセス権を読み、一覧に出すものなら記録を作って足す
function Add-Item([string]$Display, $Security, [bool]$IsFile, [int]$Level, [bool]$IsRoot, [string]$ShareName, [string]$System) {
    $item = [WbFolderPermItem]::new()
    $item.Display = $Display
    $item.IsFile = $IsFile
    $item.Depth = $Level
    $item.IsRoot = $IsRoot
    $item.ShareName = $ShareName
    $item.Stale = $Security.Stale
    $item.LookupSystem = $System
    [WbFolderPermNative]::Analyze($item, $Security.Descriptor, $(if ($IsFile) { 1 } else { 0 }), $System)
    $items.Add($item)
}

# 調べているフォルダが共有フォルダなら、その名前を返す。この PC は実際の場所で、ほかのサーバーはサーバーの上のパスで比べる
function Get-ShareNameFor($Target, [string]$Display) {
    $rest = $Display.Substring($Target.Full.TrimEnd('\').Length)
    $name = $null
    if ($Target.System) {
        if (-not $Target.ServerPath -or $serverShareByPath.Count -eq 0) { return '' }
        $key = $Target.System + '|' + ($Target.ServerPath.TrimEnd('\') + $rest).TrimEnd('\')
        if ($serverShareByPath.TryGetValue($key, [ref]$name)) { return $name }
        return ''
    }
    if ($localShareByReal.Count -eq 0) { return '' }
    $real = $Target.Real.TrimEnd('\') + $rest
    if ($localShareByReal.TryGetValue($real.TrimEnd('\'), [ref]$name)) { return $name }
    return ''
}

# -Path のフォルダが継承するはずの設定（親のフォルダのアクセス権から作る）。ドライブのルートと共有のルートには親がないので $null
function Get-RootExpected($Target) {
    $full = $Target.Full.TrimEnd('\')
    if ($full -match '^[A-Za-z]:$' -or $full -match '^\\\\[^\\]+\\[^\\]+$') { return $null }
    $cut = $full.LastIndexOf('\')
    if ($cut -le 0) { return $null }
    $parent = $full.Substring(0, $cut)
    if ($parent -match '^[A-Za-z]:$') { $parent += '\' }
    $parentSecurity = [WbFolderPermNative]::ReadSecurity((ConvertTo-LongPath $parent))
    if ($parentSecurity.Error -ne 0) { return $null }
    return [WbFolderPermNative]::ExpectedFor($parentSecurity.Descriptor, $true)
}

# 共有のルート（\\fs01\営業。ネットワークドライブのルートを含む）を共有越しに読むと、Windows は親のフォルダをたどれないため、
# 継承の印のない古い形の ACL（フォーマットしたままのドライブの直下に作ったフォルダなど）では、親から継承した設定も
# 「このフォルダで設定した」として返す。そのときは、サーバーの管理共有越し（\\fs01\D$\share\営業）に読み直す。
# 戻り値は、読み直したパス（読み直さなくてよいときは $null、読み直せないときは空）
function Get-ShareRootRereadPath($Target) {
    if (-not $Target.System -or -not [regex]::IsMatch($Target.Real, '^\\\\[^\\]+\\[^\\]+\\?$')) { return $null }
    # サーバーの上の場所は、サーバーの管理者でないと分からないことがある
    $match = [regex]::Match([string]$Target.ServerPath, '^([A-Za-z]):\\?(.*)$')
    if (-not $match.Success) { return '' }
    # ドライブのルートには親がないので、読み直さなくてよい
    if (-not $match.Groups[2].Value.TrimEnd('\')) { return $null }
    return ConvertTo-LongPath ('\\{0}\{1}$\{2}' -f $Target.System, $match.Groups[1].Value.ToUpperInvariant(), $match.Groups[2].Value.TrimEnd('\'))
}

# 管理共有越しに読み直した共有のルートの、親のフォルダから継承するはずの設定（CREATOR OWNER から作られた設定を見分けるため）。
# 読めないときは $null
function Get-RereadParentExpected([string]$RereadPath) {
    $cut = $RereadPath.TrimEnd('\').LastIndexOf('\')
    if ($cut -le 0) { return $null }
    $parent = $RereadPath.Substring(0, $cut)
    # 管理共有のルート（\\?\UNC\fs01\D$）は、最後に \ を付ける
    if ($parent -match '^\\\\\?\\UNC\\[^\\]+\\[^\\]+$') { $parent += '\' }
    $parentSecurity = [WbFolderPermNative]::ReadSecurity($parent)
    if ($parentSecurity.Error -ne 0) { return $null }
    return [WbFolderPermNative]::ExpectedFor($parentSecurity.Descriptor, $true)
}

#endregion

$finished   = $false
$unexpected = $false
$saveError  = $null
$lines      = [System.Collections.Generic.List[string]]::new()   # summary.txt に書く行
$sortedItems = @()
$groupRows   = [System.Collections.Generic.List[object]]::new()

function Save-Result {
    if (-not $runDir) { return }
    try {
        # 保存先は \\?\ 付きで作っているため、長いパスでも書けるよう、ここも \\?\ を付ける
        [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
        $shareRows = [System.Collections.Generic.List[object]]::new()
        foreach ($share in $shares) {
            if ($null -eq $share.Item) {
                $note = if ($share.Error) { '共有のアクセス許可を読めません（{0}）' -f $share.Error } else { $share.Note }
                $shareRows.Add([object[]]@((Protect-ExcelText $share.Server), (Protect-ExcelText $share.Name), $share.Path, (Protect-ExcelText $share.Remark), '', '', '', '', '', '', $note))
                continue
            }
            if ($share.Item.Aces.Count -eq 0) {
                $shareRows.Add([object[]]@((Protect-ExcelText $share.Server), (Protect-ExcelText $share.Name), $share.Path, (Protect-ExcelText $share.Remark), '', '', '', '', '', '', 'だれにも許可していません'))
                continue
            }
            foreach ($ace in $share.Item.Aces) {
                $note = $ace.Detail
                $shareRows.Add([object[]]@((Protect-ExcelText $share.Server), (Protect-ExcelText $share.Name), $share.Path, (Protect-ExcelText $share.Remark),
                    $(if ($ace.Deny) { '拒否' } else { '許可' }), (Protect-ExcelText $ace.Account.Name), $ace.Account.Kind, $ace.Rights,
                    ('0x' + $ace.Mask.ToString('X8')), $ace.Account.Sid, $note))
            }
        }
        Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'shares.csv')) $ShareColumns $shareRows
        # フォルダとアクセス権の行は数が多くなるため、C# で書く
        $itemWriter = [IO.StreamWriter]::new((ConvertTo-LongPath (Join-Path $runDir 'folders.csv')), $false, $Utf8Bom)
        try {
            $aceWriter = [IO.StreamWriter]::new((ConvertTo-LongPath (Join-Path $runDir 'permissions.csv')), $false, $Utf8Bom)
            try {
                $itemWriter.WriteLine((ConvertTo-CsvLine $ItemColumns))
                $aceWriter.WriteLine((ConvertTo-CsvLine $AceColumns))
                [WbFolderPermNative]::WriteItemRows($itemWriter, $aceWriter, [WbFolderPermItem[]]$sortedItems)
            } finally {
                $aceWriter.Dispose()
            }
        } finally {
            $itemWriter.Dispose()
        }
        if ($ExpandGroup) { Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'groups.csv')) $GroupColumns $groupRows }
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

# summary.txt にだけ足す（画面に出す行数を絞るとき）
function Add-SummaryLine([string]$Text, [bool]$Show) {
    $lines.Add($Text)
    if ($Show) { Write-Host $Text }
}

#region グループのメンバー

$domainJoined = $false
# ローカルのアカウント（FS01\Administrator など）で実行しているか。ドメインに参加している PC でも、ドメインコントローラーに問い合わせられない
$localAccountRun = $false
$computerNames = @{}
$disabledCache = @{}
$groupFailures = [System.Collections.Generic.List[string]]::new()

# 名前を問い合わせる PC のコンピューター名（ローカルのアカウントのドメインの部分と比べるため）
function Get-ComputerNameOf([string]$System) {
    if (-not $System) { return $machineName }
    if ($computerNames.ContainsKey($System)) { return $computerNames[$System] }
    $name = [WbFolderPermNative]::GetComputerName($System)
    if (-not $name) { $name = $System }
    $computerNames[$System] = $name
    return $name
}

# ドメインのユーザーが無効か（読めないときは $false）
function Test-DomainDisabled([string]$Sid) {
    if (-not $domainJoined) { return $false }
    if ($disabledCache.ContainsKey($Sid)) { return $disabledCache[$Sid] }
    $disabled = $false
    try {
        $entry = [ADSI]('LDAP://<SID={0}>' -f $Sid)
        $value = $entry.Properties['userAccountControl']
        if ($value.Count -gt 0) { $disabled = ([int]$value[0] -band 2) -ne 0 }
    } catch { }
    $disabledCache[$Sid] = $disabled
    return $disabled
}

# ローカルのグループのメンバー（WinNT プロバイダー。Get-LocalGroupMember は、名前を解決できない SID があると失敗する）
function Get-WinNTMember($Account, [string]$System, [string]$Computer) {
    $machine = if ($System) { $System } else { '.' }
    $group = [ADSI]('WinNT://{0}/{1},group' -f $machine, $Account.ShortName)
    $result = [System.Collections.Generic.List[object]]::new()
    foreach ($member in @($group.Invoke('Members'))) {
        $type = $member.GetType()
        $sidBytes = [byte[]]$type.InvokeMember('objectSid', 'GetProperty', $null, $member, $null)
        $memberAccount = [WbFolderPermNative]::GetAccountBySid([WbFolderPermNative]::SidToText($sidBytes), $System)
        $disabled = $false
        if ($memberAccount.Kind -eq $KindUser) {
            if ([string]::Equals($memberAccount.Domain, $Computer, [StringComparison]::OrdinalIgnoreCase)) {
                # メンバーの一覧から得たオブジェクトは UserFlags を持たないため、ユーザーの AdsPath で開き直して読む
                try {
                    $adsPath = [string]$type.InvokeMember('AdsPath', 'GetProperty', $null, $member, $null)
                    $userFlags = ([ADSI]$adsPath).UserFlags
                    if ($userFlags.Count -gt 0) { $disabled = ([int]$userFlags.Value -band 2) -ne 0 }
                } catch { }
            } else {
                $disabled = Test-DomainDisabled $memberAccount.Sid
            }
        }
        $result.Add([pscustomobject]@{ Account = $memberAccount; Disabled = $disabled })
    }
    return , $result
}

# ドメインのグループのメンバー（LDAP の属性スコープ検索で、member の先をまとめて読む）
function Get-LdapMember($Account, [string]$System) {
    $entry = [ADSI]('LDAP://<SID={0}>' -f $Account.Sid)
    $searcher = [DirectoryServices.DirectorySearcher]::new($entry)
    $searcher.SearchScope = [DirectoryServices.SearchScope]::Base
    $searcher.AttributeScopeQuery = 'member'
    $searcher.PageSize = 1000
    [void]$searcher.PropertiesToLoad.AddRange([string[]]@('objectSid', 'userAccountControl'))
    $result = [System.Collections.Generic.List[object]]::new()
    $found = $searcher.FindAll()
    try {
        foreach ($entryResult in $found) {
            $sidValues = $entryResult.Properties['objectsid']
            if ($sidValues.Count -eq 0) { continue }
            $memberAccount = [WbFolderPermNative]::GetAccountBySid([WbFolderPermNative]::SidToText([byte[]]$sidValues[0]), $System)
            $disabled = $false
            $uac = $entryResult.Properties['useraccountcontrol']
            if ($memberAccount.Kind -eq $KindUser -and $uac.Count -gt 0) { $disabled = ([int]$uac[0] -band 2) -ne 0 }
            $result.Add([pscustomobject]@{ Account = $memberAccount; Disabled = $disabled })
        }
    } finally {
        $found.Dispose()
        $searcher.Dispose()
    }
    return , $result
}

# グループの直接のメンバー。ローカルのグループ（BUILTIN と、その PC のグループ）は WinNT で、ドメインのグループは LDAP で読む
function Get-DirectMember($Account, [string]$System) {
    $computer = Get-ComputerNameOf $System
    $isLocal = [string]::Equals($Account.Domain, 'BUILTIN', [StringComparison]::OrdinalIgnoreCase) -or
        [string]::Equals($Account.Domain, $computer, [StringComparison]::OrdinalIgnoreCase)
    if ($isLocal) { return Get-WinNTMember $Account $System $computer }
    # ドメインに参加していない PC からは、ドメインコントローラーに問い合わせられない（ローカルのグループとして探すと、分かりにくいエラーになる）
    if (-not $domainJoined) { throw 'この PC はドメインに参加していないため、ドメインのグループのメンバーを調べられません。ドメインに参加している PC で実行してください' }
    # ローカルのアカウントのままだと、Windows のエラーは「指定されたドメインがないか、またはアクセスできません。」になり、理由が分からない
    if ($localAccountRun) { throw ('ローカルのアカウント（{0}\{1}）で実行しているため、ドメインのグループのメンバーを調べられません。ドメインのアカウントで実行してください' -f [Environment]::UserDomainName, [Environment]::UserName) }
    return Get-LdapMember $Account $System
}

function Add-GroupRow([string]$Top, [string]$Member, [string]$Kind, [string]$Via, [string]$State, [string]$Sid, [string]$Note) {
    $groupRows.Add([object[]]@((Protect-ExcelText $Top), (Protect-ExcelText $Member), $Kind, (Protect-ExcelText $Via), $State, $Sid, $Note))
}

# グループのメンバーを、入れ子のグループもたどって足す。Via は、たどってきた間のグループ
function Add-GroupMember($Top, $Group, [string]$System, [string]$Via, $Visited) {
    try {
        $members = Get-DirectMember $Group $System
    } catch {
        $reason = Get-ErrorText (Get-InnerException $_.Exception)
        Add-GroupRow $Top.Name '' '' $Via '' $Group.Sid ('{0} のメンバーを調べられませんでした（{1}）' -f $Group.Name, $reason)
        $groupFailures.Add(('{0}（{1}）' -f $Group.Name, $reason))
        return
    }
    if ($members.Count -eq 0 -and -not $Via) {
        Add-GroupRow $Top.Name '' '' '' '' $Group.Sid 'メンバーはいません'
        return
    }
    foreach ($member in $members) {
        $account = $member.Account
        $note = ''
        $recurse = $false
        if ($account.Kind -eq $KindGroup) {
            if ($account.IsWide) { $note = 'ほぼ全員が入っているグループのため、メンバーを調べません' }
            elseif ($Visited.Contains($account.Sid)) { $note = 'ほかの経路でたどったグループのため、ここでは繰り返しません' }
            else { $recurse = $true }
        }
        Add-GroupRow $Top.Name $account.Name $account.Kind $Via $(if ($member.Disabled) { '無効' } else { '' }) $account.Sid $note
        if ($recurse) {
            [void]$Visited.Add($account.Sid)
            $nextVia = if ($Via) { $Via + ' → ' + $account.Name } else { $account.Name }
            Add-GroupMember $Top $account $System $nextVia $Visited
        }
    }
}

#endregion

try {
    #region 調べる

    # 調べるフォルダを順に調べる。最後に、外側に含まれるとして外したのに、外側の途中を読めずに届かなかった内側のフォルダを足して調べる
    $scanList = [System.Collections.Generic.List[object]]::new()
    foreach ($target in $targets) { $scanList.Add($target) }
    $scanIndex = 0
    $pendingAdded = $false
    while ($true) {
        if ($scanIndex -ge $scanList.Count) {
            if ($pendingAdded) { break }
            $pendingAdded = $true
            foreach ($dropped in $droppedTargets) { if ($pendingKeys.Contains($dropped.Key)) { $scanList.Add($dropped.Target) } }
            $pendingKeys.Clear()
            continue
        }
        $target = $scanList[$scanIndex]
        $scanIndex++
        # 共有フォルダを並べて調べるときは、開けない共有があっても止めずに残す（-Path に書いたフォルダは、調べる前に確かめてある）
        if (-not $target.Strict) {
            $openError = $null
            try { [void][IO.Directory]::EnumerateFileSystemEntries($target.Long).GetEnumerator().MoveNext() } catch { $openError = Get-InnerException $_.Exception }
            if ($openError) {
                if ($openError -is [UnauthorizedAccessException]) { Add-Skip '権限なし' $target.Full '共有フォルダを開けません（アクセスが拒否されました）' }
                else { Add-Skip 'その他のエラー' $target.Full ('共有フォルダを開けません（{0}）' -f (Get-ErrorText $openError)) }
                continue
            }
        }
        $prefixLen = $target.Long.TrimEnd('\').Length + 1
        # 再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる。
        # 各要素は、パス・階層・親から継承するはずの設定（同じドライブの中で移動したフォルダなど、親と合わない継承を見つけるため）
        $stack = [System.Collections.Generic.Stack[object]]::new()
        $stack.Push([object[]]@($target.Long, 0, (Get-RootExpected $target)))
        while ($stack.Count -gt 0) {
            $frame = $stack.Pop()
            $long = [string]$frame[0]
            $level = [int]$frame[1]
            $display = ConvertTo-DisplayPath $long
            if ($level -eq 0) { $display = $target.Full }
            if ($pendingKeys.Count -gt 0) { [void]$pendingKeys.Remove((Get-FolderOverlapKey $target $display)) }
            $folderCount++
            if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                Write-Progress -Activity $ToolTitle -Status (Format-Invariant '調べています: フォルダ {0:N0} 件  アクセス権を変えてあるフォルダ {1:N0} 件' $folderCount $changedFolders) -CurrentOperation $display
                $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
            }
            $security = [WbFolderPermNative]::ReadSecurity($long)
            # 共有のルートで、継承の印も自動継承の印もなく、継承も切っていない（古い形の ACL）ときは、管理共有越しに読み直す（Get-ShareRootRereadPath を参照）。
            # 自動継承の印があれば、継承の印のない設定は本当にこのフォルダで設定したもの
            $creatorExpected = $frame[2]
            if ($level -eq 0 -and $security.Error -eq 0 -and -not $security.Protected -and -not $security.AutoInherited -and $security.HasExplicit -and -not $security.HasInherited) {
                $rereadPath = Get-ShareRootRereadPath $target
                if ($null -ne $rereadPath) {
                    $reread = $null
                    if ($rereadPath) { $reread = [WbFolderPermNative]::ReadSecurity($rereadPath) }
                    if ($null -ne $reread -and $reread.Error -eq 0) {
                        $security = $reread
                        if ($null -eq $creatorExpected) { $creatorExpected = Get-RereadParentExpected $rereadPath }
                    } else {
                        $unsureRoots.Add($target.Full)
                    }
                }
            }
            # 古い形の ACL で、CREATOR OWNER から作られた所有者の設定が「このフォルダ」に見えるときは、継承した設定として扱う
            if ($security.Error -eq 0 -and $null -ne $creatorExpected) { [WbFolderPermNative]::MarkCreatorAces($security, $creatorExpected) }
            if ($security.Error -eq 0 -and $null -ne $frame[2]) {
                $security.Stale = -not [WbFolderPermNative]::InheritanceMatches($frame[2], $security.Descriptor)
            }
            # 同じフォルダを 2 つの名前で共有しているときは、両方の名前（「data、data2」）を出すため、表から引いた名前を先に使う
            $shareName = Get-ShareNameFor $target $display
            if (-not $shareName -and $level -eq 0) { $shareName = $target.ShareName }
            if ($security.Changed) { $changedFolders++ }
            if ($security.Error -eq 0 -and ($level -eq 0 -or $shareName -or $All -or $security.Changed)) {
                Add-Item $display $security $false $level ($level -eq 0) $shareName $target.System
            }
            $listError = $null
            if ($depthLimit -eq 0 -or $level -lt $depthLimit) {
                $entries = @()
                try {
                    $dirInfo = [IO.DirectoryInfo]::new($long)
                    # if 文の出力を代入すると要素 1 つの配列がほどけるため、@( ) で囲む
                    $entries = @(if ($IncludeFile) { $dirInfo.GetFileSystemInfos() } else { $dirInfo.GetDirectories() })
                } catch {
                    $listError = Get-InnerException $_.Exception
                }
                # 子が継承するはずの設定（このフォルダのアクセス権を読めないときは、比べない）
                $dirExpected = $null
                $fileExpected = $null
                if ($security.Error -eq 0 -and $entries.Count -gt 0) {
                    $dirExpected = [WbFolderPermNative]::ExpectedFor($security.Descriptor, $true)
                    if ($IncludeFile) { $fileExpected = [WbFolderPermNative]::ExpectedFor($security.Descriptor, $false) }
                }
                foreach ($entry in $entries) {
                    $isDir = $entry -is $dirType
                    $entryDisplay = ConvertTo-DisplayPath $entry.FullName
                    if ($level -eq 0) { $entryDisplay = $target.Full.TrimEnd('\') + '\' + $entry.Name }
                    # ごみ箱と System Volume Information は、Windows が管理するため調べない
                    if ($isDir -and $SystemFolders -contains $entry.Name) {
                        $excludedCount++
                        Add-Skip '除外指定' $entryDisplay 'ごみ箱・システムの管理用のフォルダ'
                        continue
                    }
                    if ($hasExclude) {
                        $matched = Get-ExcludeMatch $entry.FullName.Substring($prefixLen) $entry.Name $entryDisplay
                        if ($matched) {
                            $excludedCount++
                            Add-Skip '除外指定' $entryDisplay (Protect-ExcelText $matched)
                            continue
                        }
                    }
                    if (($entry.Attributes -band $rpFlag) -eq $rpFlag) {
                        $linkName = Get-LinkName $entry
                        if ($linkName) {
                            $linkCount++
                            Add-Skip 'リンク' $entryDisplay $linkName
                            continue
                        }
                    }
                    if ($isDir) {
                        $stack.Push([object[]]@($entry.FullName, ($level + 1), $dirExpected))
                        continue
                    }
                    $fileCount++
                    # ファイルの多いフォルダでも、表示が止まって見えないよう、ファイルごとにも進捗を確かめる
                    if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                        Write-Progress -Activity $ToolTitle -Status (Format-Invariant '調べています: フォルダ {0:N0} 件  ファイル {1:N0} 件' $folderCount $fileCount) -CurrentOperation $entryDisplay
                        $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
                    }
                    $fileSecurity = [WbFolderPermNative]::ReadSecurity($entry.FullName)
                    if ($fileSecurity.Error -ne 0) {
                        $failure = Get-ReadFailure $fileSecurity.Error 'アクセス権を読めません'
                        Add-Skip $failure[0] $entryDisplay $failure[1]
                        continue
                    }
                    if ($null -ne $fileExpected) {
                        $fileSecurity.Stale = -not [WbFolderPermNative]::InheritanceMatches($fileExpected, $fileSecurity.Descriptor)
                    }
                    if ($fileSecurity.Changed) {
                        $changedFiles++
                        Add-Item $entryDisplay $fileSecurity $true ($level + 1) $false '' $target.System
                    }
                }
            }
            # 読めなかったものは、1 つのフォルダにつき 1 行にまとめて残す
            if ($security.Error -ne 0 -and $listError) {
                $failure = Get-ReadFailure $security.Error 'アクセス権も中の一覧も読めません'
                Add-Skip $failure[0] $display $failure[1]
            } elseif ($security.Error -ne 0) {
                $failure = Get-ReadFailure $security.Error 'アクセス権を読めません'
                Add-Skip $failure[0] $display $failure[1]
            } elseif ($listError) {
                # 種類を判定できないリンクは、読み取りが禁止されていると普通のフォルダに見える。権限なしには数えない
                if ($listError -is [UnauthorizedAccessException]) { Add-Skip '権限なし' $display '中の一覧を読めません（アクセスが拒否されました）' }
                else { Add-Skip 'その他のエラー' $display ('中の一覧を読めません（{0}）' -f (Get-ErrorText $listError)) }
            }
        }
    }
    $scanElapsed = $sw.Elapsed
    # 調べ終わった後の並べ替え・表示・保存も、数が多いと時間がかかるため、そのことを出しておく
    Write-Progress -Activity $ToolTitle -Status '結果をまとめています'

    # フォルダとファイルは、パスの順に並べる（フォルダの中身は、そのフォルダのすぐ後ろ）
    $itemArray = $items.ToArray()
    $itemKeys = [string[]]::new($itemArray.Count)
    for ($k = 0; $k -lt $itemArray.Count; $k++) { $itemKeys[$k] = Get-PathSortText $itemArray[$k].Display }
    $sortedItems = Get-SortedList $itemKeys $itemArray

    # 共有フォルダは、名前の順に並べる
    $shareArray = $shares.ToArray()
    $shareKeyTexts = [string[]]@(foreach ($share in $shareArray) { Get-SortText $share.Display })
    $shares = Get-SortedList $shareKeyTexts $shareArray

    #endregion

    #region グループのメンバーを調べる

    $topGroups = @()
    if ($ExpandGroup) {
        Write-Progress -Activity $ToolTitle -Status 'グループのメンバーを調べています'
        try { $domainJoined = [bool](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).PartOfDomain } catch { }
        $localAccountRun = [string]::Equals([Environment]::UserDomainName, $machineName, [StringComparison]::OrdinalIgnoreCase)
        # アクセス権（共有のアクセス許可を含む）に出てくるグループと組み込みのアカウント。問い合わせる PC ごとに分ける
        $sourceItems = [System.Collections.Generic.List[object]]::new()
        foreach ($share in $shares) { if ($share.Item) { $sourceItems.Add($share.Item) } }
        foreach ($item in $sortedItems) { $sourceItems.Add($item) }
        $topArray = @(foreach ($pair in [WbFolderPermNative]::CollectGroups([WbFolderPermItem[]]$sourceItems.ToArray())) {
            [pscustomobject]@{ Account = $pair[0]; System = $(if ($pair[1]) { [string]$pair[1] } else { $null }) }
        })
        $topKeys = [string[]]@(foreach ($top in $topArray) { (Get-SortText $top.Account.Name) + "`0" + [string]$top.System })
        $topGroups = Get-SortedList $topKeys $topArray
        foreach ($top in $topGroups) {
            $account = $top.Account
            if ($account.Kind -eq $KindBuiltin) {
                Add-GroupRow $account.Name '' '' '' '' $account.Sid '組み込みのアカウントのため、メンバーを調べません（Windows が条件で決めるため）'
                continue
            }
            if ($account.IsWide) {
                Add-GroupRow $account.Name '' '' '' '' $account.Sid 'ほぼ全員が入っているグループのため、メンバーを調べません'
                continue
            }
            $visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            [void]$visited.Add($account.Sid)
            Add-GroupMember $account $account $top.System '' $visited
        }
        Write-Progress -Activity $ToolTitle -Completed
    }

    #endregion

    #region 一覧の表示

    Write-Line (Format-Invariant '{0} (Get-FolderPermission {1})' $ToolTitle $ScriptVersion)
    Write-Line ('対象: {0}' -f ($targetTexts -join ', '))
    Write-Line (Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（管理者: {2}）' $startTime $machineName $(if ($isAdmin) { 'はい' } else { 'いいえ' }))
    if ($depthLimit -gt 0) { Write-Line (Format-Invariant '階層: {0} 階層下まで' $depthLimit) }
    if ($hasExclude) { Write-Line ('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', ')) }
    if ($IncludeFile) {
        Write-Line (Format-Invariant 'フォルダ {0:N0} 件  ファイル {1:N0} 件  アクセス権を変えてあるフォルダ {2:N0} 件・ファイル {3:N0} 件  所要 {4}' $folderCount $fileCount $changedFolders $changedFiles (Format-Duration $scanElapsed))
    } else {
        Write-Line (Format-Invariant 'フォルダ {0:N0} 件  アクセス権を変えてあるフォルダ {1:N0} 件  所要 {2}' $folderCount $changedFolders (Format-Duration $scanElapsed))
    }
    $noteLine = Get-NoteLine
    if ($noteLine) { Write-Line $noteLine }
    $deniedShare = @($shares | Where-Object { $_.Error }).Count -gt 0
    if (-not $isAdmin -and ($skipKinds['権限なし'] -gt 0 -or $deniedShare)) { Write-Line '※ 管理者として実行すると、読める範囲が広がります' }
    # 「確かめたほうがよいところ」の集計（数が多いと 5.1 のループでは遅いため、C# で数える）
    $summary = [WbFolderPermNative]::Summarize([WbFolderPermItem[]]$sortedItems)
    # 名前を調べられなかったアカウント（削除されたとは限らないので、「名前が見つからない」とは分ける）
    $lookupFailed = $summary.LookupFailed
    foreach ($share in $shares) {
        if ($share.Item) { foreach ($ace in $share.Item.Aces) { if ($ace.Account.Kind -eq [WbFolderPermNative]::KindLookupFailed) { $lookupFailed = $true } } }
    }
    if ($lookupFailed) { Write-Line '※ アカウントの名前を調べられなかったものがあります（ドメインコントローラーにつながらないなど）。アカウントの列には SID を出しています' }
    if ($unsureRoots.Count -gt 0) {
        $unsureText = (@($unsureRoots | Select-Object -First 3) -join '、') + $(if ($unsureRoots.Count -gt 3) { Format-Invariant ' など {0:N0} 件' $unsureRoots.Count } else { '' })
        Write-Line ('※ 共有フォルダのフォルダで、親から継承した設定を見分けられなかったものがあります（{0}）。継承した設定も「このフォルダ」で設定したものとして出しています。サーバーの管理者として実行すると見分けられます' -f $unsureText)
    }
    Write-Line ''

    # 共有フォルダ
    Write-Line '■ 共有フォルダ'
    if ($shares.Count -eq 0) {
        Write-Line '    （このフォルダにつながる共有フォルダはありません）'
    } else {
        $shareLines = [System.Collections.Generic.List[object]]::new()
        foreach ($share in $shares) {
            if ($null -eq $share.Item) {
                $text = if ($share.Error) { '（読めません: {0}）' -f $share.Error } else { '（{0}）' -f $share.Note }
                $shareLines.Add([object[]]@($share.Display, $share.Path, '', $text, ''))
                continue
            }
            $first = $true
            foreach ($ace in $share.Item.Aces) {
                $shareLines.Add([object[]]@($(if ($first) { $share.Display } else { '' }), $(if ($first) { $share.Path } else { '' }), $(if ($ace.Deny) { '拒否' } else { '許可' }), $ace.Account.Name, $ace.Rights))
                $first = $false
            }
            if ($first) { $shareLines.Add([object[]]@($share.Display, $share.Path, '', '（だれにも許可していません）', '')) }
        }
        # 列は値に合わせて広げるが、長い値が 1 つあっても全体がずれないよう、上限（フォルダは 50、ほかは 40 文字分）を設ける
        $widths = @(6, 8, 4, 10)
        $limits = @(40, 50, 40, 40)
        foreach ($row in $shareLines) { for ($c = 0; $c -lt 4; $c++) { $widths[$c] = [math]::Max($widths[$c], [math]::Min($limits[$c], (Get-DisplayWidth $row[$c]))) } }
        Write-Line (' ' + (Format-Column '共有名' $widths[0] $false) + '  ' + (Format-Column 'フォルダ' $widths[1] $false) + '  ' + (Format-Column '種類' $widths[2] $false) + '  ' + (Format-Column 'アカウント' $widths[3] $false) + '  アクセス許可')
        foreach ($row in $shareLines) {
            Write-Line ((' ' + (Format-Column $row[0] $widths[0] $false) + '  ' + (Format-Column $row[1] $widths[1] $false) + '  ' + (Format-Column $row[2] $widths[2] $false) + '  ' + (Format-Column $row[3] $widths[3] $false) + '  ' + $row[4]).TrimEnd())
        }
    }
    Write-Line ''

    # アクセス権を変えてあるフォルダ（-All のときは、すべてのフォルダ）
    $sectionTitle = if ($All) { '■ フォルダのアクセス権（すべてのフォルダ）' } elseif ($IncludeFile) { '■ アクセス権を変えてあるフォルダとファイル' } else { '■ アクセス権を変えてあるフォルダ' }
    Write-Line $sectionTitle
    # 一覧の行は、数が多いと 5.1 では組み立てが遅いため、C# で作る。
    # -Path と共有フォルダのフォルダ、継承が親と合わないフォルダは、継承した分も含めてすべて。それ以外は、そのフォルダで設定した分だけ
    $blockLines = [WbFolderPermNative]::FormatItems([WbFolderPermItem[]]$sortedItems)
    if ($blockLines.Count -eq 0) {
        Write-Line '    （アクセス権を変えてあるフォルダはありません）'
    }
    $shownLimit = if ($runDir) { $ItemRows } else { [int]::MaxValue }
    # -NoSave で数が多いときも速く出せるよう、画面にはまとめて書く（1 行ずつ Write-Host すると、5.1 では何分もかかる）
    $screen = [Text.StringBuilder]::new()
    for ($k = 0; $k -lt $blockLines.Count; $k++) {
        $lines.AddRange([string[]]$blockLines[$k])
        if ($k -lt $shownLimit) {
            [void]$screen.Append([string]::Join("`r`n", [string[]]$blockLines[$k])).Append("`r`n")
            if ($screen.Length -ge 65536) {
                Write-Host $screen.ToString().TrimEnd([char[]]"`r`n")
                [void]$screen.Clear()
            }
        } elseif ($k -eq $shownLimit) {
            if ($screen.Length -gt 0) { Write-Host $screen.ToString().TrimEnd([char[]]"`r`n"); [void]$screen.Clear() }
            $unit = if ($IncludeFile) { '件' } else { 'フォルダ' }
            Write-Host (Format-Invariant '（ほか {0:N0} {1}は permissions.csv を参照してください）' ($blockLines.Count - $shownLimit) $unit)
        }
    }
    if ($screen.Length -gt 0) { Write-Host $screen.ToString().TrimEnd([char[]]"`r`n") }
    Write-Line ''

    # 確かめたほうがよいところ。アクセス権の行を、設定した場所で数える（継承した設定は、-Path のフォルダと、継承が親と合わないフォルダの行でだけ数える）
    Write-Line '■ 確かめたほうがよいところ'
    $findings = [System.Collections.Generic.List[object]]::new()
    if ($summary.ProtectedFolders + $summary.ProtectedFiles -gt 0) {
        $countText = if ($summary.ProtectedFiles -gt 0) { Format-Invariant '{0:N0} フォルダ・{1:N0} ファイル' $summary.ProtectedFolders $summary.ProtectedFiles } else { Format-Invariant '{0:N0} フォルダ' $summary.ProtectedFolders }
        $findings.Add([object[]]@([WbFolderPermNative]::NoteName($NoteProtected), $countText))
    }
    $staleItems = $summary.Stale
    $orderItems = $summary.Order
    if ($staleItems.Count -gt 0) {
        $exampleText = (@($staleItems | Select-Object -First $ExampleCount)) -join '、'
        if ($staleItems.Count -gt $ExampleCount) { $exampleText += Format-Invariant ' ほか {0:N0}' ($staleItems.Count - $ExampleCount) }
        $findings.Add([object[]]@([WbFolderPermNative]::NoteName([WbFolderPermNative]::NoteStale), (Format-Invariant '{0:N0} 件（{1}）' $staleItems.Count $exampleText)))
    }
    $noteBits = @($NoteUser, $NoteUnknown, $NoteWide, $NoteDeny)
    for ($b = 0; $b -lt $noteBits.Count; $b++) {
        $bit = $noteBits[$b]
        if ($summary.Counts[$b] -eq 0) { continue }
        $examples = $summary.Examples[$b]
        $shown = @($examples | Select-Object -First $ExampleCount)
        $exampleText = $shown -join '、'
        if ($examples.Count -gt $ExampleCount) { $exampleText += Format-Invariant ' ほか {0:N0}' ($examples.Count - $ExampleCount) }
        $findings.Add([object[]]@([WbFolderPermNative]::NoteName($bit), (Format-Invariant '{0:N0} 件（{1}）' $summary.Counts[$b] $exampleText)))
    }
    if ($orderItems.Count -gt 0) {
        $exampleText = (@($orderItems | Select-Object -First $ExampleCount)) -join '、'
        if ($orderItems.Count -gt $ExampleCount) { $exampleText += Format-Invariant ' ほか {0:N0}' ($orderItems.Count - $ExampleCount) }
        $findings.Add([object[]]@([WbFolderPermNative]::NoteName($NoteOrder), (Format-Invariant '{0:N0} 件（{1}）' $orderItems.Count $exampleText)))
    }
    if ($findings.Count -eq 0) {
        Write-Line '    （確かめたほうがよいところはありません）'
    } else {
        $nameWidth = 0
        foreach ($finding in $findings) { $nameWidth = [math]::Max($nameWidth, (Get-DisplayWidth $finding[0])) }
        foreach ($finding in $findings) { Write-Line (' ' + (Format-Column $finding[0] $nameWidth $false) + '  ' + $finding[1]) }
    }

    # グループのメンバー
    if ($ExpandGroup) {
        Write-Line ''
        Write-Line '■ グループのメンバー'
        $memberCount = @($groupRows | Where-Object { $_[1] }).Count
        if ($runDir) {
            Write-Line (Format-Invariant ' グループ {0:N0} 件  メンバー {1:N0} 件（groups.csv を参照してください）' $topGroups.Count $memberCount)
        } else {
            Write-Line (Format-Invariant ' グループ {0:N0} 件  メンバー {1:N0} 件' $topGroups.Count $memberCount)
        }
        # 保存するときは summary.txt にだけ、保存しないときは画面にも、メンバーの一覧を出す
        foreach ($row in $groupRows) {
            $text = '    ' + $row[0]
            if ($row[1]) { $text += '  ' + $row[1] + '（' + $row[2] + '）' }
            if ($row[3]) { $text += '  経由: ' + $row[3] }
            if ($row[4]) { $text += '  ' + $row[4] }
            if ($row[6]) { $text += '  ' + $row[6] }
            Add-SummaryLine $text (-not $runDir)
        }
        foreach ($failure in $groupFailures) { Write-Line (' 調べられなかったグループ: ' + $failure) }
    }
    Write-Line ''

    #endregion

    if ($runDir) { $lines.Add(('保存先: {0}' -f $runDir)) }
    Save-Result
    $finished = $true
} catch {
    # 想定外のエラーは、Ctrl+C の「中断しました」と区別して知らせる
    $unexpected = $true
    throw
} finally {
    Write-Progress -Activity $ToolTitle -Completed
    if (-not $finished -or $saveError) {
        # 途中で止まったときと、保存の途中で失敗したときは、書きかけの結果を残さない（読むだけなので、ファイルは変わっていない）
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        # 止めた理由は、Stop-WithError が表示している。保存の失敗は、この後で警告を出す
        if (-not $finished -and -not $stopping) {
            if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
            else { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
        }
    }
}

if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}（すべてを画面で見るときは、-NoSave を付けて実行し直してください）' -f $saveError)
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
} elseif ($saveWarning) {
    Write-Warning $saveWarning
}
exit 0
