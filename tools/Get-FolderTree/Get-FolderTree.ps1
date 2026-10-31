<#
.SYNOPSIS
    フォルダの構成を、罫線を使ったツリー形式のテキストで出力します。

.DESCRIPTION
    指定したフォルダの配下を、指定した階層の深さまでツリー形式で画面に表示します。
    結果はテキストと CSV（BOM 付き UTF-8）でも保存します。設計書や手順書にフォルダ構成を貼るときに使います。

    調べる場所のファイルを変更・削除することはありません。
    権限がなくて読めないフォルダはスキップし、ツリーに印を付けて件数を表示します。
    ジャンクションやシンボリックリンクはたどりません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Get-FolderTree.ps1 -Path D:\project

.PARAMETER Path
    調べるフォルダ。ドライブ（C:\）や共有フォルダ（\\server\share）も指定できます。
    省略すると、現在のフォルダを調べます。

.PARAMETER Depth
    出力する階層の深さ（1～100）。Path の直下が 1 です。既定値は 3 です。

.PARAMETER Files
    フォルダに加えて、ファイルも出力します。

.PARAMETER Hidden
    隠し属性またはシステム属性の付いたフォルダ・ファイルも出力します。

.PARAMETER Detail
    フォルダに合計サイズとファイル数を、ファイルにサイズと最終更新日時を添えます。
    合計は Depth より下の階層も含めて数えるため、全階層を読みます。

.PARAMETER Exclude
    出力しないフォルダ。フォルダのフルパスに対して、ワイルドカード（*）で指定します。
    * 以外の文字（[ ] など）は、そのままの文字として扱います。
    大文字と小文字は区別しません。複数指定するときは、カンマで区切ります。

.PARAMETER Ascii
    罫線（├─ └─ │）の代わりに、ASCII 文字（+-- \-- |）で描きます。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.PARAMETER ConfigFile
    設定ファイルのパス。Path～NoSave の値を「項目 = 値」の形で書いておけます。
    コマンドラインで指定した値は、設定ファイルより優先します。書き方は README を参照してください。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-FolderTree.ps1 -Path D:\project

    D:\project の 3 階層までのフォルダを出力します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-FolderTree.ps1 -Path D:\project -Files -Detail

    ファイルも出力し、フォルダの合計サイズとファイル数、ファイルのサイズと最終更新日時を添えます。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-FolderTree.ps1 -Path D:\project -Exclude "*\.git,*\node_modules"

    すべての .git フォルダと node_modules フォルダを出力しません。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-FolderTree.ps1 -ConfigFile .\project.ini

    設定ファイル project.ini に書いた条件で出力します。

.NOTES
    Get-FolderTree 1.0.0
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
    [string]$Path = '.',

    [ValidateRange(1, 100)]
    [int]$Depth = 3,

    [switch]$Files,

    [switch]$Hidden,

    [switch]$Detail,

    [string[]]$Exclude = @(),

    [switch]$Ascii,

    [string]$OutputFolder,

    [switch]$NoSave,

    [string]$ConfigFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$ScreenLines   = 500    # 画面に表示するツリーの行数（保存するときだけ。超えた分はファイルを見てもらう）
$ProgressMs    = 250    # 進捗表示の間隔（ミリ秒）
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$OnWindows     = [IO.Path]::DirectorySeparatorChar -eq '\'   # 開発用に Mac の pwsh でも動かせるようにするため

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

# 5.1 と 7 のどちらでも BOM 付き UTF-8 で保存する（Export-Csv の -Encoding UTF8 は 7 だと BOM なしになる）。
# 行が 0 件のときは見出しだけを書く
function Save-Csv($Rows, [string[]]$Columns, [string]$FilePath) {
    if ($null -eq $Rows) { $Rows = @() }
    $lines = @($Rows | ConvertTo-Csv -NoTypeInformation)
    if ($lines.Count -eq 0) { $lines = @(($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') }
    [IO.File]::WriteAllLines($FilePath, [string[]]$lines, $Utf8Bom)
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

# ツリーの 1 行分（フォルダかファイル）。Children と FileChildren は、表示するときに名前の順に並べる
function New-TreeNode([string]$Name, [string]$FullPath, [int]$Level, [bool]$IsDir) {
    # ファイル数が多いと遅くなるため、New-Object ではなく ::new で作り、ファイルの行には作らない。
    # $(if ...) の中に書くと空のリストが展開されて $null になるため、変数に入れてから渡す
    $dirs  = $null
    $leafs = $null
    if ($IsDir) {
        $dirs  = [Collections.Generic.List[object]]::new()
        $leafs = [Collections.Generic.List[object]]::new()
    }
    return [pscustomobject]@{
        Name         = $Name
        Path         = $FullPath
        Level        = $Level
        IsDir        = $IsDir
        IsMore       = $false
        Status       = ''
        Size         = 0L
        FileCount    = 0L
        Time         = $null
        HasMore      = $false
        Children     = $dirs
        FileChildren = $leafs
    }
}

# 並べ替えのたびにスクリプトブロックをデリゲートに変えると遅いため、1 回だけ作って使い回す
$DigitRun         = [Text.RegularExpressions.Regex]::new('[0-9]+')
$NaturalEvaluator = [Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Value.TrimStart('0').PadLeft(20, '0') }

# エクスプローラーと同じく、名前の中の数字は数値として比べる（file2 が file10 より先）。
# 数字の並びを 0 で桁をそろえた文字列に置き換えて、序数で比べる。5.1 と 7 で並びが変わらないよう、カルチャには頼らない。
# 桁をそろえると file01 と file1 が同じになるため、元の名前を後ろに付けて順番を決める。
# [Array]::Sort に PSObject の配列を渡すと、並べ替わるのはコピーのほうになるため、文字列だけを並べ替え、末尾の番号で元の項目を引く
function Get-SortedNode($Nodes) {
    $items = $Nodes.ToArray()
    if ($items.Count -lt 2) { return , $items }
    $keys = New-Object 'string[]' $items.Count
    for ($i = 0; $i -lt $items.Count; $i++) {
        $natural = $DigitRun.Replace($items[$i].Name, $NaturalEvaluator)
        $keys[$i] = $natural + [char]0 + $items[$i].Name + [char]0 + $i
    }
    [Array]::Sort($keys, [StringComparer]::OrdinalIgnoreCase)
    $sorted = New-Object 'object[]' $items.Count
    for ($i = 0; $i -lt $keys.Count; $i++) { $sorted[$i] = $items[[int]$keys[$i].Substring($keys[$i].LastIndexOf([char]0) + 1)] }
    return , $sorted
}

# 子の行を、フォルダ・ファイル・「この先にもある」印の順に返す
function Get-ChildList($Node) {
    $list = [Collections.Generic.List[object]]::new()
    $list.AddRange([object[]](Get-SortedNode $Node.Children))
    $list.AddRange([object[]](Get-SortedNode $Node.FileChildren))
    if ($Node.HasMore) { $list.Add([pscustomobject]@{ IsMore = $true }) }
    return , $list
}

# 名前の後ろに付ける印（状態、または -Detail のサイズなど）
function Get-NodeSuffix($Node) {
    if ($Node.Status) { return '  [' + $Node.Status + ']' }
    if (-not $Detail) { return '' }
    if ($Node.IsDir) { return Format-Invariant '  ({0}・{1:N0} ファイル)' (Format-Size $Node.Size) $Node.FileCount }
    return Format-Invariant '  ({0}・{1:yyyy-MM-dd HH:mm})' (Format-Size $Node.Size) $Node.Time
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの。数値の範囲は param の ValidateRange と同じにする
$configItems   = 'Path', 'Depth', 'Files', 'Hidden', 'Detail', 'Exclude', 'Ascii', 'OutputFolder', 'NoSave'
$configSwitch  = 'Files', 'Hidden', 'Detail', 'Ascii', 'NoSave'
$configRanges  = @{ Depth = 1, 100 }
$configPath    = $null
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
        # エクスプローラーで拡張子を表示していないと、メモ帳で project.ini.txt として保存してしまうことがある
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

        # パスの項目は、コマンドプロンプトで書くときと同じように %TEMP% などの環境変数を展開する
        if ($itemName -eq 'Exclude') {
            # "a,b" のように全体を囲む書き方と、"a","b" のように 1 つずつ囲む書き方のどちらも受け付ける
            $source = if ($itemValue.Contains('"')) { $rawValue } else { $itemValue }
            foreach ($part in ($source -split ',')) {
                $pattern = Remove-ConfigQuote $part
                if ($pattern.Contains('"')) { Stop-WithConfigError $lineNumber ('Exclude の値に引用符（"）が残っています: {0}' -f $pattern) }
                if ($pattern) { $excludeItems.Add((Expand-ConfigVariable $lineNumber $pattern)) }
            }
        } elseif ($configSwitch -contains $itemName) {
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
        } else {
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

$Path = Remove-Quote $Path
if ($Path.Contains('"')) { Stop-WithQuoteError '-Path' $Path }
if ($Path -match '^(\\\\\?\\)?[A-Za-z]:$') { $Path += '\' }
$Path = Resolve-StartPath $Path

# -File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# コマンドプロンプトでは ' がそのまま届くため、'*\a,*\b' のように全体を囲んだ場合も、各要素の前後の ' を外す
# ワイルドカードは * だけにして、それ以外の文字は文字どおりに扱う。5.1 の WildcardPattern.Escape は
# バッククォートをエスケープしないため、自分でエスケープしてから照合用のパターンを作る
$wildcardOptions = [Management.Automation.WildcardOptions]'IgnoreCase, CultureInvariant'
$excludeRules = @(
    foreach ($item in $Exclude) {
        foreach ($part in ($item -split ',')) {
            $text = (Remove-Quote $part).Trim("'")
            if ($text.Contains('"')) { Stop-WithQuoteError '-Exclude' $text }
            if ($text -notmatch '^[A-Za-z]:\\$') { $text = $text.TrimEnd('\') }
            if ($text) {
                $escaped = $text.Replace('`', '``').Replace('[', '`[').Replace(']', '`]').Replace('?', '`?')
                [pscustomobject]@{ Text = $text; Matcher = [Management.Automation.WildcardPattern]::new($escaped, $wildcardOptions) }
            }
        }
    }
)
$hasExclude = $excludeRules.Count -gt 0

$outputText = $null
if ($OutputFolder) {
    $outputText = Remove-Quote $OutputFolder
    if ($outputText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outputText }
    $outputText = Resolve-StartPath $outputText
}

$target   = $null
$resolved = $null
try { $resolved = Resolve-Path -LiteralPath $Path } catch { }
if ($resolved -and $resolved.Provider.Name -eq 'FileSystem' -and (Test-Path -LiteralPath $resolved.ProviderPath -PathType Container)) {
    $target = ConvertTo-DisplayPath $resolved.ProviderPath
    # 通常の書き方では末尾のピリオドが無視されるが、\\?\ 付きでは残って読めなくなるため、ここで正規化する
    try { $target = [IO.Path]::GetFullPath($target) } catch { }
} else {
    # 5.1 の Resolve-Path は 260 文字を超えるパスを扱えないため、\\?\ を付けて確かめ直す（\\?\ 付きの指定も受け付ける）
    try {
        $full = ConvertTo-DisplayPath $Path
        if (-not [IO.Path]::IsPathRooted($full)) { $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($full) }
        if ([IO.Directory]::Exists((ConvertTo-LongPath $full))) { $target = $full }
    } catch { }
}
if (-not $target) { Stop-WithError (('指定したフォルダが見つかりません: {0}' -f $Path) + (Get-ConfigNote 'Path')) }
if ($target -notmatch '^[A-Za-z]:\\$') { $target = $target.TrimEnd('\') }

$startTime = Get-Date

$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

#endregion

#region 保存先の準備

# 調べ終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir      = $null
$saveWarning = $null
if (-not $NoSave) {
    $baseDir = if ($outputText) { $outputText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        $runName = Format-Invariant 'FolderTree_{0}_{1:yyyyMMdd_HHmmss}' $env:COMPUTERNAME $startTime
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

#region フォルダをたどる

$skipped       = New-Object 'System.Collections.Generic.List[object]'
$dirCount      = 0L   # ツリーに出したフォルダの数（-Path 自身を除く）
$fileCount     = 0L   # ツリーに出したファイルの数
$deniedCount   = 0
$errorCount    = 0
$linkCount     = 0
$excludedCount = 0

$dirType      = [IO.DirectoryInfo]
$rpFlag       = [IO.FileAttributes]::ReparsePoint
$hiddenSystem = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System

$rootNode = New-TreeNode $target $target 0 $true

# 再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる。
# 要素は (フォルダ, 階層, -Detail の合計を足すフォルダの記録, ツリーの行（Depth より下では $null）, リパースポイントかどうか)
$stack = New-Object 'System.Collections.Generic.Stack[object]'
$stack.Push(((New-Object IO.DirectoryInfo (ConvertTo-LongPath $target)), 0, @(), $rootNode, $false))

$sw           = [Diagnostics.Stopwatch]::StartNew()
$nextProgress = 0L
$completed    = $false
$fatal        = $null
try {
    while ($stack.Count -gt 0) {
        # PowerShell の変数名は大文字と小文字を区別しないため、パラメーターの $Depth や $Files と別の名前にする
        $frame     = $stack.Pop()
        $dir       = $frame[0]
        $level     = $frame[1]
        $chain     = $frame[2]
        $node      = $frame[3]
        $isReparse = $frame[4]

        if ($sw.ElapsedMilliseconds -ge $nextProgress) {
            Write-Progress -Activity 'フォルダツリー出力' `
                -Status ('フォルダ {0:N0} 件 / ファイル {1:N0} 件' -f $dirCount, $fileCount) `
                -CurrentOperation (ConvertTo-DisplayPath $dir.FullName)
            $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
        }

        try {
            $entries = $dir.GetFileSystemInfos()
        } catch {
            $ex   = Get-InnerException $_.Exception
            $disp = ConvertTo-DisplayPath $dir.FullName
            if ($level -eq 0) {
                $fatal = 'フォルダを読めません: {0}（{1}）' -f $disp, (Get-ErrorText $ex)
                break
            }
            if ($ex -is [UnauthorizedAccessException] -and $isReparse) {
                # 種類を判定できず、読み取りも禁止されたリパースポイントは、リンクとして扱う
                $linkCount++
                $status = 'リンク'
                $skipped.Add([pscustomobject]@{ '理由' = 'リンク'; 'パス' = $disp; '詳細' = '読み取りが禁止されたリンク' })
            } elseif ($ex -is [UnauthorizedAccessException]) {
                $deniedCount++
                $status = '権限なし'
                $skipped.Add([pscustomobject]@{ '理由' = '権限なし'; 'パス' = $disp; '詳細' = 'アクセスが拒否されました' })
            } else {
                $errorCount++
                $status = 'エラー'
                $skipped.Add([pscustomobject]@{ '理由' = 'その他のエラー'; 'パス' = $disp; '詳細' = (Get-ErrorText $ex) })
            }
            if ($node) { $node.Status = $status }
            continue
        }

        if ($Detail -and $node) { $chain = $chain + $node }
        $showChildren = $node -and $level -lt $Depth   # 子をツリーに出すか（出さないときは「この先にもある」かだけを見る）
        $countBelow   = $showChildren -or $Detail      # 子のリンクや除外を、件数と skipped.csv に入れるか
        $dirSize      = 0L
        $dirFiles     = 0L
        foreach ($entry in $entries) {
            # 隠し・システム属性の項目は、-Hidden を指定しないときは、ないものとして扱う（ツリーにも合計にも入れない）
            if (-not $Hidden -and ($entry.Attributes -band $hiddenSystem) -ne 0) { continue }
            $entryPath = ConvertTo-DisplayPath $entry.FullName

            if ($entry -is $dirType) {
                # 除外は、リンクかどうかより先に判定する（除外したリンクを [リンク] として表示しないため）
                if ($hasExclude) {
                    $matched = $null
                    foreach ($rule in $excludeRules) {
                        if ($rule.Matcher.IsMatch($entryPath)) { $matched = $rule.Text; break }
                    }
                    if ($matched) {
                        if ($countBelow) {
                            $excludedCount++
                            $skipped.Add([pscustomobject]@{ '理由' = '除外指定'; 'パス' = $entryPath; '詳細' = $matched })
                        }
                        continue
                    }
                }
                if ($node -and -not $showChildren) { $node.HasMore = $true }
                # Depth より下で -Detail もないときは、「この先にもある」かを決めるだけで、件数には入れない
                if (-not $countBelow) { continue }

                $entryIsReparse = ($entry.Attributes -band $rpFlag) -eq $rpFlag
                $linkName = $null
                if ($entryIsReparse) {
                    # OneDrive のフォルダもリパースポイントなので、リンクと判定できたものだけを飛ばす
                    $linkType = $null
                    try { $linkType = $entry.LinkType } catch { }
                    if ($linkType -eq 'Junction') {
                        $linkName = 'ジャンクション'
                    } elseif ($linkType -eq 'SymbolicLink') {
                        $linkName = 'シンボリックリンク'
                    } elseif (($entry.Attributes -band $hiddenSystem) -eq $hiddenSystem) {
                        # Windows が互換用に作ったリンク（All Users、Application Data など）は読み取りが禁止されていて、
                        # 5.1 では種類を判定できない。どれも隠し・システム属性付きなので、それを目印にする
                        $linkName = 'Windows の互換用リンク'
                    }
                }
                $childNode = $null
                if ($showChildren) {
                    $childNode = New-TreeNode $entry.Name $entryPath ($level + 1) $true
                    $node.Children.Add($childNode)
                    $dirCount++
                }
                if ($linkName) {
                    $linkCount++
                    $skipped.Add([pscustomobject]@{ '理由' = 'リンク'; 'パス' = $entryPath; '詳細' = $linkName })
                    if ($childNode) { $childNode.Status = 'リンク' }
                    continue
                }
                # 子をツリーに出すとき（Depth まで）と、-Detail で合計を数えるときだけ、中を読む
                if ($childNode -or $Detail) { $stack.Push(($entry, ($level + 1), $chain, $childNode, $entryIsReparse)) }
            } else {
                $length = $entry.Length
                $dirSize += $length
                $dirFiles++
                if ($Files) {
                    if ($showChildren) {
                        $fileNode = New-TreeNode $entry.Name $entryPath ($level + 1) $false
                        $fileNode.Size = $length
                        $fileNode.Time = $entry.LastWriteTime
                        $node.FileChildren.Add($fileNode)
                        $fileCount++
                    } elseif ($node) {
                        $node.HasMore = $true
                    }
                }
            }
        }

        foreach ($record in $chain) {
            $record.Size      += $dirSize
            $record.FileCount += $dirFiles
        }
    }
    if (-not $fatal) { $completed = $true }
} finally {
    Write-Progress -Activity 'フォルダツリー出力' -Completed
    if (-not $completed) {
        if ($runDir) { try { [IO.Directory]::Delete($runDir) } catch { } }
        if (-not $fatal) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}
if ($fatal) { Stop-WithError $fatal }
$elapsed = $sw.Elapsed

#endregion

# ツリーの組み立て・画面表示・保存の途中で Ctrl+C されたときや、保存に失敗したときは、書きかけの結果を残さない
$finished  = $false
$saveError = $null
try {
    #region ツリーの組み立て

    # 罫線は、Word や Excel の ＭＳ ゴシックで縦にそろうよう、全角幅の文字として組む（Windows の tree コマンドと同じ）
    if ($Ascii) {
        $glyph = @{ Tee = '+-- '; Last = '\-- '; Pipe = '|   '; Blank = '    '; More = '...' }
    } else {
        $glyph = @{ Tee = '├─'; Last = '└─'; Pipe = '│  '; Blank = '    '; More = '…' }
    }

    $treeLines = New-Object 'System.Collections.Generic.List[string]'
    $itemRows  = New-Object 'System.Collections.Generic.List[object]'
    $lineStack = New-Object 'System.Collections.Generic.Stack[object]'   # (行, 罫線の前置き, 最後の子かどうか)
    $treeLines.Add($target + (Get-NodeSuffix $rootNode))
    $lineStack.Push(($rootNode, $null, $true))
    while ($lineStack.Count -gt 0) {
        $frame  = $lineStack.Pop()
        $item   = $frame[0]
        $prefix = $frame[1]
        if ($null -ne $prefix) {
            $head = $prefix + $(if ($frame[2]) { $glyph.Last } else { $glyph.Tee })
            if ($item.IsMore) { $treeLines.Add($head + $glyph.More); continue }
            $treeLines.Add($head + $item.Name + (Get-NodeSuffix $item))
            $childPrefix = $prefix + $(if ($frame[2]) { $glyph.Blank } else { $glyph.Pipe })
        } else {
            $childPrefix = ''
        }

        # Excel の「標準」書式は 12 桁以上の数値を指数で表示するため、数値は 3 桁区切りで出す。
        # 名前は「2026-10-31」が日付に、「-old」が数式に変わるなど、数値以外でも別の値になるため、常に ="..." の形で出す
        $showSize = -not $item.Status -and ($Detail -or -not $item.IsDir)
        $itemRows.Add([pscustomobject]@{
            '種類'         = $(if ($item.IsDir) { 'フォルダ' } else { 'ファイル' })
            'パス'         = $item.Path
            '名前'         = '="' + $item.Name + '"'
            '階層'         = $item.Level
            'サイズ(MB)'   = $(if ($showSize) { ($item.Size / 1MB).ToString('N2', $Invariant) } else { '' })
            'サイズ(バイト)' = $(if ($showSize) { $item.Size.ToString('N0', $Invariant) } else { '' })
            'ファイル数'   = $(if ($showSize -and $item.IsDir) { $item.FileCount.ToString('N0', $Invariant) } else { '' })
            '最終更新日時' = $(if ($item.IsDir) { '' } else { $item.Time.ToString('yyyy/MM/dd HH:mm:ss', $Invariant) })
            '状態'         = $item.Status
        })

        if ($item.IsDir) {
            $kids = Get-ChildList $item
            for ($i = $kids.Count - 1; $i -ge 0; $i--) { $lineStack.Push(($kids[$i], $childPrefix, ($i -eq $kids.Count - 1))) }
        }
    }

    #endregion

    #region ヘッダー

    $conditions = @(('{0} 階層まで' -f $Depth), $(if ($Files) { 'フォルダとファイル' } else { 'フォルダのみ' }))
    if ($Hidden) { $conditions += '隠し項目を含む' }

    $header = New-Object 'System.Collections.Generic.List[string]'
    $header.Add((Format-Invariant 'フォルダツリー出力 (Get-FolderTree {0})' $ScriptVersion))
    $header.Add((Format-Invariant '対象: {0}   {1:yyyy-MM-dd HH:mm:ss}   {2}' $target $startTime $env:COMPUTERNAME))
    if ($configPath) { $header.Add(('設定ファイル: {0}' -f $configPath)) }
    $header.Add(('条件: {0}' -f ($conditions -join '・')))
    if ($hasExclude) { $header.Add(('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', '))) }
    if ($Files) {
        $header.Add((Format-Invariant 'フォルダ {0:N0} 件  ファイル {1:N0} 件  所要 {2}' $dirCount $fileCount (Format-Duration $elapsed)))
    } else {
        $header.Add((Format-Invariant 'フォルダ {0:N0} 件  所要 {1}' $dirCount (Format-Duration $elapsed)))
    }

    $notes = @()
    $kinds = @()
    if ($deniedCount -gt 0) { $kinds += , @('権限なし', $deniedCount) }
    if ($errorCount -gt 0) { $kinds += , @('その他のエラー', $errorCount) }
    if ($kinds.Count -eq 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $kinds[0][1] $kinds[0][0]
    } elseif ($kinds.Count -gt 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（権限なし {1:N0}・その他のエラー {2:N0}）' ($deniedCount + $errorCount) $deniedCount $errorCount
    }
    if ($linkCount -gt 0) { $notes += Format-Invariant 'たどらなかったリンク {0:N0} 件' $linkCount }
    if ($excludedCount -gt 0) { $notes += Format-Invariant '除外したフォルダ {0:N0} 件' $excludedCount }
    if ($notes.Count -gt 0) { $header.Add($notes -join '  ') }
    if ($deniedCount -gt 0 -and -not $isAdmin) { $header.Add('※ 管理者として実行すると、読める範囲が広がります') }

    #endregion

    #region 保存

    # 画面で「続きは tree.txt を」と案内するため、表示より先に保存して、保存できたかを確かめておく

    if ($runDir) {
        Write-Progress -Activity 'フォルダツリー出力' -Status '結果を保存しています'
        try {
            [IO.File]::WriteAllLines((Join-Path $runDir 'tree.txt'), [string[]]$treeLines, $Utf8Bom)
            [IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), [string[]](@($header) + '' + @($treeLines)), $Utf8Bom)
            Save-Csv $itemRows @('種類', 'パス', '名前', '階層', 'サイズ(MB)', 'サイズ(バイト)', 'ファイル数', '最終更新日時', '状態') (Join-Path $runDir 'items.csv')
            Save-Csv $skipped @('理由', 'パス', '詳細') (Join-Path $runDir 'skipped.csv')
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
        }
        Write-Progress -Activity 'フォルダツリー出力' -Completed
    }

    #endregion

    #region 画面表示

    foreach ($line in $header) { Write-Host $line }
    Write-Host ''
    # 保存できなかったときは、ほかに見る場所がないので全部表示する
    $shown = if ($runDir -and -not $saveError -and $treeLines.Count -gt $ScreenLines) { $ScreenLines } else { $treeLines.Count }
    for ($i = 0; $i -lt $shown; $i++) { Write-Host $treeLines[$i] }
    if ($shown -lt $treeLines.Count) {
        Write-Host (Format-Invariant '（ほか {0:N0} 行は tree.txt を参照してください）' ($treeLines.Count - $shown)) -ForegroundColor Yellow
    }

    #endregion

    $finished = $true
} finally {
    if (-not $finished -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete($runDir, $true) } catch { } }
        if (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

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

exit 0
