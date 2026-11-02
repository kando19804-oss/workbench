<#
.SYNOPSIS
    テキストファイルの文字コードと改行を判定し、一覧にします。指定すれば、確認したうえで変換します。

.DESCRIPTION
    指定したフォルダの中のテキストファイルについて、文字コード（UTF-8・Shift_JIS・UTF-16 など）と
    改行（CRLF・LF・CR）を判定し、画面に表示して files.csv（BOM 付き UTF-8）に保存します。

    -To や -LineEnding を指定したときは、変換するファイルの件数を表示して確認を求め、Y を入力すると変換します。
    元のファイルは、結果のフォルダの backup に控えます。変換した後に読み直し、内容が変わっていないかを確かめます。
    -To も -LineEnding も指定しないときは、ファイルを変更しません。

    権限がなくて読めないフォルダやファイルはスキップし、件数を表示します。
    ジャンクションやシンボリックリンクはたどりません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Convert-TextEncoding.ps1 -Path D:\scripts

.PARAMETER Path
    調べるフォルダかファイル。フォルダを指定すると、サブフォルダも含めて調べます。省略すると、現在のフォルダを調べます。

.PARAMETER Filter
    調べるファイルの名前の条件。ワイルドカード（* と ?）で、カンマで区切って書けます。
    省略すると、よくあるテキストの拡張子（*.txt、*.csv、*.ps1、*.bat など）を調べます。

.PARAMETER Exclude
    調べないフォルダやファイル。フルパスに対して、ワイルドカード（*）で指定します。
    * 以外の文字（[ ] など）は、そのままの文字として扱います。大文字と小文字は区別しません。
    複数指定するときは、カンマで区切ります。

.PARAMETER To
    変換後の文字コード。UTF8BOM（UTF-8・BOM あり）、UTF8（UTF-8・BOM なし）、SJIS（Shift_JIS）のどれかです。

.PARAMETER LineEnding
    変換後の改行。CRLF か LF です。

.PARAMETER Force
    確認せずに変換します。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。変換するときは使えません。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Convert-TextEncoding.ps1 -Path D:\scripts

    D:\scripts の中のテキストファイルの、文字コードと改行を判定します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Convert-TextEncoding.ps1 -Path D:\scripts -To UTF8BOM -LineEnding CRLF

    UTF-8（BOM あり）・改行 CRLF にそろえます（確認してから変換します）。

.NOTES
    Convert-TextEncoding 1.0.0
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

    [string]$Filter,

    [string[]]$Exclude = @(),

    [ValidateSet('UTF8BOM', 'UTF8', 'SJIS')]
    [string]$To,

    [ValidateSet('CRLF', 'LF')]
    [string]$LineEnding,

    [switch]$Force,

    [string]$OutputFolder,

    [switch]$NoSave
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$ScreenRows    = 200      # 画面に表示するファイルの件数（保存するときだけ。超えた分はファイルを見てもらう）
$LargeBytes    = 100MB    # これを超えるファイルは、先頭だけで判定して変換しない
$SniffBytes    = 1MB      # 大きなファイルの判定に読む長さ
$BinaryBytes   = 8KB      # バイナリかどうかの判定に見る長さ
$ProgressMs    = 250      # 進捗表示の間隔（ミリ秒）
$OutputPrefix  = 'TextEncoding_'
$DefaultFilter = '*.txt,*.csv,*.tsv,*.log,*.ini,*.cfg,*.conf,*.xml,*.json,*.yml,*.yaml,*.md,*.htm,*.html,*.css,*.js,*.sql,*.ps1,*.psm1,*.psd1,*.bat,*.cmd,*.vbs,*.sh,*.py'
$FileColumns   = 'ファイル', '文字コード', '改行', 'サイズ(バイト)', '最終更新日時', '結果', '備考'
$SkipColumns   = '理由', 'パス', '詳細'
$TimeFormat    = 'yyyy/MM/dd HH:mm:ss'
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$OnWindows     = [IO.Path]::DirectorySeparatorChar -eq '\'   # 開発用に Mac の pwsh でも動かせるようにするため
$machineName   = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

# 文字コードの名前。変換できるのは $Convertible のものだけ
$NameUtf8Bom  = 'UTF-8（BOM あり）'
$NameUtf8     = 'UTF-8（BOM なし）'
$NameUtf16    = 'UTF-16（BOM あり）'
$NameSjis     = 'Shift_JIS'
$NameAscii    = 'ASCII のみ'
$NameEuc      = 'EUC-JP'
$NameJis      = 'JIS'
$NameBinary   = 'バイナリ'
$NameUnknown  = '判定できない'
$Convertible  = $NameUtf8Bom, $NameUtf8, $NameUtf16, $NameSjis, $NameAscii

# 変換できない文字やバイトがあると例外を出す文字コード（判定と変換の確認に使う）
$strictUtf8 = New-Object Text.UTF8Encoding $false, $true
$strictSjis = [Text.Encoding]::GetEncoding(932, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
$strictLE   = New-Object Text.UnicodeEncoding $false, $false, $true
$strictBE   = New-Object Text.UnicodeEncoding $true, $false, $true
$latin1     = [Text.Encoding]::GetEncoding(28591)   # バイトをそのまま文字にする（改行を数えるため）

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

# ワイルドカードは * と ? だけにして、それ以外の文字は文字どおりに扱う。5.1 の WildcardPattern.Escape は
# バッククォートをエスケープしないため、自分でエスケープしてから照合用のパターンを作る
function New-Wildcard([string]$Text, [bool]$AllowQuestion) {
    $escaped = $Text.Replace('`', '``').Replace('[', '`[').Replace(']', '`]')
    if (-not $AllowQuestion) { $escaped = $escaped.Replace('?', '`?') }
    return [Management.Automation.WildcardPattern]::new($escaped, [Management.Automation.WildcardOptions]'IgnoreCase, CultureInvariant')
}

# -File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。各要素の前後の ' も外す
function Split-List([string[]]$Items) {
    $list = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $Items) {
        foreach ($part in ($item -split ',')) {
            $text = (Remove-Quote $part).Trim("'")
            if ($text) { $list.Add($text) }
        }
    }
    return , $list
}

# JIS（ISO-2022-JP）のエスケープ（ESC $ B、ESC $ @、ESC ( I、ESC $ ( D）があるか。1 件ずつ探すと遅いため、正規表現で 1 回だけ探す
$jisEscape = New-Object Text.RegularExpressions.Regex ('\x1B(?:\$[B@]|\(I|\$\(D)')

# EUC-JP として正しいか（0x80 以上のバイトが、2 バイト文字・半角カナ・補助漢字の形になっているか）。
# 1 バイトずつ調べると 5.1 では遅いため、バイトをそのまま文字にしたものから、正しい並びを正規表現で取り除き、残りがあるかで判定する。
# 先頭のバイトで長さが決まるので、左から取り除けば、正しい並びはすべて消える。$Partial のときは、末尾で文字が途中で切れていてもよい
$eucToken = New-Object Text.RegularExpressions.Regex ('[\x00-\x7F]+|[\xA1-\xFE][\xA1-\xFE]|\x8E[\xA1-\xDF]|\x8F[\xA1-\xFE][\xA1-\xFE]')
function Test-EucJp([string]$Raw, [bool]$Partial) {
    $rest = $eucToken.Replace($Raw, '')
    if ($rest.Length -eq 0) { return $true }
    return $Partial -and $rest.Length -le 2 -and $Raw.EndsWith($rest, [StringComparison]::Ordinal)
}

# 末尾の何バイトかを外して、厳密に文字に直す（大きなファイルを先頭だけ読んだとき、末尾で文字が切れていることがあるため）
function ConvertFrom-Bytes([Text.Encoding]$Encoding, [byte[]]$Bytes, [int]$Offset, [int]$Count, [int]$MaxCut) {
    for ($cut = 0; $cut -le $MaxCut -and $cut -le $Count; $cut++) {
        try { return $Encoding.GetString($Bytes, $Offset, $Count - $cut) } catch { }
    }
    return $null
}

# 改行の種類（CRLF・LF・CR・混在・なし）
function Get-LineKind([string]$Text) {
    $crlf = [regex]::Matches($Text, "`r`n").Count
    $lf   = [regex]::Matches($Text, "(?<!`r)`n").Count
    $cr   = [regex]::Matches($Text, "`r(?!`n)").Count
    $kinds = @()
    if ($crlf -gt 0) { $kinds += 'CRLF' }
    if ($lf -gt 0) { $kinds += 'LF' }
    if ($cr -gt 0) { $kinds += 'CR' }
    if ($kinds.Count -eq 0) { return 'なし' }
    if ($kinds.Count -gt 1) { return '混在' }
    return $kinds[0]
}

# 文字コードを判定する。Text は文字に直した内容（変換できる文字コードのときだけ）
function Get-TextInfo([byte[]]$Bytes, [int]$Count, [bool]$Partial) {
    $info = @{ Name = $NameUnknown; Text = $null; Line = '' }
    $maxCut = if ($Partial) { 3 } else { 0 }
    if ($Count -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        $text = ConvertFrom-Bytes $strictUtf8 $Bytes 3 ($Count - 3) $maxCut
        if ($null -ne $text) { $info.Name = $NameUtf8Bom; $info.Text = $text }
    } elseif ($Count -ge 2 -and (($Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) -or ($Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF))) {
        $encoding = if ($Bytes[0] -eq 0xFF) { $strictLE } else { $strictBE }
        $text = ConvertFrom-Bytes $encoding $Bytes 2 ($Count - 2) $maxCut
        if ($null -ne $text) { $info.Name = $NameUtf16; $info.Text = $text; $info.BigEndian = $Bytes[0] -eq 0xFE }
    } elseif ([Array]::IndexOf($Bytes, [byte]0, 0, [math]::Min($Count, $BinaryBytes)) -ge 0) {
        $info.Name = $NameBinary
        return $info
    } elseif ([Array]::IndexOf($Bytes, [byte]0x1B, 0, $Count) -ge 0 -and $jisEscape.IsMatch($latin1.GetString($Bytes, 0, $Count))) {
        $info.Name = $NameJis
    } else {
        # ASCII だけのファイルは UTF-8 としても正しいので、UTF-8 として読めたら、0x80 以上の文字があるかで分ける
        $text = ConvertFrom-Bytes $strictUtf8 $Bytes 0 $Count $maxCut
        if ($null -ne $text -and $text -notmatch '[^\x00-\x7F]') {
            $info.Name = $NameAscii
            $info.Text = $text
        } else {
            if ($null -ne $text) {
                $info.Name = $NameUtf8; $info.Text = $text
            } else {
                $sjisText = ConvertFrom-Bytes $strictSjis $Bytes 0 $Count ([math]::Min($maxCut, 1))
                $isEuc    = Test-EucJp ($latin1.GetString($Bytes, 0, $Count)) $Partial
                if ($null -ne $sjisText -and $isEuc) {
                    # EUC-JP の日本語を Shift_JIS として読むと、半角カナばかりになる
                    $kana = [regex]::Matches($sjisText, '[\uFF61-\uFF9F]').Count
                    $other = [regex]::Matches($sjisText, '[^\x00-\x7F\uFF61-\uFF9F]').Count
                    if ($kana -gt $other) { $info.Name = $NameEuc } else { $info.Name = $NameSjis; $info.Text = $sjisText }
                } elseif ($null -ne $sjisText) {
                    $info.Name = $NameSjis; $info.Text = $sjisText
                } elseif ($isEuc) {
                    $info.Name = $NameEuc
                }
            }
        }
    }
    # 改行は、UTF-16 は文字に直した内容で、それ以外はバイトのまま数える（CR と LF のバイトは、ASCII と同じ文字コードでは同じ値のため）
    if ($info.Name -eq $NameUtf16) { $info.Line = Get-LineKind $info.Text }
    elseif ($info.Name -ne $NameUnknown) { $info.Line = Get-LineKind ($latin1.GetString($Bytes, 0, $Count)) }
    return $info
}

# 変換後の文字コードの、名前・書き込み用のエンコーディング・BOM
function Get-TargetEncoding([string]$Key, $Info) {
    switch ($Key) {
        'UTF8BOM' { return @{ Name = $NameUtf8Bom; Encoding = $strictUtf8; Preamble = [byte[]](0xEF, 0xBB, 0xBF) } }
        'UTF8'    { return @{ Name = $NameUtf8; Encoding = $strictUtf8; Preamble = [byte[]]@() } }
        'SJIS'    { return @{ Name = $NameSjis; Encoding = $strictSjis; Preamble = [byte[]]@() } }
    }
    # -To を指定しないときは、元の文字コードのまま（BOM の有無も変えない）
    switch ($Info.Name) {
        $NameUtf8Bom { return @{ Name = $NameUtf8Bom; Encoding = $strictUtf8; Preamble = [byte[]](0xEF, 0xBB, 0xBF) } }
        $NameUtf16   {
            if ($Info.BigEndian) { return @{ Name = $NameUtf16; Encoding = $strictBE; Preamble = [byte[]](0xFE, 0xFF) } }
            return @{ Name = $NameUtf16; Encoding = $strictLE; Preamble = [byte[]](0xFF, 0xFE) }
        }
        $NameSjis    { return @{ Name = $NameSjis; Encoding = $strictSjis; Preamble = [byte[]]@() } }
        $NameAscii   { return @{ Name = $NameAscii; Encoding = $strictUtf8; Preamble = [byte[]]@() } }
    }
    return @{ Name = $NameUtf8; Encoding = $strictUtf8; Preamble = [byte[]]@() }
}

# バイト列が同じか。1 バイトずつ比べると 5.1 では遅いため、バイトをそのまま文字にして比べる。
# -ceq は NUL や C1 制御文字（0x80～0x9F）を無視して比べるため、序数で比べる
function Test-SameBytes([byte[]]$A, [byte[]]$B) {
    if ($A.Length -ne $B.Length) { return $false }
    return [string]::Equals($latin1.GetString($A), $latin1.GetString($B), [StringComparison]::Ordinal)
}

# 改行をそろえる
function ConvertTo-LineEnding([string]$Text, [string]$Newline) {
    $Text = $Text.Replace("`r`n", "`n").Replace("`r", "`n")
    if ($Newline -ne "`n") { $Text = $Text.Replace("`n", $Newline) }
    return $Text
}

# 開いたファイルの中身をすべて読む（配列のまま返すため、先頭にカンマを付ける）
function Read-StreamBytes([IO.Stream]$Stream) {
    $buffer = New-Object 'byte[]' ([int]$Stream.Length)
    $count  = 0
    while ($count -lt $buffer.Length) {
        $n = $Stream.Read($buffer, $count, $buffer.Length - $count)
        if ($n -le 0) { break }
        $count += $n
    }
    if ($count -lt $buffer.Length) { $short = New-Object 'byte[]' $count; [Array]::Copy($buffer, $short, $count); $buffer = $short }
    return , $buffer
}

# 変換の予定を立てたときの中身と、書き込む直前の中身が同じかを確かめるためのハッシュ
$sha256 = New-Object Security.Cryptography.SHA256CryptoServiceProvider
function Get-BytesHash([byte[]]$Bytes) { return [BitConverter]::ToString($sha256.ComputeHash($Bytes)) }

# 中身から、変換後のバイト列を作る。Status は「変換する」「すでに同じ」「対象外」「変換できない」のどれか
function Get-Conversion([byte[]]$Bytes) {
    $info   = Get-TextInfo $Bytes $Bytes.Length $false
    $result = @{ Status = ''; Info = $info; NewBytes = $null; NewText = $null; Target = $null; Change = ''; Note = ''; EncodingChanged = $false }
    if (-not ($Convertible -contains $info.Name)) {
        $result.Status = '対象外'
        $result.Note = '{0} は変換しません' -f $info.Name
        return $result
    }
    $targetEnc = Get-TargetEncoding $To $info
    $result.Target = $targetEnc
    $newline = if ($LineEnding -eq 'CRLF') { "`r`n" } else { "`n" }
    # Shift_JIS のまま（ASCII のみで BOM を付けないときも）は、文字に直さず、バイトのまま改行だけを置き換える。
    # CP932 には同じ文字になる重複した字（NEC と IBM の拡張文字など）があり、文字に直してから戻すとバイトが変わるため。
    # Shift_JIS の 2 バイト目に CR と LF の値は現れないので、バイトのまま置き換えてよい
    $byteMode = ($info.Name -eq $NameSjis -and $targetEnc.Name -eq $NameSjis) -or ($info.Name -eq $NameAscii -and $targetEnc.Preamble.Length -eq 0)
    if ($byteMode) {
        $raw = $latin1.GetString($Bytes)
        if ($LineEnding) { $raw = ConvertTo-LineEnding $raw $newline }
        $newBytes = $latin1.GetBytes($raw)
    } else {
        $newText = $info.Text
        if ($LineEnding) { $newText = ConvertTo-LineEnding $newText $newline }
        try {
            $body = $targetEnc.Encoding.GetBytes($newText)
        } catch {
            # EncoderFallbackException は MethodInvocationException に包まれて届く
            $ex = Get-InnerException $_.Exception
            if ($ex -isnot [Text.EncoderFallbackException]) { throw }
            $char = if ($ex.CharUnknownHigh -ne [char]0) { [string]$ex.CharUnknownHigh + [string]$ex.CharUnknownLow } else { [string]$ex.CharUnknown }
            $lineNo = 1 + [regex]::Matches($newText.Substring(0, [math]::Min($ex.Index, $newText.Length)), "`n").Count
            $result.Status = '変換できない'
            $result.Note = '{0} にない文字: {1}（{2} 行目）' -f $targetEnc.Name, $char, $lineNo
            return $result
        }
        $newBytes = New-Object 'byte[]' ($targetEnc.Preamble.Length + $body.Length)
        [Array]::Copy($targetEnc.Preamble, 0, $newBytes, 0, $targetEnc.Preamble.Length)
        [Array]::Copy($body, 0, $newBytes, $targetEnc.Preamble.Length, $body.Length)
        $result.NewText = $newText
    }
    $result.NewBytes = $newBytes
    if (Test-SameBytes $Bytes $newBytes) {
        $result.Status = 'すでに同じ'
        return $result
    }
    # 変換の内容。ASCII のみのファイルは、BOM を付けない限り、UTF-8 でも Shift_JIS でも同じバイトのままなので、文字コードの変更とは書かない
    $parts = @()
    $newLineKind = Get-LineKind $(if ($targetEnc.Name -eq $NameUtf16) { $result.NewText } else { $latin1.GetString($newBytes) })
    $result.EncodingChanged = $targetEnc.Name -ne $info.Name -and -not ($info.Name -eq $NameAscii -and $targetEnc.Preamble.Length -eq 0)
    if ($result.EncodingChanged) { $parts += '{0} → {1}' -f $info.Name, $targetEnc.Name }
    if ($newLineKind -ne $info.Line) { $parts += '改行 {0} → {1}' -f $info.Line, $newLineKind }
    $result.Change = $parts -join '・'
    $result.Status = '変換する'
    return $result
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

#endregion

#region 引数の確認

$converting = [bool]$To -or [bool]$LineEnding

$pathText = Remove-Quote $Path
if ($pathText.Contains('"')) { Stop-WithQuoteError '-Path' $pathText }
if ($pathText -match '^(\\\\\?\\)?[A-Za-z]:$') { $pathText += '\' }
$pathText = Resolve-StartPath $pathText
$target = ConvertTo-DisplayPath $pathText
try {
    $target = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($target)
    $target = [IO.Path]::GetFullPath($target)
} catch { }
if ($target -notmatch '^[A-Za-z]:\\$' -and $target -ne '/') { $target = $target.TrimEnd('\') }
$targetIsFile = $false
try { $targetIsFile = [IO.File]::Exists((ConvertTo-LongPath $target)) } catch { }
$targetIsDir = $false
if (-not $targetIsFile) { try { $targetIsDir = [IO.Directory]::Exists((ConvertTo-LongPath $target)) } catch { } }
if (-not $targetIsFile -and -not $targetIsDir) { Stop-WithError ('指定したフォルダかファイルが見つかりません: {0}' -f $Path) }
# 控えるときの相対パスの基準（ファイルを指定したときは、そのフォルダ）
$baseFolder = if ($targetIsFile) { Split-Path -LiteralPath $target } else { $target }
$basePrefixLen = $baseFolder.TrimEnd([char[]]@('\', '/')).Length + 1

$filterTexts = Split-List @($(if ($Filter) { $Filter } else { $DefaultFilter }))
foreach ($text in $filterTexts) { if ($text.Contains('"')) { Stop-WithQuoteError '-Filter' $text } }
$filterRules = @(foreach ($text in $filterTexts) { New-Wildcard $text $true })
$filterLabel = if ($Filter) { $filterTexts -join ', ' } else { '既定の拡張子' }

$excludeRules = @(
    foreach ($text in (Split-List $Exclude)) {
        if ($text.Contains('"')) { Stop-WithQuoteError '-Exclude' $text }
        if ($text -notmatch '^[A-Za-z]:\\$') { $text = $text.TrimEnd('\') }
        if ($text) { [pscustomobject]@{ Text = $text; Matcher = (New-Wildcard $text $false) } }
    }
)
$hasExclude = $excludeRules.Count -gt 0

$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    $outText = Resolve-StartPath $outText
}
if ($converting -and $NoSave) { Stop-WithError '変換するときは、-NoSave は指定できません（元のファイルを控える場所がないため）' }

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
        # 末尾のピリオドや .. を整えておく（\\?\ 付きで書く backup と、通常の書き方で書く summary が、別の場所に分かれないように）
        try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
        $runName = Format-Invariant '{0}{1}_{2:yyyyMMdd_HHmmss}' $OutputPrefix $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする。長いパスでも作れるよう \\?\ を付ける
        $suffix = 1
        while ([IO.Directory]::Exists((ConvertTo-LongPath $runDir)) -or [IO.File]::Exists((ConvertTo-LongPath $runDir))) { $suffix++; $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix) }
        [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $runDir))
    } catch {
        $reason = Get-ErrorText (Get-InnerException $_.Exception)
        $runDir = $null
        # 変換するときは、元のファイルを控えられないため止める
        if ($converting) { Stop-WithError ('保存先に書き込めないため、変換しません（元のファイルを控えられないため）: {0}（{1}）' -f $baseDir, $reason) }
        $saveWarning = '保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, $reason
        Write-Warning $saveWarning
    }
}

#endregion

$dirType   = [IO.DirectoryInfo]
$rpFlag    = [IO.FileAttributes]::ReparsePoint
$hsFlag    = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System
$files     = New-Object 'System.Collections.Generic.List[object]'
$skipped   = New-Object 'System.Collections.Generic.List[object]'
$skipKinds = [ordered]@{ '権限なし' = 0; 'その他のエラー' = 0; 'クラウドのみ' = 0 }
$OnlineMask = 0x00441000   # クラウドのみのファイルの属性（OFFLINE・RECALL_ON_OPEN・RECALL_ON_DATA_ACCESS）
$linkCount     = 0
$excludedCount = 0

function Add-Skip([string]$Reason, [string]$ItemPath, [string]$Detail) {
    $skipped.Add([object[]]@($Reason, $ItemPath, $Detail))
    if ($skipKinds.Contains($Reason)) { $skipKinds[$Reason]++ }
}

function Add-SkipError([Exception]$Exception, [string]$ItemPath) {
    if ($Exception -is [UnauthorizedAccessException]) { Add-Skip '権限なし' $ItemPath 'アクセスが拒否されました' }
    else { Add-Skip 'その他のエラー' $ItemPath (Get-ErrorText $Exception) }
}

function Get-ExcludeMatch([string]$FullPath) {
    foreach ($rule in $excludeRules) { if ($rule.Matcher.IsMatch($FullPath)) { return $rule.Text } }
    return $null
}

# ファイルを読んで判定し、一覧に加える。読めないときは skipped に残す。
# 中身は持ち続けない（判定だけのときは使わず、変換するときは、書き込む直前にもう一度読むため）
$readBuffer = New-Object 'byte[]' $SniffBytes
function Add-FileInfo($Entry) {
    Stop-IfCtrlC
    $display = ConvertTo-DisplayPath $Entry.FullName
    # OneDrive などのクラウドのみのファイルは、読むとダウンロードが始まるため、調べない
    if (([int]$Entry.Attributes -band $OnlineMask) -ne 0) {
        Add-Skip 'クラウドのみ' $display 'クラウドのみのファイル（読むとダウンロードが始まるため、調べていません）'
        return
    }
    $length  = $Entry.Length
    $partial = $length -gt $LargeBytes
    $bytes   = $null
    $count   = 0
    try {
        # ほかのプログラムが書き込み中のファイルも読めるように開く
        $stream = New-Object IO.FileStream($Entry.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]'ReadWrite, Delete'))
        try {
            if ($partial) { $bytes = $readBuffer } else { $bytes = New-Object 'byte[]' $stream.Length }
            while ($count -lt $bytes.Length) {
                $n = $stream.Read($bytes, $count, $bytes.Length - $count)
                if ($n -le 0) { break }
                $count += $n
            }
        } finally {
            $stream.Dispose()
        }
    } catch {
        Add-SkipError (Get-InnerException $_.Exception) $display
        return
    }
    $info = Get-TextInfo $bytes $count $partial
    $note = if ($partial) { '先頭の 1 MB で判定' } else { '' }
    $files.Add([pscustomobject]@{
        Path = $display; Full = $Entry.FullName; Size = $length; Time = $Entry.LastWriteTime; ReadOnly = $Entry.IsReadOnly
        Name = $info.Name; Line = $info.Line; Partial = $partial; Result = ''; Note = $note
        Planned = $false; Hash = ''; Change = ''; TargetName = ''; Ext = $Entry.Extension.ToLowerInvariant()
    })
}

# 5.1 は、Ctrl+C（パイプラインの停止）を取りこぼして動き続けることがあり、変換の途中ではほとんど止まらない。
# 変換するときは、コンソールで動かしている間、Ctrl+C をキーの入力として受け取り、ファイルの区切りごとに確かめて止める
# （書き込みの途中では止めない）。止めたあとの片付けと結果の保存は、finally で行う
$ctrlCSaved = $null   # 変える前の TreatControlCAsInput（変えていないときは $null）
function Enable-CtrlCInput {
    if (-not $converting -or $null -ne $script:ctrlCSaved) { return }
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

# 調べている途中や変換の途中で Ctrl+C されたときに、どこまで進んだかを残せるようにする
$finished   = $false
$unexpected = $false
$saveError  = $null
$convertStarted = $false
$sw     = [Diagnostics.Stopwatch]::StartNew()
$lines  = New-Object 'System.Collections.Generic.List[string]'

function Save-Result {
    if (-not $runDir) { return }
    try {
        # 保存先は \\?\ 付きで作っているため、長いパスでも書けるよう、ここも \\?\ を付ける
        [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
        $rows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($file in $sortedFiles) {
            $rows.Add([object[]]@($file.Path, $file.Name, $file.Line, $file.Size.ToString('N0', $Invariant), $file.Time.ToString($TimeFormat, $Invariant), $file.Result, $file.Note))
        }
        Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'files.csv')) $FileColumns $rows
        Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $skipped
    } catch {
        $script:saveError = Get-ErrorText (Get-InnerException $_.Exception)
    }
}

$sortedFiles = @()
$answeredNo  = $false
$currentFile = $null
$nextProgress = 0L
try {
    #region 調べる

    Enable-CtrlCInput
    if ($targetIsFile) {
        Add-FileInfo (New-Object IO.FileInfo (ConvertTo-LongPath $target))
    } else {
        # 再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる
        $stack = New-Object 'System.Collections.Generic.Stack[object]'
        $stack.Push((New-Object IO.DirectoryInfo (ConvertTo-LongPath $target)))
        $isRoot = $true
        while ($stack.Count -gt 0) {
            Stop-IfCtrlC
            $dir = $stack.Pop()
            if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                Write-Progress -Activity '文字コード判定・変換' -Status (Format-Invariant 'ファイル {0:N0} 件' $files.Count) -CurrentOperation (ConvertTo-DisplayPath $dir.FullName)
                $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
            }
            try {
                $entries = $dir.GetFileSystemInfos()
            } catch {
                if ($isRoot) { Stop-WithError ('フォルダを読めません: {0}（{1}）' -f $target, (Get-ErrorText (Get-InnerException $_.Exception))) }
                Add-SkipError (Get-InnerException $_.Exception) (ConvertTo-DisplayPath $dir.FullName)
                continue
            }
            $isRoot = $false
            foreach ($entry in $entries) {
                $isDir = $entry -is $dirType
                # 前回の結果（控えたファイル）を調べないよう、このツールの結果のフォルダは中を見ない
                if ($isDir -and $entry.Name.StartsWith($OutputPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
                $linkName = $null
                if (($entry.Attributes -band $rpFlag) -eq $rpFlag) { $linkName = Get-LinkName $entry }
                if ($linkName) {
                    $linkCount++
                    Add-Skip 'リンク' (ConvertTo-DisplayPath $entry.FullName) $linkName
                    continue
                }
                if ($hasExclude) {
                    $matched = Get-ExcludeMatch (ConvertTo-DisplayPath $entry.FullName)
                    if ($matched) {
                        $excludedCount++
                        Add-Skip '除外指定' (ConvertTo-DisplayPath $entry.FullName) $matched
                        continue
                    }
                }
                if ($isDir) {
                    $stack.Push($entry)
                } else {
                    $hit = $false
                    foreach ($rule in $filterRules) { if ($rule.IsMatch($entry.Name)) { $hit = $true; break } }
                    if ($hit) { Add-FileInfo $entry }
                }
            }
        }
    }
    Write-Progress -Activity '文字コード判定・変換' -Completed
    $scanElapsed = $sw.Elapsed

    # 文字コード・改行・パスの順に並べる
    $sortedFiles = @($files | Sort-Object -Property Name, Line, Path)

    #endregion

    #region 判定の結果

    $lines.Add((Format-Invariant '文字コード判定・変換 (Convert-TextEncoding {0})' $ScriptVersion))
    $lines.Add((Format-Invariant '対象: {0}（{1}）   {2:yyyy-MM-dd HH:mm:ss}   {3}' $target $filterLabel $startTime $machineName))
    if ($hasExclude) { $lines.Add(('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', '))) }
    $lines.Add('')

    # 内訳は件数の多い順
    $groups = @($sortedFiles | Group-Object -Property Name, Line | Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, Name)
    $widthName = 10; $widthLine = 4
    foreach ($file in $sortedFiles) {
        $widthName = [math]::Max($widthName, (Get-DisplayWidth $file.Name))
        $widthLine = [math]::Max($widthLine, (Get-DisplayWidth $(if ($file.Line) { $file.Line } else { '-' })))
    }
    $lines.Add('■ 文字コードと改行の内訳')
    if ($groups.Count -eq 0) {
        $lines.Add('    （対象のファイルはありません）')
    } else {
        $lines.Add(' ' + (Format-Column '文字コード' $widthName $false) + '  ' + (Format-Column '改行' $widthLine $false) + '  ' + (Format-Column '件数' 6 $true))
        foreach ($group in $groups) {
            $first = $group.Group[0]
            $lineText = if ($first.Line) { $first.Line } else { '-' }
            $lines.Add(' ' + (Format-Column $first.Name $widthName $false) + '  ' + (Format-Column $lineText $widthLine $false) + '  ' + (Format-Column ($group.Count.ToString('N0', $Invariant)) 6 $true))
        }
    }
    $lines.Add('')
    $listIndex = $lines.Count
    $lines.Add('■ ファイル（文字コードの順）')
    $fileHeader = ' ' + (Format-Column '文字コード' $widthName $false) + '  ' + (Format-Column '改行' $widthLine $false) + '  ' + (Format-Column 'サイズ' 10 $true) + '  ファイル'
    $fileLines = @(foreach ($file in $sortedFiles) {
        ' ' + (Format-Column $file.Name $widthName $false) + '  ' + (Format-Column $(if ($file.Line) { $file.Line } else { '-' }) $widthLine $false) + '  ' + (Format-Column (Format-Size $file.Size) 10 $true) + '  ' + $file.Path
    })
    if ($fileLines.Count -eq 0) {
        $lines.Add('    （対象のファイルはありません）')
    } else {
        $lines.Add($fileHeader)
        foreach ($line in $fileLines) { $lines.Add($line) }
    }
    $listEnd = $lines.Count
    $lines.Add('')
    $lines.Add((Format-Invariant 'ファイル {0:N0} 件  所要 {1}' $sortedFiles.Count (Format-Duration $scanElapsed)))
    $notes = @()
    $kinds = @(foreach ($name in $skipKinds.Keys) { if ($skipKinds[$name] -gt 0) { , @($name, $skipKinds[$name]) } })
    if ($kinds.Count -eq 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $kinds[0][1] $kinds[0][0]
    } elseif ($kinds.Count -gt 1) {
        $skipTotal = 0
        foreach ($kind in $kinds) { $skipTotal += $kind[1] }
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $skipTotal (($kinds | ForEach-Object { Format-Invariant '{0} {1:N0}' $_[0] $_[1] }) -join '・')
    }
    if ($linkCount -gt 0) { $notes += Format-Invariant 'たどらなかったリンク {0:N0} 件' $linkCount }
    if ($excludedCount -gt 0) { $notes += Format-Invariant '除外 {0:N0} 件' $excludedCount }
    if ($notes.Count -gt 0) { $lines.Add($notes -join '  ') }

    # 画面に出す。保存するときは、ファイルの一覧は先頭の分だけ（保存しないときは、ほかに見る場所がないためすべて）
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($i -gt $listIndex + 1 + $ScreenRows -and $i -lt $listEnd -and $runDir) {
            if ($i -eq $listIndex + 2 + $ScreenRows) { Write-Host (Format-Invariant '（ほか {0:N0} 件は files.csv を参照してください）' ($listEnd - $i)) }
            continue
        }
        if ($line.StartsWith('■')) { Write-Host $line -ForegroundColor Cyan } else { Write-Host $line }
    }

    #endregion

    #region 変換

    if ($converting) {
        # 変換の予定を立てる。中身はファイルごとに読み直し、予定を立てたら手放す（すべてをメモリに持たないため）
        $plan = New-Object 'System.Collections.Generic.List[object]'
        $sameCount = 0; $outCount = 0; $cannotCount = 0
        $ps1Count = 0; $batCount = 0
        $changeKinds = [ordered]@{}
        foreach ($file in $sortedFiles) {
            Stop-IfCtrlC
            if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                Write-Progress -Activity '文字コード判定・変換' -Status '変換の予定を立てています' -CurrentOperation $file.Path
                $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
            }
            if ($file.Partial -or -not ($Convertible -contains $file.Name)) {
                $file.Result = '対象外'
                if (-not $file.Note) { $file.Note = '{0} は変換しません' -f $file.Name }
                $outCount++
                continue
            }
            try {
                $stream = New-Object IO.FileStream($file.Full, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]'ReadWrite, Delete'))
                try { $bytes = Read-StreamBytes $stream } finally { $stream.Dispose() }
            } catch {
                $file.Result = '変換できない'
                $file.Note = '読めません: ' + (Get-ErrorText (Get-InnerException $_.Exception))
                $cannotCount++
                continue
            }
            try {
                $conv = Get-Conversion $bytes
            } catch {
                $file.Result = '変換できない'
                $file.Note = '変換の準備でエラーになりました: ' + (Get-ErrorText (Get-InnerException $_.Exception))
                $cannotCount++
                continue
            }
            switch ($conv.Status) {
                '対象外' { $file.Result = '対象外'; $file.Note = $conv.Note; $outCount++ }
                '変換できない' { $file.Result = '変換できない'; $file.Note = $conv.Note; $cannotCount++ }
                'すでに同じ' { $file.Result = 'すでに同じ'; $sameCount++ }
                default {
                    if ($file.ReadOnly) {
                        $file.Result = '変換できない'
                        $file.Note = '読み取り専用のファイル'
                        $cannotCount++
                    } else {
                        $file.Planned = $true
                        $file.Hash = Get-BytesHash $bytes
                        $file.Change = $conv.Change
                        $file.TargetName = $conv.Target.Name
                        $plan.Add($file)
                        foreach ($part in ($conv.Change -split '・')) { if ($changeKinds.Contains($part)) { $changeKinds[$part]++ } else { $changeKinds[$part] = 1 } }
                        # よくある落とし穴を知らせる（日本語などを含み、文字コードを変えるものだけ）
                        if ($conv.EncodingChanged -and $conv.Info.Name -ne $NameAscii) {
                            if ($conv.Target.Name -eq $NameUtf8 -and @('.ps1', '.psm1', '.psd1') -contains $file.Ext) { $ps1Count++ }
                            if (($conv.Target.Name -eq $NameUtf8 -or $conv.Target.Name -eq $NameUtf8Bom) -and @('.bat', '.cmd') -contains $file.Ext) { $batCount++ }
                        }
                    }
                }
            }
            $bytes = $null
            $conv  = $null
        }
        Write-Progress -Activity '文字コード判定・変換' -Completed

        $targetLabel = @()
        if ($To) { $targetLabel += (Get-TargetEncoding $To $null).Name }
        if ($LineEnding) { $targetLabel += $LineEnding }
        $planLines = New-Object 'System.Collections.Generic.List[string]'
        $planLines.Add('')
        $planLines.Add(('■ 変換の予定（{0} にそろえる）' -f ($targetLabel -join '・')))
        $planLines.Add((Format-Invariant '変換するファイル {0:N0} 件' $plan.Count))
        $widthChange = 4
        foreach ($key in $changeKinds.Keys) { $widthChange = [math]::Max($widthChange, (Get-DisplayWidth $key)) }
        foreach ($key in $changeKinds.Keys) { $planLines.Add('  ' + (Format-Column $key $widthChange $false) + '  ' + (Format-Invariant '{0:N0} 件' $changeKinds[$key])) }
        $restParts = @()
        if ($sameCount -gt 0) { $restParts += Format-Invariant 'すでに同じ {0:N0} 件' $sameCount }
        if ($outCount -gt 0) { $restParts += Format-Invariant '対象外 {0:N0} 件' $outCount }
        if ($cannotCount -gt 0) { $restParts += Format-Invariant '変換できない {0:N0} 件' $cannotCount }
        $restCount = $sameCount + $outCount + $cannotCount
        if ($restCount -gt 0) { $planLines.Add((Format-Invariant '変換しないファイル {0:N0} 件（{1}）' $restCount ($restParts -join '・'))) }
        if ($ps1Count -gt 0) { $planLines.Add((Format-Invariant '※ .ps1 を BOM なしの UTF-8 にすると、Windows PowerShell 5.1 で日本語が化けます（{0:N0} 件）' $ps1Count)) }
        if ($batCount -gt 0) { $planLines.Add((Format-Invariant '※ .bat・.cmd を UTF-8 にすると、コマンドプロンプトで日本語が化けます（{0:N0} 件）' $batCount)) }
        foreach ($line in $planLines) {
            $lines.Add($line)
            if ($line.StartsWith('■')) { Write-Host $line -ForegroundColor Cyan } elseif ($line.StartsWith('※')) { Write-Host $line -ForegroundColor Yellow } else { Write-Host $line }
        }

        if ($plan.Count -eq 0) {
            # 変換できないファイルがあるときは、「そろっている」とは言わない
            $message = if ($cannotCount -gt 0) { '変換するファイルはありません' } else { '変換するファイルはありません（すべてそろっています）' }
            $lines.Add($message)
            Write-Host $message
        } else {
            $message = '元のファイルは、保存先の backup フォルダに控えます'
            $lines.Add($message)
            Write-Host $message
            $answer = 'Y'
            if (-not $Force) {
                # 画面のない実行（タスクスケジューラの「ログオンしているかどうかにかかわらず実行する」など）で、確認待ちのまま止まり続けないようにする
                $cannotAsk = '確認できないため、変換しません。確認せずに変換するときは -Force を付けてください'
                if (-not [Environment]::UserInteractive) { Stop-WithError $cannotAsk }
                # 確認を待つ間は、通常の Ctrl+C に戻す（キーの入力として受け取ると、Read-Host に Ctrl+C の文字が入るため）
                Stop-IfCtrlC
                Disable-CtrlCInput
                $answer = ''
                while ($answer -notmatch '^[YyNnＹｙＮｎ]') {
                    # -NonInteractive で起動されていると、Read-Host は英語のエラーで止まる
                    try { $answer = Read-Host '変換しますか？ [Y] はい  [N] いいえ' } catch { Stop-WithError $cannotAsk }
                    if ($null -eq $answer) { Stop-WithError $cannotAsk }   # 入力が終わっている（リダイレクトされた入力の終わり）
                    $answer = $answer.Trim()
                }
            }
            if ($answer -match '^[NnＮｎ]') {
                $answeredNo = $true
                foreach ($file in $plan) { $file.Result = '変換しなかった'; $file.Note = $file.Change }
                $lines.Add('変換しませんでした')
                Write-Host '変換しませんでした'
            } else {
                Enable-CtrlCInput
                $convertStarted = $true
                $convertWatch = [Diagnostics.Stopwatch]::StartNew()
                $doneCount = 0; $failCount = 0
                foreach ($file in $plan) {
                    # 押されていたら、ここ（前のファイルを書き終えた後）で止める
                    Stop-IfCtrlC
                    if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                        Write-Progress -Activity '文字コード判定・変換' -Status (Format-Invariant '変換しています: {0:N0} / {1:N0} 件' ($doneCount + $failCount) $plan.Count) -CurrentOperation $file.Path
                        $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
                    }
                    # ほかのプログラムが書き込めないように開いたまま、読む・控える・書く・読み直すまでを行う
                    try {
                        $stream = New-Object IO.FileStream($file.Full, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
                    } catch {
                        $file.Result = '失敗'
                        $file.Note = '開けないため、変換していません: ' + (Get-ErrorText (Get-InnerException $_.Exception))
                        $failCount++
                        continue
                    }
                    try {
                        # 1 件読めなくても、ほかのファイルの変換は続ける
                        try {
                            $bytes = Read-StreamBytes $stream
                            $sameAsPlan = (Get-BytesHash $bytes) -eq $file.Hash
                            if ($sameAsPlan) { $conv = Get-Conversion $bytes }
                        } catch {
                            $file.Result = '失敗'
                            $file.Note = '読めないため、変換していません: ' + (Get-ErrorText (Get-InnerException $_.Exception))
                            $failCount++
                            continue
                        }
                        # 予定を立てた後（確認を待っている間など）に変わったファイルは、追記した分などを消さないよう、書かない
                        if (-not $sameAsPlan) {
                            $file.Result = '失敗'
                            $file.Note = '調べた後に変更されたため、変換していません'
                            $failCount++
                            continue
                        }
                        # 元のファイルを控える。控えられなければ変換しない
                        $relative = $file.Path.Substring([math]::Min($basePrefixLen, $file.Path.Length))
                        if ($targetIsFile) { $relative = Split-Path -Leaf $file.Path }
                        $backupPath = Join-Path (Join-Path $runDir 'backup') $relative
                        try {
                            [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath (Split-Path -LiteralPath $backupPath)))
                            [IO.File]::WriteAllBytes((ConvertTo-LongPath $backupPath), $bytes)
                        } catch {
                            $file.Result = '失敗'
                            $file.Note = '元のファイルを控えられないため、変換していません: ' + (Get-ErrorText (Get-InnerException $_.Exception))
                            $failCount++
                            continue
                        }
                        # 書き込みの途中で止まったときに、そのファイルを「失敗」と記録できるようにしておく
                        $currentFile = $file
                        try {
                            $stream.Position = 0
                            $stream.Write($conv.NewBytes, 0, $conv.NewBytes.Length)
                            $stream.SetLength($conv.NewBytes.Length)
                            $stream.Flush()
                            # 読み直して、書いた内容と、文字に直した内容が変わっていないかを確かめる
                            $stream.Position = 0
                            $check = Read-StreamBytes $stream
                            $same = Test-SameBytes $check $conv.NewBytes
                            if ($same -and $null -ne $conv.NewText) {
                                $preambleLength = $conv.Target.Preamble.Length
                                $same = [string]::Equals($conv.Target.Encoding.GetString($check, $preambleLength, $check.Length - $preambleLength), $conv.NewText, [StringComparison]::Ordinal)
                            }
                            if (-not $same) { throw '読み直した内容が、変換した内容と一致しません' }
                            $file.Result = '変換した'
                            $file.Note = $file.Change
                            $doneCount++
                        } catch {
                            $reason = if ($_.Exception -is [Management.Automation.RuntimeException] -and -not $_.Exception.InnerException) { $_.Exception.Message } else { Get-ErrorText (Get-InnerException $_.Exception) }
                            # 元の内容に戻す
                            $restored = '（元に戻しました）'
                            try {
                                $stream.Position = 0
                                $stream.Write($bytes, 0, $bytes.Length)
                                $stream.SetLength($bytes.Length)
                                $stream.Flush()
                            } catch {
                                $restored = '（元に戻せませんでした。backup のファイルから戻してください）'
                            }
                            $file.Result = '失敗'
                            $file.Note = $reason + $restored
                            $failCount++
                        }
                        $currentFile = $null
                    } finally {
                        $stream.Dispose()
                    }
                    # 変換した後のサイズと最終更新日時を、一覧に書く
                    if ($file.Result -eq '変換した') {
                        try {
                            $newInfo = New-Object IO.FileInfo $file.Full
                            $file.Size = $newInfo.Length
                            $file.Time = $newInfo.LastWriteTime
                        } catch { }
                    }
                }
                Write-Progress -Activity '文字コード判定・変換' -Completed
                $resultLine = if ($failCount -eq 0) { '結果: 変換しました  {0:N0} 件  失敗 0 件  所要 {1}' } else { '結果: 変換できなかったファイルがあります  変換 {0:N0} 件  失敗 {2:N0} 件  所要 {1}（files.csv を確認してください）' }
                $resultLine = Format-Invariant $resultLine $doneCount (Format-Duration $convertWatch.Elapsed) $failCount
                $lines.Add('')
                $lines.Add($resultLine)
                Write-Host ''
                Write-Host $resultLine
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
    Write-Progress -Activity '文字コード判定・変換' -Completed
    if (-not $finished -and $stopping) {
        # 止めた理由は、Stop-WithError が表示している
        if ($runDir -and -not $convertStarted) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
    } elseif (-not $finished) {
        if ($convertStarted) {
            # 変換の途中で止まったときは、変換したファイルと控えが残っているため、結果も残す
            # 書き込みの途中だったファイルは、壊れているおそれがあるため「失敗」にする
            if ($currentFile) {
                $currentFile.Result = '失敗'
                $currentFile.Note = '書き込みの途中で止まりました。backup のファイルから戻してください'
            }
            foreach ($file in $sortedFiles) { if ($file.Result -eq '' -and $file.Planned) { $file.Result = '変換しなかった'; $file.Note = '中断したため、変換していません' } }
            $stopText = if ($unexpected) { 'エラーで止まりました' } else { '中断しました' }
            $lines.Add(('{0}。それまでに変換したファイルと、控えたファイルは残っています' -f $stopText))
            Save-Result
            Write-Host ('{0}。それまでに変換したファイルと、控えたファイルは残っています: {1}' -f $stopText, $runDir) -ForegroundColor Yellow
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

# 確認で N を選んだときは、変換できないファイルがあっても 0
$failed = @($sortedFiles | Where-Object { $_.Result -eq '変換できない' -or $_.Result -eq '失敗' }).Count
if ($converting -and -not $answeredNo -and $failed -gt 0) { exit 1 }
exit 0
