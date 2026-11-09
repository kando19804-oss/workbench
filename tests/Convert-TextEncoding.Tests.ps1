<#
.SYNOPSIS
    Convert-TextEncoding の回帰テストです。

.DESCRIPTION
    一時フォルダに、文字コードと改行の違うダミーのファイル（UTF-8 の BOM あり／なし、UTF-16 の LE と BE、Shift_JIS、
    ASCII、EUC-JP、JIS、バイナリ、壊れたもの、絵文字を含むもの、改行の混在、読み取り専用・隠しファイル、
    260 文字を超えるパス、名前に [ ] を含むフォルダ、ジャンクションのループ、読み取り禁止のフォルダとファイル、
    100 MB を超えるファイルなど）を作り、Convert-TextEncoding で判定と変換を行って、期待値と照合します。
    変換したファイルは、バイト単位で期待値と比べます。

    「権限なし」を再現するため、テストが作ったフォルダとファイルにだけアクセス拒否を設定します（終了時に戻します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境（100 MB を超えるファイル 2 つを含む）は、終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Convert-TextEncoding を実行する PowerShell（powershell.exe または pwsh）。

.PARAMETER KeepTemp
    ダミー環境と結果を削除せずに残します（調査用）。

.NOTES
    Copyright (c) 2026 結進舎
    Released under the MIT License.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Exe,

    [switch]$KeepTemp
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Continue'

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Convert-TextEncoding\Convert-TextEncoding.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Convert-TextEncoding_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Convert-TextEncoding.ps1'
$det       = Join-Path $workRoot 'det'      # 判定だけに使う（変換しない）
$conv      = Join-Path $workRoot 'conv'     # 変換に使う（変換のたびに作り直す）
$many      = Join-Path $workRoot 'many'     # 200 件の打ち切りと Ctrl+C に使う
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
$sjis      = [Text.Encoding]::GetEncoding(932)
$le        = [Text.Encoding]::Unicode
$be        = [Text.Encoding]::BigEndianUnicode
$bom       = [byte[]](0xEF, 0xBB, 0xBF)
$fileHead  = '"ファイル","文字コード","改行","サイズ(バイト)","最終更新日時","結果","備考"'
$oldTime   = New-Object DateTime 2020, 1, 2, 3, 4, 5
$script:failCount = 0

#region 補助関数

function Assert-Equal([string]$Name, $Actual, $Expected) {
    if ("$Actual" -ceq "$Expected") {
        '  PASS  {0} = {1}' -f $Name, $Actual
    } else {
        '  FAIL  {0} : 実際={1} 期待={2}' -f $Name, $Actual, $Expected
        $script:failCount++
    }
}

function Write-Skip([string]$Name, [string]$Reason) { '  SKIP  {0}（{1}）' -f $Name, $Reason }

function Write-Section([string]$Title) { ''; ('■ ' + $Title) }

# コンソールで実行すると、長い行は画面の幅で折り返され、出力に改行が入る。
# 画面の幅によって結果が変わらないように、文字を探す前に改行を取り除く
function ConvertTo-FlatText([string]$Text) { $Text -replace '\r?\n', '' }

# バイト列をつなぐ
function Join-Bytes {
    $list = New-Object 'System.Collections.Generic.List[byte]'
    foreach ($part in $args) { $list.AddRange([byte[]]$part) }
    , $list.ToArray()
}

function ConvertTo-Hex([byte[]]$Bytes) { if ($null -eq $Bytes) { return '(なし)' }; [BitConverter]::ToString($Bytes) }

# 260 文字を超えるパスにも書けるように、\\?\ を付けて書く
function New-TestFile([string]$FilePath, [byte[]]$Bytes) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllBytes('\\?\' + $FilePath, $Bytes)
}

function Get-Bytes([string]$FilePath) {
    if (-not [IO.File]::Exists('\\?\' + $FilePath)) { return $null }
    , [IO.File]::ReadAllBytes('\\?\' + $FilePath)
}

function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$InputText = $null) {
    $out = Join-Path $workRoot "out\$Label"
    if ($null -ne $InputText) {
        $lines = @($InputText | & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }
    } else {
        $lines = @(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }
    }
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = (ConvertTo-FlatText (@($lines) -join "`n")); Lines = @($lines); Dir = (Get-RunDir $out) }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行。変換で確認を待つときに止まらないよう、標準入力は空にする
function Invoke-Plain([string[]]$ToolArgs) {
    $log = ConvertTo-FlatText ((@('' | & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 相対パスの基準を確かめるため、-OutputFolder を足さず、指定したフォルダから起動する
function Invoke-From([string]$WorkDir, [string[]]$ToolArgs) {
    Push-Location -LiteralPath $WorkDir
    try {
        $log  = ConvertTo-FlatText ((@('' | & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
        $code = $LASTEXITCODE
    } finally {
        Pop-Location
    }
    [pscustomobject]@{ Code = $code; Log = $log }
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する（標準入力は NUL）
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { $utf8NoBom } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', ($CommandLine + ' < NUL')), $encoding)
    $log = ConvertTo-FlatText ((@(& cmd.exe /d /c $batch 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# [powershell]::Create() で動かすスクリプトを、一時フォルダに書く（Ctrl+C の代わりに Stop() で止めるため）
function New-HelperScript([string]$Name, [string[]]$Lines) {
    $file = Join-Path $workRoot $Name
    [IO.File]::WriteAllLines($file, $Lines, $utf8Bom)
    $file
}

# 補助スクリプトの「名前=値」の出力を受け取る
function Invoke-Helper([string]$Script, [string[]]$HelperArgs) {
    $result = @{}
    & $Exe -NoProfile -ExecutionPolicy Bypass -File $Script @HelperArgs 2>&1 | ForEach-Object {
        $line = "$_"
        $eq = $line.IndexOf('=')
        if ($eq -gt 0) { $result[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
    }
    $result
}

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'TextEncoding_*' -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($found.Count -gt 0) { return $found[-1].FullName }
    $null
}

# 結果のファイルがないとき（ツールが保存に失敗したときなど）も止まらないように、空として読む
function Read-Result([string]$Dir) {
    function Read-CsvOrEmpty([string]$Name) {
        if (-not $Dir) { return @() }
        $file = Join-Path $Dir $Name
        if (-not (Test-Path -LiteralPath $file)) { return @() }
        @(Import-Csv -LiteralPath $file -Encoding UTF8)
    }
    $lines = @()
    if ($Dir -and (Test-Path -LiteralPath (Join-Path $Dir 'summary.txt'))) { $lines = @([IO.File]::ReadAllLines((Join-Path $Dir 'summary.txt'))) }
    [pscustomobject]@{ Files = @(Read-CsvOrEmpty 'files.csv'); Skipped = @(Read-CsvOrEmpty 'skipped.csv'); Lines = $lines }
}

function Join-Sorted($Items) {
    $array = [string[]]@($Items | Where-Object { $null -ne $_ })
    [Array]::Sort($array, [StringComparer]::Ordinal)
    $array -join ','
}

# 保存されたファイルとフォルダの名前の一覧
function Get-ResultNames([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir | ForEach-Object { $_.Name })
}

function Get-BomFlags([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir -File | ForEach-Object {
        $bytes = [IO.File]::ReadAllBytes($_.FullName)
        '{0}={1}' -f $_.Name, ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    })
}

function Get-FirstLine([string]$Dir, [string]$FileName) {
    if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir $FileName))) { return '(なし)' }
    @([IO.File]::ReadAllLines((Join-Path $Dir $FileName)))[0]
}

# files.csv の、ダミーのフォルダから見たファイルの行の列
function Get-Field($Rows, [string]$Base, [string]$Rel, [string]$Column) {
    $full = Join-Path $Base $Rel
    (@($Rows | Where-Object { $_.'ファイル' -eq $full }) | ForEach-Object { $_.$Column }) -join '|'
}

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $used -notcontains $_ -and -not [IO.Directory]::Exists("${_}:\") } | Select-Object -First 1
}

# 自分自身にアクセス拒否を設定する。終了時に戻すため、設定したパスを覚えておく
function Add-Deny([string]$Path, [string]$Rights) {
    $null = & icacls.exe $Path /deny "*${userSid}:$Rights" 2>&1
    if ($LASTEXITCODE -ne 0) { return $false }
    $denied.Add($Path)
    $true
}

# 変換に使うダミーを、元の内容・属性・最終更新日時で作り直す（変換のテストのたびに、同じ状態から始めるため）
function Reset-Conv {
    foreach ($rel in $convFiles.Keys) {
        $full = '\\?\' + (Join-Path $conv $rel)
        if ([IO.File]::Exists($full)) { [IO.File]::SetAttributes($full, [IO.FileAttributes]::Normal) }
    }
    foreach ($rel in $convFiles.Keys) {
        $full = Join-Path $conv $rel
        New-TestFile $full $convFiles[$rel]
        [IO.File]::SetLastWriteTime('\\?\' + $full, $oldTime)
    }
    foreach ($rel in $convAttributes.Keys) { [IO.File]::SetAttributes('\\?\' + (Join-Path $conv $rel), $convAttributes[$rel]) }
}

# 変換の結果のファイルが、期待したバイト列か（期待値が $null なら、元のまま）
function Test-ConvBytes([string]$Rel, $Expected) {
    if ($null -eq $Expected) { $Expected = $convFiles[$Rel] }
    (ConvertTo-Hex (Get-Bytes (Join-Path $conv $Rel))) -eq (ConvertTo-Hex $Expected)
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
# 標準入力に渡す文字は、コンソールの入力の文字コードで書かれる。UTF-8 のとき先頭に BOM が付き、7 の Read-Host は
# それを外さないため、BOM を付けない UTF-8 にする。全角の Ｙ を渡せるよう、パイプの文字コードも UTF-8 にする
# （外部のプログラムに渡すときはグローバルの $OutputEncoding が使われるため、tests\Invoke-Tests.ps1 から呼ばれたときも効くように global にする）
$savedInputEncoding = $null
try { $savedInputEncoding = [Console]::InputEncoding; [Console]::InputEncoding = $utf8NoBom } catch { }
$savedOutputEncoding = $global:OutputEncoding
$global:OutputEncoding = $utf8NoBom
$userSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied  = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したパス（終了時に戻す）
$holds   = New-Object 'System.Collections.Generic.List[object]'   # テストの中で開いたままにしたファイル（終了時に閉じる）
$longRel = ('L' * 120) + '\' + ('M' * 120) + '\long.txt'

# 判定に使うダミー（相対パス → 中身と、期待する「文字コード/改行」）
$detFiles = [ordered]@{
    'utf8bom.txt'     = @((Join-Bytes $bom $utf8NoBom.GetBytes("日本語`r`nです`r`n")), 'UTF-8（BOM あり）/CRLF')
    'utf8bom-bad.txt' = @((Join-Bytes $bom ([byte[]](0xC3, 0x28))), '判定できない/')
    'utf8.txt'        = @($utf8NoBom.GetBytes("日本語`nです`n"), 'UTF-8（BOM なし）/LF')
    'short.txt'       = @([byte[]](0xC3, 0xA9, 0x0A), 'UTF-8（BOM なし）/LF')                      # Shift_JIS としても正しいが、UTF-8 を優先する
    'utf16le.txt'     = @((Join-Bytes ([byte[]](0xFF, 0xFE)) $le.GetBytes("日本語`r`n")), 'UTF-16（BOM あり）/CRLF')
    'utf16be.txt'     = @((Join-Bytes ([byte[]](0xFE, 0xFF)) $be.GetBytes("日本語`n")), 'UTF-16（BOM あり）/LF')
    'utf16-bad.txt'   = @([byte[]](0xFF, 0xFE, 0x00, 0xD8), '判定できない/')                       # 対になっていないサロゲート
    'utf16-nobom.txt' = @($le.GetBytes('abc'), 'バイナリ/')                                         # BOM のない UTF-16 は NUL を含む
    'bin.txt'         = @([byte[]](0x41, 0x00, 0x42), 'バイナリ/')
    'nul-late.txt'    = @((Join-Bytes $utf8NoBom.GetBytes('a' * 9000) ([byte[]](0x00, 0x62))), 'ASCII のみ/なし')   # NUL が先頭の 8 KB より後ろ
    'jis.txt'         = @([byte[]](0x1B, 0x24, 0x42, 0x46, 0x7C, 0x4B, 0x5C, 0x38, 0x6C, 0x1B, 0x28, 0x42, 0x0D, 0x0A), 'JIS/CRLF')
    'ascii.csv'       = @($utf8NoBom.GetBytes("a,b`r`n1,2`r`n"), 'ASCII のみ/CRLF')
    'mixed.txt'       = @($utf8NoBom.GetBytes("a`r`nb`nc`r"), 'ASCII のみ/混在')
    'cr.txt'          = @($utf8NoBom.GetBytes("a`rb`r"), 'ASCII のみ/CR')
    'noline.txt'      = @($utf8NoBom.GetBytes('abc'), 'ASCII のみ/なし')
    'empty.txt'       = @([byte[]]@(), 'ASCII のみ/なし')
    'sjis.bat'        = @($sjis.GetBytes("@echo off`r`necho 日本語`r`n"), 'Shift_JIS/CRLF')
    'sjis-kana.txt'   = @($sjis.GetBytes("ｶﾅと漢字`n"), 'Shift_JIS/LF')
    'euc.txt'         = @([byte[]](0xA4, 0xB3, 0xA4, 0xEC, 0xA4, 0xCF, 0xC6, 0xFC, 0xCB, 0xDC, 0xB8, 0xEC, 0xA4, 0xCE, 0xCA, 0xB8, 0xBE, 0xCF, 0xA4, 0xC7, 0xA4, 0xB9, 0x0A), 'EUC-JP/LF')   # 「これは日本語の文章です」
    # EUC-JP の「アイウ」。Shift_JIS としても正しい（半角カナ 6 文字になる）ため、半角カナばかりかどうかで EUC-JP と判定する
    'euc-kana.txt'    = @([byte[]](0xA5, 0xA2, 0xA5, 0xA4, 0xA5, 0xA6, 0x0A), 'EUC-JP/LF')
    'unknown.txt'     = @([byte[]](0x81, 0x0A), '判定できない/')
    # Shift_JIS の中で同じ文字が 2 か所にある字（∵ の NEC 特殊文字 87 9A）。厳密に読むと、5.1 でも 7 でも読めない
    'nec.txt'         = @([byte[]](0x87, 0x9A, 0x0A), '判定できない/')
    'sub\run.ps1'     = @($sjis.GetBytes("Write-Host '日本語'`r`n"), 'Shift_JIS/CRLF')
    '日本語 フォルダ\メモ.txt' = @($utf8NoBom.GetBytes("メモ`n"), 'UTF-8（BOM なし）/LF')
    '[a]\f[1].txt'    = @($utf8NoBom.GetBytes("[x]`n"), 'ASCII のみ/LF')
    'hidden.txt'      = @($utf8NoBom.GetBytes("隠し`n"), 'UTF-8（BOM なし）/LF')
    $longRel          = @($utf8NoBom.GetBytes("長いパス`n"), 'UTF-8（BOM なし）/LF')
}

# 変換に使うダミー（相対パス → 中身）
$convFiles = [ordered]@{
    'u8.txt'       = $utf8NoBom.GetBytes("日本語`nです`n")
    'u8bom.txt'    = Join-Bytes $bom $utf8NoBom.GetBytes("日本語`r`nです`r`n")
    'sj.txt'       = $sjis.GetBytes("日本語`r`n")
    'u16le.txt'    = Join-Bytes ([byte[]](0xFF, 0xFE)) $le.GetBytes("日本語`r`n")
    'u16be.txt'    = Join-Bytes ([byte[]](0xFE, 0xFF)) $be.GetBytes("日本語`n")
    'ascii.txt'    = $utf8NoBom.GetBytes("abc`n")
    'mixed.txt'    = $utf8NoBom.GetBytes("a`r`nb`nc`r")
    'emoji.txt'    = Join-Bytes $utf8NoBom.GetBytes("line1`nスマイル ") ([byte[]](0xF0, 0x9F, 0x99, 0x82)) $utf8NoBom.GetBytes("`n")
    'euc.txt'      = [byte[]](0xA4, 0xB3, 0xA4, 0xEC, 0xA4, 0xCF, 0xC6, 0xFC, 0xCB, 0xDC, 0xB8, 0xEC, 0xA4, 0xCE, 0xCA, 0xB8, 0xBE, 0xCF, 0xA4, 0xC7, 0xA4, 0xB9, 0x0A)
    'bin.txt'      = [byte[]](0x41, 0x00, 0x42)
    'readonly.txt' = $utf8NoBom.GetBytes("読み取り専用`n")
    'hidden.txt'   = $utf8NoBom.GetBytes("隠し`n")
    'acl.txt'      = $utf8NoBom.GetBytes("アクセス権`n")
    'run.ps1'      = $sjis.GetBytes("Write-Host '日本語'`r`n")
    'run.bat'      = $sjis.GetBytes("@echo off`r`necho 日本語`r`n")
    '[b]\g.txt'    = $utf8NoBom.GetBytes("かっこ`n")
    'sub\deep.txt' = $utf8NoBom.GetBytes("サブ`n")
    $longRel       = $utf8NoBom.GetBytes("長いパス`n")
}
$convAttributes = @{ 'readonly.txt' = [IO.FileAttributes]::ReadOnly; 'hidden.txt' = [IO.FileAttributes]::Hidden }

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool
    foreach ($rel in $detFiles.Keys) { New-TestFile (Join-Path $det $rel) $detFiles[$rel][0] }
    [IO.File]::SetAttributes((Join-Path $det 'hidden.txt'), [IO.FileAttributes]::Hidden)
    New-TestFile (Join-Path $det 'skip.dat') $utf8NoBom.GetBytes("filter で対象外`n")
    New-TestFile (Join-Path $det 'TextEncoding_old\x.txt') $utf8NoBom.GetBytes("前回の結果`n")
    New-TestFile (Join-Path $det 'denied\secret.txt') $utf8NoBom.GetBytes("秘密`n")
    New-TestFile (Join-Path $det 'noread.txt') $utf8NoBom.GetBytes("読めない`n")
    # 100 MB を超えるファイル。先頭の 1 MB だけで判定する。big-utf8.log は、1 MB の境目で文字（3 バイト）が切れる
    $bigHead = @{ 'big.log' = $utf8NoBom.GetBytes(('x' * 9215) + "`n"); 'big-utf8.log' = $utf8NoBom.GetBytes('あ' * 400000) }
    foreach ($name in $bigHead.Keys) {
        $stream = [IO.File]::Create((Join-Path $det $name))
        try { $stream.Write($bigHead[$name], 0, $bigHead[$name].Length); $stream.SetLength(101MB) } finally { $stream.Dispose() }
    }
    $null = & cmd.exe /d /c mklink /J "$det\jloop" "$det" 2>&1
    $hasJunction = [IO.Directory]::Exists("$det\jloop")
    $deniedOk = (Add-Deny "$det\denied" '(RD)') -and (Add-Deny "$det\noread.txt" '(R)')
    Assert-Equal 'ジャンクションとアクセス拒否を作れた' ('{0}/{1}' -f $hasJunction, $deniedOk) 'True/True'

    Reset-Conv
    # アクセス権を変えないことを確かめるため、Users に読み取りを足しておく
    $null = & icacls.exe (Join-Path $conv 'acl.txt') /grant '*S-1-5-32-545:(R)' 2>&1
    $aclBefore = (Get-Acl -LiteralPath (Join-Path $conv 'acl.txt')).Sddl

    # 200 件の打ち切りと Ctrl+C に使う、小さなファイル 1,500 件
    $manyText = $utf8NoBom.GetBytes((1..60 | ForEach-Object { '行 {0:00} のテキストです abcdefghijklmnopqrstuvwxyz' -f $_ }) -join "`n")
    for ($i = 1; $i -le 1500; $i++) { New-TestFile (Join-Path $many ('f{0:0000}.txt' -f $i)) $manyText }

    Write-Section '判定（既定の拡張子）'
    $run = Invoke-Tool 'detect' @('-Path', $det)
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '保存するファイル' (Get-ResultNames $run.Dir) 'files.csv,skipped.csv,summary.txt'
    Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'files.csv=True,skipped.csv=True,summary.txt=True'
    Assert-Equal 'files.csv の見出し' (Get-FirstLine $run.Dir 'files.csv') $fileHead
    Assert-Equal 'skipped.csv の見出し' (Get-FirstLine $run.Dir 'skipped.csv') '"理由","パス","詳細"'
    foreach ($rel in $detFiles.Keys) {
        Assert-Equal ('判定: {0}' -f $rel.Substring(0, [math]::Min(40, $rel.Length))) ('{0}/{1}' -f (Get-Field $x.Files $det $rel '文字コード'), (Get-Field $x.Files $det $rel '改行')) $detFiles[$rel][1]
    }
    Assert-Equal '大きなファイルは先頭の 1 MB で判定（文字が切れていても UTF-8）' ((@('big.log', 'big-utf8.log') | ForEach-Object { '{0}/{1}/{2}' -f (Get-Field $x.Files $det $_ '文字コード'), (Get-Field $x.Files $det $_ '改行'), (Get-Field $x.Files $det $_ '備考') }) -join ' | ') 'ASCII のみ/LF/先頭の 1 MB で判定 | UTF-8（BOM なし）/なし/先頭の 1 MB で判定'
    Assert-Equal 'サイズは 3 桁区切り' (Get-Field $x.Files $det 'big.log' 'サイズ(バイト)') '105,906,176'
    Assert-Equal '最終更新日時' ((Get-Field $x.Files $det 'utf8.txt' '最終更新日時') -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
    Assert-Equal '判定だけのときは、結果の列が空欄' ((@($x.Files | Where-Object { $_.'結果' }).Count)) 0
    Assert-Equal '-Filter に当たらないファイルは調べない' (Get-Field $x.Files $det 'skip.dat' '文字コード') ''
    Assert-Equal '結果のフォルダ（TextEncoding_*）の中は調べない' (Get-Field $x.Files $det 'TextEncoding_old\x.txt' '文字コード') ''
    Assert-Equal 'ファイルの件数' $x.Files.Count ($detFiles.Count + 2)
    $skipText = Join-Sorted ($x.Skipped | ForEach-Object { '{0}:{1}' -f $_.'理由', $_.'パス'.Replace($det + '\', '') })
    Assert-Equal 'skipped.csv（リンク・権限なしのフォルダとファイル）' $skipText 'リンク:jloop,権限なし:denied,権限なし:noread.txt'
    Assert-Equal 'ジャンクションの種類' (@($x.Skipped | Where-Object { $_.'理由' -eq 'リンク' } | ForEach-Object { $_.'詳細' }) -join ',') 'ジャンクション'
    Assert-Equal '注記の行（スキップの種類が 1 つのときは件数を書かない）' ([regex]::Match($run.Log, 'スキップ [^所]*?件（[^）]*）  たどらなかったリンク \S+ 件').Value) 'スキップ 2 件（権限なし）  たどらなかったリンク 1 件'
    Assert-Equal '判定だけのときは、ファイルを変えない' (ConvertTo-Hex (Get-Bytes (Join-Path $det 'utf8.txt'))) (ConvertTo-Hex $detFiles['utf8.txt'][0])
    # 5.1 と 7 で同じ並びになる（文字コード・改行・パスの順）
    $order = @($x.Files | ForEach-Object { $_.'ファイル'.Replace($det + '\', '') })
    Assert-Equal 'files.csv の並び（文字コードと改行の順）' ((@($x.Files | ForEach-Object { '{0}/{1}' -f $_.'文字コード', $_.'改行' }) | Select-Object -Unique) -join ',') 'ASCII のみ/CR,ASCII のみ/CRLF,ASCII のみ/LF,ASCII のみ/なし,ASCII のみ/混在,EUC-JP/LF,JIS/CRLF,Shift_JIS/CRLF,Shift_JIS/LF,UTF-16（BOM あり）/CRLF,UTF-16（BOM あり）/LF,UTF-8（BOM あり）/CRLF,UTF-8（BOM なし）/LF,UTF-8（BOM なし）/なし,バイナリ/,判定できない/'
    Assert-Equal 'files.csv の並び（判定できないもの）' (($order | Where-Object { @('utf8bom-bad.txt', 'utf16-bad.txt', 'unknown.txt', 'nec.txt') -contains $_ }) -join ',') 'nec.txt,unknown.txt,utf16-bad.txt,utf8bom-bad.txt'
    Assert-Equal 'files.csv の並び（同じ文字コードと改行の中はパスの順）' (($order | Where-Object { @('ascii.csv', 'noline.txt', 'empty.txt', 'nul-late.txt', 'mixed.txt', 'cr.txt', 'big.log', '[a]\f[1].txt') -contains $_ }) -join ',') 'cr.txt,ascii.csv,[a]\f[1].txt,big.log,empty.txt,noline.txt,nul-late.txt,mixed.txt'
    $summaryText = $x.Lines -join "`n"
    Assert-Equal 'summary の 1 行目' (@($x.Lines) + '')[0] '文字コード判定・変換 (Convert-TextEncoding 1.0.0)'
    Assert-Equal 'summary の対象の行' ([regex]::IsMatch((@($x.Lines) + '' + '')[1], ('^対象: {0}（既定の拡張子）   \d{{4}}-\d{{2}}-\d{{2}} \d{{2}}:\d{{2}}:\d{{2}}   {1}$' -f [regex]::Escape($det), [regex]::Escape($env:COMPUTERNAME)))) 'True'
    Assert-Equal '内訳は件数の多い順' ([regex]::IsMatch($summaryText, '■ 文字コードと改行の内訳\n 文字コード\s+改行\s+件数\n UTF-8（BOM なし）\s+LF\s+5\n')) 'True'
    Assert-Equal 'summary の件数の行' ([regex]::IsMatch($summaryText, ('\nファイル {0} 件  所要 \d+\.\d秒\nスキップ 2 件（権限なし）  たどらなかったリンク 1 件\n保存先: ' -f ($detFiles.Count + 2)))) 'True'
    Assert-Equal 'summary の最後は保存先' (@($x.Lines) + '')[-2] ('保存先: {0}' -f $run.Dir)
    Assert-Equal '画面のサイズの表示' ([regex]::IsMatch($run.Log, 'ASCII のみ\s+LF\s+101\.0 MB  ' + [regex]::Escape("$det\big.log"))) 'True'
    Assert-Equal '画面のサイズの表示（バイト）' ([regex]::IsMatch($run.Log, 'ASCII のみ\s+なし\s+0 B  ' + [regex]::Escape("$det\empty.txt"))) 'True'

    Write-Section '判定: -Filter・-Exclude・-Path にファイル'
    $run = Invoke-Tool 'filter' @('-Path', $det, '-Filter', '*', '-Exclude', '*\SUB,*\big*.log,*\[a]')
    $x = Read-Result $run.Dir
    Assert-Equal '-Filter "*" はすべて調べる' (Get-Field $x.Files $det 'skip.dat' '文字コード') 'UTF-8（BOM なし）'
    Assert-Equal '-Exclude（フォルダ・ファイル・[ ] は文字のまま・大文字と小文字は区別しない）' ((@('sub\run.ps1', 'big.log', 'big-utf8.log', '[a]\f[1].txt') | ForEach-Object { Get-Field $x.Files $det $_ '文字コード' }) -join ',') ',,,'
    Assert-Equal '除外の行と件数' ('{0}/{1}' -f $run.Log.Contains('除外: *\SUB, *\big*.log, *\[a]'), ([regex]::Match($run.Log, '除外 \d+ 件').Value)) 'True/除外 4 件'
    Assert-Equal 'skipped.csv の除外指定' (Join-Sorted ($x.Skipped | Where-Object { $_.'理由' -eq '除外指定' } | ForEach-Object { '{0}={1}' -f $_.'パス'.Replace($det + '\', ''), $_.'詳細' })) '[a]=*\[a],big-utf8.log=*\big*.log,big.log=*\big*.log,sub=*\SUB'
    $run = Invoke-Tool 'filter2' @('-Path', $det, '-Filter', 'utf1?le.txt, *.BAT ,f[1].txt')
    $x = Read-Result $run.Dir
    Assert-Equal '-Filter（? は 1 文字・カンマ区切り・空白・[ ] は文字のまま・大文字と小文字は区別しない）' (Join-Sorted ($x.Files | ForEach-Object { $_.'ファイル'.Replace($det + '\', '') })) '[a]\f[1].txt,sjis.bat,utf16le.txt'
    Assert-Equal '-Filter の対象の行' ([regex]::IsMatch($run.Log, '対象: ' + [regex]::Escape($det) + '（utf1\?le\.txt, \*\.BAT, f\[1\]\.txt）')) 'True'
    $run = Invoke-Tool 'onefile' @('-Path', (Join-Path $det 'sjis-kana.txt'), '-Filter', '*.csv')
    $x = Read-Result $run.Dir
    Assert-Equal '-Path にファイル（-Filter は使わない）' ('{0}/{1}' -f $x.Files.Count, (Get-Field $x.Files $det 'sjis-kana.txt' '文字コード')) '1/Shift_JIS'
    $run = Invoke-Tool 'longfile' @('-Path', (Join-Path $det $longRel))
    Assert-Equal '-Path に 260 文字を超えるファイル' ('{0}/{1}' -f $run.Code, (Get-Field (Read-Result $run.Dir).Files $det $longRel '文字コード')) '0/UTF-8（BOM なし）'

    Write-Section '判定: 画面は 200 件まで'
    $run = Invoke-Tool 'many' @('-Path', $many)
    $x = Read-Result $run.Dir
    Assert-Equal '保存したときは、画面は 200 件まで' ('{0}/{1}/{2}' -f $run.Log.Contains('（ほか 1,300 件は files.csv を参照してください）'), $run.Log.Contains("$many\f0200.txt"), $run.Log.Contains("$many\f0201.txt")) 'True/True/False'
    Assert-Equal 'summary.txt はすべて' ('{0}/{1}' -f @($x.Lines | Where-Object { $_.EndsWith("$many\f1500.txt") }).Count, @($x.Lines | Where-Object { $_ -like '（ほか*' }).Count) '1/0'
    $run = Invoke-Plain @('-Path', $many, '-NoSave')
    Assert-Equal '-NoSave のときは、画面にすべて' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('（ほか'), $run.Log.Contains("$many\f1500.txt")) '0/False/True'

    Write-Section '変換: -To UTF8BOM -LineEnding CRLF -Force'
    $run = Invoke-Tool 'utf8bom' @('-Path', $conv, '-To', 'UTF8BOM', '-LineEnding', 'CRLF', '-Force')
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード（読み取り専用で変換できないファイルがある）' $run.Code 1
    $expected = [ordered]@{
        'u8.txt'       = @((Join-Bytes $bom $utf8NoBom.GetBytes("日本語`r`nです`r`n")), '変換した', 'UTF-8（BOM なし） → UTF-8（BOM あり）・改行 LF → CRLF')
        'u8bom.txt'    = @($null, 'すでに同じ', '')
        'sj.txt'       = @((Join-Bytes $bom $utf8NoBom.GetBytes("日本語`r`n")), '変換した', 'Shift_JIS → UTF-8（BOM あり）')
        'u16le.txt'    = @((Join-Bytes $bom $utf8NoBom.GetBytes("日本語`r`n")), '変換した', 'UTF-16（BOM あり） → UTF-8（BOM あり）')
        'u16be.txt'    = @((Join-Bytes $bom $utf8NoBom.GetBytes("日本語`r`n")), '変換した', 'UTF-16（BOM あり） → UTF-8（BOM あり）・改行 LF → CRLF')
        'ascii.txt'    = @((Join-Bytes $bom $utf8NoBom.GetBytes("abc`r`n")), '変換した', 'ASCII のみ → UTF-8（BOM あり）・改行 LF → CRLF')
        'mixed.txt'    = @((Join-Bytes $bom $utf8NoBom.GetBytes("a`r`nb`r`nc`r`n")), '変換した', 'ASCII のみ → UTF-8（BOM あり）・改行 混在 → CRLF')
        'emoji.txt'    = @((Join-Bytes $bom $utf8NoBom.GetBytes("line1`r`nスマイル ") ([byte[]](0xF0, 0x9F, 0x99, 0x82)) $utf8NoBom.GetBytes("`r`n")), '変換した', 'UTF-8（BOM なし） → UTF-8（BOM あり）・改行 LF → CRLF')
        'euc.txt'      = @($null, '対象外', 'EUC-JP は変換しません')
        'bin.txt'      = @($null, '対象外', 'バイナリ は変換しません')
        'readonly.txt' = @($null, '変換できない', '読み取り専用のファイル')
        'hidden.txt'   = @((Join-Bytes $bom $utf8NoBom.GetBytes("隠し`r`n")), '変換した', 'UTF-8（BOM なし） → UTF-8（BOM あり）・改行 LF → CRLF')
        'acl.txt'      = @((Join-Bytes $bom $utf8NoBom.GetBytes("アクセス権`r`n")), '変換した', 'UTF-8（BOM なし） → UTF-8（BOM あり）・改行 LF → CRLF')
        'run.ps1'      = @((Join-Bytes $bom $utf8NoBom.GetBytes("Write-Host '日本語'`r`n")), '変換した', 'Shift_JIS → UTF-8（BOM あり）')
        'run.bat'      = @((Join-Bytes $bom $utf8NoBom.GetBytes("@echo off`r`necho 日本語`r`n")), '変換した', 'Shift_JIS → UTF-8（BOM あり）')
        '[b]\g.txt'    = @((Join-Bytes $bom $utf8NoBom.GetBytes("かっこ`r`n")), '変換した', 'UTF-8（BOM なし） → UTF-8（BOM あり）・改行 LF → CRLF')
        'sub\deep.txt' = @((Join-Bytes $bom $utf8NoBom.GetBytes("サブ`r`n")), '変換した', 'UTF-8（BOM なし） → UTF-8（BOM あり）・改行 LF → CRLF')
        $longRel       = @((Join-Bytes $bom $utf8NoBom.GetBytes("長いパス`r`n")), '変換した', 'UTF-8（BOM なし） → UTF-8（BOM あり）・改行 LF → CRLF')
    }
    foreach ($rel in $expected.Keys) {
        $label = $rel.Substring(0, [math]::Min(30, $rel.Length))
        Assert-Equal ('{0}: 中身' -f $label) (Test-ConvBytes $rel $expected[$rel][0]) 'True'
        Assert-Equal ('{0}: 結果と備考' -f $label) ('{0}/{1}' -f (Get-Field $x.Files $conv $rel '結果'), (Get-Field $x.Files $conv $rel '備考')) ('{0}/{1}' -f $expected[$rel][1], $expected[$rel][2])
    }
    $converted = @($expected.Keys | Where-Object { $expected[$_][1] -eq '変換した' })
    Assert-Equal '控え: 変換したファイルだけを、フォルダの構成で、元の中身のまま' (@($converted | Where-Object { (ConvertTo-Hex (Get-Bytes (Join-Path $run.Dir ('backup\' + $_)))) -ne (ConvertTo-Hex $convFiles[$_]) }).Count) 0
    Assert-Equal '控え: 変換しなかったファイルは控えない' (@(Get-ChildItem -LiteralPath ('\\?\' + (Join-Path $run.Dir 'backup')) -Recurse -File).Count) $converted.Count
    # 書き換えると Windows がアーカイブの属性を付けるため、隠しと読み取り専用が付いているかだけを見る
    Assert-Equal '隠し・読み取り専用の属性は変えない' ('{0}/{1}' -f (([IO.File]::GetAttributes((Join-Path $conv 'hidden.txt')) -band [IO.FileAttributes]::Hidden) -ne 0), (([IO.File]::GetAttributes((Join-Path $conv 'readonly.txt')) -band [IO.FileAttributes]::ReadOnly) -ne 0)) 'True/True'
    Assert-Equal 'アクセス権は変えない' ((Get-Acl -LiteralPath (Join-Path $conv 'acl.txt')).Sddl -eq $aclBefore) 'True'
    Assert-Equal '最終更新日時: 変換したファイルは新しく、ほかは元のまま' ('{0}/{1}' -f ([IO.File]::GetLastWriteTime((Join-Path $conv 'u8.txt')) -gt (Get-Date).AddMinutes(-10)), ([IO.File]::GetLastWriteTime((Join-Path $conv 'u8bom.txt')) -eq $oldTime)) 'True/True'
    Assert-Equal 'files.csv のサイズと日時は、変換した後の値' ('{0}/{1}' -f (Get-Field $x.Files $conv 'u8.txt' 'サイズ(バイト)'), ((Get-Field $x.Files $conv 'u8.txt' '最終更新日時') -ne $oldTime.ToString('yyyy/MM/dd HH:mm:ss'))) ('{0}/True' -f $expected['u8.txt'][0].Length)
    $summaryText = $run.Log
    Assert-Equal '予定の見出し' $run.Log.Contains('■ 変換の予定（UTF-8（BOM あり）・CRLF にそろえる）変換するファイル 14 件') 'True'
    Assert-Equal '予定の内訳' ([regex]::IsMatch($run.Log, 'UTF-8（BOM なし） → UTF-8（BOM あり）\s+7 件')) 'True'
    Assert-Equal '変換しないファイル' $run.Log.Contains('変換しないファイル 4 件（すでに同じ 1 件・対象外 2 件・変換できない 1 件）') 'True'
    Assert-Equal '.bat の注意（BOM ありでも）' $run.Log.Contains('※ .bat・.cmd を UTF-8 にすると、コマンドプロンプトで日本語が化けます（1 件）') 'True'
    Assert-Equal '.ps1 の注意は出さない（BOM あり）' $run.Log.Contains('※ .ps1') 'False'
    Assert-Equal '結果の行' ([regex]::IsMatch($run.Log, '結果: 変換しました  14 件  失敗 0 件  所要 \d+\.\d秒保存先: ')) 'True'
    Assert-Equal 'summary にも予定と結果' ('{0}/{1}' -f ((Read-Result $run.Dir).Lines -contains '■ 変換の予定（UTF-8（BOM あり）・CRLF にそろえる）'), (@((Read-Result $run.Dir).Lines | Where-Object { $_ -like '結果: 変換しました  14 件*' }).Count)) 'True/1'

    Write-Section '変換: -To SJIS（改行はそのまま）'
    Reset-Conv
    $run = Invoke-Tool 'sjis' @('-Path', $conv, '-To', 'SJIS', '-Force')
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 1
    Assert-Equal 'UTF-8 → Shift_JIS（LF のまま）' ('{0}/{1}' -f (Test-ConvBytes 'u8.txt' $sjis.GetBytes("日本語`nです`n")), (Get-Field $x.Files $conv 'u8.txt' '備考')) 'True/UTF-8（BOM なし） → Shift_JIS'
    Assert-Equal 'UTF-8（BOM あり）→ Shift_JIS（BOM を外す）' (Test-ConvBytes 'u8bom.txt' $sjis.GetBytes("日本語`r`nです`r`n")) 'True'
    Assert-Equal 'UTF-16 BE → Shift_JIS' (Test-ConvBytes 'u16be.txt' $sjis.GetBytes("日本語`n")) 'True'
    Assert-Equal 'Shift_JIS と ASCII のみは、すでに同じ' ((@('sj.txt', 'ascii.txt', 'run.ps1') | ForEach-Object { Get-Field $x.Files $conv $_ '結果' }) -join ',') 'すでに同じ,すでに同じ,すでに同じ'
    $smile = $utf8NoBom.GetString([byte[]](0xF0, 0x9F, 0x99, 0x82))
    Assert-Equal 'Shift_JIS にない文字' ('{0}/{1}/{2}' -f (Get-Field $x.Files $conv 'emoji.txt' '結果'), (Get-Field $x.Files $conv 'emoji.txt' '備考'), (Test-ConvBytes 'emoji.txt' $null)) ('変換できない/Shift_JIS にない文字: {0}（2 行目）/True' -f $smile)

    Write-Section '変換: -LineEnding LF（文字コードはそのまま）'
    Reset-Conv
    $run = Invoke-Tool 'lf' @('-Path', $conv, '-LineEnding', 'LF', '-Force')
    $x = Read-Result $run.Dir
    Assert-Equal 'Shift_JIS のまま改行だけ' ('{0}/{1}' -f (Test-ConvBytes 'sj.txt' $sjis.GetBytes("日本語`n")), (Get-Field $x.Files $conv 'sj.txt' '備考')) 'True/改行 CRLF → LF'
    Assert-Equal 'UTF-8 の BOM は残す' (Test-ConvBytes 'u8bom.txt' (Join-Bytes $bom $utf8NoBom.GetBytes("日本語`nです`n"))) 'True'
    Assert-Equal 'UTF-16 LE のまま' (Test-ConvBytes 'u16le.txt' (Join-Bytes ([byte[]](0xFF, 0xFE)) $le.GetBytes("日本語`n"))) 'True'
    Assert-Equal '改行の混在をそろえる' (Test-ConvBytes 'mixed.txt' $utf8NoBom.GetBytes("a`nb`nc`n")) 'True'
    Assert-Equal 'すでに LF' ((@('u8.txt', 'u16be.txt', 'ascii.txt') | ForEach-Object { Get-Field $x.Files $conv $_ '結果' }) -join ',') 'すでに同じ,すでに同じ,すでに同じ'
    Assert-Equal '予定の見出し（改行だけ）' $run.Log.Contains('■ 変換の予定（LF にそろえる）') 'True'

    Write-Section '変換: -To UTF8（BOM なし）と確認で N'
    Reset-Conv
    $run = Invoke-Tool 'no' @('-Path', $conv, '-To', 'UTF8') 'N'
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード（N のときは 0）' $run.Code 0
    Assert-Equal '変換しない' (@($convFiles.Keys | Where-Object { -not (Test-ConvBytes $_ $null) }).Count) 0
    Assert-Equal '結果と備考' ('{0}/{1}' -f (Get-Field $x.Files $conv 'sj.txt' '結果'), (Get-Field $x.Files $conv 'sj.txt' '備考')) '変換しなかった/Shift_JIS → UTF-8（BOM なし）'
    Assert-Equal '表示' ('{0}/{1}' -f $run.Log.Contains('変換しませんでした'), $run.Log.Contains('結果: ')) 'True/False'
    Assert-Equal '.ps1 の注意' $run.Log.Contains('※ .ps1 を BOM なしの UTF-8 にすると、Windows PowerShell 5.1 で日本語が化けます（1 件）') 'True'
    Assert-Equal '.bat の注意' $run.Log.Contains('※ .bat・.cmd を UTF-8 にすると、コマンドプロンプトで日本語が化けます（1 件）') 'True'
    Assert-Equal '控えは作らない' ([bool]$run.Dir -and -not (Test-Path -LiteralPath (Join-Path $run.Dir 'backup'))) 'True'

    Write-Section '確認の入力'
    $answers = @(
        @('y', '小文字の y'),
        @(([string][char]0xFF39), '全角の Ｙ'),
        @("x`nYes", 'Y/N 以外のあとに Yes')
    )
    $n = 0
    foreach ($answer in $answers) {
        Reset-Conv
        $n++
        $run = Invoke-Tool "answer$n" @('-Path', (Join-Path $conv 'sub'), '-LineEnding', 'CRLF') $answer[0]
        Assert-Equal ('確認: {0} で変換する' -f $answer[1]) ('{0}/{1}' -f $run.Code, (Test-ConvBytes 'sub\deep.txt' $utf8NoBom.GetBytes("サブ`r`n"))) '0/True'
    }
    Reset-Conv
    $run = Invoke-Tool 'eof' @('-Path', (Join-Path $conv 'sub'), '-LineEnding', 'CRLF') ''
    Assert-Equal '入力が終わったら変換せずに止まり、結果を残さない' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('確認できないため、変換しません'), (Test-ConvBytes 'sub\deep.txt' $null), [bool]$run.Dir) '1/True/True/False'

    Write-Section '変換: そろった後と、-Path にファイル'
    Reset-Conv
    $clean = Join-Path $conv 'sub'
    $run = Invoke-Tool 'clean1' @('-Path', $clean, '-To', 'UTF8BOM', '-Force')
    $run = Invoke-Tool 'clean2' @('-Path', $clean, '-To', 'UTF8BOM', '-Force')
    Assert-Equal 'そろった後は、確認せずに終わる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('変換するファイル 0 件'), $run.Log.Contains('変換するファイルはありません（すべてそろっています）')) '0/True/True'
    $run = Invoke-Tool 'cannot' @('-Path', $conv, '-To', 'UTF8BOM', '-Force')
    $run = Invoke-Tool 'cannot2' @('-Path', $conv, '-To', 'UTF8BOM', '-Force')
    Assert-Equal '変換できないファイルがあるときは「そろっています」と言わない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('変換するファイルはありません'), $run.Log.Contains('すべてそろっています')) '1/True/False'
    Reset-Conv
    $run = Invoke-Tool 'single' @('-Path', (Join-Path $conv 'sub\deep.txt'), '-LineEnding', 'CRLF', '-Force')
    Assert-Equal '-Path にファイル: 控えはファイル名だけ' ('{0}/{1}/{2}' -f $run.Code, (Test-ConvBytes 'sub\deep.txt' $utf8NoBom.GetBytes("サブ`r`n")), (ConvertTo-Hex (Get-Bytes (Join-Path $run.Dir 'backup\deep.txt')))) ('0/True/{0}' -f (ConvertTo-Hex $convFiles['sub\deep.txt']))

    Write-Section '確認を待つ間に変わったファイル（追記・削除・ほかのプログラムが開いている）'
    Reset-Conv
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = (Get-Command $Exe).Source
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Path "{1}" -LineEnding CRLF -OutputFolder "{2}"' -f $tool, $conv, (Join-Path $workRoot 'out\wait')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $proc = [Diagnostics.Process]::Start($psi)
    $errTask = $proc.StandardError.ReadToEndAsync()
    # 確認の直前に表示する「backup フォルダに控えます」の行が出るまで待つ（表示の文字コードによらず、backup は読める）
    $prompted = $false
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not $prompted -and $watch.Elapsed.TotalSeconds -lt 60) {
        $lineTask = $proc.StandardOutput.ReadLineAsync()
        if (-not $lineTask.Wait(60000)) { break }
        if ($null -eq $lineTask.Result) { break }
        if ($lineTask.Result.Contains('backup')) { $prompted = $true }
    }
    if ($prompted) {
        [IO.File]::AppendAllText((Join-Path $conv 'u8.txt'), "追記`n", $utf8NoBom)
        [IO.File]::Delete((Join-Path $conv 'ascii.txt'))
        $lock = New-Object IO.FileStream((Join-Path $conv 'mixed.txt'), [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $holds.Add($lock)
        $proc.StandardInput.WriteLine('Y')
        $proc.StandardInput.Close()
        $restTask = $proc.StandardOutput.ReadToEndAsync()
        if (-not $proc.WaitForExit(120000)) { $proc.Kill() }
        $lock.Dispose()
        $x = Read-Result (Get-RunDir (Join-Path $workRoot 'out\wait'))
        Assert-Equal '終了コード' $proc.ExitCode 1
        Assert-Equal '追記されたファイルは失敗にして、追記を残す' ('{0}/{1}/{2}' -f (Get-Field $x.Files $conv 'u8.txt' '結果'), (Get-Field $x.Files $conv 'u8.txt' '備考'), (Test-ConvBytes 'u8.txt' (Join-Bytes $convFiles['u8.txt'] $utf8NoBom.GetBytes("追記`n")))) '失敗/調べた後に変更されたため、変換していません/True'
        Assert-Equal '消されたファイルは失敗' ('{0}/{1}' -f (Get-Field $x.Files $conv 'ascii.txt' '結果'), ((Get-Field $x.Files $conv 'ascii.txt' '備考') -like '開けないため、変換していません: *')) '失敗/True'
        Assert-Equal 'ほかのプログラムが開いているファイルは失敗にして、変えない' ('{0}/{1}/{2}' -f (Get-Field $x.Files $conv 'mixed.txt' '結果'), ((Get-Field $x.Files $conv 'mixed.txt' '備考') -like '開けないため、変換していません: *'), (Test-ConvBytes 'mixed.txt' $null)) '失敗/True/True'
        Assert-Equal 'ほかのファイルは変換する' ('{0}/{1}' -f (Get-Field $x.Files $conv 'u16be.txt' '結果'), (Test-ConvBytes 'sub\deep.txt' $utf8NoBom.GetBytes("サブ`r`n"))) '変換した/True'
        # 子のプロセスの画面の出力は、文字コードがずれることがあるため、summary.txt で確かめる
        Assert-Equal '失敗したファイルがあるときの結果の行' (@($x.Lines | Where-Object { $_ -match '^結果: 変換できなかったファイルがあります  変換 \d+ 件  失敗 3 件  所要 \S+（files.csv を確認してください）$' }).Count) 1
    } else {
        Write-Skip '確認を待つ間に変わったファイル' '確認の表示が出なかったため'
        try { $proc.Kill() } catch { }
    }

    Write-Section '控えを保存できないとき・保存先に書き込めないとき'
    Reset-Conv
    # 保存先のフォルダは作れるが、その中にファイルを作れないようにする
    $guardOut = Join-Path $workRoot 'out\guard'
    [void][IO.Directory]::CreateDirectory($guardOut)
    if (Add-Deny $guardOut '(OI)(CI)(WD)') {
        $log = ConvertTo-FlatText ((@('' | & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path (Join-Path $conv 'sub') -LineEnding CRLF -Force -OutputFolder $guardOut 2>&1) | ForEach-Object { "$_" }) -join "`n")
        Assert-Equal '控えられないときは変換しない' ('{0}/{1}/{2}' -f $LASTEXITCODE, (Test-ConvBytes 'sub\deep.txt' $null), $log.Contains('結果を保存できませんでした')) '1/True/True'
    } else {
        Write-Skip '控えを保存できないとき' 'アクセス拒否を設定できなかったため'
    }
    $freeLetter = Get-FreeLetter
    if ($freeLetter) {
        $run = Invoke-Plain @('-Path', $det, '-OutputFolder', "${freeLetter}:\nope")
        Assert-Equal '書き込めない保存先: 判定だけなら警告して、すべて表示する' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("保存先に書き込めないため、結果は保存しません: ${freeLetter}:\nope"), $run.Log.Contains("$det\utf8.txt")) '0/True/True'
        $run = Invoke-Plain @('-Path', $conv, '-To', 'UTF8BOM', '-Force', '-OutputFolder', "${freeLetter}:\nope")
        Assert-Equal '書き込めない保存先: 変換するときは止まる' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため、変換しません（元のファイルを控えられないため）'), (Test-ConvBytes 'sj.txt' $null), $run.Log.Contains('■ 文字コードと改行の内訳')) '1/True/True/False'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }

    Write-Section '保存先'
    $jpOut = Join-Path $workRoot 'out\保存 先'
    $run = Invoke-Plain @('-Path', (Join-Path $det 'sub'), '-OutputFolder', $jpOut)
    Assert-Equal '日本語とスペースを含む保存先' ('{0}/{1}' -f $run.Code, (Get-ResultNames (Get-RunDir $jpOut))) '0/files.csv,skipped.csv,summary.txt'
    $run = Invoke-Plain @('-Path', (Join-Path $det 'sub'), '-OutputFolder', (Join-Path $workRoot 'out\trail.'))
    Assert-Equal '末尾にピリオドが付いた保存先' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultNames (Get-RunDir (Join-Path $workRoot 'out\trail'))), [IO.Directory]::Exists('\\?\' + (Join-Path $workRoot 'out\trail.'))) '0/files.csv,skipped.csv,summary.txt/False'
    $run = Invoke-Plain @('-Path', $det, '-NoSave', '-OutputFolder', (Join-Path $workRoot 'out\nosave'))
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}/{2}' -f $run.Code, (Test-Path -LiteralPath (Join-Path $workRoot 'out\nosave')), $run.Log.Contains('保存先')) '0/False/False'
    $cwd = Join-Path $workRoot 'cwd'
    New-TestFile "$cwd\rel\a.txt" $utf8NoBom.GetBytes("あ`n")
    $run = Invoke-From $cwd @('-Path', 'rel')
    Assert-Equal '既定の保存先はスクリプトのフォルダ・相対パスの -Path' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultNames (Get-RunDir (Split-Path $tool))), $run.Log.Contains("$cwd\rel\a.txt")) '0/files.csv,skipped.csv,summary.txt/True'
    $run = Invoke-From $cwd @('-OutputFolder', 'out')
    Assert-Equal '-Path を省略すると現在のフォルダ・相対パスの -OutputFolder' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultNames (Get-RunDir "$cwd\out")), $run.Log.Contains("対象: $cwd（")) '0/files.csv,skipped.csv,summary.txt/True'

    Write-Section '260 文字を超えるパス'
    $longOut = "$workRoot\" + ('O' * 150) + '\' + ('P' * 150)
    Reset-Conv
    $run = Invoke-Plain @('-Path', (Join-Path $conv ($longRel.Substring(0, $longRel.LastIndexOf('\')))), '-LineEnding', 'CRLF', '-Force', '-OutputFolder', $longOut)
    $longRun = @(Get-ChildItem -LiteralPath ('\\?\' + $longOut) -Directory -ErrorAction SilentlyContinue)
    $longNames = if ($longRun.Count -gt 0) { Join-Sorted (Get-ChildItem -LiteralPath $longRun[0].FullName | ForEach-Object { $_.Name }) } else { '(なし)' }
    Assert-Equal '長い -Path と -OutputFolder で変換する' ('{0}/{1}/{2}' -f $run.Code, $longNames, (Test-ConvBytes $longRel $utf8NoBom.GetBytes("長いパス`r`n"))) '0/backup,files.csv,skipped.csv,summary.txt/True'
    # 5.1 の Join-Path は \\?\ で始まるパスを扱えないため、文字列でつなぐ
    $longBackup = if ($longRun.Count -gt 0) { Get-Bytes ($longRun[0].FullName.Substring(4) + '\backup\long.txt') } else { $null }
    Assert-Equal '長い保存先に控える' (ConvertTo-Hex $longBackup) (ConvertTo-Hex $convFiles[$longRel])

    Write-Section '名前に [ ] を含むフォルダから起動（-Path と -OutputFolder を相対パスで指定）'
    # 5.1 は、名前に [ ] を含むフォルダで起動すると、名前をワイルドカードとして扱い、別の場所から始まる
    $startRoot = Join-Path $workRoot 'start'
    $startCases = @(
        @('一致するフォルダがない（$PSHOME から始まる）', 'br[1] x', @()),
        @('隣に 1 つ一致するフォルダがある', 'sib\[a]', @('sib\a')),
        @('[ が閉じていない（ドライブのルートから始まる）', 'ub\a[b', @('ub\ab')),
        @('隣に 2 つ一致する（ドライブのルートから始まる）', 'mm\[ab]', @('mm\a', 'mm\b'))
    )
    foreach ($case in $startCases) {
        $dir = "$startRoot\$($case[1])"
        New-TestFile "$dir\here\here.txt" $utf8NoBom.GetBytes("ここ`n")
        # 間違ったフォルダを調べたら分かるように、隣のフォルダには別のファイルを置く
        foreach ($decoy in $case[2]) { New-TestFile "$startRoot\$decoy\here\decoy.txt" $utf8NoBom.GetBytes("隣`n") }
        $run = Invoke-From $dir @('-Path', 'here', '-OutputFolder', 'out')
        Assert-Equal $case[0] ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('here.txt'), $run.Log.Contains('decoy.txt'), (Get-ResultNames (Get-RunDir "$dir\out"))) '0/True/False/files.csv,skipped.csv,summary.txt'
    }

    Write-Section 'コマンドプロンプトから起動'
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $run = Invoke-ViaCmd ('{0} -NoSave -Filter "*.csv,*.bat" -Path "{1}\"' -f $cmdBase, $det)
    Assert-Equal '最後の引数が "パス\"・カンマ区切りの -Filter' ('{0}/{1}' -f $run.Code, [regex]::Match($run.Log, 'ファイル \S+ 件').Value) '0/ファイル 2 件'
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -NoSave' -f $cmdBase, $det)
    Assert-Equal '"パス\" の後ろに引数: ヒントを出して止まる（%ERRORLEVEL% は 1）' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('-Path の値に引用符'), $run.Log.Contains('最後に \ を付けないでください'), $run.Log.Contains('文字コード判定・変換 (')) '1/True/True/False'
    $run = Invoke-ViaCmd ('{0} -Path ''{1}'' -Exclude ''*\sub'',''*\big*'' -NoSave' -f $cmdBase, $det)
    Assert-Equal 'cmd で '' で囲んだ -Path と -Exclude' ('{0}/{1}' -f $run.Code, [regex]::Match($run.Log, '除外 \d+ 件').Value) '0/除外 3 件'
    # -NonInteractive は powershell.exe の引数なので、-File より前に書く
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{1}" -Path "{2}" -To UTF8BOM -OutputFolder "{3}"' -f $Exe, $tool, (Join-Path $conv 'sub'), (Join-Path $workRoot 'out\nonint'))
    Assert-Equal '-NonInteractive では確認できないので変換しない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('確認できないため、変換しません'), (Test-ConvBytes 'sub\deep.txt' $null)) '1/True/True'
    # 保存先の名前を nul にすると、デバイスの NUL になるため、別の名前にする
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -To UTF8BOM -OutputFolder "{2}"' -f $cmdBase, (Join-Path $conv 'sub'), (Join-Path $workRoot 'out\stdin-empty'))
    Assert-Equal '標準入力が NUL のときも変換しない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('確認できないため、変換しません'), (Test-ConvBytes 'sub\deep.txt' $null)) '1/True/True'

    Write-Section '引数の誤り'
    $cases = @(
        @('存在しないパス', @('-Path', (Join-Path $workRoot 'no\such')), "指定したフォルダかファイルが見つかりません: $workRoot\no\such"),
        @('変換と -NoSave', @('-Path', $det, '-To', 'UTF8', '-NoSave'), '変換するときは、-NoSave は指定できません'),
        @('-To の値の誤り', @('-Path', $det, '-To', 'EUC'), 'To'),
        @('-LineEnding の値の誤り', @('-Path', $det, '-LineEnding', 'CR'), 'LineEnding')
    )
    foreach ($case in $cases) {
        $run = Invoke-Plain $case[1]
        Assert-Equal $case[0] ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains($case[2]), $run.Log.Contains('■ 文字コードと改行の内訳')) '1/True/False'
    }

    Write-Section 'Ctrl+C で中断'
    $stopScript = New-HelperScript 'stop.ps1' @(
        'param($Tool, $Target, $OutRoot, $Mode)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $Target).AddParameter(''OutputFolder'', $OutRoot)'
        'if ($Mode -eq ''convert'') { [void]$ps.AddParameter(''LineEnding'', ''CRLF'').AddParameter(''Force'', $true) }'
        '$handle = $ps.BeginInvoke()'
        '# 調べている途中なら結果のフォルダができたら、変換の途中なら backup ができたら止める'
        '$wait = if ($Mode -eq ''convert'') { ''*\backup'' } else { ''*'' }'
        'for ($i = 0; $i -lt 1500 -and -not $handle.IsCompleted; $i++) {'
        '    if (@(Get-ChildItem -LiteralPath $OutRoot -Directory -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.FullName -like $wait }).Count -gt 0) { break }'
        '    Start-Sleep -Milliseconds 20'
        '}'
        'Start-Sleep -Milliseconds $(if ($Mode -eq ''convert'') { 300 } else { 30 })'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
    )
    # Stop() は、動いている最中に呼んでも、止める合図を取りこぼして最後まで動き続けることがある（5.1 でも 7 でも）。止まるまで 3 回まで試す
    $stopped = $false
    for ($attempt = 1; $attempt -le 3 -and -not $stopped; $attempt++) {
        $result = Invoke-Helper $stopScript @('-Tool', $tool, '-Target', $many, '-OutRoot', (Join-Path $workRoot "out\stop-scan$attempt"), '-Mode', 'scan')
        $stopped = $result['RUNNING'] -eq 'True' -and $result['MESSAGE'] -ne ''
    }
    if ($stopped) {
        Assert-Equal '調べている途中: 結果を残さない' ('{0}/{1}' -f $result['AFTER'], $result['MESSAGE']) '0/中断しました。結果は保存していません。'
    } else {
        Write-Skip '調べている途中で中断' '止める前に調べ終わったか、Stop() が効かなかったため（3 回試した）'
    }
    # 変換の途中で止めたときも、止められなかったときも、どのファイルも「元のまま」か「控えを取ったうえで変換した後」で、結果の列と合うこと
    $converted = $utf8NoBom.GetBytes($utf8NoBom.GetString($manyText).Replace("`n", "`r`n"))
    $broken   = New-Object 'System.Collections.Generic.List[string]'
    $noBackup = New-Object 'System.Collections.Generic.List[string]'
    $stopped  = $false
    # 取りこぼしは、このテストの条件だと 10 回に 5 回ほど起きる（3 回では 1 割ほどが全部外れて、たまに失敗していた）。止まるまで 10 回まで試す
    for ($attempt = 1; $attempt -le 10 -and -not $stopped; $attempt++) {
        for ($i = 1; $i -le 1500; $i++) { [IO.File]::WriteAllBytes((Join-Path $many ('f{0:0000}.txt' -f $i)), $manyText) }
        $stopOut = Join-Path $workRoot "out\stop-conv$attempt"
        $result = Invoke-Helper $stopScript @('-Tool', $tool, '-Target', $many, '-OutRoot', $stopOut, '-Mode', 'convert')
        $stopDir = Get-RunDir $stopOut
        $x = Read-Result $stopDir
        $states = @{}
        foreach ($row in $x.Files) {
            $states[$row.'結果'] = 1 + [int]$states[$row.'結果']
            $bytes = ConvertTo-Hex (Get-Bytes $row.'ファイル')
            $isNew = $bytes -eq (ConvertTo-Hex $converted)
            if (-not $isNew -and $bytes -ne (ConvertTo-Hex $manyText)) { $broken.Add($row.'ファイル') }
            if ($row.'結果' -eq '変換した' -and -not $isNew) { $broken.Add($row.'ファイル') }
            if ($row.'結果' -eq '変換しなかった' -and $isNew) { $broken.Add($row.'ファイル') }
            $backupHex = ConvertTo-Hex (Get-Bytes (Join-Path $stopDir ('backup\' + (Split-Path -Leaf $row.'ファイル'))))
            if ($isNew -and $backupHex -ne (ConvertTo-Hex $manyText)) { $noBackup.Add($row.'ファイル') }
        }
        $stopped = $result['RUNNING'] -eq 'True' -and $states.ContainsKey('変換しなかった')
        '  （{0} 回目: 結果 {1}）' -f $attempt, ((@($states.Keys | Sort-Object) | ForEach-Object { '{0} {1} 件' -f $_, $states[$_] }) -join '・')
    }
    Assert-Equal '変換の途中で止めても: どのファイルも元のままか変換した後で、結果の列と合う' ($broken -join ',') ''
    Assert-Equal '変換の途中で止めても: 変換したファイルはすべて控えてある' ($noBackup -join ',') ''
    if ($stopped) {
        Assert-Equal '変換の途中: 結果と控えを残す' ('{0}/{1}' -f $x.Files.Count, $states.ContainsKey('変換した')) '1500/True'
        Assert-Equal '変換の途中: 変換しなかったファイルの備考' (@($x.Files | Where-Object { $_.'結果' -eq '変換しなかった' } | ForEach-Object { $_.'備考' } | Select-Object -Unique) -join ',') '中断したため、変換していません'
        Assert-Equal '変換の途中: 表示' ((ConvertTo-FlatText $result['MESSAGE']) -like '中断しました。それまでに変換したファイルと、控えたファイルは残っています*') 'True'
    } elseif ($hostTag -eq '7') {
        # 10 回続けて止まらなければ、取りこぼしではなく、中断したときの処理の誤り
        Assert-Equal '変換の途中で中断（7）' '止まらなかった' '止まった'
    } else {
        Write-Skip '変換の途中で中断' '止める前に変換し終わったか、5.1 の Stop() が効かなかったため（10 回試した）'
    }

    Write-Section 'Ctrl+C（変換するときは、コンソールのキーの入力として受け取って止める）'
    # 新しい（見えない）コンソールでツールを動かし、そのコンソールの入力に Ctrl+C のキーを書き込む
    $keyScript = New-HelperScript 'send-ctrlc.ps1' @(
        'param([int]$TargetPid)'
        'Add-Type -TypeDefinition @'''
        'using System;'
        'using System.Runtime.InteropServices;'
        'public static class WorkbenchConsoleKeys {'
        '    [DllImport("kernel32.dll", SetLastError = true)] static extern bool FreeConsole();'
        '    [DllImport("kernel32.dll", SetLastError = true)] static extern bool AttachConsole(uint pid);'
        '    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)] static extern IntPtr CreateFileW(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);'
        '    [DllImport("kernel32.dll", SetLastError = true)] static extern bool WriteConsoleInputW(IntPtr handle, INPUT_RECORD[] records, uint length, out uint written);'
        '    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);'
        '    [DllImport("kernel32.dll", SetLastError = true)] static extern bool GetConsoleMode(IntPtr handle, out uint mode);'
        '    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] struct KEY_EVENT_RECORD { public int KeyDown; public ushort RepeatCount; public ushort VirtualKeyCode; public ushort VirtualScanCode; public char UnicodeChar; public uint ControlKeyState; }'
        '    [StructLayout(LayoutKind.Explicit)] struct INPUT_RECORD { [FieldOffset(0)] public ushort EventType; [FieldOffset(4)] public KEY_EVENT_RECORD KeyEvent; }'
        '    public static string SendCtrlC(uint pid) {'
        '        FreeConsole();'
        '        if (!AttachConsole(pid)) { return "AttachConsole " + Marshal.GetLastWin32Error(); }'
        '        IntPtr input = CreateFileW("CONIN$", 0xC0000000, 3, IntPtr.Zero, 3, 0, IntPtr.Zero);'
        '        if (input == new IntPtr(-1)) { FreeConsole(); return "CONIN$ " + Marshal.GetLastWin32Error(); }'
        '        // Ctrl+C をキーの入力として受け取るときは、入力のモードから ENABLE_PROCESSED_INPUT（1）が外れている'
        '        uint mode;'
        '        if (!GetConsoleMode(input, out mode)) { mode = 0xFFFFFFFF; }'
        '        INPUT_RECORD[] records = new INPUT_RECORD[2];'
        '        for (int i = 0; i < 2; i++) {'
        '            records[i].EventType = 1;'
        '            records[i].KeyEvent.KeyDown = i == 0 ? 1 : 0;'
        '            records[i].KeyEvent.RepeatCount = 1;'
        '            records[i].KeyEvent.VirtualKeyCode = 0x43;'
        '            records[i].KeyEvent.VirtualScanCode = 0x2E;'
        '            records[i].KeyEvent.UnicodeChar = (char)3;'
        '            records[i].KeyEvent.ControlKeyState = 0x0008;'
        '        }'
        '        uint written;'
        '        bool ok = WriteConsoleInputW(input, records, 2, out written);'
        '        int error = Marshal.GetLastWin32Error();'
        '        CloseHandle(input);'
        '        FreeConsole();'
        '        return (ok ? "ok" : "WriteConsoleInput " + error) + "|" + mode;'
        '    }'
        '}'
        '''@'
        '$parts = [WorkbenchConsoleKeys]::SendCtrlC([uint32]$TargetPid).Split(''|'')'
        '''RESULT='' + $parts[0]'
        '''MODE='' + $(if ($parts.Count -gt 1) { $parts[1] } else { '''' })'
    )
    # $Wait: scan（結果のフォルダができたら＝調べ始めたら）か convert（backup ができたら＝変換を始めたら）
    function Invoke-ConsoleStop([string]$Label, [string]$Wait) {
        for ($i = 1; $i -le 1500; $i++) { [IO.File]::WriteAllBytes((Join-Path $many ('f{0:0000}.txt' -f $i)), $manyText) }
        $out  = Join-Path $workRoot "out\$Label"
        $proc = Start-Process -FilePath (Get-Command $Exe).Source -WindowStyle Hidden -PassThru -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $tool), '-Path', ('"{0}"' -f $many), '-LineEnding', 'CRLF', '-Force', '-OutputFolder', ('"{0}"' -f $out))
        $null = $proc.Handle   # 終了コードを受け取れるように、終わる前にハンドルを開いておく
        $pattern = if ($Wait -eq 'convert') { '*\backup' } else { '*\TextEncoding_*' }
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $ready = $false
        while (-not $ready -and -not $proc.HasExited -and $watch.Elapsed.TotalSeconds -lt 60) {
            $ready = @(Get-ChildItem -LiteralPath $out -Directory -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.FullName -like $pattern }).Count -gt 0
            if (-not $ready) { Start-Sleep -Milliseconds 20 }
        }
        $sent = '(送る前に終わった)'
        $mode = ''
        if ($ready -and -not $proc.HasExited) {
            $result = Invoke-Helper $keyScript @('-TargetPid', $proc.Id)
            $sent = $result['RESULT']
            $mode = $result['MODE']
        }
        if (-not $proc.WaitForExit(120000)) { $proc.Kill(); $sent += '（止まらなかった）' }
        # ENABLE_PROCESSED_INPUT（1）が外れていれば、キーボードの Ctrl+C も、合図ではなくキーの入力になる
        $keyMode = if ($mode -match '^\d+$') { ([long]$mode -band 1) -eq 0 } else { "(読めない: $mode)" }
        [pscustomobject]@{ Code = $proc.ExitCode; Sent = $sent; KeyMode = $keyMode; Dir = (Get-RunDir $out) }
    }
    function Get-ManyStates($Rows, [string]$Dir) {
        $broken = 0; $noBackup = 0
        foreach ($row in $Rows) {
            $hex = ConvertTo-Hex (Get-Bytes $row.'ファイル')
            $isNew = $hex -eq (ConvertTo-Hex $converted)
            if (-not $isNew -and $hex -ne (ConvertTo-Hex $manyText)) { $broken++ }
            if (($row.'結果' -eq '変換した') -ne $isNew) { $broken++ }
            if ($isNew -and (ConvertTo-Hex (Get-Bytes (Join-Path $Dir ('backup\' + (Split-Path -Leaf $row.'ファイル'))))) -ne (ConvertTo-Hex $manyText)) { $noBackup++ }
        }
        '{0}/{1}' -f $broken, $noBackup
    }
    $run = Invoke-ConsoleStop 'ctrlc-convert' 'convert'
    $x = Read-Result $run.Dir
    $kinds = Join-Sorted ($x.Files | ForEach-Object { $_.'結果' } | Select-Object -Unique)
    Assert-Equal '変換の途中: Ctrl+C をキーの入力として受け取り（合図にしない）、そのキーで止まる（終了コード 1）' ('{0}/{1}/{2}/{3}' -f $run.Sent, $run.KeyMode, $run.Code, $kinds) 'ok/True/1/変換した,変換しなかった'
    Assert-Equal '変換の途中: 書き込みの途中では止めない（失敗はない）・壊れたファイルも控えのないファイルもない' ('{0}/{1}' -f @($x.Files | Where-Object { $_.'結果' -eq '失敗' }).Count, (Get-ManyStates $x.Files $run.Dir)) '0/0/0'
    Assert-Equal '変換の途中: 残りの備考と、中断の行' ('{0}/{1}' -f (@($x.Files | Where-Object { $_.'結果' -eq '変換しなかった' } | ForEach-Object { $_.'備考' } | Select-Object -Unique) -join ','), (@($x.Lines | Where-Object { $_ -eq '中断しました。それまでに変換したファイルと、控えたファイルは残っています' }).Count)) '中断したため、変換していません/1'
    $run = Invoke-ConsoleStop 'ctrlc-scan' 'scan'
    $unchanged = @(1..1500 | Where-Object { (ConvertTo-Hex (Get-Bytes (Join-Path $many ('f{0:0000}.txt' -f $_)))) -ne (ConvertTo-Hex $manyText) }).Count
    Assert-Equal '変換の前（調べている途中）: Ctrl+C をキーの入力として受け取り、止まって、結果を残さず、ファイルを変えない' ('{0}/{1}/{2}/{3}/{4}' -f $run.Sent, $run.KeyMode, $run.Code, [bool]$run.Dir, $unchanged) 'ok/True/1/False/0'

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $det.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        $summaryText = $x.Lines -join "`n"
        Assert-Equal 'fi-FI で実行して保存できた' ($x.Files.Count -gt 0) 'True'
        Assert-Equal 'fi-FI でも対象の行の日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($summaryText, '対象: .*   \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   ')) 'True'
        Assert-Equal 'fi-FI でも所要時間とサイズの小数点は .' ('{0}/{1}' -f [regex]::IsMatch($summaryText, '所要 \d+\.\d秒'), $summaryText.Contains('101.0 MB')) 'True/True'
        Assert-Equal 'fi-FI でも files.csv のサイズは 3 桁区切り（,）' (Get-Field $x.Files $det 'big.log' 'サイズ(バイト)') '105,906,176'
        Assert-Equal 'fi-FI でも files.csv の日時は yyyy/MM/dd HH:mm:ss' ((Get-Field $x.Files $det 'utf8.txt' '最終更新日時') -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }

    Write-Section 'レビューで見つかった問題の再発防止'
    # 1. Shift_JIS のまま改行だけを変えるときは、バイトのまま置き換える（上の「-LineEnding LF」で確かめている）
    # 2. 確認を待っている間に追記されたファイルを上書きしない（上の「確認を待つ間に変わったファイル」で確かめている）
    # 3. 変換できないファイルがあるときは「そろっています」と言わない（上の「そろった後」で確かめている）
    # 4. 長い保存先でも summary.txt と CSV を保存する（上の「260 文字を超えるパス」で確かめている）
    '  （ほかの項目で確かめている）'
} finally {
    # 開いたままのファイルを閉じ、アクセス拒否を戻してから消す
    foreach ($h in $holds) { try { $h.Dispose() } catch { } }
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    $global:OutputEncoding = $savedOutputEncoding
    if ($savedInputEncoding) { try { [Console]::InputEncoding = $savedInputEncoding } catch { } }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            # ジャンクションの先（det 自身）を消さないよう、先にジャンクションだけを外す。読み取り専用は外してから消す
            if ([IO.Directory]::Exists("$det\jloop")) { [IO.Directory]::Delete("$det\jloop") }
            Get-ChildItem -LiteralPath ('\\?\' + $workRoot) -Recurse -File -Force -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Attributes = 'Normal' } catch { } }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

'  END'
exit $script:failCount
