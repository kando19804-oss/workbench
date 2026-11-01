<#
.SYNOPSIS
    Convert-TextEncoding の回帰テストです。

.DESCRIPTION
    一時フォルダに、文字コードと改行の違うダミーのファイル（UTF-8 の BOM あり／なし、UTF-16、Shift_JIS、ASCII、
    EUC-JP、JIS、バイナリ、絵文字を含むもの、改行の混在、読み取り専用・隠しファイル、260 文字を超えるパス、
    名前に [ ] を含むフォルダ、ジャンクションのループ、読み取り禁止のフォルダ、100 MB を超えるファイルなど）を作り、
    Convert-TextEncoding で判定と変換を行って期待値と照合します。ダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダにだけアクセス拒否を設定します（終了時に戻します）。

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

$tool     = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Convert-TextEncoding\Convert-TextEncoding.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Convert-TextEncoding_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tree     = Join-Path $workRoot 'tree'
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

# 260 文字を超えるパスも作れるように、\\?\ を付けて書く
function New-TestFile([string]$Rel, [byte[]]$Bytes) {
    $full = Join-Path $tree $Rel
    [void][IO.Directory]::CreateDirectory('\\?\' + (Split-Path $full))
    [IO.File]::WriteAllBytes('\\?\' + $full, $Bytes)
}

function Get-FileBytes([string]$Rel) { [IO.File]::ReadAllBytes('\\?\' + (Join-Path $tree $Rel)) }

function ConvertTo-Hex([byte[]]$Bytes) { [BitConverter]::ToString($Bytes) }

# 一時フォルダのダミーを、元の内容で作り直す（変換のテストのたびに、同じ状態から始めるため）
function Reset-Tree {
    # 読み取り専用のファイルは書き直せないため、先に属性を外す
    foreach ($key in $attributes.Keys) {
        $full = '\\?\' + (Join-Path $tree $key)
        if ([IO.File]::Exists($full)) { [IO.File]::SetAttributes($full, [IO.FileAttributes]::Normal) }
    }
    foreach ($key in $original.Keys) { New-TestFile $key $original[$key] }
    foreach ($key in $attributes.Keys) { [IO.File]::SetAttributes('\\?\' + (Join-Path $tree $key), $attributes[$key]) }
}

function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$InputText = $null) {
    $out = Join-Path $workRoot "out_$Label"
    if ($null -ne $InputText) {
        $log = ConvertTo-FlatText ($InputText | & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String)
    } else {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String)
    }
    $code = $LASTEXITCODE
    $dir  = $null
    if (Test-Path -LiteralPath $out) {
        $found = @(Get-ChildItem -LiteralPath $out -Directory -Filter 'TextEncoding_*')
        if ($found.Count -gt 0) { $dir = $found[0].FullName }
    }
    [pscustomobject]@{ Code = $code; Log = $log; Dir = $dir }
}

function Read-Result([string]$Dir) {
    function Read-CsvOrEmpty([string]$Name) {
        if (-not $Dir) { return @() }
        $file = Join-Path $Dir $Name
        if (-not (Test-Path -LiteralPath $file)) { return @() }
        @(Import-Csv -LiteralPath $file -Encoding UTF8)
    }
    [pscustomobject]@{ Files = @(Read-CsvOrEmpty 'files.csv'); Skipped = @(Read-CsvOrEmpty 'skipped.csv') }
}

# files.csv の、ダミーのフォルダから見たファイルの行
function Get-Field($Rows, [string]$Rel, [string]$Column) {
    (@($Rows | Where-Object { $_.'ファイル' -eq (Join-Path $tree $Rel) }) | ForEach-Object { $_.$Column }) -join '|'
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $log = ConvertTo-FlatText (& cmd.exe /d /c $batch 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$userSid   = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$deniedDir = Join-Path $tree 'denied'
$script:denyApplied = $false
$sjis   = [Text.Encoding]::GetEncoding(932)
$utf8   = New-Object Text.UTF8Encoding $false
$bom    = [byte[]](0xEF, 0xBB, 0xBF)
$longRel = ('L' * 120) + '\' + ('M' * 120) + '\long.txt'

# ダミーのファイル（相対パス → 中身）
$original = [ordered]@{}
$original['utf8bom.txt']      = $bom + $utf8.GetBytes("日本語`r`nです`r`n")
$original['utf8.txt']         = $utf8.GetBytes("日本語`nです`n")
$original['sjis.bat']         = $sjis.GetBytes("@echo off`r`necho 日本語`r`n")
$original['ascii.csv']        = $utf8.GetBytes("a,b`r`n1,2`r`n")
$original['mixed.txt']        = $utf8.GetBytes("a`r`nb`nc`r")
$original['noline.txt']       = $utf8.GetBytes('abc')
$original['utf16.txt']        = [byte[]](0xFF, 0xFE) + [Text.Encoding]::Unicode.GetBytes("日本語`r`n")
$original['bin.txt']          = [byte[]](0x41, 0, 0x42)
$original['euc.txt']          = [byte[]](0xA4, 0xB3, 0xA4, 0xEC, 0xA4, 0xCF, 0xC6, 0xFC, 0xCB, 0xDC, 0xB8, 0xEC, 0xA4, 0xCE, 0xCA, 0xB8, 0xBE, 0xCF, 0xA4, 0xC7, 0xA4, 0xB9, 0x0A)   # 「これは日本語の文章です」
$original['jis.txt']          = [byte[]](0x1B, 0x24, 0x42, 0x46, 0x7C, 0x4B, 0x5C, 0x38, 0x6C, 0x1B, 0x28, 0x42, 0x0D, 0x0A)                                                  # 「日本語」
$original['emoji.txt']        = $utf8.GetBytes("line1`nスマイル ") + [byte[]](0xF0, 0x9F, 0x99, 0x82) + $utf8.GetBytes("`n")
$original['sub\run.ps1']      = $sjis.GetBytes("Write-Host '日本語'`r`n")
$original['日本語 フォルダ\メモ.txt'] = $utf8.GetBytes("メモ`n")
$original['[a]\f.txt']        = $utf8.GetBytes("x`n")
$original['readonly.txt']     = $utf8.GetBytes("読み取り専用`n")
$original['hidden.txt']       = $utf8.GetBytes("隠し`n")
$original['nec.txt']          = [byte[]](0x87, 0x9A, 0x0A)   # Shift_JIS の重複した文字（∵ の NEC 特殊文字）
$original['skip.dat']         = $utf8.GetBytes("filter で対象外`n")
$original['TextEncoding_old\x.txt'] = $utf8.GetBytes("前回の結果`n")
$original[$longRel]           = $utf8.GetBytes("長いパス`n")
$original['denied\secret.txt'] = $utf8.GetBytes("秘密`n")
$attributes = @{ 'readonly.txt' = [IO.FileAttributes]::ReadOnly; 'hidden.txt' = [IO.FileAttributes]::Hidden }

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    Reset-Tree
    # 100 MB を超えるファイル（先頭の 9 KB だけが文字で、残りは 0）。先頭の 1 MB だけで判定する
    $bigPath = Join-Path $tree 'big.log'
    $stream = [IO.File]::Create($bigPath)
    $head = $utf8.GetBytes(('x' * 9215) + "`n")
    $stream.Write($head, 0, $head.Length)
    $stream.SetLength(101MB)
    $stream.Close()
    $null = New-Item -ItemType Junction -Path "$tree\jloop" -Target $tree
    $null = & icacls.exe $deniedDir /deny "*${userSid}:(RD)" 2>&1
    $script:denyApplied = ($LASTEXITCODE -eq 0)
    Assert-Equal 'アクセス拒否の設定' $script:denyApplied 'True'

    Write-Section '判定（既定の拡張子）'
    $run = Invoke-Tool 'detect' @('-Path', $tree)
    Assert-Equal '終了コード' $run.Code 0
    $x = Read-Result $run.Dir
    $expected = [ordered]@{
        'utf8bom.txt' = 'UTF-8（BOM あり）/CRLF'; 'utf8.txt' = 'UTF-8（BOM なし）/LF'; 'sjis.bat' = 'Shift_JIS/CRLF'
        'ascii.csv' = 'ASCII のみ/CRLF'; 'mixed.txt' = 'ASCII のみ/混在'; 'noline.txt' = 'ASCII のみ/なし'
        'utf16.txt' = 'UTF-16（BOM あり）/CRLF'; 'bin.txt' = 'バイナリ/'; 'euc.txt' = 'EUC-JP/LF'; 'jis.txt' = 'JIS/CRLF'
        'emoji.txt' = 'UTF-8（BOM なし）/LF'; 'sub\run.ps1' = 'Shift_JIS/CRLF'; '日本語 フォルダ\メモ.txt' = 'UTF-8（BOM なし）/LF'
        '[a]\f.txt' = 'UTF-8（BOM なし）/LF'; $longRel = 'UTF-8（BOM なし）/LF'; 'big.log' = 'ASCII のみ/LF'
    }
    foreach ($key in $expected.Keys) {
        Assert-Equal ('判定: {0}' -f $key) ('{0}/{1}' -f (Get-Field $x.Files $key '文字コード'), (Get-Field $x.Files $key '改行')) $expected[$key]
    }
    # .NET Framework（5.1）の CP932 は NEC 特殊文字の重複を読めるが、.NET（7）は厳密に読むと受け付けない
    $necExpected = if ($hostTag -eq '51') { 'Shift_JIS' } else { '判定できない' }
    Assert-Equal ('判定: nec.txt（{0}）' -f $hostTag) (Get-Field $x.Files 'nec.txt' '文字コード') $necExpected
    Assert-Equal '大きなファイルは先頭で判定' (Get-Field $x.Files 'big.log' '備考') '先頭の 1 MB で判定'
    Assert-Equal '-Filter に当たらないファイルは調べない' (Get-Field $x.Files 'skip.dat' '文字コード') ''
    Assert-Equal '結果のフォルダの中は調べない' (Get-Field $x.Files 'TextEncoding_old\x.txt' '文字コード') ''
    Assert-Equal 'ループするジャンクションをたどらない' ((@($x.Skipped | Where-Object { $_.'理由' -eq 'リンク' }) | ForEach-Object { $_.'パス' }) -join ',') "$tree\jloop"
    if ($script:denyApplied) {
        Assert-Equal '権限なし' ((@($x.Skipped | Where-Object { $_.'理由' -eq '権限なし' }) | ForEach-Object { $_.'パス' }) -join ',') $deniedDir
    } else {
        Write-Skip '権限なし' 'アクセス拒否を設定できなかったため'
    }
    Assert-Equal '判定だけのときは、ファイルを変えない' (ConvertTo-Hex (Get-FileBytes 'utf8.txt')) (ConvertTo-Hex $original['utf8.txt'])
    Assert-Equal '判定だけのときは、結果の列が空欄' (Get-Field $x.Files 'utf8.txt' '結果') ''
    Assert-Equal '内訳の見出し' $run.Log.Contains('■ 文字コードと改行の内訳') 'True'

    Write-Section '-Filter "*"、-Exclude'
    $run = Invoke-Tool 'filter' @('-Path', $tree, '-Filter', '*', '-Exclude', '*\sub,*\big.log')
    $x = Read-Result $run.Dir
    Assert-Equal '-Filter "*" はすべて調べる' (Get-Field $x.Files 'skip.dat' '文字コード') 'UTF-8（BOM なし）'
    Assert-Equal '-Exclude のフォルダ' (Get-Field $x.Files 'sub\run.ps1' '文字コード') ''
    Assert-Equal '-Exclude のファイル' (Get-Field $x.Files 'big.log' '文字コード') ''
    Assert-Equal '除外の件数' $run.Log.Contains('除外 2 件') 'True'

    Write-Section '変換: -To SJIS -LineEnding CRLF -Force'
    $run = Invoke-Tool 'sjis' @('-Path', $tree, '-To', 'SJIS', '-LineEnding', 'CRLF', '-Force')
    Assert-Equal '終了コード（変換できないファイルがある）' $run.Code 1
    $x = Read-Result $run.Dir
    Assert-Equal 'UTF-8 → Shift_JIS・CRLF' (ConvertTo-Hex (Get-FileBytes 'utf8.txt')) (ConvertTo-Hex $sjis.GetBytes("日本語`r`nです`r`n"))
    Assert-Equal 'UTF-8 の BOM を外す' (ConvertTo-Hex (Get-FileBytes 'utf8bom.txt')) (ConvertTo-Hex $sjis.GetBytes("日本語`r`nです`r`n"))
    Assert-Equal 'UTF-16 → Shift_JIS' (ConvertTo-Hex (Get-FileBytes 'utf16.txt')) (ConvertTo-Hex $sjis.GetBytes("日本語`r`n"))
    Assert-Equal '改行の混在をそろえる' (ConvertTo-Hex (Get-FileBytes 'mixed.txt')) (ConvertTo-Hex $utf8.GetBytes("a`r`nb`r`nc`r`n"))
    Assert-Equal '備考に変換の内容' (Get-Field $x.Files 'utf8.txt' '備考') 'UTF-8（BOM なし） → Shift_JIS・改行 LF → CRLF'
    Assert-Equal 'すでに同じ' (Get-Field $x.Files 'sjis.bat' '結果') 'すでに同じ'
    Assert-Equal '対象外（EUC-JP・JIS・バイナリ・大きなファイル）' ((@('euc.txt', 'jis.txt', 'bin.txt', 'big.log') | ForEach-Object { Get-Field $x.Files $_ '結果' }) -join ',') '対象外,対象外,対象外,対象外'
    Assert-Equal 'Shift_JIS にない文字' ('{0}/{1}' -f (Get-Field $x.Files 'emoji.txt' '結果'), (Get-Field $x.Files 'emoji.txt' '備考')) ('変換できない/Shift_JIS にない文字: {0}（2 行目）' -f $utf8.GetString([byte[]](0xF0, 0x9F, 0x99, 0x82)))
    Assert-Equal '変換できないファイルは変えない' (ConvertTo-Hex (Get-FileBytes 'emoji.txt')) (ConvertTo-Hex $original['emoji.txt'])
    Assert-Equal '読み取り専用は変換しない' ('{0}/{1}' -f (Get-Field $x.Files 'readonly.txt' '結果'), (ConvertTo-Hex (Get-FileBytes 'readonly.txt'))) ('変換できない/{0}' -f (ConvertTo-Hex $original['readonly.txt']))
    Assert-Equal '隠しファイルも変換し、隠しのまま' ('{0}/{1}' -f (Get-Field $x.Files 'hidden.txt' '結果'), (([IO.File]::GetAttributes((Join-Path $tree 'hidden.txt')) -band [IO.FileAttributes]::Hidden) -ne 0)) '変換した/True'
    Assert-Equal '260 文字を超えるパスも変換' (ConvertTo-Hex (Get-FileBytes $longRel)) (ConvertTo-Hex $sjis.GetBytes("長いパス`r`n"))
    Assert-Equal '名前に [ ] を含むフォルダ' (Get-Field $x.Files '[a]\f.txt' '結果') '変換した'
    Assert-Equal '控え: フォルダの構成で、元の中身' (ConvertTo-Hex ([IO.File]::ReadAllBytes((Join-Path $run.Dir 'backup\sub\run.ps1')))) (ConvertTo-Hex $original['sub\run.ps1'])
    Assert-Equal '控え: 長いパス' (Test-Path -LiteralPath ('\\?\' + (Join-Path $run.Dir ('backup\' + $longRel)))) 'True'
    Assert-Equal '変換した後のサイズ' (Get-Field $x.Files 'utf8.txt' 'サイズ(バイト)') ($sjis.GetBytes("日本語`r`nです`r`n").Length)
    # レビューで見つかった問題: Shift_JIS のまま改行だけを変えるときに、重複した文字のバイトが変わっていた
    if ($hostTag -eq '51') {
        Assert-Equal 'Shift_JIS の重複した文字のバイトを変えない' (ConvertTo-Hex (Get-FileBytes 'nec.txt')) '87-9A-0D-0A'
    } else {
        Assert-Equal '7 では重複した文字のファイルは対象外' ('{0}/{1}' -f (Get-Field $x.Files 'nec.txt' '結果'), (ConvertTo-Hex (Get-FileBytes 'nec.txt'))) '対象外/87-9A-0A'
    }

    Write-Section 'そろった後にもう一度'
    $run = Invoke-Tool 'again' @('-Path', $tree, '-To', 'SJIS', '-LineEnding', 'CRLF', '-Force')
    Assert-Equal '変換するファイルがない' $run.Log.Contains('変換するファイル 0 件') 'True'
    Assert-Equal '変換できないファイルがあるときは「そろっています」と言わない' $run.Log.Contains('すべてそろっています') 'False'

    Write-Section '確認で N を選ぶ'
    Reset-Tree
    $run = Invoke-Tool 'no' @('-Path', $tree, '-To', 'UTF8') 'N'
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '変換しない' (ConvertTo-Hex (Get-FileBytes 'sjis.bat')) (ConvertTo-Hex $original['sjis.bat'])
    Assert-Equal '結果' (Get-Field (Read-Result $run.Dir).Files 'sjis.bat' '結果') '変換しなかった'
    Assert-Equal '.ps1 の注意' $run.Log.Contains('※ .ps1 を BOM なしの UTF-8 にすると') 'True'
    Assert-Equal '.bat の注意' $run.Log.Contains('※ .bat・.cmd を UTF-8 にすると') 'True'

    Write-Section 'レビューで見つかった問題: 確認を待っている間に追記されたファイルを上書きしない'
    Reset-Tree
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = (Get-Command $Exe).Source
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Path "{1}" -LineEnding CRLF -OutputFolder "{2}"' -f $tool, $tree, (Join-Path $workRoot 'out_append')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $proc = [Diagnostics.Process]::Start($psi)
    $outputTask = $proc.StandardOutput.ReadToEndAsync()
    Start-Sleep -Seconds 5
    [IO.File]::AppendAllText((Join-Path $tree 'utf8.txt'), "追記`n")
    $proc.StandardInput.WriteLine('Y')
    $proc.StandardInput.Close()
    $proc.WaitForExit()
    $dir = @(Get-ChildItem -LiteralPath (Join-Path $workRoot 'out_append') -Directory)[0].FullName
    $x = Read-Result $dir
    Assert-Equal '追記されたファイルは失敗にする' ('{0}/{1}' -f (Get-Field $x.Files 'utf8.txt' '結果'), (Get-Field $x.Files 'utf8.txt' '備考')) '失敗/調べた後に変更されたため、変換していません'
    Assert-Equal '追記した分が残っている' ($utf8.GetString((Get-FileBytes 'utf8.txt')).EndsWith("追記`n")) 'True'
    Assert-Equal 'ほかのファイルは変換する' (Get-Field $x.Files 'mixed.txt' '結果') '変換した'

    Write-Section '止まるとき・保存できないとき'
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path (Join-Path $workRoot 'no\such') 2>&1 | Out-String)
    Assert-Equal '存在しないパス' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('指定したフォルダかファイルが見つかりません')) '1/True'
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $tree -To UTF8 -NoSave 2>&1 | Out-String)
    Assert-Equal '変換と -NoSave' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('-NoSave は指定できません')) '1/True'
    $noSaveOut = Join-Path $workRoot 'out_nosave'
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $tree -NoSave -OutputFolder $noSaveOut 2>&1
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}' -f $LASTEXITCODE, (Test-Path -LiteralPath $noSaveOut)) '0/False'
    $usedLetters = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    $freeLetter  = @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $usedLetters -notcontains $_ } | Select-Object -First 1
    if ($freeLetter) {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $tree -OutputFolder "${freeLetter}:\nope" 2>&1 | Out-String)
        Assert-Equal '書き込めない保存先: 判定だけなら続ける' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('保存先に書き込めないため、結果は保存しません')) '0/True'
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $tree -To UTF8BOM -Force -OutputFolder "${freeLetter}:\nope" 2>&1 | Out-String)
        Assert-Equal '書き込めない保存先: 変換しない' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('変換しません')) '1/True'
        Assert-Equal '書き込めない保存先: ファイルを変えない' (ConvertTo-Hex (Get-FileBytes 'sjis.bat')) (ConvertTo-Hex $original['sjis.bat'])
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }

    Write-Section 'コマンドプロンプトから起動'
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -NoSave -Filter "*.csv,*.bat" -Path "{2}\"' -f $Exe, $tool, $tree)
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal 'カンマ区切りの -Filter と末尾の \' ([regex]::Match($run.Log, 'ファイル \S+ 件').Value) 'ファイル 2 件'
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -Path "{2}\" -NoSave' -f $Exe, $tool, $tree)
    Assert-Equal '"パス\" の後ろに引数: 終了コード' $run.Code 1
    Assert-Equal '"パス\" の後ろに引数: ヒント' ([regex]::IsMatch($run.Log, '最後に \\ を付けないでください')) 'True'
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "{1}" -Path "{2}" -To UTF8BOM -OutputFolder "{3}"' -f $Exe, $tool, $tree, (Join-Path $workRoot 'out_nonint'))
    Assert-Equal '確認できないときは変換しない' ('{0}/{1}' -f $run.Code, $run.Log.Contains('確認できないため、変換しません')) '1/True'
    Assert-Equal '確認できないとき: ファイルを変えない' (ConvertTo-Hex (Get-FileBytes 'sjis.bat')) (ConvertTo-Hex $original['sjis.bat'])
} finally {
    # アクセス拒否を戻してから消す
    if ($script:denyApplied) { $null = & icacls.exe $deniedDir /remove:d "*$userSid" 2>&1 }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        # ループするジャンクションは、リンク先をたどらないように先にリンクだけ消す。読み取り専用は外してから消す
        try {
            $junction = Join-Path $tree 'jloop'
            if (Test-Path -LiteralPath $junction) { [IO.Directory]::Delete($junction) }
            Get-ChildItem -LiteralPath $workRoot -Recurse -File -Force -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Attributes = 'Normal' } catch { } }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

exit $script:failCount
