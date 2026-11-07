<#
.SYNOPSIS
    Search-Log の回帰テストです。

.DESCRIPTION
    一時フォルダにダミーのログ（文字コードの違うもの、前後の行の確かめ用、フォルダの階層、リンク、
    日本語や [ ] を含む名前、260 文字を超えるパス、Excel が数式や日付に変える値を含むもの）を作り、
    Search-Log を実行して期待値と照合します。
    「権限なし」を再現するため、テストが作ったファイルとフォルダにだけアクセス拒否を設定します（終了時に戻します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Search-Log を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Search-Log\Search-Log.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Search-Log_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Search-Log.ps1'
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
$csvHead   = '"ファイル","行番号","一致した語","一致した行","文字コード","最終更新日時"'
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

# Log には、画面への出力から改行を取り除いたものを入れる
function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out  = Join-Path $workRoot "out\$Label"
    $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = $log; Dir = (Get-RunDir $out) }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行
function Invoke-Plain([string[]]$ToolArgs) {
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 相対パスの基準を確かめるため、-OutputFolder を足さず、指定したフォルダから起動する
function Invoke-From([string]$WorkDir, [string[]]$ToolArgs) {
    Push-Location -LiteralPath $WorkDir
    try {
        $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
        $code = $LASTEXITCODE
    } finally {
        Pop-Location
    }
    [pscustomobject]@{ Code = $code; Log = $log }
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { $utf8NoBom } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
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
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'SearchLog_*' -ErrorAction SilentlyContinue | Sort-Object Name)
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
    function Read-LinesOrEmpty([string]$Name) {
        if (-not $Dir) { return @() }
        $file = Join-Path $Dir $Name
        if (-not (Test-Path -LiteralPath $file)) { return @() }
        @([IO.File]::ReadAllLines($file))
    }
    [pscustomobject]@{
        Matches = @(Read-CsvOrEmpty 'matches.csv')
        Skipped = @(Read-CsvOrEmpty 'skipped.csv')
        Text    = @(Read-LinesOrEmpty 'matches.txt')
        Lines   = @(Read-LinesOrEmpty 'summary.txt')
    }
}

# matches.csv の「ファイル:行番号」を、指定したフォルダからの相対パスで並べる
function Get-MatchKeys($Result, [string]$Base) {
    (@($Result.Matches | ForEach-Object { '{0}:{1}' -f $_.'ファイル'.Replace($Base + '\', ''), $_.'行番号' })) -join ' | '
}

function Get-SkipKeys($Result, [string]$Base) {
    (@($Result.Skipped | ForEach-Object { '{0} {1}' -f $_.'理由', $_.'パス'.Replace($Base + '\', '') })) -join ' | '
}

function Join-Sorted($Items) {
    $array = [string[]]@($Items | Where-Object { $null -ne $_ })
    [Array]::Sort($array, [StringComparer]::Ordinal)
    $array -join ','
}

function Get-ResultFileNames([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { $_.Name })
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

# 画面の結果の行（例: 結果: 一致 5 行（3 ファイル）   調べたファイル 42 件）。所要時間は除く
function Get-ResultText([string]$Log) { [regex]::Match($Log, '結果: .+?件(?=   所要)').Value }

# 結果の行の下に出る注記（スキップ・リンク・除外・最終更新）
function Get-NoteText([string]$Log) { [regex]::Match($Log, '所要 [\d.]+秒(.*?)(?=■ 語ごとの件数|保存先|$)').Groups[1].Value.Trim() }

function Get-HeaderLine([string]$Log, [string]$Prefix) { [regex]::Match($Log, [regex]::Escape($Prefix) + '[^■]*?(?=(ログ横断|対象: |設定ファイル: |検索: |除く: |条件: |除外: |■|結果: |$))').Value.Trim() }

# 画面に出た一致の行の数（> と行番号で始まる部分を数える）
function Get-ShownCount([string]$Log, [string]$Word) { [regex]::Matches($Log, '>\s*\d+  ' + [regex]::Escape($Word)).Count }

# 260 文字を超えるパスにも書けるように、\\?\ を付けて書く
function New-TextFile([string]$FilePath, [string[]]$Lines, [Text.Encoding]$Encoding = $utf8NoBom) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllText('\\?\' + $FilePath, ($Lines -join "`r`n") + "`r`n", $Encoding)
    $FilePath
}

function New-ByteFile([string]$FilePath, [byte[]]$Bytes) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllBytes('\\?\' + $FilePath, $Bytes)
    $FilePath
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

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$userSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied  = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したパス（終了時に戻す）
$holds   = New-Object 'System.Collections.Generic.List[object]'   # テストの中で開いたままにしたファイル（終了時に閉じる）

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    # 前後の行の確かめ用。5・7・20・30 行目に一致する
    $ctxDir  = Join-Path $workRoot 'ctx'
    $ctxFile = New-TextFile "$ctxDir\ctx.log" @(1..30 | ForEach-Object { if (5, 7, 20, 30 -contains $_) { 'line {0:00} hit' -f $_ } else { 'line {0:00}' -f $_ } })

    # フォルダの階層・ファイル名の条件・リンク・権限なし
    $tree = Join-Path $workRoot 'tree'
    foreach ($name in 'a.log', 'B.txt', '_x.log', 'a2.log', 'data.csv', 'noext', 'sub\deep\c.log', '[br]\d.log', '日本語 フォルダ\ログ.log', 'SearchLog_old\old.log', 'archive\e.log', 'denied\f.log', 'noread.log') {
        $null = New-TextFile "$tree\$name" @('first line', ('ERR {0}' -f $name), 'last line')
    }
    $null = & cmd.exe /d /c mklink /J "$tree\loop" "$tree" 2>&1
    $hasJunction = [IO.Directory]::Exists("$tree\loop")
    $deniedOk = (Add-Deny "$tree\denied" '(RD)') -and (Add-Deny "$tree\noread.log" '(R)')

    Write-Section '前後の行（-Context）'
    $run = Invoke-Tool 'ctx' @('-Path', $ctxDir, '-Pattern', 'hit')
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード（一致あり）' $run.Code 0
    Assert-Equal '保存するファイル' (Get-ResultFileNames $run.Dir) 'matches.csv,matches.txt,skipped.csv,summary.txt'
    Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'matches.csv=True,matches.txt=True,skipped.csv=True,summary.txt=True'
    Assert-Equal 'matches.csv の見出し' (Get-FirstLine $run.Dir 'matches.csv') $csvHead
    Assert-Equal 'skipped.csv の見出し' (Get-FirstLine $run.Dir 'skipped.csv') '"理由","パス","詳細"'
    $expected = @(
        "■ $ctxFile（UTF-8）",
        '      3  line 03', '      4  line 04', '>     5  line 05 hit', '      6  line 06', '>     7  line 07 hit', '      8  line 08', '      9  line 09',
        '  --',
        '     18  line 18', '     19  line 19', '>    20  line 20 hit', '     21  line 21', '     22  line 22',
        '  --',
        '     28  line 28', '     29  line 29', '>    30  line 30 hit',
        ''
    )
    Assert-Equal 'matches.txt（前後 2 行・重なりは 1 つに・続かないところに --）' ($x.Text -join ' | ') ($expected -join ' | ')
    Assert-Equal '画面にも同じ形' $run.Log.Contains(($expected[1..9] -join '')) 'True'
    Assert-Equal 'matches.csv' (Get-MatchKeys $x $ctxDir) 'ctx.log:5 | ctx.log:7 | ctx.log:20 | ctx.log:30'
    $row = $x.Matches | Select-Object -First 1
    Assert-Equal 'matches.csv の 1 行' ('{0}|{1}|{2}|{3}' -f $row.'一致した語', $row.'一致した行', $row.'文字コード', ($row.'最終更新日時' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$')) 'hit|line 05 hit|UTF-8|True'
    Assert-Equal '結果の行' (Get-ResultText $run.Log) '結果: 一致 4 行（1 ファイル）   調べたファイル 1 件'
    Assert-Equal 'summary の 1 行目' $x.Lines[0] 'ログ横断検索 (Search-Log 1.0.0)'
    Assert-Equal 'summary の対象の行' ([regex]::IsMatch($x.Lines[1], ('^対象: {0}（\*\.log, \*\.txt）   \d{{4}}-\d{{2}}-\d{{2}} \d{{2}}:\d{{2}}:\d{{2}}   {1}$' -f [regex]::Escape($ctxDir), [regex]::Escape([Environment]::MachineName)))) 'True'
    Assert-Equal 'summary の検索の行' $x.Lines[2] '検索: hit（大文字と小文字を区別しない）   前 2 行・後ろ 2 行'
    Assert-Equal 'summary のファイルごとの件数' ($x.Lines[-1]) ('         4 行  UTF-8      {0}' -f $ctxFile)
    Assert-Equal '画面に保存先' $run.Log.Contains(('保存先: {0}' -f $run.Dir)) 'True'

    $run = Invoke-Tool 'ctx-1-3' @('-Path', $ctxDir, '-Pattern', 'hit', '-Context', '1,3')
    $nos = @((Read-Result $run.Dir).Text | Where-Object { $_ -match '^[ >]\s*\d+  ' } | ForEach-Object { [int]($_.Substring(1, 6)) })
    Assert-Equal '-Context "1,3"（前 1 行・後ろ 3 行）' ($nos -join ',') '4,5,6,7,8,9,10,19,20,21,22,23,29,30'
    $run = Invoke-Tool 'ctx-0' @('-Path', $ctxDir, '-Pattern', 'hit', '-Context', '0')
    Assert-Equal '-Context 0（一致した行だけ。間に --）' (((Read-Result $run.Dir).Text | Select-Object -Skip 1) -join ' | ') '>     5  line 05 hit |   -- | >     7  line 07 hit |   -- | >    20  line 20 hit |   -- | >    30  line 30 hit | '
    # PowerShell から引用符なしで 2,5 と書くと、配列として届く
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $ctxDir -Pattern hit -Context 2,5 -NoSave 2>&1) | ForEach-Object { "$_" }) -join "`n")
    Assert-Equal 'PowerShell から引用符なしの -Context 2,5' ($log.Contains('前 2 行・後ろ 5 行') -and $log.Contains('     12  line 12')) 'True'
    foreach ($bad in '51', '2,51', 'a', '-1', '1,2,3') {
        $run = Invoke-Plain @('-Path', $ctxDir, '-Pattern', 'hit', '-Context', $bad, '-NoSave')
        Assert-Equal ('-Context {0}: 止まる' -f $bad) ('{0}/{1}' -f $run.Code, $run.Log.Contains('-Context は 0～50 の数値か')) '1/True'
    }
    $run = Invoke-Plain @('hit', '-Path', $ctxDir, '-NoSave')
    Assert-Equal '-Pattern は最初の引数として書ける' ('{0}/{1}' -f $run.Code, (Get-ResultText $run.Log)) '0/結果: 一致 4 行（1 ファイル）   調べたファイル 1 件'

    Write-Section '文字コード'
    $encDir = Join-Path $workRoot 'enc'
    $line   = 'あいう タイムアウト発生 ①'
    $null = New-TextFile "$encDir\utf8.log" @('first', "$line utf8") $utf8NoBom
    $null = New-TextFile "$encDir\utf8bom.log" @('first', "$line utf8bom") $utf8Bom
    $null = New-TextFile "$encDir\sjis.log" @('first', 'あいう タイムアウト発生 sjis') ([Text.Encoding]::GetEncoding(932))
    $null = New-TextFile "$encDir\utf16le.log" @('first', "$line utf16le") ([Text.Encoding]::Unicode)
    $null = New-TextFile "$encDir\utf16be.log" @('first', "$line utf16be") ([Text.Encoding]::BigEndianUnicode)
    $null = New-ByteFile "$encDir\nobom16.log" ((New-Object Text.UnicodeEncoding $false, $false).GetBytes("first`r`n$line nobom16`r`n"))
    $null = New-ByteFile "$encDir\bin.log" ([byte[]](0x41, 0x42, 0x00, 0x43) + $utf8NoBom.GetBytes("`r`n$line bin`r`n"))
    $null = New-ByteFile "$encDir\empty.log" ([byte[]]@())
    # 文字コードの判定に読む 1 MB の境目で、UTF-8 の文字が切れる（切れた文字を Shift_JIS と取り違えない）
    $head = New-Object Text.StringBuilder
    while ($head.Length -lt 1048575 - 100) { [void]$head.Append(('x' * 98) + "`r`n") }
    [void]$head.Append('y' * (1048575 - $head.Length))
    $null = New-ByteFile "$encDir\boundary.log" ($utf8NoBom.GetBytes($head.ToString() + "あ`r`n$line boundary`r`n"))
    $run = Invoke-Tool 'enc' @('-Path', $encDir, '-Pattern', 'タイムアウト発生')
    $x = Read-Result $run.Dir
    Assert-Equal '判定した文字コード' ((@($x.Matches | ForEach-Object { '{0}={1}' -f (Split-Path $_.'ファイル' -Leaf), $_.'文字コード' })) -join ', ') 'boundary.log=UTF-8, sjis.log=Shift_JIS, utf16be.log=UTF-16, utf16le.log=UTF-16, utf8.log=UTF-8, utf8bom.log=UTF-8'
    Assert-Equal '日本語が化けない' ((@($x.Matches | ForEach-Object { $_.'一致した行' })) -join ' | ') "$line boundary | あいう タイムアウト発生 sjis | $line utf16be | $line utf16le | $line utf8 | $line utf8bom"
    Assert-Equal '行番号（1 MB の境目のあとも数える）' ((@($x.Matches | ForEach-Object { $_.'行番号' })) -join ',') ('{0},2,2,2,2,2' -f ((Get-Content -LiteralPath "$encDir\boundary.log" -Encoding UTF8 | Measure-Object).Count))
    Assert-Equal '見出しに文字コード' ($run.Log.Contains("■ $encDir\sjis.log（Shift_JIS）") -and $run.Log.Contains("■ $encDir\utf16le.log（UTF-16）")) 'True'
    Assert-Equal 'バイナリ（NUL を含む・BOM のない UTF-16）は飛ばす' (Get-SkipKeys $x $encDir) 'バイナリ bin.log | バイナリ nobom16.log'
    Assert-Equal '結果の行（空のファイルも調べた数に入る）' (Get-ResultText $run.Log) '結果: 一致 6 行（6 ファイル）   調べたファイル 7 件'
    Assert-Equal 'スキップの注記' (Get-NoteText $run.Log) 'スキップ 2 件（バイナリ）'

    Write-Section 'フォルダの階層・名前の条件・リンク・権限なし'
    $run = Invoke-Tool 'tree' @('-Path', $tree, '-Pattern', 'ERR')
    $x = Read-Result $run.Dir
    # パスの大文字を序数で比べた順（_ は英字より後、日本語はさらに後）。5.1 と 7 で同じ順になる
    Assert-Equal '探したファイルと順番（*.log と *.txt、サブフォルダも）' (Get-MatchKeys $x $tree) 'a.log:2 | a2.log:2 | archive\e.log:2 | B.txt:2 | sub\deep\c.log:2 | [br]\d.log:2 | _x.log:2 | 日本語 フォルダ\ログ.log:2'
    Assert-Equal '結果のフォルダ（SearchLog_*）の中は探さない' ($x.Matches | Where-Object { $_.'ファイル' -like '*SearchLog_old*' } | Measure-Object).Count 0
    $skips = @()
    if ($deniedOk) { $skips += '権限なし denied', '権限なし noread.log' }
    if ($hasJunction) { $skips += 'リンク loop' }
    Assert-Equal 'skipped.csv（権限なし・リンク）' (Join-Sorted (@($x.Skipped | ForEach-Object { '{0} {1}' -f $_.'理由', $_.'パス'.Replace($tree + '\', '') }))) (Join-Sorted $skips)
    $notes = @()
    if ($deniedOk) { $notes += 'スキップ 2 件（権限なし）' }
    if ($hasJunction) { $notes += 'たどらなかったリンク 1 件' }
    Assert-Equal '結果の行の下の注記' (Get-NoteText $run.Log) ($notes -join '  ')
    if (-not $deniedOk) { Write-Skip '権限なし' 'アクセス拒否を設定できなかったため' }
    if (-not $hasJunction) { Write-Skip 'リンク（ジャンクション）' 'ジャンクションを作れなかったため' }
    else { Assert-Equal 'リンクの詳細' (@($x.Skipped | Where-Object { $_.'理由' -eq 'リンク' } | ForEach-Object { $_.'詳細' }) -join ',') 'ジャンクション' }
    Assert-Equal 'summary のファイルごとの件数（8 ファイル）' @($x.Lines | Where-Object { $_ -match '^\s+1 行  UTF-8' }).Count 8

    $fileLink = Join-Path $tree 'filelink.log'
    $linkMade = $false
    try { $null = New-Item -ItemType SymbolicLink -Path $fileLink -Target "$tree\a.log" -ErrorAction Stop; $linkMade = $true } catch { }
    if ($linkMade) {
        $run = Invoke-Plain @('-Path', $tree, '-Pattern', 'ERR a.log', '-NoSave')
        Assert-Equal 'ファイルのシンボリックリンクもたどらない' ('{0}/{1}' -f (Get-ResultText $run.Log), $run.Log.Contains('たどらなかったリンク 2 件')) '結果: 一致 1 行（1 ファイル）   調べたファイル 8 件/True'
        Remove-Item -LiteralPath $fileLink -Force
    } else {
        Write-Skip 'ファイルのシンボリックリンク' 'シンボリックリンクを作れないため（管理者か開発者モードが必要）'
    }

    $filterCases = @(
        @('*.csv', 'data.csv:2'),
        @('*.LOG', 'a.log:2 | a2.log:2 | archive\e.log:2 | sub\deep\c.log:2 | [br]\d.log:2 | _x.log:2 | 日本語 フォルダ\ログ.log:2'),
        @('?.log', 'a.log:2 | archive\e.log:2 | sub\deep\c.log:2 | [br]\d.log:2'),
        @('*.csv,noext', 'data.csv:2 | noext:2')
    )
    foreach ($case in $filterCases) {
        $run = Invoke-Tool 'filter' @('-Path', $tree, '-Pattern', 'ERR', '-Filter', $case[0])
        Assert-Equal ('-Filter {0}' -f $case[0]) (Get-MatchKeys (Read-Result $run.Dir) $tree) $case[1]
    }
    $run = Invoke-Tool 'exclude' @('-Path', $tree, '-Pattern', 'ERR', '-Exclude', '*\archive\,*\_x.log,*\[br]')
    $x = Read-Result $run.Dir
    Assert-Equal '-Exclude（[ ] は文字のまま・末尾の \ は無視）' (Get-MatchKeys $x $tree) 'a.log:2 | a2.log:2 | B.txt:2 | sub\deep\c.log:2 | 日本語 フォルダ\ログ.log:2'
    Assert-Equal '-Exclude: 除外の行' (Get-HeaderLine $run.Log '除外: ') '除外: *\archive, *\_x.log, *\[br]'
    Assert-Equal '-Exclude: skipped.csv' (Join-Sorted (@($x.Skipped | Where-Object { $_.'理由' -eq '除外指定' } | ForEach-Object { '{0}={1}' -f $_.'パス'.Replace($tree + '\', ''), $_.'詳細' }))) '[br]=*\[br],_x.log=*\_x.log,archive=*\archive'
    Assert-Equal '-Exclude: 注記' $run.Log.Contains('除外したフォルダ・ファイル 3 件') 'True'

    Write-Section '-Path の指定'
    $run = Invoke-Plain @('-Path', "$tree\data.csv", '-Pattern', 'ERR', '-NoSave')
    Assert-Equal 'ファイルを直接指定（-Filter は使わない）' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultText $run.Log), ([regex]::IsMatch($run.Log, ('対象: {0}   \d{{4}}' -f [regex]::Escape("$tree\data.csv"))))) '0/結果: 一致 1 行（1 ファイル）   調べたファイル 1 件/True'
    $run = Invoke-Tool 'multi' @('-Path', "$tree\a.log,$ctxDir,$workRoot\nosuch", '-Pattern', 'ERR|hit')
    $x = Read-Result $run.Dir
    Assert-Equal '複数の場所（カンマ区切り）・見つからない場所は飛ばす' ('{0}/{1}' -f $run.Code, (Get-MatchKeys $x $workRoot)) '0/tree\a.log:2 | ctx\ctx.log:5 | ctx\ctx.log:7 | ctx\ctx.log:20 | ctx\ctx.log:30'
    Assert-Equal '見つからない場所: 警告と skipped.csv' ('{0}/{1}/{2}' -f $run.Log.Contains("見つからないため、探しません: $workRoot\nosuch"), (Get-SkipKeys $x $workRoot), (Get-NoteText $run.Log)) 'True/見つからない nosuch/スキップ 1 件（見つからない）'
    $run = Invoke-Plain @('-Path', "$workRoot\nosuch1,$workRoot\nosuch2", '-Pattern', 'ERR', '-NoSave')
    Assert-Equal 'すべて見つからない: 止まる' ('{0}/{1}' -f $run.Code, $run.Log.Contains("指定したフォルダかファイルが見つかりません: $workRoot\nosuch1, $workRoot\nosuch2")) '1/True'
    $run = Invoke-Plain @('-Path', "$tree,$tree\a.log,$tree\", '-Pattern', 'ERR a.log', '-NoSave')
    Assert-Equal '重なった場所の同じファイルは 1 回だけ' (Get-ResultText $run.Log) '結果: 一致 1 行（1 ファイル）   調べたファイル 8 件'
    $run = Invoke-From $ctxDir @('-Pattern', 'hit', '-NoSave')
    Assert-Equal '-Path を省略すると現在のフォルダ' ('{0}/{1}' -f $run.Code, (Get-ResultText $run.Log)) '0/結果: 一致 4 行（1 ファイル）   調べたファイル 1 件'
    $run = Invoke-Plain @('-Path', $ctxDir, '-Pattern', 'no-such-word', '-NoSave')
    Assert-Equal '一致なし: 終了コード 2' ('{0}/{1}' -f $run.Code, (Get-ResultText $run.Log)) '2/結果: 一致なし   調べたファイル 1 件'

    Write-Section '探す文字列（-Pattern・-Regex・-CaseSensitive・-IgnorePattern）'
    $patDir = Join-Path $workRoot 'pat'
    $null = New-TextFile "$patDir\p.log" @(
        'connect 192.168.1.10 ok', 'connect 192x168x1x10 ok', 'path C:\data\file (x) end', 'Timeout occurred',
        'TIMEOUT occurred healthcheck', 'status=5| pipe', 'aa repeated', 'bb repeated', 'refused here')
    $patCases = @(
        @('文字のまま（. はただの文字）', @('-Pattern', '192.168.1.10'), '1'),
        @('文字のまま（\ と ( ）', @('-Pattern', 'C:\data\file (x)'), '3'),
        @('| で区切ると、どれか', @('-Pattern', 'timeout|refused'), '4,5,9'),
        @('-CaseSensitive', @('-Pattern', 'TIMEOUT', '-CaseSensitive'), '5'),
        @('-IgnorePattern', @('-Pattern', 'timeout|refused', '-IgnorePattern', 'HEALTH|here'), '4'),
        @('-Regex', @('-Pattern', '192\.168\.\d+\.\d+|^(aa|bb) ', '-Regex'), '1,7,8'),
        @('-Regex と -IgnorePattern', @('-Pattern', 'o\w+', '-Regex', '-IgnorePattern', 'ok$|here'), '4,5')
    )
    foreach ($case in $patCases) {
        $run = Invoke-Tool 'pat' (@('-Path', $patDir) + $case[1])
        Assert-Equal ('探す文字列: {0}' -f $case[0]) ((@((Read-Result $run.Dir).Matches | ForEach-Object { $_.'行番号' })) -join ',') $case[2]
    }
    $run = Invoke-Tool 'pat-terms' @('-Path', $patDir, '-Pattern', 'timeout|refused|nothing')
    $x = Read-Result $run.Dir
    Assert-Equal '検索の行' (Get-HeaderLine $run.Log '検索: ') '検索: timeout | refused | nothing（大文字と小文字を区別しない）   前 2 行・後ろ 2 行'
    Assert-Equal '語ごとの件数' (@($x.Lines | Where-Object { $_ -match '^  (timeout|refused|nothing) ' }) -join ' | ') '  timeout  2 行（1 ファイル） |   refused  1 行（1 ファイル） |   nothing  0 行（0 ファイル）'
    Assert-Equal '一致した語の列' ((@($x.Matches | ForEach-Object { $_.'一致した語' })) -join ',') 'timeout,timeout,refused'
    $run = Invoke-Plain @('-Path', $patDir, '-Pattern', 'timeout', '-IgnorePattern', 'healthcheck|x', '-CaseSensitive', '-Regex', '-NoSave')
    Assert-Equal '除くの行・正規表現と大文字小文字の表示' ('{0}|{1}' -f (Get-HeaderLine $run.Log '除く: '), (Get-HeaderLine $run.Log '検索: ')) '除く: healthcheck|x|検索: timeout（正規表現・大文字と小文字を区別する）   前 2 行・後ろ 2 行'
    $errorCases = @(
        @('正規表現の誤り', @('-Pattern', '(abc', '-Regex'), 'Pattern の正規表現の書き方が正しくありません: (abc'),
        @('除く語の正規表現の誤り', @('-Pattern', 'abc', '-IgnorePattern', '[x', '-Regex'), 'IgnorePattern の正規表現の書き方が正しくありません: [x'),
        @('-Pattern なし', @(), '探す文字列を -Pattern で指定してください'),
        @('| だけ', @('-Pattern', '|'), '探す文字列を -Pattern で指定してください')
    )
    foreach ($case in $errorCases) {
        $run = Invoke-Plain (@('-Path', $patDir, '-NoSave') + $case[1])
        Assert-Equal ('止まる: {0}' -f $case[0]) ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains($case[2]), $run.Log.Contains('ログ横断検索 (')) '1/True/False'
    }

    Write-Section '最終更新（-Since）'
    $sinceDir = Join-Path $workRoot 'since'
    $null = New-TextFile "$sinceDir\new.log" @('ERR new')
    $null = New-TextFile "$sinceDir\old.log" @('ERR old')
    [IO.File]::SetLastWriteTime("$sinceDir\old.log", (Get-Date).AddDays(-3))
    $twoDays = (Get-Date).AddDays(-2)
    foreach ($since in '1d', '2h', '30m', '48h', $twoDays.ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture), $twoDays.ToString('yyyy/M/d', [Globalization.CultureInfo]::InvariantCulture)) {
        $run = Invoke-Plain @('-Path', $sinceDir, '-Pattern', 'ERR', '-Since', $since, '-NoSave')
        Assert-Equal ('-Since {0}' -f $since) ('{0}/{1}/{2}' -f $run.Code, (Get-ResultText $run.Log), $run.Log.Contains('最終更新で除いたファイル 1 件')) '0/結果: 一致 1 行（1 ファイル）   調べたファイル 1 件/True'
    }
    $run = Invoke-Plain @('-Path', $sinceDir, '-Pattern', 'ERR', '-Since', '4d', '-NoSave')
    Assert-Equal '-Since 4d（どちらも対象）' ('{0}/{1}' -f (Get-ResultText $run.Log), $run.Log.Contains('最終更新で除いたファイル 0 件')) '結果: 一致 2 行（2 ファイル）   調べたファイル 2 件/True'
    Assert-Equal '-Since の条件の行' ([regex]::IsMatch($run.Log, '条件: 最終更新が \d{4}-\d{2}-\d{2} \d{2}:\d{2} 以降のファイル')) 'True'
    foreach ($bad in 'yesterday', '1w', '2026-13-01', '2026/10/31 25:00') {
        $run = Invoke-Plain @('-Path', $sinceDir, '-Pattern', 'ERR', '-Since', $bad, '-NoSave')
        Assert-Equal ('-Since {0}: 止まる' -f $bad) ('{0}/{1}' -f $run.Code, $run.Log.Contains('-Since は "2026-10-31 09:00" のような日時か')) '1/True'
    }

    Write-Section 'CSV（Excel で数式・数値・日付に変わる値）と長い行'
    $csvDir = Join-Path $workRoot 'csv'
    $long305 = '-' + ('x' * 300) + ' ERR'
    $null = New-TextFile "$csvDir\x.log" @('=SUM(1+1) ERR', '@cmd ERR', $long305, ('ERR ' + ('y' * 33000)), 'normal "quoted" ERR')
    $run = Invoke-Tool 'csv' @('-Path', $csvDir, '-Pattern', 'ERR')
    $x = Read-Result $run.Dir
    $cells = @($x.Matches | ForEach-Object { $_.'一致した行' })
    Assert-Equal '= で始まる行は ="..."' $cells[0] '="=SUM(1+1) ERR"'
    Assert-Equal '@ で始まる行は ="..."' $cells[1] '="@cmd ERR"'
    Assert-Equal '255 文字を超える - で始まる行は、先頭に空白' $cells[2] (' ' + $long305)
    Assert-Equal 'セルに入る長さ（32,000 文字）まで' $cells[3].Length 32000
    Assert-Equal '値の中の "' $cells[4] 'normal "quoted" ERR'
    Assert-Equal 'matches.txt には長い行もそのまま' (@($x.Text | Where-Object { $_.Length -gt 33000 }).Count) 1
    Assert-Equal '画面では 500 文字までにして …' ($run.Log.Contains('ERR ' + ('y' * 496) + '…') -and -not $run.Log.Contains('y' * 497)) 'True'
    $null = New-TextFile "$csvDir\num.log" @('0012', '2026-10-31', '12:30', '1,234', 'plain 0012 text')
    $run = Invoke-Tool 'csv-num' @('-Path', "$csvDir\num.log", '-Pattern', '0012|2026-10-31|12:30|1,234')
    $x = Read-Result $run.Dir
    Assert-Equal '数値や日付に見える行は ="..."' ((@($x.Matches | ForEach-Object { $_.'一致した行' })) -join ' | ') '="0012" | ="2026-10-31" | ="12:30" | ="1,234" | plain 0012 text'
    Assert-Equal '数値や日付に見える語も ="..."' ((@($x.Matches | ForEach-Object { $_.'一致した語' })) -join ' | ') '="0012" | ="2026-10-31" | ="12:30" | ="1,234" | ="0012"'
    $run = Invoke-Tool 'csv-one' @('-Path', "$csvDir\num.log", '-Pattern', '0012')
    Assert-Equal '語が 1 つのときの一致した語' ((@((Read-Result $run.Dir).Matches | ForEach-Object { $_.'一致した語' })) -join ' | ') '="0012" | ="0012"'

    Write-Section '画面に出す一致は 200 行まで'
    # 10 行ごとに一致する 3,000 行（一致 300 行）。200 行目の一致は 2,000 行目
    $manyDir = Join-Path $workRoot 'many'
    $null = New-TextFile "$manyDir\m.log" @(1..3000 | ForEach-Object { if ($_ % 10 -eq 0) { 'hit {0}' -f $_ } else { 'ctx {0}' -f $_ } })
    $run = Invoke-Tool 'many' @('-Path', $manyDir, '-Pattern', 'hit')
    $x = Read-Result $run.Dir
    Assert-Equal '画面の一致は 200 行・残りは案内' ('{0}/{1}' -f (Get-ShownCount $run.Log 'hit'), $run.Log.Contains('（ほか 100 行は matches.txt を参照してください）')) '200/True'
    Assert-Equal '200 行目の一致の後ろの行までは出し、次の一致の前の行と -- は出さない' ('{0}/{1}/{2}' -f $run.Log.Contains('   2002  ctx 2002'), $run.Log.Contains('2008  ctx 2008'), $run.Log.Contains('ctx 2002  --')) 'True/False/False'
    Assert-Equal 'matches.txt にはすべて' @($x.Text | Where-Object { $_.StartsWith('>') }).Count 300
    $run = Invoke-Plain @('-Path', $manyDir, '-Pattern', 'hit', '-NoSave')
    Assert-Equal '-NoSave ならすべて画面に出す' ('{0}/{1}' -f (Get-ShownCount $run.Log 'hit'), $run.Log.Contains('matches.txt を参照')) '300/False'

    Write-Section '書き込み中のファイル・開けないファイル'
    $lockDir = Join-Path $workRoot 'lock'
    [void][IO.Directory]::CreateDirectory($lockDir)
    # ログを書いているプログラムの多くは、ほかのプログラムに読むことだけを許して開く
    $writer = New-Object IO.FileStream ("$lockDir\live.log", [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    $holds.Add($writer)
    $bytes = $utf8NoBom.GetBytes("ERR live`r`n")
    $writer.Write($bytes, 0, $bytes.Length); $writer.Flush()
    $null = New-TextFile "$lockDir\excl.log" @('ERR excl')
    $exclusive = New-Object IO.FileStream ("$lockDir\excl.log", [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $holds.Add($exclusive)
    $run = Invoke-Tool 'lock' @('-Path', $lockDir, '-Pattern', 'ERR')
    $x = Read-Result $run.Dir
    Assert-Equal '書き込み中のログも読める' (Get-MatchKeys $x $lockDir) 'live.log:1'
    Assert-Equal 'ほかのプログラムが開いていて読めないファイルは飛ばす' ('{0}/{1}' -f (Get-SkipKeys $x $lockDir), (Get-NoteText $run.Log)) 'その他のエラー excl.log/スキップ 1 件（その他のエラー）'
    foreach ($h in $holds) { $h.Dispose() }
    $holds.Clear()

    Write-Section '保存'
    $noSaveOut = Join-Path $workRoot 'out\nosave'
    $run = Invoke-Plain @('-Path', $ctxDir, '-Pattern', 'hit', '-NoSave', '-OutputFolder', $noSaveOut)
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}/{2}' -f $run.Code, (Test-Path -LiteralPath $noSaveOut), $run.Log.Contains('保存先')) '0/False/False'
    $jpOut = Join-Path $workRoot 'out\保存 先'
    $run = Invoke-Plain @('-Path', $ctxDir, '-Pattern', 'hit', '-OutputFolder', $jpOut)
    Assert-Equal '日本語とスペースを含む保存先' ('{0}/{1}' -f $run.Code, (Get-ResultFileNames (Get-RunDir $jpOut))) '0/matches.csv,matches.txt,skipped.csv,summary.txt'
    # 通常の書き方では末尾のピリオドは無視される。\\?\ 付きで作ると、ピリオドの付いたフォルダができてしまう
    $run = Invoke-Plain @('-Path', $ctxDir, '-Pattern', 'hit', '-OutputFolder', (Join-Path $workRoot 'out\trail.'))
    Assert-Equal '末尾にピリオドが付いた保存先' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultFileNames (Get-RunDir (Join-Path $workRoot 'out\trail'))), [IO.Directory]::Exists('\\?\' + (Join-Path $workRoot 'out\trail.'))) '0/matches.csv,matches.txt,skipped.csv,summary.txt/False'
    $twiceOut = Join-Path $workRoot 'out\twice'
    $null = Invoke-Plain @('-Path', $ctxDir, '-Pattern', 'hit', '-OutputFolder', $twiceOut)
    $null = Invoke-Plain @('-Path', $ctxDir, '-Pattern', 'hit', '-OutputFolder', $twiceOut)
    Assert-Equal '2 回実行すると 2 つのフォルダ' @(Get-ChildItem -LiteralPath $twiceOut -Directory -Filter 'SearchLog_*').Count 2
    $freeLetter = Get-FreeLetter
    if ($freeLetter) {
        $run = Invoke-Plain @('-Path', $manyDir, '-Pattern', 'hit', '-OutputFolder', "${freeLetter}:\nope")
        Assert-Equal '書き込めない保存先: 警告して、一致はすべて画面に' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("保存先に書き込めないため、結果は保存しません: ${freeLetter}:\nope"), (Get-ShownCount $run.Log 'hit')) '0/True/300'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }
    # 保存先のフォルダは作れるが、その中にファイルを作れないようにして、保存に失敗させる
    $guardOut = Join-Path $workRoot 'out\save-fail'
    [void][IO.Directory]::CreateDirectory($guardOut)
    if (Add-Deny $guardOut '(OI)(CI)(WD)') {
        $run = Invoke-Plain @('-Path', $manyDir, '-Pattern', 'hit', '-OutputFolder', $guardOut)
        Assert-Equal '保存に失敗: 警告して、一致はすべて画面に' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした'), $run.Log.Contains('保存先:'), (Get-ShownCount $run.Log 'hit')) '0/True/False/300'
        Assert-Equal '保存に失敗: 書きかけの保存先を残さない' @(Get-ChildItem -LiteralPath $guardOut -Force).Count 0
    } else {
        Write-Skip '保存に失敗したとき' 'アクセス拒否を設定できなかったため'
    }

    Write-Section 'コマンドプロンプトから起動'
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $cmdOut  = Join-Path $workRoot 'out\cmd'
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -Pattern "timeout|refused" -Context "1,0" -OutputFolder "{2}\"' -f $cmdBase, $patDir, $cmdOut)
    Assert-Equal '| を含む -Pattern・-Context "1,0"・最後の引数が "パス\"' ('{0}/{1}/{2}' -f $run.Code, (Get-HeaderLine $run.Log '検索: '), (Get-ResultFileNames (Get-RunDir $cmdOut))) '0/検索: timeout | refused（大文字と小文字を区別しない）   前 1 行・後ろ 0 行/matches.csv,matches.txt,skipped.csv,summary.txt'
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -Pattern no-such-word -NoSave' -f $cmdBase, $ctxDir)
    Assert-Equal '一致なしの終了コード（%ERRORLEVEL%）' $run.Code 2
    # 後ろの -Pattern まで -Path の値になる。「探す文字列がない」ではなく、引用符のヒントを出す
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -Pattern hit' -f $cmdBase, $ctxDir)
    Assert-Equal '"パス\" の後ろに引数: ヒントを出して止まる' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('-Path の値に引用符'), $run.Log.Contains('最後に \ を付けないでください'), $run.Log.Contains('-Pattern で指定してください')) '1/True/True/False'
    $run = Invoke-ViaCmd ('{0} -Path ''{1}'' -Pattern hit -Since 1d -NoSave' -f $cmdBase, $ctxDir)
    Assert-Equal 'cmd で '' で囲んだ -Path と -Since' ('{0}/{1}' -f $run.Code, (Get-ResultText $run.Log)) '0/結果: 一致 4 行（1 ファイル）   調べたファイル 1 件'

    Write-Section '名前に [ ] を含むフォルダから起動（-Path・-ConfigFile・-OutputFolder を相対パスで指定）'
    # 5.1 は、名前に [ ] を含むフォルダで起動すると、名前をワイルドカードとして扱い、別の場所から始まる
    $start = Join-Path $workRoot 'start'
    $startCases = @(
        @('一致するフォルダがない（$PSHOME から始まる）', 'br[1] x', @()),
        @('隣に 1 つ一致するフォルダがある', 'sib\[a]', @('sib\a')),
        @('[ が閉じていない（ドライブのルートから始まる）', 'ub\a[b', @('ub\ab')),
        @('隣に 2 つ一致する（ドライブのルートから始まる）', 'mm\[ab]', @('mm\a', 'mm\b'))
    )
    foreach ($case in $startCases) {
        $null = New-TextFile "$start\$($case[1])\here.log" @('ERR here')
        $null = New-TextFile "$start\$($case[1])\search.ini" @('Path = here.log', 'Pattern = ERR', 'NoSave = true')
        # 間違ったフォルダを探したら分かるように、隣のフォルダには別の行を書く
        foreach ($decoy in $case[2]) { $null = New-TextFile "$start\$decoy\here.log" @('ERR decoy') }
    }
    foreach ($case in $startCases) {
        $dir = "$start\$($case[1])"
        $run = Invoke-From $dir @('-Path', 'here.log', '-Pattern', 'ERR', '-OutputFolder', 'out')
        Assert-Equal ('[ ] を含むフォルダから -Path と -OutputFolder: {0}' -f $case[0]) ('{0}/{1}/{2}' -f $run.Code, ($run.Log.Contains('ERR here') -and -not $run.Log.Contains('ERR decoy')), (Get-ResultFileNames (Get-RunDir "$dir\out"))) '0/True/matches.csv,matches.txt,skipped.csv,summary.txt'
        $run = Invoke-From $dir @('-Pattern', 'ERR', '-NoSave')
        Assert-Equal ('[ ] を含むフォルダから -Path を省略: {0}' -f $case[0]) ('{0}/{1}' -f $run.Code, ($run.Log.Contains('ERR here') -and -not $run.Log.Contains('ERR decoy'))) '0/True'
        $run = Invoke-From $dir @('-ConfigFile', 'search.ini')
        Assert-Equal ('[ ] を含むフォルダから -ConfigFile: {0}' -f $case[0]) ('{0}/{1}' -f $run.Code, ($run.Log.Contains('ERR here') -and -not $run.Log.Contains('ERR decoy'))) '0/True'
    }

    Write-Section '260 文字を超えるパス'
    $longDir  = "$workRoot\" + ('L' * 150) + '\' + ('M' * 150)
    $longFile = New-TextFile "$longDir\deep\long.log" @('first', 'ERR long')
    $run = Invoke-Plain @('-Path', $longDir, '-Pattern', 'ERR', '-NoSave')
    Assert-Equal '-Path（フォルダ）' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultText $run.Log), $run.Log.Contains("■ $longFile（UTF-8）")) '0/結果: 一致 1 行（1 ファイル）   調べたファイル 1 件/True'
    $run = Invoke-Plain @('-Path', $longFile, '-Pattern', 'ERR', '-NoSave')
    Assert-Equal '-Path（ファイル）' ('{0}/{1}' -f $run.Code, (Get-ResultText $run.Log)) '0/結果: 一致 1 行（1 ファイル）   調べたファイル 1 件'
    $run = Invoke-Plain @('-Path', $longDir, '-Pattern', 'ERR', '-OutputFolder', "$longDir\out")
    $longRun = @(Get-ChildItem -LiteralPath ('\\?\' + "$longDir\out") -Directory -ErrorAction SilentlyContinue)
    $longFiles = if ($longRun.Count -gt 0) { Join-Sorted (Get-ChildItem -LiteralPath $longRun[0].FullName -File | ForEach-Object { $_.Name }) } else { '(なし)' }
    Assert-Equal '-OutputFolder' ('{0}/{1}/{2}' -f $run.Code, $longFiles, $run.Log.Contains('保存先に書き込めない')) '0/matches.csv,matches.txt,skipped.csv,summary.txt/False'
    $longIni = New-TextFile "$longDir\search.ini" @('Path = deep', 'Pattern = ERR', 'OutputFolder = 結果')
    $run = Invoke-Plain @('-ConfigFile', $longIni)
    $longRun = @(Get-ChildItem -LiteralPath ('\\?\' + "$longDir\結果") -Directory -ErrorAction SilentlyContinue)
    Assert-Equal '-ConfigFile（Path と OutputFolder は設定ファイルのフォルダが基準）' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultText $run.Log), $longRun.Count) '0/結果: 一致 1 行（1 ファイル）   調べたファイル 1 件/1'

    Write-Section '設定ファイル（-ConfigFile）'
    $cfgDir = Join-Path $workRoot 'cfg\設定 フォルダ'
    $cwd    = Join-Path $workRoot 'cwd'
    [void][IO.Directory]::CreateDirectory($cwd)
    $basic = New-TextFile "$cfgDir\basic.ini" @(
        '# 設定ファイルのテスト（日本語のコメント）'
        '; セミコロンのコメント'
        'Path = ..\..\pat'
        'path = "..\..\ctx\ctx.log"'
        'Pattern = timeout'
        'Pattern = status=5|'
        'pattern = "hit"'
        'IgnorePattern = healthcheck'
        'Filter = *.log'
        'Exclude = *\nothing'
        'Since = 30d'
        'Context = 1,1'
        'OutputFolder = .\結果'
        'NoSave = false'
        'CaseSensitive = false'
        'Regex = false'
        'Pattern ='
    )
    # 別のフォルダから起動して、相対パスが設定ファイルのフォルダを基準にすることを確かめる
    $run = Invoke-From $cwd @('-ConfigFile', $basic)
    Assert-Equal '終了コード' $run.Code 0
    $runDir = Get-RunDir "$cfgDir\結果"
    Assert-Equal '保存先は設定ファイルのフォルダが基準' (Get-ResultFileNames $runDir) 'matches.csv,matches.txt,skipped.csv,summary.txt'
    Assert-Equal '起動したフォルダには保存しない' (Test-Path -LiteralPath "$cwd\結果") 'False'
    $x = Read-Result $runDir
    Assert-Equal '「対象」の次の行に設定ファイル' $x.Lines[2] "設定ファイル: $basic"
    Assert-Equal 'Pattern は 1 行が 1 語（| も文字のまま）' $x.Lines[3] '検索: timeout | status=5| | hit（大文字と小文字を区別しない）   前 1 行・後ろ 1 行'
    Assert-Equal 'IgnorePattern・Since・Exclude の行' ('{0}|{1}|{2}' -f $x.Lines[4], ($x.Lines[5] -match '^条件: 最終更新が '), $x.Lines[6]) '除く: healthcheck|True|除外: *\nothing'
    Assert-Equal 'Path（2 行）の一致' (Get-MatchKeys $x $workRoot) 'pat\p.log:4 | pat\p.log:6 | ctx\ctx.log:5 | ctx\ctx.log:7 | ctx\ctx.log:20 | ctx\ctx.log:30'
    # コマンドラインの値が優先する。-Pattern と -Path は置き換え（両方を合わせない）、-NoSave は NoSave = false より優先
    $run = Invoke-From $cwd @('-ConfigFile', $basic, '-Pattern', 'refused', '-Path', $patDir, '-Context', '0', '-NoSave')
    Assert-Equal 'コマンドラインが優先' ('{0}/{1}/{2}/{3}' -f $run.Code, (Get-HeaderLine $run.Log '検索: '), (Get-ResultText $run.Log), @(Get-ChildItem -LiteralPath "$cfgDir\結果" -Directory).Count) '0/検索: refused（大文字と小文字を区別しない）   前 0 行・後ろ 0 行/結果: 一致 1 行（1 ファイル）   調べたファイル 1 件/1'
    # 正規表現の語が 2 つ以上あるときは、1 つずつ試す（まとめると、2 つ目の語の \1 が 1 つ目の語の ( ) を指してしまう）
    $regexIni = New-TextFile "$cfgDir\regex.ini" @("Path = $patDir", 'Regex = true', 'Pattern = ^(a)\1', 'Pattern = ^(b)\1', 'NoSave = true')
    $run = Invoke-From $cwd @('-ConfigFile', $regexIni)
    Assert-Equal '正規表現の語ごとの後方参照' (Get-ResultText $run.Log) '結果: 一致 2 行（1 ファイル）   調べたファイル 1 件'
    $caseIni = New-TextFile "$cfgDir\case.ini" @("Path = $patDir", 'CaseSensitive = true', 'Pattern = TIMEOUT', 'IgnorePattern = nothing', 'NoSave = true')
    $run = Invoke-From $cwd @('-ConfigFile', $caseIni)
    Assert-Equal 'CaseSensitive = true' (Get-ResultText $run.Log) '結果: 一致 1 行（1 ファイル）   調べたファイル 1 件'

    $encodings = [ordered]@{
        'UTF-8（BOM あり）'             = $utf8Bom
        'UTF-8（BOM なし）'             = $utf8NoBom
        'Shift_JIS'                     = [Text.Encoding]::GetEncoding(932)
        'UTF-16 LE（メモ帳の Unicode）' = [Text.Encoding]::Unicode
        'UTF-16 BE'                     = [Text.Encoding]::BigEndianUnicode
    }
    $i = 0
    foreach ($label in $encodings.Keys) {
        $file = New-TextFile "$cfgDir\enc$i.ini" @('Path = ..\..\enc\sjis.log', 'Pattern = タイムアウト発生', 'NoSave = true') $encodings[$label]
        $run  = Invoke-From $cwd @('-ConfigFile', $file)
        Assert-Equal ('文字コード {0}: 日本語のパスと語' -f $label) ('{0}/{1}' -f $run.Code, (Get-ResultText $run.Log)) '0/結果: 一致 1 行（1 ファイル）   調べたファイル 1 件'
        $i++
    }
    # %TEMP% などの環境変数を展開する
    $env:WORKBENCH_TEST_ROOT = $workRoot
    $envIni = New-TextFile "$cfgDir\env.ini" @('Path = %WORKBENCH_TEST_ROOT%\ctx', 'Pattern = hit', 'OutputFolder = %WORKBENCH_TEST_ROOT%\out\cfg-env')
    $run = Invoke-From $cwd @('-ConfigFile', $envIni)
    Assert-Equal '環境変数' ('{0}/{1}' -f $run.Code, (Get-ResultFileNames (Get-RunDir "$workRoot\out\cfg-env"))) '0/matches.csv,matches.txt,skipped.csv,summary.txt'

    Write-Section '設定ファイル: 書き間違い（探し始める前に止まる）'
    $mistakes = @(
        @('知らない項目', @('Pattern = x', 'Paterns = y'), @('設定ファイルの 2 行目: 知らない項目です: Paterns', '使える項目: Path, Pattern, Regex, CaseSensitive, IgnorePattern, Filter, Exclude, Since, Context, OutputFolder, NoSave')),
        @('ConfigFile は書けない', @('ConfigFile = other.ini'), @('設定ファイルの 1 行目: 知らない項目です: ConfigFile')),
        @('同じ項目を 2 回（複数書けない項目）', @('Context = 2', 'Pattern = x', 'context = 3'), @('設定ファイルの 3 行目: Context が 2 回書かれています（1 行目と 3 行目）')),
        @('Regex が true / false 以外', @('Regex = yes'), @('設定ファイルの 1 行目: Regex は true か false で指定してください: yes')),
        @('NoSave が true / false 以外', @('NoSave = 1'), @('設定ファイルの 1 行目: NoSave は true か false で指定してください: 1')),
        @('= がない行', @('Pattern timeout'), @('「項目 = 値」の形になっていません: Pattern timeout')),
        @('セクションの見出し', @('[Search-Log]', 'Pattern = x'), @('設定ファイルの 1 行目: 「項目 = 値」の形になっていません: [Search-Log]')),
        @('Path の引用符が片方だけ', @('Path = "C:\logs', 'Pattern = x'), @('Path の値に引用符（"）が残っています: "C:\logs')),
        @('定義されていない環境変数', @('Path = %WORKBENCH_NO_SUCH_VAR%\x', 'Pattern = x'), @('定義されていない環境変数です: %WORKBENCH_NO_SUCH_VAR%')),
        @('正規表現の誤り', @('Regex = true', 'Pattern = (x'), @('Pattern の正規表現の書き方が正しくありません: (x'))
    )
    for ($i = 0; $i -lt $mistakes.Count; $i++) {
        $file    = New-TextFile "$cfgDir\err$i.ini" $mistakes[$i][1]
        $run     = Invoke-From $cwd @('-ConfigFile', $file)
        $missing = @(@($mistakes[$i][2]) | Where-Object { -not $run.Log.Contains($_) })
        if ($i -lt $mistakes.Count - 1) { if (-not $run.Log.Contains("ファイル: $file")) { $missing += "ファイル: $file" } }
        Assert-Equal ('書き間違い: {0}' -f $mistakes[$i][0]) ('{0}/{1}/{2}' -f $run.Code, ($missing -join ' | '), $run.Log.Contains('ログ横断検索 (')) '1//False'
    }
    $noteCases = @(
        @('Path が見つからない', @('Pattern = x', 'Path = .\nosuch'), "指定したフォルダかファイルが見つかりません: $cfgDir\nosuch（設定ファイルの 2 行目）"),
        @('Context の誤り', @('Pattern = x', "Path = $ctxDir", 'Context = 99'), '-Context は 0～50 の数値か、"2,10" のように前と後ろの行数で指定してください: 99（設定ファイルの 3 行目）'),
        @('Since の誤り', @('Pattern = x', "Path = $ctxDir", 'Since = 1y'), '-Since は "2026-10-31 09:00" のような日時か、2h・30m・3d のような今からの時間で指定してください: 1y（設定ファイルの 3 行目）')
    )
    for ($i = 0; $i -lt $noteCases.Count; $i++) {
        $file = New-TextFile "$cfgDir\note$i.ini" $noteCases[$i][1]
        $run  = Invoke-From $cwd @('-ConfigFile', $file)
        Assert-Equal ('行番号を添える: {0}' -f $noteCases[$i][0]) ('{0}/{1}' -f $run.Code, $run.Log.Contains($noteCases[$i][2])) '1/True'
    }
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\nothing.ini")
    Assert-Equal '設定ファイルが見つからない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("設定ファイルが見つかりません: $cfgDir\nothing.ini")) '1/True'
    $null = New-TextFile "$cfgDir\memo.ini.txt" @('Pattern = x')
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\memo.ini")
    Assert-Equal '設定ファイルの拡張子が .txt' ('{0}/{1}' -f $run.Code, $run.Log.Contains("拡張子が .txt になっていませんか: $cfgDir\memo.ini.txt")) '1/True'
    if ($freeLetter) {
        $badOut = New-TextFile "$cfgDir\badout.ini" @("Path = $ctxDir", 'Pattern = hit', "OutputFolder = ${freeLetter}:\nope")
        $run = Invoke-From $cwd @('-ConfigFile', $badOut)
        Assert-Equal '設定ファイルの OutputFolder に書き込めない: 行番号' ('{0}/{1}' -f $run.Code, $run.Log.Contains('（設定ファイルの 3 行目）')) '0/True'
    }
    $cmdIni = New-TextFile "$workRoot\cmdcfg\cmd.ini" @("Path = $patDir", 'Pattern = TIMEOUT', 'CaseSensitive = true', 'NoSave = true')
    $run = Invoke-ViaCmd ('{0} -ConfigFile "{1}"' -f $cmdBase, $cmdIni)
    Assert-Equal 'コマンドプロンプトから -ConfigFile（スイッチの項目）' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultText $run.Log), $run.Log.Contains('保存先')) '0/結果: 一致 1 行（1 ファイル）   調べたファイル 1 件/False'

    Write-Section 'Ctrl+C で中断（大きなファイルを読んでいる途中で停止）'
    # 1 MB ずつ書いて、数秒かかる大きさ（約 150 MB）のログを作る
    $bigDir = Join-Path $workRoot 'big'
    [void][IO.Directory]::CreateDirectory($bigDir)
    $block = $utf8NoBom.GetBytes(((1..16000 | ForEach-Object { '2026-10-31 09:00:00 INFO  request ok id={0}' -f $_ }) -join "`r`n") + "`r`n")
    $bigStream = [IO.File]::Create("$bigDir\big.log")
    try { for ($i = 0; $i -lt 150; $i++) { $bigStream.Write($block, 0, $block.Length) } } finally { $bigStream.Dispose() }
    $stopScript = New-HelperScript 'stop-during-search.ps1' @(
        'param($Tool, $LogDir, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $LogDir).AddParameter(''Pattern'', ''no-such-word'').AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'Start-Sleep -Milliseconds 1500'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
    )
    $result = Invoke-Helper $stopScript @('-Tool', $tool, '-LogDir', $bigDir, '-OutRoot', (Join-Path $workRoot 'out\ctrlc'))
    if ($result['RUNNING'] -eq 'True') {
        Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $result['BEFORE'], $result['AFTER']) '1/0'
        Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' '1.5 秒より前に探し終わったため'
    }

    Write-Section 'レビューで見つかった問題の再発防止'
    # 1. 一致を 1 ファイル分ためてから書いていた（上の「画面に出す一致は 200 行まで」と、Ctrl+C の項目で確かめている）
    # 2. 保存の途中で書けなくなったら、そこから先の一致をすべて画面に出す（上の「保存に失敗」で確かめている）
    # 3. 引用符の確認を、探す文字列の確認より先に行う（上の「コマンドプロンプトから起動」で確かめている）
    # 4. ファイルのシンボリックリンクもたどらない（上の「フォルダの階層」で確かめている。作れないときは飛ばす）
    # 5. 正規表現の語が複数あるときは 1 つずつ試す（上の「設定ファイル」の後方参照で確かめている）
    # 6. CSV で、数値や日付に見える値も ="..." で出す（上の「CSV」で確かめている）
    # 7. 画面の 200 行の打ち切りで、末尾に -- と次の一致の前の行を残さない（上の「画面に出す一致は 200 行まで」で確かめている）
    # 8. -Path の一部が見つからなくても、ほかの場所は探す（上の「-Path の指定」で確かめている）
    # 9. -Context 2,5 を PowerShell から書ける（上の「前後の行」で確かめている）
    # 10. 5.1 で 260 文字を超える -OutputFolder に保存できなかった（上の「260 文字を超えるパス」で確かめている）
    # 保存の途中でディスクがいっぱいになる・読んでいる途中で読めなくなる、は再現できないため、実機でも確かめない

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -Pattern hit -Since 1d -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $manyDir.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        $summaryText = $x.Lines -join "`n"
        Assert-Equal 'fi-FI でも対象の行の日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($summaryText, '対象: .*   \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   ')) 'True'
        Assert-Equal 'fi-FI でも条件の行の日時は yyyy-MM-dd HH:mm' ([regex]::IsMatch($summaryText, '条件: 最終更新が \d{4}-\d{2}-\d{2} \d{2}:\d{2} 以降')) 'True'
        Assert-Equal 'fi-FI でも所要時間の小数点は .'([regex]::IsMatch($summaryText, '結果: 一致 300 行（1 ファイル）   調べたファイル 1 件   所要 \d+\.\d秒')) 'True'
        Assert-Equal 'fi-FI でも matches.csv の日時は yyyy/MM/dd HH:mm:ss' ((@($x.Matches)[0]).'最終更新日時' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }
} finally {
    # 開いたままのファイルを閉じ、アクセス拒否を戻してから消す
    foreach ($h in $holds) { try { $h.Dispose() } catch { } }
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    Remove-Item Env:\WORKBENCH_TEST_ROOT -ErrorAction SilentlyContinue
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            # ジャンクションの先（tree 自身）を消さないよう、先にジャンクションだけを外す
            if ([IO.Directory]::Exists("$workRoot\tree\loop")) { [IO.Directory]::Delete("$workRoot\tree\loop") }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

'  END'
exit $script:failCount
