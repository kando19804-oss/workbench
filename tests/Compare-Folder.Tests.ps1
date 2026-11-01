<#
.SYNOPSIS
    Compare-Folder の回帰テストです。

.DESCRIPTION
    一時フォルダに、違いを仕込んだ 2 つのダミーのフォルダ（日本語のパス、260 文字を超えるパス、名前に [ ] を含むフォルダ、
    ジャンクションのループ、読み取り禁止のフォルダ、内容だけが違うファイル、更新日時だけが違うファイルなど）を作り、
    Compare-Folder を実行して期待値と照合します。スナップショットの保存と比較、設定ファイル、止まる場合も確かめます。
    ダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダにだけアクセス拒否を設定します（終了時に戻します）。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Compare-Folder を実行する PowerShell（powershell.exe または pwsh）。

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

$tool     = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Compare-Folder\Compare-Folder.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Compare-Folder_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$ref      = Join-Path $workRoot 'ref'
$dif      = Join-Path $workRoot 'dif'
$baseTime = New-Object DateTime 2026, 10, 1, 10, 0, 0
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

# 260 文字を超えるパスも作れるように、\\?\ を付けて作る。中身を指定しないときは、指定したサイズの 0 で埋める
function New-TestFile([string]$FilePath, [long]$Size = 10, [string]$Text = $null, $Time = $baseTime) {
    $parent = $FilePath.Substring(0, $FilePath.LastIndexOf('\'))
    [void][IO.Directory]::CreateDirectory('\\?\' + $parent)
    if ($null -ne $Text) {
        [IO.File]::WriteAllText('\\?\' + $FilePath, $Text)
    } else {
        $stream = [IO.File]::Create('\\?\' + $FilePath)
        $stream.SetLength($Size)
        $stream.Close()
    }
    [IO.File]::SetLastWriteTime('\\?\' + $FilePath, $Time)
}

function New-TestFolder([string]$FolderPath) { [void][IO.Directory]::CreateDirectory('\\?\' + $FolderPath) }

# コンソールで実行すると、長い行は画面の幅で折り返され、出力に改行が入る。
# 画面の幅によって結果が変わらないように、文字を探す前に改行を取り除く
function ConvertTo-FlatText([string]$Text) { $Text -replace '\r?\n', '' }

# Log には、画面への出力から改行を取り除いたものを入れる。Dir は作られた結果のフォルダ
function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out  = Join-Path $workRoot "out_$Label"
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String)
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = $log; Dir = (Get-RunDir $out) }
}

# 保存先を指定せずに（設定ファイルの OutputFolder を確かめるため）、指定したフォルダから起動する
function Invoke-From([string]$WorkDir, [string[]]$ToolArgs) {
    Push-Location -LiteralPath $WorkDir
    try {
        $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1 | Out-String)
        $code = $LASTEXITCODE
    } finally {
        Pop-Location
    }
    [pscustomobject]@{ Code = $code; Log = $log }
}

# 保存先のフォルダの下に作られた、結果のフォルダ（FolderDiff_* か FolderSnapshot_*）
function Get-RunDir([string]$Parent) {
    if (-not (Test-Path -LiteralPath $Parent)) { return $null }
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory | Where-Object { $_.Name -like 'FolderDiff_*' -or $_.Name -like 'FolderSnapshot_*' })
    if ($found.Count -gt 0) { return $found[0].FullName }
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
    $summary = ''
    if ($Dir -and (Test-Path -LiteralPath (Join-Path $Dir 'summary.txt'))) { $summary = [IO.File]::ReadAllText((Join-Path $Dir 'summary.txt')) }
    [pscustomobject]@{
        Diffs    = @(Read-CsvOrEmpty 'differences.csv')
        Snapshot = @(Read-CsvOrEmpty 'snapshot.csv')
        Skipped  = @(Read-CsvOrEmpty 'skipped.csv')
        Summary  = $summary
    }
}

# 画面の「結果:」の行（件数まで）
function Get-ResultLine([string]$Log) { [regex]::Match($Log, '結果: .*?同じ \S+ 件').Value }

# differences.csv の、相対パスの行
function Get-Row($Rows, [string]$Rel) { @($Rows | Where-Object { $_.'相対パス' -eq $Rel }) }

function Get-Field($Rows, [string]$Rel, [string]$Column) { (Get-Row $Rows $Rel | ForEach-Object { $_.$Column }) -join '|' }

# ツールと同じ並び順（\ をどの文字よりも前に並べ、大文字と小文字を区別しない）になっているか
function Test-SortedRel($Rows) {
    $rels = [string[]]@($Rows | ForEach-Object { $_.'相対パス' -replace '^="(.*)"$', '$1' })
    if ($rels.Count -lt 2) { return $true }
    $keys = [string[]]@($rels | ForEach-Object { $_.Replace('\', [string][char]1) })
    $sorted = [string[]]$rels.Clone()
    [Array]::Sort($keys, $sorted, [StringComparer]::OrdinalIgnoreCase)
    ($rels -join '|') -eq ($sorted -join '|')
}

function Get-FileNames([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    $names = [string[]]@(Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { $_.Name })
    [Array]::Sort($names, [StringComparer]::Ordinal)
    $names -join ','
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

function New-ConfigFile([string]$FilePath, [string[]]$Lines) {
    $parent = $FilePath.Substring(0, $FilePath.LastIndexOf('\'))
    [void][IO.Directory]::CreateDirectory('\\?\' + $parent)
    [IO.File]::WriteAllText('\\?\' + $FilePath, ($Lines -join "`r`n"), (New-Object Text.UTF8Encoding $false))
    $FilePath
}

# スナップショットを、ツールと同じ形（すべての値を引用符で囲む、BOM 付き UTF-8）で書き直す
function Save-Snapshot($Rows, [string]$FilePath) {
    $lines = @($Rows | ConvertTo-Csv -NoTypeInformation)
    [IO.File]::WriteAllLines($FilePath, [string[]]$lines, (New-Object Text.UTF8Encoding $true))
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$userSid   = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$bref      = Join-Path $workRoot 'bref'
$bdif      = Join-Path $workRoot 'bdif'
$deniedDir = Join-Path $bdif 'sec'
$script:denyApplied = $false
$junctions = New-Object 'System.Collections.Generic.List[string]'
$longDir   = ('L' * 120) + '\' + ('M' * 120)

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    foreach ($side in $ref, $dif) {
        New-TestFile "$side\same.txt" 100
        New-TestFile "$side\2026-10-31" 10
        New-TestFile "$side\sub\deep\.keep" 0
        New-TestFile "$side\[a]\f.bin" 20
        New-TestFile "$side\$longDir\deep.bin" 1234
        New-TestFolder "$side\logs"
        $null = New-Item -ItemType Junction -Path "$side\jloop" -Target $side
        $junctions.Add("$side\jloop")
    }
    # .keep は、サブフォルダを作るためだけのもの（両方にあって同じ）
    New-TestFile "$ref\change.txt" 100;                 New-TestFile "$dif\change.txt" 200
    New-TestFile "$ref\dt.txt" 50;                      New-TestFile "$dif\dt.txt" 50 -Time $baseTime.AddHours(1)
    New-TestFile "$ref\fat.txt" 60;                     New-TestFile "$dif\fat.txt" 60 -Time $baseTime.AddSeconds(1)
    New-TestFile "$ref\hashonly.txt" -Text ('a' * 70);  New-TestFile "$dif\hashonly.txt" -Text ('b' * 70)
    New-TestFile "$ref\logs\a.log" 10;                  New-TestFile "$dif\logs\a.log" 11
    New-TestFile "$ref\日本語 フォルダ\メモ.txt" 30
    New-TestFile "$dif\sub\deep\new.txt" 40
    New-TestFile "$ref\kind\inner.txt" 5;               New-TestFile "$dif\kind" 5
    $null = New-Item -ItemType Junction -Path "$ref\jonly" -Target "$ref\sub"
    $junctions.Add("$ref\jonly")

    # 読めないフォルダ用。中身は基準と比較で違うが、比較の側は読めないので比べない
    foreach ($side in $bref, $bdif) {
        foreach ($name in 'a.txt', 'B.txt', 'c\d.txt', 'c.txt', 'c-d.txt', 'Z\y.txt') { New-TestFile "$side\$name" 3 }
    }
    New-TestFile "$bref\sec\x.txt" 3
    New-TestFile "$bdif\sec\y.txt" 3
    $null = & icacls.exe $deniedDir /deny "*${userSid}:(RD)" 2>&1
    $script:denyApplied = ($LASTEXITCODE -eq 0)
    Assert-Equal 'アクセス拒否の設定' $script:denyApplied 'True'

    Write-Section '既定（サイズと更新日時）'
    $run = Invoke-Tool 'default' @('-Reference', $ref, '-Difference', $dif)
    Assert-Equal '終了コード' $run.Code 2
    Assert-Equal '結果の行' (Get-ResultLine $run.Log) '結果: 違いあり   追加 1 件  削除 4 件  変更 3 件  種類が違う 1 件  同じ 13 件'
    Assert-Equal 'リンクの件数' ([regex]::Match($run.Log, 'たどらなかったリンク \S+ 件').Value) 'たどらなかったリンク 2 件'
    $x = Read-Result $run.Dir
    Assert-Equal '保存したファイル' (Get-FileNames $run.Dir) 'differences.csv,skipped.csv,summary.txt'
    Assert-Equal 'differences.csv の行数（違いだけ）' $x.Diffs.Count 9
    Assert-Equal '追加' (Get-Field $x.Diffs 'sub\deep\new.txt' '結果') '追加'
    Assert-Equal '削除（日本語のフォルダと中身）' ('{0}/{1}' -f (Get-Field $x.Diffs '日本語 フォルダ' '種類'), (Get-Field $x.Diffs '日本語 フォルダ\メモ.txt' '結果')) 'フォルダ/削除'
    Assert-Equal '変更（サイズ）' ('{0}/{1}/{2}' -f (Get-Field $x.Diffs 'change.txt' '違い'), (Get-Field $x.Diffs 'change.txt' '基準のサイズ(バイト)'), (Get-Field $x.Diffs 'change.txt' '比較のサイズ(バイト)')) 'サイズ/100/200'
    Assert-Equal '変更（更新日時）' (Get-Field $x.Diffs 'dt.txt' '違い') '更新日時'
    Assert-Equal '2 秒以内の日時の違いは同じ' (Get-Row $x.Diffs 'fat.txt').Count 0
    Assert-Equal '内容だけの違いは、既定では見つけない' (Get-Row $x.Diffs 'hashonly.txt').Count 0
    Assert-Equal '種類が違う' ('{0}/{1}' -f (Get-Field $x.Diffs 'kind' '結果'), (Get-Field $x.Diffs 'kind' '種類')) '種類が違う/フォルダ→ファイル'
    Assert-Equal '種類が違うフォルダの中身は削除' (Get-Field $x.Diffs 'kind\inner.txt' '結果') '削除'
    Assert-Equal '片方にだけあるリンク' ('{0}/{1}/{2}' -f (Get-Field $x.Diffs 'jonly' '結果'), (Get-Field $x.Diffs 'jonly' '種類'), (Get-Field $x.Diffs 'jonly' '備考')) '削除/リンク/ジャンクション'
    Assert-Equal 'ループするジャンクションをたどらない' @($x.Diffs | Where-Object { $_.'相対パス' -like 'jloop\*' }).Count 0
    Assert-Equal '日時の書式' (Get-Field $x.Diffs 'dt.txt' '比較の最終更新日時') '2026/10/01 11:00:00'
    Assert-Equal '相対パスの順' (Test-SortedRel $x.Diffs) 'True'
    Assert-Equal 'summary.txt に違いと保存先' ('{0}/{1}' -f $x.Summary.Contains('sub\deep\new.txt'), $x.Summary.Contains("保存先: $($run.Dir)")) 'True/True'
    $allText = ($x.Diffs + $x.Skipped | ForEach-Object { ($_.PSObject.Properties | ForEach-Object { $_.Value }) -join '|' }) -join '|'
    Assert-Equal '\\?\ が残っていない' $allText.Contains('\\?\') 'False'

    Write-Section '-Hash -IncludeSame'
    $run = Invoke-Tool 'hash' @('-Reference', $ref, '-Difference', $dif, '-Hash', '-IncludeSame')
    Assert-Equal '終了コード' $run.Code 2
    Assert-Equal '比べ方' $run.Log.Contains('比べ方: サイズと更新日時とハッシュ（SHA-256）') 'True'
    Assert-Equal '結果の行' (Get-ResultLine $run.Log) '結果: 違いあり   追加 1 件  削除 4 件  変更 3 件  種類が違う 1 件  日時のみ 1 件  同じ 12 件'
    $x = Read-Result $run.Dir
    Assert-Equal 'differences.csv の行数（同じも含む）' $x.Diffs.Count 22
    Assert-Equal '内容だけの違い' ('{0}/{1}' -f (Get-Field $x.Diffs 'hashonly.txt' '結果'), (Get-Field $x.Diffs 'hashonly.txt' '違い')) '変更/内容'
    Assert-Equal '日時のみ' ('{0}/{1}' -f (Get-Field $x.Diffs 'dt.txt' '結果'), (Get-Field $x.Diffs 'dt.txt' '違い')) '日時のみ/更新日時'
    Assert-Equal 'サイズが違うファイルはハッシュを計算しない' (Get-Field $x.Diffs 'change.txt' '基準のハッシュ') ''
    $expected = (Get-FileHash -LiteralPath "$ref\hashonly.txt" -Algorithm SHA256).Hash
    Assert-Equal 'ハッシュは Get-FileHash と同じ' (Get-Field $x.Diffs 'hashonly.txt' '基準のハッシュ') $expected
    # 1 つのハッシュの計算を使い回すため、2 つ目以降のファイルの値も確かめる
    $mismatch = 0
    foreach ($row in @($x.Diffs | Where-Object { $_.'基準のハッシュ' })) {
        $rel = $row.'相対パス' -replace '^="(.*)"$', '$1'
        if ((Get-FileHash -LiteralPath ('\\?\' + "$ref\$rel") -Algorithm SHA256).Hash -ne $row.'基準のハッシュ') { $mismatch++ }
    }
    Assert-Equal 'すべてのハッシュが Get-FileHash と同じ' $mismatch 0
    Assert-Equal '260 文字を超えるパスも比べる' (Get-Field $x.Diffs "$longDir\deep.bin" '結果') '同じ'
    Assert-Equal '名前に [ ] を含むフォルダ' (Get-Field $x.Diffs '[a]\f.bin' '結果') '同じ'
    Assert-Equal '日付に見える名前は ="..."' (Get-Row $x.Diffs '="2026-10-31"').Count 1
    Assert-Equal '相対パスの順（同じも含む）' (Test-SortedRel $x.Diffs) 'True'

    Write-Section '-Exclude（名前、先頭の \ で直下だけ）'
    $run = Invoke-Tool 'exclude' @('-Reference', $ref, '-Difference', $dif, '-Exclude', '*.log,\kind')
    Assert-Equal '終了コード' $run.Code 2
    Assert-Equal '結果の行' (Get-ResultLine $run.Log) '結果: 違いあり   追加 1 件  削除 3 件  変更 2 件  同じ 13 件'
    Assert-Equal '除外の件数（基準と比較の合計）' ([regex]::Match($run.Log, '除外 \S+ 件').Value) '除外 4 件'
    $x = Read-Result $run.Dir
    $names = [string[]]@(@($x.Skipped | Where-Object { $_.'理由' -eq '除外指定' }) | ForEach-Object { '{0}:{1}' -f $_.'側', $_.'パス'.Substring($workRoot.Length + 1) })
    [Array]::Sort($names, [StringComparer]::Ordinal)
    Assert-Equal 'skipped.csv の除外' ($names -join ',') '基準:ref\kind,基準:ref\logs\a.log,比較:dif\kind,比較:dif\logs\a.log'

    Write-Section '読めないフォルダ（片方でも読めなければ、中を比べない）'
    if ($script:denyApplied) {
        $run = Invoke-Tool 'blocked' @('-Reference', $bref, '-Difference', $bdif, '-IncludeSame')
        Assert-Equal '終了コード（違いはないが、比べられなかった）' $run.Code 1
        Assert-Equal '結果の行' $run.Log.Contains('結果: 違いなし（ただし、比べられなかった場所が 1 件あります。skipped.csv を確認してください）') 'True'
        $x = Read-Result $run.Dir
        Assert-Equal '中身を追加・削除に出さない' @($x.Diffs | Where-Object { $_.'相対パス' -like 'sec\*' }).Count 0
        Assert-Equal 'skipped.csv' (@($x.Skipped) | ForEach-Object { '{0}/{1}/{2}' -f $_.'理由', $_.'側', $_.'パス' }) "権限なし/比較/$deniedDir"
        # レビューで見つかった問題: 読めないフォルダがあると、相対パスの配列が object[] になり、並べ替わらないおそれがあった
        Assert-Equal '相対パスの順（読めないフォルダがあるとき）' (Test-SortedRel $x.Diffs) 'True'
    } else {
        Write-Skip '読めないフォルダ' 'アクセス拒否を設定できなかったため'
    }

    Write-Section 'スナップショット'
    $work = Join-Path $workRoot 'work'
    New-TestFile "$work\a.txt" -Text 'aaa'
    New-TestFile "$work\b.txt" -Text 'bbb'
    New-TestFile "$work\sub\c.txt" -Text 'ccc'
    New-TestFile "$work\old.tmp" -Text 'tmp'
    $snapOut = Join-Path $workRoot 'out_snap$x'
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Reference $work -Snapshot -Hash -Exclude '*.tmp' -OutputFolder $snapOut 2>&1 | Out-String)
    Assert-Equal '保存: 終了コード' $LASTEXITCODE 0
    $snapDir = Get-RunDir $snapOut
    $snapCsv = Join-Path $snapDir 'snapshot.csv'
    Assert-Equal '保存: フォルダ名' ((Split-Path $snapDir -Leaf) -like 'FolderSnapshot_*') 'True'
    Assert-Equal '保存: ファイル' (Get-FileNames $snapDir) 'skipped.csv,snapshot.csv,summary.txt'
    # レビューで見つかった問題: $ を含むパスを "..." で囲むと、PowerShell に貼ったときに変数として読まれる
    Assert-Equal '保存: 比べるコマンド（$ を含むパスは '' で囲む）' $log.Contains(("-Reference '{0}'" -f $snapCsv)) 'True'
    $s = Read-Result $snapDir
    $info = @{}
    foreach ($row in @($s.Snapshot | Where-Object { $_.'種類' -eq '情報' })) { $info[$row.'相対パス'] = $row.'状態' }
    Assert-Equal '保存: 情報の行' ('{0}/{1}/{2}/{3}' -f $info['ツール'], $info['フォルダ'], $info['ハッシュ'], $info['除外']) "Compare-Folder 1.0.0/$work/SHA-256/*.tmp"
    Assert-Equal '保存: 時間帯' $info['時間帯'] ([TimeZoneInfo]::Local.Id)
    Assert-Equal '保存: 除外したファイルは載せない' (@($s.Snapshot | Where-Object { $_.'相対パス' -eq 'old.tmp' }).Count) 0
    Assert-Equal '保存: ハッシュ' (@($s.Snapshot | Where-Object { $_.'相対パス' -eq 'sub\c.txt' }) | ForEach-Object { $_.'ハッシュ' }) (Get-FileHash -LiteralPath "$work\sub\c.txt").Hash
    Assert-Equal '保存: 相対パスの順' (Test-SortedRel @($s.Snapshot | Where-Object { $_.'種類' -ne '情報' })) 'True'
    Assert-Equal '保存: summary.txt に比べるコマンド' $s.Summary.Contains('作業の後で、次のコマンドで比べられます') 'True'

    $run = Invoke-Tool 'snap_same' @('-Reference', $snapCsv, '-Difference', $work)
    Assert-Equal '変わっていない: 終了コード' $run.Code 0
    Assert-Equal '変わっていない: 結果の行' (Get-ResultLine $run.Log) '結果: 違いなし   追加 0 件  削除 0 件  変更 0 件  同じ 4 件'
    Assert-Equal '変わっていない: ハッシュで比べる' $run.Log.Contains('比べ方: サイズと更新日時とハッシュ（SHA-256）') 'True'
    Assert-Equal '変わっていない: 除外を引き継ぐ' $run.Log.Contains('除外: *.tmp') 'True'

    # レビューで見つかった問題: 時間帯の違う PC で保存したスナップショットを読むと、更新日時がずれていた
    $otherZoneId = if ([TimeZoneInfo]::Local.Id -eq 'UTC') { 'Tokyo Standard Time' } else { 'UTC' }
    $otherZone   = $null
    try { $otherZone = [TimeZoneInfo]::FindSystemTimeZoneById($otherZoneId) } catch { }
    if ($otherZone) {
        $rows = @(Import-Csv -LiteralPath $snapCsv -Encoding UTF8)
        foreach ($row in $rows) {
            if ($row.'種類' -eq '情報' -and $row.'相対パス' -eq '時間帯') { $row.'状態' = $otherZoneId }
            if ($row.'種類' -eq 'ファイル') {
                $local = [datetime]::ParseExact($row.'最終更新日時', 'yyyy/MM/dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
                $other = [TimeZoneInfo]::ConvertTime($local, [TimeZoneInfo]::Local, $otherZone)
                $row.'最終更新日時' = $other.ToString('yyyy/MM/dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
            }
        }
        $zoneCsv = Join-Path $workRoot 'zone.csv'
        Save-Snapshot $rows $zoneCsv
        $run = Invoke-Tool 'snap_zone' @('-Reference', $zoneCsv, '-Difference', $work)
        Assert-Equal '別の時間帯で保存: ずれない' ('{0}/{1}' -f $run.Code, (Get-ResultLine $run.Log)) '0/結果: 違いなし   追加 0 件  削除 0 件  変更 0 件  同じ 4 件'
        # 時間帯の行だけを変えると、日時がずれたとみなす（ハッシュは同じなので「日時のみ」）
        $rows = @(Import-Csv -LiteralPath $snapCsv -Encoding UTF8)
        foreach ($row in $rows) { if ($row.'種類' -eq '情報' -and $row.'相対パス' -eq '時間帯') { $row.'状態' = $otherZoneId } }
        Save-Snapshot $rows $zoneCsv
        $run = Invoke-Tool 'snap_zone2' @('-Reference', $zoneCsv, '-Difference', $work)
        Assert-Equal '時間帯だけ変える: 日時のみ' ('{0}/{1}' -f $run.Code, $run.Log.Contains('日時のみ 3 件')) '0/True'
    } else {
        Write-Skip '別の時間帯で保存' ('時間帯 {0} が見つからないため' -f $otherZoneId)
    }
    $rows = @(Import-Csv -LiteralPath $snapCsv -Encoding UTF8)
    foreach ($row in $rows) { if ($row.'種類' -eq '情報' -and $row.'相対パス' -eq '時間帯') { $row.'状態' = 'Nowhere Standard Time' } }
    $zoneCsv = Join-Path $workRoot 'zone_unknown.csv'
    Save-Snapshot $rows $zoneCsv
    $run = Invoke-Tool 'snap_zone3' @('-Reference', $zoneCsv, '-Difference', $work)
    Assert-Equal '時間帯が見つからない: 警告して続ける' ('{0}/{1}' -f $run.Code, $run.Log.Contains('時間帯（Nowhere Standard Time）が見つからないため')) '0/True'

    # 作業の後: 追加・変更・除外したファイルの追加
    New-TestFile "$work\new.txt" -Text 'new'
    New-TestFile "$work\b.txt" -Text 'bbbb'
    New-TestFile "$work\junk.tmp" -Text 'junk'
    $run = Invoke-Tool 'snap_after' @('-Reference', $snapCsv, '-Difference', $work)
    Assert-Equal '作業の後: 終了コード' $run.Code 2
    Assert-Equal '作業の後: 結果の行' (Get-ResultLine $run.Log) '結果: 違いあり   追加 1 件  削除 0 件  変更 1 件  同じ 3 件'
    Assert-Equal '作業の後: 基準の行' $run.Log.Contains("基準: スナップショット $snapCsv") 'True'
    $x = Read-Result $run.Dir
    Assert-Equal '作業の後: 除外したファイルは出ない' (Get-Row $x.Diffs 'junk.tmp').Count 0
    # レビューで見つかった問題: 比べるときに足した除外で基準の側から外れたものが、記録に残らなかった
    $run = Invoke-Tool 'snap_excl' @('-Reference', $snapCsv, '-Difference', $work, '-Exclude', 'a.txt')
    $x = Read-Result $run.Dir
    Assert-Equal 'スナップショットの側の除外も記録する' ((@($x.Skipped | Where-Object { $_.'理由' -eq '除外指定' }) | ForEach-Object { $_.'側' }) -join ',') '基準,比較'

    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Reference $work -Snapshot -OutputFolder (Join-Path $workRoot 'out_snap_nohash') 2>&1 | Out-String)
    $noHashCsv = Join-Path (Get-RunDir (Join-Path $workRoot 'out_snap_nohash')) 'snapshot.csv'
    $run = Invoke-Tool 'snap_nohash' @('-Reference', $noHashCsv, '-Difference', $work, '-Hash')
    Assert-Equal 'ハッシュのないスナップショットで -Hash' ('{0}/{1}' -f $run.Code, $run.Log.Contains('このスナップショットにはハッシュがありません')) '1/True'

    $broken = Join-Path $workRoot 'broken.csv'
    $lines = @([IO.File]::ReadAllLines($snapCsv))
    [IO.File]::WriteAllLines($broken, [string[]]($lines + '"ファイル","x.txt"'), (New-Object Text.UTF8Encoding $true))
    $run = Invoke-Tool 'snap_broken' @('-Reference', $broken, '-Difference', $work)
    Assert-Equal '列の足りない行: 何行目か' ('{0}/{1}' -f $run.Code, $run.Log.Contains(('{0} 行目が正しくありません' -f ($lines.Count + 1)))) '1/True'
    $notSnap = Join-Path $workRoot 'not.csv'
    [IO.File]::WriteAllText($notSnap, "a,b`r`n1,2`r`n")
    $run = Invoke-Tool 'snap_not' @('-Reference', $notSnap, '-Difference', $work)
    Assert-Equal 'スナップショットでない CSV' ('{0}/{1}' -f $run.Code, $run.Log.Contains('スナップショットのファイルとして読めません')) '1/True'

    Write-Section 'オンライン専用ファイル（-Hash でもハッシュを計算しない）'
    $onA = Join-Path $workRoot 'onA'; $onB = Join-Path $workRoot 'onB'
    New-TestFile "$onA\f.txt" -Text 'same'; New-TestFile "$onB\f.txt" -Text 'same'
    try {
        [IO.File]::SetAttributes("$onB\f.txt", [IO.FileAttributes]::Offline)
        $run = Invoke-Tool 'online' @('-Reference', $onA, '-Difference', $onB, '-Hash', '-IncludeSame')
        Assert-Equal '終了コード' $run.Code 0
        Assert-Equal '件数の案内' $run.Log.Contains('ハッシュで比べられなかったファイルが 1 件あります') 'True'
        Assert-Equal '備考' (Get-Field (Read-Result $run.Dir).Diffs 'f.txt' '備考') 'オンライン専用のため、ハッシュを計算していない'
    } catch {
        Write-Skip 'オンライン専用ファイル' ('属性を付けられなかったため: {0}' -f $_.Exception.Message)
    }

    Write-Section '止まるとき'
    $cases = @(
        @('比較を指定していない', @('-Reference', $ref), '比較のフォルダを指定してください'),
        @('基準が見つからない', @('-Reference', (Join-Path $workRoot 'no\such'), '-Difference', $dif), '基準のフォルダを開けません'),
        @('比較が見つからない', @('-Reference', $ref, '-Difference', (Join-Path $workRoot 'no\such')), '比較のフォルダを開けません'),
        @('基準と比較が同じ', @('-Reference', $ref, '-Difference', "$ref\."), '基準と比較が同じフォルダです'),
        @('比較が基準の中にある', @('-Reference', $ref, '-Difference', "$ref\sub"), '片方が、もう片方の中にあります'),
        @('基準が比較の中にある', @('-Reference', "$ref\sub", '-Difference', $ref), '片方が、もう片方の中にあります'),
        @('-Snapshot と -Difference', @('-Reference', $ref, '-Snapshot', '-Difference', $dif), '-Snapshot のときは、-Difference は指定できません'),
        @('-Snapshot と -NoSave', @('-Reference', $ref, '-Snapshot', '-NoSave'), '-Snapshot のときは、-NoSave は指定できません'),
        @('-Snapshot でスナップショットのファイル', @('-Reference', $snapCsv, '-Snapshot'), '-Reference にフォルダを指定してください')
    )
    foreach ($case in $cases) {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @($case[1]) -NoSave 2>&1 | Out-String)
        $code = $LASTEXITCODE
        Assert-Equal $case[0] ('{0}/{1}' -f $code, $log.Contains($case[2])) '1/True'
    }
    # 名前が似ているだけのフォルダ（ref と ref2）は、入れ子とみなさない
    New-TestFile "$workRoot\ref2\same.txt" 100
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Reference $ref -Difference "$workRoot\ref2" -NoSave 2>&1 | Out-String)
    Assert-Equal '名前が似ているだけのフォルダは入れ子でない' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('もう片方の中にあります')) '2/False'

    $noSaveOut = Join-Path $workRoot 'out_nosave'
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Reference $ref -Difference $dif -NoSave -OutputFolder $noSaveOut 2>&1 | Out-String)
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}' -f $LASTEXITCODE, (Test-Path -LiteralPath $noSaveOut)) '2/False'
    $usedLetters = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    $freeLetter  = @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $usedLetters -notcontains $_ } | Select-Object -First 1
    if ($freeLetter) {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Reference $ref -Difference $dif -OutputFolder "${freeLetter}:\nope" 2>&1 | Out-String)
        Assert-Equal '書き込めない保存先: 比べて終わる' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('保存先に書き込めないため')) '2/True'
        Assert-Equal '書き込めない保存先: skipped.csv を案内しない' $log.Contains('skipped.csv') 'False'
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Reference $ref -Snapshot -OutputFolder "${freeLetter}:\nope" 2>&1 | Out-String)
        Assert-Equal '書き込めない保存先: スナップショットは止まる' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('スナップショットを保存できません')) '1/True'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }

    Write-Section 'コマンドプロンプトから起動'
    # 名前を省略した 2 つのパス、カンマ区切りの -Exclude、最後の引数が "パス\"
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -NoSave -Exclude "*.log,\kind" "{2}" "{3}\"' -f $Exe, $tool, $ref, $dif)
    Assert-Equal '終了コード' $run.Code 2
    Assert-Equal '結果の行' (Get-ResultLine $run.Log) '結果: 違いあり   追加 1 件  削除 3 件  変更 2 件  同じ 13 件'
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -Reference "{2}\" -Difference "{3}" -NoSave' -f $Exe, $tool, $ref, $dif)
    Assert-Equal '"パス\" の後ろに引数: 終了コード' $run.Code 1
    Assert-Equal '"パス\" の後ろに引数: ヒント' ([regex]::IsMatch($run.Log, '最後に \\ を付けないでください')) 'True'

    Write-Section '設定ファイル（-ConfigFile）'
    $cfgDir = Join-Path $workRoot 'cfg\設定 フォルダ'
    $cfg = New-ConfigFile "$cfgDir\c.ini" @('# テスト', 'Reference = ..\..\ref', 'Difference = "..\..\dif"', 'Hash = true', 'Exclude = *.log', 'Exclude = \kind', 'OutputFolder = out')
    $run = Invoke-From $workRoot @('-ConfigFile', $cfg)
    Assert-Equal '終了コード' $run.Code 2
    Assert-Equal '結果の行' (Get-ResultLine $run.Log) '結果: 違いあり   追加 1 件  削除 3 件  変更 2 件  日時のみ 1 件  同じ 12 件'
    Assert-Equal '設定ファイルの行' $run.Log.Contains("設定ファイル: $cfg") 'True'
    Assert-Equal '相対パスは設定ファイルのフォルダが基準' ('{0}/{1}' -f $run.Log.Contains("基準: $ref"), ($null -ne (Get-RunDir "$cfgDir\out"))) 'True/True'
    # 同じ設定ファイルで、スナップショットも保存できる（Difference は使わない）
    $run = Invoke-From $workRoot @('-ConfigFile', $cfg, '-Snapshot', '-OutputFolder', (Join-Path $workRoot 'out_cfgsnap'))
    Assert-Equal '-Snapshot のときは Difference を使わない' $run.Code 0
    $bad = New-ConfigFile "$cfgDir\bad.ini" @('Reference = ..\..\ref', 'Snapshot = true')
    $run = Invoke-From $workRoot @('-ConfigFile', $bad)
    Assert-Equal 'Snapshot は書けない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('設定ファイルの 2 行目'), $run.Log.Contains('Snapshot は設定ファイルには書けません')) '1/True/True'
    $ns = New-ConfigFile "$cfgDir\nosave.ini" @('Reference = ..\..\ref', 'NoSave = true')
    $run = Invoke-From $workRoot @('-ConfigFile', $ns, '-Snapshot')
    Assert-Equal '-Snapshot と設定ファイルの NoSave' ('{0}/{1}' -f $run.Code, $run.Log.Contains('設定ファイルの 2 行目に NoSave = true があります')) '1/True'
    $typo = New-ConfigFile "$cfgDir\typo.ini" @('Reference = ..\..\ref', 'Diference = ..\..\dif')
    $run = Invoke-From $workRoot @('-ConfigFile', $typo)
    Assert-Equal '知らない項目' ('{0}/{1}' -f $run.Code, $run.Log.Contains('知らない項目です: Diference')) '1/True'

    Write-Section 'Ctrl+C で中断（ハッシュを計算している途中で停止）'
    $bigA = Join-Path $env:SystemRoot 'System32'
    $bigB = Join-Path $env:SystemRoot 'SysWOW64'
    if (Test-Path -LiteralPath $bigB) {
        # Ctrl+C と同じ「パイプラインの停止」を、対象の PowerShell の中から送る
        $stopScript = Join-Path $workRoot 'stop-during-hash.ps1'
        $stopOut    = Join-Path $workRoot 'out_ctrlc'
        [IO.File]::WriteAllLines($stopScript, [string[]]@(
            'param($Tool, $A, $B, $OutRoot)'
            '$ps = [powershell]::Create()'
            '[void]$ps.AddCommand($Tool).AddParameter(''Reference'', $A).AddParameter(''Difference'', $B).AddParameter(''Hash'', $true).AddParameter(''OutputFolder'', $OutRoot)'
            '$handle = $ps.BeginInvoke()'
            'Start-Sleep -Seconds 5'
            '''RUNNING='' + (-not $handle.IsCompleted)'
            '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
            '$ps.Stop()'
            '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
            '''MESSAGE='' + (($ps.Streams.Information | ForEach-Object { $_.MessageData.ToString() }) -join '' / '')'
        ), (New-Object Text.UTF8Encoding $true))
        $result = @{}
        & $Exe -NoProfile -ExecutionPolicy Bypass -File $stopScript -Tool $tool -A $bigA -B $bigB -OutRoot $stopOut 2>&1 | ForEach-Object {
            $line = "$_"
            $eq = $line.IndexOf('=')
            if ($eq -gt 0) { $result[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
        }
        if ($result['RUNNING'] -eq 'True') {
            Assert-Equal '止める前は保存先を作ってある' $result['BEFORE'] '1'
            Assert-Equal '止めた後は保存先が残らない' $result['AFTER'] '0'
            # レビューで見つかった問題: 想定外のエラーと Ctrl+C の表示を分けた。Ctrl+C は「中断しました」のまま
            Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
        } else {
            Write-Skip 'Ctrl+C で中断' '5 秒以内に比べ終わったため、途中で止められなかった'
        }
    } else {
        Write-Skip 'Ctrl+C で中断' 'SysWOW64 がないため'
    }

    Write-Section '実機: 大文字と小文字を区別するフォルダ'
    $csA = Join-Path $workRoot 'csA'; $csB = Join-Path $workRoot 'csB'
    New-TestFolder $csA
    $null = & fsutil.exe file setCaseSensitiveInfo $csA enable 2>&1
    if ($LASTEXITCODE -eq 0) {
        New-TestFile "$csA\a.txt" 1; New-TestFile "$csA\A.txt" 2
        New-TestFile "$csB\a.txt" 1
        $run = Invoke-Tool 'case' @('-Reference', $csA, '-Difference', $csB)
        $x = Read-Result $run.Dir
        Assert-Equal '大文字と小文字だけ違う名前を記録に残す' (@($x.Skipped | Where-Object { $_.'詳細' -like '大文字と小文字だけが違う名前*' }).Count) 1
        $null = & fsutil.exe file setCaseSensitiveInfo $csA disable 2>&1
    } else {
        Write-Skip '大文字と小文字を区別するフォルダ' 'このPCでは設定できないため'
    }

    Write-Section '実機: Windows の互換用リンク（Default ユーザーの Application Data）'
    $defaultProfile = Join-Path (Split-Path $env:PUBLIC -Parent) 'Default'
    if (Test-Path -LiteralPath (Join-Path $defaultProfile 'Application Data')) {
        $out = Join-Path $workRoot 'out_compat'
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Reference $defaultProfile -Snapshot -OutputFolder $out 2>&1
        $s = Read-Result (Get-RunDir $out)
        $row = @($s.Snapshot | Where-Object { $_.'相対パス' -eq 'Application Data' })
        Assert-Equal 'リンクとして記録し、中をたどらない' ('{0}/{1}' -f ($row | ForEach-Object { $_.'種類' }), @($s.Snapshot | Where-Object { $_.'相対パス' -like 'Application Data\*' }).Count) 'リンク/0'
    } else {
        Write-Skip '互換用リンク' 'このPCにないため'
    }
} finally {
    # アクセス拒否を戻してから消す
    if ($script:denyApplied) { $null = & icacls.exe $deniedDir /remove:d "*$userSid" 2>&1 }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        # ループするジャンクションは、リンク先をたどらないように先にリンクだけ消す
        try {
            foreach ($junction in $junctions) { if (Test-Path -LiteralPath $junction) { [IO.Directory]::Delete($junction) } }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

exit $script:failCount
