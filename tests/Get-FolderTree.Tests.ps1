<#
.SYNOPSIS
    Get-FolderTree の回帰テストです。

.DESCRIPTION
    一時フォルダにダミー環境（並び順と罫線を確かめる小さなツリー、日本語のパス、260 文字を超えるパス、
    ジャンクションのループ、読み取り禁止のフォルダ、隠し・システム属性、Excel で別の値に変わる名前、
    文字コードや書き方の違う設定ファイルなど）を作り、Get-FolderTree を実行して期待値と照合します。
    ダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダにだけアクセス拒否を設定します（終了時に戻します）。
    ドライブのルートを再現するため、空いているドライブ文字に subst でダミー環境のフォルダを割り当てます（終了時に外します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Get-FolderTree を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Get-FolderTree\Get-FolderTree.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Get-FolderTree_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool     = Join-Path $workRoot 'tool\Get-FolderTree.ps1'
$gold     = Join-Path $workRoot 'gold'
$tree     = Join-Path $workRoot 'tree'
$allFiles = 'items.csv,skipped.csv,summary.txt,tree.txt'
$fixedTime = Get-Date '2026-01-02 03:04:05'
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

# 260 文字を超えるパスも作れるように、\\?\ を付けて作る
function New-SizedFile([string]$FilePath, [long]$Size, $Time = $null) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    $stream = [IO.File]::Create('\\?\' + $FilePath)
    $stream.SetLength($Size)
    $stream.Close()
    if ($Time) { [IO.File]::SetLastWriteTime('\\?\' + $FilePath, $Time) }
}

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

# 設定ファイルの OutputFolder や相対パスの基準を確かめるため、-OutputFolder を足さず、指定したフォルダから起動する
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
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $log = ConvertTo-FlatText ((@(& cmd.exe /d /c $batch 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'FolderTree_*' -ErrorAction SilentlyContinue | Sort-Object Name)
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
    function Read-TextOrEmpty([string]$Name) {
        if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir $Name))) { return '' }
        [IO.File]::ReadAllText((Join-Path $Dir $Name))
    }
    [pscustomobject]@{
        Items   = @(Read-CsvOrEmpty 'items.csv')
        Skipped = @(Read-CsvOrEmpty 'skipped.csv')
        Tree    = (Read-TextOrEmpty 'tree.txt') -replace '\r\n', "`n"
        Summary = Read-TextOrEmpty 'summary.txt'
    }
}

# 5.1 と 7 で並び順が変わらないように、文字コード順に並べてカンマでつなぐ
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

# tree.txt の期待値を組み立てる（1 行目は起点のパス）
function Join-Tree([string]$First, [string[]]$Rest) { (@($First) + $Rest) -join "`n" }

# items.csv の、パスが一致する行。ないとき（ツールが保存に失敗したときなど）も止まらないように、空の行を返す
function Get-ItemRow($Result, [string]$FullPath) {
    $found = @($Result.Items | Where-Object { $_.'パス' -eq $FullPath }) | Select-Object -First 1
    if ($found) { return $found }
    [pscustomobject]@{ '種類' = '(なし)'; 'パス' = ''; '名前' = '(なし)'; '階層' = ''; 'サイズ(MB)' = ''; 'サイズ(バイト)' = ''; 'ファイル数' = ''; '最終更新日時' = ''; '状態' = '(なし)' }
}

# skipped.csv を「理由|Root から見たパス|詳細」にして並べる
function Get-SkippedText($Result, [string]$Root) {
    Join-Sorted (@($Result.Skipped) | ForEach-Object { '{0}|{1}|{2}' -f $_.'理由', $_.'パス'.Substring($Root.Length + 1), $_.'詳細' })
}

# 画面の「スキップ…」の行（ない場合は空）
function Get-NoteText([string]$Log) {
    $part = 'スキップ \S+ 件（[^）]+）|たどらなかったリンク \S+ 件|除外したフォルダ \S+ 件'
    [regex]::Match($Log, "($part)(  ($part))*").Value
}

# 画面の「対象」の行に出たフォルダ
function Get-TargetText([string]$Log) { [regex]::Match($Log, '対象: (.+?)   \d{4}-').Groups[1].Value }

function Get-CountText([string]$Log) { [regex]::Match($Log, 'フォルダ \S+ 件(  ファイル \S+ 件)?').Value }

function Get-ConditionText([string]$Log) { [regex]::Match($Log, '条件: .+?(?=除外: |フォルダ \d[\d,]* 件)').Value }

function New-ConfigFile([string]$FilePath, [string[]]$Lines, [Text.Encoding]$Encoding = (New-Object Text.UTF8Encoding $false)) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllText('\\?\' + $FilePath, ($Lines -join "`r`n"), $Encoding)
    $FilePath
}

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $used -notcontains $_ -and -not [IO.Directory]::Exists("${_}:\") } | Select-Object -First 1
}

# 自分自身にアクセス拒否を設定する。終了時に戻すため、設定したフォルダを覚えておく
function Add-Deny([string]$Path, [string]$Rights) {
    $null = & icacls.exe $Path /deny "*${userSid}:$Rights" 2>&1
    if ($LASTEXITCODE -ne 0) { return $false }
    $denied.Add($Path)
    $true
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$userSid    = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$denied     = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したフォルダ（終了時に戻す）
$links      = New-Object 'System.Collections.Generic.List[string]'   # 作ったリンク（消すときに先に外す）
$substDrive = $null

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    # 並び順と罫線を確かめる小さなツリー
    New-SizedFile "$gold\b\x.txt" 10 $fixedTime
    New-SizedFile "$gold\a\sub\deep\f.txt" 20 $fixedTime
    foreach ($name in 'file10', 'file2', 'File1') { [void][IO.Directory]::CreateDirectory("$gold\$name") }
    New-SizedFile "$gold\z.txt" 40 $fixedTime
    New-SizedFile "$gold\A.txt" 30 $fixedTime

    # 機能を確かめるツリー
    New-SizedFile "$tree\a.txt" 100
    New-SizedFile "$tree\hidden.txt" 50
    [IO.File]::SetAttributes("$tree\hidden.txt", 'Hidden')
    New-SizedFile "$tree\sys.txt" 60
    [IO.File]::SetAttributes("$tree\sys.txt", 'System')
    New-SizedFile "$tree\日本語 フォルダ\データ.bak" 2000
    New-SizedFile "$tree\日本語 フォルダ\sub\deep.LOG" 500
    New-SizedFile "$tree\2026-10-31\d.txt" 10
    New-SizedFile "$tree\-old\o.txt" 10
    New-SizedFile "$tree\001\z.txt" 10
    New-SizedFile "$tree\[br]\b.txt" 10
    New-SizedFile "$tree\node_modules\x.js" 700
    New-SizedFile "$tree\pkg\node_modules\y.js" 100
    New-SizedFile "$tree\lvl1\lvl2\lvl3\lvl4\deep.txt" 5
    New-SizedFile "$tree\lvl1\node_modules\n.js" 7
    New-SizedFile ("$tree\" + ('L' * 150) + '\' + ('M' * 150) + '\long.bin') 1234
    New-SizedFile "$tree\.git\HEAD" 3
    [IO.File]::SetAttributes("$tree\.git", 'Directory, Hidden')
    New-SizedFile "$tree\denied\secret.txt" 400
    [void][IO.Directory]::CreateDirectory("$tree\empty")
    $null = New-Item -ItemType Junction -Path "$tree\jloop" -Target $tree
    $links.Add("$tree\jloop")
    $null = New-Item -ItemType Junction -Path "$tree\lvl1\jdeep" -Target $tree
    $links.Add("$tree\lvl1\jdeep")
    $denyOk = Add-Deny "$tree\denied" '(RD)'
    Assert-Equal 'アクセス拒否の設定' $denyOk 'True'
    $longDir = "$tree\" + ('L' * 150) + '\' + ('M' * 150)

    Write-Section '並び順と罫線（既定・-Files・-Detail・-Ascii）'
    $run = Invoke-Tool 'gold' @('-Path', $gold)
    Assert-Equal '既定: 終了コード' $run.Code 0
    Assert-Equal '既定: tree.txt' (Read-Result $run.Dir).Tree.TrimEnd("`n") (Join-Tree $gold @(
        '├─a', '│  └─sub', '│      └─deep', '├─b', '├─File1', '├─file2', '└─file10'))
    $run = Invoke-Tool 'gold-files' @('-Path', $gold, '-Files', '-Depth', '2')
    Assert-Equal '-Files -Depth 2: tree.txt' (Read-Result $run.Dir).Tree.TrimEnd("`n") (Join-Tree $gold @(
        '├─a', '│  └─sub', '│      └─…', '├─b', '│  └─x.txt', '├─File1', '├─file2', '├─file10', '├─A.txt', '└─z.txt'))
    $run = Invoke-Tool 'gold-detail' @('-Path', $gold, '-Files', '-Detail', '-Depth', '2')
    Assert-Equal '-Files -Detail -Depth 2: tree.txt' (Read-Result $run.Dir).Tree.TrimEnd("`n") (Join-Tree "$gold  (100 B・4 ファイル)" @(
        '├─a  (20 B・1 ファイル)', '│  └─sub  (20 B・1 ファイル)', '│      └─…', '├─b  (10 B・1 ファイル)', '│  └─x.txt  (10 B・2026-01-02 03:04)',
        '├─File1  (0 B・0 ファイル)', '├─file2  (0 B・0 ファイル)', '├─file10  (0 B・0 ファイル)', '├─A.txt  (30 B・2026-01-02 03:04)', '└─z.txt  (40 B・2026-01-02 03:04)'))
    $run = Invoke-Tool 'gold-ascii' @('-Path', $gold, '-Ascii', '-Depth', '1')
    Assert-Equal '-Ascii -Depth 1: tree.txt' (Read-Result $run.Dir).Tree.TrimEnd("`n") (Join-Tree $gold @('+-- a', '|   \-- ...', '+-- b', '+-- File1', '+-- file2', '\-- file10'))
    Assert-Equal '-Ascii: 画面にも同じツリー' $run.Log.Contains("$gold+-- a|   \-- ...+-- b+-- File1+-- file2\-- file10") 'True'

    Write-Section '既定（オプションなし）'
    $run = Invoke-Tool 'default' @('-Path', $tree)
    Assert-Equal '終了コード' $run.Code 0
    $x = Read-Result $run.Dir
    Assert-Equal '保存するファイル' (Get-ResultFileNames $run.Dir) $allFiles
    Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'items.csv=True,skipped.csv=True,summary.txt=True,tree.txt=True'
    Assert-Equal '条件の行' (Get-ConditionText $run.Log) '条件: 3 階層まで・フォルダのみ'
    Assert-Equal '件数の行' (Get-CountText $run.Log) 'フォルダ 19 件'
    Assert-Equal 'スキップの行' (Get-NoteText $run.Log) 'スキップ 1 件（権限なし）  たどらなかったリンク 2 件'
    Assert-Equal '管理者の注記' $run.Log.Contains('※ 管理者として実行すると、読める範囲が広がります') (-not $isElevated)
    Assert-Equal '権限なしの印' $x.Tree.Contains("├─denied  [権限なし]") 'True'
    Assert-Equal 'リンクの印' ('{0}/{1}' -f $x.Tree.Contains('├─jloop  [リンク]'), $x.Tree.Contains('jdeep  [リンク]')) 'True/True'
    Assert-Equal '隠しフォルダ（.git）は出さない' $x.Tree.Contains('.git') 'False'
    Assert-Equal '260 文字を超えるフォルダ' $x.Tree.Contains('M' * 150) 'True'
    Assert-Equal '4 階層目は … にする' $x.Tree.Contains("lvl3`n│  │      └─…") 'True'
    Assert-Equal 'skipped.csv' (Get-SkippedText $x $tree) (Join-Sorted @('リンク|jloop|ジャンクション', 'リンク|lvl1\jdeep|ジャンクション', '権限なし|denied|アクセスが拒否されました'))
    Assert-Equal 'items.csv の見出し' (Get-FirstLine $run.Dir 'items.csv') '"種類","パス","名前","階層","サイズ(MB)","サイズ(バイト)","ファイル数","最終更新日時","状態"'
    Assert-Equal 'items.csv の行数（起点＋フォルダ 19）' $x.Items.Count 20
    $rootRow = Get-ItemRow $x $tree
    Assert-Equal 'items.csv の起点の行' ('{0}/{1}/{2}/{3}' -f $rootRow.'種類', $rootRow.'名前', $rootRow.'階層', $rootRow.'状態') 'フォルダ/="tree"/0/'
    $deniedRow = Get-ItemRow $x "$tree\denied"
    Assert-Equal 'items.csv の権限なし' ('{0}/{1}' -f $deniedRow.'状態', $deniedRow.'サイズ(バイト)') '権限なし/'
    Assert-Equal 'items.csv のリンク' (Get-ItemRow $x "$tree\jloop").'状態' 'リンク'
    Assert-Equal '\\?\ が残っていない' ($x.Tree + $x.Summary + (($x.Items + $x.Skipped | ForEach-Object { $_.'パス' }) -join '|')).Contains('\\?\') 'False'
    Assert-Equal 'summary.txt にヘッダーとツリー' ('{0}/{1}' -f $x.Summary.StartsWith('フォルダツリー出力 (Get-FolderTree 1.0.0)'), $x.Summary.Contains($x.Tree.Split("`n")[1])) 'True/True'

    Write-Section '-Files・-Detail・-Hidden'
    $run = Invoke-Tool 'files' @('-Path', $tree, '-Files')
    Assert-Equal '-Files: 件数の行' (Get-CountText $run.Log) 'フォルダ 19 件  ファイル 11 件'
    Assert-Equal '-Files: 条件の行' (Get-ConditionText $run.Log) '条件: 3 階層まで・フォルダとファイル'
    $x = Read-Result $run.Dir
    Assert-Equal '-Files: 隠し・システムのファイルは出さない' ('{0}/{1}' -f $x.Tree.Contains('hidden.txt'), $x.Tree.Contains('sys.txt')) 'False/False'
    $fileRow = Get-ItemRow $x "$tree\日本語 フォルダ\データ.bak"
    Assert-Equal '-Files: ファイルの行' ('{0}/{1}/{2}/{3}' -f $fileRow.'種類', $fileRow.'階層', $fileRow.'サイズ(バイト)', ($fileRow.'最終更新日時' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$')) 'ファイル/2/2,000/True'
    Assert-Equal '-Files: フォルダの行にサイズを出さない' (Get-ItemRow $x "$tree\001").'サイズ(バイト)' ''

    $run = Invoke-Tool 'detail' @('-Path', $tree, '-Detail')
    $x = Read-Result $run.Dir
    Assert-Equal '-Detail: 起点の合計（隠し・リンク・権限なしを除く）' $x.Tree.Split("`n")[0] "$tree  (4.6 KB・12 ファイル)"
    $rootRow = Get-ItemRow $x $tree
    Assert-Equal '-Detail: items.csv の起点の合計' ('{0}/{1}/{2}' -f $rootRow.'サイズ(バイト)', $rootRow.'ファイル数', $rootRow.'サイズ(MB)') '4,686/12/0.00'
    Assert-Equal '-Detail: 260 文字を超えるフォルダの合計' $x.Tree.Contains(('M' * 150) + '  (1.2 KB・1 ファイル)') 'True'
    Assert-Equal '-Detail: 空のフォルダ' $x.Tree.Contains('├─empty  (0 B・0 ファイル)') 'True'
    $run = Invoke-Tool 'detail-hidden' @('-Path', $tree, '-Detail', '-Hidden')
    $x = Read-Result $run.Dir
    Assert-Equal '-Detail -Hidden: 起点の合計に隠し項目を含む' $x.Tree.Split("`n")[0] "$tree  (4.7 KB・15 ファイル)"
    Assert-Equal '-Hidden: 条件の行' (Get-ConditionText $run.Log) '条件: 3 階層まで・フォルダのみ・隠し項目を含む'
    Assert-Equal '-Hidden: .git を出す' $x.Tree.Contains('├─.git  (3 B・1 ファイル)') 'True'

    Write-Section '-Exclude'
    $run = Invoke-Tool 'exclude' @('-Path', $tree, '-Exclude', '*\NODE_MODULES,*\jloop')
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '除外の行' $run.Log.Contains('除外: *\NODE_MODULES, *\jloop') 'True'
    Assert-Equal '件数の行' (Get-CountText $run.Log) 'フォルダ 15 件'
    Assert-Equal 'スキップの行' (Get-NoteText $run.Log) 'スキップ 1 件（権限なし）  たどらなかったリンク 1 件  除外したフォルダ 4 件'
    Assert-Equal '除外したフォルダを出さない' ('{0}/{1}' -f $x.Tree.Contains('node_modules'), $x.Tree.Contains('jloop')) 'False/False'
    Assert-Equal 'skipped.csv の除外指定' (Join-Sorted (@($x.Skipped | Where-Object { $_.'理由' -eq '除外指定' }) | ForEach-Object { $_.'パス'.Substring($tree.Length + 1) + '=' + $_.'詳細' })) (Join-Sorted @('jloop=*\jloop', 'lvl1\node_modules=*\NODE_MODULES', 'node_modules=*\NODE_MODULES', 'pkg\node_modules=*\NODE_MODULES'))
    $run = Invoke-Tool 'exclude-detail' @('-Path', $tree, '-Detail', '-Exclude', '*\node_modules')
    Assert-Equal '-Detail の合計に除外したフォルダを含めない' (Read-Result $run.Dir).Tree.Split("`n")[0] "$tree  (3.8 KB・9 ファイル)"
    $run = Invoke-Tool 'exclude-bracket' @('-Path', $tree, '-Exclude', '*\[br]')
    Assert-Equal '[ ] をそのままの文字として扱う' ('{0}/{1}' -f (Read-Result $run.Dir).Tree.Contains('[br]'), $run.Log.Contains('除外したフォルダ 1 件')) 'False/True'

    Write-Section '異常系と起動'
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path (Join-Path $workRoot 'no\such') -NoSave 2>&1) | ForEach-Object { "$_" }) -join "`n")
    Assert-Equal '存在しないパス' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains(('指定したフォルダが見つかりません: ' + (Join-Path $workRoot 'no\such')))) '1/True'
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $gold -Depth 0 -NoSave 2>&1
    Assert-Equal '-Depth 0 は止まる' $LASTEXITCODE 1
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $gold -Depth 101 -NoSave 2>&1
    Assert-Equal '-Depth 101 は止まる' $LASTEXITCODE 1
    $noSaveOut = Join-Path $workRoot 'out\nosave'
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $gold -NoSave -OutputFolder $noSaveOut 2>&1) | ForEach-Object { "$_" }) -join "`n")
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}/{2}' -f $LASTEXITCODE, (Test-Path -LiteralPath $noSaveOut), $log.Contains('保存先')) '0/False/False'
    $freeLetter = Get-FreeLetter
    if ($freeLetter) {
        $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $gold -OutputFolder "${freeLetter}:\nope" 2>&1) | ForEach-Object { "$_" }) -join "`n")
        Assert-Equal '書き込めない保存先' ('{0}/{1}/{2}' -f $LASTEXITCODE, $log.Contains('保存先に書き込めないため、結果は保存しません'), $log.Contains('└─file10')) '0/True/True'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }
    $run = Invoke-Tool 'empty-root' @('-Path', "$tree\empty", '-Files', '-Detail')
    Assert-Equal '空のフォルダを起点にする' ('{0}/{1}' -f $run.Code, (Read-Result $run.Dir).Tree.TrimEnd("`n")) "0/$tree\empty  (0 B・0 ファイル)"

    Write-Section 'コマンドプロンプトから起動'
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $run = Invoke-ViaCmd ('{0} -NoSave -Exclude "*\node_modules,*\jloop" -Path "{1}\"' -f $cmdBase, $tree)
    Assert-Equal '最後の引数が "パス\" でも読める・-Exclude のカンマ区切り' ('{0}/{1}' -f $run.Code, $run.Log.Contains('除外したフォルダ 4 件')) '0/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -NoSave' -f $cmdBase, $tree)
    Assert-Equal '"パス\" の後ろに引数: ヒントを出して止まる' ('{0}/{1}' -f $run.Code, $run.Log.Contains('最後に \ を付けないでください')) '1/True'
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}" -Exclude ''*\node_modules,*\jloop''' -f $cmdBase, $tree)
    Assert-Equal 'cmd で ''…,…'' と囲んだ -Exclude' $run.Log.Contains('除外したフォルダ 4 件') 'True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -OutputFolder "{2}\" -Files' -f $cmdBase, $gold, (Join-Path $workRoot 'out\cmd-q'))
    Assert-Equal '-OutputFolder "…\" の後ろに引数' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-OutputFolder の値に引用符')) '1/True'

    Write-Section 'レビューで見つかった問題の再発防止'
    # 2. -Exclude がジャンクションに効かず、[リンク] と表示されていた（上の -Exclude で確かめている）。シンボリックリンクも確かめる
    $linkTree = Join-Path $workRoot 'linktest'
    [void][IO.Directory]::CreateDirectory("$linkTree\real")
    $symOk = $false
    try { $null = New-Item -ItemType SymbolicLink -Path "$linkTree\slink" -Target $gold -ErrorAction Stop; $links.Add("$linkTree\slink"); $symOk = $true } catch { }
    if ($symOk) {
        $run = Invoke-Tool 'symlink' @('-Path', $linkTree)
        $x = Read-Result $run.Dir
        Assert-Equal 'シンボリックリンクをたどらない' ('{0}/{1}' -f $x.Tree.Contains('└─slink  [リンク]'), (Get-SkippedText $x $linkTree)) 'True/リンク|slink|シンボリックリンク'
        $run = Invoke-Tool 'symlink-ex' @('-Path', $linkTree, '-Exclude', '*\slink')
        Assert-Equal '除外したシンボリックリンクを出さない' ('{0}/{1}' -f (Read-Result $run.Dir).Tree.Contains('slink'), $run.Log.Contains('たどらなかったリンク')) 'False/False'
    } else {
        Write-Skip 'シンボリックリンク' 'シンボリックリンクを作れないため（管理者か開発者モードが必要）'
    }

    # 3. items.csv の名前が Excel で日付や数式に変わっていた（常に ="..." の形）
    $x = Read-Result (Get-RunDir (Join-Path $workRoot 'out\default'))
    Assert-Equal '名前は ="..." の形' ('{0}|{1}|{2}' -f (Get-ItemRow $x "$tree\2026-10-31").'名前', (Get-ItemRow $x "$tree\-old").'名前', (Get-ItemRow $x "$tree\001").'名前') '="2026-10-31"|="-old"|="001"'

    # 4. -Detail なしのとき、-Depth の 1 つ下の除外とリンクだけが件数に入っていた
    $run = Invoke-Tool 'below-depth' @('-Path', $tree, '-Depth', '1', '-Exclude', '*\node_modules')
    Assert-Equal '-Depth 1: ツリーに出した範囲だけで数える' (Get-NoteText $run.Log) 'スキップ 1 件（権限なし）  たどらなかったリンク 1 件  除外したフォルダ 1 件'
    $run = Invoke-Tool 'below-depth-detail' @('-Path', $tree, '-Depth', '1', '-Exclude', '*\node_modules', '-Detail')
    Assert-Equal '-Depth 1 -Detail: 全階層で数える' (Get-NoteText $run.Log) 'スキップ 1 件（権限なし）  たどらなかったリンク 2 件  除外したフォルダ 3 件'

    # 5. 保存に失敗したとき、消えた tree.txt を案内していた。500 行を超えるツリーで、保存できたとき・-NoSave・保存できなかったときを比べる
    $many = Join-Path $workRoot 'many'
    for ($i = 1; $i -le 600; $i++) { [void][IO.Directory]::CreateDirectory(('{0}\d{1:0000}' -f $many, $i)) }
    $run = Invoke-Tool 'many' @('-Path', $many)
    Assert-Equal '500 行を超えるとき: 画面は 500 行まで' ('{0}/{1}' -f [regex]::Matches($run.Log, '─d\d{4}').Count, $run.Log.Contains('（ほか 101 行は tree.txt を参照してください）')) '499/True'
    Assert-Equal '500 行を超えるとき: tree.txt はすべて' (Read-Result $run.Dir).Tree.TrimEnd("`n").Split("`n").Count 601
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $many -NoSave 2>&1) | ForEach-Object { "$_" }) -join "`n")
    Assert-Equal '-NoSave: 全部を画面に出す' ('{0}/{1}' -f [regex]::Matches($log, '─d\d{4}').Count, $log.Contains('tree.txt')) '600/False'
    # 保存先のフォルダは作れるが、その中にファイルを作れないようにして、保存に失敗させる
    $guardOut = Join-Path $workRoot 'out\save-fail'
    [void][IO.Directory]::CreateDirectory($guardOut)
    if (Add-Deny $guardOut '(OI)(CI)(WD)') {
        $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $many -OutputFolder $guardOut 2>&1) | ForEach-Object { "$_" }) -join "`n")
        Assert-Equal '保存に失敗: 全部を画面に出し、tree.txt を案内しない' ('{0}/{1}/{2}' -f [regex]::Matches($log, '─d\d{4}').Count, $log.Contains('tree.txt を参照'), $log.Contains('結果を保存できませんでした')) '600/False/True'
        Assert-Equal '保存に失敗: 書きかけの保存先を残さない' @(Get-ChildItem -LiteralPath $guardOut -Force).Count 0
    } else {
        Write-Skip '保存に失敗したとき' 'アクセス拒否を設定できなかったため'
    }

    # 6. items.csv の起点の行の名前がフルパスだった（上の「items.csv の起点の行」で確かめている）。260 文字を超える起点でも最後のフォルダ名
    $run = Invoke-Tool 'long-root' @('-Path', $longDir)
    Assert-Equal '260 文字を超える -Path: 起点の名前' ('{0}/{1}' -f $run.Code, (Get-ItemRow (Read-Result $run.Dir) $longDir).'名前') ('0/="' + ('M' * 150) + '"')

    # 7. 空のフォルダで止まっていた（上の「空のフォルダを起点にする」と -Detail の empty で確かめている）

    Write-Section 'パスの書き方'
    $run = Invoke-Tool 'prefix' @('-Path', ('\\?\' + $gold), '-NoSave')
    Assert-Equal '\\?\ 付きの -Path' ('{0}/{1}/{2}' -f $run.Code, (Get-TargetText $run.Log), $run.Log.Contains('\\?\')) "0/$gold/False"
    $run = Invoke-Tool 'period' @('-Path', "$gold.", '-NoSave')
    Assert-Equal '末尾にピリオドが付いた -Path' ('{0}/{1}' -f $run.Code, (Get-TargetText $run.Log)) "0/$gold"
    $run = Invoke-Tool 'jp' @('-Path', "$tree\日本語 フォルダ", '-Files')
    Assert-Equal '日本語とスペースを含む -Path' ('{0}/{1}' -f $run.Code, (Read-Result $run.Dir).Tree.Contains('└─データ.bak')) '0/True'

    Write-Section 'ドライブのルート（subst で再現）'
    $letter = Get-FreeLetter
    if ($letter) {
        $null = & subst.exe "${letter}:" $gold 2>&1
        if ($LASTEXITCODE -eq 0) { $substDrive = "${letter}:" }
    }
    if ($substDrive) {
        $run = Invoke-Tool 'root' @('-Path', $substDrive)
        $x = Read-Result $run.Dir
        Assert-Equal 'ドライブ文字だけでもルートとして扱う' ('{0}/{1}/{2}' -f $run.Code, $x.Tree.Split("`n")[0], (Get-ItemRow $x "$substDrive\").'名前') "0/$substDrive\/=""$substDrive\"""
        $null = & subst.exe $substDrive /D 2>&1
        $substDrive = $null
    } else {
        Write-Skip 'ドライブのルート' 'subst でドライブ文字を割り当てられなかったため'
    }

    Write-Section '名前に [ ] を含むフォルダから起動（-Path を省略）'
    # 5.1 は、名前に [ ] を含むフォルダで起動すると、名前をワイルドカードとして扱い、別の場所から始まる
    $start = Join-Path $workRoot 'start'
    $startCases = @(
        @('一致するフォルダがない（$PSHOME から始まる）', 'br[1] x', @()),
        @('隣に 1 つ一致するフォルダがある', 'sib\[a]', @('sib\a')),
        @('[ が閉じていない（ドライブのルートから始まる）', 'ub\a[b', @('ub\ab')),
        @('隣に 2 つ一致する（ドライブのルートから始まる）', 'mm\[ab]', @('mm\a', 'mm\b'))
    )
    foreach ($case in $startCases) {
        New-SizedFile "$start\$($case[1])\inner\f.bin" 11
        foreach ($decoy in $case[2]) { New-SizedFile "$start\$decoy\decoy\d.bin" 22 }
    }
    foreach ($case in $startCases) {
        $dir = "$start\$($case[1])"
        $run = Invoke-From $dir @('-NoSave')
        Assert-Equal ('[ ] を含むフォルダから起動: {0}' -f $case[0]) ('{0}/{1}' -f (Get-TargetText $run.Log), $run.Log.Contains('└─inner')) "$dir/True"
    }

    Write-Section '設定ファイル（-ConfigFile）'
    $cfgDir = Join-Path $workRoot 'cfg\設定 フォルダ'
    $cwd    = Join-Path $workRoot 'cwd'
    [void][IO.Directory]::CreateDirectory($cwd)
    $basic = New-ConfigFile "$cfgDir\basic.ini" @(
        '# 設定ファイルのテスト（日本語のコメント）'
        '; セミコロンのコメント'
        'path = ..\..\tree'
        'DEPTH=2'
        'Files = true'
        'Hidden = false'
        'Detail = true'
        'Exclude = *\node_modules'
        'Exclude = "*\JLOOP"'
        'Ascii = false'
        'OutputFolder = .\結果'
        'NoSave = false'
    )
    # 別のフォルダから起動して、相対パスが設定ファイルのフォルダを基準にすることを確かめる
    $run = Invoke-From $cwd @('-ConfigFile', $basic)
    Assert-Equal '終了コード' $run.Code 0
    $runDir = Get-RunDir "$cfgDir\結果"
    Assert-Equal '保存先は設定ファイルのフォルダが基準' (Get-ResultFileNames $runDir) $allFiles
    Assert-Equal '起動したフォルダには保存しない' (Test-Path -LiteralPath "$cwd\結果") 'False'
    $x = Read-Result $runDir
    Assert-Equal 'summary の 3 行目に設定ファイル' (($x.Summary -split '\r?\n')[2]) "設定ファイル: $basic"
    Assert-Equal 'Depth と Files と Detail' ('{0}/{1}' -f (Get-ConditionText $run.Log), $x.Tree.Split("`n")[0].StartsWith("$tree  (")) '条件: 2 階層まで・フォルダとファイル/True'
    Assert-Equal 'Exclude を 2 行に分けて書く' ('{0}/{1}' -f $run.Log.Contains('除外: *\node_modules, *\JLOOP'), (Get-NoteText $run.Log)) 'True/スキップ 1 件（権限なし）  たどらなかったリンク 1 件  除外したフォルダ 4 件'

    # コマンドラインの値が優先する。-Exclude は置き換え（両方を合わせない）、-NoSave は NoSave = false より優先
    $run = Invoke-From $cwd @('-ConfigFile', $basic, '-Depth', '1', '-Exclude', '*\pkg', '-NoSave')
    Assert-Equal 'コマンドラインが優先' ('{0}/{1}/{2}/{3}' -f $run.Code, (Get-ConditionText $run.Log), $run.Log.Contains('除外: *\pkg'), @(Get-ChildItem -LiteralPath "$cfgDir\結果" -Directory).Count) '0/条件: 1 階層まで・フォルダとファイル/True/1'

    # Windows 11 のメモ帳は BOM なしの UTF-8、Server 2016 / 2019 のメモ帳は Shift_JIS で保存する
    $encodings = [ordered]@{
        'UTF-8（BOM なし）' = New-Object Text.UTF8Encoding $false
        'UTF-8（BOM 付き）' = New-Object Text.UTF8Encoding $true
        'Shift_JIS'         = [Text.Encoding]::GetEncoding(932)
        'UTF-16 LE'         = [Text.Encoding]::Unicode
        'UTF-16 BE'         = [Text.Encoding]::BigEndianUnicode
    }
    $i = 0
    foreach ($label in $encodings.Keys) {
        $file = New-ConfigFile "$cfgDir\enc$i.ini" @('Path = ..\..\tree\日本語 フォルダ', 'NoSave = true', 'Files = true') $encodings[$label]
        $run  = Invoke-From $cwd @('-ConfigFile', $file)
        Assert-Equal ('文字コード {0}: 日本語のパス' -f $label) ('{0}/{1}/{2}' -f $run.Code, (Get-TargetText $run.Log), $run.Log.Contains('データ.bak')) "0/$tree\日本語 フォルダ/True"
        $i++
    }

    # %TEMP% などの環境変数を展開する
    $env:WORKBENCH_TEST_ROOT = $workRoot
    $envIni = New-ConfigFile "$cfgDir\env.ini" @('Path = %WORKBENCH_TEST_ROOT%\gold', 'OutputFolder = %WORKBENCH_TEST_ROOT%\out\cfg-env')
    $run = Invoke-From $cwd @('-ConfigFile', $envIni)
    Assert-Equal '環境変数' ('{0}/{1}/{2}' -f $run.Code, (Get-TargetText $run.Log), (Get-ResultFileNames (Get-RunDir "$workRoot\out\cfg-env"))) "0/$gold/$allFiles"

    Write-Section '設定ファイル: 書き間違い（調べ始める前に止まる）'
    $mistakes = @(
        @('知らない項目', @('Path = ..\..\gold', 'Exlude = *\a'), @('設定ファイルの 2 行目: 知らない項目です: Exlude', '使える項目: Path, Depth, Files, Hidden, Detail, Exclude, Ascii, OutputFolder, NoSave')),
        @('ConfigFile は書けない', @('ConfigFile = other.ini'), @('設定ファイルの 1 行目: 知らない項目です: ConfigFile')),
        @('同じ項目を 2 回', @('Depth = 2', 'Files = true', 'depth = 3'), @('設定ファイルの 3 行目: Depth が 2 回書かれています（1 行目と 3 行目）')),
        @('Depth が範囲外', @('Depth = 101'), @('設定ファイルの 1 行目: Depth は 1～100 の数値で指定してください: 101')),
        @('スイッチが true / false 以外', @('Files = yes'), @('設定ファイルの 1 行目: Files は true か false で指定してください: yes')),
        @('= がない行', @('Depth 4'), @('「項目 = 値」の形になっていません: Depth 4')),
        @('セクションの見出し', @('[Get-FolderTree]', 'Depth = 2'), @('設定ファイルの 1 行目: 「項目 = 値」の形になっていません: [Get-FolderTree]')),
        @('Path の引用符が片方だけ', @('Path = "D:\data'), @('Path の値に引用符（"）が残っています: "D:\data')),
        @('Exclude の引用符が片方だけ', @('Exclude = "*\a'), @('Exclude の値に引用符（"）が残っています: "*\a')),
        @('定義されていない環境変数', @('OutputFolder = %WORKBENCH_NO_SUCH_VAR%\x'), @('定義されていない環境変数です: %WORKBENCH_NO_SUCH_VAR%'))
    )
    for ($i = 0; $i -lt $mistakes.Count; $i++) {
        $file    = New-ConfigFile "$cfgDir\err$i.ini" $mistakes[$i][1]
        $run     = Invoke-From $cwd @('-ConfigFile', $file)
        $missing = @(@($mistakes[$i][2]) + "ファイル: $file" | Where-Object { -not $run.Log.Contains($_) })
        # 終了コード / 表示に足りない文 / 調べ始めたか
        Assert-Equal ('書き間違い: {0}' -f $mistakes[$i][0]) ('{0}/{1}/{2}' -f $run.Code, ($missing -join ' | '), $run.Log.Contains('フォルダツリー出力 (')) '1//False'
    }
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\nothing.ini")
    Assert-Equal '設定ファイルが見つからない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("設定ファイルが見つかりません: $cfgDir\nothing.ini")) '1/True'
    $noPath = New-ConfigFile "$cfgDir\nopath.ini" @('NoSave = true', 'Path = .\nosuch')
    $run = Invoke-From $cwd @('-ConfigFile', $noPath)
    Assert-Equal '設定ファイルの Path が見つからない: 行番号' ('{0}/{1}' -f $run.Code, $run.Log.Contains("指定したフォルダが見つかりません: $cfgDir\nosuch（設定ファイルの 2 行目）")) '1/True'
    $cmdIni = New-ConfigFile "$workRoot\cfg\cmd.ini" @('Path = ..\gold', 'NoSave = true', 'Files = true')
    $run = Invoke-ViaCmd ('{0} -ConfigFile "{1}"' -f $cmdBase, $cmdIni)
    Assert-Equal 'コマンドプロンプトから -ConfigFile（スイッチの項目）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('└─z.txt')) '0/True'

    Write-Section '共有フォルダ経由（\\localhost\<ドライブ>$）'
    $uncGold = '\\localhost\{0}${1}' -f $gold.Substring(0, 1), $gold.Substring(2)
    if (Test-Path -LiteralPath $uncGold) {
        $local = (Read-Result (Get-RunDir (Join-Path $workRoot 'out\gold-files'))).Tree.TrimEnd("`n").Split("`n")
        $run   = Invoke-Tool 'unc' @('-Path', $uncGold, '-Files', '-Depth', '2')
        $remote = (Read-Result $run.Dir).Tree.TrimEnd("`n").Split("`n")
        Assert-Equal '共有フォルダ: 1 行目は共有のパス' $remote[0] $uncGold
        Assert-Equal '共有フォルダ: ツリーはローカルと同じ' (($remote | Select-Object -Skip 1) -join '|') (($local | Select-Object -Skip 1) -join '|')
    } else {
        Write-Skip '共有フォルダ経由' '管理共有を読めないため'
    }

    Write-Section '実機: Windows の互換用リンク（All Users）'
    $usersDir = Split-Path $env:PUBLIC -Parent
    $allUsers = Join-Path $usersDir 'All Users'
    if (Test-Path -LiteralPath $allUsers) {
        $run = Invoke-Tool 'allusers' @('-Path', $usersDir, '-Depth', '1')
        Assert-Equal '-Hidden なし: All Users を出さない' ('{0}/{1}' -f $run.Code, (Read-Result $run.Dir).Tree.Contains('All Users')) '0/False'
        $run = Invoke-Tool 'allusers-hidden' @('-Path', $usersDir, '-Depth', '1', '-Hidden')
        $x = Read-Result $run.Dir
        $allUsersRow = @($x.Skipped | Where-Object { $_.'パス' -eq $allUsers }) | Select-Object -First 1
        if (-not $allUsersRow) { $allUsersRow = [pscustomobject]@{ '理由' = '(なし)'; '詳細' = '(なし)' } }
        Assert-Equal '-Hidden: All Users は [リンク]' ('{0}/{1}' -f $x.Tree.Contains('All Users  [リンク]'), $allUsersRow.'理由') 'True/リンク'
        # 5.1 は読み取りが禁止されたリンクの種類を判定できず、隠し・システム属性で互換用リンクと判定する。7 は種類を読める（Get-DiskUsage も同じ）
        $expectedDetail = if ($hostTag -eq '51') { 'Windows の互換用リンク' } else { 'シンボリックリンク' }
        Assert-Equal '-Hidden: All Users の詳細' $allUsersRow.'詳細' $expectedDetail
    } else {
        Write-Skip 'All Users' 'このPCにないため'
    }
    $oneDrive = $env:OneDrive
    if ($oneDrive -and [IO.Directory]::Exists($oneDrive) -and ([IO.File]::GetAttributes($oneDrive) -band [IO.FileAttributes]::ReparsePoint)) {
        # OneDrive のフォルダもリパースポイントなので、リンクとして飛ばしてはいけない（画面に出す名前は確かめない）
        $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path (Split-Path $oneDrive -Parent) -Depth 1 -NoSave 2>&1) | ForEach-Object { "$_" }) -join "`n")
        Assert-Equal 'OneDrive のフォルダはリンクにしない' $log.Contains((Split-Path $oneDrive -Leaf) + '  [リンク]') 'False'
    } else {
        Write-Skip 'OneDrive' 'リパースポイントの OneDrive のフォルダがないため'
    }

    Write-Section '地域設定と警告'
    # 日付の書式が、OS の地域設定によって変わらないこと
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -Files -Detail -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $gold.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        Assert-Equal 'fi-FI でも summary の日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($x.Summary, '対象: .*   \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   ')) 'True'
        Assert-Equal 'fi-FI でも ツリーの日時は yyyy-MM-dd HH:mm' $x.Tree.Contains('z.txt  (40 B・2026-01-02 03:04)') 'True'
        Assert-Equal 'fi-FI でも items.csv の日時は yyyy/MM/dd HH:mm:ss' (Get-ItemRow $x "$gold\z.txt").'最終更新日時' '2026/01/02 03:04:05'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }
    # 保存先を作れないときの警告に、改行が混ざらないこと。画面の文字は折り返されるため、ストリームから受け取る
    $warnScript = Join-Path $workRoot 'collect-warnings.ps1'
    $warnFile   = Join-Path $workRoot 'warnings.txt'
    [IO.File]::WriteAllLines($warnScript, [string[]]@(
        'param($Tool, $Path, $OutputFolder, $ResultFile)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $Path).AddParameter(''OutputFolder'', $OutputFolder)'
        '[void]$ps.Invoke()'
        '$lines = @($ps.Streams.Warning | ForEach-Object { $_.Message.Replace("`r", ''<CR>'').Replace("`n", ''<LF>'') })'
        '[IO.File]::WriteAllLines($ResultFile, [string[]]$lines, (New-Object Text.UTF8Encoding $true))'
    ), (New-Object Text.UTF8Encoding $true))
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $warnScript -Tool $tool -Path $gold -OutputFolder '\\localhost\workbench_no_such_share\out' -ResultFile $warnFile 2>&1
    $saveWarnings = @()
    if (Test-Path -LiteralPath $warnFile) { $saveWarnings = @([IO.File]::ReadAllLines($warnFile) | Where-Object { $_ -like '*保存先に書き込めないため*' }) }
    Assert-Equal '保存先を作れないときの警告が出る' ($saveWarnings.Count -gt 0) 'True'
    Assert-Equal '警告に改行が入らない' (@($saveWarnings | Where-Object { $_.Contains('<CR>') -or $_.Contains('<LF>') }).Count) 0

    Write-Section 'Ctrl+C で中断（C:\ を調べている途中で停止）'
    # Ctrl+C と同じ「パイプラインの停止」を、対象の PowerShell の中から送る
    $stopScript = Join-Path $workRoot 'stop-during-scan.ps1'
    $stopOut    = Join-Path $workRoot 'out\ctrlc'
    [IO.File]::WriteAllLines($stopScript, [string[]]@(
        'param($Tool, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $env:SystemDrive + ''\'').AddParameter(''Depth'', 100).AddParameter(''Detail'', $true).AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'Start-Sleep -Seconds 3'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + (($ps.Streams.Information | ForEach-Object { $_.MessageData.ToString() }) -join '' / '')'
    ), (New-Object Text.UTF8Encoding $true))
    $result = @{}
    & $Exe -NoProfile -ExecutionPolicy Bypass -File $stopScript -Tool $tool -OutRoot $stopOut 2>&1 | ForEach-Object {
        $line = "$_"
        $eq = $line.IndexOf('=')
        if ($eq -gt 0) { $result[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
    }
    if ($result['RUNNING'] -eq 'True') {
        Assert-Equal '止める前は保存先を作ってある' $result['BEFORE'] '1'
        Assert-Equal '止めた後は保存先が残らない' $result['AFTER'] '0'
        Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' '3 秒以内に調べ終わったため、途中で止められなかった'
    }
} finally {
    # アクセス拒否と subst を戻し、リンクを先に外してから消す
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    if ($substDrive) { $null = & subst.exe $substDrive /D 2>&1 }
    Remove-Item Env:\WORKBENCH_TEST_ROOT -ErrorAction SilentlyContinue
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            foreach ($link in $links) { if (Test-Path -LiteralPath $link) { [IO.Directory]::Delete($link) } }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

exit $script:failCount
