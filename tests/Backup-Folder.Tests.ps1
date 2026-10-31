<#
.SYNOPSIS
    Backup-Folder の回帰テストです。

.DESCRIPTION
    一時フォルダにダミー環境（日本語とスペースを含むパス、260 文字を超えるパス、ジャンクションのループ、
    読み取り禁止のフォルダ、隠し・システム属性のフォルダ、壊れた文字のファイル名、設定ファイルなど）を作り、
    Backup-Folder でコピーして、コピー先と結果のファイルを期待値と照合します。ダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダにだけアクセス拒否を設定します（終了時に戻します）。
    ドライブのルートを再現するため、空いているドライブ文字に subst でダミー環境のフォルダを割り当てます（終了時に外します）。

    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。また、ツールが一時的な記録を残さないかを確かめるため、TEMP と TMP をダミー環境の中に向けます。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Backup-Folder を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Backup-Folder\Backup-Folder.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Backup-Folder_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool     = Join-Path $workRoot 'tool\Backup-Folder.ps1'
$toolTemp = Join-Path $workRoot 'temp'
$src      = Join-Path $workRoot 'コピー元 フォルダ'
$dst      = Join-Path $workRoot 'コピー先 USB\バックアップ'
$allFiles = 'files.csv,robocopy.log,skipped.csv,summary.txt'
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

function ConvertTo-LongPath([string]$FullPath) {
    if ($FullPath.StartsWith('\\?\')) { return $FullPath }
    if ($FullPath.StartsWith('\\')) { return '\\?\UNC\' + $FullPath.Substring(2) }
    return '\\?\' + $FullPath
}

# 260 文字を超えるパスも作れるように、\\?\ を付けて作る
function New-TestFile([string]$FilePath, [long]$Size, $Time = $null) {
    [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $FilePath.Substring(0, $FilePath.LastIndexOf('\'))))
    $stream = [IO.File]::Create((ConvertTo-LongPath $FilePath))
    $stream.SetLength($Size)
    $stream.Close()
    if ($Time) { [IO.File]::SetLastWriteTime((ConvertTo-LongPath $FilePath), $Time) }
}

# コンソールで実行すると、長い行は画面の幅で折り返され、出力に改行が入る。
# 画面の幅によって結果が変わらないように、文字を探す前に改行を取り除く
function ConvertTo-FlatText([string]$Text) { $Text -replace '\r?\n', '' }

# 確認の入力を待って止まらないように、いつも標準入力に InputText を渡す（何も渡さないと、コンソールからの入力を待つ）。
# Log には、画面への出力から改行を取り除いたものを入れる
function Invoke-Raw([string[]]$ToolArgs, [string]$InputText = '', [string]$ToolPath = $tool, [string[]]$HostArgs = @()) {
    $lines = @($InputText | & $Exe -NoProfile @HostArgs -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1 | ForEach-Object { "$_" })
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = (ConvertTo-FlatText ($lines -join "`n")) }
}

# 結果の保存先を out\<Label> にして動かし、作られた結果のフォルダを Dir に入れる
function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$InputText = '') {
    $out = Join-Path $workRoot "out\$Label"
    $run = Invoke-Raw -ToolArgs ($ToolArgs + @('-OutputFolder', $out)) -InputText $InputText
    [pscustomobject]@{ Code = $run.Code; Log = $run.Log; Dir = (Get-RunDir $out) }
}

# 設定ファイルの相対パスの基準を確かめるため、指定したフォルダから起動する
function Invoke-From([string]$WorkDir, [string[]]$ToolArgs) {
    Push-Location -LiteralPath $WorkDir
    try { Invoke-Raw -ToolArgs $ToolArgs } finally { Pop-Location }
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $lines = @('' | & cmd.exe /d /c $batch 2>&1 | ForEach-Object { "$_" })
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = (ConvertTo-FlatText ($lines -join "`n")) }
}

# 保存先のフォルダの下に作られた、いちばん新しい結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'Backup_*' -ErrorAction SilentlyContinue | Sort-Object CreationTimeUtc, Name)
    if ($found.Count -gt 0) { return $found[-1].FullName }
    $null
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
        Files   = @(Read-CsvOrEmpty 'files.csv')
        Skipped = @(Read-CsvOrEmpty 'skipped.csv')
        Summary = $summary
        Names   = Get-ResultFileNames $Dir
    }
}

function Get-FileText([string]$Dir, [string]$FileName) {
    if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir $FileName))) { return '' }
    [IO.File]::ReadAllText((Join-Path $Dir $FileName))
}

function Get-FirstLine([string]$Dir, [string]$FileName) {
    if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir $FileName))) { return '(なし)' }
    @([IO.File]::ReadAllLines((Join-Path $Dir $FileName)))[0]
}

function Get-LineCount([string]$Dir, [string]$FileName) {
    if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir $FileName))) { return -1 }
    @([IO.File]::ReadAllLines((Join-Path $Dir $FileName))).Count
}

# summary.txt の、Prefix で始まる最初の行
function Get-SummaryLine([string]$Summary, [string]$Prefix) {
    ($Summary -split '\r?\n') | Where-Object { $_.StartsWith($Prefix) } | Select-Object -First 1
}

function Test-Bom([string]$FilePath) {
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
}

function Get-BomFlags([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { '{0}={1}' -f $_.Name, (Test-Bom $_.FullName) })
}

# files.csv のファイルの列を、Base から見た相対パスにして並べる
function Get-RelativeFiles($Rows, [string]$Base) {
    $prefix = $Base.TrimEnd('\') + '\'
    Join-Sorted (@($Rows) | ForEach-Object {
        $path = $_.'ファイル'
        if ($path.StartsWith($prefix)) { $path.Substring($prefix.Length) } else { $path }
    })
}

# files.csv の行を「相対パス=列1/列2」にして並べる
function Get-RowText($Rows, [string]$Base, [string[]]$Columns) {
    $prefix = $Base.TrimEnd('\') + '\'
    Join-Sorted (@($Rows) | ForEach-Object {
        $row  = $_
        $path = $row.'ファイル'
        if ($path.StartsWith($prefix)) { $path = $path.Substring($prefix.Length) }
        $path + '=' + (($Columns | ForEach-Object { $row.$_ }) -join '/')
    })
}

# コピー先にない、またはサイズか更新日時が違うファイルの一覧（すべて同じなら空）
function Get-CopyMismatch([string]$From, [string]$To, [string[]]$Relatives) {
    $bad = foreach ($rel in $Relatives) {
        $a = [IO.FileInfo]::new((ConvertTo-LongPath ($From.TrimEnd('\') + '\' + $rel)))
        $b = [IO.FileInfo]::new((ConvertTo-LongPath ($To.TrimEnd('\') + '\' + $rel)))
        if (-not $b.Exists) { "$rel(ない)" }
        elseif ($a.Length -ne $b.Length) { "$rel(サイズ)" }
        elseif ([math]::Abs(($a.LastWriteTimeUtc - $b.LastWriteTimeUtc).TotalSeconds) -gt 2) { "$rel(更新日時)" }
    }
    Join-Sorted $bad
}

function Get-Attr([string]$Path) { [string][IO.File]::GetAttributes((ConvertTo-LongPath $Path)) }

# ツールが一時フォルダに残した robocopy の記録の数
function Get-TempLogCount { @(Get-ChildItem -LiteralPath $toolTemp -Filter 'Backup-Folder_*.log' -ErrorAction SilentlyContinue).Count }

# 自分自身にアクセス拒否を設定する。終了時に戻すため、設定したフォルダを覚えておく
function Add-Deny([string]$Path, [string]$Rights) {
    $null = & icacls.exe $Path /deny "*${userSid}:$Rights" 2>&1
    if ($LASTEXITCODE -ne 0) { return $false }
    $denied.Add($Path)
    $true
}

# 設定ファイルを書く。文字コードの判定を確かめるため、文字コードを選べるようにする
function New-ConfigFile([string]$FilePath, [string[]]$Lines, [Text.Encoding]$Encoding = (New-Object Text.UTF8Encoding $false)) {
    [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $FilePath.Substring(0, $FilePath.LastIndexOf('\'))))
    [IO.File]::WriteAllText((ConvertTo-LongPath $FilePath), ($Lines -join "`r`n"), $Encoding)
    $FilePath
}

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $used -notcontains $_ -and -not [IO.Directory]::Exists("${_}:\") } | Select-Object -First 1
}

# スパースファイル（容量を使わずに大きなファイルを作る）にする。できなければ $false
function Set-Sparse([string]$FilePath, [long]$Size) {
    $null = & fsutil.exe sparse setflag $FilePath 2>&1
    if (([IO.File]::GetAttributes($FilePath) -band [IO.FileAttributes]::SparseFile) -eq 0) { return $false }
    $stream = [IO.File]::Open($FilePath, 'Open', 'ReadWrite')
    try { $stream.SetLength($Size) } finally { $stream.Close() }
    $true
}

function Format-Arg([string]$Text) { '"' + $Text + '"' }

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
# 標準入力に渡す文字は、コンソールの入力の文字コードで書かれる。UTF-8 のとき先頭に BOM が付き、7 の Read-Host は
# それを外さないため、BOM を付けない UTF-8 にする
$savedInputEncoding = $null
try { $savedInputEncoding = [Console]::InputEncoding; [Console]::InputEncoding = New-Object Text.UTF8Encoding $false } catch { }
$savedTmp   = $env:TMP
$savedTemp  = $env:TEMP
$userSid    = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied     = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したフォルダ（終了時に戻す）
$readOnly   = New-Object 'System.Collections.Generic.List[string]'   # 読み取り専用にしたファイル（消す前に戻す）
$sparseFiles = New-Object 'System.Collections.Generic.List[string]'
$substDrive = $null

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool
    [void][IO.Directory]::CreateDirectory($toolTemp)
    $env:TMP  = $toolTemp
    $env:TEMP = $toolTemp

    $longRel = ('L' * 150) + '\' + ('M' * 150) + '\long.bin'
    New-TestFile "$src\a.txt" 100 (Get-Date '2025-01-02 03:04:05')
    New-TestFile "$src\日本語 ファイル😀.txt" 200
    New-TestFile "$src\sub\deep\b.bin" 300
    New-TestFile "$src\sub\big.bin" 1234567
    New-TestFile "$src\sub\hidden.txt" 50
    [IO.File]::SetAttributes("$src\sub\hidden.txt", 'Hidden')
    New-TestFile "$src\$longRel" 400
    New-TestFile "$src\Thumbs.db" 5
    New-TestFile "$src\sub\Thumbs.db" 5
    New-TestFile "$src\desktop.ini" 5
    New-TestFile "$src\~`$doc.docx" 5
    [void][IO.Directory]::CreateDirectory("$src\empty")
    $null = New-Item -ItemType Junction -Path "$src\jn" -Target $src
    $mainRels = @('a.txt', '日本語 ファイル😀.txt', 'sub\deep\b.bin', 'sub\big.bin', 'sub\hidden.txt', $longRel)
    Assert-Equal 'スクリプトのコピー' (Test-Path -LiteralPath $tool) 'True'

    Write-Section '-ListOnly（一覧だけを保存し、コピーしない）'
    $run = Invoke-Tool 'listonly' @('-Source', $src, '-Destination', $dst, '-ListOnly')
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal 'コピー先を作らない' (Test-Path -LiteralPath $dst) 'False'
    Assert-Equal '件数と容量の行' $run.Log.Contains('コピーするファイル 6 件（1.2 MB）  新規 6 件・上書き 0 件') 'True'
    Assert-Equal '一覧だけの行' $run.Log.Contains('一覧だけを保存しました。コピーはしていません') 'True'
    Assert-Equal '確認しない' $run.Log.Contains('コピーしますか') 'False'
    $x = Read-Result $run.Dir
    Assert-Equal '保存するファイル' $x.Names $allFiles
    Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'files.csv=True,robocopy.log=True,skipped.csv=True,summary.txt=True'
    Assert-Equal 'files.csv の見出し' (Get-FirstLine $run.Dir 'files.csv') '"種類","結果","ファイル","サイズ(MB)","サイズ(バイト)","最終更新日時"'
    Assert-Equal 'files.csv のファイル（除外したファイルとリンクの先は載らない）' (Get-RelativeFiles $x.Files $src) (Join-Sorted $mainRels)
    Assert-Equal '種類と結果' (Join-Sorted ($x.Files | ForEach-Object { $_.'種類' + '/' + $_.'結果' } | Sort-Object -Unique)) '新規/一覧のみ'
    $big = @($x.Files | Where-Object { $_.'ファイル' -eq "$src\sub\big.bin" })
    Assert-Equal 'サイズの列（3 桁区切りと MB）' ('{0}/{1}' -f ($big | ForEach-Object { $_.'サイズ(バイト)' }), ($big | ForEach-Object { $_.'サイズ(MB)' })) '1,234,567/1.18'
    Assert-Equal '最終更新日時の列' (@($x.Files | Where-Object { $_.'ファイル' -eq "$src\a.txt" }) | ForEach-Object { $_.'最終更新日時' }) '2025/01/02 03:04:05'
    Assert-Equal 'skipped.csv は見出しだけ' ('{0}/{1}' -f (Get-FirstLine $run.Dir 'skipped.csv'), (Get-LineCount $run.Dir 'skipped.csv')) '"理由","パス","詳細"/1'
    $sumLines = @($x.Summary -split '\r?\n')
    Assert-Equal 'summary の 1 行目' $sumLines[0] 'フォルダのバックアップ (Backup-Folder 1.0.0)'
    Assert-Equal 'summary のコピー元' $sumLines[1] "コピー元: $src"
    Assert-Equal 'summary のコピー先（形式と空き）' ("$($sumLines[2])" -match ('^' + [regex]::Escape("コピー先: $dst（") + '\w+・空き [\d.]+ [KMGTP]?B）$')) 'True'
    Assert-Equal 'summary の日時と管理者' ("$($sumLines[3])" -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   \S+（管理者: (はい|いいえ)）$') 'True'
    Assert-Equal 'robocopy.log は調べたときの記録' (Get-FileText $run.Dir 'robocopy.log').Contains("$src\a.txt") 'True'

    Write-Section '確認で N（コピー元とコピー先を位置で指定）'
    $run = Invoke-Tool 'answer-n' @($src, $dst) 'N'
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal 'コピーしませんでした' $run.Log.Contains('コピーしませんでした') 'True'
    Assert-Equal 'コピー先を作らない' (Test-Path -LiteralPath $dst) 'False'
    Assert-Equal '結果を保存しない' ($null -eq $run.Dir) 'True'

    Write-Section '確認で Y（はじめに違う文字を入力）'
    $run = Invoke-Tool 'answer-y' @($src, $dst) "x`r`ny"
    Assert-Equal '終了コード' $run.Code 0
    $x = Read-Result $run.Dir
    Assert-Equal '結果の行' ((Get-SummaryLine $x.Summary '結果: ') -match '^結果: 正常に終わりました  コピー 6 件  失敗 0 件  所要 [\d.]+秒$') 'True'
    Assert-Equal 'files.csv の結果' (Join-Sorted ($x.Files | ForEach-Object { $_.'結果' } | Sort-Object -Unique)) 'コピーした'
    Assert-Equal 'コピー先のファイル（サイズと更新日時）' (Get-CopyMismatch $src $dst $mainRels) ''
    Assert-Equal 'コピーしないファイル（Thumbs.db・desktop.ini・~$）' (Join-Sorted (@('Thumbs.db', 'sub\Thumbs.db', 'desktop.ini', '~$doc.docx') | Where-Object { Test-Path -LiteralPath "$dst\$_" })) ''
    Assert-Equal 'ジャンクションをたどらない' (Test-Path -LiteralPath "$dst\jn") 'False'
    Assert-Equal '空のフォルダもコピー' ([IO.Directory]::Exists("$dst\empty")) 'True'
    Assert-Equal '隠しファイルの属性もコピー' (([IO.File]::GetAttributes("$dst\sub\hidden.txt") -band [IO.FileAttributes]::Hidden) -ne 0) 'True'
    Assert-Equal 'コピー先のフォルダは見える' (Get-Attr $dst) 'Directory'
    Assert-Equal '一時的な記録の行を出さない' ('{0}/{1}' -f $run.Log.Contains($toolTemp), ($run.Log -match 'Log File|ログ ファイル')) 'False/False'
    Assert-Equal '一時的な記録を残さない' (Get-TempLogCount) 0
    Assert-Equal 'robocopy.log はコピーの記録' (Get-FileText $run.Dir 'robocopy.log').Contains("$src\a.txt") 'True'

    Write-Section 'もう一度（すべてコピー済み）'
    $run = Invoke-Tool 'again' @($src, $dst)
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal 'コピーするファイルはない' $run.Log.Contains('コピーするファイルはありません（すべてコピー済みです）') 'True'
    Assert-Equal '確認しない' $run.Log.Contains('コピーしますか') 'False'
    $x = Read-Result $run.Dir
    Assert-Equal '保存するファイル' $x.Names $allFiles
    Assert-Equal 'files.csv は見出しだけ' (Get-LineCount $run.Dir 'files.csv') 1

    Write-Section '差分（サイズの変更・コピー先のほうが新しい・新規・コピー先にだけあるファイル）'
    New-TestFile "$src\a.txt" 150
    [IO.File]::SetLastWriteTime((ConvertTo-LongPath "$dst\sub\deep\b.bin"), (Get-Date).AddDays(1))
    New-TestFile "$src\new2.txt" 10
    New-TestFile "$dst\only-in-dest.txt" 7
    $run = Invoke-Tool 'diff' @($src, $dst, '-Force')
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '件数と容量の行' $run.Log.Contains('コピーするファイル 3 件（460 B）  新規 1 件・上書き 2 件') 'True'
    Assert-Equal 'files.csv' (Get-RowText (Read-Result $run.Dir).Files $src @('種類', '結果')) (Join-Sorted @('a.txt=上書き/コピーした', 'new2.txt=新規/コピーした', 'sub\deep\b.bin=上書き/コピーした'))
    Assert-Equal 'コピー先のファイル' (Get-CopyMismatch $src $dst @('a.txt', 'new2.txt', 'sub\deep\b.bin')) ''
    Assert-Equal 'コピー先にだけあるファイルを消さない' (Test-Path -LiteralPath "$dst\only-in-dest.txt") 'True'

    Write-Section '更新日時の違いは 2 秒まで同じとみなす'
    $srcTime = [IO.File]::GetLastWriteTimeUtc("$src\a.txt")
    [IO.File]::SetLastWriteTimeUtc("$dst\a.txt", $srcTime.AddSeconds(1))
    $run = Invoke-Tool 'time1' @($src, $dst, '-ListOnly')
    Assert-Equal '1 秒の違いはコピーしない' ('{0}/{1}' -f $run.Code, $run.Log.Contains('コピーするファイルはありません（すべてコピー済みです）')) '0/True'
    [IO.File]::SetLastWriteTimeUtc("$dst\a.txt", $srcTime.AddSeconds(3))
    $run = Invoke-Tool 'time3' @($src, $dst, '-ListOnly')
    Assert-Equal '3 秒の違いはコピーする' ('{0}/{1}' -f $run.Code, (Get-RowText (Read-Result $run.Dir).Files $src @('種類'))) '0/a.txt=上書き'
    [IO.File]::SetLastWriteTimeUtc("$dst\a.txt", $srcTime)

    Write-Section '新しい空のフォルダだけ'
    [void][IO.Directory]::CreateDirectory("$src\newempty")
    $run = Invoke-Tool 'newdir-list' @($src, $dst, '-ListOnly')
    Assert-Equal '一覧' ('{0}/{1}' -f $run.Code, $run.Log.Contains('コピーするファイルはありません。新しく作るフォルダ 1 個')) '0/True'
    $run = Invoke-Tool 'newdir' @($src, $dst, '-Force')
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal 'フォルダを作る' ([IO.Directory]::Exists("$dst\newempty")) 'True'
    Assert-Equal '結果の行' ((Get-SummaryLine (Read-Result $run.Dir).Summary '結果: ') -match '^結果: 正常に終わりました  コピー 0 件  失敗 0 件') 'True'

    Write-Section '確認できないとき（入力の終わり・-NonInteractive）'
    New-TestFile "$src\pending.txt" 5
    $cannotAsk = '確認できないため、コピーしません。確認せずにコピーするときは -Force を付けてください'
    $run = Invoke-Tool 'eof' @($src, $dst) ''
    Assert-Equal '入力の終わり' ('{0}/{1}' -f $run.Code, $run.Log.Contains($cannotAsk)) '1/True'
    $run = Invoke-Raw -ToolArgs @($src, $dst, '-OutputFolder', "$workRoot\out\noninteractive") -HostArgs @('-NonInteractive')
    Assert-Equal '-NonInteractive' ('{0}/{1}' -f $run.Code, $run.Log.Contains($cannotAsk)) '1/True'
    Assert-Equal 'コピーしていない' (Test-Path -LiteralPath "$dst\pending.txt") 'False'

    Write-Section '止まるとき（コピーを始める前・終了コード 1・結果を保存しない）'
    New-TestFile "$workRoot\afile" 1
    # スペースを含むパスの最後が \ だと、PowerShell が外部コマンドに渡すときに \" になってしまうため、スペースのないフォルダで試す
    $plain = Join-Path $workRoot 'plain'
    New-TestFile "$plain\p.txt" 1
    $stops = @(
        @('コピー元なし', @('-Destination', $dst), 'コピー元を指定してください（-Source で指定するか、設定ファイルの Source に書きます）'),
        @('コピー先なし', @('-Source', $src), 'コピー先を指定してください（-Destination で指定するか、設定ファイルの Destination に書きます）'),
        @('コピー元が見つからない', @("$workRoot\no\such", $dst), "コピー元のフォルダを開けません: $workRoot\no\such"),
        @('共有フォルダが見つからない', @('\\localhost\workbench_no_such_share\x', $dst), 'コピー元のフォルダを開けません: \\localhost\workbench_no_such_share\x共有フォルダのときは、次の順に確かめてください。'),
        @('コピー元とコピー先が同じ（末尾の \）', @($plain, "$plain\"), "コピー元とコピー先が同じです: $plain"),
        @('コピー元とコピー先が同じ（大文字と小文字）', @($src, $src.ToUpperInvariant()), "コピー元とコピー先が同じです: $src"),
        @('コピー先がコピー元の中', @($src, "$src\backup"), 'コピー先がコピー元の中にあります。コピーを繰り返すたびに、コピー先の中身もコピーされてしまいます。'),
        @('コピー先が同じ名前のファイル', @($src, "$workRoot\afile"), "コピー先のフォルダを作れません: $workRoot\afile（同じ名前のファイルがあります）"),
        @('コピー先のフォルダを作れない', @($src, "$workRoot\afile\x", '-Force'), "コピー先のフォルダを作れません: $workRoot\afile\x（")
    )
    for ($i = 0; $i -lt $stops.Count; $i++) {
        if ($stops[$i][0] -like '*末尾の \*' -and $plain.Contains(' ')) { Write-Skip $stops[$i][0] '一時フォルダのパスにスペースがあるため'; continue }
        $run = Invoke-Tool "stop$i" $stops[$i][1]
        Assert-Equal $stops[$i][0] ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains($stops[$i][2]), ($null -eq $run.Dir)) '1/True/True'
    }
    $freeLetter = Get-FreeLetter
    if ($freeLetter) {
        $run = Invoke-Tool 'stop-drive' @($src, "${freeLetter}:\backup")
        Assert-Equal 'コピー先のドライブがない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("コピー先のドライブ ${freeLetter}: が見つかりません。USB メモリなどがつながっているか確かめてください")) '1/True'
    } else {
        Write-Skip 'コピー先のドライブがない' '空いているドライブ文字がないため'
    }
    New-TestFile "$workRoot\locked-src\f.txt" 1
    if (Add-Deny "$workRoot\locked-src" '(RD)') {
        $run = Invoke-Tool 'stop-locked' @("$workRoot\locked-src", "$workRoot\locked-dst")
        Assert-Equal 'コピー元を開けない（権限なし）' ('{0}/{1}' -f $run.Code, $run.Log.Contains("コピー元のフォルダを開けません: $workRoot\locked-src（")) '1/True'
    } else {
        Write-Skip 'コピー元を開けない（権限なし）' 'アクセス拒否を設定できなかったため'
    }

    Write-Section '読めないフォルダ（権限なし）'
    $dn = Join-Path $workRoot 'dn'
    New-TestFile "$dn\ok.txt" 1
    New-TestFile "$dn\denied\secret.txt" 2
    if (Add-Deny "$dn\denied" '(RD)') {
        $run = Invoke-Tool 'dn-list' @($dn, "$workRoot\dn-dst", '-ListOnly')
        Assert-Equal '-ListOnly: 終了コード' $run.Code 0
        Assert-Equal '-ListOnly: ※ の行' $run.Log.Contains('※ 読めなかったフォルダが 1 件あります。そのフォルダの中はコピーされません（skipped.csv を確認してください）') 'True'
        $x = Read-Result $run.Dir
        Assert-Equal 'skipped.csv' (Join-Sorted ($x.Skipped | ForEach-Object { $_.'理由' + '|' + $_.'パス' })) "権限なし|$dn\denied"
        Assert-Equal 'skipped.csv の詳細' (@($x.Skipped | Where-Object { $_.'詳細' }).Count) 1
        $run = Invoke-Tool 'dn-copy' @($dn, "$workRoot\dn-dst", '-Force')
        Assert-Equal 'コピー: 終了コード' $run.Code 1
        Assert-Equal 'コピー: 結果の行' ((Get-SummaryLine (Read-Result $run.Dir).Summary '結果: ') -match '^結果: 読めなかったフォルダが 1 件あります。そのフォルダの中はコピーしていません。skipped\.csv を確認してください（コピー 1 件  所要 [\d.]+秒）$') 'True'
        Assert-Equal 'コピー: 読めるファイルはコピー' (Get-CopyMismatch $dn "$workRoot\dn-dst" @('ok.txt')) ''
        # コピーするものがない日も 1 にしないと、タスクスケジューラで見る結果が日によって変わる
        $run = Invoke-Tool 'dn-again' @($dn, "$workRoot\dn-dst", '-Force')
        Assert-Equal 'コピーするものがなくても終了コード 1' ('{0}/{1}' -f $run.Code, $run.Log.Contains('コピーするファイルはありません（すべてコピー済みです）')) '1/True'
        $run = Invoke-Tool 'dn-list2' @($dn, "$workRoot\dn-dst", '-ListOnly')
        Assert-Equal '-ListOnly は終了コード 0' $run.Code 0
    } else {
        Write-Skip '読めないフォルダ' 'アクセス拒否を設定できなかったため'
    }

    Write-Section '名前の変わったファイルとフォルダ（壊れた文字・末尾の . と空白・[ ]・バッククォート）'
    $names    = Join-Path $workRoot 'names'
    $namesDst = Join-Path $workRoot 'names-dst'
    $bad      = [string][char]0xD800
    New-TestFile "$names\b$bad.txt" 3
    New-TestFile "$names\[y].txt" 4
    New-TestFile "$names\back``tick.txt" 5
    New-TestFile "$names\d$bad\f.txt" 6
    $emptyNames = @('[x]', '末尾 ', '末尾.', "s$bad")
    foreach ($dirName in $emptyNames) { [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath "$names\$dirName")) }
    $run = Invoke-Tool 'names' @($names, $namesDst, '-Force')
    Assert-Equal '終了コード' $run.Code 0
    $x = Read-Result $run.Dir
    Assert-Equal 'files.csv の件数と結果' ('{0}/{1}' -f $x.Files.Count, (Join-Sorted ($x.Files | ForEach-Object { $_.'結果' } | Sort-Object -Unique))) '4/コピーした'
    Assert-Equal '空のフォルダを読めなかったと誤って判定しない' $x.Skipped.Count 0
    Assert-Equal 'コピー先のファイル' (Get-CopyMismatch $names $namesDst @("b$bad.txt", '[y].txt', 'back`tick.txt', "d$bad\f.txt")) ''
    Assert-Equal 'コピー先の空のフォルダ' (Join-Sorted ($emptyNames | Where-Object { -not [IO.Directory]::Exists((ConvertTo-LongPath "$namesDst\$_")) })) ''
    $run = Invoke-Tool 'names-again' @($names, $namesDst, '-ListOnly')
    Assert-Equal '2 回目はすべてコピー済み' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('コピーするファイルはありません（すべてコピー済みです）'), (Read-Result $run.Dir).Skipped.Count) '0/True/0'

    Write-Section '読み取り専用のファイルの上書き'
    $ro    = Join-Path $workRoot 'ro'
    $roDst = Join-Path $workRoot 'ro-dst'
    New-TestFile "$ro\ro.txt" 10
    [IO.File]::SetAttributes("$ro\ro.txt", 'ReadOnly')
    $readOnly.Add("$ro\ro.txt")
    $readOnly.Add("$roDst\ro.txt")
    $run = Invoke-Tool 'ro1' @($ro, $roDst, '-Force')
    Assert-Equal '1 回目' ('{0}/{1}' -f $run.Code, (([IO.File]::GetAttributes("$roDst\ro.txt") -band [IO.FileAttributes]::ReadOnly) -ne 0)) '0/True'
    [IO.File]::SetAttributes("$ro\ro.txt", 'Normal')
    New-TestFile "$ro\ro.txt" 20
    [IO.File]::SetAttributes("$ro\ro.txt", 'ReadOnly')
    $run = Invoke-Tool 'ro2' @($ro, $roDst, '-Force')
    Assert-Equal '読み取り専用のコピー先を上書き' ('{0}/{1}/{2}' -f $run.Code, (Get-RowText (Read-Result $run.Dir).Files $ro @('種類', '結果')), (Get-CopyMismatch $ro $roDst @('ro.txt'))) '0/ro.txt=上書き/コピーした/'

    Write-Section 'コピー中の中断（Ctrl+C）と、コピー先のフォルダの隠し・システム属性'
    # ドライブのルートには隠し・システム属性が付いていて、robocopy はそれをコピー先のフォルダにも付ける。
    # 同じ属性を付けたフォルダをコピー元にして再現する
    $hs    = Join-Path $workRoot 'hs'
    $hsDst = Join-Path $workRoot 'hs-dst'
    New-TestFile "$hs\a.txt" 10
    New-TestFile "$hs\sub\locked.txt" 20
    [IO.File]::SetAttributes($hs, 'Directory, Hidden, System')
    # Ctrl+C と同じ「パイプラインの停止」を、対象の PowerShell の中から送る。サブフォルダのファイルをロックして
    # robocopy を待たせ、コピー先にルートの属性が付いた後で止める
    $stopScript = Join-Path $workRoot 'stop-during-copy.ps1'
    [IO.File]::WriteAllLines($stopScript, [string[]]@(
        'param($Tool, $Source, $Destination, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Source'', $Source).AddParameter(''Destination'', $Destination).AddParameter(''Force'', $true).AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        '$sw     = [Diagnostics.Stopwatch]::StartNew()'
        '$during = '''''
        'while (-not $handle.IsCompleted -and $sw.Elapsed.TotalSeconds -lt 60) {'
        '    if ([IO.Directory]::Exists($Destination) -and ([IO.File]::GetAttributes($Destination) -band [IO.FileAttributes]::Hidden)) { $during = [string][IO.File]::GetAttributes($Destination); break }'
        '    Start-Sleep -Milliseconds 100'
        '}'
        '''DURING='' + $during'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''ATTR='' + [IO.File]::GetAttributes($Destination)'
        '''MESSAGE='' + (($ps.Streams.Information | ForEach-Object { $_.MessageData.ToString() } | Where-Object { $_ -like ''*中断*'' }) -join '' / '')'
        '$left = 1'
        'for ($i = 0; $i -lt 50 -and $left -gt 0; $i++) {'
        '    $left = @(Get-CimInstance Win32_Process -Filter "Name=''Robocopy.exe''" | Where-Object { "$($_.CommandLine)".Contains($Destination) }).Count'
        '    if ($left -gt 0) { Start-Sleep -Milliseconds 100 }'
        '}'
        '''ROBOCOPY='' + $left'
    ), (New-Object Text.UTF8Encoding $true))
    $result = @{}
    $lock = [IO.File]::Open("$hs\sub\locked.txt", 'Open', 'ReadWrite', 'None')
    try {
        & $Exe -NoProfile -ExecutionPolicy Bypass -File $stopScript -Tool $tool -Source $hs -Destination $hsDst -OutRoot "$workRoot\out\stop" 2>&1 | ForEach-Object {
            $line = "$_"
            $eq = $line.IndexOf('=')
            if ($eq -gt 0) { $result[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
        }
    } finally {
        $lock.Close()
    }
    if ($result['DURING']) {
        Assert-Equal 'コピー中はコピー元の属性が付いている' $result['DURING'] 'Hidden, System, Directory'
        Assert-Equal '止める前は保存先を作ってある' $result['BEFORE'] '1'
        Assert-Equal '止めた後は保存先が残らない' $result['AFTER'] '0'
        Assert-Equal '止めた後は属性を外す' $result['ATTR'] 'Directory'
        Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。コピーし終わったファイルは、コピー先に残っています。もう一度実行すると、残りのファイルだけをコピーします。'
        Assert-Equal 'robocopy が残らない' $result['ROBOCOPY'] '0'
        Assert-Equal '一時的な記録を残さない' (Get-TempLogCount) 0
    } else {
        Write-Skip 'コピー中の中断' 'コピー先に属性が付く前にコピーが終わったため'
    }
    $run = Invoke-Tool 'hs-copy' @($hs, $hsDst, '-Force')
    Assert-Equal 'もう一度: 終了コード' $run.Code 0
    Assert-Equal 'もう一度: 残りのファイルをコピー' (Get-CopyMismatch $hs $hsDst @('a.txt', 'sub\locked.txt')) ''
    Assert-Equal 'もう一度: コピー先のフォルダは見える' (Get-Attr $hsDst) 'Directory'
    # ウィンドウを閉じるなどで強制終了すると、属性が残る。コピーするものがない日の実行でも外す
    [IO.File]::SetAttributes($hsDst, 'Directory, Hidden, System')
    $run = Invoke-Tool 'hs-nothing' @($hs, $hsDst, '-Force')
    Assert-Equal '属性が残っていて、コピーするものがない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('コピーするファイルはありません（すべてコピー済みです）'), (Get-Attr $hsDst)) '0/True/Directory'
    # 確認で N のときと -ListOnly のときは、コピー先を変えない
    [IO.File]::SetAttributes($hsDst, 'Directory, Hidden, System')
    New-TestFile "$hs\new.txt" 5
    $run = Invoke-Tool 'hs-n' @($hs, $hsDst) 'N'
    Assert-Equal '確認で N: 属性を変えない' ('{0}/{1}' -f $run.Code, (Get-Attr $hsDst)) '0/Hidden, System, Directory'
    $run = Invoke-Tool 'hs-list' @($hs, $hsDst, '-ListOnly')
    Assert-Equal '-ListOnly: 属性を変えない' ('{0}/{1}' -f $run.Code, (Get-Attr $hsDst)) '0/Hidden, System, Directory'
    $run = Invoke-Tool 'hs-copy2' @($hs, $hsDst, '-Force')
    Assert-Equal 'コピーした後は属性を外す' ('{0}/{1}' -f $run.Code, (Get-Attr $hsDst)) '0/Directory'

    Write-Section 'コピー中にコピー元のファイルが更新される'
    $mod    = Join-Path $workRoot 'mod'
    $modDst = Join-Path $workRoot 'mod-dst'
    New-TestFile "$mod\f.txt" 6
    # ファイルをロックして robocopy を待たせ、その間に書き換えてからロックを外す
    $lock = [IO.File]::Open("$mod\f.txt", 'Open', 'ReadWrite', 'None')
    $proc = $null
    try {
        $procArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Format-Arg $tool), (Format-Arg $mod), (Format-Arg $modDst), '-Force', '-OutputFolder', (Format-Arg "$workRoot\out\mod"))
        $proc = Start-Process -FilePath $Exe -ArgumentList $procArgs -PassThru -WindowStyle Hidden
        $null = $proc.Handle   # 5.1 は、先にハンドルを取っておかないと終了コードを取れない
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while (-not [IO.Directory]::Exists($modDst) -and $sw.Elapsed.TotalSeconds -lt 30) { Start-Sleep -Milliseconds 100 }
        Start-Sleep -Seconds 1
        $bytes = [Text.Encoding]::ASCII.GetBytes('changed during backup')
        $lock.SetLength(0)
        $lock.Write($bytes, 0, $bytes.Length)
    } finally {
        $lock.Close()
    }
    if ($proc) {
        $proc.WaitForExit()
        $x = Read-Result (Get-RunDir "$workRoot\out\mod")
        Assert-Equal '「コピーした」と判定' ('{0}/{1}' -f $proc.ExitCode, (Get-RowText $x.Files $mod @('結果'))) '0/f.txt=コピーした'
        Assert-Equal 'コピー先は更新後の内容' ([IO.File]::ReadAllText("$modDst\f.txt")) 'changed during backup'
    } else {
        Write-Skip 'コピー中の更新' 'ツールを起動できなかったため'
    }

    Write-Section '空き容量（容量を使わないスパースファイルで再現）'
    $bigDir  = Join-Path $workRoot 'big'
    $sparse  = Join-Path $bigDir 'huge.bin'
    $fullDst = Join-Path $workRoot 'full-dst'
    New-TestFile $sparse 0
    $sparseFiles.Add($sparse)
    [void][IO.Directory]::CreateDirectory($fullDst)
    # 万一コピーが始まってもディスクを埋めないよう、コピー先への書き込みを拒否しておく
    $guarded = Add-Deny $fullDst '(OI)(CI)(WD,AD)'
    $free    = (New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($workRoot))).AvailableFreeSpace
    if ($guarded -and (Set-Sparse $sparse ($free + 100GB))) {
        $run = Invoke-Tool 'full' @($bigDir, $fullDst, '-Force')
        Assert-Equal '-Force でも止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('コピー先の空きが足りません（必要 '), ($null -eq $run.Dir)) '1/True/True'
        Assert-Equal 'コピーしていない' @([IO.Directory]::GetFileSystemEntries($fullDst)).Count 0
        $run = Invoke-Tool 'full-list' @($bigDir, $fullDst, '-ListOnly')
        Assert-Equal '-ListOnly は知らせるだけ' ('{0}/{1}' -f $run.Code, $run.Log.Contains('※ コピー先の空きが足りません（必要 ')) '0/True'
    } else {
        Write-Skip '空き容量' 'スパースファイルかアクセス拒否を設定できなかったため'
    }
    $fat = @([IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -ne 'Network' -and $_.IsReady -and $_.DriveFormat -match '^FAT(32)?$' }) | Select-Object -First 1
    $fatSrc = Join-Path $workRoot 'fat'
    New-TestFile "$fatSrc\big5g.bin" 0
    $sparseFiles.Add("$fatSrc\big5g.bin")
    if ($fat -and (Set-Sparse "$fatSrc\big5g.bin" 5GB)) {
        # -ListOnly なので、FAT32 のドライブには何も書き込まない
        $run = Invoke-Tool 'fat32' @($fatSrc, ($fat.RootDirectory.FullName + 'workbench-test-' + [Guid]::NewGuid().ToString('N')), '-ListOnly')
        Assert-Equal 'FAT32 の 4 GB の警告' ('{0}/{1}' -f $run.Code, $run.Log.Contains('※ コピー先は FAT32 のため、4 GB 以上のファイル 1 件はコピーできません')) '0/True'
    } else {
        Write-Skip 'FAT32 の 4 GB の警告' 'FAT32 のドライブがないため'
    }
    # 以降の実行で大きなファイルをコピーしないよう、大きさを 0 に戻す
    foreach ($file in $sparseFiles) { try { $stream = [IO.File]::Open($file, 'Open', 'ReadWrite'); $stream.SetLength(0); $stream.Close() } catch { } }

    Write-Section 'ドライブのルート（subst で再現）'
    $rootSim = Join-Path $workRoot 'rootsim'
    New-TestFile "$rootSim\a.txt" 1
    New-TestFile "$rootSim\`$RECYCLE.BIN\trash.txt" 2
    New-TestFile "$rootSim\System Volume Information\svi.txt" 3
    New-TestFile "$rootSim\sub\`$RECYCLE.BIN\keep.txt" 4
    $rootRels = @('a.txt', 'sub\$RECYCLE.BIN\keep.txt')
    $letter = Get-FreeLetter
    if ($letter) {
        $null = & subst.exe "${letter}:" $rootSim 2>&1
        if ($LASTEXITCODE -eq 0) { $substDrive = "${letter}:" }
    }
    if ($substDrive) {
        # ドライブ文字だけ（D:）を指定しても、ルートとして扱う
        $run = Invoke-Tool 'root-src' @($substDrive, "$workRoot\root-src-dst", '-ListOnly')
        Assert-Equal 'コピー元: ドライブ文字だけ' ('{0}/{1}' -f $run.Code, $run.Log.Contains("コピー元: $substDrive\")) '0/True'
        Assert-Equal 'コピー元: 直下の $RECYCLE.BIN と System Volume Information を除く' (Get-RelativeFiles (Read-Result $run.Dir).Files "$substDrive\") (Join-Sorted $rootRels)
        $null = & subst.exe $substDrive /D 2>&1
        $rootDst = Join-Path $workRoot 'rootdst'
        [void][IO.Directory]::CreateDirectory($rootDst)
        $null = & subst.exe $substDrive $rootDst 2>&1
        if ($LASTEXITCODE -ne 0) { $substDrive = $null }
    }
    if ($substDrive) {
        $toRoot = Join-Path $workRoot 'toroot'
        [void][IO.Directory]::CreateDirectory("$toRoot\newdir")
        $run = Invoke-Tool 'root-dst-list' @($toRoot, "$substDrive\", '-ListOnly')
        Assert-Equal 'コピー先: 新しく作るフォルダ' ('{0}/{1}' -f $run.Code, $run.Log.Contains('コピーするファイルはありません。新しく作るフォルダ 1 個')) '0/True'
        New-TestFile "$toRoot\b.txt" 5
        $run = Invoke-Tool 'root-dst' @($toRoot, "$substDrive\", '-Force')
        Assert-Equal 'コピー先: コピー' ('{0}/{1}/{2}' -f $run.Code, [IO.File]::Exists("$rootDst\b.txt"), [IO.Directory]::Exists("$rootDst\newdir")) '0/True/True'
        $run = Invoke-Tool 'root-dst-again' @($toRoot, "$substDrive\", '-ListOnly')
        Assert-Equal 'コピー先: 2 回目はすべてコピー済み' ('{0}/{1}' -f $run.Code, $run.Log.Contains('コピーするファイルはありません（すべてコピー済みです）')) '0/True'
        # 後のテストで「ドライブがない」に使うドライブ文字と重ならないよう、すぐに外す
        $null = & subst.exe $substDrive /D 2>&1
        $substDrive = $null
    } else {
        Write-Skip 'ドライブのルート' 'subst でドライブ文字を割り当てられなかったため'
    }

    Write-Section '共有フォルダ経由（\\localhost\<ドライブ>$）'
    $uncRoot = '\\localhost\{0}${1}' -f $rootSim.Substring(0, 1), $rootSim.Substring(2)
    $uncDst  = Join-Path $workRoot 'unc-dst'
    if (Test-Path -LiteralPath $uncRoot) {
        # ドライブのルートを共有したフォルダ（\\server\D$）は、パスの形ではルートと分からない
        $run = Invoke-Tool 'unc-list' @($uncRoot, $uncDst, '-ListOnly')
        Assert-Equal '直下の $RECYCLE.BIN と System Volume Information を除く' ('{0}/{1}' -f $run.Code, (Get-RelativeFiles (Read-Result $run.Dir).Files $uncRoot)) ('0/' + (Join-Sorted $rootRels))
        $run = Invoke-Tool 'unc-copy' @($uncRoot, $uncDst, '-Force')
        Assert-Equal 'コピー' ('{0}/{1}' -f $run.Code, (Get-CopyMismatch $rootSim $uncDst $rootRels)) '0/'
        $run = Invoke-Tool 'unc-again' @($uncRoot, $uncDst, '-ListOnly')
        Assert-Equal '2 回目はすべてコピー済み' ('{0}/{1}' -f $run.Code, $run.Log.Contains('コピーするファイルはありません（すべてコピー済みです）')) '0/True'
    } else {
        Write-Skip '共有フォルダ経由' '管理共有を読めないため'
    }

    Write-Section '260 文字を超えるパスと、\\?\ 付きのパス'
    $small = Join-Path $workRoot 'small'
    New-TestFile "$small\s.txt" 3
    $run = Invoke-Tool 'long-src' @(("$src\" + (Split-Path $longRel)), "$workRoot\long-src-dst", '-ListOnly')
    Assert-Equal '260 文字を超えるコピー元' ('{0}/{1}' -f $run.Code, $run.Log.Contains('コピーするファイル 1 件')) '0/True'
    $longDst = "$workRoot\" + ('D' * 200) + '\' + ('E' * 100)
    $run = Invoke-Tool 'long-dst' @($small, $longDst, '-Force')
    Assert-Equal '260 文字を超えるコピー先' ('{0}/{1}' -f $run.Code, (Get-CopyMismatch $small $longDst @('s.txt'))) '0/'
    $run = Invoke-Tool 'prefix' @(('\\?\' + $small), "$workRoot\prefix-dst", '-ListOnly')
    Assert-Equal '\\?\ 付きのコピー元' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("コピー元: $small"), $run.Log.Contains('\\?\')) '0/True/False'

    Write-Section '名前に [ ] を含むフォルダから、相対パスで起動'
    # 5.1 は、名前に [ ] を含むフォルダで起動すると、名前をワイルドカードとして扱い、1 つだけ一致した隣のフォルダから始まる
    $sib = Join-Path $workRoot 'start\sib'
    New-TestFile "$sib\[a]\data\f.txt" 11
    New-TestFile "$sib\a\data\decoy.txt" 22
    $run = Invoke-From "$sib\[a]" @('-Source', 'data', '-Destination', '..\rel-dst', '-OutputFolder', 'out', '-ListOnly')
    Assert-Equal 'コピー元とコピー先' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("コピー元: $sib\[a]\data"), $run.Log.Contains("コピー先: $sib\rel-dst")) '0/True/True'
    Assert-Equal '隣のフォルダを読まない' (Get-RelativeFiles (Read-Result (Get-RunDir "$sib\[a]\out")).Files "$sib\[a]\data") 'f.txt'
    $null = New-ConfigFile "$sib\[a]\rel.ini" @('Source = data', 'Destination = ..\rel-cfg-dst', 'OutputFolder = out2')
    $run = Invoke-From "$sib\[a]" @('-ConfigFile', 'rel.ini', '-ListOnly')
    Assert-Equal '相対パスの -ConfigFile' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("設定ファイル: $sib\[a]\rel.ini"), (Get-ResultFileNames (Get-RunDir "$sib\[a]\out2"))) "0/True/$allFiles"

    Write-Section '結果の保存先がコピー元の中'
    New-TestFile "$workRoot\inout\f.txt" 1
    $run = Invoke-Raw -ToolArgs @("$workRoot\inout", "$workRoot\inout-dst", '-ListOnly', '-OutputFolder', "$workRoot\inout\記録")
    Assert-Equal '※ の行' ('{0}/{1}' -f $run.Code, $run.Log.Contains('※ 結果の保存先がコピー元の中にあるため、結果のフォルダもコピーされます（-OutputFolder で、コピー元の外を指定してください）')) '0/True'

    Write-Section '設定ファイル（-ConfigFile）'
    $cfgDir = Join-Path $workRoot 'cfg\設定 フォルダ'
    $cwd    = Join-Path $workRoot 'cwd'
    [void][IO.Directory]::CreateDirectory($cwd)
    $basic = New-ConfigFile "$cfgDir\usb.ini" @(
        '# 設定ファイルのテスト（日本語のコメント）'
        '; セミコロンのコメント'
        'source = ..\..\コピー元 フォルダ'
        'DESTINATION=..\..\cfg-dst'
        'Destination ='
        'OutputFolder = .\記録'
    )
    # 別のフォルダから起動して、相対パスが設定ファイルのフォルダを基準にすることを確かめる
    $run = Invoke-From $cwd @('-ConfigFile', $basic, '-ListOnly')
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal 'コピー元とコピー先（設定ファイルのフォルダが基準）' ('{0}/{1}' -f $run.Log.Contains("コピー元: $src"), $run.Log.Contains("コピー先: $workRoot\cfg-dst")) 'True/True'
    $runDir = Get-RunDir "$cfgDir\記録"
    Assert-Equal '保存先（設定ファイルのフォルダが基準）' (Get-ResultFileNames $runDir) $allFiles
    Assert-Equal '起動したフォルダには保存しない' (Test-Path -LiteralPath "$cwd\記録") 'False'
    Assert-Equal 'summary の 5 行目に設定ファイル' (@((Read-Result $runDir).Summary -split '\r?\n')[4]) "設定ファイル: $basic"
    $run = Invoke-From $cwd @('-ConfigFile', $basic, '-Destination', "$workRoot\cmd-line-dst", '-OutputFolder', "$workRoot\out\cfg-prio", '-ListOnly')
    Assert-Equal 'コマンドラインが優先' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("コピー先: $workRoot\cmd-line-dst"), (Get-ResultFileNames (Get-RunDir "$workRoot\out\cfg-prio"))) "0/True/$allFiles"

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
        $file = New-ConfigFile "$cfgDir\enc$i.ini" @('Source = ..\..\コピー元 フォルダ', 'Destination = ..\..\enc-dst', "OutputFolder = ..\..\out\cfg-enc$i") $encodings[$label]
        $run  = Invoke-From $cwd @('-ConfigFile', $file, '-ListOnly')
        Assert-Equal ('文字コード {0}: 日本語のパス' -f $label) ('{0}/{1}' -f $run.Code, $run.Log.Contains("コピー元: $src")) '0/True'
        $i++
    }

    # %TEMP% などの環境変数を展開する
    $env:WORKBENCH_TEST_ROOT = $workRoot
    $envIni = New-ConfigFile "$cfgDir\env.ini" @('Source = %WORKBENCH_TEST_ROOT%\コピー元 フォルダ', 'Destination = %WORKBENCH_TEST_ROOT%\env-dst', 'OutputFolder = %WORKBENCH_TEST_ROOT%\out\cfg-env')
    $run = Invoke-From $cwd @('-ConfigFile', $envIni, '-ListOnly')
    Assert-Equal '環境変数' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("コピー先: $workRoot\env-dst"), (Get-ResultFileNames (Get-RunDir "$workRoot\out\cfg-env"))) "0/True/$allFiles"
    # URL エンコードされたフォルダ名（%E3%81%82）は、定義されていない環境変数として止めない
    New-TestFile "$workRoot\cfg\pe\%E3%81%82\f.bin" 10
    $urlIni = New-ConfigFile "$cfgDir\urlenc.ini" @('Source = ..\pe\%E3%81%82', 'Destination = ..\pe-dst', 'OutputFolder = ..\..\out\cfg-url')
    $run = Invoke-From $cwd @('-ConfigFile', $urlIni, '-ListOnly')
    Assert-Equal 'URL エンコードされたフォルダ名' ('{0}/{1}' -f $run.Code, $run.Log.Contains('コピーするファイル 1 件')) '0/True'

    Write-Section '設定ファイル: 書き間違い（調べ始める前に止まる）'
    $mistakes = @(
        @('ListOnly は書けない', @('ListOnly = true'), @('設定ファイルの 1 行目: 知らない項目です: ListOnly', '使える項目: Source, Destination, OutputFolder')),
        @('Force は書けない', @('Source = ..\..\コピー元 フォルダ', 'Force = true'), @('設定ファイルの 2 行目: 知らない項目です: Force')),
        @('ConfigFile は書けない', @('ConfigFile = other.ini'), @('設定ファイルの 1 行目: 知らない項目です: ConfigFile')),
        @('同じ項目を 2 回', @('Source = a', 'Destination = b', 'source = c'), @('設定ファイルの 3 行目: Source が 2 回書かれています（1 行目と 3 行目）')),
        @('= がない行', @('Source C:\data'), @('設定ファイルの 1 行目: 「項目 = 値」の形になっていません: Source C:\data')),
        @('セクションの見出し', @('[Backup-Folder]', 'Source = a'), @('設定ファイルの 1 行目: 「項目 = 値」の形になっていません: [Backup-Folder]')),
        @('引用符が片方だけ', @('Source = "C:\data'), @('設定ファイルの 1 行目: Source の値に引用符（"）が残っています: "C:\data')),
        @('定義されていない環境変数', @('Destination = %WORKBENCH_NO_SUCH_VAR%\x'), @('設定ファイルの 1 行目: 定義されていない環境変数です: %WORKBENCH_NO_SUCH_VAR%'))
    )
    for ($i = 0; $i -lt $mistakes.Count; $i++) {
        $file    = New-ConfigFile "$cfgDir\err$i.ini" $mistakes[$i][1]
        $run     = Invoke-From $cwd @('-ConfigFile', $file)
        $missing = @(@($mistakes[$i][2]) + "ファイル: $file" | Where-Object { -not $run.Log.Contains($_) })
        # 終了コード / 表示に足りない文 / 調べ始めたか
        Assert-Equal ('書き間違い: {0}' -f $mistakes[$i][0]) ('{0}/{1}/{2}' -f $run.Code, ($missing -join ' | '), $run.Log.Contains('コピー元: ')) '1//False'
    }
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\nothing.ini")
    Assert-Equal '設定ファイルが見つからない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("設定ファイルが見つかりません: $cfgDir\nothing.ini")) '1/True'
    $null = New-ConfigFile "$cfgDir\hidden.ini.txt" @('Source = a')
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\hidden.ini")
    Assert-Equal '拡張子が .txt になっている: ヒント' ('{0}/{1}' -f $run.Code, $run.Log.Contains("拡張子が .txt になっていませんか: $cfgDir\hidden.ini.txt")) '1/True'
    # 設定ファイルの値が原因で止まったときは、何行目かを添える
    $noSrc = New-ConfigFile "$cfgDir\nosrc.ini" @('Destination = ..\..\x-dst', 'Source = .\nosuch')
    $run = Invoke-From $cwd @('-ConfigFile', $noSrc)
    Assert-Equal '設定ファイルの Source が見つからない: 行番号' ('{0}/{1}' -f $run.Code, $run.Log.Contains("コピー元のフォルダを開けません: $cfgDir\nosuch（設定ファイルの 2 行目）")) '1/True'
    if ($freeLetter) {
        $noDrive = New-ConfigFile "$cfgDir\nodrive.ini" @('Source = ..\..\コピー元 フォルダ', "Destination = ${freeLetter}:\backup")
        $run = Invoke-From $cwd @('-ConfigFile', $noDrive)
        Assert-Equal '設定ファイルの Destination のドライブがない: 行番号' ('{0}/{1}' -f $run.Code, $run.Log.Contains("コピー先のドライブ ${freeLetter}: が見つかりません。USB メモリなどがつながっているか確かめてください（設定ファイルの 2 行目）")) '1/True'
    }

    Write-Section '起動ファイル（-CreateLauncher）'
    $launchDir  = Join-Path $workRoot '起動 [x]'
    $launchTool = Join-Path $launchDir 'Backup-Folder.ps1'
    $launchCmd  = Join-Path $launchDir 'Backup-Folder.cmd'
    $launchIni  = Join-Path $launchDir 'Backup-Folder.ini'
    [void][IO.Directory]::CreateDirectory($launchDir)
    Copy-Item -LiteralPath $repoTool -Destination $launchTool
    $run = Invoke-Raw -ToolArgs @('-CreateLauncher') -ToolPath $launchTool
    Assert-Equal '作る: 終了コード' $run.Code 0
    Assert-Equal '作る: メッセージ' ('{0}/{1}' -f $run.Log.Contains("起動ファイルを作りました: $launchCmd"), $run.Log.Contains("設定ファイルのひな形を作りました: $launchIni")) 'True/True'
    Assert-Equal '.cmd の中身' (@([IO.File]::ReadAllLines($launchCmd)) -join ' | ') '@echo off | powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dpn0.ps1" -ConfigFile "%~dpn0.ini" | pause'
    Assert-Equal '.cmd は ASCII だけ（cmd.exe がどのコードページで読んでも壊れない）' @([IO.File]::ReadAllBytes($launchCmd) | Where-Object { $_ -gt 0x7F }).Count 0
    Assert-Equal '.ini は BOM 付き UTF-8' (Test-Bom $launchIni) 'True'
    Assert-Equal '.ini の項目は空' (@([IO.File]::ReadAllLines($launchIni) | Where-Object { $_ -match '^\w+ =$' }) -join ',') 'Source =,Destination =,OutputFolder ='
    # ひな形のままダブルクリックすると、コピー元を書くよう案内して止まる（pause への Enter を送る）
    $log = ConvertTo-FlatText ((@("`r`n" | & cmd.exe /d /c $launchCmd 2>&1) | ForEach-Object { "$_" }) -join "`n")
    Assert-Equal 'ひな形のまま起動' $log.Contains('コピー元を指定してください') 'True'
    $null = New-ConfigFile $launchIni @('Source = ..\コピー元 フォルダ', 'Destination = ..\launch-dst', 'OutputFolder = .\記録') (New-Object Text.UTF8Encoding $true)
    $before = [IO.File]::ReadAllText($launchIni)
    $run = Invoke-Raw -ToolArgs @('-CreateLauncher') -ToolPath $launchTool
    Assert-Equal 'もう一度作る: 上書きしない' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains("起動ファイルはすでにあるため、作りませんでした: $launchCmd"), $run.Log.Contains("設定ファイルのひな形はすでにあるため、作りませんでした: $launchIni"), ([IO.File]::ReadAllText($launchIni) -eq $before)) '0/True/True/True'
    # ダブルクリックと同じく、cmd.exe から .cmd を起動する（確認への Y と、pause への Enter を送る）。.cmd は 5.1 で動く
    $log = ConvertTo-FlatText ((@("Y`r`n`r`n" | & cmd.exe /d /c $launchCmd 2>&1) | ForEach-Object { "$_" }) -join "`n")
    Assert-Equal '.cmd で起動: 結果の行' $log.Contains('結果: 正常に終わりました') 'True'
    Assert-Equal '.cmd で起動: コピー先と記録' ('{0}/{1}' -f [IO.File]::Exists("$workRoot\launch-dst\a.txt"), (Get-ResultFileNames (Get-RunDir "$launchDir\記録"))) "True/$allFiles"
    # -OutputFolder も設定ファイルもないときは、スクリプトと同じフォルダに保存する
    $run = Invoke-Raw -ToolArgs @($src, "$workRoot\launch-dst2", '-ListOnly') -ToolPath $launchTool
    Assert-Equal '保存先の既定はスクリプトのフォルダ' ('{0}/{1}' -f $run.Code, (Get-ResultFileNames (Get-RunDir $launchDir))) "0/$allFiles"

    Write-Section 'コマンドプロンプトから起動'
    $cmdSrc = Join-Path $workRoot 'cmdsrc'
    New-TestFile "$cmdSrc\c.txt" 1
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $run = Invoke-ViaCmd ('{0} -Source "{1}" -Destination "{2}" -ListOnly -OutputFolder "{3}"' -f $cmdBase, $cmdSrc, "$workRoot\cmd-dst", "$workRoot\out\cmd")
    Assert-Equal '引用符で囲んだパス' ('{0}/{1}' -f $run.Code, $run.Log.Contains("コピー元: $cmdSrc")) '0/True'
    # 最後が \" だと、引用符が閉じずに後ろの引数まで巻き込む。空かどうかより先に、引用符の問題として知らせる
    $hint = '最後に \ を付けないでください'
    $run = Invoke-ViaCmd ('{0} -Source "{1}\" -Destination "{2}"' -f $cmdBase, $cmdSrc, "$workRoot\cmd-dst")
    Assert-Equal '-Source "…\"' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('-Source の値に引用符（"）が残っています'), $run.Log.Contains($hint)) '1/True/True'
    $run = Invoke-ViaCmd ('{0} -Source "{1}" -Destination "{2}\" -ListOnly' -f $cmdBase, $cmdSrc, "$workRoot\cmd-dst2")
    Assert-Equal '-Destination "…\"（後ろの -ListOnly ごと巻き込む）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('-Destination の値に引用符（"）が残っています'), (Test-Path -LiteralPath "$workRoot\cmd-dst2")) '1/True/False'
    $run = Invoke-ViaCmd ('{0} -OutputFolder "{1}\" -ListOnly -Source "{2}" -Destination "{3}"' -f $cmdBase, "$workRoot\out\cmd-q", $cmdSrc, "$workRoot\cmd-dst")
    Assert-Equal '-OutputFolder "…\"' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-OutputFolder の値に引用符（"）が残っています')) '1/True'
    $run = Invoke-ViaCmd ('{0} -ConfigFile "{1}\" -ListOnly' -f $cmdBase, "$workRoot\cfg")
    Assert-Equal '-ConfigFile "…\"' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-ConfigFile の値に引用符（"）が残っています')) '1/True'

    Write-Section '一時的な記録'
    Assert-Equal 'ツールが一時的な記録を残さない（すべての実行の後）' (Get-TempLogCount) 0
} finally {
    # アクセス拒否と subst を戻してから消す
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    if ($substDrive) { $null = & subst.exe $substDrive /D 2>&1 }
    $env:TMP  = $savedTmp
    $env:TEMP = $savedTemp
    Remove-Item Env:\WORKBENCH_TEST_ROOT -ErrorAction SilentlyContinue
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            # ループするジャンクションは、リンク先をたどらないように先にリンクだけ消す。読み取り専用のファイルは消せないため戻す
            if (Test-Path -LiteralPath "$src\jn") { [IO.Directory]::Delete("$src\jn") }
            foreach ($file in $readOnly) { if ([IO.File]::Exists($file)) { [IO.File]::SetAttributes($file, 'Normal') } }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
    if ($savedInputEncoding) { try { [Console]::InputEncoding = $savedInputEncoding } catch { } }
}

exit $script:failCount
