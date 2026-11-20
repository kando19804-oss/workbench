<#
.SYNOPSIS
    Remove-OldFile の回帰テストです。

.DESCRIPTION
    一時フォルダにダミーのファイル（古い・新しい・コピーして置いた・未来の日時・読み取り専用・システム・隠し、名前に [ ] や
    日本語・空白を含むもの、260 文字を超えるパスのもの、このツールの結果のフォルダ、ごみ箱、ジャンクション、読めないフォルダ）を作り、
    Remove-OldFile の一覧・削除・移動の結果を期待値と照合します。削除・移動したかは、ファイルが残っているかで確かめます。

    「削除の拒否」は、icacls.exe で付けると、ウイルス対策ソフトなどがプロセスごと止めることがあるため（2026-11-18、開発の PC）、
    .NET の ACL の API で付けます。テストが作ったフォルダとファイルにだけ設定し、終了時に戻します。
    別のドライブへの移動は、この PC のループバックの管理共有（\\127.0.0.1\C$）を移す先にして確かめます（外へは通信しません）。
    SUBST のドライブの確認では、空いているドライブ文字を一時的に割り当て、終了時に外します。

    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Remove-OldFile を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Remove-OldFile\Remove-OldFile.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Remove-OldFile_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Remove-OldFile.ps1'
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
$isAdmin   = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$filesHead   = '"ファイル","サイズ(バイト)","作成日時","最終更新日時","経過日数","結果","備考","移した先"'
$foldersHead = '"フォルダ","件数","サイズ(バイト)","いちばん古い日時","いちばん新しい日時"'
$emptyHead   = '"フォルダ","結果","備考"'
$script:failCount = 0

$now     = Get-Date
$old     = $now.AddDays(-200)
$older   = $now.AddDays(-300)
$new     = $now.AddDays(-1)
$future  = $now.AddDays(30)
$longRel = 'longdir\' + ('L' * 120) + '\' + ('M' * 120) + '\old.log'
$longDirs = @('longdir', ('longdir\' + ('L' * 120)), ('longdir\' + ('L' * 120) + '\' + ('M' * 120)))

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

function Join-Sorted($Items) {
    $array = [string[]]@($Items | Where-Object { $null -ne $_ })
    [Array]::Sort($array, [StringComparer]::Ordinal)
    $array -join ','
}

# ローカルのパスを、指定したサーバーの管理共有のパスにする（C:\a → \\127.0.0.1\C$\a）
function ConvertTo-AdminShare([string]$Server, [string]$LocalPath) {
    '\\' + $Server + '\' + $LocalPath.Substring(0, 1).ToUpperInvariant() + '$\' + $LocalPath.Substring(3)
}

# 日時と属性を決めたファイルを作る（260 文字を超えるパスにも書けるように、\\?\ を付ける）
function New-AgedFile([string]$FilePath, [int]$Size, [datetime]$Created, [datetime]$Written, [string]$Attr = '', [int]$Fill = 0) {
    $long = '\\?\' + $FilePath
    [void][IO.Directory]::CreateDirectory($long.Substring(0, $long.LastIndexOf('\')))
    $bytes = New-Object 'byte[]' $Size
    if ($Fill) { for ($i = 0; $i -lt $Size; $i++) { $bytes[$i] = [byte]$Fill } }
    [IO.File]::WriteAllBytes($long, $bytes)
    [IO.File]::SetCreationTime($long, $Created)
    [IO.File]::SetLastWriteTime($long, $Written)
    if ($Attr) { [IO.File]::SetAttributes($long, [IO.FileAttributes]$Attr) }
}

# インターネットから取得したファイルの印（代替データストリーム）を付ける。書くと最終更新日時が変わるので、元に戻す
function Set-ZoneMark([string]$FilePath) {
    $created = [IO.File]::GetCreationTime($FilePath)
    $written = [IO.File]::GetLastWriteTime($FilePath)
    Set-Content -LiteralPath $FilePath -Stream 'Zone.Identifier' -Value "[ZoneTransfer]`r`nZoneId=3"
    [IO.File]::SetCreationTime($FilePath, $created)
    [IO.File]::SetLastWriteTime($FilePath, $written)
}

function Test-ZoneMark([string]$FilePath) {
    @(Get-Item -LiteralPath $FilePath -Stream * -ErrorAction SilentlyContinue | Where-Object { $_.Stream -eq 'Zone.Identifier' }).Count -gt 0
}

# フォルダの中身の一覧（相対パス。フォルダは末尾に \、リンクは末尾に @。リンクの先と読めないフォルダの中はたどらない）
function Get-Tree([string]$Dir) {
    $result = New-Object 'System.Collections.Generic.List[string]'
    $stack = New-Object 'System.Collections.Generic.Stack[string]'
    $stack.Push($Dir)
    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try { $entries = (New-Object IO.DirectoryInfo ('\\?\' + $current)).GetFileSystemInfos() } catch { continue }
        foreach ($entry in $entries) {
            $full = $entry.FullName.Substring(4)
            $rel = $full.Substring($Dir.Length + 1)
            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { $result.Add($rel + '@'); continue }
            if ($entry -is [IO.DirectoryInfo]) { $result.Add($rel + '\'); $stack.Push($full); continue }
            $result.Add($rel)
        }
    }
    Join-Sorted $result
}

function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$InputText = '', [string]$ToolPath = $tool) {
    $out = Join-Path $workRoot "out\$Label"
    $lines = @($InputText | & $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = (ConvertTo-FlatText (@($lines) -join "`n")); Lines = @($lines); Dir = (Get-RunDir $out) }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行。確認を待つときに止まらないよう、標準入力は空にする
function Invoke-Plain([string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $log = ConvertTo-FlatText ((@('' | & $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 相対パスの基準を確かめるため、指定したフォルダから起動する
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

# ツールの処理を 1 か所だけ書き換えたコピーを作る（本物の環境では起きない条件を確かめるため）
function New-PatchedTool([string]$Name, [string]$From, [string]$To) {
    $text = [IO.File]::ReadAllText($repoTool)
    if (-not $text.Contains($From)) { throw "書き換える場所が見つかりません: $From" }
    $file = Join-Path $workRoot "patched\$Name\Remove-OldFile.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text.Replace($From, $To), $utf8Bom)
    $file
}

# 保存先のフォルダの下に作られた、結果のフォルダ（260 文字を超えても探せるように \\?\ を付ける）
function Get-RunDir([string]$Parent) {
    if (-not [IO.Directory]::Exists('\\?\' + $Parent)) { return $null }
    $found = @([IO.Directory]::GetDirectories('\\?\' + $Parent, 'OldFile_*') | Sort-Object)
    if ($found.Count -gt 0) { return $found[-1].Substring(4) }
    $null
}

# 結果のファイルがないとき（ツールが保存に失敗したときなど）も止まらないように、空として読む
function Read-Result([string]$Dir) {
    function Read-CsvOrEmpty([string]$Name) {
        if (-not $Dir) { return @() }
        $file = '\\?\' + (Join-Path $Dir $Name)
        if (-not [IO.File]::Exists($file)) { return @() }
        @([IO.File]::ReadAllLines($file, $utf8NoBom) | ConvertFrom-Csv)
    }
    $lines = @()
    if ($Dir -and [IO.File]::Exists('\\?\' + (Join-Path $Dir 'summary.txt'))) { $lines = @([IO.File]::ReadAllLines('\\?\' + (Join-Path $Dir 'summary.txt'))) }
    [pscustomobject]@{
        Files = @(Read-CsvOrEmpty 'files.csv'); Folders = @(Read-CsvOrEmpty 'folders.csv'); Empty = @(Read-CsvOrEmpty 'emptyfolders.csv')
        Skipped = @(Read-CsvOrEmpty 'skipped.csv'); Lines = $lines; Summary = ($lines -join "`n")
    }
}

function Get-ResultNames([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted ([IO.Directory]::GetFileSystemEntries('\\?\' + $Dir) | ForEach-Object { Split-Path -Leaf $_ })
}

function Get-BomFlags([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted ([IO.Directory]::GetFiles('\\?\' + $Dir) | ForEach-Object {
        $bytes = [IO.File]::ReadAllBytes($_)
        '{0}={1}' -f (Split-Path -Leaf $_), ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    })
}

function Get-FirstLine([string]$Dir, [string]$FileName) {
    if (-not $Dir -or -not [IO.File]::Exists('\\?\' + (Join-Path $Dir $FileName))) { return '(なし)' }
    @([IO.File]::ReadAllLines('\\?\' + (Join-Path $Dir $FileName)))[0]
}

# files.csv（emptyfolders.csv）の、ダミーのフォルダから見た相対パスと列の値の一覧
function Get-RowList($Rows, [string]$Base, [string]$PathColumn, [string]$Column) {
    Join-Sorted ($Rows | ForEach-Object { '{0}={1}' -f $_.$PathColumn.Substring($Base.Length + 1), $_.$Column })
}

function Get-Row($Rows, [string]$Base, [string]$Rel, [string]$PathColumn, [string]$Column) {
    $full = Join-Path $Base $Rel
    (@($Rows | Where-Object { $_.$PathColumn -eq $full }) | ForEach-Object { $_.$Column }) -join '|'
}

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $used -notcontains $_ -and -not [IO.Directory]::Exists("${_}:\") } | Select-Object -First 1
}

# 一覧の拒否（RD）は icacls で付ける。終了時に戻すため、設定したパスを覚えておく
function Add-Deny([string]$Path, [string]$Rights) {
    $null = & icacls.exe $Path /deny "*${userSid}:$Rights" 2>&1
    if ($LASTEXITCODE -ne 0) { return $false }
    $denied.Add($Path)
    $true
}

# 削除の拒否は、icacls.exe で付けるとウイルス対策ソフトなどにプロセスごと止められることがあるため、.NET の ACL の API で付ける。
# 本人の「削除」と、親のフォルダの「子の削除」の両方を拒否する（片方だけだと、もう片方の権限で削除できる）
function Get-ItemAcl([string]$Path) {
    if ([IO.Directory]::Exists($Path)) { return (New-Object IO.DirectoryInfo $Path).GetAccessControl() }
    return (New-Object IO.FileInfo $Path).GetAccessControl()
}

function Set-ItemAcl([string]$Path, $Acl) {
    if ([IO.Directory]::Exists($Path)) { (New-Object IO.DirectoryInfo $Path).SetAccessControl($Acl) }
    else { (New-Object IO.FileInfo $Path).SetAccessControl($Acl) }
}

function Add-DenyDelete([string]$Path) {
    try {
        $sid = New-Object Security.Principal.SecurityIdentifier $userSid
        $deny = [Security.AccessControl.AccessControlType]::Deny
        $rule = New-Object Security.AccessControl.FileSystemAccessRule($sid, [Security.AccessControl.FileSystemRights]::Delete, $deny)
        $parentRule = New-Object Security.AccessControl.FileSystemAccessRule($sid, [Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles, $deny)
        $parent = Split-Path -Parent $Path
        $acl = Get-ItemAcl $Path; $acl.AddAccessRule($rule); Set-ItemAcl $Path $acl
        $aclDenies.Add([pscustomobject]@{ Path = $Path; Rule = $rule })
        $acl = Get-ItemAcl $parent; $acl.AddAccessRule($parentRule); Set-ItemAcl $parent $acl
        $aclDenies.Add([pscustomobject]@{ Path = $parent; Rule = $parentRule })
        return $true
    } catch {
        return $false
    }
}

# 主な場面のダミー。古い（作成・最終更新とも 200 日前）、新しい、コピーして置いた（作成だけ新しい）、未来の日時、
# 読み取り専用・システム（対象外）・隠し（対象）、空になるフォルダ、名前に [ ] や日本語・空白を含むもの、260 文字を超えるパス、
# このツールの結果のフォルダとごみ箱（除外）、ジャンクション（たどらない）、読めないフォルダ（スキップ）
function Build-Tree([string]$Dir) {
    New-AgedFile "$Dir\a.log" 5000 $old $old '' 65
    New-AgedFile "$Dir\b.log" 1000 $old $new
    New-AgedFile "$Dir\copied.log" 1000 $new $old
    New-AgedFile "$Dir\future.log" 1000 $future $future
    New-AgedFile "$Dir\ro.log" 1000 $old $old 'ReadOnly'
    New-AgedFile "$Dir\sys.log" 1000 $old $old 'Hidden, System'
    New-AgedFile "$Dir\hidden.log" 2000 $old $old 'Hidden'
    New-AgedFile "$Dir\note.txt" 3000 $old $old
    New-AgedFile "$Dir\sub1\x.log" 4000 $old $old
    New-AgedFile "$Dir\sub1\y.txt" 1000 $old $old
    New-AgedFile "$Dir\sub2\z.log" 1000 $old $old
    New-AgedFile "$Dir\sub2\keep.txt" 1000 $new $new
    [void][IO.Directory]::CreateDirectory("$Dir\sub3")
    New-AgedFile "$Dir\sub4\deep\w.log" 7000 $old $older
    New-AgedFile "$Dir\[br]\v.log" 1000 $old $old
    New-AgedFile "$Dir\日本語 フォルダ\資料 1.log" 1500 $old $old
    New-AgedFile "$Dir\OldFile_memo\m.log" 1000 $old $old
    New-AgedFile "$Dir\OldFile_PC01_20260101_000000\r.log" 1000 $old $old
    New-AgedFile "$Dir\`$RECYCLE.BIN\r.log" 1000 $old $old
    New-AgedFile "$Dir\$longRel" 6000 $old $old
    New-AgedFile "$Dir\noread\n.log" 1000 $old $old
    $null = & cmd.exe /d /c mklink /J "$Dir\jlink" "$workRoot\outside" 2>&1
    $junctions.Add("$Dir\jlink")
    (Add-Deny "$Dir\noread" '(RD)') -and [IO.Directory]::Exists("$Dir\jlink")
}

# 主な場面の、古いファイルと結果（一覧だけのとき）
$listResults = @(
    'a.log=', 'hidden.log=', "$longRel=", 'note.txt=', 'OldFile_memo\m.log=', 'ro.log=対象外', 'sub1\x.log=', 'sub1\y.txt=', 'sub2\z.log=',
    'sub4\deep\w.log=', 'sys.log=対象外', '[br]\v.log=', '日本語 フォルダ\資料 1.log='
)
$targetNames = @('a.log', 'hidden.log', $longRel, 'note.txt', 'OldFile_memo\m.log', 'sub1\x.log', 'sub1\y.txt', 'sub2\z.log', 'sub4\deep\w.log', '[br]\v.log', '日本語 フォルダ\資料 1.log')
$emptyNames = @('OldFile_memo', 'sub1', 'sub4', 'sub4\deep', '[br]', '日本語 フォルダ') + $longDirs

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
# 標準入力に渡す文字は、コンソールの入力の文字コードで書かれる。UTF-8 のとき先頭に BOM が付き、7 の Read-Host は
# それを外さないため、BOM を付けない UTF-8 にする（外部のプログラムに渡すときはグローバルの $OutputEncoding が使われる）
$savedInputEncoding = $null
try { $savedInputEncoding = [Console]::InputEncoding; [Console]::InputEncoding = $utf8NoBom } catch { }
$savedOutputEncoding = $global:OutputEncoding
$global:OutputEncoding = $utf8NoBom
$userSid   = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied    = New-Object 'System.Collections.Generic.List[string]'   # icacls で拒否を設定したパス（終了時に戻す）
$aclDenies = New-Object 'System.Collections.Generic.List[object]'   # .NET の API で付けた拒否（終了時に戻す）
$junctions = New-Object 'System.Collections.Generic.List[string]'   # 作ったジャンクション（終了時に先に外す）
$holders   = New-Object 'System.Collections.Generic.List[object]'   # 開いたままにしたファイルとプロセス（終了時に閉じる）
$substLetter = $null
$main = Join-Path $workRoot 'main'

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool
    New-AgedFile "$workRoot\outside\o.log" 1000 $old $old
    $mainOk = Build-Tree $main
    Assert-Equal 'ジャンクションとアクセス拒否を作れた' $mainOk 'True'
    $O = ConvertTo-AdminShare '127.0.0.1' $workRoot
    $hasShare = [IO.Directory]::Exists($O)
    if (-not $hasShare) { Write-Skip '管理共有' "$O を開けないため、別のドライブへの移動のテストを飛ばします" }
    $mainTree = Get-Tree $main

    Write-Section '一覧だけ（既定）'
    $run = Invoke-Tool 'list' @('-Path', $main, '-OlderThan', '90')
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '保存するファイル（emptyfolders.csv は -RemoveEmptyFolder のときだけ）' (Get-ResultNames $run.Dir) 'files.csv,folders.csv,skipped.csv,summary.txt'
    Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'files.csv=True,folders.csv=True,skipped.csv=True,summary.txt=True'
    Assert-Equal 'files.csv の見出し' (Get-FirstLine $run.Dir 'files.csv') $filesHead
    Assert-Equal 'folders.csv の見出し' (Get-FirstLine $run.Dir 'folders.csv') $foldersHead
    Assert-Equal 'skipped.csv の見出し' (Get-FirstLine $run.Dir 'skipped.csv') '"理由","パス","詳細"'
    Assert-Equal '古いファイルと結果（新しい・コピーして置いた・未来の日時は出ない。読み取り専用とシステムは対象外。隠しは対象）' (Get-RowList $x.Files $main 'ファイル' '結果') (Join-Sorted $listResults)
    $order = @($x.Files | ForEach-Object { $_.'ファイル'.Substring($main.Length + 1) })
    $expectedOrder = 'a.log', 'hidden.log', $longRel, 'note.txt', 'OldFile_memo\m.log', 'ro.log', 'sub1\x.log', 'sub1\y.txt', 'sub2\z.log', 'sub4\deep\w.log', 'sys.log', '[br]\v.log', '日本語 フォルダ\資料 1.log'
    Assert-Equal 'files.csv の並び（パスの順。大文字にそろえた文字の番号の順）' ($order -join ',') ($expectedOrder -join ',')
    Assert-Equal '対象外の備考' ('{0}|{1}' -f (Get-Row $x.Files $main 'ro.log' 'ファイル' '備考'), (Get-Row $x.Files $main 'sys.log' 'ファイル' '備考')) '読み取り専用のファイルです|システムファイルです'
    Assert-Equal 'サイズは 3 桁区切り・経過日数は新しいほうの日時から（作成のほうが新しい w.log も 200）' ('{0}|{1}|{2}' -f (Get-Row $x.Files $main 'a.log' 'ファイル' 'サイズ(バイト)'), (Get-Row $x.Files $main 'a.log' 'ファイル' '経過日数'), (Get-Row $x.Files $main 'sub4\deep\w.log' 'ファイル' '経過日数')) '5,000|200|200'
    Assert-Equal '日時の形（yyyy/MM/dd HH:mm:ss）と、作成日時・最終更新日時を別々に出す' ('{0}/{1}' -f ((Get-Row $x.Files $main 'a.log' 'ファイル' '作成日時') -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$'), ((Get-Row $x.Files $main 'sub4\deep\w.log' 'ファイル' '作成日時') -ne (Get-Row $x.Files $main 'sub4\deep\w.log' 'ファイル' '最終更新日時'))) 'True/True'
    Assert-Equal '一覧だけのときは、移した先が空' (@($x.Files | Where-Object { $_.'移した先' }).Count) 0
    $folderList = @($x.Folders | ForEach-Object {
        $rel = if ($_.'フォルダ' -eq $main) { '.' } else { $_.'フォルダ'.Substring($main.Length + 1) }
        '{0}|{1}|{2}' -f $rel, $_.'件数', $_.'サイズ(バイト)'
    })
    $expectedFolders = '.|5|12,000', 'sub4\deep|1|7,000', ('longdir\' + ('L' * 120) + '\' + ('M' * 120) + '|1|6,000'), 'sub1|2|5,000', '日本語 フォルダ|1|1,500', 'OldFile_memo|1|1,000', 'sub2|1|1,000', '[br]|1|1,000'
    Assert-Equal 'folders.csv（直下の古いファイルの合計。サイズの大きい順、同じならパスの順）' ($folderList -join ',') ($expectedFolders -join ',')
    $skipText = Join-Sorted ($x.Skipped | ForEach-Object { '{0}:{1}:{2}' -f $_.'理由', $_.'パス'.Substring($main.Length + 1), $_.'詳細' })
    Assert-Equal 'skipped.csv（読めないフォルダ・結果のフォルダとごみ箱・ジャンクション）' $skipText (Join-Sorted @('権限なし:noread:アクセスが拒否されました', '除外指定:OldFile_PC01_20260101_000000:このツールの結果のフォルダ', '除外指定:$RECYCLE.BIN:ごみ箱・システムの管理用のフォルダ', 'リンク:jlink:ジャンクション'))
    $summary = $x.Summary
    Assert-Equal 'summary の 1 行目' (@($x.Lines) + '')[0] '古いファイルの整理 (Remove-OldFile 1.0.0)'
    Assert-Equal '対象と日時の行' ('{0}|{1}' -f (@($x.Lines) + '' + '')[1], [regex]::IsMatch((@($x.Lines) + '' + '' + '')[2], ('^\d{{4}}-\d{{2}}-\d{{2}} \d{{2}}:\d{{2}}:\d{{2}}   {0}（管理者: (はい|いいえ)）$' -f [regex]::Escape($env:COMPUTERNAME)))) "対象: $main|True"
    Assert-Equal '条件の行' ([regex]::IsMatch($summary, '\n条件: 作成・最終更新とも 90 日以上前（\d{4}-\d{2}-\d{2} \d{2}:\d{2} より前）のファイル\n')) 'True'
    Assert-Equal '件数とスキップの行' ([regex]::IsMatch($summary, '\nファイル 17 件  古いファイル 13 件（33\.7 KB）  所要 \d+\.\d秒\nスキップ 1 件（権限なし）  たどらなかったリンク 1 件  除外 2 件\n')) 'True'
    Assert-Equal 'フォルダの表の見出しと 1 行目' ([regex]::IsMatch($summary, '■ 古いファイルのあるフォルダ（サイズの大きい順）\n      サイズ      件数  いちばん古い  フォルダ\n     11\.7 KB         5  \d{4}-\d{2}-\d{2}    ' + [regex]::Escape($main) + '\n')) 'True'
    Assert-Equal '対象の行と案内' ([regex]::IsMatch($summary, '\n対象 11 件（31\.7 KB）  対象外 2 件（読み取り専用 1・システムファイル 1）\n(※ 管理者として実行すると、対象にできる範囲が広がります\n)?削除するときは -Delete を、移すときは -MoveTo を付けて、もう一度実行してください\n')) 'True'
    Assert-Equal '管理者でないときは、管理者の案内を出す' ($summary.Contains('※ 管理者として実行すると、対象にできる範囲が広がります')) (-not $isAdmin)
    Assert-Equal 'summary の最後は保存先' (@($x.Lines) + '')[-2] ('保存先: {0}' -f $run.Dir)
    Assert-Equal '一覧だけでは、ファイルを変えない' (Get-Tree $main) $mainTree

    Write-Section '一覧だけ（-RemoveEmptyFolder）'
    $run = Invoke-Tool 'list-empty' @('-Path', $main, '-OlderThan', '90', '-RemoveEmptyFolder')
    $x = Read-Result $run.Dir
    Assert-Equal '終了コードと保存するファイル' ('{0}/{1}' -f $run.Code, (Get-ResultNames $run.Dir)) '0/emptyfolders.csv,files.csv,folders.csv,skipped.csv,summary.txt'
    Assert-Equal 'emptyfolders.csv の見出し' (Get-FirstLine $run.Dir 'emptyfolders.csv') $emptyHead
    Assert-Equal '空になる予定のフォルダ（もともと空・新しいファイルが残る・読めない・-Path そのものは入らない）' (Get-RowList $x.Empty $main 'フォルダ' '結果') (Join-Sorted ($emptyNames | ForEach-Object { "$_=" }))
    Assert-Equal '空になる予定のフォルダの行' ($x.Summary.Contains("`n空になる予定のフォルダ 9 件`n")) 'True'
    Assert-Equal '一覧だけでは、ファイルもフォルダも変えない' (Get-Tree $main) $mainTree

    Write-Section '絞り込み'
    $run = Invoke-Plain @('-Path', $main, '-OlderThan', '90', '-Include', '*.log', '-Exclude', 'sub1,*.txt', '-NoSave')
    Assert-Equal '-Include と -Exclude（除外したものも除外に数える）' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('ファイル 13 件  古いファイル 10 件'), $run.Log.Contains('除外: sub1, *.txt'), $run.Log.Contains('除外 5 件')) '0/True/True/True'
    Assert-Equal '-Include の条件の行' ($run.Log.Contains('のファイル  名前: *.log')) 'True'
    $run = Invoke-Plain @('-Path', "$main\sub1,$main\sub2", '-OlderThan', '90', '-NoSave')
    Assert-Equal '-Path をカンマで複数' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("対象: $main\sub1, $main\sub2"), $run.Log.Contains('ファイル 4 件  古いファイル 3 件')) '0/True/True'
    $run = Invoke-Plain @('-Path', $main, '-OlderThan', '９０', '-NoSave')
    Assert-Equal '全角の数字の -OlderThan' ('{0}/{1}' -f $run.Code, $run.Log.Contains('古いファイル 13 件')) '0/True'
    $run = Invoke-Plain @('-Path', $main, '-OlderThan', '36500', '-NoSave')
    Assert-Equal '古いファイルがないとき' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('（古いファイルはありません）'), $run.Log.Contains('対象のファイルはありません')) '0/True/True'

    Write-Section '削除（-Delete -Force -RemoveEmptyFolder）'
    $del = Join-Path $workRoot 'del'
    $null = Build-Tree $del
    $run = Invoke-Tool 'delete' @('-Path', $del, '-OlderThan', '90', '-Delete', '-Force', '-RemoveEmptyFolder')
    $x = Read-Result $run.Dir
    Assert-Equal '終了コードと保存するファイル' ('{0}/{1}' -f $run.Code, (Get-ResultNames $run.Dir)) '0/emptyfolders.csv,files.csv,folders.csv,skipped.csv,summary.txt'
    Assert-Equal '結果（対象は削除した・対象外はそのまま）' (Get-RowList $x.Files $del 'ファイル' '結果') (Join-Sorted ($listResults | ForEach-Object { if ($_ -like '*=') { $_ + '削除した' } else { $_ } }))
    $expectedLeft = Join-Sorted @('$RECYCLE.BIN\', '$RECYCLE.BIN\r.log', 'OldFile_PC01_20260101_000000\', 'OldFile_PC01_20260101_000000\r.log', 'b.log', 'copied.log', 'future.log', 'jlink@', 'noread\', 'ro.log', 'sub2\', 'sub2\keep.txt', 'sub3\', 'sys.log')
    Assert-Equal '残ったもの（空になったフォルダも削除。もともと空のフォルダ・リンク・除外したもの・読めないものは残る）' (Get-Tree $del) $expectedLeft
    Assert-Equal 'ジャンクションの先は消さない' ([IO.File]::Exists("$workRoot\outside\o.log")) 'True'
    Assert-Equal '空になったフォルダの結果' (Get-RowList $x.Empty $del 'フォルダ' '結果') (Join-Sorted ($emptyNames | ForEach-Object { "$_=削除した" }))
    $summary = $x.Summary
    Assert-Equal '予定の行' ([regex]::IsMatch($summary, '\n削除するファイル 11 件（31\.7 KB）\n削除しないファイル 2 件（読み取り専用 1・システムファイル 1）\n(※ 管理者として実行すると、対象にできる範囲が広がります\n)?空になるフォルダ 9 件も削除します\n※ 削除したファイルは元に戻せません（ごみ箱には入りません）\n-Force を指定したため、確認していません\n')) 'True'
    Assert-Equal '結果の行' ([regex]::IsMatch($summary, '\n結果: 削除しました  11 件（31\.7 KB）  失敗 0 件  所要 \d+\.\d秒\n空になったフォルダ 9 件を削除しました\n')) 'True'

    Write-Section '確認（Y と N）'
    $yn = Join-Path $workRoot 'yn'
    $null = Build-Tree $yn
    $ynTree = Get-Tree $yn
    $run = Invoke-Tool 'answer-n' @('-Path', $yn, '-OlderThan', '90', '-Delete', '-RemoveEmptyFolder') 'N'
    $x = Read-Result $run.Dir
    Assert-Equal 'N: 終了コード 0・ファイルを変えない' ('{0}/{1}' -f $run.Code, ((Get-Tree $yn) -eq $ynTree)) '0/True'
    Assert-Equal 'N: 結果は「削除しなかった」と「対象外」・空のフォルダも「削除しなかった」' ('{0}/{1}' -f (Join-Sorted ($x.Files | ForEach-Object { $_.'結果' } | Select-Object -Unique)), (Join-Sorted ($x.Empty | ForEach-Object { $_.'結果' } | Select-Object -Unique))) '削除しなかった,対象外/削除しなかった'
    Assert-Equal 'N: 確認の答えと「削除しませんでした」を残す' ('{0}/{1}' -f $x.Summary.Contains("`n削除しますか？ [Y] はい  [N] いいえ: N`n削除しませんでした`n"), $run.Log.Contains('削除しませんでした')) 'True/True'
    $run = Invoke-Tool 'answer-y' @('-Path', $yn, '-OlderThan', '90', '-Delete') 'y'
    $x = Read-Result $run.Dir
    Assert-Equal 'Y（小文字）: 削除する' ('{0}/{1}/{2}' -f $run.Code, @($x.Files | Where-Object { $_.'結果' -eq '削除した' }).Count, [IO.File]::Exists("$yn\a.log")) '0/11/False'
    Assert-Equal 'Y: -RemoveEmptyFolder なしでは、空になったフォルダを残す' ([IO.Directory]::Exists("$yn\sub1")) 'True'

    Write-Section '移す（同じドライブ）'
    $mv1 = Join-Path $workRoot 'mv1'
    $dest1 = Join-Path $workRoot 'mvdest1'
    $null = Build-Tree $mv1
    Set-ZoneMark "$mv1\a.log"
    New-AgedFile "$dest1\sub1\x.log" 10 $new $new
    $run = Invoke-Tool 'move-same' @('-Path', $mv1, '-OlderThan', '90', '-MoveTo', $dest1, '-Force', '-RemoveEmptyFolder')
    $x = Read-Result $run.Dir
    $expectedMove = $listResults | ForEach-Object { if ($_ -eq 'sub1\x.log=') { $_ + '移せない' } elseif ($_ -like '*=') { $_ + '移した' } else { $_ } }
    Assert-Equal '終了コード 1（移せないファイルがある）と結果' ('{0}/{1}' -f $run.Code, (Get-RowList $x.Files $mv1 'ファイル' '結果')) ('1/' + (Join-Sorted $expectedMove))
    Assert-Equal '移す先に同じ名前があるときは上書きしない' ('{0}|{1}' -f (Get-Row $x.Files $mv1 'sub1\x.log' 'ファイル' '備考'), ([IO.FileInfo]"$dest1\sub1\x.log").Length) ('移す先に同じ名前のファイルがあります: {0}\sub1\x.log|10' -f $dest1)
    Assert-Equal '移した先の列' ('{0}|{1}' -f (Get-Row $x.Files $mv1 'a.log' 'ファイル' '移した先'), (Get-Row $x.Files $mv1 'ro.log' 'ファイル' '移した先')) "$dest1\a.log|"
    Assert-Equal '同じ構成で移る（260 文字を超えるパスも）' ('{0}/{1}/{2}' -f [IO.File]::Exists("$dest1\sub4\deep\w.log"), [IO.File]::Exists('\\?\' + "$dest1\$longRel"), [IO.File]::Exists("$mv1\a.log")) 'True/True/False'
    $w = New-Object IO.FileInfo "$dest1\sub4\deep\w.log"
    # 名前を変えると、Windows がアーカイブの属性を付ける（エクスプローラーで移したときと同じ）ので、隠しの属性だけを見る
    Assert-Equal '日時・隠しの属性・代替データストリームは元のまま' ('{0}/{1}/{2}/{3}' -f ($w.CreationTime -eq $old), ($w.LastWriteTime -eq $older), (([IO.File]::GetAttributes("$dest1\hidden.log") -band [IO.FileAttributes]::Hidden) -ne 0), (Test-ZoneMark "$dest1\a.log")) 'True/True/True/True'
    Assert-Equal '空になったフォルダは削除し、移せないファイルの残るフォルダは残す' ('{0}/{1}' -f [IO.Directory]::Exists("$mv1\sub4"), [IO.Directory]::Exists("$mv1\sub1")) 'False/True'
    $summary = $x.Summary
    Assert-Equal '移す先・予定・結果の行' ('{0}/{1}/{2}/{3}' -f $summary.Contains("`n移す先: $dest1`n"), [regex]::IsMatch($summary, '\n移すファイル 10 件（27\.8 KB）\n移さないファイル 3 件（移す先に同じ名前がある 1・読み取り専用 1・システムファイル 1）\n'), [regex]::IsMatch($summary, '\n結果: 移しました  10 件（27\.8 KB）  失敗 0 件  所要 \d+\.\d秒\n空になったフォルダ 8 件を削除しました\n'), $summary.Contains('※ 削除したファイルは元に戻せません')) 'True/True/True/False'
    Assert-Equal '移せなかったファイルの表' ([regex]::IsMatch($summary, '\n■ 移せなかったファイル\n 理由            ファイル\n 同じ名前がある  ' + [regex]::Escape("$mv1\sub1\x.log") + '\n')) 'True'

    Write-Section '移す（別のドライブ）'
    if ($hasShare) {
        $mv2 = Join-Path $workRoot 'mv2'
        $dest2 = Join-Path $workRoot 'mvdest2'
        $null = Build-Tree $mv2
        Set-ZoneMark "$mv2\a.log"
        [void][IO.Directory]::CreateDirectory($dest2)
        $dest2Unc = ConvertTo-AdminShare '127.0.0.1' $dest2
        $run = Invoke-Tool 'move-cross' @('-Path', $mv2, '-OlderThan', '90', '-MoveTo', $dest2Unc, '-Force')
        $x = Read-Result $run.Dir
        Assert-Equal '終了コード 0・すべて移した' ('{0}/{1}' -f $run.Code, @($x.Files | Where-Object { $_.'結果' -eq '移した' }).Count) '0/11'
        $w = New-Object IO.FileInfo "$dest2\sub4\deep\w.log"
        $a = [IO.File]::ReadAllBytes("$dest2\a.log")
        Assert-Equal '中身・日時・属性を写し、元を削除する' ('{0}/{1}/{2}/{3}/{4}/{5}' -f $a.Length, $a[4999], ($w.CreationTime -eq $old), ($w.LastWriteTime -eq $older), ([IO.File]::GetAttributes("$dest2\hidden.log")), [IO.File]::Exists("$mv2\a.log")) '5000/65/True/True/Hidden/False'
        Assert-Equal '代替データストリームは写さない（README どおり）' (Test-ZoneMark "$dest2\a.log") 'False'
        Assert-Equal '移した先の列は、指定した共有フォルダのパス' (Get-Row $x.Files $mv2 'a.log' 'ファイル' '移した先') "$dest2Unc\a.log"
        Assert-Equal '260 文字を超えるパスも移す' ([IO.File]::Exists('\\?\' + "$dest2\$longRel")) 'True'
    } else {
        Write-Skip '別のドライブへの移動' '管理共有を開けないため'
    }

    Write-Section '確認を待つ間の変化（ツールの確認の後に処理を差し込んだコピー）'
    $hook = Join-Path $workRoot 'hook'
    $victim = Join-Path $workRoot 'victim'
    $null = Build-Tree $hook
    New-AgedFile "$victim\x.log" 4000 $old $old
    $junctions.Add("$hook\sub1")
    $hookScript = New-HelperScript 'hook-change.ps1' @(
        ('[IO.File]::AppendAllText(''{0}\a.log'', ''more'')' -f $hook)
        ('[IO.File]::Delete(''{0}\note.txt'')' -f $hook)
        ('[IO.Directory]::Move(''{0}\sub1'', ''{0}\sub1_moved'')' -f $hook)
        ('$null = & cmd.exe /d /c mklink /J "{0}\sub1" "{1}" 2>&1' -f $hook, $victim)
        ('[IO.File]::SetAttributes(''{0}\hidden.log'', ''Hidden, ReadOnly'')' -f $hook)
        ('[IO.File]::WriteAllText(''{0}\sub4\deep\new.txt'', ''x'')' -f $hook)
    )
    $marker = '                # 何を削除しようとしているかを、削除する前に保存しておく（途中で止まっても残るように）'
    $hookTool = New-PatchedTool 'hook' $marker (("                & '{0}'`n" -f $hookScript) + $marker)
    $run = Invoke-Tool 'hook' @('-Path', $hook, '-OlderThan', '90', '-Delete', '-Force', '-RemoveEmptyFolder') '' $hookTool
    $x = Read-Result $run.Dir
    $changed = '予定を立てた後に変更されたため、削除していません'
    $gone = '予定を立てた後になくなっていました'
    Assert-Equal '終了コード 1' $run.Code 1
    Assert-Equal '更新されたファイル・読み取り専用になったファイルは削除しない' ('{0}|{1}' -f (Get-Row $x.Files $hook 'a.log' 'ファイル' '結果'), (Get-Row $x.Files $hook 'hidden.log' 'ファイル' '備考')) "失敗|$changed"
    Assert-Equal 'なくなっていたファイル' ('{0}|{1}' -f (Get-Row $x.Files $hook 'note.txt' 'ファイル' '結果'), (Get-Row $x.Files $hook 'note.txt' 'ファイル' '備考')) "削除しなかった|$gone"
    Assert-Equal '途中のフォルダがジャンクションに変わった: -Path の外を消さない' ('{0}|{1}|{2}|{3}' -f (Get-Row $x.Files $hook 'sub1\x.log' 'ファイル' '結果'), (Get-Row $x.Files $hook 'sub1\x.log' 'ファイル' '備考'), [IO.File]::Exists("$victim\x.log"), [IO.File]::Exists("$hook\sub1_moved\x.log")) ('失敗|ファイルの実際の場所が、調べたときと違います（途中のフォルダがリンクに変わったなど）: {0}\x.log|True|True' -f $victim)
    Assert-Equal 'ジャンクションの先にないファイルは、なくなっていた扱い' (Get-Row $x.Files $hook 'sub1\y.txt' 'ファイル' '結果') '削除しなかった'
    Assert-Equal 'ほかのファイルは削除した' (@($x.Files | Where-Object { $_.'結果' -eq '削除した' }).Count) 6
    Assert-Equal '空のフォルダ: リンクに変わったもの・ファイルが増えたものは削除しない' ('{0}|{1}|{2}|{3}' -f (Get-Row $x.Empty $hook 'sub1' 'フォルダ' '備考'), (Get-Row $x.Empty $hook 'sub4\deep' 'フォルダ' '備考'), (Get-Row $x.Empty $hook 'sub4' 'フォルダ' '結果'), [IO.Directory]::Exists("$hook\sub1")) '調べた後で、フォルダがリンクに変わっていました|中にファイルかフォルダが残っています|削除しなかった|True'
    Assert-Equal '結果の行（なくなっていた件数）と、削除できなかったファイルの表' ('{0}/{1}' -f [regex]::IsMatch($x.Summary, '\n結果: 削除しました  6 件（17\.1 KB）  失敗 3 件  所要 \d+\.\d秒  なくなっていた 2 件\n'), [regex]::IsMatch($x.Summary, '\n■ 削除できなかったファイル\n 理由            ファイル\n 予定の後に変更  ')) 'True/True'

    Write-Section '削除の途中の差し替え（別のプロセスが、途中のフォルダをジャンクションに差し替える）'
    $attack = Join-Path $workRoot 'attack'
    $victim2 = Join-Path $workRoot 'victim2'
    for ($i = 0; $i -lt 3000; $i++) { New-AgedFile ("$attack\sub\a{0:D4}.dat" -f $i) 1 $old $old }
    New-AgedFile "$attack\sub\secret.dat" 4 $old $old
    New-AgedFile "$victim2\secret.dat" 4 $old $old
    $junctions.Add("$attack\sub")
    $watcher = New-HelperScript 'watcher.ps1' @(
        'param([string]$Base, [string]$Victim)'
        '$first = Join-Path $Base ''sub\a0000.dat'''
        '$deadline = [DateTime]::UtcNow.AddSeconds(60)'
        'while ([IO.File]::Exists($first) -and [DateTime]::UtcNow -lt $deadline) { }'
        '$swapped = $false'
        'while (-not $swapped -and [DateTime]::UtcNow -lt $deadline) { try { [IO.Directory]::Move((Join-Path $Base ''sub''), (Join-Path $Base ''sub_moved'')); $swapped = $true } catch { } }'
        'if ($swapped) { $null = & cmd.exe /d /c mklink /J (Join-Path $Base ''sub'') $Victim 2>&1 }'
    )
    $watchProc = Start-Process -FilePath (Get-Command $Exe).Source -WindowStyle Hidden -PassThru -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $watcher), '-Base', ('"{0}"' -f $attack), '-Victim', ('"{0}"' -f $victim2))
    Start-Sleep -Milliseconds 1500
    $run = Invoke-Tool 'attack' @('-Path', $attack, '-OlderThan', '90', '-Delete', '-Force')
    if (-not $watchProc.WaitForExit(60000)) { $watchProc.Kill() }
    $x = Read-Result $run.Dir
    $swappedDone = [IO.Directory]::Exists("$attack\sub_moved")
    Assert-Equal '-Path の外のファイルは消さない' ([IO.File]::Exists("$victim2\secret.dat")) 'True'
    if ($swappedDone) {
        Assert-Equal '差し替えた後の同じパスのファイルは、失敗（実際の場所が違う）' (Get-Row $x.Files $attack 'sub\secret.dat' 'ファイル' '備考') ('ファイルの実際の場所が、調べたときと違います（途中のフォルダがリンクに変わったなど）: {0}\secret.dat' -f $victim2)
    } else {
        Write-Skip '差し替えた後のファイル' '差し替える前に削除が終わったため'
    }

    Write-Section '削除できないもの（権限なし・使用中・読み取り専用のフォルダ）'
    $deny = Join-Path $workRoot 'deny'
    New-AgedFile "$deny\p\old1.log" 1000 $old $old
    New-AgedFile "$deny\q\old2.log" 1000 $old $old
    $denyOk = (Add-DenyDelete "$deny\p\old1.log") -and (Add-DenyDelete "$deny\q")
    if ($denyOk) {
        $run = Invoke-Tool 'deny' @('-Path', $deny, '-OlderThan', '90', '-Delete', '-Force', '-RemoveEmptyFolder')
        $x = Read-Result $run.Dir
        Assert-Equal '削除を拒否されたファイルとフォルダは失敗（日本語の理由）' ('{0}/{1}/{2}/{3}' -f $run.Code, (Get-Row $x.Files $deny 'p\old1.log' 'ファイル' '備考'), (Get-Row $x.Files $deny 'q\old2.log' 'ファイル' '結果'), (Get-Row $x.Empty $deny 'q' 'フォルダ' '備考')) '1/アクセスが拒否されました/削除した/アクセスが拒否されました'
        # p は old1.log が残るので削除しない（Windows が空でないと断る）。q は削除を拒否されて失敗
        Assert-Equal '理由の表と、空のフォルダの結果の行' ('{0}/{1}/{2}' -f [regex]::IsMatch($x.Summary, '\n アクセス拒否  ' + [regex]::Escape("$deny\p\old1.log") + '\n'), $x.Summary.Contains('空になったフォルダ 0 件を削除しました（削除しなかった 1 件・失敗 1 件）'), (Get-Row $x.Empty $deny 'p' 'フォルダ' '備考')) 'True/True/中にファイルかフォルダが残っています'
        Assert-Equal '管理者でないときは、管理者の案内を出す' ($x.Summary.Contains('※ 管理者として実行すると、削除できるものがあります')) (-not $isAdmin)
    } else {
        Write-Skip '削除を拒否されたファイル' '削除の拒否を設定できなかったため'
    }
    $busy = Join-Path $workRoot 'busy'
    New-AgedFile "$busy\held.log" 1000 $old $old
    New-AgedFile "$busy\cwdir\old.log" 1000 $old $old
    New-AgedFile "$busy\keepro\old.log" 1000 $old $old
    [IO.File]::SetAttributes("$busy\keepro", [IO.FileAttributes]'Directory, ReadOnly')
    $held = New-Object IO.FileStream("$busy\held.log", [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $holders.Add($held)
    $cwdProc = Start-Process -FilePath (Get-Command $Exe).Source -WindowStyle Hidden -PassThru -WorkingDirectory "$busy\cwdir" -ArgumentList @('-NoProfile', '-Command', 'Start-Sleep -Seconds 120')
    $holders.Add($cwdProc)
    Start-Sleep -Seconds 2
    $run = Invoke-Tool 'busy' @('-Path', $busy, '-OlderThan', '90', '-Delete', '-Force', '-RemoveEmptyFolder')
    $x = Read-Result $run.Dir
    Assert-Equal 'ほかのプロセスが開いているファイルは失敗（使用中）' ('{0}/{1}/{2}' -f $run.Code, (Get-Row $x.Files $busy 'held.log' 'ファイル' '備考'), [IO.File]::Exists("$busy\held.log")) '1/ほかのプロセスが使用中です/True'
    Assert-Equal 'ほかのプロセスが作業フォルダにしているフォルダは失敗（使用中）・読み取り専用の属性のフォルダは対象外' ('{0}|{1}|{2}|{3}' -f (Get-Row $x.Empty $busy 'cwdir' 'フォルダ' '備考'), (Get-Row $x.Empty $busy 'keepro' 'フォルダ' '結果'), (Get-Row $x.Empty $busy 'keepro' 'フォルダ' '備考'), [IO.Directory]::Exists("$busy\keepro")) 'ほかのプロセスが使用中です|対象外|読み取り専用のフォルダです|True'
    Assert-Equal '理由の表（使用中）' ([regex]::IsMatch($x.Summary, '\n 使用中  ' + [regex]::Escape("$busy\held.log") + '\n')) 'True'
    $held.Dispose()
    try { $cwdProc.Kill() } catch { }

    Write-Section '指定できない場所'
    $sysRoot = $env:SystemRoot
    $osText = 'この場所は指定できません（OS やアプリが動くのに必要なファイルがあるため）: '
    $userText = 'この場所は指定できません（ユーザーのフォルダには、設定やデータのファイルがあるため）: '
    $cases = @(
        @('ドライブのルート', "$env:SystemDrive\", "ドライブのルートは指定できません: $env:SystemDrive\"),
        @('ドライブ名だけ', $env:SystemDrive, "ドライブのルートは指定できません: $env:SystemDrive\"),
        @('Windows のフォルダ', $sysRoot, ($osText + $sysRoot)),
        @('System32', "$sysRoot\System32", ($osText + "$sysRoot\System32")),
        @('Program Files', $env:ProgramFiles, ($osText + $env:ProgramFiles)),
        @('ProgramData', $env:ProgramData, ($osText + $env:ProgramData)),
        @('Users', "$env:SystemDrive\Users", ($osText + "$env:SystemDrive\Users")),
        @('自分のプロファイル', $env:USERPROFILE, ($userText + $env:USERPROFILE)),
        @('Public', "$env:SystemDrive\Users\Public", ($userText + "$env:SystemDrive\Users\Public")),
        @('-Path がジャンクション', "$main\jlink", ('-Path の途中にリンク（ジャンクション）があります。リンク先の実際のフォルダを指定してください: {0}\jlink' -f $main))
    )
    if ($hasShare) {
        $cases += , @('管理共有のルート', '\\127.0.0.1\C$', 'ドライブのルートは指定できません: \\127.0.0.1\C$')
        $cases += , @('管理共有の Windows', '\\127.0.0.1\C$\Windows', ($osText + '\\127.0.0.1\C$\Windows'))
        $cases += , @('管理共有のユーザーのフォルダ', (ConvertTo-AdminShare '127.0.0.1' $env:USERPROFILE), ($userText + (ConvertTo-AdminShare '127.0.0.1' $env:USERPROFILE)))
    }
    if ([IO.Directory]::Exists('\\127.0.0.1\ADMIN$\System32')) {
        $cases += , @('ADMIN$', '\\127.0.0.1\ADMIN$', ($osText + '\\127.0.0.1\ADMIN$'))
        $cases += , @('ADMIN$ の System32', '\\127.0.0.1\ADMIN$\System32', ($osText + '\\127.0.0.1\ADMIN$\System32'))
    } else {
        Write-Skip 'ADMIN$' '\\127.0.0.1\ADMIN$ を開けないため'
    }
    foreach ($case in $cases) {
        $run = Invoke-Plain @('-Path', $case[1], '-OlderThan', '90', '-NoSave')
        Assert-Equal $case[0] ('{0}/{1}/{2}' -f $run.Code, $run.Log.ToUpperInvariant().Contains($case[2].ToUpperInvariant()), $run.Log.Contains('古いファイルの整理 (')) '1/True/False'
    }
    $run = Invoke-Plain @('-Path', "$env:LOCALAPPDATA\Temp\..\..\..", '-OlderThan', '90', '-NoSave')
    Assert-Equal '.. で上がった先も判定する' ('{0}/{1}' -f $run.Code, $run.Log.ToUpperInvariant().Contains(($userText + $env:USERPROFILE).ToUpperInvariant())) '1/True'
    $substLetter = Get-FreeLetter
    if ($substLetter) {
        $null = & subst.exe "${substLetter}:" $sysRoot 2>&1
        if ([IO.Directory]::Exists("${substLetter}:\System32")) {
            $run = Invoke-Plain @('-Path', "${substLetter}:\System32", '-OlderThan', '90', '-NoSave')
            Assert-Equal 'SUBST のドライブは、実際の場所で判定する' ('{0}/{1}/{2}' -f $run.Code, $run.Log.ToUpperInvariant().Contains(($osText + "$sysRoot\System32").ToUpperInvariant()), $run.Log.Contains("指定したパス ${substLetter}:\System32 の実際の場所です")) '1/True/True'
        } else {
            Write-Skip 'SUBST のドライブ' 'SUBST で割り当てられなかったため'
        }
        $null = & subst.exe "${substLetter}:" /d 2>&1
        $substLetter = $null
    } else {
        Write-Skip 'SUBST のドライブ' '空いているドライブ文字がないため'
    }

    Write-Section '引数の誤り'
    $cases = @(
        @('-Path なし', @('-OlderThan', '90'), '調べるフォルダを -Path で指定してください'),
        @('-OlderThan なし', @('-Path', $main), '何日より古いファイルを対象にするかを、-OlderThan で指定してください（例: -OlderThan 90）'),
        @('-OlderThan 0', @('-Path', $main, '-OlderThan', '0'), '-OlderThan は 1～36500 の数値で指定してください: 0'),
        @('-OlderThan が数値でない', @('-Path', $main, '-OlderThan', 'abc'), '-OlderThan は 1～36500 の数値で指定してください: abc'),
        @('-OlderThan 36501', @('-Path', $main, '-OlderThan', '36501'), '-OlderThan は 1～36500 の数値で指定してください: 36501'),
        @('-Delete と -MoveTo', @('-Path', $main, '-OlderThan', '90', '-Delete', '-MoveTo', $workRoot), '-Delete と -MoveTo は一緒に使えません'),
        @('-Force だけ', @('-Path', $main, '-OlderThan', '90', '-Force'), '-Force は -Delete か -MoveTo と一緒に指定してください'),
        @('-NoSave と -Delete', @('-Path', $main, '-OlderThan', '90', '-Delete', '-NoSave'), '削除するとき・移すときは -NoSave を使えません（記録を残すため）'),
        @('-MoveTo で -Path が 2 つ', @('-Path', "$main\sub1,$main\sub2", '-OlderThan', '90', '-MoveTo', $dest1), '-MoveTo を使うときは、-Path を 1 つだけ指定してください'),
        @('-MoveTo のフォルダがない', @('-Path', $main, '-OlderThan', '90', '-MoveTo', "$workRoot\nope"), "移す先のフォルダが見つかりません: $workRoot\nope。先に作ってから実行してください"),
        @('-MoveTo が -Path の中', @('-Path', $main, '-OlderThan', '90', '-MoveTo', "$main\sub2"), "移す先が、調べるフォルダと重なっています: $main と $main\sub2"),
        @('-Path にファイル', @('-Path', "$main\b.log", '-OlderThan', '90'), "フォルダを指定してください（ファイルは指定できません）: $main\b.log"),
        @('-Path がない', @('-Path', "$workRoot\no\such", '-OlderThan', '90'), "フォルダが見つかりません: $workRoot\no\such"),
        @('片方がもう片方の中', @('-Path', "$main,$main\sub1", '-OlderThan', '90'), "指定したフォルダの片方が、もう片方の中にあります: $main と $main\sub1"),
        @('どのパラメーターにも当たらない引数', @('-Path', $main, '-OlderThan', '90', '-Foo', 'bar'), 'どのパラメーターにも当たらない引数があります: -Foo bar')
    )
    foreach ($case in $cases) {
        $run = Invoke-Plain $case[1]
        Assert-Equal $case[0] ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains($case[2]), $run.Log.Contains('古いファイルの整理 (')) '1/True/False'
    }

    Write-Section '保存先'
    $blocker = Join-Path $workRoot 'blocker'
    [IO.File]::WriteAllText($blocker, 'x')
    $run = Invoke-Plain @('-Path', $main, '-OlderThan', '90', '-OutputFolder', "$blocker\out")
    Assert-Equal '一覧だけで保存先に書き込めない: 警告して画面だけ（フォルダをすべて表示）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため、結果は保存しません'), $run.Log.Contains('■ 古いファイルのあるフォルダ')) '0/True/True'
    $sv = Join-Path $workRoot 'sv'
    $null = Build-Tree $sv
    $svTree = Get-Tree $sv
    $run = Invoke-Plain @('-Path', $sv, '-OlderThan', '90', '-Delete', '-Force', '-OutputFolder', "$blocker\out")
    Assert-Equal '削除で保存先に書き込めない: 削除せずに止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため、削除せずに止めました'), ((Get-Tree $sv) -eq $svTree)) '1/True/True'
    $longOut = Join-Path $workRoot ('out\' + ('o' * 120) + '\' + ('p' * 120))
    $run = Invoke-Plain @('-Path', $main, '-OlderThan', '90', '-OutputFolder', $longOut)
    Assert-Equal '260 文字を超える保存先' ('{0}/{1}' -f $run.Code, (Get-ResultNames (Get-RunDir $longOut))) '0/files.csv,folders.csv,skipped.csv,summary.txt'

    Write-Section '本物では起こせない条件（ツールを 1 か所だけ書き換えたコピー）'
    # 削除した後の記録の保存に失敗したら、終了コード 1
    $p = New-PatchedTool 'save-after' "    Save-Result`n    `$finished = `$true" "    `$script:runDir = Join-Path `$runDir 'missing\deeper'`n    Save-Result`n    `$finished = `$true"
    $sa = Join-Path $workRoot 'sa'
    New-AgedFile "$sa\old.log" 10 $old $old
    $run = Invoke-Plain @('-Path', $sa, '-OlderThan', '90', '-Delete', '-Force', '-OutputFolder', "$workRoot\out\save-after") $p
    Assert-Equal '削除した後に記録を保存できない: 終了コード 1 と警告' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした'), [IO.File]::Exists("$sa\old.log")) '1/True/False'
    # 削除する前の予定の保存に失敗したら、削除しない
    $p = New-PatchedTool 'save-before' "                `$actStarted = `$true`n                Save-Result" "                `$actStarted = `$true`n                `$script:runDir = Join-Path `$runDir 'missing\deeper'`n                Save-Result"
    $sb = Join-Path $workRoot 'sb'
    New-AgedFile "$sb\old.log" 10 $old $old
    $run = Invoke-Plain @('-Path', $sb, '-OlderThan', '90', '-Delete', '-Force', '-OutputFolder', "$workRoot\out\save-before") $p
    Assert-Equal '予定を保存できない: 削除せずに止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('記録を保存できないため、削除せずに止めました'), [IO.File]::Exists("$sb\old.log")) '1/True/True'
    # Add-Type を使えない環境
    $p = New-PatchedTool 'no-addtype' "    if (-not ('WbRemoveOldFileNative' -as [type])) { Add-Type" "    throw 'テストのため Add-Type を使えない'`n    if (-not ('WbRemoveOldFileNative' -as [type])) { Add-Type"
    $run = Invoke-Tool 'no-addtype' @('-Path', $main, '-OlderThan', '90', '-RemoveEmptyFolder') '' $p
    $x = Read-Result $run.Dir
    Assert-Equal 'Add-Type を使えない: 一覧は同じ結果' ('{0}/{1}/{2}' -f $run.Code, ((Get-RowList $x.Files $main 'ファイル' '結果') -eq (Join-Sorted $listResults)), @($x.Empty).Count) '0/True/9'
    $run = Invoke-Plain @('-Path', $sv, '-OlderThan', '90', '-Delete', '-Force', '-OutputFolder', "$workRoot\out\no-addtype-del") $p
    Assert-Equal 'Add-Type を使えない: 削除は止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('Windows の API を呼ぶ準備ができません（Add-Type を使えません）。この環境では、削除・移動はできません'), ((Get-Tree $sv) -eq $svTree)) '1/True/True'
    # 確認を待つ間の Ctrl+C（Read-Host が止められたことを、パイプラインの停止の例外で再現する。catch には捕まらず、finally だけが動く）
    $p = New-PatchedTool 'ask-ctrlc' 'try { $answer = Read-Host $Words.Ask } catch { Stop-WithError $Words.CannotAsk }' 'throw (New-Object Management.Automation.PipelineStoppedException)'
    $ac = Join-Path $workRoot 'ac'
    New-AgedFile "$ac\old.log" 10 $old $old
    $run = Invoke-Tool 'ask-ctrlc' @('-Path', $ac, '-OlderThan', '90', '-Delete') '' $p
    $i = [array]::IndexOf([string[]]$run.Lines, '中断しました。結果は保存していません。')
    Assert-Equal '確認を待つ間の Ctrl+C: 確認の行の後ろで改行してから「中断しました」・保存しない・削除しない' ('{0}/{1}/{2}' -f $(if ($i -gt 0) { $run.Lines[$i - 1] -eq '' } else { "(行がない: $i)" }), ($null -eq $run.Dir), [IO.File]::Exists("$ac\old.log")) 'True/True/True'
    # -Path の実際の場所を読めない（古い共有フォルダの機器など）
    $p = New-PatchedTool 'no-final' "try { `$finalLong = ([WbRemoveOldFileNative]::GetFinalPath(`$rootLong)).TrimEnd('\') }" "try { throw 'テストのため実際の場所を読めない' }"
    $run = Invoke-Plain @('-Path', $sv, '-OlderThan', '90', '-Delete', '-Force', '-OutputFolder', "$workRoot\out\no-final") $p
    Assert-Equal '実際の場所を読めない: 削除は止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("フォルダの実際の場所を確かめられないため、削除せずに止めました: $sv（"), ((Get-Tree $sv) -eq $svTree)) '1/True/True'
    $run = Invoke-Plain @('-Path', $sv, '-OlderThan', '90', '-NoSave') $p
    Assert-Equal '実際の場所を読めない: 一覧は動く' ('{0}/{1}' -f $run.Code, $run.Log.Contains('古いファイル 13 件')) '0/True'
    # 暗号化（EFS）されたファイル（属性だけを書き換える）
    $p = New-PatchedTool 'efs' "                `$attrs = [int]`$entry.Attributes" "                `$attrs = [int]`$entry.Attributes; if (`$entry.Name -eq 'a.log') { `$attrs = `$attrs -bor 0x4000 }"
    $efs = Join-Path $workRoot 'efs'
    New-AgedFile "$efs\a.log" 10 $old $old
    New-AgedFile "$efs\b.log" 10 $old $old
    [void][IO.Directory]::CreateDirectory("$workRoot\efsdest1")
    $run = Invoke-Tool 'efs-same' @('-Path', $efs, '-OlderThan', '90', '-MoveTo', "$workRoot\efsdest1", '-Force') '' $p
    Assert-Equal '暗号化されたファイルも、同じドライブの中なら移す' ('{0}/{1}' -f $run.Code, [IO.File]::Exists("$workRoot\efsdest1\a.log")) '0/True'
    if ($hasShare) {
        New-AgedFile "$efs\a.log" 10 $old $old
        New-AgedFile "$efs\b.log" 10 $old $old
        [void][IO.Directory]::CreateDirectory("$workRoot\efsdest2")
        $run = Invoke-Tool 'efs-cross' @('-Path', $efs, '-OlderThan', '90', '-MoveTo', (ConvertTo-AdminShare '127.0.0.1' "$workRoot\efsdest2"), '-Force') '' $p
        $x = Read-Result $run.Dir
        Assert-Equal '暗号化されたファイルは、別のドライブへは移せない' ('{0}/{1}/{2}/{3}' -f $run.Code, (Get-Row $x.Files $efs 'a.log' 'ファイル' '結果'), (Get-Row $x.Files $efs 'a.log' 'ファイル' '備考'), (Get-Row $x.Files $efs 'b.log' 'ファイル' '結果')) '1/移せない/暗号化（EFS）されたファイルは、別のドライブへ移せません（暗号化が外れて写るため）/移した'
        Assert-Equal '移さないファイルの行' ($x.Summary.Contains('移さないファイル 1 件（暗号化されている 1）')) 'True'
        # 移す先の空き容量が足りない
        $p = New-PatchedTool 'no-free' "                `$free = Get-FreeSpace `$moveFull" "                `$free = 10L"
        $nf = Join-Path $workRoot 'nf'
        New-AgedFile "$nf\old.log" 2000 $old $old
        [void][IO.Directory]::CreateDirectory("$workRoot\nfdest")
        $run = Invoke-Plain @('-Path', $nf, '-OlderThan', '90', '-MoveTo', (ConvertTo-AdminShare '127.0.0.1' "$workRoot\nfdest"), '-Force', '-OutputFolder', "$workRoot\out\no-free") $p
        Assert-Equal '空き容量が足りない: 移さずに止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('移す先の空き容量が足りないため、移さずに止めました（必要 2.0 KB・空き 10 B）: \\127.0.0.1\C$'), [IO.File]::Exists("$nf\old.log")) '1/True/True'
    } else {
        Write-Skip '暗号化されたファイルの別のドライブへの移動と、空き容量' '管理共有を開けないため'
    }

    Write-Section 'Ctrl+C（削除の途中は、コンソールのキーの入力として受け取って止める）'
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
    $many = Join-Path $workRoot 'ctrlc'
    # キーを送る補助のプログラムは起動と C# のコンパイルに数秒かかるので、その間に削除が終わらないだけの件数にする
    for ($i = 0; $i -lt 10000; $i++) { New-AgedFile ("$many\f{0:D5}.log" -f $i) 1 $old $old }
    $out = Join-Path $workRoot 'out\ctrlc'
    $proc = Start-Process -FilePath (Get-Command $Exe).Source -WindowStyle Hidden -PassThru -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $tool), '-Path', ('"{0}"' -f $many), '-OlderThan', '90', '-Delete', '-Force', '-OutputFolder', ('"{0}"' -f $out))
    $null = $proc.Handle   # 終了コードを受け取れるように、終わる前にハンドルを開いておく
    # 削除する前に予定を files.csv に保存するので、それができたら削除が始まっている
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $ready = $false
    while (-not $ready -and -not $proc.HasExited -and $watch.Elapsed.TotalSeconds -lt 60) {
        $runDir = Get-RunDir $out
        $ready = $runDir -and [IO.File]::Exists((Join-Path $runDir 'files.csv'))
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
    $keyMode = if ($mode -match '^\d+$') { ([long]$mode -band 1) -eq 0 } else { "(読めない: $mode)" }
    $x = Read-Result (Get-RunDir $out)
    if ($sent -like 'AttachConsole*' -or $sent -like 'CONIN$*' -or $sent -eq '(送る前に終わった)') {
        Write-Skip '削除の途中で Ctrl+C' ('キーを送れなかったため（{0}）' -f $sent)
    } else {
        $kinds = Join-Sorted ($x.Files | ForEach-Object { $_.'結果' } | Select-Object -Unique)
        Assert-Equal '削除の途中: Ctrl+C をキーの入力として受け取り、そのキーで止まる（終了コード 1）' ('{0}/{1}/{2}/{3}' -f $sent, $keyMode, $proc.ExitCode, $kinds) 'ok/True/1/削除した,削除しなかった'
        Assert-Equal '削除の途中: 残りの備考と、中断の行' ('{0}/{1}' -f (@($x.Files | Where-Object { $_.'結果' -eq '削除しなかった' } | ForEach-Object { $_.'備考' } | Select-Object -Unique) -join ','), (@($x.Lines | Where-Object { $_ -eq '中断しました。それまでに削除したファイルの記録は残っています' }).Count)) '中断したため、削除していません/1'
    }
    # キーが届いたかどうかにかかわらず、結果の列とファイルは合う
    $broken = @($x.Files | Where-Object { ($_.'結果' -eq '削除した') -eq [IO.File]::Exists($_.'ファイル') }).Count
    Assert-Equal '削除の途中で止めても、記録とファイルが合う（「削除する予定」が残らない）' ('{0}/{1}' -f $broken, @($x.Files | Where-Object { $_.'結果' -eq '削除する予定' }).Count) '0/0'

    Write-Section 'コマンドプロンプトから起動'
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $run = Invoke-ViaCmd ('{0} -OlderThan 90 -NoSave -Path "{1}\"' -f $cmdBase, $main)
    Assert-Equal '最後の引数が "パス\"' ('{0}/{1}' -f $run.Code, $run.Log.Contains('古いファイル 13 件')) '0/True'
    $unknownOut = Join-Path $workRoot 'out_unknown'
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -OlderThan 90 -OutputFolder "{2}"' -f $cmdBase, $main, $unknownOut)
    Assert-Equal '"パス\" の後ろに引数: ヒントを出して止まる（%ERRORLEVEL% は 1）' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('-Path の値に引用符'), $run.Log.Contains('最後に \ を付けないでください'), (Test-Path -LiteralPath $unknownOut)) '1/True/True/False'
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -OlderThan "9 0" -NoSave' -f $cmdBase, $main)
    Assert-Equal '"パス\" の後ろに空白を含む値: 引用符のヒント' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-Path の値に引用符（"）が残っています')) '1/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -OlderThen 90 -NoSave' -f $cmdBase, $main)
    Assert-Equal 'パラメーターの名前の書き間違い' ('{0}/{1}' -f $run.Code, $run.Log.Contains('どのパラメーターにも当たらない引数があります: -OlderThen 90')) '1/True'
    $run = Invoke-ViaCmd ('{0} -Path ''{1}'' -OlderThan ''90'' -Include ''*.log'',''*.txt'' -NoSave' -f $cmdBase, $main)
    Assert-Equal 'cmd で '' で囲んだ値' ('{0}/{1}' -f $run.Code, $run.Log.Contains('古いファイル 13 件')) '0/True'
    $ni = Join-Path $workRoot 'ni'
    $null = Build-Tree $ni
    $niTree = Get-Tree $ni
    $run = Invoke-ViaCmd ('"{0}" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{1}" -Path "{2}" -OlderThan 90 -Delete -OutputFolder "{3}"' -f $Exe, $tool, $ni, (Join-Path $workRoot 'out\nonint'))
    Assert-Equal '-NonInteractive では確認できないので削除しない・結果を残さない' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('確認できないため、削除せずに止めました。確認せずに削除するときは -Force を付けてください'), ((Get-Tree $ni) -eq $niTree), [bool](Get-RunDir (Join-Path $workRoot 'out\nonint'))) '1/True/True/False'
    # 移す先は空のフォルダにする（前の項目で移したファイルがあると、すべて「移せない」になって確認まで進まない）
    $niDest = Join-Path $workRoot 'nidest'
    [void][IO.Directory]::CreateDirectory($niDest)
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -OlderThan 90 -MoveTo "{2}" -OutputFolder "{3}"' -f $cmdBase, $ni, $niDest, (Join-Path $workRoot 'out\stdin-empty'))
    Assert-Equal '標準入力が NUL のときも移さない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('確認できないため、移さずに止めました'), ((Get-Tree $ni) -eq $niTree)) '1/True/True'

    Write-Section 'そのほかの起動'
    $cwd = Join-Path $workRoot '[cwd]'
    New-AgedFile "$cwd\rel\old.log" 10 $old $old
    New-AgedFile "$workRoot\cwd\rel\decoy.log" 10 $old $old
    $run = Invoke-From $cwd @('-Path', 'rel', '-OlderThan', '90', '-NoSave')
    Assert-Equal '[ ] を含むフォルダから相対パスで起動する' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("対象: $cwd\rel"), $run.Log.Contains('古いファイル 1 件')) '0/True/True'
    $self = Join-Path $workRoot 'selfdir'
    [void][IO.Directory]::CreateDirectory($self)
    Copy-Item -LiteralPath $repoTool -Destination "$self\Remove-OldFile.ps1"
    [IO.File]::SetCreationTime("$self\Remove-OldFile.ps1", $old)
    [IO.File]::SetLastWriteTime("$self\Remove-OldFile.ps1", $old)
    $run = Invoke-Plain @('-Path', $self, '-OlderThan', '90', '-NoSave') "$self\Remove-OldFile.ps1"
    Assert-Equal '実行しているスクリプトは対象にしない（除外に数える）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('ファイル 0 件  古いファイル 0 件'), $run.Log.Contains('除外 1 件')) '0/True/True'
    $manyDirs = Join-Path $workRoot 'manydirs'
    for ($i = 1; $i -le 25; $i++) { New-AgedFile ("$manyDirs\f{0:D2}\a.log" -f $i) (10 * $i) $old $old }
    $run = Invoke-Tool 'manydirs' @('-Path', $manyDirs, '-OlderThan', '90')
    $x = Read-Result $run.Dir
    Assert-Equal '画面のフォルダは 20 件まで・folders.csv と summary はすべて' ('{0}/{1}/{2}/{3}' -f $run.Log.Contains('（ほか 5 フォルダは folders.csv を参照してください）'), $run.Log.Contains("$manyDirs\f01"), @($x.Folders).Count, $x.Summary.Contains("$manyDirs\f01")) 'True/False/25/True'

    Write-Section '制約付き言語モードとヘルプ'
    $clmCommand = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '{0}' -Path '{1}' -OlderThan 90 -OutputFolder '{2}'; exit `$LASTEXITCODE" -f $tool, $main, "$workRoot\clm"
    $clm = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command $clmCommand 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $clmCode = $LASTEXITCODE
    Assert-Equal '制約付き言語モード: 止まる' ('{0}/{1}/{2}' -f $clmCode, $clm.Contains('制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）'), (Test-Path -LiteralPath "$workRoot\clm")) '1/True/False'
    $help = (@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command ("Get-Help '{0}' -Full | Out-String -Width 300" -f $tool) 2>&1) | ForEach-Object { "$_" }) -join "`n"
    Assert-Equal 'ヘルプ（概要・ツール名と版・著作権）' ('{0}/{1}/{2}' -f $help.Contains('指定した日数より古いファイルを一覧にし'), $help.Contains('Remove-OldFile 1.0.0'), $help.Contains('Copyright (c) 2026 結進舎')) 'True/True/True'

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -OlderThan 90 -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $main.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        Assert-Equal 'fi-FI でも同じ結果' ((Get-RowList $x.Files $main 'ファイル' '結果') -eq (Join-Sorted $listResults)) 'True'
        Assert-Equal 'fi-FI でも日時・3 桁区切り・小数点の書式は同じ' ('{0}/{1}/{2}' -f ((Get-Row $x.Files $main 'a.log' 'ファイル' '作成日時') -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$'), (Get-Row $x.Files $main 'a.log' 'ファイル' 'サイズ(バイト)'), [regex]::IsMatch($x.Summary, '古いファイル 13 件（33\.7 KB）  所要 \d+\.\d秒')) 'True/5,000/True'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }

    Write-Section 'レビューで見つかった問題の再発防止'
    # 1. 調べた後の差し替え（確認を待つ間・削除の途中）は「確認を待つ間の変化」「削除の途中の差し替え」で確かめている
    # 2. ADMIN$・SUBST のドライブは「指定できない場所」、読み取り専用の属性のフォルダ・削除の拒否は「削除できないもの」で確かめている
    # 3. 記録の保存の失敗・Add-Type を使えない・実際の場所を読めない・EFS は「本物では起こせない条件」で確かめている
    # 4. PowerShell は C# の string の引数に渡した $null を "" にする（対象がすべて対象外になった）。一覧の結果の確かめで見ている
    '  （ほかの項目で確かめている）'
} catch {
    '  FAIL  テストが途中で止まりました: {0}（{1} 行目）' -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber
    $script:failCount++
} finally {
    foreach ($holder in $holders) { try { if ($holder -is [Diagnostics.Process]) { if (-not $holder.HasExited) { $holder.Kill() } } else { $holder.Dispose() } } catch { } }
    if ($substLetter) { $null = & subst.exe "${substLetter}:" /d 2>&1 }
    # アクセス拒否を戻し、ジャンクションを外してから消す
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    for ($i = $aclDenies.Count - 1; $i -ge 0; $i--) {
        try { $acl = Get-ItemAcl $aclDenies[$i].Path; [void]$acl.RemoveAccessRule($aclDenies[$i].Rule); Set-ItemAcl $aclDenies[$i].Path $acl } catch { }
    }
    $global:OutputEncoding = $savedOutputEncoding
    if ($savedInputEncoding) { try { [Console]::InputEncoding = $savedInputEncoding } catch { } }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            Start-Sleep -Milliseconds 500
            foreach ($junction in $junctions) { if ([IO.Directory]::Exists($junction)) { try { [IO.Directory]::Delete($junction) } catch { } } }
            Get-ChildItem -LiteralPath ('\\?\' + $workRoot) -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { try { if ($_ -is [IO.DirectoryInfo]) { $_.Attributes = 'Directory' } else { $_.Attributes = 'Normal' } } catch { } }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

'  END'
exit $script:failCount
