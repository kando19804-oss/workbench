<#
.SYNOPSIS
    Find-DuplicateFile の回帰テストです。

.DESCRIPTION
    一時フォルダに、内容が同じファイル、先頭だけが同じファイル、コピーらしい名前のファイル、ハードリンク、
    日本語のパス、260 文字を超えるパス、名前に [ ] を含むフォルダ、ジャンクションのループ、読み取り禁止のフォルダと
    ファイルなどを作り、Find-DuplicateFile を実行して期待値と照合します。設定ファイルと、止まる場合も確かめます。
    ダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダとファイルにだけアクセス拒否を設定します（終了時に戻します）。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Find-DuplicateFile を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Find-DuplicateFile\Find-DuplicateFile.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Find-DuplicateFile_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
# 結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーしたスクリプトを動かす
$tool     = Join-Path $workRoot 'tool\Find-DuplicateFile.ps1'
$root     = Join-Path $workRoot 'data'
$baseTime = New-Object DateTime 2026, 10, 1, 10, 0, 0
$userSid  = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied   = New-Object 'System.Collections.Generic.List[string]'
$junctions = New-Object 'System.Collections.Generic.List[string]'
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

# 内容を種（$Seed）で決める。同じ種とサイズなら同じ内容になる。260 文字を超えるパスも作れるように、\\?\ を付けて作る
function New-DataFile([string]$FilePath, [int]$Size, [int]$Seed, $Time = $baseTime) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    $bytes = New-Object 'byte[]' $Size
    (New-Object Random $Seed).NextBytes($bytes)
    [IO.File]::WriteAllBytes('\\?\' + $FilePath, $bytes)
    [IO.File]::SetLastWriteTime('\\?\' + $FilePath, $Time)
}

function New-TextFile([string]$FilePath, [string]$Text) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllText('\\?\' + $FilePath, $Text, (New-Object Text.UTF8Encoding $false))
}

# コンソールで実行すると、長い行は画面の幅で折り返され、出力に改行が入る。
# 画面の幅によって結果が変わらないように、文字を探す前に改行を取り除く
function ConvertTo-FlatText([string]$Text) { $Text -replace '\r?\n', '' }

function Get-RunDir([string]$Parent) {
    if (-not (Test-Path -LiteralPath $Parent)) { return $null }
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory | Where-Object { $_.Name -like 'DuplicateFile_*' })
    if ($found.Count -gt 0) { return $found[0].FullName }
    $null
}

# Log には、画面への出力から改行を取り除いたものを入れる
function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out  = Join-Path $workRoot "out_$Label"
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $dir  = Get-RunDir $out
    $read = {
        param([string]$Name)
        if (-not $dir) { return @() }
        $file = Join-Path $dir $Name
        if (-not (Test-Path -LiteralPath $file)) { return @() }
        @(Import-Csv -LiteralPath $file -Encoding UTF8)
    }
    [pscustomobject]@{
        Code = $code; Log = $log; Dir = $dir
        Groups = @(& $read 'groups.csv'); Files = @(& $read 'duplicates.csv')
        Copies = @(& $read 'copynames.csv'); Skipped = @(& $read 'skipped.csv')
    }
}

# 保存しないで動かす（画面と終了コードだけを見る）
function Invoke-NoSave([string[]]$ToolArgs) {
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -NoSave 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

function Test-Contains([string]$Text, [string]$Part) { $Text.IndexOf($Part, [StringComparison]::Ordinal) -ge 0 }

# 結果の相対パス（data\ より後ろ）を、並べ替えてつなぐ
function Get-MemberText($Rows) {
    $names = [string[]]@($Rows | Where-Object { $null -ne $_ } | ForEach-Object { $_.'ファイル'.Replace($root + '\', '') })
    [Array]::Sort($names, [StringComparer]::Ordinal)
    $names -join '|'
}

# あるファイルを含むグループの番号
function Get-GroupOf($Run, [string]$Rel) {
    $hit = @($Run.Files | Where-Object { $_.'ファイル' -eq "$root\$Rel" })
    if ($hit.Count -eq 0) { return '' }
    $hit[0].'グループ'
}

function Get-GroupFiles($Run, [string]$Rel) {
    $number = Get-GroupOf $Run $Rel
    @($Run.Files | Where-Object { $_.'グループ' -eq $number })
}

function Get-SkipReason($Run, [string]$FullPath) { (@($Run.Skipped | Where-Object { $_.'パス' -eq $FullPath }) | ForEach-Object { $_.'理由' }) -join '|' }

function Get-CopyRow($Run, [string]$Rel) { @($Run.Copies | Where-Object { $_.'ファイル' -eq "$root\$Rel" }) }

function Deny-Access([string]$TargetPath, [string]$Rights) {
    $null = & icacls.exe $TargetPath /deny "*${userSid}:$Rights" 2>&1
    $denied.Add($TargetPath)
}

function Test-FileBom([string]$FilePath) {
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

try {
    #region ダミー環境

    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    # 200 KB の同じ内容が 3 件（名前も場所も違う）。古い順に a → b → c
    New-DataFile "$root\a.bin" 200000 1 $baseTime
    New-DataFile "$root\sub\renamed.bin" 200000 1 $baseTime.AddDays(1)
    New-DataFile "$root\日本語 フォルダ\[x]\c.bin" 200000 1 $baseTime.AddDays(2)
    # 100 KB の同じ内容が 2 件（片方は 260 文字を超えるパス）
    $longDir = "$root\long\" + (('x' * 60 + '\') * 5)
    New-DataFile "$root\b.bin" 100000 2
    New-DataFile ($longDir + 'b.bin') 100000 2
    # サイズは同じで内容が違う
    New-DataFile "$root\diff1.bin" 150000 3
    New-DataFile "$root\diff2.bin" 150000 4
    # 先頭 64 KB は同じで、後ろだけが違う
    New-DataFile "$root\head1.bin" 90000 5
    $headBytes = [IO.File]::ReadAllBytes("$root\head1.bin")
    $headBytes[$headBytes.Length - 1] = [byte]($headBytes[$headBytes.Length - 1] -bxor 0xFF)
    [IO.File]::WriteAllBytes("$root\head2.bin", $headBytes)
    # 64 KB 以下の同じ内容（先頭の絞り込みを通らない）と、-MinSize より小さい同じ内容、0 バイト
    New-DataFile "$root\mid1.bin" 5000 6
    New-DataFile "$root\mid2.bin" 5000 6
    New-DataFile "$root\tiny1.bin" 100 7
    New-DataFile "$root\tiny2.bin" 100 7
    New-TextFile "$root\zero1.bin" ''
    New-TextFile "$root\zero2.bin" ''
    # 隠し・システム属性のファイルも調べる
    New-DataFile "$root\visible.bin" 60000 8
    New-DataFile "$root\hidden.bin" 60000 8
    [IO.File]::SetAttributes("$root\hidden.bin", [IO.FileAttributes]'Hidden, System')
    # Excel で値が変わる名前
    New-DataFile "$root\excel\001" 70000 9
    New-DataFile "$root\excel\-old" 70000 9
    # 飛ばすフォルダ（a.bin と同じ内容）と、飛ばさないフォルダ
    New-DataFile "$root\DuplicateFile_PC_20260101_120000\result.bin" 200000 1
    New-DataFile "$root\`$RECYCLE.BIN\deleted.bin" 200000 1
    New-DataFile "$root\System Volume Information\tracking.bin" 200000 1
    New-DataFile "$root\DuplicateFile_backup\user.bin" 80000 10
    New-DataFile "$root\user.bin" 80000 10

    # ジャンクションのループ
    $null = New-Item -ItemType Junction -Path "$root\jloop" -Target $root
    $junctions.Add("$root\jloop")

    # コピーらしい名前
    $base  = 'base-content'
    $other = 'base-contenX'
    New-TextFile "$root\names\報告書.xlsx" $base
    New-TextFile "$root\names\設定.ini" 'ini'
    New-TextFile "$root\names\二重.a" 'x'
    New-TextFile "$root\names\二重.b" 'y'
    $copyCases = @(
        @('報告書 - コピー.xlsx', 'コピー', '同じ', $base),
        @('報告書 - コピー (2).xlsx', 'コピー', '違う', $other),
        @('報告書のコピー.xlsx', 'コピー', '違う', 'longer-content-here'),
        @('コピー ～ 報告書.xlsx', 'コピー', '同じ', $base),
        @('コピー (2) ～ 報告書.xlsx', 'コピー', '同じ', $base),
        @('報告書 - Copy.xlsx', 'コピー', '同じ', $base),
        @('報告書 - Copy (2).xlsx', 'コピー', '同じ', $base),
        @('Copy of 報告書.xlsx', 'コピー', '同じ', $base),
        @('Copy (2) of 報告書.xlsx', 'コピー', '同じ', $base),
        @('報告書 - コピー - コピー.xlsx', 'コピー', '同じ', $base),
        @('孤児 - コピー.xlsx', 'コピー', '元が見つからない', 'x'),
        @('報告書 (1).xlsx', '番号付き', '同じ', $base),
        @('報告書(2).xlsx', '番号付き', '違う', $other),
        @('報告書（3）.xlsx', '番号付き', '同じ', $base),
        @('報告書_old.xlsx', '古い版', '同じ', $base),
        @('報告書-bak.xlsx', '古い版', '違う', $other),
        @('報告書 旧.xlsx', '古い版', '同じ', $base),
        @('報告書_バックアップ.xlsx', '古い版', '同じ', $base),
        @('報告書.xlsx.bak', '古い版', '同じ', $base),
        @('報告書.xlsx.old', '古い版', '同じ', $base),
        @('設定.bak', '古い版', '同じ', 'ini'),
        @('二重.bak', '古い版', '元が見つからない', 'z'),
        @('孤児_old.ini', '古い版', '元が見つからない', 'x'),
        @('報告書_20261102.xlsx', '日付付き', '同じ', $base),
        @('報告書-2026-11-02.xlsx', '日付付き', '同じ', $base),
        @('報告書 261102.xlsx', '日付付き', '違う', $other),
        @('報告書_1102.xlsx', '日付付き', '同じ', $base)
    )
    foreach ($case in $copyCases) { New-TextFile "$root\names\$($case[0])" $case[3] }
    # コピーの形ではない、または元がないので出さない名前
    $notCopies = '孤児 (1).xlsx', '孤児_20261102.xlsx', '報告書_1399.xlsx', '報告書 (2026).xlsx', '第1回.docx', '第2回.docx', '報告書（下書き）.xlsx'
    foreach ($name in $notCopies) { New-TextFile "$root\names\$name" 'n' }
    # 64 KB を超えるコピー（先頭から違う、すべて同じ）
    New-DataFile "$root\names\big.dat" 100000 11
    New-DataFile "$root\names\big - コピー.dat" 100000 12
    New-DataFile "$root\names\big (1).dat" 100000 11

    #endregion

    Write-Section '内容が同じファイル'
    $run = Invoke-Tool 'base' @('-Path', $root, '-MinSize', '1KB')
    Assert-Equal '終了コード（重複あり）' $run.Code 2
    Assert-Equal '保存したファイル' ((Get-ChildItem -LiteralPath $run.Dir | ForEach-Object { $_.Name } | Sort-Object) -join ',') 'copynames.csv,duplicates.csv,groups.csv,skipped.csv,summary.txt'
    Assert-Equal 'CSV と summary の BOM' (@(Get-ChildItem -LiteralPath $run.Dir | Where-Object { -not (Test-FileBom $_.FullName) }).Count) 0
    $members = Get-GroupFiles $run 'a.bin'
    Assert-Equal '名前と場所が違っても同じグループ' (Get-MemberText $members) 'a.bin|sub\renamed.bin|日本語 フォルダ\[x]\c.bin'
    Assert-Equal 'グループの中は古い順' (($members | ForEach-Object { $_.'ファイル名' }) -join ',') 'a.bin,renamed.bin,c.bin'
    Assert-Equal 'ハッシュは Get-FileHash と同じ' $members[0].'ハッシュ' (Get-FileHash -LiteralPath "$root\a.bin" -Algorithm SHA256).Hash
    Assert-Equal 'サイズは 3 桁区切り' $members[0].'サイズ(バイト)' '200,000'
    $groupRow = @($run.Groups | Where-Object { $_.'ハッシュ' -eq $members[0].'ハッシュ' })
    Assert-Equal '件数・実体の数・減らせる容量' ('{0}/{1}/{2}' -f $groupRow[0].'件数', $groupRow[0].'実体の数', $groupRow[0].'減らせる容量(バイト)') '3/3/400,000'
    $wastes = [long[]]@($run.Groups | ForEach-Object { $_.'減らせる容量(バイト)'.Replace(',', '') })
    $sortedWastes = [long[]]$wastes.Clone()
    [Array]::Sort($sortedWastes)
    [Array]::Reverse($sortedWastes)
    Assert-Equal '減らせる容量の大きい順' ($wastes -join ',') ($sortedWastes -join ',')
    Assert-Equal '260 文字を超えるパスも見つかる' (@(Get-GroupFiles $run 'b.bin').Count) 2
    Assert-Equal '\\?\ が残っていない' ((Get-ChildItem -LiteralPath $run.Dir | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join '').Contains('\\?\') 'False'
    Assert-Equal '64 KB 以下の同じ内容' (@(Get-GroupFiles $run 'mid1.bin').Count) 2
    Assert-Equal '隠し・システムのファイルも調べる' (Get-MemberText (Get-GroupFiles $run 'hidden.bin')) 'hidden.bin|visible.bin'
    Assert-Equal 'サイズだけ同じなら重複にしない' (Get-GroupOf $run 'diff1.bin') ''
    Assert-Equal '先頭だけ同じなら重複にしない' (Get-GroupOf $run 'head1.bin') ''
    Assert-Equal '-MinSize より小さいと比べない' (Get-GroupOf $run 'tiny1.bin') ''
    Assert-Equal '0 バイトは比べない' (Get-GroupOf $run 'zero1.bin') ''
    Assert-Equal 'Excel で変わる名前は ="..." で守る' ((Get-GroupFiles $run 'excel\001' | ForEach-Object { $_.'ファイル名' } | Sort-Object) -join ',') '="-old",="001"'

    Write-Section '飛ばすもの'
    Assert-Equal '結果のフォルダは調べず、記録する' ('{0}/{1}' -f (Get-GroupOf $run 'DuplicateFile_PC_20260101_120000\result.bin'), (Get-SkipReason $run "$root\DuplicateFile_PC_20260101_120000")) '/除外指定'
    Assert-Equal '同じ文字で始まるだけのフォルダは調べる' (Get-MemberText (Get-GroupFiles $run 'user.bin')) 'DuplicateFile_backup\user.bin|user.bin'
    Assert-Equal 'ごみ箱は調べず、記録する' ('{0}/{1}' -f (Get-GroupOf $run '$RECYCLE.BIN\deleted.bin'), (Get-SkipReason $run "$root\`$RECYCLE.BIN")) '/除外指定'
    Assert-Equal 'System Volume Information' ('{0}/{1}' -f (Get-GroupOf $run 'System Volume Information\tracking.bin'), (Get-SkipReason $run "$root\System Volume Information")) '/除外指定'
    Assert-Equal 'ジャンクションはたどらず、記録する' (Get-SkipReason $run "$root\jloop") 'リンク'
    Assert-Equal '画面の除外の件数' (Test-Contains $run.Log '除外 2 件') 'True'
    Assert-Equal '画面のリンクの件数' (Test-Contains $run.Log 'たどらなかったリンク 1 件') 'True'
    Assert-Equal '調べる場所のファイルを変えていない' ([IO.File]::GetLastWriteTime("$root\a.bin").ToString('s')) ($baseTime.ToString('s'))

    Write-Section 'コピーらしい名前'
    foreach ($case in $copyCases) {
        $row = @(Get-CopyRow $run "names\$($case[0])")
        $actual = if ($row.Count -eq 1) { '{0}/{1}' -f $row[0].'名前の形', $row[0].'元との比較' } else { "{0} 行" -f $row.Count }
        Assert-Equal $case[0] $actual ('{0}/{1}' -f $case[1], $case[2])
    }
    foreach ($name in $notCopies) { Assert-Equal "出さない: $name" (@(Get-CopyRow $run "names\$name").Count) 0 }
    Assert-Equal '元と思われるファイル' (@(Get-CopyRow $run 'names\報告書 - コピー.xlsx')[0].'元と思われるファイル') "$root\names\報告書.xlsx"
    Assert-Equal '.bak の元は拡張子だけが違うファイル' (@(Get-CopyRow $run 'names\設定.bak')[0].'元と思われるファイル') "$root\names\設定.ini"
    Assert-Equal '64 KB を超えて先頭から違うコピー' (@(Get-CopyRow $run 'names\big - コピー.dat')[0].'元との比較') '違う'
    Assert-Equal '64 KB を超えて同じコピー' (@(Get-CopyRow $run 'names\big (1).dat')[0].'元との比較') '同じ'
    Assert-Equal '名前の形の順に並ぶ' (($run.Copies | ForEach-Object { $_.'名前の形' } | Select-Object -Unique) -join ',') 'コピー,番号付き,古い版,日付付き'

    Write-Section '-MinSize の既定値と、名前の判定'
    New-DataFile "$workRoot\size\one1.bin" 1100000 20
    New-DataFile "$workRoot\size\one2.bin" 1100000 20
    New-DataFile "$workRoot\size\small1.bin" 900000 21
    New-DataFile "$workRoot\size\small2.bin" 900000 21
    New-TextFile "$workRoot\size\memo.txt" 'memo'
    New-TextFile "$workRoot\size\memo - コピー.txt" 'memo'
    $sizeRun = Invoke-Tool 'size' @('-Path', "$workRoot\size")
    Assert-Equal '既定は 1 MB 以上を比べる' ('{0}/{1}' -f @($sizeRun.Groups).Count, @($sizeRun.Files | Where-Object { $_.'ファイル名' -like 'one*' }).Count) '1/2'
    Assert-Equal '条件の表示' (Test-Contains $sizeRun.Log '内容を比べるのは 1 MB 以上のファイル') 'True'
    Assert-Equal '小さなコピーも名前で見つかる' (@($sizeRun.Copies | Where-Object { $_.'ファイル' -like '*memo - コピー.txt' })[0].'元との比較') '同じ'
    Assert-Equal '500KB の書き方' (@((Invoke-Tool 'size500' @('-Path', "$workRoot\size", '-MinSize', '500KB')).Groups).Count) 2
    Assert-Equal '小数の書き方（1.5GB）' (@((Invoke-Tool 'size15' @('-Path', "$workRoot\size", '-MinSize', '1.5GB')).Groups).Count) 0

    Write-Section '-Path・-Include・-Exclude・-OlderThan・-Top'
    New-DataFile "$workRoot\m1\x.bin" 3000 30
    New-DataFile "$workRoot\m2\y.bin" 3000 30
    New-DataFile "$workRoot\m,comma\z.bin" 3000 31
    New-DataFile "$workRoot\m,comma\w.bin" 3000 31
    $multi = Invoke-Tool 'multi' @('-Path', "$workRoot\m1,$workRoot\m2", '-MinSize', '1KB')
    Assert-Equal 'カンマ区切りでフォルダをまたぐ' ('{0}/{1}' -f $multi.Code, @($multi.Files).Count) '2/2'
    $comma = Invoke-Tool 'comma' @('-Path', "$workRoot\m,comma", '-MinSize', '1KB')
    Assert-Equal 'カンマを含む名前のフォルダ' ('{0}/{1}' -f $comma.Code, @($comma.Files).Count) '2/2'

    $inc = Invoke-Tool 'include' @('-Path', $root, '-MinSize', '1KB', '-Include', 'a.bin,renamed.*')
    Assert-Equal '-Include で絞る' (Get-MemberText $inc.Files) 'a.bin|sub\renamed.bin'
    $incCopy = Invoke-Tool 'includecopy' @('-Path', "$root\names", '-Include', '*コピー.xlsx')
    Assert-Equal '元は -Include に当てはまらなくても探す' (@($incCopy.Copies | Where-Object { $_.'ファイル' -like '*\報告書 - コピー.xlsx' })[0].'元との比較') '同じ'

    $ex = Invoke-Tool 'exclude' @('-Path', $root, '-MinSize', '1KB', '-Exclude', 'renamed.bin,[x],.\long,*\excel\*')
    Assert-Equal '名前・[ ]・.\ の除外' (Get-MemberText (Get-GroupFiles $ex 'b.bin')) ''
    Assert-Equal '除外の記録' (Get-SkipReason $ex "$root\sub\renamed.bin") '除外指定'
    Assert-Equal '[x] はそのままの文字' (Get-SkipReason $ex "$root\日本語 フォルダ\[x]") '除外指定'
    $fullEx = Invoke-Tool 'excludefull' @('-Path', $root, '-MinSize', '1KB', '-Exclude', "$root\sub")
    Assert-Equal 'フルパスの除外' (Get-MemberText (Get-GroupFiles $fullEx 'a.bin')) 'a.bin|日本語 フォルダ\[x]\c.bin'
    $exOrig = Invoke-Tool 'excludeorig' @('-Path', "$root\names", '-Exclude', '報告書.xlsx')
    Assert-Equal '除外したファイルは元にならない' (@($exOrig.Copies | Where-Object { $_.'ファイル' -like '*\報告書 - コピー.xlsx' })[0].'元との比較') '元が見つからない'
    $exText = Invoke-Tool 'excludetext' @('-Path', $root, '-Exclude', '001')
    Assert-Equal '除外の書き方も Excel で変わらない' (@($exText.Skipped | Where-Object { $_.'パス' -eq "$root\excel\001" })[0].'詳細') '="001"'

    New-DataFile "$workRoot\age\old1.bin" 3000 40 (Get-Date).AddDays(-400)
    New-DataFile "$workRoot\age\old2.bin" 3000 40 (Get-Date).AddDays(-400)
    New-DataFile "$workRoot\age\new.bin" 3000 40 (Get-Date)
    $age = Invoke-Tool 'older' @('-Path', "$workRoot\age", '-MinSize', '1KB', '-OlderThan', '365')
    Assert-Equal '-OlderThan で絞る' (($age.Files | ForEach-Object { $_.'ファイル名' } | Sort-Object) -join ',') 'old1.bin,old2.bin'

    $top = Invoke-Tool 'top' @('-Path', $root, '-MinSize', '1KB', '-Top', '1')
    Assert-Equal '-Top を超えた分の案内' ('{0}/{1}' -f (Test-Contains $top.Log 'グループは groups.csv を参照してください'), (Test-Contains $top.Log '件は copynames.csv を参照してください')) 'True/True'
    $summaryText = [IO.File]::ReadAllText((Join-Path $top.Dir 'summary.txt'))
    Assert-Equal 'summary.txt は省略しない' (Test-Contains $summaryText 'b.bin') 'True'

    $noSave = Invoke-NoSave @('-Path', "$workRoot\m1,$workRoot\m2", '-MinSize', '1KB', '-OutputFolder', "$workRoot\out_nosave")
    Assert-Equal '-NoSave は保存しない' ('{0}/{1}' -f $noSave.Code, (Test-Path -LiteralPath "$workRoot\out_nosave")) '2/False'

    Write-Section 'ハードリンク'
    New-DataFile "$workRoot\hard\one.bin" 3000 50
    New-DataFile "$workRoot\hard\pair1.bin" 4000 51
    New-DataFile "$workRoot\hard\pair3.bin" 4000 51
    New-DataFile "$workRoot\hard\grow1.bin" 4000 52
    $linked = $true
    try {
        $null = New-Item -ItemType HardLink -Path "$workRoot\hard\one_link.bin" -Target "$workRoot\hard\one.bin" -ErrorAction Stop
        $null = New-Item -ItemType HardLink -Path "$workRoot\hard\pair2.bin" -Target "$workRoot\hard\pair1.bin" -ErrorAction Stop
        $null = New-Item -ItemType HardLink -Path "$workRoot\hard\grow2.bin" -Target "$workRoot\hard\grow1.bin" -ErrorAction Stop
    } catch {
        $linked = $false
        Write-Skip 'ハードリンク' ('作れなかったため: {0}' -f $_.Exception.Message)
    }
    if ($linked) {
        # 片方の名前から書き足すと、もう片方の名前のフォルダの一覧は古いままになる
        $stream = [IO.File]::Open("$workRoot\hard\grow1.bin", [IO.FileMode]::Append)
        $stream.WriteByte(1)
        $stream.Close()
        $hard = Invoke-Tool 'hard' @('-Path', "$workRoot\hard", '-MinSize', '1KB')
        Assert-Equal 'すべて同じ実体のグループは出さない' (@($hard.Files | Where-Object { $_.'ファイル名' -like 'one*' }).Count) 0
        $pairRows = @($hard.Files | Where-Object { $_.'ファイル名' -like 'pair*' })
        Assert-Equal '同じ実体は 1 つとして数える' (@($hard.Groups | Where-Object { $_.'実体の数' -eq '2' }).Count) 1
        Assert-Equal '減らせる容量は実体で数える' (@($hard.Groups | Where-Object { $_.'実体の数' -eq '2' })[0].'減らせる容量(バイト)') '4,000'
        Assert-Equal '3 件で実体は 2 つ' ('{0}/{1}' -f $pairRows.Count, (($pairRows | ForEach-Object { $_.'実体' } | Sort-Object -Unique) -join ',')) '3/1,2'
        Assert-Equal 'ハードリンクの件数を表示する' (Test-Contains $hard.Log 'は、減らせる容量に数えていません') 'True'
        Assert-Equal '判定できなかったと出ない' (Test-Contains $hard.Log 'ハードリンクを判定できませんでした') 'False'
        Assert-Equal '一覧が古くても「調査中に変更」にしない' (@($hard.Skipped | Where-Object { $_.'理由' -eq '調査中に変更' }).Count) 0
    }

    Write-Section '読めない場所と、終了コード'
    New-DataFile "$workRoot\clean\a.bin" 3000 60
    New-DataFile "$workRoot\clean\b.bin" 3001 61
    $clean = Invoke-Tool 'clean' @('-Path', "$workRoot\clean", '-MinSize', '1KB')
    Assert-Equal '重複なしは 0' ('{0}/{1}' -f $clean.Code, (Test-Contains $clean.Log '（重複はありません）')) '0/True'
    $emptyDir = "$workRoot\emptydir"
    [void][IO.Directory]::CreateDirectory($emptyDir)
    Assert-Equal '空のフォルダは 0' (Invoke-Tool 'empty' @('-Path', $emptyDir)).Code 0

    New-DataFile "$workRoot\deny\ok\a.bin" 3000 62
    New-DataFile "$workRoot\deny\ok\b.bin" 3000 62
    New-DataFile "$workRoot\deny\secret\c.bin" 3000 62
    Deny-Access "$workRoot\deny\secret" '(RD)'
    New-DataFile "$workRoot\deny2\a.bin" 3000 63
    New-DataFile "$workRoot\deny2\b.bin" 3000 63
    Deny-Access "$workRoot\deny2\b.bin" '(R)'
    New-TextFile "$workRoot\deny2\s.txt" 'aaa'
    New-TextFile "$workRoot\deny2\s - コピー.txt" 'aaa'
    Deny-Access "$workRoot\deny2\s.txt" '(R)'
    if (Test-Path -LiteralPath "$workRoot\deny\secret\c.bin" -ErrorAction SilentlyContinue) {
        Write-Skip 'アクセス拒否' '設定できなかったため'
    } else {
        $deny = Invoke-Tool 'deny' @('-Path', "$workRoot\deny", '-MinSize', '1KB')
        Assert-Equal '読めないフォルダは飛ばして続ける' ('{0}/{1}/{2}' -f $deny.Code, @($deny.Files).Count, (Get-SkipReason $deny "$workRoot\deny\secret")) '2/2/権限なし'
        $deny2 = Invoke-Tool 'deny2' @('-Path', "$workRoot\deny2", '-MinSize', '1KB')
        Assert-Equal '読めないファイルは「権限なし」' (Get-SkipReason $deny2 "$workRoot\deny2\b.bin") '権限なし'
        Assert-Equal '重複がなくて読めない場所があると 1' $deny2.Code 1
        Assert-Equal '比べられなかった案内' (Test-Contains $deny2.Log '比べられなかった場所が') 'True'
        $denyCopy = @($deny2.Copies | Where-Object { $_.'ファイル' -like '*s - コピー.txt' })
        Assert-Equal '元を読めないコピー' ('{0}/{1}' -f $denyCopy[0].'元との比較', (Test-Contains $denyCopy[0].'備考' '権限なし')) '比べられなかった/True'
    }

    Write-Section 'オンライン専用ファイル'
    New-DataFile "$workRoot\online\a.bin" 3000 70
    New-DataFile "$workRoot\online\b.bin" 3000 70
    New-DataFile "$workRoot\online\d.bin" 2500 71
    New-TextFile "$workRoot\online\d - コピー.bin" 'short'
    New-DataFile "$workRoot\online\e.bin" 2500 71
    New-DataFile "$workRoot\online\e (1).bin" 2500 71
    try {
        foreach ($offline in 'b.bin', 'd.bin', 'e.bin') { [IO.File]::SetAttributes("$workRoot\online\$offline", [IO.FileAttributes]::Offline) }
        $online = Invoke-Tool 'online' @('-Path', "$workRoot\online", '-MinSize', '1KB')
        Assert-Equal '読まずに記録する' (Get-SkipReason $online "$workRoot\online\b.bin") 'オンライン専用'
        Assert-Equal '重複に出さない' (@($online.Files | Where-Object { $_.'ファイル名' -eq 'a.bin' }).Count) 0
        Assert-Equal 'サイズが違えば「違う」' (@($online.Copies | Where-Object { $_.'ファイル' -like '*d - コピー.bin' })[0].'元との比較') '違う'
        Assert-Equal 'サイズが同じなら「比べられなかった」' (@($online.Copies | Where-Object { $_.'ファイル' -like '*e (1).bin' })[0].'元との比較') '比べられなかった'
    } catch {
        Write-Skip 'オンライン専用ファイル' ('属性を付けられなかったため: {0}' -f $_.Exception.Message)
    }

    Write-Section '更新日時が壊れたファイル'
    New-DataFile "$workRoot\badtime\a.bin" 3000 80
    New-DataFile "$workRoot\badtime\b.bin" 3000 80
    New-TextFile "$workRoot\badtime\m.txt" 'mmm'
    New-TextFile "$workRoot\badtime\m - コピー.txt" 'mmm'
    $timeSet = $false
    try {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
public static class WorkbenchTestFileTime {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool SetFileTime(SafeFileHandle handle, IntPtr creation, IntPtr access, ref long write);
}
'@ -ErrorAction Stop
        foreach ($broken in 'b.bin', 'm.txt') {
            $handle = [IO.File]::Open("$workRoot\badtime\$broken", [IO.FileMode]::Open, [IO.FileAccess]::Write)
            $farFuture = [long]::MaxValue - 1
            $timeSet = [WorkbenchTestFileTime]::SetFileTime($handle.SafeFileHandle, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$farFuture)
            $handle.Close()
        }
    } catch { }
    if ($timeSet) {
        $bad = Invoke-Tool 'badtime' @('-Path', "$workRoot\badtime", '-MinSize', '1KB')
        Assert-Equal '止まらずに結果を保存する' ('{0}/{1}' -f [bool]$bad.Dir, (Test-Contains $bad.Log 'エラーで止まりました')) 'True/False'
        Assert-Equal 'コピーの行は残る' (@($bad.Copies | Where-Object { $_.'ファイル' -like '*m - コピー.txt' }).Count) 1
        Assert-Equal '元が見つからないとは言わない' (@($bad.Copies | Where-Object { $_.'ファイル' -like '*m - コピー.txt' })[0].'元との比較' -ne '元が見つからない') 'True'
        $badOld = Invoke-Tool 'badold' @('-Path', "$workRoot\badtime", '-MinSize', '1KB', '-OlderThan', '1')
        Assert-Equal '-OlderThan でも止まらない' (Test-Contains $badOld.Log 'エラーで止まりました') 'False'
    } else {
        Write-Skip '更新日時が壊れたファイル' '日時を設定できなかったため'
    }

    Write-Section '設定ファイル'
    $conf = "$workRoot\conf.ini"
    [IO.File]::WriteAllText($conf, ("# コメント`r`nPath = m1`r`nPath = m2`r`nPath = m,comma`r`nMinSize = 1KB`r`nTop = 1`r`nExclude = w.bin`r`nInclude = *.bin`r`nNoSave = false`r`n"), (New-Object Text.UTF8Encoding $false))
    $confRun = Invoke-Tool 'conf' @('-ConfigFile', $conf)
    Assert-Equal '複数の行の Path と、カンマを含む Path' ('{0}/{1}' -f $confRun.Code, @($confRun.Files).Count) '2/2'
    Assert-Equal '設定ファイルの表示' (Test-Contains $confRun.Log "設定ファイル: $conf") 'True'
    Assert-Equal 'Exclude を設定ファイルから読む' (Get-SkipReason $confRun "$workRoot\m,comma\w.bin") '除外指定'
    $cliRun = Invoke-Tool 'confcli' @('-ConfigFile', $conf, '-Path', "$workRoot\clean")
    Assert-Equal 'コマンドラインの -Path を優先する' $cliRun.Code 0
    $badCases = @(
        @('Exlude = a', '知らない項目です: Exlude'),
        @('MinSize = 10M', 'MinSize は 500KB、10MB のように指定してください: 10M'),
        @('Top = 0', 'Top は 1～1000 の数値で指定してください: 0'),
        @('OlderThan = x', 'OlderThan は 1～36500 の数値で指定してください: x'),
        @('NoSave = yes', 'NoSave は true か false で指定してください: yes'),
        @('Depth 3', '「項目 = 値」の形になっていません'),
        @("Top = 1`r`nTop = 2", 'Top が 2 回書かれています'),
        @('Path = %NO_SUCH_VARIABLE_X%', '定義されていない環境変数です')
    )
    foreach ($case in $badCases) {
        [IO.File]::WriteAllText("$workRoot\bad.ini", $case[0], (New-Object Text.UTF8Encoding $false))
        $result = Invoke-NoSave @('-ConfigFile', "$workRoot\bad.ini")
        Assert-Equal "書き間違い: $($case[1])" ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $case[1])) '1/True'
    }

    Write-Section '止まるとき'
    $stops = @(
        @('同じフォルダ', @('-Path', "$workRoot\m1,$workRoot\m1\"), '同じフォルダが 2 回指定されています'),
        @('中にあるフォルダ', @('-Path', "$root,$root\sub"), '片方が、もう片方の中にあります'),
        @('見つからないフォルダ', @('-Path', "$workRoot\no_such"), 'フォルダを開けません'),
        @('-MinSize の書き間違い', @('-Path', $root, '-MinSize', '10M'), '-MinSize は 500KB、10MB のように指定してください: 10M'),
        @('-MinSize が 0', @('-Path', $root, '-MinSize', '0'), '-MinSize は')
    )
    foreach ($stop in $stops) {
        $result = Invoke-NoSave $stop[1]
        Assert-Equal $stop[0] ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log $stop[2])) '1/True'
    }

    Write-Section '起動のしかた'
    # コマンドプロンプトで "C:\data\" のように書くと、\" が引用符のまま届く
    $bat = "$workRoot\quote.cmd"
    [IO.File]::WriteAllText($bat, ("@echo off`r`nchcp 65001 > nul`r`n{0} -NoProfile -ExecutionPolicy Bypass -File ""{1}"" -Path ""{2}\m1\"" -NoSave`r`nexit /b %errorlevel%`r`n" -f $Exe, $tool, $workRoot), (New-Object Text.UTF8Encoding $false))
    $quoteLog = ConvertTo-FlatText (& cmd.exe /c $bat 2>&1 | Out-String)
    Assert-Equal '最後が \" の引数' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $quoteLog '引用符（"）が残っています')) '1/True'

    # 名前に [ ] を含むフォルダから、相対パスで起動する
    New-DataFile "$workRoot\[br]\p.bin" 3000 90
    New-DataFile "$workRoot\[br]\q.bin" 3000 90
    $brBat = "$workRoot\bracket.cmd"
    [IO.File]::WriteAllText($brBat, ("@echo off`r`nchcp 65001 > nul`r`ncd /d ""{0}\[br]""`r`n{1} -NoProfile -ExecutionPolicy Bypass -File ""{2}"" -MinSize 1KB -NoSave`r`nexit /b %errorlevel%`r`n" -f $workRoot, $Exe, $tool), (New-Object Text.UTF8Encoding $false))
    $brLog = ConvertTo-FlatText (& cmd.exe /c $brBat 2>&1 | Out-String)
    Assert-Equal '[ ] を含むフォルダから起動する' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $brLog 'p.bin')) '2/True'

    # 保存先に書き込めないときは、画面表示だけで終わる（保存先がファイル）
    New-TextFile "$workRoot\notdir" 'x'
    $unwritable = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$workRoot\m1,$workRoot\m2" -MinSize 1KB -OutputFolder "$workRoot\notdir" 2>&1 | Out-String)
    Assert-Equal '書き込めない保存先' ('{0}/{1}' -f $LASTEXITCODE, (Test-Contains $unwritable '保存先に書き込めないため')) '2/True'
} finally {
    # アクセス拒否を戻してから消す
    foreach ($target in $denied) { $null = & icacls.exe $target /remove:d "*$userSid" 2>&1 }
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
