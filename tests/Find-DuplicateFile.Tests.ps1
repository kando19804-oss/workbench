<#
.SYNOPSIS
    Find-DuplicateFile の回帰テストです。

.DESCRIPTION
    一時フォルダに、内容が同じファイル、先頭だけが同じファイル、コピーらしい名前のファイル、ハードリンク、
    日本語のパス、260 文字を超えるパス、名前に [ ] を含むフォルダ、ジャンクションのループ、読み取り禁止のフォルダと
    ファイルなどを作り、Find-DuplicateFile を実行して期待値と照合します。設定ファイルと、止まる場合も確かめます。
    ダミー環境は終了時に削除します。

    「権限なし」を再現するため、テストが作ったフォルダとファイルにだけアクセス拒否を設定します（終了時に戻します）。
    ドライブのルートの確認のため、空いているドライブ文字を subst で一時的に割り当てます（終了時に外します）。
    C:\Users\Default の互換用リンクと、管理共有（\\localhost\C$）を読み取ります。読めない環境では、そのテストを飛ばします。
    本物の環境では起きない条件（保存の失敗、調べている間の変更など）は、ツールの処理を書き換えたコピーで確かめます。

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
$substDrive = $null
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
function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $out  = Join-Path $workRoot "out_$Label"
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1 | Out-String)
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

# 相対パスの基準を確かめるため、指定したフォルダから起動する
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

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch = Join-Path $workRoot 'run.cmd'
    [IO.File]::WriteAllText($batch, ("@echo off`r`nchcp 65001 > nul`r`n{0}`r`nexit /b %errorlevel%`r`n" -f $CommandLine), (New-Object Text.UTF8Encoding $false))
    $log = ConvertTo-FlatText (& cmd.exe /d /c $batch 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 定数や処理を 1 か所だけ書き換えたツールのコピーを作る（本物の環境では起きない条件を確かめるため）
function New-PatchedTool([string]$Name, [string]$From, [string]$To) {
    $text = [IO.File]::ReadAllText($repoTool)
    if (-not $text.Contains($From)) { throw "書き換える場所が見つかりません: $From" }
    $file = Join-Path $workRoot "patched\$Name\Find-DuplicateFile.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text.Replace($From, $To), (New-Object Text.UTF8Encoding $true))
    $file
}

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $used -notcontains $_ -and -not [IO.Directory]::Exists("${_}:\") } | Select-Object -First 1
}

# 「調べたファイル N 件」の N を返す
function Get-Scanned([string]$Log) { [regex]::Match($Log, '調べたファイル ([\d,]+) 件').Groups[1].Value }

# ファイルの列を、指定したフォルダからの相対パスにして、序数の順に並べてつなぐ
function Get-RelText($Rows, [string]$Base) {
    $names = [string[]]@(@($Rows) | ForEach-Object { $_.'ファイル'.Replace($Base + '\', '') })
    [Array]::Sort($names, [StringComparer]::Ordinal)
    $names -join '|'
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

function Get-CopyRowIn($Run, [string]$FullPath) { @($Run.Copies | Where-Object { $_.'ファイル' -eq $FullPath }) }

function Get-SkipRow($Run, [string]$FullPath) { @($Run.Skipped | Where-Object { $_.'パス' -eq $FullPath }) }

# skipped.csv の行数
function Get-SkipCount([string]$File) { @(Import-Csv -LiteralPath $File -Encoding UTF8).Count }

# -NoSave を足さずに動かす（保存先を自分で指定するとき）
function Invoke-NoSaveless([string[]]$ToolArgs) {
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

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
    $excelNames = [string[]]@(Get-GroupFiles $run 'excel\001' | ForEach-Object { $_.'ファイル名' })
    [Array]::Sort($excelNames, [StringComparer]::Ordinal)
    Assert-Equal 'Excel で変わる名前は ="..." で守る' ($excelNames -join ',') '="-old",="001"'

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
    Assert-Equal 'コピーの形が重なっているときは、すべて取り除いて元を探す' (@(Get-CopyRow $run 'names\報告書 - コピー - コピー.xlsx')[0].'元と思われるファイル') "$root\names\報告書.xlsx"
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
    # 一覧の読み取りを拒否しても、名前を指定すれば読める。一覧を読めないことで、拒否できたかを確かめる
    $listDenied = $false
    try { [void][IO.Directory]::GetFileSystemEntries("$workRoot\deny\secret") } catch { $listDenied = $true }
    if (-not $listDenied) {
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

    Write-Section '画面と CSV の形'
    $view = "$workRoot\view"
    New-DataFile "$view\d1.bin" 2000 100
    New-DataFile "$view\d2.bin" 2000 100
    New-DataFile "$view\a1.bin" 3000 101 $baseTime
    New-DataFile "$view\a2.bin" 3000 101 $baseTime.AddDays(1)
    New-DataFile "$view\sub\a3.bin" 3000 101 $baseTime.AddDays(2)
    New-DataFile "$view\solo.bin" 2500 102
    New-TextFile "$view\memo.txt" 'memo'
    New-TextFile "$view\memo - コピー.txt" 'memo'
    New-TextFile "$view\memo (1).txt" 'memo2'
    New-TextFile "$view\孤児 - コピー.txt" 'x'
    $vr = Invoke-Tool 'view' @('-Path', $view, '-MinSize', '1KB')
    Assert-Equal '先頭の行' $vr.Log.StartsWith('重複ファイル検出 (Find-DuplicateFile 1.0.0)') 'True'
    Assert-Equal '対象の行' (Test-Contains $vr.Log "対象: $view") 'True'
    Assert-Equal '条件の行（1 KB）' (Test-Contains $vr.Log '条件: 内容を比べるのは 1 KB 以上のファイル') 'True'
    Assert-Equal '調べたファイルの行' ([regex]::IsMatch($vr.Log, '調べたファイル 10 件（15\.2 KB）  フォルダ 1 件  所要 \d+\.\d秒')) 'True'
    Assert-Equal '内容まで比べたファイルの行' (Test-Contains $vr.Log '内容まで比べたファイル 5 件（12.7 KB）') 'True'
    Assert-Equal '重複の行' (Test-Contains $vr.Log '重複 2 グループ・5 件  減らせる容量 7.8 KB') 'True'
    Assert-Equal 'コピーらしい名前の行' (Test-Contains $vr.Log 'コピーらしい名前 3 件（元と同じ 1・元と違う 1・元が見つからない 1）') 'True'
    Assert-Equal '見出し（上位 20）' ('{0}/{1}' -f (Test-Contains $vr.Log '■ 減らせる容量の大きいグループ（上位 20）'), (Test-Contains $vr.Log '■ コピーらしい名前のファイル（上位 20）')) 'True/True'
    Assert-Equal 'グループの 1 行目は減らせる容量の大きいもの' (Test-Contains $vr.Log ' 1. 減らせる容量 5.9 KB（2.9 KB × 3 件）') 'True'
    Assert-Equal 'グループの 2 行目' (Test-Contains $vr.Log ' 2. 減らせる容量 2.0 KB（2.0 KB × 2 件）') 'True'
    Assert-Equal 'グループの中のファイルの行' (Test-Contains $vr.Log ('    2026-10-01 10:00  {0}\a1.bin' -f $view)) 'True'
    Assert-Equal 'コピーの表の見出し' ([regex]::IsMatch($vr.Log, '名前の形\s+元との比較\s+サイズ\s+最終更新\s+ファイル')) 'True'
    Assert-Equal 'コピーの表の行' ([regex]::IsMatch($vr.Log, ('コピー\s+同じ\s+4 B\s+\d{4}-\d{2}-\d{2} \d{2}:\d{2}\s+' + [regex]::Escape("$view\memo - コピー.txt")))) 'True'
    Assert-Equal '保存先の行' (Test-Contains $vr.Log "保存先: $($vr.Dir)") 'True'
    $heads = @(
        @('groups.csv', '"グループ","件数","実体の数","1 件のサイズ(MB)","1 件のサイズ(バイト)","減らせる容量(MB)","減らせる容量(バイト)","ハッシュ"'),
        @('duplicates.csv', '"グループ","実体","ファイル","ファイル名","サイズ(MB)","サイズ(バイト)","最終更新日時","ハッシュ"'),
        @('copynames.csv', '"名前の形","元との比較","ファイル","元と思われるファイル","サイズ(MB)","サイズ(バイト)","元のサイズ(バイト)","最終更新日時","元の最終更新日時","備考"'),
        @('skipped.csv', '"理由","パス","詳細"')
    )
    foreach ($head in $heads) { Assert-Equal ('{0} の見出し' -f $head[0]) ([IO.File]::ReadAllLines((Join-Path $vr.Dir $head[0]))[0]) $head[1] }
    $row = @($vr.Groups)[0]
    Assert-Equal 'groups.csv の 1 行目' ('{0}/{1}/{2}/{3}/{4}/{5}/{6}' -f $row.'グループ', $row.'件数', $row.'実体の数', $row.'1 件のサイズ(MB)', $row.'1 件のサイズ(バイト)', $row.'減らせる容量(MB)', $row.'減らせる容量(バイト)') '1/3/3/0.00/3,000/0.01/6,000'
    $row = @($vr.Files)[0]
    Assert-Equal 'duplicates.csv の 1 行目' ('{0}/{1}/{2}/{3}/{4}/{5}' -f $row.'グループ', $row.'実体', $row.'ファイル', $row.'ファイル名', $row.'サイズ(MB)', $row.'最終更新日時') "1/1/$view\a1.bin/a1.bin/0.00/2026/10/01 10:00:00"
    Assert-Equal 'duplicates.csv の実体の番号' ((@($vr.Files) | Where-Object { $_.'グループ' -eq '1' } | ForEach-Object { $_.'実体' }) -join ',') '1,2,3'
    Assert-Equal 'ハッシュは 64 文字の大文字' ($row.'ハッシュ' -cmatch '^[0-9A-F]{64}$') 'True'
    $row = @(Get-CopyRowIn $vr "$view\memo - コピー.txt")[0]
    Assert-Equal 'copynames.csv の行' ('{0}/{1}/{2}/{3}/{4}/{5}' -f $row.'名前の形', $row.'元との比較', $row.'元と思われるファイル', $row.'サイズ(バイト)', $row.'元のサイズ(バイト)', $row.'備考') "コピー/同じ/$view\memo.txt/4/4/"
    Assert-Equal 'copynames.csv の日時の形' (($row.'最終更新日時' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') -and ($row.'元の最終更新日時' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$')) 'True'
    $row = @(Get-CopyRowIn $vr "$view\孤児 - コピー.txt")[0]
    Assert-Equal '元が見つからない行は元の列が空' ('{0}/{1}/{2}' -f $row.'元と思われるファイル', $row.'元のサイズ(バイト)', $row.'元の最終更新日時') '//'
    $summary = @([IO.File]::ReadAllLines((Join-Path $vr.Dir 'summary.txt')))
    Assert-Equal 'summary.txt の先頭' $summary[0] '重複ファイル検出 (Find-DuplicateFile 1.0.0)'
    Assert-Equal 'summary.txt の最後は保存先' $summary[-1] "保存先: $($vr.Dir)"
    Assert-Equal 'summary.txt の見出しに上位と書かない' (@($summary | Where-Object { $_ -like '■*上位*' }).Count) 0
    Assert-Equal 'スキップがなければスキップの行を出さない' (Test-Contains $vr.Log 'スキップ') 'False'
    # -Top
    $vt = Invoke-Tool 'viewtop' @('-Path', $view, '-MinSize', '1KB', '-Top', '1')
    Assert-Equal '-Top 1 の見出し' (Test-Contains $vt.Log '■ 減らせる容量の大きいグループ（上位 1）') 'True'
    Assert-Equal '-Top 1 のグループ' ('{0}/{1}' -f (Test-Contains $vt.Log ' 1. 減らせる容量'), (Test-Contains $vt.Log ' 2. 減らせる容量')) 'True/False'
    Assert-Equal '-Top 1 のグループの案内' (Test-Contains $vt.Log '（ほか 1 グループは groups.csv を参照してください）') 'True'
    Assert-Equal '-Top 1 のコピーの案内' (Test-Contains $vt.Log '（ほか 2 件は copynames.csv を参照してください）') 'True'
    Assert-Equal '-Top 1 でも CSV はすべて' ('{0}/{1}' -f @($vt.Groups).Count, @($vt.Copies).Count) '2/3'
    $vn = Invoke-NoSave @('-Path', $view, '-MinSize', '1KB', '-Top', '1')
    Assert-Equal '-NoSave の案内に CSV を書かない' ('{0}/{1}' -f (Test-Contains $vn.Log '（ほか 1 グループ）'), (Test-Contains $vn.Log 'groups.csv')) 'True/False'
    foreach ($bad in '0', '1001', 'abc') {
        $result = Invoke-NoSave @('-Path', $view, '-Top', $bad)
        Assert-Equal ('-Top {0} は止まる' -f $bad) $result.Code 1
    }
    Assert-Equal '-Top 1000 は使える' (Invoke-NoSave @('-Path', $view, '-MinSize', '1KB', '-Top', '1000')).Code 2
    # 3 桁区切りと、大きなサイズの表示（1 GB を超えるファイルは作らず、件数の区切りだけを確かめる）
    $thousand = "$workRoot\thousand"
    [void][IO.Directory]::CreateDirectory($thousand)
    foreach ($i in 1..1100) { [IO.File]::WriteAllText("$thousand\f$i.txt", 'x') }
    $tr = Invoke-NoSave @('-Path', $thousand)
    Assert-Equal '件数は 3 桁区切り' (Test-Contains $tr.Log '調べたファイル 1,100 件') 'True'

    # 全体を読むのは、先頭 64 KB も同じファイルだけ
    New-DataFile "$workRoot\headfilter\p.bin" 70000 170
    New-DataFile "$workRoot\headfilter\q.bin" 70000 171
    $r = Invoke-NoSave @('-Path', "$workRoot\headfilter", '-MinSize', '1KB')
    Assert-Equal '先頭が違えば全体を読まない' (Test-Contains $r.Log '内容まで比べたファイル 0 件') 'True'
    New-DataFile "$workRoot\headfilter2\p.bin" 70000 172
    $headBytes = [IO.File]::ReadAllBytes("$workRoot\headfilter2\p.bin")
    $headBytes[$headBytes.Length - 1] = [byte]($headBytes[$headBytes.Length - 1] -bxor 0xFF)
    [IO.File]::WriteAllBytes("$workRoot\headfilter2\q.bin", $headBytes)
    $r = Invoke-NoSave @('-Path', "$workRoot\headfilter2", '-MinSize', '1KB')
    Assert-Equal '先頭が同じなら全体を読む（重複にはしない）' ('{0}/{1}' -f (Test-Contains $r.Log '内容まで比べたファイル 2 件'), $r.Code) 'True/0'
    # ファイル名の Excel の対策
    $excelDir = "$workRoot\excelnames"
    $excelNames = @(
        @('全角の小数', '１．５', '="１．５"'), @('小数', '.5', '=".5"'), @('指数', '1e5', '="1e5"'), @('日付', '2026-01-01', '="2026-01-01"'), @('パーセント', '100%', '="100%"'),
        @('@', '@a', '="@a"'), @('=', '=a', '="=a"'), @('+', '+a', '="+a"'), @('ふつうの名前', 'normal', 'normal'), @('数字で始まる名前', '12 abc', '12 abc')
    )
    foreach ($case in $excelNames) { New-DataFile "$excelDir\$($case[1])" 2000 173 }
    $r = Invoke-Tool 'excelnames' @('-Path', $excelDir, '-MinSize', '1KB')
    foreach ($case in $excelNames) {
        Assert-Equal ('ファイル名の Excel 対策: {0}' -f $case[0]) (@($r.Files | Where-Object { $_.'ファイル' -eq "$excelDir\$($case[1])" })[0].'ファイル名') $case[2]
    }

    Write-Section '除外の書き方'
    $ex2 = "$workRoot\ex2"
    foreach ($rel in 'keep.bin', 'logs\a.bin', 'sub\logs\b.bin', 'sub\cache\c.bin', 'cache\d.bin', '[z].bin', 'z.bin') { New-DataFile "$ex2\$rel" 2000 110 }
    $allRel = '[z].bin|cache\d.bin|keep.bin|logs\a.bin|sub\cache\c.bin|sub\logs\b.bin|z.bin'
    $exCases = @(
        @('何も除外しない', '', $allRel),
        @('名前はどの階層でも', 'logs', '[z].bin|cache\d.bin|keep.bin|sub\cache\c.bin|z.bin'),
        @('大文字小文字を区別しない', 'LOGS', '[z].bin|cache\d.bin|keep.bin|sub\cache\c.bin|z.bin'),
        @('末尾の \ は無視', 'logs\', '[z].bin|cache\d.bin|keep.bin|sub\cache\c.bin|z.bin'),
        @('先頭の \ は直下だけ', '\logs', '[z].bin|cache\d.bin|keep.bin|sub\cache\c.bin|sub\logs\b.bin|z.bin'),
        @('先頭の .\ は直下だけ', '.\logs', '[z].bin|cache\d.bin|keep.bin|sub\cache\c.bin|sub\logs\b.bin|z.bin'),
        @('logs\* は直下の logs の中身', 'logs\*', '[z].bin|cache\d.bin|keep.bin|sub\cache\c.bin|sub\logs\b.bin|z.bin'),
        @('*\cache は 2 階層目より下', '*\cache', '[z].bin|cache\d.bin|keep.bin|logs\a.bin|sub\logs\b.bin|z.bin'),
        @('* のワイルドカード', '*.bin', ''),
        @('[ ] はそのままの文字', '[z].bin', 'cache\d.bin|keep.bin|logs\a.bin|sub\cache\c.bin|sub\logs\b.bin|z.bin'),
        @('? はそのままの文字', '?.bin', $allRel),
        @('カンマで複数', 'keep.bin,z.bin', '[z].bin|cache\d.bin|logs\a.bin|sub\cache\c.bin|sub\logs\b.bin'),
        @('フルパス', "$ex2\sub", '[z].bin|cache\d.bin|keep.bin|logs\a.bin|z.bin')
    )
    foreach ($case in $exCases) {
        $toolArgs = @('-Path', $ex2, '-MinSize', '1KB')
        if ($case[1]) { $toolArgs += @('-Exclude', $case[1]) }
        $result = Invoke-Tool ('ex' + [Array]::IndexOf($exCases, $case)) $toolArgs
        Assert-Equal ('-Exclude {0}: {1}' -f $case[1], $case[0]) (Get-RelText $result.Files $ex2) $case[2]
    }
    $result = Invoke-Tool 'exline' @('-Path', $ex2, '-MinSize', '1KB', '-Exclude', 'logs,keep.bin')
    Assert-Equal '除外の行と件数' ('{0}/{1}' -f (Test-Contains $result.Log '除外: logs, keep.bin'), (Test-Contains $result.Log '除外 3 件')) 'True/True'
    Assert-Equal '除外の詳細は当てはまった書き方' (@($result.Skipped | Where-Object { $_.'パス' -eq "$ex2\keep.bin" })[0].'詳細') 'keep.bin'
    $incCases = @(
        @('名前で絞る', 'keep.bin', '1'),
        @('大文字小文字を区別しない', 'KEEP.BIN', '1'),
        @('* のワイルドカード', 'k*', '1'),
        @('[ ] はそのままの文字', '[z].bin', '1'),
        @('? はそのままの文字', '?.bin', '0'),
        @('カンマで複数', 'keep.bin,z.bin', '2')
    )
    foreach ($case in $incCases) {
        $result = Invoke-NoSave @('-Path', $ex2, '-MinSize', '1KB', '-Include', $case[1])
        Assert-Equal ('-Include {0}: {1}' -f $case[1], $case[0]) (Get-Scanned $result.Log) $case[2]
    }
    $result = Invoke-NoSave @('-Path', $ex2, '-MinSize', '1KB', '-Include', '*.bin', '-Exclude', 'keep.bin')
    Assert-Equal '-Include より -Exclude を優先' ('{0}/{1}' -f (Get-Scanned $result.Log), (Test-Contains $result.Log '絞り込み: *.bin')) '6/True'

    Write-Section '-OlderThan と -MinSize の書き方'
    foreach ($bad in '0', '36501', 'abc') {
        Assert-Equal ('-OlderThan {0} は止まる' -f $bad) (Invoke-NoSave @('-Path', "$workRoot\age", '-OlderThan', $bad)).Code 1
    }
    $result = Invoke-NoSave @('-Path', "$workRoot\age", '-MinSize', '1KB', '-OlderThan', '365')
    Assert-Equal '-OlderThan の条件の行' ([regex]::IsMatch($result.Log, '条件: 最終更新が 365 日以上前のファイルのみ（\d{4}-\d{2}-\d{2} \d{2}:\d{2} より前）')) 'True'
    $sizeCases = @(@('バイト（単位なし）', '3000', 'True'), @('小文字', '2kb', 'True'), @('境目ちょうど', '3000', 'True'), @('1 バイト大きい', '3001', 'False'))
    foreach ($case in $sizeCases) {
        $result = Invoke-NoSave @('-Path', "$workRoot\m1,$workRoot\m2", '-MinSize', $case[1])
        Assert-Equal ('-MinSize {0}: {1}' -f $case[1], $case[0]) ($result.Code -eq 2) $case[2]
    }
    foreach ($bad in '-1', '1XB', '1.5.5MB', 'KB') {
        $result = Invoke-NoSave @('-Path', "$workRoot\m1", '-MinSize', $bad)
        Assert-Equal ('-MinSize {0} は止まる' -f $bad) ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log '-MinSize')) '1/True'
    }

    Write-Section 'コピーらしい名前の細かい形'
    $names2 = "$workRoot\names2"
    New-TextFile "$names2\a.txt" 'same'
    New-TextFile "$names2\b.txt" 'same'
    New-TextFile "$names2\zero.txt" ''
    New-TextFile "$names2\zero - コピー.txt" ''
    $extraCases = @(
        @('COPY OF a.txt', 'コピー', '同じ', 'same'), @('a - COPY.txt', 'コピー', '同じ', 'same'), @('a_OLD.txt', '古い版', '同じ', 'same'), @('a_backup.txt', '古い版', '同じ', 'same'),
        @('a_古.txt', '古い版', '同じ', 'same'), @('a_控え.txt', '古い版', '同じ', 'same'), @('a_0229.txt', '日付付き', '同じ', 'same'), @('a-2026-02-28.txt', '日付付き', '同じ', 'same'),
        @('b.old', '古い版', '同じ', 'same'), @('c.bak', '古い版', '元が見つからない', 'q')
    )
    foreach ($case in $extraCases) { New-TextFile "$names2\$($case[0])" $case[3] }
    # 全角の数字の番号は番号付きにしない（old を含む名前は、コピーらしい名前の下調べを通るので、番号の判定まで進む）
    New-TextFile "$names2\old.txt" 'same'
    $notExtra = 'a_20261301.txt', 'a_0230.txt', 'a (１).txt', 'old (１).txt', 'a (1000).txt', 'a_1.txt', 'copya.txt', 'anold.txt', 'a_19000101x.txt'
    foreach ($name in $notExtra) { New-TextFile "$names2\$name" 'n' }
    $extra = Invoke-Tool 'names2' @('-Path', $names2)
    foreach ($case in $extraCases) {
        $row = @(Get-CopyRowIn $extra "$names2\$($case[0])")
        Assert-Equal $case[0] $(if ($row.Count -eq 1) { '{0}/{1}' -f $row[0].'名前の形', $row[0].'元との比較' } else { '{0} 行' -f $row.Count }) ('{0}/{1}' -f $case[1], $case[2])
    }
    foreach ($name in $notExtra) { Assert-Equal "出さない: $name" (@(Get-CopyRowIn $extra "$names2\$name").Count) 0 }
    Assert-Equal '0 バイトのコピーは「同じ」' (@(Get-CopyRowIn $extra "$names2\zero - コピー.txt")[0].'元との比較') '同じ'
    Assert-Equal 'コピーらしい名前は終了コードに数えない' $extra.Code 0

    Write-Section '読めないファイルと、調べている間の変更'
    $lockDir = "$workRoot\lock"
    New-DataFile "$lockDir\a.bin" 3000 120
    New-DataFile "$lockDir\b.bin" 3000 120
    New-DataFile "$lockDir\c.bin" 3000 120
    $hold = [IO.File]::Open("$lockDir\b.bin", [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try {
        $lr = Invoke-Tool 'lock' @('-Path', $lockDir, '-MinSize', '1KB')
        Assert-Equal 'ほかのプログラムが使っているファイルは「読めない」' (Get-SkipReason $lr "$lockDir\b.bin") '読めない'
        Assert-Equal '残りで重複を出す' ('{0}/{1}' -f $lr.Code, (Get-RelText $lr.Files $lockDir)) '2/a.bin|c.bin'
        Assert-Equal 'スキップの行' (Test-Contains $lr.Log 'スキップ 1 件（読めない）') 'True'
    } finally {
        $hold.Dispose()
    }
    # 一覧を作った後でなくなったファイル（比べる直前に消すコピーを動かす）
    $goneDir = "$workRoot\gone"
    New-DataFile "$goneDir\a.bin" 3000 121
    New-DataFile "$goneDir\b.bin" 3000 121
    New-DataFile "$goneDir\c.bin" 3000 121
    $gonePath = "$goneDir\b.bin"
    $patched = New-PatchedTool 'gone' '$fullGroups = New-Object ''System.Collections.Generic.List[object]''' ("[IO.File]::Delete('{0}'); `$fullGroups = New-Object 'System.Collections.Generic.List[object]'" -f $gonePath)
    $result = Invoke-Tool 'gone' @('-Path', $goneDir, '-MinSize', '1KB') $patched
    Assert-Equal '読む前になくなったファイルは「調査中に変更」' (Get-SkipReason $result $gonePath) '調査中に変更'
    Assert-Equal 'なくなったときの詳細' (@(Get-SkipRow $result $gonePath)[0].'詳細') '調べている間にファイルがなくなりました'
    Assert-Equal 'なくなったファイルは重複に出さない' (Get-RelText $result.Files $goneDir) 'a.bin|c.bin'
    # 読んでいる間に更新日時が変わったファイル
    $changeDir = "$workRoot\change"
    New-DataFile "$changeDir\a.bin" 3000 122
    New-DataFile "$changeDir\b.bin" 3000 122
    New-DataFile "$changeDir\c.bin" 3000 122
    $patched = New-PatchedTool 'change' '[void]$sha256.TransformFinalBlock($hashBuffer, 0, 0)' ("if (`$Item.Name -eq 'b.bin') { [IO.File]::SetLastWriteTime(`$Item.Full, (Get-Date)) }; [void]`$sha256.TransformFinalBlock(`$hashBuffer, 0, 0)")
    $result = Invoke-Tool 'change' @('-Path', $changeDir, '-MinSize', '1KB') $patched
    Assert-Equal '読んでいる間に変わったファイルは「調査中に変更」' (Get-SkipReason $result "$changeDir\b.bin") '調査中に変更'
    Assert-Equal '変わったファイルは重複に出さない' (Get-RelText $result.Files $changeDir) 'a.bin|c.bin'
    # 一覧を作った後に変わっただけのファイルは、その時点の内容で比べる（サイズが同じグループのまま）
    $listDir = "$workRoot\listchange"
    New-DataFile "$listDir\a.bin" 3000 123
    New-DataFile "$listDir\b.bin" 3000 123
    $patched = New-PatchedTool 'listchange' '$fullGroups = New-Object ''System.Collections.Generic.List[object]''' ("[IO.File]::SetLastWriteTime('{0}\a.bin', (Get-Date)); `$fullGroups = New-Object 'System.Collections.Generic.List[object]'" -f $listDir)
    $result = Invoke-Tool 'listchange' @('-Path', $listDir, '-MinSize', '1KB') $patched
    Assert-Equal '読む前に変わっただけなら比べる' ('{0}/{1}' -f (Get-RelText $result.Files $listDir), @($result.Skipped).Count) 'a.bin|b.bin/0'
    # ハードリンクを判定できないとき
    New-DataFile "$workRoot\noid\a.bin" 3000 124
    New-DataFile "$workRoot\noid\b.bin" 3000 124
    $patched = New-PatchedTool 'noid' 'Add-Type -TypeDefinition $NativeSource -ErrorAction Stop' 'throw ''not allowed'''
    $result = Invoke-Tool 'noid' @('-Path', "$workRoot\noid", '-MinSize', '1KB') $patched
    Assert-Equal 'ハードリンクを判定できないときは知らせて続ける' ('{0}/{1}' -f $result.Code, (Test-Contains $result.Log '※ ハードリンクを判定できませんでした')) '2/True'
    if ($listDenied) {
        Assert-Equal '管理者でなければ案内する' (Test-Contains $deny.Log '※ 管理者として実行すると、読める範囲が広がります') ([string](-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)))
        Assert-Equal 'スキップの行（権限なし）' (Test-Contains $deny.Log 'スキップ 1 件（権限なし）') 'True'
        $row = @(Get-SkipRow $deny2 "$workRoot\deny2\b.bin")
        Assert-Equal 'skipped.csv の理由' $row[0].'理由' '権限なし'
    }

    Write-Section 'リンクと名前'
    $real = 'C:\Users\Default'
    if ([IO.Directory]::Exists("$real\Application Data") -or [IO.Directory]::Exists("$real\AppData")) {
        $result = Invoke-Tool 'compat' @('-Path', $real, '-MinSize', '1TB')
        $links = @($result.Skipped | Where-Object { $_.'詳細' -eq 'Windows の互換用リンク' })
        Assert-Equal 'Windows の互換用リンクをたどらない' ($links.Count -gt 0) 'True'
        Assert-Equal '互換用リンクの理由はリンク' (@($links | Where-Object { $_.'理由' -ne 'リンク' }).Count) 0
    } else {
        Write-Skip 'Windows の互換用リンク' 'C:\Users\Default がないため'
    }
    $symDir = "$workRoot\sym"
    New-DataFile "$symDir\real\a.bin" 3000 130
    New-DataFile "$symDir\real\b.bin" 3000 130
    $symMade = $false
    try {
        $null = New-Item -ItemType SymbolicLink -Path "$symDir\link" -Target "$symDir\real" -ErrorAction Stop
        $symMade = $true
    } catch { }
    if ($symMade) {
        $result = Invoke-Tool 'sym' @('-Path', $symDir, '-MinSize', '1KB')
        Assert-Equal 'シンボリックリンクをたどらない' ('{0}/{1}/{2}' -f (Get-SkipReason $result "$symDir\link"), @($result.Skipped)[0].'詳細', @($result.Files).Count) 'リンク/シンボリックリンク/2'
        [IO.Directory]::Delete("$symDir\link")
    } else {
        Write-Skip 'シンボリックリンク' '作れないため（管理者か開発者モードが必要）'
    }
    $caseDir = "$workRoot\case"
    New-DataFile "$caseDir\`$Recycle.Bin\a.bin" 3000 131
    New-DataFile "$caseDir\a.bin" 3000 131
    New-DataFile "$caseDir\DuplicateFile_PC_20260101_120000_2\b.bin" 3000 131
    $result = Invoke-Tool 'case' @('-Path', $caseDir, '-MinSize', '1KB')
    Assert-Equal '$Recycle.Bin も大文字小文字を区別しない' (Get-SkipReason $result "$caseDir\`$Recycle.Bin") '除外指定'
    Assert-Equal '_2 の付いた結果のフォルダも調べない' (Get-SkipReason $result "$caseDir\DuplicateFile_PC_20260101_120000_2") '除外指定'
    Assert-Equal '飛ばしたフォルダの中は数えない' (Get-Scanned $result.Log) '1'
    $emojiDir = "$workRoot\emoji"
    $emoji = [char]::ConvertFromUtf32(0x1F600)
    New-DataFile "$emojiDir\$emoji.bin" 3000 132
    New-DataFile "$emojiDir\a&b'c.bin" 3000 132
    $result = Invoke-Tool 'emoji' @('-Path', $emojiDir, '-MinSize', '1KB')
    Assert-Equal '絵文字と記号を含む名前' (Get-RelText $result.Files $emojiDir) "a&b'c.bin|$emoji.bin"

    Write-Section 'パスの指定のしかた'
    $r = Invoke-From "$workRoot\m1" @('-MinSize', '1KB', '-NoSave')
    Assert-Equal '-Path を省略すると現在のフォルダ' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $workRoot\m1")) '0/True'
    $r = Invoke-From $workRoot @('-Path', 'm1,m2', '-MinSize', '1KB', '-NoSave')
    Assert-Equal '-Path に相対パス' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $workRoot\m1, $workRoot\m2")) '2/True'
    $r = Invoke-From $workRoot @('-Path', '.\m1\..\m2', '-MinSize', '1KB', '-NoSave')
    Assert-Equal '-Path の .. を解決する' (Test-Contains $r.Log "対象: $workRoot\m2") 'True'
    $r = Invoke-NoSave @('-Path', ('\\?\' + "$workRoot\m1"), '-MinSize', '1KB')
    Assert-Equal '-Path に \\?\ を付けてもよい' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $workRoot\m1")) '0/True'
    $r = Invoke-NoSave @('-Path', "$workRoot\m1\", '-MinSize', '1KB')
    Assert-Equal '-Path の最後に \' (Test-Contains $r.Log "対象: $workRoot\m1") 'True'
    $r = Invoke-NoSave @('-Path', "$workRoot\M1,$workRoot\m1", '-MinSize', '1KB')
    Assert-Equal '大文字小文字だけ違う同じフォルダ' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '同じフォルダが 2 回指定されています')) '1/True'
    New-DataFile "$workRoot\m1x\y.bin" 3000 30
    $r = Invoke-NoSave @('-Path', "$workRoot\m1,$workRoot\m1x", '-MinSize', '1KB')
    Assert-Equal '名前が前方一致するだけのフォルダは別のフォルダ' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log 'もう片方の中')) '2/False'
    $letter = Get-FreeLetter
    if ($letter) {
        $null = & subst.exe "${letter}:" "$workRoot\ex2" 2>&1
        if ($LASTEXITCODE -eq 0) {
            $substDrive = "${letter}:"
            $r = Invoke-Tool 'root1' @('-Path', "$substDrive\", '-MinSize', '1KB', '-Exclude', '\logs')
            Assert-Equal 'ドライブのルート（\ あり）と、直下の除外' (Get-RelText $r.Files $substDrive) '[z].bin|cache\d.bin|keep.bin|sub\cache\c.bin|sub\logs\b.bin|z.bin'
            $r = Invoke-Tool 'root2' @('-Path', $substDrive, '-MinSize', '1KB')
            Assert-Equal 'ドライブのルート（ドライブ文字だけ）' (@($r.Files).Count) 7
            $r = Invoke-Tool 'root3' @('-Path', "$substDrive\", '-MinSize', '1KB', '-Exclude', "$substDrive\sub")
            Assert-Equal 'ドライブのルートとフルパスの除外' (Get-RelText $r.Files $substDrive) '[z].bin|cache\d.bin|keep.bin|logs\a.bin|z.bin'
        } else {
            Write-Skip 'ドライブのルート' 'subst を使えなかったため'
        }
    } else {
        Write-Skip 'ドライブのルート' '空いているドライブ文字がないため'
    }
    $uncRoot = '\\localhost\{0}$\{1}' -f $workRoot.Substring(0, 1), $workRoot.Substring(3)
    if ([IO.Directory]::Exists("$uncRoot\m1")) {
        $r = Invoke-Tool 'unc' @('-Path', "$uncRoot\m1,$uncRoot\m2", '-MinSize', '1KB')
        Assert-Equal 'UNC のパス' ('{0}/{1}' -f $r.Code, (Get-RelText $r.Files $uncRoot)) '2/m1\x.bin|m2\y.bin'
        Assert-Equal 'UNC のパスに \\?\ が残らない' (([IO.File]::ReadAllText((Join-Path $r.Dir 'duplicates.csv'))).Contains('\\?\')) 'False'
        $r = Invoke-Tool 'unclong' @('-Path', "$uncRoot\data\long", '-MinSize', '1KB')
        Assert-Equal 'UNC の 260 文字を超えるパス' (Get-Scanned $r.Log) '1'
    } else {
        Write-Skip 'UNC のパス' '管理共有を読めないため'
    }
    $start = "$workRoot\start"
    $startCases = @(@('隣に一致するフォルダ', '[a]', @('a')), @('一致するフォルダなし', 'b[1]', @()), @('パターンとして正しくない', 'c[d', @()), @('2 つ以上一致', '[ef]', @('e', 'f')))
    foreach ($case in $startCases) {
        $dir = "$start\$($case[1])"
        New-DataFile "$dir\p.bin" 3000 140
        New-DataFile "$dir\q.bin" 3000 140
        # 間違ったフォルダを調べたら分かるように、隣のフォルダには別のファイルを置く
        foreach ($decoy in $case[2]) { New-DataFile "$start\$decoy\wrong.bin" 3000 141 }
        $r = Invoke-From $dir @('-MinSize', '1KB', '-OutputFolder', 'out')
        $x = @(Import-Csv -LiteralPath (Join-Path (Get-RunDir "$dir\out") 'duplicates.csv') -Encoding UTF8)
        Assert-Equal ('[ ] を含むフォルダから -Path の省略と -OutputFolder: {0}' -f $case[0]) ('{0}/{1}' -f $r.Code, (Get-RelText $x $dir)) '2/p.bin|q.bin'
    }

    Write-Section '設定ファイルの細かい動き'
    $cfgDir = "$workRoot\cfg"
    New-DataFile "$cfgDir\data\a.bin" 3000 150
    New-DataFile "$cfgDir\data\b.bin" 3000 150
    $utf8 = New-Object Text.UTF8Encoding $false
    [IO.File]::WriteAllText("$cfgDir\rel.ini", "Path = data`r`nMinSize = 1KB`r`nOutputFolder = 記録`r`n", $utf8)
    $r = Invoke-From $workRoot @('-ConfigFile', "$cfgDir\rel.ini")
    Assert-Equal '相対パスは設定ファイルのあるフォルダが基準' ('{0}/{1}' -f $r.Code, [bool](Get-RunDir "$cfgDir\記録")) '2/True'
    New-DataFile "$cfgDir\日本語\a.bin" 3000 151
    New-DataFile "$cfgDir\日本語\b.bin" 3000 151
    [IO.File]::WriteAllText("$cfgDir\sjis.ini", "# 日本語のコメント`r`nPath = 日本語`r`nMinSize = 1KB`r`nNoSave = true`r`n", [Text.Encoding]::GetEncoding(932))
    $r = Invoke-From $workRoot @('-ConfigFile', "$cfgDir\sjis.ini")
    Assert-Equal 'Shift_JIS の設定ファイルと NoSave = true' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '保存先')) '2/False'
    [IO.File]::WriteAllText("$cfgDir\bom.ini", "Path = `"data`"`r`nMinSize = '1KB'`r`nNoSave = true`r`n", (New-Object Text.UTF8Encoding $true))
    Assert-Equal 'BOM 付きと引用符' (Invoke-From $workRoot @('-ConfigFile', "$cfgDir\bom.ini")).Code 2
    [IO.File]::WriteAllText("$cfgDir\env.ini", "Path = %WB_FDF_TEST%\data`r`nMinSize = 1KB`r`nNoSave = true`r`n", $utf8)
    $env:WB_FDF_TEST = $cfgDir
    try { Assert-Equal '環境変数を展開する' (Invoke-From $workRoot @('-ConfigFile', "$cfgDir\env.ini")).Code 2 } finally { Remove-Item Env:\WB_FDF_TEST -ErrorAction SilentlyContinue }
    [IO.File]::WriteAllText("$cfgDir\cli.ini", "Path = data`r`nMinSize = 1MB`r`nTop = 1`r`nNoSave = true`r`n", $utf8)
    $r = Invoke-From $workRoot @('-ConfigFile', "$cfgDir\cli.ini", '-MinSize', '1KB', '-Top', '5')
    Assert-Equal 'コマンドラインの値を優先する' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '（上位 5）')) '2/True'
    $r = Invoke-NoSave @('-ConfigFile', "$cfgDir\nothing.ini")
    Assert-Equal 'ない設定ファイル' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '設定ファイルが見つかりません')) '1/True'
    [IO.File]::WriteAllText("$cfgDir\memo.ini.txt", 'Top = 1', $utf8)
    $r = Invoke-NoSave @('-ConfigFile', "$cfgDir\memo.ini")
    Assert-Equal '拡張子が .txt のときの案内' (Test-Contains $r.Log '拡張子が .txt になっていませんか') 'True'
    [IO.File]::WriteAllText("$cfgDir\badpath.ini", "Path = no_such_folder`r`n", $utf8)
    $r = Invoke-NoSave @('-ConfigFile', "$cfgDir\badpath.ini")
    Assert-Equal 'Path が見つからないときは行番号を添える' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '（設定ファイルの 1 行目）')) '1/True'

    Write-Section '保存'
    # 調べるフォルダの中に保存しても、結果のフォルダは調べない
    $selfDir = "$workRoot\selfout"
    New-DataFile "$selfDir\a.bin" 3000 160
    New-DataFile "$selfDir\b.bin" 3000 160
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $selfDir -MinSize 1KB -OutputFolder $selfDir 2>&1
    $firstRun = Get-RunDir $selfDir
    Assert-Equal '今回の結果のフォルダは skipped に残さない' (Get-SkipCount (Join-Path $firstRun 'skipped.csv')) 0
    Start-Sleep -Milliseconds 1100
    $out = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $selfDir -MinSize 1KB -OutputFolder $selfDir 2>&1 | Out-String)
    $runs = @(Get-ChildItem -LiteralPath $selfDir -Directory | Where-Object { $_.Name -like 'DuplicateFile_*' } | Sort-Object Name)
    Assert-Equal '前回の結果のフォルダは調べず、記録する' ('{0}/{1}' -f $runs.Count, (Get-SkipCount (Join-Path $runs[-1].FullName 'skipped.csv'))) '2/1'
    Assert-Equal '結果のフォルダの中身は数えない' (Get-Scanned $out) '2'
    # 既定の保存先はスクリプトのフォルダ
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$workRoot\m1,$workRoot\m2" -MinSize 1KB 2>&1
    Assert-Equal '既定の保存先はスクリプトのフォルダ' ([bool](Get-RunDir (Split-Path -Parent $tool))) 'True'
    Assert-Equal '保存先のフォルダの名前' ((Split-Path -Leaf (Get-RunDir (Split-Path -Parent $tool))) -replace '\d{8}_\d{6}$', 'DATE') ('DuplicateFile_{0}_DATE' -f $env:COMPUTERNAME)
    # 同じ秒に実行しても、前の結果を消さない（Ctrl+C や保存の失敗のときはフォルダごと消すため）
    $sameDir = "$workRoot\out_same"
    [void][IO.Directory]::CreateDirectory($sameDir)
    $stampNow = Get-Date
    foreach ($s in 0..40) {
        $older = Join-Path $sameDir ('DuplicateFile_{0}_{1}' -f $env:COMPUTERNAME, $stampNow.AddSeconds($s).ToString('yyyyMMdd_HHmmss'))
        [void][IO.Directory]::CreateDirectory($older)
        [IO.File]::WriteAllText("$older\keep.txt", 'keep')
    }
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$workRoot\m1" -OutputFolder $sameDir 2>&1
    Assert-Equal '同じ名前のフォルダがあれば _2 を付ける' (@(Get-ChildItem -LiteralPath $sameDir -Directory -Filter 'DuplicateFile_*_2').Count) 1
    Assert-Equal '前の結果を残す' (@(Get-ChildItem -LiteralPath $sameDir -Recurse -Filter 'keep.txt').Count) 41
    $longOut = "$workRoot\" + ('O' * 120) + '\' + ('P' * 120)
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$workRoot\m1,$workRoot\m2" -MinSize 1KB -OutputFolder $longOut 2>&1
    $longRuns = @(); try { $longRuns = @([IO.Directory]::GetDirectories('\\?\' + $longOut)) } catch { }
    Assert-Equal '-OutputFolder が長くても保存する' ('{0}/{1}' -f $LASTEXITCODE, $longRuns.Count) '2/1'
    if ($longRuns.Count -eq 1) { Assert-Equal '長い保存先のファイル' (@([IO.Directory]::GetFiles($longRuns[0])).Count) 5 }
    $spaceOut = "$workRoot\out 日本語 [x]"
    $r = Invoke-NoSaveless @('-Path', "$workRoot\m1,$workRoot\m2", '-MinSize', '1KB', '-OutputFolder', $spaceOut)
    Assert-Equal '保存先に日本語・スペース・[ ]' ('{0}/{1}' -f $r.Code, [bool](Get-RunDir $spaceOut)) '2/True'
    $letter = Get-FreeLetter
    if ($letter) {
        $r = Invoke-NoSaveless @('-Path', "$workRoot\m1,$workRoot\m2", '-MinSize', '1KB', '-OutputFolder', "${letter}:\nope")
        Assert-Equal 'ないドライブ: 警告を出して画面に表示' ('{0}/{1}/{2}' -f $r.Code, (Test-Contains $r.Log '保存先に書き込めないため'), (Test-Contains $r.Log '重複 1 グループ')) '2/True/True'
    } else {
        Write-Skip 'ないドライブ' '空いているドライブ文字がないため'
    }
    $patched = New-PatchedTool 'savefail' "Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'duplicates.csv'))" "throw 'disk full'; Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'duplicates.csv'))"
    $r = Invoke-Tool 'savefail' @('-Path', $view, '-MinSize', '1KB') $patched
    Assert-Equal '保存に失敗: 警告を出して書きかけを残さない' ('{0}/{1}/{2}' -f $r.Code, (Test-Contains $r.Log '結果を保存できませんでした: disk full'), [bool]$r.Dir) '2/True/False'
    Assert-Equal '保存に失敗: 画面には表を出す' ('{0}/{1}' -f (Test-Contains $r.Log ' 1. 減らせる容量'), (Test-Contains $r.Log '保存先:')) 'True/False'
    $patched = New-PatchedTool 'crash' '$groups = @($groupsRaw | Sort-Object' "throw 'boom'; `$groups = @(`$groupsRaw | Sort-Object"
    $r = Invoke-Tool 'crash' @('-Path', $view, '-MinSize', '1KB') $patched
    Assert-Equal '想定外のエラー: 結果を残さず、中断とは区別する' ('{0}/{1}/{2}/{3}' -f $r.Code, (Test-Contains $r.Log 'エラーで止まりました。結果は保存していません。'), (Test-Contains $r.Log '中断しました'), [bool]$r.Dir) '1/True/False/False'

    Write-Section 'コマンドプロンプトから起動'
    $cmdHead = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $r = Invoke-ViaCmd ('{0} -NoSave -MinSize 1KB -Path "{1},{2}"' -f $cmdHead, "$workRoot\m1", "$workRoot\m2")
    Assert-Equal 'カンマ区切りの -Path と終了コード 2' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '重複 1 グループ')) '2/True'
    $r = Invoke-ViaCmd ("{0} -NoSave -MinSize 1KB -Path '{1}' -Include '*.bin,*.txt' -Exclude 'logs,keep.bin'" -f $cmdHead, $ex2)
    Assert-Equal "cmd の ' はそのまま届く" ('{0}/{1}/{2}' -f $r.Code, (Test-Contains $r.Log '絞り込み: *.bin, *.txt'), (Test-Contains $r.Log '除外: logs, keep.bin')) '2/True/True'
    $r = Invoke-ViaCmd ('{0} -NoSave -MinSize 1KB -Path "{1}\" -Top 5' -f $cmdHead, "$workRoot\m1")
    Assert-Equal '"パス\" の後ろに引数（止まってヒントを出す）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '最後に \ を付けないでください')) '1/True'
    $r = Invoke-ViaCmd ('{0} -NoSave -MinSize 1KB -Path "{1}\"' -f $cmdHead, "$workRoot\m1")
    Assert-Equal '最後の引数が \"（引用符を外して調べる）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log "対象: $workRoot\m1")) '0/True'
    $r = Invoke-ViaCmd ('{0} -Path "{1}" -OutputFolder "{2}\" -Top 5' -f $cmdHead, "$workRoot\m1", "$workRoot\out_q")
    Assert-Equal '-OutputFolder の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '-OutputFolder の値に引用符')) '1/True'
    $r = Invoke-ViaCmd ('{0} -NoSave -ConfigFile "{1}\" -Top 5' -f $cmdHead, "$cfgDir\rel.ini")
    Assert-Equal '-ConfigFile の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '-ConfigFile の値に引用符')) '1/True'
    $r = Invoke-ViaCmd ('{0} -NoSave -Path "{1}" -Exclude "{2}\" -Top 5' -f $cmdHead, "$workRoot\m1", 'logs')
    Assert-Equal '-Exclude の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $r.Code, (Test-Contains $r.Log '-Exclude の値に引用符')) '1/True'

    Write-Section 'Ctrl+C で中断'
    # 比べ始める前に長く待つコピーで止める（読むだけなので、途中で止めても何も壊れない）
    $slowTool = New-PatchedTool 'slow' '$fullGroups = New-Object ''System.Collections.Generic.List[object]''' 'Start-Sleep -Seconds 60; $fullGroups = New-Object ''System.Collections.Generic.List[object]'''
    $stopFile = "$workRoot\stop-during-scan.ps1"
    [IO.File]::WriteAllLines($stopFile, [string[]]@(
        'param($Tool, $Target, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $Target).AddParameter(''MinSize'', ''1KB'').AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'Start-Sleep -Milliseconds 3000'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
    ), (New-Object Text.UTF8Encoding $true))
    $stopResult = @{}
    & $Exe -NoProfile -ExecutionPolicy Bypass -File $stopFile -Tool $slowTool -Target $view -OutRoot "$workRoot\out_ctrlc" 2>&1 | ForEach-Object {
        $line = "$_"
        $eq = $line.IndexOf('=')
        if ($eq -gt 0) { $stopResult[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
    }
    if ($stopResult['RUNNING'] -eq 'True') {
        Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $stopResult['BEFORE'], $stopResult['AFTER']) '1/0'
        Assert-Equal '中断のメッセージ' $stopResult['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' '3 秒より前に終わったため'
    }

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $cultureProbe = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $cultureProbe)"
    if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
        $cultureOut = "$workRoot\out_culture"
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -MinSize 1KB -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $view.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $cultureDir = Get-RunDir $cultureOut
        $summaryText = if ($cultureDir) { [IO.File]::ReadAllText((Join-Path $cultureDir 'summary.txt')) } else { '' }
        Assert-Equal 'fi-FI でもサイズの小数点は .' (Test-Contains $summaryText '15.2 KB') 'True'
        Assert-Equal 'fi-FI でも日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($summaryText, '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   ')) 'True'
        Assert-Equal 'fi-FI でも所要時間の小数点は .' ([regex]::IsMatch($summaryText, '所要 \d+\.\d秒')) 'True'
        $row = if ($cultureDir) { @(Import-Csv -LiteralPath (Join-Path $cultureDir 'duplicates.csv') -Encoding UTF8)[0] } else { $null }
        Assert-Equal 'fi-FI でも CSV の日時とサイズ' $(if ($row) { '{0}/{1}' -f $row.'最終更新日時', $row.'サイズ(バイト)' } else { '' }) '2026/10/01 10:00:00/3,000'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }
} finally {
    # アクセス拒否を戻してから消す
    foreach ($target in $denied) { $null = & icacls.exe $target /remove:d "*$userSid" 2>&1 }
    if ($substDrive) { $null = & subst.exe $substDrive /D 2>&1 }
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

'  END'
exit $script:failCount
