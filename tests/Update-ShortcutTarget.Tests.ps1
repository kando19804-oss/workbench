<#
.SYNOPSIS
    Update-ShortcutTarget の回帰テストです。

.DESCRIPTION
    一時フォルダにダミーのショートカット（共有フォルダ・サーバーそのもの・引数やコメントにパスがあるもの、読み取り専用・隠し・
    書き込み禁止、名前に [ ] や日本語・空白を含むもの、壊れたもの、長すぎる項目のあるもの、260 文字を超えるパスのものなど）を作り、
    Update-ShortcutTarget の棚卸し・一覧・置き換えの結果を期待値と照合します。置き換えたショートカットは、ファイルの中身を読んで確かめます。

    共有フォルダには、この PC のループバックの管理共有（\\127.0.0.1\C$ を置き換える前、\\localhost\C$ を置き換えた後）を使います。
    外へは通信しません。管理共有を開けない環境では、置き換えのテストを飛ばします（SKIP）。
    つながらないサーバーには、すぐに「届かない」が返る 127.0.0.2 を使います。

    「権限なし」を再現するため、テストが作ったフォルダとファイルにだけアクセス拒否を設定します（終了時に戻します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    注意: このテストは、管理共有を通して数百個のショートカットを作り・書き換え・消し、PowerShell や cmd を数十回起動します。
    ウイルス対策ソフト（Trend Micro）の入った Windows 11（Insider Preview）で、このテストの最中に、ウイルス対策ソフトの
    ドライバーがカーネルで例外を起こしてブルースクリーンになったことがあります（2026-11-13。ツールを使うだけでは起きていません）。
    流す前に、作業中のファイルを保存してください。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Update-ShortcutTarget を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Update-ShortcutTarget\Update-ShortcutTarget.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Update-ShortcutTarget_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Update-ShortcutTarget.ps1'
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
$latin1    = [Text.Encoding]::GetEncoding(28591)
$isAdmin   = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$shortcutHead = '"ショートカット","サーバー","リンク先","引数","作業フォルダ","アイコンの場所","コメント","最終更新日時","結果","備考","控え"'
$changeHead   = '"ショートカット","項目","置き換える前","置き換えた後","結果"'
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

function Join-Sorted($Items) {
    $array = [string[]]@($Items | Where-Object { $null -ne $_ })
    [Array]::Sort($array, [StringComparer]::Ordinal)
    $array -join ','
}

# ローカルのパスを、指定したサーバーの管理共有のパスにする（C:\a → \\127.0.0.1\C$\a）
function ConvertTo-AdminShare([string]$Server, [string]$LocalPath) {
    '\\' + $Server + '\' + $LocalPath.Substring(0, 1).ToUpperInvariant() + '$\' + $LocalPath.Substring(3)
}

# ショートカットを作る（リンク先を最初に設定し、ほかの項目はその後）
function New-Link([string]$File, [string]$Target, [hashtable]$Props = @{}) {
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $File))
    $link = $wsh.CreateShortcut($File)
    $link.TargetPath = $Target
    foreach ($key in $Props.Keys) { $link.$key = $Props[$key] }
    $link.Save()
}

# 260 文字を超えるパスにも書けるように、\\?\ を付けて書く
function New-TestFile([string]$FilePath, [byte[]]$Bytes) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllBytes('\\?\' + $FilePath, $Bytes)
}

# 書き換えた直後のショートカットは、ウイルス対策ソフトなどが削除の権限付きで開いていることがあるため、書き込みと削除も共有して読む
function Get-Bytes([string]$FilePath) {
    $stream = New-Object IO.FileStream(('\\?\' + $FilePath), [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]'ReadWrite, Delete'))
    try {
        $buffer = New-Object 'byte[]' $stream.Length
        $count = 0
        while ($count -lt $buffer.Length) { $n = $stream.Read($buffer, $count, $buffer.Length - $count); if ($n -le 0) { break }; $count += $n }
        , $buffer
    } finally {
        $stream.Dispose()
    }
}

function Get-Hash([string]$FilePath) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { [BitConverter]::ToString($sha.ComputeHash((Get-Bytes $FilePath))) } finally { $sha.Dispose() }
}

# フォルダの中の .lnk を、リンクと読めないフォルダを飛ばしてすべて集める（相対パス → ハッシュ）
function Get-LinkHashes([string]$Dir) {
    $result = @{}
    $stack = New-Object 'System.Collections.Generic.Stack[string]'
    $stack.Push($Dir)
    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try { $entries = (New-Object IO.DirectoryInfo ('\\?\' + $current)).GetFileSystemInfos() } catch { continue }
        foreach ($entry in $entries) {
            $full = $entry.FullName.Substring(4)
            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
            if ($entry -is [IO.DirectoryInfo]) { $stack.Push($full); continue }
            if ($entry.Name.EndsWith('.lnk', [StringComparison]::OrdinalIgnoreCase)) {
                try { $result[$full.Substring($Dir.Length + 1)] = Get-Hash $full } catch { }
            }
        }
    }
    $result
}

# 2 つのハッシュの一覧で、違うもの（相対パス）
function Get-ChangedNames([hashtable]$Before, [hashtable]$After, [string[]]$Names = $null) {
    $keys = if ($Names) { $Names } else { @($Before.Keys) }
    Join-Sorted ($keys | Where-Object { $Before[$_] -ne $After[$_] })
}

# ショートカットのファイルから、文字の項目と、環境変数の形のリンク先を読む（ツールとは別に書く）
function Read-Link([string]$File) {
    $b = Get-Bytes $File
    $flags = [BitConverter]::ToUInt32($b, 0x14)
    $pos = 0x4C
    if ($flags -band 1) { $pos += 2 + [BitConverter]::ToUInt16($b, $pos) }
    if ($flags -band 2) { $pos += [BitConverter]::ToUInt32($b, $pos) }
    $r = @{ Comment = ''; Relative = ''; WorkDir = ''; Arguments = ''; IconPath = ''; EnvTarget = ''; Wsh = '' }
    foreach ($field in @(@(0x4, 'Comment'), @(0x8, 'Relative'), @(0x10, 'WorkDir'), @(0x20, 'Arguments'), @(0x40, 'IconPath'))) {
        if ($flags -band $field[0]) {
            $count = [BitConverter]::ToUInt16($b, $pos)
            $r[$field[1]] = [Text.Encoding]::Unicode.GetString($b, $pos + 2, 2 * $count)
            $pos += 2 + 2 * $count
        }
    }
    $r.IconIndex = [BitConverter]::ToInt32($b, 0x38)
    while ($pos + 8 -le $b.Length) {
        $size = [BitConverter]::ToUInt32($b, $pos)
        if ($size -lt 8) { break }
        if ([BitConverter]::ToInt32($b, $pos + 4) -eq 0xA0000001) {
            $text = [Text.Encoding]::Unicode.GetString($b, $pos + 268, 520).Split([char]0)[0]
            if (-not $text) { $text = [Text.Encoding]::Default.GetString($b, $pos + 8, 260).Split([char]0)[0] }
            $r.EnvTarget = $text
        }
        $pos += $size
    }
    if ($File.Length -lt 260) { $r.Wsh = [string]$wsh.CreateShortcut($File).TargetPath }
    $r.Attr = [string][IO.File]::GetAttributes('\\?\' + $File)
    [pscustomobject]$r
}

# ファイルの中に、文字が ANSI か UTF-16 で何回あるか（大文字と小文字は区別しない）
function Get-TextCount([string]$File, [string]$Text) {
    $raw = $latin1.GetString((Get-Bytes $File))
    $count = 0
    foreach ($needle in @($latin1.GetString([Text.Encoding]::ASCII.GetBytes($Text)), $latin1.GetString([Text.Encoding]::Unicode.GetBytes($Text)))) {
        $i = 0
        while (($i = $raw.IndexOf($needle, $i, [StringComparison]::OrdinalIgnoreCase)) -ge 0) { $count++; $i++ }
    }
    $count
}

function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$InputText = $null, [string]$ToolPath = $tool) {
    $out = Join-Path $workRoot "out\$Label"
    if ($null -ne $InputText) {
        $lines = @($InputText | & $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }
    } else {
        $lines = @('' | & $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }
    }
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = (ConvertTo-FlatText (@($lines) -join "`n")); Lines = @($lines); Dir = (Get-RunDir $out) }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行。確認を待つときに止まらないよう、標準入力は空にする
function Invoke-Plain([string[]]$ToolArgs) {
    $log = ConvertTo-FlatText ((@('' | & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
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
    $file = Join-Path $workRoot "patched\$Name\Update-ShortcutTarget.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text.Replace($From, $To), $utf8Bom)
    $file
}

# 保存先のフォルダの下に作られた、結果のフォルダ（260 文字を超えても探せるように \\?\ を付ける）
function Get-RunDir([string]$Parent) {
    if (-not [IO.Directory]::Exists('\\?\' + $Parent)) { return $null }
    $found = @([IO.Directory]::GetDirectories('\\?\' + $Parent, 'ShortcutTarget_*') | Sort-Object)
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
    [pscustomobject]@{ Shortcuts = @(Read-CsvOrEmpty 'shortcuts.csv'); Changes = @(Read-CsvOrEmpty 'changes.csv'); Skipped = @(Read-CsvOrEmpty 'skipped.csv'); Lines = $lines }
}

# 保存されたファイルとフォルダの名前の一覧
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

# shortcuts.csv の、ダミーのフォルダから見たショートカットの行の列
function Get-Row($Rows, [string]$Base, [string]$Rel, [string]$Column) {
    $full = Join-Path $Base $Rel
    (@($Rows | Where-Object { $_.'ショートカット' -eq $full }) | ForEach-Object { $_.$Column }) -join '|'
}

# shortcuts.csv の、ショートカットの名前と結果の一覧（名前の順）
function Get-ResultList($Rows) {
    Join-Sorted ($Rows | ForEach-Object { '{0}={1}' -f (Split-Path -Leaf $_.'ショートカット'), $_.'結果' })
}

# 自分自身にアクセス拒否を設定する。終了時に戻すため、設定したパスを覚えておく
function Add-Deny([string]$Path, [string]$Rights) {
    $null = & icacls.exe $Path /deny "*${userSid}:$Rights" 2>&1
    if ($LASTEXITCODE -ne 0) { return $false }
    $denied.Add($Path)
    $true
}

# 主な場面のダミー。どれも置き換える前のサーバー（127.0.0.1）か、ほかのサーバーを指す
function Build-Main([string]$Dir) {
    New-Link "$Dir\a-folder.lnk" "$O\data\dept" @{ WorkingDirectory = "$O\data" }
    New-Link "$Dir\b-file.lnk" "$O\data\dept\report.xlsx" @{ IconLocation = "$O\data\dept\app.ico,2"; Description = "場所: $O\data" }
    New-Link "$Dir\c-missing.lnk" "$O\data\missing"
    New-Link "$Dir\d-explorer.lnk" 'C:\Windows\explorer.exe' @{ Arguments = "`"$O\data\dept`" /select" }
    # 127.0.0.10 と 127.0.0.1x は、127.0.0.1 とは別のサーバー
    New-Link "$Dir\e-other.lnk" 'C:\Windows\notepad.exe' @{ Arguments = '\\127.0.0.10\x\y.txt \\127.0.0.1x\z' }
    New-Link "$Dir\f-local.lnk" 'C:\Windows\notepad.exe'
    New-Link "$Dir\g-readonly.lnk" "$O\data\dept"
    New-Link "$Dir\h-hidden.lnk" "$O\data\dept"
    New-Link "$Dir\j-server.lnk" '\\127.0.0.1'
    New-Link "$Dir\k-indirect.lnk" "$O\data\dept" @{ Description = '@%SystemRoot%\system32\shell32.dll,-22579' }
    New-Link "$Dir\l-pipe.lnk" 'C:\Windows\explorer.exe' @{ Arguments = "\\.\pipe\x /e,$O\data\dept" }
    New-Link "$Dir\m-near.lnk" 'C:\Windows\notepad.exe' @{ WorkingDirectory = '\\127.0.0.10\s' }
    New-Link "$Dir\n-noshare.lnk" '\\127.0.0.1\wbnoshare\x'
    New-Link "$Dir\t-denied.lnk" "$O\data\denied\f.txt"
    New-Link "$Dir\w-nowrite.lnk" "$O\data\dept"
    New-Link "$Dir\sub\[a]\i-bracket.lnk" "$O\data\dept"
    New-Link "$Dir\日本語 フォルダ\営業 共有.lnk" "$O\data\dept"
    New-Link "$Dir\ShortcutTarget_memo\keep.lnk" "$O\data\dept"
    New-Link "$Dir\ShortcutTarget_PC01_20260101_000000\backup\0001_old.lnk" "$O\data\dept"
    New-Link "$Dir\`$RECYCLE.BIN\r.lnk" "$O\data\dept"
    New-Link "$Dir\noread\x.lnk" "$O\data\dept"
    New-Link "$Dir\long-comment.lnk" "$O\data\dept" @{ Description = ('d' * 295) + 'TAIL!' }
    New-Link "$Dir\long-wd.lnk" "$O\data\dept" @{ WorkingDirectory = 'C:\' + ('w' * 258) }
    New-TestFile "$Dir\z-corrupt.lnk" ([byte[]](1..100))
    New-TestFile "$Dir\zero.lnk" ([byte[]]@())
    $folderBytes = Get-Bytes "$Dir\a-folder.lnk"
    $big = New-Object 'byte[]' (1MB + 1)
    [Array]::Copy($folderBytes, $big, $folderBytes.Length)
    New-TestFile "$Dir\big.lnk" $big
    New-TestFile "$Dir\$longRel" $folderBytes
    New-TestFile "$Dir\memo.txt" $utf8NoBom.GetBytes("ショートカットではない`n")
    [IO.File]::SetAttributes("$Dir\g-readonly.lnk", [IO.FileAttributes]::ReadOnly)
    [IO.File]::SetAttributes("$Dir\h-hidden.lnk", [IO.FileAttributes]::Hidden)
    $null = & cmd.exe /d /c mklink /J "$Dir\jloop" "$Dir" 2>&1
    $script:junctions.Add("$Dir\jloop")
    # (W) には同期（SYNCHRONIZE）も含まれ、読むこともできなくなるため、書き込みだけを拒否する
    $ok = (Add-Deny "$Dir\noread" '(RD)') -and (Add-Deny "$Dir\w-nowrite.lnk" '(WD,AD)')
    $ok -and [IO.Directory]::Exists("$Dir\jloop")
}

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
$denied    = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したパス（終了時に戻す）
$junctions = New-Object 'System.Collections.Generic.List[string]'   # 作ったジャンクション（終了時に先に外す）
$wsh       = New-Object -ComObject WScript.Shell
$longRel   = ('L' * 120) + '\' + ('M' * 120) + '\long.lnk'
$O = ConvertTo-AdminShare '127.0.0.1' $workRoot   # 置き換える前
$N = ConvertTo-AdminShare 'localhost' $workRoot   # 置き換えた後
$main  = Join-Path $workRoot 'main'

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool
    foreach ($dir in 'data\dept', 'data\denied', 'data2\x', 'lim\longer') { [void][IO.Directory]::CreateDirectory((Join-Path $workRoot $dir)) }
    [IO.File]::WriteAllText((Join-Path $workRoot 'data\dept\report.xlsx'), 'x')
    [IO.File]::WriteAllText((Join-Path $workRoot 'data\denied\f.txt'), 'x')
    $hasShare = [IO.Directory]::Exists("$O\data\dept") -and [IO.Directory]::Exists("$N\data\dept")
    if (-not $hasShare) { Write-Skip '管理共有' "$O と $N を開けないため、置き換えのテストを飛ばします" }
    $mainOk = Build-Main $main
    Assert-Equal 'ジャンクションとアクセス拒否を作れた' $mainOk 'True'
    $mainHashes = Get-LinkHashes $main

    Write-Section '棚卸し（-From なし）'
    $run = Invoke-Tool 'inventory' @('-Path', $main)
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '保存するファイル（changes.csv と backup は作らない）' (Get-ResultNames $run.Dir) 'shortcuts.csv,skipped.csv,summary.txt'
    Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'shortcuts.csv=True,skipped.csv=True,summary.txt=True'
    Assert-Equal 'shortcuts.csv の見出し' (Get-FirstLine $run.Dir 'shortcuts.csv') $shortcutHead
    Assert-Equal 'skipped.csv の見出し' (Get-FirstLine $run.Dir 'skipped.csv') '"理由","パス","詳細"'
    $order = @($x.Shortcuts | ForEach-Object { $_.'ショートカット'.Substring($main.Length + 1) })
    $expectedOrder = 'a-folder.lnk', 'b-file.lnk', 'c-missing.lnk', 'd-explorer.lnk', 'g-readonly.lnk', 'h-hidden.lnk', 'j-server.lnk', 'k-indirect.lnk', 'l-pipe.lnk', $longRel, 'n-noshare.lnk', 'ShortcutTarget_memo\keep.lnk', 'sub\[a]\i-bracket.lnk', 't-denied.lnk', 'w-nowrite.lnk', '日本語 フォルダ\営業 共有.lnk', 'e-other.lnk', 'm-near.lnk'
    Assert-Equal '一覧の並び（サーバー・パスの順。文字の番号の順）' ($order -join ',') ($expectedOrder -join ',')
    Assert-Equal 'サーバー（IP アドレスは ="..." で囲む）' ('{0}|{1}|{2}' -f (Get-Row $x.Shortcuts $main 'a-folder.lnk' 'サーバー'), (Get-Row $x.Shortcuts $main 'e-other.lnk' 'サーバー'), (Get-Row $x.Shortcuts $main 'm-near.lnk' 'サーバー')) '="127.0.0.1"|127.0.0.10, 127.0.0.1x|="127.0.0.10"'
    Assert-Equal 'アイコンの場所とコメント（ファイルにある値のまま。@ で始まる値は Excel の数式にならないよう ="..." で囲む）' ('{0}|{1}|{2}' -f (Get-Row $x.Shortcuts $main 'b-file.lnk' 'アイコンの場所'), (Get-Row $x.Shortcuts $main 'b-file.lnk' 'コメント'), (Get-Row $x.Shortcuts $main 'k-indirect.lnk' 'コメント')) "$O\data\dept\app.ico,2|場所: $O\data|=`"@%SystemRoot%\system32\shell32.dll,-22579`""
    Assert-Equal '引数と作業フォルダ' ('{0}|{1}' -f (Get-Row $x.Shortcuts $main 'd-explorer.lnk' '引数'), (Get-Row $x.Shortcuts $main 'a-folder.lnk' '作業フォルダ')) "`"$O\data\dept`" /select|$O\data"
    Assert-Equal 'サーバーそのものを指すショートカット（リンク先はファイルから読む）' (Get-Row $x.Shortcuts $main 'j-server.lnk' 'リンク先') '\\127.0.0.1'
    Assert-Equal '最終更新日時の形' ((Get-Row $x.Shortcuts $main 'a-folder.lnk' '最終更新日時') -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
    Assert-Equal '棚卸しのときは結果・備考・控えが空' (@($x.Shortcuts | Where-Object { $_.'結果' -or $_.'備考' -or $_.'控え' }).Count) 0
    $skipText = Join-Sorted ($x.Skipped | ForEach-Object { '{0}:{1}' -f $_.'理由', $_.'パス'.Substring($main.Length + 1) })
    Assert-Equal 'skipped.csv' $skipText (Join-Sorted @('権限なし:noread', '除外指定:$RECYCLE.BIN', '除外指定:ShortcutTarget_PC01_20260101_000000', '読めないショートカット:big.lnk', '読めないショートカット:long-comment.lnk', '読めないショートカット:long-wd.lnk', '読めないショートカット:z-corrupt.lnk', '読めないショートカット:zero.lnk', 'リンク:jloop'))
    $details = @{}
    foreach ($row in $x.Skipped) { $details[$row.'パス'.Substring($main.Length + 1)] = $row.'詳細' }
    Assert-Equal '読めないショートカットの詳細' ('{0}|{1}|{2}|{3}|{4}' -f $details['z-corrupt.lnk'], $details['zero.lnk'], $details['big.lnk'], $details['long-comment.lnk'], $details['long-wd.lnk']) 'ショートカットの形式ではありません|ショートカットの形式ではありません|1 MB を超えるため、ショートカットとして読んでいません|コメントが 260 文字を超えるため、Windows の部品（WScript.Shell）で正しく読めません|作業フォルダが 260 文字を超えるため、Windows の部品（WScript.Shell）で正しく読めません'
    Assert-Equal '除外とリンクの詳細' ('{0}|{1}|{2}' -f $details['ShortcutTarget_PC01_20260101_000000'], $details['$RECYCLE.BIN'], $details['jloop']) 'このツールの結果のフォルダ|ごみ箱・システムの管理用のフォルダ|ジャンクション'
    $summary = $x.Lines -join "`n"
    Assert-Equal 'summary の 1 行目' (@($x.Lines) + '')[0] 'ショートカットのリンク先置き換え (Update-ShortcutTarget 1.0.0)'
    Assert-Equal 'summary の対象と日時の行' ('{0}|{1}' -f (@($x.Lines) + '' + '')[1], [regex]::IsMatch((@($x.Lines) + '' + '' + '')[2], ('^\d{{4}}-\d{{2}}-\d{{2}} \d{{2}}:\d{{2}}:\d{{2}}   {0}（管理者: (はい|いいえ)）$' -f [regex]::Escape($env:COMPUTERNAME)))) "対象: $main|True"
    Assert-Equal '件数の行' ([regex]::IsMatch($summary, '\nショートカット 19 件  共有フォルダを指すもの 18 件  所要 \d+\.\d秒\nスキップ 6 件（権限なし 1・読めないショートカット 5）  たどらなかったリンク 1 件  除外 2 件\n')) 'True'
    Assert-Equal 'サーバーごとの件数（件数の多い順）' ([regex]::IsMatch($summary, '■ サーバーごとの件数\n サーバー\s+件数\n 127\.0\.0\.1\s+16\n 127\.0\.0\.10\s+2\n 127\.0\.0\.1x\s+1\n')) 'True'
    Assert-Equal '一覧の見出しと、引数・作業フォルダで見つけたものの表示' ('{0}/{1}/{2}/{3}' -f [regex]::IsMatch($summary, '■ 共有フォルダを指すショートカット（サーバーの順）\n サーバー\s+共有フォルダのパス\s+ショートカット\n'), $summary.Contains("（引数）$O\data\dept"), $summary.Contains('（引数）\\127.0.0.10\x\y.txt'), $summary.Contains('（作業フォルダ）\\127.0.0.10\s')) 'True/True/True/True'
    Assert-Equal '\\.\pipe はサーバーとして扱わない' ('{0}/{1}' -f $summary.Contains('\\.\pipe'), [regex]::IsMatch($summary, '\n \.\s')) 'False/False'
    Assert-Equal 'summary の最後は保存先' (@($x.Lines) + '')[-2] ('保存先: {0}' -f $run.Dir)
    Assert-Equal '棚卸しでは、ショートカットを変えない' (Get-ChangedNames $mainHashes (Get-LinkHashes $main)) ''

    Write-Section '一覧（-From だけ）'
    $run = Invoke-Plain @('-Path', $main, '-From', '127.0.0.1', '-NoSave')
    Assert-Equal '終了コード・件数の行・見出し' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('ショートカット 19 件  127.0.0.1 を指すもの 16 件  所要'), $run.Log.Contains('■ 127.0.0.1 を指すショートカット リンク先')) '0/True/True'
    Assert-Equal '127.0.0.10・127.0.0.1x は 127.0.0.1 に当たらない' ('{0}/{1}' -f $run.Log.Contains('e-other.lnk'), $run.Log.Contains('m-near.lnk')) 'False/False'
    $run = Invoke-Plain @('-Path', $main, '-From', '\\127.0.0.1\', '-NoSave')
    Assert-Equal '-From の \\ と末尾の \ は外して読む' ('{0}/{1}' -f $run.Code, $run.Log.Contains('127.0.0.1 を指すもの 16 件')) '0/True'
    $run = Invoke-Plain @('-Path', $main, '-From', '127.0.0.1, 127.0.0.10', '-NoSave')
    Assert-Equal '-From をカンマで複数' ('{0}/{1}' -f $run.Code, $run.Log.Contains('127.0.0.1, 127.0.0.10 を指すもの 18 件')) '0/True'
    $run = Invoke-Plain @('-Path', $main, '-From', "\\127.0.0.1\C`$\$($workRoot.Substring(3))\data", '-NoSave')
    Assert-Equal '-From に共有フォルダのパス（大文字と小文字は区別しない・サーバーそのものと \wbnoshare は当たらない）' ('{0}/{1}' -f $run.Code, [regex]::Match($run.Log, 'を指すもの \d+ 件').Value) '0/を指すもの 14 件'
    $run = Invoke-Plain @('-Path', (Join-Path $main 'a-folder.lnk'), '-From', '127.0.0.1', '-NoSave')
    Assert-Equal '-Path にショートカットのファイル' ('{0}/{1}' -f $run.Code, $run.Log.Contains('ショートカット 1 件  127.0.0.1 を指すもの 1 件')) '0/True'
    Assert-Equal '一覧では、ショートカットを変えない' (Get-ChangedNames $mainHashes (Get-LinkHashes $main)) ''

    Write-Section '-Exclude・複数の -Path・環境変数'
    $run = Invoke-Tool 'exclude' @('-Path', $main, '-Exclude', 'SUB,日本語*,G-READONLY.LNK')
    $x = Read-Result $run.Dir
    Assert-Equal '-Exclude（名前・* ・大文字と小文字は区別しない）' ('{0}/{1}' -f @($x.Shortcuts | Where-Object { $_.'ショートカット' -like '*i-bracket*' -or $_.'ショートカット' -like '*日本語*' -or $_.'ショートカット' -like '*g-readonly*' }).Count, $x.Shortcuts.Count) '0/15'
    Assert-Equal '除外の行と件数（結果のフォルダとごみ箱を含む）' ('{0}/{1}' -f $run.Log.Contains('除外: SUB, 日本語*, G-READONLY.LNK'), [regex]::Match($run.Log, '除外 \d+ 件').Value) 'True/除外 5 件'
    $comma = Join-Path $workRoot 'c,omma'
    New-Link "$comma\x.lnk" "$O\data\dept"
    $run = Invoke-Plain @('-Path', $comma, '-NoSave')
    Assert-Equal 'カンマを含む名前のフォルダは 1 つのフォルダとして扱う' ('{0}/{1}' -f $run.Code, $run.Log.Contains('ショートカット 1 件')) '0/True'
    $run = Invoke-Plain @('-Path', "$main\sub,$main\日本語 フォルダ", '-NoSave')
    Assert-Equal '-Path をカンマで複数（全体が 1 つのフォルダでなければ区切る）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('ショートカット 2 件'), $run.Log.Contains("対象: $main\sub, $main\日本語 フォルダ")) '0/True/True'
    # 環境変数は、テストが作った変数で確かめる（%TEMP% は、PC によって一時フォルダと文字が一致しない形（8.3 形式の名前など）のことがあるため）。
    # ツールは子のプロセスで動くので、この変数を引き継ぐ
    $env:WB_UST_TEST_ROOT = $workRoot
    try {
        $run = Invoke-Plain @('-Path', '%WB_UST_TEST_ROOT%\main\sub', '-OutputFolder', '%WB_UST_TEST_ROOT%\out\env')
    } finally {
        Remove-Item Env:\WB_UST_TEST_ROOT -ErrorAction SilentlyContinue
    }
    Assert-Equal '-Path と -OutputFolder の環境変数を展開する' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("対象: $main\sub"), [bool](Get-RunDir "$workRoot\out\env")) '0/True/True'

    Write-Section '画面は 200 件まで'
    $many = Join-Path $workRoot 'many'
    New-Link "$many\seed.lnk" "$O\data\dept"
    $seed = Get-Bytes "$many\seed.lnk"
    for ($i = 1; $i -le 600; $i++) { New-TestFile ("$many\s{0:0000}.lnk" -f $i) $seed }
    [IO.File]::Delete("$many\seed.lnk")
    $run = Invoke-Tool 'many' @('-Path', $many)
    $x = Read-Result $run.Dir
    Assert-Equal '保存したときは、画面は 200 件まで' ('{0}/{1}/{2}' -f $run.Log.Contains('（ほか 400 件は shortcuts.csv を参照してください）'), $run.Log.Contains("$many\s0200.lnk"), $run.Log.Contains("$many\s0201.lnk")) 'True/True/False'
    Assert-Equal 'summary.txt と shortcuts.csv はすべて' ('{0}/{1}/{2}' -f @($x.Lines | Where-Object { $_.EndsWith("$many\s0600.lnk") }).Count, @($x.Lines | Where-Object { $_ -like '（ほか*' }).Count, $x.Shortcuts.Count) '1/0/600'
    $run = Invoke-Plain @('-Path', $many, '-NoSave')
    Assert-Equal '-NoSave のときは、画面にすべて' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('（ほか'), $run.Log.Contains("$many\s0600.lnk")) '0/False/True'

    if ($hasShare) {
        Write-Section '置き換え: 確認で N'
        $run = Invoke-Tool 'answer-no' @('-Path', $main, '-From', '127.0.0.1', '-To', 'localhost') 'N'
        $x = Read-Result $run.Dir
        Assert-Equal 'N のときは終了コード 0・置き換えない・控えを作らない' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('置き換えませんでした'), (Get-ResultNames $run.Dir), (Get-ChangedNames $mainHashes (Get-LinkHashes $main))) '0/True/changes.csv,shortcuts.csv,skipped.csv,summary.txt/'
        Assert-Equal 'N のときの結果の列' (Get-ResultList $x.Shortcuts) (Join-Sorted @('a-folder.lnk=置き換えなかった', 'b-file.lnk=置き換えなかった', 'c-missing.lnk=置き換え先が見つからない', 'd-explorer.lnk=置き換えなかった', 'g-readonly.lnk=置き換えられない', 'h-hidden.lnk=置き換えなかった', 'j-server.lnk=置き換えなかった', 'k-indirect.lnk=置き換えなかった', 'l-pipe.lnk=置き換えなかった', 'long.lnk=置き換えられない', 'n-noshare.lnk=置き換え先が見つからない', 'keep.lnk=置き換えなかった', 'i-bracket.lnk=置き換えなかった', 't-denied.lnk=置き換えなかった', 'w-nowrite.lnk=置き換えられない', '営業 共有.lnk=置き換えなかった'))
        # a-folder と、その写しの long.lnk は 2 項目、b-file は 3 項目、ほかの 13 件は 1 項目
        Assert-Equal 'changes.csv の見出しと行の数（項目ごとに 1 行）' ('{0}/{1}' -f (Get-FirstLine $run.Dir 'changes.csv'), $x.Changes.Count) "$changeHead/20"

        Write-Section '置き換え: -Force'
        $repl = Join-Path $workRoot 'repl'
        $ok = Build-Main $repl
        $before = Get-LinkHashes $repl
        $run = Invoke-Tool 'force' @('-Path', $repl, '-From', '127.0.0.1', '-To', 'localhost', '-Force')
        $x = Read-Result $run.Dir
        $after = Get-LinkHashes $repl
        $replaced = 'a-folder.lnk', 'b-file.lnk', 'd-explorer.lnk', 'h-hidden.lnk', 'j-server.lnk', 'k-indirect.lnk', 'l-pipe.lnk', 'ShortcutTarget_memo\keep.lnk', 'sub\[a]\i-bracket.lnk', 't-denied.lnk', '日本語 フォルダ\営業 共有.lnk'
        Assert-Equal '置き換えられないものがあれば終了コード 1' $run.Code 1
        Assert-Equal '結果の列' (Get-ResultList $x.Shortcuts) (Join-Sorted @('a-folder.lnk=置き換えた', 'b-file.lnk=置き換えた', 'c-missing.lnk=置き換え先が見つからない', 'd-explorer.lnk=置き換えた', 'g-readonly.lnk=置き換えられない', 'h-hidden.lnk=置き換えた', 'j-server.lnk=置き換えた', 'k-indirect.lnk=置き換えた', 'l-pipe.lnk=置き換えた', 'long.lnk=置き換えられない', 'n-noshare.lnk=置き換え先が見つからない', 'keep.lnk=置き換えた', 'i-bracket.lnk=置き換えた', 't-denied.lnk=置き換えた', 'w-nowrite.lnk=置き換えられない', '営業 共有.lnk=置き換えた'))
        Assert-Equal '置き換えたものだけ、ファイルが変わった' (Get-ChangedNames $before $after) (Join-Sorted $replaced)
        Assert-Equal '備考（置き換えられない・見つからない）' ('{0}|{1}|{2}' -f (Get-Row $x.Shortcuts $repl 'g-readonly.lnk' '備考'), (Get-Row $x.Shortcuts $repl 'w-nowrite.lnk' '備考'), (Get-Row $x.Shortcuts $repl $longRel '備考')) '読み取り専用のファイルです|書き込む権限がありません|パスが長すぎるため、置き換えられません（ショートカットを書き換える Windows の部品が、260 文字以上のパスを扱えないため）'
        Assert-Equal '備考（置き換え先）' ('{0}|{1}' -f (Get-Row $x.Shortcuts $repl 'c-missing.lnk' '備考'), (Get-Row $x.Shortcuts $repl 'n-noshare.lnk' '備考')) ("$N\data\missing が見つかりません|共有フォルダ \\localhost\wbnoshare が見つかりません" -replace 'C\$', 'c$')
        $backups = @($x.Shortcuts | Where-Object { $_.'控え' } | ForEach-Object { $_.'控え' })
        Assert-Equal '控えは backup\番号_名前（予定の順）' ($backups -join ',') 'backup\0001_a-folder.lnk,backup\0002_b-file.lnk,backup\0003_d-explorer.lnk,backup\0004_h-hidden.lnk,backup\0005_j-server.lnk,backup\0006_k-indirect.lnk,backup\0007_l-pipe.lnk,backup\0008_keep.lnk,backup\0009_i-bracket.lnk,backup\0010_t-denied.lnk,backup\0011_営業 共有.lnk'
        $backupOk = @($x.Shortcuts | Where-Object { $_.'控え' } | Where-Object { (Get-Hash (Join-Path $run.Dir $_.'控え')) -ne $before[$_.'ショートカット'.Substring($repl.Length + 1)] }).Count
        Assert-Equal '控えは置き換える前のファイルと同じ' $backupOk 0
        $leftOld = @($replaced | Where-Object { (Get-TextCount (Join-Path $repl $_) '127.0.0.1') -gt 0 }).Count
        Assert-Equal '置き換えたファイルに、古いサーバーの文字が残っていない（ANSI・UTF-16）' $leftOld 0
        $a = Read-Link "$repl\a-folder.lnk"
        $b = Read-Link "$repl\b-file.lnk"
        Assert-Equal 'リンク先と作業フォルダ' ('{0}|{1}' -f $a.Wsh.ToUpperInvariant(), $a.WorkDir) ("$N\data\dept".ToUpperInvariant() + "|$N\data")
        Assert-Equal 'アイコンの場所（番号はそのまま）とコメント' ('{0},{1}|{2}' -f $b.IconPath, $b.IconIndex, $b.Comment) "$N\data\dept\app.ico,2|場所: $N\data"
        Assert-Equal '引数（" で囲んだパス・\\.\pipe はそのまま）' ('{0}|{1}' -f (Read-Link "$repl\d-explorer.lnk").Arguments, (Read-Link "$repl\l-pipe.lnk").Arguments) "`"$N\data\dept`" /select|\\.\pipe\x /e,$N\data\dept"
        $j = Read-Link "$repl\j-server.lnk"
        Assert-Equal 'サーバーそのもの（リンク先は環境変数の形の欄にある）' ('{0}|{1}' -f $j.EnvTarget, $j.Wsh) '\\localhost|'
        Assert-Equal '@ で始まるコメントは、ファイルの値のまま残る' (Read-Link "$repl\k-indirect.lnk").Comment '@%SystemRoot%\system32\shell32.dll,-22579'
        Assert-Equal '隠し属性は残る・読み取り専用は変えない' ('{0}|{1}' -f (Read-Link "$repl\h-hidden.lnk").Attr, (Read-Link "$repl\g-readonly.lnk").Attr) 'Hidden, Archive|ReadOnly'
        Assert-Equal 'changes.csv（b-file の 3 項目）' (Join-Sorted ($x.Changes | Where-Object { $_.'ショートカット' -like '*b-file.lnk' } | ForEach-Object { '{0}:{1}>{2}:{3}' -f $_.'項目', $_.'置き換える前'.Replace($O, 'O').Replace($O.Replace('C$', 'c$'), 'O'), $_.'置き換えた後'.Replace($N, 'N').Replace($N.Replace('C$', 'c$'), 'N'), $_.'結果' })) 'アイコンの場所:O\data\dept\app.ico,2>N\data\dept\app.ico,2:置き換えた,コメント:場所: O\data>場所: N\data:置き換えた,リンク先:O\data\dept\report.xlsx>N\data\dept\report.xlsx:置き換えた'
        $summary = $x.Lines -join "`n"
        Assert-Equal '予定の件数の行' ([regex]::IsMatch($summary, '\n置き換えるショートカット 11 件\n置き換えないショートカット 5 件（置き換え先が見つからない 2 件・置き換えられない 3 件）\n※ 共有名が変わったときは、-From と -To を \\\\サーバー\\共有 の形で指定してください\n')) 'True'
        Assert-Equal '管理者の案内（管理者でないときだけ）' $summary.Contains('※ 管理者として実行すると、置き換えられるものがあります（パブリックのデスクトップなど）') (-not $isAdmin)
        Assert-Equal '結果の行' ([regex]::IsMatch($summary, '\n結果: 置き換えました  11 件  失敗 0 件  所要 \d+\.\d秒\n保存先: ')) 'True'
        Assert-Equal '置き換えの行' $summary.Contains("`n置き換え: \\127.0.0.1 → \\localhost`n") 'True'
        Assert-Equal '予定の表（置き換えた後のリンク先・項目の名前）' ('{0}/{1}' -f [regex]::IsMatch($summary, '■ 置き換えの予定\n 予定\s+置き換えた後のリンク先\s+ショートカット\n 置き換える\s+'), $summary.Contains("（引数）$N\data\dept")) 'True/True'

        Write-Section '元に戻す（README の方法）'
        $restore = "Set-Location -LiteralPath '{0}'; Import-Csv .\shortcuts.csv | Where-Object {{ `$_.結果 -eq '置き換えた' }} | ForEach-Object {{ Copy-Item -LiteralPath `$_.控え -Destination `$_.ショートカット -Force }}" -f $run.Dir.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $restore 2>&1
        $restoredHashes = Get-LinkHashes $repl
        # 書き換えた直後の隠しショートカットは、ウイルス対策ソフトなどが開いていて戻せないことがある（README に書いたとおり）
        Assert-Equal '置き換えたものが、置き換える前と同じ中身に戻る（隠しショートカットを除く）' (Get-ChangedNames $before $restoredHashes @($replaced | Where-Object { $_ -ne 'h-hidden.lnk' })) ''
        if ($before['h-hidden.lnk'] -eq $restoredHashes['h-hidden.lnk']) {
            Assert-Equal '戻すと隠し属性は外れる（README に書いたとおり）' (Read-Link "$repl\h-hidden.lnk").Attr 'Archive'
        } else {
            Write-Skip '隠しショートカットを戻す' 'ほかのプロセス（ウイルス対策ソフトなど）が開いていて、戻せなかったため'
        }

        Write-Section '置き換え: 置き換え先につながらない（127.0.0.2）'
        $run = Invoke-Tool 'unreachable' @('-Path', $main, '-From', '127.0.0.1', '-To', '127.0.0.2', '-Force')
        $x = Read-Result $run.Dir
        $unreach = @($x.Shortcuts | Where-Object { $_.'結果' -eq '置き換え先に接続できない' }).Count
        Assert-Equal '終了コード 1・接続できない 13 件・置き換えられない 3 件' ('{0}/{1}/{2}' -f $run.Code, $unreach, @($x.Shortcuts | Where-Object { $_.'結果' -eq '置き換えられない' }).Count) '1/13/3'
        Assert-Equal '接続できない注意は、サーバーごとに 1 行' ('{0}/{1}' -f @($x.Lines | Where-Object { $_ -like '※ * に接続できません（*' }).Count, @($x.Lines | Where-Object { $_ -like '※ \\127.0.0.2\c$ に接続できません（*）。新しいサーバーが動いているか、サーバーの名前や IP アドレスが正しいかを確かめてください' }).Count) '1/1'
        Assert-Equal 'サーバーそのものの備考' ((Get-Row $x.Shortcuts $main 'j-server.lnk' '備考') -like '\\127.0.0.2 に接続できません（*）') 'True'
        Assert-Equal 'つながらないときは、ショートカットを変えない・控えを作らない' ('{0}/{1}' -f (Get-ChangedNames $mainHashes (Get-LinkHashes $main)), (Get-ResultNames $run.Dir)) '/changes.csv,shortcuts.csv,skipped.csv,summary.txt'

        # サーバーそのものを指すショートカットだけのとき（ほかの共有の結果を覚えていないので、サーバーそのものを確かめる）
        $srvOnly = Join-Path $workRoot 'srvonly'
        New-Link "$srvOnly\s.lnk" '\\127.0.0.1'
        $srvBefore = Get-LinkHashes $srvOnly
        $run = Invoke-Tool 'srvonly' @('-Path', $srvOnly, '-From', '127.0.0.1', '-To', '127.0.0.2', '-Force')
        $x = Read-Result $run.Dir
        Assert-Equal 'サーバーそのものだけ: つながらなければ置き換えない' ('{0}/{1}/{2}' -f $run.Code, (Get-Row $x.Shortcuts $srvOnly 's.lnk' '結果'), (Get-ChangedNames $srvBefore (Get-LinkHashes $srvOnly))) '1/置き換え先に接続できない/'
        $run = Invoke-Tool 'srvonly-ok' @('-Path', $srvOnly, '-From', '127.0.0.1', '-To', 'localhost', '-Force')
        Assert-Equal 'サーバーそのものだけ: つながれば置き換える' ('{0}/{1}' -f $run.Code, (Read-Link "$srvOnly\s.lnk").EnvTarget) '0/\\localhost'

        Write-Section '置き換え: 共有フォルダのパスで指定'
        $share = Join-Path $workRoot 'share'
        New-Link "$share\p1.lnk" "$O\data\dept"
        New-Link "$share\p2.lnk" "$O\data2\x"
        New-Link "$share\p3.lnk" 'C:\Windows\notepad.exe' @{ Arguments = "`"$O\data`" /x" }
        New-Link "$share\p4.lnk" 'C:\Windows\notepad.exe' @{ Arguments = "$O\data more" }
        $shareBefore = Get-LinkHashes $share
        $run = Invoke-Tool 'share-form' @('-Path', $share, '-From', "$O\data", '-To', "$N\data", '-Force')
        $x = Read-Result $run.Dir
        Assert-Equal '終了コード・結果' ('{0}/{1}' -f $run.Code, (Get-ResultList $x.Shortcuts)) '0/p1.lnk=置き換えた,p3.lnk=置き換えた'
        Assert-Equal '\data2 と、後ろに空白が続く \data は当たらない' (Get-ChangedNames $shareBefore (Get-LinkHashes $share)) 'p1.lnk,p3.lnk'
        Assert-Equal '置き換えた後（一致した部分だけ・-To に書いたとおり）' ('{0}|{1}' -f (Read-Link "$share\p1.lnk").Wsh.ToUpperInvariant(), (Read-Link "$share\p3.lnk").Arguments) ("$N\data\dept".ToUpperInvariant() + "|`"$N\data`" /x")

        Write-Section '置き換え: -To が -From の中・長くなりすぎる'
        $lim = Join-Path $workRoot 'limdir'
        New-Link "$lim\q1.lnk" "$O\lim"
        $head = "$O\lim\x "
        New-Link "$lim\q2.lnk" 'C:\Windows\notepad.exe' @{ Arguments = $head + ('a' * (1023 - $head.Length)) }
        $head = "$O\lim\y "
        New-Link "$lim\q3.lnk" 'C:\Windows\notepad.exe' @{ Description = $head + ('d' * (260 - $head.Length)) }
        $limBefore = Get-LinkHashes $lim
        $run = Invoke-Tool 'to-inside' @('-Path', $lim, '-From', "$O\lim", '-To', "$O\lim\longer") 'N'
        $x = Read-Result $run.Dir
        Assert-Equal '-To が -From の中にあるときの注意（N なので終了コード 0）' ('{0}/{1}' -f $run.Code, $run.Log.Contains("※ -To が -From の中にあります。もう一度実行すると、二重に置き換わります（$O\lim\longer\longer）")) '0/True'
        Assert-Equal '置き換えると長くなりすぎるもの' ('{0}|{1}' -f (Get-Row $x.Shortcuts $lim 'q2.lnk' '備考'), (Get-Row $x.Shortcuts $lim 'q3.lnk' '備考')) '置き換えると引数が長くなりすぎるため、置き換えられません（Windows の部品（WScript.Shell）が扱える長さを超えるため）|置き換えるとコメントが長くなりすぎるため、置き換えられません（Windows の部品（WScript.Shell）が扱える長さを超えるため）'
        Assert-Equal 'q1 は置き換える予定・ファイルは変えない' ('{0}/{1}' -f (Get-Row $x.Shortcuts $lim 'q1.lnk' '結果'), (Get-ChangedNames $limBefore (Get-LinkHashes $lim))) '置き換えなかった/'

        Write-Section '置き換え: 共有名の案内は、共有が見つからないときだけ'
        $hint = Join-Path $workRoot 'hint'
        New-Link "$hint\c.lnk" "$O\data\missing"
        $run = Invoke-Tool 'hint' @('-Path', $hint, '-From', '127.0.0.1', '-To', 'localhost', '-Force')
        Assert-Equal 'フォルダがないだけのときは共有名の案内を出さない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('置き換え先が見つからない 1 件'), $run.Log.Contains('共有名が変わったときは')) '1/True/False'

        Write-Section '置き換え: エクスプローラーで作ったショートカット（リンク先の控えを持つもの）'
        # エクスプローラーの「新規作成 → ショートカット」で \\127.0.0.1\C$\Windows を指して作ったもの（作った PC の名前が入る追跡情報は取り除いた）。
        # リンク先の表示名とパスの控え（PropertyStoreDataBlock）を持っていて、WScript.Shell で書き換えると、控えに古いサーバーが残る
        $explorerLink = [Convert]::FromBase64String((@(
            'TAAAAAEUAgAAAAAAwAAAAAAAAEaCAggAEAAAAIAU8P0BhNoBpL97DJ9y3QGkv3sMn3LdAQBQAAAAAAAAAQAAAAAAAAAAAAAAAAAAAEcAAAAcAAAAAgAAAAAA'
            'AAAAAAAAHAAAAD8AAAAjAAAAAgAAABQAAAAAAAAAAAACAFxcMTI3LjAuMC4xXEMkAFdpbmRvd3MAXQEAAAwAAKAUAB9YDRos8CG+UEOIsHNn/JbvPLMAAACt'
            'ALuvkzufAAQAAAAAAC0AAAAxU1BTc0PlCr5DrU+F5GnchjOYbhEAAAALAAAAAAsAAAD//wAAAAAAAEEAAAAxU1BTMPElt+9HGhCl8QJgjJ7rrCUAAAAKAAAA'
            'AB8AAAAKAAAAMQAyADcALgAwAC4AMAAuADEAAAAAAAAALQAAADFTUFM6pL3eszeDQ5HnRJjaKZWrEQAAAAMAAAAAEwAAAAAAAAAAAAAAAAAAAAAANgDDAMVc'
            'XDEyNy4wLjAuMVxjJABNaWNyb3NvZnQgTmV0d29yawBEZWZhdWx0IHNoYXJlAAIAVgAxAAAAAABsXYdhEABXaW5kb3dzAEAACQAEAO++gVhNN2xdh2EuAAAA'
            'qQcAAAAAAQAAAAAAAAAAAAAAAAAAAFzj2ABXAGkAbgBkAG8AdwBzAAAAFgAAABQDAAABAACgXFwxMjcuMC4wLjFcYyRcV2luZG93cwAAAAAAAAAAAAAAAAAA'
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABcAFwAMQAyADcALgAwAC4AMAAuADEAXABjACQAXABXAGkAbgBkAG8AdwBzAAAA'
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
            'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAASgEAAAkAAKBRAAAAMVNQU+0wvdpDAIlHp/jQE6RzZiI1AAAAZAAAAAAfAAAAEQAAAGMAJAAgACgAXABcADEAMgA3'
            'AC4AMAAuADAALgAxACkAAAAAAAAAAACQAAAAMVNQUzDxJbfvRxoQpfECYIye66whAAAACgAAAAAfAAAACAAAAFcAaQBuAGQAbwB3AHMAAAAVAAAADwAAAABA'
            'AAAAAJV4/gGE2gEpAAAABAAAAAAfAAAACwAAANUwoTCkMOswIADVMKkw6zDAMPwwAAAAABUAAAAOAAAAAEAAAACkv3sMn3LdAQAAAABdAAAAMVNQU6ZqYyg9'
            'ldIRtdYAwE/ZGNBBAAAAHgAAAAAfAAAAFwAAAFwAXAAxADIANwAuADAALgAwAC4AMQBcAGMAJABcAFcAaQBuAGQAbwB3AHMAAAAAAAAAAAAAAAAAAAAAAA=='
        ) -join ''))
        $exp = Join-Path $workRoot 'explorer'
        New-TestFile "$exp\made-by-explorer.lnk" $explorerLink
        Assert-Equal 'エクスプローラーで作ったショートカットのデータ（控えに古いサーバーがある）' ('{0}/{1}' -f (Read-Link "$exp\made-by-explorer.lnk").Wsh, ((Get-TextCount "$exp\made-by-explorer.lnk" '127.0.0.1') -gt 5)) '\\127.0.0.1\c$\Windows/True'
        $run = Invoke-Tool 'explorer' @('-Path', $exp, '-From', '127.0.0.1', '-To', 'localhost', '-Force')
        $x = Read-Result $run.Dir
        Assert-Equal 'エクスプローラーで作ったショートカット: 置き換えて、控えにも古いサーバーが残らない' ('{0}/{1}/{2}/{3}' -f $run.Code, (Get-Row $x.Shortcuts $exp 'made-by-explorer.lnk' '結果'), (Get-TextCount "$exp\made-by-explorer.lnk" '127.0.0.1'), (Read-Link "$exp\made-by-explorer.lnk").Wsh.ToUpperInvariant()) '0/置き換えた/0/\\LOCALHOST\C$\WINDOWS'

        Write-Section '本物の環境では起きない条件（ツールを書き換えたコピー）'
        $pt = Join-Path $workRoot 'pt'
        function Reset-Pt { foreach ($name in 'p-a', 'p-b', 'p-c') { if ([IO.File]::Exists("$pt\$name.lnk")) { [IO.File]::Delete("$pt\$name.lnk") }; New-Link "$pt\$name.lnk" "$O\data\dept" } }
        Reset-Pt
        $ptBefore = Get-LinkHashes $pt
        $patched = New-PatchedTool 'verify' 'if ($wrong.Count -gt 0) {' 'if ($true) {'
        $run = Invoke-Tool 'pt-verify' @('-Path', $pt, '-From', '127.0.0.1', '-To', 'localhost', '-Force') -ToolPath $patched
        $x = Read-Result $run.Dir
        Assert-Equal '読み直しが合わない: 失敗・元に戻す・終了コード 1' ('{0}/{1}/{2}/{3}' -f $run.Code, (Join-Sorted ($x.Shortcuts | ForEach-Object { $_.'結果' } | Select-Object -Unique)), (@($x.Shortcuts | ForEach-Object { $_.'備考' }) -join '|'), (Get-ChangedNames $ptBefore (Get-LinkHashes $pt))) '1/失敗/読み直したが、置き換えた内容と一致しません（元に戻しました）|読み直したが、置き換えた内容と一致しません（元に戻しました）|読み直したが、置き換えた内容と一致しません（元に戻しました）/'
        Assert-Equal '読み直しが合わない: 結果の行' $run.Log.Contains('結果: 置き換えられなかったショートカットがあります  置き換え 0 件  失敗 3 件') 'True'
        Reset-Pt
        $ptBefore = Get-LinkHashes $pt
        $patched = New-PatchedTool 'save' '$link.Save()' "`$link.Save(); throw 'テストで起こした失敗'"
        $run = Invoke-Tool 'pt-save' @('-Path', $pt, '-From', '127.0.0.1', '-To', 'localhost', '-Force') -ToolPath $patched
        $x = Read-Result $run.Dir
        Assert-Equal '保存した後のエラー: 元に戻す' ('{0}/{1}/{2}' -f $run.Code, (Get-Row $x.Shortcuts $pt 'p-a.lnk' '備考'), (Get-ChangedNames $ptBefore (Get-LinkHashes $pt))) '1/テストで起こした失敗（元に戻しました）/'
        Reset-Pt
        $ptBefore = Get-LinkHashes $pt
        $patched = New-PatchedTool 'before-save' '$link.Save()' "throw 'テストで起こした保存の前の失敗'; `$link.Save()"
        $run = Invoke-Tool 'pt-before-save' @('-Path', $pt, '-From', '127.0.0.1', '-To', 'localhost', '-Force') -ToolPath $patched
        $x = Read-Result $run.Dir
        Assert-Equal '保存する前のエラー: ファイルは変わっていないので戻さない' ('{0}/{1}/{2}' -f $run.Code, (Get-Row $x.Shortcuts $pt 'p-a.lnk' '備考'), (Get-ChangedNames $ptBefore (Get-LinkHashes $pt))) '1/テストで起こした保存の前の失敗（ファイルは変わっていません）/'
        $patched = New-PatchedTool 'hash' 'if ([BitConverter]::ToString($sha256.ComputeHash($bytes)) -ne $entry.Hash) {' 'if ($true) {'
        $run = Invoke-Tool 'pt-hash' @('-Path', $pt, '-From', '127.0.0.1', '-To', 'localhost', '-Force') -ToolPath $patched
        $x = Read-Result $run.Dir
        Assert-Equal '調べた後に変わった: 置き換えず、控えも作らない' ('{0}/{1}/{2}/{3}' -f $run.Code, (Get-Row $x.Shortcuts $pt 'p-b.lnk' '備考'), (Get-ResultNames $run.Dir), (Get-ChangedNames $ptBefore (Get-LinkHashes $pt))) '1/調べた後に変更されたため、置き換えていません/changes.csv,shortcuts.csv,skipped.csv,summary.txt/'
        $patched = New-PatchedTool 'unexpected' '$entry.Backup = $backupRelative' "`$entry.Backup = `$backupRelative; if (`$backupNumber -eq 2) { throw 'テストの想定外のエラー' }"
        $run = Invoke-Tool 'pt-unexpected' @('-Path', $pt, '-From', '127.0.0.1', '-To', 'localhost', '-Force') -ToolPath $patched
        $x = Read-Result $run.Dir
        Assert-Equal '置き換えの途中の想定外のエラー: 結果を残す' ('{0}/{1}/{2}' -f ($run.Code -ne 0), (Get-ResultList $x.Shortcuts), $run.Log.Contains('エラーで止まりました。それまでに置き換えたショートカットと、控えたファイルは残っています')) 'True/p-a.lnk=置き換えた,p-b.lnk=置き換えなかった,p-c.lnk=置き換えなかった/True'
        Assert-Equal '置き換えの途中の想定外のエラー: 残りの備考' (Get-Row $x.Shortcuts $pt 'p-c.lnk' '備考') '中断したため、置き換えていません'
        # 置き換え先が権限なしで確かめられないとき（共有フォルダ越しのアクセス拒否は、Windows が前の結果を覚えていてすぐには効かないため、書き換えたコピーで確かめる）
        $ptd = Join-Path $workRoot 'ptd'
        New-Link "$ptd\p-denied.lnk" "$O\data\denied\f.txt"
        $patched = New-PatchedTool 'denied' '        [void][IO.File]::GetAttributes($checkPath)' "        if (`$checkPath -like '*\denied\*') { throw (New-Object UnauthorizedAccessException 'テストの権限なし') }`r`n        [void][IO.File]::GetAttributes(`$checkPath)"
        $run = Invoke-Tool 'pt-denied' @('-Path', $ptd, '-From', '127.0.0.1', '-To', 'localhost', '-Force') -ToolPath $patched
        $x = Read-Result $run.Dir
        Assert-Equal '置き換え先を権限なしで確かめられない: 置き換えて、備考に書く' ('{0}/{1}/{2}' -f $run.Code, (Get-Row $x.Shortcuts $ptd 'p-denied.lnk' '結果'), (Get-Row $x.Shortcuts $ptd 'p-denied.lnk' '備考')) '0/置き換えた/権限がないため、置き換え先を確かめられませんでした'
        $patched = New-PatchedTool 'save-result' "Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'shortcuts.csv'))" "throw 'テストの保存の失敗'; Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'shortcuts.csv'))"
        $run = Invoke-Tool 'pt-save-result' @('-Path', $pt) -ToolPath $patched
        Assert-Equal '結果を保存できないときは警告' ('{0}/{1}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした: テストの保存の失敗')) '0/True'

        Write-Section '保存先'
        $blocker = Join-Path $workRoot 'blocker.txt'
        [IO.File]::WriteAllText($blocker, 'x')
        Reset-Pt
        $ptBefore = Get-LinkHashes $pt
        $run = Invoke-Plain @('-Path', $pt, '-OutputFolder', "$blocker\out")
        Assert-Equal '棚卸しで保存先に書き込めない: 警告して画面だけ' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため、結果は保存しません'), $run.Log.Contains('■ サーバーごとの件数')) '0/True/True'
        $run = Invoke-Plain @('-Path', $pt, '-From', '127.0.0.1', '-To', 'localhost', '-Force', '-OutputFolder', "$blocker\out")
        Assert-Equal '置き換えで保存先に書き込めない: 置き換えずに止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため、置き換えずに止めました'), (Get-ChangedNames $ptBefore (Get-LinkHashes $pt))) '1/True/'
        $longOut = Join-Path $workRoot ('out\' + ('o' * 120) + '\' + ('p' * 120))
        $run = Invoke-Plain @('-Path', $pt, '-From', '127.0.0.1', '-To', 'localhost', '-Force', '-OutputFolder', $longOut)
        $longDir = Get-RunDir $longOut
        Assert-Equal '260 文字を超える保存先にも、結果と控えを保存する' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultNames $longDir), (@((Read-Result $longDir).Shortcuts | Where-Object { $_.'結果' -eq '置き換えた' }).Count)) '0/backup,changes.csv,shortcuts.csv,skipped.csv,summary.txt/3'

        Write-Section 'Ctrl+C（置き換えるときは、コンソールのキーの入力として受け取って止める）'
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
        # $Wait: scan（結果のフォルダができたら＝調べ始めたら）か replace（backup ができたら＝置き換えを始めたら）
        function Invoke-ConsoleStop([string]$Label, [string]$Wait) {
            for ($i = 1; $i -le 600; $i++) { New-TestFile ("$many\s{0:0000}.lnk" -f $i) $seed }
            $out  = Join-Path $workRoot "out\$Label"
            $proc = Start-Process -FilePath (Get-Command $Exe).Source -WindowStyle Hidden -PassThru -ArgumentList @(
                '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $tool), '-Path', ('"{0}"' -f $many), '-From', '127.0.0.1', '-To', 'localhost', '-Force', '-OutputFolder', ('"{0}"' -f $out))
            $null = $proc.Handle   # 終了コードを受け取れるように、終わる前にハンドルを開いておく
            $pattern = if ($Wait -eq 'replace') { '*\backup' } else { '*\ShortcutTarget_*' }
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
            $keyMode = if ($mode -match '^\d+$') { ([long]$mode -band 1) -eq 0 } else { "(読めない: $mode)" }
            [pscustomobject]@{ Code = $proc.ExitCode; Sent = $sent; KeyMode = $keyMode; Dir = (Get-RunDir $out) }
        }
        $seedHash = Get-Hash "$many\s0001.lnk"
        $run = Invoke-ConsoleStop 'ctrlc-replace' 'replace'
        $x = Read-Result $run.Dir
        $kinds = Join-Sorted ($x.Shortcuts | ForEach-Object { $_.'結果' } | Select-Object -Unique)
        # キーを送る補助のプログラムが、ツールのコンソールにつなげないことがある（サーバーの PowerShell 7 で AttachConsole がエラー 5）。
        # そのときはキーが届いていないので、ツールの動きは確かめられない
        $cannotSend = $run.Sent -like 'AttachConsole*' -or $run.Sent -like 'CONIN$*'
        if ($cannotSend) {
            Write-Skip '置き換えの途中で Ctrl+C' ('キーを送れなかったため（{0}）' -f $run.Sent)
        } else {
            Assert-Equal '置き換えの途中: Ctrl+C をキーの入力として受け取り、そのキーで止まる（終了コード 1）' ('{0}/{1}/{2}/{3}' -f $run.Sent, $run.KeyMode, $run.Code, $kinds) 'ok/True/1/置き換えた,置き換えなかった'
            Assert-Equal '置き換えの途中: 残りの備考と、中断の行' ('{0}/{1}' -f (@($x.Shortcuts | Where-Object { $_.'結果' -eq '置き換えなかった' } | ForEach-Object { $_.'備考' } | Select-Object -Unique) -join ','), (@($x.Lines | Where-Object { $_ -eq '中断しました。それまでに置き換えたショートカットと、控えたファイルは残っています' }).Count)) '中断したため、置き換えていません/1'
        }
        # キーが届いたかどうかにかかわらず、結果の列とファイルは合う
        $broken = 0
        foreach ($row in $x.Shortcuts) {
            $isOld = (Get-Hash $row.'ショートカット') -eq $seedHash
            if (($row.'結果' -eq '置き換えた') -eq $isOld) { $broken++ }
            if ($row.'結果' -eq '置き換えた' -and (Get-Hash (Join-Path $run.Dir $row.'控え')) -ne $seedHash) { $broken++ }
        }
        Assert-Equal '置き換えの途中: 結果の列とファイルが合い、置き換えたものは控えてある' $broken 0
        $run = Invoke-ConsoleStop 'ctrlc-scan' 'scan'
        $changed = @(1..600 | Where-Object { (Get-Hash ("$many\s{0:0000}.lnk" -f $_)) -ne $seedHash }).Count
        if ($run.Sent -like 'AttachConsole*' -or $run.Sent -like 'CONIN$*') {
            Write-Skip '調べている途中で Ctrl+C' ('キーを送れなかったため（{0}）' -f $run.Sent)
        } else {
            Assert-Equal '置き換えの前（調べている途中）: 止まって、結果を残さず、ファイルを変えない' ('{0}/{1}/{2}/{3}/{4}' -f $run.Sent, $run.KeyMode, $run.Code, [bool]$run.Dir, $changed) 'ok/True/1/False/0'
        }
    } else {
        Write-Skip '置き換えのテスト' '管理共有を開けないため'
    }

    Write-Section 'コマンドプロンプトから起動'
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $run = Invoke-ViaCmd ('{0} -NoSave -From 127.0.0.1 -Path "{1}\"' -f $cmdBase, $main)
    Assert-Equal '最後の引数が "パス\"' ('{0}/{1}' -f $run.Code, $run.Log.Contains('127.0.0.1 を指すもの 16 件')) '0/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -From 127.0.0.1 -NoSave' -f $cmdBase, $main)
    Assert-Equal '"パス\" の後ろに引数: ヒントを出して止まる（%ERRORLEVEL% は 1）' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('-Path の値に引用符'), $run.Log.Contains('最後に \ を付けないでください'), $run.Log.Contains('ショートカットのリンク先置き換え (')) '1/True/True/False'
    $run = Invoke-ViaCmd ('{0} -Path ''{1}'' -From ''127.0.0.1'',''127.0.0.10'' -NoSave' -f $cmdBase, $main)
    Assert-Equal 'cmd で '' で囲んだ -Path と -From' ('{0}/{1}' -f $run.Code, $run.Log.Contains('を指すもの 18 件')) '0/True'
    if ($hasShare) {
        # 前の項目で置き換えた後なので、置き換える前の状態に作り直す
        Reset-Pt
        $ptBefore = Get-LinkHashes $pt
        $run = Invoke-ViaCmd ('"{0}" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{1}" -Path "{2}" -From 127.0.0.1 -To localhost -OutputFolder "{3}"' -f $Exe, $tool, $pt, (Join-Path $workRoot 'out\nonint'))
        Assert-Equal '-NonInteractive では確認できないので置き換えない・結果を残さない' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('確認できないため、置き換えずに止めました。確認せずに置き換えるときは -Force を付けてください'), (Get-ChangedNames $ptBefore (Get-LinkHashes $pt)), [bool](Get-RunDir (Join-Path $workRoot 'out\nonint'))) '1/True//False'
        $run = Invoke-ViaCmd ('{0} -Path "{1}" -From 127.0.0.1 -To localhost -OutputFolder "{2}"' -f $cmdBase, $pt, (Join-Path $workRoot 'out\stdin-empty'))
        Assert-Equal '標準入力が NUL のときも置き換えない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('確認できないため、置き換えずに止めました'), (Get-ChangedNames $ptBefore (Get-LinkHashes $pt))) '1/True/'
    }

    Write-Section '引数の誤り'
    $noread = Join-Path $workRoot 'noread-root'
    New-Link "$noread\x.lnk" "$O\data\dept"
    $noreadOk = Add-Deny $noread '(RD)'
    $memo = Join-Path $workRoot 'memo.txt'
    [IO.File]::WriteAllText($memo, 'x')
    $cases = @(
        @('-Path なし', @('-From', '127.0.0.1'), '調べるフォルダかショートカットを -Path で指定してください'),
        @('存在しないパス', @('-Path', (Join-Path $workRoot 'no\such')), "フォルダかファイルが見つかりません: $workRoot\no\such"),
        @('.lnk でないファイル', @('-Path', $memo), "ショートカット（.lnk）ではありません: $memo"),
        @('片方がもう片方の中', @('-Path', "$main,$main\sub"), "指定したフォルダの片方が、もう片方の中にあります: $main と $main\sub"),
        @('同じフォルダを 2 回', @('-Path', "$main,$main\"), "指定したフォルダの片方が、もう片方の中にあります: $main と $main"),
        @('-To だけ', @('-Path', $main, '-To', 'localhost'), '-To を使うときは、置き換える前のサーバーを -From で指定してください'),
        @('-To に複数', @('-Path', $main, '-From', 'a', '-To', 'b,c'), '-To は 1 つだけ指定してください'),
        @('形がそろっていない', @('-Path', $main, '-From', 'a', '-To', '\\b\s'), '-From と -To は、どちらもサーバー、またはどちらも \\サーバー\共有 の形で指定してください'),
        @('-From と -To が同じ（大文字と小文字は区別しない）', @('-Path', $main, '-From', '\\SRV\', '-To', 'srv'), '-From と -To が同じです: SRV'),
        @('-From にドライブのパス', @('-Path', $main, '-From', 'C:\share'), '-From は IP アドレスかサーバー名、または \\サーバー\共有 の形で指定してください: C:\share'),
        @('-From が \\ だけ', @('-Path', $main, '-From', '\\'), '-From は IP アドレスかサーバー名、または \\サーバー\共有 の形で指定してください: \\'),
        @('-To に空白を含むサーバー名', @('-Path', $main, '-From', 'a', '-To', 'b c'), '-To は IP アドレスかサーバー名、または \\サーバー\共有 の形で指定してください: b c'),
        @('-To と -NoSave', @('-Path', $main, '-From', 'a', '-To', 'b', '-NoSave'), '置き換えるときは -NoSave を使えません（元のファイルを控えられないため）')
    )
    foreach ($case in $cases) {
        $run = Invoke-Plain $case[1]
        Assert-Equal $case[0] ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains($case[2]), $run.Log.Contains('ショートカットのリンク先置き換え (')) '1/True/False'
    }
    if ($noreadOk) {
        $run = Invoke-Plain @('-Path', $noread)
        Assert-Equal '開けないフォルダ' ('{0}/{1}' -f $run.Code, $run.Log.Contains("フォルダを開けません: $noread（")) '1/True'
    } else {
        Write-Skip '開けないフォルダ' 'アクセス拒否を設定できなかったため'
    }

    Write-Section '制約付き言語モードとヘルプ'
    $clmCommand = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '{0}' -Path '{1}' -OutputFolder '{2}'; exit `$LASTEXITCODE" -f $tool, $main, "$workRoot\clm"
    $clm = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command $clmCommand 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $clmCode = $LASTEXITCODE
    Assert-Equal '制約付き言語モード: 止まる' ('{0}/{1}/{2}' -f $clmCode, $clm.Contains('制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）'), (Test-Path -LiteralPath "$workRoot\clm")) '1/True/False'
    $help = (@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command ("Get-Help '{0}' -Full | Out-String -Width 300" -f $tool) 2>&1) | ForEach-Object { "$_" }) -join "`n"
    Assert-Equal 'ヘルプ（概要・ツール名と版・著作権）' ('{0}/{1}/{2}' -f $help.Contains('ショートカット（.lnk）のリンク先にある共有フォルダのサーバーを'), $help.Contains('Update-ShortcutTarget 1.0.0'), $help.Contains('Copyright (c) 2026 結進舎')) 'True/True/True'

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $main.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        $summary = $x.Lines -join "`n"
        Assert-Equal 'fi-FI で実行して保存できた' ($x.Shortcuts.Count -gt 0) 'True'
        Assert-Equal 'fi-FI でも日時は yyyy-MM-dd HH:mm:ss・所要時間の小数点は .' ('{0}/{1}' -f [regex]::IsMatch($summary, '\n\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}   '), [regex]::IsMatch($summary, '所要 \d+\.\d秒')) 'True/True'
        Assert-Equal 'fi-FI でも shortcuts.csv の日時は yyyy/MM/dd HH:mm:ss' ((Get-Row $x.Shortcuts $main 'a-folder.lnk' '最終更新日時') -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }

    Write-Section 'レビューで見つかった問題の再発防止'
    # 1. コメント・作業フォルダが 261 文字以上のもの（long-comment・long-wd）と、@ で始まるコメント（k-indirect）は、上の棚卸しと置き換えで確かめている
    # 2. 控えの名前と、README の方法で戻せること（「元に戻す」）
    # 3. サーバーそのもの（j-server）、\\.\pipe（l-pipe）、前の結果のフォルダの除外、長くなりすぎるもの（limdir）、つながらないサーバー（127.0.0.2）
    '  （ほかの項目で確かめている）'
} catch {
    '  FAIL  テストが途中で止まりました: {0}（{1} 行目）' -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber
    $script:failCount++
} finally {
    # アクセス拒否を戻し、ジャンクションを外してから消す
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    $global:OutputEncoding = $savedOutputEncoding
    if ($savedInputEncoding) { try { [Console]::InputEncoding = $savedInputEncoding } catch { } }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            foreach ($junction in $junctions) { if ([IO.Directory]::Exists($junction)) { [IO.Directory]::Delete($junction) } }
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
