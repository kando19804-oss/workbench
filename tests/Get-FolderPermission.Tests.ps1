<#
.SYNOPSIS
    Get-FolderPermission の回帰テストです。

.DESCRIPTION
    一時フォルダにダミーのフォルダ（継承を切ったもの、アクセス権を追加したもの、個人・拒否・広い範囲への書き込み・名前の見つからない SID・
    CREATOR OWNER・特殊な権限・空の DACL・NULL の DACL・並び順の誤り・アプリの SID、同じドライブの中で移動したフォルダとファイル、
    所有者を変えた子、汎用のアクセス権で分けて継承した子、260 文字を超えるパス、名前に [ ] や日本語を含むもの、ジャンクション、
    一覧を拒否したフォルダ、継承の印のない古い形の ACL と CREATOR OWNER から作られた所有者の設定）を作り、
    Get-FolderPermission の結果を期待値と照合します。

    アクセス権は、テストが作ったフォルダとファイルにだけ設定し、終了時に元に戻してから削除します。
    共有フォルダは、作るのに管理者が要るため、ツールの処理のうち共有の一覧を読むところだけを書き換えたコピーで確かめます
    （偽の共有は、テストが作ったフォルダを指します）。サーバーの共有は、この PC の管理共有（\\127.0.0.1\C$）を使います（外へは通信しません）。
    共有のルートを共有越しに読んだときの見え方（継承の印が外れる）と、ドメインのグループを調べられないときの理由も、書き換えたコピーで確かめます。

    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Get-FolderPermission を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Get-FolderPermission\Get-FolderPermission.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Get-FolderPermission_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Get-FolderPermission.ps1'
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
$isAdmin   = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$me        = [Security.Principal.WindowsIdentity]::GetCurrent().User
$itemsHead = '"パス","種類","階層","継承","直接の設定","継承した設定","所有者","共有名","注意"'
$acesHead  = '"パス","種類","階層","継承","許可・拒否","アカウント","アカウントの種類","アクセス権","適用先","設定した場所","詳細","値","注意","SID"'
$sharesHead = '"サーバー","共有名","フォルダ","説明","種類","アカウント","アカウントの種類","アクセス許可","値","SID","備考"'
$groupsHead = '"グループ","メンバー","メンバーの種類","経由","状態","SID","備考"'
$orphanSid = 'S-1-5-21-1111111111-2222222222-3333333333-1105'
$longRel   = 'long\' + ((0..5 | ForEach-Object { ('d' * 50) + $_ }) -join '\')
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

function New-Dir([string]$Path) {
    [void][IO.Directory]::CreateDirectory('\\?\' + $Path)
    $Path
}

# ACL の読み書き（5.1 は DirectoryInfo のメソッド、7 は FileSystemAclExtensions の拡張メソッド）
function Get-PathInfo([string]$Path) {
    if ([IO.Directory]::Exists('\\?\' + $Path)) { return New-Object IO.DirectoryInfo ('\\?\' + $Path) }
    New-Object IO.FileInfo ('\\?\' + $Path)
}

function Get-PathAcl([string]$Path) {
    $info = Get-PathInfo $Path
    if ($info.PSObject.Methods['GetAccessControl']) { return $info.GetAccessControl('Access') }
    [IO.FileSystemAclExtensions]::GetAccessControl($info, [Security.AccessControl.AccessControlSections]::Access)
}

function Set-PathAcl([string]$Path, $Acl) {
    $info = Get-PathInfo $Path
    if ($info.PSObject.Methods['SetAccessControl']) { $info.SetAccessControl($Acl) }
    elseif ($info -is [IO.DirectoryInfo]) { [IO.FileSystemAclExtensions]::SetAccessControl([IO.DirectoryInfo]$info, [Security.AccessControl.DirectorySecurity]$Acl) }
    else { [IO.FileSystemAclExtensions]::SetAccessControl([IO.FileInfo]$info, [Security.AccessControl.FileSecurity]$Acl) }
}

# フォルダ（ファイル）にアクセス権を足す（.NET の API。継承は Windows が伝える）
function Add-Rule([string]$Path, [string]$Sid, [string]$Rights, [string]$Type = 'Allow', [string]$Inherit = 'ContainerInherit, ObjectInherit', [string]$Propagate = 'None') {
    $sidObject = New-Object Security.Principal.SecurityIdentifier $Sid
    $acl = Get-PathAcl $Path
    if ([IO.Directory]::Exists('\\?\' + $Path)) {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sidObject, [Security.AccessControl.FileSystemRights]$Rights, [Security.AccessControl.InheritanceFlags]$Inherit, [Security.AccessControl.PropagationFlags]$Propagate, [Security.AccessControl.AccessControlType]$Type)))
    } else {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sidObject, [Security.AccessControl.FileSystemRights]$Rights, [Security.AccessControl.AccessControlType]$Type)))
    }
    Set-PathAcl $Path $acl
}

# DACL を SDDL のとおりに書く（低レベルの SetFileSecurity。継承は伝えず、並び順もそのまま）。終了時に戻すため、パスを覚えておく
function Set-RawDacl([string]$Path, [string]$Sddl, [bool]$KeepAutoInherit = $false) {
    $raw = New-Object Security.AccessControl.RawSecurityDescriptor $Sddl
    if ($KeepAutoInherit) { $raw.SetFlags($raw.ControlFlags -bor [Security.AccessControl.ControlFlags]::DiscretionaryAclAutoInheritRequired -bor [Security.AccessControl.ControlFlags]::DiscretionaryAclAutoInherited) }
    $bytes = New-Object byte[] $raw.BinaryLength
    $raw.GetBinaryForm($bytes, 0)
    $ok = [WbGfpTestSecurity]::SetFileSecurityW('\\?\' + $Path, 4, $bytes)
    if ($ok) { $rawDacls.Add($Path) }
    $ok
}

function Get-Sddl([string]$Path) {
    (Get-PathAcl $Path).GetSecurityDescriptorSddlForm('Access')
}

# 所有者の SID（管理者として実行すると、作ったフォルダの所有者は自分ではなく Administrators になる）
function Get-OwnerSid([string]$Path) {
    $info = Get-PathInfo $Path
    $acl = if ($info.PSObject.Methods['GetAccessControl']) { $info.GetAccessControl('Owner') } else { [IO.FileSystemAclExtensions]::GetAccessControl($info, [Security.AccessControl.AccessControlSections]::Owner) }
    $acl.GetOwner([Security.Principal.SecurityIdentifier])
}

# SID の SDDL での書き方（よく知られた SID は BA などの略号になる）
function ConvertTo-SddlSid([string]$Sid) {
    (New-Object Security.AccessControl.RawSecurityDescriptor ('O:' + $Sid)).GetSddlForm('Owner').Substring(2)
}

function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $out = Join-Path $workRoot "out\$Label"
    $lines = @(& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = (ConvertTo-FlatText (@($lines) -join "`n")); Lines = @($lines); Dir = (Get-RunDir $out) }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行
function Invoke-Plain([string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $lines = @(& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1) | ForEach-Object { "$_" }
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = (ConvertTo-FlatText (@($lines) -join "`n")); Lines = @($lines) }
}

# 相対パスの基準を確かめるため、指定したフォルダから起動する
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

# ツールの処理を書き換えたコピーを作る（本物の環境では起きない条件を確かめるため）。$Pairs は「書き換える前, 後」の組の並び
function New-PatchedTool([string]$Name, [string[]]$Pairs) {
    $text = [IO.File]::ReadAllText($repoTool)
    for ($i = 0; $i -lt $Pairs.Count; $i += 2) {
        if (-not $text.Contains($Pairs[$i])) { throw "書き換える場所が見つかりません: $($Pairs[$i])" }
        $text = $text.Replace($Pairs[$i], $Pairs[$i + 1])
    }
    $file = Join-Path $workRoot "patched\$Name\Get-FolderPermission.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text, $utf8Bom)
    $file
}

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    if (-not [IO.Directory]::Exists('\\?\' + $Parent)) { return $null }
    $found = @([IO.Directory]::GetDirectories('\\?\' + $Parent, 'FolderPermission_*') | Sort-Object)
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
        Shares = @(Read-CsvOrEmpty 'shares.csv'); Items = @(Read-CsvOrEmpty 'folders.csv'); Aces = @(Read-CsvOrEmpty 'permissions.csv')
        Groups = @(Read-CsvOrEmpty 'groups.csv'); Skipped = @(Read-CsvOrEmpty 'skipped.csv'); Lines = $lines; Summary = ($lines -join "`n")
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

# CSV の行のうち、ダミーのフォルダから見た相対パスと列の値の一覧（ダミーのフォルダそのものは「.」）
function Get-RowList($Rows, [string]$Base, [string]$PathColumn, [string]$Column) {
    Join-Sorted ($Rows | ForEach-Object {
        $rel = if ($_.$PathColumn -eq $Base) { '.' } else { $_.$PathColumn.Substring($Base.Length + 1) }
        '{0}={1}' -f $rel, $_.$Column
    })
}

function Get-Rows($Rows, [string]$Base, [string]$Rel, [string]$PathColumn) {
    $full = if ($Rel -eq '.') { $Base } else { Join-Path $Base $Rel }
    @($Rows | Where-Object { $_.$PathColumn -eq $full })
}

# permissions.csv の、あるフォルダのアクセス権の行を「アカウント:アクセス権:適用先:設定した場所」の形で
function Get-AceList($Aces, [string]$Base, [string]$Rel) {
    Join-Sorted (Get-Rows $Aces $Base $Rel 'パス' | ForEach-Object { '{0}:{1}:{2}:{3}' -f $_.'SID', $_.'アクセス権', $_.'適用先', $_.'設定した場所' })
}

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $used -notcontains $_ -and -not [IO.Directory]::Exists("${_}:\") } | Select-Object -First 1
}

# 主な場面のダミー。それぞれのフォルダで、継承を切る・アクセス権を足す・ACL をそのまま書くなどを行う
function Build-Tree([string]$Dir, [string]$Outside) {
    $ok = $true
    $null = New-Dir "$Dir\plain\sub1"
    # 継承を切る（継承していた設定はコピーする）。Users に読み取りと実行を足す
    $p = New-Dir "$Dir\protected"
    $acl = Get-PathAcl $p; $acl.SetAccessRuleProtection($true, $true); Set-PathAcl $p $acl
    Add-Rule $p 'S-1-5-32-545' 'ReadAndExecute'
    $null = New-Dir "$Dir\protected\inner"
    Add-Rule (New-Dir "$Dir\added") 'S-1-1-0' 'Modify'
    Add-Rule (New-Dir "$Dir\added\deny") 'S-1-5-32-546' 'Write' 'Deny'
    Add-Rule (New-Dir "$Dir\orphan") $orphanSid 'Modify'
    Add-Rule (New-Dir "$Dir\user") $me.Value 'FullControl'
    $p = New-Dir "$Dir\creator"
    Add-Rule $p 'S-1-3-0' 'FullControl' 'Allow' 'ContainerInherit, ObjectInherit' 'InheritOnly'
    Add-Rule $p 'S-1-5-32-545' 'ReadAndExecute' 'Allow' 'ContainerInherit' 'None'
    Add-Rule (New-Dir "$Dir\special") 'S-1-5-11' 'ReadAttributes, ReadPermissions, CreateDirectories' 'Allow' 'None' 'None'
    $null = New-Dir "$Dir\files\only"
    [IO.File]::WriteAllText("$Dir\files\a.txt", 'a')
    [IO.File]::WriteAllText("$Dir\files\b.txt", 'b')
    [IO.File]::WriteAllText("$Dir\files\only\one.txt", '1')
    Add-Rule "$Dir\files\a.txt" 'S-1-1-0' 'Read'
    Add-Rule (New-Dir "$Dir\[br]") 'S-1-5-32-545' 'Write'
    Add-Rule (New-Dir "$Dir\日本語 フォルダ") 'S-1-5-32-544' 'Modify'
    $null = New-Dir "$Dir\$longRel"
    Add-Rule "$Dir\$longRel" 'S-1-5-32-545' 'Modify'
    # アプリの SID（ストアのアプリの権限と、そこから作られるグループ）
    $p = New-Dir "$Dir\appsid"
    Add-Rule $p 'S-1-15-3-1024-1111111111-2222222222' 'Read'
    Add-Rule $p 'S-1-5-32-1111111111-2222222222' 'Read'
    Add-Rule (New-Dir "$Dir\backupops") 'S-1-5-32-551' 'Read'
    # 一覧の拒否（アクセス権は所有者として読める）
    $p = New-Dir "$Dir\noread"
    $null = New-Dir "$Dir\noread\hidden"
    Add-Rule $p $me.Value 'ListDirectory' 'Deny' 'None' 'None'
    $denied.Add([pscustomobject]@{ Path = $p; Sid = $me.Value; Rights = 'ListDirectory' })
    # 空の DACL（だれにも許可していない）・NULL の DACL（全員が操作できる）・並び順の誤り（許可の後に拒否）
    $ok = (Set-RawDacl (New-Dir "$Dir\emptydacl") 'D:P') -and $ok
    $ok = (Set-RawDacl (New-Dir "$Dir\nulldacl") 'D:NO_ACCESS_CONTROL') -and $ok
    $ok = (Set-RawDacl (New-Dir "$Dir\order") ('D:P(A;OICI;FA;;;SY)(A;OICI;FA;;;{0})(D;;0x2;;;BG)' -f $me.Value)) -and $ok
    # 同じドライブの中で移動したフォルダとファイル（Everyone の変更を持つ場所から、持たない場所へ）
    $src = New-Dir "$Outside\movesrc"
    Add-Rule $src 'S-1-1-0' 'Modify'
    $null = New-Dir "$src\mv\deep"
    [IO.File]::WriteAllText("$src\mv\f.txt", 'f')
    [IO.File]::WriteAllText("$src\mf.txt", 'm')
    $null = New-Dir "$Dir\moved"
    [IO.Directory]::Move("$src\mv", "$Dir\moved\mv")
    [IO.File]::Move("$src\mf.txt", "$Dir\moved\mf.txt")
    # 所有者を変えた子：親は継承を切って SYSTEM・Authenticated Users（変更）・CREATOR OWNER だけにし、
    # Authenticated Users として作った子の、CREATOR OWNER から作られた設定（作った子の所有者の SID）を、別の SID に書き換える（所有者はそのまま）。
    # 管理者として実行すると所有者は Administrators になるので、親には Administrators を入れない（同じ設定が 2 つになり、書き換える行を決められない）
    $p = New-Dir "$Dir\ownerchg"
    $ok = (Set-RawDacl $p 'D:PAI(A;OICI;FA;;;SY)(A;OICI;0x1301bf;;;AU)(A;OICIIO;FA;;;CO)' $true) -and $ok
    $null = New-Dir "$p\child"
    $childSddl = Get-Sddl "$p\child"
    $creatorAces = @(@((Get-OwnerSid "$p\child").Value, $me.Value) | ForEach-Object { '(A;ID;FA;;;{0})' -f (ConvertTo-SddlSid $_) } | Where-Object { [regex]::Matches($childSddl, [regex]::Escape($_)).Count -eq 1 })
    if ($creatorAces.Count -eq 0) { $ok = $false }
    else { $ok = (Set-RawDacl "$p\child" ($childSddl.Replace($creatorAces[0], ('(A;ID;FA;;;{0}-500)' -f $me.AccountDomainSid.Value))) $true) -and $ok }
    # 汎用のアクセス権（GENERIC_ALL）を下へ渡す設定。子では、効く分と下へ渡す分の 2 つに分かれる
    $p = New-Dir "$Dir\generic"
    $acl = Get-PathAcl $p
    $sddl = $acl.GetSecurityDescriptorSddlForm('Access')
    $acl.SetSecurityDescriptorSddlForm([regex]::Replace($sddl, '^D:([A-Z]*)', 'D:$1(A;OICIIO;GA;;;BU)'), 'Access')
    Set-PathAcl $p $acl
    $null = New-Dir "$p\c"
    if ((Get-Sddl "$p\c") -notmatch 'GA;;;BU') { $ok = $false }
    # ジャンクション
    $null = & cmd.exe /d /c mklink /J "$Dir\link" "$Dir\plain" 2>&1
    $junctions.Add("$Dir\link")
    $ok -and [IO.Directory]::Exists("$Dir\link")
}

# 主な場面で、一覧に出すフォルダと、その「注意」（一覧だけ・既定）
$listNotes = @(
    '.=個人を直接指定', 'added=広い範囲に書き込みを許可', 'added\deny=拒否', 'appsid=', 'backupops=', 'creator=', 'emptydacl=継承を切っている',
    'generic=広い範囲に書き込みを許可', "$longRel=広い範囲に書き込みを許可", 'moved\mv=個人を直接指定・広い範囲に書き込みを許可・継承が親と合わない',
    'noread=個人を直接指定・拒否', 'nulldacl=広い範囲に書き込みを許可', 'order=継承を切っている・個人を直接指定・拒否・並び順が正しくない',
    'orphan=名前が見つからない', 'ownerchg=継承を切っている・広い範囲に書き込みを許可', 'protected=継承を切っている・個人を直接指定',
    'special=広い範囲に書き込みを許可', 'user=個人を直接指定', '[br]=広い範囲に書き込みを許可', '日本語 フォルダ='
)

# 偽の共有（ツールのコピーで、この PC の共有の一覧の代わりに使う）。名前・種類・説明・フォルダ・SDDL（DENIED は読めない、空は設定なし）
function Get-FakeShareCode([object[]]$Specs) {
    # 1 行だけのときに @( @(...) ) がほどけないよう、行を 1 つずつ List に足す
    $rows = @(foreach ($spec in $Specs) {
        "`$FakeShareSpecs.Add(@('{0}', {1}, '{2}', '{3}', '{4}'))" -f ([string]$spec[0]).Replace("'", "''"), $spec[1], ([string]$spec[2]).Replace("'", "''"), ([string]$spec[3]).Replace("'", "''"), ([string]$spec[4]).Replace("'", "''")
    })
    @(
        '$FakeShareSpecs = [System.Collections.Generic.List[object]]::new()'
        ($rows -join "`r`n")
        'function Get-FakeShares {'
        '    $list = [System.Collections.Generic.List[object]]::new()'
        '    foreach ($spec in $FakeShareSpecs) { $s = [WbFolderPermShare]::new(); $s.Name = $spec[0]; $s.Type = [uint32][long]$spec[1]; $s.Remark = $spec[2]; $list.Add($s) }'
        '    return , [WbFolderPermShare[]]$list.ToArray()'
        '}'
        'function Get-FakeShareInfo([string]$Name) {'
        '    $info = [WbFolderPermShareInfo]::new()'
        '    foreach ($spec in $FakeShareSpecs) {'
        '        if ($spec[0] -ne $Name) { continue }'
        '        $info.Path = $spec[3]; $info.Remark = $spec[2]'
        '        if ($spec[4] -eq ''DENIED'') { $info.Error = 5 }'
        '        elseif ($spec[4]) { $raw = [Security.AccessControl.RawSecurityDescriptor]::new($spec[4]); $bytes = [byte[]]::new($raw.BinaryLength); $raw.GetBinaryForm($bytes, 0); $info.Descriptor = $bytes }'
        '    }'
        '    return $info'
        '}'
        ''
    ) -join "`r`n"
}

function New-FakeShareTool([string]$Name, [object[]]$Specs) {
    $anchor = '# 「すべての共有」に含める共有か。'
    New-PatchedTool $Name @(
        $anchor, ((Get-FakeShareCode $Specs) + $anchor),
        '    $found = [WbFolderPermNative]::EnumShares($null, [ref]$enumError)', '    $found = Get-FakeShares',
        '        $info = [WbFolderPermNative]::GetShareInfo($null, $share.Name)', '        $info = Get-FakeShareInfo $share.Name'
    )
}

#endregion

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WbGfpTestSecurity {
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool SetFileSecurityW(string name, uint info, byte[] descriptor);
}
'@

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$denied    = New-Object 'System.Collections.Generic.List[object]'   # 拒否を足したパス（終了時に外す）
$rawDacls  = New-Object 'System.Collections.Generic.List[string]'   # DACL をそのまま書いたパス（終了時に全員の操作を許してから消す）
$junctions = New-Object 'System.Collections.Generic.List[string]'   # 作ったジャンクション（終了時に先に外す）
$substLetter = $null
$data = Join-Path $workRoot 'data'

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool
    $treeOk = Build-Tree $data $workRoot
    Assert-Equal 'ダミーのアクセス権・ジャンクション・移動・所有者を変えた子・汎用のアクセス権を作れた' $treeOk 'True'
    $O = ConvertTo-AdminShare '127.0.0.1' $workRoot
    $hasShare = [IO.Directory]::Exists($O)
    if (-not $hasShare) { Write-Skip '管理共有' "$O を開けないため、共有フォルダ越しのテストを飛ばします" }

    Write-Section '一覧（既定）'
    $run = Invoke-Tool 'list' @('-Path', $data)
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード 0・保存するファイル' ('{0}/{1}' -f $run.Code, (Get-ResultNames $run.Dir)) '0/folders.csv,permissions.csv,shares.csv,skipped.csv,summary.txt'
    Assert-Equal 'すべて BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'folders.csv=True,permissions.csv=True,shares.csv=True,skipped.csv=True,summary.txt=True'
    Assert-Equal 'CSV の見出し' ('{0}|{1}|{2}' -f (Get-FirstLine $run.Dir 'folders.csv'), (Get-FirstLine $run.Dir 'permissions.csv'), (Get-FirstLine $run.Dir 'shares.csv')) ('{0}|{1}|{2}' -f $itemsHead, $acesHead, $sharesHead)
    Assert-Equal '一覧に出すフォルダと注意（継承だけのフォルダ・移動した先の中・所有者を変えた子・汎用で分かれた子は出さない）' (Get-RowList $x.Items $data 'パス' '注意') (Join-Sorted $listNotes)
    Assert-Equal '件数の行' ([regex]::IsMatch($run.Log, 'フォルダ 35 件  アクセス権を変えてあるフォルダ 19 件  所要 \d+\.\d秒')) 'True'
    Assert-Equal 'スキップとリンクの行' $run.Log.Contains('スキップ 2 件（権限なし）  たどらなかったリンク 1 件') 'True'
    Assert-Equal 'skipped.csv' (Get-RowList $x.Skipped $data 'パス' '理由') 'emptydacl=権限なし,link=リンク,noread=権限なし'
    Assert-Equal 'skipped.csv の詳細（一覧だけ読めない・リンクの種類）' ('{0}|{1}' -f (Get-Rows $x.Skipped $data 'noread' 'パス')[0].'詳細', (Get-Rows $x.Skipped $data 'link' 'パス')[0].'詳細') '中の一覧を読めません（アクセスが拒否されました）|ジャンクション'
    Assert-Equal '共有フォルダがないときの表示' $run.Log.Contains('（このフォルダにつながる共有フォルダはありません）') 'True'
    Assert-Equal '管理者でないときの案内' ($run.Log.Contains('※ 管理者として実行すると、読める範囲が広がります') -eq (-not $isAdmin)) 'True'
    Assert-Equal '確かめたほうがよいところ（件数）' ([regex]::IsMatch($x.Summary, '継承を切っている\s+4 フォルダ') -and [regex]::IsMatch($x.Summary, '継承が親と合わない\s+1 件') -and
        [regex]::IsMatch($x.Summary, '個人を直接指定\s+6 件') -and [regex]::IsMatch($x.Summary, '名前が見つからない\s+1 件（' + [regex]::Escape($orphanSid) + '）') -and
        [regex]::IsMatch($x.Summary, '広い範囲に書き込みを許可\s+8 件') -and [regex]::IsMatch($x.Summary, '拒否\s+3 件') -and [regex]::IsMatch($x.Summary, '並び順が正しくない\s+1 件')) 'True'
    Assert-Equal '画面の状態（継承だけ・継承を切っている・追加している・継承が親と合わない・アクセス権の設定がない・並び順）' ('{0}/{1}/{2}/{3}/{4}/{5}' -f $x.Summary.Contains(" $data  "), [regex]::IsMatch($x.Summary, [regex]::Escape("$data\protected") + '\s+継承を切っている'),
        [regex]::IsMatch($x.Summary, [regex]::Escape("$data\added") + '\s+追加している'), [regex]::IsMatch($x.Summary, [regex]::Escape("$data\moved\mv") + '\s+継承が親と合わない'),
        [regex]::IsMatch($x.Summary, [regex]::Escape("$data\nulldacl") + '\s+アクセス権の設定がない'), [regex]::IsMatch($x.Summary, [regex]::Escape("$data\order") + '\s+継承を切っている・並び順が正しくない')) 'True/True/True/True/True/True'
    Assert-Equal '-Path のフォルダは継承した設定も「許可（継承）」で出す・空の DACL は「だれにも許可していません」' ('{0}/{1}' -f [regex]::IsMatch($x.Summary, '許可（継承）\s+NT AUTHORITY\\SYSTEM'), $x.Summary.Contains('（だれにも許可していません）')) 'True/True'
    Assert-Equal '画面は 20 フォルダまで' ('{0}/{1}' -f $run.Log.Contains('（ほか'), (@($x.Items).Count -le 20)) ('{0}/True' -f (@($x.Items).Count -gt 20))

    Write-Section 'アクセス権の名前・適用先・アカウントの種類'
    Assert-Equal '追加した設定と継承した設定' (Get-AceList $x.Aces $data 'added') (Join-Sorted @('S-1-1-0:変更:このフォルダ・サブフォルダ・ファイル:このフォルダ', 'S-1-5-18:フル コントロール:このフォルダ・サブフォルダ・ファイル:親から継承', 'S-1-5-32-544:フル コントロール:このフォルダ・サブフォルダ・ファイル:親から継承', ('{0}:フル コントロール:このフォルダ・サブフォルダ・ファイル:親から継承' -f $me.Value)))
    $creatorRows = Get-Rows $x.Aces $data 'creator' 'パス' | Where-Object { $_.'設定した場所' -eq 'このフォルダ' }
    Assert-Equal 'CREATOR OWNER（サブフォルダ・ファイルだけ）と、フォルダにだけ効く読み取りと実行（フォルダの内容の一覧表示）' (Join-Sorted ($creatorRows | ForEach-Object { '{0}:{1}:{2}:{3}' -f $_.'アカウント', $_.'アクセス権', $_.'適用先', $_.'アカウントの種類' })) (Join-Sorted @('CREATOR OWNER:フル コントロール:サブフォルダ・ファイルだけ:組み込み', 'BUILTIN\Users:フォルダの内容の一覧表示:このフォルダ・サブフォルダ:グループ'))
    $special = @(Get-Rows $x.Aces $data 'special' 'パス' | Where-Object { $_.'設定した場所' -eq 'このフォルダ' })[0]
    Assert-Equal '特殊（詳細と値）' ('{0}|{1}|{2}|{3}' -f $special.'アクセス権', $special.'詳細', $special.'値', $special.'適用先') '特殊|属性の読み取り、フォルダの作成/データの追加、アクセス許可の読み取り|0x00120084|このフォルダだけ'
    $deny = @(Get-Rows $x.Aces $data 'added\deny' 'パス' | Where-Object { $_.'設定した場所' -eq 'このフォルダ' })[0]
    Assert-Equal '拒否と書き込み' ('{0}|{1}|{2}' -f $deny.'許可・拒否', $deny.'アクセス権', $deny.'注意') '拒否|書き込み|拒否'
    $orphan = @(Get-Rows $x.Aces $data 'orphan' 'パス' | Where-Object { $_.'設定した場所' -eq 'このフォルダ' })[0]
    Assert-Equal '名前が見つからない SID' ('{0}|{1}|{2}' -f $orphan.'アカウント', $orphan.'アカウントの種類', $orphan.'SID') "$orphanSid|名前が見つからない|$orphanSid"
    Assert-Equal 'アプリの SID（S-1-15- と S-1-5-32-<長い数>）は「Windows のアプリ」で印を付けない' (Join-Sorted (Get-Rows $x.Aces $data 'appsid' 'パス' | Where-Object { $_.'設定した場所' -eq 'このフォルダ' } | ForEach-Object { '{0}:{1}:{2}' -f $_.'アカウント', $_.'アカウントの種類', $_.'注意' })) (Join-Sorted @('S-1-15-3-1024-1111111111-2222222222:Windows のアプリ:', 'S-1-5-32-1111111111-2222222222:Windows のアプリ:'))
    $nullRow = @(Get-Rows $x.Aces $data 'nulldacl' 'パス')
    Assert-Equal 'NULL の DACL は Everyone のフル コントロールとして出す' ('{0}|{1}|{2}' -f $nullRow.Count, $nullRow[0].'アカウント', $nullRow[0].'詳細') '1|Everyone|アクセス権の設定（DACL）がないため、全員が操作できます'
    $orderItem = @(Get-Rows $x.Items $data 'order' 'パス')[0]
    Assert-Equal 'folders.csv（継承・直接の設定・継承した設定・所有者）' ('{0}|{1}|{2}|{3}' -f $orderItem.'継承', $orderItem.'直接の設定', $orderItem.'継承した設定', $orderItem.'所有者'.ToUpperInvariant()) ('切っている|3|0|{0}' -f (Get-OwnerSid "$data\order").Translate([Security.Principal.NTAccount]).Value.ToUpperInvariant())
    $genericRows = @(Get-Rows $x.Aces $data 'generic' 'パス' | Where-Object { $_.'アカウント' -eq 'BUILTIN\Users' -and $_.'設定した場所' -eq 'このフォルダ' })
    Assert-Equal '汎用のアクセス権は直して名前を付ける（GENERIC_ALL → フル コントロール）' ('{0}|{1}|{2}' -f $genericRows.Count, $genericRows[0].'アクセス権', $genericRows[0].'適用先') '1|フル コントロール|サブフォルダ・ファイルだけ'
    Assert-Equal '[ ] と日本語と 260 文字を超えるパス' ('{0}/{1}/{2}' -f @(Get-Rows $x.Items $data '[br]' 'パス').Count, @(Get-Rows $x.Items $data '日本語 フォルダ' 'パス').Count, (Get-Rows $x.Items $data $longRel 'パス')[0].'階層') '1/1/7'

    Write-Section '継承が親と合わないフォルダ（同じドライブの中で移動したもの）'
    $mv = @(Get-Rows $x.Aces $data 'moved\mv' 'パス')
    Assert-Equal '移動したフォルダは、前の親から継承した設定を出して数える' ('{0}/{1}' -f @($mv | Where-Object { $_.'アカウント' -eq 'Everyone' -and $_.'設定した場所' -eq '親から継承' }).Count, [regex]::IsMatch($x.Summary, '継承が親と合わない\s+1 件（' + [regex]::Escape("$data\moved\mv") + '）')) '1/True'
    Assert-Equal '移動したフォルダの中・所有者を変えた子・汎用で分かれた子は合っている' ('{0}/{1}/{2}' -f @(Get-Rows $x.Items $data 'moved\mv\deep' 'パス').Count, @(Get-Rows $x.Items $data 'ownerchg\child' 'パス').Count, @(Get-Rows $x.Items $data 'generic\c' 'パス').Count) '0/0/0'
    $run = Invoke-Tool 'files' @('-Path', "$data\moved", '-IncludeFile')
    $y = Read-Result $run.Dir
    Assert-Equal '-IncludeFile: 移動したファイルは合わない、移動したフォルダの中のファイルは合っている' (Get-RowList $y.Items "$data\moved" 'パス' '注意') (Join-Sorted @('.=個人を直接指定', 'mf.txt=個人を直接指定・広い範囲に書き込みを許可・継承が親と合わない', 'mv=個人を直接指定・広い範囲に書き込みを許可・継承が親と合わない'))
    $run = Invoke-Plain @('-Path', "$data\moved\mv", '-NoSave')
    Assert-Equal '-Path に移動したフォルダそのものを指定しても、本物の親と比べる' ([regex]::IsMatch($run.Log, '継承が親と合わない\s+1 件')) 'True'

    Write-Section '古い形のアクセス権（CREATOR OWNER から作られた所有者の設定）'
    # 自動継承の印のない親（Windows Server の C:\ など）の下に作ると、継承の印のない古い形の ACL になる。CREATOR OWNER から作られた
    # 所有者の設定は、親から継承した同じアカウントの設定と重なるため、Windows は継承として付け直さず「このフォルダ」の設定として返す
    # （ファイルでは、Windows が作るときに同じ設定を 1 つにまとめるため、この形は残らない。古い形のフォルダの中に作ったファイルが出ないことだけ確かめる）
    $legacy = New-Dir "$workRoot\legacy"
    $made = New-Dir "$legacy\made"
    $ownerSddl = ConvertTo-SddlSid (Get-OwnerSid $made).Value
    $legacyOk = Set-RawDacl $legacy ('D:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;{0})(A;OICI;0x1200a9;;;BU)(A;OICIIO;GA;;;CO)' -f $ownerSddl) $true
    $legacyOk = (Set-RawDacl $made ('D:(A;OICI;FA;;;SY)(A;OICI;FA;;;{0})(A;OICI;0x1200a9;;;BU)(A;;FA;;;{0})(A;OICIIO;GA;;;CO)' -f $ownerSddl)) -and $legacyOk
    [IO.File]::WriteAllText("$made\f.txt", 'f')
    Assert-Equal '準備: 古い形の ACL で、CREATOR OWNER から作られた所有者の設定を、Windows は「このフォルダ」の設定として返す' ('{0}/{1}' -f $legacyOk, (Get-Sddl $made).Contains(('(A;;FA;;;{0})' -f $ownerSddl))) 'True/True'
    $run = Invoke-Plain @('-Path', $legacy, '-IncludeFile', '-NoSave')
    Assert-Equal 'CREATOR OWNER から作られた所有者の設定だけなら、継承した設定として扱い、一覧に出さない（中に作ったファイルも出ない）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('フォルダ 2 件  ファイル 1 件  アクセス権を変えてあるフォルダ 1 件・ファイル 0 件')) '0/True'
    $run = Invoke-Plain @('-Path', $made, '-NoSave')
    Assert-Equal '-Path に指定したフォルダでも、その設定は「許可（継承）」で出す' ('{0}/{1}' -f [regex]::IsMatch($run.Log, [regex]::Escape($made) + '\s+継承だけ'), [regex]::IsMatch($run.Log, '許可（継承）\s+\S+\s+フル コントロール\s+このフォルダだけ')) 'True/True'

    Write-Section 'ファイル（-IncludeFile）'
    $run = Invoke-Tool 'include' @('-Path', "$data\files", '-IncludeFile')
    $y = Read-Result $run.Dir
    Assert-Equal 'アクセス権を変えてあるファイルだけを出す（中身が 1 件だけのフォルダも読める）' ('{0}/{1}' -f $run.Code, (Get-RowList $y.Items "$data\files" 'パス' '種類')) ('0/' + (Join-Sorted @('.=フォルダ', 'a.txt=ファイル')))
    Assert-Equal 'ファイルの適用先と設定した場所' (Get-AceList $y.Aces "$data\files" 'a.txt' | ForEach-Object { $_ -split ',' } | Where-Object { $_ -like 'S-1-1-0:*' }) 'S-1-1-0:読み取り:このファイル:このファイル'
    Assert-Equal '件数の行と見出し' ('{0}/{1}' -f [regex]::IsMatch($run.Log, 'フォルダ 2 件  ファイル 3 件  アクセス権を変えてあるフォルダ 0 件・ファイル 1 件'), $run.Log.Contains('■ アクセス権を変えてあるフォルダとファイル')) 'True/True'

    Write-Section '-All・-Depth・-Exclude・-NoSave'
    $run = Invoke-Tool 'all' @('-Path', $data, '-All', '-Depth', '1', '-Exclude', 'long,link')
    $y = Read-Result $run.Dir
    Assert-Equal '-All は継承だけのフォルダも出す・-Depth 1 は直下まで' ('{0}/{1}/{2}' -f @(Get-Rows $y.Items $data 'plain' 'パス').Count, @(Get-Rows $y.Items $data 'plain\sub1' 'パス').Count, $run.Log.Contains('■ フォルダのアクセス権（すべてのフォルダ）')) '1/0/True'
    Assert-Equal '階層と除外の行' ('{0}/{1}/{2}' -f $run.Log.Contains('階層: 1 階層下まで'), $run.Log.Contains('除外: long, link'), $run.Log.Contains('除外 2 件')) 'True/True/True'
    Assert-Equal '除外したものは skipped.csv に「除外指定」' (Get-RowList ($y.Skipped | Where-Object { $_.'理由' -eq '除外指定' }) $data 'パス' '詳細') 'link=link,long=long'
    $run = Invoke-Plain @('-Path', $data, '-NoSave')
    Assert-Equal '-NoSave はすべて画面に出す・保存しない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("$data\[br]"), $run.Log.Contains('保存先')) '0/True/False'
    $run = Invoke-Plain @('-Path', $data, '-Depth', '２', '-NoSave')
    Assert-Equal '-Depth に全角の数字' ('{0}/{1}' -f $run.Code, $run.Log.Contains('階層: 2 階層下まで')) '0/True'

    Write-Section 'グループのメンバー（-ExpandGroup）'
    $run = Invoke-Tool 'groups' @('-Path', $data, '-ExpandGroup')
    $y = Read-Result $run.Dir
    Assert-Equal 'groups.csv を作る・見出し' ('{0}/{1}' -f (Get-ResultNames $run.Dir).Contains('groups.csv'), (Get-FirstLine $run.Dir 'groups.csv')) "True/$groupsHead"
    $adminMembers = @($y.Groups | Where-Object { $_.'グループ' -eq 'BUILTIN\Administrators' -and $_.'メンバー' })
    $localAdmins = @(([ADSI]'WinNT://./Administrators,group').Invoke('Members') | ForEach-Object { (New-Object Security.Principal.SecurityIdentifier ([byte[]]$_.GetType().InvokeMember('objectSid', 'GetProperty', $null, $_, $null)), 0).Value })
    Assert-Equal 'BUILTIN\Administrators のメンバー（ローカルのグループ）' (Join-Sorted ($adminMembers | ForEach-Object { $_.'SID' })) (Join-Sorted $localAdmins)
    Assert-Equal '広い範囲のグループと組み込みは展開しない' ('{0}/{1}' -f (($y.Groups | Where-Object { $_.'グループ' -eq 'BUILTIN\Users' } | ForEach-Object { $_.'備考' }) -join '|'), (($y.Groups | Where-Object { $_.'グループ' -eq 'Everyone' } | ForEach-Object { $_.'備考' }) -join '|')) 'ほぼ全員が入っているグループのため、メンバーを調べません/組み込みのアカウントのため、メンバーを調べません（Windows が条件で決めるため）'
    $backupCount = @(([ADSI]'WinNT://./Backup Operators,group').Invoke('Members')).Count
    if ($backupCount -eq 0) {
        Assert-Equal 'メンバーのいないグループは「メンバーはいません」' (($y.Groups | Where-Object { $_.'グループ' -eq 'BUILTIN\Backup Operators' } | ForEach-Object { $_.'備考' }) -join '|') 'メンバーはいません'
    } else {
        Write-Skip 'メンバーのいないグループ' 'この PC の Backup Operators にメンバーがいるため'
    }
    $adminRid = $me.AccountDomainSid.Value + '-500'
    $adminRow = @($adminMembers | Where-Object { $_.'SID' -eq $adminRid })
    if ($adminRow.Count -eq 1) {
        $disabled = ([int](([ADSI]('WinNT://./{0},user' -f ($adminRow[0].'メンバー' -split '\\')[-1])).UserFlags.Value) -band 2) -ne 0
        Assert-Equal '無効なユーザーは「状態」が無効' $adminRow[0].'状態' $(if ($disabled) { '無効' } else { '' })
    } else {
        Write-Skip '無効なユーザー' '組み込みの Administrator が Administrators のメンバーでないため'
    }
    Assert-Equal '画面の件数' ([regex]::IsMatch($run.Log, 'グループ \d+ 件  メンバー \d+ 件（groups.csv を参照してください）')) 'True'
    $domainJoined = $false
    try { $domainJoined = [bool](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).PartOfDomain } catch { }
    if (-not $domainJoined) { Write-Skip 'ドメインのグループ（LDAP）' 'この PC はドメインに参加していないため（FGTEST-DC で確かめる）' }
    # ドメインのグループを調べられないときの理由。本物のドメインのグループはないので、BUILTIN\Administrators をドメインのグループとして扱わせる
    $asDomain = @('    if ($isLocal) { return Get-WinNTMember $Account $System $computer }', '    if ($isLocal -and $Account.Sid -ne ''S-1-5-32-544'') { return Get-WinNTMember $Account $System $computer }')
    $joinLine = '        try { $domainJoined = [bool](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).PartOfDomain } catch { }'
    $run = Invoke-Plain @('-Path', $data, '-ExpandGroup', '-NoSave') (New-PatchedTool 'domain-notjoined' ($asDomain + @($joinLine, '        $domainJoined = $false')))
    Assert-Equal 'ドメインに参加していない PC では、ドメインのグループを調べられない理由を表示する' ([regex]::IsMatch($run.Log, '調べられなかったグループ: BUILTIN\\Administrators（この PC はドメインに参加していないため、ドメインのグループのメンバーを調べられません。')) 'True'
    if ([string]::Equals([Environment]::UserDomainName, $env:COMPUTERNAME, [StringComparison]::OrdinalIgnoreCase)) {
        $run = Invoke-Plain @('-Path', $data, '-ExpandGroup', '-NoSave') (New-PatchedTool 'domain-localuser' ($asDomain + @($joinLine, '        $domainJoined = $true')))
        Assert-Equal 'ローカルのアカウントで実行したときは、ドメインのグループを調べられない理由を表示する' ($run.Log.Contains(('調べられなかったグループ: BUILTIN\Administrators（ローカルのアカウント（{0}\{1}）で実行しているため、ドメインのグループのメンバーを調べられません。' -f [Environment]::UserDomainName, [Environment]::UserName))) 'True'
    } else {
        Write-Skip 'ローカルのアカウントで実行したとき' 'ドメインのアカウントで実行しているため'
    }

    Write-Section '共有フォルダ（この PC の共有の一覧を、偽の共有に差し替えたコピー）'
    $S = Join-Path $workRoot 'shares'
    foreach ($rel in 'outer\mid\inner\deep', 'same', 'soumu', 'nosd', 'denied', 'emptysh', 'printd') { $null = New-Dir "$S\$rel" }
    Add-Rule "$S\outer\mid" $me.Value 'ListDirectory' 'Deny' 'None' 'None'
    $denied.Add([pscustomobject]@{ Path = "$S\outer\mid"; Sid = $me.Value; Rights = 'ListDirectory' })
    $specs = @(
        @('outer', 0, '外側', "$S\outer", 'D:(A;;FA;;;WD)'),
        @('inner', 0, '内側', "$S\outer\mid\inner", 'D:(A;;FA;;;BA)'),
        @('same1', 0, '', "$S\same", 'D:(A;;0x1200a9;;;WD)'),
        @('same2', 0, '', "$S\same", 'D:(A;;0x1200a9;;;WD)'),
        @('総務$', 0, '総務部', "$S\soumu", 'D:(A;;0x1301bf;;;BU)'),
        @('nosd', 0, '', "$S\nosd", ''),
        @('denied', 0, '', "$S\denied", 'DENIED'),
        @('emptysh', 0, '', "$S\emptysh", 'D:'),
        @('print$', 0, 'プリンター ドライバー', "$S\printd", 'D:(A;;FA;;;WD)'),
        @('C$', 2147483648, '既定の共有', 'C:\', ''),
        @('printer', 1, '', '', '')
    )
    $fakeTool = New-FakeShareTool 'shares' $specs
    $run = Invoke-Tool 'shares' @() $fakeTool
    $y = Read-Result $run.Dir
    Assert-Equal '-Path を省略すると、この PC の共有すべて（管理用・print$・プリンターは除く）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('対象: この PC の共有フォルダ 8 件'), (Join-Sorted ($y.Shares | ForEach-Object { $_.'共有名' } | Select-Object -Unique))) ('0/True/' + (Join-Sorted @('denied', 'emptysh', 'inner', 'nosd', 'outer', 'same1', 'same2', '総務$')))
    Assert-Equal 'shares.csv（アクセス許可・説明・読めない・設定なし・だれにも許可しない）' ('{0}|{1}|{2}|{3}|{4}' -f
        (($y.Shares | Where-Object { $_.'共有名' -eq 'outer' } | ForEach-Object { '{0}:{1}:{2}:{3}' -f $_.'アカウント', $_.'アクセス許可', $_.'説明', $_.'フォルダ' }) -join ';'),
        (($y.Shares | Where-Object { $_.'共有名' -eq '総務$' } | ForEach-Object { '{0}:{1}' -f $_.'アカウント', $_.'アクセス許可' }) -join ';'),
        ([string](@($y.Shares | Where-Object { $_.'共有名' -eq 'denied' })[0].'備考')).StartsWith('共有のアクセス許可を読めません（'),
        @($y.Shares | Where-Object { $_.'共有名' -eq 'nosd' })[0].'備考', @($y.Shares | Where-Object { $_.'共有名' -eq 'emptysh' })[0].'備考') ("Everyone:フル コントロール:外側:$S\outer|BUILTIN\Users:変更|True|設定なし（Windows の既定のアクセス許可が使われます）|だれにも許可していません")
    Assert-Equal '同じフォルダの 2 つの共有名・重なった共有は外側を 1 回・外側から届かない内側は別に調べる' (Get-RowList $y.Items $S 'パス' '共有名') (Join-Sorted @('denied=denied', 'emptysh=emptysh', 'nosd=nosd', 'outer=outer', 'outer\mid=', 'outer\mid\inner=inner', 'same=same1、same2', 'soumu=総務$'))
    Assert-Equal '内側は階層 0 として調べる・外側の途中の読めないフォルダは skipped.csv に' ('{0}/{1}' -f (Get-Rows $y.Items $S 'outer\mid\inner' 'パス')[0].'階層', (Get-RowList $y.Skipped $S 'パス' '理由')) '0/outer\mid=権限なし'
    Assert-Equal '画面（共有の表と共有名）' ('{0}/{1}/{2}' -f [regex]::IsMatch($y.Summary, ' 総務\$\s+' + [regex]::Escape("$S\soumu") + '\s+許可\s+BUILTIN\\Users\s+変更'), $y.Summary.Contains('（共有: same1、same2）'), $y.Summary.Contains('（読めません: ')) 'True/True/True'
    $run = Invoke-Plain @('-Path', "$S\outer\mid\inner", '-NoSave') $fakeTool
    Assert-Equal '-Path がこの PC のフォルダのときは、含む共有と同じ共有を出す' ('{0}/{1}/{2}' -f $run.Log.Contains(' inner '), $run.Log.Contains(' outer '), $run.Log.Contains(' same1 ')) 'True/True/False'
    $run = Invoke-Plain @('-Path', $S, '-NoSave') $fakeTool
    Assert-Equal '-Path の中にある共有も出す' ('{0}/{1}' -f $run.Log.Contains(' same1 '), $run.Log.Contains(' print$ ')) 'True/False'
    $noShareTool = New-FakeShareTool 'noshare' @(, @('C$', 2147483648, '', 'C:\', ''))
    $run = Invoke-Plain @() $noShareTool
    Assert-Equal 'この PC に共有がないときに -Path を省略すると止まる' ('{0}/{1}' -f $run.Code, $run.Log.Contains('この PC には共有フォルダがありません。調べるフォルダを -Path で指定してください')) '1/True'

    Write-Section 'サーバーの共有フォルダ（\\127.0.0.1）'
    if ($hasShare) {
        $run = Invoke-Plain @('-Path', (ConvertTo-AdminShare '127.0.0.1' "$data\added"), '-NoSave')
        Assert-Equal '共有フォルダの中を指定: 共有のアクセス許可（設定のない管理共有は既定の設定）と、サーバーに問い合わせた名前' ('{0}/{1}/{2}' -f $run.Code, [regex]::IsMatch($run.Log, '\\\\127\.0\.0\.1\\C\$\s+C:\\\s+許可\s+BUILTIN\\Administrators'), $run.Log.Contains('Everyone')) '0/True/True'
        $run = Invoke-Plain @('-Path', (ConvertTo-AdminShare '127.0.0.1' "$data\moved"), '-NoSave')
        Assert-Equal '共有フォルダ越しでも、継承が親と合わないフォルダを見つける' ([regex]::IsMatch($run.Log, '継承が親と合わない\s+1 件')) 'True'
        $adminTool = New-PatchedTool 'admin-shares' @(
            'return ($type -band 0xFF) -eq 0 -and ($type -band 2147483648L) -eq 0 -and -not ($AutoShares -contains $Share.Name)', 'return @(''C$'', ''ADMIN$'') -contains $Share.Name')
        $run = Invoke-Plain @('-Path', '\\127.0.0.1', '-Depth', '1', '-NoSave') $adminTool
        Assert-Equal '\\サーバー: 重なった共有（ADMIN$ は C$ の中）は外側を 1 回だけ調べ、中の共有のフォルダに共有名を付ける' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('対象: \\127.0.0.1（共有フォルダ 2 件）'), [regex]::IsMatch($run.Log, '\\\\127\.0\.0\.1\\C\$\\Windows\s+.*（共有: ADMIN\$）')) '0/True/True'
        # 共有のルートを共有越しに読むと、Windows は親をたどれず、継承の印のない古い形の ACL では継承した設定も「このフォルダ」になる。
        # 本物の共有は作れないので、共有の中のフォルダを共有のルートとして扱わせ、読んだ結果から継承の印（と自動継承の印）を外して、その見え方をまねる
        $stripCode = @'
function Remove-InheritedMark($Security) {
    if ($Security.Error -ne 0 -or $Security.NullDacl) { return $Security }
    $d = $Security.Descriptor
    $d[3] = $d[3] __AUTO__
    $Security.AutoInherited = ($d[3] -band 0x04) -ne 0
    $pos = [BitConverter]::ToInt32($d, 16) + 8
    $count = [BitConverter]::ToUInt16($d, $pos - 4)
    for ($i = 0; $i -lt $count; $i++) { $d[$pos + 1] = $d[$pos + 1] -band 0xEF; $pos += [BitConverter]::ToUInt16($d, $pos + 2) }
    $Security.HasExplicit = $count -gt 0
    $Security.HasInherited = $false
    return $Security
}

'@
        $rootPairs = @(
            'function Get-ShareRootRereadPath($Target) {', ($stripCode.Replace('__AUTO__', '-band 0xFB') + 'function Get-ShareRootRereadPath($Target) {'),
            '[regex]::IsMatch($Target.Real, ''^\\\\[^\\]+\\[^\\]+\\?$'')', '[regex]::IsMatch($Target.Real, ''^\\\\'')',
            '    if ($full -match ''^[A-Za-z]:$'' -or $full -match ''^\\\\[^\\]+\\[^\\]+$'') { return $null }', '    if ($full -match ''^[A-Za-z]:$'' -or $full -match ''^\\\\'') { return $null }',
            '            $security = [WbFolderPermNative]::ReadSecurity($long)', '            $security = [WbFolderPermNative]::ReadSecurity($long); if ($level -eq 0) { $security = Remove-InheritedMark $security }'
        )
        $plainUnc = ConvertTo-AdminShare '127.0.0.1' "$data\plain"
        $run = Invoke-Plain @('-Path', $plainUnc, '-NoSave') (New-PatchedTool 'share-root' $rootPairs)
        Assert-Equal '共有のルートで継承の印が外れていたら、管理共有越しに読み直して「継承だけ」にする' ('{0}/{1}/{2}/{3}' -f $run.Code, [regex]::IsMatch($run.Log, [regex]::Escape($plainUnc) + '\s+継承だけ'), $run.Log.Contains('アクセス権を変えてあるフォルダ 0 件'), $run.Log.Contains('※ 共有フォルダのフォルダで')) '0/True/True/False'
        $unsureTool = New-PatchedTool 'share-root-unsure' ($rootPairs + @('            if ($sharePath) { $target.ServerPath = $sharePath.TrimEnd(''\'') + $restPath }', ''))
        $run = Invoke-Plain @('-Path', $plainUnc, '-NoSave') $unsureTool
        Assert-Equal 'サーバーの上の場所が分からず読み直せないときは、「このフォルダ」のまま出して、見分けられなかったことを表示する' ('{0}/{1}/{2}' -f $run.Code, [regex]::IsMatch($run.Log, [regex]::Escape($plainUnc) + '\s+追加している'), $run.Log.Contains(('※ 共有フォルダのフォルダで、親から継承した設定を見分けられなかったものがあります（{0}）' -f $plainUnc))) '0/True/True'
        # 自動継承の印がある DACL（新しい形）なら、継承の印のない設定は本当にこのフォルダで設定したもの。読み直さず、見分けられなかったとも出さない
        $autoPairs = @($rootPairs[0], ($stripCode.Replace('__AUTO__', '-bor 0x04') + $rootPairs[0])) + @($rootPairs[2..($rootPairs.Count - 1)]) + @('            if ($sharePath) { $target.ServerPath = $sharePath.TrimEnd(''\'') + $restPath }', '')
        $run = Invoke-Plain @('-Path', $plainUnc, '-NoSave') (New-PatchedTool 'share-root-auto' $autoPairs)
        Assert-Equal '自動継承の印がある共有のルートは、継承の印がなくても「見分けられなかった」と出さない' ('{0}/{1}/{2}' -f $run.Code, [regex]::IsMatch($run.Log, [regex]::Escape($plainUnc) + '\s+追加している'), $run.Log.Contains('※ 共有フォルダのフォルダで')) '0/True/False'
    } else {
        Write-Skip 'サーバーの共有フォルダ' '管理共有を開けないため'
    }
    $run = Invoke-Plain @('-Path', '\\127.0.0.1')
    Assert-Equal '\\サーバーに共有がない（管理用の共有だけ）と止まる' ('{0}/{1}' -f $run.Code, $run.Log.Contains('\\127.0.0.1 には共有フォルダがありません（C$ や ADMIN$ などの管理用の共有は含めません')) '1/True'
    $run = Invoke-Plain @('-Path', '\\127.0.0.2')
    Assert-Equal '\\サーバーにつながらないと止まる' ('{0}/{1}' -f $run.Code, $run.Log.Contains('サーバーの共有フォルダの一覧を読めません: \\127.0.0.2（')) '1/True'

    Write-Section '引数の誤り'
    $run = Invoke-Plain @('-Path', "$workRoot\nothing")
    Assert-Equal 'フォルダが見つからない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("フォルダが見つかりません: $workRoot\nothing")) '1/True'
    $run = Invoke-Plain @('-Path', "$data\files\a.txt")
    Assert-Equal 'ファイルを指定した' ('{0}/{1}' -f $run.Code, $run.Log.Contains('フォルダを指定してください（ファイルは指定できません）')) '1/True'
    foreach ($bad in '0', '101', 'abc') {
        $run = Invoke-Plain @('-Path', $data, '-Depth', $bad)
        Assert-Equal "-Depth $bad" ('{0}/{1}' -f $run.Code, $run.Log.Contains("-Depth は 1～100 の数値で指定してください: $bad")) '1/True'
    }
    $run = Invoke-Plain @('-Path', $data, '-Bogus', 'x')
    Assert-Equal 'どのパラメーターにも当たらない引数' ('{0}/{1}' -f $run.Code, $run.Log.Contains('どのパラメーターにも当たらない引数があります: -Bogus x')) '1/True'

    Write-Section '保存先'
    $blocker = Join-Path $workRoot 'blocker.txt'
    [IO.File]::WriteAllText($blocker, 'x')
    $lines = @(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path "$data\files" -OutputFolder $blocker 2>&1) | ForEach-Object { "$_" }
    $log = ConvertTo-FlatText ($lines -join "`n")
    Assert-Equal '保存先に書き込めないときは、警告して画面表示だけ（終了コード 0）' ('{0}/{1}/{2}' -f $LASTEXITCODE, $log.Contains('保存先に書き込めないため、結果は保存しません'), $log.Contains('■ アクセス権を変えてあるフォルダ')) '0/True/True'
    $saveFail = New-PatchedTool 'save-fail' @('                $aceWriter.WriteLine((ConvertTo-CsvLine $AceColumns))', '                throw [IO.IOException]::new(''テストのため保存に失敗'')')
    $run = Invoke-Tool 'save-fail' @('-Path', "$data\files") $saveFail
    Assert-Equal '保存の途中で失敗したら、書きかけの結果を消して警告する' ('{0}/{1}/{2}' -f $run.Code, [bool]$run.Dir, $run.Log.Contains('結果を保存できませんでした: テストのため保存に失敗（すべてを画面で見るときは、-NoSave を付けて実行し直してください）')) '0/False/True'
    $envOut = Join-Path $workRoot 'out\env'
    $lines = @(& $Exe -NoProfile -ExecutionPolicy Bypass -Command ('$env:GFPTEST = ''{0}''; & ''{1}'' -Path ''%GFPTEST%\files'' -OutputFolder ''%GFPTEST%\..\out\env''; exit $LASTEXITCODE' -f $data, $tool) 2>&1) | ForEach-Object { "$_" }
    Assert-Equal '-Path と -OutputFolder の環境変数を展開する' ('{0}/{1}' -f $LASTEXITCODE, [bool](Get-RunDir $envOut)) '0/True'

    Write-Section '本物では起こせない条件（ツールを 1 か所だけ書き換えたコピー）'
    $aclFail = New-PatchedTool 'acl-fail' @('            $security = [WbFolderPermNative]::ReadSecurity($long)', '            $security = [WbFolderPermNative]::ReadSecurity($long); if ($long.EndsWith(''\plain'')) { $security = [WbFolderPermSecurity]::new(); $security.Error = 5 }')
    $run = Invoke-Tool 'acl-fail' @('-Path', $data) $aclFail
    $y = Read-Result $run.Dir
    Assert-Equal 'アクセス権を読めないフォルダは「権限なし」で残し、中は調べる' ('{0}/{1}' -f (Get-Rows $y.Skipped $data 'plain' 'パス')[0].'詳細', [regex]::IsMatch($run.Log, 'フォルダ 35 件')) 'アクセス権を読めません（アクセスが拒否されました）/True'
    $lookupFail = New-PatchedTool 'lookup-fail' @('        int error = Lookup(system, sid, out name, out domain, out use);', ('        int error; if (text == "{0}") {{ error = 1722; name = null; domain = null; use = 0; }} else {{ error = Lookup(system, sid, out name, out domain, out use); }}' -f $orphanSid))
    $run = Invoke-Tool 'lookup-fail' @('-Path', "$data\orphan") $lookupFail
    $y = Read-Result $run.Dir
    $row = @(Get-Rows $y.Aces $data 'orphan' 'パス' | Where-Object { $_.'SID' -eq $orphanSid })[0]
    Assert-Equal '名前を調べられない（削除されたとは限らない）: 分けて、印を付けず、案内を出す' ('{0}|{1}|{2}' -f $row.'アカウントの種類', $row.'注意', $run.Log.Contains('※ アカウントの名前を調べられなかったものがあります')) '名前を調べられない||True'
    $fat = New-PatchedTool 'fat' @('    if ([WbFolderPermNative]::GetVolumeAcl($root, [ref]$fileSystem, [ref]$hasAcl) -ne 0) { return }', '    $fileSystem = ''FAT32''; $hasAcl = $false')
    $run = Invoke-Plain @('-Path', $data) $fat
    Assert-Equal 'アクセス権のないファイルシステムでは止まる' ('{0}/{1}' -f $run.Code, $run.Log.Contains("このドライブ（FAT32）には、アクセス権がありません: $data")) '1/True'
    $noAddType = New-PatchedTool 'no-addtype' @('    if (-not (''WbFolderPermNative'' -as [type])) { Add-Type', '    throw ''テストのため Add-Type を使えない''; if (-not (''WbFolderPermNative'' -as [type])) { Add-Type')
    $run = Invoke-Plain @('-Path', $data) $noAddType
    Assert-Equal 'Add-Type を使えないと止まる' ('{0}/{1}' -f $run.Code, $run.Log.Contains('Windows の API を呼ぶ準備ができません（Add-Type を使えません）')) '1/True'

    Write-Section 'Ctrl+C（調べている途中で止めると、結果を保存しない）'
    $slow = New-PatchedTool 'slow' @('            $folderCount++', '            $folderCount++; Start-Sleep -Milliseconds 100')
    $stopped = $false
    for ($attempt = 1; $attempt -le 3 -and -not $stopped; $attempt++) {
        $stopOut = Join-Path $workRoot "out\stop$attempt"
        $ps = [powershell]::Create()
        [void]$ps.AddCommand($slow).AddParameter('Path', $data).AddParameter('OutputFolder', $stopOut)
        $handle = $ps.BeginInvoke()
        $watch = [Diagnostics.Stopwatch]::StartNew()
        while (-not (Get-RunDir $stopOut) -and -not $handle.IsCompleted -and $watch.Elapsed.TotalSeconds -lt 30) { Start-Sleep -Milliseconds 50 }
        Start-Sleep -Milliseconds 300
        $ps.Stop()
        $stopped = -not (Get-RunDir $stopOut)
        $ps.Dispose()
    }
    Assert-Equal '止めると、作りかけの結果のフォルダを消す' $stopped 'True'

    Write-Section 'コマンドプロンプトから起動'
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}\"' -f $cmdBase, "$data\files")
    Assert-Equal '最後の引数が "パス\"' ('{0}/{1}' -f $run.Code, $run.Log.Contains("対象: $data\files")) '0/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -All -NoSave' -f $cmdBase, $data)
    Assert-Equal '"パス\" の後ろに引数: ヒントを出して止まる（%ERRORLEVEL% は 1）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-Path の値に引用符（"）が残っています')) '1/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -Exclude "a b"' -f $cmdBase, $data)
    Assert-Equal '"パス\" の後ろに空白を含む値: 引用符のヒント' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-Path の値に引用符（"）が残っています')) '1/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -Dpth 2 -NoSave' -f $cmdBase, $data)
    Assert-Equal 'パラメーターの名前の書き間違い' ('{0}/{1}' -f $run.Code, $run.Log.Contains('どのパラメーターにも当たらない引数があります: -Dpth 2')) '1/True'
    $run = Invoke-ViaCmd ('{0} -Path ''{1}'' -Depth ''1'' -NoSave' -f $cmdBase, "$data\files")
    Assert-Equal 'cmd で '' で囲んだ値' ('{0}/{1}' -f $run.Code, $run.Log.Contains('階層: 1 階層下まで')) '0/True'

    Write-Section 'そのほかの起動'
    $cwd = Join-Path $workRoot '[cwd]'
    $null = New-Dir "$cwd\rel"
    $null = New-Dir "$workRoot\cwd\rel"
    $run = Invoke-From $cwd @('-Path', 'rel', '-NoSave')
    Assert-Equal '[ ] を含むフォルダから相対パスで起動する' ('{0}/{1}' -f $run.Code, $run.Log.Contains("対象: $cwd\rel")) '0/True'
    $comma = New-Dir "$workRoot\a,b"
    $run = Invoke-Plain @('-Path', $comma, '-NoSave')
    Assert-Equal 'カンマを含む名前のフォルダ' ('{0}/{1}' -f $run.Code, $run.Log.Contains("対象: $comma")) '0/True'
    $run = Invoke-Plain @('-Path', ('{0},{1}' -f "$data\files", "$data\creator"), '-NoSave')
    Assert-Equal 'カンマ区切りで 2 つ' ('{0}/{1}' -f $run.Code, $run.Log.Contains("対象: $data\files, $data\creator")) '0/True'
    $run = Invoke-Plain @('-Path', ('{0},{1}' -f $data, "$data\creator"), '-NoSave')
    Assert-Equal '重なった -Path は外側を 1 回だけ調べる' ([regex]::IsMatch($run.Log, 'フォルダ 35 件')) 'True'
    $letter = Get-FreeLetter
    if ($letter) {
        $null = & subst.exe "${letter}:" $data 2>&1
        $substLetter = $letter
        $run = Invoke-Plain @('-Path', "${letter}:\moved", '-NoSave')
        Assert-Equal 'SUBST のドライブを通しても、本物の親と比べる' ('{0}/{1}' -f $run.Code, [regex]::IsMatch($run.Log, '継承が親と合わない\s+1 件')) '0/True'
        $null = & subst.exe "${letter}:" /d 2>&1
        $substLetter = $null
    } else {
        Write-Skip 'SUBST' '空いているドライブ文字がないため'
    }

    Write-Section '制約付き言語モードとヘルプ'
    $clmCommand = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '{0}' -Path '{1}' -OutputFolder '{2}'; exit `$LASTEXITCODE" -f $tool, $data, "$workRoot\clm"
    $clm = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command $clmCommand 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $clmCode = $LASTEXITCODE
    Assert-Equal '制約付き言語モード: 止まる' ('{0}/{1}/{2}' -f $clmCode, $clm.Contains('制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）'), (Test-Path -LiteralPath "$workRoot\clm")) '1/True/False'
    $help = (@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command ("Get-Help '{0}' -Full | Out-String -Width 300" -f $tool) 2>&1) | ForEach-Object { "$_" }) -join "`n"
    Assert-Equal 'ヘルプ（概要・ツール名と版・著作権）' ('{0}/{1}/{2}' -f $help.Contains('共有のアクセス許可と NTFS のアクセス権の両方から一覧にします'), $help.Contains('Get-FolderPermission 1.0.0'), $help.Contains('Copyright (c) 2026 結進舎')) 'True/True/True'

    Write-Section '地域設定'
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (1234).ToString('N0')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '1,234') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $data.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $y = Read-Result (Get-RunDir $cultureOut)
        Assert-Equal 'fi-FI でも同じ結果' ((Get-RowList $y.Items $data 'パス' '注意') -eq (Join-Sorted $listNotes)) 'True'
        Assert-Equal 'fi-FI でも値の書式は同じ' (@(Get-Rows $y.Aces $data 'special' 'パス' | Where-Object { $_.'設定した場所' -eq 'このフォルダ' })[0].'値') '0x00120084'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }

    Write-Section 'レビューで見つかった問題の再発防止'
    # 1. 同じドライブの中で移動したフォルダ・ファイル、所有者を変えた子、汎用で分かれた子は「継承が親と合わないフォルダ」で確かめている
    # 2. 重なった共有・外側から届かない内側・同じフォルダの 2 つの共有名は「共有フォルダ」「サーバーの共有フォルダ」で確かめている
    # 3. 保存の途中の失敗、名前を調べられない SID は「保存先」「本物では起こせない条件」で確かめている
    # 4. 中身が 1 件だけのフォルダ（if 文の代入で配列がほどけて止まった）は「ファイル」の files\only で確かめている
    $line = @($x.Lines | Where-Object { $_ -match '^\s+許可（継承）\s+BUILTIN\\Administrators' })[0]
    Assert-Equal 'アカウントの列は 40 文字分まで（長い SID が 1 つあっても、ほかの行はずれない）' ([regex]::IsMatch($line, 'BUILTIN\\Administrators {20}フル コントロール')) 'True'
} catch {
    '  FAIL  テストが途中で止まりました: {0}（{1} 行目）' -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber
    $script:failCount++
} finally {
    if ($substLetter) { $null = & subst.exe "${substLetter}:" /d 2>&1 }
    # 拒否を外し、DACL をそのまま書いたものは全員の操作を許してから、ジャンクションを外して消す
    foreach ($entry in $denied) {
        try {
            $acl = Get-PathAcl $entry.Path
            [void]$acl.RemoveAccessRule((New-Object Security.AccessControl.FileSystemAccessRule((New-Object Security.Principal.SecurityIdentifier $entry.Sid), [Security.AccessControl.FileSystemRights]$entry.Rights, 'None', 'None', 'Deny')))
            Set-PathAcl $entry.Path $acl
        } catch { }
    }
    foreach ($path in $rawDacls) {
        try {
            $raw = New-Object Security.AccessControl.RawSecurityDescriptor 'D:NO_ACCESS_CONTROL'
            $bytes = New-Object byte[] $raw.BinaryLength
            $raw.GetBinaryForm($bytes, 0)
            [void][WbGfpTestSecurity]::SetFileSecurityW('\\?\' + $path, 4, $bytes)
        } catch { }
    }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            Start-Sleep -Milliseconds 500
            foreach ($junction in $junctions) { if ([IO.Directory]::Exists($junction)) { try { [IO.Directory]::Delete($junction) } catch { } } }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

'  END'
exit $script:failCount
