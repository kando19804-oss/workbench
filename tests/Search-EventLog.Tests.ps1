<#
.SYNOPSIS
    Search-EventLog の回帰テストです。

.DESCRIPTION
    テスト用のイベントを持つ .evtx ファイルを一時フォルダに組み立て、Search-EventLog で取り出した結果を期待値と照合します。
    本物のイベントログには書き込みません（読むだけです）。管理者でなくても、すべてのレベル（重大・詳細を含む）を確かめられます。
    メッセージの確認には、どの Windows にもあるソース WSH の情報を使います（イベントの本文がそのままメッセージになります）。
    本物のイベントログは、wevtutil で数えた件数と突き合わせます。
    「権限なし」を再現するため、テストが作ったファイルとフォルダにだけアクセス拒否を設定します（終了時に戻します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Search-EventLog を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Search-EventLog\Search-EventLog.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Search-EventLog_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Search-EventLog.ps1'
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
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
function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $out  = Join-Path $workRoot "out\$Label"
    $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = $log; Dir = (Get-RunDir $out) }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行
function Invoke-Plain([string[]]$ToolArgs, [string]$ToolPath = $tool) {
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
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

# [powershell]::Create() で動かすスクリプトを、一時フォルダに書く
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

# 定数や処理を 1 か所だけ書き換えたツールのコピーを作る（本物のログでは起きない条件を確かめるため）
function New-PatchedTool([string]$Name, [string]$From, [string]$To) {
    $text = [IO.File]::ReadAllText($repoTool)
    if (-not $text.Contains($From)) { throw "書き換える場所が見つかりません: $From" }
    $file = Join-Path $workRoot "patched\$Name\Search-EventLog.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text.Replace($From, $To), $utf8Bom)
    $file
}

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'EventLog_*' -ErrorAction SilentlyContinue | Sort-Object Name)
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
    $lines = @()
    if ($Dir -and (Test-Path -LiteralPath (Join-Path $Dir 'summary.txt'))) { $lines = @([IO.File]::ReadAllLines((Join-Path $Dir 'summary.txt'))) }
    [pscustomobject]@{ Events = @(Read-CsvOrEmpty 'events.csv'); Groups = @(Read-CsvOrEmpty 'summary.csv'); Skipped = @(Read-CsvOrEmpty 'skipped.csv'); Lines = $lines }
}

# events.csv のイベント ID を、ファイルの順（新しい順）に並べる
function Get-Ids($Rows) { (@($Rows) | ForEach-Object { $_.'イベントID' }) -join ',' }

# 結果のイベント ID を、終了コードと一緒に返す
function Get-CodeAndIds($Run) { '{0}/{1}' -f $Run.Code, (Get-Ids (Read-Result $Run.Dir).Events) }

function Get-ResultFileNames([string]$Dir) {
    if (-not $Dir) { return '' }
    (@(Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { $_.Name } | Sort-Object)) -join ','
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

#region テスト用の .evtx を組み立てる

# 本物のログに書き込むと、管理者でないとソースを作れず、重大と詳細のレベルや決まった日時のイベントも作れないため、
# ファイルを直接組み立てる。形式は [MS-EVEN6] の BinXml で、1 つのチャンク（64 KB）に 1 つのテンプレートを置く

$crcTable = New-Object 'uint32[]' 256
foreach ($n in 0..255) {
    [uint32]$c = $n
    foreach ($k in 1..8) { if ($c -band 1) { $c = [uint32]3988292384 -bxor ($c -shr 1) } else { $c = $c -shr 1 } }
    $crcTable[$n] = $c
}

function Get-Crc32([byte[]]$Data) {
    [uint32]$crc = 4294967295
    foreach ($b in $Data) { $crc = $crcTable[($crc -bxor $b) -band 255] -bxor ($crc -shr 8) }
    [uint32]($crc -bxor [uint32]4294967295)
}

function Set-Int32At($Buf, [int]$Offset, [int]$Value) {
    $bytes = [BitConverter]::GetBytes($Value)
    for ($i = 0; $i -lt 4; $i++) { $Buf[$Offset + $i] = $bytes[$i] }
}

# 要素名・属性名。オフセットのすぐ後に名前の本体を置く
function Add-XmlName($Buf, [int]$Base, [string]$Name) {
    $Buf.AddRange([BitConverter]::GetBytes([int]($Base + $Buf.Count + 4)))
    $Buf.AddRange([BitConverter]::GetBytes([int]0))
    [uint32]$hash = 0
    foreach ($ch in $Name.ToCharArray()) { $hash = [uint32](([uint64]$hash * 65599 + [int]$ch) -band 4294967295) }
    $Buf.AddRange([BitConverter]::GetBytes([uint16]($hash -band 65535)))
    $Buf.AddRange([BitConverter]::GetBytes([uint16]$Name.Length))
    $Buf.AddRange([Text.Encoding]::Unicode.GetBytes($Name + [char]0))
}

# 要素。Attr は @(名前, 値の番号, 型) か @(名前, 文字列)。Value は @(値の番号, 型)
function Add-XmlElement($Buf, [int]$Base, [string]$Name, $Attr, $Value, [scriptblock]$Children) {
    $Buf.Add([byte]$(if ($Attr) { 0x41 } else { 0x01 }))
    $Buf.AddRange([byte[]](0xFF, 0xFF))
    $sizeAt = $Buf.Count
    $Buf.AddRange([byte[]](0, 0, 0, 0))
    Add-XmlName $Buf $Base $Name
    if ($Attr) {
        $listAt = $Buf.Count
        $Buf.AddRange([byte[]](0, 0, 0, 0))
        $Buf.Add([byte]0x06)
        Add-XmlName $Buf $Base $Attr[0]
        if ($Attr.Count -eq 3) {
            $Buf.Add([byte]0x0E); $Buf.AddRange([BitConverter]::GetBytes([uint16]$Attr[1])); $Buf.Add([byte]$Attr[2])
        } else {
            $Buf.AddRange([byte[]](0x05, 0x01)); $Buf.AddRange([BitConverter]::GetBytes([uint16]$Attr[1].Length)); $Buf.AddRange([Text.Encoding]::Unicode.GetBytes($Attr[1]))
        }
        Set-Int32At $Buf $listAt ($Buf.Count - $listAt - 4)
    }
    if ($Value -or $Children) {
        $Buf.Add([byte]0x02)
        if ($Value) { $Buf.Add([byte]0x0E); $Buf.AddRange([BitConverter]::GetBytes([uint16]$Value[0])); $Buf.Add([byte]$Value[1]) }
        if ($Children) { & $Children }
        $Buf.Add([byte]0x04)
    } else {
        $Buf.Add([byte]0x03)
    }
    Set-Int32At $Buf $sizeAt ($Buf.Count - $sizeAt - 4)
}

# チャンクの最初のレコードにテンプレートの本体を置き、後のレコードはその位置を指す
$templateAt = 550
$templateId = [byte[]](0x57, 0x42, 0x54, 0x31)

function New-EvtxRecord($Event, [long]$RecordId, [int]$Start) {
    $rec = New-Object 'System.Collections.Generic.List[byte]'
    $fileTime = ([datetime]$Event.Time).ToUniversalTime().ToFileTimeUtc()
    $rec.AddRange([byte[]](0x2A, 0x2A, 0, 0, 0, 0, 0, 0))
    $rec.AddRange([BitConverter]::GetBytes([int64]$RecordId))
    $rec.AddRange([BitConverter]::GetBytes([int64]$fileTime))
    $rec.AddRange([byte[]](0x0F, 1, 1, 0, 0x0C, 1))
    $rec.AddRange($templateId)
    $rec.AddRange([BitConverter]::GetBytes([int]$templateAt))
    if ($Start -eq 512) {
        $rec.AddRange([byte[]](0, 0, 0, 0))
        $rec.AddRange($templateId)
        $rec.AddRange((New-Object byte[] 12))
        $sizeAt = $rec.Count
        $rec.AddRange([byte[]](0, 0, 0, 0))
        $rec.AddRange([byte[]](0x0F, 1, 1, 0))
        Add-XmlElement $rec $Start 'Event' @('xmlns', 'http://schemas.microsoft.com/win/2004/08/events/event') $null {
            Add-XmlElement $rec $Start 'System' $null $null {
                Add-XmlElement $rec $Start 'Provider' @('Name', 0, 0x01) $null $null
                Add-XmlElement $rec $Start 'EventID' @('Qualifiers', 13, 0x06) @(1, 0x06) $null
                Add-XmlElement $rec $Start 'Version' $null @(2, 0x04) $null
                Add-XmlElement $rec $Start 'Level' $null @(3, 0x04) $null
                Add-XmlElement $rec $Start 'Task' $null @(4, 0x06) $null
                Add-XmlElement $rec $Start 'Opcode' $null @(5, 0x04) $null
                Add-XmlElement $rec $Start 'Keywords' $null @(6, 0x15) $null
                Add-XmlElement $rec $Start 'TimeCreated' @('SystemTime', 7, 0x11) $null $null
                Add-XmlElement $rec $Start 'EventRecordID' $null @(8, 0x0A) $null
                Add-XmlElement $rec $Start 'Channel' $null @(9, 0x01) $null
                Add-XmlElement $rec $Start 'Computer' $null @(10, 0x01) $null
                Add-XmlElement $rec $Start 'Security' @('UserID', 11, 0x13) $null $null
            }
            Add-XmlElement $rec $Start 'EventData' $null $null {
                Add-XmlElement $rec $Start 'Data' $null @(12, 0x01) $null
            }
        }
        $rec.Add([byte]0x00)
        Set-Int32At $rec $sizeAt ($rec.Count - $sizeAt - 4)
    }
    $sidBytes = [byte[]]@()
    if ($Event.Sid) {
        $sid = New-Object Security.Principal.SecurityIdentifier ([string]$Event.Sid)
        $sidBytes = New-Object byte[] $sid.BinaryLength
        $sid.GetBinaryForm($sidBytes, 0)
    }
    $values = @(
        @(0x01, [Text.Encoding]::Unicode.GetBytes([string]$Event.Provider)),
        @(0x06, [BitConverter]::GetBytes([uint16]$Event.Id)),
        @(0x04, [byte[]]@(0)),
        @(0x04, [byte[]]@([byte]$Event.Level)),
        @(0x06, [BitConverter]::GetBytes([uint16]0)),
        @(0x04, [byte[]]@(0)),
        @(0x15, [BitConverter]::GetBytes([uint64]36028797018963968)),
        @(0x11, [BitConverter]::GetBytes([int64]$fileTime)),
        @(0x0A, [BitConverter]::GetBytes([uint64]$RecordId)),
        @(0x01, [Text.Encoding]::Unicode.GetBytes([string]$Event.Channel)),
        @(0x01, [Text.Encoding]::Unicode.GetBytes([string]$Event.Computer)),
        @(0x13, $sidBytes),
        @(0x01, [Text.Encoding]::Unicode.GetBytes([string]$Event.Data)),
        @(0x06, [BitConverter]::GetBytes([uint16]$Event.Qualifiers))
    )
    $rec.AddRange([BitConverter]::GetBytes([int]$values.Count))
    foreach ($v in $values) {
        $rec.AddRange([BitConverter]::GetBytes([uint16]$v[1].Length))
        $rec.Add([byte]$(if ($v[1].Length) { $v[0] } else { 0 }))
        $rec.Add([byte]0)
    }
    foreach ($v in $values) { if ($v[1].Length) { $rec.AddRange([byte[]]$v[1]) } }
    $rec.Add([byte]0x00)
    while (($rec.Count + 4) % 8) { $rec.Add([byte]0) }
    $total = $rec.Count + 4
    $rec.AddRange([BitConverter]::GetBytes([int]$total))
    Set-Int32At $rec 4 $total
    , $rec
}

function New-EvtxChunk([object[]]$Records, [long]$FirstId, [int[]]$Offsets, [int]$FreeOffset) {
    $bytes = New-Object byte[] 65536
    foreach ($i in 0..($Records.Count - 1)) { $Records[$i].CopyTo($bytes, $Offsets[$i]) }
    $lastId = $FirstId + $Records.Count - 1
    [Text.Encoding]::ASCII.GetBytes("ElfChnk`0").CopyTo($bytes, 0)
    [BitConverter]::GetBytes([int64]$FirstId).CopyTo($bytes, 8)
    [BitConverter]::GetBytes([int64]$lastId).CopyTo($bytes, 16)
    [BitConverter]::GetBytes([int64]$FirstId).CopyTo($bytes, 24)
    [BitConverter]::GetBytes([int64]$lastId).CopyTo($bytes, 32)
    [BitConverter]::GetBytes([int]128).CopyTo($bytes, 40)
    [BitConverter]::GetBytes([int]$Offsets[$Offsets.Count - 1]).CopyTo($bytes, 44)
    [BitConverter]::GetBytes([int]$FreeOffset).CopyTo($bytes, 48)
    $body = New-Object byte[] ($FreeOffset - 512)
    [Array]::Copy($bytes, 512, $body, 0, $body.Length)
    [BitConverter]::GetBytes([uint32](Get-Crc32 $body)).CopyTo($bytes, 52)
    [BitConverter]::GetBytes([int]1).CopyTo($bytes, 120)
    $head = New-Object byte[] 504
    [Array]::Copy($bytes, 0, $head, 0, 120)
    [Array]::Copy($bytes, 128, $head, 120, 384)
    [BitConverter]::GetBytes([uint32](Get-Crc32 $head)).CopyTo($bytes, 124)
    , $bytes
}

function New-EvtxFile([string]$File, [object[]]$Events) {
    $chunks  = New-Object 'System.Collections.Generic.List[object]'
    $records = New-Object 'System.Collections.Generic.List[object]'
    $offsets = New-Object 'System.Collections.Generic.List[int]'
    $free    = 512
    $firstId = 1L
    $id      = 0L
    foreach ($ev in $Events) {
        $id++
        $rec = New-EvtxRecord $ev $id $free
        if ($free + $rec.Count -gt 65536) {
            $chunks.Add((New-EvtxChunk $records.ToArray() $firstId $offsets.ToArray() $free))
            $records.Clear(); $offsets.Clear()
            $free = 512
            $firstId = $id
            $rec = New-EvtxRecord $ev $id $free
        }
        $records.Add($rec.ToArray())
        $offsets.Add($free)
        $free += $rec.Count
    }
    $chunks.Add((New-EvtxChunk $records.ToArray() $firstId $offsets.ToArray() $free))
    $head = New-Object byte[] 4096
    [Text.Encoding]::ASCII.GetBytes("ElfFile`0").CopyTo($head, 0)
    [BitConverter]::GetBytes([int64]($chunks.Count - 1)).CopyTo($head, 16)
    [BitConverter]::GetBytes([int64]($id + 1)).CopyTo($head, 24)
    [BitConverter]::GetBytes([int]128).CopyTo($head, 32)
    [BitConverter]::GetBytes([uint16]2).CopyTo($head, 36)
    [BitConverter]::GetBytes([uint16]3).CopyTo($head, 38)
    [BitConverter]::GetBytes([uint16]4096).CopyTo($head, 40)
    [BitConverter]::GetBytes([uint16]$chunks.Count).CopyTo($head, 42)
    $checked = New-Object byte[] 120
    [Array]::Copy($head, 0, $checked, 0, 120)
    [BitConverter]::GetBytes([uint32](Get-Crc32 $checked)).CopyTo($head, 124)
    [void][IO.Directory]::CreateDirectory('\\?\' + (Split-Path -Parent $File))
    $stream = [IO.File]::Create('\\?\' + $File)
    try {
        $stream.Write($head, 0, 4096)
        foreach ($chunk in $chunks) { $stream.Write($chunk, 0, 65536) }
    } finally {
        $stream.Dispose()
    }
}

# WSH の情報には、この組み合わせの ID のメッセージだけがある（本文がそのままメッセージになる）
$wshQualifiers = @{ 0 = 0; 1 = 49152; 2 = 32768; 4 = 16384 }

function New-TestEvent([datetime]$Time, [int]$Level, [int]$Id, [string]$Provider = 'wb-other', [string]$Data = '', [string]$Channel = 'System', [string]$Sid = '', [string]$Computer = 'SERVER01') {
    $qualifiers = if ($Provider -eq 'WSH') { $wshQualifiers[$Id] } else { 0 }
    @{ Time = $Time; Level = $Level; Id = $Id; Provider = $Provider; Data = $Data; Channel = $Channel; Sid = $Sid; Computer = $Computer; Qualifiers = $qualifiers }
}

function Get-TestTime([string]$HourMinute) { [datetime]::ParseExact("2026-06-15 $HourMinute", 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture) }

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$userSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied  = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したパス（終了時に戻す）
$substDrive = $null
[void][IO.Directory]::CreateDirectory((Split-Path -Parent $tool))
[IO.File]::Copy($repoTool, $tool)

try {
    Write-Section '準備: テスト用の .evtx を作る'
    $range = @('-Since', '2026-06-15 09:00', '-Until', '2026-06-15 11:00')
    $allLevels = 'Critical,Error,Warning,Information,Verbose'
    $mainEvents = @(
        (New-TestEvent (Get-TestTime '10:00') 1 9001 'Microsoft-Windows-Kernel-Power'),
        (New-TestEvent (Get-TestTime '10:01') 2 1 'WSH' 'エラー Timeout 1' 'Application'),
        (New-TestEvent (Get-TestTime '10:02') 2 1 'WSH' 'エラー Timeout 2' 'Application'),
        (New-TestEvent (Get-TestTime '10:03') 2 1 'WSH' "エラー Timeout 3`r`n2 行目`r`n3 行目" 'Application'),
        (New-TestEvent (Get-TestTime '10:04') 2 9002 'wb-other' '' 'System' '' 'SERVER02'),
        (New-TestEvent (Get-TestTime '10:05') 3 2 'WSH' '警告 メッセージ' 'Application'),
        (New-TestEvent (Get-TestTime '10:06') 3 2 'WSH' '警告 メッセージ' 'Application'),
        (New-TestEvent (Get-TestTime '10:07') 4 4 'WSH' '情報 メッセージ' 'Application'),
        (New-TestEvent (Get-TestTime '10:08') 0 0 'WSH' 'レベル 0 のメッセージ' 'Security'),
        (New-TestEvent (Get-TestTime '10:09') 5 4 'WSH' '詳細 メッセージ' 'Application'),
        (New-TestEvent (Get-TestTime '10:10') 2 0 'WSH' '=1+1 "quoted"' 'Application'),
        (New-TestEvent (Get-TestTime '10:11') 2 9003 'wb-other' '' 'System' 'S-1-5-18'),
        (New-TestEvent (Get-TestTime '10:12') 2 9004 'wb-other' '' 'System' 'S-1-5-21-1-2-3-1000')
    )
    $main = Join-Path $workRoot 'evtx\main.evtx'
    New-EvtxFile $main $mainEvents
    $check = @(Get-WinEvent -Path $main -Oldest -ErrorAction SilentlyContinue)
    Assert-Equal '作った .evtx を Windows が読める' $check.Count $mainEvents.Count
    Assert-Equal 'レベルと日時も読める' ('{0}/{1:yyyy-MM-dd HH:mm}' -f $check[0].Level, $check[0].TimeCreated) '1/2026-06-15 10:00'
    $wshReady = (@($check | Where-Object { $_.Id -eq 1 })[0].Message) -eq 'エラー Timeout 1'
    if (-not $wshReady) { Write-Skip 'メッセージの確認' 'ソース WSH の情報がないため' }
    $mainArgs = @('-Path', $main) + $range

    Write-Section '既定のレベル（重大・エラー）'
    $run = Invoke-Tool 'default' $mainArgs
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '保存先のフォルダの名前' ((Split-Path -Leaf $run.Dir) -replace '\d{8}_\d{6}$', 'DATE') ('EventLog_{0}_DATE' -f $env:COMPUTERNAME)
    Assert-Equal '保存するファイル' (Get-ResultFileNames $run.Dir) 'events.csv,skipped.csv,summary.csv,summary.txt'
    Assert-Equal '新しい順の ID' (Get-Ids $x.Events) '9004,9003,0,9002,1,1,1,9001'
    Assert-Equal '日時の形' (@($x.Events)[0].'日時') '2026/06/15 10:12:00'
    Assert-Equal '新しい順の日時' ((@($x.Events) | ForEach-Object { $_.'日時'.Substring(11, 5) }) -join ',') '10:12,10:11,10:10,10:04,10:03,10:02,10:01,10:00'
    Assert-Equal 'レベルの名前' ((@($x.Events) | ForEach-Object { $_.'レベル' } | Select-Object -Unique) -join ',') 'エラー,重大'
    Assert-Equal 'ログの名前' ((@($x.Events) | ForEach-Object { $_.'ログ' } | Select-Object -Unique) -join ',') 'System,Application'
    Assert-Equal 'ソースは Microsoft-Windows- を省く' (@($x.Events)[-1].'ソース') 'Kernel-Power'
    Assert-Equal 'コンピューター' ((@($x.Events) | Where-Object { $_.'イベントID' -eq '9002' }).'コンピューター') 'SERVER02'
    Assert-Equal 'タスクのカテゴリ（ソースの情報がないと空欄）' ((@($x.Events) | Where-Object { $_.'イベントID' -eq '9002' }).'タスクのカテゴリ') ''
    Assert-Equal 'ユーザー: なければ空欄' ((@($x.Events) | Where-Object { $_.'イベントID' -eq '9002' }).'ユーザー') ''
    Assert-Equal 'ユーザー: 名前に直す' ((@($x.Events) | Where-Object { $_.'イベントID' -eq '9003' }).'ユーザー' -match '\\SYSTEM$') 'True'
    Assert-Equal 'ユーザー: 名前に直せない SID はそのまま' ((@($x.Events) | Where-Object { $_.'イベントID' -eq '9004' }).'ユーザー') 'S-1-5-21-1-2-3-1000'
    Assert-Equal 'メッセージがないイベント' ((@($x.Events) | Where-Object { $_.'イベントID' -eq '9002' }).'メッセージ') '（メッセージを取得できません）'
    Assert-Equal '件数の行' ([regex]::Match($run.Log, 'イベント \S+ 件（[^）]*）  種類 \S+').Value) 'イベント 8 件（重大 1・エラー 7）  種類 6'
    Assert-Equal '対象の行' $run.Log.Contains(('対象: {0}期間' -f $main)) 'True'
    Assert-Equal '期間の行' $run.Log.Contains('期間: 2026-06-15 09:00 ～ 2026-06-15 11:00   レベル: 重大, エラー') 'True'
    Assert-Equal '条件がないときは条件の行を出さない' $run.Log.Contains('条件:') 'False'
    Assert-Equal '集計の見出し' $run.Log.Contains('■ 集計（件数の多い順・上位 30）') 'True'
    Assert-Equal '新しい順の見出し' $run.Log.Contains('■ 新しい順（20 件まで）') 'True'
    Assert-Equal '保存先の行' $run.Log.Contains(('保存先: {0}' -f $run.Dir)) 'True'
    Assert-Equal 'スキップがないときはスキップの行を出さない' $run.Log.Contains('スキップ') 'False'
    Assert-Equal 'skipped.csv は見出しだけ' @($x.Skipped).Count 0

    Write-Section 'CSV とテキストの形'
    foreach ($case in @(
        @('events.csv', '"日時","ログ","レベル","イベントID","ソース","タスクのカテゴリ","コンピューター","ユーザー","メッセージ"'),
        @('summary.csv', '"件数","レベル","ログ","ソース","イベントID","最初に発生","最後に発生","メッセージ（1 行目）"'),
        @('skipped.csv', '"理由","対象","詳細"')
    )) {
        Assert-Equal ('{0} の見出し' -f $case[0]) ([IO.File]::ReadAllLines((Join-Path $run.Dir $case[0]))[0]) $case[1]
    }
    foreach ($name in 'events.csv', 'summary.csv', 'skipped.csv', 'summary.txt') {
        $head = New-Object byte[] 3
        $stream = [IO.File]::OpenRead((Join-Path $run.Dir $name))
        try { [void]$stream.Read($head, 0, 3) } finally { $stream.Dispose() }
        Assert-Equal ('{0} は BOM 付き UTF-8' -f $name) ($head -join ',') '239,187,191'
    }
    $summaryText = $x.Lines -join "`n"
    Assert-Equal 'summary.txt の先頭' $x.Lines[0] 'イベントログ抽出 (Search-EventLog 1.0.0)'
    Assert-Equal 'summary.txt の集計' $summaryText.Contains('■ 集計（件数の多い順）') 'True'
    Assert-Equal 'summary.txt の件数の行' $summaryText.Contains('イベント 8 件（重大 1・エラー 7）') 'True'
    Assert-Equal 'summary.txt の保存先' $x.Lines[-1] ('保存先: {0}' -f $run.Dir)

    Write-Section '集計'
    $top = @($x.Groups)[0]
    Assert-Equal '1 行目は件数の多いもの' ('{0}/{1}/{2}/{3}/{4}' -f $top.'件数', $top.'レベル', $top.'ログ', $top.'ソース', $top.'イベントID') '3/エラー/Application/WSH/1'
    Assert-Equal '最初に発生・最後に発生' ('{0}/{1}' -f $top.'最初に発生', $top.'最後に発生') '2026/06/15 10:01:00/2026/06/15 10:03:00'
    Assert-Equal '件数が同じなら最後に発生が新しい順' ((@($x.Groups) | ForEach-Object { $_.'イベントID' }) -join ',') '1,9004,9003,0,9002,9001'
    if ($wshReady) { Assert-Equal '集計のメッセージは、いちばん新しいイベントの 1 行目' $top.'メッセージ（1 行目）' 'エラー Timeout 3' }
    # ログ・レベル・ソース・イベント ID のどれかが違えば、別の行にする
    $groupFile = Join-Path $workRoot 'groups\groups.evtx'
    New-EvtxFile $groupFile @(
        (New-TestEvent (Get-TestTime '10:00') 2 500 'wb-a' '' 'System'),
        (New-TestEvent (Get-TestTime '10:01') 2 500 'wb-a' '' 'System'),
        (New-TestEvent (Get-TestTime '10:02') 2 500 'wb-a' '' 'Application'),
        (New-TestEvent (Get-TestTime '10:03') 3 500 'wb-a' '' 'System'),
        (New-TestEvent (Get-TestTime '10:04') 2 500 'wb-b' '' 'System'),
        (New-TestEvent (Get-TestTime '10:05') 2 501 'wb-a' '' 'System')
    )
    $run = Invoke-Tool 'groups' (@('-Path', $groupFile, '-Level', $allLevels) + $range)
    $rows = @((Read-Result $run.Dir).Groups)
    Assert-Equal 'ログ・レベル・ソース・ID のどれかが違えば別の行' (($rows | ForEach-Object { '{0}/{1}/{2}/{3}/{4}' -f $_.'件数', $_.'ログ', $_.'レベル', $_.'ソース', $_.'イベントID' }) -join ',') '2/System/エラー/wb-a/500,1/System/エラー/wb-a/501,1/System/エラー/wb-b/500,1/System/警告/wb-a/500,1/Application/エラー/wb-a/500'

    Write-Section 'レベル'
    $levelCases = @(
        @('Critical', 'Critical', '9001'),
        @('Error', 'Error', '9004,9003,0,9002,1,1,1'),
        @('Warning', 'Warning', '2,2'),
        @('Information（レベル 0 を含む）', 'Information', '0,4'),
        @('Verbose', 'Verbose', '4'),
        @('日本語', '重大,エラー', '9004,9003,0,9002,1,1,1,9001'),
        @('日本語（残りの 3 つ）', '警告,情報,詳細', '4,0,4,2,2'),
        @('大文字小文字と空白', 'critical, WARNING', '2,2,9001'),
        @('同じレベルを 2 回', 'Information,情報', '0,4'),
        @('すべて', $allLevels, '9004,9003,0,4,0,4,2,2,9002,1,1,1,9001')
    )
    foreach ($case in $levelCases) {
        $run = Invoke-Tool ('level' + [Array]::IndexOf($levelCases, $case)) ($mainArgs + @('-Level', $case[1]))
        Assert-Equal ('-Level {0}' -f $case[0]) (Get-CodeAndIds $run) "0/$($case[2])"
    }
    $x = Read-Result $run.Dir
    Assert-Equal 'すべてのレベル: 件数の行' ([regex]::Match($run.Log, 'イベント \S+ 件（[^）]*）  種類 \S+').Value) 'イベント 13 件（重大 1・エラー 7・警告 2・情報 2・詳細 1）  種類 10'
    Assert-Equal 'すべてのレベル: 期間の行' $run.Log.Contains('レベル: 重大, エラー, 警告, 情報, 詳細') 'True'
    Assert-Equal 'レベル 0 は「情報」' ((@($x.Events) | Where-Object { $_.'ログ' -eq 'Security' }).'レベル') '情報'
    Assert-Equal 'レベルの名前（新しい順）' ((@($x.Events) | ForEach-Object { $_.'レベル' }) -join ',') 'エラー,エラー,エラー,詳細,情報,情報,警告,警告,エラー,エラー,エラー,エラー,重大'
    $run = Invoke-Tool 'levelinfo' ($mainArgs + @('-Level', 'Information'))
    Assert-Equal '情報だけ: 期間の行と件数の行' ('{0}/{1}' -f $run.Log.Contains('レベル: 情報'), $run.Log.Contains('イベント 2 件（情報 2）  種類 2')) 'True/True'

    Write-Section '画面の色'
    $colorScript = New-HelperScript 'colors.ps1' @(
        'param($Tool, $File)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $File).AddParameter(''Since'', ''2026-06-15 09:00'').AddParameter(''Until'', ''2026-06-15 11:00'').AddParameter(''Level'', ''Critical,Error,Warning'').AddParameter(''NoSave'', $true)'
        '[void]$ps.Invoke()'
        'foreach ($record in $ps.Streams.Information) {'
        '    $m = $record.MessageData'
        '    if ($m.Message -match ''^ .*(重大|エラー|警告)  '') { ''{0}_{1}={2}'' -f $Matches[1], $(if ($m.Message -match ''^ \d{4}-'') { ''一覧'' } else { ''集計'' }), $m.ForegroundColor }'
        '}'
    )
    $colors = Invoke-Helper $colorScript @('-Tool', $tool, '-File', $main)
    Assert-Equal '重大は赤' ('{0}/{1}' -f $colors['重大_集計'], $colors['重大_一覧']) 'Red/Red'
    Assert-Equal 'エラーは黄色' ('{0}/{1}' -f $colors['エラー_集計'], $colors['エラー_一覧']) 'Yellow/Yellow'
    Assert-Equal '警告は色なし' ('{0}/{1}' -f $colors['警告_集計'], $colors['警告_一覧']) '/'

    Write-Section '期間'
    $rangeCases = @(
        @('境界の日時を含む', '2026-06-15 10:03', '2026-06-15 10:05', '2,9002,1'),
        @('秒まで', '2026-06-15 10:03:00', '2026-06-15 10:03:59', '1'),
        @('秒まで: 1 秒後は含まない', '2026-06-15 10:03:01', '2026-06-15 10:03:59', ''),
        @('/ の日付', '2026/06/15 10:03', '2026/06/15 10:04', '9002,1'),
        @('月日が 1 桁', '2026-6-15 10:03', '2026-6-15 10:04', '9002,1'),
        @('時が 1 桁', '2026-6-15 9:00', '2026/6/15 10:00', '9001'),
        @('日付だけ', '2026-06-15', '2026-06-16', '9004,9003,0,4,0,4,2,2,9002,1,1,1,9001')
    )
    foreach ($case in $rangeCases) {
        $run = Invoke-Tool ('range' + [Array]::IndexOf($rangeCases, $case)) @('-Path', $main, '-Level', $allLevels, '-Since', $case[1], '-Until', $case[2])
        Assert-Equal ('-Since {0} -Until {1}: {2}' -f $case[1], $case[2], $case[0]) (Get-Ids (Read-Result $run.Dir).Events) $case[3]
    }
    $now = Get-Date
    $recent = Join-Path $workRoot 'recent\recent.evtx'
    New-EvtxFile $recent @(
        (New-TestEvent $now.AddMinutes(-30) 2 1),
        (New-TestEvent $now.AddHours(-3) 2 2),
        (New-TestEvent $now.AddHours(-30) 2 3),
        (New-TestEvent $now.AddDays(-40) 2 4)
    )
    $relativeCases = @(
        @('省略は 24 時間', @(), '1,2', '（24 時間）'),
        @('1h', @('-Since', '1h'), '1', '（1 時間）'),
        @('45m', @('-Since', '45m'), '1', '（45 分）'),
        @('4H（大文字）', @('-Since', '4H'), '1,2', '（4 時間）'),
        @('3d', @('-Since', '3d'), '1,2,3', '（3 日）'),
        @('2 d（空白）', @('-Since', '2 d'), '1,2,3', '（2 日）'),
        @('100d', @('-Since', '100d'), '1,2,3,4', '（100 日）')
    )
    foreach ($case in $relativeCases) {
        $run = Invoke-Tool ('rel' + [Array]::IndexOf($relativeCases, $case)) (@('-Path', $recent) + $case[1])
        Assert-Equal ('-Since {0}' -f $case[0]) ('{0}/{1}' -f (Get-Ids (Read-Result $run.Dir).Events), $run.Log.Contains($case[3])) "$($case[2])/True"
    }
    $run = Invoke-Tool 'relboth' @('-Path', $recent, '-Since', '4h', '-Until', '1h')
    Assert-Equal '-Since 4h -Until 1h' (Get-Ids (Read-Result $run.Dir).Events) '2'
    Assert-Equal '-Until があるときは（4 時間）を添えない' $run.Log.Contains('（4 時間）') 'False'

    Write-Section '該当するイベントがないとき'
    # Windows でのテストで見つかった問題: 該当なしが「読めるログがありませんでした」になり、終了コード 1 で止まっていた
    $run = Invoke-Tool 'nomatch' @('-Path', $main, '-Since', '2000-01-01', '-Until', '2000-01-02')
    $x = Read-Result $run.Dir
    Assert-Equal '終了コードは 2' $run.Code 2
    Assert-Equal '該当なしの表示' ('{0}/{1}' -f $run.Log.Contains('該当するイベントはありません'), $run.Log.Contains('読めるログがありません')) 'True/False'
    Assert-Equal '表を出さない' $run.Log.Contains('■') 'False'
    Assert-Equal '結果は見出しだけを保存する' ('{0}/{1}' -f (Get-ResultFileNames $run.Dir), @($x.Events).Count) 'events.csv,skipped.csv,summary.csv,summary.txt/0'
    Assert-Equal 'summary.txt の集計は（該当なし）' ($x.Lines -join "`n").Contains('    （該当なし）') 'True'
    $run = Invoke-Plain @('-Path', $main, '-Since', '2000-01-01', '-Until', '2000-01-02', '-NoSave')
    Assert-Equal '-NoSave でも終了コードは 2' ('{0}/{1}' -f $run.Code, $run.Log.Contains('該当するイベントはありません')) '2/True'

    Write-Section 'イベント ID'
    $manyIds = (@(1) + (2001..2015)) -join ','
    $ids15   = (@(1) + (2001..2014)) -join ','
    $ids40   = (@(1) + (2001..2039)) -join ','
    $idCases = @(
        @('-Id 1 つ', @('-Id', '9002'), '0/9002'),
        @('-Id 2 つ', @('-Id', '9002,1'), '0/9002,1,1,1'),
        @('-Id 前後の空白', @('-Id', ' 9002 , 9003 '), '0/9003,9002'),
        @('-Id 0', @('-Id', '0'), '0/0'),
        @('-Id 同じ ID を 2 回', @('-Id', '9002,9002'), '0/9002'),
        @('-ExcludeId', @('-ExcludeId', '1'), '0/9004,9003,0,9002,9001'),
        @('-ExcludeId 2 つ', @('-ExcludeId', '1,9004'), '0/9003,0,9002,9001'),
        @('-ExcludeId 0', @('-ExcludeId', '0'), '0/9004,9003,9002,1,1,1,9001'),
        @('-Id と -ExcludeId', @('-Id', '1,9002', '-ExcludeId', '1'), '0/9002'),
        @('-Id と -ExcludeId が同じ', @('-Id', '1', '-ExcludeId', '1'), '2/'),
        @('ID 16 個（取り出した後に絞る）', @('-Level', $allLevels, '-Id', $manyIds), '0/1,1,1'),
        @('ID 15 個とすべてのレベル（式に書く上限）', @('-Level', $allLevels, '-Id', $ids15), '0/1,1,1'),
        @('ID 40 個', @('-Level', $allLevels, '-Id', $ids40), '0/1,1,1'),
        @('ID 16 個と -ExcludeId', @('-Id', $manyIds, '-ExcludeId', '1'), '2/'),
        @('ID 65535', @('-Id', '65535'), '2/')
    )
    foreach ($case in $idCases) {
        $run = Invoke-Tool ('id' + [Array]::IndexOf($idCases, $case)) ($mainArgs + $case[1])
        Assert-Equal $case[0] (Get-CodeAndIds $run) $case[2]
    }

    Write-Section 'ソース'
    $sourceCases = @(
        @('イベント ビューアーの名前', 'Kernel-Power', '0/9001'),
        @('元の名前', 'Microsoft-Windows-Kernel-Power', '0/9001'),
        @('大文字小文字を区別しない', 'kernel-power', '0/9001'),
        @('元の名前の大文字小文字', 'MICROSOFT-WINDOWS-kernel-POWER', '0/9001'),
        @('この PC にないソース', 'WB-OTHER', '0/9004,9003,9002'),
        @('カンマで複数', 'Kernel-Power, wb-other', '0/9004,9003,9002,9001'),
        @('前方一致では探さない', 'Kernel', '2/'),
        @('省く部分だけでは探さない', 'Microsoft-Windows-', '2/'),
        @('ないソース', 'NoSuchSource', '2/')
    )
    foreach ($case in $sourceCases) {
        $run = Invoke-Tool ('source' + [Array]::IndexOf($sourceCases, $case)) ($mainArgs + @('-Source', $case[1]))
        Assert-Equal ('-Source {0}: {1}' -f $case[1], $case[0]) (Get-CodeAndIds $run) $case[2]
    }
    $run = Invoke-Tool 'cond' ($mainArgs + @('-Id', '9002,1', '-ExcludeId', '1,5', '-Source', 'wb-other,WSH', '-Pattern', ' Timeout | 警告 '))
    Assert-Equal '条件の行（すべて）' $run.Log.Contains('条件: ID 9002, 1・除く ID 1, 5・ソース wb-other, WSH・メッセージ Timeout | 警告') 'True'
    $run = Invoke-Tool 'cond2' ($mainArgs + @('-Source', 'WSH'))
    Assert-Equal '条件の行（ソースだけ）' $run.Log.Contains('条件: ソース WSH期間') 'False'
    Assert-Equal '条件の行（ソースだけ）の中身' $run.Log.Contains('条件: ソース WSH') 'True'

    if ($wshReady) {
        Write-Section 'メッセージ'
        $run = Invoke-Tool 'message' $mainArgs
        $x = Read-Result $run.Dir
        Assert-Equal '全文（改行も含む）' ((@($x.Events) | Where-Object { $_.'日時' -like '*10:03:00' }).'メッセージ' -replace "`r`n", '|') 'エラー Timeout 3|2 行目|3 行目'
        Assert-Equal 'ソースの情報がないイベントの件数' $run.Log.Contains('メッセージを取得できなかったイベント 4 件') 'True'
        Assert-Equal '画面は 1 行目だけ' ('{0}/{1}' -f $run.Log.Contains('エラー Timeout 3'), $run.Log.Contains('2 行目')) 'True/False'
        # レビューで見つかった問題: ="..." の中の " を重ねていなかった
        Assert-Equal 'Excel 対策（中の " を重ねる）' ((@($x.Events) | Where-Object { $_.'イベントID' -eq '0' }).'メッセージ') '="=1+1 ""quoted"""'
        Assert-Equal 'summary.csv も Excel 対策' ((@($x.Groups) | Where-Object { $_.'イベントID' -eq '0' }).'メッセージ（1 行目）') '="=1+1 ""quoted"""'

        $patternCases = @(
            @('大文字小文字を区別しない', 'TIMEOUT', '0/1,1,1'),
            @('| でどれか', 'nomatch|timeout', '0/1,1,1'),
            @('どれも一致しない', 'nomatch|nothing', '2/'),
            @('前後の空白を外す', ' timeout 3 ', '0/1'),
            @('2 行目以降も探す', '3 行目', '0/1'),
            @('正規表現として読まない', 'Time.ut', '2/'),
            @('* を文字として読む', 'Time*', '2/'),
            @('記号', '=1+1', '0/0'),
            @('| だけなら絞らない', '|', '0/9004,9003,0,9002,1,1,1,9001')
        )
        foreach ($case in $patternCases) {
            $run = Invoke-Tool ('pattern' + [Array]::IndexOf($patternCases, $case)) ($mainArgs + @('-Pattern', $case[1]))
            Assert-Equal ('-Pattern {0}: {1}' -f $case[1], $case[0]) (Get-CodeAndIds $run) $case[2]
        }
        $run = Invoke-Tool 'patternjp' ($mainArgs + @('-Pattern', '警告|情報', '-Level', $allLevels))
        Assert-Equal '-Pattern 日本語' (Get-CodeAndIds $run) '0/4,2,2'
        Assert-Equal '-Pattern で除いたメッセージなしは数えない' $run.Log.Contains('メッセージを取得できなかった') 'False'

        $base = Get-TestTime '10:00'
        $longMessage = 'A' * 100
        $excelEvents = @(
            @('数字', '12345', '="12345"'), @('小数', '.5', '=".5"'), @('全角の数字', '１２３', '="１２３"'), @('全角の小数', '１．５', '="１．５"'), @('日付', '2026-01-01', '="2026-01-01"'),
            @('負の数', '-1', '="-1"'), @('プラス', '+1 abc', '="+1 abc"'), @('@', '@sum', '="@sum"'), @('パーセント', '100%', '="100%"'), @('指数', '1e5', '="1e5"'),
            @('ふつうの文', 'normal text', 'normal text'), @('数字で始まる文', '123 abc', '123 abc'), @('カンマと引用符', 'a,b "c"', 'a,b "c"'),
            @('255 文字の = は式にする', ('=' + 'x' * 254), ('="=' + ('x' * 254) + '"')), @('256 文字の = は先頭に空白', ('=' + 'x' * 255), (' =' + ('x' * 255)))
        )
        $excelList = New-Object 'System.Collections.Generic.List[object]'
        foreach ($i in 0..($excelEvents.Count - 1)) { $excelList.Add((New-TestEvent $base.AddMinutes($i) 2 1 'WSH' $excelEvents[$i][1] 'Application')) }
        $excel = Join-Path $workRoot 'excel\excel.evtx'
        New-EvtxFile $excel $excelList.ToArray()
        $run = Invoke-Tool 'excel' (@('-Path', $excel) + $range)
        $rows = @((Read-Result $run.Dir).Events)
        foreach ($i in 0..($excelEvents.Count - 1)) {
            Assert-Equal ('Excel 対策: {0}' -f $excelEvents[$i][0]) $rows[$excelEvents.Count - 1 - $i].'メッセージ' $excelEvents[$i][2]
        }

        Write-Section '見えない文字'
        # 実機確認で見つかった問題: Windows はメッセージの日時に見えない文字（U+200E）を入れるので、-Pattern で日時を探せなかった
        # （-ceq はこの文字を無視することがあるので、Contains と -Pattern の結果で確かめる）
        $lrm = [string][char]0x200E
        $rlm = [string][char]0x200F
        $markFile = Join-Path $workRoot 'marks\marks.evtx'
        New-EvtxFile $markFile @(New-TestEvent (Get-TestTime '10:00') 2 1 'WSH' ("開始 ${lrm}2026${lrm}-${lrm}11${rlm}-${lrm}01 終了") 'Application')
        $run = Invoke-Tool 'marks' (@('-Path', $markFile, '-Pattern', '2026-11-01') + $range)
        Assert-Equal '見えない文字を除いて探す' (Get-CodeAndIds $run) '0/1'
        $saved = (@((Read-Result $run.Dir).Events) | ForEach-Object { $_.'メッセージ' }) -join ''
        Assert-Equal '見えない文字を除いて保存する' ('{0}/{1}/{2}' -f $saved.Contains($lrm), $saved.Contains($rlm), $saved.Contains('開始 2026-11-01 終了')) 'False/False/True'

        Write-Section 'メッセージの長さ'
        $lengthFile = Join-Path $workRoot 'length\length.evtx'
        New-EvtxFile $lengthFile @(
            (New-TestEvent (Get-TestTime '10:00') 2 1 'WSH' $longMessage 'Application'),
            (New-TestEvent (Get-TestTime '10:01') 2 2 'WSH' ('B' * 80) 'Application'),
            (New-TestEvent (Get-TestTime '10:02') 2 4 'WSH' "`r`n`r`n   本文の 1 行目   `r`r`n2 行目" 'Application'),
            (New-TestEvent (Get-TestTime '10:03') 2 0 'WSH' ('C' * 5000) 'Application')
        )
        $run = Invoke-Tool 'length' (@('-Path', $lengthFile) + $range)
        $x = Read-Result $run.Dir
        Assert-Equal '画面は 80 文字まで' ('{0}/{1}' -f $run.Log.Contains(('A' * 80) + '…'), $run.Log.Contains('A' * 81)) 'True/False'
        Assert-Equal '80 文字ちょうどは … を付けない' ('{0}/{1}' -f $run.Log.Contains('B' * 80), $run.Log.Contains(('B' * 80) + '…')) 'True/False'
        Assert-Equal 'events.csv は全文' ((@($x.Events) | Where-Object { $_.'イベントID' -eq '1' }).'メッセージ') $longMessage
        Assert-Equal '空の行を飛ばして 1 行目を取る' ((@($x.Groups) | Where-Object { $_.'イベントID' -eq '4' }).'メッセージ（1 行目）') '本文の 1 行目'
        Assert-Equal '長いメッセージも全文' (@(@($x.Events) | Where-Object { $_.'イベントID' -eq '0' })[0].'メッセージ'.Length) 5000
        $patched = New-PatchedTool 'cell' '$CellLimit     = 32000' '$CellLimit     = 50'
        $run = Invoke-Tool 'cell' (@('-Path', $lengthFile) + $range) $patched
        $x = Read-Result $run.Dir
        Assert-Equal 'セルの上限で切る' ((@($x.Events) | Where-Object { $_.'イベントID' -eq '1' }).'メッセージ') ('A' * 50)
        Assert-Equal '切った後の値も数える' ((@($x.Groups) | Where-Object { $_.'イベントID' -eq '1' }).'メッセージ（1 行目）') $longMessage
    } else {
        Write-Skip 'メッセージ' 'ソース WSH の情報がないため'
    }

    Write-Section 'ログの中の値の Excel 対策'
    # ソースの名前は、管理者でなくても好きな名前で書き込める（ReportEvent）。数式を仕込まれても Excel で動かないようにする
    $injection = Join-Path $workRoot 'inject\inject.evtx'
    $formula   = "=cmd|' /C calc'!A0"
    New-EvtxFile $injection @(
        (New-TestEvent (Get-TestTime '10:00') 2 7001 $formula '' 'Application' '' '@host'),
        (New-TestEvent (Get-TestTime '10:01') 2 7002 '12345' '' '+log' '' '001')
    )
    $run = Invoke-Tool 'inject' (@('-Path', $injection) + $range)
    $x = Read-Result $run.Dir
    $row = @($x.Events | Where-Object { $_.'イベントID' -eq '7001' })[0]
    Assert-Equal 'events.csv のソース（数式）' $row.'ソース' ('="{0}"' -f $formula)
    Assert-Equal 'events.csv のコンピューター' $row.'コンピューター' '="@host"'
    $row = @($x.Events | Where-Object { $_.'イベントID' -eq '7002' })[0]
    Assert-Equal 'events.csv のソース（数字）' $row.'ソース' '="12345"'
    Assert-Equal 'events.csv のログ' $row.'ログ' '="+log"'
    Assert-Equal 'events.csv のコンピューター（数字）' $row.'コンピューター' '="001"'
    $row = @($x.Groups | Where-Object { $_.'イベントID' -eq '7001' })[0]
    Assert-Equal 'summary.csv のソース' $row.'ソース' ('="{0}"' -f $formula)
    $row = @($x.Groups | Where-Object { $_.'イベントID' -eq '7002' })[0]
    Assert-Equal 'summary.csv のログ' $row.'ログ' '="+log"'
    Assert-Equal '画面はそのまま' $run.Log.Contains("  $formula  ") 'True'

    Write-Section '表示の上限と 3 桁区切り'
    $base = Get-TestTime '10:00'
    $manyList = New-Object 'System.Collections.Generic.List[object]'
    foreach ($i in 1..35) { $manyList.Add((New-TestEvent $base.AddSeconds($i) 2 (3000 + $i) 'wb-many')) }
    foreach ($i in 1..3) { $manyList.Add((New-TestEvent $base.AddSeconds(100 + $i) 1 100 'wb-many')) }
    $many = Join-Path $workRoot 'many\many.evtx'
    New-EvtxFile $many $manyList.ToArray()
    $run = Invoke-Tool 'many' (@('-Path', $many) + $range)
    $x = Read-Result $run.Dir
    Assert-Equal '集計は 30 件まで' ([regex]::Matches($run.Log, 'wb-many\s+\d+\s+\d{4}-\d{2}-\d{2} \d{2}:\d{2}  ').Count) 30
    Assert-Equal '集計の続きの案内' $run.Log.Contains('（ほか 6 件は summary.csv を参照してください）') 'True'
    Assert-Equal '新しい順は 20 件まで' ([regex]::Matches($run.Log, '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}  ').Count) 20
    Assert-Equal '新しい順の続きの案内' $run.Log.Contains('（ほか 18 件は events.csv を参照してください）') 'True'
    Assert-Equal 'CSV はすべて' ('{0}/{1}' -f @($x.Events).Count, @($x.Groups).Count) '38/36'
    Assert-Equal 'summary.txt の集計はすべて' (@($x.Lines | Where-Object { $_ -match 'wb-many\s+\d+\s+\d{4}-' }).Count) 36
    Assert-Equal '件数の多い順で、件数は右にそろえる' $run.Log.Contains('■ 集計（件数の多い順・上位 30） 件数  レベル') 'True'
    $run = Invoke-Plain (@('-Path', $many, '-NoSave') + $range)
    Assert-Equal '-NoSave: 集計はすべて' ([regex]::Matches($run.Log, 'wb-many\s+\d+\s+\d{4}-\d{2}-\d{2} \d{2}:\d{2}  ').Count) 36
    Assert-Equal '-NoSave: 新しい順はすべて' ([regex]::Matches($run.Log, '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}  ').Count) 38
    Assert-Equal '-NoSave: 見出しに上限を書かない' ('{0}/{1}/{2}' -f $run.Log.Contains('■ 集計（件数の多い順）'), $run.Log.Contains('■ 新しい順 '), $run.Log.Contains('上位')) 'True/True/False'
    Assert-Equal '-NoSave: 続きの案内を出さない' $run.Log.Contains('（ほか') 'False'
    Assert-Equal '-NoSave: 保存先の行を出さない' $run.Log.Contains('保存先') 'False'
    $bulkList = New-Object 'System.Collections.Generic.List[object]'
    foreach ($i in 1..1234) { $bulkList.Add((New-TestEvent $base.AddSeconds($i) 2 100 'wb-bulk')) }
    $bulk = Join-Path $workRoot 'bulk\bulk.evtx'
    New-EvtxFile $bulk $bulkList.ToArray()
    Assert-Equal '複数のチャンクの .evtx ができた' ((Get-Item -LiteralPath $bulk).Length -gt (4096 + 65536)) 'True'
    $run = Invoke-Tool 'bulk' (@('-Path', $bulk) + $range)
    $x = Read-Result $run.Dir
    Assert-Equal '件数の行は 3 桁区切り' $run.Log.Contains('イベント 1,234 件（エラー 1,234）  種類 1') 'True'
    Assert-Equal '集計の件数は 3 桁区切り' ('{0}/{1}' -f @($x.Groups)[0].'件数', @($x.Events).Count) '1,234/1234'
    Assert-Equal '画面の集計の件数は 3 桁区切りで列をそろえる' $run.Log.Contains(' 1,234  エラー') 'True'
    Assert-Equal '新しい順の続きの案内は 3 桁区切り' $run.Log.Contains('（ほか 1,214 件は events.csv を参照してください）') 'True'

    Write-Section 'ファイルとフォルダ'
    # レビューで見つかった問題: Get-WinEvent は Path をワイルドカードとして読むため、[a].evtx が隣の a.evtx に一致していた
    $nameDir = Join-Path $workRoot 'names'
    New-EvtxFile "$nameDir\[a].evtx" @(New-TestEvent (Get-TestTime '10:00') 2 11)
    New-EvtxFile "$nameDir\a.evtx" @(New-TestEvent (Get-TestTime '10:00') 2 12)
    $run = Invoke-Tool 'bracket' (@('-Path', "$nameDir\[a].evtx") + $range)
    Assert-Equal '名前に [ ] を含むファイル' (Get-CodeAndIds $run) '0/11'
    Assert-Equal '[ ] のファイルの対象の行' $run.Log.Contains(('対象: {0}\[a].evtx' -f $nameDir)) 'True'
    $nameCases = @(
        # Windows でのテストで見つかった問題: -FilterHashtable だと、名前に & を含むファイルは XML のエラーで読めなかった
        @('&', 'a&b.evtx'), @("'", "it's.evtx"), @('`', 'a`b.evtx'), @('スペース', 'a b.evtx'), @('日本語', '日本語のログ.evtx'), @('%', '100%.evtx'), @('絵文字', ('log' + [char]::ConvertFromUtf32(0x1F600) + '.evtx')),
        @(';', 'a;b.evtx'), @('大文字の拡張子', 'UPPER.EVTX')
    )
    foreach ($case in $nameCases) {
        $dir = Join-Path $workRoot ('name' + [Array]::IndexOf($nameCases, $case))
        New-EvtxFile "$dir\$($case[1])" @(New-TestEvent (Get-TestTime '10:00') 2 21)
        $run = Invoke-Tool ('namefile' + [Array]::IndexOf($nameCases, $case)) (@('-Path', "$dir\$($case[1])") + $range)
        Assert-Equal ('名前に {0}: ファイルを指定' -f $case[0]) (Get-CodeAndIds $run) '0/21'
        $run = Invoke-Tool ('namedir' + [Array]::IndexOf($nameCases, $case)) (@('-Path', $dir) + $range)
        Assert-Equal ('名前に {0}: フォルダを指定' -f $case[0]) (Get-CodeAndIds $run) '0/21'
    }
    $spaceDir = Join-Path $workRoot '日本語 と スペース [x]'
    New-EvtxFile "$spaceDir\in.evtx" @(New-TestEvent (Get-TestTime '10:00') 2 22)
    $run = Invoke-Tool 'spacedir' (@('-Path', $spaceDir) + $range)
    Assert-Equal 'フォルダの名前に日本語・スペース・[ ]' (Get-CodeAndIds $run) '0/22'

    $treeDir = Join-Path $workRoot 'tree'
    New-EvtxFile "$treeDir\b.evtx" @(New-TestEvent (Get-TestTime '10:00') 2 31)
    New-EvtxFile "$treeDir\a.evtx" @(New-TestEvent (Get-TestTime '10:01') 2 32)
    New-EvtxFile "$treeDir\sub\deep\c.evtx" @(New-TestEvent (Get-TestTime '10:02') 2 33)
    [IO.File]::WriteAllText("$treeDir\note.txt", 'not a log')
    [IO.File]::WriteAllText("$treeDir\fake.evtx.bak", 'not a log')
    $junction = "$treeDir\loop"
    $null = New-Item -ItemType Junction -Path $junction -Target $treeDir
    $run = Invoke-Tool 'tree' (@('-Path', $treeDir) + $range)
    $x = Read-Result $run.Dir
    Assert-Equal 'フォルダの下の .evtx をすべて読む（サブフォルダも）' (Get-CodeAndIds $run) '0/33,32,31'
    Assert-Equal '対象の行（ファイルは名前の順）' $run.Log.Contains(('対象: {0}\a.evtx ほか 2 ファイル' -f $treeDir)) 'True'
    Assert-Equal 'ループするジャンクションをたどらない' ((@($x.Skipped) | ForEach-Object { '{0}/{1}/{2}' -f $_.'理由', $_.'対象', $_.'詳細' }) -join ',') "リンク/$junction/ジャンクション"
    Assert-Equal 'スキップの行' $run.Log.Contains('スキップ 1 件（リンク）') 'True'
    $run = Invoke-Tool 'treeslash' (@('-Path', "$treeDir\") + $range)
    Assert-Equal 'フォルダの最後に \' (Get-CodeAndIds $run) '0/33,32,31'
    $run = Invoke-Tool 'multi' (@('-Path', ('{0}\a.evtx,{1}\[a].evtx' -f $treeDir, $nameDir)) + $range)
    Assert-Equal 'カンマで複数のファイル' ('{0}/{1}' -f (Get-CodeAndIds $run), $run.Log.Contains('ほか 1 ファイル')) '0/32,11/True'
    $run = Invoke-Tool 'multiquote' (@('-Path', ("'{0}\a.evtx', '{1}\[a].evtx'" -f $treeDir, $nameDir)) + $range)
    Assert-Equal 'カンマの後の空白と '' を外す' (Get-CodeAndIds $run) '0/32,11'
    $run = Invoke-Tool 'missing' (@('-Path', ('{0}\a.evtx,{0}\no.evtx' -f $treeDir)) + $range)
    Assert-Equal 'ないファイルは飛ばしてほかを読む' ('{0}/{1}' -f (Get-CodeAndIds $run), ((@((Read-Result $run.Dir).Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', $_.'詳細' }) -join ',')) '0/32/ファイルを読めない/見つかりません'
    $emptyDir = Join-Path $workRoot 'emptydir'
    [void][IO.Directory]::CreateDirectory($emptyDir)
    $run = Invoke-Tool 'emptydir' (@('-Path', $emptyDir) + $range)
    Assert-Equal '.evtx のないフォルダ' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('読める .evtx ファイルがありません'), [bool]$run.Dir) '1/True/False'

    Write-Section '読めないファイル'
    # Windows でのテストで見つかった問題: -FilterHashtable だと、壊れた .evtx が「該当なし」になっていた
    $badDir = Join-Path $workRoot 'bad'
    [void][IO.Directory]::CreateDirectory($badDir)
    [IO.File]::WriteAllBytes("$badDir\broken.evtx", [byte[]](1..200 | ForEach-Object { 7 }))
    [IO.File]::WriteAllBytes("$badDir\zero.evtx", [byte[]]@())
    [IO.File]::WriteAllText("$badDir\text.log", 'not a log')
    foreach ($case in @(@('壊れた .evtx', 'broken.evtx'), @('0 バイトの .evtx', 'zero.evtx'), @('.evtx でないファイル', 'text.log'))) {
        $run = Invoke-Tool ('bad' + $case[1]) (@('-Path', "$badDir\$($case[1])") + $range)
        Assert-Equal ('{0}: 終了コード 1 で止まり、結果を残さない' -f $case[0]) ('{0}/{1}' -f $run.Code, [bool]$run.Dir) '1/False'
        Assert-Equal ('{0}: 理由を表示する' -f $case[0]) ('{0}/{1}' -f $run.Log.Contains('読めるログがありませんでした'), $run.Log.Contains(('{0}\{1}: ファイルを読めない' -f $badDir, $case[1]))) 'True/True'
    }
    $run = Invoke-Tool 'badmix' (@('-Path', ('{0}\broken.evtx,{1}\a.evtx' -f $badDir, $treeDir)) + $range)
    $x = Read-Result $run.Dir
    Assert-Equal '壊れたファイルは飛ばしてほかを読む' (Get-CodeAndIds $run) '0/32'
    Assert-Equal 'skipped.csv の理由' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', $_.'対象' }) -join ',') "ファイルを読めない/$badDir\broken.evtx"
    Assert-Equal 'スキップの行（ファイルを読めない）' $run.Log.Contains('スキップ 1 件（ファイルを読めない）') 'True'
    $run = Invoke-Tool 'badnomatch' @('-Path', ('{0}\broken.evtx,{1}\a.evtx' -f $badDir, $treeDir), '-Since', '2000-01-01', '-Until', '2000-01-02')
    Assert-Equal '読めたファイルに該当がなければ、該当なし' ('{0}/{1}' -f $run.Code, $run.Log.Contains('スキップ 1 件（ファイルを読めない）')) '2/True'
    $run = Invoke-Tool 'baddir' (@('-Path', $badDir) + $range)
    Assert-Equal '壊れたファイルしかないフォルダ' ('{0}/{1}' -f $run.Code, [bool]$run.Dir) '1/False'
    Assert-Equal 'すべて読めないときは理由を並べる' ($run.Log.Contains('broken.evtx: ファイルを読めない') -and $run.Log.Contains('zero.evtx: ファイルを読めない')) 'True'
    # Windows でのテストで見つかった問題: 新しい順に読むと、最後のほうが壊れた .evtx は 1 件も読めなかった（壊れた場所の手前までは読める）
    $bytes = [IO.File]::ReadAllBytes($bulk)
    $lastChunk = $bytes.Length - 65536
    foreach ($i in ($lastChunk + 10000)..($lastChunk + 10300)) { $bytes[$i] = 0xEE }
    $half = Join-Path $badDir 'half\half.evtx'
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $half))
    [IO.File]::WriteAllBytes($half, $bytes)
    $run = Invoke-Tool 'half' (@('-Path', $half) + $range)
    $x = Read-Result $run.Dir
    $halfRead = @($x.Events).Count
    Assert-Equal '途中で壊れた .evtx: 読めた分を出す' ('{0}/{1}' -f $run.Code, ($halfRead -gt 0 -and $halfRead -lt 1234)) '0/True'
    Assert-Equal '途中で壊れた .evtx: skipped.csv に残す' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', ($_.'詳細' -replace ':.*$', '') }) -join ',') ('ファイルを読めない/途中まで読んだ（イベント {0} 件は結果に含む）' -f $halfRead.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))

    Write-Section '権限なし'
    $denyFile = Join-Path $workRoot 'deny\denied.evtx'
    New-EvtxFile $denyFile @(New-TestEvent (Get-TestTime '10:00') 2 41)
    New-EvtxFile "$workRoot\deny\open.evtx" @(New-TestEvent (Get-TestTime '10:01') 2 42)
    if (Add-Deny $denyFile '(R)') {
        $run = Invoke-Tool 'denyfile' (@('-Path', "$workRoot\deny") + $range)
        $x = Read-Result $run.Dir
        Assert-Equal '読めないファイルは飛ばしてほかを読む' (Get-CodeAndIds $run) '0/42'
        Assert-Equal '読めないファイルは権限なし' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', $_.'対象' }) -join ',') "権限なし/$denyFile"
    } else {
        Write-Skip '読めないファイル' 'アクセス拒否を設定できなかったため'
    }
    $denyDir = Join-Path $workRoot 'denydir'
    New-EvtxFile "$denyDir\closed\in.evtx" @(New-TestEvent (Get-TestTime '10:00') 2 43)
    New-EvtxFile "$denyDir\open.evtx" @(New-TestEvent (Get-TestTime '10:01') 2 44)
    if (Add-Deny "$denyDir\closed" '(RD)') {
        $run = Invoke-Tool 'denydir' (@('-Path', $denyDir) + $range)
        $x = Read-Result $run.Dir
        Assert-Equal '読めないフォルダは飛ばしてほかを読む' (Get-CodeAndIds $run) '0/44'
        Assert-Equal '読めないフォルダは権限なし' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', $_.'対象' }) -join ',') "権限なし/$denyDir\closed"
        Assert-Equal 'スキップの行（権限なし）' $run.Log.Contains('スキップ 1 件（権限なし）') 'True'
    } else {
        Write-Skip '読めないフォルダ' 'アクセス拒否を設定できなかったため'
    }

    Write-Section '260 文字を超えるパス'
    $longDir  = "$workRoot\" + ('L' * 120) + '\' + ('M' * 120)
    New-EvtxFile "$longDir\deep\long&[x].evtx" @(New-TestEvent (Get-TestTime '10:00') 2 51)
    $run = Invoke-Tool 'longdir' (@('-Path', $longDir) + $range)
    Assert-Equal '-Path にフォルダ' (Get-CodeAndIds $run) '0/51'
    $run = Invoke-Tool 'longfile' (@('-Path', "$longDir\deep\long&[x].evtx") + $range)
    Assert-Equal '-Path にファイル' ('{0}/{1}' -f (Get-CodeAndIds $run), $run.Log.Contains(('対象: {0}\deep\long&[x].evtx' -f $longDir))) '0/51/True'
    $edgeRoot = Join-Path $workRoot 'edge'
    foreach ($length in 249, 250, 259, 260, 261) {
        $file = "$edgeRoot\" + ('z' * ($length - $edgeRoot.Length - 6)) + '.evtx'
        New-EvtxFile $file @(New-TestEvent (Get-TestTime '10:00') 2 $length)
        $run = Invoke-Tool "edge$length" (@('-Path', $file) + $range)
        Assert-Equal ('パスの長さ {0}（実際 {1}）' -f $length, $file.Length) (Get-CodeAndIds $run) "0/$length"
    }
    $longOut = "$workRoot\" + ('O' * 120) + '\' + ('P' * 120)
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $main @range -OutputFolder $longOut 2>&1
    $longRuns = @([IO.Directory]::GetDirectories('\\?\' + $longOut))
    Assert-Equal '-OutputFolder が長くても保存する' ('{0}/{1}' -f $LASTEXITCODE, $longRuns.Count) '0/1'
    if ($longRuns.Count -eq 1) { Assert-Equal '長い保存先のファイル' (@([IO.Directory]::GetFiles($longRuns[0])).Count) 4 }

    Write-Section 'ドライブのルート'
    # Windows でのテストで見つかった問題: 最後の \ を外すと \\?\Q: になり、読めなかった
    $letter = Get-FreeLetter
    if ($letter) {
        $rootDir = Join-Path $workRoot 'rootdrive'
        New-EvtxFile "$rootDir\sub\in.evtx" @(New-TestEvent (Get-TestTime '10:00') 2 61)
        $null = & subst.exe "${letter}:" $rootDir 2>&1
        if ($LASTEXITCODE -eq 0) {
            $substDrive = "${letter}:"
            $run = Invoke-Tool 'driveroot' (@('-Path', "$substDrive\") + $range)
            Assert-Equal '-Path にドライブのルート' (Get-CodeAndIds $run) '0/61'
            $run = Invoke-Tool 'driveonly' (@('-Path', $substDrive) + $range)
            Assert-Equal '-Path にドライブ文字だけ' (Get-CodeAndIds $run) '0/61'
        } else {
            Write-Skip 'ドライブのルート' 'subst を使えなかったため'
        }
    } else {
        Write-Skip 'ドライブのルート' '空いているドライブ文字がないため'
    }

    Write-Section 'UNC のパス'
    # 管理共有を読めない環境では飛ばす
    $uncRoot = '\\localhost\{0}$\{1}' -f $workRoot.Substring(0, 1), $workRoot.Substring(3)
    if ([IO.File]::Exists("$uncRoot\evtx\main.evtx")) {
        $run = Invoke-Tool 'uncfile' (@('-Path', "$uncRoot\evtx\main.evtx") + $range)
        Assert-Equal 'UNC のファイル' ('{0}/{1}' -f (Get-CodeAndIds $run), $run.Log.Contains(('対象: {0}\evtx\main.evtx' -f $uncRoot))) '0/9004,9003,0,9002,1,1,1,9001/True'
        $run = Invoke-Tool 'uncdir' (@('-Path', "$uncRoot\tree") + $range)
        Assert-Equal 'UNC のフォルダ' (Get-CodeAndIds $run) '0/33,32,31'
        $run = Invoke-Tool 'unclong' (@('-Path', ('{0}\{1}\{2}\deep' -f $uncRoot, ('L' * 120), ('M' * 120))) + $range)
        Assert-Equal 'UNC の 260 文字を超えるパス' (Get-CodeAndIds $run) '0/51'
    } else {
        Write-Skip 'UNC のパス' '管理共有を読めないため'
    }

    Write-Section '相対パスと [ ] を含むフォルダからの起動'
    $relDir = Join-Path $workRoot 'relative'
    New-EvtxFile "$relDir\here.evtx" @(New-TestEvent (Get-TestTime '10:00') 2 71)
    $run = Invoke-From $relDir (@('-Path', 'here.evtx', '-OutputFolder', 'out') + $range)
    Assert-Equal '-Path と -OutputFolder に相対パス' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('イベント 1 件'), (Get-ResultFileNames (Get-RunDir "$relDir\out"))) '0/True/events.csv,skipped.csv,summary.csv,summary.txt'
    $run = Invoke-From $relDir (@('-Path', '.', '-NoSave') + $range)
    Assert-Equal '-Path に .' ('{0}/{1}' -f $run.Code, $run.Log.Contains('イベント 1 件')) '0/True'
    $start = Join-Path $workRoot 'start'
    $startCases = @(
        @('隣に一致するフォルダ', '[a]', @('a')),
        @('一致するフォルダなし', 'b[1]', @()),
        @('パターンとして正しくない', 'c[d', @()),
        @('2 つ以上一致', '[ef]', @('e', 'f'))
    )
    foreach ($case in $startCases) {
        New-EvtxFile "$start\$($case[1])\here.evtx" @(New-TestEvent (Get-TestTime '10:00') 2 81)
        # 間違ったフォルダを読んだら分かるように、隣のフォルダには別のイベントを書く
        foreach ($decoy in $case[2]) { New-EvtxFile "$start\$decoy\here.evtx" @(New-TestEvent (Get-TestTime '10:00') 2 82) }
        $dir = "$start\$($case[1])"
        $run = Invoke-From $dir (@('-Path', 'here.evtx', '-OutputFolder', 'out') + $range)
        Assert-Equal ('[ ] を含むフォルダから: {0}' -f $case[0]) ('{0}/{1}' -f $run.Code, (Get-Ids (Read-Result (Get-RunDir "$dir\out")).Events)) '0/81'
    }

    Write-Section 'この PC のイベントログ（読むだけ）'
    $liveUntil = (Get-Date).AddMinutes(-2)
    $liveSince = $liveUntil.AddDays(-7)
    $liveRange = @('-Since', $liveSince.ToString('yyyy-MM-dd HH:mm:ss'), '-Until', $liveUntil.ToString('yyyy-MM-dd HH:mm:ss'))
    $queryTime = { param($t) $t.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z') }
    $liveQuery = "*[System[(Level=1 or Level=2) and TimeCreated[@SystemTime>='{0}' and @SystemTime<='{1}']]]" -f (& $queryTime $liveSince), (& $queryTime $liveUntil)
    $liveCount = 0
    $liveLogs  = @('System', 'Application')
    foreach ($logName in $liveLogs) {
        $xml = (& wevtutil.exe qe $logName "/q:$liveQuery" /f:xml) -join ''
        $liveCount += [regex]::Matches($xml, '<Event ').Count
    }
    $run = Invoke-Tool 'live' $liveRange
    $x = Read-Result $run.Dir
    Assert-Equal '既定の System と Application: wevtutil と同じ件数' @($x.Events).Count $liveCount
    Assert-Equal '既定の対象の行' $run.Log.Contains(('対象: {0}（管理者: {1}）   ログ: System, Application' -f $env:COMPUTERNAME, $(if ($isAdmin) { 'はい' } else { 'いいえ' }))) 'True'
    Assert-Equal '既定の終了コード' $run.Code $(if ($liveCount -gt 0) { 0 } else { 2 })
    if ($liveCount -gt 0) {
        Assert-Equal '取り出したのは System と Application だけ' (@(@($x.Events) | Where-Object { $liveLogs -notcontains $_.'ログ' }).Count) 0
        Assert-Equal '取り出したのは重大とエラーだけ' (@(@($x.Events) | Where-Object { @('重大', 'エラー') -notcontains $_.'レベル' }).Count) 0
    } else {
        Write-Skip '本物のイベントの中身' '直近 7 日に重大とエラーがないため'
    }
    $run = Invoke-Tool 'livelist' (@('-LogName', 'NoSuchLog,Application') + $liveRange)
    $x = Read-Result $run.Dir
    Assert-Equal 'ないログは飛ばす' ((@($x.Skipped) | ForEach-Object { '{0}/{1}/{2}' -f $_.'理由', $_.'対象', $_.'詳細' }) -join ',') 'ログがない/NoSuchLog/この PC にはないログです'
    Assert-Equal 'スキップの行（ログがない）' $run.Log.Contains('スキップ 1 件（ログがない）') 'True'
    Assert-Equal 'ログの行に書いたログを並べる' $run.Log.Contains('ログ: NoSuchLog, Application') 'True'
    $run = Invoke-Tool 'liveslash' (@('-LogName', 'Microsoft-Windows-TaskScheduler/Operational,Windows PowerShell', '-Level', 'Warning') + $liveRange)
    Assert-Equal '/ やスペースを含むログの名前' (@((Read-Result $run.Dir).Skipped).Count) 0
    Assert-Equal '/ やスペースを含むログの名前: 終了コード' ($run.Code -eq 0 -or $run.Code -eq 2) 'True'
    # 分析とデバッグのログは -Oldest でないと読めない（Get-WinEvent -ListLog は、Select-Object -First で受けるとスクリプトごと黙って終わる）
    $debugLog = @(Get-WinEvent -ListLog '*' -Force -ErrorAction SilentlyContinue | Where-Object { @('Analytical', 'Debug') -contains [string]$_.LogType -and -not $_.IsEnabled })
    if ($debugLog.Count -gt 0) {
        $run = Invoke-Tool 'livedebug' (@('-LogName', $debugLog[0].LogName) + $liveRange)
        Assert-Equal ('分析とデバッグのログ（{0}）' -f $debugLog[0].LogName) ('{0}/{1}' -f $run.Code, @((Read-Result $run.Dir).Skipped).Count) '2/0'
        # -Oldest を外すとエラーになることを使って、「その他のエラー」を確かめる
        $patched = New-PatchedTool 'nooldest' '-FilterXPath $query -Oldest' '-FilterXPath $query'
        $run = Invoke-Tool 'othererror' (@('-LogName', ('{0},System' -f $debugLog[0].LogName)) + $liveRange) $patched
        Assert-Equal 'ログを読むときのほかのエラーは、その他のエラー' ((@((Read-Result $run.Dir).Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', $_.'対象' }) -join ',') ('その他のエラー/{0}' -f $debugLog[0].LogName)
    } else {
        Write-Skip '分析とデバッグのログ' 'この PC にないため'
    }
    # Windows でのテストで見つかった問題: -FilterHashtable だと、権限のないログが「該当なし」になっていた
    $run = Invoke-Tool 'security' (@('-LogName', 'Security,System') + $liveRange)
    $x = Read-Result $run.Dir
    if ($isAdmin) {
        Assert-Equal '管理者は Security を読める' (@($x.Skipped).Count) 0
        Assert-Equal '管理者には案内を出さない' $run.Log.Contains('※ 管理者として実行すると') 'False'
        $run = Invoke-Tool 'security2' @('-LogName', 'Security', '-Level', 'Information', '-Since', '1h')
        Assert-Equal '管理者: レベル 0 の監査イベントを「情報」として読む' ($run.Code -eq 2 -or (@(@((Read-Result $run.Dir).Events) | Where-Object { $_.'レベル' -ne '情報' }).Count -eq 0)) 'True'
    } else {
        Assert-Equal '管理者でないときの Security は権限なし' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', $_.'対象' }) -join ',') '権限なし/Security'
        Assert-Equal 'スキップの行と管理者の案内' ('{0}/{1}' -f $run.Log.Contains('スキップ 1 件（権限なし）'), $run.Log.Contains('※ 管理者として実行すると、Security などのログも読めます')) 'True/True'
        $run = Invoke-Tool 'security2' (@('-LogName', 'Security') + $liveRange)
        Assert-Equal 'Security だけ: 止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('Security: 権限なし（アクセスが拒否されました）'), $run.Log.Contains('※ 管理者として実行すると')) '1/True/True'
        Assert-Equal 'Security だけ: 結果を残さない' ([bool]$run.Dir) 'False'
    }

    Write-Section '止まるとき'
    $stopCases = @(
        @('ないログだけ', @('-LogName', 'NoSuchLog'), 'NoSuchLog: ログがない（この PC にはないログです）'),
        @('ログの名前が空', @('-LogName', ' , '), '-LogName を指定してください'),
        @('レベルの誤り', @('-Level', 'Fatal'), '-Level は Critical'),
        @('レベルが空', @('-Level', ' , '), '-Level を指定してください'),
        @('期間の誤り', @('-Since', 'yesterday'), '-Since は "2026-10-31 09:00"'),
        @('-Until の誤り', @('-Until', 'tomorrow'), '-Until は "2026-10-31 09:00"'),
        @('ありえない日付', @('-Since', '2026-13-01'), '-Since は "2026-10-31 09:00"'),
        @('期間が逆', @('-Since', '2h', '-Until', '3h'), 'と同じか、それより後になっています'),
        @('-Since が未来', @('-Since', '2099-01-01'), '今（'),
        @('-Since と -Until が同じ', @('-Since', '2026-06-15', '-Until', '2026-06-15'), 'と同じか、それより後になっています'),
        @('-Id の誤り（大きすぎる）', @('-Id', '70000'), '-Id は 0～65535'),
        @('-Id の誤り（文字）', @('-Id', 'abc'), '-Id は 0～65535'),
        @('-Id の誤り（マイナス）', @('-Id', '1,-5'), '-Id は 0～65535'),
        @('-ExcludeId の誤り', @('-ExcludeId', '99999'), '-ExcludeId は 0～65535'),
        @('.evtx がない', @('-Path', (Join-Path $workRoot 'no\such.evtx')), '読める .evtx ファイルがありません')
    )
    foreach ($case in $stopCases) {
        $run = Invoke-Plain ($case[1] + @('-NoSave'))
        Assert-Equal $case[0] ('{0}/{1}' -f $run.Code, $run.Log.Contains($case[2])) '1/True'
    }
    $run = Invoke-Plain @('positional-value')
    Assert-Equal '名前のない引数は受け付けない' $run.Code 1

    Write-Section '保存'
    $run = Invoke-Plain ($mainArgs + @('-NoSave', '-OutputFolder', "$workRoot\out\nosave"))
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}' -f $run.Code, (Test-Path -LiteralPath "$workRoot\out\nosave")) '0/False'
    # 既定の保存先はスクリプトのフォルダ
    $run = Invoke-Plain $mainArgs
    Assert-Equal '既定の保存先はスクリプトのフォルダ' ('{0}/{1}' -f $run.Code, (Get-ResultFileNames (Get-RunDir (Split-Path -Parent $tool)))) '0/events.csv,skipped.csv,summary.csv,summary.txt'
    # 同じ秒に実行しても、前の結果を消さない（Ctrl+C や保存の失敗のときはフォルダごと消すため）
    $sameDir = Join-Path $workRoot 'out\same'
    [void][IO.Directory]::CreateDirectory($sameDir)
    $stampNow = Get-Date
    foreach ($s in 0..40) {
        $older = Join-Path $sameDir ('EventLog_{0}_{1}' -f $env:COMPUTERNAME, $stampNow.AddSeconds($s).ToString('yyyyMMdd_HHmmss'))
        [void][IO.Directory]::CreateDirectory($older)
        [IO.File]::WriteAllText("$older\keep.txt", 'keep')
    }
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @mainArgs -OutputFolder $sameDir 2>&1
    Assert-Equal '同じ名前のフォルダがあれば _2 を付ける' (@(Get-ChildItem -LiteralPath $sameDir -Directory -Filter 'EventLog_*_2').Count) 1
    Assert-Equal '前の結果を残す' (@(Get-ChildItem -LiteralPath $sameDir -Recurse -Filter 'keep.txt').Count) 41
    $fileOut = Join-Path $workRoot 'out\afile'
    [IO.File]::WriteAllText($fileOut, 'x')
    $run = Invoke-Plain ($mainArgs + @('-OutputFolder', $fileOut))
    Assert-Equal '保存先がファイル: 警告を出して画面に表示' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため、結果は保存しません'), $run.Log.Contains('■ 集計（件数の多い順）')) '0/True/True'
    $letter = Get-FreeLetter
    if ($letter) {
        $run = Invoke-Plain ($mainArgs + @('-OutputFolder', "${letter}:\nope"))
        Assert-Equal 'ないドライブ: 警告を出して画面に表示' ('{0}/{1}' -f $run.Code, $run.Log.Contains('保存先に書き込めないため')) '0/True'
    } else {
        Write-Skip 'ないドライブ' '空いているドライブ文字がないため'
    }
    $patched = New-PatchedTool 'savefail' "Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'summary.csv'))" "throw 'disk full'; Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'summary.csv'))"
    $run = Invoke-Tool 'savefail' (@('-Path', $many) + $range) $patched
    Assert-Equal '保存に失敗: 警告を出して書きかけを残さない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした: disk full'), [bool]$run.Dir) '0/True/False'
    Assert-Equal '保存に失敗: すべて画面に出す' ('{0}/{1}' -f [regex]::Matches($run.Log, 'wb-many\s+\d+\s+\d{4}-\d{2}-\d{2} \d{2}:\d{2}  ').Count, [regex]::Matches($run.Log, '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}  ').Count) '36/38'
    Assert-Equal '保存に失敗: 保存先の行を出さない' $run.Log.Contains('保存先:') 'False'
    $patched = New-PatchedTool 'crash' '$sorted = @($events | Sort-Object -Property Time -Descending)' "throw 'boom'"
    $run = Invoke-Tool 'crash' $mainArgs $patched
    Assert-Equal '想定外のエラー: 結果を残さず、中断とは区別する' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('エラーで止まりました。結果は保存していません。'), $run.Log.Contains('中断しました'), [bool]$run.Dir) '1/True/False/False'

    Write-Section 'コマンドプロンプトから起動'
    $cmdHead = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $rangeText = '-Since "2026-06-15 09:00" -Until "2026-06-15 11:00"'
    $run = Invoke-ViaCmd ('{0} -NoSave {1} -Level "Critical,Error" -Path "{2}\"' -f $cmdHead, $rangeText, $treeDir)
    Assert-Equal 'カンマ区切りの -Level と、最後の引数が \"' ('{0}/{1}' -f $run.Code, $run.Log.Contains('イベント 3 件')) '0/True'
    # 後ろの引数に引用符があると、引数の組み立てが崩れて PowerShell のエラーになる（ツールの前で止まる）ので、引用符のない引数を続ける
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -NoSave -Since 3d' -f $cmdHead, $treeDir)
    Assert-Equal '-Path の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('最後に \ を付けないでください')) '1/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -OutputFolder "{2}\" -Since 3d' -f $cmdHead, $main, "$workRoot\out\cmdq")
    Assert-Equal '-OutputFolder の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-OutputFolder の値に引用符')) '1/True'
    $run = Invoke-ViaCmd ("{0} -NoSave -Path '{1}' -Since '2026-06-15' -Until '2026-06-16' -Level 'Critical,Error' -Id '9001,1'" -f $cmdHead, $main)
    Assert-Equal "cmd の ' はそのまま届く" ('{0}/{1}' -f $run.Code, $run.Log.Contains('イベント 4 件')) '0/True'
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}" -Since 2000-01-01 -Until 2000-01-02' -f $cmdHead, $main)
    Assert-Equal 'cmd で受け取る終了コード（該当なしは 2）' $run.Code 2
    $run = Invoke-ViaCmd ('{0} -NoSave -LogName NoSuchLog' -f $cmdHead)
    Assert-Equal 'cmd で受け取る終了コード（止まると 1）' $run.Code 1
    $spaced = Join-Path $workRoot 'スペース ある\in.evtx'
    New-EvtxFile $spaced @(New-TestEvent (Get-TestTime '10:00') 2 91)
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}" {2}' -f $cmdHead, $spaced, $rangeText)
    Assert-Equal 'cmd からスペースと日本語を含むパス' ('{0}/{1}' -f $run.Code, $run.Log.Contains('イベント 1 件')) '0/True'

    Write-Section 'Ctrl+C で中断'
    # 読むのに時間がかかるよう、同じ .evtx のコピーをたくさん置いて読ませる
    $slowDir = Join-Path $workRoot 'slow'
    [void][IO.Directory]::CreateDirectory($slowDir)
    foreach ($i in 1..200) { [IO.File]::Copy($bulk, "$slowDir\slow$i.evtx") }
    $stopScript = New-HelperScript 'stop-during-read.ps1' @(
        'param($Tool, $LogDir, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $LogDir).AddParameter(''Since'', ''2026-06-15 09:00'').AddParameter(''Until'', ''2026-06-15 11:00'').AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'Start-Sleep -Milliseconds 2500'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
    )
    $result = Invoke-Helper $stopScript @('-Tool', $tool, '-LogDir', $slowDir, '-OutRoot', (Join-Path $workRoot 'out\ctrlc'))
    if ($result['RUNNING'] -eq 'True') {
        Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $result['BEFORE'], $result['AFTER']) '1/0'
        Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' '2.5 秒より前に読み終わったため'
    }

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -Since '2026-06-15 09:00' -Until '2026-06-15 11:00' -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $bulk.Replace("'", "''"), $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        $summaryText = $x.Lines -join "`n"
        Assert-Equal 'fi-FI でも期間の行は yyyy-MM-dd HH:mm' $summaryText.Contains('期間: 2026-06-15 09:00 ～ 2026-06-15 11:00') 'True'
        Assert-Equal 'fi-FI でも件数の区切りはカンマ' $summaryText.Contains('イベント 1,234 件（エラー 1,234）') 'True'
        Assert-Equal 'fi-FI でも所要時間の小数点は .' ([regex]::IsMatch($summaryText, '所要 \d+\.\d秒')) 'True'
        Assert-Equal 'fi-FI でも events.csv の日時は yyyy/MM/dd HH:mm:ss' (@($x.Events)[0].'日時' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
        Assert-Equal 'fi-FI でも summary.csv の件数の区切りはカンマ' (@($x.Groups)[0].'件数') '1,234'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }

    Write-Section 'レビューで見つかった問題の再発防止'
    # 1. ="..." の中の " を重ねる（上の「メッセージ」で確かめている）
    # 2. ID が多いと検索の式の上限を超える（上の「イベント ID」で確かめている）
    # 3. [a].evtx が隣の a.evtx に一致する（上の「ファイルとフォルダ」で確かめている）
    # 4. Security の権限エラーと管理者のヒント（上の「この PC のイベントログ」で確かめている）
    # 5. 保存の失敗と想定外のエラーを、中断と区別する（上の「保存」で確かめている）
} finally {
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    if ($substDrive) { $null = & subst.exe $substDrive /D 2>&1 }
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
