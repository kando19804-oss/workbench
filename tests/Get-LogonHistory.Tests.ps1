<#
.SYNOPSIS
    Get-LogonHistory の回帰テストです。

.DESCRIPTION
    テスト用の Security ログとセッションのログ（TerminalServices-LocalSessionManager/Operational）の .evtx を
    一時フォルダに組み立て、Get-LogonHistory の -Path で読んだ結果を期待値と照合します。
    本物のイベントログには書き込みません。この PC のイベントログは、管理者のときだけ読みます（管理者でないときは、止まることを確かめます）。
    実際の記録はテストに入れません。記録の形（値の名前と型）は、実際の Windows の記録に合わせています。
    「権限なし」を再現するため、テストが作ったファイルとフォルダにだけアクセス拒否を設定します（終了時に戻します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Get-LogonHistory を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Get-LogonHistory\Get-LogonHistory.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Get-LogonHistory_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Get-LogonHistory.ps1'
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
$inv       = [Globalization.CultureInfo]::InvariantCulture
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

# 処理を 1 か所だけ書き換えたツールのコピーを作る（本物のログでは起きない条件を確かめるため）
function New-PatchedTool([string]$Name, [string]$From, [string]$To) {
    $text = [IO.File]::ReadAllText($repoTool)
    if (-not $text.Contains($From)) { throw "書き換える場所が見つかりません: $From" }
    $file = Join-Path $workRoot "patched\$Name\Get-LogonHistory.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text.Replace($From, $To), $utf8Bom)
    $file
}

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'LogonHistory_*' -ErrorAction SilentlyContinue | Sort-Object Name)
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
    [pscustomobject]@{
        Sessions = @(Read-CsvOrEmpty 'sessions.csv'); Events = @(Read-CsvOrEmpty 'events.csv'); Users = @(Read-CsvOrEmpty 'users.csv')
        Skipped = @(Read-CsvOrEmpty 'skipped.csv'); Lines = $lines
    }
}

# sessions.csv の行を「ログオンの時刻（HH:mm）|ユーザー|種類|状態」にして、ファイルの順に並べる
function Get-SessionKeys($Rows) { (@($Rows) | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.'ログオン'.Substring(11, 5), $_.'ユーザー', $_.'種類', $_.'状態' }) -join ',' }

# events.csv の行を「日時（HH:mm）操作」にして、ファイルの順に並べる
function Get-EventKeys($Rows) { (@($Rows) | ForEach-Object { '{0} {1}' -f $_.'日時'.Substring(11, 5), $_.'操作' }) -join ',' }

# 条件に合う最初の行
function Get-Row($Rows, [scriptblock]$Where) { @(@($Rows) | Where-Object $Where)[0] }

function Get-ResultFileNames([string]$Dir) {
    if (-not $Dir) { return '' }
    (@(Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { $_.Name } | Sort-Object)) -join ','
}

# 画面の幅（ツールと同じく、全角を 2 桁、半角カタカナを 1 桁として数える）
function Get-DisplayWidth([string]$Text) {
    $width = 0
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int]$ch
        if ($code -lt 0x100 -or ($code -ge 0xFF61 -and $code -le 0xFF9F)) { $width += 1 }
        elseif ($code -ge 0xDC00 -and $code -le 0xDFFF) { }
        else { $width += 2 }
    }
    $width
}

# summary.txt の「■ ログオン」の表の行（見出しを含む）
function Get-SessionTableLines($Lines) {
    $inTable = $false
    foreach ($line in $Lines) {
        if ($line.StartsWith('■ ログオン（')) { $inTable = $true; continue }
        if ($inTable) { if ($line -eq '' -or $line.StartsWith('■')) { break }; $line }
    }
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

# Security ログには書き込めない（管理者でも、監査の記録を好きな内容で作ることはできない）ため、ファイルを直接組み立てる。
# 形式は [MS-EVEN6] の BinXml。tests\Search-EventLog.Tests.ps1 の組み立て方を、名前付きの Data と UserData を持つ記録に広げた。
# 記録の形（イベント ID・Data の名前と型・UserData の形）ごとに、チャンク（64 KB）の中で最初の記録にテンプレートを置き、
# 同じチャンクの同じ形の記録はそのテンプレートを指す（Windows が書く .evtx と同じ作り）

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

# 値の型（BinXml の型）。0x01 文字、0x04 UInt8、0x06 UInt16、0x08 UInt32、0x0A UInt64、0x11 FileTime、
# 0x13 SID、0x14 HexInt32（ログオンの失敗の状態コード）、0x15 HexInt64（ログオン ID）
function Get-ValueBytes([int]$Type, $Value) {
    if ($null -eq $Value) { return , [byte[]]@() }
    switch ($Type) {
        0x01 { return , [Text.Encoding]::Unicode.GetBytes([string]$Value) }
        0x04 { return , [byte[]]@([byte]$Value) }
        0x06 { return , [BitConverter]::GetBytes([uint16]$Value) }
        0x08 { return , [BitConverter]::GetBytes([uint32]$Value) }
        0x0A { return , [BitConverter]::GetBytes([uint64]$Value) }
        0x11 { return , [BitConverter]::GetBytes([int64]$Value) }
        0x13 {
            if (-not [string]$Value) { return , [byte[]]@() }
            $sid = New-Object Security.Principal.SecurityIdentifier ([string]$Value)
            $bytes = New-Object byte[] $sid.BinaryLength
            $sid.GetBinaryForm($bytes, 0)
            return , $bytes
        }
        0x14 { return , [BitConverter]::GetBytes([uint32]$Value) }
        0x15 { return , [BitConverter]::GetBytes([uint64]$Value) }
    }
    throw "知らない型: $Type"
}

# 記録の形のキー（同じキーの記録は、同じテンプレートを使える）
function Get-EventShape($Event) {
    $parts = New-Object 'System.Collections.Generic.List[string]'
    $parts.Add(('{0}|{1}|{2}' -f $Event.Id, $Event.Provider, $Event.Channel))
    if ($null -ne $Event.Data) { $parts.Add('D'); foreach ($d in $Event.Data) { $parts.Add(('{0}:{1}' -f $d[0], $d[1])) } }
    if ($null -ne $Event.UserData) { $parts.Add(('U:{0}:{1}' -f $Event.UserData.Root, $Event.UserData.Ns)); foreach ($d in $Event.UserData.Items) { $parts.Add(('{0}:{1}' -f $d[0], $d[1])) } }
    $parts -join ';'
}

$script:templateCount = 0

# 記録を 1 つ組み立てる。Templates はチャンクの中で置いたテンプレート（形のキー → 位置と ID）。
# 新しく置いたテンプレートは、記録をこのチャンクに入れると決めてから、呼び出し側で Templates に足す
function New-EvtxRecord($Event, [long]$RecordId, [int]$Start, $Templates) {
    $shape = Get-EventShape $Event
    $known = $Templates.ContainsKey($shape)
    if ($known) {
        $templateAt = $Templates[$shape].Offset
        $templateId = $Templates[$shape].Id
    } else {
        $script:templateCount++
        $templateAt = $Start + 38
        $templateId = [BitConverter]::GetBytes([int](0x57420000 + $script:templateCount))
    }
    $rec = New-Object 'System.Collections.Generic.List[byte]'
    $fileTime = ([datetime]$Event.Time).ToUniversalTime().ToFileTimeUtc()
    $rec.AddRange([byte[]](0x2A, 0x2A, 0, 0, 0, 0, 0, 0))
    $rec.AddRange([BitConverter]::GetBytes([int64]$RecordId))
    $rec.AddRange([BitConverter]::GetBytes([int64]$fileTime))
    $rec.AddRange([byte[]](0x0F, 1, 1, 0, 0x0C, 1))
    $rec.AddRange([byte[]]$templateId)
    $rec.AddRange([BitConverter]::GetBytes([int]$templateAt))
    # 値の一覧（System の 11 個のあとに、Data か UserData の値）
    $values = New-Object 'System.Collections.Generic.List[object]'
    $values.Add(@(0x01, [string]$Event.Provider))
    $values.Add(@(0x06, [int]$Event.Id))
    $values.Add(@(0x04, [int]$Event.Version))
    $values.Add(@(0x04, [int]$Event.Level))
    $values.Add(@(0x06, [int]$Event.Task))
    $values.Add(@(0x04, 0))
    $values.Add(@(0x15, [uint64]$Event.Keywords))
    $values.Add(@(0x11, $fileTime))
    $values.Add(@(0x0A, [uint64]$RecordId))
    $values.Add(@(0x01, [string]$Event.Channel))
    $values.Add(@(0x01, [string]$Event.Computer))
    $items = if ($null -ne $Event.Data) { @($Event.Data) } elseif ($null -ne $Event.UserData) { @($Event.UserData.Items) } else { @() }
    foreach ($d in $items) { $values.Add(@($d[1], $d[2])) }
    if (-not $known) {
        $rec.AddRange([byte[]](0, 0, 0, 0))
        $rec.AddRange([byte[]]$templateId)
        $rec.AddRange((New-Object byte[] 12))
        $sizeAt = $rec.Count
        $rec.AddRange([byte[]](0, 0, 0, 0))
        $rec.AddRange([byte[]](0x0F, 1, 1, 0))
        Add-XmlElement $rec $Start 'Event' @('xmlns', 'http://schemas.microsoft.com/win/2004/08/events/event') $null {
            Add-XmlElement $rec $Start 'System' $null $null {
                Add-XmlElement $rec $Start 'Provider' @('Name', 0, 0x01) $null $null
                Add-XmlElement $rec $Start 'EventID' $null @(1, 0x06) $null
                Add-XmlElement $rec $Start 'Version' $null @(2, 0x04) $null
                Add-XmlElement $rec $Start 'Level' $null @(3, 0x04) $null
                Add-XmlElement $rec $Start 'Task' $null @(4, 0x06) $null
                Add-XmlElement $rec $Start 'Opcode' $null @(5, 0x04) $null
                Add-XmlElement $rec $Start 'Keywords' $null @(6, 0x15) $null
                Add-XmlElement $rec $Start 'TimeCreated' @('SystemTime', 7, 0x11) $null $null
                Add-XmlElement $rec $Start 'EventRecordID' $null @(8, 0x0A) $null
                Add-XmlElement $rec $Start 'Channel' $null @(9, 0x01) $null
                Add-XmlElement $rec $Start 'Computer' $null @(10, 0x01) $null
            }
            if ($null -ne $Event.Data) {
                $dataList = @($Event.Data)
                if ($dataList.Count -eq 0) {
                    Add-XmlElement $rec $Start 'EventData' $null $null $null
                } else {
                    Add-XmlElement $rec $Start 'EventData' $null $null {
                        for ($i = 0; $i -lt $dataList.Count; $i++) { Add-XmlElement $rec $Start 'Data' @('Name', [string]$dataList[$i][0]) @((11 + $i), $dataList[$i][1]) $null }
                    }
                }
            } elseif ($null -ne $Event.UserData) {
                $userData = $Event.UserData
                Add-XmlElement $rec $Start 'UserData' $null $null {
                    Add-XmlElement $rec $Start $userData.Root @('xmlns', $userData.Ns) $null {
                        $list = @($userData.Items)
                        for ($i = 0; $i -lt $list.Count; $i++) { Add-XmlElement $rec $Start ([string]$list[$i][0]) $null @((11 + $i), $list[$i][1]) $null }
                    }
                }
            }
        }
        $rec.Add([byte]0x00)
        Set-Int32At $rec $sizeAt ($rec.Count - $sizeAt - 4)
    }
    $encoded = @(foreach ($v in $values) { , (Get-ValueBytes $v[0] $v[1]) })
    $rec.AddRange([BitConverter]::GetBytes([int]$values.Count))
    for ($i = 0; $i -lt $values.Count; $i++) {
        $bytes = $encoded[$i]
        $rec.AddRange([BitConverter]::GetBytes([uint16]$bytes.Length))
        $rec.Add([byte]$(if ($bytes.Length) { $values[$i][0] } else { 0 }))
        $rec.Add([byte]0)
    }
    foreach ($bytes in $encoded) { if ($bytes.Length) { $rec.AddRange([byte[]]$bytes) } }
    $rec.Add([byte]0x00)
    while (($rec.Count + 4) % 8) { $rec.Add([byte]0) }
    $total = $rec.Count + 4
    $rec.AddRange([BitConverter]::GetBytes([int]$total))
    Set-Int32At $rec 4 $total
    [pscustomobject]@{ Bytes = $rec.ToArray(); Shape = $shape; NewTemplate = (-not $known); Offset = $templateAt; Id = $templateId }
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

# 記録を日時の順に並べて .evtx を作る（記録の番号は 1 から。同じ日時のものは、書いた順）
function New-EvtxFile([string]$File, [object[]]$Events) {
    $chunks    = New-Object 'System.Collections.Generic.List[object]'
    $records   = New-Object 'System.Collections.Generic.List[object]'
    $offsets   = New-Object 'System.Collections.Generic.List[int]'
    $templates = @{}
    $free      = 512
    $firstId   = 1L
    $id        = 0L
    $ordered   = @(for ($i = 0; $i -lt $Events.Count; $i++) { [pscustomobject]@{ Event = $Events[$i]; Time = [datetime]$Events[$i].Time; Index = $i } }) | Sort-Object -Property Time, Index
    foreach ($item in $ordered) {
        $ev = $item.Event
        $id++
        $rec = New-EvtxRecord $ev $id $free $templates
        if ($free + $rec.Bytes.Length -gt 65536) {
            $chunks.Add((New-EvtxChunk $records.ToArray() $firstId $offsets.ToArray() $free))
            $records.Clear(); $offsets.Clear(); $templates = @{}
            $free = 512
            $firstId = $id
            $rec = New-EvtxRecord $ev $id $free $templates
        }
        if ($rec.NewTemplate) { $templates[$rec.Shape] = @{ Offset = $rec.Offset; Id = $rec.Id } }
        $records.Add($rec.Bytes)
        $offsets.Add($free)
        $free += $rec.Bytes.Length
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

#endregion

#region 記録を作る

$SecurityProvider = 'Microsoft-Windows-Security-Auditing'
$LsmChannel       = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
$LsmFileName      = 'Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx'
$AuditSuccess     = [uint64]::Parse('8020000000000000', 'HexNumber')
$AuditFailure     = [uint64]::Parse('8010000000000000', 'HexNumber')
$Server01         = 'SERVER01.corp.example.local'
$script:computer  = $Server01

# 日時。"HH:mm"・"HH:mm:ss"・"HH:mm:ss.fff" は 2026-06-15、"-1 HH:mm" は前日、"yyyy-MM-dd HH:mm" はその日時
function Get-T([string]$Text) {
    $formats = [string[]]@('yyyy-MM-dd HH:mm:ss.fff', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm')
    if ($Text -match '^\d{4}-') { return [datetime]::ParseExact($Text, $formats, $inv, [Globalization.DateTimeStyles]::AssumeLocal) }
    $day = 15
    if ($Text -match '^([+-]\d+) (.*)$') { $day += [int]$Matches[1]; $Text = $Matches[2] }
    $date = (New-Object datetime 2026, 6, $day).ToString('yyyy-MM-dd', $inv)
    [datetime]::ParseExact("$date $Text", $formats, $inv, [Globalization.DateTimeStyles]::AssumeLocal)
}

# 「ドメイン\ユーザー名」をドメインとユーザー名に分ける（ドメインがないときは -）
function Split-Account([string]$Name) {
    $at = $Name.IndexOf('\')
    if ($at -lt 0) { return , @('-', $Name) }
    , @($Name.Substring(0, $at), $Name.Substring($at + 1))
}

function New-SecEvent([string]$When, [int]$Id, [int]$Version, $Data, [uint64]$Keywords, [string]$Computer) {
    @{ Time = (Get-T $When); Id = $Id; Version = $Version; Level = 0; Task = 12544; Keywords = $Keywords; Provider = $SecurityProvider; Channel = 'Security'; Computer = $Computer; Data = $Data; UserData = $null }
}

# ログオン（4624 の版 2）。Id と Linked はログオン ID、Elevated は %%1842（管理者の権限あり）か %%1843
function New-Logon([string]$When, [string]$Name, [string]$Sid, [uint64]$Id, [int]$Type, [string]$Ip = '-', [string]$Workstation = '-',
                   [uint64]$Linked = 0, [string]$Elevated = '%%1843', [string]$Process = 'User32 ', [string]$Auth = 'Negotiate', [string]$Computer = $script:computer) {
    $account = Split-Account $Name
    New-SecEvent $When 4624 2 @(
        @('SubjectUserSid', 0x13, 'S-1-5-18'), @('SubjectUserName', 0x01, 'SERVER01$'), @('SubjectDomainName', 0x01, 'CORP'), @('SubjectLogonId', 0x15, 999),
        @('TargetUserSid', 0x13, $Sid), @('TargetUserName', 0x01, $account[1]), @('TargetDomainName', 0x01, $account[0]), @('TargetLogonId', 0x15, $Id),
        @('LogonType', 0x08, $Type), @('LogonProcessName', 0x01, $Process), @('AuthenticationPackageName', 0x01, $Auth), @('WorkstationName', 0x01, $Workstation),
        @('LogonGuid', 0x01, '{00000000-0000-0000-0000-000000000000}'), @('IpAddress', 0x01, $Ip), @('IpPort', 0x01, '0'),
        @('TargetLinkedLogonId', 0x15, $Linked), @('ElevatedToken', 0x01, $Elevated)
    ) $AuditSuccess $Computer
}

# ログオフ（4634）
function New-Logoff([string]$When, [string]$Name, [string]$Sid, [uint64]$Id, [int]$Type, [string]$Computer = $script:computer) {
    $account = Split-Account $Name
    New-SecEvent $When 4634 0 @(@('TargetUserSid', 0x13, $Sid), @('TargetUserName', 0x01, $account[1]), @('TargetDomainName', 0x01, $account[0]), @('TargetLogonId', 0x15, $Id), @('LogonType', 0x08, $Type)) $AuditSuccess $Computer
}

# ユーザーがログオフした記録（4647）。LogonType はない
function New-UserLogoff([string]$When, [string]$Name, [string]$Sid, [uint64]$Id, [string]$Computer = $script:computer) {
    $account = Split-Account $Name
    New-SecEvent $When 4647 0 @(@('TargetUserSid', 0x13, $Sid), @('TargetUserName', 0x01, $account[1]), @('TargetDomainName', 0x01, $account[0]), @('TargetLogonId', 0x15, $Id)) $AuditSuccess $Computer
}

# ログオンの失敗（4625）。Status と SubStatus は 16 進数の文字（C000006D など）
function New-Failure([string]$When, [string]$Name, [int]$Type, [string]$Status, [string]$SubStatus, [string]$Ip = '-', [string]$Workstation = '-', [string]$Computer = $script:computer) {
    $account = Split-Account $Name
    New-SecEvent $When 4625 0 @(
        @('SubjectUserSid', 0x13, 'S-1-5-18'), @('SubjectUserName', 0x01, 'SERVER01$'), @('SubjectDomainName', 0x01, 'CORP'), @('SubjectLogonId', 0x15, 999),
        @('TargetUserSid', 0x13, 'S-1-0-0'), @('TargetUserName', 0x01, $account[1]), @('TargetDomainName', 0x01, $account[0]),
        @('Status', 0x14, [Convert]::ToUInt32($Status, 16)), @('FailureReason', 0x01, '%%2313'), @('SubStatus', 0x14, [Convert]::ToUInt32($SubStatus, 16)),
        @('LogonType', 0x08, $Type), @('LogonProcessName', 0x01, 'NtLmSsp '), @('AuthenticationPackageName', 0x01, 'NTLM'), @('WorkstationName', 0x01, $Workstation),
        @('IpAddress', 0x01, $Ip), @('IpPort', 0x01, '0')
    ) $AuditFailure $Computer
}

# OS の起動（4608）。値はない
function New-Boot([string]$When, [string]$Computer = $script:computer) { New-SecEvent $When 4608 0 @() $AuditSuccess $Computer }

# Security ログの消去（1102）。値は UserData の LogFileCleared にある
function New-Clear([string]$When, [string]$Name, [string]$Computer = $script:computer) {
    $account = Split-Account $Name
    @{ Time = (Get-T $When); Id = 1102; Version = 1; Level = 4; Task = 104; Keywords = [uint64]::Parse('4020000000000000', 'HexNumber'); Provider = 'Microsoft-Windows-Eventlog'; Channel = 'Security'; Computer = $Computer; Data = $null
       UserData = @{ Root = 'LogFileCleared'; Ns = 'http://manifests.microsoft.com/win/2004/08/windows/eventlog'
                     Items = @(@('SubjectUserSid', 0x13, 'S-1-5-21-1000-2000-3000-500'), @('SubjectUserName', 0x01, $account[1]), @('SubjectDomainName', 0x01, $account[0]), @('SubjectLogonId', 0x15, 4660)) } }
}

# セッションのログ（21 ログオン、23 ログオフ、24 切断、25 再接続）。23 には Address がない
function New-Lsm([string]$When, [int]$Id, [string]$User, [int]$Session, [string]$Address = '', [string]$Computer = $script:computer) {
    $items = New-Object 'System.Collections.Generic.List[object]'
    $items.Add(@('User', 0x01, $User))
    $items.Add(@('SessionID', 0x08, $Session))
    if ($Id -ne 23) { $items.Add(@('Address', 0x01, $Address)) }
    @{ Time = (Get-T $When); Id = $Id; Version = 0; Level = 4; Task = 0; Keywords = [uint64]::Parse('1000000000000000', 'HexNumber'); Provider = 'Microsoft-Windows-TerminalServices-LocalSessionManager'
       Channel = $LsmChannel; Computer = $Computer; Data = $null; UserData = @{ Root = 'EventXML'; Ns = 'Event_NS'; Items = $items.ToArray() } }
}

# 対象外のログの記録（System ログ）
function New-SystemEvent([string]$When) {
    @{ Time = (Get-T $When); Id = 6005; Version = 0; Level = 4; Task = 0; Keywords = [uint64]::Parse('8000000000000000', 'HexNumber'); Provider = 'EventLog'; Channel = 'System'; Computer = $script:computer; Data = @(); UserData = $null }
}

# ユーザーの SID
$sidYamada = 'S-1-5-21-1000-2000-3000-1101'
$sidSuzuki = 'S-1-5-21-1000-2000-3000-1102'
$sidTanaka = 'S-1-5-21-1000-2000-3000-1103'
$sidKato   = 'S-1-5-21-1000-2000-3000-1104'
$sidSvc    = 'S-1-5-21-1000-2000-3000-1200'
$sidAdmin  = 'S-1-5-21-4000-5000-6000-500'
$sidMs     = 'S-1-5-21-4000-5000-6000-1001'

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$userSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied  = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したパス（終了時に戻す）
[void][IO.Directory]::CreateDirectory((Split-Path -Parent $tool))
[IO.File]::Copy($repoTool, $tool)

try {
    #region 準備

    Write-Section '準備: テスト用の .evtx を作る'
    # 主な場面（SERVER01、2026-06-14 08:00 に起動）。期間は 2026-06-15 09:00～18:00
    $mainSec = @(
        (New-Boot '-1 08:00'),
        (New-Logon '-1 08:00:05' 'Window Manager\DWM-1' 'S-1-5-90-0-1' 0x100 2),                                          # 画面の表示のアカウント（除く）
        (New-Logon '-1 20:00' 'CORP\kato' $sidKato 0x500 2 '127.0.0.1' 'SERVER01'),                                         # 期間の前に終わった（出さない）
        (New-UserLogoff '-1 21:00' 'CORP\kato' $sidKato 0x500),
        (New-Logon '-1 22:00' 'CORP\kato' $sidKato 0x600 10 '10.0.0.11' 'PC-KATO'),                                        # 期間の中の再起動で終わった（出す）
        (New-Logon '08:30' 'CORP\yamada' $sidYamada 0x1000 10 '192.168.10.21' 'PC-YAMADA'),                                 # 期間の前から続くリモートデスクトップ
        (New-Failure '08:50' 'CORP\yamada' 3 'C000006D' 'C000006A' '192.168.10.99' 'PC-X'),                                 # 期間の前の失敗（出さない）
        (New-Clear '09:05' 'CORP\admin2'),
        (New-Logon '09:10' 'SERVER01\admin' $sidAdmin 0x2000 2 '127.0.0.1' 'SERVER01' 0x2001 '%%1843'),                    # UAC の 2 つのログオン
        (New-Logon '09:10' 'SERVER01\admin' $sidAdmin 0x2001 2 '127.0.0.1' 'SERVER01' 0x2000 '%%1842'),
        (New-Logon '09:10:01' 'SERVER01\admin' $sidAdmin 0x2100 7 '127.0.0.1' 'SERVER01'),                                  # サインインと同時のロック解除（数えない）
        (New-Logoff '09:10:01.050' 'SERVER01\admin' $sidAdmin 0x2100 7),
        (New-Logon '09:15' 'CORP\svc-app' $sidSvc 0x9000 3 '10.0.0.8' 'APP01' 0 '%%1843' 'NtLmSsp ' 'NTLM'),               # ネットワーク
        (New-Logoff '09:15:01' 'CORP\svc-app' $sidSvc 0x9000 3),
        (New-Logon '09:16' 'CORP\SERVER01$' 'S-1-5-21-1000-2000-3000-2001' 0x9100 3 '10.0.0.1'),                           # コンピューターのアカウント（除く）
        (New-Logon '09:17' 'NT AUTHORITY\SYSTEM' 'S-1-5-18' 0x3e7 5),                                                      # SYSTEM（除く）
        (New-Failure '09:20' 'CORP\yamada' 3 'C000006D' 'C000006A' '192.168.10.99' 'PC-X'),
        (New-Failure '09:21' 'CORP\yamada' 3 'C000006D' 'C000006A' '192.168.10.99' 'PC-X'),
        (New-Failure '09:22' 'CORP\yamada' 3 'C000006D' 'C000006A' '192.168.10.99' 'PC-X'),
        (New-Failure '09:25' 'SERVER01\nouser' 2 'C000006D' 'C0000064' '127.0.0.1' 'SERVER01'),                             # コンソールの失敗（127.0.0.1 は空欄）
        (New-Failure '09:26' 'CORP\tanaka' 3 'C0000234' '0' '10.0.0.7' 'PC-TANAKA'),
        (New-Failure '09:27' 'CORP\tanaka' 10 'C0000999' '0' '10.0.0.7' 'PC-TANAKA'),                                       # 表にない状態コード
        (New-Logon '10:15' 'SERVER01\admin' $sidAdmin 0x2300 7 '127.0.0.1' 'SERVER01' 0x2301 '%%1843'),                    # ロック解除（type 7 の組と type 11 の組）
        (New-Logon '10:15' 'SERVER01\admin' $sidAdmin 0x2301 7 '127.0.0.1' 'SERVER01' 0x2300 '%%1842'),
        (New-Logon '10:15:00.010' 'SERVER01\admin' $sidAdmin 0x2200 11 '127.0.0.1' 'SERVER01' 0x2201 '%%1843'),
        (New-Logon '10:15:00.010' 'SERVER01\admin' $sidAdmin 0x2201 11 '127.0.0.1' 'SERVER01' 0x2200 '%%1842'),
        (New-Logoff '10:15:00.050' 'SERVER01\admin' $sidAdmin 0x2300 7),
        (New-Logoff '10:15:00.050' 'SERVER01\admin' $sidAdmin 0x2301 7),
        (New-Logoff '10:15:00.050' 'SERVER01\admin' $sidAdmin 0x2200 2),
        (New-Logoff '10:15:00.050' 'SERVER01\admin' $sidAdmin 0x2201 2),
        (New-Logon '10:30' 'CORP\suzuki' $sidSuzuki 0x3000 2 '-' 'SERVER01' 0 '%%1843' 'seclogo'),                         # 3 秒で終わる runas（ログオンに数える）
        (New-Logoff '10:30:03' 'CORP\suzuki' $sidSuzuki 0x3000 2),
        (New-UserLogoff '11:05' 'SERVER01\admin' $sidAdmin 0x2000),
        (New-Logoff '11:05:01' 'SERVER01\admin' $sidAdmin 0x2000 2),
        (New-Logoff '11:05:01' 'SERVER01\admin' $sidAdmin 0x2001 2),
        (New-Logon '11:30' 'CORP\yamada' $sidYamada 0x1100 10 '192.168.10.21' 'PC-YAMADA'),                                 # 再接続の確かめ（0.4 秒で終わる）
        (New-Logoff '11:30:00.400' 'CORP\yamada' $sidYamada 0x1100 10),
        (New-UserLogoff '12:00' 'CORP\yamada' $sidYamada 0x1000),
        (New-Logoff '12:00:01' 'CORP\yamada' $sidYamada 0x1000 10),
        (New-Logon '13:00' 'CORP\suzuki' $sidSuzuki 0x5000 10 '10.0.0.5' 'PC-SUZUKI'),                                     # 切断したまま再起動
        (New-Logon '14:00' 'CORP\suzuki' $sidSuzuki 0x6000 2 '127.0.0.1' 'SERVER01'),                                      # ログオフせずに再起動
        (New-Boot '15:00'),
        (New-Logon '15:05' 'CORP\tanaka' $sidTanaka 0x6000 2 '127.0.0.1' 'SERVER01'),                                      # 再起動の後に、同じログオン ID
        (New-UserLogoff '15:10' 'CORP\tanaka' $sidTanaka 0x6000),
        (New-Logoff '15:10:01' 'CORP\tanaka' $sidTanaka 0x6000 2),
        (New-Logon '16:00' 'CORP\suzuki' $sidSuzuki 0x7000 10 '10.0.0.5' 'PC-SUZUKI'),                                     # 切断中
        (New-Logon '16:10' 'CORP\suzuki' $sidSuzuki 0x7100 7 '10.0.0.5' 'PC-SUZUKI'),                                      # リモートデスクトップの中のロック解除
        (New-Logon '16:10:00.035' 'CORP\suzuki' $sidSuzuki 0x7101 7 '10.0.0.5' 'PC-SUZUKI' 0 '%%1843' 'Negotiate'),       # （組にならない 2 つが 35 ミリ秒差で記録される）
        (New-Logoff '16:10:00.040' 'CORP\suzuki' $sidSuzuki 0x7100 7),
        (New-Logoff '16:10:00.070' 'CORP\suzuki' $sidSuzuki 0x7101 7),
        (New-Logon '17:00' 'CORP\tanaka' $sidTanaka 0x8000 10 '10.0.0.9' 'PC-TANAKA'),                                     # 記録なし
        (New-Logon '17:30' 'MicrosoftAccount\user@example.com' $sidMs 0xA000 11 '127.0.0.1' 'SERVER01'),                   # Microsoft アカウント
        (New-UserLogoff '17:45' 'SERVER01\localuser' $sidMs 0xA000),
        (New-Logoff '17:45:01' 'SERVER01\localuser' $sidMs 0xA000 2),
        (New-Logon '18:30' 'CORP\yamada' $sidYamada 0xB000 2 '127.0.0.1' 'SERVER01'),                                      # 期間の後（出さない）
        (New-Failure '18:30' 'CORP\yamada' 3 'C000006D' 'C000006A' '192.168.10.99' 'PC-X')
    )
    $mainLsm = @(
        (New-Lsm '08:30:03' 21 'CORP\yamada' 2 '192.168.10.21'),
        (New-Lsm '08:45' 24 'CORP\yamada' 2 '192.168.10.21'),        # 期間の前の切断・再接続（行に数えない）
        (New-Lsm '08:50' 25 'CORP\yamada' 2 '192.168.10.21'),
        (New-Lsm '09:10:02' 21 'SERVER01\admin' 1 'LOCAL'),
        (New-Lsm '11:00' 24 'CORP\yamada' 2 '192.168.10.21'),
        (New-Lsm '11:05:03' 23 'SERVER01\admin' 1),
        (New-Lsm '11:30:01' 25 'CORP\yamada' 2 '192.168.10.21'),
        (New-Lsm '12:00:03' 23 'CORP\yamada' 2),
        (New-Lsm '12:00:04' 24 'CORP\yamada' 2 '192.168.10.21'),      # ログオフの後の切断（数えない）
        (New-Lsm '13:00:02' 21 'CORP\suzuki' 3 '10.0.0.5'),
        (New-Lsm '13:30' 24 'CORP\suzuki' 3 '10.0.0.5'),
        (New-Lsm '16:00:03' 21 'CORP\suzuki' 2 '10.0.0.5'),
        (New-Lsm '16:30' 24 'CORP\suzuki' 2 '10.0.0.5')
    )
    $mainDir = Join-Path $workRoot 'main'
    $mainSecFile = Join-Path $mainDir 'Security.evtx'
    $mainLsmFile = Join-Path $mainDir $LsmFileName
    New-EvtxFile $mainSecFile $mainSec
    New-EvtxFile $mainLsmFile $mainLsm

    # 作った .evtx を、ツールと同じ読み方で、値の型まで読めるか
    $query  = New-Object Diagnostics.Eventing.Reader.EventLogQuery($mainSecFile, [Diagnostics.Eventing.Reader.PathType]::FilePath, '*[System[EventID=4624]]')
    $reader = New-Object Diagnostics.Eventing.Reader.EventLogReader($query)
    $selector = New-Object Diagnostics.Eventing.Reader.EventLogPropertySelector (, [string[]]@("Event/EventData/Data[@Name='TargetUserSid']", "Event/EventData/Data[@Name='TargetLogonId']", "Event/EventData/Data[@Name='LogonType']"))
    $first = $reader.ReadEvent()
    $values = $first.GetPropertyValues($selector)
    Assert-Equal '作った .evtx を読める（SID・ログオン ID・種類の型）' ('{0}/{1}/{2}' -f $values[0].GetType().Name, $values[1].GetType().Name, $values[2].GetType().Name) 'SecurityIdentifier/UInt64/UInt32'
    $first.Dispose(); $reader.Dispose()
    $count = 0
    $reader = New-Object Diagnostics.Eventing.Reader.EventLogReader((New-Object Diagnostics.Eventing.Reader.EventLogQuery($mainLsmFile, [Diagnostics.Eventing.Reader.PathType]::FilePath, '*')))
    while ($e = $reader.ReadEvent()) { $count++; $e.Dispose() }
    $reader.Dispose()
    Assert-Equal '作ったセッションのログを読める' $count $mainLsm.Count
    $range = @('-Since', '2026-06-15 09:00', '-Until', '2026-06-15 18:00')
    $mainArgs = @('-Path', $mainDir) + $range

    #endregion

    #region 主な場面

    Write-Section '既定の種類（対話・リモートデスクトップ・ロック解除）'
    $run = Invoke-Tool 'main' $mainArgs
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '保存先のフォルダの名前' ((Split-Path -Leaf $run.Dir) -replace '\d{8}_\d{6}$', 'DATE') ('LogonHistory_{0}_DATE' -f $env:COMPUTERNAME)
    Assert-Equal '保存するファイル' (Get-ResultFileNames $run.Dir) 'events.csv,sessions.csv,skipped.csv,summary.txt,users.csv'
    Assert-Equal '対象の行' $run.Log.Contains(('対象: {0}\Security.evtx ほか 1 ファイル   Security ログ: 2026-06-14 08:00 から残っています' -f $mainDir)) 'True'
    Assert-Equal '期間の行' $run.Log.Contains('期間: 2026-06-15 09:00 ～ 2026-06-15 18:00   種類: 対話, リモートデスクトップ, ロック解除') 'True'
    Assert-Equal '条件がないときは条件の行を出さない' $run.Log.Contains('条件:') 'False'
    Assert-Equal '合計の行' $run.Log.Contains('ログオン 10 件（ロック解除 2 回・切断 3 回・再接続 1 回）  失敗 6 件  所要 ') 'True'
    Assert-Equal 'ログの消去の注意' $run.Log.Contains('※ 2026-06-15 09:05 に Security ログが消去されています（CORP\admin2）') 'True'
    Assert-Equal '起動の記録があるときは「続いていたログオンは分からない」を出さない' $run.Log.Contains('より前から続いていたログオンは分かりません') 'False'
    Assert-Equal '残っている期間の注意を出さない' $run.Log.Contains('残っていません') 'False'
    Assert-Equal 'スキップの行を出さない' $run.Log.Contains('スキップ') 'False'
    Assert-Equal '保存先の行' $run.Log.Contains(('保存先: {0}' -f $run.Dir)) 'True'
    Assert-Equal 'ログオンの一覧（新しい順・期間に重なるもの・状態）' (Get-SessionKeys $x.Sessions) ('17:30|MicrosoftAccount\user@example.com|対話|ログオフ,17:00|CORP\tanaka|リモートデスクトップ|記録なし,' +
        '16:00|CORP\suzuki|リモートデスクトップ|切断中,15:05|CORP\tanaka|対話|ログオフ,14:00|CORP\suzuki|対話|再起動,13:00|CORP\suzuki|リモートデスクトップ|再起動,' +
        '10:30|CORP\suzuki|対話（runas）|ログオフ,09:10|SERVER01\admin|対話|ログオフ,08:30|CORP\yamada|リモートデスクトップ|ログオフ,22:00|CORP\kato|リモートデスクトップ|再起動')

    $yamada = Get-Row $x.Sessions { $_.'ログオン ID' -eq '0x1000' }
    Assert-Equal '期間の前から続くリモートデスクトップ' ('{0}|{1}|{2}|{3}|{4}|{5}' -f $yamada.'ログオン', $yamada.'ログオフ', $yamada.'時間', $yamada.'接続元 IP アドレス', $yamada.'接続元の名前', $yamada.'管理者権限') '2026/06/15 08:30:00|2026/06/15 12:00:00|="3:30:00"|="192.168.10.21"|PC-YAMADA|いいえ'
    Assert-Equal '切断・再接続は期間の中のものだけを行に数える' ('{0}/{1}/{2}/{3}' -f $yamada.'切断の回数', $yamada.'再接続の回数', $yamada.'最後の切断', $yamada.'セッション ID') '1/1/2026/06/15 11:00:00/2'
    Assert-Equal '認証とコンピューター' ('{0}|{1}' -f $yamada.'認証', $yamada.'コンピューター') "Negotiate|$Server01"
    $admin = Get-Row $x.Sessions { $_.'ユーザー' -eq 'SERVER01\admin' }
    Assert-Equal 'UAC の 2 つのログオンを 1 行に' ('{0}|{1}|{2}|{3}' -f $admin.'ログオン ID', $admin.'管理者権限', $admin.'ログオフ', $admin.'時間') '0x2000 / 0x2001|はい|2026/06/15 11:05:00|="1:55:00"'
    Assert-Equal 'コンソールの 127.0.0.1 は空欄' ('[{0}]|{1}' -f $admin.'接続元 IP アドレス', $admin.'接続元の名前') '[]|SERVER01'
    Assert-Equal 'ロック解除は 1 回（サインインと同時のものは数えない）' $admin.'ロック解除の回数' '1'
    Assert-Equal 'コンソールもセッションのログと組にする' $admin.'セッション ID' '1'
    $runas = Get-Row $x.Sessions { $_.'ログオン ID' -eq '0x3000' }
    Assert-Equal '短い runas はログオンに数える' ('{0}|{1}' -f $runas.'種類', $runas.'時間') '対話（runas）|="0:00:03"'
    $reused = @($x.Sessions | Where-Object { $_.'ログオン ID' -eq '0x6000' } | ForEach-Object { '{0}|{1}|{2}' -f $_.'ユーザー', $_.'状態', $_.'ログオフ' }) -join ','
    Assert-Equal '再起動の後に使い回されたログオン ID を区別する' $reused 'CORP\tanaka|ログオフ|2026/06/15 15:10:00,CORP\suzuki|再起動|'
    $disc = Get-Row $x.Sessions { $_.'ログオン ID' -eq '0x5000' }
    Assert-Equal '切断したまま再起動したものは「再起動」' ('{0}|{1}|{2}' -f $disc.'状態', $disc.'切断の回数', $disc.'セッション ID') '再起動|1|3'
    $now = Get-Row $x.Sessions { $_.'ログオン ID' -eq '0x7000' }
    Assert-Equal '切断中' ('{0}|{1}|{2}|[{3}]|[{4}]' -f $now.'状態', $now.'切断の回数', $now.'最後の切断', $now.'ログオフ', $now.'時間') '切断中|1|2026/06/15 16:30:00|[]|[]'
    Assert-Equal '35 ミリ秒差の 2 つのロック解除は 1 回' $now.'ロック解除の回数' '1'
    $ms = Get-Row $x.Sessions { $_.'ログオン ID' -eq '0xa000' }
    Assert-Equal 'Microsoft アカウント（ログオフの名前が違っても SID で組にする）' ('{0}|{1}' -f $ms.'ログオフ', $ms.'時間') '2026/06/15 17:45:00|="0:15:00"'

    Assert-Equal 'events.csv（期間の中・新しい順）' (Get-EventKeys $x.Events) ('17:45 ログオフ,17:30 ログオン,17:00 ログオン,16:30 切断,16:10 ロック解除,16:00 ログオン,15:10 ログオフ,15:05 ログオン,15:00 OS の起動,' +
        '14:00 ログオン,13:30 切断,13:00 ログオン,12:00 ログオフ,11:30 再接続,11:05 ログオフ,11:00 切断,10:30 ログオフ,10:30 ログオン,10:15 ロック解除,' +
        '09:27 ログオンの失敗,09:26 ログオンの失敗,09:25 ログオンの失敗,09:22 ログオンの失敗,09:21 ログオンの失敗,09:20 ログオンの失敗,09:10 ログオン,09:05 ログの消去')
    $e = Get-Row $x.Events { $_.'操作' -eq 'ログオン' -and $_.'ユーザー' -eq 'SERVER01\admin' }
    Assert-Equal 'events.csv のログオンの行' ('{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f $e.'日時', $e.'種類', $e.'詳細', $e.'ログオン ID', $e.'セッション ID', $e.'イベント ID', $e.'コンピューター') "2026/06/15 09:10:00|対話|管理者権限: はい|0x2000|1|4624|$Server01"
    $e = Get-Row $x.Events { $_.'操作' -eq 'ロック解除' -and $_.'ユーザー' -eq 'SERVER01\admin' }
    Assert-Equal 'events.csv のロック解除の行'('{0}|{1}|{2}|{3}|{4}' -f $e.'日時', $e.'ユーザー', $e.'種類', $e.'ログオン ID', $e.'セッション ID') '2026/06/15 10:15:00|SERVER01\admin|ロック解除|0x2300|1'
    $e = Get-Row $x.Events { $_.'操作' -eq '再接続' }
    Assert-Equal 'events.csv の再接続（セッションのログの 25 を使い、Security の確かめのログオンと二重に数えない）' ('{0}|{1}|{2}|{3}' -f $e.'日時', $e.'イベント ID', $e.'ログオン ID', $e.'セッション ID') '2026/06/15 11:30:01|25|0x1000|2'
    $e = Get-Row $x.Events { $_.'操作' -eq 'ログオフ' -and $_.'ユーザー' -eq 'CORP\yamada' }
    Assert-Equal 'events.csv のログオフ（4647 を優先）' ('{0}|{1}' -f $e.'日時', $e.'イベント ID') '2026/06/15 12:00:00|4647'
    $failures = @($x.Events | Where-Object { $_.'操作' -eq 'ログオンの失敗' } | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.'ユーザー', $_.'種類', $_.'接続元 IP アドレス', $_.'接続元の名前', $_.'詳細' }) -join ','
    Assert-Equal 'events.csv の失敗（理由・コード・127.0.0.1 は空欄）' $failures ('CORP\tanaka|リモートデスクトップ|="10.0.0.7"|PC-TANAKA|その他（0xC0000999）,CORP\tanaka|ネットワーク|="10.0.0.7"|PC-TANAKA|アカウントがロックされている（0xC0000234）,' +
        'SERVER01\nouser|対話||SERVER01|ユーザー名がない（0xC000006D / 0xC0000064）,' + ((@(1..3) | ForEach-Object { 'CORP\yamada|ネットワーク|="192.168.10.99"|PC-X|パスワードが違う（0xC000006D / 0xC000006A）' }) -join ','))
    $e = Get-Row $x.Events { $_.'操作' -eq 'ログの消去' }
    Assert-Equal 'events.csv のログの消去' ('{0}|{1}|{2}' -f $e.'ユーザー', $e.'詳細', $e.'イベント ID') 'CORP\admin2|消した人: CORP\admin2|1102'
    $e = Get-Row $x.Events { $_.'操作' -eq 'OS の起動' }
    Assert-Equal 'events.csv の OS の起動' ('{0}|{1}' -f $e.'イベント ID', $e.'コンピューター') "4608|$Server01"

    $users = @($x.Users | ForEach-Object { '{0}|{1}|{2}' -f $_.'ユーザー', $_.'ログオン回数', $_.'失敗回数' }) -join ','
    Assert-Equal 'users.csv（ログオンの多い順、同じなら失敗の多い順、名前の順）' $users ('CORP\suzuki|4|0,CORP\tanaka|2|2,CORP\yamada|1|3,CORP\kato|1|0,MicrosoftAccount\user@example.com|1|0,SERVER01\admin|1|0,SERVER01\nouser|0|1')
    $u = Get-Row $x.Users { $_.'ユーザー' -eq 'CORP\tanaka' }
    Assert-Equal 'users.csv の日時と合計の時間（ログオフの記録があるものだけ）' ('{0}|{1}|{2}|{3}' -f $u.'最初のログオン', $u.'最後のログオン', $u.'合計の時間', $u.'最後の失敗') '2026/06/15 15:05:00|2026/06/15 17:00:00|="0:05:00"|2026/06/15 09:27:00'
    $u = Get-Row $x.Users { $_.'ユーザー' -eq 'SERVER01\nouser' }
    Assert-Equal 'users.csv: ログオンのないユーザーの合計の時間は空欄' ('[{0}]|[{1}]' -f $u.'合計の時間', $u.'最初のログオン') '[]|[]'

    Write-Section '画面と summary.txt'
    $summary = $x.Lines
    Assert-Equal 'summary.txt の先頭' $summary[0] 'ログオンの履歴 (Get-LogonHistory 1.0.0)'
    Assert-Equal 'summary.txt の最後は保存先' $summary[-1] ('保存先: {0}' -f $run.Dir)
    $table = @(Get-SessionTableLines $summary)
    Assert-Equal 'ログオンの表の見出し' (($table[0] -replace '\s+', ' ').Trim()) 'ログオン ログオフ 時間 ユーザー 種類 接続元 管理者'
    Assert-Equal 'ログオフが同じ日なら時刻だけ' (($table | Where-Object { $_ -like ' 2026-06-15 09:10*' }) -replace '\s+', ' ').Trim() '2026-06-15 09:10 11:05 1:55:00 SERVER01\admin 対話 はい'
    Assert-Equal 'ログオフの記録がないときは状態' (($table | Where-Object { $_ -like ' 2026-06-15 16:00*' }) -replace '\s+', ' ').Trim() '2026-06-15 16:00 （切断中） CORP\suzuki リモートデスクトップ 10.0.0.5 いいえ'
    # 表の桁がそろう（ログオフの列の始まりが、どの行も同じ幅の位置）
    $logoffAt = @($table | ForEach-Object { Get-DisplayWidth ([regex]::Match($_, '^ \S+( \S+)?\s{2,}').Value) } | Select-Object -Unique)
    Assert-Equal '表の桁がそろう（全角を 2 桁として）' ($logoffAt -join ',') '19'
    $failureTable = @($summary | Where-Object { $_ -match '^\s+\d+  (CORP|SERVER01)\\' } | ForEach-Object { ($_ -replace '\s+', ' ').Trim() })
    Assert-Equal 'ログオンの失敗の表（ユーザーと接続元ごと・件数の多い順・端末の名前で補う）' ($failureTable -join ',') ('3 CORP\yamada 192.168.10.99 2026-06-15 09:22 パスワードが違う,2 CORP\tanaka 10.0.0.7 2026-06-15 09:27 その他（0xC0000999）,1 SERVER01\nouser SERVER01 2026-06-15 09:25 ユーザー名がない')
    $wide = @($summary | Where-Object { $_.StartsWith(' ') -and (Get-DisplayWidth $_) -gt 120 })
    Assert-Equal '表の行は 120 桁に収まる' $wide.Count 0
    $head = New-Object byte[] 3
    $bomOk = @(foreach ($name in 'events.csv', 'sessions.csv', 'users.csv', 'skipped.csv', 'summary.txt') {
        $stream = [IO.File]::OpenRead((Join-Path $run.Dir $name))
        try { [void]$stream.Read($head, 0, 3) } finally { $stream.Dispose() }
        if (($head -join ',') -eq '239,187,191') { $name }
    }).Count
    Assert-Equal '5 つのファイルは BOM 付き UTF-8' $bomOk 5
    foreach ($case in @(
        @('sessions.csv', '"ログオン","ログオフ","時間","ユーザー","種類","接続元 IP アドレス","接続元の名前","管理者権限","状態","ロック解除の回数","切断の回数","再接続の回数","最後の切断","認証","ログオン ID","セッション ID","コンピューター"'),
        @('events.csv', '"日時","操作","ユーザー","種類","接続元 IP アドレス","接続元の名前","詳細","ログオン ID","セッション ID","イベント ID","コンピューター"'),
        @('users.csv', '"ユーザー","ログオン回数","最初のログオン","最後のログオン","合計の時間","失敗回数","最後の失敗"'),
        @('skipped.csv', '"理由","対象","詳細"')
    )) {
        Assert-Equal ('{0} の見出し' -f $case[0]) ([IO.File]::ReadAllLines((Join-Path $run.Dir $case[0]))[0]) $case[1]
    }

    Write-Section 'ログオンの種類'
    $run = Invoke-Tool 'network' ($mainArgs + @('-LogonType', 'Network'))
    $x = Read-Result $run.Dir
    $svc = Get-Row $x.Sessions { $_.'ユーザー' -eq 'CORP\svc-app' }
    Assert-Equal 'Network: ネットワークのログオンだけ（コンピューターのアカウントは除く）' ('{0}/{1}/{2}|{3}|{4}|{5}' -f $run.Code, @($x.Sessions).Count, $svc.'種類', $svc.'接続元 IP アドレス', $svc.'認証', $svc.'時間') '0/1/ネットワーク|="10.0.0.8"|NTLM|="0:00:01"'
    Assert-Equal 'Network: 合計の行に切断とロック解除を書かない' $run.Log.Contains('ログオン 1 件  失敗 6 件') 'True'
    Assert-Equal 'Network: セッションのログの注意を出さない' $run.Log.Contains('セッションのログ') 'False'
    $run = Invoke-Tool 'all' ($mainArgs + @('-LogonType', 'All'))
    $x = Read-Result $run.Dir
    Assert-Equal 'All: すべての種類（SYSTEM とコンピューターのアカウントは除く）' ('{0}/{1}/{2}' -f $run.Code, @($x.Sessions).Count, $run.Log.Contains('種類: すべて')) '0/11/True'
    $run = Invoke-Tool 'unlock' ($mainArgs + @('-LogonType', 'Unlock'))
    Assert-Equal 'Unlock: ロック解除だけ（ログオンの行はない）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('ログオン 0 件（ロック解除 2 回）  失敗 6 件')) '0/True'
    $run = Invoke-Tool 'japanese' ($mainArgs + @('-LogonType', '対話,リモートデスクトップ'))
    $x = Read-Result $run.Dir
    Assert-Equal '日本語の種類（ロック解除を除くと、行にも数えない）' ('{0}/{1}/{2}/{3}' -f $run.Log.Contains('種類: 対話, リモートデスクトップ'), $run.Log.Contains('ログオン 10 件（切断 3 回・再接続 1 回）'), (Get-Row $x.Sessions { $_.'ユーザー' -eq 'SERVER01\admin' }).'ロック解除の回数', @($x.Events | Where-Object { $_.'操作' -eq 'ロック解除' }).Count) 'True/True/0/0'
    $run = Invoke-Tool 'remote' ($mainArgs + @('-LogonType', 'RemoteDesktop'))
    $x = Read-Result $run.Dir
    Assert-Equal 'RemoteDesktop: リモートデスクトップだけ' (@($x.Sessions | ForEach-Object { $_.'ログオン'.Substring(11, 5) }) -join ',') '17:00,16:00,13:00,08:30,22:00'

    Write-Section '-User'
    foreach ($case in @(
        @('yamada', 'ログオン 1 件（ロック解除 0 回・切断 1 回・再接続 1 回）  失敗 3 件'),
        @('CORP\yamada', 'ログオン 1 件（ロック解除 0 回・切断 1 回・再接続 1 回）  失敗 3 件'),
        @('YAMADA', 'ログオン 1 件（ロック解除 0 回・切断 1 回・再接続 1 回）  失敗 3 件'),
        @('OTHER\yamada', '該当する記録はありません'),
        @('*admin*', 'ログオン 1 件（ロック解除 1 回・切断 0 回・再接続 0 回）  失敗 0 件'),
        @('localuser', 'ログオン 1 件（ロック解除 0 回・切断 0 回・再接続 0 回）  失敗 0 件'),
        @('yamada,suzuki', 'ログオン 5 件（ロック解除 1 回・切断 3 回・再接続 1 回）  失敗 3 件'),
        @('[a]*', '該当する記録はありません')
    )) {
        $run = Invoke-Plain ($mainArgs + @('-User', $case[0], '-NoSave'))
        Assert-Equal ('-User {0}' -f $case[0]) $run.Log.Contains($case[1]) 'True'
    }
    $run = Invoke-Plain ($mainArgs + @('-User', 'yamada,*admin*', '-NoSave'))
    Assert-Equal '-User の条件の行' $run.Log.Contains('条件: ユーザー yamada, *admin*') 'True'
    $run = Invoke-Plain ($mainArgs + @('-User', 'nobody', '-NoSave'))
    Assert-Equal '該当なしは終了コード 2' ('{0}/{1}' -f $run.Code, $run.Log.Contains('該当する記録はありません')) '2/True'

    Write-Section '期間'
    $run = Invoke-Tool 'nomatch' @('-Path', $mainDir, '-Since', '2000-01-01', '-Until', '2000-01-02')
    Assert-Equal '期間に記録がない: 終了コード 2 で、結果は保存する' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('該当する記録はありません'), (Get-ResultFileNames $run.Dir)) '2/True/events.csv,sessions.csv,skipped.csv,summary.txt,users.csv'
    $run = Invoke-Plain @('-Path', $mainDir, '-Since', '200d', '-NoSave')
    Assert-Equal '今からの時間（200d）と、期間の始めより前の記録がないときの注意' ('{0}/{1}' -f [regex]::IsMatch($run.Log, '期間: \d{4}-\d{2}-\d{2} \d{2}:\d{2} ～ \d{4}-\d{2}-\d{2} \d{2}:\d{2}（200 日）'), $run.Log.Contains('※ 2026-06-14 08:00 より前の記録は、Security ログに残っていません（上書きされています）')) 'True/True'
    $run = Invoke-Plain @('-Path', $mainDir, '-Since', '2026/6/15 9:00', '-Until', '2026/06/15 18:00', '-NoSave')
    Assert-Equal '日時の書き方（2026/6/15 9:00）' $run.Log.Contains('期間: 2026-06-15 09:00 ～ 2026-06-15 18:00') 'True'
    $run = Invoke-Plain @('-Path', $mainDir, '-Since', '2026-06-15 13:00', '-Until', '2026-06-15 14:30', '-NoSave')
    Assert-Equal '期間の後に再起動で終わったものも出す' ('{0}/{1}' -f $run.Log.Contains('2026-06-15 13:00  （再起動）'), $run.Log.Contains('2026-06-15 14:00  （再起動）')) 'True/True'
    # 期間の前の「記録なし」（17:00）は、いつ終わったか分からないので出さない。「切断中」（16:00）は出す
    $run = Invoke-Tool 'late' @('-Path', $mainDir, '-Since', '2026-06-15 17:10', '-Until', '2026-06-15 18:00')
    Assert-Equal '期間の前の記録なしは出さず、切断中は出す' (Get-SessionKeys (Read-Result $run.Dir).Sessions) '17:30|MicrosoftAccount\user@example.com|対話|ログオフ,16:00|CORP\suzuki|リモートデスクトップ|切断中'

    #endregion

    #region レビューで見つかった問題

    Write-Section '期間の前の OS の起動の記録がない（ログオンの記録が残っていないログオフ）'
    # Security ログは 07:00 から（起動の記録はない）。期間より前からのログオンは、ログオフだけが残っている
    $orphanDir = Join-Path $workRoot 'orphan'
    New-EvtxFile "$orphanDir\Security.evtx" @(
        (New-Failure '07:00' 'CORP\x' 3 'C000006D' 'C000006A' '10.0.0.99'),
        (New-Logon '09:30' 'CORP\kato' $sidKato 0xC100 7 '127.0.0.1' 'SERVER01'),
        (New-Logoff '09:30:00.040' 'CORP\kato' $sidKato 0xC100 7),
        (New-UserLogoff '10:00' 'CORP\kato' $sidKato 0xC000),
        (New-Logoff '10:00:01' 'CORP\kato' $sidKato 0xC000 2),
        (New-Logoff '10:30' 'Window Manager\DWM-1' 'S-1-5-90-0-1' 0xC400 2),
        (New-Logoff '10:40' 'CORP\kato' $sidKato 0xC500 7),
        (New-Logoff '10:50' 'CORP\svc-app' $sidSvc 0xC600 3),
        (New-Logoff '11:00' 'SERVER01\admin' $sidAdmin 0xC200 2),
        (New-Logoff '11:00:02' 'SERVER01\admin' $sidAdmin 0xC201 2),
        (New-UserLogoff '11:30' 'CORP\suzuki' $sidSuzuki 0xC300)
    )
    $orphanRange = @('-Path', $orphanDir, '-Since', '2026-06-15 09:00', '-Until', '2026-06-15 12:00')
    $run = Invoke-Tool 'orphan' $orphanRange
    $x = Read-Result $run.Dir
    Assert-Equal 'ログオフしか残っていなくても、該当する記録に数える（終了コード 0）' $run.Code 0
    Assert-Equal '注意を出す' $run.Log.Contains('※ 2026-06-15 07:00 より前から続いていたログオンは分かりません（ログオフだけ 3 件を events.csv に書いています）') 'True'
    $orphans = @($x.Events | Where-Object { $_.'操作' -eq 'ログオフ' } | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}' -f $_.'日時'.Substring(11, 5), $_.'ユーザー', $_.'種類', $_.'詳細', $_.'ログオン ID', $_.'イベント ID' }) -join ','
    Assert-Equal 'events.csv にログオフを書く（4647 を優先、UAC の 2 つは 1 つ、4647 だけなら種類は空欄、SYSTEM 系とロック解除は書かない）' $orphans ('11:30|CORP\suzuki||ログオンの記録が残っていません|0xc300|4647,' +
        '11:00|SERVER01\admin|対話|ログオンの記録が残っていません|0xc200|4634,10:00|CORP\kato|対話|ログオンの記録が残っていません|0xc000|4647')
    Assert-Equal 'ログオンの表と sessions.csv には出さない' ('{0}/{1}' -f @($x.Sessions).Count, $run.Log.Contains('ログオン 0 件（ロック解除 1 回・切断 0 回・再接続 0 回）  失敗 0 件')) '0/True'
    $run = Invoke-Plain ($orphanRange + @('-User', 'kato', '-NoSave'))
    Assert-Equal '-User で絞る' $run.Log.Contains('（ログオフだけ 1 件を events.csv に書いています）') 'True'
    $run = Invoke-Tool 'orphanall' ($orphanRange + @('-LogonType', 'All'))
    Assert-Equal '-LogonType All ならネットワークのログオフも書く' (@((Read-Result $run.Dir).Events | Where-Object { $_.'操作' -eq 'ログオフ' } | ForEach-Object { $_.'種類' }) -join ',') ',対話,ネットワーク,対話'
    $run = Invoke-Plain ($orphanRange + @('-LogonType', 'Network', '-NoSave'))
    Assert-Equal '-LogonType Network なら 4647 だけのものは書かない' ('{0}/{1}' -f $run.Code, $run.Log.Contains('（ログオフだけ 1 件を events.csv に書いています）')) '0/True'
    $run = Invoke-Plain @('-Path', $orphanDir, '-Since', '2026-06-15 06:00', '-Until', '2026-06-15 12:00', '-NoSave')
    Assert-Equal '期間の始めより前の記録がないときは、上書きの注意だけ' ('{0}/{1}' -f $run.Log.Contains('※ 2026-06-15 07:00 より前の記録は、Security ログに残っていません（上書きされています）'), $run.Log.Contains('続いていたログオン')) 'True/False'
    $run = Invoke-Plain ($orphanRange + @('-User', 'nobody', '-NoSave'))
    Assert-Equal 'ログオフがなくても注意は出す' ('{0}/{1}' -f $run.Code, $run.Log.Contains('※ 2026-06-15 07:00 より前から続いていたログオンは分かりません（Security ログに記録が残っていません）')) '2/True'

    Write-Section 'セッションのログのログオン（21）が残っていない'
    $lsmOrphanDir = Join-Path $workRoot 'lsmorphan'
    New-EvtxFile "$lsmOrphanDir\Security.evtx" @(
        (New-Boot '-1 08:00'),
        (New-Logon '-1 09:00' 'CORP\yamada' $sidYamada 0xD000 10 '192.168.10.21' 'PC-YAMADA')
    )
    New-EvtxFile "$lsmOrphanDir\$LsmFileName" @(
        (New-Lsm '09:30' 24 'CORP\yamada' 3 '192.168.10.21'),
        (New-Lsm '09:45' 24 'CORP\ghost' 5 '10.9.9.9'),
        (New-Lsm '09:50' 24 'CORP\ghost2' 6 'LOCAL'),
        (New-Lsm '10:00' 25 'CORP\yamada' 3 '192.168.10.22'),
        (New-Lsm '10:30' 24 'CORP\yamada' 3 '192.168.10.22')
    )
    $lsmRange = @('-Path', $lsmOrphanDir, '-Since', '2026-06-15 09:00', '-Until', '2026-06-15 12:00')
    $run = Invoke-Tool 'lsmorphan' $lsmRange
    $x = Read-Result $run.Dir
    $row = @($x.Sessions)[0]
    Assert-Equal '続いているログオンに付けて、最後が切断なら切断中' ('{0}|{1}|{2}|{3}|{4}|{5}' -f $row.'ログオン', $row.'状態', $row.'切断の回数', $row.'再接続の回数', $row.'最後の切断', $row.'セッション ID') '2026/06/14 09:00:00|切断中|2|1|2026/06/15 10:30:00|3'
    $unmatched = @($x.Events | Where-Object { $_.'ユーザー' -like 'CORP\ghost*' } | ForEach-Object { '{0}|{1}|[{2}]|{3}' -f $_.'ユーザー', $_.'種類', $_.'ログオン ID', $_.'セッション ID' }) -join ','
    Assert-Equal '組にできない切断も events.csv に書く（接続元の有無で種類を決め、セッション ID を書く）' $unmatched 'CORP\ghost2|対話|[]|6,CORP\ghost|リモートデスクトップ|[]|5'
    Assert-Equal '合計とセッションのログの注意' ('{0}/{1}' -f $run.Log.Contains('ログオン 1 件（ロック解除 0 回・切断 4 回・再接続 1 回）'), $run.Log.Contains('※ 2026-06-15 09:30 より前の切断・再接続は分かりません（セッションのログに残っていません）')) 'True/True'
    $run = Invoke-Plain ($lsmRange + @('-LogonType', 'RemoteDesktop', '-NoSave'))
    Assert-Equal '組にできないものは、選んだ種類のときだけ数える' $run.Log.Contains('切断 3 回') 'True'

    Write-Section 'コンピューターの名前を変えた（変える前の記録は前の名前）'
    $renameDir = Join-Path $workRoot 'rename'
    New-EvtxFile "$renameDir\Security.evtx" @(
        (New-Boot '08:00' 'DESKTOP-OLD'),
        (New-Logon '09:00' 'CORP\yamada' $sidYamada 0xE000 2 '127.0.0.1' 'DESKTOP-OLD' 0 '%%1843' 'User32 ' 'Negotiate' 'DESKTOP-OLD'),
        (New-Boot '12:00' 'PC01.corp.example.local'),
        (New-Logon '12:05' 'CORP\yamada' $sidYamada 0xE100 2 '127.0.0.1' 'PC01' 0 '%%1843' 'User32 ' 'Negotiate' 'PC01.corp.example.local'),
        (New-UserLogoff '12:30' 'CORP\yamada' $sidYamada 0xE100 'PC01.corp.example.local')
    )
    $run = Invoke-Tool 'rename' @('-Path', $renameDir, '-Since', '2026-06-15 07:00', '-Until', '2026-06-15 13:00')
    $x = Read-Result $run.Dir
    Assert-Equal '名前を変える前のログオンも「再起動」' (@($x.Sessions | ForEach-Object { '{0}|{1}|{2}' -f $_.'ログオン'.Substring(11, 5), $_.'状態', $_.'コンピューター' }) -join ',') '12:05|ログオフ|PC01.corp.example.local,09:00|再起動|DESKTOP-OLD'
    Assert-Equal '1 台なら表にコンピューターの列を足さない' $run.Log.Contains('管理者  コンピューター') 'False'
    Assert-Equal 'OS の起動の行は、記録に書かれた名前' (@($x.Events | Where-Object { $_.'操作' -eq 'OS の起動' } | ForEach-Object { $_.'コンピューター' }) -join ',') 'PC01.corp.example.local,DESKTOP-OLD'

    Write-Section '複数のコンピューターの .evtx'
    $multiDir = Join-Path $workRoot 'multi'
    New-EvtxFile "$multiDir\a\Security.evtx" @(
        (New-Boot '-1 08:00'),
        (New-Logon '10:00' 'CORP\yamada' $sidYamada 0xF000 2 '127.0.0.1' 'SERVER01'),
        (New-UserLogoff '10:30' 'CORP\yamada' $sidYamada 0xF000),
        (New-Failure '10:05' 'CORP\yamada' 3 'C000006D' 'C000006A' '192.168.10.99' 'PC-X'),
        (New-Logon '10:40' 'CORP\admin' 'S-1-5-21-1000-2000-3000-500' 0xF100 2 '127.0.0.1' 'SERVER01'),
        (New-UserLogoff '10:50' 'CORP\admin' 'S-1-5-21-1000-2000-3000-500' 0xF100),
        (New-Logon '10:45' 'CORP-FS\admin' 'S-1-5-21-7000-8000-9000-500' 0xF200 2 '127.0.0.1' 'SERVER01'),
        (New-UserLogoff '10:55' 'CORP-FS\admin' 'S-1-5-21-7000-8000-9000-500' 0xF200)
    )
    # SERVER02 は 09:30 からしか残っていない。いちばん古い記録は、構築したときの名前（WIN-BUILD）
    New-EvtxFile "$multiDir\b\Security.evtx" @(
        (New-Failure '09:30' 'CORP\yamada' 3 'C000006D' 'C000006A' '192.168.10.99' 'PC-X' 'WIN-BUILD'),
        (New-Logon '11:00' 'CORP\yamada' $sidYamada 0xF000 10 '192.168.10.21' 'PC-YAMADA' 0 '%%1843' 'User32 ' 'Negotiate' 'SERVER02.corp.example.local'),
        (New-UserLogoff '11:20' 'CORP\yamada' $sidYamada 0xF000 'SERVER02.corp.example.local')
    )
    $run = Invoke-Tool 'multi' @('-Path', $multiDir, '-Since', '2026-06-15 09:00', '-Until', '2026-06-15 12:00')
    $x = Read-Result $run.Dir
    Assert-Equal '同じログオン ID でも、コンピューターごとに組にする' (@($x.Sessions | ForEach-Object { '{0}|{1}|{2}' -f $_.'ログオン'.Substring(11, 5), $_.'ログオフ'.Substring(11, 5), $_.'コンピューター' }) -join ',') ('11:00|11:20|SERVER02.corp.example.local,10:45|10:55|SERVER01.corp.example.local,10:40|10:50|SERVER01.corp.example.local,10:00|10:30|SERVER01.corp.example.local')
    Assert-Equal '対象の行（ほか 1 ファイル、Security ログの日付は出さない）' ('{0}/{1}' -f $run.Log.Contains(('対象: {0}\a\Security.evtx ほか 1 ファイル期間' -f $multiDir)), $run.Log.Contains('から残っています')) 'True/False'
    Assert-Equal '表にコンピューターの列を足す' ('{0}/{1}' -f [regex]::IsMatch($run.Log, '管理者  コンピューター'), [regex]::IsMatch($run.Log, '理由 +コンピューター')) 'True/True'
    Assert-Equal '注意の先頭にコンピューターの名前（いちばん新しい記録の名前）' $run.Log.Contains('※ SERVER02: 2026-06-15 09:30 より前の記録は、Security ログに残っていません（上書きされています）') 'True'
    Assert-Equal 'セッションのログを読んでいない注意もコンピューターごと' ('{0}/{1}' -f $run.Log.Contains('※ SERVER01: リモートデスクトップのセッションのログを読んでいないため、切断・再接続は分かりません'), $run.Log.Contains('※ SERVER02: リモートデスクトップのセッションのログを読んでいないため')) 'True/True'
    $failureTable = @(Read-Result $run.Dir).Lines | Where-Object { $_ -match '^\s+1  CORP\\yamada' }
    Assert-Equal '失敗はコンピューターごとにまとめる' (@($failureTable).Count) 2
    # Sort-Object の文字列の比べ方は 5.1 と 7 で違い、ハイフンの扱いが変わる（CORP\admin と CORP-FS\admin の順が入れ替わっていた）
    Assert-Equal 'users.csv の名前の順は、5.1 と 7 で同じ（文字の番号の順）' (@($x.Users | ForEach-Object { $_.'ユーザー' }) -join ',') 'CORP\yamada,CORP-FS\admin,CORP\admin'

    Write-Section 'ファイルを読む順（5.1 と 7 で同じ）'
    $orderDir = Join-Path $workRoot 'fileorder'
    [void][IO.Directory]::CreateDirectory($orderDir)
    [IO.File]::Copy($mainSecFile, "$orderDir\Security.evtx")
    [IO.File]::Copy($mainSecFile, "$orderDir\Security-old.evtx")
    $run = Invoke-Plain (@('-Path', $orderDir, '-NoSave') + $range)
    Assert-Equal '名前は文字の番号の順（Security-old.evtx が先）' $run.Log.Contains(('対象: {0}\Security-old.evtx ほか 1 ファイル' -f $orderDir)) 'True'
    # セッションのログがないので、再接続は Security のすぐ終わるログオンから数える
    Assert-Equal '重なったファイルの同じ記録は 1 回だけ数える' $run.Log.Contains('ログオン 10 件（ロック解除 2 回・切断 0 回・再接続 1 回）  失敗 6 件') 'True'

    Write-Section '画面の幅（いちばん長くなる組み合わせ）'
    $widthDir = Join-Path $workRoot 'width'
    New-EvtxFile "$widthDir\Security.evtx" @(
        (New-Boot '2026-12-19 08:00'),
        (New-Logon '2026-12-20 09:00' 'CORP\administrator-longname' 'S-1-5-21-1000-2000-3000-1300' 0x10 10 'fe80::1234:5678:9abc:def0%12' 'PC'),
        (New-UserLogoff '2027-04-01 10:00' 'CORP\administrator-longname' 'S-1-5-21-1000-2000-3000-1300' 0x10),
        (New-Logon '2026-12-21 09:00' 'CORP\b' 'S-1-5-21-1000-2000-3000-1301' 0x20 10 '10.0.0.1' 'PC')
    )
    $run = Invoke-Tool 'width' @('-Path', $widthDir, '-Since', '2026-12-01', '-Until', '2027-04-30')
    $table = @(Get-SessionTableLines (Read-Result $run.Dir).Lines)
    Assert-Equal '年をまたぐログオフ・100 日を超える時間・長い名前と IPv6 を切る' (($table | Where-Object { $_ -like ' 2026-12-20*' }) -replace '\s+', ' ').Trim() '2026-12-20 09:00 2027-04-01 10:00 102日 1:00:00 CORP\administrator… リモートデスクトップ fe80::1234:56… いいえ'
    Assert-Equal 'それでも 120 桁に収まる' (@($table | ForEach-Object { Get-DisplayWidth $_ }) | Measure-Object -Maximum).Maximum 119
    $row = Get-Row (Read-Result $run.Dir).Sessions { $_.'ログオン ID' -eq '0x10' }
    Assert-Equal 'CSV には切らずに書く' ('{0}|{1}' -f $row.'ユーザー', $row.'接続元 IP アドレス') 'CORP\administrator-longname|fe80::1234:5678:9abc:def0%12'

    #endregion

    #region 一覧の行数

    Write-Section '一覧の行数'
    $manyEvents = New-Object 'System.Collections.Generic.List[object]'
    $manyEvents.Add((New-Boot '-1 08:00'))
    foreach ($i in 1..25) {
        $sid = 'S-1-5-21-1000-2000-3000-{0}' -f (2000 + $i)
        $name = 'CORP\user{0:00}' -f $i
        $manyEvents.Add((New-Logon ('{0:00}:{1:00}' -f (9 + [math]::Floor($i / 60)), ($i % 60)) $name $sid (0x30000 + $i) 2 '127.0.0.1' 'SERVER01'))
        $manyEvents.Add((New-UserLogoff ('{0:00}:{1:00}' -f (10 + [math]::Floor($i / 60)), ($i % 60)) $name $sid (0x30000 + $i)))
    }
    $manyDir = Join-Path $workRoot 'many'
    New-EvtxFile "$manyDir\Security.evtx" $manyEvents.ToArray()
    $manyArgs = @('-Path', $manyDir) + $range
    $run = Invoke-Tool 'many' $manyArgs
    $x = Read-Result $run.Dir
    # user05 は、ユーザーの表（名前の順）には出るが、ログオンの表（新しい順）では 21 件目以降になる。user25 はその逆
    $user05 = [regex]::Escape('CORP\user05')
    $user25 = [regex]::Escape('CORP\user25')
    Assert-Equal '保存したときは、画面は 20 件まで' ('{0}/{1}/{2}/{3}' -f $run.Log.Contains('（ほか 5 件は sessions.csv を参照してください）'), $run.Log.Contains('（ほか 5 件は users.csv を参照してください）'), [regex]::Matches($run.Log, $user05).Count, [regex]::Matches($run.Log, $user25).Count) 'True/True/1/1'
    Assert-Equal 'summary.txt は省略しない' ('{0}/{1}' -f @($x.Lines | Where-Object { $_ -like '*CORP\user05*' }).Count, @($x.Lines | Where-Object { $_ -like '*参照してください*' }).Count) '2/0'
    $run = Invoke-Plain ($manyArgs + @('-NoSave'))
    Assert-Equal '-NoSave のときは、すべて表示する' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('参照してください'), [regex]::Matches($run.Log, $user05).Count) '0/False/2'

    #endregion

    #region 読めないファイルとパス

    Write-Section '読めないファイル'
    $badDir = Join-Path $workRoot 'bad'
    [void][IO.Directory]::CreateDirectory($badDir)
    [IO.File]::WriteAllBytes("$badDir\broken.evtx", [byte[]](1..200 | ForEach-Object { 7 }))
    [IO.File]::WriteAllBytes("$badDir\zero.evtx", [byte[]]@())
    [IO.File]::WriteAllText("$badDir\text.evtx", 'not a log')
    New-EvtxFile "$badDir\system.evtx" @(New-SystemEvent '10:00')
    $run = Invoke-Tool 'baddir' (@('-Path', $badDir) + $range)
    Assert-Equal '読める .evtx がない: 止まって結果を残さない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains(('読める .evtx ファイルがありません: {0}' -f $badDir)), [bool]$run.Dir) '1/True/False'
    Assert-Equal '読めない理由を添える' ('{0}/{1}' -f $run.Log.Contains(('{0}\broken.evtx: ファイルを読めない' -f $badDir)), $run.Log.Contains(('{0}\system.evtx: 対象外のログ（Security でもセッションのログでもありません（System））' -f $badDir))) 'True/True'
    $run = Invoke-Tool 'badmix' (@('-Path', ('{0},{1}' -f $badDir, $mainSecFile)) + $range)
    $x = Read-Result $run.Dir
    Assert-Equal '読めないファイルは飛ばしてほかを読む' ('{0}/{1}' -f $run.Code, $run.Log.Contains('スキップ 4 件（ファイルを読めない 3・対象外のログ 1）')) '0/True'
    Assert-Equal 'skipped.csv の理由' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', (Split-Path -Leaf $_.'対象') }) -join ',') 'ファイルを読めない/broken.evtx,対象外のログ/system.evtx,ファイルを読めない/text.evtx,ファイルを読めない/zero.evtx'
    Assert-Equal 'セッションのログを読まなかった注意' $run.Log.Contains('※ リモートデスクトップのセッションのログを読んでいないため、切断・再接続は分かりません') 'True'
    $row = Get-Row $x.Sessions { $_.'ログオン ID' -eq '0x1000' }
    Assert-Equal 'セッションのログがないときは、Security のすぐ終わるログオンを再接続として行に数える' ('{0}/{1}/[{2}]' -f $row.'再接続の回数', $row.'切断の回数', $row.'セッション ID') '1/0/[]'
    Assert-Equal 'セッションのログがないときの再接続の行' ((@($x.Events | Where-Object { $_.'操作' -eq '再接続' }) | ForEach-Object { '{0}|{1}|{2}' -f $_.'日時', $_.'ログオン ID', $_.'イベント ID' }) -join ',') '2026/06/15 11:30:00|0x1100|4624'
    $run = Invoke-Tool 'lsmonly' (@('-Path', $mainLsmFile) + $range)
    Assert-Equal 'セッションのログだけ: 止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains(('-Path に、Security ログの .evtx がありません: {0}' -f $mainLsmFile)), [bool]$run.Dir) '1/True/False'
    $run = Invoke-Tool 'notfound' (@('-Path', "$workRoot\nothing.evtx") + $range)
    Assert-Equal 'ないファイル: 止まって理由を出す' ('{0}/{1}' -f $run.Code, $run.Log.Contains(('{0}\nothing.evtx: ファイルを読めない（見つかりません）' -f $workRoot))) '1/True'
    $emptyDir = Join-Path $workRoot 'emptydir'
    [void][IO.Directory]::CreateDirectory($emptyDir)
    $run = Invoke-Plain (@('-Path', $emptyDir, '-NoSave') + $range)
    Assert-Equal '.evtx のないフォルダ' ('{0}/{1}' -f $run.Code, $run.Log.Contains('読める .evtx ファイルがありません')) '1/True'

    # 最後のほうが壊れた .evtx: 新しい順に読めないので、起動を探せなくても最初から読む（壊れた場所の手前まで）
    $bigEvents = New-Object 'System.Collections.Generic.List[object]'
    $bigEvents.Add((New-Boot '-1 08:00'))
    foreach ($i in 1..300) {
        $at = (Get-T '09:00').AddSeconds($i * 20)
        $bigEvents.Add((New-Logon $at.ToString('yyyy-MM-dd HH:mm:ss', $inv) 'CORP\yamada' $sidYamada (0x40000 + $i) 2 '127.0.0.1' 'SERVER01'))
        $bigEvents.Add((New-Logoff $at.AddSeconds(10).ToString('yyyy-MM-dd HH:mm:ss', $inv) 'CORP\yamada' $sidYamada (0x40000 + $i) 2))
    }
    $bigFile = Join-Path $workRoot 'big\Security.evtx'
    New-EvtxFile $bigFile $bigEvents.ToArray()
    $bytes = [IO.File]::ReadAllBytes($bigFile)
    Assert-Equal '複数のチャンクの .evtx ができた' ($bytes.Length -gt (4096 + 65536 * 2)) 'True'
    $lastChunk = $bytes.Length - 65536
    foreach ($i in ($lastChunk + 10000)..($lastChunk + 10300)) { $bytes[$i] = 0xEE }
    $half = Join-Path $workRoot 'half\Security.evtx'
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $half))
    [IO.File]::WriteAllBytes($half, $bytes)
    $run = Invoke-Tool 'half' (@('-Path', $half) + $range)
    $x = Read-Result $run.Dir
    $halfCount = @($x.Sessions).Count
    Assert-Equal '最後が壊れた .evtx: 読めた分を出す' ('{0}/{1}' -f $run.Code, ($halfCount -gt 100 -and $halfCount -lt 300)) '0/True'
    Assert-Equal '最後が壊れた .evtx: skipped.csv に残す' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', ($_.'詳細' -replace '^(途中まで読んだ（記録 )[\d,]+( 件は結果に含む）).*$', '$1N$2') }) -join ',') 'ファイルを読めない/途中まで読んだ（記録 N 件は結果に含む）'
    # 期間に合う記録の前で壊れている: Security ログを 1 件も読めなかったので、「該当なし」と区別して止める
    $failEvents = New-Object 'System.Collections.Generic.List[object]'
    foreach ($i in 1..400) { $failEvents.Add((New-Failure (Get-T '07:00').AddSeconds($i).ToString('yyyy-MM-dd HH:mm:ss', $inv) 'CORP\x' 3 'C000006D' 'C000006A' '10.0.0.99')) }
    foreach ($i in 1..100) { $failEvents.Add((New-Failure (Get-T '10:00').AddSeconds($i).ToString('yyyy-MM-dd HH:mm:ss', $inv) 'CORP\x' 3 'C000006D' 'C000006A' '10.0.0.99')) }
    $allBad = Join-Path $workRoot 'allbad\Security.evtx'
    New-EvtxFile $allBad $failEvents.ToArray()
    $bytes = [IO.File]::ReadAllBytes($allBad)
    foreach ($chunkAt in @(for ($at = 4096 + 65536; $at -lt $bytes.Length; $at += 65536) { $at })) { foreach ($i in ($chunkAt + 10000)..($chunkAt + 10300)) { $bytes[$i] = 0xEE } }
    [IO.File]::WriteAllBytes($allBad, $bytes)
    $run = Invoke-Tool 'allbad' (@('-Path', $allBad) + $range)
    Assert-Equal 'Security ログを 1 件も読めない: 止まって結果を残さない' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('Security ログを読めませんでした'), $run.Log.Contains(('{0}: ファイルを読めない' -f $allBad)), [bool]$run.Dir) '1/True/True/False'

    Write-Section 'リンク・権限なし・パス'
    $linkDir = Join-Path $workRoot 'linkdir'
    [void][IO.Directory]::CreateDirectory($linkDir)
    [IO.File]::Copy($mainSecFile, "$linkDir\Security.evtx")
    $null = & cmd.exe /d /c mklink /J "$linkDir\loop" "$linkDir" 2>&1
    if ([IO.Directory]::Exists("$linkDir\loop")) {
        $run = Invoke-Tool 'link' (@('-Path', $linkDir) + $range)
        $x = Read-Result $run.Dir
        Assert-Equal 'ジャンクションはたどらない' ('{0}/{1}/{2}' -f $run.Log.Contains('ログオン 10 件'), ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', $_.'詳細' }) -join ','), $run.Log.Contains('スキップ 1 件（リンク）')) 'True/リンク/ジャンクション/True'
    } else {
        Write-Skip 'ジャンクション' 'ジャンクションを作れなかったため'
    }
    $denyDir = Join-Path $workRoot 'deny'
    [void][IO.Directory]::CreateDirectory("$denyDir\closed")
    [IO.File]::Copy($mainSecFile, "$denyDir\Security.evtx")
    [IO.File]::Copy($orphanDir + '\Security.evtx', "$denyDir\denied.evtx")
    [IO.File]::Copy($mainSecFile, "$denyDir\closed\Security.evtx")
    if ((Add-Deny "$denyDir\denied.evtx" '(R)') -and (Add-Deny "$denyDir\closed" '(RD)')) {
        $run = Invoke-Tool 'deny' (@('-Path', $denyDir) + $range)
        $x = Read-Result $run.Dir
        Assert-Equal '読めないファイルとフォルダは飛ばして、権限なし' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', (Split-Path -Leaf $_.'対象') }) -join ',') '権限なし/closed,権限なし/denied.evtx'
        Assert-Equal 'ほかは読む' ('{0}/{1}' -f $run.Log.Contains('ログオン 10 件'), $run.Log.Contains('スキップ 2 件（権限なし）')) 'True/True'
    } else {
        Write-Skip '権限なし' 'アクセス拒否を設定できなかったため'
    }
    $longDir = "$workRoot\" + ('L' * 120) + '\' + ('M' * 120)
    [void][IO.Directory]::CreateDirectory('\\?\' + $longDir)
    [IO.File]::Copy($mainSecFile, '\\?\' + "$longDir\Security.evtx")
    $run = Invoke-Tool 'longdir' (@('-Path', $longDir) + $range)
    Assert-Equal '260 文字を超えるパス（フォルダ）' $run.Log.Contains('ログオン 10 件') 'True'
    $run = Invoke-Tool 'longfile' (@('-Path', "$longDir\Security.evtx") + $range)
    Assert-Equal '260 文字を超えるパス（ファイル）' ('{0}/{1}' -f $run.Log.Contains('ログオン 10 件'), $run.Log.Contains(('対象: {0}\Security.evtx' -f $longDir))) 'True/True'
    $longOut = "$workRoot\" + ('O' * 120) + '\' + ('P' * 120)
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @mainArgs -OutputFolder $longOut 2>&1
    $longRun = @(if ([IO.Directory]::Exists('\\?\' + $longOut)) { [IO.Directory]::GetDirectories('\\?\' + $longOut) })
    Assert-Equal '260 文字を超える保存先' $(if ($longRun.Count -eq 1) { @([IO.Directory]::GetFiles($longRun[0])).Count } else { 0 }) 5
    $jpDir = Join-Path $workRoot 'スペース ある\ログ'
    [void][IO.Directory]::CreateDirectory($jpDir)
    [IO.File]::Copy($mainSecFile, "$jpDir\Security.evtx")
    $run = Invoke-Tool 'japanese path' (@('-Path', "$jpDir\Security.evtx") + $range)
    Assert-Equal '日本語とスペースを含むパス' $run.Log.Contains('ログオン 10 件') 'True'
    $bracketDir = Join-Path $workRoot 'br'
    [void][IO.Directory]::CreateDirectory("$bracketDir\[a]")
    [void][IO.Directory]::CreateDirectory("$bracketDir\a")
    [IO.File]::Copy($mainSecFile, "$bracketDir\[a]\Security.evtx")
    [IO.File]::Copy("$orphanDir\Security.evtx", "$bracketDir\a\Security.evtx")
    $run = Invoke-Plain (@('-Path', "$bracketDir\[a]\Security.evtx", '-NoSave') + $range)
    Assert-Equal '[a] を含むパス（隣の a を読まない）' $run.Log.Contains('ログオン 10 件') 'True'
    $run = Invoke-From "$bracketDir\[a]" (@('-Path', 'Security.evtx', '-NoSave') + $range)
    Assert-Equal '[a] のフォルダから相対パス' $run.Log.Contains('ログオン 10 件') 'True'
    $run = Invoke-Plain (@('-Path', ('{0},{1}' -f $mainSecFile, $mainLsmFile), '-NoSave') + $range)
    Assert-Equal '-Path にカンマで複数のファイル' $run.Log.Contains('ログオン 10 件（ロック解除 2 回・切断 3 回・再接続 1 回）') 'True'

    #endregion

    #region 引数の誤りと保存先

    Write-Section '引数の誤り'
    foreach ($case in @(
        @(@('-LogonType', 'Remote'), '-LogonType に使えない種類があります: Remote'),
        @(@('-LogonType', ','), '-LogonType を指定してください'),
        @(@('-Since', '2026-13-01'), '-Since の日時を読めません: 2026-13-01'),
        @(@('-Since', '2026-06-15 18:00', '-Until', '2026-06-15 09:00'), '-Since の日時（2026-06-15 18:00）が、-Until の日時（2026-06-15 09:00）と同じか、それより後になっています'),
        @(@('-Since', '2099-01-01'), 'と同じか、それより後になっています'),
        @(@('-Until', 'abc'), '-Until の日時を読めません: abc'),
        @(@('-NoSave', '-OutputFolder', "$workRoot\out\x"), '-NoSave のときは、-OutputFolder は指定できません')
    )) {
        $run = Invoke-Plain (@('-Path', $mainDir) + $case[0])
        Assert-Equal ('止まる: {0}' -f ($case[0] -join ' ')) ('{0}/{1}' -f $run.Code, $run.Log.Contains($case[1])) '1/True'
    }
    $run = Invoke-Plain @('-Path', $mainDir, '-LogonType', 'Remote')
    Assert-Equal '使える種類を並べる' $run.Log.Contains('使える種類: Interactive（対話）、RemoteDesktop（リモートデスクトップ）、Unlock（ロック解除）、Network（ネットワーク）、Batch（バッチ）、Service（サービス）、NewCredentials（別の資格情報）、All（すべて）') 'True'

    Write-Section '保存先'
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @mainArgs 2>&1
    $defaultRun = Get-RunDir (Split-Path -Parent $tool)
    Assert-Equal '-OutputFolder なしはスクリプトのフォルダ' $(if ($defaultRun) { @(Get-ChildItem -LiteralPath $defaultRun -File).Count } else { 0 }) 5
    $fileOut = Join-Path $workRoot 'out\afile'
    [IO.File]::WriteAllText($fileOut, 'x')
    $run = Invoke-Plain ($mainArgs + @('-OutputFolder', $fileOut))
    Assert-Equal '保存先がファイル: 集める前に止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains(('保存先に書き込めません: {0}' -f $fileOut)), $run.Log.Contains('-NoSave を付けると、画面に表示だけします')) '1/True/True'
    $letter = Get-FreeLetter
    if ($letter) {
        $run = Invoke-Plain ($mainArgs + @('-OutputFolder', "${letter}:\nope"))
        Assert-Equal 'ないドライブ: 止まる' ('{0}/{1}' -f $run.Code, $run.Log.Contains('保存先に書き込めません')) '1/True'
    } else {
        Write-Skip 'ないドライブ' '空いているドライブ文字がないため'
    }
    $sameDir = Join-Path $workRoot 'out\same'
    [void][IO.Directory]::CreateDirectory($sameDir)
    $stampNow = Get-Date
    foreach ($s in 0..40) {
        $older = Join-Path $sameDir ('LogonHistory_{0}_{1}' -f $env:COMPUTERNAME, $stampNow.AddSeconds($s).ToString('yyyyMMdd_HHmmss'))
        [void][IO.Directory]::CreateDirectory($older)
        [IO.File]::WriteAllText("$older\keep.txt", 'keep')
    }
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @mainArgs -OutputFolder $sameDir 2>&1
    Assert-Equal '同じ名前のフォルダがあれば _2 を付ける' (@(Get-ChildItem -LiteralPath $sameDir -Directory -Filter 'LogonHistory_*_2').Count) 1
    Assert-Equal '前の結果を残す' (@(Get-ChildItem -LiteralPath $sameDir -Recurse -Filter 'keep.txt').Count) 41
    $patched = New-PatchedTool 'savefail' "Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'events.csv'))" "throw 'disk full'; Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'events.csv'))"
    $run = Invoke-Tool 'savefail' $manyArgs $patched
    Assert-Equal '保存に失敗: 警告を出して、書きかけを残さない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした: disk full'), [bool]$run.Dir) '1/True/False'
    Assert-Equal '保存に失敗: すべて画面に出し、保存先の行を出さない' ('{0}/{1}/{2}' -f $run.Log.Contains('CORP\user05'), $run.Log.Contains('参照してください'), $run.Log.Contains('保存先:')) 'True/False/False'
    $patched = New-PatchedTool 'crash' '$sortedEvents = @(' "throw 'boom'; `$sortedEvents = @("
    $run = Invoke-Tool 'crash' $mainArgs $patched
    Assert-Equal '想定外のエラー: 結果を残さず、中断とは区別する' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('エラーで止まりました。結果は保存していません。'), $run.Log.Contains('中断しました'), [bool]$run.Dir) '1/True/False/False'

    #endregion

    #region 起動

    Write-Section 'コマンドプロンプトから起動'
    $cmdHead = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $rangeText = '-Since "2026-06-15 09:00" -Until "2026-06-15 18:00"'
    $run = Invoke-ViaCmd ('{0} -NoSave {1} -LogonType "Interactive,RemoteDesktop,Unlock" -Path "{2}\"' -f $cmdHead, $rangeText, $mainDir)
    Assert-Equal 'カンマ区切りの -LogonType と、最後の引数が \"' ('{0}/{1}' -f $run.Code, $run.Log.Contains('ログオン 10 件')) '0/True'
    # 後ろの引数に、空白を含む引用符付きの値があると、引数の組み立てが崩れて PowerShell のエラーになる（README の「止まるとき」に書いた）ので、引用符のない引数を続ける
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -NoSave -Since 200d' -f $cmdHead, $mainDir)
    Assert-Equal '-Path の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-Path の値に引用符（"）が残っています')) '1/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -OutputFolder "{2}\" -Since 200d' -f $cmdHead, $mainDir, "$workRoot\out\cmdq")
    Assert-Equal '-OutputFolder の途中で \"（止まって、保存先を作らない）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('-OutputFolder の値に引用符'), (Test-Path -LiteralPath "$workRoot\out\cmdq")) '1/True/False'
    $run = Invoke-ViaCmd ('{0} -User "CORP\" -Path "{1}" -NoSave' -f $cmdHead, $mainDir)
    Assert-Equal '-User の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-User の値に引用符（"）が残っています')) '1/True'
    $run = Invoke-ViaCmd ("{0} -NoSave -Path '{1}' -Since '2026-06-15' -Until '2026-06-16' -LogonType 'Interactive,RemoteDesktop' -User 'yamada'" -f $cmdHead, $mainDir)
    # 06-15 の 1 日分なので、期間の前・後のもの（08:30 の前の切断、18:30 のログオンと失敗）も入る
    Assert-Equal "cmd の ' はそのまま届く" ('{0}/{1}' -f $run.Code, $run.Log.Contains('ログオン 2 件（切断 2 回・再接続 2 回）  失敗 5 件')) '0/True'
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}" -Since 2000-01-01 -Until 2000-01-02' -f $cmdHead, $mainDir)
    Assert-Equal 'cmd で受け取る終了コード（該当なしは 2）' $run.Code 2
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}" -LogonType Remote' -f $cmdHead, $mainDir)
    Assert-Equal 'cmd で受け取る終了コード（止まると 1）' $run.Code 1
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}\Security.evtx" {2}' -f $cmdHead, $jpDir, $rangeText)
    Assert-Equal 'cmd からスペースと日本語を含むパス' ('{0}/{1}' -f $run.Code, $run.Log.Contains('ログオン 10 件')) '0/True'

    Write-Section '制約付き言語モードとヘルプ'
    $clmCommand = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '{0}' -Path '{1}' -OutputFolder '{2}'; exit `$LASTEXITCODE" -f $tool, $mainDir, "$workRoot\clm"
    $clm = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command $clmCommand 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $clmCode = $LASTEXITCODE
    Assert-Equal '制約付き言語モード: 止まる' ('{0}/{1}/{2}' -f $clmCode, $clm.Contains('制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）'), (Test-Path -LiteralPath "$workRoot\clm")) '1/True/False'
    $help = (@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command ("Get-Help '{0}' -Full | Out-String -Width 300" -f $tool) 2>&1) | ForEach-Object { "$_" }) -join "`n"
    Assert-Equal 'ヘルプの .NOTES' ('{0}/{1}/{2}/{3}' -f $help.Contains('Get-LogonHistory 1.0.0'), $help.Contains('Copyright (c) 2026 結進舎'), $help.Contains('Released under the MIT License.'), $help.Contains('https://github.com/kando19804-oss/workbench/blob/main/LICENSE')) 'True/True/True/True'
    Assert-Equal 'ヘルプのパラメーター' (@(@('Since', 'Until', 'User', 'LogonType', 'Path', 'OutputFolder', 'NoSave') | Where-Object { -not $help.Contains("-$_") }) -join ',') ''

    Write-Section 'Ctrl+C で中断'
    # 集め始める前に長く待つコピーで止める（読むだけなので、途中で止めても何も壊れない）
    $slowTool = New-PatchedTool 'slow' '    $sinceText = ConvertTo-QueryTime $sinceInfo.Time' "    Start-Sleep -Seconds 60`r`n    `$sinceText = ConvertTo-QueryTime `$sinceInfo.Time"
    $stopScript = New-HelperScript 'stop.ps1' @(
        'param($Tool, $LogDir, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $LogDir).AddParameter(''Since'', ''2026-06-15 09:00'').AddParameter(''Until'', ''2026-06-15 18:00'').AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'Start-Sleep -Milliseconds 3000'
        '''RUNNING='' + (-not $handle.IsCompleted)'
        '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '$ps.Stop()'
        '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
        '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
    )
    # Stop() は、動いている最中に呼んでも取りこぼして最後まで動き続けることがあるので、何度か試す
    $stopped = $false
    for ($try = 1; $try -le 5 -and -not $stopped; $try++) {
        $result = Invoke-Helper $stopScript @('-Tool', $slowTool, '-LogDir', $mainDir, '-OutRoot', (Join-Path $workRoot "out\ctrlc$try"))
        if ($result['RUNNING'] -eq 'True' -and $result['AFTER'] -eq '0') { $stopped = $true }
    }
    if ($stopped) {
        Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $result['BEFORE'], $result['AFTER']) '1/0'
        Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' 'Stop() が効かなかったため（5 回試した）'
    }

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -Since '2026-06-15 09:00' -Until '2026-06-15 18:00' -OutputFolder '{2}' | Out-Null" -f $tool, $mainDir, $cultureOut
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        $summaryText = $x.Lines -join "`n"
        Assert-Equal 'fi-FI でも期間の行は yyyy-MM-dd HH:mm' $summaryText.Contains('期間: 2026-06-15 09:00 ～ 2026-06-15 18:00') 'True'
        Assert-Equal 'fi-FI でも所要時間の小数点は .' ([regex]::IsMatch($summaryText, '所要 \d+\.\d秒')) 'True'
        Assert-Equal 'fi-FI でも CSV の日時は yyyy/MM/dd HH:mm:ss' (@($x.Sessions)[0].'ログオン') '2026/06/15 17:30:00'
        Assert-Equal 'fi-FI でも時間は H:mm:ss' (Get-Row $x.Sessions { $_.'ログオン ID' -eq '0x1000' }).'時間' '="3:30:00"'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }

    #endregion

    #region この PC のイベントログ

    Write-Section 'この PC のイベントログ'
    if (-not $isAdmin) {
        $run = Invoke-Tool 'local' @('-Since', '1d')
        Assert-Equal '管理者でない: 案内を出して止まる' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('Security ログを読む権限がありません。管理者として実行してください'), $run.Log.Contains('ほかの PC から集めた .evtx ファイルは、-Path で読めます（管理者でなくても読めます）'), [bool]$run.Dir) '1/True/True/False'
        Write-Skip 'この PC の Security ログ' '管理者でないため（管理者で実行すると確かめる）'
    } else {
        $run = Invoke-Tool 'local' @('-Since', '1d', '-LogonType', 'All')
        $x = Read-Result $run.Dir
        Assert-Equal '本物: 終了コード（0 か 2）' ($run.Code -eq 0 -or $run.Code -eq 2) 'True'
        Assert-Equal '本物: 対象の行' $run.Log.Contains(('対象: {0}（管理者: はい）' -f $env:COMPUTERNAME)) 'True'
        $labels = '対話', '対話（runas）', 'リモートデスクトップ', 'ネットワーク', 'ネットワーク（平文）', 'バッチ', 'サービス', '別の資格情報'
        Assert-Equal '本物: 種類はすべて訳してある' (@($x.Sessions | Where-Object { $labels -notcontains $_.'種類' } | ForEach-Object { $_.'種類' } | Select-Object -Unique) -join ',') ''
        Assert-Equal '本物: 状態はすべて決まった値' (@($x.Sessions | Where-Object { @('ログオフ', 'ログオン中', '切断中', '再起動', '記録なし') -notcontains $_.'状態' } | ForEach-Object { $_.'状態' } | Select-Object -Unique) -join ',') ''
        Assert-Equal '本物: 操作はすべて決まった値' (@($x.Events | Where-Object { @('ログオン', 'ログオフ', 'ロック解除', '切断', '再接続', 'ログオンの失敗', 'OS の起動', 'ログの消去') -notcontains $_.'操作' } | ForEach-Object { $_.'操作' } | Select-Object -Unique) -join ',') ''
        Assert-Equal '本物: その他のエラーはない' (@($x.Skipped | Where-Object { $_.'理由' -eq 'その他のエラー' }).Count) 0
        Assert-Equal '本物: SYSTEM とコンピューターのアカウントを出さない' (@($x.Sessions | Where-Object { $_.'ユーザー' -like '*$' -or $_.'ユーザー' -like 'NT AUTHORITY\SYSTEM' }).Count) 0
        $live = @($x.Sessions | Where-Object { $_.'状態' -eq 'ログオン中' })
        if ($live.Count -gt 0) { Assert-Equal '本物: 今のログオンは「ログオン中」' 'True' 'True' } else { Write-Skip '本物: ログオン中' '期間の中に、今も続いているログオンがないため' }
    }

    #endregion

    #region 再発防止の一覧

    Write-Section 'レビューで見つかった問題の再発防止（上で確かめている）'
    # 1. 期間の前にログオンして、期間に重なるもの（主な場面の 08:30）・期間の前に終わったものは出さない（kato）
    # 2. 再起動の後に使い回されたログオン ID（0x6000）・切断したまま再起動（0x5000）
    # 3. 短い runas（0x3000）・ロック解除のまとめ方とサインインのとき（admin）・Security の再接続とセッションのログの再接続
    # 4. 行の切断・再接続の回数は期間の中だけ（yamada の 08:45・08:50）・ログオフの後の切断（12:00:04）
    # 5. 組にできないセッションのログ（ghost・ghost2）と種類の絞り込み・21 が残っていないセッション
    # 6. 失敗の 127.0.0.1・表にない状態コード・期間の前と後の失敗
    # 7. 期間の前の OS の起動がないときの注意と、ログオンの記録がないログオフ
    # 8. 名前を変えたコンピューター・複数のコンピューター（列・注意・失敗のまとめ・ファイルごとの名前）
    # 9. 5.1 と 7 で並びが同じ（ユーザー名のハイフン・ファイル名のハイフン）
    # 10. 最後が壊れた .evtx・Security を 1 件も読めない・セッションのログだけ・読めない理由
    # 11. 画面の幅（いちばん長くなる組み合わせで 119 桁）
    # 12. -User・-LogonType・-Since・-Until・-Path・-OutputFolder の \"

    #endregion
} catch {
    '  FAIL  テストが途中で止まりました: {0}（{1} 行目）' -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber
    $script:failCount++
} finally {
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            # ジャンクションの先（linkdir 自身）を消さないよう、先にジャンクションだけを外す
            if ([IO.Directory]::Exists("$workRoot\linkdir\loop")) { [IO.Directory]::Delete("$workRoot\linkdir\loop") }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

'  END'
exit $script:failCount
