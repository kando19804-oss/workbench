<#
.SYNOPSIS
    Get-RebootHistory の回帰テストです。

.DESCRIPTION
    テスト用の System ログの .evtx を一時フォルダに組み立て、Get-RebootHistory の -Path で読んだ結果を期待値と照合します。
    この PC を調べるときの動き（今の起動、最後の起動のあとの停止）は、読むログと OS の情報の起動の時刻を差し替えたコピーで確かめます。
    本物のイベントログには書き込みません。この PC の System ログは、読むだけです（管理者でなくても読めます）。
    実際の記録はテストに入れません。記録の形（値の名前と型）は、実際の Windows の記録に合わせています。
    「権限なし」を再現するため、テストが作ったファイルとフォルダにだけアクセス拒否を設定します（終了時に戻します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Get-RebootHistory を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Get-RebootHistory\Get-RebootHistory.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Get-RebootHistory_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Get-RebootHistory.ps1'
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

# 処理を書き換えたツールのコピーを作る（本物のログでは起きない条件や、この PC を調べるときの動きを確かめるため）。
# Pairs は、書き換える前の文字と後の文字を交互に並べたもの
function New-PatchedTool([string]$Name, [string[]]$Pairs) {
    $text = [IO.File]::ReadAllText($repoTool)
    for ($i = 0; $i -lt $Pairs.Count; $i += 2) {
        if (-not $text.Contains($Pairs[$i])) { throw "書き換える場所が見つかりません: $($Pairs[$i])" }
        $text = $text.Replace($Pairs[$i], $Pairs[$i + 1])
    }
    $file = Join-Path $workRoot "patched\$Name\Get-RebootHistory.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text, $utf8Bom)
    $file
}

# この PC の System ログの代わりに File を読み、OS の情報の最後の起動の時刻を Boot にしたコピー（この PC を調べるときの動きを確かめる）
function New-LocalTool([string]$Name, [string]$File, [string]$Boot) {
    New-PatchedTool $Name @(
        '$source = @{ Label = $SystemLog; Name = $SystemLog; IsFile = $false;', ('$source = @{{ Label = $SystemLog; Name = ''{0}''; IsFile = $true;' -f $File),
        '(Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime', ('[datetime]::ParseExact(''{0}'', ''yyyy-MM-dd HH:mm:ss'', $null)' -f $Boot)
    )
}

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'RebootHistory_*' -ErrorAction SilentlyContinue | Sort-Object Name)
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
    [pscustomobject]@{ Reboots = @(Read-CsvOrEmpty 'reboots.csv'); Events = @(Read-CsvOrEmpty 'events.csv'); Skipped = @(Read-CsvOrEmpty 'skipped.csv'); Lines = $lines }
}

# reboots.csv の行を「停止|起動|種類」にして、ファイルの順に並べる
function Get-RebootKeys($Rows) { (@($Rows) | ForEach-Object { '{0}|{1}|{2}' -f $_.'停止', $_.'起動', $_.'種類' }) -join ',' }

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

# summary.txt の「■ 停止と起動」の表の行（見出しを含む）
function Get-TableLines($Lines) {
    $inTable = $false
    foreach ($line in $Lines) {
        if ($line.StartsWith('■ 停止と起動（')) { $inTable = $true; continue }
        if ($inTable) { if ($line -eq '' -or $line.StartsWith('■') -or $line.StartsWith('（ほか')) { break }; $line }
    }
}

# 表の行を、空白をまとめて 1 つにした文字にする（列の幅に左右されずに比べるため）
function Get-TableRow($Lines, [string]$Start) {
    $line = @(Get-TableLines $Lines | Where-Object { $_.StartsWith(' ' + $Start) })[0]
    if ($null -eq $line) { return '' }
    ($line -replace '\s+', ' ').Trim()
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

# System ログには、好きな内容の記録を書き込めない（書けても本物のログを汚す）ため、ファイルを直接組み立てる。
# 形式は [MS-EVEN6] の BinXml。tests\Get-LogonHistory.Tests.ps1 の組み立て方に、名前のない Data（6008 など）・
# Binary（6005・6006・6008）・Boolean（41）を足した。記録の形ごとに、チャンク（64 KB）の中で最初の記録にテンプレートを置き、
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

# 値の型（BinXml の型）。0x01 文字、0x04 UInt8、0x06 UInt16、0x08 UInt32、0x0A UInt64、0x0D Boolean、0x0E Binary、
# 0x11 FileTime、0x15 HexInt64
function Get-ValueBytes([int]$Type, $Value) {
    if ($null -eq $Value) { return , [byte[]]@() }
    switch ($Type) {
        0x01 { return , [Text.Encoding]::Unicode.GetBytes([string]$Value) }
        0x04 { return , [byte[]]@([byte]$Value) }
        0x06 { return , [BitConverter]::GetBytes([uint16]$Value) }
        0x08 { return , [BitConverter]::GetBytes([uint32]$Value) }
        0x0A { return , [BitConverter]::GetBytes([uint64]$Value) }
        0x0D { return , [BitConverter]::GetBytes([int]$(if ($Value) { 1 } else { 0 })) }
        0x0E { return , [byte[]]$Value }
        0x11 { return , [BitConverter]::GetBytes([int64]$Value) }
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
# Data の名前が '' なら名前のない Data、'#Binary' なら Binary の要素にする
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
    # if の出力は配列をほどくので、値が 1 つだけの記録（13 など）でも組の配列のままになるよう、@( ) の中で受ける
    $items = @(if ($null -ne $Event.Data) { $Event.Data } elseif ($null -ne $Event.UserData) { $Event.UserData.Items })
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
                        for ($i = 0; $i -lt $dataList.Count; $i++) {
                            $dataName = [string]$dataList[$i][0]
                            if ($dataName -eq '#Binary') { Add-XmlElement $rec $Start 'Binary' $null @((11 + $i), $dataList[$i][1]) $null }
                            elseif ($dataName -eq '') { Add-XmlElement $rec $Start 'Data' $null @((11 + $i), $dataList[$i][1]) $null }
                            else { Add-XmlElement $rec $Start 'Data' @('Name', $dataName) @((11 + $i), $dataList[$i][1]) $null }
                        }
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

# 記録を書いた順に並べて .evtx を作る（記録の番号は 1 から）。書いた時刻は、WriteAt があればそれ、なければ記録の日時。
# 停止のときの記録（13・109）は、本物と同じく次の起動のあとに書かれる（WriteAt）ので、記録の番号の順と日時の順が一致しない
function New-EvtxFile([string]$File, [object[]]$Events) {
    $chunks    = New-Object 'System.Collections.Generic.List[object]'
    $records   = New-Object 'System.Collections.Generic.List[object]'
    $offsets   = New-Object 'System.Collections.Generic.List[int]'
    $templates = @{}
    $free      = 512
    $firstId   = 1L
    $id        = 0L
    $ordered   = @(for ($i = 0; $i -lt $Events.Count; $i++) {
        $at = if ($null -ne $Events[$i].WriteAt) { [datetime]$Events[$i].WriteAt } else { [datetime]$Events[$i].Time }
        [pscustomobject]@{ Event = $Events[$i]; Time = $at; Index = $i }
    }) | Sort-Object -Property Time, Index
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

# .evtx のチャンクを壊したコピーを作る（Chunks は壊すチャンクの番号。0 が最初のチャンク。Offset はチャンクの中の壊す位置で、
# 512 ならチャンクの最初の記録から壊れる）
function New-BrokenCopy([string]$From, [string]$To, [int[]]$Chunks, [int]$Offset = 10000) {
    $bytes = [IO.File]::ReadAllBytes($From)
    foreach ($chunk in $Chunks) {
        $at = 4096 + 65536 * $chunk
        foreach ($i in ($at + $Offset)..($at + $Offset + 300)) { $bytes[$i] = 0xEE }
    }
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $To))
    [IO.File]::WriteAllBytes($To, $bytes)
}

#endregion

#region 記録を作る

$KernelGeneral   = 'Microsoft-Windows-Kernel-General'
$KernelPower     = 'Microsoft-Windows-Kernel-Power'
$KernelBoot      = 'Microsoft-Windows-Kernel-Boot'
$ClassicKeywords = [uint64]::Parse('8080000000000000', 'HexNumber')
$script:computer = 'SERVER01.corp.example.local'

# 日時（"yyyy-MM-dd HH:mm"・"yyyy-MM-dd HH:mm:ss"・"yyyy-MM-dd HH:mm:ss.fff"）
function Get-T([string]$Text) {
    $formats = [string[]]@('yyyy-MM-dd HH:mm:ss.fff', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm')
    [datetime]::ParseExact($Text, $formats, $inv, [Globalization.DateTimeStyles]::AssumeLocal)
}

function Format-T([datetime]$Time) { $Time.ToString('yyyy-MM-dd HH:mm:ss.fff', $inv) }

# System ログの記録。WriteAt は、書かれた時刻（記録の日時と違うとき）
function New-SysEvent([string]$When, [int]$Id, [string]$Provider, $Data, $UserData = $null, [string]$WriteAt = '', [int]$Version = 0) {
    $e = @{ Time = (Get-T $When); Id = $Id; Version = $Version; Level = 4; Task = 0; Keywords = $ClassicKeywords; Provider = $Provider; Channel = 'System'
            Computer = $script:computer; Data = $Data; UserData = $UserData; WriteAt = $null }
    if ($WriteAt) { $e.WriteAt = Get-T $WriteAt }
    $e
}

# 関係のない記録（停止の前の最後の記録の目安や、ログの始まりに使う）
function New-Noise([string]$When) { New-SysEvent $When 7036 'Service Control Manager' @(@('param1', 0x01, 'Windows Update'), @('param2', 0x01, 'running')) }

# SYSTEMTIME（16 バイト）
function New-SystemTimeBytes([datetime]$Time) {
    $bytes = New-Object byte[] 16
    $parts = @($Time.Year, $Time.Month, [int]$Time.DayOfWeek, $Time.Day, $Time.Hour, $Time.Minute, $Time.Second, $Time.Millisecond)
    for ($i = 0; $i -lt 8; $i++) { [BitConverter]::GetBytes([uint16]$parts[$i]).CopyTo($bytes, $i * 2) }
    , $bytes
}

# OS の起動（12）。記録の日時は起動の 1 秒後（本物と同じく、StartTime のほうが少し前）
function New-Start([string]$When, [int]$Qfe = 2227) {
    $t = Get-T $When
    New-SysEvent (Format-T $t.AddSeconds(1)) 12 $KernelGeneral @(
        @('MajorVersion', 0x08, 10), @('MinorVersion', 0x08, 0), @('BuildVersion', 0x08, 20348), @('QfeVersion', 0x08, $Qfe), @('ServiceVersion', 0x08, 0),
        @('BootMode', 0x08, 0), @('StartTime', 0x11, $t.ToUniversalTime().ToFileTimeUtc()))
}

# OS の停止（13）
function New-KernelStop([string]$When, [string]$WriteAt = '') { New-SysEvent $When 13 $KernelGeneral @(, @('StopTime', 0x11, (Get-T $When).ToUniversalTime().ToFileTimeUtc())) $null $WriteAt }

# 起動の種類（27）。0 通常、1 高速スタートアップ、2 休止状態からの再開
function New-BootType([string]$When, [int]$Type = 0) { New-SysEvent $When 27 $KernelBoot @(@('BootType', 0x08, $Type), @('LoadOptions', 0x01, ' NOEXECUTE=OPTOUT')) $null '' 1 }

# 停止の開始（109）。4 シャットダウン、5 再起動、6 電源を切る
function New-PowerStart([string]$When, [int]$Action, [string]$WriteAt = '') { New-SysEvent $When 109 $KernelPower @(@('ShutdownActionType', 0x08, $Action), @('ShutdownEventCode', 0x08, 0), @('ShutdownReason', 0x08, 5)) $null $WriteAt }

# 正常に停止しないまま起動（41 の版 10）
function New-Dirty([string]$When, [uint32]$Bugcheck = 0, [uint64]$PowerButton = 0, [uint32]$Sleep = 0, [bool]$LongPress = $false) {
    New-SysEvent $When 41 $KernelPower @(
        @('BugcheckCode', 0x08, $Bugcheck), @('BugcheckParameter1', 0x15, 0), @('BugcheckParameter2', 0x15, 0), @('BugcheckParameter3', 0x15, 0), @('BugcheckParameter4', 0x15, 0),
        @('SleepInProgress', 0x08, $Sleep), @('PowerButtonTimestamp', 0x0A, $PowerButton), @('BootAppStatus', 0x08, 0), @('Checkpoint', 0x04, 0),
        @('ConnectedStandbyInProgress', 0x0D, $false), @('SystemSleepTransitionsToOn', 0x08, 0), @('CsEntryScenarioInstanceId', 0x04, 0), @('BugcheckInfoFromEFI', 0x0D, $false),
        @('CheckpointStatus', 0x04, 0), @('CsEntryScenarioInstanceIdV2', 0x04, 0), @('LongPowerButtonPressDetected', 0x0D, $LongPress)
    ) $null '' 10
}

# 停止の指示（1074）。Process はパス（記録には「パス (コンピューター名)」の形で入る）
function New-Order([string]$When, [string]$Process, [string]$Reason, [string]$Code, [string]$Type, [string]$Comment, [string]$User) {
    $name = $script:computer.Split('.')[0]
    New-SysEvent $When 1074 'User32' @(
        @('param1', 0x01, ('{0} ({1})' -f $Process, $name)), @('param2', 0x01, $name), @('param3', 0x01, $Reason), @('param4', 0x01, $Code),
        @('param5', 0x01, $Type), @('param6', 0x01, $Comment), @('param7', 0x01, $User))
}

# 停止の取り消し（1075）
function New-Cancel([string]$When, [string]$User) { New-SysEvent $When 1075 'User32' @(@('param1', 0x01, $script:computer.Split('.')[0]), @('param2', 0x01, $User)) }

# 予期しない停止の理由の入力（1076）
function New-LaterReason([string]$When, [string]$Reason, [string]$Code, [string]$Comment, [string]$User) {
    New-SysEvent $When 1076 'User32' @(@('param1', 0x01, $Reason), @('param2', 0x01, $Code), @('param3', 0x01, ''), @('param4', 0x01, ''), @('param5', 0x01, $Comment), @('param6', 0x01, $User))
}

# イベントログの起動（6005）と停止（6006）。値は Binary だけ
function New-LogStart([string]$When) {
    $bin = New-Object byte[] 24
    (New-SystemTimeBytes (Get-T $When).ToUniversalTime()).CopyTo($bin, 0)
    New-SysEvent $When 6005 'EventLog' @(, @('#Binary', 0x0E, $bin))
}
function New-LogStop([string]$When) { New-SysEvent $When 6006 'EventLog' @(, @('#Binary', 0x0E, [byte[]](New-Object byte[] 8))) }

# 予期しない停止（6008）。名前のない文字の値 7 つ（時刻・日付は地域の書式で、見えない文字 U+200E が入る）と、
# Binary（ローカルの SYSTEMTIME、UTC の SYSTEMTIME、そのあと 32 バイト）
function New-Unexpected([string]$When, [string]$StopAt) {
    $stop = Get-T $StopAt
    $bin = New-Object byte[] 64
    (New-SystemTimeBytes $stop).CopyTo($bin, 0)
    (New-SystemTimeBytes $stop.ToUniversalTime()).CopyTo($bin, 16)
    [BitConverter]::GetBytes([int]60).CopyTo($bin, 32)
    [BitConverter]::GetBytes([int]60).CopyTo($bin, 36)
    $mark = [string][char]0x200E
    $date = $mark + $stop.ToString('yyyy', $inv) + '/' + $mark + $stop.ToString('MM', $inv) + '/' + $mark + $stop.ToString('dd', $inv)
    New-SysEvent $When 6008 'EventLog' @(@('', 0x01, $stop.ToString('HH:mm:ss', $inv)), @('', 0x01, $date), @('', 0x01, ''), @('', 0x01, ''), @('', 0x01, '600'), @('', 0x01, ''), @('', 0x01, ''), @('#Binary', 0x0E, $bin))
}

# ブルースクリーン（1001）
function New-BugCheck([string]$When, [string]$Text, [string]$Dump) {
    New-SysEvent $When 1001 'Microsoft-Windows-WER-SystemErrorReporting' @(@('param1', 0x01, $Text), @('param2', 0x01, $Dump), @('param3', 0x01, '11111111-2222-3333-4444-555555555555')) $null '' 1
}

# ログの消去（104）。値は UserData の LogFileCleared にある
function New-Clear([string]$When, [string]$Name, [string]$Channel) {
    $at = $Name.IndexOf('\')
    New-SysEvent $When 104 'Microsoft-Windows-Eventlog' $null @{ Root = 'LogFileCleared'; Ns = 'http://manifests.microsoft.com/win/2004/08/windows/eventlog'
        Items = @(@('SubjectUserName', 0x01, $Name.Substring($at + 1)), @('SubjectDomainName', 0x01, $Name.Substring(0, $at)), @('Channel', 0x01, $Channel), @('BackupPath', 0x01, '')) } '' 1
}

# 起動（12・27・6005）
function Add-Boot($List, [string]$When, [int]$Qfe = 2227) {
    $t = Get-T $When
    $List.Add((New-Start $When $Qfe))
    $List.Add((New-BootType (Format-T $t.AddSeconds(1))))
    $List.Add((New-LogStart (Format-T $t.AddSeconds(10))))
}

# 停止の指示から OS の停止まで（1074・6006・109・13）。109 と 13 は、次の起動の 5 秒後に書かれる
function Add-CleanStop($List, [string]$OrderAt, [string]$Process, [string]$Reason, [string]$Code, [string]$Type, [string]$Comment, [string]$User, [int]$Action, [string]$StopAt, [string]$NextBoot) {
    $stop = Get-T $StopAt
    $write = Format-T (Get-T $NextBoot).AddSeconds(5)
    $List.Add((New-Order $OrderAt $Process $Reason $Code $Type $Comment $User))
    $List.Add((New-LogStop (Format-T $stop.AddSeconds(-10))))
    $List.Add((New-PowerStart (Format-T $stop.AddSeconds(-5)) $Action $write))
    $List.Add((New-KernelStop $StopAt $write))
}

$ShutdownExe = 'C:\Windows\system32\shutdown.exe'
$Installer   = 'C:\Windows\servicing\TrustedInstaller.exe'
$StartMenu   = 'C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartMenuExperienceHost.exe'
$Planned     = 'その他 (計画済)'

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
    # 主な場面（SERVER01、2026-06-01 08:00 に起動）
    $main = New-Object 'System.Collections.Generic.List[object]'
    $main.Add((New-Noise '2026-06-01 07:59:00'))
    Add-Boot $main '2026-06-01 08:00:00'
    $main.Add((New-Clear '2026-06-02 09:00:00' 'CORP\admin2' 'System'))
    $main.Add((New-Clear '2026-06-02 09:05:00' 'CORP\admin2' 'Application'))              # ほかのログの消去（読まない）
    # 再起動（コメントあり）
    Add-CleanStop $main '2026-06-03 22:00:00' $ShutdownExe $Planned '0x80000000' '再起動' 'アプリの設定変更' 'CORP\yamada' 5 '2026-06-03 22:00:30' '2026-06-03 22:02:00'
    Add-Boot $main '2026-06-03 22:02:00'
    # Windows Update による再起動（OS のバージョンが変わる）
    Add-CleanStop $main '2026-06-05 03:00:00' $Installer 'オペレーティング システム: アップグレード (計画済)' '0x80020003' '再起動' '' 'NT AUTHORITY\SYSTEM' 5 '2026-06-05 03:00:30' '2026-06-05 03:05:00'
    Add-Boot $main '2026-06-05 03:05:00' 2340
    # ブルースクリーン（41・6008・遅れて書かれる 1001）
    $main.Add((New-Noise '2026-06-07 10:00:00'))
    Add-Boot $main '2026-06-07 10:01:00' 2340
    $main.Add((New-Dirty '2026-06-07 10:01:03' 0x3B))
    $main.Add((New-Unexpected '2026-06-07 10:01:10' '2026-06-07 10:00:30'))
    $main.Add((New-BugCheck '2026-06-07 10:02:30' '0x0000003b (0x00000000c0000005, 0xfffff80000001234, 0xffff800000005678, 0x0000000000000000)' 'C:\Windows\MEMORY.DMP'))
    # 電源断（41 は停止コード 0）と、管理者があとで入力した理由（1076）
    Add-Boot $main '2026-06-08 09:30:00' 2340
    $main.Add((New-Dirty '2026-06-08 09:30:05'))
    $main.Add((New-Unexpected '2026-06-08 09:30:10' '2026-06-08 09:10:00'))
    $main.Add((New-LaterReason '2026-06-08 09:40:00' 'ハードウェア: 保守 (計画外)' '0x10001' 'UPS の故障' 'CORP\admin'))
    # 電源ボタンの長押し
    Add-Boot $main '2026-06-09 12:00:00' 2340
    $main.Add((New-Dirty '2026-06-09 12:00:05' 0 133000000000000000 0 $true))
    $main.Add((New-Unexpected '2026-06-09 12:00:10' '2026-06-09 11:55:00'))
    # 停止の記録がまったくない（停止の前の最後の記録は、書かれた順ではなく時刻で決める。14:00 の記録は 16:00 に書かれた）
    $late = New-Noise '2026-06-10 14:00:00'
    $late.WriteAt = Get-T '2026-06-10 16:00:00'
    $main.Add($late)
    $main.Add((New-Noise '2026-06-10 15:00:00'))
    Add-Boot $main '2026-06-10 18:00:00' 2340
    # 英語の OS が書いた停止の指示（109 がないので、文字で判定する）
    $main.Add((New-Order '2026-06-11 23:00:00' $StartMenu 'Other (Unplanned)' '0x0' 'power off' '' 'CORP\suzuki'))
    $main.Add((New-LogStop '2026-06-11 23:00:20'))
    $main.Add((New-KernelStop '2026-06-11 23:00:30' '2026-06-12 08:00:05'))
    Add-Boot $main '2026-06-12 08:00:00' 2340
    # 取り消した指示（1075）のあとで、もう一度指示して再起動
    $main.Add((New-Order '2026-06-12 12:00:00' $ShutdownExe $Planned '0x80000000' '再起動' '' 'CORP\yamada'))
    $main.Add((New-Cancel '2026-06-12 12:05:00' 'CORP\yamada'))
    Add-CleanStop $main '2026-06-12 20:00:00' $ShutdownExe $Planned '0x80000000' '再起動' '' 'CORP\yamada' 5 '2026-06-12 20:00:30' '2026-06-12 20:02:00'
    Add-Boot $main '2026-06-12 20:02:00' 2340
    # イベントログのサービスだけの再起動（行にしない）
    $main.Add((New-LogStop '2026-06-13 10:00:00'))
    $main.Add((New-LogStart '2026-06-13 10:00:05'))
    # シャットダウンの途中で止まった（6006 のあと、41 だけ。6008 はない）
    $main.Add((New-Order '2026-06-14 23:00:00' $ShutdownExe $Planned '0x80000000' 'シャットダウン' 'メンテナンス' 'CORP\admin'))
    $main.Add((New-LogStop '2026-06-14 23:00:10'))
    Add-Boot $main '2026-06-15 07:00:00' 2340
    $main.Add((New-Dirty '2026-06-15 07:00:05'))
    # まだ再起動していない指示
    $main.Add((New-Order '2026-06-15 13:00:00' $ShutdownExe $Planned '0x80000000' '再起動' '' 'CORP\yamada'))
    $main.Add((New-Noise '2026-06-15 14:00:00'))
    $mainDir = Join-Path $workRoot 'main'
    $mainFile = Join-Path $mainDir 'System.evtx'
    New-EvtxFile $mainFile $main.ToArray()

    # 作った .evtx を、ツールと同じ読み方で、値の型まで読めるか
    $query  = New-Object Diagnostics.Eventing.Reader.EventLogQuery($mainFile, [Diagnostics.Eventing.Reader.PathType]::FilePath, "*[System[(EventID=12 or EventID=6008 or EventID=41)]]")
    $reader = New-Object Diagnostics.Eventing.Reader.EventLogReader($query)
    $types = New-Object 'System.Collections.Generic.List[string]'
    $select12 = New-Object Diagnostics.Eventing.Reader.EventLogPropertySelector (, [string[]]@("Event/EventData/Data[@Name='StartTime']", "Event/EventData/Data[@Name='QfeVersion']"))
    $select41 = New-Object Diagnostics.Eventing.Reader.EventLogPropertySelector (, [string[]]@("Event/EventData/Data[@Name='BugcheckCode']", "Event/EventData/Data[@Name='LongPowerButtonPressDetected']"))
    while ($e = $reader.ReadEvent()) {
        $names = switch ($e.Id) {
            12 { @($e.GetPropertyValues($select12) | ForEach-Object { $_.GetType().Name }) }
            6008 { @($e.Properties | Where-Object { $_.Value -is [byte[]] } | ForEach-Object { $_.Value.Length }) }
            41 { @($e.GetPropertyValues($select41) | ForEach-Object { $_.GetType().Name }) }
        }
        $types.Add(('{0}:{1}' -f $e.Id, (@($names) -join '/')))
        $e.Dispose()
    }
    $reader.Dispose()
    Assert-Equal '作った .evtx を読める（起動の時刻・6008 の Binary・41 の値の型）' ((@($types | Sort-Object -Unique)) -join ',') '12:DateTime/UInt32,41:UInt32/Boolean,6008:64'

    #endregion

    #region 主な場面

    Write-Section '主な場面'
    $run = Invoke-Tool 'main' @('-Path', $mainDir)
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '保存先のフォルダの名前' ((Split-Path -Leaf $run.Dir) -replace '\d{8}_\d{6}$', 'DATE') ('RebootHistory_{0}_DATE' -f $env:COMPUTERNAME)
    Assert-Equal '保存するファイル' (Get-ResultFileNames $run.Dir) 'events.csv,reboots.csv,skipped.csv,summary.txt'
    Assert-Equal '対象の行' $run.Log.Contains(('対象: {0}   System ログ: 2026-06-01 07:59 から残っています' -f $mainFile)) 'True'
    Assert-Equal '期間の行（指定なし）' $run.Log.Contains('期間: すべて（System ログに残っている記録）') 'True'
    Assert-Equal '-Path のときは今の起動を出さない' $run.Log.Contains('今の起動') 'False'
    Assert-Equal '合計の行' $run.Log.Contains('再起動 3 回（うち Windows Update 1 回）・シャットダウン 1 回・予期しない停止 4 回・不明 1 回   所要 ') 'True'
    Assert-Equal '停止の行（新しい順。サービスの再起動・取り消した指示は行にしない）' (Get-RebootKeys $x.Reboots) (@(
        '2026/06/14 23:00:10|2026/06/15 07:00:00|予期しない停止', '2026/06/12 20:00:30|2026/06/12 20:02:00|再起動', '2026/06/11 23:00:30|2026/06/12 08:00:00|シャットダウン',
        '|2026/06/10 18:00:00|不明', '2026/06/09 11:55:00|2026/06/09 12:00:00|予期しない停止', '2026/06/08 09:10:00|2026/06/08 09:30:00|予期しない停止',
        '2026/06/07 10:00:30|2026/06/07 10:01:00|予期しない停止', '2026/06/05 03:00:30|2026/06/05 03:05:00|再起動', '2026/06/03 22:00:30|2026/06/03 22:02:00|再起動') -join ',')
    Assert-Equal 'reboots.csv の列' ((Get-Content -LiteralPath (Join-Path $run.Dir 'reboots.csv') -Encoding UTF8 -TotalCount 1)) '"停止","起動","停止していた時間","その前の稼働時間","種類","Windows Update","きっかけ","ユーザー","理由","理由のコード","コメント","原因","停止コード","ダンプ","あとで入力した理由","停止の指示","プロセス","起動の種類","OS のバージョン","停止の前の最後の記録","備考","コンピューター"'
    Assert-Equal 'events.csv の列' ((Get-Content -LiteralPath (Join-Path $run.Dir 'events.csv') -Encoding UTF8 -TotalCount 1)) '"日時","操作","詳細","ユーザー","プロセス","イベント ID","ソース","コンピューター"'
    Assert-Equal 'skipped.csv の列' ((Get-Content -LiteralPath (Join-Path $run.Dir 'skipped.csv') -Encoding UTF8 -TotalCount 1)) '"理由","対象","詳細"'

    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/03 22:02:00' }
    Assert-Equal '再起動: 時間・きっかけ・ユーザー' ('{0}|{1}|{2}|{3}|{4}' -f $row.'停止していた時間', $row.'その前の稼働時間', $row.'きっかけ', $row.'ユーザー', $row.'Windows Update') '="0:01:30"|2日 14:00:30|shutdown コマンド|CORP\yamada|'
    Assert-Equal '再起動: 理由・コメント・指示・プロセス（コンピューター名を外す）' ('{0}|{1}|{2}|{3}|{4}' -f $row.'理由', $row.'理由のコード', $row.'コメント', $row.'停止の指示', $row.'プロセス') ('{0}|0x80000000|アプリの設定変更|2026/06/03 22:00:00|{1}' -f $Planned, $ShutdownExe)
    Assert-Equal '再起動: 起動の種類・OS のバージョン・コンピューター' ('{0}|{1}|{2}|[{3}]' -f $row.'起動の種類', $row.'OS のバージョン', $row.'コンピューター', $row.'備考') '通常|="10.0.20348.2227"|SERVER01.corp.example.local|[]'
    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/05 03:05:00' }
    Assert-Equal 'Windows Update: 印・きっかけ・OS のバージョンの変化' ('{0}|{1}|{2}|{3}' -f $row.'Windows Update', $row.'きっかけ', $row.'ユーザー', $row.'OS のバージョン') 'はい|Windows Update|NT AUTHORITY\SYSTEM|10.0.20348.2227 → 10.0.20348.2340'
    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/07 10:01:00' }
    Assert-Equal 'ブルースクリーン: 原因・停止コード・ダンプ' ('{0}|{1}|{2}' -f $row.'原因', $row.'停止コード', $row.'ダンプ') 'ブルースクリーン 0x0000003B（SYSTEM_SERVICE_EXCEPTION）|0x0000003B|C:\Windows\MEMORY.DMP'
    Assert-Equal 'ブルースクリーン: 停止の時刻は 6008 のバイナリ（UTC）から・停止時間' ('{0}|{1}' -f $row.'停止', $row.'停止していた時間') '2026/06/07 10:00:30|="0:00:30"'
    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/08 09:30:00' }
    Assert-Equal '電源断: 原因とあとで入力した理由' ('{0}|{1}' -f $row.'原因', $row.'あとで入力した理由') '不明（電源断・フリーズ・強制的な電源オフなど）|ハードウェア: 保守 (計画外)「UPS の故障」（CORP\admin）'
    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/09 12:00:00' }
    Assert-Equal '電源ボタンの長押し' $row.'原因' '電源ボタンの長押し'
    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/10 18:00:00' }
    Assert-Equal '停止の記録がない: 停止の前の最後の記録と備考' ('[{0}]|{1}|{2}|[{3}]' -f $row.'停止', $row.'停止の前の最後の記録', $row.'備考', $row.'停止していた時間') '[]|2026/06/10 15:00:00|停止の記録がありません|[]'
    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/12 08:00:00' }
    Assert-Equal '英語の停止の指示（power off）はシャットダウン・スタート メニュー' ('{0}|{1}|{2}' -f $row.'種類', $row.'きっかけ', $row.'理由') 'シャットダウン|スタート メニュー|Other (Unplanned)'
    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/12 20:02:00' }
    Assert-Equal '取り消した指示ではなく、取り消しのない最後の指示を使う' $row.'停止の指示' '2026/06/12 20:00:00'
    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/15 07:00:00' }
    Assert-Equal 'シャットダウンの途中で止まった: 停止は 6006、ユーザーと備考' ('{0}|{1}|{2}|{3}' -f $row.'停止', $row.'ユーザー', $row.'コメント', $row.'備考') '2026/06/14 23:00:10|CORP\admin|メンテナンス|シャットダウンの途中で止まりました'

    # 注意
    Assert-Equal 'System ログの消去の注意（ほかのログの消去は出さない）' ('{0}/{1}' -f $run.Log.Contains('※ 2026-06-02 09:00 に System ログが消去されています（CORP\admin2）'), $run.Log.Contains('2026-06-02 09:05')) 'True/False'
    Assert-Equal 'まだ再起動していない指示の注意' $run.Log.Contains('※ 2026-06-15 13:00 に再起動の指示（CORP\yamada）がありますが、ログを書き出した時点で、まだ再起動していません') 'True'
    Assert-Equal '取り消した指示は注意に出さない' $run.Log.Contains('2026-06-12 12:00 に') 'False'

    # events.csv
    $ops = @($x.Events | ForEach-Object { $_.'操作' } | Sort-Object -Unique) -join ','
    Assert-Equal 'events.csv の操作（13 種類すべて）' $ops ((@('イベントログの起動', 'イベントログの停止', 'ブルースクリーン', 'ログの消去', '起動の種類', '停止の開始', '停止の指示', '停止の取り消し', '正常に停止しないまま起動', '予期しない停止', '理由の入力', 'OS の起動', 'OS の停止') | Sort-Object) -join ',')
    Assert-Equal 'events.csv の件数（ほかのログの消去と関係のない記録は除く）' @($x.Events).Count ($main.Count - 6)
    Assert-Equal 'events.csv は新しい順' (@($x.Events)[0].'日時') '2026/06/15 13:00:00'
    Assert-Equal 'サービスの再起動の 6006 と 6005' ((@($x.Events | Where-Object { $_.'日時' -like '2026/06/13 10:00:0*' }) | ForEach-Object { '{0}/{1}' -f $_.'操作', $_.'詳細' }) -join ',') 'イベントログの起動/イベントログのサービスの再起動（OS は再起動していません）,イベントログの停止/イベントログのサービスの再起動（OS は再起動していません）'
    $evt = Get-Row $x.Events { $_.'日時' -eq '2026/06/11 23:00:00' }
    Assert-Equal '停止の指示の詳細は記録の文字のまま' ('{0}|{1}|{2}' -f $evt.'詳細', $evt.'ユーザー', $evt.'イベント ID') 'power off: Other (Unplanned)（0x0）|CORP\suzuki|1074'
    $evt = Get-Row $x.Events { $_.'操作' -eq '停止の取り消し' }
    Assert-Equal '取り消しの行' ('{0}|{1}' -f $evt.'日時', $evt.'ユーザー') '2026/06/12 12:05:00|CORP\yamada'
    $evt = Get-Row $x.Events { $_.'操作' -eq '予期しない停止' -and $_.'日時' -like '2026/06/07*' }
    Assert-Equal '6008 の詳細（停止の時刻）' $evt.'詳細' '停止の時刻: 2026/06/07 10:00:30'
    $evt = Get-Row $x.Events { $_.'操作' -eq 'OS の起動' -and $_.'日時' -like '2026/06/05*' }
    Assert-Equal 'OS の起動の時刻は記録の値（記録の日時より 1 秒前）と詳細' ('{0}|{1}' -f $evt.'日時', $evt.'詳細') '2026/06/05 03:05:00|OS のバージョン: 10.0.20348.2340'

    Write-Section '画面と summary.txt'
    Assert-Equal '画面: 再起動の行' (Get-TableRow $x.Lines '2026-06-03 22:00') '2026-06-03 22:00 22:02 0:01:30 再起動 CORP\yamada shutdown コマンド: アプリの設定変更'
    Assert-Equal '画面: 違う日の起動は月日と時刻・35 桁で切る' (Get-TableRow $x.Lines '2026-06-11 23:00') '2026-06-11 23:00 06-12 08:00 8:59:30 シャットダウン CORP\suzuki スタート メニュー: Other (Unplann…'
    Assert-Equal '画面: 停止の記録がない' (Get-TableRow $x.Lines '（不明）') '（不明） 06-10 18:00 不明 停止の記録がありません'
    Assert-Equal '画面: 入力された理由を先に書く' (Get-TableRow $x.Lines '2026-06-08 09:10') '2026-06-08 09:10 09:30 0:20:00 予期しない停止 入力された理由: UPS の故障 ／ 不…'
    Assert-Equal '画面: 停止の途中で止まった' (Get-TableRow $x.Lines '2026-06-14 23:00') '2026-06-14 23:00 06-15 07:00 7:59:50 予期しない停止 CORP\admin シャットダウンの途中で停止（shutd…'
    Assert-Equal '画面: Windows Update' (Get-TableRow $x.Lines '2026-06-05 03:00') '2026-06-05 03:00 03:05 0:04:30 再起動 NT AUTHORITY\SYSTEM Windows Update: オペレーティング …'
    $table = @(Get-TableLines $x.Lines)
    Assert-Equal '表の見出し' (($table[0] -replace '\s+', ' ').Trim()) '停止 起動 停止時間 種類 ユーザー きっかけ・理由'
    Assert-Equal '表は 120 桁に収まる' ((@($table | ForEach-Object { Get-DisplayWidth $_ }) | Measure-Object -Maximum).Maximum -le 120) 'True'
    Assert-Equal 'summary.txt の最後は保存先' ($x.Lines[-1]) ('保存先: {0}' -f $run.Dir)
    Assert-Equal 'summary.txt と画面の見出しが同じ' $run.Log.Contains(($x.Lines[0..2] -join '')) 'True'

    #endregion

    #region 期間

    Write-Section '期間'
    $run = Invoke-Plain @('-Path', $mainDir, '-Since', '2026-06-08 00:00', '-Until', '2026-06-10 23:59', '-NoSave')
    Assert-Equal '期間の中の停止だけ' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('予期しない停止 2 回・不明 1 回'), $run.Log.Contains('2026-06-07')) '0/True/False'
    Assert-Equal '期間の行' $run.Log.Contains('期間: 2026-06-08 00:00 ～ 2026-06-10 23:59') 'True'
    $run = Invoke-Tool 'since' @('-Path', $mainDir, '-Since', '2026-06-12 00:00', '-Until', '2026-06-12 23:59')
    $x = Read-Result $run.Dir
    Assert-Equal '停止が期間の前でも、起動が期間に入れば出す' (Get-RebootKeys $x.Reboots) '2026/06/12 20:00:30|2026/06/12 20:02:00|再起動,2026/06/11 23:00:30|2026/06/12 08:00:00|シャットダウン'
    Assert-Equal 'events.csv は期間の中の記録だけ' (@($x.Events | Where-Object { $_.'日時' -notlike '2026/06/12*' }).Count) 0
    $run = Invoke-Plain @('-Path', $mainDir, '-Since', '2026-05-01', '-Until', '2026-06-02', '-NoSave')
    Assert-Equal '期間の始めより前の記録が残っていない注意' $run.Log.Contains('※ 2026-06-01 07:59 より前の記録は、System ログに残っていません（上書きされています）') 'True'
    $run = Invoke-Plain @('-Path', $mainDir, '-Until', '2026-06-04', '-NoSave')
    Assert-Equal '-Until だけ' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('期間: ～ 2026-06-04 00:00'), $run.Log.Contains('再起動 1 回   所要')) '0/True/True'
    $run = Invoke-Plain @('-Path', $mainDir, '-Since', '2026-07-01', '-Until', '2026-07-02', '-NoSave')
    Assert-Equal '該当なし: 終了コード 2' ('{0}/{1}' -f $run.Code, $run.Log.Contains('該当する停止はありません')) '2/True'
    $run = Invoke-Plain @('-Path', $mainDir, '-Since', '3000d', '-NoSave')
    Assert-Equal '今からの時間（-Since 3000d）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('（3000 日）')) '0/True'

    #endregion

    #region 短い間隔の起動（1 回目のレビュー）

    Write-Section '短い間隔の起動・12 が読めない起動・サービスの再起動'
    $script:computer = 'QUICK01'
    $quick = New-Object 'System.Collections.Generic.List[object]'
    Add-Boot $quick '2026-06-20 09:00:00'
    # 40 秒おきのブルースクリーンの繰り返し
    Add-Boot $quick '2026-06-20 10:00:00'
    $quick.Add((New-Dirty '2026-06-20 10:00:03' 0x50))
    $quick.Add((New-Unexpected '2026-06-20 10:00:10' '2026-06-20 09:59:50'))
    Add-Boot $quick '2026-06-20 10:00:40'
    $quick.Add((New-Dirty '2026-06-20 10:00:43' 0x3B))
    $quick.Add((New-Unexpected '2026-06-20 10:00:50' '2026-06-20 10:00:35'))
    # 停止と起動の時刻に秒の端数がある（本物と同じ。秒を切り捨ててから引く）
    Add-CleanStop $quick '2026-06-20 10:01:00' $ShutdownExe $Planned '0x80000000' '再起動' '' 'CORP\admin' 5 '2026-06-20 10:01:08.990' '2026-06-20 10:01:20.500'
    Add-Boot $quick '2026-06-20 10:01:20.500'
    # 予期しない停止のあと、理由を入力してすぐ再起動
    Add-Boot $quick '2026-06-21 10:04:00'
    $quick.Add((New-Unexpected '2026-06-21 10:04:10' '2026-06-21 10:03:00'))
    $quick.Add((New-LaterReason '2026-06-21 10:05:00' 'ハードウェア: 保守 (計画外)' '0x10001' 'UPS の故障' 'CORP\admin'))
    Add-CleanStop $quick '2026-06-21 10:05:20' $ShutdownExe $Planned '0x80000000' '再起動' '' 'CORP\admin' 5 '2026-06-21 10:05:35' '2026-06-21 10:05:50'
    Add-Boot $quick '2026-06-21 10:05:50'
    # 1001 が起動から 90 秒後に書かれ、すぐ再起動
    Add-Boot $quick '2026-06-22 08:00:00'
    $quick.Add((New-Dirty '2026-06-22 08:00:05' 0x3B))
    $quick.Add((New-Unexpected '2026-06-22 08:00:10' '2026-06-22 07:59:00'))
    $quick.Add((New-BugCheck '2026-06-22 08:01:30' '0x0000003b (0x0, 0x0, 0x0, 0x0)' 'C:\Windows\MEMORY.DMP'))
    $quick.Add((New-Order '2026-06-22 08:01:40' 'C:\Program Files\App\app.exe' 'レガシ API シャットダウン' '0x80070000' '再起動' '' 'CORP\other'))   # 停止の前の指示が 2 つ（最後のものを使う）
    Add-CleanStop $quick '2026-06-22 08:01:50' $ShutdownExe $Planned '0x80000000' '再起動' '' 'CORP\admin' 5 '2026-06-22 08:02:05' '2026-06-22 08:02:20'
    Add-Boot $quick '2026-06-22 08:02:20'
    # 12 だけが読めず、27 と 6005 が残る起動（27 の時刻を起動にする）。指示の種類は日本語でも英語でもない（109 の値で再起動と分かる）
    Add-CleanStop $quick '2026-06-23 01:00:00' $ShutdownExe $Planned '0x80000000' 'Neustart' '' 'CORP\admin' 5 '2026-06-23 01:00:20' '2026-06-23 01:00:40'
    $quick.Add((New-BootType '2026-06-23 01:00:40'))
    $quick.Add((New-LogStart '2026-06-23 01:00:50'))
    # 12 と 27 が読めず、6005 だけが残る起動（サービスの再起動と取り違えない）
    Add-CleanStop $quick '2026-06-24 01:00:00' $ShutdownExe $Planned '0x80000000' 'シャットダウン' '' 'CORP\admin' 6 '2026-06-24 01:00:20' '2026-06-24 09:00:10'
    $quick.Add((New-LogStart '2026-06-24 09:00:10'))
    # サービスが落ちて、6006 なしで 6005（前の起動のあとに停止の記録がない）
    Add-Boot $quick '2026-06-25 08:00:00'
    $quick.Add((New-LogStart '2026-06-25 12:00:00'))
    # 最後の起動のあとに、取り消した指示（注意に出さない）
    $quick.Add((New-Order '2026-06-25 13:00:00' $ShutdownExe $Planned '0x80000000' '再起動' '' 'CORP\admin'))
    $quick.Add((New-Cancel '2026-06-25 13:05:00' 'CORP\admin'))
    $quickDir = Join-Path $workRoot 'quick'
    New-EvtxFile "$quickDir\System.evtx" $quick.ToArray()
    $run = Invoke-Tool 'quick' @('-Path', $quickDir)
    $x = Read-Result $run.Dir
    # 06-24 の起動と 06-25 08:00 の起動の間には停止の記録がないので、06-25 08:00 の前は「不明」
    Assert-Equal '短い間隔でも、それぞれの停止に付く' (Get-RebootKeys $x.Reboots) (@(
        '|2026/06/25 08:00:00|不明', '2026/06/24 01:00:20|2026/06/24 09:00:10|シャットダウン', '2026/06/23 01:00:20|2026/06/23 01:00:40|再起動',
        '2026/06/22 08:02:05|2026/06/22 08:02:20|再起動', '2026/06/22 07:59:00|2026/06/22 08:00:00|予期しない停止', '2026/06/21 10:05:35|2026/06/21 10:05:50|再起動',
        '2026/06/21 10:03:00|2026/06/21 10:04:00|予期しない停止', '2026/06/20 10:01:08|2026/06/20 10:01:20|再起動', '2026/06/20 10:00:35|2026/06/20 10:00:40|予期しない停止',
        '2026/06/20 09:59:50|2026/06/20 10:00:00|予期しない停止') -join ',')
    Assert-Equal '停止コードはそれぞれの 41 から' ((@($x.Reboots | Where-Object { $_.'起動' -like '2026/06/20 10:00:*' }) | ForEach-Object { $_.'停止コード' }) -join ',') '0x0000003B,0x00000050'
    Assert-Equal '入力した理由は予期しない停止に付き、すぐあとの再起動には付かない' ((@($x.Reboots | Where-Object { $_.'起動' -like '2026/06/21*' }) | ForEach-Object { '{0}/[{1}]' -f $_.'種類', $_.'あとで入力した理由' }) -join ',') '再起動/[],予期しない停止/[ハードウェア: 保守 (計画外)「UPS の故障」（CORP\admin）]'
    Assert-Equal '遅れて書かれた 1001 は元の停止に付く' ((@($x.Reboots | Where-Object { $_.'起動' -like '2026/06/22*' }) | ForEach-Object { '{0}/[{1}]' -f $_.'種類', $_.'ダンプ' }) -join ',') '再起動/[],予期しない停止/[C:\Windows\MEMORY.DMP]'
    $row = Get-Row $x.Reboots { $_.'起動' -eq '2026/06/24 09:00:10' }
    Assert-Equal '6005 だけの起動の前の停止の記録は、サービスの再起動に数えない' ('{0}|[{1}]' -f $row.'種類', $row.'備考') 'シャットダウン|[]'
    Assert-Equal 'サービスが落ちた 6005（6006 なし）は行にしない' (@($x.Reboots | Where-Object { $_.'起動' -like '2026/06/25 12*' }).Count) 0
    Assert-Equal 'サービスが落ちた 6005 の詳細' (Get-Row $x.Events { $_.'日時' -eq '2026/06/25 12:00:00' }).'詳細' 'イベントログのサービスの再起動（OS は再起動していません）'
    Assert-Equal '本物の停止の 6006 には、サービスの再起動と書かない' (@($x.Events | Where-Object { $_.'詳細' -like '*サービスの再起動*' }).Count) 1
    Assert-Equal '6005 で補った起動に 27 があれば、27 の時刻を起動にする' (Get-Row $x.Reboots { $_.'停止' -eq '2026/06/23 01:00:20' }).'起動の種類' '通常'
    Assert-Equal '秒の端数を切り捨ててから引く（10:01:08.990 → 10:01:20.500）' (Get-Row $x.Reboots { $_.'停止' -eq '2026/06/20 10:01:08' }).'停止していた時間' '="0:00:12"'
    Assert-Equal '停止の前の指示が 2 つなら、最後のものを使う' (Get-Row $x.Reboots { $_.'停止' -eq '2026/06/22 08:02:05' }).'ユーザー' 'CORP\admin'
    Assert-Equal '取り消した指示は、まだ再起動していない注意に出さない' $run.Log.Contains('まだ') 'False'

    #endregion

    #region 12 のないコンピューター・年をまたぐ停止

    Write-Section '12 の記録がないコンピューター'
    $script:computer = 'OLD01'
    $old = New-Object 'System.Collections.Generic.List[object]'
    $old.Add((New-LogStart '2026-06-25 08:00:10'))
    $old.Add((New-Order '2026-06-25 12:00:00' $ShutdownExe $Planned '0x80000000' '再起動' '' 'CORP\admin'))
    $old.Add((New-LogStop '2026-06-25 12:00:10'))
    $old.Add((New-LogStart '2026-06-25 12:01:00'))
    $old.Add((New-Dirty '2026-06-25 15:00:04' 0 133000000000000000))
    $old.Add((New-LogStart '2026-06-25 15:00:12'))
    $old.Add((New-Unexpected '2026-06-25 15:00:13' '2026-06-25 14:50:00'))
    $oldDir = Join-Path $workRoot 'old'
    New-EvtxFile "$oldDir\System.evtx" $old.ToArray()
    $run = Invoke-Tool 'old' @('-Path', $oldDir)
    $x = Read-Result $run.Dir
    Assert-Equal '6005 で起動を補う（41 が 6005 より先でも、その起動に付く）' (Get-RebootKeys $x.Reboots) '2026/06/25 14:50:00|2026/06/25 15:00:12|予期しない停止,2026/06/25 12:00:10|2026/06/25 12:01:00|再起動'
    Assert-Equal '予期しない停止の原因' (@($x.Reboots)[0].'原因') '電源ボタンの長押し'

    Write-Section 'Windows Server のきっかけ（評価版の期限切れ・WMI）'
    $script:computer = 'SERVER02'
    $server = New-Object 'System.Collections.Generic.List[object]'
    Add-Boot $server '2026-07-25 11:00:00'
    Add-CleanStop $server '2026-07-25 16:14:00' 'C:\WINDOWS\system32\wbem\wmiprvse.exe' 'この理由のタイトルが見つかりません' '0x80070015' '再起動' '' 'SERVER02\Administrator' 5 '2026-07-25 16:14:20' '2026-07-25 16:14:30'
    Add-Boot $server '2026-07-25 16:14:30'
    Add-CleanStop $server '2026-08-04 22:10:00' 'C:\WINDOWS\system32\wlms\wlms.exe' $Planned '0x80000000' 'シャットダウン' 'この Windows インストールのライセンスの有効期限が切れています。オペレーティング システムはシャットダウンします。' 'NT AUTHORITY\SYSTEM' 6 '2026-08-04 22:10:30' '2026-11-11 21:42:00'
    Add-Boot $server '2026-11-11 21:42:00'
    $serverDir = Join-Path $workRoot 'server'
    New-EvtxFile "$serverDir\System.evtx" $server.ToArray()
    $run = Invoke-Tool 'server' @('-Path', $serverDir)
    $x = Read-Result $run.Dir
    Assert-Equal 'きっかけの名前（評価版の期限切れ・WMI）' ((@($x.Reboots) | ForEach-Object { '{0}/{1}' -f $_.'きっかけ', $_.'プロセス' }) -join ',') '評価版の期限切れ/C:\WINDOWS\system32\wlms\wlms.exe,WMI（Restart-Computer など）/C:\WINDOWS\system32\wbem\wmiprvse.exe'
    Assert-Equal '画面: 評価版の期限切れはコメントを出す' (Get-TableRow $x.Lines '2026-08-04 22:10') '2026-08-04 22:10 11-11 21:42 98日 23:31:30 シャットダウン NT AUTHORITY\SYSTEM 評価版の期限切れ: この Windows イ…'

    Write-Section '年をまたぐ停止と画面の幅'
    $script:computer = 'YEAR01'
    $year = New-Object 'System.Collections.Generic.List[object]'
    Add-Boot $year '2025-12-01 08:00:00'
    Add-CleanStop $year '2025-12-20 18:00:00' 'C:\Windows\System32\svchost.exe' 'オペレーティング システム: サービス パック (計画済)' '0x80020010' '電源を切る' '' 'NT AUTHORITY\SYSTEM' 6 '2025-12-20 18:00:30' '2026-01-05 08:30:00'
    Add-Boot $year '2026-01-05 08:30:00'
    $yearDir = Join-Path $workRoot 'year'
    New-EvtxFile "$yearDir\System.evtx" $year.ToArray()
    $run = Invoke-Tool 'year' @('-Path', $yearDir)
    $x = Read-Result $run.Dir
    Assert-Equal '年をまたぐ起動も月日と時刻・理由のコードで Windows Update' (Get-TableRow $x.Lines '2025-12-20 18:00') '2025-12-20 18:00 01-05 08:30 15日 14:29:30 シャットダウン NT AUTHORITY\SYSTEM svchost.exe: オペレーティング シ…'
    Assert-Equal 'それでも 120 桁に収まる' ((@(Get-TableLines $x.Lines | ForEach-Object { Get-DisplayWidth $_ }) | Measure-Object -Maximum).Maximum -le 120) 'True'
    Assert-Equal '合計の行（シャットダウンの Windows Update）' $run.Log.Contains('シャットダウン 1 回（うち Windows Update 1 回）') 'True'

    #endregion

    #region 一覧の行数

    Write-Section '一覧の行数'
    $script:computer = 'MANY01'
    $many = New-Object 'System.Collections.Generic.List[object]'
    Add-Boot $many '2026-06-01 00:00:00'
    foreach ($i in 1..25) {
        $day = (Get-T '2026-06-01 00:00:00').AddDays($i)
        Add-CleanStop $many (Format-T $day.AddMinutes(-2)) $ShutdownExe $Planned '0x80000000' '再起動' ('作業{0:00}' -f $i) 'CORP\yamada' 5 (Format-T $day.AddMinutes(-1)) (Format-T $day)
        Add-Boot $many (Format-T $day)
    }
    $manyDir = Join-Path $workRoot 'many'
    New-EvtxFile "$manyDir\System.evtx" $many.ToArray()
    $run = Invoke-Tool 'many' @('-Path', $manyDir)
    $x = Read-Result $run.Dir
    Assert-Equal '保存したときは、画面は 20 件まで' ('{0}/{1}/{2}' -f $run.Log.Contains('（ほか 5 件は reboots.csv を参照してください）'), $run.Log.Contains('作業25'), $run.Log.Contains('作業05')) 'True/True/False'
    Assert-Equal 'summary.txt と reboots.csv は省略しない' ('{0}/{1}/{2}' -f @($x.Lines | Where-Object { $_ -like '*作業05*' }).Count, @($x.Lines | Where-Object { $_ -like '*参照してください*' }).Count, @($x.Reboots).Count) '1/0/25'
    $run = Invoke-Plain @('-Path', $manyDir, '-NoSave')
    Assert-Equal '-NoSave のときは、すべて表示する' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('参照してください'), $run.Log.Contains('作業01')) '0/False/True'

    #endregion

    #region 複数のコンピューター・重複

    Write-Section '複数のコンピューターの .evtx'
    $multiDir = Join-Path $workRoot 'multi'
    [void][IO.Directory]::CreateDirectory("$multiDir\a"); [void][IO.Directory]::CreateDirectory("$multiDir\b")
    [IO.File]::Copy($mainFile, "$multiDir\a\System.evtx")
    [IO.File]::Copy("$oldDir\System.evtx", "$multiDir\b\System.evtx")
    $run = Invoke-Tool 'multi' @('-Path', $multiDir)
    $x = Read-Result $run.Dir
    Assert-Equal '対象はほか 1 ファイル・System ログの行なし（コンピューターが複数）' ('{0}/{1}' -f $run.Log.Contains(('対象: {0}\a\System.evtx ほか 1 ファイル' -f $multiDir)), $run.Log.Contains('から残っています')) 'True/False'
    Assert-Equal '表の最後にコンピューターの列' ((@(Get-TableLines $x.Lines)[0] -replace '\s+', ' ').Trim()) '停止 起動 停止時間 種類 ユーザー きっかけ・理由 コンピューター'
    Assert-Equal 'コンピューターごとに組み合わせる' (@($x.Reboots).Count) 11
    Assert-Equal '注意の先頭にコンピューターの名前' $run.Log.Contains('※ SERVER01: 2026-06-02 09:00 に System ログが消去されています（CORP\admin2）') 'True'
    Assert-Equal '行のコンピューター' ((@($x.Reboots | ForEach-Object { $_.'コンピューター' } | Sort-Object -Unique)) -join ',') 'OLD01,SERVER01.corp.example.local'

    Write-Section '同じファイルを重ねて指定'
    $run = Invoke-Plain @('-Path', ('{0},{1},{2}' -f $mainFile, $mainDir, $mainFile.ToUpperInvariant()), '-NoSave')
    Assert-Equal '1 回だけ読む（ほか N ファイルを出さない）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('ほか'), $run.Log.Contains('再起動 3 回（うち Windows Update 1 回）・シャットダウン 1 回・予期しない停止 4 回・不明 1 回')) '0/False/True'

    #endregion

    #region この PC を調べるときの動き（読むログと OS の情報の起動の時刻を差し替えたコピー）

    Write-Section 'この PC を調べるとき（差し替えたコピー）'
    $script:computer = 'LOCALPC'
    # 最後の起動の記録のあとに停止の記録がある（今の起動の記録が残っていない）
    $local1 = New-Object 'System.Collections.Generic.List[object]'
    $local1.Add((New-Noise '2026-06-15 06:59:00'))
    Add-Boot $local1 '2026-06-15 07:00:00'
    $local1.Add((New-Order '2026-06-15 20:00:00' $StartMenu $Planned '0x0' 'シャットダウン' '' 'LOCALPC\user'))
    $local1.Add((New-LogStop '2026-06-15 20:00:10'))
    $local1.Add((New-KernelStop '2026-06-15 20:00:20'))
    $local1File = Join-Path $workRoot 'local1\System.evtx'
    New-EvtxFile $local1File $local1.ToArray()
    $localTool = New-LocalTool 'local1' $local1File '2026-06-16 08:00:00'
    $run = Invoke-Plain @('-NoSave') $localTool
    Assert-Equal '今の起動（OS の情報）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains(('対象: {0}   System ログ: 2026-06-15 06:59 から残っています' -f $env:COMPUTERNAME)), $run.Log.Contains('今の起動: 2026-06-16 08:00（稼働 ')) '0/True/True'
    $run = Invoke-Tool 'local1' @() $localTool
    $x = Read-Result $run.Dir
    Assert-Equal '起動の記録がない停止は、OS の情報の起動を使う（行と備考）' ('{0}|{1}|{2}' -f (Get-RebootKeys $x.Reboots), @($x.Reboots)[0].'備考', @($x.Reboots)[0].'停止していた時間') '2026/06/15 20:00:20|2026/06/16 08:00:00|シャットダウン|起動の記録が残っていないため、起動の時刻は OS の情報（最後の起動の時刻）から取りました|="11:59:40"'
    # OS の情報の起動が停止より前（休止状態など）なら、行にしない
    $localTool = New-LocalTool 'local1b' $local1File '2026-06-15 19:00:00'
    $run = Invoke-Plain @('-NoSave') $localTool
    Assert-Equal '起動が停止より前なら行にしない（今の起動は新しいほう）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('該当する停止はありません'), $run.Log.Contains('今の起動: 2026-06-15 19:00')) '2/True/True'
    # 今の起動の記録が残っていない（長く再起動していない）
    $local2 = @((New-Noise '2026-06-01 00:00:00'), (New-Noise '2026-06-15 00:00:00'))
    $local2File = Join-Path $workRoot 'local2\System.evtx'
    New-EvtxFile $local2File $local2
    $run = Invoke-Plain @('-NoSave') (New-LocalTool 'local2' $local2File '2026-05-01 03:00:00')
    Assert-Equal '今の起動の記録が残っていない注意' ('{0}/{1}' -f $run.Code, $run.Log.Contains('※ 今の起動（2026-05-01 03:00）の記録は、System ログに残っていません。それより前の停止は分かりません')) '2/True'
    # 高速スタートアップの起動（OS の情報の起動の時刻は変わらない。OS の停止とイベントログの停止は記録されない）
    $local3 = New-Object 'System.Collections.Generic.List[object]'
    Add-Boot $local3 '2026-06-10 08:00:00'
    $local3.Add((New-Order '2026-06-15 22:00:00' $StartMenu $Planned '0x0' '電源を切る' '' 'LOCALPC\user'))
    $local3.Add((New-PowerStart '2026-06-15 22:00:05' 6))
    $local3.Add((New-BootType '2026-06-16 09:00:00' 1))
    $local3File = Join-Path $workRoot 'local3\System.evtx'
    New-EvtxFile $local3File $local3.ToArray()
    $run = Invoke-Tool 'local3' @() (New-LocalTool 'local3' $local3File '2026-06-10 08:00:00')
    $x = Read-Result $run.Dir
    Assert-Equal '高速スタートアップ: 今の起動は 27 の時刻' $run.Log.Contains('今の起動: 2026-06-16 09:00') 'True'
    Assert-Equal '高速スタートアップ: 起動の種類と、停止の完了の記録の備考を付けない' ('{0}|{1}|[{2}]' -f (Get-RebootKeys $x.Reboots), @($x.Reboots)[0].'起動の種類', @($x.Reboots)[0].'備考') '2026/06/15 22:00:05|2026/06/16 09:00:00|シャットダウン|高速スタートアップ|[]'
    # 休止状態からの再開（27 の 2）は、起動に数えない
    $local4 = New-Object 'System.Collections.Generic.List[object]'
    Add-Boot $local4 '2026-06-15 07:00:00'
    $local4.Add((New-BootType '2026-06-15 12:00:00' 2))
    # イベントログのサービスだけの再起動と、まだ再起動していない指示（この PC のときの言い方）
    $local4.Add((New-LogStop '2026-06-15 13:00:00'))
    $local4.Add((New-LogStart '2026-06-15 13:00:05'))
    $local4.Add((New-Order '2026-06-15 14:00:00' $ShutdownExe $Planned '0x80000000' '再起動' '' 'LOCALPC\user'))
    $local4File = Join-Path $workRoot 'local4\System.evtx'
    New-EvtxFile $local4File $local4.ToArray()
    $run = Invoke-Plain @('-NoSave') (New-LocalTool 'local4' $local4File '2026-06-15 07:00:00')
    Assert-Equal 'サービスの再起動と休止状態からの再開: 行なし・今の起動はずれない' ('{0}/{1}' -f $run.Code, $run.Log.Contains('今の起動: 2026-06-15 07:00')) '2/True'
    Assert-Equal 'まだ再起動していない指示（この PC）' $run.Log.Contains('※ 2026-06-15 14:00 に再起動の指示（LOCALPC\user）がありますが、まだ再起動していません') 'True'

    #endregion

    #region 読めないファイル

    Write-Section '読めないファイル'
    $badDir = Join-Path $workRoot 'bad'
    [void][IO.Directory]::CreateDirectory($badDir)
    [IO.File]::WriteAllBytes("$badDir\broken.evtx", [byte[]](1..200 | ForEach-Object { 7 }))
    [IO.File]::WriteAllBytes("$badDir\zero.evtx", [byte[]]@())
    [IO.File]::WriteAllText("$badDir\text.evtx", 'not a log')
    $app = New-SysEvent '2026-06-01 10:00' 1000 'Application Error' @(, @('param1', 0x01, 'app.exe'))
    $app.Channel = 'Application'
    New-EvtxFile "$badDir\application.evtx" @($app)
    $run = Invoke-Tool 'baddir' @('-Path', $badDir)
    Assert-Equal '読める .evtx がない: 止まって結果を残さない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains(('読める .evtx ファイルがありません: {0}' -f $badDir)), [bool]$run.Dir) '1/True/False'
    Assert-Equal '読めない理由を添える' ('{0}/{1}' -f $run.Log.Contains(('{0}\broken.evtx: ファイルを読めない' -f $badDir)), $run.Log.Contains(('{0}\application.evtx: 対象外のログ（System ログではありません（Application））' -f $badDir))) 'True/True'
    $run = Invoke-Tool 'apponly' @('-Path', "$badDir\application.evtx")
    Assert-Equal 'System でないログだけ: 止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains(('-Path に、System ログの .evtx がありません: {0}\application.evtx' -f $badDir)), [bool]$run.Dir) '1/True/False'
    $run = Invoke-Tool 'badmix' @('-Path', ('{0},{1}' -f $badDir, $mainFile))
    $x = Read-Result $run.Dir
    Assert-Equal '読めないファイルは飛ばしてほかを読む' ('{0}/{1}' -f $run.Code, $run.Log.Contains('スキップ 4 件（ファイルを読めない 3・対象外のログ 1）')) '0/True'
    Assert-Equal 'skipped.csv の理由' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', (Split-Path -Leaf $_.'対象') }) -join ',') '対象外のログ/application.evtx,ファイルを読めない/broken.evtx,ファイルを読めない/text.evtx,ファイルを読めない/zero.evtx'
    $run = Invoke-Tool 'notfound' @('-Path', "$workRoot\nothing.evtx")
    Assert-Equal 'ないファイル: 止まって理由を出す' ('{0}/{1}' -f $run.Code, $run.Log.Contains(('{0}\nothing.evtx: ファイルを読めない（見つかりません）' -f $workRoot))) '1/True'
    $emptyDir = Join-Path $workRoot 'emptydir'
    [void][IO.Directory]::CreateDirectory($emptyDir)
    $run = Invoke-Plain @('-Path', $emptyDir, '-NoSave')
    Assert-Equal '.evtx のないフォルダ' ('{0}/{1}' -f $run.Code, $run.Log.Contains('読める .evtx ファイルがありません')) '1/True'

    # 途中が壊れた .evtx: 壊れた場所の手前まで、1 件ずつ読む（64 件ずつまとめて読むと、手前の記録まで失う。3 回目のレビュー）
    $script:computer = 'BIG01'
    $big = New-Object 'System.Collections.Generic.List[object]'
    Add-Boot $big '2026-01-01 00:00:00'
    foreach ($i in 1..120) {
        $day = (Get-T '2026-01-01 00:00:00').AddDays($i)
        foreach ($n in 1..3) { $big.Add((New-Noise (Format-T $day.AddHours(-10 + $n)))) }
        Add-CleanStop $big (Format-T $day.AddMinutes(-2)) $ShutdownExe $Planned '0x80000000' '再起動' ('作業{0:000}' -f $i) 'CORP\yamada' 5 (Format-T $day.AddMinutes(-1)) (Format-T $day)
        Add-Boot $big (Format-T $day)
    }
    $bigFile = Join-Path $workRoot 'big\System.evtx'
    New-EvtxFile $bigFile $big.ToArray()
    $chunkCount = [int][math]::Floor(([IO.FileInfo]$bigFile).Length / 65536)
    Assert-Equal '複数のチャンクの .evtx ができた' ($chunkCount -ge 4) 'True'
    $run = Invoke-Plain @('-Path', $bigFile, '-NoSave')
    Assert-Equal '壊れていない .evtx は全部読む' $run.Log.Contains('再起動 120 回') 'True'
    $halfFile = Join-Path $workRoot 'half\System.evtx'
    New-BrokenCopy $bigFile $halfFile @(2)
    $run = Invoke-Tool 'half' @('-Path', $halfFile)
    $x = Read-Result $run.Dir
    $halfCount = @($x.Reboots).Count
    # 壊れた場所の手前で、読める最後の記録の時刻（すべての記録を 1 件ずつ読んで調べる）
    $lastReadable = $null
    $allReader = New-Object Diagnostics.Eventing.Reader.EventLogReader((New-Object Diagnostics.Eventing.Reader.EventLogQuery($halfFile, [Diagnostics.Eventing.Reader.PathType]::FilePath, '*')))
    $allReader.BatchSize = 1
    try { while ($e = $allReader.ReadEvent()) { $lastReadable = $e.TimeCreated; $e.Dispose() } } catch { }
    $allReader.Dispose()
    # 読み取りが停止の途中で止まると、いちばん新しい行は起動の記録のない停止になるので、起動のある行から探す
    $newestBoot = @($x.Reboots | Where-Object { $_.'起動' } | ForEach-Object { [datetime]::ParseExact($_.'起動', 'yyyy/MM/dd HH:mm:ss', $inv) } | Sort-Object)[-1]
    # 64 件ずつ読むと、壊れた場所の手前の記録を最大 63 件（停止と起動 9 回分）まとめて失う。1 件ずつなら、壊れた場所の直前の起動（1 日の内）まで行に出る
    Assert-Equal '途中が壊れた .evtx: 壊れた場所の直前の起動まで読む' ('{0}/{1}/{2}' -f $run.Code, ($halfCount -lt 120), ($null -ne $lastReadable -and $newestBoot -ge $lastReadable.AddDays(-1))) '0/True/True'
    Assert-Equal '途中が壊れた .evtx: skipped.csv に残す' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', ($_.'詳細' -replace '^(途中まで読んだ（記録 )[\d,]+( 件は結果に含む）).*$', '$1N$2') }) -join ',') 'ファイルを読めない/途中まで読んだ（記録 N 件は結果に含む）'
    Assert-Equal '途中が壊れた .evtx: 最後の起動の前の停止を、サービスの再起動と取り違えない' (@($x.Events | Where-Object { $_.'詳細' -like '*サービスの再起動*' }).Count) 0
    # 停止と起動の記録より前で壊れている: System ログを 1 件も読めなかったので、「該当なし」と区別して止める
    $noise = New-Object 'System.Collections.Generic.List[object]'
    foreach ($i in 1..400) { $noise.Add((New-Noise (Format-T (Get-T '2026-01-01 00:00:00').AddMinutes($i)))) }
    Add-Boot $noise '2026-01-02 00:00:00'
    foreach ($i in 1..300) { $noise.Add((New-Noise (Format-T (Get-T '2026-01-02 01:00:00').AddMinutes($i)))) }
    $noiseFile = Join-Path $workRoot 'noise\System.evtx'
    New-EvtxFile $noiseFile $noise.ToArray()
    $allBad = Join-Path $workRoot 'allbad\System.evtx'
    $noiseChunks = [int](([IO.FileInfo]$noiseFile).Length / 65536)
    New-BrokenCopy $noiseFile $allBad @(1..($noiseChunks - 1)) 512
    $run = Invoke-Tool 'allbad' @('-Path', $allBad)
    Assert-Equal '停止と起動の記録を 1 件も読めない: 止まって結果を残さない' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('System ログを読めませんでした'), $run.Log.Contains(('{0}: ファイルを読めない' -f $allBad)), [bool]$run.Dir) '1/True/True/False'

    #endregion

    #region リンク・権限なし・パス

    Write-Section 'リンク・権限なし・パス'
    $linkDir = Join-Path $workRoot 'linkdir'
    [void][IO.Directory]::CreateDirectory($linkDir)
    [IO.File]::Copy($mainFile, "$linkDir\System.evtx")
    $null = & cmd.exe /d /c mklink /J "$linkDir\loop" "$linkDir" 2>&1
    if ([IO.Directory]::Exists("$linkDir\loop")) {
        $run = Invoke-Tool 'link' @('-Path', $linkDir)
        $x = Read-Result $run.Dir
        Assert-Equal 'ジャンクションはたどらない' ('{0}/{1}/{2}' -f $run.Log.Contains('予期しない停止 4 回'), ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', $_.'詳細' }) -join ','), $run.Log.Contains('スキップ 1 件（リンク）')) 'True/リンク/ジャンクション/True'
    } else {
        Write-Skip 'ジャンクション' 'ジャンクションを作れなかったため'
    }
    $denyDir = Join-Path $workRoot 'deny'
    [void][IO.Directory]::CreateDirectory("$denyDir\closed")
    [IO.File]::Copy($mainFile, "$denyDir\System.evtx")
    [IO.File]::Copy($mainFile, "$denyDir\denied.evtx")
    [IO.File]::Copy($mainFile, "$denyDir\closed\System.evtx")
    if ((Add-Deny "$denyDir\denied.evtx" '(R)') -and (Add-Deny "$denyDir\closed" '(RD)')) {
        $run = Invoke-Tool 'deny' @('-Path', $denyDir)
        $x = Read-Result $run.Dir
        Assert-Equal '読めないファイルとフォルダは飛ばして、権限なし' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', (Split-Path -Leaf $_.'対象') }) -join ',') '権限なし/closed,権限なし/denied.evtx'
        Assert-Equal 'ほかは読む' ('{0}/{1}' -f $run.Log.Contains('予期しない停止 4 回'), $run.Log.Contains('スキップ 2 件（権限なし）')) 'True/True'
    } else {
        Write-Skip '権限なし' 'アクセス拒否を設定できなかったため'
    }
    $longDir = "$workRoot\" + ('L' * 120) + '\' + ('M' * 120)
    [void][IO.Directory]::CreateDirectory('\\?\' + $longDir)
    [IO.File]::Copy($mainFile, '\\?\' + "$longDir\System.evtx")
    $run = Invoke-Tool 'longdir' @('-Path', $longDir)
    Assert-Equal '260 文字を超えるパス（フォルダ）' $run.Log.Contains('予期しない停止 4 回') 'True'
    $run = Invoke-Tool 'longfile' @('-Path', "$longDir\System.evtx")
    Assert-Equal '260 文字を超えるパス（ファイル）' ('{0}/{1}' -f $run.Log.Contains('予期しない停止 4 回'), $run.Log.Contains(('対象: {0}\System.evtx' -f $longDir))) 'True/True'
    $longOut = "$workRoot\" + ('O' * 120) + '\' + ('P' * 120)
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $mainDir -OutputFolder $longOut 2>&1
    $longRun = @(if ([IO.Directory]::Exists('\\?\' + $longOut)) { [IO.Directory]::GetDirectories('\\?\' + $longOut) })
    Assert-Equal '260 文字を超える保存先' $(if ($longRun.Count -eq 1) { @([IO.Directory]::GetFiles($longRun[0])).Count } else { 0 }) 4
    $jpDir = Join-Path $workRoot 'スペース ある\ログ'
    [void][IO.Directory]::CreateDirectory($jpDir)
    [IO.File]::Copy($mainFile, "$jpDir\System.evtx")
    $run = Invoke-Tool 'japanese path' @('-Path', "$jpDir\System.evtx")
    Assert-Equal '日本語とスペースを含むパス' $run.Log.Contains('予期しない停止 4 回') 'True'
    $bracketDir = Join-Path $workRoot 'br'
    [void][IO.Directory]::CreateDirectory("$bracketDir\[a]")
    [void][IO.Directory]::CreateDirectory("$bracketDir\a")
    [IO.File]::Copy($mainFile, "$bracketDir\[a]\System.evtx")
    [IO.File]::Copy("$oldDir\System.evtx", "$bracketDir\a\System.evtx")
    $run = Invoke-Plain @('-Path', "$bracketDir\[a]\System.evtx", '-NoSave')
    Assert-Equal '[a] を含むパス（隣の a を読まない）' $run.Log.Contains('予期しない停止 4 回') 'True'
    $run = Invoke-From "$bracketDir\[a]" @('-Path', 'System.evtx', '-NoSave')
    Assert-Equal '[a] のフォルダから相対パス' $run.Log.Contains('予期しない停止 4 回') 'True'

    #endregion

    #region 引数の誤りと保存先

    Write-Section '引数の誤り'
    foreach ($case in @(
        @(@('-Since', '2026-13-01'), '-Since の日時を読めません: 2026-13-01'),
        @(@('-Since', '2026-06-15 18:00', '-Until', '2026-06-15 09:00'), '-Since の日時（2026-06-15 18:00）が、-Until の日時（2026-06-15 09:00）と同じか、それより後になっています'),
        @(@('-Since', '2099-01-01'), 'と同じか、それより後になっています'),
        @(@('-Until', 'abc'), '-Until の日時を読めません: abc'),
        @(@('-NoSave', '-OutputFolder', "$workRoot\out\x"), '-NoSave のときは、-OutputFolder は指定できません'),
        @(@('-Sinse', '30d'), 'どのパラメーターにも当たらない引数があります: -Sinse 30d')
    )) {
        $run = Invoke-Plain (@('-Path', $mainDir) + $case[0])
        Assert-Equal ('止まる: {0}' -f ($case[0] -join ' ')) ('{0}/{1}' -f $run.Code, $run.Log.Contains($case[1])) '1/True'
    }

    Write-Section '保存先'
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $mainDir 2>&1
    $defaultRun = Get-RunDir (Split-Path -Parent $tool)
    Assert-Equal '-OutputFolder なしはスクリプトのフォルダ' $(if ($defaultRun) { @(Get-ChildItem -LiteralPath $defaultRun -File).Count } else { 0 }) 4
    $fileOut = Join-Path $workRoot 'out\afile'
    [IO.File]::WriteAllText($fileOut, 'x')
    $run = Invoke-Plain @('-Path', $mainDir, '-OutputFolder', $fileOut)
    Assert-Equal '保存先がファイル: 集める前に止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains(('保存先に書き込めません: {0}' -f $fileOut)), $run.Log.Contains('-NoSave を付けると、画面に表示だけします')) '1/True/True'
    $letter = Get-FreeLetter
    if ($letter) {
        $run = Invoke-Plain @('-Path', $mainDir, '-OutputFolder', "${letter}:\nope")
        Assert-Equal 'ないドライブ: 止まる' ('{0}/{1}' -f $run.Code, $run.Log.Contains('保存先に書き込めません')) '1/True'
    } else {
        Write-Skip 'ないドライブ' '空いているドライブ文字がないため'
    }
    $sameDir = Join-Path $workRoot 'out\same'
    [void][IO.Directory]::CreateDirectory($sameDir)
    $stampNow = Get-Date
    foreach ($s in 0..40) {
        $older = Join-Path $sameDir ('RebootHistory_{0}_{1}' -f $env:COMPUTERNAME, $stampNow.AddSeconds($s).ToString('yyyyMMdd_HHmmss'))
        [void][IO.Directory]::CreateDirectory($older)
        [IO.File]::WriteAllText("$older\keep.txt", 'keep')
    }
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $mainDir -OutputFolder $sameDir 2>&1
    Assert-Equal '同じ名前のフォルダがあれば _2 を付ける' (@(Get-ChildItem -LiteralPath $sameDir -Directory -Filter 'RebootHistory_*_2').Count) 1
    Assert-Equal '前の結果を残す' (@(Get-ChildItem -LiteralPath $sameDir -Recurse -Filter 'keep.txt').Count) 41
    $patched = New-PatchedTool 'savefail' @("Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'events.csv'))", "throw 'disk full'; Save-Csv (ConvertTo-LongPath (Join-Path `$runDir 'events.csv'))")
    $run = Invoke-Tool 'savefail' @('-Path', $manyDir) $patched
    Assert-Equal '保存に失敗: 警告を出して、書きかけを残さない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした: disk full'), [bool]$run.Dir) '1/True/False'
    Assert-Equal '保存に失敗: すべて画面に出し、保存先の行を出さない' ('{0}/{1}/{2}' -f $run.Log.Contains('作業01'), $run.Log.Contains('参照してください'), $run.Log.Contains('保存先:')) 'True/False/False'
    $patched = New-PatchedTool 'crash' @('$periodRecords = @(', "throw 'boom'; `$periodRecords = @(")
    $run = Invoke-Tool 'crash' @('-Path', $mainDir) $patched
    Assert-Equal '想定外のエラー: 結果を残さず、中断とは区別する' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('エラーで止まりました。結果は保存していません。'), $run.Log.Contains('中断しました'), [bool]$run.Dir) '1/True/False/False'

    #endregion

    #region 起動

    Write-Section 'コマンドプロンプトから起動'
    $cmdHead = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $unknownOut = Join-Path $workRoot 'out_unknown'
    $run = Invoke-ViaCmd ('{0} -OutputFolder "{1}\" -NoSuchName "a b"' -f $cmdHead, $unknownOut)
    Assert-Equal '"パス\" の後ろに空白を含む値（引用符のヒントを出し、保存先を作らない）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('-OutputFolder の値に引用符（"）が残っています'), (Test-Path -LiteralPath $unknownOut)) '1/True/False'
    $run = Invoke-ViaCmd ('{0} -NoSuchName' -f $cmdHead)
    Assert-Equal 'どのパラメーターにも当たらない引数（日本語の案内を出して止める）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('どのパラメーターにも当たらない引数があります: -NoSuchName')) '1/True'
    $run = Invoke-ViaCmd ('{0} -NoSave -Since "2026-06-01 00:00" -Path "{1}\"' -f $cmdHead, $mainDir)
    Assert-Equal '最後の引数が \"（引用符を外して動く）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('予期しない停止 4 回')) '0/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -NoSave -Since 3000d' -f $cmdHead, $mainDir)
    Assert-Equal '-Path の途中で \"（止まってヒントを出す）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('-Path の値に引用符（"）が残っています')) '1/True'
    $run = Invoke-ViaCmd ('{0} -Path "{1}\" -NoSave -Since "2026-06-01 00:00"' -f $cmdHead, $mainDir)
    Assert-Equal '-Path の途中で \" の後ろに空白を含む値（英語のエラーでなく、ヒントを出す）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('-Path の値に引用符（"）が残っています'), $run.Log.Contains('positional')) '1/True/False'
    $run = Invoke-ViaCmd ('{0} -Path "{1}" -OutputFolder "{2}\" -Since 3000d' -f $cmdHead, $mainDir, "$workRoot\out\cmdq")
    Assert-Equal '-OutputFolder の途中で \"（止まって、保存先を作らない）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('-OutputFolder の値に引用符'), (Test-Path -LiteralPath "$workRoot\out\cmdq")) '1/True/False'
    $run = Invoke-ViaCmd ("{0} -NoSave -Path '{1}' -Since '2026-06-01' -Until '2026-06-30'" -f $cmdHead, $mainDir)
    Assert-Equal "cmd の ' はそのまま届く" ('{0}/{1}' -f $run.Code, $run.Log.Contains('予期しない停止 4 回')) '0/True'
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}" -Since 2000-01-01 -Until 2000-01-02' -f $cmdHead, $mainDir)
    Assert-Equal 'cmd で受け取る終了コード（該当なしは 2）' $run.Code 2
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}" -Since abc' -f $cmdHead, $mainDir)
    Assert-Equal 'cmd で受け取る終了コード（止まると 1）' $run.Code 1
    $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}\System.evtx"' -f $cmdHead, $jpDir)
    Assert-Equal 'cmd からスペースと日本語を含むパス' ('{0}/{1}' -f $run.Code, $run.Log.Contains('予期しない停止 4 回')) '0/True'

    Write-Section '制約付き言語モードとヘルプ'
    $clmCommand = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '{0}' -Path '{1}' -OutputFolder '{2}'; exit `$LASTEXITCODE" -f $tool, $mainDir, "$workRoot\clm"
    $clm = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command $clmCommand 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $clmCode = $LASTEXITCODE
    Assert-Equal '制約付き言語モード: 止まる' ('{0}/{1}/{2}' -f $clmCode, $clm.Contains('制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）'), (Test-Path -LiteralPath "$workRoot\clm")) '1/True/False'
    $help = (@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command ("Get-Help '{0}' -Full | Out-String -Width 300" -f $tool) 2>&1) | ForEach-Object { "$_" }) -join "`n"
    Assert-Equal 'ヘルプの .NOTES' ('{0}/{1}/{2}/{3}' -f $help.Contains('Get-RebootHistory 1.0.0'), $help.Contains('Copyright (c) 2026 結進舎'), $help.Contains('Released under the MIT License.'), $help.Contains('https://github.com/kando19804-oss/workbench/blob/main/LICENSE')) 'True/True/True/True'
    Assert-Equal 'ヘルプのパラメーター' (@(@('Since', 'Until', 'Path', 'OutputFolder', 'NoSave') | Where-Object { -not $help.Contains("-$_") }) -join ',') ''

    Write-Section 'Ctrl+C で中断'
    # 集め始める前に長く待つコピーで止める（読むだけなので、途中で止めても何も壊れない）
    $slowTool = New-PatchedTool 'slow' @('    $readCount = 0', "    Start-Sleep -Seconds 60`r`n    `$readCount = 0")
    $stopScript = New-HelperScript 'stop.ps1' @(
        'param($Tool, $LogDir, $OutRoot)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $LogDir).AddParameter(''OutputFolder'', $OutRoot)'
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
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -OutputFolder '{2}' | Out-Null" -f $tool, $mainDir, $cultureOut
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        $summaryText = $x.Lines -join "`n"
        Assert-Equal 'fi-FI でも表の日時は yyyy-MM-dd HH:mm' $summaryText.Contains(' 2026-06-03 22:00  22:02') 'True'
        Assert-Equal 'fi-FI でも所要時間の小数点は .' ([regex]::IsMatch($summaryText, '所要 \d+\.\d秒')) 'True'
        Assert-Equal 'fi-FI でも CSV の日時は yyyy/MM/dd HH:mm:ss' (@($x.Reboots)[0].'起動') '2026/06/15 07:00:00'
        Assert-Equal 'fi-FI でも時間は H:mm:ss' (Get-Row $x.Reboots { $_.'起動' -eq '2026/06/03 22:02:00' }).'停止していた時間' '="0:01:30"'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }

    #endregion

    #region この PC のイベントログ

    Write-Section 'この PC のイベントログ'
    $run = Invoke-Tool 'local' @()
    $x = Read-Result $run.Dir
    Assert-Equal '本物: 終了コード（0 か 2）' ($run.Code -eq 0 -or $run.Code -eq 2) 'True'
    Assert-Equal '本物: 対象の行と今の起動の行' ('{0}/{1}' -f $run.Log.Contains(('対象: {0}   System ログ: ' -f $env:COMPUTERNAME)), $run.Log.Contains('今の起動: ')) 'True/True'
    $cimBoot = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    $shown = [regex]::Match($run.Log, '今の起動: (\d{4}-\d\d-\d\d \d\d:\d\d)').Groups[1].Value
    Assert-Equal '本物: 今の起動は OS の情報の起動と同じか後' ($shown -ge $cimBoot.ToString('yyyy-MM-dd HH:mm', $inv)) 'True'
    Assert-Equal '本物: 種類はすべて決まった値' (@($x.Reboots | Where-Object { @('再起動', 'シャットダウン', '予期しない停止', '不明') -notcontains $_.'種類' } | ForEach-Object { $_.'種類' } | Select-Object -Unique) -join ',') ''
    Assert-Equal '本物: 起動の種類はすべて決まった値' (@($x.Reboots | Where-Object { @('通常', '高速スタートアップ', '') -notcontains $_.'起動の種類' } | ForEach-Object { $_.'起動の種類' } | Select-Object -Unique) -join ',') ''
    $operations = '停止の指示', '停止の取り消し', '停止の開始', 'イベントログの停止', 'OS の停止', 'OS の起動', '起動の種類', 'イベントログの起動', '予期しない停止', '正常に停止しないまま起動', 'ブルースクリーン', '理由の入力', 'ログの消去'
    Assert-Equal '本物: 操作はすべて決まった値' (@($x.Events | Where-Object { $operations -notcontains $_.'操作' } | ForEach-Object { $_.'操作' } | Select-Object -Unique) -join ',') ''
    Assert-Equal '本物: 停止の時刻がある行は、停止が起動より前' (@($x.Reboots | Where-Object { $_.'停止' -and $_.'起動' -and $_.'停止' -gt $_.'起動' }).Count) 0
    Assert-Equal '本物: その他のエラーはない' (@($x.Skipped | Where-Object { $_.'理由' -eq 'その他のエラー' }).Count) 0

    #endregion

    #region 再発防止の一覧

    Write-Section 'レビューで見つかった問題の再発防止（上で確かめている）'
    # 1. 短い間隔の起動（40 秒おきのブルースクリーン・理由を入力してすぐの再起動・遅れて書かれた 1001）: quick
    # 2. 年をまたぐ起動の表示と 120 桁: year
    # 3. 停止の指示のある予期しない停止のユーザー・6006 のあと 41 だけ（6008 なし）の停止の時刻: main の 06-14
    # 4. 秒の端数（停止時間は CSV の時刻の差と同じ）: main の 06-03（22:00:30 → 22:02:00 の 12 の StartTime は端数なし）、本物
    # 5. 同じファイルの重複指定: 重複
    # 6. 先頭は読めるが途中が壊れた .evtx（1 件ずつ読む）・停止と起動の記録より前で壊れた .evtx: half・allbad
    # 7. イベントログのサービスだけの再起動（行にしない・今の起動がずれない）: main の 06-13、quick の 06-25、local4
    # 8. 12 だけが読めない起動（6005 で補う・27 の時刻を使う）: quick の 06-23・06-24
    # 9. 入力された理由（1076）を画面で先に書く: main の 06-08
    # 10. 高速スタートアップの起動の前の備考・今の起動・最後の起動のあとの停止: local1・local3

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
