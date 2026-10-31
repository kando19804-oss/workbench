<#
.SYNOPSIS
    指定したポートを使っているプロセスを突き止めます。

.DESCRIPTION
    このコンピューターの TCP と UDP のポートについて、使っているプロセス（PID、プロセス名、サービス、
    実行ユーザー、起動日時、実行ファイル、コマンドライン）を一覧にし、画面に表示します。
    結果は CSV（BOM 付き UTF-8）でも保存します。ポートを指定しなければ、待ち受けているすべてのポートを一覧にします。

    読み取りだけを行い、プロセスを止めたり設定を変えたりすることはありません。
    コマンドラインはパスワードなどが含まれることがあるため、画面には表示せず CSV にだけ書きます。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Get-PortProcess.ps1 -Port 8080

.PARAMETER Port
    調べるポート。"80,443,8000-8100" のように、カンマと範囲（-）で書けます（1～65535）。
    ローカルのポートかリモートのポートのどちらかが一致すれば対象にします。省略すると、すべてのポートが対象です。

.PARAMETER Established
    TCP の待ち受けに加えて、接続中などの通信（待ち受け以外の状態）も出します。UDP は常に出します。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.PARAMETER ConfigFile
    設定ファイルのパス。Port、Established、OutputFolder、NoSave を「項目 = 値」の形で書いておけます。
    コマンドラインで指定した値は、設定ファイルより優先します。書き方は README を参照してください。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-PortProcess.ps1

    待ち受けているすべてのポートを一覧にします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-PortProcess.ps1 -Port "8080,443"

    8080 番と 443 番を使っているプロセスを調べます。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-PortProcess.ps1 -Port 1433 -Established

    1433 番で待ち受けているプロセスと、ほかのサーバーの 1433 番につないでいるプロセスを調べます。

.NOTES
    Get-PortProcess 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Position = 0)]
    [string[]]$Port = @(),

    [switch]$Established,

    [string]$OutputFolder,

    [switch]$NoSave,

    [string]$ConfigFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$ServiceShown  = 3     # 画面に表示するサービスの数（残りは「ほか N 個」）
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$Columns       = @('プロトコル', 'ローカルアドレス', 'ローカルポート', 'リモートアドレス', 'リモートポート', '状態', 'PID',
                   'プロセス', 'サービス', '実行ユーザー', '起動日時', '実行ファイル', 'コマンドライン')

#region 補助関数

# コマンドプロンプトから渡された引数に残る引用符を外す
# （"C:\" が C:" として届く問題や、cmd では ' がそのまま届く問題への対策）
function Remove-Quote([string]$Text) {
    $t = $Text.Trim().Trim('"')
    if ($t.Length -ge 2 -and $t.StartsWith("'") -and $t.EndsWith("'")) { $t = $t.Substring(1, $t.Length - 2) }
    return $t
}

# 日付や数値の区切り文字が、OS の地域設定によって変わらないようにする
function Format-Invariant([string]$Format) {
    return [string]::Format($Invariant, $Format, [object[]]$args)
}

# .NET のメソッドで起きた例外は MethodInvocationException に包まれて届くため、元の例外を取り出す
function Get-InnerException([Exception]$Exception) {
    while ($Exception -is [Management.Automation.MethodInvocationException] -and $Exception.InnerException) {
        $Exception = $Exception.InnerException
    }
    return $Exception
}

# 例外メッセージに含まれる \\?\ 付きのパスを、表示用のパスに戻す。
# .NET のメッセージは途中や末尾に改行が入ることがあるため、1 行にまとめる
function Get-ErrorText([Exception]$Exception) {
    $text = $Exception.Message.Replace('\\?\UNC\', '\\').Replace('\\?\', '')
    return ($text -replace '\s*[\r\n]+\s*', ' ').Trim()
}

# 5.1 は 260 文字を超えるパスを通常の書き方では読めないため、ファイルを読み書きするときは \\?\ 付きのパスを使う
function ConvertTo-LongPath([string]$FullPath) {
    if ($FullPath.StartsWith('\\?\')) { return $FullPath }
    if ($FullPath.StartsWith('\\')) { return '\\?\UNC\' + $FullPath.Substring(2) }
    return '\\?\' + $FullPath
}

# Excel は = + - @ で始まる値を数式として扱うため、その形の値は ="..." で囲む。
# 値の中の " は "" に重ね、Excel の式に書けない 255 文字を超える値は、先頭に ' を付けて文字として読ませる
function ConvertTo-ExcelText([string]$Text) {
    if ($Text -notmatch '^[=+\-@]') { return $Text }
    if ($Text.Length -gt 255) { return "'" + $Text }
    return '="' + $Text.Replace('"', '""') + '"'
}

# 5.1 と 7 のどちらでも BOM 付き UTF-8 で保存する（Export-Csv の -Encoding UTF8 は 7 だと BOM なしになる）。
# 行が 0 件のときは見出しだけを書く
function Save-Csv($Rows, [string[]]$Columns, [string]$FilePath) {
    if ($null -eq $Rows) { $Rows = @() }
    $lines = @($Rows | ConvertTo-Csv -NoTypeInformation)
    if ($lines.Count -eq 0) { $lines = @(($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') }
    [IO.File]::WriteAllLines($FilePath, [string[]]$lines, $Utf8Bom)
}

function Stop-WithError([string]$Message) {
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value) {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -OutputFolder "C:\report"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
}

# 5.1 の Get-Content は BOM のない UTF-8 を Shift_JIS として読んでしまうため、文字コードを自分で判定する。
# BOM があればそれに従い、なければ UTF-8 として読めるかを確かめ、読めなければ Shift_JIS とみなす
function Read-ConfigText([string]$FilePath) {
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return (New-Object Text.UTF8Encoding $false).GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2) }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return [Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2) }
    try { return (New-Object Text.UTF8Encoding $false, $true).GetString($bytes) } catch { }
    return [Text.Encoding]::GetEncoding(932).GetString($bytes)
}

# 設定ファイルの値は、前後を同じ引用符で囲んだときだけ外す（片方だけの引用符は書き間違いとして残す）
function Remove-ConfigQuote([string]$Text) {
    $t = $Text.Trim()
    if ($t.Length -ge 2 -and (($t.StartsWith('"') -and $t.EndsWith('"')) -or ($t.StartsWith("'") -and $t.EndsWith("'")))) {
        $t = $t.Substring(1, $t.Length - 2)
    }
    return $t
}

# 設定ファイルの書き間違いは、何行目かと設定ファイルのパスを添えて止める
function Stop-WithConfigError([int]$LineNumber, [string]$Message, [string]$Hint) {
    $text = '設定ファイルの {0} 行目: {1}' -f $LineNumber, $Message
    if ($Hint) { $text += "`r`n" + $Hint }
    Stop-WithError ($text + "`r`n" + ('ファイル: {0}' -f $configPath))
}

# 設定ファイルから来た値なら、何行目かを添える（見つからない・書き込めないときに原因を探しやすくするため）
function Get-ConfigNote([string]$Name) {
    if ($configApplied.ContainsKey($Name)) { return '（設定ファイルの {0} 行目）' -f $configApplied[$Name] }
    return ''
}

# .NET の IsPathRooted は | などの文字で例外を出すため、先頭の形で判定する
function Test-RootedPath([string]$Text) {
    return $Text -match '^([A-Za-z]:|[\\/])'
}

# 5.1 は、名前に [ や ] を含むフォルダで起動すると、その名前をワイルドカードとして PowerShell の現在の場所を決める
# （プロセスの現在のフォルダは正しい）。移る先は、一致するフォルダがなければ $PSHOME、1 つだけ一致すればそのフォルダ
# （[a] なら隣の a）、パターンとして正しくない（a[b）か 2 つ以上一致する（[ab]）とドライブのルートになる。
# そのときだけ、相対パスをプロセスの現在のフォルダを基準に解決する。対話の PowerShell で cd したあとは
# プロセスの現在のフォルダが古いままなので、「違っていたら常にプロセスの現在のフォルダを使う」とはしない
$startDir   = $null
$processDir = [Environment]::CurrentDirectory.TrimEnd('\')
$location   = (Get-Location).ProviderPath.TrimEnd('\')
if ($processDir -match '[\[\]]' -and $location -ne $processDir) {
    $misplaced = $location -eq $PSHOME.TrimEnd('\') -or $location -match '^[A-Za-z]:$'
    if (-not $misplaced) {
        try { $misplaced = [Management.Automation.WildcardPattern]::new($processDir, [Management.Automation.WildcardOptions]::IgnoreCase).IsMatch($location) } catch { }
    }
    if ($misplaced) { $startDir = $processDir }
}

function Resolve-StartPath([string]$Text) {
    if ($startDir -and $Text -and -not (Test-RootedPath $Text)) { return $startDir + '\' + $Text }
    return $Text
}

# %TEMP% などの環境変数を展開する。名前を打ち間違えると %名前% のまま残り、その名前のフォルダに保存してしまうため止める。
# URL エンコードされたフォルダ名（%E3%81%82 など）を誤って止めないよう、16 進 2 桁の %XX% は対象にしない
function Expand-ConfigVariable([int]$LineNumber, [string]$Text) {
    $expanded = [Environment]::ExpandEnvironmentVariables($Text)
    if ($expanded -match '%(?![0-9A-Fa-f]{2}%)[A-Za-z_][^%\\=\s]*%') { Stop-WithConfigError $LineNumber ('定義されていない環境変数です: {0}' -f $Matches[0]) }
    return $expanded
}

# 画面の表で列をそろえるため、全角文字を 2 文字分の幅として数える（半角カタカナは 1 文字分）
function Get-DisplayWidth([string]$Text) {
    $width = 0
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int]$ch
        if ($code -lt 0x100 -or ($code -ge 0xFF61 -and $code -le 0xFF9F)) { $width += 1 }
        elseif ($code -ge 0xDC00 -and $code -le 0xDFFF) { }   # サロゲートペアの後半は、前半で 2 文字分を数えた
        else { $width += 2 }
    }
    return $width
}

function Format-Cell([string]$Text, [int]$Width, [bool]$AlignRight) {
    $pad = ' ' * [math]::Max(0, $Width - (Get-DisplayWidth $Text))
    if ($AlignRight) { return $pad + $Text }
    return $Text + $pad
}

# 見出しと行（文字列の配列）から、列をそろえた表の行を作る
function Format-Table([string[]]$Head, [bool[]]$Right, $Rows) {
    $widths = @(foreach ($i in 0..($Head.Count - 1)) {
        $w = Get-DisplayWidth $Head[$i]
        foreach ($row in $Rows) { $w = [math]::Max($w, (Get-DisplayWidth $row[$i])) }
        $w
    })
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($row in @(, $Head) + @($Rows)) {
        $cells = @(foreach ($i in 0..($Head.Count - 1)) { Format-Cell $row[$i] $widths[$i] $Right[$i] })
        $lines.Add((' ' + ($cells -join '  ')).TrimEnd())
    }
    return , $lines.ToArray()
}

# ポートの指定 1 つ（80 や 8000-8100）を、@(開始, 終了) にする。書き方が誤っていれば $null
function ConvertTo-PortRange([string]$Text) {
    $from = 0
    $to   = 0
    if ($Text -match '^(\d+)\s*-\s*(\d+)$') {
        if (-not [int]::TryParse($Matches[1], [ref]$from) -or -not [int]::TryParse($Matches[2], [ref]$to)) { return $null }
    } elseif ($Text -match '^\d+$') {
        if (-not [int]::TryParse($Text, [ref]$from)) { return $null }
        $to = $from
    } else {
        return $null
    }
    if ($from -lt 1 -or $to -gt 65535 -or $from -gt $to) { return $null }
    return , @($from, $to)
}

# IPv6 のアドレスは、ポートと区別できるよう [ ] で囲む
function Format-Endpoint([string]$Address, [int]$PortNumber) {
    if ($Address.Contains(':')) { return '[{0}]:{1}' -f $Address, $PortNumber }
    return '{0}:{1}' -f $Address, $PortNumber
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの
$configItems   = 'Port', 'Established', 'OutputFolder', 'NoSave'
$configSwitch  = 'Established', 'NoSave'
$configPath    = $null
$configApplied = @{}   # 設定ファイルの値を使った項目 → 行番号

if ($ConfigFile) {
    $configText = Remove-Quote $ConfigFile
    if ($configText.Contains('"')) { Stop-WithQuoteError '-ConfigFile' $configText }
    $configText = Resolve-StartPath $configText
    $configPath = $configText
    try {
        $configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($configText)
        $configPath = [IO.Path]::GetFullPath($configPath)
    } catch { }
    if (-not [IO.File]::Exists((ConvertTo-LongPath $configPath))) {
        # エクスプローラーで拡張子を表示していないと、メモ帳で port.ini.txt として保存してしまうことがある
        $message = '設定ファイルが見つかりません: {0}' -f $configPath
        if ([IO.File]::Exists((ConvertTo-LongPath ($configPath + '.txt')))) { $message += "`r`n" + ('拡張子が .txt になっていませんか: {0}.txt' -f $configPath) }
        Stop-WithError $message
    }
    try {
        $configLines = (Read-ConfigText (ConvertTo-LongPath $configPath)) -split '\r\n|\n|\r'
    } catch {
        Stop-WithError ('設定ファイルを読めません: {0}（{1}）' -f $configPath, (Get-ErrorText (Get-InnerException $_.Exception)))
    }

    $configDir    = Split-Path -LiteralPath $configPath   # -LiteralPath には -Parent を付けられない（親フォルダを返すのが既定）
    $configValues = @{}
    $configLineOf = @{}
    $portItems    = New-Object 'System.Collections.Generic.List[string]'
    $lineNumber   = 0
    foreach ($rawLine in $configLines) {
        $lineNumber++
        $lineText = $rawLine.Trim()
        # # と ; はパスに使える文字なので、行の先頭にあるときだけコメントとみなす
        if ($lineText -eq '' -or $lineText.StartsWith('#') -or $lineText.StartsWith(';')) { continue }
        $separator = $lineText.IndexOf('=')
        if ($separator -le 0) { Stop-WithConfigError $lineNumber ('「項目 = 値」の形になっていません: {0}' -f $lineText) }

        $itemKey  = $lineText.Substring(0, $separator).Trim()
        $itemName = $configItems | Where-Object { $_ -eq $itemKey }
        if (-not $itemName) { Stop-WithConfigError $lineNumber ('知らない項目です: {0}' -f $itemKey) ('使える項目: {0}' -f ($configItems -join ', ')) }

        $rawValue  = $lineText.Substring($separator + 1).Trim()
        $itemValue = Remove-ConfigQuote $rawValue
        # 値が空の行は、書かなかったのと同じ扱い（ひな形の空の行の下に書き足しても、2 回書いたことにしない）
        if ($itemValue -eq '') { continue }

        # Port は長くなりやすいので複数の行に分けて書ける。ほかの項目は、どちらが効いているか分からなくなるため 1 回だけ
        if ($itemName -ne 'Port' -and $configLineOf.ContainsKey($itemName)) {
            Stop-WithConfigError $lineNumber ('{0} が 2 回書かれています（{1} 行目と {2} 行目）' -f $itemName, $configLineOf[$itemName], $lineNumber)
        }
        if (-not $configLineOf.ContainsKey($itemName)) { $configLineOf[$itemName] = $lineNumber }

        if ($itemName -eq 'Port') {
            # Port は複数の行に書けるため、どの行の誤りかが分かるよう、ここで行ごとに確かめる
            foreach ($part in ($itemValue -split ',')) {
                $text = $part.Trim()
                if ($text -and -not (ConvertTo-PortRange $text)) {
                    Stop-WithConfigError $lineNumber ('Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: {0}' -f $text)
                }
            }
            $portItems.Add($itemValue)
        } elseif ($configSwitch -contains $itemName) {
            if ($itemValue -eq 'true') { $configValues[$itemName] = $true }
            elseif ($itemValue -eq 'false') { $configValues[$itemName] = $false }
            else { Stop-WithConfigError $lineNumber ('{0} は true か false で指定してください: {1}' -f $itemName, $itemValue) }
        } else {
            # パスの項目は、コマンドプロンプトで書くときと同じように %TEMP% などの環境変数を展開する
            if ($itemValue.Contains('"')) { Stop-WithConfigError $lineNumber ('{0} の値に引用符（"）が残っています: {1}' -f $itemName, $itemValue) }
            $itemValue = Expand-ConfigVariable $lineNumber $itemValue
            # 相対パスは、設定ファイルのあるフォルダを基準にする（タスクスケジューラでは現在のフォルダが System32 になるため）
            if (-not (Test-RootedPath $itemValue)) {
                $itemValue = $configDir.TrimEnd('\') + '\' + $itemValue
                try { $itemValue = [IO.Path]::GetFullPath($itemValue) } catch { }
            }
            $configValues[$itemName] = $itemValue
        }
    }
    if ($portItems.Count -gt 0) { $configValues['Port'] = $portItems.ToArray() }

    # コマンドラインで指定した項目は、設定ファイルより優先する
    foreach ($itemName in @($configValues.Keys)) {
        if (-not $PSBoundParameters.ContainsKey($itemName)) {
            Set-Variable -Name $itemName -Value $configValues[$itemName]
            $configApplied[$itemName] = $configLineOf[$itemName]
        }
    }
}

#endregion

#region 引数の確認

# -File で起動すると配列を渡せないため、カンマ区切りの文字列で受ける。範囲は 8000-8100 の形で書ける
$portSet   = New-Object 'System.Collections.Generic.HashSet[int]'
$portTexts = @()
foreach ($entry in $Port) {
    foreach ($part in ($entry -split ',')) {
        $text = (Remove-Quote $part).Trim("'").Trim()
        if ($text -eq '') { continue }
        $range = ConvertTo-PortRange $text
        if (-not $range) { Stop-WithError ('-Port は 1～65535 の数値か、8000-8100 のような範囲で指定してください: {0}' -f $text) }
        for ($n = $range[0]; $n -le $range[1]; $n++) { [void]$portSet.Add($n) }
        $portTexts += $(if ($range[0] -eq $range[1]) { [string]$range[0] } else { '{0}-{1}' -f $range[0], $range[1] })
    }
}
# "," のように、値はあるのにポートが 1 つもないときは、絞り込みが消えて全件になるのを防ぐため止める
if ($portSet.Count -eq 0 -and @($Port | Where-Object { "$_".Trim() -ne '' }).Count -gt 0) {
    Stop-WithError ('-Port にポートがありません: {0}' -f ($Port -join ','))
}
$hasPortFilter = $portSet.Count -gt 0

$outputText = $null
if ($OutputFolder) {
    $outputText = Remove-Quote $OutputFolder
    if ($outputText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outputText }
    $outputText = Resolve-StartPath $outputText
}

$startTime   = Get-Date
$machineName = [Environment]::MachineName
$isAdmin     = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

#endregion

#region ポートの一覧

if (-not (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue)) {
    Stop-WithError 'Get-NetTCPConnection が見つかりません。このツールは Windows 10 / Windows Server 2016 以降で動きます。'
}

$entries = New-Object 'System.Collections.Generic.List[object]'
try {
    foreach ($c in @(Get-NetTCPConnection -ErrorAction Stop)) {
        $state = "$($c.State)"
        if (-not $Established -and $state -ne 'Listen') { continue }
        if ($hasPortFilter -and -not ($portSet.Contains([int]$c.LocalPort) -or ($state -ne 'Listen' -and $portSet.Contains([int]$c.RemotePort)))) { continue }
        $entries.Add([pscustomobject]@{
            Protocol      = 'TCP'
            LocalAddress  = [string]$c.LocalAddress
            LocalPort     = [int]$c.LocalPort
            RemoteAddress = $(if ($state -eq 'Listen') { '' } else { [string]$c.RemoteAddress })
            RemotePort    = $(if ($state -eq 'Listen') { 0 } else { [int]$c.RemotePort })
            State         = $(if ($state -eq 'Listen') { '待ち受け' } elseif ($state -eq 'Established') { '接続中' } else { $state })
            ProcessId     = [int]$c.OwningProcess
        })
    }
    foreach ($u in @(Get-NetUDPEndpoint -ErrorAction Stop)) {
        if ($hasPortFilter -and -not $portSet.Contains([int]$u.LocalPort)) { continue }
        $entries.Add([pscustomobject]@{
            Protocol = 'UDP'; LocalAddress = [string]$u.LocalAddress; LocalPort = [int]$u.LocalPort
            RemoteAddress = ''; RemotePort = 0; State = '-'; ProcessId = [int]$u.OwningProcess
        })
    }
} catch {
    Stop-WithError ('ポートの一覧を取れません（{0}）。Windows Management Instrumentation のサービスが動いているか確かめてください。' -f (Get-ErrorText (Get-InnerException $_.Exception)))
}

# プロトコル（TCP、UDP の順）、ローカルポートの番号、ローカルアドレス（IPv4、IPv6 の順）、リモートアドレス、リモートポートの順に並べる。
# アドレスの文字列だけで比べると、記号の扱いが 5.1 と 7 で違うことがあるため、IPv4 と IPv6 を先に分ける
$sorted = @($entries | Sort-Object -Property `
    @{ Expression = { if ($_.Protocol -eq 'TCP') { 0 } else { 1 } } }, LocalPort,
    @{ Expression = { if ($_.LocalAddress.Contains(':')) { 1 } else { 0 } } }, LocalAddress, RemoteAddress, RemotePort)

#endregion

#region プロセスとサービスの情報

$pidSet = New-Object 'System.Collections.Generic.HashSet[int]'
foreach ($e in $sorted) { [void]$pidSet.Add($e.ProcessId) }

# PID → 実行ユーザー。Win32_Process の GetOwner は、呼ぶたびにすべてのプロセスを調べ直すため、1 回 0.3 秒ほどかかる。
# Get-Process -IncludeUserName はまとめて取れるので、こちらを先に使う（5.1 では、管理者でないと使えない）
$owners     = @{}
$ownerByWmi = $false
try {
    foreach ($gp in @(Get-Process -IncludeUserName -ErrorAction Stop)) {
        if ($pidSet.Contains($gp.Id) -and $gp.UserName) { $owners[$gp.Id] = [string]$gp.UserName }
    }
} catch {
    $ownerByWmi = $true
}

# PID → プロセスの情報。管理者でないと、ほかのユーザーのプロセスの実行ファイル・コマンドライン・実行ユーザーは取れない
$processes = @{}
try {
    foreach ($p in @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop)) {
        $processId = [int]$p.ProcessId
        if (-not $pidSet.Contains($processId)) { continue }
        $owner = ''
        if ($owners.ContainsKey($processId)) {
            $owner = $owners[$processId]
        } elseif ($ownerByWmi -and $p.ExecutablePath) {
            # 実行ファイルが取れないプロセス（ほかのユーザーのものなど）は、ほとんどが GetOwner も失敗し、
            # 7 の Get-Process -IncludeUserName でも取れない。時間を短くし、5.1 と 7 で結果をそろえるため、呼ばない
            try {
                $result = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction Stop
                if ($result.ReturnValue -eq 0 -and $result.User) { $owner = $(if ($result.Domain) { '{0}\{1}' -f $result.Domain, $result.User } else { $result.User }) }
            } catch { }
        }
        $processes[$processId] = [pscustomobject]@{
            Name        = [string]$p.Name
            Path        = [string]$p.ExecutablePath
            CommandLine = [string]$p.CommandLine
            Started     = $p.CreationDate
            Owner       = $owner
        }
    }
} catch { }

# CIM で取れなかったプロセスは、名前だけでも Get-Process で補う。どちらでも見つからなければ、調べている間に終わったもの
foreach ($processId in $pidSet) {
    if ($processes.ContainsKey($processId)) { continue }
    $name = '（終了したプロセス）'
    $started = $null
    try {
        $gp = Get-Process -Id $processId -ErrorAction Stop
        $name = $gp.ProcessName + '.exe'
        if ($processId -eq 0 -or $processId -eq 4) { $name = $gp.ProcessName }
        try { $started = $gp.StartTime } catch { }
    } catch { }
    $processes[$processId] = [pscustomobject]@{ Name = $name; Path = ''; CommandLine = ''; Started = $started; Owner = '' }
}

# PID → そのプロセスで動いているサービスの名前
$services = @{}
try {
    foreach ($s in @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop)) {
        $processId = [int]$s.ProcessId
        if ($processId -eq 0 -or -not $pidSet.Contains($processId)) { continue }
        if (-not $services.ContainsKey($processId)) { $services[$processId] = New-Object 'System.Collections.Generic.List[string]' }
        $services[$processId].Add([string]$s.Name)
    }
} catch { }
foreach ($key in @($services.Keys)) { $services[$key] = @($services[$key] | Sort-Object) }

# 画面のサービスの列。PID 4 と 0 は、プロセスではないことが分かるように注記する
function Get-ServiceText([int]$ProcessId, [bool]$All) {
    if ($ProcessId -eq 4) { return '（カーネル：HTTP.sys や SMB など）' }
    if ($ProcessId -eq 0) { return '（切断の後始末：プロセスなし）' }
    if (-not $services.ContainsKey($ProcessId)) { return '' }
    $names = @($services[$ProcessId])
    if ($All -or $names.Count -le $ServiceShown) { return $names -join ', ' }
    return '{0} ほか {1} 個' -f (($names[0..($ServiceShown - 1)]) -join ', '), ($names.Count - $ServiceShown)
}

# 管理者でないため取れなかった情報があるか（PID 0 と 4 は、管理者でも取れない項目があるので数えない）
$missingInfo = $false
if (-not $isAdmin) {
    foreach ($processId in $pidSet) {
        if ($processId -le 4) { continue }
        $info = $processes[$processId]
        if ($info.Name -ne '（終了したプロセス）' -and (-not $info.Path -or -not $info.Owner)) { $missingInfo = $true }
    }
}

#endregion

$finished  = $false
$saveError = $null
$runDir    = $null
try {
    #region 画面表示

    $header = New-Object 'System.Collections.Generic.List[string]'
    $header.Add((Format-Invariant 'ポート使用プロセス特定 (Get-PortProcess {0})' $ScriptVersion))
    $header.Add((Format-Invariant '対象: {0}   {1:yyyy-MM-dd HH:mm:ss}   管理者: {2}' $machineName $startTime $(if ($isAdmin) { 'はい' } else { 'いいえ' })))
    if ($configPath) { $header.Add(('設定ファイル: {0}' -f $configPath)) }
    $header.Add(('条件: {0}   {1}' -f $(if ($hasPortFilter) { 'ポート ' + ($portTexts -join ', ') } else { 'すべてのポート' }),
        $(if ($Established) { 'TCP は待ち受けと接続中など' } else { 'TCP は待ち受けのみ' })))
    if ($sorted.Count -gt 0) {
        $header.Add((Format-Invariant '該当 {0:N0} 件（プロセス {1:N0} 個）' $sorted.Count $pidSet.Count))
    } else {
        $header.Add('該当なし')
    }
    if ($missingInfo) { $header.Add('※ 管理者として実行すると、ほかのユーザーのプロセスの実行ユーザー・実行ファイル・コマンドラインも取れます（コマンドラインは CSV にだけ書きます）') }

    $body = New-Object 'System.Collections.Generic.List[string]'
    if ($sorted.Count -gt 0) {
        if ($Established) {
            $head  = @('プロトコル', 'ローカル', 'リモート', '状態', 'PID', 'プロセス', 'サービス')
            $right = @($false, $false, $false, $false, $true, $false, $false)
        } else {
            $head  = @('プロトコル', 'ローカル', '状態', 'PID', 'プロセス', 'サービス')
            $right = @($false, $false, $false, $true, $false, $false)
        }
        $rows = @(
            foreach ($e in $sorted) {
                $cells = @($e.Protocol, (Format-Endpoint $e.LocalAddress $e.LocalPort))
                if ($Established) { $cells += $(if ($e.RemoteAddress) { Format-Endpoint $e.RemoteAddress $e.RemotePort } else { '' }) }
                $cells += @($e.State, [string]$e.ProcessId, $processes[$e.ProcessId].Name, (Get-ServiceText $e.ProcessId $false))
                , $cells
            }
        )
        $body.Add('■ ポート')
        $body.AddRange([string[]](Format-Table $head $right $rows))

        $procRows = @(
            foreach ($processId in @($pidSet | Sort-Object)) {
                $info = $processes[$processId]
                $startedText = if ($info.Started) { ([datetime]$info.Started).ToString('yyyy-MM-dd HH:mm', $Invariant) } else { '' }
                , @([string]$processId, $info.Name, $info.Owner, $startedText, $info.Path)
            }
        )
        $body.Add('■ プロセス')
        $body.AddRange([string[]](Format-Table @('PID', 'プロセス', '実行ユーザー', '起動日時', '実行ファイル') @($true, $false, $false, $false, $false) $procRows))
    }

    foreach ($line in $header) { Write-Host $line }
    foreach ($line in $body) {
        if ($line -eq '') { Write-Host '' }
        elseif ($line.StartsWith('■')) { Write-Host '' ; Write-Host $line -ForegroundColor Cyan }
        else { Write-Host $line }
    }

    #endregion

    #region 保存

    if (-not $NoSave) {
        $baseDir = if ($outputText) { $outputText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
        try {
            $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
            # \\?\ 付きのパスでは末尾のピリオドや空白が残るため、先に通常の書き方で整える（260 文字を超えると例外になるので、そのときはそのまま）
            try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
            $runName = Format-Invariant 'PortProcess_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
            $runDir  = Join-Path $baseDir $runName
            # 失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
            $suffix = 1
            while ([IO.Directory]::Exists((ConvertTo-LongPath $runDir)) -or [IO.File]::Exists((ConvertTo-LongPath $runDir))) {
                $suffix++
                $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix)
            }
            [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $runDir))
        } catch {
            $saveError = ('保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, (Get-ErrorText (Get-InnerException $_.Exception))) + (Get-ConfigNote 'OutputFolder')
            $runDir = $null
        }
        if ($runDir) {
            try {
                $summary = New-Object 'System.Collections.Generic.List[string]'
                $summary.AddRange([string[]]$header)
                foreach ($line in $body) { if ($line.StartsWith('■')) { $summary.Add('') }; $summary.Add($line) }
                [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$summary, $Utf8Bom)

                $csvRows = @(
                    foreach ($e in $sorted) {
                        $info = $processes[$e.ProcessId]
                        [pscustomobject]@{
                            'プロトコル'       = $e.Protocol
                            'ローカルアドレス' = $e.LocalAddress
                            'ローカルポート'   = $e.LocalPort
                            'リモートアドレス' = $e.RemoteAddress
                            'リモートポート'   = $(if ($e.RemoteAddress) { $e.RemotePort } else { '' })
                            '状態'             = $e.State
                            'PID'              = $e.ProcessId
                            'プロセス'         = (ConvertTo-ExcelText $info.Name)
                            'サービス'         = (ConvertTo-ExcelText (Get-ServiceText $e.ProcessId $true))
                            '実行ユーザー'     = (ConvertTo-ExcelText $info.Owner)
                            '起動日時'         = $(if ($info.Started) { ([datetime]$info.Started).ToString('yyyy/MM/dd HH:mm:ss', $Invariant) } else { '' })
                            '実行ファイル'     = (ConvertTo-ExcelText $info.Path)
                            'コマンドライン'   = (ConvertTo-ExcelText $info.CommandLine)
                        }
                    }
                )
                Save-Csv $csvRows $Columns (ConvertTo-LongPath (Join-Path $runDir 'ports.csv'))
            } catch {
                $saveError = '結果を保存できませんでした: {0}' -f (Get-ErrorText (Get-InnerException $_.Exception))
                try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { }
                $runDir = $null
            }
        }
    }

    #endregion

    $finished = $true
} finally {
    if (-not $finished) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow
    }
}

if ($saveError) {
    Write-Host ''
    Write-Warning $saveError
} elseif ($runDir) {
    Write-Host ''
    Write-Host ('保存先: {0}' -f $runDir)
}

exit 0
