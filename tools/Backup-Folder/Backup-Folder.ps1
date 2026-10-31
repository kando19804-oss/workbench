<#
.SYNOPSIS
    フォルダの中身を、USB メモリや別のドライブへ差分でコピーします。

.DESCRIPTION
    コピー元のフォルダを調べ、コピーするファイルの件数と容量を表示してから、確認のうえでコピーします。
    2 回目からは、サイズか更新日時が変わったファイルだけをコピーします。コピーには robocopy を使います。

    コピー元のファイルを変更・削除することはありません。コピー先のファイルを削除することもありません。
    結果（コピーしたファイルの一覧と robocopy の記録）は、CSV（BOM 付き UTF-8）などで保存します。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Backup-Folder.ps1 -Source D:\data -Destination E:\backup\data

.PARAMETER Source
    コピー元のフォルダ。ドライブ（D:\）や共有フォルダ（\\server\share）も指定できます。

.PARAMETER Destination
    コピー先のフォルダ。なければ作ります。

.PARAMETER ListOnly
    コピーする予定のファイルの件数と容量を表示し、一覧を files.csv に保存するだけで、コピーしません。
    コピー先のフォルダも作りません。

.PARAMETER Force
    確認せずにコピーします。タスクスケジューラなど、画面を見ずに実行するときに使います。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER ConfigFile
    設定ファイルのパス。Source・Destination・OutputFolder を「項目 = 値」の形で書いておけます。
    コマンドラインで指定した値は、設定ファイルより優先します。書き方は README を参照してください。

.PARAMETER CreateLauncher
    スクリプトと同じフォルダに、ダブルクリック用の起動ファイル（.cmd）と設定ファイルのひな形（.ini）を作ります。
    コピーはしません。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Backup-Folder.ps1 -Source \\server01\share\経理 -Destination E:\backup\経理

    共有フォルダの中身を、USB メモリへコピーします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Backup-Folder.ps1 -Source D:\data -Destination E:\backup\data -ListOnly

    コピーする予定のファイルを確かめるだけで、コピーはしません。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Backup-Folder.ps1 -CreateLauncher

    ダブルクリック用の起動ファイルと、設定ファイルのひな形を作ります。

.NOTES
    Backup-Folder 1.0.0
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
    [string]$Source,

    [Parameter(Position = 1)]
    [string]$Destination,

    [switch]$ListOnly,

    [switch]$Force,

    [string]$OutputFolder,

    [string]$ConfigFile,

    [switch]$CreateLauncher
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$RobocopyExe   = Join-Path ([Environment]::SystemDirectory) 'Robocopy.exe'
$Fat32Limit    = 4GB   # FAT32 に置けるのは 4 GB 未満のファイルだけ
$TimeTolerance = 2     # 更新日時の違いを同じとみなす秒数（robocopy の /FFT と同じ）
# コピーしないファイル（Windows が作るサムネイルの記録とフォルダの設定、Office が編集中に作る一時ファイル）
$ExcludeFiles  = 'Thumbs.db', 'desktop.ini', '~$*'

#region 補助関数

# コマンドプロンプトから渡された引数に残る引用符を外す
# （"C:\" が C:" として届く問題や、cmd では ' がそのまま届く問題への対策）
function Remove-Quote([string]$Text) {
    $t = $Text.Trim().Trim('"')
    if ($t.Length -ge 2 -and $t.StartsWith("'") -and $t.EndsWith("'")) { $t = $t.Substring(1, $t.Length - 2) }
    return $t
}

# 5.1 は 260 文字を超えるパスを通常の書き方では読めないため、内部では \\?\ 付きのパスを使う
function ConvertTo-LongPath([string]$FullPath) {
    if ($FullPath.StartsWith('\\?\')) { return $FullPath }
    if ($FullPath.StartsWith('\\')) { return '\\?\UNC\' + $FullPath.Substring(2) }
    return '\\?\' + $FullPath
}

# 表示・保存用に \\?\ を外す
function ConvertTo-DisplayPath([string]$FullPath) {
    if ($FullPath.StartsWith('\\?\UNC\')) { return '\\' + $FullPath.Substring(8) }
    if ($FullPath.StartsWith('\\?\')) { return $FullPath.Substring(4) }
    return $FullPath
}

function Format-Size([long]$Bytes) {
    $units = 'B', 'KB', 'MB', 'GB', 'TB', 'PB'
    $value = [double]$Bytes
    $i = 0
    while ($value -ge 1024 -and $i -lt $units.Count - 1) { $value /= 1024; $i++ }
    if ($i -eq 0) { return '{0} B' -f $Bytes }
    return [string]::Format($Invariant, '{0:N1} {1}', $value, $units[$i])
}

# 日付や数値の区切り文字が、OS の地域設定によって変わらないようにする
function Format-Invariant([string]$Format) {
    return [string]::Format($Invariant, $Format, [object[]]$args)
}

function Format-Duration([TimeSpan]$Span) {
    if ($Span.TotalSeconds -lt 60) { return [string]::Format($Invariant, '{0:N1}秒', $Span.TotalSeconds) }
    if ($Span.TotalHours -lt 1) { return '{0}分{1}秒' -f $Span.Minutes, $Span.Seconds }
    return '{0}時間{1}分{2}秒' -f [math]::Floor($Span.TotalHours), $Span.Minutes, $Span.Seconds
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
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -Source "C:\data"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# 入力されたパスをフルパスにする。末尾の \ は、ドライブのルート以外は外す
# （robocopy は "D:\data\" の最後の \" を引用符として読んでしまうため、末尾に \ を残さない）
function Resolve-FullPath([string]$Text) {
    $full = ConvertTo-DisplayPath $Text
    try {
        $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($full)
        $full = [IO.Path]::GetFullPath($full)
    } catch { }
    if ($full -notmatch '^[A-Za-z]:\\$') { $full = $full.TrimEnd('\') }
    return $full
}

# コピー先のドライブの形式と空き容量。共有フォルダは DriveInfo が対応していないため、FileSystemObject で取得する。
# 取得できないときは $null を返す
function Get-DriveSpace([string]$FullPath) {
    try {
        if ($FullPath.StartsWith('\\')) {
            $fso   = New-Object -ComObject Scripting.FileSystemObject
            $drive = $fso.GetDrive($fso.GetDriveName($FullPath))
            return [pscustomobject]@{ Format = [string]$drive.FileSystem; Free = [long]$drive.FreeSpace }
        }
        $drive = New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($FullPath))
        if ($drive.IsReady) { return [pscustomobject]@{ Format = $drive.DriveFormat; Free = $drive.AvailableFreeSpace } }
    } catch { }
    return $null
}

# 共有フォルダやネットワークドライブを開けないときの案内。管理者として実行しているときは、そのことも添える
function Get-NetworkHint([string]$FullPath) {
    $isNetwork = $FullPath.StartsWith('\\')
    if (-not $isNetwork -and $FullPath -match '^[A-Za-z]:') {
        try { $isNetwork = (New-Object IO.DriveInfo ($FullPath.Substring(0, 2))).DriveType -eq 'Network' } catch { }
        # 管理者として実行していると、通常のユーザーで割り当てたネットワークドライブはドライブごと見えない
        if (-not [IO.Directory]::Exists($FullPath.Substring(0, 2) + '\')) { $isNetwork = $isNetwork -or $isAdmin }
    }
    $hint = @()
    if ($isNetwork) {
        $hint += '共有フォルダのときは、次の順に確かめてください。'
        $hint += '  1. エクスプローラーのアドレス欄にパスを貼り付けて、開けるか試す'
        $hint += '  2. ユーザー名とパスワードを聞かれたら入力する（「資格情報を記憶する」にチェックを入れると、次からは聞かれません）'
        $hint += '  3. 開けたら、もう一度実行する'
        if ($isAdmin) {
            $hint += '※ 管理者として実行しています。通常のユーザーで接続した共有フォルダやネットワークドライブは、管理者として開いた PowerShell からは見えないことがあります。管理者としてではなく実行してください。'
        }
    }
    return $hint
}

# robocopy の記録（UTF-16）を読む。ReadAllText は壊れた文字（対になっていないサロゲート）を置き換えてしまい、
# ファイル名が一致しなくなるため、バイト列をそのまま文字に移す
function Read-UnicodeText([string]$FilePath) {
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    $start = if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { 2 } else { 0 }
    $chars = New-Object char[] ([int][math]::Floor(($bytes.Length - $start) / 2))
    [Buffer]::BlockCopy($bytes, $start, $chars, 0, $chars.Length * 2)
    return New-Object string (, $chars)
}

# ファイルがコピー先で使う容量。クラスタ（割り当ての単位）の大きさに切り上げる。
# exFAT の USB メモリはクラスタが大きく、小さなファイルが多いと、ファイルの合計よりかなり多くの容量を使うため
function Get-AllocatedSize([long]$Bytes) {
    if ($clusterSize -le 0 -or $Bytes -le 0) { return $Bytes }
    return [long][math]::Ceiling($Bytes / [double]$clusterSize) * $clusterSize
}

# サイズが同じで、更新日時の違いが 2 秒以内なら、同じファイルとみなす（更新日時が分からないときは、サイズだけで判定する）
function Test-SameFile([IO.FileInfo]$Info, [long]$Size, $TimeUtc) {
    if ($Info.Length -ne $Size) { return $false }
    if ($null -eq $TimeUtc) { return $true }
    return [math]::Abs(($Info.LastWriteTimeUtc - $TimeUtc).TotalSeconds) -le $TimeTolerance
}

# フォルダから属性を外す。外せなくても（フォルダがないときも）、コピーの結果には関わらないため、知らせずに続ける
function Clear-FolderAttribute([string]$FolderPath, [int]$Mask) {
    if ($Mask -eq 0) { return }
    try {
        $current = [int][IO.File]::GetAttributes($FolderPath)
        if (($current -band $Mask) -ne 0) { [IO.File]::SetAttributes($FolderPath, [IO.FileAttributes]($current -band -bnot $Mask)) }
    } catch { }
}

function Write-Line([string]$Text) {
    $lines.Add($Text)
    Write-Host $Text
}

#endregion

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

#region 起動ファイルの作成

# 実行ポリシーにかかわらず起動できるよう、-ExecutionPolicy Bypass を付けて起動するバッチを作る。
# %~dpn0 は「バッチのあるフォルダ＋バッチの名前（拡張子なし）」なので、フォルダごと移しても（USB メモリのドライブ文字が
# 変わっても）動く。中身が ASCII だけになるため、cmd.exe がどのコードページで読んでも壊れない
if ($CreateLauncher) {
    $baseName  = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
    $launchers = @(
        [pscustomobject]@{
            Path     = Join-Path $PSScriptRoot "$baseName.cmd"
            Label    = '起動ファイル'
            Lines    = @(
                '@echo off'
                'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dpn0.ps1" -ConfigFile "%~dpn0.ini"'
                'pause'
            )
            Encoding = [Text.Encoding]::ASCII
        },
        [pscustomobject]@{
            Path     = Join-Path $PSScriptRoot "$baseName.ini"
            Label    = '設定ファイルのひな形'
            Lines    = @(
                '# Backup-Folder の設定ファイル'
                '# 「項目 = 値」の形で書きます。# か ; で始まる行はコメントです'
                ''
                '# コピー元のフォルダ（例: \\server01\共有\経理）'
                'Source ='
                ''
                '# コピー先のフォルダ（例: E:\backup\経理）'
                'Destination ='
                ''
                '# 結果の保存先（書かなければ、スクリプトと同じフォルダに保存します）'
                'OutputFolder ='
            )
            Encoding = $Utf8Bom
        }
    )
    foreach ($launcher in $launchers) {
        if (Test-Path -LiteralPath $launcher.Path) {
            Write-Host ('{0}はすでにあるため、作りませんでした: {1}' -f $launcher.Label, $launcher.Path)
            continue
        }
        try {
            [IO.File]::WriteAllLines($launcher.Path, [string[]]$launcher.Lines, $launcher.Encoding)
        } catch {
            Stop-WithError ('{0}を作れません: {1}（{2}）' -f $launcher.Label, $launcher.Path, (Get-ErrorText (Get-InnerException $_.Exception)))
        }
        Write-Host ('{0}を作りました: {1}' -f $launcher.Label, $launcher.Path)
    }
    Write-Host ''
    Write-Host ('次に、{0}.ini をメモ帳で開いて、Source（コピー元）と Destination（コピー先）を書いてください。' -f $baseName)
    Write-Host ('その後は、{0}.cmd をダブルクリックするとバックアップできます。' -f $baseName)
    exit 0
}

#endregion

#region 設定ファイル

# 書ける項目は、パラメーターの名前から - を除いたもの。-ListOnly と -Force は、確認なしでコピーする設定が
# 知らないうちに効くのを防ぐため、書けないようにする
$configItems   = 'Source', 'Destination', 'OutputFolder'
$configPath    = $null
$configApplied = @{}   # 設定ファイルの値を使った項目 → 行番号

if ($ConfigFile) {
    $configText = Remove-Quote $ConfigFile
    if ($configText.Contains('"')) { Stop-WithQuoteError '-ConfigFile' $configText }
    $configText = Resolve-StartPath $configText
    $configPath = $configText
    try {
        $configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((ConvertTo-DisplayPath $configText))
        $configPath = [IO.Path]::GetFullPath($configPath)
    } catch { }
    if (-not [IO.File]::Exists((ConvertTo-LongPath $configPath))) {
        # エクスプローラーで拡張子を表示していないと、メモ帳で usb.ini.txt として保存してしまうことがある
        $message = '設定ファイルが見つかりません: {0}' -f $configPath
        if ([IO.File]::Exists((ConvertTo-LongPath ($configPath + '.txt')))) {
            $message += "`r`n" + ('拡張子が .txt になっていませんか: {0}.txt' -f $configPath)
        }
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
    $lineNumber   = 0
    foreach ($rawLine in $configLines) {
        $lineNumber++
        $lineText = $rawLine.Trim()
        # # と ; はフォルダ名に使える文字なので、行の先頭にあるときだけコメントとみなす
        if ($lineText -eq '' -or $lineText.StartsWith('#') -or $lineText.StartsWith(';')) { continue }
        $separator = $lineText.IndexOf('=')
        if ($separator -le 0) { Stop-WithConfigError $lineNumber ('「項目 = 値」の形になっていません: {0}' -f $lineText) }

        $itemKey  = $lineText.Substring(0, $separator).Trim()
        $itemName = $configItems | Where-Object { $_ -eq $itemKey }
        if (-not $itemName) { Stop-WithConfigError $lineNumber ('知らない項目です: {0}' -f $itemKey) ('使える項目: {0}' -f ($configItems -join ', ')) }

        $itemValue = Remove-ConfigQuote $lineText.Substring($separator + 1)
        # 値が空の行は、書かなかったのと同じ扱い（ひな形の空の行の下に書き足しても、2 回書いたことにしない）
        if ($itemValue -eq '') { continue }

        # どちらの値が効いているか分からなくなるため、同じ項目は 1 回だけ
        if ($configLineOf.ContainsKey($itemName)) {
            Stop-WithConfigError $lineNumber ('{0} が 2 回書かれています（{1} 行目と {2} 行目）' -f $itemName, $configLineOf[$itemName], $lineNumber)
        }
        $configLineOf[$itemName] = $lineNumber

        if ($itemValue.Contains('"')) { Stop-WithConfigError $lineNumber ('{0} の値に引用符（"）が残っています: {1}' -f $itemName, $itemValue) }
        # コマンドプロンプトで書くときと同じように、%TEMP% などの環境変数を展開する
        $itemValue = Expand-ConfigVariable $lineNumber $itemValue
        # 相対パスは、設定ファイルのあるフォルダを基準にする（タスクスケジューラでは現在のフォルダが System32 になるため）
        if (-not (Test-RootedPath $itemValue)) {
            $itemValue = $configDir.TrimEnd('\') + '\' + $itemValue
            try { $itemValue = [IO.Path]::GetFullPath($itemValue) } catch { }
        }
        $configValues[$itemName] = $itemValue
    }

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

# コマンドプロンプトで -Source "C:\data\" と書くと、後ろの引数まで -Source の値に入り、-Destination が空になる。
# 本当の原因（最後の \"）を知らせるため、空かどうかより先に引用符を確かめる
$srcText = Remove-Quote $Source
if ($srcText.Contains('"')) { Stop-WithQuoteError '-Source' $srcText }
$dstText = Remove-Quote $Destination
if ($dstText.Contains('"')) { Stop-WithQuoteError '-Destination' $dstText }
$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    $outText = Resolve-StartPath $outText
}

# 設定ファイルにも書けるため、どちらで指定するかを案内する
if (-not $srcText) { Stop-WithError 'コピー元を指定してください（-Source で指定するか、設定ファイルの Source に書きます）' }
if (-not $dstText) { Stop-WithError 'コピー先を指定してください（-Destination で指定するか、設定ファイルの Destination に書きます）' }
if ($srcText -match '^(\\\\\?\\)?[A-Za-z]:$') { $srcText += '\' }
if ($dstText -match '^(\\\\\?\\)?[A-Za-z]:$') { $dstText += '\' }
$srcPath = Resolve-FullPath (Resolve-StartPath $srcText)
$dstPath = Resolve-FullPath (Resolve-StartPath $dstText)

# コピー元は、中を読めるところまで確かめる（共有フォルダは、接続していないと読めない）
$srcError = $null
$srcOk    = $false
try { $srcOk = [IO.Directory]::Exists((ConvertTo-LongPath $srcPath)) } catch { }
if ($srcOk) {
    try {
        [void][IO.Directory]::EnumerateFileSystemEntries((ConvertTo-LongPath $srcPath)).GetEnumerator().MoveNext()
    } catch {
        $srcOk    = $false
        $srcError = Get-ErrorText (Get-InnerException $_.Exception)
    }
}
if (-not $srcOk) {
    $message = 'コピー元のフォルダを開けません: {0}' -f $srcPath
    if ($srcError) { $message += '（{0}）' -f $srcError }
    $message += Get-ConfigNote 'Source'
    Stop-WithError ((@($message) + (Get-NetworkHint $srcPath)) -join "`r`n")
}

if ($dstPath -match '^[A-Za-z]:' -and -not [IO.Directory]::Exists($dstPath.Substring(0, 2) + '\')) {
    $message = ('コピー先のドライブ {0} が見つかりません。USB メモリなどがつながっているか確かめてください' -f $dstPath.Substring(0, 2)) + (Get-ConfigNote 'Destination')
    Stop-WithError ((@($message) + (Get-NetworkHint $dstPath)) -join "`r`n")
}
# robocopy に任せると「コピー先のフォルダにアクセスできない」という分かりにくいエラーになるため、先に確かめる
if ([IO.File]::Exists((ConvertTo-LongPath $dstPath))) {
    Stop-WithError (('コピー先のフォルダを作れません: {0}（同じ名前のファイルがあります）' -f $dstPath) + (Get-ConfigNote 'Destination'))
}

# コピー先がコピー元の中にあると、実行するたびにコピー先の中身もコピーされて増えていく。
# ネットワークドライブと UNC パスのように、同じ場所を別の名前で指定した場合は判定できない
$srcPrefix = $srcPath.TrimEnd('\') + '\'
$dstPrefix = $dstPath.TrimEnd('\') + '\'
if ($dstPrefix -eq $srcPrefix) { Stop-WithError ('コピー元とコピー先が同じです: {0}' -f $srcPath) }
if ($dstPrefix.StartsWith($srcPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    Stop-WithError ('コピー先がコピー元の中にあります。コピーを繰り返すたびに、コピー先の中身もコピーされてしまいます。' + "`r`n" + ('コピー元: {0}' -f $srcPath) + "`r`n" + ('コピー先: {0}' -f $dstPath))
}

if (-not (Test-Path -LiteralPath $RobocopyExe)) { Stop-WithError ('robocopy が見つかりません: {0}' -f $RobocopyExe) }

$startTime = Get-Date

#endregion


# robocopy の記録（/UNILOG）は、いったん一時フォルダに書く。途中で止まったときも、終わったときも消す
$planLog     = Join-Path ([IO.Path]::GetTempPath()) ('Backup-Folder_{0}.log' -f [Guid]::NewGuid().ToString('N'))
$copyLog     = Join-Path ([IO.Path]::GetTempPath()) ('Backup-Folder_{0}.log' -f [Guid]::NewGuid().ToString('N'))
$runDir      = $null
$exitCode    = 0
$copyStarted = $false
$completed   = $false
$hideAttributes = 0
try {
    #region コピーするファイルを調べる

    # robocopy を一覧だけのモード（/L）で動かし、コピーする予定のファイルを記録から読む。
    # 記録は UTF-16 なので、どの言語のファイル名も壊れない。/NC で種類の列を空にし、/BYTES でサイズをバイトで、
    # /TS でコピー元の更新日時（UTC）を出す（共有フォルダのファイルを 1 件ずつ調べ直すと遅いため）
    $commonArgs = @($srcPath, $dstPath, '/E', '/COPY:DAT', '/DCOPY:T', '/FFT', '/XJ', '/XX', '/XF') + $ExcludeFiles
    # ドライブのルートには、ごみ箱（$RECYCLE.BIN）と、標準ユーザーでは読めない System Volume Information がある。
    # ドライブのルートを共有したフォルダ（\\server\D$ など）はパスの形では分からないため、いつもコピー元の直下のものを除く。
    # 下の階層にある同じ名前のフォルダを巻き込まないよう、フルパスで指定する
    $commonArgs += @('/XD', ($srcPrefix + '$RECYCLE.BIN'), ($srcPrefix + 'System Volume Information'))
    $commonArgs += @('/NP', '/FP')
    $plan    = New-Object 'System.Collections.Generic.List[object]'
    $skipped = New-Object 'System.Collections.Generic.List[object]'
    $newDirs = 0
    # \\?\ 付きのパスを 1 件ずつ関数で作ると遅いため、先頭の部分を先に作っておく
    $srcLong = ConvertTo-LongPath $srcPrefix
    $dstLong = ConvertTo-LongPath $dstPrefix
    try {
        Write-Progress -Activity 'フォルダのバックアップ' -Status 'コピーするファイルを調べています'
        $null = & $RobocopyExe @commonArgs /L /NC /BYTES /TS /NJH /NJS /R:0 /W:0 "/UNILOG:$planLog" 2>&1
        $listExit = $LASTEXITCODE
        if ($listExit -ge 16) {
            $detail = ''
            if (Test-Path -LiteralPath $planLog) { $detail = (((Read-UnicodeText $planLog) -split '\r?\n') | Where-Object { $_.Trim() } | Select-Object -Last 3) -join ' ' }
            Stop-WithError ('robocopy でコピー元を調べられませんでした（終了コード {0}）: {1}' -f $listExit, $detail)
        }

        # 行の形: フォルダは「タブ 件数 タブ パス\」、ファイルは「タブ (種類) タブ タブ サイズ 空白 更新日時 タブ パス」。
        # エラーの説明など、ほかの形の行は飛ばす
        foreach ($line in ((Read-UnicodeText $planLog) -split '\r?\n')) {
            $cols = $line.Split("`t")
            if ($cols.Count -lt 3) { continue }
            $itemPath = $cols[$cols.Count - 1]
            if (-not $itemPath.StartsWith($srcPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
            $relative = $itemPath.Substring($srcPrefix.Length)
            if ($itemPath.EndsWith('\')) {
                # 読めないフォルダは「ファイル 0 件のフォルダ」として記録されるだけで、エラーの行は出ない。
                # 件数 0 のフォルダは開けるか試し、読めなければ、飛ばした場所として残す
                if ($cols[$cols.Count - 2].Trim() -eq '0') {
                    try {
                        [void][IO.Directory]::EnumerateFileSystemEntries($srcLong + $relative).GetEnumerator().MoveNext()
                    } catch {
                        $ex     = Get-InnerException $_.Exception
                        $reason = if ($ex -is [UnauthorizedAccessException]) { '権限なし' } else { 'その他のエラー' }
                        $skipped.Add([pscustomobject]@{ '理由' = $reason; 'パス' = $itemPath.TrimEnd('\'); '詳細' = (Get-ErrorText $ex) })
                        continue
                    }
                }
                if (-not [IO.Directory]::Exists($dstLong + $relative)) { $newDirs++ }
                continue
            }
            $sizeCol  = $cols[$cols.Count - 2].Trim()
            $gap      = $sizeCol.IndexOf(' ')
            $sizeText = if ($gap -gt 0) { $sizeCol.Substring(0, $gap) } else { $sizeCol }
            $size     = 0L
            if (-not [long]::TryParse($sizeText, [ref]$size)) { continue }
            $timeUtc = $null
            $parsed  = [datetime]::MinValue
            if ($gap -gt 0 -and [datetime]::TryParseExact($sizeCol.Substring($gap + 1).Trim(), 'yyyy/MM/dd HH:mm:ss', $Invariant,
                    [Globalization.DateTimeStyles]'AssumeUniversal, AdjustToUniversal', [ref]$parsed)) {
                $timeUtc = $parsed
            }
            $dstInfo = [IO.FileInfo]::new($dstLong + $relative)
            $exists  = $dstInfo.Exists
            $plan.Add([pscustomobject]@{
                Kind     = if ($exists) { '上書き' } else { '新規' }
                Path     = $itemPath
                Relative = $relative
                Size     = $size
                OldSize  = if ($exists) { $dstInfo.Length } else { 0L }
                TimeUtc  = $timeUtc
                Result   = ''
            })
        }
    } finally {
        Write-Progress -Activity 'フォルダのバックアップ' -Completed
    }
    # 終了コードの 8 は、読めなかったフォルダかファイルがあったことを示す
    $listErrors = ($listExit -band 8) -ne 0

    $space   = Get-DriveSpace $dstPath
    $isFat32 = $space -and $space.Format -match '^FAT(32)?$'
    $clusterSize = 0L
    if ($dstPath -match '^[A-Za-z]:') {
        try { $clusterSize = [long]@(Get-CimInstance Win32_Volume -Filter ("DriveLetter='{0}'" -f $dstPath.Substring(0, 2)) -ErrorAction Stop)[0].BlockSize } catch { }
    }
    $planBytes = 0L
    $needBytes = 0L
    $overCount = 0
    $bigCount  = 0
    foreach ($item in $plan) {
        $planBytes += $item.Size
        if ($item.Kind -eq '上書き') { $overCount++ }
        if ($item.Size -ge $Fat32Limit) {
            $bigCount++
            if ($isFat32) { continue }   # FAT32 にはコピーできないので、空きの計算に入れない
        }
        # 上書きするファイルは、元の大きさとの差だけ空きが要る
        $newSize = Get-AllocatedSize $item.Size
        $oldSize = Get-AllocatedSize $item.OldSize
        if ($newSize -gt $oldSize) { $needBytes += $newSize - $oldSize }
    }

    #endregion

    #region 確認

    $lines = New-Object 'System.Collections.Generic.List[string]'
    Write-Line (Format-Invariant 'フォルダのバックアップ (Backup-Folder {0})' $ScriptVersion)
    Write-Line ('コピー元: {0}' -f $srcPath)
    $driveText = if ($space) { Format-Invariant '（{0}・空き {1}）' $space.Format (Format-Size $space.Free) } else { '' }
    Write-Line ('コピー先: {0}{1}' -f $dstPath, $driveText)
    Write-Line (Format-Invariant '{0:yyyy-MM-dd HH:mm:ss}   {1}（管理者: {2}）' $startTime $env:COMPUTERNAME $(if ($isAdmin) { 'はい' } else { 'いいえ' }))
    if ($configPath) { Write-Line ('設定ファイル: {0}' -f $configPath) }
    Write-Line ''

    $hasWork = $plan.Count -gt 0 -or $newDirs -gt 0
    if ($plan.Count -gt 0) {
        Write-Line (Format-Invariant 'コピーするファイル {0:N0} 件（{1}）  新規 {2:N0} 件・上書き {3:N0} 件' $plan.Count (Format-Size $planBytes) ($plan.Count - $overCount) $overCount)
    } elseif ($newDirs -gt 0) {
        Write-Line (Format-Invariant 'コピーするファイルはありません。新しく作るフォルダ {0:N0} 個' $newDirs)
    } else {
        Write-Line 'コピーするファイルはありません（すべてコピー済みです）'
    }
    if ($skipped.Count -gt 0) {
        Write-Line (Format-Invariant '※ 読めなかったフォルダが {0:N0} 件あります。そのフォルダの中はコピーされません（skipped.csv を確認してください）' $skipped.Count)
    } elseif ($listErrors) {
        Write-Line '※ 読めなかったフォルダかファイルがあります（robocopy.log を確認してください）'
    }
    # 結果の保存先がコピー元の中にあると、次からは結果のフォルダもコピーされる
    $baseDir = if ($outText) { $outText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    $baseDir = Resolve-FullPath $baseDir
    if (($baseDir.TrimEnd('\') + '\').StartsWith($srcPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        Write-Line '※ 結果の保存先がコピー元の中にあるため、結果のフォルダもコピーされます（-OutputFolder で、コピー元の外を指定してください）'
    }
    if ($isFat32 -and $bigCount -gt 0) { Write-Line (Format-Invariant '※ コピー先は FAT32 のため、4 GB 以上のファイル {0:N0} 件はコピーできません' $bigCount) }
    if ($space -and $needBytes -gt $space.Free) {
        $shortageText = 'コピー先の空きが足りません（必要 {0}・空き {1}）' -f (Format-Size $needBytes), (Format-Size $space.Free)
        # 一覧だけのときはコピーしないので、知らせるだけにする
        if ($ListOnly) { Write-Line ('※ ' + $shortageText) } else { Stop-WithError $shortageText }
    }

    $doCopy = $hasWork -and -not $ListOnly
    if ($doCopy -and -not $Force) {
        # 画面のない実行（タスクスケジューラの「ログオンしているかどうかにかかわらず実行する」など）で、確認待ちのまま止まり続けないようにする
        $cannotAsk = '確認できないため、コピーしません。確認せずにコピーするときは -Force を付けてください'
        if (-not [Environment]::UserInteractive) { Stop-WithError $cannotAsk }
        $answer = ''
        while ($answer -notmatch '^[YyNnＹｙＮｎ]') {
            # -NonInteractive で起動されていると、Read-Host は英語のエラーで止まる
            try { $answer = Read-Host 'コピーしますか？ [Y] はい  [N] いいえ' } catch { Stop-WithError $cannotAsk }
            if ($null -eq $answer) { Stop-WithError $cannotAsk }   # 入力が終わっている（リダイレクトされた入力の終わり）
            $answer = $answer.Trim()
        }
        if ($answer -match '^[NnＮｎ]') {
            Write-Host 'コピーしませんでした'
            exit 0
        }
    }

    # robocopy は、コピー元のフォルダの属性をコピー先のフォルダにも付ける。ドライブのルートには隠し・システム属性が
    # 付いているため、そのままだとコピー先のフォルダがエクスプローラーに表示されなくなる。コピー元に付いている
    # 隠し・システム属性を、robocopy の後と最後に（中断したときも）コピー先から外す。コピー前の属性に戻す方法だと、
    # 前の実行を途中で止めて残った属性まで覚えてしまうため、外す属性はコピー元で決める。前の実行が強制終了されて
    # 属性が残ったときも直るよう、コピーするものがない日も外す
    if (-not $ListOnly -and $dstPath -notmatch '^[A-Za-z]:\\$') {
        try { $hideAttributes = [int]([IO.File]::GetAttributes((ConvertTo-LongPath $srcPath)) -band [IO.FileAttributes]'Hidden, System') } catch { }
    }

    # robocopy に任せると、作れなかったことが分かりにくいため、コピーする前に自分で作っておく
    if ($doCopy) {
        try {
            [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $dstPath))
        } catch {
            Stop-WithError (('コピー先のフォルダを作れません: {0}（{1}）' -f $dstPath, (Get-ErrorText (Get-InnerException $_.Exception))) + (Get-ConfigNote 'Destination'))
        }
    }

    #endregion

    #region 保存先の準備

    # コピーが終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく。
    # ここから後で中断したときは、作った保存先のフォルダを消す
    $copyStarted = $true
    $saveWarning = $null
    try {
        $runName = Format-Invariant 'Backup_{0}_{1:yyyyMMdd_HHmmss}' $env:COMPUTERNAME $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while (Test-Path -LiteralPath $runDir) { $suffix++; $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix) }
        [void][IO.Directory]::CreateDirectory($runDir)
    } catch {
        $saveWarning = ('保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, (Get-ErrorText (Get-InnerException $_.Exception))) + (Get-ConfigNote 'OutputFolder')
        $runDir = $null
        Write-Warning $saveWarning
    }

    #endregion

    #region コピー

    $resultLog = $planLog
    if ($doCopy) {
        Write-Host ''
        $sw = [Diagnostics.Stopwatch]::StartNew()
        # robocopy の表示は、コピー中のファイルを知らせるために画面にだけ流す（出力として返さない）。
        # 一時的な記録の場所を示す行は、利用者には意味がないため出さない。robocopy の表示は、コンソールのコードページが
        # 932 なら日本語（ログ ファイル:）、65001 なら英語（Log File :）になるため、記録のパスを含むかどうかで判定する
        & $RobocopyExe @commonArgs /MT:8 /R:3 /W:5 /NJH /NJS /NDL /TEE "/UNILOG:$copyLog" 2>&1 | ForEach-Object {
            $text = "$_"
            if (-not $text.Contains($copyLog) -and $text -notmatch '^\s*(Log File|ログ ファイル)\s*:') { Write-Host $text }
        }
        $copyExit  = $LASTEXITCODE
        $elapsed   = $sw.Elapsed
        $resultLog = $copyLog
        Clear-FolderAttribute (ConvertTo-LongPath $dstPath) $hideAttributes

        # robocopy の表示に頼らず、コピー先のファイルのサイズと更新日時を見て、1 件ずつ確かめる。
        # バックアップの途中でコピー元が更新されたときは、計画を立てた時点ではなく、今のコピー元と比べる
        $copied = 0
        foreach ($item in $plan) {
            $ok = $false
            try {
                $info = [IO.FileInfo]::new($dstLong + $item.Relative)
                if ($info.Exists) {
                    $ok = Test-SameFile $info $item.Size $item.TimeUtc
                    if (-not $ok) {
                        # robocopy は、コピー元を調べた時点の更新日時をコピー先に付ける。そのため、途中で更新されたファイルは
                        # 「今の内容で、調べた時点の更新日時」になる。今のコピー元と、どちらの更新日時で比べてもよいことにする
                        $now = [IO.FileInfo]::new($srcLong + $item.Relative)
                        if ($now.Exists) { $ok = (Test-SameFile $info $now.Length $now.LastWriteTimeUtc) -or (Test-SameFile $info $now.Length $item.TimeUtc) }
                    }
                }
            } catch { }
            if ($ok) { $item.Result = 'コピーした'; $copied++ } else { $item.Result = '失敗' }
        }
        $failed = $plan.Count - $copied

        Write-Host ''
        Write-Line ''
        if ($copyExit -ge 16) {
            Write-Line ('結果: robocopy が途中で止まりました（終了コード {0}）。robocopy.log を確認してください' -f $copyExit)
            $exitCode = 1
        } elseif ($failed -gt 0) {
            Write-Line (Format-Invariant '結果: コピーできなかったファイルが {0:N0} 件あります。files.csv と robocopy.log を確認してください（コピー {1:N0} 件  所要 {2}）' $failed $copied (Format-Duration $elapsed))
            $exitCode = 1
        } elseif ($skipped.Count -gt 0) {
            Write-Line (Format-Invariant '結果: 読めなかったフォルダが {0:N0} 件あります。そのフォルダの中はコピーしていません。skipped.csv を確認してください（コピー {1:N0} 件  所要 {2}）' $skipped.Count $copied (Format-Duration $elapsed))
            $exitCode = 1
        } elseif ($copyExit -ge 8) {
            Write-Line (Format-Invariant '結果: コピーできなかったフォルダかファイルがあります。robocopy.log を確認してください（コピー {0:N0} 件  所要 {1}）' $copied (Format-Duration $elapsed))
            $exitCode = 1
        } else {
            Write-Line (Format-Invariant '結果: 正常に終わりました  コピー {0:N0} 件  失敗 0 件  所要 {1}' $copied (Format-Duration $elapsed))
        }
    } else {
        foreach ($item in $plan) { $item.Result = '一覧のみ' }
        if ($ListOnly -and $hasWork) {
            Write-Line ''
            Write-Line '一覧だけを保存しました。コピーはしていません'
        }
        # コピーするものがない日も、読めなかったフォルダがあることを終了コードで分かるようにする（-ListOnly のときは 0）
        if (-not $ListOnly -and ($skipped.Count -gt 0 -or $listErrors)) { $exitCode = 1 }
    }

    #endregion

    #region 保存

    if ($runDir) {
        try {
            [IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), [string[]]$lines, $Utf8Bom)
            $rows = @(
                foreach ($item in $plan) {
                    $timeText = if ($item.TimeUtc) { $item.TimeUtc.ToLocalTime().ToString('yyyy/MM/dd HH:mm:ss', $Invariant) } else { '' }
                    [pscustomobject]@{ '種類' = $item.Kind; '結果' = $item.Result; 'ファイル' = $item.Path; 'サイズ(MB)' = ($item.Size / 1MB).ToString('N2', $Invariant); 'サイズ(バイト)' = $item.Size.ToString('N0', $Invariant); '最終更新日時' = $timeText }
                }
            )
            Save-Csv $rows @('種類', '結果', 'ファイル', 'サイズ(MB)', 'サイズ(バイト)', '最終更新日時') (Join-Path $runDir 'files.csv')
            Save-Csv $skipped @('理由', 'パス', '詳細') (Join-Path $runDir 'skipped.csv')
            # robocopy の記録は UTF-16 なので、ほかの出力とそろえて BOM 付き UTF-8 にする
            if (Test-Path -LiteralPath $resultLog) {
                [IO.File]::WriteAllText((Join-Path $runDir 'robocopy.log'), (Read-UnicodeText $resultLog), $Utf8Bom)
            }
            Write-Host ''
            Write-Host ('保存先: {0}' -f $runDir)
        } catch {
            Write-Warning ('結果を保存できませんでした: {0}' -f (Get-ErrorText (Get-InnerException $_.Exception)))
            try { [IO.Directory]::Delete($runDir, $true) } catch { }
        }
    } elseif ($saveWarning) {
        Write-Host ''
        Write-Warning $saveWarning
    }

    #endregion

    $completed = $true
} finally {
    Clear-FolderAttribute (ConvertTo-LongPath $dstPath) $hideAttributes
    foreach ($tempLog in $planLog, $copyLog) {
        if (Test-Path -LiteralPath $tempLog) { try { Remove-Item -LiteralPath $tempLog -Force } catch { } }
    }
    if ($copyStarted -and -not $completed) {
        # 中断したときは、結果を残さない（このツールが作った保存先のフォルダだけを消す）
        if ($runDir) { try { [IO.Directory]::Delete($runDir, $true) } catch { } }
        Write-Host '中断しました。コピーし終わったファイルは、コピー先に残っています。もう一度実行すると、残りのファイルだけをコピーします。' -ForegroundColor Yellow
    }
}

exit $exitCode
