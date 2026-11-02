<#
.SYNOPSIS
    指定したファイルやフォルダを使っているプロセスとクライアントを突き止めます。

.DESCRIPTION
    指定したファイルやフォルダ（中身を含む）を使っているプロセス（PID、プロセス名、サービス、実行ユーザー、
    起動日時、実行ファイル、コマンドライン）を一覧にし、画面に表示します。結果は CSV（BOM 付き UTF-8）でも保存します。
    ファイルサーバーの上で管理者として実行すると、共有フォルダ経由で開いているクライアントとユーザーも表示します。

    読み取りだけを行い、プロセスを止めたり、共有の接続を切ったりすることはありません。
    開けるかの確認のためにファイルを書き込みのモードで開きますが、何も書かずにすぐに閉じます。
    コマンドラインはパスワードなどが含まれることがあるため、画面には表示せず CSV にだけ書きます。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Find-FileLock.ps1 -Path D:\共有\経理\売上.xlsx

.PARAMETER Path
    調べるファイルかフォルダ。フォルダを指定すると、中のファイルとフォルダもすべて調べます。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Find-FileLock.ps1 -Path D:\共有\経理\売上.xlsx

    売上.xlsx を使っているプロセスと、共有フォルダ経由で開いているクライアントを調べます。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Find-FileLock.ps1 -Path D:\app\logs

    フォルダの中身も含めて調べます。

.NOTES
    Find-FileLock 1.0.0
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
    [string]$Path,

    [string]$OutputFolder,

    [switch]$NoSave
)

# 制約付き言語モード（AppLocker や WDAC の環境）では、.NET の型も Add-Type も使えないので、最初に止める
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host 'Windows の API を呼ぶ準備ができません。制約付き言語モードか、コンパイラーの実行が禁止されている環境では使えません'
    exit 1
}

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$ToolTitle     = 'ファイルロック特定'
$ScreenRows    = 200       # 画面に表示する行数（欄ごと。保存するときだけ。超えた分はファイルを見てもらう）
$ProgressMs    = 250       # 進捗表示の間隔（ミリ秒）
# オンライン専用ファイルの属性（OFFLINE・RECALL_ON_OPEN・RECALL_ON_DATA_ACCESS）。読むとダウンロードが始まる
$OnlineMask    = 0x00441000
$InUseText     = '不可（ほかのプログラムが使用中）'
$LockColumns   = '種類', 'ファイル', 'PID', 'プロセス', 'サービス', '実行ユーザー', '起動日時', '実行ファイル', 'コマンドライン',
                 'クライアント', 'クライアントのユーザー', '権限', '読み取り', '書き込み'
$SkipColumns   = '理由', 'パス', '詳細'
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$OnWindows     = [IO.Path]::DirectorySeparatorChar -eq '\'   # 開発用に Mac の pwsh でも動かせるようにするため
$machineName   = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

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
    if (-not $OnWindows) { return $FullPath }
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

# CSV の 1 行
function ConvertTo-CsvLine([object[]]$Fields) {
    $parts = New-Object 'string[]' $Fields.Count
    for ($i = 0; $i -lt $Fields.Count; $i++) { $parts[$i] = ([string]$Fields[$i]).Replace('"', '""') }
    return '"' + ($parts -join '","') + '"'
}

# 5.1 と 7 のどちらでも BOM 付き UTF-8 で保存する（Export-Csv の -Encoding UTF8 は 7 だと BOM なしになる）
function Save-Csv([string]$FilePath, [string[]]$Columns, $Rows) {
    $writer = New-Object IO.StreamWriter($FilePath, $false, $Utf8Bom)
    try {
        $writer.WriteLine((ConvertTo-CsvLine $Columns))
        foreach ($row in $Rows) { $writer.WriteLine((ConvertTo-CsvLine $row)) }
    } finally {
        $writer.Dispose()
    }
}

# Excel は、= + - @ で始まる値を数式に、数字と区切りの記号だけの値（2026-10-31、001）や「.1」「1e5」を
# 数値や日付に変えてしまうため、その形の値は ="..." で囲む。式に書けない 255 文字を超える値は、先頭に空白を足す
$excelRisk = New-Object Text.RegularExpressions.Regex ('^(?:[=+\-@]|[\d\s.,/:+\-]+$)')
function Protect-ExcelText([string]$Text) {
    if ($Text -eq '') { return $Text }
    $risky = $excelRisk.IsMatch($Text)
    if (-not $risky) {
        # 全角の数字も Excel は数値にするため、半角にそろえてから判定する。
        # 壊れた文字（対になっていないサロゲート）を含むと Normalize が例外を出すので、そのときはそのまま判定する
        $normalized = $Text
        try { $normalized = $Text.Normalize([Text.NormalizationForm]::FormKC) } catch { }
        $number = 0.0
        $risky = [double]::TryParse($normalized.TrimEnd('%'), [Globalization.NumberStyles]::Float, $Invariant, [ref]$number)
    }
    if (-not $risky) { return $Text }
    if ($Text.Length -le 255) { return '="' + $Text.Replace('"', '""') + '"' }
    return ' ' + $Text
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

function Stop-WithError([string]$Message) {
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value) {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: -Path "C:\data"）。最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# 入力されたパスをフルパスにする。末尾の \ は、ドライブのルート以外は外す
function Resolve-FullPath([string]$Text) {
    $full = ConvertTo-DisplayPath $Text
    try {
        $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($full)
        $full = [IO.Path]::GetFullPath($full)
    } catch { }
    if ($full -notmatch '^[A-Za-z]:\\$' -and $full -ne '/') { $full = $full.TrimEnd('\') }
    return $full
}

# 共有フォルダかネットワークドライブのパスか
function Test-NetworkPath([string]$FullPath) {
    if ($FullPath.StartsWith('\\')) { return $true }
    if ($FullPath -match '^[A-Za-z]:') {
        try { return (New-Object IO.DriveInfo ($FullPath.Substring(0, 2))).DriveType -eq 'Network' } catch { }
    }
    return $false
}

# 共有フォルダやネットワークドライブを開けないときの案内。管理者として実行しているときは、そのことも添える
function Get-NetworkHint([string]$FullPath) {
    $isNetwork = Test-NetworkPath $FullPath
    if (-not $isNetwork -and $FullPath -match '^[A-Za-z]:') {
        # 管理者として実行していると、通常のユーザーで割り当てたネットワークドライブはドライブごと見えない
        if (-not [IO.Directory]::Exists($FullPath.Substring(0, 2) + '\')) { $isNetwork = $isAdmin }
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

#endregion

#region 引数の確認

$startTime = Get-Date
$isAdmin   = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

if (-not $Path) { Stop-WithError '調べるファイルかフォルダを -Path で指定してください' }
$targetText = Remove-Quote $Path
if ($targetText -eq '') { Stop-WithError '調べるファイルかフォルダを -Path で指定してください' }
if ($targetText.Contains('"')) { Stop-WithQuoteError '-Path' $targetText }
# コマンドプロンプトで "D:\" と書くと D: で届き、ドライブのルートではなく、そのドライブの現在のフォルダを指してしまう
if ($targetText -match '^(\\\\\?\\)?[A-Za-z]:$') { $targetText += '\' }
$targetFull = Resolve-FullPath (Resolve-StartPath $targetText)
$targetLong = ConvertTo-LongPath $targetFull

$targetIsFile = $false
$targetIsDir  = $false
try { $targetIsFile = [IO.File]::Exists($targetLong) } catch { }
try { $targetIsDir = [IO.Directory]::Exists($targetLong) } catch { }
if (-not $targetIsFile -and -not $targetIsDir) {
    Stop-WithError ((@(('ファイルもフォルダも見つかりません: {0}' -f $targetFull)) + @(Get-NetworkHint $targetFull)) -join "`r`n")
}
$targetIsNetwork = Test-NetworkPath $targetFull

$outputText = $null
if ($OutputFolder) {
    $outputText = Remove-Quote $OutputFolder
    if ($outputText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outputText }
    if ($outputText -match '^(\\\\\?\\)?[A-Za-z]:$') { $outputText += '\' }
    $outputText = Resolve-StartPath $outputText
}

#endregion

#region Windows の API

# NtQueryInformationFile の FileProcessIdsUsingFileInformation（47）で、ファイルやフォルダを開いているプロセスの番号を取る。
# FILE_READ_ATTRIBUTES だけで開くので、共有違反にならず、ファイルの中身も読まない（オンライン専用ファイルもダウンロードされない）。
# openLink が真のときは FILE_FLAG_OPEN_REPARSE_POINT で、リンクの先ではなくそのものを開く（フォルダの中をたどるとき）。
# -Path に指定した場所は、リンクなら先を調べる。戻り値は Windows のエラーコードで、0 が成功
$nativeSource = @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Workbench.FindFileLock {
    public static class Native {
        [StructLayout(LayoutKind.Sequential)]
        private struct IoStatusBlock { public IntPtr Status; public UIntPtr Information; }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateFileW")]
        private static extern SafeFileHandle CreateFile(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);

        [DllImport("ntdll.dll")]
        private static extern int NtQueryInformationFile(SafeFileHandle handle, out IoStatusBlock status, IntPtr info, uint length, int infoClass);

        [DllImport("ntdll.dll")]
        private static extern int RtlNtStatusToDosError(int status);

        public static int Query(string path, bool openLink, out int[] ids) {
            ids = new int[0];
            const uint FileReadAttributes = 0x80, ShareAll = 7, OpenExisting = 3, BackupSemantics = 0x02000000, OpenReparsePoint = 0x00200000;
            const int FileProcessIdsUsingFileInformation = 47;
            int infoLengthMismatch = unchecked((int)0xC0000004), bufferOverflow = unchecked((int)0x80000005);
            using (SafeFileHandle handle = CreateFile(path, FileReadAttributes, ShareAll, IntPtr.Zero, OpenExisting, BackupSemantics | (openLink ? OpenReparsePoint : 0), IntPtr.Zero)) {
                if (handle.IsInvalid) { return Marshal.GetLastWin32Error(); }
                int size = 4096;
                while (true) {
                    IntPtr buffer = Marshal.AllocHGlobal(size);
                    try {
                        IoStatusBlock status;
                        int result = NtQueryInformationFile(handle, out status, buffer, (uint)size, FileProcessIdsUsingFileInformation);
                        if ((result == infoLengthMismatch || result == bufferOverflow) && size < 64 * 1024 * 1024) {
                            size *= 2;
                            continue;
                        }
                        if (result < 0) { return RtlNtStatusToDosError(result); }
                        // ULONG 個数、続いて ULONG_PTR の配列（構造体の境界に合わせるため、配列はポインターの大きさの位置から始まる）
                        int count = Marshal.ReadInt32(buffer, 0);
                        ids = new int[count];
                        for (int i = 0; i < count; i++) {
                            ids[i] = (int)Marshal.ReadIntPtr(buffer, IntPtr.Size + i * IntPtr.Size).ToInt64();
                        }
                        return 0;
                    } finally {
                        Marshal.FreeHGlobal(buffer);
                    }
                }
            }
        }
    }
}
'@

# 同じ PowerShell でもう一度実行したときは、読み込み済みの型を使う（同じ名前の型は読み込み直せない）
if (-not ('Workbench.FindFileLock.Native' -as [type])) {
    try {
        Add-Type -TypeDefinition $nativeSource -ErrorAction Stop
    } catch {
        Stop-WithError ('Windows の API を呼ぶ準備ができません。制約付き言語モードか、コンパイラーの実行が禁止されている環境では使えません（{0}）' -f (Get-ErrorText (Get-InnerException $_.Exception)))
    }
}

#endregion

$hits       = New-Object 'System.Collections.Generic.List[object]'   # プロセスが使っているもの
$clients    = New-Object 'System.Collections.Generic.List[object]'   # 共有フォルダから開いているもの
$skipped    = New-Object 'System.Collections.Generic.List[object]'
$skipKinds  = [ordered]@{ '権限なし' = 0; 'その他のエラー' = 0 }
$linkCount  = 0
$fileCount  = 0L
$dirCount   = 0L
$openChecks = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
$smbNote    = $null
$sw         = [Diagnostics.Stopwatch]::StartNew()
$nextShow   = 0L
$dirType    = [IO.DirectoryInfo]
$rpFlag     = [IO.FileAttributes]::ReparsePoint
$hsFlag     = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System

function Add-Skip([string]$Reason, [string]$ItemPath, [string]$Detail) {
    $skipped.Add([object[]]@($Reason, $ItemPath, $Detail))
    if ($skipKinds.Contains($Reason)) { $skipKinds[$Reason]++ }
}

# リパースポイントのうち、リンクと判定できたものの種類を返す（OneDrive のファイルやフォルダもリパースポイントなので、それ以外は $null）
function Get-LinkName($Entry) {
    if (($Entry.Attributes -band $rpFlag) -ne $rpFlag) { return $null }
    # Windows が互換用に作ったリンク（All Users、Application Data など）は、5.1 では種類を判定できない。どれも隠し・システム属性付き。
    # 7 は種類を判定できるが、5.1 と同じ結果になるよう、この目印を先に見る
    if ($Entry -is $dirType -and ($Entry.Attributes -band $hsFlag) -eq $hsFlag) { return 'Windows の互換用リンク' }
    $linkType = $null
    try { $linkType = $Entry.LinkType } catch { }
    if ($linkType -eq 'Junction') { return 'ジャンクション' }
    if ($linkType -eq 'SymbolicLink') { return 'シンボリックリンク' }
    return $null
}

# 1 つのファイルやフォルダを使っているプロセスを問い合わせる。問い合わせに使ったハンドルが数に入るので、このツール自身は除く
function Find-UsingProcess([string]$LongPath, [bool]$IsFile, [bool]$OpenLink) {
    $display = ConvertTo-DisplayPath $LongPath
    $ids  = $null
    $code = 0
    try {
        $code = [Workbench.FindFileLock.Native]::Query($LongPath, $OpenLink, [ref]$ids)
    } catch {
        Add-Skip 'その他のエラー' $display (Get-ErrorText (Get-InnerException $_.Exception))
        return
    }
    # 調べている間に消えたもの（ERROR_FILE_NOT_FOUND、ERROR_PATH_NOT_FOUND）は、もう使われていないので数えない
    if ($code -eq 2 -or $code -eq 3) { return }
    if ($code -ne 0) {
        $detail = Get-ErrorText (New-Object ComponentModel.Win32Exception ([int]$code))
        Add-Skip $(if ($code -eq 5) { '権限なし' } else { 'その他のエラー' }) $display $detail
        return
    }
    foreach ($usingId in $ids) {
        if ($usingId -eq $PID) { continue }
        $hits.Add([pscustomobject]@{ File = $display; ProcessId = [int]$usingId; IsFile = $IsFile })
    }
}

# フォルダの中をすべて問い合わせる。再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる
function Search-Folder([string]$RootLong) {
    $stack = New-Object 'System.Collections.Generic.Stack[object]'
    $stack.Push((New-Object IO.DirectoryInfo $RootLong))
    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        try {
            $entries = $dir.GetFileSystemInfos()
        } catch {
            $ex = Get-InnerException $_.Exception
            if ($ex -is [IO.DirectoryNotFoundException] -or $ex -is [IO.FileNotFoundException]) { continue }
            if ($ex -is [UnauthorizedAccessException]) {
                Add-Skip '権限なし' (ConvertTo-DisplayPath $dir.FullName) 'アクセスが拒否されました'
            } else {
                Add-Skip 'その他のエラー' (ConvertTo-DisplayPath $dir.FullName) (Get-ErrorText $ex)
            }
            continue
        }
        foreach ($entry in $entries) {
            if ($sw.ElapsedMilliseconds -ge $script:nextShow) {
                Write-Progress -Activity $ToolTitle -Status (Format-Invariant '調べています: ファイル {0:N0} 件 / フォルダ {1:N0} 件' $script:fileCount $script:dirCount) -CurrentOperation (ConvertTo-DisplayPath $entry.FullName)
                $script:nextShow = $sw.ElapsedMilliseconds + $ProgressMs
            }
            $linkName = $null
            if (($entry.Attributes -band $rpFlag) -eq $rpFlag) { $linkName = Get-LinkName $entry }
            if ($linkName) {
                $script:linkCount++
                $skipped.Add([object[]]@('たどらなかったリンク', (ConvertTo-DisplayPath $entry.FullName), $linkName))
            } elseif ($entry -is $dirType) {
                $script:dirCount++
                Find-UsingProcess $entry.FullName $false $true
                $stack.Push($entry)
            } else {
                $script:fileCount++
                Find-UsingProcess $entry.FullName $true $true
            }
        }
    }
}

# 指定のモードで実際に開けるかを確かめて、すぐに閉じる。読み書きと削除をすべて許して開くので、ほかのプログラムの邪魔をしない
function Test-OpenFile([string]$LongPath, [IO.FileAccess]$Access) {
    $stream = $null
    try {
        $stream = [IO.File]::Open($LongPath, [IO.FileMode]::Open, $Access, ([IO.FileShare]'ReadWrite, Delete'))
        return '可'
    } catch {
        $ex = Get-InnerException $_.Exception
        if ($ex -is [UnauthorizedAccessException]) {
            $readOnly = $false
            try { $readOnly = ([IO.File]::GetAttributes($LongPath) -band [IO.FileAttributes]::ReadOnly) -ne 0 } catch { }
            if ($readOnly -and $Access -ne [IO.FileAccess]::Read) { return '不可（読み取り専用）' }
            return '不可（権限なし）'
        }
        # ERROR_SHARING_VIOLATION（32）と ERROR_LOCK_VIOLATION（33）
        $code = $ex.HResult -band 0xFFFF
        if ($code -eq 32 -or $code -eq 33) { return $InUseText }
        return '不可（その他のエラー）'
    } finally {
        if ($stream) { $stream.Dispose() }
    }
}

# 開けるかの確認。オンライン専用ファイルは、開くとダウンロードが始まるので確かめない
function Add-OpenCheck([string]$DisplayPath) {
    if ($openChecks.ContainsKey($DisplayPath)) { return }
    $long = ConvertTo-LongPath $DisplayPath
    $online = $false
    try { $online = ([int][IO.File]::GetAttributes($long) -band $OnlineMask) -ne 0 } catch { }
    if ($online) {
        $openChecks[$DisplayPath] = [pscustomobject]@{ Read = '確かめていない'; Write = '確かめていない' }
        return
    }
    $openChecks[$DisplayPath] = [pscustomobject]@{
        Read  = Test-OpenFile $long ([IO.FileAccess]::Read)
        Write = Test-OpenFile $long ([IO.FileAccess]::Write)
    }
}

# 共有フォルダの接続の権限（アクセスマスク）を、読み取り・読み書き・その他にする
function Get-AccessText($Mask) {
    $value  = [long]$Mask
    $reads  = (0x1L -bor 0x80000000L -bor 0x10000000L)   # FILE_READ_DATA、GENERIC_READ、GENERIC_ALL
    $writes = (0x2L -bor 0x4L -bor 0x40000000L -bor 0x10000000L)   # FILE_WRITE_DATA、FILE_APPEND_DATA、GENERIC_WRITE、GENERIC_ALL
    $canRead  = ($value -band $reads) -ne 0
    $canWrite = ($value -band $writes) -ne 0
    if ($canRead -and $canWrite) { return '読み書き' }
    if ($canRead) { return '読み取り' }
    return 'その他'
}

$finished   = $false
$unexpected = $false
$saveError  = $null
$saveWarning = $null
$runDir     = $null
$lines      = New-Object 'System.Collections.Generic.List[string]'
$exitCode   = 0
try {
    #region 調べる

    Write-Progress -Activity $ToolTitle -Status '調べています'
    if ($targetIsDir) {
        $dirCount = 1L
        Find-UsingProcess $targetLong $false $false
        Search-Folder $targetLong
    } else {
        $fileCount = 1L
        Find-UsingProcess $targetLong $true $false
    }

    # 共有フォルダの接続は、このコンピューターが共有しているものだけが分かる。ネットワーク上のパスとは突き合わせられない
    if (-not $targetIsNetwork -and (Get-Command Get-SmbOpenFile -ErrorAction SilentlyContinue)) {
        Write-Progress -Activity $ToolTitle -Status '共有フォルダの接続を調べています'
        $targetTrim = $targetFull.TrimEnd('\')
        try {
            foreach ($open in @(Get-SmbOpenFile -ErrorAction Stop)) {
                $openPath = ([string]$open.Path).TrimEnd('\')
                $isUnder  = $openPath.Equals($targetTrim, [StringComparison]::OrdinalIgnoreCase) -or
                    ($targetIsDir -and $openPath.StartsWith($targetTrim + '\', [StringComparison]::OrdinalIgnoreCase))
                if (-not $isUnder) { continue }
                $clients.Add([pscustomobject]@{
                    File   = $openPath
                    Client = [string]$open.ClientComputerName
                    User   = [string]$open.ClientUserName
                    Access = Get-AccessText $open.Permissions
                })
            }
        } catch {
            if (-not $isAdmin) {
                $smbNote = '※ 共有フォルダの接続元は、管理者として実行すると調べられます'
            } else {
                $smbNote = '※ 共有フォルダの接続元を調べられませんでした（{0}）' -f (Get-ErrorText (Get-InnerException $_.Exception))
            }
        }
    } elseif (-not $targetIsNetwork) {
        $smbNote = '※ 共有フォルダの接続元を調べられませんでした（Get-SmbOpenFile が見つかりません）'
    }

    # 開けるかの確認の対象は、指定したファイルと、使っているものが見つかったファイル
    Write-Progress -Activity $ToolTitle -Status '開けるかを確かめています'
    if ($targetIsFile) { Add-OpenCheck $targetFull }
    foreach ($hit in $hits) { if ($hit.IsFile) { Add-OpenCheck $hit.File } }
    foreach ($client in $clients) {
        $isFile = $false
        try { $isFile = [IO.File]::Exists((ConvertTo-LongPath $client.File)) } catch { }
        if ($isFile) { Add-OpenCheck $client.File }
    }

    #endregion

    #region プロセスとサービスの情報

    $pidSet = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($hit in $hits) { [void]$pidSet.Add($hit.ProcessId) }
    $processes = @{}
    $services  = @{}
    if ($pidSet.Count -gt 0) {
        Write-Progress -Activity $ToolTitle -Status 'プロセスの情報を取っています'

        # PID → 実行ユーザー。Win32_Process の GetOwner は、呼ぶたびにすべてのプロセスを調べ直すため、1 回 0.3 秒ほどかかる。
        # Get-Process -IncludeUserName はまとめて取れるので、こちらを先に使い、取れなかったプロセスだけ GetOwner で取る
        # （5.1 では、管理者でないと Get-Process -IncludeUserName を使えないので、すべて GetOwner で取る）
        $owners = @{}
        try {
            foreach ($gp in @(Get-Process -IncludeUserName -ErrorAction Stop)) {
                if ($pidSet.Contains($gp.Id) -and $gp.UserName) { $owners[$gp.Id] = [string]$gp.UserName }
            }
        } catch { }

        # PID → プロセスの情報。管理者でないと、ほかのユーザーのプロセスの実行ファイル・コマンドライン・実行ユーザーは取れない
        try {
            foreach ($p in @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop)) {
                $processId = [int]$p.ProcessId
                if (-not $pidSet.Contains($processId)) { continue }
                $owner = ''
                if ($owners.ContainsKey($processId)) {
                    $owner = $owners[$processId]
                } elseif ($processId -gt 4 -and ($isAdmin -or $p.ExecutablePath)) {
                    # 管理者なら、Get-Process で取れない保護されたプロセス（lsass.exe など）も GetOwner で取れる。
                    # 管理者でないときは、実行ファイルが取れないプロセス（ほかのユーザーのものなど）では GetOwner もほとんど失敗し、
                    # 7 の Get-Process -IncludeUserName でも取れない。時間を短くし、5.1 と 7 で結果をそろえるため、呼ばない。
                    # PID 0 と 4 は、管理者でも取れない
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
                if ($processId -eq 4) { $name = $gp.ProcessName }
                try { $started = $gp.StartTime } catch { }
            } catch { }
            $processes[$processId] = [pscustomobject]@{ Name = $name; Path = ''; CommandLine = ''; Started = $started; Owner = '' }
        }

        # PID → そのプロセスで動いているサービスの名前
        try {
            foreach ($s in @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop)) {
                $processId = [int]$s.ProcessId
                if ($processId -eq 0 -or -not $pidSet.Contains($processId)) { continue }
                if (-not $services.ContainsKey($processId)) { $services[$processId] = New-Object 'System.Collections.Generic.List[string]' }
                $services[$processId].Add([string]$s.Name)
            }
        } catch { }
        # 行ごとに並べ替えないよう、ここで 1 つの文字列にしておく
        foreach ($key in @($services.Keys)) { $services[$key] = @($services[$key].ToArray() | Sort-Object) -join ', ' }
    }

    # サービスの列。PID 4 は、プロセスではないことが分かるように注記する
    function Get-ServiceText([int]$ProcessId) {
        if ($ProcessId -eq 4) { return '（カーネル：共有フォルダ経由の接続など）' }
        if (-not $services.ContainsKey($ProcessId)) { return '' }
        return $services[$ProcessId]
    }

    # 管理者でないため取れなかった情報があるか（PID 4 は、管理者でも取れない項目があるので数えない）
    $missingInfo = $false
    if (-not $isAdmin) {
        foreach ($processId in $pidSet) {
            if ($processId -le 4) { continue }
            $info = $processes[$processId]
            if ($info.Name -ne '（終了したプロセス）' -and (-not $info.Path -or -not $info.Owner)) { $missingInfo = $true }
        }
    }

    #endregion

    #region 結果をまとめる

    # 並びは 5.1 と 7 で同じになるよう、文字コードの順にする
    $hits.Sort([Comparison[object]] {
        param($x, $y)
        $order = [string]::Compare($x.File, $y.File, [StringComparison]::OrdinalIgnoreCase)
        if ($order -ne 0) { return $order }
        return $x.ProcessId.CompareTo($y.ProcessId)
    })
    $clients.Sort([Comparison[object]] {
        param($x, $y)
        $order = [string]::Compare($x.File, $y.File, [StringComparison]::OrdinalIgnoreCase)
        if ($order -ne 0) { return $order }
        $order = [string]::Compare($x.Client, $y.Client, [StringComparison]::OrdinalIgnoreCase)
        if ($order -ne 0) { return $order }
        return [string]::Compare($x.User, $y.User, [StringComparison]::OrdinalIgnoreCase)
    })
    $hitList    = $hits.ToArray()
    $clientList = $clients.ToArray()
    $checkNames = New-Object 'string[]' $openChecks.Count
    $openChecks.Keys.CopyTo($checkNames, 0)
    [Array]::Sort($checkNames, [StringComparer]::OrdinalIgnoreCase)

    # 共有違反で開けないのに、使っているものが見つからないことがある（別の PC からの接続、ドライバーなど）。使われていることは確かなので「使用中」とする
    $blockedUnknown = $false
    foreach ($name in $checkNames) {
        if ($openChecks[$name].Read -eq $InUseText -or $openChecks[$name].Write -eq $InUseText) { $blockedUnknown = $true }
    }
    $foundAny = $hitList.Count -gt 0 -or $clientList.Count -gt 0
    $skipTotal = $skipKinds['権限なし'] + $skipKinds['その他のエラー']
    if ($foundAny -or $blockedUnknown) { $exitCode = 2 }
    elseif ($skipTotal -gt 0) { $exitCode = 1 }

    $processRows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($hit in $hitList) {
        $info = $processes[$hit.ProcessId]
        $startedText = if ($info.Started) { ([datetime]$info.Started).ToString('yyyy-MM-dd HH:mm', $Invariant) } else { '' }
        $processRows.Add([string[]]@([string]$hit.ProcessId, $info.Name, $info.Owner, $startedText, (Get-ServiceText $hit.ProcessId), $hit.File))
    }
    $clientRows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($client in $clientList) { $clientRows.Add([string[]]@($client.Client, $client.User, $client.Access, $client.File)) }
    $checkRows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($name in $checkNames) { $checkRows.Add([string[]]@($openChecks[$name].Read, $openChecks[$name].Write, $name)) }

    $sections = @(
        [pscustomobject]@{ Title = '■ このコンピューターのプロセス'; Head = [string[]]@('PID', 'プロセス', '実行ユーザー', '起動日時', 'サービス', 'ファイル'); Right = [bool[]]@($true, $false, $false, $false, $false, $false); Rows = $processRows.ToArray() }
        [pscustomobject]@{ Title = '■ 共有フォルダから開いているクライアント'; Head = [string[]]@('クライアント', 'ユーザー', '権限', 'ファイル'); Right = [bool[]]@($false, $false, $false, $false); Rows = $clientRows.ToArray() }
        [pscustomobject]@{ Title = '■ 開けるかの確認'; Head = [string[]]@('読み取り', '書き込み', 'ファイル'); Right = [bool[]]@($false, $false, $false); Rows = $checkRows.ToArray() }
    )

    $header = New-Object 'System.Collections.Generic.List[string]'
    $header.Add((Format-Invariant '{0} (Find-FileLock {1})' $ToolTitle $ScriptVersion))
    $header.Add(('対象: {0}' -f $targetFull))
    if ($targetIsDir) { $header.Add((Format-Invariant '調べた数: ファイル {0:N0} 件  フォルダ {1:N0} 件' $fileCount $dirCount)) }
    $header.Add((Format-Invariant '{0}   {1:yyyy-MM-dd HH:mm:ss}   管理者: {2}' $machineName $startTime $(if ($isAdmin) { 'はい' } else { 'いいえ' })))

    $elapsed = Format-Duration $sw.Elapsed
    $tail = New-Object 'System.Collections.Generic.List[string]'
    if ($foundAny) {
        $tail.Add((Format-Invariant '結果: 使用中   プロセス {0:N0} 個  クライアント {1:N0} 件   所要 {2}' $pidSet.Count $clientList.Count $elapsed))
    } elseif ($blockedUnknown) {
        $tail.Add(('結果: 使用中（使っているものは特定できません）   所要 {0}' -f $elapsed))
        $tail.Add('※ 使っているものは見つかりませんが、書き込みで開けないファイルがあります。別の PC から共有フォルダ経由で使われているかもしれません。ファイルサーバーの上で実行してください')
    } else {
        $tail.Add(('結果: 使用中のものは見つかりません   所要 {0}' -f $elapsed))
    }
    if ($skipTotal -gt 0) {
        $parts = @(foreach ($kind in $skipKinds.Keys) { if ($skipKinds[$kind] -gt 0) { Format-Invariant '{0} {1:N0}' $kind $skipKinds[$kind] } })
        $tail.Add((Format-Invariant 'スキップ {0:N0} 件（{1}）' $skipTotal ($parts -join '・')))
    }
    if ($linkCount -gt 0) { $tail.Add((Format-Invariant 'たどらなかったリンク {0:N0} 件' $linkCount)) }
    if ($smbNote) { $tail.Add($smbNote) }
    if ($targetIsNetwork) { $tail.Add('※ 別の PC からは、この PC のプロセスしか分かりません。誰が開いているかは、ファイルサーバーの上で、そのサーバーでのパス（D:\共有\... など）を指定して実行してください') }
    if ($targetIsDir) {
        $ownDir  = [Environment]::CurrentDirectory.TrimEnd('\')
        $rootDir = $targetFull.TrimEnd('\')
        if ($ownDir.Equals($rootDir, [StringComparison]::OrdinalIgnoreCase) -or $ownDir.StartsWith($rootDir + '\', [StringComparison]::OrdinalIgnoreCase)) {
            $tail.Add('※ このツールを、調べるフォルダの中で実行しています。別のフォルダから実行してください')
        }
    }
    if ($missingInfo) { $tail.Add('※ 管理者として実行すると、ほかのユーザーのプロセスの実行ユーザー・実行ファイル・コマンドラインも取れます（コマンドラインは CSV にだけ書きます）') }

    # 欄ごとの表の行。Rows が空のときは（なし）だけにする
    $sectionLines = @(foreach ($section in $sections) {
        $tableLines = if ($section.Rows.Count -gt 0) { Format-Table $section.Head $section.Right $section.Rows } else { [string[]]@('    （なし）') }
        , $tableLines
    })

    #endregion

    #region 保存

    if (-not $NoSave) {
        $baseDir = if ($outputText) { $outputText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
        try {
            $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
            # \\?\ 付きのパスでは末尾のピリオドや空白が残るため、先に通常の書き方で整える（260 文字を超えると例外になるので、そのときはそのまま）
            try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
            $runName = Format-Invariant 'FileLock_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
            $runDir  = Join-Path $baseDir $runName
            # 失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
            $suffix = 1
            while ([IO.Directory]::Exists((ConvertTo-LongPath $runDir)) -or [IO.File]::Exists((ConvertTo-LongPath $runDir))) {
                $suffix++
                $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix)
            }
            [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $runDir))
        } catch {
            $saveWarning = '保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, (Get-ErrorText (Get-InnerException $_.Exception))
            $runDir = $null
        }
    }

    if ($runDir) {
        # summary.txt には、すべての行を書く
        $lines.AddRange([string[]]$header)
        for ($i = 0; $i -lt $sections.Count; $i++) {
            $lines.Add('')
            $lines.Add($sections[$i].Title)
            $lines.AddRange([string[]]$sectionLines[$i])
        }
        $lines.Add('')
        $lines.AddRange([string[]]$tail.ToArray())
        $lines.Add(('保存先: {0}' -f $runDir))

        $csvRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($hit in $hitList) {
            $info = $processes[$hit.ProcessId]
            $read = ''; $write = ''
            if ($openChecks.ContainsKey($hit.File)) { $read = $openChecks[$hit.File].Read; $write = $openChecks[$hit.File].Write }
            $csvRows.Add([object[]]@('プロセス', (Protect-ExcelText $hit.File), $hit.ProcessId, (Protect-ExcelText $info.Name),
                (Protect-ExcelText (Get-ServiceText $hit.ProcessId)), (Protect-ExcelText $info.Owner),
                $(if ($info.Started) { ([datetime]$info.Started).ToString('yyyy/MM/dd HH:mm:ss', $Invariant) } else { '' }),
                (Protect-ExcelText $info.Path), (Protect-ExcelText $info.CommandLine), '', '', '', $read, $write))
        }
        foreach ($client in $clientList) {
            $read = ''; $write = ''
            if ($openChecks.ContainsKey($client.File)) { $read = $openChecks[$client.File].Read; $write = $openChecks[$client.File].Write }
            $csvRows.Add([object[]]@('クライアント', (Protect-ExcelText $client.File), '', '', '', '', '', '', '',
                (Protect-ExcelText $client.Client), (Protect-ExcelText $client.User), $client.Access, $read, $write))
        }
        $skipRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($row in $skipped) { $skipRows.Add([object[]]@($row[0], (Protect-ExcelText $row[1]), $row[2])) }

        Write-Progress -Activity $ToolTitle -Status '結果を保存しています'
        try {
            [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$lines, $Utf8Bom)
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'locks.csv')) $LockColumns $csvRows
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $skipRows
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
            try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { }
            $runDir = $null
        }
    }

    #endregion

    $finished = $true
} catch {
    # 想定外のエラーは、Ctrl+C の「中断しました」と区別して知らせる
    $unexpected = $true
    throw
} finally {
    Write-Progress -Activity $ToolTitle -Completed
    if (-not $finished) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        else { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

#region 画面表示

foreach ($line in $header) { Write-Host $line }
for ($i = 0; $i -lt $sections.Count; $i++) {
    Write-Host ''
    Write-Host $sections[$i].Title -ForegroundColor Cyan
    $tableLines = @($sectionLines[$i])   # （なし）の 1 行だけのときは、配列ではなく文字列で入っている
    # 保存できたときは、画面には先頭の分だけを出す。保存しなかったときは、ほかに見る場所がないためすべて出す
    # （表の 1 行目は見出し）
    $limit = if ($runDir) { $ScreenRows + 1 } else { $tableLines.Count }
    $shown = [math]::Min($limit, $tableLines.Count)
    for ($j = 0; $j -lt $shown; $j++) { Write-Host $tableLines[$j] }
    if ($tableLines.Count -gt $shown) {
        Write-Host (Format-Invariant '（ほか {0:N0} 行は locks.csv を参照してください）' ($tableLines.Count - $shown))
    }
}
Write-Host ''
foreach ($line in $tail) { Write-Host $line }

if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($saveWarning) {
    Write-Warning $saveWarning
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
}

#endregion

exit $exitCode
