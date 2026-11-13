<#
.SYNOPSIS
    このサーバーに、誰が・いつ・どこからログオンしたかと、ログオンの失敗を、Security ログから一覧にします。

.DESCRIPTION
    Security ログのログオン（4624）とログオフ（4634・4647）をログオン ID で組にして、ログオンごとの時間を出します。
    ログオンの失敗（4625）は、ユーザーと接続元ごとにまとめます。リモートデスクトップの切断と再接続は、
    セッションのログ（TerminalServices-LocalSessionManager/Operational）から読みます。
    結果は画面に表示し、summary.txt・sessions.csv・events.csv・users.csv・skipped.csv（BOM 付き UTF-8）に保存します。
    期間の前にログオンして、期間中も続いていたログオンも出します。

    イベントログを変更・削除することはありません。
    Security ログは、管理者でないと読めません。ほかの PC から集めた .evtx ファイルは、-Path で読めます。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Get-LogonHistory.ps1

.PARAMETER Since
    期間の始め。期間に重なるログオン（期間より前にログオンして、期間の始めにまだ続いていたものを含む）と、期間の中の失敗などを取り出します。
    "2026-11-10 09:00" のような日時か、2h・30m・3d のような今からの時間で書きます。既定値は 7d（直近 7 日）です。

.PARAMETER Until
    期間の終わり。書き方は Since と同じです。

.PARAMETER User
    取り出すユーザー。カンマで区切って複数書けます。* は任意の文字列です。ドメインを付けないと、どのドメインのユーザーにも当たります。

.PARAMETER LogonType
    取り出すログオンの種類。Interactive（対話）、RemoteDesktop（リモートデスクトップ）、Unlock（ロック解除）、
    Network（ネットワーク）、Batch（バッチ）、Service（サービス）、NewCredentials（別の資格情報）を、カンマで区切って書きます。
    All ですべての種類です。既定値は Interactive,RemoteDesktop,Unlock です。ログオンの失敗は、この指定に関係なくすべて取り出します。

.PARAMETER Path
    読む .evtx ファイル、または .evtx ファイルのあるフォルダ（サブフォルダも探します）。カンマで区切って複数書けます。
    Security ログと、リモートデスクトップのセッションのログを読みます。指定すると、この PC のイベントログは読みません。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。-OutputFolder と一緒には使えません。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-LogonHistory.ps1

    直近 7 日のログオン（対話・リモートデスクトップ・ロック解除）と、ログオンの失敗を一覧にします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-LogonHistory.ps1 -Since 30d -User "yamada,*admin*"

    直近 30 日の、指定したユーザーのログオンを一覧にします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-LogonHistory.ps1 -Path D:\collected\ap01 -Since 30d

    ほかのサーバーから集めた .evtx ファイルを読みます。

.NOTES
    Get-LogonHistory 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$Since = '7d',

    [string]$Until,

    [string]$User,

    [string]$LogonType = 'Interactive,RemoteDesktop,Unlock',

    [string]$Path,

    [string]$OutputFolder,

    [switch]$NoSave
)

# 制約付き言語モード（AppLocker や WDAC の環境）では、このツールが使う .NET の型を使えないので、最初に止める
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host '制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）'
    exit 1
}

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion    = '1.0.0'
$ToolName         = 'Get-LogonHistory'
$ToolTitle        = 'ログオンの履歴'
$ScreenRows       = 20       # 画面に表示する各表の行数（保存するときだけ）
$NameWidth        = 20       # 画面の表のユーザーの幅（半角の文字数）。超える分は … にする
$SourceWidth      = 15       # 画面の表の接続元の幅。IPv4 アドレスと NetBIOS の名前はすべて入る。ログオンの表が 120 桁に収まるように、ユーザーより狭くする
$ShortSeconds     = 5        # これより短く終わった対話・リモートデスクトップのログオンは、ロック解除や再接続のときの確かめとみなす
$UnlockGapSeconds = 2        # 1 回のロック解除で記録される、ロック解除のログオンの間隔
$SignInSeconds    = 5        # ログオンとこの秒数の内にあるロック解除は、サインインのときに記録されるもの
$LsmBeforeSeconds = 5        # セッションのログのログオンが、Security のログオンより前にずれてよい秒数
$LsmAfterSeconds  = 300      # 同じく、後にずれてよい秒数（プロファイルを作るときは時間がかかる）
$ReconnectSeconds = 10       # Security の再接続とセッションのログの再接続を、同じものとみなす秒数
$ProgressMs       = 250      # 進捗表示の間隔（ミリ秒）
$LongPathLength   = 250      # この長さ以上の .evtx のパスは、\\?\ を付けて読む
$SecurityLog      = 'Security'
$LsmLog           = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
$LogonAuditGuid   = [guid]'0cce9215-69ae-11d9-bed3-505054503030'   # 監査ポリシーの「ログオン」
$LogoffAuditGuid  = [guid]'0cce9216-69ae-11d9-bed3-505054503030'   # 監査ポリシーの「ログオフ」
$SessionColumns   = 'ログオン', 'ログオフ', '時間', 'ユーザー', '種類', '接続元 IP アドレス', '接続元の名前', '管理者権限', '状態', 'ロック解除の回数', '切断の回数', '再接続の回数', '最後の切断', '認証', 'ログオン ID', 'セッション ID', 'コンピューター'
$EventColumns     = '日時', '操作', 'ユーザー', '種類', '接続元 IP アドレス', '接続元の名前', '詳細', 'ログオン ID', 'セッション ID', 'イベント ID', 'コンピューター'
$UserColumns      = 'ユーザー', 'ログオン回数', '最初のログオン', '最後のログオン', '合計の時間', '失敗回数', '最後の失敗'
$SkipColumns      = '理由', '対象', '詳細'
$TimeFormat       = 'yyyy/MM/dd HH:mm:ss'
$ScreenTime       = 'yyyy-MM-dd HH:mm'      # 画面の表は秒を省く（表を 120 桁に収めるため。CSV には秒まで書く）
$Invariant        = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom          = New-Object Text.UTF8Encoding $true
$machineName      = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }

# ログオンの種類。キーは -LogonType の名前（日本語の名前でも書ける）、Types は Windows の記録の LogonType
$TypeGroups = [ordered]@{
    'Interactive'    = @{ Label = '対話'; Types = @(2, 11) }
    'RemoteDesktop'  = @{ Label = 'リモートデスクトップ'; Types = @(10, 12) }
    'Unlock'         = @{ Label = 'ロック解除'; Types = @(7, 13) }
    'Network'        = @{ Label = 'ネットワーク'; Types = @(3, 8) }
    'Batch'          = @{ Label = 'バッチ'; Types = @(4) }
    'Service'        = @{ Label = 'サービス'; Types = @(5) }
    'NewCredentials' = @{ Label = '別の資格情報'; Types = @(9) }
}
$SessionTypes = @(2, 10, 11, 12)   # 画面の前に人が座る（またはリモートデスクトップでつなぐ）ログオン
$UnlockTypes  = @(7, 13)
$RemoteTypes  = @(10, 12)

# ログオンの失敗の状態コード（Status か SubStatus）と理由。
# キーは 16 進数の文字（0xC000006A のような数値のリテラルは Int32 の負の数になり、記録の値と比べられないため）
$FailureReasons = @{
    'C000006D' = 'ユーザー名かパスワードが違う'
    'C000006A' = 'パスワードが違う'
    'C0000064' = 'ユーザー名がない'
    'C0000072' = 'アカウントが無効'
    'C0000234' = 'アカウントがロックされている'
    'C0000071' = 'パスワードの期限切れ'
    'C0000193' = 'アカウントの期限切れ'
    'C0000224' = 'パスワードの変更が必要'
    'C000006F' = 'ログオンできる時間外'
    'C0000070' = 'ログオンできない端末'
    'C000015B' = 'このログオンの種類は許可されていない'
    'C0000133' = '時刻がずれている'
    'C000005E' = 'ログオン サーバーがない'
    'C000006E' = 'アカウントの制限'
    'C0000413' = '認証ファイアウォール'
    'C0000192' = 'Netlogon サービスが開始されていない'
}
# まとめのコード（ユーザー名かパスワードが違う、アカウントの制限）。このときは SubStatus で理由を見る
$SummaryStatus = @('C000006D', 'C000006E')

# 監査ポリシーを読む（AuditQuerySystemPolicy）。管理者でも SeSecurityPrivilege が要ることがあるので、そのときは一時的に有効にする
$AuditSource = @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
public static class WbLogonHistoryAudit {
    [StructLayout(LayoutKind.Sequential)]
    private struct AuditPolicyInformation { public Guid SubCategory; public uint AuditingInformation; public Guid Category; }
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    private struct TokenPrivileges { public uint Count; public long Luid; public uint Attributes; }
    // 戻り値は BOOLEAN（1 バイト）。既定の bool（4 バイトの BOOL）で受けると、失敗しても成功に見えることがある
    [DllImport("advapi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.U1)]
    private static extern bool AuditQuerySystemPolicy([In] Guid[] subCategoryGuids, uint policyCount, out IntPtr auditPolicy);
    [DllImport("advapi32.dll")]
    private static extern void AuditFree(IntPtr buffer);
    [DllImport("kernel32.dll")]
    private static extern IntPtr GetCurrentProcess();
    [DllImport("kernel32.dll")]
    private static extern bool CloseHandle(IntPtr handle);
    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);
    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool LookupPrivilegeValue(string systemName, string name, out long luid);
    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool AdjustTokenPrivileges(IntPtr token, bool disableAll, ref TokenPrivileges newState, uint length, ref TokenPrivileges previousState, out uint returnLength);

    private static uint QueryOnce(Guid subCategory, out int error) {
        IntPtr buffer;
        error = 0;
        if (!AuditQuerySystemPolicy(new Guid[] { subCategory }, 1, out buffer)) { error = Marshal.GetLastWin32Error(); return 0; }
        if (buffer == IntPtr.Zero) { error = 13; return 0; }
        try { return ((AuditPolicyInformation)Marshal.PtrToStructure(buffer, typeof(AuditPolicyInformation))).AuditingInformation; }
        finally { AuditFree(buffer); }
    }

    // 1 = 成功を記録する、2 = 失敗を記録する（両方なら 3）
    public static uint Query(Guid subCategory) {
        int error;
        uint value = QueryOnce(subCategory, out error);
        if (error == 0) return value;
        if (error != 5 && error != 1314) throw new Win32Exception(error);
        IntPtr token;
        if (!OpenProcessToken(GetCurrentProcess(), 0x0028, out token)) throw new Win32Exception(Marshal.GetLastWin32Error());
        try {
            TokenPrivileges state = new TokenPrivileges();
            state.Count = 1;
            state.Attributes = 2;
            if (!LookupPrivilegeValue(null, "SeSecurityPrivilege", out state.Luid)) throw new Win32Exception(Marshal.GetLastWin32Error());
            TokenPrivileges previous = new TokenPrivileges();
            uint length;
            uint size = (uint)Marshal.SizeOf(typeof(TokenPrivileges));
            if (!AdjustTokenPrivileges(token, false, ref state, size, ref previous, out length)) throw new Win32Exception(Marshal.GetLastWin32Error());
            int adjust = Marshal.GetLastWin32Error();
            try {
                if (adjust != 0) throw new Win32Exception(adjust);
                value = QueryOnce(subCategory, out error);
                if (error != 0) throw new Win32Exception(error);
                return value;
            } finally {
                if (adjust == 0 && previous.Count > 0) {
                    TokenPrivileges ignored = new TokenPrivileges();
                    uint dummy;
                    AdjustTokenPrivileges(token, false, ref previous, size, ref ignored, out dummy);
                }
            }
        } finally {
            CloseHandle(token);
        }
    }
}
'@

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

# 日付や数値の区切り文字が、OS の地域設定によって変わらないようにする
function Format-Invariant([string]$Format) {
    return [string]::Format($Invariant, $Format, [object[]]$args)
}

function Format-Duration([TimeSpan]$Span) {
    if ($Span.TotalSeconds -lt 60) { return [string]::Format($Invariant, '{0:N1}秒', $Span.TotalSeconds) }
    if ($Span.TotalHours -lt 1) { return '{0}分{1}秒' -f $Span.Minutes, $Span.Seconds }
    return '{0}時間{1}分{2}秒' -f [math]::Floor($Span.TotalHours), $Span.Minutes, $Span.Seconds
}

# ログオンの長さ（0:22:39、1 日を超えると 1日 2:03:00）
function Format-Span([TimeSpan]$Span) {
    if ($Span.Ticks -lt 0) { $Span = [TimeSpan]::Zero }
    if ($Span.TotalDays -ge 1) { return Format-Invariant '{0}日 {1}:{2:00}:{3:00}' ([math]::Floor($Span.TotalDays)) $Span.Hours $Span.Minutes $Span.Seconds }
    return Format-Invariant '{0}:{1:00}:{2:00}' $Span.Hours $Span.Minutes $Span.Seconds
}

# .NET のメソッドで起きた例外は MethodInvocationException に包まれて届くため、元の例外を取り出す
function Get-InnerException([Exception]$Exception) {
    while (($Exception -is [Management.Automation.MethodInvocationException] -or $Exception -is [Reflection.TargetInvocationException]) -and $Exception.InnerException) {
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

# 例外の連なり（InnerException）のどこかに、指定した種類の例外があるか
function Test-InnerException([Exception]$Exception, [type]$Type) {
    while ($Exception) {
        if ($Exception -is $Type) { return $true }
        $Exception = $Exception.InnerException
    }
    return $false
}

# CSV の 1 行
function ConvertTo-CsvLine([object[]]$Fields) {
    $parts = [string[]]::new($Fields.Count)
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

# 値を 1 行の文字列にそろえる。途中の改行は空白 1 つに、ほかの制御文字は ? にする（画面や CSV を崩さないため）
function ConvertTo-CellText($Value) {
    if ($null -eq $Value) { return '' }
    $text = ([string]$Value) -replace '\s*[\r\n]+\s*', ' '
    return ($text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '?').Trim()
}

# 画面の表で列をそろえるため、全角文字を 2 文字分の幅として数える（半角カタカナは 1 文字分）。
# 表の値はほとんどが半角だけなので、そのときは 1 文字ずつ数えない（件数が多いときに遅くならないように）
$wideChar = New-Object Text.RegularExpressions.Regex ('[^\u0000-ÿ｡-ﾟ]')
function Get-DisplayWidth([string]$Text) {
    if (-not $wideChar.IsMatch($Text)) { return $Text.Length }
    $width = 0
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int]$ch
        if ($code -lt 0x100 -or ($code -ge 0xFF61 -and $code -le 0xFF9F)) { $width += 1 }
        elseif ($code -ge 0xDC00 -and $code -le 0xDFFF) { }   # サロゲートペアの後半は、前半で 2 文字分を数えた
        else { $width += 2 }
    }
    return $width
}

# 画面の表の値を、半角 $Width 文字分までにする（超える分は … にする）
function Limit-Width([string]$Text, [int]$Width) {
    if ((Get-DisplayWidth $Text) -le $Width) { return $Text }
    $result = ''
    foreach ($ch in $Text.ToCharArray()) {
        if ((Get-DisplayWidth ($result + $ch)) -gt $Width - 2) { break }
        $result += $ch
    }
    return $result + '…'
}

# 見出しと行（文字列の配列）から、列をそろえた表の行を作る。
# summary.txt にはすべての行を書くので、件数が多くても遅くならないよう、値ごとの幅は 1 回だけ数え、関数を呼ぶ回数を減らす
function Format-TextTable([string[]]$Head, [bool[]]$Right, $Rows) {
    $all = @(, $Head) + @($Rows)
    $columns = $Head.Count
    $widths = [int[]]::new($columns)
    $cellWidths = [object[]]::new($all.Count)
    for ($r = 0; $r -lt $all.Count; $r++) {
        $rowWidths = [int[]]::new($columns)
        for ($i = 0; $i -lt $columns; $i++) {
            $text = [string]$all[$r][$i]
            $w = if ($wideChar.IsMatch($text)) { Get-DisplayWidth $text } else { $text.Length }
            $rowWidths[$i] = $w
            if ($w -gt $widths[$i]) { $widths[$i] = $w }
        }
        $cellWidths[$r] = $rowWidths
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $cells = [string[]]::new($columns)
    for ($r = 0; $r -lt $all.Count; $r++) {
        for ($i = 0; $i -lt $columns; $i++) {
            $pad = ' ' * ($widths[$i] - $cellWidths[$r][$i])
            $cells[$i] = if ($Right[$i]) { $pad + [string]$all[$r][$i] } else { [string]$all[$r][$i] + $pad }
        }
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
function Stop-WithQuoteError([string]$Name, [string]$Value, [string]$Example = 'D:\collected') {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        ('値を引用符で囲むときは、最後に \ を付けないでください（例: {0} "{1}"）。' -f $Name, $Example) + '最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# -File で起動すると配列を渡せないため、カンマ区切りの文字列で受ける。各要素の前後の ' と空白も外す
function Split-List([string]$Text) {
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($part in ($Text -split ',')) {
        $item = (Remove-Quote $part).Trim("'").Trim()
        if ($item) { $list.Add($item) }
    }
    return , $list
}

# 日時か、2h・30m・3d のような今からの時間を読む。Relative は「（7 日）」のように添える文字
function ConvertTo-PointInTime([string]$Text, [string]$Name) {
    $value = (Remove-Quote $Text).Trim("'")
    if ($value -match '^(\d{1,5})\s*([mhd])$') {
        $amount = [int]$Matches[1]
        switch ($Matches[2]) {
            'm' { return @{ Time = $startTime.AddMinutes(-$amount); Relative = '{0} 分' -f $amount } }
            'h' { return @{ Time = $startTime.AddHours(-$amount); Relative = '{0} 時間' -f $amount } }
            'd' { return @{ Time = $startTime.AddDays(-$amount); Relative = '{0} 日' -f $amount } }
        }
    }
    $formats = 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd', 'yyyy/MM/dd HH:mm:ss', 'yyyy/MM/dd HH:mm', 'yyyy/MM/dd',
               'yyyy-M-d H:mm:ss', 'yyyy-M-d H:mm', 'yyyy-M-d', 'yyyy/M/d H:mm:ss', 'yyyy/M/d H:mm', 'yyyy/M/d'
    $parsed = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($value, [string[]]$formats, $Invariant, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
        Stop-WithError ('{0} の日時を読めません: {1}（"2026-11-10 09:00" のような日時か、2h・30m・3d のような今からの時間で指定してください）' -f $Name, $value)
    }
    return @{ Time = $parsed; Relative = '' }
}

# イベントログ（または .evtx）を読む。Get-WinEvent は使わない（使わないメッセージを 1 件ごとに組み立てるため、記録が多いと
# 何十倍も遅い。Path をワイルドカードとして読む問題もない）。Source の Name はログの名前かファイルのパス
function New-EventReader($Source, [string]$Query, [bool]$Reverse = $false) {
    $pathType = if ($Source.IsFile) { [Diagnostics.Eventing.Reader.PathType]::FilePath } else { [Diagnostics.Eventing.Reader.PathType]::LogName }
    $eventQuery = New-Object Diagnostics.Eventing.Reader.EventLogQuery($Source.Name, $pathType, $Query)
    $eventQuery.ReverseDirection = $Reverse
    return New-Object Diagnostics.Eventing.Reader.EventLogReader($eventQuery)
}

# 条件に合う最初の記録（$Reverse なら最後の記録）の、日時・ログの名前・コンピューター。ないときは $null
function Read-FirstEvent($Source, [string]$Query, [bool]$Reverse = $false) {
    $reader = New-EventReader $Source $Query $Reverse
    try {
        $record = $reader.ReadEvent()
        if ($null -eq $record) { return $null }
        try { return @{ Time = $record.TimeCreated; Log = [string]$record.LogName; Machine = [string]$record.MachineName } } finally { $record.Dispose() }
    } finally {
        $reader.Dispose()
    }
}

# 読み取りの条件（XPath）の日時の部分。From と To は省略できる
function Get-TimeCondition([string]$From, [string]$To) {
    $parts = @()
    if ($From) { $parts += "@SystemTime>='$From'" }
    if ($To) { $parts += "@SystemTime<='$To'" }
    if ($parts.Count -eq 0) { return '' }
    return ' and TimeCreated[' + ($parts -join ' and ') + ']'
}

# イベントログの側で絞るための式（XPath）に書く日時
function ConvertTo-QueryTime([datetime]$Time) {
    return $Time.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ', $Invariant)
}

# リパースポイントのうち、リンクと判定できたものの種類を返す（OneDrive のフォルダもリパースポイントなので、それ以外は $null）
function Get-LinkName($Entry) {
    if (($Entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) { return $null }
    # Windows が互換用に作ったリンク（All Users、Application Data など）は、5.1 では種類を判定できない。どれも隠し・システム属性付き
    $hiddenSystem = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System
    if ($Entry -is [IO.DirectoryInfo] -and ($Entry.Attributes -band $hiddenSystem) -eq $hiddenSystem) { return 'Windows の互換用リンク' }
    $linkType = $null
    try { $linkType = $Entry.LinkType } catch { }
    if ($linkType -eq 'Junction') { return 'ジャンクション' }
    if ($linkType -eq 'SymbolicLink') { return 'シンボリックリンク' }
    return $null
}

# コンピューターの名前（server01.corp.example.local）を、組にするためのキー（SERVER01）にする
function Get-ComputerKey([string]$Name) {
    return $Name.Split('.')[0].ToUpperInvariant()
}

# 名前で並べるときの順を、5.1 と 7 でそろえるためのキー。Sort-Object の文字列の比べ方は 5.1（.NET Framework）と 7 で違い、
# ハイフンなどの記号の扱いが変わる（CORP\admin と CORP-FS\admin の順が入れ替わる）。
# 大文字にそろえた文字の番号を 16 進数で並べた文字列にすると、どちらでも文字の番号の順になる
$sortKeys = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
function ConvertTo-SortKey([string]$Text) {
    $key = $null
    if ($sortKeys.TryGetValue($Text, [ref]$key)) { return $key }
    $builder = New-Object Text.StringBuilder ($Text.Length * 4)
    foreach ($ch in $Text.ToUpperInvariant().ToCharArray()) { [void]$builder.Append(([int]$ch).ToString('X4', $Invariant)) }
    $key = $builder.ToString()
    $sortKeys[$Text] = $key
    return $key
}

# 記録の数値（ログオン ID、状態コードなど）。数値でも "0x3e7" のような文字でも届くので、どちらも読む
function ConvertTo-UInt64Value($Value) {
    if ($null -eq $Value) { return [uint64]0 }
    if ($Value -is [string]) {
        $text = $Value.Trim()
        if ($text -match '^0[xX]([0-9A-Fa-f]{1,16})$') { return [Convert]::ToUInt64($Matches[1], 16) }
        $number = [uint64]0
        if ([uint64]::TryParse($text, [ref]$number)) { return $number }
        return [uint64]0
    }
    # 負の Int32（0xC000006D など）は、ビットのまま符号なしにする
    if ($Value -is [int]) { return [uint64][BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$Value), 0) }
    if ($Value -is [long]) { return [BitConverter]::ToUInt64([BitConverter]::GetBytes([long]$Value), 0) }
    return [uint64]$Value
}

# ログオンの種類の表示名
function Get-TypeLabel([int]$Type, [string]$Process) {
    switch ($Type) {
        { $_ -eq 2 -or $_ -eq 11 } { if ($Process.Trim() -like 'seclogo*') { return '対話（runas）' }; return '対話' }
        { $_ -eq 10 -or $_ -eq 12 } { return 'リモートデスクトップ' }
        { $_ -eq 7 -or $_ -eq 13 } { return 'ロック解除' }
        3 { return 'ネットワーク' }
        8 { return 'ネットワーク（平文）' }
        4 { return 'バッチ' }
        5 { return 'サービス' }
        9 { return '別の資格情報' }
    }
    return '種類 {0}' -f $Type
}

# OS の内部で使うアカウント（SYSTEM、LOCAL SERVICE、NETWORK SERVICE、ANONYMOUS LOGON、DWM-1、UMFD-0）と、コンピューターのアカウント
function Test-SystemAccount([string]$Sid, [string]$Name) {
    if ($Sid -eq 'S-1-5-18' -or $Sid -eq 'S-1-5-19' -or $Sid -eq 'S-1-5-20' -or $Sid -eq 'S-1-5-7') { return $true }
    if ($Sid.StartsWith('S-1-5-90-', [StringComparison]::Ordinal) -or $Sid.StartsWith('S-1-5-96-', [StringComparison]::Ordinal)) { return $true }
    return $Name.EndsWith('$', [StringComparison]::Ordinal)
}

# 「ドメイン\ユーザー名」。ドメインがないときは、ユーザー名だけ
function Join-AccountName([string]$Domain, [string]$Name) {
    $d = ConvertTo-CellText $Domain
    $n = ConvertTo-CellText $Name
    if ($d -and $d -ne '-') { return $d + '\' + $n }
    return $n
}

# 接続元の IP アドレス。記録がない（-）ときは空。コンソールのログオンなどに記録される自分自身のアドレスも空にする
function ConvertTo-SourceAddress([string]$Address, [int]$Type) {
    $text = ConvertTo-CellText $Address
    if ($text -eq '-') { return '' }
    if (($Type -eq 2 -or $Type -eq 11 -or $Type -eq 7 -or $Type -eq 13) -and ($text -eq '127.0.0.1' -or $text -eq '::1')) { return '' }
    return $text
}

# セッションのログの Address が IP アドレスなら、その文字（コンソールは「ローカル」「LOCAL」や空で記録される）
function Get-LsmAddress([string]$Address) {
    $text = ConvertTo-CellText $Address
    $ip = $null
    if ($text -and [Net.IPAddress]::TryParse($text, [ref]$ip)) { return $text }
    return ''
}

# ログオンの失敗の理由。Status が「ユーザー名かパスワードが違う」などのまとめのコードのときは、SubStatus で詳しく見る
function Get-FailureReason([uint64]$Status, [uint64]$SubStatus) {
    $code = $Status.ToString('X8', $Invariant)
    if ($SummaryStatus -contains $code -and $SubStatus -ne 0) { $code = $SubStatus.ToString('X8', $Invariant) }
    if ($FailureReasons.ContainsKey($code)) { return $FailureReasons[$code] }
    return 'その他（0x{0}）' -f $code
}

# -User のパターン。* だけを任意の文字列として使うため、ほかのワイルドカードの文字はエスケープする
# （5.1 の WildcardPattern.Escape はバッククォートをエスケープしない）
function New-UserPattern([string]$Text) {
    $escaped = $Text.Replace('`', '``').Replace('[', '`[').Replace(']', '`]').Replace('?', '`?')
    return @{ Full = $Text.Contains('\'); Pattern = New-Object Management.Automation.WildcardPattern($escaped, [Management.Automation.WildcardOptions]::IgnoreCase) }
}

# -User に当たるか。名前はいくつか持てる（Microsoft アカウントは、ログオンとログオフで名前が違う）
function Test-UserMatch($Names) {
    if ($userPatterns.Count -eq 0) { return $true }
    foreach ($name in $Names) {
        if (-not $name) { continue }
        $short = $name.Substring($name.LastIndexOf('\') + 1)
        foreach ($pattern in $userPatterns) {
            if ($pattern.Full) { if ($pattern.Pattern.IsMatch($name)) { return $true } }
            elseif ($pattern.Pattern.IsMatch($short)) { return $true }
        }
    }
    return $false
}

function Format-Time($Time) {
    if ($null -eq $Time) { return '' }
    return ([datetime]$Time).ToString($TimeFormat, $Invariant)
}

function Format-ScreenTime($Time) {
    if ($null -eq $Time) { return '' }
    return ([datetime]$Time).ToString($ScreenTime, $Invariant)
}

# 画面のログオフの時刻。ログオンと同じ日なら時刻だけ、違う日なら月日と時刻（表を 120 桁に収めるため）
function Format-LogoffTime([datetime]$Logon, [datetime]$Logoff) {
    if ($Logon.Date -eq $Logoff.Date) { return $Logoff.ToString('HH:mm', $Invariant) }
    if ($Logon.Year -eq $Logoff.Year) { return $Logoff.ToString('MM-dd HH:mm', $Invariant) }
    return $Logoff.ToString($ScreenTime, $Invariant)
}

function Format-LogonId([uint64]$Id) {
    return '0x' + $Id.ToString('x', $Invariant)
}

#endregion

#region 引数の確認

$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

if ($NoSave -and $OutputFolder) { Stop-WithError '-NoSave のときは、-OutputFolder は指定できません' }

# コマンドプロンプトで "CORP\" のように最後を \" にすると、後ろの引数まで値に入る（-NoSave が効かないまま保存するなど）。
# " はユーザー名にも日時にも使わないので、残っていたら止める
foreach ($pair in @(@('-User', $User, 'CORP\yamada'), @('-LogonType', $LogonType, 'Interactive'), @('-Since', $Since, '2026-11-10 09:00'), @('-Until', $Until, '2026-11-10 18:00'))) {
    if ($pair[1] -and (Remove-Quote $pair[1]).Contains('"')) { Stop-WithQuoteError $pair[0] (Remove-Quote $pair[1]) $pair[2] }
}

$startTime = Get-Date
$sinceInfo = ConvertTo-PointInTime $Since '-Since'
$untilInfo = $null
if ($Until) { $untilInfo = ConvertTo-PointInTime $Until '-Until' }
$endTime = if ($untilInfo) { $untilInfo.Time } else { $startTime }
if ($sinceInfo.Time -ge $endTime) {
    $endLabel = if ($untilInfo) { '-Until の日時' } else { '今' }
    Stop-WithError (Format-Invariant '-Since の日時（{0:yyyy-MM-dd HH:mm}）が、{1}（{2:yyyy-MM-dd HH:mm}）と同じか、それより後になっています' $sinceInfo.Time $endLabel $endTime)
}

# ログオンの種類
$typeNames = @{}
foreach ($key in $TypeGroups.Keys) { $typeNames[$key.ToLowerInvariant()] = $key; $typeNames[$TypeGroups[$key].Label] = $key }
$selectedGroups = [System.Collections.Generic.List[string]]::new()
$allTypes = $false
foreach ($item in (Split-List $LogonType)) {
    $lower = $item.ToLowerInvariant()
    if ($lower -eq 'all' -or $item -eq 'すべて') { $allTypes = $true; continue }
    if (-not $typeNames.ContainsKey($lower) -and -not $typeNames.ContainsKey($item)) {
        $names = @(foreach ($key in $TypeGroups.Keys) { '{0}（{1}）' -f $key, $TypeGroups[$key].Label }) -join '、'
        Stop-WithError ('-LogonType に使えない種類があります: {0}' -f $item + "`r`n" + '使える種類: ' + $names + '、All（すべて）')
    }
    $key = if ($typeNames.ContainsKey($lower)) { $typeNames[$lower] } else { $typeNames[$item] }
    if (-not $selectedGroups.Contains($key)) { $selectedGroups.Add($key) }
}
if ($allTypes) { $selectedGroups.Clear(); foreach ($key in $TypeGroups.Keys) { $selectedGroups.Add($key) } }
if ($selectedGroups.Count -eq 0) { Stop-WithError '-LogonType を指定してください' }
$selectedTypes = New-Object 'System.Collections.Generic.HashSet[int]'
foreach ($key in $selectedGroups) { foreach ($t in $TypeGroups[$key].Types) { [void]$selectedTypes.Add($t) } }
$unlockSelected = $selectedGroups.Contains('Unlock')
$sessionSelected = $selectedGroups.Contains('Interactive') -or $selectedGroups.Contains('RemoteDesktop')
# ロック解除やサインインを見分けるため、対話・リモートデスクトップ・ロック解除のどれかを取り出すときは、3 つとも読む
$readTypes = New-Object 'System.Collections.Generic.HashSet[int]'
foreach ($t in $selectedTypes) { [void]$readTypes.Add($t) }
if ($sessionSelected -or $unlockSelected) { foreach ($t in @($SessionTypes) + @($UnlockTypes)) { [void]$readTypes.Add($t) } }
# ログオフの記録（4634）は、キャッシュの種類を元の種類で記録する（11 のログオンが 2 のログオフになる）
$logoffTypes = New-Object 'System.Collections.Generic.HashSet[int]'
foreach ($t in $readTypes) {
    [void]$logoffTypes.Add($t)
    if ($t -eq 11) { [void]$logoffTypes.Add(2) } elseif ($t -eq 12) { [void]$logoffTypes.Add(10) } elseif ($t -eq 13) { [void]$logoffTypes.Add(7) }
}
$typeLabelText = if ($allTypes) { 'すべて' } else { (@($selectedGroups | ForEach-Object { $TypeGroups[$_].Label }) -join ', ') }

# ユーザー
$userItems = Split-List $User
$userPatterns = [System.Collections.Generic.List[object]]::new()
foreach ($item in $userItems) { $userPatterns.Add((New-UserPattern $item)) }

$outText = $null
if ($OutputFolder) {
    $outText = Remove-Quote $OutputFolder
    if ($outText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outText }
    if ($outText -match '^(\\\\\?\\)?[A-Za-z]:$') { $outText += '\' }
    $outText = Resolve-StartPath $outText
}

$skipped   = [System.Collections.Generic.List[object]]::new()
$skipKinds = [ordered]@{ '権限なし' = 0; 'ログがない' = 0; 'ファイルを読めない' = 0; '対象外のログ' = 0; 'リンク' = 0; 'その他のエラー' = 0 }
function Add-Skip([string]$Reason, [string]$Target, [string]$Detail) {
    $skipped.Add([object[]]@($Reason, $Target, $Detail))
    $skipKinds[$Reason]++
}

# 読む対象。Kind は Security かセッションのログ（Lsm）。Oldest はログに残っているいちばん古い記録の日時
$sources = [System.Collections.Generic.List[object]]::new()
$useFiles = [bool]$Path
if ($useFiles) {
    $files = [System.Collections.Generic.List[string]]::new()
    foreach ($item in (Split-List $Path)) {
        if ($item.Contains('"')) { Stop-WithQuoteError '-Path' $item }
        $full = ConvertTo-DisplayPath (Resolve-StartPath $item)
        try {
            $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($full)
            $full = [IO.Path]::GetFullPath($full)
        } catch { }
        if ([IO.File]::Exists((ConvertTo-LongPath $full))) {
            $files.Add($full)
        } elseif ([IO.Directory]::Exists((ConvertTo-LongPath $full))) {
            # フォルダの下の .evtx を探す。リンクはたどらず、読めないフォルダは飛ばす
            # （ドライブのルートは、最後の \ を外すと \\?\D: になって読めない）
            $stack = New-Object 'System.Collections.Generic.Stack[string]'
            $top = $full.TrimEnd('\')
            if ($top -match '^[A-Za-z]:$') { $top += '\' }
            $stack.Push((ConvertTo-LongPath $top))
            while ($stack.Count -gt 0) {
                $dir = $stack.Pop()
                try { $entries = (New-Object IO.DirectoryInfo $dir).GetFileSystemInfos() } catch {
                    $ex = Get-InnerException $_.Exception
                    $reason = if ($ex -is [UnauthorizedAccessException]) { '権限なし' } else { 'その他のエラー' }
                    Add-Skip $reason (ConvertTo-DisplayPath $dir) (Get-ErrorText $ex)
                    continue
                }
                # ファイルは名前の順に加え、サブフォルダは名前の順にたどれるよう、逆の順に積む
                $subDirs = [System.Collections.Generic.List[string]]::new()
                foreach ($entry in ($entries | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_.Name } })) {
                    $linkName = Get-LinkName $entry
                    if ($linkName -and $entry -is [IO.DirectoryInfo]) {
                        Add-Skip 'リンク' (ConvertTo-DisplayPath $entry.FullName) $linkName
                        continue
                    }
                    if ($entry -is [IO.DirectoryInfo]) { $subDirs.Add($entry.FullName) }
                    elseif ($entry.Extension -ieq '.evtx') { $files.Add((ConvertTo-DisplayPath $entry.FullName)) }
                }
                for ($i = $subDirs.Count - 1; $i -ge 0; $i--) { $stack.Push($subDirs[$i]) }
            }
        } else {
            Add-Skip 'ファイルを読めない' $full '見つかりません'
        }
    }
    # 何のログかは、ファイルの最初の記録で見分ける
    foreach ($file in $files) {
        Write-Progress -Activity $ToolTitle -Status 'ファイルを調べています' -CurrentOperation $file
        # 260 文字以上のパスは \\?\ を付けないと読めない
        $source = @{ Kind = ''; Label = $file; Name = $(if ($file.Length -ge $LongPathLength) { ConvertTo-LongPath $file } else { $file }); IsFile = $true; Oldest = $null; Computer = ''; ReadFrom = $null; Failed = $false }
        try {
            $first = Read-FirstEvent $source '*'
        } catch {
            if (Test-InnerException $_.Exception ([UnauthorizedAccessException])) { Add-Skip '権限なし' $file 'アクセスが拒否されました'; continue }
            Add-Skip 'ファイルを読めない' $file (Get-ErrorText (Get-InnerException $_.Exception))
            continue
        }
        if ($null -eq $first) { Add-Skip '対象外のログ' $file '記録がありません'; continue }
        $kind = if ($first.Log -ieq $SecurityLog) { 'Security' } elseif ($first.Log -ieq $LsmLog) { 'Lsm' } else { $null }
        if (-not $kind) { Add-Skip '対象外のログ' $file ('Security でもセッションのログでもありません（{0}）' -f $first.Log); continue }
        $source.Kind = $kind
        $source.Oldest = $first.Time
        # コンピューターの名前は、いちばん新しい記録のもの（構築の後に名前を変えたサーバーは、古い記録が前の名前になっているため）。
        # 最後のほうが壊れていて新しい順に読めないときは、いちばん古い記録の名前を使う
        $machine = $first.Machine
        try {
            $last = Read-FirstEvent $source '*' $true
            if ($last) { $machine = $last.Machine }
        } catch { }
        $source.Computer = Get-ComputerKey $machine
        $sources.Add($source)
    }
    Write-Progress -Activity $ToolTitle -Completed
    if ($sources.Count -eq 0) {
        # 打ち間違いか、権限か、形式の違うファイルかが分かるように、理由を添える
        $reasons = @($skipped | Select-Object -First 5 | ForEach-Object { '  {0}: {1}（{2}）' -f $_[1], $_[0], $_[2] })
        $message = '読める .evtx ファイルがありません: {0}' -f $Path
        if ($reasons.Count -gt 0) { $message += "`r`n" + ($reasons -join "`r`n") }
        if ($skipped.Count -gt $reasons.Count) { $message += "`r`n" + (Format-Invariant '  ほか {0:N0} 件' ($skipped.Count - $reasons.Count)) }
        Stop-WithError $message
    }
    # セッションのログだけでは、ログオンもログオンの失敗も分からない（「該当なし」と誤解されないように止める）
    if (@($sources | Where-Object { $_.Kind -eq 'Security' }).Count -eq 0) {
        Stop-WithError ('-Path に、Security ログの .evtx がありません: {0}' -f $Path + "`r`n" + 'ログオンとログオンの失敗は Security ログにあります。セッションのログ（切断・再接続）だけでは調べられません')
    }
} else {
    # この PC の Security ログ。読めなければ、集め始める前に止める
    $source = @{ Kind = 'Security'; Label = $SecurityLog; Name = $SecurityLog; IsFile = $false; Oldest = $null; Computer = (Get-ComputerKey $machineName); ReadFrom = $null; Failed = $false }
    try {
        $first = Read-FirstEvent $source '*'
    } catch {
        if (Test-InnerException $_.Exception ([UnauthorizedAccessException])) {
            Stop-WithError ('Security ログを読む権限がありません。管理者として実行してください' + "`r`n" + 'ほかの PC から集めた .evtx ファイルは、-Path で読めます（管理者でなくても読めます）')
        }
        Stop-WithError ('Security ログを読めません: {0}' -f (Get-ErrorText (Get-InnerException $_.Exception)))
    }
    if ($first) { $source.Oldest = $first.Time }
    $sources.Add($source)
    # リモートデスクトップのセッションのログ。読めなくても続ける
    if ($sessionSelected) {
        $source = @{ Kind = 'Lsm'; Label = $LsmLog; Name = $LsmLog; IsFile = $false; Oldest = $null; Computer = (Get-ComputerKey $machineName); ReadFrom = $null; Failed = $false }
        try {
            $first = Read-FirstEvent $source '*'
            if ($first) { $source.Oldest = $first.Time }
            $sources.Add($source)
        } catch {
            if (Test-InnerException $_.Exception ([Diagnostics.Eventing.Reader.EventLogNotFoundException])) {
                Add-Skip 'ログがない' $LsmLog 'この PC にはないログです（切断・再接続は分かりません）'
            } elseif (Test-InnerException $_.Exception ([UnauthorizedAccessException])) {
                Add-Skip '権限なし' $LsmLog 'アクセスが拒否されました（切断・再接続は分かりません）'
            } else {
                Add-Skip 'その他のエラー' $LsmLog (Get-ErrorText (Get-InnerException $_.Exception))
            }
        }
    }
}

#endregion

#region 保存先の準備

# 読み終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir = $null
if (-not $NoSave) {
    $baseDir = if ($outText) { $outText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        # \\?\ 付きのパスでは末尾のピリオドや空白が残るため、先に通常の書き方で整える（260 文字を超えると例外になるので、そのときはそのまま）
        try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
        $runName = Format-Invariant 'LogonHistory_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while ([IO.Directory]::Exists((ConvertTo-LongPath $runDir)) -or [IO.File]::Exists((ConvertTo-LongPath $runDir))) {
            $suffix++
            $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix)
        }
        [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $runDir))
    } catch {
        $ex = Get-InnerException $_.Exception
        $message = '保存先に書き込めません: {0}（{1}）' -f $baseDir, (Get-ErrorText $ex)
        # 管理者として開いた PowerShell からは、通常の権限で割り当てたネットワークドライブが見えない
        if ($isAdmin -and $ex -is [Management.Automation.DriveNotFoundException]) {
            $message += "`r`n" + '管理者として実行した PowerShell からは、通常の権限で割り当てたネットワークドライブは見えません。\\server\share の形で指定してください'
        }
        Stop-WithError ($message + "`r`n" + '-NoSave を付けると、画面に表示だけします')
    }
}

#endregion

$finished   = $false
$fatal      = $false
$unexpected = $false
$saveError  = $null
$sw         = [Diagnostics.Stopwatch]::StartNew()
try {
    #region 読み取り

    # 期間に重なるログオン（期間より前にログオンして、期間中も続いていたもの）を出すため、ログオンとログオフは
    # 期間の前の最後の OS の起動から読む（ログオンは、OS を起動し直すと終わるため）。起動の記録がなければ、ログの最初から読む。
    # ログオンの失敗は期間の中だけ。ログオフと OS の起動は、期間の後のものも読む（ログオンが、いつ終わったかを知るため）
    $sinceText = ConvertTo-QueryTime $sinceInfo.Time
    $untilText = ConvertTo-QueryTime $endTime
    $logonTypeCond = ''
    $logoffTypeCond = ''
    if (-not $allTypes) {
        $logonTypeCond  = ' and EventData[' + (@($readTypes | Sort-Object | ForEach-Object { "Data[@Name='LogonType']='$_'" }) -join ' or ') + ']'
        $logoffTypeCond = ' and EventData[' + (@($logoffTypes | Sort-Object | ForEach-Object { "Data[@Name='LogonType']='$_'" }) -join ' or ') + ']'
    }
    function Get-SecurityQuery([string]$From) {
        return "*[(System[EventID=4624$(Get-TimeCondition $From $untilText)]$logonTypeCond)" +
            " or (System[EventID=4634$(Get-TimeCondition $From '')]$logoffTypeCond)" +
            " or System[(EventID=4647 or EventID=4608 or EventID=1102)$(Get-TimeCondition $From '')]" +
            " or System[EventID=4625$(Get-TimeCondition $sinceText $untilText)]]"
    }
    # セッションのログは小さい（既定で 1 MB）ので、すべて読む
    $lsmQuery = '*[System[EventID=21 or EventID=23 or EventID=24 or EventID=25]]'

    # 記録の値は、名前で取り出す（記録の版によって値の並びが違うため）
    function New-Selector([string[]]$Names) {
        return New-Object Diagnostics.Eventing.Reader.EventLogPropertySelector (, [string[]]@($Names | ForEach-Object { "Event/EventData/Data[@Name='$_']" }))
    }
    $selectLogon  = New-Selector @('TargetUserSid', 'TargetUserName', 'TargetDomainName', 'TargetLogonId', 'LogonType', 'TargetLinkedLogonId', 'ElevatedToken', 'IpAddress', 'WorkstationName', 'LogonProcessName', 'AuthenticationPackageName')
    $selectLogoff = New-Selector @('TargetUserSid', 'TargetUserName', 'TargetDomainName', 'TargetLogonId', 'LogonType')
    $selectFail   = New-Selector @('TargetUserName', 'TargetDomainName', 'LogonType', 'IpAddress', 'WorkstationName', 'Status', 'SubStatus', 'AuthenticationPackageName')
    $selectClear  = New-Object Diagnostics.Eventing.Reader.EventLogPropertySelector (, [string[]]@('Event/UserData/*/SubjectDomainName', 'Event/UserData/*/SubjectUserName'))
    $selectLsm    = New-Object Diagnostics.Eventing.Reader.EventLogPropertySelector (, [string[]]@('Event/UserData/*/User', 'Event/UserData/*/SessionID', 'Event/UserData/*/Address'))

    $logons   = [System.Collections.Generic.List[object]]::new()
    $logoffs  = @{}   # コンピューター|ログオン ID → ログオフの記録（時刻とイベント ID）の一覧
    $boots    = @{}   # コンピューター → OS の起動の時刻の一覧
    $failures = [System.Collections.Generic.List[object]]::new()
    $clears   = [System.Collections.Generic.List[object]]::new()
    $lsmEvents = [System.Collections.Generic.List[object]]::new()
    $sidNames = @{}   # コンピューター|SID → 記録に出てきた名前（Microsoft アカウントは、ログオンとログオフで名前が違う）
    $seen     = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $bootNames = @{}   # コンピューターのキー|OS の起動の時刻 → その記録に書かれたコンピューターの名前（ほかの行と同じく、記録のままを書く）

    function Add-SidName([string]$Computer, [string]$Sid, [string]$Name) {
        if (-not $Sid -or -not $Name) { return }
        $key = $Computer + '|' + $Sid
        if (-not $sidNames.ContainsKey($key)) { $sidNames[$key] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
        [void]$sidNames[$key].Add($Name)
    }

    # 記録を 1 件、種類ごとの一覧に入れる。
    # 組にするときのコンピューターは、読んだログ（ファイル）ごとに決める。記録に書かれた名前は、名前を変える前の記録だと前の名前なので、
    # 記録ごとに見ると、名前を変える前のログオンから、変えた後の OS の起動が見えなくなる
    function Add-EventRecord($Record, $Source) {
        $computerText = [string]$Record.MachineName
        $computer = $Source.Computer
        $time = $Record.TimeCreated
        # 同じ記録を、重なった .evtx から 2 回読まないようにする。ログを消去すると記録の番号が振り直されることがあるので、日時も含める。
        # ファイルごとに決めた名前は、重なったファイルでも同じとは限らないので、ここは記録に書かれた名前で見る
        if (-not $seen.Add(('{0}|{1}|{2}|{3}' -f $computerText, $Record.LogName, $Record.RecordId, $time.Ticks))) { return }
        $recordId = [long]$Record.RecordId
        switch ([int]$Record.Id) {
            4624 {
                $v = $Record.GetPropertyValues($selectLogon)
                $sid = [string]$v[0]
                $name = Join-AccountName ([string]$v[2]) ([string]$v[1])
                if (Test-SystemAccount $sid ([string]$v[1])) { return }
                Add-SidName $computer $sid $name
                $elevated = $null
                if ([string]$v[6] -eq '%%1842') { $elevated = $true } elseif ([string]$v[6] -eq '%%1843') { $elevated = $false }
                $type = [int](ConvertTo-UInt64Value $v[4])
                # Index は並べ替えた後の番号。後から Add-Member で足すと、件数が多いときに遅いので、ここで用意しておく
                $logons.Add([pscustomobject]@{
                    Computer = $computer; ComputerText = $computerText; Time = $time; RecordId = $recordId; Sid = $sid; Name = $name; Type = $type
                    Id = (ConvertTo-UInt64Value $v[3]); Linked = (ConvertTo-UInt64Value $v[5]); Elevated = $elevated
                    Ip = (ConvertTo-SourceAddress ([string]$v[7]) $type); Workstation = (ConvertTo-SourceAddress ([string]$v[8]) 0)
                    Process = (ConvertTo-CellText $v[9]); Auth = (ConvertTo-CellText $v[10]); Index = 0
                })
            }
            { $_ -eq 4634 -or $_ -eq 4647 } {
                $v = $Record.GetPropertyValues($selectLogoff)
                $sid = [string]$v[0]
                $name = Join-AccountName ([string]$v[2]) ([string]$v[1])
                Add-SidName $computer $sid $name
                $id = ConvertTo-UInt64Value $v[3]
                $key = $computer + '|' + $id
                if (-not $logoffs.ContainsKey($key)) { $logoffs[$key] = [System.Collections.Generic.List[object]]::new() }
                # 4647 には種類がない（0 にする）。Used は、ログオンと組にできたかの印（組にできなかったものは、ログオンの記録が残っていない）
                $type = if ([int]$Record.Id -eq 4634) { [int](ConvertTo-UInt64Value $v[4]) } else { 0 }
                $logoffs[$key].Add([pscustomobject]@{
                    Time = $time; EventId = [int]$Record.Id; Computer = $computer; ComputerText = $computerText; Sid = $sid; Name = $name; Id = $id; Type = $type
                    System = (Test-SystemAccount $sid ([string]$v[1])); Used = $false
                })
            }
            4608 {
                if (-not $boots.ContainsKey($computer)) { $boots[$computer] = [System.Collections.Generic.List[datetime]]::new() }
                $boots[$computer].Add($time)
                $bootNames[$computer + '|' + $time.Ticks] = $computerText
            }
            1102 {
                $v = $Record.GetPropertyValues($selectClear)
                $clears.Add([pscustomobject]@{ Computer = $computer; ComputerText = $computerText; Time = $time; By = (Join-AccountName ([string]$v[0]) ([string]$v[1])) })
            }
            4625 {
                $v = $Record.GetPropertyValues($selectFail)
                $type = [int](ConvertTo-UInt64Value $v[2])
                $failures.Add([pscustomobject]@{
                    Computer = $computer; ComputerText = $computerText; Time = $time; RecordId = $recordId; Name = (Join-AccountName ([string]$v[1]) ([string]$v[0])); Type = $type
                    Ip = (ConvertTo-SourceAddress ([string]$v[3]) $type); Workstation = (ConvertTo-SourceAddress ([string]$v[4]) 0)
                    Status = (ConvertTo-UInt64Value $v[5]); SubStatus = (ConvertTo-UInt64Value $v[6]); Auth = (ConvertTo-CellText $v[7]); Reason = ''
                })
            }
            { $_ -eq 21 -or $_ -eq 23 -or $_ -eq 24 -or $_ -eq 25 } {
                $v = $Record.GetPropertyValues($selectLsm)
                $lsmEvents.Add([pscustomobject]@{
                    Computer = $computer; ComputerText = $computerText; Time = $time; RecordId = $recordId; Id = [int]$Record.Id; User = (ConvertTo-CellText $v[0])
                    SessionId = (ConvertTo-CellText $v[1]); Address = (ConvertTo-CellText $v[2]); Ip = (Get-LsmAddress ([string]$v[2]))
                })
            }
        }
    }

    $readCount = 0
    $nextProgress = 0L
    foreach ($source in $sources) {
        $status = if ($source.Kind -eq 'Security') { 'Security ログを読んでいます' } else { 'セッションのログを読んでいます' }
        Write-Progress -Activity $ToolTitle -Status $status -CurrentOperation $source.Label
        $countBefore = $readCount
        $reader = $null
        try {
            if ($source.Kind -eq 'Security') {
                # 期間の前の最後の OS の起動（新しい順に探して、最初の 1 件）。最後のほうが壊れた .evtx では新しい順に読めないので、
                # 失敗したら最初から（古い順に、壊れた場所の手前まで）読む
                try {
                    $boot = Read-FirstEvent $source "*[System[EventID=4608 and TimeCreated[@SystemTime<'$sinceText']]]" $true
                    if ($boot) { $source.ReadFrom = $boot.Time }
                } catch {
                    $source.ReadFrom = $null
                }
                $query = Get-SecurityQuery $(if ($source.ReadFrom) { ConvertTo-QueryTime $source.ReadFrom } else { '' })
            } else {
                $query = $lsmQuery
            }
            # 古い順に読む（壊れた .evtx でも、壊れた場所の手前まで読めるため）
            $reader = New-EventReader $source $query
            while ($true) {
                $record = $reader.ReadEvent()
                if ($null -eq $record) { break }
                try { Add-EventRecord $record $source } finally { $record.Dispose() }
                $readCount++
                if ($sw.ElapsedMilliseconds -ge $nextProgress) {
                    Write-Progress -Activity $ToolTitle -Status (Format-Invariant '{0}（記録 {1:N0} 件）' $status $readCount) -CurrentOperation $source.Label
                    $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
                }
            }
        } catch {
            $ex = Get-InnerException $_.Exception
            # 1 件も読めなかった Security のログは、後で「すべて読めなかった」かを判断するために印を付ける
            if ($readCount -eq $countBefore) { $source.Failed = $true }
            if (Test-InnerException $_.Exception ([UnauthorizedAccessException])) {
                Add-Skip '権限なし' $source.Label 'アクセスが拒否されました'
            } elseif (Test-InnerException $_.Exception ([Diagnostics.Eventing.Reader.EventLogNotFoundException])) {
                Add-Skip 'ログがない' $source.Label (Get-ErrorText $ex)
            } else {
                $detail = Get-ErrorText $ex
                if ($readCount -gt $countBefore) { $detail = Format-Invariant '途中まで読んだ（記録 {0:N0} 件は結果に含む）: {1}' ($readCount - $countBefore) $detail }
                Add-Skip $(if ($useFiles) { 'ファイルを読めない' } else { 'その他のエラー' }) $source.Label $detail
            }
        } finally {
            if ($reader) { $reader.Dispose() }
        }
    }

    # Security ログを 1 件も読めなかったときは、「該当なし」と区別して止める（作った保存先のフォルダは finally で消す）
    if (@($sources | Where-Object { $_.Kind -eq 'Security' -and -not $_.Failed }).Count -eq 0) {
        $reasons = @($skipped | Select-Object -First 5 | ForEach-Object { '  {0}: {1}（{2}）' -f $_[1], $_[0], $_[2] })
        $fatal = $true
        $finished = $true
        Stop-WithError (('Security ログを読めませんでした' + "`r`n") + ($reasons -join "`r`n"))
    }

    # OS の起動の時刻は、複数のファイルから読むと順番が混ざるので並べ替える（次の起動を探すときに、古い順で見るため）
    foreach ($computer in @($boots.Keys)) { $boots[$computer].Sort() }

    # この PC のログオンで、今も続いているもの（ログオフの記録がないときに「ログオン中」と判定するため）
    $liveLogons = New-Object 'System.Collections.Generic.HashSet[uint64]'
    if (-not $useFiles) {
        try {
            foreach ($session in @(Get-CimInstance -ClassName Win32_LogonSession -ErrorAction Stop)) {
                $number = [uint64]0
                if ([uint64]::TryParse([string]$session.LogonId, [ref]$number)) { [void]$liveLogons.Add($number) }
            }
        } catch { }
    }

    # 監査ポリシー（管理者として、この PC を調べるときだけ分かる）
    $auditNotes = [System.Collections.Generic.List[string]]::new()
    if (-not $useFiles -and $isAdmin) {
        Write-Progress -Activity $ToolTitle -Status '監査ポリシーを確かめています'
        try {
            if (-not ('WbLogonHistoryAudit' -as [type])) { Add-Type -TypeDefinition $AuditSource -IgnoreWarnings -WarningAction SilentlyContinue -ErrorAction Stop }
            $logonAudit  = [WbLogonHistoryAudit]::Query($LogonAuditGuid)
            $logoffAudit = [WbLogonHistoryAudit]::Query($LogoffAuditGuid)
            if (($logonAudit -band 1) -eq 0) { $auditNotes.Add('※ ログオンの成功を記録しない設定です（監査ポリシー）。ログオンは表示されません') }
            if (($logonAudit -band 2) -eq 0) { $auditNotes.Add('※ ログオンの失敗を記録しない設定です（監査ポリシー）。失敗は表示されません') }
            if (($logoffAudit -band 1) -eq 0) { $auditNotes.Add('※ ログオフを記録しない設定です（監査ポリシー）。ログオフの時刻は分かりません') }
        } catch { }
    }

    #endregion

    #region ログオンを組にする

    Write-Progress -Activity $ToolTitle -Status 'ログオンを組にしています'
    # Sort-Object は同じキーのものの並びが 5.1 と 7 で違うので、記録の番号まで並べて、結果をそろえる
    $sortedLogons = @($logons.ToArray() | Sort-Object -Property Time, Computer, RecordId)

    # UAC の有効な環境で管理者がログオンすると、昇格したログオンとしないログオンが同時に記録され、お互いを指す。1 つにまとめる。
    # ログオン ID は OS を起動し直すと使い回されるので、同じ ID のログオンが何件あってもよいように、記録ごとの番号で「使った」を覚え、
    # 相方は同じ ID のうち 2 秒以内で、まだ使っていないものから選ぶ
    for ($i = 0; $i -lt $sortedLogons.Count; $i++) { $sortedLogons[$i].Index = $i }
    $byId = @{}
    foreach ($logon in $sortedLogons) {
        $key = $logon.Computer + '|' + $logon.Id
        if (-not $byId.ContainsKey($key)) { $byId[$key] = [System.Collections.Generic.List[object]]::new() }
        $byId[$key].Add($logon)
    }
    $used = New-Object 'System.Collections.Generic.HashSet[int]'
    $units = [System.Collections.Generic.List[object]]::new()
    foreach ($logon in $sortedLogons) {
        if (-not $used.Add($logon.Index)) { continue }
        $ids = [System.Collections.Generic.List[uint64]]::new()
        $ids.Add($logon.Id)
        $elevated = $logon.Elevated
        if ($logon.Linked -ne 0) {
            $linkedKey = $logon.Computer + '|' + $logon.Linked
            if ($byId.ContainsKey($linkedKey)) {
                foreach ($other in $byId[$linkedKey]) {
                    if ($used.Contains($other.Index) -or [math]::Abs(($other.Time - $logon.Time).TotalSeconds) -gt 2) { continue }
                    [void]$used.Add($other.Index)
                    $ids.Add($other.Id)
                    if ($other.Elevated -eq $true) { $elevated = $true } elseif ($null -eq $elevated) { $elevated = $other.Elevated }
                    break
                }
            }
        }
        $ids.Sort()
        $units.Add([pscustomobject]@{
            Computer = $logon.Computer; ComputerText = $logon.ComputerText; Time = $logon.Time; RecordId = $logon.RecordId; Sid = $logon.Sid; Name = $logon.Name; Type = $logon.Type
            Ids = $ids; Elevated = $elevated; Ip = $logon.Ip; Workstation = $logon.Workstation; Process = $logon.Process; Auth = $logon.Auth
            Logoff = $null; LogoffEventId = 0; Kind = 'Session'; Status = ''; Lsm = $null; Unlocks = 0; Disconnects = 0; Reconnects = 0; LastDisconnect = $null
        })
    }

    # ログオフ。ログオン ID は OS を起動し直すと使い回されることがあるので、次の起動より前の記録だけを見る
    function Get-NextBoot([string]$Computer, [datetime]$Time) {
        if (-not $boots.ContainsKey($Computer)) { return $null }
        foreach ($boot in $boots[$Computer]) { if ($boot -gt $Time) { return $boot } }
        return $null
    }
    # ユーザーがログオフした記録（4647）があればその時刻、なければ 4634 の早いほう。
    # UAC の 2 つのログオンは、片方の 4634 だけが早く記録されることがあるかもしれないので、4647 を優先する
    foreach ($unit in $units) {
        $limit = Get-NextBoot $unit.Computer $unit.Time
        $userLogoff = $null
        $sessionEnd = $null
        $sessionEndId = 0
        foreach ($id in $unit.Ids) {
            $key = $unit.Computer + '|' + $id
            if (-not $logoffs.ContainsKey($key)) { continue }
            foreach ($entry in $logoffs[$key]) {
                $t = $entry.Time
                if ($t -lt $unit.Time.AddSeconds(-1)) { continue }
                if ($null -ne $limit -and $t -gt $limit) { continue }
                $entry.Used = $true
                if ($entry.EventId -eq 4647) { if ($null -eq $userLogoff -or $t -lt $userLogoff) { $userLogoff = $t } }
                elseif ($null -eq $sessionEnd -or $t -lt $sessionEnd) { $sessionEnd = $t; $sessionEndId = $entry.EventId }
            }
        }
        if ($null -ne $userLogoff) { $unit.Logoff = $userLogoff; $unit.LogoffEventId = 4647 }
        elseif ($null -ne $sessionEnd) { $unit.Logoff = $sessionEnd; $unit.LogoffEventId = $sessionEndId }
    }

    # ログオンと組にできなかったログオフ。ログオンの記録が、もう Security ログに残っていない（期間より前から続いていた）ログオンのもの。
    # 同じログオン ID の 4647 と 4634 は 1 つにまとめ（ログオン ID は OS を起動し直すと使い回されるので、起動をまたぐものは別）、
    # ログオンと同じく、4647 があればその時刻にする
    $orphanGroups = @{}
    foreach ($list in $logoffs.Values) {
        foreach ($entry in $list) {
            if ($entry.Used -or $entry.System) { continue }
            $next = Get-NextBoot $entry.Computer $entry.Time
            $key = '{0}|{1}|{2}' -f $entry.Computer, $entry.Id, $(if ($null -ne $next) { $next.Ticks } else { 0 })
            if (-not $orphanGroups.ContainsKey($key)) { $orphanGroups[$key] = [System.Collections.Generic.List[object]]::new() }
            $orphanGroups[$key].Add($entry)
        }
    }
    $orphanItems = [System.Collections.Generic.List[object]]::new()
    foreach ($group in $orphanGroups.Values) {
        $chosen = $null
        $type = 0
        foreach ($entry in $group) {
            if ($type -eq 0) { $type = $entry.Type }
            if ($null -eq $chosen -or ($entry.EventId -eq 4647 -and $chosen.EventId -ne 4647) -or ($entry.EventId -eq $chosen.EventId -and $entry.Time -lt $chosen.Time)) { $chosen = $entry }
        }
        $orphanItems.Add([pscustomobject]@{ Computer = $chosen.Computer; ComputerText = $chosen.ComputerText; Time = $chosen.Time; EventId = $chosen.EventId; Sid = $chosen.Sid; Name = $chosen.Name; Id = $chosen.Id; Type = $type })
    }
    # UAC の 2 つのログオンは、ログオンの記録がないとお互いが分からないので、同じユーザーの近い（$SignInSeconds 秒以内の）ものを 1 つにまとめる
    $orphanLogoffs = [System.Collections.Generic.List[object]]::new()
    $orphan = $null
    foreach ($item in @($orphanItems.ToArray() | Sort-Object -Property Computer, Sid, Time, Id)) {
        if ($orphan -and $orphan.Computer -eq $item.Computer -and $orphan.Sid -eq $item.Sid -and ($item.Time - $orphan.First).TotalSeconds -le $SignInSeconds) {
            if ($item.EventId -eq 4647 -and $orphan.EventId -ne 4647) { $orphan.Time = $item.Time; $orphan.EventId = 4647; $orphan.Id = $item.Id; $orphan.Name = $item.Name }
            if ($orphan.Type -eq 0) { $orphan.Type = $item.Type }
            continue
        }
        $orphan = [pscustomobject]@{ Computer = $item.Computer; ComputerText = $item.ComputerText; First = $item.Time; Time = $item.Time; EventId = $item.EventId; Sid = $item.Sid; Name = $item.Name; Id = $item.Id; Type = $item.Type }
        $orphanLogoffs.Add($orphan)
    }

    # ロック解除。1 回のロック解除で、ロック解除のログオンが 2～4 つ同時に記録されるので、同じユーザーの近いものを 1 回にまとめる
    $unlocks = [System.Collections.Generic.List[object]]::new()
    $current = $null
    foreach ($unit in @($units | Where-Object { $UnlockTypes -contains $_.Type } | Sort-Object -Property Computer, Sid, Time, RecordId)) {
        if ($current -and $current.Computer -eq $unit.Computer -and $current.Sid -eq $unit.Sid -and ($unit.Time - $current.Time).TotalSeconds -le $UnlockGapSeconds) { continue }
        $current = [pscustomobject]@{
            Computer = $unit.Computer; ComputerText = $unit.ComputerText; Time = $unit.Time; Sid = $unit.Sid; Name = $unit.Name; Ip = $unit.Ip; Workstation = $unit.Workstation
            Id = $unit.Ids[0]; Real = $true; Session = $null
        }
        $unlocks.Add($current)
    }
    $unlocksByUser = @{}
    foreach ($unlock in $unlocks) {
        $key = $unlock.Computer + '|' + $unlock.Sid
        if (-not $unlocksByUser.ContainsKey($key)) { $unlocksByUser[$key] = [System.Collections.Generic.List[object]]::new() }
        $unlocksByUser[$key].Add($unlock)
    }

    # 数秒で終わった対話・リモートデスクトップのログオンは、ロック解除や再接続のときに資格情報を確かめるためのもの。ログオンには数えない
    foreach ($unit in $units) {
        if ($UnlockTypes -contains $unit.Type) { $unit.Kind = 'Unlock'; continue }
        if ($SessionTypes -notcontains $unit.Type -or $null -eq $unit.Logoff) { continue }
        if (($unit.Logoff - $unit.Time).TotalSeconds -gt $ShortSeconds) { continue }
        $key = $unit.Computer + '|' + $unit.Sid
        $near = $false
        if ($unlocksByUser.ContainsKey($key)) {
            foreach ($unlock in $unlocksByUser[$key]) { if ([math]::Abs(($unlock.Time - $unit.Time).TotalSeconds) -le $SignInSeconds) { $near = $true; break } }
        }
        if ($near) { $unit.Kind = 'UnlockPart' }
        elseif ($RemoteTypes -contains $unit.Type) { $unit.Kind = 'Reconnect' }
    }

    # 続くログオンの一覧（コンピューター|SID ごと、ログオンの古い順）
    $sessionsByUser = @{}
    foreach ($unit in $units) {
        if ($unit.Kind -ne 'Session' -or $SessionTypes -notcontains $unit.Type) { continue }
        $key = $unit.Computer + '|' + $unit.Sid
        if (-not $sessionsByUser.ContainsKey($key)) { $sessionsByUser[$key] = [System.Collections.Generic.List[object]]::new() }
        $sessionsByUser[$key].Add($unit)
    }
    # その時刻に続いているログオン（ログオンが後のものを優先）
    function Find-ActiveSession([string]$Computer, [string]$Sid, [datetime]$Time) {
        $key = $Computer + '|' + $Sid
        if (-not $sessionsByUser.ContainsKey($key)) { return $null }
        $found = $null
        foreach ($session in $sessionsByUser[$key]) {
            if ($session.Time -gt $Time) { break }
            if ($null -eq $session.Logoff -or $session.Logoff -ge $Time) { $found = $session }
        }
        return $found
    }

    # サインインのときにも、ロック解除のログオンが記録される。ログオンと同じ時刻のものは、ロック解除に数えない
    foreach ($unlock in $unlocks) {
        $key = $unlock.Computer + '|' + $unlock.Sid
        if ($sessionsByUser.ContainsKey($key)) {
            foreach ($session in $sessionsByUser[$key]) { if ([math]::Abs(($session.Time - $unlock.Time).TotalSeconds) -le $SignInSeconds) { $unlock.Real = $false; break } }
        }
        if (-not $unlock.Real) { continue }
        # ログオンの行に数えるのは、一覧にするとき（期間の中のものだけを数えるため）
        $unlock.Session = Find-ActiveSession $unlock.Computer $unlock.Sid $unlock.Time
    }

    #endregion

    #region 切断・再接続（セッションのログ）

    # セッションのログを、セッション ID ごとに、ログオン（21）からログオフ（23）までにまとめる。
    # ログオフの後にも切断（24）が記録されるので、それは数えない
    $lsmSessions = [System.Collections.Generic.List[object]]::new()
    $lsmOrphans = [System.Collections.Generic.List[object]]::new()   # ログオン（21）が、古い記録から消えて残っていないセッション
    $lsmRows = [System.Collections.Generic.List[object]]::new()
    $lsmState = @{}
    foreach ($evt in @($lsmEvents.ToArray() | Sort-Object -Property Time, Computer, RecordId)) {
        $key = $evt.Computer + '|' + $evt.SessionId
        if (-not $lsmState.ContainsKey($key)) { $lsmState[$key] = @{ Current = $null; Closed = $false } }
        $state = $lsmState[$key]
        switch ($evt.Id) {
            21 {
                $session = [pscustomobject]@{ Computer = $evt.Computer; SessionId = $evt.SessionId; User = $evt.User; Ip = $evt.Ip; Start = $evt.Time; End = $null; Events = ([System.Collections.Generic.List[object]]::new()); Match = $null }
                $lsmSessions.Add($session)
                $state.Current = $session
                $state.Closed = $false
            }
            23 {
                if ($state.Current) { $state.Current.End = $evt.Time }
                $state.Current = $null
                $state.Closed = $true
            }
            default {
                # switch の中の continue は foreach ではなく switch に効くので、if で分ける
                if (-not $state.Closed) {
                    # ログオン（21）より前に切断・再接続があるのは、21 が古い記録から消えたセッション。後で、続いているログオンと組にする
                    if ($null -eq $state.Current) {
                        $state.Current = [pscustomobject]@{ Computer = $evt.Computer; SessionId = $evt.SessionId; User = $evt.User; Ip = $evt.Ip; Start = $null; End = $null; Events = ([System.Collections.Generic.List[object]]::new()); Match = $null }
                        $lsmOrphans.Add($state.Current)
                    }
                    $lsmRows.Add([pscustomobject]@{ Event = $evt; Session = $state.Current })
                    $state.Current.Events.Add($evt)
                }
            }
        }
    }

    # セッションのログのログオンと、Security のログオンを組にする（ユーザー・時刻・接続元）。
    # 件数が多い（リモートデスクトップのサーバーなど）ときに遅くならないよう、コンピューターごとに時刻の順に並べ、
    # 時刻の範囲に入る候補だけを見る（ログオンもセッションのログも古い順なので、範囲の始まりは前に進むだけ）
    $candidatesByComputer = @{}
    foreach ($unit in $units) {
        if ($unit.Kind -ne 'Session' -or $SessionTypes -notcontains $unit.Type) { continue }
        if (-not $candidatesByComputer.ContainsKey($unit.Computer)) { $candidatesByComputer[$unit.Computer] = [System.Collections.Generic.List[object]]::new() }
        $candidatesByComputer[$unit.Computer].Add($unit)
    }
    $firstIndex = @{}
    foreach ($ls in $lsmSessions) {
        if (-not $candidatesByComputer.ContainsKey($ls.Computer)) { continue }
        $list = $candidatesByComputer[$ls.Computer]
        $from = $ls.Start.AddSeconds(-$LsmAfterSeconds)
        $to = $ls.Start.AddSeconds($LsmBeforeSeconds)
        $i = if ($firstIndex.ContainsKey($ls.Computer)) { $firstIndex[$ls.Computer] } else { 0 }
        while ($i -lt $list.Count -and $list[$i].Time -lt $from) { $i++ }
        $firstIndex[$ls.Computer] = $i
        $best = $null
        $bestScore = [double]::MaxValue
        for ($j = $i; $j -lt $list.Count -and $list[$j].Time -le $to; $j++) {
            $unit = $list[$j]
            if ($unit.Lsm) { continue }
            $diff = ($ls.Start - $unit.Time).TotalSeconds
            # 接続元：リモートデスクトップなら IP アドレスが同じ、コンソールならどちらも IP アドレスがない（Hyper-V の拡張セッションは 0.0.0.0）
            if ($ls.Ip) { if ($RemoteTypes -notcontains $unit.Type -or $unit.Ip -ne $ls.Ip) { continue } }
            elseif ($unit.Ip -and $unit.Ip -ne '0.0.0.0') { continue }
            $names = $sidNames[$unit.Computer + '|' + $unit.Sid]
            $nameMatch = $names -and $names.Contains($ls.User)
            $score = [math]::Abs($diff) + $(if ($nameMatch) { 0 } else { 100000 })
            if ($score -lt $bestScore) { $best = $unit; $bestScore = $score }
        }
        if ($best) {
            # 切断・再接続の回数は、一覧にするときに数える（期間の中のものだけを数えるため）
            $best.Lsm = $ls
            $ls.Match = $best
            Add-SidName $best.Computer $best.Sid $ls.User
        }
    }

    # ログオン（21）が残っていないセッションは、最初の切断・再接続の時刻に続いていた、同じユーザーのログオンと組にする。
    # 再接続はほかの端末からのこともあるので、接続元は同じものを優先するだけにする（リモートデスクトップとコンソールの区別は見る）
    foreach ($ls in $lsmOrphans) {
        if (-not $candidatesByComputer.ContainsKey($ls.Computer)) { continue }
        $at = $ls.Events[0].Time
        $best = $null
        $bestIp = $false
        foreach ($unit in $candidatesByComputer[$ls.Computer]) {
            if ($unit.Time -gt $at) { break }
            if ($unit.Lsm -or ($null -ne $unit.Logoff -and $unit.Logoff -lt $at)) { continue }
            $next = Get-NextBoot $unit.Computer $unit.Time
            if ($null -ne $next -and $next -le $at) { continue }
            if ($ls.Ip) { if ($RemoteTypes -notcontains $unit.Type) { continue } }
            elseif ($unit.Ip -and $unit.Ip -ne '0.0.0.0') { continue }
            $names = $sidNames[$unit.Computer + '|' + $unit.Sid]
            if (-not ($names -and $names.Contains($ls.User)) -and -not [string]::Equals($unit.Name, $ls.User, [StringComparison]::OrdinalIgnoreCase)) { continue }
            # 候補は古い順なので、後のもので上書きする（接続元が同じものがあれば、それを優先する）
            $sameIp = [bool]$ls.Ip -and $unit.Ip -eq $ls.Ip
            if ($sameIp -or -not $bestIp) { $best = $unit; $bestIp = $sameIp }
        }
        if ($best) {
            $best.Lsm = $ls
            $ls.Match = $best
        }
    }

    # Security の再接続（数秒で終わったリモートデスクトップのログオン）。セッションのログに同じ再接続があれば、そちらを使う。
    # セッションのログの再接続の時刻は、コンピューターごとに並べて、二分探索で近いものを探す
    $reconnectTimes = @{}
    foreach ($evt in $lsmEvents) {
        if ($evt.Id -ne 25) { continue }
        if (-not $reconnectTimes.ContainsKey($evt.Computer)) { $reconnectTimes[$evt.Computer] = [System.Collections.Generic.List[datetime]]::new() }
        $reconnectTimes[$evt.Computer].Add($evt.Time)
    }
    foreach ($computer in @($reconnectTimes.Keys)) { $reconnectTimes[$computer].Sort() }
    function Test-NearReconnect([string]$Computer, [datetime]$Time) {
        if (-not $reconnectTimes.ContainsKey($Computer)) { return $false }
        $list = $reconnectTimes[$Computer]
        $index = $list.BinarySearch($Time)
        if ($index -lt 0) { $index = -bnot $index }
        foreach ($k in @(($index - 1), $index)) {
            if ($k -ge 0 -and $k -lt $list.Count -and [math]::Abs(($list[$k] - $Time).TotalSeconds) -le $ReconnectSeconds) { return $true }
        }
        return $false
    }
    $secReconnects = [System.Collections.Generic.List[object]]::new()
    foreach ($unit in @($units | Where-Object { $_.Kind -eq 'Reconnect' })) {
        if (Test-NearReconnect $unit.Computer $unit.Time) { continue }
        $unit | Add-Member -NotePropertyName ActiveSession -NotePropertyValue (Find-ActiveSession $unit.Computer $unit.Sid $unit.Time)
        $secReconnects.Add($unit)
    }

    # ログオフの記録がないログオンの状態。OS を起動し直すとログオンは終わるので、再起動を先に見る
    # （切断したまま夜間に再起動した RDP を「切断中」にしないため。ログオン ID も再起動の後に使い回されることがある）
    foreach ($unit in $units) {
        if ($unit.Kind -ne 'Session') { continue }
        if ($null -ne $unit.Logoff) { $unit.Status = 'ログオフ'; continue }
        if ($null -ne (Get-NextBoot $unit.Computer $unit.Time)) { $unit.Status = '再起動'; continue }
        if ($unit.Lsm -and $null -eq $unit.Lsm.End -and $unit.Lsm.Events.Count -gt 0 -and $unit.Lsm.Events[$unit.Lsm.Events.Count - 1].Id -eq 24) { $unit.Status = '切断中'; continue }
        $live = $false
        foreach ($id in $unit.Ids) { if ($liveLogons.Contains($id)) { $live = $true; break } }
        if ($live) { $unit.Status = 'ログオン中'; continue }
        $unit.Status = '記録なし'
    }

    #endregion

    #region 一覧にする

    function Get-UnitNames($Unit) {
        $names = [System.Collections.Generic.List[string]]::new()
        $names.Add($Unit.Name)
        $key = $Unit.Computer + '|' + $Unit.Sid
        if ($sidNames.ContainsKey($key)) { foreach ($name in $sidNames[$key]) { $names.Add($name) } }
        return , $names
    }

    # 期間に重なるログオンか（期間より前にログオンして、期間の始めにまだ続いていたものも含める）
    function Test-InPeriod($Unit) {
        if ($Unit.Time -gt $endTime) { return $false }
        if ($Unit.Time -ge $sinceInfo.Time) { return $true }
        switch ($Unit.Status) {
            'ログオフ' { return $Unit.Logoff -ge $sinceInfo.Time }
            '再起動' { return (Get-NextBoot $Unit.Computer $Unit.Time) -ge $sinceInfo.Time }
            '切断中' { return $true }
            'ログオン中' { return $true }
        }
        # 記録なし：いつ終わったか分からないので、期間の中にログオンしたものだけを出す
        return $false
    }
    function Test-InWindow([datetime]$Time) {
        return $Time -ge $sinceInfo.Time -and $Time -le $endTime
    }

    # -User があるか（名前の一覧を作るのは、絞り込むときだけにする。件数が多いときに遅くならないように）
    $userFilter = $userPatterns.Count -gt 0
    $periodUnits = [System.Collections.Generic.List[object]]::new()
    foreach ($unit in $units) {
        if ($unit.Kind -ne 'Session' -or -not $selectedTypes.Contains($unit.Type) -or -not (Test-InPeriod $unit)) { continue }
        if ($userFilter -and -not (Test-UserMatch (Get-UnitNames $unit))) { continue }
        $periodUnits.Add($unit)
    }
    $sessions = @($periodUnits.ToArray() | Sort-Object -Property @{ Expression = 'Time'; Descending = $true }, @{ Expression = 'Computer'; Descending = $false }, @{ Expression = 'RecordId'; Descending = $false })
    # events.csv には、期間の中の記録だけを書く。並びを 5.1 と 7 でそろえるため、足した順の番号も持つ
    $events = [System.Collections.Generic.List[object]]::new()
    function Add-EventRow($Time, [string]$Operation, [string]$Name, [string]$Type, [string]$Ip, [string]$Workstation, [string]$Detail, [string]$LogonId, [string]$SessionId, [int]$EventId, [string]$Computer) {
        $events.Add([pscustomobject]@{ Time = $Time; Seq = $events.Count; Operation = $Operation; Name = $Name; Type = $Type; Ip = $Ip; Workstation = $Workstation; Detail = $Detail; LogonId = $LogonId; SessionId = $SessionId; EventId = $EventId; Computer = $Computer })
    }
    function Get-AdminText($Elevated) {
        if ($Elevated -eq $true) { return 'はい' }
        if ($Elevated -eq $false) { return 'いいえ' }
        return ''
    }
    foreach ($session in $sessions) {
        $label = Get-TypeLabel $session.Type $session.Process
        $sessionId = if ($session.Lsm) { $session.Lsm.SessionId } else { '' }
        $admin = Get-AdminText $session.Elevated
        $detail = if ($admin) { '管理者権限: ' + $admin } else { '' }
        if (Test-InWindow $session.Time) {
            Add-EventRow $session.Time 'ログオン' $session.Name $label $session.Ip $session.Workstation $detail (Format-LogonId $session.Ids[0]) $sessionId 4624 $session.ComputerText
        }
        if ($null -ne $session.Logoff -and (Test-InWindow $session.Logoff)) {
            Add-EventRow $session.Logoff 'ログオフ' $session.Name $label $session.Ip $session.Workstation '' (Format-LogonId $session.Ids[0]) $sessionId $session.LogoffEventId $session.ComputerText
        }
    }
    # ログオンの記録が残っていないログオフ（期間より前から続いていたログオン）。ログオンの行は作れないので、events.csv にだけ書く。
    # ロック解除のログオンのログオフは、組にできたときと同じく書かない
    $orphanCounts = @{}   # コンピューター → 件数（注意の行に書く）
    $orphanTotal = 0      # 該当する記録に数える（ログオンの表には出ないが、期間の中に誰かがログオンしていた記録なので）
    foreach ($orphan in $orphanLogoffs) {
        if (-not (Test-InWindow $orphan.Time)) { continue }
        if ($orphan.Type -ne 0) {
            if ($UnlockTypes -contains $orphan.Type -or -not $selectedTypes.Contains($orphan.Type)) { continue }
            $label = Get-TypeLabel $orphan.Type ''
        } else {
            # 4647（ユーザーがログオフした記録）だけのときは、種類が分からない。対話かリモートデスクトップのどちらか
            if (-not $sessionSelected) { continue }
            $label = ''
        }
        if ($userFilter) {
            $names = [System.Collections.Generic.List[string]]::new()
            $names.Add($orphan.Name)
            if ($sidNames.ContainsKey($orphan.Computer + '|' + $orphan.Sid)) { foreach ($name in $sidNames[$orphan.Computer + '|' + $orphan.Sid]) { $names.Add($name) } }
            if (-not (Test-UserMatch $names)) { continue }
        }
        Add-EventRow $orphan.Time 'ログオフ' $orphan.Name $label '' '' 'ログオンの記録が残っていません' (Format-LogonId $orphan.Id) '' $orphan.EventId $orphan.ComputerText
        $orphanCounts[$orphan.Computer] = 1 + $(if ($orphanCounts.ContainsKey($orphan.Computer)) { $orphanCounts[$orphan.Computer] } else { 0 })
        $orphanTotal++
    }
    $unlockCount = 0
    if ($unlockSelected) {
        foreach ($unlock in $unlocks) {
            if (-not $unlock.Real -or -not (Test-InWindow $unlock.Time)) { continue }
            $names = [System.Collections.Generic.List[string]]::new()
            $names.Add($unlock.Name)
            if ($sidNames.ContainsKey($unlock.Computer + '|' + $unlock.Sid)) { foreach ($name in $sidNames[$unlock.Computer + '|' + $unlock.Sid]) { $names.Add($name) } }
            if (-not (Test-UserMatch $names)) { continue }
            $sessionId = if ($unlock.Session -and $unlock.Session.Lsm) { $unlock.Session.Lsm.SessionId } else { '' }
            Add-EventRow $unlock.Time 'ロック解除' $unlock.Name 'ロック解除' $unlock.Ip $unlock.Workstation '' (Format-LogonId $unlock.Id) $sessionId 4624 $unlock.ComputerText
            $unlockCount++
            # ログオンの行の回数も、合計の行・events.csv と同じく、期間の中のものだけを数える
            if ($unlock.Session) { $unlock.Session.Unlocks++ }
        }
    }
    $disconnectCount = 0
    $reconnectCount = 0
    if ($sessionSelected) {
        foreach ($row in $lsmRows) {
            $evt = $row.Event
            if (-not (Test-InWindow $evt.Time)) { continue }
            $match = if ($row.Session) { $row.Session.Match } else { $null }
            if ($match) {
                if (-not $selectedTypes.Contains($match.Type) -or ($userFilter -and -not (Test-UserMatch (Get-UnitNames $match)))) { continue }
                $name = $match.Name
                $label = Get-TypeLabel $match.Type $match.Process
            } else {
                # 組にできなかったものは、接続元の IP アドレスの有無で種類を決め、その種類を選んでいるときだけ出す
                $group = if ($evt.Ip) { 'RemoteDesktop' } else { 'Interactive' }
                if (-not $selectedGroups.Contains($group)) { continue }
                $names = [System.Collections.Generic.List[string]]::new()
                $names.Add($evt.User)
                if (-not (Test-UserMatch $names)) { continue }
                $name = $evt.User
                $label = $TypeGroups[$group].Label
            }
            $operation = if ($evt.Id -eq 24) { '切断' } else { '再接続' }
            $logonId = if ($match) { Format-LogonId $match.Ids[0] } else { '' }
            Add-EventRow $evt.Time $operation $name $label $evt.Ip '' '' $logonId $evt.SessionId $evt.Id $evt.ComputerText
            if ($evt.Id -eq 24) { $disconnectCount++ } else { $reconnectCount++ }
            if ($match) {
                if ($evt.Id -eq 24) { $match.Disconnects++; $match.LastDisconnect = $evt.Time } else { $match.Reconnects++ }
            }
        }
        foreach ($unit in $secReconnects) {
            if (-not (Test-InWindow $unit.Time) -or -not $selectedTypes.Contains($unit.Type) -or ($userFilter -and -not (Test-UserMatch (Get-UnitNames $unit)))) { continue }
            Add-EventRow $unit.Time '再接続' $unit.Name (Get-TypeLabel $unit.Type $unit.Process) $unit.Ip $unit.Workstation '' (Format-LogonId $unit.Ids[0]) '' 4624 $unit.ComputerText
            $reconnectCount++
            if ($unit.ActiveSession) { $unit.ActiveSession.Reconnects++ }
        }
    }
    $shownFailures = [System.Collections.Generic.List[object]]::new()
    $sortedFailures = @($failures.ToArray() | Sort-Object -Property @{ Expression = 'Time'; Descending = $true }, @{ Expression = 'Computer'; Descending = $false }, @{ Expression = 'RecordId'; Descending = $false })
    foreach ($failure in $sortedFailures) {
        $names = [System.Collections.Generic.List[string]]::new()
        $names.Add($failure.Name)
        if (-not (Test-UserMatch $names)) { continue }
        $reason = Get-FailureReason $failure.Status $failure.SubStatus
        $code = Format-Invariant '0x{0:X8}' $failure.Status
        if ($failure.SubStatus -ne 0 -and $failure.SubStatus -ne $failure.Status) { $code += Format-Invariant ' / 0x{0:X8}' $failure.SubStatus }
        $failure.Reason = $reason
        $shownFailures.Add($failure)
        # 理由を引けない状態コードは、理由の側にコードが入っているので、重ねて付けない
        $detail = if ($reason.StartsWith('その他', [StringComparison]::Ordinal)) { 'その他（{0}）' -f $code } else { '{0}（{1}）' -f $reason, $code }
        Add-EventRow $failure.Time 'ログオンの失敗' $failure.Name (Get-TypeLabel $failure.Type '') $failure.Ip $failure.Workstation $detail '' '' 4625 $failure.ComputerText
    }
    foreach ($computer in @($boots.Keys | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } })) {
        foreach ($boot in $boots[$computer]) {
            if (-not (Test-InWindow $boot)) { continue }
            $nameKey = $computer + '|' + $boot.Ticks
            $text = if ($bootNames.ContainsKey($nameKey)) { $bootNames[$nameKey] } else { $computer }
            Add-EventRow $boot 'OS の起動' '' '' '' '' '' '' '' 4608 $text
        }
    }
    foreach ($clear in $clears) {
        if (-not (Test-InWindow $clear.Time)) { continue }
        Add-EventRow $clear.Time 'ログの消去' $clear.By '' '' '' ('消した人: {0}' -f $clear.By) '' '' 1102 $clear.ComputerText
    }
    $sortedEvents = @($events.ToArray() | Sort-Object -Property @{ Expression = 'Time'; Descending = $true }, @{ Expression = 'Seq'; Descending = $false })

    # ユーザーごと（ログオンの記録の名前でまとめる）
    $userTable = @{}
    function Get-UserEntry([string]$Name) {
        $key = $Name.ToLowerInvariant()
        if (-not $userTable.ContainsKey($key)) {
            $userTable[$key] = [pscustomobject]@{ Name = $Name; Logons = 0; First = $null; Last = $null; Total = [TimeSpan]::Zero; Completed = 0; Failures = 0; LastFailure = $null }
        }
        return $userTable[$key]
    }
    foreach ($session in $sessions) {
        $entry = Get-UserEntry $session.Name
        $entry.Logons++
        if ($null -eq $entry.First -or $session.Time -lt $entry.First) { $entry.First = $session.Time }
        if ($null -eq $entry.Last -or $session.Time -gt $entry.Last) { $entry.Last = $session.Time }
        if ($null -ne $session.Logoff) { $entry.Total += ($session.Logoff - $session.Time); $entry.Completed++ }
    }
    foreach ($failure in $shownFailures) {
        $entry = Get-UserEntry $failure.Name
        $entry.Failures++
        if ($null -eq $entry.LastFailure -or $failure.Time -gt $entry.LastFailure) { $entry.LastFailure = $failure.Time }
    }
    $userList = @($userTable.Values | Sort-Object -Property @{ Expression = 'Logons'; Descending = $true }, @{ Expression = 'Failures'; Descending = $true }, @{ Expression = { ConvertTo-SortKey $_.Name }; Descending = $false })

    # ログオンの失敗（ユーザーと接続元ごと）
    $failureGroups = @{}
    foreach ($failure in $shownFailures) {
        $sourceText = if ($failure.Ip) { $failure.Ip } else { $failure.Workstation }
        $key = $failure.Computer + '|' + $failure.Name.ToLowerInvariant() + '|' + $sourceText
        if (-not $failureGroups.ContainsKey($key)) {
            # 新しい順に見ているので、最初に見たものが最後の失敗
            $failureGroups[$key] = [pscustomobject]@{ Count = 0; Name = $failure.Name; Source = $sourceText; Last = $failure.Time; Reason = $failure.Reason; Computer = $failure.Computer }
        }
        $failureGroups[$key].Count++
    }
    $failureList = @($failureGroups.Values | Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Last'; Descending = $true })

    #endregion

    #region 画面と summary.txt の内容

    $head = [System.Collections.Generic.List[string]]::new()
    $head.Add((Format-Invariant '{0} ({1} {2})' $ToolTitle $ToolName $ScriptVersion))
    $securitySources = @($sources | Where-Object { $_.Kind -eq 'Security' })
    $oldestByComputer = @{}
    foreach ($source in $securitySources) {
        if ($null -eq $source.Oldest) { continue }
        if (-not $oldestByComputer.ContainsKey($source.Computer) -or $source.Oldest -lt $oldestByComputer[$source.Computer]) { $oldestByComputer[$source.Computer] = $source.Oldest }
    }
    $oldestText = ''
    if ($oldestByComputer.Count -eq 1) { $oldestText = Format-Invariant '   Security ログ: {0:yyyy-MM-dd HH:mm} から残っています' (@($oldestByComputer.Values)[0]) }
    if ($useFiles) {
        $more = if ($sources.Count -gt 1) { Format-Invariant ' ほか {0:N0} ファイル' ($sources.Count - 1) } else { '' }
        # Security のファイルを先に出す
        $firstLabel = if ($securitySources.Count -gt 0) { $securitySources[0].Label } else { $sources[0].Label }
        $head.Add(('対象: {0}{1}{2}' -f $firstLabel, $more, $oldestText))
    } else {
        $head.Add(('対象: {0}（管理者: {1}）{2}' -f $machineName, $(if ($isAdmin) { 'はい' } else { 'いいえ' }), $oldestText))
    }
    $span = if ($sinceInfo.Relative -and -not $untilInfo) { '（{0}）' -f $sinceInfo.Relative } else { '' }
    $head.Add((Format-Invariant '期間: {0:yyyy-MM-dd HH:mm} ～ {1:yyyy-MM-dd HH:mm}{2}   種類: {3}' $sinceInfo.Time $endTime $span $typeLabelText))
    if ($userItems.Count -gt 0) { $head.Add(('条件: ユーザー {0}' -f ($userItems -join ', '))) }
    $head.Add('')

    # 表の行
    $userHead = [string[]]@('ユーザー', 'ログオン', '最後のログオン', '失敗', '最後の失敗')
    $userRight = [bool[]]@($false, $true, $false, $true, $false)
    $userRows = @(foreach ($entry in $userList) {
        , [string[]]@((Limit-Width $entry.Name $NameWidth), $entry.Logons.ToString('N0', $Invariant), (Format-ScreenTime $entry.Last), $entry.Failures.ToString('N0', $Invariant), (Format-ScreenTime $entry.LastFailure))
    })
    # 複数のコンピューターの .evtx を読んだときは、どのサーバーの行かが分かるように、表の最後にコンピューターの列を足す
    $readComputers = @($sources | ForEach-Object { $_.Computer } | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } } -Unique)
    $multiComputer = $readComputers.Count -gt 1
    $sessionHead = [string[]]@('ログオン', 'ログオフ', '時間', 'ユーザー', '種類', '接続元', '管理者')
    $sessionRight = [bool[]]@($false, $false, $true, $false, $false, $false, $false)
    if ($multiComputer) { $sessionHead += 'コンピューター'; $sessionRight += $false }
    $sessionRows = @(foreach ($session in $sessions) {
        $logoffText = if ($null -ne $session.Logoff) { Format-LogoffTime $session.Time $session.Logoff } else { '（{0}）' -f $session.Status }
        $spanText = if ($null -ne $session.Logoff) { Format-Span ($session.Logoff - $session.Time) } else { '' }
        $sourceText = if ($session.Ip) { $session.Ip } else { '' }
        $cells = [string[]]@((Format-ScreenTime $session.Time), $logoffText, $spanText, (Limit-Width $session.Name $NameWidth), (Get-TypeLabel $session.Type $session.Process), (Limit-Width $sourceText $SourceWidth), (Get-AdminText $session.Elevated))
        if ($multiComputer) { $cells += $session.Computer }
        , $cells
    })
    $failureHead = [string[]]@('件数', 'ユーザー', '接続元', '最後', '理由')
    $failureRight = [bool[]]@($true, $false, $false, $false, $false)
    if ($multiComputer) { $failureHead += 'コンピューター'; $failureRight += $false }
    $failureRows = @(foreach ($group in $failureList) {
        $cells = [string[]]@($group.Count.ToString('N0', $Invariant), (Limit-Width $group.Name $NameWidth), (Limit-Width $group.Source $SourceWidth), (Format-ScreenTime $group.Last), $group.Reason)
        if ($multiComputer) { $cells += $group.Computer }
        , $cells
    })

    $tail = [System.Collections.Generic.List[string]]::new()
    $total = $sessions.Count + $shownFailures.Count + $unlockCount + $orphanTotal
    # 見出しの最後が空行なので、表があるときだけ、表の後に空行を入れる
    if ($total -gt 0) { $tail.Add('') }
    $timeText = '所要 {0}' -f (Format-Duration $sw.Elapsed)
    if ($total -eq 0) {
        $tail.Add(('該当する記録はありません  {0}' -f $timeText))
    } else {
        $parts = @()
        if ($unlockSelected) { $parts += Format-Invariant 'ロック解除 {0:N0} 回' $unlockCount }
        if ($sessionSelected) { $parts += Format-Invariant '切断 {0:N0} 回' $disconnectCount; $parts += Format-Invariant '再接続 {0:N0} 回' $reconnectCount }
        $partText = if ($parts.Count -gt 0) { '（{0}）' -f ($parts -join '・') } else { '' }
        $tail.Add((Format-Invariant 'ログオン {0:N0} 件{1}  失敗 {2:N0} 件  {3}' $sessions.Count $partText $shownFailures.Count $timeText))
    }
    # 注意
    # 読んだコンピューターが複数あるときは、注意の先頭にコンピューターの名前を書く
    $allComputers = @($sources | ForEach-Object { $_.Computer } | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } } -Unique)
    function Get-ComputerPrefix([string]$Computer) { if ($allComputers.Count -gt 1) { return '{0}: ' -f $Computer }; return '' }
    foreach ($computer in @($oldestByComputer.Keys | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } })) {
        $oldest = $oldestByComputer[$computer]
        if ($oldest -gt $sinceInfo.Time) {
            # 複数のコンピューターの .evtx を読んだときは、どのコンピューターのことかを先頭に書く
            $tail.Add((Format-Invariant '※ {1}{0:yyyy-MM-dd HH:mm} より前の記録は、Security ログに残っていません（上書きされています）' $oldest (Get-ComputerPrefix $computer)))
            continue
        }
        # 期間の始めより前の OS の起動の記録がないと、最も古い記録より前にログオンして、期間まで続いていたものは分からない
        # （起動の記録があれば、期間の始めに続いていたログオンは、その起動より後のもので、すべて読んでいる）
        $bootBefore = $false
        if ($boots.ContainsKey($computer)) { foreach ($boot in $boots[$computer]) { if ($boot -lt $sinceInfo.Time) { $bootBefore = $true; break } } }
        if ($bootBefore) { continue }
        $count = if ($orphanCounts.ContainsKey($computer)) { $orphanCounts[$computer] } else { 0 }
        $reason = if ($count -gt 0) { Format-Invariant 'ログオフだけ {0:N0} 件を events.csv に書いています' $count } else { 'Security ログに記録が残っていません' }
        $tail.Add((Format-Invariant '※ {1}{0:yyyy-MM-dd HH:mm} より前から続いていたログオンは分かりません（{2}）' $oldest (Get-ComputerPrefix $computer) $reason))
    }
    if ($sessionSelected) {
        $lsmSources = @($sources | Where-Object { $_.Kind -eq 'Lsm' })
        # Security を読んだのに、セッションのログを読んでいないコンピューター
        foreach ($computer in @($sources | Where-Object { $_.Kind -eq 'Security' } | ForEach-Object { $_.Computer } | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } } -Unique)) {
            if (@($lsmSources | Where-Object { $_.Computer -eq $computer }).Count -eq 0) {
                $tail.Add(('※ {0}リモートデスクトップのセッションのログを読んでいないため、切断・再接続は分かりません' -f (Get-ComputerPrefix $computer)))
            }
        }
        # 同じコンピューターのアーカイブと今のログを一緒に読んだときは、いちばん古い記録で判断する
        $lsmOldest = @{}
        foreach ($source in $lsmSources) {
            if ($null -eq $source.Oldest) { continue }
            if (-not $lsmOldest.ContainsKey($source.Computer) -or $source.Oldest -lt $lsmOldest[$source.Computer]) { $lsmOldest[$source.Computer] = $source.Oldest }
        }
        foreach ($computer in @($lsmOldest.Keys | Sort-Object -Property @{ Expression = { ConvertTo-SortKey $_ } })) {
            if ($lsmOldest[$computer] -gt $sinceInfo.Time) {
                $tail.Add((Format-Invariant '※ {1}{0:yyyy-MM-dd HH:mm} より前の切断・再接続は分かりません（セッションのログに残っていません）' $lsmOldest[$computer] (Get-ComputerPrefix $computer)))
            }
        }
    }
    foreach ($clear in @($clears | Where-Object { Test-InWindow $_.Time } | Sort-Object -Property Time, Computer)) {
        $tail.Add((Format-Invariant '※ {2}{0:yyyy-MM-dd HH:mm} に Security ログが消去されています（{1}）' $clear.Time $clear.By (Get-ComputerPrefix $clear.Computer)))
    }
    foreach ($note in $auditNotes) { $tail.Add($note) }
    $kinds = @(foreach ($name in $skipKinds.Keys) { if ($skipKinds[$name] -gt 0) { , @($name, $skipKinds[$name]) } })
    $skipTotal = 0
    foreach ($kind in $kinds) { $skipTotal += $kind[1] }
    if ($kinds.Count -eq 1) { $tail.Add((Format-Invariant 'スキップ {0:N0} 件（{1}）' $skipTotal $kinds[0][0])) }
    elseif ($kinds.Count -gt 1) { $tail.Add((Format-Invariant 'スキップ {0:N0} 件（{1}）' $skipTotal (($kinds | ForEach-Object { Format-Invariant '{0} {1:N0}' $_[0] $_[1] }) -join '・'))) }

    # 表（Limit が 0 なら、すべての行）
    function Get-Section([string]$Title, [string[]]$Head, [bool[]]$Right, $Rows, [int]$Limit, [string]$MoreFile) {
        $lines = [System.Collections.Generic.List[string]]::new()
        $shown = @($Rows)
        $titleText = $Title
        if ($Limit -gt 0 -and $shown.Count -gt $Limit) { $shown = @($shown[0..($Limit - 1)]) }
        if ($Limit -gt 0) { $titleText = $Title -f (Format-Invariant '・{0} 件まで' $Limit) } else { $titleText = $Title -f '' }
        $lines.Add($titleText)
        if ($shown.Count -eq 0) { $lines.Add('    （該当なし）'); return , $lines }
        $lines.AddRange([string[]](Format-TextTable $Head $Right $shown))
        $rest = @($Rows).Count - $shown.Count
        if ($rest -gt 0) { $lines.Add((Format-Invariant '（ほか {0:N0} 件は {1} を参照してください）' $rest $MoreFile)) }
        return , $lines
    }
    function Get-Body([int]$Limit) {
        $body = [System.Collections.Generic.List[string]]::new()
        $body.AddRange((Get-Section '■ ユーザーごと（ログオンの多い順{0}）' $userHead $userRight $userRows $Limit 'users.csv'))
        $body.Add('')
        $body.AddRange((Get-Section '■ ログオン（新しい順{0}）' $sessionHead $sessionRight $sessionRows $Limit 'sessions.csv'))
        $body.Add('')
        $body.AddRange((Get-Section '■ ログオンの失敗（ユーザーと接続元ごと{0}）' $failureHead $failureRight $failureRows $Limit 'events.csv'))
        return , $body
    }

    #endregion

    #region 保存

    if ($runDir) {
        Write-Progress -Activity $ToolTitle -Status '結果を保存しています'
        try {
            $summary = [System.Collections.Generic.List[string]]::new()
            $summary.AddRange([string[]]$head.ToArray())
            if ($total -gt 0) { $summary.AddRange([string[]](Get-Body 0).ToArray()) }
            $summary.AddRange([string[]]$tail.ToArray())
            $summary.Add(('保存先: {0}' -f $runDir))
            [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$summary, $Utf8Bom)

            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($session in $sessions) {
                $spanText = if ($null -ne $session.Logoff) { Format-Span ($session.Logoff - $session.Time) } else { '' }
                $sessionId = if ($session.Lsm) { $session.Lsm.SessionId } else { '' }
                # ログオン ID（1 つか 2 つ）。行ごとにパイプラインを使うと、件数が多いときに遅い
                $idTexts = [string[]]::new($session.Ids.Count)
                for ($k = 0; $k -lt $idTexts.Count; $k++) { $idTexts[$k] = Format-LogonId $session.Ids[$k] }
                $rows.Add([object[]]@((Format-Time $session.Time), (Format-Time $session.Logoff), (Protect-ExcelText $spanText), (Protect-ExcelText $session.Name), (Get-TypeLabel $session.Type $session.Process),
                    (Protect-ExcelText $session.Ip), (Protect-ExcelText $session.Workstation), (Get-AdminText $session.Elevated), $session.Status,
                    $session.Unlocks.ToString('N0', $Invariant), $session.Disconnects.ToString('N0', $Invariant), $session.Reconnects.ToString('N0', $Invariant), (Format-Time $session.LastDisconnect),
                    (Protect-ExcelText $session.Auth), ($idTexts -join ' / '), $sessionId, (Protect-ExcelText $session.ComputerText)))
            }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'sessions.csv')) $SessionColumns $rows

            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($evt in $sortedEvents) {
                $rows.Add([object[]]@((Format-Time $evt.Time), $evt.Operation, (Protect-ExcelText $evt.Name), $evt.Type, (Protect-ExcelText $evt.Ip), (Protect-ExcelText $evt.Workstation),
                    (Protect-ExcelText $evt.Detail), $evt.LogonId, $evt.SessionId, $evt.EventId, (Protect-ExcelText $evt.Computer)))
            }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'events.csv')) $EventColumns $rows

            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in $userList) {
                # ログオフの記録があるログオンだけを足す。1 つもなければ空欄
                $totalText = if ($entry.Completed -gt 0) { Format-Span $entry.Total } else { '' }
                $rows.Add([object[]]@((Protect-ExcelText $entry.Name), $entry.Logons.ToString('N0', $Invariant), (Format-Time $entry.First), (Format-Time $entry.Last), (Protect-ExcelText $totalText),
                    $entry.Failures.ToString('N0', $Invariant), (Format-Time $entry.LastFailure)))
            }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'users.csv')) $UserColumns $rows

            $rows = [System.Collections.Generic.List[object]]::new()
            foreach ($skip in $skipped) { $rows.Add([object[]]@($skip[0], (Protect-ExcelText $skip[1]), (Protect-ExcelText $skip[2]))) }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $rows
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
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
    if (-not $finished -or $saveError -or $fatal) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        $runDir = $null
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

#region 画面表示

foreach ($line in $head) { Write-Host $line }
if ($total -gt 0) {
    # 保存できたときは、画面には先頭の分だけを出す。保存しなかったときは、ほかに見る場所がないためすべて出す
    $limit = if ($runDir) { $ScreenRows } else { 0 }
    foreach ($line in (Get-Body $limit)) {
        if ($line.StartsWith('■', [StringComparison]::Ordinal)) { Write-Host $line -ForegroundColor Cyan } else { Write-Host $line }
    }
}
foreach ($line in $tail) {
    if ($line.StartsWith('※', [StringComparison]::Ordinal)) { Write-Host $line -ForegroundColor Yellow } else { Write-Host $line }
}

$exitCode = if ($total -gt 0) { 0 } else { 2 }
if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
    $exitCode = 1
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
}

#endregion

exit $exitCode
