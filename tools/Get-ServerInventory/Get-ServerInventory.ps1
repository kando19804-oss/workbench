<#
.SYNOPSIS
    サーバーの構成を集め、設計書やパラメータシートに貼れる形の CSV にまとめます。

.DESCRIPTION
    OS・CPU・メモリ・ディスク・ネットワーク・役割と機能・インストール済みソフト・更新プログラム・サービス・
    ローカルユーザーとグループ・共有フォルダと、リモートデスクトップやファイアウォールなどの主な設定を、
    1 回の実行で集めます。主な内容を画面に表示し、すべてを CSV（BOM 付き UTF-8）に保存します。

    読み取りだけを行い、サーバーの設定を変更することはありません。
    プロダクトキー、パスワード、環境変数の値、サービスの実行ファイルと引数は取りません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Get-ServerInventory.ps1

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoSave
    結果を保存せず、画面表示だけにします。-OutputFolder と一緒には使えません。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-ServerInventory.ps1

    このサーバーの構成を集めて、スクリプトと同じフォルダに保存します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-ServerInventory.ps1 -OutputFolder \\fileserver\inventory

    結果を共有フォルダに保存します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-ServerInventory.ps1 -NoSave

    結果を保存せず、画面表示だけにします。

.NOTES
    Get-ServerInventory 1.0.0
    Copyright (c) 2026 結進舎
    Released under the MIT License.
    https://github.com/kando19804-oss/workbench/blob/main/LICENSE

.LINK
    https://github.com/kando19804-oss/workbench
#>
#Requires -Version 5.1
[CmdletBinding(PositionalBinding = $false)]
param(
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

$ScriptVersion = '1.0.0'
$ToolName      = 'Get-ServerInventory'
$ToolTitle     = 'パラメータシート出力'
$CellWidth     = 30       # 画面の表の IP アドレスや DNS サーバーの幅（半角の文字数）。超える分は「ほか N 件」にまとめる。3 列あるので、120 桁に収まるように 30 にする
$ParamColumns  = '分類', '項目', '値', '備考'
$SkipColumns   = '理由', '項目', '詳細'
$TimeFormat    = 'yyyy-MM-dd HH:mm:ss'
$DateFormat    = 'yyyy-MM-dd'
$WindowsAppId  = '55c92734-d682-4d71-983e-d6ec3f16059f'   # SoftwareLicensingProduct の Windows の ApplicationID
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true
$machineName   = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }
# 32 ビットの PowerShell でも 64 ビットのレジストリを読む
$RegView       = if ([Environment]::Is64BitOperatingSystem) { 'Registry64' } else { 'Default' }

# parameters.csv の分類と項目。どのサーバーでも同じ名前・同じ順にする（メモリのスロットとページファイルは、この後ろに足す）
$Categories = [ordered]@{
    'コンピューター'       = @('ホスト名', '完全修飾ドメイン名', 'ドメイン・ワークグループ', 'ドメインの役割', 'メーカー', 'モデル', 'シリアル番号', 'BIOS のバージョン', 'BIOS の日付', 'ファームウェア', 'セキュアブート')
    'OS'                   = @('名前', 'エディション', 'バージョン', 'ビルド', 'アーキテクチャ', 'インストールの種類', 'インストール日', '最終起動', 'Windows フォルダ', 'OS の言語', 'システムロケール', 'タイムゾーン', 'ライセンス認証')
    'CPU'                  = @('名前', 'ソケット数', 'コア数', '論理プロセッサ数', '最大クロック')
    'メモリ'               = @('実装メモリ', '使えるメモリ')
    '仮想メモリ'           = @('管理')
    'メモリダンプ'         = @('種類', '保存先', '自動的に再起動')
    'リモートデスクトップ' = @('接続', 'ネットワークレベル認証', 'ポート')
    'ファイアウォール'     = @('ドメイン', 'プライベート', 'パブリック')
    'UAC'                  = @('UAC', '通知のレベル')
    '時刻同期'             = @('種類', 'NTP サーバー')
    'Windows Update'       = @('WSUS サーバー', '自動更新')
    '電源'                 = @('電源プラン')
    'ネットワーク'         = @('DNS サフィックスの検索一覧')
}

# 値を取る単位。1 つの単位が取れなくても、ほかの単位は取る。Owns はその単位で値を書く項目（取れなかったときに備考を書く）
function Get-ParamKeys([string]$Category, [string[]]$Names) { return , @(foreach ($name in $Names) { $Category + '|' + $name }) }
$ParamUnits = @(
    # Get-ParamKeys は配列を 1 つの値として返すので、@( ) で囲まずに ( ) でつなぐ（@( ) だと配列の配列になる）
    @{ Label = 'コンピューター'; Reader = 'Read-ParamComputer'; Owns = (Get-ParamKeys 'コンピューター' @('ホスト名', '完全修飾ドメイン名', 'ドメイン・ワークグループ', 'ドメインの役割', 'メーカー', 'モデル')) + (Get-ParamKeys 'メモリ' @('使えるメモリ')) }
    @{ Label = 'BIOS'; Reader = 'Read-ParamBios'; Owns = Get-ParamKeys 'コンピューター' @('シリアル番号', 'BIOS のバージョン', 'BIOS の日付') }
    @{ Label = 'ファームウェア'; Reader = 'Read-ParamFirmware'; Owns = Get-ParamKeys 'コンピューター' @('ファームウェア', 'セキュアブート') }
    @{ Label = 'OS'; Reader = 'Read-ParamOs'; Owns = Get-ParamKeys 'OS' @('名前', 'エディション', 'バージョン', 'ビルド', 'アーキテクチャ', 'インストールの種類', 'インストール日', '最終起動', 'Windows フォルダ', 'OS の言語', 'システムロケール') }
    @{ Label = 'タイムゾーン'; Reader = 'Read-ParamTimeZone'; Owns = Get-ParamKeys 'OS' @('タイムゾーン') }
    @{ Label = 'ライセンス認証'; Reader = 'Read-ParamLicense'; Owns = Get-ParamKeys 'OS' @('ライセンス認証') }
    @{ Label = 'CPU'; Reader = 'Read-ParamCpu'; Owns = Get-ParamKeys 'CPU' $Categories['CPU'] }
    @{ Label = 'メモリ'; Reader = 'Read-ParamMemory'; Owns = Get-ParamKeys 'メモリ' @('実装メモリ') }
    @{ Label = '仮想メモリ'; Reader = 'Read-ParamPageFile'; Owns = Get-ParamKeys '仮想メモリ' $Categories['仮想メモリ'] }
    @{ Label = 'メモリダンプ'; Reader = 'Read-ParamDump'; Owns = Get-ParamKeys 'メモリダンプ' $Categories['メモリダンプ'] }
    @{ Label = 'リモートデスクトップ'; Reader = 'Read-ParamRdp'; Owns = Get-ParamKeys 'リモートデスクトップ' $Categories['リモートデスクトップ'] }
    @{ Label = 'ファイアウォール'; Reader = 'Read-ParamFirewall'; Owns = Get-ParamKeys 'ファイアウォール' $Categories['ファイアウォール'] }
    @{ Label = 'UAC'; Reader = 'Read-ParamUac'; Owns = Get-ParamKeys 'UAC' $Categories['UAC'] }
    @{ Label = '時刻同期'; Reader = 'Read-ParamTime'; Owns = Get-ParamKeys '時刻同期' $Categories['時刻同期'] }
    @{ Label = 'Windows Update'; Reader = 'Read-ParamUpdate'; Owns = Get-ParamKeys 'Windows Update' $Categories['Windows Update'] }
    @{ Label = '電源プラン'; Reader = 'Read-ParamPower'; Owns = Get-ParamKeys '電源' $Categories['電源'] }
    @{ Label = 'DNS サフィックスの検索一覧'; Reader = 'Read-ParamDnsSuffix'; Owns = Get-ParamKeys 'ネットワーク' $Categories['ネットワーク'] }
)

# 一覧の CSV。Numeric は数値の列の位置（Excel 向けの ="..." にしない）
$ListDefs = [ordered]@{
    Disk       = @{ Label = 'ディスク'; File = 'disks.csv'; Numeric = @(0, 2, 3)
                    Columns = 'ディスク番号', 'モデル', '容量(GB)', '容量(バイト)', 'パーティションの形式', '接続', '状態' }
    Volume     = @{ Label = 'ボリューム'; File = 'volumes.csv'; Numeric = @(0, 1, 6, 7, 8, 9, 10, 11)
                    Columns = 'ディスク番号', 'パーティション番号', '種類', 'ドライブ', 'ラベル', 'ファイルシステム', '容量(GB)', '容量(バイト)', '空き(GB)', '空き(バイト)', '使用率', 'クラスターサイズ' }
    Network    = @{ Label = 'ネットワーク'; File = 'network.csv'; Numeric = @(16)
                    Columns = '接続の名前', 'アダプター', '状態', '速度', 'MAC アドレス', 'DHCP', 'IPv4 アドレス', 'サブネットマスク', 'ゲートウェイ', 'DNS サーバー', 'IPv6', 'IPv6 アドレス', 'DNS サフィックス', 'DNS に登録', 'NetBIOS over TCP/IP', 'ネットワークの種類', 'MTU' }
    Route      = @{ Label = '静的ルート'; File = 'routes.csv'; Numeric = @(3)
                    Columns = '宛先', 'ゲートウェイ', '接続の名前', 'メトリック' }
    Feature    = @{ Label = '役割と機能'; File = 'features.csv'; Numeric = @()
                    Columns = '名前', '表示名', '種類', '親' }
    Software   = @{ Label = 'インストール済みソフト'; File = 'software.csv'; Numeric = @()
                    Columns = '名前', 'ビット', 'バージョン', '発行元', 'インストール日' }
    HotFix     = @{ Label = '更新プログラム'; File = 'hotfixes.csv'; Numeric = @()
                    Columns = 'KB 番号', '種類', 'インストール日', 'インストールした人' }
    Service    = @{ Label = 'サービス'; File = 'services.csv'; Numeric = @()
                    Columns = 'サービス名', '表示名', 'スタートアップ', '状態', '実行ユーザー' }
    LocalUser  = @{ Label = 'ローカルユーザー'; File = 'localusers.csv'; Numeric = @()
                    Columns = 'ユーザー名', '有効', 'パスワードの期限', '説明' }
    LocalGroup = @{ Label = 'ローカルグループ'; File = 'localgroups.csv'; Numeric = @()
                    Columns = 'グループ名', 'メンバー名' }
    Share      = @{ Label = '共有フォルダ'; File = 'shares.csv'; Numeric = @()
                    Columns = '共有名', 'パス', '種類', '説明' }
}

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

# 日付や数値の区切り文字が、OS の地域設定によって変わらないようにする
function Format-Invariant([string]$Format) {
    return [string]::Format($Invariant, $Format, [object[]]$args)
}

function Format-Duration([TimeSpan]$Span) {
    if ($Span.TotalSeconds -lt 60) { return [string]::Format($Invariant, '{0:N1}秒', $Span.TotalSeconds) }
    if ($Span.TotalHours -lt 1) { return '{0}分{1}秒' -f $Span.Minutes, $Span.Seconds }
    return '{0}時間{1}分{2}秒' -f [math]::Floor($Span.TotalHours), $Span.Minutes, $Span.Seconds
}

# .NET のメソッドで起きた例外は MethodInvocationException に包まれて届くため、元の例外を取り出す。
# InvokeMember（グループのメンバーを読むとき）の例外は TargetInvocationException に包まれる
function Get-InnerException([Exception]$Exception) {
    while (($Exception -is [Management.Automation.ExtendedTypeSystemException] -or $Exception -is [Reflection.TargetInvocationException]) -and $Exception.InnerException) {
        $Exception = $Exception.InnerException
    }
    return $Exception
}

# .NET のメッセージは途中や末尾に改行が入ることがあるため、1 行にまとめる
function Get-ErrorText([Exception]$Exception) {
    $text = $Exception.Message.Replace('\\?\UNC\', '\\').Replace('\\?\', '')
    return ($text -replace '\s*[\r\n]+\s*', ' ').Trim()
}

# 「アクセスが拒否されました」は、UnauthorizedAccessException、SecurityException（レジストリ）、
# HRESULT 0x80070005 の例外、CimException のどれかで届く
function Test-AccessDenied([Exception]$Exception) {
    if ($Exception -is [UnauthorizedAccessException] -or $Exception -is [Security.SecurityException] -or $Exception.HResult -eq -2147024891) { return $true }
    if ($Exception.GetType().Name -eq 'CimException') { return "$($Exception.NativeErrorCode)" -eq 'AccessDenied' }
    return $false
}

# skipped.csv の理由と詳細
function Get-SkipReason([Exception]$Exception) {
    $ex = Get-InnerException $Exception
    if ($ex -is [InvalidOperationException] -and $ex.Message -eq 'NotApplicable') {
        return , @('対象外', 'ドメインコントローラーには、ローカルユーザーとローカルグループがありません')
    }
    if (Test-AccessDenied $ex) { return , @('権限なし', (Get-ErrorText $ex)) }
    return , @('その他のエラー', (Get-ErrorText $ex))
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
# 数値や日付に変えてしまうため、その形の値は ="..." で囲む
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
    return '="' + $Text.Replace('"', '""') + '"'
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
function Format-TextTable([string[]]$Head, [bool[]]$Right, $Rows) {
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

# 「名前  値」の組を、名前の幅をそろえて並べる
function Format-Pairs($Pairs) {
    $width = 0
    foreach ($pair in $Pairs) { $width = [math]::Max($width, (Get-DisplayWidth $pair[0])) }
    return , @(foreach ($pair in $Pairs) { (' ' + (Format-Cell $pair[0] $width $false) + '  ' + $pair[1]).TrimEnd() })
}

# 画面の表に出す値の一覧。数が多いと表の列がすべての行で広がって折り返されるので、幅を超える分は「ほか N 件」にまとめる。
# 最初の 1 つは必ず出す。すべての値は CSV に書く
function Format-ListCell($Values) {
    $list = @(foreach ($value in @($Values)) { if ([string]$value -ne '') { [string]$value } })
    if ($list.Count -eq 0) { return '' }
    $text = $list[0]
    for ($i = 1; $i -lt $list.Count; $i++) {
        $next = $text + ', ' + $list[$i]
        $rest = $list.Count - $i - 1
        $more = if ($rest -gt 0) { Format-Invariant ' ほか {0:N0} 件' $rest } else { '' }
        if ((Get-DisplayWidth ($next + $more)) -gt $CellWidth) { return $text + (Format-Invariant ' ほか {0:N0} 件' ($list.Count - $i)) }
        $text = $next
    }
    return $text
}

function Stop-WithError([string]$Message) {
    $Host.UI.WriteErrorLine($Message)
    exit 1
}

# コマンドプロンプトで "C:\data\" のように最後が \" だと、引用符が閉じずに後ろの引数まで 1 つの値として届く。
# " は Windows のパスに使えない文字なので、値に残っていたらこの問題とみなして止める
function Stop-WithQuoteError([string]$Name, [string]$Value) {
    Stop-WithError (('{0} の値に引用符（"）が残っています: {1}' -f $Name, $Value) + "`r`n" +
        'パスを引用符で囲むときは、最後に \ を付けないでください（例: {0} "C:\tools\結果"）。' -f $Name + '最後が \" になっていると、引用符が閉じずに後ろの引数までその値として扱われます。')
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

# 値を 1 行の文字列にそろえる。途中の改行は空白 1 つに、ほかの制御文字は ? にする（画面や CSV を崩さないため）
function ConvertTo-CellText($Value) {
    if ($null -eq $Value) { return '' }
    $text = (([string]$Value) -replace '\s*[\r\n]+\s*', ' ').Trim()
    return [regex]::Replace($text, '[\x00-\x1F\x7F]', '?')
}

# 配列の値を、並べ替えて 1 つの文字列にする。順序に意味がないものに使う（空の値と重なりは除く）
function Join-Sorted($Values) {
    $list = New-Object 'System.Collections.Generic.List[string]'
    foreach ($value in @($Values)) {
        $text = ConvertTo-CellText $value
        if ($text -ne '' -and -not $list.Contains($text)) { $list.Add($text) }
    }
    $names = $list.ToArray()
    [Array]::Sort($names, [StringComparer]::OrdinalIgnoreCase)
    return $names -join ', '
}

# 順序に意味があるものは、そのままつなぐ
function Join-Ordered($Values) {
    return (@(foreach ($value in @($Values)) { $text = ConvertTo-CellText $value; if ($text -ne '') { $text } }) -join ', ')
}

# 列の順に値を並べた行を作って加える
function Add-Row($Rows) {
    $Rows.Add([string[]]@(foreach ($value in $args) { ConvertTo-CellText $value }))
}

# プロパティがないかもしれない値を読む（OS や PowerShell の版で、ないことがある）
function Get-OptionalValue($Object, [string]$Name) {
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Format-Gb($Bytes) {
    return Format-Invariant '{0:N1} GB' ([double]$Bytes / 1GB)
}

function Format-GbNumber($Bytes) {
    if ($null -eq $Bytes) { return '' }
    return Format-Invariant '{0:N1}' ([double]$Bytes / 1GB)
}

function Format-Number($Value) {
    if ($null -eq $Value) { return '' }
    return Format-Invariant '{0:N0}' ([double]$Value)
}

# 日時（CIM の日時や .NET の DateTime）を書く。読めない値はそのまま
function Format-Time($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return '' }
    try { return ([datetime]$Value).ToString($TimeFormat, $Invariant) } catch { return ConvertTo-CellText $Value }
}

function Format-Date($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return '' }
    try { return ([datetime]$Value).ToString($DateFormat, $Invariant) } catch { return ConvertTo-CellText $Value }
}

# 並べ替え用。読めない日時は最も古いものとして扱う
function Get-SortTime($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return [datetime]::MinValue }
    try { return [datetime]$Value } catch { return [datetime]::MinValue }
}

function Get-CultureName([int]$Lcid) {
    try { return [Globalization.CultureInfo]::GetCultureInfo($Lcid).Name } catch { return [string]$Lcid }
}

# プレフィックス長（24）をサブネットマスク（255.255.255.0）にする
# （0xFFFFFFFF は PowerShell では Int32 の -1 になるので、大きな数は使わず 8 ビットずつ組み立てる）
function ConvertTo-SubnetMask([int]$Prefix) {
    $octets = @(for ($i = 0; $i -lt 4; $i++) {
        $bits = [math]::Max(0, [math]::Min(8, $Prefix - 8 * $i))
        (256 - (1 -shl (8 - $bits))) -band 255
    })
    return $octets -join '.'
}

# IPv4 アドレスを数の順に並べるためのキー。読めないものは後ろに回す
function Get-Ipv4SortKey([string]$Address) {
    $parsed = $null
    if ([version]::TryParse($Address, [ref]$parsed)) { return $parsed }
    return [version]'999.999.999.999'
}

# Windows の上でだけ開けるレジストリ。ビューを指定して HKLM を開く
function Open-MachineKey([string]$SubKey, [string]$View) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::$View)
    return $base.OpenSubKey($SubKey)
}

# HKLM の値を読む。キーや値がなければ Default を返す。Raw のときは %SystemRoot% などを展開しない
function Get-RegValue([string]$SubKey, [string]$Name, $Default, [bool]$Raw = $false) {
    $key = Open-MachineKey $SubKey $RegView
    if (-not $key) { return $Default }
    try {
        $options = if ($Raw) { [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames } else { [Microsoft.Win32.RegistryValueOptions]::None }
        return $key.GetValue($Name, $Default, $options)
    } finally {
        $key.Dispose()
    }
}

# CIM の一覧は、同じクラスを何度も問い合わせないように覚えておく。
# 配列を 1 つの値として返すので、パイプに流すときは、いったん変数に入れてから流す（そのまま流すと一覧全体が 1 件になる）
$cimCache = @{}
function Get-CimList([string]$ClassName) {
    if (-not $cimCache.ContainsKey($ClassName)) { $cimCache[$ClassName] = @(Get-CimInstance -ClassName $ClassName -ErrorAction Stop) }
    return , $cimCache[$ClassName]
}

function Get-FirstCim([string]$ClassName) {
    $list = Get-CimList $ClassName
    if ($list.Count -eq 0) { throw ('{0} の情報がありません' -f $ClassName) }
    return $list[0]
}

#endregion

#region 値の入れ物

$paramValues = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
$paramExtra  = @{}
foreach ($category in $Categories.Keys) { $paramExtra[$category] = New-Object 'System.Collections.Generic.List[string]' }
$skips        = New-Object 'System.Collections.Generic.List[object]'
$currentSkips = $null
$listResults  = @{}

# parameters.csv の値を書く。決まった項目にないもの（スロットやページファイル）は、分類の最後に足す
function Set-Param([string]$Category, [string]$Name, $Value, [string]$Note = '') {
    $key = $Category + '|' + $Name
    if (-not $paramValues.ContainsKey($key) -and $Categories[$Category] -notcontains $Name) { $paramExtra[$Category].Add($Name) }
    $paramValues[$key] = @{ Value = (ConvertTo-CellText $Value); Note = $Note }
}

function Get-ParamValue([string]$Category, [string]$Name) {
    $key = $Category + '|' + $Name
    if ($paramValues.ContainsKey($key)) { return $paramValues[$key].Value }
    return ''
}

# 一部だけ読めなかったものを残す。$currentSkips は単位を読む間だけ差し替える
function Add-PartialSkip([string]$Name, [Exception]$Exception) {
    $reason = Get-SkipReason $Exception
    $currentSkips.Add([object[]]@($reason[0], ('{0}: {1}' -f $Name, $reason[1])))
}

# 補助の問い合わせ。失敗しても単位全体は止めず、空の一覧を返して skipped.csv に残す
function Invoke-Partial([string]$Name, [scriptblock]$Body) {
    try { return , @(& $Body) } catch { Add-PartialSkip $Name $_.Exception; return , @() }
}

#endregion

#region parameters.csv の値

function Read-ParamComputer {
    $cs = Get-FirstCim 'Win32_ComputerSystem'
    $script:domainJoined = [bool]$cs.PartOfDomain
    Set-Param 'コンピューター' 'ホスト名' $cs.Name
    $dnsName = [string]$cs.DNSHostName
    if (-not $dnsName) { $dnsName = [string]$cs.Name }
    # プライマリ DNS サフィックス。名前解決はしない（ネットワークに出ない）
    $suffix = ''
    try { $suffix = [string][Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName } catch { }
    Set-Param 'コンピューター' '完全修飾ドメイン名' $(if ($suffix) { $dnsName + '.' + $suffix } else { $dnsName })
    Set-Param 'コンピューター' 'ドメイン・ワークグループ' $(if ($cs.PartOfDomain) { $cs.Domain } else { $cs.Workgroup })
    $roleNames = @{ 0 = 'スタンドアロンのワークステーション'; 1 = 'メンバーのワークステーション'; 2 = 'スタンドアロンのサーバー'; 3 = 'メンバーサーバー'; 4 = 'ドメインコントローラー'; 5 = 'ドメインコントローラー' }
    $role = [int]$cs.DomainRole
    Set-Param 'コンピューター' 'ドメインの役割' $(if ($roleNames.ContainsKey($role)) { $roleNames[$role] } else { [string]$role })
    Set-Param 'コンピューター' 'メーカー' $cs.Manufacturer
    Set-Param 'コンピューター' 'モデル' $cs.Model
    Set-Param 'メモリ' '使えるメモリ' (Format-Gb $cs.TotalPhysicalMemory)
}

function Read-ParamBios {
    $bios = Get-FirstCim 'Win32_BIOS'
    Set-Param 'コンピューター' 'シリアル番号' $bios.SerialNumber
    Set-Param 'コンピューター' 'BIOS のバージョン' $bios.SMBIOSBIOSVersion
    Set-Param 'コンピューター' 'BIOS の日付' (Format-Date $bios.ReleaseDate)
}

# ファームウェアの種類（1 = BIOS、2 = UEFI）。管理者でなくても読める Windows の API を使う
function Get-FirmwareType {
    if (-not ('WbInventory.NativeMethods' -as [type])) {
        Add-Type -Namespace WbInventory -Name NativeMethods -MemberDefinition '[DllImport("kernel32.dll", SetLastError = true)] public static extern bool GetFirmwareType(out uint firmwareType);'
    }
    $value = [uint32]0
    if (-not [WbInventory.NativeMethods]::GetFirmwareType([ref]$value)) { throw (New-Object ComponentModel.Win32Exception) }
    return [int]$value
}

function Read-ParamFirmware {
    $type = Get-FirmwareType
    if ($type -eq 1) {
        Set-Param 'コンピューター' 'ファームウェア' 'BIOS'
        Set-Param 'コンピューター' 'セキュアブート' '対象外（BIOS）'
        return
    }
    if ($type -ne 2) {
        Set-Param 'コンピューター' 'ファームウェア' '' ('ファームウェアの種類が分かりません（{0}）' -f $type)
        Set-Param 'コンピューター' 'セキュアブート' '' 'ファームウェアの種類が分かりません'
        return
    }
    Set-Param 'コンピューター' 'ファームウェア' 'UEFI'
    $state = Get-RegValue 'SYSTEM\CurrentControlSet\Control\SecureBoot\State' 'UEFISecureBootEnabled' $null
    if ($null -eq $state) {
        Set-Param 'コンピューター' 'セキュアブート' '' 'セキュアブートの情報がありません'
    } else {
        Set-Param 'コンピューター' 'セキュアブート' $(if ([int]$state -eq 1) { '有効' } else { '無効' })
    }
}

function Read-ParamOs {
    $os = Get-FirstCim 'Win32_OperatingSystem'
    $versionKey = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    # レジストリの ProductName は Windows 11 でも「Windows 10」のままなので、名前は Caption を使う
    Set-Param 'OS' '名前' $os.Caption
    Set-Param 'OS' 'エディション' (Get-RegValue $versionKey 'EditionID' '')
    # 21H2 のような表示のバージョン。Windows Server 2016 / 2019 にはないので、1607 のような ReleaseId を使う
    $version = [string](Get-RegValue $versionKey 'DisplayVersion' '')
    if (-not $version) { $version = [string](Get-RegValue $versionKey 'ReleaseId' '') }
    Set-Param 'OS' 'バージョン' $version
    $build = [string](Get-RegValue $versionKey 'CurrentBuild' '')
    if (-not $build) { $build = [string]$os.BuildNumber }
    $ubr = Get-RegValue $versionKey 'UBR' $null
    if ($null -ne $ubr -and $build) { $build += '.' + [string]$ubr }
    Set-Param 'OS' 'ビルド' $build
    $architecture = [string]$os.OSArchitecture
    $bits = [regex]::Match($architecture, '\d+')
    Set-Param 'OS' 'アーキテクチャ' $(if ($bits.Success) { $bits.Value + ' ビット' } else { $architecture })
    $typeNames = @{ 'Server' = 'デスクトップ エクスペリエンス'; 'Server Core' = 'Server Core'; 'Client' = 'クライアント'; 'Nano Server' = 'Nano Server' }
    $installType = [string](Get-RegValue $versionKey 'InstallationType' '')
    Set-Param 'OS' 'インストールの種類' $(if ($typeNames.ContainsKey($installType)) { $typeNames[$installType] } else { $installType })
    Set-Param 'OS' 'インストール日' (Format-Time $os.InstallDate)
    Set-Param 'OS' '最終起動' (Format-Time $os.LastBootUpTime)
    Set-Param 'OS' 'Windows フォルダ' $os.WindowsDirectory
    # インストールしたときの OS の言語。英語版に日本語の言語パックを入れて表示している場合は en-US になる
    Set-Param 'OS' 'OS の言語' (Get-CultureName ([int]$os.OSLanguage))
    # システムロケール（Unicode 対応でないプログラムの言語）は、Nls\Locale の既定の値（00000411 の形）
    $locale = [string](Get-RegValue 'SYSTEM\CurrentControlSet\Control\Nls\Locale' '' '')
    $lcid = 0
    if ([int]::TryParse($locale, [Globalization.NumberStyles]::HexNumber, $Invariant, [ref]$lcid)) {
        Set-Param 'OS' 'システムロケール' (Get-CultureName $lcid)
    } else {
        Set-Param 'OS' 'システムロケール' '' 'システムロケールの情報がありません'
    }
}

function Read-ParamTimeZone {
    Set-Param 'OS' 'タイムゾーン' (Get-FirstCim 'Win32_TimeZone').Caption
}

# ライセンス認証の状態とチャネル。評価版と猶予期間のときは、期限の日付も書き添える。
# KMS の認証は 7 日ごとに自動で更新されて期限が変わるので、書き添えない
function Get-LicenseText($Product, [datetime]$Now) {
    $statusNames = @{ 0 = '未認証'; 1 = '認証済み'; 2 = '猶予期間'; 3 = '猶予期間（ハードウェアの変更）'; 4 = '猶予期間（正規品でない）'; 5 = '通知'; 6 = '猶予期間（延長）' }
    $status = [int]$Product.LicenseStatus
    $text = if ($statusNames.ContainsKey($status)) { $statusNames[$status] } else { '状態 {0}' -f $status }
    $parts = @()
    # Description は「Windows(R) Operating System, VOLUME_KMSCLIENT channel」の形
    $match = [regex]::Match([string]$Product.Description, ',\s*([A-Za-z0-9_]+)\s+channel', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $channel = if ($match.Success) { $match.Groups[1].Value } else { '' }
    if ($channel) { $parts += $channel }
    $grace = [long](Get-OptionalValue $Product 'GracePeriodRemaining')   # 分
    $inGrace = @(2, 3, 4, 6) -contains $status
    if ($grace -gt 0 -and ($channel -match 'EVAL' -or $inGrace)) { $parts += '期限 ' + $Now.AddMinutes($grace).ToString($DateFormat, $Invariant) }
    if ($parts.Count -gt 0) { $text += '（' + ($parts -join '、') + '）' }
    return $text
}

function Read-ParamLicense {
    $filter = "ApplicationID='{0}' AND PartialProductKey IS NOT NULL" -f $WindowsAppId
    $products = @(Get-CimInstance -ClassName SoftwareLicensingProduct -Filter $filter -ErrorAction Stop)
    if ($products.Count -eq 0) {
        Set-Param 'OS' 'ライセンス認証' '' 'Windows のライセンスの情報が見つかりません'
        return
    }
    # ESU などのアドオンのライセンスも同じ ApplicationID で入るので除く（LicenseIsAddon のない古い OS もあるので、WQL では絞らない）
    $mains = @($products | Where-Object { -not (Get-OptionalValue $_ 'LicenseIsAddon') })
    if ($mains.Count -gt 0) { $products = $mains }
    $product = $products[0]
    foreach ($candidate in $products) { if ([int]$candidate.LicenseStatus -eq 1) { $product = $candidate; break } }
    Set-Param 'OS' 'ライセンス認証' (Get-LicenseText $product (Get-Date))
}

function Read-ParamCpu {
    $cpus = Get-CimList 'Win32_Processor'
    if ($cpus.Count -eq 0) { throw 'Win32_Processor の情報がありません' }
    $names   = New-Object 'System.Collections.Generic.List[string]'
    $cores   = 0
    $logical = 0
    foreach ($cpu in $cpus) {
        $name = ConvertTo-CellText $cpu.Name
        if ($name -and -not $names.Contains($name)) { $names.Add($name) }
        $cores   += [int]$cpu.NumberOfCores
        $logical += [int]$cpu.NumberOfLogicalProcessors
    }
    Set-Param 'CPU' '名前' ($names.ToArray() -join ', ')
    Set-Param 'CPU' 'ソケット数' (Format-Number $cpus.Count)
    Set-Param 'CPU' 'コア数' (Format-Number $cores)
    Set-Param 'CPU' '論理プロセッサ数' (Format-Number $logical)
    Set-Param 'CPU' '最大クロック' (Format-Invariant '{0:N0} MHz' ([int]$cpus[0].MaxClockSpeed))
}

# 画面の「8.0 GB × 2」のために、スロットごとの容量を覚えておく
$memorySlots = New-Object 'System.Collections.Generic.List[long]'
function Read-ParamMemory {
    $total = [long]0
    foreach ($module in (Get-CimList 'Win32_PhysicalMemory')) {
        $capacity = [long]$module.Capacity
        $total += $capacity
        $script:memorySlots.Add($capacity)
        $slot = ConvertTo-CellText $module.DeviceLocator
        if (-not $slot) { $slot = ConvertTo-CellText $module.BankLabel }
        if (-not $slot) { $slot = [string]$script:memorySlots.Count }
        # 同じ名前のスロットが重なったときは、番号を付けて分ける
        $item = 'スロット ' + $slot
        $baseItem = $item
        $number = 2
        while ($paramValues.ContainsKey('メモリ|' + $item)) { $item = '{0} ({1})' -f $baseItem, $number; $number++ }
        $parts = @(Format-Gb $capacity)
        $speed = [int](Get-OptionalValue $module 'ConfiguredClockSpeed')
        if ($speed -le 0) { $speed = [int](Get-OptionalValue $module 'Speed') }
        if ($speed -gt 0) { $parts += Format-Invariant '{0:N0} MHz' $speed }
        $maker = (@((ConvertTo-CellText $module.Manufacturer), (ConvertTo-CellText $module.PartNumber)) | Where-Object { $_ }) -join ' '
        if ($maker) { $parts += $maker }
        Set-Param 'メモリ' $item ($parts -join '、')
    }
    if ($total -gt 0) {
        Set-Param 'メモリ' '実装メモリ' (Format-Gb $total)
    } else {
        Set-Param 'メモリ' '実装メモリ' '' 'メモリのスロットの情報がありません'
    }
}

function Read-ParamPageFile {
    $cs = Get-FirstCim 'Win32_ComputerSystem'
    $auto = [bool]$cs.AutomaticManagedPagefile
    Set-Param '仮想メモリ' '管理' $(if ($auto) { '自動' } else { '手動' })
    $usage = @{}
    foreach ($file in (Get-CimList 'Win32_PageFileUsage')) { $usage[[string]$file.Name] = [long]$file.AllocatedBaseSize }
    $done = @{}
    if (-not $auto) {
        # 設定はレジストリの PagingFiles（「C:\pagefile.sys 4096 8192」の形の複数行）から読む。管理者でなくても読める。
        # 「?:\pagefile.sys」は自動の印。サイズが 0 0 か、サイズのない行は「システム管理サイズ」
        $lines = @(Get-RegValue 'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' 'PagingFiles' @())
        foreach ($line in $lines) {
            $text = ([string]$line).Trim()
            if (-not $text -or $text.StartsWith('?:', [StringComparison]::Ordinal)) { continue }
            $sizeMatch = [regex]::Match($text, '^(.+?)\s+(\d+)\s+(\d+)$')
            $name = if ($sizeMatch.Success) { $sizeMatch.Groups[1].Value } else { $text }
            $initial = if ($sizeMatch.Success) { [long]$sizeMatch.Groups[2].Value } else { 0 }
            $maximum = if ($sizeMatch.Success) { [long]$sizeMatch.Groups[3].Value } else { 0 }
            $done[$name] = $true
            $current = if ($usage.ContainsKey($name)) { Format-Invariant '（現在 {0:N0} MB）' $usage[$name] } else { '' }
            $sizeText = if ($initial -eq 0 -and $maximum -eq 0) { 'システム管理サイズ' } else { Format-Invariant '初期 {0:N0} MB、最大 {1:N0} MB' $initial $maximum }
            Set-Param '仮想メモリ' ('ページファイル ' + $name) ($sizeText + $current)
        }
    }
    $names = @($usage.Keys | Sort-Object)
    foreach ($name in $names) {
        if ($done.ContainsKey($name)) { continue }
        $done[$name] = $true
        $head = if ($auto) { '自動' } else { '設定なし' }
        Set-Param '仮想メモリ' ('ページファイル ' + $name) (Format-Invariant '{0}（現在 {1:N0} MB）' $head $usage[$name])
    }
    if ($done.Count -eq 0) { Set-Param '仮想メモリ' 'ページファイル' 'なし' }
}

function Read-ParamDump {
    $key = 'SYSTEM\CurrentControlSet\Control\CrashControl'
    $enabled = Get-RegValue $key 'CrashDumpEnabled' $null
    if ($null -eq $enabled) {
        Set-Param 'メモリダンプ' '種類' '' 'メモリダンプの設定がありません'
        Set-Param 'メモリダンプ' '保存先' '' 'メモリダンプの設定がありません'
    } else {
        $kind = [int]$enabled
        $kindText = switch ($kind) {
            0 { 'なし' }
            1 { if ([int](Get-RegValue $key 'FilterPages' 0) -eq 1) { 'アクティブメモリダンプ' } else { '完全メモリダンプ' } }
            2 { 'カーネルメモリダンプ' }
            3 { '最小メモリダンプ' }
            7 { '自動メモリダンプ' }
            default { '種類 {0}' -f $kind }
        }
        Set-Param 'メモリダンプ' '種類' $kindText
        # %SystemRoot% などは展開せず、設定のとおりに書く。最小メモリダンプはフォルダに保存する
        $place = ''
        if ($kind -eq 3) { $place = Get-RegValue $key 'MinidumpDir' '' $true } elseif ($kind -ne 0) { $place = Get-RegValue $key 'DumpFile' '' $true }
        Set-Param 'メモリダンプ' '保存先' $place
    }
    Set-Param 'メモリダンプ' '自動的に再起動' $(if ([int](Get-RegValue $key 'AutoReboot' 1) -eq 1) { 'する' } else { 'しない' })
}

function Read-ParamRdp {
    $policyKey = 'SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
    $tcpKey    = 'SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'
    # グループポリシーで指定されていれば、そちらが効く
    $deny = Get-RegValue $policyKey 'fDenyTSConnections' $null
    $denyNote = 'グループポリシー'
    if ($null -eq $deny) {
        $deny = Get-RegValue 'SYSTEM\CurrentControlSet\Control\Terminal Server' 'fDenyTSConnections' 1
        $denyNote = ''
    }
    Set-Param 'リモートデスクトップ' '接続' $(if ([int]$deny -eq 0) { '許可' } else { '許可しない' }) $denyNote
    $nla = Get-RegValue $policyKey 'UserAuthentication' $null
    $nlaNote = 'グループポリシー'
    if ($null -eq $nla) {
        $nla = Get-RegValue $tcpKey 'UserAuthentication' 1
        $nlaNote = ''
    }
    Set-Param 'リモートデスクトップ' 'ネットワークレベル認証' $(if ([int]$nla -eq 1) { '必要' } else { '不要' }) $nlaNote
    $port = Get-RegValue $tcpKey 'PortNumber' $null
    if ($null -eq $port) {
        Set-Param 'リモートデスクトップ' 'ポート' '' 'ポートの設定がありません'
    } else {
        Set-Param 'リモートデスクトップ' 'ポート' ([string][int]$port)
    }
}

# 実際に効いている設定（グループポリシーを含む）は管理者でないと読めないので、管理者でないときはこの PC の設定を読む
function Read-ParamFirewall {
    if ($isAdmin) {
        $profiles = @(Get-NetFirewallProfile -PolicyStore ActiveStore -ErrorAction Stop)
        $note = ''
    } else {
        $profiles = @(Get-NetFirewallProfile -ErrorAction Stop)
        $note = 'この PC の設定。グループポリシーの分を含みません'
    }
    $names = @{ 'Domain' = 'ドメイン'; 'Private' = 'プライベート'; 'Public' = 'パブリック' }
    foreach ($profile in $profiles) {
        $name = [string]$profile.Name
        if (-not $names.ContainsKey($name)) { continue }
        $enabled = switch ("$($profile.Enabled)") { 'True' { '有効' } 'False' { '無効' } 'NotConfigured' { '未構成' } default { "$($profile.Enabled)" } }
        Set-Param 'ファイアウォール' $names[$name] $enabled $note
    }
}

function Read-ParamUac {
    $key = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
    if ([int](Get-RegValue $key 'EnableLUA' 1) -eq 0) {
        Set-Param 'UAC' 'UAC' '無効'
        Set-Param 'UAC' '通知のレベル' '' 'UAC が無効です'
        return
    }
    Set-Param 'UAC' 'UAC' '有効'
    # コントロールパネルの 4 段階のどれに当たるか
    $consent = [int](Get-RegValue $key 'ConsentPromptBehaviorAdmin' 5)
    $secure  = [int](Get-RegValue $key 'PromptOnSecureDesktop' 1)
    if ($consent -eq 2 -and $secure -eq 1) { $level = '常に通知する' }
    elseif ($consent -eq 5 -and $secure -eq 1) { $level = '既定' }
    elseif ($consent -eq 5 -and $secure -eq 0) { $level = '既定（デスクトップを暗転しない）' }
    elseif ($consent -eq 0 -and $secure -eq 0) { $level = '通知しない' }
    else { $level = 'カスタム（ConsentPromptBehaviorAdmin={0}、PromptOnSecureDesktop={1}）' -f $consent, $secure }
    Set-Param 'UAC' '通知のレベル' $level
}

# 時刻同期の設定。W32Time のサービスが止まっていると w32tm /query は失敗するので、レジストリを読む
function Read-ParamTime {
    $policyKey = 'SOFTWARE\Policies\Microsoft\W32Time\Parameters'
    $localKey  = 'SYSTEM\CurrentControlSet\Services\W32Time\Parameters'
    $typeNames = @{ 'NTP' = 'NTP'; 'NT5DS' = 'ドメインの階層（NT5DS）'; 'AllSync' = 'NTP とドメインの階層（AllSync）'; 'NoSync' = '同期しない（NoSync）' }
    $type = Get-RegValue $policyKey 'Type' $null
    $typeNote = 'グループポリシー'
    if ($null -eq $type) {
        $type = Get-RegValue $localKey 'Type' ''
        $typeNote = 'この PC の設定'
    }
    $typeText = [string]$type
    if (-not $typeText) { $typeNote = '時刻同期の設定がありません' }
    Set-Param '時刻同期' '種類' $(if ($typeNames.ContainsKey($typeText)) { $typeNames[$typeText] } else { $typeText }) $typeNote
    $server = Get-RegValue $policyKey 'NtpServer' $null
    $serverNote = 'グループポリシー'
    if ($null -eq $server) {
        $server = Get-RegValue $localKey 'NtpServer' ''
        $serverNote = 'この PC の設定'
    }
    # ドメインに参加したサーバーは、種類が NT5DS でも既定の time.windows.com が残っていることが多い。使われていると読まれないように書き添える
    if (-not [string]$server) { $serverNote = '時刻同期の設定がありません' }
    elseif ($typeText -eq 'NT5DS' -or $typeText -eq 'NoSync') { $serverNote += '。種類が {0} のため使われません' -f $typeText }
    Set-Param '時刻同期' 'NTP サーバー' $server $serverNote
}

function Read-ParamUpdate {
    $policyKey = 'SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $auKey     = $policyKey + '\AU'
    $server    = [string](Get-RegValue $policyKey 'WUServer' '')
    $useServer = [int](Get-RegValue $auKey 'UseWUServer' 0)
    if ($server -and $useServer -eq 1) {
        Set-Param 'Windows Update' 'WSUS サーバー' $server 'グループポリシー'
    } elseif ($server) {
        Set-Param 'Windows Update' 'WSUS サーバー' 'なし' ('WUServer（{0}）はありますが、UseWUServer が 1 でないため使われません' -f $server)
    } else {
        Set-Param 'Windows Update' 'WSUS サーバー' 'なし'
    }
    $noAuto  = Get-RegValue $auKey 'NoAutoUpdate' $null
    $options = Get-RegValue $auKey 'AUOptions' $null
    if ($null -ne $noAuto -and [int]$noAuto -eq 1) {
        Set-Param 'Windows Update' '自動更新' '無効' 'グループポリシー'
    } elseif ($null -ne $options) {
        $option = [int]$options
        $text = switch ($option) {
            2 { 'ダウンロードとインストールを通知' }
            3 { '自動ダウンロードして通知' }
            4 { '自動ダウンロードしてインストールを予約' }
            5 { 'ローカルの管理者が選ぶ' }
            7 { '自動ダウンロードして、インストールと再起動を通知' }
            default { 'AUOptions={0}' -f $option }
        }
        if ($option -eq 4) {
            # 0 = 毎日、1～7 = 日曜～土曜
            $day  = [int](Get-RegValue $auKey 'ScheduledInstallDay' 0)
            $hour = [int](Get-RegValue $auKey 'ScheduledInstallTime' 3)
            $dayNames = '日', '月', '火', '水', '木', '金', '土'
            $dayText = if ($day -ge 1 -and $day -le 7) { '毎週{0}曜' -f $dayNames[$day - 1] } else { '毎日' }
            $text += '（{0} {1}:00）' -f $dayText, $hour
        }
        Set-Param 'Windows Update' '自動更新' $text 'グループポリシー'
    } else {
        Set-Param 'Windows Update' '自動更新' 'ポリシーなし'
    }
}

# 電源プラン。Win32_PowerPlan は管理者でないと読めないので、powercfg を使う
function Read-ParamPower {
    $exe = Join-Path $env:SystemRoot 'System32\powercfg.exe'
    # 外部コマンドの標準エラーを 2>&1 で受けると、5.1 は $ErrorActionPreference が Stop のとき例外にするため、ここだけ変える
    $ErrorActionPreference = 'Continue'
    $output = @(& $exe /getactivescheme 2>&1)
    $exitCode = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    $text = ConvertTo-CellText (@(foreach ($line in $output) { [string]$line }) -join ' ')
    if ($exitCode -ne 0) { throw ('powercfg が失敗しました（終了コード {0}）: {1}' -f $exitCode, $text) }
    # 「電源設定の GUID: 381b4222-...  (バランス)」の形。言葉は OS の表示の言語で変わるので、GUID とかっこで読む
    $guid = '[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}'
    $match = [regex]::Match($text, '(' + $guid + ')\s*\((.*)\)\s*$')
    if ($match.Success) {
        Set-Param '電源' '電源プラン' $match.Groups[2].Value.Trim()
        return
    }
    $only = [regex]::Match($text, $guid)
    if (-not $only.Success) { throw ('powercfg の結果を読めません: {0}' -f $text) }
    Set-Param '電源' '電源プラン' $only.Value '電源プランの名前が分かりません'
}

function Read-ParamDnsSuffix {
    $text = Join-Ordered (Get-DnsClientGlobalSetting -ErrorAction Stop).SuffixSearchList
    Set-Param 'ネットワーク' 'DNS サフィックスの検索一覧' $(if ($text) { $text } else { 'なし' })
}

#endregion

#region 一覧

function Read-ListDisk($Rows) {
    $statusNames = @{ 'Online' = 'オンライン'; 'Offline' = 'オフライン' }
    foreach ($disk in @(Get-Disk -ErrorAction Stop | Sort-Object Number)) {
        $status = Join-Ordered @(foreach ($state in @($disk.OperationalStatus)) { $t = [string]$state; if ($statusNames.ContainsKey($t)) { $statusNames[$t] } else { $t } })
        Add-Row $Rows $disk.Number $disk.FriendlyName (Format-GbNumber $disk.Size) (Format-Number $disk.Size) "$($disk.PartitionStyle)" "$($disk.BusType)" $status
    }
}

# パーティションごとに 1 行。ボリュームの情報（ラベル、空きなど）は Win32_Volume から、ボリュームの GUID のパスで突き合わせる
function Read-ListVolume($Rows) {
    $volumes = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($volume in (Get-CimList 'Win32_Volume')) { $volumes[[string]$volume.DeviceID] = $volume }
    $typeNames = @{
        'System' = 'EFI システム'; 'Reserved' = '予約（MSR）'; 'Basic' = '基本'; 'Recovery' = '回復'
        'IFS' = '基本'; 'FAT32' = '基本'; 'FAT16' = '基本'; 'FAT12' = '基本'; 'Huge' = '基本'
        'Extended' = '拡張'; 'Logical' = '論理ドライブ'; 'Unknown' = '不明'
    }
    foreach ($partition in @(Get-Partition -ErrorAction Stop | Sort-Object DiskNumber, PartitionNumber)) {
        $type = [string]$partition.Type
        $typeText = if ($typeNames.ContainsKey($type)) { $typeNames[$type] } else { $type }
        if ($partition.IsActive) { $typeText += '（アクティブ）' }
        $drives = New-Object 'System.Collections.Generic.List[string]'
        $volume = $null
        foreach ($path in @($partition.AccessPaths)) {
            $text = [string]$path
            if (-not $text) { continue }
            if ($text.StartsWith('\\?\Volume', [StringComparison]::OrdinalIgnoreCase)) {
                if ($volumes.ContainsKey($text)) { $volume = $volumes[$text] }
                continue
            }
            # ドライブ文字は「C:」、マウント先のフォルダは末尾の \ を外す
            if ($text -match '^[A-Za-z]:\\$') { $drives.Add($text.Substring(0, 2)) } else { $drives.Add($text.TrimEnd('\')) }
        }
        $capacity = [long]$partition.Size
        $label = ''; $fileSystem = ''; $free = $null; $cluster = $null
        if ($volume) {
            $label      = $volume.Label
            $fileSystem = $volume.FileSystem
            if ($null -ne $volume.Capacity) { $capacity = [long]$volume.Capacity }
            if ($null -ne $volume.FreeSpace) { $free = [long]$volume.FreeSpace }
            if ($null -ne $volume.BlockSize) { $cluster = [long]$volume.BlockSize }
        }
        $usage = ''
        if ($null -ne $free -and $capacity -gt 0) { $usage = Format-Invariant '{0:N1}%' (100.0 * ($capacity - $free) / $capacity) }
        Add-Row $Rows $partition.DiskNumber $partition.PartitionNumber $typeText ($drives.ToArray() -join ', ') $label $fileSystem `
            (Format-GbNumber $capacity) (Format-Number $capacity) (Format-GbNumber $free) (Format-Number $free) $usage (Format-Number $cluster)
    }
}

# 画面のネットワークの表のために、アダプターごとの主な値を覚えておく
$networkScreen = New-Object 'System.Collections.Generic.List[object]'

# アダプターごとに 1 行。IPv4 の設定を中心に、いくつかのコマンドの結果をインターフェイスの番号で突き合わせる。
# 補助のコマンドが失敗しても、アダプターの一覧は出す（その列は空になり、skipped.csv に残る）
function Read-ListNetwork($Rows) {
    $adapters = @(Get-NetAdapter -ErrorAction Stop | Sort-Object Name)
    # 絞り込みのパラメーターを付けると、該当なしのときにエラーになるコマンドがあるので、すべて取ってから絞る
    $addresses   = Invoke-Partial 'Get-NetIPAddress' { Get-NetIPAddress -ErrorAction Stop }
    $interfaces  = Invoke-Partial 'Get-NetIPInterface' { Get-NetIPInterface -ErrorAction Stop }
    $routes      = Invoke-Partial 'Get-NetRoute' { Get-NetRoute -ErrorAction Stop }
    $dnsServers  = Invoke-Partial 'Get-DnsClientServerAddress' { Get-DnsClientServerAddress -ErrorAction Stop }
    $dnsClients  = Invoke-Partial 'Get-DnsClient' { Get-DnsClient -ErrorAction Stop }
    $bindings    = Invoke-Partial 'Get-NetAdapterBinding' { Get-NetAdapterBinding -ErrorAction Stop }
    $profiles    = Invoke-Partial 'Get-NetConnectionProfile' { Get-NetConnectionProfile -ErrorAction Stop }
    $configs     = Invoke-Partial 'Win32_NetworkAdapterConfiguration' { Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ErrorAction Stop }

    $statusNames   = @{ 'Up' = '接続'; 'Disconnected' = '切断'; 'Disabled' = '無効'; 'Not Present' = 'なし' }
    $categoryNames = @{ 'DomainAuthenticated' = 'ドメイン'; 'Private' = 'プライベート'; 'Public' = 'パブリック' }
    $netbiosNames  = @{ 0 = '既定（DHCP サーバーの設定に従う）'; 1 = '有効'; 2 = '無効' }

    foreach ($adapter in $adapters) {
        $index = [int]$adapter.InterfaceIndex
        $name  = [string]$adapter.Name
        $status = [string]$adapter.Status
        $statusText = if ($statusNames.ContainsKey($status)) { $statusNames[$status] } else { $status }
        $speed = if ($status -eq 'Up') { [string]$adapter.LinkSpeed } else { '' }

        # IPv4 は数の順に、サブネットマスクも同じ順に。IPv6 はリンクローカル（fe80::）と、数時間ごとに変わる一時アドレスを除く
        $v4 = @($addresses | Where-Object { [int]$_.InterfaceIndex -eq $index -and "$($_.AddressFamily)" -eq 'IPv4' } | Sort-Object { Get-Ipv4SortKey ([string]$_.IPAddress) })
        $v6 = @($addresses | Where-Object { [int]$_.InterfaceIndex -eq $index -and "$($_.AddressFamily)" -eq 'IPv6' -and -not ([string]$_.IPAddress).StartsWith('fe80:', [StringComparison]::OrdinalIgnoreCase) -and "$($_.SuffixOrigin)" -ne 'Random' })
        $ipText   = Join-Ordered @(foreach ($a in $v4) { $a.IPAddress })
        $maskText = Join-Ordered @(foreach ($a in $v4) { ConvertTo-SubnetMask ([int]$a.PrefixLength) })
        $ipScreen = @(foreach ($a in $v4) { '{0}/{1}' -f $a.IPAddress, $a.PrefixLength })
        $v6Text   = Join-Sorted @(foreach ($a in $v6) { $a.IPAddress })

        $dhcp = ''; $mtu = ''
        foreach ($interface in $interfaces) {
            if ([int]$interface.InterfaceIndex -ne $index -or "$($interface.AddressFamily)" -ne 'IPv4') { continue }
            $dhcp = if ("$($interface.Dhcp)" -eq 'Enabled') { '有効' } else { '無効' }
            $mtu  = Format-Number $interface.NlMtu
        }
        # 画面と CSV で同じ順にするため、並べ替えてから両方に使う
        $gatewayText = Join-Sorted @(foreach ($route in $routes) {
            if ([int]$route.InterfaceIndex -eq $index -and [string]$route.DestinationPrefix -eq '0.0.0.0/0' -and [string]$route.NextHop -ne '0.0.0.0') { [string]$route.NextHop }
        })
        $gateways = @($gatewayText -split ', ' | Where-Object { $_ })
        $dnsList = @()
        foreach ($server in $dnsServers) {
            if ([int]$server.InterfaceIndex -eq $index -and "$($server.AddressFamily)" -eq '2') { $dnsList = @($server.ServerAddresses) }
        }
        # AddressFamily は 2（IPv4）や 23（IPv6）の数で届くことも、IPv4 の名前で届くこともある
        if ($dnsList.Count -eq 0) {
            foreach ($server in $dnsServers) {
                if ([int]$server.InterfaceIndex -eq $index -and "$($server.AddressFamily)" -eq 'IPv4') { $dnsList = @($server.ServerAddresses) }
            }
        }
        $dnsText = Join-Ordered $dnsList
        $ipv6 = ''
        foreach ($binding in $bindings) {
            if ([string]$binding.Name -eq $name -and [string]$binding.ComponentID -eq 'ms_tcpip6') { $ipv6 = if ($binding.Enabled) { '有効' } else { '無効' } }
        }
        $suffix = ''; $register = ''
        foreach ($client in $dnsClients) {
            if ([int]$client.InterfaceIndex -ne $index) { continue }
            $suffix   = [string]$client.ConnectionSpecificSuffix
            $register = if ($client.RegisterThisConnectionsAddress) { 'する' } else { 'しない' }
        }
        $netbios = ''
        foreach ($config in $configs) {
            if ([int](Get-OptionalValue $config 'InterfaceIndex') -ne $index) { continue }
            $option = Get-OptionalValue $config 'TcpipNetbiosOptions'
            if ($null -ne $option) { $netbios = if ($netbiosNames.ContainsKey([int]$option)) { $netbiosNames[[int]$option] } else { [string]$option } }
        }
        $category = ''
        foreach ($profile in $profiles) {
            if ([int]$profile.InterfaceIndex -ne $index) { continue }
            $value = "$($profile.NetworkCategory)"
            $category = if ($categoryNames.ContainsKey($value)) { $categoryNames[$value] } else { $value }
        }

        Add-Row $Rows $name $adapter.InterfaceDescription $statusText $speed $adapter.MacAddress $dhcp $ipText $maskText $gatewayText $dnsText $ipv6 $v6Text $suffix $register $netbios $category $mtu
        $script:networkScreen.Add([string[]]@((ConvertTo-CellText $name), $statusText, (Format-ListCell $ipScreen), (Format-ListCell $gateways), (Format-ListCell $dnsList), $dhcp))
    }
}

# 再起動しても残る静的ルート（route -p や New-NetRoute で足したもの）
function Read-ListRoute($Rows) {
    foreach ($route in @(Get-NetRoute -PolicyStore PersistentStore -ErrorAction Stop | Sort-Object DestinationPrefix, NextHop)) {
        Add-Row $Rows $route.DestinationPrefix $route.NextHop $route.InterfaceAlias (Format-Number $route.RouteMetric)
    }
}

# Windows Server かどうか（ProductType 1 はクライアント、2 はドメインコントローラー、3 はサーバー）
function Test-ServerOs {
    return [int](Get-FirstCim 'Win32_OperatingSystem').ProductType -ne 1
}

# Windows Server はサーバーマネージャーの「役割と機能」（Get-WindowsFeature）、Windows 10 / 11 は「Windows の機能」（Win32_OptionalFeature）
function Read-ListFeature($Rows) {
    if (Test-ServerOs) {
        # PowerShell 7 は、Windows PowerShell の互換の仕組みでモジュールを読み込み、その旨の警告を出すので抑える
        Import-Module ServerManager -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false
        $typeNames = @{ 'Role' = '役割'; 'Role Service' = '役割サービス'; 'Feature' = '機能' }
        foreach ($feature in @(Get-WindowsFeature -ErrorAction Stop -WarningAction SilentlyContinue)) {
            if (-not $feature.Installed) { continue }
            $type = [string]$feature.FeatureType
            Add-Row $Rows $feature.Name $feature.DisplayName $(if ($typeNames.ContainsKey($type)) { $typeNames[$type] } else { $type }) (Get-OptionalValue $feature 'Parent')
        }
        return
    }
    foreach ($feature in @(Get-CimInstance -ClassName Win32_OptionalFeature -Filter 'InstallState=1' -ErrorAction Stop | Sort-Object Name)) {
        Add-Row $Rows $feature.Name $feature.Caption 'Windows の機能' ''
    }
}

# 「プログラムと機能」に出るもの（コンピューター全体にインストールしたもの）。同じ名前の複数のバージョンは 1 行にまとめる
function Read-ListSoftware($Rows) {
    $views = New-Object 'System.Collections.Generic.List[object]'
    if ([Environment]::Is64BitOperatingSystem) { $views.Add([string[]]@('Registry64', '64 ビット')) }
    $views.Add([string[]]@('Registry32', '32 ビット'))
    foreach ($view in $views) {
        $root = Open-MachineKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' $view[0]
        if (-not $root) { continue }
        $found = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
        try {
            foreach ($subName in $root.GetSubKeyNames()) {
                $key = $null
                try {
                    $key = $root.OpenSubKey($subName)
                    if (-not $key) { continue }
                    $displayName = ConvertTo-CellText $key.GetValue('DisplayName', '')
                    # 「プログラムと機能」に出ないもの（システムの部品と、別のソフトの更新）は除く
                    if (-not $displayName -or [int]$key.GetValue('SystemComponent', 0) -eq 1 -or [string]$key.GetValue('ParentKeyName', '')) { continue }
                    if (-not $found.ContainsKey($displayName)) {
                        $found[$displayName] = [pscustomobject]@{
                            Versions = New-Object 'System.Collections.Generic.List[string]'
                            Dates = New-Object 'System.Collections.Generic.List[string]'
                            Publisher = [string]$key.GetValue('Publisher', '')
                        }
                    }
                    $found[$displayName].Versions.Add([string]$key.GetValue('DisplayVersion', ''))
                    # インストール日は 20261015 の形
                    $date = ([string]$key.GetValue('InstallDate', '')).Trim()
                    if ($date -match '^(\d{4})(\d{2})(\d{2})$') { $date = '{0}-{1}-{2}' -f $Matches[1], $Matches[2], $Matches[3] }
                    $found[$displayName].Dates.Add($date)
                } catch {
                    Add-PartialSkip $subName $_.Exception
                } finally {
                    if ($key) { $key.Dispose() }
                }
            }
        } finally {
            $root.Dispose()
        }
        foreach ($displayName in $found.get_Keys()) {
            $entry = $found[$displayName]
            Add-Row $Rows $displayName $view[1] (Join-Sorted $entry.Versions.ToArray()) $entry.Publisher (Join-Sorted $entry.Dates.ToArray())
        }
    }
    # 名前の順に並べる
    $sorted = @($Rows.ToArray() | Sort-Object { $_[0] }, { $_[1] })
    $Rows.Clear()
    foreach ($row in $sorted) { $Rows.Add($row) }
}

# 画面の「最新」のために、いちばん新しい更新プログラムを覚えておく
$latestHotFix = $null
function Read-ListHotFix($Rows) {
    $hotfixes = Get-CimList 'Win32_QuickFixEngineering'
    $sorted = @($hotfixes | Sort-Object @{ Expression = { Get-SortTime $_.InstalledOn }; Descending = $true }, @{ Expression = { [string]$_.HotFixID } })
    foreach ($hotfix in $sorted) {
        Add-Row $Rows $hotfix.HotFixID $hotfix.Description (Format-Date $hotfix.InstalledOn) $hotfix.InstalledBy
        if ($null -eq $script:latestHotFix -and (Get-SortTime $hotfix.InstalledOn) -ne [datetime]::MinValue) {
            $script:latestHotFix = '{0}  {1}' -f (ConvertTo-CellText $hotfix.HotFixID), (Format-Date $hotfix.InstalledOn)
        }
    }
}

# サービスのレジストリから、種類（Type）とトリガーの有無を調べる。
# Type の SERVICE_USER_SERVICE（0x40）はユーザーごとのサービス、SERVICE_USERSERVICE_INSTANCE（0x80）はログオンごとの実体（末尾が変わる）。
# TriggerInfo があるサービスは、Windows が必要なときに開始する（services.msc の「トリガー開始」）
function Get-ServiceKeyInfo([string]$Name) {
    $info = [pscustomobject]@{ Bits = 0; Trigger = $false }
    try {
        $key = Open-MachineKey ('SYSTEM\CurrentControlSet\Services\' + $Name) $RegView
        if ($key) {
            try {
                $info.Bits = [int]$key.GetValue('Type', 0)
                $triggerKey = $key.OpenSubKey('TriggerInfo')
                if ($triggerKey) { $info.Trigger = $true; $triggerKey.Dispose() }
            } finally {
                $key.Dispose()
            }
        }
    } catch { }
    return $info
}

# サービス。実行ファイルと引数は、パスワードが書かれていることがあるので取らない
function Read-ListService($Rows) {
    $stateNames = @{ 'Running' = '実行中'; 'Stopped' = '停止'; 'Paused' = '一時停止'; 'Start Pending' = '開始中'; 'Stop Pending' = '停止中'; 'Pause Pending' = '一時停止中'; 'Continue Pending' = '再開中' }
    $modeNames  = @{ 'Auto' = '自動'; 'Manual' = '手動'; 'Disabled' = '無効'; 'Boot' = 'ブート'; 'System' = 'システム' }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $list = New-Object 'System.Collections.Generic.List[object]'
    $unreadable = New-Object 'System.Collections.Generic.List[string]'
    foreach ($service in (Get-CimList 'Win32_Service')) {
        $name  = [string]$service.Name
        $state = [string]$service.State
        if ($stateNames.ContainsKey($state)) { $state = $stateNames[$state] }
        $mode = [string]$service.StartMode
        $user = [string]$service.StartName
        $displayName = [string]$service.DisplayName
        # Windows は、設定を読む権限のないサービス（管理者でないときの LSM など）を、スタートアップ Unknown・表示名はサービス名・実行ユーザーは空で返す。
        # 表示名はサービス名で代用されているだけなので書かず、skipped.csv に残す
        $isUnreadable = $mode -eq 'Unknown' -or -not $mode
        if ($isUnreadable) {
            $mode = '不明'
            $displayName = ''
            $user = ''
        }
        if ($modeNames.ContainsKey($mode)) { $mode = $modeNames[$mode] }
        $keyInfo = Get-ServiceKeyInfo $name
        # services.msc と同じく、遅延開始とトリガー開始を書き添える（「自動（遅延・トリガー）」「手動（トリガー）」）
        $marks = @()
        if ($mode -eq '自動' -and (Get-OptionalValue $service 'DelayedAutoStart')) { $marks += '遅延' }
        if ($keyInfo.Trigger -and ($mode -eq '自動' -or $mode -eq '手動')) { $marks += 'トリガー' }
        if ($marks.Count -gt 0) { $mode = '{0}（{1}）' -f $mode, ($marks -join '・') }
        # ユーザーごとのサービスは、ログオンのたびに名前の末尾と状態が変わるので、末尾を除いた名前にまとめ、状態と実行ユーザーは書かない。
        # 表示名にも同じ末尾が付く（「Agent Activation Runtime_2f882f」）ので外す
        $bits = $keyInfo.Bits
        if (($bits -band 0x40) -ne 0) {
            if (($bits -band 0x80) -ne 0) {
                $suffixMatch = [regex]::Match($name, '_[0-9a-fA-F]+$')
                if ($suffixMatch.Success) {
                    $name = $name.Substring(0, $suffixMatch.Index)
                    if ($displayName.EndsWith($suffixMatch.Value, [StringComparison]::OrdinalIgnoreCase)) { $displayName = $displayName.Substring(0, $displayName.Length - $suffixMatch.Value.Length) }
                }
            }
            $state = ''
            $user  = ''
        }
        if (-not $seen.Add($name)) { continue }
        if ($isUnreadable) { $unreadable.Add($name) }
        Add-Row $list $name $displayName $mode $state $user
    }
    foreach ($row in @($list.ToArray() | Sort-Object { $_[0] })) { $Rows.Add($row) }
    foreach ($name in @($unreadable.ToArray() | Sort-Object)) {
        Add-PartialSkip $name (New-Object UnauthorizedAccessException 'スタートアップ・表示名・実行ユーザーを読めません（サービスの設定を読む権限がありません）')
    }
}

# ドメインコントローラーには、ローカルユーザーがない
function Assert-NotDomainController {
    if ([int](Get-FirstCim 'Win32_ComputerSystem').DomainRole -ge 4) { throw (New-Object InvalidOperationException 'NotApplicable') }
}

function Read-ListLocalUser($Rows) {
    Assert-NotDomainController
    foreach ($user in @(Get-CimInstance -ClassName Win32_UserAccount -Filter 'LocalAccount=TRUE' -ErrorAction Stop | Sort-Object Name)) {
        Add-Row $Rows $user.Name $(if ($user.Disabled) { '無効' } else { '有効' }) $(if ($user.PasswordExpires) { 'あり' } else { 'なし' }) $user.Description
    }
}

# グループのメンバーを、WinNT プロバイダーで読む（Get-LocalGroupMember は、名前を解決できない SID があると失敗する）
function Get-GroupMemberNames([string]$GroupName) {
    $group = [ADSI]('WinNT://./{0},group' -f $GroupName)
    foreach ($member in @($group.Invoke('Members'))) {
        $adsPath = [string]$member.GetType().InvokeMember('AdsPath', 'GetProperty', $null, $member, $null)
        # WinNT://DOMAIN/user、WinNT://WORKGROUP/HOST/user、解決できないものは WinNT://S-1-5-...
        $parts = @($adsPath -replace '^WinNT://', '' -split '/')
        if ($parts.Count -ge 2) { $parts[$parts.Count - 2] + '\' + $parts[$parts.Count - 1] } else { $parts[0] }
    }
}

function Read-ListLocalGroup($Rows) {
    Assert-NotDomainController
    foreach ($group in @(Get-CimInstance -ClassName Win32_Group -Filter 'LocalAccount=TRUE' -ErrorAction Stop | Sort-Object Name)) {
        $groupName = [string]$group.Name
        try {
            foreach ($memberName in @(Get-GroupMemberNames $groupName | Sort-Object)) { Add-Row $Rows $groupName $memberName }
        } catch {
            Add-PartialSkip $groupName $_.Exception
        }
    }
}

function Read-ListShare($Rows) {
    $shares = Get-CimList 'Win32_Share'
    foreach ($share in @($shares | Sort-Object Name)) {
        # 種類は UInt32。ハッシュテーブルで引くと、キーの 0 や 1（Int32）と Int64 の値が一致しないので、switch で比べる
        $type = [long]$share.Type
        $typeText = switch ($type) {
            0 { 'ディスク' }
            1 { 'プリンター' }
            2 { 'デバイス' }
            3 { 'IPC' }
            2147483648 { '管理共有' }
            2147483649 { '管理共有（プリンター）' }
            2147483650 { '管理共有（デバイス）' }
            2147483651 { 'IPC' }
            default { [string]$type }
        }
        Add-Row $Rows $share.Name $share.Path $typeText $share.Description
    }
}

#endregion

#region 単位ごとの実行

function Invoke-ParamUnit($Unit) {
    Write-Progress -Activity $ToolTitle -Status ('{0}を読んでいます' -f $Unit.Label)
    $script:currentSkips = New-Object 'System.Collections.Generic.List[object]'
    $failure = $null
    try {
        & $Unit.Reader
    } catch {
        $failure = Get-SkipReason $_.Exception
        $skips.Add([object[]]@($failure[0], $Unit.Label, $failure[1]))
    }
    foreach ($skip in $script:currentSkips) { $skips.Add([object[]]@($skip[0], $Unit.Label, $skip[1])) }
    # 書けなかった項目も、空欄の行として残す（どのサーバーでも同じ項目が並ぶように）
    foreach ($key in $Unit.Owns) {
        if ($paramValues.ContainsKey($key)) { continue }
        $parts = $key.Split('|')
        $note = if ($failure) { '取れませんでした（{0}）' -f $failure[0] } else { '情報がありません' }
        Set-Param $parts[0] $parts[1] '' $note
    }
}

function Invoke-ListUnit([string]$Name) {
    $def = $ListDefs[$Name]
    Write-Progress -Activity $ToolTitle -Status ('{0}を読んでいます' -f $def.Label)
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $script:currentSkips = New-Object 'System.Collections.Generic.List[object]'
    $status = ''
    try {
        & "Read-List$Name" $rows
    } catch {
        $failure = Get-SkipReason $_.Exception
        $status = $failure[0]
        $skips.Add([object[]]@($failure[0], $def.Label, $failure[1]))
        # 途中まで読めた分は、ほかと区別できないので書かない
        $rows.Clear()
    }
    foreach ($skip in $script:currentSkips) { $skips.Add([object[]]@($skip[0], $def.Label, $skip[1])) }
    $script:listResults[$Name] = @{ Rows = $rows; Status = $status }
}

#endregion

#region 画面と summary.txt

# 画面に出す値。取れなかったときは、そのことを書く
function Get-ShownValue([string]$Text) {
    if ($Text) { return $Text }
    return '（取れませんでした）'
}

# 一覧の件数か、取れなかった理由
function Get-ListCountText([string]$Name) {
    $result = $listResults[$Name]
    if ($result.Status) { return $result.Status }
    return Format-Invariant '{0:N0}' $result.Rows.Count
}

function Get-ScreenLines {
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add((Format-Invariant '{0} ({1} {2})' $ToolTitle $ToolName $ScriptVersion))
    $lines.Add((Format-Invariant '{0}   {1:yyyy-MM-dd HH:mm:ss}   管理者: {2}' $machineName $startTime $(if ($isAdmin) { 'はい' } else { 'いいえ' })))

    # コンピューター
    $lines.Add('')
    $lines.Add('■ コンピューター')
    $pairs = New-Object 'System.Collections.Generic.List[object]'
    $hostName = Get-ParamValue 'コンピューター' 'ホスト名'
    $fqdn     = Get-ParamValue 'コンピューター' '完全修飾ドメイン名'
    $hostText = if ($hostName -and $fqdn -and -not $fqdn.Equals($hostName, [StringComparison]::OrdinalIgnoreCase)) { '{0}（{1}）' -f $hostName, $fqdn } else { Get-ShownValue $hostName }
    $pairs.Add([string[]]@('ホスト名', $hostText))
    $domain = Get-ParamValue 'コンピューター' 'ドメイン・ワークグループ'
    $role   = Get-ParamValue 'コンピューター' 'ドメインの役割'
    $domainLabel = if ($script:domainJoined -eq $false) { 'ワークグループ' } else { 'ドメイン' }
    $pairs.Add([string[]]@($domainLabel, $(if ($domain) { '{0}（{1}）' -f $domain, $role } else { Get-ShownValue '' })))
    $maker = Get-ParamValue 'コンピューター' 'メーカー'
    $model = Get-ParamValue 'コンピューター' 'モデル'
    $pairs.Add([string[]]@('メーカー・モデル', (Get-ShownValue ((@($maker, $model) | Where-Object { $_ }) -join '  '))))
    $osName = Get-ParamValue 'OS' '名前'
    $osParts = @()
    $osVersion = Get-ParamValue 'OS' 'バージョン'
    if ($osVersion) { $osParts += $osVersion }
    $osBuild = Get-ParamValue 'OS' 'ビルド'
    if ($osBuild) { $osParts += 'ビルド ' + $osBuild }
    $osBits = Get-ParamValue 'OS' 'アーキテクチャ'
    if ($osBits) { $osParts += $osBits }
    $osText = if ($osName -and $osParts.Count -gt 0) { '{0}（{1}）' -f $osName, ($osParts -join '、') } else { Get-ShownValue $osName }
    $pairs.Add([string[]]@('OS', $osText))
    $pairs.Add([string[]]@('最終起動', (Get-ShownValue (Get-ParamValue 'OS' '最終起動'))))
    $cpuName = Get-ParamValue 'CPU' '名前'
    $cpuText = if ($cpuName) {
        '{0}（{1} ソケット・{2} コア・{3} 論理プロセッサ）' -f $cpuName, (Get-ParamValue 'CPU' 'ソケット数'), (Get-ParamValue 'CPU' 'コア数'), (Get-ParamValue 'CPU' '論理プロセッサ数')
    } else { Get-ShownValue '' }
    $pairs.Add([string[]]@('CPU', $cpuText))
    $installed = Get-ParamValue 'メモリ' '実装メモリ'
    if ($installed -and $memorySlots.Count -gt 0) {
        # 同じ容量のスロットをまとめて「8.0 GB × 2」の形にする（出てきた順）
        $order = New-Object 'System.Collections.Generic.List[long]'
        $counts = @{}
        foreach ($size in $memorySlots) {
            if (-not $counts.ContainsKey($size)) { $counts[$size] = 0; $order.Add($size) }
            $counts[$size]++
        }
        $memoryText = '{0}（{1}）' -f $installed, ((@(foreach ($size in $order) { '{0} × {1}' -f (Format-Gb $size), $counts[$size] })) -join '、')
    } else {
        $memoryText = Get-ShownValue (Get-ParamValue 'メモリ' '使えるメモリ')
    }
    $pairs.Add([string[]]@('メモリ', $memoryText))
    $pairs.Add([string[]]@('タイムゾーン', (Get-ShownValue (Get-ParamValue 'OS' 'タイムゾーン'))))
    $lines.AddRange([string[]](Format-Pairs $pairs))

    # ディスク
    $lines.Add('')
    $lines.Add('■ ディスク')
    $disk = $listResults['Disk']
    if ($disk.Status) {
        $lines.Add(('    （取れませんでした: {0}）' -f $disk.Status))
    } elseif ($disk.Rows.Count -eq 0) {
        $lines.Add('    （ありません）')
    } else {
        $rows = @(foreach ($row in $disk.Rows) { , [string[]]@($row[0], ($row[2] + ' GB'), $row[4], $row[1]) })
        $lines.AddRange([string[]](Format-TextTable ([string[]]@('ディスク', '容量', '形式', 'モデル')) ([bool[]]@($true, $true, $false, $false)) $rows))
    }

    # ボリューム（ドライブ文字かマウント先のあるものだけ）
    $lines.Add('')
    $lines.Add('■ ボリューム')
    $volume = $listResults['Volume']
    $volumeRows = @(foreach ($row in $volume.Rows) {
        if ($row[3]) { , [string[]]@($row[3], $row[4], $row[5], $(if ($row[6]) { $row[6] + ' GB' } else { '' }), $(if ($row[8]) { $row[8] + ' GB' } else { '' }), $row[0]) }
    })
    if ($volume.Status) {
        $lines.Add(('    （取れませんでした: {0}）' -f $volume.Status))
    } elseif ($volumeRows.Count -eq 0) {
        $lines.Add('    （ありません）')
    } else {
        $lines.AddRange([string[]](Format-TextTable ([string[]]@('ドライブ', 'ラベル', 'ファイルシステム', '容量', '空き', 'ディスク')) ([bool[]]@($false, $false, $false, $true, $true, $true)) $volumeRows))
    }

    # ネットワーク
    $lines.Add('')
    $lines.Add('■ ネットワーク')
    $network = $listResults['Network']
    if ($network.Status) {
        $lines.Add(('    （取れませんでした: {0}）' -f $network.Status))
    } elseif ($networkScreen.Count -eq 0) {
        $lines.Add('    （ありません）')
    } else {
        $lines.AddRange([string[]](Format-TextTable ([string[]]@('接続の名前', '状態', 'IPv4 アドレス', 'ゲートウェイ', 'DNS サーバー', 'DHCP')) ([bool[]]@($false, $false, $false, $false, $false, $false)) $networkScreen.ToArray()))
    }

    # 設定
    $lines.Add('')
    $lines.Add('■ 設定')
    $pairs = New-Object 'System.Collections.Generic.List[object]'
    $rdp = Get-ParamValue 'リモートデスクトップ' '接続'
    $rdpText = if ($rdp -eq '許可') {
        $port = Get-ParamValue 'リモートデスクトップ' 'ポート'
        '許可（ネットワークレベル認証: {0}{1}）' -f (Get-ParamValue 'リモートデスクトップ' 'ネットワークレベル認証'), $(if ($port) { '、ポート ' + $port } else { '' })
    } else { Get-ShownValue $rdp }
    $pairs.Add([string[]]@('リモートデスクトップ', $rdpText))
    $firewallParts = @(foreach ($name in $Categories['ファイアウォール']) { '{0}: {1}' -f $name, (Get-ShownValue (Get-ParamValue 'ファイアウォール' $name)) })
    $firewallText = $firewallParts -join '  '
    # 画面の幅（120 桁）に収まるように、画面では短く書く。CSV の備考には「この PC の設定。」も書く
    if (-not $isAdmin) { $firewallText += '（グループポリシーの分を含みません）' }
    $pairs.Add([string[]]@('ファイアウォール', $firewallText))
    $timeType = Get-ParamValue '時刻同期' '種類'
    $timeServer = Get-ParamValue '時刻同期' 'NTP サーバー'
    $timeText = if (($timeType -eq 'NTP' -or $timeType -like 'NTP とドメイン*') -and $timeServer) { '{0}（{1}）' -f $timeType, $timeServer } else { Get-ShownValue $timeType }
    $pairs.Add([string[]]@('時刻同期', $timeText))
    $pairs.Add([string[]]@('Windows Update', ('WSUS: {0}  自動更新: {1}' -f (Get-ShownValue (Get-ParamValue 'Windows Update' 'WSUS サーバー')), (Get-ShownValue (Get-ParamValue 'Windows Update' '自動更新')))))
    $pairs.Add([string[]]@('電源プラン', (Get-ShownValue (Get-ParamValue '電源' '電源プラン'))))
    $pairs.Add([string[]]@('ライセンス認証', (Get-ShownValue (Get-ParamValue 'OS' 'ライセンス認証'))))
    $lines.AddRange([string[]](Format-Pairs $pairs))

    # 一覧の件数
    $lines.Add('')
    $lines.Add('■ 一覧の件数')
    $countItems = @('Feature', 'Software', 'HotFix', 'Service', 'LocalUser', 'Share')
    $countWidth = 0
    foreach ($name in $countItems) { $countWidth = [math]::Max($countWidth, (Get-DisplayWidth (Get-ListCountText $name))) }
    $pairs = New-Object 'System.Collections.Generic.List[object]'
    foreach ($name in $countItems) {
        $text = Format-Cell (Get-ListCountText $name) $countWidth $true
        if ($name -eq 'HotFix' -and -not $listResults[$name].Status -and $latestHotFix) { $text += '（最新 {0}）' -f $latestHotFix }
        $pairs.Add([string[]]@($ListDefs[$name].Label, $text))
    }
    $lines.AddRange([string[]](Format-Pairs $pairs))
    return , $lines
}

#endregion

#region 引数の確認

$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

if ($NoSave -and $OutputFolder) { Stop-WithError '-NoSave のときは、-OutputFolder は指定できません' }

$outputText = $null
if ($OutputFolder) {
    $outputText = Remove-Quote $OutputFolder
    if ($outputText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outputText }
    if ($outputText -match '^(\\\\\?\\)?[A-Za-z]:$') { $outputText += '\' }
    $outputText = Resolve-StartPath $outputText
}

$startTime    = Get-Date
$domainJoined = $null

#endregion

#region 保存先の準備

# 集め終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir = $null
if (-not $NoSave) {
    $baseDir = if ($outputText) { $outputText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        # \\?\ 付きのパスでは末尾のピリオドや空白が残るため、先に通常の書き方で整える（260 文字を超えると例外になるので、そのときはそのまま）
        try { $baseDir = [IO.Path]::GetFullPath($baseDir) } catch { }
        $runName = Format-Invariant 'ServerInventory_{0}_{1:yyyyMMdd_HHmmss}' $machineName $startTime
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
$unexpected = $false
$saveError  = $null
$sw         = [Diagnostics.Stopwatch]::StartNew()
$screen     = $null
$tail       = New-Object 'System.Collections.Generic.List[string]'
try {
    #region 集める

    foreach ($unit in $ParamUnits) { Invoke-ParamUnit $unit }
    foreach ($name in $ListDefs.Keys) { Invoke-ListUnit $name }

    $screen = Get-ScreenLines
    $failedCount = @($skips | Where-Object { $_[0] -ne '対象外' }).Count
    if ($failedCount -gt 0) { $tail.Add((Format-Invariant '※ 取れなかった項目が {0:N0} 件あります（skipped.csv を確認してください）' $failedCount)) }
    if (-not $isAdmin -and @($skips | Where-Object { $_[0] -eq '権限なし' }).Count -gt 0) { $tail.Add('※ 管理者として実行すると、取れる情報が増えます') }
    $tail.Add(('所要 {0}' -f (Format-Duration $sw.Elapsed)))

    #endregion

    #region 保存

    if ($runDir) {
        Write-Progress -Activity $ToolTitle -Status '結果を保存しています'
        try {
            $summary = New-Object 'System.Collections.Generic.List[string]'
            $summary.AddRange([string[]]$screen)
            $summary.Add('')
            $summary.AddRange([string[]]$tail.ToArray())
            $summary.Add(('保存先: {0}' -f $runDir))
            [IO.File]::WriteAllLines((ConvertTo-LongPath (Join-Path $runDir 'summary.txt')), [string[]]$summary, $Utf8Bom)

            $paramRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($category in $Categories.Keys) {
                foreach ($name in @($Categories[$category]) + @($paramExtra[$category].ToArray())) {
                    $entry = $paramValues[$category + '|' + $name]
                    $paramRows.Add([object[]]@($category, (Protect-ExcelText $name), (Protect-ExcelText $entry.Value), $entry.Note))
                }
            }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'parameters.csv')) $ParamColumns $paramRows

            foreach ($name in $ListDefs.Keys) {
                $def = $ListDefs[$name]
                $csvRows = New-Object 'System.Collections.Generic.List[object]'
                foreach ($row in $listResults[$name].Rows) {
                    $csvRows.Add([object[]]@(for ($i = 0; $i -lt $row.Count; $i++) { if ($def.Numeric -contains $i) { $row[$i] } else { Protect-ExcelText $row[$i] } }))
                }
                Save-Csv (ConvertTo-LongPath (Join-Path $runDir $def.File)) $def.Columns $csvRows
            }

            $skipRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($skip in $skips) { $skipRows.Add([object[]]@($skip[0], $skip[1], (Protect-ExcelText $skip[2]))) }
            Save-Csv (ConvertTo-LongPath (Join-Path $runDir 'skipped.csv')) $SkipColumns $skipRows
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
    if (-not $finished -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete((ConvertTo-LongPath $runDir), $true) } catch { } }
        $runDir = $null
        if ($unexpected) { Write-Host 'エラーで止まりました。結果は保存していません。' -ForegroundColor Yellow }
        elseif (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

#region 画面表示

foreach ($line in $screen) {
    if ($line.StartsWith('■', [StringComparison]::Ordinal)) { Write-Host $line -ForegroundColor Cyan } else { Write-Host $line }
}
Write-Host ''
# 保存しなかったときは skipped.csv がないので、理由を画面に出す
foreach ($line in $tail) { Write-Host $(if ($runDir) { $line } else { $line.Replace('skipped.csv を確認してください', '下の一覧を確認してください') }) }
if (-not $runDir -and $skips.Count -gt 0) {
    Write-Host ''
    Write-Host '■ 取れなかったもの' -ForegroundColor Cyan
    $skipScreen = @(foreach ($skip in $skips) { , [string[]]@($skip[0], $skip[1], $skip[2]) })
    foreach ($line in (Format-TextTable ([string[]]@('理由', '項目', '詳細')) ([bool[]]@($false, $false, $false)) $skipScreen)) { Write-Host $line }
}

$exitCode = 0
if ($saveError) {
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
    $exitCode = 1
} elseif ($runDir) {
    Write-Host ('保存先: {0}' -f $runDir)
}

#endregion

exit $exitCode
