<#
.SYNOPSIS
    Test-Connectivity の回帰テストです。

.DESCRIPTION
    このテストの中で、ループバック（127.0.0.1 と ::1）に TCP の待ち受けを作り、Test-Connectivity を実行して期待値と照合します。
    「接続できた」は待ち受けているポート、「拒否された」は待ち受けていないポート、「名前解決失敗」は .invalid の名前
    （RFC 6761 で、どこにも存在しないことが決まっている名前）で再現します。外部のサーバーには接続しません。
    「タイムアウト」は、どこにもつながらない文書用のアドレス 192.0.2.1（RFC 5737）で再現します。
    応答が返ってくるネットワークでは、タイムアウトを使う項目を飛ばします。

    一時フォルダに、対象の CSV（日本語のパス、260 文字を超えるパス、文字コードや書き方の違うもの）と設定ファイルを作ります。
    「権限なし」を再現するため、テストが作ったファイルとフォルダにだけアクセス拒否を設定します（終了時に戻します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Test-Connectivity を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Test-Connectivity\Test-Connectivity.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Test-Connectivity_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool     = Join-Path $workRoot 'tool\Test-Connectivity.ps1'
$lists    = Join-Path $workRoot 'lists'
$utf8Bom  = New-Object Text.UTF8Encoding $true
$csvHead  = '"No","判定","ホスト","IP アドレス","ping","応答時間(ms)","ポート","TCP","接続時間(ms)","メモ","確認日時","詳細"'
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
function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out  = Join-Path $workRoot "out\$Label"
    $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = $log; Dir = (Get-RunDir $out) }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行
function Invoke-Plain([string[]]$ToolArgs) {
    $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 設定ファイルの OutputFolder や相対パスの基準を確かめるため、-OutputFolder を足さず、指定したフォルダから起動する
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
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
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

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'Connectivity_*' -ErrorAction SilentlyContinue | Sort-Object Name)
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
    $summary = ''
    if ($Dir -and (Test-Path -LiteralPath (Join-Path $Dir 'summary.txt'))) { $summary = [IO.File]::ReadAllText((Join-Path $Dir 'summary.txt')) }
    [pscustomobject]@{
        Results = @(Read-CsvOrEmpty 'results.csv')
        Changes = @(Read-CsvOrEmpty 'changes.csv')
        Summary = $summary
        Lines   = @($summary -split '\r?\n')
    }
}

# results.csv の、No が一致する行。ないとき（ツールが保存に失敗したときなど）も止まらないように、空の行を返す
function Get-Row($Result, [int]$No) {
    $found = @($Result.Results | Where-Object { $_.'No' -eq "$No" }) | Select-Object -First 1
    if ($found) { return $found }
    [pscustomobject]@{ 'No' = ''; '判定' = '(なし)'; 'ホスト' = ''; 'IP アドレス' = ''; 'ping' = ''; '応答時間(ms)' = ''; 'ポート' = ''; 'TCP' = ''; '接続時間(ms)' = ''; 'メモ' = ''; '確認日時' = ''; '詳細' = '' }
}

# 判定/IP アドレス/ping/ポート/TCP
function Get-RowText($Result, [int]$No) {
    $row = Get-Row $Result $No
    '{0}/{1}/{2}/{3}/{4}' -f $row.'判定', $row.'IP アドレス', $row.'ping', $row.'ポート', $row.'TCP'
}

function Join-Sorted($Items) {
    $array = [string[]]@($Items | Where-Object { $null -ne $_ })
    [Array]::Sort($array, [StringComparer]::Ordinal)
    $array -join ','
}

function Get-ResultFileNames([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { $_.Name })
}

function Get-BomFlags([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir -File | ForEach-Object {
        $bytes = [IO.File]::ReadAllBytes($_.FullName)
        '{0}={1}' -f $_.Name, ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    })
}

function Get-FirstLine([string]$Dir, [string]$FileName) {
    if (-not $Dir -or -not (Test-Path -LiteralPath (Join-Path $Dir $FileName))) { return '(なし)' }
    @([IO.File]::ReadAllLines((Join-Path $Dir $FileName)))[0]
}

# 画面の「対象」の行（例: C:\x\targets.csv（2 件）、-Target（1 件）   タイムアウト 2 秒）
function Get-TargetText([string]$Log) { [regex]::Match($Log, '対象: (.+?秒)').Groups[1].Value }

function Get-CountText([string]$Log) { [regex]::Match($Log, '結果: OK \S+ 件  NG \S+ 件').Value }

# 260 文字を超えるパスにも書けるように、\\?\ を付けて書く
function New-TextFile([string]$FilePath, [string[]]$Lines, [Text.Encoding]$Encoding = $utf8Bom, [string]$NewLine = "`r`n") {
    [void][IO.Directory]::CreateDirectory('\\?\' + $FilePath.Substring(0, $FilePath.LastIndexOf('\')))
    [IO.File]::WriteAllText('\\?\' + $FilePath, ($Lines -join $NewLine) + $NewLine, $Encoding)
    $FilePath
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

# 待ち受けていないポート（待ち受けを作ってすぐに止める）
function Get-ClosedPort([Net.IPAddress]$Address) {
    $probe = New-Object Net.Sockets.TcpListener ($Address, 0)
    $probe.Start()
    $port = $probe.LocalEndpoint.Port
    $probe.Stop()
    $port
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$userSid    = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied     = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したパス（終了時に戻す）
$listeners  = New-Object 'System.Collections.Generic.List[object]'   # 作った待ち受け（終了時に止める）
$acceptLoop = $null

try {
    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool

    # 接続を受け付けて、すぐに閉じる待ち受け。受け付けないと待ち行列があふれて、後の実行で拒否されるため、別のスレッドで受け付け続ける
    $listener4 = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, 0)
    $listener4.Start()
    $listeners.Add($listener4)
    $port4   = $listener4.LocalEndpoint.Port
    $closed4 = Get-ClosedPort ([Net.IPAddress]::Loopback)
    $port6   = $null
    $closed6 = $null
    try {
        $listener6 = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::IPv6Loopback, 0)
        $listener6.Start()
        $listeners.Add($listener6)
        $port6   = $listener6.LocalEndpoint.Port
        $closed6 = Get-ClosedPort ([Net.IPAddress]::IPv6Loopback)
    } catch { }
    $acceptLoop = [powershell]::Create()
    [void]$acceptLoop.AddScript({
        param($Listeners)
        while ($true) {
            foreach ($l in $Listeners) { try { while ($l.Pending()) { $l.AcceptTcpClient().Close() } } catch { } }
            Start-Sleep -Milliseconds 20
        }
    }.ToString()).AddArgument($listeners.ToArray())
    $null = $acceptLoop.BeginInvoke()
    Assert-Equal '待ち受けの準備' ('{0}/{1}' -f ($port4 -gt 0), ($closed4 -ne $port4)) 'True/True'

    # 応答のないアドレス（TCP の接続が終わらない）を使えるか。応答が返るネットワークでは、タイムアウトの項目を飛ばす
    $blackHole = $null
    try {
        $probeClient = New-Object Net.Sockets.TcpClient
        $probeTask   = $probeClient.ConnectAsync('192.0.2.1', 443)
        if (-not $probeTask.Wait(1500)) { $blackHole = '192.0.2.1:443' }
        $probeClient.Close()
    } catch { }

    Write-Section '判定（-Target）'
    $run = Invoke-Tool 'basic' @('-Target', "localhost,127.0.0.1:$port4,127.0.0.1:$closed4,wbtest-no-such-host.invalid:443")
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード（NG がある）' $run.Code 2
    Assert-Equal '保存するファイル' (Get-ResultFileNames $run.Dir) 'results.csv,summary.txt'
    Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'results.csv=True,summary.txt=True'
    Assert-Equal 'results.csv の見出し' (Get-FirstLine $run.Dir 'results.csv') $csvHead
    Assert-Equal 'ポートなし: ping で判定・IPv4 を優先' (Get-RowText $x 1) 'OK/127.0.0.1/応答あり//-'
    Assert-Equal 'ポートなし: 応答時間は数値' ((Get-Row $x 1).'応答時間(ms)' -match '^\d[\d,]*$') 'True'
    Assert-Equal '待ち受けているポート: 接続できた' (Get-RowText $x 2) "OK/127.0.0.1/応答あり/$port4/接続できた"
    Assert-Equal '待ち受けているポート: 接続時間は数値' ((Get-Row $x 2).'接続時間(ms)' -match '^\d[\d,]*$') 'True'
    Assert-Equal '待ち受けていないポート: 拒否された（ping の結果は問わず NG）' (Get-RowText $x 3) "NG/127.0.0.1/応答あり/$closed4/拒否された"
    Assert-Equal '待ち受けていないポート: 接続時間は空' (Get-Row $x 3).'接続時間(ms)' ''
    Assert-Equal '名前解決失敗: ping と TCP は確かめない' (Get-RowText $x 4) 'NG//-/443/-'
    Assert-Equal '名前解決失敗: 詳細に理由' ((Get-Row $x 4).'詳細'.StartsWith('名前解決: ')) 'True'
    Assert-Equal 'ホストは ="..." の形' ('{0}|{1}' -f (Get-Row $x 1).'ホスト', (Get-Row $x 4).'ホスト') '="localhost"|="wbtest-no-such-host.invalid"'
    Assert-Equal '確認日時の形' ((Get-Row $x 1).'確認日時' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
    Assert-Equal 'summary の 1 行目' $x.Lines[0] '疎通一括確認 (Test-Connectivity 1.0.0)'
    Assert-Equal 'summary の 2 行目（実行元）' $x.Lines[1].StartsWith(('実行元: {0}（' -f [Environment]::MachineName)) 'True'
    Assert-Equal '対象の行' (Get-TargetText $run.Log) '-Target（4 件）   タイムアウト 2 秒'
    Assert-Equal '結果の行' (Get-CountText $run.Log) '結果: OK 2 件  NG 2 件'
    Assert-Equal 'summary に表' ($x.Summary.Contains(' No  判定  ホスト') -and $x.Summary.Contains("  3  NG    127.0.0.1")) 'True'
    Assert-Equal '画面に保存先' $run.Log.Contains(('保存先: {0}' -f $run.Dir)) 'True'

    $run = Invoke-Plain @('-Target', "127.0.0.1:$port4,localhost", '-NoSave')
    Assert-Equal 'すべて OK: 終了コード 0' ('{0}/{1}' -f $run.Code, (Get-CountText $run.Log)) '0/結果: OK 2 件  NG 0 件'

    # Windows は拒否されても SYN を送り直すため、そのままでは拒否と分かるまで 2 秒ほどかかり、短い待ち時間ではタイムアウトになっていた
    $run = Invoke-Tool 'refused-1s' @('-Target', "127.0.0.1:$closed4", '-Timeout', '1')
    Assert-Equal '-Timeout 1 でも拒否されたと分かる' (Get-RowText (Read-Result $run.Dir) 1) "NG/127.0.0.1/応答あり/$closed4/拒否された"
    foreach ($wait in '5', '20') {
        $run = Invoke-Tool "refused-$wait" @('-Target', "127.0.0.1:$closed4", '-Timeout', $wait)
        $elapsed = [double]([regex]::Match($run.Log, '所要 ([\d.]+)秒').Groups[1].Value)
        Assert-Equal ('-Timeout {0}: 拒否されたと分かり、待ち時間まで待たない' -f $wait) ('{0}/{1}' -f (Get-Row (Read-Result $run.Dir) 1).'TCP', ($elapsed -gt 0 -and $elapsed -lt 4)) '拒否された/True'
    }
    if ($blackHole) {
        $run = Invoke-Tool 'timeout' @('-Target', $blackHole, '-Timeout', '1')
        $x = Read-Result $run.Dir
        Assert-Equal '応答のないアドレス: タイムアウト' ('{0}/{1}' -f $run.Code, (Get-RowText $x 1)) '2/NG/192.0.2.1/応答なし/443/タイムアウト'
        $run = Invoke-Tool 'timeout-4s' @('-Target', $blackHole, '-Timeout', '4')
        $elapsed = [double]([regex]::Match($run.Log, '所要 ([\d.]+)秒').Groups[1].Value)
        Assert-Equal '-Timeout 4: 待ち時間まで待ってからタイムアウト' ('{0}/{1}' -f (Get-Row (Read-Result $run.Dir) 1).'TCP', ($elapsed -ge 3.9 -and $elapsed -lt 6)) 'タイムアウト/True'
    } else {
        Write-Skip 'タイムアウト' '192.0.2.1 から応答が返るネットワークのため'
    }

    Write-Section 'IP アドレスの書き方'
    $run = Invoke-Tool 'address' @('-Target', ('127.000.000.001:{0},127.0.0.010,{1}' -f $port4, [Environment]::MachineName))
    $x = Read-Result $run.Dir
    Assert-Equal '0 で桁をそろえた IPv4' (Get-RowText $x 1) "OK/127.0.0.1/応答あり/$port4/接続できた"
    Assert-Equal '0 を付けた値は 10 進数として読む（8 進数にしない）' (Get-Row $x 2).'IP アドレス' '127.0.0.10'
    Assert-Equal 'コンピューター名: IPv4 で確かめる' ('{0}/{1}' -f (Get-Row $x 3).'判定', ((Get-Row $x 3).'IP アドレス' -match '^\d+\.\d+\.\d+\.\d+$')) 'OK/True'
    if ($port6) {
        $run = Invoke-Tool 'ipv6' @('-Target', "[::1]:$port6,::1,[::1]:$closed6,[::1]")
        $x = Read-Result $run.Dir
        Assert-Equal 'IPv6: [::1]:ポート' (Get-RowText $x 1) "OK/::1/応答あり/$port6/接続できた"
        Assert-Equal 'IPv6: ::1（ping だけ）' (Get-RowText $x 2) 'OK/::1/応答あり//-'
        Assert-Equal 'IPv6: 待ち受けていないポート' (Get-RowText $x 3) "NG/::1/応答あり/$closed6/拒否された"
        Assert-Equal 'IPv6: [::1]（ポートなし）' ('{0}/{1}' -f (Get-RowText $x 4), (Get-Row $x 4).'ホスト') 'OK/::1/応答あり//-/="::1"'
    } else {
        Write-Skip 'IPv6' '::1 で待ち受けられなかったため'
    }

    Write-Section '対象の CSV（-List）'
    $longMemo = 'あ' * 300
    $listFile = New-TextFile "$lists\list.csv" @(
        'ホスト,ポート,メモ,担当'
        '# コメントの行,,,'
        "127.0.0.1,$port4,Web サーバー（HTTPS）,田中"
        ',,ホストが空の行,'
        "localhost,,`"セル内`n改行`","
        "127.0.0.1,$closed4,`"カンマ, 入り`","
        '127.0.0.1,,19" ラック,'
        "127.0.0.1, $closed4 ,001,"
        '127.0.0.1,,"""引用""",'
        "127.0.0.1,,$longMemo,"
        '  # 空白のあとの #,,,'
    )
    $run = Invoke-Tool 'list' @('-List', $listFile, '-Target', 'localhost')
    $x = Read-Result $run.Dir
    Assert-Equal '終了コード' $run.Code 2
    Assert-Equal '対象の行（CSV の次に -Target）' (Get-TargetText $run.Log) "$listFile（7 件）、-Target（1 件）   タイムアウト 2 秒"
    Assert-Equal '番号は CSV、-Target の順' ('{0}|{1}' -f (Get-Row $x 1).'ホスト', (Get-Row $x 8).'ホスト') '="127.0.0.1"|="localhost"'
    Assert-Equal '判定' (($x.Results | ForEach-Object { $_.'判定' }) -join ',') 'OK,OK,NG,OK,NG,OK,OK,OK'
    Assert-Equal 'ポートの前後の空白は無視する' (Get-RowText $x 5) "NG/127.0.0.1/応答あり/$closed4/拒否された"
    Assert-Equal 'メモ' ((1..6 | ForEach-Object { (Get-Row $x $_).'メモ' }) -join '|') '="Web サーバー（HTTPS）"|="セル内 改行"|="カンマ, 入り"|="19"" ラック"|="001"|="""引用"""'
    Assert-Equal '255 文字を超えるメモは ="..." にしない' ((Get-Row $x 7).'メモ' -ceq $longMemo) 'True'
    Assert-Equal '-Target の対象のメモは空' (Get-Row $x 8).'メモ' ''

    $english = New-TextFile "$lists\english.csv" @('Memo,Host,Port', 'メモ1,localhost,', "メモ2,127.0.0.1,$port4")
    $run = Invoke-Tool 'english' @('-List', $english)
    $x = Read-Result $run.Dir
    Assert-Equal '英語の見出し・列の順番が違う' ('{0}/{1}/{2}' -f $run.Code, (Get-RowText $x 2), (Get-Row $x 1).'メモ') "0/OK/127.0.0.1/応答あり/$port4/接続できた/=""メモ1"""
    $excel = New-TextFile "$lists\excel.csv" @('ホスト,ポート,メモ,,', 'localhost,,a,,', ',,,,')
    $run = Invoke-Plain @('-List', $excel, '-NoSave')
    Assert-Equal 'Excel が付ける空の見出しと空の行' ('{0}/{1}' -f $run.Code, (Get-TargetText $run.Log)) "0/$excel（1 件）   タイムアウト 2 秒"
    $lfOnly = New-TextFile "$lists\lf.csv" @('ホスト,ポート,メモ', 'localhost,,LF だけ', "127.0.0.1,$port4,") $utf8Bom "`n"
    $run = Invoke-Plain @('-List', $lfOnly, '-NoSave')
    Assert-Equal '改行が LF だけの CSV' ('{0}/{1}' -f $run.Code, (Get-TargetText $run.Log)) "0/$lfOnly（2 件）   タイムアウト 2 秒"

    # Windows 11 のメモ帳は BOM なしの UTF-8、Excel の「CSV（コンマ区切り）」は Shift_JIS で保存する
    $encodings = [ordered]@{
        'UTF-8（BOM なし）' = New-Object Text.UTF8Encoding $false
        'UTF-8（BOM 付き）' = $utf8Bom
        'Shift_JIS'         = [Text.Encoding]::GetEncoding(932)
        'UTF-16 LE'         = [Text.Encoding]::Unicode
        'UTF-16 BE'         = [Text.Encoding]::BigEndianUnicode
    }
    $i = 0
    foreach ($label in $encodings.Keys) {
        $file = New-TextFile "$lists\enc$i.csv" @('ホスト,ポート,メモ', "127.0.0.1,$port4,日本語のメモ$i") $encodings[$label]
        $run  = Invoke-Plain @('-List', $file, '-NoSave')
        Assert-Equal ('文字コード {0}' -f $label) ('{0}/{1}/{2}' -f $run.Code, (Get-TargetText $run.Log), $run.Log.Contains("接続できた  日本語のメモ$i")) "0/$file（1 件）   タイムアウト 2 秒/True"
        $i++
    }

    Write-Section '対象の CSV: 書き間違い（確かめ始める前に止まる）'
    $listMistakes = @(
        @('ポートが範囲外（何行目かはセル内改行を数える）', @('ホスト,ポート,メモ', 'localhost,,"1 行目', '2 行目"', '127.0.0.1,80800,ポート誤り'), '対象の CSV の 4 行目: ポートは 1～65535 の数値で指定してください: 80800'),
        @('ポートが 65536', @('ホスト,ポート', 'localhost,65536'), '対象の CSV の 2 行目: ポートは 1～65535 の数値で指定してください: 65536'),
        @('ポートが 0', @('ホスト,ポート', 'localhost,0'), '対象の CSV の 2 行目: ポートは 1～65535 の数値で指定してください: 0'),
        @('ポートに + が付いている', @('ホスト,ポート', 'localhost,+80'), 'ポートは 1～65535 の数値で指定してください: +80'),
        @('ポートが小数', @('ホスト,ポート', 'localhost,80.0'), 'ポートは 1～65535 の数値で指定してください: 80.0'),
        @('ポートが全角の数字', @('ホスト,ポート', 'localhost,４４３'), 'ポートは 1～65535 の数値で指定してください: ４４３'),
        @('ホストの列がない', @('Name,Port', 'localhost,443'), '対象の CSV に「ホスト」の列がありません。1 行目を見出し（ホスト,ポート,メモ）にしてください'),
        @('引用符が閉じていない', @('ホスト,ポート,メモ', 'localhost,,ok', '127.0.0.1,,"閉じていない', 'localhost,,ここも消える'), '対象の CSV の 3 行目: 引用符（"）が閉じていません')
    )
    for ($i = 0; $i -lt $listMistakes.Count; $i++) {
        $file    = New-TextFile "$lists\err$i.csv" $listMistakes[$i][1]
        $run     = Invoke-Plain @('-List', $file, '-NoSave')
        $missing = @(@($listMistakes[$i][2], "ファイル: $file") | Where-Object { -not $run.Log.Contains($_) })
        # 終了コード / 表示に足りない文 / 確かめ始めたか
        Assert-Equal ('書き間違い: {0}' -f $listMistakes[$i][0]) ('{0}/{1}/{2}' -f $run.Code, ($missing -join ' | '), $run.Log.Contains('疎通一括確認 (')) '1//False'
    }
    $onlyComment = New-TextFile "$lists\comment.csv" @('ホスト,ポート,メモ', '# web01,443,')
    $run = Invoke-Plain @('-List', $onlyComment, '-NoSave')
    Assert-Equal '確かめる対象がない' ('{0}/{1}' -f $run.Code, $run.Log.Contains('確かめる対象がありません')) '1/True'
    [IO.File]::WriteAllBytes("$lists\empty.csv", [byte[]]@())
    $run = Invoke-Plain @('-List', "$lists\empty.csv", '-NoSave')
    Assert-Equal '空のファイル' ('{0}/{1}' -f $run.Code, $run.Log.Contains('確かめる対象がありません')) '1/True'
    $run = Invoke-Plain @('-List', "$lists\nothing.csv", '-NoSave')
    Assert-Equal '見つからない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("対象の CSV が見つかりません: $lists\nothing.csv")) '1/True'
    $null = New-TextFile "$lists\memo.csv.txt" @('ホスト', 'localhost')
    $run = Invoke-Plain @('-List', "$lists\memo.csv", '-NoSave')
    Assert-Equal '拡張子が .txt になっている' ('{0}/{1}' -f $run.Code, $run.Log.Contains("拡張子が .txt になっていませんか: $lists\memo.csv.txt")) '1/True'
    $deniedList = New-TextFile "$lists\denied.csv" @('ホスト', 'localhost')
    if (Add-Deny $deniedList '(R)') {
        $run = Invoke-Plain @('-List', $deniedList, '-NoSave')
        Assert-Equal '読めない CSV' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("対象の CSV を読めません: $deniedList（"), $run.Log.Contains('\\?\')) '1/True/False'
    } else {
        Write-Skip '読めない CSV' 'アクセス拒否を設定できなかったため'
    }

    Write-Section '-Target の書き方と誤り'
    $run = Invoke-Plain @('-Target', "localhost, ,,'127.0.0.1:$port4'", '-NoSave')
    Assert-Equal '空の項目と引用符を飛ばす' ('{0}/{1}' -f $run.Code, (Get-TargetText $run.Log)) '0/-Target（2 件）   タイムアウト 2 秒'
    $targetMistakes = @(
        @('ポートが数値でない（何つ目か）', @('-Target', 'localhost,db01:abc'), '-Target の 2 つ目: ポートは 1～65535 の数値で指定してください: db01:abc'),
        @('ポートが空', @('-Target', 'web01:'), '-Target の 1 つ目: ポートは 1～65535 の数値で指定してください: web01:'),
        @('ポートが範囲外', @('-Target', 'web01:65536'), 'ポートは 1～65535 の数値で指定してください: web01:65536'),
        @('IPv6 のポートが空', @('-Target', '[::1]:'), 'ポートは 1～65535 の数値で指定してください: [::1]:'),
        @('ホストがない', @('-Target', ':443'), '-Target の 1 つ目: ホストがありません: :443'),
        @('対象が空', @('-Target', ' , '), '確かめる対象がありません'),
        @('対象を指定しない', @(), '確かめる対象を、-List（CSV ファイル）か -Target で指定してください。ひな形は -NewList で作れます。')
    )
    foreach ($case in $targetMistakes) {
        $run = Invoke-Plain (@($case[1]) + '-NoSave')
        Assert-Equal ('誤り: {0}' -f $case[0]) ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains($case[2]), $run.Log.Contains('疎通一括確認 (')) '1/True/False'
    }
    foreach ($case in @(@('-Timeout', '0'), @('-Timeout', '61'), @('-Repeat', '0'), @('-Repeat', '3601'))) {
        $run = Invoke-Plain @('-Target', 'localhost', $case[0], $case[1], '-NoSave')
        Assert-Equal ('{0} {1} は止まる' -f $case[0], $case[1]) $run.Code 1
    }

    Write-Section '表の表示（全角は 2 文字分の幅）'
    # 名前解決できない対象だけにして、ping の応答時間のように毎回変わる値を表に出さない
    $tableList = New-TextFile "$lists\table.csv" @('ホスト,ポート,メモ', 'a.invalid,443,Web サーバー（HTTPS）', 'サーバー.invalid,,𠮷野家ｻｰﾊﾞｰ')
    $run = Invoke-Tool 'table' @('-List', $tableList)
    $x = Read-Result $run.Dir
    $head = [array]::IndexOf($x.Lines, @($x.Lines | Where-Object { $_.StartsWith(' No') })[0])
    $expectedTable = @(
        (' No  判定  ホスト' + (' ' * 12) + 'IP アドレス   ping  ポート  TCP  メモ'),
        ('  1  NG    a.invalid' + (' ' * 9) + '名前解決失敗  -' + (' ' * 8) + '443  -    Web サーバー（HTTPS）'),
        ('  2  NG    サーバー.invalid  名前解決失敗  -' + (' ' * 13) + '-    𠮷野家ｻｰﾊﾞｰ')
    )
    $actualTable = if ($head -ge 0) { @($x.Lines[$head..($head + 2)]) } else { @() }
    for ($i = 0; $i -lt 3; $i++) {
        $actualLine = if ($i -lt $actualTable.Count) { $actualTable[$i] } else { '(なし)' }
        Assert-Equal ('表の {0} 行目' -f ($i + 1)) $actualLine $expectedTable[$i]
    }
    Assert-Equal '画面にも同じ表' $run.Log.Contains($expectedTable[1]) 'True'

    Write-Section 'ひな形（-NewList）'
    $newList = Join-Path $workRoot 'new\対象 一覧.csv'
    [void][IO.Directory]::CreateDirectory((Split-Path $newList))
    $run = Invoke-Plain @('-NewList', $newList)
    Assert-Equal '作る: 終了コードと表示' ('{0}/{1}' -f $run.Code, $run.Log.Contains("ひな形を作りました: $newList")) '0/True'
    $templateBytes = [IO.File]::ReadAllBytes($newList)
    Assert-Equal 'ひな形は BOM 付き UTF-8' ($templateBytes[0] -eq 0xEF -and $templateBytes[1] -eq 0xBB -and $templateBytes[2] -eq 0xBF) 'True'
    Assert-Equal 'ひな形の中身' ([IO.File]::ReadAllLines($newList) -join '|') 'ホスト,ポート,メモ|# 先頭が # の行は確かめません。ポートを空欄にすると ping だけを確かめます,,|web01,443,Web サーバー（HTTPS）|db01,1433,SQL Server|192.168.1.1,,ゲートウェイ（ping だけ）'
    [IO.File]::AppendAllText($newList, "localhost,,書き足した行`r`n", (New-Object Text.UTF8Encoding $false))
    $sizeBefore = (Get-Item -LiteralPath $newList).Length
    $run = Invoke-Plain @('-NewList', $newList)
    Assert-Equal '同じ名前のファイルがある: 上書きしない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("同じ名前のファイルがあるため、ひな形を作りません（上書きしません）: $newList"), ((Get-Item -LiteralPath $newList).Length -eq $sizeBefore)) '1/True/True'
    $run = Invoke-Plain @('-NewList', (Split-Path $newList))
    Assert-Equal '同じ名前のフォルダがある' ('{0}/{1}' -f $run.Code, $run.Log.Contains('同じ名前のファイルがあるため')) '1/True'
    $run = Invoke-Plain @('-NewList', (Join-Path $workRoot 'no\such\new.csv'))
    Assert-Equal '作れない場所' ('{0}/{1}' -f $run.Code, $run.Log.Contains(('ひな形を作れません: {0}（' -f (Join-Path $workRoot 'no\such\new.csv')))) '1/True'
    # ひな形をそのまま読めること（web01・db01 は名前解決に失敗してよい。待ち時間は短くする）
    $run = Invoke-Plain @('-List', $newList, '-Timeout', '1', '-NoSave')
    Assert-Equal 'ひな形を -List で読める' (Get-TargetText $run.Log) "$newList（4 件）   タイムアウト 1 秒"

    Write-Section '保存'
    $noSaveOut = Join-Path $workRoot 'out\nosave'
    $run = Invoke-Plain @('-Target', 'localhost', '-NoSave', '-OutputFolder', $noSaveOut)
    Assert-Equal '-NoSave: 保存しない' ('{0}/{1}/{2}' -f $run.Code, (Test-Path -LiteralPath $noSaveOut), $run.Log.Contains('保存先')) '0/False/False'
    $jpOut = Join-Path $workRoot 'out\保存 先'
    $run = Invoke-Plain @('-Target', 'localhost', '-OutputFolder', $jpOut)
    Assert-Equal '日本語とスペースを含む保存先' ('{0}/{1}' -f $run.Code, (Get-ResultFileNames (Get-RunDir $jpOut))) '0/results.csv,summary.txt'
    # 通常の書き方では末尾のピリオドは無視される。\\?\ 付きで作ると、ピリオドの付いたフォルダができてしまう
    $run = Invoke-Plain @('-Target', 'localhost', '-OutputFolder', (Join-Path $workRoot 'out\trail.'))
    Assert-Equal '末尾にピリオドが付いた保存先' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultFileNames (Get-RunDir (Join-Path $workRoot 'out\trail'))), [IO.Directory]::Exists('\\?\' + (Join-Path $workRoot 'out\trail.'))) '0/results.csv,summary.txt/False'
    $freeLetter = Get-FreeLetter
    if ($freeLetter) {
        $run = Invoke-Plain @('-Target', 'localhost', '-OutputFolder', "${freeLetter}:\nope")
        Assert-Equal '書き込めない保存先: 警告して画面表示だけ' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("保存先に書き込めないため、結果は保存しません: ${freeLetter}:\nope"), $run.Log.Contains('  1  OK    localhost')) '0/True/True'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }
    # 保存先のフォルダは作れるが、その中にファイルを作れないようにして、保存に失敗させる
    $guardOut = Join-Path $workRoot 'out\save-fail'
    [void][IO.Directory]::CreateDirectory($guardOut)
    if (Add-Deny $guardOut '(OI)(CI)(WD)') {
        $run = Invoke-Plain @('-Target', 'localhost', '-OutputFolder', $guardOut)
        Assert-Equal '保存に失敗: 警告する' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした'), $run.Log.Contains('保存先:')) '0/True/False'
        Assert-Equal '保存に失敗: 書きかけの保存先を残さない' @(Get-ChildItem -LiteralPath $guardOut -Force).Count 0
    } else {
        Write-Skip '保存に失敗したとき' 'アクセス拒否を設定できなかったため'
    }
    # 保存先を作れないときの警告に、改行が混ざらないこと。画面の文字は折り返されるため、ストリームから受け取る
    $warnScript = New-HelperScript 'collect-warnings.ps1' @(
        'param($Tool, $OutputFolder)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Target'', ''localhost'').AddParameter(''OutputFolder'', $OutputFolder)'
        '[void]$ps.Invoke()'
        '$lines = @($ps.Streams.Warning | ForEach-Object { $_.Message.Replace("`r", ''<CR>'').Replace("`n", ''<LF>'') })'
        '''COUNT='' + @($lines | Where-Object { $_ -like ''*保存先に書き込めないため*'' }).Count'
        '''BROKEN='' + @($lines | Where-Object { $_.Contains(''<CR>'') -or $_.Contains(''<LF>'') }).Count'
    )
    $result = Invoke-Helper $warnScript @('-Tool', $tool, '-OutputFolder', '\\localhost\workbench_no_such_share\out')
    # 確かめる前と、結果の表のあとの 2 回出す（表が長いと、最初の警告が見えなくなるため）
    Assert-Equal '保存先を作れないときの警告' ('{0}/{1}' -f $result['COUNT'], $result['BROKEN']) '2/0'

    Write-Section 'コマンドプロンプトから起動'
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
    $cmdOut  = Join-Path $workRoot 'out\cmd'
    $run = Invoke-ViaCmd ('{0} -Target "127.0.0.1:{1},localhost" -Timeout 1 -OutputFolder "{2}\"' -f $cmdBase, $port4, $cmdOut)
    Assert-Equal '最後の引数が "パス\" でも保存できる・-Target のカンマ区切り' ('{0}/{1}/{2}' -f $run.Code, (Get-TargetText $run.Log), (Get-ResultFileNames (Get-RunDir $cmdOut))) '0/-Target（2 件）   タイムアウト 1 秒/results.csv,summary.txt'
    $run = Invoke-ViaCmd ('{0} -Target "127.0.0.1:{1}"' -f $cmdBase, $closed4)
    Assert-Equal 'NG があるときの終了コード（%ERRORLEVEL%）' $run.Code 2
    # 後ろの -Target まで保存先の値になるが、「確かめる対象がない」ではなく、引用符のヒントを出す
    $run = Invoke-ViaCmd ('{0} -OutputFolder "{1}\" -Target localhost' -f $cmdBase, $cmdOut)
    Assert-Equal '"パス\" の後ろに引数: ヒントを出して止まる'('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('-OutputFolder の値に引用符'), $run.Log.Contains('最後に \ を付けないでください')) '1/True/True'
    $run = Invoke-ViaCmd ('{0} -NoSave -Target ''localhost,127.0.0.1:{1}''' -f $cmdBase, $port4)
    Assert-Equal 'cmd で ''…,…'' と囲んだ -Target' ('{0}/{1}' -f $run.Code, (Get-TargetText $run.Log)) '0/-Target（2 件）   タイムアウト 2 秒'
    $cmdList = New-TextFile "$workRoot\cmdlist\targets.csv" @('ホスト', 'localhost')
    $run = Invoke-ViaCmd ('{0} -NoSave -List "{1}"' -f $cmdBase, $cmdList)
    Assert-Equal 'cmd から -List' ('{0}/{1}' -f $run.Code, (Get-TargetText $run.Log)) "0/$cmdList（1 件）   タイムアウト 2 秒"

    Write-Section '名前に [ ] を含むフォルダから起動（-List と -NewList を相対パスで指定）'
    # 5.1 は、名前に [ ] を含むフォルダで起動すると、名前をワイルドカードとして扱い、別の場所から始まる
    $start = Join-Path $workRoot 'start'
    $startCases = @(
        @('一致するフォルダがない（$PSHOME から始まる）', 'br[1] x', @()),
        @('隣に 1 つ一致するフォルダがある', 'sib\[a]', @('sib\a')),
        @('[ が閉じていない（ドライブのルートから始まる）', 'ub\a[b', @('ub\ab')),
        @('隣に 2 つ一致する（ドライブのルートから始まる）', 'mm\[ab]', @('mm\a', 'mm\b'))
    )
    foreach ($case in $startCases) {
        $null = New-TextFile "$start\$($case[1])\targets.csv" @('ホスト', 'localhost')
        # 間違ったフォルダの CSV を読んだら件数で分かるように、隣のフォルダには 2 件の CSV を置く
        foreach ($decoy in $case[2]) { $null = New-TextFile "$start\$decoy\targets.csv" @('ホスト', 'localhost', '127.0.0.1') }
    }
    foreach ($case in $startCases) {
        $dir = "$start\$($case[1])"
        $run = Invoke-From $dir @('-List', 'targets.csv', '-NoSave')
        Assert-Equal ('[ ] を含むフォルダから -List: {0}' -f $case[0]) (Get-TargetText $run.Log) "$dir\targets.csv（1 件）   タイムアウト 2 秒"
        $run = Invoke-From $dir @('-NewList', 'new.csv')
        Assert-Equal ('[ ] を含むフォルダから -NewList: {0}' -f $case[0]) ('{0}/{1}' -f $run.Code, [IO.File]::Exists("$dir\new.csv")) '0/True'
    }

    Write-Section '260 文字を超えるパス'
    $longDir  = "$workRoot\" + ('L' * 150) + '\' + ('M' * 150)
    $longList = New-TextFile "$longDir\targets.csv" @('ホスト', 'localhost')
    $run = Invoke-Plain @('-List', $longList, '-NoSave')
    Assert-Equal '-List' ('{0}/{1}' -f $run.Code, (Get-TargetText $run.Log)) "0/$longList（1 件）   タイムアウト 2 秒"
    $run = Invoke-Plain @('-Target', 'localhost', '-OutputFolder', "$longDir\out")
    $longRun = @(Get-ChildItem -LiteralPath ('\\?\' + "$longDir\out") -Directory -ErrorAction SilentlyContinue)
    $longFiles = if ($longRun.Count -gt 0) { Join-Sorted (Get-ChildItem -LiteralPath $longRun[0].FullName -File | ForEach-Object { $_.Name }) } else { '(なし)' }
    Assert-Equal '-OutputFolder' ('{0}/{1}/{2}' -f $run.Code, $longFiles, $run.Log.Contains('保存先に書き込めない')) '0/results.csv,summary.txt/False'
    $run = Invoke-Plain @('-NewList', "$longDir\new.csv")
    Assert-Equal '-NewList' ('{0}/{1}' -f $run.Code, [IO.File]::Exists('\\?\' + "$longDir\new.csv")) '0/True'
    $longIni = New-TextFile "$longDir\check.ini" @('List = targets.csv', 'NoSave = true')
    $run = Invoke-Plain @('-ConfigFile', $longIni)
    Assert-Equal '-ConfigFile（List は設定ファイルのフォルダが基準）' ('{0}/{1}' -f $run.Code, (Get-TargetText $run.Log)) "0/$longList（1 件）   タイムアウト 2 秒"
    $run = Invoke-Plain @('-NewList', "$longDir\new.csv")
    Assert-Equal '-NewList: 同じ名前のファイルがある' ('{0}/{1}' -f $run.Code, $run.Log.Contains('同じ名前のファイルがあるため')) '1/True'

    Write-Section '設定ファイル（-ConfigFile）'
    $cfgDir = Join-Path $workRoot 'cfg\設定 フォルダ'
    $cwd    = Join-Path $workRoot 'cwd'
    [void][IO.Directory]::CreateDirectory($cwd)
    $null = New-TextFile "$cfgDir\対象.csv" @('ホスト,ポート,メモ', "127.0.0.1,$port4,設定ファイルの CSV")
    $basic = New-TextFile "$cfgDir\basic.ini" @(
        '# 設定ファイルのテスト（日本語のコメント）'
        '; セミコロンのコメント'
        'List = 対象.csv'
        'Target = localhost'
        "target = `"127.0.0.1:$closed4`""
        'TIMEOUT=3'
        'OutputFolder = .\結果'
        'NoSave = false'
        'Repeat ='
    )
    # 別のフォルダから起動して、相対パスが設定ファイルのフォルダを基準にすることを確かめる
    $run = Invoke-From $cwd @('-ConfigFile', $basic)
    Assert-Equal '終了コード' $run.Code 2
    $runDir = Get-RunDir "$cfgDir\結果"
    Assert-Equal '保存先は設定ファイルのフォルダが基準' (Get-ResultFileNames $runDir) 'results.csv,summary.txt'
    Assert-Equal '起動したフォルダには保存しない' (Test-Path -LiteralPath "$cwd\結果") 'False'
    $x = Read-Result $runDir
    $targetLine = [array]::IndexOf($x.Lines, @($x.Lines | Where-Object { $_.StartsWith('対象: ') })[0])
    Assert-Equal '「対象」の前の行に設定ファイル' $(if ($targetLine -gt 0) { $x.Lines[$targetLine - 1] } else { '(なし)' }) "設定ファイル: $basic"
    Assert-Equal 'List・Target（2 行）・Timeout' (Get-TargetText $run.Log) "$cfgDir\対象.csv（1 件）、-Target（2 件）   タイムアウト 3 秒"
    Assert-Equal '設定ファイルの対象の判定' ('{0}|{1}|{2}' -f (Get-RowText $x 1), (Get-RowText $x 2), (Get-RowText $x 3)) "OK/127.0.0.1/応答あり/$port4/接続できた|OK/127.0.0.1/応答あり//-|NG/127.0.0.1/応答あり/$closed4/拒否された"

    # コマンドラインの値が優先する。-Target は置き換え（両方を合わせない）、-NoSave は NoSave = false より優先
    $run = Invoke-From $cwd @('-ConfigFile', $basic, '-Timeout', '1', '-Target', 'localhost', '-NoSave')
    Assert-Equal 'コマンドラインが優先' ('{0}/{1}/{2}' -f $run.Code, (Get-TargetText $run.Log), @(Get-ChildItem -LiteralPath "$cfgDir\結果" -Directory).Count) "0/$cfgDir\対象.csv（1 件）、-Target（1 件）   タイムアウト 1 秒/1"

    $i = 0
    foreach ($label in $encodings.Keys) {
        $file = New-TextFile "$cfgDir\enc$i.ini" @('List = ..\設定 フォルダ\対象.csv', 'NoSave = true', 'Timeout = 1') $encodings[$label]
        $run  = Invoke-From $cwd @('-ConfigFile', $file)
        Assert-Equal ('文字コード {0}: 日本語のパス' -f $label) ('{0}/{1}' -f $run.Code, (Get-TargetText $run.Log)) "0/$cfgDir\対象.csv（1 件）   タイムアウト 1 秒"
        $i++
    }

    # %TEMP% などの環境変数を展開する
    $env:WORKBENCH_TEST_ROOT = $workRoot
    $envIni = New-TextFile "$cfgDir\env.ini" @('List = %WORKBENCH_TEST_ROOT%\lists\english.csv', 'OutputFolder = %WORKBENCH_TEST_ROOT%\out\cfg-env')
    $run = Invoke-From $cwd @('-ConfigFile', $envIni)
    Assert-Equal '環境変数' ('{0}/{1}/{2}' -f $run.Code, (Get-TargetText $run.Log), (Get-ResultFileNames (Get-RunDir "$workRoot\out\cfg-env"))) "0/$english（2 件）   タイムアウト 2 秒/results.csv,summary.txt"

    Write-Section '設定ファイル: 書き間違い（確かめ始める前に止まる）'
    $mistakes = @(
        @('知らない項目', @('Target = localhost', 'Tiemout = 3'), @('設定ファイルの 2 行目: 知らない項目です: Tiemout', '使える項目: List, Target, Timeout, Repeat, OutputFolder, NoSave')),
        @('NewList は書けない', @('NewList = new.csv'), @('設定ファイルの 1 行目: 知らない項目です: NewList')),
        @('ConfigFile は書けない', @('ConfigFile = other.ini'), @('設定ファイルの 1 行目: 知らない項目です: ConfigFile')),
        @('同じ項目を 2 回', @('Timeout = 2', 'Target = localhost', 'timeout = 3'), @('設定ファイルの 3 行目: Timeout が 2 回書かれています（1 行目と 3 行目）')),
        @('Timeout が範囲外', @('Timeout = 61'), @('設定ファイルの 1 行目: Timeout は 1～60 の数値で指定してください: 61')),
        @('Repeat が範囲外', @('Repeat = 0'), @('設定ファイルの 1 行目: Repeat は 1～3600 の数値で指定してください: 0')),
        @('NoSave が true / false 以外', @('NoSave = yes'), @('設定ファイルの 1 行目: NoSave は true か false で指定してください: yes')),
        @('= がない行', @('Timeout 3'), @('「項目 = 値」の形になっていません: Timeout 3')),
        @('セクションの見出し', @('[Test-Connectivity]', 'Target = localhost'), @('設定ファイルの 1 行目: 「項目 = 値」の形になっていません: [Test-Connectivity]')),
        @('List の引用符が片方だけ', @('List = "C:\tools\targets.csv'), @('List の値に引用符（"）が残っています: "C:\tools\targets.csv')),
        @('Target の引用符が片方だけ', @('Target = "localhost'), @('Target の値に引用符（"）が残っています: "localhost')),
        @('定義されていない環境変数', @('OutputFolder = %WORKBENCH_NO_SUCH_VAR%\x'), @('定義されていない環境変数です: %WORKBENCH_NO_SUCH_VAR%'))
    )
    for ($i = 0; $i -lt $mistakes.Count; $i++) {
        $file    = New-TextFile "$cfgDir\err$i.ini" $mistakes[$i][1]
        $run     = Invoke-From $cwd @('-ConfigFile', $file)
        $missing = @(@($mistakes[$i][2]) + "ファイル: $file" | Where-Object { -not $run.Log.Contains($_) })
        Assert-Equal ('書き間違い: {0}' -f $mistakes[$i][0]) ('{0}/{1}/{2}' -f $run.Code, ($missing -join ' | '), $run.Log.Contains('疎通一括確認 (')) '1//False'
    }
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\nothing.ini")
    Assert-Equal '設定ファイルが見つからない' ('{0}/{1}' -f $run.Code, $run.Log.Contains("設定ファイルが見つかりません: $cfgDir\nothing.ini")) '1/True'
    $null = New-TextFile "$cfgDir\memo.ini.txt" @('Target = localhost')
    $run = Invoke-From $cwd @('-ConfigFile', "$cfgDir\memo.ini")
    Assert-Equal '設定ファイルの拡張子が .txt' ('{0}/{1}' -f $run.Code, $run.Log.Contains("拡張子が .txt になっていませんか: $cfgDir\memo.ini.txt")) '1/True'
    $noList = New-TextFile "$cfgDir\nolist.ini" @('NoSave = true', 'List = .\nosuch.csv')
    $run = Invoke-From $cwd @('-ConfigFile', $noList)
    Assert-Equal '設定ファイルの List が見つからない: 行番号' ('{0}/{1}' -f $run.Code, $run.Log.Contains("対象の CSV が見つかりません: $cfgDir\nosuch.csv（設定ファイルの 2 行目）")) '1/True'
    if ($freeLetter) {
        $badOut = New-TextFile "$cfgDir\badout.ini" @('Target = localhost', "OutputFolder = ${freeLetter}:\nope")
        $run = Invoke-From $cwd @('-ConfigFile', $badOut)
        Assert-Equal '設定ファイルの OutputFolder に書き込めない: 行番号' ('{0}/{1}' -f $run.Code, $run.Log.Contains('（設定ファイルの 2 行目）')) '0/True'
    }
    $cmdIni = New-TextFile "$workRoot\cmdlist\cmd.ini" @('List = targets.csv', 'NoSave = true')
    $run = Invoke-ViaCmd ('{0} -ConfigFile "{1}"' -f $cmdBase, $cmdIni)
    Assert-Equal 'コマンドプロンプトから -ConfigFile（スイッチの項目）' ('{0}/{1}/{2}' -f $run.Code, (Get-TargetText $run.Log), $run.Log.Contains('保存先')) "0/$cmdList（1 件）   タイムアウト 2 秒/False"

    Write-Section '繰り返し（-Repeat）'
    # 待ち受けを作ったり止めたりして判定を変え、Ctrl+C の代わりに Stop() で止める。
    # 1 回の確認が -Repeat より長くかかるように、応答のないアドレスも入れる（確認の途中で止めたときの回数を確かめるため）
    $repeatScript = New-HelperScript 'repeat.ps1' @(
        'param($Tool, $OutRoot, [int]$Port, $Targets)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''Target'', $Targets).AddParameter(''Repeat'', 1).AddParameter(''Timeout'', 1).AddParameter(''OutputFolder'', $OutRoot)'
        '$handle = $ps.BeginInvoke()'
        'function Get-Info { ($ps.Streams.Information | ForEach-Object { $_.MessageData.ToString() }) -join "`n" }'
        'function Wait-Text([string]$Pattern) {'
        '    $sw = [Diagnostics.Stopwatch]::StartNew()'
        '    while ($sw.Elapsed.TotalSeconds -lt 20) { if ((Get-Info) -match $Pattern) { return $true }; Start-Sleep -Milliseconds 50 }'
        '    $false'
        '}'
        '''ROUND1='' + (Wait-Text ''秒ごとに確かめています'')'
        '$listener = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, $Port)'
        '$listener.Start()'
        '''UP='' + (Wait-Text ''NG → OK'')'
        '$listener.Stop()'
        # ツールは変化の行を出してから結果を確定するので、その後の状態の行まで待ってから止める（確定の前に止めると、1 つ前の確認の結果が残る）
        '''DOWN='' + (Wait-Text ''(?s)OK → NG.*（確認 \d+ 回目  OK 1 件'')'
        '$ps.Stop()'
        '''STOPPED='' + $handle.IsCompleted'
        '''INFO='' + (Get-Info).Replace("`r", ''<CR>'').Replace("`n", ''<LF>'')'
    )
    $repeatPort = Get-ClosedPort ([Net.IPAddress]::Loopback)
    $repeatOut  = Join-Path $workRoot 'out\repeat'
    $repeatTargets = "127.0.0.1:$repeatPort,localhost"
    if ($blackHole) { $repeatTargets += ",$blackHole" }
    $result = Invoke-Helper $repeatScript @('-Tool', $tool, '-OutRoot', $repeatOut, '-Port', $repeatPort, '-Targets', $repeatTargets)
    Assert-Equal '判定が変わるまで動いた' ('{0}/{1}/{2}/{3}' -f $result['ROUND1'], $result['UP'], $result['DOWN'], $result['STOPPED']) 'True/True/True/True'
    $info = @("$($result['INFO'])" -split '<LF>')
    $changeShown = @($info | Where-Object { $_ -match '^\d{2}:\d{2}:\d{2}  (NG → OK|OK → NG)  ' } | ForEach-Object { $_.Substring(10) })
    Assert-Equal '画面に変化だけを表示' ($changeShown -join '|') ("NG → OK  1  127.0.0.1:$repeatPort  拒否された → 接続できた|OK → NG  1  127.0.0.1:$repeatPort  接続できた → 拒否された")
    $statusLines = @($info | ForEach-Object { [regex]::Matches($_, '（確認 (\d+) 回目  OK \d+ 件  NG \d+ 件）') } | ForEach-Object { $_ })
    $ngNow = if ($blackHole) { 2 } else { 1 }
    Assert-Equal '最後の行に回数と件数' $(if ($statusLines.Count -gt 0) { $statusLines[-1].Value -replace '\d+ 回目', 'N 回目' } else { '(なし)' }) ('（確認 N 回目  OK 1 件  NG {0} 件）' -f $ngNow)
    Assert-Equal '止めたときも保存先を表示' ($info -contains ('保存先: {0}' -f (Get-RunDir $repeatOut))) 'True'
    $repeatDir = Get-RunDir $repeatOut
    $x = Read-Result $repeatDir
    Assert-Equal '保存するファイル' (Get-ResultFileNames $repeatDir) 'changes.csv,results.csv,summary.txt'
    Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $repeatDir) 'changes.csv=True,results.csv=True,summary.txt=True'
    Assert-Equal 'changes.csv の見出し' (Get-FirstLine $repeatDir 'changes.csv') '"日時","No","ホスト","ポート","変化前","変化後","詳細"'
    Assert-Equal 'changes.csv の記録' (($x.Changes | ForEach-Object { '{0}/{1}/{2}/{3}/{4}/{5}' -f $_.'No', $_.'ホスト', $_.'ポート', $_.'変化前', $_.'変化後', $_.'詳細' }) -join '|') ("1/=""127.0.0.1""/$repeatPort/NG/OK/拒否された → 接続できた|1/=""127.0.0.1""/$repeatPort/OK/NG/接続できた → 拒否された")
    Assert-Equal 'changes.csv の日時の形' (@($x.Changes | Where-Object { $_.'日時' -notmatch '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$' }).Count) 0
    $changesBytes = if ($repeatDir) { [IO.File]::ReadAllBytes((Join-Path $repeatDir 'changes.csv')) } else { [byte[]]@() }
    $bomCount = 0
    for ($i = 0; $i -le $changesBytes.Length - 3; $i++) { if ($changesBytes[$i] -eq 0xEF -and $changesBytes[$i + 1] -eq 0xBB -and $changesBytes[$i + 2] -eq 0xBF) { $bomCount++ } }
    Assert-Equal 'changes.csv の BOM は先頭の 1 つだけ' $bomCount 1
    Assert-Equal 'results.csv は最後の確認' (Get-RowText $x 1) "NG/127.0.0.1/応答あり/$repeatPort/拒否された"
    $recordAt = [array]::IndexOf($x.Lines, '■ 判定が変わった記録')
    Assert-Equal 'summary.txt に変化の一覧' $(if ($recordAt -ge 0) { (@($x.Lines[($recordAt + 1)..($recordAt + 2)]) | ForEach-Object { $_.Substring(10) }) -join '|' } else { '(なし)' }) ($changeShown -join '|')
    $summaryRound = [regex]::Match($x.Summary, '繰り返し: 1 秒ごと   結果は (\d+) 回目の確認').Groups[1].Value
    if ($blackHole) {
        # 確認の途中で止めたとき、まだ終わっていない回の番号を書いていた
        Assert-Equal 'summary.txt の回数は、表の結果を確かめた回' $summaryRound $(if ($statusLines.Count -gt 0) { $statusLines[-1].Groups[1].Value } else { '(なし)' })
    } else {
        Assert-Equal 'summary.txt に回数' ($summaryRound -match '^\d+$') 'True'
        Write-Skip '確認の途中で止めたときの回数' '192.0.2.1 から応答が返るネットワークのため'
    }

    # 設定ファイルの Repeat も効くこと
    $repeatIni = New-TextFile "$cfgDir\repeat.ini" @('Target = localhost', 'Repeat = 2', 'Timeout = 1', 'NoSave = true')
    $repeatCfgScript = New-HelperScript 'repeat-config.ps1' @(
        'param($Tool, $ConfigFile)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddCommand($Tool).AddParameter(''ConfigFile'', $ConfigFile)'
        '$handle = $ps.BeginInvoke()'
        '$sw = [Diagnostics.Stopwatch]::StartNew()'
        'function Get-Info { ($ps.Streams.Information | ForEach-Object { $_.MessageData.ToString() }) -join "`n" }'
        'while ($sw.Elapsed.TotalSeconds -lt 20 -and (Get-Info) -notmatch ''確認 2 回目'') { Start-Sleep -Milliseconds 50 }'
        '$ps.Stop()'
        '''INFO='' + (Get-Info).Replace("`r", ''<CR>'').Replace("`n", ''<LF>'')'
    )
    $result = Invoke-Helper $repeatCfgScript @('-Tool', $tool, '-ConfigFile', $repeatIni)
    Assert-Equal '設定ファイルの Repeat' ('{0}/{1}/{2}' -f "$($result['INFO'])".Contains('繰り返し: 2 秒ごと'), "$($result['INFO'])".Contains('2 秒ごとに確かめています'), "$($result['INFO'])".Contains('保存先')) 'True/True/False'

    Write-Section 'Ctrl+C で中断（1 回目の確認の途中で停止）'
    if ($blackHole) {
        $stopScript = New-HelperScript 'stop-during-check.ps1' @(
            'param($Tool, $OutRoot, $Target)'
            '$ps = [powershell]::Create()'
            '[void]$ps.AddCommand($Tool).AddParameter(''Target'', $Target).AddParameter(''Timeout'', 10).AddParameter(''OutputFolder'', $OutRoot)'
            '$handle = $ps.BeginInvoke()'
            'Start-Sleep -Seconds 2'
            '''RUNNING='' + (-not $handle.IsCompleted)'
            '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
            '$ps.Stop()'
            '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
            '''MESSAGE='' + (($ps.Streams.Information | ForEach-Object { $_.MessageData.ToString() }) -join '' / '')'
        )
        $result = Invoke-Helper $stopScript @('-Tool', $tool, '-OutRoot', (Join-Path $workRoot 'out\ctrlc'), '-Target', $blackHole)
        Assert-Equal '止める前は確かめている途中' $result['RUNNING'] 'True'
        Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $result['BEFORE'], $result['AFTER']) '1/0'
        Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
    } else {
        Write-Skip 'Ctrl+C で中断' '192.0.2.1 から応答が返るネットワークのため'
    }

    Write-Section 'レビューで見つかった問題の再発防止'
    # 1. 0 で桁をそろえた IPv4 を 8 進数として読んでいた（上の「IP アドレスの書き方」で確かめている）
    # 2. メモのセル内改行で、表と CSV がくずれていた（上の「対象の CSV」で確かめている）
    # 3. ="..." の中の " を重ねていなかった・255 文字を超える値も囲んでいた（上の「対象の CSV」で確かめている）
    # 4. 接続時間を、すべての対象で同じ時刻から測っていた。応答の遅い対象のあとに確かめる対象でも、接続時間は短い
    if ($blackHole) {
        $run = Invoke-Tool 'tcp-ms' @('-Target', "$blackHole,127.0.0.1:$port4", '-Timeout', '2')
        $tcpMs = "$((Get-Row (Read-Result $run.Dir) 2).'接続時間(ms)')" -replace ',', ''
        Assert-Equal '接続時間は対象ごとに測る' ($tcpMs -match '^\d+$' -and [int]$tcpMs -lt 500) 'True'
    } else {
        Write-Skip '接続時間は対象ごとに測る' '192.0.2.1 から応答が返るネットワークのため'
    }
    # 5. .NET Framework の名前解決でスレッドが足りなくなり、名前解決の遅い対象のあとの対象まで待ち時間を超えていた。
    #    存在しない 1 語の名前は、LLMNR などで 2 秒以上かかる。その 49 件のあとに、すぐに解決できるこの PC の名前を置く
    $slowList = New-TextFile "$lists\slow.csv" (@('ホスト') + @(1..49 | ForEach-Object { 'wbtest-nosuch-{0:00}' -f $_ }) + [Environment]::MachineName)
    $run = Invoke-Tool 'slow-names' @('-List', $slowList, '-Timeout', '2')
    $x = Read-Result $run.Dir
    $slowCount = @($x.Results | Where-Object { $_.'詳細' -eq '名前解決: 待ち時間内に応答がありません' }).Count
    Assert-Equal '名前解決の遅い対象のあとの対象' ('{0}/{1}' -f (Get-Row $x 50).'判定', (Get-Row $x 50).'詳細') 'OK/'
    if ($slowCount -eq 0) { Write-Skip '名前解決の遅い対象' 'このネットワークでは存在しない名前がすぐに失敗するため、上の項目は手がかりにならない' }
    # 6. -Repeat を止めたとき、保存先を表示していなかった（上の「止めたときも保存先を表示」で確かめている）
    # 7. 値の途中の " を引用符の始まりとして読み、後ろの行が消えていた（上の「対象の CSV」の 19" ラックで確かめている）
    # 8. 閉じていない引用符があると、後ろの対象が黙って消えていた（上の「書き間違い: 引用符が閉じていない」で確かめている）

    Write-Section '多数の対象（50 件ずつ確かめる）'
    $manyList = New-TextFile "$lists\many.csv" (@('ホスト,ポート') + @(1..120 | ForEach-Object { if ($_ % 2) { "127.0.0.1,$port4" } else { 'localhost,' } }))
    $run = Invoke-Tool 'many' @('-List', $manyList)
    $x = Read-Result $run.Dir
    $elapsed = [double]([regex]::Match($run.Log, '所要 ([\d.]+)秒').Groups[1].Value)
    Assert-Equal '120 件: すべて OK' ('{0}/{1}' -f $run.Code, (Get-CountText $run.Log)) '0/結果: OK 120 件  NG 0 件'
    Assert-Equal '120 件: 番号の順' ((($x.Results | ForEach-Object { $_.'No' }) -join ',') -eq ((1..120) -join ',')) 'True'
    Assert-Equal '120 件: 所要時間（待ち時間より短い）' ($elapsed -gt 0 -and $elapsed -lt 6) 'True'

    Write-Section '地域設定'
    # 日付と数値の書式が、OS の地域設定によって変わらないこと
    $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
    $control = "$(& $Exe -NoProfile -Command $probe)"
    if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
        $cultureOut = Join-Path $workRoot 'out\culture'
        $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Target '127.0.0.1:{1}' -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $port4, $cultureOut.Replace("'", "''")
        $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $x = Read-Result (Get-RunDir $cultureOut)
        Assert-Equal 'fi-FI でも実行元の日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($x.Summary, '実行元: .*   \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}')) 'True'
        Assert-Equal 'fi-FI でも所要時間の小数点は .' ([regex]::IsMatch($x.Summary, '所要 \d+\.\d秒')) 'True'
        Assert-Equal 'fi-FI でも results.csv の日時は yyyy/MM/dd HH:mm:ss' ((Get-Row $x 1).'確認日時' -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
    } else {
        Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
    }
} finally {
    # 待ち受けを止め、アクセス拒否を戻してから消す
    if ($acceptLoop) { try { $acceptLoop.Stop(); $acceptLoop.Dispose() } catch { } }
    foreach ($l in $listeners) { try { $l.Stop() } catch { } }
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    Remove-Item Env:\WORKBENCH_TEST_ROOT -ErrorAction SilentlyContinue
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

'  END'
exit $script:failCount
