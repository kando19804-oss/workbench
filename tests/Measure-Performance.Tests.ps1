<#
.SYNOPSIS
    Measure-Performance の回帰テストです。

.DESCRIPTION
    この PC の性能カウンターとプロセスを、短い時間だけ記録して、画面・CSV・まとめを期待値と照合します。
    -Duration の最小は 1 分なので、ほとんどのテストは -Duration を秒として扱うように書き換えたコピー（fast）を動かします。
    本物の 1 分の記録も 1 回だけ行います。
    -Process のテストには、ping.exe を別の名前にコピーしたものを、ループバック（127.0.0.1）に向けて動かします（外へは通信しません）。
    本物の環境では起きない条件（保存の失敗、想定外のエラー、カウンターが読めない、目安を超えた値など）は、
    ツールの処理を 1 か所だけ書き換えたコピーで確かめます。
    「書き込めない保存先」を再現するため、テストが作ったフォルダにだけアクセス拒否を設定します（終了時に戻します）。
    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。ダミー環境は終了時に削除します。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Measure-Performance を実行する PowerShell（powershell.exe または pwsh）。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Measure-Performance\Measure-Performance.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Measure-Performance_{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag)
$tool      = Join-Path $workRoot 'tool\Measure-Performance.ps1'   # 本物のコピー（既定の保存先と 1 分の記録を確かめる）
$fast      = Join-Path $workRoot 'fast\Measure-Performance.ps1'   # -Duration を秒として扱うコピー
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
$inv       = [Globalization.CultureInfo]::InvariantCulture
$isAdmin   = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$userSid   = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$quick     = @('-Interval', '1', '-Duration', '3')   # fast で 3 秒（3 回）記録する
$script:failCount = 0

# -Duration を秒として扱うための書き換え（fast と、書き換えたコピーのすべてに当てる）
$FastPair = @('$endAt  = if ($Duration) { [double]$Duration * 60 - $lead }', '$endAt  = if ($Duration) { [double]$Duration * 1 - $lead }')

$Heads = @{
    'system.csv'       = '"日時","CPU 使用率(%)","最も忙しいコア(%)","プロセッサ待ち行列","メモリ使用率(%)","空きメモリ(MB)","コミット(MB)","コミットの上限(MB)","コミット率(%)","ページ読み込み(回/秒)","プロセス数"'
    'disks.csv'        = '"日時","ドライブ","使用時間(%)","応答時間(ms)","読み取りの応答時間(ms)","書き込みの応答時間(ms)","回数(回/秒)","読み取り(MB/秒)","書き込み(MB/秒)","待ち行列","空き(GB)","空き(%)"'
    'network.csv'      = '"日時","アダプター","説明","受信(Mbps)","送信(Mbps)","使用率(%)","速度(Mbps)"'
    'processes.csv'    = '"日時","プロセス","PID","CPU(%)","CPU の順位","メモリ(MB)","メモリの順位","コミット(MB)","I/O(MB/秒)","ハンドル数","スレッド数","指定","サービス"'
    'process-list.csv' = '"プロセス","PID","サービス","実行ユーザー","起動日時","実行ファイル","コマンドライン","最初に見た日時","最後に見た日時","CPU の平均(%)","CPU の最大(%)","メモリの最大(MB)","コミットの最大(MB)","コミットの増減(MB)"'
    'counters.csv'     = '"日時","カウンター","インスタンス","値"'
}

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

# 保存先のフォルダの下に作られた、結果のフォルダ
function Get-RunDir([string]$Parent) {
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'Performance_*' -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($found.Count -gt 0) { return $found[-1].FullName }
    $null
}

# Log には、画面への出力から改行を取り除いたものを入れる
function Invoke-Tool([string]$Label, [string[]]$ToolArgs, [string]$ToolPath = $fast) {
    $out  = Join-Path $workRoot "out\$Label"
    $log  = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }) -join "`n")
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = $log; Dir = (Get-RunDir $out); Out = $out }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行
# -RedirectInput は標準入力に空の行を渡す。コンソールのない環境（Claude Code など）では最初からリダイレクトされているが、
# 本物のコンソールでは渡さないとコンソールのままになり、-Duration なしの実行が止まらない
function Invoke-Plain([string[]]$ToolArgs, [string]$ToolPath = $fast, [switch]$RedirectInput) {
    if ($RedirectInput) {
        $lines = @('' | & $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1)
    } else {
        $lines = @(& $Exe -NoProfile -ExecutionPolicy Bypass -File $ToolPath @ToolArgs 2>&1)
    }
    $code = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = (ConvertTo-FlatText (($lines | ForEach-Object { "$_" }) -join "`n")) }
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

# $Exe で動かす補助のスクリプトを、一時フォルダに書く
function New-HelperScript([string]$Name, [string[]]$Lines) {
    $file = Join-Path $workRoot $Name
    [IO.File]::WriteAllLines($file, $Lines, $utf8Bom)
    $file
}

# 補助スクリプトの「名前=値」の出力を受け取る
function Invoke-Helper([string]$Script, [string[]]$HelperArgs = @()) {
    $result = @{}
    & $Exe -NoProfile -ExecutionPolicy Bypass -File $Script @HelperArgs 2>&1 | ForEach-Object {
        $line = "$_"
        $eq = $line.IndexOf('=')
        if ($eq -gt 0) { $result[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
    }
    $result
}

# 処理を書き換えたツールのコピーを作る（本物の環境では起きない条件を確かめるため）。
# Pairs は、書き換える前の文字と後の文字を交互に並べたもの。-Duration を秒として扱う書き換えも必ず当てる
function New-PatchedTool([string]$Name, [string[]]$Pairs) {
    $text = [IO.File]::ReadAllText($repoTool)
    $all = @($FastPair) + @($Pairs)
    for ($i = 0; $i -lt $all.Count; $i += 2) {
        if (-not $text.Contains($all[$i])) { throw "書き換える場所が見つかりません: $($all[$i])" }
        $text = $text.Replace($all[$i], $all[$i + 1])
    }
    $file = Join-Path $workRoot "patched\$Name\Measure-Performance.ps1"
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
    [IO.File]::WriteAllText($file, $text, $utf8Bom)
    $file
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
        Dir = $Dir; Lines = $lines; Text = ($lines -join "`n")
        System = @(Read-CsvOrEmpty 'system.csv'); Disks = @(Read-CsvOrEmpty 'disks.csv'); Network = @(Read-CsvOrEmpty 'network.csv')
        Processes = @(Read-CsvOrEmpty 'processes.csv'); List = @(Read-CsvOrEmpty 'process-list.csv'); Counters = @(Read-CsvOrEmpty 'counters.csv')
    }
}

# ファイル名を文字の番号の順に並べる（Sort-Object は processes.csv を process-list.csv より前にするなど、記号の扱いが違うため）
# 結果のファイルの 1 行目（フォルダやファイルがないときは空。テストを途中で止めないため）
function Get-FirstLine([string]$Dir, [string]$Name) {
    if (-not $Dir) { return '' }
    $file = Join-Path $Dir $Name
    if (-not [IO.File]::Exists($file)) { return '' }
    @([IO.File]::ReadAllLines($file))[0]
}

function Get-ResultFileNames([string]$Dir) {
    if (-not $Dir) { return '' }
    $names = [string[]]@(Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { $_.Name })
    [Array]::Sort($names, [StringComparer]::Ordinal)
    $names -join ','
}

# まとめの 1 行目の「（3秒・3 回）」の回数
function Get-RecordedCount([string]$Text) {
    $m = [regex]::Match($Text, '・([\d,]+) 回）')
    if (-not $m.Success) { return -1 }
    [int]($m.Groups[1].Value -replace ',', '')
}

# CSV の数値（3 桁区切り）を読む
function ConvertTo-Number([string]$Text) { [double]::Parse(($Text -replace ',', ''), $inv) }

# 「■ 見出し」の表の行（見出しの行と、空行の手前まで）
function Get-SectionLines($Lines, [string]$Title) {
    $inSection = $false
    foreach ($line in $Lines) {
        if ($line -eq ('■ ' + $Title)) { $inSection = $true; continue }
        if ($inSection) { if ($line -eq '' -or $line.StartsWith('■')) { break }; $line }
    }
}

# ファイルの先頭が BOM で、途中に BOM がないか
function Test-BomOnce([string]$FilePath) {
    $bytes = [IO.File]::ReadAllBytes($FilePath)
    if ($bytes.Length -lt 3 -or $bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) { return $false }
    for ($i = 3; $i -le $bytes.Length - 3; $i++) { if ($bytes[$i] -eq 0xEF -and $bytes[$i + 1] -eq 0xBB -and $bytes[$i + 2] -eq 0xBF) { return $false } }
    $true
}

function New-ConfigFile([string]$FilePath, [string[]]$Lines, [Text.Encoding]$Encoding = $utf8NoBom) {
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $FilePath))
    [IO.File]::WriteAllText($FilePath, (($Lines -join "`r`n") + "`r`n"), $Encoding)
    $FilePath
}

function Get-FreeLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    foreach ($c in [char[]]'QRSTUVWXYZ') { if ($used -notcontains [string]$c) { return [string]$c } }
    $null
}

# system.csv が指定の行数（見出しを含む）になるまで待つ（ほかのプロセスで動かしているツールが記録を始めたかを見る）
function Wait-SystemRows([string]$Out, [int]$Lines, [Diagnostics.Process]$Proc, [int]$Seconds = 60) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.Elapsed.TotalSeconds -lt $Seconds -and -not $Proc.HasExited) {
        $dir = Get-RunDir $Out
        if ($dir -and [IO.File]::Exists((Join-Path $dir 'system.csv'))) {
            $count = 0
            try {
                $stream = New-Object IO.FileStream((Join-Path $dir 'system.csv'), 'Open', 'Read', 'ReadWrite')
                try { $reader = New-Object IO.StreamReader($stream); while ($null -ne $reader.ReadLine()) { $count++ } } finally { $stream.Dispose() }
            } catch { }
            if ($count -ge $Lines) { return $dir }
        }
        Start-Sleep -Milliseconds 100
    }
    $null
}

#endregion

$denied = [System.Collections.Generic.List[string]]::new()
$dummyProc = $null
$peakProc = $null
try {
    #region 準備

    Write-Section ('準備: ダミー環境を作成（{0}）' -f $workRoot)
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $tool))
    [IO.File]::Copy($repoTool, $tool)
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $fast))
    $fastText = [IO.File]::ReadAllText($repoTool)
    if (-not $fastText.Contains($FastPair[0])) { throw "書き換える場所が見つかりません: $($FastPair[0])" }
    [IO.File]::WriteAllText($fast, $fastText.Replace($FastPair[0], $FastPair[1]), $utf8Bom)

    # -Process で記録するプロセス。ping.exe を別の名前にコピーし、ループバックに向けて動かす
    $dummyExe = Join-Path $workRoot 'bin\mp-test-ping.exe'
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $dummyExe))
    [IO.File]::Copy((Join-Path $env:SystemRoot 'System32\PING.EXE'), $dummyExe)
    $dummyProc = Start-Process -FilePath $dummyExe -ArgumentList '-n 900 127.0.0.1' -WindowStyle Hidden -PassThru
    # コミットの値を確かめるプロセス。200 MB を確保してすぐ解放するので、今のコミットと、これまでの最大のコミットが 200 MB 離れる
    # （7 から起動した 5.1 は 7 のモジュールの場所を引き継ぎ、Start-Sleep を読めないことがあるので、.NET で待つ）
    $peakProc = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') -WindowStyle Hidden -PassThru -ArgumentList @(
        '-NoProfile', '-Command', '"$p = [Runtime.InteropServices.Marshal]::AllocHGlobal(200MB); [Runtime.InteropServices.Marshal]::FreeHGlobal($p); [Threading.Thread]::Sleep(900000)"')
    Start-Sleep -Seconds 3   # PowerShell が起動して、確保と解放を終えるのを待つ
    '  ダミーのプロセス: {0}（PID {1}）、コミットの確認（PID {2}）' -f $dummyExe, $dummyProc.Id, $peakProc.Id

    #endregion

    #region 基本

    Write-Section '基本（-Interval 1・3 秒）'
    $run = Invoke-Tool 'basic' $quick
    $x = Read-Result $run.Dir
    $n = Get-RecordedCount $x.Text
    Assert-Equal '終了コード' $run.Code 0
    Assert-Equal '保存したファイル' (Get-ResultFileNames $run.Dir) 'disks.csv,network.csv,process-list.csv,processes.csv,summary.txt,system.csv'
    foreach ($name in 'system.csv', 'disks.csv', 'network.csv', 'processes.csv', 'process-list.csv') {
        Assert-Equal ('見出し: ' + $name) (Get-FirstLine $run.Dir $name) $Heads[$name]
    }
    Assert-Equal 'すべてのファイルが BOM 付きで、途中に BOM がない' (@(Get-ChildItem -LiteralPath $run.Dir -File | Where-Object { -not (Test-BomOnce $_.FullName) }).Count) 0
    Assert-Equal '記録した回数（2 回以上）' ($n -ge 2) 'True'
    Assert-Equal 'system.csv の行は記録した回数' $x.System.Count $n
    Assert-Equal '日時は yyyy/MM/dd HH:mm:ss' (@($x.System | Where-Object { $_.'日時' -notmatch '^\d{4}/\d\d/\d\d \d\d:\d\d:\d\d$' }).Count) 0
    # -Interval 1 でも、続く 2 回が同じ秒にならない（予定の時刻を秒の区切りにそろえ、四捨五入する）
    Assert-Equal 'system.csv の日時は重ならない' @($x.System | ForEach-Object { $_.'日時' } | Select-Object -Unique).Count $x.System.Count
    Assert-Equal 'CPU 使用率は 0～100 の小数 1 桁' (@($x.System | Where-Object { $_.'CPU 使用率(%)' -notmatch '^\d{1,3}\.\d$' -or (ConvertTo-Number $_.'CPU 使用率(%)') -gt 100 }).Count) 0
    Assert-Equal '空きメモリは 3 桁区切りの小数 1 桁' (@($x.System | Where-Object { $_.'空きメモリ(MB)' -notmatch '^\d{1,3}(,\d{3})*\.\d$' }).Count) 0
    $systemDrive = $env:SystemDrive.ToUpperInvariant()
    Assert-Equal 'disks.csv にシステムのドライブが毎回ある' (@($x.Disks | Where-Object { $_.'ドライブ' -eq $systemDrive }).Count) $n
    Assert-Equal '応答時間は小数 2 桁' (@($x.Disks | Where-Object { $_.'応答時間(ms)' -notmatch '^[\d,]+\.\d\d$' }).Count) 0
    $times = @($x.Processes | ForEach-Object { $_.'日時' } | Select-Object -Unique)
    $tooMany = @($times | Where-Object { $t = $_; @($x.Processes | Where-Object { $_.'日時' -eq $t }).Count -gt 10 }).Count
    Assert-Equal 'processes.csv は毎回 10 行まで（CPU とメモリの上位 5 件ずつ）' ('{0}/{1}' -f $times.Count, $tooMany) ('{0}/0' -f $n)
    Assert-Equal 'processes.csv の行は、CPU かメモリの 5 位以内' (@($x.Processes | Where-Object { [int]$_.'CPU の順位' -gt 5 -and [int]$_.'メモリの順位' -gt 5 }).Count) 0
    Assert-Equal 'processes.csv に、毎回メモリの 1～5 位がそろう' (@($x.Processes | Where-Object { [int]$_.'メモリの順位' -le 5 }).Count) (5 * $n)
    Assert-Equal 'processes.csv に、毎回 CPU の 1～5 位がそろう' (@($x.Processes | Where-Object { [int]$_.'CPU の順位' -le 5 }).Count) (5 * $n)
    # プロセスの CPU(%) は PC 全体を 100% とした値なので、すべてのプロセスの平均の合計は、CPU 使用率の平均とほぼ同じになる
    # （論理プロセッサの数で割り忘れると、その数の倍になる）
    $cpuAverage = ($x.System | ForEach-Object { ConvertTo-Number $_.'CPU 使用率(%)' } | Measure-Object -Average).Average
    $processSum = ($x.List | ForEach-Object { ConvertTo-Number $_.'CPU の平均(%)' } | Measure-Object -Sum).Sum
    Assert-Equal ('プロセスの CPU の平均の合計（{0:N1}）は、CPU 使用率の平均（{1:N1}）× 1.5 + 3 以下' -f $processSum, $cpuAverage) ($processSum -le $cpuAverage * 1.5 + 3) 'True'
    Assert-Equal 'CPU の順位は、回ごとに重ならない' (@($x.Processes | Group-Object -Property '日時', 'CPU の順位' | Where-Object { $_.Count -gt 1 }).Count) 0
    Assert-Equal 'process-list.csv に行がある' ($x.List.Count -gt 50) 'True'
    Assert-Equal '起動日時は yyyy/MM/dd HH:mm:ss か空' (@($x.List | Where-Object { $_.'起動日時' -and $_.'起動日時' -notmatch '^\d{4}/\d\d/\d\d \d\d:\d\d:\d\d$' }).Count) 0
    Assert-Equal '指定の列は空（-Process なし）' (@($x.Processes | Where-Object { $_.'指定' -ne '' }).Count) 0
    $allText = (@(Get-ChildItem -LiteralPath $run.Dir -Filter '*.csv' | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n")
    Assert-Equal 'CSV に -0.0 がない（7 も 5.1 と同じ）' ([regex]::IsMatch($allText, '"-0(\.0+)?"')) 'False'
    Assert-Equal 'まとめ: 1 行目' $x.Lines[0] '性能の記録 (Measure-Performance 1.0.0)'
    Assert-Equal 'まとめ: 対象の行' ([regex]::IsMatch($x.Text, ('対象: {0}（管理者: {1}）' -f [regex]::Escape($env:COMPUTERNAME), $(if ($isAdmin) { 'はい' } else { 'いいえ' })))) 'True'
    Assert-Equal 'まとめ: CPU の行に論理プロセッサの数' ([regex]::IsMatch($x.Text, ('（論理プロセッサ {0}）   メモリ [\d,]+\.\d GB' -f [Environment]::ProcessorCount))) 'True'
    Assert-Equal 'まとめ: 記録の行' ($x.Text.Contains('記録: 1 秒ごと・3 分間（') -and $x.Text.Contains('プロセス: CPU とメモリの上位 5 件')) 'True'
    Assert-Equal 'まとめ: 終わり方' ([regex]::IsMatch($x.Text, '記録を終えました（指定の時間）  \d\d:\d\d:\d\d ～ \d\d:\d\d:\d\d')) 'True'
    $sections = @($x.Lines | Where-Object { $_.StartsWith('■') }) -join ','
    Assert-Equal 'まとめ: 表の並び' ($sections -replace '■ ネットワーク,', '') '■ CPU とメモリ,■ ディスク,■ CPU を多く使ったプロセス（期間の平均）,■ メモリを多く使ったプロセス（期間の最大）,■ 目安を超えたところ'
    Assert-Equal 'まとめ: CPU を多く使ったプロセスは 5 行' @(Get-SectionLines $x.Lines 'CPU を多く使ったプロセス（期間の平均）' | Where-Object { $_ -match '^\s+\d+\s' }).Count 5
    Assert-Equal 'まとめ: このツールの CPU 使用率' ([regex]::IsMatch($x.Text, 'このツールの CPU 使用率: 平均 \d+\.\d%')) 'True'
    Assert-Equal 'まとめ: 管理者でないときだけ、管理者の注意' $x.Text.Contains('※ 管理者として実行すると') (-not $isAdmin)
    Assert-Equal '画面: 記録中の行が回数分' ([regex]::Matches($run.Log, '\d\d:\d\d:\d\d +\d+%').Count) $n
    Assert-Equal '画面: 見出しと保存先' ($run.Log.Contains('時刻') -and $run.Log.Contains('ディスク（最も忙しい）') -and $run.Log.Contains('保存先: ' + $run.Dir)) 'True'
    # 「対象」の時刻・フォルダ名・まとめの始めがそろう（レビュー 軽微 5）
    $targetTime = [regex]::Match($x.Text, '対象: .+?(\d{4})-(\d\d)-(\d\d) (\d\d):(\d\d):(\d\d)')
    $folderTime = [regex]::Match((Split-Path -Leaf $run.Dir), '_(\d{8})_(\d{6})$')
    $startShown = [regex]::Match($x.Text, '記録を終えました（指定の時間）  (\d\d):(\d\d):(\d\d)')
    Assert-Equal '「対象」・フォルダ名・まとめの始めの時刻がそろう' ('{0}/{1}' -f (($targetTime.Groups[4..6] | ForEach-Object { $_.Value }) -join ''), (($startShown.Groups[1..3] | ForEach-Object { $_.Value }) -join '')) ('{0}/{0}' -f $folderTime.Groups[2].Value)
    # 最初の記録は始めの時刻の 1 間隔後、まとめの期間は始めの時刻から指定の時間の終わりまで
    # （始めの時刻を取ったあとで秒の区切りまで待っていたため、記録が 1 秒遅れ、期間が 1 秒長く見えていた。実機確認で見つけた）
    $startAt = [datetime]::ParseExact($targetTime.Value.Substring($targetTime.Value.Length - 19), 'yyyy-MM-dd HH:mm:ss', $inv)
    $firstAt = [datetime]::ParseExact($x.System[0].'日時', 'yyyy/MM/dd HH:mm:ss', $inv)
    $period  = [regex]::Match($x.Text, '記録を終えました（指定の時間）  (\d\d:\d\d:\d\d) ～ (\d\d:\d\d:\d\d)（(.+?)・')
    Assert-Equal '最初の記録は、始めの時刻の 1 秒後（-Interval 1）' ($firstAt - $startAt).TotalSeconds 1
    Assert-Equal 'まとめの期間は、始めの時刻から 3 秒（指定の時間の終わり）' ('{0} ～ {1}（{2}）' -f $period.Groups[1].Value, $period.Groups[2].Value, $period.Groups[3].Value) ('{0} ～ {1}（3秒）' -f $startAt.ToString('HH:mm:ss', $inv), $startAt.AddSeconds(3).ToString('HH:mm:ss', $inv))
    # 始めの時刻のあとの準備に 0.6 秒かかっても、最初の回は 1 秒後のまま（最初の回まで 0.2 秒以上あるので、ずらさない）
    # （待つのは起点の読み取りとストップウォッチより前。保存先の準備が長引いたときと同じ形にする）
    $slowPrep = New-PatchedTool 'slowprep' @('    # 準備に時間がかかっても、最初の回が 1 間隔の平均になるよう', ('    [Threading.Thread]::Sleep(600)' + "`n" + '    # 準備に時間がかかっても、最初の回が 1 間隔の平均になるよう'))
    $slowRun = Invoke-Tool 'slowprep' $quick $slowPrep
    $slowX = Read-Result $slowRun.Dir
    $slowStart = [regex]::Match($slowX.Text, '対象: .+?(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)')
    Assert-Equal '準備に 0.6 秒かかっても、最初の記録は始めの時刻の 1 秒後' ([datetime]::ParseExact($slowX.System[0].'日時', 'yyyy/MM/dd HH:mm:ss', $inv) - [datetime]::ParseExact($slowStart.Groups[1].Value, 'yyyy-MM-dd HH:mm:ss', $inv)).TotalSeconds 1

    # 本物の値の妥当性
    # コミットは「今の」値（これまでの最大ではない）。200 MB を確保して解放したプロセスで、Get-Process のプライベートメモリと比べる
    $peakItem = @($x.List | Where-Object { $_.'PID' -eq [string]$peakProc.Id })
    $peakNow = Get-Process -Id $peakProc.Id -ErrorAction SilentlyContinue
    if ($peakItem.Count -gt 0 -and $peakNow -and ($peakNow.PeakPagedMemorySize64 - $peakNow.PrivateMemorySize64) -gt 100MB) {
        Assert-Equal 'コミットは今の値（200 MB を解放したプロセスで、Get-Process との差が 2 MB 未満）' ([math]::Abs((ConvertTo-Number $peakItem[0].'コミットの最大(MB)') - $peakNow.PrivateMemorySize64 / 1MB) -lt 2) 'True'
    } else {
        Write-Skip 'コミットは今の値' 'コミットの確認のプロセスが、確保と解放を終えていないため'
    }
    if ($x.Network.Count -gt 0) {
        $connectionNames = @(Get-CimInstance -ClassName Win32_NetworkAdapter -Property NetConnectionID | Where-Object { $_.NetConnectionID } | ForEach-Object { [string]$_.NetConnectionID })
        Assert-Equal 'network.csv のアダプターは、接続の名前か説明' (@($x.Network | Where-Object { $connectionNames -notcontains $_.'アダプター' -and $_.'アダプター' -ne $_.'説明' }).Count) 0
    } else {
        Write-Skip 'ネットワークのアダプターの名前' 'アダプターがないため'
    }

    #endregion

    #region 既定の保存先と、本物の 1 分の記録

    Write-Section '既定の保存先（スクリプトのフォルダ）と、本物の -Duration 1（1 分）'
    $toolDir = Split-Path -Parent $tool
    $real = Invoke-Plain @('-Interval', '30', '-Duration', '1') $tool
    $realDir = Get-RunDir $toolDir
    $x = Read-Result $realDir
    Assert-Equal '本物: 終了コード' $real.Code 0
    Assert-Equal '本物: スクリプトのフォルダに保存' ([bool]$realDir) 'True'
    Assert-Equal '本物: 1 分で 2 回（30 秒ごと）・指定の時間で終わる' ('{0}/{1}' -f (Get-RecordedCount $x.Text), $x.Text.Contains('記録を終えました（指定の時間）')) '2/True'
    Assert-Equal '本物: 記録の行' $x.Text.Contains('記録: 30 秒ごと・1 分間（') 'True'
    # 記録は始めの時刻の 30 秒後と 60 秒後。「（～まで）」・最後の記録・まとめの終わりがそろう
    $realStart  = [regex]::Match($x.Text, '対象: .+?(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)')
    $realPlan   = [regex]::Match($x.Text, '記録: 30 秒ごと・1 分間（(\d\d:\d\d:\d\d) まで）')
    $realPeriod = [regex]::Match($x.Text, '記録を終えました（指定の時間）  (\d\d:\d\d:\d\d) ～ (\d\d:\d\d:\d\d)（(.+?)・')
    $realTimes  = @($x.System | ForEach-Object { $_.'日時'.Substring(11) }) -join ','
    $s = [datetime]::ParseExact($realStart.Groups[1].Value, 'yyyy-MM-dd HH:mm:ss', $inv)
    Assert-Equal '本物: 記録の時刻・「まで」・まとめの期間がそろう' ('{0}|{1}|{2} ～ {3}（{4}）' -f $realTimes, $realPlan.Groups[1].Value, $realPeriod.Groups[1].Value, $realPeriod.Groups[2].Value, $realPeriod.Groups[3].Value) ('{0},{1}|{1}|{2} ～ {1}（1分0秒）' -f $s.AddSeconds(30).ToString('HH:mm:ss', $inv), $s.AddSeconds(60).ToString('HH:mm:ss', $inv), $s.ToString('HH:mm:ss', $inv))

    #endregion

    #region -Top

    Write-Section '-Top'
    $run = Invoke-Tool 'top2' ($quick + @('-Top', '2'))
    $x = Read-Result $run.Dir
    Assert-Equal '-Top 2: 行は CPU かメモリの 2 位以内' (@($x.Processes | Where-Object { [int]$_.'CPU の順位' -gt 2 -and [int]$_.'メモリの順位' -gt 2 }).Count) 0
    Assert-Equal '-Top 2: 毎回 4 行まで' (@($x.Processes | Group-Object -Property '日時' | Where-Object { $_.Count -gt 4 }).Count) 0
    Assert-Equal '-Top 2: 記録の行とまとめの件数' ('{0}/{1}' -f $x.Text.Contains('プロセス: CPU とメモリの上位 2 件'), @(Get-SectionLines $x.Lines 'CPU を多く使ったプロセス（期間の平均）' | Where-Object { $_ -match '^\s+\d+\s' }).Count) 'True/2'
    $run = Invoke-Tool 'top0' ($quick + @('-Top', '0'))
    $x = Read-Result $run.Dir
    Assert-Equal '-Top 0: processes.csv は見出しだけ・まとめは 5 件・process-list.csv はある' ('{0}/{1}/{2}/{3}' -f $x.Processes.Count, @(Get-SectionLines $x.Lines 'CPU を多く使ったプロセス（期間の平均）' | Where-Object { $_ -match '^\s+\d+\s' }).Count, ($x.List.Count -gt 0), $x.Text.Contains('プロセス: 毎回は記録しない')) '0/5/True/True'

    #endregion

    #region -Process

    Write-Section '-Process（ping.exe のコピーをループバックに向けて動かす）'
    $dummyPid = [string]$dummyProc.Id
    $run = Invoke-Tool 'process' ($quick + @('-Top', '0', '-Process', 'mp-test-*,nosuch_mp_proc,mp-test-*'))
    $x = Read-Result $run.Dir
    $n = Get-RecordedCount $x.Text
    $rows = @($x.Processes | Where-Object { $_.'PID' -eq $dummyPid })
    Assert-Equal '-Process: 毎回、指定したプロセスだけを記録（指定は「はい」）' ('{0}/{1}/{2}' -f $rows.Count, @($rows | Where-Object { $_.'指定' -eq 'はい' -and $_.'プロセス' -eq 'mp-test-ping.exe' }).Count, $x.Processes.Count) ('{0}/{0}/{0}' -f $n)
    $item = @($x.List | Where-Object { $_.'PID' -eq $dummyPid })[0]
    # 一時フォルダは 8.3 形式の名前のことがあるので、Get-Process の実行ファイル（長い名前）と比べる
    $procNow = Get-Process -Id $dummyProc.Id
    Assert-Equal '-Process: process-list.csv の実行ファイル・コマンドライン・実行ユーザー' ('{0}/{1}/{2}' -f ($item.'実行ファイル' -eq $procNow.Path), $item.'コマンドライン'.Contains('127.0.0.1'), ($item.'実行ユーザー' -ne '')) 'True/True/True'
    Assert-Equal '-Process: コミットは Get-Process のプライベートメモリと同じくらい（差が 2 MB 未満）' ([math]::Abs((ConvertTo-Number $item.'コミットの最大(MB)') - $procNow.PrivateMemorySize64 / 1MB) -lt 2) 'True'
    Assert-Equal '-Process: 記録の行（重なった名前は 1 つにする）' ([regex]::IsMatch($x.Text, '   プロセス: 指定 mp-test-\*, nosuch_mp_proc(\r?\n|$)')) 'True'
    Assert-Equal '-Process: まとめの表と、見つからなかった名前' ('{0}/{1}' -f @(Get-SectionLines $x.Lines '指定したプロセス（-Process）' | Where-Object { $_ -match ('mp-test-ping\.exe\s+' + $dummyPid) }).Count, $x.Text.Contains('※ -Process に当たるプロセスは、期間中に見つかりませんでした: nosuch_mp_proc')) '1/True'
    $run = Invoke-Tool 'process-full' ($quick + @('-Top', '0', '-Process', 'MP-TEST-PING.EXE'))
    $x = Read-Result $run.Dir
    Assert-Equal '-Process: .exe を付けた名前・大文字小文字の違い' (@($x.Processes | Where-Object { $_.'PID' -eq $dummyPid }).Count -gt 0) 'True'
    $run = Invoke-Tool 'process-bracket' ($quick + @('-Top', '0', '-Process', '[m]p-test-ping'))
    $x = Read-Result $run.Dir
    Assert-Equal '-Process: [ ] はそのままの文字（当たらない）' ('{0}/{1}' -f $x.Processes.Count, $x.Text.Contains('期間中に見つかりませんでした: [m]p-test-ping')) '0/True'

    #endregion

    #region -Counter

    Write-Section '-Counter'
    $run = Invoke-Tool 'counter' ($quick + @('-Counter', '\Processor Information(_Total)\% Processor Time,\Memory\Available MBytes,\Processor Information(0,0)\% Processor Time'))
    $x = Read-Result $run.Dir
    $n = Get-RecordedCount $x.Text
    Assert-Equal '-Counter: 終了コード・counters.csv の見出し' ('{0}/{1}' -f $run.Code, ((Get-FirstLine $run.Dir 'counters.csv') -eq $Heads['counters.csv'])) '0/True'
    Assert-Equal '-Counter: 毎回 3 行（かっこの中のカンマは区切りにしない）' $x.Counters.Count (3 * $n)
    $instances = @($x.Counters | ForEach-Object { $_.'インスタンス' } | Select-Object -Unique)
    Assert-Equal '-Counter: インスタンスの列（_Total・なし・="0,0"）' ('{0}/{1}/{2}/{3}' -f $instances.Count, ($instances -contains '_Total'), ($instances -contains ''), ($instances -contains '="0,0"')) '3/True/True/True'
    $memRow = @($x.Counters | Where-Object { $_.'カウンター' -eq '\Memory\Available MBytes' })[0]
    Assert-Equal '-Counter: 値は 3 桁区切り・小数 3 桁まで' ($memRow.'値' -match '^\d{1,3}(,\d{3})*(\.\d{1,3})?$') 'True'
    Assert-Equal '-Counter: 記録の行の次に「追加のカウンター」・まとめの表' ('{0}/{1}' -f $x.Text.Contains('追加のカウンター: \Processor Information(_Total)\% Processor Time, \Memory\Available MBytes, \Processor Information(0,0)\% Processor Time'), @(Get-SectionLines $x.Lines '追加のカウンター').Count) 'True/4'

    # 同じ名前のインスタンスを #1、#2 で区別する（レビュー 中 1）
    $run = Invoke-Tool 'counter-dup' ($quick + @('-Counter', '\Process(*)\ID Process'))
    $x = Read-Result $run.Dir
    $first = @($x.Counters | Where-Object { $_.'日時' -eq $x.Counters[0].'日時' })
    $names = @($first | ForEach-Object { $_.'インスタンス' })
    Assert-Equal '-Counter \Process(*): 1 回の中のインスタンスは重ならない・svchost#1 がある' ('{0}/{1}' -f ($names.Count -eq @($names | Select-Object -Unique).Count), ($names -contains 'svchost#1')) 'True/True'
    Assert-Equal '-Counter \Process(*): まとめは 20 件と「ほか N 件は counters.csv」' ('{0}/{1}' -f @(Get-SectionLines $x.Lines '追加のカウンター' | Where-Object { $_.StartsWith(' \Process') }).Count, [regex]::IsMatch($x.Text, '（ほか [\d,]+ 件は counters\.csv を参照してください）')) '20/True'

    # 表示の言語の名前（この PowerShell の Get-Counter で読める名前で、英語でないもの）
    $localPath = "$(& $Exe -NoProfile -Command '$s = @(Get-Counter -ListSet * -ErrorAction SilentlyContinue | Where-Object { $_.CounterSetName -match ''[^\x00-\x7F]'' }); if ($s.Count -gt 0) { $s[0].Paths[0] }' 2>$null)".Trim()
    if ($localPath) {
        $run = Invoke-Tool 'counter-local' ($quick + @('-Counter', $localPath))
        Assert-Equal ('-Counter: 表示の言語の名前でも読める（{0}）' -f $localPath) ('{0}/{1}' -f $run.Code, $run.Log.Contains('追加のカウンター: ' + $localPath)) '0/True'
    } else {
        Write-Skip '-Counter の表示の言語の名前' '英語でない名前のカウンターがないため'
    }

    $errOut = Join-Path $workRoot 'out\counter-errors'
    $cases = @(
        @('Processor(_Total)\% Processor Time', '-Counter の書き方が違います: Processor(_Total)\% Processor Time（\オブジェクト(インスタンス)\カウンター の形で書いてください）'),
        @('\\SERVER02\Memory\Available MBytes', '-Counter には、ほかのコンピューターのカウンターは書けません: \\SERVER02\Memory\Available MBytes'),
        @('\Web Servce(_Total)\Current Connections', '-Counter のカウンターが見つかりません: \Web Servce(_Total)\Current Connections（'),
        @('\LogicalDisk(Q9:)\% Free Space', '-Counter のカウンターが見つかりません: \LogicalDisk(Q9:)\% Free Space（')
    )
    foreach ($case in $cases) {
        $r = Invoke-Plain @('-Duration', '3', '-Counter', $case[0], '-OutputFolder', $errOut)
        Assert-Equal ('-Counter の誤り: ' + $case[0]) ('{0}/{1}/{2}' -f $r.Code, $r.Log.Contains($case[1]), [bool](Get-RunDir $errOut)) '1/True/False'
    }
    Assert-Equal '-Counter が見つからないときの案内' $r.Log.Contains('英語の名前か、この PC のパフォーマンス モニターに表示される名前で書いてください') 'True'

    #endregion

    #region -NoSave

    Write-Section '-NoSave'
    $fastDir = Split-Path -Parent $fast
    $r = Invoke-Plain ($quick + @('-NoSave', '-Counter', '\Process(*)\ID Process'))
    Assert-Equal '-NoSave: 終了コード・保存しない・保存先の行がない' ('{0}/{1}/{2}' -f $r.Code, [bool](Get-RunDir $fastDir), $r.Log.Contains('保存先:')) '0/False/False'
    Assert-Equal '-NoSave: 表示しきれない分は件数だけ' ('{0}/{1}' -f [regex]::IsMatch($r.Log, '（ほか [\d,]+ 件）'), $r.Log.Contains('counters.csv を参照')) 'True/False'
    Assert-Equal '-NoSave: 管理者の注意は出さない' $r.Log.Contains('※ 管理者として実行すると') 'False'

    #endregion

    #region 設定ファイル

    Write-Section '設定ファイル（-ConfigFile）'
    $cfgDir = Join-Path $workRoot 'cfg\設定 フォルダ'
    $cfg = New-ConfigFile (Join-Path $cfgDir 'nightly.ini') @(
        '# Measure-Performance の設定ファイル'
        '; コメント'
        'Interval = 1'
        'Duration = 3'
        'Top = 1'
        'Process = mp-test-ping'
        'Process = "nosuch_a","nosuch_b"'
        'Counter = \Processor Information(0,0)\% Processor Time, \Memory\Available MBytes'
        'OutputFolder = 結果 フォルダ'
        'NoSave ='
    )
    $r = Invoke-Plain @('-ConfigFile', $cfg)
    $cfgRun = Get-RunDir (Join-Path $cfgDir '結果 フォルダ')
    $x = Read-Result $cfgRun
    Assert-Equal '設定ファイル: 終了コード・相対パスは設定ファイルのフォルダが基準' ('{0}/{1}' -f $r.Code, [bool]$cfgRun) '0/True'
    Assert-Equal '設定ファイル: 設定ファイルの行・記録の行' ('{0}/{1}' -f $x.Text.Contains('設定ファイル: ' + $cfg), $x.Text.Contains('記録: 1 秒ごと・3 分間（')) 'True/True'
    Assert-Equal '設定ファイル: Process は複数の行を合わせる' $x.Text.Contains('プロセス: CPU とメモリの上位 1 件、指定 mp-test-ping, nosuch_a, nosuch_b') 'True'
    Assert-Equal '設定ファイル: Counter（かっこの中のカンマ）' $x.Text.Contains('追加のカウンター: \Processor Information(0,0)\% Processor Time, \Memory\Available MBytes') 'True'
    # コマンドラインが優先（Process は設定ファイルのものを使わない）
    $r = Invoke-Plain @('-ConfigFile', $cfg, '-Process', 'nosuch_c', '-Duration', '2')
    $x = Read-Result (Get-RunDir (Join-Path $cfgDir '結果 フォルダ'))
    Assert-Equal '設定ファイル: コマンドラインの -Process と -Duration が優先' ('{0}/{1}/{2}' -f $x.Text.Contains('、指定 nosuch_c'), $x.Text.Contains('nosuch_a'), $x.Text.Contains('記録: 1 秒ごと・2 分間（')) 'True/False/True'

    # 文字コード（Shift_JIS・UTF-16）と環境変数
    $env:MP_TEST_OUT = Join-Path $workRoot 'out\環境変数'
    $cfgSjis = New-ConfigFile (Join-Path $cfgDir 'sjis.ini') @('# 日本語のコメント', 'Duration = 2', 'Interval = 1', 'OutputFolder = %MP_TEST_OUT%\シフトJIS') ([Text.Encoding]::GetEncoding(932))
    $r = Invoke-Plain @('-ConfigFile', $cfgSjis)
    Assert-Equal '設定ファイル: Shift_JIS・環境変数' ('{0}/{1}' -f $r.Code, [bool](Get-RunDir (Join-Path $env:MP_TEST_OUT 'シフトJIS'))) '0/True'
    $cfgUtf16 = New-ConfigFile (Join-Path $cfgDir 'utf16.ini') @('Duration = 2', 'Interval = 1', 'NoSave = true') ([Text.Encoding]::Unicode)
    $r = Invoke-Plain @('-ConfigFile', $cfgUtf16)
    Assert-Equal '設定ファイル: UTF-16・NoSave = true' ('{0}/{1}' -f $r.Code, $r.Log.Contains('保存先:')) '0/False'

    # 書き間違い（記録を始める前に止まる）
    $badCases = @(
        @(@('Intervl = 5'), '設定ファイルの 1 行目: 知らない項目です: Intervl'),
        @(@('Interval 5'), '設定ファイルの 1 行目: 「項目 = 値」の形になっていません: Interval 5'),
        @(@('Interval = 0'), '設定ファイルの 1 行目: Interval は 1～3600 の数値で指定してください: 0'),
        @(@('Duration = 10081'), '設定ファイルの 1 行目: Duration は 1～10080 の数値で指定してください: 10081'),
        @(@('Top = 51'), '設定ファイルの 1 行目: Top は 0～50 の数値で指定してください: 51'),
        @(@('Interval = 5', 'Interval = 6'), '設定ファイルの 2 行目: Interval が 2 回書かれています（1 行目と 2 行目）'),
        @(@('NoSave = yes'), '設定ファイルの 1 行目: NoSave は true か false で指定してください: yes'),
        @(@('OutputFolder = "D:\perf'), '設定ファイルの 1 行目: OutputFolder の値に引用符（"）が残っています: "D:\perf'),
        @(@('OutputFolder = %NO_SUCH_MP_VAR%\perf'), '設定ファイルの 1 行目: 定義されていない環境変数です: %NO_SUCH_MP_VAR%')
    )
    $i = 0
    foreach ($case in $badCases) {
        $i++
        $bad = New-ConfigFile (Join-Path $cfgDir ('bad{0}.ini' -f $i)) $case[0]
        $r = Invoke-Plain @('-ConfigFile', $bad, '-Duration', '2')
        Assert-Equal ('設定ファイルの誤り: ' + ($case[0] -join ' / ')) ('{0}/{1}/{2}' -f $r.Code, $r.Log.Contains($case[1]), $r.Log.Contains('ファイル: ' + $bad)) '1/True/True'
    }
    New-ConfigFile (Join-Path $cfgDir 'typo.ini.txt') @('Duration = 2') | Out-Null
    $r = Invoke-Plain @('-ConfigFile', (Join-Path $cfgDir 'typo.ini'))
    Assert-Equal '設定ファイルがない（.txt のヒント）' ('{0}/{1}/{2}' -f $r.Code, $r.Log.Contains('設定ファイルが見つかりません: ' + (Join-Path $cfgDir 'typo.ini')), $r.Log.Contains('拡張子が .txt になっていませんか')) '1/True/True'

    # NoSave と OutputFolder（レビュー 中 3・2 回目 軽微 3）
    $cfgOut = New-ConfigFile (Join-Path $cfgDir 'out-only.ini') @('Interval = 1', 'OutputFolder = 設定の保存先')
    $r = Invoke-Plain @('-ConfigFile', $cfgOut, '-NoSave', '-Duration', '2')
    Assert-Equal '設定ファイルの OutputFolder ＋ コマンドラインの -NoSave: 保存しない' ('{0}/{1}' -f $r.Code, (Test-Path -LiteralPath (Join-Path $cfgDir '設定の保存先'))) '0/False'
    $cfgNoSave = New-ConfigFile (Join-Path $cfgDir 'nosave.ini') @('Interval = 1', 'NoSave = true')
    $lineOut = Join-Path $workRoot 'out\line-out'
    $r = Invoke-Plain @('-ConfigFile', $cfgNoSave, '-OutputFolder', $lineOut, '-Duration', '2')
    Assert-Equal '設定ファイルの NoSave = true ＋ コマンドラインの -OutputFolder: 保存する' ('{0}/{1}' -f $r.Code, [bool](Get-RunDir $lineOut)) '0/True'
    $cfgBoth = New-ConfigFile (Join-Path $cfgDir 'both.ini') @('# x', 'OutputFolder = a', '# y', 'NoSave = true')
    $r = Invoke-Plain @('-ConfigFile', $cfgBoth, '-Duration', '2')
    Assert-Equal '設定ファイルに両方: 行番号を小さい順に出して止まる' ('{0}/{1}' -f $r.Code, $r.Log.Contains('設定ファイルの 2 行目と 4 行目: NoSave = true のときは、OutputFolder は書けません')) '1/True'
    $r = Invoke-Plain @('-NoSave', '-OutputFolder', (Join-Path $workRoot 'out\both-line'), '-Duration', '2')
    Assert-Equal 'コマンドラインに両方: 止まる' ('{0}/{1}' -f $r.Code, $r.Log.Contains('-NoSave のときは、-OutputFolder は指定できません')) '1/True'

    #endregion

    #region 保存先

    Write-Section '保存先'
    $run = Invoke-Tool '日本語 と 空白' $quick
    Assert-Equal '日本語と空白を含む保存先' ('{0}/{1}' -f $run.Code, (Get-ResultFileNames $run.Dir)) '0/disks.csv,network.csv,process-list.csv,processes.csv,summary.txt,system.csv'
    $longOut = Join-Path $workRoot ('out\long\' + ('L' * 120) + '\' + ('M' * 120))
    $r = Invoke-Plain ($quick + @('-OutputFolder', $longOut))
    $longRun = @()
    try { $longRun = @([IO.Directory]::GetDirectories(('\\?\' + $longOut), 'Performance_*')) } catch { }
    $longFiles = if ($longRun.Count -gt 0) { @([IO.Directory]::GetFiles($longRun[0])).Count } else { 0 }
    Assert-Equal '260 文字を超える保存先' ('{0}/{1}/{2}' -f $r.Code, ($longOut.Length -gt 260), $longFiles) '0/True/6'

    $lockedOut = Join-Path $workRoot 'out\denied'
    [void][IO.Directory]::CreateDirectory($lockedOut)
    $null = & icacls.exe $lockedOut /deny "*${userSid}:(WD,AD)" 2>&1
    $denied.Add($lockedOut)
    $r = Invoke-Plain ($quick + @('-OutputFolder', $lockedOut))
    Assert-Equal '書き込めない保存先: 記録を始める前に止まる' ('{0}/{1}/{2}/{3}' -f $r.Code, $r.Log.Contains('保存先に書き込めません: ' + $lockedOut), $r.Log.Contains('-NoSave を付けると、画面に表示だけします'), $r.Log.Contains('記録を終えました')) '1/True/True/False'
    $letter = Get-FreeLetter
    if ($letter) {
        $r = Invoke-Plain ($quick + @('-OutputFolder', ('{0}:\perf' -f $letter)))
        Assert-Equal 'ないドライブの保存先' ('{0}/{1}' -f $r.Code, $r.Log.Contains(('保存先に書き込めません: {0}:\perf' -f $letter))) '1/True'
    } else {
        Write-Skip 'ないドライブの保存先' '空いているドライブ文字がないため'
    }

    # 記録の途中で CSV を書き込み禁止で開く（Excel で開いたときと同じ）と、次のファイルに分けて続ける
    $lockOut = Join-Path $workRoot 'out\lock'
    $stdout = Join-Path $workRoot 'lock-stdout.txt'
    $proc = Start-Process -FilePath (Get-Command $Exe).Source -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $fast), '-Interval', '1', '-Duration', '8', '-OutputFolder', ('"{0}"' -f $lockOut))
    $null = $proc.Handle
    $lockDir = Wait-SystemRows $lockOut 3 $proc
    if ($lockDir) {
        $held = [IO.File]::Open((Join-Path $lockDir 'system.csv'), 'Open', 'Read', 'Read')
        Start-Sleep -Milliseconds 2500
        $held.Dispose()
    }
    if (-not $proc.WaitForExit(60000)) { $proc.Kill() }
    $x = Read-Result $lockDir
    $n = Get-RecordedCount $x.Text
    $part2 = if ($lockDir -and (Test-Path -LiteralPath (Join-Path $lockDir 'system_2.csv'))) { @(Import-Csv -LiteralPath (Join-Path $lockDir 'system_2.csv') -Encoding UTF8).Count } else { -1 }
    Assert-Equal 'CSV を開かれたら system_2.csv に分ける（取りこぼしなし）' ('{0}/{1}/{2}' -f $proc.ExitCode, ($part2 -gt 0), ($x.System.Count + $part2 -eq $n)) '0/True/True'
    Assert-Equal 'CSV を分けた注意' $x.Text.Contains('※ system_2.csv に分けました（system.csv を開いているほかのプログラムがあり、書き足せなかったため）') 'True'

    # 1 つの CSV の行が上限に達したら分ける（process-list.csv も。レビュー 2 回目 軽微 4）
    $split = New-PatchedTool 'split' @('$MaxCsvRows     = 1000000', '$MaxCsvRows     = 2')
    # 3 回目で分かれる。VM で回が減っても 3 回以上記録できるよう、5 秒記録する（FGTEST-FS の 5.1 で、3 秒では 2 回のことがあった）
    $run = Invoke-Tool 'split' @('-Interval', '1', '-Duration', '5') $split
    $x = Read-Result $run.Dir
    $n = Get-RecordedCount $x.Text
    $systemParts = @(Get-ChildItem -LiteralPath $run.Dir -Filter 'system*.csv')
    $systemRows = 0
    foreach ($p in $systemParts) { $systemRows += @(Import-Csv -LiteralPath $p.FullName -Encoding UTF8).Count }
    $listParts = @(Get-ChildItem -LiteralPath $run.Dir -Filter 'process-list*.csv')
    Assert-Equal '上限で分ける: system.csv の行の合計・どのファイルも見出しと BOM' ('{0}/{1}/{2}' -f ($systemParts.Count -ge 2), ($systemRows -eq $n), @($systemParts + $listParts | Where-Object { -not (Test-BomOnce $_.FullName) -or [IO.File]::ReadAllLines($_.FullName)[0] -notlike '"*' }).Count) 'True/True/0'
    Assert-Equal '上限で分ける: 注意（CSV と process-list.csv）' ('{0}/{1}/{2}' -f $x.Text.Contains('※ system_2.csv に分けました（行が 2 行に達したため）'), $x.Text.Contains('※ process-list.csv は、2 件ごとに process-list_2.csv などに分けました'), ($listParts.Count -gt 2)) 'True/True/True'

    # 記録の途中で書き込めなくなったとき（ディスクがいっぱいなど）
    $fail = New-PatchedTool 'savefail' @('        $text = ($Lines -join "`r`n") + "`r`n"', '        $text = ($Lines -join "`r`n") + "`r`n"; if ($Key -eq ''disks'' -and $file.Rows -gt 0) { throw (New-Object IO.IOException ''テストの書き込みの失敗'') }')
    $run = Invoke-Tool 'savefail' $quick $fail
    $x = Read-Result $run.Dir
    Assert-Equal '途中で保存できなくなった: 画面の表示だけで続け、終了コード 1' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('記録を保存できなくなりました。画面の表示だけで続けます: テストの書き込みの失敗'), [regex]::IsMatch($run.Log, '※ \d\d:\d\d:\d\d から記録を保存できていません: テストの書き込みの失敗'), $run.Log.Contains('記録を終えました（指定の時間）')) '1/True/True/True'
    Assert-Equal '途中で保存できなくなった: summary.txt は保存しない' (Test-Path -LiteralPath (Join-Path $run.Dir 'summary.txt')) 'False'

    #endregion

    #region 想定外のエラー・カウンターが読めない

    Write-Section '想定外のエラー・カウンターが読めない'
    # 目印は 2 行で特定する（まとめの側の「# 追加のカウンター（-Counter）」に当たらないように）
    $loopAnchor = '            # 追加のカウンター' + "`n" + '            if ($extraHandles.Count -gt 0) {'
    $loopErr = New-PatchedTool 'loop-error' @($loopAnchor, ('            if ($recorded -eq 2) { throw ''テストの想定外のエラー'' }' + "`n" + $loopAnchor))
    $run = Invoke-Tool 'loop-error' $quick $loopErr
    $x = Read-Result $run.Dir
    Assert-Equal '記録の途中の例外: まとめを保存して終了コード 1' ('{0}/{1}/{2}/{3}' -f $run.Code, $x.Text.Contains('記録を終えました（エラー）'), $x.Text.Contains('※ エラーで止まりました: テストの想定外のエラー'), (Get-RecordedCount $x.Text)) '1/True/True/2'

    # まとめの途中の例外でも、Ctrl+C の扱いを戻す行まで届く（レビュー 2 回目 軽微 1）
    $marker = Join-Path $workRoot 'restored.txt'
    # @( ) の中の + は、カンマより後に結び付くので、括弧で囲む
    $restoreAnchor = '        if ($keyInput) {' + "`n" + '            try { while'
    $sumErr = New-PatchedTool 'summary-error' @(
        '        # 注意', ('        throw ''テストのまとめの例外''' + "`n" + '        # 注意'),
        $restoreAnchor, (('        [IO.File]::WriteAllText(''{0}'', ''x'')' -f $marker) + "`n" + $restoreAnchor)
    )
    $r = Invoke-Plain ($quick + @('-NoSave')) $sumErr
    Assert-Equal 'まとめの途中の例外: Ctrl+C の扱いを戻す行まで届く' ('{0}/{1}' -f ($r.Code -ne 0), (Test-Path -LiteralPath $marker)) 'True/True'

    $core = New-PatchedTool 'core' @("CpuTotal    = '\Processor Information(_Total)\% Processor Time'", "CpuTotal    = '\Processor Information(_Total)\% No Such Counter'")
    $r = Invoke-Plain ($quick + @('-OutputFolder', (Join-Path $workRoot 'out\core'))) $core
    Assert-Equal 'CPU のカウンターが読めない: 止まる' ('{0}/{1}/{2}/{3}' -f $r.Code, $r.Log.Contains('性能カウンターを読めません: \Processor Information(_Total)\% No Such Counter（'), $r.Log.Contains('性能カウンターが無効になっているか、壊れています'), [bool](Get-RunDir (Join-Path $workRoot 'out\core'))) '1/True/True/False'
    $disk = New-PatchedTool 'disk' @("Idle     = '\LogicalDisk(*)\% Idle Time'", "Idle     = '\LogicalDisk(*)\% No Such Counter'")
    $run = Invoke-Tool 'disk' $quick $disk
    $x = Read-Result $run.Dir
    Assert-Equal 'ディスクのカウンターが読めない: 飛ばして続ける' ('{0}/{1}/{2}/{3}/{4}' -f $run.Code, $x.Text.Contains('ディスク: 記録しない'), $x.Text.Contains('※ ディスクのカウンターを読めないため、記録しません（\LogicalDisk(*)\% No Such Counter: '), $x.Disks.Count, $x.Text.Contains('■ ディスク')) '0/True/True/0/False'
    $addType = New-PatchedTool 'addtype' @('namespace Workbench.MeasurePerformance', 'namespace Workbench.MeasurePerformance }}} broken')
    $r = Invoke-Plain ($quick + @('-NoSave')) $addType
    Assert-Equal 'Windows の API を呼ぶ準備ができない' ('{0}/{1}' -f $r.Code, $r.Log.Contains('Windows の API を呼ぶ準備ができません。コンパイラーの実行が禁止されている環境では使えません（')) '1/True'

    #endregion

    #region 目安

    Write-Section '目安（ディスクの応答時間と CPU の値を差し替えたコピー）'
    # MP_TEST_DISK は「応答時間(ms)|回数(回/秒)」で、システムのドライブの値を差し替える。MP_TEST_CPUHIT は CPU の目安を超えたことにする回
    $limitTool = New-PatchedTool 'limit' @(
        '                    $latency = if ($null -ne $d.Latency) { $d.Latency * 1000 } else { $null }',
        ('                    $latency = if ($null -ne $d.Latency) { $d.Latency * 1000 } else { $null }' + "`n" + '                    if ($env:MP_TEST_DISK -and $name -eq $env:SystemDrive) { $parts = $env:MP_TEST_DISK.Split(''|''); $latency = [double]$parts[0]; $d.Iops = [double]$parts[1] }'),
        '            $cpuHit    = $null -ne $cpuTotal -and $cpuTotal -ge $CpuLimit',
        ('            $cpuHit    = $null -ne $cpuTotal -and $cpuTotal -ge $CpuLimit' + "`n" + '            if ($env:MP_TEST_CPUHIT) { $cpuHit = @($env:MP_TEST_CPUHIT.Split('','')) -contains [string]$recorded }')
    )
    $drive = $env:SystemDrive
    $limitCases = @(
        # 遅くて回数の少ないディスク（132.52 ms・8.5 回/秒。レビュー 中 2）は判定する。CPU は 1 回だけ
        @('132.52|8.5', '1', $true, ('※ CPU 使用率 90% 以上: 1 回（')),
        # 読み書きのほとんどないディスクの 1 回の遅い読み書き（68 ms・0.5 回/秒）は判定しない。CPU は 2 回で続けてはいない
        @('68.19|0.5', '1,3', $false, '※ CPU 使用率 90% 以上: 2 回（続けて超えたことはありません）'),
        # -Interval 1 で 1 回だけの 120 ms は判定しない（レビュー 2 回目 軽微 2）。CPU は 3 回続けて
        @('120|1', '1,2,3', $false, '※ CPU 使用率 90% 以上: 3 回（最も長く続いたのは '),
        # 2 回・60 ms（回数 × 応答時間 0.12）は判定する
        @('60|2', '', $true, ''),
        # 3 回・25 ms（回数 × 応答時間 0.075。1 秒のうち待っていたのは 0.1 秒未満）は判定しない
        @('25|3', '', $false, '')
    )
    $k = 0
    foreach ($case in $limitCases) {
        $k++
        $env:MP_TEST_DISK = $case[0]
        $env:MP_TEST_CPUHIT = $case[1]
        $run = Invoke-Tool ('limit{0}' -f $k) $quick $limitTool
        $x = Read-Result $run.Dir
        $diskLine = ('※ {0} の応答時間 20 ms 以上: ' -f $drive)
        Assert-Equal ('ディスクの応答時間 {0} ms・{1} 回/秒（まとめ）' -f $case[0].Split('|')) $x.Text.Contains($diskLine) $case[2]
        # 記録中の行の * は、本物のほかのドライブが目安を超えても付くので、「付くはず」のときだけ確かめる
        if ($case[2]) { Assert-Equal ('ディスクの応答時間 {0} ms・{1} 回/秒（記録中の行の *）' -f $case[0].Split('|')) $run.Log.Contains(' ms * ') 'True' }
        if ($case[3]) {
            Assert-Equal ('CPU の目安の表示（{0} 回目）' -f $case[1]) $x.Text.Contains($case[3]) 'True'
            Assert-Equal ('CPU の目安の記録中の行の *（{0} 回目）' -f $case[1]) ([regex]::IsMatch($run.Log, '\d\d:\d\d:\d\d +\d+% \* ')) 'True'
        }
    }
    $env:MP_TEST_DISK = $null
    $env:MP_TEST_CPUHIT = $null

    #endregion

    #region 地域設定

    Write-Section '地域設定（de-DE でも小数点は . ・3 桁区切りは ,）'
    $cultureOut = Join-Path $workRoot 'out\culture'
    $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'de-DE'; & '{0}' -Interval 1 -Duration 3 -OutputFolder '{1}' | Out-Null" -f $fast, $cultureOut
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
    $x = Read-Result (Get-RunDir $cultureOut)
    Assert-Equal 'de-DE: CSV の数値' ('{0}/{1}' -f ($x.System.Count -gt 0), @($x.System | Where-Object { $_.'空きメモリ(MB)' -notmatch '^\d{1,3}(,\d{3})*\.\d$' -or $_.'日時' -notmatch '^\d{4}/\d\d/\d\d \d\d:\d\d:\d\d$' }).Count) 'True/0'
    Assert-Equal 'de-DE: まとめの数値' ('{0}/{1}' -f [regex]::IsMatch($x.Text, 'メモリ [\d,]+\.\d GB'), [regex]::IsMatch($x.Text, 'このツールの CPU 使用率: 平均 \d+\.\d%')) 'True/True'

    #endregion

    #region 引数の誤り

    Write-Section '引数の誤り（記録を始める前に止まる）'
    $r = Invoke-Plain @('-Interval', '120', '-Duration', '1') $tool
    Assert-Equal '-Duration が -Interval より短い' ('{0}/{1}' -f $r.Code, $r.Log.Contains('-Duration（1 分）が -Interval（120 秒）より短いため、1 回も記録できません')) '1/True'
    $r = Invoke-Plain @('-Interval', '1') -RedirectInput
    Assert-Equal '入力をリダイレクトして -Duration なし' ('{0}/{1}' -f $r.Code, $r.Log.Contains('Ctrl+C で止められない起動のしかたです。-Duration で記録する時間を指定してください')) '1/True'
    $r = Invoke-Plain @('-Duration', '2', '-Foo', 'bar')
    Assert-Equal 'どのパラメーターにも当たらない引数' ('{0}/{1}' -f $r.Code, $r.Log.Contains('どのパラメーターにも当たらない引数があります: -Foo bar')) '1/True'
    $r = Invoke-Plain @('-Duration', '2', '-Top', '51')
    Assert-Equal '-Top 51（範囲外）' $r.Code 1

    $clm = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '{0}' -Interval 1 -Duration 2 -NoSave" -f $fast
    $clmLog = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -Command $clm 2>&1) | ForEach-Object { "$_" }) -join "`n")
    Assert-Equal '制約付き言語モード' $clmLog.Contains('制約付き言語モードでは使えません（AppLocker や WDAC でスクリプトの実行が制限されている環境です）') 'True'

    #endregion

    #region コマンドプロンプト

    Write-Section 'コマンドプロンプトから起動'
    $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -Interval 1' -f (Get-Command $Exe).Source, $fast
    $cmdOut = Join-Path $workRoot 'out\cmd'
    $r = Invoke-ViaCmd ('{0} -Duration 2 -OutputFolder "{1}\"' -f $cmdBase, $cmdOut)
    Assert-Equal '最後の引数が "保存先\"' ('{0}/{1}' -f $r.Code, [bool](Get-RunDir $cmdOut)) '0/True'
    $cmdOut2 = Join-Path $workRoot 'out\cmd2'
    $r = Invoke-ViaCmd ('{0} -OutputFolder "{1}\" -Duration 2' -f $cmdBase, $cmdOut2)
    Assert-Equal '"保存先\" の後ろに引数: ヒントを出して止まる' ('{0}/{1}/{2}/{3}' -f $r.Code, $r.Log.Contains('-OutputFolder の値に引用符（"）が残っています'), $r.Log.Contains('最後に \ を付けないでください'), (Test-Path -LiteralPath $cmdOut2)) '1/True/True/False'
    $r = Invoke-ViaCmd ('{0} -OutputFolder "{1}\" -Process "a b" -Duration 2' -f $cmdBase, $cmdOut2)
    Assert-Equal '"保存先\" の後ろに空白を含む値: 引用符のヒント' ('{0}/{1}' -f $r.Code, $r.Log.Contains('-OutputFolder の値に引用符（"）が残っています')) '1/True'
    $r = Invoke-ViaCmd ('{0} -Intervl 1 -Duration 2 -NoSave' -f $cmdBase)
    Assert-Equal 'パラメーターの名前の書き間違い' ('{0}/{1}' -f $r.Code, $r.Log.Contains('どのパラメーターにも当たらない引数があります: -Intervl 1')) '1/True'
    $r = Invoke-ViaCmd ('{0} -Duration 2 -NoSave -NoSuchName "a b"' -f $cmdBase)
    Assert-Equal '名前の合わない引数と空白を含む値' ('{0}/{1}' -f $r.Code, $r.Log.Contains('どのパラメーターにも当たらない引数があります: -NoSuchName "a b"')) '1/True'
    $r = Invoke-ViaCmd ('{0} -Duration 2 -NoSave -Counter "\Processor Information(_Total)\%% Processor Time"' -f $cmdBase)
    Assert-Equal 'バッチファイルの %% は % になる（README の注意）' ('{0}/{1}' -f $r.Code, $r.Log.Contains('追加のカウンター: \Processor Information(_Total)\% Processor Time')) '0/True'

    #endregion

    #region Ctrl+C

    Write-Section 'Ctrl+C'
    $keyScript = New-HelperScript 'send-ctrlc.ps1' @(
        'param([int]$TargetPid)'
        'Add-Type -TypeDefinition @'''
        'using System;'
        'using System.Runtime.InteropServices;'
        'public static class WorkbenchConsoleKeys {'
        '    [DllImport("kernel32.dll", SetLastError = true)] static extern bool FreeConsole();'
        '    [DllImport("kernel32.dll", SetLastError = true)] static extern bool AttachConsole(uint pid);'
        '    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)] static extern IntPtr CreateFileW(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);'
        '    [DllImport("kernel32.dll", SetLastError = true)] static extern bool WriteConsoleInputW(IntPtr handle, INPUT_RECORD[] records, uint length, out uint written);'
        '    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);'
        '    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] struct KEY_EVENT_RECORD { public int KeyDown; public ushort RepeatCount; public ushort VirtualKeyCode; public ushort VirtualScanCode; public char UnicodeChar; public uint ControlKeyState; }'
        '    [StructLayout(LayoutKind.Explicit)] struct INPUT_RECORD { [FieldOffset(0)] public ushort EventType; [FieldOffset(4)] public KEY_EVENT_RECORD KeyEvent; }'
        '    public static string SendCtrlC(uint pid) {'
        '        FreeConsole();'
        '        if (!AttachConsole(pid)) { return "AttachConsole " + Marshal.GetLastWin32Error(); }'
        '        IntPtr input = CreateFileW("CONIN$", 0xC0000000, 3, IntPtr.Zero, 3, 0, IntPtr.Zero);'
        '        if (input == new IntPtr(-1)) { FreeConsole(); return "CONIN$ " + Marshal.GetLastWin32Error(); }'
        '        INPUT_RECORD[] records = new INPUT_RECORD[2];'
        '        for (int i = 0; i < 2; i++) {'
        '            records[i].EventType = 1; records[i].KeyEvent.KeyDown = i == 0 ? 1 : 0; records[i].KeyEvent.RepeatCount = 1;'
        '            records[i].KeyEvent.VirtualKeyCode = 0x43; records[i].KeyEvent.VirtualScanCode = 0x2E; records[i].KeyEvent.UnicodeChar = (char)3; records[i].KeyEvent.ControlKeyState = 0x0008;'
        '        }'
        '        uint written;'
        '        bool ok = WriteConsoleInputW(input, records, 2, out written);'
        '        int error = Marshal.GetLastWin32Error();'
        '        CloseHandle(input);'
        '        FreeConsole();'
        '        return ok ? "ok" : "WriteConsoleInput " + error;'
        '    }'
        '}'
        '''@'
        '''RESULT='' + [WorkbenchConsoleKeys]::SendCtrlC([uint32]$TargetPid)'
    )
    # -Duration なし（Ctrl+C で止めるまで）。隠したコンソールで動かし、記録が始まったら Ctrl+C のキーを送る
    $ctrlOut = Join-Path $workRoot 'out\ctrlc'
    $proc = Start-Process -FilePath (Get-Command $Exe).Source -WindowStyle Hidden -PassThru -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $fast), '-Interval', '1', '-OutputFolder', ('"{0}"' -f $ctrlOut))
    $null = $proc.Handle
    $ctrlDir = Wait-SystemRows $ctrlOut 3 $proc
    $sent = '(記録が始まらなかった)'
    if ($ctrlDir -and -not $proc.HasExited) { $sent = (Invoke-Helper $keyScript @('-TargetPid', $proc.Id))['RESULT'] }
    if (-not $proc.WaitForExit(60000)) { $proc.Kill(); $sent += '（止まらなかった）' }
    if ($sent -like 'AttachConsole*' -or $sent -like 'CONIN$*') {
        Write-Skip 'Ctrl+C で止める' ('キーを送れなかったため（{0}）' -f $sent)
    } else {
        $x = Read-Result $ctrlDir
        Assert-Equal 'Ctrl+C: 終了コード 0・まとめと process-list.csv を保存' ('{0}/{1}/{2}/{3}' -f $sent, $proc.ExitCode, $x.Text.Contains('記録を終えました（Ctrl+C）'), ($x.List.Count -gt 0)) 'ok/0/True/True'
    }

    # 入力をリダイレクトしているとき（パイプラインの停止）も、まとめを保存する。Stop() は取りこぼすことがあるので 3 回まで試す
    $stopScript = New-HelperScript 'stop.ps1' @(
        'param([string]$Tool, [string]$Out)'
        '$ps = [powershell]::Create()'
        '[void]$ps.AddScript("& ''$Tool'' -Interval 1 -Duration 60 -OutputFolder ''$Out''")'
        '$handle = $ps.BeginInvoke()'
        '$watch = [Diagnostics.Stopwatch]::StartNew()'
        'while ($watch.Elapsed.TotalSeconds -lt 40 -and -not $handle.IsCompleted) {'
        '    $dir = @(Get-ChildItem -LiteralPath $Out -Directory -ErrorAction SilentlyContinue)'
        '    if ($dir.Count -gt 0 -and (Test-Path -LiteralPath (Join-Path $dir[0].FullName ''system.csv''))) {'
        '        $lines = 0; try { $lines = @(Get-Content -LiteralPath (Join-Path $dir[0].FullName ''system.csv'')).Count } catch { }'
        '        if ($lines -ge 3) { break }'
        '    }'
        '    Start-Sleep -Milliseconds 100'
        '}'
        '$ps.Stop()'
        '''STATE='' + $ps.InvocationStateInfo.State'
    )
    $stopped = $false
    for ($try = 1; $try -le 3 -and -not $stopped; $try++) {
        $stopOut = Join-Path $workRoot ('out\stop{0}' -f $try)
        [void][IO.Directory]::CreateDirectory($stopOut)
        $result = Invoke-Helper $stopScript @('-Tool', $fast, '-Out', $stopOut)
        $x = Read-Result (Get-RunDir $stopOut)
        if ($result['STATE'] -eq 'Stopped' -and $x.Text.Contains('記録を終えました（Ctrl+C）')) { $stopped = $true }
    }
    Assert-Equal 'パイプラインの停止: まとめと process-list.csv を保存（3 回まで試す）' ('{0}/{1}' -f $stopped, ($x.List.Count -gt 0)) 'True/True'

    #endregion

    #region C# と PowerShell の CSV の決まり・まとめの速さ

    Write-Section 'C# の CSV の決まりがスクリプトと同じ・10 万件のまとめが速い（レビュー 中 4）'
    $parityScript = New-HelperScript 'parity.ps1' @(
        'param([string]$Tool)'
        '$ast = [Management.Automation.Language.Parser]::ParseFile($Tool, [ref]$null, [ref]$null)'
        '$Invariant = [Globalization.CultureInfo]::InvariantCulture'
        '$src = $ast.FindAll({ $args[0] -is [Management.Automation.Language.AssignmentStatementAst] -and $args[0].Left.Extent.Text -eq ''$nativeSource'' }, $true) | Select-Object -First 1'
        '. ([scriptblock]::Create($src.Extent.Text))'
        'Add-Type -TypeDefinition $nativeSource -IgnoreWarnings -WarningAction SilentlyContinue'
        'foreach ($a in $ast.FindAll({ $args[0] -is [Management.Automation.Language.AssignmentStatementAst] -and $args[0].Left.Extent.Text -eq ''$excelRisk'' }, $true)) { . ([scriptblock]::Create($a.Extent.Text)) }'
        'foreach ($f in $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) { if ($f.Name -in ''Protect-ExcelText'', ''ConvertTo-CellText'', ''Format-Number'', ''ConvertTo-CsvLine'') { . ([scriptblock]::Create($f.Extent.Text)) } }'
        '$texts = @(''001'', ''2026-10-31'', ''=SUM(1)'', ''+81'', ''-x'', ''@a'', ''.1'', ''1e5'', ([string][char]0xFF11 + [char]0xFF12), ''abc'', ''50%'', (''='' + (''a'' * 300)), "a`r`nb", ("x" + [char]7 + "y"), ''  pad  '', '''', ''svchost.exe'', ''0,1'', ([string][char]0xD800 + ''z''), ''a"b'')'
        '$diff = 0'
        'foreach ($t in $texts) { if (-not [string]::Equals((Protect-ExcelText (ConvertTo-CellText $t)), [Workbench.MeasurePerformance.CsvText]::Excel([Workbench.MeasurePerformance.CsvText]::Cell($t)), [StringComparison]::Ordinal)) { $diff++ } }'
        'foreach ($v in @(-0.04, -0.0, 0.0, 12.345, -4.06, 1234567.89, -1234.5)) { foreach ($fmt in ''N1'', ''N0'', ''#,##0.###'') { if ((Format-Number $v $fmt) -ne [Workbench.MeasurePerformance.CsvText]::Number($v, $fmt)) { $diff++ } } }'
        '$fields = [string[]]@(''a'', ''b"c'', '''', ''1,2'')'
        'if ((ConvertTo-CsvLine $fields) -ne [Workbench.MeasurePerformance.CsvText]::Line($fields)) { $diff++ }'
        '''DIFF='' + $diff'
        '$sampler = New-Object Workbench.MeasurePerformance.ProcessSampler (0, 4, [Management.Automation.WildcardPattern[]]@(), [bool[]]@())'
        '$now = Get-Date'
        'for ($i = 0; $i -lt 100000; $i++) { $st = New-Object Workbench.MeasurePerformance.ProcStat; $st.Name = ''app.exe''; $st.Pid = $i; $st.CpuSum = $i % 97; $st.MemoryMax = $i % 89; $st.CommandLine = ''app.exe --id='' + $i; $st.FirstSeen = $now; $st.LastSeen = $now; $sampler.StatList.Add($st) }'
        '$sw = [Diagnostics.Stopwatch]::StartNew()'
        '$byCpu = $sampler.SortedByCpu(); $null = $sampler.SortedByMemory(); $lines = [Workbench.MeasurePerformance.ProcessSampler]::ListLines($byCpu, 1000)'
        '''SECONDS='' + $sw.Elapsed.TotalSeconds.ToString(''0.0'', $Invariant)'
        '''LINES='' + $lines.Count'
    )
    $result = Invoke-Helper $parityScript @('-Tool', $repoTool)
    Assert-Equal 'C# の CsvText とスクリプトの関数の結果が同じ' $result['DIFF'] '0'
    Assert-Equal '10 万件の並べ替えと process-list.csv の行づくりが 15 秒未満' ('{0}/{1}' -f ([double]::Parse("0$($result['SECONDS'])", $inv) -lt 15), $result['LINES']) 'True/100000'

    #endregion

    #region 画面によって幅の変わる文字

    # ※ と … は、Windows PowerShell のウィンドウでは 2 文字分、Windows ターミナルでは 1 文字分で描かれ、後ろの列がずれる（実機確認で見つけた）
    Write-Section '記録中の行と表に、画面によって幅の変わる文字（※ …）を使わない'
    $widthScript = New-HelperScript 'width.ps1' @(
        'param([string]$Tool)'
        '$ast = [Management.Automation.Language.Parser]::ParseFile($Tool, [ref]$null, [ref]$null)'
        'foreach ($a in $ast.FindAll({ $args[0] -is [Management.Automation.Language.AssignmentStatementAst] -and $args[0].Left.Extent.Text -eq ''$wideChar'' }, $true)) { . ([scriptblock]::Create($a.Extent.Text)) }'
        'foreach ($f in $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) { if ($f.Name -in ''Get-DisplayWidth'', ''Limit-Width'', ''Get-Mark'') { . ([scriptblock]::Create($f.Extent.Text)) } }'
        '''MARK=['' + (Get-Mark $true) + ''][''+ (Get-Mark $false) + '']'''
        '''ASCII='' + (Limit-Width ''AsusFanControlService.exe'' 20)'
        '$wide = Limit-Width ''あいうえおかきくけこさしすせそ'' 20'
        '''WIDE='' + $wide.Length + ''/'' + (Get-DisplayWidth $wide) + ''/'' + $wide.EndsWith(''...'', [StringComparison]::Ordinal)'
    )
    $result = Invoke-Helper $widthScript @('-Tool', $repoTool)
    Assert-Equal '目安の印（超えた・超えていない）は、どちらも半角 3 文字' $result['MARK'] '[ * ][   ]'
    Assert-Equal '長い名前は ... で切る（半角 20 文字分）' $result['ASCII'] 'AsusFanControlSer...'
    Assert-Equal '全角の長い名前も ... で切り、20 文字分を超えない（文字数/幅/... で終わる）' $result['WIDE'] '11/19/True'

    #endregion

    #region 32 ビットの PowerShell

    Write-Section '32 ビットの Windows PowerShell（構造体の読み取り）'
    $wow = Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if ($Exe -like 'powershell*' -and (Test-Path -LiteralPath $wow)) {
        $wowOut = Join-Path $workRoot 'out\wow64'
        $wowLog = ConvertTo-FlatText ((@(& $wow -NoProfile -ExecutionPolicy Bypass -File $fast -Interval 1 -Duration 3 -Top 0 -Process 'mp-test-ping' -Counter '\Memory\Available MBytes' -OutputFolder $wowOut 2>&1) | ForEach-Object { "$_" }) -join "`n")
        $x = Read-Result (Get-RunDir $wowOut)
        $item = @($x.List | Where-Object { $_.'PID' -eq $dummyPid })[0]
        $procNow = Get-Process -Id $dummyProc.Id
        Assert-Equal '32 ビット: プロセスの値と、カウンターの値' ('{0}/{1}/{2}' -f ($null -ne $item -and [math]::Abs((ConvertTo-Number $item.'コミットの最大(MB)') - $procNow.PrivateMemorySize64 / 1MB) -lt 2), ($x.Counters.Count -gt 0 -and (ConvertTo-Number $x.Counters[0].'値') -gt 0), ($x.System.Count -gt 0)) 'True/True/True'
    } else {
        Write-Skip '32 ビットの Windows PowerShell' '5.1 のテストで、SysWOW64 の PowerShell があるときだけ試すため'
    }

    #endregion

    #region 再発防止の一覧

    Write-Section 'レビューで見つかった問題の再発防止（上で確かめている）'
    # 1 回目 中 1. -Counter の同じ名前のインスタンス: -Counter \Process(*)
    # 1 回目 中 2・2 回目 軽微 2. ディスクの応答時間の判定: 目安
    # 1 回目 中 3・2 回目 軽微 3. NoSave と OutputFolder: 設定ファイル
    # 1 回目 中 4. まとめの速さ: C# の CSV の決まり・まとめの速さ。2 回目の Ctrl+C を受けないのは実機確認
    # 1 回目 軽微 5. 時刻をそろえる: 基本
    # 1 回目 軽微 6・7. -NoSave の案内と注意: -NoSave（「止めるときは…」の行は、入力をリダイレクトすると出ないので実機確認）
    # 1 回目 軽微 8. 地域設定: 地域設定
    # 1 回目 軽微 9. 目安の表示: 目安
    # 2 回目 軽微 1. まとめの途中の例外: 想定外のエラー
    # 2 回目 軽微 4. process-list.csv の分割: 保存先
    # 開発中に見つけたもの: 7 の -0.0（基本）

    #endregion
} catch {
    '  FAIL  テストが途中で止まりました: {0}（{1} 行目）' -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber
    $script:failCount++
} finally {
    $env:MP_TEST_DISK = $null
    $env:MP_TEST_CPUHIT = $null
    $env:MP_TEST_OUT = $null
    if ($dummyProc -and -not $dummyProc.HasExited) { try { $dummyProc.Kill() } catch { } }
    if ($peakProc -and -not $peakProc.HasExited) { try { $peakProc.Kill() } catch { } }
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    if ($KeepTemp) {
        Write-Section ('ダミー環境を残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        try {
            Start-Sleep -Milliseconds 500   # ダミーのプロセスが終わるのを待つ
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('ダミー環境を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
}

'  END'
exit $script:failCount
