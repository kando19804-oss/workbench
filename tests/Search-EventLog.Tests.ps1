<#
.SYNOPSIS
    Search-EventLog の回帰テストです。

.DESCRIPTION
    テスト専用のイベントソース（workbench-tests-<日時>）を作って Application ログにイベントを書き込み、
    それを .evtx ファイルに書き出して、Search-EventLog で取り出した結果を期待値と照合します。
    .evtx は、名前に [ ] を含むファイル、ジャンクションのループを含むフォルダ、260 文字を超えるパスなどに置いて試します。
    テスト用のイベントソースと一時フォルダは、終了時に削除します（書き込んだイベントは Application ログに残ります）。

    イベントソースを作るには、管理者の権限が要ります。管理者でないときは、イベントを使うテストを飛ばします。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Search-EventLog を実行する PowerShell（powershell.exe または pwsh）。

.PARAMETER KeepTemp
    一時フォルダと結果を削除せずに残します（調査用）。

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

$tool     = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Search-EventLog\Search-EventLog.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$stamp    = '{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Search-EventLog_' + $stamp)
$source   = 'workbench-tests-' + $stamp
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

function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out  = Join-Path $workRoot "out_$Label"
    $log  = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1 | Out-String)
    $code = $LASTEXITCODE
    $dir  = $null
    if (Test-Path -LiteralPath $out) {
        $found = @(Get-ChildItem -LiteralPath $out -Directory -Filter 'EventLog_*')
        if ($found.Count -gt 0) { $dir = $found[0].FullName }
    }
    [pscustomobject]@{ Code = $code; Log = $log; Dir = $dir }
}

function Read-Result([string]$Dir) {
    function Read-CsvOrEmpty([string]$Name) {
        if (-not $Dir) { return @() }
        $file = Join-Path $Dir $Name
        if (-not (Test-Path -LiteralPath $file)) { return @() }
        @(Import-Csv -LiteralPath $file -Encoding UTF8)
    }
    [pscustomobject]@{ Events = @(Read-CsvOrEmpty 'events.csv'); Groups = @(Read-CsvOrEmpty 'summary.csv'); Skipped = @(Read-CsvOrEmpty 'skipped.csv') }
}

# テスト用のイベントだけの件数（ID ごと、ID の小さい順）
function Get-IdCounts($Rows) {
    $ids = @($Rows | Where-Object { $_.'ソース' -eq $source } | ForEach-Object { [int]$_.'イベントID' } | Sort-Object)
    ($ids | Group-Object | ForEach-Object { '{0}x{1}' -f $_.Name, $_.Count }) -join ','
}

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $log = ConvertTo-FlatText (& cmd.exe /d /c $batch 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
[void][IO.Directory]::CreateDirectory($workRoot)
$sourceCreated = $false
$junction = Join-Path $workRoot 'evtx\sub\loop'

try {
    Write-Section ('準備: テスト用のイベントを書き込む（ソース {0}）' -f $source)
    $ready = $false
    if ($isAdmin) {
        try {
            [Diagnostics.EventLog]::CreateEventSource($source, 'Application')
            $sourceCreated = $true
            $log = New-Object Diagnostics.EventLog 'Application'
            $log.Source = $source
            foreach ($n in 1..3) { $log.WriteEntry(('テスト エラー Timeout {0}' -f $n), 'Error', 1001) }
            foreach ($n in 1..2) { $log.WriteEntry(('テスト 警告 {0}' -f $n), 'Warning', 1002) }
            $log.WriteEntry('テスト 情報', 'Information', 1003)
            $log.WriteEntry('=1+1 "quoted"', 'Error', 1004)
            $log.WriteEntry("1 行目`r`n2 行目", 'Error', 1005)
            $log.Dispose()
            Start-Sleep -Seconds 2
            # テスト用のイベントだけを .evtx に書き出す
            $evtxDir = Join-Path $workRoot 'evtx'
            [void][IO.Directory]::CreateDirectory("$evtxDir\sub")
            $query = "*[System[Provider[@Name='$source']]]"
            $null = & wevtutil.exe epl Application "$evtxDir\all.evtx" "/q:$query" 2>&1
            # 名前に [ ] を含むファイルと、ワイルドカードとして一致してしまう隣のファイル（警告だけ）
            Copy-Item -LiteralPath "$evtxDir\all.evtx" -Destination "$evtxDir\sub\[a].evtx"
            $null = & wevtutil.exe epl Application "$evtxDir\sub\a.evtx" "/q:*[System[Provider[@Name='$source'] and Level=3]]" 2>&1
            $null = New-Item -ItemType Junction -Path $junction -Target $evtxDir
            $ready = Test-Path -LiteralPath "$evtxDir\all.evtx"
        } catch {
            Write-Skip 'テスト用のイベント' ('書き込めなかったため: {0}' -f $_.Exception.Message)
        }
    } else {
        Write-Skip 'テスト用のイベント' '管理者でないため（イベントソースを作れない）'
    }
    Assert-Equal 'テスト用のイベントを用意できた（管理者のときは必ず）' ($ready -or -not $isAdmin) 'True'
    $all = Join-Path $workRoot 'evtx\all.evtx'

    if ($ready) {
        Write-Section '.evtx を読む（既定のレベル: 重大・エラー）'
        $run = Invoke-Tool 'evtx' @('-Path', $all, '-Since', '1h')
        Assert-Equal '終了コード' $run.Code 0
        $x = Read-Result $run.Dir
        Assert-Equal '取り出した件数（ID ごと）' (Get-IdCounts $x.Events) '1001x3,1004x1,1005x1'
        Assert-Equal '集計の 1 行目' ('{0}/{1}/{2}' -f $x.Groups[0].'件数', $x.Groups[0].'イベントID', $x.Groups[0].'レベル') '3/1001/エラー'
        Assert-Equal '新しい順' (([datetime]$x.Events[0].'日時') -ge ([datetime]$x.Events[-1].'日時')) 'True'
        Assert-Equal 'メッセージは全文（改行も含む）' (@($x.Events | Where-Object { $_.'イベントID' -eq '1005' }) | ForEach-Object { $_.'メッセージ' -replace "`r`n", '|' }) '1 行目|2 行目'
        Assert-Equal '集計のメッセージは 1 行目' (@($x.Groups | Where-Object { $_.'イベントID' -eq '1005' }) | ForEach-Object { $_.'メッセージ（1 行目）' }) '1 行目'
        # レビューで見つかった問題: ="..." の中の " を重ねていなかった
        Assert-Equal 'Excel 対策（中の " を重ねる）' (@($x.Events | Where-Object { $_.'イベントID' -eq '1004' }) | ForEach-Object { $_.'メッセージ' }) '="=1+1 ""quoted"""'
        Assert-Equal '件数の行' ([regex]::Match($run.Log, 'イベント \S+ 件（[^）]*）  種類 \S+').Value) 'イベント 5 件（エラー 5）  種類 3'

        Write-Section 'レベル・ID・ソース・メッセージで絞る'
        $cases = @(
            @('-Level Warning', @('-Level', 'Warning'), '1002x2'),
            @('-Level 日本語', @('-Level', '警告,情報'), '1002x2,1003x1'),
            @('-Id', @('-Id', '1001,1003'), '1001x3'),
            @('-ExcludeId', @('-ExcludeId', '1001'), '1004x1,1005x1'),
            @('-Source', @('-Source', $source.ToUpperInvariant()), '1001x3,1004x1,1005x1'),
            @('-Pattern（大文字小文字を区別しない、| でどれか）', @('-Pattern', 'nomatch|TIMEOUT'), '1001x3')
        )
        foreach ($case in $cases) {
            $run = Invoke-Tool ('f' + [Array]::IndexOf($cases, $case)) (@('-Path', $all, '-Since', '1h') + $case[1])
            Assert-Equal $case[0] ('{0}/{1}' -f $run.Code, (Get-IdCounts (Read-Result $run.Dir).Events)) "0/$($case[2])"
        }
        $run = Invoke-Tool 'nosource' @('-Path', $all, '-Since', '1h', '-Source', 'NoSuchSource')
        Assert-Equal 'ソースがないとき: 該当なし' ('{0}/{1}' -f $run.Code, $run.Log.Contains('該当するイベントはありません')) '2/True'
        # レビューで見つかった問題: ID が多いと、検索の式の上限を超えてエラーになる
        $manyIds = (@(1001) + (2001..2015)) -join ','
        $run = Invoke-Tool 'manyids' @('-Path', $all, '-Since', '1h', '-Level', 'Critical,Error,Warning,Information,Verbose', '-Id', $manyIds)
        Assert-Equal 'ID 16 個（取り出した後に絞る）' ('{0}/{1}' -f $run.Code, (Get-IdCounts (Read-Result $run.Dir).Events)) '0/1001x3'
        $run = Invoke-Tool 'ids15' @('-Path', $all, '-Since', '1h', '-Level', 'Critical,Error,Warning,Information,Verbose', '-Id', ((@(1001) + (2001..2014)) -join ','))
        Assert-Equal 'ID 15 個とすべてのレベル（検索の式に入れる）' ('{0}/{1}' -f $run.Code, (Get-IdCounts (Read-Result $run.Dir).Events)) '0/1001x3'
        $run = Invoke-Tool 'old' @('-Path', $all, '-Since', '2000-01-01', '-Until', '2000-01-02')
        Assert-Equal '期間に該当なし' ('{0}/{1}' -f $run.Code, $run.Log.Contains('該当するイベントはありません')) '2/True'

        Write-Section 'ファイルとフォルダ'
        # レビューで見つかった問題: Get-WinEvent は Path をワイルドカードとして読むため、[a].evtx が隣の a.evtx に一致していた
        $run = Invoke-Tool 'bracket' @('-Path', (Join-Path $workRoot 'evtx\sub\[a].evtx'), '-Since', '1h')
        Assert-Equal '名前に [ ] を含むファイル' ('{0}/{1}' -f $run.Code, (Get-IdCounts (Read-Result $run.Dir).Events)) '0/1001x3,1004x1,1005x1'
        $run = Invoke-Tool 'folder' @('-Path', (Join-Path $workRoot 'evtx'), '-Since', '1h', '-Level', 'Warning')
        $x = Read-Result $run.Dir
        Assert-Equal 'フォルダの下の .evtx をすべて読む' (Get-IdCounts $x.Events) '1002x6'
        Assert-Equal 'ループするジャンクションをたどらない' ((@($x.Skipped | Where-Object { $_.'理由' -eq 'リンク' }) | ForEach-Object { $_.'対象' }) -join ',') $junction
        Assert-Equal '対象の行' $run.Log.Contains(('対象: {0} ほか 2 ファイル' -f $all)) 'True'
        $longDir = Join-Path $workRoot (('L' * 120) + '\' + ('M' * 120))
        [void][IO.Directory]::CreateDirectory('\\?\' + $longDir)
        [IO.File]::Copy($all, '\\?\' + (Join-Path $longDir 'long.evtx'))
        $run = Invoke-Tool 'long' @('-Path', $longDir, '-Since', '1h')
        $x = Read-Result $run.Dir
        $readOrReported = ($run.Code -eq 0 -and (Get-IdCounts $x.Events) -eq '1001x3,1004x1,1005x1') -or $run.Log.Contains('ファイルを読めない')
        Assert-Equal '260 文字を超えるパス: 読めるか、読めないと知らせる' $readOrReported 'True'
        $run = Invoke-Tool 'missing' @('-Path', ('{0},{1}' -f $all, (Join-Path $workRoot 'no.evtx')), '-Since', '1h')
        Assert-Equal 'ないファイルは飛ばしてほかを読む' ('{0}/{1}' -f $run.Code, ((@((Read-Result $run.Dir).Skipped) | ForEach-Object { $_.'理由' }) -join ',')) '0/ファイルを読めない'

        Write-Section 'この PC のイベントログ'
        $run = Invoke-Tool 'live' @('-LogName', 'NoSuchLog,Application', '-Source', $source, '-Since', '1h')
        $x = Read-Result $run.Dir
        Assert-Equal 'Application から読む' ('{0}/{1}' -f $run.Code, (Get-IdCounts $x.Events)) '0/1001x3,1004x1,1005x1'
        Assert-Equal 'ないログは飛ばす' ((@($x.Skipped) | ForEach-Object { '{0}/{1}' -f $_.'理由', $_.'対象' }) -join ',') 'ログがない/NoSuchLog'

        Write-Section 'コマンドプロンプトから起動'
        $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -NoSave -Since 1h -Level "Critical,Error" -Path "{2}\"' -f $Exe, $tool, (Join-Path $workRoot 'evtx\sub'))
        Assert-Equal 'カンマ区切りの -Level と、末尾の \' ('{0}/{1}' -f $run.Code, $run.Log.Contains('イベント 5 件')) '0/True'
    }

    Write-Section '止まるとき'
    $cases = @(
        @('ないログだけ', @('-LogName', 'NoSuchLog'), '読めるログがありませんでした'),
        @('レベルの誤り', @('-Level', 'Fatal'), '-Level は Critical'),
        @('期間の誤り', @('-Since', 'yesterday'), '-Since は "2026-10-31 09:00"'),
        @('期間が逆', @('-Since', '2h', '-Until', '3h'), 'と同じか、それより後になっています'),
        @('ID の誤り', @('-Id', '70000'), '-Id は 0～65535'),
        @('.evtx がない', @('-Path', (Join-Path $workRoot 'no\such.evtx')), '読める .evtx ファイルがありません')
    )
    foreach ($case in $cases) {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @($case[1]) -NoSave 2>&1 | Out-String)
        Assert-Equal $case[0] ('{0}/{1}' -f $LASTEXITCODE, $log.Contains($case[2])) '1/True'
    }
    # レビューで見つかった問題: Security ログの権限エラーが「その他のエラー」になり、管理者のヒントが出なかった
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -LogName 'Security,System' -Since 10m -NoSave 2>&1 | Out-String)
    if ($isAdmin) {
        Assert-Equal '管理者は Security を読める' $log.Contains('Security: 権限なし') 'False'
    } else {
        Assert-Equal '管理者でないときの Security は権限なし' ('{0}/{1}' -f $log.Contains('スキップ 1 件（権限なし）'), $log.Contains('※ 管理者として実行すると')) 'True/True'
    }
    $noSaveOut = Join-Path $workRoot 'out_nosave'
    $null = & $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Since 1h -Level 'Critical,Error,Warning,Information' -NoSave -OutputFolder $noSaveOut 2>&1
    Assert-Equal '-NoSave: 保存しない' (Test-Path -LiteralPath $noSaveOut) 'False'
    $usedLetters = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    $freeLetter  = @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $usedLetters -notcontains $_ } | Select-Object -First 1
    if ($freeLetter) {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Since 1h -Level 'Critical,Error,Warning,Information' -OutputFolder "${freeLetter}:\nope" 2>&1 | Out-String)
        Assert-Equal '書き込めない保存先' $log.Contains('保存先に書き込めないため') 'True'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }
} finally {
    if ($sourceCreated) { try { [Diagnostics.EventLog]::DeleteEventSource($source) } catch { Write-Section ('イベントソースを削除できませんでした: {0}' -f $source) } }
    if ($KeepTemp) {
        Write-Section ('一時フォルダを残しました: {0}' -f $workRoot)
    } elseif (Test-Path -LiteralPath $workRoot) {
        # ループするジャンクションは、リンク先をたどらないように先にリンクだけ消す
        try {
            if (Test-Path -LiteralPath $junction) { [IO.Directory]::Delete($junction) }
            [IO.Directory]::Delete('\\?\' + $workRoot, $true)
        } catch {
            Write-Section ('一時フォルダを削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message)
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

exit $script:failCount
