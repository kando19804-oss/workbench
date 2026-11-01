<#
.SYNOPSIS
    Export-TaskList の回帰テストです。

.DESCRIPTION
    タスクスケジューラに、テスト用のフォルダ（\workbench-tests\ の下）を作り、いろいろなトリガー・状態・名前の
    テスト用タスクを登録して、Export-TaskList を実行し、期待値と照合します。
    テスト用のタスクは、トリガーを 2099 年など先の日時にしてあり、勝手には動きません。前回の結果を作るため、
    「cmd /c exit 0」と「cmd /c exit 1」だけを行うタスクを、テストの中で 1 回ずつ実行します。
    テスト用のタスクとフォルダ、一時フォルダの結果は、終了時に削除します。

    タスクのフォルダを作るには、管理者の権限が要ります。管理者でないときは、タスクを使うテストを飛ばします。

    通常は tests\Invoke-Tests.ps1 から呼び出します。終了コードは失敗した件数です。

.PARAMETER Exe
    Export-TaskList を実行する PowerShell（powershell.exe または pwsh）。

.PARAMETER KeepTemp
    結果のフォルダとテスト用のタスクを削除せずに残します（調査用）。

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

$tool     = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Export-TaskList\Export-TaskList.ps1'))
$hostTag  = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$stamp    = '{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Export-TaskList_' + $stamp)
$testTop  = '\workbench-tests'
$testDir  = '{0}\ET_{1}' -f $testTop, $stamp
$testRel  = $testDir.TrimStart('\')
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
        $found = @(Get-ChildItem -LiteralPath $out -Directory -Filter 'TaskList_*')
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
    $summary = ''
    if ($Dir -and (Test-Path -LiteralPath (Join-Path $Dir 'summary.txt'))) { $summary = [IO.File]::ReadAllText((Join-Path $Dir 'summary.txt')) }
    [pscustomobject]@{ Tasks = @(Read-CsvOrEmpty 'tasks.csv'); Skipped = @(Read-CsvOrEmpty 'skipped.csv'); Summary = $summary }
}

# tasks.csv の、テスト用フォルダから見たタスクの行（名前は ="..." の形のこともある）
function Get-TaskRow($Rows, [string]$SubPath) {
    $full = $testDir + '\' + $SubPath
    @($Rows | Where-Object { ($_.'フォルダ'.TrimEnd('\') + '\' + ($_.'タスク名' -replace '^="(.*)"$', '$1')) -eq $full })
}

function Get-TaskField($Rows, [string]$SubPath, [string]$Column) { (Get-TaskRow $Rows $SubPath | ForEach-Object { $_.$Column }) -join '|' }

# コマンドプロンプトから起動した場合の引数の渡り方を確かめるため、バッチファイル経由で実行する
function Invoke-ViaCmd([string]$CommandLine) {
    $batch    = Join-Path $workRoot 'run.cmd'
    $codePage = [Console]::OutputEncoding.CodePage
    $encoding = if ($codePage -eq 65001) { New-Object Text.UTF8Encoding $false } else { [Text.Encoding]::GetEncoding($codePage) }
    [IO.File]::WriteAllLines($batch, [string[]]@('@echo off', $CommandLine), $encoding)
    $log = ConvertTo-FlatText (& cmd.exe /d /c $batch 2>&1 | Out-String)
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# テスト用のタスクを登録する。Configure でトリガーなどを設定する。実行ユーザーは自分（ログオンしているときのみ）
function Register-TestTask([string]$SubPath, [scriptblock]$Configure, [string]$Command = 'exit 0', [string]$Sddl = $null) {
    $definition = $service.NewTask(0)
    $definition.RegistrationInfo.Author = 'workbench-tests'
    $definition.Principal.LogonType = 3
    $definition.Principal.UserId = $currentUser
    $action = $definition.Actions.Create(0)
    $action.Path = 'cmd.exe'
    $action.Arguments = '/c ' + $Command
    & $Configure $definition
    $cut    = $SubPath.LastIndexOf('\')
    $folder = $testFolder
    if ($cut -gt 0) {
        $folderPath = $testDir + '\' + $SubPath.Substring(0, $cut)
        try { $folder = $service.GetFolder($folderPath) } catch { $folder = $testFolder.CreateFolder($SubPath.Substring(0, $cut)) }
    }
    [void]$folder.RegisterTaskDefinition($SubPath.Substring($cut + 1), $definition, 6, $null, $null, 3, $Sddl)
}

# 実行して終わるまで待つ（最大 30 秒）
function Invoke-TestTask([string]$SubPath) {
    $task = $service.GetFolder($testDir).GetTask($SubPath)
    [void]$task.Run($null)
    for ($i = 0; $i -lt 60; $i++) {
        Start-Sleep -Milliseconds 500
        $task = $service.GetFolder($testDir).GetTask($SubPath)
        if ($task.State -ne 4 -and $task.LastTaskResult -ne 0x41301 -and $task.LastRunTime.Year -ge 2000) { return }
    }
}

# テスト用のフォルダの下を、タスク → フォルダの順に消す
function Remove-TestFolder($Folder) {
    foreach ($sub in @($Folder.GetFolders(0))) { Remove-TestFolder $sub }
    foreach ($task in @($Folder.GetTasks(1))) { try { $Folder.DeleteTask($task.Name, 0) } catch { } }
    if ($Folder.Path -ne '\') {
        $parentPath = $Folder.Path.Substring(0, $Folder.Path.LastIndexOf('\'))
        if (-not $parentPath) { $parentPath = '\' }
        try { $service.GetFolder($parentPath).DeleteFolder($Folder.Name, 0) } catch { }
    }
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$userSid     = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$service     = New-Object -ComObject Schedule.Service
$service.Connect()
$testFolder  = $null
$taskReady   = $false
[void][IO.Directory]::CreateDirectory($workRoot)

try {
    Write-Section ('準備: テスト用のタスクを登録（{0}）' -f $testDir)
    try {
        $top = $null
        try { $top = $service.GetFolder($testTop) } catch { $top = $service.GetFolder('\').CreateFolder($testTop.TrimStart('\')) }
        $testFolder = $top.CreateFolder(($testDir.Substring($testTop.Length + 1)))
        $start = '2099-01-01T03:00:00'
        Register-TestTask 'daily' {
            param($d)
            $t = $d.Triggers.Create(2); $t.StartBoundary = $start; $t.DaysInterval = 1
            $t.Repetition.Interval = 'PT1H'; $t.Repetition.Duration = 'PT12H'
            $d.RegistrationInfo.Description = '2026-10-31'
        }
        Register-TestTask 'weekly' {
            param($d)
            $t = $d.Triggers.Create(3); $t.StartBoundary = $start; $t.WeeksInterval = 1; $t.DaysOfWeek = 2 + 8 + 32
            $d.Settings.Enabled = $false
        }
        Register-TestTask 'monthly' {
            param($d)
            $t = $d.Triggers.Create(4); $t.StartBoundary = $start; $t.DaysOfMonth = 1 + (1 -shl 14); $t.MonthsOfYear = 1 + 8 + 64 + 512
        }
        Register-TestTask 'dow' {
            param($d)
            $t = $d.Triggers.Create(5); $t.StartBoundary = $start; $t.WeeksOfMonth = 1; $t.DaysOfWeek = 2
            $t2 = $d.Triggers.Create(1); $t2.StartBoundary = '2099-11-05T09:00:00'; $t2.Enabled = $false
        }
        Register-TestTask 'sub\fail' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start } 'exit 1'
        Register-TestTask 'sub\ok' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start }
        Register-TestTask '日本語 タスク' {
            param($d)
            $t = $d.Triggers.Create(9); $t.UserId = $currentUser
            $d.Settings.Hidden = $true
        }
        Register-TestTask '-dash' { param($d) $t = $d.Triggers.Create(8) }
        Invoke-TestTask 'sub\fail'
        Invoke-TestTask 'sub\ok'
        $taskReady = $true
    } catch {
        Write-Skip 'テスト用のタスクの登録' ('登録できなかったため: {0}' -f $_.Exception.Message)
    }
    Assert-Equal 'テスト用のタスクを登録できた（管理者のときは必ず）' ($taskReady -or -not $isAdmin) 'True'

    if ($taskReady) {
        Write-Section '-Path でテスト用のフォルダを指定'
        $run = Invoke-Tool 'path' @('-Path', $testDir)
        Assert-Equal '終了コード' $run.Code 0
        Assert-Equal '件数の行' ([regex]::Match($run.Log, 'タスク \S+ 件（[^）]*）(  前回失敗 \S+ 件)?').Value) 'タスク 8 件（有効 7・無効 1）  前回失敗 1 件'
        Assert-Equal '範囲の行' ([regex]::IsMatch($run.Log, [regex]::Escape("範囲: $testDir の下") + '(?!（)')) 'True'
        $x = Read-Result $run.Dir
        Assert-Equal 'tasks.csv の行数' $x.Tasks.Count 8
        Assert-Equal '毎日と繰り返し' (Get-TaskField $x.Tasks 'daily' 'トリガー') '毎日 03:00（1 時間ごとに 12 時間繰り返す）'
        Assert-Equal '毎週' (Get-TaskField $x.Tasks 'weekly' 'トリガー') '毎週 月・水・金 03:00'
        Assert-Equal '毎月（月の一部）' (Get-TaskField $x.Tasks 'monthly' 'トリガー') '毎月（1・4・7・10 月）1・15 日 03:00'
        Assert-Equal '毎月の曜日と、無効の 1 回のトリガー' (Get-TaskField $x.Tasks 'dow' 'トリガー') '毎月 第 1 月曜 03:00 / 1 回 2099-11-05 09:00（無効）'
        Assert-Equal 'ログオン時' (Get-TaskField $x.Tasks '日本語 タスク' 'トリガー') ('ログオン時（{0}）' -f $currentUser)
        Assert-Equal '起動時' (Get-TaskField $x.Tasks '-dash' 'トリガー') '起動時'
        Assert-Equal '無効' (Get-TaskField $x.Tasks 'weekly' '状態') '無効'
        Assert-Equal '準備完了' (Get-TaskField $x.Tasks 'daily' '状態') '準備完了'
        Assert-Equal '失敗' ('{0}/{1}' -f (Get-TaskField $x.Tasks 'sub\fail' '前回の結果'), (Get-TaskField $x.Tasks 'sub\fail' '結果のコード')) '失敗/0x00000001'
        Assert-Equal '成功' ('{0}/{1}' -f (Get-TaskField $x.Tasks 'sub\ok' '前回の結果'), (Get-TaskField $x.Tasks 'sub\ok' '結果のコード')) '成功/0x00000000'
        Assert-Equal 'まだ実行していない' ('{0}/{1}' -f (Get-TaskField $x.Tasks 'daily' '前回の結果'), (Get-TaskField $x.Tasks 'daily' '前回の実行日時')) 'まだ実行していない/'
        Assert-Equal '次回の実行日時' (Get-TaskField $x.Tasks 'daily' '次回の実行日時') '2099/01/01 03:00:00'
        Assert-Equal '実行ユーザー' ((Get-TaskField $x.Tasks 'daily' '実行ユーザー') -like ('*' + $env:USERNAME)) 'True'
        Assert-Equal '実行の条件' (Get-TaskField $x.Tasks 'daily' '実行の条件') 'ログオンしているときのみ'
        Assert-Equal '最上位の特権' (Get-TaskField $x.Tasks 'daily' '最上位の特権') 'いいえ'
        Assert-Equal '操作' (Get-TaskField $x.Tasks 'sub\fail' '操作') 'cmd.exe /c exit 1'
        Assert-Equal '作成者' (Get-TaskField $x.Tasks 'daily' '作成者') 'workbench-tests'
        Assert-Equal '隠し' ('{0}/{1}' -f (Get-TaskField $x.Tasks '日本語 タスク' '隠し'), (Get-TaskField $x.Tasks 'daily' '隠し')) 'はい/'
        Assert-Equal '日付に見える説明は ="..."' (Get-TaskField $x.Tasks 'daily' '説明') '="2026-10-31"'
        Assert-Equal '- で始まる名前は ="..."' @($x.Tasks | Where-Object { $_.'タスク名' -eq '="-dash"' }).Count 1
        $paths = [string[]]@($x.Tasks | ForEach-Object { ($_.'フォルダ'.TrimEnd('\') + '\' + ($_.'タスク名' -replace '^="(.*)"$', '$1')).Substring($testDir.Length + 1) })
        Assert-Equal 'パスの順（フォルダの中身はフォルダのすぐ後）' ($paths -join ',') '-dash,daily,dow,monthly,sub\fail,sub\ok,weekly,日本語 タスク'

        Write-Section 'XML'
        $xmlFile = Join-Path $run.Dir ('xml\{0}\sub\fail.xml' -f $testRel)
        Assert-Equal 'タスクのフォルダの構成で保存' (Test-Path -LiteralPath $xmlFile) 'True'
        Assert-Equal 'XML ファイルの列' (Get-TaskField $x.Tasks 'sub\fail' 'XML ファイル') ('{0}\sub\fail.xml' -f $testRel)
        $bytes = [IO.File]::ReadAllBytes($xmlFile)
        Assert-Equal 'UTF-16（BOM 付き）' ('{0:X2}{1:X2}' -f $bytes[0], $bytes[1]) 'FFFE'
        $doc = New-Object Xml.XmlDocument
        $loaded = $true
        try { $doc.Load($xmlFile) } catch { $loaded = $false }
        Assert-Equal 'XML として読める' $loaded 'True'
        Assert-Equal 'XML の中身は元のタスク' ($doc.Task.Actions.Exec.Arguments) '/c exit 1'
        Assert-Equal 'XML の数' @(Get-ChildItem -LiteralPath (Join-Path $run.Dir 'xml') -Recurse -File -Filter '*.xml').Count 8
        # 保存した XML を、schtasks で取り込み直せるか
        $importName = '{0}\reimported' -f $testDir
        $null = & schtasks.exe /create /xml $xmlFile /tn $importName 2>&1
        Assert-Equal 'schtasks /create /xml で取り込める' $LASTEXITCODE 0
        # 後のテストの件数が変わらないよう、取り込んだタスクは消しておく
        try { $testFolder.DeleteTask('reimported', 0) } catch { }

        Write-Section '-NoXml、-NoSave'
        $run = Invoke-Tool 'noxml' @('-Path', $testDir, '-NoXml')
        Assert-Equal '-NoXml: 終了コード' $run.Code 0
        Assert-Equal '-NoXml: XML を保存しない' (Test-Path -LiteralPath (Join-Path $run.Dir 'xml')) 'False'
        Assert-Equal '-NoXml: XML ファイルの列は空欄' (Get-TaskField (Read-Result $run.Dir).Tasks 'daily' 'XML ファイル') ''
        Assert-Equal '-NoXml: 表示' $run.Log.Contains('XML: 保存しない') 'True'
        $noSaveOut = Join-Path $workRoot 'out_nosave'
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path $testDir -NoSave -OutputFolder $noSaveOut 2>&1 | Out-String)
        Assert-Equal '-NoSave: 保存しない' ('{0}/{1}' -f $LASTEXITCODE, (Test-Path -LiteralPath $noSaveOut)) '0/False'
        Assert-Equal '-NoSave: 一覧を表示する' $log.Contains("$testDir\sub\fail") 'True'

        Write-Section '既定の範囲と -All'
        $run = Invoke-Tool 'default' @()
        $x = Read-Result $run.Dir
        Assert-Equal '既定: 終了コード' $run.Code 0
        Assert-Equal '既定: テスト用のタスクを含む' (@(Get-TaskRow $x.Tasks 'daily').Count) 1
        Assert-Equal '既定: \Microsoft\ の下を含まない' @($x.Tasks | Where-Object { $_.'フォルダ' -like '\Microsoft*' }).Count 0
        Assert-Equal '既定: 範囲の行' $run.Log.Contains('範囲: \ の下（\Microsoft\ の下を除く）') 'True'
        $run = Invoke-Tool 'all' @('-All', '-NoXml')
        $x = Read-Result $run.Dir
        Assert-Equal '-All: \Microsoft\ の下を含む' (@($x.Tasks | Where-Object { $_.'フォルダ' -like '\Microsoft\*' }).Count -gt 0) 'True'
        Assert-Equal '-All: 範囲の行' $run.Log.Contains('範囲: \ の下（Windows 標準のタスクを含む）') 'True'
        # -All のときは 200 件を超えることが多い。保存したときは、画面は 200 件まで
        if ($x.Tasks.Count -gt 200) {
            Assert-Equal '-All: 画面は 200 件まで' $run.Log.Contains(('（ほか {0:N0} 件は tasks.csv を参照してください）' -f ($x.Tasks.Count - 200))) 'True'
        } else {
            Write-Skip '画面は 200 件まで' 'タスクが 200 件以下のため'
        }
        $defrag = '\Microsoft\Windows\Defrag'
        $exists = $true
        try { [void]$service.GetFolder($defrag) } catch { $exists = $false }
        if ($exists) {
            $run = Invoke-Tool 'ms' @('-Path', $defrag)
            Assert-Equal '-Path で \Microsoft\ の下を指定すると出力する' ((@((Read-Result $run.Dir).Tasks).Count -gt 0)) 'True'
        } else {
            Write-Skip '-Path で \Microsoft\ の下' "$defrag がないため"
        }

        Write-Section 'コマンドプロンプトから起動'
        $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -NoSave -Path "{2}\"' -f $Exe, $tool, $testDir)
        Assert-Equal '末尾の \ 付きの -Path: 終了コード' $run.Code 0
        Assert-Equal '末尾の \ 付きの -Path: 件数' $run.Log.Contains('タスク 8 件') 'True'
        $run = Invoke-ViaCmd ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -Path "{2}" -OutputFolder "{3}\" -NoSave' -f $Exe, $tool, $testDir, $workRoot)
        Assert-Equal '"パス\" の後ろに引数: 終了コード' $run.Code 1
        Assert-Equal '"パス\" の後ろに引数: ヒント' ([regex]::IsMatch($run.Log, '最後に \\ を付けないでください')) 'True'

        Write-Section 'レビューで見つかった問題の再発防止'
        # XML のファイル名が重なったら _2 を付ける（末尾の . はファイル名にできず、消すため）
        $collision = $true
        try {
            Register-TestTask 'col\b' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start }
            Register-TestTask 'col\b.' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start }
        } catch { $collision = $false }
        if ($collision -and @($service.GetFolder("$testDir\col").GetTasks(1)).Count -eq 2) {
            $run = Invoke-Tool 'collision' @('-Path', "$testDir\col")
            $files = [string[]]@((Read-Result $run.Dir).Tasks | ForEach-Object { $_.'XML ファイル' })
            [Array]::Sort($files, [StringComparer]::Ordinal)
            Assert-Equal 'XML の名前が重なったら _2' ($files -join ',') ('{0}\col\b.xml,{0}\col\b_2.xml' -f $testRel)
        } else {
            Write-Skip 'XML の名前の重なり' '末尾が . のタスクを登録できなかったため'
        }
        # 読み取りを拒否したタスク。プロパティを読むときの権限エラーも「権限なし」に分ける
        $denied = $true
        try { Register-TestTask 'deny\secret' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start } 'exit 0' ('D:(D;;FR;;;{0})(A;;FA;;;BA)(A;;FA;;;SY)' -f $userSid) } catch { $denied = $false }
        if ($denied) {
            $run = Invoke-Tool 'denied' @('-Path', "$testDir\deny")
            $row = @((Read-Result $run.Dir).Skipped | Where-Object { $_.'パス' -like '*\deny*' })
            if ($row.Count -gt 0) {
                Assert-Equal '読めないタスクは権限なし' (($row | ForEach-Object { $_.'理由' }) -join ',') '権限なし'
            } else {
                Write-Skip '読めないタスク' '一覧からも消えたため（エラーにならなかった）'
            }
        } else {
            Write-Skip '読めないタスク' 'アクセス拒否のタスクを登録できなかったため'
        }
    }

    Write-Section '止まるとき・保存できないとき'
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -Path '\no\such\folder' -NoSave 2>&1 | Out-String)
    Assert-Equal '存在しないフォルダ' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('タスクのフォルダが見つかりません: \no\such\folder')) '1/True'
    $usedLetters = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    $freeLetter  = @('Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z') | Where-Object { $usedLetters -notcontains $_ } | Select-Object -First 1
    if ($freeLetter) {
        $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -OutputFolder "${freeLetter}:\nope" 2>&1 | Out-String)
        Assert-Equal '書き込めない保存先' ('{0}/{1}' -f $LASTEXITCODE, $log.Contains('保存先に書き込めないため')) '0/True'
    } else {
        Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
    }
    $log = ConvertTo-FlatText (& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool -NoSave 2>&1 | Out-String)
    Assert-Equal '管理者でないときの注意' $log.Contains('※ 管理者として実行していないため') (-not $isAdmin)
} finally {
    if ($KeepTemp) {
        Write-Section ('結果とテスト用のタスクを残しました: {0} / {1}' -f $workRoot, $testDir)
    } else {
        if ($testFolder) {
            Remove-TestFolder $testFolder
            # ほかのテストが使っていなければ、\workbench-tests も消す
            try {
                $top = $service.GetFolder($testTop)
                if (@($top.GetFolders(0)).Count -eq 0 -and @($top.GetTasks(1)).Count -eq 0) { $service.GetFolder('\').DeleteFolder($testTop.TrimStart('\'), 0) }
            } catch { }
            $left = $true
            try { [void]$service.GetFolder($testDir) } catch { $left = $false }
            if ($left) { Write-Section ('テスト用のタスクを削除できませんでした: {0}' -f $testDir) }
        }
        if (Test-Path -LiteralPath $workRoot) {
            try { [IO.Directory]::Delete('\\?\' + $workRoot, $true) } catch { Write-Section ('結果を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message) }
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

exit $script:failCount
