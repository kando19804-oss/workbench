<#
.SYNOPSIS
    Export-TaskList の回帰テストです。

.DESCRIPTION
    タスクスケジューラに、テスト用のフォルダ（\workbench-tests\ の下）を作り、いろいろなトリガー・状態・名前の
    テスト用のタスクを登録して、Export-TaskList を実行し、期待値と照合します。
    テスト用のタスクは、トリガーを 2099 年など先の日時にしてあり、勝手には動きません。前回の結果を作るため、
    終了コードを返すだけの VBScript を、テストの中で実行します（wscript.exe で動かすので、画面に窓は出ません）。
    XML ファイルの列が - で始まる場合を確かめるため、ルート（\）にも、名前に目印（-wbt_ET_<日時>）を付けた
    タスクを 1 つだけ登録します。

    管理者だけが登録できるもの（起動時のトリガー、SYSTEM など）は、管理者のときだけ試します。
    読めないタスク（権限なし）は、タスクスケジューラが一覧から除くため、再現できず試しません。

    結果の既定の保存先はスクリプトのフォルダなので、リポジトリに結果を作らないよう、一時フォルダにコピーした
    スクリプトを動かします。テスト用のタスクとフォルダ、一時フォルダの結果は、終了時に削除します。

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

$repoTool  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tools\Export-TaskList\Export-TaskList.ps1'))
$hostTag   = if ($Exe -like 'pwsh*') { '7' } else { '51' }
$stamp     = '{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), $hostTag
$workRoot  = Join-Path ([IO.Path]::GetTempPath()) ('workbench-tests\Export-TaskList_' + $stamp)
$tool      = Join-Path $workRoot 'tool\Export-TaskList.ps1'
$utf8Bom   = New-Object Text.UTF8Encoding $true
$utf8NoBom = New-Object Text.UTF8Encoding $false
$testTop   = '\workbench-tests'
$testDir   = '{0}\ET_{1}' -f $testTop, $stamp
$testRel   = $testDir.TrimStart('\')
$rootName  = '-wbt_ET_' + $stamp                       # ルートに置くテスト用のタスク
$rootLike  = '^-wbt_ET_\d{8}_\d{6}_(51|7)$'            # 途中で止まった前回のテストの残りも、これで見分けて消す
$start     = '2099-01-01T03:00:00'
$taskHead  = '"フォルダ","タスク名","状態","実行ユーザー","実行の条件","最上位の特権","トリガー","操作","作業フォルダ","前回の実行日時","前回の結果","結果のコード","次回の実行日時","作成者","作成日","説明","隠し","XML ファイル"'
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

# Log には、画面への出力から改行を取り除いたものを入れる。Lines は画面の 1 行ずつ（表の行を数えるため）
function Invoke-Tool([string]$Label, [string[]]$ToolArgs) {
    $out   = Join-Path $workRoot "out\$Label"
    $lines = @(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs -OutputFolder $out 2>&1) | ForEach-Object { "$_" }
    $code  = $LASTEXITCODE
    [pscustomobject]@{ Code = $code; Log = (ConvertTo-FlatText (@($lines) -join "`n")); Lines = @($lines); Dir = (Get-RunDir $out) }
}

# 保存しない実行や、引数の誤りで止まることを確かめる実行
function Invoke-Plain([string[]]$ToolArgs) {
    $log = ConvertTo-FlatText ((@(& $Exe -NoProfile -ExecutionPolicy Bypass -File $tool @ToolArgs 2>&1) | ForEach-Object { "$_" }) -join "`n")
    [pscustomobject]@{ Code = $LASTEXITCODE; Log = $log }
}

# 相対パスの基準を確かめるため、-OutputFolder を足さず、指定したフォルダから起動する
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

# [powershell]::Create() で動かすスクリプトを、一時フォルダに書く（Write-Host の色を受け取るためと、Ctrl+C の代わりに Stop() で止めるため）
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
    $found = @(Get-ChildItem -LiteralPath $Parent -Directory -Filter 'TaskList_*' -ErrorAction SilentlyContinue | Sort-Object Name)
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
    [pscustomobject]@{ Tasks = @(Read-CsvOrEmpty 'tasks.csv'); Skipped = @(Read-CsvOrEmpty 'skipped.csv'); Lines = $lines }
}

function Join-Sorted($Items) {
    $array = [string[]]@($Items | Where-Object { $null -ne $_ })
    [Array]::Sort($array, [StringComparer]::Ordinal)
    $array -join ','
}

# 保存されたファイルとフォルダの名前の一覧
function Get-ResultNames([string]$Dir) {
    if (-not $Dir) { return '(なし)' }
    Join-Sorted (Get-ChildItem -LiteralPath $Dir | ForEach-Object { $_.Name })
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

# summary.txt の n 行目（ないときは空）
function Get-Line($Lines, [int]$Index) {
    $all = @($Lines)
    if ($Index -ge $all.Count -or -$Index -gt $all.Count) { return '' }
    $all[$Index]
}

# 画面の件数の行（例: タスク 4 件（有効 3・無効 1）  前回失敗 1 件）。所要時間は除く
function Get-CountText([string]$Log) { [regex]::Match($Log, 'タスク [\d,]+ 件（[^）]*）(  前回失敗 [\d,]+ 件)?(?=  所要)').Value }

# tasks.csv から数えた、件数の行。前回失敗は、有効なタスクだけで数える
function Get-ExpectedCount($Rows) {
    $all  = @($Rows).Count
    $off  = @($Rows | Where-Object { $_.'状態' -eq '無効' }).Count
    $fail = @($Rows | Where-Object { $_.'状態' -ne '無効' -and $_.'前回の結果' -eq '失敗' }).Count
    $text = 'タスク {0:N0} 件（有効 {1:N0}・無効 {2:N0}）' -f $all, ($all - $off), $off
    if ($fail -gt 0) { $text += '  前回失敗 {0:N0} 件' -f $fail }
    $text
}

# 画面の表の行の数（状態で始まる行）
function Get-TableRowCount($Lines) { @($Lines | Where-Object { $_ -match '^ (準備完了|無効|実行中|キューに登録|不明)  ' }).Count }

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

function Get-TaskPath([string]$Name) { "$testDir\$Name" }

# xml\ から見た XML のパス
function Get-XmlRel([string]$Name) { "$testRel\$Name.xml" }

# tasks.csv の行のタスクのパス（名前は ="..." の形のこともある）
function Get-RowPath($Row) { $Row.'フォルダ'.TrimEnd('\') + '\' + ($Row.'タスク名' -replace '^="(.*)"$', '$1') }

function Get-TaskRow($Rows, [string]$Name) {
    $full = Get-TaskPath $Name
    @($Rows | Where-Object { (Get-RowPath $_) -eq $full })
}

function Get-TaskField($Rows, [string]$Name, [string]$Column) { (@(Get-TaskRow $Rows $Name) | ForEach-Object { $_.$Column }) -join '|' }

# tasks.csv に出たテスト用のタスクを、出た順に、テストの中の名前で並べる
function Get-TestOrder($Rows) {
    (@($Rows | ForEach-Object { Get-RowPath $_ } | Where-Object { $_.StartsWith($testDir + '\') } | ForEach-Object { $_.Substring($testDir.Length + 1) })) -join ','
}

function Get-TestSubFolder([string]$Sub) {
    try { return $service.GetFolder("$testDir\$Sub") } catch { return $testFolder.CreateFolder($Sub) }
}

# テスト用のタスクを登録する。既定の操作は、終了コード 0 を返す VBScript。実行ユーザーは自分（ログオンしているときのみ）。
# 同じ名前のタスクがあれば、上書きせずに失敗させる（TASK_CREATE）
function Register-TestTask([string]$Name, [scriptblock]$Configure, [string]$Sddl = $null, [string]$User = $currentUser, [int]$LogonType = 3, $Folder = $null) {
    $definition = $service.NewTask(0)
    $definition.RegistrationInfo.Author = 'workbench-tests'
    $definition.Principal.LogonType = $LogonType
    $action = $definition.Actions.Create(0)
    $action.Path = 'wscript.exe'
    $action.Arguments = '//B //Nologo "{0}"' -f $okScript
    if ($Configure) { & $Configure $definition }
    $cut = $Name.LastIndexOf('\')
    if (-not $Folder) { $Folder = if ($cut -gt 0) { Get-TestSubFolder $Name.Substring(0, $cut) } else { $testFolder } }
    [void]$Folder.RegisterTaskDefinition($Name.Substring($cut + 1), $definition, 2, $User, $null, $LogonType, $Sddl)
}

# 登録できなかったら、その理由を残して続ける（ほかのテストは動かす）
function Add-TestTask([string]$Name, [scriptblock]$Configure, [string]$User = $currentUser, [int]$LogonType = 3) {
    try { Register-TestTask $Name $Configure $null $User $LogonType } catch { $registerErrors.Add(('{0}: {1}' -f $Name, ($_.Exception.Message -replace '\s+', ' '))) }
}

function Get-TestTask([string]$Name) { $service.GetFolder('\').GetTask((Get-TaskPath $Name)) }

# 実行して、終わるまで待つ（最大 30 秒）
function Invoke-TestTask([string]$Name) {
    [void](Get-TestTask $Name).Run($null)
    for ($i = 0; $i -lt 60; $i++) {
        Start-Sleep -Milliseconds 500
        $task = Get-TestTask $Name
        if ($task.State -ne 4 -and $task.LastTaskResult -ne 0x41301 -and $task.LastRunTime.Year -ge 2000) { return $true }
    }
    $false
}

function Remove-TestTask($Folder, $Task) {
    try { if ($Task.State -eq 4) { $Task.Stop(0) } } catch { }
    try { $Folder.DeleteTask($Task.Name, 0) } catch { }
}

# テスト用のフォルダの下を、タスク → フォルダの順に消す
function Remove-TestFolder([string]$FolderPath) {
    $folder = $null
    try { $folder = $service.GetFolder($FolderPath) } catch { return }
    foreach ($sub in @($folder.GetFolders(0))) { Remove-TestFolder $sub.Path }
    foreach ($task in @($folder.GetTasks(1))) { Remove-TestTask $folder $task }
    $cut = $FolderPath.LastIndexOf('\')
    $parentPath = if ($cut -gt 0) { $FolderPath.Substring(0, $cut) } else { '\' }
    try { $service.GetFolder($parentPath).DeleteFolder($FolderPath.Substring($cut + 1), 0) } catch { }
}

# COM から消せなかったもの（読み取りを拒否したタスクや、登録の途中で拒否されて残ったタスクのファイル）は、
# タスクのファイル（%windir%\System32\Tasks の下）のアクセス拒否を外してから消す。Sub を指定したときは、そのフォルダの下だけ
function Clear-TestLeftovers([string]$Sub) {
    $folderPath = if ($Sub) { "$testDir\$Sub" } else { $testDir }
    $disk = Join-Path $env:windir ('System32\Tasks' + $folderPath)
    if (-not [IO.Directory]::Exists($disk)) { return }
    $null = & icacls.exe $disk /remove:d "*$userSid" /t /c /q 2>&1
    foreach ($file in @(Get-ChildItem -LiteralPath $disk -Recurse -File -Force -ErrorAction SilentlyContinue)) {
        $taskPath = $folderPath + $file.FullName.Substring($disk.Length)
        $cut = $taskPath.LastIndexOf('\')
        try {
            $service.GetFolder($taskPath.Substring(0, $cut)).DeleteTask($taskPath.Substring($cut + 1), 0)
        } catch {
            # タスクスケジューラに登録されていない（ファイルだけが残った）ときは、ファイルを消す
            $ex = $_.Exception
            while ($ex.InnerException) { $ex = $ex.InnerException }
            if ($ex.HResult -eq -2147024894) { try { $file.Delete() } catch { } }
        }
    }
    Remove-TestFolder $folderPath
}

function Remove-RootTestTasks {
    $root = $service.GetFolder('\')
    foreach ($task in @($root.GetTasks(1))) { if ($task.Name -match $rootLike) { Remove-TestTask $root $task } }
}

#endregion

$savedEncoding = $null
try { $savedEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$isAdmin     = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$userSid     = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied      = New-Object 'System.Collections.Generic.List[string]'   # アクセス拒否を設定したパス（終了時に戻す）
$registerErrors = New-Object 'System.Collections.Generic.List[string]'
$service     = New-Object -ComObject Schedule.Service
$service.Connect()
$testFolder  = $null
$okScript    = Join-Path $workRoot 'exit0.vbs'
$failScript  = Join-Path $workRoot 'exit1.vbs'
$sleepScript = Join-Path $workRoot 'sleep.vbs'

try {
    Write-Section ('準備: テスト用のタスクを登録（{0}）' -f $testDir)
    [void][IO.Directory]::CreateDirectory((Split-Path $tool))
    Copy-Item -LiteralPath $repoTool -Destination $tool
    [IO.File]::WriteAllText($okScript, "WScript.Quit 0`r`n", [Text.Encoding]::ASCII)
    [IO.File]::WriteAllText($failScript, "WScript.Quit 1`r`n", [Text.Encoding]::ASCII)
    [IO.File]::WriteAllText($sleepScript, "WScript.Sleep 120000`r`n", [Text.Encoding]::ASCII)
    Remove-RootTestTasks
    try {
        $top = $null
        try { $top = $service.GetFolder($testTop) } catch { $top = $service.GetFolder('\').CreateFolder($testTop.TrimStart('\')) }
        $testFolder = $top.CreateFolder($testDir.Substring($testTop.Length + 1))
    } catch {
        Write-Skip 'テスト用のフォルダ' ('作れなかったため: {0}' -f ($_.Exception.Message -replace '\s+', ' '))
    }
    # タスクのフォルダは、Windows の既定では管理者でなくても作れる
    Assert-Equal 'テスト用のフォルダを作れた' ([bool]$testFolder) 'True'
} catch {
    '  テストの準備で止まりました: {0}' -f $_.Exception.Message
    $script:failCount++
}

try {
    if ($testFolder) {
        $hasWscript = [bool](Get-Command wscript.exe -ErrorAction SilentlyContinue)
        if ($isAdmin) {
            # 管理者だけが登録できるもの
            Add-TestTask 'sub\boot' { param($d) $null = $d.Triggers.Create(8) }
            Add-TestTask 'sub\logonall' { param($d) $null = $d.Triggers.Create(9) }
            Add-TestTask 'sub\session' { param($d) $t = $d.Triggers.Create(11); $t.StateChange = 7 }
            Add-TestTask 'sub\system' { param($d) $d.Principal.RunLevel = 1; $t = $d.Triggers.Create(1); $t.StartBoundary = $start } 'SYSTEM' 5
            Add-TestTask 'sub\s4u' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start } $currentUser 2
        }
        Add-TestTask 'daily' {
            param($d)
            $t = $d.Triggers.Create(2); $t.StartBoundary = $start; $t.DaysInterval = 1
            $t.Repetition.Interval = 'PT1H'; $t.Repetition.Duration = 'PT12H'
            foreach ($a in $d.Actions) { $a.WorkingDirectory = 'C:\作業 フォルダ' }
            $d.RegistrationInfo.Description = '2026-10-31'
            $d.RegistrationInfo.Date = '2026-10-31T09:15:00'
        }
        Add-TestTask 'days' {
            param($d)
            $t = $d.Triggers.Create(2); $t.StartBoundary = $start; $t.DaysInterval = 3; $t.Repetition.Interval = 'PT30M'
            $t2 = $d.Triggers.Create(2); $t2.StartBoundary = '2099-01-01T09:30:00'; $t2.DaysInterval = 1
            $t2.Repetition.Interval = 'PT5M'; $t2.Repetition.Duration = 'P1DT2H'
        }
        Add-TestTask 'weekly' {
            param($d)
            $t = $d.Triggers.Create(3); $t.StartBoundary = $start; $t.WeeksInterval = 1; $t.DaysOfWeek = 2 + 8 + 32
            $d.Settings.Enabled = $false
        }
        Add-TestTask 'weeks2' { param($d) $t = $d.Triggers.Create(3); $t.StartBoundary = $start; $t.WeeksInterval = 2; $t.DaysOfWeek = 1 + 64 }
        Add-TestTask 'monthly' {
            param($d)
            $t = $d.Triggers.Create(4); $t.StartBoundary = $start; $t.DaysOfMonth = 1 + (1 -shl 14); $t.MonthsOfYear = 1 + 8 + 64 + 512
        }
        Add-TestTask 'lastday' {
            param($d)
            $t = $d.Triggers.Create(4); $t.StartBoundary = $start; $t.DaysOfMonth = 1 -shl 9; $t.RunOnLastDayOfMonth = $true; $t.MonthsOfYear = 4095
        }
        Add-TestTask 'dow' {
            param($d)
            $t = $d.Triggers.Create(5); $t.StartBoundary = $start; $t.WeeksOfMonth = 1; $t.DaysOfWeek = 2; $t.MonthsOfYear = 4095
            $t2 = $d.Triggers.Create(1); $t2.StartBoundary = '2099-11-05T09:00:00'; $t2.Enabled = $false
        }
        Add-TestTask 'lastweek' {
            param($d)
            $t = $d.Triggers.Create(5); $t.StartBoundary = $start; $t.WeeksOfMonth = 2 + 8; $t.RunOnLastWeekOfMonth = $true; $t.DaysOfWeek = 4 + 16; $t.MonthsOfYear = 2
        }
        Add-TestTask 'once' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = '2099-11-05T09:00:00+09:00' }
        Add-TestTask 'event' {
            param($d)
            $t = $d.Triggers.Create(0)
            $t.Subscription = '<QueryList><Query Id="0" Path="Application"><Select Path="Application">*[System[Provider[@Name=''workbench-tests''] and EventID=9999]]</Select></Query></QueryList>'
        }
        Add-TestTask 'idle' { param($d) $null = $d.Triggers.Create(6) }
        Add-TestTask 'regist' { param($d) $null = $d.Triggers.Create(7) }
        Add-TestTask '日本語 タスク' { param($d) $t = $d.Triggers.Create(9); $t.UserId = $currentUser; $d.Settings.Hidden = $true }
        Add-TestTask 'actions' {
            param($d)
            $d.Actions.Clear()
            $c = $d.Actions.Create(5); $c.ClassId = '{11111111-2222-3333-4444-555555555555}'
            $e = $d.Actions.Create(0); $e.Path = 'C:\Program Files\Example\tool.exe'; $e.Arguments = '-a "b c"'; $e.WorkingDirectory = 'D:\work'
            $e2 = $d.Actions.Create(0); $e2.Path = 'cmd.exe'; $e2.Arguments = '/c exit 0'
            # 全角のピリオドは数字の判定の正規表現に当たらないため、半角にそろえてから判定しないと見落とす
            $d.RegistrationInfo.Description = '１．５'
        }
        Add-TestTask '-dash' {
            param($d)
            $t = $d.Triggers.Create(0)
            $t.Subscription = '<QueryList><Query Id="0" Path="Application"><Select Path="Application">*[System[Provider[@Name=''workbench-tests''] and EventID=9998]]</Select></Query></QueryList>'
            $d.RegistrationInfo.Author = '+81 3'
            $d.RegistrationInfo.Description = '=' + ('x' * 299)
        }
        foreach ($name in 'ok', 'b', 'b.', 'b ', 'sub x', 'sub-x', 'sub\plain', 'Zeta') { Add-TestTask $name { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start } }
        foreach ($name in 'fail', 'offfail') {
            Add-TestTask $name { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start; foreach ($a in $d.Actions) { $a.Arguments = '//B //Nologo "{0}"' -f $failScript } }
        }
        try {
            Register-TestTask $rootName { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start } $null $currentUser 3 $service.GetFolder('\')
        } catch { $registerErrors.Add(('{0}: {1}' -f $rootName, ($_.Exception.Message -replace '\s+', ' '))) }
        Assert-Equal 'テスト用のタスクを登録できた' ($registerErrors -join ' | ') ''
        $ran = $false
        if ($hasWscript) {
            $ran = (Invoke-TestTask 'ok') -and (Invoke-TestTask 'fail') -and (Invoke-TestTask 'offfail')
            Assert-Equal 'テスト用のタスクを実行できた' $ran 'True'
        } else {
            Write-Skip 'テスト用のタスクの実行' 'wscript.exe がないため'
        }
        # 前回失敗したあとで無効にしたタスク（昔の失敗は「前回失敗」に数えない）
        (Get-TestTask 'offfail').Enabled = $false
        $extra = if ($isAdmin) { 5 } else { 0 }
        function Get-ExactCount([int]$Add) {
            $text = 'タスク {0} 件（有効 {1}・無効 2）' -f (25 + $extra + $Add), (23 + $extra + $Add)
            if ($ran) { $text += '  前回失敗 1 件' }
            $text
        }
        $exactCount = Get-ExactCount 0

        Write-Section '一覧（-Path でテスト用のフォルダを指定）'
        $run = Invoke-Tool 'main' @('-Path', $testDir)
        $x = Read-Result $run.Dir
        Assert-Equal '終了コード' $run.Code 0
        Assert-Equal '保存するファイル' (Get-ResultNames $run.Dir) 'skipped.csv,summary.txt,tasks.csv,xml'
        Assert-Equal 'BOM 付き UTF-8' (Get-BomFlags $run.Dir) 'skipped.csv=True,summary.txt=True,tasks.csv=True'
        Assert-Equal 'tasks.csv の見出し' (Get-FirstLine $run.Dir 'tasks.csv') $taskHead
        Assert-Equal 'skipped.csv の見出し' (Get-FirstLine $run.Dir 'skipped.csv') '"理由","パス","詳細"'
        Assert-Equal 'skipped.csv は見出しだけ' $x.Skipped.Count 0
        Assert-Equal 'summary の 1 行目' (Get-Line $x.Lines 0) 'タスクスケジューラ出力 (Export-TaskList 1.0.0)'
        $adminText = if ($isAdmin) { 'はい' } else { 'いいえ' }
        Assert-Equal 'summary の対象の行' ([regex]::IsMatch((Get-Line $x.Lines 1), ('^対象: {0}（管理者: {1}）   \d{{4}}-\d{{2}}-\d{{2}} \d{{2}}:\d{{2}}:\d{{2}}$' -f [regex]::Escape($env:COMPUTERNAME), $adminText))) 'True'
        Assert-Equal 'summary の範囲の行' (Get-Line $x.Lines 2) "範囲: $testDir の下"
        Assert-Equal 'summary の最後は保存先' (Get-Line $x.Lines -1) ('保存先: {0}' -f $run.Dir)
        Assert-Equal '画面の最後は保存先' $run.Log.EndsWith(('保存先: {0}' -f $run.Dir)) 'True'
        $expectedNames = @('-dash', 'actions', 'b', 'b ', 'b.', 'daily', 'days', 'dow', 'event', 'fail', 'idle', 'lastday', 'lastweek', 'monthly', 'offfail', 'ok', 'once', 'regist')
        if ($isAdmin) { $expectedNames += @('sub\boot', 'sub\logonall') }
        $expectedNames += 'sub\plain'
        if ($isAdmin) { $expectedNames += @('sub\s4u', 'sub\session', 'sub\system') }
        $expectedNames += @('sub x', 'sub-x', 'weekly', 'weeks2', 'Zeta', '日本語 タスク')
        Assert-Equal 'タスクの順（\ はどの文字よりも前、大文字と小文字は区別しない）' (Get-TestOrder $x.Tasks) ($expectedNames -join ',')
        $rowCount = $x.Tasks.Count
        Assert-Equal '件数の行' (Get-CountText $run.Log) $exactCount
        Assert-Equal '件数の行は tasks.csv と合う' (Get-CountText $run.Log) (Get-ExpectedCount $x.Tasks)
        Assert-Equal 'summary の件数の行' ((@($x.Lines | Where-Object { $_ -like 'タスク * 件（*' }) | ForEach-Object { Get-CountText $_ }) -join '|') $exactCount
        Assert-Equal '画面の表の行の数（200 件以下はすべて）' (Get-TableRowCount $run.Lines) $rowCount

        Write-Section 'トリガー'
        Assert-Equal '毎日と繰り返し' (Get-TaskField $x.Tasks 'daily' 'トリガー') '毎日 03:00（1 時間ごとに 12 時間繰り返す）'
        Assert-Equal '何日ごと・無期限の繰り返し・日と時間の期間' (Get-TaskField $x.Tasks 'days' 'トリガー') '3 日ごと 03:00（30 分ごとに無期限に繰り返す） / 毎日 09:30（5 分ごとに 1 日 2 時間繰り返す）'
        Assert-Equal '毎週' (Get-TaskField $x.Tasks 'weekly' 'トリガー') '毎週 月・水・金 03:00'
        Assert-Equal '何週ごと（日曜と土曜）' (Get-TaskField $x.Tasks 'weeks2' 'トリガー') '2 週ごと 日・土 03:00'
        Assert-Equal '毎月（月の一部）' (Get-TaskField $x.Tasks 'monthly' 'トリガー') '毎月（1・4・7・10 月）1・15 日 03:00'
        Assert-Equal '毎月の最終日' (Get-TaskField $x.Tasks 'lastday' 'トリガー') '毎月 10 日・最終日 03:00'
        Assert-Equal '毎月の曜日と、無効の 1 回のトリガー' (Get-TaskField $x.Tasks 'dow' 'トリガー') '毎月 第 1 月曜 03:00 / 1 回 2099-11-05 09:00（無効）'
        Assert-Equal '毎月の最終週と、月の一部' (Get-TaskField $x.Tasks 'lastweek' 'トリガー') '毎月（2 月）第 2・第 4・最終 火・木曜 03:00'
        Assert-Equal '1 回（時差付きの日時）' (Get-TaskField $x.Tasks 'once' 'トリガー') '1 回 2099-11-05 09:00'
        Assert-Equal 'イベント' (Get-TaskField $x.Tasks 'event' 'トリガー') 'イベント発生時'
        Assert-Equal 'アイドル時' (Get-TaskField $x.Tasks 'idle' 'トリガー') 'アイドル時'
        Assert-Equal '作成・変更時' (Get-TaskField $x.Tasks 'regist' 'トリガー') 'タスクの作成・変更時'
        Assert-Equal 'ログオン時（ユーザーを指定）' ((Get-TaskField $x.Tasks '日本語 タスク' 'トリガー') -match ('^ログオン時（(.+\\)?' + [regex]::Escape($env:USERNAME) + '）$')) 'True'
        Assert-Equal 'トリガーなし' (Get-TaskField $x.Tasks 'actions' 'トリガー') ''
        if ($isAdmin) {
            Assert-Equal '起動時' (Get-TaskField $x.Tasks 'sub\boot' 'トリガー') '起動時'
            Assert-Equal 'ログオン時（すべてのユーザー）' (Get-TaskField $x.Tasks 'sub\logonall' 'トリガー') 'ログオン時（すべてのユーザー）'
            Assert-Equal 'セッションの状態の変更時' (Get-TaskField $x.Tasks 'sub\session' 'トリガー') 'セッションの状態の変更時'
        } else {
            Write-Skip '起動時・すべてのユーザーのログオン時・セッションの状態の変更時' '管理者でないと登録できないため'
        }

        Write-Section '状態・結果・実行ユーザー・操作'
        Assert-Equal '無効' (Get-TaskField $x.Tasks 'weekly' '状態') '無効'
        Assert-Equal '準備完了' (Get-TaskField $x.Tasks 'daily' '状態') '準備完了'
        Assert-Equal 'まだ実行していない' ('{0}/{1}' -f (Get-TaskField $x.Tasks 'daily' '前回の結果'), (Get-TaskField $x.Tasks 'daily' '前回の実行日時')) 'まだ実行していない/'
        Assert-Equal '次回の実行日時' (Get-TaskField $x.Tasks 'daily' '次回の実行日時') '2099/01/01 03:00:00'
        Assert-Equal '無効のタスクの次回の実行日時は空欄' (Get-TaskField $x.Tasks 'weekly' '次回の実行日時') ''
        if ($ran) {
            Assert-Equal '失敗' ('{0}/{1}' -f (Get-TaskField $x.Tasks 'fail' '前回の結果'), (Get-TaskField $x.Tasks 'fail' '結果のコード')) '失敗/0x00000001'
            Assert-Equal '成功' ('{0}/{1}' -f (Get-TaskField $x.Tasks 'ok' '前回の結果'), (Get-TaskField $x.Tasks 'ok' '結果のコード')) '成功/0x00000000'
            Assert-Equal '前回の実行日時' ((Get-TaskField $x.Tasks 'ok' '前回の実行日時') -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True'
            Assert-Equal '無効にしたタスクの前回の結果は、そのまま書く' ('{0}/{1}/{2}' -f (Get-TaskField $x.Tasks 'offfail' '状態'), (Get-TaskField $x.Tasks 'offfail' '前回の結果'), (Get-TaskField $x.Tasks 'offfail' '結果のコード')) '無効/失敗/0x00000001'
        } else {
            Write-Skip '前回の結果（成功・失敗）' 'テスト用のタスクを実行できなかったため'
        }
        Assert-Equal '実行ユーザー' ((Get-TaskField $x.Tasks 'daily' '実行ユーザー') -match ('^(.+\\)?' + [regex]::Escape($env:USERNAME) + '$')) 'True'
        Assert-Equal '実行の条件' (Get-TaskField $x.Tasks 'daily' '実行の条件') 'ログオンしているときのみ'
        Assert-Equal '最上位の特権' (Get-TaskField $x.Tasks 'daily' '最上位の特権') 'いいえ'
        if ($isAdmin) {
            Assert-Equal 'SYSTEM・サービス アカウント・最上位の特権' ('{0}/{1}/{2}' -f ((Get-TaskField $x.Tasks 'sub\system' '実行ユーザー') -match 'SYSTEM$'), (Get-TaskField $x.Tasks 'sub\system' '実行の条件'), (Get-TaskField $x.Tasks 'sub\system' '最上位の特権')) 'True/サービス アカウント/はい'
            Assert-Equal 'パスワードを保存しない（S4U）' (Get-TaskField $x.Tasks 'sub\s4u' '実行の条件') 'パスワードを保存しない'
        } else {
            Write-Skip 'SYSTEM・最上位の特権・パスワードを保存しない' '管理者でないと登録できないため'
        }
        Assert-Equal '操作' (Get-TaskField $x.Tasks 'fail' '操作') ('wscript.exe //B //Nologo "{0}"' -f $failScript)
        Assert-Equal '複数の操作と COM ハンドラー' (Get-TaskField $x.Tasks 'actions' '操作') 'COM ハンドラー / C:\Program Files\Example\tool.exe -a "b c" / cmd.exe /c exit 0'
        Assert-Equal '作業フォルダ（指定した操作の分だけ）' ('{0}|{1}' -f (Get-TaskField $x.Tasks 'actions' '作業フォルダ'), (Get-TaskField $x.Tasks 'daily' '作業フォルダ')) 'D:\work|C:\作業 フォルダ'
        Assert-Equal '作成者' (Get-TaskField $x.Tasks 'daily' '作成者') 'workbench-tests'
        Assert-Equal '作成日' ('{0}|{1}' -f (Get-TaskField $x.Tasks 'daily' '作成日'), (Get-TaskField $x.Tasks 'ok' '作成日')) '2026/10/31 09:15:00|'
        Assert-Equal '隠し' ('{0}/{1}' -f (Get-TaskField $x.Tasks '日本語 タスク' '隠し'), (Get-TaskField $x.Tasks 'daily' '隠し')) 'はい/'

        Write-Section 'Excel 対策（="..."）'
        Assert-Equal '日付に見える説明' (Get-TaskField $x.Tasks 'daily' '説明') '="2026-10-31"'
        Assert-Equal '全角の数字の説明' (Get-TaskField $x.Tasks 'actions' '説明') '="１．５"'
        Assert-Equal '+ で始まる作成者' (Get-TaskField $x.Tasks '-dash' '作成者') '="+81 3"'
        Assert-Equal '255 文字を超える = で始まる説明は、先頭に空白' (Get-TaskField $x.Tasks '-dash' '説明') (' =' + ('x' * 299))
        Assert-Equal '- で始まる名前' @($x.Tasks | Where-Object { $_.'タスク名' -ceq '="-dash"' }).Count 1
        Assert-Equal 'フォルダや XML ファイルの列は、- で始まらなければそのまま' ('{0}|{1}' -f @(Get-TaskRow $x.Tasks '-dash')[0].'フォルダ', (Get-TaskField $x.Tasks '-dash' 'XML ファイル')) ('{0}|{1}' -f $testDir, (Get-XmlRel '-dash'))

        Write-Section '画面の表示（色と、無効のタスクの -）'
        $colorScript = New-HelperScript 'colors.ps1' @(
            'param($Tool, $TaskPath)'
            '$ps = [powershell]::Create()'
            '[void]$ps.AddCommand($Tool).AddParameter(''Path'', $TaskPath).AddParameter(''NoSave'', $true)'
            '$null = $ps.Invoke()'
            'foreach ($r in $ps.Streams.Information) { $m = $r.MessageData; if ($m -is [Management.Automation.HostInformationMessage]) { ''{0}|{1}'' -f $m.ForegroundColor, $m.Message } }'
        )
        $screen = @(& $Exe -NoProfile -ExecutionPolicy Bypass -File $colorScript -Tool $tool -TaskPath $testDir 2>&1 | ForEach-Object { "$_" })
        $rowPattern = '^(?<color>\w*)\| (?<state>\S+)\s+(?<result>.+?)\s{2,}(?<next>\d{4}-\d{2}-\d{2} \d{2}:\d{2}|-)\s+(?<user>\S+)\s+'
        function Get-ScreenCells([string]$Name) {
            $line = @($screen | Where-Object { $_.EndsWith('  ' + (Get-TaskPath $Name)) }) | Select-Object -First 1
            $m = [regex]::Match([string]$line, $rowPattern)
            if (-not $m.Success) { return '(行がない)' }
            '{0}|{1}|{2}|{3}' -f $m.Groups['color'].Value, $m.Groups['state'].Value, $m.Groups['result'].Value, $m.Groups['next'].Value
        }
        if ($ran) {
            Assert-Equal '前回失敗は黄色' (Get-ScreenCells 'fail') 'Yellow|準備完了|失敗 (0x1)|2099-01-01 03:00'
            Assert-Equal '成功は色なし' (Get-ScreenCells 'ok') '|準備完了|成功 (0x0)|2099-01-01 03:00'
            Assert-Equal '無効は灰色で、前回失敗でも - にする' (Get-ScreenCells 'offfail') 'DarkGray|無効|-|-'
        }
        Assert-Equal 'まだ実行していない' (Get-ScreenCells 'daily') '|準備完了|まだ実行していない|2099-01-01 03:00'
        Assert-Equal '無効（まだ実行していない）' (Get-ScreenCells 'weekly') 'DarkGray|無効|-|-'
        Assert-Equal '次回の実行がないタスク' (Get-ScreenCells 'event') '|準備完了|まだ実行していない|-'
        Assert-Equal '見出しの行' (@($screen | Where-Object { $_ -match '^\| 状態\s+前回の結果\s+次回の実行\s+実行ユーザー\s+タスク$' }).Count) 1

        Write-Section 'XML'
        $xmlFile = Join-Path $run.Dir ('xml\' + (Get-XmlRel 'fail'))
        Assert-Equal 'タスクのフォルダの構成で保存' (Test-Path -LiteralPath $xmlFile) 'True'
        Assert-Equal 'XML ファイルの列' (Get-TaskField $x.Tasks 'fail' 'XML ファイル') (Get-XmlRel 'fail')
        $bytes = if (Test-Path -LiteralPath $xmlFile) { [IO.File]::ReadAllBytes($xmlFile) } else { [byte[]](0, 0) }
        Assert-Equal 'UTF-16（BOM 付き）' ('{0:X2}{1:X2}' -f $bytes[0], $bytes[1]) 'FFFE'
        $doc = New-Object Xml.XmlDocument
        $loaded = $true
        try { $doc.Load($xmlFile) } catch { $loaded = $false }
        Assert-Equal 'XML として読める' $loaded 'True'
        Assert-Equal 'XML の中身は元のタスク' $(if ($loaded) { $doc.Task.Actions.Exec.Arguments } else { '' }) ('//B //Nologo "{0}"' -f $failScript)
        $xmlCount = @(Get-ChildItem -LiteralPath (Join-Path $run.Dir 'xml') -Recurse -File -Filter '*.xml').Count
        Assert-Equal 'XML の数は tasks.csv の行の数' $xmlCount $rowCount
        Assert-Equal 'XML ファイルの列が空欄の行はない' @($x.Tasks | Where-Object { -not $_.'XML ファイル' }).Count 0
        # 末尾の . と空白はファイル名にできず消すため、名前が重なる。重なったら _2 などを付ける
        $bRels = @('b', 'b ', 'b.') | ForEach-Object { Get-TaskField $x.Tasks $_ 'XML ファイル' }
        Assert-Equal 'XML の名前が重なったら _2・_3' ($bRels -join ',') ('{0}\b.xml,{0}\b_2.xml,{0}\b_3.xml' -f $testRel)
        Assert-Equal 'XML の名前が重なっても、それぞれのタスクの中身' (@($bRels | ForEach-Object {
            $file = Join-Path $run.Dir ('xml\' + $_)
            if ($_ -and (Test-Path -LiteralPath $file)) { ([xml][IO.File]::ReadAllText($file)).Task.RegistrationInfo.URI } else { '(なし)' }
        }) -join ',') ((@('b', 'b ', 'b.') | ForEach-Object { Get-TaskPath $_ }) -join ',')
        # 保存した XML を、schtasks で取り込み直せるか
        $null = & schtasks.exe /create /xml $xmlFile /tn (Get-TaskPath 'reimported') 2>&1
        Assert-Equal 'schtasks /create /xml で取り込める' $LASTEXITCODE 0
        $imported = $null
        try { $imported = Get-TestTask 'reimported' } catch { }
        Assert-Equal '取り込んだタスクの操作' $(if ($imported) { $imported.Definition.Actions.Item(1).Arguments } else { '(なし)' }) ('//B //Nologo "{0}"' -f $failScript)
        # 後のテストの件数が変わらないよう、取り込んだタスクは消しておく
        if ($imported) { Remove-TestTask $testFolder $imported }

        Write-Section '-NoXml・-NoSave'
        $run = Invoke-Tool 'noxml' @('-Path', $testDir, '-NoXml')
        Assert-Equal '-NoXml: 保存するファイル' ('{0}/{1}' -f $run.Code, (Get-ResultNames $run.Dir)) '0/skipped.csv,summary.txt,tasks.csv'
        Assert-Equal '-NoXml: XML ファイルの列は空欄' (Get-TaskField (Read-Result $run.Dir).Tasks 'fail' 'XML ファイル') ''
        Assert-Equal '-NoXml: 表示（範囲の次の行）' $run.Log.Contains("範囲: $testDir の下" + 'XML: 保存しない') 'True'
        Assert-Equal '-NoXml: summary にも' (Get-Line (Read-Result $run.Dir).Lines 3) 'XML: 保存しない'
        $noSaveOut = Join-Path $workRoot 'out\nosave'
        $run = Invoke-Plain @('-Path', $testDir, '-NoSave', '-OutputFolder', $noSaveOut)
        Assert-Equal '-NoSave: 保存しない' ('{0}/{1}/{2}' -f $run.Code, (Test-Path -LiteralPath $noSaveOut), $run.Log.Contains('保存先')) '0/False/False'
        Assert-Equal '-NoSave: 一覧と件数を表示する' ('{0}/{1}' -f $run.Log.Contains((Get-TaskPath 'fail')), (Get-CountText $run.Log)) ('True/' + $exactCount)
        Assert-Equal '-NoSave: XML の行は出さない' $run.Log.Contains('XML:') 'False'

        Write-Section '範囲（-Path と -All）'
        $run = Invoke-Tool 'default' @()
        $x = Read-Result $run.Dir
        Assert-Equal '既定: テスト用のタスクを含む' ('{0}/{1}' -f $run.Code, @(Get-TaskRow $x.Tasks 'daily').Count) '0/1'
        Assert-Equal '既定: 範囲の行' $run.Log.Contains('範囲: \ の下（\Microsoft\ の下を除く）') 'True'
        Assert-Equal '既定: \Microsoft\ の下を含まない' @($x.Tasks | Where-Object { $_.'フォルダ' -like '\Microsoft*' }).Count 0
        Assert-Equal '既定: 件数の行は tasks.csv と合う' (Get-CountText $run.Log) (Get-ExpectedCount $x.Tasks)
        # xml\ から見たパスが - で始まるのは、ルートのタスクだけ
        $rootRow = @($x.Tasks | Where-Object { $_.'フォルダ' -eq '\' -and $_.'タスク名' -ceq ('="{0}"' -f $rootName) })
        Assert-Equal 'XML ファイルの列も ="..." で書く（- で始まるとき）' ('{0}/{1}' -f $rootRow.Count, $(if ($rootRow.Count -gt 0) { $rootRow[0].'XML ファイル' } else { '' })) ('1/="{0}.xml"' -f $rootName)
        $run = Invoke-Tool 'all' @('-All', '-NoXml')
        $x = Read-Result $run.Dir
        $allCount = $x.Tasks.Count
        Assert-Equal '-All: \Microsoft\ の下を含む' ('{0}/{1}' -f $run.Code, (@($x.Tasks | Where-Object { $_.'フォルダ' -like '\Microsoft\*' }).Count -gt 0)) '0/True'
        Assert-Equal '-All: 範囲の行' $run.Log.Contains('範囲: \ の下（Windows 標準のタスクを含む）') 'True'
        Assert-Equal '-All: テスト用のタスクも含む' @(Get-TaskRow $x.Tasks 'daily').Count 1
        Assert-Equal '-All: 件数の行は tasks.csv と合う' (Get-CountText $run.Log) (Get-ExpectedCount $x.Tasks)
        # -All のときは 200 件を超えることが多い。保存したときは、画面は 200 件まで。summary.txt と -NoSave はすべて
        if ($allCount -gt 200) {
            $lastPath = Get-RowPath $x.Tasks[-1]
            Assert-Equal '-All: 画面は 200 件まで' ('{0}/{1}/{2}' -f (Get-TableRowCount $run.Lines), $run.Log.Contains(('（ほか {0:N0} 件は tasks.csv を参照してください）' -f ($allCount - 200))), $run.Log.Contains($lastPath)) '200/True/False'
            Assert-Equal '-All: summary.txt はすべて' ('{0}/{1}' -f @($x.Lines | Where-Object { $_.EndsWith('  ' + $lastPath) }).Count, @($x.Lines | Where-Object { $_ -like '（ほか*' }).Count) '1/0'
            $run = Invoke-Plain @('-All', '-NoSave')
            Assert-Equal '-All -NoSave: 画面にすべて' ('{0}/{1}' -f $run.Log.Contains('（ほか'), $run.Log.Contains($lastPath)) 'False/True'
        } else {
            Write-Skip '画面は 200 件まで' ('-All でもタスクが {0} 件で、200 件以下のため' -f $allCount)
        }
        # -Path で \Microsoft\ の下を指定したときは、-All がなくても、その下のフォルダまで出力する
        $msOk = $true
        try { [void]$service.GetFolder('\Microsoft\Windows') } catch { $msOk = $false }
        if ($msOk) {
            foreach ($case in @(@('\Microsoft\Windows', '\Microsoft\Windows'), @('/microsoft/windows/', '\microsoft\windows'))) {
                $run = Invoke-Plain @('-Path', $case[0], '-NoSave')
                Assert-Equal ('-Path で \Microsoft\ の下を指定すると出力する: {0}' -f $case[0]) ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains(('範囲: {0} の下' -f $case[1])), $run.Log.Contains('範囲: \ '), $run.Log.Contains('タスク 0 件')) '0/True/False/False'
            }
        } else {
            Write-Skip '-Path で \Microsoft\ の下' '\Microsoft\Windows がないため'
        }
        # \ で始めない・/ で区切る・末尾に \ を付ける、も受け付ける
        $run = Invoke-Plain @('-Path', ($testRel.Replace('\', '/') + '/'), '-NoSave')
        Assert-Equal '-Path の書き方（\ で始めない・/・末尾の /）' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains("範囲: $testDir の下"), (Get-CountText $run.Log)) ('0/True/' + $exactCount)
        $run = Invoke-Plain @($testDir, '-NoSave')
        Assert-Equal '-Path は名前を付けずに書ける' ('{0}/{1}' -f $run.Code, (Get-CountText $run.Log)) ('0/' + $exactCount)
        $run = Invoke-Plain @('-Path', "$testDir\sub x", '-NoSave')
        Assert-Equal '-Path にタスクを指定（フォルダではない）' ('{0}/{1}' -f $run.Code, $run.Log.Contains("タスクのフォルダが見つかりません: $testDir\sub x（")) '1/True'
        [void]$testFolder.CreateFolder('empty')
        $run = Invoke-Plain @('-Path', "$testDir\empty", '-NoSave')
        Assert-Equal 'タスクのないフォルダ' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('（タスクはありません）'), $run.Log.Contains('状態  '), (Get-CountText $run.Log)) '0/True/False/タスク 0 件（有効 0・無効 0）'
        $testFolder.DeleteFolder('empty', 0)

        Write-Section '保存'
        $jpOut = Join-Path $workRoot 'out\保存 先'
        $run = Invoke-Plain @('-Path', $testDir, '-OutputFolder', $jpOut)
        Assert-Equal '日本語とスペースを含む保存先' ('{0}/{1}' -f $run.Code, (Get-ResultNames (Get-RunDir $jpOut))) '0/skipped.csv,summary.txt,tasks.csv,xml'
        # 通常の書き方では末尾のピリオドは無視される。\\?\ 付きで作ると、ピリオドの付いたフォルダができてしまう
        $run = Invoke-Plain @('-Path', $testDir, '-NoXml', '-OutputFolder', (Join-Path $workRoot 'out\trail.'))
        Assert-Equal '末尾にピリオドが付いた保存先' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultNames (Get-RunDir (Join-Path $workRoot 'out\trail'))), [IO.Directory]::Exists('\\?\' + (Join-Path $workRoot 'out\trail.'))) '0/skipped.csv,summary.txt,tasks.csv/False'
        # 同じ名前の結果のフォルダがあるときは、_2 を付ける（これから 20 秒の時刻の名前のフォルダを、先に作っておく）
        $twiceOut = Join-Path $workRoot 'out\twice'
        $now = Get-Date
        foreach ($s in 0..20) { [void][IO.Directory]::CreateDirectory((Join-Path $twiceOut ('TaskList_{0}_{1:yyyyMMdd_HHmmss}' -f $env:COMPUTERNAME, $now.AddSeconds($s)))) }
        $run = Invoke-Plain @('-Path', $testDir, '-NoXml', '-OutputFolder', $twiceOut)
        $made = @(Get-ChildItem -LiteralPath $twiceOut -Directory | Where-Object { $_.Name -like '*_2' })
        Assert-Equal '同じ名前の結果のフォルダがあるときは _2' ('{0}/{1}/{2}' -f $run.Code, $made.Count, $(if ($made.Count -gt 0) { Get-ResultNames $made[0].FullName } else { '(なし)' })) '0/1/skipped.csv,summary.txt,tasks.csv'
        $freeLetter = Get-FreeLetter
        if ($freeLetter) {
            $run = Invoke-Plain @('-Path', $testDir, '-OutputFolder', "${freeLetter}:\nope")
            Assert-Equal '書き込めない保存先: 警告して、画面に表示する' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains("保存先に書き込めないため、結果は保存しません: ${freeLetter}:\nope"), (Get-CountText $run.Log), $run.Log.Contains('保存先: ')) ('0/True/{0}/False' -f $exactCount)
        } else {
            Write-Skip '書き込めない保存先' '空いているドライブ文字がないため'
        }
        # 保存先のフォルダは作れるが、その中にファイルを作れないようにして、保存に失敗させる
        $guardOut = Join-Path $workRoot 'out\save-fail'
        [void][IO.Directory]::CreateDirectory($guardOut)
        if (Add-Deny $guardOut '(OI)(CI)(WD)') {
            $run = Invoke-Plain @('-Path', $testDir, '-NoXml', '-OutputFolder', $guardOut)
            Assert-Equal '保存に失敗: 警告して、画面に表示する' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした'), $run.Log.Contains('保存先:'), (Get-CountText $run.Log)) ('0/True/False/' + $exactCount)
            Assert-Equal '保存に失敗: 書きかけの保存先を残さない' @(Get-ChildItem -LiteralPath $guardOut -Force).Count 0
            # XML も保存できないが、skipped.csv も残らないため、XML の注意は出さない
            $run = Invoke-Plain @('-Path', $testDir, '-OutputFolder', $guardOut)
            Assert-Equal '保存に失敗（XML も）: skipped.csv を見るようには書かない' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('結果を保存できませんでした'), $run.Log.Contains('skipped.csv')) '0/True/False'
        } else {
            Write-Skip '保存に失敗したとき' 'アクセス拒否を設定できなかったため'
        }
        # 既定の保存先は、スクリプトと同じフォルダ（別のフォルダから起動しても）
        $cwd = Join-Path $workRoot 'cwd'
        [void][IO.Directory]::CreateDirectory($cwd)
        $run = Invoke-From $cwd @('-Path', $testDir, '-NoXml')
        Assert-Equal '既定の保存先はスクリプトのフォルダ' ('{0}/{1}/{2}' -f $run.Code, (Get-ResultNames (Get-RunDir (Split-Path $tool))), @(Get-ChildItem -LiteralPath $cwd).Count) '0/skipped.csv,summary.txt,tasks.csv/0'
        $run = Invoke-From $cwd @('-Path', $testDir, '-NoXml', '-OutputFolder', 'rel')
        Assert-Equal '相対パスの -OutputFolder は、起動したフォルダが基準' ('{0}/{1}' -f $run.Code, (Get-ResultNames (Get-RunDir "$cwd\rel"))) '0/skipped.csv,summary.txt,tasks.csv'

        Write-Section '260 文字を超えるパス'
        $longOut = "$workRoot\" + ('L' * 150) + '\' + ('M' * 150)
        $run = Invoke-Plain @('-Path', $testDir, '-OutputFolder', $longOut)
        $longRun = @(Get-ChildItem -LiteralPath ('\\?\' + $longOut) -Directory -ErrorAction SilentlyContinue)
        $longNames = if ($longRun.Count -gt 0) { Join-Sorted (Get-ChildItem -LiteralPath $longRun[0].FullName | ForEach-Object { $_.Name }) } else { '(なし)' }
        Assert-Equal '-OutputFolder' ('{0}/{1}/{2}' -f $run.Code, $longNames, $run.Log.Contains('保存先に書き込めない')) '0/skipped.csv,summary.txt,tasks.csv,xml/False'
        # 5.1 の Join-Path は \\?\ で始まるパスを扱えないため、文字列でつなぐ
        $longXml = if ($longRun.Count -gt 0) { @(Get-ChildItem -LiteralPath ($longRun[0].FullName + '\xml') -Recurse -File -ErrorAction SilentlyContinue).Count } else { 0 }
        Assert-Equal '-OutputFolder: XML もすべて保存する' ('{0}/{1}' -f $longXml, $run.Log.Contains('XML を保存できなかった')) ('{0}/False' -f $rowCount)

        Write-Section '名前に [ ] を含むフォルダから起動（-OutputFolder を相対パスで指定）'
        # 5.1 は、名前に [ ] を含むフォルダで起動すると、名前をワイルドカードとして扱い、別の場所から始まる
        $startRoot = Join-Path $workRoot 'start'
        $startCases = @(
            @('一致するフォルダがない（$PSHOME から始まる）', 'br[1] x', @()),
            @('隣に 1 つ一致するフォルダがある', 'sib\[a]', @('sib\a')),
            @('[ が閉じていない（ドライブのルートから始まる）', 'ub\a[b', @('ub\ab')),
            @('隣に 2 つ一致する（ドライブのルートから始まる）', 'mm\[ab]', @('mm\a', 'mm\b'))
        )
        foreach ($case in $startCases) {
            $dir = "$startRoot\$($case[1])"
            [void][IO.Directory]::CreateDirectory($dir)
            foreach ($decoy in $case[2]) { [void][IO.Directory]::CreateDirectory("$startRoot\$decoy") }
            $run = Invoke-From $dir @('-Path', $testDir, '-NoXml', '-OutputFolder', 'out')
            $decoyOut = @($case[2] | Where-Object { Test-Path -LiteralPath "$startRoot\$_\out" }).Count
            Assert-Equal $case[0] ('{0}/{1}/{2}' -f $run.Code, (Get-ResultNames (Get-RunDir "$dir\out")), $decoyOut) '0/skipped.csv,summary.txt,tasks.csv/0'
        }

        Write-Section 'コマンドプロンプトから起動'
        $cmdBase = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $Exe, $tool
        $run = Invoke-ViaCmd ('{0} -NoSave -Path "{1}\"' -f $cmdBase, $testDir)
        Assert-Equal '最後の引数が "パス\" の -Path' ('{0}/{1}' -f $run.Code, (Get-CountText $run.Log)) ('0/' + $exactCount)
        $run = Invoke-ViaCmd ('{0} -Path "{1}\" -NoSave' -f $cmdBase, $testDir)
        Assert-Equal '"パス\" の後ろに引数: ヒントを出して止まる（%ERRORLEVEL% は 1）' ('{0}/{1}/{2}/{3}' -f $run.Code, $run.Log.Contains('-Path の値に引用符'), $run.Log.Contains('最後に \ を付けないでください'), $run.Log.Contains('タスクスケジューラ出力 (')) '1/True/True/False'
        $cmdOut = Join-Path $workRoot 'out\cmd'
        $run = Invoke-ViaCmd ('{0} -Path "{1}" -OutputFolder "{2}\" -NoXml' -f $cmdBase, $testDir, $cmdOut)
        Assert-Equal '"-OutputFolder\" の後ろに引数: ヒントを出して止まる' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('-OutputFolder の値に引用符'), (Test-Path -LiteralPath $cmdOut)) '1/True/False'
        $run = Invoke-ViaCmd ('{0} -Path ''{1}'' -NoXml -OutputFolder ''{2}''' -f $cmdBase, $testDir, $cmdOut)
        Assert-Equal 'cmd で '' で囲んだ -Path と -OutputFolder' ('{0}/{1}' -f $run.Code, (Get-ResultNames (Get-RunDir $cmdOut))) '0/skipped.csv,summary.txt,tasks.csv'

        Write-Section '実行中のタスク'
        if ($hasWscript) {
            $running = $false
            try {
                Register-TestTask 'running' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start; foreach ($a in $d.Actions) { $a.Arguments = '//B //Nologo "{0}"' -f $sleepScript } }
                [void](Get-TestTask 'running').Run($null)
                for ($i = 0; $i -lt 30 -and -not $running; $i++) { Start-Sleep -Milliseconds 500; $running = (Get-TestTask 'running').State -eq 4 }
            } catch { }
            if ($running) {
                $run = Invoke-Tool 'running' @('-Path', $testDir, '-NoXml')
                $x = Read-Result $run.Dir
                Assert-Equal '実行中のタスク' ('{0}/{1}/{2}' -f (Get-TaskField $x.Tasks 'running' '状態'), (Get-TaskField $x.Tasks 'running' '前回の結果'), (Get-TaskField $x.Tasks 'running' '結果のコード')) '実行中/実行中/0x00041301'
                Assert-Equal '実行中のタスクは失敗に数えない' (Get-CountText $run.Log) (Get-ExactCount 1)
                Assert-Equal '画面の表示' ([regex]::IsMatch($run.Log, ' 実行中\s+実行中\s+2099-01-01 03:00\s+\S+\s+' + [regex]::Escape((Get-TaskPath 'running')))) 'True'
            } else {
                Write-Skip '実行中のタスク' '実行中にならなかったため'
            }
            try { Remove-TestTask $testFolder (Get-TestTask 'running') } catch { }
        } else {
            Write-Skip '実行中のタスク' 'wscript.exe がないため'
        }

        Write-Section 'XML を保存できないタスク'
        # フォルダ b.xml の中のタスクの XML は、タスク b の XML（b.xml）と重なって保存できない
        $xmlFailOk = $true
        try {
            Register-TestTask 'xmlfail\b' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start }
            Register-TestTask 'xmlfail\b.xml\in' { param($d) $t = $d.Triggers.Create(1); $t.StartBoundary = $start }
        } catch { $xmlFailOk = $false }
        if ($xmlFailOk) {
            $run = Invoke-Tool 'xmlfail' @('-Path', "$testDir\xmlfail")
            $x = Read-Result $run.Dir
            $note = '※ XML を保存できなかったタスクが 1 件あります（skipped.csv を確認してください）'
            Assert-Equal '一覧には載せる（XML ファイルの列は空欄）' ('{0}/{1}/{2}' -f $run.Code, (Get-TaskField $x.Tasks 'xmlfail\b' 'XML ファイル'), (Get-TaskField $x.Tasks 'xmlfail\b.xml\in' 'XML ファイル')) ('0/{0}\xmlfail\b.xml/' -f $testRel)
            Assert-Equal '表示（スキップの件数には含めない）' ('{0}/{1}' -f $run.Log.Contains($note + '保存先: '), $run.Log.Contains('スキップ ')) 'True/False'
            Assert-Equal 'summary にも' @($x.Lines | Where-Object { $_ -eq $note }).Count 1
            Assert-Equal 'skipped.csv' ((@($x.Skipped | ForEach-Object { '{0}|{1}|{2}' -f $_.'理由', $_.'パス', $_.'詳細'.StartsWith('XML を保存できません: ') })) -join ',') ("その他のエラー|$testDir\xmlfail\b.xml\in|True")
        } else {
            Write-Skip 'XML を保存できないタスク' 'テスト用のタスクを登録できなかったため'
        }

        # 読めないタスク（権限なし）は再現できないため、試さない。タスクスケジューラは、読めないタスクを一覧から除く。
        # 読み取りを拒否して登録しようとすると、管理者でも登録の時点で拒否される（タスクのファイルだけが残る）。
        # 管理者でないときは、登録したタスクやフォルダのアクセス拒否の変更も、拒否されるのに途中まで効いて、消せなくなる

        Write-Section 'Ctrl+C で中断（読んでいる途中で停止）'
        $stopScript = New-HelperScript 'stop-during-read.ps1' @(
            'param($Tool, $OutRoot)'
            '$ps = [powershell]::Create()'
            '[void]$ps.AddCommand($Tool).AddParameter(''All'', $true).AddParameter(''OutputFolder'', $OutRoot)'
            '$handle = $ps.BeginInvoke()'
            'for ($i = 0; $i -lt 100 -and @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count -eq 0; $i++) { Start-Sleep -Milliseconds 20 }'
            'Start-Sleep -Milliseconds 300'
            '''RUNNING='' + (-not $handle.IsCompleted)'
            '''BEFORE='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
            '$ps.Stop()'
            '''AFTER='' + @(Get-ChildItem -LiteralPath $OutRoot -Directory -ErrorAction SilentlyContinue).Count'
            '''MESSAGE='' + @($ps.Streams.Information | Where-Object { $_.MessageData.ToString() -like ''中断*'' } | ForEach-Object { $_.MessageData.ToString() })'
        )
        $result = Invoke-Helper $stopScript @('-Tool', $tool, '-OutRoot', (Join-Path $workRoot 'out\ctrlc'))
        if ($result['RUNNING'] -eq 'True') {
            Assert-Equal '止める前は保存先を作ってある・止めた後は残らない' ('{0}/{1}' -f $result['BEFORE'], $result['AFTER']) '1/0'
            Assert-Equal '中断のメッセージ' $result['MESSAGE'] '中断しました。結果は保存していません。'
        } else {
            Write-Skip 'Ctrl+C で中断' '止める前に読み終わったため'
        }

        Write-Section '地域設定'
        # 日付と数値の書式が、OS の地域設定によって変わらないこと
        $probe   = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; (Get-Date '2026-10-29 08:05:00').ToString('yyyy/MM/dd HH:mm:ss') + '|' + (1.5).ToString('N1')"
        $control = "$(& $Exe -NoProfile -Command $probe)"
        if ($control -and $control -ne '2026/10/29 08:05:00|1.5') {
            $cultureOut = Join-Path $workRoot 'out\culture'
            $command = "[Threading.Thread]::CurrentThread.CurrentCulture = 'fi-FI'; & '{0}' -Path '{1}' -NoXml -OutputFolder '{2}' | Out-Null" -f $tool.Replace("'", "''"), $testDir, $cultureOut.Replace("'", "''")
            $null = & $Exe -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
            Assert-Equal 'fi-FI で実行して保存できた' ([bool](Get-RunDir $cultureOut)) 'True'
            $x = Read-Result (Get-RunDir $cultureOut)
            $summaryText = $x.Lines -join "`n"
            Assert-Equal 'fi-FI でも対象の行の日時は yyyy-MM-dd HH:mm:ss' ([regex]::IsMatch($summaryText, '対象: .*   \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}')) 'True'
            Assert-Equal 'fi-FI でも所要時間の小数点は .' ([regex]::IsMatch($summaryText, '  所要 \d+\.\d秒')) 'True'
            Assert-Equal 'fi-FI でも次回の実行は yyyy-MM-dd HH:mm' ([regex]::IsMatch($summaryText, '2099-01-01 03:00  ')) 'True'
            Assert-Equal 'fi-FI でも tasks.csv の日時は yyyy/MM/dd HH:mm:ss' (Get-TaskField $x.Tasks 'daily' '次回の実行日時') '2099/01/01 03:00:00'
            if ($ran) { Assert-Equal 'fi-FI でも前回の実行日時は yyyy/MM/dd HH:mm:ss' ((Get-TaskField $x.Tasks 'ok' '前回の実行日時') -match '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') 'True' }
        } else {
            Write-Skip '地域設定' 'fi-FI に切り替えられなかったため'
        }

        Write-Section 'レビューで見つかった問題の再発防止'
        # 1. プロパティを読むときの権限エラーも「権限なし」に分ける（読めないタスクを再現できないため、確かめられない）
        # 2. XML の名前が重なったら _2 を付け、1 件保存できなくても続ける（上の「XML」と「XML を保存できないタスク」で確かめている）
        # 3. XML ファイルの列も Excel 対策する（上の「範囲」の、ルートのタスクで確かめている）
        # 4. -Path を開けないときと見つからないときの表示を分ける（見つからないときは「止まるとき」で確かめている。
        #    開けないときは、読み取りを拒否したフォルダでもタスクスケジューラが一覧を返すため、再現できない）
        # 5. -Path を指定したときの範囲の表示（上の「範囲」で確かめている）
        # 6. XML を保存できなかったタスクは、スキップの件数に含めない（上の「XML を保存できないタスク」で確かめている）
        # 7. Windows でのテストで見つけた問題: 5.1 で 260 文字を超える -OutputFolder に保存できなかった、保存に失敗したときも
        #    skipped.csv を見るように表示していた、無効のタスクの次回の実行日時を tasks.csv に書いていた（上でそれぞれ確かめている）
        '  （ほかの項目で確かめている）'
    }

    Write-Section '止まるとき'
    $run = Invoke-Plain @('-Path', '\no\such\folder', '-NoSave')
    Assert-Equal '存在しないフォルダ' ('{0}/{1}/{2}' -f $run.Code, $run.Log.Contains('タスクのフォルダが見つかりません: \no\such\folder（'), $run.Log.Contains('タスクスケジューラ出力 (')) '1/True/False'
    $run = Invoke-Plain @('-Path', 'no-such-folder', '-NoSave')
    Assert-Equal '\ で始めないフォルダ' ('{0}/{1}' -f $run.Code, $run.Log.Contains('タスクのフォルダが見つかりません: \no-such-folder（')) '1/True'

    Write-Section '管理者でないときの注意'
    $run = Invoke-Plain @('-Path', $(if ($testFolder) { $testDir } else { '\' }), '-NoSave')
    Assert-Equal '管理者でないときの注意（管理者のときは出さない）' ('{0}/{1}' -f $run.Code, $run.Log.Contains('※ 管理者として実行していないため、ほかのユーザーや SYSTEM のタスクが出ないことがあります')) ('0/{0}' -f (-not $isAdmin))
} finally {
    foreach ($path in $denied) { $null = & icacls.exe $path /remove:d "*$userSid" 2>&1 }
    if ($KeepTemp) {
        Write-Section ('結果とテスト用のタスクを残しました: {0} / {1} / \{2}' -f $workRoot, $testDir, $rootName)
    } else {
        if ($testFolder) {
            Remove-TestFolder $testDir
            Clear-TestLeftovers ''
            # ほかのテストが使っていなければ、\workbench-tests も消す
            try {
                $top = $service.GetFolder($testTop)
                if (@($top.GetFolders(0)).Count -eq 0 -and @($top.GetTasks(1)).Count -eq 0) { $service.GetFolder('\').DeleteFolder($testTop.TrimStart('\'), 0) }
            } catch { }
            $left = $true
            try { [void]$service.GetFolder($testDir) } catch { $left = $false }
            if ($left -or [IO.Directory]::Exists((Join-Path $env:windir ('System32\Tasks\' + $testRel)))) { Write-Section ('テスト用のタスクを削除できませんでした: {0}' -f $testDir) }
        }
        Remove-RootTestTasks
        if (@($service.GetFolder('\').GetTasks(1) | Where-Object { $_.Name -match $rootLike }).Count -gt 0) { Write-Section ('テスト用のタスクを削除できませんでした: \{0}' -f $rootName) }
        if (Test-Path -LiteralPath $workRoot) {
            try { [IO.Directory]::Delete('\\?\' + $workRoot, $true) } catch { Write-Section ('結果を削除できませんでした: {0}（{1}）' -f $workRoot, $_.Exception.Message) }
        }
    }
    if ($savedEncoding) { try { [Console]::OutputEncoding = $savedEncoding } catch { } }
}

exit $script:failCount
