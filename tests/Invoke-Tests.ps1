<#
.SYNOPSIS
    ツールの回帰テストを、Windows PowerShell 5.1 と PowerShell 7 の両方で実行します。

.DESCRIPTION
    tests フォルダにある *.Tests.ps1 を、Windows PowerShell 5.1 と PowerShell 7（入っている場合）で
    順に実行し、結果をまとめて表示します。失敗した項目があると、終了コードが 1 になります。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\tests\Invoke-Tests.ps1

.PARAMETER Name
    実行するテストの名前（例: Get-DiskUsage）。省略すると、すべて実行します。

.PARAMETER Detailed
    合格した項目も含めて、すべての結果を表示します。

.NOTES
    Copyright (c) 2026 結進舎
    Released under the MIT License.
#>
param(
    [string]$Name = '*',

    [switch]$Detailed
)

Set-StrictMode -Version 2.0

$EndMark = '  END'   # 各テストが、最後の exit の直前に出す行
$targets = @(@{ Exe = 'powershell.exe'; Label = 'Windows PowerShell 5.1' })
if (Get-Command pwsh -ErrorAction SilentlyContinue) {
    $targets += @{ Exe = 'pwsh'; Label = 'PowerShell 7' }
} else {
    Write-Warning 'PowerShell 7（pwsh）が見つからないため、Windows PowerShell 5.1 だけで実行します。'
}

$testFiles = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter "$Name.Tests.ps1" -File | Sort-Object Name)
if ($testFiles.Count -eq 0) {
    $Host.UI.WriteErrorLine(('テストが見つかりません: {0}.Tests.ps1' -f $Name))
    exit 1
}

# BOM のない .ps1 は、5.1 が Shift_JIS として読むため日本語が壊れる。フックを通らない編集（Mac のエディタなど）で落ちることがあるので、最初に確かめる
$repoRoot   = Split-Path -Parent $PSScriptRoot
$noBomFiles = @(foreach ($dir in 'tools', 'tests') {
    Get-ChildItem -LiteralPath (Join-Path $repoRoot $dir) -Filter '*.ps1' -File -Recurse | Where-Object {
        $head   = New-Object byte[] 3
        $stream = [IO.File]::OpenRead($_.FullName)
        try { $read = $stream.Read($head, 0, 3) } finally { $stream.Dispose() }
        -not ($read -eq 3 -and $head[0] -eq 0xEF -and $head[1] -eq 0xBB -and $head[2] -eq 0xBF)
    }
})
if ($noBomFiles.Count -gt 0) {
    Write-Host ''
    Write-Host '===== BOM の確認' -ForegroundColor Cyan
    foreach ($f in $noBomFiles) { Write-Host ('  FAIL BOM がありません: {0}' -f $f.FullName) -ForegroundColor Red }
}

$results = foreach ($file in $testFiles) {
    $testName = $file.Name -replace '\.Tests\.ps1$', ''
    foreach ($target in $targets) {
        Write-Host ''
        Write-Host ('===== {0} / {1}' -f $testName, $target.Label) -ForegroundColor Cyan
        $sw      = [Diagnostics.Stopwatch]::StartNew()
        # テストが途中で止まっても、それまでの結果を残すため、1 行ずつ受け取る
        $lines   = New-Object 'System.Collections.Generic.List[object]'
        $crashed = $false
        try {
            & $file.FullName -Exe $target.Exe | ForEach-Object { $lines.Add($_) }
            # テストは、例外にならずに黙って終わることがある（Get-WinEvent -ListLog を Select-Object -First で受けたときなど）。
            # そのときは合格の数が減るだけで気づけないので、テストの最後の行まで来たかを確かめる。
            # テストの外に印を置いても、黙って終わるのはテストのスクリプトだけなので、印は出てしまう
            if (@($lines | Where-Object { "$_" -ceq $EndMark }).Count -eq 0) {
                $crashed = $true
                Write-Host ('  テストが最後まで走りませんでした（例外は出ていません）。テストの exit の直前に ''{0}'' を出しているか、途中で黙って終わっていないかを確かめてください' -f $EndMark) -ForegroundColor Red
            }
        } catch {
            $crashed = $true
            Write-Host ('  テストが途中で止まりました: {0}' -f $_.Exception.Message) -ForegroundColor Red
        }
        foreach ($line in $lines) {
            $text = "$line"
            if ($text -match '^\s+FAIL') {
                Write-Host $text -ForegroundColor Red
            } elseif ($text -match '^\s+SKIP') {
                Write-Host $text -ForegroundColor Yellow
            } elseif ($Detailed -or $text.StartsWith('■')) {
                Write-Host $text
            }
        }
        $failCount = @($lines | Where-Object { "$_" -match '^\s+FAIL' }).Count
        if ($crashed) { $failCount++ }
        [pscustomobject]@{
            Test    = $testName
            Target  = $target.Label
            Pass    = @($lines | Where-Object { "$_" -match '^\s+PASS' }).Count
            Fail    = $failCount
            Skip    = @($lines | Where-Object { "$_" -match '^\s+SKIP' }).Count
            Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
        }
    }
}

Write-Host ''
Write-Host '===== 結果' -ForegroundColor Cyan
foreach ($r in $results) {
    $color = if ($r.Fail -gt 0) { 'Red' } else { 'Green' }
    Write-Host ('{0} / {1}: 合格 {2}  失敗 {3}  スキップ {4}（{5} 秒）' -f $r.Test, $r.Target, $r.Pass, $r.Fail, $r.Skip, $r.Seconds) -ForegroundColor $color
}
$bomColor = if ($noBomFiles.Count -gt 0) { 'Red' } else { 'Green' }
Write-Host ('BOM の確認: BOM のない .ps1 {0} 件' -f $noBomFiles.Count) -ForegroundColor $bomColor

$totalFail = $noBomFiles.Count
foreach ($r in $results) { $totalFail += $r.Fail }
if ($totalFail -gt 0) { exit 1 }
exit 0
