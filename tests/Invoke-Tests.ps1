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

$totalFail = 0
foreach ($r in $results) { $totalFail += $r.Fail }
if ($totalFail -gt 0) { exit 1 }
exit 0
