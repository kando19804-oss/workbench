<#
.SYNOPSIS
    容量を多く使っているフォルダ・ファイル・拡張子を洗い出します。

.DESCRIPTION
    指定したフォルダの配下をすべて調べ、合計サイズの大きいフォルダ、大きいファイル、
    拡張子ごとの合計を画面に表示します。結果は CSV（BOM 付き UTF-8）でも保存します。

    調べる場所のファイルを変更・削除することはありません。
    権限がなくて読めないフォルダはスキップし、件数を表示します。
    ジャンクションやシンボリックリンクはたどりません。

    起動方法:
        powershell -ExecutionPolicy Bypass -File .\Get-DiskUsage.ps1 -Path C:\

.PARAMETER Path
    調べるフォルダ。ドライブ（C:\）や共有フォルダ（\\server\share）も指定できます。
    省略すると、現在のフォルダを調べます。

.PARAMETER Depth
    フォルダの順位に含める階層の深さ（1～100）。Path の直下が 1 です。既定値は 3 です。
    集計そのものは、常に全階層が対象です。

.PARAMETER Top
    画面に表示する件数（1～1000）。フォルダ・ファイル・拡張子それぞれに適用します。
    既定値は 20 です。

.PARAMETER OlderThan
    最終更新から指定した日数以上たったファイルだけで集計します（1～36500）。
    フォルダ・ファイル・拡張子のすべての集計が対象です。

.PARAMETER Exclude
    調べないフォルダ。フォルダのフルパスに対して、ワイルドカード（*）で指定します。
    * 以外の文字（[ ] など）は、そのままの文字として扱います。
    大文字と小文字は区別しません。複数指定するときは、カンマで区切ります。

.PARAMETER OutputFolder
    結果の保存先。省略すると、スクリプトと同じフォルダに保存します。

.PARAMETER NoCsv
    結果を保存せず、画面表示だけにします。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-DiskUsage.ps1 -Path C:\

    C ドライブ全体を調べます。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-DiskUsage.ps1 -Path D:\logs -OlderThan 90

    最終更新から 90 日以上たったファイルだけで集計します。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Get-DiskUsage.ps1 -Path C:\ -Exclude "C:\Windows,*\node_modules"

    C:\Windows と、すべての node_modules フォルダを調べません。

.NOTES
    Get-DiskUsage 1.0.0
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
    [string]$Path = '.',

    [ValidateRange(1, 100)]
    [int]$Depth = 3,

    [ValidateRange(1, 1000)]
    [int]$Top = 20,

    [ValidateRange(1, 36500)]
    [int]$OlderThan,

    [string[]]$Exclude = @(),

    [string]$OutputFolder,

    [switch]$NoCsv
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$CsvFileTop    = 1000   # files.csv に残す件数
$ProgressMs    = 250    # 進捗表示の間隔（ミリ秒）
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8Bom       = New-Object Text.UTF8Encoding $true

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

function Format-Size([long]$Bytes) {
    $units = 'B', 'KB', 'MB', 'GB', 'TB', 'PB'
    $value = [double]$Bytes
    $i = 0
    while ($value -ge 1024 -and $i -lt $units.Count - 1) { $value /= 1024; $i++ }
    if ($i -eq 0) { return '{0} B' -f $Bytes }
    return [string]::Format($Invariant, '{0:N1} {1}', $value, $units[$i])
}

# 日付や数値の区切り文字が、OS の地域設定によって変わらないようにする
function Format-Invariant([string]$Format) {
    return [string]::Format($Invariant, $Format, [object[]]$args)
}

# Excel は「.1」「.001」のような値を数値に変えてしまうため、数値として読める値は ="..." の形で出す
function ConvertTo-ExcelText([string]$Text) {
    # 全角の数字も Excel は数値にするため、半角にそろえてから判定する。
    # 壊れた文字（対になっていないサロゲート）を含むと Normalize が例外を出すので、そのときはそのまま判定する
    $normalized = $Text
    try { $normalized = $Text.Normalize([Text.NormalizationForm]::FormKC) } catch { }
    $number = 0.0
    if ([double]::TryParse($normalized.TrimEnd('%'), [Globalization.NumberStyles]::Float, $Invariant, [ref]$number)) {
        return '="' + $Text + '"'
    }
    return $Text
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

# 5.1 と 7 のどちらでも BOM 付き UTF-8 で保存する（Export-Csv の -Encoding UTF8 は 7 だと BOM なしになる）。
# 行が 0 件のときは見出しだけを書く
function Save-Csv($Rows, [string[]]$Columns, [string]$FilePath) {
    if ($null -eq $Rows) { $Rows = @() }
    $lines = @($Rows | ConvertTo-Csv -NoTypeInformation)
    if ($lines.Count -eq 0) { $lines = @(($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') }
    [IO.File]::WriteAllLines($FilePath, [string[]]$lines, $Utf8Bom)
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

#endregion

#region 引数の確認

$Path = Remove-Quote $Path
if ($Path.Contains('"')) { Stop-WithQuoteError '-Path' $Path }
if ($Path -match '^(\\\\\?\\)?[A-Za-z]:$') { $Path += '\' }

# -File で起動すると配列を渡せないため、カンマ区切りの文字列も受け付ける。
# コマンドプロンプトでは ' がそのまま届くため、'*\a,*\b' のように全体を囲んだ場合も、各要素の前後の ' を外す
# ワイルドカードは * だけにして、それ以外の文字は文字どおりに扱う。5.1 の WildcardPattern.Escape は
# バッククォートをエスケープしないため、自分でエスケープしてから照合用のパターンを作る
$wildcardOptions = [Management.Automation.WildcardOptions]'IgnoreCase, CultureInvariant'
$excludeRules = @(
    foreach ($item in $Exclude) {
        foreach ($part in ($item -split ',')) {
            $text = (Remove-Quote $part).Trim("'")
            if ($text.Contains('"')) { Stop-WithQuoteError '-Exclude' $text }
            if ($text -notmatch '^[A-Za-z]:\\$') { $text = $text.TrimEnd('\') }
            if ($text) {
                $escaped = $text.Replace('`', '``').Replace('[', '`[').Replace(']', '`]').Replace('?', '`?')
                [pscustomobject]@{ Text = $text; Matcher = [Management.Automation.WildcardPattern]::new($escaped, $wildcardOptions) }
            }
        }
    }
)
$hasExclude = $excludeRules.Count -gt 0

$outputText = $null
if ($OutputFolder) {
    $outputText = Remove-Quote $OutputFolder
    if ($outputText.Contains('"')) { Stop-WithQuoteError '-OutputFolder' $outputText }
}

$target   = $null
$resolved = $null
try { $resolved = Resolve-Path -LiteralPath $Path } catch { }
if ($resolved -and $resolved.Provider.Name -eq 'FileSystem' -and (Test-Path -LiteralPath $resolved.ProviderPath -PathType Container)) {
    $target = ConvertTo-DisplayPath $resolved.ProviderPath
    # 通常の書き方では末尾のピリオドが無視されるが、\\?\ 付きでは残って読めなくなるため、ここで正規化する
    try { $target = [IO.Path]::GetFullPath($target) } catch { }
} else {
    # 5.1 の Resolve-Path は 260 文字を超えるパスを扱えないため、\\?\ を付けて確かめ直す（\\?\ 付きの指定も受け付ける）
    try {
        $full = ConvertTo-DisplayPath $Path
        if (-not [IO.Path]::IsPathRooted($full)) { $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($full) }
        if ([IO.Directory]::Exists((ConvertTo-LongPath $full))) { $target = $full }
    } catch { }
}
if (-not $target) { Stop-WithError ('指定したフォルダが見つかりません: {0}' -f $Path) }
if ($target -notmatch '^[A-Za-z]:\\$') { $target = $target.TrimEnd('\') }

$startTime = Get-Date
$useCutoff = $PSBoundParameters.ContainsKey('OlderThan')
$cutoff    = $startTime.AddDays(-$OlderThan)
$cutoffUtc = $cutoff.ToUniversalTime()

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

#endregion

#region ドライブの使用量

$driveLine = $null
try {
    if ($target.StartsWith('\\')) {
        # DriveInfo は共有フォルダに対応していないため、FileSystemObject で取得する
        $fso   = New-Object -ComObject Scripting.FileSystemObject
        $name  = $fso.GetDriveName($target)
        $drive = $fso.GetDrive($name)
        $total = [long]$drive.TotalSize
        $free  = [long]$drive.FreeSpace
    } else {
        $drive = New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($target))
        $name  = $drive.Name.TrimEnd('\')
        $total = $drive.TotalSize
        $free  = $drive.TotalFreeSpace
    }
    $driveLine = 'ドライブ {0}  使用 {1} / 全体 {2}（空き {3}）' -f $name, (Format-Size ($total - $free)), (Format-Size $total), (Format-Size $free)
} catch { }

#endregion

#region 保存先の準備

# 調べ終わってから書き込めないと分かるのを避けるため、先に保存先のフォルダを作っておく
$runDir      = $null
$saveWarning = $null
if (-not $NoCsv) {
    $baseDir = if ($outputText) { $outputText } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    try {
        $baseDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($baseDir)
        $runName = Format-Invariant 'DiskUsage_{0}_{1:yyyyMMdd_HHmmss}' $env:COMPUTERNAME $startTime
        $runDir  = Join-Path $baseDir $runName
        # 中断や失敗のときはこのフォルダごと消すため、既存のフォルダと名前が重ならないようにする
        $suffix = 1
        while (Test-Path -LiteralPath $runDir) { $suffix++; $runDir = Join-Path $baseDir ('{0}_{1}' -f $runName, $suffix) }
        [void][IO.Directory]::CreateDirectory($runDir)
    } catch {
        $saveWarning = '保存先に書き込めないため、結果は保存しません: {0}（{1}）' -f $baseDir, (Get-ErrorText (Get-InnerException $_.Exception))
        $runDir = $null
        Write-Warning $saveWarning
    }
}

#endregion

#region 集計

$folders    = New-Object 'System.Collections.Generic.List[object]'
$skipped    = New-Object 'System.Collections.Generic.List[object]'
$candidates = New-Object 'System.Collections.Generic.List[object]'   # 大きいファイルの候補
$extTable   = @{}                                                     # 拡張子 → [サイズ, 件数]
$threshold  = -1L                                                     # これより大きいファイルだけを候補にする
$totalSize  = 0L
$totalFiles = 0L
$totalDirs  = 0L
$deniedCount   = 0
$errorCount    = 0
$linkCount     = 0
$excludedCount = 0

$dirType      = [IO.DirectoryInfo]
$rpFlag       = [IO.FileAttributes]::ReparsePoint
$hiddenSystem = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System

# 再帰呼び出しは深い階層でスタックが尽きるため、自前のスタックでたどる。
# 要素は (フォルダ, 階層, 順位に載せる祖先フォルダの記録, リパースポイントかどうか)
$stack = New-Object 'System.Collections.Generic.Stack[object]'
$stack.Push(((New-Object IO.DirectoryInfo (ConvertTo-LongPath $target)), 0, @(), $false))

$sw           = [Diagnostics.Stopwatch]::StartNew()
$nextProgress = 0L
$completed    = $false
$fatal        = $null
try {
    while ($stack.Count -gt 0) {
        # PowerShell の変数名は大文字と小文字を区別しないため、パラメーターの $Depth と別の名前にする
        $node      = $stack.Pop()
        $dir       = $node[0]
        $level     = $node[1]
        $chain     = $node[2]
        $isReparse = $node[3]

        if ($sw.ElapsedMilliseconds -ge $nextProgress) {
            Write-Progress -Activity '容量検索' `
                -Status ('フォルダ {0:N0} 件 / ファイル {1:N0} 件 / {2}' -f $totalDirs, $totalFiles, (Format-Size $totalSize)) `
                -CurrentOperation (ConvertTo-DisplayPath $dir.FullName)
            $nextProgress = $sw.ElapsedMilliseconds + $ProgressMs
        }

        try {
            $entries = $dir.GetFileSystemInfos()
        } catch {
            $ex   = Get-InnerException $_.Exception
            $disp = ConvertTo-DisplayPath $dir.FullName
            if ($level -eq 0) {
                $fatal = 'フォルダを読めません: {0}（{1}）' -f $disp, (Get-ErrorText $ex)
                break
            }
            if ($ex -is [UnauthorizedAccessException] -and $isReparse) {
                # 種類を判定できず、読み取りも禁止されたリパースポイントは、リンクとして扱う
                $linkCount++
                $skipped.Add([pscustomobject]@{ '理由' = 'リンク'; 'パス' = $disp; '詳細' = '読み取りが禁止されたリンク' })
            } elseif ($ex -is [UnauthorizedAccessException]) {
                $deniedCount++
                $skipped.Add([pscustomobject]@{ '理由' = '権限なし'; 'パス' = $disp; '詳細' = 'アクセスが拒否されました' })
            } else {
                $errorCount++
                $skipped.Add([pscustomobject]@{ '理由' = 'その他のエラー'; 'パス' = $disp; '詳細' = (Get-ErrorText $ex) })
            }
            continue
        }

        if ($level -gt 0) {
            $totalDirs++
            if ($level -le $Depth) {
                $folderRecord = [pscustomobject]@{ Path = (ConvertTo-DisplayPath $dir.FullName); Depth = $level; Size = 0L; Files = 0L }
                $folders.Add($folderRecord)
                $chain = $chain + $folderRecord
            }
        }

        $dirSize  = 0L
        $dirFiles = 0L
        foreach ($entry in $entries) {
            if ($entry -is $dirType) {
                $entryIsReparse = ($entry.Attributes -band $rpFlag) -eq $rpFlag
                if ($entryIsReparse) {
                    # OneDrive のフォルダもリパースポイントなので、リンクと判定できたものだけを飛ばす
                    $linkType = $null
                    try { $linkType = $entry.LinkType } catch { }
                    $linkName = $null
                    if ($linkType -eq 'Junction') {
                        $linkName = 'ジャンクション'
                    } elseif ($linkType -eq 'SymbolicLink') {
                        $linkName = 'シンボリックリンク'
                    } elseif (($entry.Attributes -band $hiddenSystem) -eq $hiddenSystem) {
                        # Windows が互換用に作ったリンク（All Users、Application Data など）は読み取りが禁止されていて、
                        # 5.1 では種類を判定できない。どれも隠し・システム属性付きなので、それを目印にする
                        $linkName = 'Windows の互換用リンク'
                    }
                    if ($linkName) {
                        $linkCount++
                        $skipped.Add([pscustomobject]@{ '理由' = 'リンク'; 'パス' = (ConvertTo-DisplayPath $entry.FullName); '詳細' = $linkName })
                        continue
                    }
                }
                if ($hasExclude) {
                    $entryPath = ConvertTo-DisplayPath $entry.FullName
                    $matched   = $null
                    foreach ($rule in $excludeRules) {
                        if ($rule.Matcher.IsMatch($entryPath)) { $matched = $rule.Text; break }
                    }
                    if ($matched) {
                        $excludedCount++
                        $skipped.Add([pscustomobject]@{ '理由' = '除外指定'; 'パス' = $entryPath; '詳細' = $matched })
                        continue
                    }
                }
                $stack.Push(($entry, ($level + 1), $chain, $entryIsReparse))
            } else {
                if ($useCutoff -and $entry.LastWriteTimeUtc -ge $cutoffUtc) { continue }
                $length = $entry.Length
                $dirSize += $length
                $dirFiles++

                $ext = $entry.Extension
                $acc = $extTable[$ext]
                if ($null -eq $acc) { $acc = New-Object 'long[]' 2; $extTable[$ext] = $acc }
                $acc[0] += $length
                $acc[1]++

                if ($length -gt $threshold) {
                    $candidates.Add([pscustomobject]@{ Path = $entry.FullName; Size = $length; Time = $entry.LastWriteTime })
                    # 候補が増えすぎたら上位だけ残し、足切りのサイズを上げる
                    if ($candidates.Count -ge 2 * $CsvFileTop) {
                        $kept = @($candidates | Sort-Object -Property Size -Descending | Select-Object -First $CsvFileTop)
                        $candidates.Clear()
                        $candidates.AddRange($kept)
                        $threshold = $kept[-1].Size
                    }
                }
            }
        }

        $totalSize  += $dirSize
        $totalFiles += $dirFiles
        foreach ($record in $chain) {
            $record.Size  += $dirSize
            $record.Files += $dirFiles
        }
    }
    if (-not $fatal) { $completed = $true }
} finally {
    Write-Progress -Activity '容量検索' -Completed
    if (-not $completed) {
        if ($runDir) { try { [IO.Directory]::Delete($runDir) } catch { } }
        if (-not $fatal) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}
if ($fatal) { Stop-WithError $fatal }
$elapsed = $sw.Elapsed

#endregion

# 並べ替え・画面表示・保存の途中で Ctrl+C されたときや、保存に失敗したときは、書きかけの結果を残さない
$finished  = $false
$saveError = $null
try {
    $bySizeThenPath = @{ Expression = 'Size'; Descending = $true }, @{ Expression = 'Path'; Descending = $false }
    $folderRank = @($folders | Sort-Object -Property $bySizeThenPath)
    $fileRank   = @($candidates | Sort-Object -Property $bySizeThenPath | Select-Object -First $CsvFileTop)
    $extList    = @(
        foreach ($key in $extTable.Keys) {
            $extName = if ($key) { $key.ToLowerInvariant() } else { '（拡張子なし）' }
            [pscustomobject]@{ Path = $extName; Size = $extTable[$key][0]; Files = $extTable[$key][1] }
        }
    )
    $extRank = @($extList | Sort-Object -Property $bySizeThenPath)

    #region 画面表示

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add((Format-Invariant '容量検索 (Get-DiskUsage {0})' $ScriptVersion))
    $lines.Add((Format-Invariant '対象: {0}   {1:yyyy-MM-dd HH:mm:ss}   {2}（管理者: {3}）' $target $startTime $env:COMPUTERNAME $(if ($isAdmin) { 'はい' } else { 'いいえ' })))
    if ($useCutoff) { $lines.Add((Format-Invariant '条件: 最終更新が {0} 日以上前のファイルのみ（{1:yyyy-MM-dd HH:mm} より前）' $OlderThan $cutoff)) }
    if ($hasExclude) { $lines.Add(('除外: {0}' -f (($excludeRules | ForEach-Object { $_.Text }) -join ', '))) }
    if ($driveLine) { $lines.Add($driveLine) }
    $lines.Add((Format-Invariant '集計 {0}  ファイル {1:N0} 件  フォルダ {2:N0} 件  所要 {3}' (Format-Size $totalSize) $totalFiles $totalDirs (Format-Duration $elapsed)))

    $notes = @()
    $kinds = @()
    if ($deniedCount -gt 0) { $kinds += , @('権限なし', $deniedCount) }
    if ($errorCount -gt 0) { $kinds += , @('その他のエラー', $errorCount) }
    if ($kinds.Count -eq 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（{1}）' $kinds[0][1] $kinds[0][0]
    } elseif ($kinds.Count -gt 1) {
        $notes += Format-Invariant 'スキップ {0:N0} 件（権限なし {1:N0}・その他のエラー {2:N0}）' ($deniedCount + $errorCount) $deniedCount $errorCount
    }
    if ($linkCount -gt 0) { $notes += Format-Invariant 'たどらなかったリンク {0:N0} 件' $linkCount }
    if ($excludedCount -gt 0) { $notes += Format-Invariant '除外したフォルダ {0:N0} 件' $excludedCount }
    if ($notes.Count -gt 0) { $lines.Add($notes -join '  ') }
    if ($deniedCount -gt 0 -and -not $isAdmin) { $lines.Add('※ 管理者として実行すると、読める範囲が広がります') }

    $lines.Add('')
    $lines.Add(('■ 大きいフォルダ（{0} 階層まで・上位 {1}）' -f $Depth, $Top))
    $lines.Add(' 順位  合計サイズ  ファイル数  フォルダ')
    $rank = 0
    foreach ($row in ($folderRank | Select-Object -First $Top)) {
        $rank++
        $lines.Add((Format-Invariant '{0,5}  {1,10}  {2,10:N0}  {3}' $rank (Format-Size $row.Size) $row.Files $row.Path))
    }
    if ($rank -eq 0) { $lines.Add('    （該当なし）') }

    $lines.Add('')
    $lines.Add(('■ 大きいファイル（上位 {0}）' -f $Top))
    $lines.Add(' 順位      サイズ  最終更新          ファイル')
    $rank = 0
    foreach ($row in ($fileRank | Select-Object -First $Top)) {
        $rank++
        $lines.Add((Format-Invariant '{0,5}  {1,10}  {2:yyyy-MM-dd HH:mm}  {3}' $rank (Format-Size $row.Size) $row.Time (ConvertTo-DisplayPath $row.Path)))
    }
    if ($rank -eq 0) { $lines.Add('    （該当なし）') }

    $lines.Add('')
    $lines.Add(('■ 拡張子ごとの合計（上位 {0}）' -f $Top))
    $lines.Add(' 順位  合計サイズ  ファイル数  拡張子')
    $rank = 0
    foreach ($row in ($extRank | Select-Object -First $Top)) {
        $rank++
        $lines.Add((Format-Invariant '{0,5}  {1,10}  {2,10:N0}  {3}' $rank (Format-Size $row.Size) $row.Files $row.Path))
    }
    if ($rank -eq 0) { $lines.Add('    （該当なし）') }

    foreach ($line in $lines) {
        if ($line.StartsWith('■')) { Write-Host $line -ForegroundColor Cyan } else { Write-Host $line }
    }

    #endregion

    #region 保存

    if ($runDir) {
        Write-Progress -Activity '容量検索' -Status '結果を保存しています'
        try {
            [IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), [string[]]$lines, $Utf8Bom)

            # Excel の「標準」書式は 12 桁以上の数値を指数で表示するため、数値は 3 桁区切りで出す
            # （Excel は桁区切り付きの数値として読み込む）。行数が多いため、関数を呼ばずに書式を直接指定する
            $rank = 0
            $rows = @(
                foreach ($row in $folderRank) {
                    $rank++
                    [pscustomobject]@{ '順位' = $rank; 'フォルダ' = $row.Path; '階層' = $row.Depth; 'サイズ(MB)' = ($row.Size / 1MB).ToString('N2', $Invariant); 'サイズ(バイト)' = $row.Size.ToString('N0', $Invariant); 'ファイル数' = $row.Files.ToString('N0', $Invariant) }
                }
            )
            Save-Csv $rows @('順位', 'フォルダ', '階層', 'サイズ(MB)', 'サイズ(バイト)', 'ファイル数') (Join-Path $runDir 'folders.csv')

            $rank = 0
            $rows = @(
                foreach ($row in $fileRank) {
                    $rank++
                    $filePath = ConvertTo-DisplayPath $row.Path
                    [pscustomobject]@{ '順位' = $rank; 'ファイル' = $filePath; 'サイズ(MB)' = ($row.Size / 1MB).ToString('N2', $Invariant); 'サイズ(バイト)' = $row.Size.ToString('N0', $Invariant); '最終更新日時' = $row.Time.ToString('yyyy/MM/dd HH:mm:ss', $Invariant); '拡張子' = (ConvertTo-ExcelText ([IO.Path]::GetExtension($filePath).ToLowerInvariant())) }
                }
            )
            Save-Csv $rows @('順位', 'ファイル', 'サイズ(MB)', 'サイズ(バイト)', '最終更新日時', '拡張子') (Join-Path $runDir 'files.csv')

            $rank = 0
            $rows = @(
                foreach ($row in $extRank) {
                    $rank++
                    [pscustomobject]@{ '順位' = $rank; '拡張子' = (ConvertTo-ExcelText $row.Path); 'サイズ(MB)' = ($row.Size / 1MB).ToString('N2', $Invariant); 'サイズ(バイト)' = $row.Size.ToString('N0', $Invariant); 'ファイル数' = $row.Files.ToString('N0', $Invariant) }
                }
            )
            Save-Csv $rows @('順位', '拡張子', 'サイズ(MB)', 'サイズ(バイト)', 'ファイル数') (Join-Path $runDir 'extensions.csv')

            Save-Csv $skipped @('理由', 'パス', '詳細') (Join-Path $runDir 'skipped.csv')
        } catch {
            $saveError = Get-ErrorText (Get-InnerException $_.Exception)
        }
        Write-Progress -Activity '容量検索' -Completed
    }

    #endregion

    $finished = $true
} finally {
    if (-not $finished -or $saveError) {
        # 消すのは、このツールが作った保存先のフォルダだけ
        if ($runDir) { try { [IO.Directory]::Delete($runDir, $true) } catch { } }
        if (-not $finished) { Write-Host '中断しました。結果は保存していません。' -ForegroundColor Yellow }
    }
}

if ($saveError) {
    Write-Host ''
    Write-Warning ('結果を保存できませんでした: {0}' -f $saveError)
} elseif ($runDir) {
    Write-Host ''
    Write-Host ('保存先: {0}' -f $runDir)
} elseif ($saveWarning) {
    Write-Host ''
    Write-Warning $saveWarning
}

exit 0
