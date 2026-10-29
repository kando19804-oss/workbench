# Claude Code のフック（PostToolUse: Write|Edit）
# .ps1 / .psm1 / .psd1 を保存したあとに、次の 2 つを行う。
#   1. BOM がなければ付ける（Windows PowerShell 5.1 で日本語が化けないように）
#   2. Windows PowerShell 5.1 で文法エラーがないか確かめる。エラーがあれば終了コード 2 で Claude に知らせる
# BOM がないと 5.1 が日本語を読み違えて誤った文法エラーを出すため、必ず 1 → 2 の順に行う。

try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

$reader  = New-Object IO.StreamReader([Console]::OpenStandardInput(), [Text.Encoding]::UTF8)
$payload = $reader.ReadToEnd() | ConvertFrom-Json
$file    = $payload.tool_input.file_path
if (-not $file -or $file -notmatch '\.ps[dm]?1$' -or -not (Test-Path -LiteralPath $file)) { exit 0 }

# 1. BOM を付ける（UTF-16 のファイルはそのまま）
$bytes   = [IO.File]::ReadAllBytes($file)
$hasBom  = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
$isUtf16 = $bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF))
if (-not $hasBom -and -not $isUtf16) {
    $withBom = New-Object byte[] ($bytes.Length + 3)
    $withBom[0] = 0xEF; $withBom[1] = 0xBB; $withBom[2] = 0xBF
    [Array]::Copy($bytes, 0, $withBom, 3, $bytes.Length)
    [IO.File]::WriteAllBytes($file, $withBom)
}

# 2. 5.1 の構文チェック。5.1 のエラーメッセージは日本語なので、文字化けしないようにファイル経由で受け取る
$resultFile = [IO.Path]::GetTempFileName()
try {
    $target = $file.Replace("'", "''")
    $output = $resultFile.Replace("'", "''")
    $check  = @"
`$errors = `$null
[void][Management.Automation.Language.Parser]::ParseFile('$target', [ref]`$null, [ref]`$errors)
`$lines = @(`$errors | ForEach-Object { '{0} 行目: {1}' -f `$_.Extent.StartLineNumber, `$_.Message })
[IO.File]::WriteAllLines('$output', [string[]]`$lines, (New-Object Text.UTF8Encoding `$false))
"@
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($check))
    $null = & powershell.exe -NoProfile -NonInteractive -EncodedCommand $encoded 2>&1
    $errors = @([IO.File]::ReadAllLines($resultFile, [Text.Encoding]::UTF8) | Where-Object { $_ })
} finally {
    [IO.File]::Delete($resultFile)
}

if ($errors.Count -gt 0) {
    [Console]::Error.WriteLine(('Windows PowerShell 5.1 で文法エラーがあります: {0}' -f $file))
    foreach ($line in $errors) { [Console]::Error.WriteLine('  ' + $line) }
    [Console]::Error.WriteLine('5.1 にない文法（三項演算子、??、?.、&& など）を使っていないか確認してください。')
    exit 2
}
exit 0
