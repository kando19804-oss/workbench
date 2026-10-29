---
name: tool-dev
description: workbench のツールの開発フェーズ（開発の流れ 2/4）。ユーザーが承認した tools/<名前>/README.md に沿って、Windows PowerShell 5.1 で動く 1 ファイル完結の .ps1 を実装し、5.1 と 7 で一度動かす。終わったら /tool-review に進む。
argument-hint: "<ツール名>"
---

# 開発（開発の流れ 2/4）

対象: $ARGUMENTS

## 前提

- `tools/<名前>/README.md` が、要件フェーズでユーザーに承認されていること。承認されていなければ、`/tool-spec` から始めるよう伝えて止まる

## 手順

1. `CLAUDE.md` の開発ルールと「5.1 の落とし穴」を読み直す
2. `tools/<名前>/<名前>.ps1` を書く。構成と書き方は `tools/Get-DiskUsage/Get-DiskUsage.ps1` に合わせる
   - 先頭に日本語のヘルプと `.NOTES` のライセンス表示
   - `#Requires -Version 5.1`、`Set-StrictMode -Version 2.0`
   - パラメーターと同じ名前のローカル変数を使わない
   - 共通処理（長いパスの扱い、リンクの判定、CSV の保存など）は Get-DiskUsage からコピーして使う
3. 保存するとフックが BOM を付け、5.1 の文法を確認する。エラーが返ってきたら、その場で直す
4. 小さなダミーデータで、`powershell.exe`（5.1）と `pwsh`（7）の両方から一度動かし、README の表示例や出力ファイルの形どおりかを確かめる
5. 実装中に README と違う判断をしたら、README も直して、その理由を記録しておく

## 終わり方

- 作ったファイルと、README から変えた点を短く報告する
- ユーザーの確認は待たずに、`/tool-review` で実装のレビューに進む
