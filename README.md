# 資格取得ヒストリー PowerPoint Generator

履歴書の「免許・資格」欄をもとに、資格取得ヒストリーのPowerPointを生成するための作業フォルダです。

## 現在の成果物

- `out/資格取得ヒストリー_太地稔_20260103.pptx`
  - PowerShell/OpenXMLで作成した現行版のPowerPoint
- `out/assets/certification-history-bg.png`
  - ChatGPT画像生成で作成した背景画像
- `data/certifications.json`
  - 履歴書から抽出した資格取得データ
- `build_cert_history_ppt.ps1`
  - Pythonなしでpptxを作るためのPowerShell版スクリプト
- `src/resume_cert_ppt/generate_ppt.py`
  - `python-pptx`へ移行するためのPython版スクリプト

## Windowsローカルでの推奨セットアップ

Python 3.11以降をインストールし、PowerShellで次を実行します。

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

PowerPointを生成します。

```powershell
.\.venv\Scripts\python.exe -m resume_cert_ppt.generate_ppt
```

出力先:

```text
out/資格取得ヒストリー_太地稔_20260103_python.pptx
```

## Web版Codexへ依頼する内容

Web版Codexでは、ローカルPCへのPython/Node/Officeインストールはできません。代わりに、このリポジトリのコード整備と検証を依頼するのが適しています。

依頼文の例:

```text
このリポジトリを、Windowsローカル環境で履歴書docxから資格取得ヒストリーPowerPointを生成できるツールに整備してください。

要件:
- data/certifications.json を入力として、python-pptxでPowerPointを生成する
- 既存の out/資格取得ヒストリー_太地稔_20260103.pptx と同等以上の見栄えにする
- setup.ps1、requirements.txt、README.mdを最新化する
- PowerPoint COMが使える環境ではPDFまたはPNGプレビューも出力できるようにする
- 生成物の検証手順をREADMEに明記する
```

## 今後の改善候補

- `python-docx`で履歴書docxから資格欄を自動抽出する
- PowerPoint COMが使える場合にPDF/PNGプレビューを書き出す
- 画像生成プロンプトをJSON管理し、背景画像の差し替えを簡単にする
- GitHub Actions上でJSON整合性とPython構文チェックを行う

## Web版Codex用タスク

Web版Codexには [TASK_FOR_WEB_CODEX.md](TASK_FOR_WEB_CODEX.md) の内容を貼り付けて依頼してください。

GitHubへ渡す具体的な手順は [GITHUB_UPLOAD_GUIDE.md](GITHUB_UPLOAD_GUIDE.md) を参照してください。

## ローカル検証

Pythonが未導入でも、現行成果物とデータの整合性は次で確認できます。

```powershell
powershell -ExecutionPolicy Bypass -File .\validate.ps1
```
