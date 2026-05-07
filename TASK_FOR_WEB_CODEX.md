# Web版Codexへの依頼

このリポジトリを、Windowsローカル環境で履歴書docxから資格取得ヒストリーPowerPointを生成できるツールに整備してください。

## 背景

元の履歴書:

```text
E:\My Resources\My Historys\履歴書_太地稔_20260103.docx
```

ローカルCodex環境ではPython、Node.js、PowerPoint COMが使えなかったため、現行版はPowerShellでOpenXMLを直接生成しています。

現行成果物:

```text
out/資格取得ヒストリー_太地稔_20260103.pptx
```

## 期待する成果

- `data/certifications.json` を入力としてPowerPointを生成する
- `python-pptx`版の `src/resume_cert_ppt/generate_ppt.py` を完成させる
- `out/assets/certification-history-bg.png` を背景画像として埋め込む
- 現行PowerShell版と同等以上の見栄えにする
- Windowsで動く `setup.ps1` と `README.md` を整備する
- 可能なら `python-docx` を追加し、履歴書docxから資格欄を抽出できるようにする
- PowerPoint COMが使える環境ではPDFまたはPNGプレビューを書き出せるようにする
- JSON整合性、Python構文、PowerPoint生成の検証手順を追加する

## 制約

- Web版Codexのクラウド環境から、ユーザーのWindows PCには直接インストールしない
- 履歴書docxの原本はローカルPC上にあるため、必要ならサンプルデータとして `data/certifications.json` を使う
- 生成画像はすでに `out/assets/certification-history-bg.png` に保存済み

## 推奨コマンド

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
.\.venv\Scripts\python.exe -m resume_cert_ppt.generate_ppt
```

## 受け入れ条件

- `out/資格取得ヒストリー_太地稔_20260103_python.pptx` が生成される
- 5枚構成のPowerPointになる
- 生成画像が埋め込まれる
- 資格データ14件が欠落しない
- READMEだけでWindows上の再生成手順が分かる
