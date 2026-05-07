# GitHubへ渡す手順

このPCにはGit for Windowsを導入済みで、ローカルリポジトリは `main` ブランチの初回コミットまで完了しています。

GitHubコネクタから見えるリポジトリは0件のため、現時点ではGitHub側で新規リポジトリを作成してURLを取得する必要があります。

## 手順

1. GitHubで新しいリポジトリを作成します。
   - 例: `resume-cert-ppt-generator`
   - Public / Private はどちらでも構いません。

2. 次のZIPを展開します。

   ```text
   E:\AI\Codex\動画編集\out\resume-cert-ppt-web-codex-package.zip
   ```

3. 展開した中身をGitHubリポジトリへアップロードします。

4. ChatGPTのWeb版Codexでそのリポジトリを選びます。

5. [TASK_FOR_WEB_CODEX.md](TASK_FOR_WEB_CODEX.md) の本文をWeb版Codexへ貼り付けて依頼します。

## GitHubコネクタについて

この環境からGitHubコネクタを確認したところ、ユーザー `taiji-3nr` として認証されていますが、Codexからアクセス可能なリポジトリは0件でした。

既存リポジトリをCodexに扱わせるには、GitHub側でCodex/GitHub Appに対象リポジトリへのアクセスを許可してください。

## ローカルでpushする場合

GitHubで空のリポジトリを作成したら、次を実行します。

```powershell
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```
