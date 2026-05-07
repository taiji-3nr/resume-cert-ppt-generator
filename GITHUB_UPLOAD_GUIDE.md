# GitHubへ渡す手順

このPCでは `git` と `gh` が見つからないため、現時点では手動アップロードが最短です。

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

## ローカルでgitを使う場合

このPCにGit for Windowsを入れると、次の流れにできます。

```powershell
git init
git add .
git commit -m "Add resume certification PowerPoint generator"
git branch -M main
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```

