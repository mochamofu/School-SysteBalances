# ランディングページ（独立サイト版）

`site/index.html` は、積立金会計 入力アシスタントの紹介用ランディングページです。
**1ファイル完結の静的サイト**なので、どのホスティングサービスにもそのまま置けます。

## 公開にお金はかかる？

**基本は0円です。** 静的サイト（HTMLだけのサイト）の公開は、主要サービスすべてに無料枠があります。

| 方法 | 費用 | 手間 | URL |
|---|---|---|---|
| **GitHub Pages**（設定済み・おすすめ） | 無料 | 設定1回 | `https://mochamofu.github.io/School-SysteBalances/` |
| **Netlify Drop** | 無料 | ドラッグ&ドロップだけ | `https://好きな名前.netlify.app` |
| **Cloudflare Pages** | 無料 | GitHub連携で自動 | `https://好きな名前.pages.dev` |

お金がかかるのは、**独自ドメイン**（例: `tsumitate-assistant.jp`）を使いたい場合だけです。
これはどのサービスを使っても共通で、ドメイン代として**年1,000〜2,000円程度**（お名前.com、Cloudflare Registrar等）。
上記の `~.github.io` / `~.netlify.app` などのURLのままなら完全無料です。

## 公開手順

### A. GitHub Pages（このリポジトリだけで完結・おすすめ）

デプロイ用ワークフロー（`.github/workflows/deploy-pages.yml`）は設定済みです。あとは1回だけ:

1. GitHubリポジトリの **Settings → Pages** を開く
2. **Source** を **「GitHub Actions」** に変更する

これだけで `https://mochamofu.github.io/School-SysteBalances/` に公開され、
以後 `site/` を更新してmainにpushするたびに自動で反映されます。

### B. Netlify Drop（アカウント登録すら最小限・一番手軽）

1. https://app.netlify.com/drop を開く
2. `site` フォルダをブラウザにドラッグ&ドロップ
3. その場でURLが発行される（サイト名は後から変更可）

### C. Cloudflare Pages

1. https://pages.cloudflare.com でGitHubリポジトリを連携
2. Build command: なし ／ Output directory: `site`
3. 自動でビルド・公開される

## 内容を直したいとき

`site/index.html` を直接編集してください（HTML+CSS+少量のJSが1ファイルに入っています）。
文言・色・セクションの追加削除はすべてこのファイル内で完結します。
