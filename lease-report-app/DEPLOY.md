# GitHub と Cloudflare Pages へのデプロイ

対象リポジトリ例: [dezigozi/kasouhin_uriagezumi_Web](https://github.com/dezigozi/kasouhin_uriagezumi_Web)

## 前提（重要）

- **Cloudflare Pages は静的サイト用**です。Excel を読む **Node API（`server.js`）は Pages 上では動きません**（ファイル共有・Excel 復号が必要なため）。
- 構成は次の二段です。
  1. **Pages**: フロント（`build/`）
  2. **別ホスト**: API（Docker や社内サーバー、Cloudflare Tunnel など）

## 1. GitHub にプッシュ

このフォルダ（`lease-report-app`）をリポジトリのルートにする想定です。

```bash
cd lease-report-app
git init
git add .
git commit -m "Initial commit: lease report web + deploy config"
git branch -M main
git remote add origin https://github.com/dezigozi/kasouhin_uriagezumi_Web.git
git push -u origin main
```

すでに上位フォルダで Git 管理している場合は、サブフォルダだけを別リポジトリにしたいときは [git subtree](https://git-scm.com/book/en/v2/Git-Tools-Subtree-Merging) や新規クローンへファイルをコピーする方法を使ってください。

## 2. API の設定（これをしないとクラウドの画面はデータを読めません）

Cloudflare に載るのは **ブラウザ用の画面（HTML/JS）だけ**です。Excel は **ユーザーの PC ではなく「API を動かしているサーバー」**が読みます。次の **①〜④** をセットします。

| 順番 | どこで | 何をする |
|------|--------|----------|
| ① | **API を動かすマシン** | `server.js`（または Docker）で API を起動し、共有フォルダにそのマシンからアクセスできるようにする |
| ② | **API サーバー** | 環境変数 `LEASE_REPORT_CORS_ORIGIN` に **フロントの URL**（Workers/Pages の `https://….workers.dev` など）を指定 |
| ③ | **Cloudflare** | **Variables and secrets** に `LEASE_REPORT_API_BASE` を追加し、**② の API の公開 URL**（`https://api.会社.com` など）を書く |
| ④ | **Cloudflare** | 変数を追加したあと **再ビルド・再デプロイ**（ビルド時に URL が JS に埋め込まれるため） |

### 2.1 API をどこで動かすか

- **社内の Windows サーバー or 常駐 PC** が現実的です（`\\192.…\share\…` の UNC に、そのマシンからアクセスできる必要があります）。
- API を **Docker** で動かす場合も、**コンテナの中からその共有が見える**ようにマウントするか、ホストでパスが通る場所に置いてください。

### 2.2 起動方法 A: Node で直接（開発・本番どちらでも可）

リポジトリルートで:

```bash
npm ci
node server.js
```

既定では **ポート 3001**（`SERVE_STATIC` 未設定時）。別ポートにする例:

```powershell
$env:PORT="3001"
node server.js
```

### 2.3 起動方法 B: Docker

```bash
docker build -t lease-report-api .
docker run -d -p 3001:3001 ^
  -e LEASE_REPORT_CORS_ORIGIN=https://kasouhinuriagezumiweb.<あなたのサブドメイン>.workers.dev ^
  lease-report-api
```

（Linux のシェルでは `^` を `\` 改行に読み替え。）

### 2.4 動作確認

ブラウザまたは curl で（API を起動したマシン上からでも可）:

- `http://localhost:3001/api/health` → `{"ok":true,"service":"lease-report-api"}` なら API は起動しています。

### 2.5 CORS（必須に近い）

ブラウザは **フロントのドメイン** と **API のドメイン** が違うと、API 側の許可が無いと通信できません。

- API の環境変数 **`LEASE_REPORT_CORS_ORIGIN`** に、**実際にユーザーが開くフロントの URL のオリジン**を書きます（**末尾スラッシュなし**）。
- **複数**（本番・プレビュー）なら **カンマ区切り**。

例:

```text
LEASE_REPORT_CORS_ORIGIN=https://kasouhinuriagezumiweb.xxxxx.workers.dev,https://main.kasouhin-uriagezumi-web.pages.dev
```

`LEASE_REPORT_CORS_ORIGIN` を **未設定**にすると、API は `Access-Control-Allow-Origin: *` で応答します（手早いが、本番ではオリジン指定を推奨）。

### 2.6 Cloudflare に `LEASE_REPORT_API_BASE` を設定して再ビルド

1. Cloudflare ダッシュボード → 該当プロジェクト → **Settings** → **Variables and secrets**。
2. **Production**（および必要なら Preview）に **Variable** を追加:
   - **名前:** `LEASE_REPORT_API_BASE`
   - **値:** API の **https のオリジンだけ**（例: `https://api.会社.com`）。**`/api` は付けない。** 末尾スラッシュも不要。
3. **Save** したあと、**新しいデプロイを走らせる**（「Retry deployment」や空コミットで push など）。  
   この値は **`npm run build` 実行時**に `webpack` が `bundle.js` に埋め込みます。設定しただけでは既存の JS は変わりません。

### 2.7（任意）API トークンで保護する

API 側で `LEASE_REPORT_API_TOKEN` を設定すると、リクエストに  
`Authorization: Bearer <同じ文字列>` が必要になります。

フロントは **`public/index.html`** の `<head>` 内などで、ビルド・デプロイ前に次を入れる方法があります（リポジトリにトークンをコミットしないでください。CI のシークレットで HTML を生成する運用が安全です）。

```html
<script>window.__LEASE_REPORT_API_TOKEN__ = 'ここにトークン';</script>
```

### 2.8 アプリ内「設定」の Excel パスについて

画面の **システム設定** に入れるパスは、**ブラウザではなく API サーバーから見たパス**です。社内共有なら **そのサーバーで使っている UNC**（例: `\\192.1.1.103\share\…`）をそのまま指定します。

### 2.9 その他の API 用環境変数（参照）

| 変数 | 説明 |
|------|------|
| `PORT` | 待受ポート（API のみ時は既定 3001） |
| `LEASE_REPORT_CACHE_DIR` | キャッシュ保存先ディレクトリ |
| `EXCEL_PASSWORD` | 暗号化 Excel 用パスワード（未設定時はアプリ既定） |
| `LEASE_REPORT_API_TOKEN` | 設定時は Bearer 認証必須 |
| `LEASE_REPORT_CORS_ORIGIN` | 許可するフロントのオリジン（カンマ区切り可） |

## 3. Cloudflare の設定（Pages と Workers のどちらか）

このリポジトリは **静的フロント（`build/`）** なので、Cloudflare では次のどちらかです。

### A. Cloudflare **Pages**（推奨・シンプル）

1. [Cloudflare Dashboard](https://dash.cloudflare.com/) → **Workers & Pages** → **Create** → **Pages**（※「Workers」ではなく **Pages**）→ **Connect to Git**。
2. ビルド設定:

| 項目 | 値 |
|------|-----|
| Build command | `npm ci && npm run build` |
| Build output directory | `build` |
| **Deploy command** | **空欄** |
| Root directory | `/` |

### B. いまのように **Workers の Git 連携**（Build 画面に Deploy command があるタイプ）

ダッシュボードの **Runtime → Build** で次を確認します。

| 項目 | 値 |
|------|-----|
| Build command | `npm run build`（または `npm ci && npm run build`） |
| Deploy command | `npx wrangler deploy`（そのままで可） |
| Root directory | `/` |

リポジトリの **`wrangler.toml`** で **`name`** を Workers プロジェクトの **Name**（例: `kasouhinuriagezumiweb`）と **完全一致**させ、`[assets]` で **`./build`** を指定します（このリポジトリに同梱済み）。SPA のため `not_found_handling = "single-page-application"` を付けています。詳細は [Workers Static Assets / SPA](https://developers.cloudflare.com/workers/static-assets/routing/single-page-application) を参照。

**Pages 用の `wrangler pages deploy`** を CLI で使う例: `npx wrangler pages deploy build --project-name=<Pages のプロジェクト名>`（[Direct Upload](https://developers.cloudflare.com/pages/get-started/direct-upload/)）。

3. **Variables and secrets** に **`LEASE_REPORT_API_BASE`** を入れ、**再ビルド**する手順は **§2.6** と同じです（フロントが API を呼ぶために必須）。

4. API 側の **CORS**（`LEASE_REPORT_CORS_ORIGIN`）は **§2.5** を参照してください。

## 4. wrangler.toml

- **Workers + `npx wrangler deploy`**: リポジトリの `wrangler.toml` がそのまま使われます（`name` と `[assets].directory = "./build"`）。
- **Pages**: **Deploy command は空**。`wrangler.toml` は必須ではありません（参考用）。

## 5. 社内ネットワークだけで見せる場合

API を社内 PC で動かし、[Cloudflare Tunnel](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/) で `https://report.example.com` のように公開する方法もあります。その場合もフロントは Pages、API は Tunnel の先のオリジン、という組み合わせが可能です。
