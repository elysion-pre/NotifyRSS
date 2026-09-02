# NotifyRSS
- RSSの更新をDiscordのチャンネルへ通知
- **Google Apps Script(以降GAS)** のソースコードを
**GitHub** で編集してコミットしたら自動でGASのソースコードへ反映させる環境の構築

---

## 🗒️ はじめに
ブログやニュースサイトの更新をDiscordに通知したい時、  
既存のIFTTTやZapierなどのツールだと
- 「画像が表示されない」
- 「更新頻度に遅延がある」
- 「無料枠の制限が厳しい」

といった悩みに直面しがちです。

そこで今回は **GAS** を使用し、無料で自作できる **Discord通知ボット** の作成方法とソースコードを公開します。

スマホのブラウザだけでも構築可能な **GitHub連携による自動デプロイ環境の作り方** もあわせて解説します。

---

## 💡 このスクリプトの特徴

- **完全無料＆制限なし**
  - GASの標準機能のみで動くため費用がかかりません。
- **画像抽出の精度が抜群**
  - アイキャッチ画像（OGP）を優先的に取得し、直リンク拒否対策としてDiscordに直接画像ファイルを添付送信します。
- **主要なRSS形式を自動判定**
  - RSS 2.0、Atom、RDF形式を判別して処理します。
- **除外キーワード・拡張子フィルター**
  - 記事内の絵文字、コメント欄のアイコン、広告画像などの不要な画像を除外します。
- **GitHub Actions連携**
  - GitHub上でコードを編集・コミットするだけでGASへ自動送信（デプロイ）され、結果がDiscordに届きます。

---

## 📜 ソースコードと全体の構成
最新のソースコードおよびファイル構成は以下の **GitHubリポジトリ** で管理・公開しています。

https://github.com/elysion-pre/NotifyRSS

```text
.
├── .github/
│   └── workflows/
│       └── deploy.yml   # GitHub Actions (自動デプロイ＆Discord通知)
├── .clasp.json          # clasp設定ファイル
├── appsscript.json      # GASプロジェクト設定
├── Config.gs            # 設定値および除外キーワードの定義ファイル
└── NotifyRSS.js              # RSS取得・パース・Discord送信を行うメイン処理
```

---

## 💬 1. Discord Webhookの取得手順

### ステップ 1：Discordで「**Webhook URL**」を取得する
1. **Discord アプリ** を開きます
2. 通知を飛ばしたいサーバーのチャンネル名を長押しし、「**チャンネルの編集（歯車マーク）**」 をタップします
3. 「**連携サービス（またはウェブフック）**」 をタップし、「**新しいウェブフックを作成**」 をタップします
4. 新しく作成された Webhook をタップし、「**ウェブフックURLをコピー**」 をコピーしてメモ帳に控えておきます。

---

## 📊 2. スプレッドシートの設定手順

### ステップ 1: スプレッドシートの準備

新しい **Googleスプレッドシート** を作成し、  
シート名を **`シート1`**（デフォルト）にします。  
2行目以降に通知したいRSSの情報を入力します。

**シート1**
| 列 | 項目名 | 説明 | 例 |
| :--- | :--- | :--- | :--- |
| **A列** | Site Name | サイト名(D列から自動取得) | `=IMPORTFEED(D2, "feed title", false, 1)` |
| **B列** | Type | E列のdiscord webhook URLを自動入力させるもの | `news` or `game` etc... |
| **C列** | Site URL | サイトのURL | `https://hogehoge.com/` |
| **D列** | RSS URL | 通知したいサイトのRSSフィードURL | `https://hogehoge.com/feed` |
| **E列** | Webhook URL | 通知先のDiscord Webhook URL | `=VLOOKUP(B2, 'シート2'!$A$2:$B$6, 2, false)` |
| **F列** | アイコンURL | （任意）フッターおよびBotアバター用のカスタム画像URL | `https://hogehoge.com/icon.png` |
| **G列** | カラーコード | （任意）Discord Embedの左端のバーの色 | #3498db |
| **H列** | 最終更新日時 | スクリプトが自動で最終通知日時を記録します（初期値は空欄でOK） | *自動割り当て* |

> 💡 **初回実行時について**  
> H列が空欄の状態で実行された場合、過去記事の連投を防ぐため **「現在の最新記事の日時」** を記録して次回以降の更新チェックの基準とします。
> **初回検知時はDiscordへ通知されません**

**シート2**
| 列 | 項目名 | 説明 | 例 |
| :--- | :--- | :--- | :--- |
| **A列** | Type | 記事の種類 | `news` or `game` etc... |
| **B列** | Webhook URL | 通知先のDiscord Webhook URL | `https://discord.com/api/webhooks/...` |

> 💡 シート2を使用せず、シート1に手動で入力しても完結可能です

---

## 🚀 3. GASの設定手順

### ステップ 1: スクリプトプロパティの設定

1. **GASエディタ** の左メニューから **「プロジェクトの設定（歯車アイコン）」** を開きます。
2. **「スクリプト プロパティ」** の項目で以下を追加します：
   - **プロパティ**: SPREADSHEET_ID
   - **値**: 準備したスプレッドシートのID（URLの `https://docs.google.com/spreadsheets/d/【ここの文字列】/edit` の部分）

---

### ステップ 2: 定期実行（トリガー）の設定

1. GASエディタの左メニューから **「トリガー（時計アイコン）」** を開きます。
2. 右下の **「トリガーを追加」** をクリックし、以下のように設定します：
   - **実行する関数を選択**
     - checkRssAndNotifyDiscord
   - **実行するデプロイを選択**
     - Head
   - **イベントのソースを選択**
     - 時間駆動型
   - **時間ベースのトリガーのタイプを選択**
     - 分ベースのタイマー
   - **時間の間隔を選択**
     - 10分おき や 15分おき（お好みに合わせて調整）

---

## 🛠 4. GitHubからGASへの自動デプロイ環境構築

**GitHub** 上でコードを編集・コミットした際に、自動で  
**clasp(Command Line Apps Script Projects)** を経由して **デプロイ（GASへ反映）** し、  
結果を Discordに通知する **CI/CD** 環境の構築手順です。

> 💡 PCはもちろん、スマホのブラウザのみでも構築可能です。

### ステップ 1： GAS側で API を「オン」にする
1. ブラウザで [Google Apps Script ユーザー設定](https://script.google.com/home/usersettings)にアクセスします。
2. **「Google Apps Script API」** という項目のスイッチを **「オン」** に変更します。

---

### ステップ 2： デプロイ先 GAS の「スクリプトID」をコピーする
1. 目的の GAS プロジェクトをブラウザで開きます。
2. 画面左側のメニューにある **歯車マーク（プロジェクトの設定）** をタップします。
3. 画面を下にスクロールし、**「スクリプト ID」** の文字列をコピーしてメモ帳等に控えます。

---

### ステップ 3： Google Cloud Shell で CLASPRC_JSON を取得する
1. ブラウザで [Google Cloud コンソール](https://console.cloud.google.com/)にアクセスします。
2. 画面右上にある **`>_`（アクティベート Cloud Shell）アイコン** をタップし、ターミナルが立ち上がるまで数秒待ちます。
3. 以下のコマンドを順に入力・実行します。

**claspをシステムにインストール**
```bash
npm install -g @google/clasp
```
**ログインを開始**
```bash
clasp login --no-localhost
```

4. 画面に `Authorize clasp by visiting this url:` という文字と一緒に長い URL が出力されるので、コピーして新しいタブで開きます。
5. Googleのログインと権限の許可画面が出るので、すべて許可します。
6. 完了するとアドレスバーに `http://localhost:8505/?code=...` という URL が表示されるので、アドレスバーの **localhost の URL 全体** をコピーします。（※`サイトにアクセスできません`と出ても問題ありません）
7. Cloud Shell のタブに戻り、画面右上にある **「＋」（新しいタブを開く）ボタン** から **2つ目のターミナル** を開きます。
8. 新しい画面で **curlコマンド** を実行します。（※URLは **ダブルクォーテーション** で囲んでください）

```bash
curl -L "ここにコピーしたlocalhostのURLを貼り付け"
```

9. 実行後、画面に `You are logged in!` と出たら1つ目のターミナルタブに戻り、以下のコマンドで認証情報を出力します。

```bash
cat ~/.clasprc.json
```

10. 画面に出力された `{ "token": ... }` から始まるテキストを全選択してコピーします。

> ⚠️ **ターミナル上で選択が難しい場合**  
> `cat ~/.clasprc.json > token.txt` を実行し、画面右上の鉛筆マーク（エディタ）から `token.txt` を開いてコピーしてください。

---

### ステップ 4： GitHub に「シークレット（秘密鍵）」を登録する
1. ブラウザでGitHubリポジトリの **「Settings」 ＞ 「Secrets and variables」 ＞ 「Actions」** を開きます。
2. **「New repository secret」** ボタンをタップし、以下の3つを登録します。

| Secret 名 | 設定する値 |
| :--- | :--- |
| **SCRIPT_ID** | ステップ2で控えた「スクリプトID」 |
| **CLASPRC_JSON** | ステップ3で取得した `{ "token": ... }` の全文 |
| **DISCORD_WEBHOOK** | デプロイ結果を通知したい Discord の Webhook URL |

---

### ステップ 5： GitHub 上で構成ファイルを作成する
リポジトリのトップ画面（main ブランチ）で **「Add file」 ＞ 「Create new file」** を選び、以下のファイルをそれぞれ作成・コミットします。

**.clasp.json**
```json
{
  "scriptId": "${SCRIPT_ID}"
}
```
> 💡 ファイル名の先頭に **ドット** が必要なので注意

**appsscript.json**
```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```
> 💡 「app**ss**cript」と、sが2つ重なるので注意

---

### ステップ 6： GitHub Actions ワークフローを作成する

リポジトリトップで **「Add file」 ＞ 「Create new file」** を選択し、ファイル名に `.github/workflows/deploy.yml` と入力して以下のコードを貼り付け、コミットします。

```yml
name: Deploy to GAS

on:
  push:
    branches:
      - main

jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '20'

      - name: Install clasp
        run: npm install -g @google/clasp

      - name: Create clasprc.json
        run: |
          echo '${{ secrets.CLASPRC_JSON }}' > ~/.clasprc.json

      - name: Replace Script ID
        run: |
          sed -i "s/\${SCRIPT_ID}/${{ secrets.SCRIPT_ID }}/g" .clasp.json

      - name: Push to GAS
        run: clasp push --force

      - name: Notify Discord
        if: always()
        uses: sarisia/actions-status-discord@v1
        with:
          webhook: ${{ secrets.DISCORD_WEBHOOK }}
          status: ${{ job.status }}
          title: "GAS自動デプロイ結果"
          nofail: false
          username: "GitHub Actions"
          description: |
            **リポジトリ:** ${{ github.repository }}
            **実行者:** ${{ github.actor }}
            **コミットメッセージ:**
            ```
            ${{ github.event.head_commit.message }}
            ```
```

コミット完了後、自動的に **GitHub Actions** が起動し、デプロイ結果が **Discord チャンネル**に届きます。

> ⚠️ 注意
> コード修正は **必ず** GitHub側で行うこと
> GASのエディタ画面を開いて直接書き換えても、次にGitHub Actionsが動いたときに上書きされて消えてしまいます。

---

## ⚙️ 5. スクリプトの設定

**Config.gs** の設定値を書き換えることで、  
除外キーワードや動作パラメータの調整が可能です。

```js
/**
 * 処理の設定および除外キーワードの定義ファイル
 */
const CONFIG = {
  // 1回の実行で送信する最大記事数（Discordレートリミット対策）
  MAX_POSTS_PER_RUN: 5,

  // デフォルトの埋め込みカラー（16進数数値）
  DEFAULT_COLOR: 0x3498db,

  // 除外する動画・非対応拡張子
  INVALID_EXTENSIONS: [
    '.mp4', '.webm', '.mkv', '.avi', '.mov', '.flv', '.wmv', '.m4v', '.svg', '.gif'
  ],

  // 有効とみなす画像拡張子
  VALID_IMAGE_EXTENSIONS: [
    '.jpg', '.jpeg', '.png', '.webp'
  ],

  // コメントアイコン・装飾画像・広告・ロゴ等の除外キーワード
  IGNORE_KEYWORDS: [
    'comment', 'chara', 'icon', 'avatar', 'res_', 'thumb_comment',
    'amazon.com', 'amazon-adsystem.com', 'm.media-amazon.com', 'ssl-images-amazon.com',
    'pochipp', 'pochipp-logo', 'plugins/pochipp',
    '/smilies/', 'emoji', 'counter', 'facebook.com', 'twitter.com',
    'line.me', 'hatena', 'share', 'clear.gif',
    'blank.gif', 'pixel', 'ad_banner', 'logo_publisher'
  ]
};
```

---

## 🎯 まとめ

本スクリプトとGitHub Actions自動デプロイ環境を使えば、GitHub上のコードを更新するだけで簡単にGASの処理をアップデートできます。

特定サイトで画像が拾えない・余計な画像が混ざる場合は、Config.gs の **IGNORE_KEYWORDS** に単語を追加して調整してみてください。

---

## 📄 ライセンス
[MIT License](https://github.com/elysion-pre/NotifyRSS/blob/main/LICENSE)
