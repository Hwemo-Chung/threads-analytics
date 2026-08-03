# Threads Analyzer

[한국어](./README.md) · [English](./README.en.md) · **日本語**

Meta Threads API で自分のアカウントの投稿を全件収集し、12シートの Excel レポートを作ります。ランキング、時間帯別の成績、730日分の日次表示回数、バイラルの集中度、フォロワー属性まで。

**Threads では平均値が嘘をつきます。** バイラル1件が平均を引き上げ、結論をひっくり返すからです。このツールは平均と並べて中央値とデッド率を出します。平均で最良に見えた時間帯が、中央値では最悪だったということが実際に起きます。

作者自身のアカウント（オリジナル投稿1,954件、累計表示440万）での実測値：

| | |
|---|---|
| 表示回数の中央値 | **412** |
| 表示回数の平均 | **2,254** — 中央値の5.5倍 |
| 上位1%の投稿が占める表示回数 | **48.3%** |
| 500回に届かなかった投稿 | **56.5%** |
| 10時（JST）— 投稿数2位の時間帯 | 中央値 **278**、デッド率 **70.0%** |
| 1時（JST） | 中央値 **666**、デッド率 **43.8%** |

## 機能

- **OAuth 2.0** — ローカルHTTPSサーバーで長期トークンを取得
- **投稿の全件収集** — ページネーション、投稿ごとのインサイトを並列取得＋キャッシュ
- **アカウントインサイト** — フォロワー数、30日表示回数、属性（国・都市・性別・年齢）
- **12シートのExcelレポート** — 下の表を参照
- **3言語出力** — 韓国語 / English / 日本語（`--lang`）
- **トークン更新** — 60日の長期トークンを期限前に延長

## まず目で見る

Meta アプリなしで実際のレポートを作れます。サンプルデータを同梱しています。

```bash
pip install -r requirements.txt
python3 export_excel.py -i samples/sample_analysis.json -o sample.xlsx
```

シートの抜粋は **[docs/SAMPLE.md](./docs/SAMPLE.md)** にあります。
サンプルは合成データです。本文とアカウント名は架空で、数値の分布だけを実アカウントから取っています。

## 準備

### 1. Meta 開発者アプリ

自分のアプリが必要です。これは避けられません。Meta はアカウント所有者本人のアプリ経由でしかインサイトを開放していないためです。10分ほどかかります。

1. [Meta for Developers](https://developers.facebook.com/) でアプリを作成
2. **Use case** → 「Other」→「Consumer」
3. **Threads API** プロダクトを追加
   > ⚠️ 検索ボックスに「Threads」と入力・貼り付けしても出てこないことがあります。**ドロップダウンの「Threads API」を直接クリック**してください。
4. **App Settings > Basic** で `Threads App ID` と `Threads App Secret` を控える
5. **Threads API > Settings** でコールバックURLを3つとも登録：
   - Callback URL: `https://localhost:8888/callback`
   - Deauthorize URL: `https://localhost:8888/deauthorize`
   - Data Deletion URL: `https://localhost:8888/delete`
6. **Threads API > Permissions** で自分のThreadsアカウントをテスターに追加し、Threadsアプリ側で招待を承認

> ⚠️ Meta は3つとも **HTTPS 必須**です。セットアップが自己署名証明書を作るのはこのためです。

アプリの所有者は自分なので開発モードで動作し、**Meta のアプリ審査は不要**です。

### 2. Python 3.8 以上

## インストール

```bash
git clone https://github.com/Hwemo-Chung/threads-analytics.git
cd threads-analytics
python3 setup.py
```

`setup.py` が依存関係のインストール、SSL証明書の生成、`.env` の作成、OAuth認証までを行います。App ID と Secret を貼り付けるだけです。再実行しても既存のトークンと設定は保持されます。

<details>
<summary>手動セットアップ</summary>

```bash
cp .env.example .env
pip install -r requirements.txt

openssl req -x509 -newkey rsa:2048 \
  -keyout localhost.key -out localhost.crt \
  -days 365 -nodes -subj "/CN=localhost"

python3 auth.py
```

`.env` に記入：

```
THREADS_APP_ID=your_app_id_here
THREADS_APP_SECRET=your_app_secret_here
REDIRECT_URI=https://localhost:8888/callback
INSIGHTS_CACHE_TTL_DAYS=7
```
</details>

## 使い方

### 収集と分析

```bash
python3 analyze.py
```

全投稿をページネーションで取得し、投稿ごとのインサイトを5並列で収集、アカウントインサイトとフォロワー属性を取得します。トークン期限を事前警告し、429/5xx はバックオフ再試行、結果を `output/analysis_*.json` に保存します。

```bash
python3 analyze.py --refresh-insights     # キャッシュ無視で全件再取得
python3 analyze.py --ttl-days 30          # 今回だけTTLを変更
python3 analyze.py --max-posts 100        # サンプル取得
python3 analyze.py --skip-demographics
python3 analyze.py --fail-on-api-error    # APIエラー時に exit 2
python3 analyze.py --workers 3
python3 analyze.py --export-excel         # 分析後にExcelまで生成
```

> 約2,000件で初回は **10分** ほどかかります。2回目以降はキャッシュが効いて1〜2分です。

### Excelレポートの生成

```bash
python3 export_excel.py
python3 export_excel.py -i output/analysis_YYYYMMDD_HHMMSS.json
python3 export_excel.py -o path/to/report.xlsx
```

**出力言語：**

```bash
python3 export_excel.py --lang ja      # 日本語
python3 export_excel.py --lang en      # English
python3 export_excel.py                # 韓国語（既定）
```

`.env` に `THREADS_LANG=ja` を書いて固定もできます。`--lang` が優先されます。

シート名・見出し・判定ラベルに加えて、**成長戦略の生成文まで**翻訳されます。
**投稿本文と数値は収集したまま**です。

### 12のシート

| シート | 内容 |
|---|---|
| 全投稿 | 全件（日付・表示・いいね・エンゲージメント、フィルタ可） |
| ランキング | 表示 / エンゲージメント / いいね率 / バイラル の各TOP 30 |
| 時間帯分析 | 24時間 × 曜日、**中央値**とデッド率（表示500未満の割合） |
| フォロワー属性 | 国・都市・性別・年齢層 |
| 成長インサイト | KPI要約、メディアタイプ別成績、バイラルのパターン |
| バイラル詳細分析 | 固定値ではなくアカウント規模に合わせた閾値（P99）、表示の集中度 |
| いいね率詳細分析 | 分布、TOP/WORST 30、メディアタイプ別中央値 |
| コンテンツ最適化 | 文字数別の表示回数 vs エンゲージメント率、トピックタグ別成績 |
| 月次トレンド | 前月比、四半期サマリー |
| 日次表示回数 | 最大730日の日次データ、7日移動平均、表示日基準の月別 |
| スナップショット推移 | 実行のたびに蓄積されるフォロワー履歴 — **APIからは取得できないデータ** |
| 10万への戦略 | 現在地、今すぐやること5件、ロードマップ |

2つだけ補足します。

- **日次表示回数** は投稿日ではなく *表示日* 基準です。月次トレンドは投稿の累計表示を公開月に紐づけるため、直近の月を構造的に過小評価します。この2枚は意図的に別の問いに答えています。
- **スナップショット推移** が存在するのは、`followers_count` が `since`/`until` を渡してもスカラーしか返さないからです。APIにフォロワーの時系列はありません。ローカルのスナップショットが唯一の履歴なので、このシートを活かすには `analyze.py` を定期実行してください。

### トークン更新（60日ごと）

```bash
python3 refresh_token.py
```

`analyze.py` は期限の `TOKEN_WARN_DAYS`（既定7）日前に警告し、期限切れなら停止します。すでに切れている場合は `auth.py` を再実行してください。

### テキストアーカイブ

```bash
python3 archive.py
```

全投稿の本文を `output/저장함/` に Markdown・JSONL・CSV で書き出し、月別とメディアタイプ別に分けます。`analyze.py` 実行時に自動生成されます（`--skip-archive` で無効化）。

> Threads API に **保存済み投稿（他人の投稿のブックマーク）を取得するエンドポイントはありません**。このアーカイブは自分の投稿のみです。

### テスト

```bash
python3 -m unittest discover -s tests -v
```

ネットワークアクセスなし、シークレットなし。

## 既知の制限

- 長文のセルは列幅の上限（60）に達して途中で切れて見えます。セルをクリックすれば全文を確認できます。
- フォロワー属性は **100人以上** から提供されます。Meta の仕様であり不具合ではありません。
- 時間帯分析と文字数分析は、サンプルが30件未満の場合は推測せず判定を保留します。

## API メモ

- **レート制限** — 約4,800 calls/hour。429と5xxはバックオフ再試行
- **トークン** — 長期トークンは60日。`TOKEN_EXPIRES_AT` を `.env` に記録
- **インサイトキャッシュ** — v2 + TTL（`INSIGHTS_CACHE_TTL_DAYS`、既定7日）
- **未知のフィールドは黙って無視されます** — `GET /{user-id}/threads` は `fields=id,存在しない項目` でも 200 を返します。200はフィールド存在の証拠になりません。実投稿を並べてキーの有無を数えて確認してください。`/me` は未知フィールドでエラーを返しますが、`/threads` と `/replies` は返しません
- **APIに存在しないもの** — `saves`、`impressions`、`reach`、`profile_visits`、そしてフォロワーの日次時系列

## プライバシー

サーバーはありません。すべて自分のマシン上で動き、Meta API を直接呼び、`output/` に書き出します。トークンはローカルの `.env`（パーミッション600）に保存されます。`.env`、証明書、`output/` は gitignore 済みです。

## ライセンス

MIT
