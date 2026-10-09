# 競合分析ツール

Google Places APIを使って周辺店舗を収集し、Streamlitで競合マップを表示するツールです。

`collector`（データ収集・実行する人自身のAPIキーを使用）と `viewer`（既存データの表示のみ・
APIキー不要）を明確に分離しています。**viewerは実行時にGoogle APIを一切呼び出しません。**
一般公開しても、あなたのAPIキー・アカウントが使われることはありません。

このリポジトリは丸ごと公開しています。`viewer/app.py`のデモは架空データで動いていますが、
`collector`をcloneして**ご自身のGoogle Maps Platform APIキー**で実行すればExcelが出力でき、
それをデモのサイドバーからアップロードすると、ご自身のエリアの実データで動作を確認できます
（詳しくは下記「STEP 1」「STEP 2」を参照）。

---

## 構成

```
collector/                # コードは公開・実行は各自のAPIキー(.env)で行う
├── export.py              # 周辺店舗をAPIで取得しExcel・CSVに出力するCLI
├── usage.py               # 無料枠を超えないための呼び出し回数ガード
├── verify_one.py          # API疎通確認用（1件だけ取得するテスト）
├── config.py               # 収集プロファイル(TOML)の読み込み
└── configs/
    └── bar_izakaya.toml    # 検索キーワード・ジャンル判定辞書などの設定例

viewer/                   # APIキー不要・公開デプロイ可能
├── app.py                 # Streamlitダッシュボード本体
└── data/
    └── sample_places.xlsx  # 公開ポートフォリオ用のダミーデータ（架空の店舗）

scripts/
└── generate_sample_data.py # ダミーデータの生成（APIは使わない）

.env                       # collector専用のAPIキー（Gitにはcommitしない）
```

---

## 事前準備

### 1. 依存関係のインストール

```bash
uv sync
```

### 2. APIキーの設定（collectorを使う場合のみ）

プロジェクトルートに `.env` ファイルを作成します（**Gitにはcommitしないこと**）。

```
GOOGLE_MAPS_API_KEY=あなたのAPIキー
```

必要なGoogle Cloud APIは以下の2つです。

- **Geocoding API**（住所 → 緯度経度変換）
- **Google Places API**（周辺店舗検索・詳細取得）

`viewer` だけを使う（既存データを見るだけの）場合はAPIキーは不要です。

---

## STEP 1: 店舗データの取得（collector・要APIキー）

```bash
uv run python -m collector.export
```

デフォルトは `collector/configs/bar_izakaya.toml` の設定（バー・居酒屋向けキーワード、
汎用サンプル住所「大阪府大阪市北区梅田1丁目」中心・半径500m）で検索します。

| オプション | 説明 | 例 |
|---|---|---|
| `--config` | 収集プロファイル（検索キーワード・ジャンル判定辞書）のTOMLファイル | `--config collector/configs/beauty_salon.toml` |
| `--address` | 検索基点の住所（省略時は設定ファイルの`default_address`） | `--address "東京都渋谷区"` |
| `--label` | 出力列・ファイル名に使う拠点ラベル（省略時は`--address`をそのまま使用） | `--label 渋谷` |
| `--radius` | 半径（300 / 500 / 1000 m） | `--radius 1000` |
| `--output` | 出力Excelパス（省略時は日時付き自動生成） | `--output result.xlsx` |
| `--keywords` | 検索キーワードをカンマ区切りで上書き（小規模な動作確認向け） | `--keywords 居酒屋,bar` |
| `--estimate` | APIを呼ばずに、今月の使用回数と今回の最大呼び出し回数だけを表示 | `--estimate` |

### 費用（無料枠）と安全装置

Google Maps Platformは料金区分ごとに月間の無料枠があります（2025年3月以降の料金体系）。
collectorが使う区分と、1回の実行（7キーワード・半径500m）での呼び出し回数の目安は次のとおりです。

| 料金区分 | 無料枠/月 | 1回の実行での目安 |
|---|---|---|
| Geocoding | 10,000回 | 1回 |
| Text Search Enterprise | 1,000回 | 最大189回（キーワード数 × 検索地点9 × 最大3ページ） |
| Place Details Enterprise + Atmosphere | 1,000回 | 半径内の店舗数（西中島・500mで約200回） |

- `collector/usage.py` が今月の呼び出し回数を `.api_usage.json`（Git管理対象外）に記録し、
  無料枠の80%に達する呼び出しは**APIを呼ぶ前に中止**します（費用は発生しません）。
  キーワード検索は最大回数で、店舗詳細は件数が確定した時点で、事前に上限チェックします。
- 記録できるのはこのツールの呼び出しだけです。同じ請求先アカウントで他にもAPIを使っている場合は、
  Cloud Consoleで利用状況を確認し、APIごとのクォータ（1日の上限）と予算アラートも設定してください。
- 詳細取得は半径内の店舗だけに行います（検索ヒットのうち半径外の店は詳細を取得しない）。

### 業種・エリアを変えて使う

`collector/configs/bar_izakaya.toml` をコピーして、検索キーワード・ジャンル判定辞書・
サブジャンル/特徴キーワードを書き換えれば、飲食店以外の業種にも流用できます。

```bash
cp collector/configs/bar_izakaya.toml collector/configs/my_profile.toml
# my_profile.toml を編集後
uv run python -m collector.export --config collector/configs/my_profile.toml --address "..."
```

### 実際の調査エリアをローカルだけで保持する

クライアント案件など、実際の調査対象エリアをリポジトリに残したくない場合は、
ファイル名を `*.local.toml` にすると `.gitignore` で自動的に除外されます。

```bash
cp collector/configs/bar_izakaya.toml collector/configs/bar_izakaya.local.toml
# default_address を実際の調査エリアに書き換えてから
uv run python -m collector.export --config collector/configs/bar_izakaya.local.toml
```

### 出力ファイル

- `places_<ラベル>_<半径>m_<日時>.xlsx`（viewerへのアップロード用。以下の3シート構成）
  - `店舗`: 1店舗1行。カテゴリ、評価・口コミ件数、最新口コミ日、営業時間・深夜営業、
    属性（`属性_テラス席` など。値は あり/なし/空欄=未登録）、口コミ要約など
  - `キーワード順位`: 検索キーワードごとに、基点で検索したときの表示順位
    （Googleマップの実際の表示順位の目安。順位が空欄の店は周辺地点の検索でのみ見つかった店）
  - `口コミ`: 1店舗あたり最大5件の口コミ（投稿者名は保存しない）
- `data/places.csv`（店舗シートのみのCSV）

いずれも `.gitignore` によりGit管理対象外です。実データをリポジトリにcommitしないでください。

---

## STEP 2: ダッシュボードの起動（viewer・開発時）

```bash
uv run streamlit run viewer/app.py
```

ブラウザで `http://localhost:8501` が開きます。サイドバーの「調査結果ファイルをアップロード」
から `collector` の出力（Excel/CSV）を選ぶと、その場で読み込んで表示されます
（アップロードされたファイルはディスクに保存されず、ブラウザセッション内のメモリでのみ扱われます）。

アップロードしなかった場合は、`viewer/`直下または`viewer/data/`にあるExcel/CSVを自動検出します
（`viewer/data/sample_places.xlsx` が既定のサンプルデータとして入っています）。

### 自店舗の選択と差別化分析

サイドバーの「自店舗を選択」で自店を選ぶと、地図・散布図・ランキングで自店が強調表示され、
「差別化分析」タブで周辺競合との差を確認できます。

- **比較対象**: 自店と同じメインカテゴリの競合（初期値。10店以上いる場合）か、周辺の全競合かを選べます。
- **属性の保有率**: 属性ごとに、競合の何%が「あり」かを表示し、自店の状況と合わせて判定します。
  - 競合の60%以上が持つのに自店はなし/未登録 → **要対応**
  - 自店にあり、競合の20%以下しか持たない → **アピールポイント**
  - 多くの競合が持ち、自店にもある → **対応済み（説明文にも明記）**
- **メインカテゴリの分布**: 自店と同じカテゴリにどれだけ競合が集中しているか。
- **プロフィール充実度**: 電話番号・Webサイト・営業時間・価格帯・属性の登録状況から算出した目安。
  ビジネスの説明文や写真など、APIで取得できない項目は含みません。

`店舗`シートに`自店`列を作り、自店の行に`○`を入れておくと、その店舗が初期選択されます。

---

## デプロイ（Streamlit Community Cloud・無料）

`viewer/app.py`をクライアント・ポートフォリオ閲覧者向けに公開します。APIキーは一切不要で、
Secretsに何も登録しなくてもそのまま動きます。

1. このリポジトリをGitHubにpush（公開リポジトリでも問題ありません。理由は下記
   「セキュリティに関する注意」を参照）。
2. [share.streamlit.io](https://share.streamlit.io) にGitHubアカウントでログイン。
3. 「Create app」→ 対象リポジトリ・ブランチ（`main`）を選択し、
   **Main file path に `viewer/app.py`** を指定してデプロイ。
4. Secretsの設定は不要（空のまま）。viewerはGoogle APIを一切呼び出さないため。
5. 依存関係はリポジトリ直下の `uv.lock`（`pyproject.toml`）が自動的に使われます。
   もし依存解決でエラーが出る場合は、リポジトリ直下に以下のような`requirements.txt`を
   追加してください（Community Cloudはこちらを優先して使うようになります）。

   ```
   streamlit>=1.61.1
   pandas>=3.0.5
   plotly>=6.9.0
   pydeck>=0.9.3
   openpyxl>=3.1.5
   ```

6. デプロイ完了後に発行されるURL（`https://<任意名>.streamlit.app`）をポートフォリオに掲載。

### クライアントへの渡し方

`collector`で出力したExcel/CSVをメールやチャットでクライアントに送り、
上記でデプロイした公開URLを開いてサイドバーからアップロードしてもらうだけです。
インストールもzip解凍も不要で、ブラウザだけで完結します。

### 一般公開デモとしての使い方

ポートフォリオ訪問者は、`collector`をcloneして自分のGoogle Maps Platform APIキーで
Excelを出力し、同じ公開URLにアップロードすれば、架空のサンプルデータではなく
自分の指定したエリアの実データで動作を確認できます。クライアントに対する差別化は
「Excel出力（APIキー取得・実行）の手間をこちらが代行すること」です。

---

## セキュリティに関する注意

- `collector`（APIキーを使う処理）と `viewer`（表示のみ）はコード上完全に分離されています。
  `viewer/app.py` はGoogle APIのURLやキー参照を一切含みません。両方とも同じ公開リポジトリに
  含まれていますが、`collector`はcloneした人自身の`.env`（コミット対象外）を読むだけなので、
  あなたのAPIキーが公開経路や他人の実行に紛れ込むことはありません。
- `viewer/app.py`のファイルアップロード機能は、アップロードされたファイルをディスクに
  保存せず、そのブラウザセッション内のメモリ上でのみ処理します。公開URLを知っていれば
  誰でも同じ形式のファイルをアップロードして閲覧できる（クライアント専用ページにはならない）
  点は把握した上で運用してください。
- `.env` は `.gitignore` で除外されており、Git履歴上も一度もcommitされていないことを確認済みです。
  引き続きcommitしないよう注意してください。
- `.gitignore` は collectorの出力（`*.csv`・`places_*.xlsx`）を除外しています。公開用ダミーデータ
  `viewer/data/sample_places.xlsx` はcommit対象です。実データを `viewer/data/` に別名で置くと
  commit対象になってしまうため、commit前に必ず `git status` で確認してください。
- 万一APIキーが漏洩した場合は、Google Cloud ConsoleでキーをRegenerate（失効・再発行）し、
  必要に応じてHTTPリファラ制限・API制限をかけてください。

---

## verify_one.py（API疎通確認）

```bash
uv run python collector/verify_one.py
```

Geocoding → Places検索 → Places詳細取得の3リクエストのみを発行し、APIキーとAPI有効化状況を
確認します。
