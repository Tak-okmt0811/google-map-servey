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
├── verify_one.py          # API疎通確認用（1件だけ取得するテスト）
├── config.py               # 収集プロファイル(TOML)の読み込み
└── configs/
    └── bar_izakaya.toml    # 検索キーワード・ジャンル判定辞書などの設定例

viewer/                   # APIキー不要・公開デプロイ可能
├── app.py                 # Streamlitダッシュボード本体
└── data/
    └── sample_places.csv   # 公開ポートフォリオ用のダミーデータ（架空の店舗）

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

- `places_<ラベル>_<半径>m_<日時>.xlsx`（手元確認用。緯度・経度・拠点座標を含む）
- `data/places.csv`（同内容のCSV。viewerへのアップロード用）

いずれも `.gitignore` によりGit管理対象外です。実データをリポジトリにcommitしないでください。

---

## STEP 2: ダッシュボードの起動（viewer・開発時）

```bash
uv run streamlit run viewer/app.py
```

ブラウザで `http://localhost:8501` が開きます。サイドバーの「調査結果ファイルをアップロード」
から `collector` の出力（CSV/Excel）を選ぶと、その場で読み込んで表示されます
（アップロードされたファイルはディスクに保存されず、ブラウザセッション内のメモリでのみ扱われます）。

アップロードしなかった場合は、`viewer/`直下または`viewer/data/`にあるCSV/Excelを自動検出します
（`viewer/data/sample_places.csv` が既定のサンプルデータとして入っています）。

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
- `.gitignore` は `*.csv` を既定で除外しつつ、`viewer/data/sample_places.csv`（公開用ダミーデータ）
  のみ例外的にcommit対象としています。実データを `viewer/data/` に置く場合は
  ファイル名を変えるか、commit前に必ず `git status` で確認してください。
- 万一APIキーが漏洩した場合は、Google Cloud ConsoleでキーをRegenerate（失効・再発行）し、
  必要に応じてHTTPリファラ制限・API制限をかけてください。

---

## verify_one.py（API疎通確認）

```bash
uv run python collector/verify_one.py
```

Geocoding → Places検索 → Places詳細取得の3リクエストのみを発行し、APIキーとAPI有効化状況を
確認します。
