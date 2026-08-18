# 競合分析ツール

Google Places APIを使って周辺店舗を収集し、Streamlitで競合マップを表示するツールです。

`collector`（データ収集・自分のAPIキーを使用）と `viewer`（既存データの表示のみ・APIキー不要）を
明確に分離しています。**viewerは実行時にGoogle APIを一切呼び出しません。** クライアント向けexe配布や
一般公開の際にも、あなたのAPIキー・アカウントが使われることはありません。

---

## 構成

```
collector/                # 非公開・自分のPC(または管理者専用環境)でのみ実行
├── export.py              # 周辺店舗をAPIで取得しExcel・CSVに出力するCLI
├── verify_one.py          # API疎通確認用（1件だけ取得するテスト）
├── config.py               # 収集プロファイル(TOML)の読み込み
└── configs/
    └── bar_izakaya.toml    # 検索キーワード・ジャンル判定辞書などの設定例

viewer/                   # APIキー不要・公開/exe配布可能
├── app.py                 # Streamlitダッシュボード本体
├── launcher.py             # exe化用の起動エントリポイント
├── build_exe.spec          # PyInstaller用ビルド設定（Windows向け）
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

## STEP 1: 店舗データの取得（collector・非公開）

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

### 実際の調査エリアをローカルだけで保持する

クライアント案件など、実際の調査対象エリアをリポジトリに残したくない場合は、
ファイル名を `*.local.toml` にすると `.gitignore` で自動的に除外されます。

```bash
cp collector/configs/bar_izakaya.toml collector/configs/bar_izakaya.local.toml
# default_address を実際の調査エリアに書き換えてから
uv run python -m collector.export --config collector/configs/bar_izakaya.local.toml
```

```bash
cp collector/configs/bar_izakaya.toml collector/configs/my_profile.toml
# my_profile.toml を編集後
uv run python -m collector.export --config collector/configs/my_profile.toml --address "..."
```

### 出力ファイル

- `places_<ラベル>_<半径>m_<日時>.xlsx`（手元確認用）
- `data/places.csv`（緯度・経度・拠点座標を含む。viewer/exe配布用の実データ）

いずれも `.gitignore` によりGit管理対象外です。実データをリポジトリにcommitしないでください。

---

## STEP 2: ダッシュボードの起動（viewer・開発時）

```bash
uv run streamlit run viewer/app.py
```

ブラウザで `http://localhost:8501` が開きます。次の場所を自動的にスキャンし、
見つかったCSV/Excelをサイドバーの一覧から選べます。

1. `viewer/input/`（最優先）
2. `viewer/`（app.pyと同階層）
3. `viewer/data/`（開発時のデフォルト。`sample_places.csv` が入っています）

`collector`で生成した `data/places.csv` を表示したい場合は、`viewer/input/` にコピーするか、
サイドバーの「または直接パスを指定」に絶対パスを入力してください。

---

## 成果物① Windowsクライアント向けexe化

`viewer/` だけをexe化して配布します。**APIキーは一切含まれず、実行時も通信しません。**

### ビルド（Windows環境で実行）

PyInstallerはクロスコンパイルできないため、**Windows機（または Windows GitHub Actions runner）で
ビルドしてください。**

```bash
uv sync --group dev
uv run pyinstaller viewer/build_exe.spec
```

出力: `dist/competitor-dashboard/` フォルダ一式（`competitor-dashboard.exe` を含む）

### クライアントへの渡し方

1. `dist/competitor-dashboard/` フォルダをまるごと渡す（zip圧縮推奨）。
2. `collector`で生成した最新の `places.csv`（または`.xlsx`）を、そのフォルダ直下か
   `input/` サブフォルダに配置してもらう。
3. `competitor-dashboard.exe` をダブルクリックするとブラウザでダッシュボードが開く。
4. 調査結果を更新する場合は、`input/` フォルダの中身を新しいファイルに差し替えるだけでよい
   （exeの再ビルドは不要）。

macOS上ではmacOS版バイナリとしてビルド・動作確認は可能ですが、Windows向け配布物としては
使えません（実際のexe生成は必ずWindows環境で行ってください）。

### ビルド（GitHub Actionsで自動化する場合）

Windows機を用意しなくても、[.github/workflows/build-exe.yml](.github/workflows/build-exe.yml) が
GitHub上のWindows runnerで自動的にビルドします。

**手動実行する場合**

1. GitHubリポジトリの `Actions` タブ → `Build Windows exe` を選択
2. `Run workflow` ボタンを押す（`workflow_dispatch`）
3. 実行が終わったら、そのRunのページ下部 `Artifacts` から
   `competitor-dashboard-windows` をダウンロード（zip、`dist/competitor-dashboard/` 一式を圧縮したもの）

**タグをpushしてリリースとして公開する場合**

```bash
git tag v1.0.0
git push origin v1.0.0
```

`v` から始まるタグをpushすると、ビルド後に自動でGitHub Releaseが作成され、
zipが添付されます。クライアントには、そのReleaseページのダウンロードリンクを共有できます。

ビルド生成物（`dist/`, `build/`）はリポジトリにcommitしません。常にActions側で
都度ビルドし、Artifacts/ReleasesからDLする運用にしてください。

---

## 成果物② ポートフォリオ公開（無料ホスティング）

`viewer/`（+ `viewer/data/sample_places.csv`）だけを含む**公開用リポジトリ/ブランチ**を用意し、
Streamlit Community Cloud（無料）や Hugging Face Spaces（無料）にデプロイします。

- 表示するのは架空データ（`viewer/data/sample_places.csv`）のみ。実在店舗の情報や
  クライアントの実データは含めないでください。
- デプロイ設定のSecretsには**何も登録しない**でください。viewerはAPIキーを参照しないため不要です。
- リポジトリに `collector/` や `.env` を含めない（別リポジトリ、または `.gitignore` で除外）ことで、
  あなたのAPIキーが公開経路に紛れ込むリスクをなくせます。

### デプロイ手順（Streamlit Community Cloudの例）

1. `viewer/` を含む公開用リポジトリをGitHubにpush。
2. [share.streamlit.io](https://share.streamlit.io) でリポジトリを接続し、
   メインファイルパスに `viewer/app.py` を指定。
3. Secretsは設定不要（空のまま）。
4. デプロイ後のURLをポートフォリオに掲載。

---

## セキュリティに関する注意

- `collector`（APIキーを使う処理）と `viewer`（表示のみ）はコード上完全に分離されています。
  `viewer/app.py` はGoogle APIのURLやキー参照を一切含みません。
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
