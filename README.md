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

クライアントへの配布は **`CompetitorDashboardSetup.exe`（インストーラ）を渡すのが基本**です。
zip解凍もフォルダ管理も不要で、「ダブルクリック → インストールウィザード → デスクトップに
アイコン」という、非エンジニアに馴染みのある体験になります。

### インストーラの中身・挙動

- インストール先は `%LOCALAPPDATA%\CompetitorDashboard`（ユーザー単位）。管理者権限や
  UACプロンプトは不要。
- インストール時に空の `input` フォルダも一緒に作成される（README代わりの`readme.txt`入り）。
  `collector`で生成した最新の `places.csv`（または`.xlsx`）をこのフォルダに置けば、
  アプリ起動時に自動検出される。
- 完了画面の「起動する」にチェックが入った状態でインストールを終えると、そのまま
  ブラウザでダッシュボードが開く。
- スタートメニュー・デスクトップ（任意）にアイコンが作成され、Windowsの
  「アプリと機能」から通常のアプリと同じようにアンインストールできる
  （アンインストール時、ユーザーが`input`に追加したファイルは削除されない）。
- コマンドプロンプトは表示されず、初回起動時のメールアドレス入力を求められることもない
  （`~/.streamlit/credentials.toml` を起動時に自動生成して回避している）。

**アプリの終了方法**: コンソールを表示しない設計のため、ブラウザタブを閉じただけでは
裏でプロセスが起動したままになる。完全に終了させるには、タスクマネージャーで
`competitor-dashboard.exe` を終了するか、PCを再起動する必要がある。この点は
将来的にpywebviewベースの実装に切り替えることでネイティブウィンドウの
「閉じるボタン」で解決できる想定（未着手）。

### ビルド（GitHub Actionsで自動化・推奨）

Windows機を用意しなくても、[.github/workflows/build-exe.yml](.github/workflows/build-exe.yml) が
GitHub上のWindows runnerでPyInstaller実行 → Inno Setupインストーラ生成まで自動的に行います。

**タグをpushしてリリースとして公開する場合（クライアント配布はこちら）**

```bash
git tag v1.0.0
git push origin v1.0.0
```

`v` から始まるタグをpushすると、ビルド後に自動でGitHub Releaseが作成され、
**`CompetitorDashboardSetup.exe`（インストーラ、zipなし）と `competitor-dashboard.exe`
（生の単一exe、上級者向け）の両方が添付**されます。GitHub Releaseの添付ファイルは
zip化されずに公開されるため、クライアントは解凍が一切不要です。クライアントには
そのReleaseページのダウンロードリンクと`CompetitorDashboardSetup.exe`を使うよう案内してください。

タグのバージョン番号（`v1.0.0`の`1.0.0`部分）がそのままインストーラのバージョンとして
使われます。

**手動実行する場合（動作確認用）**

1. GitHubリポジトリの `Actions` タブ → `Build Windows exe` を選択
2. `Run workflow` ボタンを押す（`workflow_dispatch`）
3. 実行が終わったら、そのRunのページ下部 `Artifacts` から `CompetitorDashboardSetup`
   （インストーラ）または `competitor-dashboard-windows`（生exe、上級者向け）をダウンロード。
   いずれもzipを1回解凍すればexeが出てくる（Artifactsのダウンロードは仕組み上GitHub側で
   必ず1回zip化されるため、これ以上は減らせない）。この場合バージョンは`0.0.0-dev`になる。

ビルド生成物（`dist/`, `build/`, `installer/output/`）はリポジトリにcommitしません。
常にActions側で都度ビルドし、Artifacts/ReleasesからDLする運用にしてください。

### ローカル（Windows機）でビルドする場合

PyInstallerはクロスコンパイルできないため、Windows機で実行する必要があります。

```bash
uv sync --group dev
uv run pyinstaller viewer/build_exe.spec
```

出力: `dist/competitor-dashboard.exe`（onefile方式の単一ファイル。`_internal`のような
付随フォルダは生成されない）。続けて[Inno Setup](https://jrsoftware.org/isinfo.php)を
インストールした上で、以下でインストーラ化できる。

```bash
ISCC.exe /DMyAppVersion=1.0.0 installer\setup.iss
```

出力: `installer\output\CompetitorDashboardSetup.exe`

macOS上ではmacOS版バイナリとしてビルド・動作確認は可能ですが、Windows向け配布物としては
使えません（実際のexe/インストーラ生成は必ずWindows環境で行ってください）。onefile方式は
起動のたびに一時フォルダへ展開するため、初回表示まで数秒かかる点にも留意してください。

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
