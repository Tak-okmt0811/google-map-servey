# 競合分析ツール

Google Places APIを使って周辺店舗を収集し、Streamlitで競合マップを表示するツールです。

---

## 構成

| ファイル | 役割 |
|---|---|
| `google_places_export.py` | 周辺店舗をAPIで取得しExcel・CSVに出力 |
| `streamlit_competitor_analysis.py` | CSVを読み込んでブラウザでダッシュボード表示 |

---

## 事前準備

### 1. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 2. APIキーの設定

プロジェクトルートに `.env` ファイルを作成します（**Gitにはcommitしないこと**）。

```
GOOGLE_MAPS_API_KEY=あなたのAPIキー
```

必要なGoogle Cloud APIは以下の2つです。

- **Geocoding API**（住所 → 緯度経度変換）
- **Google Places API**（周辺店舗検索・詳細取得）

---

## STEP 1: 店舗データの取得

```bash
python google_places_export.py
```

デフォルトは「大阪府淀川市西中島3丁目」中心・半径500mで検索します。

| オプション | 説明 | 例 |
|---|---|---|
| `--address` | 検索基点の住所 | `--address "大阪府淀川市西中島3丁目"` |
| `--radius` | 半径（300 / 500 / 1000 m） | `--radius 1000` |
| `--output` | 出力Excelパス（省略時は日時付き自動生成） | `--output result.xlsx` |

実行後、以下の2ファイルが生成されます。

- `places_西中島3丁目_500m_YYYYMMDD_HHMMSS.xlsx`（手元確認用）
- `data/places.csv`（Streamlit用・**こちらをGitにcommitする**）

---

## STEP 2: ダッシュボードの起動

```bash
streamlit run streamlit_competitor_analysis.py
```

ブラウザで `http://localhost:8501` が開きます。

`data/places.csv` が存在すれば自動的に読み込まれます。  
**APIキーがなくても閲覧のみ可能**です（緯度経度はCSVに含まれているため）。

---

## Streamlit Community Cloudへのデプロイ（閲覧者への共有）

1. `data/places.csv` をGitHubリポジトリにpush（`.env` はpushしない）
2. [share.streamlit.io](https://share.streamlit.io) でリポジトリを接続
3. **Settings → Secrets** に以下を追加（地図の中心座標取得に使用）

   ```toml
   GOOGLE_MAPS_API_KEY = "あなたのAPIキー"
   ```

4. デプロイ後のURLをクライアントに共有

> **注意**: APIキーは `.gitignore` によりリポジトリには含まれません。Secrets機能で安全に管理してください。

---

## .gitignore 対象ファイル

以下はリポジトリに含めません。

```
.env                  # APIキー
places_*.xlsx         # 生成されたExcel
.geocode_cache.csv    # ジオコードキャッシュ
.venv/                # 仮想環境
```
