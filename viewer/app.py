#!/usr/bin/env python3
"""Streamlit viewer for competitor analysis exports.

このアプリはGoogle APIを一切呼び出しません。collector/export.py が事前に生成した
Excel/CSV、またはブラウザからアップロードされたファイルを読み込んで表示するだけです。
APIキーは不要で、公開デプロイしても安全です。
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pydeck as pdk
import streamlit as st

APP_DIR = Path(__file__).resolve().parent

# ファイル探索対象: アプリと同階層、次にdata/（開発時のデフォルト運用・サンプルデータ格納先）
SEARCH_DIRS = [APP_DIR, APP_DIR / "data"]

STORE_SHEET = "店舗"
ATTR_PREFIX = "属性_"
NO_OWN_STORE = "（選択しない）"

# 差別化分析のしきい値
COMMON_RATE = 0.6  # これ以上の競合が持つ属性 = 持っていないと不利
RARE_RATE = 0.2  # これ以下の競合しか持たない属性 = 打ち出せば差別化になる
MIN_REGISTERED = 5  # 登録済みの競合がこれ未満の属性は判定しない

BASIC_PROFILE_ITEMS = {
    "電話番号": "電話番号",
    "公式サイト/SNS": "Webサイト",
    "営業時間": "営業時間",
    "価格帯": "価格帯",
}

OWN_STATUS_COLORS = {
    "あり": "#2e7d32",
    "なし": "#c62828",
    "未登録": "#9e9e9e",
    "自店未選択": "#1565c0",
}


def find_distance_col(df: pd.DataFrame) -> Optional[str]:
    for col in df.columns:
        if "からの距離(m)" in str(col):
            return str(col)
    return None


def infer_origin_label(distance_col: str) -> str:
    return distance_col.replace("からの距離(m)", "")


def find_lat_lng_cols(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    lat_candidates = ["緯度", "lat", "latitude"]
    lng_candidates = ["経度", "lng", "lon", "longitude"]

    lat_col = next((c for c in df.columns if str(c).lower() in lat_candidates), None)
    lng_col = next((c for c in df.columns if str(c).lower() in lng_candidates), None)

    if lat_col and lng_col:
        return str(lat_col), str(lng_col)
    return None, None


def normalize_scores(series: pd.Series, reverse: bool = False) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    min_v = s.min(skipna=True)
    max_v = s.max(skipna=True)

    if pd.isna(min_v) or pd.isna(max_v) or min_v == max_v:
        out = pd.Series([0.5] * len(s), index=s.index)
    else:
        out = (s - min_v) / (max_v - min_v)

    if reverse:
        out = 1 - out

    return out.fillna(0.0)


def build_competitor_score(df: pd.DataFrame, distance_col: str, w_dist: float, w_rate: float, w_reviews: float) -> pd.DataFrame:
    out = df.copy()

    out[distance_col] = pd.to_numeric(out[distance_col], errors="coerce")
    out["Google評価"] = pd.to_numeric(out["Google評価"], errors="coerce")
    out["口コミ件数"] = pd.to_numeric(out["口コミ件数"], errors="coerce")

    out["distance_score"] = normalize_scores(out[distance_col], reverse=True)
    out["rating_score"] = normalize_scores(out["Google評価"], reverse=False)
    out["reviews_score"] = normalize_scores((out["口コミ件数"].fillna(0) + 1).map(math.log), reverse=False)

    weight_sum = w_dist + w_rate + w_reviews
    if weight_sum <= 0:
        w_dist, w_rate, w_reviews = 0.4, 0.35, 0.25
        weight_sum = 1.0

    out["競合スコア"] = (
        out["distance_score"] * w_dist
        + out["rating_score"] * w_rate
        + out["reviews_score"] * w_reviews
    ) / weight_sum

    return out


def score_band(score: float) -> str:
    if score >= 0.67:
        return "高"
    if score >= 0.34:
        return "中"
    return "低"


def build_circle_path(lat: float, lng: float, radius_m: int, points: int = 72) -> List[List[float]]:
    coords: List[List[float]] = []
    earth_radius = 6378137.0
    lat_rad = math.radians(lat)

    for i in range(points + 1):
        theta = 2 * math.pi * i / points
        dy = radius_m * math.sin(theta)
        dx = radius_m * math.cos(theta)

        d_lat = (dy / earth_radius) * (180 / math.pi)
        d_lng = (dx / (earth_radius * math.cos(lat_rad))) * (180 / math.pi)
        coords.append([lng + d_lng, lat + d_lat])

    return coords


def list_export_files() -> List[str]:
    seen: dict[str, float] = {}
    for d in SEARCH_DIRS:
        if not d.is_dir():
            continue
        for pattern in ("*.csv", "*.xlsx"):
            for p in d.glob(pattern):
                if p.name.startswith("."):
                    continue
                seen[str(p)] = p.stat().st_mtime
    return [path for path, _ in sorted(seen.items(), key=lambda kv: kv[1], reverse=True)]


def resolve_path(path_text: str) -> Path:
    p = Path(path_text).expanduser()
    if p.is_absolute():
        return p
    return (APP_DIR / p).resolve()


def read_sheets(name: str, data) -> Dict[str, pd.DataFrame]:
    """CSVなら店舗シートのみ、Excelなら全シートを返す。店舗シートが無い旧形式は先頭シートを店舗扱いにする。"""
    if str(name).lower().endswith(".csv"):
        return {STORE_SHEET: pd.read_csv(data, encoding="utf-8-sig")}
    sheets = pd.read_excel(data, sheet_name=None)
    if STORE_SHEET not in sheets:
        first = next(iter(sheets))
        sheets[STORE_SHEET] = sheets.pop(first)
    return sheets


@st.cache_data(show_spinner=False)
def read_sheets_cached(path_str: str) -> Dict[str, pd.DataFrame]:
    return read_sheets(path_str, path_str)


def attribute_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if str(c).startswith(ATTR_PREFIX)]


def attribute_frame(df: pd.DataFrame) -> pd.DataFrame:
    """属性列（あり/なし/空=未登録）と深夜営業（○/×）を、属性名 -> あり/なし/'' の表にそろえる。"""
    out = pd.DataFrame(index=df.index)
    for col in attribute_cols(df):
        out[col.removeprefix(ATTR_PREFIX)] = df[col].fillna("").astype(str).str.strip()
    if "深夜営業" in df.columns:
        out["深夜営業"] = df["深夜営業"].map({"○": "あり", "×": "なし"}).fillna("")
    return out


def judge_attribute(rate: Optional[float], registered: int, own: Optional[str]) -> str:
    if registered < MIN_REGISTERED or rate is None:
        return "データ不足"
    if own is None:
        if rate >= COMMON_RATE:
            return "多くの競合が対応"
        if rate <= RARE_RATE:
            return "対応している競合は少数"
        return "—"
    if own == "あり":
        if rate <= RARE_RATE:
            return "アピールポイント"
        if rate >= COMMON_RATE:
            return "対応済み（説明文にも明記）"
        return "対応済み"
    if rate >= COMMON_RATE:
        return "要対応（未登録）" if own == "" else "要対応"
    if own == "":
        return "未登録"
    if rate <= RARE_RATE:
        return "差別化の候補"
    return "—"


def build_attribute_table(competitors: pd.DataFrame, own_row: Optional[pd.Series]) -> pd.DataFrame:
    comp_attrs = attribute_frame(competitors)
    own_attrs = attribute_frame(own_row.to_frame().T) if own_row is not None else None
    records = []
    for attr in comp_attrs.columns:
        values = comp_attrs[attr]
        has = int((values == "あり").sum())
        registered = int(values.isin(["あり", "なし"]).sum())
        rate = has / registered if registered else None
        own = own_attrs[attr].iloc[0] if own_attrs is not None else None
        records.append(
            {
                "属性": attr,
                "競合の保有率(%)": round(rate * 100, 1) if rate is not None else None,
                "保有店数": has,
                "登録済みの競合数": registered,
                "自店": ("未登録" if own == "" else own) if own is not None else "自店未選択",
                "判定": judge_attribute(rate, registered, own),
            }
        )
    return pd.DataFrame(records).sort_values("競合の保有率(%)", ascending=False, na_position="last")


def profile_scores(df: pd.DataFrame) -> pd.DataFrame:
    """API上で確認できる項目に基づくプロフィール充実度（0-100）。基本項目60点 + 属性の登録率40点。"""
    out = pd.DataFrame(index=df.index)
    basic_cols = [c for c in BASIC_PROFILE_ITEMS if c in df.columns]
    if basic_cols:
        filled = df[basic_cols].fillna("").astype(str).apply(lambda s: s.str.strip() != "")
        basic_ratio = filled.mean(axis=1)
    else:
        basic_ratio = pd.Series(0.0, index=df.index)
    attrs = attribute_frame(df).drop(columns=["深夜営業"], errors="ignore")
    attr_ratio = attrs.isin(["あり", "なし"]).mean(axis=1) if not attrs.empty else pd.Series(0.0, index=df.index)
    out["プロフィール充実度"] = (basic_ratio * 60 + attr_ratio * 40).round(1)
    return out


def render_map(
    df: pd.DataFrame,
    origin_lat: float,
    origin_lng: float,
    radius_m: int,
    dot_size: int,
    own_row: Optional[pd.Series],
) -> None:
    map_df = df.dropna(subset=["緯度", "経度"]).copy()
    rank_scale = {"高": 1.15, "中": 1.0, "低": 0.9}
    map_df["radius"] = map_df["競合ランク"].map(rank_scale).fillna(1.0) * float(dot_size)

    color_map = {"高": [220, 20, 60, 180], "中": [255, 165, 0, 160], "低": [30, 144, 255, 140]}
    map_df["color"] = map_df["競合ランク"].map(color_map)
    map_df["color"] = map_df["color"].apply(
        lambda v: v if isinstance(v, list) else [128, 128, 128, 120]
    )

    layers = []
    circle_data = [{"path": build_circle_path(origin_lat, origin_lng, r), "radius": r} for r in [300, 500, 1000]]
    layers.append(
        pdk.Layer(
            "PathLayer",
            data=circle_data,
            get_path="path",
            get_width=2,
            get_color=[70, 70, 70, 120],
            width_min_pixels=1,
        )
    )
    layers.append(
        pdk.Layer(
            "ScatterplotLayer",
            data=map_df[["店舗名", "緯度", "経度", "radius", "color"]],
            get_position="[経度, 緯度]",
            get_radius="radius",
            get_fill_color="color",
            pickable=True,
            stroked=False,
        )
    )
    layers.append(
        pdk.Layer(
            "ScatterplotLayer",
            data=[{"店舗名": "拠点", "緯度": origin_lat, "経度": origin_lng}],
            get_position="[経度, 緯度]",
            get_radius=max(20, int(dot_size * 1.8)),
            get_fill_color=[0, 0, 0, 220],
            pickable=True,
        )
    )
    if own_row is not None and pd.notna(own_row.get("緯度")) and pd.notna(own_row.get("経度")):
        layers.append(
            pdk.Layer(
                "ScatterplotLayer",
                data=[{"店舗名": f"自店: {own_row['店舗名']}", "緯度": float(own_row["緯度"]), "経度": float(own_row["経度"])}],
                get_position="[経度, 緯度]",
                get_radius=max(24, int(dot_size * 2.2)),
                get_fill_color=[0, 200, 83, 240],
                get_line_color=[255, 255, 255, 255],
                line_width_min_pixels=2,
                stroked=True,
                pickable=True,
            )
        )

    view = pdk.ViewState(latitude=origin_lat, longitude=origin_lng, zoom=15 if radius_m <= 500 else 14, pitch=0)
    st.pydeck_chart(pdk.Deck(layers=layers, initial_view_state=view, tooltip={"text": "{店舗名}"}))


def render_differentiation_tab(competitors: pd.DataFrame, own_row: Optional[pd.Series]) -> None:
    st.subheader("差別化分析")
    if own_row is None:
        st.info("サイドバーで「自店舗」を選ぶと、自店と周辺競合の差（要対応・アピールポイント）を判定します。")

    has_attrs = bool(attribute_cols(competitors))
    if not has_attrs:
        st.warning(
            "このファイルには属性データ（属性_○○列）がありません。最新のcollectorで出力し直すと、"
            "属性の保有率・プロフィール充実度が表示されます。"
        )

    attr_table = build_attribute_table(competitors, own_row) if (has_attrs or "深夜営業" in competitors.columns) else None

    # 1. 説明文に入れる言葉の候補（自店選択時のみ）
    if own_row is not None and attr_table is not None:
        st.markdown("#### 説明文・プロフィールに反映する候補")
        appeal = attr_table[attr_table["判定"] == "アピールポイント"]["属性"].tolist()
        must_state = attr_table[attr_table["判定"] == "対応済み（説明文にも明記）"]["属性"].tolist()
        todo = attr_table[attr_table["判定"].str.startswith("要対応")]["属性"].tolist()
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.markdown("**打ち出す強み**（自店にあり、競合の" + f"{int(RARE_RATE * 100)}%以下しか持たない）")
            st.markdown("\n".join(f"- {a}" for a in appeal) if appeal else "該当なし")
        with col_b:
            st.markdown("**明記して取りこぼさない**（多くの競合が持ち、自店にもある）")
            st.markdown("\n".join(f"- {a}" for a in must_state) if must_state else "該当なし")
        with col_c:
            st.markdown("**対応・登録を検討**（" + f"{int(COMMON_RATE * 100)}%以上の競合が持つが、自店はなし/未登録）")
            st.markdown("\n".join(f"- {a}" for a in todo) if todo else "該当なし")

    # 2. 属性の保有率
    if attr_table is not None:
        st.markdown("#### 属性の保有率（周辺競合のうち、その属性が「あり」の割合）")
        chart_df = attr_table.dropna(subset=["競合の保有率(%)"])
        fig = px.bar(
            chart_df.sort_values("競合の保有率(%)"),
            x="競合の保有率(%)",
            y="属性",
            color="自店",
            orientation="h",
            color_discrete_map=OWN_STATUS_COLORS,
            hover_data={"登録済みの競合数": True, "判定": True},
        )
        fig.add_vline(x=COMMON_RATE * 100, line_dash="dot", line_color="#888")
        fig.add_vline(x=RARE_RATE * 100, line_dash="dot", line_color="#888")
        fig.update_layout(height=max(360, 24 * len(chart_df)), legend_title_text="自店の状況")
        st.plotly_chart(fig, width="stretch")
        st.caption(
            f"点線は{int(RARE_RATE * 100)}%と{int(COMMON_RATE * 100)}%。保有率は「未登録」の店を除いて計算しています"
            f"（登録済みの競合が{MIN_REGISTERED}店未満の属性は「データ不足」）。"
        )
        st.dataframe(attr_table, width="stretch", hide_index=True)

    # 3. カテゴリ分布
    if "カテゴリ" in competitors.columns and competitors["カテゴリ"].fillna("").astype(str).str.strip().ne("").any():
        st.markdown("#### メインカテゴリの分布")
        cats = competitors["カテゴリ"].fillna("").astype(str).str.strip()
        counts = cats[cats != ""].value_counts().head(15).rename_axis("カテゴリ").reset_index(name="店舗数")
        own_cat = str(own_row.get("カテゴリ", "")).strip() if own_row is not None else ""
        counts["自店のカテゴリ"] = counts["カテゴリ"].eq(own_cat).map({True: "自店と同じ", False: "その他"})
        fig_cat = px.bar(
            counts.sort_values("店舗数"),
            x="店舗数",
            y="カテゴリ",
            orientation="h",
            color="自店のカテゴリ",
            color_discrete_map={"自店と同じ": "#2e7d32", "その他": "#90a4ae"},
        )
        fig_cat.update_layout(height=max(320, 26 * len(counts)))
        st.plotly_chart(fig_cat, width="stretch")
        if own_cat:
            same = int((cats == own_cat).sum())
            st.caption(
                f"自店のメインカテゴリ「{own_cat}」の競合は {same}店（周辺の{same / max(len(cats), 1) * 100:.0f}%）。"
                "メインカテゴリは順位への影響が大きい要素です。競合が集中している場合は、より具体的なカテゴリや"
                "サブカテゴリの追加で検索される場面を広げられないか検討してください。"
            )
    else:
        st.caption("カテゴリ列がないため、カテゴリ分布は表示できません（最新のcollectorで出力すると表示されます）。")

    # 4. プロフィール充実度
    if has_attrs:
        st.markdown("#### プロフィール充実度")
        comp_scores = profile_scores(competitors)["プロフィール充実度"]
        if own_row is not None:
            own_score = float(profile_scores(own_row.to_frame().T)["プロフィール充実度"].iloc[0])
            m1, m2, m3 = st.columns(3)
            m1.metric("自店", f"{own_score:.0f}点")
            m2.metric("競合の中央値", f"{comp_scores.median():.0f}点")
            m3.metric("自店より充実している競合", f"{(comp_scores > own_score).mean() * 100:.0f}%")

            missing_basic = [
                label for col, label in BASIC_PROFILE_ITEMS.items()
                if col in own_row.index and str(own_row.get(col, "") if pd.notna(own_row.get(col)) else "").strip() == ""
            ]
            comp_attr = attribute_frame(competitors).drop(columns=["深夜営業"], errors="ignore")
            own_attr = attribute_frame(own_row.to_frame().T).drop(columns=["深夜営業"], errors="ignore")
            registered_rate = comp_attr.isin(["あり", "なし"]).mean()
            missing_attrs = [a for a in own_attr.columns if own_attr[a].iloc[0] == "" and registered_rate.get(a, 0) >= 0.5]
            st.markdown("**自店の未入力項目**")
            st.write("・基本情報: " + ("、".join(missing_basic) if missing_basic else "なし"))
            st.write("・属性（競合の半数以上が登録済みなのに自店は未登録）: " + ("、".join(missing_attrs) if missing_attrs else "なし"))
        else:
            st.metric("競合の中央値", f"{comp_scores.median():.0f}点")
        st.caption(
            "Places APIで確認できる項目（電話番号・Webサイト・営業時間・価格帯・属性の登録状況）に基づく目安です。"
            "ビジネスの説明文や写真枚数など、APIで取得できない項目は含みません。"
        )


st.set_page_config(page_title="競合分析ダッシュボード", page_icon="📍", layout="wide")
st.title("競合分析ダッシュボード")
st.caption("MEO対策のための周辺競合・差別化分析（距離・評価・口コミ・属性・カテゴリ）")

with st.sidebar:
    st.header("設定")
    uploaded_file = st.file_uploader(
        "調査結果ファイルをアップロード（Excel/CSV）",
        type=["csv", "xlsx"],
        help=(
            "アップロードしたファイルはこのブラウザセッション内でのみ使用され、"
            "サーバーに保存されません。collector/export.py の出力をそのまま使えます。"
        ),
    )

    export_files = list_export_files()
    if uploaded_file is None:
        if export_files:
            selected_file = st.selectbox("または検出したファイルを使う", export_files, index=0)
        else:
            selected_file = ""
            st.info(
                "ファイルが自動検出されませんでした。上のアップローダーから、"
                "collectorで出力したExcel/CSVを読み込んでください。"
            )
        manual_path = st.text_input("またはサーバー上のパスを直接指定（任意）", value="")
    else:
        selected_file = ""
        manual_path = ""


if uploaded_file is not None:
    try:
        sheets = read_sheets(uploaded_file.name, uploaded_file)
    except Exception as exc:  # pragma: no cover
        st.error(f"ファイル読み込みに失敗しました: {exc}")
        st.stop()
else:
    effective_path = manual_path.strip() or selected_file
    target_file = resolve_path(effective_path) if effective_path else Path("")

    if not effective_path or not target_file.exists():
        st.error("分析対象のファイルが見つかりません。サイドバーでファイルをアップロード・選択・指定してください。")
        st.stop()

    try:
        sheets = read_sheets_cached(str(target_file))
    except Exception as exc:  # pragma: no cover
        st.error(f"ファイル読み込みに失敗しました: {exc}")
        st.stop()

df_raw = sheets[STORE_SHEET]

distance_col = find_distance_col(df_raw)
if distance_col is None:
    st.error("距離列（xxxからの距離(m)）が見つかりません。")
    st.stop()
assert distance_col is not None

required_cols = ["店舗名", "Google評価", "口コミ件数", "ジャンル", "住所"]
missing = [c for c in required_cols if c not in df_raw.columns]
if missing:
    st.error(f"必要列が不足しています: {', '.join(missing)}")
    st.stop()

origin_label = infer_origin_label(distance_col)

lat_col, lng_col = find_lat_lng_cols(df_raw)
df = df_raw.copy()
if lat_col and lng_col:
    df["緯度"] = pd.to_numeric(df[lat_col], errors="coerce")
    df["経度"] = pd.to_numeric(df[lng_col], errors="coerce")

# 閉業店は競合から除外する
if "営業状態" in df.columns:
    df = df[df["営業状態"].fillna("") != "閉業"]

# 拠点座標: エクスポート済みの列があればそれを使用（推奨）。
# 無い古い形式のファイルの場合は店舗座標の平均で代用する。
origin_latlng: Optional[Tuple[float, float]] = None
if "拠点緯度" in df.columns and "拠点経度" in df.columns:
    o_lat = pd.to_numeric(df["拠点緯度"], errors="coerce").dropna()
    o_lng = pd.to_numeric(df["拠点経度"], errors="coerce").dropna()
    if not o_lat.empty and not o_lng.empty:
        origin_latlng = (float(o_lat.iloc[0]), float(o_lng.iloc[0]))

origin_address = ""
if "拠点住所" in df.columns:
    addr_series = df["拠点住所"].dropna()
    if not addr_series.empty:
        origin_address = str(addr_series.iloc[0])

if origin_latlng is None and "緯度" in df.columns and "経度" in df.columns:
    mean_lat = df["緯度"].mean(skipna=True)
    mean_lng = df["経度"].mean(skipna=True)
    if pd.notna(mean_lat) and pd.notna(mean_lng):
        origin_latlng = (float(mean_lat), float(mean_lng))
        st.info("拠点座標の列が見つからないため、店舗座標の平均を地図中心として使用しています。")

df[distance_col] = pd.to_numeric(df[distance_col], errors="coerce")

with st.sidebar:
    st.caption(f"拠点: {origin_address or origin_label}")

    # 自店の選択（「自店」列に○がある行を初期値にする）
    store_options = [NO_OWN_STORE] + df.sort_values(distance_col).index.tolist()
    default_index = 0
    if "自店" in df.columns:
        marked = df.index[df["自店"].fillna("").astype(str).str.strip() == "○"].tolist()
        if marked:
            default_index = store_options.index(marked[0])
    own_choice = st.selectbox(
        "自店舗を選択",
        store_options,
        index=default_index,
        format_func=lambda v: v if v == NO_OWN_STORE else f"{df.at[v, '店舗名']}（{int(df.at[v, distance_col]) if pd.notna(df.at[v, distance_col]) else '-'}m）",
        help="選んだ店舗を自店として強調表示し、周辺競合との差を分析します。",
    )

    st.markdown("---")
    w_dist = st.slider("重み: 距離", min_value=0.0, max_value=1.0, value=0.40, step=0.05)
    w_rate = st.slider("重み: Google評価", min_value=0.0, max_value=1.0, value=0.35, step=0.05)
    w_reviews = st.slider("重み: 口コミ件数", min_value=0.0, max_value=1.0, value=0.25, step=0.05)
    st.caption("競合スコアの推奨: 距離40% / 評価35% / 口コミ25%")

scored = build_competitor_score(df, distance_col, w_dist, w_rate, w_reviews)
scored["競合ランク"] = scored["競合スコア"].apply(score_band)

own_row: Optional[pd.Series] = None if own_choice == NO_OWN_STORE else scored.loc[own_choice]

with st.sidebar:
    radius_filter = st.selectbox("分析半径", [300, 500, 1000], index=1)
    max_points = st.slider("地図表示件数（上位スコア）", min_value=30, max_value=300, value=120, step=10)
    dot_size = st.slider("地図のドットサイズ", min_value=8, max_value=28, value=12, step=2)
    min_rating = st.slider("評価の下限", min_value=0.0, max_value=5.0, value=0.0, step=0.1)
    min_reviews = st.number_input("口コミ件数の下限", min_value=0, value=0, step=10)
    genres = sorted([g for g in scored["ジャンル"].dropna().unique().tolist()])
    selected_genres = st.multiselect("ジャンル", genres, default=genres)

filtered = scored[
    (scored[distance_col].fillna(999999) <= radius_filter)
    & (scored["Google評価"].fillna(0) >= min_rating)
    & (scored["口コミ件数"].fillna(0) >= min_reviews)
    & (scored["ジャンル"].isin(selected_genres))
].copy()

# 競合 = フィルタ後の店舗から自店を除いたもの
competitors = filtered.drop(index=own_choice, errors="ignore") if own_row is not None else filtered

if competitors.empty:
    st.warning("条件に一致する店舗がありません。フィルタ条件を緩めてください。")
    st.stop()

if origin_latlng is None:
    origin_lat = float(pd.to_numeric(filtered["緯度"], errors="coerce").mean()) if "緯度" in filtered.columns else 34.726893
    origin_lng = float(pd.to_numeric(filtered["経度"], errors="coerce").mean()) if "経度" in filtered.columns else 135.501036
else:
    origin_lat, origin_lng = origin_latlng

# Summary
c1, c2, c3, c4 = st.columns(4)
c1.metric("競合店舗数", f"{len(competitors):,}")
c2.metric(
    "平均Google評価",
    f"{competitors['Google評価'].mean(skipna=True):.2f}",
    delta=(f"自店 {own_row['Google評価']:.1f}" if own_row is not None and pd.notna(own_row["Google評価"]) else None),
    delta_color="off",
)
c3.metric(
    "平均口コミ件数",
    f"{competitors['口コミ件数'].mean(skipna=True):.1f}",
    delta=(f"自店 {int(own_row['口コミ件数'])}件" if own_row is not None and pd.notna(own_row["口コミ件数"]) else None),
    delta_color="off",
)
c4.metric("高競合比率", f"{(competitors['競合ランク'] == '高').mean() * 100:.1f}%")

tab_map, tab_diff, tab_scatter, tab_rank = st.tabs(["地図", "差別化分析", "散布図", "ランキング"])

with tab_map:
    st.subheader("拠点中心マッピング")
    st.caption("凡例: 赤=高競合 / 橙=中競合 / 青=低競合 / 緑=自店 / 黒=拠点。地図はスコア上位から表示。")
    if "緯度" in competitors.columns and "経度" in competitors.columns:
        map_view_df = competitors.sort_values("競合スコア", ascending=False).head(max_points)
        render_map(map_view_df, origin_lat, origin_lng, radius_filter, dot_size, own_row)
        st.caption(f"地図表示: {len(map_view_df)}件 / フィルタ後 {len(competitors)}件")
    else:
        st.info("緯度/経度列がないため地図を描画できません。")

with tab_diff:
    render_differentiation_tab(competitors, own_row)

with tab_scatter:
    st.subheader("距離 × Google評価 × 口コミ件数")
    scatter_df = competitors.copy()
    scatter_df["口コミ件数_plot"] = pd.to_numeric(scatter_df["口コミ件数"], errors="coerce").fillna(0).clip(lower=0) + 1
    fig = px.scatter(
        scatter_df,
        x=distance_col,
        y="Google評価",
        size="口コミ件数_plot",
        color="ジャンル",
        hover_name="店舗名",
        hover_data={"競合スコア": ":.3f", "競合ランク": True, "住所": True},
        size_max=36,
    )
    if own_row is not None:
        fig.add_trace(
            go.Scatter(
                x=[own_row[distance_col]],
                y=[own_row["Google評価"]],
                mode="markers+text",
                marker={"symbol": "star", "size": 22, "color": "#00c853", "line": {"width": 1, "color": "white"}},
                text=["自店"],
                textposition="top center",
                name="自店",
                hovertext=[own_row["店舗名"]],
            )
        )
    fig.update_layout(height=520)
    st.plotly_chart(fig, width="stretch")

with tab_rank:
    st.subheader("競合ランキング")
    top_n = st.slider("表示件数", min_value=1, max_value=min(100, len(competitors)), value=min(30, len(competitors)))
    rank_cols = [
        c for c in [
            "店舗名", "ジャンル", "カテゴリ", distance_col, "Google評価", "口コミ件数",
            "競合スコア", "競合ランク", "価格帯", "深夜営業", "住所",
        ]
        if c in competitors.columns
    ]
    rank_df = competitors.sort_values("競合スコア", ascending=False)[rank_cols].head(top_n)
    st.dataframe(rank_df, width="stretch")
    if own_row is not None:
        better = int((competitors["競合スコア"] > own_row["競合スコア"]).sum())
        st.caption(f"自店の競合スコア: {own_row['競合スコア']:.3f}（競合{len(competitors)}店中、自店より高い店は{better}店）")
