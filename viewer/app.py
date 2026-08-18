#!/usr/bin/env python3
"""Streamlit viewer for competitor analysis exports.

このアプリはGoogle APIを一切呼び出しません。collector/export.py が事前に生成した
CSV/Excel（緯度・経度・拠点座標を含む）を読み込んで表示するだけです。
APIキーは不要で、公開デプロイ・exe配布のどちらでも安全に配布できます。
"""

from __future__ import annotations

import math
import sys
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import plotly.express as px
import pydeck as pdk
import streamlit as st

# PyInstallerでexe化した場合、実行ファイルと同じ場所を基準にする。
# 通常のstreamlit run実行時はこのスクリプトの場所を基準にする。
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

# ファイル探索対象: input/ を最優先、次にアプリと同階層、最後にdata/（開発時のデフォルト運用）
SEARCH_DIRS = [APP_DIR / "input", APP_DIR, APP_DIR / "data"]


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


def render_map(df: pd.DataFrame, origin_lat: float, origin_lng: float, radius_m: int, dot_size: int) -> None:
    map_df = df.dropna(subset=["緯度", "経度"]).copy()
    rank_scale = {"高": 1.15, "中": 1.0, "低": 0.9}
    map_df["radius"] = map_df["競合ランク"].map(rank_scale).fillna(1.0) * float(dot_size)

    color_map = {"高": [220, 20, 60, 180], "中": [255, 165, 0, 160], "低": [30, 144, 255, 140]}
    map_df["color"] = map_df["競合ランク"].map(color_map)
    map_df["color"] = map_df["color"].apply(
        lambda v: v if isinstance(v, list) else [128, 128, 128, 120]
    )

    stores_layer = pdk.Layer(
        "ScatterplotLayer",
        data=map_df,
        get_position="[経度, 緯度]",
        get_radius="radius",
        get_fill_color="color",
        pickable=True,
        stroked=False,
    )

    origin_layer = pdk.Layer(
        "ScatterplotLayer",
        data=[{"緯度": origin_lat, "経度": origin_lng}],
        get_position="[経度, 緯度]",
        get_radius=max(20, int(dot_size * 1.8)),
        get_fill_color=[0, 0, 0, 220],
        pickable=True,
    )

    circle_data = []
    for r in [300, 500, 1000]:
        circle_data.append({"path": build_circle_path(origin_lat, origin_lng, r), "radius": r})

    circles_layer = pdk.Layer(
        "PathLayer",
        data=circle_data,
        get_path="path",
        get_width=2,
        get_color=[70, 70, 70, 120],
        width_min_pixels=1,
    )

    view = pdk.ViewState(latitude=origin_lat, longitude=origin_lng, zoom=15 if radius_m <= 500 else 14, pitch=0)

    st.pydeck_chart(pdk.Deck(layers=[circles_layer, stores_layer, origin_layer], initial_view_state=view))


st.set_page_config(page_title="競合分析ダッシュボード", page_icon="📍", layout="wide")
st.title("競合分析ダッシュボード")
st.caption("距離・Google評価・口コミ件数を軸に競争圧を可視化")

with st.sidebar:
    st.header("設定")
    export_files = list_export_files()
    if export_files:
        selected_file = st.selectbox("検出したファイル", export_files, index=0)
    else:
        selected_file = ""
        st.warning(
            "CSV/Excelファイルが見つかりません。\n\n"
            f"次のいずれかに調査結果ファイルを置いてください:\n"
            + "\n".join(f"- {d}" for d in SEARCH_DIRS)
        )

    manual_path = st.text_input("または直接パスを指定（任意）", value="")

    w_dist = st.slider("重み: 距離", min_value=0.0, max_value=1.0, value=0.40, step=0.05)
    w_rate = st.slider("重み: Google評価", min_value=0.0, max_value=1.0, value=0.35, step=0.05)
    w_reviews = st.slider("重み: 口コミ件数", min_value=0.0, max_value=1.0, value=0.25, step=0.05)

    st.markdown("---")
    st.write("推奨: 距離40% / 評価35% / 口コミ25%")


@st.cache_data(show_spinner=False)
def read_export_cached(path_str: str) -> pd.DataFrame:
    if path_str.endswith(".csv"):
        return pd.read_csv(path_str, encoding="utf-8-sig")
    return pd.read_excel(path_str)


effective_path = manual_path.strip() or selected_file
target_file = resolve_path(effective_path) if effective_path else Path("")

if not effective_path or not target_file.exists():
    st.error("分析対象のファイルが見つかりません。サイドバーでファイルを選択・指定してください。")
    st.stop()

try:
    df_raw = read_export_cached(str(target_file))
except Exception as exc:  # pragma: no cover
    st.error(f"ファイル読み込みに失敗しました: {exc}")
    st.stop()

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

with st.sidebar:
    st.caption(f"拠点: {origin_address or origin_label}")

scored = build_competitor_score(df, distance_col, w_dist, w_rate, w_reviews)
scored["競合ランク"] = scored["競合スコア"].apply(score_band)
scored[distance_col] = pd.to_numeric(scored[distance_col], errors="coerce")

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

if filtered.empty:
    st.warning("条件に一致する店舗がありません。フィルタ条件を緩めてください。")
    st.stop()

if origin_latlng is None:
    origin_lat = float(pd.to_numeric(filtered["緯度"], errors="coerce").mean()) if "緯度" in filtered.columns else 34.726893
    origin_lng = float(pd.to_numeric(filtered["経度"], errors="coerce").mean()) if "経度" in filtered.columns else 135.501036
else:
    origin_lat, origin_lng = origin_latlng

# Summary
c1, c2, c3, c4 = st.columns(4)
c1.metric("対象店舗数", f"{len(filtered):,}")
c2.metric("平均Google評価", f"{filtered['Google評価'].mean(skipna=True):.2f}")
c3.metric("平均口コミ件数", f"{filtered['口コミ件数'].mean(skipna=True):.1f}")
c4.metric("高競合比率", f"{(filtered['競合ランク'] == '高').mean() * 100:.1f}%")
st.caption("凡例: 赤=高競合 / 橙=中競合 / 青=低競合。地図はスコア上位から表示。")

tab1, tab2, tab3, tab4 = st.tabs(["地図", "散布図", "ランキング", "KPI比較"])

with tab1:
    st.subheader("拠点中心マッピング")
    if "緯度" in filtered.columns and "経度" in filtered.columns:
        map_view_df = filtered.sort_values("競合スコア", ascending=False).head(max_points)
        render_map(map_view_df, origin_lat, origin_lng, radius_filter, dot_size)
        st.caption(f"地図表示: {len(map_view_df)}件 / フィルタ後 {len(filtered)}件")
    else:
        st.info("緯度/経度列がないため地図を描画できません。")

with tab2:
    st.subheader("距離 × Google評価 × 口コミ件数")
    scatter_df = filtered.copy()
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
    fig.update_layout(height=520)
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.subheader("競合ランキング")
    top_n = st.slider("表示件数", min_value=10, max_value=min(100, len(filtered)), value=min(30, len(filtered)), step=5)
    rank_cols = [
        "店舗名",
        "ジャンル",
        distance_col,
        "Google評価",
        "口コミ件数",
        "競合スコア",
        "競合ランク",
        "価格帯",
        "深夜営業",
        "住所",
    ]
    rank_df = filtered.sort_values("競合スコア", ascending=False)[rank_cols].head(top_n)
    st.dataframe(rank_df, use_container_width=True)

with tab4:
    st.subheader("半径別KPI比較")
    records = []
    for r in [300, 500, 1000]:
        temp = scored[scored[distance_col].fillna(999999) <= r]
        if len(temp) == 0:
            continue
        records.append(
            {
                "半径(m)": r,
                "店舗数": len(temp),
                "平均Google評価": temp["Google評価"].mean(skipna=True),
                "平均口コミ件数": temp["口コミ件数"].mean(skipna=True),
                "高競合比率(%)": (temp["競合ランク"] == "高").mean() * 100,
            }
        )

    kpi_df = pd.DataFrame(records)
    st.dataframe(kpi_df, use_container_width=True)

    if not kpi_df.empty:
        fig_kpi = px.line(
            kpi_df,
            x="半径(m)",
            y=["平均Google評価", "平均口コミ件数"],
            markers=True,
        )
        fig_kpi.update_layout(height=420)
        st.plotly_chart(fig_kpi, use_container_width=True)
