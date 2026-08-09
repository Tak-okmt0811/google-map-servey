#!/usr/bin/env python3
"""Streamlit app for competitor analysis from Google Places export Excel."""

from __future__ import annotations

import math
import os
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import pydeck as pdk
import requests
import streamlit as st

GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"
BASE_DIR = Path(__file__).resolve().parent
GEOCODE_CACHE_PATH = BASE_DIR / ".geocode_cache.csv"


def load_env_file(env_path: str = ".env") -> None:
    path = Path(env_path)
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def geocode_address(address: str, api_key: str) -> Optional[Tuple[float, float]]:
    params = {"address": address, "language": "ja", "key": api_key}
    try:
        res = requests.get(GEOCODE_URL, params=params, timeout=20)
        res.raise_for_status()
        data = res.json()
    except requests.RequestException:
        return None

    if data.get("status") != "OK" or not data.get("results"):
        return None

    loc = data["results"][0]["geometry"]["location"]
    return float(loc["lat"]), float(loc["lng"])


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


def load_cache() -> pd.DataFrame:
    if not GEOCODE_CACHE_PATH.exists():
        return pd.DataFrame(columns=["住所", "緯度", "経度"])
    return pd.read_csv(GEOCODE_CACHE_PATH)


def save_cache(cache_df: pd.DataFrame) -> None:
    cache_df.drop_duplicates(subset=["住所"], keep="last").to_csv(GEOCODE_CACHE_PATH, index=False)


def geocode_dataframe(df: pd.DataFrame, address_col: str, api_key: str) -> pd.DataFrame:
    out = df.copy()
    if "緯度" not in out.columns:
        out["緯度"] = pd.NA
    if "経度" not in out.columns:
        out["経度"] = pd.NA

    cache = load_cache()
    cache_map: Dict[str, Tuple[float, float]] = {}
    for _, row in cache.iterrows():
        addr = str(row.get("住所", "")).strip()
        lat = row.get("緯度")
        lng = row.get("経度")
        if addr and pd.notna(lat) and pd.notna(lng):
            cache_map[addr] = (float(lat), float(lng))

    new_rows: List[Dict[str, object]] = []

    missing_mask = out["緯度"].isna() | out["経度"].isna()
    out.loc[missing_mask, address_col] = out.loc[missing_mask, address_col].astype(str)

    # Fill from local cache first.
    for idx, row in out.loc[missing_mask].iterrows():
        address = str(row.get(address_col, "")).strip()
        if address in cache_map:
            lat, lng = cache_map[address]
            out.at[idx, "緯度"] = lat
            out.at[idx, "経度"] = lng

    still_missing_mask = out["緯度"].isna() | out["経度"].isna()
    addresses = (
        out.loc[still_missing_mask, address_col]
        .dropna()
        .astype(str)
        .str.strip()
    )
    unique_missing = [a for a in addresses.unique().tolist() if a and a not in cache_map]

    # Avoid long waits: geocode at most 60 new addresses per run.
    request_targets = unique_missing[:60]
    if request_targets:
        progress = st.progress(0, text="住所を座標変換中...")
    else:
        progress = None

    for i, address in enumerate(request_targets, start=1):
        result = geocode_address(address, api_key)
        if result is None:
            continue
        lat, lng = result
        cache_map[address] = (lat, lng)
        new_rows.append({"住所": address, "緯度": lat, "経度": lng})
        if progress is not None:
            progress.progress(i / len(request_targets), text=f"住所を座標変換中... {i}/{len(request_targets)}")

    if progress is not None:
        progress.empty()

    if len(unique_missing) > 60:
        st.info(f"未変換住所が多いため、今回は60件まで変換しました（残り {len(unique_missing) - 60} 件）。")

    # Apply newly cached coordinates.
    for idx, row in out.loc[out["緯度"].isna() | out["経度"].isna()].iterrows():
        address = str(row.get(address_col, "")).strip()
        if address in cache_map:
            lat, lng = cache_map[address]
            out.at[idx, "緯度"] = lat
            out.at[idx, "経度"] = lng

    if new_rows:
        cache = pd.concat([cache, pd.DataFrame(new_rows)], ignore_index=True)
        save_cache(cache)

    return out


BUNDLED_CSV = BASE_DIR / "data" / "places.csv"


def latest_export_file() -> str:
    # bundled CSV takes priority over generated Excel files
    if BUNDLED_CSV.exists():
        return str(BUNDLED_CSV)
    files = sorted(BASE_DIR.glob("places_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return str(files[0]) if files else ""


def list_export_files() -> List[str]:
    if BUNDLED_CSV.exists():
        return [str(BUNDLED_CSV)]
    files = sorted(BASE_DIR.glob("places_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return [str(p) for p in files]


def resolve_excel_path(path_text: str) -> Path:
    p = Path(path_text).expanduser()
    if p.is_absolute():
        return p
    return (BASE_DIR / p).resolve()


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

load_env_file()

with st.sidebar:
    st.header("設定")
    default_path = latest_export_file()
    export_files = list_export_files()
    if export_files:
        selected_file = st.selectbox("検出したExcel", export_files, index=0)
        use_detected = st.checkbox("検出ファイルを使う", value=True)
    else:
        selected_file = ""
        use_detected = False

    excel_path = st.text_input("分析対象Excel", value=default_path)

    w_dist = st.slider("重み: 距離", min_value=0.0, max_value=1.0, value=0.40, step=0.05)
    w_rate = st.slider("重み: Google評価", min_value=0.0, max_value=1.0, value=0.35, step=0.05)
    w_reviews = st.slider("重み: 口コミ件数", min_value=0.0, max_value=1.0, value=0.25, step=0.05)

    st.markdown("---")
    st.write("推奨: 距離40% / 評価35% / 口コミ25%")


@st.cache_data(show_spinner=False)
def read_excel_cached(path_str: str) -> pd.DataFrame:
    if path_str.endswith(".csv"):
        return pd.read_csv(path_str, encoding="utf-8-sig")
    return pd.read_excel(path_str)

effective_excel_path = selected_file if (use_detected and selected_file) else excel_path
excel_file = resolve_excel_path(effective_excel_path) if effective_excel_path else Path("")

if not effective_excel_path or not excel_file.exists():
    st.error("分析対象のExcelファイルが見つかりません。")
    if export_files:
        st.info(f"検出済み候補: {export_files[0]}")
    else:
        st.info(f"{BASE_DIR} に places_*.xlsx が見つかりません。")
    st.stop()

try:
    df_raw = read_excel_cached(str(excel_file))
except Exception as exc:  # pragma: no cover
    st.error(f"Excel読み込みに失敗しました: {exc}")
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
default_origin_addr = f"大阪府大阪市淀川区{origin_label}"
with st.sidebar:
    origin_address = st.text_input("拠点住所（地図中心）", value=default_origin_addr)

api_key = os.getenv("GOOGLE_MAPS_API_KEY", "").strip()
origin_latlng = geocode_address(origin_address, api_key) if api_key else None

lat_col, lng_col = find_lat_lng_cols(df_raw)
df = df_raw.copy()

if lat_col and lng_col:
    df["緯度"] = pd.to_numeric(df[lat_col], errors="coerce")
    df["経度"] = pd.to_numeric(df[lng_col], errors="coerce")
    # lat/lng already embedded: derive origin from data mean if API key absent
    if origin_latlng is None:
        mean_lat = df["緯度"].mean(skipna=True)
        mean_lng = df["経度"].mean(skipna=True)
        if pd.notna(mean_lat) and pd.notna(mean_lng):
            origin_latlng = (float(mean_lat), float(mean_lng))
else:
    with st.sidebar:
        use_geocode = st.checkbox("住所から座標変換して地図表示", value=True)
        rerun_geocode = st.button("座標変換を再実行")

    if use_geocode:
        if api_key:
            geo_state_key = f"geo::{excel_file}"
            should_run = rerun_geocode or geo_state_key not in st.session_state
            if should_run:
                with st.spinner("住所を座標変換しています..."):
                    st.session_state[geo_state_key] = geocode_dataframe(df, "住所", api_key)
            df = st.session_state.get(geo_state_key, df)
        else:
            st.warning("GOOGLE_MAPS_API_KEY がないため、店舗座標の自動変換をスキップしました。")

if origin_latlng is None:
    st.warning("拠点住所を座標変換できませんでした。地図中心は店舗平均座標を使います。")

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
