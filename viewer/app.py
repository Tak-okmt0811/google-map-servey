#!/usr/bin/env python3
"""Streamlit viewer for competitor analysis exports.

このアプリはGoogle APIを一切呼び出しません。collector/export.py が事前に生成した
Excel/CSV、またはブラウザからアップロードされたファイルを読み込んで表示するだけです。
APIキーは不要で、公開デプロイしても安全です。
（出店候補地の住所検索だけは、キー不要の国土地理院 住所検索APIを使います）
"""

from __future__ import annotations

import json
import math
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from urllib.error import URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pydeck as pdk
import streamlit as st

APP_DIR = Path(__file__).resolve().parent

# ファイルがアップロードされていないときに表示する架空のサンプルデータ（大阪市中心部・本町付近）
SAMPLE_PATH = APP_DIR / "data" / "sample_places.xlsx"

STORE_SHEET = "店舗"
KEYWORD_SHEET = "キーワード順位"
GRID_SHEET = "地点別順位"
ATTR_PREFIX = "属性_"
NO_OWN_STORE = "（選択しない）"

# 差別化分析のしきい値
COMMON_RATE = 0.6  # これ以上の競合が持つ属性 = 持っていないと不利
RARE_RATE = 0.2  # これ以下の競合しか持たない属性 = 打ち出せば差別化になる
MIN_REGISTERED = 5  # 登録済みの競合がこれ未満の属性は判定しない
MIN_SAME_CATEGORY = 10  # 同カテゴリの競合がこれ以上いれば、同カテゴリとの比較を初期値にする
# 「ある＝強み」とは言えない属性は判定から外す
EXCLUDED_FROM_JUDGEMENT = {"現金のみ"}

BASIC_PROFILE_ITEMS = {
    "電話番号": "電話番号",
    "公式サイト/SNS": "Webサイト",
    "営業時間": "営業時間",
    "価格帯": "価格帯",
}

# 順位マップ: 地点番号 -> 方角（collector/export.py の _sub_centers の並び順）
GRID_POINT_NAMES = ["基点", "北", "南", "東", "西", "北東", "北西", "南東", "南西"]
# (この順位以内, 表示名, 色)。どれにも入らない = 圏外（検索結果の上位60件に出てこない）
RANK_BANDS = [
    (3, "1〜3位", [46, 125, 50, 230]),
    (10, "4〜10位", [249, 168, 37, 230]),
    (20, "11〜20位", [239, 108, 0, 230]),
    (60, "21〜60位", [198, 40, 40, 230]),
]
OUT_OF_RANK = ("圏外", [120, 120, 120, 200])

# 出店候補地: 住所 -> 緯度経度は国土地理院の住所検索API（キー不要）を使う。Google APIは呼ばない
GSI_ADDRESS_SEARCH_URL = "https://msearch.gsi.go.jp/address-search/AddressSearch"
MAX_CANDIDATES = 5
# 競合度（地図の色）のルール
STRONG_RATING = 4.0
STRONG_OTHER_QUANTILE = 0.75  # 別カテゴリでも、口コミ件数が全店舗の上位25%なら「中」
COMPETITION_COLORS = {"高": [220, 20, 60, 190], "中": [255, 165, 0, 170], "低": [30, 144, 255, 140]}
MAP_DOT_SIZE = 12  # 地図の点の大きさの初期値
# キーワード順位: 上位10店のうち、比較するカテゴリの店がこの数以上いれば「関係のあるキーワード」とみなす
KEYWORD_RELATED_MIN = 3

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


def category_label(categories: List[str]) -> str:
    if not categories:
        return "全店舗"
    if len(categories) <= 3:
        return "・".join(categories)
    return f"{'・'.join(categories[:2])}など{len(categories)}カテゴリ"


def category_mask(df: pd.DataFrame, categories: List[str]) -> pd.Series:
    """比較するカテゴリに含まれる店舗。カテゴリ未指定なら全店舗。"""
    if not categories:
        return pd.Series(True, index=df.index)
    if "カテゴリ" not in df.columns:
        return pd.Series(False, index=df.index)
    return df["カテゴリ"].fillna("").astype(str).str.strip().isin(categories)


def assign_competition_level(df: pd.DataFrame, categories: List[str]) -> pd.DataFrame:
    """Googleの3つの評価基準に沿って、各店舗の競合度（高/中/低）とその理由を付ける。

    関連性 = 比較するカテゴリ（複数可・まとめて1つのグループとして扱う）に含まれるか、
    知名度 = 評価と口コミ件数。距離は分析半径での絞り込みで扱う。
    """
    out = df.copy()
    rating = pd.to_numeric(out["Google評価"], errors="coerce")
    reviews_raw = pd.to_numeric(out["口コミ件数"], errors="coerce")
    reviews = reviews_raw.fillna(0)
    high_rating = rating >= STRONG_RATING
    rating_text = f"評価{STRONG_RATING}以上"

    if not categories:
        # 中央値・上位25%は口コミのある店だけで計算する（口コミ0件の店で基準が下がらないように）
        median = int(reviews_raw.median())
        many_reviews = reviews >= median
        out["競合度"] = np.select([high_rating & many_reviews, high_rating | many_reviews], ["高", "中"], "低")
        out["競合度の理由"] = np.select(
            [high_rating & many_reviews, high_rating | many_reviews],
            [f"{rating_text}・口コミ{median}件以上（全店舗の中央値）", f"{rating_text}か口コミ{median}件以上の一方のみ"],
            "評価・口コミとも基準未満",
        )
        return out

    same = category_mask(out, categories)
    same_median = int(reviews_raw[same].median()) if reviews_raw[same].notna().any() else int(reviews_raw.median())
    top_quarter = int(reviews_raw.quantile(STRONG_OTHER_QUANTILE))
    strong = same & high_rating & (reviews >= same_median)
    strong_other = ~same & high_rating & (reviews >= top_quarter)
    out["競合度"] = np.select([strong, same | strong_other], ["高", "中"], "低")
    out["競合度の理由"] = np.select(
        [strong, same, strong_other],
        [
            f"比較するカテゴリ・{rating_text}・口コミ{same_median}件以上（比較するカテゴリの中央値）",
            "比較するカテゴリ（評価か口コミが基準未満）",
            f"別カテゴリだが{rating_text}・口コミ{top_quarter}件以上（全店舗の上位25%）",
        ],
        "別カテゴリ",
    )
    return out


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
def read_sheets_cached(path_str: str, mtime: float) -> Dict[str, pd.DataFrame]:
    """mtime はキャッシュのキー用。ファイルを作り直したときに古い内容が表示され続けないようにする。"""
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
        if attr in EXCLUDED_FROM_JUDGEMENT:
            continue
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
    center_lat: float,
    center_lng: float,
    radius_m: int,
    origin_lat: float,
    origin_lng: float,
    own_row: Optional[pd.Series],
    dot_size: int = MAP_DOT_SIZE,
) -> None:
    map_df = df.dropna(subset=["緯度", "経度"]).copy()
    size_scale = {"高": 1.3, "中": 1.0, "低": 0.85}
    map_df["radius"] = map_df["競合度"].map(size_scale).fillna(1.0) * float(dot_size)
    map_df["color"] = map_df["競合度"].map(COMPETITION_COLORS)
    map_df["color"] = map_df["color"].apply(lambda v: v if isinstance(v, list) else [128, 128, 128, 120])
    map_df["tooltip"] = map_df["店舗名"].astype(str) + "（競合度: " + map_df["競合度"].astype(str) + "）"
    # 赤が青に隠れないよう、競合度の低い店から描く
    map_df = map_df.sort_values("競合度", key=lambda s: s.map({"低": 0, "中": 1, "高": 2}))

    layers = [
        pdk.Layer(
            "PathLayer",
            data=[{"path": build_circle_path(center_lat, center_lng, radius_m)}],
            get_path="path",
            get_width=2,
            get_color=[70, 70, 70, 140],
            width_min_pixels=1,
        ),
        pdk.Layer(
            "ScatterplotLayer",
            data=map_df[["tooltip", "緯度", "経度", "radius", "color"]],
            get_position="[経度, 緯度]",
            get_radius="radius",
            get_fill_color="color",
            pickable=True,
            stroked=False,
        ),
        pdk.Layer(
            "ScatterplotLayer",
            data=[{"tooltip": "拠点（データ収集の中心）", "緯度": origin_lat, "経度": origin_lng}],
            get_position="[経度, 緯度]",
            get_radius=int(dot_size * 1.8),
            get_fill_color=[0, 0, 0, 220],
            pickable=True,
        ),
    ]
    if own_row is not None and pd.notna(own_row.get("緯度")) and pd.notna(own_row.get("経度")):
        layers.append(
            pdk.Layer(
                "ScatterplotLayer",
                data=[{"tooltip": f"自店: {own_row['店舗名']}", "緯度": float(own_row["緯度"]), "経度": float(own_row["経度"])}],
                get_position="[経度, 緯度]",
                get_radius=int(dot_size * 2.2),
                get_fill_color=[0, 200, 83, 240],
                get_line_color=[255, 255, 255, 255],
                line_width_min_pixels=2,
                stroked=True,
                pickable=True,
            )
        )

    view = pdk.ViewState(latitude=center_lat, longitude=center_lng, zoom=15 if radius_m <= 500 else 14, pitch=0)
    st.pydeck_chart(pdk.Deck(layers=layers, initial_view_state=view, tooltip={"text": "{tooltip}"}))


def render_differentiation_tab(competitors: pd.DataFrame, own_row: Optional[pd.Series], categories: List[str]) -> None:
    st.subheader("差別化分析")
    if own_row is None:
        st.info("サイドバーで「自店舗」を選ぶと、自店と周辺競合の差（要対応・アピールポイント）を判定します。")

    has_attrs = bool(attribute_cols(competitors))
    if not has_attrs:
        st.warning(
            "このファイルには属性データ（属性_○○列）がありません。最新のcollectorで出力し直すと、"
            "属性の保有率・プロフィール充実度が表示されます。"
        )

    # 比較対象: 業態が違う店（例: バーに対するラーメン店）まで含めると判定がぶれるため、
    # 同カテゴリの競合が十分いれば同カテゴリとの比較を初期値にする
    base = competitors
    if categories and "カテゴリ" in competitors.columns:
        same = competitors[category_mask(competitors, categories)]
        scope_options = [f"比較するカテゴリ（{category_label(categories)}・{len(same)}店）", f"周辺の全競合（{len(competitors)}店）"]
        scope = st.radio(
            "比較対象",
            scope_options,
            index=0 if len(same) >= MIN_SAME_CATEGORY else 1,
            horizontal=True,
            help="属性の保有率・判定・プロフィール充実度をどの競合と比べるか。カテゴリ分布は常に全競合で表示します。",
        )
        if scope == scope_options[0]:
            base = same

    attr_table = build_attribute_table(base, own_row) if (has_attrs or "深夜営業" in base.columns) else None

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
        comp_scores = profile_scores(base)["プロフィール充実度"]
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
            comp_attr = attribute_frame(base).drop(columns=["深夜営業"], errors="ignore")
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


def distance_from(lat: float, lng: float, lats: pd.Series, lngs: pd.Series) -> pd.Series:
    """1地点から各店舗までの距離(m)。collector と同じハバーサイン式。"""
    lat1, lng1 = np.radians(lat), np.radians(lng)
    lat2, lng2 = np.radians(pd.to_numeric(lats, errors="coerce")), np.radians(pd.to_numeric(lngs, errors="coerce"))
    a = np.sin((lat2 - lat1) / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin((lng2 - lng1) / 2) ** 2
    return 6371000 * 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))


def collected_radius(df: pd.DataFrame, distance_col: str) -> float:
    """データを収集した半径。列が無い古い形式のファイルは、最も遠い店舗までの距離で代用する。"""
    if "収集半径(m)" in df.columns:
        r = pd.to_numeric(df["収集半径(m)"], errors="coerce").dropna()
        if not r.empty:
            return float(r.iloc[0])
    return float(pd.to_numeric(df[distance_col], errors="coerce").max())


def rank_band(rank: Optional[float]) -> Tuple[str, List[int]]:
    if rank is not None and pd.notna(rank):
        for limit, label, color in RANK_BANDS:
            if rank <= limit:
                return label, color
    return OUT_OF_RANK


def median_or_none(series: pd.Series) -> Optional[float]:
    s = pd.to_numeric(series, errors="coerce").dropna()
    return float(s.median()) if not s.empty else None


def render_keyword_tab(
    keyword_sheet: Optional[pd.DataFrame], stores: pd.DataFrame, own_row: Optional[pd.Series], categories: List[str]
) -> None:
    st.subheader("キーワード順位")
    st.caption(
        "キーワードで検索したとき（拠点で検索）に、各店舗が何位に出たかを表示します。"
        "上位の店の口コミ件数・評価と比べると、自店が上位に入るための目安が分かります。"
    )
    if keyword_sheet is None or keyword_sheet.empty or "place_id" not in stores.columns:
        st.info(
            "このファイルには「キーワード順位」シートがありません。最新のcollectorで出力し直すと表示されます。"
        )
        return

    info_cols = [c for c in ["カテゴリ", "Google評価", "口コミ件数"] if c in stores.columns]
    info = stores.assign(place_id=stores["place_id"].astype(str)).drop_duplicates("place_id").set_index("place_id")[info_cols]
    kw = keyword_sheet.assign(place_id=keyword_sheet["place_id"].astype(str))
    kw["順位"] = pd.to_numeric(kw["順位"], errors="coerce")
    # 閉業などで店舗シートから除いた店は順位表にも出さない
    kw = kw.join(info, on="place_id", how="inner")
    kw["比較カテゴリ"] = category_mask(kw, categories).map({True: "○", False: ""}) if categories else ""
    ranked = kw.dropna(subset=["順位"]).sort_values("順位")
    all_keywords = sorted(kw["キーワード"].dropna().unique().tolist())
    if not all_keywords:
        st.info("キーワード順位のデータがありません。")
        return

    own_id = str(own_row.get("place_id", "")) if own_row is not None else ""

    # 比較するカテゴリと関係のあるキーワード = 上位10店に比較カテゴリの店が一定数いる（または自店が表示される）
    top10_hits = {
        k: int((g.head(10)["比較カテゴリ"] == "○").sum()) for k, g in ranked.groupby("キーワード")
    }
    own_keywords = set(ranked.loc[ranked["place_id"] == own_id, "キーワード"]) if own_id else set()
    related = [k for k in all_keywords if top10_hits.get(k, 0) >= KEYWORD_RELATED_MIN or k in own_keywords]
    keywords = all_keywords
    if categories:
        label = category_label(categories)
        show_all = st.toggle(
            f"「{label}」と関係の薄いキーワードも表示",
            value=False,
            key="kw_show_all",
            help=(
                f"初期状態では、上位10店のうち{KEYWORD_RELATED_MIN}店以上が「{label}」のキーワード"
                "（と自店が表示されるキーワード）だけを表示します。"
            ),
        )
        if not show_all:
            if related:
                keywords = related
            else:
                st.info(f"「{label}」の店が上位に多いキーワードがないため、すべてのキーワードを表示しています。")
    own_reviews = pd.to_numeric(pd.Series([own_row.get("口コミ件数")]), errors="coerce").iloc[0] if own_row is not None else None

    # 1. 自店のキーワード別まとめ
    summary_rows = []
    for k in keywords:
        g = ranked[ranked["キーワード"] == k]
        own_rank = g.loc[g["place_id"] == own_id, "順位"].min() if own_id else None
        top3 = median_or_none(g[g["順位"] <= 3]["口コミ件数"])
        row = {
            "キーワード": k,
            "自店の順位": (f"{int(own_rank)}位" if own_rank is not None and pd.notna(own_rank) else "圏外") if own_id else "自店未選択",
            "1位の店": g["店舗名"].iloc[0] if not g.empty else "",
            **({f"上位10店のうち{category_label(categories)}": f"{top10_hits.get(k, 0)}店"} if categories else {}),
            "上位3店の口コミ件数（中央値）": top3,
            "上位10店の口コミ件数（中央値）": median_or_none(g[g["順位"] <= 10]["口コミ件数"]),
            "上位10店の評価（中央値）": median_or_none(g[g["順位"] <= 10]["Google評価"]),
        }
        if own_id:
            # 圏外のキーワードは口コミ件数より先に関連性（カテゴリ・説明文）の問題なので、件数の差は出さない
            is_ranked = own_rank is not None and pd.notna(own_rank)
            gap = None if not is_ranked or top3 is None or own_reviews is None or pd.isna(own_reviews) else max(0, int(round(top3 - own_reviews)))
            row["上位3店の中央値まで（口コミ）"] = "—" if gap is None else (f"あと{gap}件" if gap else "到達")
        summary_rows.append(row)
    summary = pd.DataFrame(summary_rows)
    if own_id:
        # 自店が表示されるキーワードを上にまとめる
        summary = summary.sort_values("自店の順位", key=lambda s: s.eq("圏外"), kind="stable")

    st.markdown("#### キーワード別のまとめ" + ("（自店）" if own_id else ""))
    st.dataframe(summary, width="stretch", hide_index=True)
    if own_id:
        st.caption(
            f"自店の口コミ件数は{int(own_reviews) if own_reviews is not None and pd.notna(own_reviews) else '-'}件。"
            "「上位3店の中央値まで」は、上位3店の口コミ件数の中央値に届くまでの件数の目安です。"
            "圏外（上位60件に出てこない）のキーワードは、口コミ件数より先に、カテゴリ・説明文・口コミの内容が"
            "その言葉と結びついていない可能性があります（自店の業態と関係のないキーワードは気にしなくて構いません）。"
        )

    # 2. キーワードを選んで詳しく見る（初期値: 自店の順位が最も高いキーワード）
    st.markdown("#### キーワードを選んで詳しく見る")
    own_ranked = ranked[ranked["place_id"] == own_id].sort_values("順位") if own_id else ranked.iloc[0:0]
    default_kw = own_ranked["キーワード"].iloc[0] if not own_ranked.empty else keywords[0]
    keyword = st.selectbox("キーワード", keywords, index=keywords.index(default_kw), key=f"kw_keyword_{own_id}")
    g = ranked[ranked["キーワード"] == keyword].copy()
    g["順位帯"] = g["順位"].map(lambda r: rank_band(r)[0])
    g["自店"] = g["place_id"].eq(own_id).map({True: "○", False: ""})
    g["店の種類"] = g["比較カテゴリ"].map({"○": category_label(categories)}).fillna("その他") if categories else "全店舗"

    plot = g.copy()
    plot["口コミ件数_plot"] = pd.to_numeric(plot["口コミ件数"], errors="coerce").fillna(0).clip(lower=0) + 1
    band_colors = {label: f"rgb({c[0]},{c[1]},{c[2]})" for _, label, c in RANK_BANDS}
    fig = px.scatter(
        plot,
        x="口コミ件数_plot",
        y="Google評価",
        color="順位帯",
        category_orders={"順位帯": [label for _, label, _ in RANK_BANDS]},
        color_discrete_map=band_colors,
        symbol="店の種類" if categories else None,
        symbol_map={category_label(categories): "circle", "その他": "x"} if categories else None,
        hover_name="店舗名",
        hover_data={"順位": True, "口コミ件数": True, "口コミ件数_plot": False, "カテゴリ": True},
        log_x=True,
        labels={"口コミ件数_plot": "口コミ件数（+1・対数目盛）"},
    )
    own_point = plot[plot["place_id"] == own_id]
    if not own_point.empty:
        fig.add_trace(
            go.Scatter(
                x=own_point["口コミ件数_plot"],
                y=own_point["Google評価"],
                mode="markers+text",
                marker={"symbol": "star", "size": 22, "color": "#00c853", "line": {"width": 1, "color": "white"}},
                text=["自店"],
                textposition="top center",
                name="自店",
                hovertext=own_point["店舗名"],
            )
        )
    fig.update_layout(height=460, legend_title_text="順位")
    st.plotly_chart(fig, width="stretch")
    st.caption(
        "見方: 緑（1〜3位）の店が右側（口コミが多い）に集まっていれば、このキーワードは口コミ件数が効いています。"
        "口コミの少ない店が上位にいる場合は、店名・カテゴリとキーワードの一致や距離が効いています。"
    )

    if categories:
        st.caption(f"●=比較するカテゴリ（{category_label(categories)}）の店、×=それ以外の店。")

    cols = [c for c in ["順位", "自店", "比較カテゴリ" if categories else "", "店舗名", "カテゴリ", "Google評価", "口コミ件数", "距離(m)"] if c in g.columns]
    st.dataframe(g[cols].head(30), width="stretch", hide_index=True)
    st.caption(
        "順位はPlaces APIで拠点を中心に検索した結果の並び順にもとづく目安で、上位60件まで記録しています。"
        "距離(m)は拠点からの距離です。場所ごとの順位の違いは「順位マップ」タブで確認できます。"
    )


def render_rank_map_tab(
    grid: Optional[pd.DataFrame],
    stores: pd.DataFrame,
    own_row: Optional[pd.Series],
    origin_lat: float,
    origin_lng: float,
    area_radius: float,
) -> None:
    st.subheader("順位マップ")
    st.caption(
        "同じキーワードでも、検索する場所によって表示順位は変わります。基点と周辺8地点のそれぞれで検索したときに、"
        "選んだ店舗が何位に出たかを地図に表示します。"
    )
    if grid is None or grid.empty:
        st.info(
            "このファイルには「地点別順位」シートがありません。最新のcollectorで出力し直すと、"
            "地点ごとの順位マップが表示されます。"
        )
        return

    grid = grid.copy()
    grid["順位"] = pd.to_numeric(grid["順位"], errors="coerce")
    points = (
        grid[["地点番号", "地点緯度", "地点経度"]]
        .drop_duplicates("地点番号")
        .sort_values("地点番号")
        .reset_index(drop=True)
    )
    points["地点"] = points["地点番号"].map(lambda i: GRID_POINT_NAMES[int(i)] if int(i) < len(GRID_POINT_NAMES) else f"地点{i}")
    if len(points) <= 1:
        st.caption("収集半径が300m以下のため、検索地点は基点の1か所だけです。")

    names = grid.drop_duplicates("place_id").set_index("place_id")["店舗名"]
    best_origin = grid[grid["地点番号"] == 0].groupby("place_id")["順位"].min()
    store_ids = sorted(names.index, key=lambda pid: (best_origin.get(pid, 999), names[pid]))
    own_id = str(own_row.get("place_id", "")) if own_row is not None else ""
    if own_id in names.index:
        store_ids.remove(own_id)
        store_ids.insert(0, own_id)

    col_kw, col_store = st.columns(2)
    keyword = col_kw.selectbox("キーワード", sorted(grid["キーワード"].dropna().unique().tolist()), key="rank_keyword")
    store_id = col_store.selectbox(
        "順位を見る店舗",
        store_ids,
        format_func=lambda pid: f"自店: {names[pid]}" if pid == own_id else str(names[pid]),
        key=f"rank_store_{own_id}",  # 自店を切り替えたら、選択も新しい自店に戻す
        help="初期値は自店（サイドバーで選択した店舗）。競合店を選ぶと、その店がどの場所で強いかが分かります。",
    )

    kw_grid = grid[grid["キーワード"] == keyword]
    store_ranks = kw_grid[kw_grid["place_id"] == store_id].groupby("地点番号")["順位"].min()
    points["順位"] = points["地点番号"].map(store_ranks)
    points["順位帯"] = points["順位"].map(lambda r: rank_band(r)[0])
    points["color"] = points["順位"].map(lambda r: rank_band(r)[1])
    # TextLayerは既定で英数字しか描画できないため、圏外は「-」で表す
    points["label"] = points["順位"].map(lambda r: str(int(r)) if pd.notna(r) else "-")

    ranked = points["順位"].dropna()
    origin_rank = points.loc[points["地点番号"] == 0, "順位"]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("基点での順位", f"{int(origin_rank.iloc[0])}位" if not origin_rank.empty and pd.notna(origin_rank.iloc[0]) else "圏外")
    m2.metric("上位3位以内の地点", f"{int((ranked <= 3).sum())} / {len(points)}")
    m3.metric("上位10位以内の地点", f"{int((ranked <= 10).sum())} / {len(points)}")
    m4.metric("平均順位（表示された地点）", f"{ranked.mean():.1f}位" if not ranked.empty else "-")

    point_radius = max(40.0, area_radius * 0.12)
    layers = [
        pdk.Layer(
            "PathLayer",
            data=[{"path": build_circle_path(origin_lat, origin_lng, int(area_radius))}],
            get_path="path",
            get_width=2,
            get_color=[70, 70, 70, 120],
            width_min_pixels=1,
        ),
        pdk.Layer(
            "ScatterplotLayer",
            data=points[["地点", "地点緯度", "地点経度", "順位帯", "color"]],
            get_position="[地点経度, 地点緯度]",
            get_radius=point_radius,
            get_fill_color="color",
            pickable=True,
        ),
        pdk.Layer(
            "TextLayer",
            data=points[["地点緯度", "地点経度", "label"]],
            get_position="[地点経度, 地点緯度]",
            get_text="label",
            get_size=18,
            get_color=[255, 255, 255, 255],
            font_weight="bold",
        ),
    ]
    store = stores[stores["place_id"].astype(str) == str(store_id)] if "place_id" in stores.columns else stores.iloc[0:0]
    if not store.empty and pd.notna(store.iloc[0].get("緯度")):
        layers.append(
            pdk.Layer(
                "ScatterplotLayer",
                data=[{"地点": f"店舗: {names[store_id]}", "順位帯": "", "緯度": float(store.iloc[0]["緯度"]), "経度": float(store.iloc[0]["経度"])}],
                get_position="[経度, 緯度]",
                get_radius=max(20.0, point_radius * 0.3),
                get_fill_color=[0, 200, 83, 240],
                get_line_color=[255, 255, 255, 255],
                line_width_min_pixels=2,
                stroked=True,
                pickable=True,
            )
        )
    view = pdk.ViewState(latitude=origin_lat, longitude=origin_lng, zoom=15 if area_radius <= 500 else 14, pitch=0)
    st.pydeck_chart(pdk.Deck(layers=layers, initial_view_state=view, tooltip={"text": "{地点}: {順位帯}"}))
    st.caption(
        "凡例: 緑=1〜3位 / 黄=4〜10位 / 橙=11〜20位 / 赤=21〜60位 / 灰（-）=圏外（上位60件に出てこない） / 小さな緑の点=選んだ店舗。"
        "円はデータの収集範囲。"
    )

    # 地点ごとの内訳（その地点で最上位だった店も並べる）
    top = kw_grid.sort_values("順位").drop_duplicates("地点番号").set_index("地点番号")
    detail = points[["地点", "地点番号", "順位"]].copy()
    detail["順位"] = detail["順位"].map(lambda r: f"{int(r)}位" if pd.notna(r) else "圏外")
    detail["この地点の最上位の店"] = detail["地点番号"].map(top["店舗名"])
    detail["その店の順位"] = detail["地点番号"].map(top["順位"]).map(lambda r: f"{int(r)}位" if pd.notna(r) else "")
    st.dataframe(detail.drop(columns=["地点番号"]), width="stretch", hide_index=True)

    # 選んだ店舗のキーワード別まとめ
    st.markdown("#### キーワード別のまとめ（選んだ店舗）")
    n_points = len(points)
    rows = []
    for kw, g in grid.groupby("キーワード"):
        ranks = g[g["place_id"] == store_id].groupby("地点番号")["順位"].min()
        origin = ranks.get(0)
        rows.append(
            {
                "キーワード": kw,
                "基点での順位": f"{int(origin)}位" if origin is not None and pd.notna(origin) else "圏外",
                "上位3位以内の地点数": f"{int((ranks <= 3).sum())} / {n_points}",
                "上位10位以内の地点数": f"{int((ranks <= 10).sum())} / {n_points}",
                "平均順位（表示された地点）": round(float(ranks.mean()), 1) if not ranks.empty else None,
            }
        )
    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
    st.caption(
        "順位はPlaces APIで検索した結果の並び順にもとづく目安です。実際のGoogleマップの順位は、"
        "検索する人の正確な位置・時間・検索履歴などでも変わります。"
    )


@st.cache_data(ttl=86400, max_entries=200, show_spinner=False)
def geocode_candidate(query: str) -> Optional[Tuple[float, float, str]]:
    """住所（または「緯度,経度」）を緯度経度に変換する。住所は国土地理院の住所検索APIを使う。"""
    m = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)\s*[,，]\s*(-?\d+(?:\.\d+)?)\s*", query)
    if m:
        return float(m.group(1)), float(m.group(2)), "座標で指定"
    req = Request(f"{GSI_ADDRESS_SEARCH_URL}?{urlencode({'q': query})}", headers={"User-Agent": "competitor-viewer"})
    with urlopen(req, timeout=10) as res:
        results = json.load(res)
    if not results:
        return None
    lng, lat = results[0]["geometry"]["coordinates"]
    return float(lat), float(lng), str(results[0]["properties"].get("title", ""))


def summarize_area(pool: pd.DataFrame, lat: float, lng: float, radius: int) -> Tuple[Dict[str, object], pd.DataFrame]:
    """1地点の半径内にある競合の数・評価・口コミ件数をまとめる。"""
    dist = distance_from(lat, lng, pool["緯度"], pool["経度"])
    near = pool.assign(**{"候補地からの距離(m)": dist.round()}).loc[dist <= radius]
    near = near.sort_values("候補地からの距離(m)")
    rating = pd.to_numeric(near["Google評価"], errors="coerce")
    reviews = pd.to_numeric(near["口コミ件数"], errors="coerce")
    summary = {
        "競合数": len(near),
        "強い競合（地図の赤）": int((near["競合度"] == "高").sum()),
        "平均評価": round(float(rating.mean()), 2) if rating.notna().any() else None,
        "口コミ件数の中央値": round(float(reviews.median()), 0) if reviews.notna().any() else None,
        "最寄りの競合まで(m)": int(dist.min()) if dist.notna().any() else None,
    }
    return summary, near


def render_candidate_tab(
    pool: pd.DataFrame,
    categories: List[str],
    own_row: Optional[pd.Series],
    origin_lat: float,
    origin_lng: float,
    origin_name: str,
    area_radius: float,
) -> None:
    st.subheader("出店候補地の比較")
    st.caption(
        "まだ出店していない場所の住所を入れると、その周りにある競合の数・評価・口コミ件数を、"
        "読み込んだデータから集計して比べます。候補地は、データの収集範囲（地図の円）の中にある必要があります。"
    )
    if "緯度" not in pool.columns or "経度" not in pool.columns:
        st.info("緯度/経度列がないため、候補地の比較はできません。")
        return

    st.caption(f"数える競合: **{category_label(categories)}**（サイドバーの「比較するカテゴリ」で変更できます）")

    with st.form("candidate_form"):
        text = st.text_area(
            f"候補地の住所（1行に1か所・最大{MAX_CANDIDATES}か所）",
            placeholder="大阪府大阪市北区梅田1丁目\n34.7025, 135.4959（緯度, 経度でも指定できます）",
            key="candidate_text",
        )
        radius = st.segmented_control(
            "候補地からの半径",
            [300, 500],
            default=300,
            format_func=lambda v: f"{v}m",
            key="candidate_radius",
        )
        st.form_submit_button("比較する", type="primary")

    radius = int(radius or 300)
    queries = [q.strip() for q in text.splitlines() if q.strip()][:MAX_CANDIDATES]
    if not queries:
        st.info("候補地の住所を入力して「比較する」を押してください。")
        return

    target = pool[category_mask(pool, categories)]
    if own_row is not None:
        target = target.drop(index=own_row.name, errors="ignore")

    sites = [{"候補地": f"（参考）基点: {origin_name}", "lat": origin_lat, "lng": origin_lng, "label": "0"}]
    for i, q in enumerate(queries, start=1):
        try:
            with st.spinner(f"住所を検索中: {q}"):
                found = geocode_candidate(q)
        except (URLError, TimeoutError, ValueError, KeyError) as exc:
            st.error(f"「{q}」の住所検索に失敗しました（{exc}）。時間をおいて再度お試しください。")
            continue
        if found is None:
            st.warning(f"「{q}」に該当する住所が見つかりませんでした。表記を変えるか「緯度, 経度」で指定してください。")
            continue
        lat, lng, title = found
        sites.append({"候補地": f"{i}. {q}", "lat": lat, "lng": lng, "label": str(i), "検索結果": title})

    rows, nearby = [], {}
    for site in sites:
        from_origin = float(distance_from(origin_lat, origin_lng, pd.Series([site["lat"]]), pd.Series([site["lng"]])).iloc[0])
        if from_origin > area_radius:
            coverage = "範囲外（集計できません）"
        elif from_origin + radius > area_radius:
            coverage = "一部が範囲外（競合数は少なめに出ます）"
        else:
            coverage = "OK"
        row = {"候補地": site["候補地"], "基点からの距離(m)": int(from_origin), "データ範囲": coverage}
        if not coverage.startswith("範囲外"):
            summary, near = summarize_area(target, site["lat"], site["lng"], radius)
            row.update(summary)
            nearby[site["候補地"]] = near
        rows.append(row)

    st.markdown(f"#### 半径{radius}m以内の競合（{category_label(categories)}）")
    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
    st.caption(
        "強い競合は、地図タブの赤と同じ基準です。"
        "見方の例: 競合数が少なくても強い競合が1店あると、口コミで追いつくまで時間がかかります。"
        "競合数は多くても平均評価が低い・口コミ件数が少ない場所は、評価を積み上げれば上位を狙える余地があります。"
    )

    target_ids = set(target.index)
    map_df = pool.dropna(subset=["緯度", "経度"])[["店舗名", "緯度", "経度"]].copy()
    map_df["color"] = [[230, 81, 0, 200] if idx in target_ids else [144, 164, 174, 120] for idx in map_df.index]
    site_df = pd.DataFrame(sites)
    site_df["店舗名"] = site_df["候補地"]
    layers = [
        pdk.Layer(
            "PathLayer",
            data=[{"path": build_circle_path(origin_lat, origin_lng, int(area_radius))}]
            + [{"path": build_circle_path(s["lat"], s["lng"], radius)} for s in sites[1:]],
            get_path="path",
            get_width=2,
            get_color=[70, 70, 70, 140],
            width_min_pixels=1,
        ),
        pdk.Layer(
            "ScatterplotLayer",
            data=map_df,
            get_position="[経度, 緯度]",
            get_radius=12,
            get_fill_color="color",
            pickable=True,
        ),
        pdk.Layer(
            "ScatterplotLayer",
            data=site_df[["店舗名", "lat", "lng"]],
            get_position="[lng, lat]",
            get_radius=36,
            get_fill_color=[21, 101, 192, 230],
            pickable=True,
        ),
        pdk.Layer(
            "TextLayer",
            data=site_df[["lat", "lng", "label"]],
            get_position="[lng, lat]",
            get_text="label",
            get_size=16,
            get_color=[255, 255, 255, 255],
            font_weight="bold",
        ),
    ]
    view = pdk.ViewState(latitude=origin_lat, longitude=origin_lng, zoom=14, pitch=0)
    st.pydeck_chart(pdk.Deck(layers=layers, initial_view_state=view, tooltip={"text": "{店舗名}"}))
    st.caption("凡例: 青=候補地（0は基点） / 橙=数える競合 / 灰=その他の店舗。大きい円はデータの収集範囲、小さい円は候補地からの半径。")

    for site in sites[1:]:
        near = nearby.get(site["候補地"])
        if near is None:
            continue
        with st.expander(f"{site['候補地']} の周辺の競合（{len(near)}店）"):
            if site.get("検索結果"):
                st.caption(f"住所検索の結果: {site['検索結果']}（{site['lat']:.5f}, {site['lng']:.5f}）")
            cols = [c for c in ["店舗名", "カテゴリ", "候補地からの距離(m)", "Google評価", "口コミ件数", "最新口コミ日", "住所"] if c in near.columns]
            st.dataframe(near[cols], width="stretch", hide_index=True)

    st.caption("住所検索: 国土地理院 住所検索API。住所によっては町名・丁目の代表点になります。")


def render_store_list_tab(stores: pd.DataFrame, own_row: Optional[pd.Series], radius: int, dist_col: str) -> None:
    st.subheader("店舗一覧")
    st.caption("読み込んだデータの全店舗を、検索・絞り込みして確認できます。見出しをクリックすると並べ替え、絞り込んだ結果はCSVで保存できます。")

    data = stores.drop(columns=["自店"], errors="ignore")
    data.insert(0, "自店", data.index.map(lambda idx: "○" if own_row is not None and idx == own_row.name else ""))
    cats = data["カテゴリ"].fillna("").astype(str).str.strip() if "カテゴリ" in data.columns else pd.Series("", index=data.index)
    attr_labels = [c.removeprefix(ATTR_PREFIX) for c in attribute_cols(data)]

    with st.container(border=True):
        query = st.text_input("店名・住所・カテゴリで検索", placeholder="例: パレノール", key="list_query")
        c1, c2, c3 = st.columns(3)
        selected_cats = c1.multiselect(
            "カテゴリ", cats[cats != ""].value_counts().index.tolist(), placeholder="すべて", key="list_categories"
        )
        levels = c2.pills("競合度", ["高", "中", "低"], selection_mode="multi", default=["高", "中", "低"], key="list_levels")
        only_radius = c3.toggle(f"分析半径（{radius}m）内だけ", value=False, key="list_in_radius")
        c4, c5, c6 = st.columns(3)
        min_rating = c4.slider("評価の下限", min_value=0.0, max_value=5.0, value=0.0, step=0.1, key="list_min_rating")
        min_reviews = c5.number_input("口コミ件数の下限", min_value=0, value=0, step=10, key="list_min_reviews")
        late_only = c6.toggle("深夜営業の店だけ", value=False, key="list_late") if "深夜営業" in data.columns else False
        required_attrs = st.multiselect(
            "この属性が「あり」の店だけ", attr_labels, placeholder="指定しない", key="list_attrs"
        ) if attr_labels else []

    mask = pd.Series(True, index=data.index)
    if query.strip():
        text_cols = [c for c in ["店舗名", "住所", "カテゴリ"] if c in data.columns]
        haystack = data[text_cols].fillna("").astype(str).agg(" ".join, axis=1)
        mask &= haystack.str.contains(query.strip(), case=False, regex=False)
    if selected_cats:
        mask &= cats.isin(selected_cats)
    mask &= data["競合度"].isin(levels or [])
    if only_radius:
        mask &= pd.to_numeric(data[dist_col], errors="coerce").fillna(999999) <= radius
    mask &= pd.to_numeric(data["Google評価"], errors="coerce").fillna(0) >= min_rating
    mask &= pd.to_numeric(data["口コミ件数"], errors="coerce").fillna(0) >= min_reviews
    if late_only:
        mask &= data["深夜営業"].fillna("").astype(str).str.strip() == "○"
    for attr in required_attrs:
        mask &= data[f"{ATTR_PREFIX}{attr}"].fillna("").astype(str).str.strip() == "あり"
    result = data[mask].sort_values("口コミ件数", ascending=False)

    default_cols = [
        c for c in [
            "自店", "店舗名", "カテゴリ", "競合度", "競合度の理由", dist_col, "Google評価", "口コミ件数",
            "最新口コミ日", "価格帯", "深夜営業", "住所", "GoogleマップURL",
        ]
        if c in data.columns
    ]
    hidden = {"緯度", "経度", "拠点住所", "拠点緯度", "拠点経度", "収集半径(m)", "place_id"}
    all_cols = [c for c in data.columns if c not in hidden]
    with st.expander("表示する列を選ぶ"):
        show_cols = st.multiselect("表示する列", all_cols, default=default_cols, key=f"list_columns_{dist_col}", label_visibility="collapsed")
    show_cols = show_cols or default_cols

    st.caption(f"{len(result)}件 / 全{len(data)}件（口コミ件数の多い順）")
    st.dataframe(
        result[show_cols],
        width="stretch",
        hide_index=True,
        column_config={
            "GoogleマップURL": st.column_config.LinkColumn("Googleマップ", display_text="開く"),
            "Google評価": st.column_config.NumberColumn(format="%.1f"),
            "口コミ件数": st.column_config.NumberColumn(format="%d"),
        },
    )
    st.download_button(
        "表示中の一覧をCSVで保存",
        result[show_cols].to_csv(index=False).encode("utf-8-sig"),
        file_name="店舗一覧.csv",
        mime="text/csv",
        icon=":material/download:",
        on_click="ignore",
        help="Excelで開いても文字化けしない形式（BOM付きUTF-8）で保存します。",
    )


st.set_page_config(page_title="競合分析ダッシュボード", page_icon="📍", layout="wide")
st.title("競合分析ダッシュボード")
st.caption("MEO対策のための周辺競合・差別化分析（距離・評価・口コミ・属性・カテゴリ）")

with st.expander("このツールについて（どんな人に役立つか）"):
    st.markdown(
        """
Googleはローカル検索の順位を主に **関連性・距離・知名度** の3つで決めていると公表しています。
このダッシュボードは、周辺の同業店と比べて「自分の店は何ができるか」を3つの観点で確認するためのものです。

- **すでに店を営業しているオーナーの方**: 「多くの競合が登録しているのに自店は未登録」の属性や、自店だけの強みが一覧で分かります（差別化分析タブ）。
  キーワードごとの順位と、上位に入るための口コミ件数の目安（キーワード順位タブ）、場所による順位の違い（順位マップタブ）も確認できます。
- **これから出店を考えている方**: サイドバーで予定の業態（カテゴリ）を選び、候補地の住所を入れると、
  周りの同業店の数・評価・口コミ件数を契約前に比べられます（出店候補地タブ）。
- **MEO対策・Web集客を支援している方**: お客さまの初回診断や提案資料の材料に使えます。

**このツールで分からないこと**: キーワードの月間検索回数は分かりません（分かるのは「何位に出るか」の目安です）。
説明文・写真・投稿など、APIで取得できない項目も含みません。
"""
    )

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
    if uploaded_file is None:
        st.info("サンプル（架空の店舗）データを表示中です。調査結果ファイルをアップロードすると、そのデータに切り替わります。")


if uploaded_file is not None:
    try:
        sheets = read_sheets(uploaded_file.name, uploaded_file)
    except Exception as exc:  # pragma: no cover
        st.error(f"ファイル読み込みに失敗しました: {exc}")
        st.stop()
else:
    try:
        sheets = read_sheets_cached(str(SAMPLE_PATH), SAMPLE_PATH.stat().st_mtime)
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
    own_base: Optional[pd.Series] = None if own_choice == NO_OWN_STORE else df.loc[own_choice]

    # 比較するカテゴリ: 初期値は自店のカテゴリ。自店がない（出店検討中の）場合は予定の業態を選ぶ
    all_cats = df["カテゴリ"].fillna("").astype(str).str.strip() if "カテゴリ" in df.columns else pd.Series("", index=df.index)
    category_options = all_cats[all_cats != ""].value_counts().index.tolist()
    own_cat = str(own_base.get("カテゴリ", "") or "").strip() if own_base is not None else ""
    target_categories: List[str] = st.multiselect(
        "比較するカテゴリ（複数可）",
        category_options,
        default=[own_cat] if own_cat in category_options else [],
        key=f"target_categories_{own_choice}",
        placeholder="指定しない（全店舗）",
        help=(
            "地図の色（競合度）・差別化分析の比較対象・出店候補地で数える競合の基準になります。"
            "複数選ぶと1つのグループとして比べます（例: バー＋ラウンジバー）。出店前なら予定している業態を選んでください。"
        ),
    )
    radius_filter = st.selectbox(
        "分析半径",
        [300, 500, 1000],
        index=1,
        format_func=lambda v: f"{v}m",
        help="自店を選んだ場合は自店から、選ばない場合は拠点からの半径です。",
    )

df = assign_competition_level(df, target_categories)
own_row: Optional[pd.Series] = None if own_base is None else df.loc[own_choice]

if origin_latlng is None:
    origin_lat = float(pd.to_numeric(df["緯度"], errors="coerce").mean()) if "緯度" in df.columns else 34.6822
    origin_lng = float(pd.to_numeric(df["経度"], errors="coerce").mean()) if "経度" in df.columns else 135.5008
else:
    origin_lat, origin_lng = origin_latlng

# 距離は自店を選んだら自店から、選ばなければ拠点から測る
if own_row is not None and pd.notna(own_row.get("緯度")) and "緯度" in df.columns:
    center_lat, center_lng = float(own_row["緯度"]), float(own_row["経度"])
    dist_col = "自店からの距離(m)"
    df[dist_col] = distance_from(center_lat, center_lng, df["緯度"], df["経度"]).round()
    own_row = df.loc[own_choice]
else:
    center_lat, center_lng = origin_lat, origin_lng
    dist_col = distance_col

in_radius = df[pd.to_numeric(df[dist_col], errors="coerce").fillna(999999) <= radius_filter]
# 競合 = 半径内の店舗から自店を除いたもの
competitors = in_radius.drop(index=own_choice, errors="ignore") if own_row is not None else in_radius

if competitors.empty:
    st.warning("分析半径内に店舗がありません。半径を広げてください。")
    st.stop()

# Summary: 比較するカテゴリの競合（指定しない場合は全店舗）と自店を比べる
target = competitors[category_mask(competitors, target_categories)]
target_label = category_label(target_categories)
c1, c2, c3, c4 = st.columns(4)
c1.metric(f"競合（{target_label}・半径{radius_filter}m）", f"{len(target):,}店")
c2.metric("強い競合（地図の赤）", f"{int((competitors['競合度'] == '高').sum())}店")
c3.metric(
    "評価の中央値",
    f"{target['Google評価'].median(skipna=True):.1f}" if target["Google評価"].notna().any() else "-",
    delta=(f"自店 {own_row['Google評価']:.1f}" if own_row is not None and pd.notna(own_row["Google評価"]) else None),
    delta_color="off",
)
c4.metric(
    "口コミ件数の中央値",
    f"{target['口コミ件数'].median(skipna=True):.0f}件" if target["口コミ件数"].notna().any() else "-",
    delta=(f"自店 {int(own_row['口コミ件数'])}件" if own_row is not None and pd.notna(own_row["口コミ件数"]) else None),
    delta_color="off",
)

tab_map, tab_diff, tab_kw, tab_grid, tab_site, tab_list = st.tabs(
    ["地図", "差別化分析", "キーワード順位", "順位マップ", "出店候補地", "店舗一覧"]
)
area_radius = collected_radius(df, distance_col)

with tab_map:
    st.subheader("競合マップ")
    if "緯度" in competitors.columns and "経度" in competitors.columns:
        dot_size = st.slider("点の大きさ", min_value=6, max_value=30, value=MAP_DOT_SIZE, step=2, key="map_dot_size")
        render_map(competitors, center_lat, center_lng, radius_filter, origin_lat, origin_lng, own_row, dot_size)
        if not target_categories:
            st.caption(
                f"凡例: 赤=評価{STRONG_RATING}以上で口コミも全店舗の中央値以上 / 橙=どちらか一方を満たす / 青=どちらも満たさない / "
                "緑=自店 / 黒=拠点。サイドバーで「比較するカテゴリ」を選ぶと、業態を考えた色分けになります。"
            )
        else:
            st.caption(
                f"凡例: 赤=「{target_label}」で評価{STRONG_RATING}以上・口コミが比較するカテゴリの中央値以上（上位を取り合う強い相手） / "
                f"橙=「{target_label}」だがそれ以外、または別カテゴリでも評価{STRONG_RATING}以上で口コミが全店舗の上位25% / "
                "青=それ以外 / 緑=自店 / 黒=拠点。点にカーソルを合わせると店名が出ます。"
            )
        st.caption(
            "色はGoogleの評価基準のうち「関連性」（同じカテゴリか）と「知名度」（評価・口コミ件数）で決めています。"
            "「距離」は分析半径の円で表しています。各店の判定理由は「店舗一覧」タブで確認できます。"
        )
    else:
        st.info("緯度/経度列がないため地図を描画できません。")

with tab_diff:
    render_differentiation_tab(competitors, own_row, target_categories)

with tab_kw:
    render_keyword_tab(sheets.get(KEYWORD_SHEET), df, own_row, target_categories)

with tab_grid:
    render_rank_map_tab(sheets.get(GRID_SHEET), df, own_row, origin_lat, origin_lng, area_radius)

with tab_site:
    render_candidate_tab(df, target_categories, own_row, origin_lat, origin_lng, origin_address or origin_label, area_radius)

with tab_list:
    render_store_list_tab(df, own_row, radius_filter, dist_col)
