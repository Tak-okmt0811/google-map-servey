#!/usr/bin/env python3
"""Export nearby place information using Google Places API (New / v1) & Geocoding API to Excel & CSV.

このスクリプトは実行する人自身のGoogle Maps Platform APIキーを消費します。
無料枠を超えないよう、collector/usage.py の呼び出し回数ガードを通してAPIを呼びます。

出力するExcelのシート:
  - 店舗:           1店舗1行。カテゴリ・属性（属性_○○列）・口コミ要約などを含む
  - キーワード順位: 検索キーワードごとに、基点で検索したときの表示順位
  - 口コミ:         1店舗あたり最大5件の口コミ（投稿者名は保存しない）

Usage example:
    python -m collector.export --estimate          # API を呼ばずに呼び出し回数の見積もりだけ表示
    python -m collector.export
    python -m collector.export --address "大阪府淀川市西中島3丁目" --radius 500
    python -m collector.export --radius 300 --keywords 居酒屋,bar   # 小規模な動作確認
"""

from __future__ import annotations

import argparse
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from math import atan2, cos, radians, sin, sqrt
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests

try:
    from .config import CollectorConfig, DEFAULT_CONFIG_PATH, load_config
    from .usage import GEOCODING, PLACE_DETAILS, TEXT_SEARCH, BudgetExceeded, UsageGuard
except ImportError:  # `python collector/export.py` のように直接実行された場合
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from collector.config import CollectorConfig, DEFAULT_CONFIG_PATH, load_config
    from collector.usage import GEOCODING, PLACE_DETAILS, TEXT_SEARCH, BudgetExceeded, UsageGuard


GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"
PLACES_SEARCH_TEXT_URL = "https://places.googleapis.com/v1/places:searchText"
PLACES_DETAILS_BASE_URL = "https://places.googleapis.com/v1/places/"

# Text Search は 1ページ最大20件・最大3ページ
MAX_SEARCH_PAGES = 3

# 検索: rating/userRatingCount を含むため Text Search Enterprise 区分
SEARCH_FIELD_MASK = ",".join(
    [
        "places.id",
        "places.displayName",
        "places.location",
        "places.formattedAddress",
        "places.types",
        "places.primaryType",
        "places.primaryTypeDisplayName",
        "places.rating",
        "places.userRatingCount",
        "nextPageToken",
    ]
)

# Google Places の固定enum値なので、業種に関わらず共通（設定ファイル化しない）
PRICE_MAP = {
    "PRICE_LEVEL_FREE": "無料",
    "PRICE_LEVEL_INEXPENSIVE": "￥1,000未満",
    "PRICE_LEVEL_MODERATE": "￥1,000-2,000",
    "PRICE_LEVEL_EXPENSIVE": "￥2,000-4,000",
    "PRICE_LEVEL_VERY_EXPENSIVE": "￥4,000以上",
}

BUSINESS_STATUS_MAP = {
    "OPERATIONAL": "営業中",
    "CLOSED_TEMPORARILY": "臨時休業",
    "CLOSED_PERMANENTLY": "閉業",
}

# Places API の属性（Enterprise + Atmosphere 区分）-> 表示名。値は true/false/未登録 の3通り
ATTRIBUTE_FIELDS = {
    "outdoorSeating": "テラス席",
    "goodForGroups": "団体向け",
    "reservable": "予約可",
    "servesCocktails": "カクテル",
    "servesWine": "ワイン",
    "servesBeer": "ビール",
    "liveMusic": "ライブ音楽",
    "goodForWatchingSports": "スポーツ観戦",
    "takeout": "テイクアウト",
    "delivery": "デリバリー",
    "dineIn": "店内飲食",
    "servesLunch": "ランチ",
    "servesDinner": "ディナー",
    "servesDessert": "デザート",
    "servesVegetarianFood": "ベジタリアン",
    "goodForChildren": "子連れ歓迎",
    "menuForChildren": "キッズメニュー",
    "allowsDogs": "ペット可",
    "restroom": "トイレ",
}

NESTED_ATTRIBUTE_FIELDS = {
    ("paymentOptions", "acceptsCreditCards"): "クレジットカード",
    ("paymentOptions", "acceptsNfc"): "タッチ決済",
    ("paymentOptions", "acceptsCashOnly"): "現金のみ",
    ("accessibilityOptions", "wheelchairAccessibleEntrance"): "車椅子対応入口",
}

# 詳細: editorialSummary 等を含むため Place Details Enterprise + Atmosphere 区分
DETAILS_FIELD_MASK = ",".join(
    [
        "id",
        "displayName",
        "formattedAddress",
        "nationalPhoneNumber",
        "internationalPhoneNumber",
        "websiteUri",
        "googleMapsUri",
        "businessStatus",
        "rating",
        "userRatingCount",
        "priceLevel",
        "regularOpeningHours",
        "types",
        "primaryType",
        "primaryTypeDisplayName",
        "editorialSummary",
        "reviewSummary",
        "generativeSummary",
        "reviews",
        "paymentOptions",
        "accessibilityOptions",
        *ATTRIBUTE_FIELDS.keys(),
    ]
)

STORE_SHEET = "店舗"
KEYWORD_SHEET = "キーワード順位"
REVIEW_SHEET = "口コミ"


def attribute_columns() -> List[str]:
    labels = list(ATTRIBUTE_FIELDS.values()) + list(NESTED_ATTRIBUTE_FIELDS.values())
    return [f"属性_{label}" for label in labels]


def output_columns(origin_label: str) -> List[str]:
    return [
        "place_id",
        "店舗名",
        "ジャンル",
        "サブジャンル",
        "カテゴリ",
        "住所",
        "緯度",
        "経度",
        f"{origin_label}からの距離(m)",
        "Google評価",
        "口コミ件数",
        "最新口コミ日",
        "価格帯",
        "営業時間",
        "定休日",
        "電話番号",
        "公式サイト/SNS",
        "GoogleマップURL",
        "営業状態",
        "特徴",
        "深夜営業",
        "競合度",
        "ターゲット",
        "紹介文",
        "口コミ要約",
        *attribute_columns(),
        "備考",
        "拠点住所",
        "拠点緯度",
        "拠点経度",
    ]


@dataclass
class PlaceSummary:
    place_id: str
    name: str
    lat: float
    lng: float
    vicinity: str
    types: List[str]
    rating: Optional[float]
    user_ratings_total: Optional[int]
    category: str = ""
    keywords: Dict[str, Optional[int]] = field(default_factory=dict)  # キーワード -> 基点検索での順位


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Google Places API (New) -> Excel/CSV exporter")
    parser.add_argument(
        "--config",
        default=str(DEFAULT_CONFIG_PATH),
        help="収集プロファイル（検索キーワード・ジャンル判定辞書）のTOMLファイルパス",
    )
    parser.add_argument(
        "--address",
        default=None,
        help="検索基点の住所（未指定時は設定ファイルのdefault_addressを使用）",
    )
    parser.add_argument(
        "--label",
        default=None,
        help="出力列・ファイル名に使う拠点ラベル（未指定時は--addressの値をそのまま使用）",
    )
    parser.add_argument(
        "--keywords",
        default=None,
        help="検索キーワードをカンマ区切りで上書き（例: 居酒屋,bar）。小規模な動作確認向け",
    )
    parser.add_argument(
        "--radius",
        type=int,
        default=500,
        choices=[300, 500, 1000],
        help="検索半径 (m): 300 / 500 / 1000",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="出力Excelパス。未指定時は日時付きファイル名を生成",
    )
    parser.add_argument(
        "--detail-workers",
        type=int,
        default=8,
        help="詳細取得の並列数 (デフォルト: 8)",
    )
    parser.add_argument(
        "--estimate",
        action="store_true",
        help="APIを呼ばずに、今月の使用回数と今回の最大呼び出し回数の見積もりだけを表示する",
    )
    return parser.parse_args()


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


def get_api_key() -> str:
    load_env_file()
    key = os.getenv("GOOGLE_MAPS_API_KEY", "").strip()
    if not key:
        raise RuntimeError(
            "環境変数 GOOGLE_MAPS_API_KEY が設定されていません。"
        )
    return key


def geocode_address(address: str, api_key: str, guard: UsageGuard) -> Tuple[float, float, str]:
    guard.record(GEOCODING)
    params = {"address": address, "language": "ja", "key": api_key}
    res = requests.get(GEOCODE_URL, params=params, timeout=20)
    res.raise_for_status()
    data = res.json()
    status = data.get("status")
    if status != "OK" or not data.get("results"):
        raise RuntimeError(f"Geocoding失敗: status={status}")

    first = data["results"][0]
    loc = first["geometry"]["location"]
    return loc["lat"], loc["lng"], first.get("formatted_address", address)


def haversine_m(lat1: float, lon1: float, lat2: float, lon2: float) -> int:
    r = 6371000
    d_lat = radians(lat2 - lat1)
    d_lon = radians(lon2 - lon1)
    a = sin(d_lat / 2) ** 2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(d_lon / 2) ** 2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))
    return int(round(r * c))


def _sub_centers(lat: float, lng: float, radius: int) -> List[Tuple[float, float, int]]:
    """中心点 + オフセット点を返す（検索漏れ防止用）。先頭は必ず基点そのもの。"""
    if radius <= 300:
        return [(lat, lng, radius)]
    offset_m = radius * 0.44
    d_lat = offset_m / 111000
    d_lng = offset_m / (111000 * cos(radians(lat)))
    sub_r = int(radius * 0.6)
    diag_lat = d_lat * 0.707
    diag_lng = d_lng * 0.707
    return [
        (lat, lng, radius),
        (lat + d_lat, lng, sub_r),
        (lat - d_lat, lng, sub_r),
        (lat, lng + d_lng, sub_r),
        (lat, lng - d_lng, sub_r),
        (lat + diag_lat, lng + diag_lng, sub_r),
        (lat + diag_lat, lng - diag_lng, sub_r),
        (lat - diag_lat, lng + diag_lng, sub_r),
        (lat - diag_lat, lng - diag_lng, sub_r),
    ]


def max_search_calls(keywords: List[str], radius: int) -> int:
    return len(keywords) * len(_sub_centers(0.0, 0.0, radius)) * MAX_SEARCH_PAGES


def fetch_nearby_places(
    lat: float,
    lng: float,
    radius: int,
    api_key: str,
    keywords: List[str],
    guard: UsageGuard,
) -> Dict[str, PlaceSummary]:
    places: Dict[str, PlaceSummary] = {}
    centers = _sub_centers(lat, lng, radius)

    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": SEARCH_FIELD_MASK,
    }

    for kw in keywords:
        print(f"[INFO] Places API (New) キーワード検索: {kw}")
        for center_index, (c_lat, c_lng, c_radius) in enumerate(centers):
            # 基点（先頭の中心点）での並び順を、そのキーワードでの表示順位として記録する
            is_origin = center_index == 0
            rank = 0
            page_token = None
            for _ in range(MAX_SEARCH_PAGES):
                payload = {
                    "textQuery": kw,
                    "locationBias": {
                        "circle": {
                            "center": {"latitude": c_lat, "longitude": c_lng},
                            "radius": float(c_radius),
                        }
                    },
                    "languageCode": "ja",
                }
                if page_token:
                    payload["pageToken"] = page_token

                guard.record(TEXT_SEARCH)
                try:
                    res = requests.post(PLACES_SEARCH_TEXT_URL, json=payload, headers=headers, timeout=20)
                    res.raise_for_status()
                    data = res.json()
                except requests.RequestException as exc:
                    print(f"[WARN] リクエストエラー: {exc}")
                    break

                for p in data.get("places", []):
                    rank += 1
                    place_id = p.get("id")
                    if not place_id:
                        continue
                    loc = p.get("location", {})
                    p_lat = loc.get("latitude")
                    p_lng = loc.get("longitude")
                    if p_lat is None or p_lng is None:
                        continue

                    summary = places.get(place_id)
                    if summary is None:
                        category = p.get("primaryTypeDisplayName", {})
                        summary = PlaceSummary(
                            place_id=place_id,
                            name=p.get("displayName", {}).get("text", ""),
                            lat=p_lat,
                            lng=p_lng,
                            vicinity=p.get("formattedAddress", ""),
                            types=p.get("types", []),
                            rating=p.get("rating"),
                            user_ratings_total=p.get("userRatingCount"),
                            category=category.get("text", "") if isinstance(category, dict) else "",
                        )
                        places[place_id] = summary

                    if is_origin:
                        prev = summary.keywords.get(kw)
                        summary.keywords[kw] = rank if prev is None else min(prev, rank)
                    else:
                        summary.keywords.setdefault(kw, None)

                page_token = data.get("nextPageToken")
                if not page_token:
                    break
                time.sleep(1.5)

    return places


def fetch_place_details(place_id: str, api_key: str, guard: UsageGuard) -> Dict[str, object]:
    url = f"{PLACES_DETAILS_BASE_URL}{place_id}"
    headers = {
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": DETAILS_FIELD_MASK,
    }
    params = {"languageCode": "ja"}

    guard.record(PLACE_DETAILS)
    try:
        res = requests.get(url, headers=headers, params=params, timeout=20)
        res.raise_for_status()
        return res.json()
    except requests.RequestException:
        return {}


def fetch_all_place_details(
    place_ids: List[str], api_key: str, workers: int, guard: UsageGuard
) -> Dict[str, Dict[str, object]]:
    if not place_ids:
        return {}

    # 件数が確定した時点で、まとめて上限チェックしてから取得を始める
    guard.check(PLACE_DETAILS, len(place_ids))

    safe_workers = max(1, min(workers, 16, len(place_ids)))
    details_map: Dict[str, Dict[str, object]] = {}

    with ThreadPoolExecutor(max_workers=safe_workers) as executor:
        future_map = {
            executor.submit(fetch_place_details, place_id, api_key, guard): place_id
            for place_id in place_ids
        }

        done_count = 0
        total = len(future_map)
        for future in as_completed(future_map):
            place_id = future_map[future]
            try:
                details_map[place_id] = future.result()
            except BudgetExceeded:
                raise
            except Exception:
                details_map[place_id] = {}

            done_count += 1
            if done_count % 25 == 0 or done_count == total:
                print(f"[INFO] 詳細取得進捗: {done_count}/{total}")

    return details_map


def infer_genre(types: List[str], name: str, config: CollectorConfig) -> str:
    for t in types:
        if t in config.genre_map:
            return config.genre_map[t]

    n = name.lower()
    for needle, genre in config.genre_name_fallback.items():
        if needle.lower() in n or needle in name:
            return genre
    return "その他"


def infer_subgenre(name: str, types: List[str], config: CollectorConfig) -> str:
    text = f"{name} {' '.join(types)}".lower()
    found = [
        label
        for label, keys in config.subgenre_keywords.items()
        if any(k.lower() in text for k in keys)
    ]
    return " / ".join(found) if found else ""


def infer_features(name: str, weekday_text: List[str], reviews_text: str, config: CollectorConfig) -> str:
    corpus = f"{name} {' '.join(weekday_text)} {reviews_text}".lower()
    feats = [
        label
        for label, keys in config.feature_keywords.items()
        if any(k.lower() in corpus for k in keys)
    ]
    return "・".join(feats)


def infer_closed_day(weekday_text: List[str]) -> str:
    if not weekday_text:
        return ""
    for line in weekday_text:
        if any(x in line for x in ["休業", "定休日", "休み", "closed"]):
            return line
    return ""


LATE_NIGHT_CLOSE_HOURS = range(0, 6)  # 0時〜5時台に閉店 = 深夜営業


def infer_late_night(weekday_text: List[str], periods: Optional[List[Dict[str, object]]] = None) -> str:
    """閉店時刻が0〜5時台、または24時間営業なら深夜営業とみなす。

    構造化データ（regularOpeningHours.periods）を優先し、無い場合は
    "18時00分～1時00分" / "18:00～1:00" 形式の営業時間テキストから閉店時刻を読む。
    """
    if periods:
        for period in periods:
            close = period.get("close") if isinstance(period, dict) else None
            if not isinstance(close, dict):
                return "○"  # close が無い period は24時間営業
            if close.get("hour") in LATE_NIGHT_CLOSE_HOURS:
                return "○"
        return "×"

    text = " ".join(weekday_text)
    if "24 時間営業" in text or "24時間営業" in text:
        return "○"
    for close_hour in re.findall(r"[～~\-–]\s*(\d{1,2})\s*(?:時|:)", text):
        if int(close_hour) in LATE_NIGHT_CLOSE_HOURS:
            return "○"
    return "×"


def infer_competitiveness(distance_m: int, rating: Optional[float], reviews: Optional[int], genre: str, config: CollectorConfig) -> str:
    score = 0
    if distance_m <= 300:
        score += 2
    elif distance_m <= 500:
        score += 1

    if rating and rating >= 4.0:
        score += 2
    elif rating and rating >= 3.5:
        score += 1

    if reviews and reviews >= 100:
        score += 2
    elif reviews and reviews >= 30:
        score += 1

    if genre in config.salaryman_genres:
        score += 1

    if score >= 5:
        return "高"
    if score >= 3:
        return "中"
    return "低"


def infer_target(genre: str, price_text: str, late_night: str, config: CollectorConfig) -> str:
    targets: List[str] = []
    if genre in config.salaryman_genres:
        targets.append("サラリーマン")
    if genre in config.women_genres:
        targets.append("女性")
    if "￥2,000" in price_text or "￥4,000" in price_text:
        targets.append("デート/会食")
    if late_night == "○":
        targets.append("2次会")
    if not targets:
        targets.append("近隣住民")
    return "・".join(dict.fromkeys(targets))


def _bool_label(value: object) -> str:
    if value is True:
        return "あり"
    if value is False:
        return "なし"
    return ""  # 未登録


def extract_attributes(details: Dict[str, object]) -> Dict[str, str]:
    attrs = {f"属性_{label}": _bool_label(details.get(field_name)) for field_name, label in ATTRIBUTE_FIELDS.items()}
    for (parent, child), label in NESTED_ATTRIBUTE_FIELDS.items():
        group = details.get(parent)
        value = group.get(child) if isinstance(group, dict) else None
        attrs[f"属性_{label}"] = _bool_label(value)
    return attrs


def _nested_text(obj: object, *keys: str) -> str:
    for key in keys:
        if not isinstance(obj, dict):
            return ""
        obj = obj.get(key)
    return obj if isinstance(obj, str) else ""


def extract_reviews(place_id: str, name: str, details: Dict[str, object]) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    reviews = details.get("reviews")
    if not isinstance(reviews, list):
        return rows
    for r in reviews:
        if not isinstance(r, dict):
            continue
        published = str(r.get("publishTime", ""))[:10]
        rows.append(
            {
                "place_id": place_id,
                "店舗名": name,
                "評価": r.get("rating"),
                "投稿日": published,
                "本文": _nested_text(r, "text", "text") or _nested_text(r, "originalText", "text"),
            }
        )
    return rows


def build_rows(
    origin_label: str,
    origin_address: str,
    origin_lat: float,
    origin_lng: float,
    places: Dict[str, PlaceSummary],
    details_map: Dict[str, Dict[str, object]],
    config: CollectorConfig,
) -> Tuple[List[Dict[str, object]], List[Dict[str, object]], List[Dict[str, object]]]:
    """店舗・キーワード順位・口コミの3シート分の行を作る。places は半径内に絞り込み済みであること。"""
    store_rows: List[Dict[str, object]] = []
    keyword_rows: List[Dict[str, object]] = []
    review_rows: List[Dict[str, object]] = []

    for p in places.values():
        distance_m = haversine_m(origin_lat, origin_lng, p.lat, p.lng)
        details = details_map.get(p.place_id, {})
        types_val = details.get("types")
        types = types_val if isinstance(types_val, list) else p.types

        disp = details.get("displayName")
        name = disp.get("text", p.name) if isinstance(disp, dict) else p.name

        address = str(details.get("formattedAddress") or p.vicinity)

        opening_hours = details.get("regularOpeningHours", {})
        weekday_text = opening_hours.get("weekdayDescriptions", []) if isinstance(opening_hours, dict) else []
        opening_hours_str = " | ".join(weekday_text) if weekday_text else ""

        phone = details.get("nationalPhoneNumber") or details.get("internationalPhoneNumber", "")
        website = details.get("websiteUri", "")

        rating_val = details.get("rating", p.rating)
        rating = float(rating_val) if isinstance(rating_val, (int, float)) else p.rating

        reviews_val = details.get("userRatingCount", p.user_ratings_total)
        reviews = int(reviews_val) if isinstance(reviews_val, (int, float)) else p.user_ratings_total

        price_level = details.get("priceLevel")
        price_text = PRICE_MAP.get(str(price_level), "") if price_level else ""

        editorial = _nested_text(details, "editorialSummary", "text")
        review_summary = _nested_text(details, "reviewSummary", "text", "text") or _nested_text(
            details, "generativeSummary", "overview", "text"
        )
        category = _nested_text(details, "primaryTypeDisplayName", "text") or p.category

        place_reviews = extract_reviews(p.place_id, name, details)
        review_rows.extend(place_reviews)
        latest_review = max((r["投稿日"] for r in place_reviews if r["投稿日"]), default="")

        genre = infer_genre(types, name, config)
        subgenre = infer_subgenre(name, types, config)
        closed_day = infer_closed_day(weekday_text)
        periods = opening_hours.get("periods") if isinstance(opening_hours, dict) else None
        late_night = infer_late_night(weekday_text, periods)
        features = infer_features(name, weekday_text, editorial, config)
        competitiveness = infer_competitiveness(distance_m, rating, reviews, genre, config)
        target = infer_target(genre, price_text, late_night, config)

        store_rows.append(
            {
                "place_id": p.place_id,
                "店舗名": name,
                "ジャンル": genre,
                "サブジャンル": subgenre,
                "カテゴリ": category,
                "住所": address,
                "緯度": p.lat,
                "経度": p.lng,
                f"{origin_label}からの距離(m)": distance_m,
                "Google評価": rating,
                "口コミ件数": reviews,
                "最新口コミ日": latest_review,
                "価格帯": price_text,
                "営業時間": opening_hours_str,
                "定休日": closed_day,
                "電話番号": phone,
                "公式サイト/SNS": website,
                "GoogleマップURL": details.get("googleMapsUri", ""),
                "営業状態": BUSINESS_STATUS_MAP.get(str(details.get("businessStatus", "")), ""),
                "特徴": features,
                "深夜営業": late_night,
                "競合度": competitiveness,
                "ターゲット": target,
                "紹介文": editorial,
                "口コミ要約": review_summary,
                **extract_attributes(details),
                "備考": "",
                "拠点住所": origin_address,
                "拠点緯度": origin_lat,
                "拠点経度": origin_lng,
            }
        )

        for kw, rank in p.keywords.items():
            keyword_rows.append(
                {
                    "キーワード": kw,
                    "順位": rank,
                    "place_id": p.place_id,
                    "店舗名": name,
                    "距離(m)": distance_m,
                }
            )

    distance_col = f"{origin_label}からの距離(m)"
    store_rows.sort(key=lambda row: int(row[distance_col]))
    keyword_rows.sort(key=lambda row: (row["キーワード"], row["順位"] is None, row["順位"] or 0))
    return store_rows, keyword_rows, review_rows


def print_estimate(guard: UsageGuard, keywords: List[str], radius: int) -> None:
    print(guard.summary())
    print("[ESTIMATE] 今回の実行で想定される最大呼び出し回数:")
    print("  - Geocoding: 1")
    print(f"  - Text Search: 最大 {max_search_calls(keywords, radius)}（{len(keywords)}キーワード × "
          f"{len(_sub_centers(0.0, 0.0, radius))}地点 × 最大{MAX_SEARCH_PAGES}ページ）")
    print(f"  - Place Details: 半径{radius}m内の店舗数（検索後に確定。取得前に上限チェックを行う）")


def main() -> None:
    args = parse_args()
    config = load_config(args.config)
    keywords = [k.strip() for k in args.keywords.split(",")] if args.keywords else config.keywords
    keywords = [k for k in keywords if k]
    guard = UsageGuard()

    if args.estimate:
        print_estimate(guard, keywords, args.radius)
        return

    address = args.address or config.default_address
    if not address:
        raise RuntimeError("--address が未指定で、設定ファイルにも default_address がありません。")
    origin_label = args.label or address

    # 検索の最大回数が上限内に収まることを、APIを1回も呼ぶ前に確認する
    guard.check(GEOCODING, 1)
    guard.check(TEXT_SEARCH, max_search_calls(keywords, args.radius))
    api_key = get_api_key()
    print(guard.summary())

    print(f"[INFO] 住所をジオコーディング中: {address}")
    lat, lng, formatted = geocode_address(address, api_key, guard)
    print(f"[INFO] 基点座標: {lat:.6f}, {lng:.6f} ({formatted})")

    print(f"[INFO] Places API (New) 実行中: radius={args.radius}m, keywords={keywords}")
    found = fetch_nearby_places(lat, lng, args.radius, api_key, keywords, guard)
    if not found:
        raise RuntimeError("検索結果が0件でした。条件を見直してください。")

    # 詳細取得の前に半径外を除外する（詳細取得は1件ごとに課金対象になるため）
    places = {
        pid: p for pid, p in found.items()
        if haversine_m(lat, lng, p.lat, p.lng) <= args.radius
    }
    print(f"[INFO] 検索ヒット {len(found)}件のうち、半径{args.radius}m内の {len(places)}件の詳細を取得します")
    details_map = fetch_all_place_details(list(places.keys()), api_key, args.detail_workers, guard)

    store_rows, keyword_rows, review_rows = build_rows(
        origin_label, formatted, lat, lng, places, details_map, config
    )

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = args.output or f"places_{origin_label}_{args.radius}m_{ts}.xlsx"
    store_df = pd.DataFrame(store_rows).reindex(columns=output_columns(origin_label))
    keyword_df = pd.DataFrame(keyword_rows, columns=["キーワード", "順位", "place_id", "店舗名", "距離(m)"])
    review_df = pd.DataFrame(review_rows, columns=["place_id", "店舗名", "評価", "投稿日", "本文"])

    with pd.ExcelWriter(output) as writer:
        store_df.to_excel(writer, sheet_name=STORE_SHEET, index=False)
        keyword_df.to_excel(writer, sheet_name=KEYWORD_SHEET, index=False)
        review_df.to_excel(writer, sheet_name=REVIEW_SHEET, index=False)
    print(f"[DONE] Excel出力完了: {output}（店舗 {len(store_df)}件 / キーワード順位 {len(keyword_df)}行 / 口コミ {len(review_df)}件）")

    csv_dir = Path(output).resolve().parent / "data"
    csv_dir.mkdir(exist_ok=True)
    csv_path = csv_dir / "places.csv"
    store_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"[DONE] CSV出力完了: {csv_path}（店舗シートのみ）")
    print(guard.summary())


if __name__ == "__main__":
    try:
        main()
    except BudgetExceeded as exc:
        print(f"[STOP] {exc}")
        raise SystemExit(1)
