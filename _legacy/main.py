#!/usr/bin/env python3
"""Export nearby place information using Google Places API (New / v1) & Geocoding API to Excel & CSV.

Usage example:
    GOOGLE_MAPS_API_KEY=xxxxx python google_places_export_new.py
    GOOGLE_MAPS_API_KEY=xxxxx python google_places_export_new.py --address "大阪府淀川市西中島3丁目" --radius 500
"""

from __future__ import annotations

import argparse
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime
from math import atan2, cos, radians, sin, sqrt
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests


GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"
PLACES_SEARCH_TEXT_URL = "https://places.googleapis.com/v1/places:searchText"
PLACES_DETAILS_BASE_URL = "https://places.googleapis.com/v1/places/"

GENRE_MAP = {
    "bar": "BAR",
    "night_club": "BAR",
    "liquor_store": "BAR",
    "restaurant": "レストラン",
    "meal_takeaway": "レストラン",
    "cafe": "カフェ",
    "bakery": "ベーカリー",
}

SUBGENRE_KEYWORDS = {
    "ショットバー": ["shot", "ショット"],
    "ダイニングバー": ["dining", "ダイニング"],
    "ワインバー": ["wine", "ワイン"],
    "スポーツバー": ["sports", "スポーツ"],
    "居酒屋": ["izakaya", "居酒屋"],
    "焼鳥": ["yakitori", "焼き鳥", "焼鳥"],
    "イタリアン": ["italian", "イタリアン"],
}

FEATURE_KEYWORDS = {
    "個室": ["個室", "private"],
    "喫煙": ["喫煙", "smoking"],
    "ダーツ": ["ダーツ", "darts"],
    "カラオケ": ["カラオケ", "karaoke"],
    "テラス": ["テラス", "terrace"],
}

PRICE_MAP = {
    "PRICE_LEVEL_FREE": "無料",
    "PRICE_LEVEL_INEXPENSIVE": "￥1,000未満",
    "PRICE_LEVEL_MODERATE": "￥1,000-2,000",
    "PRICE_LEVEL_EXPENSIVE": "￥2,000-4,000",
    "PRICE_LEVEL_VERY_EXPENSIVE": "￥4,000以上",
}

WEEKDAYS_JA = ["月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"]


def output_columns(origin_label: str) -> List[str]:
    return [
        "店舗名",
        "ジャンル",
        "サブジャンル",
        "住所",
        f"{origin_label}からの距離(m)",
        "Google評価",
        "口コミ件数",
        "価格帯",
        "営業時間",
        "定休日",
        "電話番号",
        "公式サイト/SNS",
        "特徴",
        "深夜営業",
        "競合度",
        "ターゲット",
        "備考",
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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Google Places API (New) -> Excel exporter")
    parser.add_argument(
        "--address",
        default="大阪府淀川市西中島3丁目",
        help="検索基点の住所 (デフォルト: 大阪府淀川市西中島3丁目)",
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


def geocode_address(address: str, api_key: str) -> Tuple[float, float, str]:
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
    """中心点 + オフセット点を返す（検索漏れ防止用）"""
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


def fetch_nearby_places(lat: float, lng: float, radius: int, api_key: str) -> Dict[str, PlaceSummary]:
    keywords = ["bar", "居酒屋", "焼き鳥", "イタリアン", "ダイニング", "スナック", "和食", "カフェ", "レストラン", "ラーメン"]
    places: Dict[str, PlaceSummary] = {}
    centers = _sub_centers(lat, lng, radius)

    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": "places.id,places.displayName,places.location,places.formattedAddress,places.types,places.rating,places.userRatingCount,nextPageToken",
    }

    for kw in keywords:
        print(f"[INFO] Places API (New) キーワード検索: {kw}")
        for c_lat, c_lng, c_radius in centers:
            page_token = None
            while True:
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

                try:
                    res = requests.post(PLACES_SEARCH_TEXT_URL, json=payload, headers=headers, timeout=20)
                    res.raise_for_status()
                    data = res.json()
                except requests.RequestException as exc:
                    print(f"[WARN] リクエストエラー: {exc}")
                    break

                results = data.get("places", [])
                for p in results:
                    place_id = p.get("id")
                    if not place_id:
                        continue
                    loc = p.get("location", {})
                    p_lat = loc.get("latitude")
                    p_lng = loc.get("longitude")
                    if p_lat is None or p_lng is None:
                        continue

                    display_name = p.get("displayName", {}).get("text", "")
                    places[place_id] = PlaceSummary(
                        place_id=place_id,
                        name=display_name,
                        lat=p_lat,
                        lng=p_lng,
                        vicinity=p.get("formattedAddress", ""),
                        types=p.get("types", []),
                        rating=p.get("rating"),
                        user_ratings_total=p.get("userRatingCount"),
                    )

                page_token = data.get("nextPageToken")
                if not page_token:
                    break
                time.sleep(1.5)  # ページネーション時のウェイト

    return places


def fetch_place_details(place_id: str, api_key: str) -> Dict[str, object]:
    url = f"{PLACES_DETAILS_BASE_URL}{place_id}"
    headers = {
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": "id,displayName,formattedAddress,nationalPhoneNumber,internationalPhoneNumber,websiteUri,rating,userRatingCount,priceLevel,regularOpeningHours,types,editorialSummary",
    }
    params = {"languageCode": "ja"}

    try:
        res = requests.get(url, headers=headers, params=params, timeout=20)
        res.raise_for_status()
        return res.json()
    except Exception:
        return {}


def fetch_all_place_details(place_ids: List[str], api_key: str, workers: int) -> Dict[str, Dict[str, object]]:
    if not place_ids:
        return {}

    safe_workers = max(1, min(workers, 16, len(place_ids)))
    details_map: Dict[str, Dict[str, object]] = {}

    with ThreadPoolExecutor(max_workers=safe_workers) as executor:
        future_map = {
            executor.submit(fetch_place_details, place_id, api_key): place_id
            for place_id in place_ids
        }

        done_count = 0
        total = len(future_map)
        for future in as_completed(future_map):
            place_id = future_map[future]
            try:
                details_map[place_id] = future.result()
            except Exception:
                details_map[place_id] = {}

            done_count += 1
            if done_count % 25 == 0 or done_count == total:
                print(f"[INFO] 詳細取得進捗: {done_count}/{total}")

    return details_map


def infer_genre(types: List[str], name: str) -> str:
    for t in types:
        if t in GENRE_MAP:
            return GENRE_MAP[t]

    n = name.lower()
    if "居酒屋" in name:
        return "居酒屋"
    if "焼鳥" in name or "焼き鳥" in name:
        return "焼鳥"
    if "bar" in n:
        return "BAR"
    if "イタリアン" in name:
        return "イタリアン"
    return "その他"


def infer_subgenre(name: str, types: List[str]) -> str:
    text = f"{name} {' '.join(types)}".lower()
    found = [label for label, keys in SUBGENRE_KEYWORDS.items() if any(k.lower() in text for k in keys)]
    return " / ".join(found) if found else ""


def infer_features(name: str, weekday_text: List[str], reviews_text: str) -> str:
    corpus = f"{name} {' '.join(weekday_text)} {reviews_text}".lower()
    feats = [label for label, keys in FEATURE_KEYWORDS.items() if any(k.lower() in corpus for k in keys)]
    return "・".join(feats)


def infer_closed_day(weekday_text: List[str]) -> str:
    if not weekday_text:
        return ""
    for line in weekday_text:
        if any(x in line for x in ["休業", "定休日", "休み", "closed"]):
            return line
    return ""


def infer_late_night(weekday_text: List[str]) -> str:
    text = " ".join(weekday_text)
    for marker in ["0:00", "00:00", "1:00", "01:00", "2:00", "02:00", "3:00", "03:00", "4:00", "04:00", "5:00", "05:00"]:
        if marker in text:
            return "○"
    return "×"


def infer_competitiveness(distance_m: int, rating: Optional[float], reviews: Optional[int], genre: str) -> str:
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

    if genre == "BAR":
        score += 1

    if score >= 5:
        return "高"
    if score >= 3:
        return "中"
    return "低"


def infer_target(genre: str, price_text: str, late_night: str) -> str:
    targets: List[str] = []
    if genre in {"BAR", "居酒屋"}:
        targets.append("サラリーマン")
    if genre in {"イタリアン", "カフェ"}:
        targets.append("女性")
    if "￥2,000" in price_text or "￥4,000" in price_text:
        targets.append("デート/会食")
    if late_night == "○":
        targets.append("2次会")
    if not targets:
        targets.append("近隣住民")
    return "・".join(dict.fromkeys(targets))


def build_rows(
    origin_label: str,
    origin_lat: float,
    origin_lng: float,
    places: Dict[str, PlaceSummary],
    api_key: str,
    detail_workers: int,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    details_map = fetch_all_place_details(
        [p.place_id for p in places.values()],
        api_key,
        workers=detail_workers,
    )

    for p in places.values():
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

        summary = details.get("editorialSummary", {})
        review_text = summary.get("text", "") if isinstance(summary, dict) else ""

        distance_m = haversine_m(origin_lat, origin_lng, p.lat, p.lng)
        genre = infer_genre(types, name)
        subgenre = infer_subgenre(name, types)
        closed_day = infer_closed_day(weekday_text)
        late_night = infer_late_night(weekday_text)
        features = infer_features(name, weekday_text, review_text)
        competitiveness = infer_competitiveness(distance_m, rating, reviews, genre)
        target = infer_target(genre, price_text, late_night)

        rows.append(
            {
                "店舗名": name,
                "ジャンル": genre,
                "サブジャンル": subgenre,
                "住所": address,
                f"{origin_label}からの距離(m)": distance_m,
                "Google評価": rating,
                "口コミ件数": reviews,
                "価格帯": price_text,
                "営業時間": opening_hours_str,
                "定休日": closed_day,
                "電話番号": phone,
                "公式サイト/SNS": website,
                "特徴": features,
                "深夜営業": late_night,
                "競合度": competitiveness,
                "ターゲット": target,
                "備考": "",
            }
        )

    sort_key = f"{origin_label}からの距離(m)"

    def sort_distance(row: Dict[str, object]) -> int:
        val = row.get(sort_key)
        return int(val) if isinstance(val, (int, float)) else 999999

    return sorted(rows, key=sort_distance)


def normalize_origin_label(address: str) -> str:
    if "西中島3丁目" in address:
        return "西中島3丁目"
    return address


def main() -> None:
    args = parse_args()
    api_key = get_api_key()

    print(f"[INFO] 住所をジオコーディング中: {args.address}")
    lat, lng, formatted = geocode_address(args.address, api_key)
    origin_label = normalize_origin_label(args.address)
    print(f"[INFO] 基点座標: {lat:.6f}, {lng:.6f} ({formatted})")

    print(f"[INFO] Places API (New) 実行中: radius={args.radius}m")
    places = fetch_nearby_places(lat, lng, args.radius, api_key)
    if not places:
        raise RuntimeError("検索結果が0件でした。条件を見直してください。")

    print(f"[INFO] 詳細取得中: {len(places)}件")
    rows = build_rows(origin_label, lat, lng, places, api_key, args.detail_workers)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = args.output or f"places_{origin_label}_{args.radius}m_{ts}.xlsx"
    df = pd.DataFrame(rows)
    expected_cols = output_columns(origin_label)
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""
    df = df[expected_cols]
    df.to_excel(output, index=False)
    print(f"[DONE] Excel出力完了: {output} ({len(df)}件)")

    csv_dir = Path(output).resolve().parent / "data"
    csv_dir.mkdir(exist_ok=True)
    csv_path = csv_dir / "places.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"[DONE] CSV出力完了: {csv_path} (Community Cloud配布用)")


if __name__ == "__main__":
    main()
    
"""
with open('google_places_export_new.py', 'w', encoding='utf-8') as f:
    f.write(script_content)

print("Saved google_places_export_new.py successfully!")
"""