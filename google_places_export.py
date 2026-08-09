#!/usr/bin/env python3
"""Export nearby place information from Google Places API to Excel.

Usage example:
    GOOGLE_MAPS_API_KEY=xxxxx python google_places_export.py
    GOOGLE_MAPS_API_KEY=xxxxx python google_places_export.py --address "大阪府淀川市西中島3丁目" --radius 500
"""

from __future__ import annotations

import argparse
import os
import re
import time
from dataclasses import dataclass
from datetime import datetime
from math import atan2, cos, radians, sin, sqrt
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests


GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"
NEARBY_URL = "https://maps.googleapis.com/maps/api/place/nearbysearch/json"
DETAILS_URL = "https://maps.googleapis.com/maps/api/place/details/json"


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
    0: "無料",
    1: "￥1,000未満",
    2: "￥1,000-2,000",
    3: "￥2,000-4,000",
    4: "￥4,000以上",
}

WEEKDAYS_JA = ["月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"]


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
    parser = argparse.ArgumentParser(description="Google Places API -> Excel exporter")
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
    return parser.parse_args()


def get_api_key() -> str:
    load_env_file()
    key = os.getenv("GOOGLE_MAPS_API_KEY", "").strip()
    if not key:
        raise RuntimeError(
            "環境変数 GOOGLE_MAPS_API_KEY が設定されていません。"
        )
    return key


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


def split_postal_and_address(address: str) -> Tuple[str, str]:
    cleaned = re.sub(r"^\s*日本[、,\s]*", "", address.strip())
    postal_code = ""

    m = re.search(r"〒?\s*(\d{3}-\d{4})", cleaned)
    if m:
        postal_code = m.group(1)
        cleaned = (cleaned[: m.start()] + cleaned[m.end() :]).strip(" 、,")

    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    return postal_code, cleaned


def normalize_time_text(text: str) -> str:
    t = text
    t = t.replace("～", "-").replace("−", "-").replace("ー", "-").replace("—", "-").replace("–", "-")
    t = t.replace("：", ":")
    t = t.replace("翌", "")
    t = re.sub(r"(\d{1,2})時(\d{1,2})分", lambda m: f"{int(m.group(1)):02d}:{int(m.group(2)):02d}", t)
    t = re.sub(r"(\d{1,2})時", lambda m: f"{int(m.group(1)):02d}:00", t)
    t = re.sub(r"\s+", "", t)
    return t


def parse_opening_hours_by_day(weekday_text: List[str]) -> Dict[str, Dict[str, str]]:
    day_hours: Dict[str, Dict[str, str]] = {
        d: {"raw": "", "start": "", "end": ""} for d in WEEKDAYS_JA
    }

    for line in weekday_text:
        m = re.match(r"^(月曜日|火曜日|水曜日|木曜日|金曜日|土曜日|日曜日)\s*[:：]\s*(.+)$", line.strip())
        if not m:
            continue

        day = m.group(1)
        value = m.group(2).strip()
        normalized = normalize_time_text(value)
        day_hours[day]["raw"] = normalized

        if any(x in normalized.lower() for x in ["休業", "定休日", "休み", "closed"]):
            day_hours[day]["raw"] = "休業"
            continue

        if "24時間営業" in normalized:
            day_hours[day]["start"] = "00:00"
            day_hours[day]["end"] = "24:00"
            continue

        ranges = re.findall(r"(\d{1,2}:\d{2})-(\d{1,2}:\d{2})", normalized)
        if ranges:
            start, _ = ranges[0]
            _, end = ranges[-1]
            day_hours[day]["start"] = start
            day_hours[day]["end"] = end

    return day_hours


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


def fetch_nearby_places(lat: float, lng: float, radius: int, api_key: str) -> Dict[str, PlaceSummary]:
    keywords = ["bar", "居酒屋", "焼き鳥", "イタリアン"]
    places: Dict[str, PlaceSummary] = {}

    for kw in keywords:
        next_token: Optional[str] = None
        page_count = 0

        while True:
            params = {
                "location": f"{lat},{lng}",
                "radius": radius,
                "keyword": kw,
                "language": "ja",
                "key": api_key,
            }
            if next_token:
                params = {"pagetoken": next_token, "key": api_key, "language": "ja"}

            res = requests.get(NEARBY_URL, params=params, timeout=20)
            res.raise_for_status()
            payload = res.json()
            status = payload.get("status")

            if status not in {"OK", "ZERO_RESULTS"}:
                if status == "INVALID_REQUEST" and next_token:
                    time.sleep(2)
                    continue
                raise RuntimeError(f"Nearby Search失敗: status={status}")

            for r in payload.get("results", []):
                place_id = r.get("place_id")
                if not place_id:
                    continue
                geo = r.get("geometry", {}).get("location", {})
                if "lat" not in geo or "lng" not in geo:
                    continue

                places[place_id] = PlaceSummary(
                    place_id=place_id,
                    name=r.get("name", ""),
                    lat=geo["lat"],
                    lng=geo["lng"],
                    vicinity=r.get("vicinity", ""),
                    types=r.get("types", []),
                    rating=r.get("rating"),
                    user_ratings_total=r.get("user_ratings_total"),
                )

            next_token = payload.get("next_page_token")
            page_count += 1
            if not next_token or page_count >= 3:
                break
            time.sleep(2)

    return places


def fetch_place_details(place_id: str, api_key: str) -> Dict[str, object]:
    fields = [
        "name",
        "formatted_address",
        "international_phone_number",
        "website",
        "rating",
        "user_ratings_total",
        "price_level",
        "opening_hours",
        "types",
        "editorial_summary",
    ]
    params = {
        "place_id": place_id,
        "fields": ",".join(fields),
        "language": "ja",
        "key": api_key,
    }
    res = requests.get(DETAILS_URL, params=params, timeout=20)
    res.raise_for_status()
    payload = res.json()
    status = payload.get("status")
    if status != "OK":
        return {}
    return payload.get("result", {})


def build_rows(
    origin_label: str,
    origin_lat: float,
    origin_lng: float,
    places: Dict[str, PlaceSummary],
    api_key: str,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []

    for p in places.values():
        details = fetch_place_details(p.place_id, api_key)
        types_val = details.get("types")
        types = types_val if isinstance(types_val, list) else p.types
        name = str(details.get("name") or p.name)
        raw_address = str(details.get("formatted_address") or p.vicinity)
        postal_code, address = split_postal_and_address(raw_address)

        opening_hours_val = details.get("opening_hours")
        opening_hours = opening_hours_val if isinstance(opening_hours_val, dict) else {}
        weekday_text_val = opening_hours.get("weekday_text")
        weekday_text = weekday_text_val if isinstance(weekday_text_val, list) else []
        day_hours = parse_opening_hours_by_day(weekday_text)
        phone = details.get("international_phone_number", "")
        website = details.get("website", "")
        rating_val = details.get("rating", p.rating)
        rating = float(rating_val) if isinstance(rating_val, (int, float)) else p.rating
        reviews_val = details.get("user_ratings_total", p.user_ratings_total)
        reviews = int(reviews_val) if isinstance(reviews_val, (int, float)) else p.user_ratings_total
        price_level_val = details.get("price_level")
        price_level = int(price_level_val) if isinstance(price_level_val, int) else None
        price_text = PRICE_MAP.get(price_level, "") if price_level is not None else ""

        summary = details.get("editorial_summary", {})
        review_text = summary.get("overview", "") if isinstance(summary, dict) else ""

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
                "郵便番号": postal_code,
                "住所": address,
                "緯度": p.lat,
                "経度": p.lng,
                f"{origin_label}からの距離(m)": distance_m,
                "Google評価": rating,
                "口コミ件数": reviews,
                "価格帯": price_text,
                "月曜営業時間": day_hours["月曜日"]["raw"],
                "月曜営業開始": day_hours["月曜日"]["start"],
                "月曜営業終了": day_hours["月曜日"]["end"],
                "火曜営業時間": day_hours["火曜日"]["raw"],
                "火曜営業開始": day_hours["火曜日"]["start"],
                "火曜営業終了": day_hours["火曜日"]["end"],
                "水曜営業時間": day_hours["水曜日"]["raw"],
                "水曜営業開始": day_hours["水曜日"]["start"],
                "水曜営業終了": day_hours["水曜日"]["end"],
                "木曜営業時間": day_hours["木曜日"]["raw"],
                "木曜営業開始": day_hours["木曜日"]["start"],
                "木曜営業終了": day_hours["木曜日"]["end"],
                "金曜営業時間": day_hours["金曜日"]["raw"],
                "金曜営業開始": day_hours["金曜日"]["start"],
                "金曜営業終了": day_hours["金曜日"]["end"],
                "土曜営業時間": day_hours["土曜日"]["raw"],
                "土曜営業開始": day_hours["土曜日"]["start"],
                "土曜営業終了": day_hours["土曜日"]["end"],
                "日曜営業時間": day_hours["日曜日"]["raw"],
                "日曜営業開始": day_hours["日曜日"]["start"],
                "日曜営業終了": day_hours["日曜日"]["end"],
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

    print(f"[INFO] Nearby Search実行中: radius={args.radius}m")
    places = fetch_nearby_places(lat, lng, args.radius, api_key)
    if not places:
        raise RuntimeError("検索結果が0件でした。条件を見直してください。")

    print(f"[INFO] 詳細取得中: {len(places)}件")
    rows = build_rows(origin_label, lat, lng, places, api_key)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = args.output or f"places_{origin_label}_{args.radius}m_{ts}.xlsx"
    df = pd.DataFrame(rows)
    df.to_excel(output, index=False)
    print(f"[DONE] Excel出力完了: {output} ({len(df)}件)")

    # Community Cloud配布用: data/places.csv に上書き保存
    csv_dir = Path(output).resolve().parent / "data"
    csv_dir.mkdir(exist_ok=True)
    csv_path = csv_dir / "places.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"[DONE] CSV出力完了: {csv_path} (Community Cloud配布用)")


if __name__ == "__main__":
    main()