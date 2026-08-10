#!/usr/bin/env python3
"""Google Places API (New) & Geocoding API 疎通確認用（1件取得テスト）

1回の実行で以下の3リクエストのみを発行します。
1. Geocoding API (住所 -> 緯度経度)
2. Places API (New) searchText (周辺検索 -> 1件のみ取得)
3. Places API (New) details (詳細情報取得)
"""

from __future__ import annotations

import os
from pathlib import Path
import requests

GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"
PLACES_SEARCH_TEXT_URL = "https://places.googleapis.com/v1/places:searchText"
PLACES_DETAILS_BASE_URL = "https://places.googleapis.com/v1/places/"


def load_env_file(env_path: str = ".env") -> None:
    """.env ファイルから環境変数を読み込む"""
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
        raise RuntimeError("環境変数 GOOGLE_MAPS_API_KEY が設定されていません。")
    return key


def test_geocoding(address: str, api_key: str) -> tuple[float, float, str]:
    """1. Geocoding API のテスト"""
    print(f"\n--- [1/3] Geocoding API テスト: {address} ---")
    params = {"address": address, "language": "ja", "key": api_key}
    res = requests.get(GEOCODE_URL, params=params, timeout=10)
    
    print(f"ステータスコード: {res.status_code}")
    res.raise_for_status()
    
    data = res.json()
    status = data.get("status")
    print(f"API応答ステータス: {status}")
    if status != "OK" or not data.get("results"):
        raise RuntimeError(f"Geocoding失敗: {data}")

    first = data["results"][0]
    loc = first["geometry"]["location"]
    formatted_address = first.get("formatted_address", address)
    print(f"成功 -> 緯度: {loc['lat']}, 経度: {loc['lng']}, 変換後の住所: {formatted_address}")
    return loc["lat"], loc["lng"], formatted_address


def test_places_search(lat: float, lng: float, radius: int, api_key: str) -> dict:
    """2. Places API (New) searchText のテスト（1件取得）"""
    print(f"\n--- [2/3] Places API (New) SearchText テスト ---")
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress",
    }
    
    payload = {
        "textQuery": "居酒屋",
        "pageSize": 1,
        "locationBias": {
            "circle": {
                "center": {"latitude": lat, "longitude": lng},
                "radius": float(radius),
            }
        },
        "languageCode": "ja",
    }

    res = requests.post(PLACES_SEARCH_TEXT_URL, json=payload, headers=headers, timeout=10)
    print(f"ステータスコード: {res.status_code}")
    
    # 200以外の場合、APIから返されたエラー詳細を表示する
    if res.status_code != 200:
        print(f"\n[エラー詳細内容]\n{res.text}\n")
        
    res.raise_for_status()

    data = res.json()
    places = data.get("places", [])
    if not places:
        raise RuntimeError("検索結果が0件でした。")

    first_place = places[0]
    place_id = first_place.get("id")
    name = first_place.get("displayName", {}).get("text", "")
    print(f"成功 -> Place ID: {place_id}, 店舗名: {name}")
    return first_place


def test_place_details(place_id: str, api_key: str) -> dict:
    """3. Places API (New) Details のテスト"""
    print(f"\n--- [3/3] Places API (New) Place Details テスト ---")
    url = f"{PLACES_DETAILS_BASE_URL}{place_id}"
    headers = {
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": "id,displayName,formattedAddress,nationalPhoneNumber,rating",
    }
    params = {"languageCode": "ja"}

    res = requests.get(url, headers=headers, params=params, timeout=10)
    print(f"ステータスコード: {res.status_code}")
    
    if res.status_code != 200:
        print(f"エラーレスポンス内容: {res.text}")
    res.raise_for_status()

    data = res.json()
    print(f"成功 -> 名前: {data.get('displayName', {}).get('text')}")
    print(f"       住所: {data.get('formattedAddress')}")
    print(f"       電話: {data.get('nationalPhoneNumber')}")
    print(f"       評価: {data.get('rating')}")
    return data


def main() -> None:
    try:
        api_key = get_api_key()
        test_address = "大阪府淀川市西中島3丁目"
        
        # 1. ジオコーディング
        lat, lng, _ = test_geocoding(test_address, api_key)
        
        # 2. キーワード検索（1件のみ）
        place = test_places_search(lat, lng, 500, api_key)
        
        # 3. 詳細検索
        place_id = place.get("id")
        if place_id:
            test_place_details(place_id, api_key)
            
        print("\n==========================================")
        print(" [成功] すべてのAPIの疎通が正常に完了しました！")
        print("==========================================")

    except Exception as e:
        print("\n==========================================")
        print(f" [失敗] エラーが発生しました: {e}")
        print("==========================================")


if __name__ == "__main__":
    main()