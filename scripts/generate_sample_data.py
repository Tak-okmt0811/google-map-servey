#!/usr/bin/env python3
"""公開デモ用の架空データ（viewer/data/sample_places.xlsx）を生成する。

collector/export.py と同じシート構成・列構成で、実在しない店舗・住所・口コミを作る。
Google APIは一切呼ばない。

    uv run python scripts/generate_sample_data.py
"""

from __future__ import annotations

import math
import random
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from collector.export import (  # noqa: E402
    KEYWORD_SHEET,
    REVIEW_SHEET,
    STORE_SHEET,
    attribute_columns,
    haversine_m,
    output_columns,
)

OUT_PATH = ROOT / "viewer" / "data" / "sample_places.xlsx"
ORIGIN_LABEL = "サンプル拠点"
ORIGIN_ADDRESS = "（デモ用ダミー住所）大阪府サンプル市サンプル1丁目"
ORIGIN_LAT = 34.726893
ORIGIN_LNG = 135.501036
TODAY = date(2026, 10, 1)
N_STORES = 60
KEYWORDS = ["bar", "居酒屋", "焼き鳥", "イタリアン", "ダイニング", "スナック", "ラーメン"]

# カテゴリ -> (ジャンル, 店名の接尾辞, ヒットしやすいキーワード)
CATEGORIES = {
    "居酒屋": ("居酒屋", ["酒場", "食堂"], ["居酒屋", "ダイニング", "焼き鳥"]),
    "バー": ("BAR", ["Bar", "Lounge"], ["bar", "スナック"]),
    "焼き鳥店": ("焼鳥", ["焼鳥", "串焼き"], ["焼き鳥", "居酒屋"]),
    "イタリア料理店": ("イタリアン", ["トラットリア", "バル"], ["イタリアン", "ダイニング", "bar"]),
    "ラーメン屋": ("レストラン", ["らーめん", "麺処"], ["ラーメン"]),
    "ダイニングバー": ("BAR", ["ダイニング", "ビストロ"], ["ダイニング", "bar", "居酒屋"]),
    "スナック": ("BAR", ["スナック", "パブ"], ["スナック", "bar"]),
    "カフェ": ("カフェ", ["カフェ", "喫茶"], ["ダイニング"]),
}
CATEGORY_WEIGHTS = [26, 18, 12, 10, 10, 10, 8, 6]

# 属性ごとの「あり」の確率（周辺でよくある/珍しいの差が出るように設定）
ATTRIBUTE_RATES = {
    "テラス席": 0.10, "団体向け": 0.70, "予約可": 0.65, "カクテル": 0.55, "ワイン": 0.45,
    "ビール": 0.90, "ライブ音楽": 0.05, "スポーツ観戦": 0.08, "テイクアウト": 0.30,
    "デリバリー": 0.15, "店内飲食": 0.97, "ランチ": 0.35, "ディナー": 0.92, "デザート": 0.25,
    "ベジタリアン": 0.10, "子連れ歓迎": 0.30, "キッズメニュー": 0.05, "ペット可": 0.03,
    "トイレ": 0.85, "クレジットカード": 0.75, "タッチ決済": 0.45, "現金のみ": 0.10,
    "車椅子対応入口": 0.20,
}
REGISTRATION_RATE = 0.7  # 属性が登録されている（未登録でない）確率

# サンプルの自店: 珍しい属性を持ち、多くの競合が登録済みの項目が未登録、という差が見える設定
OWN_STORE_ATTRIBUTES = {
    "テラス席": "あり", "ライブ音楽": "あり", "スポーツ観戦": "あり", "ビール": "あり",
    "カクテル": "あり", "店内飲食": "あり", "ディナー": "あり", "クレジットカード": "なし",
    "団体向け": "", "予約可": "", "タッチ決済": "",
}

REVIEW_TEMPLATES = [
    "雰囲気が良く、ゆっくり飲めました。", "料理がどれも美味しかったです。", "スタッフの対応が丁寧でした。",
    "少し待ち時間が長かったです。", "コスパが良いと思います。", "駅から近くて便利です。",
    "団体でも利用しやすかったです。", "お酒の種類が豊富でした。", "深夜まで開いていて助かります。",
    "テラス席が気持ちよかったです。",
]


def jitter(lat: float, lng: float, max_radius_m: float) -> tuple[float, float]:
    r = random.uniform(30, max_radius_m)
    theta = random.uniform(0, 2 * math.pi)
    earth = 6378137.0
    d_lat = (r * math.sin(theta) / earth) * (180 / math.pi)
    d_lng = (r * math.cos(theta) / (earth * math.cos(math.radians(lat)))) * (180 / math.pi)
    return lat + d_lat, lng + d_lng


def attribute_value(label: str) -> str:
    if random.random() > REGISTRATION_RATE:
        return ""
    return "あり" if random.random() < ATTRIBUTE_RATES[label] else "なし"


def main() -> None:
    random.seed(42)
    distance_col = f"{ORIGIN_LABEL}からの距離(m)"
    stores, keyword_rows, review_rows = [], [], []
    names_used: set[str] = set()

    for i in range(N_STORES):
        is_own = i == 0
        category = "ダイニングバー" if is_own else random.choices(list(CATEGORIES), CATEGORY_WEIGHTS)[0]
        genre, suffixes, hit_keywords = CATEGORIES[category]
        if is_own:
            name = "【サンプル自店】デモダイニング"
        else:
            name = f"{random.choice(['サンプル', 'デモ', 'テスト', 'モデル', '架空'])}{random.choice(suffixes)} {chr(65 + i % 26)}{i}"
        names_used.add(name)

        lat, lng = (jitter(ORIGIN_LAT, ORIGIN_LNG, 120) if is_own else jitter(ORIGIN_LAT, ORIGIN_LNG, 950))
        distance = haversine_m(ORIGIN_LAT, ORIGIN_LNG, lat, lng)
        rating = 4.1 if is_own else round(random.uniform(3.0, 4.8), 1)
        review_count = 38 if is_own else random.choice(
            [random.randint(1, 20), random.randint(20, 120), random.randint(120, 600)]
        )
        late = "○" if (is_own or random.random() < 0.35) else "×"

        attrs = {}
        for col in attribute_columns():
            label = col.removeprefix("属性_")
            attrs[col] = OWN_STORE_ATTRIBUTES.get(label, attribute_value(label)) if is_own else attribute_value(label)

        n_reviews = 5 if review_count >= 5 else review_count
        days_back = random.choice([5, 20, 45, 120, 300]) if not is_own else 150
        dates = sorted((TODAY - timedelta(days=days_back + random.randint(0, 200)) for _ in range(n_reviews)), reverse=True)
        for d in dates:
            review_rows.append({
                "place_id": f"sample_{i}", "店舗名": name, "評価": random.randint(3, 5),
                "投稿日": d.isoformat(), "本文": random.choice(REVIEW_TEMPLATES),
            })

        stores.append({
            "place_id": f"sample_{i}",
            "店舗名": name,
            "ジャンル": genre,
            "サブジャンル": "",
            "カテゴリ": category,
            "住所": f"（ダミー住所）大阪府サンプル市サンプル{i % 5 + 1}丁目{i + 1}-{i % 9 + 1}",
            "緯度": round(lat, 6),
            "経度": round(lng, 6),
            distance_col: distance,
            "Google評価": rating,
            "口コミ件数": review_count,
            "最新口コミ日": dates[0].isoformat() if dates else "",
            "価格帯": random.choice(["￥1,000-2,000", "￥2,000-4,000", "￥4,000以上", ""]),
            "営業時間": "17:00-翌2:00" if late == "○" else "17:00-23:00",
            "定休日": random.choice(["月曜日", "日曜日", "不定休", ""]),
            "電話番号": "" if (not is_own and random.random() < 0.1) else f"06-0000-{1000 + i:04d}",
            "公式サイト/SNS": "" if (is_own or random.random() < 0.4) else "https://example.com/",
            "GoogleマップURL": "",
            "営業状態": "閉業" if (not is_own and random.random() < 0.03) else "営業中",
            "特徴": "",
            "深夜営業": late,
            "競合度": "",
            "ターゲット": "",
            "紹介文": "",
            "口コミ要約": "",
            **attrs,
            "備考": "サンプルデータ（実店舗ではありません）",
            "拠点住所": ORIGIN_ADDRESS,
            "拠点緯度": ORIGIN_LAT,
            "拠点経度": ORIGIN_LNG,
            "自店": "○" if is_own else "",
        })

        own_ranks = {"ダイニング": 4, "bar": 17, "居酒屋": 31}
        for kw in KEYWORDS:
            if is_own:
                if kw in own_ranks or kw in hit_keywords:
                    keyword_rows.append({"キーワード": kw, "順位": own_ranks.get(kw), "place_id": f"sample_{i}",
                                         "店舗名": name, "距離(m)": distance})
                continue
            if kw in hit_keywords or random.random() < 0.08:
                rank = random.randint(1, 60) if random.random() < 0.6 else None
                keyword_rows.append({"キーワード": kw, "順位": rank, "place_id": f"sample_{i}",
                                     "店舗名": name, "距離(m)": distance})

    store_df = pd.DataFrame(stores).reindex(columns=output_columns(ORIGIN_LABEL) + ["自店"])
    store_df = store_df.sort_values(distance_col)
    keyword_df = pd.DataFrame(keyword_rows).sort_values(["キーワード", "順位"], na_position="last")
    review_df = pd.DataFrame(review_rows)

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_PATH) as writer:
        store_df.to_excel(writer, sheet_name=STORE_SHEET, index=False)
        keyword_df.to_excel(writer, sheet_name=KEYWORD_SHEET, index=False)
        review_df.to_excel(writer, sheet_name=REVIEW_SHEET, index=False)
    print(f"wrote {OUT_PATH}: 店舗 {len(store_df)} / キーワード順位 {len(keyword_df)} / 口コミ {len(review_df)}")


if __name__ == "__main__":
    main()
