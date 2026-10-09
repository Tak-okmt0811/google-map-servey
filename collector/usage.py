"""Google Maps Platform の無料枠を超えないための呼び出し回数ガード。

料金区分（SKU）ごとに今月の呼び出し回数を `.api_usage.json` に記録し、
無料枠に安全率をかけた上限を超えそうな場合は、APIを呼ぶ前に BudgetExceeded を送出する。

- 無料枠は「請求先アカウント単位・月単位」で、月の区切りは米国太平洋時間。
- 記録できるのはこのツールがこのPCで行った呼び出しだけ。同じ請求先アカウントで
  他のアプリがAPIを使っている場合、その分は把握できないため安全率で余裕を持たせている。
- 回数はリクエスト送信前に記録する（失敗したリクエストも数える＝安全側）。
"""

from __future__ import annotations

import json
import threading
from datetime import datetime
from pathlib import Path
from typing import Dict
from zoneinfo import ZoneInfo

LEDGER_PATH = Path(__file__).resolve().parent.parent / ".api_usage.json"

GEOCODING = "geocoding"
TEXT_SEARCH = "text_search_enterprise"
PLACE_DETAILS = "place_details_enterprise_atmosphere"

# 2025年3月以降の料金体系での月間無料枠（https://developers.google.com/maps/billing-and-pricing/pricing）
FREE_CAPS: Dict[str, int] = {
    GEOCODING: 10_000,
    TEXT_SEARCH: 1_000,
    PLACE_DETAILS: 1_000,
}

SKU_LABELS: Dict[str, str] = {
    GEOCODING: "Geocoding",
    TEXT_SEARCH: "Text Search Enterprise",
    PLACE_DETAILS: "Place Details Enterprise + Atmosphere",
}

# 他アプリでの利用など把握できない分のための余裕
SAFETY_RATIO = 0.8


class BudgetExceeded(RuntimeError):
    pass


def billing_month() -> str:
    return datetime.now(ZoneInfo("America/Los_Angeles")).strftime("%Y-%m")


class UsageGuard:
    def __init__(self, ledger_path: Path = LEDGER_PATH) -> None:
        self.ledger_path = ledger_path
        self.month = billing_month()
        self._lock = threading.Lock()
        self._ledger = self._load()

    def _load(self) -> Dict[str, Dict[str, int]]:
        if not self.ledger_path.exists():
            return {}
        return json.loads(self.ledger_path.read_text(encoding="utf-8"))

    def _save(self) -> None:
        self.ledger_path.write_text(
            json.dumps(self._ledger, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def used(self, sku: str) -> int:
        return int(self._ledger.get(self.month, {}).get(sku, 0))

    def limit(self, sku: str) -> int:
        return int(FREE_CAPS[sku] * SAFETY_RATIO)

    def remaining(self, sku: str) -> int:
        return max(0, self.limit(sku) - self.used(sku))

    def check(self, sku: str, planned: int) -> None:
        """planned 回の呼び出しが上限内に収まるかを、呼び出し前に確認する。"""
        if self.used(sku) + planned > self.limit(sku):
            raise BudgetExceeded(
                f"{SKU_LABELS[sku]}: 今月の使用 {self.used(sku)} 回 + 予定 {planned} 回が、"
                f"安全上限 {self.limit(sku)} 回（無料枠 {FREE_CAPS[sku]} 回の{int(SAFETY_RATIO * 100)}%）を超えるため中止しました。"
                "費用は発生していません。"
            )

    def record(self, sku: str, count: int = 1) -> None:
        with self._lock:
            self.check(sku, count)
            month = self._ledger.setdefault(self.month, {})
            month[sku] = int(month.get(sku, 0)) + count
            self._save()

    def summary(self) -> str:
        lines = [f"[USAGE] {self.month}（米国太平洋時間）のこのツールによる呼び出し回数:"]
        for sku in FREE_CAPS:
            lines.append(
                f"  - {SKU_LABELS[sku]}: {self.used(sku)} / 安全上限 {self.limit(sku)}（無料枠 {FREE_CAPS[sku]}）"
            )
        return "\n".join(lines)
