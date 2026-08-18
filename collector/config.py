"""収集プロファイル（検索キーワード・ジャンル判定辞書など）の読み込み。"""

from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List

DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent / "configs" / "bar_izakaya.toml"


@dataclass
class CollectorConfig:
    default_address: str
    keywords: List[str]
    genre_map: Dict[str, str] = field(default_factory=dict)
    genre_name_fallback: Dict[str, str] = field(default_factory=dict)
    subgenre_keywords: Dict[str, List[str]] = field(default_factory=dict)
    feature_keywords: Dict[str, List[str]] = field(default_factory=dict)
    salaryman_genres: List[str] = field(default_factory=list)
    women_genres: List[str] = field(default_factory=list)


def load_config(path: Path | str = DEFAULT_CONFIG_PATH) -> CollectorConfig:
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"設定ファイルが見つかりません: {config_path}")

    with config_path.open("rb") as f:
        raw = tomllib.load(f)

    search = raw.get("search", {})
    target_rules = raw.get("target_rules", {})

    return CollectorConfig(
        default_address=search.get("default_address", ""),
        keywords=list(search.get("keywords", [])),
        genre_map=dict(raw.get("genre_map", {})),
        genre_name_fallback=dict(raw.get("genre_name_fallback", {})),
        subgenre_keywords={k: list(v) for k, v in raw.get("subgenre_keywords", {}).items()},
        feature_keywords={k: list(v) for k, v in raw.get("feature_keywords", {}).items()},
        salaryman_genres=list(target_rules.get("salaryman_genres", [])),
        women_genres=list(target_rules.get("women_genres", [])),
    )
