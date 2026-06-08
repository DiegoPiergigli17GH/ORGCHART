"""Indice mercati / aree da guidelines.xlsx per ricerca RN, Romania, ecc."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class MarketMatch:
    market: str
    description: str
    division: str = ""
    division_name: str = ""
    area: str = ""
    area_name: str = ""
    sub_area: str = ""
    sub_area_name: str = ""
    search_terms: list[str] = field(default_factory=list)

    @property
    def breadcrumb(self) -> str:
        parts = [
            self.division_name or self.division,
            self.area_name or self.area,
            self.sub_area_name or self.sub_area,
            f"{self.description} ({self.market})",
        ]
        return " › ".join(p for p in parts if p and p != "nan")

    @property
    def label(self) -> str:
        return f"{self.market} — {self.description}"


class GuidelinesIndex:
    LEVELS = [
        ("Division", "Description"),
        ("Cluster", "Description.1"),
        ("Area", "Description.2"),
        ("Sub Area", "Description.3"),
        ("Sub Sub", "Description.4"),
        ("Market", "Description.5"),
        ("SS", "Description.6"),
    ]

    def __init__(self):
        self.path: Path | None = None
        self._markets: list[MarketMatch] = []

    @property
    def loaded(self) -> bool:
        return bool(self._markets)

    @property
    def count(self) -> int:
        return len(self._markets)

    def load(self, path: Path) -> int:
        self.path = path
        raw = pd.read_excel(path)
        self._markets = []

        for _, row in raw.iterrows():
            market = _cell(row.get("Market"))
            desc = _cell(row.get("Description.5"))
            if not market:
                continue

            terms = {market.lower(), desc.lower()}
            for code_col, name_col in self.LEVELS:
                code = _cell(row.get(code_col))
                name = _cell(row.get(name_col))
                if code:
                    terms.add(code.lower())
                if name:
                    terms.add(name.lower())

            self._markets.append(MarketMatch(
                market=market,
                description=desc or market,
                division=_cell(row.get("Division")),
                division_name=_cell(row.get("Description")),
                area=_cell(row.get("Area")),
                area_name=_cell(row.get("Description.2")),
                sub_area=_cell(row.get("Sub Area")),
                sub_area_name=_cell(row.get("Description.3")),
                search_terms=sorted(t for t in terms if t),
            ))

        return len(self._markets)

    def search(self, query: str, limit: int = 25) -> list[MarketMatch]:
        q = query.strip().lower()
        if not q:
            return []

        scored: list[tuple[int, MarketMatch]] = []
        for m in self._markets:
            score = _score_match(m, q)
            if score > 0:
                scored.append((score, m))

        scored.sort(key=lambda x: (-x[0], x[1].market))
        return [m for _, m in scored[:limit]]


def _cell(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _score_match(m: MarketMatch, q: str) -> int:
    if m.market.lower() == q:
        return 100
    if m.description.lower() == q:
        return 95
    if m.market.lower().startswith(q):
        return 80
    if m.description.lower().startswith(q):
        return 75
    # Sigle corte (RN, EES): niente match parziali dentro altre parole
    if len(q) <= 3:
        for t in m.search_terms:
            if t == q:
                return 50
        return 0
    if q in m.description.lower():
        return 60
    for t in m.search_terms:
        if t == q:
            return 50
        if len(t) >= 4 and q in t:
            return 40
    return 0


def default_guidelines_path(app_dir: Path) -> Path:
    return app_dir / "guidelines.xlsx"
