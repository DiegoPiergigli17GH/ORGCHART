"""Campi derivati da hierarchy_path: country, market, org_path (export + app)."""

from __future__ import annotations

import base64
from pathlib import Path

import pandas as pd

from guidelines_index import GuidelinesIndex, default_guidelines_path

_GENERIC_DEPT_NAMES = frozenset({
    "group", "corporate", "amea", "eur", "emea", "apac", "americas",
    "europe", "east and west europe", "west europe", "east europe",
    "supply chain & operations", "group human resources",
})


def dept_chain(hierarchy_path: str, dept_names: dict[str, str]) -> list[str]:
    chain: list[str] = []
    for part in hierarchy_path.split(" > "):
        if not part.startswith("FODepartment/"):
            continue
        dept_id = part.split("/", 1)[-1].split(":")[0]
        name = dept_names.get(dept_id, "")
        if not name:
            try:
                pad = "=" * (4 - len(dept_id) % 4) if len(dept_id) % 4 else ""
                name = base64.b64decode(dept_id + pad).decode("utf-8", errors="ignore")
            except Exception:
                name = dept_id
        if name and (not chain or chain[-1] != name):
            chain.append(name)
    return chain


def detect_market(
    chain: list[str],
    name_lookup: dict[str, str],
    code_by_country: dict[str, str],
    code_to_country: dict[str, str],
) -> tuple[str, str]:
    """Ritorna (country, sigla Market guidelines es. RN)."""
    for name in reversed(chain):
        key = name.lower().strip()
        if not key:
            continue
        if key in name_lookup:
            country = name_lookup[key]
            code = code_by_country.get(country.lower(), "")
            if not code and key in code_to_country:
                code = _code_for_key(key, code_to_country, code_by_country)
            return country, code
        for mk, canonical in name_lookup.items():
            if len(mk) < 4:
                continue
            if mk in key or key in mk:
                return canonical, code_by_country.get(canonical.lower(), "")

    for name in reversed(chain):
        low = name.lower().strip()
        if low and low not in _GENERIC_DEPT_NAMES and len(low) > 3:
            return name, code_by_country.get(low, "")
    return "", ""


def _code_for_key(key: str, code_to_country: dict[str, str], code_by_country: dict[str, str]) -> str:
    country = code_to_country.get(key, "")
    if country:
        return code_by_country.get(country.lower(), key.upper())
    return key.upper()


def load_market_lookups(
    app_dir: Path,
) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    path = default_guidelines_path(app_dir)
    if not path.exists():
        return {}, {}, {}
    idx = GuidelinesIndex()
    idx.load(path)
    return idx.market_lookups()


def load_market_lookup(app_dir: Path) -> dict[str, str]:
    """Compat: solo name_lookup."""
    return load_market_lookups(app_dir)[0]


def add_org_columns(
    df: pd.DataFrame,
    dept_names: dict[str, str],
    market_lookup: dict[str, str] | None = None,
    market_lookups: tuple[dict[str, str], dict[str, str], dict[str, str]] | None = None,
) -> pd.DataFrame:
    if df.empty or "hierarchy_path" not in df.columns:
        return df

    if market_lookups:
        name_lookup, code_by_country, code_to_country = market_lookups
    else:
        name_lookup = market_lookup or {}
        code_by_country = {}
        code_to_country = {}

    org_paths: list[str] = []
    countries: list[str] = []
    markets: list[str] = []

    for _, row in df.iterrows():
        chain = dept_chain(str(row.get("hierarchy_path", "")), dept_names)
        org_paths.append(" › ".join(chain) if chain else "")
        country, market = detect_market(
            chain, name_lookup, code_by_country, code_to_country
        )
        countries.append(country)
        markets.append(market)

    out = df.copy()
    out["org_path"] = org_paths
    out["country"] = countries
    out["market"] = markets
    return out
