"""Campi derivati da hierarchy_path: country, org_path (condivisi export + app)."""

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


def detect_country(chain: list[str], market_lookup: dict[str, str]) -> str:
    for name in reversed(chain):
        key = name.lower().strip()
        if key in market_lookup:
            return market_lookup[key]
        for mk, canonical in market_lookup.items():
            if len(mk) < 4:
                continue
            if mk in key or key in mk:
                return canonical
    for name in reversed(chain):
        low = name.lower().strip()
        if low and low not in _GENERIC_DEPT_NAMES and len(low) > 3:
            return name
    return ""


def load_market_lookup(app_dir: Path) -> dict[str, str]:
    path = default_guidelines_path(app_dir)
    if not path.exists():
        return {}
    idx = GuidelinesIndex()
    idx.load(path)
    return idx.market_name_lookup()


def add_org_columns(
    df: pd.DataFrame,
    dept_names: dict[str, str],
    market_lookup: dict[str, str] | None = None,
) -> pd.DataFrame:
    if df.empty or "hierarchy_path" not in df.columns:
        return df

    lookup = market_lookup or {}
    org_paths: list[str] = []
    countries: list[str] = []

    for _, row in df.iterrows():
        chain = dept_chain(str(row.get("hierarchy_path", "")), dept_names)
        org_paths.append(" › ".join(chain) if chain else "")
        countries.append(detect_country(chain, lookup))

    out = df.copy()
    out["org_path"] = org_paths
    out["country"] = countries
    return out
