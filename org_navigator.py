"""Navigazione organigramma da Excel estratto + match guidelines."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from guidelines_index import GuidelinesIndex, MarketMatch
from org_enrich import add_org_columns, dept_chain, load_market_lookups


class OrgNavigator:
    def __init__(self):
        self.df: pd.DataFrame = pd.DataFrame()
        self.dept_df: pd.DataFrame = pd.DataFrame()
        self.source_path: Path | None = None
        self._market_lookup: dict[str, str] = {}

    @property
    def loaded(self) -> bool:
        return not self.df.empty

    @property
    def employee_count(self) -> int:
        return int(self.df["user_id"].nunique()) if self.loaded else 0

    def load_excel(self, path: Path, guidelines: GuidelinesIndex | None = None) -> int:
        self.source_path = path
        xl = pd.ExcelFile(path)
        sheet = "All_Employees" if "All_Employees" in xl.sheet_names else xl.sheet_names[0]
        self.df = pd.read_excel(path, sheet_name=sheet)
        if "Departments" in xl.sheet_names:
            self.dept_df = pd.read_excel(path, sheet_name="Departments")
        else:
            self.dept_df = pd.DataFrame()
        if guidelines:
            self.enrich(guidelines)
        return len(self.df)

    def load_dataframe(self, df: pd.DataFrame, guidelines: GuidelinesIndex | None = None) -> int:
        self.source_path = None
        self.df = df.copy()
        if guidelines:
            self.enrich(guidelines)
        return len(self.df)

    def find_latest_export(self, output_dir: Path) -> Path | None:
        files = sorted(output_dir.glob("orgchart_*.xlsx"), key=lambda p: p.stat().st_mtime)
        return files[-1] if files else None

    def auto_load(self, output_dir: Path, guidelines: GuidelinesIndex | None = None) -> Path | None:
        latest = self.find_latest_export(output_dir)
        if latest:
            self.load_excel(latest, guidelines=guidelines)
        return latest

    def enrich(self, guidelines: GuidelinesIndex) -> None:
        """Aggiunge country e org_path analizzando hierarchy_path."""
        if not self.loaded:
            return

        self._market_lookup = guidelines.market_name_lookup()
        self.df = add_org_columns(
            self.df,
            self.dept_name_map(),
            market_lookups=guidelines.market_lookups(),
        )

    def filter_by_market(self, match: MarketMatch) -> pd.DataFrame:
        """Solo persone il cui percorso org contiene il mercato (es. Romania), non tutta l'area EUR."""
        if not self.loaded:
            return pd.DataFrame()
        return self._filter_org_terms(match.org_filter_terms, match_exact_short_codes=True)

    def filter_by_text(
        self,
        query: str,
        include_roles: bool = False,
    ) -> pd.DataFrame:
        """
        Ricerca libera: tutte le parole devono comparire (AND).
        Default: solo organigramma (country, org_path, department) — non il ruolo.
        """
        if not self.loaded or not query.strip():
            return pd.DataFrame()

        tokens = [t for t in query.strip().split() if t]
        if not tokens:
            return pd.DataFrame()

        mask = pd.Series(True, index=self.df.index)
        for token in tokens:
            mask &= self._token_mask(token, include_roles=include_roles)
        return self.df[mask].copy()

    def _filter_org_terms(
        self,
        terms: list[str],
        match_exact_short_codes: bool = False,
    ) -> pd.DataFrame:
        if not terms:
            return pd.DataFrame()

        mask = pd.Series(False, index=self.df.index)
        for term in terms:
            if not term:
                continue
            mask |= self._token_mask(
                term,
                include_roles=False,
                org_only=True,
                exact_short_code=match_exact_short_codes and len(term) <= 3,
            )
        return self.df[mask].copy()

    def _token_mask(
        self,
        token: str,
        include_roles: bool = False,
        org_only: bool = False,
        exact_short_code: bool = False,
    ) -> pd.Series:
        q = token.strip().lower()
        if not q:
            return pd.Series(True, index=self.df.index)

        org_cols = ["country", "market", "org_path", "department_name", "hierarchy_path"]
        extra_cols = ["job_title", "full_name", "email"] if include_roles and not org_only else []
        cols = [c for c in org_cols + extra_cols if c in self.df.columns]

        mask = pd.Series(False, index=self.df.index)
        for col in cols:
            if col in ("country", "market", "org_path", "department_name"):
                mask |= self.df[col].astype(str).apply(
                    lambda v, c=col: _org_text_matches(q, v, exact_short_code)
                )
            elif col == "hierarchy_path":
                mask |= self.df[col].astype(str).apply(
                    lambda v: _hierarchy_matches(q, v, exact_short_code)
                )
            else:
                mask |= self.df[col].astype(str).str.lower().str.contains(
                    q, na=False, regex=False
                )
        return mask

    def dept_name_map(self) -> dict[str, str]:
        m: dict[str, str] = {}
        if not self.dept_df.empty:
            for _, r in self.dept_df.iterrows():
                m[str(r.get("department_id", ""))] = str(r.get("department_name", ""))
        if self.loaded and "department_id" in self.df.columns:
            for _, r in self.df.drop_duplicates("department_id").iterrows():
                m[str(r["department_id"])] = str(r.get("department_name", ""))
        return m

    def build_dept_tree(self, subset: pd.DataFrame) -> dict:
        """Albero annidato: {nome_dept: {_employees: [...], child_name: {...}}}."""
        root: dict = {}
        names = self.dept_name_map()

        for _, row in subset.iterrows():
            chain = dept_chain(str(row.get("hierarchy_path", "")), names)
            node = root
            for dept_name in chain:
                node = node.setdefault(dept_name, {})
            node.setdefault("_employees", []).append({
                "full_name": row.get("full_name", ""),
                "job_title": row.get("job_title", ""),
                "email": row.get("email", ""),
                "department_name": row.get("department_name", ""),
                "country": row.get("country", ""),
                "market": row.get("market", ""),
                "org_path": row.get("org_path", ""),
            })

        return root

    def employees_in_dept(self, subset: pd.DataFrame, dept_path: str) -> list[dict]:
        if dept_path:
            m = subset["department_name"].astype(str) == dept_path
            rows = subset[m]
        else:
            rows = subset
        return [_employee_row_dict(r) for _, r in rows.iterrows()]


def _employee_row_dict(r) -> dict:
    return {
        "full_name": r.get("full_name", ""),
        "job_title": r.get("job_title", ""),
        "email": r.get("email", ""),
        "department_name": r.get("department_name", ""),
        "country": r.get("country", ""),
        "market": r.get("market", ""),
        "org_path": r.get("org_path", ""),
    }


def _org_text_matches(token: str, text: str, exact_short_code: bool) -> bool:
    text_l = str(text).lower()
    if not text_l or text_l == "nan":
        return False
    if exact_short_code and len(token) <= 3:
        parts = re.split(r"[\s›>/\-]+", text_l)
        return token in parts
    if token in text_l:
        return True
    return any(_word_match(token, seg) for seg in re.split(r"[\s›>]+", text_l))


def _hierarchy_matches(token: str, hierarchy_path: str, exact_short_code: bool) -> bool:
    if exact_short_code and len(token) <= 3:
        return False
    return token in str(hierarchy_path).lower()


def _word_match(token: str, segment: str) -> bool:
    segment = segment.strip()
    if not segment:
        return False
    if token == segment:
        return True
    return bool(re.search(rf"\b{re.escape(token)}\b", segment))


# compat: tree builder import
_dept_chain = dept_chain


def tree_to_tk(parent_dict: dict, parent_id: str, tree, path: str = "") -> None:
    """Popola ttk.Treeview da albero dept."""
    for key in sorted(parent_dict.keys()):
        if key == "_employees":
            continue
        node_path = f"{path}/{key}" if path else key
        iid = tree.insert(parent_id, "end", text=key, values=(node_path,))
        tree_to_tk(parent_dict[key], iid, tree, node_path)
