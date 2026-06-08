"""Navigazione organigramma da Excel estratto + match guidelines."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from guidelines_index import MarketMatch


class OrgNavigator:
    def __init__(self):
        self.df: pd.DataFrame = pd.DataFrame()
        self.dept_df: pd.DataFrame = pd.DataFrame()
        self.source_path: Path | None = None

    @property
    def loaded(self) -> bool:
        return not self.df.empty

    @property
    def employee_count(self) -> int:
        return int(self.df["user_id"].nunique()) if self.loaded else 0

    def load_excel(self, path: Path) -> int:
        self.source_path = path
        xl = pd.ExcelFile(path)
        sheet = "All_Employees" if "All_Employees" in xl.sheet_names else xl.sheet_names[0]
        self.df = pd.read_excel(path, sheet_name=sheet)
        if "Departments" in xl.sheet_names:
            self.dept_df = pd.read_excel(path, sheet_name="Departments")
        else:
            self.dept_df = pd.DataFrame()
        return len(self.df)

    def load_dataframe(self, df: pd.DataFrame) -> int:
        self.source_path = None
        self.df = df.copy()
        return len(self.df)

    def find_latest_export(self, output_dir: Path) -> Path | None:
        files = sorted(output_dir.glob("orgchart_*.xlsx"), key=lambda p: p.stat().st_mtime)
        return files[-1] if files else None

    def auto_load(self, output_dir: Path) -> Path | None:
        latest = self.find_latest_export(output_dir)
        if latest:
            self.load_excel(latest)
        return latest

    def filter_by_market(self, match: MarketMatch) -> pd.DataFrame:
        if not self.loaded:
            return pd.DataFrame()

        terms = list({
            match.market.lower(),
            match.description.lower(),
            match.sub_area_name.lower(),
            match.area_name.lower(),
            match.division_name.lower(),
        } - {""})

        mask = pd.Series(False, index=self.df.index)
        for col in ("department_name", "hierarchy_path", "job_title", "full_name"):
            if col not in self.df.columns:
                continue
            col_mask = pd.Series(False, index=self.df.index)
            for t in terms:
                if len(t) >= 2:
                    col_mask |= self.df[col].astype(str).str.lower().str.contains(
                        t, na=False, regex=False
                    )
            mask |= col_mask

        return self.df[mask].copy()

    def filter_by_text(self, query: str) -> pd.DataFrame:
        if not self.loaded or not query.strip():
            return pd.DataFrame()

        q = query.strip().lower()
        mask = pd.Series(False, index=self.df.index)
        for col in ("department_name", "hierarchy_path", "job_title", "full_name", "email"):
            if col in self.df.columns:
                mask |= self.df[col].astype(str).str.lower().str.contains(
                    q, na=False, regex=False
                )
        return self.df[mask].copy()

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
            chain = _dept_chain(str(row.get("hierarchy_path", "")), names)
            node = root
            for dept_name in chain:
                node = node.setdefault(dept_name, {})
            node.setdefault("_employees", []).append({
                "full_name": row.get("full_name", ""),
                "job_title": row.get("job_title", ""),
                "email": row.get("email", ""),
                "department_name": row.get("department_name", ""),
            })

        return root

    def employees_in_dept(self, subset: pd.DataFrame, dept_path: str) -> list[dict]:
        if dept_path:
            m = subset["department_name"].astype(str) == dept_path
            rows = subset[m]
        else:
            rows = subset
        return [
            {
                "full_name": r.get("full_name", ""),
                "job_title": r.get("job_title", ""),
                "email": r.get("email", ""),
                "department_name": r.get("department_name", ""),
            }
            for _, r in rows.iterrows()
        ]


def _dept_chain(hierarchy_path: str, names: dict[str, str]) -> list[str]:
    chain: list[str] = []
    for part in hierarchy_path.split(" > "):
        if not part.startswith("FODepartment/"):
            continue
        dept_id = part.split("/", 1)[-1].split(":")[0]
        name = names.get(dept_id, "")
        if not name:
            try:
                import base64
                pad = "=" * (4 - len(dept_id) % 4) if len(dept_id) % 4 else ""
                name = base64.b64decode(dept_id + pad).decode("utf-8", errors="ignore")
            except Exception:
                name = dept_id
        if name and (not chain or chain[-1] != name):
            chain.append(name)
    return chain


def tree_to_tk(parent_dict: dict, parent_id: str, tree, path: str = "") -> None:
    """Popola ttk.Treeview da albero dept."""
    for key in sorted(parent_dict.keys()):
        if key == "_employees":
            continue
        node_path = f"{path}/{key}" if path else key
        iid = tree.insert(parent_id, "end", text=key, values=(node_path,))
        tree_to_tk(parent_dict[key], iid, tree, node_path)
