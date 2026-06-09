"""
Ariston Org Chart Extractor — GUI

Tab Estrazione: incolla token → scarica Excel
Tab Esplora: cerca mercato (RN, Romania…) e naviga l'organigramma caricato
"""

from __future__ import annotations

import logging
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

from guidelines_index import GuidelinesIndex, default_guidelines_path
from org_navigator import OrgNavigator, tree_to_tk
from orgchart_crawler import (
    APP_DIR,
    LOCAL_SESSION_PATH,
    _load_yaml,
    build_runtime_config,
    parse_clipboard_tokens,
    run_extraction,
)


class QueueLogHandler(logging.Handler):
    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue
        self.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))

    def emit(self, record):
        self.log_queue.put(self.format(record))


class OrgChartApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ariston Org Chart")
        self.minsize(720, 620)
        self.geometry("960x720")

        self._log_queue: queue.Queue[str] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._advanced_visible = False

        self.guidelines = GuidelinesIndex()
        self.navigator = OrgNavigator()
        self._market_matches: list = []
        self._current_subset = pd.DataFrame()
        self._dept_tree: dict = {}

        self._build_ui()
        self._setup_logging()
        self._load_saved_session()
        self._load_guidelines_default()
        self._load_org_default()
        self.after(100, self._poll_log_queue)

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 12, "pady": 4}
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_extract = ttk.Frame(self.notebook)
        self.tab_explore = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_extract, text="  Estrazione  ")
        self.notebook.add(self.tab_explore, text="  Esplora organigramma  ")

        self._build_extract_tab(pad)
        self._build_explore_tab(pad)

    def _build_extract_tab(self, pad: dict):
        header = ttk.Frame(self.tab_extract)
        header.pack(fill="x", **pad)
        ttk.Label(header, text="Estrazione organigramma", font=("Segoe UI", 15, "bold")).pack(
            anchor="w"
        )
        ttk.Label(
            header,
            text="Copia gli header da DevTools (Network → data) e incolla con un click.",
            wraplength=700,
        ).pack(anchor="w", pady=(2, 0))

        csrf_frame = ttk.LabelFrame(self.tab_extract, text="Token CSRF")
        csrf_frame.pack(fill="x", **pad)
        row = ttk.Frame(csrf_frame)
        row.pack(fill="x", padx=10, pady=8)
        self.csrf_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.csrf_var, font=("Consolas", 10)).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(row, text="Incolla da clipboard", command=self._paste_from_clipboard).pack(
            side="left", padx=(8, 0)
        )

        self.btn_row = ttk.Frame(self.tab_extract)
        self.btn_row.pack(fill="x", **pad)
        self.run_btn = ttk.Button(
            self.btn_row, text="Avvia estrazione", command=self._start_extraction
        )
        self.run_btn.pack(side="left")
        ttk.Button(self.btn_row, text="Apri cartella output", command=self._open_output).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(self.btn_row, text="Impostazioni SESSION", command=self._toggle_advanced).pack(
            side="right"
        )

        self.advanced_frame = ttk.LabelFrame(
            self.tab_extract, text="Cookie SESSION (prima volta o errore 401)"
        )
        self.session_var = tk.StringVar()
        ttk.Entry(
            self.advanced_frame, textvariable=self.session_var, font=("Consolas", 10)
        ).pack(fill="x", padx=10, pady=8)

        self.status_var = tk.StringVar(value="Pronto.")
        ttk.Label(self.tab_extract, textvariable=self.status_var).pack(anchor="w", padx=12)

        self.progress = ttk.Progressbar(self.tab_extract, mode="indeterminate")
        self.progress.pack(fill="x", **pad)

        log_frame = ttk.LabelFrame(self.tab_extract, text="Log")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=10, state="disabled", font=("Consolas", 9)
        )
        self.log_text.pack(fill="both", expand=True, padx=8, pady=8)

    def _build_explore_tab(self, pad: dict):
        top = ttk.Frame(self.tab_explore)
        top.pack(fill="x", **pad)

        self.guidelines_status = tk.StringVar(value="Guidelines: non caricato")
        ttk.Label(top, textvariable=self.guidelines_status).pack(side="left")
        ttk.Button(top, text="Ricarica guidelines", command=self._load_guidelines_default).pack(
            side="left", padx=8
        )
        ttk.Button(top, text="Scegli file…", command=self._pick_guidelines).pack(side="left")

        self.org_status = tk.StringVar(value="Organigramma: non caricato")
        ttk.Label(top, textvariable=self.org_status).pack(side="left", padx=(20, 0))
        ttk.Button(top, text="Carica Excel…", command=self._pick_org_excel).pack(
            side="left", padx=8
        )
        ttk.Button(top, text="Ultimo export", command=self._load_org_default).pack(side="left")

        search_frame = ttk.LabelFrame(
            self.tab_explore,
            text="Cerca (es. RN, Romania, «Romania Marketing» = tutte le parole nell'organigramma)",
        )
        search_frame.pack(fill="x", **pad)
        srow = ttk.Frame(search_frame)
        srow.pack(fill="x", padx=10, pady=8)
        self.search_var = tk.StringVar()
        ent = ttk.Entry(srow, textvariable=self.search_var, font=("Segoe UI", 11))
        ent.pack(side="left", fill="x", expand=True)
        ent.bind("<Return>", lambda _: self._run_search())
        ttk.Button(srow, text="Cerca", command=self._run_search).pack(side="left", padx=8)
        self.search_roles_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            search_frame,
            text="Includi anche ruolo / nome nella ricerca libera",
            variable=self.search_roles_var,
        ).pack(anchor="w", padx=10, pady=(0, 6))

        mid = ttk.Panedwindow(self.tab_explore, orient="horizontal")
        mid.pack(fill="both", expand=True, **pad)

        left = ttk.Frame(mid)
        right = ttk.Frame(mid)
        mid.add(left, weight=1)
        mid.add(right, weight=2)

        ttk.Label(left, text="Mercati trovati (guidelines)").pack(anchor="w")
        self.market_list = tk.Listbox(left, height=6, font=("Segoe UI", 10))
        self.market_list.pack(fill="both", expand=True, pady=4)
        self.market_list.bind("<<ListboxSelect>>", lambda _: self._on_market_select())

        self.breadcrumb_var = tk.StringVar(value="")
        ttk.Label(left, textvariable=self.breadcrumb_var, wraplength=280, foreground="#555").pack(
            anchor="w", pady=4
        )

        ttk.Label(left, text="Struttura dipartimenti").pack(anchor="w")
        self.dept_tree = ttk.Treeview(left, columns=("path",), displaycolumns=())
        self.dept_tree.pack(fill="both", expand=True, pady=4)
        self.dept_tree.bind("<<TreeviewSelect>>", lambda _: self._on_dept_select())

        ttk.Label(right, text="Persone (filtrate)").pack(anchor="w")
        cols = ("name", "country", "market", "role", "email", "dept")
        self.emp_tree = ttk.Treeview(right, columns=cols, show="headings", height=20)
        for c, w, label in [
            ("name", 130, "Nome"),
            ("country", 85, "Country"),
            ("market", 50, "Sigla"),
            ("role", 140, "Ruolo"),
            ("email", 160, "Email"),
            ("dept", 130, "Dipartimento"),
        ]:
            self.emp_tree.heading(c, text=label)
            self.emp_tree.column(c, width=w, minwidth=80)
        vsb = ttk.Scrollbar(right, orient="vertical", command=self.emp_tree.yview)
        self.emp_tree.configure(yscrollcommand=vsb.set)
        self.emp_tree.pack(side="left", fill="both", expand=True, pady=4)
        vsb.pack(side="left", fill="y", pady=4)

        self.explore_summary = tk.StringVar(value="")
        ttk.Label(self.tab_explore, textvariable=self.explore_summary).pack(anchor="w", padx=12)

    # ── shared ────────────────────────────────────────────────────────────────

    def _show_advanced(self):
        if not self._advanced_visible:
            self._advanced_visible = True
            self.advanced_frame.pack(fill="x", padx=12, pady=6, before=self.progress)

    def _toggle_advanced(self):
        if self._advanced_visible:
            self.advanced_frame.pack_forget()
            self._advanced_visible = False
        else:
            self._show_advanced()

    def _load_saved_session(self):
        data = _load_yaml(LOCAL_SESSION_PATH)
        if data.get("session_token"):
            self.session_var.set(data["session_token"])

    def _setup_logging(self):
        root = logging.getLogger()
        root.setLevel(logging.INFO)
        for h in root.handlers[:]:
            root.removeHandler(h)
        root.addHandler(QueueLogHandler(self._log_queue))

    def _append_log(self, line: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _poll_log_queue(self):
        while True:
            try:
                line = self._log_queue.get_nowait()
                self._append_log(line)
            except queue.Empty:
                break
        self.after(100, self._poll_log_queue)

    def _paste_from_clipboard(self):
        try:
            clip = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Clipboard vuota", "Nessun testo negli appunti.")
            return

        tokens = parse_clipboard_tokens(clip)
        filled = []
        if tokens["csrf_token"]:
            self.csrf_var.set(tokens["csrf_token"])
            filled.append("CSRF")
        if tokens["session_token"]:
            self.session_var.set(tokens["session_token"])
            filled.append("SESSION")

        if filled:
            self.status_var.set(f"Incollati da clipboard: {', '.join(filled)}")
            if tokens["session_token"]:
                self._show_advanced()
        else:
            messagebox.showwarning(
                "Nessun token riconosciuto",
                "Copia da DevTools la richiesta 'data' (Request Headers)\n"
                "oppure il valore di x-csrf-token / cookie SESSION.",
            )

    def _open_output(self):
        out = APP_DIR / "output"
        out.mkdir(exist_ok=True)
        if sys.platform == "win32":
            subprocess.run(["explorer", str(out)], check=False)
        elif sys.platform == "darwin":
            subprocess.run(["open", str(out)], check=False)
        else:
            subprocess.run(["xdg-open", str(out)], check=False)

    def _set_running(self, running: bool):
        self.run_btn.configure(state="disabled" if running else "normal")
        if running:
            self.progress.start(12)
        else:
            self.progress.stop()

    # ── extraction ────────────────────────────────────────────────────────────

    def _start_extraction(self):
        if self._worker and self._worker.is_alive():
            return

        csrf = self.csrf_var.get().strip()
        if not csrf:
            messagebox.showwarning("Token mancante", "Incolla il token CSRF prima di avviare.")
            return

        try:
            cfg = build_runtime_config(csrf, self.session_var.get().strip() or None)
        except ValueError as e:
            messagebox.showerror("Configurazione", str(e))
            self._show_advanced()
            return

        self._set_running(True)
        self.status_var.set("Estrazione in corso…")
        from orgchart_crawler import CRAWLER_VERSION
        self._append_log(f"——— Avvio estrazione (crawler v{CRAWLER_VERSION}) ———")

        def progress(calls: int, emp: int, q: int):
            self.after(
                0,
                lambda: self.status_var.set(
                    f"Chiamate: {calls}  |  Dipendenti: {emp}  |  Coda: {q}"
                ),
            )

        def work():
            try:
                df, path = run_extraction(cfg, progress_callback=progress)

                def done_ok():
                    self._set_running(False)
                    if df.empty or path is None:
                        self.status_var.set("Completato senza dipendenti.")
                        messagebox.showwarning(
                            "Nessun dato",
                            "0 dipendenti — verifica token o VPN.",
                        )
                    else:
                        self.status_var.set(f"Fatto — {df['user_id'].nunique()} dipendenti")
                        self.navigator.load_dataframe(
                            df, guidelines=self.guidelines if self.guidelines.loaded else None
                        )
                        self._update_org_status()
                        messagebox.showinfo(
                            "Completato",
                            f"Dipendenti: {df['user_id'].nunique()}\nFile:\n{path}\n\n"
                            "Vai al tab Esplora per navigare.",
                        )
                        self.notebook.select(self.tab_explore)

                self.after(0, done_ok)
            except SystemExit:
                def done_auth():
                    self._set_running(False)
                    self.status_var.set("Token scaduto.")
                    messagebox.showerror("Token scaduto", "401/403 — incolla token freschi.")
                    self._show_advanced()

                self.after(0, done_auth)
            except Exception as e:
                def done_err():
                    self._set_running(False)
                    self.status_var.set("Errore.")
                    messagebox.showerror("Errore", str(e))

                self.after(0, done_err)

        self._worker = threading.Thread(target=work, daemon=True)
        self._worker.start()

    # ── guidelines & org load ─────────────────────────────────────────────────

    def _load_guidelines_default(self):
        path = default_guidelines_path(APP_DIR)
        if path.exists():
            n = self.guidelines.load(path)
            self.guidelines_status.set(f"Guidelines: {path.name} ({n} mercati)")
            if self.navigator.loaded:
                self.navigator.enrich(self.guidelines)
        else:
            self.guidelines_status.set("Guidelines: file mancante (guidelines.xlsx)")

    def _pick_guidelines(self):
        path = filedialog.askopenfilename(
            title="Seleziona guidelines",
            filetypes=[("Excel", "*.xlsx"), ("Tutti", "*.*")],
            initialdir=str(APP_DIR),
        )
        if path:
            n = self.guidelines.load(Path(path))
            self.guidelines_status.set(f"Guidelines: {Path(path).name} ({n} mercati)")
            if self.navigator.loaded:
                self.navigator.enrich(self.guidelines)

    def _load_org_default(self):
        out = APP_DIR / "output"
        out.mkdir(exist_ok=True)
        latest = self.navigator.auto_load(
            out, guidelines=self.guidelines if self.guidelines.loaded else None
        )
        self._update_org_status()
        if latest:
            self.explore_summary.set(f"Caricato ultimo export: {latest.name}")

    def _pick_org_excel(self):
        path = filedialog.askopenfilename(
            title="Organigramma Excel",
            filetypes=[("Excel", "*.xlsx"), ("Tutti", "*.*")],
            initialdir=str(APP_DIR / "output"),
        )
        if path:
            self.navigator.load_excel(
                Path(path), guidelines=self.guidelines if self.guidelines.loaded else None
            )
            self._update_org_status()

    def _update_org_status(self):
        if self.navigator.loaded:
            src = self.navigator.source_path.name if self.navigator.source_path else "memoria"
            self.org_status.set(
                f"Organigramma: {self.navigator.employee_count} persone ({src})"
            )
        else:
            self.org_status.set("Organigramma: non caricato — estrai o carica Excel")

    # ── explore / search ──────────────────────────────────────────────────────

    def _run_search(self):
        query = self.search_var.get().strip()
        if not query:
            return
        if not self.navigator.loaded:
            messagebox.showinfo(
                "Nessun organigramma",
                "Esegui prima un'estrazione o carica un file Excel.",
            )
            return

        self.market_list.delete(0, "end")
        self._market_matches = []

        if self.guidelines.loaded:
            self._market_matches = self.guidelines.search(query)
            for m in self._market_matches:
                self.market_list.insert("end", m.label)

        if self._market_matches:
            self.market_list.selection_set(0)
            self._on_market_select()
        else:
            self._current_subset = self.navigator.filter_by_text(
                query, include_roles=self.search_roles_var.get()
            )
            self._rebuild_dept_tree()
            n = len(self._current_subset)
            u = self._current_subset["user_id"].nunique() if n else 0
            self.breadcrumb_var.set(f"Ricerca libera: «{query}»")
            self.explore_summary.set(
                f"{u} persone, {n} righe — filtro organigramma"
                + (" + ruolo/nome" if self.search_roles_var.get() else "")
            )
            self._show_all_employees()

    def _on_market_select(self):
        sel = self.market_list.curselection()
        if not sel or not self._market_matches:
            return
        match = self._market_matches[sel[0]]
        self.breadcrumb_var.set(match.breadcrumb)
        self._current_subset = self.navigator.filter_by_market(match)
        self._rebuild_dept_tree()
        n = len(self._current_subset)
        u = self._current_subset["user_id"].nunique() if n else 0
        self.explore_summary.set(
            f"{match.label}: {u} persone uniche, {n} righe (posizioni)"
        )
        self._show_all_employees()

    def _rebuild_dept_tree(self):
        for item in self.dept_tree.get_children():
            self.dept_tree.delete(item)
        self._dept_tree = self.navigator.build_dept_tree(self._current_subset)
        tree_to_tk(self._dept_tree, "", self.dept_tree)
        for child in self.dept_tree.get_children():
            self.dept_tree.item(child, open=True)

    def _on_dept_select(self):
        sel = self.dept_tree.selection()
        if not sel:
            return
        item = sel[0]
        dept_name = self.dept_tree.item(item, "text")
        self._fill_employee_table(dept_name)

    def _show_all_employees(self):
        self._fill_employee_table(None)

    def _fill_employee_table(self, dept_name: str | None):
        for row in self.emp_tree.get_children():
            self.emp_tree.delete(row)

        if dept_name:
            rows = _find_employees(self._dept_tree, dept_name)
        else:
            rows = []
            if not self._current_subset.empty:
                for _, r in self._current_subset.iterrows():
                    rows.append(_employee_display_row(r))

        seen = set()
        for e in rows:
            key = (e.get("full_name"), e.get("email"))
            if key in seen:
                continue
            seen.add(key)
            self.emp_tree.insert(
                "",
                "end",
                values=(
                    e.get("full_name", ""),
                    e.get("country", ""),
                    e.get("market", ""),
                    e.get("job_title", ""),
                    e.get("email", ""),
                    e.get("department_name", ""),
                ),
            )


def _employee_display_row(r) -> dict:
    return {
        "full_name": r.get("full_name", ""),
        "job_title": r.get("job_title", ""),
        "email": r.get("email", ""),
        "department_name": r.get("department_name", ""),
        "country": r.get("country", ""),
        "market": r.get("market", ""),
    }


def _find_employees(tree: dict, dept_name: str) -> list[dict]:
    if dept_name in tree:
        node = tree[dept_name]
        out = list(node.get("_employees", []))
        for k, v in node.items():
            if k != "_employees":
                out.extend(_collect_employees(v))
        return out
    for k, v in tree.items():
        if k == "_employees":
            continue
        if k == dept_name:
            return _collect_employees(v)
        found = _find_employees(v, dept_name)
        if found:
            return found
    return []


def _collect_employees(node: dict) -> list[dict]:
    out = list(node.get("_employees", []))
    for k, v in node.items():
        if k != "_employees":
            out.extend(_collect_employees(v))
    return out


def main():
    app = OrgChartApp()
    app.mainloop()


if __name__ == "__main__":
    main()
