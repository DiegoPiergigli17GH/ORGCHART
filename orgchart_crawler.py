"""
SAP SuccessFactors / Ingentis Org Chart — Crawler v11.0

In molte prospettive Ingentis (es. Ariston) i dipendenti NON compaiono come
nodi EmpJob nei children: sono in Position.data._IOM_INTERNAL_.calculations
(es. EmpJobFullNamePOS, CountEmpJobsperPosition).

Crawl flow:
  1. BFS FODepartment (downLinks come il browser) → dept + Position inline
  2. Estrae dipendenti da nodi Position (calculations + eventuali EmpJob figli)
  3. BFS Position_to_Position per la gerarchia
  4. Batch Position opzionale per arricchire con userId da EmpJob
  5. Export Excel

Setup: copy config.example.yaml → config.yaml and fill SESSION / CSRF from DevTools.
"""

CRAWLER_VERSION = "11.0"

import base64
import json
import logging
import re
import sys
import time
import unicodedata
from collections import deque
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

import pandas as pd
import requests
import yaml

try:
    from tqdm import tqdm
except ImportError:
    class tqdm:
        def __init__(self, *_, **kw):
            self._desc = kw.get("desc", "")
            self._n = 0

        def __enter__(self):
            return self

        def __exit__(self, *_):
            print()

        def update(self, n=1):
            self._n += n
            print(f"\r  {self._desc}: {self._n} …", end="", flush=True)

        def set_postfix(self, d):
            print(
                f"\r  {self._desc}: {self._n}  "
                f"[{', '.join(f'{k}={v}' for k, v in d.items())}]  ",
                end="",
                flush=True,
            )


# ── paths (works as script or PyInstaller .exe) ───────────────────────────────

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
else:
    APP_DIR = Path(__file__).resolve().parent
    BUNDLE_DIR = APP_DIR

CONFIG_PATH = APP_DIR / "config.yaml"
LOCAL_SESSION_PATH = APP_DIR / "local_session.yaml"


def _defaults_path() -> Path:
    local = APP_DIR / "orgchart_defaults.yaml"
    if local.exists():
        return local
    bundled = BUNDLE_DIR / "orgchart_defaults.yaml"
    return bundled if bundled.exists() else local


# ── config ────────────────────────────────────────────────────────────────────

def _load_yaml(path: Path) -> dict:
    if not path.exists():
        return {}
    with path.open(encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def parse_csrf(text: str) -> str:
    """Estrae x-csrf-token da incolla grezzo DevTools."""
    text = text.strip().strip('"').strip("'")
    if not text:
        return ""
    m = re.search(r"x-csrf-token[:\s]+([^\s;,\r\n]+)", text, re.I)
    if m:
        return m.group(1)
    return text.split()[0]


def parse_session(text: str) -> str:
    """Estrae cookie SESSION da incolla grezzo DevTools."""
    text = text.strip().strip('"').strip("'")
    if not text:
        return ""
    m = re.search(r"SESSION[=:\s]+([a-f0-9-]{30,})", text, re.I)
    if m:
        return m.group(1)
    if re.match(r"^[a-f0-9-]{30,}$", text, re.I):
        return text
    return text.split()[0]


def parse_clipboard_tokens(text: str) -> dict[str, str]:
    """Legge CSRF e SESSION da incolla DevTools (header, cookie o valori singoli)."""
    return {
        "csrf_token": parse_csrf(text),
        "session_token": parse_session(text),
    }


def save_local_session(session_token: str) -> None:
    LOCAL_SESSION_PATH.write_text(
        yaml.safe_dump({"session_token": session_token}, allow_unicode=True),
        encoding="utf-8",
    )


def build_runtime_config(
    csrf_token: str,
    session_token: Optional[str] = None,
) -> dict:
    """Unisce defaults Ariston + session salvata + CSRF incollato ora."""
    cfg = _load_yaml(_defaults_path())
    cfg.update(_load_yaml(CONFIG_PATH))
    cfg.update(_load_yaml(LOCAL_SESSION_PATH))

    csrf = parse_csrf(csrf_token)
    if not csrf:
        raise ValueError("Token CSRF mancante.")

    session = parse_session(session_token) if session_token else ""
    if not session:
        session = parse_session(str(cfg.get("session_token", "")))

    if not session or "YOUR" in session:
        raise ValueError(
            "Sessione mancante. Apri Impostazioni e incolla il cookie SESSION "
            "(solo la prima volta o se ricevi errore 401)."
        )

    cfg["csrf_token"] = csrf
    cfg["session_token"] = session
    save_local_session(session)
    return cfg


def load_config() -> dict:
    """CLI: legge config.yaml completo (modalità legacy)."""
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Missing {CONFIG_PATH.name}\n"
            f"  Usa orgchart_app.py (GUI) oppure copia config.example.yaml → config.yaml"
        )
    cfg = _load_yaml(_defaults_path())
    cfg.update(_load_yaml(CONFIG_PATH))

    required = ["base_url", "session_token", "csrf_token", "root_node_id"]
    missing = [k for k in required if not cfg.get(k) or "YOUR" in str(cfg.get(k, ""))]
    if missing:
        raise ValueError(f"config.yaml — fill in: {', '.join(missing)}")

    return cfg


def _cfg(cfg: dict, key: str, default: Any = None) -> Any:
    return cfg.get(key, default)


# ── logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ── utilities ─────────────────────────────────────────────────────────────────

def _b64decode(payload: str) -> str:
    pad = "=" * (4 - len(payload) % 4) if len(payload) % 4 else ""
    try:
        d = base64.b64decode(payload + pad).decode("utf-8")
        return d if re.match(r"^\d+$", d) else payload
    except Exception:
        return payload


def parse_node_id(nid: str) -> list[dict]:
    segs = []
    for part in nid.split("."):
        m = re.match(r"^(.+?)::(.+?)(?::(\d+))?$", part)
        if m:
            segs.append({
                "type": m.group(1),
                "raw": m.group(2),
                "value": _b64decode(m.group(2)),
            })
    return segs


def node_to_path(nid: str) -> str:
    return " > ".join(f"{s['type']}/{s['value']}" for s in parse_node_id(nid))


def dept_id_from_node(nid: str) -> str:
    segs = [s for s in parse_node_id(nid) if s["type"] == "FODepartment"]
    return segs[-1]["value"] if segs else ""


def dept_name_from_data(d: dict) -> str:
    return (
        d.get("name_defaultValue")
        or d.get("name_en_defaultValue")
        or d.get("externalCode")
        or ""
    )


def parse_name(name_f: str, caps_f: str) -> tuple[str, str]:
    if not name_f:
        return "", ""
    clean = name_f.strip().title()
    caps = caps_f.strip().title() if caps_f else ""
    parts = clean.split()
    if len(parts) == 1:
        return parts[0], ""
    if len(parts) == 2:
        return parts[0], parts[1]
    if caps:
        for i, p in enumerate(parts):
            if p.lower() == caps.split()[0].lower():
                return (" ".join(parts[:i]) or parts[0]), " ".join(parts[i:])
    return parts[0], " ".join(parts[1:])


def _norm(s: str) -> str:
    n = unicodedata.normalize("NFKD", s)
    return re.sub(r"[^a-z0-9]", "", n.encode("ascii", "ignore").decode().lower())


def build_email(first: str, last: str, domain: str) -> str:
    f = _norm(first.split()[0]) if first else ""
    l = _norm(last.split()[0]) if last else ""
    if f and l:
        return f"{f}.{l}{domain}"
    if f:
        return f"{f}{domain}"
    return f"{l}{domain}" if l else ""


# ── session ─────────────────────────────────────────────────────────────────

DOWN_LINKS_BY_TYPE = {
    "FODepartment": ["FODepartment_to_FODepartment", "ROOTLINK"],
    "FODepartment_positions": ["FODepartment_to_Position"],
    "Position": ["Position_to_EmpJob", "Position_to_Position"],
}

# DownLinks usati dal browser Ariston su expand FODepartment (orgchart_defaults.yaml)
DEFAULT_DEPARTMENT_DOWN_LINKS = [
    "FODepartment_to_FODepartment",
    "FODepartment_to_Position",
    "FODepartment_to_ManagerPosition_to_Matrix",
    "PositionMatrix",
    "ROOTLINK",
]


def _department_down_links(cfg: dict) -> list[str]:
    links = _cfg(cfg, "department_down_links")
    return links if links else DEFAULT_DEPARTMENT_DOWN_LINKS


def dept_level(node_id: str) -> int:
    return node_id.count("FODepartment::")


def position_level(node_id: str) -> int:
    return node_id.count("Position::")


def _chunks(items: list, size: int):
    for i in range(0, len(items), size):
        yield items[i : i + size]


def build_session(cfg: dict) -> requests.Session:
    base_url = _cfg(cfg, "base_url").rstrip("/")
    domain = base_url.replace("https://", "").replace("http://", "").split("/")[0]
    s = requests.Session()
    s.cookies.set("SESSION", _cfg(cfg, "session_token"), domain=domain)
    s.headers.update({
        "Content-Type": "application/json",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "x-csrf-token": _cfg(cfg, "csrf_token"),
        "x-asofdate": datetime.now().strftime("%Y-%m-%d"),
        "x-dataset-type": "preloaded",
        "x-requested-with": "XMLHttpRequest",
        "Origin": base_url,
        "Referer": f"{base_url}/public/orghtml/index.html?idp=sso",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
    })
    log.info("Session → %s/public/api/data", base_url)
    return s


# ── API ───────────────────────────────────────────────────────────────────────

def _retry_pause(cfg: dict, attempt: int, extra: float = 0) -> None:
    """Pausa tra retry — più lunga dopo reset connessione (rate limit / WAF)."""
    base = float(_cfg(cfg, "retry_backoff", 4))
    time.sleep(base * attempt + extra)


def _post(
    session: requests.Session,
    cfg: dict,
    body: dict,
    label: str = "",
    save_num: int = 0,
) -> Optional[dict]:
    base_url = _cfg(cfg, "base_url").rstrip("/")
    url = f"{base_url}/public/api/data"
    max_retries = int(_cfg(cfg, "max_retries", 8))
    timeout = int(_cfg(cfg, "timeout", 60))
    debug_raw = bool(_cfg(cfg, "debug_raw", False))
    debug_dir = APP_DIR / "debug"

    for attempt in range(1, max_retries + 1):
        try:
            r = session.post(url, json=body, timeout=timeout)

            if r.status_code == 401:
                log.error("401 — SESSION scaduto. Aggiorna session_token in config.yaml.")
                sys.exit(1)
            if r.status_code == 403:
                log.error("403 — CSRF non valido. Aggiorna csrf_token in config.yaml.")
                sys.exit(1)
            if r.status_code == 429:
                wait = 8 * attempt
                log.warning("Rate limit HTTP 429 — attendo %ds", wait)
                time.sleep(wait)
                continue
            if r.status_code != 200:
                log.warning(
                    "HTTP %d attempt %d/%d — body: %s",
                    r.status_code, attempt, max_retries,
                    (r.text[:120] + "…") if r.text else "(vuoto)",
                )
                _retry_pause(cfg, attempt)
                continue

            try:
                data = r.json()
            except json.JSONDecodeError:
                snippet = (r.text[:150] + "…") if r.text else "(risposta vuota)"
                log.warning(
                    "Non-JSON (spesso HTML/WAF) attempt %d/%d: %s",
                    attempt, max_retries, snippet,
                )
                _retry_pause(cfg, attempt, extra=6)
                continue

            if debug_raw and save_num > 0 and save_num <= 12:
                debug_dir.mkdir(parents=True, exist_ok=True)
                fname = debug_dir / f"raw_{save_num:02d}_{label}.json"
                fname.write_text(json.dumps(data, indent=2, ensure_ascii=False))
                log.info("DEBUG → %s  (%d items)", fname.name, len(data.get("data", [])))

            return data

        except requests.exceptions.Timeout:
            log.warning("Timeout attempt %d/%d", attempt, max_retries)
            _retry_pause(cfg, attempt)
        except requests.exceptions.ConnectionError as e:
            log.warning(
                "Connessione interrotta dal server attempt %d/%d — %s",
                attempt, max_retries, e,
            )
            _retry_pause(cfg, attempt, extra=10)
        except requests.exceptions.RequestException as e:
            log.warning("Errore rete attempt %d/%d — %s", attempt, max_retries, e)
            _retry_pause(cfg, attempt, extra=5)

    return None


def expand_node(
    session: requests.Session,
    cfg: dict,
    node_id: str,
    node_type: str,
    level: int,
    call_num: int = 0,
    mode: str = "dept",
) -> Optional[dict]:
    if mode in ("dept", "dept_pos"):
        api_type = "FODepartment"
        down_links = (
            DOWN_LINKS_BY_TYPE["FODepartment_positions"]
            if mode == "dept_pos"
            else _department_down_links(cfg)
        )
    else:
        api_type = node_type
        down_links = DOWN_LINKS_BY_TYPE.get(
            node_type, ["FODepartment_to_FODepartment", "ROOTLINK"]
        )

    body = {
        "perspective": int(_cfg(cfg, "perspective", 4928)),
        "chart": str(_cfg(cfg, "chart_id", "4925")),
        "type": api_type,
        "level": level,
        "runId": 0,
        "asOfDate": datetime.now().strftime("%Y-%m-%d"),
        "downLinks": down_links,
        "ids": [node_id],
        "isListRequest": False,
        "liveConditions": [],
    }
    label = f"EXPAND_{mode}_{api_type}_lvl{level}"
    return _post(session, cfg, body, label=label, save_num=call_num)


def fetch_position_batch(
    session: requests.Session,
    cfg: dict,
    position_ids: list[str],
    level: int,
    call_num: int = 0,
) -> Optional[dict]:
    """Batch Position → EmpJob (isListRequest=true, come DevTools)."""
    body = {
        "perspective": int(_cfg(cfg, "perspective", 4928)),
        "chart": str(_cfg(cfg, "chart_id", "4925")),
        "type": "Position",
        "level": level,
        "runId": 0,
        "asOfDate": datetime.now().strftime("%Y-%m-%d"),
        "downLinks": ["Position_to_EmpJob"],
        "ids": position_ids,
        "isListRequest": True,
        "liveConditions": [],
    }
    label = f"BATCH_Position_lvl{level}_n{len(position_ids)}"
    return _post(session, cfg, body, label=label, save_num=call_num)


# ── parser ────────────────────────────────────────────────────────────────────

def _flatten(items: list) -> list:
    """Flatten nested FODepartment / Position trees."""
    flat = []
    for item in items:
        flat.append(item)
        if item.get("type") in ("FODepartment", "Position"):
            for child in item.get("children", []):
                flat.extend(_flatten([child]))
    return flat


def _process_item(
    item: dict,
    dept_names: dict,
    manager: str,
    to_enqueue: list,
    employees: list,
    email_domain: str,
    collected_positions: set[str],
    seen_employees: set[tuple[str, str]],
) -> None:
    itype = item.get("type", "")
    iid = item.get("id", "")
    idata = item.get("data", {})
    counts = item.get("childCounts", {})
    children = item.get("children", [])

    if itype == "FODepartment":
        name = dept_name_from_data(idata)
        dept_val = dept_id_from_node(iid)
        if name and dept_val:
            dept_names[dept_val] = name

        for child in children:
            _process_item(
                child, dept_names, manager, to_enqueue, employees, email_domain,
                collected_positions, seen_employees,
            )

        n_sub = counts.get("FODepartment_to_FODepartment", 0)
        if n_sub > 0 and not children:
            to_enqueue.append((iid, "FODepartment", "dept", manager))

    elif itype == "Position":
        if "Position::" in iid:
            collected_positions.add(iid)

        pos_code = idata.get("code", "")
        job_title = idata.get("jobTitle") or idata.get("externalName_defaultValue", "")
        dept_id = dept_id_from_node(iid)
        h_path = node_to_path(iid)

        pos_uids = []
        for child in children:
            if child.get("type") == "EmpJob":
                emp = _extract_emp(
                    child, pos_code, job_title, dept_id, h_path, manager, email_domain
                )
                if emp:
                    key = (emp["user_id"], emp["node_id"])
                    if key not in seen_employees:
                        seen_employees.add(key)
                        employees.append(emp)
                    pos_uids.append(emp["user_id"])
            elif child.get("type") == "Position":
                mgr = pos_uids[0] if pos_uids else manager
                _process_item(
                    child, dept_names, mgr, to_enqueue, employees, email_domain,
                    collected_positions, seen_employees,
                )

        mgr = pos_uids[0] if pos_uids else manager

        if not pos_uids:
            emp = _extract_emp_from_position(
                iid, idata, pos_code, job_title, dept_id, h_path, mgr, email_domain
            )
            if emp:
                key = (emp["user_id"], iid)
                if key not in seen_employees:
                    seen_employees.add(key)
                    employees.append(emp)
                    pos_uids.append(emp["user_id"])

        has_emp_children = any(c.get("type") == "EmpJob" for c in children)
        has_pos_children = any(c.get("type") == "Position" for c in children)

        if counts.get("Position_to_Position", 0) > 0 and not has_pos_children:
            to_enqueue.append((iid, "Position", "pos", mgr))

        if counts.get("Position_to_EmpJob", 0) > 0 and not has_emp_children:
            collected_positions.add(iid)


def _get_calculations(idata: dict) -> dict:
    """Campi calcolati Ingentis (chiave con o senza underscore iniziale)."""
    for key in ("_IOM_INTERNAL_", "IOM_INTERNAL"):
        block = idata.get(key, {})
        if isinstance(block, dict):
            calcs = block.get("calculations", {})
            if isinstance(calcs, dict) and calcs:
                return calcs
    calcs = idata.get("calculations", {})
    return calcs if isinstance(calcs, dict) else {}


def _calc_str(calcs: dict, *keys: str) -> str:
    for key in keys:
        val = calcs.get(key)
        if val is not None and str(val).strip():
            return str(val).strip()
    return ""


def _position_employee_count(calcs: dict) -> int:
    raw = _calc_str(
        calcs,
        "CountEmpJobsperPosition",
        "CountEmpJobsPerPosition",
        "CountEmpJobs",
    )
    if not raw:
        return 0
    try:
        return int(float(raw))
    except ValueError:
        return 0


def _extract_emp_from_position(
    position_id: str,
    idata: dict,
    pos_code: str,
    job_title: str,
    dept_id: str,
    h_path: str,
    manager: str,
    email_domain: str,
) -> Optional[dict]:
    """
    Dipendente embedded nella Position (prospettiva senza nodi EmpJob espansi).
    """
    calcs = _get_calculations(idata)
    full_name = _calc_str(
        calcs,
        "EmpJobFullNamePOS",
        "Name",
        "FullName",
        "EmployeeName",
    )
    emp_count = _position_employee_count(calcs)
    if not full_name and emp_count <= 0:
        return None
    if not full_name:
        return None

    first, last = parse_name(full_name, _calc_str(calcs, "FullNameAllCaps", "NameAllCaps"))
    uid = _calc_str(
        calcs,
        "EmpJobUserIdPOS",
        "UserIdPOS",
        "userId",
        "EmpJobUserId",
    ) or str(idata.get("userId") or "").strip()
    if not uid:
        uid = f"pos:{pos_code}" if pos_code else f"posnode:{position_id.rsplit('Position::', 1)[-1]}"

    status = _calc_str(calcs, "EmpJobStatusPOS", "emplStatus", "EmplStatus")
    title = job_title or _calc_str(calcs, "EmpJobTitlePOS", "JobTitle", "externalName_defaultValue")

    return {
        "user_id": uid,
        "first_name": first,
        "last_name": last,
        "full_name": full_name.title(),
        "job_title": title,
        "position_id": pos_code,
        "department_id": dept_id,
        "manager_user_id": manager,
        "empl_status": status,
        "email": build_email(first, last, email_domain),
        "hierarchy_path": h_path,
        "node_id": position_id,
        "data_source": "position_calculations",
    }


def _extract_emp(
    node: dict,
    pos_code: str,
    job_title: str,
    dept_id: str,
    h_path: str,
    manager: str,
    email_domain: str,
) -> Optional[dict]:
    d = node.get("data", {})
    calcs = d.get("_IOM_INTERNAL_", {}).get("calculations", {})
    uid = d.get("userId", "")
    if not uid:
        return None
    first, last = parse_name(calcs.get("Name", ""), calcs.get("FullNameAllCaps", ""))
    return {
        "user_id": uid,
        "first_name": first,
        "last_name": last,
        "full_name": calcs.get("Name", "").title(),
        "job_title": job_title,
        "position_id": pos_code,
        "department_id": dept_id,
        "manager_user_id": manager,
        "empl_status": d.get("emplStatus", ""),
        "email": build_email(first, last, email_domain),
        "hierarchy_path": h_path,
        "node_id": node.get("id", ""),
    }


def process_response(
    response: dict,
    dept_names: dict,
    inherited_manager: str,
    email_domain: str,
    collected_positions: set[str],
    seen_employees: set[tuple[str, str]],
) -> tuple[list[dict], list[tuple]]:
    employees: list[dict] = []
    to_enqueue: list[tuple] = []

    for item in _flatten(response.get("data", [])):
        _process_item(
            item, dept_names, inherited_manager, to_enqueue, employees, email_domain,
            collected_positions, seen_employees,
        )

    return employees, to_enqueue


# ── crawler ───────────────────────────────────────────────────────────────────

class _SilentBar:
    def update(self, n=1):
        pass

    def set_postfix(self, d):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *_):
        pass


def crawl(
    session: requests.Session,
    cfg: dict,
    progress_callback: Optional[Callable[[int, int, int], None]] = None,
) -> tuple[list[dict], dict]:
    """
    BFS con visited (node_id, node_type, mode).
    Fase finale: batch Position ids con isListRequest=true.
    """
    visited: set[tuple[str, str, str]] = set()
    employees: list[dict] = []
    dept_names: dict = {}
    collected_positions: set[str] = set()
    seen_employees: set[tuple[str, str]] = set()
    call_num = 0
    empty_calls = 0

    root_id = _cfg(cfg, "root_node_id")
    email_domain = _cfg(cfg, "email_domain", "@company.com")
    delay = float(_cfg(cfg, "request_delay", 0.4))
    batch_size = int(_cfg(cfg, "position_batch_size", 20))

    queue: deque = deque([(root_id, "FODepartment", "dept", "")])
    log.info("Avvio crawl v%s root=%s", CRAWLER_VERSION, root_id)

    pbar_ctx = _SilentBar() if progress_callback else tqdm(desc="Chiamate API", unit=" calls")
    with pbar_ctx as pbar:
        fail_streak = 0

        while queue:
            node_id, ntype, mode, manager = queue.popleft()
            visit_key = (node_id, ntype, mode)

            if visit_key in visited:
                continue

            call_num += 1
            if mode == "pos":
                level = position_level(node_id)
            else:
                level = dept_level(node_id)

            resp = expand_node(
                session, cfg, node_id, ntype, level, call_num, mode=mode
            )
            pbar.update(1)

            if resp is None:
                queue.append((node_id, ntype, mode, manager))
                fail_streak += 1
                pause = delay * min(5 + fail_streak, 20)
                log.warning(
                    "Chiamata fallita — riaccodo nodo (pausa %.0fs, coda: %d)",
                    pause, len(queue),
                )
                pbar.set_postfix({
                    "emp": len(employees),
                    "calls": call_num,
                    "q": len(queue),
                })
                if progress_callback:
                    progress_callback(call_num, len(employees), len(queue))
                time.sleep(pause)
                continue

            fail_streak = 0
            visited.add(visit_key)

            if not resp.get("data"):
                empty_calls += 1
                log.debug("Vuoto — %s %s lvl=%d", mode, ntype, level)
                pbar.set_postfix({
                    "emp": len(employees),
                    "calls": call_num,
                    "q": len(queue),
                })
                time.sleep(delay)
                continue

            items = resp["data"]
            tc: dict[str, int] = {}
            for it in items:
                t = it.get("type", "?")
                tc[t] = tc.get(t, 0) + 1

            log.info(
                "Call %3d | %-8s %-12s lvl=%d → %d items %s",
                call_num, mode, ntype, level, len(items), tc,
            )

            new_employees, to_enqueue = process_response(
                resp, dept_names, manager, email_domain,
                collected_positions, seen_employees,
            )
            employees.extend(new_employees)

            if new_employees:
                log.info("  +%d dipendenti (tot: %d)", len(new_employees), len(employees))

            added = 0
            for nid, nt, md, mgr in to_enqueue:
                if (nid, nt, md) not in visited:
                    queue.append((nid, nt, md, mgr))
                    added += 1
            if added:
                log.info("  +%d nodi accodati (coda: %d)", added, len(queue))

            stats = {
                "emp": len(employees),
                "calls": call_num,
                "q": len(queue),
                "pos": len(collected_positions),
            }
            pbar.set_postfix(stats)
            if progress_callback:
                progress_callback(call_num, len(employees), len(queue))
            time.sleep(delay)

        # Fase batch: Position ids noti → EmpJob (come il browser)
        if collected_positions:
            log.info(
                "Fase batch — %d Position da interrogare (batch_size=%d)",
                len(collected_positions), batch_size,
            )
            by_level: dict[int, list[str]] = {}
            for pid in sorted(collected_positions):
                by_level.setdefault(position_level(pid), []).append(pid)

            for lv, ids in sorted(by_level.items()):
                for chunk in _chunks(ids, batch_size):
                    call_num += 1
                    resp = fetch_position_batch(session, cfg, chunk, lv, call_num)
                    pbar.update(1)

                    if resp is None:
                        log.warning("Batch fallito — %d Position lvl=%d", len(chunk), lv)
                        time.sleep(delay * 3)
                        continue

                    if not resp.get("data"):
                        empty_calls += 1
                        time.sleep(delay)
                        continue

                    tc = {}
                    for it in resp["data"]:
                        t = it.get("type", "?")
                        tc[t] = tc.get(t, 0) + 1
                    log.info(
                        "Call %3d | batch     Position     lvl=%d → %d items %s",
                        call_num, lv, len(resp["data"]), tc,
                    )

                    new_employees, _ = process_response(
                        resp, dept_names, "", email_domain,
                        collected_positions, seen_employees,
                    )
                    employees.extend(new_employees)
                    if new_employees:
                        log.info(
                            "  +%d dipendenti batch (tot: %d)",
                            len(new_employees), len(employees),
                        )

                    stats = {"emp": len(employees), "calls": call_num, "pos": len(collected_positions)}
                    pbar.set_postfix(stats)
                    if progress_callback:
                        progress_callback(call_num, len(employees), 0)
                    time.sleep(delay)

    log.info(
        "Fine — calls:%d  dipendenti:%d  dept:%d  position:%d  vuote:%d",
        call_num, len(employees), len(dept_names),
        len(collected_positions), empty_calls,
    )
    return employees, dept_names


def run_extraction(
    cfg: dict,
    progress_callback: Optional[Callable[[int, int, int], None]] = None,
) -> tuple[pd.DataFrame, Optional[Path]]:
    """Esegue crawl completo e export Excel. Usato da GUI e CLI."""
    session = build_session(cfg)
    employees, dept_names = crawl(session, cfg, progress_callback=progress_callback)

    if not employees:
        log.error("Nessun dipendente estratto.")
        return pd.DataFrame(), None

    df = build_dataframe(employees, dept_names)
    path = export_excel(df, dept_names, cfg)
    return df, path


# ── export ────────────────────────────────────────────────────────────────────

def build_dataframe(employees: list[dict], dept_names: dict) -> pd.DataFrame:
    if not employees:
        return pd.DataFrame()
    df = pd.DataFrame(employees)
    uid_to_name = df.set_index("user_id")["full_name"].to_dict()
    df["manager_full_name"] = df["manager_user_id"].map(uid_to_name).fillna("")
    df["department_name"] = df["department_id"].map(dept_names).fillna("")
    df["multi_position"] = df.duplicated(subset=["user_id"], keep=False)
    cols = [
        "user_id", "first_name", "last_name", "full_name", "email",
        "job_title", "position_id", "department_id", "department_name",
        "manager_user_id", "manager_full_name", "empl_status",
        "hierarchy_path", "multi_position", "node_id",
    ]
    df = df[[c for c in cols if c in df.columns]]
    df = df.sort_values(
        ["hierarchy_path", "last_name", "first_name"]
    ).reset_index(drop=True)
    log.info(
        "DataFrame: %d righe | %d unici | %d dept",
        len(df), df["user_id"].nunique(), df["department_id"].nunique(),
    )
    return df


def export_excel(df: pd.DataFrame, dept_names: dict, cfg: dict) -> Path:
    out = Path(_cfg(cfg, "output_dir", "."))
    if not out.is_absolute():
        out = APP_DIR / out
    out.mkdir(parents=True, exist_ok=True)

    company = _cfg(cfg, "company_name", "Company")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = out / f"orgchart_{company}_{ts}.xlsx"
    as_of = datetime.now().strftime("%Y-%m-%d")

    df_u = df.drop_duplicates(subset=["user_id"], keep="first").copy()
    df_d = pd.DataFrame([
        {"department_id": k, "department_name": v}
        for k, v in sorted(dept_names.items())
    ])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet, frame in [
            ("All_Employees", df),
            ("Unique_Employees", df_u),
            ("Departments", df_d),
        ]:
            if frame.empty:
                continue
            frame.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.freeze_panes = "A2"
            for i, col in enumerate(frame.columns, 1):
                w = max(
                    len(str(col)),
                    frame[col].astype(str).str.len().max() if len(frame) else 0,
                ) + 2
                ws.column_dimensions[ws.cell(1, i).column_letter].width = min(w, 65)

        pd.DataFrame({
            "Metrica": [
                "Righe totali", "Dipendenti unici", "Multi-posizione",
                "Dipartimenti", "Posizioni", "Data", "Timestamp",
            ],
            "Valore": [
                len(df),
                df["user_id"].nunique(),
                int(df[df["multi_position"]]["user_id"].nunique())
                if "multi_position" in df.columns else 0,
                df["department_id"].nunique(),
                df["position_id"].nunique(),
                as_of,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ],
        }).to_excel(writer, sheet_name="Riepilogo", index=False)

    log.info("Excel → %s", path)
    return path


# ── main ──────────────────────────────────────────────────────────────────────

def main() -> pd.DataFrame:
    print("=" * 70)
    print(f"  SAP SuccessFactors / Ingentis Org Chart Extractor  v{CRAWLER_VERSION}")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    cfg = load_config()
    df, path = run_extraction(cfg)

    if df.empty:
        log.error("Controlla debug/raw_*.json (se debug_raw: true) e i token.")
        return pd.DataFrame()

    print("\n" + "=" * 70)
    print(f"  ✓  Dipendenti totali  : {len(df)}")
    print(f"  ✓  Dipendenti unici   : {df['user_id'].nunique()}")
    print(f"  ✓  Dipartimenti       : {df['department_id'].nunique()}")
    print(f"  ✓  Excel              : {path}")
    print("=" * 70)
    return df


if __name__ == "__main__":
    main()
