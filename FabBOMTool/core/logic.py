"""
core/logic.py
All business logic preserved exactly from the original app.py.
Zero Tkinter dependencies. Pure Python.
"""

import re
import json
import csv
from pathlib import Path

import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from werkzeug.security import check_password_hash, generate_password_hash

# Embedded legend (optional - place legend_embedded.py next to this file or in project root)
try:
    import legend_embedded  # type: ignore
except Exception:
    legend_embedded = None

APP_NAME    = "Fabrication BOM Tool"
APP_VERSION = "2026.3.3"
DEFAULT_ADMIN_PASSWORD = "FBT2026!"

# ============================
# Paths
# ============================

def get_app_dir() -> Path:
    return Path(__file__).resolve().parent.parent

APP_DIR                      = get_app_dir()
DATA_DIR                     = APP_DIR / "data"
LEGEND_CACHE_PATH            = DATA_DIR / "Legend.cache.json"
LEGEND_XLSX_PATH             = APP_DIR / "Legend.xlsx"
SETTINGS_FILE                = DATA_DIR / "settings.json"
COMPANY_DEFAULTS_LOCAL_PATH  = DATA_DIR / "CompanyDefaults.json"
EXPORTS_DIR                  = APP_DIR / "exports"

# ============================
# Material Types preset
# ============================

MATERIAL_TYPE_PRESET_RAW = [
    "PVDF","Cast Iron","Forged Steel","Carbon Steel","Malleable Iron","Copper",
    "Bronze","PVC","Stainless Steel","Brass","Cast Steel","Ductile Iron",
    "Polypropylene","PVC SDR35","Polyethylene","PEX","Victaulic (OEM) Imperial",
    "304 Stainless Steel","Cast Brass","Aluminum","Nickel iron","CPVC",
    "Rubber","FPM","Copper B280","Epoxy Steel/Polyethylene","Forged Brass",
    "Steel","Neoprene","Engineered Polymer","Cast Copper","Cast Bronze",
    "Polypropelene","ABS","MDPE","HDPE","Forged Bronze","Black Iron",
    "Forged Carbon Steel","Cast Carbon Steel",
]


def dedupe_case_insensitive_keep_first(values: list) -> list:
    seen = set()
    out = []
    for v in values:
        s = (v or "").strip()
        if not s:
            continue
        key = s.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(s)
    return out


MATERIAL_TYPE_PRESET = dedupe_case_insensitive_keep_first(MATERIAL_TYPE_PRESET_RAW)

# ============================
# Fitting Types
# ============================

_BASE_FITTING_TYPES = [
    "Adapter","Cap","Coupling","Elbow","Flange","Nipple","Olet","Pipe",
    "Reducer","Sleeve","Tee","Union","Valve","Wye","Unclassified",
]


def _legend_bucket_values_from_embedded() -> set:
    out = set()
    try:
        if legend_embedded is None:
            return out
        payload = getattr(legend_embedded, "LEGEND_PAYLOAD", None)
        if not isinstance(payload, dict):
            return out
        for key in ("concat","concat_nospace","desc","desc_nospace"):
            d = payload.get(key, {})
            if isinstance(d, dict):
                for v in d.values():
                    s = str(v or "").strip()
                    if s and s.lower() != "n/a":
                        out.add(s)
    except Exception:
        return set()
    return out


def _build_fitting_types() -> list:
    buckets   = _legend_bucket_values_from_embedded()
    all_types = dedupe_case_insensitive_keep_first([*_BASE_FITTING_TYPES, *sorted(list(buckets))])
    sorted_types = sorted([t for t in all_types if t.lower() != "unclassified"], key=lambda x: x.lower())
    if any(t.lower() == "unclassified" for t in all_types):
        sorted_types.append("Unclassified")
    return sorted_types


FITTING_TYPES = _build_fitting_types()

# ============================
# Regex patterns
# ============================

LENGTH_RE = re.compile(
    r"("
    r"\d+'\s*-\s*[\d\s/]+\""
    r"|"
    r"\d+\s+\d+\s*/\s*\d+\""
    r"|"
    r"\d+\s*-\s*\d+\s*/\s*\d+\""
    r"|"
    r"\d+\s*/\s*\d+\""
    r"|"
    r"\d+(?:\.\d+)?\""
    r")"
)

BATCH_IN_FILENAME_RE   = re.compile(r"\bB\d{3,6}\b", re.IGNORECASE)
BATCH_IN_ROW_RE        = re.compile(r"\bB\d{3,6}\b", re.IGNORECASE)
FALLBACK_BATCH_TOKEN_RE = re.compile(r"[A-Z0-9]{6,}")

PROGRAM_DEFAULT_MULTIPLIER = 2.00
FALLBACK_MULTIPLIER        = 2.00

# ============================
# Hardcoded company multipliers
# ============================

HARDCODED_COMPANY_SIDE_MULTIPLIERS = {
    'PVDF':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Cast Iron':{'Adapter':1.0,'Cap':1.0,'Coupling':0.0,'Elbow':1.0,'Flange':1.0,'Nipple':2.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Forged Steel':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Carbon Steel':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Malleable Iron':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Copper':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Bronze':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'PVC':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Stainless Steel':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Brass':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Cast Steel':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Ductile Iron':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Polypropylene':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'PVC SDR35':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Polyethylene':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'PEX':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Victaulic (OEM) Imperial':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    '304 Stainless Steel':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Cast Brass':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Aluminum':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Nickel iron':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'CPVC':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Rubber':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'FPM':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Copper B280':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Epoxy Steel/Polyethylene':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Forged Brass':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Steel':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Neoprene':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Engineered Polymer':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Cast Copper':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Cast Bronze':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'ABS':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'MDPE':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'HDPE':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Forged Bronze':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Black Iron':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Forged Carbon Steel':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
    'Cast Carbon Steel':{'Adapter':1.0,'Cap':1.0,'Coupling':1.0,'Elbow':1.0,'Flange':1.0,'Nipple':1.0,'Olet':3.0,'Pipe':0.0,'Reducer':2.0,'Sleeve':0.0,'Tee':2.0,'Union':2.0,'Valve':2.0,'Wye':2.0,'Unclassified':0.0},
}

# ============================
# Helpers
# ============================

def clean_token(x) -> str:
    return str(x).replace("\u201c",'"').replace("\u201d",'"').replace("\n"," ").strip()

def _norm_text(s: str) -> str:
    s = str(s or "").upper()
    s = re.sub(r"[^A-Z0-9]+"," ",s)
    return re.sub(r"\s+"," ",s).strip()

def _norm_text_nospace(s: str) -> str:
    s = str(s or "").upper()
    return re.sub(r"[^A-Z0-9]+","",s).strip()

def _ci_get(d: dict, key: str):
    if not isinstance(d,dict): return None
    k = str(key or "").strip().lower()
    if not k: return None
    for kk,vv in d.items():
        if str(kk).strip().lower() == k:
            return vv
    return None

# ============================
# Legend
# ============================

def build_legend_maps_from_xlsx(xlsx_path: Path):
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is not None:
            headers[str(v).strip().lower()] = c
    col_desc   = headers.get("description", 4)
    col_concat = headers.get("description.1", 5)
    col_fit    = headers.get("multiplier reference", 9)
    col_mfg    = headers.get("manufacturer", 3)
    concat_map = {}
    desc_to_fit = {}
    manufacturers = set()
    for r in range(2, ws.max_row + 1):
        desc   = ws.cell(r, col_desc).value
        concat = ws.cell(r, col_concat).value
        fit    = ws.cell(r, col_fit).value
        mfg    = ws.cell(r, col_mfg).value if col_mfg else None
        fit_s  = str(fit or "").strip()
        if not fit_s or fit_s.lower() == "n/a":
            continue
        if mfg:
            manufacturers.add(_norm_text(mfg))
        if concat:
            k = _norm_text(concat)
            if k: concat_map[k] = fit_s
        if desc:
            k2 = _norm_text(desc)
            if k2: desc_to_fit.setdefault(k2, set()).add(fit_s)
    desc_map = {k2: next(iter(fits)) for k2, fits in desc_to_fit.items() if len(fits) == 1}
    return concat_map, desc_map, manufacturers


def load_legend_maps():
    # 1) Embedded
    try:
        if legend_embedded is not None:
            p = getattr(legend_embedded, "LEGEND_PAYLOAD", None)
            if isinstance(p, dict):
                cm  = p.get("concat", {}) or {}
                cmn = p.get("concat_nospace", {}) or {}
                dm  = p.get("desc", {}) or {}
                dmn = p.get("desc_nospace", {}) or {}
                mfg = set(p.get("manufacturers", []) or [])
                if isinstance(cm, dict) and (cm or cmn or dm or dmn):
                    return cm, cmn, dm, dmn, set(map(_norm_text, mfg))
    except Exception:
        pass

    # 2) Cached JSON
    try:
        if LEGEND_CACHE_PATH.exists():
            p = json.loads(LEGEND_CACHE_PATH.read_text(encoding="utf-8"))
            cm  = p.get("concat", {}) or {}
            cmn = p.get("concat_nospace", {}) or {}
            dm  = p.get("desc", {}) or {}
            dmn = p.get("desc_nospace", {}) or {}
            mfg = set(p.get("manufacturers", []) or [])
            if isinstance(cm, dict) and (cm or cmn or dm or dmn):
                return cm, cmn, dm, dmn, set(map(_norm_text, mfg))
    except Exception:
        pass

    # 3) Legacy XLSX fallback
    if LEGEND_XLSX_PATH.exists():
        cm, dm, mfg = build_legend_maps_from_xlsx(LEGEND_XLSX_PATH)
        cmn = {_norm_text_nospace(k): v for k, v in cm.items()}
        dmn = {_norm_text_nospace(k): v for k, v in dm.items()}
        return cm, cmn, dm, dmn, mfg

    return {}, {}, {}, {}, set()


def classify_fitting(description: str, manufacturers: set, concat_map: dict, concat_map_ns: dict, desc_map: dict, desc_map_ns: dict) -> str:
    raw = clean_token(description)
    if not raw:
        return "Unclassified"

    norm = _norm_text(raw)
    norm_ns = _norm_text_nospace(raw)

    if norm in desc_map:
        return desc_map[norm]
    if norm_ns in desc_map_ns:
        return desc_map_ns[norm_ns]
    if norm in concat_map:
        return concat_map[norm]
    if norm_ns in concat_map_ns:
        return concat_map_ns[norm_ns]

    # strip manufacturer prefix if present
    tokens = norm.split()
    if tokens and manufacturers:
        for n in range(min(3, len(tokens)), 0, -1):
            prefix = " ".join(tokens[:n])
            if prefix in manufacturers:
                trimmed = " ".join(tokens[n:]).strip()
                trimmed_ns = _norm_text_nospace(trimmed)
                if trimmed in desc_map:
                    return desc_map[trimmed]
                if trimmed_ns in desc_map_ns:
                    return desc_map_ns[trimmed_ns]
                if trimmed in concat_map:
                    return concat_map[trimmed]
                if trimmed_ns in concat_map_ns:
                    return concat_map_ns[trimmed_ns]
                break

    return "Unclassified"

# ============================
# Material normalization / matching
# ============================

def normalize_material_token(s: str) -> str:
    s = _norm_text(s)
    s = s.replace("POLYPROPELENE", "POLYPROPYLENE")
    return s


def build_material_alias_map(known_material_types: list) -> dict:
    alias = {}
    canon_norm = {}
    for m in known_material_types:
        canonical = str(m or "").strip()
        if not canonical:
            continue
        norm = normalize_material_token(canonical)
        canon_norm[norm] = canonical
        alias[norm] = canonical
        alias[_norm_text_nospace(canonical)] = canonical

    def add(a, canonical):
        alias[_norm_text(a)]               = canonical
        alias[_norm_text_nospace(a)]       = canonical

    if "COPPER"         in canon_norm: add("COOPER", canon_norm["COPPER"]); add("CU", canon_norm["COPPER"])
    if "PVC"            in canon_norm: add("POLYVINYL CHLORIDE", canon_norm["PVC"])
    for k, v in canon_norm.items():
        if k in ("STAINLESS STEEL","304 STAINLESS STEEL"): add("SS", v); add("S S", v)
    if "CAST IRON"      in canon_norm: add("CI", canon_norm["CAST IRON"])
    if "CARBON STEEL"   in canon_norm: add("CS", canon_norm["CARBON STEEL"])
    if "COPPER B280"    in canon_norm: add("COPPER R280", canon_norm["COPPER B280"]); add("B280", canon_norm["COPPER B280"]); add("R280", canon_norm["COPPER B280"])
    if "NICKEL IRON"    in canon_norm: add("NICKELIRON", canon_norm["NICKEL IRON"])
    return alias


def material_type_from_material(material: str, known_material_types: list) -> str:
    raw  = clean_token(material)
    if not raw: return "UNKNOWN"
    mats = [str(x).strip() for x in (known_material_types or []) if str(x).strip()]
    if not mats: return "UNKNOWN"
    norm  = normalize_material_token(raw)
    ns    = _norm_text_nospace(raw)
    alias = build_material_alias_map(mats)
    if norm in alias: return alias[norm]
    if ns   in alias: return alias[ns]
    parts = norm.split()
    for n in range(min(4,len(parts)),0,-1):
        key    = " ".join(parts[:n])
        key_ns = _norm_text_nospace(key)
        if key    in alias: return alias[key]
        if key_ns in alias: return alias[key_ns]
    for m in mats:
        mk = normalize_material_token(m)
        if mk and mk in norm: return m
    first = raw.split()[0].strip()
    return first if first else "UNKNOWN"

# ============================
# Settings / Defaults
# ============================

DEFAULT_SETTINGS = {
    "admin_password_hash":        "",
    "multiplier_mode":            "Company",
    "selected_project":           "",
    "material_types":             MATERIAL_TYPE_PRESET,
    "company_side_multipliers":   {},
    "project_side_multipliers":   {},
    "exclude_fitting_types":      {ft: False for ft in _BASE_FITTING_TYPES},
}


def _clean_multiplier_table(tbl) -> dict:
    if not isinstance(tbl, dict): return {}
    out = {}
    for mat, mp in tbl.items():
        mat_s = str(mat or "").strip()
        if not mat_s or not isinstance(mp, dict): continue
        out[mat_s] = {}
        for ft, val in mp.items():
            ft_s = str(ft or "").strip()
            if not ft_s: continue
            try: out[mat_s][ft_s] = round(float(val), 2)
            except Exception: pass
    return out


def _deep_merge_defaults(data: dict, defaults: dict) -> dict:
    out = dict(defaults)
    for k, v in data.items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = _deep_merge_defaults(v, out[k])
        else:
            out[k] = v
    return out


def build_full_default_table(material_types: list) -> dict:
    out = {}
    for mat in material_types:
        out[mat] = {}
        base_mat = _ci_get(HARDCODED_COMPANY_SIDE_MULTIPLIERS, mat)
        for ft in FITTING_TYPES:
            if isinstance(base_mat, dict) and ft in base_mat:
                out[mat][ft] = round(float(base_mat[ft]), 2)
            else:
                out[mat][ft] = round(float(PROGRAM_DEFAULT_MULTIPLIER), 2)
    return out


def ensure_admin_password_hash(settings: dict) -> dict:
    password_hash = str(settings.get("admin_password_hash", "") or "").strip()
    legacy_password = str(settings.get("admin_password", "") or "").strip()
    if not password_hash:
        seed_password = legacy_password or DEFAULT_ADMIN_PASSWORD
        settings["admin_password_hash"] = generate_password_hash(seed_password)
    settings.pop("admin_password", None)
    return settings


def password_matches(settings: dict, password: str) -> bool:
    ensure_admin_password_hash(settings)
    provided = str(password or "")
    password_hash = str(settings.get("admin_password_hash", "") or "")
    if not password_hash:
        return False
    try:
        return check_password_hash(password_hash, provided)
    except ValueError:
        return False


def set_admin_password(settings: dict, password: str) -> dict:
    settings["admin_password_hash"] = generate_password_hash(str(password or ""))
    settings.pop("admin_password", None)
    return settings


def ensure_company_defaults(settings: dict) -> dict:
    settings = ensure_admin_password_hash(settings)
    mats = settings.get("material_types", [])
    if not isinstance(mats, list) or not mats:
        mats = MATERIAL_TYPE_PRESET
    mats = dedupe_case_insensitive_keep_first([str(x).strip() for x in mats if str(x).strip()])
    settings["material_types"] = mats

    existing = _clean_multiplier_table(settings.get("company_side_multipliers", {}))
    full     = build_full_default_table(mats)
    for mat in mats:
        full.setdefault(mat, {})
        mat_tbl = _ci_get(existing, mat)
        if isinstance(mat_tbl, dict):
            for ft, val in mat_tbl.items():
                try: full[mat][ft] = round(float(val), 2)
                except Exception: pass
        for ft in FITTING_TYPES:
            full[mat].setdefault(ft, round(float(PROGRAM_DEFAULT_MULTIPLIER), 2))
    settings["company_side_multipliers"] = full

    ex = settings.get("exclude_fitting_types", {})
    if not isinstance(ex, dict): ex = {}
    for ft in FITTING_TYPES:
        ex.setdefault(ft, False)
    settings["exclude_fitting_types"] = ex
    return settings


def load_settings() -> dict:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if SETTINGS_FILE.exists():
        try:
            s = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            s = _deep_merge_defaults(s, DEFAULT_SETTINGS)
            mats = s.get("material_types", [])
            if not isinstance(mats, list) or not mats:
                mats = MATERIAL_TYPE_PRESET
            s["material_types"] = dedupe_case_insensitive_keep_first([str(x).strip() for x in mats if str(x).strip()])
            s["company_side_multipliers"] = _clean_multiplier_table(s.get("company_side_multipliers", {}))
            psm = s.get("project_side_multipliers", {})
            s["project_side_multipliers"] = {str(k).strip(): _clean_multiplier_table(v) for k, v in psm.items() if str(k).strip()}
            mode = str(s.get("multiplier_mode","Company") or "Company")
            s["multiplier_mode"] = mode if mode in ("Company","Project") else "Company"
            s = ensure_company_defaults(s)
            return s
        except Exception:
            pass
    s = json.loads(json.dumps(DEFAULT_SETTINGS))
    s = ensure_company_defaults(s)
    return s


def save_settings(settings: dict) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    sanitized = ensure_company_defaults(dict(settings))
    sanitized.pop("admin_password", None)
    SETTINGS_FILE.write_text(json.dumps(sanitized, indent=2), encoding="utf-8")

# ============================
# Multiplier lookup
# ============================

def get_effective_multiplier(settings: dict, project: str, material: str, fitting: str) -> float:
    mat = str(material).strip()
    fit = str(fitting).strip()
    if project and project != "Company Default":
        psm = settings.get("project_side_multipliers", {})
        proj_tbl = _ci_get(psm, project)
        if isinstance(proj_tbl, dict):
            mp = _ci_get(proj_tbl, mat)
            if isinstance(mp, dict):
                v = _ci_get(mp, fit)
                if v is not None:
                    try: return round(float(v), 2)
                    except Exception: pass
    csm = settings.get("company_side_multipliers", {})
    if isinstance(csm, dict):
        mp = _ci_get(csm, mat)
        if isinstance(mp, dict):
            v = _ci_get(mp, fit)
            if v is not None:
                try: return round(float(v), 2)
                except Exception: pass
    return round(float(PROGRAM_DEFAULT_MULTIPLIER), 2)

# ============================
# Size parsing
# ============================

def _clean_size_text(s: str) -> str:
    s = clean_token(s).replace("ø","").replace("Ø","").replace('"',"").lower()
    s = re.sub(r"[^0-9x/\-.\s]"," ",s)
    return re.sub(r"\s+"," ",s).strip()


def parse_single_size(s: str) -> float:
    s = _clean_size_text(s)
    m = re.fullmatch(r"(\d+)\s+(\d+)\s*/\s*(\d+)", s)
    if m:
        w,n,d = m.groups(); return float(w)+float(n)/float(d)
    m = re.fullmatch(r"(\d+)\s*-\s*(\d+)\s*/\s*(\d+)", s)
    if m:
        w,n,d = m.groups(); return float(w)+float(n)/float(d)
    m = re.fullmatch(r"(\d+)\s*/\s*(\d+)", s)
    if m:
        n,d = m.groups(); return float(n)/float(d)
    m = re.fullmatch(r"\d+(\.\d+)?", s)
    if m:
        return float(s)
    raise ValueError(f"Unparseable size: {s}")


def size_to_diameter_in(size_str: str) -> float:
    s = _clean_size_text(size_str)
    s = re.sub(r"\s*x\s*"," x ",s)
    parts = [p.strip() for p in s.split(" x ") if p.strip()]
    if not parts:
        raise ValueError(f"No parsable size parts: {size_str}")
    parsed = []
    for p in parts:
        try:
            parsed.append(parse_single_size(p))
        except Exception:
            m = re.search(r"\d+(\.\d+)?|\d+\s*/\s*\d+|\d+\s*-\s*\d+\s*/\s*\d+|\d+\s+\d+\s*/\s*\d+", p)
            if m: parsed.append(parse_single_size(m.group(0)))
    if not parsed:
        raise ValueError(f"Could not parse any diameter from: {size_str}")
    return max(parsed)

# ============================
# Batch finding
# ============================

def _batch_from_filename(filename: str):
    m = BATCH_IN_FILENAME_RE.search(filename.upper())
    return m.group(0).upper() if m else None


def _find_batch_in_row(items: list, filename_batch):
    for tok in items:
        m = BATCH_IN_ROW_RE.search(tok.upper())
        if m: return m.group(0).upper()
    if filename_batch: return filename_batch
    for tok in reversed(items):
        if FALLBACK_BATCH_TOKEN_RE.fullmatch(tok.upper()): return tok.upper()
    return None

# ============================
# Row normalization
# ============================

def _split_length_and_trailing_material(token: str):
    raw = clean_token(token)
    m = LENGTH_RE.search(raw.replace("Ø","").replace("ø",""))
    if not m: return None, raw
    length = m.group(1)
    after  = raw[m.end():].strip(" -–—\t")
    return length, after


def normalize_row_tokens(tokens: list) -> list:
    return [clean_token(t) for t in tokens if clean_token(t)]


def _extract_length_from_tokens(tokens: list):
    for i, tok in enumerate(tokens):
        m = LENGTH_RE.search(tok.replace("Ø","").replace("ø",""))
        if m:
            return i, m.group(1)
    return None, None


def _extract_description_material_size(tokens: list):
    idx, length = _extract_length_from_tokens(tokens)
    if idx is None:
        desc = " ".join(tokens[:-2]).strip() if len(tokens) >= 2 else " ".join(tokens).strip()
        size = tokens[-2] if len(tokens) >= 2 else ""
        material = tokens[-1] if len(tokens) >= 1 else ""
        return desc, size, material, ""

    before = tokens[:idx]
    after  = tokens[idx + 1:]
    trailing_material = ""
    if after:
        trailing_material = after[-1]
        after = after[:-1]

    material = trailing_material
    if before:
        if len(before) >= 2:
            desc = " ".join(before[:-1]).strip()
            size = before[-1]
        else:
            desc = before[0]
            size = ""
    else:
        desc = ""
        size = ""

    return desc, size, material, length


def row_to_record(tokens: list, filename: str, settings: dict, legend_maps) -> dict:
    concat_map, concat_map_ns, desc_map, desc_map_ns, manufacturers = legend_maps
    items = normalize_row_tokens(tokens)
    if not items:
        raise ValueError("Empty row")

    desc, size, material, length = _extract_description_material_size(items)
    batch = _find_batch_in_row(items, _batch_from_filename(filename))
    material_type = material_type_from_material(material, settings.get("material_types", []))
    fitting_type = classify_fitting(desc, manufacturers, concat_map, concat_map_ns, desc_map, desc_map_ns)

    inches = 0.0
    warnings = []
    errors = []

    try:
        inches = round(size_to_diameter_in(size) * get_effective_multiplier(settings, settings.get("selected_project", ""), material_type, fitting_type), 2)
    except Exception as exc:
        errors.append(str(exc))

    if not batch:
        warnings.append("Missing batch")
    if fitting_type == "Unclassified":
        warnings.append("Unclassified fitting")

    return {
        "batch": batch or "",
        "description": desc,
        "size": size,
        "material": material,
        "material_type": material_type,
        "fitting_type": fitting_type,
        "length": length,
        "inches": inches,
        "warnings": "; ".join(warnings),
        "errors": "; ".join(errors),
    }


def records_from_pdf(pdf_path: str, settings: dict, legend_maps) -> list:
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for tbl in tables:
                if not tbl:
                    continue
                for row in tbl[1:]:
                    if not row:
                        continue
                    try:
                        rows.append(row_to_record(list(row), Path(pdf_path).name, settings, legend_maps))
                    except Exception:
                        continue
    return rows


def summarize_records(df: pd.DataFrame) -> str:
    if df.empty:
        return "No rows found."
    lines = [
        f"Rows: {len(df)}",
        f"Total inches: {df['inches'].fillna(0).sum():,.2f}",
        f"Warnings: {(df['warnings'].fillna('') != '').sum()}",
        f"Errors: {(df['errors'].fillna('') != '').sum()}",
    ]
    return "\n".join(lines)


def export_records(df: pd.DataFrame, export_path: Path) -> None:
    export_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(export_path, index=False)
    wb = load_workbook(export_path)
    ws = wb.active
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for idx, column_cells in enumerate(ws.columns, start=1):
        max_len = max(len(str(c.value or "")) for c in column_cells)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 40)
    table_ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    tbl = Table(displayName="BOMExport", ref=table_ref)
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tbl)
    wb.save(export_path)


def run_bom(pdf_paths: list, settings: dict, export_filename: str, mode: str, project: str) -> dict:
    settings = ensure_company_defaults(dict(settings))
    settings["multiplier_mode"] = mode if mode in ("Company", "Project") else "Company"
    settings["selected_project"] = project if settings["multiplier_mode"] == "Project" else ""

    legend_maps = load_legend_maps()
    records = []
    for pdf_path in pdf_paths:
        records.extend(records_from_pdf(pdf_path, settings, legend_maps))

    df = pd.DataFrame(records)
    if not df.empty:
        excluded = settings.get("exclude_fitting_types", {})
        df = df[~df["fitting_type"].map(lambda ft: bool(excluded.get(ft, False)))]
        df = df.reset_index(drop=True)

    export_base = str(export_filename or "BOM_Export").strip() or "BOM_Export"
    export_path = EXPORTS_DIR / f"{export_base}.xlsx"
    export_records(df, export_path)

    summary = summarize_records(df)
    warn_rows = int((df["warnings"].fillna("") != "").sum()) if not df.empty else 0
    err_rows = int((df["errors"].fillna("") != "").sum()) if not df.empty else 0

    return {
        "ok": True,
        "rows": int(len(df)),
        "ok_rows": int(len(df) - warn_rows - err_rows),
        "warn_rows": warn_rows,
        "err_rows": err_rows,
        "total_inches": round(float(df["inches"].fillna(0).sum()), 2) if not df.empty else 0.0,
        "summary": summary,
        "output": str(export_path),
        "output_filename": export_path.name,
    }
