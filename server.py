"""
Электоральная карта НСО.
Источник данных — Excel-файл с листами:
  • «УИК»   — результаты по каждому участку
  • «ТИК»   — агрегаты по территориальным комиссиям (районы/города)
  • «ОблИК» — агрегат по всей области

Запуск: python server.py
Открыть: http://localhost:8000
"""
from pathlib import Path
from typing import Optional, List, Dict, Any

import pandas as pd
import numpy as np
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse
from fastapi.middleware.cors import CORSMiddleware

BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "база результатов для редакции .xlsx"

# Порядок партий в новой базе (столбцы I..R):
# I/J   — ЕР абс/%
# K/L   — КПРФ абс/%
# M/N   — Новые люди абс/%
# O/P   — ЛДПР абс/%
# Q/R   — СРЗП абс/%
PARTY_ORDER = ["ER", "KPRF", "NEWPEOPLE", "LDPR", "SRZP"]
PARTY_NAMES = {
    "ER": "Единая Россия",
    "KPRF": "КПРФ",
    "NEWPEOPLE": "Новые люди",
    "LDPR": "ЛДПР",
    "SRZP": "Справедливая Россия",
}
PARTY_SHORT = {
    "ER": "ЕР", "KPRF": "КПРФ", "NEWPEOPLE": "НЛ",
    "LDPR": "ЛДПР", "SRZP": "СРЗП",
}
PARTY_COLORS = {
    "ER": "#1e6bb8", "KPRF": "#c1121f", "NEWPEOPLE": "#0b8a3a",
    "LDPR": "#0a3d91", "SRZP": "#f5a623",
}


def _normalize_pct(x):
    """Приводит значение из базы к «человеческому» проценту."""
    if x is None:
        return None
    try:
        v = float(x)
    except (ValueError, TypeError):
        return None
    if v <= 2:
        v = v * 100.0
    return round(v, 2)


def _parse_pct(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace("%", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _parse_int(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    try:
        return int(float(str(x).replace(" ", "").replace(",", ".")))
    except (ValueError, TypeError):
        return None


def _party_dict(votes, pct):
    return {"votes": votes, "pct": pct}


# ---------- Загрузка листов ----------
def load_uik_sheet(xl: pd.ExcelFile) -> List[dict]:
    """Лист «УИК». Строки — участки."""
    df = xl.parse("УИК")
    records = []
    for _, row in df.iterrows():
        code = row.iloc[0]
        year = row.iloc[1]
        territory = row.iloc[2]
        uik = row.iloc[4]
        if pd.isna(code) or pd.isna(uik):
            continue
        try:
            uik_num = int(float(uik))
        except (ValueError, TypeError):
            continue
        rec = {
            "level": str(code).strip(),
            "year": _parse_int(year),
            "territory": str(territory).strip() if not pd.isna(territory) else None,
            "uik": uik_num,
            "turnout_abs": _parse_int(row.iloc[5]),
            "turnout_pct": _parse_pct(row.iloc[6]),
            "voters": _parse_int(row.iloc[7]),
            "parties": {
                "ER":        _party_dict(_parse_int(row.iloc[8]),  _parse_pct(row.iloc[9])),
                "KPRF":      _party_dict(_parse_int(row.iloc[10]), _parse_pct(row.iloc[11])),
                "NEWPEOPLE": _party_dict(_parse_int(row.iloc[12]), _parse_pct(row.iloc[13])),
                "LDPR":      _party_dict(_parse_int(row.iloc[14]), _parse_pct(row.iloc[15])),
                "SRZP":      _party_dict(_parse_int(row.iloc[16]), _parse_pct(row.iloc[17])),
            },
        }
        records.append(rec)
    return records


def load_tik_sheet(xl: pd.ExcelFile) -> List[dict]:
    """Лист «ТИК». Строки — территории."""
    df = xl.parse("ТИК")
    records = []
    for _, row in df.iterrows():
        code = row.iloc[0]
        year = row.iloc[1]
        territory = row.iloc[2]
        if pd.isna(code) or pd.isna(territory):
            continue
        rec = {
            "level": str(code).strip(),
            "year": _parse_int(year),
            "territory": str(territory).strip(),
            "uik_count": _parse_int(row.iloc[3]),
            "turnout_abs": _parse_int(row.iloc[5]),
            "turnout_pct": _parse_pct(row.iloc[6]),
            "voters": _parse_int(row.iloc[7]),
            "parties": {
                "ER":        _party_dict(_parse_int(row.iloc[8]),  _parse_pct(row.iloc[9])),
                "KPRF":      _party_dict(_parse_int(row.iloc[10]), _parse_pct(row.iloc[11])),
                "NEWPEOPLE": _party_dict(_parse_int(row.iloc[12]), _parse_pct(row.iloc[13])),
                "LDPR":      _party_dict(_parse_int(row.iloc[14]), _parse_pct(row.iloc[15])),
                "SRZP":      _party_dict(_parse_int(row.iloc[16]), _parse_pct(row.iloc[17])),
            },
        }
        records.append(rec)
    return records


def load_oblik_sheet(xl: pd.ExcelFile) -> List[dict]:
    """Лист «ОблИК». Одна строка на выборы (вся область)."""
    df = xl.parse("ОблИК")
    records = []
    for _, row in df.iterrows():
        code = row.iloc[0]
        if pd.isna(code):
            continue
        rec = {
            "level": str(code).strip(),
            "year": _parse_int(row.iloc[1]),
            "territory": str(row.iloc[2]).strip(),
            "turnout_abs": _parse_int(row.iloc[5]),
            "turnout_pct": _parse_pct(row.iloc[6]),
            "voters": _parse_int(row.iloc[7]),
            "parties": {
                "ER":        _party_dict(_parse_int(row.iloc[8]),  _parse_pct(row.iloc[9])),
                "KPRF":      _party_dict(_parse_int(row.iloc[10]), _parse_pct(row.iloc[11])),
                "NEWPEOPLE": _party_dict(_parse_int(row.iloc[12]), _parse_pct(row.iloc[13])),
                "LDPR":      _party_dict(_parse_int(row.iloc[14]), _parse_pct(row.iloc[15])),
                "SRZP":      _party_dict(_parse_int(row.iloc[16]), _parse_pct(row.iloc[17])),
            },
        }
        records.append(rec)
    return records


print("Загрузка Excel...")
if not EXCEL_FILE.exists():
    raise FileNotFoundError(f"Не найден файл: {EXCEL_FILE}")

xl = pd.ExcelFile(EXCEL_FILE)
UIK_DATA = load_uik_sheet(xl)
TIK_DATA = load_tik_sheet(xl)
OBLIK_DATA = load_oblik_sheet(xl)
print(f"  УИК:   {len(UIK_DATA)}")
print(f"  ТИК:   {len(TIK_DATA)}")
print(f"  ОблИК: {len(OBLIK_DATA)}")

# ---- Реестр выборов ----
def _election_key(level: str, year: int) -> str:
    return f"{level}|{year}"

ELECTIONS: Dict[str, Dict[str, Any]] = {}
for r in TIK_DATA + OBLIK_DATA + UIK_DATA:
    if not r.get("year"):
        continue
    key = _election_key(r["level"], r["year"])
    if key not in ELECTIONS:
        ELECTIONS[key] = {
            "code": key,
            "level": r["level"],
            "year": r["year"],
            "title": f'{r["level"]} — {r["year"]}',
            "date": f'{r["year"]}-01-01',
        }

ELECTIONS = dict(sorted(ELECTIONS.items(), key=lambda kv: (kv[1]["year"], kv[1]["level"])))


def find_code(code: str) -> dict:
    if code not in ELECTIONS:
        raise HTTPException(404, f"Неизвестный код выборов: {code}")
    return ELECTIONS[code]


def get_tik(code: str) -> List[dict]:
    meta = find_code(code)
    return [r for r in TIK_DATA
            if r["level"] == meta["level"] and r["year"] == meta["year"]]


def get_uik(code: str, territory: Optional[str] = None) -> List[dict]:
    meta = find_code(code)
    rows = [r for r in UIK_DATA
            if r["level"] == meta["level"] and r["year"] == meta["year"]]
    if territory:
        rows = [r for r in rows if r["territory"] == territory]
    return rows


def get_oblik(code: str) -> Optional[dict]:
    meta = find_code(code)
    for r in OBLIK_DATA:
        if r["level"] == meta["level"] and r["year"] == meta["year"]:
            return r
    return None


# ---------- FastAPI ----------
app = FastAPI(title="Электоральная карта НСО")
app.add_middleware(CORSMiddleware, allow_origins=["*"],
                   allow_methods=["*"], allow_headers=["*"])


@app.get("/api/elections")
def list_elections():
    return [{"code": v["code"], "title": v["title"],
             "date": v["date"], "year": v["year"], "level": v["level"]}
            for v in ELECTIONS.values()]


@app.get("/api/territories/{code}")
def territories(code: str):
    return sorted({r["territory"] for r in get_tik(code)})


# ---------- /api/multi-results ----------
@app.get("/api/multi-results")
def multi_results(codes: str = Query(...),
                  territory: Optional[str] = None,
                  source: str = Query("TIK", pattern="^(TIK|UIK)$")):
    code_list = [c.strip() for c in codes.split(",") if c.strip()]
    out: Dict[str, Any] = {}
    for c in code_list:
        if c not in ELECTIONS:
            continue
        if source == "UIK":
            rows = get_uik(c, territory)
            agg = _aggregate_rows(rows)
            out[c] = [{"territory": territory or "—", **agg}]
        else:
            rows = get_tik(c)
            out[c] = [
                {
                    "territory": r["territory"],
                    "turnout_abs": r["turnout_abs"],
                    "turnout_pct": _normalize_pct(r["turnout_pct"]),
                    "voters": r["voters"],
                    "parties": [
                        {"party_code": pc, "party_name": PARTY_NAMES[pc],
                         "color": PARTY_COLORS[pc],
                         "votes": r["parties"][pc]["votes"],
                         "pct": _normalize_pct(r["parties"][pc]["pct"])}
                        for pc in PARTY_ORDER
                    ],
                } for r in rows
            ]
    return out


def _aggregate_rows(rows: List[dict]) -> dict:
    if not rows:
        return {"turnout_abs": 0, "turnout_pct": 0, "voters": 0,
                "parties": [{"party_code": pc, "party_name": PARTY_NAMES[pc],
                             "color": PARTY_COLORS[pc], "votes": 0, "pct": 0}
                            for pc in PARTY_ORDER]}
    t_abs = sum(r["turnout_abs"] or 0 for r in rows)
    voters = sum(r["voters"] or 0 for r in rows)
    totals = {pc: 0 for pc in PARTY_ORDER}
    pct_sums = {pc: 0.0 for pc in PARTY_ORDER}
    pct_n = {pc: 0 for pc in PARTY_ORDER}
    for r in rows:
        for pc in PARTY_ORDER:
            totals[pc] += r["parties"][pc]["votes"] or 0
            p = _normalize_pct(r["parties"][pc]["pct"])
            if p is not None:
                pct_sums[pc] += p
                pct_n[pc] += 1
    return {
        "turnout_abs": t_abs,
        "turnout_pct": round(t_abs / voters * 100, 2) if voters else None,
        "voters": voters,
        "parties": [
            {"party_code": pc, "party_name": PARTY_NAMES[pc],
             "color": PARTY_COLORS[pc],
             "votes": totals[pc],
             "pct": round(pct_sums[pc] / pct_n[pc], 2) if pct_n[pc] else None}
            for pc in PARTY_ORDER
        ],
    }


# ---------- /api/oblik/{code} — единый областной округ ----------
@app.get("/api/oblik/{code}")
def oblik_result(code: str):
    r = get_oblik(code)
    if not r:
        raise HTTPException(404, "Нет данных ОблИК для выбранных выборов")
    return {
        "territory": r["territory"],
        "turnout_abs": r["turnout_abs"],
        "turnout_pct": _normalize_pct(r["turnout_pct"]),
        "voters": r["voters"],
        "parties": [
            {"party_code": pc, "party_name": PARTY_NAMES[pc],
             "color": PARTY_COLORS[pc],
             "votes": r["parties"][pc]["votes"],
             "pct": _normalize_pct(r["parties"][pc]["pct"])}
            for pc in PARTY_ORDER
        ],
    }


# ---------- /api/results/{code} ----------
@app.get("/api/results/{code}")
def results(code: str, territory: Optional[str] = None):
    if territory:
        rows = get_uik(code, territory)
        return [_aggregate_rows(rows)]
    rows = get_tik(code)
    return [
        {
            "territory": r["territory"],
            "turnout_abs": r["turnout_abs"],
            "turnout_pct": _normalize_pct(r["turnout_pct"]),
            "voters": r["voters"],
            "parties": [
                {"party_code": pc, "party_name": PARTY_NAMES[pc],
                 "color": PARTY_COLORS[pc],
                 "votes": r["parties"][pc]["votes"],
                 "pct": _normalize_pct(r["parties"][pc]["pct"])}
                for pc in PARTY_ORDER
            ],
        } for r in rows
    ]


# ---------- /api/uik-list/{code} ----------
@app.get("/api/uik-list/{code}")
def uik_list(code: str, territory: str):
    rows = get_uik(code, territory)
    return [
        {
            "uik": r["uik"],
            "turnout_abs": r["turnout_abs"],
            "turnout_pct": _normalize_pct(r["turnout_pct"]),
            "voters": r["voters"],
            "parties": [
                {"party_code": pc, "party_name": PARTY_NAMES[pc],
                 "color": PARTY_COLORS[pc],
                 "votes": r["parties"][pc]["votes"],
                 "pct": _normalize_pct(r["parties"][pc]["pct"])}
                for pc in PARTY_ORDER
            ],
        } for r in rows
    ]


# ---------- /api/dynamics ----------
@app.get("/api/dynamics")
def dynamics(territory: Optional[str] = None):
    out = []
    for code, meta in ELECTIONS.items():
        if territory:
            rows = [r for r in get_tik(code) if r["territory"] == territory]
            if not rows:
                continue
            r = rows[0]
        else:
            r = get_oblik(code)
            if not r:
                continue
        for pc in PARTY_ORDER:
            votes = r["parties"][pc]["votes"] or 0
            pct = _normalize_pct(r["parties"][pc]["pct"])
            out.append({
                "election_code": code,
                "election_date": str(meta["year"]),
                "party_code": pc,
                "party_name": PARTY_NAMES[pc],
                "color": PARTY_COLORS[pc],
                "votes": votes,
                "pct": pct,
            })
    out.sort(key=lambda x: (x["election_date"], x["party_code"]))
    return out


# ---------- /api/turnout-dynamics ----------
@app.get("/api/turnout-dynamics")
def turnout_dynamics(territory: Optional[str] = None):
    out = []
    for code, meta in ELECTIONS.items():
        if territory:
            rows = [r for r in get_tik(code) if r["territory"] == territory]
            if not rows:
                continue
            r = rows[0]
            out.append({
                "election_code": code,
                "election_date": str(meta["year"]),
                "turnout_pct": _normalize_pct(r["turnout_pct"]),
                "turnout_abs": r["turnout_abs"],
                "voters": r["voters"],
            })
        else:
            r = get_oblik(code)
            if not r:
                continue
            out.append({
                "election_code": code,
                "election_date": str(meta["year"]),
                "turnout_pct": _normalize_pct(r["turnout_pct"]),
                "turnout_abs": r["turnout_abs"],
                "voters": r["voters"],
            })
    out.sort(key=lambda x: x["election_date"])
    return out


# ---------- /api/top10 ----------
@app.get("/api/top10/{code}/{party}")
def top10(code: str, party: str, by: str = "pct"):
    rows = get_tik(code)
    rows = [r for r in rows if r["parties"].get(party, {}).get(by) is not None]
    rows.sort(key=lambda r: r["parties"][party][by], reverse=True)
    top = rows[:10]
    return [
        {
            "territory": r["territory"],
            "parties": [
                {"party_code": pc, "party_name": PARTY_NAMES[pc],
                 "color": PARTY_COLORS[pc],
                 "votes": r["parties"][pc]["votes"],
                 "pct": _normalize_pct(r["parties"][pc]["pct"])}
                for pc in PARTY_ORDER
            ],
        } for r in top
    ]


# ---------- /api/heatmap-data ----------
@app.get("/api/heatmap-data")
def heatmap_data(codes: str = Query(...),
                 territory: Optional[str] = None):
    code_list = [c.strip() for c in codes.split(",") if c.strip()]
    if territory:
        uik_set = set()
        data: Dict[str, Dict[str, Optional[float]]] = {}
        for c in code_list:
            if c not in ELECTIONS:
                continue
            for r in get_uik(c, territory):
                key = f'УИК {r["uik"]}'
                uik_set.add(key)
                data.setdefault(key, {})[c] = _normalize_pct(r["turnout_pct"])
        return {"rows": sorted(uik_set),
                "elections": code_list,
                "data": data,
                "mode": "uik"}
    else:
        terr_set = set()
        data = {}
        for c in code_list:
            if c not in ELECTIONS:
                continue
            for r in get_tik(c):
                terr_set.add(r["territory"])
                data.setdefault(r["territory"], {})[c] = _normalize_pct(r["turnout_pct"])
        return {"rows": sorted(terr_set),
                "elections": code_list,
                "data": data,
                "mode": "tik"}


# ---------- /api/kprf-wins/{code} ----------
@app.get("/api/kprf-wins/{code}")
def kprf_wins(code: str):
    """
    Возвращает список районов, отсортированных по убыванию количества УИК,
    на которых КПРФ заняла первое место по числу голосов.
    Для каждого района — список номеров таких УИК.
    """
    meta = find_code(code)
    # Собираем все УИК по всем районам данного года
    rows = [r for r in UIK_DATA
            if r["level"] == meta["level"] and r["year"] == meta["year"]]
    # Группируем по районам
    by_terr: Dict[str, List[int]] = {}
    for r in rows:
        terr = r["territory"]
        if not terr:
            continue
        # Определяем победителя по голосам
        best_party = None
        best_votes = -1
        for pc in PARTY_ORDER:
            v = r["parties"][pc]["votes"] or 0
            if v > best_votes:
                best_votes = v
                best_party = pc
        if best_party == "KPRF" and best_votes > 0:
            by_terr.setdefault(terr, []).append(r["uik"])

    # Формируем результат
    result = []
    for terr, uiks in by_terr.items():
        result.append({
            "territory": terr,
            "count": len(uiks),
            "uiks": sorted(uiks),
        })
    result.sort(key=lambda x: x["count"], reverse=True)
    return result


@app.get("/", response_class=HTMLResponse)
def index():
    return (BASE_DIR / "index.html").read_text(encoding="utf-8")


if __name__ == "__main__":
    import uvicorn
    import os
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))