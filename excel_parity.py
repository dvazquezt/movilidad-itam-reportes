"""
Réplica EXACTA de las fórmulas actuales del Excel (semántica COUNTIFS) para
demostrar que Python reproduce los números de hoy ANTES de corregir nada.

Si la réplica coincide con los valores cacheados del Excel -> entendimos la
lógica actual.  Después, las diferencias entre "réplica" y "versión corregida"
se explican regla por regla (ver reports.balance_explicado).

Semántica COUNTIFS emulada:
  * comparación de texto sin distinguir mayúsculas
  * comodines * y ? en criterios de igualdad / desigualdad
  * operadores <>, >=, <=, >, <, =
  * celdas vacías: sólo cumplen "<>valor"
"""
from __future__ import annotations

import re

import pandas as pd

from load_data import cell

COL = {  # letra de columna en In-Out (Todos) -> encabezado
    "A": "Flujo", "B": "Tipo", "D": "clave carrera", "I": "Semestre", "J": "Semestre Principal",
    "K": "año", "L": "ciclo escolar", "M": "Nivel", "N": "País", "O": "PIM",
    "P": "Región (PIM)", "Q": "Campo Patlani", "R": "Universidad", "S": "Sexo",
}


def _norm(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return str(v).casefold()


def _wild(p: str) -> re.Pattern:
    out = ""
    i = 0
    while i < len(p):
        ch = p[i]
        if ch == "~" and i + 1 < len(p):
            out += re.escape(p[i + 1]); i += 2; continue
        out += ".*" if ch == "*" else "." if ch == "?" else re.escape(ch)
        i += 1
    return re.compile(rf"^{out}$", re.DOTALL)


def crit_mask(series: pd.Series, crit) -> pd.Series:
    s = series.map(_norm)
    if not isinstance(crit, str):
        crit = str(crit)
    op = "="
    for o in ("<>", ">=", "<=", ">", "<", "="):
        if crit.startswith(o):
            op, crit = o, crit[len(o):]
            break
    c = crit.casefold()
    if op in ("=", "<>"):
        if c == "":
            eq = s.isna()
        elif "*" in c or "?" in c:
            rx = _wild(c)
            eq = s.map(lambda x: x is not None and bool(rx.match(x)))
        else:
            eq = s == c
        return eq if op == "=" else ~eq
    cmp = {">=": lambda x: x >= c, "<=": lambda x: x <= c, ">": lambda x: x > c, "<": lambda x: x < c}[op]
    return s.map(lambda x: x is not None and cmp(x))


def countifs(raw: pd.DataFrame, *pairs) -> int:
    m = pd.Series(True, index=raw.index)
    for col, crit in zip(pairs[::2], pairs[1::2]):
        m &= crit_mask(raw[COL[col]], crit)
    return int(m.sum())


# ---------------------------------------------------------------------------
# Balance
# ---------------------------------------------------------------------------
BALANCE_COLS = {  # columna Balance -> criterios (además de universidad y rango de ciclo)
    "C": [("A", "In*"), ("B", "Normal"), ("I", "<>Verano")],
    "D": [("A", "Out*"), ("B", "Normal"), ("I", "<>Verano")],
    "E": [("A", "In*"), ("B", "Normal"), ("I", "Verano")],
    "F": [("A", "Out*"), ("B", "Normal"), ("I", "Verano")],
    "G": [("B", "Verano *")],
    "H": [("A", "Out*"), ("B", "Especial"), ("I", "Verano")],
    "I": [("A", "In*"), ("B", "Especial"), ("I", "<>Verano")],
    "J": [("A", "Out*"), ("B", "Especial"), ("I", "<>Verano")],
    "K": [("B", "Visiting"), ("I", "<>Verano")],
    "L": [("B", "Visiting"), ("I", "Verano")],
    "M": [("B", "Sin*"), ("I", "<>Verano")],
    "N": [("B", "Sin*"), ("I", "Verano")],
}


def replicate_balance(raw: pd.DataFrame, grid: pd.DataFrame, ciclo_ini: str | None = None,
                      ciclo_fin: str | None = None) -> pd.DataFrame:
    ciclo_fin = ciclo_fin or cell(grid, "C1")
    ciclo_ini = ciclo_ini or cell(grid, "A1")
    rows = []
    uni_l = raw["Universidad"].map(_norm)
    cic = raw["ciclo escolar"].map(_norm)
    win = cic.map(lambda x: x is not None and ciclo_ini.casefold() <= x <= ciclo_fin.casefold())
    for r in range(6, len(grid) + 1):
        name = cell(grid, f"B{r}")
        if name is None or (isinstance(name, float) and pd.isna(name)):
            continue
        base = win & crit_mask(raw["Universidad"], f"{name}*")
        rec = {"fila_balance": r, "universidad_balance": name}
        for col, crits in BALANCE_COLS.items():
            m = base.copy()
            for c, k in crits:
                m &= crit_mask(raw[COL[c]], k)
            rec[f"py_{col}"] = int(m.sum())
            rec[f"xl_{col}"] = cell(grid, f"{col}{r}")
        rec["py_O"] = rec["py_D"] - rec["py_C"] + 0.5 * (rec["py_F"] - rec["py_E"])
        rec["xl_O"] = cell(grid, f"O{r}")
        rec["xl_A_incluir"] = cell(grid, f"A{r}")
        rows.append(rec)
    out = pd.DataFrame(rows)
    out["coincide"] = True
    for col in list(BALANCE_COLS) + ["O"]:
        out["coincide"] &= pd.to_numeric(out[f"xl_{col}"], errors="coerce").fillna(-999) == out[f"py_{col}"]
    out.attrs["ventana"] = (ciclo_ini, ciclo_fin)
    return out


# ---------------------------------------------------------------------------
# Reporte Anual y PIM: celdas con COUNTIFS escalares
# ---------------------------------------------------------------------------
def _reporte_anual_specs(c1):
    return {
        "C6 Semestral con convenio": [("I", "<>Verano"), ("B", "*al"), ("L", c1), ("A", "Out*")],
        "C7 Semestral sin convenio": [("I", "<>Verano"), ("B", "Sin convenio"), ("L", c1), ("A", "Out*")],
        "C8 Verano con convenio": [("I", "Verano"), ("B", "*al"), ("L", c1), ("A", "Out*")],
        "C9 Verano sin convenio": [("I", "Verano"), ("B", "Sin convenio"), ("L", c1), ("A", "Out*")],
        "C10 Dobles grados": [("B", "Doble Grado"), ("L", c1)],
        "C11 Verano en el extranjero": [("B", "Verano*"), ("L", c1), ("A", "Out*")],
        "D27 Lic sin convenio": [("M", "Lic*"), ("B", "Sin*"), ("L", c1), ("A", "Out*")],
        "E27 Lic verano extranjero": [("M", "Lic*"), ("B", "Verano*"), ("L", c1), ("A", "Out*")],
        "F27 Maestría out": [("M", "Maestría"), ("L", c1), ("A", "Out*")],
        "C53 Lic in": [("A", "In*"), ("L", c1), ("M", "Lic*")],
        "D53 Lic out": [("M", "Lic*"), ("L", c1), ("A", "Out*")],
        "E53 Mae in": [("A", "In*"), ("L", c1), ("M", "Mae*")],
        "F53 Mae out": [("M", "Mae*"), ("L", c1), ("A", "Out*")],
    }


def replicate_reporte_anual(raw, grid) -> pd.DataFrame:
    c1 = cell(grid, "C1")
    rows = []
    for k, crits in _reporte_anual_specs(c1).items():
        ref = k.split()[0]
        py = countifs(raw, *[x for p in crits for x in p])
        rows.append({"celda": ref, "concepto": k, "ciclo": c1, "excel": cell(grid, ref), "python": py})
    # C27 = SUM(COUNTIFS(B:B,{"*al";"Doble*"}, M "Lic*", L c1, A "Out*"))
    py = sum(countifs(raw, "B", b, "M", "Lic*", "L", c1, "A", "Out*") for b in ("*al", "Doble*"))
    rows.append({"celda": "C27", "concepto": "C27 Lic por convenio (+doble grado)", "ciclo": c1,
                 "excel": cell(grid, "C27"), "python": py})
    out = pd.DataFrame(rows)
    out["coincide"] = pd.to_numeric(out["excel"], errors="coerce") == out["python"]
    return out


def replicate_pim(raw, grid) -> pd.DataFrame:
    b2 = cell(grid, "B2")
    rows = []
    for first, flujo in ((6, "Out*"), (18, "In*")):
        for r in range(first, first + 7):
            reg = cell(grid, f"A{r}")
            for col, niv, pim in (("B", "Lic*", True), ("C", "Ma*", True), ("F", "Lic*", False), ("G", "Ma*", False)):
                crits = ["L", b2, "M", niv, "P", reg, "A", flujo] + (["O", "Sí"] if pim else [])
                rows.append({"celda": f"{col}{r}", "flujo": flujo.strip("*"), "region": reg, "nivel": niv.strip("*"),
                             "solo_pim": pim, "ciclo": b2, "excel": cell(grid, f"{col}{r}"),
                             "python": countifs(raw, *crits)})
    out = pd.DataFrame(rows)
    out["coincide"] = pd.to_numeric(out["excel"], errors="coerce") == out["python"]
    return out
