"""
Reportes del MVP (versión "corregida" / propuesta).

  1. balance_convenios      Balance por universidad (con trazabilidad vs Excel)
  2. outgoing_por_carrera   Salientes por carrera: principal vs componente
  3. pim_reporte            Movilidad con la red PIM y % sobre el total

Cada reporte devuelve DataFrames listos para exportar y SIEMPRE distingue
`registros` (filas de movilidad) de `alumnos` (Claves Únicas distintas).
"""
from __future__ import annotations

import numpy as np
import pandas as pd



def _window(df, ciclo_ini, ciclo_fin):
    return df[(df["ciclo"] >= ciclo_ini) & (df["ciclo"] <= ciclo_fin)]


def ciclos_atras(ciclo_fin: str, n: int) -> str:
    """'2024-2025', 5 -> '2020-2021'  (misma lógica que Balance!A1)."""
    a, b = int(ciclo_fin[:4]), int(ciclo_fin[-4:])
    return f"{a - n + 1}-{b - n + 1}"


# ---------------------------------------------------------------------------
# 1. Balance
# ---------------------------------------------------------------------------
def balance_convenios(mov: pd.DataFrame, ciclo_ini: str, ciclo_fin: str, umbral: float = 0.0):
    d = _window(mov, ciclo_ini, ciclo_fin)

    def piv(mask, name):
        return d[mask].groupby("uni_key").size().rename(name)

    N, V, T = d["tipo"] == "Normal", d["es_verano"], d["es_trimestre_maestria"]
    IN, OUT = d["flujo"] == "IN", d["flujo"] == "OUT"
    parts = [
        piv(N & IN & ~V & ~T, "conv_sem_in"), piv(N & OUT & ~V & ~T, "conv_sem_out"),
        piv(N & IN & T, "conv_trim_in"), piv(N & OUT & T, "conv_trim_out"),
        piv(N & IN & V, "conv_ver_in"), piv(N & OUT & V, "conv_ver_out"),
        piv((d["tipo"] == "Especial") & IN, "especial_in"), piv((d["tipo"] == "Especial") & OUT, "especial_out"),
        piv(d["tipo"] == "Visiting", "visiting_in"),
        piv((d["tipo"] == "Sin convenio") & IN, "sin_conv_in"), piv((d["tipo"] == "Sin convenio") & OUT, "sin_conv_out"),
        piv(d["tipo"].isin(["Verano en el extranjero", "Verano LSGL"]), "verano_extranjero_out"),
        piv(d["tipo"] == "Doble Grado", "doble_grado_out"),
        d[d["cuenta_balance"]].groupby("uni_key")["aporte_balance"].sum().rename("balance"),
        d[d["cuenta_balance"] & OUT].groupby("uni_key")["clave_unica"].nunique().rename("alumnos_enviados_unicos"),
        d[d["cuenta_balance"] & IN].groupby("uni_key")["clave_unica"].nunique().rename("alumnos_recibidos_unicos"),
    ]
    t = pd.concat(parts, axis=1).fillna(0)
    meta = d.groupby("uni_key").agg(
        universidad=("universidad_canon", "first"),
        pais=("pais", lambda s: s.value_counts().index[0]),
        pim=("pim", "max"),
        balance_en_unis_info=("unis_info_balance", lambda s: s.dropna().iloc[0] if s.notna().any() else None),
        grafias=("universidad", lambda s: " | ".join(sorted(set(s)))),
    )
    t = meta.join(t).reset_index()
    num = [c for c in t.columns if c.endswith(("_in", "_out")) or c.startswith("alumnos_")]
    t[num] = t[num].astype(int)
    # [COMPROBADO] fórmula de la wishlist (trimestres pesan como semestre, ver config)
    t["semaforo"] = np.where(t["balance"].abs() <= umbral, "Balanceado",
                             np.where(t["balance"] > 0, "En contra del ITAM (enviamos más)",
                                      "A favor del ITAM (recibimos más)"))
    # [AMBIGÜEDAD] la wishlist usa "a favor / en contra" sin definir el signo; aquí
    # balance>0 = ITAM envió más de lo que recibió.  Confirmar con el área.
    t["incluir"] = (t["balance_en_unis_info"] == "Sí") & (t[num].sum(axis=1) > 0)
    t = t.sort_values(["incluir", "balance"], ascending=[False, True])

    hist = (d[d["cuenta_balance"]].pivot_table(index="uni_key", columns="ciclo", values="aporte_balance",
                                                aggfunc="sum", fill_value=0))
    hist.columns = [f"balance_{c}" for c in hist.columns]
    acum = hist.cumsum(axis=1)
    acum.columns = [c.replace("balance_", "acumulado_") for c in acum.columns]
    historico = meta[["universidad", "pais"]].join(hist).join(acum).reset_index()
    return t, historico


def balance_explicado(mov: pd.DataFrame, raw: pd.DataFrame, parity: pd.DataFrame, ciclo_ini, ciclo_fin):
    """Explica por qué el balance del Excel difiere del de Python por universidad."""
    d = _window(mov, ciclo_ini, ciclo_fin)
    d = d[d["tipo"] == "Normal"]
    names = parity["universidad_balance"].tolist()
    rows = []
    # (a) cada registro Normal de la ventana: ¿cuántas filas de Balance lo capturan?
    lower = d["universidad"].str.casefold()
    hits = {i: [n for n in names if l.startswith(str(n).casefold())] for i, l in lower.items()}
    d = d.assign(filas_balance_que_lo_cuentan=[len(hits[i]) for i in d.index],
                 filas=[" | ".join(hits[i]) for i in d.index])
    no_cap = d[d["filas_balance_que_lo_cuentan"] == 0].groupby("universidad").size()
    multi = d[d["filas_balance_que_lo_cuentan"] > 1].groupby(["universidad", "filas"]).size()
    for u, n in no_cap.items():
        rows.append({"tipo_diferencia": "Registro NO capturado por ninguna fila de Balance (queda fuera del balance)",
                     "universidad": u, "registros": n, "detalle": ""})
    for (u, f), n in multi.items():
        rows.append({"tipo_diferencia": "Registro contado en VARIAS filas de Balance (prefijo/fila repetida)",
                     "universidad": u, "registros": n, "detalle": f})
    dup = parity[parity["universidad_balance"].duplicated(keep=False)]
    for u, g in dup.groupby("universidad_balance"):
        rows.append({"tipo_diferencia": "Universidad repetida en la lista de Balance",
                     "universidad": u, "registros": len(g), "detalle": "filas " + ", ".join(map(str, g["fila_balance"]))})
    nan = parity[parity["xl_A_incluir"].astype(str).str.startswith("#")]
    for _, r in nan.iterrows():
        rows.append({"tipo_diferencia": "Nombre de Balance no existe en unis-info (#N/A en 'Incluir')",
                     "universidad": r["universidad_balance"], "registros": 0, "detalle": f"fila {r['fila_balance']}"})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 2. Outgoing por carrera
# ---------------------------------------------------------------------------
def outgoing_por_carrera(mov, bridge, ciclo_ini, ciclo_fin, nivel=None, carreras_cat=None):
    d = _window(mov[mov["flujo"] == "OUT"], ciclo_ini, ciclo_fin)
    b = _window(bridge[bridge["flujo"] == "OUT"], ciclo_ini, ciclo_fin)
    if nivel:
        d, b = d[d["nivel"] == nivel], b[b["nivel"] == nivel]
    a = d.groupby("carrera_principal").agg(registros_carrera_principal=("fuente_fila", "size"),
                                           alumnos_carrera_principal=("clave_unica", "nunique"))
    c = b.groupby("programa").agg(registros_por_componente=("fuente_fila", "size"),
                                  alumnos_por_componente=("clave_unica", "nunique"),
                                  registros_de_conjuntas=("es_conjunta", "sum"))
    t = c.join(a, how="outer").fillna(0).astype(int)
    t["registros_no_contados_en_principal"] = t["registros_por_componente"] - t["registros_carrera_principal"]
    if carreras_cat is not None:
        nm = carreras_cat.set_index("Clave")["Carrera"]
        t.insert(0, "nombre", t.index.map(nm))
    t = t.sort_values("registros_por_componente", ascending=False).reset_index().rename(columns={"index": "programa"})
    totales = pd.DataFrame([{
        "concepto": "Total institucional (cada registro una vez)", "registros": len(d), "alumnos": d["clave_unica"].nunique()},
        {"concepto": "Suma de la columna por componente (NO es un total: conjuntas cuentan 2 veces)",
         "registros": int(t["registros_por_componente"].sum()), "alumnos": None},
        {"concepto": "Registros de carreras conjuntas", "registros": int(b.drop_duplicates("fuente_fila")["es_conjunta"].sum()),
         "alumnos": b[b["es_conjunta"]]["clave_unica"].nunique()},
    ])
    matriz = b.pivot_table(index="programa", columns="pais", values="clave_unica", aggfunc="nunique", fill_value=0)
    return t, totales, matriz


# ---------------------------------------------------------------------------
# 3. PIM
# ---------------------------------------------------------------------------
def pim_reporte(mov, ciclo, meta_pct=20.0):
    d = mov[mov["ciclo"] == ciclo]

    def agg(g):
        return pd.Series({
            "registros_total": len(g), "registros_pim": int(g["pim"].fillna(False).sum()),
            "alumnos_total": g["clave_unica"].nunique(),
            "alumnos_pim": g.loc[g["pim"].fillna(False), "clave_unica"].nunique(),
            "registros_pim_sin_dato": int(g["pim"].isna().sum()),
        })

    res = d.groupby(["flujo", "nivel", "plazo"]).apply(agg, include_groups=False).reset_index()
    tot_f = d.groupby("flujo").apply(agg, include_groups=False).reset_index().assign(nivel="(todos)", plazo="(todos)")
    tot = agg(d).to_frame().T.assign(flujo="(todos)", nivel="(todos)", plazo="(todos)")
    res = pd.concat([res, tot_f, tot], ignore_index=True)
    for k in ("registros", "alumnos"):
        res[f"pct_pim_{k}"] = (res[f"{k}_pim"] / res[f"{k}_total"] * 100).round(1)
    res["cumple_meta_20pct_registros"] = res["pct_pim_registros"] >= meta_pct
    por_region = (d[d["pim"].fillna(False)].pivot_table(index=["region"], columns=["flujo", "nivel"],
                                                       values="fuente_fila", aggfunc="size", fill_value=0))
    por_uni = (d[d["pim"].fillna(False)].groupby(["universidad_canon", "flujo"]).agg(
        registros=("fuente_fila", "size"), alumnos=("clave_unica", "nunique")).unstack(fill_value=0))
    por_uni.columns = [f"{a}_{b}" for a, b in por_uni.columns]
    return res, por_region, por_uni.reset_index()
