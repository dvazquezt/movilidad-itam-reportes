"""
Reglas de negocio explícitas.  Todo lo que hoy está escondido en criterios
de COUNTIFS ("*al", "Verano*", "Sin*", B6&"*") se expresa aquí con nombre.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

import config as C


def classify(df: pd.DataFrame, excluir_cancelados: bool = False) -> pd.DataFrame:
    d = df.copy()
    t, ver, tri = d["tipo"], d["es_verano"], d["es_trimestre_maestria"]

    cat = np.select(
        [
            (t == "Normal") & ~ver,
            (t == "Normal") & ver,
            (t == "Especial") & ~ver,
            (t == "Especial") & ver,
            (t == "Sin convenio") & ~ver,
            (t == "Sin convenio") & ver,
            t == "Visiting",
            t == "Doble Grado",
            t == "Verano en el extranjero",
            t.isin(["LSGL", "Verano LSGL"]),
        ],
        [
            "Convenio semestral", "Convenio verano", "Especial semestral", "Especial verano",
            "Sin convenio semestral", "Sin convenio verano", "Visiting", "Doble grado",
            "Programa de verano en el extranjero", "LSGL",
        ],
        default="Sin clasificar",
    )
    d["categoria"] = cat

    # ¿Con convenio?  [COMPROBADO] Reporte Anual usa "*al" = Normal + Especial.
    d["con_convenio"] = t.isin(["Normal", "Especial"])
    # ¿Paga? [COMPROBADO wishlist] visiting / sin convenio = pagan, no cuentan en balance.
    d["paga"] = t.isin(["Visiting", "Sin convenio"])

    # Balance  [COMPROBADO] sólo Tipo = Normal; verano pesa 0.5
    d["cuenta_balance"] = t.isin(C.BALANCE_TIPOS_CUENTAN)
    if excluir_cancelados:
        d["cuenta_balance"] &= ~d["flag_cancelado"]
    d["peso_balance"] = np.where(ver, C.BALANCE_PESO_VERANO,
                                 np.where(tri, C.BALANCE_PESO_TRIMESTRE, 1.0))
    d["signo"] = np.where(d["flujo"] == "OUT", 1, -1)
    d["aporte_balance"] = np.where(d["cuenta_balance"], d["signo"] * d["peso_balance"], 0.0)

    # Plazo PIM.  [AMBIGÜEDAD] el Excel no lo define con fórmula; propuesta: corto = verano.
    d["plazo"] = np.where(ver | t.isin(["Verano en el extranjero", "Verano LSGL", "LSGL"]), "Corto", "Largo")

    # Duración de la estancia (por alumno-universidad-ciclo)
    grp = d.groupby(["clave_unica", "flujo", "uni_key", "ciclo"])
    n_sem = grp["termino_principal"].transform(lambda s: s[s != "Verano"].nunique())
    d["duracion"] = np.select(
        [ver, d["nivel"].eq("Maestría") & tri, n_sem >= 2],
        ["Corto plazo / verano", "Trimestre(s) de maestría", "Un año"],
        default="Un semestre",
    )
    return d


def carrera_bridge(df: pd.DataFrame, componentes: pd.DataFrame) -> pd.DataFrame:
    """Tabla puente registro -> programa componente.
    Un registro de carrera conjunta genera 2 filas; uno simple, 1 fila.
    Usar SOLO para reportes 'por carrera'. Para totales institucionales usar
    la tabla de movilidad original (1 fila por registro)."""
    comp = componentes.set_index("clave")["componentes"]
    b = df[["fuente_fila", "clave_unica", "flujo", "ciclo", "nivel", "tipo", "categoria",
            "clave_carrera", "carrera_principal", "pais", "universidad_canon"]].copy()
    b["programa"] = b["clave_carrera"].map(comp).fillna(b["clave_carrera"])
    b["programa"] = b["programa"].str.split("|")
    b = b.explode("programa")
    b["es_conjunta"] = b.groupby("fuente_fila")["programa"].transform("size") > 1
    return b


def alumnos_dim(df: pd.DataFrame) -> pd.DataFrame:
    """Dimensión alumno (una fila por Clave Única)."""
    g = df.sort_values(["anio_orig"]).groupby("clave_unica")
    return pd.DataFrame({
        "alumno": g["alumno"].agg(lambda s: s.value_counts().index[0]),
        "variantes_nombre": g["alumno_key"].nunique(),
        "sexo": g["sexo"].agg(lambda s: s.dropna().iloc[-1] if s.notna().any() else None),
        "nacionalidad": g["nacionalidad"].agg(lambda s: s.dropna().iloc[-1] if s.notna().any() else None),
        "fecha_nacimiento": g["fecha_nacimiento"].max(),
        "flujos": g["flujo"].agg(lambda s: "|".join(sorted(set(s)))),
        "n_registros": g.size(),
        "n_periodos": g["periodo_cod"].nunique(),
        "primer_ciclo": g["ciclo"].min(),
        "ultimo_ciclo": g["ciclo"].max(),
    }).reset_index()
