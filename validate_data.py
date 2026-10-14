"""
Validaciones sobre la tabla limpia.  No corrige: produce un inventario
`issues` (una fila por registro x regla) que se exporta para revisión.

Cada regla declara:
  id, severidad (ALTA/MEDIA/BAJA/INFO), reportes afectados y si la corrección
  puede ser automática ("auto"), semiautomática ("semi") o humana ("humana").
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

import pandas as pd

# Entidades que aparecen en la columna Universidad pero son programas / becas,
# no universidades.  [COMPROBADO] presentes en los datos; [AMBIGÜEDAD] cómo
# deben contarse en "universidades de origen/destino".
NO_UNIVERSIDADES = {"fulbright", "esen", "daad"}


@dataclass
class Rule:
    id: str
    severidad: str
    titulo: str
    reportes: str
    correccion: str
    fn: Callable[[pd.DataFrame], pd.Series]   # devuelve máscara booleana
    detalle: Callable[[pd.DataFrame], pd.Series] | None = None


def _dup_mask(df, keys):
    return df.duplicated(keys, keep=False) & df[keys].notna().all(axis=1)


def _spelling_variants(df):
    n = df.groupby("uni_key")["universidad"].transform("nunique")
    return n > 1


def _uni_multi_country(df):
    g = df.groupby(["uni_key", "pais"]).size().rename("n").reset_index()
    top = g.sort_values("n", ascending=False).drop_duplicates("uni_key").set_index("uni_key")["pais"]
    multi = df.groupby("uni_key")["pais"].transform("nunique") > 1
    return multi & (df["pais"] != df["uni_key"].map(top))


def _cu_multi_name(df):
    return df.groupby("clave_unica")["alumno_key"].transform("nunique") > 1


def _cu_both_flows(df):
    return df.groupby("clave_unica")["flujo"].transform("nunique") > 1


PROG_MAESTRIA = ("MBA", "MTIA", "DR ")

RULES: list[Rule] = [
    Rule("V01", "MEDIA", "Tipo escrito de varias formas ('Sin convenio'/'Sin Convenio')",
         "Reporte anual, Balance, sin convenio", "auto",
         lambda d: d["tipo_orig"] != d["tipo"]),
    Rule("V02", "MEDIA", "Clave Única guardada como número (perdió ceros a la izquierda)",
         "Alumnos únicos, repetidores, cruces con Banner", "auto",
         lambda d: d["cu_era_numero"]),
    Rule("V03", "MEDIA", "Sexo nulo (0 en CRM)", "Patlani, SEP, FIMPES, equidad", "humana (CRM/Banner)",
         lambda d: d["sexo"].isna()),
    Rule("V04", "BAJA", "Nacionalidad nula o '0'", "Patlani-Países, 911, listados", "humana (CRM/Banner)",
         lambda d: d["nacionalidad"].isna()),
    Rule("V05", "BAJA", "Fecha de nacimiento con error #N/A pegado como valor o vacía",
         "Edad (FIMPES)", "humana (CRM/Banner)",
         lambda d: d["fecha_nacimiento"].isna()),
    Rule("V06", "ALTA", "Clave de carrera vacía", "Todos los reportes por carrera, Patlani campo", "humana",
         lambda d: d["clave_carrera"].isna()),
    Rule("V07", "ALTA", "Año del Periodo (texto CRM) ≠ columna año", "Cualquier corte por ciclo/año",
         "humana", lambda d: d["anio_periodo"].notna() & (d["anio_periodo"] != d["anio_orig"]),
         lambda d: d["periodo_orig"].astype(str) + " | año=" + d["anio_orig"].astype(str) + " | ciclo=" + d["ciclo_orig"].astype(str)),
    Rule("V08", "ALTA", "Ciclo escolar no corresponde a año+semestre", "Todos los reportes por ciclo", "semi",
         lambda d: d["ciclo_calc"] != d["ciclo_orig"],
         lambda d: "ciclo=" + d["ciclo_orig"].astype(str) + " esperado=" + d["ciclo_calc"].astype(str)),
    Rule("V09", "MEDIA", "Nivel ≠ nivel del Periodo (Undergraduate/Graduate)", "Reportes por nivel, PIM, Patlani",
         "humana", lambda d: d["nivel_periodo"].notna() & (d["nivel_periodo"] != d["nivel"])),
    Rule("V10", "MEDIA", "Nivel inconsistente con el programa (MBA/MTIA en licenciatura o viceversa)",
         "Reportes por nivel y carrera", "humana",
         lambda d: (d["clave_carrera"].fillna("").str.startswith(PROG_MAESTRIA) & (d["nivel"] == "Licenciatura")) |
                   (~d["clave_carrera"].fillna("").str.startswith(PROG_MAESTRIA) & d["clave_carrera"].notna()
                    & (d["nivel"] == "Maestría") & (d["flujo"] == "OUT"))),
    Rule("V11", "ALTA", "Misma universidad escrita de varias formas (mayúsculas/espacios/marca sin convenio)",
         "Balance, top universidades, nuevos destinos", "auto",
         _spelling_variants, lambda d: d["universidad"]),
    Rule("V12", "ALTA", "Universidad no encontrada en unis-info (PIM/Balance indefinido)",
         "PIM, Balance", "humana (alta en catálogo)",
         lambda d: d["pim"].isna(), lambda d: d["universidad"]),
    Rule("V13", "ALTA", "Marca '(Sin Convenio)' en el nombre contradice el Tipo",
         "Balance, sin convenio, reporte anual", "humana",
         lambda d: (d["marca_sin_convenio"] & d["tipo"].isin(["Normal", "Especial"])) |
                   (~d["marca_sin_convenio"] & (d["tipo"] == "Sin convenio")),
         lambda d: d["tipo"] + " | " + d["universidad"].astype(str)),
    Rule("V14", "MEDIA", "Universidad registrada con un país distinto al habitual", "Reportes por país/región, PIM región",
         "semi", _uni_multi_country, lambda d: d["universidad"].astype(str) + " -> " + d["pais"].astype(str)),
    Rule("V15", "MEDIA", "País sin región en país-región", "Región, PIM", "auto (catálogo)",
         lambda d: d["region"].isna(), lambda d: d["pais"]),
    Rule("V16", "MEDIA", "Misma Clave Única con nombres distintos", "Alumnos únicos, listados", "semi",
         _cu_multi_name, lambda d: d["alumno"]),
    Rule("V17", "ALTA", "Registro duplicado (misma CU, flujo, periodo y universidad)",
         "Todos los conteos", "semi",
         lambda d: _dup_mask(d, ["clave_unica", "flujo", "periodo_cod", "uni_key"]),
         lambda d: d["tipo"] + " | " + d["universidad"].astype(str)),
    Rule("V18", "INFO", "Mismo alumno, mismo periodo, dos registros (p.ej. dos cursos de verano)",
         "Conteo alumnos vs registros", "regla de negocio",
         lambda d: _dup_mask(d, ["clave_unica", "flujo", "periodo_cod"]) & ~_dup_mask(d, ["clave_unica", "flujo", "periodo_cod", "uni_key"]),
         lambda d: d["tipo"] + " | " + d["universidad"].astype(str)),
    Rule("V19", "ALTA", "Comentario dice CANCELÓ pero el registro sigue en la base (se cuenta)",
         "Todos los conteos", "humana", lambda d: d["flag_cancelado"], lambda d: d["comentario"]),
    Rule("V20", "INFO", "Comentario documenta diferencia con el CRM", "Conciliación CRM", "humana",
         lambda d: d["comentario"].notna() & ~d["flag_cancelado"] & (d["comentario"] != "Tomó 2 cursos"),
         lambda d: d["comentario"]),
    Rule("V21", "INFO", "Clave Única aparece como incoming y outgoing", "Alumnos únicos", "ninguna",
         _cu_both_flows),
    Rule("V22", "MEDIA", "La columna Universidad contiene un programa/beca, no una universidad (Fulbright, ESEN, DAAD)",
         "Top universidades, PIM, convenios", "humana (definir regla)",
         lambda d: d["uni_key"].isin(NO_UNIVERSIDADES), lambda d: d["universidad"]),
    Rule("V23", "MEDIA", "Valor cacheado en Excel ≠ valor recalculado desde catálogo (PIM / región / carrera principal)",
         "PIM, región, SEP", "auto",
         lambda d: ((d["xl_pim"].map({"Sí": True, "No": False}) != d["pim"]) & d["pim"].notna()) |
                   ((d["xl_region"] != d["region"]) & d["region"].notna()) |
                   ((d["xl_carrera_principal"].astype(str).str.strip() != d["carrera_principal"].astype(str)) & d["carrera_principal"].notna()),
         lambda d: "xl_pim=" + d["xl_pim"].astype(str) + " calc=" + d["pim"].astype(str) +
                   " | xl_region=" + d["xl_region"].astype(str) + " calc=" + d["region"].astype(str)),
]


def run_validations(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    out = []
    for r in RULES:
        mask = r.fn(df).fillna(False).astype(bool)
        if not mask.any():
            continue
        sub = df.loc[mask, ["fuente_fila", "flujo", "tipo", "clave_unica", "alumno", "periodo_orig", "universidad"]].copy()
        sub["regla"] = r.id
        sub["severidad"] = r.severidad
        sub["problema"] = r.titulo
        sub["reportes_afectados"] = r.reportes
        sub["correccion"] = r.correccion
        sub["detalle"] = r.detalle(df).loc[mask].astype(str) if r.detalle else ""
        out.append(sub)
    issues = pd.concat(out, ignore_index=True) if out else pd.DataFrame()
    resumen = (issues.groupby(["regla", "severidad", "problema", "reportes_afectados", "correccion"])
               .agg(registros=("fuente_fila", "nunique")).reset_index()
               .sort_values(["severidad", "regla"], key=lambda s: s.map({"ALTA": 0, "MEDIA": 1, "BAJA": 2, "INFO": 3}) if s.name == "severidad" else s))
    resumen["pct_base"] = (resumen["registros"] / len(df) * 100).round(2)
    return issues, resumen
