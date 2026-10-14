"""
Limpieza y normalización de In-Out (Todos).

Entrada : DataFrame crudo de load_data (valores cacheados del Excel).
Salida  : DataFrame `movilidad` limpio, una fila por registro de movilidad,
          con columnas originales preservadas (sufijo _orig) para auditoría.

Reglas:
  * Nunca se borra un registro aquí. Los problemas se MARCAN (validate_data)
    y las exclusiones se deciden en business_rules / reports.
  * Toda corrección automática deja rastro en la columna `_correcciones`.
"""
from __future__ import annotations

import re
import unicodedata

import numpy as np
import pandas as pd

import config as C


# ---------------------------------------------------------------------------
# utilidades de texto
# ---------------------------------------------------------------------------
def collapse_spaces(s):
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return None
    return re.sub(r"\s+", " ", str(s)).strip()


def strip_accents(s: str) -> str:
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()


def norm_key(s) -> str | None:
    """Llave de comparación: sin acentos, minúsculas, espacios colapsados."""
    s = collapse_spaces(s)
    if s is None:
        return None
    return strip_accents(s).casefold()


SIN_CONV_RE = re.compile(r"\s*\(\s*sin\s+convenio\s*\)\s*", re.IGNORECASE)


def uni_key(name) -> str | None:
    """Llave de universidad: ignora mayúsculas, espacios, acentos y la marca
    '(Sin Convenio)'.  [COMPROBADO] el Balance del Excel ya agrupa con y sin
    marca a propósito (comentario en Balance!B4)."""
    s = collapse_spaces(name)
    if s is None:
        return None
    s = SIN_CONV_RE.sub(" ", s)
    return norm_key(s)


# ---------------------------------------------------------------------------
# catálogos
# ---------------------------------------------------------------------------
def build_catalogs(src) -> dict:
    unis = src.unis.copy()
    unis["uni_nombre"] = unis["Universidad"].map(collapse_spaces)
    unis["lookup_key"] = unis["Universidad"].map(lambda x: collapse_spaces(x).casefold())
    unis["uni_key"] = unis["Universidad"].map(uni_key)

    car = src.carreras.copy()
    car["Clave"] = car["Clave"].map(collapse_spaces)
    car["Carrera Principal"] = car["Carrera Principal"].map(collapse_spaces)

    pais = src.pais_region.copy()
    pais["PAÍS"] = pais["PAÍS"].map(lambda x: collapse_spaces(x).upper())
    return {"unis": unis, "carreras": car, "pais": pais}


def build_carrera_componentes(carreras: pd.DataFrame) -> pd.DataFrame:
    """[PROPUESTA] Descompone cada clave en sus programas componentes.
    Una carrera conjunta 'ECO DFI' -> ECO | DFI.  Las claves cuyo token no
    existe como programa individual se marcan para revisión humana."""
    standalone = set(carreras["Clave"])
    rows = []
    for clave in carreras["Clave"]:
        nota, revisar = "", False
        if clave in C.PROGRAMAS_MULTIPALABRA or " " not in clave:
            comps = [clave]
        else:
            toks = clave.split(" ")
            if all(t in standalone for t in toks):
                comps = toks
                if len(toks) > 2:
                    revisar, nota = True, "3 componentes: confirmar si es triple o conjunta de 2"
            else:
                faltan = [t for t in toks if t not in standalone]
                revisar = True
                nota = f"token(s) {faltan} no existen como programa individual en clave-carrera"
                # propuesta conservadora: [primer programa, resto] sólo si el resto
                # también es una clave del catálogo; si no, la clave completa es 1 programa
                resto = " ".join(toks[1:])
                comps = [toks[0], resto] if (toks[0] in standalone and resto in standalone) else [clave]
        rows.append({"clave": clave, "componentes": "|".join(comps),
                     "n_componentes": len(comps), "es_conjunta": len(comps) > 1,
                     "requiere_validacion": revisar, "nota": nota})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# limpieza principal
# ---------------------------------------------------------------------------
PERIODO_RE = re.compile(
    r"^(FALL|SPRING|SUMMER|JANUARY-MARCH|APRIL-JUNE|JULY-SEPTEMBER|OCTOBER-DECEMBER)\s+(\d{4})-(UNDERGRADUATE|GRADUATE)",
    re.IGNORECASE)


def _ciclo(anio: int, termino: str) -> str:
    if termino in C.TERMINOS_INICIO_CICLO:
        return f"{anio}-{anio + 1}"
    return f"{anio - 1}-{anio}"


def clean_inout(raw: pd.DataFrame, cats: dict, aliases: pd.DataFrame | None = None) -> pd.DataFrame:
    df = pd.DataFrame(index=raw.index)
    corr = [[] for _ in range(len(raw))]

    df["fuente_fila"] = raw["_fuente_fila"].astype(int)

    # --- Flujo / Tipo -------------------------------------------------------
    df["flujo"] = raw["Flujo"].map(collapse_spaces).str.upper().map({"INCOMING": "IN", "OUTGOING": "OUT"})
    df["tipo_orig"] = raw["Tipo"]
    df["tipo"] = raw["Tipo"].map(lambda x: C.TIPO_CANON.get(norm_key(x), collapse_spaces(x)))
    for i, (a, b) in enumerate(zip(df["tipo_orig"], df["tipo"])):
        if a != b:
            corr[i].append(f"tipo:{a!r}->{b!r}")

    # --- Alumno / Clave Única ----------------------------------------------
    df["alumno"] = raw["Alumno"].map(collapse_spaces)
    df["alumno_key"] = df["alumno"].map(lambda x: norm_key(x).upper() if x else None)
    cu_raw = raw["Clave Única"]
    df["cu_era_numero"] = cu_raw.map(lambda x: isinstance(x, (int, float)) and not pd.isna(x))
    df["clave_unica"] = cu_raw.map(lambda x: None if x is None or str(x).strip() in ("", "0")
                                   else str(int(x) if isinstance(x, float) else x).strip().zfill(9))
    for i, f in enumerate(df["cu_era_numero"]):
        if f:
            corr[i].append("clave_unica: número->texto con ceros a la izquierda")

    # --- Periodo ------------------------------------------------------------
    df["periodo_orig"] = raw["Periodo"]
    m = raw["Periodo"].astype(str).str.strip().str.extract(PERIODO_RE)
    df["termino"] = m[0].str.upper().map(C.PERIODO_TERM)
    df["anio_periodo"] = pd.to_numeric(m[1], errors="coerce").astype("Int64")
    df["nivel_periodo"] = m[2].str.upper().map({"UNDERGRADUATE": "Licenciatura", "GRADUATE": "Maestría"})
    df["semestre_orig"] = raw["Semestre"].map(collapse_spaces)
    df["anio_orig"] = pd.to_numeric(raw["año"], errors="coerce").astype("Int64")
    df["ciclo_orig"] = raw["ciclo escolar"].map(collapse_spaces)
    df["termino_principal"] = df["semestre_orig"].map(C.SEMESTRE_PRINCIPAL)
    df["es_verano"] = df["semestre_orig"].eq("Verano")
    df["es_trimestre_maestria"] = df["semestre_orig"].isin(
        ["Enero-Marzo", "Abril-Junio", "Julio-Septiembre", "Octubre-Diciembre"])
    # [PROPUESTA] ciclo calculado a partir de (año, semestre) del propio registro.
    df["ciclo_calc"] = [(_ciclo(int(a), s) if pd.notna(a) and s else None)
                        for a, s in zip(df["anio_orig"], df["semestre_orig"])]
    # [PROPUESTA] ciclo "oficial" del pipeline = el del Excel (para no mover
    # cifras históricas); las discrepancias se reportan en validate_data.
    df["ciclo"] = df["ciclo_orig"]
    df["periodo_cod"] = df["anio_orig"].astype(str) + "-" + df["semestre_orig"].fillna("")

    # --- Nivel ----------------------------------------------------------------
    df["nivel"] = raw["Nivel"].map(collapse_spaces)

    # --- Carrera (recalculada desde catálogo) ---------------------------------
    car = cats["carreras"].set_index("Clave")
    df["clave_carrera"] = raw["clave carrera"].map(collapse_spaces)
    df["carrera_nombre"] = df["clave_carrera"].map(car["Carrera"])
    df["carrera_principal"] = df["clave_carrera"].map(car["Carrera Principal"])
    df["campo_patlani"] = df["clave_carrera"].map(car["Campo de estudios Patlani"])

    # --- País / región --------------------------------------------------------
    df["pais"] = raw["País"].map(lambda x: collapse_spaces(x).upper() if x else None)
    reg = cats["pais"].set_index("PAÍS")["REGIÓN PIM"]
    df["region"] = df["pais"].map(reg)

    # --- Universidad ------------------------------------------------------------
    df["universidad_orig"] = raw["Universidad"]
    df["universidad"] = raw["Universidad"].map(collapse_spaces)
    if aliases is not None and len(aliases):
        a = aliases[aliases["aplicar"].astype(str).str.lower().isin(["sí", "si", "true", "1"])]
        amap = dict(zip(a["nombre_en_datos"].map(collapse_spaces), a["nombre_canonico"].map(collapse_spaces)))
        df["universidad"] = df["universidad"].map(lambda x: amap.get(x, x))
    df["marca_sin_convenio"] = df["universidad"].fillna("").str.contains(SIN_CONV_RE)
    df["uni_key"] = df["universidad"].map(uni_key)
    # nombre canónico de despliegue = grafía de unis-info si existe, si no la más frecuente
    unis = cats["unis"]
    pref = unis.drop_duplicates("uni_key").set_index("uni_key")["uni_nombre"].map(lambda s: SIN_CONV_RE.sub(" ", s).strip())
    freq = (df.assign(base=df["universidad"].map(lambda s: SIN_CONV_RE.sub(" ", s).strip() if s else s))
              .groupby("uni_key")["base"].agg(lambda s: s.value_counts().index[0]))
    df["universidad_canon"] = df["uni_key"].map(pref).fillna(df["uni_key"].map(freq))
    # PIM / Balance desde unis-info: igual que VLOOKUP (exacto, sin distinguir mayúsculas)
    lk = unis.drop_duplicates("lookup_key").set_index("lookup_key")
    key = df["universidad"].map(lambda x: x.casefold() if x else None)
    df["pim"] = key.map(lk["PIM"]).map({"Sí": True, "No": False})
    df["unis_info_balance"] = key.map(lk["Balance"])
    # [PROPUESTA] si no hay match exacto, heredar por uni_key (p.ej. variante "(sin convenio)")
    by_key = unis.drop_duplicates("uni_key").set_index("uni_key")
    miss = df["pim"].isna()
    df.loc[miss, "pim"] = df.loc[miss, "uni_key"].map(by_key["PIM"]).map({"Sí": True, "No": False})
    df["pim_heredado_por_llave"] = miss & df["pim"].notna()

    # --- Sexo / nacionalidad / nacimiento ---------------------------------------
    df["sexo_orig"] = raw["Sexo"]
    df["sexo"] = raw["Sexo"].map(lambda x: C.SEXO_CANON.get(str(x).strip().lower()) if x is not None else None)
    df["nacionalidad"] = raw["Nacionalidad"].map(
        lambda x: None if x is None or str(x).strip() in ("0", "") else collapse_spaces(x).upper())
    fn = raw["Fecha de Nacimiento"]
    df["fecha_nac_error_excel"] = fn.astype(str).str.startswith("#")
    df["fecha_nacimiento"] = pd.to_datetime(fn.where(~df["fecha_nac_error_excel"]), errors="coerce")
    mes = df["semestre_orig"].map(C.MES_INICIO_TERMINO)
    inicio = pd.to_datetime(dict(year=df["anio_orig"].astype("float"), month=mes, day=1), errors="coerce")
    bm, bd = df["fecha_nacimiento"].dt.month, df["fecha_nacimiento"].dt.day
    aun_no = (inicio.dt.month < bm) | ((inicio.dt.month == bm) & (bd > 1))
    edad = inicio.dt.year - df["fecha_nacimiento"].dt.year - aun_no.astype(int)
    df["edad_inicio"] = edad.where(df["fecha_nacimiento"].notna()).astype("Int64")

    # --- Comentario ---------------------------------------------------------------
    df["comentario"] = raw["Comentario"].map(collapse_spaces)
    df["flag_cancelado"] = df["comentario"].fillna("").str.upper().str.contains("CANCEL") & \
        ~df["comentario"].fillna("").str.contains("sí vino", case=False)

    # --- Valores cacheados del Excel (sólo para comparar) -----------------------
    for c_excel, c_new in [("Carrera Principal", "xl_carrera_principal"), ("PIM", "xl_pim"),
                           ("Región (PIM)", "xl_region"), ("Edad", "xl_edad")]:
        df[c_new] = raw[c_excel]

    df["_correcciones"] = [";".join(x) for x in corr]
    return df


def load_aliases(path) -> pd.DataFrame | None:
    try:
        return pd.read_csv(path, dtype=str)
    except FileNotFoundError:
        return None
