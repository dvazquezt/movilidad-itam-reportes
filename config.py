"""
Configuración central del pipeline de Movilidad Internacional (ITAM).

Todo lo que es "conocimiento del negocio" que hoy vive en fórmulas o en la
cabeza de quien opera la macro debe quedar aquí o en /catalogos, no regado
en el código.

Convención usada en los comentarios:
  [COMPROBADO]  = comportamiento observado en el Excel "datos_macro_actual al 6-10-2026.xlsx"
  [PROPUESTA]   = decisión de diseño de este pipeline (debe validarla el área)
  [AMBIGÜEDAD]  = no se puede decidir con los archivos; requiere a una persona
"""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
CATALOGOS_DIR = BASE_DIR / "catalogos"

# ---------------------------------------------------------------------------
# Hojas del libro fuente
# ---------------------------------------------------------------------------
SHEET_INOUT = "In-Out (Todos)"
SHEET_UNIS = "unis-info"
SHEET_CARRERAS = "clave-carrera"
SHEET_PAIS_REGION = "país-región"
SHEET_BALANCE = "Balance"
SHEET_REPORTE_ANUAL = "Reporte Anual"
SHEET_PIM = "PIM"

# [COMPROBADO] Encabezados de la fila 1 de In-Out (Todos), columnas A..W
INOUT_COLUMNS = [
    "Flujo", "Tipo", "Alumno", "clave carrera", "Carrera", "Carrera Principal",
    "Clave Única", "Periodo", "Semestre", "Semestre Principal", "año",
    "ciclo escolar", "Nivel", "País", "PIM", "Región (PIM)", "Campo Patlani",
    "Universidad", "Sexo", "Edad", "Fecha de Nacimiento", "Nacionalidad",
    "Comentario",
]

# [COMPROBADO] Columnas que en el Excel son fórmulas (VLOOKUP / IF / DATEDIF).
# El pipeline NO confía en su valor cacheado: las recalcula desde catálogos.
DERIVED_IN_EXCEL = {
    "Carrera": "VLOOKUP(clave carrera, clave-carrera!A:B, 2)",
    "Carrera Principal": "VLOOKUP(clave carrera, clave-carrera!A:D, 4)",
    "Campo Patlani": "VLOOKUP(clave carrera, clave-carrera!A:C, 3)",
    "Semestre Principal": "IF(Semestre -> Otoño/Primavera/Verano)",
    "PIM": "VLOOKUP(Universidad, unis-info!A:C, 2)",
    "Región (PIM)": "VLOOKUP(País, país-región!A:B, 2)",
    "Edad": "IFNA(DATEDIF(FechaNac, DATE(año, mes_inicio, DAY(TODAY()))),0)",
}

# ---------------------------------------------------------------------------
# Normalización de categorías
# ---------------------------------------------------------------------------
# [COMPROBADO] 9 valores en Tipo; "Sin convenio" y "Sin Convenio" son el mismo.
TIPO_CANON = {
    "normal": "Normal",
    "especial": "Especial",
    "sin convenio": "Sin convenio",
    "visiting": "Visiting",
    "doble grado": "Doble Grado",
    "verano en el extranjero": "Verano en el extranjero",
    "lsgl": "LSGL",
    "verano lsgl": "Verano LSGL",
}

# [COMPROBADO] Periodo (texto del CRM) -> término. 22 variantes de escritura
# (mayúsculas, "(January to May)", etc.) se reducen a 7 términos.
PERIODO_TERM = {
    "FALL": "Otoño",
    "SPRING": "Primavera",
    "SUMMER": "Verano",
    "JANUARY-MARCH": "Enero-Marzo",
    "APRIL-JUNE": "Abril-Junio",
    "JULY-SEPTEMBER": "Julio-Septiembre",
    "OCTOBER-DECEMBER": "Octubre-Diciembre",
}

# [COMPROBADO] Fórmula de la columna J "Semestre Principal".
# Nota: Julio-Septiembre (trimestre de maestría) se agrupa como "Verano".
SEMESTRE_PRINCIPAL = {
    "Otoño": "Otoño", "Octubre-Diciembre": "Otoño",
    "Primavera": "Primavera", "Enero-Marzo": "Primavera", "Abril-Junio": "Primavera",
    "Verano": "Verano", "Julio-Septiembre": "Verano",
}

# [COMPROBADO] Términos que abren ciclo escolar (ciclo = año-año+1).
# El resto pertenece al ciclo (año-1)-año. Regla: el ciclo va de otoño a verano.
TERMINOS_INICIO_CICLO = {"Otoño", "Octubre-Diciembre"}

# Mes de inicio usado para la edad (la fórmula del Excel usa 1, 6 u 8).
# [COMPROBADO] El Excel asigna 8 a TODO lo que no sea Primavera/Verano, incluidos
# Enero-Marzo y Abril-Junio (error). [PROPUESTA] usar el mes real del término.
MES_INICIO_TERMINO = {
    "Primavera": 1, "Verano": 6, "Otoño": 8,
    "Enero-Marzo": 1, "Abril-Junio": 4, "Julio-Septiembre": 7, "Octubre-Diciembre": 10,
}

SEXO_CANON = {"male": "M", "female": "F"}  # [COMPROBADO] "0" = nulo en CRM (comentario S1)

# Claves de programa de varias palabras que NO son carreras conjuntas.
PROGRAMAS_MULTIPALABRA = {"MBA ADM", "MTIA CDS", "MTIA COM", "MTIA DER", "MTIA ECO",
                          "MTIA FIN", "MTIA POLPUB", "MTIA TEO ECO", "TIC ADM", "DR ECO"}

# ---------------------------------------------------------------------------
# Reglas de balance
# ---------------------------------------------------------------------------
BALANCE_PESO_VERANO = 0.5           # [COMPROBADO] wishlist + fórmula =D-C+0.5*(F-E)
BALANCE_TIPOS_CUENTAN = {"Normal"}  # [COMPROBADO] COUNTIFS(... B:B,"Normal" ...)
# [AMBIGÜEDAD] Cada trimestre de maestría pesa 1 (como un semestre) en el Excel,
# porque sólo se filtra Semestre <> "Verano". ¿Es lo que se quiere?
BALANCE_PESO_TRIMESTRE = 1.0

# Tolerancia de comparación Excel vs Python (conteos enteros => 0)
PARITY_TOLERANCE = 0.0
