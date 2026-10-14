# movilidad-itam-reportes

Capa de limpieza, validación y reporteo en Python para la Dirección de Movilidad Internacional del ITAM.
Lee el Excel histórico (la hoja **In-Out (Todos)**), sin modificarlo, y genera reportes nuevos y reproducibles.

> Versión 0.1 (borrador). El diagnóstico técnico completo está en [`docs/diagnostico_tecnico.md`](docs/diagnostico_tecnico.md).

---

## Cómo correrlo (paso a paso)

### 1. Instalar Python (sólo la primera vez)
- **Windows:** descarga Python 3.11 o más reciente de <https://www.python.org/downloads/>.
  En el instalador marca la casilla **"Add python.exe to PATH"** antes de dar *Install Now*.
- **Mac:** descárgalo de la misma página, o usa `brew install python`.

Para comprobarlo, abre una terminal (Windows: *Símbolo del sistema* o *PowerShell*; Mac: *Terminal*) y escribe:
```
python --version      (en Mac: python3 --version)
```
Debe aparecer `Python 3.11` o mayor.

### 2. Poner el Excel en la carpeta `datos/`
Copia el Excel exportado (por ejemplo `datos_macro_actual al 6-10-2026.xlsx`) dentro de `datos/`.
Debe haber **un solo** `.xlsx` ahí. El programa sólo lo lee.

### 3. Correr
- **Windows:** doble clic en **`correr_windows.bat`**.
- **Mac / Linux:** en la terminal, dentro de esta carpeta: `./correr_mac_linux.sh`

La primera vez crea un entorno `.venv` e instala `pandas`, `openpyxl` y `numpy` (necesita internet, ~1–2 min).
Después cada corrida tarda alrededor de 1 minuto.

### 4. Ver los resultados en `salidas/`
| Archivo | Qué contiene |
|---|---|
| `reportes_mvp_<ciclo>_<fecha>.xlsx` | 1) Balance por universidad + histórico · 2) Outgoing por carrera (principal vs. conjuntas) + carrera × país · 3) PIM |
| `paridad_excel_python_<fecha>.xlsx` | Réplica de las fórmulas del Excel actual y explicación de cada diferencia del balance corregido |
| `diagnostico_calidad_<fecha>.xlsx` | Problemas de datos por regla y por registro (con la fila del Excel), fórmulas pegadas a mano, carreras conjuntas, alumnos |
| `movilidad_limpia_<fecha>.csv` | La tabla limpia, una fila por registro, con `fuente_fila` para rastrearla en el Excel |

Cada libro trae una hoja **LEEME** con el archivo fuente, su huella SHA-256 y los parámetros usados.

### Opciones
Desde la terminal (con el entorno activado, o agregándolas al final del `.bat`/`.sh`):
```
python run_pipeline.py --ciclo 2025-2026            # ciclo final (por defecto, el de Balance!C1)
python run_pipeline.py --ciclo 2025-2026 --anios 3  # ventana del balance
python run_pipeline.py --input "C:\ruta\otro.xlsx"  # usar otro archivo
python run_pipeline.py --sin-formulas               # más rápido, sin auditar fórmulas
python inspect_workbook.py "datos/archivo.xlsx" --out salidas/mapa_hojas.xlsx   # mapa de hojas y tablas dinámicas
python tests/test_parity.py                         # verifica que la réplica coincide 100 % con el Excel
```
Para activar el entorno a mano: Windows `.venv\Scripts\activate` · Mac/Linux `source .venv/bin/activate`.

### Problemas comunes
| Mensaje | Solución |
|---|---|
| `'py' no se reconoce…` / `python: command not found` | Reinstala Python marcando "Add to PATH", o cierra y vuelve a abrir la terminal |
| `Pon exactamente UN archivo .xlsx en datos/` | Deja un solo Excel en `datos/` (cierra Excel: los temporales `~$…` se ignoran) |
| `Faltan columnas esperadas en In-Out (Todos)` | Cambió la estructura del Excel; revisa los encabezados contra `config.INOUT_COLUMNS` |
| `PermissionError` al escribir en `salidas/` | Cierra el reporte anterior si está abierto en Excel |

---

## Cómo está organizado

```
Excel ─► load_data ─► clean_data ─► validate_data ─► business_rules ─► reports ─► export_excel
                                 └─► excel_parity (réplica exacta de los COUNTIFS actuales)
```

| Archivo | Responsabilidad |
|---|---|
| `config.py` | Hojas, columnas esperadas, mapeos y parámetros de balance (todas las reglas con nombre) |
| `load_data.py` | Única puerta de entrada (sólo lectura). Para Dataverse se agrega aquí `load_from_dataverse()` con el mismo esquema |
| `clean_data.py` | Normaliza textos, Clave Única, periodo; recalcula carrera, PIM y región desde catálogos |
| `validate_data.py` | 23 reglas de calidad con severidad, reportes afectados y tipo de corrección |
| `business_rules.py` | Categorías de movilidad, balance (verano = 0.5), plazo, duración, puente de carreras conjuntas |
| `excel_parity.py` | Emula COUNTIFS para reproducir Balance, Reporte Anual y PIM tal como hoy |
| `reports.py` | Los 3 reportes del MVP en versión corregida |
| `export_excel.py` | Escribe libros nuevos con fecha |
| `run_pipeline.py` | Orquesta todo |
| `catalogos/carrera_componentes.csv` | Qué claves son carreras conjuntas. **Revisar filas con `requiere_validacion = True`** |
| `catalogos/alias_universidades.csv` | Grafías que son la misma universidad (sólo se aplican filas con `aplicar = sí`) |

Convenciones en el código: `[COMPROBADO]` observado en el Excel · `[PROPUESTA]` decisión de diseño · `[AMBIGÜEDAD]` decisión pendiente del área.

## Datos personales
`datos/` y `salidas/` están excluidas por `.gitignore`: contienen nombres y claves únicas de alumnos.
**No subas Excel ni CSV de datos a este repositorio.** Los únicos CSV versionados son los catálogos.
