# Reportes de Movilidad Internacional (ITAM)

Este programa toma el Excel que el área exporta del CRM y genera automáticamente los reportes de movilidad: **Balance de convenios**, **Outgoing por carrera** y **Red PIM**. También entrega una lista de los problemas de datos que encuentre.

- **No modifica el Excel original:** solo lo lee y guarda los resultados en archivos nuevos.
- **Tarda alrededor de un minuto.**
- **No hace falta saber programar** para usarlo.

> **¿Quieres entender primero qué hace y por qué?** Lee los documentos «Qué hace el programa de reportes de Movilidad» y «Diagnóstico y pasos a seguir». Esta guía explica cómo **usarlo**.

---

## Índice
1. [Inicio rápido](#1-inicio-rápido)
2. [Instalación (solo la primera vez)](#2-instalación-solo-la-primera-vez)
3. [Cómo generar los reportes cada ciclo](#3-cómo-generar-los-reportes-cada-ciclo)
4. [Qué archivos obtienes](#4-qué-archivos-obtienes)
5. [Cómo leer las cifras](#5-cómo-leer-las-cifras)
6. [Si algo sale mal](#6-si-algo-sale-mal)
7. [Opciones adicionales](#7-opciones-adicionales)
8. [Cómo mantenerlo actualizado](#8-cómo-mantenerlo-actualizado)
9. [Qué hace por dentro (para quien programa)](#9-qué-hace-por-dentro-para-quien-programa)
10. [Datos personales](#10-datos-personales)

---

## 1. Inicio rápido

Si Python ya está instalado:

1. Copia el Excel exportado del CRM a la carpeta **`datos/`**.
2. **Windows:** doble clic en **`correr_windows.bat`**. **Mac:** ejecuta `./correr_mac_linux.sh` en la Terminal.
3. Abre la carpeta **`salidas/`**: ahí están los reportes.

Si es la primera vez, sigue la sección 2.

---

## 2. Instalación (solo la primera vez)

### Paso 1. Descargar el programa
- **Sin usar git:** en la página del repositorio en GitHub, botón verde **Code → Download ZIP**. Descomprime el archivo en una carpeta fácil de encontrar, por ejemplo `Documentos/movilidad-itam-reportes`.
- **Con git:** `git clone https://github.com/dvazquezt/movilidad-itam-reportes.git`

### Paso 2. Instalar Python
Necesitas **Python 3.11 o más reciente**. Es gratuito.

**En Windows**
1. Entra a <https://www.python.org/downloads/> y descarga la versión más reciente.
2. Abre el instalador. **Antes de dar clic en *Install Now*, marca la casilla «Add python.exe to PATH»** (abajo, en la primera pantalla). Si no la marcas, el programa no encontrará Python.
3. Termina la instalación.

**En Mac**
1. Entra a <https://www.python.org/downloads/>, descarga el instalador para macOS y ábrelo.
2. Sigue los pasos del instalador.

### Paso 3. Comprobar que quedó instalado
Abre una terminal:
- **Windows:** tecla Inicio, escribe `cmd` y presiona Enter.
- **Mac:** Cmd + Espacio, escribe `Terminal` y presiona Enter.

Escribe esto y presiona Enter:
```
python --version
```
En Mac escribe `python3 --version`. Debe aparecer algo como `Python 3.12.4`. Si aparece un error, ve a la sección 6.

**Listo.** Las librerías que necesita el programa (pandas, openpyxl y numpy) se instalan solas la primera vez que lo corres. Para eso necesitas internet.

---

## 3. Cómo generar los reportes cada ciclo

### Paso 1. Exporta el Excel del CRM
Exporta el Excel como siempre. Debe tener la hoja **In-Out (Todos)** con sus columnas de siempre, y las hojas **unis-info**, **clave-carrera** y **país-región**.

### Paso 2. Ponlo en la carpeta `datos/`
- Copia el archivo dentro de `datos/`.
- **Quita el Excel anterior.** Solo puede haber un `.xlsx` en esa carpeta.
- El nombre del archivo no importa.

### Paso 3. Corre el programa
- **Windows:** doble clic en `correr_windows.bat`. Se abre una ventana negra que muestra el avance; al final dice **«Listo»**. Presiona cualquier tecla para cerrarla.
- **Mac:** abre la Terminal y escribe `cd ` (con un espacio al final). Arrastra la carpeta del programa a la ventana y presiona Enter. Luego escribe `./correr_mac_linux.sh` y presiona Enter.

La primera vez tarda 1–2 minutos más porque instala las librerías. Después, alrededor de un minuto.

Al terminar verás algo así:
```
Registros: 7,760 | alumnos únicos: 6,474
Paridad Balance: 332/332 filas idénticas
Paridad Reporte Anual: 14/14 celdas idénticas
Paridad PIM: 56/56 celdas idénticas
```
Las líneas de **«Paridad»** indican si el programa reproduce los números que hoy da el Excel. Si alguna no dice que todo coincide, revisa la sección 6.

### Paso 4. Revisa los resultados
Abre la carpeta `salidas/`. Este es el orden recomendado:
1. **`diagnostico_calidad_…xlsx`:** problemas de datos. Si hay algo que corregir en el CRM, corrígelo, vuelve a exportar y repite.
2. **`reportes_mvp_…xlsx`:** los reportes.

Cada corrida crea archivos nuevos con la fecha y hora en el nombre, así que nunca se sobrescribe un reporte anterior.

---

## 4. Qué archivos obtienes

| Archivo | Para qué sirve | Qué tiene adentro |
|---|---|---|
| **`reportes_mvp_<ciclo>_<fecha>.xlsx`** | **Los reportes para informar** | `1_balance`: balance por universidad. `1_balance_historico`: balance por ciclo y acumulado. `2_outgoing_carrera`: salientes por carrera. `2_totales`: totales institucionales. `2_carrera_x_pais`: a dónde va cada carrera. `3_pim_resumen`, `3_pim_region`, `3_pim_universidad`: red PIM |
| `diagnostico_calidad_<fecha>.xlsx` | Encontrar y corregir errores de datos | `resumen_reglas`: cuántos registros tiene cada tipo de problema. `detalle_problemas`: cada problema con la **fila del Excel** donde está. Además, fórmulas pegadas a mano, carreras conjuntas y una ficha por alumno |
| `paridad_excel_python_<fecha>.xlsx` | Comprobar las cifras contra el Excel actual | Los números del Excel junto a los del programa. `balance_por_que_difiere` explica cada diferencia del balance corregido |
| `movilidad_limpia_<fecha>.csv` | Otros análisis o un tablero en Power BI | Toda la tabla ya limpia, una fila por registro |

Todos los libros traen una hoja **LEEME** que indica con qué archivo, en qué fecha y con qué parámetros se generaron.

---

## 5. Cómo leer las cifras

- **Registros y alumnos no son lo mismo.** Un *registro* es una fila: un alumno en un periodo y una universidad. Un alumno de maestría que estuvo tres trimestres son **3 registros y 1 alumno**. Los reportes muestran las dos cifras.
- **Carreras conjuntas.** En `2_outgoing_carrera`, un alumno de *Economía y Dirección Financiera* cuenta en **las dos** carreras. Por eso la suma de todas las carreras es mayor que el total. El total real está en `2_totales`, donde cada alumno cuenta una sola vez. La columna `registros_carrera_principal` muestra cómo lo contaba el Excel (solo la primera carrera).
- **Balance.**
  - Se calcula así: `salientes semestrales − entrantes semestrales + 0.5 × (verano saliente − verano entrante)`.
  - Solo cuenta la movilidad por convenio (tipo *Normal*). Visiting, sin convenio, especial y doble grado se muestran en su propia columna, pero no suman al balance.
  - **Positivo** = el ITAM envió más alumnos de los que recibió. **Negativo** = recibió más.
- **Ciclo escolar:** va de otoño a verano. Por ejemplo, 2024-2025 = otoño 2024 + primavera 2025 + verano 2025.
- **PIM:** `pct_pim_registros` es el porcentaje de registros con escuelas PIM; `cumple_meta_20pct_registros` indica si llega al 20 %.

---

## 6. Si algo sale mal

| Lo que ves | Qué significa | Qué hacer |
|---|---|---|
| `"python" no se reconoce como un comando…` o `command not found` | Python no está instalado o no está en el PATH | Reinstala Python y marca **«Add python.exe to PATH»**. Cierra y vuelve a abrir la terminal |
| `Pon exactamente UN archivo .xlsx en …datos` | En `datos/` no hay ningún Excel, o hay más de uno | Deja solo el Excel que quieres usar |
| `Faltan columnas esperadas en In-Out (Todos)` | Cambió la estructura del Excel (una columna nueva o renombrada) | Revisa los encabezados de la hoja: deben ser los de siempre (Flujo, Tipo, Alumno, clave carrera…) |
| `PermissionError` o `Permission denied` | Un reporte anterior está abierto en Excel | Ciérralo y vuelve a correr |
| `No matching distribution found` o errores de `pip` | No hubo internet al instalar las librerías | Conéctate y vuelve a correr |
| Alguna línea de «Paridad» no dice que todo coincide | El Excel cambió su lógica, o alguien cambió una regla del programa | Avisa a quien mantiene el programa antes de publicar cifras |
| En Mac: `permission denied: ./correr_mac_linux.sh` | El archivo perdió el permiso para ejecutarse | Escribe `chmod +x correr_mac_linux.sh` y vuelve a intentar |

Si nada de esto funciona, borra la carpeta oculta `.venv` dentro del programa y vuelve a correrlo. Se reinstala solo.

---

## 7. Opciones adicionales

Por defecto el programa usa el ciclo que tenga la hoja Balance del Excel (celda C1) y una ventana de 5 ciclos para el balance. Para cambiarlo, escribe la opción después del comando.

**Windows:** abre la terminal en la carpeta del programa (clic derecho en la carpeta → «Abrir en Terminal»).
```
correr_windows.bat --ciclo 2025-2026
correr_windows.bat --ciclo 2025-2026 --anios 3
```
**Mac:**
```
./correr_mac_linux.sh --ciclo 2025-2026 --anios 3
```

| Opción | Qué hace | Ejemplo |
|---|---|---|
| `--ciclo` | Ciclo final de los reportes | `--ciclo 2025-2026` |
| `--anios` | Cuántos ciclos hacia atrás incluye el balance | `--anios 5` |
| `--input` | Usar un Excel que no está en `datos/` | `--input "C:\Users\yo\Desktop\export.xlsx"` |
| `--sin-formulas` | Más rápido; omite la revisión de fórmulas pegadas a mano | `--sin-formulas` |

Herramientas extra (con el entorno activado: Windows `.venv\Scripts\activate`, Mac `source .venv/bin/activate`):
```
python inspect_workbook.py "datos/archivo.xlsx" --out salidas/mapa_hojas.xlsx   # mapa de hojas y tablas dinámicas
python tests/test_parity.py                                                     # verifica que se reproduce el Excel actual
```

---

## 8. Cómo mantenerlo actualizado

Casi todo el mantenimiento se hace editando tablas en Excel, sin programar.

| Si pasa esto… | …edita esto | Cómo |
|---|---|---|
| Se abre una carrera o carrera conjunta nueva | Hoja `clave-carrera` del Excel **y** `catalogos/carrera_componentes.csv` | En el CSV agrega una fila con la clave (p. ej. `CDS DER`) y sus componentes separados por `\|` (p. ej. `CDS\|DER`) |
| Aparece una universidad nueva | Hoja `unis-info` del Excel | Agrega la universidad con `PIM` (Sí/No) y `Balance` (Sí/No) |
| Una universidad aparece con dos nombres distintos | `catalogos/alias_universidades.csv` | Pon el nombre como aparece en los datos, el nombre correcto y `aplicar = sí` |
| Un país nuevo | Hoja `país-región` del Excel | Agrega el país con su región |

Los CSV se abren con Excel. Guárdalos como **CSV UTF-8** para no perder los acentos.

---

## 9. Qué hace por dentro (para quien programa)

```
Excel ─► load_data ─► clean_data ─► validate_data ─► business_rules ─► reports ─► export_excel
                                 └─► excel_parity  (réplica de las fórmulas actuales, para comparar)
```

| Archivo | Qué hace |
|---|---|
| `run_pipeline.py` | Orquesta todo; es lo que ejecutan los scripts `correr_*` |
| `config.py` | **Todas las reglas de negocio con nombre**: hojas, columnas esperadas, tipos de movilidad, peso del verano (0.5), peso del trimestre, etc. |
| `load_data.py` | Lee el Excel en modo de solo lectura. **Es el único archivo que hay que reescribir para conectar Dataverse** |
| `clean_data.py` | Normaliza textos, claves únicas y periodos; recalcula carrera, PIM y región desde los catálogos |
| `validate_data.py` | Las 23 reglas de calidad: severidad, reportes afectados, tipo de corrección |
| `business_rules.py` | Categoría de movilidad, balance, plazo, duración y tabla puente de carreras conjuntas |
| `excel_parity.py` | Emula `COUNTIFS` (comodines, sin distinguir mayúsculas) para reproducir Balance, Reporte Anual y PIM |
| `reports.py` | Los tres reportes |
| `export_excel.py` | Escribe los libros de salida |
| `inspect_workbook.py` | Mapa de hojas, fórmulas y tablas dinámicas del Excel |
| `tests/test_parity.py` | Prueba de regresión: falla si se deja de reproducir el Excel |

**Reglas de trabajo**
- Cambia reglas solo en `config.py`, y corre `tests/test_parity.py` antes de publicar cifras.
- La limpieza **nunca borra registros**: los marca. Excluir un registro es una decisión del área.
- En los comentarios del código: `[COMPROBADO]` = observado en el Excel, `[PROPUESTA]` = decisión de diseño, `[AMBIGÜEDAD]` = pendiente de que el área decida.
- El diagnóstico técnico completo está en [`docs/diagnostico_tecnico.md`](docs/diagnostico_tecnico.md).

**Requisitos:** Python ≥ 3.11, `pandas`, `openpyxl` y `numpy` (ver `requirements.txt`).

---

## 10. Datos personales

El Excel y los resultados contienen **nombres y claves únicas de alumnos**:
- Las carpetas `datos/` y `salidas/` están excluidas del repositorio (`.gitignore`). **Nunca subas Excel ni CSV con datos.**
- Comparte los resultados solo dentro del área y por medios institucionales.
- Este repositorio debe ser **privado**.
