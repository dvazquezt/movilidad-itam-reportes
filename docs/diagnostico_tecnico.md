# Diagnóstico técnico — Reportes de Movilidad Internacional

Oct 14, 2026

Python ya reproduce exactamente los números actuales del Balance, el Reporte Anual y el PIM: coinciden las 332 filas del Balance (13 cifras cada una), las 14 celdas revisadas del Reporte Anual y las 56 del PIM. Esa misma réplica muestra que el Balance deja fuera a socios con decenas de registros y que las carreras conjuntas se cuentan sólo en su primera carrera. Aquí, **\[Comprobado\]** marca lo que se verificó en los archivos (fuente: «datos\_macro\_actual al 6-10-2026.xlsx» y «Wish list de reportes.docx»), **\[Propuesta\]** marca decisiones de diseño y **\[Ambigüedad\]** marca decisiones que le tocan al área.

## 1. Diagnóstico técnico

El archivo es un libro de cálculo de 26 hojas sin código VBA: el nombre «macro» es costumbre, no hay macros. Todo sale de una sola tabla, **In-Out (Todos)**: 7,760 registros que van del ciclo 2007-2008 al 2025-2026 (3,936 outgoing y 3,824 incoming). A su alrededor hay 3 catálogos, 5 hojas de reporte con fórmulas COUNTIFS, 22 tablas dinámicas repartidas en 15 hojas y varias hojas de trabajo.

**Hallazgos que hoy cambian cifras \[Comprobado\]**

1. **El Balance deja fuera a socios activos y cuenta doble a otros.** Cada universidad se busca por prefijo (`nombre&"*"`) dentro de una lista escrita a mano. En la ventana 2020-21 a 2024-25 quedan fuera 69 registros de tipo Normal: Universität Zu Köln (26), Prague University of Business and Economics (32), Sciences Po Paris/Poitiers (7), Ca’ Foscari (2), UPC (1) y Católica Porto (1). Köln y Praga sí están en la lista de Balance, pero con otra grafía («…Köln, WiSo Facultat», «Prague University of Economics and Business»). Rennes, Bocconi y Macalester aparecen dos veces, y «Stockholm University» también cuenta a sus facultades, que además tienen fila propia.
2. **Las tablas dinámicas tienen rangos fijos y desactualizados.** Sus rangos fuente terminan entre la fila 5,826 y la 7,203, pero los datos llegan a la fila 7,761 y no están ordenados por fecha. Si hoy se actualizan, dejan fuera 558 filas; entre ellas, 179 de 2023-24, 156 de 2024-25 y 74 de 2025-26. Sus cachés se actualizaron por última vez entre el 19-sep-2023 y el 20-nov-2025.
3. **Las carreras conjuntas sólo cuentan en su primera carrera.** «Carrera Principal» toma el primer código (ECO DFI → ECO), así que 949 registros pierden la segunda carrera. Las que más pierden son DFI (228), CPOL (158), RRI (131) y MAT (121). Las hojas SEP y SEP-TOT lo hacen a propósito («se toma la carrera de la izquierda, para no duplicar»); el reporte por carrera lo hereda sin querer.
4. **«Alumnos» casi siempre quiere decir «registros».** Las tablas dinámicas cuentan filas («Cuenta de Clave Única»), no alumnos distintos. Además, 145 claves únicas están guardadas como número y perdieron los ceros a la izquierda. Por eso un conteo de distintos da 6,509 alumnos en vez de 6,474.
5. **El histórico publicado está tecleado a mano.** En el Reporte Anual sólo el último ciclo es fórmula. 2019-20 tiene un valor estimado con asterisco, y un comentario explica que no se recalcula porque después aparecen cancelaciones y alumnos sin convenio.
6. **Hay fórmulas frágiles o rotas.** Buscador tiene 14 fórmulas con #REF!, porque apuntan a una hoja «Incoming» que ya no existe. «Edad» depende de TODAY(). Los trimestres Enero-Marzo y Abril-Junio usan agosto como mes de inicio. La ventana del Balance compara ciclos como texto, y C1 sigue en 2024-2025 aunque ya hay datos de 2025-26.

La buena noticia: la lógica actual es reproducible. La réplica en Python da exactamente los números del Excel, así que cada corrección se puede medir y explicar universidad por universidad.

## 2. Mapa de hojas

De las 26 hojas, sólo 4 son fuente o catálogo: todo lo demás se puede regenerar desde In-Out (Todos) y los catálogos. Las 22 tablas dinámicas dependen de 10 cachés distintas, cada una con su propio rango fijo \[Comprobado\]. La lista completa, con el rango fuente de cada dinámica, está en `mapa_hojas.xlsx`.

| Hoja | Tipo | Qué hace | Depende de | Estado |
| --- | --- | --- | --- | --- |
| In-Out (Todos) | **Datos (fuente)** | Un registro por alumno × periodo × universidad; 7,760 filas, 23 columnas | clave-carrera, unis-info, país-región (BUSCARV) | 53,429 fórmulas; 5,217 #N/A pegados como valor en Fecha de Nacimiento |
| clave-carrera | Catálogo | 81 claves → nombre, campo Patlani, carrera principal; incluye cursos (BCM, BPGM, LSGL, MBP) y «NA» | — | «MBA ADM » lleva un espacio al final; la carrera principal es el primer código |
| unis-info | Catálogo | 365 universidades → PIM (Sí/No) y Balance (Sí/No) | — | 10 sin valor en Balance; 2 nombres con espacio al final; 52 que nunca aparecen en los datos |
| país-región | Catálogo | 215 países → región PIM | — | Completo para los países actuales |
| Balance | Reporte (fórmulas) | Enviados/recibidos por universidad en 5 ciclos; 12 columnas COUNTIFS + balance | In-Out, unis-info | Lista de 332 universidades escrita a mano; 5 #N/A; 3 duplicadas; omite socios (ver §1); columna P «Balance» tecleada, sin fórmula |
| Reporte Anual | Reporte (mixto) | Salientes por tipo; Lic/Maestría in/out; verano en el extranjero; Buddy Program | In-Out | Sólo el último ciclo es fórmula; histórico tecleado; Buddy dice «Total 42» pero 65 + 33 = 98 |
| PIM | Reporte (fórmulas + 4 dinámicas) | PIM por región y nivel, in/out | In-Out | Dice «Long term» pero la fórmula no filtra plazo; el % está tecleado (=53\*100/156) |
| Patlani | Reporte (fórmulas) | Formato SEP-ANUIES: nivel × sexo, campo de formación | In-Out | Ciclo 2022-2023; secciones «va en 0s» |
| Patlani-Países, SEP, SEP-TOT, Fimpes, PIM-Unis, PIM Datos %, País-Nivel, País-Unis, DG, Outgoing por Carrera, Registrados | Reportes (dinámicas) | Cortes por país, carrera principal, sexo, edad, PIM, doble grado, periodo | In-Out (rango fijo) | Cachés de 2023 a 2025; cuentan registros, no alumnos |
| Doble Grado (oculta), DG 20-22 | Datos manuales | Listas de doble grado con estatus y seguimiento | — | CU numérica, universidad en minúsculas («essex»); no conectadas a In-Out |
| Buscador | Herramienta | Ficha de un alumno por nombre | In-Out, hoja «Incoming» borrada | 14 fórmulas #REF! |
| CalculosRandom | Borrador | Conteos sueltos y TRANSLATE de países | — | Sin uso en reportes |
| Hoja1, Hoja2 | Borradores (dinámicas) | Incoming por nivel × semestre × año; outgoing por país × año | In-Out | Sin nombre ni propósito documentado |
| Carreras Alumnos Inter | Reporte (dinámica) | Conteo por carrera | In-Out!E7:E7203 | Rango mal definido: arranca en la fila 7 y usa un dato («CIENCIA POLÍTICA») como encabezado; mezcla incoming y outgoing |

Respecto a lo que pediste revisar: ninguna hoja es un respaldo histórico en sentido estricto. Doble Grado, DG 20-22, Hoja1, Hoja2 y CalculosRandom son restos de trabajo manual. Las hojas «Incoming» y «Outgoing» que antes alimentaban In-Out ya no existen: lo delatan el #REF! de Buscador y un comentario en clave-carrera.

## 3. Tabla principal: In-Out (Todos)

De las 23 columnas, 15 vienen del CRM, 7 se calculan con fórmulas desde los catálogos y una es manual (Comentario). Las calculadas tienen valores pegados a mano en algunas filas: Semestre Principal en 346, Carrera y Campo Patlani en 3, PIM y Región en 2. Edad no tiene fórmula en 532 filas \[Comprobado\]. Una fila equivale a una estancia en un periodo; no hay identificador propio de registro.

| Col | Campo | Origen | Qué representa | Problemas encontrados | Lo usan |
| --- | --- | --- | --- | --- | --- |
| A | Flujo | CRM | Incoming / Outgoing | Limpio | Todos |
| B | Tipo | CRM | Modalidad: Normal 5,235; Verano en el extranjero 1,342; Especial 791; Sin convenio 206 + «Sin Convenio» 9; Visiting 98; Doble Grado 45; LSGL 29; Verano LSGL 5 | «Sin convenio» y «Sin Convenio» conviven; Especial mezcla Fulbright, ESEN y programas cortos | Balance, Anual, PIM |
| C | Alumno | CRM | Nombre | 10 claves únicas con 2 grafías del nombre (con y sin acentos, nombre completo o abreviado) | Listados |
| D | clave carrera | CRM | **Out:** carrera ITAM del alumno. **In:** programa o área donde se inscribe en el ITAM, o curso de verano (BCM, MBP, BPGM, LSGL) | 2 vacías; 7 «NA»; 789 out y 158 in son conjuntas | Por carrera, Patlani, SEP |
| E | Carrera | Fórmula (clave-carrera) | Nombre de la clave | 1 tecleada («ACTUARIA» sin acento) | Outgoing por carrera |
| F | Carrera Principal | Fórmula | Primer código de la clave | Pierde la 2.ª carrera (949 registros) | SEP, SEP-TOT |
| G | Clave Única | CRM | ID del alumno en Banner (9 dígitos) | 145 guardadas como número (sin ceros); 1 vale «0» | Alumnos únicos |
| H | Periodo | CRM | Texto del periodo en el CRM | 22 escrituras para 7 términos («FALL 2022-Undergraduate», «FALL 2025-Undergraduate (August to December)», «Spring…») | Registrados |
| I | Semestre | CRM | Otoño/Primavera/Verano o trimestre de maestría | Cambia de significado por nivel: trimestre en maestría | Balance («<>Verano»), Anual |
| J | Semestre Principal | Fórmula | Trimestre → semestre (Julio-Sep → Verano) | 346 tecleados (coinciden con la fórmula) | PIM, PIM-Unis |
| K | año | CRM | Año del periodo | 62 no coinciden con el año del Periodo (SUMMER 2017 con año 2018: 29 LSGL, 17 Visiting, 12 Especial) | Hoja2, DG |
| L | ciclo escolar | CRM / captura | Ciclo otoño→verano | 2 contradicen año+semestre (Otoño 2021 en 2019-2020; Primavera 2022 en 2022-2023) | Todos los reportes |
| M | Nivel | CRM | Licenciatura / Maestría | 7 contradicen el Periodo; 55 contradicen el programa (MBA con nivel Licenciatura, etc.) | PIM, Patlani, Anual |
| N | País | CRM | **Out:** país de destino. **In:** país de la universidad de origen | 15 universidades con 2+ países (IE en 5; Wirtschaftsuniversität Wien con Eslovaquia); 2 con MEXICO (UNAM) | País, Patlani-Países |
| O | PIM | Fórmula (unis-info) | Universidad del consorcio | 2 tecleados; Iscte no está en unis-info | PIM |
| P | Región (PIM) | Fórmula (país-región) | Región | 2 tecleados | PIM |
| Q | Campo Patlani | Fórmula | Campo SEP | Depende sólo de la primera carrera | Patlani |
| R | Universidad | CRM + edición | **Out:** destino. **In:** origen. A veces un programa (Fulbright 571, ESEN 198, DAAD 18) | 14 universidades con varias grafías; «(Sin Convenio)» dentro del nombre en 123 registros, escrito de 3 formas | Balance, PIM, top |
| S | Sexo | CRM | Male / Female | 2,247 en «0» (nulo del CRM) + 4 vacíos: **todos incoming de 2007-08 a 2018-19** | Patlani, SEP, Fimpes |
| T | Edad | Fórmula | Edad al inicio | 0 cuando falta la fecha (5,221); usa TODAY(); mes de trimestre erróneo; 532 sin fórmula | Fimpes |
| U | Fecha de Nacimiento | CRM | — | 5,217 #N/A pegados como valor (todos los registros hasta 2018-19 y algunos hasta 2022-23) y 420 vacías | Edad |
| V | Nacionalidad | CRM | **Out:** casi siempre MEXICO. **In:** nacionalidad del alumno | 806 en «0» (todos incoming, sobre todo 2007-2012), 223 vacías (outgoing 2023-2025) | Patlani-Países |
| W | Comentario | Manual | Aclaraciones frente al CRM | Mezcla datos («CANCELÓ» 8, «Tomó 2 cursos» 165) con notas libres | Conciliación |

**Duplicados \[Comprobado\].** Hay 16 registros duplicados exactos (misma CU, flujo, periodo y universidad). Entre ellos: 3 alumnos de King’s College en verano 2023 capturados como «Verano LSGL» y también como «Verano en el extranjero», 3 de LSE en verano 2023 con dos grafías del nombre, y un registro marcado CANCELÓ junto a su duplicado sin marca. Aparte, 42 registros son el mismo alumno en el mismo periodo en dos universidades: casi todos son cursos de verano dobles, que pueden ser legítimos («Se fue a dos pero cuenta como una sola alumna»).

**Alumnos y registros \[Comprobado\].** Hay 6,474 alumnos distintos. Incoming: 2,538 con 1 registro, 399 con 2, 158 con 3 y uno con 10 (diez trimestres en Torcuato di Tella). Outgoing: 466 de 3,379 alumnos tienen más de un periodo, la mayoría semestre + verano. En maestría incoming hay 1,216 registros para 751 alumnos.

**Columnas por reporte del MVP.** El Balance usa Flujo, Tipo, Semestre, ciclo, Universidad y unis-info.Balance. Outgoing por carrera usa Flujo, clave carrera, ciclo, Nivel, Clave Única y País. PIM usa Flujo, ciclo, Nivel, Semestre, Universidad→PIM y País→Región.

## 4. Problemas de datos

Hay 22 tipos de problema, detectados por 23 reglas automáticas. El detalle por registro (con la fila de Excel) está en `diagnostico_calidad_*.xlsx`. Los que más pesan en las cifras son los de universidad y duplicados (afectan el Balance) y los de carrera (afectan cualquier reporte por programa). «Auto» significa que Python lo corrige sin intervención; «semi», que propone la corrección y una persona la confirma; «humana», que el dato debe venir de otra fuente.

| # | Problema | Registros | Impacto | Reportes afectados | Corrección en Python | Tipo |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Universidades omitidas o duplicadas en la lista del Balance | 69 fuera + 4 universidades duplicadas o solapadas | Balance de Köln y Praga en 0 cuando deberían ser −18 y −16 | Balance | Construir la lista desde los datos + unis-info, no a mano | Auto |
| 2 | Misma universidad con varias grafías (mayúsculas, espacios, marca «(Sin Convenio)») | 672 | Top de universidades y «nuevos destinos» inflados | Balance, top, nuevos destinos | Llave normalizada `uni_key` | Auto |
| 3 | Grafías distintas que sí son la misma universidad («Californa-San Diego» / «CALIFORNIA SAN DIEGO», «UNIVARSITY») | 4 propuestas | Lo mismo que el anterior | Igual | Tabla `alias_universidades.csv` | Semi |
| 4 | «(Sin Convenio)» en el nombre contradice el Tipo | 117 | No se sabe si cuenta para el balance | Balance, sin convenio | Convenio como campo, no como texto | Humana |
| 5 | Registros duplicados exactos | 16 | Conteos inflados | Todos | Marcar y excluir con lista revisada | Semi |
| 6 | Comentario «CANCELÓ» y el registro sigue contando | 8 | Enviados/recibidos inflados | Todos | Bandera `flag_cancelado`; excluir sólo cuando el área lo confirme | Humana |
| 7 | Año del Periodo ≠ columna año | 62 | El registro cae en otro ciclo | Por ciclo | Reportar; decidir qué campo manda | Humana |
| 8 | Ciclo ≠ año + semestre | 2 | Ídem | Por ciclo | Recalcular ciclo | Semi |
| 9 | Clave Única numérica (sin ceros) | 145 | 34 alumnos contados dos veces | Alumnos únicos, repetidores | Texto de 9 dígitos | Auto |
| 10 | Misma CU con nombres distintos | 22 | Listados inconsistentes | Listados | Nombre de Banner | Semi |
| 11 | Carrera conjunta reducida a la primera | 949 | Subconteo de la 2.ª carrera | Por carrera, Patlani campo | Tabla puente de componentes | Auto (con catálogo revisado) |
| 12 | Claves de carrera ambiguas (CON CEF, FIN CON, ADM CON CEF, ADM CON DFI) | 4 claves | No se sabe si son conjuntas | Por carrera | Marcadas en `carrera_componentes.csv` | Humana (Servicios Escolares) |
| 13 | Clave de carrera vacía o «NA» | 9 | Sin carrera | Por carrera | Listar para captura | Humana |
| 14 | Nivel contradice Periodo o programa | 7 + 55 | Mal clasificado Lic/Maestría | PIM, Patlani, Anual | Reportar | Humana |
| 15 | Sexo = 0 o vacío | 2,251 | Sin sexo para todo incoming hasta 2018-19 | Patlani, SEP, equidad | Nulo explícito; recuperar de Banner | Humana |
| 16 | Nacionalidad 0 o vacía | 1,029 | Patlani-Países incompleto | Patlani, 911 | Nulo explícito | Humana |
| 17 | Fecha de nacimiento #N/A o vacía | 5,637 | Edad = 0 en lugar de vacía | Fimpes | Nulo + edad recalculada | Humana |
| 18 | Universidad con país atípico | 22 | Región PIM equivocada | País, región, PIM | País desde catálogo de universidades | Semi |
| 19 | Programas en la columna Universidad (Fulbright, ESEN, DAAD) | 806 | Fulbright sale como «universidad» n.º 1 | Top, PIM | Separar programa de institución | Humana (regla) |
| 20 | Universidad sin registro en unis-info (Iscte) | 1 | PIM indefinido | PIM | Alta en catálogo | Humana |
| 21 | Valores pegados donde debería haber fórmula | 346 + 532 + 13 | Cambios silenciosos al modificar catálogos | PIM, Fimpes | Python recalcula siempre | Auto |
| 22 | Fórmulas rotas: Buscador #REF!, rangos de dinámicas, Edad con TODAY() | 14 fórmulas, 22 dinámicas | Resultados variables o incompletos | Buscador, todas las dinámicas | Se reemplazan con código | Auto |

También hay comentarios que contradicen el CRM: «Estaba cancelado pero sí vino – lo agregué a mano», «no tiene solicitud en el CRM», «CRM dice ECO y COM, eso no existe» y «No distingue cds lic vs mtia», entre otros (18 registros). Son la memoria de correcciones manuales: el pipeline los conserva y los lista, pero no los interpreta.

La wishlist dice 956 registros con Carrera ≠ Carrera Principal; hoy son 949. La diferencia es menor y probablemente viene de cambios posteriores al conteo. Las demás cifras de la wishlist coinciden: Sexo 0 = 2,247; 123 nombres con «(Sin Convenio)»; nacionalidad vacía = 223; carrera vacía = 9 (2 vacías + 7 «NA»).

## 5. Reglas de negocio

Encontré 26 reglas: 7 ya las conocían, 12 están sólo dentro de fórmulas o comentarios, y 7 vienen de la wishlist pero hoy no se aplican igual en el Excel. En Python todas viven con nombre en `config.py` y `business_rules.py`; ningún criterio queda escondido en un comodín.

**Las que ya conocían, confirmadas en fórmulas \[Comprobado\]**

| # | Regla | Dónde está |
| --- | --- | --- |
| R1 | Balance = salientes semestrales − entrantes semestrales + 0.5 × (verano saliente − verano entrante) | Balance!O: `=D-C+0.5*(F-E)` |
| R2 | Sólo cuenta para el balance el Tipo «Normal»; Especial, Visiting, Sin convenio, LSGL y Doble Grado se muestran aparte | Criterio `B:B,"Normal"` en Balance C–F |
| R3 | Ciclo escolar = otoño a verano; Oct-Dic abre ciclo; Ene-Mar, Abr-Jun y Jul-Sep cierran | Columna L (7,758 de 7,760 lo cumplen) |
| R4 | Contar alumnos ≠ contar registros | Wishlist; el Excel cuenta registros |
| R5 | Carrera conjunta: debe poder contarse por cualquiera de las dos | Wishlist; el Excel sólo usa la primera |
| R6 | PIM se marca por universidad en unis-info | Columna O = BUSCARV(unis-info, 2) |
| R7 | Meta PIM: ≥ 20 % de la movilidad con escuelas del consorcio | Wishlist |

**Implícitas en fórmulas y comentarios \[Comprobado\]**

| # | Regla | Dónde está |
| --- | --- | --- |
| R8 | «Con convenio» = Normal + Especial (criterio `"*al"`) | Reporte Anual C6/C8 y su comentario |
| R9 | «Programa de verano en el extranjero» = todo Tipo que empiece con «Verano», incluido Verano LSGL | Reporte Anual C11 y comentario |
| R10 | Universidad con y sin «(Sin Convenio)» se agrupan en el Balance porque el Tipo ya filtra | Comentario en Balance!B4; criterio `B6&"*"` |
| R11 | Una universidad entra al Balance si unis-info.Balance = «Sí» y tuvo al menos un movimiento; doble grado y «sin convenio» tienen «No» | Balance!A y su comentario |
| R12 | Ventana del Balance = últimos N ciclos hasta C1 (hoy 5: 2020-21 a 2024-25) | Balance!A1, C1, C2 |
| R13 | Cada trimestre de maestría cuenta como un semestre completo en el balance (sólo se excluye «Verano») | Criterio `I:I,"<>Verano"` |
| R14 | Jul-Sep de maestría se agrupa como «Verano»; Ene-Mar y Abr-Jun como «Primavera»; Oct-Dic como «Otoño» | Columna J |
| R15 | Carrera principal = la de la izquierda, «para no duplicar alumnos con 2 programas» | Comentario en SEP y SEP-TOT |
| R16 | Edad = años cumplidos al inicio del periodo (mes 1, 6 u 8) | Comentario en In-Out!T1 |
| R17 | Sexo «0» = nulo del CRM | Comentario en In-Out!S1 |
| R18 | Las cifras de años ya reportados no se recalculan: se congelan | Reporte Anual!B14 y B28 |
| R19 | Total del Reporte Anual = intercambio + doble grado; la tabla histórica incluye doble grado dentro de «por convenio» | Reporte Anual G6, G7, C27 |

**De la wishlist, aún no aplicadas igual en el Excel \[Comprobado: la diferencia\]**

| # | Regla | Situación actual |
| --- | --- | --- |
| R20 | Los que pagan (Visiting, Sin convenio) se muestran en el Balance pero no cuentan | Se muestran, pero sin filtro de Flujo: las columnas Visiting y Sin convenio mezclan in y out |
| R21 | Duración: corto plazo/verano, un semestre, un año | No existe; Python la deriva (110 registros outgoing son «un año») |
| R22 | Plazo PIM largo/corto | La hoja PIM dice «Long term» y no filtra plazo; la tabla Largo/Corto está tecleada |
| R23 | Semáforo: balanceado / a favor / en contra | No existe |
| R24 | Alumnos con más de una movilidad | No existe; hay 466 outgoing con 2 o más periodos |
| R25 | Estatus del proceso (interesado → concluido) | Sólo «CANCELÓ» dentro de Comentario |
| R26 | Tipos de movilidad: Normal, Sin convenio, Verano en el extranjero, Especial, Visiting, Doble Grado, LSGL, Verano LSGL | Coinciden con los 8 valores («Sin Convenio» es variante) |

**\[Ambigüedad\] Decisiones que hacen falta**

1. R13: ¿un trimestre de maestría debe pesar 1 en el balance, o 0.5? Python lo deja configurable (`BALANCE_PESO_TRIMESTRE`) y por ahora replica el Excel.
2. R23: ¿qué significa «a favor del ITAM», enviar más o recibir más? ¿Con qué tolerancia algo es «balanceado»?
3. R2 y R11: ¿los Especial (Fulbright, ESEN, NUS) cuentan en el balance de algún convenio?
4. ¿Los registros «CANCELÓ» se excluyen? Hoy cuentan.
5. Total del Reporte Anual: ¿con o sin doble grado? Las dos tablas de la misma hoja lo hacen distinto.
6. PIM: ¿el 20 % se mide en registros o en alumnos, por ciclo, y sobre qué total? La celda del Excel usa 53/156 tecleado.

## 6. Modelo de datos propuesto \[Propuesta\]

Propongo una tabla de hechos, `movilidad` (una fila por alumno × periodo × universidad, igual que hoy), y siete catálogos pequeños. La única tabla puente es la que resuelve las carreras conjuntas. No hace falta una base de datos: en la etapa Excel son DataFrames y CSV, y en Power BI son las mismas tablas como modelo estrella.

| Tabla | Llave primaria | Campos | Relación | Origen hoy |
| --- | --- | --- | --- | --- |
| **movilidad** | `id_movilidad` (hoy: archivo + fila; con Dataverse: GUID del registro) | flujo (IN/OUT), clave\_unica, periodo\_id, uni\_id, clave\_carrera, tipo, nivel, estatus, flag\_cancelado, comentario, fuente\_fila | Hechos; N:1 con todas las demás | In-Out (Todos) |
| **alumnos** | `clave_unica` (texto de 9 dígitos) | nombre, sexo, nacionalidad, fecha\_nacimiento | 1:N con movilidad | In-Out (deduplicado) |
| **periodos** | `periodo_id` (p. ej. 2024-OTO, 2025-ENE-MAR) | año, término, término\_principal, ciclo, esquema (semestral/trimestral), es\_verano, peso\_balance | 1:N con movilidad | Derivado de Periodo, Semestre y año |
| **universidades** | `uni_id` | nombre\_canónico, país, es\_pim, es\_programa (Fulbright, ESEN, DAAD), cuenta\_balance | 1:N con movilidad; 1:N con alias | unis-info + datos |
| **universidad\_alias** | `alias` | uni\_id, motivo, validado\_por | N:1 con universidades | Nuevo (`alias_universidades.csv`) |
| **paises** | `pais` | región\_PIM, nombre\_es | 1:N con universidades | país-región |
| **programas** | `clave_programa` (ADM, ECO, MBA ADM…) | nombre, nivel, división\*, campo\_Patlani | 1:N con plan\_componente | clave-carrera |
| **planes** (clave de carrera) | `clave_carrera` (ECO DFI) | nombre, es\_conjunta | 1:N con movilidad; 1:N con plan\_componente | clave-carrera |
| **plan\_componente** (puente) | (`clave_carrera`, `clave_programa`) | orden, es\_principal | N:1 con planes y programas | Nuevo (`carrera_componentes.csv`) |
| **tipos\_movilidad** | `tipo` | con\_convenio, paga, cuenta\_balance, es\_programa\_verano | 1:N con movilidad | Reglas R2, R8 y R9 |
| **convenios** | `convenio_id` | uni\_id, tipo, plazas\*, vigencia\*, red\* | N:1 con universidades | Sólo la bandera Balance de unis-info |
| **cifras\_publicadas** | (reporte, ciclo, métrica) | valor, fecha\_publicación, nota | — | Histórico tecleado del Reporte Anual |

\* No existe hoy en el Excel (ver §10).

**Incoming y outgoing en la misma tabla.** Se distinguen con `flujo`. Tres columnas cambian de sentido según el flujo y se documentan así: `uni_id` es destino (OUT) u origen (IN); `pais` sigue a la universidad; `clave_carrera` es la carrera del alumno (OUT) o el programa receptor en el ITAM (IN). No conviene separar dos tablas: el Balance necesita ambos flujos juntos.

**Carrera conjunta.** El registro guarda la clave tal cual (`ECO DFI`). La tabla puente la abre en ECO y DFI. Los reportes por programa usan el puente, y cada conjunta aparece una vez en cada carrera. Los totales institucionales usan `movilidad` directo, donde aparece una sola vez. «Carrera principal» queda como el componente con `es_principal`, que reproduce lo que hace hoy SEP.

**Alumno con varias movilidades.** Hay una fila por estancia y periodo, ligada al mismo `clave_unica`. Una *estancia* es la combinación (alumno, universidad, ciclo): sirve para derivar la duración (otoño + primavera en la misma universidad = un año) y para contar repetidores sin confundirlos con trimestres.

**Maestría por trimestre.** Cada trimestre es un registro con su `periodo_id`. Para cualquier corte se publican tres conteos: registros (trimestres), alumnos distintos y estancias. Para incoming de maestría dan 1,216 registros y 751 alumnos. El peso de un trimestre en el balance es un parámetro de `periodos.peso_balance` (ver Ambigüedad 1 en §5).

## 7. Arquitectura del código \[Propuesta, ya implementada como borrador\]

Son diez archivos de Python (más una prueba) que sólo usan pandas, openpyxl y numpy, y una carpeta de catálogos CSV editables. Sólo `load_data.py` sabe que la fuente es un Excel: cambiar a Dataverse es reescribir ese archivo. El pipeline completo tarda alrededor de 1 minuto y nunca escribe sobre el archivo original. Abre el libro en modo sólo lectura y guarda su huella SHA-256 en cada salida.

El flujo, en orden:

1. **Cargar** (`load_data.py`): lee In-Out (Todos) con los valores y, aparte, las fórmulas. También lee los tres catálogos y las hojas de reporte para compararlas. Valida que existan las 23 columnas esperadas y conserva la fila de Excel de cada registro.
2. **Limpiar y normalizar** (`clean_data.py`): normaliza textos y Clave Única, interpreta el Periodo y recalcula desde los catálogos carrera, PIM y región (no confía en los valores cacheados). Crea la llave de universidad, pone nulos explícitos donde había «0» y #N/A, y recalcula la edad. Cada cambio queda anotado en `_correcciones`.
3. **Validar** (`validate_data.py`): 23 reglas con severidad, reportes afectados y tipo de corrección. No borra nada; produce el inventario.
4. **Aplicar reglas de negocio** (`business_rules.py`): categoría de movilidad, con o sin convenio, si paga, si cuenta en el balance y con qué peso, plazo, duración, puente de carreras y dimensión alumno.
5. **Replicar el Excel** (`excel_parity.py`): emula COUNTIFS (comodines, sin distinguir mayúsculas) y reproduce el Balance, el Reporte Anual y el PIM tal como los calcula hoy el Excel.
6. **Reportar** (`reports.py`): los tres reportes del MVP en versión corregida, más la explicación de por qué difieren del Excel.
7. **Exportar** (`export_excel.py`): libros nuevos con fecha en el nombre, una hoja LEEME con la fuente, su huella y los parámetros, encabezados fijos y filtros.

El orquestador es `run_pipeline.py`; `inspect_workbook.py` genera el mapa de hojas y dinámicas, y `tests/test_parity.py` falla si la réplica deja de coincidir.

| Archivo | Responsabilidad | Cambia cuando… |
| --- | --- | --- |
| `config.py` | Nombres de hojas, columnas esperadas, mapeos y parámetros de balance | Cambia una regla o una categoría |
| `catalogos/carrera_componentes.csv` | Qué claves son conjuntas y sus componentes | Se abre una carrera o conjunta nueva |
| `catalogos/alias_universidades.csv` | Grafías que son la misma universidad (sólo aplica `aplicar = sí`) | Aparece una grafía nueva |
| `load_data.py` | Única puerta de entrada | Llega Dataverse |
| `clean_data.py` · `validate_data.py` · `business_rules.py` · `reports.py` | Lógica | Casi nunca |

Tres decisiones deliberadas. No uso base de datos ni orquestadores, porque con 7,760 filas pandas basta y alguien del área puede mantenerlo. Los catálogos son CSV, para que una persona sin Python pueda corregirlos. Y la limpieza nunca borra registros: marca, y la exclusión se decide en el reporte, para que siempre se pueda explicar una diferencia.

## 8. MVP recomendado: tres reportes

Recomiendo **Balance**, **Outgoing por carrera** y **PIM**. Son los tres con más riesgo de dar una cifra equivocada hoy, y los tres comparten toda la capa de limpieza. El Reporte Anual queda como prueba de paridad, no como reporte nuevo. Los tres ya corren en el borrador; las cifras de abajo son del ciclo 2024-2025 \[Comprobado sobre los datos actuales\].

### 8.1 Balance por convenio / universidad

- **Filtros:** ciclo final y número de ciclos (por defecto el del Excel: 2024-2025 y 5 ciclos); universidad; país; sólo «incluir = Sí».
- **Columnas que usa:** flujo, tipo, semestre, ciclo, universidad (llave), unis-info.Balance, clave única.
- **Lógica:** se agrupa por llave de universidad, no por una lista tecleada. Las columnas se separan en: convenio semestral, trimestre de maestría (columna propia), verano con convenio, especial, visiting, sin convenio, verano en el extranjero y doble grado. Balance = Σ (+1 si sale, −1 si entra) × peso (verano 0.5, trimestre configurable), sólo para Tipo Normal. Se agregan enviados y recibidos como alumnos únicos, el semáforo y una hoja con el balance por ciclo y su acumulado.
- **Resultado:** 151 universidades con «incluir» (el Excel marca 147). Cambian Köln (de 0 a −18), Praga (de 0 a −16), Sciences Po Paris/Poitiers (+5), UPC (−1) y Stockholm University (−7 → −10 al separar sus facultades). Rennes deja de aparecer dos veces.
- **Riesgos:** fusionar universidades que en realidad son convenios distintos (Autónoma de Barcelona tiene filas por facultad: DER, ECO, RI/CPL); el peso del trimestre; registros cancelados; el signo del semáforo.
- **Validación:** la réplica ya coincide en las 332 filas. La hoja `balance_por_que_difiere` clasifica cada diferencia en no capturado, contado varias veces, fila duplicada o nombre que no está en unis-info.

### 8.2 Outgoing por carrera (simple y conjunta)

- **Filtros:** ciclo o rango de ciclos, periodo, nivel, tipo.
- **Columnas que usa:** flujo = OUT, clave carrera, tabla puente de componentes, clave única, país y universidad.
- **Lógica:** se reportan dos vistas lado a lado. La primera es «carrera principal», que equivale a lo de hoy. La segunda es «por componente», donde una conjunta suma en sus dos carreras. Cada vista da registros y alumnos, y una columna muestra cuánto se pierde con la principal. El total institucional se calcula aparte, sin el puente, para no duplicar. Incluye la matriz carrera × país.
- **Resultado 2024-25:** 236 registros y 232 alumnos en total. 47 registros son de carreras conjuntas, por eso la suma por componente da 283 (y no es un total). Matemáticas Aplicadas pasa de 8 a 16, Ciencia Política de 1 a 9 y Dirección Financiera de 41 a 52.
- **Riesgos:** las 4 claves ambiguas del catálogo (§4, #12); claves con 3 programas; la carrera de incoming no es la carrera del alumno.
- **Validación:** la vista «principal» debe coincidir con la dinámica «Outgoing por Carrera» una vez actualizada con el rango completo. Su caché (de marzo de 2024) da 317 registros para 2022-23; los datos actuales dan 322. La suma de la vista «componente» debe ser igual al total + el número de registros conjuntos (283 = 236 + 47).

### 8.3 Red PIM

- **Filtros:** ciclo, flujo, nivel, plazo (corto = verano; largo = semestre o trimestre), región.
- **Columnas que usa:** flujo, nivel, semestre, tipo, universidad → PIM, país → región, clave única.
- **Lógica:** para cada combinación se calcula total y PIM en registros y en alumnos, el % y si se cumple la meta del 20 %. Se acompaña de los desgloses por región y por universidad.
- **Resultado 2024-25:** 135 de 462 registros son PIM (29.2 %); en alumnos, 135 de 415 (32.5 %). Outgoing llega al 34.3 % e incoming al 23.9 %. Incoming de maestría queda en 12.7 %.
- **Riesgos:** la bandera PIM de unis-info refleja la membresía de hoy, no la de cada año, así que un histórico puede contar escuelas que entraron después. Falta la definición exacta del 20 % (§5, Ambigüedad 6). Fulbright y ESEN entran como «no PIM».
- **Validación:** la réplica de las 56 celdas del PIM coincide. El 33.97 % tecleado (53/156) debe reproducirse con un filtro explícito; si no se puede, documentar qué se dividió.

## 9. Plan de validación

La regla es simple: **ninguna cifra cambia sin una causa con nombre.** Primero Python reproduce el Excel tal como es (eso ya está hecho); después, cada diferencia entre esa réplica y la versión corregida se atribuye a una regla del §4 o del §5. Lo que no se pueda explicar se considera un error del pipeline.

**Cuatro niveles**

1. **Integridad de la fuente.** Para cada corrida se guardan la huella SHA-256, el número de filas y las columnas. Si cambia la estructura (una columna nueva o renombrada), el pipeline se detiene.
2. **Paridad exacta con las fórmulas actuales.** `excel_parity.py` emula COUNTIFS y se compara contra los valores guardados en el libro. Resultado de hoy: Balance 332/332 filas, Reporte Anual 14/14 celdas y PIM 56/56 celdas \[Comprobado\]. Queda como prueba automática (`tests/test_parity.py`). Siguiente paso: añadir Patlani y rehacer cada tabla dinámica con el rango completo.
3. **Conciliación réplica → corregido.** Para cada métrica se arma un puente: valor Excel + registros no capturados − registros contados doble − duplicados − cancelados (si se decide excluirlos) ± cambios de ciclo = valor Python. Para el Balance ya existe (`balance_por_que_difiere`).
4. **Contra cifras ya publicadas.** Se recalculan los ciclos tecleados del Reporte Anual con la base de hoy y se registra la diferencia. No se «corrige» lo publicado: se guarda en `cifras_publicadas` con su fecha.

**Qué se compara en cada dimensión**

| Dimensión | Comparación | Identidad que debe cumplirse | Causa esperada de diferencia |
| --- | --- | --- | --- |
| Conteos por periodo | Registros por periodo × flujo vs. la dinámica Registrados recalculada con rango completo | Suma por periodo = total del ciclo | Filas fuera del rango de la dinámica; año del Periodo ≠ año (62) |
| Alumnos únicos | CU como en Excel vs CU normalizada | 6,509 − 34 (CU numéricas) − 1 (CU «0») = 6,474 | Formato de CU; mismo alumno con dos nombres |
| Carrera | Vista principal vs vista por componente | componente = total + registros conjuntos | Claves ambiguas del catálogo |
| Universidad | Lista de grafías fusionadas por llave | Cada fusión aprobada en `alias_universidades.csv` o por regla de mayúsculas | Grafías; facultades con convenio propio |
| Balance | Por universidad: réplica vs corregido | Diferencia 100 % atribuida en `balance_por_que_difiere` | Lista manual, prefijos, duplicados |
| PIM | Registros y alumnos PIM por ciclo vs PIM Datos % | Total igual a la suma por flujo | Iscte sin catálogo; filas fuera de rango (459 en caché vs 462 hoy para 2024-25) |

**Ya encontrado en el nivel 4 \[Comprobado\].** Las cifras publicadas de salientes de licenciatura no coinciden con la base actual en dos ciclos. Uso la definición de la fila dinámica del Reporte Anual: convenio + doble grado, sin convenio, verano en el extranjero y maestría.

| Ciclo | Publicado | Base actual | Diferencia | Explicación disponible |
| --- | --- | --- | --- | --- |
| 2023-24 | 269 | 269 | 0 | — |
| 2022-23 | 317 | 322 | +5 | Altas tardías (comentario B14) |
| 2021-22 | 225 | 227 | +2 | Altas tardías |
| 2019-20 | 449\* | 163 | −286 | El 449 era una estimación (asterisco); no hubo verano 2020 |
| 2018-19 | 415 | 361 | −54 | **Ninguna** (convenio −19, verano −35) |
| 2017-18 | 355 | 306 | −49 | **Ninguna** (convenio −29, verano −20) |
| 2016-17 | 274 | 275 | +1 | Altas tardías |

&#91;Ambigüedad\] No sé si en 2017-18 y 2018-19 se borraron registros de la base o si el informe de esos años usó otra definición. Hay que revisar los informes publicados de esos ciclos antes de cerrar la etapa 2.

## 10. Lo que no se puede resolver todavía

Con el Excel actual se pueden hacer casi todos los reportes de conteo de la wishlist. Los listados operativos («alumnos en el extranjero hoy»), las tasas, los convenios con plazas y las becas necesitan datos que hoy no están en el archivo. Cuando digo «probablemente en Dataverse» es una suposición: no he visto esa base.

| Dato | Dónde está | Reportes que lo necesitan |
| --- | --- | --- |
| Flujo, tipo, periodo, ciclo, nivel, país, universidad, carrera, sexo (desde 2018-19), nacionalidad | **Ya en el Excel** | Balance, Anual, PIM, por carrera, país, Patlani |
| Duración (verano / semestre / año) y repetidores | **Ya en el Excel**, derivable (110 registros «un año»; 466 alumnos con más de un periodo) | Salientes por duración, repetidores |
| Nuevos destinos y destinos que dejaron de usarse | **Ya en el Excel**, derivable una vez normalizadas las universidades | Destinos |
| Componentes de cada carrera conjunta | **Otra tabla** (catálogo a validar) | Por carrera |
| División o departamento de cada carrera | **Otra tabla** (catálogo pequeño) | Salientes por división |
| País canónico de cada universidad; universidad vs programa (Fulbright) | **Otra tabla** (catálogo de universidades) | País, región, top |
| Convenio como campo: tipo, plazas, vigencia, red | **Otra tabla** (hoy se deduce del Tipo y del texto «(Sin Convenio)»; los convenios bilaterales se cuentan «del brochure») | Balance contra plazas, convenios vigentes, «dormidos», demanda vs plazas |
| Fecha de alta/baja de cada escuela en PIM | **Otra tabla** | PIM histórico |
| Sexo de incoming 2007-08 a 2018-19; fechas de nacimiento previas a 2019 | **Banner** (o Dataverse, si el CRM las tiene) | Patlani, SEP, edad |
| Semestre que cursa el alumno al viajar (la columna Semestre es la temporada, no el avance) | **Servicios Escolares / Banner** | Salientes por semestre cursado; tasa sobre 5.º semestre en adelante |
| Matrícula por carrera, semestre y sexo; egresados por generación | **Servicios Escolares** | Todas las tasas de participación e internacionalización |
| Materias cursadas, equivalencias, créditos, calificaciones, promedio | **Servicios Escolares / otra plataforma** (la wishlist lo dice) | Desempeño académico; materias más solicitadas por incoming |
| Estatus del proceso (interesado → concluido), cancelaciones con motivo | **Dataverse**, probablemente (hoy sólo «CANCELÓ» en Comentario) | Pipeline de solicitudes |
| Fechas exactas de estancia, ciudad, seguro médico, alojamiento, contacto de emergencia, estatus migratorio, idioma declarado | **Dataverse**, probablemente | Listados nominales; idioma (español vs sólo inglés) |
| Buddy Program | **Dataverse** (pestaña «buddies», según nota del Reporte Anual) | Reporte Anual |
| Número de cursos de verano en el extranjero | **Otra tabla** («la tabla de cursos en el reporte») | Reporte Anual |
| Becas (fuente y monto), pagos y colegiaturas | **No existe hoy** (la wishlist lo confirma para becas) | Becas, SEP, finanzas, sin convenio |
| Movilidad docente, virtual/COIL, investigación, encuestas, incidentes, rankings | **No existe hoy** en estos archivos | Secciones 8, 11 y 12 de la wishlist |

El doble grado queda a medias. En In-Out hay 45 registros de tipo Doble Grado, pero «titulados» y el seguimiento están en las hojas manuales Doble Grado y DG 20-22 (con estatus como «rehabilitada» y «no está en CRM»). Se pueden cruzar por CU, pero no son una fuente confiable.

## 11. Plan de implementación

Las etapas 1 y 2 ya tienen borrador funcionando; lo que falta es sobre todo decisión del área, no código. La etapa 4 no debería empezar hasta que las cifras de la 2 estén aceptadas: un dashboard sobre datos sin conciliar sólo hace más visible el problema. Los esfuerzos son mi estimación para una persona de servicio social a medio tiempo.

| Etapa | Tareas | Entregables | Dificultad / esfuerzo | Depende de | Riesgos |
| --- | --- | --- | --- | --- | --- |
| **1. Diagnóstico y limpieza** | Revisar con el área el inventario de problemas; validar `carrera_componentes.csv` y `alias_universidades.csv`; decidir las 6 ambigüedades del §5; marcar duplicados y cancelados | Inventario aprobado; catálogos validados; tabla `movilidad_limpia` | Baja en código, media en coordinación; 2–3 semanas | Tiempo de la persona que conoce el proceso | Que las decisiones se queden sin dueño; corregir en Python lo que debería corregirse en el CRM |
| **2. Reproducir reportes actuales** | Paridad de Patlani, SEP, PIM-Unis y dinámicas con rango completo; conciliar 2017-18 y 2018-19 contra lo publicado; cargar `cifras_publicadas` | Paridad 100 % documentada; puente de diferencias por reporte | Media; 3–4 semanas | Etapa 1; informes publicados de años anteriores | Diferencias sin explicación en el histórico; presión por «que cuadre» con lo publicado |
| **3. Excel automático** | Un comando que genere el libro de reportes del ciclo (MVP + Anual + Patlani); hoja LEEME con fuente, huella y parámetros; manual de una página | `reportes_<ciclo>.xlsx` en un clic; manual de operación | Baja; 2 semanas | Etapa 2 | Que se siga editando el Excel a mano después de generarlo |
| **4. Dashboard en Power BI** | Publicar las tablas del §6 como modelo estrella; medidas separadas para registros y alumnos (DISTINCTCOUNT); la puente de carreras como relación de muchos a muchos sólo para la vista por carrera | Tablero con filtros de la wishlist (periodo, flujo, nivel, carrera, país, región, tipo, PIM) | Media; 3–4 semanas | Etapa 3; licencias y espacio de trabajo de Power BI | Que las medidas DAX vuelvan a implementar reglas por su cuenta: deben leer columnas ya calculadas por Python |
| **5. Fuente Dataverse** | Escribir `load_from_dataverse()` con el mismo esquema; mapear campos del CRM a las 23 columnas; correr en paralelo Excel vs Dataverse al menos un ciclo | Pipeline con dos fuentes intercambiables; reporte de diferencias | Media-alta; 3–6 semanas | Acceso de lectura a Dataverse (TI); diccionario de campos | Que Dataverse no tenga las correcciones manuales del Excel (comentarios, altas a mano); permisos |

Para la etapa 5 conviene pedir desde ya, aunque el acceso tarde: la lista de tablas y campos de Dataverse que corresponden a In-Out, y si el CRM guarda estatus, fechas de estancia e idioma. Eso decide cuántas filas del §10 se resuelven sin pedir nada a Servicios Escolares.

## 12. Lista priorizada de tareas

Las primeras cinco cambian cifras que hoy se reportan, así que conviene hacerlas antes del próximo informe; ninguna requiere tocar el CRM.

**Prioridad 1: corrige cifras actuales**

- [ ] Revisar con el área la hoja `balance_por_que_difiere` y aceptar o rechazar el balance corregido de Köln, Praga, Sciences Po Paris/Poitiers, Rennes y Stockholm.
- [ ] Validar `carrera_componentes.csv`, sobre todo CON DFI, CON CEF, FIN CON, ADM CON CEF y ADM CON DFI (con Servicios Escolares si hace falta).
- [ ] Decidir las 6 ambigüedades del §5: peso del trimestre, signo del semáforo, Especial en el balance, cancelados, total del anual y definición del 20 % PIM.
- [ ] Revisar los 16 duplicados exactos y los 8 «CANCELÓ».
- [ ] No volver a actualizar las dinámicas del Excel sin ampliar su rango a toda la tabla, o mejor convertir In-Out (Todos) en Tabla de Excel. Es el arreglo más barato mientras llega Python.

**Prioridad 2: hace confiable el proceso**

- [ ] Investigar la diferencia de 2017-18 y 2018-19 contra lo publicado.
- [ ] Revisar los 62 registros con año del Periodo ≠ año y decidir qué campo manda.
- [ ] Aprobar o rechazar las propuestas de `alias_universidades.csv` y decidir si las facultades de Autónoma de Barcelona y Stockholm son convenios separados.
- [ ] Dar de alta Iscte en unis-info y completar las 10 universidades sin valor de Balance.
- [ ] Cargar el histórico tecleado del Reporte Anual como `cifras_publicadas`.
- [ ] Añadir la paridad de Patlani y SEP a `tests/test_parity.py`.

**Prioridad 3: prepara lo que sigue**

- [ ] Pedir a TI el diccionario de campos de Dataverse y acceso de lectura.
- [ ] Pedir a Servicios Escolares la matrícula por carrera, semestre y sexo, y el sexo y la fecha de nacimiento faltantes de incoming 2007-2019.
- [ ] Crear el catálogo de divisiones y el de convenios (plazas, vigencia, red).
- [ ] Escribir el manual de una página para correr el pipeline cada ciclo.

**Archivos entregados con este documento**

- `movilidad_itam/`: el código (`config`, `load_data`, `clean_data`, `validate_data`, `business_rules`, `excel_parity`, `reports`, `export_excel`, `run_pipeline`, `inspect_workbook`, `tests/test_parity`, `catalogos/`, `README.md`).
- `diagnostico_calidad_*.xlsx`: inventario por regla y por registro, fórmulas pegadas, componentes de carrera y dimensión alumno.
- `paridad_excel_python_*.xlsx`: réplica del Excel y explicación de diferencias.
- `reportes_mvp_2024-2025_*.xlsx`: los tres reportes del MVP.
- `mapa_hojas.xlsx`: las 26 hojas y las 22 tablas dinámicas con su rango fuente.
- `movilidad_limpia_*.csv`: la tabla limpia, con la fila de Excel de origen.

El archivo original no se modificó: su huella SHA-256 (276a1a77…c8e7f4) es la misma antes y después, y viene anotada en cada salida.
