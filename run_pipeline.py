"""
Orquestador:  Excel original -> carga -> limpieza -> validación -> reglas ->
paridad con Excel -> reportes MVP -> archivos nuevos en ./salidas

Uso:
    python run_pipeline.py --input "datos_macro_actual al 6-10-2026.xlsx" \
                           --out salidas [--ciclo 2024-2025] [--anios 5]
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))

import config as C                                     # noqa: E402
from load_data import load_source, cell                 # noqa: E402
from clean_data import (build_catalogs, build_carrera_componentes, clean_inout,  # noqa: E402
                        load_aliases, uni_key)
from validate_data import run_validations               # noqa: E402
from business_rules import classify, carrera_bridge, alumnos_dim  # noqa: E402
import excel_parity as XP                               # noqa: E402
import reports as R                                     # noqa: E402
from export_excel import write_workbook, stamp          # noqa: E402


def formula_audit(src) -> pd.DataFrame:
    """Columnas que deberían ser fórmula pero tienen valores pegados a mano."""
    if src.inout_formulas is None:
        return pd.DataFrame()
    f = src.inout_formulas
    rows = []
    for col, desc in C.DERIVED_IN_EXCEL.items():
        is_f = f[col].astype(str).str.startswith("=")
        hard = f.loc[~is_f & f[col].notna(), "_fuente_fila"]
        empty = f.loc[f[col].isna(), "_fuente_fila"]
        rows.append({"columna": col, "fórmula_esperada": desc, "filas_con_fórmula": int(is_f.sum()),
                     "filas_valor_pegado": len(hard), "filas_vacías": len(empty),
                     "ejemplo_filas_pegadas": ", ".join(map(str, hard.head(8)))})
    return pd.DataFrame(rows)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default=None, help="ruta al Excel; si se omite, usa el único .xlsx de la carpeta datos/")
    ap.add_argument("--out", default=str(Path(__file__).resolve().parent / "salidas"))
    ap.add_argument("--ciclo", default=None, help="ciclo final, p.ej. 2024-2025 (default: Balance!C1)")
    ap.add_argument("--anios", type=int, default=5)
    ap.add_argument("--sin-formulas", action="store_true", help="no leer fórmulas (más rápido)")
    a = ap.parse_args(argv)

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    ts = stamp()

    # 1. carga ---------------------------------------------------------------
    if a.input is None:
        datos = Path(__file__).resolve().parent / "datos"
        xs = [p for p in datos.glob("*.xlsx") if not p.name.startswith("~$")]
        if len(xs) != 1:
            sys.exit(f"Pon exactamente UN archivo .xlsx en {datos} (hay {len(xs)}) o usa --input.")
        a.input = str(xs[0])
    src = load_source(a.input, with_formulas=not a.sin_formulas)
    ciclo_fin = a.ciclo or cell(src.raw_sheets[C.SHEET_BALANCE], "C1")
    ciclo_ini = R.ciclos_atras(ciclo_fin, a.anios)

    # 2. catálogos -----------------------------------------------------------
    cats = build_catalogs(src)
    C.CATALOGOS_DIR.mkdir(exist_ok=True)
    comp_path = C.CATALOGOS_DIR / "carrera_componentes.csv"
    if comp_path.exists():
        componentes = pd.read_csv(comp_path)
    else:
        componentes = build_carrera_componentes(cats["carreras"])
        componentes.to_csv(comp_path, index=False, encoding="utf-8-sig")
        print(f"[catálogo] propuesta de componentes escrita en {comp_path} (revisar)")
    aliases = load_aliases(C.CATALOGOS_DIR / "alias_universidades.csv")

    # 3. limpieza --------------------------------------------------------------
    mov = clean_inout(src.inout, cats, aliases)

    # 4. validación ------------------------------------------------------------
    issues, resumen = run_validations(mov)
    audit = formula_audit(src)

    # 5. reglas de negocio ------------------------------------------------------
    mov = classify(mov)
    bridge = carrera_bridge(mov, componentes)
    alumnos = alumnos_dim(mov)

    # 6. paridad Excel vs Python ----------------------------------------------
    bal_rep = XP.replicate_balance(src.inout, src.raw_sheets[C.SHEET_BALANCE])
    ra_rep = XP.replicate_reporte_anual(src.inout, src.raw_sheets[C.SHEET_REPORTE_ANUAL])
    pim_rep = XP.replicate_pim(src.inout, src.raw_sheets[C.SHEET_PIM])
    xl_ini, xl_fin = bal_rep.attrs["ventana"]
    bal_exp = R.balance_explicado(mov, src.inout, bal_rep, xl_ini, xl_fin)

    # 7. reportes MVP ------------------------------------------------------------
    bal, bal_hist = R.balance_convenios(mov, ciclo_ini, ciclo_fin)
    car, car_tot, car_matriz = R.outgoing_por_carrera(mov, bridge, ciclo_fin, ciclo_fin,
                                                      carreras_cat=cats["carreras"])
    pim, pim_reg, pim_uni = R.pim_reporte(mov, ciclo_fin)

    # comparación agregada balance Excel(réplica) vs Python(corregido)
    xl_by_key = (bal_rep.assign(uni_key=bal_rep["universidad_balance"].map(uni_key))
                 .groupby("uni_key")["py_O"].sum().rename("balance_excel"))
    bal_cmp = bal.set_index("uni_key")[["universidad", "balance", "incluir"]].join(xl_by_key, how="outer")
    bal_cmp["diferencia"] = bal_cmp["balance"].fillna(0) - bal_cmp["balance_excel"].fillna(0)
    bal_cmp = bal_cmp[bal_cmp["diferencia"] != 0].reset_index()

    # 8. exportar ----------------------------------------------------------------
    meta = {
        "generado": datetime.now().isoformat(timespec="seconds"),
        "archivo_fuente": src.source_path,
        "sha256_fuente": src.source_sha256,
        "registros_in_out": len(mov),
        "ventana_balance_python": f"{ciclo_ini} a {ciclo_fin}",
        "ventana_balance_excel": f"{xl_ini} a {xl_fin}",
        "nota": "Archivo generado automáticamente. El archivo fuente no fue modificado.",
    }
    f1 = write_workbook(out / f"diagnostico_calidad_{ts}.xlsx", {
        "resumen_reglas": resumen, "detalle_problemas": issues,
        "formulas_pegadas": audit, "carrera_componentes": componentes,
        "alumnos": alumnos,
    }, meta)
    f2 = write_workbook(out / f"paridad_excel_python_{ts}.xlsx", {
        "balance_replica": bal_rep, "reporte_anual_replica": ra_rep, "pim_replica": pim_rep,
        "balance_por_que_difiere": bal_exp, "balance_excel_vs_corregido": bal_cmp,
    }, meta)
    f3 = write_workbook(out / f"reportes_mvp_{ciclo_fin}_{ts}.xlsx", {
        "1_balance": bal, "1_balance_historico": bal_hist,
        "2_outgoing_carrera": car, "2_totales": car_tot, "2_carrera_x_pais": car_matriz,
        "3_pim_resumen": pim, "3_pim_region": pim_reg, "3_pim_universidad": pim_uni,
    }, meta, index_sheets={"2_carrera_x_pais", "3_pim_region"})
    f4 = out / f"movilidad_limpia_{ts}.csv"
    mov.to_csv(f4, index=False, encoding="utf-8-sig")

    print("\n=== RESUMEN ===")
    print(f"Registros: {len(mov):,} | alumnos únicos: {mov['clave_unica'].nunique():,}")
    print(f"Paridad Balance: {bal_rep['coincide'].sum()}/{len(bal_rep)} filas idénticas")
    print(f"Paridad Reporte Anual: {ra_rep['coincide'].sum()}/{len(ra_rep)} celdas idénticas")
    print(f"Paridad PIM: {pim_rep['coincide'].sum()}/{len(pim_rep)} celdas idénticas")
    print(resumen[["regla", "severidad", "registros", "problema"]].to_string(index=False))
    for f in (f1, f2, f3, f4):
        print("->", f)
    return dict(mov=mov, issues=issues, resumen=resumen, bal=bal, bal_rep=bal_rep, ra=ra_rep, pim_rep=pim_rep,
                bal_exp=bal_exp, bal_cmp=bal_cmp, car=car, car_tot=car_tot, pim=pim, audit=audit,
                componentes=componentes)


if __name__ == "__main__":
    main()
