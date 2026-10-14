"""Prueba de regresión: la réplica de fórmulas debe coincidir al 100% con el Excel.
Uso: python tests/test_parity.py ["ruta/al/excel.xlsx"]  (por defecto, el .xlsx de datos/) """
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config as C
from load_data import load_source
import excel_parity as XP

if len(sys.argv) > 1:
    path = sys.argv[1]
else:
    xs = [p for p in (Path(__file__).resolve().parents[1] / "datos").glob("*.xlsx") if not p.name.startswith("~$")]
    assert len(xs) == 1, "Pon exactamente un .xlsx en datos/ o pasa la ruta como argumento"
    path = str(xs[0])
src = load_source(path, with_formulas=False)
b = XP.replicate_balance(src.inout, src.raw_sheets[C.SHEET_BALANCE])
r = XP.replicate_reporte_anual(src.inout, src.raw_sheets[C.SHEET_REPORTE_ANUAL])
p = XP.replicate_pim(src.inout, src.raw_sheets[C.SHEET_PIM])
for name, df in (("Balance", b), ("Reporte Anual", r), ("PIM", p)):
    ok = df["coincide"].all()
    print(f"{name:14} {'OK' if ok else 'FALLA'}  {df['coincide'].sum()}/{len(df)}")
    assert ok, df[~df["coincide"]]
