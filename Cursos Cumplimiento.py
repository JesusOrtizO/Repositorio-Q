# Cursos_Cumplimiento.py
from __future__ import annotations

import argparse
import unicodedata
import pandas as pd

ESTADOS_CUMPLIDOS = {"TERMINADO", "CONCLUIDO", "EXENCION"}

COLUMNAS = [
    "Nombre_Colaborador", "Puesto", "Estatus", "Dirección", "Sucursal", "Unidad_Negocio",
    "Estado", "Jefe_Inmediato", "Curso", "Estado_Expediente",
    "Correo", "Jefe_Nombre", "Extra1", "Extra2", "Fecha1", "Fecha2",
]


def norm_series(x: pd.Series) -> pd.Series:
    s = x.fillna("").astype(str).str.strip().str.upper()
    return s.apply(lambda t: "".join(c for c in unicodedata.normalize("NFKD", t) if not unicodedata.combining(c)))


def norm_one(txt: str) -> str:
    s = (txt or "").strip().upper()
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True, help="Ruta al Excel .xlsx")
    ap.add_argument("--sheet", default=None, help="Nombre de hoja (default: primera hoja)")
    ap.add_argument("--skiprows", type=int, default=9, help="Filas a omitir (default: 9)")
    ap.add_argument("--direccion", required=True, help="Dirección objetivo (ej: FINANZAS)")
    ap.add_argument("--courses", nargs="*", default=None, help="Cursos exactos a considerar (opcional)")
    args = ap.parse_args()

    xls = pd.ExcelFile(args.file)
    sheet = args.sheet or xls.sheet_names[0]
    df = xls.parse(sheet, skiprows=args.skiprows)

    df.columns = COLUMNAS[: len(df.columns)]

    df["Dirección_N"] = norm_series(df["Dirección"])
    df["Curso_N"] = norm_series(df["Curso"])
    df["Estado_Expediente_N"] = norm_series(df["Estado_Expediente"])

    dir_n = norm_one(args.direccion)
    df_dir = df[df["Dirección_N"] == dir_n].copy()

    if args.courses:
        courses_n = set(norm_series(pd.Series(args.courses)).tolist())
        df_dir = df_dir[df_dir["Curso_N"].isin(courses_n)]

    # Pendiente = NO está en estados cumplidos (incluye EXENCIÓN)
    df_dir["Es_Cumplido"] = df_dir["Estado_Expediente_N"].isin(ESTADOS_CUMPLIDOS)
    df_pend = df_dir[~df_dir["Es_Cumplido"]].copy()

    print("\n=== Resumen ===")
    print(f"Dirección: {args.direccion}")
    print(f"Total registros: {len(df_dir)}")
    print(f"Cumplidos (incluye EXENCIÓN): {int(df_dir['Es_Cumplido'].sum())}")
    print(f"Pendientes: {len(df_pend)}")

    print("\n=== Pendientes por Dirección ===")
    print(df_pend.groupby("Dirección").size().reset_index(name="Cursos_Pendientes"))

    print("\n=== Pendientes por Sucursal ===")
    print(df_pend.groupby(["Dirección", "Sucursal"]).size().reset_index(name="Cursos_Pendientes"))

    print("\n=== Pendientes por Curso ===")
    print(df_pend.groupby("Curso").size().reset_index(name="Cursos_Pendientes"))


if __name__ == "__main__":
    main()
