# app.py
from __future__ import annotations

import re
import unicodedata
from datetime import date
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

from reporte import crear_reporte_excel

st.set_page_config(page_title="Reporte de Cursos", layout="wide")
st.title("📊 Reporte de Cursos por Dirección / Departamento")

# Estados cumplidos (normalizados sin acentos)
ESTADOS_CUMPLIDOS = {"TERMINADO", "CONCLUIDO", "EXENCION"}

# Columnas mínimas requeridas para el análisis
REQ_CANON = ["Nombre_Colaborador", "Dirección", "Sucursal", "Curso", "Estado_Expediente"]

# ✅ Encabezados típicos del reporte (los tuyos)
# Usaremos "keywords" para detectarlos sin depender de la fila exacta.
SYNONYMS: Dict[str, List[str]] = {
    "Nombre_Colaborador": [
        "Usuario - Nombre completo del usuario",
        "Nombre_Colaborador", "Nombre", "Colaborador", "Empleado", "Nombre del colaborador",
    ],
    "Dirección": [
        "Usuario - Dirección",
        "Dirección", "Direccion", "Dirección / Área", "Direccion / Area", "Dirección General",
    ],
    "Sucursal": [
        "Usuario - Departamento",          # lo más común
        "Usuario - Departamento Parent",   # alternativa jerárquica
        "Sucursal", "Departamento", "Área", "Area", "Unidad", "Subárea", "Subarea",
    ],
    "Curso": [
        "Capacitación - Título de la capacitación",
        "Curso", "Nombre_Curso", "Curso Asignado", "Capacitación", "Capacitacion", "Nombre del curso",
    ],
    "Estado_Expediente": [
        "Registro de capacitación - Estado del expediente",
        "Estado_Expediente", "Estatus_Expediente", "Estado", "Estatus", "Status", "Avance",
    ],
}

def norm_one(txt: str) -> str:
    s = (txt or "").strip().upper()
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def norm_series(x: pd.Series) -> pd.Series:
    s = x.fillna("").astype(str).str.strip().str.upper()
    return s.apply(lambda t: "".join(c for c in unicodedata.normalize("NFKD", t) if not unicodedata.combining(c)))

def build_colmap(columns: List[str]) -> Dict[str, str]:
    cols = list(columns)
    cols_norm = [norm_one(str(c)) for c in cols]
    idx_by_norm = {c: i for i, c in enumerate(cols_norm)}

    colmap: Dict[str, str] = {}
    for canon, alts in SYNONYMS.items():
        found = None
        for alt in alts:
            alt_n = norm_one(alt)
            if alt_n in idx_by_norm:
                found = cols[idx_by_norm[alt_n]]
                break
        if found:
            colmap[canon] = found
    return colmap

def score_row_as_header(row_values: List, expected_norm: set[str]) -> int:
    row_norm = {norm_one(str(x)) for x in row_values if str(x).strip() not in ("", "nan", "None")}
    return len(row_norm & expected_norm)

def detectar_fila_header(raw_preview: pd.DataFrame, max_scan_rows: int = 80) -> Optional[int]:
    """
    Escanea filas 0..max_scan_rows buscando la que más se parezca a un encabezado.
    Usa coincidencias con SYNONYMS (normalizados).
    """
    expected_norm = {norm_one(x) for x in sum(SYNONYMS.values(), [])}

    best_i, best_score = None, 0
    limit = min(max_scan_rows, len(raw_preview))
    for i in range(limit):
        score = score_row_as_header(raw_preview.iloc[i].tolist(), expected_norm)
        if score > best_score:
            best_score = score
            best_i = i

    # Umbral: con tus headers largos normalmente score >= 3-5 fácil
    return best_i if best_score >= 2 else None

def construir_df_minimo(df: pd.DataFrame, colmap: Dict[str, str]) -> pd.DataFrame:
    out = pd.DataFrame({
        "Nombre_Colaborador": df[colmap["Nombre_Colaborador"]],
        "Dirección": df[colmap["Dirección"]],
        "Sucursal": df[colmap["Sucursal"]],
        "Curso": df[colmap["Curso"]],
        "Estado_Expediente": df[colmap["Estado_Expediente"]],
    })

    out["Dirección_N"] = norm_series(out["Dirección"])
    out["Sucursal_N"] = norm_series(out["Sucursal"])
    out["Curso_N"] = norm_series(out["Curso"])
    out["Estado_Expediente_N"] = norm_series(out["Estado_Expediente"])

    out["Es_Cumplido"] = out["Estado_Expediente_N"].isin(ESTADOS_CUMPLIDOS)
    out["Es_Pendiente"] = ~out["Es_Cumplido"]
    return out

@st.cache_data
def cargar_excel_autoheader(archivo, sheet_name: Optional[str]) -> Tuple[pd.DataFrame, Dict[str, str], List[str], Optional[int]]:
    xls = pd.ExcelFile(archivo)
    hoja = sheet_name if sheet_name else xls.sheet_names[0]

    # Leer preview sin header para detectar fila
    preview = xls.parse(hoja, header=None, nrows=120)
    header_row = detectar_fila_header(preview, max_scan_rows=80)

    if header_row is None:
        # Fallback: intenta header=0 normal
        df = xls.parse(hoja)
        colmap = build_colmap(list(df.columns))
        return df, colmap, xls.sheet_names, None

    # Carga todo sin header y fuerza la fila detectada como header
    raw = xls.parse(hoja, header=None)
    hdr = raw.iloc[header_row].copy().fillna("").astype(str).str.strip()

    # Si hay merges que dejan blancos, rellenar a la derecha (suave)
    hdr = hdr.replace("", pd.NA).ffill().fillna("")

    df = raw.iloc[header_row + 1 :].copy()
    df.columns = hdr.tolist()
    df = df.reset_index(drop=True).dropna(axis=1, how="all")

    colmap = build_colmap(list(df.columns))
    return df, colmap, xls.sheet_names, header_row

# =========================
# UI
# =========================
archivo = st.file_uploader("Sube el archivo Excel de cursos:", type=["xlsx"])
if archivo is None:
    st.info("Por favor sube un archivo Excel para continuar.")
    st.stop()

with st.sidebar:
    st.header("⚙️ Lectura del archivo")
    sheet_name = st.text_input("Nombre de hoja (vacío = primera hoja):", value="")

df_raw, colmap_auto, sheet_names, header_row_detected = cargar_excel_autoheader(archivo, sheet_name.strip() or None)

with st.expander("🧪 Debug lectura (opcional)"):
    st.write("Hoja(s):", sheet_names)
    st.write("Fila header detectada (0-index):", header_row_detected)
    st.write("Columnas detectadas (primeras 30):", list(df_raw.columns)[:30])
    st.dataframe(df_raw.head(8), use_container_width=True)

missing = [c for c in REQ_CANON if c not in colmap_auto]

# Mapeo manual si falta algo (a prueba de balas)
if missing:
    st.warning(
        "No pude reconocer automáticamente todas las columnas. "
        "Selecciona manualmente cuáles columnas corresponden a cada campo."
    )
    cols = list(df_raw.columns)
    manual = {}
    for canon in REQ_CANON:
        default = colmap_auto.get(canon, None)
        options = ["(No existe)"] + cols
        idx = options.index(default) if default in cols else 0
        manual[canon] = st.selectbox(f"Columna para **{canon}**", options=options, index=idx)

    still_missing = [k for k, v in manual.items() if v == "(No existe)"]
    if still_missing:
        st.error(f"Faltan columnas necesarias: {still_missing}. No puedo continuar.")
        st.stop()

    colmap = {k: v for k, v in manual.items()}
else:
    colmap = colmap_auto

df = construir_df_minimo(df_raw, colmap)

# ---- Selector de dirección ----
direcciones = sorted([d for d in df["Dirección"].dropna().unique() if str(d).strip() != ""])
if not direcciones:
    st.error("No se encontraron Direcciones en el archivo.")
    st.stop()

direccion_objetivo = st.selectbox("Selecciona una Dirección a analizar:", direcciones)
df_dir = df[df["Dirección_N"] == norm_one(direccion_objetivo)].copy()

# ---- Selector de cursos ----
st.markdown("### 🎯 Filtro de cursos")
colA, colB = st.columns([2, 1])

with colB:
    modo = st.radio("Modo de filtro", ["Selección exacta", "Búsqueda por texto"], horizontal=False)

df_base = df_dir.copy()

if modo == "Selección exacta":
    cursos = sorted([c for c in df_base["Curso"].dropna().unique() if str(c).strip() != ""])
    with colA:
        cursos_sel = st.multiselect("Selecciona los cursos (vacío = todos):", options=cursos)
    if cursos_sel:
        cursos_sel_n = set(norm_series(pd.Series(cursos_sel)).tolist())
        df_base = df_base[df_base["Curso_N"].isin(cursos_sel_n)]
else:
    with colA:
        texto = st.text_input("Palabras clave (coma separadas) para buscar en el nombre del curso:")
    keywords = [t.strip() for t in (texto or "").split(",") if t.strip()]
    if keywords:
        pat = "|".join(re.escape(k) for k in keywords)
        df_base = df_base[df_base["Curso"].fillna("").astype(str).str.contains(pat, case=False, regex=True)]

# ---- Métricas ----
df_total = df_base
df_pend = df_base[df_base["Es_Pendiente"]]
df_ok = df_base[df_base["Es_Cumplido"]]

st.subheader(f"Resumen para: {direccion_objetivo}")
m1, m2, m3 = st.columns(3)
m1.metric("📋 Total registros (cursos)", df_total.shape[0])
m2.metric("✅ Cumplidos (Terminado/Concluido/Exención)", df_ok.shape[0])
m3.metric("⏳ Pendientes", df_pend.shape[0])

# ---- Reporte por sucursal ----
st.subheader("Cursos por Departamento (Sucursal)")
rep_suc = (
    df_total.groupby(["Dirección", "Sucursal"], dropna=False)
    .agg(Total_Cursos=("Curso", "size"), Cumplidos=("Es_Cumplido", "sum"), Pendientes=("Es_Pendiente", "sum"))
    .reset_index()
    .sort_values(["Pendientes", "Total_Cursos"], ascending=False)
)

st.dataframe(rep_suc, use_container_width=True)
if not rep_suc.empty:
    st.bar_chart(rep_suc.set_index("Sucursal")["Pendientes"])

# ---- Resumen por curso ----
st.subheader("Resumen por Curso")
rep_curso = (
    df_total.groupby("Curso", dropna=False)
    .agg(Total=("Curso", "size"), Cumplidos=("Es_Cumplido", "sum"), Pendientes=("Es_Pendiente", "sum"))
    .reset_index()
    .sort_values(["Pendientes", "Total"], ascending=False)
)

st.dataframe(rep_curso, use_container_width=True)
if not rep_curso.empty:
    st.bar_chart(rep_curso.set_index("Curso")["Pendientes"])

# ---- Detalle ----
st.subheader("🔎 Detalle (incluye estado)")
detalle = df_total[["Sucursal", "Nombre_Colaborador", "Curso", "Estado_Expediente", "Es_Pendiente"]].copy()
detalle = detalle.sort_values(by=["Sucursal", "Nombre_Colaborador", "Curso"])
st.dataframe(detalle, use_container_width=True)

with st.expander("Ver solo pendientes"):
    detalle_p = df_pend[["Sucursal", "Nombre_Colaborador", "Curso", "Estado_Expediente"]].sort_values(
        by=["Sucursal", "Nombre_Colaborador", "Curso"]
    )
    st.dataframe(detalle_p, use_container_width=True)

# =========================
# Export Excel visual
# =========================
st.markdown("### 📥 Exportar reporte visual a Excel")

rep_curso_export = rep_curso[rep_curso["Pendientes"] > 0].copy()
tabla_cursos = [["Curso", "Pendientes", "% Pendientes"]]
for _, r in rep_curso_export.iterrows():
    curso = r["Curso"]
    pend = int(r["Pendientes"])
    total = int(r["Total"])
    pct = f"{(pend / total * 100):.1f} %" if total else "0 %"
    tabla_cursos.append([curso, pend, pct])
if len(tabla_cursos) == 1:
    tabla_cursos.append(["(Sin pendientes)", 0, "0 %"])

rep_suc_export = rep_suc[rep_suc["Pendientes"] > 0].copy()
tabla_areas = [["Área", "Pendientes"]]
for _, r in rep_suc_export.iterrows():
    tabla_areas.append([r["Sucursal"], int(r["Pendientes"])])
if len(tabla_areas) == 1:
    tabla_areas.append(["(Sin pendientes)", 0])

excel_bytes = crear_reporte_excel(
    direccion=direccion_objetivo,
    total_pendientes=int(df_pend.shape[0]),
    tabla_cursos=tabla_cursos,
    tabla_areas=tabla_areas,
    fecha_texto=date.today().strftime("%d/%m/%Y"),
)

st.download_button(
    label="⬇️ Descargar reporte Excel",
    data=excel_bytes,
    file_name=f"Reporte_Visual_{norm_one(direccion_objetivo).replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
