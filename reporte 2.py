import pandas as pd
import streamlit as st
import unicodedata
import re

st.set_page_config(page_title="Reporte de Cursos", layout="wide")
st.title("📊 Reporte de Cursos por Dirección / Departamento")

# ---- Config: estados que cuentan como "cumplido" ----
ESTADOS_CUMPLIDOS = {"TERMINADO", "CONCLUIDO", "EXENCION"}  # normalizados sin acento

def norm_text(x: pd.Series) -> pd.Series:
    """Mayúsculas, trim, y quitar acentos para comparar de forma estable."""
    s = x.fillna("").astype(str).str.strip().str.upper()
    return s.apply(
        lambda t: "".join(c for c in unicodedata.normalize("NFKD", t) if not unicodedata.combining(c))
    )

@st.cache_data
def cargar_datos(archivo):
    xls = pd.ExcelFile(archivo)
    # OJO: en tu primer script usas skiprows=9; aquí no. Ajusta si tu archivo real lo requiere.
    df = xls.parse(xls.sheet_names[0], skiprows=9)

    columnas = [
        'Nombre_Colaborador', 'Puesto', 'Estatus', 'Dirección', 'Sucursal', 'Unidad_Negocio',
        'Estado', 'Jefe_Inmediato', 'Curso', 'Estado_Expediente',
        'Correo', 'Jefe_Nombre', 'Extra1', 'Extra2', 'Fecha1', 'Fecha2'
    ]
    df.columns = columnas[:len(df.columns)]

    # columnas normalizadas para filtros
    df["Dirección_N"] = norm_text(df["Dirección"])
    df["Sucursal_N"] = norm_text(df["Sucursal"])
    df["Curso_N"] = norm_text(df["Curso"])
    df["Estado_Expediente_N"] = norm_text(df["Estado_Expediente"])

    # flags
    df["Es_Cumplido"] = df["Estado_Expediente_N"].isin(ESTADOS_CUMPLIDOS)
    df["Es_Pendiente"] = ~df["Es_Cumplido"]

    return df

archivo = st.file_uploader("Sube el archivo Excel de cursos:", type=["xlsx"])

if archivo is None:
    st.info("Por favor sube un archivo Excel para continuar.")
    st.stop()

df = cargar_datos(archivo)

# ---- Selector de dirección ----
direcciones = sorted([d for d in df["Dirección"].dropna().unique() if str(d).strip() != ""])
direccion_objetivo = st.selectbox("Selecciona una Dirección a analizar:", direcciones)

df_dir = df[df["Dirección_N"] == norm_text(pd.Series([direccion_objetivo])).iloc[0]].copy()

# ---- Selector de cursos (tipo filtro) ----
st.markdown("### 🎯 Filtro de cursos")
colA, colB = st.columns([2, 1])

with colB:
    modo = st.radio("Modo de filtro", ["Selección exacta", "Búsqueda por texto"], horizontal=False)

df_base_cursos = df_dir.copy()

if modo == "Selección exacta":
    cursos = sorted([c for c in df_base_cursos["Curso"].dropna().unique() if str(c).strip() != ""])
    with colA:
        cursos_sel = st.multiselect(
            "Selecciona los cursos a considerar (vacío = todos):",
            options=cursos
        )
    if cursos_sel:
        cursos_sel_n = set(norm_text(pd.Series(cursos_sel)).tolist())
        df_base_cursos = df_base_cursos[df_base_cursos["Curso_N"].isin(cursos_sel_n)]

else:
    with colA:
        texto = st.text_input("Palabras clave (separadas por coma) para buscar en el nombre del curso:")
    keywords = [t.strip() for t in (texto or "").split(",") if t.strip()]
    if keywords:
        # OR entre keywords; escapamos regex
        pat = "|".join(re.escape(k) for k in keywords)
        df_base_cursos = df_base_cursos[df_base_cursos["Curso"].fillna("").astype(str).str.contains(pat, case=False, regex=True)]

# ---- Métricas ----
df_total = df_base_cursos
df_pend = df_base_cursos[df_base_cursos["Es_Pendiente"]]
df_ok = df_base_cursos[df_base_cursos["Es_Cumplido"]]

st.subheader(f"Resumen para: {direccion_objetivo}")
m1, m2, m3 = st.columns(3)
m1.metric("📋 Total registros (cursos)", df_total.shape[0])
m2.metric("✅ Cumplidos (Terminado/Concluido/Exención)", df_ok.shape[0])
m3.metric("⏳ Pendientes", df_pend.shape[0])

# ---- Reporte por sucursal (departamento) ----
st.subheader("Cursos por Departamento (Sucursal)")
rep_suc = (
    df_total.groupby(["Dirección", "Sucursal"])
    .agg(
        Total_Cursos=("Curso", "size"),
        Cumplidos=("Es_Cumplido", "sum"),
        Pendientes=("Es_Pendiente", "sum"),
    )
    .reset_index()
    .sort_values(["Pendientes", "Total_Cursos"], ascending=False)
)

st.dataframe(rep_suc, use_container_width=True)
st.bar_chart(rep_suc.set_index("Sucursal")["Pendientes"])

# ---- Resumen por curso ----
st.subheader("Resumen por Curso")
rep_curso = (
    df_total.groupby("Curso")
    .agg(
        Total=("Curso", "size"),
        Cumplidos=("Es_Cumplido", "sum"),
        Pendientes=("Es_Pendiente", "sum"),
    )
    .reset_index()
    .sort_values(["Pendientes", "Total"], ascending=False)
)
st.dataframe(rep_curso, use_container_width=True)
st.bar_chart(rep_curso.set_index("Curso")["Pendientes"])

# ---- Detalle por colaborador ----
st.subheader("🔎 Detalle (incluye estado)")
detalle = df_total[["Sucursal", "Nombre_Colaborador", "Curso", "Estado_Expediente", "Es_Pendiente"]].copy()
detalle = detalle.sort_values(by=["Sucursal", "Nombre_Colaborador", "Curso"])
st.dataframe(detalle, use_container_width=True)

# ---- Solo pendientes (vista rápida) ----
with st.expander("Ver solo pendientes"):
    detalle_p = df_pend[["Sucursal", "Nombre_Colaborador", "Curso", "Estado_Expediente"]].sort_values(
        by=["Sucursal", "Nombre_Colaborador", "Curso"]
    )
    st.dataframe(detalle_p, use_container_width=True)
