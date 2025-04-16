import pandas as pd
import streamlit as st

st.set_page_config(page_title="Reporte de Cursos Pendientes", layout="wide")
st.title("📊 Reporte de Cursos No Concluidos por Departamento")

@st.cache_data
def cargar_datos(archivo):
    xls = pd.ExcelFile(archivo)
    df = xls.parse(xls.sheet_names[0])  # Cargar la primera hoja
    columnas = [
        'Nombre_Colaborador', 'Puesto', 'Estatus', 'Dirección', 'Sucursal', 'Unidad_Negocio',
        'Estado', 'Jefe_Inmediato', 'Curso', 'Estado_Expediente',
        'Correo', 'Jefe_Nombre', 'Extra1', 'Extra2', 'Fecha1', 'Fecha2'
    ]
    df.columns = columnas[:len(df.columns)]
    return df

archivo = st.file_uploader("Sube el archivo Excel de cursos:", type=["xlsx"])

if archivo is not None:
    df = cargar_datos(archivo)
    df['Dirección'] = df['Dirección'].fillna('')
    df['Estado_Expediente'] = df['Estado_Expediente'].fillna('')

    departamentos_disponibles = df['Dirección'].dropna().unique()
    departamento_objetivo = st.selectbox("Selecciona una Dirección a analizar:", sorted(departamentos_disponibles))

    df_total = df[df['Dirección'].str.upper() == departamento_objetivo.upper()]
    df_filtrado = df_total[~df_total['Estado_Expediente'].str.upper().isin(["TERMINADO", "CONCLUIDO"])]

    st.subheader(f"Resumen general para: {departamento_objetivo}")
    col1, col2 = st.columns(2)
    col1.metric(label="📋 Total de cursos registrados", value=df_total.shape[0])
    col2.metric(label="⏳ Cursos no concluidos", value=df_filtrado.shape[0])

    reporte_departamentos = df_filtrado.groupby(['Dirección', 'Sucursal']).size().reset_index(name='Cursos_Pendientes')
    st.subheader("Cursos pendientes por Departamento (Sucursal)")
    st.dataframe(reporte_departamentos, use_container_width=True)

    st.subheader("Visualización por Departamento")
    st.bar_chart(reporte_departamentos.set_index('Sucursal')['Cursos_Pendientes'])

    st.subheader("Resumen de Cursos Pendientes")
    reporte_cursos = df_filtrado.groupby('Curso').size().reset_index(name='Total_Pendientes')
    st.dataframe(reporte_cursos, use_container_width=True)
    st.bar_chart(reporte_cursos.set_index('Curso')['Total_Pendientes'])

else:
    st.info("Por favor sube un archivo Excel para continuar.")