
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Reporte de Cursos Pendientes", layout="wide")
st.title("📊 Reporte de Cursos No Concluidos por Departamento")

@st.cache_data
def cargar_datos(archivo):
    xls = pd.ExcelFile(archivo)
    df = xls.parse('Sheet1')  # Hoja del archivo de ejemplo
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

    df_filtrado = df[
        (df['Dirección'].str.upper() == departamento_objetivo.upper()) &
        (~df['Estado_Expediente'].str.upper().isin(["TERMINADO", "CONCLUIDO"]))
    ]

    st.subheader(f"Resumen general para: {departamento_objetivo}")
    st.metric(label="Total de cursos pendientes", value=df_filtrado.shape[0])

    reporte_departamentos = df_filtrado.groupby(['Dirección', 'Sucursal']).size().reset_index(name='Cursos_Pendientes')

    st.subheader("Cursos pendientes por Departamento (Sucursal)")
    st.dataframe(reporte_departamentos, use_container_width=True)

    st.subheader("Visualización")
    st.bar_chart(reporte_departamentos.set_index('Sucursal')['Cursos_Pendientes'])

else:
    st.info("Por favor sube un archivo Excel para continuar.")