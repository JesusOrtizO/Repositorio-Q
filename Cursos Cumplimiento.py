import pandas as pd

# === 1. Cargar el archivo Excel ===
file_path = "E:/Proyectos Q/Copia de 2025_Certificación_Q_-_Colaboradores_Q_y_Qsalud_colaboradores_y_ODS_20250411_11_40_55_AM.xlsx"
xls = pd.ExcelFile(file_path)

# === 2. Leer la hoja principal omitiendo las filas de encabezado falsas ===
df = xls.parse('2024 Certificación Q - Colabora', skiprows=9)

# === 3. Renombrar columnas ===
columnas = [
    'Nombre_Colaborador', 'Puesto', 'Estatus', 'Dirección', 'Sucursal', 'Unidad_Negocio',
    'Estado', 'Jefe_Inmediato', 'Curso', 'Estado_Expediente',
    'Correo', 'Jefe_Nombre', 'Extra1', 'Extra2', 'Fecha1', 'Fecha2'
]
df.columns = columnas[:len(df.columns)]

# === 4. Elegir el departamento a analizar ===
departamento_objetivo = "FINANZAS"  # <-- Cambia aquí para otro departamento

# === 5. Filtrar datos del departamento con cursos NO concluidos ===
df_filtrado = df[
    (df['Dirección'].str.upper() == departamento_objetivo.upper()) &
    (~df['Estado_Expediente'].str.upper().isin(["TERMINADO", "CONCLUIDO"]))
]

# === 6. Reporte general por dirección (puede ser útil si hay varias direcciones) ===
reporte_direccion = df_filtrado.groupby('Dirección').size().reset_index(name='Cursos_Pendientes')
print("\n=== Reporte de Cursos Pendientes por Dirección ===")
print(reporte_direccion)

# === 7. Reporte detallado por departamento/sucursal dentro de la dirección ===
reporte_departamentos = df_filtrado.groupby(['Dirección', 'Sucursal']).size().reset_index(name='Cursos_Pendientes')
print("\n=== Reporte de Cursos Pendientes por Departamento ===")
print(reporte_departamentos)

# === 8. (Opcional) Exportar a Excel o CSV ===
# reporte_departamentos.to_excel("reporte_cursos_pendientes_por_departamento.xlsx", index=False)
# reporte_direccion.to_excel("reporte_cursos_pendientes_por_direccion.xlsx", index=False)

