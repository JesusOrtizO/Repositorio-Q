
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.drawing.image import Image

# Crear libro de Excel
wb = Workbook()
ws = wb.active
ws.title = "Dashboard Dirección"

# Estilos
bold_font = Font(bold=True)
header_font = Font(bold=True, color="FFFFFF")
purple_fill = PatternFill("solid", fgColor="800080")
center = Alignment(horizontal="center")

# Título
ws.merge_cells('A1:D1')
ws['A1'] = "CURSOS PENDIENTES POR DIRECCIÓN"
ws['A1'].font = Font(bold=True, size=14)
ws['A1'].alignment = center

# Subtítulo
ws.merge_cells('A2:D2')
ws['A2'] = "DIRECCIÓN: Operaciones"
ws['A2'].font = Font(bold=True)
ws['A2'].alignment = center

# Fecha y total
ws['F1'] = "16 de enero de 2025"
ws['F1'].alignment = Alignment(horizontal="right")
ws['F2'] = "Total pendientes"
ws['F3'] = 6
ws['F3'].font = Font(bold=True, size=14, color="FFFFFF")
ws['F3'].fill = PatternFill("solid", fgColor="FF4D6D")
ws['F3'].alignment = center

# Tabla de cursos
curso_data = [
    ["Curso", "Pendientes", "% Pendientes"],
    ["Política: Conflicto de Intereses 2024", 6, "100 %"],
    ["Medidas de seguridad en el puesto de trabajo", 0, "0 %"],
    ["PCI DSS VERSIÓN 4.0", 0, "0 %"],
    ["Protección de Datos Personales 2024", 0, "0 %"]
]

start_row = 5
for i, row in enumerate(curso_data):
    for j, value in enumerate(row):
        cell = ws.cell(row=start_row + i, column=j + 1, value=value)
        if i == 0:
            cell.font = header_font
            cell.fill = purple_fill
        cell.alignment = center

# Tabla de áreas
area_data = [
    ["Área", "Pendientes"],
    ["Operaciones Emisión", 1],
    ["Sistemas Desarrollo SEL", 1],
    ["Operaciones Autos Golfo", 1],
    ["Administración de Emisión", 1],
    ["Operaciones Centro de Contacto", 1],
    ["Otra área", 1]
]

start_row_area = 12
ws.cell(row=start_row_area, column=1, value="Cursos Pendientes por Área").font = bold_font

for i, row in enumerate(area_data):
    for j, value in enumerate(row):
        cell = ws.cell(row=start_row_area + 1 + i, column=j + 1, value=value)
        if i == 0:
            cell.font = header_font
            cell.fill = purple_fill
        cell.alignment = center

# Ajustes finales
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 20
ws.column_dimensions["F"].width = 25

# Guardar archivo
excel_path = "/mnt/data/Reporte_Visual_Direccion_Operaciones.xlsx"
wb.save(excel_path)

excel_path