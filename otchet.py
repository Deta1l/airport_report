from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

ws.merge_cells('A1:B4')
start = ws['A1']
start.value = "OnD"
start.alignment = Alignment(horizontal="center", vertical="center")

thin = Side(border_style="thin", color="000000")


ws.merge_cells('A5:B9')
ws.merge_cells('C5:C9')


ws.merge_cells('C2:J2')
s7 = ws['C2']
s7.value = "S7"
s7.alignment = Alignment(horizontal="center", vertical="center")

ws.merge_cells('C3:C4')
ws['C3'] = 'Город стыковки'
ws.merge_cells('D3:D4')
ws['D3'] = '№ рейсов'
ws.merge_cells('E3:E4')
ws['E3'] = 'период'
ws.merge_cells('F3:F4')
ws['F3'] = 'дни недели'
ws.merge_cells('G3:G4')
ws['G3'] = 'вылет'
ws.merge_cells('H3:H4')
ws['H3'] = 'прилет'
ws.merge_cells('I3:I4')
ws['I3'] = 'время стык'
ws.merge_cells('J3:J4')
ws['J3'] = 'время полета'

wb.save('otchet.xlsx')
wb.close()
print("отчет готов!")