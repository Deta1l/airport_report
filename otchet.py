from openpyxl import load_workbook

outputf = "otchet.xlsx"

wb = load_workbook(outputf)
ws = wb['Лист1']




ws.merge_cells('B2:F4')
ws['b2'] = 'Отчет стыковок'




wb.save(outputf)
wb.close()
print("otchet gotov")