from openpyxl import load_workbook

outputf = "otchet.xlsx"

wb = load_workbook(outputf)
ws = wb['Лист1']

ws['A6'] = 'Отчет стыковок'



wb.save(outputf)
wb.close()
print("otchet gotov")