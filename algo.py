from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook, open

wb = Workbook()

schedule_xlsx = open('data\Schedule_LT.xlsx', read_only=True) 
print(schedule_xlsx)

ws = wb.active