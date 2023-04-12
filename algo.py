from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import load_workbook, open


#schedule_xlsx = open('data\Schedule_LT.xlsx', read_only=True) 


schedule_xlsx  = 'data\Schedule_LT.xlsx'

wb = load_workbook(schedule_xlsx)

ws = wb.active
m_row = ws.max_row

for i in range(1, m_row + 1):
    cell_obj = ws.cell(row = i, column = 2)
    print(cell_obj.value)























'''
def dfs(graph, start, end, n, path=None):
    if path is None:
        path = []
    path = path + [start]
    if start == end or len(path) == n:
        return [path]
    paths = []
    for node in graph[start]:
        if node not in path:
            new_paths = dfs(graph, node, end, n, path)
            for new_path in new_paths:
                paths.append(new_path)
    return paths


graph = {'a':['b'],
         'b':['a', 'c', 'c', 'c'],
         'c':['b', 'b', 'b']
    
}

start = 'a'
end = 'c'
n = 4

print(dfs(graph, start, end, n+2))
'''