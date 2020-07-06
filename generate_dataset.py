import numpy as np
from openpyxl import *

excel_path = "C:\\Users\איתן אביטן\Downloads\לימודים\סימולציה לרשתות תקשורת\פרויקט\data.xlsx"

wb = load_workbook(excel_path)
ws = wb["2"]
arrival_time = np.random.poisson(25, 10000)
arrival_time = np.sort(arrival_time)

arrival_time = list(set(arrival_time))
max_value = max(arrival_time)
j = 0
for i in range(6, len(arrival_time)):
    wc_ell = ws.cell(i, 2)
    wc_ell.value = round(arrival_time[j] / max_value, 3)
    j += 1

# wc_cell = ws.cell(6, 3)
# wc_cell.value = ws.cell(6, 2).value
#
# for i in range(7, 10000):
#     wc_cell = ws.cell(i, 3)
#     wc_cell.value = ws.cell(i, 2).value - ws.cell(i - 1, 2).value

wb.save(excel_path)
