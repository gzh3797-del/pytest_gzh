"""Print each Excel row number alongside its 用例编号 (column C)."""
from pathlib import Path
import openpyxl

_EXCEL = Path(__file__).parent.parent / 'Manual_testcase' / 'AcuHMI-1-7_用户管理用例.xlsx'

wb = openpyxl.load_workbook(_EXCEL)
ws = wb['Sheet1']

for row in ws.iter_rows(min_row=2):
    r = row[0].row
    case_id = row[2].value or ""
    if case_id:
        print(f"row {r:4d}  {case_id}")
