import openpyxl
import win32com.client

def replace_in_excel(file_path, replacements):
    wb = openpyxl.load_workbook(file_path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for old, new in replacements:
                        cell.value = cell.value.replace(old, new)
    wb.save(file_path)

def replace_in_excel_xls(file_path, replacements):
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(file_path)
    for ws in wb.Worksheets:
        for old, new in replacements:
            ws.Cells.Replace(What=old, Replacement=new, LookAt=1)  # xlPart
    wb.Save()
    wb.Close()
    excel.Quit() 