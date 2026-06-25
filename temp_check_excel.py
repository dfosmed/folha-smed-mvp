import openpyxl
wb = openpyxl.load_workbook('modelo_contabilização.xlsx')
sheet = wb.active
for row in sheet.iter_rows(min_row=26, max_row=45, values_only=True):
    print(repr(row[0]))
