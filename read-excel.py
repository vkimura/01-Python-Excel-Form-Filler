from xlrd import open_workbook

wb = open_workbook("C:\\Users\\vkimura\\Documents\\Projects\\2022\\05\\09\\2022-06-08-1-colon-20PM.xls")
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
columns = []
print("Columns:")

for i in range(sheet.ncols):
    columns.append(sheet.cell_value(0, i))

# print(columns)
print(sheet.cell_value(1, 5))


