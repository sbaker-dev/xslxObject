from xlsxObject import XlsxObject

a = XlsxObject(r"C:\Users\Samuel\PycharmProjects\xslxObject\Tests\Example.xlsx")

print(a)

print(a.sheet_names)
print(a.sheet_headers)
print(a.sheet_col_count)
print(a.sheet_row_count)

print(a.sheet_data[0])


for row in a.sheet_data[0].row_data:
    print(row)

print(a.sheet_data[1])
