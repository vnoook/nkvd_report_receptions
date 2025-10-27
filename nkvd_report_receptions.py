# "Врач, оказавший услугу", "Услуга", "Отделение сотрудника", "Вид оплаты"
import openpyxl
import pandas as pd

# файл данных
file_xlsx = 'Аналитика2609-2710.xlsx'

# открывается файл
wb = openpyxl.load_workbook(file_xlsx)
wb_s = wb.active

# строки начала и конца
wb_file_sheet_row_begin = 3
wb_file_sheet_row_end = wb_s.max_row - 1
wb_file_sheet_col_begin = wb_s.min_column
wb_file_sheet_col_end = wb_s.max_column

# получение всех данных из файла и его закрытие, чтобы к нему больше не возвращаться
list_all_data_fresh = []
otdel_dict = {}
for row in range(wb_file_sheet_row_begin, wb_file_sheet_row_end + 1):
    otdel = wb_s.cell(row=row, column=10).value
    if otdel_dict.get(otdel) is None:
        otdel_dict[otdel] = 1
    else:
        otdel_dict[otdel] = otdel_dict[otdel] + 1
    doc = wb_s.cell(row=row, column=2).value
    service = wb_s.cell(row=row, column=3).value
    payment = wb_s.cell(row=row, column=22).value
    list_all_data_fresh.append([doc, service, otdel, payment])

print(*list_all_data_fresh, sep='\n')
wb.close()

