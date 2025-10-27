# "Врач, оказавший услугу", "Услуга", "Отделение сотрудника", "Статус услуги", "Вид оплаты"
import openpyxl

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
for row in range(wb_file_sheet_row_begin, wb_file_sheet_row_end + 1):
    # print(row)
    for col in range(wb_file_sheet_col_begin, wb_file_sheet_col_end + 1):
        # print(f'... {row = } ... {col = } ... {wb_s.cell(row, col).coordinate = } ... {wb_s.cell(row, col).value = }')
        pass
    list_all_data_fresh.append([wb_s.cell(row=row, column=2).value,
                                wb_s.cell(row=row, column=3).value,
                                wb_s.cell(row=row, column=10).value,
                                wb_s.cell(row=row, column=22).value
                                ])
print(*list_all_data_fresh, sep='\n')
wb.close()
