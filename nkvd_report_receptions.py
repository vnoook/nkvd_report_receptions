# "Врач, оказавший услугу", "Услуга", "Отделение сотрудника", "Статус услуги", "Вид оплаты"

import openpyxl

file_xlsx = 'Аналитика2609-2710.xlsx'

# открывается файл
wb_in = openpyxl.load_workbook(file_xlsx)
wb_in_s = wb_in.active

# строка начала и конца
wb_in_s_row_begin = 3
wb_in_s_row_end = wb_in_s.max_row - 1

list_all_data_fresh = []

# получение всех данных из файла и его закрытие, чтобы к нему больше не возвращаться
for row in range(wb_in_s_row_begin, wb_in_s_row_end + 1):
    print(row)
    # wb_in_s.cell(row=1, column=1).value
    # print(f'{row = } ... {wb_in_s_row_begin = } ... {wb_in_s_row_end = }')

    # list_all_data_fresh.append([wb_in_s.cell(row=row, column=wb_in_s).value,
    #                             wb_in_s.cell(row=row, column=wb_in_s_col_2).value,
    #                             wb_in_s.cell(row=row, column=wb_in_s_col_3).value,
    #                             wb_in_s.cell(row=row, column=wb_in_s_col_4).value,
    #                             wb_in_s.cell(row=row, column=wb_in_s_col_5).value
    #                             ])
wb_in.close()
