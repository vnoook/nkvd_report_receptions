# "Врач, оказавший услугу", "Услуга", "Отделение сотрудника", "Статус услуги", "Вид оплаты"

import openpyxl

file_xlsx = 'Аналитика2609-2710.xlsx'

# открывается файл
wb = openpyxl.load_workbook(file_xlsx)
wb_s = wb.active

# строка начала и конца
wb_file_sheet_row_begin = 3
wb_file_sheet_row_end = wb_s.max_row - 1
wb_file_sheet_col_begin = wb_s.min_column
wb_file_sheet_col_end = wb_s.max_column

list_all_data_fresh = []

# получение всех данных из файла и его закрытие, чтобы к нему больше не возвращаться
for row in range(wb_file_sheet_row_begin, wb_file_sheet_row_end + 1):
    print(row)
    for col in range(wb_file_sheet_col_begin, wb_file_sheet_col_end + 1):
        # print(f'... {row = } ... {col = } ... {wb_s.cell(row, col).value}')
        list_all_data_fresh.append([wb_in_s.cell(row=row, column=2).value,
                                    wb_in_s.cell(row=row, column=3).value,
                                    wb_in_s.cell(row=row, column=10).value,
                                    wb_in_s.cell(row=row, column=12).value,
                                    wb_in_s.cell(row=row, column=22).value
                                    ])
        # if col in (2,3,10,12,22):
        #     print(f'... {row = } ... {col = } ... {wb_s.cell(row, col).coordinate = } ... {wb_s.cell(row, col).value = }')





    # openpyxl.utils.cell.coordinate_from_string(wb_file_IC_s.cell(1, col_IC).coordinate)
    # wb_full_s.cell(wb_full_s.max_row, wb_full_s.max_column).coordinate
    # openpyxl.utils.cell.coordinate_from_string(self.wb_file_GASPS_s.cell(1, col_GASPS).coordinate)[0]
    # wb_from_s.cell(xl_row, xl_col).coordinate
    #
    # indexR_IC = wb_IC_cells_range.index(row_in_range_IC)
    # indexC_IC = row_in_range_IC.index(cell_in_row_IC)
    # wb_IC_cell_coord = wb_IC_cells_range[indexR_IC][indexC_IC].coordinate



# wb.close()
