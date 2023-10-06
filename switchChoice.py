
def my_list_shift(array, shift_num):
    return array[shift_num % len(array):] + array[:shift_num % len(array)]


def switch_front_back(sheet, shift_num):
    start_pos = 2
    # 将每列数据存储到List中
    for j in range(start_pos, sheet.max_column):
        column_list = []
        for i in range(start_pos, sheet.max_row + 1):
            column_list.append(sheet.cell(i, j).value)

        column_list = my_list_shift(column_list, shift_num)

        for i in range(start_pos, sheet.max_row + 1):
            sheet.cell(i, j).value = column_list[i - start_pos]


def switch_left_right(sheet, shift_num):
    start_pos = 2
    # 将每列数据存储到List中
    for j in range(start_pos, sheet.max_row + 1):
        row_list = []
        for i in range(start_pos, sheet.max_column):
            row_list.append(sheet.cell(j, i).value)

        row_list = my_list_shift(row_list, shift_num)

        for i in range(start_pos, sheet.max_column):
            sheet.cell(j, i).value = row_list[i - start_pos]


def switch_incline(sheet, shift_num):
    switch_front_back(sheet, shift_num)
    switch_left_right(sheet, shift_num)
