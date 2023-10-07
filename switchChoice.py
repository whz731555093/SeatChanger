
def my_list_shift(array, shift_num):
    return array[shift_num % len(array):] + array[:shift_num % len(array)]


def switch_front_back(sheet, shift_num, row_group1_member):
    start_pos = 2
    if row_group1_member == 0:
        # 将每列数据存储到List中
        for j in range(start_pos, sheet.max_column):
            column_list = []
            for i in range(start_pos, sheet.max_row + 1):
                column_list.append(sheet.cell(i, j).value)
            column_list = my_list_shift(column_list, shift_num)
            for i in range(start_pos, sheet.max_row + 1):
                sheet.cell(i, j).value = column_list[i - start_pos]
    else:
        # 将每列数据存储到List中
        for j in range(start_pos, sheet.max_column):
            group1_column_list = []
            for i in range(start_pos, start_pos + row_group1_member - 1):
                group1_column_list.append(sheet.cell(i, j).value)
            group1_column_list = my_list_shift(group1_column_list, shift_num)
            for i in range(start_pos, start_pos + row_group1_member - 1):
                sheet.cell(i, j).value = group1_column_list[i - start_pos]

            group2_column_list = []
            for i in range(start_pos + row_group1_member, sheet.max_row + 1):
                group2_column_list.append(sheet.cell(i, j).value)
            print(group2_column_list)
            print(start_pos + row_group1_member)
            print(sheet.max_row + 1)
            group2_column_list = my_list_shift(group2_column_list, shift_num)
            for i in range(start_pos + row_group1_member, sheet.max_row + 1):
                sheet.cell(i, j).value = group2_column_list[i - start_pos - row_group1_member]


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


def switch_incline(sheet, shift_num, row_group1_member):
    switch_front_back(sheet, shift_num, row_group1_member)
    switch_left_right(sheet, shift_num)
