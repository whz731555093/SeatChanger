import openpyxl
import datetime
import os
import shutil
import switchChoice
from tkinter import messagebox


AllTableList = []
try:
    # 筛选出目录下的xlsx文件，获取目标excel名称和文件类型
    for file in os.listdir():
        if file.endswith(".xlsx"):
            AllTableList.append(file)

    if len(AllTableList) == 0:
        raise FileNotFoundError
    elif len(AllTableList) > 1:
        raise IndexError

    OriginalTableName, TableType = AllTableList[0].split(".")
    OriginalTable, LastDate = OriginalTableName.split("_")

    # 获取当前日期和时间并指定格式
    DateFormat = datetime.datetime.now().strftime("%Y-%m-%d")
    NewTableName = OriginalTable + "_" + DateFormat + '.' + TableType

    # 判断该表是否存在，若存在则先删除
    if os.path.exists(NewTableName):
        raise NameError

except IndexError:
    messagebox.showerror("错误！！！", "请确保文件夹中仅有一个Excel表格！")

except FileNotFoundError:
    messagebox.showerror("错误！！！", "文件夹中不存在Excel表格")

except NameError:
    messagebox.showerror("错误！！！", "文件夹中已存在新表格，请修改名称或恢复原表格")

else:
    # 交换座位的方式：
    ChangeModeChoice = 3

    # 复制原来的表格，生成一份新的表格
    shutil.copyfile(OriginalTableName + '.' + TableType, NewTableName)

    # 打开新的表格
    NewTable = openpyxl.load_workbook(NewTableName)

    # 打开新的表格中的指定工作表
    NewSheet = NewTable["CurSeat"]

    ShiftNum = 1  # 表示向前移动的位数
    RowGroup1_Member = 0  # 0表示不分组，其余数表示组1的人数
    # 根据交换模式为新表进行赋值
    if ChangeModeChoice == 1:
        # 模式1：前后轮换
        switchChoice.switch_front_back(NewSheet, ShiftNum, RowGroup1_Member)
    elif ChangeModeChoice == 2:
        # 模式2：左右轮换
        switchChoice.switch_left_right(NewSheet, ShiftNum)
    elif ChangeModeChoice == 3:
        # 模式3：前后左右轮换
        switchChoice.switch_incline(NewSheet, ShiftNum, RowGroup1_Member)

    # 给原来的表重命名，以便于区分
    if OriginalTableName.endswith("(Original)"):
        messagebox.showinfo("提示!", "确认新表格是否有效后，请删除带\'(Original)\'的表格")
    else:
        os.rename(OriginalTableName + '.' + TableType, OriginalTableName + "(Original)." + TableType)

    NewTable.save(NewTableName)
    NewTable.close()


