import originpro as op

# 获取当前工作簿
wb = op.find_book()

# 获取两个工作表
sheet1 = wb[0]  # 第一个工作表
sheet2 = wb[1]  # 第二个工作表


# 查找指定长名称的列（结合 LabTalk）
def find_col_by_longname(sheet, longname):
    sheet.activate()  # 激活目标工作表
    for i in range(1, sheet.cols + 1):  # LabTalk 中列索引从 1 开始
        # 使用 LabTalk 获取当前列的长名称
        op.lt_exec(f"lname$ = col({i})[L]$")
        lname = op.get_lt_str("lname$")

        if lname == longname:
            return i - 1  # Python 中索引从 0 开始
    return None


# 提取指定索引的列数据
def extract_col_data(sheet, col_index):
    if col_index is not None:
        return sheet.to_list(col_index)
    else:
        print("Error: Column not found.")
        return None


# 创建新工作表并添加数据
def create_new_sheet_with_data(i, data, new_sheet, sheet):
    data_comment = sheet.name
    sheet.activate()  # 激活目标工作表

    # 使用 LabTalk 获取当前列的单位
    op.lt_exec(f"units$ = col({longname})[U]$")
    UNIT = op.get_lt_str("units$")

    new_sheet.from_list(i, data, lname=f'{longname}', units=f'{UNIT}', comments=f'{data_comment}')  # 将数据填入第 1 列
    print(f"Data has been added to the new sheet: new_sheet")


# 指定要查找的长名称
longname = "WkVelXT1D3"

# 在两个工作表中查找指定长名称的列并提取数据
col_index1 = find_col_by_longname(sheet1, longname)
col_index2 = find_col_by_longname(sheet2, longname)

data1 = extract_col_data(sheet1, col_index1)
data2 = extract_col_data(sheet2, col_index2)

# 创建新的工作表来存放数据
new_sheet = wb.add_sheet(name=f"CombinedData_{longname}'")

# 将提取的数据放入新的工作表中
if data1:
    create_new_sheet_with_data(1, data1, new_sheet, sheet1)

if data2:
    create_new_sheet_with_data(3, data2, new_sheet, sheet2)

new_sheet.activate()