#自动读取FAST.Farm数据并分列
import originpro as op

# 1. 创建一个新的工作簿
wb = op.new_book('w', lname='A5_1')
wks = wb[0]  # 获取工作簿中的第一个工作表
wks.name = 'A5'


# 获取当前工作簿
#wb = op.find_book('_m/w_', '_name_')
#wb = op.find_book('w', 'A5')
# 在工作簿中添加一个新的工作表，并命名为 "Sheet1"
#new_sheet = wb.add_sheet(name="Sheet2")
#new_sheet = wb.add_sheet(name="A5_M1")
#wks = wb[1]


# 2. 文件路径（修改为你的实际路径）
file_path = 'F:/FASTFarm/E/PrescribedSurge_SERON_A5/FAST.Farm.out'

# 3. 打开文件，手动解析前几行内容
with open(file_path, 'r') as file:
    lines = file.readlines()
    col_names = lines[6].strip().split()  # 第七行为列名称
    units = lines[7].strip().split()      # 第八行为单位
    data_lines = lines[8:]

# 4. 将数据逐行解析为二维列表
data_columns = list(zip(*[map(float, line.strip().split()) for line in data_lines]))

# 5. 将数据逐列导入 Origin 工作表
for col_index, (col_data, col_name) in enumerate(zip(data_columns, col_names)):
    # 将每列数据导入到工作表
    wks.from_list(col_index, list(col_data), lname=col_name)
    # 设置单位（如果存在单位信息）
    if col_index < len(units):
        unit_str = units[col_index]
        op.lt_exec(f'wks.col{col_index + 1}.unit$ = "{unit_str}"')  # 使用 LabTalk 设置单位


# 6. 打印调试信息
print("Column Names:", col_names)
print("Units:", units)
print(f"Successfully imported {len(data_columns)} columns into Origin.")

