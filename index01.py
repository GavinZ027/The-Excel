import pandas as pd
import os

# 用途：读取Excel，筛选相同名称，对目标文件添加对应编号

# 读取Excel表格中的数据
excel_file = ' example.xlsx'  # Excel文件的路径
sheet_name = 'Sheet1'  # Excel文件中的工作表名

# 使用pandas读取Excel表
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# 打印读取的数据以确认正确性（可选）
# print(df)

# 目标文件夹路径
target_folder = '/Downloads'

# 获取文件夹中的所有文件名（忽略大小写差异）
file_list = [f.lower() for f in os.listdir(target_folder) if not f.startswith('.')]

# 创建名字到序号的映射（忽略大小写差异）
name_to_index = {name.lower(): index for name, index in zip(df['姓名'], df['序号'])}

# 重命名文件
for filename in os.listdir(target_folder):
    # 跳过隐藏文件
    if filename.startswith('.'):
        continue

    # 获取文件名（含扩展名，但忽略大小写比较）
    name_with_ext = filename.lower()

    # 尝试获取对应的序号
    index = name_to_index.get(os.path.splitext(name_with_ext)[0])

    if index is not None:
        # 创建新的文件名（保留原扩展名）
        base, ext = os.path.splitext(filename)
        new_filename = f"{index}_{base}{ext}"

        # 获取旧文件路径和新文件路径
        old_file_path = os.path.join(target_folder, filename)
        new_file_path = os.path.join(target_folder, new_filename)

        # 检查新文件路径是否已存在（避免重命名冲突）
        if not os.path.exists(new_file_path):
            # 重命名文件
            os.rename(old_file_path, new_file_path)
            print(f"Renamed {old_file_path} to {new_file_path}")
        else:
            print(f"Skipped {old_file_path} because {new_file_path} already exists.")
    else:
        print(f"No index found for {filename}")
