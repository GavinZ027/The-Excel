import pandas as pd
import os
import re

# 用途：读取Excel，根据表格内容顺序，对目标文件进行重新命名及编号

# Step 1: 读取Excel文件，识别文字
file_path = 'example.xlsx'  # 请替换为你的Excel文件路径
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Step 2: 重命名文件并添加序号
folder_path = '/Downloads'  # 请替换为你的文件夹路径

# 确保文件夹路径末尾有斜杠
if not folder_path.endswith('/'):
    folder_path += '/'

# 获取文件夹中的所有非隐藏文件
all_files = [f for f in os.listdir(folder_path)
             if os.path.isfile(os.path.join(folder_path, f)) and not f.startswith('.')]

# 自然排序函数
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

# 确保文件按照自然顺序排序
all_files.sort(key=natural_sort_key)

# 遍历DataFrame，按照顺序重命名文件并添加序号
for index, row in df.iterrows():
    if index < len(all_files):
        old_name = all_files[index]
        new_name = f"{index + 1}_{row['name']}"  # 添加序号和新文件名

        # 获取文件扩展名
        file_extension = os.path.splitext(old_name)[1]

        old_file_path = os.path.join(folder_path, old_name)
        new_file_path = os.path.join(folder_path, new_name + file_extension)  # 保留原扩展名

        # 检查旧文件是否存在，并进行重命名
        if os.path.exists(old_file_path):
            os.rename(old_file_path, new_file_path)
            print(f'Renamed: {old_name} -> {new_name}{file_extension}')
        else:
            print(f'File not found: {old_name}')
    else:
        print(f'Not enough files for all entries in Excel.')

print("Renaming completed.")
