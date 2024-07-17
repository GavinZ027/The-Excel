import pandas as pd
import os

# Step 1: 读取Excel文件，识别文字
file_path = 'example.xlsx'  # 请替换为你的Excel文件路径
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Step 2: 重命名PDF文件并添加序号
folder_path = 'example/doc'  # 请替换为你的文件夹路径

# 确保文件夹路径末尾有斜杠
if not folder_path.endswith('/'):
    folder_path += '/'

pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# 确保PDF文件按照顺序排序
pdf_files.sort()

# 遍历DataFrame，按照顺序重命名PDF文件并添加序号
for index, row in df.iterrows():
    if index < len(pdf_files):
        old_name = pdf_files[index]
        new_name = f"{index + 1}_{row['名字']}"  # 添加序号和新文件名

        old_file_path = os.path.join(folder_path, old_name)
        new_file_path = os.path.join(folder_path, new_name + '.pdf')  # 添加.pdf扩展名

        # 检查旧文件是否存在，并进行重命名
        if os.path.exists(old_file_path):
            os.rename(old_file_path, new_file_path)
            print(f'Renamed: {old_name} -> {new_name}.pdf')
        else:
            print(f'File not found: {old_name}')
    else:
        print(f'Not enough PDF files for all entries in Excel.')

print("Renaming completed.")
