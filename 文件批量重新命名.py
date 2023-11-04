import os
import pandas as pd
from tqdm import tqdm
from datetime import datetime


def rename_files(excel_path, main_folder_path):
    # 读取Excel文件
    df = pd.read_excel(excel_path)

    # 添加新列
    df['修改成功状态'] = ''
    df['修改失败原因'] = ''
    df['修改时间'] = ''

    # 进度条
    for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="重命名文件"):
        original_file_name = row['文件名']
        new_file_name = row['文件新名称']
        file_path = os.path.join(main_folder_path, original_file_name)
        new_file_path = os.path.join(main_folder_path, new_file_name)

        # 尝试重命名文件
        try:
            os.rename(file_path, new_file_path)
            df.at[index, '修改成功状态'] = '成功'
            df.at[index, '修改时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        except Exception as e:
            df.at[index, '修改成功状态'] = '失败'
            df.at[index, '修改失败原因'] = str(e)
            df.at[index, '修改时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # 保存修改后的Excel文件
    df.to_excel(excel_path, index=False)


if __name__ == '__main__':
    excel_path = input('请输入Excel文件的路径和名称：')
    main_folder_path = input('请输入需要修改的文件的主路径：')
    rename_files(excel_path, main_folder_path)
