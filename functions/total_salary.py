import pandas as pd
import os

from times_process import time_count


def salary_sum():
    # 設定檔案路徑
    file_path = './output/老師個別表.xlsx'

    # 檢查檔案是否存在
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    # 檢查檔案是否為 .xlsx 格式
    if not file_path.endswith('.xlsx'):
        raise ValueError(f"File is not an Excel .xlsx file: {file_path}")

    # 初始化一個空的 DataFrame 來儲存最終的薪水總表
    salary_summary = pd.DataFrame(columns=['老師', '薪水總額'])

    # 使用 pd.read_excel 讀取所有工作表的數據
    try:
        sheet_data_dict = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")

    # 讀取原始 Excel 文件並準備寫入
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

    # 遍歷每個工作表並處理數據
    for sheet, sheet_data in sheet_data_dict.items():
        # 計算薪水總額
        total_salary = sheet_data['薪水'].sum()

        # 構建總薪水行，確保列數量匹配
        total_row = pd.Series(['總薪水'] + [total_salary] + [''] * (len(sheet_data.columns) - 2), index=sheet_data.columns)

        # 將總薪水行添加到工作表數據中
        sheet_data = pd.concat([sheet_data, total_row.to_frame().T], ignore_index=True)

        # 將更新後的數據寫入 Excel 文件
        sheet_data.to_excel(writer, sheet_name=sheet, index=False)

        # 添加到總表
        summary_row = pd.DataFrame({'老師': [sheet], '薪水總額': [total_salary]})
        salary_summary = pd.concat([salary_summary, summary_row], ignore_index=True)

    # 保存新的 Excel 文件
    writer.close()

    # 將薪水總表保存為一個新的 Excel 文件
    summary_file_path = './output/薪水總表.xlsx'
    try:
        salary_summary.to_excel(summary_file_path, index=False)
        print(f"薪水總表已成功保存到 {summary_file_path}")
    except Exception as e:
        raise Exception(f"Error saving summary Excel file: {e}")
    time_count()
