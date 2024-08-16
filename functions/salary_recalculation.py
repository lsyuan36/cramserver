from operator import index

import pandas as pd


def salary_recalculation():
    # 讀取用戶上傳的 Excel 檔案
    teacher_salary_file = "./output/老師個別表.xlsx"
    base_salary_file = "./input/基底薪資.xlsx"

    # 讀取老師個別表的所有試算頁
    teacher_sheets = pd.read_excel(teacher_salary_file,sheet_name=None, engine='openpyxl')

    # 讀取基底薪資表
    base_salary_df = pd.read_excel(base_salary_file, engine='openpyxl')

    # 將基底薪資表轉換為字典方便查找
    base_salary_dict = pd.Series(base_salary_df.基底薪資.values, index=base_salary_df.老師姓名).to_dict()

    # 定義未出現在基底薪資表中的老師的薪水
    default_salary = 250

    # 處理每個老師個別表中的頁面
    result_sheets = {}
    for sheet_name, df in teacher_sheets.items():
        if sheet_name in base_salary_dict:
            base_salary = base_salary_dict[sheet_name]
        else:
            base_salary = default_salary

        # 計算薪水 = (基底薪資 + (人數-1)*50) * 時數
        df['薪水'] = (base_salary + (df['學生人數'] - 1) * 50) * df['時數']

        # 將結果保存到新的 DataFrame 中
        result_sheets[sheet_name] = df

    # 將所有結果保存至一個新的 Excel 檔案
    output_file_corrected = "./output/老師個別表.xlsx"
    with pd.ExcelWriter(output_file_corrected) as writer:
        for sheet_name, df in result_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("Salary ReCalculation Done")


salary_recalculation()
