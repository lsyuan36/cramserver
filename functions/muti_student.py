import pandas as pd


def muti_student():
    # 加载Excel文件
    file_path = 'output/老師個別表.xlsx'
    xls = pd.ExcelFile(file_path)

    # 初始化字典来存储每个老师处理后的数据
    processed_data = {}

    # 合并相同日期时间但不同学生的行，并重新计算薪水
    def merge_rows(df):
        # 重命名列以匹配预期的键
        df = df.rename(columns={'上課日期': '日期', '開始時間': '時間', '學生姓名': '學生', '時長(小時)': '時數'})

        # 按日期、时间分组，并统计唯一学生数量
        merged = df.groupby(['日期', '時間'], as_index=False).agg({
            '學生': lambda x: ', '.join(x.unique()),  # 连接学生名字
            '時數': 'first',  # 取时数的第一个值
        })

        # 根据提供的公式计算薪水
        merged['學生人數'] = merged['學生'].apply(lambda x: len(x.split(', ')))
        merged['薪水'] = (250 + (merged['學生人數'] - 1) * 50) * merged['時數']

        return merged

    # 处理每个表格
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        processed_data[sheet] = merge_rows(df)

    # 创建一个新的Excel文件保存每个老师的处理后数据
    output_file_path = 'output/老師個別表.xlsx'
    with pd.ExcelWriter(output_file_path) as writer:
        for sheet, data in processed_data.items():
            data.to_excel(writer, sheet_name=sheet, index=False)
