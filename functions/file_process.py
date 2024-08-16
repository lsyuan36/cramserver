import pandas as pd
import muti_student
import salary_recalculation

# Load the uploaded Excel file
file_path = './input/上課記錄 (回覆).xlsx'
df = pd.read_excel(file_path)

# Convert time columns to datetime
df['開始時間'] = pd.to_datetime(df['開始時間'], format='%H:%M:%S').dt.time
df['結束時間'] = pd.to_datetime(df['結束時間'], format='%H:%M:%S').dt.time

# Calculate duration in hours
df['開始時間'] = pd.to_datetime(df['上課日期'].astype(str) + ' ' + df['開始時間'].astype(str))
df['結束時間'] = pd.to_datetime(df['上課日期'].astype(str) + ' ' + df['結束時間'].astype(str))
df['時長(小時)'] = ((df['結束時間'] - df['開始時間']).dt.total_seconds() / 3600).round(1)

# Split records with multiple students into separate rows for student_df
df_students = df.assign(學生姓名=df['學生姓名'].str.split(', ')).explode('學生姓名')

# Create the student table with student names
student_columns = ['學生姓名', '上課日期', '上課進度', '考試分數/回家作業分數', '開始時間', '結束時間', '時長(小時)',
                   '回家作業', '上課狀況', '老師姓名']
student_df = df_students[student_columns].copy()
student_df.columns = ['學生姓名', '上課日期', '上課內容', '考試分數', '開始時間', '結束時間', '時長(小時)', '回家作業',
                      '狀況', '老師姓名']
student_df['上課費用'] = None

# Create the teacher table with teacher names
teacher_columns = ['老師姓名', '上課日期', '開始時間', '結束時間', '學生姓名', '時長(小時)']
teacher_df = df[teacher_columns].copy()
teacher_df.columns = ['老師姓名', '上課日期', '開始時間', '結束時間', '學生姓名', '時長(小時)']
teacher_df['薪水'] = (250 + ((teacher_df['學生姓名'].str.split(',').apply(len) - 1) * 50)) * teacher_df['時長(小時)']

# Convert dates and times to the desired format
student_df['上課日期'] = pd.to_datetime(student_df['上課日期']).dt.strftime('%Y/%m/%d')
student_df['開始時間'] = pd.to_datetime(student_df['開始時間']).dt.strftime('%H:%M')
student_df['結束時間'] = pd.to_datetime(student_df['結束時間']).dt.strftime('%H:%M')

teacher_df['上課日期'] = pd.to_datetime(teacher_df['上課日期']).dt.strftime('%Y/%m/%d')
teacher_df['開始時間'] = pd.to_datetime(teacher_df['開始時間']).dt.strftime('%H:%M')
teacher_df['結束時間'] = pd.to_datetime(teacher_df['結束時間']).dt.strftime('%H:%M')

# Write each student's records to their own sheet with formatted date and time
student_writer = pd.ExcelWriter('./output/學生個別表.xlsx', engine='xlsxwriter')
for student_name, student_data in student_df.groupby('學生姓名'):
    student_data.drop(columns=['學生姓名']).to_excel(student_writer, sheet_name=student_name, index=False)

student_writer.close()

# Write each teacher's records to their own sheet with formatted date and time
teacher_writer = pd.ExcelWriter('./output/老師個別表.xlsx', engine='xlsxwriter')
for teacher_name, teacher_data in teacher_df.groupby('老師姓名'):
    teacher_data.to_excel(teacher_writer, sheet_name=teacher_name, index=False)

teacher_writer.close()

print('檔案處理完成')

# Calculate total salary for each teacher
teacher_total_salary = teacher_df.groupby('老師姓名')['薪水'].sum().reset_index()
teacher_total_salary.columns = ['老師姓名', '總薪水']

# Write the total salary to a new sheet
salary_writer = pd.ExcelWriter('./output/薪水總表.xlsx', engine='xlsxwriter')
teacher_total_salary.to_excel(salary_writer, sheet_name='薪水總表', index=False)
salary_writer.close()

print('薪水總表處理完成')


# Function to get records for a specific month
def get_records_for_month(month):
    filtered_teacher_df = teacher_df[teacher_df['上課日期'].str.startswith(month)]
    filtered_student_df = student_df[student_df['上課日期'].str.startswith(month)]

    # Write filtered records to new files
    filtered_teacher_writer = pd.ExcelWriter(f'./output/老師個別表.xlsx', engine='xlsxwriter')
    for teacher_name, teacher_data in filtered_teacher_df.groupby('老師姓名'):
        teacher_data.to_excel(filtered_teacher_writer, sheet_name=teacher_name, index=False)
    filtered_teacher_writer.close()

    filtered_student_writer = pd.ExcelWriter(f'./output/學生個別表.xlsx', engine='xlsxwriter')
    for student_name, student_data in filtered_student_df.groupby('學生姓名'):
        student_data.drop(columns=['學生姓名']).to_excel(filtered_student_writer, sheet_name=student_name, index=False)
    filtered_student_writer.close()

    print(f'{month}月份的表處理完成')


# Get user input for month
month = input('請輸入月份 (格式為 YYYY/MM): ')
get_records_for_month(month)
muti_student.multi_student()
salary_recalculation.salary_recalculation()