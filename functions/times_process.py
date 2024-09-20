import pandas as pd
from datetime import datetime, timedelta

def time_count():
    # Load the Excel file
    file_path = './output/老師個別表.xlsx'
    excel_data = pd.ExcelFile(file_path)


    # Function to calculate end time
    def calculate_end_time(start_time, duration):
        start = datetime.strptime(start_time, "%H:%M")
        end = start + timedelta(hours=duration)
        return end.strftime("%H:%M")


    # Create a new Excel writer to overwrite the original file
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # Process each sheet and save it back
        for sheet_name in excel_data.sheet_names:
            # Load each sheet into a DataFrame
            df = excel_data.parse(sheet_name)

            # Remove the row that contains '總薪水'
            df = df[df['日期'] != '總薪水'].copy()

            # Create new columns for start time and end time
            df['開始時間'] = df['時間']
            df['結束時間'] = df.apply(lambda row: calculate_end_time(row['時間'], row['時數']), axis=1)

            # Drop the old '時間' column
            df = df.drop(columns=['時間'])

            # Reorder the columns as per the user's request
            df = df[['日期', '開始時間', '結束時間', '時數', '學生人數', '薪水']]

            # Recalculate total salary (薪水)
            total_salary = df['薪水'].sum()

            # Add a row for total salary at the bottom
            total_row = ['總薪水'," ", " ", " ", " ", total_salary]
            df.loc[len(df)] = total_row

            # Reset the index without adding the index column
            df = df.reset_index(drop=True)

            # Save the modified DataFrame back to the original Excel file
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Indicate the process is done and file saved
    print("All sheets have been processed and saved back to the original file.")
