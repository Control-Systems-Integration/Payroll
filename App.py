import pandas as pd
import numpy as np
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from openpyxl import load_workbook
from openpyxl.styles import Font, Color

# Create a Tkinter root window
root = Tk()
# Hide the root window
root.withdraw()

try:
    # Open the file picker dialog
    clockIn_File = askopenfilename()
    payRoll_File = askopenfilename()

    # Check if a file path was selected
    if not clockIn_File or not payRoll_File:
        print("No file selected.")
        raise SystemExit

    # Read the Excel files
    df1 = pd.read_excel(clockIn_File)
    df2 = pd.read_excel(payRoll_File)

    # Keep only 'Employee name', 'Ticket Date' and 'Agency' columns in df2
    df2 = df2[
        ['Employee Name', 'Ticket Date', 'Agency', 'Clock-In ID', 'Supervisors Name', 'PM Assigned',
         'JobNo|Customer|Description', 'Email', 'WTL Start Date', 'WTL End Date']]

    # Convert 'Ticket Date' to datetime in both dataframes
    df1['Ticket Date'] = pd.to_datetime(df1['Ticket Date'])
    df2['Ticket Date'] = pd.to_datetime(df2['Ticket Date'])

    # If 'Clock In' and 'Clock Out' are not datetime, convert them
    df1['Clock In'] = pd.to_datetime(df1['Clock In'])
    df1['Clock Out'] = pd.to_datetime(df1['Clock Out'])

    # Calculate 'Actual Hours Worked' as the difference between 'Clock Out' and 'Clock In', converted to hours
    df1['Actual Hours Worked'] = (df1['Clock Out'] - df1['Clock In']).dt.total_seconds() / 3600
    # # Add 0.5 to 'Actual Hours Worked' column if there is a WTL Start Date and WTL End Date
    # df1.loc[~df1['WTL Start Date'].isnull() & ~df1['WTL End Date'].isnull(), 'Actual Hours Worked'] += 0.5
    # Taking off the half hour for lunch
    df1['Actual Hours Worked'] = (df1['Actual Hours Worked'] - .5)

    # Calculate overtime hours if 'Actual Hours Worked' is greater than 8
    df1['Overtime'] = np.where(df1['Actual Hours Worked'] > 8, df1['Actual Hours Worked'] - 8, 0)

    # Add 'Day of the Week' column
    df1['Day of the Week'] = df1['Ticket Date'].dt.day_name()

    # # Round 'Actual Hours Worked' to 2 decimal places
    # df1['Actual Hours Worked'] = df1['Actual Hours Worked'].round(2)

    # Merge dataframes based on 'Employee name' and 'Ticket Date'
    merged_df = pd.merge(df1, df2, on=['Employee Name', 'Ticket Date'], how='left')

    # Add 0.5 to 'Actual Hours Worked' column if there is a WTL Start Date and WTL End Date
    merged_df.loc[
        ~merged_df['WTL Start Date'].isnull() & ~merged_df['WTL End Date'].isnull(), 'Actual Hours Worked'] += 0.5

    # If 'Agency' is blank, fill with 'CSI'
    merged_df['Agency'] = merged_df['Agency'].fillna('CSI')

    # Create a new dataframe for rows with errors
    errors_df = merged_df[(merged_df['Clock In'].isna()) |
                          (merged_df['Clock Out'].isna()) |
                          (merged_df['Actual Hours Worked'] < 8)].copy()

    # Create the 'Error Description' column
    def generate_error_desc(row):
        if pd.isnull(row['Clock In']) and pd.isnull(row['Clock Out']):
            return 'No Clock In or Clock Out Time'
        elif pd.isnull(row['Clock In']):
            return 'No Clock In'
        elif pd.isnull(row['Clock Out']):
            return 'No Clock Out'
        elif row['Actual Hours Worked'] < 8:
            return 'Less Than 8 Hours'
        else:
            return np.nan


    errors_df['Error Description'] = errors_df.apply(generate_error_desc, axis=1)

    # Convert 'Ticket Date' back to 'mm/dd/yyyy' format
    merged_df['Ticket Date'] = merged_df['Ticket Date'].dt.strftime('%m/%d/%Y')
    errors_df['Ticket Date'] = errors_df['Ticket Date'].dt.strftime('%m/%d/%Y')

    # Write the dataframes into a new Excel file with two sheets
    with pd.ExcelWriter('C:/Users/tj-fo/Desktop/Test/Payroll.xlsx') as writer:
        merged_df.to_excel(writer, sheet_name='Payroll', index=False)
        errors_df.to_excel(writer, sheet_name='Errors', index=False)

    # SB
    # Load the workbook
    wb = load_workbook('C:/Users/tj-fo/Desktop/Test/Payroll.xlsx')

    # SB
    # Select the sheets
    sheet1 = wb['Payroll']

    # SB
    # Create a red bold font
    red_bold_font = Font(color="FF0000", bold=True)

    # Check each cell in column E (5th column) for both sheets
    for sheet in [sheet1]:
        for row in sheet.iter_rows(min_row=2, min_col=4, max_col=5):
            for cell in row:
                if cell.column_letter == 'D' and (cell.value is None or cell.value == ''):
                    cell.value = 'Clock In Time?'
                    cell.font = red_bold_font
                elif cell.column_letter == 'E' and (cell.value is None or cell.value == ''):
                    cell.value = 'Clock Out Time?'
                    cell.font = red_bold_font

    # SB
    # Set column widths
    sheet1.column_dimensions['A'].width = 11.26
    sheet1.column_dimensions['B'].width = 26.14
    sheet1.column_dimensions['C'].width = 31.86
    sheet1.column_dimensions['D'].width = 19
    sheet1.column_dimensions['E'].width = 20.43
    sheet1.column_dimensions['F'].width = 18.71
    sheet1.column_dimensions['G'].width = 20.86
    sheet1.column_dimensions['H'].width = 18
    sheet1.column_dimensions['I'].width = 32.57
    sheet1.column_dimensions['J'].width = 28.71
    sheet1.column_dimensions['K'].width = 20
    sheet1.column_dimensions['L'].width = 22.86
    sheet1.column_dimensions['M'].width = 33.86
    sheet1.column_dimensions['N'].width = 71.57
    sheet1.column_dimensions['O'].width = 31.86
    sheet1.column_dimensions['P'].width = 20.43
    sheet1.column_dimensions['Q'].width = 20.43

    # Select the sheets
    sheet2 = wb['Errors']

    # Set column widths
    sheet2.column_dimensions['A'].width = 11.26
    sheet2.column_dimensions['B'].width = 26.14
    sheet2.column_dimensions['C'].width = 31.86
    sheet2.column_dimensions['D'].width = 19
    sheet2.column_dimensions['E'].width = 20.43
    sheet2.column_dimensions['F'].width = 18.71
    sheet2.column_dimensions['G'].width = 20.86
    sheet2.column_dimensions['H'].width = 18
    sheet2.column_dimensions['I'].width = 32.57
    sheet2.column_dimensions['J'].width = 28.71
    sheet2.column_dimensions['K'].width = 20
    sheet2.column_dimensions['L'].width = 22.86
    sheet2.column_dimensions['M'].width = 33.86
    sheet2.column_dimensions['N'].width = 71.57
    sheet2.column_dimensions['O'].width = 31.86
    sheet2.column_dimensions['P'].width = 20.43
    sheet2.column_dimensions['Q'].width = 20.43
    sheet2.column_dimensions['R'].width = 20.43

    # Save workbook
    wb.save('C:/Users/tj-fo/Desktop/Test/Payroll.xlsx')

except Exception as e:
    print("An error occurred:", str(e))
    raise SystemExit
