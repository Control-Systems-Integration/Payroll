import pandas as pd
import datetime
import numpy as np
from openpyxl import load_workbook

def round_time(dt):
    # Calculate the number of minutes past the last 15-minute mark
    minutes = (dt.minute % 15) * 60 + dt.second

    # If the number of minutes is less than 7, round down; otherwise, round up
    if minutes < 7 * 60:
        dt = dt - datetime.timedelta(minutes=dt.minute % 15, seconds=dt.second)
    else:
        dt = dt + datetime.timedelta(minutes=15 - dt.minute % 15, seconds=-dt.second)

    return dt

# Load the Excel file
df = pd.read_excel('C:\\test\\Payroll.xlsx')

# Load the Employee Name data from Test3.xlsx file
df_test3 = pd.read_excel('C:\\test\\Test3.xlsx')
employee_names_test3 = df_test3['Employee Name'].unique()

# Convert the 'Ticket Date', 'Clock In', and 'Clock Out' columns to datetime
df['Ticket Date'] = pd.to_datetime(df['Ticket Date'])
df['Clock In'] = pd.to_datetime(df['Clock In']).apply(round_time)
df['Clock Out'] = pd.to_datetime(df['Clock Out']).apply(round_time)

# Calculate the total hours worked for each job
df['Total Hours Worked'] = (df['Clock Out'] - df['Clock In']).dt.total_seconds() / 3600

# Create a list to hold the results
results = []

# Create a DataFrame to hold missing Employees
missing_employees = pd.DataFrame(columns=['Employee Name', 'Ticket Date'])

# Group by 'Employee Name', 'JobNo|Customer|Description', 'Agency', and 'Ticket Date'
grouped = df.groupby(['Employee Name', 'JobNo|Customer|Description', 'Agency', df['Ticket Date'].dt.date])

# Group by 'Employee Name', 'JobNo|Customer|Description', 'Agency', and 'Ticket Date'
grouped = df.groupby(['Employee Name', 'JobNo|Customer|Description', 'Agency', df['Ticket Date'].dt.date])

for name, group in grouped:
    if name[0] not in employee_names_test3:
        missing_employees = missing_employees.append({'Employee Name': name[0], 'Ticket Date': name[3]}, ignore_index=True)

for name, group in grouped:
    total_hours = group['Total Hours Worked'].sum()

    # Deduct 30 minutes for lunch break if the employee worked for more than 5 hours
    if total_hours > 5:
        total_hours -= 0.5

    regular_hours = round(min(8, total_hours), 2)
    overtime_hours = round(max(0, total_hours - 8), 2)

    #SB
    print(group.columns)

    # If the Ticket Date is on a Saturday or Sunday, all hours are overtime
    if group['Ticket Date'].dt.dayofweek.iloc[0] >= 5:
        overtime_hours = round(total_hours, 2)
        regular_hours = 0

    results.append(pd.DataFrame({
        'Employee Name': [name[0]],
        'Employee ID': [group['Employee ID'].iloc[0]],
        'JobNo|Customer|Description': [name[1]],
        'Agency': [name[2]],
        'Ticket Date': [name[3]],
        'Day': [group['Ticket Date'].dt.day_name().iloc[0]],
        'Regular Hours': [regular_hours],
        'Overtime Hours': [overtime_hours],
        'Clock In': [group['Clock In'].iloc[0]],
        'Clock Out': [group['Clock Out'].iloc[0]],
        'Supervisors Name': [group['Supervisors Name'].iloc[0]],
        'PM Assigned': [group['PM Assigned'].iloc[0]],
        'Email': [group['Email'].iloc[0]],
        'WTL Approved': [group['WTL Approved'].iloc[0]]
    }))

# Concatenate all the results into a single dataframe
result = pd.concat(results)

# Save the result to a new Excel file
result.to_excel('C:\\test\\result.xlsx', index=False)

# Load the workbook
book = load_workbook('C:\\test\\result.xlsx')

# Access the sheet by name or index
sheet1 = book['Sheet1']

# Set column widths
sheet1.column_dimensions['A'].width = 29.71
sheet1.column_dimensions['B'].width = 12.57
sheet1.column_dimensions['C'].width = 72.71
sheet1.column_dimensions['D'].width = 17
sheet1.column_dimensions['E'].width = 14.29
sheet1.column_dimensions['F'].width = 15.14
sheet1.column_dimensions['G'].width = 15
sheet1.column_dimensions['H'].width = 16.71
sheet1.column_dimensions['I'].width = 20.29
sheet1.column_dimensions['J'].width = 27.71
sheet1.column_dimensions['K'].width = 22.43
sheet1.column_dimensions['L'].width = 23.57
sheet1.column_dimensions['M'].width = 35.29
sheet1.column_dimensions['N'].width = 20.57

# Save the modified workbook
book.save('C:\\test\\result.xlsx')









