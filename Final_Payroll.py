import pandas as pd

# Load the data
df = pd.read_excel('C:/test/Payroll.xlsx')

# Initialize the dictionary
cumulative_hours = {name: 0 for name in df['Employee Name'].unique()}

# Loop over the rows in the DataFrame
for index, row in df.iterrows():
    employee_name = row['Employee Name']
    lunch_adjusted = row['Lunch Adjusted']

    # Add the lunch adjusted hours to the cumulative hours
    cumulative_hours[employee_name] += lunch_adjusted

    # Calculate the regular time and overtime
    if cumulative_hours[employee_name] <= 40:
        df.loc[index, 'Regular Time'] = lunch_adjusted
        df.loc[index, 'Overtime'] = 0
    else:
        if cumulative_hours[employee_name] - lunch_adjusted < 40:
            regular_time = 40 - (cumulative_hours[employee_name] - lunch_adjusted)
            df.loc[index, 'Regular Time'] = regular_time
            df.loc[index, 'Overtime'] = lunch_adjusted - regular_time
        else:
            df.loc[index, 'Regular Time'] = 0
            df.loc[index, 'Overtime'] = lunch_adjusted

# Write the corrected data back to the file
df.to_excel('C:/test/Payroll.xlsx', index=False)
