import pandas as pd

def parse_time_duration(duration):
    if isinstance(duration, str):
        hours, minutes, seconds = map(int, duration.split(':'))
    elif isinstance(duration, float):
        if pd.isna(duration):
            hours = 0
            minutes = 0
            seconds = 0
        else:
            hours = int(duration)
            minutes = int((duration - hours) * 60)
            seconds = int(((duration - hours) * 60 - minutes) * 60)
    else:
        hours = duration.hour
        minutes = duration.minute
        seconds = duration.second
    return pd.Timedelta(hours=hours, minutes=minutes, seconds=seconds)

def find_employees(input_file):
    df = pd.read_excel(input_file)
    df.rename(columns={'Position ID': 'PositionID', 'Position Status': 'PositionStatus',
                       'Time': 'TimeIn', 'Time Out': 'TimeOut',
                       'Timecard Hours (as Time)': 'TimecardHours',
                       'Pay Cycle Start Date': 'PayCycleStartDate',
                       'Pay Cycle End Date': 'PayCycleEndDate',
                       'Employee Name': 'EmployeeName', 'File Number': 'FileNumber'}, inplace=True)

    # Sort the DataFrame by 'EmployeeName' and 'TimeIn' columns
    df.sort_values(['EmployeeName', 'TimeIn'], inplace=True)

    # Reset the index of the DataFrame
    df.reset_index(drop=True, inplace=True)

    employees_7_days = []
    employees_less_than_10_hours = []
    employees_more_than_14_hours = []

    for i in range(len(df)):
        current_employee = df.loc[i, 'EmployeeName']
        current_time_in = pd.to_datetime(df.loc[i, 'TimeIn'])
        current_duration_str = df.loc[i, 'TimecardHours']
        current_hours = parse_time_duration(current_duration_str)

        # Check if the employee has worked for 7 consecutive days
        if i < len(df) - 6:
            consecutive_dates = [current_time_in + pd.DateOffset(days=j) for j in range(7)]
            
            # Ensure that 'TimeIn' is not NaT and 'TimeOut' is not null for all 7 days
            valid_consecutive_days = all(
                pd.notnull(df.loc[i+j, 'TimeIn']) and pd.notnull(df.loc[i+j, 'TimeOut'])
                and pd.to_datetime(df.loc[i+j, 'TimeIn']).floor('D') == consecutive_date
                for j, consecutive_date in enumerate(consecutive_dates)
            )

            if valid_consecutive_days:
                employees_7_days.append((current_employee, current_time_in))

        # Check if the employee has less than 10 hours between shifts but greater than 1 hour
        if i < len(df) - 1:
            next_employee = df.loc[i+1, 'EmployeeName']
            next_time_in = pd.to_datetime(df.loc[i+1, 'TimeIn'])
            next_duration_str = df.loc[i+1, 'TimecardHours']
            next_hours = parse_time_duration(next_duration_str)

            if current_employee == next_employee and pd.notnull(next_time_in) and pd.notnull(current_time_in) \
                    and next_time_in - current_time_in < pd.Timedelta(hours=10) \
                    and next_time_in - current_time_in > pd.Timedelta(hours=1):
                employees_less_than_10_hours.append((current_employee, current_time_in))

        # Check if the employee has worked for more than 14 hours in a single shift
        if current_hours > pd.Timedelta(hours=14):
            employees_more_than_14_hours.append((current_employee, current_time_in))

    
    print("Employees who worked for 7 consecutive days:")
    for employee in employees_7_days:
        print(employee[0], employee[1])

    print("\nEmployees who have less than 10 hours between shifts but greater than 1 hour:")
    for employee in employees_less_than_10_hours:
        print(employee[0], employee[1])

    print("\nEmployees who worked for more than 14 hours in a single shift:")
    for employee in employees_more_than_14_hours:
        print(employee[0], employee[1])

if __name__ == '__main__':
    input_file = r'C:\Users\Jairaj\Downloads\Copy of Assignment_Timecard.xlsx'
    find_employees(input_file)
