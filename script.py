import pandas as pd

def read_excel():
    user_data = pd.read_excel(r"fatigue_risk_calc(decrypted).xlsm", 
    engine='openpyxl',
    sheet_name="Schedule", 
    usecols="A:L",
    header=9).values.tolist()
    for row in user_data:
        day = row[0]
        on_duty = row[1]
        off_duty = row[2]
        job_type_breaks = row[3]
        commuting_time = row[4]
        duty_length = row[5]
        rest_length = row[6]
        average_duty_per_day = row[6]
        cumulative_component = row[8]
        duty_timing_component = row[9]
        job_type_breaks_component = row[10]
        fatigue_index = row[11]
        print(row[1])
        
read_excel()