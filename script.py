import pandas as pd
import sqlite3

def read_excel():
    user_data = pd.read_excel(r"fatigue_risk_calc(decrypted).xlsm", 
    engine='openpyxl',
    sheet_name="Schedule", 
    usecols="A:L",
    header=9).values.tolist()

    return user_data

def save_data():
    try:
        conn = sqlite3.connect("fatigue.db")
    except:
        print("Database connection error!")

    sql = ""

    day = ""
    on_duty = ""
    off_duty = ""
    job_type_breaks = ""
    commuting_time = ""
    duty_length = ""
    rest_length = ""
    average_duty_per_day = ""
    cumulative_component = ""
    duty_timing_component = ""
    job_type_breaks_component = ""
    fatigue_index = ""

    user_data = read_excel()
    for index, row in enumerate(user_data):
        day = str(row[0])
        on_duty = str(row[1])
        off_duty = str(row[2])
        job_type_breaks = str(row[3])
        commuting_time = str(row[4])
        duty_length = str(row[5])
        rest_length = str(row[6])
        average_duty_per_day = str(row[6])
        cumulative_component = str(row[8])
        duty_timing_component = str(row[9])
        job_type_breaks_component = str(row[10])
        fatigue_index = str(row[11])
        cur = conn.cursor()
        compare_query = "SELECT * FROM fatigue WHERE day = " + f"'{day}'"
        exist_data = cur.execute(compare_query).fetchone()
        if exist_data == None : 
            sql = '''INSERT INTO fatigue(day, on_duty, 
                off_duty, job_type_breaks, commuting_time, duty_length, 
                rest_length, average_duty_per_day, cumulative_component, 
                duty_timing_component, job_type_breaks_component, fatigue_index)
                VALUES'''
            if index != (len(user_data) - 1) :
                sql += '(' + f'"{day}"' + ', ' + f'"{on_duty}"' + ', ' + f'"{off_duty}"' + ', ' + f'"{job_type_breaks}"' + ', ' + f'"{commuting_time}"' + ', ' + f'"{duty_length}"' + ', ' + f'"{rest_length}"' + ', ' + f'"{average_duty_per_day}"' + ', ' + f'"{cumulative_component}"' + ', ' + f'"{duty_timing_component}"' + ', ' + f'"{job_type_breaks_component}"' + ', ' + f'"{fatigue_index}"' + ')' + ', '
            else :
                sql += '(' + f'"{day}"' + ', ' + f'"{on_duty}"' + ', ' + f'"{off_duty}"' + ', ' + f'"{job_type_breaks}"' + ', ' + f'"{commuting_time}"' + ', ' + f'"{duty_length}"' + ', ' + f'"{rest_length}"' + ', ' + f'"{average_duty_per_day}"' + ', ' + f'"{cumulative_component}"' + ', ' + f'"{duty_timing_component}"' + ', ' + f'"{job_type_breaks_component}"' + ', ' + f'"{fatigue_index}"' + ')'

            try:
                cur.execute(sql)
            except sqlite3.Error as e:
                print(e)

        else :
            sql = '''UPDATE fatigue SET '''
            sql += 'day = ' + f'"{day}"' + ', ' + 'on_duty = ' + f'"{on_duty}"' + ', ' + 'off_duty = ' + f'"{off_duty}"' + ', ' + 'job_type_breaks = ' + f'"{job_type_breaks}"' + ', ' + 'commuting_time = ' + f'"{commuting_time}"' + ', ' + 'duty_length = ' + f'"{duty_length}"' + ', ' + 'rest_length = ' + f'"{rest_length}"' + ', ' + 'average_duty_per_day = ' + f'"{average_duty_per_day}"' + ', ' + 'cumulative_component = ' + f'"{cumulative_component}"' + ', ' + 'duty_timing_component = ' + f'"{duty_timing_component}"' + ', ' + 'job_type_breaks_component = ' + f'"{job_type_breaks_component}"' + ', ' + 'fatigue_index = ' + f'"{fatigue_index}"'            
            sql += f''' WHERE day = "{day}"'''

            try:
                cur.execute(sql)
            except sqlite3.Error as e:
                print(e)

    conn.commit()
    print("Successfully saved!")

save_data()