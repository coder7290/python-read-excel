import pandas as pd
import sqlite3
import tkinter as tk
from tkinter.messagebox import showinfo
from tkinter import *

interface = tk.Tk()
var = StringVar()
label = Label(interface, textvariable=var, relief=RAISED, font=("Arial", 12), bg="green", padx=2)
var.set("Please save changed state by clicking 'Save Data' button.")
label.pack()
label.place(x=149, y=5)
screen_width = interface.winfo_screenwidth()
screen_height = interface.winfo_screenheight()
position_right = int((screen_width - 700)/2)
position_top = int((screen_height - 400)/2)
window_width = 700
window_height = 400
interface.configure(bg='green')
interface.geometry('%dx%d+%d+%d' % (window_width, window_height, position_right, position_top))
interface.title('Fatigue Risk Calc(decrypted)') 
save_state = False

def read_excel():
    user_data = pd.read_excel(r"fatigue_risk_calc(decrypted).xlsm", 
    engine='openpyxl',
    sheet_name="Schedule", 
    usecols="A:L",
    header=9).values.tolist()

    return user_data

def popup_showinfo():
    showinfo("Warning", "Successfully saved!")

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
        create_table_query = '''
                CREATE TABLE IF NOT EXISTS "fatigue" (
                    "id"	INTEGER NOT NULL,
                    "day"	TEXT,
                    "on_duty"	REAL,
                    "off_duty"	TEXT,
                    "job_type_breaks"	TEXT,
                    "commuting_time"	TEXT,
                    "duty_length"	TEXT,
                    "rest_length"	TEXT,
                    "average_duty_per_day"	TEXT,
                    "cumulative_component"	TEXT,
                    "duty_timing_component"	TEXT,
                    "job_type_breaks_component"	TEXT,
                    "fatigue_index"	TEXT,
                    PRIMARY KEY("id" AUTOINCREMENT)
                )'''
        cur.execute(create_table_query)
        
        compare_query = "SELECT * FROM fatigue WHERE day = " + f"'{day}'"
        exist_data = cur.execute(compare_query).fetchone()
        if exist_data == None : 
            sql = '''INSERT INTO fatigue(day, on_duty, 
                off_duty, job_type_breaks, commuting_time, duty_length, 
                rest_length, average_duty_per_day, cumulative_component, 
                duty_timing_component, job_type_breaks_component, fatigue_index)
                VALUES'''

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
    try :
        conn.commit()
        popup_showinfo()
        save_sate = True
        print("Successfully saved!")
    except: 
        save_sate = False

save_button = tk.Button(interface, text='Save Data', width=20, height=2, bg='#54FA9B', command = save_data)
exit_button = tk.Button(interface, text='Exit', width=20, height=2, bg='red', command=interface.destroy)
save_button.pack()
exit_button.pack()
save_button.place(x=280, y=100)
exit_button.place(x=280, y=200)
interface.mainloop()
