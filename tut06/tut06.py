#MADHUR GARG 2001CE33

# In the below code we will be reading and writing in multiple .xlsx files.
import pandas as pd
from datetime import datetime
start_time = datetime.now()

#defining a funcion.

def AttEnDeNcE_RePoRt():
    try:
        
        DF1 = pd.read_csv('input_attendance.csv')
        DF2 = pd.read_csv('input_registered_students.csv')

    except:
        print("There was an error reading the file.")


    # Adding empty columns 
    DF1['Roll']  = ''
    DF1['Time']  = ''
    DF1['Date']  = ''
    DF1['Day']   = ''
    DF1['Month'] = ''
    DF1['Year']  = ''