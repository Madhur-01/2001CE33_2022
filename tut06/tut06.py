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


     # Concatenating the given data 
    DF1.loc[DF1.Roll  == '', 'Roll'] = DF1.Attendance.str.split().str.get(0)
    DF1.loc[DF1.Time  == '', 'Time'] = DF1.Timestamp.str.split().str.get(1)
    DF1.loc[DF1.Date  == '', 'Date'] = DF1.Timestamp.str.split().str.get(0)
    DF1.loc[DF1.Day   == '', 'Day'] = DF1.Date.str.split('-').str.get(0)
    DF1.loc[DF1.Month == '', 'Month'] = DF1.Date.str.split('-').str.get(1)
    DF1.loc[DF1.Year  == '', 'Year'] = DF1.Date.str.split('-').str.get(2)

    # (Modays and Thurdays)
    VaLiD_dAtEs = []
    FlAg        = ''
    
    # Indexing over the dataframe for getting valid dates, using 'datetime' library for the same.
    for i in DF1.index:
        dt = datetime(int(DF1['Year'][i]), int(DF1['Month'][i]), int(DF1['Day'][i]))

        if (dt.weekday() == 0 or dt.weekday() == 3):

            if (FlAg != DF1['Date'][i]):

                VaLiD_dAtEs.append(DF1['Date'][i])
                FlAg = DF1['Date'][i]
        else:

            DF1['Date'][i] = -1

  
    DF1 = DF1.sort_values('Roll')

    # Creating a dataframe to store to consolidated output.

    FiNaL_df = pd.DataFrame()
    FiNaL_df['Roll'] = ''
    FiNaL_df['Name'] = ''
    for i in range(0, len(VaLiD_dAtEs)):
        FiNaL_df[VaLiD_dAtEs[i]] = ''
    FiNaL_df['Actual Lecture Taken'] = ''
    FiNaL_df['Total Real'] = ''
    FiNaL_df['% Attendance'] = ''

    # Giving structure according to our valid dates stored in the list.
    
    for i in DF2.index:
        FiNaL_df.at[i+1, 'Roll'] = DF2['Roll No'][i]
        FiNaL_df.at[i+1, 'Name'] = DF2['Name'][i]
        FiNaL_df.at[i+1, 'Actual Lecture Taken'] = len(VaLiD_dAtEs)
        FiNaL_df.at[i+1, 'Total Real'] = FiNaL_df.at[i+1, '% Attendance'] = 0

