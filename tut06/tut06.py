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

    try:
        # Main Loop (iterating for each roll number one by one)
        for i in DF2.index:

            # A variable dataframe which will store data for every roll number
            dF3 = pd.DataFrame()
            dF3['Date'] = ''
            dF3['Roll'] = ''
            dF3['Name'] = ''
            dF3['Total Attendance Count'] = 0
            dF3['Real'] = ''
            dF3['Duplicate'] = ''
            dF3['Invalid'] = ''
            dF3['Absent'] = ''

            dF3.at[0, 'Roll'] = DF2['Roll No'][i]
            dF3.at[0, 'Name'] = DF2['Name'][i]

            # Writing valid dates in the dataframe
            for j in range(1, len(VaLiD_dAtEs)+1):
                dF3.at[j, 'Date'] = VaLiD_dAtEs[j-1]
                dF3.at[j, 'Total Attendance Count'] = dF3.at[j, 'Real'] = dF3.at[j,'Duplicate'] = dF3.at[j, 'Invalid'] = dF3.at[j, 'Absent'] = 0

            # Temporary variable
            temp = DF2['Roll No'][i]

            # Counting attendance according to the given criteria
            for j in DF1.index:
                if (temp == DF1['Roll'][j]):
                    if (DF1['Date'][j] != -1):
                        ind = VaLiD_dAtEs.index(DF1['Date'][j])
                        dF3.at[ind+1, 'Total Attendance Count'] += 1
                        if (DF1['Time'][j] >= '14:00' and DF1['Time'][j] <= '15:00'):
                            if (dF3['Real'][ind+1] == 0):
                                dF3.at[ind+1, 'Real'] += 1
                            else:
                                dF3.at[ind+1, 'Duplicate'] += 1
                        else:
                            dF3.at[ind+1, 'Invalid'] += 1

            # Marking absent from the above evaluated attendance.
            for j in range(1, len(VaLiD_dAtEs)+1):
                if (dF3['Real'][j] == 0):
                    dF3.at[j, 'Absent'] = 1
                    FiNaL_df.at[i+1, VaLiD_dAtEs[j-1]] = 'A'
                else:
                    FiNaL_df.at[i+1, VaLiD_dAtEs[j-1]] = 'P'
                    FiNaL_df.at[i+1, 'Total Real'] += 1
                FiNaL_df.at[i+1, '% Attendance'] = round(
                    FiNaL_df.at[i+1, 'Total Real'] * 100 / FiNaL_df.at[i+1, 'Actual Lecture Taken'], 2)

            # Saving the excel file for each roll number.
            dF3.to_excel('output/'+temp+'.xlsx', index=False)
    except:
        print("Index overflow, check the range again.")

    try:
        
        FiNaL_df.to_excel('output/attendance_report_consolidated.xlsx', index=False)
    except:
        print("Error saving the excel file in cosolidated format.")


AttEnDeNcE_RePoRt()

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))