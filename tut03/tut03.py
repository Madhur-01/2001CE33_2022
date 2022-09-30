#MADHUR GARG; 2001CE33

#importing pandas
import pandas as pd

#reading the input file
df = pd.read_excel("input_octant_transition_identify.xlsx")

#data preprocessing
df.at[0,'U_avg']  = df['U'].mean()
df.at[0,'V_avg']  = df['V'].mean()
df.at[0,'W_avg']  = df['W'].mean()


df["U'"] = df['U'] - df.at[0,'U_avg']
df["V'"] = df['V'] - df.at[0,'V_avg']
df["W'"] = df['W'] - df.at[0,'W_avg']


#definig a function to categorise data in different octant
def octant(x,y,z) :
    if x>0 :
        if y>0:
            if z>0 :
                return +1
            else :
                return -1
        else :
            if z>0 :
                return +4
            else :
                return -4
    else :
        if y> 0 :
            if z>0:
                return +2
            else :
                return -2
        else :
            if z>0:
                return +3
            else:
                return -3
#applying the above function           
df['octant']         =   df.apply([lambda x : octant(x["U'"],x["V'"],x["W'"])], axis=1)

#leaving an empty column
df.at[1,''] = 'User Input'

#counting individual octant uing value_counts function
df.at[0,'Octant ID'] =   'Overall Count'
df.at[0,'1']         =   df['octant'].value_counts()[+1]
df.at[0,'-1']        =   df['octant'].value_counts()[-1]
df.at[0,'2']         =   df['octant'].value_counts()[+2]
df.at[0,'-2']        =   df['octant'].value_counts()[-2]
df.at[0,'3']         =   df['octant'].value_counts()[+3]
df.at[0,'-3']        =   df['octant'].value_counts()[-3]
df.at[0,'4']         =   df['octant'].value_counts()[+4]
df.at[0,'-4']        =   df['octant'].value_counts()[-4]

mod = 5000

df.at[1,'Octant ID'] = 'Mod '+ str(mod)

size = len(df['octant'])
m=0
#using a while loop to split the data 
while(size>0):
    temp1 = mod
    if m == 0: #starting from value 0
        x = 0
    else:
        x = m*temp1 

    if size<mod:
        mod = size
        size = 0
        
    y = m*temp1+mod - 1
    
    
    #inserting range and their corresponding data
    m1 = str(x)
    m2= str(y)
    df.at[m+2,'Octant ID'] = m1 +'-'+m2 

    #making a new data frame
    df2 = df.loc[x:y] 
   
    df.at[m+2,'1'] = df2['octant'].value_counts()[+1]
    df.at[m+2,'2'] = df2['octant'].value_counts()[+2]
    df.at[m+2,'3'] = df2['octant'].value_counts()[+3]
    df.at[m+2,'4'] = df2['octant'].value_counts()[+4]
    df.at[m+2,'-1'] = df2['octant'].value_counts()[-1]
    df.at[m+2,'-2'] = df2['octant'].value_counts()[-2]
    df.at[m+2,'-3'] = df2['octant'].value_counts()[-3]
    df.at[m+2,'-4'] = df2['octant'].value_counts()[-4]
    

    m = m + 1
    size = size - mod
    
