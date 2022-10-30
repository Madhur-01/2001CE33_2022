#MADHUR GARG, 2001CE33

#importing numpy and pandas
import numpy as np
import pandas as pd

#reading the input file
df = pd.read_excel("octant_input.xlsx")

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
df.at[0,'1']        =   df['octant'].value_counts()[+1]
df.at[0,'-1']        =   df['octant'].value_counts()[-1]
df.at[0,'2']        =   df['octant'].value_counts()[+2]
df.at[0,'-2']        =   df['octant'].value_counts()[-2]
df.at[0,'3']        =   df['octant'].value_counts()[+3]
df.at[0,'-3']        =   df['octant'].value_counts()[-3]
df.at[0,'4']        =   df['octant'].value_counts()[+4]
df.at[0,'-4']        =   df['octant'].value_counts()[-4]


#asking user for input
mod = int(input('enter the value of mod: '))


df.at[1,'Octant ID'] = 'Mod '+ str(mod)


size = len(df['octant'])
m=0
#using a while loop to split the data 
while(size>0):
    temp = mod
    if m == 0: #starting from value 0
        x = 0
    else:
        x = m*temp + 1 

    y = m*temp+mod
    if size<mod:
        mod = size
        size = 0
    
    
    #inserting range and their corresponding data
    m1 = str(x)
    m2= str(y)
    df.at[m+2,'Octant ID'] = m1 +'-'+m2 
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
    
    
#defining a dictionary of name of octants
dict = {'1':'Internal outward interaction','-1':'External outward interaction','2':'External Ejection','-2':'Internal Ejection','3':'External inward interaction','-3':'Internal inward interaction','4':'Internal sweep','-4':'External sweep'}

#making a list of counts in ascending order
list = []
for i in range(-4,0):
    list.append(df.at[0,str(i)])
for i in range(1,5):
    list.append(df.at[0,str(i)])
list.sort()

#filling the rank of counts
k = 8
for i in range(len(list)):
    df.at[0,'rank'+(df == list[i]).idxmax(axis=1)[0]] = k
    k = k-1
#filling the octant with highest count
df.at[0,'Rank1 Octant ID'] = (df == list[7]).idxmax(axis=1)[0]
df.at[1,'Rank1 Octant ID'] = 0
#filling the name of octant
df.at[0,'Rank1 Octant Name'] = dict[(df == list[7]).idxmax(axis=1)[0]]


#making and filling list of counts in ascending order
for j in range(2,int(len(df)/mod)+2):
      
    list = []
    for i in range(-4,0):
        list.append(df.at[j,str(i)])
    for i in range(1,5):
        list.append(df.at[j,str(i)])
    list.sort()
    print(list)
    k = 8
    for i in range(len(list)):
        df.at[j,'rank'+(df == list[i]).idxmax(axis=1)[j]] = k
        k = k-1
    #filling the octant with highest count
    df.at[j,'Rank1 Octant ID'] = (df == list[7]).idxmax(axis=1)[j]
    
    #filling the name of octant
    df.at[j,'Rank1 Octant Name'] = dict[(df == list[7]).idxmax(axis=1)[j]]



#df.at[6+int(len(df)/mod),'1'] = 'Octant ID'
k = 1
for i in range(-4,0):
        df.at[k+6+int(len(df)/mod),'1'] = i
        k = k+1
for i in range(1,5):
        df.at[k+6+int(len(df)/mod),'1'] = i
        k = k+1
df.to_excel("Octant_output.xlsx")
print(df)
    
# df.at[6+int(len(df)/mod),'2'] = 'Octant Name'
# k= 1
# for i in range(-4,0):
#         df.at[k+6+int(len(df)/mod),'2'] = dict[str(i)]
#         k = k+1
# for i in range(1,5):
#         df.at[k+6+int(len(df)/mod),'2'] = dict[str(i)]
#         k = k+1
#df.at[6+int(len(df)/mod),'3'] = 'Count of rank1 mod values'

k= 1
for i in range(-4,0):
        df.at[k+6+int(len(df)/mod),'3'] = df['Rank1 Octant ID'].value_counts()[i]
        k = k+1
for i in range(1,5):
        df.at[k+6+int(len(df)/mod),'3'] = df['Rank1 Octant ID'].value_counts()[i]
        k = k+1


print(df)
