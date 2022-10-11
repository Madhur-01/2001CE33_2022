#MADHUR GARG; 2001CE33

#importing pandas
import pandas as pd
import os
#reading the input file
df = pd.read_excel("input_octant_longest_subsequence_with_range.xlsx")

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
df['Octant']         =   df.apply([lambda x : octant(x["U'"],x["V'"],x["W'"])], axis=1)

#making columns for subsequence
df['Octant_No'] = ''
df['Longest_Subsequence_Length'] = ''
df['Count'] = ''

l1 = [1,-1,2,-2,3,-3,4,-4] #making a list of all the octants
l = df['Octant'].tolist()
i=0
for x in l1: #finding subsequence for every octant
    df.at[i,'Octant_No'] = x
    count = 1
    temp = 1
    mx = 0
    for y in range(len(l)-1):
        if x == l[y] and x == l[y+1]:
            temp += 1
        else:
            if mx == temp:
                count += 1
            elif mx < temp:
                count = 1
            mx = max(mx,temp)
            temp = 1
    df.at[i,'Longest_Subsequence_Length'] = mx
    df.at[i,'Count'] = count
    i += 1

i=0
j=0
df[''] = ''
df['Octant_No.'] = ''
df['Longest_Subsequence_Length_'] = ''
df['count'] = ''
for x in l1: #finding subsequence for every octant
    t1 = df.at[i,'Longest_Subsequence_Length']
    df.at[j,'Octant_No.'] = df.at[i,'Octant_No']
    df.at[j,'Longest_Subsequence_Length_'] = df.at[i,'Longest_Subsequence_Length']
    df.at[j,'count'] = df.at[i,'Count']
    j += 1
    df.at[j,'Octant_No.'] = 'Time'
    df.at[j,'Longest_Subsequence_Length_'] = 'From'
    df.at[j,'count'] = 'To'
    j += 1
    temp = 1
    for y in range(len(l)-1):
        if x == l[y] and x == l[y+1]:
            temp += 1
        elif temp == df.at[i,'Longest_Subsequence_Length']:
            df.at[j,'Longest_Subsequence_Length_'] = df.at[y-temp+1,'Time']
            df.at[j,'count'] = df.at[y,'Time']
            j += 1
            temp = 1
        else:
            temp = 1
    i += 1

try: 
    df.to_excel('output_octant_longest_subsequence_with_range')
except:
    print("An exception occurred")