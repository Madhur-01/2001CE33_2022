#MADHUR GARG; 2001CE33

#importing pandas
import pandas as pd

#reading the input file
df = pd.read_excel("input_octant_longest_subsequence.xlsx")

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

# #Finding total number of rows.
total_rows=len(df.axes[0])

#Defining a function to find count of longest subsequence.
def longest_subsequence_count():
    octant_val=[1,-1,2,-2,3,-3,4,-4] #Numbers assigning octants.
    subsequence_len=[0,0,0,0,0,0,0,0] 
    max_subsequence_len=[-1,-1,-1,-1,-1,-1,-1,-1]
    temp_subsequence_len=[0,0,0,0,0,0,0,0] 
    max_subsequence_count=[0,0,0,0,0,0,0,0] 
    for t in range(total_rows-1): #Applying logic here.
        for u in range(8):
            if(df.at[t,'Octant']==octant_val[u]):
                if(df.at[t+1,'Octant']==octant_val[u]):
                    subsequence_len[u]+=1
                else:
                    max_subsequence_len[u]=max(max_subsequence_len[u],subsequence_len[u])
                    subsequence_len[u]=0
                break
    for t in range(total_rows-1):
        for u in range(8):
            if(df.at[t,'Octant']==octant_val[u]):
                if(df.at[t+ 1,'Octant']==octant_val[u]):
                    temp_subsequence_len[u]+=1
                else:
                    if(temp_subsequence_len[u]==max_subsequence_len[u]):
                        max_subsequence_count[u]+=1
                    temp_subsequence_len[u]=0
                break
    
    for i in range(8):
        df.loc[df.index[i],'Octant Num.']=octant_val[i]
        df.loc[df.index[i],'Longest Subsequence Length']=max_subsequence_len[i]+1
        df.loc[df.index[i],'Count']=max_subsequence_count[i]
    
    #Now storing this dataframe to an excel file.
    df.to_excel('output_octant_longest_subsequence.xlsx')

#Calling the main function.
longest_subsequence_count()