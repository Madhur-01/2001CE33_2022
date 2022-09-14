#MADHUR GARG, 2001CE33

#importing spandas
import pandas as pd

#reading the input file
df = pd.read_csv("octant_input.csv")

#data preprocessing
df.at[0,'U_avg']  = df['U'].mean()
df.at[0,'V_avg']  = df['V'].mean()
df.at[0,'W_avg']  = df['W'].mean()


df["U'"] = df['U'] - df.at[0,'U_avg']
df["V'"] = df['V'] - df.at[0,'V_avg']
df["W'"] = df['W'] - df.at[0,'W_avg']


