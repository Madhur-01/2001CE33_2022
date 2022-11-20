#MADHUR GARG 2001CE33
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Alignment,Border,Side
import numpy as np
import os
import re
from datetime import datetime
os.system('cls')
start_time = datetime.now()

#making function for one innings
def one_inning(inn1,bat_pl,bow_pl,s): 
    innbat = Workbook()
    innfow = Workbook()
    innbow = Workbook()
    s1 = innbat.active
    s2 = innbow.active
    s3 = innfow.active
    s1.column_dimensions['A'].width = 25


	
    #scorecard and index
    s1['A1'] = s + ' Innings'
    s1['I1'] = '0-0'
    s1['J1'] = '0 overs'
    s1['A2'] = 'Batter'
    s1['F2'] = 'R'
    s1['G2'] = 'B'
    s1['H2'] = '4s'
    s1['I2'] = '6s'
    s1['J2'] = 'SR'

    s2['A1'] = 'Bowler'
    s2['D1'] = 'O'
    s2['E1'] = 'M'
    s2['F1'] = 'R'
    s2['G1'] = 'W'
    s2['H1'] = 'NB'
    s2['I1'] = 'WD'
    s2['J1'] = 'ECO'

    s3['A1'] = 'Fall of wickets'
