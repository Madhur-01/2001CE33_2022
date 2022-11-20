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

def OnE_iNnInG(inn1,bat_pl,bow_pl,s) : #making function for one innings
    innbat = Workbook()
    innfow = Workbook()
    innbow = Workbook()
    S1 = innbat.active
    S2 = innbow.active
    S3 = innfow.active
    S1.column_dimensions['A'].width = 25


    #scorecard and index
    S1['A1']    = s + ' Innings'
    S1['I1']     = '0-0'
    S1['J1']    = '0 overs'
    S1['A2']    = 'Batter'
    S1['F2']     = 'R'
    S1['G2']     = 'B'
    S1['H2']    = '4s'
    S1['I2']    = '6s'
    S1['J2']     = 'SR'


    S2['A1']    = 'Bowler'
    S2['D1'] =  'O'
    S2['E1']    = 'M'
    S2['F1']    = 'R'
    S2['G1'] =  'W'
    S2['H1']    = 'NB'
    S2['I1'] =  'WD'
    S2['J1']     = 'ECO'

    
    S3['A1']    = 'Fall of wickets'
    

    #using regex
    Over   =   re.compile(r'(\d\d?\.\d)')
    Zero      =   re.compile(r'no run')
    No_ball   =   re.compile(r', (no ball),')
    Wide    =     re.compile(r', wide,')
    Wide2     =   re.compile(r', 2 wides,')
    wide3     =   re.compile(r', 3 wides,')
    single    =   re.compile(r', 1 run,')
    SIX       =   re.compile(r', SIX,')
    FOUR      =   re.compile(r', FOUR,')
    BYES      =   re.compile(r', (byes),')
    lBYES     =   re.compile(r', (leg byes),')
    Double    =   re.compile(r', 2 runs,')
    Triple    =   re.compile(r', 3 runs,')
    out       =   re.compile(r', out')
    player    =   re.compile(r'(\d\d?\.\d) (\w+) (to|\w+ to) (\w+)( \w+)?,')
    caught    =   re.compile(r', out Caught by (\w+)')
    lbw       =   re.compile(r', out Lbw!!')
    BOWLED    =   re.compile(r', out bowled!!')
    run_out   =   re.compile(r'Run Out!! ')
    runs      =   0
    wickets   =   0
    nb        =   0
    nlb       =   0
    nw        =   0
    nnb       =   0
    ppr       =   0

