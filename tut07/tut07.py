# MADHUR GARG  2001CE33

from datetime import datetime
start_time = datetime.now()
import os
os.system("cls")
import openpyxl
import glob

from openpyxl.styles import Color, PatternFill, Font, Border, Side

from platform import python_version
ver = python_version()
if ver == "3.8.10":
	print("Correct Version Installed")
else:
	print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

##Read all the excel files in a batch format from the input/ folder. Only xlsx to be allowed

##Save all the excel files in a the output/ folder. Only xlsx to be allowed

## output filename = input_filename[_OcTaNt_AnAlySis_mod_5000].xlsx , ie, append _OcTaNt_AnAlySis_mod_5000 to the original filename.

OcTaNt_SiGn             = [1,-1,2,-2,3,-3,4,-4]
OcTaNt_NaMe_Id_MaPpInG  = {1:"Internal outward interaction", -1:"External outward interaction", 2:"External Ejection", -2:"Internal Ejection", 3:"External inward interaction", -3:"Internal inward interaction", 4:"Internal sweep", -4:"External sweep"}
YeLLOw                  = "00FFFF00"
YeLLOw_bg               = PatternFill(start_color=YeLLOw, end_color= YeLLOw, fill_type='solid')
black                   = "00000000"
DoUbLe                  = Side(border_style="thin", color=black)
BlAcK_BoRdEr            = Border(top=DoUbLe, left=DoUbLe, right=DoUbLe, bottom=DoUbLe)

#Code
def ReSeT_cOuNt(count):
    for item in OcTaNt_SiGn:
        count[item] = 0
