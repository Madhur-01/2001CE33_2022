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

# Method to initialise dictionary with 0 for "OcTaNt_SiGn" except 'left'

def ReSeT_cOuNt_except(count, left):
    for item in OcTaNt_SiGn:
        if(item!=left):
            count[item] = 0


def SeT_FrEqUeNcY(longest, frequency, outputSheet):
    # Iterating "OcTaNt_SiGn" and updating sheet
    for i in range(9):
        for j in range(3):
            outputSheet.cell(row = 3+i, column = 45+j).border = BlAcK_BoRdEr

    outputSheet.cell(row=3, column=45).value= "Octant ##"
    outputSheet.cell(row=3, column=46).value= "Longest Subsquence Length"
    outputSheet.cell(row=3, column=47).value= "Count"

    for i, label in enumerate(OcTaNt_SiGn):
        currRow = i+3
        try:
            outputSheet.cell(row=currRow+1, column=45).value = label	
            outputSheet.cell(column=46, row=currRow+1).value = longest[label]
            outputSheet.cell(column=47, row=currRow+1).value = frequency[label]
        except FileNotFoundError:
            print("File not found!!")
            exit()


# Method to set time range for longest subsequence
def LoNgEsT_sUbSeQuEnCe_TiMe(longest, frequency, timeRange, outputSheet):
    # Naming columns number
    lengthCol = 50
    freqCol = 51
    
    # Initial row, just after the header row
    row = 4

    outputSheet.cell(row=3, column = 49).value = "Octant ###"

    outputSheet.cell(row=3, column = 50).value = "Longest Subsquence Length"

    outputSheet.cell(row=3, column = 51).value = "Count"


    # Iterating all octants 
    for octant in OcTaNt_SiGn:
        try:
            # Setting octant's longest subsequence and frequency data
            outputSheet.cell(column=49, row=row).value = octant

            outputSheet.cell(column=lengthCol, row=row).value = longest[octant]
            
            outputSheet.cell(column=freqCol, row=row).value = frequency[octant]

        except FileNotFoundError:
            
            print("File not found!!")
            exit()

        row+=1

        try:
            # Setting default labels
            outputSheet.cell(column=49, row=row).value = "Time"

            outputSheet.cell(column=lengthCol, row=row).value = "From"

            outputSheet.cell(column=freqCol, row=row).value = "To"

        except FileNotFoundError:

            print("File not found!!")

            exit()

        row+=1