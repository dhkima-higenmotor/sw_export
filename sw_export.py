import os
import sys
import win32com.client
import time
import PySimpleGUI as sg

# Working Directory
PATH_INPUT = sg.popup_get_folder("Select working directory")
if PATH_INPUT is None:
    sys.exit()
FILE_LIST = os.listdir(PATH_INPUT)
FILE_LIST_SLDDRW = [file for file in FILE_LIST if (file[0:2]!="~$") and (file.endswith(".slddrw") or file.endswith(".SLDDRW")) ]
FILE_LIST_SLDPRT = [file for file in FILE_LIST if (not "_SKEL." in file) and (file[0:2]!="~$") and (file.endswith(".sldprt") or file.endswith(".SLDPRT")) ]
FILE_LIST_SLDASM = [file for file in FILE_LIST if (file[0:1]!="~") and (file.endswith(".sldasm") or file.endswith(".SLDASM")) ]

# Make Directories
PATH_2D = PATH_INPUT + "\\2D"
PATH_DXF = PATH_2D + "\\DXF"
PATH_PDF = PATH_2D + "\\PDF"
PATH_STP = PATH_2D + "\\STEP"
PATH_STP_ASM = PATH_2D + "\\STEP_ASM"
if os.path.exists(PATH_2D) == False:
    os.makedirs(PATH_2D)
if os.path.exists(PATH_DXF) == False:
    os.makedirs(PATH_DXF)
if os.path.exists(PATH_PDF) == False:
    os.makedirs(PATH_PDF)
if os.path.exists(PATH_STP) == False:
    os.makedirs(PATH_STP)
if os.path.exists(PATH_STP_ASM) == False:
    os.makedirs(PATH_STP_ASM)

# BASENAME
BASENAME = FILE_LIST_SLDDRW.copy()
for i in range(len(FILE_LIST_SLDDRW)):
    BASENAME[i] = os.path.splitext(FILE_LIST_SLDDRW[i])[0]
BASENAME_STP = FILE_LIST_SLDPRT.copy()
for i in range(len(FILE_LIST_SLDPRT)):
    BASENAME_STP[i] = os.path.splitext(FILE_LIST_SLDPRT[i])[0]
BASENAME_STP_ASM = FILE_LIST_SLDASM.copy()
for i in range(len(FILE_LIST_SLDASM)):
    BASENAME_STP_ASM[i] = os.path.splitext(FILE_LIST_SLDASM[i])[0]

# Start Solidworks
swApp = win32com.client.Dispatch('SldWorks.Application')
swApp.Visible = True
time.sleep(10)

# Export PDF, DXF
print("1. Export PDF,DXF from")
for i in range(len(FILE_LIST_SLDDRW)):
    print('from : '+PATH_INPUT+'\\'+FILE_LIST_SLDDRW[i])
    Model = swApp.OpenDoc(PATH_INPUT+'\\'+FILE_LIST_SLDDRW[i],3)
    Result_PDF = Model.SaveAs(PATH_PDF+'\\'+BASENAME[i]+'.pdf')
    print('  to : '+PATH_INPUT+'\\'+BASENAME[i]+'.pdf')
    Result_DXF = Model.SaveAs(PATH_DXF+'\\'+BASENAME[i]+'.DXF')
    print('  to : '+PATH_INPUT+'\\'+BASENAME[i]+'.DXF')
    swApp.CloseAllDocuments(True)
print("----------------")

# Export stp as configurations from .SLDPRT
print("2. Export STP as configurations from .SLDPRT")
for i in range(len(FILE_LIST_SLDPRT)):
    print('from : '+PATH_INPUT+'\\'+FILE_LIST_SLDPRT[i])
    Model = swApp.OpenDoc(PATH_INPUT+'\\'+FILE_LIST_SLDPRT[i],1)
    ## Get Configurations
    ConfNames = Model.GetConfigurationNames
    #print(f' Configurations : {ConfNames}')
    k = 0
    for k in range(len(ConfNames)):
        if ConfNames[k] == "기본":
            SaveName = BASENAME_STP[i]
        else:
            SaveName = ConfNames[k]
        print('  to : '+PATH_INPUT+'\\'+SaveName+'.STEP')
        Model.ShowConfiguration2(ConfNames[k])
        Result_STP = Model.SaveAs(PATH_STP+'\\'+SaveName+'.STEP')
    swApp.CloseAllDocuments(True)
print("----------------")

# Export stp as configurations from .SLDASM
print("3. Export STP as configurations from .SLDASM")
for i in range(len(FILE_LIST_SLDASM)):
    print('from : '+PATH_INPUT+'\\'+FILE_LIST_SLDASM[i])
    Model = swApp.OpenDoc(PATH_INPUT+'\\'+FILE_LIST_SLDASM[i],2)
    ## Get Configurations
    ConfNames = Model.GetConfigurationNames
    #print(f' Configurations : {ConfNames}')
    k = 0
    for k in range(len(ConfNames)):
        if ConfNames[k] == "기본":
            SaveName = BASENAME_STP[i]
        else:
            SaveName = ConfNames[k]
        print('  to : '+PATH_INPUT+'\\'+SaveName+'.STEP')
        Model.ShowConfiguration2(ConfNames[k])
        Result_STP = Model.SaveAs(PATH_STP_ASM+'\\'+SaveName+'.STEP')
    swApp.CloseAllDocuments(True)
print("----------------")

# Quit Solidworks
swApp.ExitApp()
print("END!")
