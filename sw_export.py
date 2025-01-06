import os
import sys
import win32com.client
import time
import customtkinter
from customtkinter import filedialog

# Parameters
WorkingDirectory = "D:\\github"
Step = "STEP_ON"
Step_Asm = "STEP_ASM_ON"
Dxf = "DXF_ON"
Pdf = "PDF_ON"
Prefix = "*"
Out_Dir = "2D"

##############################
# Functions
def init_parameters():
    entry_wd.insert(0,"D:\\github")
    entry_prefix.insert(0,"*")
    entry_out_dir.insert(0,"2D")

def read_parameters():
    global WorkingDirectory, Prefix, Out_Dir, Step, Step_Asm, Dxf, Pdf
    WorkingDirectory = entry_wd.get()
    Step = switch_STEP.get()
    Step_Asm = switch_STEP_ASM.get()
    Dxf = switch_DXF.get()
    Pdf = switch_PDF.get()
    Prefix = entry_prefix.get()
    Out_Dir = entry_out_dir.get()
    if Out_Dir=="":
        Out_Dir = "2D"
    print("## Input Parameters")
    print("WorkingDirectory = " + WorkingDirectory)
    print("Step = " + Step)
    print("Step_Asm = " + Step_Asm)
    print("Dxf = " + Dxf)
    print("Pdf = " + Pdf)
    print("Prefix = " + Prefix)
    print("Out_Dir = " + Out_Dir)
    print("-------------------")

def run_export():
    global WorkingDirectory, Prefix, Out_Dir, Step, Step_Asm, Dxf, Pdf
    global Result_PDF, Result_DXF, Result_STP, Result_STP_ASM
    #
    # Working Directory
    PATH_INPUT = WorkingDirectory
    if PATH_INPUT == '':
        sys.exit()
    #
    # File List
    FILE_LIST = os.listdir(PATH_INPUT)
    FILE_LIST_SLDDRW = [file for file in FILE_LIST if (file[0:2]!="~$") and (file.endswith(".slddrw") or file.endswith(".SLDDRW")) ]
    FILE_LIST_SLDPRT = [file for file in FILE_LIST if (not "_SKEL." in file) and (file[0:2]!="~$") and (file.endswith(".sldprt") or file.endswith(".SLDPRT")) ]
    FILE_LIST_SLDASM = [file for file in FILE_LIST if (file[0:1]!="~") and (file.endswith(".sldasm") or file.endswith(".SLDASM")) ]
    #
    # Make Directories
    PATH_2D = PATH_INPUT + "\\" + Out_Dir
    PATH_DXF = PATH_2D + "\\DXF"
    PATH_PDF = PATH_2D + "\\PDF"
    PATH_STP = PATH_2D + "\\STEP"
    PATH_STP_ASM = PATH_2D + "\\STEP_ASM"
    if os.path.exists(PATH_2D)==False:
        os.makedirs(PATH_2D)
    if os.path.exists(PATH_DXF)==False and Dxf=="DXF_ON":
        os.makedirs(PATH_DXF)
    if os.path.exists(PATH_PDF)==False and Pdf=="PDF_ON":
        os.makedirs(PATH_PDF)
    if os.path.exists(PATH_STP)==False and Step=="STEP_ON":
        os.makedirs(PATH_STP)
    if os.path.exists(PATH_STP_ASM)==False and Step_Asm=="STEP_ASM_ON":
        os.makedirs(PATH_STP_ASM)
    #
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
    #
    # Start Solidworks
    swApp = win32com.client.Dispatch('SldWorks.Application')
    swApp.Visible = True
    time.sleep(10)
    #
    # Export PDF, DXF
    if Dxf=="DXF_ON" or Pdf=="PDF_ON":
        print("# Export PDF, DXF")
        for i in range(len(FILE_LIST_SLDDRW)):
            if (Prefix!="" or Prefix!="*") and FILE_LIST_SLDDRW[i].startswith(Prefix):
                print('from : '+PATH_INPUT+'\\'+FILE_LIST_SLDDRW[i])
                time.sleep(1)
                Model = swApp.OpenDoc(PATH_INPUT+'\\'+FILE_LIST_SLDDRW[i],3)
                time.sleep(5)
                if Pdf=="PDF_ON":
                    Result_PDF = Model.SaveAs(PATH_PDF+'\\'+BASENAME[i]+'.pdf')
                    print('  to : '+PATH_INPUT+'\\'+BASENAME[i]+'.pdf')
                if Dxf=="DXF_ON":
                    Result_DXF = Model.SaveAs(PATH_DXF+'\\'+BASENAME[i]+'.DXF')
                    print('  to : '+PATH_INPUT+'\\'+BASENAME[i]+'.DXF')
                swApp.CloseAllDocuments(True)
        print("----------------")
    #
    # Export stp as configurations from .SLDPRT
    if Step=="STEP_ON":
        print("# Export STP as configurations from .SLDPRT")
        for i in range(len(FILE_LIST_SLDPRT)):
            if (Prefix!="" or Prefix!="*") and FILE_LIST_SLDPRT[i].startswith(Prefix):
                print('from : '+PATH_INPUT+'\\'+FILE_LIST_SLDPRT[i])
                time.sleep(1)
                Model = swApp.OpenDoc(PATH_INPUT+'\\'+FILE_LIST_SLDPRT[i],1)
                ## Get Configurations
                ConfNames = Model.GetConfigurationNames
                #print(f' Configurations : {ConfNames}')
                k = 0
                for k in range(len(ConfNames)):
                    if (ConfNames[k]=="기본") or (ConfNames[k]=="Default"):
                        SaveName = BASENAME_STP[i]
                    else:
                        SaveName = ConfNames[k]
                    print('  to : '+PATH_STP+'\\'+SaveName+'.STEP')
                    Model.ShowConfiguration2(ConfNames[k])
                    Result_STP = Model.SaveAs(PATH_STP+'\\'+SaveName+'.STEP')
                swApp.CloseAllDocuments(True)
        print("----------------")
    #
    # Export stp as configurations from .SLDASM
    if Step_Asm=="STEP_ASM_ON":
        print("# Export STP as configurations from .SLDASM")
        for i in range(len(FILE_LIST_SLDASM)):
            if (Prefix!="" or Prefix!="*") and FILE_LIST_SLDASM[i].startswith(Prefix):
                print('from : '+PATH_INPUT+'\\'+FILE_LIST_SLDASM[i])
                time.sleep(1)
                Model = swApp.OpenDoc(PATH_INPUT+'\\'+FILE_LIST_SLDASM[i],2)
                ## Get Configurations
                ConfNames = Model.GetConfigurationNames
                #print(f' Configurations : {ConfNames}')
                k = 0
                for k in range(len(ConfNames)):
                    if (ConfNames[k]=="기본") or (ConfNames[k]=="Default"):
                        SaveName = BASENAME_STP_ASM[i]
                    else:
                        SaveName = ConfNames[k]
                    print('  to : '+PATH_STP_ASM+'\\'+SaveName+'.STEP')
                    Model.ShowConfiguration2(ConfNames[k])
                    Result_STP_ASM = Model.SaveAs(PATH_STP_ASM+'\\'+SaveName+'.STEP')
                swApp.CloseAllDocuments(True)
        print("----------------")
    #
    # Quit Solidworks
    swApp.ExitApp()
    print("..... Finished!")

# Callback Functions
def button_wd_callback():
    print("# button_wd pressed.")
    global WorkingDirectory
    WorkingDirectory = filedialog.askdirectory(title="Select Solidworks Working Directory", initialdir=WorkingDirectory)
    entry_wd.delete(0,last_index='end')
    entry_wd.insert(0,WorkingDirectory)
    print('Working Directory : %s'%WorkingDirectory)

def switch_STEP():
    print("# switch_STEP Toggled.")

def switch_STEP_ASM():
    print("# switch_STEP_ASM Toggled.")

def switch_DXF():
    print("# switch_DXF Toggled.")

def switch_PDF():
    print("# switch_PDF Toggled.")

def button_run_callback():
    print("# button_run pressed.")
    read_parameters()
    run_export()

def button_exit_callback():
    print("# button_exit pressed.")
    exit()

##############################
# GUI config
customtkinter.set_default_color_theme("green")
app = customtkinter.CTk()
app.title("sw_export")
#app.geometry("950x650")
app.resizable(width=False, height=False)
font_big = customtkinter.CTkFont(size=14)
if ( sys.platform.startswith('win')): app.iconbitmap('icons\\sw_export.ico')
# Gap between pads in customtkinter
PADX = 5
PADY = 1
GRIDWIDTH = 150

##############################
# GUI items
# Working Directory
label_wd = customtkinter.CTkLabel(app, text="DIR = ", fg_color="transparent", compound="right")
label_wd.grid(row=0, column=0, padx=PADX, pady=PADY, sticky="e")

entry_wd = customtkinter.CTkEntry(app, placeholder_text="D:\\github", width=GRIDWIDTH*2)
entry_wd.grid(row=0, column=1, padx=PADX, pady=PADY, columnspan=2)

button_wd = customtkinter.CTkButton(app, text="BROWSE", command=button_wd_callback, width=GRIDWIDTH/2)
button_wd.grid(row=0, column=3, padx=PADX, pady=PADY, sticky="w")

# Subject 1
label_subject1 = customtkinter.CTkLabel(app, text="# OPTIONS", fg_color="transparent", compound="right", font=font_big)
label_subject1.grid(row=1, column=0, padx=PADX, pady=PADY, sticky="w")

# Switch for STEP
switch_STEP_var = customtkinter.StringVar(value="STEP_ON")
switch_STEP = customtkinter.CTkSwitch(app, text="STEP", command=switch_STEP, variable=switch_STEP_var, onvalue="STEP_ON", offvalue="STEP_OFF")
switch_STEP.grid(row=2, column=0, padx=PADX, pady=PADY, sticky="w")

# Switch for STEP_ASM
switch_STEP_ASM_var = customtkinter.StringVar(value="STEP_ASM_ON")
switch_STEP_ASM = customtkinter.CTkSwitch(app, text="STEP_ASM", command=switch_STEP_ASM, variable=switch_STEP_ASM_var, onvalue="STEP_ASM_ON", offvalue="STEP_ASM_OFF")
switch_STEP_ASM.grid(row=3, column=0, padx=PADX, pady=PADY, sticky="w")

# Switch for DXF
switch_DXF_var = customtkinter.StringVar(value="DXF_ON")
switch_DXF = customtkinter.CTkSwitch(app, text="DXF", command=switch_DXF, variable=switch_DXF_var, onvalue="DXF_ON", offvalue="DXF_OFF")
switch_DXF.grid(row=4, column=0, padx=PADX, pady=PADY, sticky="w")

# Switch for PDF
switch_PDF_var = customtkinter.StringVar(value="PDF_ON")
switch_PDF = customtkinter.CTkSwitch(app, text="PDF", command=switch_PDF, variable=switch_PDF_var, onvalue="PDF_ON", offvalue="PDF_OFF")
switch_PDF.grid(row=5, column=0, padx=PADX, pady=PADY, sticky="w")

# PREFIX
entry_prefix = customtkinter.CTkEntry(app, placeholder_text="*", width=GRIDWIDTH*2)
entry_prefix.grid(row=2, column=1, padx=PADX, pady=PADY, sticky="e", columnspan=2)

label_prefix = customtkinter.CTkLabel(app, text="PREFIX", fg_color="transparent", compound="left")
label_prefix.grid(row=2, column=3, padx=PADX, pady=PADY, sticky="w")

# OUT_DIR
entry_out_dir = customtkinter.CTkEntry(app, placeholder_text="2D", width=GRIDWIDTH*2)
entry_out_dir.grid(row=3, column=1, padx=PADX, pady=PADY, sticky="e", columnspan=2)

label_out_dir = customtkinter.CTkLabel(app, text="OUT_DIR", fg_color="transparent", compound="left")
label_out_dir.grid(row=3, column=3, padx=PADX, pady=PADY, sticky="w")

# Run Button
button_run = customtkinter.CTkButton(app, text="RUN", command=button_run_callback, width=GRIDWIDTH*2)
button_run.grid(row=5, column=1, padx=PADX, pady=PADY, sticky="w", columnspan=2)

# Exit Button
button_exit = customtkinter.CTkButton(app, text="EXIT", command=button_exit_callback, width=GRIDWIDTH/2)
button_exit.grid(row=5, column=3, padx=PADX, pady=PADY, sticky="w")


##############################
# Init
init_parameters()

##############################
# GUI loop
app.mainloop()
