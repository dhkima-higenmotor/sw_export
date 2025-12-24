import os
import sys
import win32com.client
import pythoncom
import time
import tkinter
from tkinter import filedialog, font, ttk

# SolidWorks Constants
swUserPreferenceToggle_e_swDxfIssuingWarning = 11 # Toggle to issue warning if DXF scale is not 1:1

# Parameters
WorkingDirectory = r"D:\github"
Step = "STEP_ON"
Step_Asm = "STEP_ASM_ON"
Dxf = "DXF_ON"
Pdf = "PDF_ON"
Prefix = ""
Out_Dir = "2D"

##############################
# Functions
def init_parameters():
    entry_wd.insert(0,r"D:\github")
    entry_prefix.insert(0,"")
    entry_out_dir.insert(0,"2D")

def read_parameters():
    global WorkingDirectory, Prefix, Out_Dir, Step, Step_Asm, Dxf, Pdf
    WorkingDirectory = entry_wd.get()
    Step = switch_STEP_var.get()
    Step_Asm = switch_STEP_ASM_var.get()
    Dxf = switch_DXF_var.get()
    Pdf = switch_PDF_var.get()
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
    global DXF_MAP_FILE, WorkingDirectory, Prefix, Out_Dir, Step, Step_Asm, Dxf, Pdf
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
            try:
                if (Prefix!="" or Prefix!="*") and FILE_LIST_SLDDRW[i].startswith(Prefix):
                    print('from : '+PATH_INPUT+'\\'+FILE_LIST_SLDDRW[i])
                    time.sleep(1)
                    # OpenDoc6(FileName, Type, Options, Configuration, Errors, Warnings)
                    # Type: 3 (swDocDRAWING)
                    # Options: 1 (Silent) | 32 (LoadModel - Force Resolved)
                    # This ensures drawing is not Lightweight, which causes DXF export failure
                    # Must use VT_BYREF | VT_I4 for output arguments in late-bound calls
                    error = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                    warning = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                    Model = swApp.OpenDoc6(PATH_INPUT+'\\'+FILE_LIST_SLDDRW[i], 3, 1 | 32, "", error, warning)
                    
                    if Model is None:
                         print(f"  FAILED to open: {FILE_LIST_SLDDRW[i]} (Error: {error.value})")
                         continue

                    # Explicitly Activate the document to ensure context for SaveAs
                    # ActivateDoc3 also takes an Error byref
                    swApp.ActivateDoc3(FILE_LIST_SLDDRW[i], False, 0, error)

                    time.sleep(5)
                    
                    # Check for needed References if any (optional)
                    
                    # Use SaveAs3 for better compatibility and silent option
                    if Pdf=="PDF_ON":
                        target_pdf = PATH_PDF+'\\'+BASENAME[i]+'.pdf'
                        # SaveAs3 Silent fails to overwrite sometimes, so delete first
                        if os.path.exists(target_pdf):
                            try:
                                os.remove(target_pdf)
                            except:
                                pass
                        
                        # SaveAs3(Name, Version, Options)
                        # Version: 0 (Current), Options: 1 (Silent)
                        Result_PDF = Model.SaveAs3(target_pdf, 0, 1)
                        time.sleep(1)
                        # SaveAs3 returns 0 on success
                        if Result_PDF == 0:
                            print('  to : '+target_pdf)
                        else:
                            print(f'  FAILED to export PDF: {BASENAME[i]} (Error Code: {Result_PDF})')

                    if Dxf=="DXF_ON":
                        target_dxf = PATH_DXF+'\\'+BASENAME[i]+'.DXF'
                        if os.path.exists(target_dxf):
                            try:
                                os.remove(target_dxf)
                            except:
                                pass

                        # Disable DXF Scale Warning
                        # 11 = swDxfIssuingWarning
                        # Disable DXF Scale Warning (11) and Mapping File (143)
                        # 11 = swDxfIssuingWarning, 143 = swDxfMappingFileEnabled
                        swApp.SetUserPreferenceToggle(11, False)
                        swApp.SetUserPreferenceToggle(143, False)
                        
                        # Use SaveAs3 in Non-Silent Mode (Options=0)
                        # This avoids the "Error 1" caused by Silent Mode in some versions
                        # The preference toggle above handles the specific popup we care about
                        Result_DXF = Model.SaveAs3(target_dxf, 0, 0)
                        
                        time.sleep(5)
                        
                        # Restore Warning (Optional)
                        # swApp.SetUserPreferenceToggle(11, True) 
                        # swApp.SetUserPreferenceToggle(143, True) 

                        if Result_DXF == 0:
                            print('  to : '+target_dxf)
                        else:
                            print(f'  FAILED to export DXF: {BASENAME[i]} (Error Code: {Result_DXF})')

                    swApp.CloseAllDocuments(True)
            except Exception as e:
                print(f"  ERROR processing {FILE_LIST_SLDDRW[i]}: {str(e)}")
                try:
                    swApp.CloseAllDocuments(True)
                except:
                    pass
        
        # Restore Preferences globally at end
        swApp.SetUserPreferenceToggle(11, True)
        swApp.SetUserPreferenceToggle(143, True)
        print("----------------")
    #
    # Export stp as configurations from .SLDPRT
    if Step=="STEP_ON":
        print("# Export STP as configurations from .SLDPRT")
        for i in range(len(FILE_LIST_SLDPRT)):
            try:
                if (Prefix!="" or Prefix!="*") and FILE_LIST_SLDPRT[i].startswith(Prefix):
                    print('from : '+PATH_INPUT+'\\'+FILE_LIST_SLDPRT[i])
                    time.sleep(1)
                    Model = swApp.OpenDoc(PATH_INPUT+'\\'+FILE_LIST_SLDPRT[i],1)
                    ## Get Configurations
                    ConfNames = Model.GetConfigurationNames
                    #print(f' Configurations : {ConfNames}')
                    k = 0
                    if ConfNames is not None:
                        for k in range(len(ConfNames)):
                            try:
                                if (ConfNames[k]=="기본") or (ConfNames[k]=="Default"):
                                    SaveName = BASENAME_STP[i]
                                else:
                                    SaveName = ConfNames[k]
                                print('  to : '+PATH_STP+'\\'+SaveName+'.STEP')
                                Model.ShowConfiguration2(ConfNames[k])
                                # Use SaveAs3 with Silent option (1)
                                Result_STP = Model.SaveAs3(PATH_STP+'\\'+SaveName+'.STEP', 0, 1)
                                time.sleep(1)
                                if Result_STP != 0:
                                     print(f"  FAILED to export STEP: {SaveName} (Error Code: {Result_STP})")
                            except Exception as conf_e:
                                print(f"  ERROR processing config {ConfNames[k]} in {FILE_LIST_SLDPRT[i]}: {str(conf_e)}")

                    swApp.CloseAllDocuments(True)
            except Exception as e:
                print(f"  ERROR processing {FILE_LIST_SLDPRT[i]}: {str(e)}")
                try:
                    swApp.CloseAllDocuments(True)
                except:
                    pass
        print("----------------")
    #
    # Export stp as configurations from .SLDASM
    if Step_Asm=="STEP_ASM_ON":
        print("# Export STP as configurations from .SLDASM")
        for i in range(len(FILE_LIST_SLDASM)):
            try:
                if (Prefix!="" or Prefix!="*") and FILE_LIST_SLDASM[i].startswith(Prefix):
                    print('from : '+PATH_INPUT+'\\'+FILE_LIST_SLDASM[i])
                    time.sleep(1)
                    Model = swApp.OpenDoc(PATH_INPUT+'\\'+FILE_LIST_SLDASM[i],2)
                    ## Get Configurations
                    ConfNames = Model.GetConfigurationNames
                    #print(f' Configurations : {ConfNames}')
                    k = 0
                    if ConfNames is not None:
                        for k in range(len(ConfNames)):
                            try:
                                if (ConfNames[k]=="기본") or (ConfNames[k]=="Default"):
                                    SaveName = BASENAME_STP_ASM[i]
                                else:
                                    SaveName = ConfNames[k]
                                print('  to : '+PATH_STP_ASM+'\\'+SaveName+'.STEP')
                                print('  to : '+PATH_STP_ASM+'\\'+SaveName+'.STEP')
                                Model.ShowConfiguration2(ConfNames[k])
                                # Use SaveAs3 with Silent option (1)
                                Result_STP_ASM = Model.SaveAs3(PATH_STP_ASM+'\\'+SaveName+'.STEP', 0, 1)
                                if Result_STP_ASM != 0:
                                     print(f"  FAILED to export STEP ASM: {SaveName} (Error Code: {Result_STP_ASM})")
                            except Exception as conf_e:
                                print(f"  ERROR processing config {ConfNames[k]} in {FILE_LIST_SLDASM[i]}: {str(conf_e)}")
                    swApp.CloseAllDocuments(True)
            except Exception as e:
                print(f"  ERROR processing {FILE_LIST_SLDASM[i]}: {str(e)}")
                try:
                    swApp.CloseAllDocuments(True)
                except:
                    pass
        print("----------------")
    #
    # Quit Solidworks
    #swApp.ExitApp()
    print("..... Finished!")

# Callback Functions
def button_wd_callback():
    print("# button_wd pressed.")
    global WorkingDirectory
    WorkingDirectory = filedialog.askdirectory(title="Select Solidworks Working Directory", initialdir=WorkingDirectory)
    entry_wd.delete(0,'end')
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
app = tkinter.Tk()
app.title("sw_export")
app.resizable(width=False, height=False)
font_big = font.Font(size=14)
if ( sys.platform.startswith('win')): app.iconbitmap('icons\\sw_export.ico')
# Gap between pads in tkinter
PADX = 5
PADY = 1
GRIDWIDTH = 150

##############################
# GUI items
# Working Directory
label_wd = ttk.Label(app, text="DIR = ")
label_wd.grid(row=0, column=0, padx=PADX, pady=PADY, sticky="e")

entry_wd = ttk.Entry(app, width=40)
entry_wd.grid(row=0, column=1, padx=PADX, pady=PADY, columnspan=2)

button_wd = ttk.Button(app, text="BROWSE", command=button_wd_callback, width=10)
button_wd.grid(row=0, column=3, padx=PADX, pady=PADY, sticky="w")

# Subject 1
label_subject1 = ttk.Label(app, text="# OPTIONS")
label_subject1.grid(row=1, column=0, padx=PADX, pady=PADY, sticky="w")

# Switch for STEP
switch_STEP_var = tkinter.StringVar(value="STEP_ON")
switch_STEP = ttk.Checkbutton(app, text="STEP", command=switch_STEP, variable=switch_STEP_var, onvalue="STEP_ON", offvalue="STEP_OFF")
switch_STEP.grid(row=2, column=0, padx=PADX, pady=PADY, sticky="w")
switch_STEP.invoke() # Set default to ON

# Switch for STEP_ASM
switch_STEP_ASM_var = tkinter.StringVar(value="STEP_ASM_ON")
switch_STEP_ASM = ttk.Checkbutton(app, text="STEP_ASM", command=switch_STEP_ASM, variable=switch_STEP_ASM_var, onvalue="STEP_ASM_ON", offvalue="STEP_ASM_OFF")
switch_STEP_ASM.grid(row=3, column=0, padx=PADX, pady=PADY, sticky="w")
switch_STEP_ASM.invoke()

# Switch for DXF
switch_DXF_var = tkinter.StringVar(value="DXF_ON")
switch_DXF = ttk.Checkbutton(app, text="DXF", command=switch_DXF, variable=switch_DXF_var, onvalue="DXF_ON", offvalue="DXF_OFF")
switch_DXF.grid(row=4, column=0, padx=PADX, pady=PADY, sticky="w")
switch_DXF.invoke()

# Switch for PDF
switch_PDF_var = tkinter.StringVar(value="PDF_ON")
switch_PDF = ttk.Checkbutton(app, text="PDF", command=switch_PDF, variable=switch_PDF_var, onvalue="PDF_ON", offvalue="PDF_OFF")
switch_PDF.grid(row=5, column=0, padx=PADX, pady=PADY, sticky="w")
switch_PDF.invoke()

# PREFIX
entry_prefix = ttk.Entry(app, width=40)
entry_prefix.grid(row=2, column=1, padx=PADX, pady=PADY, sticky="e", columnspan=2)

label_prefix = ttk.Label(app, text="PREFIX")
label_prefix.grid(row=2, column=3, padx=PADX, pady=PADY, sticky="w")

# OUT_DIR
entry_out_dir = ttk.Entry(app, width=40)
entry_out_dir.grid(row=3, column=1, padx=PADX, pady=PADY, sticky="e", columnspan=2)

label_out_dir = ttk.Label(app, text="OUT_DIR")
label_out_dir.grid(row=3, column=3, padx=PADX, pady=PADY, sticky="w")

# Run Button
button_run = ttk.Button(app, text="RUN", command=button_run_callback, width=30)
button_run.grid(row=5, column=1, padx=PADX, pady=PADY, sticky="w", columnspan=2)

# Exit Button
button_exit = ttk.Button(app, text="EXIT", command=button_exit_callback, width=10)
button_exit.grid(row=5, column=3, padx=PADX, pady=PADY, sticky="w")


##############################
# Init
init_parameters()

##############################
# GUI loop
app.mainloop()
