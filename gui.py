import subprocess
import tkinter as tk
import tkinter.messagebox
import sys
import os
import shutil
from time import sleep
from shutil import copyfile
from os import path
from tkinter import filedialog
from tkinter import ttk
from glob import glob 

# Notes:
# Adding --onefile to build command breaks selenium
# to build exe run command: pyinstaller --add-data "tabula-1.0.5-jar-with-dependencies.jar;tabula" -w --upx-dir=./upx-3.96-win64 gui.py

# My imports
from classes.Printer import Printer
from classes.DocumentHandler import DocumentHandler
from classes.MaterialsList import MaterialsList
from classes.SAPHandler import SAPHandler
from classes.WebHandler import WebHandler

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Integration I Prep")
        self.master.geometry("600x325")
        self.materialsListPath = ""
        self.selectedPrinter = ""
        self.statusIDLEColor = "red" # "#e8d900"

        self.printer = Printer()
        self.documentHandler = DocumentHandler()
        self.materialsList = MaterialsList()
        self.sapHandler = SAPHandler()
        self.webHandler = WebHandler()

        self.proNumVar = tk.StringVar(value = "Prod #: ")
        self.serialNumVar = tk.StringVar(value = "Serial #: ")
        self.chassisNumVar = tk.StringVar(value = "Chassis #: ")
        self.cellNumVar = tk.StringVar(value = "Cell #: ")
        self.statusLabelTextVar = tk.StringVar(value = "READY")

        self.generatedDocuments = False
        self.checking = False
        self.deselectingCheckEverything = False
        self.proEntered = False
        self.checkButtons = []

        self.create_widgets()

        self.openSAPBatchFilePath = "openSAP.bat"
        

        print(self.openSAPBatchFilePath)

        # self.openSAPBatchFilePath = 'openSAP.bat'
  
    # Function for opening the 
    # file explorer window

    def find_ext(self, dr, ext):
        return glob(path.join(dr,"*.{}".format(ext)))

    def browseFiles(self):
        filePath = filedialog.askopenfilename(initialdir = "/",
                                            title = "Select a File",
                                            filetypes = (("PDF files", "*.pdf"),("all files","*.*"),))
        
        # Change label contents
        filename = filePath.split("/")[-1]
        print(filename)
        self.label_info.configure(text="File Chosen: \n" + filename)
        self.materialsListPath = filePath

    def executeSAP(self):
        proNum = self.serialEntered.get()

        if proNum == "":
            tk.messagebox.showerror("Error", "Please enter pro number.")
        else:
            self.sapHandler.runSAPScript(proNum)

            tk.messagebox.showinfo("Info", "SAP script ready for use.")

            self.webHandler.openSAP()

            subprocess.call(self.openSAPBatchFilePath)

            self.webHandler.closeBrowser()

    def create_widgets(self):
        self.generateDocumentsHeader =  ttk.Label(self.master, 
                                    text = "Print Documents",
                                    font="bold")

        self.masterFrame =  ttk.LabelFrame(self.master)
        self.bottomFrame =  ttk.Frame(self.master)
        self.topFrame =  ttk.Frame(self.master)

        self.statusLabel =  tk.Label(self.topFrame, 
                                    text = "Status: ",
                                    width = 15)
        self.statusLabel.grid(row=0,column=0)
        self.statusLabelText =  tk.Label(self.topFrame, 
                                    textvariable = self.statusLabelTextVar,
                                    width = 15,
                                    fg = "#e8d900")
        self.statusLabelText.grid(row=0,column=1, padx = 0, pady = 0)

        # Frames
        self.fileBrowserFrame =  ttk.Frame(self.masterFrame) # , padx=-50, pady=-300
        self.fileBrowserFrame.grid(row=0,column=0)

        self.statsFrame =  ttk.Frame(self.masterFrame) # , padx=-50
        self.statsFrame.grid(row=0,column=1)

        # Create label and searchbox for user serial input
        self.serialLabel =  tk.Label(self.fileBrowserFrame, 
                                    text = "Enter PRO #:",
                                    width = 20, # height = 4, 
                                    fg = "black")
        self.serialLabel.grid(row=0,column=0)

        self.serial = tk.StringVar()
        self.serialEntered = ttk.Entry(
                                    self.fileBrowserFrame, 
                                    width = 15, 
                                    textvariable = self.serial)
        self.serialEntered.grid(row = 0, column = 1)

        # Create button to run SAP script
        self.runSAPScript = ttk.Button(self.fileBrowserFrame, 
                                text = "Edit SAP Script",
                                command = self.executeSAP) 
        self.runSAPScript.grid(row=0,column=2)

        # Create a File Explorer label
        # self.label_info =  tk.Label(self.fileBrowserFrame, 
        #                             text = "Select materials list PDF.",
        #                             width = 20, # height = 4, 
        #                             fg = "red")
        # self.label_info.grid(row=1,column=0)
            
        # self.button_explore = ttk.Button(self.fileBrowserFrame, 
        #                         text = "Browse Files",
        #                         command = self.browseFiles)
        # self.button_explore.grid(row=1,column=1,columnspan=2,pady=25)

        self.labelWidth = 35

        self.label_proNum =  tk.Label(self.statsFrame, 
                                    textvariable = self.proNumVar,
                                    width = self.labelWidth, 
                                    height = 2)

        self.label_proNum.grid(row=0,column=0)

        self.label_serialNum =  tk.Label(self.statsFrame, 
                                    textvariable = self.serialNumVar,
                                    width = self.labelWidth, 
                                    height = 2)
        self.label_serialNum.grid(row=1,column=0)

        self.label_chassisNum =  tk.Label(self.statsFrame, 
                                    textvariable = self.chassisNumVar,
                                    width = self.labelWidth, 
                                    height = 2)
        self.label_chassisNum.grid(row=2,column=0)

        self.label_cellNum =  tk.Label(self.statsFrame, 
                                    textvariable = self.cellNumVar,
                                    width = self.labelWidth, 
                                    height = 2)
        self.label_cellNum.grid(row=3,column=0)

        self.button_generate_documents = ttk.Button(self.master, 
                            text = "Print Documents",
                            command = self.generateDocuments,
                            style = "AccentButton") 
        # self.button_generate_documents.grid(row=0,column=2) # , padx=10

        self.button_exit = ttk.Button(self.master, 
                            text = "Exit",
                            command = sys.exit) 
        # self.button_exit.grid(row=1,column=2)
  
        self.generateDocumentsHeader.pack()
        self.topFrame.pack(side="top")
        self.masterFrame.pack(pady=0)
        # self.bottomFrame.pack(pady=5)
        self.button_generate_documents.pack(pady=15)
        self.button_exit.pack()
        self.pack()

        # self.create_printing_window()

    def create_printing_window(self):
        self.printWindow = tk.Toplevel()
        self.printWindow.title("Integration I Prep")
        self.printWindow.geometry("350x350")

        self.create_printing_window_widgets()

    def create_printing_window_widgets(self):
        self.printingFrame =  ttk.Frame(self.printWindow)
        # self.printingFrame.grid(row=0,column=0)

        self.printOptionsFrame =  ttk.LabelFrame(self.printingFrame)
        self.printOptionsFrame.grid(row=0,column=0)
        
        self.printWindowHeader =  tk.Label(self.printWindow, 
                                    text = "Print Documents",
                                    font="bold", # height = 4, 
                                    ) # fg = "black")
        # self.printWindowHeader.grid(row=0,column=0)

        self.OPTIONS = self.printer.printerNames
        self.variable = tk.StringVar(self.printingFrame)
        self.variable.set(self.OPTIONS[0]) # default value

        self.printerOptionsMenu = ttk.OptionMenu(self.printOptionsFrame, self.variable, *self.OPTIONS, command=self.selectPrinter)
        self.printerOptionsMenu.grid(row=2, column=0)

        self.label_print =  tk.Label(self.printOptionsFrame, 
                                    text = "Select a printer and choose \nwhich documents to print.",
                                    width = 25, # height = 4,
                                    fg = "red")
        # self.label_print.pack()   
        self.label_print.grid(row=1,column=0)

        self.printProAndDHRSign = tk.IntVar()
        self.printInstrumentSign = tk.IntVar()
        self.printMaterialsList = tk.IntVar()
        self.printEverything = tk.IntVar()

        self.C1 = ttk.Checkbutton(self.printOptionsFrame, text = "Pro # & DHR Sign x (1)", variable = self.printProAndDHRSign, \
                        onvalue = 1, offvalue = 0,           
                        width = 20, command = self.printCheckButtonHandler)
        self.C1.grid(row = 4, column = 0)
        self.C2 = ttk.Checkbutton(self.printOptionsFrame, text = "Instrument Sign x (1)", variable = self.printInstrumentSign, \
                        onvalue = 1, offvalue = 0,           
                        width = 20, command = self.printCheckButtonHandler)
        self.C2.grid(row = 5, column = 0)
        self.C3 = ttk.Checkbutton(self.printOptionsFrame, text = "Materials List x (1)", variable = self.printMaterialsList, \
                        onvalue = 1, offvalue = 0,           
                        width = 20, command = self.printCheckButtonHandler)
        self.C3.grid(row = 6, column = 0)
        self.C4 = ttk.Checkbutton(self.printOptionsFrame, text = "All Documents", variable = self.printEverything, \
                        onvalue = 1, offvalue = 0,           
                        width = 20, command = self.printEverythingCheckButtonHandler)
        self.C4.grid(row = 3, column = 0)

        self.checkButtons.append(self.C1)
        self.checkButtons.append(self.C2)
        self.checkButtons.append(self.C3)
        self.checkButtons.append(self.C4)

        #\ height=1, for each one
        self.print_window_button_print = ttk.Button(self.printWindow, 
                            text = "Print",
                            command = self.printUserChoices,
                            style="AccentButton") 

        self.print_window_button_exit = ttk.Button(self.printWindow, 
                            text = "Exit",
                            command = sys.exit) 

        self.printWindowHeader.pack()
        self.printingFrame.pack()
        self.printOptionsFrame.pack()
        # self.C4.pack()
        # self.C1.pack()
        # self.C2.pack()
        # self.C3.pack()
        self.print_window_button_print.pack(pady=20)
        self.print_window_button_exit.pack()
        # self.pack()
    
    def printCheckButtonHandler(self):
        # If print everything checkbox is checked any other checkbox
        # is unchecked then uncheck the print all checkbox
        if self.checkButtons[3].instate(['selected']) == True and not self.checking:
            self.deselectingCheckEverything = True
            self.checkButtons[3].invoke()
            self.deselectingCheckEverything = False

    def printEverythingCheckButtonHandler(self):
        self.checking = True
        if self.checkButtons[3].instate(['selected']) == True:
            for i in range(3):
                if self.checkButtons[i].instate(['selected']) == False:
                    self.checkButtons[i].invoke()
        elif not self.deselectingCheckEverything:
            for i in range(3):
                self.checkButtons[i].invoke()
        self.checking = False

    def generateDocuments(self):
        self.statusLabelTextVar.set("IN PROGRESS")
        self.statusLabelText.configure(fg = "red") 

        # Copy materials list pdf from temp folder to exports folder
        tempPath = path.expanduser('~/AppData/Local/Temp')
        materialsListPdfSrcPath = self.find_ext(tempPath, "pdf")[0]
        materialsListPdfDstPath = os.getcwd() + "\\exports\\" + materialsListPdfSrcPath.split("\\")[-1]

        # print("src: " + src)
        # print("dst: " + dst)
        copyfile(materialsListPdfSrcPath, materialsListPdfDstPath)
        self.materialsListPath = materialsListPdfDstPath

        print(self.materialsListPath)
        self.materialsList.getData(self.materialsListPath)

        print(self.materialsListPath)
        self.materialsList.getData(self.materialsListPath)

        materialsListMsgBoxStr = self.materialsList.getMaterialsListMsgBoxStr()
        if materialsListMsgBoxStr:
            tk.messagebox.showwarning("Warning", materialsListMsgBoxStr)

        self.proNumVar.set("Pro #: " + self.materialsList.proNum)
        self.serialNumVar.set("Serial #: " + self.materialsList.serialNum)
        self.chassisNumVar.set("Chassis #: " + str(self.materialsList.chassisNum))
        self.cellNumVar.set("Cell #: " + self.materialsList.cellNum)

        if self.materialsList.chassisNum == "":
            tk.messagebox.showerror("Error", "Chassis not issued. Cannot create sign.")
        else:

            print("[INFO] Retrieved Pro Num: {} from materials list.".format(self.materialsList.proNum))
            print("[INFO] Retrieved Serial Num: {} from materials list.".format(self.materialsList.serialNum))
            print("[INFO] Retrieved Chassis Num: {} from materials list.".format(self.materialsList.chassisNum))
            print("[INFO] Calculated cell number: {}.".format(self.materialsList.cellNum))
            print("\n")

            # Modify documents to print
            self.documentHandler.createInstrumentSignPDF(self.materialsList.proNum, self.materialsList.serialNum, 
                                                        self.materialsList.chassisNum, self.materialsList.cellNum)

            self.documentHandler.createProNumAndDHR(self.materialsList.proNum, self.materialsList.serialNum)
            print("\n")

            self.statusLabelTextVar.set("COMPLETE")
            self.statusLabelText.configure(fg = "green")

            self.generatedDocuments = True

            # tk.messagebox.showinfo("Info", "Documents ready to be printed.")

            self.printUserChoices()

    def getUserChoices(self):
        userChoices = []

        if self.printEverything.get() == 1:
            userChoices.append(4)
        else:
            if self.printProAndDHRSign.get() == 1:
                userChoices.append(1)

            if self.printInstrumentSign.get() == 1:
                userChoices.append(2)

            if self.printMaterialsList.get() == 1:
                userChoices.append(3)

        return userChoices

    def selectPrinter(self, printer):
        print("Selected printer " + printer)
        self.printer.selectPrinter(printer)
        print("\n")

    def printUserChoices(self):
        if self.generatedDocuments:
            # choices = self.getUserChoices()
            choices = [3,2,2]

            if len(choices) == 0:
                tk.messagebox.showwarning("Warning", "Please select a document to print.")
                return

            for choice in choices:
                sleep(12)
                self.printer.printDocuments(self.documentHandler.newInstrumentPDFPath, self.materialsListPath, self.documentHandler.newProDHRPDFPath, choice)

            tk.messagebox.showinfo("Info", "Documents sent to printer.")
        else:
            tk.messagebox.showerror("Error", "Please generate documents first before printing.")

    def checkRequirements(self):
        if self.materialsListPath == "":
            tk.messagebox.showerror("Error", "Please select materials list PDF.")

            return False
        
        if not path.exists(self.materialsListPath):
            tk.messagebox.showerror("Error", "Please select a valid PDF file.")

            return False

        return True



root = tk.Tk()

# Create a style
style = ttk.Style(root)

# Import the tcl file
root.tk.call('source', './Azure-ttk-theme-main/azure.tcl')

# Set the theme with the theme_use method
style.theme_use('azure')

app = Application(master=root)
app.mainloop()