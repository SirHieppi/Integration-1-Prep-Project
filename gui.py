import subprocess
import tkinter as tk
import tkinter.messagebox
import sys
import os
import shutil
from time import sleep
from shutil import copyfile
from os import path
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
        self.master.geometry("600x375")
        self.materialsListPath = ""
        self.selectedPrinter = ""
        self.statusIDLEColor = "red" # "#e8d900"
        self.labelWidth = 35

        self.printer = Printer()
        self.documentHandler = DocumentHandler()
        self.materialsList = MaterialsList()
        self.sapHandler = SAPHandler()
        self.webHandler = WebHandler()

        self.instrumentVar = tk.StringVar(value = "Instrument: N/A")
        self.proNumVar = tk.StringVar(value = "Prod #: N/A")
        self.serialNumVar = tk.StringVar(value = "Serial #: N/A")
        self.chassisNumVar = tk.StringVar(value = "Chassis #: N/A")
        self.cellNumVar = tk.StringVar(value = "Cell #: N/A")
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

    def find_ext(self, dr, ext):
        return glob(path.join(dr,"*.{}".format(ext)))

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

        # Status label
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

        self.label_instrument =  tk.Label(self.statsFrame, 
                                    textvariable = self.instrumentVar,
                                    width = self.labelWidth, 
                                    height = 2)
        self.label_instrument.grid(row=0,column=0)

        self.label_proNum =  tk.Label(self.statsFrame, 
                                    textvariable = self.proNumVar,
                                    width = self.labelWidth, 
                                    height = 2)
        self.label_proNum.grid(row=1,column=0)

        self.label_serialNum =  tk.Label(self.statsFrame, 
                                    textvariable = self.serialNumVar,
                                    width = self.labelWidth, 
                                    height = 2)
        self.label_serialNum.grid(row=2,column=0)

        self.label_chassisNum =  tk.Label(self.statsFrame, 
                                    textvariable = self.chassisNumVar,
                                    width = self.labelWidth, 
                                    height = 2)
        self.label_chassisNum.grid(row=3,column=0)

        self.label_cellNum =  tk.Label(self.statsFrame, 
                                    textvariable = self.cellNumVar,
                                    width = self.labelWidth, 
                                    height = 2)
        self.label_cellNum.grid(row=4,column=0)

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

        materialsListMsgBoxStr = self.materialsList.getMaterialsListMsgBoxStr()
        if materialsListMsgBoxStr:
            tk.messagebox.showwarning("Warning", materialsListMsgBoxStr)

        self.instrumentVar.set("Instrument: " + self.materialsList.instrument)
        self.proNumVar.set("Pro #: " + self.materialsList.proNum)
        self.serialNumVar.set("Serial #: " + self.materialsList.serialNum)
        self.chassisNumVar.set("Chassis #: " + str(self.materialsList.chassisNum))
        self.cellNumVar.set("Cell #: " + self.materialsList.cellNum)

        # Only NovaSeq needs chassis number for instrument sign
        if self.materialsList.chassisNum == "" and self.materialsList.materialNumber == "20013740":
            tk.messagebox.showerror("Error", "Chassis not issued. Cannot create sign.")
        else:
            print("[INFO] Retrieved Instrument: {} from materials list.".format(self.materialsList.instrument))
            print("[INFO] Retrieved Material number: {} from materials list.".format(self.materialsList.materialNumber))
            print("[INFO] Retrieved Pro Num: {} from materials list.".format(self.materialsList.proNum))
            print("[INFO] Retrieved Serial Num: {} from materials list.".format(self.materialsList.serialNum))
            print("[INFO] Retrieved Chassis Num: {} from materials list.".format(self.materialsList.chassisNum))
            print("[INFO] Calculated cell number: {}.".format(self.materialsList.cellNum))
            print("\n")

            # Modify documents to print
            if self.materialsList.materialNumber == "20013740" or self.materialsList.materialNumber == "20046751":
                self.documentHandler.createNovaSeqInstrumentSignPDF(self.materialsList.proNum, self.materialsList.serialNum, 
                                                        self.materialsList.chassisNum, self.materialsList.materialNumber)
            elif self.materialsList.materialNumber == "15033616":
                self.documentHandler.createMiSeqInstrumentSignPDF(self.materialsList.proNum, self.materialsList.serialNum)


            print("\n")

            self.statusLabelTextVar.set("COMPLETE")
            self.statusLabelText.configure(fg = "green")

            self.generatedDocuments = True

            # tk.messagebox.showinfo("Info", "Documents ready to be printed.")

            self.printUserChoices()

    def printUserChoices(self):
        if self.generatedDocuments:
            choices = [1,1,2]

            if len(choices) == 0:
                tk.messagebox.showwarning("Warning", "Please select a document to print.")
                return

            for choice in choices:
                sleep(8)
                self.printer.printDocuments(self.documentHandler.newNovaSeqSignPDFPath, self.materialsListPath, choice)

            tk.messagebox.showinfo("Info", "Documents sent to printer.")
        else:
            tk.messagebox.showerror("Error", "Please generate documents first before printing.")

root = tk.Tk()

# Create a style
style = ttk.Style(root)

# Import the tcl file
root.tk.call('source', './Azure-ttk-theme-main/azure.tcl')

# Set the theme with the theme_use method
style.theme_use('azure')

app = Application(master=root)
app.mainloop()