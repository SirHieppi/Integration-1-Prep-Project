import subprocess
import shlex
import io
import win32print
import sys
import win32api
import os
from time import sleep

class Printer():
    def __init__(self):
        self.selectedPrinter = ""
        self.printerNames = self.getPrinters()
    
    def getPrinters(self):
        self.printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 2)
        printerNames = []

        for printer in self.printers:
            printerNames.append(printer['pPrinterName'])

        return printerNames

    def selectPrinter(self, printerChosen):
        win32print.SetDefaultPrinter(self.printers[self.printerNames.index(printerChosen)]['pPrinterName'])

        print("[INFO] Default printer is now set to {}.".format(win32print.GetDefaultPrinter()))

    def printDocument(self, filename, copies=1):
        allowPrinting = True

        print("Printing " + filename)

        name = win32print.GetDefaultPrinter()
        #printdefaults = {"DesiredAccess": win32print.PRINTER_ACCESS_ADMINISTER}
        printdefaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE}
        handle = win32print.OpenPrinter(name, printdefaults)
        level = 2
        attributes = win32print.GetPrinter(handle, level)
        print("Old Duplex = %d" % attributes['pDevMode'].Duplex)
        #attributes['pDevMode'].Duplex = 1    # no flip
        #attributes['pDevMode'].Duplex = 2    # flip up
        #attributes['pDevMode'].Duplex = 3    # flip over
        # attributes['pDevMode'].Color = 1
        attributes['pDevMode'].Copies = copies
        print("setting copies to " + str(attributes['pDevMode'].Copies) + " for " + filename)
        attributes['pDevMode'].Duplex = 2
        ## 'SetPrinter' fails because of 'Access is denied.'
        ## But the attribute 'Duplex' is set correctly

        try:
            win32print.SetPrinter(handle, level, attributes, 0)
        except:
            print("win32print.SetPrinter: set 'Duplex'")
        res = win32api.ShellExecute(0, 'print', filename, None, '.', 0)
        
        win32print.ClosePrinter(handle)

    def printInstrumentSign(self, newInstrumentPDFPath):
        # If adobe acrobat reader is open then nothing will be printed
        # so adobe acrobat reader must be terminated if it exists
        # delay = 5

        # sleep(delay)
        try:
            os.system("TASKKILL /F /IM AcroRD32.exe")
        except:
            pass
        self.printDocument(newInstrumentPDFPath)

    def printDocuments(self, newInstrumentPDFPath, materialsListPath, newProDHRPDFPath, choice):
        if choice == 1:
            self.printInstrumentSign(newInstrumentPDFPath)
        elif choice == 2:
            self.printDocument(materialsListPath)

        print("Finished printing.\n")
     
