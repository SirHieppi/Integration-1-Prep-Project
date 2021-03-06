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

    def printDocument(self, filename):
        allowPrinting = True

        if allowPrinting:
            win32api.ShellExecute (
                0,
                "printto",
                filename,
                '"%s"' % win32print.GetDefaultPrinter (),
                ".",
                0
            )
        else:
            print("[INFO] Printing disabled.")

    def printInstrumentSign(self, newInstrumentPDFPath):
        # If adobe acrobat reader is open then nothing will be printed
        # so adobe acrobat reader must be terminated if it exists
        delay = 10

        sleep(delay)
        try:
            os.system("TASKKILL /F /IM AcroRD32.exe")
        except:
            pass
        self.printDocument(newInstrumentPDFPath)
        # sleep(delay)
        # try:
        #     os.system("TASKKILL /F /IM AcroRD32.exe")
        # except:
        #     pass
        # sleep(delay)

    def printDocuments(self, newInstrumentPDFPath, materialsListPath, newProDHRPDFPath, userChoice):
        # print("[1] Pro # & DHR Sign x (1)")
        # print("[2] Instrument Sign x (1)")
        # print("[3] Materials list x (1)")
        # print("[4] All of the above.")
        # print("[5] Cancel.")

        if userChoice == 1:
            print("[INFO] Printing Pro # & DHR Sign ({}) x 1.".format(newProDHRPDFPath)) 
            self.printDocument(newProDHRPDFPath)
        elif userChoice == 2:
            print("[INFO] Printing Instrument Sign ({}) x 1.".format(newInstrumentPDFPath))
            self.printInstrumentSign(newInstrumentPDFPath)

        elif userChoice == 3:
            print("[INFO] Printing Materials list ({}) x 1.".format(materialsListPath))
            self.printDocument(materialsListPath)
        elif userChoice == 4:
            print("[INFO] Printing Pro # & DHR Sign ({}) x 1.".format(newProDHRPDFPath)) 
            self.printDocument(newProDHRPDFPath)
            sleep(1)
            print("[INFO] Printing Instrument Sign ({}) x 1.".format(newInstrumentPDFPath))

            self.printInstrumentSign(newInstrumentPDFPath)
            
            print("[INFO] Printing Materials list ({}) x 1.".format(materialsListPath))
            self.printDocument(materialsListPath)
        elif userChoice == 5:
            print("Exiting.")
            return

        print("Finished printing.\n")
     
