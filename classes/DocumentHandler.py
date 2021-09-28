import win32com.client
import os
from docx import Document

class DocumentHandler():
    def __init__(self):
        # Paths
        self.chromedriver = os.getcwd() + "\\chromedriver_win32\\chromedriver.exe"
        self.instrumentSignExcelPath = os.getcwd() + "\\templates\\InstrumentSign_New.xlsx"
        self.proDHRSignDocPath = os.getcwd() + "\\templates\\PRO # & DHR .docx"
        self.newInstrumentPDFPath = ""
        self.newProDHRPDFPath = ""

    def createInstrumentSignPDF(self, proNum, serialNum, chassisNum, cellNum): 
        excel = win32com.client.Dispatch("Excel.Application")
        # excel.Visible = False
        wb = excel.Workbooks.Open(r'{}'.format(self.instrumentSignExcelPath))
        ws = wb.Worksheets["Input"]
        
        print("[INFO] Editing {}.".format(self.instrumentSignExcelPath))
        ws.Cells(3, 2).Value = proNum
        ws.Cells(4, 2).Value = serialNum
        ws.Cells(5, 2).Value = chassisNum
        # ws.Cells(7, 2).Value = cellNum
        
        ws = wb.Worksheets["Template_printout"]
            
        pdfPath = os.getcwd() + "\\exports\\{}_Instrument_Sign".format(serialNum) + ".pdf"
        print("[INFO] Saving instrument sign to {}".format(pdfPath))
        wb.Worksheets("Template_printout").Select()
        
        wb.ActiveSheet.ExportAsFixedFormat(0, pdfPath)
        
        wb.Close(True)
        
        excel.Quit()
        
        self.newInstrumentPDFPath = pdfPath
    
    def createProNumAndDHR(self, proNum, serialNum):
        print("[INFO] Editing {}.".format(self.proDHRSignDocPath))
        document = Document(self.proDHRSignDocPath)

        debugPrint = False

        serialNum = serialNum[1:]
        serialNumIndex = 0
        foundA = False
        lineIndex = 0

        # Finds the start of serial number and replaces one by one since word doc txt may be split up
        for paragraph in document.paragraphs:
            for line in paragraph.runs:
                replaceLine = ""
                for char in line.text:
                    if char == 'A':
                        foundA = True
                    if foundA and char != 'A' and serialNumIndex < len(serialNum):
                        replaceLine += serialNum[serialNumIndex]
                        serialNumIndex += 1
                    else:
                        replaceLine += char
                if debugPrint:
                    print("'" + line.text + "'" + " -> '" + replaceLine + "'")
                lineIndex += 1
                if replaceLine != line.text:
                    line.text = replaceLine
        
        inline = document.paragraphs[2].runs
        inline[-1].text = proNum
        
        # document.save(self.proDHRSignDocPath)

        self.newProDHRPDFPath = os.getcwd() + "\\exports\\A{}_ProDHR_Sign".format(serialNum) + ".docx"
        print(self.newProDHRPDFPath)
        # convert(self.proDHRSignDocPath, proDHRPDFPath)

        document.save(self.newProDHRPDFPath)