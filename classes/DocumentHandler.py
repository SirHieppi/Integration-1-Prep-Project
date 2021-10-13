import win32com.client
import os
from docx import Document

class DocumentHandler():
    def __init__(self):
        # Paths
        self.chromedriver = os.getcwd() + "\\chromedriver_win32\\chromedriver.exe"
        self.novaSeqSignExcelPath = os.getcwd() + "\\templates\\InstrumentSign_New.xlsx"
        self.miSeqSignExcelPath = os.getcwd() + "\\templates\\DEVN 1046978 WIN10 Heimdall Template.xlsx"
        self.newNovaSeqSignPDFPath = ""
        self.newMiSeqSignPDFPath = ""

    def createNovaSeqInstrumentSignPDF(self, proNum, serialNum, chassisNum, materialNumber): 
        excel = win32com.client.Dispatch("Excel.Application")
        # excel.Visible = False
        wb = excel.Workbooks.Open(r'{}'.format(self.novaSeqSignExcelPath))
        ws = wb.Worksheets["Input"]
        
        print("[INFO] Editing {}.".format(self.novaSeqSignExcelPath))
        ws.Cells(3, 2).Value = proNum
        ws.Cells(4, 2).Value = serialNum
        ws.Cells(5, 2).Value = chassisNum
        
        # Select different template version for china unit
        if materialNumber == "20046751":
            ws.Cells(6,2).Value = "CA 2.0 (MAH)"
        elif materialNumber == "20013740":
            ws.Cells(6,2).Value = "CA 1.5"
        
        ws = wb.Worksheets["Template_printout"]
            
        pdfPath = os.getcwd() + "\\exports\\{}_NovaSeq_Instrument_Sign".format(serialNum) + ".pdf"
        print("[INFO] Saving NovaSeq instrument sign to {}".format(pdfPath))
        wb.Worksheets("Template_printout").Select()
        
        wb.ActiveSheet.ExportAsFixedFormat(0, pdfPath)
        
        wb.Close(True)
        
        excel.Quit()
        
        self.newNovaSeqSignPDFPath = pdfPath

    def createMiSeqInstrumentSignPDF(self, proNum, serialNum):
        excel = win32com.client.Dispatch("Excel.Application")
        # excel.Visible = False
        wb = excel.Workbooks.Open(r'{}'.format(self.miSeqSignExcelPath))
        ws = wb.Worksheets["MISeqRUO"]
        
        print("[INFO] Editing {}.".format(self.miSeqSignExcelPath))
        ws.Cells(2, 2).Value = proNum
        ws.Cells(1, 2).Value = serialNum
                    
        pdfPath = os.getcwd() + "\\exports\\{}_Instrument_Sign".format(serialNum) + ".pdf"
        print("[INFO] Saving MiSeq instrument sign to {}".format(pdfPath))
        wb.Worksheets("MISeqRUO").Select()

        try:
            wb.ActiveSheet.ExportAsFixedFormat(0, pdfPath)
        except:
            print("Cannot save " + pdfPath)
        
        wb.Close(True)
        
        excel.Quit()
        
        self.newNovaSeqSignPDFPath = pdfPath