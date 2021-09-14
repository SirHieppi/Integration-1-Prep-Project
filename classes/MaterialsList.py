import tabula
import PyPDF2
import math
from tabula import read_pdf
from pathlib import Path

class MaterialsList():
    def  __init__(self):
        self.totalNovaseqMaterials = {
            'SHIPPING BRKT, FLUIDICS MODULE, CUSTOM': 2, 
            'DUCT, TUBING': 1, 
            'ASSY, RCA, TL, NOVASEQ': 2, 
            'ASSY, PHANTOM BIM, TL': 2, 
            'BRACKET, TUBE SUPPORT': 1, 
            'CBL,APX_MISC_H22-OM GROUNDING_FORK': 2,
            'ASSY, FLUIDICS, MODULE,V2': 1, 
            'PCA, RFID Integrated Antenna': 1, 
            'ASSEMBLY, CABLE TRACK': 1, 
            'VIB ISOLATOR ASSY, TUNED DAMPER, REAR': 2, 
            'VIB ISOLATOR ASSY, TUNED DAMPER, FRONT': 2, 
            'MOTION CONTROLLER, C413': 1, 
            'ASSY,COMPUTE ENGINE, NORTHSHORE,NOVASEQ': 1,
            'ASSY, OPA W/ NOZZLE': 1, 
            'ASSY, XY STAGE MODULE': 1, 
            'ASSY, CAMERA MODULE (CAM)': 1,
            'ASSY, FOCUS TRACKING MODULE (FTM)': 1,
            'ASSY, EMISSION OPTICS MODULE (EOM)': 1,
            'LGM_VES': 1,
            'ASSY, CHASSIS, NOVASEQ, V1.5': 1,
            'ASSY, DUAL ACTUATION DECK 2.0': 1    
        }    
        self.proNum = ""
        self.serialNum = ""
        self.chassisNum = ""
        self.materialsListPath = ""
        self.cellNum = ""

    def getData(self, materialsListPath):
        if not self.checkMaterialsList(materialsListPath):
            print("Materials list not valid.")

        self.extractDataFromMaterialsList(materialsListPath)
    
    def getMaterialsListMsgBoxStr(self):
        ret = ""
        if self.missing:
            ret += self.missingPartsStr 

        if self.surplus:
            ret += "\n" + self.surplusPartsStr

        return ret
    
    def createMaterialsListDict(self, materialsListPath):
        ret = {}
        self.missingPartsStr = "Missing parts in materials list:\n"
        self.surplusPartsStr = "Surplus parts in materials list:\n"
        self.missing = False
        self.surplus = False
        
        tables = tabula.read_pdf(materialsListPath, pages='all', lattice=True)



        for table in tables:
            index = 0

            print(table)

            # print("table keys:")

            # print(table.keys())

            if 'Description' in table.keys():
                for description in table['Description']:
                    if description != 'N/A':
                        key = str(description).replace('\r', ' ')

                        if key in ret:
                            ret[key] += 1
                        else:
                            if 'Qty Required' in table:
                                ret[key] = int(table['Qty Required'][index])

                        index += 1
        print('ret:')
        print(ret)
        return ret

    def checkMaterialsList(self, materialsListPath):
        # materialsList = {
        #     'SHIPPING BRKT, FLUIDICS MODULE, CUSTOM': 1, 
        #     'DUCT, TUBING': 1, 
        #     'ASSY, RCA, TL, NOVASEQ': 5, 
        #     'ASSY, PHANTOM BIM, TL': 1, 
        #     'BRACKET, TUBE SUPPORT': 1, 
        #     'CBL,APX_MISC_H22-OM GROUNDING_FORK': 1, 
        #     'PCA, RFID Integrated Antenna': 1, 
        #     'VIB ISOLATOR ASSY, TUNED DAMPER, REAR': 1, 
        #     'VIB ISOLATOR ASSY, TUNED DAMPER, FRONT': 1, 
        #     'MOTION CONTROLLER, C413': 1, 
        #     'ASSY, OPA W/ NOZZLE': 1, 
        #     'ASSY, XY STAGE MODULE': 1, 
        #     'ASSY, CAMERA MODULE (CAM)': 1,
        #     'ASSY, DUAL ACTUATION DECK 2.0': 1    
        # } 
        materialsList = self.createMaterialsListDict(materialsListPath)

        print("[INFO] Checking materials list...")
        for material in self.totalNovaseqMaterials:
            # print(missingPartsStr)
            if not material in materialsList:
                self.missingPartsStr += "* " + material + ": {} \n".format(self.totalNovaseqMaterials[material])
                self.missing = True
            else:
                diff = self.totalNovaseqMaterials[material] - materialsList[material]

                if diff < 0:
                    # surplus
                    self.surplusPartsStr += "* " + material + ": {} \n".format(abs(diff))
                    self.surplus = True
                elif diff > 0:
                    # missing
                    self.missingPartsStr += "* " + material + ": {} \n".format(diff)
                    self.missing = True

        if self.missing and self.surplus:
            print(self.missingPartsStr)
            print(self.surplusPartsStr)
            return False
        elif self.missing:
            print(self.missingPartsStr)
            return False
        elif self.surplus:
            print(self.surplusPartsStr)
            return False
        else:
            print("[INFO] Materials list contains all required materials.")
            # print(self.missing)
            # print(self.surplus)
            # print(self.missingPartsStr)
            # print(self.surplusPartsStr)
            return True

    def extractDataFromMaterialsList(self, materialsListPath):    
        # creating a pdf file object  
        pdfFileObj = open(materialsListPath, 'rb')  
            
        # creating a pdf reader object  
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)   
            
        # creating a page object  
        pageObj = pdfReader.getPage(0)  
            
        # extracting text from page  
        text = pageObj.extractText()
        
        proNum = ""
        serialNum = ""
        chassisNum = ""
        cellNum = ""

        # Get pro num
        after = text.partition("Prod Order #")[2]
        index = 0
        while after[index].isdigit():
            proNum += after[index]
            index += 1

        proNum = proNum.lstrip("0")
        
        # Get serial num
        after = text.partition("Serial #")[2]
        serialNum = after[:6]
            
        # closing the pdf file object  
        pdfFileObj.close() 

        # Get chassis num from table in pdf
        tables = tabula.read_pdf(materialsListPath, pages = "all")
        
        tableIndex = 0
        for table in tables: 
            descriptionIndex = 0
            if 'Description' in table.keys():
                for description in table['Description']:
                    if isinstance(description, str) and "ASSY, CHASSIS," in description:
                        chassisNum = tables[tableIndex]["Batch # / Serial #"][descriptionIndex]

                        # sometimes serial number is on row below because each material is broken up into several rows in the table
                        if math.isnan(chassisNum):
                            # serial a row below "ASSY, CHASSIS,"
                            chassisNum = tables[tableIndex]["Batch # / Serial #"][descriptionIndex + 1]

                        break
                    descriptionIndex += 1
                    
            tableIndex += 1
            
        self.proNum = proNum
        self.serialNum = serialNum
        self.chassisNum = chassisNum
        self.cellNum = str(self.calculateCellNum())

    def calculateCellNum(self):
        cellNum = int(self.serialNum[1:]) % 3

        if cellNum == 0:
            return 3
        
        return cellNum