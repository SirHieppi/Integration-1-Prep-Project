import tabula
import PyPDF2
import math
import json

class MaterialsList():
    def  __init__(self):
        self.instrument = ""
        self.materialNumber = ""
        self.proNum = ""
        self.serialNum = ""
        self.chassisNum = ""
        self.materialsListPath = ""
        self.cellNum = ""

    def getData(self, materialsListPath):
        self.extractDataFromMaterialsList(materialsListPath)

        if not self.checkMaterialsList(materialsListPath):
            print("Materials list not valid.")
    
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

            # if 'Description' in table.keys():
            #     for description in table['Description']:
            #         if description != 'N/A':
            #             key = str(description).replace('\r', ' ')

            #             if key in ret:
            #                 if table['Mvt. Typ.'][index] == 261:
            #                     ret[key] += 1
            #                 elif table['Mvt. Typ.'][index] == 262:
            #                     ret[key] -= 1
            #             else:
            #                 if 'Qty Issued' in table:
            #                     if math.isnan(table['Qty Issued'][index]):
            #                         ret[key] = 0
            #                     else:
            #                         if not key in ret:
            #                             ret[key] = 0

            #                         if table['Mvt. Typ.'][index] == 261:
            #                             ret[key] += int(table['Qty Issued'][index])
            #                         elif table['Mvt. Typ.'][index] == 262:
            #                             ret[key] -= int(table['Qty Issued'][index])

            #             index += 1
            if 'Material #' in table.keys():
                for materialNum in table['Material #']:
                    materialNum = str(materialNum)
                    if materialNum != 'N/A':
                        description = str(table["Description"][index]).replace('\r', ' ')

                        if materialNum in ret:
                            if table['Mvt. Typ.'][index] == 261:
                                ret[materialNum][1] += 1
                            elif table['Mvt. Typ.'][index] == 262:
                                ret[materialNum][1] -= 1
                        else:
                            if 'Qty Issued' in table:
                                if math.isnan(table['Qty Issued'][index]):
                                    ret[materialNum][1] = 0
                                else:
                                    if not materialNum in ret:
                                        ret[materialNum] = [description, 0]

                                    if table['Mvt. Typ.'][index] == 261:
                                        ret[materialNum][1] += int(table['Qty Issued'][index])
                                    elif table['Mvt. Typ.'][index] == 262:
                                        ret[materialNum][1] -= int(table['Qty Issued'][index])
                    index += 1

        print('Materials list:')
        print(ret)
        return ret

    def checkMaterialsList(self, materialsListPath):
        materialsList = self.createMaterialsListDict(materialsListPath)

        f = open("./Entire BOM List.json")
        bom = json.loads(f.read())

        print(bom[self.materialNumber])

        print("[INFO] Checking materials list...")
        # for material in bom[self.materialNumber]:
        #     # print(missingPartsStr)
        #     if not material == "NAME":
        #         if not material in bom[self.materialNumber]:
        #             self.missingPartsStr += "* " + material + ": {} \n".format(bom[self.materialNumber])
        #             self.missing = True
        #         else:
        #             if material not in materialsList:
        #                 diff = bom[self.materialNumber][material]
        #             else:
        #                 diff = bom[self.materialNumber][material] - materialsList[material]

        #             if diff < 0:
        #                 # surplus
        #                 self.surplusPartsStr += "* " + material + ": {} \n".format(abs(diff))
        #                 self.surplus = True
        #             elif diff > 0:
        #                 # missing
        #                 self.missingPartsStr += "* " + material + ": {} \n".format(diff)
        #                 self.missing = True
        for material in materialsList:
            # print(missingPartsStr)
            if not material == "NAME":
                if not material in bom[self.materialNumber] and material in bom[self.materialNumber]:
                    self.missingPartsStr += "* " + str(material) + ": {} \n".format(bom[self.materialNumber][material][0])
                    self.missing = True
                
                # Ignores materials that are not on the BOM
                elif material in bom[self.materialNumber]:
                    if material not in materialsList:
                        diff = bom[self.materialNumber][material][1]
                    else:
                        diff = bom[self.materialNumber][material][1] - materialsList[material][1]

                    if diff < 0:
                        # surplus
                        self.surplusPartsStr += "* " + bom[self.materialNumber][material][0] + ": {} \n".format(abs(diff))
                        self.surplus = True
                    elif diff > 0:
                        # missing
                        self.missingPartsStr += "* " + bom[self.materialNumber][material][0] + ": {} \n".format(diff)
                        self.missing = True
        f.close()

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
        
        instrument = text.partition('Batch')[0].partition('Material Desc')[2]
        materialNumber = text.partition('Material Number')[2].partition('Header')[0]
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
        
        # # NovaSeq
        # if materialNumber == "20013740":
        #     after = text.partition("Serial #")[2]
        #     serialNum = after[:6]
        # # MiSeq
        # elif materialNumber == "15033616":
        #     serialNum = text.partition("Serial")[2].partition("Quantity")[0][8:]
            
        # closing the pdf file object  
        pdfFileObj.close() 

        # Get chassis num from table in pdf
        tables = tabula.read_pdf(materialsListPath, pages = "all")

        # Get serial num
        if materialNumber == "20013740" or "20046751":
            serialNum = tables[-1].keys()[0][4:]
        else:
            serialNum = tables[-1].keys()[0]
        
        tableIndex = 0

        # Find chassis num
        for table in tables: 
            descriptionIndex = 0
            # if 'Description' in table.keys():
            #     for description in table['Description']:
            #         if isinstance(description, str) and "ASSY, CHASSIS," in description:
            #             chassisNum = tables[tableIndex]["Batch # / Serial #"][descriptionIndex]

            #             # sometimes serial number is on row below because each material is broken up into several rows in the table
            #             if isinstance(chassisNum, int) and math.isnan(chassisNum):
            #                 # serial a row below "ASSY, CHASSIS,"
            #                 chassisNum = tables[tableIndex]["Batch # / Serial #"][descriptionIndex + 1]

            #             break
            #         descriptionIndex += 1
            if 'Material #' in table.keys():
                for material in table['Material #']:
                    if isinstance(material, int) and "ASSY, CHASSIS," in tables[tableIndex]["Description"][descriptionIndex]:
                        chassisNum = tables[tableIndex]["Batch # / Serial #"][descriptionIndex]
                    descriptionIndex += 1
                    
            tableIndex += 1
            
        self.instrument = instrument
        self.materialNumber = materialNumber
        self.proNum = proNum
        self.serialNum = serialNum
        self.chassisNum = chassisNum

        if self.materialNumber == "20013740" or self.materialNumber == "20046751":
            self.cellNum = str(self.calculateCellNum())

    def calculateCellNum(self):
        startingIndex = 0
        for char in self.serialNum:
            if not char.isdigit():
                startingIndex += 1

        cellNum = int(self.serialNum[startingIndex:]) % 3

        if cellNum == 0:
            return 3
        
        return cellNum