import os
from os import path
from pathlib import Path

class SAPHandler():
    def __init__(self):
        path = Path(os.getcwd())
        pathStr = os.path.dirname(path)

        self.sapScriptPath = pathStr + "\\SAP Script.vbs"

        print("SAP Script.vbs path: " + self.sapScriptPath)

    def runSAPScript(self, docNum):
        self.modifySAPScript(docNum)

    def modifySAPScript(self, docNum):
        sapScript = open(self.sapScriptPath, "r")
        list_of_lines = sapScript.readlines()
        list_of_lines[21] = "session.findById(\"wnd[0]/usr/ctxtP_AUFNR\").text = \"{}\"\n".format(docNum)

        sapScript = open(self.sapScriptPath, "w")
        sapScript.writelines(list_of_lines)
        sapScript.close()