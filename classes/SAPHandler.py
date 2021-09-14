import os
from os import path

class SAPHandler():
    def __init__(self):
        self.sapScriptPath = os.getcwd() + "\\integration prep.vbs"

    def runSAPScript(self, docNum):
        self.modifySAPScript(docNum)

    def modifySAPScript(self, docNum):
        sapScript = open(self.sapScriptPath, "r")
        list_of_lines = sapScript.readlines()
        list_of_lines[21] = "session.findById(\"wnd[0]/usr/ctxtP_AUFNR\").text = \"{}\"\n".format(docNum)

        sapScript = open(self.sapScriptPath, "w")
        sapScript.writelines(list_of_lines)
        sapScript.close()