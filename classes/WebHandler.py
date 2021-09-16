import selenium
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from pathlib import Path
from webdriver_manager.chrome import ChromeDriverManager

class WebHandler():
    def __init__(self) -> None:
        self.sapURL = "https://sapportalprd.illumina.com/irj/servlet/prt/portal/prteventname/Navigate/prtroot/com.sap.portal.appintegrator.sap.Transaction?AppIntegratorVariant=com.sap.portal.appintegrator.sap.Transaction&ApplicationParameter=&DynamicParameter=&AutoStart=false&OkCode=&UseSPO1=false&DebugMode=false&System=SAP_ECC&TCode=SMEN&Technique=SSF&GuiType=WinGui"

    def visit(self, url):
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
        print("[INFO] Visiting {}".format(url))
        self.driver.get(url)

        time.sleep(5)

        self.driver.close()

    def openSAP(self):
        self.visit(self.sapURL)

    def closeBrowser(self):
        self.driver.close()
