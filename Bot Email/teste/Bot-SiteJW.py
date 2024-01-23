from functools import partial
from webbrowser import get
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import StaleElementReferenceException, WebDriverException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as condicaoEsperada
from datetime import datetime
from time import sleep
from mimetypes import init
import os
import sys
import schedule
import pathlib
import shutil
import getpass
import easygui
import math

loginJw = 'cte.marba@jotaw.com' #input('Insira email de acesso: ')
passwordJw = 'ctemarba' #getpass.getpass('Password (Texto Oculto): ')

dirBaixados = os.getcwd() + '\\Baixados'
dirEnviados = os.getcwd() + '\\Enviados'
list_file = []
new_list_file = []
new_file = []

class upload:

    def esperar_elemento(self, by, element, driver):
        return bool(driver.find_elements(by,element)) 

    def valida_elemento(self,tipo, path):
        try: self.driver.find_element(by=tipo,value=path)
        except NoSuchElementException as e: return False
        return True   

    def login_email(self): #,usuario,senha

        url='https://jotaw.eslcloud.com.br/users/sign_in'
        
        options =Options()
        options.add_experimental_option('prefs',{
            "download.default_directory": dirBaixados,
            "download.Prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        options.add_argument("--start-maximized")

        self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)
        self.driver.get(url)
        self.wdw = WebDriverWait(self.driver,30)
        
        op = True 
        while op:
            userEmail = self.driver.find_element_by_id('user_email')
            passwordEmail = self.driver.find_element_by_id('user_password')
            btn_login = self.driver.find_element_by_xpath('//*[@id="new_user"]/div[2]/div/div[1]/input')
            userEmail.send_keys(loginJw)
            passwordEmail.send_keys(passwordJw)
            btn_login.click()
            op = False

        else:
            sleep(2)
        espera_btn = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div/div/div[1]/ul[2]/li[7]/a')))
        espera_btn.click()
        self.uploadXml(self.driver)
        self.driver.quit()
        sys.exit()
        
    def uploadXml(self,driver):
        mostrarMenu = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="import"]/a')))
        mostrarMenu.click()
        # menu = self.driver.find_element_by_xpath('//*[@id="ctes"]/a')
        #importacaoMenu = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="doc_e/cte/document"]/a')))
        #importacaoMenu.click()
        importacaoMenu = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="integration/import/freight"]/a/span')))
        importacaoMenu.click()
        #fretesMenu = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="edi/import/batch"]/a')))
        #fretesMenu.click()
        novaImportacao = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div[1]/div/div/div/div[2]/div/a/span')))
        novaImportacao.click()
        # novaXML = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div[1]/div/div/div/div[2]/div/ul/li[3]/a')))
        # novaXML.click()
        sleep(1)
        # procurarXML = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="integration_import_freight_documents"]')))
        # procurarXML = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.ID, 'integration_import_freight_documents')))
        procurarXML = self.driver.find_element_by_id('integration_import_freight_documents')
        # webdriver.ActionChains(self.driver).click(procurarXML).perform()
        # procurarXML.click()
        list =[r'C:\Users\Felip\Desktop\Bot Email\Baixados\135224820534183_v03.00-procCTe.xml',r'C:\Users\Felip\Desktop\Bot Email\Baixados\135224820534445_v03.00-procCTe.xml',r'C:\Users\Felip\Desktop\Bot Email\Baixados\135224821203468_v03.00-procCTe.xml']
        for l in list:
            procurarXML.send_keys(l)
            #self.driver.send_keys(Keys.ENTER)
        salvarXML = self.driver.find_element_by_id('submit')
        salvarXML.click()
        sleep(2)
        
upload_dados = upload()
upload_dados.login_email()

