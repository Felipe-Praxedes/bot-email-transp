from functools import partial
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
import openpyxl
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import StaleElementReferenceException, WebDriverException
import getpass
import easygui
from datetime import datetime
import sys
import pyautogui as gui
import time
import math
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options

class Dados:

    def dados_usuario(self):
        loginEmail = 'cte.marba@jotaw.com' #input('Insira email de acesso: ')
        passwordEmail = '@Jotaw347895' #getpass.getpass('Password (Texto Oculto): ')
        loginJw = 'cte.marba@jotaw.com' #input('Insira email de acesso: ')
        passwordJw = 'ctemarba' #getpass.getpass('Password (Texto Oculto): ')

        return loginEmail, passwordEmail

    def esperar_elemento(self, by, element, driver):
        return bool(driver.find_elements(by,element)) 

    def valida_elemento(self,tipo, path):
        try: self.driver.find_element(by=tipo,value=path)
        except NoSuchElementException as e: return False
        return True   

    def login_email(self): #,usuario,senha
        url='https://webmail1.hostinger.com.br/'
        options = webdriver.ChromeOptions() 
        options.add_argument("download.default_directory=C:\\Users\\Felip\\Desktop\\este")
        self.driver = webdriver.Chrome(executable_path=r'./chromedriver.exe', chrome_options=options)
        self.driver.get(url)
        wdw = WebDriverWait(self.driver,15)
        op = True 
        while op:
            userEmail = self.driver.find_element_by_id('rcmloginuser')
            passwordEmail = self.driver.find_element_by_id('rcmloginpwd')
            btn_login = self.driver.find_element_by_id('rcmloginsubmit')
            self.loginEmail, self.passwordEmail = self.dados_usuario()
            userEmail.send_keys(self.loginEmail)
            passwordEmail.send_keys(self.passwordEmail)
            btn_login.click()
            op = False
            # erro = self.driver.find_elements_by_xpath('//*[@id="messagestack"]/div/span and (@class ="warning content ui alert alert-warning")]')
            # if len(erro)>0:
            #    print(erro[0].text)
            #    op = True
            #    print('Programa Finalizado')
            # else:
            #    op = False
        else:
            sleep(2)
        if self.valida_elemento(By.CLASS_NAME, 'button icon toolbar-button refresh') == True:
            espera_btn = partial(self.esperar_elemento,By.CLASS_NAME, 'button icon toolbar-button refresh')
            valida_btn = wdw.until(espera_btn)
            if valida_btn == True:
                btn_att = self.driver.find_element_by_class_name('button icon toolbar-button refresh')
                btn_att.click()
        else:
            # self.driver.quit()
            time.sleep(1)
            self.baixarXml(self.driver)
            self.driver.quit()
            sys.exit()
        
    def baixarXml(self,driver):
        menu = driver.find_element_by_xpath('//*[@id="mailsearchform"]')
        menu.send_keys('cteseara')
        menu.click()

        gui.press('enter')
        time.sleep(2)

        # erro = self.driver.find_element_by_xpath('//*[@id="messagestack"]/div/span and (@class ="confirmation content ui alert alert-success")]')
        # op = True
        # while op:
        #    if len(erro)>0:
        #        op = False
        #    else:
        #        op = True
        
  
        lista_email = self.driver.find_elements_by_class_name('rcmContactAddress')
        
        for email in lista_email:
            email.click()
            # time.sleep(2)
            # xml_email = self.driver.find_elements_by_xpath('//*[@id="attachmenudownload"]')
            # xml_email2 = self.driver.find_elements_by_class_name('button icon dropdown skip-content')
            xml_email = self.driver.find_element_by_css_selector('#attach2')
            # xml_email4 = self.driver.find_elements_by_id('attachmenudownload')
            # xml_email = self.driver.find_elements_by_tag_name('a')
            for xml in xml_email:
                xml.click()

extrair_dados = Dados()
extrair_dados.login_email()
