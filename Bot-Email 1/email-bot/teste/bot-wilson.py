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

loginEmail = 'cte.marba@jotaw.com' #input('Insira email de acesso: ')
passwordEmail = '@Jotaw347895' #getpass.getpass('Password (Texto Oculto): ')
loginJw = 'cte.marba@jotaw.com' #input('Insira email de acesso: ')
passwordJw = 'ctemarba' #getpass.getpass('Password (Texto Oculto): ')

dirBaixados = os.getcwd() + '\\Baixados'
dirEnviados = os.getcwd() + '\\Enviados'
list_file = []
new_list_file = []
new_file = []
# função que lista todos os arquivos da pasta
def lista_arquivos_inicial(dirBaixados):
    lista = []
    for root, dirs, files in os.walk(dirBaixados):
        for file in files:
            lista.append(os.path.join(root,file))
    return lista
# valida se o dirBaixados do diretorio é valido
def valida_dir(self):
    if os.path.isdir(pathlib.Path(self)):
        return True
    else:
        print('Diretório inserido não é valido!!')
        return False

class Dados:

    #def dados_usuario(self):
    #    self.loginEmail = 'cte.marba@jotaw.com' #input('Insira email de acesso: ')
    #    self.passwordEmail = '@Jotaw347895' #getpass.getpass('Password (Texto Oculto): ')
    #    self.loginJw = 'cte.marba@jotaw.com' #input('Insira email de acesso: ')
    #    self.passwordJw = 'ctemarba' #getpass.getpass('Password (Texto Oculto): ')

    #    return loginEmail, passwordEmail, loginJW, passwordJW

    def esperar_elemento(self, by, element, driver):
        return bool(driver.find_elements(by,element)) 

    def valida_elemento(self,tipo, path):
        try: self.driver.find_element(by=tipo,value=path)
        except NoSuchElementException as e: return False
        return True   

    def login_email(self): #,usuario,senha
        url='https://webmail1.hostinger.com.br/'
        
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
            userEmail = self.driver.find_element_by_id('rcmloginuser')
            userPassword = self.driver.find_element_by_id('rcmloginpwd')
            btn_login = self.driver.find_element_by_id('rcmloginsubmit')
            #self.loginEmail, self.passwordEmail = self.dados_usuario()
            userEmail.send_keys(loginEmail)
            userPassword.send_keys(passwordEmail)
            btn_login.click()
            op = False

        else:
            sleep(2)
        if self.valida_elemento(By.CLASS_NAME, 'button icon toolbar-button refresh') == True:
            espera_btn = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.CLASS_NAME, 'button icon toolbar-button refresh')))
            valida_btn = self.wdw.until(espera_btn)
            if valida_btn == True:
                btn_att = self.driver.find_element_by_class_name('button icon toolbar-button refresh')
                btn_att.click()
        else:
            sleep(1)
            self.baixarXml(self.driver)
            #self.driver.quit()
            # sys.exit()
        
    def baixarXml(self,driver):
        menu = self.driver.find_element_by_xpath('//*[@id="mailsearchform"]')
        mostrarMsg = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//a[@title="Mostrar mensagens não lidas"]')))
        mostrarMsg.click()
        menu.send_keys('cteseara')
        menu.click()
        menu.send_keys(Keys.ENTER)
        sleep(2)

        while len(self.driver.find_elements_by_class_name('rcmContactAddress'))>0:
            email = self.driver.find_element_by_xpath("//table[@id='messagelist']/tbody/tr[1]/td[2]/span[@class ='subject']/a")
            # duploclick no email
            webdriver.ActionChains(self.driver).double_click(email).perform()
            # localiza o arquivo xml
            xml_email = self.driver.find_element_by_xpath("//ul[@id='attachment-list']/li/a")
            xml_email.click()
            self.driver.back()
            mostrarMsg = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//a[@title="Mostrar mensagens não lidas"]')))
        

class TransferenciaArquivos:
    
    def iniciar(self, old_dir,new_dir):
       new_list_file =  self.lista_arquivos_inicial(old_dir)
       #  new_file = self.separa_arquivos_novos(list_file, new_list_file)
       self.copia_arquivos(new_dir,new_list_file)
    # função que lista todos os arquivos dentro da pasta
    def lista_arquivos_inicial(self, dirBaixados):
        lista = []
        for root, dirs, files in os.walk(dirBaixados):
            for file in files:
                lista.append(os.path.join(root,file))
        return lista

    # função que transfere os arquivos para a pasta de destino
    def copia_arquivos(self, dirBaixados, list):
        if len(list) == 0:
            pass
        else:
            for file in list:
                name_arq = pathlib.Path(file).name
                new_dir = os.path.join(dirBaixados,name_arq)
                list_file.append(file)
                try:
                    shutil.move(file, new_dir)
                    print(f'O arquivo {file} foi transferido com sucesso')
                except Exception as e:
                    print(f'Não foi possivel copiar o arquivo {name_arq}')
                    continue

try:
    origem = dirBaixados # input("Digite o diretório de origem: ").strip()
    var = valida_dir(origem)
    if var == False:
        sys.exit()
    destino = dirEnviados # input("Digite o diretório de Destino: ").strip()
    var = valida_dir(destino)
    if var == False:
        sys.exit()
    
except KeyboardInterrupt:
    print()
    print('O Programa foi finalizado')
    sys.exit()

class upload:

    def esperar_elemento(self, by, element, driver):
        return bool(driver.find_elements(by,element)) 

    def valida_elemento(self,tipo, path):
        try: self.driver.find_element(by=tipo,value=path)
        except NoSuchElementException as e: return False
        return True   

    def login_email(self):

        url='https://jotaw.eslcloud.com.br/users/sign_in'
        
        # options =Options()
        # options.add_experimental_option('prefs',{
        #     "download.default_directory": dirBaixados,
        #     "download.Prompt_for_download": False,
        #     "download.directory_upgrade": True,
        #     "safebrowsing.enabled": True
        # })
        # options.add_argument("--start-maximized")

        # self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)
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
        # sys.exit()
        
    def uploadXml(self,driver):
        mostrarMenu = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="import"]/a')))
        mostrarMenu.click()
        importacaoMenu = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="integration/import/freight"]/a/span')))
        importacaoMenu.click()
        novaImportacao = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div[1]/div/div/div/div[2]/div/a/span')))
        novaImportacao.click()
        sleep(1)
        procurarXML = self.driver.find_element_by_id('integration_import_freight_documents')
        list =[r'C:\Users\Felip\Desktop\Bot Email\Baixados\135224820534183_v03.00-procCTe.xml',r'C:\Users\Felip\Desktop\Bot Email\Baixados\135224820534445_v03.00-procCTe.xml',r'C:\Users\Felip\Desktop\Bot Email\Baixados\135224821203468_v03.00-procCTe.xml']
        for l in list:
            procurarXML.send_keys(l)
            #self.driver.send_keys(Keys.ENTER)
        # salvarXML = self.driver.find_element_by_id('submit')
        # salvarXML.click()
        sleep(2)

# dowload dos dados        
extrair_dados = Dados()
extrair_dados.login_email()

# upload dos arquivos baixados
upload_dados = upload()
upload_dados.login_email()

program = TransferenciaArquivos()
# definição de rotina (frequencia que sera executado)
list_file = lista_arquivos_inicial(dirBaixados)
program.iniciar(origem,destino)


