from functools import partial
from tokenize import String
from typing import List
from webbrowser import get
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as condicaoEsperada
from datetime import datetime
from time import sleep
from mimetypes import init
import os
import sys
import pathlib
import shutil

loginEmail = 'cte.marba@jotaw.com'
passwordEmail = '@Jotaw347895'  

loginJw = 'cte.marba@jotaw.com'
passwordJw = 'ctemarba'

urlEmail = 'https://webmail1.hostinger.com.br/'
urlJw = 'https://jotaw.eslcloud.com.br/users/sign_in'
urlFretes = 'https://jotaw.eslcloud.com.br/freight/normals/new'

global matriz

class Dados:

    def login_email(self):

        options = Options()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_experimental_option(
            'prefs', {
            "download.Prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        options.add_argument("--start-maximized")

        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        sleep(0.5)
        self.driver.get(urlEmail)
        self.wdw = WebDriverWait(self.driver, 40)

        op = True
        while op:
            userEmail = self.driver.find_element(By.ID,'rcmloginuser')
            userPassword = self.driver.find_element(By.ID,'rcmloginpwd')
            btn_login = self.driver.find_element(By.ID,'rcmloginsubmit')
            userEmail.send_keys(loginEmail)
            print(f'Email {loginEmail} utilizado')
            userPassword.send_keys(passwordEmail)
            btn_login.click()
            op = False

        else:
            sleep(2)

        self.baixarXml(self.driver)
        
        self.driver.get(urlJw)
        sleep(2)

        op = True
        while op:
            userJw = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.ID, 'user_email')))
            userPassword = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.ID, 'user_password')))
            btn_login = self.driver.find_element(By.XPATH,'//*[@id="new_user"]/div[2]/div/div[1]/input')
            userJw.send_keys(loginJw)
            print(f'Email {loginJw} utilizado')
            userPassword.send_keys(passwordJw)
            btn_login.click()
            op = False

        else:
            sleep(2)
        self.consultaChave(self.driver)

        self.driver.get(urlFretes)

        self.createManifest(self.driver)

        self.driver.quit() 

    def baixarXml(self, driver):
        menu = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="mailsearchform"]')))
        mostrarMsg = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//a[@title="Mostrar mensagens não lidas"]')))
        mostrarMsg.click()
        menu.send_keys('cacau')
        menu.send_keys(Keys.ENTER)
        wdwToEmail = WebDriverWait(self.driver, 5)

        try:
            esperarEmail = wdwToEmail.until(condicaoEsperada.element_to_be_clickable((By.CLASS_NAME, 'rcmContactAddress')))
        except:
            pass
        sleep(0.5)

        tx = 0
        linha = [0] * 5
        hTabelaAnterior = 0
        self.matrizMapa = []
        self.matriz = []
        while len(self.driver.find_elements(By.CLASS_NAME,'rcmContactAddress')) > 0: 
            email = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, "//table[@id='messagelist']/tbody/tr[1]/td[2]/span[@class ='subject']/a")))
            webdriver.ActionChains(self.driver).double_click(email).perform()
            sleep(0.5)
            totalAtch = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, "//*[@id='message-htmlpart1']/div/div/div[2]/table[2]/tbody/tr")))
            mapaAtch = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="message-htmlpart1"]/div/div/div[2]/div[2]/table/tbody/tr')))
            hTabela = len(self.driver.find_elements(By.XPATH,"//*[@id='message-htmlpart1']/div/div/div[2]/table[2]/tbody/tr"))
            hMapa = len(self.driver.find_elements(By.XPATH,'//*[@id="message-htmlpart1"]/div/div/div[2]/div[2]/table/tbody/tr'))
            self.matriz = self.matriz + [linha] * (hTabela)
            rangeTabela = len(self.matriz) - hTabelaAnterior
            lt = 1
            transpAnterior = ''   
            print(f'{hTabela} NFs para ser tabeladas...')
            for l in range(rangeTabela):
                tabelaTransporte = []
                nf_Tabela = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, f"//*[@id='message-htmlpart1']/div/div/div[2]/table[2]/tbody/tr[{lt}]/td[2]")))
                transp_Tabela = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, f"//*[@id='message-htmlpart1']/div/div/div[2]/table[2]/tbody/tr[{lt}]/td[5]")))
                nf = nf_Tabela.accessible_name
                transp = transp_Tabela.accessible_name
                nt = 3
                if transpAnterior == transp:
                    tabelaTransporte.append(str(data_tabela))
                    tabelaTransporte.append(str(carro_tabela))
                else:
                    for n in range(hMapa):
                        carga_tabela = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, f"//*[@id='message-htmlpart1']/div/div/div[2]/div[2]/table[1]/tbody/tr[{nt}]/td[2]"))).accessible_name
                        carro_tabela = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, f"//*[@id='message-htmlpart1']/div/div/div[2]/div[2]/table[1]/tbody/tr[{nt}]/td[9]"))).accessible_name
                        data_tabela = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, f"//*[@id='message-htmlpart1']/div/div/div[2]/div[2]/table[1]/tbody/tr[{nt}]/td[11]"))).accessible_name
                        try:
                            numCarga = int(carga_tabela)
                        except: 
                            numCarga = 0
                        try:
                            numTransp = int(transp)
                        except:
                            numTransp = 0
                        if numCarga !=0 and numTransp !=0:
                            if numCarga == numTransp:
                                tabelaTransporte.append(str(data_tabela))
                                tabelaTransporte.append(str(carro_tabela))
                                break
                        elif n + 1 == hMapa or transp == 'Transporte':
                            tabelaTransporte.append(str("Data"))
                            tabelaTransporte.append(str("Carro"))
                            break
                        nt = nt +1
                tabelaTransporte.append(str(transp)) 
                tabelaTransporte.append(str(nf))
                tabelaTransporte.append(str("Chave"))
                # tabelaTransporte.append(str(nf + "  - {},{}").format(l + hTabelaAnterior,0))
                # tabelaTransporte.append(str(transp + " - {},{}").format(l + hTabelaAnterior,1)) 
                # tabelaTransporte.append(str("Chave" + " - {},{}").format(l + hTabelaAnterior,2))
                # tabelaTransporte.append(str("Chave" + " [{}][{}]").format(l + hTabelaAnterior,4))
                self.matriz[l + hTabelaAnterior] = tabelaTransporte
                transpAnterior = transp
                lt = lt + 1
            # print(self.matriz)
            hTabelaAnterior = hTabela
            self.driver.back()
            mostrarMsg = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//a[@title="Mostrar mensagens não lidas"]')))
            menu = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="mailsearchform"]')))
            menu.send_keys('cacau')
            menu.send_keys(Keys.ENTER)
            try:
                esperarEmail = wdwToEmail.until(condicaoEsperada.element_to_be_clickable((By.CLASS_NAME, 'rcmContactAddress')))
            except:
                pass
            sleep(0.5)
            tx = tx +1

        if tx != 0 and tx > 1:
            print(f'{tx} emails extraidos com sucesso')
        elif tx == 1:
            print(f'{tx} email extraido com sucesso')

    def consultaChave(self, driver):
        # print(self.matriz)
        # mostrar = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="rcmliSU5CT1guVHJhc2g"]/a')))
        operacional = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div[1]/ul[2]/li[3]/a')))
        operacional.click()
        notasFiscais = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="invoice_base"]/a/span')))
        notasFiscais.click()
        consultaNf = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="invoice_base"]/ul/li[1]')))
        consultaNf.click()
        sleep(1)

        tituloConsulta = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div[2]/div/div/div/div/div[1]/div/div/div/span')))
        data = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="search_invoices_issue_date"]')))
        data.click()
        limparData = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/div[10]/div[3]/div/button[2]')))
        limparData.click()
        
        hList = len(self.matriz)

        if hList == 0:
            pass
        else:
            for l in range(hList):
                procurarNF = self.driver.find_element(By.ID, 'search_invoices_number')
                try:
                    nfList = int(self.matriz[l][3])
                except:
                    nfList = 0
                procurarNF.send_keys(nfList)
                procurarNF.send_keys(Keys.ENTER)
                # tableChave = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="invoices_view_table"]/div[3]/div/table/tbody[1]/tr[1]')))
                # exportExcel = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="btn-export-xlsx"]/i')))
                op = True
                wdwChave = WebDriverWait(self.driver, 5)
                while op:
                    try:
                        msgNfConsulta = wdwChave.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="invoices_view_table"]/div[3]/div/table/tbody[1]/tr[1]/td[3]/div')))
                        numNf = int(msgNfConsulta.accessible_name)
                    except:
                        numNf = 0
                    try:
                        msgNf = wdwChave.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="invoices_view_table"]/div[3]/div/table/tbody[1]/tr[1]/td/div'))).text
                    except:
                        msgNf = ''
                    if (numNf == nfList and numNf != 0) or msgNf != '':
                        op = False
                    else:
                        msgNf = ''
                        numNf = 0
                
                if msgNf != 'Nenhuma nota localizada.':
                    chaveIcon = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="invoices_view_table"]/div[3]/div/table/tbody[1]/tr[1]/td[13]/div/i')))
                    chaveIcon.click()
                    okBtn = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/div[11]/div/div[3]/button[1]')))
                    chaveAtual = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="swal2-content"]'))).text
                    self.matriz[l][4] = chaveAtual
                    okBtn.click()
                    procurarNF.clear()
                else:
                    self.matriz[l][4] = "Chave"
                    procurarNF.clear()
                msgNf = ''
                numNf = 0

        # print(self.matriz)
        logger_file = open('./log/logger.log', "a", encoding='utf-8')
        hm = len(self.matriz)
        for l in range(hm):
            stringFile = str(self.matriz[l])
            logger_file.write(stringFile + '\n')
        logger_file.close()       

        if hList != 0 and hList > 1:
            print(f'{hList} chaves consultadas com sucesso')
        elif hList == 1:
            print(f'{hList} chave consultada com sucesso')

    def createManifest(self, driver):
        globalizado = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="tab-main-form"]/div/div[5]/div/div/div/div[2]/div[8]/div/div/span')))
        globalizado.click()
        pagador = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="select2-freight_normal_payer_id-container"]')))
        pagador.click()
        sleep(1)

        filtro = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/span/span/span[1]/input')))
        filtro.send_keys('i.b.a.c')
        op = True
        while op:
            cacau = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="select2-freight_normal_payer_id-results"]/li/div/div[1]')))
            nameCacau = cacau.text
            if nameCacau != '':
                cacau.click()
                op = False

        btnConfirm = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="new_freight_normal"]/div[2]/button')))
        btnConfirm.click()

        btnCloseMsg = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="toast-container"]/div/button')))
        btnCloseMsg.click()

        btnChave = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="tab-main-form"]/div/div[2]/div[1]/div[5]/button')))
        btnChave.click()

        btnInfo = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_modal"]/div/div/form/div/div[1]/div/div/div[2]/div/div/span')))
        btnInfo.click()

        logger_file = open("./log/logger.log", "r+")
        linhas = logger_file.readlines()

        hList = len(linhas)

        if hList == 0:
            pass
        else:
            for l in range(hList):
                inputChave = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_form_key"]')))
                try:
                    matriz = linhas[l].split(",")
                    if matriz[4].replace("'","").replace("]\n","").strip() == 'Chave':
                        pass
                    else:
                        chaveList = int(matriz[4].replace("'","").replace("]\n","").strip())
                        transpList = int(matriz[2].replace("'","").strip())
                        
                        inputChave.send_keys(chaveList)
                        inputChave.send_keys(Keys.ENTER)
                        btnLer = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_modal"]/div/div/form/div/div[2]/div[1]/div[2]/button')))
                        btnLer.click()

                        btnWait = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_documents_table"]/div/div/table/tbody[2]/tr/td/div/span[2]/button[3]')))

                        proxMatriz = linhas[l + 1].split(",")
                        proxTranspList = int(proxMatriz[2].replace("'","").strip())

                        if transpList == proxTranspList:
                            pass
                        else:
                            closeChave = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_modal"]/div/div/div[2]/a')))
                            closeChave.click()

                            body= self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '/html/body')))
                            body.send_keys(Keys.PAGE_DOWN)

                            tipoVeiculo = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="select2-freight_normal_vehicle_type_id-container"]')))
                            tipoVeiculo.click()
                            
                            inputVeiculo = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/span/span/span[1]/input')))

                            if transpList == 'VUC':
                                newTranspList = 'IVECO'
                            elif transpList == 'HR':
                                newTranspList = 'HR'
                            else:
                                newTranspList = 'IVECO'
                            
                            inputVeiculo.send_keys(newTranspList)
                            op = True
                            while op:
                                choseVeiculo = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="select2-freight_normal_vehicle_type_id-results"]/li')))
                                try:
                                    choseVeiculo.click()
                                    op = False
                                except:
                                    pass
                                
                            saveFrete = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="submitFreight"]')))
                            saveFrete.click()

                            saveYes = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/div[14]/div/div[3]/button[1]')))
                            #saveYes.click()
                            body.send_keys(Keys.PAGE_UP)

                            newFrete = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="new-crud-btn"]')))
                            newFrete.click()

                            globalizado = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="tab-main-form"]/div/div[5]/div/div/div/div[2]/div[8]/div/div/span')))
                            globalizado.click()
                            pagador = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="select2-freight_normal_payer_id-container"]')))
                            pagador.click()
                            sleep(1)

                            filtro = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/span/span/span[1]/input')))
                            filtro.send_keys('i.b.a.c')
                            op = True
                            while op:
                                cacau = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="select2-freight_normal_payer_id-results"]/li/div/div[1]')))
                                nameCacau = cacau.text
                                if nameCacau != '':
                                    cacau.click()
                                    op = False
                            btnConfirm = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="new_freight_normal"]/div[2]/button')))
                            btnConfirm.click()

                            btnCloseMsg = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="toast-container"]/div/button')))
                            btnCloseMsg.click()

                            btnChave = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="tab-main-form"]/div/div[2]/div[1]/div[5]/button')))
                            btnChave.click()

                            btnInfo = self.wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_modal"]/div/div/form/div/div[1]/div/div/div[2]/div/div/span')))
                            btnInfo.click()       
                except:
                    pass

extrair_dados = Dados()
extrair_dados.login_email()

sys.exit()