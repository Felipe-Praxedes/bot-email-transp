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

loginJw = 'cte.marba@jotaw.com'
passwordJw = 'ctemarba'

urlFretes = 'https://jotaw.eslcloud.com.br/freight/normals/new'

class dados:
    def createManifest():
        options = Options()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_experimental_option(
            'prefs', {
            "download.Prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        options.add_argument("--start-maximized")

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        sleep(0.5)
        wdw = WebDriverWait(driver, 30)
        driver.get(urlFretes)
        op = True
        while op:
            userJw = wdw.until(condicaoEsperada.element_to_be_clickable((By.ID, 'user_email')))
            userPassword = wdw.until(condicaoEsperada.element_to_be_clickable((By.ID, 'user_password')))
            btn_login = driver.find_element(By.XPATH,'//*[@id="new_user"]/div[2]/div/div[1]/input')
            userJw.send_keys(loginJw)
            print(f'Email {loginJw} utilizado')
            userPassword.send_keys(passwordJw)
            btn_login.click()
            op = False
        driver.get(urlFretes)

        globalizado = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="tab-main-form"]/div/div[5]/div/div/div/div[2]/div[8]/div/div/span')))
        globalizado.click()
        pagador = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="select2-freight_normal_payer_id-container"]')))
        pagador.click()
        sleep(1)

        filtro = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/span/span/span[1]/input')))
        filtro.send_keys('i.b.a.c')
        op = True
        while op:
            cacau = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="select2-freight_normal_payer_id-results"]/li/div/div[1]')))
            nameCacau = cacau.text
            if nameCacau != '':
                cacau.click()
                op = False

        btnConfirm = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="new_freight_normal"]/div[2]/button')))
        btnConfirm.click()

        btnCloseMsg = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="toast-container"]/div/button')))
        btnCloseMsg.click()

        btnChave = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="tab-main-form"]/div/div[2]/div[1]/div[5]/button')))
        btnChave.click()

        btnInfo = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_modal"]/div/div/form/div/div[1]/div/div/div[2]/div/div/span')))
        btnInfo.click()

        logger_file = open("./log/logger.log", "r+")
        linhas = logger_file.readlines()

        hList = len(linhas)

        if hList == 0:
            pass
        else:
            for l in range(hList):
                inputChave = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_form_key"]')))
                try:
                    matriz = linhas[l].split(",")
                    if matriz[4].replace("'","").replace("]\n","").strip() == 'Chave':
                        pass
                    else:
                        chaveList = int(matriz[4].replace("'","").replace("]\n","").strip())
                        transpList = int(matriz[2].replace("'","").strip())
                        
                        inputChave.send_keys(chaveList)
                        inputChave.send_keys(Keys.ENTER)
                        btnLer = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_modal"]/div/div/form/div/div[2]/div[1]/div[2]/button')))
                        btnLer.click()

                        btnWait = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_documents_table"]/div/div/table/tbody[2]/tr/td/div/span[2]/button[3]')))

                        proxMatriz = linhas[l + 1].split(",")
                        proxTranspList = int(proxMatriz[2].replace("'","").strip())

                        if transpList == proxTranspList:
                            pass
                        else:
                            closeChave = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_modal"]/div/div/div[2]/a')))
                            closeChave.click()

                            body= wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '/html/body')))
                            body.send_keys(Keys.PAGE_DOWN)

                            tipoVeiculo = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="select2-freight_normal_vehicle_type_id-container"]')))
                            tipoVeiculo.click()
                            
                            inputVeiculo = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/span/span/span[1]/input')))

                            if transpList == 'VUC':
                                newTranspList = 'IVECO'
                            elif transpList == 'HR':
                                newTranspList = 'HR'
                            else:
                                newTranspList = 'IVECO'
                            
                            inputVeiculo.send_keys(newTranspList)
                            op = True
                            while op:
                                choseVeiculo = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="select2-freight_normal_vehicle_type_id-results"]/li')))
                                try:
                                    choseVeiculo.click()
                                    op = False
                                except:
                                    pass
                                
                            saveFrete = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="submitFreight"]')))
                            saveFrete.click()

                            saveYes = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/div[14]/div/div[3]/button[1]')))
                            #saveYes.click()
                            body.send_keys(Keys.PAGE_UP)

                            newFrete = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="new-crud-btn"]')))
                            newFrete.click()

                            globalizado = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="tab-main-form"]/div/div[5]/div/div/div/div[2]/div[8]/div/div/span')))
                            globalizado.click()
                            pagador = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="select2-freight_normal_payer_id-container"]')))
                            pagador.click()
                            sleep(1)

                            filtro = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/span/span/span[1]/input')))
                            filtro.send_keys('i.b.a.c')
                            op = True
                            while op:
                                cacau = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="select2-freight_normal_payer_id-results"]/li/div/div[1]')))
                                nameCacau = cacau.text
                                if nameCacau != '':
                                    cacau.click()
                                    op = False
                            btnConfirm = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="new_freight_normal"]/div[2]/button')))
                            btnConfirm.click()

                            btnCloseMsg = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="toast-container"]/div/button')))
                            btnCloseMsg.click()

                            btnChave = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="tab-main-form"]/div/div[2]/div[1]/div[5]/button')))
                            btnChave.click()

                            btnInfo = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@id="quick_document_modal"]/div/div/form/div/div[1]/div/div/div[2]/div/div/span')))
                            btnInfo.click()       
                except:
                    pass
            
dados.createManifest()

sys.exit()