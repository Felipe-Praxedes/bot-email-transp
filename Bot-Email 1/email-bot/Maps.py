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

Origem = 'Jundiaí - Aglomeração Urbana de Jundiaí, Jundiaí - SP'
Destino= 'Carapicuíba, São Paulo'

urlFretes = 'https://www.google.com/maps/'

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

        btnRotas = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="xoLGzf-T3iPGc"]')))
        btnRotas.click()

        btnOrigem = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="sb_ifc51"]/input')))
        btnOrigem.send_keys(Origem)

        btnDestino = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="sb_ifc52"]/input')))
        btnDestino.send_keys(Destino)

        btnDestino.send_keys(Keys.ENTER)

        textoKm = wdw.until(condicaoEsperada.element_to_be_clickable((By.XPATH, '//*[@id="section-directions-trip-0"]/div/div[1]/div[1]/div[2]/div')))
        km  = textoKm.text

        print(km)
            
dados.createManifest()

sys.exit()