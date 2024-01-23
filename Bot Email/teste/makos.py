from functools import partial
from tokenize import String
from webbrowser import get
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as condicaoEsperada
from datetime import datetime
from time import sleep
from mimetypes import init
import os
import sys
import pathlib
import shutil

try:
    os.mkdir(os.getcwd() + '\\Baixados2')
    os.mkdir('./Baixados23')
except OSError:
    pass
