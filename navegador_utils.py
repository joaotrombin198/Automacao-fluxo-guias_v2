# navegador_utils.py
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import sys
import time
import threading
import openpyxl
from openpyxl.styles import PatternFill, Font

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def iniciar_navegador():
    navegador = webdriver.Chrome()
    navegador.maximize_window()
    navegador.get("https://rda-hml.unimedsc.com.br/autsc2/Login.do")
    print("[INFO] Navegador iniciado.")
    return navegador

def fazer_login(navegador, usuario, senha):
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, "ds_login"))).send_keys(usuario)
    navegador.find_element(By.ID, "passwordTemp").send_keys(senha)
    WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.ID, "Button_DoLogin"))).click()
    print("[INFO] Login realizado com sucesso!")

def clicar_menu_sadt(navegador):
    try:
        navegador.switch_to.default_content()
        WebDriverWait(navegador, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, "/html/body/div/div[1]/iframe"))
        )
        botao_sadt = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/ul/li[2]/a/div/div[2]'))
        )
        botao_sadt.click()
        print("[INFO] Menu SADT acessado para resetar tela.")
        time.sleep(0.6)
        return True
    except Exception as e:
        print(f"[ERRO] Falha ao acessar menu SADT: {e}")
        return False