# carteira_utils.py
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

from settings import CARTEIRINHA_FIXA

def ajustar_carteira(carteira):
    carteira = carteira.strip()
    if carteira.startswith("198") and len(carteira) == 15:
        carteira = "0" + carteira
    return carteira

def validar_carteirinha(carteira):
    carteira = carteira.strip()
    return carteira == CARTEIRINHA_FIXA or (carteira.startswith("0198") and len(carteira) == 16)

def extrair_segmentos_carteira(carteira):
    if carteira == CARTEIRINHA_FIXA:
        unimed = carteira[0:4]
        cartao = carteira[4:8]
        benef = carteira[8:14]
        depen = carteira[14:16]
        digitos = [carteira[16]]
    else:
        unimed = carteira[0:4]
        cartao = carteira[4:8]
        benef = carteira[8:14]
        depen = carteira[14:16]
        digitos = list(range(10))
    return unimed, cartao, benef, depen, digitos
