import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from navegador_utils import clicar_menu_sadt
import sys
import time
import threading
import openpyxl
from openpyxl.styles import PatternFill, Font


def finalizar_solicitacao_tratando_erros(navegador):
    try:
        print("[FINALIZAR] Clicando no botão 'Finalizar'...")
        botao_finalizar = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.ID, "Button_Finalizar"))
        )
        botao_finalizar.click()
        time.sleep(1)

        # Verifica se apareceu mensagem de erro
        try:
            explica_hidden = WebDriverWait(navegador, 5).until(
                EC.presence_of_element_located((By.NAME, "explica_bloq_hidden"))
            )
            texto_erro = explica_hidden.get_attribute("value")
            if "erros" in texto_erro.lower():
                print("[FINALIZAR] Erros detectados na solicitação.")
                radio_sim = navegador.find_element(By.XPATH, '//input[@name="forcar_solic" and @value="1"]')
                radio_sim.click()
                time.sleep(1)
                print("[FINALIZAR] Marcado 'Sim' para finalizar mesmo com erros, clicando novamente em 'Finalizar'...")
                botao_finalizar = WebDriverWait(navegador, 10).until(
                    EC.element_to_be_clickable((By.ID, "Button_Finalizar"))
                )
                botao_finalizar.click()
                time.sleep(1.5)
        except TimeoutException:
            print("[FINALIZAR] Nenhum erro detectado após clicar em finalizar.")

        # Verifica se mensagem "Guia em estudo" aparece
        try:
            WebDriverWait(navegador, 5).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//td[contains(text(), "Em estudo") or contains(text(), "em estudo")]')
                )
            )
            print("[FINALIZAR] Solicitação finalizada com sucesso (mensagem 'em estudo' detectada).")
            if not clicar_menu_sadt(navegador):
                print("[ERRO] Não conseguiu voltar ao menu SADT.")
                return "falha", False
            return "ok", True

        except TimeoutException:
            print("[FINALIZAR] Não encontrou mensagem 'em estudo'. Tentando voltar ao menu SADT mesmo assim...")
            if not clicar_menu_sadt(navegador):
                print("[ERRO] Não conseguiu voltar ao menu SADT.")
                return "falha", False
            return "erro_envio", True

    except Exception as e:
        print(f"[FINALIZAR] Erro ao tentar finalizar: {e}")
        return "falha", False