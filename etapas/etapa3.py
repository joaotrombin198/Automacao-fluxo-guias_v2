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

def etapa_3_preencher_unimed_e_contratado(navegador):
    try:
        print("[ETAPA 3] Aguardando campo 'Unimed Executora'...")
        unimed_exec = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.ID, "cd_unimed_executora"))
        )
        unimed_exec.clear()
        unimed_exec.send_keys("0198")
        print("[ETAPA 3] Campo 'Unimed Executora' preenchido com 0198.")
        time.sleep(0.5)
    except TimeoutException:
        print("[ERRO - ETAPA 3] Campo 'Unimed Executora' não encontrado.")
        return False

    try:
        print("[ETAPA 3] Clicando na lupa para abrir popup do Contratado Solicitante...")
        lupa_prest_solic = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.localizadorImgLink"))
        )
        lupa_prest_solic.click()
        print("[ETAPA 3] Popup de busca do contratado aberto.")
        time.sleep(0.8)
    except TimeoutException:
        print("[ERRO - ETAPA 3] Lupa do Contratado Solicitante não encontrada ou não clicável.")
        return False

    main_window = navegador.current_window_handle
    WebDriverWait(navegador, 10).until(EC.number_of_windows_to_be(2))
    for handle in navegador.window_handles:
        if handle != main_window:
            popup = handle
            break
    navegador.switch_to.window(popup)
    print("[ETAPA 3] Troca para janela popup realizada.")
    time.sleep(0.8)

    try:
        print("[ETAPA 3] Preenchendo código da Unimed na popup...")
        cd_unimed_prestador = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.ID, "s_CD_UNIMED_PRESTADOR"))
        )
        cd_unimed_prestador.clear()
        cd_unimed_prestador.send_keys("0198")
        time.sleep(0.3)

        print("[ETAPA 3] Preenchendo nome do prestador na popup...")
        nome_prestador = navegador.find_element(By.ID, "s_NM_COMPLETO")
        nome_prestador.clear()
        nome_prestador.send_keys("MICHEL FARACO")
        time.sleep(0.3)

        print("[ETAPA 3] Clicando em localizar na popup...")
        botao_localizar = navegador.find_element(By.NAME, "Button_DoSearch")
        botao_localizar.click()
        time.sleep(0.3)
    except TimeoutException:
        print("[ERRO - ETAPA 3] Campos ou botão da popup não encontrados.")
        return False

    try:
        print("[ETAPA 3] Aguardando tabela com Michel Faraco...")
        link_michel = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "MICHEL FARACO")]'))
        )
        link_michel.click()
        print("[ETAPA 3] Michel Faraco selecionado.")
        time.sleep(0.8)
    except TimeoutException:
        print("[ERRO - ETAPA 3] Link do Michel Faraco não encontrado na tabela.")
        return False

    # Voltar para janela principal
    navegador.switch_to.window(main_window)
    print("[ETAPA 3] Voltou para janela principal após seleção do contratado.")
    time.sleep(0.5)

    # Garante que está no frame correto antes de clicar no avançar
    try:
        navegador.switch_to.default_content()
        WebDriverWait(navegador, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "corpo")))
    except TimeoutException:
        print("[ERRO - ETAPA 3] Frame 'corpo' não encontrado para clicar no avançar.")
        return False

    # Tenta clicar no botão avançar com múltiplas abordagens
    try:
        print("[ETAPA 3] Tentando clicar no botão 'Avançar' da etapa 3...")

        # Tenta pelo id
        try:
            btn_avancar = WebDriverWait(navegador, 5).until(
                EC.element_to_be_clickable((By.ID, "botao_avancar"))
            )
            btn_avancar.click()
            print("[ETAPA 3] Botão 'Avançar' clicado via ID com sucesso.")
        except TimeoutException:
            # Tenta pelo atributo value (texto do botão)
            try:
                btn_avancar = WebDriverWait(navegador, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//input[@type="submit" and contains(@value, "Avançar")]'))
                )
                btn_avancar.click()
                print("[ETAPA 3] Botão 'Avançar' clicado via XPATH com sucesso.")
            except TimeoutException:
                # Fallback: clicar via javascript
                print("[ETAPA 3] Tentando clicar no botão 'Avançar' via JavaScript...")
                btn = navegador.find_element(By.XPATH, '//input[@type="submit" and contains(@value, "Avançar")]')
                navegador.execute_script("arguments[0].click();", btn)
                print("[ETAPA 3] Botão 'Avançar' clicado via JavaScript.")
        time.sleep(0.5)
    except Exception as e:
        print(f"[ERRO - ETAPA 3] Falha ao clicar no botão 'Avançar': {e}")
        return False

    return True