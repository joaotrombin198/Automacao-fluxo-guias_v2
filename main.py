# main.py

from settings import CAMINHO_PLANILHA, CARTEIRINHA_FIXA
from planilha_utils import carregar_planilha, salvar_planilha_formatada
from navegador_utils import iniciar_navegador, fazer_login, clicar_menu_sadt
from carteira_utils import ajustar_carteira, validar_carteirinha, extrair_segmentos_carteira
from etapas.etapa3 import etapa_3_preencher_unimed_e_contratado
from etapas.etapa4 import etapa_4_preencher_campos
from etapas.finalizar import finalizar_solicitacao_tratando_erros

import pandas as pd
import time
import threading
import sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def iniciar_cronometro():
    start = time.time()
    def mostrar_tempo():
        while True:
            time.sleep(30)
            elapsed = time.time() - start
            print(f"[TEMPO] Tempo de execução: {int(elapsed)} segundos")
    t = threading.Thread(target=mostrar_tempo, daemon=True)
    t.start()


if __name__ == "__main__":
    iniciar_cronometro()

    try:
        df = pd.read_csv(CAMINHO_PLANILHA, sep=';', dtype={"NR_SEQ_SEGURADO": str, "NR_SEQ_REQUISICAO": str})
        df["NR_SEQ_SEGURADO"] = df["NR_SEQ_SEGURADO"].str.strip()
        df["NR_SEQ_REQUISICAO"] = df["NR_SEQ_REQUISICAO"].str.strip()
        print("[INFO] Planilha CSV carregada com sucesso.")
    except Exception as e:
        print(f"[ERRO] Erro ao carregar planilha: {e}")
        sys.exit(1)

    grupos = df.groupby(["NR_SEQ_REQUISICAO", "NR_SEQ_SEGURADO"], sort=False)
    lista_grupos = list(grupos)
    total_grupos = len(lista_grupos)

    navegador = iniciar_navegador()
    try:
        fazer_login(navegador, "admin198", "Unimed198@")
    except Exception as e:
        print(f"[ERRO] Falha no login: {e}")
        navegador.quit()
        sys.exit(1)

    if not clicar_menu_sadt(navegador):
        navegador.quit()
        sys.exit(1)

    for idx, ((nr_seq_req, nr_seq_seg), grupo) in enumerate(lista_grupos, start=1):
        print(f"\n[GRUPO {idx}/{total_grupos}] Guia {nr_seq_req} / Segurado {nr_seq_seg} → {len(grupo)} itens")

        carteira = ajustar_carteira(nr_seq_seg)
        if not validar_carteirinha(carteira):
            print(f"[AVISO] Carteira inválida: {carteira}. Pulando guia.")
            continue

        unimed, cartao, benef, depen, digitos = extrair_segmentos_carteira(carteira)
        navegador.switch_to.default_content()
        WebDriverWait(navegador, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "corpo")))
        navegador.find_element(By.ID, "bnf_cd_unimed").clear(); navegador.find_element(By.ID, "bnf_cd_unimed").send_keys(unimed)
        navegador.find_element(By.ID, "bnf_cd_cartao").clear(); navegador.find_element(By.ID, "bnf_cd_cartao").send_keys(cartao)
        navegador.find_element(By.ID, "bnf_cd_benef").clear(); navegador.find_element(By.ID, "bnf_cd_benef").send_keys(benef)
        navegador.find_element(By.ID, "bnf_cd_dependencia").clear(); navegador.find_element(By.ID, "bnf_cd_dependencia").send_keys(depen)

        sucesso_digito = False
        for digito in digitos:
            print(f"[GRUPO {idx}] [DIGITO] Tentando dígito: {digito}")
            navegador.find_element(By.ID, "bnf_cd_digito_verificador").clear()
            navegador.find_element(By.ID, "bnf_cd_digito_verificador").send_keys(str(digito))
            time.sleep(0.6)
            try:
                WebDriverWait(navegador, 1).until(EC.presence_of_element_located((By.XPATH, '//*[contains(text(), "2º Passo") or contains(text(), "2° Passo")]')))
                print(f"[GRUPO {idx}] [DIGITO] Dígito {digito} aceito!")
                sucesso_digito = True

                # Verifica se há exclusão
                try:
                    data_exclusao_element = WebDriverWait(navegador, 5).until(
                        EC.presence_of_element_located((By.XPATH, '//td[contains(text(), "Data Exclusão:")]/following-sibling::td[1]'))
                    )
                    data_exclusao_texto = data_exclusao_element.text.strip().replace('\xa0', '')
                    print(f"[GRUPO {idx}] [INFO] Data de Exclusão: '{data_exclusao_texto}'")
                    if data_exclusao_texto != "__/__/____":
                        print(f"[GRUPO {idx}] [INFO] Beneficiário excluído. Usando carteirinha fixa.")
                        carteira = ajustar_carteira(CARTEIRINHA_FIXA)
                        unimed, cartao, benef, depen, digitos = extrair_segmentos_carteira(carteira)

                        if not clicar_menu_sadt(navegador): break
                        navegador.switch_to.default_content()
                        WebDriverWait(navegador, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "corpo")))
                        navegador.find_element(By.ID, "bnf_cd_unimed").clear(); navegador.find_element(By.ID, "bnf_cd_unimed").send_keys(unimed)
                        navegador.find_element(By.ID, "bnf_cd_cartao").clear(); navegador.find_element(By.ID, "bnf_cd_cartao").send_keys(cartao)
                        navegador.find_element(By.ID, "bnf_cd_benef").clear(); navegador.find_element(By.ID, "bnf_cd_benef").send_keys(benef)
                        navegador.find_element(By.ID, "bnf_cd_dependencia").clear(); navegador.find_element(By.ID, "bnf_cd_dependencia").send_keys(depen)

                        sucesso_digito = False
                        for digito in digitos:
                            print(f"[GRUPO {idx}] [FIXA] Tentando dígito: {digito}")
                            navegador.find_element(By.ID, "bnf_cd_digito_verificador").clear()
                            navegador.find_element(By.ID, "bnf_cd_digito_verificador").send_keys(str(digito))
                            time.sleep(0.6)
                            try:
                                WebDriverWait(navegador, 1).until(EC.presence_of_element_located((By.XPATH, '//*[contains(text(), "2º Passo") or contains(text(), "2° Passo")]')))
                                print(f"[GRUPO {idx}] [FIXA] Dígito {digito} aceito!")
                                sucesso_digito = True
                                break
                            except: continue
                        if not sucesso_digito:
                            print(f"[GRUPO {idx}] [ERRO] Nenhum dígito válido para carteirinha fixa. Pulando.")
                            continue
                except TimeoutException:
                    print(f"[GRUPO {idx}] [INFO] Nenhum campo de exclusão. Prosseguindo com carteira normal.")
                break
            except TimeoutException:
                print(f"[GRUPO {idx}] [DIGITO] Dígito {digito} rejeitado.")

        if not sucesso_digito:
            print(f"[GRUPO {idx}] [ERRO] Nenhum dígito válido. Pulando guia.")
            continue

        navegador.find_element(By.ID, "Button_Search").click()
        time.sleep(0.6)

        if not etapa_3_preencher_unimed_e_contratado(navegador):
            print(f"[GRUPO {idx}] [ERRO] Falha na Etapa 3. Pulando guia.")
            continue

        sucesso_et4 = etapa_4_preencher_campos(navegador, grupo)
        if not sucesso_et4:
            print(f"[GRUPO {idx}] [ERRO] Falha na Etapa 4. Pulando guia.")
            continue

        status, ok = finalizar_solicitacao_tratando_erros(navegador)
        print(f"[GRUPO {idx}] [FINALIZAR] Guia {nr_seq_req} → {status}")

        df.loc[(df.NR_SEQ_REQUISICAO == nr_seq_req) & (df.NR_SEQ_SEGURADO == nr_seq_seg), 'STATUS'] = status
        salvar_planilha_formatada(df, CAMINHO_PLANILHA.replace('.csv', '_atualizado.xlsx'))

        clicar_menu_sadt(navegador)

    print("\n[INFO] Todas as guias processadas.")
    navegador.quit()