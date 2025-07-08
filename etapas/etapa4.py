# etapas/etapa4.py

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time


def etapa_4_preencher_campos(navegador, grupo_df):
    """
    Preenche todos os procedimentos de um grupo (mesmo NR_SEQ_REQUISICAO e NR_SEQ_SEGURADO).
    grupo_df: pandas.DataFrame com colunas CD_PROCEDIMENTO e QT_SOLICITADO.
    """
    try:
        # Seleciona "Não" para Atendimento a RN
        print("[ETAPA 4] Aguardando dropdown 'Atendimento a RN'...")
        select_rn = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.NAME, "FG_ATENDIMENTO_RN"))
        )
        for opt in select_rn.find_elements(By.TAG_NAME, 'option'):
            if opt.get_attribute("value") == "N":
                opt.click()
                print("[ETAPA 4] Selecionado 'Não' para Atendimento a RN.")
                break
        time.sleep(0.5)

        # Campo Observações com NR_SEQ_REQUISICAO (mesmo para todo grupo)
        nr_seq = grupo_df.iloc[0].NR_SEQ_REQUISICAO
        print(f"[ETAPA 4] Preenchendo Observações com {nr_seq}...")
        obs = navegador.find_element(By.ID, "ds_obs")
        obs.clear()
        obs.send_keys(str(nr_seq))
        obs.send_keys("\t")
        time.sleep(0.5)

        # Itera sobre cada procedimento no grupo
        for idx, row in enumerate(grupo_df.itertuples(), start=1):
            cd_id = f"cd_servico_{idx}"
            qt_id = f"nr_qtd_{idx}"
            proc = row.CD_PROCEDIMENTO
            qtd = str(row.QT_SOLICITADO)

            print(f"[ETAPA 4] Inserindo procedimento #{idx}: {proc} x {qtd}")
            # Preenche código do procedimento
            cd_el = navegador.find_element(By.ID, cd_id)
            cd_el.clear()
            cd_el.send_keys(proc)
            cd_el.send_keys("\t")
            time.sleep(0.5)

            # Verifica quantidade default na tela
            qt_el = navegador.find_element(By.ID, qt_id)
            qtd_tela = qt_el.get_attribute("value").strip()
            if qtd_tela != qtd:
                print(f"[ETAPA 4] Quantidade ({qtd}) difere da tela ({qtd_tela}). Atualizando...")
                qt_el.clear()
                qt_el.send_keys(qtd)
                qt_el.send_keys("\t")
                time.sleep(0.5)
            else:
                print(f"[ETAPA 4] Quantidade ({qtd}) confere com tela.")

        return True

    except TimeoutException:
        print("[ERRO - ETAPA 4] Elementos não encontrados para preencher múltiplos procedimentos.")
        return False
