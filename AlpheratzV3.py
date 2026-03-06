import os
import time
from datetime import datetime

import pandas as pd
import pyautogui

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException


URL_BASE = ""
LEGAL_URL = ""

USUARIO = ""
SENHA = ""

TIMEOUT_LONG = 60

SEARCH_INPUT_ID = "ctl00_content_txt_search"
WORKFLOW_TAB_ID = "ctl00_content_bt_workflow"
WF_STEP_INPUT_ID = "ctl00_content_txt_ActiveLegalWorkflowStep"
WF_ADD_BTN_ID = "ctl00_content_btn_AddActiveLegalWorkflowStep"

RI_SAVE_ID = "ctl00_content_bt_step_save"
RI_DATE_ID = "ctl00_content_tb1"
RI_OBS_ID = "ctl00_content_tb2"

UPLOAD_BTN_CSS = "#table_uploads > tbody > tr:nth-child(1) > td > input:nth-child(1)"
UPLOAD_BTN_ALT_CSS = "#div_upload_save > div > div.qq-upload-button"
UPLOAD_BTN_ALT2_CSS = "div.qq-upload-button"
UPLOAD_LIST_CSS = "ul.qq-upload-list > li"

BASE_DOCS_DIR = ""

opt1 = "Litigation pendente apenas de habilitação"
opt2 = "Litigation pendente apenas de substituição/assistência"
opt3 = "Litigation pendente de habilitação e substituicao/assistência"

save = True


def log(msg: str):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def wait_visible(wait: WebDriverWait, by: By, value: str):
    return wait.until(EC.visibility_of_element_located((by, value)))


def wait_clickable(wait: WebDriverWait, by: By, value: str):
    return wait.until(EC.element_to_be_clickable((by, value)))


def safe_js_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    driver.execute_script("arguments[0].click();", el)


def clear_and_type(el, text: str):
    el.click()
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.DELETE)
    el.send_keys(str(text).strip())


def ensure_legal_search_ready(driver, timeout=TIMEOUT_LONG):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, SEARCH_INPUT_ID))
    )


def ensure_workflow_ready(driver, timeout=TIMEOUT_LONG):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, WF_STEP_INPUT_ID))
    )


def go_to_legal_search_page(driver, log_fn, timeout: int = TIMEOUT_LONG):
    log_fn("Navegando para a tela Legal...")
    driver.get(LEGAL_URL)
    log_fn("Aguardando campo de busca na tela Legal.")
    ensure_legal_search_ready(driver, timeout=timeout)
    log_fn("Tela Legal carregada.")


def attempt_login(driver, username: str, password: str, log_fn, timeout: int = 30) -> bool:
    if not username or not password:
        return False

    end_time = time.time() + timeout

    while time.time() < end_time:
        try:
            user_el = None
            pwd_el = None

            for u_id, p_id in (("txt_username", "txt_password"), ("txtUsuario", "txtSenha")):
                try:
                    user_el = WebDriverWait(driver, 1).until(
                        EC.presence_of_element_located((By.ID, u_id))
                    )
                    pwd_el = WebDriverWait(driver, 1).until(
                        EC.presence_of_element_located((By.ID, p_id))
                    )
                    break
                except Exception:
                    pass

            if user_el and pwd_el and user_el.is_displayed():
                clear_and_type(user_el, username)
                clear_and_type(pwd_el, password)
                pwd_el.send_keys(Keys.ENTER)
                log_fn("Login automático realizado.")
                return True

        except Exception:
            time.sleep(0.5)

    return False


def salvamento(driver, wait: WebDriverWait):
    try:
        btn_save = wait_clickable(wait, By.ID, RI_SAVE_ID)
        safe_js_click(driver, btn_save)
        log("Cliquei em Salvar.")
    except Exception:
        log("Erro no Salvamento, validar motivo")


def search_and_open_litigation(driver, wait: WebDriverWait, litigation_id: str):
    search = wait_visible(wait, By.ID, SEARCH_INPUT_ID)
    clear_and_type(search, litigation_id)
    search.send_keys(Keys.ENTER)

    candidates = [
        f"//span[normalize-space(text())='{litigation_id}']",
        f"//a[normalize-space(text())='{litigation_id}']",
        f"//*[self::a or self::span][contains(normalize-space(.),'{litigation_id}')][1]",
        f"//span[contains(@id,'lbl_OmniIndex_LitigationID') and normalize-space(text())='{litigation_id}']",
    ]

    for xp in candidates:
        try:
            el = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
            safe_js_click(driver, el)
            return
        except Exception:
            pass

    raise RuntimeError(f"Não consegui abrir o LitigationID={litigation_id}")


def open_legal_workflow_tab(driver, wait: WebDriverWait):
    btn = wait_clickable(wait, By.ID, WORKFLOW_TAB_ID)
    safe_js_click(driver, btn)
    ensure_workflow_ready(driver)


def resolve_step_name(alerta: str) -> str:
    a = (alerta or "").strip()
    if a == opt1:
        return "Petição - Habilitação Processual"
    if a == opt2 or a == opt3:
        return "Pedido de Substituição Processual Protocolado"
    return ""


def select_step_and_add(driver, wait: WebDriverWait, log_fn, step_name: str):
    step_input = wait_clickable(wait, By.ID, WF_STEP_INPUT_ID)

    driver.execute_script("arguments[0].focus();", step_input)
    clear_and_type(step_input, step_name)

    time.sleep(3)
    step_input.send_keys(Keys.ARROW_DOWN)
    time.sleep(1)
    step_input.send_keys(Keys.ENTER)

    add_btn = wait_clickable(wait, By.ID, WF_ADD_BTN_ID)
    safe_js_click(driver, add_btn)

    log_fn(f"Etapa adicionada: {step_name}")


def preencher_data_atual_no_passo(driver, log_fn):
    data_atual = datetime.now().strftime("%d/%m/%Y")
    campo_data = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, RI_DATE_ID))
    )

    clear_and_type(campo_data, data_atual)
    campo_data.send_keys(Keys.TAB)
    time.sleep(0.2)
    campo_data.send_keys(Keys.TAB)
    time.sleep(0.2)
    campo_data.send_keys(Keys.ENTER)

    log_fn(f"Data preenchida com o dia atual: {data_atual}")


def extrair_nome_arquivo(caminho: str) -> str:
    return os.path.basename(caminho)


def buscar_arquivo_na_pasta_litigation(litigation_id: str) -> list[str]:
    pasta_lit = os.path.join(BASE_DOCS_DIR, str(litigation_id))

    if not os.path.isdir(pasta_lit):
        log(f"não achei a pasta: {pasta_lit}")
        return []

    extensoes_validas = {
        ".jpeg", ".jpg", ".png", ".gif", ".tif", ".tiff", ".bmp",
        ".docx", ".doc", ".pdf", ".ppt",
        ".xls", ".xlsx", ".eml",
        ".zip", ".rar"
    }

    arquivos_validos = []

    for nome in os.listdir(pasta_lit):
        full = os.path.join(pasta_lit, nome)
        if os.path.isfile(full):
            ext = os.path.splitext(nome)[1].lower()
            if ext in extensoes_validas:
                arquivos_validos.append(full)

    arquivos_validos.sort(key=lambda x: os.path.basename(x).lower())

    if not arquivos_validos:
        log(f"não achei arquivo válido dentro da pasta: {pasta_lit}")
        return []

    return arquivos_validos


def upload_por_janela_windows(caminho_arquivo: str):
    time.sleep(1.2)
    pyautogui.write(caminho_arquivo)
    time.sleep(0.2)
    pyautogui.press("enter")


def contar_itens_upload(driver) -> int:
    return len(driver.find_elements(By.CSS_SELECTOR, UPLOAD_LIST_CSS))


def wait_novo_item_upload(driver, qtd_antes: int, timeout: int = 20):
    def _novo_item(_):
        itens = driver.find_elements(By.CSS_SELECTOR, UPLOAD_LIST_CSS)
        if len(itens) > qtd_antes:
            return len(itens) - 1
        return False

    try:
        return WebDriverWait(driver, timeout).until(_novo_item)
    except TimeoutException:
        return None


def wait_upload_final_state_by_index(driver, item_index: int, timeout: int = 60):
    def _state(_):
        itens = driver.find_elements(By.CSS_SELECTOR, UPLOAD_LIST_CSS)
        if item_index is None:
            return "unknown"
        if item_index >= len(itens):
            return None

        target = itens[item_index]
        cls = (target.get_attribute("class") or "").lower()

        if "qq-upload-fail" in cls:
            return "fail"

        if "qq-upload-success" in cls:
            return "success"

        spinners = target.find_elements(By.CSS_SELECTOR, "span.qq-upload-spinner")
        if not spinners:
            failed_text = target.find_elements(By.CSS_SELECTOR, "span.qq-upload-failed-text")
            if any(ft.is_displayed() for ft in failed_text):
                return "fail"
            return "success"

        return None

    return WebDriverWait(driver, timeout).until(_state)


def localizar_botao_upload(driver, timeout: int = 10):
    seletores = [
        UPLOAD_BTN_ALT_CSS,
        UPLOAD_BTN_ALT2_CSS,
        UPLOAD_BTN_CSS,
    ]

    ultimo_erro = None

    for css in seletores:
        try:
            el = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, css))
            )
            if el.is_displayed():
                log(f"Botão de upload encontrado pelo seletor: {css}")
                return el, css
        except Exception as e:
            ultimo_erro = e

    raise RuntimeError(f"Não consegui localizar o botão de upload. Último erro: {ultimo_erro}")


def clicar_botao_upload(driver):
    el, css = localizar_botao_upload(driver, timeout=5)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.3)

    try:
        driver.execute_script("arguments[0].click();", el)
        log(f"Cliquei no botão de upload usando: {css}")
        return
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(el).click(el).perform()
        log(f"Cliquei no botão de upload com ActionChains usando: {css}")
        return
    except Exception:
        pass

    try:
        input_file = el.find_element(By.CSS_SELECTOR, "input[type='file']")
        driver.execute_script("arguments[0].click();", input_file)
        log(f"Cliquei no input[type=file] interno usando: {css}")
        return
    except Exception:
        pass

    raise RuntimeError(f"Não consegui clicar no botão de upload usando o seletor: {css}")


def clicar_botao_carregar_novo_arquivo(driver, wait: WebDriverWait, lista: list[str]) -> int:
    enviados = 0
    total_arquivos = len(lista)

    log(f"Quantidade de arquivos desta pasta: {total_arquivos}")

    for idx, caminho_arquivo in enumerate(lista, start=1):
        nome_arquivo = extrair_nome_arquivo(caminho_arquivo)

        try:
            log(f"Upload {idx}/{total_arquivos} -> {caminho_arquivo}")

            qtd_antes = contar_itens_upload(driver)

            clicar_botao_upload(driver)

            tb2 = wait_visible(wait, By.ID, RI_OBS_ID)
            tb2.click()
            time.sleep(1)
            tb2.click()

            ActionChains(driver).send_keys(Keys.TAB).send_keys(Keys.ENTER).perform()
            time.sleep(2)

            upload_por_janela_windows(caminho_arquivo)
            log(f"Aguardando item novo na lista de upload: {nome_arquivo}")

            novo_indice = wait_novo_item_upload(driver, qtd_antes, timeout=5)

            if novo_indice is None:
                log(f"Nenhum item novo apareceu na lista para: {nome_arquivo}")
            else:
                log(f"Novo item identificado no índice {novo_indice}: {nome_arquivo}")

                try:
                    resultado = wait_upload_final_state_by_index(driver, novo_indice, timeout=60)
                    log(f"Upload finalizado: {resultado} | {nome_arquivo}")
                except TimeoutException:
                    log(f"Timeout no upload do arquivo: {nome_arquivo}")

            enviados += 1
            time.sleep(1)

        except Exception as e:
            log(f"Erro no upload do arquivo {nome_arquivo}: {e}")

    log(f"Fim dos uploads desta pasta. Enviados={enviados}")
    return enviados


def voltar_para_tela_inicial(driver):
    driver.get(LEGAL_URL)
    ensure_legal_search_ready(driver)


def preparar_dataframe_ordenado(df: pd.DataFrame, id_col: str, alerta_col: str) -> pd.DataFrame:
    df = df.copy()

    df[id_col] = df[id_col].astype(str).str.strip()
    df[alerta_col] = df[alerta_col].astype(str).str.strip()

    df = df[
        df[id_col].notna() &
        (df[id_col] != "") &
        (df[id_col].str.lower() != "nan")
    ].copy()

    df["_litigation_num"] = pd.to_numeric(df[id_col], errors="coerce")
    df = df[df["_litigation_num"].notna()].copy()

    df = df.sort_values(by="_litigation_num", ascending=True).reset_index(drop=True)

    return df


def main():
    excel_name = "SEU_ARQUIVO.xlsx"
    excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), excel_name)

    df = pd.read_excel(excel_path, dtype=str)

    id_col = "litigationID" if "litigationID" in df.columns else df.columns[0]
    alerta_col = "Alerta" if "Alerta" in df.columns else df.columns[17]

    df = preparar_dataframe_ordenado(df, id_col, alerta_col)

    log(f"Total de litigations válidos ordenados: {len(df)}")
    if len(df) > 0:
        log(f"Primeiro litigation da fila: {df.iloc[0][id_col]}")
        log(f"Último litigation da fila: {df.iloc[-1][id_col]}")

    options = webdriver.ChromeOptions()
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 40)

    driver.get(URL_BASE)

    pyautogui.PAUSE = 0.15
    pyautogui.FAILSAFE = True

    log("Tentando login automático...")
    attempt_login(driver, USUARIO, SENHA, log)

    go_to_legal_search_page(driver, log)

    total = len(df)
    processados = 0
    erros = 0
    sem_pasta = 0

    for i in range(total):
        litigation_id = str(df.iloc[i][id_col]).strip()
        alerta = str(df.iloc[i][alerta_col]).strip()

        if not litigation_id or litigation_id.lower() == "nan":
            continue

        step_name = resolve_step_name(alerta)
        if not step_name:
            continue

        log(f"({i+1}/{total}) Processando LitigationID={litigation_id}")

        try:
            ensure_legal_search_ready(driver)

            search_and_open_litigation(driver, wait, litigation_id)
            open_legal_workflow_tab(driver, wait)
            select_step_and_add(driver, wait, log, step_name)
            preencher_data_atual_no_passo(driver, log)

            lista = buscar_arquivo_na_pasta_litigation(litigation_id)

            if not lista:
                sem_pasta += 1
                log(f"LitigationID={litigation_id} sem pasta ou sem arquivo. Indo para o próximo.")
                processados += 1
                voltar_para_tela_inicial(driver)
                continue

            log(f"Arquivos encontrados para upload: {lista}")
            log(f"Quantidade de arquivos para este litigation: {len(lista)}")

            clicar_botao_carregar_novo_arquivo(driver, wait, lista)
            log("Uploads concluídos para este litigation.")

            if save:
                salvamento(driver, wait)

            processados += 1
            log(f"OK LitigationID={litigation_id}")

            voltar_para_tela_inicial(driver)

        except Exception as e:
            erros += 1
            log(f"ERRO LitigationID={litigation_id}: {e}")

            try:
                voltar_para_tela_inicial(driver)
            except Exception:
                pass

    log(f"Finalizado. Processados={processados} | Erros={erros} | Sem pasta/arquivo={sem_pasta}")


if __name__ == "__main__":
    main()