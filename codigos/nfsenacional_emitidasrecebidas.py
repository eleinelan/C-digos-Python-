import sys, time, os, re
from pathlib import Path
from datetime import datetime, timedelta
from xml.etree import ElementTree as ET

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

# ----------------------------
# CONFIG
# ----------------------------
LOGIN_URL = "https://www.nfse.gov.br/EmissorNacional/Login?ReturnUrl=%2fEmissorNacional"
HOME_URL = "https://www.nfse.gov.br/EmissorNacional"
HOME_URL_PATH = "/EmissorNacional"
EMITIDAS_HREF = "/EmissorNacional/Notas/Emitidas"
RECEBIDAS_HREF = "/EmissorNacional/Notas/Recebidas"

TIMEOUT_LONG = 120
TIMEOUT_MED = 25
TIMEOUT_SHORT = 8

HEADLESS = False

USE_CHROME_PROFILE = False
USER_DATA_DIR = r"C:\Users\SEUUSUARIO\AppData\Local\Google\Chrome\User Data"
PROFILE_DIR = "Default"

DOWNLOAD_DIR = str(Path.home() / "Downloads")
USE_FIRST_TWO_WORDS = True

MAX_PAGES = 50     # trava de segurança para paginação
SAVE_SCREENSHOTS = False  # só cria pasta de saída se der erro

# ----------------------------
# UTILS
# ----------------------------
def _prev_month_year(base_dt=None):
    if base_dt is None: base_dt = datetime.now()
    first = base_dt.replace(day=1)
    last_prev = first - timedelta(days=1)
    return last_prev.month, last_prev.year

def _parse_br_date(txt):
    txt = (txt or "").strip()
    try: return datetime.strptime(txt, "%d/%m/%Y")
    except Exception: return None

def _sanitize_filename(name: str) -> str:
    name = re.sub(r"[^\w \-\.]+", "", name, flags=re.UNICODE)
    return re.sub(r"\s{2,}", " ", name).strip()

def _list_downloaded_files(directory: str, ext_filter: str | None = None) -> set[str]:
    files = {f for f in os.listdir(directory) if not f.endswith(".crdownload")}
    if ext_filter: files = {f for f in files if f.lower().endswith(ext_filter.lower())}
    return files

def _wait_new_download(directory: str, before: set[str], expected_ext: str, timeout: int = 60) -> str | None:
    end = time.time() + timeout
    while time.time() < end:
        current = _list_downloaded_files(directory, expected_ext)
        new_files = list(current - before)
        if new_files:
            paths = [Path(directory) / nf for nf in new_files]
            paths.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            return str(paths[0])
        time.sleep(0.4)
    return None

def _apply_prefix(fullpath: str, prefix: str) -> str:
    if not fullpath: return ""
    p = Path(fullpath)
    safe_prefix = _sanitize_filename(prefix)
    desired = p.with_name(f"{safe_prefix} {p.name}")
    if desired.exists():
        base, ext, i = desired.stem, desired.suffix, 2
        while desired.exists():
            desired = desired.with_name(f"{base} ({i}){ext}"); i += 1
    try:
        p.rename(desired); return str(desired)
    except Exception:
        return fullpath

def _shorten_name(name: str) -> str:
    tokens = (name or "").strip().upper().split()
    name = " ".join(tokens[:2]) if (USE_FIRST_TWO_WORDS and len(tokens) >= 2) else " ".join(tokens)
    return _sanitize_filename(name)

# ----------------------------
# XML parsing
# ----------------------------
def _et_all_text_by_tail_tag(root: ET.Element, tag_endings: list[str]) -> list[str]:
    found = []
    for elem in root.iter():
        t = elem.tag
        if isinstance(t, str) and any(t.endswith(tail) for tail in tag_endings):
            if elem.text and elem.text.strip(): found.append(elem.text.strip())
    return found

def extract_names_from_xml(xml_path: str) -> dict:
    result = {"prestador": "", "tomador": ""}
    try:
        parser = ET.XMLParser(encoding="utf-8")
        tree = ET.parse(xml_path, parser=parser)
        root = tree.getroot()
    except Exception:
        return result
    names_generic = _et_all_text_by_tail_tag(root, ["RazaoSocial","xNome","NomeFantasia"])
    if names_generic:
        result["prestador"] = _shorten_name(names_generic[0])
        for nm in names_generic[1:]:
            if nm != names_generic[0]:
                result["tomador"] = _shorten_name(nm); break
    try:
        for blk in root.iter():
            tag = blk.tag if isinstance(blk.tag, str) else ""
            if tag.endswith("PrestadorServico"):
                txts=[c.text.strip() for c in blk.iter() if isinstance(c.tag,str) and c.text and c.text.strip() and (c.tag.endswith("RazaoSocial") or c.tag.endswith("xNome") or c.tag.endswith("NomeFantasia"))]
                if txts: result["prestador"]=_shorten_name(txts[0])
            if tag.endswith("TomadorServico"):
                txts=[c.text.strip() for c in blk.iter() if isinstance(c.tag,str) and c.text and c.text.strip() and (c.tag.endswith("RazaoSocial") or c.tag.endswith("xNome"))]
                if txts: result["tomador"]=_shorten_name(txts[0])
    except Exception: pass
    return result

# ----------------------------
# Selenium helpers
# ----------------------------
def make_driver(headless: bool = False) -> webdriver.Chrome:
    chrome_opts = Options()
    if headless: chrome_opts.add_argument("--headless=new")
    if USE_CHROME_PROFILE:
        chrome_opts.add_argument(f"--user-data-dir={USER_DATA_DIR}")
        chrome_opts.add_argument(f"--profile-directory={PROFILE_DIR}")
    chrome_opts.add_argument("--start-maximized")
    chrome_opts.add_argument("--disable-gpu")
    chrome_opts.add_argument("--disable-dev-shm-usage")
    chrome_opts.add_argument("--no-sandbox")
    chrome_opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_opts.add_experimental_option("useAutomationExtension", False)
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }
    chrome_opts.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=chrome_opts)
    driver.implicitly_wait(0)
    return driver

def wait_until_logged_in(driver):
    wait = WebDriverWait(driver, TIMEOUT_LONG)
    def ok(_d):
        if HOME_URL_PATH in _d.current_url: return True
        try: _d.find_element(By.CSS_SELECTOR, f'a[href="{EMITIDAS_HREF}"]'); return True
        except Exception: return False
    wait.until(ok)

def _click_menu_card(driver, href: str, label_for_debug: str):
    try:
        wait = WebDriverWait(driver, TIMEOUT_MED)
        selector = f'a[href="{href}"]'
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
        elem = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
        try:
            elem.click()
        except WebDriverException:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
            time.sleep(0.2)
            try: elem.click()
            except WebDriverException: driver.execute_script("arguments[0].click();", elem)
        try: WebDriverWait(driver, TIMEOUT_MED).until(EC.url_contains(href))
        except TimeoutException: pass
        print(f"✅ Naveguei para: {label_for_debug}")
    except Exception as e:
        print(f"⚠️ Não consegui abrir {label_for_debug}: {e}")

def coletar_linhas_mes_anterior(driver):
    alvo_mes, alvo_ano = _prev_month_year()
    # tenta localizar tabela; se não houver, retorna []
    try:
        tabela = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody")))
    except TimeoutException:
        try: tabela = driver.find_element(By.CSS_SELECTOR, "table tbody")
        except Exception: return []
    linhas = tabela.find_elements(By.CSS_SELECTOR, "tr")
    if not linhas: return []
    resultados = []
    for tr in linhas:
        tds = tr.find_elements(By.CSS_SELECTOR, "td")
        if len(tds) < 2: continue
        emissao_txt = tds[0].text.strip()
        emissao_dt = _parse_br_date(emissao_txt)
        if not emissao_dt or (emissao_dt.month != alvo_mes or emissao_dt.year != alvo_ano):
            continue
        empresa = tds[1].text.strip()
        competencia = tds[2].text.strip() if len(tds) > 2 else ""
        municipio = tds[3].text.strip() if len(tds) > 3 else ""
        preco = tds[4].text.strip() if len(tds) > 4 else ""
        situacao = tds[5].text.strip() if len(tds) > 5 else ""
        resultados.append({
            "tr": tr,
            "Emissão": emissao_txt,
            "Emitida para": empresa,
            "Competência": competencia,
            "Município Emissor": municipio,
            "Preço Serviço (R$)": preco,
            "Situação": situacao,
        })
    return resultados

def abrir_menu_linha(driver, tr):
    try: el = tr.find_element(By.CSS_SELECTOR, "a.icone-trigger")
    except Exception:
        try: el = tr.find_element(By.CSS_SELECTOR, ".glyphicon.glyphicon-option-vertical")
        except Exception: return False
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.1)
    try: el.click()
    except Exception: driver.execute_script("arguments[0].click();", el)
    try:
        WebDriverWait(driver, TIMEOUT_SHORT).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".popover .popover-content"))
        ); return True
    except TimeoutException: return False

def clicar_download_xml(driver):
    link = WebDriverWait(driver, TIMEOUT_MED).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.popover-content a[href*="/EmissorNacional/Notas/Download/NFSe/"]'))
    ); link.click()

def clicar_download_danfse(driver):
    link = WebDriverWait(driver, TIMEOUT_MED).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.popover-content a[href*="/EmissorNacional/Notas/Download/DANFSe/"]'))
    ); link.click()

# --------- paginação ----------
def _find_next_button(driver):
    candidates = [
        ("CSS", "ul.pagination li.next:not(.disabled) a"),
        ("CSS", "ul.pagination li a[rel='next']"),
        ("CSS", "ul.pagination li a[aria-label='Próximo'], ul.pagination li a[aria-label='Proximo']"),
        ("XPATH", "//ul[contains(@class,'pagination')]//a[contains(.,'Próxima') or contains(.,'Próximo') or contains(.,'›') or contains(.,'»')]"),
    ]
    for kind, sel in candidates:
        try:
            if kind == "CSS":
                el = driver.find_element(By.CSS_SELECTOR, sel)
            else:
                el = driver.find_element(By.XPATH, sel)
            # checa se não está desabilitado via pai
            try:
                li = el.find_element(By.XPATH, "./ancestor::li[1]")
                cls = li.get_attribute("class") or ""
                if "disabled" in cls.lower(): continue
            except Exception:
                pass
            return el
        except Exception:
            continue
    return None

def _go_next_page(driver):
    btn = _find_next_button(driver)
    if not btn: return False
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    try:
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)
    # espera tabela “mudar” minimamente
    time.sleep(0.6)
    return True

# ----------------------------
# Processadores por página (com paginação)
# ----------------------------
def processar_pagina(driver, pagina_tipo: str, href: str):
    """
    pagina_tipo: "emitidas" (prestados) usa prefixo do PRESTADOR;
                 "recebidas" (tomados) usa prefixo do TOMADOR.
    """
    _click_menu_card(driver, href, f"NFS-e {pagina_tipo.capitalize()}")

    alvo_mes, alvo_ano = _prev_month_year()
    planilha_rows, excel_prefix_for_batch = [], None
    pagina = 1

    while pagina <= MAX_PAGES:
        linhas = coletar_linhas_mes_anterior(driver)
        if not linhas:
            # se não tem linhas nesta página, tenta próxima; se não houver próxima, encerra
            if not _go_next_page(driver): break
            pagina += 1
            continue

        print(f"— Página {pagina} ({pagina_tipo}) —")

        for idx, item in enumerate(linhas, start=1):
            tr = item["tr"]; emissao = item["Emissão"]; empresa_coluna = item["Emitida para"]
            print(f"▶️ [{pagina_tipo}] Linha {idx}: {empresa_coluna} — Emissão {emissao}")

            if not abrir_menu_linha(driver, tr):
                print("   ⚠️ Não consegui abrir o menu desta linha. Pulando…"); continue

            # XML primeiro
            before_xml = _list_downloaded_files(DOWNLOAD_DIR, ".xml")
            try:
                clicar_download_xml(driver)
            except Exception as e:
                print(f"   ⚠️ Erro ao clicar 'Download XML': {e}"); continue
            xml_path = _wait_new_download(DOWNLOAD_DIR, before_xml, ".xml", timeout=60)
            if not xml_path:
                print("   ⚠️ XML não detectado."); continue
            print(f"   ✅ XML baixado: {xml_path}")

            names = extract_names_from_xml(xml_path)
            prefix = (names.get("prestador") or names.get("tomador") or "NFSE") if pagina_tipo=="emitidas" \
                     else (names.get("tomador") or names.get("prestador") or "NFSE")
            if not excel_prefix_for_batch and prefix: excel_prefix_for_batch = prefix

            xml_renamed = _apply_prefix(xml_path, prefix); print(f"   🏷  XML renomeado: {xml_renamed}")

            # PDF
            if not abrir_menu_linha(driver, tr):
                print("   ⚠️ Não consegui reabrir o menu para baixar o DANFS-e. Pulando PDF…")
            else:
                before_pdf = _list_downloaded_files(DOWNLOAD_DIR, ".pdf")
                try:
                    clicar_download_danfse(driver)
                    pdf_path = _wait_new_download(DOWNLOAD_DIR, before_pdf, ".pdf", timeout=60)
                    if pdf_path:
                        pdf_renamed = _apply_prefix(pdf_path, prefix); print(f"   🏷  PDF renomeado: {pdf_renamed}")
                    else:
                        print("   ⚠️ PDF não detectado.")
                except Exception as e:
                    print(f"   ⚠️ Erro ao clicar 'Download DANFS-e': {e}")

            planilha_rows.append({
                "Emissão": item["Emissão"],
                "Emitida para": empresa_coluna,
                "Competência": item["Competência"],
                "Município Emissor": item["Município Emissor"],
                "Preço Serviço (R$)": item["Preço Serviço (R$)"],
                "Situação": item["Situação"],
                "Prefixo usado": prefix,
            })

        # tenta ir para próxima página
        if not _go_next_page(driver): break
        pagina += 1

    # Planilha do mês
    try:
        import pandas as pd
        if planilha_rows:
            df = pd.DataFrame(planilha_rows)
            batch_prefix = excel_prefix_for_batch or "NFSE"
            excel_name = f"{batch_prefix} NFSe_{pagina_tipo.capitalize()}_{alvo_ano:04d}-{alvo_mes:02d}.xlsx"
            excel_path = Path(DOWNLOAD_DIR) / _sanitize_filename(excel_name)
            df.to_excel(excel_path, index=False)
            print(f"📄 Planilha gerada [{pagina_tipo}]: {excel_path}")
        else:
            print(f"ℹ️ Não havia linhas do mês anterior para registrar na planilha ({pagina_tipo}).")
    except Exception as e:
        print(f"⚠️ Não consegui salvar a planilha Excel ({pagina_tipo}): {e}\nTente: pip install pandas openpyxl")

# ----------------------------
# MAIN
# ----------------------------
def main():
    out_dir = Path("nfse_automacao_out")
    driver = make_driver(headless=HEADLESS)
    try:
        driver.get(LOGIN_URL); print("👉 Faça o login manualmente (certificado/conta).")
        wait_until_logged_in(driver)
        if HOME_URL_PATH not in driver.current_url: driver.get(HOME_URL)

        # Emitidas (prestados)
        processar_pagina(driver, "emitidas", EMITIDAS_HREF)
        # Recebidas (tomados)
        processar_pagina(driver, "recebidas", RECEBIDAS_HREF)

        time.sleep(0.6)
    except Exception as e:
        if SAVE_SCREENSHOTS:
            out_dir.mkdir(exist_ok=True)
            shot = out_dir / "erro_nfse.png"
            try:
                driver.save_screenshot(str(shot))
                print(f"❌ Erro: {e}\nScreenshot salvo em: {shot.resolve()}")
            except Exception:
                print(f"❌ Erro: {e}")
        else:
            print(f"❌ Erro: {e}")
        sys.exit(1)
    finally:
        # driver.quit()
        pass

if __name__ == "__main__":
    main()
