# fsist_end2end_todas.py
# -*- coding: utf-8 -*-

import time
import zipfile
import shutil
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# =========================
# CONFIG
# =========================
URL = "https://www.fsist.com.br/usuario/monitor-de-notas"
WAIT = 50

DOWNLOAD_DIR = Path.home() / "Downloads"

# Saídas pedidas
FINAL_DIR   = DOWNLOAD_DIR / "FSist-NFe entradas-Todas"        # extração do ZIP
FINAL_PRINT = DOWNLOAD_DIR / "FSist-NFe entradas-Todas.png"    # print ANTES dos downloads
EXCEL_FIXED = DOWNLOAD_DIR / "FSist-NFe entradas-Todas.xlsx"   # planilha renomeada (fixo)

# Padrões de arquivos gerados pela FSist
ZIP_PREFIX  = "FSist XMLs N"
ZIP_EXT     = ".zip"
XLSX_PREFIX = "FSist-NFe-Todas--"
XLSX_EXT    = ".xlsx"

# =========================
# SELETORES
# =========================
ABA_RECEBIDAS     = (By.ID, "TabPageEsqNFeRecebidas")
PERIODO_SPAN      = (By.ID, "Periodo")
MES_PASSADO       = (By.ID, "DataMesPassado")
CAL_PERIODO_TRIGG = [
    (By.XPATH, "//*[@id='Periodo']/ancestor::*[self::div or self::button][1]"),
    (By.XPATH, "//*[contains(@class,'icon-calendar')]/ancestor::*[self::div or self::button][1]"),
]

BTN_SELECIONAR_TODAS = (By.ID, "butSelecionadosQtd")

# Relatório (Excel)
BTN_RELATORIO = [
    (By.XPATH, "//*[contains(@class,'icon-excel')]/ancestor::*[self::button or self::a or self::div][1]"),
    (By.XPATH, "//button[contains(., 'Relatório') or contains(., 'Relatorio')]"),
]
BTN_GERAR_RELATORIO = [
    (By.XPATH, "//button[.//i[contains(@class,'icon-excel')] and contains(., 'GERAR')]"),
    (By.XPATH, "//button[contains(translate(., 'ÉéÍíÓóÂâÃãÁáÊêÚúÕõÇç','EeIiOoAaAaAaEeUuOoCc'),'GERAR RELATORIO')]"),
]

# Download XMLs & PDFs
BTN_DOWNLOAD  = (By.ID, "butDownload")  # botão verde da barra
BTN_XMLS_PDFS = (By.XPATH, "//button[.//span[normalize-space()='XMLs e PDFs']]")
BTN_CIENCIA   = (By.XPATH, "//button[contains(., 'Sim, efetuar ciência da operação')]")

# =========================
# UTILITÁRIOS
# =========================
def build_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

def js_click(driver, el):
    driver.execute_script("arguments[0].click();", el)

def wait_and_click(driver, locator, label, scroll=True, timeout=WAIT):
    el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))
    if scroll:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        WebDriverWait(driver, 4).until(EC.element_to_be_clickable(locator)).click()
    except Exception:
        js_click(driver, el)
    print(f"✓ {label}")

def try_click_any(driver, locators, label, timeout_each=5):
    for loc in locators:
        try:
            wait_and_click(driver, loc, label, timeout=timeout_each)
            return True
        except Exception:
            continue
    return False

def wait_download_complete(dirpath: Path, timeout=420):
    start = time.time()
    while time.time() - start < timeout:
        if not list(dirpath.glob("*.crdownload")):
            return True
        time.sleep(1)
    return False

def newest_file_in(dirpath: Path, startswith=None, endswith=None, timeout=240):
    start = time.time()
    while time.time() - start < timeout:
        files = []
        for p in dirpath.glob("*"):
            if not p.is_file():
                continue
            if startswith and not p.name.startswith(startswith):
                continue
            if endswith and not p.name.endswith(endswith):
                continue
            if Path(str(p) + ".crdownload").exists():
                continue
            files.append(p)
        if files:
            return max(files, key=lambda f: f.stat().st_mtime)
        time.sleep(1)
    return None

def extract_zip_to_named_folder(zip_path: Path, final_dir: Path):
    work_dir = zip_path.parent / "_fsist_extract_tmp"
    if work_dir.exists():
        shutil.rmtree(work_dir, ignore_errors=True)
    work_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as zf:
        # proteção contra path traversal
        for info in zf.infolist():
            target = (work_dir / info.filename).resolve()
            if not str(target).startswith(str(work_dir.resolve())):
                continue
        zf.extractall(work_dir)

    # Se houver uma pasta única, usa; senão, usa o work_dir
    tops = [p for p in work_dir.iterdir()]
    source = tops[0] if len(tops) == 1 and tops[0].is_dir() else work_dir

    if final_dir.exists():
        shutil.rmtree(final_dir, ignore_errors=True)
    shutil.move(str(source), str(final_dir))
    shutil.rmtree(work_dir, ignore_errors=True)

# =========================
# FLUXO PRINCIPAL
# =========================
def main():
    driver = build_driver()
    try:
        # 1) Acesso e (se precisar) login manual
        driver.get(URL)
        print("• Página aberta. Faça o login manualmente se necessário.")
        WebDriverWait(driver, 300).until(
            lambda d: d.find_elements(*ABA_RECEBIDAS) or d.find_elements(By.ID, "Periodo")
        )

        # 2) Ajustar período: Mês passado
        opened = False
        for loc in CAL_PERIODO_TRIGG:
            try:
                wait_and_click(driver, loc, "Abrindo seletor de 'Período'", scroll=False)
                opened = True
                break
            except Exception:
                continue
        if opened:
            wait_and_click(driver, MES_PASSADO, "Aplicando 'Mês passado'", scroll=False)
            time.sleep(0.4)
            try:
                periodo_txt = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(PERIODO_SPAN)
                ).text.strip()
            except Exception:
                periodo_txt = "(não consegui ler)"
            print(f"✓ Período atual: {periodo_txt}")
        else:
            print("⚠ Não consegui abrir o seletor de período (seguirei assim mesmo).")

        # 3) Selecionar todas
        wait_and_click(driver, BTN_SELECIONAR_TODAS, "Clique em 'Selecionar todas'")

        # 3.1) PRINT IMEDIATO (antes de abrir modais)
        driver.save_screenshot(str(FINAL_PRINT))
        print(f"✓ Print salvo em: {FINAL_PRINT}")

        # 4) RELATÓRIO → GERAR RELATÓRIO (Excel) + renomear fixo
        if try_click_any(driver, BTN_RELATORIO, "Abrindo 'Relatório'"):
            if try_click_any(driver, BTN_GERAR_RELATORIO, "Gerando relatório (Excel)", timeout_each=10):
                print("⏳ Aguardando download do Excel…")
                wait_download_complete(DOWNLOAD_DIR, timeout=420)
                xlsx = newest_file_in(DOWNLOAD_DIR, startswith=XLSX_PREFIX, endswith=XLSX_EXT, timeout=240)
                if xlsx:
                    print(f"✓ Excel baixado: {xlsx.name}")
                    try:
                        # apaga o destino se já existir e renomeia (sem data)
                        if EXCEL_FIXED.exists():
                            EXCEL_FIXED.unlink()
                        xlsx.replace(EXCEL_FIXED)
                        print(f"✓ Planilha renomeada para: {EXCEL_FIXED.name}")
                    except Exception as e:
                        print(f"⚠ Não consegui renomear a planilha: {e}")
                else:
                    print("⚠ Não encontrei o Excel baixado.")
            else:
                print("⚠ Não encontrei o botão 'Gerar relatório'. Pulando a planilha.")
        else:
            print("⚠ Não localizei o botão 'Relatório'. Pulando a planilha.")

        # 5) DOWNLOAD → XMLs e PDFs (trata ciência)
        wait_and_click(driver, BTN_DOWNLOAD, "Abrindo 'Download' da barra")
        time.sleep(1)
        try:
            btn = WebDriverWait(driver, 3).until(EC.presence_of_element_located(BTN_CIENCIA))
            js_click(driver, btn)
            print("✓ Confirmei: 'Sim, efetuar ciência da operação'")
            time.sleep(2)
            wait_and_click(driver, BTN_DOWNLOAD, "Abrindo 'Download' novamente")
            time.sleep(1)
        except Exception:
            pass

        wait_and_click(driver, BTN_XMLS_PDFS, "Clicando em 'XMLs e PDFs'")
        print("⏳ Aguardando download do ZIP…")
        wait_download_complete(DOWNLOAD_DIR, timeout=420)
        zipf = newest_file_in(DOWNLOAD_DIR, startswith=ZIP_PREFIX, endswith=ZIP_EXT, timeout=240)
        if not zipf:
            raise RuntimeError("Não encontrei o ZIP (verifique a pasta Downloads).")
        print(f"✓ ZIP baixado: {zipf.name}")

        # 6) Extrair ZIP para Downloads com nome final
        extract_zip_to_named_folder(zipf, FINAL_DIR)
        print(f"✓ Arquivos extraídos em: {FINAL_DIR}")

        print("\n========== CONCLUÍDO ==========")
        print(f"Pasta XML/PDF: {FINAL_DIR}")
        print(f"Planilha: {EXCEL_FIXED if EXCEL_FIXED.exists() else '(não criada)'}")
        print(f"Print: {FINAL_PRINT}")
        print("================================\n")

        input("Revise os arquivos em Downloads. Pressione ENTER para fechar o navegador...")

    except Exception as e:
        print(f"\n✗ Erro: {e}\n")
        input("Pressione ENTER para fechar o navegador...")
    finally:
        try:
            driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()
