# nftse_nfts_bot.py — PMSP NFTS
# NFTS – SERVIÇOS TOMADOS (PDF+TXT) para TODAS as empresas + planilha.
# LOGIN GUIADO: abre login.aspx, mostra banner com botão "Continuar execução".
# Ao prosseguir, o script navega até INÍCIO, abre "Consulta de NFTS" e
# clica em "NFTS - SERVIÇOS TOMADOS" automaticamente.

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, StaleElementReferenceException,
    UnexpectedAlertPresentException
)
from pathlib import Path
from datetime import datetime
import base64, re, time, csv, sys, traceback

URL_LOGIN   = "https://nfe.prefeitura.sp.gov.br/login.aspx"
URL_INICIO  = "https://nfe.prefeitura.sp.gov.br/contribuinte/inicio.aspx"
URL_CONSULTA_NFTS = "https://nfe.prefeitura.sp.gov.br/contribuinte/consultasnfts.aspx"

# ========================= UTIL =========================

def log(msg):
    print("[LOG]", msg, flush=True)

def sanitize(name: str) -> str:
    name = re.sub(r'[\r\n\t]+', ' ', name)
    name = re.sub(r'[\\/:*?"<>|]+', ' ', name)
    name = re.sub(r'\s{2,}', ' ', name).strip()
    return (name[:140].rstrip() if len(name) > 140 else name) or "Empresa"

def mes_ano_anterior():
    hoje = datetime.now()
    ano, mes = hoje.year, hoje.month - 1
    if mes == 0:
        mes, ano = 12, ano - 1
    return f"{mes:02d}", str(ano)


def create_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    # downloads genéricos para qualquer usuário
    prefs = {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=opts)

# ========================= IFRAME HELPERS =========================

def switch_into_iframe_with(driver, xpath_target: str, max_scan=12) -> bool:
    driver.switch_to.default_content()
    try:
        if driver.find_elements(By.XPATH, xpath_target):
            return True
    except Exception:
        pass
    for f in driver.find_elements(By.TAG_NAME, "iframe")[:max_scan]:
        try:
            driver.switch_to.default_content(); driver.switch_to.frame(f)
            if driver.find_elements(By.XPATH, xpath_target):
                return True
        except Exception:
            pass
    driver.switch_to.default_content()
    return False

# ========================= LOGIN GUIADO =========================

def _inject_login_overlay(driver):
    js = r"""
    (function(){
      try{
        var go = (sessionStorage.getItem('nfts_go') === '1');
        if (go) { window.__nfts_go__ = true; }
        if (!document.getElementById('nftsOverlay')){
          const style = document.createElement('style');
          style.textContent =
          '#nftsOverlay{position:fixed;right:16px;bottom:16px;z-index:2147483647;background:#0b7d8c;color:#fff;padding:16px 18px;max-width:380px;border-radius:14px;box-shadow:0 10px 30px rgba(0,0,0,.25);font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif}' +
          '#nftsOverlay h3{margin:0 0 8px 0;font-size:16px;font-weight:800}' +
          '#nftsOverlay ol{margin:8px 0 12px 20px;font-size:14px;line-height:1.35}' +
          '#nftsOverlay button{background:#fff;color:#0b7d8c;border:none;font-weight:800;padding:10px 14px;border-radius:10px;cursor:pointer}';
          document.head.appendChild(style);
          const box = document.createElement('div');
          box.id = 'nftsOverlay';
          box.innerHTML =
            '<h3>Passos para iniciar</h3>' +
            '<ol>'+
            '<li>Faça login no portal.</li>'+
            '<li>No menu, abra <b>Consulta de Notas → Consulta de NFTS</b>.</li>'+
            '<li>Deixe a página de filtros visível (contribuinte/Incidência/mês).</li>'+
            '</ol>'+
            '<button id="nftsGoBtn" type="button">Continuar execução</button>';
          document.body.appendChild(box);
          document.getElementById('nftsGoBtn').onclick = function(){
            sessionStorage.setItem('nfts_go','1');
            window.__nfts_go__ = true;
            try{ alert("Ok! Vou continuar."); }catch(e){}
          };
        }
      }catch(e){}
    })();
    """
    try: driver.execute_script(js)
    except Exception: pass


def _filtros_prontos(driver) -> bool:
    # Consideramos "pronto" quando estivermos na consultasnfts.aspx ou quando
    # existir o select de contribuintes/Incidência na árvore atual (mesmo em iframe)
    try:
        cur = (driver.current_url or '').lower()
    except Exception:
        cur = ''

    # 1) Se já está na página de consulta, ótimo
    if 'consultasnfts.aspx' in cur:
        # Ainda validamos a presença de algum filtro
        xp_any = (
            "//select[option[contains(.,'Selecione o contribuinte')]]"
            "| //label[contains(.,'Incidência')]"
            "| //select[option[normalize-space(.)='1' or normalize-space(.)='12' or normalize-space(.)='2']]"
        )
        try:
            if driver.find_elements(By.XPATH, xp_any):
                return True
        except Exception:
            pass

    # 2) Caso a URL não seja a de consulta, tentamos detectar os filtros por iframe
    xp = "//select[option[contains(.,'Selecione o contribuinte')]]"
    try:
        if driver.find_elements(By.XPATH, xp):
            return True
        for f in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.default_content(); driver.switch_to.frame(f)
                if driver.find_elements(By.XPATH, xp):
                    return True
            except Exception:
                pass
        driver.switch_to.default_content()
    except Exception:
        driver.switch_to.default_content()
    return False


def wait_login_and_open_nfts(driver, timeout=900):
    log("Abrindo login.aspx e aguardando você finalizar o login…")
    driver.get(URL_LOGIN)
    t0 = time.time()

    while time.time() - t0 < timeout:
        _inject_login_overlay(driver)

        # Se já estamos na tela de consulta (após o login manual), não volte para INÍCIO
        try:
            cur = (driver.current_url or '').lower()
        except Exception:
            cur = ''
        if ('consultasnfts.aspx' in cur) and _filtros_prontos(driver):
            log("Tela de filtros da NFTS pronta (consultasnfts.aspx).")
            return

        # Também aceitamos estar em INÍCIO com filtros carregados por iframe
        if _filtros_prontos(driver) and 'inicio.aspx' in cur:
            log("Tela de filtros da NFTS pronta (via INÍCIO).")
            return

        # Botão do banner: apenas navegue se ainda não está na consultasnfts.aspx
        try:
            go = driver.execute_script("return !!window.__nfts_go__ || sessionStorage.getItem('nfts_go')==='1';")
        except Exception:
            go = False

        if go:
            if 'consultasnfts.aspx' not in cur:
                # Abre INÍCIO → Consulta de NFTS → NFTS - SERVIÇOS TOMADOS
                driver.get(URL_INICIO)
                try:
                    _abrir_menu_consulta_nfts(driver)
                    _abrir_pagina_nfts_servicos_tomados(driver)
                except Exception:
                    log("Não consegui abrir automaticamente. Você pode abrir manualmente e clicar no banner de novo.")
            # Se chegamos aqui e a página de consulta está pronta, retorna
            if _filtros_prontos(driver):
                log("Tela de filtros da NFTS pronta.")
                return
        time.sleep(0.4)

    raise TimeoutException("Tempo máximo de login esgotado.")

# ========================= NAVEGAÇÃO MENU =========================

def _click_by_text(driver, patterns):
    # Clica no primeiro elemento que contenha o texto indicado
    for xp in [
        "//a[normalize-space()[contains(.,$txt)]]",
        "//*[self::a or self::span or self::div or self::li][contains(normalize-space(.),$txt)]",
    ]:
        for pat in patterns:
            pat_norm = pat.strip()
            try:
                el = driver.execute_script(
                    """
                    var xp = arguments[0].replace('$txt', arguments[1]);
                    try{ return document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; }catch(e){ return null; }
                    """,
                    xp, pat_norm
                )
                if el:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    try:
                        el.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", el)
                    return True
            except Exception:
                pass
    return False


def _abrir_menu_consulta_nfts(driver):
    # Em muitas instalações o menu abre direto em INÍCIO; basta clicar na opção.
    # Procuramos algo como "Consulta de Notas" e depois "Consulta de NFTS".
    driver.switch_to.default_content()
    _click_by_text(driver, ["Consulta de Notas"])  # hover não é necessário se o click expandir
    time.sleep(0.3)
    _click_by_text(driver, ["Consulta de NFTS", "NFTS"])  # fallback genérico
    time.sleep(0.5)


def _abrir_pagina_nfts_servicos_tomados(driver):
    # Dentro de "Consulta de NFTS", clicar em "NFTS - SERVIÇOS TOMADOS"
    driver.switch_to.default_content()
    # Em algumas telas, o link/btn pode estar em um iframe
    xp = "//*[self::a or self::button or self::input][contains(translate(normalize-space(.),'áéíóúãõç','AEIOUAOC'), 'NFTS - SERVICOS TOMADOS') or contains(@value,'NFTS - SERVI') or contains(@title,'NFTS')]"
    if not switch_into_iframe_with(driver, xp):
        # tentar encontrar sem iframe também
        pass
    try:
        btn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xp)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)
    except Exception:
        log("Aviso: não consegui clicar em 'NFTS - SERVIÇOS TOMADOS' automaticamente. Selecione manualmente e clique no banner novamente.")
    finally:
        driver.switch_to.default_content()

# ========================= FILTROS =========================

def localizar_select_contribuinte(driver):
    xp = "//select[option[contains(.,'Selecione o contribuinte')]]"
    assert switch_into_iframe_with(driver, xp), "Não achei o select de contribuintes."
    return WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.XPATH, xp)))


def listar_contribuintes(driver):
    sel = localizar_select_contribuinte(driver)
    textos = []
    for _ in range(3):
        try:
            S = Select(sel)
            if len(S.options) <= 1:
                time.sleep(0.4); continue
            for opt in S.options[1:]:
                t = (opt.text or "").strip()
                if t and "Selecione o contribuinte" not in t:
                    textos.append(t)
            break
        except StaleElementReferenceException:
            sel = localizar_select_contribuinte(driver); time.sleep(0.2)
    driver.switch_to.default_content()
    return textos


def selecionar_contribuinte(driver, texto_visivel: str) -> str:
    sel = localizar_select_contribuinte(driver)
    for _ in range(3):
        try:
            Select(sel).select_by_visible_text(texto_visivel)
            raz = texto_visivel.split(" - ", 1)[-1].strip() if " - " in texto_visivel else texto_visivel
            log(f"Contribuinte selecionado: {raz}")
            driver.switch_to.default_content()
            return sanitize(raz)
        except StaleElementReferenceException:
            sel = localizar_select_contribuinte(driver); time.sleep(0.2)
    driver.switch_to.default_content()
    return "Empresa"


def marcar_incidencia(driver):
    xp = "//label[contains(.,'Incidência')]/preceding-sibling::input[@type='radio']"
    if not switch_into_iframe_with(driver, xp):
        xp = "//input[@type='radio' and (../label[contains(.,'Incidência')])]"
        if not switch_into_iframe_with(driver, xp):
            log("Aviso: não achei 'Incidência'.")
            return
    try:
        el = driver.find_element(By.XPATH, xp)
        driver.execute_script("arguments[0].click();", el)
        log("Marcado: Incidência.")
    except Exception:
        log("Aviso: falha ao clicar em Incidência.")
    driver.switch_to.default_content()


def _pick_select_ano_mes(driver):
    xp_mes = "//select[option[normalize-space(.)='1' or normalize-space(.)='12' or normalize-space(.)='2']]"
    if not switch_into_iframe_with(driver, xp_mes):
        return None, None
    sel_ano = sel_mes = None
    for sel in driver.find_elements(By.TAG_NAME, "select"):
        try:
            texts = [(o.text or "").strip() for o in sel.find_elements(By.TAG_NAME, "option")]
        except StaleElementReferenceException:
            continue
        if not texts: continue
        if sum(1 for t in texts if re.fullmatch(r"\d{4}", t)) >= 4:
            sel_ano = sel
        if sum(1 for t in texts if re.fullmatch(r"(0?[1-9]|1[0-2])", t)) >= 8:
            sel_mes = sel
    return sel_ano, sel_mes


def set_periodo_mes_anterior(driver):
    mm, yyyy = mes_ano_anterior()
    sel_ano, sel_mes = _pick_select_ano_mes(driver)
    if sel_ano:
        for _ in range(3):
            try:
                S = Select(sel_ano)
                try: S.select_by_visible_text(yyyy)
                except: S.select_by_value(yyyy)
                break
            except StaleElementReferenceException:
                sel_ano, sel_mes = _pick_select_ano_mes(driver); time.sleep(0.2)
    alvo_mes_txt = str(int(mm))
    if sel_mes:
        for _ in range(3):
            try:
                S = Select(sel_mes)
                try: S.select_by_visible_text(alvo_mes_txt)
                except:
                    try: S.select_by_visible_text(mm)
                    except: S.select_by_value(mm)
                break
            except StaleElementReferenceException:
                sel_ano, sel_mes = _pick_select_ano_mes(driver); time.sleep(0.2)
    driver.switch_to.default_content()
    log(f"Período ajustado para {mm}/{yyyy}.")
    return mm, yyyy

# ========================= AÇÕES NA PÁGINA DE RESULTADO =========================

def _esperar_tabela(driver):
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//table | //th[contains(.,'Tomador') or contains(.,'Série')]"))
        )
    except TimeoutException:
        log("Aviso: não identifiquei claramente a tabela; vou imprimir mesmo assim.")


def imprimir_pdf(driver, nome_base: str) -> Path:
    pdf = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
    data = base64.b64decode(pdf["data"])
    out = Path.home() / "Downloads" / (sanitize(nome_base) + ".pdf")
    with open(out, "wb") as f: f.write(data)
    log(f"PDF salvo em: {out}")
    return out


def exportar_txt(driver, nome_base: str):
    downloads = Path.home() / "Downloads"
    antes = {p for p in downloads.glob("*")}

    # Seleção TXT
    xp_sel = "//select[option[normalize-space(.)='TXT'] or option[contains(.,'TXT')]]"
    if not switch_into_iframe_with(driver, xp_sel):
        log("Aviso: select do formato (TXT) não encontrado.")
        driver.switch_to.default_content()
        return None
    try:
        sel = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xp_sel)))
        Select(sel).select_by_visible_text("TXT")
    except Exception as e:
        log(f"Aviso: não consegui selecionar TXT ({e}).")
        driver.switch_to.default_content()
        return None

    # Botão Exportar
    btn = None
    for xp in [
        "//input[@type='button' or @type='submit'][@value='Exportar']",
        "//button[normalize-space()='Exportar']",
        "//*[self::a or self::span or self::div][normalize-space()='Exportar']"
    ]:
        try:
            btn = driver.find_element(By.XPATH, xp)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            driver.execute_script("arguments[0].click();", btn)
            break
        except Exception:
            btn = None
    driver.switch_to.default_content()
    if btn is None:
        log("Aviso: botão 'Exportar' não encontrado.")
        return None

    alvo = downloads / (sanitize(nome_base) + ".txt")
    t0 = time.time()
    ultimo = None
    while time.time() - t0 < 30:
        atuais = {p for p in downloads.glob("*.txt")}
        novos = [p for p in atuais - {p for p in antes if p.suffix.lower()=='.txt'} if not p.name.endswith(".crdownload")]
        if novos:
            novo = max(novos, key=lambda p: p.stat().st_mtime)
            ultimo = novo
            try:
                if alvo.exists(): alvo.unlink()
                novo.replace(alvo)
                log(f"TXT salvo em: {alvo}")
                return alvo
            except PermissionError:
                time.sleep(0.5)
        time.sleep(0.4)
    if ultimo:
        log(f"Aviso: TXT baixado como '{ultimo.name}', mas não renomeado.")
        return ultimo
    log("Aviso: não detectei o download do TXT.")
    return None


def extrair_razao_ccm(driver) -> str:
    driver.switch_to.default_content()
    try:
        hdr = driver.find_element(By.XPATH, "//*[contains(translate(.,'ccm','CCM'),'CCM')][1]")
        linhas = [ln.strip() for ln in hdr.text.splitlines() if "CCM" in ln.upper()]
        if linhas:
            ln = linhas[0].split("FILTROS", 1)[0]
            idx = ln.rfind(" - ")
            if idx != -1: return sanitize(ln[idx+3:].strip())
    except Exception:
        pass
    return "Empresa"


def extrair_valor_servicos(driver) -> str:
    driver.switch_to.default_content()
    try: txt = driver.find_element(By.TAG_NAME, "body").text
    except: return ""
    for rgx in [
        r'Valor dos Servi[cç]os[:\s]+R?\$?\s*([\d\.\s]+,\d{2})',
        r'Total[:\s]+R?\$?\s*([\d\.\s]+,\d{2})',
        r'Valor Total[:\s]+R?\$?\s*([\d\.\s]+,\d{2})',
    ]:
        m = re.search(rgx, txt, flags=re.IGNORECASE)
        if m:
            return re.sub(r'\s+', '', m.group(1).strip())
    return ""


def salvar_excel(razao: str, mm: str, yyyy: str, valor: str):
    downloads = Path.home() / "Downloads"
    xlsx = downloads / "relatorio_nftse.xlsx"
    row = {"Tipo": "NFTS - SERVIÇOS TOMADOS", "Razão Social": razao, "Período": f"{mm}/{yyyy}", "Valor dos Serviços": valor or ""}
    try:
        import pandas as pd
        if xlsx.exists():
            df = pd.read_excel(xlsx)
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        else:
            df = pd.DataFrame([row])
        df.to_excel(xlsx, index=False)
        log(f"Planilha atualizada: {xlsx}")
    except Exception as e:
        csv_path = downloads / "relatorio_nftse.csv"
        write_header = not csv_path.exists()
        with open(csv_path, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=row.keys())
            if write_header: w.writeheader()
            w.writerow(row)
        log(f"Aviso: sem pandas/openpyxl ({e}). Salvei no CSV: {csv_path}")

# ========================= FLUXO NFTS =========================

def _abrir_relatorio_nfts(driver) -> str:
    """Clica em "NFTS - SERVIÇOS TOMADOS" e retorna o handle da nova janela/aba se abrir."""
    xp_btn = (
        "//input[@type='button' or @type='submit'][contains(@value,'NFTS') or contains(@value,'Tomados')]" \
        "| //button[contains(.,'NFTS') or contains(.,'Tomados')]" \
        "| //a[contains(.,'NFTS -') or contains(.,'Serviços Tomados') or contains(.,'SERVIÇOS TOMADOS')]"
    )
    assert switch_into_iframe_with(driver, xp_btn), "Não encontrei o botão/ação de NFTS - SERVIÇOS TOMADOS."
    prev = set(driver.window_handles)
    start_url = driver.current_url

    try:
        btn = driver.find_element(By.XPATH, xp_btn)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        driver.execute_script("arguments[0].click();", btn)
    except UnexpectedAlertPresentException:
        try: driver.switch_to.alert.accept()
        except Exception: pass

    driver.switch_to.default_content()

    end = time.time() + 20
    while time.time() < end:
        cur = set(driver.window_handles)
        if len(cur) > len(prev):
            for h in cur:
                if h not in prev:
                    return h
        if driver.current_url != start_url:
            return driver.current_window_handle
        time.sleep(0.2)

    raise TimeoutException("Não consegui abrir a tela da NFTS.")


def processar_nfts(driver, razao_filtros, mm, yyyy, main_handle):
    h = _abrir_relatorio_nfts(driver)
    driver.switch_to.window(h)
    _esperar_tabela(driver)
    razao = extrair_razao_ccm(driver) or razao_filtros
    valor = extrair_valor_servicos(driver)
    base = sanitize(f"{razao} – NFTS – SERVIÇOS TOMADOS – {yyyy}-{mm}")
    imprimir_pdf(driver, base)
    exportar_txt(driver, base)
    salvar_excel(razao, mm, yyyy, valor)
    try:
        if driver.current_window_handle != main_handle:
            driver.close()
    except Exception:
        pass
    driver.switch_to.window(main_handle)


def processar_empresa(driver, texto_opt: str, main_handle: str):
    razao_filtros = selecionar_contribuinte(driver, texto_opt)
    marcar_incidencia(driver)
    mm, yyyy = set_periodo_mes_anterior(driver)

    try:
        processar_nfts(driver, razao_filtros, mm, yyyy, main_handle)
    except Exception as e:
        log(f"Atenção (NFTS) '{texto_opt}': {e}")
        traceback.print_exc()


# ========================= MAIN =========================

def main():
    driver = create_driver()
    try:
        wait_login_and_open_nfts(driver)  # LOGIN + abrir Consulta de NFTS → NFTS - SERVIÇOS TOMADOS
        main_handle = driver.current_window_handle

        empresas = listar_contribuintes(driver)
        if not empresas:
            log("Não encontrei empresas na lista (só o placeholder?).")
            return

        log(f"Total de empresas na lista: {len(empresas)}")
        for i, texto_opt in enumerate(empresas, start=1):
            log(f"----- [{i}/{len(empresas)}] {texto_opt} -----")
            try:
                driver.switch_to.window(main_handle)
                processar_empresa(driver, texto_opt, main_handle)
            except Exception as e:
                log(f"Falha ao processar '{texto_opt}': {e}")
                traceback.print_exc()
            time.sleep(0.6)

        log("Concluído para todas as empresas.")
    finally:
        pass  # mantém o navegador aberto para revisão


if __name__ == "__main__":
    main()
