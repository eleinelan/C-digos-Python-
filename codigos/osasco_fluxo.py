# -*- coding: utf-8 -*-
import os, re, time, unicodedata, traceback
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# ======================= CONFIG GERAL =======================
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
BASE = "https://nfe.osasco.sp.gov.br"
URL_LOGIN = f"{BASE}/EissnfeWebApp/Portal/Default.aspx?ReturnUrl=%2fEissnfeWebApp%2fSistema%2fGeral%2fLogin.aspx"
LOG_PATH = r"C:\NFSeOsasco\osasco_log.txt"

PT_MESES = {1:"janeiro",2:"fevereiro",3:"março",4:"abril",5:"maio",6:"junho",7:"julho",8:"agosto",9:"setembro",10:"outubro",11:"novembro",12:"dezembro"}

# ======================= LOG / UTILS =======================
def _log_file(msg: str):
    try:
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(msg.rstrip()+"\n")
    except Exception:
        pass

def _report_error(e: Exception, context: str = ""):
    msg = f"❌ Erro {('em '+context) if context else ''}: {e.__class__.__name__} - {str(e) or '(sem mensagem)'}"
    print(msg, flush=True)
    _log_file("\n"+msg+"\n"+traceback.format_exc())

def calc_intervalo_mes_anterior():
    hoje = date.today()
    inicio = (hoje.replace(day=1) - relativedelta(months=1))
    fim    = (hoje.replace(day=1) - timedelta(days=1))
    return inicio, fim

def _norm(txt):
    if txt is None: return ""
    t = unicodedata.normalize("NFKD", txt)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return t.strip()

# ======================= DRIVER =======================
def setup_driver(download_dir):
    opts = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": False,
        "safebrowsing.for_trusted_sources_enabled": False,
    }
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--start-maximized")
    opts.add_argument("--safebrowsing-disable-download-protection")
    opts.add_argument("--disable-features=DownloadBubble")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    try:
        driver.execute_cdp_cmd("Page.enable", {})
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": download_dir})
    except Exception:
        pass
    return driver

def aguardar_login_manual(driver):
    driver.get(URL_LOGIN)
    print("\n🔐 Faça LOGIN; quando cair na Home eu continuo. (ENTER também funciona)")
    start = time.time()
    while True:
        try:
            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.LINK_TEXT, "Notas Fiscais")))
            print("✅ Login detectado (elementos da Home)."); return
        except Exception:
            pass
        if os.name=="nt":
            import msvcrt
            if msvcrt.kbhit() and msvcrt.getwch()=="\r":
                print("➡️  Prosseguindo por ENTER."); return
        if time.time()-start>600:
            print("⚠️ Tempo esgotado (10 min). Prosseguindo."); return
        time.sleep(1)

# ======================= HELPERS UI =======================
def _has_visible_overlay(driver):
    for css in (".ui-widget-overlay", ".modal-backdrop", ".blockUI"):
        try:
            el = driver.find_element(By.CSS_SELECTOR, css)
            if el.is_displayed(): return True
        except Exception:
            pass
    return False

def _esperar_overlay_sumir(driver, timeout=10):
    end = time.time()+timeout
    while time.time()<end:
        if not _has_visible_overlay(driver): return True
        time.sleep(0.2)
    try:
        driver.execute_script("document.querySelectorAll('.ui-widget-overlay,.modal-backdrop,.blockUI').forEach(e=>e.remove());")
    except Exception: pass
    return True

def _fechar_todos_os_modais(driver, max_loops=4):
    fechados = 0
    for _ in range(max_loops):
        try:
            modal = WebDriverWait(driver, 1.5).until(
                EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'ui-dialog') and contains(@class,'ui-widget') and not(contains(@style,'display: none'))]"))
            )
            try:
                btn = modal.find_element(By.XPATH, ".//button[contains(.,'Fechar') or contains(.,'OK') or contains(.,'Ok')] | .//input[@type='button' and (contains(@value,'Fechar') or contains(@value,'OK') or contains(@value,'Ok'))]")
            except Exception:
                btn = modal.find_element(By.XPATH, ".//a[contains(@class,'ui-dialog-titlebar-close')]")
            driver.execute_script("arguments[0].click();", btn)
            _esperar_overlay_sumir(driver, 6)
            fechados += 1
            time.sleep(0.2)
        except Exception:
            break
    if fechados: print(f"ℹ️ Fechei {fechados} janela(s) de alerta.")
    return fechados > 0

def _force_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try: el.click()
    except Exception: driver.execute_script("arguments[0].click();", el)

def _go_home(driver):
    try:
        home = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.LINK_TEXT,"Início")))
        _force_click(driver, home)
        _esperar_overlay_sumir(driver, 5)
    except Exception: pass

def _obter_nome_empresa(driver, default="empresa"):
    try:
        cands = driver.find_elements(By.XPATH, "//*[contains(normalize-space(.), 'Contribuinte:')]")
        texto = ""
        if cands:
            texto = max((c.text for c in cands), key=lambda t: len(t or ""), default="")
        if not texto:
            texto = driver.find_element(By.TAG_NAME, "body").text
        m = re.search(r"Contribuinte:\s*(.+?)(?:\s+(?:CPF\/?CNPJ|CNPJ|CPF)\b|$)", texto, flags=re.IGNORECASE | re.DOTALL)
        if m:
            nome = re.sub(r'[\\/:*?"<>|]+', "_", m.group(1).strip())
            if nome: return nome
    except Exception: pass
    return default

# ======================= DOWNLOADS =======================
def _wait_new_download(download_dir, before_set, expect_exts=(".pdf",".xml",".zip"), timeout=150):
    end = time.time()+timeout
    last_size = {}
    while time.time()<end:
        atual = {f for f in os.listdir(download_dir) if not f.endswith(".crdownload")}
        novos = [f for f in atual - before_set if f.lower().endswith(expect_exts)]
        if novos:
            fn = max(novos, key=lambda n: os.path.getmtime(os.path.join(download_dir, n)))
            fp = os.path.join(download_dir, fn)
            size = os.path.getsize(fp)
            if fn in last_size and last_size[fn]==size:
                return fn
            last_size[fn] = size
        time.sleep(0.35)
    return None

def _rename_with_retry(orig_full, dest_full, attempts=16, wait=0.5):
    for _ in range(attempts):
        try:
            os.replace(orig_full, dest_full)
            return True
        except Exception:
            time.sleep(wait)
    return False

# ======================= EXPORTAÇÃO (PDF/XML) =======================
def _preencher_input(elem, texto):
    elem.click(); elem.send_keys(Keys.CONTROL, "a"); elem.send_keys(Keys.DELETE)
    time.sleep(0.05); elem.send_keys(texto); time.sleep(0.05); elem.send_keys(Keys.TAB); time.sleep(0.05)

def _formatar_para_input(elem, data_obj):
    return data_obj.strftime("%d/%m/%Y")

def _preencher_datas_e_horas(driver, dt_ini, dt_fim):
    campo_ini = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,"//input[contains(@id,'DataInicial') or contains(@id,'txtDataInicial') or contains(@id,'dtInicial')]")))
    campo_fim = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,"//input[contains(@id,'DataFinal') or contains(@id,'txtDataFinal') or contains(@id,'dtFinal')]")))
    _preencher_input(campo_ini, _formatar_para_input(campo_ini, dt_ini))
    _preencher_input(campo_fim, _formatar_para_input(campo_fim, dt_fim))
    for xp,val in [
        ("//input[contains(@id,'HoraInicial') or contains(@id,'txtHoraInicial') or contains(@id,'HoraIni')]", "00:00"),
        ("//input[contains(@id,'HoraFinal') or contains(@id,'txtHoraFinal') or contains(@id,'HoraFim')]", "23:59"),
    ]:
        try: _preencher_input(driver.find_element(By.XPATH,xp), val)
        except Exception: pass

def _abrir_tela_exportacao(driver):
    _esperar_overlay_sumir(driver, 8); _fechar_todos_os_modais(driver)
    _force_click(driver, WebDriverWait(driver,30).until(EC.presence_of_element_located((By.LINK_TEXT, "Notas Fiscais"))))
    _esperar_overlay_sumir(driver, 3)
    try:
        el = WebDriverWait(driver,15).until(EC.presence_of_element_located((By.LINK_TEXT, "Exportar Notas para Arquivo")))
    except Exception:
        el = WebDriverWait(driver,15).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Exportar Notas")))
    _force_click(driver, el)
    _esperar_overlay_sumir(driver, 3); _fechar_todos_os_modais(driver)

def _mark_radio_exact(driver, label_text):
    xp = f"//input[@type='radio' and (following-sibling::*[contains(.,'{label_text}')])]"
    r = WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.XPATH, xp)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", r)
    try: r.click()
    except Exception: driver.execute_script("arguments[0].click();", r)
    try:
        driver.execute_script("arguments[0].checked=true; arguments[0].dispatchEvent(new Event('change',{bubbles:true}))", r)
    except Exception: pass
    time.sleep(0.15)
    print(f"   🔘 Marcado: {label_text}")

def _abrir_exportar_e_gerar(driver, dt_ini, dt_fim, considerar, nome_empresa):
    _abrir_tela_exportacao(driver)
    try: _mark_radio_exact(driver, "Data de Emissão")
    except Exception: pass
    alvo_cons = "Emitidas pela minha Empresa" if considerar=="emitidas" else "Recebidas pela minha Empresa"
    try: _mark_radio_exact(driver, alvo_cons)
    except Exception: pass
    _preencher_datas_e_horas(driver, dt_ini, dt_fim)

    sufixo = "notas emitidas" if considerar=="emitidas" else "notas recebidas"

    # PDF
    try: _mark_radio_exact(driver, "PDF")
    except Exception: pass
    before = set(os.listdir(DOWNLOAD_DIR))
    _force_click(driver, WebDriverWait(driver,30).until(EC.element_to_be_clickable((By.XPATH,"//input[@type='submit' and (contains(@value,'Gerar Arquivo') or contains(@value,'Gerar'))]"))))
    if _fechar_todos_os_modais(driver):
        print(f"📄 {considerar.upper()} (PDF): sem notas no período.")
    else:
        arq = _wait_new_download(DOWNLOAD_DIR, before, (".pdf",".zip"), timeout=150)
        if arq:
            src = os.path.join(DOWNLOAD_DIR, arq)
            dest = os.path.join(DOWNLOAD_DIR, f"{nome_empresa}_{sufixo}{os.path.splitext(arq)[1].lower()}")
            if _rename_with_retry(src, dest):
                print(f"📥 PDF salvo: {os.path.basename(dest)}")
            else:
                print(f"📥 PDF gerado: {arq} (não consegui renomear)")
        else:
            print("⚠️ Solicitei PDF, mas não detectei download.")

    # XML
    try: _mark_radio_exact(driver, "XML")
    except Exception: pass
    before = set(os.listdir(DOWNLOAD_DIR))
    _force_click(driver, WebDriverWait(driver,30).until(EC.element_to_be_clickable((By.XPATH,"//input[@type='submit' and (contains(@value,'Gerar Arquivo') or contains(@value,'Gerar'))]"))))
    if _fechar_todos_os_modais(driver):
        print(f"🗂 {considerar.upper()} (XML): sem notas no período.")
    else:
        arq = _wait_new_download(DOWNLOAD_DIR, before, (".xml",".zip"), timeout=180)
        if arq:
            src = os.path.join(DOWNLOAD_DIR, arq)
            dest = os.path.join(DOWNLOAD_DIR, f"{nome_empresa}_{sufixo}{os.path.splitext(arq)[1].lower()}")
            if _rename_with_retry(src, dest):
                print(f"📥 XML salvo: {os.path.basename(dest)}")
            else:
                print(f"📥 XML gerado: {arq} (não consegui renomear)")
        else:
            print("⚠️ Solicitei XML, mas o navegador pode ter bloqueado.")

# ======================= LIVRO FISCAL (PDF) =======================
def _abrir_livro_fiscal(driver):
    _go_home(driver)
    _esperar_overlay_sumir(driver, 8); _fechar_todos_os_modais(driver)
    print("📄 Abrindo Relatórios → Livro Fiscal…")
    rel = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.LINK_TEXT, "Relatórios")))
    actions = ActionChains(driver)
    try:
        actions.move_to_element(rel).pause(0.4).perform()
        item = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, "//a[contains(.,'Livro Fiscal')]")))
        _force_click(driver, item)
    except Exception:
        _force_click(driver, rel)
        item = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Livro Fiscal")))
        _force_click(driver, item)
    _esperar_overlay_sumir(driver, 3); _fechar_todos_os_modais(driver)

def _select_option_by_text_flexible(sel: Select, alvo_texto: str, numero_mes: int | None = None):
    alvo_norm = _norm(alvo_texto).lower()
    for t in {alvo_texto, alvo_texto.title(), alvo_texto.upper(), alvo_texto.lower()}:
        try: sel.select_by_visible_text(t); return True
        except Exception: pass
    for t in {alvo_texto, alvo_texto.title(), alvo_texto.upper(), alvo_texto.lower()}:
        try: sel.select_by_value(t); return True
        except Exception: pass
    for i, opt in enumerate(sel.options):
        if _norm(opt.text).lower().find(alvo_norm) >= 0:
            sel.select_by_index(i); return True
    if numero_mes:
        for t in {str(numero_mes), f"{numero_mes:02d}"}:
            try: sel.select_by_value(t); return True
            except Exception: pass
    return False

def _selecionar_exercicio_mes(driver, ano, mes_num):
    sel_ano = Select(WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,"//select[contains(@id,'Exercicio') or contains(@name,'Exercicio') or contains(@id,'ddlExercicio')]"))))
    _select_option_by_text_flexible(sel_ano, str(ano))
    sel_mes = Select(WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,"//select[contains(@id,'Mes') or contains(@name,'Mes') or contains(@id,'ddlMes')]"))))
    alvo = PT_MESES.get(mes_num, "").title()
    try:
        idx = max(0, sel_mes.all_selected_options[0].index - 1)
        sel_mes.select_by_index(idx); time.sleep(0.1)
    except Exception:
        pass
    _select_option_by_text_flexible(sel_mes, alvo, numero_mes=mes_num)

def _mark_radio_exact(driver, label_text):
    xp = f"//input[@type='radio' and (following-sibling::*[contains(.,'{label_text}')])]"
    r = WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.XPATH, xp)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", r)
    try: r.click()
    except Exception: driver.execute_script("arguments[0].click();", r)
    try:
        driver.execute_script("arguments[0].checked=true; arguments[0].dispatchEvent(new Event('change',{bubbles:true}))", r)
    except Exception: pass
    time.sleep(0.15)
    print(f"   🔘 Marcado: {label_text}")

def _gerar_livro(driver, ano, mes_num, tipo_label, nome_final):
    _abrir_livro_fiscal(driver)
    _selecionar_exercicio_mes(driver, ano, mes_num)

    _mark_radio_exact(driver, tipo_label)
    try: _mark_radio_exact(driver, "PDF")
    except Exception: pass

    main = driver.current_window_handle
    existentes = set(driver.window_handles)
    before = set(os.listdir(DOWNLOAD_DIR))

    btn = WebDriverWait(driver,30).until(EC.element_to_be_clickable((By.XPATH,"//input[@type='submit' and (contains(@value,'Gerar') or contains(@id,'Gerar'))] | //button[contains(.,'Gerar')]")))
    _force_click(driver, btn)
    time.sleep(0.7)

    new = None
    try:
        WebDriverWait(driver,5).until(lambda d: len(d.window_handles) > len(existentes))
        atuais = set(driver.window_handles); dif = list(atuais-existentes)
        if dif: new = dif[0]
    except Exception:
        pass

    arq = None
    if new:
        driver.switch_to.window(new)
        before2 = set(os.listdir(DOWNLOAD_DIR))
        arq = _wait_new_download(DOWNLOAD_DIR, before2, (".pdf",), timeout=120)
        try: driver.close()
        except: pass
        try: driver.switch_to.window(main)
        except: driver.switch_to.window(driver.window_handles[0])
    else:
        arq = _wait_new_download(DOWNLOAD_DIR, before, (".pdf",), timeout=40)

    if arq:
        src = os.path.join(DOWNLOAD_DIR, arq)
        dest = os.path.join(DOWNLOAD_DIR, nome_final)
        if _rename_with_retry(src, dest):
            print(f"📄 Livro salvo: {os.path.basename(dest)}")
        else:
            print(f"⚠️ Livro baixado como {arq}, não consegui renomear para {nome_final}")
    else:
        print("⚠️ Não consegui gerar/baixar este Livro.")

# ======================= GUIA ISS (Emitidos) — com suporte a iframe =======================
def _abrir_guia_emitidos(driver):
    _go_home(driver)
    _esperar_overlay_sumir(driver, 8); _fechar_todos_os_modais(driver)
    print("💸 Abrindo Pagamentos → Gerar Guias ISS → para Doctos. Emitidos…")
    pag = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.LINK_TEXT, "Pagamentos")))
    actions = ActionChains(driver)
    href_final = None
    try:
        actions.move_to_element(pag).pause(0.4).perform()
        gerar = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, "//a[contains(.,'Gerar Guias ISS')]")))
        actions.move_to_element(gerar).pause(0.3).perform()
        item = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, "//a[contains(.,'para Doctos. Emitidos')]")))
        href_final = item.get_attribute("href") or None
        _force_click(driver, item)
    except Exception:
        # fallback: clica sequencialmente
        _force_click(driver, pag)
        try:
            _force_click(driver, WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, "//a[contains(.,'Gerar Guias ISS')]"))))
            sub = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, "//a[contains(.,'para Doctos. Emitidos')]")))
            href_final = sub.get_attribute("href") or None
            _force_click(driver, sub)
        except Exception:
            pass
    if href_final and "http" in href_final:
        driver.get(href_final)
    _esperar_overlay_sumir(driver, 6); _fechar_todos_os_modais(driver)

def _maybe_switch_to_guia_iframe(driver):
    # volta ao topo e tenta entrar em seções/iframes típicos da página de guia
    try: driver.switch_to.default_content()
    except Exception: pass
    # múltiplas tentativas de iframes comuns
    cand_xps = [
        "//iframe[contains(@src,'Pagamento') or contains(@src,'Guia') or contains(@src,'Guias') or contains(@src,'Doctos')]",
        "//iframe[contains(@id,'frame') or contains(@name,'frame')]",
        "//iframe"
    ]
    for xp in cand_xps:
        try:
            iframe = WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.XPATH, xp)))
            driver.switch_to.frame(iframe)
            # validamos se dentro há “Exerc” ou “Mês” ou “Pesquisar”
            try:
                WebDriverWait(driver, 2).until(EC.presence_of_element_located((
                    By.XPATH, "//*[contains(.,'Exerc') or contains(.,'Mês') or contains(.,'Pesquisar') or contains(.,'Imprimir')]"
                )))
                return True
            except Exception:
                driver.switch_to.default_content()
        except Exception:
            pass
    return False

def g__achar_select_exercicio(driver):
    # tentar dentro ou fora de iframe
    _maybe_switch_to_guia_iframe(driver)
    xps = [
        "//label[contains(.,'Exerc')]/following::select[1]",
        "//span[contains(.,'Exerc')]/following::select[1]",
        "//td[contains(.,'Exerc')]/following::select[1]",
        "//select[contains(@id,'Exercicio') or contains(@name,'Exercicio')]",
    ]
    for xp in xps:
        try:
            el = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, xp)))
            if el.tag_name.lower()=="select": return el
        except Exception: pass
    # plano C: pega o 1º select da página visível
    return WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.XPATH, "(//select)[1]")))

def g__achar_select_mes(driver):
    # manter no mesmo contexto (possível iframe)
    xps = [
        "//label[contains(.,'Mês') or contains(.,'Mes')]/following::select[1]",
        "//span[contains(.,'Mês') or contains(.,'Mes')]/following::select[1]",
        "//td[contains(.,'Mês') or contains(.,'Mes')]/following::select[1]",
        "//select[contains(@id,'Mes') or contains(@name,'Mes')]",
    ]
    for xp in xps:
        try:
            el = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, xp)))
            if el.tag_name.lower()=="select": return el
        except Exception: pass
    # plano C: pega o 2º select
    try:
        return WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, "(//select)[2]")))
    except Exception:
        # ou volta pro 1º se só existir um
        return WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, "(//select)[1]")))

def g__select_text_flex(sel: Select, alvo_txt: str, numero_mes: int | None = None):
    alvo_n = _norm(alvo_txt).lower()
    for t in (alvo_txt, alvo_txt.title(), alvo_txt.upper(), alvo_txt.lower()):
        try: sel.select_by_visible_text(t); return True
        except Exception: pass
    for t in (alvo_txt, alvo_txt.title(), alvo_txt.upper(), alvo_txt.lower()):
        try: sel.select_by_value(t); return True
        except Exception: pass
    for i,opt in enumerate(sel.options):
        if _norm(opt.text).lower().find(alvo_n) >= 0:
            sel.select_by_index(i); return True
    if numero_mes:
        for v in (str(numero_mes), f"{numero_mes:02d}"):
            try: sel.select_by_value(v); return True
            except Exception: pass
    return False

def _confirm_guia_context(driver):
    # tenta achar um título/indicador de que estamos na tela correta
    marcadores = [
        "Gerar Guias ISS", "para Doctos. Emitidos", "Imprimir", "Pesquisar", "Pagamento"
    ]
    try:
        body = driver.find_element(By.TAG_NAME, "body").text
        if any(m.lower() in _norm(body).lower() for m in marcadores):
            return True
    except Exception:
        pass
    return False

def g_gerar_guia(driver, ano, mes_num, mes_nome, nome_empresa):
    _abrir_guia_emitidos(driver)
    _esperar_overlay_sumir(driver, 6); _fechar_todos_os_modais(driver)

    # às vezes abrir o submenu não troca o frame; certifica o contexto
    if not _confirm_guia_context(driver):
        _maybe_switch_to_guia_iframe(driver)

    # plano B: se ainda não estamos no contexto, recarregue pelo último href capturado (quando disponível)
    # (já fizemos isso em _abrir_guia_emitidos quando havia href)

    print("🗂 Selecionando período…", flush=True)
    try:
        sel_ano = Select(g__achar_select_exercicio(driver))
        sel_mes = Select(g__achar_select_mes(driver))
    except TimeoutException:
        # tenta mais uma vez: volta ao topo, entra no possível iframe e procura de novo
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        _maybe_switch_to_guia_iframe(driver)
        sel_ano = Select(g__achar_select_exercicio(driver))
        sel_mes = Select(g__achar_select_mes(driver))

    # debug suave: quantos selects existem
    try:
        all_selects = driver.find_elements(By.TAG_NAME, "select")
        print(f"   (debug) selects visíveis: {len(all_selects)}", flush=True)
    except Exception:
        pass

    ok_ano = g__select_text_flex(sel_ano, str(ano))
    ok_mes = g__select_text_flex(sel_mes, mes_nome, numero_mes=mes_num)
    if not ok_ano or not ok_mes:
        print("🟡 Não consegui selecionar Exercício/Mês para guia (seguindo)."); return
    print(f"   ✔ Ano={ano}  Mês={mes_nome}", flush=True)

    print("🔎 Clicando em Pesquisar…", flush=True)
    try:
        btn_pesq = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
            "//input[@type='submit' and @value='Pesquisar'] | //button[normalize-space()='Pesquisar'] | //a[normalize-space()='Pesquisar']"
        )))
        _force_click(driver, btn_pesq)
        _esperar_overlay_sumir(driver, 8)
    except Exception:
        pass

    btns = driver.find_elements(By.XPATH, "//input[@type='submit' and @value='Imprimir'] | //button[normalize-space()='Imprimir'] | //a[normalize-space()='Imprimir']")
    if not btns:
        print("🧾 Guia ISS (Emitidos): já emitida (sem botão 'Imprimir').")
        return

    print("🖨️ Abrindo PDF da Guia…", flush=True)
    main = driver.current_window_handle
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    existentes = set(driver.window_handles)
    try:
        _force_click(driver, btns[0])
    except Exception:
        try:
            # se o botão estava no iframe, tenta clicar novamente no mesmo contexto
            _maybe_switch_to_guia_iframe(driver)
            btns2 = driver.find_elements(By.XPATH, "//input[@type='submit' and @value='Imprimir'] | //button[normalize-space()='Imprimir'] | //a[normalize-space()='Imprimir']")
            if btns2:
                _force_click(driver, btns2[0])
        except Exception:
            pass
    _esperar_overlay_sumir(driver, 4)

    nova = None
    try:
        WebDriverWait(driver, 12).until(lambda d: len(d.window_handles) > len(existentes))
        dif = list(set(driver.window_handles) - existentes)
        if dif: nova = dif[0]
    except Exception: pass
    if nova:
        driver.switch_to.window(nova)

    # aguarda download
    before = {f for f in os.listdir(DOWNLOAD_DIR) if not f.endswith(".crdownload")}
    end = time.time()+120
    final = None
    while time.time() < end:
        atuais = {f for f in os.listdir(DOWNLOAD_DIR) if not f.endswith(".crdownload")}
        novos = [f for f in (atuais-before) if f.lower().endswith(".pdf")]
        if novos:
            final = max(novos, key=lambda n: os.path.getmtime(os.path.join(DOWNLOAD_DIR, n))); break
        time.sleep(0.4)

    if final:
        src = os.path.join(DOWNLOAD_DIR, final)
        dest = os.path.join(DOWNLOAD_DIR, f"{nome_empresa}_Guia ISS Prestados.pdf")
        if _rename_with_retry(src, dest):
            print(f"🧾 Guia ISS salva: {os.path.basename(dest)}", flush=True)
        else:
            print(f"🧾 Guia ISS baixada como {final} (não consegui renomear).", flush=True)
    else:
        print("⚠️ Não detectei o download do PDF da Guia ISS.", flush=True)

    if nova:
        try: driver.close()
        except: pass
        try: driver.switch_to.window(main)
        except: driver.switch_to.window(driver.window_handles[0])

# ======================= MAIN =======================
def main():
    dt_ini, dt_fim = calc_intervalo_mes_anterior()
    print(f"🗓️ Período (mês anterior): {dt_ini.strftime('%d/%m/%Y')} a {dt_fim.strftime('%d/%m/%Y')}")
    driver = setup_driver(DOWNLOAD_DIR)
    try:
        aguardar_login_manual(driver)
        _esperar_overlay_sumir(driver, 8); _fechar_todos_os_modais(driver)

        nome_empresa = _obter_nome_empresa(driver, "empresa")
        print(f"🏷️ Contribuinte detectado: {nome_empresa}")

        # 1) Exportações
        try:
            _abrir_exportar_e_gerar(driver, dt_ini, dt_fim, "emitidas",  nome_empresa)
            _abrir_exportar_e_gerar(driver, dt_ini, dt_fim, "recebidas", nome_empresa)
        except Exception as e:
            _report_error(e, "Exportação de Notas")

        # 2) Livros (mês/ano do período)
        ano, mes = dt_ini.year, dt_ini.month
        try:
            _gerar_livro(driver, ano, mes, "Notas Fiscais Emitidas",  f"{nome_empresa}_Livro Notas Emitidas.pdf")
            _gerar_livro(driver, ano, mes, "Notas Fiscais Recebidas", f"{nome_empresa}_Livro Notas Recebidas.pdf")
        except Exception as e:
            _report_error(e, "Livro Fiscal")

        # 3) Guia ISS — usa o MESMO mês/ano do período (mês anterior)
        try:
            mes_nome = PT_MESES.get(mes, "janeiro").title()
            g_gerar_guia(driver, ano, mes, mes_nome, nome_empresa)
        except Exception as e:
            _report_error(e, "Guia ISS (Emitidos)")

        print("\n✅ Fluxo concluído. Verifique a pasta Downloads.")
    except Exception as e:
        _report_error(e, "Fluxo principal")
    finally:
        try: driver.quit()
        except: pass

if __name__ == "__main__":
    main()
