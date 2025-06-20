import os, re, sys, time, tempfile, tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
import subprocess, asyncio
from playwright.sync_api import sync_playwright
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from bs4 import BeautifulSoup, NavigableString

# ----------- CONFIG -----------
CHROME_PATH = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
CHROME_USER_DATA_DIR = r"C:\\temp\\chrome"
CHROME_REMOTE_DEBUGGING_PORT = 9222
CHROME_DEBUG_URL = f"http://localhost:{CHROME_REMOTE_DEBUGGING_PORT}"
SCOPES = ["https://www.googleapis.com/auth/documents"]
if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# ----------- GOOGLE DOCS ------
def autenticar_google_docs():
    if os.path.exists("token.json"):
        try:
            return Credentials.from_authorized_user_file("token.json", SCOPES)
        except Exception:
            os.remove("token.json")
    cred = filedialog.askopenfilename(title="credentials.json", filetypes=[("JSON", "*.json")])
    if not cred: return None
    flow = InstalledAppFlow.from_client_secrets_file(cred, SCOPES)
    creds = flow.run_local_server(port=0)
    open("token.json","w").write(creds.to_json())
    return creds

def extrair_document_id(url): m=re.search(r"/d/([A-Za-z0-9-_]+)",url); return m.group(1) if m else None

# ----------- HTML ➜ texto -----
def processar_html(html):
    soup,txt,est = BeautifulSoup(html,"html.parser"),"",[]
    def add(e,i,f): est.append((e,i,f))
    def walk(n,a):
        nonlocal txt
        if isinstance(n,NavigableString):
            ini=len(txt); txt+=str(n); fim=len(txt)
            for e in a: add(e,ini,fim)
        else:
            novo=a+(["bold"] if n.name in("strong","b") else [])+(["italic"] if n.name in("em","i") else [])
            if n.name=="br": txt+="\n"; return
            for c in n.children: walk(c,novo)
            if n.name=="p": txt+="\n"
    for el in soup.children: walk(el,[])
    txt+="\n\n"; return txt,est

def inserir_no_docs(svc,doc,idx,title,html):
    texto,est=processar_html(html)
    req=[{"insertText":{"location":{"index":idx},"text":title+"\n"}},
         {"updateParagraphStyle":{"range":{"startIndex":idx,"endIndex":idx+len(title)+1},
                                  "paragraphStyle":{"namedStyleType":"HEADING_1"},"fields":"namedStyleType"}},
         {"insertText":{"location":{"index":idx+len(title)+1},"text":texto}}]
    for e,s,f in est:
        style={"bold":True} if e=="bold" else {"italic":True}
        req.append({"updateTextStyle":{"range":{"startIndex":idx+len(title)+1+s,
                                               "endIndex":idx+len(title)+1+f},
                                       "textStyle":style,"fields":e}})
    svc.documents().batchUpdate(documentId=doc,body={"requests":req}).execute()
    return len(title)+1+len(texto)

# ----------- PLAYWRIGHT -------
def _stop(page): return page.locator("button:has(svg[aria-label='Stop generating'])").count()
def _stream(page): return page.locator(".result-streaming, .animate-spin").count()
def _composer(page):
    chips  = page.locator("[data-testid='file-upload-preview']").count() \
            or page.locator("div[role='listitem']").count()
    if chips:
        return True
    # textarea pode nem existir no DOM quando o composer está vazio
    ta = page.locator("textarea")
    if ta.count():
        try:
            return bool(ta.input_value().strip())
        except Exception:
            return False
    return False


def aguardar_pronto(page, stab=3, buf=1.2):
    """espera: sem Stop, sem stream, composer limpo, html estável ≥ stab s"""
    last_cnt = page.locator("[data-message-author-role='assistant']").count()
    last_html = page.locator("[data-message-author-role='assistant']").nth(-1)\
                 .locator(".markdown").inner_html() if last_cnt else ""
    fase_buffer=False; t0=time.time()
    while True:
        busy=_stop(page) or _stream(page) or _composer(page)
        cnt = page.locator("[data-message-author-role='assistant']").count()
        html= page.locator("[data-message-author-role='assistant']").nth(-1)\
                 .locator(".markdown").inner_html() if cnt else ""
        mudou = busy or cnt!=last_cnt or html!=last_html
        if mudou:
            fase_buffer=False; t0=time.time(); last_cnt,cnt=cnt,cnt; last_html,html=html,html
        else:
            now=time.time()
            if not fase_buffer and now-t0>=buf:
                fase_buffer=True; t0=now
            elif fase_buffer and now-t0>=stab:
                return
        time.sleep(0.5)

def digitar_prompt(page,prompt):
    for sel in ("textarea","div[role='textbox']"):
        try:
            page.wait_for_selector(sel,timeout=2000)
            box=page.locator(sel).first
            if sel=="textarea": box.evaluate("n=>n.value=''"); box.fill(prompt)
            else: box.click(); box.evaluate("n=>n.innerText=''"); box.type(prompt)
            page.keyboard.press("Enter"); return
        except: pass
    page.keyboard.type(prompt); page.keyboard.press("Enter")

def enviar_pdf(page,pdf,prompt):
    aguardar_pronto(page)                                     # 1
    prev_cnt=page.locator("[data-message-author-role='assistant']").count()

    # upload
    with tempfile.NamedTemporaryFile(delete=False,suffix=".pdf") as tmp:
        tmp.write(open(pdf,"rb").read()); path=tmp.name
    if page.locator("input[type='file']").count():
        page.set_input_files("input[type='file']",path)
    else:
        page.click("button:has(svg[aria-label='Upload a file'])",timeout=5000)
        page.wait_for_selector("input[type='file']",timeout=5000)
        page.set_input_files("input[type='file']",path)

    digitar_prompt(page,prompt)                               # 2

    # aguarda nova bolha existir
    while page.locator("[data-message-author-role='assistant']").count()<=prev_cnt:
        time.sleep(0.5)

    aguardar_pronto(page)                                     # 3 finaliza bolha

    # garante que .markdown exista
    nova = page.locator("[data-message-author-role='assistant']").nth(-1)
    page.wait_for_selector("[data-message-author-role='assistant'] >> .markdown", timeout=0)
    return nova.locator(".markdown").inner_html()

# helper já mostrado antes — certifique-se de que está no arquivo
def esperar_html_estavel(locator, segundos=2.0, dt=0.4):
    tam = len(locator.inner_html(timeout=0))
    t0  = time.time()
    while True:
        time.sleep(dt)
        novo = len(locator.inner_html(timeout=0))
        if novo != tam:
            tam, t0 = novo, time.time()
        elif time.time() - t0 >= segundos:
            return
        
def esperar_markdown(bolha, dt: float = 0.4):
    """
    Aguarda indefinidamente até que exista .markdown dentro da bolha.
    Devolve o locator .markdown (que ainda pode crescer depois).
    """
    while True:
        md = bolha.locator(".markdown")
        if md.count():
            return md          # encontrado → sai do laço
        time.sleep(dt)         # espera e tenta novamente

def capturar_resposta_completa(page):
    """
    Aguarda .markdown da última bolha existir e estabilizar.
    """
    bolha = page.locator("[data-message-author-role='assistant']").nth(-1)
    mark  = bolha.locator(".markdown")
    mark.wait_for(state="attached")          # ← substitui wait_for_selector

    # estabiliza por 1,5 s sem crescimento de bytes
    tam = len(mark.inner_html(timeout=0))
    t0  = time.time()
    while True:
        time.sleep(0.4)
        novo = len(mark.inner_html(timeout=0))
        if novo != tam:
            tam, t0 = novo, time.time()
        elif time.time() - t0 >= 1.5:
            return mark.inner_html()



# --------------- SUBSTITUA somente esta função -----------------------
def processar():
    pdfs = filedialog.askopenfilenames(title="PDFs", filetypes=[("PDF", "*.pdf")])
    if not pdfs:
        return

    link_gpt = simpledialog.askstring("GPT",  "Link da sala GPT:")
    link_doc = simpledialog.askstring("Docs", "Link do documento:")
    prompt   = campo_prompt.get().strip()
    doc_id   = extrair_document_id(link_doc)
    if not all((link_gpt, doc_id, prompt)):
        return

    texto_log.delete("1.0", tk.END)
    texto_log.insert(tk.END, "Início\n"); janela.update()

    svc = build("docs", "v1", credentials=autenticar_google_docs())
    idx = svc.documents().get(documentId=doc_id).execute()["body"]["content"][-1]["endIndex"] - 1

    with sync_playwright() as p:
        ctx  = p.chromium.connect_over_cdp(CHROME_DEBUG_URL).contexts[0]
        page = None

        for i, pdf in enumerate(pdfs, 1):
            if (i - 1) % 30 == 0:
                page = ctx.new_page()
                page.goto(link_gpt)
                time.sleep(5)

            nome = os.path.basename(pdf)
            texto_log.insert(tk.END, f"➡ {nome}\n"); janela.update()

            # contador antes de enviar
            cnt_before = page.locator("[data-message-author-role='assistant']").count()

            # envio: upload + prompt + aguarda estrutura
            _ = enviar_pdf(page, pdf, prompt)

            # bolha recém-criada
            bolha = page.locator("[data-message-author-role='assistant']").nth(cnt_before)

            # espera .markdown existir (ou usa bolha inteira)
            md = esperar_markdown(bolha)

            # estabilidade do HTML
            esperar_html_estavel(md, segundos=2.0)

            html = md.inner_html()

            if html.strip():
                idx += inserir_no_docs(svc, doc_id, idx, nome, html)
                texto_log.insert(tk.END, "✓ OK\n")
            else:
                texto_log.insert(tk.END, "⚠ Sem resposta\n")

            # garante chat 100 % livre antes do próximo PDF
            aguardar_pronto(page)
            janela.update()


# ----------- Interface --------
janela=tk.Tk(); janela.title("Automatizador de PDFs"); janela.geometry("700x500")
top=tk.Frame(janela); top.pack(pady=10)
def abrir_debug():
    subprocess.Popen([CHROME_PATH,
                      f"--remote-debugging-port={CHROME_REMOTE_DEBUGGING_PORT}",
                      f"--user-data-dir={CHROME_USER_DATA_DIR}"])
tk.Button(top,text="Abrir Chrome Debug",command=abrir_debug).pack(side=tk.LEFT,padx=10)
tk.Button(top,text="Trocar credenciais",
          command=lambda:[os.remove(f) for f in ("token.json","credentials.json") if os.path.exists(f)]).pack(side=tk.LEFT)

mid=tk.Frame(janela); mid.pack(pady=5)
tk.Label(mid,text="Prompt:").pack(side=tk.LEFT)
campo_prompt=tk.Entry(mid,width=80); campo_prompt.pack(side=tk.LEFT,padx=5)

tk.Button(janela,text="Selecionar e Processar PDFs",command=processar).pack(pady=10)
texto_log=scrolledtext.ScrolledText(janela,width=90,height=20); texto_log.pack(pady=10)
janela.mainloop()
