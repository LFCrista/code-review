import os, re, sys, time, tempfile, tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
import subprocess, asyncio

from playwright.sync_api import sync_playwright, TimeoutError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from bs4 import BeautifulSoup, NavigableString

# --------------------------------------------------
# CONFIGURAÇÕES DO CHROME (modo debug)
# --------------------------------------------------
CHROME_PATH = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
CHROME_USER_DATA_DIR = r"C:\\temp\\chrome"
CHROME_REMOTE_DEBUGGING_PORT = 9222
CHROME_DEBUG_URL = f"http://localhost:{CHROME_REMOTE_DEBUGGING_PORT}"

SCOPES = ["https://www.googleapis.com/auth/documents"]

if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# --------------------------------------------------
#                GOOGLE DOCS UTILS
# --------------------------------------------------
def autenticar_google_docs():
    creds = None
    if os.path.exists("token.json"):
        try:
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        except Exception:
            os.remove("token.json")
            return autenticar_google_docs()

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists("credentials.json"):
                arq = filedialog.askopenfilename(
                    title="Selecione credentials.json", filetypes=[("JSON", "*.json")]
                )
                if not arq:
                    messagebox.showerror("Erro", "credentials.json não selecionado.")
                    return None
                os.rename(arq, "credentials.json")
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        open("token.json", "w").write(creds.to_json())
    return creds


def trocar_credentials():
    for f in ("token.json", "credentials.json"):
        if os.path.exists(f):
            os.remove(f)
    messagebox.showinfo("Credenciais", "Removidas — serão solicitadas novamente.")


def extrair_document_id(link: str) -> str | None:
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", link)
    return m.group(1) if m else None


# --------------------------------------------------
#            HTML → TEXTO / ESTILOS
# --------------------------------------------------
def processar_html_com_estilos(html: str):
    soup, txt, estilos = BeautifulSoup(html, "html.parser"), "", []

    def add(est, ini, fim): estilos.append((est, ini, fim))

    def walk(n, ativos):
        nonlocal txt
        if isinstance(n, NavigableString):
            ini, txt = len(txt), txt + str(n)
            for est in ativos:
                add(est, ini, len(txt))
        else:
            novos = ativos + (["bold"] if n.name in ("strong", "b") else []) + (
                ["italic"] if n.name in ("em", "i") else []
            )
            if n.name == "br":
                txt += "\n"
                return
            for c in n.children:
                walk(c, novos)
            if n.name == "p":
                txt += "\n"

    for el in soup.children:
        walk(el, [])
    txt += "\n\n"
    return txt, estilos


def inserir_no_google_docs(svc, doc_id, idx, titulo, html):
    texto, estilos = processar_html_com_estilos(html)
    reqs = [
        {"insertText": {"location": {"index": idx}, "text": titulo + "\n"}},
        {
            "updateParagraphStyle": {
                "range": {"startIndex": idx, "endIndex": idx + len(titulo) + 1},
                "paragraphStyle": {"namedStyleType": "HEADING_1"},
                "fields": "namedStyleType",
            }
        },
        {"insertText": {"location": {"index": idx + len(titulo) + 1}, "text": texto}},
    ]
    for est, s, e in estilos:
        style, field = ({"bold": True}, "bold") if est == "bold" else ({"italic": True}, "italic")
        reqs.append(
            {
                "updateTextStyle": {
                    "range": {
                        "startIndex": idx + len(titulo) + 1 + s,
                        "endIndex": idx + len(titulo) + 1 + e,
                    },
                    "textStyle": style,
                    "fields": field,
                }
            }
        )
    svc.documents().batchUpdate(documentId=doc_id, body={"requests": reqs}).execute()
    return len(titulo) + 1 + len(texto)


# --------------------------------------------------
#        CONTROLE DE SINCRONIZAÇÃO PLAYWRIGHT
# --------------------------------------------------
def _chat_ocupado(page) -> bool:
    """
    Retorna True se o ChatGPT ainda estiver:
      • streamando a resposta,
      • com botão “Stop generating” na tela,
      • ou com algum chip de anexo no composer.
    """
    return (
        page.locator(".result-streaming").count() > 0
        or page.locator("button:has(svg[aria-label='Stop generating'])").count() > 0
        or page.locator("[data-testid='file-upload-preview']").count() > 0
        or page.locator("div[role='listitem'] svg[aria-label='Document']").count() > 0
    )


def aguardar_chat_livre(page, estabilidade_seg=3):
    """
    Espera o chat ficar totalmente livre durante 'estabilidade_seg'
    segundos consecutivos.
    """
    t0 = time.time()
    while True:
        if _chat_ocupado(page):
            t0 = time.time()  # reinicia cronômetro
        if time.time() - t0 >= estabilidade_seg:
            return
        time.sleep(0.6)


# ---------------------- PROMPT ----------------------
def digitar_prompt(page, prompt: str):
    sel_textarea = "textarea"
    sel_div = "div[role='textbox']"
    try:
        page.wait_for_selector(sel_textarea, timeout=4000)
        tb = page.locator(sel_textarea).first
        tb.evaluate("n=>n.value=''")
        tb.fill(prompt)
        page.keyboard.press("Enter")
        return
    except Exception:
        pass

    try:
        page.wait_for_selector(sel_div, timeout=3000)
        box = page.locator(sel_div).first
        box.click()
        box.evaluate("n=>n.innerText=''")
        box.type(prompt)
        page.keyboard.press("Enter")
        return
    except Exception:
        pass

    page.keyboard.type(prompt)
    page.keyboard.press("Enter")


# ------------------ ENVIO/ESPERA --------------------
def enviar_arquivo_e_esperar(page, pdf, prompt):
    # 1) garante que o chat esteja 100 % livre
    aguardar_chat_livre(page)

    # 2) faz upload
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(open(pdf, "rb").read())

    if page.locator("input[type='file']").count():
        page.set_input_files("input[type='file']", tmp.name)
    else:
        page.click("button:has(svg[aria-label='Upload a file'])", timeout=5000)
        page.wait_for_selector("input[type='file']", timeout=5000)
        page.set_input_files("input[type='file']", tmp.name)

    # 3) envia prompt
    time.sleep(0.5)  # pequena pausa para o composer atualizar
    digitar_prompt(page, prompt)

    # 4) espera resposta concluir + chip sumir
    aguardar_chat_livre(page)

    # 5) captura HTML final
    html = page.locator(".markdown").last.inner_html()
    return html


# --------------------------------------------------
#                    CHROME
# --------------------------------------------------
def abrir_chrome_debug():
    try:
        subprocess.Popen(
            [
                CHROME_PATH,
                f"--remote-debugging-port={CHROME_REMOTE_DEBUGGING_PORT}",
                f"--user-data-dir={CHROME_USER_DATA_DIR}",
            ]
        )
        messagebox.showinfo("Chrome", "Chrome (debug) iniciado.")
    except Exception as e:
        messagebox.showerror("Erro", str(e))


# --------------------------------------------------
#                  PROCESSO PRINCIPAL
# --------------------------------------------------
def processar_pdfs():
    pdfs = filedialog.askopenfilenames(title="Selecione PDFs", filetypes=[("PDF", "*.pdf")])
    if not pdfs:
        return
    link_gpt = simpledialog.askstring("GPT", "Link da sala GPT:")
    if not link_gpt:
        return
    link_doc = simpledialog.askstring("Google Docs", "Link do documento:")
    doc_id = extrair_document_id(link_doc)
    if not doc_id:
        return
    prompt = campo_prompt.get().strip()
    if not prompt:
        return

    texto_log.delete("1.0", tk.END)
    texto_log.insert(tk.END, "Iniciando...\n")
    janela.update()

    svc = build("docs", "v1", credentials=autenticar_google_docs())
    try:
        idx = (
            svc.documents()
            .get(documentId=doc_id)
            .execute()["body"]["content"][-1]["endIndex"]
            - 1
        )
    except Exception as e:
        messagebox.showerror("Erro", f"Docs: {e}")
        return

    with sync_playwright() as p:
        ctx = p.chromium.connect_over_cdp(CHROME_DEBUG_URL).contexts[0]
        page = None
        for i, pdf in enumerate(pdfs, 1):
            if (i - 1) % 30 == 0:
                page = ctx.new_page()
                page.goto(link_gpt)
                time.sleep(5)

            nome = os.path.basename(pdf)
            texto_log.insert(tk.END, f"📄 {nome}\n")
            janela.update()

            html = enviar_arquivo_e_esperar(page, pdf, prompt)
            if html.startswith("__ERRO"):
                texto_log.insert(tk.END, f"❌ {html}\n")
            else:
                texto_log.insert(tk.END, "➡️  Docs...\n")
                idx += inserir_no_google_docs(svc, doc_id, idx, nome, html)
                texto_log.insert(tk.END, "✅ OK\n")
            janela.update()


# --------------------------------------------------
#                    TKINTER
# --------------------------------------------------
janela = tk.Tk()
janela.title("Automatizador de PDFs")
janela.geometry("700x500")

top = tk.Frame(janela)
top.pack(pady=10)
tk.Button(top, text="Abrir Chrome Debug", command=abrir_chrome_debug).pack(side=tk.LEFT, padx=10)
tk.Button(top, text="Trocar credenciais", command=trocar_credentials).pack(side=tk.LEFT)

mid = tk.Frame(janela)
mid.pack(pady=5)
tk.Label(mid, text="Prompt:").pack(side=tk.LEFT)
campo_prompt = tk.Entry(mid, width=80)
campo_prompt.pack(side=tk.LEFT, padx=5)

bt = tk.Frame(janela)
bt.pack(pady=10)
tk.Button(bt, text="Selecionar e Processar PDFs", command=processar_pdfs).pack()

texto_log = scrolledtext.ScrolledText(janela, width=90, height=20)
texto_log.pack(pady=10)

janela.mainloop()
