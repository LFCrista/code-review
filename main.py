import os
import re
import sys
import time
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
import subprocess
from playwright.sync_api import sync_playwright
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import asyncio
from bs4 import BeautifulSoup, NavigableString

CHROME_PATH = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
CHROME_USER_DATA_DIR = r"C:\\temp\\chrome"
CHROME_REMOTE_DEBUGGING_PORT = 9222
CHROME_DEBUG_URL = f"http://localhost:{CHROME_REMOTE_DEBUGGING_PORT}"

SCOPES = ['https://www.googleapis.com/auth/documents']

if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# Global para armazenar link do GPT personalizado
link_gpt_custom = ""

# --- Autenticação Google Docs ---
def autenticar_google_docs():
    creds = None
    if os.path.exists('token.json'):
        try:
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        except Exception:
            os.remove('token.json')
            return autenticar_google_docs()
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists("credentials.json"):
                caminho = filedialog.askopenfilename(title="Selecione o arquivo credentials.json", filetypes=[("JSON Files", "*.json")])
                if not caminho:
                    messagebox.showerror("Erro", "Arquivo credentials.json não selecionado.")
                    return None
                os.rename(caminho, "credentials.json")
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def trocar_credentials():
    if os.path.exists("token.json"):
        os.remove("token.json")
    if os.path.exists("credentials.json"):
        os.remove("credentials.json")
    messagebox.showinfo("Troca de credenciais", "As credenciais foram removidas. Novas serão solicitadas no próximo envio.")

def extrair_document_id(link):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', link)
    return match.group(1) if match else None

def processar_html_com_estilos(html):
    soup = BeautifulSoup(html, 'html.parser')
    texto_final = ""
    estilos = []

    def adicionar_estilo(estilo, inicio, fim):
        estilos.append((estilo, True, inicio, fim))

    def processar_no(no, estilos_ativos):
        nonlocal texto_final
        if isinstance(no, NavigableString):
            ini = len(texto_final)
            texto_final += str(no)
            fim = len(texto_final)
            for estilo in estilos_ativos:
                adicionar_estilo(estilo, ini, fim)
        elif hasattr(no, 'children'):
            novos_estilos = estilos_ativos[:]
            if no.name in ['strong', 'b']:
                novos_estilos.append('bold')
            if no.name in ['em', 'i']:
                novos_estilos.append('italic')
            if no.name == 'br':
                texto_final += '\n'
                return
            for filho in no.children:
                processar_no(filho, novos_estilos)
            if no.name == 'p':
                texto_final += '\n'

    for el in soup.children:
        processar_no(el, [])

    texto_final += '\n\n'
    return texto_final, estilos

def inserir_no_google_docs(service, doc_id, index, titulo, html):
    texto, estilos = processar_html_com_estilos(html)
    requests = [
        {'insertText': {'location': {'index': index}, 'text': titulo + "\n"}},
        {
            'updateParagraphStyle': {
                'range': {'startIndex': index, 'endIndex': index + len(titulo) + 1},
                'paragraphStyle': {'namedStyleType': 'HEADING_1'},
                'fields': 'namedStyleType'
            }
        },
        {'insertText': {'location': {'index': index + len(titulo) + 1}, 'text': texto}}
    ]

    for estilo, valor, start, end in estilos:
        req = {
            'updateTextStyle': {
                'range': {'startIndex': index + len(titulo) + 1 + start, 'endIndex': index + len(titulo) + 1 + end},
                'textStyle': {},
                'fields': ''
            }
        }
        if estilo == 'bold':
            req['updateTextStyle']['textStyle']['bold'] = True
            req['updateTextStyle']['fields'] = 'bold'
        elif estilo == 'italic':
            req['updateTextStyle']['textStyle']['italic'] = True
            req['updateTextStyle']['fields'] = 'italic'
        requests.append(req)

    service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
    return len(titulo) + 1 + len(texto)

def esperar_resposta_gpt(page, tempo_maximo=180, intervalo_check=1.5, tempo_estavel=4):
    conteudo_anterior = ""
    tentativas_estaveis = 0
    tempo_decorrido = 0
    while tempo_decorrido < tempo_maximo:
        if page.url.endswith("/api/auth/error"):
            return "__ERRO_AUTENTICACAO__"
        if page.locator(".text-token-text-error").count() > 0:
            return "__ERRO_GPT__"
        if page.locator("button:has(svg[aria-label='Stop generating'])").is_visible():
            time.sleep(intervalo_check)
            tempo_decorrido += intervalo_check
            continue
        try:
            html = page.locator(".markdown").last.inner_html()
        except Exception:
            html = ""
        if html == conteudo_anterior:
            tentativas_estaveis += 1
        else:
            tentativas_estaveis = 0
        if tentativas_estaveis >= tempo_estavel:
            return html if html else "__ERRO_GPT__"
        conteudo_anterior = html
        time.sleep(intervalo_check)
        tempo_decorrido += intervalo_check
    return html if html else "__ERRO_GPT__"

def abrir_chrome_debug():
    try:
        subprocess.Popen([
            CHROME_PATH,
            f"--remote-debugging-port={CHROME_REMOTE_DEBUGGING_PORT}",
            f"--user-data-dir={CHROME_USER_DATA_DIR}"
        ])
        messagebox.showinfo("Chrome", "Chrome com depuração iniciado.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao iniciar o Chrome:\n{e}")

def processar_pdfs():
    global link_gpt_custom
    arquivos = filedialog.askopenfilenames(title="Selecione os arquivos PDF", filetypes=[("PDF files", "*.pdf")])
    if not arquivos:
        return

    link_gpt_custom = simpledialog.askstring("GPT Personalizado", "Cole o link da sala do GPT personalizado:")
    if not link_gpt_custom:
        messagebox.showerror("Erro", "Link do GPT personalizado não informado.")
        return

    link_doc = simpledialog.askstring("Google Docs", "Cole o link do documento onde deseja salvar as respostas:")
    doc_id = extrair_document_id(link_doc)
    if not doc_id:
        messagebox.showerror("Erro", "Link do Google Docs inválido.")
        return

    prompt = campo_prompt.get().strip()
    if not prompt:
        messagebox.showerror("Erro", "O prompt não pode estar vazio.")
        return

    texto_log.delete("1.0", tk.END)
    texto_log.insert(tk.END, "Iniciando processamento...\n")
    janela.update()

    creds = autenticar_google_docs()
    service = build('docs', 'v1', credentials=creds)

    try:
        doc = service.documents().get(documentId=doc_id).execute()
        index = doc['body']['content'][-1]['endIndex'] - 1
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível acessar o documento:\n{e}")
        return

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CHROME_DEBUG_URL)
        context = browser.contexts[0]

        for i, arquivo in enumerate(arquivos):
            if i % 30 == 0:
                page = context.new_page()
                page.goto(link_gpt_custom)
                time.sleep(5)

            nome = os.path.basename(arquivo)
            texto_log.insert(tk.END, f"📄 Processando: {nome}\n")
            janela.update()

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                with open(arquivo, "rb") as f:
                    tmp.write(f.read())
                tmp_path = tmp.name

            if page.locator("input[type='file']").count() > 0:
                page.set_input_files("input[type='file']", tmp_path)
            else:
                page.click("button:has(svg[aria-label='Upload a file'])", timeout=5000)
                page.wait_for_selector("input[type='file']", timeout=5000)
                page.set_input_files("input[type='file']", tmp_path)

            time.sleep(4)
            page.keyboard.type(prompt)
            page.keyboard.press("Enter")

            resposta_html = esperar_resposta_gpt(page)

            if resposta_html.startswith("__ERRO"):
                texto_log.insert(tk.END, f"❌ Erro ao processar {nome}: {resposta_html}\n")
            else:
                texto_log.insert(tk.END, f"✅ {nome} processado. Enviando ao Google Docs...\n")
                try:
                    titulo_sala = f"{nome} - parte {i // 30 + 1}"
                    incremento = inserir_no_google_docs(service, doc_id, index, titulo_sala, resposta_html)
                    index += incremento
                    texto_log.insert(tk.END, f"📜 Inserido no Docs.\n")
                except Exception as e:
                    texto_log.insert(tk.END, f"❌ Falha ao enviar {nome} ao Docs: {e}\n")

            janela.update()

# Interface Tkinter
janela = tk.Tk()
janela.title("Automatizador de PDFs para ChatGPT")
janela.geometry("700x500")

frame_topo = tk.Frame(janela)
frame_topo.pack(pady=10)

botao_abrir_chrome = tk.Button(frame_topo, text="Abrir Chrome com Debug", command=abrir_chrome_debug)
botao_abrir_chrome.pack(side=tk.LEFT, padx=10)

botao_trocar_credentials = tk.Button(frame_topo, text="Trocar credenciais", command=trocar_credentials)
botao_trocar_credentials.pack(side=tk.LEFT, padx=10)

frame_prompt = tk.Frame(janela)
frame_prompt.pack(pady=5)

label_prompt = tk.Label(frame_prompt, text="Prompt a ser enviado:")
label_prompt.pack(side=tk.LEFT)

campo_prompt = tk.Entry(frame_prompt, width=80)
campo_prompt.pack(side=tk.LEFT, padx=5)

frame_botoes = tk.Frame(janela)
frame_botoes.pack(pady=10)

botao_processar = tk.Button(frame_botoes, text="Selecionar e Processar PDFs", command=processar_pdfs)
botao_processar.pack()

texto_log = scrolledtext.ScrolledText(janela, width=90, height=20)
texto_log.pack(pady=10)

janela.mainloop()
