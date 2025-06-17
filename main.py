# Seu script completo atualizado com negrito e itálico funcionando corretamente:
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
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

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

def inserir_no_google_docs(service, doc_id, index, nome_arquivo, html):
    titulo = f"{nome_arquivo}\n"
    texto, estilos = processar_html_com_estilos(html)

    requests = [
        {'insertText': {'location': {'index': index}, 'text': titulo}},
        {
            'updateParagraphStyle': {
                'range': {'startIndex': index, 'endIndex': index + len(titulo)},
                'paragraphStyle': {'namedStyleType': 'HEADING_1'},
                'fields': 'namedStyleType'
            }
        },
        {'insertText': {'location': {'index': index + len(titulo)}, 'text': texto}}
    ]

    for estilo, valor, start, end in estilos:
        req = {
            'updateTextStyle': {
                'range': {'startIndex': index + len(titulo) + start, 'endIndex': index + len(titulo) + end},
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
    return len(titulo) + len(texto)

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

def enviar_pdf_para_gpt(page, caminho_pdf):
    tempo_maximo = 900
    intervalo = 1.5
    tempo_decorrido = 0
    while page.locator("div[role='listitem']").is_visible() and tempo_decorrido < tempo_maximo:
        time.sleep(intervalo)
        tempo_decorrido += intervalo
    if page.locator("div[role='listitem']").is_visible():
        return "__ARQUIVO_JA_ANEXADO__"
    try:
        if page.locator("input[type='file']").count() > 0:
            page.set_input_files("input[type='file']", caminho_pdf)
        else:
            page.click("button:has(svg[aria-label='Upload a file'])", timeout=5000)
            page.wait_for_selector("input[type='file']", timeout=5000)
            page.set_input_files("input[type='file']", caminho_pdf)
    except:
        return "__ERRO_ENVIO__"
    time.sleep(4)
    page.keyboard.type("R1 - T1")
    page.keyboard.press("Enter")
    return esperar_resposta_gpt(page)

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
    arquivos = filedialog.askopenfilenames(title="Selecione os arquivos PDF", filetypes=[("PDF files", "*.pdf")])
    if not arquivos:
        return

    texto_log.delete("1.0", tk.END)
    texto_log.insert(tk.END, "Iniciando processamento...\n")
    janela.update()

    link = simpledialog.askstring("Google Docs", "Cole o link do documento onde deseja salvar as respostas:")
    doc_id = extrair_document_id(link)
    if not doc_id:
        messagebox.showerror("Erro", "Link do Google Docs inválido.")
        return

    creds = autenticar_google_docs()
    service = build('docs', 'v1', credentials=creds)

    try:
        doc = service.documents().get(documentId=doc_id).execute()
        index = doc['body']['content'][-1]['endIndex'] - 1
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível acessar o documento:\n{e}")
        return

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp(CHROME_DEBUG_URL)
            context = browser.contexts[0]
            page = context.pages[0] if context.pages else context.new_page()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao conectar ao navegador Chrome com debug remoto:\n{e}")
            return

        for arquivo in arquivos:
            nome = os.path.basename(arquivo)
            texto_log.insert(tk.END, f"📄 Processando: {nome}\n")
            janela.update()

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                with open(arquivo, "rb") as f:
                    tmp.write(f.read())
                tmp_path = tmp.name

            resposta_html = enviar_pdf_para_gpt(page, tmp_path)

            if resposta_html.startswith("__ERRO"):
                texto_log.insert(tk.END, f"❌ Erro ao processar {nome}: {resposta_html}\n")
            else:
                texto_log.insert(tk.END, f"✅ {nome} processado com sucesso. Enviando ao Google Docs...\n")
                janela.update()
                try:
                    incremento = inserir_no_google_docs(service, doc_id, index, nome, resposta_html)
                    index += incremento
                    texto_log.insert(tk.END, f"📜 Inserido no Docs.\n")
                except Exception as e:
                    texto_log.insert(tk.END, f"❌ Falha ao enviar {nome} ao Docs: {e}\n")
            janela.update()

# Interface Tkinter
janela = tk.Tk()
janela.title("Automatizador de PDFs para ChatGPT")
janela.geometry("600x400")

frame_botoes = tk.Frame(janela)
frame_botoes.pack(pady=10)

botao_abrir_chrome = tk.Button(frame_botoes, text="Abrir Chrome com Debug", command=abrir_chrome_debug)
botao_abrir_chrome.pack(side=tk.LEFT, padx=10)

botao_processar = tk.Button(frame_botoes, text="Selecionar e Processar PDFs", command=processar_pdfs)
botao_processar.pack(side=tk.LEFT, padx=10)

texto_log = scrolledtext.ScrolledText(janela, width=80, height=20)
texto_log.pack(pady=10)

janela.mainloop()