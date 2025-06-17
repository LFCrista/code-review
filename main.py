# main.py
import sys
import os
import re
import time
import tempfile
import shutil
import asyncio
import subprocess
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QTextEdit, QFileDialog, QLineEdit, QLabel, QMessageBox, QHBoxLayout,
    QSlider, QGroupBox, QGridLayout
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from playwright.sync_api import sync_playwright
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from bs4 import BeautifulSoup, NavigableString

SCOPES = ['https://www.googleapis.com/auth/documents']
CHROME_PATH = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
CHROME_USER_DATA_DIR = r"C:\\temp\\chrome"
CHROME_REMOTE_DEBUGGING_PORT = 9222
CHROME_DEBUG_URL = f"http://localhost:{CHROME_REMOTE_DEBUGGING_PORT}"

if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF to GPT Automation")
        self.setMinimumSize(1000, 620)
        self.setWindowIcon(QIcon("icon.png"))

        self.prompt_input = QLineEdit()
        self.prompt_input.setPlaceholderText("Digite o prompt para o GPT")

        self.gpt_link_input = QLineEdit()
        self.gpt_link_input.setPlaceholderText("Link do GPT personalizado")

        self.docs_link_input = QLineEdit()
        self.docs_link_input.setPlaceholderText("Link do Google Docs")

        self.slider_lote = QSlider(Qt.Horizontal)
        self.slider_lote.setMinimum(10)
        self.slider_lote.setMaximum(50)
        self.slider_lote.setValue(30)
        self.slider_lote.valueChanged.connect(self.atualizar_valor_slider)

        self.label_slider_valor = QLabel("30")

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)

        self.botao_chrome = QPushButton("Abrir Chrome com Debug")
        self.botao_chrome.clicked.connect(self.abrir_chrome_debug)

        self.botao_credentials = QPushButton("Trocar credenciais")
        self.botao_credentials.clicked.connect(self.trocar_credentials)

        self.botao_upload = QPushButton("Selecionar PDFs")
        self.botao_upload.clicked.connect(self.processar_pdfs)

        self.botao_processar = QPushButton("Iniciar Automação")
        self.botao_processar.clicked.connect(self.processar_pdfs)
        self.botao_processar.setFixedHeight(40)

        layout = QVBoxLayout()

        header = QLabel("PDF to GPT Automation")
        header.setStyleSheet("font-size: 24px; font-weight: bold; color: #8c7ae6")
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)

        upload_group = QGroupBox("Upload de PDFs")
        upload_layout = QVBoxLayout()
        upload_layout.addWidget(self.botao_upload)
        upload_group.setLayout(upload_layout)

        config_group = QGroupBox("Configurações")
        config_layout = QGridLayout()
        config_layout.addWidget(QLabel("Prompt para GPT:"), 0, 0)
        config_layout.addWidget(self.prompt_input, 0, 1)
        config_layout.addWidget(QLabel("Link do GPT personalizado:"), 1, 0)
        config_layout.addWidget(self.gpt_link_input, 1, 1)
        config_layout.addWidget(QLabel("Link do Google Docs:"), 2, 0)
        config_layout.addWidget(self.docs_link_input, 2, 1)
        config_layout.addWidget(QLabel("Páginas por Lote:"), 3, 0)
        config_layout.addWidget(self.slider_lote, 3, 1)
        config_layout.addWidget(self.label_slider_valor, 3, 2)
        config_group.setLayout(config_layout)

        botoes_extra = QHBoxLayout()
        botoes_extra.addWidget(self.botao_chrome)
        botoes_extra.addWidget(self.botao_credentials)

        layout.addWidget(upload_group)
        layout.addWidget(config_group)
        layout.addLayout(botoes_extra)
        layout.addWidget(self.botao_processar)
        layout.addWidget(QLabel("Logs de Execução:"))
        layout.addWidget(self.log_output)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        self.aplicar_estilo()

    def aplicar_estilo(self):
        estilo = """QWidget { background-color: #f5f6fa; font-family: 'Segoe UI'; font-size: 10.5pt; }
        QLabel { color: #2f3640; }
        QLineEdit, QTextEdit {
            background-color: white; border: 1px solid #dcdde1;
            border-radius: 6px; padding: 5px;
        }
        QPushButton {
            background-color: #8c7ae6; color: white;
            border: none; border-radius: 5px; padding: 8px 12px;
        }
        QPushButton:hover { background-color: #9c88ff; }
        QSlider::groove:horizontal {
            border: 1px solid #dcdde1; height: 6px;
            background: #dcdde1; border-radius: 3px;
        }
        QSlider::handle:horizontal {
            background: #8c7ae6; border: 1px solid #8c7ae6;
            width: 14px; margin: -5px 0; border-radius: 7px;
        }
        QGroupBox {
            border: 1px solid #dcdde1; border-radius: 8px;
            padding: 10px; margin-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin; left: 10px;
            padding: 0 3px 0 3px; font-weight: bold;
        }
        """
        self.setStyleSheet(estilo)

    def atualizar_valor_slider(self):
        self.label_slider_valor.setText(str(self.slider_lote.value()))

    def abrir_chrome_debug(self):
        try:
            subprocess.Popen([
                CHROME_PATH,
                "--remote-debugging-port=9222",
                f"--user-data-dir={CHROME_USER_DATA_DIR}"
            ])
            QMessageBox.information(self, "Chrome", "Chrome com depuração iniciado.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao iniciar o Chrome:\n{e}")

    def trocar_credentials(self):
        if os.path.exists("token.json"):
            os.remove("token.json")
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecione o novo credentials.json", filter="JSON Files (*.json)")
        if caminho:
            destino = os.path.abspath("credentials.json")
            try:
                shutil.copyfile(caminho, destino)
                QMessageBox.information(self, "Credenciais", "Credenciais substituídas com sucesso.")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao copiar o arquivo:\n{e}")
        else:
            QMessageBox.warning(self, "Credenciais", "Nenhum arquivo selecionado.")

    def autenticar_google_docs(self):
        creds = None
        if os.path.exists('token.json'):
            try:
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            except Exception:
                os.remove('token.json')
                return self.autenticar_google_docs()
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists("credentials.json"):
                    QMessageBox.critical(self, "Erro", "Arquivo credentials.json não encontrado.")
                    return None
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return creds

    def extrair_document_id(self, link):
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', link)
        return match.group(1) if match else None

    def processar_html_com_estilos(self, html):
        soup = BeautifulSoup(html, 'html.parser')
        texto_final, estilos = "", []

        def adicionar_estilo(estilo, ini, fim):
            estilos.append((estilo, ini, fim))

        def processar(no, ativos):
            nonlocal texto_final
            if isinstance(no, NavigableString):
                ini = len(texto_final)
                texto_final += str(no)
                fim = len(texto_final)
                for estilo in ativos:
                    adicionar_estilo(estilo, ini, fim)
            elif hasattr(no, 'children'):
                novos = ativos[:]
                if no.name in ['strong', 'b']:
                    novos.append('bold')
                if no.name in ['em', 'i']:
                    novos.append('italic')
                if no.name == 'br':
                    texto_final += '\n'
                    return
                for filho in no.children:
                    processar(filho, novos)
                if no.name == 'p':
                    texto_final += '\n'

        for el in soup.children:
            processar(el, [])

        return texto_final + '\n\n', estilos

    def inserir_no_google_docs(self, service, doc_id, index, titulo, html):
        texto, estilos = self.processar_html_com_estilos(html)
        reqs = [
            {'insertText': {'location': {'index': index}, 'text': titulo + '\n'}},
            {'updateParagraphStyle': {
                'range': {'startIndex': index, 'endIndex': index + len(titulo) + 1},
                'paragraphStyle': {'namedStyleType': 'HEADING_1'},
                'fields': 'namedStyleType'
            }},
            {'insertText': {'location': {'index': index + len(titulo) + 1}, 'text': texto}}
        ]
        for estilo, ini, fim in estilos:
            style = {'bold': estilo == 'bold', 'italic': estilo == 'italic'}
            reqs.append({
                'updateTextStyle': {
                    'range': {'startIndex': index + len(titulo) + 1 + ini, 'endIndex': index + len(titulo) + 1 + fim},
                    'textStyle': style,
                    'fields': ','.join(style.keys())
                }
            })
        service.documents().batchUpdate(documentId=doc_id, body={'requests': reqs}).execute()
        return len(titulo) + 1 + len(texto)

    def esperar_resposta_gpt(self, page, tempo_max=180, intervalo=1.5, estavel=4):
        anterior = ""
        tentativas = 0
        decorrido = 0
        while decorrido < tempo_max:
            if page.url.endswith("/api/auth/error"):
                return "__ERRO_AUTENTICACAO__"
            if page.locator(".text-token-text-error").count() > 0:
                return "__ERRO_GPT__"
            if page.locator("button:has(svg[aria-label='Stop generating'])").is_visible():
                time.sleep(intervalo)
                decorrido += intervalo
                continue
            try:
                html = page.locator(".markdown").last.inner_html()
            except:
                html = ""
            if html == anterior:
                tentativas += 1
            else:
                tentativas = 0
            if tentativas >= estavel:
                return html if html else "__ERRO_GPT__"
            anterior = html
            time.sleep(intervalo)
            decorrido += intervalo
        return html if html else "__ERRO_GPT__"

    def processar_pdfs(self):
        arquivos, _ = QFileDialog.getOpenFileNames(self, "Selecione os arquivos PDF", filter="PDF Files (*.pdf)")
        if not arquivos:
            return

        prompt = self.prompt_input.text().strip()
        link_gpt = self.gpt_link_input.text().strip()
        link_docs = self.docs_link_input.text().strip()
        paginas_lote = self.slider_lote.value()

        if not prompt or not link_docs or not link_gpt:
            QMessageBox.warning(self, "Campos obrigatórios", "Preencha todos os campos obrigatórios.")
            return

        doc_id = self.extrair_document_id(link_docs)
        if not doc_id:
            QMessageBox.warning(self, "Erro", "Link do Google Docs inválido.")
            return

        creds = self.autenticar_google_docs()
        if not creds:
            return

        service = build('docs', 'v1', credentials=creds)
        doc = service.documents().get(documentId=doc_id).execute()
        index = doc['body']['content'][-1]['endIndex'] - 1

        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(CHROME_DEBUG_URL)
            context = browser.contexts[0]

            for i, arquivo in enumerate(arquivos):
                if i % paginas_lote == 0:
                    page = context.new_page()
                    page.goto(link_gpt)
                    time.sleep(5)

                nome = os.path.basename(arquivo)
                self.log_output.append(f"📄 Processando: {nome}")

                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    with open(arquivo, "rb") as f:
                        tmp.write(f.read())
                    tmp_path = tmp.name

                page.set_input_files("input[type='file']", tmp_path)
                time.sleep(3)
                page.keyboard.type(prompt)
                page.keyboard.press("Enter")

                html = self.esperar_resposta_gpt(page)
                if html.startswith("__ERRO"):
                    self.log_output.append(f"❌ Erro ao processar {nome}: {html}")
                else:
                    self.log_output.append(f"✅ Processado. Enviando ao Docs...")
                    try:
                        titulo = f"{nome} - parte {i // paginas_lote + 1}"
                        incremento = self.inserir_no_google_docs(service, doc_id, index, titulo, html)
                        index += incremento
                        self.log_output.append("📜 Inserido com sucesso.")
                    except Exception as e:
                        self.log_output.append(f"❌ Erro ao inserir no Docs: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
