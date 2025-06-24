import os
import re
from tkinter import filedialog
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from config import SCOPES


def autenticar_google_docs():
    """
    Faz OAuth2, devolvendo um objeto Credentials.
    Guarda/recupera o token em token.json.
    """
    if os.path.exists("token.json"):
        try:
            return Credentials.from_authorized_user_file("token.json", SCOPES)
        except Exception:
            os.remove("token.json")

    cred = filedialog.askopenfilename(
        title="Selecione o credentials.json", filetypes=[("JSON", "*.json")]
    )
    if not cred:
        return None

    flow = InstalledAppFlow.from_client_secrets_file(cred, SCOPES)
    creds = flow.run_local_server(port=0)
    with open("token.json", "w") as fp:
        fp.write(creds.to_json())
    return creds


def extrair_document_id(url: str) -> str | None:
    """
    Extrai o ID de um Google Document da URL.
    """
    m = re.search(r"/d/([A-Za-z0-9-_]+)", url)
    return m.group(1) if m else None
