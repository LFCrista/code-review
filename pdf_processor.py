import os
import time
from googleapiclient.discovery import build
from playwright.sync_api import sync_playwright

import config
import auth
import html_utils
import playwright_utils


def process_pdfs(
    pdfs: tuple[str],
    link_gpt: str,
    doc_id: str,
    prompt: str,
    max_per_tab: int,
    texto_log,
    janela,
):
    """
    Envia cada PDF ao ChatGPT, captura a resposta e grava no Google Docs.
    """
    if not pdfs:
        return

    texto_log.insert("end", "Início\n")
    janela.update()

    svc = build("docs", "v1", credentials=auth.autenticar_google_docs())
    idx = (
        svc.documents()
        .get(documentId=doc_id)
        .execute()["body"]["content"][-1]["endIndex"]
        - 1
    )

    with sync_playwright() as p:
        # --------- Tenta conectar ao Chrome em modo CDP -------------
        try:
            ctx = p.chromium.connect_over_cdp(config.CHROME_DEBUG_URL).contexts[0]
        except Exception as e:
            raise RuntimeError(
                "Não foi possível conectar ao Chrome em modo debug.\n"
                "Clique em 'Abrir Chrome Debug' (ou abra manualmente) e tente de novo."
            ) from e

        page = None

        for i, pdf in enumerate(pdfs, 1):
            # ---- Controle de abas a cada N arquivos ----------------
            if (i - 1) % max_per_tab == 0:
                if page and not page.is_closed():
                    try:
                        page.close()
                    except Exception:
                        pass
                page = ctx.new_page()
                page.goto(link_gpt)
                time.sleep(5)

            nome = os.path.basename(pdf)
            texto_log.insert("end", f"➡ {nome}\n")
            janela.update()

            cnt_before = page.locator("[data-message-author-role='assistant']").count()
            _ = playwright_utils.enviar_pdf(page, pdf, prompt)

            bolha = page.locator("[data-message-author-role='assistant']").nth(cnt_before)
            md = playwright_utils.esperar_markdown(bolha)
            playwright_utils.esperar_html_estavel(md)
            html = md.inner_html()

            if html.strip():
                idx += html_utils.inserir_no_docs(svc, doc_id, idx, nome, html)
                texto_log.insert("end", "✓ OK\n")
            else:
                texto_log.insert("end", "⚠ Sem resposta\n")

            playwright_utils.aguardar_pronto(page)
            janela.update()

        # Fecha última aba
        if page and not page.is_closed():
            try:
                page.close()
            except Exception:
                pass
