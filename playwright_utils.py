import time
from playwright.sync_api import sync_playwright
import config


# ------------------- Funções de estado -------------------------------
def _stop(page):
    return page.locator("button:has(svg[aria-label='Stop generating'])").count()


def _stream(page):
    return page.locator(".result-streaming, .animate-spin").count()


def _composer(page):
    chips = page.locator("[data-testid='file-upload-preview']").count() or page.locator(
        "div[role='listitem']"
    ).count()
    if chips:
        return True
    ta = page.locator("textarea")
    if ta.count():
        try:
            return bool(ta.input_value().strip())
        except Exception:
            return False
    return False


# ---------------------- Helpers principais ---------------------------
def aguardar_pronto(page, stab: float = 3.0, buf: float = 1.2):
    """
    Espera indefinidamente até que ChatGPT finalize a resposta
    (sem timeout).
    """
    last_cnt = page.locator("[data-message-author-role='assistant']").count()
    last_html = (
        page.locator("[data-message-author-role='assistant']")
        .nth(-1)
        .locator(".markdown")
        .inner_html()
        if last_cnt
        else ""
    )
    fase_buffer, t0 = False, time.time()

    while True:
        busy = _stop(page) or _stream(page) or _composer(page)
        cnt = page.locator("[data-message-author-role='assistant']").count()
        html = (
            page.locator("[data-message-author-role='assistant']")
            .nth(-1)
            .locator(".markdown")
            .inner_html()
            if cnt
            else ""
        )
        mudou = busy or cnt != last_cnt or html != last_html
        if mudou:
            fase_buffer, t0 = False, time.time()
            last_cnt, last_html = cnt, html
        else:
            now = time.time()
            if not fase_buffer and now - t0 >= buf:
                fase_buffer, t0 = True, now
            elif fase_buffer and now - t0 >= stab:
                return
        time.sleep(0.5)


def digitar_prompt(page, prompt: str):
    """Digita o prompt (textarea ou div[role='textbox'])."""
    for sel in ("textarea", "div[role='textbox']"):
        try:
            page.wait_for_selector(sel, timeout=2000)
            box = page.locator(sel).first
            if sel == "textarea":
                box.evaluate("n=>n.value=''")
                box.fill(prompt)
            else:
                box.click()
                box.evaluate("n=>n.innerText=''")
                box.type(prompt)
            page.keyboard.press("Enter")
            return
        except Exception:
            pass
    page.keyboard.type(prompt)
    page.keyboard.press("Enter")


def enviar_pdf(page, pdf: str, prompt: str):
    """
    Envia um PDF + prompt ao ChatGPT e retorna o HTML da resposta.
    Aguarda indefinidamente pelos elementos necessários.
    """
    aguardar_pronto(page)
    prev_cnt = page.locator("[data-message-author-role='assistant']").count()

    # Upload
    if page.locator("input[type='file']").count():
        page.set_input_files("input[type='file']", pdf)
    else:
        page.click("button:has(svg[aria-label='Upload a file'])", timeout=5000)
        page.wait_for_selector("input[type='file']", timeout=5000)
        page.set_input_files("input[type='file']", pdf)

    digitar_prompt(page, prompt)

    while page.locator("[data-message-author-role='assistant']").count() <= prev_cnt:
        time.sleep(0.5)

    aguardar_pronto(page)

    bolha = page.locator("[data-message-author-role='assistant']").nth(-1)
    page.wait_for_selector(
        "[data-message-author-role='assistant'] >> .markdown", timeout=0
    )
    return bolha.locator(".markdown").inner_html()


def esperar_html_estavel(locator, segundos: float = 2.0, dt: float = 0.4):
    """
    Aguarda indefinidamente até o innerHTML estabilizar por 'segundos'.
    """
    tam = len(locator.inner_html(timeout=0))
    t0 = time.time()
    while True:
        time.sleep(dt)
        novo = len(locator.inner_html(timeout=0))
        if novo != tam:
            tam, t0 = novo, time.time()
        elif time.time() - t0 >= segundos:
            return


def esperar_markdown(bolha, dt: float = 0.4):
    """
    Aguarda indefinidamente o aparecimento do elemento .markdown
    dentro da bolha de resposta.
    """
    while True:
        md = bolha.locator(".markdown, .prose")  # inclui fallback
        if md.count():
            return md
        time.sleep(dt)
