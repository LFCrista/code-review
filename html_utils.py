from bs4 import BeautifulSoup, NavigableString


def processar_html(html: str):
    """
    Converte HTML em (texto plano, estilos).
    estilos = lista de tuplas ( "bold"/"italic", ini, fim )
    """
    soup, txt, est = BeautifulSoup(html, "html.parser"), "", []

    def add(estilo, ini, fim):
        est.append((estilo, ini, fim))

    def walk(node, ativos):
        nonlocal txt
        if isinstance(node, NavigableString):
            ini, txt = len(txt), txt + str(node)
            fim = len(txt)
            for e in ativos:
                add(e, ini, fim)
        else:
            novos = ativos + (
                ["bold"] if node.name in ("strong", "b") else []
            ) + (
                ["italic"] if node.name in ("em", "i") else []
            )
            if node.name == "br":
                txt += "\n"
                return
            for child in node.children:
                walk(child, novos)
            if node.name == "p":
                txt += "\n"

    for el in soup.children:
        walk(el, [])
    txt += "\n\n"
    return txt, est


def inserir_no_docs(svc, doc_id, idx, title, html):
    """
    Insere título (Heading 1) + corpo no Google Docs e devolve
    quantos caracteres foram acrescentados ao índice.
    """
    texto, est = processar_html(html)
    req = [
        {"insertText": {"location": {"index": idx}, "text": title + "\n"}},
        {
            "updateParagraphStyle": {
                "range": {"startIndex": idx, "endIndex": idx + len(title) + 1},
                "paragraphStyle": {"namedStyleType": "HEADING_1"},
                "fields": "namedStyleType",
            }
        },
        {
            "insertText": {
                "location": {"index": idx + len(title) + 1},
                "text": texto,
            }
        },
    ]
    for estilo, s, f in est:
        style = {"bold": True} if estilo == "bold" else {"italic": True}
        req.append(
            {
                "updateTextStyle": {
                    "range": {
                        "startIndex": idx + len(title) + 1 + s,
                        "endIndex": idx + len(title) + 1 + f,
                    },
                    "textStyle": style,
                    "fields": estilo,
                }
            }
        )
    svc.documents().batchUpdate(documentId=doc_id, body={"requests": req}).execute()
    return len(title) + 1 + len(texto)
