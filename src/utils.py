import unicodedata

def normalizar_texto(texto):
    if not texto:
        return ""

    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ASCII", "ignore").decode("ASCII")
    return texto.upper()

def valor_brl_para_numero(valor):
    """
    Converte '120.148,48' em 120148.48
    """
    if valor is None:
        return None

    valor = valor.replace(".", "").replace(",", ".")

    try:
        return float(valor)
    except:
        return None

def agrupar_words_por_linha(words, tolerancia_y=3):
    """
    Agrupa palavras extraídas pelo pdfplumber em linhas visuais,
    mantendo a posição X de cada palavra.
    """
    if not words:
        return []

    words = sorted(words, key=lambda w: (w["top"], w["x0"]))

    linhas = []
    linha_atual = []
    y_atual = None

    for word in words:
        y = word["top"]

        if y_atual is None:
            linha_atual = [word]
            y_atual = y
        elif abs(y - y_atual) <= tolerancia_y:
            linha_atual.append(word)
        else:
            linhas.append(sorted(linha_atual, key=lambda w: w["x0"]))
            linha_atual = [word]
            y_atual = y

    if linha_atual:
        linhas.append(sorted(linha_atual, key=lambda w: w["x0"]))

    return linhas

def texto_linha(words_linha):
    return " ".join(w["text"] for w in sorted(words_linha, key=lambda w: w["x0"]))
