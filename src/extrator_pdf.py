import re
import os
import pdfplumber
import pandas as pd
from .utils import normalizar_texto, valor_brl_para_numero, agrupar_words_por_linha, texto_linha

def localizar_cabecalho_e_divisao(linhas_words, largura_pagina):
    """
    Localiza a linha do cabeçalho e a posição X onde começa a tabela da direita.
    """
    for idx, linha_words in enumerate(linhas_words):
        texto = texto_linha(linha_words)
        texto_norm = normalizar_texto(texto)

        if "CODIGO" in texto_norm and "VENCIMENTOS" in texto_norm and "DESCONTOS" in texto_norm:
            codigos = [
                w for w in linha_words
                if normalizar_texto(w["text"]) == "CODIGO"
            ]

            if len(codigos) >= 2:
                # O segundo "CÓDIGO" é o início da tabela de descontos
                inicio_direita = codigos[1]["x0"] - 2
            else:
                # Fallback: metade da página
                inicio_direita = largura_pagina / 2

            return idx, inicio_direita

    return None, largura_pagina / 2

def extrair_registro_do_trecho(trecho):
    """
    Extrai:
    Código, Descrição, Quantidade e Valor numérico.
    """
    trecho = trecho.strip()

    if not trecho:
        return None

    match_codigo = re.match(r'^(\d{5})\s+(.*)$', trecho)

    if not match_codigo:
        return None

    codigo = match_codigo.group(1)
    restante = match_codigo.group(2).strip()

    # Captura a quantidade e o valor no final da linha
    match_final = re.search(
        r'(\d+)\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s*$',
        restante
    )

    if not match_final:
        return None

    quantidade = int(match_final.group(1))
    valor_brl = match_final.group(2)
    valor = valor_brl_para_numero(valor_brl)

    descricao = restante[:match_final.start()].strip()

    return {
        "Codigo": codigo,
        "Descricao": descricao,
        "Quantidade": quantidade,
        "Valor": valor
    }

def extrair_dados_pdf(caminho_pdf):
    """
    Extrai dados da folha respeitando a posição visual:
    lado esquerdo = VENCIMENTOS
    lado direito = DESCONTOS
    """

    dados = []

    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina_num, page in enumerate(pdf.pages, start=1):
            texto = page.extract_text() or ""
            linhas_texto = texto.split("\n")

            # --------------------------------------------------
            # 1. Identificar Centro de Custo e Regime
            # --------------------------------------------------
            ccusto = None
            regime = None

            for linha in linhas_texto:
                linha_norm = normalizar_texto(linha)

                if "C. CUSTO:" in linha_norm or "C.CUSTO:" in linha_norm:
                    match = re.search(
                        r'C\.?\s*CUSTO:\s*(\d+)\s*-\s*(.*?)(?:\s+REGIME:|$)',
                        linha,
                        re.IGNORECASE
                    )

                    if match:
                        ccusto = f"{match.group(1)} - {match.group(2).strip()}"
                    else:
                        parts = re.split(r'C\.?\s*CUSTO:', linha, flags=re.IGNORECASE)
                        if len(parts) > 1:
                            ccusto = parts[1].strip()

                if "REGIME:" in linha_norm:
                    match = re.search(
                        r'REGIME:\s*(\d+)\s*-\s*(.*?)$',
                        linha,
                        re.IGNORECASE
                    )

                    if match:
                        regime = f"{match.group(1)} - {match.group(2).strip()}"
                    else:
                        parts = re.split(r'REGIME:', linha, flags=re.IGNORECASE)
                        if len(parts) > 1:
                            regime = parts[1].strip()

            if not ccusto or not regime:
                break

            # --------------------------------------------------
            # 2. Extrair palavras com coordenadas reais do PDF
            # --------------------------------------------------
            words = page.extract_words(
                x_tolerance=1,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False
            )

            linhas_words = agrupar_words_por_linha(words)

            indice_cabecalho, inicio_direita = localizar_cabecalho_e_divisao(
                linhas_words,
                page.width
            )

            if indice_cabecalho is None:
                continue

            vencimentos_pagina = []
            descontos_pagina = []

            # --------------------------------------------------
            # 3. Processar cada linha após o cabeçalho
            # --------------------------------------------------
            for linha_words in linhas_words[indice_cabecalho + 1:]:
                texto_completo = texto_linha(linha_words)
                texto_norm = normalizar_texto(texto_completo)

                if not texto_completo.strip():
                    continue

                # Ignorar totais, rodapés e separadores
                if (
                    "TOTAL DOS VENCIMENTOS" in texto_norm
                    or "TOTAL DOS DESCONTOS" in texto_norm
                    or "TOTAL LIQUIDO" in texto_norm
                    or "BASE DE CALCULO" in texto_norm
                    or "TOTAL GERAL" in texto_norm
                    or "RESUMO" in texto_norm
                ):
                    continue

                # Separar fisicamente pelo eixo X
                words_esquerda = [
                    w for w in linha_words
                    if w["x0"] < inicio_direita
                ]

                words_direita = [
                    w for w in linha_words
                    if w["x0"] >= inicio_direita
                ]

                texto_esquerda = texto_linha(words_esquerda).strip()
                texto_direita = texto_linha(words_direita).strip()

                # Lado esquerdo = vencimentos
                registro_esquerda = extrair_registro_do_trecho(texto_esquerda)

                if registro_esquerda:
                    vencimentos_pagina.append({
                        "Arquivo": os.path.basename(caminho_pdf),
                        "Página": pagina_num,
                        "C.Custo": ccusto,
                        "Regime": regime,
                        "Tipo": "VENCIMENTO",
                        **registro_esquerda
                    })

                # Lado direito = descontos
                registro_direita = extrair_registro_do_trecho(texto_direita)

                if registro_direita:
                    descontos_pagina.append({
                        "Arquivo": os.path.basename(caminho_pdf),
                        "Página": pagina_num,
                        "C.Custo": ccusto,
                        "Regime": regime,
                        "Tipo": "DESCONTO",
                        **registro_direita
                    })

            # Primeiro todos os vencimentos, depois todos os descontos
            dados.extend(vencimentos_pagina)
            dados.extend(descontos_pagina)

    return pd.DataFrame(dados)
