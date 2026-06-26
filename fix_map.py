with open('src/gerador_contabilizacao.py', 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("'SALRIO FAMILIA (INSS/CONTR.)': 'SALRIO FAMLIA'", "'SALÁRIO FAMILIA (INSS/CONTR.)': 'SALÁRIO FAMÍLIA'")
content = content.replace("'ABONO FAMILIA': 'ABONO FAMLIA'", "'ABONO FAMILIA': 'ABONO FAMÍLIA'")
content = content.replace("'AUXILIO DOENA (LICENA SADE)': 'AUXLIO DOENA'", "'AUXILIO DOENÇA (LICENÇA SAÚDE)': 'AUXÍLIO DOENÇA'")
content = content.replace("'FRIAS 1/3 (ABONO CONSTITUCIONAL)': 'FRIAS - ABONO CONSTITUCIONAL'", "'FÉRIAS 1/3 (ABONO CONSTITUCIONAL)': 'FÉRIAS - ABONO CONSTITUCIONAL'")
content = content.replace("'FRIAS 1/3  (ABONO CONSTITUCIONAL)': 'FRIAS - ABONO CONSTITUCIONAL'", "'FÉRIAS 1/3  (ABONO CONSTITUCIONAL)': 'FÉRIAS - ABONO CONSTITUCIONAL'")
content = content.replace("'PARCELA PROP. (13SLR)': '13 SALRIO'", "'PARCELA PROP. (13ºSLR)': '13º SALÁRIO'")
content = content.replace("'INDENIZ. E RESTITUIES TRAB.': 'IDENIZAES E RESTITUIES'", "'INDENIZ. E RESTITUIÇÕES TRAB.': 'IDENIZAÇÕES E RESTITUIÇÕES'")
content = content.replace("'AUXLIO NATALIDADE': 'AUXLIO NATALIDADE'", "'AUXÍLIO NATALIDADE': 'AUXÍLIO NATALIDADE'")
content = content.replace("'DEDUO ART. 37, X': 'DEDUO ART. 37, X'", "'DEDUÇÃO ART. 37, X': 'DEDUÇÃO ART. 37, X'")
content = content.replace("'DIFERENA CARGA HORRIA': 'DIFERENA CARGA HORRIA'", "'DIFERENÇA CARGA HORÁRIA': 'DIFERENÇA CARGA HORÁRIA'")

with open('src/gerador_contabilizacao.py', 'w', encoding='utf-8') as f:
    f.write(content)
