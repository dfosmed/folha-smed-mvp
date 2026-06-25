import pandas as pd
import unicodedata
import re

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')
    text = text.upper()
    text = re.sub(r'[^A-Z0-9]', '', text)
    return text

xls2 = pd.ExcelFile('outputs/SMED - C. Custo e Regime (5)_classificado.xlsx')
df2 = pd.read_excel(xls2, 'Resumo Geral')

# Block 1 rules
c_custo = '038 - OUTROS SMED'
regimes = ['002 - EFETIVO', '015 - EFETIVO/COMISSIONADO', '019 - PROFISSIONAL EDUCACAO ( EST.)', '031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)']

df_block = df2[(df2['C.Custo'] == c_custo) & (df2['Regime'].isin(regimes))]

folha_bruta = df_block[df_block['Grupo'] == 'Provento']['Valor'].sum()
print(f"Calculated FOLHA BRUTA: {folha_bruta}")

# Let's list the known DESCRICAO from Block 1 and their expected values (from temp_analyze.py)
# ('ABONO FAMILIA', 67.54)
# ('AFAST.MATERNIDADE (SAL. MATERN.)', None)
# ('AUXILIO DOENA (LICENA SADE)', 5780.49)
# ('FRIAS 1/3  (ABONO CONSTITUCIONAL)', 2307.66)
# ('PARCELA PROP. (13SLR)', 4611.05)
# ('HORA EXTRA', None)
# ('INDENIZ. E RESTITUIES TRAB.', 30253.04)
# ('RESSARCIMENTO', 241.46)
# ('FALTAS/FALTAS HORAS', 17784.23)

natureza_map = {
    'ABONO FAMILIA': 'ABONO FAMLIA',
    'AFAST.MATERNIDADE (SAL. MATERN.)': 'AFASTAMENTO MATERNIDADE',
    'AUXILIO DOENA (LICENA SADE)': 'AUXLIO DOENA',
    'FRIAS 1/3  (ABONO CONSTITUCIONAL)': 'FRIAS - ABONO CONSTITUCIONAL',
    'PARCELA PROP. (13SLR)': '13 SALRIO',
    'HORA EXTRA': 'HORA EXTRA',
    'INDENIZ. E RESTITUIES TRAB.': 'IDENIZAES E RESTITUIES',
    'RESSARCIMENTO': 'RESSARCIMENTO',
    'FALTAS/FALTAS HORAS': 'FALTAS/FALTAS HORAS',
}

mapped_naturezas = set()

for desc, nat in natureza_map.items():
    # normalize just in case, but here we can just use df_block
    # wait, the naturezas in temp_natureza output have strange characters because of utf-8 terminal issue
    # I will just match by normalized text
    pass

for nat in df_block['Natureza'].unique():
    val = df_block[df_block['Natureza'] == nat]['Valor'].sum()
    grupo = df_block[df_block['Natureza'] == nat]['Grupo'].iloc[0]
    print(f"Natureza: {nat} | Grupo: {grupo} | Valor: {val}")

