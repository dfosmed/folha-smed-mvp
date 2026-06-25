import pandas as pd
from src.gerador_contabilizacao import parse_and_fill_contabilizacao

df = pd.read_excel('outputs/SMED - C. Custo e Regime (5)_classificado.xlsx', 'Resumo Geral')
res = parse_and_fill_contabilizacao(df, 'modelo_contabilização.xlsx')
with open('outputs/Contabilização Folha.xlsx', 'wb') as f:
    f.write(res)
print('Saved output')
