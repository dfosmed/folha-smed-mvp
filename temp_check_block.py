import pandas as pd
df = pd.read_excel('outputs/SMED - C. Custo e Regime (5)_classificado.xlsx', 'Resumo Geral')
c_custo = "038 - OUTROS SMED"
regimes = ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"]
df_block = df[(df['C.Custo'] == c_custo) & (df['Regime'].isin(regimes))]
print(df_block)
