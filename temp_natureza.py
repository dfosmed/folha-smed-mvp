import pandas as pd
xls2 = pd.ExcelFile('outputs/SMED - C. Custo e Regime (5)_classificado.xlsx')
df2 = pd.read_excel(xls2, 'Resumo Geral')

# check naturezas that are Proventos
print("Proventos:")
print(df2[df2['Grupo'] == 'Provento']['Natureza'].unique())

print("\nDescontos:")
print(df2[df2['Grupo'] == 'Desconto']['Natureza'].unique())
