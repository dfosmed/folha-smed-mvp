import json
import pandas as pd
import numpy as np
import os
import io

JSON_FILE = 'src/consignados_descontos.json'

def gerar_consignados_excel(df):
    """
    Recebe um DataFrame (correspondente à aba 'Resumo Geral') e retorna os bytes do Excel gerado.
    """
    df['Natureza_Upper'] = df['Natureza'].astype(str).str.strip().str.upper()

    listas_por_grupo = {
        "EFETIVOS": [
            (10, "SINDICATO"), (11, "PENSÃO"), (39, "CAIXA"), (84, "BIG CARD"), (86, "BANCO BRASIL"), 
            (92, "COUNTRY"), (124, "MINAS CLUBE"), (125, "PARANA"), (1276, "VALE TRANSPORTE"), 
            (1279, "SINSEM CLUBE"), (1280, "SICOOB AC CR"), (1284, "BC COOPERATIVO"), (1292, "BRADESCO"), 
            (1293, "SANTANDER"), (1294, "IRRF"), (1298, "UP BRASIL"), (1299, "NOTRE DAME"), 
            (1454, "IPREM LEI 316/2023"), (1455, "BC DAYCOVAL"), (1457, "BC PAN"), (1462, "PAM"), 
            (1463, "B. NIO"), (1466, "DESC JUDICIAL"), (1469, "BC MASTER"), (1472, "SICREDI"), 
            (1473, "BR CARD"), (1476, "CASH CARD"), (1478, "IPREM")
        ],
        "CONTRATADOS": [
            (10, "SINDICATO"), (11, "PENSÃO"), (92, "COUNTRY"), (1276, "VALE TRANSPORTE"),
            (1279, "SINSEM CLUBE"), (1294, "IRRF"), (1479, "INSS PF")
        ],
        "COMISSIONADOS": [
            (10, "SINDICATO"), (11, "PENSÃO"), (92, "COUNTRY"), (1276, "VALE TRANSPORTE"),
            (1279, "SINSEM CLUBE"), (1294, "IRRF"), (1479, "INSS PF")
        ]
    }

    ccustos_unicos = df['C.Custo'].dropna().unique()
    def get_ccusto_by_prefix(prefix):
        for c in ccustos_unicos:
            if str(c).startswith(prefix):
                return c
        return None

    regras = {
        "OUTROS SMED": {
            "prefix": "038 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO', '015 - EFETIVO/COMISSIONADO', '019 - PROFISSIONAL EDUCACAO ( EST.)', '031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)'],
                "CONTRATADOS": ['009 - CONTRATO', '020 - PROFISSIONAL EDUCACAO (CONT.)'],
                "COMISSIONADOS": ['003 - COMISSIONADO'],
                "AGENTE POLÍTICO": ['011 - AGENTE POLITICO']
            }
        },
        "RESTANTE DA SMED 30%": {
            "prefix": "098 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO','015 - EFETIVO/COMISSIONADO', '019 - PROFISSIONAL EDUCACAO ( EST.)'],
                "CONTRATADOS": ['009 - CONTRATO', '020 - PROFISSIONAL EDUCACAO (CONT.)'],
                "COMISSIONADOS/CONTR.": ['003 - COMISSIONADO']
            }
        },
        "ED.FUNDAMENTAL 70%": {
            "prefix": "093 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO', '015 - EFETIVO/COMISSIONADO', '019 - PROFISSIONAL EDUCACAO ( EST.)', '030 - PROF EDUCACAO EST./COMISSIONADO', '031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)', '032 - EFETIVO/COMISSIONADO (CARREIRA)'],
                "CONTRATADOS": ['009 - CONTRATO', '020 - PROFISSIONAL EDUCACAO (CONT.)'],
                "COMISSIONADO": ['003 - COMISSIONADO']
            }
        },
        "ENSINO FUNDAMENTAL 30%": {
            "prefix": "092 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO'],
                "CONTRATADOS": ['009 - CONTRATO']
            }
        },
        "ED.INFANTIL PRE 70%": {
            "prefix": "083 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO', '015 - EFETIVO/COMISSIONADO', '019 - PROFISSIONAL EDUCACAO ( EST.)', '030 - PROF EDUCACAO EST./COMISSIONADO', '031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)'],
                "CONTRATADOS": ['009 - CONTRATO', '020 - PROFISSIONAL EDUCACAO (CONT.)']
            }
        },
        "ED.INFANTIL PRE 30%": {
            "prefix": "082 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO', '019 - PROFISSIONAL EDUCACAO ( EST.)']
            }
        },
        "ED.INFANTIL CRECHE 70%": {
            "prefix": "077 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO', '015 - EFETIVO/COMISSIONADO', '019 - PROFISSIONAL EDUCACAO ( EST.)', '030 - PROF EDUCACAO EST./COMISSIONADO', '031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)', '032 - EFETIVO/COMISSIONADO (CARREIRA)'],
                "CONTRATADOS": ['009 - CONTRATO', '020 - PROFISSIONAL EDUCACAO (CONT.)']
            }
        },
        "ED.INFANTIL CRECHE 30%": {
            "prefix": "076 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO', '019 - PROFISSIONAL EDUCACAO ( EST.)', '030 - PROF EDUCACAO EST./COMISSIONADO']
            }
        },
        "EDUCAÇÃO ESPECIAL 70%": {
            "prefix": "072 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO', '019 - PROFISSIONAL EDUCACAO ( EST.)'],
                "CONTRATADOS": ['009 - CONTRATO', '020 - PROFISSIONAL EDUCACAO (CONT.)']
            }
        },
        "EJA 70%": {
            "prefix": "088 - ",
            "grupos": {
                "EFETIVOS": ['002 - EFETIVO', '015 - EFETIVO/COMISSIONADO', '019 - PROFISSIONAL EDUCACAO ( EST.)', '031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)'],
                "CONTRATADOS": ['009 - CONTRATO', '020 - PROFISSIONAL EDUCACAO (CONT.)']
            }
        }
    }

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book

    format_header = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    format_title = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    format_total = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'right',
        'valign': 'vcenter'
    })

    format_total_value = workbook.add_format({
        'bold': True,
        'border': 1,
        'num_format': 'R$ #,##0.00'
    })

    format_money = workbook.add_format({
        'border': 1,
        'num_format': 'R$ #,##0.00'
    })

    format_cell = workbook.add_format({
        'border': 1
    })

    for sheet_name, regras_aba in regras.items():
        prefix = regras_aba['prefix']
        grupos = regras_aba['grupos']
        
        c_custo_real = get_ccusto_by_prefix(prefix)
        if not c_custo_real:
            continue

        df_ccusto = df[df['C.Custo'] == c_custo_real]
        worksheet = workbook.add_worksheet(sheet_name)
        
        worksheet.set_column('A:A', 8)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 18)
        
        current_row = 0
        
        for nome_grupo, regimes in grupos.items():
            df_grupo = df_ccusto[df_ccusto['Regime'].isin(regimes)]
            
            dados_saida = []
            total_grupo = 0
            
            if "EFETIVO" in nome_grupo:
                lista_itens = listas_por_grupo["EFETIVOS"]
            elif "CONTRATADO" in nome_grupo:
                lista_itens = listas_por_grupo["CONTRATADOS"]
            elif "COMISSIONADO" in nome_grupo or "AGENTE" in nome_grupo:
                lista_itens = listas_por_grupo["COMISSIONADOS"]
            else:
                lista_itens = listas_por_grupo["EFETIVOS"]
            
            for codigo, descricao in lista_itens:
                natureza_upper = descricao.strip().upper()
                
                valor_soma = df_grupo.loc[df_grupo['Natureza_Upper'] == natureza_upper, 'Valor'].sum()
                
                dados_saida.append((codigo, descricao, valor_soma))
                total_grupo += valor_soma

            worksheet.merge_range(current_row, 0, current_row, 2, nome_grupo, format_title)
            current_row += 1
            
            worksheet.write(current_row, 0, 'CÓD', format_header)
            worksheet.write(current_row, 1, 'DESCRIÇÃO', format_header)
            worksheet.write(current_row, 2, 'VALOR', format_header)
            current_row += 1
            
            for row_data in dados_saida:
                worksheet.write(current_row, 0, row_data[0], format_cell) 
                worksheet.write(current_row, 1, row_data[1], format_cell) 
                
                valor_val = row_data[2]
                if pd.isna(valor_val):
                    valor_val = 0
                
                worksheet.write(current_row, 2, valor_val, format_money)
                current_row += 1
            
            worksheet.merge_range(current_row, 0, current_row, 1, 'TOTAL: R$', format_total)
            worksheet.write(current_row, 2, total_grupo, format_total_value)
            
            current_row += 3

    writer.close()
    return output.getvalue()

def main():
    EXCEL_IN = 'outputs/SMED - C. Custo e Regime (5)_classificado.xlsx'
    EXCEL_OUT = 'outputs/Consignado por centro de custo e regime.xlsx'
    print(f"Lendo arquivo base: {EXCEL_IN}")
    df = pd.read_excel(EXCEL_IN, sheet_name='Resumo Geral')
    
    bytes_excel = gerar_consignados_excel(df)
    
    with open(EXCEL_OUT, 'wb') as f:
        f.write(bytes_excel)
    print("Planilha gerada com sucesso!")

if __name__ == '__main__':
    main()
