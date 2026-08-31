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

    try:
        with open(JSON_FILE, 'r', encoding='utf-8') as f:
            efetivos_json = json.load(f)
            efetivos_list = [(item['codigo'], item['descricao']) for item in efetivos_json]
    except Exception:
        efetivos_list = [
            (142, "SINDICATO"), (470, "PENSÃO"), (136, "CAIXA"), (104, "BIG CARD"), (147, "BANCO BRASIL"), 
            (134, "COUNTRY"), (176, "MINAS CLUBE"), (232, "PARANA"), (559, "VALE TRANSPORTE"), 
            (266, "SINSEM CLUBE"), (277, "SICOOB AC CR"), (423, "BC COOPERATIVO"), (441, "BRADESCO"), 
            (309, "SANTANDER"), (531, "IRRF"), (448, "UP BRASIL"), (447, "NOTRE DAME"), 
            (3781, "IPREM LEI 316/2023"), (3759, "BC DAYCOVAL"), (3784, "BC PAN"), (101, "PAM"), 
            (315, "B. NIO"), (188, "DESC JUDICIAL"), (3882, "BC MASTER"), (3887, "SICREDI"), 
            (3886, "BR CARD"), (3909, "CASH CARD"), (554, "IPREM"), (3888, "SICRED 2")
        ]

    listas_por_grupo = {
        "EFETIVOS": efetivos_list,
        "CONTRATADOS": [
            (142, "SINDICATO"), (470, "PENSÃO"), (134, "COUNTRY"), (559, "VALE TRANSPORTE"),
            (266, "SINSEM CLUBE"), (531, "IRRF"), (528, "INSS PF")
        ],
        "COMISSIONADOS": [
            (142, "SINDICATO"), (470, "PENSÃO"), (134, "COUNTRY"), (559, "VALE TRANSPORTE"),
            (266, "SINSEM CLUBE"), (531, "IRRF"), (528, "INSS PF")
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
                "EFETIVOS": ['002 - EFETIVO', '015 - EFETIVO/COMISSIONADO', '019 - PROFISSIONAL EDUCACAO ( EST.)', '030 - PROF EDUCACAO EST./COMISSIONADO', '031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)'],
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
            # Garantir que consideramos apenas Descontos no relatório de consignados
            df_grupo_descontos = df_grupo[df_grupo['Grupo'].astype(str).str.strip().str.upper() == 'DESCONTO']
            
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
                
                valor_soma = df_grupo_descontos.loc[df_grupo_descontos['Natureza_Upper'] == natureza_upper, 'Valor'].sum()
                
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
