import io
import re
import unicodedata
import pandas as pd
import openpyxl

CONFIG_BLOCOS = {
    "OUTROS SMED - EFETIVOS - CREDOR 105": {
        "C.Custo": "038 - OUTROS SMED",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)"]
    },
    "OUTROS SMED - CONTRATADOS - CREDOR 543": {
        "C.Custo": "038 - OUTROS SMED",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"]
    },
    "OUTROS SMED - COMISSIONADOS - CREDOR 543": {
        "C.Custo": "038 - OUTROS SMED",
        "Regimes": ["003 - COMISSIONADO"]
    },
    "OUTROS SMED - AGENTE POLÍTICO - CREDOR 543": {
        "C.Custo": "038 - OUTROS SMED",
        "Regimes": ["011 - AGENTE POLITICO"]
    },
    "RESTANTE DA SMED 30% - COMISSIONADOS/CONTR. - CREDOR 543": {
        "C.Custo": "098 - RESTANTE SMED 30%",
        "Regimes": ["003 - COMISSIONADO"]
    },
    "RESTANTE DA SMED 30% EFETIVOS - CREDOR 105": {
        "C.Custo": "098 - RESTANTE SMED 30%",
        "Regimes": ["015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)"]
    },
    "RESTANTE DA SMED 30% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "098 - RESTANTE SMED 30%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"]
    },
    "ED.FUNDAMENTAL 70% - COMISSIONADO- CREDOR 543": {
        "C.Custo": "093 - ENSINO FUNDAMENTAL 70%",
        "Regimes": ["003 - COMISSIONADO"]
    },
    "ED.FUNDAMENTAL 70%  - EFETIVOS - CREDOR 105": {
        "C.Custo": "093 - ENSINO FUNDAMENTAL 70%",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "030 - PROF EDUCACAO EST./COMISSIONADO", "031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)", "032 - EFETIVO/COMISSIONADO (CARREIRA)"]
    },
    "ED.FUNDAMENTAL 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "093 - ENSINO FUNDAMENTAL 70%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"]
    },
    "ENSINO FUNDAMENTAL  30% EFETIVOS - CREDOR 105": {
        "C.Custo": "092 - ENSINO FUNDAMENTAL 30%",
        "Regimes": ["002 - EFETIVO"]
    },
    "ENSINO FUNDAMENTAL 30% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "092 - ENSINO FUNDAMENTAL 30%",
        "Regimes": ["009 - CONTRATO"]
    },
    "ED.INFANTIL  PRE ESCOLA 70% - EFETIVOS - CREDOR 105": {
        "C.Custo": "083 - E.I. PRÉ-ESCOLA 70%",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "030 - PROF EDUCACAO EST./COMISSIONADO", "031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)"]
    },
    "ED.INFANTIL  PRE ESCOLA 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "083 - E.I. PRÉ-ESCOLA 70%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"]
    },
    "ED.INFANTIL  PRE ESCOLA 30% EFETIVOS - CREDOR 105": {
        "C.Custo": "082 - E.I. PRÉ-ESCOLA 30%",
        "Regimes": ["002 - EFETIVO", "019 - PROFISSIONAL EDUCACAO ( EST.)"]
    },
    "ED.INFANTIL CRECHE 70% - EFETIVOS - CREDOR 105": {
        "C.Custo": "077 - E.I. CRECHE 70%",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "030 - PROF EDUCACAO EST./COMISSIONADO", "031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)", "032 - EFETIVO/COMISSIONADO (CARREIRA)"]
    },
    "ED.INFANTIL  CRECHE 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "077 - E.I. CRECHE 70%",
        "Regimes": ["002 - EFETIVO", "020 - PROFISSIONAL EDUCACAO (CONT.)"]
    },
    "ED.INFANTIL  CRECHE 30% EFETIVOS - CREDOR 105": {
        "C.Custo": "076 - E.I. CRECHE 30%",
        "Regimes": ["002 - EFETIVO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "030 - PROF EDUCACAO EST./COMISSIONADO"]
    },
    "EDUCAÇÃO ESPECIAL 70% EFETIVOS - CREDOR 105": {
        "C.Custo": "072 - EDUCAÇÃO ESPECIAL 70%",
        "Regimes": ["002 - EFETIVO", "019 - PROFISSIONAL EDUCACAO ( EST.)"]
    },
    "EDUCAÇÃO ESPECIAL 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "072 - EDUCAÇÃO ESPECIAL 70%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"]
    },
    "EJA 70%  EFETIVOS - CREDOR 105": {
        "C.Custo": "088 - EJA 70%",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)"]
    },
    "EJA 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "088 - EJA 70%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"]
    }
}

NATUREZA_MAP = {
    'ABONO FAMILIA': 'ABONO FAMÍLIA',
    'SALÁRIO FAMILIA (INSS/CONTR.)': 'SALÁRIO FAMÍLIA',
    'AFAST.MATERNIDADE (SAL. MATERN.)': 'AFASTAMENTO MATERNIDADE',
    'AFAST. MATERNIDADE(INSS/CONTR.)': 'AFASTAMENTO MATERNIDADE',
    'AUXILIO DOENÇA (LICENÇA SAÚDE)': 'AUXÍLIO DOENÇA',
    'FÉRIAS 1/3 (ABONO CONSTITUCIONAL)': 'FÉRIAS - ABONO CONSTITUCIONAL',
    'FÉRIAS 1/3  (ABONO CONSTITUCIONAL)': 'FÉRIAS - ABONO CONSTITUCIONAL',
    'PARCELA PROP. (13ºSLR)': '13º SALÁRIO',
    'HORA EXTRA': 'HORA EXTRA',
    'INDENIZ. E RESTITUIÇÕES TRAB.': 'IDENIZAÇÕES E RESTITUIÇÕES',
    'AUXÍLIO NATALIDADE': 'AUXÍLIO NATALIDADE',
    'RESSARCIMENTO': 'RESSARCIMENTO',
    'DEDUÇÃO ART. 37, X': 'DEDUÇÃO ART. 37, X',
    'DIFERENÇA CARGA HORÁRIA': 'DIFERENÇA CARGA HORÁRIA',
    'DESCONTO DIAS/HORAS': 'DESCONTO DIAS/HORAS',
    'FALTAS/FALTAS HORAS': 'FALTAS/FALTAS HORAS',
}

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    # Tratar os espaços e unicode
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')
    text = text.upper()
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def normalize_key(text):
    text = normalize_text(text)
    text = re.sub(r'[^A-Z0-9]', '', text)
    return text

def parse_and_fill_contabilizacao(df_resumo: pd.DataFrame, path_modelo) -> bytes:
    if isinstance(path_modelo, str):
        wb = openpyxl.load_workbook(path_modelo)
    else:
        wb = openpyxl.load_workbook(io.BytesIO(path_modelo))
    sheet = wb.active
    
    # Precompute keys for block identification
    block_keys = {normalize_key(k): v for k, v in CONFIG_BLOCOS.items()}
    
    # Dicionario para acesso rápido às naturezas pelo texto
    natureza_map_norm = {normalize_key(k): v for k, v in NATUREZA_MAP.items()}
    
    current_block_config = None
    current_block_df = None
    start_data_row = None
    
    in_block = False
    folha_bruta_row = None
    mapped_naturezas_in_block = set()
    
    for row in range(1, sheet.max_row + 1):
        cell_val = sheet.cell(row=row, column=1).value
        if not cell_val:
            continue
            
        norm_cell = normalize_key(str(cell_val))
        
        # Detect block start
        if norm_cell in block_keys:
            current_block_config = block_keys[norm_cell]
            # Extrair df
            c_custo = current_block_config["C.Custo"]
            regimes = current_block_config["Regimes"]
            current_block_df = df_resumo[(df_resumo['C.Custo'] == c_custo) & (df_resumo['Regime'].isin(regimes))]
            
            in_block = True
            folha_bruta_row = None
            start_data_row = None
            mapped_naturezas_in_block = set()
            continue
            
        if in_block:
            # Look for Folha Bruta
            if "FOLHABRUTA" in norm_cell or "BRUTOSUBISDIO" in norm_cell or "BRUTOSUBSIDIO" in norm_cell:
                # Se ainda n achou, é a primeira e principal (ex: FOLHA BRUTA, FOLHA BRUTA (COMISSIONADOS), BRUTO SUBISIDIO)
                if not folha_bruta_row and "EMPENHO" not in norm_cell:
                    folha_bruta_row = row
                    # Calculo: soma dos proventos
                    val_folha = current_block_df[current_block_df['Grupo'] == 'Provento']['Valor'].sum()
                    if val_folha > 0:
                        sheet.cell(row=row, column=2).value = val_folha
                    else:
                        sheet.cell(row=row, column=2).value = None # Clear previous if any
                    continue
            
            # Look for block end / Folha Bruta para Empenho
            if "EMPENHO" in norm_cell and "FOLHA" in norm_cell and "BRUTA" in norm_cell:
                # É a ultima linha (FOLHA BRUTA PARA EMPENHO)
                # Somar todos os descontos n mapeados neste bloco
                unmapped_df = current_block_df[(current_block_df['Grupo'] == 'Desconto') & (~current_block_df['Natureza'].isin(mapped_naturezas_in_block))]
                sum_unmapped = unmapped_df['Valor'].sum()
                if sum_unmapped > 0:
                    sheet.cell(row=row, column=4).value = sum_unmapped
                else:
                    sheet.cell(row=row, column=4).value = None
                
                # Fechar bloco
                in_block = False
                continue
                
            # Else, see if it's a known mapped row
            if norm_cell in natureza_map_norm:
                expected_nat = natureza_map_norm[norm_cell]
                val_nat = current_block_df[current_block_df['Natureza'] == expected_nat]['Valor'].sum()
                
                if val_nat > 0:
                    sheet.cell(row=row, column=2).value = val_nat
                    mapped_naturezas_in_block.add(expected_nat)
                else:
                    sheet.cell(row=row, column=2).value = None
                    
    # Salvar em bytes
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
