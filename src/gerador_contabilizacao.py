import io
import re
import os
import json
import unicodedata
import pandas as pd
import openpyxl
import logging
import datetime

logger = logging.getLogger(__name__)

MESES_MAP = {
    1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARÇO', 4: 'ABRIL',
    5: 'MAIO', 6: 'JUNHO', 7: 'JULHO', 8: 'AGOSTO',
    9: 'SETEMBRO', 10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO'
}

NOMES_MESES = [
    ('JANEIRO', 1), ('FEVEREIRO', 2), ('MARCO', 3), ('MARÇO', 3),
    ('ABRIL', 4), ('MAIO', 5), ('JUNHO', 6), ('JULHO', 7),
    ('AGOSTO', 8), ('SETEMBRO', 9), ('OUTUBRO', 10), ('NOVEMBRO', 11),
    ('DEZEMBRO', 12),
    ('JAN', 1), ('FEV', 2), ('MAR', 3), ('ABR', 4), ('MAI', 5),
    ('JUN', 6), ('JUL', 7), ('AGO', 8), ('SET', 9), ('OUT', 10),
    ('NOV', 11), ('DEZ', 12)
]

def extrair_mes_ano_nome_arquivo(nome_arquivo):
    """
    Extrai o mês e ano a partir do nome do arquivo (ex: 'FOLHA OUTUBRO 2026.pdf' -> 'OUTUBRO/2026').
    """
    if not nome_arquivo:
        return None
        
    base = re.sub(r'\.[a-zA-Z0-9]+$', '', os.path.basename(str(nome_arquivo)))
    norm = unicodedata.normalize('NFKD', base).encode('ASCII', 'ignore').decode('utf-8').upper()
    
    ano = None
    match_ano = re.search(r'\b(20[2-3]\d)\b', norm)
    if match_ano:
        ano = int(match_ano.group(1))
        
    mes_num = None
    for nome_m, num in NOMES_MESES:
        norm_m = unicodedata.normalize('NFKD', nome_m).encode('ASCII', 'ignore').decode('utf-8').upper()
        if re.search(r'(?<![A-Z0-9])' + norm_m + r'(?![A-Z0-9])', norm):
            mes_num = num
            break
            
    if not mes_num:
        match_num = re.search(r'(?<!\d)(0?[1-9]|1[0-2])[._/-](20[2-3]\d|\d{2})(?!\d)', norm)
        if match_num:
            mes_num = int(match_num.group(1))
            ano_cand = int(match_num.group(2))
            if ano_cand < 100:
                ano = 2000 + ano_cand
            else:
                ano = ano_cand
        else:
            match_inv = re.search(r'(?<!\d)(20[2-3]\d)[._/-](0?[1-9]|1[0-2])(?!\d)', norm)
            if match_inv:
                ano = int(match_inv.group(1))
                mes_num = int(match_inv.group(2))
            else:
                match_mes_sozinho = re.search(r'(?:FOLHA|SMED|PAGAMENTO)[^\d]*(0?[1-9]|1[0-2])(?!\d)', norm)
                if match_mes_sozinho:
                    mes_num = int(match_mes_sozinho.group(1))
                
    if mes_num:
        if not ano:
            match_2dig = re.search(r'[\s_.-](\d{2})(?!\d)', norm)
            if match_2dig:
                val2 = int(match_2dig.group(1))
                if 20 <= val2 <= 40:
                    ano = 2000 + val2
            if not ano:
                ano = datetime.datetime.now().year
        return f"{MESES_MAP[mes_num]}/{ano}"
    return None

CONFIG_BLOCOS = {
    "OUTROS SMED - EFETIVOS - CREDOR 105": {
        "C.Custo": "038 - OUTROS SMED",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "030 - PROF EDUCACAO EST./COMISSIONADO","031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)"],
        "Dotacao_CCusto": "OUTROS SMED",
        "Dotacao_Regime": "EFETIVO"
    },
    "OUTROS SMED - CONTRATADOS - CREDOR 543": {
        "C.Custo": "038 - OUTROS SMED",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"],
        "Dotacao_CCusto": "OUTROS SMED",
        "Dotacao_Regime": "CONTRATADO"
    },
    "OUTROS SMED - COMISSIONADOS - CREDOR 543": {
        "C.Custo": "038 - OUTROS SMED",
        "Regimes": ["003 - COMISSIONADO"],
        "Dotacao_CCusto": "OUTROS SMED",
        "Dotacao_Regime": "COMISSIONADOS"
    },
    "OUTROS SMED - AGENTE POLÍTICO - CREDOR 543": {
        "C.Custo": "038 - OUTROS SMED",
        "Regimes": ["011 - AGENTE POLITICO"],
        "Dotacao_CCusto": "OUTROS SMED",
        "Dotacao_Regime": "AGENTE POLÍTICO"
    },
    "RESTANTE DA SMED 30% - COMISSIONADOS/CONTR. - CREDOR 543": {
        "C.Custo": "098 - RESTANTE SMED 30%",
        "Regimes": ["003 - COMISSIONADO"],
        "Dotacao_CCusto": "RESTANTE DA SMED 30%",
        "Dotacao_Regime": "COMISSIONADOS"
    },
    "RESTANTE DA SMED 30% - EFETIVOS - CREDOR 105": {
        "C.Custo": "098 - RESTANTE SMED 30%",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)"],
        "Dotacao_CCusto": "RESTANTE DA SMED 30%",
        "Dotacao_Regime": "EFETIVO"
    },
    "RESTANTE DA SMED 30% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "098 - RESTANTE SMED 30%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"],
        "Dotacao_CCusto": "RESTANTE DA SMED 30%",
        "Dotacao_Regime": "CONTRATADO"
    },
    "ED.FUNDAMENTAL 70% - COMISSIONADO- CREDOR 543": {
        "C.Custo": "093 - ENSINO FUNDAMENTAL 70%",
        "Regimes": ["003 - COMISSIONADO"],
        "Dotacao_CCusto": "ED.FUNDAMENTAL 70%",
        "Dotacao_Regime": "COMISSIONADOS"
    },
    "ED.FUNDAMENTAL 70%  - EFETIVOS - CREDOR 105": {
        "C.Custo": "093 - ENSINO FUNDAMENTAL 70%",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "030 - PROF EDUCACAO EST./COMISSIONADO", "031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)", "032 - EFETIVO/COMISSIONADO (CARREIRA)"],
        "Dotacao_CCusto": "ED.FUNDAMENTAL 70%",
        "Dotacao_Regime": "EFETIVO"
    },
    "ED.FUNDAMENTAL 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "093 - ENSINO FUNDAMENTAL 70%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"],
        "Dotacao_CCusto": "ED.FUNDAMENTAL 70%",
        "Dotacao_Regime": "CONTRATADO"
    },
    "ENSINO FUNDAMENTAL  30% EFETIVOS - CREDOR 105": {
        "C.Custo": "092 - ENSINO FUNDAMENTAL 30%",
        "Regimes": ["002 - EFETIVO"],
        "Dotacao_CCusto": "ENSINO FUNDAMENTAL 30%",
        "Dotacao_Regime": "EFETIVO"
    },
    "ENSINO FUNDAMENTAL 30% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "092 - ENSINO FUNDAMENTAL 30%",
        "Regimes": ["009 - CONTRATO"],
        "Dotacao_CCusto": "ENSINO FUNDAMENTAL 30%",
        "Dotacao_Regime": "CONTRATADO"
    },
    "ED.INFANTIL  PRE ESCOLA 70% - EFETIVOS - CREDOR 105": {
        "C.Custo": "083 - E.I. PRÉ-ESCOLA 70%",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "030 - PROF EDUCACAO EST./COMISSIONADO", "031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)"],
        "Dotacao_CCusto": "ED.INFANTIL PRE ESCOLA 70%",
        "Dotacao_Regime": "EFETIVO"
    },
    "ED.INFANTIL  PRE ESCOLA 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "083 - E.I. PRÉ-ESCOLA 70%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"],
        "Dotacao_CCusto": "ED.INFANTIL PRE ESCOLA 70%",
        "Dotacao_Regime": "CONTRATADO"
    },
    "ED.INFANTIL  PRE ESCOLA 30% EFETIVOS - CREDOR 105": {
        "C.Custo": "082 - E.I. PRÉ-ESCOLA 30%",
        "Regimes": ["002 - EFETIVO", "019 - PROFISSIONAL EDUCACAO ( EST.)"],
        "Dotacao_CCusto": "ED.INFANTIL PRE ESCOLA 30%",
        "Dotacao_Regime": "EFETIVO"
    },
    "ED.INFANTIL PRE ESCOLA 30% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "082 - E.I. PRÉ-ESCOLA 30%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"],
        "Dotacao_CCusto": "ED.INFANTIL PRE ESCOLA 30%",
        "Dotacao_Regime": "CONTRATADO"
    },
    "ED.INFANTIL CRECHE 70% - EFETIVOS - CREDOR 105": {
        "C.Custo": "077 - E.I. CRECHE 70%",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "030 - PROF EDUCACAO EST./COMISSIONADO", "031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)", "032 - EFETIVO/COMISSIONADO (CARREIRA)"],
        "Dotacao_CCusto": "ED.INFANTIL CRECHE 70%",
        "Dotacao_Regime": "EFETIVO"
    },
    "ED.INFANTIL  CRECHE 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "077 - E.I. CRECHE 70%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"],
        "Dotacao_CCusto": "ED.INFANTIL CRECHE 70%",
        "Dotacao_Regime": "CONTRATADO"
    },
    "ED.INFANTIL  CRECHE 30% EFETIVOS - CREDOR 105": {
        "C.Custo": "076 - E.I. CRECHE 30%",
        "Regimes": ["002 - EFETIVO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "030 - PROF EDUCACAO EST./COMISSIONADO"],
        "Dotacao_CCusto": "ED.INFANTIL CRECHE 30%",
        "Dotacao_Regime": "EFETIVO"
    },
    "ED.INFANTIL  CRECHE 30% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "076 - E.I. CRECHE 30%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"],
        "Dotacao_CCusto": "ED.INFANTIL CRECHE 30%",
        "Dotacao_Regime": "CONTRATADO"
    },
    "EDUCAÇÃO ESPECIAL 70% EFETIVOS - CREDOR 105": {
        "C.Custo": "072 - EDUCAÇÃO ESPECIAL 70%",
        "Regimes": ["002 - EFETIVO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "026 - EFETIVO/CEDIDO", "030 - PROF EDUCACAO EST./COMISSIONADO"],
        "Dotacao_CCusto": "EDUCAÇÃO ESPECIAL 70%",
        "Dotacao_Regime": "EFETIVO"
    },
    "EDUCAÇÃO ESPECIAL 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "072 - EDUCAÇÃO ESPECIAL 70%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"],
        "Dotacao_CCusto": "EDUCAÇÃO ESPECIAL 70%",
        "Dotacao_Regime": "CONTRATADO"
    },
    "EJA 70%  EFETIVOS - CREDOR 105": {
        "C.Custo": "088 - EJA 70%",
        "Regimes": ["002 - EFETIVO", "015 - EFETIVO/COMISSIONADO", "019 - PROFISSIONAL EDUCACAO ( EST.)", "031 - PROF EDUCACAO EST./COMISSIONADO (CARREIRA)"],
        "Dotacao_CCusto": "EJA 70%",
        "Dotacao_Regime": "EFETIVO"
    },
    "EJA 70% - CONTRATADOS - CREDOR 543": {
        "C.Custo": "088 - EJA 70%",
        "Regimes": ["009 - CONTRATO", "020 - PROFISSIONAL EDUCACAO (CONT.)"],
        "Dotacao_CCusto": "EJA 70%",
        "Dotacao_Regime": "CONTRATADO"
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
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')
    text = text.upper()
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def normalize_key(text):
    text = normalize_text(text)
    text = re.sub(r'[^A-Z0-9]', '', text)
    return text

def get_dotacao_value(dotacoes_json, ccusto, regime, natureza_planilha, tipo_dotacao):
    if dotacoes_json is None or not tipo_dotacao:
        return None
        
    natureza_norm = normalize_key(natureza_planilha)
    tipo_norm = normalize_key(tipo_dotacao)
    
    # Encontrar no json iterando
    for row_data in dotacoes_json:
        r_ccusto = str(row_data.get('C. CUSTO', '')).strip().upper()
        r_regime = str(row_data.get('REGIME', '')).strip().upper()
        r_desc = str(row_data.get('DESCRIÇÃO', row_data.get('DESCRIÇO', ''))).strip() 
        
        if r_ccusto == ccusto.strip().upper() and r_regime == regime.strip().upper():
            if normalize_key(r_desc) == natureza_norm:
                # Encontrou a linha, agora acha a coluna
                for col, val in row_data.items():
                    if normalize_key(col) == tipo_norm:
                        if val is not None and not (isinstance(val, float) and pd.isna(val)):
                            if str(val).strip().lower() != "nan":
                                return str(val).strip()
                break
    return None

def parse_and_fill_contabilizacao(df_resumo: pd.DataFrame, path_modelo, config_dotacoes_dict=None, nome_arquivo: str = None, competencia: str = None) -> bytes:
    if config_dotacoes_dict is None:
        config_dotacoes_dict = {}
    if isinstance(path_modelo, str):
        wb = openpyxl.load_workbook(path_modelo)
    else:
        wb = openpyxl.load_workbook(io.BytesIO(path_modelo))
    sheet = wb.active
    
    # Atualizar o mês/ano no cabeçalho superior da contabilização
    if not competencia and nome_arquivo:
        competencia = extrair_mes_ano_nome_arquivo(nome_arquivo)
        
    if competencia:
        for r in range(1, 10):
            for c in range(1, 6):
                cell = sheet.cell(row=r, column=c)
                val_str = str(cell.value or "")
                if "CONTABILIZA" in val_str.upper() and "FOLHA" in val_str.upper():
                    novo_val = re.sub(
                        r'(CONTABILIZA.*?FOLHA\s+DE\s+PAGAMENTO\s*-\s*)([^\n\r]+)',
                        rf'\g<1>{competencia}',
                        val_str,
                        flags=re.IGNORECASE
                    )
                    cell.value = novo_val
                    break
    
    # Carregar detalhamento de dotações via JSON
    dotacoes_json = None
    try:
        json_path = os.path.join(os.path.dirname(__file__), 'dotacoes.json')
        if os.path.exists(json_path):
            with open(json_path, 'r', encoding='utf-8') as f:
                dotacoes_json = json.load(f)
    except Exception as e:
        logger.error(f"Não foi possivel carregar Detalhamento de Dotações JSON: {e}")
    
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
            # Atualiza dinamicamente o campo "PAGO COM XXX" de acordo com o recurso selecionado para o bloco
            if "PAGO COM" in str(cell_val).upper():
                d_ccusto = current_block_config.get("Dotacao_CCusto")
                d_regime = current_block_config.get("Dotacao_Regime")
                if d_ccusto and d_regime:
                    key = f"{d_ccusto.upper()}_{d_regime.upper()}"
                    tipo_selecionado = config_dotacoes_dict.get(key, None)
                    if tipo_selecionado:
                        recurso = re.sub(r'^(PAGO\s+COM|DOTA[CÇ][AÃ]O)\s*', '', str(tipo_selecionado).strip(), flags=re.IGNORECASE).strip()
                        if type(sheet.cell(row=row, column=1)).__name__ != 'MergedCell':
                            sheet.cell(row=row, column=1).value = f"PAGO COM {recurso}"
                continue

            dotacao_atual = sheet.cell(row=row, column=3).value
            manter_dotacao = False
            
            # Força o NÃO EMPENHAR e OP EXTRA para rubricas específicas que sempre o utilizam
            rubricas_nao_empenhar = [
                "PARCELAANT13", "RESSARCIMENTO", "DEDUCAOART37X", 
                "DIFERENCACARGAHORARIA", "DESCONTODIASHORAS", "FALTASFALTASHORAS"
            ]
            rubricas_op_extra = [
                "SALARIOFAMILIAINSSCONTR",
                "AFASTMATERNIDADEINSSCONTR",
                "AFASTMATERNIDADESALMATERN"
            ]
            
            if norm_cell in rubricas_nao_empenhar:
                # Trata erro de merged cell
                if type(sheet.cell(row=row, column=3)).__name__ != 'MergedCell':
                    sheet.cell(row=row, column=3).value = "NÃO EMPENHAR"
                    import copy
                    if sheet.cell(row=row, column=3).font:
                        new_font = copy.copy(sheet.cell(row=row, column=3).font)
                        new_font.bold = True
                        sheet.cell(row=row, column=3).font = new_font
                manter_dotacao = True
                
            if norm_cell in rubricas_op_extra:
                if type(sheet.cell(row=row, column=3)).__name__ != 'MergedCell':
                    sheet.cell(row=row, column=3).value = "OP EXTRA"
                    import copy
                    if sheet.cell(row=row, column=3).font:
                        new_font = copy.copy(sheet.cell(row=row, column=3).font)
                        new_font.bold = True
                        sheet.cell(row=row, column=3).font = new_font
                manter_dotacao = True
                
            if dotacao_atual and ("NǟO EMPENHAR" in str(dotacao_atual).upper() or "NAO EMPENHAR" in str(dotacao_atual).upper() or "NÃO EMPENHAR" in str(dotacao_atual).upper() or "OP EXTRA" in str(dotacao_atual).upper()):
                manter_dotacao = True
                
            # Look for Folha Bruta
            if "FOLHABRUTA" in norm_cell or "BRUTOSUBISDIO" in norm_cell or "BRUTOSUBSIDIO" in norm_cell or ("BRUTO" in norm_cell and "SECRETARIO" in norm_cell):
                if not folha_bruta_row and "EMPENHO" not in norm_cell:
                    folha_bruta_row = row
                    
                    if current_block_config.get("Dotacao_Regime") == "AGENTE POLÍTICO":
                        val_folha = current_block_df[current_block_df['Natureza'] == 'VENCIMENTOS']['Valor'].sum()
                    else:
                        val_folha = current_block_df[current_block_df['Grupo'] == 'Provento']['Valor'].sum()
                        
                    if val_folha > 0:
                        sheet.cell(row=row, column=2).value = val_folha
                    else:
                        sheet.cell(row=row, column=2).value = None 
                    continue
            
            # Look for block end / Folha Bruta para Empenho
            if "EMPENHO" in norm_cell and (("FOLHA" in norm_cell and "BRUTA" in norm_cell) or "SUBSIDIO" in norm_cell):
                if current_block_config.get("Dotacao_Regime") == "AGENTE POLÍTICO":
                    val_folha_empenho = current_block_df[current_block_df['Natureza'] == 'VENCIMENTOS']['Valor'].sum()
                else:
                    val_folha_empenho = current_block_df[current_block_df['Grupo'] == 'Provento']['Valor'].sum()
                    
                unmapped_df = current_block_df[(current_block_df['Grupo'] == 'Desconto') & (~current_block_df['Natureza'].isin(mapped_naturezas_in_block))]
                sum_unmapped = unmapped_df['Valor'].sum()
                
                if val_folha_empenho > 0 and sum_unmapped > 0:
                    sheet.cell(row=row, column=4).value = sum_unmapped
                else:
                    sheet.cell(row=row, column=4).value = None
                    
                nova_dotacao_empenho = None
                if dotacoes_json is not None:
                    d_ccusto = current_block_config.get("Dotacao_CCusto")
                    d_regime = current_block_config.get("Dotacao_Regime")
                    if d_ccusto and d_regime:
                        key = f"{d_ccusto.upper()}_{d_regime.upper()}"
                        tipo_selecionado = config_dotacoes_dict.get(key, None)
                        # Busca do JSON a rubrica "FOLHA BRUTA PARA EMPENHO"
                        nova_dotacao_empenho = get_dotacao_value(dotacoes_json, d_ccusto, d_regime, "FOLHA BRUTA PARA EMPENHO", tipo_selecionado)

                is_special_dynamic_empenho = nova_dotacao_empenho and nova_dotacao_empenho.upper() in ["OP EXTRA", "NÃO EMPENHAR", "NAO EMPENHAR"]
                
                def set_dotacao_empenho(val, bold=False):
                    if type(sheet.cell(row=row, column=3)).__name__ != 'MergedCell':
                        sheet.cell(row=row, column=3).value = val
                        import copy
                        if sheet.cell(row=row, column=3).font:
                            new_font = copy.copy(sheet.cell(row=row, column=3).font)
                            new_font.bold = bold
                            sheet.cell(row=row, column=3).font = new_font
                
                if val_folha_empenho > 0 and sum_unmapped > 0:
                    if is_special_dynamic_empenho:
                        set_dotacao_empenho(nova_dotacao_empenho, bold=True)
                    elif nova_dotacao_empenho:
                        set_dotacao_empenho(nova_dotacao_empenho, bold=False)
                    else:
                        set_dotacao_empenho(None, bold=False)
                else:
                    if is_special_dynamic_empenho:
                        set_dotacao_empenho(nova_dotacao_empenho, bold=True)
                    else:
                        set_dotacao_empenho(None, bold=False)
                
                in_block = False
                continue
                
            # Else, see if it's a known mapped row
            if norm_cell in natureza_map_norm:
                expected_nat = natureza_map_norm[norm_cell]
                val_nat = current_block_df[current_block_df['Natureza'] == expected_nat]['Valor'].sum()
                
                nova_dotacao = None
                if dotacoes_json is not None:
                    d_ccusto = current_block_config.get("Dotacao_CCusto")
                    d_regime = current_block_config.get("Dotacao_Regime")
                    if d_ccusto and d_regime:
                        key = f"{d_ccusto.upper()}_{d_regime.upper()}"
                        tipo_selecionado = config_dotacoes_dict.get(key, None)
                        nova_dotacao = get_dotacao_value(dotacoes_json, d_ccusto, d_regime, cell_val, tipo_selecionado)
                        
                is_special_dynamic = nova_dotacao and nova_dotacao.upper() in ["OP EXTRA"]
                is_special_rubrica = norm_cell in rubricas_nao_empenhar
                is_special_rubrica_op = norm_cell in rubricas_op_extra
                
                def set_dotacao(val, bold=False):
                    if type(sheet.cell(row=row, column=3)).__name__ != 'MergedCell':
                        sheet.cell(row=row, column=3).value = val
                        import copy
                        if sheet.cell(row=row, column=3).font:
                            new_font = copy.copy(sheet.cell(row=row, column=3).font)
                            new_font.bold = bold
                            sheet.cell(row=row, column=3).font = new_font
                            
                if val_nat > 0:
                    sheet.cell(row=row, column=2).value = val_nat
                    mapped_naturezas_in_block.add(expected_nat)
                    
                    if is_special_rubrica:
                        pass # Já forçado no início do bloco
                    elif is_special_rubrica_op:
                        pass # Já forçado no início do bloco como OP EXTRA
                    elif is_special_dynamic:
                        set_dotacao(nova_dotacao, bold=True)
                    elif nova_dotacao:
                        set_dotacao(nova_dotacao, bold=False)
                    else:
                        set_dotacao(None, bold=False)
                else:
                    sheet.cell(row=row, column=2).value = None
                    if is_special_rubrica:
                        pass # Já forçado no início do bloco
                    elif is_special_rubrica_op:
                        set_dotacao(None, bold=False) # Se for zero, tira o OP EXTRA
                    elif is_special_dynamic:
                        set_dotacao(nova_dotacao, bold=True)
                    else:
                        set_dotacao(None, bold=False)
                            
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
