import os
import re
import json
import pandas as pd
from src.utils import normalizar_texto

def simplificar_texto(texto):
    if not texto:
        return ""
    txt = str(texto).upper()
    # Padroniza abreviações de número antes de remover acentos
    txt = txt.replace('Nº', 'N').replace('N.º', 'N').replace('N°', 'N')
    txt = re.sub(r'\b(NO|NUM|NUMERO)\b', 'N', txt)
    
    # Agora normaliza para remover acentos e caracteres especiais
    txt = normalizar_texto(txt)
    # Remove qualquer caractere que não seja letra ou número
    txt = re.sub(r'[^A-Z0-9]', '', txt)
    return txt


def classificar_rubricas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Classifica as rubricas de folha de pagamento baseadas em regras de negócio.
    Retorna um novo DataFrame com as colunas adicionais:
    Grupo, Natureza, Status e Observação.
    """
    df_classificado = df.copy()
    
    # Inicializa as novas colunas
    df_classificado['Grupo'] = ''
    df_classificado['Natureza'] = ''
    df_classificado['Status'] = ''
    df_classificado['Observação'] = ''
    
    # Carregar a base de dados de verbas para classificação de vencimentos
    dict_verbas_codigo = {}
    dict_verbas_desc_simpl = {}
    
    # 1. Carregar do banco de dados JSON local padrão
    caminho_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tabela_verbas_smed.json")
    if os.path.exists(caminho_json):
        try:
            with open(caminho_json, "r", encoding="utf-8") as f:
                db_json = json.load(f)
                for cod_str, info in db_json.items():
                    natureza_db = info.get("natureza", "")
                    desc_raw = info.get("descricao", "")
                    
                    if cod_str:
                        dict_verbas_codigo[cod_str] = natureza_db
                    if desc_raw:
                        desc_simpl = simplificar_texto(desc_raw)
                        if desc_simpl:
                            dict_verbas_desc_simpl[desc_simpl] = natureza_db
        except Exception as e:
            print(f"Erro ao carregar banco de dados JSON '{caminho_json}': {str(e)}")

    for idx, row in df_classificado.iterrows():
        tipo = str(row.get('Tipo', ''))
        codigo_raw = row.get('Codigo', '')
        
        # Normaliza o código do dataframe atual
        try:
            codigo = str(int(float(codigo_raw)))
        except ValueError:
            codigo = str(codigo_raw).strip()
            
        descricao = str(row.get('Descricao', ''))
        desc_norm = normalizar_texto(descricao)
        desc_simpl = simplificar_texto(descricao)
        
        grupo = ''
        natureza = ''
        status = ''
        obs = ''
        
        if tipo.upper() == 'VENCIMENTO':
            grupo = 'Provento'
        elif tipo.upper() == 'DESCONTO':
            grupo = 'Desconto'
        else:
            grupo = tipo.capitalize()
            
        # Busca estrita por Código na tabela oficial
        if codigo in dict_verbas_codigo and dict_verbas_codigo[codigo] != "":
            natureza = dict_verbas_codigo[codigo]
            status = 'Classificado'
        else:
            # Fallback para o tipo de grupo, ou deixar como Revisar
            natureza = 'Revisar'
            status = 'Pendente'
            obs = 'Rubrica não encontrada na tabela padrão'
            
        df_classificado.at[idx, 'Grupo'] = grupo
        df_classificado.at[idx, 'Natureza'] = natureza
        df_classificado.at[idx, 'Status'] = status
        df_classificado.at[idx, 'Observação'] = obs
        
    return df_classificado
