import json

with open('src/tabela_verbas_smed.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# The exact string used elsewhere
ferias_abono = "FÉRIAS - ABONO CONSTITUCIONAL"

if '3770' in data:
    data['3770']['natureza'] = ferias_abono

if '3780' in data:
    data['3780']['natureza'] = ferias_abono

with open('src/tabela_verbas_smed.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("JSON atualizado com sucesso.")
