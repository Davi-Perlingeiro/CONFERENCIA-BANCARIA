import pandas as pd
import unicodedata
import re

def normalizar_nome(nome):
    if pd.isna(nome):
        return ''
    nome = str(nome).strip().upper()
    nome = unicodedata.normalize('NFD', nome)
    nome = ''.join(c for c in nome if unicodedata.category(c) != 'Mn')
    nome = re.sub(r'\s+', ' ', nome)
    return nome

# ============ BANCO ============
df_raw = pd.read_excel('ConsultaPagamentos 25.03.26.xls', header=None)
df_banco = df_raw.iloc[18:].copy()
df_banco.columns = ['nome', 'cpf_cnpj', 'tipo_pag', 'ref_empresa', 'data_pag', 'valor', 'status']
df_banco = df_banco.dropna(subset=['nome'])
df_banco = df_banco[df_banco['nome'] != 'Total:']
df_banco = df_banco[~df_banco['nome'].astype(str).str.contains('favorecido')]
df_banco['valor'] = pd.to_numeric(df_banco['valor'], errors='coerce')
df_banco['nome_norm'] = df_banco['nome'].apply(normalizar_nome)

# ============ FOLHAS ============
folhas = []
arquivos_folha = [
    ('0060_LIQUIDO_PG_RJ - RODANDO.xlsx', 'RJ 0060'),
    ('0041_LIQUIDO_PG_QUINZ - OP RJ - RODANDO.xlsx', 'QUINZ RJ 0041'),
    ('0041_LIQUIDO_PG_BH - RODANDO.xlsx', 'BH 0041'),
]

for arq, origem in arquivos_folha:
    df = pd.read_excel(arq, header=None)
    dados = df.iloc[8:].copy()
    dados = dados[dados[0].notna()]
    dados = dados[~dados[0].astype(str).str.contains('Total|Dpto|TOTAL|Resumo', na=False)]
    for i, row in dados.iterrows():
        nome = str(row[0]).strip()
        valor = row[15] if pd.notna(row[15]) else (row[16] if pd.notna(row[16]) else None)
        if nome and valor is not None:
            folhas.append({
                'nome': nome,
                'nome_norm': normalizar_nome(nome),
                'valor': float(valor),
                'origem': origem
            })

df_folhas = pd.DataFrame(folhas)

print('=' * 80)
print('RELATORIO DE COMPARACAO - BANCO vs FOLHAS DE PAGAMENTO')
print('=' * 80)
print(f'Data: 25/03/2026')
print(f'Registros no banco: {len(df_banco)}')
print(f'Registros nas folhas: {len(df_folhas)}')
print(f'  - RJ 0060: {len(df_folhas[df_folhas.origem == "RJ 0060"])}')
print(f'  - QUINZ RJ 0041: {len(df_folhas[df_folhas.origem == "QUINZ RJ 0041"])}')
print(f'  - BH 0041: {len(df_folhas[df_folhas.origem == "BH 0041"])}')
print(f'Valor total banco: R$ {df_banco["valor"].sum():,.2f}')
print(f'Valor total folhas: R$ {df_folhas["valor"].sum():,.2f}')
print()

# ============ COMPARACAO ============
correspondidos = []
nao_encontrados_banco = []
divergencia_valor = []
divergencia_nome = []
banco_usado = set()

for idx, folha_row in df_folhas.iterrows():
    nome_folha = folha_row['nome_norm']
    valor_folha = folha_row['valor']

    # Match exato por nome normalizado
    match_exato = df_banco[
        (df_banco['nome_norm'] == nome_folha) &
        (~df_banco.index.isin(banco_usado))
    ]

    if len(match_exato) > 0:
        banco_row = match_exato.iloc[0]
        banco_idx = match_exato.index[0]
        banco_usado.add(banco_idx)

        if abs(banco_row['valor'] - valor_folha) < 0.02:
            correspondidos.append({
                'nome_folha': folha_row['nome'],
                'nome_banco': banco_row['nome'],
                'valor_folha': valor_folha,
                'valor_banco': banco_row['valor'],
                'origem': folha_row['origem']
            })
        else:
            divergencia_valor.append({
                'nome_folha': folha_row['nome'],
                'nome_banco': banco_row['nome'],
                'valor_folha': valor_folha,
                'valor_banco': banco_row['valor'],
                'diferenca': banco_row['valor'] - valor_folha,
                'origem': folha_row['origem']
            })
    else:
        # Tentar match parcial (nome truncado no banco)
        match_parcial = None
        for bidx, brow in df_banco[~df_banco.index.isin(banco_usado)].iterrows():
            nome_banco = brow['nome_norm']
            # Banco trunca nomes - verificar inicio igual
            if len(nome_banco) >= 10:
                if nome_folha.startswith(nome_banco) or nome_banco.startswith(nome_folha[:len(nome_banco)]):
                    match_parcial = (bidx, brow)
                    break
            # SOUZA/SOUSA
            nome_folha_v1 = nome_folha.replace('SOUZA', 'SOUSA')
            nome_folha_v2 = nome_folha.replace('SOUSA', 'SOUZA')
            if nome_banco in (nome_folha_v1, nome_folha_v2):
                match_parcial = (bidx, brow)
                break
            # Truncado + SOUZA/SOUSA
            if len(nome_banco) >= 10:
                if nome_folha_v1.startswith(nome_banco) or nome_folha_v2.startswith(nome_banco):
                    match_parcial = (bidx, brow)
                    break

        if match_parcial:
            bidx, brow = match_parcial
            banco_usado.add(bidx)
            if abs(brow['valor'] - valor_folha) < 0.02:
                divergencia_nome.append({
                    'nome_folha': folha_row['nome'],
                    'nome_banco': brow['nome'],
                    'valor': valor_folha,
                    'origem': folha_row['origem'],
                    'status': 'OK (mesmo valor)'
                })
            else:
                divergencia_nome.append({
                    'nome_folha': folha_row['nome'],
                    'nome_banco': brow['nome'],
                    'valor_folha': valor_folha,
                    'valor_banco': brow['valor'],
                    'origem': folha_row['origem'],
                    'status': f'DIVERGENCIA VALOR: dif R$ {brow["valor"] - valor_folha:.2f}'
                })
        else:
            nao_encontrados_banco.append({
                'nome': folha_row['nome'],
                'valor': valor_folha,
                'origem': folha_row['origem']
            })

# Banco sem correspondencia na folha
banco_sem_folha = df_banco[~df_banco.index.isin(banco_usado)]

print('=' * 80)
print(f'CORRESPONDIDOS OK: {len(correspondidos)} registros')
print('=' * 80)
for c in correspondidos:
    print(f'  OK | {c["nome_folha"]} | R$ {c["valor_folha"]:.2f} | {c["origem"]}')

print()
print('=' * 80)
print(f'NAO ENCONTRADOS NO BANCO: {len(nao_encontrados_banco)} registros')
print('=' * 80)
if nao_encontrados_banco:
    for n in nao_encontrados_banco:
        print(f'  FALTA | {n["nome"]} | R$ {n["valor"]:.2f} | {n["origem"]}')
else:
    print('  Nenhum - todos foram encontrados!')

print()
print('=' * 80)
print(f'DIVERGENCIA DE NOME (mas correspondido): {len(divergencia_nome)} registros')
print('=' * 80)
for d in divergencia_nome:
    print(f'  NOME | Folha: {d["nome_folha"]}')
    print(f'         Banco: {d["nome_banco"]}')
    print(f'         {d.get("status", "")} | {d["origem"]}')

print()
print('=' * 80)
print(f'DIVERGENCIA DE VALOR: {len(divergencia_valor)} registros')
print('=' * 80)
if divergencia_valor:
    for d in divergencia_valor:
        print(f'  VALOR | {d["nome_folha"]}')
        print(f'          Folha: R$ {d["valor_folha"]:.2f} | Banco: R$ {d["valor_banco"]:.2f} | Dif: R$ {d["diferenca"]:.2f} | {d["origem"]}')
else:
    print('  Nenhuma divergencia de valor encontrada!')

print()
print('=' * 80)
print(f'NO BANCO MAS NAO NAS FOLHAS: {len(banco_sem_folha)} registros')
print('=' * 80)
for i, row in banco_sem_folha.iterrows():
    print(f'  EXTRA | {row["nome"]} | R$ {row["valor"]:.2f} | Ref: {row["ref_empresa"]}')

print()
print('=' * 80)
print('RESUMO FINAL')
print('=' * 80)
print(f'  Correspondidos OK:           {len(correspondidos)}')
print(f'  Divergencia de nome:         {len(divergencia_nome)}')
print(f'  Divergencia de valor:        {len(divergencia_valor)}')
print(f'  Folha sem banco:             {len(nao_encontrados_banco)}')
print(f'  Banco sem folha:             {len(banco_sem_folha)}')
print(f'  Total folhas:                {len(df_folhas)}')
print(f'  Total banco:                 {len(df_banco)}')
