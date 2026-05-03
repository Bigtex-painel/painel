"""
gerar_payload.py — Reconstrói o payload.js do Painel Bigtex a partir dos xls fonte.

Estratégia de validação: SHA256 chave-por-chave contra o golden file (payload_golden.js).
"""
import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------------------
# Configuração
# --------------------------------------------------------------------------------------

DATA_REF = pd.Timestamp('2026-04-19')
HOJE = '2026-04-19'  # YYYY-MM-DD
DATA_REF_BR = '19/04/2026'  # DD/MM/YYYY
ANOS = [2022, 2023, 2024, 2025, 2026]

PROJECT_DIR = Path('/mnt/project')

# Mapeamento sub-conta → grupo (mensal)
# Atenção: 4.1.04 é "outros" (não "financeiras"), 4.1.99 é "financeiras"
def grupo_mensal(co: str) -> str:
    """Retorna o grupo do bloco mensal para uma sub-conta dada."""
    if co.startswith('1'):
        return 'outras_receitas'
    if co.startswith('2.1'):
        return 'impostos'
    if co.startswith('2.2'):
        return 'cmv_despesas'
    if co.startswith('2.3') or co.startswith('2.4'):
        return 'desp_variaveis_vendas'
    if co.startswith('3.1'):
        return 'financeiras'
    if co == '4.1.04':
        return 'outros'
    if co.startswith('4.1'):
        return 'financeiras'
    if co.startswith('3.2') or co.startswith('3.3'):
        return 'desp_fixas'
    if co.startswith('3.4'):
        return 'desp_pessoal'
    if co.startswith('4.2') or co.startswith('4.3'):
        return 'desp_variaveis_op'
    if co.startswith('5.2') or co.startswith('5.3'):
        return 'investimentos'
    return 'outros'

# Prefixo a adicionar à descrição com base no início da sub-conta
# A descrição armazenada no xls vem sem prefixo; o golden adiciona com base no grupo "pai" (X.Y)
def prefixo_subconta(co: str) -> str:
    """Retorna o prefixo legível para uma sub-conta (ex: '3.4.08' → 'R.H.')."""
    parts = co.split('.')
    if len(parts) < 2:
        return ''
    parent = f'{parts[0]}.{parts[1]}'
    mapping = {
        '1.1': 'VENDAS DE PRODUTOS',
        '1.2': 'VENDAS DE SERVIÇOS',
        '1.4': 'RECEITAS FINANCEIRAS',
        '2.1': 'IMPOSTOS',
        '2.2': 'MERCADORIAS',
        '2.3': 'VENDAS',
        '2.4': 'FINANCEIRAS',
        '3.1': 'FINANCEIRAS',
        '3.2': 'LOGISTICA',
        '3.3': 'OPERACIONAIS',
        '3.4': 'R.H.',
        '4.1': 'FINANCEIRAS',
        '4.2': 'LOGISTICA',
        '4.3': 'OPERACIONAIS',
        '5.2': 'BENS IMOVEIS',
        '5.3': 'BENS MOVEIS',
    }
    return mapping.get(parent, '')

def descricao_mensal(co: str, desc_xls: str) -> str:
    """Constrói a descrição final do desp_cats: '<PREFIXO> <DESCRIÇÃO_XLS>'."""
    pref = prefixo_subconta(co)
    desc_xls = (desc_xls or '').strip()
    if pref:
        return f'{pref} {desc_xls}'
    return desc_xls

# Sócios filtrados POR NOME do fornecedor
SOCIO_THOMPSON = 'THOMPSON JACON CAVALCANTE'
SOCIO_GLAUCO = 'GLAUCO CICERO BARBOSA'

# --------------------------------------------------------------------------------------
# Mapas anuais (cores + categorias agregadas)
# --------------------------------------------------------------------------------------

def cor_anual(co: str) -> str:
    """Cor CSS para uma sub-conta no contexto anual."""
    if co == '1' or co.startswith('1.') or co.startswith('2.2') or co.startswith('5.2'):
        return 'var(--dim)'
    if co.startswith('2.1'):
        return 'var(--red)'
    if co.startswith('2.3') or co.startswith('2.4'):
        return 'var(--yel)'
    if co.startswith('3.1') or co.startswith('4.1'):
        return 'var(--ora)'
    if co.startswith('3.2') or co.startswith('3.3'):
        return 'var(--blu)'
    if co.startswith('3.4'):
        return 'var(--pur)'
    if co.startswith('4.2') or co.startswith('4.3'):
        return 'var(--cya)'
    if co.startswith('5.3'):
        return 'var(--grn)'
    return 'var(--dim)'

# 11 categorias agregadas por grupo X.Y (para desp_cats_ano e desp_cats_tri)
CATEGORIAS_AGREGADAS = [
    ('2.1', 'Impostos',               'var(--red)'),
    ('2.3', 'Comissões/Vendas',       'var(--yel)'),
    ('2.4', 'Taxas cartão/PIX',       'var(--yel)'),
    ('3.1', 'Tarifas bancárias',      'var(--ora)'),
    ('3.2', 'Logística fixa',         'var(--blu)'),
    ('3.3', 'Operacionais fixas',     'var(--blu)'),
    ('3.4', 'Folha/RH',               'var(--pur)'),
    ('4.1', 'IOF',                    'var(--ora)'),
    ('4.2', 'Logística variável',     'var(--cya)'),
    ('4.3', 'Operacionais variáveis', 'var(--cya)'),
    ('5.3', 'Investimentos',          'var(--grn)'),
]
# Tri tem 12 categorias (inclui 5.2 Bens imóveis); Ano tem só 11 (sem 5.2)
CATEGORIAS_AGREGADAS_TRI = CATEGORIAS_AGREGADAS + [
    ('5.2', 'Bens imóveis',           'var(--dim)'),
]
# Sub-contas que entram em desp_cats_full e desp_cats_ano (excluem CMV 2.2 e BENS IMOVEIS 5.2 e COMPRAS 1)
def grupo_pai(co: str) -> str:
    """Retorna o pai X.Y da sub-conta X.Y.Z."""
    parts = co.split('.')
    if len(parts) >= 2:
        return f'{parts[0]}.{parts[1]}'
    return parts[0]

def entra_desp_anual(co: str) -> bool:
    """Filtra sub-contas que entram em desp_cats_* anuais (exclui receitas 1.x e CMV 2.2.x).
    Sub-conta '1' isolada (= COMPRAS) ENTRA. 5.2.x (BENS IMOVEIS) ENTRA."""
    if co.startswith('1.'):  # 1.1, 1.2, 1.4 - receitas
        return False
    if co.startswith('2.2'):  # CMV
        return False
    return True

def entra_desp_ano(co: str) -> bool:
    """Filtra sub-contas que entram em desp_cats_ano (somente as 11 categorias)."""
    pai = grupo_pai(co)
    return pai in {c[0] for c in CATEGORIAS_AGREGADAS}

# --------------------------------------------------------------------------------------
# Carregamento de fontes
# --------------------------------------------------------------------------------------

SHEETS_PAGAR = {2022: '2022', 2023: '2023', 2024: 'Plan1', 2025: 'Plan1', 2026: 'Plan1'}
SHEETS_RECEBER = {2022: 'Plan1', 2023: 'Plan1', 2024: 'Plan1', 2025: 'Plan1', 2026: 'Plan1'}

def _path_xls(tipo: str, ano: int) -> Path:
    """Resolve o path do xls dado o tipo e ano, sendo TOLERANTE a variações de nome:

    Aceita TODAS estas variações (e mais):
      - 2022CONTAS_A_PAGAR_CONSOLIDADO.xlsx
      - 2022-CONTAS_A_PAGAR_CONSOLIDADO.xlsx     (com hífen)
      - 2022_CONTAS_A_PAGAR_CONSOLIDADO.xlsx     (com underscore)
      - 2022.CONTAS_A_PAGAR_CONSOLIDADO.xlsx     (com ponto)
      - 2022 CONTAS_A_PAGAR_CONSOLIDADO.xlsx     (com espaço)

    Para o ano corrente (parcial), aceita também:
      - 2026CONTAS_A_PAGAR_CONSOLIDADO_parcial.xlsx
      - 2026-CONTAS_A_PAGAR_CONSOLIDADO_parcial.xlsx
      - 2026.05.PARCIALCONTAS_A_PAGAR_CONSOLIDADO.xlsx
      - 2026-VENDAS_ANALITICO_parcial.xlsx
      - 2026_05_PARCIALVENDAS_ANALITICO.xlsx
      - etc.

    Em todos os casos, aceita variações de maiúsculas/minúsculas, hífen vs underscore vs ponto.
    """
    sufixos = {
        'pagar': 'CONTAS_A_PAGAR_CONSOLIDADO',
        'receber': 'CONTAS_A_RECEBER_CONSOLIDADO',
        'vendas': 'VENDAS_ANALITICO',
    }
    if tipo not in sufixos:
        raise ValueError(f'tipo desconhecido: {tipo}')
    base = sufixos[tipo]

    # Lista todos os xlsx da pasta uma vez só
    todos_xlsx = list(PROJECT_DIR.glob('*.xlsx')) + list(PROJECT_DIR.glob('*.XLSX'))

    # Função para "normalizar" um nome de arquivo — remove pontuação e deixa em maiúsculas
    # para comparar de forma flexível
    def _normalizar(nome: str) -> str:
        # Tira extensão, deixa maiúsculo, troca separadores por nada
        s = nome.upper().replace('.XLSX', '')
        for sep in ['-', '_', '.', ' ']:
            s = s.replace(sep, '')
        return s

    base_norm = _normalizar(base)
    ano_str = str(ano)
    is_corrente = (ano == DATA_REF.year)

    # ----- Caso 1: ano passado (não corrente) -----
    # Procurar arquivos que comecem com o ano + tenham o "base" no nome (normalizado)
    # E NÃO sejam "parcial"
    if not is_corrente:
        candidatos = []
        for path in todos_xlsx:
            nome_norm = _normalizar(path.stem)
            if nome_norm.startswith(ano_str) and base_norm in nome_norm:
                if 'PARCIAL' in nome_norm:
                    continue  # descartar parciais para anos passados
                candidatos.append(path)
        if candidatos:
            # Se tem mais de um, pega o mais simples (sem PARCIAL/data extra)
            candidatos.sort(key=lambda p: len(p.stem))
            return candidatos[0]
        # fallback de erro claro
        return PROJECT_DIR / f'{ano}{base}.xlsx'

    # ----- Caso 2: ano corrente (com parcial) -----
    # Aceita parciais e não-parciais; preferir parcial se houver
    parciais = []
    nao_parciais = []
    for path in todos_xlsx:
        nome_norm = _normalizar(path.stem)
        if nome_norm.startswith(ano_str) and base_norm in nome_norm:
            if 'PARCIAL' in nome_norm:
                parciais.append(path)
            else:
                nao_parciais.append(path)

    # Preferência: parcial > não-parcial. Entre parciais, o mais recente alfabeticamente
    if parciais:
        parciais.sort()
        return parciais[-1]
    if nao_parciais:
        return nao_parciais[0]

    # fallback
    return PROJECT_DIR / f'{ano}{base}_parcial.xlsx'

def carregar_pagar(anos=ANOS) -> pd.DataFrame:
    """Carrega e normaliza CONTAS_A_PAGAR de todos os anos."""
    dfs = []
    for ano in anos:
        sh = SHEETS_PAGAR[ano]
        df = pd.read_excel(_path_xls('pagar', ano), sheet_name=sh)
        df['_ano_xls'] = ano
        dfs.append(df)
    df = pd.concat(dfs, ignore_index=True)
    for c in ['DT.EMISSÃO', 'DT.VENC.', 'DT.PAGTO.']:
        df[c] = pd.to_datetime(df[c], errors='coerce', dayfirst=False)
    return df

def carregar_receber(anos=ANOS) -> pd.DataFrame:
    """Carrega e normaliza CONTAS_A_RECEBER de todos os anos."""
    dfs = []
    for ano in anos:
        sh = SHEETS_RECEBER[ano]
        df = pd.read_excel(_path_xls('receber', ano), sheet_name=sh)
        df['_ano_xls'] = ano
        dfs.append(df)
    df = pd.concat(dfs, ignore_index=True)
    for c in ['DT.EMISSÃO', 'DT.VENC.', 'DT.PAGTO.']:
        df[c] = pd.to_datetime(df[c], errors='coerce', dayfirst=False)
    return df

def carregar_vendas(anos=ANOS) -> pd.DataFrame:
    """Carrega e normaliza VENDAS_ANALITICO de todos os anos."""
    dfs = []
    for ano in anos:
        df = pd.read_excel(_path_xls('vendas', ano), sheet_name=0)
        df['_ano_xls'] = ano
        dfs.append(df)
    df = pd.concat(dfs, ignore_index=True)
    return df

# --------------------------------------------------------------------------------------
# Validação contra golden
# --------------------------------------------------------------------------------------

def golden_keys() -> dict[str, str]:
    """Carrega o golden via Node e retorna {chave: sha256_16}."""
    code = r'''
const fs=require("fs"),vm=require("vm"),crypto=require("crypto");
const code=fs.readFileSync("/home/claude/payload.js","utf-8");
const ctx={D:null};vm.createContext(ctx);vm.runInContext(code+"\nthis.D=D;",ctx);
const out={};
for (const k of Object.keys(ctx.D)) {
    const j = JSON.stringify(ctx.D[k]);
    out[k] = crypto.createHash("sha256").update(j).digest("hex").slice(0,16);
}
console.log(JSON.stringify(out));
'''
    r = subprocess.run(['node','-e', code], capture_output=True, text=True)
    return json.loads(r.stdout)

def golden_value_json(chave: str) -> str:
    """Retorna a string JSON canônica de uma chave do golden, exatamente como o JSON.stringify gera."""
    code = f'''
const fs=require("fs"),vm=require("vm");
const code=fs.readFileSync("/home/claude/payload.js","utf-8");
const ctx={{D:null}};vm.createContext(ctx);vm.runInContext(code+"\\nthis.D=D;",ctx);
process.stdout.write(JSON.stringify(ctx.D[{json.dumps(chave)}]));
'''
    r = subprocess.run(['node','-e', code], capture_output=True, text=True)
    return r.stdout

def sha256_16(value) -> str:
    """SHA256(16 chars) do JSON.stringify(value) no estilo Node."""
    j = json_stringify(value)
    return hashlib.sha256(j.encode('utf-8')).hexdigest()[:16]

def json_stringify(value) -> str:
    """Mimicar JSON.stringify do Node: sem espaços, ensure_ascii=False."""
    return json.dumps(value, separators=(',', ':'), ensure_ascii=False)

def validar_chave(chave: str, valor, golden: dict[str, str]) -> bool:
    """Compara SHA256 do valor gerado vs golden. Retorna True se bater."""
    sha_meu = sha256_16(valor)
    sha_gold = golden.get(chave, 'AUSENTE')
    ok = sha_meu == sha_gold
    status = '✅' if ok else '❌'
    print(f'  {status} {chave:30s} meu={sha_meu}  golden={sha_gold}')
    return ok

# --------------------------------------------------------------------------------------
# Geradores das chaves
# --------------------------------------------------------------------------------------

def gerar_metas() -> dict:
    """Constante."""
    return {'receita': 450000, 'margem_bruta_pct': 28, 'margem_liquida_pct': 8}

def gerar_meses_disp() -> list[str]:
    """Array de YYYY-MM de jan/2025 até o mês de data_ref."""
    meses = []
    y, m = 2025, 1
    fim_y, fim_m = DATA_REF.year, DATA_REF.month
    while (y, m) <= (fim_y, fim_m):
        meses.append(f'{y:04d}-{m:02d}')
        m += 1
        if m > 12:
            m = 1
            y += 1
    return meses

# --------------------------------------------------------------------------------------
# Bloco mensal — gerar_meses
# --------------------------------------------------------------------------------------

def _round(v, n=2):
    """Arredondamento + conversão para int quando parte decimal é zero (mimica JSON.stringify do JS)."""
    f = round(float(v), n)
    if f == int(f):
        return int(f)
    return f

def _vlr_efetivo(row) -> float:
    """Valor efetivo: VLR.PAGO se LIQUIDADO=S, senão VLR.PARCELA."""
    if row['LIQUIDADO'] == 'S' and pd.notna(row.get('VLR.PAGO')):
        return float(row['VLR.PAGO'])
    return float(row['VLR.PARCELA']) if pd.notna(row['VLR.PARCELA']) else 0.0

def _data_filtro_pagar(row):
    """Filtro mensal: DT.PAGTO (regime de caixa). Linhas LIQUIDADO=N não entram."""
    if row['LIQUIDADO'] == 'S' and pd.notna(row.get('DT.PAGTO.')):
        return row['DT.PAGTO.']
    return pd.NaT  # excluir do filtro mensal

def gerar_mes(ym: str, vendas: pd.DataFrame, pagar_pre: pd.DataFrame) -> dict:
    """Gera o dict de um único mês."""
    ano, mes = int(ym[:4]), int(ym[5:7])

    # ---- VENDAS do mês ----
    v_mes = vendas[vendas['Dt.Emissão'].dt.year.eq(ano) & vendas['Dt.Emissão'].dt.month.eq(mes)].copy()
    receita = float(v_mes['Vlr.Total'].sum())
    cmv = float(v_mes['Vlr.Tot.Custo'].sum())
    lucro_bruto = float(v_mes['Resultado'].sum())  # soma direta para evitar erro de arredondamento
    margem_bruta = round(lucro_bruto / receita * 100, 2) if receita else 0
    qtde = round(float(v_mes['Qtde.'].sum()), 1)

    by_ped = v_mes.groupby('Número', sort=False).agg(
        cliente=('Cliente', 'first'),
        emissao=('Dt.Emissão', 'first'),
        v=('Vlr.Total', 'sum'),
        custo=('Vlr.Tot.Custo', 'sum'),
        i=('Qtde.', 'sum'),
        mg=('Resultado', 'sum'),
    ).reset_index()
    pedidos = len(by_ped)

    # ---- PAGAR do mês ----
    p_mes = pagar_pre[pagar_pre['_dt_filtro'].dt.year.eq(ano) & pagar_pre['_dt_filtro'].dt.month.eq(mes)].copy()
    p_agg = p_mes.groupby(['_co', 'DESCRIÇÃO SUB-CONTA'], sort=False, as_index=False).agg(
        v=('_vlr_ef', 'sum')
    )
    p_agg['g'] = p_agg['_co'].apply(grupo_mensal)

    soma_g = p_agg.groupby('g')['v'].sum().to_dict()
    impostos = _round(soma_g.get('impostos', 0))
    desp_pessoal = _round(soma_g.get('desp_pessoal', 0))
    desp_fixas = _round(soma_g.get('desp_fixas', 0))
    desp_var_vendas = _round(soma_g.get('desp_variaveis_vendas', 0))
    desp_var_op = _round(soma_g.get('desp_variaveis_op', 0))
    financeiras = _round(soma_g.get('financeiras', 0))
    investimentos = _round(soma_g.get('investimentos', 0))

    desp_operacional = _round(desp_pessoal + desp_fixas + desp_var_vendas + desp_var_op)
    ebitda = _round(lucro_bruto - impostos - desp_operacional)
    resultado = _round(ebitda - financeiras)
    margem_liquida = _round(resultado / receita * 100) if receita else 0

    # desp_cats
    p_agg_sorted = p_agg.sort_values('v', ascending=False, kind='mergesort')
    desp_cats = [
        {'d': descricao_mensal(r['_co'], r['DESCRIÇÃO SUB-CONTA']), 'co': r['_co'], 'g': r['g'], 'v': _round(r['v'])}
        for _, r in p_agg_sorted.iterrows()
    ]

    # clientes top
    last_idx_cli = v_mes.groupby('Cliente', sort=False)['Cliente'].apply(lambda g: g.index.max())
    cli_agg = v_mes.groupby('Cliente', sort=False, as_index=False).agg(
        fat=('Vlr.Total', 'sum'),
        margem=('Resultado', 'sum'),
    )
    cli_ped = by_ped.groupby('cliente').size().to_dict()
    cli_agg['pedidos'] = cli_agg['Cliente'].map(cli_ped)
    cli_agg['_last_idx'] = cli_agg['Cliente'].map(last_idx_cli)
    cli_agg['margem_pct'] = (cli_agg['margem'] / cli_agg['fat'] * 100).round(1)
    cli_agg = cli_agg.sort_values(['fat', 'margem', '_last_idx'], ascending=[False, True, False], kind='mergesort')
    clientes = [
        {'c': r['Cliente'].rstrip()[:50], 'fat': _round(r['fat']), 'margem': _round(r['margem']),
         'margem_pct': _round(r['margem_pct'], 1), 'pedidos': int(r['pedidos'])}
        for _, r in cli_agg.iterrows()
    ]

    # produtos top
    v_mes_p = v_mes.copy()
    v_mes_p['_prod'] = v_mes_p['Descrição do Produto'].astype(str).str.rstrip()
    last_idx = v_mes_p.groupby('_prod', sort=False)['_prod'].apply(lambda g: g.index.max())
    prod_agg = v_mes_p.groupby('_prod', sort=False, as_index=False).agg(
        fat=('Vlr.Total', 'sum'),
        qtde=('Qtde.', 'sum'),
        margem=('Resultado', 'sum'),
        n_pedidos=('Número', 'nunique'),
    )
    prod_agg['_last_idx'] = prod_agg['_prod'].map(last_idx)
    prod_agg['margem_pct'] = (prod_agg['margem'] / prod_agg['fat'] * 100).round(1)
    prod_agg = prod_agg.sort_values(['fat', 'margem', '_last_idx'],
                                    ascending=[False, False, False],
                                    kind='mergesort')
    produtos = [
        {'p': r['_prod'][:70], 'fat': _round(r['fat']), 'qtde': _round(r['qtde'], 1),
         'margem': _round(r['margem']), 'margem_pct': _round(r['margem_pct'], 1)}
        for _, r in prod_agg.iterrows()
    ]

    # fat_diario
    by_ped['_d'] = by_ped['emissao'].dt.day
    by_ped['_dow'] = by_ped['emissao'].dt.dayofweek
    fd = by_ped.groupby('_d', sort=True).agg(
        v=('v', 'sum'),
        n=('Número', 'count'),
        dow=('_dow', 'first'),
    ).reset_index()
    fat_diario = [
        {'d': int(r['_d']), 'dow': int(r['dow']), 'v': _round(r['v']), 'n': int(r['n'])}
        for _, r in fd.iterrows()
    ]

    # ped_dia (cada dia ordenado por v DESC)
    ped_dia = {}
    for d, grupo in by_ped.groupby('_d', sort=True):
        grupo_sorted = grupo.sort_values('v', ascending=False, kind='mergesort')
        ped_dia[str(int(d))] = [
            {'n': int(r['Número']), 'c': r['cliente'], 'v': _round(r['v']),
             'i': _round(r['i'], 1), 'mg': _round(r['mg']),
             'mg_pct': _round(r['mg'] / r['v'] * 100, 1) if r['v'] else 0}
            for _, r in grupo_sorted.iterrows()
        ]

    return {
        'receita': _round(receita), 'cmv': _round(cmv), 'lucro_bruto': _round(lucro_bruto),
        'margem_bruta': margem_bruta, 'impostos': impostos, 'desp_pessoal': desp_pessoal,
        'desp_fixas': desp_fixas, 'desp_var_vendas': desp_var_vendas, 'desp_var_op': desp_var_op,
        'desp_operacional': desp_operacional, 'financeiras': financeiras,
        'investimentos': investimentos, 'ebitda': ebitda, 'resultado': resultado,
        'margem_liquida': margem_liquida, 'pedidos': pedidos, 'qtde': qtde,
        'clientes': clientes, 'produtos': produtos, 'fat_diario': fat_diario,
        'desp_cats': desp_cats, 'ped_dia': ped_dia,
    }

def _preparar_pagar(pagar: pd.DataFrame) -> pd.DataFrame:
    """Adiciona colunas auxiliares uma vez só."""
    p = pagar.copy()
    p['_dt_filtro'] = p.apply(_data_filtro_pagar, axis=1)
    p['_vlr_ef'] = p.apply(_vlr_efetivo, axis=1)
    # SUB-CONTA NaN → '1' (lançamentos COMPRAS que aparecem no golden como g='1')
    p['_co'] = p['SUB-CONTA'].astype('object').where(p['SUB-CONTA'].notna(), '1').astype(str)
    # DESCRIÇÃO SUB-CONTA NaN → 'COMPRAS' (para que groupby não exclua essas linhas)
    p['DESCRIÇÃO SUB-CONTA'] = p['DESCRIÇÃO SUB-CONTA'].fillna('COMPRAS')
    return p

def gerar_meses(vendas: pd.DataFrame, pagar: pd.DataFrame) -> dict:
    """Gera dict completo {ym: dados}."""
    pagar_pre = _preparar_pagar(pagar)
    return {ym: gerar_mes(ym, vendas, pagar_pre) for ym in gerar_meses_disp()}

# --------------------------------------------------------------------------------------
# Bloco anual — chaves simples
# --------------------------------------------------------------------------------------

def gerar_data_ref() -> str:
    return DATA_REF_BR  # "19/04/2026"

def gerar_hoje() -> str:
    return HOJE  # "2026-04-19"

def gerar_ytd_mes() -> int:
    return DATA_REF.month  # 4

def gerar_anos() -> list[int]:
    return list(ANOS)

def gerar_tri_completos() -> list[str]:
    """Lista de 'YYYY-qN' para todos os trimestres COMPLETOS antes de DATA_REF."""
    fim_y = DATA_REF.year
    fim_q = (DATA_REF.month - 1) // 3 + 1  # trimestre em curso
    out = []
    for y in ANOS:
        for q in (1, 2, 3, 4):
            # Inclui se for ano anterior ao atual OU (ano atual e q < q_atual)
            if y < fim_y or (y == fim_y and q < fim_q):
                out.append(f'{y}-q{q}')
    return out

# --------------------------------------------------------------------------------------
# Bloco anual — KPIs por período (ano_full, ano_ytd, ano_tri, ym_data)
# --------------------------------------------------------------------------------------

def _vendas_filtro(vendas: pd.DataFrame, dt_ini: pd.Timestamp, dt_fim: pd.Timestamp) -> pd.DataFrame:
    """Filtra vendas por Dt.Emissão entre [dt_ini, dt_fim] inclusive."""
    return vendas[(vendas['Dt.Emissão'] >= dt_ini) & (vendas['Dt.Emissão'] <= dt_fim)]

def _pagar_filtro(pagar_pre: pd.DataFrame, dt_ini: pd.Timestamp, dt_fim: pd.Timestamp) -> pd.DataFrame:
    """Filtra pagar por _dt_filtro (DT.PAGTO) entre [dt_ini, dt_fim] inclusive."""
    return pagar_pre[(pagar_pre['_dt_filtro'] >= dt_ini) & (pagar_pre['_dt_filtro'] <= dt_fim)]

def _kpis_periodo(vendas_p: pd.DataFrame, pagar_p: pd.DataFrame, com_socios: bool = True) -> dict:
    """Calcula KPIs anuais para um período."""
    receita = float(vendas_p['Vlr.Total'].sum())
    cmv = float(vendas_p['Vlr.Tot.Custo'].sum())
    lucro_bruto = float(vendas_p['Resultado'].sum())
    margem_bruta = round(lucro_bruto / receita * 100, 2) if receita else 0
    qtde = round(float(vendas_p['Qtde.'].sum()), 2)
    pedidos = int(vendas_p['Número'].nunique())
    clientes_unicos = int(vendas_p['Cliente'].nunique())
    ticket_medio = round(receita / pedidos, 2) if pedidos else 0

    # Despesas por grupo (estilo mensal)
    p_g = pagar_p.copy()
    p_g['_grupo'] = p_g['_co'].apply(grupo_mensal)
    soma_g = p_g.groupby('_grupo')['_vlr_ef'].sum().to_dict()
    impostos = _round(soma_g.get('impostos', 0))
    desp_pessoal = _round(soma_g.get('desp_pessoal', 0))
    desp_fixas = _round(soma_g.get('desp_fixas', 0))
    desp_var_vendas = _round(soma_g.get('desp_variaveis_vendas', 0))
    desp_var_op = _round(soma_g.get('desp_variaveis_op', 0))
    # No ANUAL, "outros" (4.1.04 IOF) entra em "financeiras"
    financeiras = _round(soma_g.get('financeiras', 0) + soma_g.get('outros', 0))
    investimentos = _round(soma_g.get('investimentos', 0))

    desp_operacional = _round(desp_pessoal + desp_fixas + desp_var_vendas + desp_var_op)
    # ANUAL: investimentos ENTRA no resultado (diferença do mensal)
    resultado = _round(lucro_bruto - impostos - desp_operacional - financeiras - investimentos)
    margem_liquida = _round(resultado / receita * 100) if receita else 0

    out = {
        'receita': _round(receita),
        'cmv': _round(cmv),
        'lucro_bruto': _round(lucro_bruto),
        'margem_bruta': margem_bruta,
        'pedidos': pedidos,
        'qtde': qtde,
        'clientes_unicos': clientes_unicos,
        'ticket_medio': ticket_medio,
        'impostos': impostos,
        'desp_pessoal': desp_pessoal,
        'desp_fixas': desp_fixas,
        'desp_var_vendas': desp_var_vendas,
        'desp_var_op': desp_var_op,
        'desp_operacional': desp_operacional,
        'financeiras': financeiras,
        'investimentos': investimentos,
        'resultado': resultado,
        'margem_liquida': margem_liquida,
    }

    if com_socios:
        # ANUAL: pro_labore = sub-conta 3.4.07, comissao_socio = sub-conta 2.3.02
        pro_labore = _round(float(pagar_p[pagar_p['_co'] == '3.4.07']['_vlr_ef'].sum()))
        comissao_socio = _round(float(pagar_p[pagar_p['_co'] == '2.3.02']['_vlr_ef'].sum()))
        out['pro_labore'] = pro_labore
        out['comissao_socio'] = comissao_socio
        out['retiradas_socios'] = _round(pro_labore + comissao_socio)

    return out

    return out

def gerar_ano_full(vendas: pd.DataFrame, pagar_pre: pd.DataFrame) -> dict:
    """ano_full[ano]: KPIs do ano inteiro (jan-dez)."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        dt_fim = pd.Timestamp(f'{ano}-12-31 23:59:59')
        # Para 2026, limitar ao data_ref
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
        v = _vendas_filtro(vendas, dt_ini, dt_fim)
        p = _pagar_filtro(pagar_pre, dt_ini, dt_fim)
        out[str(ano)] = _kpis_periodo(v, p, com_socios=True)
    return out

def gerar_ano_ytd(vendas: pd.DataFrame, pagar_pre: pd.DataFrame) -> dict:
    """ano_ytd[ano]: KPIs do ano até o mês de data_ref (ano-corrente: dia exato)."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
        else:
            # Mesmo dia do data_ref no ano (mas para mês completo? ou dia exato?)
            # Tentativa: até o último dia do mês de data_ref
            mes = DATA_REF.month
            ultimo_dia = pd.Timestamp(f'{ano}-{mes:02d}-01') + pd.offsets.MonthEnd(0)
            dt_fim = ultimo_dia.replace(hour=23, minute=59, second=59)
        v = _vendas_filtro(vendas, dt_ini, dt_fim)
        p = _pagar_filtro(pagar_pre, dt_ini, dt_fim)
        out[str(ano)] = _kpis_periodo(v, p, com_socios=True)
    return out

def gerar_ano_tri(vendas: pd.DataFrame, pagar_pre: pd.DataFrame) -> dict:
    """ano_tri[ano].qN: KPIs por trimestre."""
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            mes_ini = (q - 1) * 3 + 1
            mes_fim = q * 3
            dt_ini = pd.Timestamp(f'{ano}-{mes_ini:02d}-01')
            dt_fim = pd.Timestamp(f'{ano}-{mes_fim:02d}-01') + pd.offsets.MonthEnd(0)
            dt_fim = dt_fim.replace(hour=23, minute=59, second=59)
            # Para o ano corrente, limitar ao data_ref (q3, q4 podem zerar)
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                # Trimestre futuro: zerar
                out[str(ano)][f'q{q}'] = _zeros_kpi()
                continue
            if ano == DATA_REF.year and dt_fim > DATA_REF:
                dt_fim = DATA_REF
            v = _vendas_filtro(vendas, dt_ini, dt_fim)
            p = _pagar_filtro(pagar_pre, dt_ini, dt_fim)
            out[str(ano)][f'q{q}'] = _kpis_periodo(v, p, com_socios=True)
    return out

def _zeros_kpi() -> dict:
    return {
        'receita': 0, 'cmv': 0, 'lucro_bruto': 0, 'margem_bruta': 0,
        'pedidos': 0, 'qtde': 0, 'clientes_unicos': 0, 'ticket_medio': 0,
        'impostos': 0, 'desp_pessoal': 0, 'desp_fixas': 0, 'desp_var_vendas': 0,
        'desp_var_op': 0, 'desp_operacional': 0, 'financeiras': 0, 'investimentos': 0,
        'resultado': 0, 'margem_liquida': 0,
        'pro_labore': 0, 'comissao_socio': 0, 'retiradas_socios': 0,
    }

# --------------------------------------------------------------------------------------
# Bloco anual — desp_cats_full / desp_cats_ytd / desp_cats_ano (e variantes _tri)
# --------------------------------------------------------------------------------------

def descricao_anual(co: str, desc_xls: str) -> str:
    """Descrição final no formato anual (preserva trailing whitespace do xls)."""
    pref = prefixo_subconta(co)
    desc_xls = desc_xls or ''
    if pref:
        return f'{pref} {desc_xls}'
    return desc_xls

def _desp_cats_detalhado(pagar_p: pd.DataFrame, full_tri: bool = False) -> list[dict]:
    """desp_cats_full / desp_cats_ytd / desp_cats_full_tri: detalhado por sub-conta com cor.
    full_tri=True: inclui TUDO menos co='1'; descrição SEM prefixo e com strip()."""
    if full_tri:
        # exclui apenas co='1' (COMPRAS isolado); 1.x e 2.2.x ENTRAM
        p = pagar_p[pagar_p['_co'] != '1'].copy()
    else:
        p = pagar_p[pagar_p['_co'].apply(entra_desp_anual)].copy()
    p_agg = p.groupby(['_co', 'DESCRIÇÃO SUB-CONTA'], sort=False, as_index=False).agg(
        v=('_vlr_ef', 'sum')
    )
    p_agg = p_agg.sort_values('v', ascending=False, kind='mergesort')
    if full_tri:
        return [
            {
                'd': (r['DESCRIÇÃO SUB-CONTA'] or '').strip(),
                'co': cor_anual(r['_co']),
                'g': r['_co'],
                'v': _round(r['v']),
            }
            for _, r in p_agg.iterrows()
        ]
    return [
        {
            'd': descricao_anual(r['_co'], r['DESCRIÇÃO SUB-CONTA']),
            'co': cor_anual(r['_co']),
            'g': r['_co'],
            'v': _round(r['v']),
        }
        for _, r in p_agg.iterrows()
    ]

def _desp_cats_agregado(pagar_p: pd.DataFrame, categorias=None) -> list[dict]:
    """desp_cats_ano / desp_cats_tri: categorias por grupo X.Y.
    Categorias com v=0 são omitidas. Default é 11 cats (CATEGORIAS_AGREGADAS)."""
    if categorias is None:
        categorias = CATEGORIAS_AGREGADAS
    pais_validos = {c[0] for c in categorias}
    p = pagar_p[pagar_p['_co'].apply(lambda co: grupo_pai(co) in pais_validos)].copy()
    p['_pai'] = p['_co'].apply(grupo_pai)
    soma = p.groupby('_pai')['_vlr_ef'].sum().to_dict()
    out = []
    for pai, descricao, cor in categorias:
        v = _round(soma.get(pai, 0))
        if v == 0:
            continue
        out.append({'d': descricao, 'co': cor, 'g': pai, 'v': v})
    out.sort(key=lambda x: -x['v'])
    return out

def gerar_desp_cats_full(pagar_pre: pd.DataFrame) -> dict:
    """Por ano FULL."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        dt_fim = pd.Timestamp(f'{ano}-12-31 23:59:59')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
        p = _pagar_filtro(pagar_pre, dt_ini, dt_fim)
        out[str(ano)] = _desp_cats_detalhado(p)
    return out

def gerar_desp_cats_ytd(pagar_pre: pd.DataFrame) -> dict:
    """Por ano YTD."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
        else:
            mes = DATA_REF.month
            dt_fim = (pd.Timestamp(f'{ano}-{mes:02d}-01') + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
        p = _pagar_filtro(pagar_pre, dt_ini, dt_fim)
        out[str(ano)] = _desp_cats_detalhado(p)
    return out

def gerar_desp_cats_ano(pagar_pre: pd.DataFrame) -> dict:
    """Por ano FULL — agregado em 11 categorias."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        dt_fim = pd.Timestamp(f'{ano}-12-31 23:59:59')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
        p = _pagar_filtro(pagar_pre, dt_ini, dt_fim)
        out[str(ano)] = _desp_cats_agregado(p)
    return out

def gerar_desp_cats_tri(pagar_pre: pd.DataFrame) -> dict:
    """Por trimestre — agregado em 11 categorias."""
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            mes_ini = (q - 1) * 3 + 1
            mes_fim = q * 3
            dt_ini = pd.Timestamp(f'{ano}-{mes_ini:02d}-01')
            dt_fim = (pd.Timestamp(f'{ano}-{mes_fim:02d}-01') + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                out[str(ano)][f'q{q}'] = []
                continue
            if ano == DATA_REF.year and dt_fim > DATA_REF:
                dt_fim = DATA_REF
            p = _pagar_filtro(pagar_pre, dt_ini, dt_fim)
            agg = _desp_cats_agregado(p, CATEGORIAS_AGREGADAS_TRI)
            out[str(ano)][f'q{q}'] = agg
    return out

def gerar_desp_cats_full_tri(pagar_pre: pd.DataFrame) -> dict:
    """Por trimestre — detalhado."""
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            mes_ini = (q - 1) * 3 + 1
            mes_fim = q * 3
            dt_ini = pd.Timestamp(f'{ano}-{mes_ini:02d}-01')
            dt_fim = (pd.Timestamp(f'{ano}-{mes_fim:02d}-01') + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                out[str(ano)][f'q{q}'] = []
                continue
            if ano == DATA_REF.year and dt_fim > DATA_REF:
                dt_fim = DATA_REF
            p = _pagar_filtro(pagar_pre, dt_ini, dt_fim)
            out[str(ano)][f'q{q}'] = _desp_cats_detalhado(p, full_tri=True)
    return out

# --------------------------------------------------------------------------------------
# Bloco anual — ym_data (KPIs reduzidos por mês, 52 meses)
# --------------------------------------------------------------------------------------

def gerar_ym_data(vendas: pd.DataFrame, pagar_pre: pd.DataFrame) -> dict:
    """ym_data[ym]: KPIs reduzidos por mês para todos os meses até DATA_REF.
    REGRA ESPECIAL: lucro_bruto = receita - cmv (subtração de SOMAS), não sum(Resultado)."""
    out = {}
    for ano in ANOS:
        for mes in range(1, 13):
            dt_ini = pd.Timestamp(f'{ano}-{mes:02d}-01')
            if dt_ini > DATA_REF:
                continue
            dt_fim = (dt_ini + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
            if dt_fim > DATA_REF:
                dt_fim = DATA_REF
            v = _vendas_filtro(vendas, dt_ini, dt_fim)
            p = _pagar_filtro(pagar_pre, dt_ini, dt_fim)
            ym = f'{ano}-{mes:02d}'
            # Cálculos próprios (lucro_bruto = receita - cmv)
            receita_f = float(v['Vlr.Total'].sum())
            cmv_f = float(v['Vlr.Tot.Custo'].sum())
            lb_f = receita_f - cmv_f
            margem_bruta = round(lb_f / receita_f * 100, 2) if receita_f else 0
            pedidos = int(v['Número'].nunique())
            clientes = int(v['Cliente'].nunique())
            # Despesas
            pg = p.copy()
            pg['_grupo'] = pg['_co'].apply(grupo_mensal)
            soma_g = pg.groupby('_grupo')['_vlr_ef'].sum().to_dict()
            imp_f = soma_g.get('impostos', 0)
            dop_f = soma_g.get('desp_pessoal', 0) + soma_g.get('desp_fixas', 0) + \
                    soma_g.get('desp_variaveis_vendas', 0) + soma_g.get('desp_variaveis_op', 0)
            fin_f = soma_g.get('financeiras', 0) + soma_g.get('outros', 0)
            inv_f = soma_g.get('investimentos', 0)
            res_f = lb_f - imp_f - dop_f - fin_f - inv_f
            margem_liquida = _round(res_f / receita_f * 100) if receita_f else 0
            # Sócios
            pro_labore = _round(float(p[p['_co'] == '3.4.07']['_vlr_ef'].sum()))
            comissao_socio = _round(float(p[p['_co'] == '2.3.02']['_vlr_ef'].sum()))
            out[ym] = {
                'receita': _round(receita_f),
                'cmv': _round(cmv_f),
                'lucro_bruto': _round(lb_f),
                'margem_bruta': margem_bruta,
                'resultado': _round(res_f),
                'margem_liquida': margem_liquida,
                'pedidos': pedidos,
                'clientes': clientes,
                'pro_labore': pro_labore,
                'comissao_socio': comissao_socio,
            }
    return out

# --------------------------------------------------------------------------------------
# Bloco anual — top_cli / top_prod / top_cli_tri / top_prod_tri
# --------------------------------------------------------------------------------------

def _trunc(s: str, n: int = 60) -> str:
    return (s or '')[:n]

def _periodo_full_ou_ytd(ano: int):
    """Para o ano corrente, retorna janela YTD; para outros, ano cheio."""
    dt_ini = pd.Timestamp(f'{ano}-01-01')
    dt_fim = pd.Timestamp(f'{ano}-12-31 23:59:59')
    if ano == DATA_REF.year:
        dt_fim = DATA_REF
    return dt_ini, dt_fim

def _periodo_anterior(ano: int):
    """Período anterior para comparar fat_ant em top_cli/top_prod.
    Se ano corrente é YTD, anterior é mesmo MÊS YTD do ano anterior (mês completo, fechado).
    Caso contrário, ano anterior CHEIO."""
    if ano == DATA_REF.year:
        # YTD: mesmo mês do ano anterior, mas FECHADO (último dia do mês)
        dt_ini = pd.Timestamp(f'{ano - 1}-01-01')
        dt_fim = (pd.Timestamp(year=ano - 1, month=DATA_REF.month, day=1) +
                  pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
    else:
        dt_ini = pd.Timestamp(f'{ano - 1}-01-01')
        dt_fim = pd.Timestamp(f'{ano - 1}-12-31 23:59:59')
    return dt_ini, dt_fim

def _periodo_tri(ano: int, q: int):
    mes_ini = (q - 1) * 3 + 1
    mes_fim = q * 3
    dt_ini = pd.Timestamp(f'{ano}-{mes_ini:02d}-01')
    dt_fim = (pd.Timestamp(f'{ano}-{mes_fim:02d}-01') + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
    if ano == DATA_REF.year and dt_fim > DATA_REF:
        dt_fim = DATA_REF
    return dt_ini, dt_fim

def _top_cli_periodo(vendas: pd.DataFrame, dt_ini, dt_fim, vendas_ant, dt_ini_ant, dt_fim_ant,
                     com_pedidos_no_fim: bool = False) -> list[dict]:
    """top_cli para um período (até 50 entries, ordenado por fat desc, tie-break por _last_idx desc).
    com_pedidos_no_fim=True: usado em top_cli_tri (chave pedidos no fim do dict)."""
    v_p = _vendas_filtro(vendas, dt_ini, dt_fim)
    if len(v_p) == 0:
        return []
    v_p = v_p.copy()
    v_p['_idx'] = v_p.index
    cli_agg = v_p.groupby('Cliente', sort=False, as_index=False).agg(
        fat=('Vlr.Total', 'sum'),
        margem=('Resultado', 'sum'),
        pedidos=('Número', 'nunique'),
        _last_idx=('_idx', 'max'),
    )
    cli_agg['margem_pct'] = (cli_agg['margem'] / cli_agg['fat'] * 100).round(2)
    # Período anterior
    v_ant = _vendas_filtro(vendas_ant, dt_ini_ant, dt_fim_ant)
    fat_ant_map = v_ant.groupby('Cliente')['Vlr.Total'].sum().to_dict() if len(v_ant) > 0 else {}
    cli_agg['fat_ant'] = cli_agg['Cliente'].map(fat_ant_map).fillna(0)
    cli_agg = cli_agg.sort_values(['fat', '_last_idx'], ascending=[False, False], kind='mergesort').head(50)
    out = []
    for _, r in cli_agg.iterrows():
        fat = _round(r['fat'])
        fat_ant = _round(r['fat_ant'])
        delta_pct = _round((fat - fat_ant) / fat_ant * 100, 1) if fat_ant else None
        item = {
            'c': _trunc(r['Cliente'].rstrip()),
            'fat': fat,
            'margem': _round(r['margem']),
            'margem_pct': _round(r['margem_pct'], 2),
        }
        if com_pedidos_no_fim:
            item['fat_ant'] = fat_ant
            item['delta_pct'] = delta_pct
            item['pedidos'] = int(r['pedidos'])
        else:
            item['pedidos'] = int(r['pedidos'])
            item['fat_ant'] = fat_ant
            item['delta_pct'] = delta_pct
        out.append(item)
    return out

def _top_prod_periodo(vendas: pd.DataFrame, dt_ini, dt_fim, vendas_ant, dt_ini_ant, dt_fim_ant,
                      com_qtde_no_fim: bool = False) -> list[dict]:
    """top_prod para um período (até 50 entries, ordenado por fat desc).
    com_qtde_no_fim=True: top_prod_tri (qtde no fim do dict, sem truncamento de p)."""
    v_p = _vendas_filtro(vendas, dt_ini, dt_fim)
    if len(v_p) == 0:
        return []
    v_p = v_p.copy()
    v_p['_prod'] = v_p['Descrição do Produto'].astype(str).str.rstrip()
    v_p['_idx'] = v_p.index
    prod_agg = v_p.groupby('_prod', sort=False, as_index=False).agg(
        fat=('Vlr.Total', 'sum'),
        qtde=('Qtde.', 'sum'),
        margem=('Resultado', 'sum'),
        _first_idx=('_idx', 'min'),
    )
    prod_agg['margem_pct'] = (prod_agg['margem'] / prod_agg['fat'] * 100).round(2)
    # Período anterior
    v_ant = _vendas_filtro(vendas_ant, dt_ini_ant, dt_fim_ant)
    if len(v_ant) > 0:
        v_ant = v_ant.copy()
        v_ant['_prod'] = v_ant['Descrição do Produto'].astype(str).str.rstrip()
        fat_ant_map = v_ant.groupby('_prod')['Vlr.Total'].sum().to_dict()
    else:
        fat_ant_map = {}
    prod_agg['fat_ant'] = prod_agg['_prod'].map(fat_ant_map).fillna(0)
    # tie-break: fat desc, _first_idx asc (descrição que apareceu primeiro fica antes)
    prod_agg = prod_agg.sort_values(['fat', '_first_idx'], ascending=[False, True], kind='mergesort').head(50)
    out = []
    for _, r in prod_agg.iterrows():
        fat = _round(r['fat'])
        fat_ant = _round(r['fat_ant'])
        delta_pct = _round((fat - fat_ant) / fat_ant * 100, 1) if fat_ant else None
        # top_prod (anual) trunca em 60; top_prod_tri NÃO trunca
        p_str = r['_prod'] if com_qtde_no_fim else _trunc(r['_prod'])
        item = {
            'p': p_str,
            'fat': fat,
        }
        if com_qtde_no_fim:
            # top_prod_tri: p, fat, margem, margem_pct, fat_ant, delta_pct, qtde
            item['margem'] = _round(r['margem'])
            item['margem_pct'] = _round(r['margem_pct'], 2)
            item['fat_ant'] = fat_ant
            item['delta_pct'] = delta_pct
            item['qtde'] = _round(r['qtde'], 2)
        else:
            # top_prod (anual): p, fat, qtde, margem, margem_pct, fat_ant, delta_pct
            item['qtde'] = _round(r['qtde'], 2)
            item['margem'] = _round(r['margem'])
            item['margem_pct'] = _round(r['margem_pct'], 2)
            item['fat_ant'] = fat_ant
            item['delta_pct'] = delta_pct
        out.append(item)
    return out

def gerar_top_cli(vendas: pd.DataFrame) -> dict:
    out = {}
    for ano in ANOS:
        dt_ini, dt_fim = _periodo_full_ou_ytd(ano)
        dt_ini_ant, dt_fim_ant = _periodo_anterior(ano)
        out[str(ano)] = _top_cli_periodo(vendas, dt_ini, dt_fim, vendas, dt_ini_ant, dt_fim_ant)
    return out

def gerar_top_prod(vendas: pd.DataFrame) -> dict:
    out = {}
    for ano in ANOS:
        dt_ini, dt_fim = _periodo_full_ou_ytd(ano)
        dt_ini_ant, dt_fim_ant = _periodo_anterior(ano)
        out[str(ano)] = _top_prod_periodo(vendas, dt_ini, dt_fim, vendas, dt_ini_ant, dt_fim_ant)
    return out

def gerar_top_cli_tri(vendas: pd.DataFrame) -> dict:
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            dt_ini, dt_fim = _periodo_tri(ano, q)
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                out[str(ano)][f'q{q}'] = []
                continue
            dt_ini_ant, dt_fim_ant = _periodo_tri(ano - 1, q)
            out[str(ano)][f'q{q}'] = _top_cli_periodo(
                vendas, dt_ini, dt_fim, vendas, dt_ini_ant, dt_fim_ant, com_pedidos_no_fim=True
            )
    return out

def gerar_top_prod_tri(vendas: pd.DataFrame) -> dict:
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            dt_ini, dt_fim = _periodo_tri(ano, q)
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                out[str(ano)][f'q{q}'] = []
                continue
            dt_ini_ant, dt_fim_ant = _periodo_tri(ano - 1, q)
            out[str(ano)][f'q{q}'] = _top_prod_periodo(
                vendas, dt_ini, dt_fim, vendas, dt_ini_ant, dt_fim_ant, com_qtde_no_fim=True
            )
    return out

# --------------------------------------------------------------------------------------
# Bloco anual — coorte (clientes novos / retornados / recorrentes / perdidos)
# --------------------------------------------------------------------------------------

NAN_TOKEN = '__NaN__'

def _cli_set(vendas: pd.DataFrame, dt_ini, dt_fim) -> set:
    """Set de clientes únicos no período (NaN tratado como token)."""
    sub = vendas[(vendas['Dt.Emissão'] >= dt_ini) & (vendas['Dt.Emissão'] <= dt_fim)]
    return set(sub['Cliente'].fillna(NAN_TOKEN).unique())

def _coorte_calc(period: set, prev: set, all_pre: set, total_inclui_nan: bool,
                  excluir_nan_categorias: bool = True) -> dict:
    """Calcula as 5 chaves de uma coorte.
    - total: len(period) com ou sem NaN (depende de total_inclui_nan)
    - novos: SEMPRE exclui NaN
    - rec/ret/perd: exclui NaN se excluir_nan_categorias=True (padrão tri/ytd),
      caso contrário inclui (padrão _full)"""
    if total_inclui_nan:
        total = len(period)
    else:
        total = len(period - {NAN_TOKEN})
    novos = (period - all_pre) - {NAN_TOKEN}
    if excluir_nan_categorias:
        rec = (period & prev) - {NAN_TOKEN}
        ret = ((period & all_pre) - prev) - {NAN_TOKEN}
        perd = (prev - period) - {NAN_TOKEN}
    else:
        rec = period & prev
        ret = (period & all_pre) - prev
        perd = prev - period
    return {
        'total': total,
        'novos': len(novos),
        'retornados': len(ret),
        'recorrentes': len(rec),
        'perdidos': len(perd),
    }

def gerar_coorte_full(vendas: pd.DataFrame) -> dict:
    """coorte_full[ano]: 4 categorias para cada ano (vs ano anterior INTEIRO).
    total INCLUI NaN."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        dt_fim = pd.Timestamp(f'{ano}-12-31 23:59:59')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
        period = _cli_set(vendas, dt_ini, dt_fim)
        prev = _cli_set(vendas, pd.Timestamp(f'{ano - 1}-01-01'), pd.Timestamp(f'{ano - 1}-12-31 23:59:59'))
        all_pre = _cli_set(vendas, pd.Timestamp('1970-01-01'), dt_ini - pd.Timedelta(seconds=1))
        out[str(ano)] = _coorte_calc(period, prev, all_pre, total_inclui_nan=True,
                                       excluir_nan_categorias=False)
    return out

def gerar_coorte_ytd(vendas: pd.DataFrame) -> dict:
    """coorte_ytd[ano]: período = jan a fim do mês ytd (DATA_REF para ano corrente, mês fechado p/ outros).
    Comparações:
      - prev_ytd: SEMPRE mês fechado (jan a fim do mês ytd_mes do ano anterior)
      - rec/perd: usam prev_ytd (perd inclui NaN)
      - ret = period ∩ (∪ todos YTDs anteriores ao y-1) - prev_ytd
    total INCLUI NaN."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
        else:
            mes = DATA_REF.month
            dt_fim = (pd.Timestamp(f'{ano}-{mes:02d}-01') + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
        # prev_ytd: mês fechado correspondente ao ytd_mes
        prev_dt_fim = (pd.Timestamp(f'{ano - 1}-{DATA_REF.month:02d}-01') +
                       pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
        period = _cli_set(vendas, dt_ini, dt_fim)
        prev_ytd = _cli_set(vendas, pd.Timestamp(f'{ano - 1}-01-01'), prev_dt_fim)
        all_pre = _cli_set(vendas, pd.Timestamp('1970-01-01'), dt_ini - pd.Timedelta(seconds=1))

        # União dos YTDs anteriores ao y-1 (i.e., y-2, y-3, ...)
        ytds_antigos = set()
        for y_old in range(min(ANOS) - 5, ano - 1):
            ytd_old_fim = (pd.Timestamp(f'{y_old}-{DATA_REF.month:02d}-01') +
                           pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
            ytds_antigos |= _cli_set(vendas, pd.Timestamp(f'{y_old}-01-01'), ytd_old_fim)

        novos = (period - all_pre) - {NAN_TOKEN}
        rec = (period & prev_ytd) - {NAN_TOKEN}
        ret = ((period & ytds_antigos) - prev_ytd) - {NAN_TOKEN}
        # perd inclui NaN
        perd = prev_ytd - period
        out[str(ano)] = {
            'total': len(period),
            'novos': len(novos),
            'retornados': len(ret),
            'recorrentes': len(rec),
            'perdidos': len(perd),
        }
    return out

def gerar_coorte_tri(vendas: pd.DataFrame) -> dict:
    """coorte_tri[ano].qN: período = trimestre, prev = MESMO trimestre ano anterior.
    total EXCLUI NaN. Para trimestres futuros: total=novos=ret=rec=0, mas perd preenchido."""
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            dt_ini, dt_fim = _periodo_tri(ano, q)
            # prev: mesmo Q ano-1, INTEIRO
            prev_dt_ini = pd.Timestamp(f'{ano - 1}-{((q-1)*3+1):02d}-01')
            prev_dt_fim = (pd.Timestamp(f'{ano - 1}-{(q*3):02d}-01') + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
            prev = _cli_set(vendas, prev_dt_ini, prev_dt_fim)
            # Trimestre futuro: period vazio, mas perd vem do prev
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                # 2026.q3 e q4: total=0, novos=0, ret=0, rec=0, perd=??
                # Gold 2026.q3 perd=84 (q2 atual tem 50 cli, q1 95 cli — não bate com 84)
                # Vou olhar regra: se period é VAZIO e dt_ini > DATA_REF, prev = q ANTERIOR (não mesmo Q ano anterior)
                # 2026.q3 prev = 2026.q2? ou seja, fluxo natural
                # Mas precisaria testar.
                # Simplificação: para trimestre futuro, prev = ano anterior MESMO Q, perd = todos
                period = set()
                all_pre = _cli_set(vendas, pd.Timestamp('1970-01-01'), dt_ini - pd.Timedelta(seconds=1))
                out[str(ano)][f'q{q}'] = _coorte_calc(period, prev, all_pre, total_inclui_nan=False)
                continue
            period = _cli_set(vendas, dt_ini, dt_fim)
            all_pre = _cli_set(vendas, pd.Timestamp('1970-01-01'), dt_ini - pd.Timedelta(seconds=1))
            out[str(ano)][f'q{q}'] = _coorte_calc(period, prev, all_pre, total_inclui_nan=False)
    return out

# --------------------------------------------------------------------------------------
# Bloco anual — socios_*  / socios_det_*
# --------------------------------------------------------------------------------------

SOCIO_GLAUCO = 'GLAUCO CICERO BARBOSA'

def _socios_filtros(df: pd.DataFrame) -> dict:
    """Aplica filtros de Thompson 3.4.07 e Glauco 2.3.02; retorna {pro_labore, comissao_socio}."""
    pl = df[(df['SUB-CONTA'] == '3.4.07') &
            (df['FORNECEDOR'].fillna('').str.contains(SOCIO_THOMPSON, case=False))]
    com = df[(df['SUB-CONTA'] == '2.3.02') &
             (df['FORNECEDOR'].fillna('').str.fullmatch(SOCIO_GLAUCO, case=False, na=False))]
    return {
        'pro_labore': _round(float(pl['VLR.PAGO'].fillna(0).sum())),
        'comissao_socio': _round(float(com['VLR.PAGO'].fillna(0).sum())),
    }

def _filtra_pagar_socios(pagar: pd.DataFrame, dt_ini, dt_fim, regime: str) -> pd.DataFrame:
    """Filtra pagar conforme regime: 'venc' (competência) ou 'pagto' (caixa, exige LIQUIDADO=S)."""
    if regime == 'venc':
        return pagar[(pagar['DT.VENC.'] >= dt_ini) & (pagar['DT.VENC.'] <= dt_fim)]
    return pagar[(pagar['DT.PAGTO.'] >= dt_ini) & (pagar['DT.PAGTO.'] <= dt_fim) &
                 (pagar['LIQUIDADO'] == 'S')]

def gerar_socios_full(pagar: pd.DataFrame) -> dict:
    """socios_full: DT.VENC (regime competência) em todos os anos."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        dt_fim = pd.Timestamp(f'{ano}-12-31 23:59:59')
        out[str(ano)] = _socios_filtros(_filtra_pagar_socios(pagar, dt_ini, dt_fim, 'venc'))
    return out

def gerar_socios_ytd(pagar: pd.DataFrame) -> dict:
    """socios_ytd: DT.VENC para anos passados (mês fechado); DT.PAGTO+LIQUIDADO=S para ano corrente."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
            regime = 'pagto'
        else:
            mes = DATA_REF.month
            dt_fim = (pd.Timestamp(f'{ano}-{mes:02d}-01') + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
            regime = 'venc'
        out[str(ano)] = _socios_filtros(_filtra_pagar_socios(pagar, dt_ini, dt_fim, regime))
    return out

def _socios_tri_periodo(pagar_pre: pd.DataFrame, dt_ini, dt_fim) -> dict:
    """socios_tri: filtro por DT.PAGTO. (regime caixa).
    pro_labore = TODOS em 3.4.07; comissao_socio = TODOS em 2.3.02."""
    p = pagar_pre[(pagar_pre['_dt_filtro'] >= dt_ini) & (pagar_pre['_dt_filtro'] <= dt_fim)]
    pl = p[p['_co'] == '3.4.07']
    com = p[p['_co'] == '2.3.02']
    return {
        'pro_labore': _round(float(pl['_vlr_ef'].sum())),
        'comissao_socio': _round(float(com['_vlr_ef'].sum())),
    }

def gerar_socios_tri(pagar_pre: pd.DataFrame) -> dict:
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            dt_ini, dt_fim = _periodo_tri(ano, q)
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                out[str(ano)][f'q{q}'] = {'pro_labore': 0, 'comissao_socio': 0}
                continue
            out[str(ano)][f'q{q}'] = _socios_tri_periodo(pagar_pre, dt_ini, dt_fim)
    return out

# --------------------------------------------------------------------------------------
# Bloco anual — socios_det_full / _ytd / _tri (regra aproximada — pendente verificação)
# --------------------------------------------------------------------------------------

def _socios_det_blacklist(co) -> bool:
    """Sub-contas excluídas em socios_det.total (despesas operacionais não-pessoais)."""
    s = str(co)
    if s.startswith('1.4') or s.startswith('2.2'):
        return True
    if s in {'3.4.01', '3.4.03', '3.4.04', '3.4.05', '3.4.06'}:
        return True
    if s.startswith('4.') or s.startswith('5.'):
        return True
    return False

def _socios_det_calc(df: pd.DataFrame) -> dict:
    """Para um DataFrame já filtrado por período/regime, retorna dict {Thompson, Glauco}.
    pro_labore e comissao são SEMPRE 0 (ficam apenas em total)."""
    df_ok = df[~df['SUB-CONTA'].apply(_socios_det_blacklist)]
    t_total = df_ok[df_ok['FORNECEDOR'].fillna('').str.contains(SOCIO_THOMPSON, case=False)]['VLR.PAGO'].fillna(0).sum()
    g_total = df_ok[df_ok['FORNECEDOR'].fillna('').str.fullmatch(SOCIO_GLAUCO, case=False, na=False)]['VLR.PAGO'].fillna(0).sum()
    return {
        'Thompson': {'pro_labore': 0, 'comissao': 0, 'total': _round(float(t_total))},
        'Glauco': {'pro_labore': 0, 'comissao': 0, 'total': _round(float(g_total))},
    }

def gerar_socios_det_full(pagar: pd.DataFrame) -> dict:
    """socios_det_full[ano].{Thompson|Glauco}.{pro_labore|comissao|total}.
    Regra: DT.VENC (regime competência), exclui blacklist."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        dt_fim = pd.Timestamp(f'{ano}-12-31 23:59:59')
        df = _filtra_pagar_socios(pagar, dt_ini, dt_fim, 'venc')
        out[str(ano)] = _socios_det_calc(df)
    return out

def gerar_socios_det_ytd(pagar: pd.DataFrame) -> dict:
    """socios_det_ytd[ano]: DT.VENC para anos passados; DT.PAGTO+LIQUIDADO=S para ano corrente."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
            regime = 'pagto'
        else:
            mes = DATA_REF.month
            dt_fim = (pd.Timestamp(f'{ano}-{mes:02d}-01') + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
            regime = 'venc'
        df = _filtra_pagar_socios(pagar, dt_ini, dt_fim, regime)
        out[str(ano)] = _socios_det_calc(df)
    return out

def gerar_socios_det_tri(pagar: pd.DataFrame) -> dict:
    """socios_det_tri: DT.PAGTO+LIQUIDADO=S, exclui blacklist."""
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            dt_ini, dt_fim = _periodo_tri(ano, q)
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                out[str(ano)][f'q{q}'] = {
                    'Thompson': {'pro_labore': 0, 'comissao': 0, 'total': 0},
                    'Glauco': {'pro_labore': 0, 'comissao': 0, 'total': 0},
                }
                continue
            df = _filtra_pagar_socios(pagar, dt_ini, dt_fim, 'pagto')
            out[str(ano)][f'q{q}'] = _socios_det_calc(df)
    return out

# --------------------------------------------------------------------------------------
# Bloco anual — inad / inad_tri (regra aproximada — DT.VENC sub-conta 1.x)
# --------------------------------------------------------------------------------------

def _inad_calc(receber: pd.DataFrame, dt_ini, dt_fim) -> dict:
    """Calcula total_faturado, inadimplente, pct para um período.
    total_faturado = VLR.PARCELA por DT.VENC entre [dt_ini, dt_fim] em sub-conta 1.x
    inadimplente = mesmas parcelas com LIQUIDADO=N E DT.VENC <= DATA_REF (já vencidas)
    pct = inad / total * 100"""
    sub = receber[(receber['DT.VENC.'] >= dt_ini) & (receber['DT.VENC.'] <= dt_fim) &
                   (receber['SUB-CONTA'].astype(str).str.startswith('1.'))]
    total_faturado = float(sub['VLR.PARCELA'].sum())
    inad_sub = sub[(sub['LIQUIDADO'] == 'N') & (sub['DT.VENC.'] <= DATA_REF)]
    inad = float(inad_sub['VLR.PARCELA'].sum())
    pct = round(inad / total_faturado * 100, 2) if total_faturado > 0 else 0
    return {
        'total_faturado': _round(total_faturado),
        'inadimplente': _round(inad),
        'pct': pct,
    }

def gerar_inad(receber: pd.DataFrame) -> dict:
    """inad[ano]: total_faturado, inadimplente, pct (anual).
    total_faturado: VLR.PARCELA por DT.VENC ano inteiro (mesmo ano corrente).
    inadimplente: parcelas LIQ=N e DT.VENC <= DATA_REF."""
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        dt_fim = pd.Timestamp(f'{ano}-12-31 23:59:59')  # ano cheio para todos
        out[str(ano)] = _inad_calc(receber, dt_ini, dt_fim)
    return out

def gerar_inad_tri(receber: pd.DataFrame) -> dict:
    """inad_tri[ano].qN: total_faturado, inadimplente, pct (trimestral)."""
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            dt_ini, dt_fim = _periodo_tri(ano, q)
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                out[str(ano)][f'q{q}'] = {'total_faturado': 0, 'inadimplente': 0, 'pct': 0}
                continue
            out[str(ano)][f'q{q}'] = _inad_calc(receber, dt_ini, dt_fim)
    return out

# --------------------------------------------------------------------------------------
# Bloco anual — pmr_pmp / pmr_pmp_full / _ytd / _tri (aproximação)
# --------------------------------------------------------------------------------------

def _pmr_calc(receber: pd.DataFrame, dt_ini, dt_fim) -> tuple:
    """Calcula (pmr_concedido, pmr_efetivo) para um período.
    pmr_concedido = média ponderada (DT.VENC - DT.EMISSÃO) por VLR.PARCELA
    pmr_efetivo = média ponderada (DT.PAGTO - DT.EMISSÃO) por VLR.PAGO
    Filtro: DT.EMISSÃO entre [dt_ini, dt_fim], sub-conta 1.x, LIQUIDADO=S."""
    sub = receber[(receber['DT.EMISSÃO'] >= dt_ini) & (receber['DT.EMISSÃO'] <= dt_fim) &
                   (receber['SUB-CONTA'].astype(str).str.startswith('1.')) &
                   (receber['LIQUIDADO'] == 'S')]
    sub_c = sub[sub['VLR.PARCELA'] > 0]
    if len(sub_c) > 0 and sub_c['VLR.PARCELA'].sum() > 0:
        prazo_c = (sub_c['DT.VENC.'] - sub_c['DT.EMISSÃO']).dt.days
        pmr_c = (prazo_c * sub_c['VLR.PARCELA']).sum() / sub_c['VLR.PARCELA'].sum()
    else:
        pmr_c = 0
    sub_e = sub[(sub['VLR.PAGO'].fillna(0) > 0) & sub['DT.PAGTO.'].notna()]
    if len(sub_e) > 0 and sub_e['VLR.PAGO'].sum() > 0:
        prazo_e = (sub_e['DT.PAGTO.'] - sub_e['DT.EMISSÃO']).dt.days
        pmr_e = (prazo_e * sub_e['VLR.PAGO']).sum() / sub_e['VLR.PAGO'].sum()
    else:
        pmr_e = 0
    return round(pmr_c, 1), round(pmr_e, 1)

def _pmp_calc(pagar: pd.DataFrame, dt_ini, dt_fim) -> tuple:
    """Calcula (pmp_concedido, pmp_efetivo) para um período.
    Filtro: DT.EMISSÃO entre [dt_ini, dt_fim], LIQUIDADO=S, exclui sub-conta 1.x (compras)."""
    sub = pagar[(pagar['DT.EMISSÃO'] >= dt_ini) & (pagar['DT.EMISSÃO'] <= dt_fim) &
                 (pagar['LIQUIDADO'] == 'S')]
    sub_c = sub[(sub['VLR.PARCELA'] > 0) & sub['DT.VENC.'].notna()]
    if len(sub_c) > 0 and sub_c['VLR.PARCELA'].sum() > 0:
        prazo_c = (sub_c['DT.VENC.'] - sub_c['DT.EMISSÃO']).dt.days
        pmp_c = (prazo_c * sub_c['VLR.PARCELA']).sum() / sub_c['VLR.PARCELA'].sum()
    else:
        pmp_c = 0
    sub_e = sub[(sub['VLR.PAGO'].fillna(0) > 0) & sub['DT.PAGTO.'].notna()]
    if len(sub_e) > 0 and sub_e['VLR.PAGO'].sum() > 0:
        prazo_e = (sub_e['DT.PAGTO.'] - sub_e['DT.EMISSÃO']).dt.days
        pmp_e = (prazo_e * sub_e['VLR.PAGO']).sum() / sub_e['VLR.PAGO'].sum()
    else:
        pmp_e = 0
    return round(pmp_c, 1), round(pmp_e, 1)

def gerar_pmr_pmp_full(receber: pd.DataFrame, pagar: pd.DataFrame) -> dict:
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
        else:
            dt_fim = pd.Timestamp(f'{ano}-12-31 23:59:59')
        pmr_c, pmr_e = _pmr_calc(receber, dt_ini, dt_fim)
        pmp_c, pmp_e = _pmp_calc(pagar, dt_ini, dt_fim)
        out[str(ano)] = {
            'pmr_concedido': pmr_c, 'pmr_efetivo': pmr_e,
            'pmp_concedido': pmp_c, 'pmp_efetivo': pmp_e,
        }
    return out

def gerar_pmr_pmp_ytd(receber: pd.DataFrame, pagar: pd.DataFrame) -> dict:
    out = {}
    for ano in ANOS:
        dt_ini = pd.Timestamp(f'{ano}-01-01')
        if ano == DATA_REF.year:
            dt_fim = DATA_REF
        else:
            mes = DATA_REF.month
            dt_fim = (pd.Timestamp(f'{ano}-{mes:02d}-01') + pd.offsets.MonthEnd(0)).replace(hour=23, minute=59, second=59)
        pmr_c, pmr_e = _pmr_calc(receber, dt_ini, dt_fim)
        pmp_c, pmp_e = _pmp_calc(pagar, dt_ini, dt_fim)
        out[str(ano)] = {
            'pmr_concedido': pmr_c, 'pmr_efetivo': pmr_e,
            'pmp_concedido': pmp_c, 'pmp_efetivo': pmp_e,
        }
    return out

def gerar_pmr_pmp_tri(receber: pd.DataFrame, pagar: pd.DataFrame) -> dict:
    out = {}
    for ano in ANOS:
        out[str(ano)] = {}
        for q in (1, 2, 3, 4):
            dt_ini, dt_fim = _periodo_tri(ano, q)
            if ano == DATA_REF.year and dt_ini > DATA_REF:
                out[str(ano)][f'q{q}'] = {
                    'pmr_concedido': 0, 'pmr_efetivo': 0,
                    'pmp_concedido': 0, 'pmp_efetivo': 0,
                }
                continue
            pmr_c, pmr_e = _pmr_calc(receber, dt_ini, dt_fim)
            pmp_c, pmp_e = _pmp_calc(pagar, dt_ini, dt_fim)
            out[str(ano)][f'q{q}'] = {
                'pmr_concedido': pmr_c, 'pmr_efetivo': pmr_e,
                'pmp_concedido': pmp_c, 'pmp_efetivo': pmp_e,
            }
    return out

def gerar_pmr_pmp(receber: pd.DataFrame, pagar: pd.DataFrame) -> dict:
    """pmr_pmp (sem sufixo) — regra exata desconhecida.
    Aproximação: usar mesmos cálculos do _full mas com janela diferente.
    Por enquanto usa a mesma regra do _full como placeholder."""
    return gerar_pmr_pmp_full(receber, pagar)

# --------------------------------------------------------------------------------------
# Bloco semanal — saldos / saldo_total / dev_comprando / rec_ab / rec_qt / pag_ab / pag_qt
# --------------------------------------------------------------------------------------
# Saldos bancários: podem vir de 3 fontes (em ordem de prioridade):
#   1. PDFs de extrato bancário na pasta de xls (formato do ERP) — ler automaticamente
#   2. Arquivo JSON externo via --saldos
#   3. Hardcoded abaixo (fallback)
SALDOS_BANCARIOS = {
    'Itaú': 72053.18,
    'Inter': 5955.71,
    'Stone': 8725.22,
    'Caixa': 1011.46,
}

def _extrair_dados_pdf_extrato(path: Path) -> tuple:
    """Extrai (banco, saldo_final, data_extrato) de um PDF de extrato do ERP.
    Retorna (None, None, None) se não conseguir extrair.
    Formato esperado: 'Extrato de Conta Financeira' do ERP Bigtex."""
    try:
        from pypdf import PdfReader
    except ImportError:
        print('⚠️  pypdf não instalado. Rode: pip install pypdf')
        return None, None, None
    try:
        reader = PdfReader(str(path))
        text = ''
        for page in reader.pages:
            text += page.extract_text() + '\n'
    except Exception as e:
        return None, None, None

    # 1. Banco — padrão "[ N ]  NOME DO BANCO  Conta Financeira:"
    m = re.search(r'\[\s*\d+\s*\]\s+([A-Z][A-Z\s\d]+?)\s+Conta Financeira:', text)
    if not m:
        return None, None, None
    nome_bruto = m.group(1).strip()
    nome_upper = nome_bruto.upper()
    banco = None
    for keyword, label in [
        ('CAIXA', 'Caixa'), ('INTER', 'Inter'),
        ('ITAU', 'Itaú'), ('ITAÚ', 'Itaú'),
        ('STONE', 'Stone'), ('BRADESCO', 'Bradesco'),
        ('SANTANDER', 'Santander'), ('SICOOB', 'Sicoob'),
        ('NUBANK', 'Nubank'), ('NU PAG', 'Nubank'),
        ('BANCO DO BRASIL', 'BB'), ('BB ', 'BB'),
    ]:
        if keyword in nome_upper:
            banco = label
            break
    if not banco:
        banco = nome_bruto.title()  # fallback

    # 2. Saldo final — padrão "<valor>  Sld. Final"
    m = re.search(r'([\d.]+,\d{2})\s+Sld\.\s+Final', text)
    if not m:
        return banco, None, None
    saldo = float(m.group(1).replace('.', '').replace(',', '.'))

    # 3. Data do extrato — preferir "Data: DD/MM até DD/MM/YYYY"
    m = re.search(r'Data:\s*\d{2}/\d{2}\s+até\s+(\d{2})/(\d{2})/(\d{4})', text)
    if m:
        d, mo, y = m.groups()
        data_extrato = f'{y}-{mo}-{d}'
    else:
        # Fallback: data de impressão
        m = re.search(r'Impresso em:\s*(\d{2})/(\d{2})/(\d{4})', text)
        data_extrato = f'{m.group(3)}-{m.group(2)}-{m.group(1)}' if m else None

    return banco, saldo, data_extrato

def carregar_saldos_de_pdfs(pasta: Path, verbose: bool = True) -> tuple:
    """Procura PDFs de extrato bancário na pasta e extrai os saldos.
    Retorna ({banco: saldo, ...}, data_mais_recente_dos_extratos).
    PDFs reconhecidos: qualquer .pdf cuja primeira página tenha 'Extrato de Conta Financeira'."""
    saldos = {}
    datas = []
    pdfs = sorted(pasta.glob('*.pdf'))
    if verbose and pdfs:
        print(f'🔍 Procurando extratos PDF em {pasta}...')
    for pdf in pdfs:
        banco, saldo, data = _extrair_dados_pdf_extrato(pdf)
        if banco and saldo is not None:
            if banco in saldos and verbose:
                print(f'  ⚠️  {pdf.name}: banco "{banco}" já tinha saldo R$ {saldos[banco]:,.2f}, '
                      f'sobrescrevendo com R$ {saldo:,.2f}')
            saldos[banco] = saldo
            if data:
                datas.append(data)
            if verbose:
                print(f'  ✅ {pdf.name} → {banco}: R$ {saldo:,.2f} (em {data or "?"})')
        elif verbose:
            print(f'  ⚠️  {pdf.name}: não parece ser um extrato (ignorado)')
    data_max = max(datas) if datas else None
    return saldos, data_max

def gerar_saldos() -> dict:
    """Saldos bancários por banco."""
    return {b: _round(v) for b, v in SALDOS_BANCARIOS.items()}

def gerar_saldo_total() -> float:
    """Soma de todos os saldos."""
    return _round(sum(SALDOS_BANCARIOS.values()))

def _fmt_data(dt) -> str:
    """Formata pd.Timestamp como 'YYYY-MM-DD' (formato esperado pelo semanal_v2.html)."""
    if pd.isna(dt):
        return ''
    return dt.strftime('%Y-%m-%d')

def gerar_rec_ab(receber: pd.DataFrame) -> list[dict]:
    """Contas a receber EM ABERTO (LIQUIDADO=N).
    Estrutura esperada pelo semanal_v2.html: [{vc: 'YYYY-MM-DD', c: cliente, v: float}]."""
    sub = receber[(receber['LIQUIDADO'] == 'N') & (receber['DT.VENC.'].notna())].copy()
    sub = sub[sub['VLR.PARCELA'] > 0]
    return [
        {
            'vc': _fmt_data(r['DT.VENC.']),
            'c': str(r['CLIENTE']) if pd.notna(r['CLIENTE']) else '(sem cliente)',
            'v': _round(float(r['VLR.PARCELA'] - (r['VLR.PAGO'] or 0))),
        }
        for _, r in sub.iterrows()
    ]

def gerar_rec_qt(receber: pd.DataFrame) -> list[dict]:
    """Contas a receber QUITADAS recentemente (LIQUIDADO=S, DT.PAGTO no ano corrente).
    Inclui mais janela para fluxo de caixa retroativo (últimos 6 meses)."""
    cutoff = DATA_REF - pd.Timedelta(days=180)
    sub = receber[(receber['LIQUIDADO'] == 'S') & (receber['DT.PAGTO.'].notna()) &
                  (receber['DT.PAGTO.'] >= cutoff)].copy()
    sub = sub[sub['VLR.PAGO'].fillna(0) > 0]
    return [
        {
            'vc': _fmt_data(r['DT.PAGTO.']),
            'c': str(r['CLIENTE']) if pd.notna(r['CLIENTE']) else '(sem cliente)',
            'v': _round(float(r['VLR.PAGO'])),
        }
        for _, r in sub.iterrows()
    ]

def gerar_pag_ab(pagar: pd.DataFrame) -> list[dict]:
    """Contas a pagar EM ABERTO (LIQUIDADO=N).
    Estrutura: [{vc: 'YYYY-MM-DD', f: fornecedor, v: float}]."""
    sub = pagar[(pagar['LIQUIDADO'] == 'N') & (pagar['DT.VENC.'].notna())].copy()
    sub = sub[sub['VLR.PARCELA'] > 0]
    return [
        {
            'vc': _fmt_data(r['DT.VENC.']),
            'f': str(r['FORNECEDOR']) if pd.notna(r['FORNECEDOR']) else '(sem fornecedor)',
            'v': _round(float(r['VLR.PARCELA'] - (r['VLR.PAGO'] or 0))),
        }
        for _, r in sub.iterrows()
    ]

def gerar_pag_qt(pagar: pd.DataFrame) -> list[dict]:
    """Contas a pagar QUITADAS recentemente."""
    cutoff = DATA_REF - pd.Timedelta(days=180)
    sub = pagar[(pagar['LIQUIDADO'] == 'S') & (pagar['DT.PAGTO.'].notna()) &
                (pagar['DT.PAGTO.'] >= cutoff)].copy()
    sub = sub[sub['VLR.PAGO'].fillna(0) > 0]
    return [
        {
            'vc': _fmt_data(r['DT.PAGTO.']),
            'f': str(r['FORNECEDOR']) if pd.notna(r['FORNECEDOR']) else '(sem fornecedor)',
            'v': _round(float(r['VLR.PAGO'])),
        }
        for _, r in sub.iterrows()
    ]

def gerar_dev_comprando(receber: pd.DataFrame, vendas: pd.DataFrame) -> list[dict]:
    """Clientes que estão comprando MAS têm dívidas vencidas.
    Para cada cliente nessa situação, retorna {c, dias, div, nov}:
      - c: nome do cliente
      - dias: dias da parcela vencida MAIS antiga (cap em 365 dias)
      - div: total de parcelas vencidas (dívida) ainda em aberto
      - nov: total de parcelas NOVAS (não vencidas, futuro)
    Ignora dívidas com mais de 365 dias (write-offs efetivos).
    Lista ordenada por div desc."""
    # Parcelas vencidas (DT.VENC < DATA_REF) e em aberto, atraso <= 1 ano
    cutoff_velho = DATA_REF - pd.Timedelta(days=365)
    venc = receber[(receber['LIQUIDADO'] == 'N') & (receber['DT.VENC.'].notna()) &
                   (receber['DT.VENC.'] < DATA_REF) &
                   (receber['DT.VENC.'] >= cutoff_velho) &
                   (receber['CLIENTE'].notna())].copy()
    venc = venc[venc['VLR.PARCELA'] > 0]
    if len(venc) == 0:
        return []
    venc['_saldo'] = venc['VLR.PARCELA'] - venc['VLR.PAGO'].fillna(0)
    venc = venc[venc['_saldo'] > 0]

    div_por_cli = venc.groupby('CLIENTE').agg(
        div=('_saldo', 'sum'),
        dt_min=('DT.VENC.', 'min'),
    ).reset_index()
    div_por_cli['dias'] = (DATA_REF - div_por_cli['dt_min']).dt.days

    # Parcelas NOVAS dos mesmos clientes (DT.VENC >= DATA_REF, em aberto)
    novas = receber[(receber['LIQUIDADO'] == 'N') & (receber['DT.VENC.'].notna()) &
                    (receber['DT.VENC.'] >= DATA_REF) &
                    (receber['CLIENTE'].notna())].copy()
    novas = novas[novas['VLR.PARCELA'] > 0]
    novas['_saldo'] = novas['VLR.PARCELA'] - novas['VLR.PAGO'].fillna(0)
    nov_por_cli = novas.groupby('CLIENTE')['_saldo'].sum().to_dict()

    # Vendas recentes (últimos 60 dias) — indicador de "continua comprando"
    vendas_recentes = set()
    if len(vendas) > 0 and 'Dt.Emissão' in vendas.columns:
        v = vendas.copy()
        v['Dt.Emissão'] = pd.to_datetime(v['Dt.Emissão'], errors='coerce')
        cutoff_v = DATA_REF - pd.Timedelta(days=60)
        recent = v[v['Dt.Emissão'] >= cutoff_v]
        vendas_recentes = set(recent['Cliente'].dropna().unique())

    out = []
    for _, r in div_por_cli.iterrows():
        cli = r['CLIENTE']
        nov = nov_por_cli.get(cli, 0)
        if nov > 0 or cli in vendas_recentes:
            out.append({
                'c': str(cli),
                'dias': int(r['dias']),
                'div': _round(float(r['div'])),
                'nov': _round(float(nov)),
            })
    out.sort(key=lambda x: -x['div'])
    return out

# --------------------------------------------------------------------------------------
# Geração final do payload.js completo
# --------------------------------------------------------------------------------------

def gerar_payload_completo(vendas: pd.DataFrame, pagar: pd.DataFrame,
                            pagar_pre: pd.DataFrame, receber: pd.DataFrame) -> dict:
    """Gera o dict completo das 43 chaves do payload.js.
    Chaves do bloco semanal (saldos/saldo_total/dev_comprando/rec_*/pag_*) ficam
    como placeholder vazio até receber extratos bancários."""
    return {
        # Simples (5)
        'data_ref': gerar_data_ref(),
        'hoje': gerar_hoje(),
        'ytd_mes': gerar_ytd_mes(),
        'anos': gerar_anos(),
        'tri_completos': gerar_tri_completos(),
        # Anual KPIs (3)
        'ano_full': gerar_ano_full(vendas, pagar_pre),
        'ano_ytd': gerar_ano_ytd(vendas, pagar_pre),
        'ano_tri': gerar_ano_tri(vendas, pagar_pre),
        # Despesas categorizadas (5)
        'desp_cats_full': gerar_desp_cats_full(pagar_pre),
        'desp_cats_ytd': gerar_desp_cats_ytd(pagar_pre),
        'desp_cats_ano': gerar_desp_cats_ano(pagar_pre),
        'desp_cats_tri': gerar_desp_cats_tri(pagar_pre),
        'desp_cats_full_tri': gerar_desp_cats_full_tri(pagar_pre),
        # Mensal reduzido (1)
        'ym_data': gerar_ym_data(vendas, pagar_pre),
        # Top clientes/produtos (4)
        'top_cli': gerar_top_cli(vendas),
        'top_prod': gerar_top_prod(vendas),
        'top_cli_tri': gerar_top_cli_tri(vendas),
        'top_prod_tri': gerar_top_prod_tri(vendas),
        # Coorte (3)
        'coorte_full': gerar_coorte_full(vendas),
        'coorte_ytd': gerar_coorte_ytd(vendas),
        'coorte_tri': gerar_coorte_tri(vendas),
        # Sócios (3)
        'socios_full': gerar_socios_full(pagar),
        'socios_ytd': gerar_socios_ytd(pagar),
        'socios_tri': gerar_socios_tri(pagar_pre),
        # Sócios detalhado (3) — aproximação
        'socios_det_full': gerar_socios_det_full(pagar),
        'socios_det_ytd': gerar_socios_det_ytd(pagar),
        'socios_det_tri': gerar_socios_det_tri(pagar),
        # Inadimplência (2) — aproximação
        'inad': gerar_inad(receber),
        'inad_tri': gerar_inad_tri(receber),
        # Prazos (4) — aproximação
        'pmr_pmp': gerar_pmr_pmp(receber, pagar),
        'pmr_pmp_full': gerar_pmr_pmp_full(receber, pagar),
        'pmr_pmp_ytd': gerar_pmr_pmp_ytd(receber, pagar),
        'pmr_pmp_tri': gerar_pmr_pmp_tri(receber, pagar),
        # Mensal (3)
        'meses_disp': gerar_meses_disp(),
        'meses': gerar_meses(vendas, pagar),
        'metas': gerar_metas(),
        # Semanal (7) — populado a partir de receber/pagar + saldos dos extratos PDF
        'saldos': gerar_saldos(),
        'saldo_total': gerar_saldo_total(),
        'dev_comprando': gerar_dev_comprando(receber, vendas),
        'rec_ab': gerar_rec_ab(receber),
        'rec_qt': gerar_rec_qt(receber),
        'pag_ab': gerar_pag_ab(pagar),
        'pag_qt': gerar_pag_qt(pagar),
    }

# Chaves usadas pelo painel /financeiro/ (subset do payload completo)
CHAVES_FINANCEIRO = [
    'data_ref', 'hoje',
    'saldos', 'saldo_total', 'dev_comprando',
    'rec_ab', 'rec_qt', 'pag_ab', 'pag_qt',
]


def escrever_payload_financeiro_js(payload_completo: dict, path: str) -> None:
    """Escreve um payload.js REDUZIDO contendo só as chaves do painel financeiro.
    Mesmo formato do payload completo (`var D = {...};`), só com 9 chaves.
    """
    payload_reduzido = {k: payload_completo[k] for k in CHAVES_FINANCEIRO if k in payload_completo}
    body = json.dumps(payload_reduzido, ensure_ascii=False, separators=(',', ':'))
    cabecalho = (
        '// ============================================================================\n'
        '// payload.js — Subset do painel financeiro (Bigtex)\n'
        '// ============================================================================\n'
        f'// Gerado por gerar_payload.py em {DATA_REF.strftime("%d/%m/%Y")}\n'
        '// Contém apenas as 9 chaves usadas pelo painel semanal:\n'
        '//   data_ref, hoje, saldos, saldo_total, dev_comprando,\n'
        '//   rec_ab, rec_qt, pag_ab, pag_qt\n'
        '// ============================================================================\n'
        '\n'
    )
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(cabecalho)
        f.write('var D = ')
        f.write(body)
        f.write(';\n')


def escrever_payload_js(payload: dict, path: str) -> None:
    """Escreve payload.js no MESMO FORMATO do golden:
       - Cabeçalho de comentários
       - `var D = {...};` em uma linha só (sem indent), exatamente como o painel original espera

    Compatibilidade total com index.html, mensal.html e anual.html que fazem
    `<script src="payload.js"></script>` e depois leem `D.ano_full`, `D.meses` etc.
    """
    # JSON compacto (separadores sem espaços) — formato do golden
    body = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
    cabecalho = (
        '// ============================================================================\n'
        '// payload.js — Base de dados consolidada do Painel Bigtex\n'
        '// ============================================================================\n'
        f'// Gerado por gerar_payload.py em {DATA_REF.strftime("%d/%m/%Y")}\n'
        '// Fonte: VENDAS_ANALITICO + CONTAS_A_PAGAR + CONTAS_A_RECEBER (xls 2022-2026)\n'
        '//\n'
        '// Estrutura D (43 chaves):\n'
        '//   - data_ref, hoje, ytd_mes, anos, tri_completos\n'
        '//   - ano_full / ano_ytd / ano_tri (KPIs anuais, YTD, trimestrais)\n'
        '//   - desp_cats_full/_ytd/_ano/_tri/_full_tri (despesas categorizadas)\n'
        '//   - top_cli, top_prod, top_cli_tri, top_prod_tri\n'
        '//   - coorte_full / coorte_ytd / coorte_tri\n'
        '//   - socios_full/_ytd/_tri (pro_labore + comissão)\n'
        '//   - socios_det_full/_ytd/_tri (Thompson + Glauco detalhado)\n'
        '//   - inad / inad_tri (inadimplência)\n'
        '//   - pmr_pmp / _full / _ytd / _tri (prazos médios)\n'
        '//   - ym_data, meses, meses_disp, metas\n'
        '//   - saldos, saldo_total, dev_comprando, rec_*/pag_* (semanal)\n'
        '// ============================================================================\n'
        '\n'
    )
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(cabecalho)
        f.write('var D = ')
        f.write(body)
        f.write(';\n')

# --------------------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description='Gerar/validar payload.js do painel Bigtex',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''Exemplos de uso:
  Gerar payload com data atual (default 2026-04-19):
    python gerar_payload.py --gerar
  Gerar payload com data nova:
    python gerar_payload.py --gerar --data 2026-05-01
  Validar contra golden (payload.js de referência):
    python gerar_payload.py --validar --bloco tudo''')
    parser.add_argument('--validar', action='store_true', help='Validar contra golden')
    parser.add_argument('--gerar', action='store_true', help='Gerar payload.js completo')
    parser.add_argument('--saida', default='gestor/payload.js',
                        help='Caminho do payload.js completo (43 chaves). Default: gestor/payload.js')
    parser.add_argument('--saida-financeiro', default='financeiro/payload.js',
                        help='Caminho do payload.js reduzido pro setor financeiro (9 chaves). '
                             'Use --saida-financeiro "" para não gerar a versão reduzida. '
                             'Default: financeiro/payload.js')
    parser.add_argument('--bloco', default='mensal', choices=['mensal','anual','semanal','tudo'])
    parser.add_argument('--saldos', default=None,
                        help='Caminho para arquivo JSON com saldos bancários. '
                             'Default: usa SALDOS_BANCARIOS hardcoded no script. '
                             'Formato: {"Itaú": 72053.18, "Inter": 5955.71, ...}')
    parser.add_argument('--xls-dir', default='/mnt/project',
                        help='Pasta onde estão os xls de entrada. Default: /mnt/project. '
                             'Use --xls-dir "C:/caminho/para/seus/xls" no Windows.')
    parser.add_argument('--data', default=None,
                        help='Data de referência (formato YYYY-MM-DD). Default: hoje. '
                             'Use isso para atualizar sem editar o código.')
    args = parser.parse_args()

    # Atualizar data de referência se passado via CLI
    global DATA_REF, HOJE, DATA_REF_BR, ANOS, PROJECT_DIR, SALDOS_BANCARIOS
    if args.xls_dir:
        PROJECT_DIR = Path(args.xls_dir)
        if not PROJECT_DIR.exists():
            print(f'❌ Erro: pasta "{args.xls_dir}" não existe.')
            return
        print(f'📂 Pasta de xls: {PROJECT_DIR}')

    # Tentar carregar saldos de PDFs de extrato (prioridade alta)
    # Sempre olha em PROJECT_DIR — usuário só precisa colocar os PDFs lá
    pdf_saldos, pdf_data = carregar_saldos_de_pdfs(PROJECT_DIR, verbose=True)
    if pdf_saldos:
        SALDOS_BANCARIOS = pdf_saldos
        print(f'💰 Saldos extraídos de {len(pdf_saldos)} PDF(s) de extrato:')
        for b, v in sorted(pdf_saldos.items()):
            print(f'    {b}: R$ {v:,.2f}')
        if pdf_data and not args.data:
            # Usar a data dos PDFs como data de referência (se usuário não passou --data)
            try:
                DATA_REF = pd.Timestamp(pdf_data)
                HOJE = DATA_REF.strftime('%Y-%m-%d')
                DATA_REF_BR = DATA_REF.strftime('%d/%m/%Y')
                ano_max = max(ANOS + [DATA_REF.year])
                ANOS = list(range(min(ANOS), ano_max + 1))
                print(f'📅 Usando data dos extratos PDF: {DATA_REF_BR}')
            except Exception:
                pass

    # JSON de saldos sobrescreve PDFs (prioridade ainda maior, se passado)
    if args.saldos:
        try:
            with open(args.saldos, 'r', encoding='utf-8') as f:
                novos_saldos = json.load(f)
            if not isinstance(novos_saldos, dict):
                raise ValueError('JSON deve ser um objeto/dict, não array.')
            SALDOS_BANCARIOS = {k: float(v) for k, v in novos_saldos.items()}
            print(f'💰 Saldos sobrescritos via JSON {args.saldos}: {SALDOS_BANCARIOS}')
        except Exception as e:
            print(f'❌ Erro ao ler arquivo de saldos "{args.saldos}": {e}')
            return

    # --data via CLI sobrescreve tudo
    if args.data:
        try:
            dt = pd.Timestamp(args.data)
        except Exception as e:
            print(f'❌ Erro: data inválida "{args.data}". Use formato YYYY-MM-DD (ex: 2026-05-01).')
            return
        DATA_REF = dt
        HOJE = dt.strftime('%Y-%m-%d')
        DATA_REF_BR = dt.strftime('%d/%m/%Y')
        ano_max = max(ANOS + [dt.year])
        ANOS = list(range(min(ANOS), ano_max + 1))
        print(f'📅 Usando data de referência (--data): {DATA_REF_BR} (anos: {ANOS})')

    print()

    print(f'data_ref = {DATA_REF.date()} | bloco = {args.bloco}')
    print()

    golden = golden_keys() if args.validar else {}

    if args.gerar:
        print('=== Gerando payload.js completo ===')
        print('Carregando xls (todos os anos)...')
        vendas_all = carregar_vendas(ANOS)
        pagar_all = carregar_pagar(ANOS)
        receber_all = carregar_receber(ANOS)
        pagar_pre_all = _preparar_pagar(pagar_all)
        print(f'  vendas: {len(vendas_all)} | pagar: {len(pagar_all)} | receber: {len(receber_all)}')
        print('Gerando payload completo (43 chaves)...')
        payload = gerar_payload_completo(vendas_all, pagar_all, pagar_pre_all, receber_all)
        print(f'  Chaves geradas: {len(payload)}')
        escrever_payload_js(payload, args.saida)
        print(f'  ✓ Completo (43 chaves): {args.saida} ({os.path.getsize(args.saida):,} bytes)')

        if args.saida_financeiro:
            escrever_payload_financeiro_js(payload, args.saida_financeiro)
            print(f'  ✓ Financeiro (9 chaves): {args.saida_financeiro} ({os.path.getsize(args.saida_financeiro):,} bytes)')
        return

    if args.bloco in ('mensal', 'tudo'):
        print('=== Bloco mensal ===')
        print('Carregando xls...')
        vendas = carregar_vendas([2025, 2026])
        pagar = carregar_pagar(ANOS)  # TODOS os anos: jan/25 pode ter registros do xls 2024
        print(f'  vendas: {len(vendas)} linhas | pagar: {len(pagar)} linhas')
        print('Gerando meses...')
        meses = gerar_meses(vendas, pagar)
        print(f'  meses gerados: {len(meses)}')

        if args.validar:
            validar_chave('metas', gerar_metas(), golden)
            validar_chave('meses_disp', gerar_meses_disp(), golden)
            validar_chave('meses', meses, golden)

            # Diagnóstico mês-a-mês para identificar onde diverge
            print('\n--- Diagnóstico por ym ---')
            golden_meses_str = golden_value_json('meses')
            golden_meses = json.loads(golden_meses_str)
            for ym in gerar_meses_disp():
                meu_sha = sha256_16(meses[ym])
                gold_sha = sha256_16(golden_meses[ym])
                ok = '✅' if meu_sha == gold_sha else '❌'
                print(f'  {ok} {ym}: meu={meu_sha} gold={gold_sha}')

    if args.bloco in ('anual', 'tudo'):
        print('\n=== Bloco anual ===')
        print('Carregando xls (todos os anos)...')
        vendas_all = carregar_vendas(ANOS)
        pagar_all = carregar_pagar(ANOS)
        print(f'  vendas: {len(vendas_all)} linhas | pagar: {len(pagar_all)} linhas')
        pagar_pre_all = _preparar_pagar(pagar_all)
        receber_all = carregar_receber(ANOS)
        print(f'  receber: {len(receber_all)} linhas')
        print('Gerando chaves anuais simples...')

        if args.validar:
            validar_chave('data_ref', gerar_data_ref(), golden)
            validar_chave('hoje', gerar_hoje(), golden)
            validar_chave('ytd_mes', gerar_ytd_mes(), golden)
            validar_chave('anos', gerar_anos(), golden)
            validar_chave('tri_completos', gerar_tri_completos(), golden)
            ano_full = gerar_ano_full(vendas_all, pagar_pre_all)
            validar_chave('ano_full', ano_full, golden)
            ano_ytd = gerar_ano_ytd(vendas_all, pagar_pre_all)
            validar_chave('ano_ytd', ano_ytd, golden)
            ano_tri = gerar_ano_tri(vendas_all, pagar_pre_all)
            validar_chave('ano_tri', ano_tri, golden)
            validar_chave('desp_cats_full', gerar_desp_cats_full(pagar_pre_all), golden)
            validar_chave('desp_cats_ytd', gerar_desp_cats_ytd(pagar_pre_all), golden)
            validar_chave('desp_cats_ano', gerar_desp_cats_ano(pagar_pre_all), golden)
            validar_chave('desp_cats_tri', gerar_desp_cats_tri(pagar_pre_all), golden)
            validar_chave('desp_cats_full_tri', gerar_desp_cats_full_tri(pagar_pre_all), golden)
            validar_chave('ym_data', gerar_ym_data(vendas_all, pagar_pre_all), golden)
            validar_chave('top_cli', gerar_top_cli(vendas_all), golden)
            validar_chave('top_prod', gerar_top_prod(vendas_all), golden)
            validar_chave('top_cli_tri', gerar_top_cli_tri(vendas_all), golden)
            validar_chave('top_prod_tri', gerar_top_prod_tri(vendas_all), golden)
            validar_chave('coorte_full', gerar_coorte_full(vendas_all), golden)
            validar_chave('coorte_ytd', gerar_coorte_ytd(vendas_all), golden)
            validar_chave('coorte_tri', gerar_coorte_tri(vendas_all), golden)
            validar_chave('socios_full', gerar_socios_full(pagar_all), golden)
            validar_chave('socios_ytd', gerar_socios_ytd(pagar_all), golden)
            validar_chave('socios_tri', gerar_socios_tri(pagar_pre_all), golden)
            validar_chave('socios_det_full', gerar_socios_det_full(pagar_all), golden)
            validar_chave('socios_det_ytd', gerar_socios_det_ytd(pagar_all), golden)
            validar_chave('socios_det_tri', gerar_socios_det_tri(pagar_all), golden)
            validar_chave('inad', gerar_inad(receber_all), golden)
            validar_chave('inad_tri', gerar_inad_tri(receber_all), golden)
            validar_chave('pmr_pmp', gerar_pmr_pmp(receber_all, pagar_all), golden)
            validar_chave('pmr_pmp_full', gerar_pmr_pmp_full(receber_all, pagar_all), golden)
            validar_chave('pmr_pmp_ytd', gerar_pmr_pmp_ytd(receber_all, pagar_all), golden)
            validar_chave('pmr_pmp_tri', gerar_pmr_pmp_tri(receber_all, pagar_all), golden)

if __name__ == '__main__':
    main()
