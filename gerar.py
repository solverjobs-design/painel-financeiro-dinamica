#!/usr/bin/env python3
"""Lê a planilha de fluxo de caixa no SharePoint, trata os lançamentos e gera site/index.html.

Abas usadas: fluxo de caixa (lançamentos), créditos (contratos de empréstimo) e de_para
(categoria de cada subcategoria). Os nomes ficam em config.json.

O painel sai criptografado com a senha do secret PAINEL_SENHA. Se qualquer etapa falhar,
o script termina com erro antes de gravar o site, e o painel publicado continua o anterior.

Uso na automação:   python gerar.py
Teste local:        python gerar.py --xlsx copia.xlsx --hoje 2026-10-29
"""

import argparse
import base64
import collections
import datetime as dt
import html
import http.cookiejar
import io
import json
import os
import re
import sys
import traceback
import unicodedata
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
BRT = dt.timezone(dt.timedelta(hours=-3))  # Brasília (sem horário de verão desde 2019)
EPOCA = dt.date(1899, 12, 30)              # dia zero das datas do Excel
ITERACOES = 600_000
SAL = b'painel-dinamica-v1'
NATUREZAS = {
    'OPER': 'Operacional',
    'EMP': 'Empréstimos e financiamentos',
    'MUT': 'Mútuo entre empresas do grupo',
    'SOC': 'Sócios (pró-labore, retiradas, lucros, mútuos)',
    'APL': 'Aplicações financeiras',
    'OUT': 'Outras não operacionais',
    'INT': 'Movimentações internas',
    'NC': 'Não classificado',
}
NC_CAT = ('Não classificado', 'NC', '99')
FATURA = 'Cartão de crédito (fatura)'
CUSTOS = 'Custos de contratação (IOF, tarifas, seguro)'
CAMPOS = ['m', 'w', 'd', 'st', 'mv', 'emp', 'nat', 'cat', 'sub', 'cp', 'conta', 'meio', 'card', 'pf', 'fv',
          'tax', 'taxp', 'doc', 'loan', 'lender', 'lkind', 'desc', 'v', 'mdt']
COMDIM = ['emp', 'nat', 'cat', 'sub', 'cp', 'conta', 'meio', 'card', 'pf', 'fv', 'tax', 'taxp', 'doc',
          'loan', 'lender', 'lkind', 'desc']
# colunas da aba de fluxo (cabeçalho comparado sem acento, espaço ou caixa)
COL_FLUXO = {
    'fc': 'TIPO FC', 'mov': 'TIPO MOV', 'fv': 'FIXA OU VAR', 'empresa': 'EMPRESA', 'data': 'DATA',
    'semana': 'SEMANA', 'tipodoc': 'TIPO DOC', 'lanc': 'LANÇAMENTO', 'valor': 'VALOR', 'pf': 'PJ /PF',
    'cat': 'CATEGORIA', 'sub': 'SUBCATEGORIA', 'banco': 'BANCO', 'desc': 'DESCRIÇÃO', 'info': 'INFO',
}
OBRIG_FLUXO = ('fc', 'mov', 'empresa', 'data', 'lanc', 'valor', 'cat', 'sub', 'banco', 'desc')
COL_CRED = {
    'empresa': 'EMPRESA', 'liberacao': 'DATA LIBERAÇÃO', 'contrato': 'CONTRATO', 'descricao': 'DESCRIÇÃO',
    'nparc': 'Nº PARCELAS', 'npagas': 'Nº PARCELAS PAGAS', 'liberado': 'VALOR LIBERADO', 'parcela': 'VALOR DA PARCELA',
    'saldo': 'SALDO DEVEDOR', 'banco': 'BANCO', 'venc': 'DATA VENC.', 'juros': 'JUROS AO MÊS', 'indexador': 'CORREÇÃO MONETÁRIA MENSAL',
    'situacao': 'SITUAÇÃO', 'info': 'INFO',
}
OBRIG_CRED = ('empresa', 'contrato', 'descricao', 'parcela', 'banco')
MN = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']


class ErroPainel(Exception):
    """Erro esperado. A mensagem vai para o log público, então nunca leva dados da planilha."""


# ---------------------------------------------------------------- utilitários

def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFKD', str(s)) if not unicodedata.combining(c))


def norm(s):
    """Maiúsculas, sem acento, espaços simples."""
    return ' '.join(sem_acento(s or '').upper().split())


def chave(s):
    """Só letras e números, para comparar cabeçalhos e nomes."""
    return re.sub(r'[^A-Z0-9]', '', norm(s))


def limpo(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = ' '.join(str(v).split())
    return '' if s in ('—', '-', '–', '#N/A', '#N/D', '#VALUE!', '#VALOR!', '#REF!', 'None') else s


def data(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return EPOCA + dt.timedelta(days=int(v))
    s = limpo(v)
    for f in ('%d/%m/%Y', '%Y-%m-%d', '%d/%m/%y', '%Y-%m-%d %H:%M:%S'):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def numero(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = limpo(v).replace('R$', '').replace(' ', '')
    if not s:
        return None
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def brl(v):
    s = f'{abs(v):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return ('−' if v < -0.004 else '') + 'R$ ' + s


def dma(d):
    return d.strftime('%d/%m/%y')


def contem(texto, termo):
    """Termo (já normalizado) aparece como palavra(s) inteira(s) no texto normalizado."""
    return re.search(r'(?<![A-Z0-9])' + re.escape(termo) + r'(?![A-Z0-9])', texto) is not None


def codigo(s):
    """'3.3.01 Manutenção' -> ['3','3','01']; sem código -> []."""
    m = re.match(r'\s*(\d+(?:\.\d+)*)\.?(?:\s|$)', s or '')
    return m.group(1).split('.') if m else []


def sem_codigo(s):
    return re.sub(r'^\s*\d+(?:\.\d+)*\.?\s*', '', s or '').strip(' -–')


CONECTIVOS = {'DE', 'DA', 'DO', 'DAS', 'DOS', 'E', 'EM', 'NA', 'NO', 'PARA', 'COM', 'A', 'O'}
SIGLAS = {'SA', 'S/A', 'LTDA', 'ME', 'EPP', 'EIRELI', 'MEI', 'PJ', 'PF', 'CEF', 'PIX', 'TED', 'ISS', 'IOF', 'FGTS',
          'INSS', 'IPTU', 'IPVA', 'ICMS', 'IRPJ', 'CSLL', 'IRRF', 'CNPJ', 'CPF', 'TV', 'XP', 'MG', 'BH', 'SP', 'RJ', 'CEMIG'}
VOGAIS = set('AEIOUY')


def titulo(s):
    """Nome em caixa mista, mantendo siglas: 'VERO SA' -> 'Vero SA'."""
    out = []
    for i, p in enumerate(s.split()):
        b = sem_acento(p).upper().strip('.,;:()')
        if b in ('AS',) and i > 0 and i == len(s.split()) - 1:
            out.append('SA')
        elif b in SIGLAS or (len(b) <= 3 and b.isalpha() and not (set(b) & VOGAIS)):
            out.append(p.upper())
        elif i > 0 and b in CONECTIVOS:
            out.append(p.lower())
        elif any(c.isdigit() for c in p):
            out.append(p[:1].upper() + p[1:].lower())
        else:
            out.append(p[:1].upper() + p[1:].lower())
    return ' '.join(out)


# ---------------------------------------------------------------- leitura da planilha

def baixar(url):
    """Baixa o .xlsx de um link de compartilhamento do SharePoint/OneDrive ("qualquer pessoa com o link")."""
    if not url:
        raise ErroPainel('o secret PLANILHA_URL não está cadastrado no GitHub')
    partes = urllib.parse.urlsplit(url)
    q = urllib.parse.parse_qs(partes.query)
    q['download'] = ['1']
    alvo = urllib.parse.urlunsplit(partes._replace(query=urllib.parse.urlencode(q, doseq=True)))
    # o SharePoint grava um cookie no primeiro acesso e redireciona para o arquivo; sem o cookie ele devolve a tela de login
    abridor = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    abridor.addheaders = [('User-Agent', 'Mozilla/5.0 (painel-automatico)')]
    try:
        with abridor.open(alvo, timeout=120) as r:
            conteudo = r.read()
    except Exception as e:  # noqa: BLE001
        raise ErroPainel(f'não consegui baixar a planilha ({type(e).__name__}); confira o link e o compartilhamento') from None
    if not conteudo.startswith(b'PK'):
        raise ErroPainel('o link da planilha não devolveu um arquivo Excel (o compartilhamento "qualquer pessoa com o link" pode ter sido desligado)')
    return conteudo


def abrir_planilha(conteudo):
    import openpyxl
    import warnings
    warnings.filterwarnings('ignore', module='openpyxl')
    try:
        return openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    except (zipfile.BadZipFile, KeyError, ValueError):
        raise ErroPainel('o arquivo baixado não é uma planilha Excel válida') from None


def aba(wb, nome):
    alvo = chave(nome)
    for ws in wb.worksheets:
        if chave(ws.title) == alvo:
            return ws
    raise ErroPainel(f'aba "{nome}" não encontrada na planilha')


def ler_tabela(ws, colunas, obrig, nome_aba):
    linhas = ws.iter_rows(values_only=True)
    cab = None
    for _ in range(10):  # o cabeçalho pode não estar na linha 1
        r = next(linhas, None)
        if r is None:
            break
        ks = [chave(c) for c in r]
        if sum(1 for c in colunas.values() if chave(c) in ks) >= len(obrig):
            cab = ks
            break
    if cab is None:
        raise ErroPainel(f'aba "{nome_aba}": cabeçalho não encontrado')
    pos = {}
    for k, c in colunas.items():
        if chave(c) in cab:
            pos[k] = cab.index(chave(c))
        elif k in obrig:
            raise ErroPainel(f'a coluna "{c}" sumiu da aba "{nome_aba}"')
    out = []
    for r in linhas:
        if r is None or all(v is None for v in r):
            continue
        out.append({k: (r[i] if i < len(r) else None) for k, i in pos.items()})
    return out


def ler(conteudo, cfg):
    wb = abrir_planilha(conteudo)
    A = cfg['abas']
    fluxo = ler_tabela(aba(wb, A['fluxo']), COL_FLUXO, OBRIG_FLUXO, A['fluxo'])
    cred = ler_tabela(aba(wb, A['creditos']), COL_CRED, OBRIG_CRED, A['creditos'])
    try:
        dp = aba(wb, A['de_para'])
        de_para = {}
        for r in dp.iter_rows(min_row=2, values_only=True):
            if r and r[0] and len(r) > 1 and r[1]:
                de_para.setdefault(chave(r[0]), limpo(r[1]))
    except ErroPainel:
        de_para = {}  # aba opcional
    return fluxo, cred, de_para


# ---------------------------------------------------------------- classificação

class Regras:
    def __init__(self, cfg, de_para):
        self.cfg = cfg
        self.de_para = de_para
        self.plano = [(p.split('.'), c, n, cc) for p, c, n, cc in cfg['plano_de_contas']]
        self.texto = [(norm(t), c, n, cc) for t, c, n, cc in cfg['regras_por_texto']]
        self.catcode = {NC_CAT[0]: NC_CAT[2]}
        for _, c, _, cc in cfg['plano_de_contas'] + cfg['regras_por_texto']:
            self.catcode.setdefault(c, cc)
        ms = cfg['mutuo_com_socios']
        self.catcode.setdefault(ms[0], ms[2])
        self.renomear = {chave(k): v for k, v in cfg.get('renomear_subcategoria', {}).items()}
        self.fixos = {chave(x) for x in cfg['custos_fixos']}
        self.impostos = [(t, [norm(k) for k in ks]) for t, ks in cfg['impostos']]
        self.bancos = {norm(k): v for k, v in cfg['bancos'].items()}
        self.cartoes = [(norm(k), v) for k, v in cfg['cartao_pelo_nome']]
        self.credores = [(norm(k), v) for k, v in cfg['credores']]
        self.apelidos = [(v, [norm(k) for k in ks]) for v, ks in cfg['apelidos']]
        self.grupo = [norm(x) for x in cfg['nomes_do_grupo']]

    def pelo_codigo(self, cod):
        melhor = None
        for p, c, n, cc in self.plano:
            if cod[:len(p)] == p and (melhor is None or len(p) > melhor[0]):
                melhor = (len(p), c, n)
        return melhor

    def categoria(self, cat, sub):
        """Devolve (categoria gerencial, sigla da natureza) pelo plano de contas da planilha."""
        nc, ns = norm(cat), norm(sub)
        for t, c, n, _ in self.texto:
            if contem(nc, t) or contem(ns, t):
                return c, n
        cc, cs = codigo(cat), codigo(sub)
        if not cc and sub and chave(sub) in self.de_para:      # categoria vazia: usa a aba de_para
            cc = codigo(self.de_para[chave(sub)])
        if len(cs) < 2:                                         # subcategoria sem código válido
            cs = []
        mc = self.pelo_codigo(cc) if cc else None
        ms = self.pelo_codigo(cs) if cs else None
        if ms and (not cc or cs[:len(cc)] == cc) and (not mc or ms[0] > mc[0]):
            return ms[1], ms[2]
        if mc:
            return mc[1], mc[2]
        if ms and not cc:
            return ms[1], ms[2]
        return NC_CAT[0], NC_CAT[1]

    def subcategoria(self, sub, cat_ger):
        s = sem_codigo(limpo(sub))
        if not s or s == '?' or cat_ger == NC_CAT[0]:
            return 'Não classificado'
        s = self.renomear.get(chave(s), s)
        if s.isupper() and len(s) > 4:
            s = s[:1] + s[1:].lower()
        return s

    def imposto(self, sub, lanc, mv):
        if mv != 1:
            return '', ''
        ns, nl = norm(sub), norm(lanc)
        for tipo, ks in self.impostos:
            if any(contem(ns, k) for k in ks) or any(nl.startswith(k) for k in ks):
                return tipo, ('Parcelamento' if 'PARCEL' in ns or 'PARCEL' in nl else 'Corrente')
        return '', ''

    def conta(self, banco, previsto):
        b = norm(banco)
        if b in self.bancos:
            return tuple(self.bancos[b])
        if not b:
            return ('A definir (previsto)' if previsto else 'Não informado', 'Conta corrente', '')
        if 'CART' in b or 'CREDITO' in b:
            return ('Cartão ' + titulo(banco), 'Cartão de crédito', self.cartao(banco) or 'Outro')
        return ('Conta ' + titulo(banco), 'Conta corrente', '')

    def cartao(self, texto):
        t = norm(texto)
        for k, v in self.cartoes:
            if contem(t, k):
                return v
        return ''

    def credor(self, *textos):
        t = ' '.join(norm(x) for x in textos)
        nome = ''
        for k, v in self.credores:
            if contem(t, k):
                nome = v
                break
        return nome

    def apelido(self, texto):
        t = norm(texto) + ' '
        for v, ks in self.apelidos:
            if any(t.startswith(k) or contem(t, k.strip()) for k in ks):
                return v
        return None

    def do_grupo(self, nome):
        n = norm(nome)
        return any(contem(n, g) for g in self.grupo)


PARCELA_RE = [
    re.compile(r'\s*-?\s*\(?\d{1,3}\s*/\s*\d{1,3}\)?\s*$'),   # " - 05/10", " 01/04"
    re.compile(r'\s*\d{2}d\d{2}\s*$', re.I),                   # "Auto05d12"
]
SUFIXOS = {'SA', 'S A', 'AS', 'LTDA', 'ME', 'EPP', 'EIRELI', 'MEI', 'S/A'}


def nome_base(s):
    for rx in PARCELA_RE:
        s = rx.sub('', s)
    return s.strip(' -–')


def chave_nome(s):
    n = re.sub(r'[^A-Z0-9 ]', ' ', norm(s)).split()
    while n and n[-1] in SUFIXOS:
        n.pop()
    while len(n) > 1 and ' '.join(n[-2:]) in SUFIXOS:
        n = n[:-2]
    return ' '.join(n)


# ---------------------------------------------------------------- contratos (aba Créditos)

def ler_contratos(cred, R):
    out = []
    for c in cred:
        desc = limpo(c.get('descricao'))
        emp = limpo(c.get('empresa'))
        if not desc or not emp:
            continue
        contrato = limpo(c.get('contrato'))
        banco_txt = limpo(c.get('banco'))
        credor = R.credor(banco_txt) or R.credor(desc) or (titulo(banco_txt) if banco_txt else '') or titulo(desc.split()[0])
        if 'CONSORCIO' in norm(desc):
            credor = f'{credor} (consórcio)'
        lib = data(c.get('liberacao'))
        nome = f'{desc} ({contrato})' if contrato else (f'{desc} (CEF)' if 'CEF' in credor else desc)
        nparc, npagas = numero(c.get('nparc')), numero(c.get('npagas'))
        out.append({
            'empresa': emp, 'liberacao': lib.strftime('%d/%m/%Y') if lib else '—', 'contrato': contrato or '—',
            'descricao': desc, 'nparc': nparc, 'npagas': npagas,
            'liberado': numero(c.get('liberado')), 'parcela': numero(c.get('parcela')), 'saldo': numero(c.get('saldo')),
            'banco': credor, 'venc': numero(c.get('venc')), 'juros': numero(c.get('juros')),
            'indexador': limpo(c.get('indexador')) or '—', 'situacao_planilha': limpo(c.get('situacao')) or '—',
            'info': limpo(c.get('info')) or '—',
            # uso interno
            '_nome': nome, '_num': re.sub(r'[^A-Z0-9]', '', contrato.upper())[:9] if contrato else '',
            '_chave': re.sub(r'[^A-Z ]', ' ', norm(desc)).split(), '_emp': chave(emp), '_lib': lib,
        })
    return out


GENERICAS = {'EMPRESTIMO', 'EMPREST', 'EMP', 'CREDITO', 'GIRO', 'PARCELA', 'CONTRATO', 'OPERACAO', 'PJ', 'PF'}
NUM_CONTRATO = re.compile(r'(?<![A-Z0-9])C\d{8,10}(?![0-9])')


def credor_base(c):
    return c['banco'].replace(' (consórcio)', '')


def achar_contrato(r, contratos, credor_linha):
    """Liga um lançamento de empréstimo a um contrato da aba Créditos.
    Ordem: número do contrato no texto; valor da parcela (ou da liberação) na mesma empresa e no mesmo banco;
    nome do contrato no texto (pelo menos duas palavras próprias). Um número de contrato que não está na aba
    Créditos impede as demais ligações (é outro contrato)."""
    texto = norm(' '.join([r['_lanc'], r['_desc'], r['_info']]))
    compacto = re.sub(r'[^A-Z0-9]', '', texto)
    for tam in (9, 8):
        for c in contratos:
            if len(c['_num']) >= 8 and c['_num'][:tam] in compacto:
                return c, 'numero'
    if NUM_CONTRATO.search(texto):
        return None, None
    mesma = [c for c in contratos if c['_emp'] == r['_empk'] and (not credor_linha or credor_base(c) == credor_linha)]
    alvo = 'liberado' if r['mv'] == 0 else 'parcela'
    for c in mesma:
        v = c[alvo]
        if v and abs(r['v'] - v) <= max(0.01, v * 0.002):
            return c, 'valor'
    palavras = [w for w in re.sub(r'[^A-Z ]', ' ', texto).split() if len(w) >= 3]
    acha = lambda k: any(w == k or (len(w) >= 6 and len(k) >= 6 and (k.startswith(w) or w.startswith(k))) for w in palavras)
    for c in mesma:
        ks = [k for k in c['_chave'] if len(k) > 2 and k not in GENERICAS]
        if ks and sum(map(len, ks)) >= 7 and all(acha(k) for k in ks):
            return c, 'nome'
    return None, None


def nome_fora_da_aba(num, textos):
    """Nome de um contrato que não está na aba Créditos, tirado das descrições dos lançamentos.
    'C324201202 - INVESTIMENTO EMPRESARIAL - PARC 33/48' -> 'Investimento empresarial (C324201202), fora da aba Créditos'."""
    ruins = ('LIQUIDA', 'EMPRESTIMO', 'PARCELA', 'PARC ', 'AMORTIZA', 'CONTRATO', 'PAGAMENTO')
    cont = collections.Counter()
    for t in textos:
        for p in re.split(r'\s+-\s+|-(?=C\d)|·', t):
            p = p.strip()
            if p and not re.search(r'\d', p) and len(p) > 3 and not any(b in p for b in ruins):
                cont[p] += 1
    nome = cont.most_common(1)[0][0].capitalize() if cont else 'Contrato'
    return f'{nome} ({num}), fora da aba Créditos'


# ---------------------------------------------------------------- montagem

def montar(fluxo, cred, de_para, cfg, hoje):
    R = Regras(cfg, de_para)
    empresas = {chave(e): titulo(e) for e in cfg['empresas']}
    q = collections.Counter()
    brutas = []
    for x in fluxo:
        lanc = limpo(x.get('lanc'))
        d = data(x.get('data'))
        v = numero(x.get('valor'))
        if norm(lanc) == 'SALDO INICIAL' or not d or not v:
            continue
        ek = chave(x.get('empresa'))
        if ek not in empresas:
            q['fora_empresa'] += 1
            continue
        mov = norm(x.get('mov'))
        mv = 0 if mov.startswith('ENTRADA') else 1 if mov.startswith('SAIDA') else (0 if v > 0 else 1)
        previsto = norm(x.get('fc')).startswith('PREVISTO')
        brutas.append({
            '_d': d, 'v': round(abs(v), 2), 'mv': mv, 'st': 1 if previsto else 0, '_emp': empresas[ek], '_empk': ek,
            '_lanc': lanc, '_desc': limpo(x.get('desc')), '_info': limpo(x.get('info')),
            '_cat': limpo(x.get('cat')), '_sub': limpo(x.get('sub')), '_banco': limpo(x.get('banco')),
            '_pf': norm(x.get('pf')), '_fv': norm(x.get('fv')), '_doc': limpo(x.get('tipodoc')),
            '_sem': numero(x.get('semana')),
        })
    if not brutas:
        raise ErroPainel('a aba de fluxo não tem lançamentos com data e valor')

    # data-base = último dia com lançamento realizado (sem passar de hoje)
    realizados = [b['_d'] for b in brutas if not b['st'] and b['_d'] <= hoje]
    if not realizados:
        raise ErroPainel('nenhum lançamento realizado encontrado na aba de fluxo')
    base = max(realizados)
    ano = base.year if str(cfg.get('ano', 'auto')).lower() == 'auto' else int(cfg['ano'])
    fora_ano = [b for b in brutas if b['_d'].year != ano]
    L = [b for b in brutas if b['_d'].year == ano]
    if len(L) < cfg.get('linhas_minimas', 1):
        raise ErroPainel(f'a aba de fluxo tem só {len(L)} lançamentos em {ano} (mínimo esperado: {cfg.get("linhas_minimas")})')
    contratos = ler_contratos(cred, R)

    # 1) categoria, natureza, subcategoria, contraparte
    for r in L:
        cat, nat = R.categoria(r['_cat'], r['_sub'])
        if nat == 'MUT' and not R.do_grupo(r['_lanc']):
            cat, nat, _ = R.cfg['mutuo_com_socios']
        r['cat'], r['_nat'] = cat, nat
        r['sub'] = R.subcategoria(r['_sub'], cat)
        r['_orig'] = ' › '.join(x for x in (r['_cat'], r['_sub']) if x) or 'sem categoria'

    # 2) transferências sem categoria que batem em data e valor com o lado oposto em outra conta
    idx = collections.defaultdict(list)
    for r in L:
        idx[(r['_d'], r['v'], r['mv'])].append(r)
    q['transf_n'] = 0
    q['transf_v'] = 0.0
    for r in L:
        if r['_nat'] == 'NC' and not r['st'] and re.search(r'\b(PIX|TED|DOC|TRANSF\w*)\b', norm(r['_lanc'] + ' ' + r['_desc'])):
            par = [o for o in idx[(r['_d'], r['v'], 1 - r['mv'])]
                   if norm(o['_banco']) != norm(r['_banco']) and not o['st'] and o['_nat'] in ('INT', 'NC')]
            if par:
                r['cat'], r['_nat'], r['sub'] = 'Movimentações internas', 'INT', 'Transferência entre contas'
                q['transf_n'] += 1
                q['transf_v'] += r['v']

    # 3) empréstimos: liga ao contrato, separa parcela / custos / liquidação / liberação
    reclass = []
    so_numero = [c for c in contratos if c['_num']]
    pendentes = []
    for r in L:
        texto = ' '.join((r['_lanc'], r['_desc'], r['_info']))
        r['_credor'] = R.credor(texto) or R.credor(r['_banco'])
        e_emp = r['_nat'] == 'EMP'
        c, como = achar_contrato(r, contratos if e_emp else so_numero, r['_credor'])
        if not e_emp and como != 'numero':
            r['loan'] = r['lender'] = r['lkind'] = ''
            continue
        if not e_emp:  # número de contrato no texto, mas lançado em outra categoria
            if r['mv'] == 1 and not re.match(r'(IOF|TARIFA|SEGURO)', norm(r['_lanc'])):
                reclass.append({'data': r['_d'].strftime('%d/%m/%Y'), 'valor': r['v'], 'de': r['_orig'],
                                'desc': ' '.join(texto.split())[:60], 'contrato': c['_nome']})
                r['cat'], r['_nat'], r['sub'] = 'Empréstimos e financiamentos', 'EMP', 'Empréstimos'
            else:
                r['loan'] = r['lender'] = r['lkind'] = ''
                continue
        nl = norm(texto)
        r['_como'] = como
        if c:
            r['loan'], r['lender'] = c['_nome'], c['banco']
        else:
            cred = r['_credor'] or 'Não identificado'
            num = NUM_CONTRATO.search(nl)
            if r['cat'] == 'Consórcios' or 'CONSORCIO' in nl:
                r['loan'], r['lender'] = titulo(nome_base(r['_lanc'] or 'Consórcio')).capitalize(), f'{cred} (consórcio)'
            elif num:
                r['loan'], r['lender'] = '__NUM__' + num.group(0), cred
            elif r['mv'] == 0:
                r['loan'], r['lender'] = 'Liberação de crédito, contrato não identificado', cred
            else:
                r['loan'], r['lender'] = f'Empréstimo {cred}, contrato não identificado', cred
                pendentes.append(r)
    por_num = collections.defaultdict(list)
    for r in L:
        if r.get('loan', '').startswith('__NUM__'):
            por_num[r['loan'][7:]].append(norm(' · '.join((r['_lanc'], r['_desc'], r['_info']))))
    nomes_num = {n: nome_fora_da_aba(n, ts) for n, ts in por_num.items()}
    for r in L:
        if r.get('loan', '').startswith('__NUM__'):
            r['loan'] = nomes_num[r['loan'][7:]]
    # previstos sem número: liga pelo valor da última parcela paga de cada contrato (parcela variável, até 3%)
    ultima = {}
    for r in sorted((r for r in L if r.get('_como') and not r['st'] and r['mv'] == 1), key=lambda r: r['_d']):
        if not re.match(r'(IOF|TARIFA|SEGURO)', norm(r['_lanc'])) and 'LIQUIDACAO CONTRATO' not in norm(r['_desc']):
            ultima[r['loan']] = (r['v'], r['_empk'], r['lender'])
    for r in pendentes:
        cands = [(k, cr) for k, (v, ek, cr) in ultima.items()
                 if ek == r['_empk'] and abs(r['v'] - v) <= v * 0.03 and (not r['_credor'] or cr.replace(' (consórcio)', '') == r['_credor'])]
        if len(cands) == 1:
            r['loan'], r['lender'] = cands[0]
            r['_como'] = 'valor'
    for r in L:
        if not r.get('loan'):
            continue
        nl = norm(' '.join((r['_lanc'], r['_desc'], r['_info'])))
        como = r.get('_como')
        if r['mv'] == 0:
            r['lkind'] = 'Liberação de crédito'
        elif 'LIQUIDACAO CONTRATO' in nl or 'LIQUIDADO' in norm(r['_lanc']):
            r['lkind'] = 'Liquidação antecipada'
        elif como == 'numero' and re.match(r'(IOF|TARIFA|SEGURO)', norm(r['_lanc'])):
            r['lkind'] = CUSTOS
        else:
            r['lkind'] = 'Parcela'

    # 4) conta, cartão, PF/PJ, fixo/variável, imposto, documento, contraparte
    for r in L:
        conta, meio, card = R.conta(r['_banco'], r['st'])
        nl = norm(r['_lanc'])
        if r['cat'] == FATURA or nl.startswith('CARTAO'):
            meio = 'Cartão de crédito'
            card = R.cartao(r['_lanc'] + ' ' + r['_desc']) or card or 'Outro'
        r['conta'], r['meio'], r['card'] = conta, meio, card
        if r['_pf'] in ('PF', 'PJ'):
            r['pf'] = r['_pf']
        elif r['cat'] in ('Retiradas / despesas pessoais de sócios', 'Distribuição de lucros'):
            r['pf'] = 'PF'
        elif r['cat'] == FATURA:
            r['pf'] = 'Não identificado'
        else:
            r['pf'] = 'PJ'
        if r['mv'] == 0:
            r['fv'] = 'Não se aplica (entrada)'
        elif r['_fv'].startswith('FIX'):
            r['fv'] = 'Fixo'
        elif r['_fv'].startswith('VAR'):
            r['fv'] = 'Variável'
        elif chave(r['sub']) in R.fixos or r.get('lkind') == 'Parcela':
            r['fv'] = 'Fixo'
        else:
            r['fv'] = 'Variável'
        r['tax'], r['taxp'] = R.imposto(r['_sub'], r['_lanc'], r['mv'])
        doc = norm(r['_doc'])
        r['doc'] = 'Não informado' if doc in ('', '-') else 'Nota de Debito' if doc == 'ND' else 'Nota Fiscal' if doc == 'NF' else titulo(r['_doc'])
        r['desc'] = ' · '.join(x for x in (r['_desc'], r['_info']) if x) if r['_desc'] else ('· ' + r['_info'] if r['_info'] else '')
        r['nat'], r['emp'] = NATUREZAS[r['_nat']], r['_emp']
        r['m'], r['d'], r['mdt'] = r['_d'].month, r['_d'].day, r['_d'].month
        r['w'] = int(r['_sem']) if r['_sem'] and 1 <= r['_sem'] <= 6 else (r['_d'].day - 1) // 7 + 1

    # contraparte: apelidos, sem nº de parcela, variações de grafia unificadas
    for r in L:
        base_nome = nome_base(r['_lanc']) or nome_base(r['_desc']) or nome_base(r['_info']) or 'Não informado'
        apel = R.apelido(base_nome)
        if apel:
            r['_cp'], r['_fixo'] = apel, True
        elif r['mv'] == 0 and re.match(r'^CRED? ?PIX', norm(base_nome)):
            r['_cp'], r['_fixo'] = 'PIX não identificado', True
        else:
            letras = [c for c in base_nome if c.isalpha()]
            caixa_alta = letras and sum(c.isupper() for c in letras) / len(letras) > 0.6
            r['_cp'], r['_fixo'] = (titulo(base_nome) if caixa_alta else base_nome), False
    grupos = collections.defaultdict(collections.Counter)
    for r in L:
        if not r['_fixo']:
            grupos[chave_nome(r['_cp'])][r['_cp']] += 1
    rotulo = {k: c.most_common(1)[0][0] for k, c in grupos.items()}
    for r in L:
        r['cp'] = r['_cp'] if r['_fixo'] else rotulo[chave_nome(r['_cp'])]

    # 5) cartão sem dupla contagem: fatura sai quando o mês tem as compras detalhadas daquele cartão
    detalhe = {(r['m'], r['card']) for r in L if r['meio'] == 'Cartão de crédito' and r['cat'] != FATURA
               and norm(r['_banco']) in R.bancos and R.bancos[norm(r['_banco'])][1] == 'Cartão de crédito'}
    detalhe_mes = {m for m, _ in detalhe}
    fat_fora = []
    for r in L:
        if r['cat'] == FATURA and ((r['m'], r['card']) in detalhe or (r['card'] in ('', 'Outro') and r['m'] in detalhe_mes)):
            fat_fora.append(r)
    ids = {id(r) for r in fat_fora}
    L = [r for r in L if id(r) not in ids]
    meses_detalhe = sorted(detalhe_mes)

    # 6) possíveis duplicidades em empréstimos previstos (mesmo contrato, data e valor)
    dup = collections.Counter((r['loan'], r['_d'], r['v']) for r in L if r.get('loan') and r['st'] and r['mv'] == 1)
    dups = [(k, n) for k, n in dup.items() if n > 1]

    # 7) contratos: status e conferência com o fluxo
    por_loan = collections.defaultdict(list)
    for r in L:
        if r.get('loan'):
            por_loan[r['loan']].append(r)
    avisos = []
    for c in contratos:
        rs = por_loan.get(c['_nome'], [])
        liq = [r for r in rs if r['lkind'] == 'Liquidação antecipada' and not r['st']]
        pag = sorted((r for r in rs if r['lkind'] == 'Parcela' and r['mv'] == 1 and not r['st']), key=lambda r: r['_d'])
        prev = [r for r in rs if r['mv'] == 1 and r['st']]
        libs = [r for r in rs if r['mv'] == 0 and not r['st']]
        meses_pagos = sorted({r['m'] for r in pag})
        inicia = re.search(r'INICIA\w*\s+(?:EM\s+)?(\d{4})', norm(c['info'] + ' ' + c['situacao_planilha']))
        if liq:
            u = max(liq, key=lambda r: r['_d'])
            c['status'] = 'Encerrado'
            c['obs'] = f'Liquidado em {dma(u["_d"])} ({brl(sum(r["v"] for r in liq))} no fluxo).'
        elif inicia and int(inicia.group(1)) > ano and not pag:
            c['status'] = f'Carência (inicia {inicia.group(1)})'
            c['obs'] = (f'Liberado em {dma(libs[0]["_d"])} ({brl(libs[0]["v"])}). ' if libs else '') + f'Sem parcela no fluxo de {ano}.'
        elif pag and not prev and (pag[-1]['m'] < base.month):
            c['status'] = 'Encerrado'
            c['obs'] = f'Última parcela no fluxo em {dma(pag[-1]["_d"])}; nenhuma prevista.'
        elif not pag and not prev and norm(c['situacao_planilha']) in ('QUITADO', 'LIQUIDADO', 'ENCERRADO'):
            c['status'] = 'Encerrado'
            c['obs'] = f'Sem pagamento no fluxo de {ano}; a aba marca "{c["situacao_planilha"]}".'
        else:
            c['status'] = 'Ativo'
            partes = []
            if pag:
                vals = [r['v'] for r in pag]
                faixa = brl(min(vals)) if max(vals) - min(vals) < 1 else f'{brl(min(vals))} a {brl(max(vals))}'
                partes.append(f'{len(pag)} parcela(s) paga(s) no fluxo em {ano} ({faixa}), desde {MN[pag[0]["m"] - 1]}.')
            else:
                partes.append(f'Nenhuma parcela paga no fluxo de {ano}.')
            if prev:
                partes.append(f'{len(prev)} prevista(s) até dez.')
            c['obs'] = ' '.join(partes)
        # conferências que viram ponto de atenção
        if c['status'] != 'Encerrado' and c['npagas'] is not None and c['_lib'] and c['_lib'].year == ano and len(meses_pagos) > c['npagas']:
            c['obs'] += f' A aba indica {int(c["npagas"])} parcela(s) paga(s), mas o fluxo já registra {len(meses_pagos)}.'
            avisos.append(f'a aba Créditos indica {int(c["npagas"])} parcela(s) paga(s) em "{c["descricao"]}", mas o fluxo já registra {len(meses_pagos)}')
        if norm(c['situacao_planilha']) in ('QUITADO', 'LIQUIDADO') and prev:
            avisos.append(f'"{c["descricao"]}" aparece como {c["situacao_planilha"]} na aba Créditos, mas o fluxo ainda prevê {len(prev)} pagamento(s)')
    sem_contrato = collections.defaultdict(lambda: [0.0, 0.0])
    liquidados = {r['loan'] for r in L if r.get('lkind') == 'Liquidação antecipada' and 'não identificado' not in r['loan']}
    for r in L:
        if (r.get('loan') and r['mv'] == 1 and r['lkind'] == 'Parcela' and r['loan'] not in liquidados
                and ('não identificado' in r['loan'] or 'fora da aba' in r['loan'])):
            sem_contrato[r['loan']][r['st']] += r['v']
    for nome, (pago, prev) in sorted(sem_contrato.items(), key=lambda x: -sum(x[1])):
        partes = ' e '.join(x for x in (f'{brl(pago)} pagos' if pago else '', f'{brl(prev)} previstos' if prev else '') if x)
        avisos.append(f'"{nome}" tem {partes} em {ano} sem contrato correspondente na aba Créditos')
    if dups:
        porv = collections.defaultdict(list)
        for (loan, d, v), n in dups:
            porv[(loan, v)].append(d)
        for (loan, v), ds in porv.items():
            meses = ', '.join(MN[d.month - 1] for d in sorted(ds))
            avisos.append(f'a previsão de {meses} traz {brl(v)} em duplicidade na mesma data para "{loan}"; se for erro, as saídas previstas estão {brl(v)} acima em cada mês')

    # 8) saída no formato do painel
    dims = {k: [] for k in COMDIM}
    pos = {k: {} for k in COMDIM}
    rows = []
    for r in sorted(L, key=lambda r: (r['_d'], r['mv'])):
        for k in ('loan', 'lender', 'lkind', 'tax', 'taxp'):
            r.setdefault(k, '')
        lin = []
        for f in CAMPOS:
            val = r[f]
            if f in pos:
                if val not in pos[f]:
                    pos[f][val] = len(dims[f])
                    dims[f].append(val)
                val = pos[f][val]
            lin.append(val)
        rows.append(lin)
    nc = [r for r in L if r['_nat'] == 'NC' and r['mv'] == 1]
    socios = sorted({r['cp'].split()[0] for r in L if r['cat'] == R.cfg['mutuo_com_socios'][0]})
    quality = {
        'data_base': base.strftime('%d/%m/%Y'), 'ano': ano, 'linhas': len(L), 'linhas_fora_ano': len(fora_ano),
        'linhas_fora_empresa': q['fora_empresa'],
        'fatura_excluida_valor': round(sum(r['v'] for r in fat_fora), 2), 'fatura_excluida_n': len(fat_fora),
        'meses_cartao_detalhado': meses_detalhe,
        'transf_reclass_valor': round(q['transf_v'], 2), 'transf_reclass_n': q['transf_n'],
        'nao_classificado_valor': round(sum(r['v'] for r in nc), 2), 'nao_classificado_n': len(nc),
        'reclass_emprestimo': reclass, 'socios': socios, 'avisos_divida': avisos,
    }
    credits = [{k: v for k, v in c.items() if not k.startswith('_')} for c in contratos]
    return {
        'fields': CAMPOS, 'dims': dims, 'rows': rows, 'credits': credits, 'quality': quality,
        'catcode': R.catcode, 'nat': NATUREZAS, 'cur': base.month, 'ano': ano,
    }


# ---------------------------------------------------------------- saída

def criptografar(texto, senha):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    chave_aes = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=SAL, iterations=ITERACOES).derive(senha.encode('utf-8'))
    iv = os.urandom(12)
    ct = AESGCM(chave_aes).encrypt(iv, texto.encode('utf-8'), None)
    b64 = lambda b: base64.b64encode(b).decode('ascii')
    return {'s': b64(SAL), 'i': b64(iv), 'c': b64(ct), 'n': ITERACOES}


def gravar(D, cfg, saida):
    senha = os.environ.get('PAINEL_SENHA', '')
    if len(senha) < 8:
        raise ErroPainel('o secret PAINEL_SENHA não está cadastrado ou tem menos de 8 caracteres')
    modelo = (RAIZ / 'modelo.html').read_text(encoding='utf-8')
    if '/*__CIFRA__*/null' not in modelo:
        raise ErroPainel('modelo.html sem o marcador /*__CIFRA__*/null')
    cifra = criptografar(json.dumps(D, ensure_ascii=False, separators=(',', ':')), senha)
    pagina = modelo.replace('/*__CIFRA__*/null', json.dumps(cifra)).replace('__TITULO__', html.escape(cfg.get('titulo', 'Painel')))
    saida.mkdir(parents=True, exist_ok=True)
    (saida / 'index.html').write_text(pagina, encoding='utf-8')
    (saida / 'robots.txt').write_text('User-agent: *\nDisallow: /\n', encoding='utf-8')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--xlsx', help='lê uma cópia local (.xlsx) em vez da planilha online (teste)')
    ap.add_argument('--hoje', help='data de referência AAAA-MM-DD (teste)')
    ap.add_argument('--saida', default=str(RAIZ / 'site'), help='pasta do site gerado')
    ap.add_argument('--json', help='grava também os dados tratados, sem criptografia, neste arquivo (só para teste local)')
    a = ap.parse_args()
    try:
        try:
            cfg = json.loads((RAIZ / 'config.json').read_text(encoding='utf-8-sig'))
        except json.JSONDecodeError as e:
            raise ErroPainel(f'config.json com erro de formato na linha {e.lineno} (confira vírgulas e aspas)') from None
        agora = dt.datetime.now(BRT)
        hoje = dt.date.fromisoformat(a.hoje) if a.hoje else agora.date()
        if a.xlsx:
            conteudo = Path(a.xlsx).read_bytes()
        else:
            conteudo = baixar(os.environ.get(cfg.get('segredo_planilha', 'PLANILHA_URL'), '').strip())
        fluxo, cred, de_para = ler(conteudo, cfg)
        D = montar(fluxo, cred, de_para, cfg, hoje)
        D['gerado'] = agora.strftime('%d/%m/%Y %H:%M')
        D['abas'] = cfg['abas']
        if a.json:
            Path(a.json).write_text(json.dumps(D, ensure_ascii=False, indent=1), encoding='utf-8')
        gravar(D, cfg, Path(a.saida))
    except ErroPainel as e:
        print(f'ERRO: {e}. O painel publicado não foi alterado.', file=sys.stderr)
        sys.exit(1)
    except Exception as e:  # noqa: BLE001
        # Sem a mensagem original: ela pode conter dados da planilha e o log do GitHub é público.
        print(f'ERRO inesperado ({type(e).__name__}). O painel publicado não foi alterado.', file=sys.stderr)
        for fr in traceback.extract_tb(e.__traceback__):
            print(f'  {Path(fr.filename).name}:{fr.lineno} em {fr.name}', file=sys.stderr)
        sys.exit(1)
    Q = D['quality']
    print(f'Painel gerado: {Q["linhas"]} lançamentos de {Q["ano"]}, realizado até {Q["data_base"]}.')
    if Q['nao_classificado_n']:
        print(f'Aviso: {Q["nao_classificado_n"]} saídas sem categoria aparecem como "Não classificado".')
    if Q['linhas_fora_empresa']:
        print(f'Aviso: {Q["linhas_fora_empresa"]} lançamentos de empresa fora da lista de config.json ficaram de fora.')


if __name__ == '__main__':
    main()
