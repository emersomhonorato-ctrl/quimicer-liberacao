#!/usr/bin/env python3
"""Backup Quimicer - Etapa 7: gera as planilhas FR 001, FR 002 e FR 011 a partir do banco,
no mesmo formato do botao "Exportar as 3 FR" do app.
Entrada: pasta com itens.json, registros.json, lotes.json, config.json (exportados via psql).
Uso: python3 gerar_excel.py <pasta_json> <pasta_saida>"""
import json, sys, os, re, math, unicodedata, datetime as dt
from zoneinfo import ZoneInfo
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

SRC, OUT = sys.argv[1], sys.argv[2]
HOJE = dt.datetime.now(ZoneInfo('America/Sao_Paulo')).date()
MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

F_DENS = dict(k='dens', nome='Densidade', unit='g/cm³', tipo='faixa', mn='densMin', mx='densMax')
F_PH = dict(k='ph', nome='pH', tipo='faixa', mn='phMin', mx='phMax')
F_VISC = dict(k='visc', nome='Viscosidade', unit='s', tipo='faixa', mn='viscMin', mx='viscMax')
MODS = {
  'mp': dict(key='mp', fr='FR 001', rev='00', title='Liberação de Matéria-Prima', itemCol='Matéria-prima', aux=('fornecedor','Fornecedor'), cfg='mp', validade=True, laudo=True,
    fields=[F_DENS, F_PH, F_VISC, dict(k='colagem', nome='Colagem', tipo='aprov', txt='colagem'), dict(k='aparencia', nome='Aparência', unit='visual', tipo='aprov', txt='visual'), dict(k='odor', nome='Odor', tipo='aprov', txt='odor')]),
  'pa': dict(key='pa', fr='FR 002', rev='00', title='Liberação de Produto Acabado', itemCol='Produto', aux=('op','Ordem de produção'), cfg='geral', validade=False, laudo=False,
    fields=[F_DENS, F_PH, F_VISC, dict(k='umid', nome='Umidade', unit='%', tipo='livre'), dict(k='imperm', nome='Impermeabilidade', tipo='aprov', txt='diluicao'),
            dict(k='colagem', nome='Colagem', tipo='aprov', txt='colagem'), dict(k='aparencia', nome='Aparência', unit='visual', tipo='aprov', txt='visual')]),
  'pi': dict(key='pi', fr='FR 011', rev='01', title='Liberação de Produto Intermediário', itemCol='Intermediário', aux=('op','Ordem de produção'), cfg='pi', validade=False, laudo=False,
    fields=[F_DENS, F_PH, F_VISC, dict(k='umid', nome='Umidade', unit='%', tipo='faixa', mn='umidMin', mx='umidMax'), dict(k='residuo', nome='Resíduo', unit='%', tipo='faixa', mn='resMin', mx='resMax'),
            dict(k='colagem', nome='Colagem', tipo='aprov', txt='colagem'), dict(k='defloc', nome='Defloculação', tipo='aprov', txt='defloculacao'), dict(k='aparencia', nome='Aparência', unit='visual', tipo='aprov', txt='visual'),
            dict(k='pesoIni', nome='Peso inicial', unit='kg', tipo='livre'), dict(k='pesoFin', nome='Peso final', unit='kg', tipo='livre'), dict(k='perda', nome='Perda na reação', unit='%', tipo='calc')]),
}
AV = {'ok': 'CONFORME', 'nc': 'N/C', 'na': 'N/A', 'pend': 'NÃO AVALIADO', 'free': ''}
VSTAT = {'venc': 'Vencida', 'prox': 'Vence em breve', 'ok': 'Dentro da validade', 'lib': 'Liberada pela direção', 'sem': 'Sem validade', 'enc': 'Inativo'}

def load(n):
    with open(os.path.join(SRC, n), encoding='utf-8') as f: return json.load(f) or []
def slug(s):
    s = unicodedata.normalize('NFKD', s); s = ''.join(c for c in s if not unicodedata.combining(c))
    return (re.sub(r'^-|-$', '', re.sub(r'[^A-Z0-9]+', '-', s.upper()))[:120]) or 'X'
def normNome(s): return re.sub(r'\s+', ' ', str(s or '').strip()).upper()
def parseNum(v):
    if v is None or isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v)
    t = str(v).strip().replace(',', '.')
    try: n = float(t) if t else None
    except ValueError: return None
    return n if n is not None and math.isfinite(n) else None
def fmt(n):
    if n is None or n == '': return '—'
    if isinstance(n, (int, float)) and not isinstance(n, bool):
        s = ('%.4f' % n).rstrip('0').rstrip('.'); return s.replace('.', ',')
    return str(n)
def check(v, mn, mx):
    if mn is None or mx is None: return 'na' if v is None or v == '' else 'free'
    n = parseNum(v)
    if n is None: return 'pend'
    return 'ok' if mn - 1e-9 <= n <= mx + 1e-9 else 'nc'
def checkDens(v, mn, mx):
    b = check(v, mn, mx)
    if b != 'nc': return b
    n = parseNum(v); dec = lambda x: len(repr(float(x)).split('.')[1].rstrip('0')) if float(x) != int(float(x)) else 0
    d = max(dec(mn), dec(mx), 2); m = 10 ** d
    tr = math.floor(n * m + 1e-9) / m; ro = math.floor(n * m + 1e-7 + 0.5) / m
    ok = lambda x: mn - 1e-9 <= x <= mx + 1e-9
    return 'ok' if ok(tr) or ok(ro) else 'nc'
def textCheck(v):
    if v is None or v == '': return 'na'
    t = str(v).upper()
    return 'nc' if re.search(r'N[ÃA]O|NOK|REPROV', t) else ('ok' if t == 'OK' else 'free')
def perda(r):
    a, b = parseNum(r.get('pesoIni')), parseNum(r.get('pesoFin'))
    return round((a - b) / a * 10000) / 100 if a is not None and b is not None and a > 0 else None
def dateOf(s):
    try: return dt.date.fromisoformat(str(s)[:10])
    except Exception: return None

itens, regs, lotes, cfgs = load('itens.json'), load('registros.json'), load('lotes.json'), load('config.json')
CFG = {c['chave']: c['dados'] or {} for c in cfgs}
FRREV = CFG.get('frrev', {})
ITENS = {}
for it in itens: ITENS.setdefault(it['modulo'], {})[it['id']] = it['dados'] or {}
LOTES = {l['id']: l['dados'] or {} for l in lotes}

def specOf(rec, mod):
    p = ITENS.get(mod['key'], {}).get(slug(normNome(rec.get('produto'))))
    if p:
        o = {'extras': p.get('extras') or []}
        for f in mod['fields']:
            if f['tipo'] == 'faixa': o[f['mn']] = p.get(f['mn']); o[f['mx']] = p.get(f['mx'])
        return o, p
    return rec.get('spec') or {}, {}

def evalRec(rec, mod):
    sp, item = specOf(rec, mod); r = {'x': {}}; fora = []
    for f in mod['fields']:
        v = rec.get(f['k'])
        if f['tipo'] == 'faixa': s = (checkDens if f['k'] == 'dens' else check)(v, sp.get(f['mn']), sp.get(f['mx']))
        elif f['tipo'] in ('livre', 'calc'): s = 'na' if v is None or v == '' else 'free'
        else: s = (('pend' if f.get('txt') and item.get(f['txt']) else 'na') if v is None or v == '' else textCheck(v))
        r[f['k']] = s
        if s == 'nc': fora.append(f['nome'])
    for e in (sp.get('extras') or []) if isinstance(sp.get('extras'), list) else []:
        v = (rec.get('extras') or {}).get(e.get('id'))
        s = (textCheck(v) if v else 'pend') if e.get('tipo') == 'aprovacao' else check(v, e.get('min'), e.get('max'))
        r['x'][e.get('id')] = s
        if s == 'nc': fora.append(e.get('nome') or e.get('id'))
    if mod['validade'] and rec.get('validade') and rec.get('data') and rec['validade'] < rec['data']: fora.append('Validade')
    r['fora'] = fora; return r, sp

def lotInfo(rec):
    L = LOTES.get(re.sub(r'[^A-Za-z0-9_\-.~:@+]', '_', rec['_id']), {})
    val = L.get('validade') or rec.get('validade') or ''; d = L.get('direcao')
    vd = dateOf(val)
    if L.get('encerrado'): st = 'enc'
    elif not vd: st = 'sem'
    elif vd < HOJE: st = 'lib' if d and (not d.get('ate') or (dateOf(d['ate']) or HOJE) >= HOJE) else 'venc'
    else: st = 'prox' if (vd - HOJE).days <= 30 else 'ok'
    return vd, VSTAT[st]

def recs_by_month(mod):
    meses = {}
    rows = sorted([g for g in regs if g['modulo'] == mod['key']], key=lambda g: (g['mes'], g['origem'] != 'planilha', g['ordem'] if g['ordem'] is not None else 0))
    idx = {}
    for g in rows:
        r = dict(g['dados'] or {}); m = g['mes']
        if g['origem'] == 'planilha':
            i = idx.get(m, 0); idx[m] = i + 1; r['_src'] = 'arquivo'; r['_id'] = f'{m}:a{i}'
        else: r['_src'] = 'live'; r['_id'] = f"{m}:{g['id']}"
        if mod['key'] == 'pi' and r.get('perda') is None: r['perda'] = perda(r)
        meses.setdefault(m, []).append(r)
    return meses

def planilha(mod):
    cols = ['Data', mod['itemCol'], 'Lote', mod['aux'][1]] + (['Validade', 'Situação da validade'] if mod['validade'] else []) + (['Destino'] if mod['key'] == 'pa' else [])
    for f in mod['fields']:
        u = f" ({'s ou cP' if f['k'] == 'visc' else f['unit']})" if f.get('unit') and f['tipo'] != 'aprov' else ''
        cols += [f['nome'] + u, f['nome'] + ' mín', f['nome'] + ' máx', f['nome'] + ' avaliação'] if f['tipo'] == 'faixa' else [f['nome'] + u]
    cols += ['Outros ensaios', 'Situação', 'Fora da especificação', 'Nº RNC', 'Status da RNC', 'Analisado por', 'Observação', 'Origem'] + (['Laudo'] if mod['laudo'] else [])
    num = lambda v: '' if v is None or v == '' else v
    def linha(r):
        ev, sp = evalRec(r, mod)
        v = [dateOf(r.get('data')) or '', r.get('produto') or '', r.get('lote') or '', r.get(mod['aux'][0]) or '']
        if mod['validade']: vd, st = lotInfo(r); v += [vd or '', st]
        if mod['key'] == 'pa': v.append('Cliente: ' + r['cliente'] if r.get('cliente') else ('Estoque' if r.get('destino') == 'estoque' else ''))
        for f in mod['fields']:
            if f['tipo'] == 'faixa': v += [num(r.get(f['k'])), num(sp.get(f['mn'])), num(sp.get(f['mx'])), AV.get(ev[f['k']], '')]
            else: v.append(num(r.get(f['k'])))
        ex = []
        for e in (sp.get('extras') or []):
            xv = (r.get('extras') or {}).get(e.get('id'))
            ex.append(f"{e.get('nome')}={'-' if xv is None else (fmt(xv) if isinstance(xv, (int, float)) else xv)}{(' ' + e['unidade']) if e.get('unidade') and e.get('tipo') != 'aprovacao' else ''} [{AV.get(ev['x'].get(e.get('id')), '')}]")
        v += [' | '.join(ex), r.get('situacao') or '', ', '.join(ev['fora']), r.get('rnc') or '', r.get('rncStatus') or '', r.get('analista') or '', r.get('obs') or '',
              ('Planilha' + (' (editado no app)' if r.get('editadoNoApp') else '')) if r['_src'] == 'arquivo' else 'App']
        if mod['laudo']: v.append((r.get('laudo') or {}).get('nome') if isinstance(r.get('laudo'), dict) else (r.get('laudoRef') or ''))
        return [x if not isinstance(x, (dict, list)) else json.dumps(x, ensure_ascii=False) for x in v]
    def chave(r):
        lote = str(r.get('lote') or ''); m = re.match(r'(\D*)(\d*)(.*)', lote)
        return (r.get('data') or '', m.group(1), int(m.group(2)) if m.group(2) else -1, m.group(3))
    meses = recs_by_month(mod)
    rev = (FRREV.get(mod['key']) or {}).get('rev') or mod['rev']
    titulo = f"{mod['fr']} · Rev{rev} · {mod['title']}"
    wb = Workbook(); wb.remove(wb.active)
    def folha(nome, recs, sub):
        ws = wb.create_sheet(nome[:31]); ws.append([titulo]); ws.append([sub]); ws.append([]); ws.append(cols)
        for r in recs: ws.append(linha(r))
        for i, c in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = 34 if i == 2 else (40 if re.search('Observação|Outros|Laudo', c) else max(9, min(22, len(c) + 2)))
        ws.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{4 + max(len(recs), 1)}"
        dcols = [1] + ([5] if mod['validade'] else [])
        for row in ws.iter_rows(min_row=5):
            for c in dcols:
                if isinstance(row[c - 1].value, dt.date): row[c - 1].number_format = 'dd/mm/yyyy'
    todos = sorted([r for m in meses for r in meses[m]], key=chave)
    folha('Todos', todos, f"Todos os meses · {len(todos)} lotes · exportado em {HOJE.strftime('%d/%m/%Y')} pela rotina de backup")
    for m in sorted(meses, reverse=True):
        l = sorted(meses[m], key=chave)
        if l: folha(f"{MESES[int(m[5:7]) - 1]} {m[:4]}".upper(), l, f"{MESES[int(m[5:7]) - 1]} {m[:4]} · {len(l)} lotes")
    wb.properties.title = titulo; wb.properties.creator = 'Quimicer'
    nome = f"{mod['fr'].replace(' ', '_')}_{slug(mod['title'])}_{HOJE.isoformat()}.xlsx"
    wb.save(os.path.join(OUT, nome)); return nome, len(todos)

os.makedirs(OUT, exist_ok=True)
for k in ('mp', 'pa', 'pi'):
    nome, n = planilha(MODS[k]); print(f'{nome}: {n} lotes')
