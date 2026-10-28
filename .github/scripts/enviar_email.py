#!/usr/bin/env python3
"""Backup Quimicer - Etapa 9: envia o e-mail diario (com as planilhas) ou o aviso de falha.
Uso: python3 enviar_email.py ok|falha"""
import os, sys, glob, smtplib, ssl, datetime as dt
from zoneinfo import ZoneInfo
from email.message import EmailMessage

REMETENTE = 'quimicercontroledequalidade@gmail.com'
DESTINO = ['quimicercontroledequalidade@gmail.com', 'laboratoriocq@quimicer.com.br']
senha = ''.join((os.environ.get('GMAIL_APP_PASSWORD') or '').split())
if len(senha) != 16:
    print('::warning::Segredo GMAIL_APP_PASSWORD ausente ou invalido - e-mail nao enviado'); sys.exit(0)

modo = sys.argv[1] if len(sys.argv) > 1 else 'ok'
hoje = dt.datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y')
run = f"{os.environ.get('GITHUB_SERVER_URL')}/{os.environ.get('GITHUB_REPOSITORY')}/actions/runs/{os.environ.get('GITHUB_RUN_ID')}"
def ler(p, d=''):
    try: return open(p, encoding='utf-8').read().strip()
    except Exception: return d

msg = EmailMessage(); msg['From'] = f'Backup Quimicer <{REMETENTE}>'; msg['To'] = ', '.join(DESTINO)
if modo == 'falha':
    msg['Subject'] = f'⚠️ FALHA no backup do sistema de liberação de lotes - {hoje}'
    msg.set_content(f"""O backup automático de hoje ({hoje}) NÃO foi concluído.

Detalhes técnicos da execução: {run}

O backup anterior continua guardado no Google Drive (Backup Quimicer → copias).
A rotina tenta de novo amanhã às 3h. Se a falha se repetir, verifique o sistema.

— Rotina automática de backup · Liberação de Lotes Quimicer""")
else:
    conf = dict(l.split(';', 1) for l in ler('out/conferencia.txt').splitlines() if ';' in l)
    arq = (ler('out/arquivos_resumo.txt', ';;').split(';') + ['', '', ''])[:3]
    zipn = ler('zipinfo.txt', '-')
    msg['Subject'] = f'Backup diário OK - Liberação de Lotes Quimicer - {hoje}'
    msg.set_content(f"""Backup automático de {hoje} concluído com sucesso.

Banco de dados
  • Registros de liberação: {conf.get('registros', '-')}
  • Especificações (itens): {conf.get('itens', '-')}
  • Laudos: {conf.get('laudos', '-')}
  • RNC: {conf.get('rnc', '-')}

PDFs anexados no Drive: {arq[0] or '-'} arquivos ({arq[1] or '0'} novos hoje)
Pacote do dia: {zipn}
Local: Google Drive da conta {REMETENTE} → pasta "Backup Quimicer"
  • copias: banco + fontes + planilhas (as 5 últimas)
  • arquivos/laudos-mp: espelho dos PDFs

Em anexo: planilhas FR 001, FR 002 e FR 011 atualizadas.

— Rotina automática de backup · Liberação de Lotes Quimicer""")
    for f in sorted(glob.glob('out/excel/*.xlsx')):
        msg.add_attachment(open(f, 'rb').read(), maintype='application',
                           subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=os.path.basename(f))

with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ssl.create_default_context(), timeout=60) as s:
    s.login(REMETENTE, senha); s.send_message(msg)
print(f'E-mail ({modo}) enviado para: ' + ', '.join(DESTINO))
