# Roteiro de restauração – Liberação de Lotes Quimicer

Use este roteiro se o sistema (Supabase ou GitHub) for perdido ou corrompido.
Recomenda-se ter apoio de alguém com conhecimento técnico (banco PostgreSQL).

## Onde estão os backups
Google Drive da conta quimicercontroledequalidade@gmail.com → pasta **Backup Quimicer**
- `copias/backup_quimicer_AAAA-MM-DD.zip` – as 5 últimas cópias diárias. Cada uma contém:
  - `banco_sistema.sql.gz` – estrutura + todos os dados do sistema (liberações, especificações, laudos, RNC, revisões ESP, perfis, configurações)
  - `banco_usuarios_arquivos.sql.gz` – usuários de login e o índice dos PDFs anexados
  - `fontes_AAAA-MM-DD.zip` – código do app (site) + `fontes_versao.txt` com a versão exata
  - `excel/` – planilhas FR 001, FR 002 e FR 011 (leitura imediata, sem restaurar nada)
  - `conferencia.txt` – quantidade de registros no dia do backup (para conferir a restauração)
- `arquivos/laudos-mp/` – cópia de todos os PDFs anexados (laudos de MP, FDS, informativos técnicos)

## 1. Consultar os dados sem restaurar
Abra as planilhas da pasta `excel/` do zip mais recente. Elas têm todos os lotes, mês a mês.

## 2. Restaurar o banco num projeto Supabase novo
1. Crie um projeto novo no Supabase (região São Paulo, sa-east-1).
2. Pegue a "connection string" do projeto novo (botão **Connect** → Session pooler).
3. Num computador com o cliente `psql` (PostgreSQL 17):
   ```
   gunzip banco_sistema.sql.gz
   psql "<connection string>" -f banco_sistema.sql
   gunzip banco_usuarios_arquivos.sql.gz
   psql "<connection string>" -f banco_usuarios_arquivos.sql
   ```
4. Confira as contagens com o arquivo `conferencia.txt`:
   `select count(*) from registros;` (e o mesmo para laudos, itens, rnc).
5. Recrie a rotina diária interna: o agendamento `backup-diario-quimicer` (pg_cron) e a função `backup-arquivos` (Edge Function) – ver o histórico do projeto.

## 3. Restaurar os PDFs
No Supabase novo, crie o bucket **laudos-mp** (privado) e envie para ele o conteúdo da pasta
`arquivos/laudos-mp` do Drive, mantendo as subpastas (pode ser pelo painel do Supabase ou com o rclone).

## 4. Publicar o app de novo
1. Descompacte `fontes_AAAA-MM-DD.zip` num repositório do GitHub e ative o GitHub Pages.
2. Atualize o `config.js` com o endereço (URL) e a chave pública (anon key) do Supabase novo.
3. Abra o site, entre com um usuário administrador e confira algumas liberações.

## Teste automático
Todo domingo a rotina de backup restaura o banco num servidor temporário e confere as contagens.
Se a restauração falhar, chega um e-mail "FALHA no backup".
