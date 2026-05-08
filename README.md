# Painel Financeiro Bigtex

Painel web hospedado em Cloudflare Pages com acesso restrito por Cloudflare Access.

## Estrutura:

```
painel/
├── index.html              ← Landing pública (raiz)
├── _redirects              ← Redireciona /financeiro/ → /financeiro/semanal.html
│
├── financeiro/             ← Acesso: financeiro + gestores
│   ├── semanal.html        ← Painel semanal (operacional)
│   └── payload.js          ← Subset com 9 chaves
│
├── gestor/                 ← Acesso: só gestores/sócios
│   ├── semanal.html        ← Painel semanal completo
│   ├── mensal.html         ← Painel mensal
│   ├── anual.html          ← Painel anual
│   └── payload.js          ← Payload completo (43 chaves)
│
└── gerar_payload.py        ← Gera ambos payloads a partir dos xls
```

## Como atualizar os dados

1. Coloca os xls atualizados na pasta `/mnt/project` (ou outra, via `--xls-dir`):
   - `2025VENDAS_ANALITICO.xlsx`, `2026...VENDAS_ANALITICO.xlsx`
   - `2025CONTAS_A_PAGAR_CONSOLIDADO.xlsx`, etc
   - `2025CONTAS_A_RECEBER_CONSOLIDADO.xlsx`, etc

2. Coloca os PDFs de extrato bancário do mês na mesma pasta (Caixa, Inter, Itaú, Stone). Eles são lidos automaticamente.

3. Roda o gerador, que produz os 2 payloads (gestor completo + financeiro reduzido):

```bash
python gerar_payload.py --gerar --data 2026-MM-DD
```

4. Commita e dá push. Cloudflare Pages publica automaticamente.

```bash
git add gestor/payload.js financeiro/payload.js
git commit -m "Atualizar dados do payload"
git push
```

## Hospedagem e acesso

- **Hospedagem:** Cloudflare Pages (deploy automático a cada push no `main`)
- **Domínio atual:** `bigtex-painel.pages.dev`
- **Autenticação:** Cloudflare Access com One-time PIN por e-mail
- **Políticas de acesso:**
  - `/financeiro/*` → e-mails autorizados (financeiro + gestores)
  - `/gestor/*` → só e-mails de gestores/sócios
  - `/` (landing) → público

## Arquitetura de dados

O `payload.js` contém 43 chaves consolidadas a partir dos xls de venda e contas:

- **Anual** (5 chaves): `ano_full`, `ano_ytd`, `ano_tri`, `desp_cats_*`, `coorte_*`
- **Mensal** (12+ chaves): `meses`, `meses_disp`, `top_cli`, `top_prod`, `socios_*`, `pmr_pmp_*`, `inad`, `metas`, etc
- **Semanal** (7 chaves): `saldos`, `saldo_total`, `dev_comprando`, `rec_ab`, `rec_qt`, `pag_ab`, `pag_qt`

# Painel Financeiro Bigtex

Painel web do dashboard financeiro, hospedado em Cloudflare Pages com acesso restrito via Cloudflare Access. Atualização diária automatizada via GitHub Actions.

> 🔄 **Próxima atualização automática:** todo dia útil às **9h57 BR** (12h57 UTC).

---

## 👥 Para usuários (funcionários)

### Como o sistema funciona

Todo dia às **9h57 da manhã**, automaticamente:

1. O sistema lê os xls que estão na pasta `INBOX` do Drive
2. Atualiza as bases consolidadas (vendas, contas a pagar, contas a receber)
3. Republica o painel com os números novos
4. Move os xls processados pra pasta `LIXEIRA` (limpando o INBOX)

### Como mandar arquivos

1. Abra o Google Drive da empresa (`bigtex.server@gmail.com`)
2. Vá em `BIGTEX_DATABASE/DASHBOARDS/INBOX`
3. Arraste os xls pra dentro

**Pode mandar com qualquer nome.** O sistema identifica vendas/pagar/receber pelo CONTEÚDO do arquivo, não pelo nome. Os nomes "estranhos" do ERP (tipo `relatorio_X8472.xlsx`) funcionam normalmente.

### Quando mandar

**Antes das 9h57 BR.** Se mandar depois, o painel só vai refletir no dia seguinte.

### Notificações WhatsApp

| Mensagem recebida | O que significa | O que você deve fazer |
|----|----|----|
| 📭 **INBOX vazio** | Ninguém mandou xls antes das 9h57 | Mandar os xls do dia (vai entrar no painel só amanhã), ou disparar manualmente (peça ao mantenedor) |
| ⚠️ **Erro parcial** | Algum xls deu problema (corrompido, formato errado, etc) | Clicar no link da mensagem, ver detalhes, corrigir o xls e reenviar pro INBOX |
| ❌ **Atualização falhou** | Sistema com problema infraestrutural | Avisar o mantenedor — não tente "consertar" sozinho |

### O que **NÃO** fazer

| ❌ Não... | Por quê |
|----|----|
| Renomear arquivos na pasta `BASES` | Quebra o sistema (perde a identificação canônica) |
| Apagar arquivos da pasta `BASES` | Idem |
| Criar arquivos novos manualmente em `BASES` | Pode causar duplicatas e conflitos |
| Mover arquivos de volta da `LIXEIRA` pro `INBOX` | Vai dobrar o processamento e bagunçar |

**Regra simples:** só mexa na pasta `INBOX`. O resto, o sistema cuida.

---

## 🛠️ Para mantenedores

### Arquitetura geral

```
                     ┌─────────────────────────┐
                     │   cron-job.org (9:57)   │
                     └────────────┬────────────┘
                                  │ POST com PAT
                                  ▼
┌──────────────┐         ┌────────────────────┐         ┌──────────────────┐
│ Google Drive │◄────────┤  GitHub Actions    │────────►│      Git         │
│              │  Drive  │  (atualizar.yml)   │  commit │ payloads, status │
│  INBOX       │  API    │                    │         │  processados     │
│  BASES       │         │  drive_sync.py     │         └────────┬─────────┘
│  LIXEIRA     │         │  gerar_payload.py  │                  │
└──────────────┘         │  notify.py         │                  ▼
                         └─────────┬──────────┘         ┌──────────────────┐
                                   │ WhatsApp           │ Cloudflare Pages │
                                   ▼ (em caso de erro)  │  (deploy auto)   │
                         ┌──────────────────┐           └──────────────────┘
                         │   CallMeBot      │
                         └──────────────────┘
```

### Componentes externos

| Componente | Onde | Pra quê |
|----|----|----|
| Google Drive | Conta `bigtex.server@gmail.com` | Armazena INBOX/BASES/LIXEIRA |
| Service Account | Google Cloud, projeto `bigtex-dashboard` | Auth da API do Drive |
| cron-job.org | Conta `bigtex.server@gmail.com` | Dispara o workflow pontualmente |
| GitHub Actions | Repo atual | Roda a atualização |
| CallMeBot | Serviço gratuito de WhatsApp via HTTP | Envia notificações |
| Cloudflare Pages | Conta `bigtex.server@gmail.com` | Hospeda o painel |

### Estrutura de pastas no Drive

```
BIGTEX_DATABASE/
└── DASHBOARDS/
    ├── INBOX/      ← funcionários colocam xls aqui (qualquer nome)
    ├── BASES/      ← arquivos canônicos (NÃO MEXER MANUALMENTE)
    └── LIXEIRA/    ← xls processados (esvaziar manualmente quando lotar)
```

A pasta `DASHBOARDS` é compartilhada com a Service Account como **Editor**. Permissão herda pras filhas.

### Estrutura do repo

```
.github/workflows/
  └── atualizar.yml             # workflow do GitHub Actions

gerar_payload.py                # gera os payloads JS do dashboard
drive_sync.py                   # auth + sync com Google Drive
notify.py                       # envia WhatsApp via CallMeBot
requirements-actions.txt        # deps Python pro workflow

gestor/payload.js               # dados do dashboard "gestor" (gerado)
financeiro/payload.js           # dados do dashboard "financeiro" (gerado)

status/<data>.json              # registro JSON de cada execução
processados/<data>/             # cópia dos xls processados (auditoria)
index.html                      # painel (servido pelo Cloudflare Pages)
_redirects                      # config do Cloudflare Pages
```

### Secrets do GitHub

Em **Settings → Secrets and variables → Actions**:

| Nome | Pra quê |
|----|----|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | JSON inteiro da Service Account (auth Drive API) |
| `DRIVE_INBOX_ID` | ID da pasta INBOX no Drive |
| `DRIVE_BASES_ID` | ID da pasta BASES no Drive |
| `DRIVE_LIXEIRA_ID` | ID da pasta LIXEIRA no Drive |
| `WHATSAPP_PHONE_1` | Telefone destino WhatsApp 1 (com DDI, formato `+5511...`) |
| `WHATSAPP_APIKEY_1` | Apikey CallMeBot do telefone 1 |
| `WHATSAPP_PHONE_2` | (opcional) telefone destino 2 |
| `WHATSAPP_APIKEY_2` | (opcional) apikey do telefone 2 |

### Convenções de nomes canônicos em BASES

| Tipo | Padrão | Exemplo |
|----|----|----|
| Vendas | `vendas_<ano>.xlsx` | `vendas_2026.xlsx` |
| Contas a pagar | `pagar_<ano>.xlsx` | `pagar_2026.xlsx` |
| Contas a receber | `receber_<ano>.xlsx` | `receber_2026.xlsx` |
| Extratos bancários | `extrato_<banco>.pdf` | `extrato_itau.pdf` |

Bancos suportados: `itau`, `inter`, `stone`, `caixa` (caixas baixas).

O `gerar_payload.py` identifica os xls do INBOX pelo CONTEÚDO (colunas distintivas) e os PDFs por conteúdo (parsing das primeiras linhas), e renomeia pra esses padrões na hora de mover pra BASES.

### Categorização do status JSON

`gerar_payload.py` grava `status/<data>.json` em toda execução. O workflow lê esse arquivo e categoriza:

| Categoria | Quando | O que acontece |
|----|----|----|
| `sucesso` | Tudo OK | Commit + WhatsApp silencioso |
| `parcial` | Algum xls com erro, mas demais OK | Commit + WhatsApp ⚠️ |
| `falha_total` | Nenhum payload gerado | Sem commit + WhatsApp ❌ |
| `catastrofica` | `status.json` nem foi escrito | Sem commit + WhatsApp ❌ |

### Procedimentos de manutenção

#### Disparar workflow manualmente

- **Via GitHub:** Aba Actions → "Atualizar Dashboard Bigtex" → "Run workflow"
- **Via cron-job.org:** clica no cronjob → "Test run"

Em ambos os casos, o trigger fica registrado como `workflow_dispatch` na aba Actions.

#### Rotacionar PAT do cron-job.org

⚠️ **O PAT atual vence em maio/2027.** O cron-job.org vai começar a falhar com 401 a partir dessa data.

1. Vá em https://github.com/settings/personal-access-tokens
2. Encontre o token `cron-job-bigtex-dispatch`
3. Clica em **Regenerate token**, define nova validade (recomendo 1 ano)
4. Copia o novo token
5. No cron-job.org, edita o cronjob, vai na aba **Advanced**
6. Substitui o header `Authorization` pelo novo: `Bearer <novo-PAT>`
7. Faz **Test run** pra validar (esperado: 204 No Content)

#### Adicionar segundo telefone WhatsApp

1. No celular do destinatário:
   - Salvar contato `+34 644 51 95 23` como "CallMeBot"
   - Mandar mensagem **literal**: `I allow callmebot to send me messages`
   - Aguardar resposta com apikey (até 2 minutos)
2. No GitHub, **Settings → Secrets and variables → Actions**:
   - Cria secret `WHATSAPP_PHONE_2` com o telefone (formato `+55119...`)
   - Cria secret `WHATSAPP_APIKEY_2` com a apikey recebida
3. Próxima execução já notifica os 2 telefones

Não precisa mudar código.

#### Virada de ano: criar stubs canônicos novos

A Service Account não pode CRIAR arquivos no Drive pessoal (sem Workspace). Quando o ano virar e começarem a chegar xls de vendas/pagar/receber do novo ano, o `upload-bases` vai logar "FALTA NO DRIVE" e mandar WhatsApp ⚠️.

**Procedimento (~3 minutos, 1x por ano):**

1. Em janeiro do ano novo, antes de mandar o primeiro xls do ano:
2. Cria localmente 3 arquivos vazios chamados `vendas_<ano>.xlsx`, `pagar_<ano>.xlsx`, `receber_<ano>.xlsx` (pode ser planilha em branco mesmo, qualquer xlsx serve)
3. Faz upload manual em `BASES/` no Drive (você é o dono, pode criar)
4. Pronto, o sistema vai escrever neles na próxima execução

#### Esvaziar a LIXEIRA periodicamente

A pasta `LIXEIRA` acumula os xls processados (com o nome aleatório do ERP). Cópia auditável fica no Git em `processados/<data>/`, então a LIXEIRA é descartável.

Periodicamente (mensal/trimestral), abra a pasta no Drive, seleciona tudo, deleta. Sem perda.

### Troubleshooting

#### Workflow não roda no horário esperado

1. Cron-job.org tá ativo? Login → veja se o job tá com status verde
2. Last execution do cronjob mostra erro? Códigos comuns:
   - **401** → PAT expirado ou inválido
   - **404** → URL errada ou repo movido (precisa atualizar URL)
   - **307/301** → Redirect detectado, clicar em "Update job URL" no cron-job.org

#### Notificação WhatsApp não chega

1. Telefone fora de área / sem internet?
2. Apikey ainda válido? Pra testar isolado: rode `python notify.py "teste"` localmente com `WHATSAPP_PHONE_1` + `WHATSAPP_APIKEY_1` no env
3. CallMeBot às vezes desativa apikeys por inatividade. Pra reativar: manda `I allow callmebot to send me messages` de novo na conversa do CallMeBot

#### "FALTA NO DRIVE: vendas_<ano>.xlsx"

Ver "Virada de ano" acima.

#### Saldo bancário desatualizado no painel

PDF de extrato em `BASES` está velho. Pra atualizar: coloca o PDF novo no `INBOX` (sistema identifica por conteúdo), aguarda próximo cron OU dispara manualmente.

#### Workflow rodou mas commit falhou

Verifique se há **branch protection rules** em `main` exigindo PR. Workflow usa `GITHUB_TOKEN` com `contents: write`, que normalmente basta — mas pode falhar se há protection. Solução: adicionar exceção pro `github-actions[bot]`.

### Decisões arquiteturais (por que cada coisa)

#### Por que Service Account em Drive pessoal (não Workspace)?

Pra evitar custo de licença Workspace (~R$36/mês). Limitação: SAs em Drive pessoal não têm cota de armazenamento, logo **não podem CRIAR arquivos**. Trade-offs:

- `BASES` é pré-populado pelo proprietário (uma vez no setup)
- `upload-bases` só faz UPDATE (otimizado com checagem MD5)
- Pasta `LIXEIRA` foi criada pelo proprietário como destino dos arquivos do INBOX após processamento (em vez de mover pra lixeira do Drive — operação que SA também não pode fazer em arquivos do dono)

Se um dia migrar pra Workspace + Shared Drive, dá pra simplificar (SA cria livremente, sem precisar de pré-população nem da LIXEIRA).

#### Por que cron externo (cron-job.org) e não cron interno do GitHub Actions?

Cron interno do GitHub free tier atrasa de 1 a 6 horas em horário comercial dos EUA. cron-job.org dispara via API com pontualidade ±1min. Custo zero, conta gratuita basta.

#### Por que MD5 check no upload-bases?

Sem isso, todo dia o sistema fazia round-trip de upload dos 19 arquivos (mesmo dos não-modificados). Reduz uploads desnecessários e mantém data de modificação no Drive realista (só atualiza o que realmente mudou).

#### Por que `processados/` no Git em vez de pasta no Drive?

Auditoria via Git é superior: cada execução é um commit datado, com diff visível. Acessível pelo histórico do GitHub para sempre. Drive viraria backup secundário sem propósito claro.
- **Metadados** (2 chaves): `data_ref`, `hoje`

O painel `/financeiro/` consome apenas 9 dessas (saldos + 2 metadados), o que reduz o payload em ~78%.
