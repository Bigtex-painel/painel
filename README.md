[README.md](https://github.com/user-attachments/files/27317767/README.md)
# Painel Financeiro Bigtex

Painel web hospedado em Cloudflare Pages com acesso restrito por Cloudflare Access.

## Estrutura

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
- **Metadados** (2 chaves): `data_ref`, `hoje`

O painel `/financeiro/` consome apenas 9 dessas (saldos + 2 metadados), o que reduz o payload em ~78%.
