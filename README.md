# Controle de SLA CLT/DLT (Jira → Google Chat / E-mail)

Apps Script vinculado à planilha de controle de SLA. Extrai o backlog do Jira, calcula CLT/DLT por prioridade e envia alertas por time no Google Chat e por e-mail.

## Estrutura

| Arquivo | Conteúdo |
|---|---|
| `src/Codigo.gs` | Menu, extração do Jira, regras de SLA, mapeamento de times, gatilho diário |
| `src/Disparos.gs` | Templates e envio (Chat e e-mail), pré-visualização, diagnóstico |
| `src/DashboardApi.gs` + `src/Dashboard.html` | Dashboard (KPIs, chamados, disparos). Nomes diferentes: o Apps Script não aceita .gs e .html com o mesmo nome |
| `src/Parametros.html` | Painel (Análise, Disparos, Regras, Looker, Manual) |
| `src/appsscript.json` | Manifesto. Na primeira publicação é copiado do projeto remoto, se não existir aqui |

## Fluxo de publicação

- **Push na `main` alterando `src/`** → workflow *Deploy Apps Script* faz `clasp push` e, se `DEPLOYMENT_ID` existir, atualiza a implantação.
- **Editou direto no editor do Apps Script?** Rode *Sincronizar do Apps Script* em Actions. Ele abre um PR com as diferenças; o merge publica de novo.

> O `clasp push` substitui **todos** os arquivos do projeto remoto pelos de `src/`. Arquivo que existir só no editor é apagado na próxima publicação. Sincronize antes se alguém editou por lá.

## Secrets (Settings → Secrets and variables → Actions)

| Secret | Valor |
|---|---|
| `CLASPRC_JSON` | Conteúdo do `~/.clasprc.json` gerado por `clasp login` |
| `SCRIPT_ID` | ID do projeto (Configurações do projeto no editor do Apps Script) |
| `DEPLOYMENT_ID` | Opcional. ID da implantação a atualizar |

## Configuração na planilha (após a publicação)

1. Recarregar a planilha e aceitar a autorização.
2. Menu **1** (token Jira) e **2** (webhook padrão), se ainda não configurados.
3. Aba `Mapeamento_Times`: sigla (A), time (B), TL (C), webhook do Chat (D), e-mails (E).
4. Menu **🔎 Diagnóstico da Instalação** — todas as linhas devem estar ✅.
