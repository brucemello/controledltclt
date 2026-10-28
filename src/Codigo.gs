/**
 * PROJETO: MONITORAMENTO E ALERTA DE SLA (CLT / DLT) JIRA - MULTITIMES & INCREMENTAL
 * v2.1
 *  - Disparos movidos para Disparos.gs (template novo, retry, paginação, pré-visualização)
 *  - Painel carregado via template (URLs centralizadas no CONFIG)
 *  - Lista de projetos/times lida da aba Mapeamento_Times
 *  - Validação de regras de SLA no servidor
 *  - Registro da data de extração por projeto (alerta de base desatualizada)
 *  - Gatilho diário configurável pelo menu/painel
 *  - Sanitização das siglas de projeto usadas na JQL
 */

const CONFIG = {
  JIRA_BASE_URL: 'https://jiraproducao.totvs.com.br',
  MANUAL_PDF_URL: 'https://drive.google.com/file/d/1cWrCxB2eHk9ZvnH4g97EV9CCJJ_p7OUC/view?usp=drive_link',
  LOOKER_URL: '', // <<< cole aqui o link do relatório no Looker Studio

  ABA_INDICADORES: 'Indicadores_CLT_DLT',
  ABA_REGRAS: 'Regras_SLA',
  ABA_MAPEAMENTO: 'Mapeamento_Times',
  ABA_LOG: 'Logs',

  MAX_RESULTS: 50,
  TOTAL_COLUNAS: 26, // A..X + Y/Z (data e hora de abertura do ticket e de criação da issue)
  HORA_ENVIO_DIARIO: 8, // hora (fuso do projeto) do envio automático

  CUSTOM_FIELDS: {
    DATA_ABERTURA_TICKET: 'customfield_10100', // Data e Hora de Abertura do Ticket
    PRIORIDADE_INTERNA: 'customfield_10010'    // Prioridade Interna
  }
};

/* ===================== MENU ===================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Jira - Controle CLT/DLT')
    .addItem('1. Configurar Token Jira', 'configurarTokenJira')
    .addItem('2. Configurar Webhook Padrão do Chat', 'configurarWebhookChat')
    .addItem('3. Testar Conexão Jira', 'testarConexaoJira')
    .addSeparator()
    .addItem('4. Painel Geral, Disparos & Regras SLA', 'abrirTelaParametros')
    .addItem('5. Pré-visualizar Alertas do Chat (sem enviar)', 'previsualizarAlertasChat')
    .addItem('6. Enviar Alertas Agora (Chat + E-mail)', 'executarEReportarChat')
    .addItem('6b. Enviar E-mail de Teste para Mim', 'enviarEmailTesteMenu')
    .addSeparator()
    .addItem('7. Ativar Envio Diário Automático', 'configurarGatilhoDiario')
    .addItem('8. Desativar Envio Diário Automático', 'removerGatilhoDiario')
    .addItem('🔎 Diagnóstico da Instalação', 'diagnosticoMenu')
    .addSeparator()
    .addItem('9. Manual do Usuário (PDF)', 'abrirManual')
    .addToUi();
}

function configurarTokenJira() {
  const token = Browser.inputBox('Informe o Personal Access Token (PAT) do Jira:');
  if (!token || token === 'cancel') return;
  PropertiesService.getScriptProperties().setProperty('JIRA_TOKEN', token.trim());
  SpreadsheetApp.getUi().alert('Token salvo com sucesso!');
}

function configurarWebhookChat() {
  const url = Browser.inputBox('Informe a URL do Webhook Padrão do Chat:');
  if (!url || url === 'cancel') return;
  PropertiesService.getScriptProperties().setProperty('WEBHOOK_CHAT_URL', url.trim());
  SpreadsheetApp.getUi().alert('Webhook padrão salvo com sucesso!');
}

function testarConexaoJira() {
  try {
    const resp = chamarJiraApi_('/rest/api/2/myself', 'get');
    registrarLog_('TESTE_CONEXAO', 'OK', JSON.stringify(resp).substring(0, 100));
    SpreadsheetApp.getUi().alert('Conexão OK!\nUsuário: ' + (resp.displayName || resp.name));
  } catch (e) {
    SpreadsheetApp.getUi().alert('Falha na Conexão: ' + e.message);
  }
}

function abrirTelaParametros() {
  garantirAbaRegras_();
  garantirAbaMapeamento_();
  const tpl = HtmlService.createTemplateFromFile('Parametros');
  tpl.urlManual = CONFIG.MANUAL_PDF_URL;
  tpl.urlLooker = CONFIG.LOOKER_URL;
  const html = tpl.evaluate().setWidth(760).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Controle de SLA CLT / DLT');
}

function abrirManual() {
  const html = HtmlService.createHtmlOutput(
    '<script>window.open(' + JSON.stringify(CONFIG.MANUAL_PDF_URL) + ', "_blank"); google.script.host.close();</script>'
  ).setWidth(250).setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(html, 'Abrindo Manual...');
}

/* ===================== DADOS DO PAINEL (google.script.run) ===================== */

function obterDadosPainel() {
  const leitura = lerMapeamento_();
  const mapa = leitura.mapa;
  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();

  const projetos = Object.keys(mapa).map(sigla => {
    const iso = props.getProperty('EXTRACAO_' + sigla);
    return {
      sigla: sigla,
      time: mapa[sigla].time,
      ultimaExtracao: iso ? Utilities.formatDate(new Date(iso), tz, 'dd/MM HH:mm') : 'nunca'
    };
  });

  const timesMap = {};
  Object.keys(mapa).forEach(sigla => {
    const t = mapa[sigla].time;
    if (!timesMap[t]) timesMap[t] = { nome: t, webhookProprio: false, emails: [] };
    if (mapa[sigla].webhook) timesMap[t].webhookProprio = true;
    mapa[sigla].emails.forEach(e => {
      if (timesMap[t].emails.indexOf(e) < 0) timesMap[t].emails.push(e);
    });
  });

  return {
    projetos: projetos,
    times: Object.keys(timesMap).sort().map(t => timesMap[t]),
    regras: obterRegrasSLA(),
    temWebhookPadrao: !!props.getProperty('WEBHOOK_CHAT_URL'),
    gatilhoAtivo: gatilhoRotinaAtivo_(),
    horaEnvio: CONFIG.HORA_ENVIO_DIARIO,
    avisosMapeamento: leitura.invalidas,
    versao: VERSAO_DISPAROS,
    usuarioEmail: emailUsuarioAtual_()
  };
}

/* ===================== GATILHO DIÁRIO ===================== */

function configurarGatilhoDiario() { alertaSeguro_(ativarEnvioDiario()); }
function removerGatilhoDiario() { alertaSeguro_(desativarEnvioDiario()); }

function ativarEnvioDiario() {
  removerGatilhosRotina_();
  ScriptApp.newTrigger('rotinaDiariaAtendimentos')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.HORA_ENVIO_DIARIO)
    .nearMinute(0)
    .create();
  const msg = 'Envio diário ativado: por volta das ' + CONFIG.HORA_ENVIO_DIARIO + 'h, apenas em dias úteis. ' +
    'Responsável pelo gatilho: ' + Session.getEffectiveUser().getEmail();
  registrarLog_('GATILHO', 'OK', msg);
  return msg;
}

function desativarEnvioDiario() {
  const n = removerGatilhosRotina_();
  const msg = n ? 'Envio diário desativado (' + n + ' gatilho(s) removido(s)).' :
    'Nenhum gatilho seu estava ativo. Gatilhos criados por outras pessoas só podem ser removidos por elas.';
  registrarLog_('GATILHO', 'OK', msg);
  return msg;
}

function removerGatilhosRotina_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'rotinaDiariaAtendimentos') {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  return n;
}

function gatilhoRotinaAtivo_() {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'rotinaDiariaAtendimentos');
}

/* ===================== ROTINA AGENDADA ===================== */

function rotinaDiariaAtendimentos() {
  if (CHAT_CFG.PULAR_FIM_DE_SEMANA && ehFimDeSemana_()) {
    registrarLog_('ROTINA_DIARIA', 'SKIP', 'Fim de semana - extração e envio ignorados.');
    return;
  }

  const listaProjetos = Object.keys(obterMapeamentoTimes_());
  if (listaProjetos.length === 0) {
    registrarLog_('ROTINA_DIARIA', 'ERRO', 'Nenhum projeto cadastrado na aba Mapeamento_Times.');
    return;
  }

  try {
    const r = processarRelatorioCLTDLT({ projeto: listaProjetos.join(', ') });
    registrarLog_('ROTINA_DIARIA', 'OK', r.mensagem);
  } catch (e) {
    // Não dispara com base antiga: falha na extração interrompe a rotina (o dono do gatilho recebe e-mail).
    registrarLog_('ROTINA_DIARIA', 'ERRO', 'Falha na extração, envio cancelado: ' + e.message);
    throw e;
  }

  const envio = dispararAlertasSLA_({ origem: 'AGENDADO', canais: CHAT_CFG.CANAIS_AGENDADO });
  registrarLog_('ROTINA_DIARIA', envio.falhas.length ? 'PARCIAL' : 'OK', envio.mensagem.substring(0, 300));
}

/* ===================== REGRAS DE SLA ===================== */

function obterRegrasSLA() {
  const padrao = {
    CLT: { Critical: 23, High: 30, Medium: 60, Low: 70 },
    DLT: { Critical: 13, High: 21, Medium: 40, Low: 'N/A' }
  };

  try {
    garantirAbaRegras_();
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA_REGRAS);
    const dados = aba.getRange(2, 1, 4, 3).getValues();

    if (!dados || dados.length === 0 || !dados[0][0]) return padrao;

    const regras = { CLT: {}, DLT: {} };
    dados.forEach(linha => {
      const prio = normalizarPrioridade_(linha[0]);
      if (prio) {
        regras.CLT[prio] = Number(linha[1]) || padrao.CLT[prio];
        const dltVal = String(linha[2]).toUpperCase().trim();
        regras.DLT[prio] = (dltVal === 'N/A' || dltVal === '' || dltVal === 'INFINITY') ? 'N/A' : (Number(linha[2]) || padrao.DLT[prio]);
      }
    });

    return regras;
  } catch (e) {
    registrarLog_('ERRO_OBTER_REGRAS', 'ERRO', e.message);
    return padrao;
  }
}

function salvarRegrasSLA(novasRegras) {
  const regras = validarRegras_(novasRegras);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = prepararAba_(ss, CONFIG.ABA_REGRAS);

  aba.clear();
  aba.appendRow(['Prioridade', 'Limite CLT (Dias)', 'Limite DLT (Dias)']);
  aba.getRange('A1:C1').setFontWeight('bold').setBackground('#1C3144').setFontColor('#FFFFFF');

  const linhas = ['Critical', 'High', 'Medium', 'Low'].map(p => [p, regras.CLT[p], regras.DLT[p]]);
  aba.getRange(2, 1, 4, 3).setValues(linhas);
  aba.autoResizeColumns(1, 3);

  const usuario = Session.getActiveUser().getEmail() || 'desconhecido';
  registrarLog_('ATUALIZAR_REGRAS', 'OK', 'Regras atualizadas por ' + usuario + ': ' + JSON.stringify(regras));
  return 'Regras de SLA atualizadas com sucesso! Elas valem a partir da próxima análise.';
}

function validarRegras_(r) {
  if (!r || !r.CLT || !r.DLT) throw new Error('Estrutura de regras inválida.');
  const out = { CLT: {}, DLT: {} };

  ['Critical', 'High', 'Medium', 'Low'].forEach(p => {
    const clt = Number(r.CLT[p]);
    if (!Number.isInteger(clt) || clt <= 0 || clt > 999) {
      throw new Error('Limite CLT inválido para ' + p + ': "' + r.CLT[p] + '". Use um número inteiro entre 1 e 999.');
    }
    out.CLT[p] = clt;

    const dltBruto = String(r.DLT[p] === undefined || r.DLT[p] === null ? '' : r.DLT[p]).trim().toUpperCase();
    if (dltBruto === 'N/A' || dltBruto === '') {
      out.DLT[p] = 'N/A';
      return;
    }
    const dlt = Number(dltBruto);
    if (!Number.isInteger(dlt) || dlt <= 0 || dlt > 999) {
      throw new Error('Limite DLT inválido para ' + p + ': "' + r.DLT[p] + '". Use um inteiro entre 1 e 999 ou N/A.');
    }
    out.DLT[p] = dlt;
  });

  return out;
}

/* ===================== EXTRAÇÃO / RELATÓRIO ===================== */

function normalizarListaProjetos_(entrada) {
  const lista = String(entrada || '')
    .split(/[,;\s]+/)
    .map(p => p.trim().toUpperCase())
    .filter(Boolean);

  const invalidos = lista.filter(p => !/^[A-Z][A-Z0-9_]*$/.test(p));
  if (invalidos.length) throw new Error('Sigla(s) de projeto inválida(s): ' + invalidos.join(', '));

  return lista.filter((p, i) => lista.indexOf(p) === i);
}

function processarRelatorioCLTDLT(parametros) {
  const mapaTimes = obterMapeamentoTimes_();
  let listaProjetosAtualizar = normalizarListaProjetos_(parametros && parametros.projeto);

  if (listaProjetosAtualizar.length === 0) {
    listaProjetosAtualizar = Object.keys(mapaTimes);
  }
  if (listaProjetosAtualizar.length === 0) {
    throw new Error('Nenhum projeto foi informado ou configurado na aba Mapeamento_Times.');
  }

  const projetoInput = listaProjetosAtualizar.join(', ');
  const regrasVigentes = obterRegrasSLA();

  // statusCategory != Done evita erro HTTP 400 de status inexistentes em alguns projetos
  const jql = `project in (${projetoInput}) AND issuetype in (Manutenção, "Rejeição - Manutenção", Apoio, "Apoio - Cliente", Documentação) AND statusCategory != Done AND status not in (Expedido, Cancelled, Cancelado, "Pronto para Expedição", Recusada, Associado, Closed, Reprovado) AND Ticket is not EMPTY AND "Código da Organização" != TOTVS AND "Data Solução Paliativa" is EMPTY AND (labels not in (FX3) OR labels is EMPTY)`;

  registrarLog_('JQL_EXECUTADA', 'INFO', jql);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = prepararAba_(ss, CONFIG.ABA_INDICADORES);
  if (aba.getLastRow() === 0) inicializarAbaIndicadores_(aba);

  let dadosExistentes = [];
  if (aba.getLastRow() > 1) {
    dadosExistentes = aba.getRange(2, 1, aba.getLastRow() - 1, CONFIG.TOTAL_COLUNAS).getValues();
  }

  const dadosPreservados = dadosExistentes.filter(linha => {
    const key = String(linha[0]);
    if (!key) return false;
    const sigla = key.split('-')[0].toUpperCase();
    return listaProjetosAtualizar.indexOf(sigla) < 0;
  });

  const issues = buscarIssuesPorJQL_(jql);
  const novasLinhas = [];
  const hoje = new Date();

  let totalDentro = 0;
  let totalProximo = 0;
  let totalEstourado = 0;

  issues.forEach((issue) => {
    const key = issue.key;
    const fields = issue.fields || {};

    const siglaProjeto = key.split('-')[0].toUpperCase();
    const infoTime = mapaTimes[siglaProjeto] || { time: 'Outros / Não Mapeado', lead: 'N/A' };
    const nomeTime = infoTime.time;

    const prioridadeInternaObj = fields[CONFIG.CUSTOM_FIELDS.PRIORIDADE_INTERNA];
    const prioridadePadraoObj = fields.priority || {};
    const nomePrioridade = (prioridadeInternaObj && prioridadeInternaObj.value) ? prioridadeInternaObj.value : (prioridadePadraoObj.name || 'Baixa');
    const prioridade = normalizarPrioridade_(nomePrioridade);

    // CLT conta da ABERTURA DO TICKET (campo customizado). Sem essa data, CLT = N/A
    //     (não usa a criação da issue como substituta, o que transformaria CLT em DLT).
    // DLT conta da CRIAÇÃO DA ISSUE no Jira (campo created).
    // Regra: dias completos de 24h desde a data/hora de abertura (mesma do painel da sprint).
    const aberturaTicket = parseJiraDataHora_(fields[CONFIG.CUSTOM_FIELDS.DATA_ABERTURA_TICKET]);
    const criacaoIssue = parseJiraDataHora_(fields.created);
    const resolucao = parseJiraDataHora_(fields.resolutiondate);
    const dataFinal = resolucao || hoje;
    const temCLT = !!aberturaTicket;

    const dataAberturaTicketDia = dataLocal_(aberturaTicket);
    const dataCriacaoIssueDia = dataLocal_(criacaoIssue);
    const dataResolucaoDia = dataLocal_(resolucao);

    const diasDecorridosCLT = temCLT ? diasCompletos_(aberturaTicket, dataFinal) : 'N/A';
    const diasDecorridosDLT = diasCompletos_(criacaoIssue, dataFinal);

    const limiteCLT = regrasVigentes.CLT[prioridade] !== undefined ? regrasVigentes.CLT[prioridade] : 70;
    const limiteDLT = regrasVigentes.DLT[prioridade] !== undefined ? regrasVigentes.DLT[prioridade] : Infinity;
    const semLimiteDLT = (limiteDLT === Infinity || limiteDLT === 'N/A');

    const diasRestantesCLT = temCLT ? limiteCLT - diasDecorridosCLT : 'N/A';
    const diasRestantesDLT = semLimiteDLT ? 999 : limiteDLT - diasDecorridosDLT;

    const alertaCLT = temCLT ? calcularAlerta_(diasRestantesCLT) : 'N/A (SEM DATA DO TICKET)';
    const alertaDLT = semLimiteDLT ? 'DENTRO DO PRAZO' : calcularAlerta_(diasRestantesDLT);

    let statusGeral = 'DENTRO DO PRAZO';
    if (alertaCLT === 'ESTOURADO' || alertaDLT === 'ESTOURADO') {
      statusGeral = 'ESTOURADO (PASSOU DO PRAZO)';
      totalEstourado++;
    } else if (alertaCLT === 'PRÓXIMO DE VENCER' || alertaDLT === 'PRÓXIMO DE VENCER') {
      statusGeral = 'PRÓXIMO DE VENCER';
      totalProximo++;
    } else {
      totalDentro++;
    }

    // Ano/Mês de referência: abertura do ticket; sem ela, criação da issue.
    const dataReferencia = dataAberturaTicketDia || dataCriacaoIssueDia;
    const ano = Number(dataReferencia.substring(0, 4));
    const anoMes = dataReferencia.substring(0, 7);

    novasLinhas.push([
      key,
      fields.summary || '',
      (fields.status && fields.status.name) || '',
      prioridade,
      (fields.assignee && fields.assignee.displayName) || 'Sem responsável',
      statusGeral,
      temCLT ? dataAberturaTicketDia : 'N/A',
      diasDecorridosCLT,
      limiteCLT,
      diasRestantesCLT,
      alertaCLT,
      dataCriacaoIssueDia,
      diasDecorridosDLT,
      semLimiteDLT ? 'N/A' : limiteDLT,
      semLimiteDLT ? 'N/A' : diasRestantesDLT,
      alertaDLT,
      dataResolucaoDia || 'Em Aberto',
      ano,
      anoMes,
      temCLT && alertaCLT !== 'ESTOURADO' ? 1 : 0,
      temCLT && alertaCLT === 'ESTOURADO' ? 1 : 0,
      alertaDLT !== 'ESTOURADO' ? 1 : 0,
      alertaDLT === 'ESTOURADO' ? 1 : 0,
      nomeTime,
      temCLT ? aberturaTicket : 'N/A',
      criacaoIssue
    ]);
  });

  const todasAsLinhas = dadosPreservados.concat(novasLinhas);

  aba.clearContents();
  inicializarAbaIndicadores_(aba);

  if (todasAsLinhas.length > 0) {
    aba.getRange(2, 1, todasAsLinhas.length, todasAsLinhas[0].length).setValues(todasAsLinhas);
    formatarLinhasResultado_(aba, todasAsLinhas.length);
  }

  // Registro da data de extração (usado no cabeçalho das mensagens e no painel)
  const props = PropertiesService.getScriptProperties();
  const agoraIso = new Date().toISOString();
  const marcas = { ULTIMA_EXTRACAO: agoraIso, ULTIMO_PROJETO_EXTRAIDO: projetoInput };
  listaProjetosAtualizar.forEach(p => { marcas['EXTRACAO_' + p] = agoraIso; });
  props.setProperties(marcas);

  registrarLog_('GERAR_RELATORIO', 'OK', `${novasLinhas.length} issues atualizadas para: ${projetoInput}. Total consolidado na planilha: ${todasAsLinhas.length}.`);

  return {
    sucesso: true,
    total: novasLinhas.length,
    estourados: totalEstourado,
    proximos: totalProximo,
    dentro: totalDentro,
    mensagem: `Atualização concluída para [${projetoInput}]! Processados: ${novasLinhas.length} ` +
      `(🔴 ${totalEstourado} | 🟡 ${totalProximo} | 🟢 ${totalDentro}) | Base consolidada: ${todasAsLinhas.length} chamados.`
  };
}

/* ===================== MAPEAMENTO / ABAS ===================== */

function obterMapeamentoTimes_() {
  return lerMapeamento_().mapa;
}

/**
 * Lê a aba Mapeamento_Times ignorando linhas cuja coluna A não seja uma sigla válida
 * (ex.: URL de webhook colada na coluna errada). As linhas ignoradas são devolvidas em "invalidas".
 */
function lerMapeamento_() {
  garantirAbaMapeamento_();
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA_MAPEAMENTO);
  const dados = aba.getRange(2, 1, Math.max(aba.getLastRow() - 1, 1), 5).getValues();
  const mapa = {};
  const invalidas = [];

  dados.forEach((linha, i) => {
    const sigla = String(linha[0] || '').trim().toUpperCase();
    if (!sigla) return;

    if (!/^[A-Z][A-Z0-9_]*$/.test(sigla)) {
      const resumo = /^HTTPS?:/.test(sigla) ? 'URL na coluna A (deveria estar na coluna D)' : '"' + sigla.substring(0, 30) + '"';
      invalidas.push('Linha ' + (i + 2) + ': ' + resumo);
      return;
    }

    const emailsBrutos = String(linha[4] || '').split(/[,;\s]+/).map(x => x.trim().toLowerCase()).filter(Boolean);
    const emails = emailsBrutos.filter(x => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
    if (emails.length < emailsBrutos.length) {
      invalidas.push('Linha ' + (i + 2) + ': e-mail inválido na coluna E (ignorado)');
    }

    mapa[sigla] = {
      time: String(linha[1] || 'Outros Squads').trim(),
      lead: linha[2] || 'Sem Lead',
      webhook: String(linha[3] || '').trim(),
      emails: emails
    };
  });

  if (invalidas.length) {
    registrarLog_('MAPEAMENTO', 'AVISO', 'Linhas ignoradas em ' + CONFIG.ABA_MAPEAMENTO + ': ' + invalidas.join(' | '));
  }

  return { mapa: mapa, invalidas: invalidas };
}

function garantirAbaMapeamento_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(CONFIG.ABA_MAPEAMENTO);
  if (!aba) {
    aba = ss.insertSheet(CONFIG.ABA_MAPEAMENTO);
    aba.appendRow(['Projeto (Sigla Jira)', 'Nome do Time / Squad', 'Tech Lead', 'Webhook Chat do Time (Opcional)', 'E-mails para Alerta (separados por vírgula)']);
    aba.getRange('A1:E1').setFontWeight('bold').setBackground('#1C3144').setFontColor('#FFFFFF');

    aba.appendRow(['DSERTAF4', 'Novas Apurações', 'Bruce Mello', '', '']);
    aba.appendRow(['DTRIPRO', 'Novas Apurações', 'Bruce Mello', '', '']);
    aba.appendRow(['DTRIPRO_CAPEX', 'Novas Apurações', 'Bruce Mello', '', '']);
    aba.appendRow(['DJOBRIG', 'Obrigações Acessórias', 'Adriana', '', '']);
    aba.autoResizeColumns(1, 5);
  } else if (!String(aba.getRange(1, 5).getValue()).trim()) {
    aba.getRange(1, 5).setValue('E-mails para Alerta (separados por vírgula)')
      .setFontWeight('bold').setBackground('#1C3144').setFontColor('#FFFFFF');
  }
}

function garantirAbaRegras_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA_REGRAS);
  if (!aba || aba.getLastRow() <= 1) {
    salvarRegrasSLA({
      CLT: { Critical: 23, High: 30, Medium: 60, Low: 70 },
      DLT: { Critical: 13, High: 21, Medium: 40, Low: 'N/A' }
    });
  }
}

/* ===================== JIRA ===================== */

function calcularAlerta_(diasRestantes) {
  if (diasRestantes < 0) return 'ESTOURADO';
  if (diasRestantes <= 5) return 'PRÓXIMO DE VENCER';
  return 'DENTRO DO PRAZO';
}

function buscarIssuesPorJQL_(jql) {
  let startAt = 0, total = 0, todasIssues = [];
  do {
    const endpoint = '/rest/api/2/search?jql=' + encodeURIComponent(jql) +
      '&startAt=' + startAt +
      '&maxResults=' + CONFIG.MAX_RESULTS +
      '&fields=' + encodeURIComponent(`summary,status,assignee,created,priority,resolutiondate,${CONFIG.CUSTOM_FIELDS.DATA_ABERTURA_TICKET},${CONFIG.CUSTOM_FIELDS.PRIORIDADE_INTERNA}`);

    const resp = chamarJiraApi_(endpoint, 'get');
    const issues = resp.issues || [];
    total = resp.total || 0;
    todasIssues = todasIssues.concat(issues);
    startAt += CONFIG.MAX_RESULTS;
  } while (startAt < total);

  return todasIssues;
}

function normalizarPrioridade_(nomePrioridade) {
  const p = String(nomePrioridade).toLowerCase();
  if (p.includes('crític') || p.includes('critic')) return 'Critical';
  if (p.includes('alt') || p.includes('high')) return 'High';
  if (p.includes('méd') || p.includes('med')) return 'Medium';
  return 'Low';
}

function chamarJiraApi_(endpoint, method, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('JIRA_TOKEN');
  if (!token) throw new Error('Token do Jira não configurado.');

  const options = {
    method: method,
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/json', 'Content-Type': 'application/json' }
  };
  if (payload) options.payload = JSON.stringify(payload);

  const response = UrlFetchApp.fetch(CONFIG.JIRA_BASE_URL + endpoint, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    registrarLog_('ERRO_API_JIRA', code, text.substring(0, 300));
    throw new Error(`Erro HTTP ${code}: ${text.substring(0, 300)}`);
  }

  return text ? JSON.parse(text) : {};
}

/* ===================== PLANILHA / LOG ===================== */

function inicializarAbaIndicadores_(aba) {
  aba.clearContents();
  aba.appendRow([
    'Issue Key', 'Resumo', 'Status', 'Prioridade', 'Responsável', 'Status Geral SLA',
    'Abertura Ticket (CLT)', 'Dias CLT', 'Limite CLT', 'Dias Restantes CLT', 'Alerta CLT',
    'Criação Issue (DLT)', 'Dias DLT', 'Limite DLT', 'Dias Restantes DLT', 'Alerta DLT', 'Data Resolução',
    'Ano', 'Ano_Mes', 'Flag_CLT_Dentro', 'Flag_CLT_Estourado', 'Flag_DLT_Dentro', 'Flag_DLT_Estourado', 'Time_Squad',
    'Abertura Ticket Data/Hora (CLT)', 'Criação Issue Data/Hora (DLT)'
  ]);
  aba.getRange(1, 1, 1, CONFIG.TOTAL_COLUNAS).setFontWeight('bold').setBackground('#1C3144').setFontColor('#FFFFFF');
}

function formatarLinhasResultado_(aba, qtdLinhas) {
  aba.autoResizeColumns(1, CONFIG.TOTAL_COLUNAS);
  aba.getRange(2, 6, qtdLinhas, 19).setHorizontalAlignment('center');
  aba.getRange(2, 25, qtdLinhas, 2).setNumberFormat('dd/MM/yyyy HH:mm');
}

function prepararAba_(ss, nome) {
  return ss.getSheetByName(nome) || ss.insertSheet(nome);
}

function registrarLog_(acao, status, detalhe) {
  const aba = prepararAba_(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.ABA_LOG);
  if (aba.getLastRow() === 0) aba.appendRow(['Data/Hora', 'Ação', 'Status', 'Detalhe']);
  aba.appendRow([new Date(), acao, status, detalhe]);
}

/**
 * Converte uma data do Jira para o dia civil 'yyyy-MM-dd' no fuso do projeto.
 * Aceita: Date, '2026-09-08' (campo só data) e '2026-09-08T14:30:00.000-0300' (data e hora).
 * Campo só data é mantido como está: passar por new Date() o leria como UTC e recuaria um dia.
 */
function dataLocal_(valor) {
  const d = parseJiraDataHora_(valor);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}

/**
 * Converte data do Jira em Date. Campo só data ('2026-09-08') vira 00:00 no fuso do projeto
 * (new Date('2026-09-08') seria meia-noite UTC = dia anterior em São Paulo).
 */
function parseJiraDataHora_(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return isNaN(valor.getTime()) ? null : valor;
  const s = String(valor).trim();
  const soData = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (soData) return new Date(Number(soData[1]), Number(soData[2]) - 1, Number(soData[3]));
  // Jira usa offset sem dois-pontos (-0300); o padrão ISO pede -03:00.
  const d = new Date(s.replace(/([+-]\d{2})(\d{2})$/, '$1:$2'));
  return isNaN(d.getTime()) ? null : d;
}

/** Dias completos de 24h entre duas datas/horas (regra do painel da sprint). */
function diasCompletos_(inicio, fim) {
  return Math.floor((fim.getTime() - inicio.getTime()) / 86400000);
}

function extrairData_(dataHora) {
  return dataHora ? String(dataHora).substring(0, 10) : '';
}
