/**
 * DASHBOARD v1.0 - KPIs, chamados e disparos em tela cheia.
 * Abre como janela na planilha (menu "📊 Dashboard") e, se o projeto estiver implantado
 * como App da Web, também pela URL da implantação (doGet).
 * Depende de Código.gs e Disparos.gs.
 */

const DASH_CFG = {
  DOMINIO_PERMITIDO: 'totvs.com.br', // só e-mails deste domínio abrem o App da Web ('' = sem filtro)
  MAX_HISTORICO: 40
};

/* ===================== ENTRADAS ===================== */

function abrirDashboard() {
  const html = montarHtmlDashboard_().setWidth(1400).setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard SLA CLT / DLT');
}

function doGet() {
  if (!usuarioPermitido_()) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:40px;color:#172b4d"><h3>Acesso restrito</h3>' +
      '<p>Entre com uma conta @' + escaparHtml_(DASH_CFG.DOMINIO_PERMITIDO) + ' para ver o dashboard.</p></div>'
    ).setTitle('Acesso restrito');
  }
  return montarHtmlDashboard_()
    .setTitle('SLA CLT/DLT')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function montarHtmlDashboard_() {
  const tpl = HtmlService.createTemplateFromFile('Dashboard');
  tpl.urlLooker = CONFIG.LOOKER_URL || '';
  return tpl.evaluate();
}

function usuarioPermitido_() {
  if (!DASH_CFG.DOMINIO_PERMITIDO) return true;
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  return email.endsWith('@' + DASH_CFG.DOMINIO_PERMITIDO.toLowerCase());
}

/** Toda chamada do dashboard passa por aqui (vale também quando aberto pela URL). */
function exigirAcesso_() {
  if (!usuarioPermitido_()) throw new Error('Acesso restrito a contas @' + DASH_CFG.DOMINIO_PERMITIDO + '.');
}

/* ===================== DADOS (google.script.run) ===================== */

function obterDashboard() {
  exigirAcesso_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  const agora = new Date();
  const leitura = lerMapeamento_();
  const mapa = leitura.mapa;

  const aba = ss.getSheetByName(CONFIG.ABA_INDICADORES);
  const dados = (aba && aba.getLastRow() > 1)
    ? aba.getRange(2, 1, aba.getLastRow() - 1, CONFIG.TOTAL_COLUNAS).getValues()
    : [];
  const grupos = agruparPorTime_(dados, agora);

  const itens = [];
  Object.keys(grupos).forEach(time => {
    grupos[time].forEach(it => itens.push(serializarItem_(it, time, tz)));
  });
  itens.sort((a, b) => a.pior - b.pior);

  const webhooksTime = montarWebhooksPorTime_(mapa);
  const emailsTime = montarEmailsPorTime_(mapa);
  const webhookPadrao = !!props.getProperty('WEBHOOK_CHAT_URL');
  const nomesTimes = {};
  Object.keys(mapa).forEach(s => { nomesTimes[mapa[s].time] = true; });
  Object.keys(grupos).forEach(t => { nomesTimes[t] = true; });

  const times = Object.keys(nomesTimes).sort().map(t => {
    const doTime = itens.filter(i => i.time === t);
    const ctx = montarContextoTime_(t, mapa, props, tz);
    return {
      nome: t,
      total: doTime.length,
      est: doTime.filter(i => i.faixa === 'ESTOURADO').length,
      prox: doTime.filter(i => i.faixa === 'PROXIMO').length,
      dentro: doTime.filter(i => i.faixa === 'DENTRO').length,
      projetos: Object.keys(mapa).filter(s => mapa[s].time === t),
      webhook: webhooksTime[t] ? 'time' : (webhookPadrao ? 'padrao' : 'nenhum'),
      emails: emailsTime[t] || [],
      base: ctx.linhaBase.replace('🕒 ', ''),
      desatualizada: ctx.desatualizada
    };
  });

  const ultima = props.getProperty('ULTIMA_EXTRACAO');
  const horasBase = ultima ? (agora.getTime() - new Date(ultima).getTime()) / 3600000 : null;

  return {
    geradoEm: Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm'),
    ultimaExtracao: ultima ? Utilities.formatDate(new Date(ultima), tz, 'dd/MM/yyyy HH:mm') : null,
    baseDesatualizada: horasBase === null || horasBase > CHAT_CFG.HORAS_BASE_DESATUALIZADA,
    kpis: calcularKpis_(itens),
    porPrioridade: resumoPorPrioridade_(itens),
    times: times,
    itens: itens,
    regras: obterRegrasSLA(),
    projetos: Object.keys(mapa).map(s => {
      const iso = props.getProperty('EXTRACAO_' + s);
      return { sigla: s, time: mapa[s].time, ultimaExtracao: iso ? Utilities.formatDate(new Date(iso), tz, 'dd/MM HH:mm') : 'nunca' };
    }),
    avisosMapeamento: leitura.invalidas,
    gatilhoAtivo: gatilhoRotinaAtivo_(),
    horaEnvio: CONFIG.HORA_ENVIO_DIARIO,
    canaisAgendado: CHAT_CFG.CANAIS_AGENDADO,
    historico: obterHistoricoEnvios_(tz),
    usuario: emailUsuarioAtual_(),
    versao: VERSAO_DISPAROS,
    urlPlanilha: ss.getUrl(),
    jiraBase: CONFIG.JIRA_BASE_URL
  };
}

function serializarItem_(it, time, tz) {
  const prazo = p => ({
    semData: !!p.semData,
    semLimite: !p.semData && p.limite === null,
    dias: p.dias,
    limite: p.limite,
    rest: isFinite(p.rest) ? p.rest : null,
    venc: p.venc ? Utilities.formatDate(p.venc, tz, 'dd/MM/yy') : null
  });
  return {
    key: it.key,
    resumo: it.resumo,
    status: it.status,
    prioridade: it.prioridade,
    responsavel: it.responsavel,
    time: time,
    clt: prazo(it.clt),
    dlt: prazo(it.dlt),
    pior: isFinite(it.pior) ? it.pior : 9999,
    faixa: it.faixa
  };
}

function calcularKpis_(itens) {
  const comCLT = itens.filter(i => !i.clt.semData);
  const comDLT = itens.filter(i => !i.dlt.semLimite);
  const pct = (parte, total) => total ? Math.round((parte / total) * 1000) / 10 : null;
  const cltOk = comCLT.filter(i => i.clt.rest >= 0).length;
  const dltOk = comDLT.filter(i => i.dlt.rest >= 0).length;
  const idades = itens.map(i => i.dlt.dias).filter(d => typeof d === 'number');

  return {
    total: itens.length,
    estourados: itens.filter(i => i.faixa === 'ESTOURADO').length,
    proximos: itens.filter(i => i.faixa === 'PROXIMO').length,
    dentro: itens.filter(i => i.faixa === 'DENTRO').length,
    cltEstourados: comCLT.length - cltOk,
    dltEstourados: comDLT.length - dltOk,
    pctCLT: pct(cltOk, comCLT.length),
    pctDLT: pct(dltOk, comDLT.length),
    baseCLT: comCLT.length,
    baseDLT: comDLT.length,
    semDataTicket: itens.length - comCLT.length,
    semResponsavel: itens.filter(i => /^sem respons/i.test(i.responsavel)).length,
    idadeMediaDLT: idades.length ? Math.round(idades.reduce((a, b) => a + b, 0) / idades.length) : null,
    maiorAtraso: itens.length && itens[0].pior < 0 ? { key: itens[0].key, dias: -itens[0].pior } : null
  };
}

function resumoPorPrioridade_(itens) {
  return ['Critical', 'High', 'Medium', 'Low'].map(p => {
    const l = itens.filter(i => i.prioridade === p);
    return {
      prioridade: p,
      total: l.length,
      est: l.filter(i => i.faixa === 'ESTOURADO').length,
      prox: l.filter(i => i.faixa === 'PROXIMO').length,
      dentro: l.filter(i => i.faixa === 'DENTRO').length
    };
  });
}

function obterHistoricoEnvios_(tz) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA_LOG);
  if (!aba || aba.getLastRow() <= 1) return [];
  const qtd = Math.min(aba.getLastRow() - 1, 400);
  const linhas = aba.getRange(aba.getLastRow() - qtd + 1, 1, qtd, 4).getValues();
  const relevantes = /^(NOTIFICAR|ROTINA_DIARIA|GERAR_RELATORIO|GATILHO|ATUALIZAR_REGRAS)/;
  return linhas
    .filter(l => relevantes.test(String(l[1])))
    .reverse()
    .slice(0, DASH_CFG.MAX_HISTORICO)
    .map(l => ({
      quando: l[0] instanceof Date ? Utilities.formatDate(l[0], tz, 'dd/MM HH:mm') : String(l[0]),
      acao: String(l[1]),
      status: String(l[2]),
      detalhe: String(l[3]).substring(0, 220)
    }));
}

/* ===================== AÇÕES (protegidas) ===================== */

function dashAnalisar(projetos) {
  exigirAcesso_();
  return processarRelatorioCLTDLT({ projeto: (projetos || []).join(', ') }).mensagem;
}

function dashPrevisualizar(times, canais) {
  exigirAcesso_();
  return previsualizarAlertasPainel(times, canais);
}

function dashEnviar(times, canais) {
  exigirAcesso_();
  return enviarAlertasPainel(times, canais);
}

function dashEmailTeste(times) {
  exigirAcesso_();
  return enviarEmailTestePainel(times);
}

function dashAlternarGatilho(ativar) {
  exigirAcesso_();
  return ativar ? ativarEnvioDiario() : desativarEnvioDiario();
}

function dashSalvarRegras(regras) {
  exigirAcesso_();
  return salvarRegrasSLA(regras);
}

function dashDiagnostico() {
  exigirAcesso_();
  return diagnosticarInstalacao();
}
