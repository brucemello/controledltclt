/**
 * MÓDULO DE DISPAROS v2.2 - Alertas de SLA (CLT/DLT) no Google Chat e por E-mail
 * Depende de Código.gs: CONFIG, lerMapeamento_, calcularAlerta_, registrarLog_
 */

const VERSAO_DISPAROS = '2.2';

const CHAT_CFG = {
  LIMITE_CHARS: 3800,            // margem abaixo do limite de texto do Google Chat
  MAX_ITENS_POR_GRUPO: 15,       // acima disso, mostra "...e mais N"
  LISTAR_DENTRO_PRAZO: false,    // false = só a contagem dos verdes
  ENVIAR_QUANDO_TUDO_OK: true,   // envia uma linha curta se o time não tiver pendências
  PAUSA_ENTRE_ENVIOS_MS: 1100,   // webhooks do Chat aceitam ~1 msg/s por espaço
  TENTATIVAS: 3,
  PULAR_FIM_DE_SEMANA: true,     // vale para a rotina agendada
  HORAS_BASE_DESATUALIZADA: 26,
  CANAIS_AGENDADO: { chat: true, email: true } // canais usados pela rotina diária
};

const EMAIL_CFG = {
  NOME_REMETENTE: 'Alerta SLA CLT/DLT',
  PREFIXO_ASSUNTO: '[SLA CLT/DLT]',
  LISTAR_DENTRO_PRAZO: true,     // no e-mail cabe a lista completa
  MAX_ITENS_POR_GRUPO: 100,
  ENVIAR_QUANDO_TUDO_OK: false   // não manda e-mail se o time não tiver pendências
};

const STATUS_ENCERRADOS_ = ['CONCLUÍDO', 'CONCLUIDO', 'RESOLVIDO', 'EXPEDIDO', 'CANCELADO',
  'CANCELLED', 'CLOSED', 'DONE', 'RECUSADA', 'REPROVADO', 'ASSOCIADO'];

/* ===================== ENTRADAS - MENU ===================== */

function executarEReportarChat() {
  alertaSeguro_(dispararAlertasSLA_({ origem: 'MANUAL', canais: { chat: true, email: true } }).mensagem);
}

function enviarEmailTesteMenu() {
  try {
    alertaSeguro_(enviarEmailTestePainel([]));
  } catch (e) {
    alertaSeguro_('Erro: ' + e.message);
  }
}

function diagnosticoMenu() {
  alertaSeguro_(diagnosticarInstalacao());
}

function previsualizarAlertasChat() {
  const r = dispararAlertasSLA_({ origem: 'MANUAL', dryRun: true, canais: { chat: true, email: false } });
  const corpo = r.previews.length ? r.previews.filter(p => p.chat).map(p =>
    '<h4 style="margin:12px 0 4px">' + escaparHtml_(p.time) + ' → ' + escaparHtml_(p.chat.destino) + '</h4>' +
    p.chat.mensagens.map((m, i) =>
      '<pre style="white-space:pre-wrap;background:#f4f6f9;padding:8px;border-radius:4px;font-size:12px">' +
      '[Parte ' + (i + 1) + '/' + p.chat.mensagens.length + ' · ' + m.length + ' chars]\n' + escaparHtml_(m) + '</pre>'
    ).join('')
  ).join('') : '<p>' + escaparHtml_(r.mensagem) + '</p>';

  const html = HtmlService.createHtmlOutput('<div style="font-family:Arial">' + corpo + '</div>')
    .setWidth(700).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Pré-visualização dos Alertas (Chat)');
}

/* ===================== ENTRADAS - PAINEL (google.script.run) ===================== */

function previsualizarAlertasPainel(times, canais) {
  const r = dispararAlertasSLA_({ origem: 'PAINEL', dryRun: true, times: times || [], canais: normalizarCanais_(canais) });
  return { previews: r.previews, mensagem: r.mensagem };
}

function enviarAlertasPainel(times, canais) {
  return dispararAlertasSLA_({ origem: 'PAINEL', times: times || [], canais: normalizarCanais_(canais) }).mensagem;
}

/** Envia o e-mail de TODOS os times selecionados apenas para quem está usando o painel. */
function enviarEmailTestePainel(times) {
  const eu = emailUsuarioAtual_();
  if (!eu) throw new Error('Não foi possível identificar seu e-mail (conta fora do domínio?).');
  return dispararAlertasSLA_({
    origem: 'TESTE', times: times || [], canais: { chat: false, email: true }, emailTeste: eu
  }).mensagem;
}

function normalizarCanais_(c) {
  if (!c) return { chat: true, email: false };
  const out = { chat: !!c.chat, email: !!c.email };
  if (!out.chat && !out.email) throw new Error('Selecione ao menos um canal (Chat ou E-mail).');
  return out;
}

function emailUsuarioAtual_() {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
}

/* ===================== DIAGNÓSTICO ===================== */

/** Mostra qual versão das funções está de fato ativa (detecta código antigo duplicado). */
function diagnosticarInstalacao() {
  const props = PropertiesService.getScriptProperties();
  const linhas = ['Módulo de disparos: v' + VERSAO_DISPAROS];

  const envioNovo = String(executarEReportarChat).indexOf('dispararAlertasSLA_') >= 0;
  linhas.push(envioNovo ? '✅ Envio manual: versão nova'
    : '❌ Envio manual: VERSÃO ANTIGA ativa. Apague a função executarEReportarChat do Código.gs.');

  const rotinaNova = String(rotinaDiariaAtendimentos).indexOf('dispararAlertasSLA_') >= 0;
  linhas.push(rotinaNova ? '✅ Rotina diária: versão nova'
    : '❌ Rotina diária: VERSÃO ANTIGA ativa. Substitua o Código.gs inteiro.');

  const gatilhos = ScriptApp.getProjectTriggers()
    .map(t => t.getHandlerFunction())
    .filter(h => h === 'rotinaDiariaAtendimentos' || h === 'executarEReportarChat');
  linhas.push('Seus gatilhos: ' + (gatilhos.length ? gatilhos.join(', ') : 'nenhum') +
    ' (gatilhos de outras pessoas ou de cópias da planilha não aparecem aqui)');

  linhas.push('Webhook padrão: ' + (props.getProperty('WEBHOOK_CHAT_URL') ? 'configurado' : 'NÃO configurado'));
  linhas.push('Seu e-mail: ' + (emailUsuarioAtual_() || 'não identificado'));
  linhas.push('Cota de e-mail restante hoje: ' + MailApp.getRemainingDailyQuota());
  linhas.push('Fuso do projeto: ' + Session.getScriptTimeZone());

  const leitura = lerMapeamento_();
  const comEmail = Object.keys(leitura.mapa).filter(s => leitura.mapa[s].emails.length).length;
  linhas.push('Projetos mapeados: ' + Object.keys(leitura.mapa).length + ' (' + comEmail + ' com e-mail)' +
    (leitura.invalidas.length ? ' | problemas: ' + leitura.invalidas.length : ''));

  const ultima = props.getProperty('ULTIMA_EXTRACAO');
  linhas.push('Última extração: ' + (ultima ? Utilities.formatDate(new Date(ultima), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : 'nunca (na versão nova)'));

  return linhas.join('\n');
}

/* ===================== ORQUESTRAÇÃO ===================== */

function dispararAlertasSLA_(opcoes) {
  opcoes = opcoes || {};
  const canais = opcoes.canais || { chat: true, email: false };
  const res = { enviados: [], ignorados: [], falhas: [], previews: [], mensagem: '' };
  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  const agora = new Date();
  const ctx = {
    props: props,
    tz: tz,
    hojeStr: Utilities.formatDate(agora, tz, 'yyyy-MM-dd'),
    agendado: opcoes.origem === 'AGENDADO',
    dryRun: !!opcoes.dryRun,
    emailTeste: opcoes.emailTeste || '',
    res: res
  };

  if (ctx.agendado && CHAT_CFG.PULAR_FIM_DE_SEMANA && ehFimDeSemana_()) {
    res.mensagem = 'Fim de semana - envio agendado ignorado.';
    registrarLog_('NOTIFICAR', 'SKIP', res.mensagem);
    return res;
  }

  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA_INDICADORES);
  if (!aba || aba.getLastRow() <= 1) {
    res.mensagem = 'A aba de indicadores está vazia. Execute primeiro a análise de SLA.';
    return res;
  }

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 24).getValues();
  const hoje0 = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());
  const grupos = agruparPorTime_(dados, hoje0);

  let times = Object.keys(grupos).sort();
  if (opcoes.times && opcoes.times.length) {
    times = times.filter(t => opcoes.times.indexOf(t) >= 0);
  }

  if (times.length === 0) {
    res.mensagem = 'Nenhuma issue em atendimento ativo encontrada para os times selecionados.';
    return res;
  }

  const mapa = lerMapeamento_().mapa;
  const webhookPadrao = props.getProperty('WEBHOOK_CHAT_URL');
  const webhooksTime = montarWebhooksPorTime_(mapa);
  const emailsTime = montarEmailsPorTime_(mapa);

  times.forEach(time => {
    const itens = grupos[time];
    const contexto = montarContextoTime_(time, mapa, props, tz);
    const prev = { time: time, chat: null, email: null };

    if (canais.chat) processarChat_(time, itens, contexto, ctx, prev, webhooksTime[time], webhookPadrao);
    if (canais.email) processarEmail_(time, itens, contexto, ctx, prev, emailsTime[time] || []);

    if (ctx.dryRun && (prev.chat || prev.email)) res.previews.push(prev);
  });

  res.mensagem = montarResumo_(res, ctx.dryRun);
  return res;
}

function processarChat_(time, itens, contexto, ctx, prev, webhookTime, webhookPadrao) {
  const res = ctx.res;
  const mensagens = montarMensagensTime_(time, itens, contexto, ctx.tz);
  if (!mensagens.length) {
    res.ignorados.push(time + ' / Chat (sem pendências)');
    return;
  }

  const webhook = webhookTime || webhookPadrao;
  if (ctx.dryRun) {
    prev.chat = {
      destino: webhookTime ? 'Webhook do time' : (webhookPadrao ? 'Webhook padrão' : 'SEM WEBHOOK - não será enviado'),
      mensagens: mensagens
    };
    return;
  }

  if (!webhook) {
    res.ignorados.push(time + ' / Chat (sem webhook)');
    registrarLog_('NOTIFICAR_CHAT', 'SEM_WEBHOOK', 'Time sem webhook: ' + time);
    return;
  }

  const chave = 'ENVIO_CHAT_' + slug_(time);
  if (ctx.agendado && ctx.props.getProperty(chave) === ctx.hojeStr) {
    res.ignorados.push(time + ' / Chat (já enviado hoje)');
    return;
  }

  try {
    const threadKey = 'sla-' + slug_(time) + '-' + ctx.hojeStr;
    mensagens.forEach((texto, i) => {
      if (i > 0) Utilities.sleep(CHAT_CFG.PAUSA_ENTRE_ENVIOS_MS);
      enviarChatComRetry_(webhook, texto, threadKey);
    });
    ctx.props.setProperty(chave, ctx.hojeStr);
    res.enviados.push(time + ' (Chat)');
    registrarLog_('NOTIFICAR_CHAT', 'OK', time + ': ' + itens.length + ' chamados em ' + mensagens.length + ' mensagem(ns).');
  } catch (e) {
    res.falhas.push(time + ' (Chat): ' + e.message);
    registrarLog_('NOTIFICAR_CHAT', 'ERRO', time + ': ' + e.message);
  }
  Utilities.sleep(CHAT_CFG.PAUSA_ENTRE_ENVIOS_MS);
}

function processarEmail_(time, itens, contexto, ctx, prev, emailsDoTime) {
  const res = ctx.res;
  const teste = !!ctx.emailTeste;
  const email = montarEmailTime_(time, itens, contexto, ctx.tz, teste);
  if (!email) {
    res.ignorados.push(time + ' / E-mail (sem pendências)');
    return;
  }

  const destinatarios = teste ? [ctx.emailTeste] : emailsDoTime;

  if (ctx.dryRun) {
    prev.email = {
      destino: destinatarios.length ? destinatarios.join(', ') : 'SEM E-MAIL - não será enviado',
      assunto: email.assunto,
      html: email.html
    };
    return;
  }

  if (!destinatarios.length) {
    res.ignorados.push(time + ' / E-mail (sem destinatário na coluna E)');
    registrarLog_('NOTIFICAR_EMAIL', 'SEM_EMAIL', 'Time sem e-mail: ' + time);
    return;
  }

  const chave = 'ENVIO_EMAIL_' + slug_(time);
  if (ctx.agendado && ctx.props.getProperty(chave) === ctx.hojeStr) {
    res.ignorados.push(time + ' / E-mail (já enviado hoje)');
    return;
  }

  try {
    const cota = MailApp.getRemainingDailyQuota();
    if (cota < destinatarios.length) throw new Error('Cota diária de e-mail insuficiente (restam ' + cota + ').');

    MailApp.sendEmail({
      to: destinatarios.join(','),
      subject: email.assunto,
      htmlBody: email.html,
      body: email.texto,
      name: EMAIL_CFG.NOME_REMETENTE
    });

    if (!teste) ctx.props.setProperty(chave, ctx.hojeStr);
    res.enviados.push(time + ' (E-mail → ' + destinatarios.join(', ') + ')');
    registrarLog_('NOTIFICAR_EMAIL', teste ? 'TESTE' : 'OK', time + ' → ' + destinatarios.join(', '));
  } catch (e) {
    res.falhas.push(time + ' (E-mail): ' + e.message);
    registrarLog_('NOTIFICAR_EMAIL', 'ERRO', time + ': ' + e.message);
  }
}

function montarResumo_(res, dryRun) {
  if (dryRun) {
    let s = res.previews.length + ' time(s) com conteúdo para enviar.';
    if (res.ignorados.length) s += '\nSem envio: ' + res.ignorados.join(', ');
    return s;
  }
  let s = 'Enviados: ' + (res.enviados.length ? res.enviados.join(', ') : 'nenhum');
  if (res.ignorados.length) s += '\nIgnorados: ' + res.ignorados.join(', ');
  if (res.falhas.length) s += '\nFALHAS:\n' + res.falhas.join('\n');
  return s;
}

/* ===================== DADOS ===================== */

function agruparPorTime_(dados, hoje0) {
  const grupos = {};
  dados.forEach(linha => {
    const status = String(linha[2]).trim().toUpperCase();
    if (!linha[0] || !status) return;
    if (STATUS_ENCERRADOS_.some(s => status.includes(s))) return;

    const time = String(linha[23] || 'Outros / Não Mapeado').trim();
    (grupos[time] = grupos[time] || []).push(avaliarLinha_(linha, hoje0));
  });
  return grupos;
}

/** Recalcula os dias NO MOMENTO DO ENVIO (dias corridos de calendário). */
function avaliarLinha_(linha, hoje0) {
  const abertura = parseDataPlanilha_(linha[6]);
  const criacao = parseDataPlanilha_(linha[11]);

  const limCLT = Number(linha[8]);
  const limDLTnum = Number(linha[13]);
  const temDLT = String(linha[13]).toUpperCase() !== 'N/A' && linha[13] !== '' && !isNaN(limDLTnum);

  const diasCLT = abertura ? diasEntre_(abertura, hoje0) : Number(linha[7]);
  const diasDLT = criacao ? diasEntre_(criacao, hoje0) : Number(linha[12]);

  const restCLT = limCLT - diasCLT;
  const restDLT = temDLT ? limDLTnum - diasDLT : Infinity;
  const pior = Math.min(restCLT, restDLT);
  const alerta = calcularAlerta_(pior);

  return {
    key: String(linha[0]),
    resumo: String(linha[1] || ''),
    status: String(linha[2] || ''),
    prioridade: String(linha[3] || ''),
    responsavel: String(linha[4] || 'Sem responsável'),
    clt: { dias: diasCLT, limite: limCLT, rest: restCLT, venc: abertura ? somarDias_(abertura, limCLT) : null },
    dlt: { dias: diasDLT, limite: temDLT ? limDLTnum : null, rest: restDLT, venc: (criacao && temDLT) ? somarDias_(criacao, limDLTnum) : null },
    pior: pior,
    faixa: alerta === 'ESTOURADO' ? 'ESTOURADO' : (alerta === 'PRÓXIMO DE VENCER' ? 'PROXIMO' : 'DENTRO')
  };
}

function separarFaixas_(itens) {
  return {
    est: itens.filter(i => i.faixa === 'ESTOURADO').sort((a, b) => a.pior - b.pior),
    prox: itens.filter(i => i.faixa === 'PROXIMO').sort((a, b) => a.pior - b.pior),
    dentro: itens.filter(i => i.faixa === 'DENTRO').sort((a, b) => a.pior - b.pior)
  };
}

function montarWebhooksPorTime_(mapa) {
  const out = {};
  Object.keys(mapa).forEach(proj => {
    const t = mapa[proj].time;
    const w = String(mapa[proj].webhook || '').trim();
    if (w && !out[t]) out[t] = w;
  });
  return out;
}

function montarEmailsPorTime_(mapa) {
  const out = {};
  Object.keys(mapa).forEach(proj => {
    const t = mapa[proj].time;
    (mapa[proj].emails || []).forEach(e => {
      out[t] = out[t] || [];
      if (out[t].indexOf(e) < 0) out[t].push(e);
    });
  });
  return out;
}

/** Usa a extração MAIS ANTIGA entre os projetos do time (pior caso). */
function montarContextoTime_(time, mapa, props, tz) {
  let maisAntiga = null;
  Object.keys(mapa).filter(s => mapa[s].time === time).forEach(s => {
    const iso = props.getProperty('EXTRACAO_' + s);
    if (!iso) return;
    const d = new Date(iso);
    if (!maisAntiga || d < maisAntiga) maisAntiga = d;
  });

  if (!maisAntiga) {
    const isoGlobal = props.getProperty('ULTIMA_EXTRACAO');
    if (isoGlobal) maisAntiga = new Date(isoGlobal);
  }
  if (!maisAntiga) return { linhaBase: '🕒 Data da extração desconhecida', desatualizada: true };

  const horas = (Date.now() - maisAntiga.getTime()) / 3600000;
  return {
    linhaBase: '🕒 Base Jira de ' + Utilities.formatDate(maisAntiga, tz, 'dd/MM HH:mm'),
    desatualizada: horas > CHAT_CFG.HORAS_BASE_DESATUALIZADA
  };
}

/* ===================== TEMPLATE - CHAT ===================== */

function montarMensagensTime_(time, itens, contexto, tz) {
  const f = separarFaixas_(itens);
  const est = f.est, prox = f.prox, dentro = f.dentro;

  if (est.length === 0 && prox.length === 0 && !CHAT_CFG.ENVIAR_QUANDO_TUDO_OK) return [];

  const hojeFmt = Utilities.formatDate(new Date(), tz, 'dd/MM/yy');
  const cab = [
    '📊 *SLA CLT/DLT - ' + limparTexto_(time, 50).toUpperCase() + '* · ' + hojeFmt,
    '🔴 ' + est.length + ' estourados   🟡 ' + prox.length + ' a vencer   🟢 ' + dentro.length + ' no prazo',
    contexto.linhaBase + (contexto.desatualizada ? '  ⚠️ *base desatualizada - rode a análise*' : '')
  ].join('\n');

  if (est.length === 0 && prox.length === 0) {
    return [cab + '\n\n✅ Nenhum chamado vencido ou próximo de vencer.' + rodape_()];
  }

  const blocos = [];
  adicionarGrupo_(blocos, '🔴 *PASSARAM DO PRAZO (' + est.length + ')*', est, tz);
  adicionarGrupo_(blocos, '🟡 *PRÓXIMOS DE VENCER (' + prox.length + ')*', prox, tz);

  if (dentro.length) {
    if (CHAT_CFG.LISTAR_DENTRO_PRAZO) {
      adicionarGrupo_(blocos, '🟢 *DENTRO DO PRAZO (' + dentro.length + ')*', dentro, tz);
    } else {
      blocos.push('🟢 *Dentro do prazo:* ' + dentro.length + ' chamado(s) - detalhes na planilha/dashboard.');
    }
  }

  const msgs = paginar_(cab, blocos);
  msgs[msgs.length - 1] += rodape_();
  return msgs;
}

function adicionarGrupo_(blocos, titulo, itens, tz) {
  if (!itens.length) return;
  blocos.push(titulo);
  itens.slice(0, CHAT_CFG.MAX_ITENS_POR_GRUPO).forEach(i => blocos.push(formatarItem_(i, tz)));
  if (itens.length > CHAT_CFG.MAX_ITENS_POR_GRUPO) {
    blocos.push('_…e mais ' + (itens.length - CHAT_CFG.MAX_ITENS_POR_GRUPO) + ' chamado(s) neste grupo._');
  }
}

function formatarItem_(it, tz) {
  const link = '<' + CONFIG.JIRA_BASE_URL + '/browse/' + it.key + '|' + it.key + '>';
  return '• ' + link + ' · *' + it.prioridade + '* · ' + limparTexto_(it.status, 30) + '\n' +
    '    ' + limparTexto_(it.resumo, 70) + '\n' +
    '    👤 ' + limparTexto_(it.responsavel, 40) + '\n' +
    '    ' + descreverPrazo_('CLT', it.clt, tz) + '  |  ' + descreverPrazo_('DLT', it.dlt, tz);
}

function descreverPrazo_(rotulo, p, tz) {
  if (p.limite === null) return rotulo + ' ' + p.dias + 'd (sem limite)';
  const venc = p.venc ? Utilities.formatDate(p.venc, tz, 'dd/MM') : '?';
  const base = rotulo + ' ' + p.dias + '/' + p.limite + 'd';
  if (p.rest < 0) return '⛔ *' + base + ' - venceu há ' + Math.abs(p.rest) + 'd* (' + venc + ')';
  if (p.rest <= 5) return '⚠️ *' + base + ' - vence em ' + p.rest + 'd* (' + venc + ')';
  return base + ' (vence ' + venc + ')';
}

function rodape_() {
  return CONFIG.LOOKER_URL ? '\n\n📈 <' + CONFIG.LOOKER_URL + '|Abrir dashboard no Looker>' : '';
}

function paginar_(cabecalho, blocos) {
  const msgs = [];
  let atual = cabecalho;
  blocos.forEach(b => {
    if ((atual + '\n\n' + b).length > CHAT_CFG.LIMITE_CHARS) {
      msgs.push(atual);
      atual = '_(continuação)_';
    }
    atual += '\n\n' + b;
  });
  msgs.push(atual);
  if (msgs.length > 1) return msgs.map((m, i) => m + '\n_[' + (i + 1) + '/' + msgs.length + ']_');
  return msgs;
}

/* ===================== TEMPLATE - E-MAIL ===================== */

function montarEmailTime_(time, itens, contexto, tz, teste) {
  const f = separarFaixas_(itens);
  const est = f.est, prox = f.prox, dentro = f.dentro;

  if (est.length === 0 && prox.length === 0 && !EMAIL_CFG.ENVIAR_QUANDO_TUDO_OK && !teste) return null;

  const e = escaparHtml_;
  const hojeFmt = Utilities.formatDate(new Date(), tz, 'dd/MM/yy');
  const situacao = est.length ? est.length + ' estourado(s)' : (prox.length ? prox.length + ' a vencer' : 'tudo no prazo');
  const assunto = (teste ? '[TESTE] ' : '') + EMAIL_CFG.PREFIXO_ASSUNTO + ' ' + time + ' · ' + situacao + ' · ' + hojeFmt;

  let html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#172b4d;font-size:13px;max-width:960px">';

  if (teste) {
    html += '<div style="background:#fffae6;border:1px solid #ffe380;padding:8px 10px;border-radius:4px;margin-bottom:12px;font-size:12px">' +
      '🧪 <b>E-mail de teste.</b> Na operação real, ele vai para os destinatários da coluna E da aba Mapeamento_Times.</div>';
  }

  html += '<h2 style="margin:0 0 4px;font-size:18px">📊 SLA CLT/DLT · ' + e(time) + '</h2>';
  html += '<div style="color:#6b778c;font-size:12px;margin-bottom:14px">' + hojeFmt + ' · ' +
    e(contexto.linhaBase.replace('🕒 ', '')) +
    (contexto.desatualizada ? ' · <b style="color:#bf2600">⚠️ base desatualizada</b>' : '') + '</div>';

  html += '<table role="presentation" cellspacing="0" cellpadding="0" style="margin-bottom:16px"><tr>' +
    cartaoEmail_('Estourados', est.length, '#bf2600', '#ffebe6') +
    cartaoEmail_('A vencer (≤5d)', prox.length, '#974f0c', '#fffae6') +
    cartaoEmail_('No prazo', dentro.length, '#006644', '#e3fcef') +
    '</tr></table>';

  if (!est.length && !prox.length) html += '<p>✅ Nenhum chamado vencido ou próximo de vencer.</p>';

  html += tabelaEmail_('🔴 Passaram do prazo', '#bf2600', est, tz);
  html += tabelaEmail_('🟡 Próximos de vencer', '#974f0c', prox, tz);
  if (EMAIL_CFG.LISTAR_DENTRO_PRAZO) html += tabelaEmail_('🟢 Dentro do prazo', '#006644', dentro, tz);

  html += '<p style="color:#6b778c;font-size:11px;margin-top:20px">' +
    (CONFIG.LOOKER_URL ? '<a href="' + e(CONFIG.LOOKER_URL) + '" style="color:#0052CC">📈 Abrir dashboard no Looker</a> · ' : '') +
    'Mensagem automática do controle de SLA CLT/DLT. Prazos em dias corridos, recalculados no momento do envio.</p></div>';

  const txt = [assunto, contexto.linhaBase, ''];
  est.concat(prox).forEach(i => {
    txt.push('- ' + i.key + ' | ' + i.prioridade + ' | CLT ' + i.clt.dias + '/' + i.clt.limite + 'd | DLT ' +
      i.dlt.dias + '/' + (i.dlt.limite === null ? '∞' : i.dlt.limite + 'd') + ' | ' + CONFIG.JIRA_BASE_URL + '/browse/' + i.key);
  });
  txt.push('', 'No prazo: ' + dentro.length + ' chamado(s).');

  return { assunto: assunto, html: html, texto: txt.join('\n') };
}

function cartaoEmail_(rotulo, n, cor, fundo) {
  return '<td style="padding-right:8px"><div style="background:' + fundo + ';border-radius:6px;padding:8px 14px;text-align:center;min-width:90px">' +
    '<div style="font-size:20px;font-weight:bold;color:' + cor + '">' + n + '</div>' +
    '<div style="font-size:11px;color:' + cor + '">' + rotulo + '</div></div></td>';
}

function tabelaEmail_(titulo, cor, itens, tz) {
  if (!itens.length) return '';
  const e = escaparHtml_;
  const th = 'style="background:#f4f5f7;text-align:left;padding:6px 8px;font-size:11px;border-bottom:2px solid #dfe1e6"';
  const td = 'style="padding:6px 8px;border-bottom:1px solid #ebecf0;vertical-align:top;font-size:12px"';

  let h = '<h3 style="color:' + cor + ';font-size:14px;margin:18px 0 6px">' + titulo + ' (' + itens.length + ')</h3>';
  h += '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%">';
  h += '<tr><th ' + th + '>Issue</th><th ' + th + '>Prioridade</th><th ' + th + '>Status</th><th ' + th +
    '>Resumo</th><th ' + th + '>Responsável</th><th ' + th + '>CLT</th><th ' + th + '>DLT</th></tr>';

  itens.slice(0, EMAIL_CFG.MAX_ITENS_POR_GRUPO).forEach(i => {
    h += '<tr>' +
      '<td ' + td + '><a href="' + e(CONFIG.JIRA_BASE_URL + '/browse/' + i.key) + '" style="color:#0052CC;font-weight:bold;white-space:nowrap">' + e(i.key) + '</a></td>' +
      '<td ' + td + '>' + e(i.prioridade) + '</td>' +
      '<td ' + td + '>' + e(i.status) + '</td>' +
      '<td ' + td + '>' + e(truncar_(i.resumo, 90)) + '</td>' +
      '<td ' + td + '>' + e(i.responsavel) + '</td>' +
      '<td ' + td + '>' + celulaPrazoEmail_(i.clt, tz) + '</td>' +
      '<td ' + td + '>' + celulaPrazoEmail_(i.dlt, tz) + '</td>' +
      '</tr>';
  });
  h += '</table>';

  if (itens.length > EMAIL_CFG.MAX_ITENS_POR_GRUPO) {
    h += '<p style="font-size:11px;color:#6b778c">…e mais ' + (itens.length - EMAIL_CFG.MAX_ITENS_POR_GRUPO) + ' chamado(s). Veja a planilha.</p>';
  }
  return h;
}

function celulaPrazoEmail_(p, tz) {
  if (p.limite === null) return p.dias + 'd<br><span style="color:#6b778c;font-size:11px">sem limite</span>';
  const venc = p.venc ? Utilities.formatDate(p.venc, tz, 'dd/MM') : '?';
  const base = p.dias + '/' + p.limite + 'd';
  if (p.rest < 0) {
    return '<b style="color:#bf2600">' + base + '</b><br><span style="color:#bf2600;font-size:11px;white-space:nowrap">venceu há ' + Math.abs(p.rest) + 'd (' + venc + ')</span>';
  }
  if (p.rest <= 5) {
    return '<b style="color:#974f0c">' + base + '</b><br><span style="color:#974f0c;font-size:11px;white-space:nowrap">vence em ' + p.rest + 'd (' + venc + ')</span>';
  }
  return base + '<br><span style="color:#6b778c;font-size:11px;white-space:nowrap">vence ' + venc + '</span>';
}

/* ===================== ENVIO - CHAT ===================== */

function enviarChatComRetry_(webhook, texto, threadKey) {
  const sep = webhook.indexOf('?') >= 0 ? '&' : '?';
  const url = webhook + sep + 'threadKey=' + encodeURIComponent(threadKey) +
    '&messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD';

  for (let t = 1; t <= CHAT_CFG.TENTATIVAS; t++) {
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      payload: JSON.stringify({ text: texto }),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) return;
    if (code === 429 || code >= 500) {
      Utilities.sleep(1000 * Math.pow(2, t));
      continue;
    }
    throw new Error('HTTP ' + code + ': ' + resp.getContentText().substring(0, 200));
  }
  throw new Error('Falha após ' + CHAT_CFG.TENTATIVAS + ' tentativas (rate limit/servidor).');
}

/* ===================== HELPERS ===================== */

/** getUi() lança exceção em gatilhos/editor: nunca usar sem proteção. */
function alertaSeguro_(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    registrarLog_('UI_INDISPONIVEL', 'INFO', String(msg).substring(0, 300));
  }
}

function ehFimDeSemana_() {
  const dia = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'u')); // 1=seg ... 7=dom
  return dia >= 6;
}

function parseDataPlanilha_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const p = String(v || '').substring(0, 10).split('-');
  if (p.length === 3) return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  return null;
}

function diasEntre_(inicio, fim) {
  return Math.round((fim - inicio) / 86400000);
}

function somarDias_(data, dias) {
  const d = new Date(data.getTime());
  d.setDate(d.getDate() + Number(dias));
  return d;
}

/** Remove caracteres que quebram a formatação do Chat e trunca o texto. */
function limparTexto_(txt, max) {
  let s = String(txt || '').replace(/[*_~`<>|]/g, '').replace(/\s+/g, ' ').trim();
  if (max && s.length > max) s = s.substring(0, max - 1) + '…';
  return s;
}

function truncar_(txt, max) {
  const s = String(txt || '').replace(/\s+/g, ' ').trim();
  return s.length > max ? s.substring(0, max - 1) + '…' : s;
}

function slug_(txt) {
  return String(txt).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '-').toLowerCase().substring(0, 40);
}

function escaparHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
