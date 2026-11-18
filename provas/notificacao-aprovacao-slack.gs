/**
 * Notificação no Slack quando o Status de um aluno vira "Aprovado" ou "Reprovado"
 * na aba TESTS da planilha "Controle de Correções".
 *
 * Como funciona:
 *  - Um gatilho INSTALÁVEL onEdit dispara a cada edição na planilha.
 *  - Se a edição for na aba TESTS, na coluna de Status, e o novo valor for
 *    "Aprovado" ou "Reprovado", envia uma mensagem para o Slack via Incoming Webhook,
 *    marcando (@) as 2 colaboradoras e informando Nome, Nível, Prova, E-mail e Resultado.
 *
 * IMPORTANTE: precisa ser um gatilho INSTALÁVEL (não o onEdit simples), porque
 * o onEdit simples NÃO tem permissão para chamar serviços externos (UrlFetchApp).
 * A função criarGatilho() abaixo cria o gatilho instalável para você.
 */

// ===================== CONFIG — PREENCHA AQUI =====================
var CONFIG = {
  // Nome EXATO da aba (guia) onde fica a tabela de correções.
  ABA: 'TESTS',

  // Coluna de Status (a que vira "Aprovado"/"Reprovado"/"Refazer").
  COLUNA_STATUS: 'X',

  // Colunas dos dados a incluir na mensagem (use as LETRAS das colunas).
  COLUNA_NOME:  'B',  // Aluno
  COLUNA_NIVEL: 'C',  // Nível
  COLUNA_PROVA: 'D',  // Prova
  COLUNA_EMAIL: 'E',  // E-mail

  // Linha do cabeçalho (os dados começam na linha seguinte). Normalmente 1.
  LINHA_CABECALHO: 1,

  // Valores que disparam a notificação (comparação sem diferenciar maiúsc/minúsc).
  STATUS_GATILHO: ['Aprovado', 'Reprovado', 'Refazer'],

  // IDs no Slack (Perfil > ... > "Copiar ID do membro"): você + as 2 colaboradoras.
  SLACK_USER_IDS: ['SEU_USER_ID_1', 'SEU_USER_ID_2', 'SEU_USER_ID_3'],
};

// A URL do Incoming Webhook fica em Script Properties (não no código), por segurança.
// Rode uma vez a função guardarWebhook() depois de colar a URL nela. Veja abaixo.
function getWebhookUrl_() {
  var url = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!url) throw new Error('SLACK_WEBHOOK_URL não configurada. Rode guardarWebhook() uma vez.');
  return url;
}
// =================================================================

/**
 * Gatilho instalável de edição. NÃO renomeie sem atualizar criarGatilho().
 */
function aoEditar(e) {
  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() !== CONFIG.ABA) return;

  var colStatus = colunaParaIndice_(CONFIG.COLUNA_STATUS);
  var linha = range.getRow();

  // Só reage à edição da célula de Status, abaixo do cabeçalho.
  if (range.getColumn() !== colStatus) return;
  if (linha <= CONFIG.LINHA_CABECALHO) return;

  var valor = (e.value || '').toString().trim();
  if (!valor) return;

  var disparou = CONFIG.STATUS_GATILHO.some(function (s) {
    return s.toLowerCase() === valor.toLowerCase();
  });
  if (!disparou) return;

  var dados = {
    nome:  lerCelula_(sheet, linha, CONFIG.COLUNA_NOME),
    nivel: lerCelula_(sheet, linha, CONFIG.COLUNA_NIVEL),
    prova: lerCelula_(sheet, linha, CONFIG.COLUNA_PROVA),
    email: lerCelula_(sheet, linha, CONFIG.COLUNA_EMAIL),
    resultado: valor,
  };

  enviarSlack_(dados);
}

function lerCelula_(sheet, linha, colunaLetra) {
  var col = colunaParaIndice_(colunaLetra);
  return sheet.getRange(linha, col).getDisplayValue().toString().trim();
}

/** Converte letra de coluna ('A','B','X','AA') em índice numérico (1-based). */
function colunaParaIndice_(letra) {
  var s = letra.toUpperCase();
  var n = 0;
  for (var i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

function enviarSlack_(d) {
  var mencoes = CONFIG.SLACK_USER_IDS.map(function (id) { return '<@' + id + '>'; }).join(' ');
  var emoji = /aprovad/i.test(d.resultado) ? ':white_check_mark:'
            : /refazer/i.test(d.resultado) ? ':arrows_counterclockwise:'
            : ':x:';

  var payload = {
    text: emoji + ' Aluno com resultado *' + d.resultado + '* — ' + d.nome,
    blocks: [
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: emoji + ' *Resultado de prova atualizado*\n' + mencoes,
        },
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: '*Nome:*\n' + (d.nome || '—') },
          { type: 'mrkdwn', text: '*Nível:*\n' + (d.nivel || '—') },
          { type: 'mrkdwn', text: '*Prova:*\n' + (d.prova || '—') },
          { type: 'mrkdwn', text: '*E-mail:*\n' + (d.email || '—') },
          { type: 'mrkdwn', text: '*Resultado:*\n' + (d.resultado || '—') },
        ],
      },
    ],
  };

  var resp = UrlFetchApp.fetch(getWebhookUrl_(), {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) {
    console.error('Falha ao enviar ao Slack: ' + resp.getResponseCode() + ' ' + resp.getContentText());
  }
}

// ===================== FUNÇÕES DE SETUP (rodar uma vez) =====================

/**
 * 1) Cole a URL do Incoming Webhook entre as aspas e rode esta função UMA vez.
 *    Depois apague a URL daqui (ela já fica salva nas Script Properties).
 */
function guardarWebhook() {
  var URL = 'COLE_AQUI_A_URL_DO_WEBHOOK_DO_SLACK';
  PropertiesService.getScriptProperties().setProperty('SLACK_WEBHOOK_URL', URL);
  console.log('Webhook salvo com sucesso.');
}

/**
 * 2) Rode esta função UMA vez para criar o gatilho instalável de edição.
 *    Vai pedir autorização — aceite com a conta que tem acesso à planilha.
 */
function criarGatilho() {
  var ss = SpreadsheetApp.getActive();
  // Remove gatilhos duplicados da mesma função, se já existirem.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'aoEditar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('aoEditar')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  console.log('Gatilho criado com sucesso.');
}

/**
 * 3) (Opcional) Teste o envio ao Slack sem precisar editar a planilha.
 */
function testarEnvio() {
  enviarSlack_({
    nome: 'Aluno Teste',
    nivel: 'Nível 3',
    prova: 'Test 2',
    email: 'aluno.teste@exemplo.com',
    resultado: 'Aprovado',
  });
}
