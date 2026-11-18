/**
 * Placement Check -> Planilha de Matrícula (UPSERT por e-mail)
 *
 * Quando o professor preenche/edita o Feedback (col. J) ou o Level (col. I) de um
 * aluno na aba "06/26" da Placement Check, o script envia os dados para a planilha
 * de matrícula:
 *   - Se o e-mail já existe lá  -> ATUALIZA Nome e Nível naquela linha.
 *   - Se o e-mail não existe     -> CRIA uma nova linha (Nome, E-mail, Nível).
 *   - Celular (col. C do destino) nunca é tocado.
 *   - Sem e-mail na origem        -> não envia.
 *
 * Este script deve ficar na Placement Check (Extensões > Apps Script) e usa um
 * gatilho INSTALÁVEL onEdit (necessário para escrever em outra planilha).
 */

// ===================== CONFIG =====================
var CFG = {
  // --- ORIGEM (Placement Check) ---
  SRC_ABA: '06/26',          // nome EXATO da aba de origem
  SRC_LINHA_CABECALHO: 1,    // ajuste se o cabeçalho não estiver na linha 1
  SRC_COL_NOME:     4,       // D
  SRC_COL_EMAIL:    5,       // E
  SRC_COL_LEVEL:    9,       // I
  SRC_COL_FEEDBACK: 10,      // J
  // Edições nestas colunas disparam a sincronização (Nome, E-mail, Level, Feedback):
  COLS_GATILHO: [4, 5, 9, 10],

  // --- DESTINO (Matrícula) ---
  DEST_ID: 'COLE_AQUI_O_ID_DA_PLANILHA_DE_MATRICULA',
  DEST_GID: 0,               // 1ª aba (gid=0)
  DEST_LINHA_CABECALHO: 1,   // dados começam na linha seguinte
  DEST_COL_NOME:  1,         // A
  DEST_COL_EMAIL: 2,         // B  (chave de identificação)
  DEST_COL_NIVEL: 4,         // D  (Celular = C: nunca é tocado)
};
// ==================================================

/**
 * Gatilho instalável de edição. NÃO renomeie sem atualizar criarGatilho().
 */
function aoEditarPlacement(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CFG.SRC_ABA) return;

  // A edição precisa tocar em pelo menos uma das colunas-gatilho.
  var startCol = e.range.getColumn();
  var endCol = startCol + e.range.getNumColumns() - 1;
  var intersecta = CFG.COLS_GATILHO.some(function (c) { return c >= startCol && c <= endCol; });
  if (!intersecta) return;

  var startRow = e.range.getRow();
  for (var r = 0; r < e.range.getNumRows(); r++) {
    var linha = startRow + r;
    if (linha <= CFG.SRC_LINHA_CABECALHO) continue;
    processarLinhaOrigem_(sheet, linha);
  }
}

/** Lê uma linha da origem e faz o upsert no destino, se estiver "pronta". */
function processarLinhaOrigem_(sheet, linha) {
  var nome     = sheet.getRange(linha, CFG.SRC_COL_NOME).getDisplayValue().toString().trim();
  var email    = sheet.getRange(linha, CFG.SRC_COL_EMAIL).getDisplayValue().toString().trim();
  var nivel    = sheet.getRange(linha, CFG.SRC_COL_LEVEL).getDisplayValue().toString().trim();
  var feedback = sheet.getRange(linha, CFG.SRC_COL_FEEDBACK).getDisplayValue().toString().trim();

  // "Pronta" = tem e-mail E tem feedback. Sem isso, não sobe.
  if (!email || !feedback) return;

  upsertMatricula_({ nome: nome, email: email, nivel: nivel });
}

/** Cria ou atualiza a linha do aluno na planilha de matrícula, chaveando por e-mail. */
function upsertMatricula_(dados) {
  var ss = SpreadsheetApp.openById(CFG.DEST_ID);
  var sheet = pegarAbaPorGid_(ss, CFG.DEST_GID);
  var ultimaLinha = sheet.getLastRow();
  var primeiraDados = CFG.DEST_LINHA_CABECALHO + 1;
  var alvo = dados.email.toLowerCase();

  var linhaEncontrada = -1;
  if (ultimaLinha >= primeiraDados) {
    var emails = sheet
      .getRange(primeiraDados, CFG.DEST_COL_EMAIL, ultimaLinha - primeiraDados + 1, 1)
      .getValues();
    for (var i = 0; i < emails.length; i++) {
      if ((emails[i][0] || '').toString().trim().toLowerCase() === alvo) {
        linhaEncontrada = primeiraDados + i;
        break;
      }
    }
  }

  if (linhaEncontrada > 0) {
    // Atualiza Nome e Nível; não toca em Celular (C) nem no E-mail (chave).
    sheet.getRange(linhaEncontrada, CFG.DEST_COL_NOME).setValue(dados.nome);
    sheet.getRange(linhaEncontrada, CFG.DEST_COL_NIVEL).setValue(dados.nivel);
  } else {
    // Cria nova linha: Nome (A), E-mail (B), Nível (D). Celular (C) fica em branco.
    var nova = ultimaLinha + 1;
    sheet.getRange(nova, CFG.DEST_COL_NOME).setValue(dados.nome);
    sheet.getRange(nova, CFG.DEST_COL_EMAIL).setValue(dados.email);
    sheet.getRange(nova, CFG.DEST_COL_NIVEL).setValue(dados.nivel);
  }
}

/** Acha a aba pelo gid (sheetId); cai na 1ª aba se não encontrar. */
function pegarAbaPorGid_(ss, gid) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  return sheets[0];
}

// ===================== SETUP / TESTE =====================

/**
 * Rode UMA vez para criar o gatilho instalável de edição.
 * Vai pedir autorização (inclui permissão para a planilha de destino).
 */
function criarGatilhoPlacement() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'aoEditarPlacement') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('aoEditarPlacement')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  console.log('Gatilho criado com sucesso.');
}

/**
 * Teste: processa UMA linha da aba de origem (sem precisar editar a planilha).
 * Altere LINHA_TESTE para o número da linha de um aluno que já tenha e-mail e feedback.
 */
function testarLinhaPlacement() {
  var LINHA_TESTE = 2; // <-- ajuste para a linha de um aluno real (com e-mail + feedback)
  var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.SRC_ABA);
  if (!sheet) { console.error('Aba "' + CFG.SRC_ABA + '" não encontrada.'); return; }
  processarLinhaOrigem_(sheet, LINHA_TESTE);
  console.log('Processada a linha ' + LINHA_TESTE + '. Confira a planilha de matrícula.');
}
