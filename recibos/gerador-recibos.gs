/**
 * Gerador Automático de Recibos — Escola de Idiomas Exemplo
 *
 * Gera um recibo em PDF a partir de cada linha "Pago" da aba "Pagamentos",
 * salva o PDF numa pasta do Drive, escreve de volta na planilha o Nº do recibo
 * e o link do PDF, e (opcional) envia o recibo por e-mail para o aluno.
 *
 * Duas formas de disparar:
 *   1) MENU  "Recibos" > "Gerar recibos pendentes"  — gera em lote todas as
 *      linhas que estão "Pago" e ainda não têm recibo.
 *   2) AUTOMÁTICO ao editar — quando o Status de uma linha vira "Pago", o recibo
 *      é gerado sozinho (gatilho INSTALÁVEL onEdit, igual ao script do Slack).
 *
 * IMPORTANTE: gerar PDF, salvar no Drive e enviar e-mail exigem gatilho
 * INSTALÁVEL (não o onEdit simples). Rode criarGatilhoRecibos() uma vez.
 *
 * SETUP (rode uma vez, nesta ordem):
 *   1) Ajuste o bloco CONFIG abaixo (aba, colunas, dados da empresa).
 *   2) Crie uma pasta no Drive para os PDFs e cole o ID em CFG.PASTA_DRIVE_ID.
 *   3) Rode criarGatilhoRecibos()  -> autoriza e cria o gatilho automático.
 *   4) (Opcional) testarRecibo()   -> gera 1 PDF de exemplo para conferir o layout.
 */

// ===================== CONFIG — PREENCHA AQUI =====================
var CFG = {
  // --- PLANILHA DE PAGAMENTOS ---
  ABA: 'Pagamentos',          // nome EXATO da aba
  LINHA_CABECALHO: 1,         // dados começam na linha seguinte

  // Colunas de ENTRADA (use as LETRAS das colunas)
  COL_DATA:      'A',  // Data do pagamento
  COL_NOME:      'B',  // Nome do aluno / pagador
  COL_EMAIL:     'C',  // E-mail (para envio; pode ficar vazio)
  COL_DESCRICAO: 'D',  // Referência (ex.: "Mensalidade Julho/2026 - Inglês")
  COL_VALOR:     'E',  // Valor pago (número ou "R$ 350,00")
  COL_STATUS:    'F',  // Status: gera quando vira "Pago"

  // Colunas de SAÍDA (o script escreve aqui; deixe-as vazias na planilha)
  COL_NUM_RECIBO: 'G', // Nº do recibo gerado
  COL_LINK_PDF:   'H', // Link do PDF no Drive
  COL_DATA_GER:   'I', // Data/hora em que o recibo foi gerado

  // Status que dispara a geração (comparação sem diferenciar maiúsc/minúsc)
  STATUS_GATILHO: 'Pago',

  // --- SAÍDA ---
  PASTA_DRIVE_ID: 'COLE_AQUI_O_ID_DA_PASTA_DO_DRIVE', // pasta onde os PDFs são salvos
  ENVIAR_EMAIL: false,        // true = envia o PDF por e-mail para o aluno (COL_EMAIL)

  // --- DADOS DA EMPRESA (aparecem no recibo) ---
  EMPRESA: {
    nome:   'Escola de Idiomas Exemplo',
    cnpj:   '',               // ex.: '12.345.678/0001-90' (deixe vazio se não usar)
    cidade: 'São Paulo',      // cidade que aparece na data ("São Paulo, 03 de julho de 2026")
    responsavel: '',          // nome de quem assina (deixe vazio para linha em branco)
  },
};
// =================================================================


// ===================== MENU =====================

/** Cria o menu "Recibos" ao abrir a planilha. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Recibos')
    .addItem('Gerar recibos pendentes', 'gerarRecibosPendentes')
    .addSeparator()
    .addItem('Gerar recibo da linha selecionada', 'gerarReciboSelecionado')
    .addToUi();
}


// ===================== GATILHO AUTOMÁTICO =====================

/**
 * Gatilho instalável de edição. NÃO renomeie sem atualizar criarGatilhoRecibos().
 * Gera o recibo quando o Status de uma linha vira "Pago".
 */
function aoEditarPagamento(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CFG.ABA) return;

  var colStatus = colunaParaIndice_(CFG.COL_STATUS);
  if (e.range.getColumn() !== colStatus) return;

  var linha = e.range.getRow();
  if (linha <= CFG.LINHA_CABECALHO) return;

  var valor = (e.value || '').toString().trim();
  if (valor.toLowerCase() !== CFG.STATUS_GATILHO.toLowerCase()) return;

  gerarReciboLinha_(sheet, linha, { silencioso: true });
}


// ===================== AÇÕES DO MENU =====================

/** Gera recibos para todas as linhas "Pago" que ainda não têm recibo. */
function gerarRecibosPendentes() {
  var sheet = pegarAbaPagamentos_();
  var ultima = sheet.getLastRow();
  var inicio = CFG.LINHA_CABECALHO + 1;
  if (ultima < inicio) { alerta_('Nenhuma linha de dados encontrada.'); return; }

  var colStatus = colunaParaIndice_(CFG.COL_STATUS);
  var colNum    = colunaParaIndice_(CFG.COL_NUM_RECIBO);

  var gerados = 0, ignorados = 0, erros = 0;
  for (var linha = inicio; linha <= ultima; linha++) {
    var status = sheet.getRange(linha, colStatus).getDisplayValue().toString().trim();
    var jaTem  = sheet.getRange(linha, colNum).getDisplayValue().toString().trim();
    if (status.toLowerCase() !== CFG.STATUS_GATILHO.toLowerCase() || jaTem) { ignorados++; continue; }
    try {
      gerarReciboLinha_(sheet, linha, { silencioso: true });
      gerados++;
    } catch (err) {
      erros++;
      console.error('Linha ' + linha + ': ' + err.message);
    }
  }
  alerta_('Recibos gerados: ' + gerados + '\nIgnorados (sem "Pago" ou já com recibo): ' + ignorados + '\nErros: ' + erros);
}

/** Gera o recibo da linha atualmente selecionada (mesmo sem estar "Pago"). */
function gerarReciboSelecionado() {
  var sheet = pegarAbaPagamentos_();
  var linha = sheet.getActiveRange().getRow();
  if (linha <= CFG.LINHA_CABECALHO) { alerta_('Selecione uma linha de dados (abaixo do cabeçalho).'); return; }
  try {
    var r = gerarReciboLinha_(sheet, linha, { silencioso: false, forcar: true });
    alerta_('Recibo ' + r.numero + ' gerado.\n' + r.url);
  } catch (err) {
    alerta_('Erro: ' + err.message);
  }
}


// ===================== NÚCLEO: GERAR 1 RECIBO =====================

/**
 * Lê a linha, gera o PDF, salva no Drive, escreve de volta e (opcional) envia e-mail.
 * opts.forcar = gera mesmo que a linha já tenha um Nº de recibo (regera).
 * Retorna { numero, url }.
 */
function gerarReciboLinha_(sheet, linha, opts) {
  opts = opts || {};

  var dados = {
    data:      lerCelula_(sheet, linha, CFG.COL_DATA),
    nome:      lerCelula_(sheet, linha, CFG.COL_NOME),
    email:     lerCelula_(sheet, linha, CFG.COL_EMAIL),
    descricao: lerCelula_(sheet, linha, CFG.COL_DESCRICAO),
    valorTxt:  lerCelula_(sheet, linha, CFG.COL_VALOR),
  };

  if (!dados.nome)     throw new Error('Sem nome do pagador (col. ' + CFG.COL_NOME + ').');
  var valorNum = parsearValor_(dados.valorTxt);
  if (!(valorNum > 0)) throw new Error('Valor inválido na col. ' + CFG.COL_VALOR + ': "' + dados.valorTxt + '".');

  // Não regera se já existe recibo, a menos que forcar=true.
  var jaTem = lerCelula_(sheet, linha, CFG.COL_NUM_RECIBO);
  if (jaTem && !opts.forcar) return { numero: jaTem, url: lerCelula_(sheet, linha, CFG.COL_LINK_PDF) };

  var numero = proximoNumeroRecibo_();
  var html   = montarHtmlRecibo_(numero, dados, valorNum);
  var nomeArq = 'Recibo ' + numero + ' - ' + sanitizar_(dados.nome) + '.pdf';

  var pdf = Utilities.newBlob(html, 'text/html', 'recibo.html')
    .getAs('application/pdf')
    .setName(nomeArq);

  var pasta = pegarPasta_();
  var file  = pasta.createFile(pdf);
  file.setDescription('Recibo ' + numero + ' — ' + dados.nome + ' — ' + formatarBRL_(valorNum));
  var url = file.getUrl();

  // Escreve de volta na planilha.
  escrever_(sheet, linha, CFG.COL_NUM_RECIBO, numero);
  escrever_(sheet, linha, CFG.COL_LINK_PDF, url);
  escrever_(sheet, linha, CFG.COL_DATA_GER, agoraFormatado_());

  // Envia por e-mail, se ligado e houver e-mail.
  if (CFG.ENVIAR_EMAIL && dados.email) {
    enviarEmailRecibo_(dados, numero, valorNum, pdf);
  }

  return { numero: numero, url: url };
}


// ===================== TEMPLATE HTML DO RECIBO =====================

function montarHtmlRecibo_(numero, dados, valorNum) {
  var e = CFG.EMPRESA;
  var valorBRL   = formatarBRL_(valorNum);
  var extenso    = valorPorExtenso_(valorNum);
  var dataRecibo = dados.data ? dados.data : dataHojeExtenso_();
  var dataLocal  = e.cidade + ', ' + dataHojeExtenso_();
  var cnpjLinha  = e.cnpj ? ('<div class="cnpj">CNPJ: ' + escaparHtml_(e.cnpj) + '</div>') : '';
  var assinaNome = e.responsavel ? escaparHtml_(e.responsavel) : '&nbsp;';

  return '' +
  '<!doctype html><html><head><meta charset="utf-8"><style>' +
  '  * { box-sizing: border-box; }' +
  '  body { font-family: Georgia, "Times New Roman", serif; color: #1a1a1a; margin: 0; padding: 48px 56px; }' +
  '  .rec { border: 2px solid #0b5cad; border-radius: 10px; padding: 40px 44px; max-width: 720px; margin: 0 auto; }' +
  '  .topo { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 1px solid #ddd; padding-bottom: 18px; margin-bottom: 26px; }' +
  '  .empresa { font-size: 26px; font-weight: bold; color: #0b5cad; letter-spacing: .5px; }' +
  '  .cnpj { font-size: 12px; color: #555; margin-top: 4px; }' +
  '  .meta { text-align: right; font-size: 13px; color: #333; }' +
  '  .meta .num { font-size: 15px; font-weight: bold; color: #0b5cad; }' +
  '  h1 { font-size: 20px; letter-spacing: 4px; text-align: center; margin: 8px 0 28px; color: #0b5cad; }' +
  '  .valor-box { text-align: center; font-size: 30px; font-weight: bold; border: 1px dashed #0b5cad; border-radius: 8px; padding: 12px; margin: 0 auto 28px; width: 260px; }' +
  '  .corpo { font-size: 16px; line-height: 1.9; text-align: justify; }' +
  '  .corpo b { color: #0b1a2b; }' +
  '  .extenso { font-style: italic; }' +
  '  .data { margin-top: 40px; font-size: 15px; }' +
  '  .assinatura { margin-top: 64px; text-align: center; }' +
  '  .assinatura .linha { border-top: 1px solid #333; width: 320px; margin: 0 auto 6px; }' +
  '  .assinatura .nome { font-size: 14px; }' +
  '  .rodape { margin-top: 34px; text-align: center; font-size: 11px; color: #999; }' +
  '</style></head><body>' +
  '<div class="rec">' +
  '  <div class="topo">' +
  '    <div><div class="empresa">' + escaparHtml_(e.nome) + '</div>' + cnpjLinha + '</div>' +
  '    <div class="meta"><div class="num">RECIBO Nº ' + escaparHtml_(numero) + '</div><div>Data: ' + escaparHtml_(dataRecibo) + '</div></div>' +
  '  </div>' +
  '  <h1>R E C I B O</h1>' +
  '  <div class="valor-box">' + valorBRL + '</div>' +
  '  <div class="corpo">' +
  '    Recebi(emos) de <b>' + escaparHtml_(dados.nome) + '</b> a importância de ' +
  '    <b>' + valorBRL + '</b> (<span class="extenso">' + escaparHtml_(extenso) + '</span>), ' +
  '    referente a <b>' + escaparHtml_(dados.descricao || '—') + '</b>, ' +
  '    dando plena e total quitação pelo valor recebido.' +
  '  </div>' +
  '  <div class="data">' + escaparHtml_(dataLocal) + '</div>' +
  '  <div class="assinatura"><div class="linha"></div><div class="nome">' + assinaNome + '</div>' +
  '    <div class="nome">' + escaparHtml_(e.nome) + '</div></div>' +
  '  <div class="rodape">Documento gerado automaticamente.</div>' +
  '</div>' +
  '</body></html>';
}


// ===================== E-MAIL =====================

function enviarEmailRecibo_(dados, numero, valorNum, pdfBlob) {
  var assunto = 'Recibo ' + numero + ' — ' + CFG.EMPRESA.nome;
  var corpo =
    'Olá, ' + dados.nome + '!\n\n' +
    'Segue em anexo o recibo referente a: ' + (dados.descricao || '—') + '.\n' +
    'Valor: ' + formatarBRL_(valorNum) + '\n' +
    'Recibo nº: ' + numero + '\n\n' +
    'Qualquer dúvida, é só responder este e-mail.\n\n' +
    'Atenciosamente,\n' + CFG.EMPRESA.nome;

  GmailApp.sendEmail(dados.email, assunto, corpo, {
    name: CFG.EMPRESA.nome,
    attachments: [pdfBlob],
  });
}


// ===================== NUMERAÇÃO DOS RECIBOS =====================

/**
 * Retorna o próximo número sequencial no formato ANO-0001, guardando o contador
 * nas Script Properties. Zera a cada virada de ano.
 */
function proximoNumeroRecibo_() {
  var props = PropertiesService.getScriptProperties();
  var ano = new Date().getFullYear().toString();
  var chave = 'RECIBO_SEQ_' + ano;
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var seq = parseInt(props.getProperty(chave) || '0', 10) + 1;
    props.setProperty(chave, seq.toString());
    return ano + '-' + ('0000' + seq).slice(-4);
  } finally {
    lock.releaseLock();
  }
}


// ===================== HELPERS =====================

function pegarAbaPagamentos_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.ABA);
  if (!sheet) throw new Error('Aba "' + CFG.ABA + '" não encontrada.');
  return sheet;
}

function pegarPasta_() {
  if (!CFG.PASTA_DRIVE_ID || CFG.PASTA_DRIVE_ID.indexOf('COLE_AQUI') === 0) {
    throw new Error('Configure CFG.PASTA_DRIVE_ID com o ID de uma pasta do Drive.');
  }
  return DriveApp.getFolderById(CFG.PASTA_DRIVE_ID);
}

function lerCelula_(sheet, linha, colunaLetra) {
  return sheet.getRange(linha, colunaParaIndice_(colunaLetra)).getDisplayValue().toString().trim();
}

function escrever_(sheet, linha, colunaLetra, valor) {
  sheet.getRange(linha, colunaParaIndice_(colunaLetra)).setValue(valor);
}

/** Converte letra de coluna ('A','B','AA') em índice numérico (1-based). */
function colunaParaIndice_(letra) {
  var s = letra.toUpperCase(), n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

/** Aceita "R$ 1.234,56", "1234.56", "1.234,56", "1234,5" ou número puro. */
function parsearValor_(txt) {
  if (typeof txt === 'number') return txt;
  var s = (txt || '').toString().trim().replace(/[R$\s]/g, '');
  if (!s) return NaN;
  if (s.indexOf(',') > -1) {
    // Formato brasileiro: ponto é milhar, vírgula é decimal.
    s = s.replace(/\./g, '').replace(',', '.');
  }
  var n = parseFloat(s);
  return isNaN(n) ? NaN : Math.round(n * 100) / 100;
}

function formatarBRL_(n) {
  var neg = n < 0;
  var inteiro = Math.floor(Math.abs(n));
  var cent = Math.round((Math.abs(n) - inteiro) * 100);
  var s = inteiro.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (neg ? '-' : '') + 'R$ ' + s + ',' + ('0' + cent).slice(-2);
}

function sanitizar_(s) {
  return (s || '').toString().replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
}

function escaparHtml_(s) {
  return (s == null ? '' : s.toString())
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function agoraFormatado_() {
  return Utilities.formatDate(new Date(), fusoBR_(), 'dd/MM/yyyy HH:mm');
}

function fusoBR_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'America/Sao_Paulo';
}

var MESES_BR = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
function dataHojeExtenso_() {
  var d = new Date();
  return ('0' + d.getDate()).slice(-2) + ' de ' + MESES_BR[d.getMonth()] + ' de ' + d.getFullYear();
}

function alerta_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}


// ===================== VALOR POR EXTENSO (pt-BR) =====================

function valorPorExtenso_(valor) {
  var reais = Math.floor(valor);
  var centavos = Math.round((valor - reais) * 100);
  var partes = [];
  if (reais > 0) {
    // Milhão/bilhão exato exige "de reais" (ex.: "um milhão de reais").
    var sufixo = (reais === 1) ? ' real' : ' reais';
    if (reais >= 1000000 && reais % 1000000 === 0) sufixo = ' de reais';
    partes.push(extensoInteiro_(reais) + sufixo);
  }
  if (centavos > 0) partes.push(extensoInteiro_(centavos) + (centavos === 1 ? ' centavo' : ' centavos'));
  if (partes.length === 0) return 'zero real';
  return partes.join(' e ');
}

function extensoInteiro_(n) {
  if (n === 0) return 'zero';
  var unidades = ['','um','dois','três','quatro','cinco','seis','sete','oito','nove','dez',
                  'onze','doze','treze','quatorze','quinze','dezesseis','dezessete','dezoito','dezenove'];
  var dezenas  = ['','','vinte','trinta','quarenta','cinquenta','sessenta','setenta','oitenta','noventa'];
  var centenas = ['','cento','duzentos','trezentos','quatrocentos','quinhentos',
                  'seiscentos','setecentos','oitocentos','novecentos'];

  function ate999(x) {
    if (x === 0) return '';
    if (x === 100) return 'cem';
    var c = Math.floor(x / 100), resto = x % 100, out = [];
    if (c > 0) out.push(centenas[c]);
    if (resto > 0) {
      if (resto < 20) out.push(unidades[resto]);
      else {
        var d = Math.floor(resto / 10), u = resto % 10;
        out.push(u > 0 ? dezenas[d] + ' e ' + unidades[u] : dezenas[d]);
      }
    }
    return out.join(' e ');
  }

  var milhoes  = Math.floor(n / 1000000);
  var milhares = Math.floor((n % 1000000) / 1000);
  var resto    = n % 1000;

  var grupos = [];
  if (milhoes > 0)  grupos.push(ate999(milhoes) + (milhoes === 1 ? ' milhão' : ' milhões'));
  if (milhares > 0) grupos.push(milhares === 1 ? 'mil' : ate999(milhares) + ' mil');
  if (resto > 0)    grupos.push(ate999(resto));

  // Junta os grupos: " e " antes do último se ele for < 100 ou centena redonda; senão ", ".
  var texto = grupos[0];
  for (var i = 1; i < grupos.length; i++) {
    var ehUltimo = (i === grupos.length - 1);
    var ligar = ehUltimo && (resto < 100 || resto % 100 === 0);
    texto += (ligar ? ' e ' : ', ') + grupos[i];
  }
  return texto;
}


// ===================== SETUP / TESTE (rodar uma vez) =====================

/**
 * Rode UMA vez para criar o gatilho instalável de edição.
 * Vai pedir autorização (Drive, Gmail e a planilha). Aceite com a conta certa.
 */
function criarGatilhoRecibos() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'aoEditarPagamento') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('aoEditarPagamento').forSpreadsheet(ss).onEdit().create();
  console.log('Gatilho criado. Ao marcar "' + CFG.STATUS_GATILHO + '" na col. ' + CFG.COL_STATUS + ', o recibo é gerado.');
}

/**
 * Teste: gera 1 recibo de exemplo (não mexe na planilha; salva o PDF na pasta do Drive).
 * Confira o layout e o "valor por extenso" antes de usar de verdade.
 */
function testarRecibo() {
  var numero = 'TESTE-' + Utilities.formatDate(new Date(), fusoBR_(), 'HHmmss');
  var dados = {
    data: dataHojeExtenso_(),
    nome: 'Maria da Silva',
    email: '',
    descricao: 'Mensalidade Julho/2026 — Curso de Inglês',
    valorTxt: 'R$ 1.234,56',
  };
  var html = montarHtmlRecibo_(numero, dados, 1234.56);
  var pdf = Utilities.newBlob(html, 'text/html', 'recibo.html')
    .getAs('application/pdf').setName('Recibo ' + numero + ' (TESTE).pdf');
  var file = pegarPasta_().createFile(pdf);
  console.log('Recibo de teste gerado: ' + file.getUrl());
  console.log('Valor por extenso: ' + valorPorExtenso_(1234.56));
}
