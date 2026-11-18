/**
 * Backend do HISTÓRICO do Gerador de Recibos
 * ------------------------------------------------------------------
 * COMO PUBLICAR (uma vez):
 *  1. Crie uma Planilha Google (Google Sheets) nova, no seu Workspace.
 *  2. Menu Extensões > Apps Script. Apague o conteúdo e cole ESTE arquivo.
 *  3. Defina a senha em Script Properties (NUNCA no código):
 *       Configurações do projeto (engrenagem) > Propriedades do script >
 *       Adicionar propriedade:  HISTORICO_SENHA = <senha forte>
 *     (ou rode uma vez definirSenha('...') no editor e apague o valor depois).
 *  4. Implantar > Nova implantação > Tipo: "App da Web".
 *       - Executar como:  Eu (sua conta)
 *       - Quem pode acessar:  Qualquer pessoa
 *  5. Copie a URL que termina em /exec e cole no drawer "Recibos Gerados"
 *     da ferramenta (primeiro acesso pede a URL; fica salva no navegador).
 *
 * Segurança: SALVAR é automático e sem senha (só grava). LER e EXCLUIR
 * exigem a SENHA e têm bloqueio de 15 min após 5 tentativas erradas.
 * A planilha guarda CPF/nome — use uma senha FORTE e só com quem consulta.
 */

// A senha de leitura/exclusão fica em Script Properties (chave HISTORICO_SENHA).
// Sem ela configurada, LER e EXCLUIR ficam bloqueados (falha fechada).
var PROP_SENHA = 'HISTORICO_SENHA';
function senhaConfigurada_(){
  return String(PropertiesService.getScriptProperties().getProperty(PROP_SENHA) || '');
}
/* Utilitário de setup: rode UMA vez no editor com a senha desejada e depois
   apague o argumento do histórico de execução/código. */
function definirSenha(senha){
  if(!senha || String(senha).length < 12) throw new Error('Use uma senha forte (12+ caracteres).');
  PropertiesService.getScriptProperties().setProperty(PROP_SENHA, String(senha));
}
var ABA   = 'Recibos';
var MAX_CAMPO = 300;   // limite de caracteres por campo (anti-abuso)

/* Colunas do histórico. Para adicionar um campo novo, ACRESCENTE no fim desta
   lista: o cabeçalho da planilha existente se auto-completa na próxima gravação
   (ver sheet_), e as linhas são montadas na ordem do cabeçalho real da planilha,
   não na ordem daqui — então nada desalinha. */
var COLS  = ['id','criadoEm','tipo','nome','cpf','email','idioma','curso',
             'mes','ano','valor','forma','cidade','data','instituicao','cnpjInst',
             // Termo de Distrato
             'dataContrato','dataSolicitacao','mesesContratados','mesesUtilizados',
             'tipoPagamento','valorServico','valorMaterial','parcelasPagas',
             'parcelasAbertas','multa','reembolso','saldoDevedor','prazo','ajusteManual'];

function doGet(){
  return ContentService
    .createTextOutput('OK — backend do histórico ativo.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e){
  var out = {ok:false};
  try{
    var body = (e && e.postData) ? JSON.parse(e.postData.contents) : {};
    var acao = body.action;

    // SALVAR é automático e NÃO exige senha (só grava, não expõe dados).
    if(acao==='registrar'){
      return json(registrar(body.registro||{}, body.id||''));
    }

    // LER / EXCLUIR expõem CPF/nome -> exigem senha + bloqueio anti-tentativas.
    var chk = checkSenha_(body.senha);
    if(chk !== 'ok'){ return json({ok:false, error:chk}); }  // 'senha' ou 'bloqueado'
    if(acao==='listar')        out = listar();
    else if(acao==='excluir')  out = excluir(body.id||'');
    else                       out = {ok:false, error:'ação desconhecida'};
  }catch(err){
    out = {ok:false, error:'erro interno'};   // não vaza detalhes ao cliente
  }
  return json(out);
}

function json(o){
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Bloqueio anti-força-bruta: 5 senhas erradas -> trava 15 min.
   Retorna 'ok', 'senha' (errada) ou 'bloqueado'. */
var MAX_TENTATIVAS = 5;
var BLOQUEIO_MS    = 15 * 60 * 1000;
function checkSenha_(senha){
  var props = PropertiesService.getScriptProperties();
  if(Date.now() < Number(props.getProperty('travadoAte') || 0)){ return 'bloqueado'; }

  var esperada = senhaConfigurada_();
  if(!esperada){ return 'senha'; }   // sem senha configurada: nunca libera

  // caminho da senha CORRETA: sem lock, para abrir rápido (comparação tempo-constante)
  if(igualSeguro_(String(senha||''), esperada)){
    props.deleteProperty('tentativas');
    return 'ok';
  }

  // caminho da FALHA: lock só aqui, para o contador não ter corrida
  var lock = LockService.getScriptLock();
  try { lock.waitLock(2000); } catch(e) {}
  try {
    var t = Number(props.getProperty('tentativas') || 0) + 1;
    if(t >= MAX_TENTATIVAS){
      props.setProperty('travadoAte', String(Date.now() + BLOQUEIO_MS));
      props.deleteProperty('tentativas');
      return 'bloqueado';
    }
    props.setProperty('tentativas', String(t));
    return 'senha';
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function invalidarCache_(){ try { CacheService.getScriptCache().remove('lista'); } catch(e) {} }

function igualSeguro_(a, b){
  if(a.length !== b.length) return false;
  var r = 0;
  for(var i=0;i<a.length;i++){ r |= (a.charCodeAt(i) ^ b.charCodeAt(i)); }
  return r === 0;
}

/* Limpa cada campo antes de gravar: corta tamanho e neutraliza fórmula no Sheets. */
function limpaCampo_(v){
  v = (v == null ? '' : String(v)).slice(0, MAX_CAMPO);
  if(/^[=+\-@]/.test(v)){ v = "'" + v; }
  return v;
}

function sheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA);
  if(!sh){ sh = ss.insertSheet(ABA); }
  if(sh.getLastRow() === 0){ sh.appendRow(COLS); return sh; }

  // Planilha antiga (criada antes do Termo de Distrato): acrescenta ao cabeçalho
  // as colunas que faltam, sem tocar nas linhas já gravadas.
  var head = headOf_(sh);
  var faltando = COLS.filter(function(c){ return head.indexOf(c) === -1; });
  if(faltando.length){
    sh.getRange(1, head.length + 1, 1, faltando.length).setValues([faltando]);
  }
  return sh;
}

/* O cabeçalho REAL da planilha manda na ordem das colunas — nunca o array COLS.
   É o que impede um campo novo de cair na coluna errada numa planilha antiga. */
function headOf_(sh){
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}
function linhaPara_(head, reg, id, criadoEm){
  return head.map(function(c){
    if(c === 'id')       return id;
    if(c === 'criadoEm') return criadoEm;
    return limpaCampo_(reg[c]);
  });
}

function listar(){
  var cache = CacheService.getScriptCache();
  var hit = cache.get('lista');
  if(hit){ return {ok:true, registros: JSON.parse(hit)}; }

  var sh = sheet_();
  var rng = sh.getDataRange().getValues();
  var head = rng[0] || COLS;
  var regs = [];
  for(var i=1;i<rng.length;i++){
    var o = {};
    for(var c=0;c<head.length;c++){ o[head[c]] = rng[i][c]; }
    if(o.id) regs.push(o);
  }
  try {
    var str = JSON.stringify(regs);
    if(str.length < 95000){ cache.put('lista', str, 60); }  // 60s (limite ~100KB)
  } catch(e) {}
  return {ok:true, registros:regs};
}

function registrar(reg, id){
  var sh = sheet_();
  var head = headOf_(sh);
  var tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  var iso = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");

  if(id){  // upsert: se já existe esse id, atualiza a linha
    var vals = sh.getDataRange().getValues();
    for(var i=1;i<vals.length;i++){
      if(String(vals[i][0]) === String(id)){
        var criado = vals[i][1] || iso;
        sh.getRange(i+1, 1, 1, head.length)
          .setValues([linhaPara_(head, reg, id, criado)]);
        invalidarCache_();
        return {ok:true, id:id, criadoEm:criado};
      }
    }
  }

  var novoId = Utilities.getUuid();
  sh.appendRow(linhaPara_(head, reg, novoId, iso));
  invalidarCache_();
  return {ok:true, id:novoId, criadoEm:iso};
}

function excluir(id){
  var sh = sheet_();
  var vals = sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][0]) === String(id)){
      sh.deleteRow(i+1);
      invalidarCache_();
      return {ok:true};
    }
  }
  return {ok:false, error:'id não encontrado'};
}
