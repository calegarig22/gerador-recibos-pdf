// Função serverless (Vercel) — envia o recibo em PDF por e-mail,
// pela caixa configurada nas variáveis de ambiente (SMTP_USER / SMTP_PASS).
// Provedor: Google Workspace (smtp.gmail.com) via Senha de App.

const nodemailer = require('nodemailer');

// Hosts autorizados a usar este endpoint (bloqueia uso cross-site).
// ALLOWED_HOSTS: lista separada por vírgula, ex.: "recibos.exemplo.com.br,localhost:3000".
// ALLOWED_HOST_PREFIX (opcional): prefixo dos deploys de preview da Vercel,
// ex.: "gerador-recibos-" aceita "gerador-recibos-<hash>.vercel.app".
const ALLOWED_HOSTS = String(process.env.ALLOWED_HOSTS || '')
  .split(',').map((h) => h.trim().toLowerCase()).filter(Boolean);
const PREVIEW_PREFIX = String(process.env.ALLOWED_HOST_PREFIX || '').trim().toLowerCase();
function hostAutorizado(h) {
  h = String(h || '').toLowerCase();
  if (!h) return false;
  if (ALLOWED_HOSTS.includes(h)) return true;
  return !!PREVIEW_PREFIX && h.startsWith(PREVIEW_PREFIX) && /^[a-z0-9-]+\.vercel\.app$/.test(h);
}
function hostDe(u) { try { return new URL(u).host; } catch (e) { return ''; } }

// A Vercel corta requisições acima de 4,5 MB e o base64 infla o PDF em ~33%,
// então o limite útil do anexo fica bem abaixo disso.
const MAX_PDF_BYTES = 3 * 1024 * 1024; // 3 MB

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// Cria o transporte SMTP. O 465 (TLS direto) é o caminho normal; o 587 entra
// como alternativa quando a saída na 465 é bloqueada na rede do servidor.
function criarTransporte(user, pass, alternativo) {
  return nodemailer.createTransport(alternativo
    ? { host: 'smtp.gmail.com', port: 587, secure: false, requireTLS: true, auth: { user, pass },
        connectionTimeout: 15000, greetingTimeout: 15000, socketTimeout: 20000 }
    : { host: 'smtp.gmail.com', port: 465, secure: true, auth: { user, pass },
        connectionTimeout: 15000, greetingTimeout: 15000, socketTimeout: 20000 });
}

// Erros de conexão (e não de autenticação/destinatário) justificam a 2ª tentativa.
function ehErroDeConexao(err) {
  const c = String((err && err.code) || '');
  return c === 'ETIMEDOUT' || c === 'ECONNECTION' || c === 'ESOCKET' || c === 'ECONNREFUSED';
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'Método não permitido' });
    return;
  }

  // Só aceita chamadas originadas do próprio site.
  const oHost = hostDe(req.headers.origin || '');
  const rHost = hostDe(req.headers.referer || '');
  if (!hostAutorizado(oHost) && !hostAutorizado(rHost)) {
    res.status(403).json({ ok: false, error: 'Origem não autorizada' });
    return;
  }

  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
    const to   = String(body.to   || '').trim().slice(0, 200);
    const nome = String(body.nome || '').trim().slice(0, 120);
    const mes  = String(body.mes  || '').trim().slice(0, 20);
    const ano  = String(body.ano  || '').trim().slice(0, 8);
    const distrato = String(body.tipo || '') === 'Distrato';

    if (!to || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) {
      res.status(400).json({ ok: false, error: 'E-mail do aluno inválido' });
      return;
    }
    if (!body.pdf || typeof body.pdf !== 'string') {
      res.status(400).json({ ok: false, error: 'PDF não recebido' });
      return;
    }

    // Valida o anexo: base64 -> bytes, tamanho e assinatura de PDF.
    let buf = null;
    try { buf = Buffer.from(body.pdf, 'base64'); } catch (e) { buf = null; }
    if (!buf || buf.length === 0 || buf.length > MAX_PDF_BYTES) {
      res.status(400).json({ ok: false, error: 'Anexo inválido ou grande demais (máx. 3 MB)' });
      return;
    }
    if (buf.slice(0, 5).toString('latin1') !== '%PDF-') {
      res.status(400).json({ ok: false, error: 'O anexo não é um PDF válido' });
      return;
    }

    const user = String(process.env.SMTP_USER || '').trim();
    // A Senha de App do Google é exibida em blocos ("abcd efgh ijkl mnop"); se for
    // colada com os espaços, o Gmail recusa a autenticação. Removemos qualquer
    // espaço em branco antes de usar.
    const pass = String(process.env.SMTP_PASS || '').replace(/\s+/g, '');
    if (!user || !pass) {
      res.status(500).json({ ok: false, error: 'Credenciais de e-mail não configuradas no servidor' });
      return;
    }

    const remetente = String(process.env.REMETENTE_NOME || 'Escola de Idiomas Exemplo').trim();
    const primeiroNome = nome.split(/\s+/)[0] || 'aluno(a)';
    const mesLabel = mes + (ano ? ' de ' + ano : '');
    const refText = mes ? ` referente ao mês de ${mesLabel}` : '';
    const refHtml = mes ? ` referente ao mês de <b>${esc(mesLabel)}</b>` : '';
    const nomeArq = (nome.replace(/[^\p{L}\p{N}\-_. ]/gu, '').trim()) || 'aluno';

    // O distrato é um documento a ser assinado, não um comprovante: o texto do
    // e-mail pede a devolução assinada em vez de só anunciar o anexo.
    const mensagem = {
      from: `"${remetente.replace(/"/g, '')}" <${user}>`,
      to,
      replyTo: user,
      subject: distrato
        ? 'Termo de Distrato — ' + remetente
        : 'Recibo de Pagamento' + (mes ? ' - ' + mesLabel : '') + ' — ' + remetente,
      text: distrato
        ? `Olá, ${primeiroNome}!\n\n` +
          `Segue em anexo o Termo de Distrato referente ao cancelamento do seu contrato, ` +
          `com o demonstrativo dos valores apurados.\n\n` +
          `Pedimos que revise o documento e nos devolva uma via assinada respondendo a este e-mail.\n\n` +
          `Qualquer dúvida, é só responder aqui.\n\n` +
          `Atenciosamente,\n${remetente}`
        : `Olá, ${primeiroNome}!\n\n` +
          `Segue em anexo o seu recibo de pagamento${refText}.\n\n` +
          `Qualquer dúvida, é só responder este e-mail.\n\n` +
          `Atenciosamente,\n${remetente}`,
      html: distrato
        ? `<p>Olá, <b>${esc(primeiroNome)}</b>!</p>` +
          `<p>Segue em anexo o <b>Termo de Distrato</b> referente ao cancelamento do seu contrato, ` +
          `com o demonstrativo dos valores apurados.</p>` +
          `<p>Pedimos que revise o documento e nos devolva uma <b>via assinada</b> respondendo a este e-mail.</p>` +
          `<p>Qualquer dúvida, é só responder aqui.</p>` +
          `<p>Atenciosamente,<br><b>${esc(remetente)}</b></p>`
        : `<p>Olá, <b>${esc(primeiroNome)}</b>!</p>` +
          `<p>Segue em anexo o seu <b>recibo de pagamento</b>${refHtml}.</p>` +
          `<p>Qualquer dúvida, é só responder este e-mail.</p>` +
          `<p>Atenciosamente,<br><b>${esc(remetente)}</b></p>`,
      attachments: [
        {
          filename: `${distrato ? 'Termo de Distrato' : 'Recibo'} - ${nomeArq}.pdf`,
          content: buf,
          contentType: 'application/pdf',
        },
      ],
    };

    try {
      await criarTransporte(user, pass).sendMail(mensagem);
    } catch (err) {
      if (!ehErroDeConexao(err)) throw err;
      await criarTransporte(user, pass, true).sendMail(mensagem);
    }

    res.status(200).json({ ok: true });
  } catch (err) {
    // O detalhe completo fica só no log da função; ao cliente vai apenas o
    // código do erro SMTP (EAUTH, ETIMEDOUT, EENVELOPE…), que não é sensível
    // e é o que permite diagnosticar a falha sem abrir o painel da Vercel.
    console.error('Falha no envio de e-mail:', err);
    const code = String((err && (err.code || err.responseCode)) || 'DESCONHECIDO').slice(0, 40);
    const dica = code === 'EAUTH'
      ? 'Verifique SMTP_USER e a Senha de App (SMTP_PASS) na Vercel.'
      : (ehErroDeConexao(err) ? 'Não foi possível conectar ao servidor SMTP.' : '');
    res.status(500).json({ ok: false, error: 'Erro ao enviar (' + code + ')' + (dica ? ' — ' + dica : '') });
  }
};
