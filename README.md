# Gerador de Recibos PDF

Ferramenta web para emitir recibos de pagamento e termos de distrato em PDF/Word, com envio por e-mail e histórico em Google Sheets.

> Versão de portfólio anonimizada de uma ferramenta desenvolvida para uma escola de idiomas; dados e marcas reais removidos.

## O que faz

- **Recibo de pagamento:** formulário → pré-visualização editável → baixar PDF, baixar Word (.docx), imprimir ou enviar por e-mail ao aluno.
- **Termo de distrato:** calcula multa, reembolso ou saldo devedor conforme a regra de cancelamento (arrependimento em até 7 dias, modalidade parcelada ou recorrente), com ajuste manual para exceções.
- **Histórico:** todo documento gerado é salvo numa planilha Google via Web App do Apps Script; consultar, reabrir e excluir exige senha.
- **Automações extras (`provas/`):** aviso no Slack quando o resultado de uma prova muda e sincronização (upsert por e-mail) entre a planilha de nivelamento e a de matrícula.

## Destaques técnicos

- Página única sem build (HTML/CSS/JS puro); PDF com `html2pdf.js` e `.docx` gerado no navegador (OOXML + JSZip) a partir do DOM já editado.
- Valor por extenso em pt-BR, máscara de CPF, data por extenso.
- Função serverless de e-mail com validação do anexo (tamanho e assinatura `%PDF-`), checagem de origem, escape de HTML e fallback de porta SMTP (465 → 587).
- Backend Apps Script: senha em Script Properties, comparação em tempo constante, bloqueio de 15 min após 5 tentativas, neutralização de injeção de fórmula, cache de leitura e cabeçalho de planilha autoajustável.
- Cabeçalhos de segurança na Vercel (CSP, HSTS, X-Frame-Options, Permissions-Policy).

## Stack

HTML/CSS/JavaScript · html2pdf.js · JSZip · Node.js (Vercel Functions) + Nodemailer · Google Apps Script (Sheets, Drive, Gmail) · Slack Incoming Webhooks

## Como rodar

1. **Dados da empresa:** edite o objeto `EMPRESA` em `recibos/index.html` (os valores atuais são fictícios).
2. **Frontend local:** abra `recibos/index.html` no navegador, ou rode `npx vercel dev` em `recibos/` para testar também o envio de e-mail.
3. **E-mail (Vercel):** `cd recibos && npm install`, faça o deploy e cadastre as variáveis de `recibos/.env.example` (`SMTP_USER`, `SMTP_PASS`, `REMETENTE_NOME`, `ALLOWED_HOSTS`).
4. **Histórico:** crie uma planilha, cole `recibos/historico.gs` em Extensões > Apps Script, defina a propriedade de script `HISTORICO_SENHA` e implante como App da Web. Cole a URL `/exec` no painel "Documentos Gerados" (fica salva só no navegador).
5. **Alternativa por planilha:** `recibos/gerador-recibos.gs` gera recibos a partir de linhas de uma planilha (configure `CFG.PASTA_DRIVE_ID`).
6. **Provas:** nos scripts de `provas/`, preencha o `CONFIG`/`CFG` e rode as funções de setup (`guardarWebhook`, `criarGatilho`, `criarGatilhoPlacement`).

## Estrutura

```
recibos/
  index.html            # app completo (formulário, documentos, PDF/Word, histórico)
  api/enviar.js         # função serverless de envio por e-mail
  historico.gs          # backend do histórico (Apps Script Web App)
  gerador-recibos.gs    # gerador alternativo via planilha + Drive + Gmail
  vercel.json           # cabeçalhos de segurança e config da função
  .env.example          # variáveis de ambiente
provas/
  notificacao-aprovacao-slack.gs
  placement-para-matricula.gs
```

## Direitos autorais

© 2026 Gustavo Fiedler Calegari. Todos os direitos reservados.

Este código está publicado apenas para avaliação de portfólio. Não é permitido copiar, modificar, distribuir ou usar comercialmente, no todo ou em parte, sem autorização por escrito do autor.
