/**
 * Comanda Aí — controle de acesso por convite (grátis, no Google Planilhas).
 *
 * Como instalar (uma vez só, ~5 min):
 * 1. Crie uma planilha nova no Google Sheets (ex.: "Comanda Aí — Acessos").
 * 2. Menu Extensões → Apps Script. Apague o que estiver lá e cole este arquivo inteiro.
 * 3. Clique em Salvar (ícone de disquete).
 * 4. Implantar → Nova implantação → engrenagem → "App da Web".
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 *    Clique em Implantar e autorize (Avançado → Acessar projeto, se aparecer aviso).
 * 5. Copie a "URL do app da Web" (termina em /exec) e mande para colocar no app.
 *
 * Como funciona:
 * - Cada pedido de acesso vira uma linha na aba "Acessos" e você recebe um e-mail
 *   com os botões LIBERAR e RECUSAR.
 * - Para bloquear alguém depois, mude a coluna "status" para: bloqueado
 *   (ou use o link "Bloquear" do e-mail). Para liberar de novo: liberado
 * - A coluna "ultimo_acesso" mostra quando cada aparelho abriu o app pela última vez.
 */

const ABA = 'Acessos';
const COLS = ['id', 'pedido_em', 'bar', 'nome', 'whatsapp', 'email', 'status', 'token', 'ultimo_acesso', 'aparelho'];
const C = Object.fromEntries(COLS.map((c, i) => [c, i]));

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
  }
  return sh;
}

// chave secreta dos links de liberar/recusar (criada sozinha, fica só na sua conta Google)
function chave_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('CHAVE_ADMIN');
  if (!k) { k = Utilities.getUuid().replace(/-/g, ''); p.setProperty('CHAVE_ADMIN', k); }
  return k;
}

const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const pagina_ = (titulo, texto, cor) => HtmlService.createHtmlOutput(
  `<meta name="viewport" content="width=device-width,initial-scale=1">
   <div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:420px;margin:60px auto;padding:28px;border-radius:20px;background:#11131c;color:#f4f6fb;text-align:center">
   <div style="font-size:56px">${cor}</div><h2 style="margin:10px 0">${titulo}</h2><p style="color:#8a90a3;line-height:1.5">${texto}</p></div>`
).setTitle('Comanda Aí');
const esc_ = s => String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function acharLinha_(sh, id) {
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().flat().map(String);
  const i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

function doGet(e) {
  const p = e.parameter || {};
  const sh = aba_();
  // 1) o app pergunta o status deste aparelho
  if (p.action === 'status') {
    const row = acharLinha_(sh, p.id);
    if (row < 0) return json_({ status: 'desconhecido' });
    const r = sh.getRange(row, 1, 1, COLS.length).getValues()[0];
    if (String(r[C.token]) !== String(p.token)) return json_({ status: 'desconhecido' });
    sh.getRange(row, C.ultimo_acesso + 1).setValue(new Date());
    return json_({ status: String(r[C.status]).trim().toLowerCase() });
  }
  // 2) você clica em Liberar / Recusar / Bloquear no e-mail
  const acoes = { liberar: 'liberado', recusar: 'recusado', bloquear: 'bloqueado' };
  if (acoes[p.action]) {
    if (p.k !== chave_()) return pagina_('Link inválido', 'Este link não é válido.', '⛔');
    const row = acharLinha_(sh, p.id);
    if (row < 0) return pagina_('Pedido não encontrado', 'Ele pode ter sido apagado da planilha.', '❓');
    const r = sh.getRange(row, 1, 1, COLS.length).getValues()[0];
    sh.getRange(row, C.status + 1).setValue(acoes[p.action]);
    if (p.action === 'liberar' && r[C.email]) {
      try {
        MailApp.sendEmail({ to: String(r[C.email]), name: 'Comanda Aí', subject: 'Seu acesso ao Comanda Aí foi liberado ✅',
          htmlBody: `Olá, ${esc_(r[C.nome])}!<br><br>O acesso do <b>${esc_(r[C.bar])}</b> ao Comanda Aí foi liberado. É só abrir o app no mesmo celular em que você fez o pedido.<br><br>Boas vendas! 🍻` });
      } catch (err) {}
    }
    const msg = { liberar: ['Acesso liberado', `${esc_(r[C.nome])} (${esc_(r[C.bar])}) já pode usar o app.`, '✅'],
                  recusar: ['Pedido recusado', `${esc_(r[C.nome])} (${esc_(r[C.bar])}) não terá acesso.`, '🚫'],
                  bloquear: ['Acesso bloqueado', `O app de ${esc_(r[C.nome])} (${esc_(r[C.bar])}) será bloqueado na próxima vez que abrir com internet.`, '🔒'] }[p.action];
    return pagina_(...msg);
  }
  return json_({ ok: true, app: 'Comanda Aí — acessos' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action !== 'solicitar') return json_({ ok: false, error: 'ação desconhecida' });
    const campo = (v, n) => String(v || '').trim().slice(0, n);
    const bar = campo(b.bar, 80), nome = campo(b.nome, 80), whatsapp = campo(b.whatsapp, 30), email = campo(b.email, 120), aparelho = campo(b.aparelho, 60);
    if (!bar || !nome || !whatsapp) return json_({ ok: false, error: 'Preencha bar, nome e WhatsApp.' });
    const sh = aba_();
    // mesmo aparelho pedindo de novo: reaproveita o pedido pendente e não manda outro e-mail
    const dados = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues() : [];
    const ant = dados.find(r => String(r[C.aparelho]) === aparelho && String(r[C.status]) === 'pendente');
    if (ant) return json_({ ok: true, id: String(ant[C.id]), token: String(ant[C.token]), status: 'pendente' });

    const id = Utilities.getUuid().slice(0, 8), token = Utilities.getUuid().replace(/-/g, '');
    sh.appendRow([id, new Date(), bar, nome, whatsapp, email, 'pendente', token, '', aparelho]);

    const url = ScriptApp.getService().getUrl(), k = chave_();
    const link = a => `${url}?action=${a}&id=${id}&k=${k}`;
    const botao = (a, txt, cor) => `<a href="${link(a)}" style="display:inline-block;padding:14px 26px;margin:6px;border-radius:12px;background:${cor};color:#fff;text-decoration:none;font-weight:700">${txt}</a>`;
    MailApp.sendEmail({
      to: Session.getEffectiveUser().getEmail(), name: 'Comanda Aí',
      subject: `🔔 Pedido de acesso: ${bar} (${nome})`,
      htmlBody: `<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;font-size:15px">
        <h2 style="margin:0 0 12px">Novo pedido de acesso ao Comanda Aí</h2>
        <table cellpadding="6" style="border-collapse:collapse">
          <tr><td><b>Bar</b></td><td>${esc_(bar)}</td></tr>
          <tr><td><b>Nome</b></td><td>${esc_(nome)}</td></tr>
          <tr><td><b>WhatsApp</b></td><td>${esc_(whatsapp)}</td></tr>
          <tr><td><b>E-mail</b></td><td>${esc_(email) || '—'}</td></tr>
        </table><br>
        ${botao('liberar', '✅ LIBERAR', '#16a34a')} ${botao('recusar', 'Recusar', '#6b7280')}
        <p style="color:#6b7280;font-size:13px;margin-top:18px">Depois de liberado, para bloquear: <a href="${link('bloquear')}">bloquear este acesso</a>, ou mude o status na planilha.</p></div>`
    });
    return json_({ ok: true, id, token, status: 'pendente' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
