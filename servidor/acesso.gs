// © 2026 Comanda Aí. Todos os direitos reservados. Código proprietário (veja LICENSE).
/**
 * Comanda Aí — controle de acesso por convite (grátis, no Google Planilhas).
 *
 * Como instalar (uma vez só, ~5 min — dá para fazer pelo celular):
 * 1. Abra script.google.com (no celular: Safari → aA → "Solicitar site para computador").
 * 2. Novo projeto. Apague o que estiver lá e cole este arquivo inteiro. Salve.
 * 3. No topo, escolha a função "configurar" e toque em Executar. Autorize
 *    (Avançado → Acessar projeto → Permitir). A planilha "Comanda Aí — Acessos" é criada no seu Drive.
 * 4. Implantar → Nova implantação → engrenagem → "App da Web".
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 * 5. Copie a "URL do app da Web" (termina em /exec) e coloque no app.
 *
 * (Também funciona colado em Extensões → Apps Script de uma planilha já existente.)
 *
 * Como funciona:
 * - Cada pedido vira uma linha na aba "Acessos" e chega um e-mail com LIBERAR e RECUSAR.
 * - Ao liberar, começa o teste grátis de 20 dias (coluna "vence_em", plano "teste").
 * - Todo dia às 9h chega um e-mail com quem vence em até 3 dias, com o botão "Renovar +30 dias".
 * - Quando vence, o app do bar mostra "teste terminou" (os dados dele continuam no celular).
 * - Para renovar à mão: mude a data em "vence_em" (e o plano para "pago"). O app libera sozinho.
 * - Para acesso sem prazo (ex.: o seu celular): deixe "vence_em" vazio.
 * - Para bloquear: mude "status" para bloqueado. Para liberar de novo: liberado
 * - "ultimo_acesso" mostra quando cada aparelho abriu o app pela última vez.
 *
 * Atualizar o servidor depois de colar uma versão nova deste arquivo:
 *   Salvar → rodar "configurar" uma vez → Implantar → Gerenciar implantações → lápis ✏️ →
 *   Versão: "Nova versão" → Implantar. (A URL /exec continua a mesma.)
 */

const ABA = 'Acessos';
const COLS = ['id', 'pedido_em', 'bar', 'nome', 'whatsapp', 'email', 'status', 'token', 'ultimo_acesso', 'aparelho', 'vence_em', 'plano'];
const DIAS_TESTE = 20;        // teste grátis a partir da liberação
const MAX_PEDIDOS_DIA = 30;   // acima disso no mesmo dia, o pedido é guardado mas não manda e-mail (proteção contra abuso)
const DIA = 24 * 3600 * 1000;
const VERSAO = '4';

// Lê a data da coluna vence_em mesmo se ela foi digitada como texto ("21/11/2026", "21/11/26", "2026-11-21").
function dataDe_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const t = String(v || '').trim();
  let m = t.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (m) { const y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return new Date(y, +m[2] - 1, +m[1]); }
  m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}
const C = Object.fromEntries(COLS.map((c, i) => [c, i]));

// Funciona colado dentro de uma planilha OU como projeto avulso em script.google.com:
// no projeto avulso, a planilha "Comanda Aí — Acessos" é criada sozinha no seu Google Drive.
function planilha_() {
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) return ativa;
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('PLANILHA_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const nova = SpreadsheetApp.create('Comanda Aí — Acessos');
  p.setProperty('PLANILHA_ID', nova.getId());
  return nova;
}

// Rode esta função uma vez pelo botão "Executar" para autorizar e criar a planilha.
function configurar() {
  const ss = planilha_(); aba_(); chave_();
  ss.setSpreadsheetLocale('pt_BR'); ss.setSpreadsheetTimeZone('America/Sao_Paulo');
  // e-mail diário às 9h com quem vence nos próximos dias
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'resumoDiario'))
    ScriptApp.newTrigger('resumoDiario').timeBased().everyDays(1).atHour(9).create();
  Logger.log('Planilha pronta: ' + ss.getUrl());
}

function aba_() {
  const ss = planilha_();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
  }
  // planilhas criadas pela versão anterior ganham as colunas novas
  if (sh.getLastColumn() < COLS.length) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.getRange(2, C.vence_em + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('dd/MM/yyyy');
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
    let status = String(r[C.status]).trim().toLowerCase();
    const vence = dataDe_(r[C.vence_em]);
    // você renovou mudando só a data na planilha: volta a liberar sozinho
    if (status === 'vencido' && vence && Date.now() <= fimDoDia_(vence)) {
      status = 'liberado';
      sh.getRange(row, C.status + 1).setValue('liberado');
    }
    if (status === 'liberado' && vence && Date.now() > fimDoDia_(vence)) {
      status = 'vencido';
      sh.getRange(row, C.status + 1).setValue('vencido');
    }
    const dias = vence ? Math.ceil((fimDoDia_(vence) - Date.now()) / DIA) : null;
    return json_({ status, vence: vence ? vence.toISOString() : null, dias, plano: String(r[C.plano] || '') });
  }
  // 2) você clica em Liberar / Recusar / Bloquear no e-mail
  const acoes = { liberar: 'liberado', recusar: 'recusado', bloquear: 'bloqueado', renovar: 'liberado' };
  if (acoes[p.action]) {
    if (p.k !== chave_()) return pagina_('Link inválido', 'Este link não é válido.', '⛔');
    const row = acharLinha_(sh, p.id);
    if (row < 0) return pagina_('Pedido não encontrado', 'Ele pode ter sido apagado da planilha.', '❓');
    const r = sh.getRange(row, 1, 1, COLS.length).getValues()[0];
    sh.getRange(row, C.status + 1).setValue(acoes[p.action]);
    // liberar: começa o teste grátis (se ainda não tem data). renovar: +30 dias a partir do vencimento (ou de hoje)
    if (p.action === 'liberar' && !dataDe_(r[C.vence_em])) {
      sh.getRange(row, C.vence_em + 1).setValue(new Date(Date.now() + DIAS_TESTE * DIA)).setNumberFormat('dd/MM/yyyy');
      sh.getRange(row, C.plano + 1).setValue('teste');
    }
    if (p.action === 'renovar') {
      const atual = dataDe_(r[C.vence_em]);
      const base = atual && atual.getTime() > Date.now() ? atual.getTime() : Date.now();
      sh.getRange(row, C.vence_em + 1).setValue(new Date(base + 30 * DIA)).setNumberFormat('dd/MM/yyyy');
      sh.getRange(row, C.plano + 1).setValue('pago');
    }
    if (p.action === 'liberar' && r[C.email]) {
      try {
        MailApp.sendEmail({ to: String(r[C.email]), name: 'Comanda Aí', subject: 'Seu acesso ao Comanda Aí foi liberado ✅',
          htmlBody: `Olá, ${esc_(r[C.nome])}!<br><br>O acesso do <b>${esc_(r[C.bar])}</b> ao Comanda Aí foi liberado, com <b>${DIAS_TESTE} dias de teste grátis</b>. É só abrir o app no mesmo celular em que você fez o pedido.<br><br>Boas vendas! 🍻` });
      } catch (err) {}
    }
    const msg = { liberar: ['Acesso liberado', `${esc_(r[C.nome])} (${esc_(r[C.bar])}) já pode usar o app.`, '✅'],
                  recusar: ['Pedido recusado', `${esc_(r[C.nome])} (${esc_(r[C.bar])}) não terá acesso.`, '🚫'],
                  renovar: ['Renovado por 30 dias', `${esc_(r[C.nome])} (${esc_(r[C.bar])}) já pode continuar usando. O app atualiza sozinho.`, '💳'],
                  bloquear: ['Acesso bloqueado', `O app de ${esc_(r[C.nome])} (${esc_(r[C.bar])}) será bloqueado na próxima vez que abrir com internet.`, '🔒'] }[p.action];
    return pagina_(...msg);
  }
  return json_({ ok: true, app: 'Comanda Aí — acessos', versao: VERSAO });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'interesse') return json_(interesse_(b));
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
    sh.appendRow([id, new Date(), bar, nome, whatsapp, email, 'pendente', token, '', aparelho, '', '']);
    const hoje = dados.filter(r => r[C.pedido_em] instanceof Date && Date.now() - r[C.pedido_em].getTime() < DIA).length;
    if (hoje >= MAX_PEDIDOS_DIA) return json_({ ok: true, id, token, status: 'pendente' }); // fica na planilha, sem e-mail

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

function fimDoDia_(d) { const x = new Date(d); x.setHours(23, 59, 59, 999); return x.getTime(); }

// Roda sozinho todo dia às 9h (criado pela função configurar): avisa quem vence em até 3 dias e quem já venceu.
function resumoDiario() {
  const sh = aba_();
  if (sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues();
  const url = ScriptApp.getService().getUrl(), k = chave_();
  const linha = r => {
    const dias = Math.ceil((fimDoDia_(dataDe_(r[C.vence_em])) - Date.now()) / DIA);
    const quando = dias < 0 ? `venceu há ${-dias} dia(s)` : dias === 0 ? 'vence hoje' : `vence em ${dias} dia(s)`;
    return `<tr><td style="padding:6px 10px"><b>${esc_(r[C.bar])}</b><br><span style="color:#6b7280">${esc_(r[C.nome])} · ${esc_(r[C.whatsapp])} · ${esc_(r[C.plano] || '')}</span></td>
      <td style="padding:6px 10px">${quando}</td>
      <td style="padding:6px 10px"><a href="${url}?action=renovar&id=${r[C.id]}&k=${k}" style="background:#16a34a;color:#fff;padding:8px 12px;border-radius:8px;text-decoration:none;font-weight:700">Renovar +30 dias</a></td></tr>`;
  };
  const alvo = rows.filter(r => dataDe_(r[C.vence_em]) && ['liberado', 'vencido'].includes(String(r[C.status]).trim().toLowerCase())
    && (fimDoDia_(dataDe_(r[C.vence_em])) - Date.now()) / DIA <= 3 && (Date.now() - fimDoDia_(dataDe_(r[C.vence_em]))) / DIA <= 7);
  if (!alvo.length) return;
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), name: 'Comanda Aí',
    subject: `⏰ Comanda Aí: ${alvo.length} acesso(s) vencendo ou vencidos`,
    htmlBody: `<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;font-size:15px"><h2>Vencimentos</h2>
      <p>Quando o bar pagar, toque em <b>Renovar +30 dias</b>. O app dele libera sozinho.</p>
      <table style="border-collapse:collapse">${alvo.map(linha).join('')}</table></div>` });
}

// Teste de interesse ("porta falsa"): o bar toca em "Quero isso" num recurso do Pro que ainda não existe.
// Cada clique vira uma linha na aba "Interesse" e um e-mail para você.
const COLS_INT = ['data', 'recurso', 'bar', 'nome', 'whatsapp', 'email', 'atendentes', 'pagaria', 'aparelho'];
function interesse_(b) {
  const campo = (v, n) => String(v || '').trim().slice(0, n);
  const d = { recurso: campo(b.recurso, 60), bar: campo(b.bar, 80), nome: campo(b.nome, 80), whatsapp: campo(b.whatsapp, 30),
              email: campo(b.email, 120), atendentes: campo(b.atendentes, 20), pagaria: campo(b.pagaria, 30), aparelho: campo(b.aparelho, 60) };
  if (!d.recurso || !d.whatsapp) return { ok: false, error: 'Faltou o WhatsApp.' };
  const ss = planilha_();
  let sh = ss.getSheetByName('Interesse');
  if (!sh) { sh = ss.insertSheet('Interesse'); sh.appendRow(COLS_INT); sh.setFrozenRows(1); sh.getRange(1, 1, 1, COLS_INT.length).setFontWeight('bold'); }
  // mesmo aparelho e mesmo recurso no mesmo dia: atualiza em vez de duplicar, e não manda outro e-mail
  const linhas = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, COLS_INT.length).getValues() : [];
  const i = linhas.findIndex(r => r[8] === d.aparelho && r[1] === d.recurso && r[0] instanceof Date && Date.now() - r[0].getTime() < DIA);
  const linha = [new Date(), d.recurso, d.bar, d.nome, d.whatsapp, d.email, d.atendentes, d.pagaria, d.aparelho];
  if (i >= 0) { sh.getRange(i + 2, 1, 1, COLS_INT.length).setValues([linha]); return { ok: true }; }
  sh.appendRow(linha);
  const hoje = linhas.filter(r => r[0] instanceof Date && Date.now() - r[0].getTime() < DIA).length;
  if (hoje < MAX_PEDIDOS_DIA) {
    const zap = d.whatsapp.replace(/\D/g, '');
    MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), name: 'Comanda Aí',
      subject: `💡 Interesse no Pro: ${d.recurso} — ${d.bar}`,
      htmlBody: `<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;font-size:15px">
        <h2 style="margin:0 0 12px">Alguém quer um recurso do Pro!</h2>
        <table cellpadding="6" style="border-collapse:collapse">
          <tr><td><b>Recurso</b></td><td>${esc_(d.recurso)}</td></tr>
          <tr><td><b>Bar</b></td><td>${esc_(d.bar)}</td></tr>
          <tr><td><b>Nome</b></td><td>${esc_(d.nome)}</td></tr>
          <tr><td><b>WhatsApp</b></td><td>${esc_(d.whatsapp)}</td></tr>
          ${d.atendentes ? `<tr><td><b>Atendentes</b></td><td>${esc_(d.atendentes)}</td></tr>` : ''}
          <tr><td><b>Pagaria</b></td><td>${esc_(d.pagaria) || '—'}</td></tr>
        </table><br>
        <a href="https://wa.me/${zap.length <= 11 ? '55' + zap : zap}" style="display:inline-block;padding:12px 22px;border-radius:12px;background:#25d366;color:#05260f;text-decoration:none;font-weight:700">Chamar no WhatsApp</a>
        <p style="color:#6b7280;font-size:13px">Tudo fica na aba "Interesse" da planilha.</p></div>` });
  }
  return { ok: true };
}
