/**
 * ICliGo — integração Apps Script / Google Sheets.
 * Ficheiro preparado para substituir o Code.gs do projeto Apps Script.
 *
 * Depois de colar: Implementar > Gerir implementações > Editar >
 * Nova versão > Implementar, mantendo "Quem tem acesso: Qualquer pessoa".
 *
 * O endpoint POST escreve a lead e atribui um ID único.
 * O endpoint GET?action=check confirma se esse ID consta efetivamente do Sheet.
 * Não expõe nome, email ou telefone em verificações públicas.
 */
const ICLIGO_SHEET_ID = '1EZ-KBxL3970VUtYh00IL6Oy9jue-Ba-WDTbcITTlUgE';
const ICLIGO_TAB = 'Leads';
const ICLIGO_ID_COL = 21; // Coluna U, sem alterar a fórmula na aba Prioridades

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'check') {
    return ContentService.createTextOutput('ICliGo Leads API ativa');
  }
  const id = String(p.submissionId || '');
  const callback = String(p.callback || '');
  // JSONP é necessário porque Apps Script não disponibiliza CORS configurável.
  if (!/^[A-Za-z_$][\w$]*$/.test(callback) ||
      !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) {
    return ContentService.createTextOutput('Pedido inválido.');
  }
  let saved = false;
  try {
    const sheet = SpreadsheetApp.openById(ICLIGO_SHEET_ID).getSheetByName(ICLIGO_TAB);
    if (!sheet) throw new Error('Aba Leads não encontrada');
    saved = icligoExists_(sheet, id);
  } catch (error) {
    console.error('Erro na verificação de lead: ' + error);
  }
  return ContentService.createTextOutput(
    callback + '(' + JSON.stringify({ success: true, saved: saved, submissionId: id }) + ');'
  ).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  try {
    const raw = e && e.parameter && e.parameter.data;
    if (!raw) throw new Error('Falta o campo data');
    const lead = JSON.parse(raw);
    const result = icligoSave_(lead);
    return icligoJson_({ success: true, saved: true, submissionId: result.id, duplicate: result.duplicate });
  } catch (error) {
    console.error('Erro ao guardar lead: ' + error);
    return icligoJson_({ success: false, saved: false, error: 'Não foi possível guardar a lead' });
  }
}

function icligoSave_(lead) {
  if (!lead || typeof lead !== 'object') throw new Error('Lead inválida');
  const name = String(lead.name || '').trim().slice(0, 150);
  const email = String(lead.email || '').trim().toLowerCase().slice(0, 254);
  const whatsapp = String(lead.whatsapp || '').replace(/\D/g, '');
  const id = String(lead.submissionId || '').trim();
  const answers = Array.isArray(lead.answers) ? lead.answers : [];

  if (name.length < 2) throw new Error('Nome inválido');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email inválido');
  if (whatsapp.length < 7 || whatsapp.length > 15) throw new Error('Telefone inválido');
  if (answers.length !== 5) throw new Error('Quiz incompleto');
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) {
    throw new Error('ID de envio inválido');
  }
  const scores = answers.map(function(a) {
    const n = Number(a && a.score);
    if (!Number.isInteger(n) || n < 0 || n > 2) throw new Error('Pontuação inválida');
    return n;
  });
  const score = scores.reduce(function(s, n) { return s + n; }, 0);
  // Escala ponderada existente: 90 para [2,1,2,2,2]; 65 para [1,1,2,2,1].
  const weighted = scores[0] * 5 + scores[1] * 10 + scores[2] * 10 +
                   scores[3] * 5 + scores[4] * 20;
  const profile = score >= 8 ? 'FORTE ALINHAMENTO' :
                  score >= 4 ? 'INTERESSE A EXPLORAR' : 'FASE DE EXPLORAÇÃO';
  const priority = weighted >= 80 ? '🔥 PRIORIDADE 1' :
                   weighted >= 60 ? '🟠 PRIORIDADE 2' :
                   weighted >= 40 ? '🟡 PRIORIDADE 3' : '⚪ PRIORIDADE 4';
  const replies = answers.map(function(a, i) {
    return (i + 1) + '. ' + String(a.question || '').slice(0, 300) +
           ' → ' + String(a.answer || '').slice(0, 500);
  }).join(' | ');

  // Evita escritas simultâneas e duplicação em reenvios.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.openById(ICLIGO_SHEET_ID).getSheetByName(ICLIGO_TAB);
    if (!sheet) throw new Error('Aba Leads não encontrada');
    if (sheet.getRange(1, ICLIGO_ID_COL).getValue() !== 'Submission ID') {
      sheet.getRange(1, ICLIGO_ID_COL).setValue('Submission ID');
    }
    if (icligoExists_(sheet, id)) return { id: id, duplicate: true };
    const date = Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy/MM/dd HH:mm:ss');
    const country = String(lead.country || '').slice(0, 100);
    const countryIso = String(lead.countryIso || '').toUpperCase().slice(0, 3);
    // Prefixa =+-@ para que texto fornecido por visitantes não crie fórmulas.
    const safe = function(x) { return /^[=+\-@]/.test(x) ? "'" + x : x; };
    sheet.appendRow([
      date, safe(name), whatsapp, safe(email), score, profile, safe(replies),
      scores[0], scores[1], scores[2], scores[3], scores[4],
      weighted, priority, 'Novo', '', '', '', safe(country), safe(countryIso), id
    ]);
    SpreadsheetApp.flush();
    return { id: id, duplicate: false };
  } finally {
    lock.releaseLock();
  }
}

function icligoExists_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const matches = sheet.getRange(2, ICLIGO_ID_COL, lastRow - 1, 1)
    .createTextFinder(id).matchEntireCell(true).findNext();
  return matches !== null;
}

function icligoJson_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
