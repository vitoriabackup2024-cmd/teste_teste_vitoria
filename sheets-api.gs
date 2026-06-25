// ═══════════════════════════════════════════════════════════════
// finança. · Vitória — Apps Script (Google Sheets backend)
// Cole este código em Extensões → Apps Script da sua planilha
// Publique como Web App: executar como Você, acesso Qualquer Pessoa
// ═══════════════════════════════════════════════════════════════

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ── Estrutura das abas ──────────────────────────────────────────
const ABA_LANCAMENTOS = 'lancamentos';
const ABA_POUPANCAS   = 'poupancas';
const ABA_CONFIG      = 'config';

const HEADER_LANC = ['data', 'tipo', 'cat', 'valor', 'desc'];
const HEADER_POUP = ['key', 'saldo'];
const HEADER_CFG  = ['key', 'value'];

// ── doGet — leitura ─────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || 'getAll';
  let result;

  try {
    if (action === 'getAll') {
      result = {
        lancamentos: getLancamentos(),
        poupancas:   getPoupancas(),
        metas:       getMetas(),
      };
    }
    return jsonResponse(result);
  } catch(err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ── doPost — escrita ────────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result   = { ok: true };

    if (action === 'addEntry') {
      addEntry(body.data);
      if (body.data.tipo === 'poupanca') {
        updatePoupanca(body.data.cat, body.data.valor);
      }
    } else if (action === 'updateMeta') {
      setConfig('meta_' + body.key, body.value);
    } else if (action === 'resgatarViagem') {
      updatePoupanca('viagem', -getPoupancaSaldo('viagem'));
    }

    return jsonResponse(result);
  } catch(err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ── Lançamentos ─────────────────────────────────────────────────
function getLancamentos() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(ABA_LANCAMENTOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_LANCAMENTOS);
    sheet.appendRow(HEADER_LANC);
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    data:  row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(row[0]),
    tipo:  row[1],
    cat:   row[2],
    valor: parseFloat(row[3]) || 0,
    desc:  row[4] || '',
  })).filter(r => r.valor > 0);
}

function addEntry(entry) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(ABA_LANCAMENTOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_LANCAMENTOS);
    sheet.appendRow(HEADER_LANC);
  }
  sheet.appendRow([
    entry.data, entry.tipo, entry.cat || '',
    parseFloat(entry.valor) || 0, entry.desc || ''
  ]);
}

// ── Poupanças ───────────────────────────────────────────────────
function getPoupancas() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(ABA_POUPANCAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_POUPANCAS);
    sheet.appendRow(HEADER_POUP);
    sheet.appendRow(['viagem', 0]);
    sheet.appendRow(['reserva', 0]);
    sheet.appendRow(['chacara', 0]);
    return { viagem: 0, reserva: 0, chacara: 0 };
  }
  const data   = sheet.getDataRange().getValues();
  const result = {};
  data.slice(1).forEach(row => { result[row[0]] = parseFloat(row[1]) || 0; });
  return result;
}

function getPoupancaSaldo(key) {
  const p = getPoupancas();
  return p[key] || 0;
}

function updatePoupanca(key, delta) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(ABA_POUPANCAS);
  if (!sheet) getPoupancas(); // cria a aba
  sheet = ss.getSheetByName(ABA_POUPANCAS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      const novo = Math.max(0, (parseFloat(data[i][1]) || 0) + parseFloat(delta));
      sheet.getRange(i + 1, 2).setValue(novo);
      return;
    }
  }
  sheet.appendRow([key, Math.max(0, parseFloat(delta) || 0)]);
}

// ── Config / Metas ──────────────────────────────────────────────
function getMetas() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(ABA_CONFIG);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_CONFIG);
    sheet.appendRow(HEADER_CFG);
    return { viagem: 3500, reserva: 0, chacara: 0 };
  }
  const data = sheet.getDataRange().getValues();
  const cfg  = {};
  data.slice(1).forEach(row => { cfg[row[0]] = row[1]; });
  return {
    viagem:  parseFloat(cfg['meta_viagem'])  || 3500,
    reserva: parseFloat(cfg['meta_reserva']) || 0,
    chacara: parseFloat(cfg['meta_chacara']) || 0,
  };
}

function setConfig(key, value) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(ABA_CONFIG);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_CONFIG);
    sheet.appendRow(HEADER_CFG);
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

// ── Helpers ─────────────────────────────────────────────────────
function jsonResponse(data, code) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
