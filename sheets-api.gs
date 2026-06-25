// ═══════════════════════════════════════════════════════════════
// finança. · Vitória — Apps Script v2 (tudo via GET para evitar CORS)
// ═══════════════════════════════════════════════════════════════

const ABA_LANCAMENTOS = 'lancamentos';
const ABA_POUPANCAS   = 'poupancas';
const ABA_CONFIG      = 'config';

// ── doGet — leitura E escrita (evita CORS do POST) ──────────────
function doGet(e) {
  const p      = e.parameter;
  const action = p.action || 'getAll';

  try {
    if (action === 'getAll') {
      return ok({ lancamentos: getLancamentos(), poupancas: getPoupancas(), metas: getMetas() });
    }
    if (action === 'addEntry') {
      const entry = {
        data:  p.data,
        tipo:  p.tipo,
        cat:   p.cat  || '',
        valor: parseFloat(p.valor) || 0,
        desc:  p.desc || ''
      };
      addEntry(entry);
      if (entry.tipo === 'poupanca') updatePoupanca(entry.cat, entry.valor);
      return ok({ saved: true });
    }
    if (action === 'resgatarViagem') {
      updatePoupanca('viagem', -getPoupancaSaldo('viagem'));
      return ok({ ok: true });
    }
    if (action === 'setMeta') {
      setConfig('meta_' + p.key, parseFloat(p.value) || 0);
      return ok({ ok: true });
    }
    return ok({ error: 'acao desconhecida' });
  } catch(err) {
    return ok({ error: err.message });
  }
}

// ── Lançamentos ─────────────────────────────────────────────────
function getLancamentos() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ABA_LANCAMENTOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_LANCAMENTOS);
    sheet.appendRow(['data','tipo','cat','valor','desc']);
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    data:  (function(v) {
             if (v instanceof Date) return Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM-dd');
             var s = String(v);
             // "Thu Jun 25 2026 ..." → extrai com regex
             var m = s.match(/(\w{3}) (\d{1,2}) (\d{4})/);
             if (m) {
               var months = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
               return m[3] + '-' + months[m[1]] + '-' + ('0'+m[2]).slice(-2);
             }
             return s.slice(0,10);
           })(row[0]),
    tipo:  row[1],
    cat:   row[2],
    valor: parseFloat(row[3]) || 0,
    desc:  row[4] || ''
  })).filter(r => r.valor > 0);
}

function addEntry(entry) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ABA_LANCAMENTOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_LANCAMENTOS);
    sheet.appendRow(['data','tipo','cat','valor','desc']);
  }
  sheet.appendRow([entry.data, entry.tipo, entry.cat, entry.valor, entry.desc]);
}

// ── Poupanças ───────────────────────────────────────────────────
function getPoupancas() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ABA_POUPANCAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_POUPANCAS);
    sheet.appendRow(['key','saldo']);
    sheet.appendRow(['viagem',0]);
    sheet.appendRow(['reserva',0]);
    sheet.appendRow(['chacara',0]);
    return { viagem:0, reserva:0, chacara:0 };
  }
  const data = sheet.getDataRange().getValues();
  const r = {};
  data.slice(1).forEach(row => { r[row[0]] = parseFloat(row[1]) || 0; });
  return r;
}

function getPoupancaSaldo(key) { return getPoupancas()[key] || 0; }

function updatePoupanca(key, delta) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ABA_POUPANCAS);
  if (!sheet) { getPoupancas(); sheet = ss.getSheetByName(ABA_POUPANCAS); }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i+1, 2).setValue(Math.max(0, (parseFloat(data[i][1])||0) + parseFloat(delta)));
      return;
    }
  }
  sheet.appendRow([key, Math.max(0, parseFloat(delta)||0)]);
}

// ── Config / Metas ──────────────────────────────────────────────
function getMetas() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ABA_CONFIG);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_CONFIG);
    sheet.appendRow(['key','value']);
    return { viagem:3500, reserva:0, chacara:0 };
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
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ABA_CONFIG);
  if (!sheet) { sheet = ss.insertSheet(ABA_CONFIG); sheet.appendRow(['key','value']); }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i+1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

// ── Helper ──────────────────────────────────────────────────────
function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
