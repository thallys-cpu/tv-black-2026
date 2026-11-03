/**
 * TV da Black 2026 — entrega as vendas da aba tv_vendas já filtradas, sem dado de aluno.
 * Publicar: Implantar > Nova implantação > Tipo "App da Web"
 *           Executar como: Eu  ·  Quem pode acessar: Qualquer pessoa
 */
const ABA = 'tv_vendas';
const INICIO = '2026-10-09';
const PROD = {'cfp-2027_55': 'CFP', 'cfpatepassar': 'CFP', 'cfg_2026': 'Trilha CFG', 'cga_2026': 'Trilha CFG', 'cge-2026': 'Trilha CFG', 'cfa_2025': 'CFA L1'};
const VEND = {'edu-': 'Edu Fattore', 'edufattore': 'Edu Fattore', 'br-': 'Bruno Schwarz', 'brunao': 'Bruno Schwarz', 'mn-': 'Medson Oliveira', 'medsono': 'Medson Oliveira',
  'mr-': 'Marco Albrecht', 'marcoa': 'Marco Albrecht', 'nl-': 'Natália', 'nataliap': 'Natália', 'gt-': 'Guilherme Toro', 'guitoro': 'Guilherme Toro'};

function doGet(e) {
  const cache = CacheService.getScriptCache();
  let corpo = cache.get('tv');
  if (!corpo) {
    corpo = JSON.stringify(montaDados_());
    try { cache.put('tv', corpo, 30); } catch (err) {}
  }
  const cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[A-Za-z_$][\w$]{0,40}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + corpo + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(corpo).setMimeType(ContentService.MimeType.JSON);
}

function montaDados_() {
  const sh = SpreadsheetApp.openById('1ujFMSXDLgB0_9oNSihJnCpB4eaDHnPthg5PiOEiMqLc').getSheetByName(ABA);
  const values = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 5).getDisplayValues();
  const cols = values[0];
  const out = [];
  values.slice(1).forEach(function (r) {
    if (!r[0]) return;
    const o = {};
    cols.forEach(function (c, i) { o[c] = r[i] || ''; });
    const valor = Number(String(o.valor).replace(/\./g, '').replace(',', '.')) || 0;
    const pa = PROD[o.curso_a], pb = PROD[o.curso_b], quando = String(o.quando);
    if (!pa || pa !== pb || quando < INICIO) return;
    if (!(valor === 3600 || valor === 4320 || (pa === 'CFP' && (valor === 5000 || valor === 6000)))) return;
    out.push([quando, pa, valor, VEND[o.vendedor] || 'Marketing']);
  });
  out.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
  const at = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm');
  return {at: at, vendas: out};
}

/** Rode uma vez pelo editor para autorizar e conferir. */
function teste() {
  const d = montaDados_();
  Logger.log(d.vendas.length + ' vendas · última: ' + JSON.stringify(d.vendas[d.vendas.length - 1]));
}
