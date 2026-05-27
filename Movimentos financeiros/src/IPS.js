//IPS.gs em "Movimentosfinanceiros"
/** CONFIG **/
// Substitua pelo ID da pasta no Google Drive (não é o caminho do PC).
const FOLDER_ID = '1TXL942FE_Z05gSCJ_f_1lj_nEPnD5DWv';
const SHEET_NAME = 'Ficheiros IPS';
const HEADER_ROW = 6;
const START_ROW = HEADER_ROW + 1;
const TZ = 'Europe/Lisbon';

// Ordenação: 'NAME' | 'EMISSAO' | 'DATA' | 'APTSEM'
const ORDER_BY = 'NAME';

const FORMAT_DATA = 'yyyy/m/d';     // B: Data (Drive)
const FORMAT_EMISSAO = 'yyyy-mm-dd';// D: Emissão

function listarPDFsIPS() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`A aba "${SHEET_NAME}" não existe.`);

  // 1) Ler PDFs da pasta (sem subpastas)
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByType(MimeType.PDF);

  const items = [];
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();             // A: Ficheiro
    const last = f.getLastUpdated();      // B: Data (Drive)
    const dataOnly = toDateOnly(last);    // só a data (00:00)

    const tipo = getTipo(name);           // C
    const emissao =
      parseDateFromName(name)                             // YYYY-MM-DD do nome
      || yearPlusMonthDayFromFile(name, dataOnly)         // YYYY do nome + mês/dia do Drive
      || dataOnly;                                        // fallback

    const pagante = /x\.pdf$/i.test(name) ? 'Outro' : 'Associação'; // E
    const aptSem = getAptSem(name);      // F: "NNN/NN"

    items.push({
      name, dataOnly, tipo, emissao, pagante, aptSem,
      primeiro: 0, dups: 0
    });
  }

  // 2) Ordenar antes de escrever
  sortItems(items, ORDER_BY);

  // 3) Calcular Primeiro/Dups apenas para FT-IPS pagas pela Associação
  const grupos = new Map(); // key = `${aptSem}|${yearEmissao}`
  items.forEach((it, i) => {
    if (!it.aptSem) return;
    if (it.pagante !== 'Associação') return;
    if (it.tipo !== 'FT-IPS') return; // <- SÓ FT-IPS contam para dups
    const key = `${it.aptSem}|${it.emissao.getFullYear()}`;
    if (!grupos.has(key)) grupos.set(key, []);
    grupos.get(key).push(i);
  });

  for (const [, idxs] of grupos) {
    // ordenar por Emissão (data) e depois por nome
    idxs.sort((i, j) => {
      const a = items[i], b = items[j];
      const t = a.emissao - b.emissao;
      return t !== 0 ? t : a.name.localeCompare(b.name, 'pt-PT', { numeric: true });
    });
    const dupsCount = Math.max(0, idxs.length - 1);
    if (idxs.length > 0) {
      items[idxs[0]].primeiro = 1;
      items[idxs[0]].dups = dupsCount; // 0 se só 1 emissão; 1 se 2 emissões; etc.
      for (let k = 1; k < idxs.length; k++) {
        items[idxs[k]].primeiro = 0;
        items[idxs[k]].dups = 0;       // zeros nas restantes linhas do grupo
      }
    }
  }

  // 4) Escrever (A:H) a partir de A7
  const lastRow = sh.getLastRow();
  const clearRows = Math.max(0, lastRow - START_ROW + 1);
  if (clearRows > 0) sh.getRange(START_ROW, 1, clearRows, 8).clearContent();

  if (!items.length) return;

  const out = items.map(it => ([
    it.name,           // A Ficheiro
    it.dataOnly,       // B Data (Drive)
    it.tipo || '???',  // C Tipo
    it.emissao,        // D Emissão
    it.pagante,        // E Pagante
    it.aptSem,         // F Apt/sem
    it.primeiro,       // G Primeiro (0/1)
    it.dups            // H Dups
  ]));

  sh.getRange(START_ROW, 1, out.length, 8).setValues(out);

  // 5) Formatos (com proteção a erros do Sheets)
  try { sh.getRange(START_ROW, 2, out.length, 1).setNumberFormat(FORMAT_DATA); } catch (e) {}
  try { sh.getRange(START_ROW, 4, out.length, 1).setNumberFormat(FORMAT_EMISSAO); } catch (e) {}

  // Autoajustes úteis
  sh.autoResizeColumn(1);
  sh.autoResizeColumn(3);
  sh.autoResizeColumn(5);
  sh.autoResizeColumn(6);
  sh.autoResizeColumn(7);
  sh.autoResizeColumn(8);
}

/*** HELPERS ***/

// Ordenação configurável
function sortItems(items, mode) {
  switch (mode) {
    case 'EMISSAO':
      items.sort((a, b) => (a.emissao - b.emissao) || a.name.localeCompare(b.name, 'pt-PT', { numeric: true }));
      break;
    case 'DATA':
      items.sort((a, b) => (a.dataOnly - b.dataOnly) || a.name.localeCompare(b.name, 'pt-PT', { numeric: true }));
      break;
    case 'APTSEM':
      items.sort((a, b) =>
        (a.aptSem || '').localeCompare((b.aptSem || ''), 'pt-PT', { numeric: true })
        || a.name.localeCompare(b.name, 'pt-PT', { numeric: true })
      );
      break;
    case 'NAME':
    default:
      items.sort((a, b) => a.name.localeCompare(b.name, 'pt-PT', { numeric: true }));
  }
}

// Só a data (00:00) mantendo dia/mês/ano
function toDateOnly(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

// Tipo
function getTipo(name) {
  const m = name.match(/\b(FT-IPS|IPS|CP-FT|DRHP)\b/i);
  return m ? m[1].toUpperCase() : '???';
}

// "NNN/NN" a partir de "NNN-NN"
function getAptSem(name) {
  const m = name.match(/\b(\d{3})-(\d{2})\b/);
  return m ? `${m[1]}/${m[2]}` : '';
}

// Emissão completa "YYYY-MM-DD[ x].pdf"
function parseDateFromName(name) {
  const m = name.match(/\s(20\d{2})-(\d{2})-(\d{2})x?\.pdf$/i);
  if (!m) return null;
  const y = +m[1], mo = +m[2], d = +m[3];
  return new Date(y, mo - 1, d);
}

// Emissão só com ano "YYYY[ x].pdf" => usa esse ano + mês/dia da "Data" (Drive)
function yearPlusMonthDayFromFile(name, fileDateOnly) {
  const m = name.match(/\s(20\d{2})x?\.pdf$/i);
  if (!m) return null;
  const y = +m[1];
  return new Date(y, fileDateOnly.getMonth(), fileDateOnly.getDate());
}
