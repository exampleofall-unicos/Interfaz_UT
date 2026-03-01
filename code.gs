/*******************************************************
 * SISTEMA DE REMITOS - UNION TECHNOLOGY (Únicos™)
 * Lógica integrada según documento “Lógica Remitos - UNION TECHNOLOGY” + Modificaciones y Agregados
 * PDFs obligatorios y guardados en Drive (RR/RE)
 *******************************************************/

// ====== HOJAS (nombres exactos) ======
const SHEET_RR = 'Remito de Recepción';
const SHEET_RE = 'Remito de Entrega';
const SHEET_TALLER = 'Equipos en Taller';
const SHEET_CLIENTES = 'Clientes';
const SHEET_RECEPCION = 'Recepción';
const SHEET_INVENTARIO = 'Inventario';
const SHEET_RESERVAS = 'Reservas';
const SHEET_HISTORIAL = 'Historial';
const SHEET_DATOS = 'Datos'; // usado por fórmulas en la RE
const SHEET_DATOS_DESPLEGABLES = 'Datos Desplegables';

const REMITO_HEADER_CANDIDATES = ['N° REMITO','NRO REMITO','NRO. REMITO','N REMITO','NRO REM','REMITO','NRO','N°'];
const ESTADO_ACTUAL_HEADER_CANDIDATES = ['ESTADO ACTUAL','ESTADO'];

// ====== CARPETAS PDF ======
const FOLDER_RR = '1c_qdZwxwmfmFa7S1ySqMTT-D_IzKZsLt';
const FOLDER_RE = '1j6S1EvSbg9dcXrM0az7Pqes8uCHvFvJB';

// ====== PDF (A4 vertical, ~1cm márgenes) ======
const PDF_EXPORT = { size:'A4', portrait:true, top:0.4, bottom:0.4, left:0.4, right:0.4 };

// ====== Rangos críticos RR ======
const RR_NRO='L2', RR_FECHA='L3', RR_CLIENTE='B9', RR_DNI='I9', RR_TEL='K9', RR_DIR='D11';
const RR_MARCA='D14', RR_MODELO='K14', RR_CAP='D15', RR_COLOR='K15', RR_IMEI='D16';
const RR_OBS='D18';                 // OBSERVACIONES (merge D18:L18)
const RR_EMPRESA='K23';             // EMPRESA (merge K23:L23) -> NO debe verse en PDF
const RR_EMPRESA_RANGE='K23:L23';
const RR_RECEP='K25';               // RECEPCIÓN visible en PDF (merge K25:L25)
const RR_RECEP_RANGE='K25:L25';

// ✅ Checkboxes RR
const RR_CHECK_RANGES = ['B22:B26','D22:D26','G22:G25']; // checkboxes
const RR_OTRO_RANGE   = 'D27:L28';                       // "OTRO"
const RR_FECHA_HORA_VIS='L53';
const RE_DOLAR_A1      = 'K53'; // valor dólar Blue (script)
const DOLAR_ROSARIO_UPDATED_AT_KEY = 'DOLAR_ROSARIO_UPDATED_AT';
const DOLAR_ROSARIO_VALUE_KEY = 'DOLAR_ROSARIO_VALUE';

// ====== Rangos críticos RE ======
const RE_NRO='L2', RE_FECHA='L3', RE_CLIENTE='B9:H9', RE_DNI='I9', RE_TEL='K9', RE_DIRECCION='D11';
const RE_MARCA='D14', RE_MODELO='K14', RE_IMEI='D15';
const RE_CAT_COLS=['B','C','D','E']; // lista vertical en B18:E35
const RE_ROW_START=18, RE_ROW_END=35;
const RE_QTY_COL='J'; // cantidades J18:J35
const RE_PRICE_COLS=['K','L']; // combinadas K:L
const RE_EXTRA_ROW_START=32, RE_EXTRA_ROW_END=35; // líneas especiales (extras UI)
const RE_TOTAL='K37:L37';
const RE_COST_TOTAL='M37';      // suma de costos (según tu implementación)
const RE_MONEDA='L16';          // moneda en entrega
const RE_FECHA_HORA_VIS='L42';
const RE_DOLAR_A2      = 'K42'; // valor dólar Blue (script)

// NUEVO: PRECIO USD y FORMA DE PAGO
const RE_PRECIO_USD_CELL = 'I37';
const RE_FORMA_PAGO_RANGE = 'D16:F16';

// ====== Inventario ======
const INV_ROW_INSERT = 6;

// CARGA (fila 2) — asumimos:
// A FECHA, B MOVIMIENTO, C MARCA, D MODELO, E CATEGORÍA, F CANTIDAD, G PRECIO, H TOTAL,
// I COSTO (con fórmula, NO se borra), J TÉCNICO
const INV_CARGA = {
  FECHA:'A2',
  MOV:'B2',
  MARCA:'C2',
  MODELO:'D2',
  CATEG:'E2',
  CANT:'F2',
  PREC:'G2',
  TOTAL:'H2',
  COSTO:'I2',     // ⚠️ TIENE FÓRMULA (se preserva)
  TECNICO:'J2'
};

// ====== Utilidades base ======
const ss = () => SpreadsheetApp.getActiveSpreadsheet();
const sh = n => {
  const s = ss().getSheetByName(n);
  if (!s) throw new Error('No existe la hoja: ' + n);
  return s;
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('UNION TECHNOLOGY')
    .addItem('Generar PDF Recepción', 'generarPDFRecepcion')
    .addItem('Generar PDF Entrega', 'generarPDFEntrega')
    .addSeparator()
    .addItem('Carga de Inventario', 'cargaDeInventario')
    .addSeparator()
    .addItem('Ver / Ocultar Inventario (PIN)', 'inventarioToggleConPIN')
    .addSeparator()
    .addItem('Actualizar Valor Dólar Rosario', 'obtenerDolarBlueRosario')
    .addSeparator()
    .addItem('Abrir Interfaz', 'abrirInterfaz')
    .addToUi();
}

function abrirInterfaz() {
  const html = HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistema de Remitos')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Mostrar/Ocultar Inventario con PIN 1111
const PIN_INVENTARIO = '1111';
function inventarioToggleConPIN(){
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Acceso requerido', 'Ingrese la clave para continuar:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  if (r.getResponseText().trim() !== PIN_INVENTARIO) { ui.alert('Clave incorrecta.'); return; }
  const S = sh(SHEET_INVENTARIO);
  if (S.isSheetHidden()) { S.showSheet(); ui.alert('Hoja "Inventario" visible.'); }
  else { S.hideSheet(); ui.alert('Hoja "Inventario" oculta.'); }
}

function readA1(sheet, a1, display=true){
  return display ? sheet.getRange(a1).getDisplayValue().toString().trim()
                 : sheet.getRange(a1).getValue();
}
function writeA1(sheet, a1, v){ sheet.getRange(a1).setValue(v); }

function getHeaderMap_(s){
  const head = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  const map={}; head.forEach((h,i)=>map[String(h).trim()]=i+1);
  return map;
}
function getCol_(map, names) { for (let i=0;i<names.length;i++){ if (map[names[i]]) return map[names[i]]; } return 0; }

function normalizeStr_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function getColFlexible_(map, keywords) {
  const headerKeys = Object.keys(map);
  const normalizedKeywords = keywords.map(kw => normalizeStr_(kw));
  
  // 1. Intenta coincidencia exacta (case-sensitive)
  for (let kw of keywords) {
    if (map[kw]) return map[kw];
  }
  
  // 2. Intenta case-insensitive exacta (sin acentos)
  for (let hkey of headerKeys) {
    const normalized = normalizeStr_(hkey);
    if (normalizedKeywords.includes(normalized)) {
      return map[hkey];
    }
  }
  
  // 3. Intenta partial match (sin acentos)
  for (let hkey of headerKeys) {
    const normalizedHeader = normalizeStr_(hkey);
    for (let kw of normalizedKeywords) {
      if (normalizedHeader.includes(kw) || kw.includes(normalizedHeader)) {
        return map[hkey];
      }
    }
  }
  
  return 0;
}

function normalizeRemitoId_(value){
  let txt = String(value || '').trim();
  if (!txt) return '';

  txt = txt
    .replace(/^RR\s*#?\s*/i, '')
    .replace(/^RE\s*#?\s*/i, '')
    .trim();

  if (/^\d+(?:\.0+)?$/.test(txt)) {
    return String(Math.trunc(Number(txt)));
  }

  return txt.toUpperCase();
}

function findHeaderCol1Based_(headers, candidates){
  const map = _headerMapFromArray_(headers || []);
  return getColFlexible_(map, candidates || []);
}

function isRecepcionEntregado_(estadoValue){
  const norm = normalizeStr_(estadoValue);
  return norm === 'entregado' || norm.includes('entregado');
}

function buildRecepcionEstadoMapByRemito_(){
  const map = new Map();
  const SR = sh(SHEET_RECEPCION);
  const values = SR.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return map;

  const headers = values[0].map(h => String(h || '').trim());
  const colRemito = findHeaderCol1Based_(headers, REMITO_HEADER_CANDIDATES);
  const colEstado = findHeaderCol1Based_(headers, ESTADO_ACTUAL_HEADER_CANDIDATES);

  if (!colRemito) {
    throw new Error('No se encontró columna de remito en hoja "Recepción" (ej: N° REMITO, NRO REMITO, REMITO).');
  }
  if (!colEstado) {
    throw new Error('No se encontró columna "Estado actual" en hoja "Recepción".');
  }

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const remitoNorm = normalizeRemitoId_(row[colRemito - 1]);
    if (!remitoNorm) continue;
    const estado = String(row[colEstado - 1] || '').trim();
    map.set(remitoNorm, estado);
  }

  return map;
}

function getDropdownOptionsFromSheet_(sheetName, columnIndex, startRow){
  const S = sh(sheetName);
  const firstRow = startRow || 2;
  const lastRow = S.getLastRow();
  if (lastRow < firstRow) return [];
  const values = S.getRange(firstRow, columnIndex, lastRow - firstRow + 1, 1)
    .getDisplayValues()
    .flat();
  const unique = new Set();
  values.forEach(v => {
    const s = String(v || '').trim();
    if (s && s !== '-' && s !== '—') {
      if (!unique.has(s)) unique.add(s);
    }
  });
  return Array.from(unique);
}

function findRowByKey_(s, header, value){
  const values = s.getDataRange().getValues();
  if (values.length<2) return -1;
  const idx = values[0].indexOf(header);
  if (idx===-1) return -1;
  for (let r=1;r<values.length;r++){
    if (values[r][idx]==value) return r+1; // 1-based
  }
  return -1;
}

function maxRemito_(){
  const S = sh(SHEET_HISTORIAL);
  const vals = S.getDataRange().getValues();
  if (vals.length<2) return 0;
  const idx = vals[0].indexOf('N° REMITO');
  if (idx===-1) return 0;
  let m=0;
  for (let r=1;r<vals.length;r++){
    const n = Number(vals[r][idx]);
    if (!isNaN(n)&&n>m) m=n;
  }
  return m;
}

function ensureActiveSheet_(name){
  const cur = ss().getActiveSheet().getName();
  if (cur!==name) throw new Error(`Esta acción solo se puede ejecutar desde "${name}".`);
}

function exportSheetPdf_(sheetName, folderId, fileName){
  const spreadsheetId = ss().getId();
  const gid = sh(sheetName).getSheetId();
  const params = {
    format:'pdf', size:PDF_EXPORT.size, portrait:PDF_EXPORT.portrait,
    top_margin:PDF_EXPORT.top, bottom_margin:PDF_EXPORT.bottom,
    left_margin:PDF_EXPORT.left, right_margin:PDF_EXPORT.right,
    fitw:true, gridlines:false, printnotes:false, sheetnames:false, pagenum:'DISABLED', gid:gid
  };
  const url = 'https://docs.google.com/spreadsheets/d/'+spreadsheetId+'/export?'+
              Object.keys(params).map(k=>k+'='+encodeURIComponent(params[k])).join('&');
  const resp = UrlFetchApp.fetch(url, {headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}, muteHttpExceptions:true});
  if (resp.getResponseCode()!==200) throw new Error('Error al exportar PDF: '+resp.getContentText());
  const blob = resp.getBlob().setName(fileName);
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  return file.getUrl();
}

// ==== Parsing números locales y lectura segura ====
function parseNumberLocal_(str){
  if (typeof str !== 'string') str = String(str||'');
  str = str.trim();
  if (!str) return 0;
  str = str.replace(/[^\d.,-]/g, '');
  if (str.indexOf(',') > -1 && str.indexOf('.') > -1) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (str.indexOf(',') > -1) {
    str = str.replace(',', '.');
  }
  const n = Number(str);
  return isNaN(n) ? 0 : n;
}
function getMergedNumber_(sheet, rangeA1){
  const disp = sheet.getRange(rangeA1).getDisplayValue();
  let n = parseNumberLocal_(disp);
  if (!n){
    const raw = sheet.getRange(rangeA1).getValue();
    n = typeof raw === 'number' ? raw : parseNumberLocal_(String(raw||''));
  }
  return n || 0;
}
function getMergedPairNumber_(sheet, row, cols){
  const disp = sheet.getRange(`${cols[0]}${row}:${cols[1]}${row}`).getDisplayValues()[0];
  let p = 0;
  if (disp[0] && String(disp[0]).trim()!=='') p = parseNumberLocal_(disp[0]);
  else if (disp[1] && String(disp[1]).trim()!=='') p = parseNumberLocal_(disp[1]);
  if (!p){
    const raw = sheet.getRange(`${cols[0]}${row}:${cols[1]}${row}`).getValues()[0];
    if (raw[0] !== '' && raw[0] != null) p = (typeof raw[0] === 'number') ? raw[0] : parseNumberLocal_(String(raw[0]));
    else if (raw[1] !== '' && raw[1] != null) p = (typeof raw[1] === 'number') ? raw[1] : parseNumberLocal_(String(raw[1]));
  }
  return p || 0;
}

/* =========================================================
   ✅ NUEVO: Helpers de validación + capitalización + WhatsApp
   (NO rompen lo actual, solo los usamos en RR/Clientes)
   ========================================================= */

// Capitaliza cada palabra (incluye tildes). Ej: "franco ariel vega" -> "Franco Ariel Vega"
function toTitleCaseEs_(input){
  const s = String(input || '').trim().replace(/\s+/g,' ');
  if (!s) return '';
  return s.split(' ').map(w=>{
    const lower = w.toLocaleLowerCase('es-AR');
    return lower.charAt(0).toLocaleUpperCase('es-AR') + lower.slice(1);
  }).join(' ');
}

// Limpia teléfono y arma link WhatsApp según regla del doc: anteponer https://wa.me/549 + teléfono
function buildWhatsAppLink_(telRaw){
  const digits = String(telRaw || '').replace(/[^\d]/g,''); // solo números
  if (!digits) return '';
  // Evitar duplicados típicos: si ya viene con 549 o 54
  let waDigits = digits;
  if (digits.startsWith('549')) {
    waDigits = digits;
  } else if (digits.startsWith('54')) {
    waDigits = '549' + digits.slice(2);
  } else {
    waDigits = '549' + digits;
  }
  return 'https://wa.me/' + waDigits;
}

// Valida campos obligatorios de RR según documento.
// Devuelve array con etiquetas faltantes.
function validarObligatoriosRR_(SRR){
  const missing = [];

  const cliente = readA1(SRR, RR_CLIENTE);
  const dni     = readA1(SRR, RR_DNI);
  const tel     = readA1(SRR, RR_TEL);

  const marca   = readA1(SRR, RR_MARCA);
  const modelo  = readA1(SRR, RR_MODELO);
  const color   = readA1(SRR, RR_COLOR);

  const empresa = SRR.getRange(RR_EMPRESA_RANGE).getDisplayValue().toString().trim();
  const recep   = SRR.getRange(RR_RECEP_RANGE).getDisplayValue().toString().trim();

  if (!cliente) missing.push('Cliente (B9)');
  if (!dni) missing.push('DNI (I9)');
  if (!tel) missing.push('Teléfono (K9)');
  if (!marca) missing.push('Marca (D14)');
  if (!modelo) missing.push('Modelo (K14)');
  if (!color) missing.push('Color (K15)');
  if (!empresa) missing.push('Empresa (K23:L23)');
  if (!recep) missing.push('Recepción (K25:L25)');

  return missing;
}

// ====== Repuestos solicitados desde RR (SIN GENÉRICOS) ======
function categoriasDesdeRR_(){
  const S = sh(SHEET_RR);
  const items = [];

  // Cada checkbox tiene su texto en la celda a la derecha:
  // B -> C, D -> E, G -> H
  const labelColByCheckboxCol = { 'B': 'C', 'D': 'E', 'G': 'H' };

  RR_CHECK_RANGES.forEach(a1 => {
    const range = S.getRange(a1);
    const vals = range.getValues(); // boolean
    const startRow = range.getRow();
    const colLetter = String.fromCharCode(64 + range.getColumn());

    const labelCol = labelColByCheckboxCol[colLetter];
    if (!labelCol) return;

    for (let i=0; i<vals.length; i++){
      if (vals[i][0] === true){
        const row = startRow + i;
        const label = S.getRange(`${labelCol}${row}`).getDisplayValue().toString().trim();
        if (label) items.push(label);
      }
    }
  });

  // OTRO: D27:L28
  const otro = S.getRange(RR_OTRO_RANGE).getDisplayValue().toString().trim();
  if (otro){
    otro.split(/[,;\n]/).map(x=>x.trim()).filter(Boolean).forEach(x=>items.push(x));
  }

  return items;
}

function countRepuestos_(arr){
  const m = new Map();
  (arr||[]).forEach(x=>{
    const k = String(x||'').trim();
    if (!k) return;
    m.set(k, (m.get(k)||0) + 1);
  });
  return m;
}

// ====== Limpieza de formularios (manteniendo fórmulas) ======
function limpiarRR_(){
  const S=sh(SHEET_RR);
  [RR_CLIENTE,RR_DNI,RR_TEL,RR_DIR,RR_MARCA,RR_MODELO,RR_CAP,RR_COLOR,RR_IMEI].forEach(a=>writeA1(S,a,''));
  S.getRange('D18:L18').clearContent();
  S.getRange(RR_EMPRESA_RANGE).clearContent();
  S.getRange(RR_RECEP_RANGE).clearContent();

  RR_CHECK_RANGES.forEach(a1 => sh(SHEET_RR).getRange(a1).setValue(false));
  S.getRange(RR_OTRO_RANGE).clearContent();
  writeA1(S, RR_NRO, '');
}

/**
 * Mantiene fórmulas RE (blindaje):
 * - K18:L35
 * - K33:L35  (incluida dentro de K18:L35, pero lo dejamos explícito)
 */
function limpiarRE_(){
  const S = sh(SHEET_RE);

  // Backup fórmulas
  const rngKL = S.getRange('K18:L35');
  const formulasKL = rngKL.getFormulas();

  writeA1(S, RE_NRO, '');

  // Inputs:
  S.getRange('B32:E35').clearContent();
  S.getRange('D16:F16').clearContent();
  S.getRange('J16').clearContent();
  S.getRange('L16').clearContent();
  S.getRange('J18:J35').clearContent();

  // Restaurar fórmulas
  rngKL.setFormulas(formulasKL);
}

// ====== Validación de duplicados (Historial/Recepción/Taller) ======
function assertNoDuplicateRemito_(nro){
  const dupH = findRowByKey_(sh(SHEET_HISTORIAL),'N° REMITO',nro)!==-1;
  const dupR = findRowByKey_(sh(SHEET_RECEPCION),'N° REMITO',nro)!==-1;
  const dupT = findRowByKey_(sh(SHEET_TALLER),'N° REMITO',nro)!==-1;
  if (dupH || dupR || dupT) {
    throw new Error(`El N° REMITO ${nro} ya existe en ${[
      dupH?'Historial':null, dupR?'Recepción':null, dupT?'Equipos en Taller':null
    ].filter(Boolean).join(', ')}. No se crearán filas duplicadas.`);
  }
}

// ============== GENERAR PDF RECEPCIÓN (RR) =================
function generarPDFRecepcion(){
  ensureActiveSheet_(SHEET_RR);
  const SRR = sh(SHEET_RR);

  // ✅ NUEVO: validar obligatorios RR (según documento)
  const missing = validarObligatoriosRR_(SRR);
  if (missing.length){
    SpreadsheetApp.getUi().alert(
      'Campos obligatorios incompletos',
      'No se puede generar el PDF de Recepción porque faltan:\n\n• ' + missing.join('\n• ') +
      '\n\nCompletá esos campos y volvé a intentar.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  let nro = Number(readA1(SRR, RR_NRO, true));
  if (!nro){
    nro = maxRemito_()+1;
    writeA1(SRR, RR_NRO, nro);
  }

  assertNoDuplicateRemito_(nro);

  // ✅ NUEVO: capitalizar nombre cliente para bases (no modifica lo que se ve en RR si no querés)
  const clienteRaw = readA1(SRR, RR_CLIENTE);
  const cliente = toTitleCaseEs_(clienteRaw);

  const dni     = readA1(SRR, RR_DNI);
  const tel     = readA1(SRR, RR_TEL);
  const dir     = readA1(SRR, RR_DIR);
  const marca   = readA1(SRR, RR_MARCA);
  const modelo  = readA1(SRR, RR_MODELO);
  const imei    = readA1(SRR, RR_IMEI);

  const obsRecep = SRR.getRange('D18:L18').getDisplayValue().toString().trim();
  const empresa  = SRR.getRange(RR_EMPRESA_RANGE).getDisplayValue().toString().trim();
  const recep    = SRR.getRange(RR_RECEP_RANGE).getDisplayValue().toString().trim();

  const fechaHoy = new Date();
  const fileName = `RR_${nro}_${cliente||'Cliente'}.pdf`;

  const empresaBackup = SRR.getRange(RR_EMPRESA_RANGE).getValues();
  SRR.getRange(RR_EMPRESA_RANGE).clearContent();

  const urlPDF = exportSheetPdf_(SHEET_RR, FOLDER_RR, fileName);

  SRR.getRange(RR_EMPRESA_RANGE).setValues(empresaBackup);

  // HISTORIAL
  {
    const SH = sh(SHEET_HISTORIAL); SH.insertRows(2,1);
    const map = getHeaderMap_(SH);

    if (map['N° REMITO']) SH.getRange(2,map['N° REMITO']).setValue(nro);
    if (map['FECHA INGRESO']) SH.getRange(2,map['FECHA INGRESO']).setValue(fechaHoy);
    if (map['CLIENTE']) SH.getRange(2,map['CLIENTE']).setValue(cliente); // ✅ capitalizado
    if (map['MARCA']) SH.getRange(2,map['MARCA']).setValue(marca);
    if (map['MODELO']) SH.getRange(2,map['MODELO']).setValue(modelo);
    if (map['LINK DE REMITO DE RECEPCIÓN']) SH.getRange(2,map['LINK DE REMITO DE RECEPCIÓN']).setValue(urlPDF);
    if (map['EMPRESA']) SH.getRange(2,map['EMPRESA']).setValue(empresa);
  }

  // TALLER
  {
    const ST = sh(SHEET_TALLER); ST.insertRows(2,1);
    const map = getHeaderMap_(ST);

    if (map['N° REMITO']) ST.getRange(2,map['N° REMITO']).setValue(nro);
    if (map['IMEI']) ST.getRange(2,map['IMEI']).setValue(imei);
    if (map['MARCA']) ST.getRange(2,map['MARCA']).setValue(marca);
    if (map['MODELO']) ST.getRange(2,map['MODELO']).setValue(modelo);
    if (map['CLIENTE']) ST.getRange(2,map['CLIENTE']).setValue(cliente); // ✅ capitalizado
    if (map['INGRESO']) ST.getRange(2,map['INGRESO']).setValue(fechaHoy);

    if (map['ESTADO ACTUAL']) ST.getRange(2,map['ESTADO ACTUAL']).setValue('Pendiente Diagnóstico');

    if (map['EMPRESA']) ST.getRange(2,map['EMPRESA']).setValue(empresa);
    if (map['OBSERVACIONES RECEPCIÓN']) ST.getRange(2,map['OBSERVACIONES RECEPCIÓN']).setValue(obsRecep);
  }

  // CLIENTES
  if (dni){
    const SC = sh(SHEET_CLIENTES);
    const exists = findRowByKey_(SC,'DNI',dni)!==-1;
    if (!exists){
      SC.insertRows(2,1);
      const map = getHeaderMap_(SC);

      if (map['NOMBRE']) SC.getRange(2,map['NOMBRE']).setValue(cliente); // ✅ capitalizado
      if (map['DNI']) SC.getRange(2,map['DNI']).setValue(dni);
      if (map['TELÉFONO']) SC.getRange(2,map['TELÉFONO']).setValue(tel);
      if (map['DIRECCIÓN']) SC.getRange(2,map['DIRECCIÓN']).setValue(dir);

      // ✅ NUEVO: LINK WHATSAPP
      const colWa = getCol_(map, ['LINK WHATSAPP','LINK WHATSAPP ']);
      if (colWa){
        const waLink = buildWhatsAppLink_(tel);
        if (waLink) SC.getRange(2, colWa).setValue(waLink);
      }
    }
  }

  // RECEPCIÓN
  const repuestosArr = categoriasDesdeRR_();
  const repSolicStr = repuestosArr.join('; ');

  {
    const SR = sh(SHEET_RECEPCION); SR.insertRows(2,1);
    const map = getHeaderMap_(SR);

    if (map['N° REMITO']) SR.getRange(2,map['N° REMITO']).setValue(nro);
    if (map['FECHA']) SR.getRange(2,map['FECHA']).setValue(fechaHoy);
    if (map['CLIENTE']) SR.getRange(2,map['CLIENTE']).setValue(cliente); // ✅ capitalizado
    if (map['DNI']) SR.getRange(2,map['DNI']).setValue(dni);
    if (map['MARCA']) SR.getRange(2,map['MARCA']).setValue(marca);
    if (map['MODELO']) SR.getRange(2,map['MODELO']).setValue(modelo);
    if (map['IMEI']) SR.getRange(2,map['IMEI']).setValue(imei);

    if (map['RECEPCIÓN']) SR.getRange(2,map['RECEPCIÓN']).setValue(recep);

    if (map['ESTADO ACTUAL']) {
      SR.getRange(2, map['ESTADO ACTUAL']).setFormula("='Equipos en Taller'!I:I");
    }

    if (map['REPUESTOS SOLICITADOS']) SR.getRange(2,map['REPUESTOS SOLICITADOS']).setValue(repSolicStr);
    if (map['OBSERVACIONES']) SR.getRange(2,map['OBSERVACIONES']).setValue(obsRecep);
    if (map['EMPRESA']) SR.getRange(2,map['EMPRESA']).setValue(empresa);
  }

  // RESERVAS (sin genéricos)
  {
    const counts = countRepuestos_(repuestosArr);
    if (counts.size){
      const SV = sh(SHEET_RESERVAS);
      const map = getHeaderMap_(SV);

      for (const [repuesto, cantidad] of counts.entries()){
        SV.insertRows(2,1);
        if (map['FECHA'])      SV.getRange(2,map['FECHA']).setValue(fechaHoy);
        if (map['N° REMITO'])  SV.getRange(2,map['N° REMITO']).setValue(nro);
        if (map['MOVIMIENTO']) SV.getRange(2,map['MOVIMIENTO']).setValue('RESERVA');
        if (map['MARCA'])      SV.getRange(2,map['MARCA']).setValue(marca);
        if (map['MODELO'])     SV.getRange(2,map['MODELO']).setValue(modelo);
        if (map['CATEGORÍA'])  SV.getRange(2,map['CATEGORÍA']).setValue(repuesto);
        if (map['CANTIDAD'])   SV.getRange(2,map['CANTIDAD']).setValue(cantidad);
      }
    }
  }

  limpiarRR_();
  SpreadsheetApp.getUi().alert(`RR #${nro} OK. PDF creado y bases actualizadas.`);
}

// ============== GENERAR PDF ENTREGA (RE) =================
function validarCamposObligatoriosRE_(SRE){
  const missing = [];
  const formaPagoCheck = SRE.getRange(RE_FORMA_PAGO_RANGE).getDisplayValue().toString().trim();
  const monedaCheck = readA1(SRE, RE_MONEDA, true);
  if (!formaPagoCheck) missing.push('Forma de Pago');
  if (!monedaCheck) missing.push('Moneda');
  return missing;
}

function generarPDFEntregaCore_(){
  const SRE = ss().getSheetByName(SHEET_RE);
  if (!SRE) throw new Error('No existe la hoja: ' + SHEET_RE);

  const missing = validarCamposObligatoriosRE_(SRE);
  if (missing.length) {
    const err = new Error('Campos obligatorios incompletos');
    err.code = 'VALIDATION';
    err.missing = missing;
    throw err;
  }

  const nro = Number(readA1(SRE, RE_NRO, true));
  if (!nro) throw new Error('Ingrese un N° REMITO en L2 (Hoja Remito de Entrega).');

  const ST = sh(SHEET_TALLER);
  const rowT = findRowByKey_(ST,'N° REMITO',nro);
  if (rowT===-1) throw new Error('N° REMITO no encontrado en "Equipos en Taller".');

  const mapT = getHeaderMap_(ST);
  
  const colTec = getCol_(mapT, ['TÉCNICO','TECNICO']);
  const tecnicoAsignado = colTec ? ST.getRange(rowT, colTec).getDisplayValue().toString().trim() : '';

  if (!tecnicoAsignado) {
    throw new Error(`No se puede generar el PDF de Entrega del remito #${nro} porque no hay un técnico asignado en "Equipos en Taller".`);
  }

  const colEstado = mapT['ESTADO ACTUAL'];
  const estado = colEstado ? ST.getRange(rowT,colEstado).getDisplayValue().toString().trim() : '';
  if (estado!=='Terminado' && estado!=='Sin Reparación') throw new Error('Solo puede generarse el PDF de Entrega cuando el ESTADO ACTUAL es "Terminado" o "Sin Reparación".');

  const cliente = readA1(SRE, RE_CLIENTE);
  const marca   = readA1(SRE, RE_MARCA);
  const modelo  = readA1(SRE, RE_MODELO);
  const fechaEntrega = (() => {
    const v = SRE.getRange(RE_FECHA).getValue();
    return v instanceof Date ? v : new Date();
  })();

  const moneda = readA1(SRE, RE_MONEDA, true);
  const costoTotal = getMergedNumber_(SRE, RE_COST_TOTAL);
  const totalGeneral = getMergedPairNumber_(SRE, 37, RE_PRICE_COLS);
  const precioUsd = getMergedNumber_(SRE, RE_PRECIO_USD_CELL);
  const formaPago = SRE.getRange(RE_FORMA_PAGO_RANGE).getDisplayValue().toString().trim();

  SpreadsheetApp.flush();
  Utilities.sleep(300);
  SpreadsheetApp.flush();

  const fileName = `RE_${nro}_${cliente||'Cliente'}.pdf`;
  const urlPDF = exportSheetPdf_(SHEET_RE, FOLDER_RE, fileName);

  // HISTORIAL: actualizar fila existente
  {
    const SH = sh(SHEET_HISTORIAL);
    const rowH = findRowByKey_(SH,'N° REMITO',nro);
    if (rowH===-1) throw new Error('No existe el N° REMITO en "Historial". (Debe haberse generado primero el RR).');

    const map = getHeaderMap_(SH);

    const colFechaEnt = getCol_(map, ['FECHA ENTREGA','FECHA DE ENTREGA']);
    if (colFechaEnt) SH.getRange(rowH, colFechaEnt).setValue(fechaEntrega);

    const colTotal = getCol_(map, ['TOTAL']);
    if (colTotal) SH.getRange(rowH, colTotal).setValue(totalGeneral);

    const colLinkRE = getCol_(map, ['LINK DE REMITO DE ENTREGA','LINK REMITO DE ENTREGA','LINK RE']);
    if (colLinkRE) SH.getRange(rowH, colLinkRE).setValue(urlPDF);

    const colMoneda = getCol_(map, ['MONEDA']);
    if (colMoneda) SH.getRange(rowH, colMoneda).setValue(moneda);

    const colCostoUsd = getCol_(map, ['COSTO USD']);
    if (colCostoUsd) SH.getRange(rowH, colCostoUsd).setValue(costoTotal);

    const colPrecioUsd = getCol_(map, ['PRECIO USD']);
    if (colPrecioUsd) SH.getRange(rowH, colPrecioUsd).setValue(precioUsd);

    const colProfit = getCol_(map, ['PROFIT']);
    if (colProfit) SH.getRange(rowH, colProfit).setValue((precioUsd || 0) - (costoTotal || 0));

    const colFormaPago = getCol_(map, ['FORMA DE PAGO']);
    if (colFormaPago) SH.getRange(rowH, colFormaPago).setValue(formaPago);
  }

  // INVENTARIO: se conserva igual (VENTA desde RE)
  {
    const SI   = sh(SHEET_INVENTARIO);
    const mapI = getHeaderMap_(SI);
    const agg = new Map();

    for (let r = RE_ROW_START; r <= RE_ROW_END; r++) {
      const importeLinea = getMergedPairNumber_(SRE, r, RE_PRICE_COLS);
      let cant = getMergedNumber_(SRE, `${RE_QTY_COL}${r}`);
      if (!cant || cant <= 0) cant = 1;

      const genericosEncontrados = new Set();
      for (const col of RE_CAT_COLS) {
        const catRaw = SRE.getRange(`${col}${r}`).getDisplayValue().toString().trim();
        if (!catRaw) continue;
        genericosEncontrados.add(catRaw);
      }
      if (genericosEncontrados.size === 0) continue;

      const importeUnitario = cant > 0 ? (importeLinea / cant) : importeLinea;

      for (const gen of genericosEncontrados) {
        const cur = agg.get(gen) || { cant: 0, total: 0 };
        cur.cant  += cant;
        cur.total += (importeUnitario * cant);
        agg.set(gen, cur);
      }
    }

    for (const [gen, info] of agg.entries()) {
      SI.insertRows(INV_ROW_INSERT, 1);

      if (mapI['FECHA'])      SI.getRange(INV_ROW_INSERT, mapI['FECHA']).setValue(fechaEntrega);
      if (mapI['MOVIMIENTO']) SI.getRange(INV_ROW_INSERT, mapI['MOVIMIENTO']).setValue('VENTA');
      if (mapI['MARCA'])      SI.getRange(INV_ROW_INSERT, mapI['MARCA']).setValue(marca);
      if (mapI['MODELO'])     SI.getRange(INV_ROW_INSERT, mapI['MODELO']).setValue(modelo);
      if (mapI['CATEGORÍA'])  SI.getRange(INV_ROW_INSERT, mapI['CATEGORÍA']).setValue(gen);
      if (mapI['CANTIDAD'])   SI.getRange(INV_ROW_INSERT, mapI['CANTIDAD']).setValue(info.cant);

      const precioUnit = info.cant > 0 ? (info.total / info.cant) : info.total;
      if (mapI['PRECIO'])     SI.getRange(INV_ROW_INSERT, mapI['PRECIO']).setValue(precioUnit);
      if (mapI['TOTAL'])      SI.getRange(INV_ROW_INSERT, mapI['TOTAL']).setValue(info.total);
    }
  }

  // RECEPCIÓN: actualizar estado + fecha
  {
    const SR = sh(SHEET_RECEPCION);
    const rowR = findRowByKey_(SR,'N° REMITO',nro);
    if (rowR!==-1){
      const map = getHeaderMap_(SR);
      if (map['ESTADO ACTUAL']) SR.getRange(rowR, map['ESTADO ACTUAL']).setValue('Entregado');
      const colFechaEnt = getCol_(map, ['FECHA DE ENTREGA','FECHA ENTREGA']);
      if (colFechaEnt) SR.getRange(rowR, colFechaEnt).setValue(fechaEntrega);
    }
  }

  // EQUIPOS EN TALLER: set ENTREGA = fechaEntrega
  {
    const ST2 = sh(SHEET_TALLER);
    const rowPivot = findRowByKey_(ST2, 'N° REMITO', nro);
    if (rowPivot !== -1) {
      const mapPivot = getHeaderMap_(ST2);
      if (mapPivot['ENTREGA']) {
        ST2.getRange(rowPivot, mapPivot['ENTREGA']).setValue(fechaEntrega);
      }
    }
  }

  // RESERVAS: borrar todas las filas con ese N° REMITO
  {
    const SV = sh(SHEET_RESERVAS);
    const vals = SV.getDataRange().getValues();
    if (vals.length>1){
      const idx = vals[0].indexOf('N° REMITO');
      if (idx!==-1){
        for (let i=vals.length-1;i>=1;i--){
          if (vals[i][idx]==nro) SV.deleteRow(i+1);
        }
      }
    }
  }

  limpiarRE_();
  return {
    ok: true,
    nro: nro,
    urlPDF: urlPDF,
    message: `RE #${nro} OK. PDF creado y bases actualizadas.`
  };
}

function generarPDFEntrega(){
  const res = generarPDFEntregaCore_();
  try {
    SpreadsheetApp.getUi().alert(res.message || `RE #${res.nro} OK. PDF creado y bases actualizadas.`);
  } catch (_) {}
  return true;
}

/**
 * Limpia la fila 2 de CARGA pero preserva la fórmula de I2.
 * (También preserva por seguridad el formato, pero acá solo cuidamos fórmula.)
 */
function limpiarCargaInventarioManteniendoFormulas_(SI){
  const rngA2 = SI.getRange('A2');
  const rngH2 = SI.getRange('H2');
  const rngI2 = SI.getRange('I2');

  const fA2 = rngA2.getFormula();
  const fH2 = rngH2.getFormula();
  const fI2 = rngI2.getFormula();

  // Limpia toda la fila 2 (A2:J2)
  SI.getRange('A2:J2').clearContent();

  // Restaura fórmulas si existían
  if (fA2) rngA2.setFormula(fA2);
  if (fH2) rngH2.setFormula(fH2);
  if (fI2) rngI2.setFormula(fI2);
}

/**
 * CARGA DE INVENTARIO:
 * - Inserta una fila en MOVIMIENTOS INVENTARIO (la misma hoja "Inventario") en INV_ROW_INSERT.
 * - Soporta movimiento SCRAP:
 *    * PRECIO y TOTAL VACÍOS
 *    * COSTO y TÉCNICO se registran
 * - Al final: limpia fila 2 excepto fórmula I2.
 */
function cargaDeInventario(){
  ensureActiveSheet_(SHEET_INVENTARIO);
  const SI = sh(SHEET_INVENTARIO);

  const fecha   = SI.getRange(INV_CARGA.FECHA).getValue() || new Date();
  const movRaw  = SI.getRange(INV_CARGA.MOV).getDisplayValue().toString().trim();
  const mov     = movRaw.toUpperCase();

  const marca   = SI.getRange(INV_CARGA.MARCA).getDisplayValue().toString().trim();
  const modelo  = SI.getRange(INV_CARGA.MODELO).getDisplayValue().toString().trim();
  const categ   = SI.getRange(INV_CARGA.CATEG).getDisplayValue().toString().trim();

  const cant    = Number(SI.getRange(INV_CARGA.CANT).getValue()) || 0;

  const precio  = Number(SI.getRange(INV_CARGA.PREC).getValue()) || 0;
  const total   = cant * precio;

  // COSTO: viene desde I2 (fórmula). Guardamos el valor mostrado, no la fórmula.
  const costo   = parseNumberLocal_(SI.getRange(INV_CARGA.COSTO).getDisplayValue());
  const tecnico = SI.getRange(INV_CARGA.TECNICO).getDisplayValue().toString().trim();

  if (!mov) throw new Error('Falta MOVIMIENTO en B2.');
  if (!marca || !modelo || !categ) throw new Error('Faltan datos obligatorios (MARCA / MODELO / CATEGORÍA).');
  if (!cant || cant <= 0) throw new Error('CANTIDAD debe ser mayor a 0.');

  // ✅ BLOQUEO: SCRAP requiere TÉCNICO
  if (mov === 'SCRAP' && !tecnico) {
    SpreadsheetApp.getUi().alert(
      'Acción bloqueada',
      'No se puede cargar un movimiento de SCRAP si no hay un TÉCNICO seleccionado.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Insertar fila en "MOVIMIENTOS INVENTARIO" (misma hoja)
  SI.insertRows(INV_ROW_INSERT,1);

  const map = getHeaderMap_(SI);

  // Campos comunes
  if (map['FECHA'])      SI.getRange(INV_ROW_INSERT, map['FECHA']).setValue(fecha);
  if (map['MOVIMIENTO']) SI.getRange(INV_ROW_INSERT, map['MOVIMIENTO']).setValue(mov);
  if (map['MARCA'])      SI.getRange(INV_ROW_INSERT, map['MARCA']).setValue(marca);
  if (map['MODELO'])     SI.getRange(INV_ROW_INSERT, map['MODELO']).setValue(modelo);
  if (map['CATEGORÍA'])  SI.getRange(INV_ROW_INSERT, map['CATEGORÍA']).setValue(categ);
  if (map['CANTIDAD'])   SI.getRange(INV_ROW_INSERT, map['CANTIDAD']).setValue(cant);

  // Nuevas columnas (si existen)
  if (map['COSTO'])      SI.getRange(INV_ROW_INSERT, map['COSTO']).setValue(costo || '');
  if (map['TÉCNICO'])    SI.getRange(INV_ROW_INSERT, map['TÉCNICO']).setValue(tecnico || '');
  if (map['TECNICO'])    SI.getRange(INV_ROW_INSERT, map['TECNICO']).setValue(tecnico || ''); // por si está sin tilde

  // Comportamiento por movimiento
  if (mov === 'SCRAP') {
    // PRECIO y TOTAL vacíos
    if (map['PRECIO']) SI.getRange(INV_ROW_INSERT, map['PRECIO']).setValue('');
    if (map['TOTAL'])  SI.getRange(INV_ROW_INSERT, map['TOTAL']).setValue('');
  } else {
    // Comportamiento normal (COMPRA/VENTA/etc)
    if (map['PRECIO']) SI.getRange(INV_ROW_INSERT, map['PRECIO']).setValue(precio);
    if (map['TOTAL'])  SI.getRange(INV_ROW_INSERT, map['TOTAL']).setValue(total);
  }

  // ✅ Limpieza fila 2 preservando fórmulas de A2, H2 e I2
  limpiarCargaInventarioManteniendoFormulas_(SI);

  SpreadsheetApp.getUi().alert('Movimiento de Inventario registrado exitosamente.');
}

/** ==== Dólar Blue en ROSARIO (VENTA) -> K53 y K42 ==== */
function getDolarBlueRosarioVenta_() {
  const url = 'https://www.infodolar.com/cotizacion-dolar-localidad-rosario-provincia-santa-fe.aspx';
  const html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();

  const regex = /<td class="colCompraVenta"[^>]*>\s*\$ ?([\d.,]+)/g;
  const matches = [...html.matchAll(regex)];

  if (matches.length < 4) {
    throw new Error('No se encontró el cuarto valor (venta) en la página de InfoDolar.');
  }

  const valorStr = matches[3][1];
  const valorNum = parseFloat(valorStr.replace(/\./g, '').replace(',', '.'));
  if (isNaN(valorNum)) {
    throw new Error('No se pudo parsear el valor del dólar blue en Rosario.');
  }
  return valorNum;
}

function obtenerDolarBlueRosario() {
  const valorNum = getDolarBlueRosarioVenta_();
  const libro = ss();
  const updatedAtIso = new Date().toISOString();

  const hojaRR = libro.getSheetByName(SHEET_RR);
  if (hojaRR) {
    const cellRR = hojaRR.getRange(RE_DOLAR_A1);
    cellRR.setValue(valorNum);
    cellRR.setNumberFormat('#,##0.00');
  }

  const hojaRE = libro.getSheetByName(SHEET_RE);
  if (hojaRE) {
    const cellRE = hojaRE.getRange(RE_DOLAR_A2);
    cellRE.setValue(valorNum);
    cellRE.setNumberFormat('#,##0.00');
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty(DOLAR_ROSARIO_UPDATED_AT_KEY, updatedAtIso);
  props.setProperty(DOLAR_ROSARIO_VALUE_KEY, String(valorNum));
  return { ok:true, value:valorNum, updatedAt:updatedAtIso };
}

function activarDolarCadaMinuto() {
  const handler = 'obtenerDolarBlueRosario';
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(1)
    .create();
}

function desactivarDolarCadaMinuto() {
  const handler = 'obtenerDolarBlueRosario';
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/** ===== APIs para la interfaz (Index.html) ===== **/
function apiGenerarRR(){ generarPDFRecepcion(); return true; }
function apiGenerarRE(){ return generarPDFEntregaCore_(); }

function apiREGuardarPagoMonedaIVA(nro, formaPago, moneda, ivaBool){
  try {
    const nroNum = Number(nro);
    if (!Number.isFinite(nroNum) || nroNum <= 0) {
      return { ok:false, message:'N° remito inválido.' };
    }

    const spreadsheet = SpreadsheetApp.openById(ss().getId());
    const SRE = spreadsheet.getSheetByName(SHEET_RE);
    if (!SRE) return { ok:false, message:'No existe la hoja: ' + SHEET_RE };

    writeA1(SRE, RE_NRO, nroNum);
    SRE.getRange(RE_FORMA_PAGO_RANGE).setValue(String(formaPago || '').trim());
    writeA1(SRE, RE_MONEDA, String(moneda || '').trim());
    SRE.getRange('J16').setValue(!!ivaBool);
    SpreadsheetApp.flush();

    return { ok:true, message:'Datos de pago persistidos.' };
  } catch (e) {
    return { ok:false, message:String((e && e.message) ? e.message : e) };
  }
}

function apiGenerarREWeb(nro){
  try {
    const nroNum = Number(nro);
    if (!Number.isFinite(nroNum) || nroNum <= 0) {
      return {
        ok:false,
        code:'VALIDATION',
        missing:['N° Remito'],
        message:'Ingresá un N° remito válido.'
      };
    }

    const SRE = sh(SHEET_RE);
    writeA1(SRE, RE_NRO, nroNum);
    SpreadsheetApp.flush();
    Utilities.sleep(300);
    SpreadsheetApp.flush();

    const missing = validarCamposObligatoriosRE_(SRE);
    if (missing.length) {
      return {
        ok:false,
        code:'VALIDATION',
        missing,
        message:'Campos obligatorios incompletos'
      };
    }

    const res = generarPDFEntregaCore_();
    return {
      ok: true,
      code:'OK',
      nro: res.nro,
      urlPDF: res.urlPDF,
      message: res.message || 'RE generado correctamente.'
    };
  } catch (e) {
    const msg = String((e && e.message) ? e.message : e);
    const missing = (e && Array.isArray(e.missing)) ? e.missing : [];
    const code = (e && e.code) ? String(e.code) : 'ERROR';
    return {
      ok: false,
      code,
      missing,
      message: msg
    };
  }
}
function apiUsdRosarioGet(){
  try {
    const props = PropertiesService.getScriptProperties();
    let value = parseDolarValue_(props.getProperty(DOLAR_ROSARIO_VALUE_KEY));
    let updatedAt = String(props.getProperty(DOLAR_ROSARIO_UPDATED_AT_KEY) || '').trim();

    if (value === null) {
      value = getDolarRosarioFromSheets_();
      if (value !== null) props.setProperty(DOLAR_ROSARIO_VALUE_KEY, String(value));
    }

    if (!updatedAt) {
      updatedAt = getDolarUpdatedAtFallback_();
    }

    return { ok:true, value:value, updatedAt: updatedAt || null };
  } catch (e) {
    return {
      ok:false,
      message:String((e && e.message) ? e.message : e)
    };
  }
}

function apiUsdRosarioRefresh(){
  try {
    const res = obtenerDolarBlueRosario();
    return {
      ok:true,
      value: Number(res && res.value),
      updatedAt: (res && res.updatedAt) ? String(res.updatedAt) : null
    };
  } catch (e) {
    return {
      ok:false,
      message:String((e && e.message) ? e.message : e)
    };
  }
}

function apiActualizarDolar(){
  const res = apiUsdRosarioRefresh();
  if (!res.ok) throw new Error(res.message || 'No se pudo actualizar USD Rosario.');
  return res;
}

function apiGetDolarRosarioStatus(){
  return apiUsdRosarioGet();
}

function getDolarRosarioFromSheets_(){
  const libro = ss();
  const hojaRE = libro.getSheetByName(SHEET_RE);
  const hojaRR = libro.getSheetByName(SHEET_RR);

  let value = null;
  if (hojaRE) {
    value = parseDolarValue_(hojaRE.getRange(RE_DOLAR_A2).getValue());
  }
  if (value === null && hojaRR) {
    value = parseDolarValue_(hojaRR.getRange(RE_DOLAR_A1).getValue());
  }
  return value;
}

function getDolarUpdatedAtFallback_(){
  const libro = ss();
  const hojaRE = libro.getSheetByName(SHEET_RE);
  const hojaRR = libro.getSheetByName(SHEET_RR);
  const fromRE = hojaRE ? String(hojaRE.getRange(RE_FECHA_HORA_VIS).getDisplayValue() || '').trim() : '';
  if (fromRE) return fromRE;
  const fromRR = hojaRR ? String(hojaRR.getRange(RR_FECHA_HORA_VIS).getDisplayValue() || '').trim() : '';
  if (fromRR) return fromRR;
  return '';
}

function parseDolarValue_(raw){
  if (raw === null || raw === undefined || raw === '') return null;
  if (typeof raw === 'number' && isFinite(raw)) return raw;
  const txt = String(raw).trim();
  if (!txt) return null;
  const parsed = Number(txt.replace(/\./g, '').replace(',', '.').replace(/[^0-9.-]/g, ''));
  return isFinite(parsed) ? parsed : null;
}

function apiListarHistorial(q, limit){
  const S = sh(SHEET_HISTORIAL);
  const lastRow = S.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = Math.max(1, S.getLastColumn());
  const headers = S.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
  const map = _headerMapFromArray_(headers);


  const IDX = {
    nro: getColFlexible_(map, ['N° REMITO', 'NRO REMITO', 'NRO', 'REMITO']) - 1,
    ingreso: getColFlexible_(map, ['FECHA INGRESO', 'INGRESO', 'FECHA DE INGRESO']) - 1,
    empresa: getColFlexible_(map, ['EMPRESA']) - 1,
    cliente: getColFlexible_(map, ['CLIENTE']) - 1,

  };

  const numRows = lastRow - 1;
  const rng = S.getRange(2, 1, numRows, lastCol);
  const values = rng.getValues();
  const displays = rng.getDisplayValues();
  const rich = rng.getRichTextValues();

  const tz = Session.getScriptTimeZone();
  const toDateStr = (raw, display) => {
    if (raw instanceof Date && !isNaN(raw.getTime())) {
      return Utilities.formatDate(raw, tz, 'dd/MM/yyyy');
    }
    const txt = String(display || raw || '').trim();
    if (!txt) return '';
    const m = txt.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
      const dd = String(m[1]).padStart(2, '0');
      const mm = String(m[2]).padStart(2, '0');
      const yy = m[3].length === 2 ? `20${m[3]}` : m[3];
      return `${dd}/${mm}/${yy}`;
    }
    const d = new Date(txt);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'dd/MM/yyyy');
    return txt;
  };

  const toLink = (richCell, rawCell, displayCell) => {
    try {
      if (richCell && typeof richCell.getLinkUrl === 'function') {
        const direct = richCell.getLinkUrl();
        if (direct) return String(direct).trim();
      }
      if (richCell && typeof richCell.getRuns === 'function') {
        const runs = richCell.getRuns() || [];
        for (let i = 0; i < runs.length; i++) {
          const runLink = runs[i].getLinkUrl();
          if (runLink) return String(runLink).trim();
        }
      }
    } catch (_) {}
    const fallback = String(rawCell || displayCell || '').trim();
    if (/^https?:\/\//i.test(fallback)) return fallback;
    return '';
  };

  const pick = (arr, idx) => (idx >= 0 && idx < arr.length) ? arr[idx] : '';

  const toNorm = (v) => String(v || '').toLowerCase();
  const query = toNorm(String(q || '').trim());
  const max = Math.max(1, Number(limit || 50));

  if (displays.length > 0) {
    try {
      const rowD0 = displays[0] || [];
      const rowV0 = values[0] || [];
      const rowR0 = rich[0] || [];
      const dbgLinkRR = toLink(pick(rowR0, IDX.linkRR), pick(rowV0, IDX.linkRR), pick(rowD0, IDX.linkRR));
      const dbgLinkRE = toLink(pick(rowR0, IDX.linkRE), pick(rowV0, IDX.linkRE), pick(rowD0, IDX.linkRE));
      Logger.log('Historial debug map idx: ' + JSON.stringify(IDX));
      Logger.log('Historial debug displayValues[0..' + (lastCol - 1) + ']: ' + JSON.stringify(rowD0.slice(0, lastCol)));
      Logger.log('Historial debug links: ' + JSON.stringify({ linkRR: dbgLinkRR, linkRE: dbgLinkRE }));
    } catch (_) {}
  }

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const rowV = values[i];
    const rowD = displays[i];
    const rowR = rich[i];

    const nro = pick(rowD, IDX.nro);
    const fechaIngreso = toDateStr(pick(rowV, IDX.ingreso), pick(rowD, IDX.ingreso));
    const empresa = String(pick(rowD, IDX.empresa) || '').trim();
    const cliente = String(pick(rowD, IDX.cliente) || '').trim();
    const modelo = String(pick(rowD, IDX.modelo) || '').trim();
    const linkRR = toLink(pick(rowR, IDX.linkRR), pick(rowV, IDX.linkRR), pick(rowD, IDX.linkRR));
    const fechaEntrega = toDateStr(pick(rowV, IDX.fechaEntrega), pick(rowD, IDX.fechaEntrega));
    const linkRE = toLink(pick(rowR, IDX.linkRE), pick(rowV, IDX.linkRE), pick(rowD, IDX.linkRE));
    const moneda = String(pick(rowD, IDX.moneda) || '').trim();

    const rawTotal = pick(rowV, IDX.total);
    const dispTotal = pick(rowD, IDX.total);
    const total = String(dispTotal || rawTotal || '').trim();

    const item = {
      nro: String(nro || '').trim(),
      ingreso: fechaIngreso,
      empresa: empresa,
      cliente: cliente,
      modelo: modelo,
      linkRR: linkRR,
      entrega: fechaEntrega,
      linkRE: linkRE,
      moneda: moneda,
      total: total
    };

    const searchable = [item.nro, item.cliente, item.modelo, item.ingreso, item.entrega]
      .map(toNorm)
      .join(' | ');

    if (!query || searchable.includes(query)) {
      out.push(item);
      if (out.length >= max) break;
    }
  }

  if (out.length) {
    try {
      Logger.log(JSON.stringify(out[0]));
    } catch (_) {}
  }

  // Se respeta el orden de la hoja (fila 2 hacia abajo).
  return out;
}

function getDashboardData(){
  const started = Date.now();

  const SH = sh(SHEET_HISTORIAL);
  const ST = sh(SHEET_TALLER);

  const histValues = SH.getDataRange().getDisplayValues(); // 1 lectura hoja Historial
  const tallValues = ST.getDataRange().getDisplayValues(); // 1 lectura hoja Taller

  const histRows = histValues.length > 1 ? histValues.slice(1) : [];
  const tallRows = tallValues.length > 1 ? tallValues.slice(1) : [];

  const headH = histValues.length ? histValues[0] : [];
  const headT = tallValues.length ? tallValues[0] : [];

  const colH = {
    nro: getColFlexible_(_headerMapFromArray_(headH), ['N° REMITO','NRO REMITO','NRO']),
    total: getColFlexible_(_headerMapFromArray_(headH), ['TOTAL']),
    costoUsd: getColFlexible_(_headerMapFromArray_(headH), ['COSTO USD','COSTO TOTAL USD']),
    precioUsd: getColFlexible_(_headerMapFromArray_(headH), ['PRECIO USD','PRECIO TOTAL USD']),
    profit: getColFlexible_(_headerMapFromArray_(headH), ['PROFIT','GANANCIA']),
    formaPago: getColFlexible_(_headerMapFromArray_(headH), ['FORMA DE PAGO','FORMA PAGO']),
    moneda: getColFlexible_(_headerMapFromArray_(headH), ['MONEDA']),
    fechaEntrega: getColFlexible_(_headerMapFromArray_(headH), ['FECHA ENTREGA','FECHA DE ENTREGA'])
  };

  const colT = {
    estado: getColFlexible_(_headerMapFromArray_(headT), ['ESTADO ACTUAL','ESTADO']),
    tecnico: getColFlexible_(_headerMapFromArray_(headT), ['TÉCNICO','TECNICO']),
    modelo: getColFlexible_(_headerMapFromArray_(headT), ['MODELO']),
    repuestos: getColFlexible_(_headerMapFromArray_(headT), ['REPUESTOS UTILIZADOS','REPUESTOS']),
    entrega: getColFlexible_(_headerMapFromArray_(headT), ['ENTREGA'])
  };

  let equiposIngresados = 0;
  let equiposEntregados = 0;
  let costoTotal = 0;
  let precioTotal = 0;
  let profitTotal = 0;
  let pesosFacturados = 0;
  let usdFacturados = 0;
  let costoRepuestosUsd = 0;

  const ingresosPorFormaPago = new Map();
  const distribucionMoneda = new Map();

  for (let i = 0; i < histRows.length; i++) {
    const row = histRows[i];
    const nro = _readCell_(row, colH.nro);
    if (String(nro || '').trim()) equiposIngresados++;

    const entrega = _readCell_(row, colH.fechaEntrega);
    if (String(entrega || '').trim()) equiposEntregados++;

    const total = parseNumberLocal_(_readCell_(row, colH.total));
    const costoUsd = parseNumberLocal_(_readCell_(row, colH.costoUsd));
    const precioUsd = parseNumberLocal_(_readCell_(row, colH.precioUsd));
    const profit = parseNumberLocal_(_readCell_(row, colH.profit));

    costoTotal += costoUsd;
    precioTotal += precioUsd;
    profitTotal += profit || (precioUsd - costoUsd);

    const formaPago = String(_readCell_(row, colH.formaPago) || '').trim() || 'Sin dato';
    ingresosPorFormaPago.set(formaPago, (ingresosPorFormaPago.get(formaPago) || 0) + (total || 0));

    const moneda = String(_readCell_(row, colH.moneda) || '').trim() || 'Sin dato';
    distribucionMoneda.set(moneda, (distribucionMoneda.get(moneda) || 0) + 1);

    const monedaNorm = normalizeStr_(moneda);
    if (monedaNorm.includes('usd') || monedaNorm.includes('dolar') || monedaNorm.includes('dólar')) {
      usdFacturados += (total || precioUsd || 0);
    } else {
      pesosFacturados += (total || 0);
    }
    costoRepuestosUsd += costoUsd;
  }

  const estadosCount = new Map([
    ['Pendiente Diagnóstico', 0],
    ['En Reparación', 0],
    ['Terminado', 0],
    ['Sin Reparación', 0],
    ['Entregado', 0]
  ]);

  const reparadosPorTecnico = new Map();
  const repuestosPorModelo = new Map();

  for (let i = 0; i < tallRows.length; i++) {
    const row = tallRows[i];
    const estadoRaw = String(_readCell_(row, colT.estado) || '').trim();
    const estadoNorm = normalizeStr_(estadoRaw);
    const estadoLabel = _normalizarEstadoDashboard_(estadoNorm, estadoRaw);
    if (estadosCount.has(estadoLabel)) {
      estadosCount.set(estadoLabel, (estadosCount.get(estadoLabel) || 0) + 1);
    }

    const tecnico = String(_readCell_(row, colT.tecnico) || '').trim() || 'Sin técnico';
    if (estadoLabel === 'Terminado' || estadoLabel === 'Sin Reparación' || estadoLabel === 'Entregado') {
      reparadosPorTecnico.set(tecnico, (reparadosPorTecnico.get(tecnico) || 0) + 1);
    }

    const modelo = String(_readCell_(row, colT.modelo) || '').trim() || 'Sin modelo';
    const repuestosRaw = String(_readCell_(row, colT.repuestos) || '').trim();
    if (!repuestosRaw) continue;

    const repuestos = repuestosRaw
      .split(/[,;\n]/)
      .map(x => String(x || '').trim())
      .filter(Boolean);

    if (!repuestos.length) continue;
    if (!repuestosPorModelo.has(modelo)) repuestosPorModelo.set(modelo, new Map());
    const byRep = repuestosPorModelo.get(modelo);
    repuestos.forEach((rep) => {
      byRep.set(rep, (byRep.get(rep) || 0) + 1);
    });
  }

  const pagosOrdenados = Array.from(ingresosPorFormaPago.entries())
    .sort((a,b)=> b[1]-a[1]);
  const monedasOrdenadas = Array.from(distribucionMoneda.entries())
    .sort((a,b)=> b[1]-a[1]);
  const tecnicosOrdenados = Array.from(reparadosPorTecnico.entries())
    .sort((a,b)=> b[1]-a[1]);

  const topRepuestos = _topRepuestosDesdeModelos_(repuestosPorModelo, 6);
  const modelosOrdenados = Array.from(repuestosPorModelo.keys())
    .sort((a,b)=> a.localeCompare(b, 'es'))
    .slice(0, 24);
  const repuestosSeries = topRepuestos.map((rep)=>({
    label: rep,
    data: modelosOrdenados.map((modelo)=>{
      const byRep = repuestosPorModelo.get(modelo);
      return byRep ? (byRep.get(rep) || 0) : 0;
    })
  }));

  return {
    ok: true,
    generatedAt: new Date().toISOString(),
    elapsedMs: Date.now() - started,
    kpis: {
      equiposIngresados,
      equiposEntregados,
      costoTotal,
      precioTotal,
      profitTotal
    },
    tecnicoReparaciones: {
      labels: tecnicosOrdenados.map(x=>x[0]),
      data: tecnicosOrdenados.map(x=>x[1])
    },
    ingresosFormaPago: {
      labels: pagosOrdenados.map(x=>x[0]),
      data: pagosOrdenados.map(x=>Number(x[1] || 0))
    },
    facturacion: {
      pesosFacturados,
      usdFacturados,
      costoRepuestosUsd
    },
    repuestosPorModelo: {
      labels: modelosOrdenados,
      datasets: repuestosSeries
    },
    distribucionMoneda: {
      labels: monedasOrdenadas.map(x=>x[0]),
      data: monedasOrdenadas.map(x=>x[1])
    },
    estadosTaller: {
      labels: Array.from(estadosCount.keys()),
      data: Array.from(estadosCount.values())
    }
  };
}

function _headerMapFromArray_(head){
  const map = {};
  (head || []).forEach((h, i)=>{
    map[String(h || '').trim()] = i + 1;
  });
  return map;
}

function _readCell_(row, col1Based){
  if (!row || !Number.isFinite(col1Based) || col1Based <= 0) return '';
  return row[col1Based - 1] || '';
}

function _normalizarEstadoDashboard_(estadoNorm, raw){
  if (estadoNorm.includes('pendiente')) return 'Pendiente Diagnóstico';
  if (estadoNorm.includes('reparacion') || estadoNorm.includes('reparación')) return 'En Reparación';
  if (estadoNorm.includes('terminado')) return 'Terminado';
  if (estadoNorm.includes('sin reparacion') || estadoNorm.includes('sin reparación')) return 'Sin Reparación';
  if (estadoNorm.includes('entregado')) return 'Entregado';
  return String(raw || '').trim() || 'Pendiente Diagnóstico';
}

function _topRepuestosDesdeModelos_(modelosMap, topN){
  const totalByRep = new Map();
  modelosMap.forEach((repMap)=>{
    repMap.forEach((count, rep)=>{
      totalByRep.set(rep, (totalByRep.get(rep) || 0) + Number(count || 0));
    });
  });
  return Array.from(totalByRep.entries())
    .sort((a,b)=> b[1]-a[1])
    .slice(0, Number(topN || 6))
    .map(x=>x[0]);
}

/** ===== helpers para UI: obtener listas únicas desde hojas ===== **/
function uniqueNonEmpty_(arr){
  const set = new Set();
  (arr||[]).forEach(v=>{ const s = String(v||'').toString().trim(); if (s) set.add(s); });
  return Array.from(set);
}

/**
 * Intenta extraer las opciones de un desplegable a partir de la validación de datos
 * aplicada a la celda A1 indicada en la hoja. Retorna array vacío si no hay
 * validación o no se pudo resolver.
 */
function getValidationOptionsFromA1_(sheetName, a1){
  try{
    const S = sh(sheetName);
    const rng = S.getRange(a1);
    const dv = rng.getDataValidation();
    if (!dv) return [];
    const crit = dv.getCriteriaType();
    const vals = dv.getCriteriaValues() || [];

    // VALUE_IN_RANGE -> criterio[0] es un Range
    if (crit === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && vals[0]){
      const r = vals[0];
      if (r && typeof r.getValues === 'function'){
        const flattened = r.getDisplayValues().flat();
        return uniqueNonEmpty_(flattened);
      }
    }

    // VALUE_IN_LIST -> criterio[0] puede ser un array o una string
    if (crit === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST && vals[0]){
      const v0 = vals[0];
      if (Array.isArray(v0)) return uniqueNonEmpty_(v0.map(x=>String(x||'').trim()));
      const s = String(v0||'').trim();
      if (!s) return [];
      // Si viene en forma "a,b,c" (posible), separo por coma
      return uniqueNonEmpty_(s.split(',').map(x=>x.trim()).filter(Boolean));
    }

    // Otros tipos: intentar si criterio contiene Range en cualquier posición
    for (let i=0;i<vals.length;i++){
      const v = vals[i];
      if (v && typeof v.getValues === 'function'){
        return uniqueNonEmpty_(v.getDisplayValues().flat());
      }
      if (Array.isArray(v)) return uniqueNonEmpty_(v.map(x=>String(x||'').trim()));
      if (typeof v === 'string' && v.includes(',')) return uniqueNonEmpty_(v.split(',').map(x=>x.trim()).filter(Boolean));
    }

    return [];
  }catch(e){
    return [];
  }
}

function colValuesByHeader_(sheetName, headerNames){
  try{
    const S = sh(sheetName);
    const map = getHeaderMap_(S);
    const col = getCol_(map, headerNames);
    if (!col) return [];
    const vals = S.getRange(2, col, Math.max(1,S.getLastRow()-1), 1).getDisplayValues().map(r=>r[0]);
    return uniqueNonEmpty_(vals);
  }catch(e){ return []; }
}

/** Devuelve lista de marcas (intenta 'Datos' col B, luego 'Recepción'/'Inventario'). */
function apiGetMarcas(){
  // Intenta columna B de 'Datos' (común en la plantilla)
  try{
    // Primero: si la hoja RR tiene validación en RR_MARCA, usarla
    const fromRR = getValidationOptionsFromA1_(SHEET_RR, RR_MARCA);
    if (fromRR && fromRR.length) return fromRR;

    const SD = sh(SHEET_DATOS);
    const vals = SD.getRange('B2:B' + Math.max(2, SD.getLastRow())).getDisplayValues().flat();
    const uniq = uniqueNonEmpty_(vals);
    if (uniq.length) return uniq;
  }catch(e){}

  // Fallback: buscar header 'MARCA' en Recepción/Inventario
  let out = colValuesByHeader_(SHEET_RECEPCION, ['MARCA','Marca']);
  if (out.length) return out;
  out = colValuesByHeader_(SHEET_INVENTARIO, ['MARCA','Marca']);
  return out;
}

/** Devuelve modelos para una marca específica.
 * Busca en la hoja 'Datos', columna B (Marca) y devuelve valores únicos de columna D (Modelo).
 */
function apiGetModelos(marca){
  marca = String(marca||'').trim();
  if (!marca) return [];

  try {
    const SD = sh(SHEET_DATOS);
    const map = getHeaderMap_(SD);
    
    // Buscar columnas MARCA (generalmente B) y MODELO (generalmente D)
    const colMarca = getCol_(map, ['MARCA','Marca','B']);
    const colModelo = getCol_(map, ['MODELO','Modelo','D']);
    
    if (!colMarca || !colModelo) return [];
    
    const rows = SD.getRange(2, 1, Math.max(1, SD.getLastRow() - 1), Math.max(colMarca, colModelo)).getDisplayValues();
    const modelos = [];
    
    rows.forEach(r => {
      const marcaRow = String(r[colMarca - 1] || '').trim();
      if (marcaRow === marca) {
        const modeloRow = String(r[colModelo - 1] || '').trim();
        if (modeloRow) modelos.push(modeloRow);
      }
    });
    
    return uniqueNonEmpty_(modelos);
  } catch(e){}

  // Fallback: intentar validación en RR
  try{
    const fromRR = getValidationOptionsFromA1_(SHEET_RR, RR_MODELO);
    if (fromRR && fromRR.length) return fromRR;
  }catch(e){}

  // Último fallback: buscar en Recepción/Inventario
  let out = colValuesByHeader_(SHEET_RECEPCION, ['MODELO','Modelo']);
  if (out.length) return out;
  out = colValuesByHeader_(SHEET_INVENTARIO, ['MODELO','Modelo']);
  return out;
}

/** Devuelve lista de TODOS los modelos desde 'Datos RR' columna D (sin filtrar por marca). */
function apiGetModelosDatos(){
  try {
    // Intenta primero con la hoja 'Datos RR'
    let s = null;
    try {
      s = ss().getSheetByName('Datos RR');
    } catch(e){}
    
    if (s && s.getLastRow() >= 2){
      const vals = s.getRange('D2:D' + s.getLastRow()).getDisplayValues().flat();
      const uniq = uniqueNonEmpty_(vals);
      if (uniq.length) return uniq;
    }
  } catch(e){}

  // Fallback: intentar obtener desde 'Inventario' columna D (MODELO)
  try {
    const inv = sh(SHEET_INVENTARIO);
    if (inv && inv.getLastRow() >= 2){
      const vals = inv.getRange('D2:D' + inv.getLastRow()).getDisplayValues().flat();
      const uniq = uniqueNonEmpty_(vals);
      if (uniq.length) return uniq;
    }
  } catch(e){}

  // Fallback: obtener desde 'Recepción' buscando columna MODELO
  try {
    const modelos = colValuesByHeader_(SHEET_RECEPCION, ['MODELO','Modelo']);
    if (modelos.length) return modelos;
  } catch(e){}

  // Última opción: devolver array vacío
  return [];
}

/** Escribe la Marca seleccionada en D14 de 'Remito de Recepción'
    Esto activa la fórmula FILTER en 'Datos RR' que filtra automáticamente los datos */
function apiSetMarcaInD14(marca){
  try {
    const SRR = sh(SHEET_RR);
    writeA1(SRR, RR_MARCA, String(marca || '').trim());
    return true;
  } catch(e){
    console.error('Error al setear marca en D14:', e);
    return false;
  }
}

/** Devuelve lista de capacidades desde validación en RR_CAP (igual que Colors) */
function apiGetCapacidades(){
  try{
    // Intentar obtener de validación en RR_CAP
    const fromRR = getValidationOptionsFromA1_(SHEET_RR, RR_CAP);
    if (fromRR && fromRR.length) return fromRR;
  }catch(e){}

  // Fallback: buscar columnas 'CAPACIDAD' en Datos / Recepción / Inventario
  let out = colValuesByHeader_(SHEET_DATOS, ['CAPACIDAD','Capacidad','CAP']);
  if (out.length) return out;
  out = colValuesByHeader_(SHEET_RECEPCION, ['CAPACIDAD','Capacidad','CAP']);
  if (out.length) return out;
  out = colValuesByHeader_(SHEET_INVENTARIO, ['CAPACIDAD','Capacidad','CAP']);
  return out;
}

/** ✅ FUNCIÓN DE DEBUGGING - Verifica qué datos hay en 'Datos RR' columna D */
function apiDebugModelosDatos(){
  const result = {
    sheetExists: false,
    sheetName: 'Datos RR',
    lastRow: 0,
    rawData: [],
    uniqueData: [],
    error: null
  };

  try {
    const s = ss().getSheetByName('Datos RR');
    if (!s) {
      result.error = "La hoja 'Datos RR' no existe";
      return result;
    }
    
    result.sheetExists = true;
    result.lastRow = s.getLastRow();
    
    if (result.lastRow < 2) {
      result.error = "La hoja 'Datos RR' está vacía (sin datos a partir de fila 2)";
      return result;
    }
    
    const vals = s.getRange('D2:D' + result.lastRow).getDisplayValues().flat();
    result.rawData = vals;
    result.uniqueData = uniqueNonEmpty_(vals);
    
  } catch(e) {
    result.error = "Error al leer: " + e.message;
  }

  return result;
}

/** Devuelve lista de empresas (intenta 'Datos','Recepción','Taller'). */
function apiGetEmpresas(){
  // intentar 'Datos' columna que contenga 'EMPRESA' o bien buscar en Recepción/Taller
  // Primero intentar validación en RR (celda RR_EMPRESA)
  try{
    const fromRR = getValidationOptionsFromA1_(SHEET_RR, RR_EMPRESA);
    if (fromRR && fromRR.length) return fromRR;
  }catch(e){}

  let out = colValuesByHeader_(SHEET_DATOS, ['EMPRESA','Empresa']);
  if (out.length) return out;
  out = colValuesByHeader_(SHEET_RECEPCION, ['EMPRESA','Empresa']);
  if (out.length) return out;
  out = colValuesByHeader_(SHEET_TALLER, ['EMPRESA','Empresa']);
  return out;
}

/** Devuelve lista de nombres de recepción (intenta 'Remito de Recepción' K25 col, luego 'Recepción' header). */
function apiGetRecepciones(){
  // 1) Intentar leer validación en Remito de Recepción (K25:L25) si hay lista en hoja oculta
  try{
    // Priorizar validación en la celda (usar la primera celda del merge)
    const fromRR = getValidationOptionsFromA1_(SHEET_RR, RR_RECEP);
    if (fromRR && fromRR.length) return fromRR;

    const SRR = sh(SHEET_RR);
    const val = SRR.getRange(RR_RECEP_RANGE).getDisplayValue().toString().trim();
    const out = [];
    if (val) out.push(val);
    // fallback: buscar columna RECEPCIÓN en 'Recepción' tabla
    const fromRecep = colValuesByHeader_(SHEET_RECEPCION, ['RECEPCIÓN','Recepción','RECEPCION']);
    return uniqueNonEmpty_(out.concat(fromRecep));
  }catch(e){
    return colValuesByHeader_(SHEET_RECEPCION, ['RECEPCIÓN','Recepción','RECEPCION']);
  }
}

/** Devuelve lista de colores para el RR (prioriza validación en K15). */
function apiGetColors(){
  try{
    const fromRR = getValidationOptionsFromA1_(SHEET_RR, RR_COLOR);
    if (fromRR && fromRR.length) return fromRR;
  }catch(e){}

  // Fallback: buscar columnas 'COLOR' en Recepción / Inventario / Datos
  let out = colValuesByHeader_(SHEET_RECEPCION, ['COLOR','Color']);
  if (out.length) return out;
  out = colValuesByHeader_(SHEET_INVENTARIO, ['COLOR','Color']);
  if (out.length) return out;
  out = colValuesByHeader_(SHEET_DATOS, ['COLOR','Color']);
  return out;
}

/** Devuelve lista de valores para OTRO (prioriza validación en D27). */
function apiGetOtros(){
  try{
    // RR_OTRO_RANGE es un merge D27:L28; usamos la primera celda D27
    const fromRR = getValidationOptionsFromA1_(SHEET_RR, 'D27');
    if (fromRR && fromRR.length) return fromRR;
  }catch(e){}

  try{
    // Si la celda D27 contiene un listado separado por comas, devolverlo
    const SRR = sh(SHEET_RR);
    const val = SRR.getRange('D27').getDisplayValue().toString().trim();
    if (val && val.includes(',')) return uniqueNonEmpty_(val.split(',').map(x=>x.trim()));
  }catch(e){}

  return [];
}


function apiListarTaller(limit){
  const ST = sh(SHEET_TALLER);
  const vals = ST.getDataRange().getDisplayValues();
  if (vals.length < 2) return [];

  const headers = vals[0].map(h => String(h || '').trim());
  const colRemito = findHeaderCol1Based_(headers, REMITO_HEADER_CANDIDATES);
  const colCliente = findHeaderCol1Based_(headers, ['CLIENTE']);
  const colModelo = findHeaderCol1Based_(headers, ['MODELO']);
  const colEstadoTaller = findHeaderCol1Based_(headers, ESTADO_ACTUAL_HEADER_CANDIDATES);

  if (!colRemito) {
    throw new Error('No se encontró columna de remito en "Equipos en Taller" (ej: N° REMITO, NRO REMITO, REMITO).');
  }
  if (!colEstadoTaller) {
    throw new Error('No se encontró columna "Estado actual" en "Equipos en Taller".');
  }

  const recepcionEstadoByRemito = buildRecepcionEstadoMapByRemito_();

  const out = [];
  const lim = Number(limit || 0);
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const remitoRaw = row[colRemito - 1];
    const remitoId = normalizeRemitoId_(remitoRaw);
    if (!remitoId) continue;

    const estadoTaller = String(row[colEstadoTaller - 1] || '').trim();
    const estadoRecepcion = recepcionEstadoByRemito.get(remitoId) || '';
    const entregado = isRecepcionEntregado_(estadoRecepcion);

    out.push({
      remitoId,
      nro: String(remitoRaw || '').trim(),
      cliente: colCliente ? String(row[colCliente - 1] || '').trim() : '',
      modelo: colModelo ? String(row[colModelo - 1] || '').trim() : '',
      estado: estadoTaller,
      estadoTaller,
      estadoRecepcion,
      entregado
    });

    if (lim > 0 && out.length >= lim) break;
  }

  return out;
}

/** Devuelve todos los detalles de un equipo desde "Equipos en Taller" */
function apiGetEquipoDetalle(nro){
  try {
    const ST = sh(SHEET_TALLER);
    const rowT = findRowByKey_(ST, 'N° REMITO', nro);
    if (rowT === -1) {
      return null;
    }
    
    const mapT = getHeaderMap_(ST);
    
    const out = {
      nro: nro,
      cliente: '',
      modelo: '',
      ingreso: '',
      empresa: '',
      observaciones: '',
      tecnico: '',
      repuestos: '',
      entrega: '',
      estado: ''
    };
    
    // Mapear columnas disponibles
    if (mapT['CLIENTE']) out.cliente = ST.getRange(rowT, mapT['CLIENTE']).getDisplayValue().toString().trim();
    if (mapT['MODELO']) out.modelo = ST.getRange(rowT, mapT['MODELO']).getDisplayValue().toString().trim();
    if (mapT['INGRESO']) out.ingreso = ST.getRange(rowT, mapT['INGRESO']).getDisplayValue().toString().trim();
    if (mapT['EMPRESA']) out.empresa = ST.getRange(rowT, mapT['EMPRESA']).getDisplayValue().toString().trim();
    if (mapT['OBSERVACIONES RECEPCIÓN']) out.observaciones = ST.getRange(rowT, mapT['OBSERVACIONES RECEPCIÓN']).getDisplayValue().toString().trim();
    
    const colTec = getCol_(mapT, ['TÉCNICO','TECNICO']);
    if (colTec) out.tecnico = ST.getRange(rowT, colTec).getDisplayValue().toString().trim();
    
    const colRep = getCol_(mapT, ['REPUESTOS UTILIZADOS','REPUESTOS']);
    if (colRep) out.repuestos = ST.getRange(rowT, colRep).getDisplayValue().toString().trim();
    
    if (mapT['ENTREGA']) out.entrega = ST.getRange(rowT, mapT['ENTREGA']).getDisplayValue().toString().trim();
    if (mapT['ESTADO ACTUAL']) out.estado = ST.getRange(rowT, mapT['ESTADO ACTUAL']).getDisplayValue().toString().trim();
    
    return out;
  } catch(e) {
    console.error('Error al obtener detalles del equipo:', e);
    return null;
  }
}

/** Obtiene opciones de técnico desde Datos Desplegables (columna D) */
function apiGetTecnicosOptions(){
  try {
    return getDropdownOptionsFromSheet_(SHEET_DATOS_DESPLEGABLES, 4, 2);
  } catch(e) {
    console.error('Error apiGetTecnicosOptions:', e.toString(), e.stack);
    return [];
  }
}

/** Obtiene opciones de repuestos desde Datos Desplegables (columna F) */
function apiGetRepuestosOptions(){
  try {
    return getDropdownOptionsFromSheet_(SHEET_DATOS_DESPLEGABLES, 6, 2);
  } catch(e) {
    console.error('Error apiGetRepuestosOptions:', e.toString(), e.stack);
    return [];
  }
}

/** Obtiene opciones de Forma de Pago desde Datos Desplegables (columna H) */
function apiGetFormasPago(){
  try {
    return getDropdownOptionsFromSheet_(SHEET_DATOS_DESPLEGABLES, 8, 2);
  } catch(e) {
    console.error('Error apiGetFormasPago:', e.toString(), e.stack);
    return [];
  }
}

/** Obtiene opciones de Moneda desde Datos Desplegables (columna J) */
function apiGetMonedas(){
  try {
    return getDropdownOptionsFromSheet_(SHEET_DATOS_DESPLEGABLES, 10, 2);
  } catch(e) {
    console.error('Error apiGetMonedas:', e.toString(), e.stack);
    return [];
  }
}

/**
 * Servicios para líneas extra de RE desde Datos Desplegables columna K (fila 2 en adelante).
 * Devuelve array simple de strings no vacíos.
 */
function apiGetServiciosColumnaK(){
  try {
    return getDropdownOptionsFromSheet_(SHEET_DATOS_DESPLEGABLES, 11, 2);
  } catch(e) {
    console.error('Error apiGetServiciosColumnaK:', e.toString(), e.stack);
    return [];
  }
}

/**
 * Devuelve servicios con precio en ARS para líneas extra de RE.
 * Fuente principal: Datos Desplegables!K2:K (apiGetServiciosColumnaK).
 * Soporta formato "Servicio|80000" cuando venga así en columna K.
 */
function apiGetServiciosPesos(){
  try {
    const fromK = apiGetServiciosColumnaK();
    if (!Array.isArray(fromK) || !fromK.length) return [];
    return fromK.map(v => {
      const s = String(v || '').trim();
      if (!s) return null;
      if (s.indexOf('|') !== -1) {
        const p = s.split('|');
        const label = String(p[0] || '').trim();
        const precio = parseNumberLocal_(p[1]);
        if (!label) return null;
        return { label, precio: Number(precio || 0) };
      }
      return { label: s, precio: 0 };
    }).filter(Boolean);
  } catch(e) {
    console.error('Error apiGetServiciosPesos:', e.toString(), e.stack);
    return [];
  }
}

/** Actualiza observaciones de un equipo en Equipos en Taller */
function apiActualizarObservaciones(nro, observaciones){
  try {
    const ST = sh(SHEET_TALLER);
    const rowT = findRowByKey_(ST, 'N° REMITO', nro);
    if (rowT === -1) return false;
    
    const mapT = getHeaderMap_(ST);
    if (!mapT['OBSERVACIONES RECEPCIÓN']) return false;
    
    ST.getRange(rowT, mapT['OBSERVACIONES RECEPCIÓN']).setValue(String(observaciones || '').trim());
    return true;
  } catch(e) {
    console.error('Error al actualizar observaciones:', e);
    return false;
  }
}

/** Actualiza técnico de un equipo en Equipos en Taller */
function apiActualizarTecnico(nro, tecnico){
  try {
    const ST = sh(SHEET_TALLER);
    const rowT = findRowByKey_(ST, 'N° REMITO', nro);
    if (rowT === -1) return false;
    
    const mapT = getHeaderMap_(ST);
    const colTec = getCol_(mapT, ['TÉCNICO','TECNICO']);
    if (!colTec) return false;
    
    ST.getRange(rowT, colTec).setValue(String(tecnico || '').trim());
    return true;
  } catch(e) {
    console.error('Error al actualizar técnico:', e);
    return false;
  }
}

/** Actualiza repuestos utilizados de un equipo en Equipos en Taller */
function apiActualizarRepuestos(nro, repuestos){
  try {
    const ST = sh(SHEET_TALLER);
    const rowT = findRowByKey_(ST, 'N° REMITO', nro);
    if (rowT === -1) return false;
    
    const mapT = getHeaderMap_(ST);
    const colRep = getCol_(mapT, ['REPUESTOS UTILIZADOS','REPUESTOS']);
    if (!colRep) return false;
    
    // Convertir array a string si es necesario
    const repStr = Array.isArray(repuestos) ? repuestos.join(', ') : String(repuestos || '').trim();
    ST.getRange(rowT, colRep).setValue(repStr);
    return true;
  } catch(e) {
    console.error('Error al actualizar repuestos:', e);
    return false;
  }
}

/** Obtiene opciones de estados desde Datos Desplegables (columna B) */
function apiGetEstadosOptions(){
  try {
    return getDropdownOptionsFromSheet_(SHEET_DATOS_DESPLEGABLES, 2, 2);
  } catch(e) {
    console.error('Error apiGetEstadosOptions:', e.toString(), e.stack);
    return [];
  }
}

function getDropdownOptionsMultiCols_(sheetName, columnsMap, startRow){
  const out = {};
  Object.keys(columnsMap || {}).forEach(k => out[k] = []);
  try {
    const S = sh(sheetName);
    const firstRow = startRow || 2;
    const lastRow = S.getLastRow();
    if (lastRow < firstRow) return out;

    const cols = Object.values(columnsMap || {}).map(Number).filter(n => Number.isFinite(n) && n > 0);
    if (!cols.length) return out;

    const maxCol = Math.max.apply(null, cols);
    const values = S.getRange(firstRow, 1, lastRow - firstRow + 1, maxCol).getDisplayValues();

    Object.keys(columnsMap).forEach((key)=>{
      const colIndex = Number(columnsMap[key]) - 1;
      const set = new Set();
      for (let r = 0; r < values.length; r++) {
        const raw = values[r][colIndex];
        const v = String(raw || '').trim();
        if (v) set.add(v);
      }
      out[key] = Array.from(set);
    });
  } catch (e) {
    console.error('Error getDropdownOptionsMultiCols_:', e.toString(), e.stack);
  }
  return out;
}

/**
 * API unificada para modal de Taller: detalle + opciones en una sola llamada.
 */
function apiGetEquipoDetalleFull(nro){
  try {
    const n = Number(nro);
    if (!Number.isFinite(n) || n <= 0) {
      return { ok:false, message:'N° remito inválido.', detalle:null, tecnicos:[], repuestosOptions:[], estados:[] };
    }

    const ST = sh(SHEET_TALLER);
    const values = ST.getDataRange().getDisplayValues();
    if (!values || values.length < 2) {
      return { ok:false, message:'No hay datos en Equipos en Taller.', detalle:null, tecnicos:[], repuestosOptions:[], estados:[] };
    }

    const headers = values[0].map(h => String(h || '').trim());
    const idxByCandidates = (names)=>{
      for (let i = 0; i < names.length; i++) {
        const idx = headers.indexOf(names[i]);
        if (idx > -1) return idx;
      }
      return -1;
    };

    const idxNro = idxByCandidates(['N° REMITO','NRO REMITO','NRO']);
    if (idxNro === -1) {
      return { ok:false, message:'No existe columna N° REMITO en Equipos en Taller.', detalle:null, tecnicos:[], repuestosOptions:[], estados:[] };
    }

    let row = null;
    for (let r = 1; r < values.length; r++) {
      if (Number(values[r][idxNro]) === n) {
        row = values[r];
        break;
      }
    }

    if (!row) {
      return { ok:false, message:'Equipo no encontrado.', detalle:null, tecnicos:[], repuestosOptions:[], estados:[] };
    }

    const detalle = {
      nro: n,
      cliente: '',
      modelo: '',
      ingreso: '',
      empresa: '',
      observaciones: '',
      tecnico: '',
      repuestos: '',
      entrega: '',
      estado: ''
    };

    const idxCliente = idxByCandidates(['CLIENTE']);
    const idxModelo = idxByCandidates(['MODELO']);
    const idxIngreso = idxByCandidates(['INGRESO']);
    const idxEmpresa = idxByCandidates(['EMPRESA']);
    const idxObs = idxByCandidates(['OBSERVACIONES RECEPCIÓN','OBSERVACIONES RECEPCION','OBSERVACIONES']);
    const idxTec = idxByCandidates(['TÉCNICO','TECNICO']);
    const idxRep = idxByCandidates(['REPUESTOS UTILIZADOS','REPUESTOS']);
    const idxEntrega = idxByCandidates(['ENTREGA']);
    const idxEstado = idxByCandidates(['ESTADO ACTUAL','ESTADO']);

    if (idxCliente > -1) detalle.cliente = String(row[idxCliente] || '').trim();
    if (idxModelo > -1) detalle.modelo = String(row[idxModelo] || '').trim();
    if (idxIngreso > -1) detalle.ingreso = String(row[idxIngreso] || '').trim();
    if (idxEmpresa > -1) detalle.empresa = String(row[idxEmpresa] || '').trim();
    if (idxObs > -1) detalle.observaciones = String(row[idxObs] || '').trim();
    if (idxTec > -1) detalle.tecnico = String(row[idxTec] || '').trim();
    if (idxRep > -1) detalle.repuestos = String(row[idxRep] || '').trim();
    if (idxEntrega > -1) detalle.entrega = String(row[idxEntrega] || '').trim();
    if (idxEstado > -1) detalle.estado = String(row[idxEstado] || '').trim();

    const options = getDropdownOptionsMultiCols_(SHEET_DATOS_DESPLEGABLES, {
      estados: 2,
      tecnicos: 4,
      repuestosOptions: 6
    }, 2);

    return {
      ok: true,
      detalle,
      tecnicos: options.tecnicos || [],
      repuestosOptions: options.repuestosOptions || [],
      estados: options.estados || []
    };
  } catch(e) {
    console.error('Error apiGetEquipoDetalleFull:', e.toString(), e.stack);
    return { ok:false, message:String((e && e.message) ? e.message : e), detalle:null, tecnicos:[], repuestosOptions:[], estados:[] };
  }
}

function _findRowByRemitoFlexible_(sheetValues, colRemito, remitoId){
  const target = normalizeRemitoId_(remitoId);
  if (!target) return -1;
  for (let r = 1; r < sheetValues.length; r++) {
    const current = normalizeRemitoId_(sheetValues[r][colRemito - 1]);
    if (current && current === target) return r + 1;
  }
  return -1;
}

function _getTallerCols_(headers){
  const colRemito = findHeaderCol1Based_(headers, REMITO_HEADER_CANDIDATES);
  const colEstado = findHeaderCol1Based_(headers, ESTADO_ACTUAL_HEADER_CANDIDATES);
  return { colRemito, colEstado };
}

/**
 * Actualiza estado de taller por remito con validación robusta y lock para evitar race conditions.
 * Escribe en hoja "Equipos en Taller" columna Estado actual (habitualmente I).
 */
function apiUpdateEstadoTaller(remitoId, nuevoEstado){
  const estado = String(nuevoEstado || '').trim();
  if (!estado) return { ok:false, message:'Estado inválido.' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(7000);

    const ST = sh(SHEET_TALLER);
    const values = ST.getDataRange().getDisplayValues();
    if (!values || values.length < 2) {
      return { ok:false, message:'No hay datos en "Equipos en Taller".' };
    }

    const headers = values[0].map(h => String(h || '').trim());
    const cols = _getTallerCols_(headers);
    if (!cols.colRemito) {
      return { ok:false, message:'No se encontró columna de remito en "Equipos en Taller".' };
    }
    if (!cols.colEstado) {
      return { ok:false, message:'No se encontró columna "Estado actual" en "Equipos en Taller".' };
    }

    const row = _findRowByRemitoFlexible_(values, cols.colRemito, remitoId);
    if (row === -1) {
      return { ok:false, message:'Equipo no encontrado para el remito indicado.' };
    }

    const current = String(values[row - 1][cols.colEstado - 1] || '').trim();
    if (normalizeStr_(current) === normalizeStr_(estado)) {
      return {
        ok:true,
        remitoId: normalizeRemitoId_(remitoId),
        row,
        estado: current || estado,
        unchanged: true
      };
    }

    ST.getRange(row, cols.colEstado).setValue(estado);
    SpreadsheetApp.flush();

    const confirmado = ST.getRange(row, cols.colEstado).getDisplayValue().toString().trim();
    if (normalizeStr_(confirmado) !== normalizeStr_(estado)) {
      return { ok:false, message:'No se pudo confirmar la actualización del estado en la hoja.' };
    }

    return {
      ok:true,
      remitoId: normalizeRemitoId_(remitoId),
      row,
      estado: confirmado,
      unchanged: false
    };
  } catch (e) {
    return {
      ok:false,
      message: String((e && e.message) ? e.message : e)
    };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Endpoint liviano para Kanban optimista:
 * - Busca por remito
 * - Actualiza solo la celda de estado en "Equipos en Taller"
 * - Devuelve únicamente bandera de éxito
 */
function updateEstado(remitoId, nuevoEstado){
  const estadoRaw = String(nuevoEstado || '').trim();
  const estadoNorm = normalizeStr_(estadoRaw);
  let estado = estadoRaw;
  if (estadoNorm.includes('sin reparacion')) estado = 'Sin Reparación';
  else if (estadoNorm.includes('en reparacion') || estadoNorm === 'reparacion') estado = 'En Reparación';
  else if (estadoNorm.includes('pendiente')) estado = 'Pendiente Diagnóstico';
  else if (estadoNorm.includes('presupuesto')) estado = 'Presupuesto espera Aprobación';
  else if (estadoNorm.includes('terminado')) estado = 'Terminado';

  if (!estado) return { success:false, message:'Estado inválido.' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);

    const ST = sh(SHEET_TALLER);
    const lastRow = ST.getLastRow();
    const lastCol = ST.getLastColumn();
    if (lastRow < 2) return { success:false, message:'No hay datos en Equipos en Taller.' };

    const headers = ST.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
    const colRemito = findHeaderCol1Based_(headers, REMITO_HEADER_CANDIDATES);
    const colEstado = findHeaderCol1Based_(headers, ESTADO_ACTUAL_HEADER_CANDIDATES);

    if (!colRemito) return { success:false, message:'No se encontró columna de remito.' };
    if (!colEstado) return { success:false, message:'No se encontró columna Estado actual.' };

    const remitoNorm = normalizeRemitoId_(remitoId);
    if (!remitoNorm) return { success:false, message:'Remito inválido.' };

    const remitosCol = ST.getRange(2, colRemito, lastRow - 1, 1).getDisplayValues();
    let targetRow = -1;
    for (let i = 0; i < remitosCol.length; i++) {
      const current = normalizeRemitoId_(remitosCol[i][0]);
      if (current && current === remitoNorm) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) return { success:false, message:'Equipo no encontrado.' };

    ST.getRange(targetRow, colEstado).setValue(estado);
    return { success:true };
  } catch (e) {
    return { success:false, message:String((e && e.message) ? e.message : e) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** Actualiza el estado de un equipo en Equipos en Taller */
function apiActualizarEstadoEquipo(nro, nuevoEstado){
  const res = apiUpdateEstadoTaller(nro, nuevoEstado);
  return !!(res && res.ok);
}

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Unión Tecno · Sistema de Gestión')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// helper para includes
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ===========================
 * APIs NUEVAS para UI (Index.html)
 * No modifican tu lógica existente. Solo escriben/leen datos.
 * ===========================
 */

/**
 * Guarda el formulario de RR escribiendo en las celdas reales de la hoja "Remito de Recepción".
 * Luego la UI puede llamar apiGenerarRR() para ejecutar el flujo completo.
 *
 * data = {
 *  cliente,dni,tel,dir,marca,modelo,cap,color,imei,obs,empresa,recep,otro
 * }
 */
function apiRRGuardarFormulario(data){
  const SRR = sh(SHEET_RR);

  data = data || {};
  const safe = v => String(v || '').toString();

  // Escribir inputs RR
  writeA1(SRR, RR_CLIENTE, safe(data.cliente));
  writeA1(SRR, RR_DNI, safe(data.dni));
  writeA1(SRR, RR_TEL, safe(data.tel));
  writeA1(SRR, RR_DIR, safe(data.dir));
  writeA1(SRR, RR_MARCA, safe(data.marca));
  writeA1(SRR, RR_MODELO, safe(data.modelo));
  writeA1(SRR, RR_CAP, safe(data.cap));
  writeA1(SRR, RR_COLOR, safe(data.color));
  writeA1(SRR, RR_IMEI, safe(data.imei));

  // Observaciones (merge)
  SRR.getRange('D18:L18').setValue(safe(data.obs));

  // Empresa / Recepción (merge)
  SRR.getRange(RR_EMPRESA_RANGE).setValue(safe(data.empresa));
  SRR.getRange(RR_RECEP_RANGE).setValue(safe(data.recep));

  // OTRO (merge)
  SRR.getRange(RR_OTRO_RANGE).setValue(safe(data.otro));

  return true;
}

/**
 * Setea el N° REMITO en la hoja "Remito de Entrega" (L2) desde la UI.
 */
function apiRESetNro(nro){
  const SRE = sh(SHEET_RE);
  const n = Number(nro);
  if (!n || n <= 0) throw new Error('N° remito inválido.');
  writeA1(SRE, RE_NRO, n);
  return true;
}

function clearREAutoRowsCache_(nro){
  try {
    const n = Number(nro);
    if (!Number.isFinite(n) || n <= 0) return;
    CacheService.getScriptCache().remove(`re_auto_rows_${n}`);
  } catch (_) {}
}

function getREHydrateMarkers_(SRE){
  const cliente = SRE.getRange(RE_CLIENTE).getDisplayValue().toString().trim();
  const dni = SRE.getRange(RE_DNI).getDisplayValue().toString().trim();
  const modelo = SRE.getRange(RE_MODELO).getDisplayValue().toString().trim();
  return { cliente, dni, modelo };
}

function isREHydrateComputed_(markers){
  return !!(markers && markers.cliente && (markers.dni || markers.modelo));
}

/**
 * Endpoint unificado para hidratar RE en un solo roundtrip.
 * - Setea L2
 * - Espera cálculo de fórmulas (polling acotado)
 * - Devuelve det + extras + autoRows + timings
 */
function apiREHydrate(nro){
  const t0 = Date.now();
  const timings = {
    setNroMs: 0,
    flushWaitMs: 0,
    readDetMs: 0,
    readExtrasMs: 0,
    readAutoMs: 0,
    totalMs: 0
  };

  try {
    const n = Number(nro);
    if (!Number.isFinite(n) || n <= 0) {
      return { ok:false, computed:false, det:null, extras:[], autoRows:[], timings, message:'N° remito inválido.' };
    }

    const SRE = sh(SHEET_RE);

    const tSet = Date.now();
    writeA1(SRE, RE_NRO, n);
    timings.setNroMs = Date.now() - tSet;

    const tWait = Date.now();
    let computed = false;
    let markers = { cliente:'', dni:'', modelo:'' };
    SpreadsheetApp.flush();
    for (let i = 0; i < 8; i++) {
      markers = getREHydrateMarkers_(SRE);
      if (isREHydrateComputed_(markers)) {
        computed = true;
        break;
      }
      Utilities.sleep(180);
      SpreadsheetApp.flush();
    }
    timings.flushWaitMs = Date.now() - tWait;

    const tDet = Date.now();
    const det = apiREReadFormData();
    timings.readDetMs = Date.now() - tDet;

    const tExtras = Date.now();
    const extras = apiREReadLineasExtra(n);
    timings.readExtrasMs = Date.now() - tExtras;

    const tAuto = Date.now();
    const autoRows = apiREReadPreciosAuto();
    timings.readAutoMs = Date.now() - tAuto;

    timings.totalMs = Date.now() - t0;

    console.log('[apiREHydrate] timings:', JSON.stringify(timings), 'computed:', computed, 'markers:', JSON.stringify(markers));

    return {
      ok: true,
      computed,
      det: det || null,
      extras: Array.isArray(extras) ? extras : [],
      autoRows: Array.isArray(autoRows) ? autoRows : [],
      timings
    };
  } catch (e) {
    timings.totalMs = Date.now() - t0;
    console.error('[apiREHydrate] Error:', e.toString(), e.stack, 'timings:', JSON.stringify(timings));
    return {
      ok: false,
      computed: false,
      det: null,
      extras: [],
      autoRows: [],
      timings,
      message: String((e && e.message) ? e.message : e)
    };
  }
}

/**
 * Devuelve un “detalle unificado” del remito para la UI:
 * - Datos principales del equipo (marca/modelo/imei)
 * - Estado y técnico (desde Equipos en Taller)
 * - Links RR/RE (desde Historial)
 * - DNI/Tel/WhatsApp (desde Recepción + Clientes)
 */
function apiGetRemitoDetalle(nro){
  const n = Number(nro);
  if (!n || n <= 0) throw new Error('N° remito inválido.');

  const out = {
    nro: n,
    cliente: '',
    dni: '',
    tel: '',
    direccion: '',
    whatsapp: '',
    marca: '',
    modelo: '',
    imei: '',
    estado: '',
    tecnico: '',
    repuestos: [],
    links: { rr:'', re:'' }
  };

  // 1) Taller: estado + tecnico + datos equipo + repuestos (si existen como columnas)
  {
    const ST = sh(SHEET_TALLER);
    const rowT = findRowByKey_(ST, 'N° REMITO', n);
    if (rowT !== -1){
      const mapT = getHeaderMap_(ST);

      if (mapT['CLIENTE']) out.cliente = ST.getRange(rowT, mapT['CLIENTE']).getDisplayValue().toString().trim();
      if (mapT['MARCA']) out.marca = ST.getRange(rowT, mapT['MARCA']).getDisplayValue().toString().trim();
      if (mapT['MODELO']) out.modelo = ST.getRange(rowT, mapT['MODELO']).getDisplayValue().toString().trim();
      if (mapT['IMEI']) out.imei = ST.getRange(rowT, mapT['IMEI']).getDisplayValue().toString().trim();
      if (mapT['ESTADO ACTUAL']) out.estado = ST.getRange(rowT, mapT['ESTADO ACTUAL']).getDisplayValue().toString().trim();

      const colTec = getCol_(mapT, ['TÉCNICO','TECNICO']);
      if (colTec) out.tecnico = ST.getRange(rowT, colTec).getDisplayValue().toString().trim();

      // Traer repuestos utilizados - convertir string a array
      const colRep = getCol_(mapT, ['REPUESTOS UTILIZADOS','REPUESTOS']);
      if (colRep) {
        const repStr = ST.getRange(rowT, colRep).getDisplayValue().toString().trim();
        if (repStr) {
          // Dividir por comas/punto y coma y limpiar espacios
          out.repuestos = repStr.split(/[,;]/).map(r => r.trim()).filter(r => r);
        }
      }
    }
  }

  // 2) Historial: links RR / RE
  {
    const SH = sh(SHEET_HISTORIAL);
    const rowH = findRowByKey_(SH, 'N° REMITO', n);
    if (rowH !== -1){
      const mapH = getHeaderMap_(SH);

      const colRR = getCol_(mapH, ['LINK DE REMITO DE RECEPCIÓN','LINK REMITO DE RECEPCION','LINK RR']);
      const colRE = getCol_(mapH, ['LINK DE REMITO DE ENTREGA','LINK REMITO DE ENTREGA','LINK RE']);

      if (colRR) out.links.rr = SH.getRange(rowH, colRR).getDisplayValue().toString().trim();
      if (colRE) out.links.re = SH.getRange(rowH, colRE).getDisplayValue().toString().trim();

      if (!out.cliente && mapH['CLIENTE']) out.cliente = SH.getRange(rowH, mapH['CLIENTE']).getDisplayValue().toString().trim();
    }
  }

  // 3) Recepción: DNI (y a veces tel) por N° remito
  {
    const SR = sh(SHEET_RECEPCION);
    const rowR = findRowByKey_(SR, 'N° REMITO', n);
    if (rowR !== -1){
      const mapR = getHeaderMap_(SR);

      if (mapR['DNI']) out.dni = SR.getRange(rowR, mapR['DNI']).getDisplayValue().toString().trim();
      if (!out.cliente && mapR['CLIENTE']) out.cliente = SR.getRange(rowR, mapR['CLIENTE']).getDisplayValue().toString().trim();

      // Si en Recepción tuvieras Teléfono, lo tomamos (si no, lo buscamos en Clientes)
      const colTel = getCol_(mapR, ['TELÉFONO','TELEFONO','TEL']);
      if (colTel) out.tel = SR.getRange(rowR, colTel).getDisplayValue().toString().trim();

      if (!out.marca && mapR['MARCA']) out.marca = SR.getRange(rowR, mapR['MARCA']).getDisplayValue().toString().trim();
      if (!out.modelo && mapR['MODELO']) out.modelo = SR.getRange(rowR, mapR['MODELO']).getDisplayValue().toString().trim();
      if (!out.imei && mapR['IMEI']) out.imei = SR.getRange(rowR, mapR['IMEI']).getDisplayValue().toString().trim();
    }
  }

  // 4) Clientes: tel + direccion + link whatsapp usando DNI
  if (out.dni){
    const SC = sh(SHEET_CLIENTES);
    const rowC = findRowByKey_(SC, 'DNI', out.dni);
    if (rowC !== -1){
      const mapC = getHeaderMap_(SC);

      const colTel = getCol_(mapC, ['TELÉFONO','TELEFONO','TEL']);
      if (colTel && !out.tel) out.tel = SC.getRange(rowC, colTel).getDisplayValue().toString().trim();

      const colDir = getCol_(mapC, ['DIRECCIÓN','DIRECCION','DIRECCIO','DIR']);
      if (colDir) out.direccion = SC.getRange(rowC, colDir).getDisplayValue().toString().trim();

      const colWa = getCol_(mapC, ['LINK WHATSAPP','WHATSAPP','LINK WA']);
      if (colWa) out.whatsapp = SC.getRange(rowC, colWa).getDisplayValue().toString().trim();

      // fallback: si no existe el link pero hay tel, lo generamos igual
      if (!out.whatsapp && out.tel) out.whatsapp = buildWhatsAppLink_(out.tel);
    } else {
      // fallback: si no existe el cliente, pero hay tel (por recepción), generamos link
      if (out.tel) out.whatsapp = buildWhatsAppLink_(out.tel);
    }
  } else {
    if (out.tel) out.whatsapp = buildWhatsAppLink_(out.tel);
  }

  return out;
}

/**
 * Lee los datos de los campos de la hoja "Remito de Entrega" en tiempo real.
 * Combina valores de las celdas con fórmulas + datos obtenidos directamente desde las otras hojas.
 */
/**
 * FUNCIÓN DE DEBUG COMPLETA: Verifica todo el flujo
 */
function apiDebugRERead(){
  try {
    const SRE = sh(SHEET_RE);
    const SR = sh(SHEET_RECEPCION);
    const SC = sh(SHEET_CLIENTES);
    const ST = sh(SHEET_TALLER);
    
    const nro = Number(readA1(SRE, 'L2'));
    
    const debug = {
      nroEnL2: nro,
      formulasEnRE: {
        B9: SRE.getRange('B9').getFormula(),
        I9: SRE.getRange('I9').getFormula(),
        K9: SRE.getRange('K9').getFormula(),
        D11: SRE.getRange('D11').getFormula(),
        D14: SRE.getRange('D14').getFormula(),
        K14: SRE.getRange('K14').getFormula(),
        D15: SRE.getRange('D15').getFormula()
      },
      valoresEnRE: {
        B9: SRE.getRange('B9').getDisplayValue().toString().trim(),
        I9: SRE.getRange('I9').getDisplayValue().toString().trim(),
        K9: SRE.getRange('K9').getDisplayValue().toString().trim(),
        D11: SRE.getRange('D11').getDisplayValue().toString().trim(),
        D14: SRE.getRange('D14').getDisplayValue().toString().trim(),
        K14: SRE.getRange('K14').getDisplayValue().toString().trim(),
        D15: SRE.getRange('D15').getDisplayValue().toString().trim()
      },
      hojaRecepcionExists: true,
      hojaRecepcionRows: SR.getLastRow(),
      hojaClientesExists: true,
      hojaClientesRows: SC.getLastRow(),
      hojaTallerExists: true,
      hojaTallerRows: ST.getLastRow()
    };
    
    // Buscar este nro en Recepción
    const rowR = findRowByKey_(SR, 'N° REMITO', nro);
    debug.nroEnRecepcion = rowR !== -1 ? 'ENCONTRADO en fila ' + rowR : 'NO ENCONTRADO';
    
    if (rowR !== -1) {
      const mapR = getHeaderMap_(SR);
      debug.datosEnRecepcion = {
        cliente: SR.getRange(rowR, mapR['CLIENTE']).getDisplayValue().toString().trim(),
        dni: SR.getRange(rowR, mapR['DNI']).getDisplayValue().toString().trim(),
        marca: SR.getRange(rowR, mapR['MARCA']).getDisplayValue().toString().trim(),
        modelo: SR.getRange(rowR, mapR['MODELO']).getDisplayValue().toString().trim(),
        imei: SR.getRange(rowR, mapR['IMEI']).getDisplayValue().toString().trim()
      };
    }
    
    console.log('[apiDebugRERead] DEBUG COMPLETO:', debug);
    return debug;
  } catch(e) {
    console.error('[apiDebugRERead] Error:', e.toString(), e.stack);
    return { error: e.toString(), stack: e.stack };
  }
}

function apiREReadFormData(){
  try {
    const SRE = sh(SHEET_RE);
    const nro = Number(readA1(SRE, RE_NRO));
    
    console.log('[apiREReadFormData] nro leído de L2:', nro);
    console.log('[apiREReadFormData] RE_NRO constante:', RE_NRO);
    console.log('[apiREReadFormData] RE_CLIENTE constante:', RE_CLIENTE);
    
    if (!nro || nro <= 0) {
      console.warn('[apiREReadFormData] nro inválido:', nro);
      return null;
    }
    
    // Leer cliente desde el rango B9:H9 (pueden estar mergeadas)
    let clienteVal = '';
    try {
      const clienteRange = SRE.getRange(RE_CLIENTE);
      clienteVal = clienteRange.getDisplayValue().toString().trim();
      console.log('[apiREReadFormData] Cliente leído:', clienteVal);
    } catch(e) {
      console.warn('[apiREReadFormData] Error al leer cliente:', e.toString());
      clienteVal = '';
    }
    
    // Lectura batch para campos simples en RE (evita múltiples roundtrips internos)
    const valueByA1 = {};
    const fieldsA1 = [RE_DNI, RE_TEL, RE_DIRECCION, RE_MARCA, RE_MODELO, RE_IMEI, RE_FECHA];
    const rangeList = SRE.getRangeList(fieldsA1).getRanges();
    rangeList.forEach((rng, i) => {
      valueByA1[fieldsA1[i]] = rng.getDisplayValue().toString().trim();
    });

    const dniVal = valueByA1[RE_DNI] || '';
    const telVal = valueByA1[RE_TEL] || '';
    const direccionVal = valueByA1[RE_DIRECCION] || '';
    const marcaVal = valueByA1[RE_MARCA] || '';
    const modeloVal = valueByA1[RE_MODELO] || '';
    const imeiVal = valueByA1[RE_IMEI] || '';
    let estadoActualVal = '';

    try {
      const ST = sh(SHEET_TALLER);
      const rowT = findRowByKey_(ST, 'N° REMITO', nro);
      if (rowT !== -1) {
        const mapT = getHeaderMap_(ST);
        const colEstado = getCol_(mapT, ['ESTADO ACTUAL', 'ESTADO']);
        if (colEstado) {
          estadoActualVal = ST.getRange(rowT, colEstado).getDisplayValue().toString().trim();
        }
      }
    } catch (e) {
      console.warn('[apiREReadFormData] No se pudo leer estado actual:', e.toString());
    }
    
    console.log('[apiREReadFormData] DNI:', dniVal);
    console.log('[apiREReadFormData] TEL:', telVal);
    console.log('[apiREReadFormData] DIRECCION:', direccionVal);
    console.log('[apiREReadFormData] MARCA:', marcaVal);
    console.log('[apiREReadFormData] MODELO:', modeloVal);
    console.log('[apiREReadFormData] IMEI:', imeiVal);
    
    const out = {
      nro: nro,
      fecha: valueByA1[RE_FECHA] || '',   // L3
      cliente: clienteVal,                // B9:H9 (merged)
      dni: dniVal,                        // I9
      tel: telVal,                        // K9
      direccion: direccionVal,            // D11
      marca: marcaVal,                    // D14
      modelo: modeloVal,                  // K14
      imei: imeiVal,                      // D15
      estadoActual: estadoActualVal,
      repuestos: []
    };
    
    console.log('[apiREReadFormData] Datos leídos de RE:', out);
    
    // Leer repuestos de la tabla de detalles (filas 18-35, columnas B-E) en batch
    let repuestosSet = new Set();
    const catValues = SRE.getRange(RE_ROW_START, 2, RE_ROW_END - RE_ROW_START + 1, 4).getDisplayValues();
    for (let r = 0; r < catValues.length; r++) {
      for (let c = 0; c < catValues[r].length; c++) {
        const val = String(catValues[r][c] || '').trim();
        if (val) repuestosSet.add(val);
      }
    }
    
    out.repuestos = Array.from(repuestosSet).filter(r => r);
    console.log('[apiREReadFormData] Repuestos encontrados:', out.repuestos);
    
    console.log('[apiREReadFormData] Datos finales completos:', out);
    return out;
  } catch(e) {
    console.error('Error en apiREReadFormData:', e.toString(), e.stack);
    return null;
  }
}

/**
 * Lee líneas automáticas de RE exactamente desde:
 * B18:B31 detalle | C18:C31 cantidad | K18:K31 precio
 * Filtra filas sin detalle.
 */
function apiREReadPreciosAuto(){
  try {
    const SRE = sh(SHEET_RE);
    const nro = Number(readA1(SRE, RE_NRO));
    if (!nro || nro <= 0) return [];

    const cache = CacheService.getScriptCache();
    const cacheKey = `re_auto_rows_${nro}`;
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        const parsed = JSON.parse(cached);
        if (Array.isArray(parsed)) return parsed;
      } catch (_) {}
    }

    const out = [];
    const totalRows = 14; // 18..31
    const detalles = SRE.getRange(18, 2, totalRows, 4).getDisplayValues(); // B:E
    const cantidades = SRE.getRange(18, 10, totalRows, 1).getDisplayValues(); // J
    const precios = SRE.getRange(18, 11, totalRows, 1).getDisplayValues(); // K

    for (let i = 0; i < totalRows; i++) {
      const rowNum = 18 + i;
      const detalle = (detalles[i] || []).map(v => String(v || '').trim()).find(v => v) || '';
      if (!detalle) continue;

      const cantidadDisplay = String((cantidades[i] && cantidades[i][0]) || '').trim();
      const cantidadNum = parseNumberLocal_(cantidadDisplay);
      const cantidad = cantidadNum ? Math.max(1, Math.min(99, Math.round(cantidadNum))) : 1;

      const precioDisplay = String((precios[i] && precios[i][0]) || '').trim();

      out.push({ row:rowNum, detalle, cantidad, precio:precioDisplay, precioDisplay });
    }

    cache.put(cacheKey, JSON.stringify(out), 90);
    return out;
  } catch(e) {
    console.error('Error apiREReadPreciosAuto:', e.toString(), e.stack);
    return [];
  }
}

/**
 * Escribe cantidad de una línea automática por índice fijo (0..13 => filas 18..31).
 */
function apiREWriteAutoQty(nro, autoIndex, cantidad){
  try {
    const SRE = sh(SHEET_RE);
    const n = Number(nro);
    if (!n || n <= 0) throw new Error('N° remito inválido.');
    if (Number(readA1(SRE, RE_NRO)) !== n) writeA1(SRE, RE_NRO, n);

    const idx = Number(autoIndex);
    if (!Number.isFinite(idx) || idx < 0) throw new Error('Índice auto inválido.');
    const qty = Math.max(1, Math.min(99, Math.round(Number(cantidad) || 1)));

    const row = 18 + idx;
    if (row < 18 || row > 31) throw new Error('Índice auto fuera de rango.');
    SRE.getRange(`J${row}`).setValue(qty);
    clearREAutoRowsCache_(n);
    SpreadsheetApp.flush();

    return { ok:true, row, cantidad:qty };
  } catch(e) {
    console.error('Error apiREWriteAutoQty:', e.toString(), e.stack);
    return { ok:false, error:e.toString() };
  }
}

/**
 * Endpoint legacy: no escribe precios en RE (K/L es solo fórmula).
 */
function apiREWriteAutoPrecio(nro, autoIndex, precio){
  try {
    const SRE = sh(SHEET_RE);
    const n = Number(nro);
    if (!n || n <= 0) throw new Error('N° remito inválido.');
    if (Number(readA1(SRE, RE_NRO)) !== n) writeA1(SRE, RE_NRO, n);

    const idx = Number(autoIndex);
    if (!Number.isFinite(idx) || idx < 0 || idx > 13) throw new Error('Índice auto fuera de rango.');
    const row = 18 + idx;
    SpreadsheetApp.flush();

    return { ok:true, row, ignoredPrice:true };
  } catch(e) {
    console.error('Error apiREWriteAutoPrecio:', e.toString(), e.stack);
    return { ok:false, error:e.toString() };
  }
}

/**
 * Escribe líneas automáticas en lote.
 * lines = [{idx, cantidad, precio}] o [{row, cantidad, precio}]
 * El campo precio se ignora para proteger fórmulas en K/L.
 */
function apiREWriteAutoLines(nro, lines){
  try {
    const SRE = sh(SHEET_RE);
    const n = Number(nro);
    if (!n || n <= 0) throw new Error('N° remito inválido.');
    if (Number(readA1(SRE, RE_NRO)) !== n) writeA1(SRE, RE_NRO, n);

    const normalized = Array.isArray(lines) ? lines : [];
    let written = 0;

    normalized.forEach((item)=>{
      const obj = (item && typeof item === 'object') ? item : {};

      let row = Number(obj.row);
      if (!Number.isFinite(row) || row < 18 || row > 31) {
        const idx = Number(obj.idx);
        if (Number.isFinite(idx)) row = 18 + idx;
      }
      if (!Number.isFinite(row) || row < 18 || row > 31) return;

      const hasCantidad = Object.prototype.hasOwnProperty.call(obj, 'cantidad');

      if (hasCantidad) {
        const qty = Math.max(1, Math.min(99, Math.round(Number(obj.cantidad) || 1)));
        SRE.getRange(`J${row}`).setValue(qty);
      }

      if (hasCantidad) written++;
    });

    clearREAutoRowsCache_(n);
    SpreadsheetApp.flush();
    return { ok:true, written };
  } catch(e) {
    console.error('Error apiREWriteAutoLines:', e.toString(), e.stack);
    return { ok:false, error:e.toString() };
  }
}

/**
 * Lee precios de extras desde K32:K35 (mapeo fijo por índice visible).
 */
function apiREReadExtraPrices(nro){
  try {
    const SRE = sh(SHEET_RE);
    const n = Number(nro || readA1(SRE, RE_NRO));
    if (!n || n <= 0) return { ok:false, prices:[] };
    if (Number(readA1(SRE, RE_NRO)) !== n) writeA1(SRE, RE_NRO, n);

    const prices = [];
    for (let r = RE_EXTRA_ROW_START; r <= RE_EXTRA_ROW_END; r++) {
      const precioDisplay = SRE.getRange(`K${r}`).getDisplayValue().toString().trim();
      const precioNum = parseNumberLocal_(precioDisplay);
      prices.push({
        idx: r - RE_EXTRA_ROW_START,
        row: r,
        precioDisplay: precioDisplay,
        precio: Number(precioNum || 0)
      });
    }
    return { ok:true, prices };
  } catch(e) {
    console.error('Error apiREReadExtraPrices:', e.toString(), e.stack);
    return { ok:false, prices:[], error:e.toString() };
  }
}

/**
 * Escribe una línea extra puntual en RE (idx 0..3 => filas 32..35).
 */
function apiREWriteExtraLine(nro, idx, payload){
  try {
    const SRE = sh(SHEET_RE);
    const n = Number(nro);
    if (!n || n <= 0) throw new Error('N° remito inválido.');
    if (Number(readA1(SRE, RE_NRO)) !== n) writeA1(SRE, RE_NRO, n);

    const i = Number(idx);
    if (!Number.isFinite(i) || i < 0 || i >= 4) throw new Error('Índice extra inválido.');
    const row = RE_EXTRA_ROW_START + i;

    const data = payload || {};
    const detalle = String(data.detalle || '').trim();
    const cantidad = Math.max(1, Math.min(99, Math.round(Number(data.cantidad) || 1)));

    if (!detalle) {
      SRE.getRange(`B${row}:E${row}`).clearContent();
      SRE.getRange(`${RE_QTY_COL}${row}`).clearContent();
    } else {
      SRE.getRange(`B${row}`).setValue(detalle);
      SRE.getRange(`${RE_QTY_COL}${row}`).setValue(cantidad);
    }

    clearREAutoRowsCache_(n);
    SpreadsheetApp.flush();
    const prices = apiREReadExtraPrices(n).prices || [];
    return { ok:true, idx:i, row, prices };
  } catch(e) {
    console.error('Error apiREWriteExtraLine:', e.toString(), e.stack);
    return { ok:false, error:e.toString() };
  }
}

/**
 * Lee líneas extra de RE (filas 32-35) en formato para UI.
 * Devuelve: [{detalle, cantidad, precio}, ...]
 */
function apiREReadLineasExtra(nro){
  try {
    const SRE = sh(SHEET_RE);
    const n = Number(nro || readA1(SRE, RE_NRO));
    if (!n || n <= 0) return [];

    const out = [];
    const rowCount = RE_EXTRA_ROW_END - RE_EXTRA_ROW_START + 1;
    const detalles = SRE.getRange(RE_EXTRA_ROW_START, 2, rowCount, 4).getDisplayValues(); // B:E
    const cantidades = SRE.getRange(RE_EXTRA_ROW_START, 10, rowCount, 1).getDisplayValues(); // J
    const precios = SRE.getRange(RE_EXTRA_ROW_START, 11, rowCount, 1).getDisplayValues(); // K

    for (let i = 0; i < rowCount; i++) {
      const r = RE_EXTRA_ROW_START + i;
      const detalleParts = (detalles[i] || []).map(v => String(v || '').trim());
      const detalle = detalleParts.find(v => v) || '';
      if (!detalle) continue;

      const cantidadDisplay = String((cantidades[i] && cantidades[i][0]) || '').trim();
      const cantidad = Math.max(1, Math.min(99, Math.round(parseNumberLocal_(cantidadDisplay) || 1)));
      const precio = String((precios[i] && precios[i][0]) || '').trim();

      out.push({ idx: r - RE_EXTRA_ROW_START, row:r, detalle, cantidad, precio });
    }
    return out;
  } catch(e) {
    console.error('Error apiREReadLineasExtra:', e.toString(), e.stack);
    return [];
  }
}

/**
 * Guarda líneas extra en RE (filas 32-35), compactadas sin huecos.
 * extras = [{detalle, cantidad, precio}, ...]
 */
function apiREGuardarLineasExtra(nro, extras){
  const SRE = sh(SHEET_RE);
  const n = Number(nro);
  if (!n || n <= 0) throw new Error('N° remito inválido para guardar líneas extra.');

  const nroActual = Number(readA1(SRE, RE_NRO));
  if (nroActual !== n) writeA1(SRE, RE_NRO, n);

  const normalized = (Array.isArray(extras) ? extras : [])
    .map(x => ({
      detalle: String((x && x.detalle) || '').trim(),
      cantidad: Math.max(1, Math.min(99, Math.round(Number((x && x.cantidad) || 1) || 1)))
    }))
    .filter(x => x.detalle)
    .slice(0, 4);

  SRE.getRange(`B${RE_EXTRA_ROW_START}:E${RE_EXTRA_ROW_END}`).clearContent();
  SRE.getRange(`${RE_QTY_COL}${RE_EXTRA_ROW_START}:${RE_QTY_COL}${RE_EXTRA_ROW_END}`).clearContent();

  normalized.forEach((row, i) => {
    const r = RE_EXTRA_ROW_START + i;
    SRE.getRange(`B${r}`).setValue(row.detalle);
    SRE.getRange(`${RE_QTY_COL}${r}`).setValue(row.cantidad);
  });

  clearREAutoRowsCache_(n);
  SpreadsheetApp.flush();
  return { success:true, count: normalized.length, prices: (apiREReadExtraPrices(n).prices || []) };
}
