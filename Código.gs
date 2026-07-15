const APP_NAME = 'AjuTroquel';
const SPREADSHEET_ID = '1uCM81cY3GgmD2Atr8Y4x6qhWx8y1sxI0tORGs0bM1vc';
const BASE_MEASURE_HEADERS = ['units', 'object', 'control', 'nominal', 'tolerance'];
const PART_RE = /^207115[45](-00)?$/;

function doGet() {
  let output;
  try {
    output = HtmlService.createHtmlOutputFromFile('Index');
  } catch (e) {
    output = HtmlService.createHtmlOutputFromFile('index');
  }
  return output.setTitle(APP_NAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppConfig() {
  const sheets = getSpreadsheet_().getSheets().filter(s => !s.isSheetHidden()).map(s => {
    const headers = getHeaders_(s, false);
    const mode = isMeasurementSheet_(headers) ? 'measurement' : 'table';
    return { name: s.getName(), headers: headers, mode: mode, rowCount: Math.max(s.getLastRow() - 1, 0) };
  }).filter(s => s.headers.length);
  return {
    appName: APP_NAME,
    spreadsheetId: SPREADSHEET_ID,
    parts: sheets.filter(s => s.mode === 'measurement' && PART_RE.test(s.name)).map(s => ({ name: s.name, objectCount: getMeasurementPlan(s.name).objectRows.length })),
    sheets: sheets
  };
}

function getMeasurementPlan(sheetName) {
  const sheet = getSheet_(resolvePartSheetName_(sheetName));
  const headers = getHeaders_(sheet, true);
  if (!isMeasurementSheet_(headers)) throw new Error('La hoja ' + sheet.getName() + ' no tiene encabezados de medicion.');
  const rows = getSheetRows_(sheet, Math.max(sheet.getLastColumn(), BASE_MEASURE_HEADERS.length), true);
  return {
    sheetName: sheet.getName(),
    headers: headers,
    rows: rows,
    objectRows: getMeasurementObjectRows_(headers, rows),
    captureColumns: getCaptureColumns_(headers, rows)
  };
}

function saveFaroCapture(payload) {
  payload = payload || {};
  const sheet = getSheet_(resolvePartSheetName_(payload.sheetName || payload.partNumber));
  const headers = getHeaders_(sheet, true);
  const rows = getSheetRows_(sheet, Math.max(sheet.getLastColumn(), BASE_MEASURE_HEADERS.length), false);
  const objectIndex = getHeaderIndex_(headers, 'object');
  const piece = normalizePiece_(payload.piece);
  const stage = String(payload.stage || 'PREVIA_AJUSTE').trim();
  const previousPiece = normalizePiece_(payload.previousPiece || payload.previaPiece || '');
  const adjustmentId = String(payload.adjustmentId || payload.ajusteId || '').trim();
  const values = payload.values || {};
  if (!piece) throw new Error('Captura la PIECE asignada por metrologia.');
  const duplicateCapture = findDuplicateCapturePiece_(headers, rows, piece);
  if (duplicateCapture) {
    throw new Error('La PIECE ' + piece + ' ya existe en ' + sheet.getName() + ' (columna ' + duplicateCapture.columnNumber + '). No se puede sobrescribir el reporte dimensional.');
  }

  const targetColumn = findCaptureColumn_(sheet, headers, rows);
  const captureTypeRowNumber = ensureCaptureTypeRow_(sheet, headers, rows);
  const now = new Date();
  sheet.getRange(1, targetColumn).setValue(piece);

  rows.forEach(row => {
    const objectName = String(row.values[objectIndex] || '').trim();
    const key = normalize_(objectName);
    const cell = sheet.getRange(row.rowNumber, targetColumn);
    if (key === 'piece') {
      cell.setValue(piece);
    } else if (isDateRegKey_(key)) {
      cell.setValue(now);
      cell.setNumberFormat('yyyy-mm-dd hh:mm:ss');
    } else if (Object.prototype.hasOwnProperty.call(values, objectName)) {
      const value = values[objectName];
      cell.setValue(value === undefined || value === null ? '' : value);
    }
  });

  
  const captureType = buildCaptureTypeLabel_(stage, piece, previousPiece);
  sheet.getRange(captureTypeRowNumber, targetColumn).setValue(captureType);
  if (normalize_(stage) === 'postajuste' && previousPiece) {
    linkAdjustmentPostPiece_({ adjustmentId: adjustmentId, previousPiece: previousPiece, postPiece: piece, partNumber: sheet.getName() });
  }

  SpreadsheetApp.flush();
  return { ok: true, sheetName: sheet.getName(), piece: piece, stage: stage, columnNumber: targetColumn, dateReg: formatDateTime_(now), captureType: captureType };
}

function savePressRecord(payload) {
  payload = payload || {};
  const piece = normalizePiece_(payload.piece);
  const part = resolvePartSheetName_(payload.partNumber || payload.np || '');
  if (!piece) throw new Error('Captura PIECE para PRENSA.');
  if (!part) throw new Error('Selecciona numero de parte para PRENSA.');
  const sheet = getRequiredTableSheet_('PRENSA');
  appendByHeaders_(sheet, {
    PIECE: piece,
    NP: part,
    DATE: new Date(),
    ETAPA: payload.stage || payload.etapa || 'ACTUAL',
    AJUSTE_ID: String(payload.adjustmentId || payload.ajusteId || '').trim(),
    PIECE_PREVIA: normalizePiece_(payload.previousPiece || payload.piecePrevia || ''),
    RESPONSABLE: payload.responsible || '',
    OBSERVACIONES: payload.notes || '',

    'Altura golpe delantero izquierdo': payload.alturaGdi,
    'Altura golpe delantero derecho': payload.alturaGdd,
    'Altura golpe trasero izquierdo': payload.alturaGti,
    'Altura golpe trasero derecho': payload.alturaGtd,
    'Altura herramienta cerrada': payload.alturaHc,

    'Si diferencia menor a:': payload.parametroDif1,
    'Factor de correccion menor a:': payload.fcDif1,

    'Si la diferencia entre min:': payload.parametroDif2Min,
    'Si la diferencia entre max:': payload.parametroDif2Max,
    'Factor de correccion entre min y max:': payload.fcDif2,

    'Si diferencia mayor a:': payload.parametroDif3,
    'Factor de correccion mayor a:': payload.fcDif3,

    'Operador_prensa': payload.operadorPrensa
  });
  return { ok: true, sheetName: 'PRENSA', piece: piece, part: part };
}

function saveShimsRecord(payload) {
  payload = payload || {};
  const piece = normalizePiece_(payload.piece);
  const part = resolvePartSheetName_(payload.partNumber || payload.np || '');
  if (!piece) throw new Error('Captura PIECE para AJUSTES.');
  if (!part) throw new Error('Selecciona numero de parte para AJUSTES.');
  const sheet = getRequiredTableSheet_('AJUSTES');
  const record = {
    PIECE: piece,
    NP: part,
    DATE: new Date(),
    ETAPA: payload.stage || payload.etapa || 'ACTUAL',
    AJUSTE_ID: String(payload.adjustmentId || payload.ajusteId || '').trim(),
    PIECE_PREVIA: normalizePiece_(payload.previousPiece || payload.piecePrevia || ''),
    RESPONSABLE: payload.responsible || '',
    OBSERVACIONES: payload.notes || ''
  };
  const values = payload.values || payload.shims || {};
  Object.keys(values).forEach(function(key) {
    const cleanKey = String(key || '').trim();
    if (cleanKey) record[cleanKey] = values[key];
  });
  [['SUP-1','sup1'], ['SUP-2','sup2'], ['INF-1','inf1'], ['INF-2','inf2'], ['INF-3','inf3'], ['INF-4','inf4'], ['INF-5','inf5']].forEach(function(pair) {
    if (payload[pair[1]] !== undefined) record[pair[0]] = payload[pair[1]];
  });
  appendByHeaders_(sheet, record);
  return { ok: true, sheetName: 'AJUSTES', piece: piece, part: part };
}

function getLatestShimsConfig(payload) {
  payload = payload || {};
  const part = resolvePartSheetName_(payload.partNumber || payload.np || '');
  const piece = normalizePiece_(payload.piece);
  if (!part) throw new Error('Selecciona numero de parte para cargar calzas.');
  const sheet = getRequiredTableSheet_('AJUSTES');
  const headers = getHeaders_(sheet, true);
  const dataRows = Math.max(sheet.getLastRow() - 1, 0);
  if (!dataRows) return { ok: true, found: false, part: part, values: {} };
  const rows = sheet.getRange(2, 1, dataRows, headers.length).getDisplayValues();
  const normalizedHeaders = headers.map(normalize_);
  const npIdx = normalizedHeaders.indexOf('np');
  const pieceIdx = normalizedHeaders.indexOf('piece');
  const dateIdx = normalizedHeaders.indexOf('date');
  let selected = null;
  for (let i = rows.length - 1; i >= 0; i--) {
    const rowPart = npIdx >= 0 ? resolvePartText_(rows[i][npIdx]) : '';
    const rowPiece = pieceIdx >= 0 ? normalizePiece_(rows[i][pieceIdx]) : '';
    if (rowPart !== part) continue;
    if (piece && rowPiece === piece) { selected = { row: rows[i], exact: true }; break; }
    if (!selected) selected = { row: rows[i], exact: false };
  }
  if (!selected) return { ok: true, found: false, part: part, values: {} };
  const values = {};
  headers.forEach(function(header, idx) {
    const h = String(header || '').trim();
    if (/^CALZA\s+/i.test(h)) values[h] = selected.row[idx];
  });
  return {
    ok: true,
    found: true,
    part: part,
    piece: pieceIdx >= 0 ? selected.row[pieceIdx] : '',
    date: dateIdx >= 0 ? selected.row[dateIdx] : '',
    exactPiece: selected.exact,
    values: values
  };
}

function resolvePartText_(part) {
  const value = String(part || '').trim().replace(/\.0$/, '');
  if (!value) return '';
  const compact = value.replace(/-00$/, '');
  const ss = getSpreadsheet_();
  if (ss.getSheetByName(value)) return value;
  if (ss.getSheetByName(compact)) return compact;
  if (ss.getSheetByName(compact + '-00')) return compact + '-00';
  return compact;
}

function saveAdjustmentSnapshot(payload) {
  payload = payload || {};
  const previousPiece = normalizePiece_(payload.piece || payload.previousPiece || '');
  const postPiece = normalizePiece_(payload.postPiece || '');
  const part = resolvePartSheetName_(payload.partNumber || '');
  if (!previousPiece) throw new Error('Captura la PIECE previa para vincular el ajuste.');
  if (!part) throw new Error('Selecciona numero de parte para el ajuste.');

  const adjustmentId = generateId_('AJ');
  const recordPiece = postPiece || previousPiece;
  const shared = {
    stage: 'NUEVO_AJUSTE',
    adjustmentId: adjustmentId,
    previousPiece: previousPiece,
    responsible: payload.responsible || '',
    notes: payload.notes || ''
  };

  const press = Object.assign({}, payload.press || {}, shared);
  press.piece = recordPiece;
  press.partNumber = part;
  savePressRecord(press);

  const shims = Object.assign({}, payload.shims || {}, shared);
  shims.piece = recordPiece;
  shims.partNumber = part;
  saveShimsRecord(shims);

  return { ok: true, id: adjustmentId, piece: recordPiece, postPiece: postPiece, previousPiece: previousPiece, part: part, pendingPostPiece: !postPiece };
}

function linkAdjustmentPostPiece(payload) {
  return linkAdjustmentPostPiece_(payload);
}

function linkAdjustmentPostPiece_(payload) {
  payload = payload || {};
  const adjustmentId = String(payload.adjustmentId || payload.ajusteId || '').trim();
  const previousPiece = normalizePiece_(payload.previousPiece || payload.piecePrevia || '');
  const postPiece = normalizePiece_(payload.postPiece || payload.piece || '');
  const part = resolvePartSheetName_(payload.partNumber || payload.np || '');
  if (!previousPiece || !postPiece || !part) return { ok: false, updated: 0 };
  const updatedPress = updateAdjustmentRowsPostPiece_('PRENSA', adjustmentId, previousPiece, postPiece, part);
  const updatedShims = updateAdjustmentRowsPostPiece_('AJUSTES', adjustmentId, previousPiece, postPiece, part);
  return { ok: true, updated: updatedPress + updatedShims, press: updatedPress, shims: updatedShims, previousPiece: previousPiece, postPiece: postPiece, part: part };
}

function updateAdjustmentRowsPostPiece_(sheetName, adjustmentId, previousPiece, postPiece, part) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const headers = ensureTableHeaders_(sheet, ['AJUSTE_ID', 'PIECE_PREVIA']);
  const normalized = headers.map(normalize_);
  const pieceCol = normalized.indexOf('piece') + 1;
  const prevCol = normalized.indexOf('pieceprevia') + 1;
  const adjustCol = normalized.indexOf('ajusteid') + 1;
  const npCol = normalized.indexOf('np') + 1;
  const stageCol = normalized.indexOf('etapa') + 1;
  if (!pieceCol) return 0;
  const rowCount = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
  let updated = 0;
  values.forEach(function(row, idx) {
    const rowPart = npCol ? resolvePartText_(row[npCol - 1]) : '';
    const rowPiece = pieceCol ? String(row[pieceCol - 1] || '').trim().toUpperCase() : '';
    const rowPrev = prevCol ? String(row[prevCol - 1] || '').trim().toUpperCase() : '';
    const rowAdjust = adjustCol ? String(row[adjustCol - 1] || '').trim() : '';
    const rowStage = stageCol ? normalize_(row[stageCol - 1]) : '';
    if (rowPart !== part) return;
    if (rowStage && rowStage !== 'nuevoajuste') return;
    const sameAdjustment = adjustmentId && rowAdjust === adjustmentId;
    const samePair = !adjustmentId && (rowPrev === previousPiece || rowPiece === previousPiece);
    if (!sameAdjustment && !samePair) return;
    const sheetRow = idx + 2;
    sheet.getRange(sheetRow, pieceCol).setValue(postPiece);
    if (prevCol) sheet.getRange(sheetRow, prevCol).setValue(previousPiece);
    updated++;
  });
  if (updated) SpreadsheetApp.flush();
  return updated;
}

function ensureTableHeaders_(sheet, requiredHeaders) {
  let headers = getHeaders_(sheet, true);
  const normalized = headers.map(normalize_);
  const missing = requiredHeaders.filter(function(header) { return normalized.indexOf(normalize_(header)) < 0; });
  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }
  return headers;
}

function extractMeasurementsFromReport(payload) {
  payload = payload || {};
  const part = resolvePartSheetName_(payload.partNumber || payload.sheetName || '');
  const plan = getMeasurementPlan(part);
  const fileName = payload.fileName || 'reporte-dimensional';
  const mimeType = payload.mimeType || 'application/octet-stream';
  const base64 = String(payload.base64 || '').replace(/^data:[^,]+,/, '');
  if (!base64) throw new Error('Selecciona un PDF o imagen para analizar.');

  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
  const text = ocrBlob_(blob, fileName);
  const parsed = parseMeasurementsFromText_(text, plan.objectRows);
  return { ok: true, partNumber: part, textPreview: text.substring(0, 4000), matches: parsed.matches, missing: parsed.missing, total: plan.objectRows.length };
}


function extractPressFromPhoto(payload) {
  payload = payload || {};
  const fileName = payload.fileName || 'parametros-prensa';
  const mimeType = payload.mimeType || 'application/octet-stream';
  const base64 = String(payload.base64 || '').replace(/^data:[^,]+,/, '');
  if (!base64) throw new Error('Selecciona una foto de parametros de prensa.');

  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
  const text = ocrBlob_(blob, fileName, 'es');
  return { ok: true, textPreview: text.substring(0, 4000), values: parsePressParametersFromText_(text) };
}
function ocrBlob_(blob, fileName, language) {
  const resource = { title: 'OCR_' + fileName };
  let doc = null;
  let lastError = null;
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      doc = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: language || 'en' });
      const text = DocumentApp.openById(doc.id).getBody().getText();
      try { Drive.Files.remove(doc.id); } catch(e) {}
      return text;
    } catch (e) {
      lastError = e;
      const message = String(e && e.message || e);
      if (doc && doc.id) { try { Drive.Files.remove(doc.id); } catch(removeError) {} }
      if (!/rate limit|user rate limit|exceeded/i.test(message) || attempt === 2) break;
      Utilities.sleep((attempt + 1) * 2500);
    }
  }
  const message = String(lastError && lastError.message || lastError || '');
  if (/rate limit|user rate limit|exceeded/i.test(message)) {
    throw new Error('Limite temporal de OCR alcanzado por Google Drive. Espera 30 a 60 segundos y vuelve a intentar; no es error del archivo.');
  }
  throw lastError;
}

function normalizeOcrLine_(line) {
  return String(line || '')
    .replace(/[|]/g, ' ')
    .replace(/\bS(?:urt|ur|utf|urf)\s*D(?:lst|ist|1st)\b/ig, 'Surf Dist')
    .replace(/\bD(?:ev|e[vw])\b/ig, 'Dev')
    .replace(/\b(?:Norn|Nora|Nomn)\b/ig, 'Nom')
    .replace(/\b(?:Mcas|Meas\.|Mea5)\b/ig, 'Meas')
    .replace(/\s+/g, ' ')
    .trim();
}

function surfDistValueRegex_() {
  return /\bSurf\s*Dist\b\s*([-+]?\d+(?:[.,]\d+)?)/ig;
}
function parseMeasurementsFromText_(text, objectRows) {
  const lines = String(text || '').split(/\r?\n/).map(normalizeOcrLine_).filter(Boolean);
  const matches = {};
  const missing = [];
  const surfMap = parseSurfDistMapFromLines_(lines);
  objectRows.forEach(row => {
    const objectName = String(row.object || '').trim();
    const lineIndex = findObjectLineIndex_(lines, objectName);
    if (lineIndex < 0) {
      missing.push(objectName);
      return;
    }
    const windowText = lines.slice(lineIndex, Math.min(lines.length, lineIndex + 12)).join(' ');
    let value = extractMeasurementValue_(windowText, row);
    const surfKey = surfPointKey_(objectName);
    if (surfKey && Object.prototype.hasOwnProperty.call(surfMap, surfKey)) {
      const blockValue = surfMap[surfKey];
      if (shouldUseBlockSurfValue_(value, blockValue, windowText, row)) value = blockValue;
    }
    if (value === null || value === undefined || value === '') {
      missing.push(objectName);
      return;
    }
    matches[objectName] = value;
  });
  return { matches: matches, missing: missing };
}

function findObjectLineIndex_(lines, objectName) {
  const pattern = objectNameRegex_(objectName);
  for (let i = 0; i < lines.length; i++) {
    if (pattern.test(lines[i])) return i;
  }
  for (let i = 0; i < lines.length - 1; i++) {
    const windowText = lines.slice(i, i + 2).join(' ');
    if (pattern.test(windowText)) return i;
  }
  return -1;
}

function objectNameRegex_(objectName) {
  const name = String(objectName || '').trim();
  const surf = name.match(/^surf\s*pt\s*(\d+)$/i);
  if (surf) return new RegExp('(^|[^a-z0-9])surf\\s*pt\\s*' + surf[1] + '(?!\\d)(?=$|[^a-z0-9])', 'i');
  const angle = name.match(/^angle\s*(\d+)$/i);
  if (angle) return new RegExp('(^|[^a-z0-9])angle\\s*' + angle[1] + '(?!\\d)(?=$|[^a-z0-9])', 'i');
  return new RegExp('(^|[^a-z0-9])' + name.split(/\s+/).map(escapeRegex_).join('\\s*') + '(?=$|[^a-z0-9])', 'i');
}

function extractMeasurementValue_(text, row) {
  const objectName = String(row.object || '');
  const isAngle = normalize_(objectName).indexOf('angle') === 0;
  const compact = String(text || '').replace(/\s+/g, ' ');

  if (isAngle) return extractAngleMeas_(compact, row);
  return extractSurfDev_(compact, row);
}

function extractAngleMeas_(text, row) {
  const nom = cleanNumber_(row.nominal);
  const numbers = extractNumbers_(text).filter(n => Math.abs(n) > 10);
  if (!numbers.length) return null;
  if (nom !== null) {
    for (let i = 0; i < numbers.length - 1; i++) {
      if (Math.abs(numbers[i] - nom) < 0.01) return numbers[i + 1];
    }
    const candidates = numbers.filter(n => Math.abs(n - nom) > 0.01 && Math.abs(n - nom) < 5);
    if (candidates.length) return candidates[0];
  }
  const measBlock = text.match(/\bNom\b[\s\S]*?\bMeas\b[\s\S]*?\bDev\b[\s\S]*?([-+]?\d+(?:[.,]\d+)?)\s+([-+]?\d+(?:[.,]\d+)?)/i);
  if (measBlock) return cleanNumber_(measBlock[2]);
  return numbers.length > 1 ? numbers[1] : numbers[0];
}

function extractSurfDev_(text, row) {
  const scoped = scopeTextAfterFirstSurfLabel_(String(text || ''));
  const objectName = String(row && row.object || '');
  const values = [];
  const valueRe = surfDistValueRegex_();
  let m;
  while ((m = valueRe.exec(scoped)) !== null) values.push(cleanNumber_(m[1]));
  if (/^surf\s*pt\s*3$/i.test(objectName) && values.length >= 3) return values[2];
  if (values.length) return values[0];
  const dev = scoped.match(/\bDev\b\s*([-+]?\d+(?:[.,]\d+)?)/i);
  if (dev) return cleanNumber_(dev[1]);
  return null;
}

function scopeTextAfterFirstSurfLabel_(text) {
  const match = String(text || '').match(/(^|[^a-z0-9])surf\s*pt\s*\d+(?!\d)/i);
  return match ? text.substring(match.index + match[0].length) : text;
}

function parseSurfDistMapFromLines_(lines) {
  const tokens = [];
  lines.forEach((line, lineIndex) => {
    const text = String(line || '');
    let m;
    const labelRe = /(^|[^a-z0-9])surf\s*pt\s*(\d+)(?!\d)/ig;
    while ((m = labelRe.exec(text)) !== null) tokens.push({ type: 'label', key: 'surfpt' + m[2], line: lineIndex, pos: m.index });
    const valueRe = surfDistValueRegex_();
    while ((m = valueRe.exec(text)) !== null) tokens.push({ type: 'value', value: cleanNumber_(m[1]), line: lineIndex, pos: m.index });
  });
  tokens.sort((a, b) => a.line - b.line || a.pos - b.pos || (a.type === 'label' ? -1 : 1));

  const result = {};
  let pending = [];
  let values = [];
  tokens.forEach(token => {
    if (token.type === 'label') pending.push(token.key);
    if (token.type === 'value' && token.value !== null) values.push(token.value);
    while (pending.length && values.length) {
      const key = pending.shift();
      const value = values.shift();
      if (!Object.prototype.hasOwnProperty.call(result, key)) result[key] = value;
    }
  });
  return result;
}

function shouldUseBlockSurfValue_(localValue, blockValue, windowText, row) {
  if (blockValue === null || blockValue === undefined || blockValue === '') return false;
  if (localValue === null || localValue === undefined || localValue === '') return true;
  if (Math.abs(Number(localValue) - Number(blockValue)) < 0.0001) return false;
  const objectName = String(row.object || '').toLowerCase();
  return /^surf\s*pt\s*6$/i.test(objectName);
}

function surfPointKey_(objectName) {
  const m = String(objectName || '').match(/^surf\s*pt\s*(\d+)$/i);
  return m ? 'surfpt' + m[1] : '';
}
function parsePressParametersFromText_(text) {
  const normalized = normalizePressOcrText_(text);
  const numbers = extractNumbers_(normalized);
  const values = {};

  values.alturaGdi = findNumberAfterLabel_(normalized, /Altura\s+golpe\s+delanter[oa]\s+izquierd[oa]/i);
  values.alturaGdd = findNumberAfterLabel_(normalized, /Altura\s+golpe\s+delanter[oa]\s+derech[oa]/i);
  values.alturaGti = findNumberAfterLabel_(normalized, /Altura\s+golpe\s+traser[oa]\s+izquierd[oa]/i);
  values.alturaGtd = findNumberAfterLabel_(normalized, /Altura\s+golpe\s+traser[oa]\s+derech[oa]/i);
  values.alturaHc = findNumberAfterLabel_(normalized, /Altura\s+herramienta\s+cerrada/i);

  applyPressCorrectionTableFallback_(normalized, values);

  if (!values.alturaGdi && numbers.length >= 5 && numbers.some(n => n > 100)) {
    values.alturaGdi = numbers[0];
    values.alturaGti = numbers[1];
    values.alturaHc = numbers[2];
    values.alturaGdd = numbers[3];
    values.alturaGtd = numbers[4];
  }
  return values;
}

function normalizePressOcrText_(text) {
  return String(text || '')
    .replace(/,/g, '.')
    .replace(/\bmm\b/ig, ' ')
    .replace(/[|]/g, ' ')
    .replace(/correccion|correcion/ig, 'correccion')
    .replace(/diferenc[a-z]*/ig, 'diferencia')
    .replace(/\s+/g, ' ')
    .trim();
}

function applyPressCorrectionTableFallback_(text, values) {
  const tableStart = firstExistingIndex_(text, [/Tabla\s+de\s+correccion/i, /Si\s+diferencia\s+menor/i]);
  const tableText = tableStart >= 0 ? text.substring(tableStart) : text;

  const menor = rowNumbers_(tableText, /Si\s+diferencia\s+menor\s+a/i, /Si\s+la\s+diferencia\s+entre|Si\s+diferencia\s+mayor\s+a/i);
  const entre = rowNumbers_(tableText, /Si\s+la\s+diferencia\s+entre/i, /Si\s+diferencia\s+mayor\s+a/i);
  const mayor = rowNumbers_(tableText, /Si\s+diferencia\s+mayor\s+a/i, /(?:Cerrar|$)/i);

  if (menor.length >= 1) values.parametroDif1 = menor[0];
  if (menor.length >= 2) values.fcDif1 = menor[menor.length - 1];
  if (values.parametroDif1 === 7) values.parametroDif1 = 1.5;
  if (values.fcDif1 === 1.5 && values.parametroDif1 === 1.5) values.fcDif1 = 1.0;

  if (entre.length >= 1) values.parametroDif2Min = entre[0];
  if (entre.length >= 2) values.parametroDif2Max = entre[1];
  if (entre.length >= 3) values.fcDif2 = entre[entre.length - 1];

  if (mayor.length >= 1) values.parametroDif3 = mayor[0];
  if (mayor.length >= 2) values.fcDif3 = mayor[mayor.length - 1];

  fillStandardPressCorrectionDefaults_(values);
}

function rowNumbers_(text, startRegex, endRegex) {
  const row = extractBetween_(text, startRegex, endRegex);
  return extractNumbers_(row).filter(isPressCorrectionNumber_);
}

function fillStandardPressCorrectionDefaults_(values) {
  if ((values.parametroDif1 === '' || values.parametroDif1 == null) && values.parametroDif2Min === 1.5) values.parametroDif1 = 1.5;
  if ((values.fcDif1 === '' || values.fcDif1 == null) && values.parametroDif1 === 1.5) values.fcDif1 = 1.0;
  if ((values.fcDif2 === '' || values.fcDif2 == null) && values.parametroDif2Min === 1.5 && values.parametroDif2Max === 2.5) values.fcDif2 = 1.1;
  if ((values.fcDif3 === '' || values.fcDif3 == null) && values.parametroDif3 === 2.5) values.fcDif3 = 1.2;
}

function firstExistingIndex_(text, patterns) {
  const indexes = patterns.map(p => String(text || '').search(p)).filter(i => i >= 0);
  return indexes.length ? Math.min.apply(null, indexes) : -1;
}

function extractBetween_(text, startRegex, endRegex) {
  const source = String(text || '');
  const start = source.search(startRegex);
  if (start < 0) return '';
  const rest = source.substring(start);
  const end = rest.substring(1).search(endRegex);
  return end >= 0 ? rest.substring(0, end + 1) : rest;
}

function isPressCorrectionNumber_(n) {
  return n > 0 && n <= 3;
}
function findNumberAfterLabel_(text, labelRegex) {
  const m = String(text || '').match(new RegExp(labelRegex.source + '[^\\d+-]{0,140}([-+]?\\d+(?:[.,]\\d+)?)', 'i'));
  return m ? cleanNumber_(m[1]) : '';
}

function findSecondNumberAfterLabel_(text, labelRegex) {
  const m = String(text || '').match(new RegExp(labelRegex.source + '(.{0,120})', 'i'));
  if (!m) return '';
  const nums = extractNumbers_(m[1]);
  return nums.length > 1 ? nums[1] : '';
}

function extractNumbers_(text) {
  const numbers = String(text || '').match(/[-+]?\d+(?:[.,]\d+)?/g) || [];
  return numbers.map(cleanNumber_).filter(n => n !== null);
}

function cleanNumber_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = Number(String(value).replace(',', '.'));
  return isFinite(n) ? n : null;
}
function getSheetData(sheetName) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet, true);
  const rows = getSheetRows_(sheet, headers.length, true);
  if (isMeasurementSheet_(headers)) return { mode: 'measurement', headers: headers, rows: rows, objectRows: getMeasurementObjectRows_(headers, rows), captureColumns: getCaptureColumns_(headers, rows) };
  return { mode: 'table', headers: headers, rows: rows, objectRows: [], captureColumns: [] };
}

function getMeasurementObjectRows_(headers, rows) {
  const unitsIndex = getHeaderIndex_(headers, 'units');
  const objectIndex = getHeaderIndex_(headers, 'object');
  const controlIndex = getHeaderIndex_(headers, 'control');
  const nominalIndex = getHeaderIndex_(headers, 'nominal');
  const toleranceIndex = getHeaderIndex_(headers, 'tolerance');
  return rows.map(row => {
    const objectName = String(row.values[objectIndex] || '').trim();
    const key = normalize_(objectName);
    return { rowNumber: row.rowNumber, units: row.values[unitsIndex] || '', object: objectName, control: row.values[controlIndex] || '', nominal: row.values[nominalIndex] || '', tolerance: row.values[toleranceIndex] || '', isMeta: key === 'piece' || isDateRegKey_(key) || isCaptureTypeKey_(key) };
  }).filter(row => row.object && !row.isMeta);
}

function isCaptureTypeKey_(key) { return key === 'tipocaptura' || key === 'tipoidentificador' || key === 'identificadorcaptura'; }

function buildCaptureTypeLabel_(stage, piece, previousPiece) {
  const key = normalize_(stage);
  if (key === 'postajuste') {
    return 'PIEZA POST AJUSTE DE ' + (previousPiece || 'PIECE PREVIA') + ' PREVIA AJUSTE';
  }
  return 'PIEZA PREVIA AL AJUSTE';
}

function ensureCaptureTypeRow_(sheet, headers, rows) {
  const objectIndex = getHeaderIndex_(headers, 'object');
  const unitsIndex = getHeaderIndex_(headers, 'units');
  const controlIndex = getHeaderIndex_(headers, 'control');
  const nominalIndex = getHeaderIndex_(headers, 'nominal');
  const toleranceIndex = getHeaderIndex_(headers, 'tolerance');
  const existing = rows.find(function(row) { return isCaptureTypeKey_(normalize_(row.values[objectIndex])); });
  if (existing) return existing.rowNumber;
  const rowNumber = sheet.getLastRow() + 1;
  const base = new Array(Math.max(sheet.getLastColumn(), BASE_MEASURE_HEADERS.length)).fill('');
  if (unitsIndex >= 0) base[unitsIndex] = '';
  if (objectIndex >= 0) base[objectIndex] = 'Tipo captura';
  if (controlIndex >= 0) base[controlIndex] = 'Identificador';
  if (nominalIndex >= 0) base[nominalIndex] = '';
  if (toleranceIndex >= 0) base[toleranceIndex] = '';
  sheet.getRange(rowNumber, 1, 1, base.length).setValues([base]);
  return rowNumber;
}

function getCaptureColumns_(headers, rows) {
  const objectIndex = getHeaderIndex_(headers, 'object');
  const pieceRow = rows.find(row => normalize_(row.values[objectIndex]) === 'piece');
  const dateRow = rows.find(row => isDateRegKey_(normalize_(row.values[objectIndex])));
  const captureColumns = [];
  const width = Math.max(headers.length, BASE_MEASURE_HEADERS.length);
  for (let columnIndex = BASE_MEASURE_HEADERS.length; columnIndex < width; columnIndex++) {
    const header = headers[columnIndex] || '';
    const piece = pieceRow ? pieceRow.values[columnIndex] : header;
    const dateReg = dateRow ? dateRow.values[columnIndex] : '';
    const displayPiece = normalizePiece_(piece || header);
    if (header || piece || dateReg) captureColumns.push({ index: columnIndex + 1, header: header, piece: displayPiece || piece || header, dateReg: dateReg || '' });
  }
  return captureColumns;
}

function findDuplicateCapturePiece_(headers, rows, piece) {
  const targetKey = canonicalPieceKey_(piece);
  if (!targetKey) return null;
  const objectIndex = getHeaderIndex_(headers, 'object');
  const pieceRow = rows.find(row => normalize_(row.values[objectIndex]) === 'piece');
  const width = Math.max(headers.length, BASE_MEASURE_HEADERS.length);
  for (let columnIndex = BASE_MEASURE_HEADERS.length; columnIndex < width; columnIndex++) {
    const header = headers[columnIndex] || '';
    const rowPiece = pieceRow ? pieceRow.values[columnIndex] : '';
    const candidate = rowPiece || header;
    if (!candidate) continue;
    if (canonicalPieceKey_(candidate) === targetKey) {
      return { columnNumber: columnIndex + 1, piece: String(candidate || '').trim() };
    }
  }
  return null;
}

function findCaptureColumn_(sheet, headers, rows) {
  const objectIndex = getHeaderIndex_(headers, 'object');
  const dateRow = rows.find(row => isDateRegKey_(normalize_(row.values[objectIndex])));
  const width = Math.max(sheet.getLastColumn(), BASE_MEASURE_HEADERS.length);
  for (let column = BASE_MEASURE_HEADERS.length + 1; column <= width; column++) {
    const headerValue = sheet.getRange(1, column).getDisplayValue();
    const dateValue = dateRow ? sheet.getRange(dateRow.rowNumber, column).getDisplayValue() : '';
    if (!headerValue && !dateValue) return column;
  }
  return width + 1;
}
function getFullReport(payload) {
  payload = payload || {};
  const ss = getSpreadsheet_();
  const selectedPart = String(payload.partNumber || '').trim();
  const startDate = parseFilterDate_(payload.startDate, false);
  const endDate = parseFilterDate_(payload.endDate, true);
  const pressRows = readTableObjects_('PRENSA');
  const shimsRows = readTableObjects_('AJUSTES');
  const captures = [];
  const detailRows = [];
  let measureHeaders = [];
  const measureUnits = {};

  getAppConfig().parts.forEach(partInfo => {
    const part = partInfo.name;
    if (selectedPart && resolvePartSheetName_(selectedPart) !== part) return;
    const plan = getMeasurementPlan(part);
    const objectRows = plan.objectRows;
    measureHeaders = mergeUnique_(measureHeaders, objectRows.map(r => r.object));
    objectRows.forEach(r => { if (r.object && !measureUnits[r.object]) measureUnits[r.object] = unitAbbrev_(r.units); });
    const rowsByObject = {};
    objectRows.forEach(r => rowsByObject[normalize_(r.object)] = r);

    plan.captureColumns.forEach(col => {
      const piece = normalizePiece_(col.piece || col.header);
      if (!piece) return;
      const dateFromFaro = parseMaybeDate_(col.dateReg);
      const relatedPress = pressRows.filter(r => normalizePiece_(r.PIECE) === piece && partKey_(r.NP) === partKey_(part));
      const relatedShims = shimsRows.filter(r => normalizePiece_(r.PIECE) === piece && partKey_(r.NP) === partKey_(part));
      const reportDate = dateFromFaro || latestDateFromRows_(relatedPress) || latestDateFromRows_(relatedShims);
      if (!dateInRange_(reportDate, startDate, endDate)) return;

      const values = {};
      const statuses = {};
      let filled = 0;
      let nok = 0;
      objectRows.forEach(obj => {
        const row = plan.rows.find(r => normalize_(r.values[getHeaderIndex_(plan.headers, 'object')]) === normalize_(obj.object));
        const value = row ? row.values[col.index - 1] : '';
        values[obj.object] = value;
        statuses[obj.object] = value === '' || value === null || value === undefined ? '' : (isNok_(value, obj.tolerance, obj.nominal) ? 'NOK' : 'OK');
        if (value !== '' && value !== null && value !== undefined) filled++;
        if (isNok_(value, obj.tolerance, obj.nominal)) nok++;
        detailRows.push({
          dateReg: reportDate ? formatDateTime_(reportDate) : '',
          part: part,
          piece: piece,
          object: obj.object,
          control: obj.control,
          nominal: obj.nominal,
          tolerance: obj.tolerance,
          units: unitAbbrev_(obj.units),
          value: value,
          status: statuses[obj.object]
        });
      });

      const press = relatedPress.length ? relatedPress[relatedPress.length - 1] : {};
      const shims = relatedShims.length ? relatedShims[relatedShims.length - 1] : {};
      captures.push({
        dateReg: reportDate ? formatDateTime_(reportDate) : '',
        part: part,
        piece: piece,
        values: values,
        statuses: statuses,
        filledMeasurements: filled,
        totalMeasurements: objectRows.length,
        nokMeasurements: nok,
        pressCount: relatedPress.length,
        shimsCount: relatedShims.length,
        press: normalizePressForReport_(press),
        shims: normalizeShimsForReport_(shims)
      });
    });
  });

  captures.sort((a, b) => String(a.dateReg).localeCompare(String(b.dateReg)) || String(a.piece).localeCompare(String(b.piece)));
  return { ok: true, captures: captures, detailRows: detailRows, measureHeaders: measureHeaders, measureUnits: measureUnits, pressHeaders: getReportPressHeaders_(), shimsHeaders: getReportShimsHeaders_(shimsRows) };
}

function unitAbbrev_(unit) {
  const text = String(unit || '').trim();
  if (!text) return '';
  if (/millimeters?|milimetros?/i.test(text)) return 'mm';
  if (/degrees?|grados?/i.test(text)) return 'deg';
  return text;
}

function normalizePressForReport_(press) {
  press = press || {};
  return {
    piece2: normalizePiece_(press.PIECE),
    previousPiece: normalizePiece_(press.PIECE_PREVIA),
    dateReg: press.DATE || '',
    alturaGdi: press['Altura golpe delantero izquierdo'] || '',
    alturaGdd: press['Altura golpe delantero derecho'] || '',
    alturaGti: press['Altura golpe trasero izquierdo'] || '',
    alturaGtd: press['Altura golpe trasero derecho'] || '',
    alturaHc: press['Altura herramienta cerrada'] || '',
    parametroDif1: press['Si diferencia menor a:'] || '',
    fcDif1: press['Factor de correccion menor a'] || press['Factor de correccion menor a:'] || '',
    parametroDif2: joinBetween_(press['Si la diferencia entre min:'], press['Si la diferencia entre max:']),
    fcDif2: press['Factor de correccion entre min y max:'] || press['Factor de correccion entre min y max:'] || '',
    parametroDif3: press['Si diferencia mayor a:'] || '',
    fcDif3: press['Factor de correccion mayor a:'] || ''
  };
}

function getReportPressHeaders_() {
  return [
    ['piece2', 'ID'],
    ['previousPiece', 'PIECE previa'],
    ['dateReg', 'Date Reg'],
    ['alturaGdi', 'Altura golpe delantero izquierdo'],
    ['alturaGdd', 'Altura golpe delantero derecho'],
    ['alturaGti', 'Altura golpe trasero izquierdo'],
    ['alturaGtd', 'Altura golpe trasero derecho'],
    ['alturaHc', 'Altura herramienta cerrada'],
    ['parametroDif1', 'Si diferencia menor a:'],
    ['fcDif1', 'Factor de correccion menor a:'],
    ['parametroDif2', 'Si la diferencia entre:'],
    ['fcDif2', 'Factor de correccion entre:'],
    ['parametroDif3', 'Si diferencia mayor a:'],
    ['fcDif3', 'Factor de correccion mayor a:']
  ];
}

function normalizeShimsForReport_(row) {
  row = row || {};
  const out = { piece: normalizePiece_(row.PIECE), previousPiece: normalizePiece_(row.PIECE_PREVIA), dateReg: row.DATE || '' };
  Object.keys(row).forEach(function(key) {
    if (/^CALZA\s+/i.test(String(key || ''))) out[key] = row[key];
  });
  return out;
}

function getReportShimsHeaders_(rows) {
  const seen = {};
  rows.forEach(function(row) {
    Object.keys(row || {}).forEach(function(key) {
      if (/^CALZA\s+/i.test(String(key || ''))) seen[key] = true;
    });
  });
  return Object.keys(seen).sort();
}

function readTableObjects_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const headers = getHeaders_(sheet, true);
  return getSheetRows_(sheet, headers.length, true).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row.values[i]);
    return obj;
  });
}

function isNok_(value, tolerance, nominal) {
  const n = cleanNumber_(value);
  if (n === null) return false;
  const tol = parseTolerance_(tolerance);
  if (!tol) return false;
  const nom = cleanNumber_(nominal);
  if (nom !== null) return n < nom - tol || n > nom + tol;
  return Math.abs(n) > tol;
}

function parseTolerance_(tolerance) {
  const nums = extractNumbers_(tolerance);
  return nums.length ? Math.abs(nums[0]) : null;
}

function parseFilterDate_(value, endOfDay) {
  if (!value) return null;
  const text = String(value || '').trim();
  const m = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date(value);
  if (isNaN(d.getTime())) return null;
  if (endOfDay) d.setHours(23, 59, 59, 999);
  else d.setHours(0, 0, 0, 0);
  return d;
}

function parseMaybeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (!value) return null;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function latestDateFromRows_(rows) {
  let latest = null;
  rows.forEach(row => {
    const d = parseMaybeDate_(row.DATE);
    if (d && (!latest || d > latest)) latest = d;
  });
  return latest;
}

function dateInRange_(date, startDate, endDate) {
  if (!startDate && !endDate) return true;
  if (!date) return false;
  if (startDate && date < startDate) return false;
  if (endDate && date > endDate) return false;
  return true;
}

function canonicalPieceKey_(value) {
  const text = String(value || '').trim().toUpperCase();
  if (!text) return '';
  let normalized = text
    .replace(/^PIECE\s*/i, '')
    .replace(/^PIEZA\s*/i, '')
    .replace(/[\s_-]+/g, '');
  const match = normalized.match(/^P?0*(\d+)$/);
  if (match) return 'P' + String(Number(match[1])).padStart(3, '0');
  return normalized;
}

function normalizePiece_(value) {
  return canonicalPieceKey_(value);
}

function partKey_(value) {
  return String(value || '').trim().replace(/\.0$/, '').replace(/-00$/, '');
}

function joinBetween_(min, max) {
  if (min && max) return min + ' y ' + max;
  return min || max || '';
}

function mergeUnique_(base, add) {
  const seen = {};
  base.concat(add).forEach(v => { if (v) seen[v] = true; });
  return Object.keys(seen);
}

function escapeRegex_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


function getSpreadsheet_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function getSheetRows_(sheet, width, display) {
  if (sheet.getLastRow() < 2) return [];
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(width, 1));
  const values = display ? range.getDisplayValues() : range.getValues();
  return values.map((row, index) => ({ rowNumber: index + 2, values: row }));
}

function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('No existe la hoja: ' + sheetName);
  return sheet;
}

function getRequiredTableSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error('No existe la hoja requerida: ' + name);
  if (!getHeaders_(sheet, true).length) throw new Error('La hoja ' + name + ' no tiene encabezados.');
  return sheet;
}

function getHeaders_(sheet, includeBlankTail) {
  const lastColumn = Math.max(sheet.getLastColumn(), BASE_MEASURE_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(v => String(v || '').trim());
  if (!includeBlankTail) return headers.filter(Boolean);
  while (headers.length && !headers[headers.length - 1]) headers.pop();
  if (!headers.length) throw new Error('La hoja ' + sheet.getName() + ' no tiene encabezados en la fila 1.');
  return headers;
}

function appendByHeaders_(sheet, values) {
  let headers = getHeaders_(sheet, true);
  const normalizedHeaders = headers.map(normalize_);
  const missing = Object.keys(values).filter(function(key) {
    return normalizedHeaders.indexOf(normalize_(key)) < 0;
  });
  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }
  const normalizedValues = {};
  Object.keys(values).forEach(k => normalizedValues[normalize_(k)] = values[k]);
  const row = headers.map(h => {
    const value = Object.prototype.hasOwnProperty.call(values, h) ? values[h] : normalizedValues[normalize_(h)];
    if (value instanceof Date) return value;
    return value === undefined || value === null ? '' : value;
  });
  sheet.appendRow(row);
  const dateCol = headers.map(normalize_).indexOf('date') + 1;
  if (dateCol > 0) sheet.getRange(sheet.getLastRow(), dateCol).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  SpreadsheetApp.flush();
}

function resolvePartSheetName_(part) {
  const value = String(part || '').trim().replace(/\.0$/, '');
  if (!value) return '';
  const ss = getSpreadsheet_();
  if (ss.getSheetByName(value)) return value;
  const compact = value.replace(/-00$/, '');
  if (ss.getSheetByName(compact)) return compact;
  const dash = compact + '-00';
  if (ss.getSheetByName(dash)) return dash;
  throw new Error('No existe la hoja del numero de parte: ' + value);
}

function isMeasurementSheet_(headers) {
  const normalized = headers.map(normalize_);
  return BASE_MEASURE_HEADERS.every(h => normalized.indexOf(h) >= 0);
}

function getHeaderIndex_(headers, normalizedName) { return headers.map(normalize_).indexOf(normalizedName); }
function normalize_(value) { return String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, ''); }
function isDateRegKey_(key) { return key === 'datereg' || key === 'fechareg' || key === 'date'; }
function formatDateTime_(date) { return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
function generateId_(prefix) { return prefix + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + String(Math.floor(Math.random() * 1000)).padStart(3, '0'); }