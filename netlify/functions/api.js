const { google } = require('googleapis');

const CORS_HEADERS = {
  'Content-Type': 'application/json; charset=utf-8',
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'GET,POST,OPTIONS'
};

const CONFIG_HEADERS = ['CODIGO','NOMBRE','SHEET_ID','PIN_ADMIN','ACTIVA','FECHA','FOLDER_ID','FOLDER_URL'];
const SHEET_CONFIG_RIFAS = 'CONFIG_RIFAS';
const SHEET_LISTAS = 'LISTAS';
const SHEET_VENTAS = 'VENTAS';
const SHEET_WINNERS = 'SORTEO_GANADORES';
const VENTAS_HEADERS = ['Numero','NombreApellido','DNI','Telefono','Vendedor','Fecha','PrecioCobrado'];
const WINNERS_HEADERS = ['Numero','NombreApellido','DNI','Telefono','Vendedor','FechaSorteo'];

function response(statusCode, data) {
  return { statusCode, headers: CORS_HEADERS, body: JSON.stringify(data) };
}
function ok(data) { return { ok: true, data: data ?? null }; }
function fail(err) { return { ok: false, error: err && err.message ? err.message : String(err) }; }
function normalizeCode(code) { return String(code || '').trim().toLowerCase().replace(/\s+/g, ''); }
function safeName(name) { return String(name || '').trim().replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').substring(0,120) || 'sin_nombre'; }
function normHeader(s) { return String(s || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,''); }
function nowIso() { return new Date().toISOString(); }
function parseBool(v) { return String(v || '').trim().toUpperCase() === 'SI' || v === true; }

function envRequired(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Falta variable de entorno ${name}.`);
  return value;
}

async function getClients() {
  const email = envRequired('GOOGLE_SERVICE_ACCOUNT_EMAIL');
  let key = envRequired('GOOGLE_PRIVATE_KEY');
  key = key.replace(/\\n/g, '\n');
  const auth = new google.auth.JWT({
    email,
    key,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive'
    ]
  });
  const sheets = google.sheets({ version: 'v4', auth });
  const drive = google.drive({ version: 'v3', auth });
  return { sheets, drive };
}

async function getValues(sheets, spreadsheetId, range) {
  try {
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    return res.data.values || [];
  } catch (e) {
    if (e.code === 400) return [];
    throw e;
  }
}

async function updateValues(sheets, spreadsheetId, range, values) {
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values }
  });
}

async function appendValues(sheets, spreadsheetId, range, values) {
  const res = await sheets.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values }
  });
  return res.data;
}

async function getSpreadsheetMeta(sheets, spreadsheetId) {
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  return res.data.sheets || [];
}

async function getSheetIdByName(sheets, spreadsheetId, title) {
  const shs = await getSpreadsheetMeta(sheets, spreadsheetId);
  const found = shs.find(s => s.properties && s.properties.title === title);
  return found ? found.properties.sheetId : null;
}

async function ensureSheet(sheets, spreadsheetId, title) {
  const existingId = await getSheetIdByName(sheets, spreadsheetId, title);
  if (existingId !== null) return existingId;
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: { title } } }] }
  });
  return res.data.replies[0].addSheet.properties.sheetId;
}

async function ensureHeaders(sheets, spreadsheetId) {
  await ensureSheet(sheets, spreadsheetId, SHEET_LISTAS);
  await ensureSheet(sheets, spreadsheetId, SHEET_VENTAS);
  await ensureSheet(sheets, spreadsheetId, SHEET_WINNERS);

  const listas = await getValues(sheets, spreadsheetId, `${SHEET_LISTAS}!A1:H8`);
  if (!listas[0] || !listas[0][0]) {
    await updateValues(sheets, spreadsheetId, `${SHEET_LISTAS}!A1:H8`, [
      ['Vendedores','Ventas','','Total $','Precio unitario','', 'Publicidad',''],
      ['', '', '', '1 número', 0, 'Desde', 1, ''],
      ['', '', '', '2 números', 0, 'Hasta', 500, ''],
      ['', '', '', 'Promo 2x activa', 'NO', 'Contacto', '', ''],
      ['', '', '', '', '', '', 'Imagen fondo ID', ''],
      ['', '', '', '', '', '', 'Texto extra', ''],
      ['', '', '', '', '', '', 'Imagen fondo nombre', ''],
      ['', '', '', '', '', '', 'Mensaje inferior general', '']
    ]);
  }
  const ventas = await getValues(sheets, spreadsheetId, `${SHEET_VENTAS}!A1:G1`);
  if (!ventas[0] || !ventas[0][0]) await updateValues(sheets, spreadsheetId, `${SHEET_VENTAS}!A1:G1`, [VENTAS_HEADERS]);
  const winners = await getValues(sheets, spreadsheetId, `${SHEET_WINNERS}!A1:F1`);
  if (!winners[0] || !winners[0][0]) await updateValues(sheets, spreadsheetId, `${SHEET_WINNERS}!A1:F1`, [WINNERS_HEADERS]);
}

async function ensureMasterHeaders(sheets) {
  const masterId = envRequired('MASTER_SS_ID');
  await ensureSheet(sheets, masterId, SHEET_CONFIG_RIFAS);
  const values = await getValues(sheets, masterId, `${SHEET_CONFIG_RIFAS}!A1:H1`);
  const row = values[0] || [];
  const okHeaders = CONFIG_HEADERS.every((h,i) => String(row[i]||'').trim().toUpperCase() === h);
  if (!okHeaders) await updateValues(sheets, masterId, `${SHEET_CONFIG_RIFAS}!A1:H1`, [CONFIG_HEADERS]);
}

async function listRifasRaw(sheets) {
  const masterId = envRequired('MASTER_SS_ID');
  await ensureMasterHeaders(sheets);
  const values = await getValues(sheets, masterId, `${SHEET_CONFIG_RIFAS}!A1:H`);
  const rows = values.slice(1).filter(r => String(r[0]||'').trim());
  return rows.map((r, idx) => ({
    rowNumber: idx + 2,
    codigo: String(r[0] || ''),
    nombre: String(r[1] || ''),
    sheetId: String(r[2] || ''),
    pinAdmin: String(r[3] || ''),
    activa: String(r[4] || ''),
    fecha: String(r[5] || ''),
    folderId: String(r[6] || ''),
    folderUrl: String(r[7] || '')
  }));
}

async function findRifa(sheets, codigo) {
  const code = normalizeCode(codigo);
  if (!code) throw new Error('Falta el código de rifa.');
  const rifas = await listRifasRaw(sheets);
  const rifa = rifas.find(r => normalizeCode(r.codigo) === code);
  if (!rifa) throw new Error('Código de rifa inválido.');
  if (String(rifa.activa || '').trim().toUpperCase() !== 'SI') throw new Error('Esta rifa no está activa.');
  return rifa;
}

async function assertAdmin(sheets, codigo, pin) {
  const rifa = await findRifa(sheets, codigo);
  if (String(pin || '') !== String(rifa.pinAdmin || '')) throw new Error('PIN incorrecto.');
  return rifa;
}

async function getListasConfig(sheets, spreadsheetId) {
  await ensureHeaders(sheets, spreadsheetId);
  const values = await getValues(sheets, spreadsheetId, `${SHEET_LISTAS}!A1:H200`);
  const sellers = values.slice(1).map(r => String(r[0] || '').trim()).filter(Boolean);
  const cell = (r,c) => values[r-1] && values[r-1][c-1] != null ? values[r-1][c-1] : '';
  let startNum = Number(cell(2,7)) || 1;
  let endNum = Number(cell(3,7)) || 500;
  if (endNum < startNum) endNum = startNum;
  const precio1 = Number(cell(2,5)) || 0;
  const precio2 = Number(cell(3,5)) || 0;
  const promoActiva = parseBool(cell(4,5));
  const publicidad = {
    titulo: String(cell(2,8) || ''),
    premios: String(cell(3,8) || ''),
    contacto: String(cell(4,8) || ''),
    imagenId: String(cell(5,8) || ''),
    textoExtra: String(cell(6,8) || ''),
    imagenNombre: String(cell(7,8) || ''),
    mensajeInferior: String(cell(8,8) || '')
  };
  return { sellers, startNum, endNum, precio1, precio2, promoActiva, publicidad };
}

function ventasHeaderMap(headers) {
  const normalized = headers.map(normHeader);
  const find = (name) => {
    const i = normalized.indexOf(normHeader(name));
    if (i === -1) throw new Error(`Falta columna ${name} en VENTAS.`);
    return i;
  };
  return {
    numero: find('Numero'), nombreApellido: find('NombreApellido'), dni: find('DNI'),
    telefono: find('Telefono'), vendedor: find('Vendedor'), fecha: find('Fecha'),
    precioCobrado: normalized.indexOf(normHeader('PrecioCobrado'))
  };
}

async function getVentas(sheets, spreadsheetId) {
  await ensureHeaders(sheets, spreadsheetId);
  const values = await getValues(sheets, spreadsheetId, `${SHEET_VENTAS}!A1:G`);
  if (!values.length) return { headers: VENTAS_HEADERS, rows: [], map: ventasHeaderMap(VENTAS_HEADERS) };
  const headers = values[0];
  const map = ventasHeaderMap(headers);
  const rows = values.slice(1).filter(r => String(r[map.numero] || '').trim());
  return { headers, rows, map };
}

async function getDashboard(sheets, spreadsheetId) {
  const cfg = await getListasConfig(sheets, spreadsheetId);
  const ventas = await getVentas(sheets, spreadsheetId);
  const soldSet = new Set();
  ventas.rows.forEach(r => {
    const n = Number(String(r[ventas.map.numero] || '').trim());
    if (Number.isFinite(n) && n >= cfg.startNum && n <= cfg.endNum) soldSet.add(n);
  });
  const soldNumbers = Array.from(soldSet).sort((a,b) => a-b);
  const availableNumbers = [];
  for (let n = cfg.startNum; n <= cfg.endNum; n++) if (!soldSet.has(n)) availableNumbers.push(n);
  const winners = await getWinners(sheets, spreadsheetId);
  return { startNum: cfg.startNum, endNum: cfg.endNum, total: cfg.endNum-cfg.startNum+1, soldCount: soldNumbers.length, availableCount: availableNumbers.length, soldNumbers, availableNumbers, winners };
}

async function getWinners(sheets, spreadsheetId) {
  await ensureHeaders(sheets, spreadsheetId);
  const values = await getValues(sheets, spreadsheetId, `${SHEET_WINNERS}!A2:F`);
  return values.filter(r => String(r[0]||'').trim()).map(r => ({
    numero: Number(r[0]), nombreApellido: String(r[1]||''), dni: String(r[2]||''), telefono: String(r[3]||''), vendedor: String(r[4]||''), fechaSorteo: String(r[5]||'')
  }));
}

async function apiCheckRifa(ctx, data) {
  const rifa = await findRifa(ctx.sheets, data.codigoRifa || data.codigo);
  return { codigo: rifa.codigo, nombre: rifa.nombre };
}

async function apiInit(ctx, data) {
  const rifa = await findRifa(ctx.sheets, data.codigoRifa || data.codigo);
  await ensureHeaders(ctx.sheets, rifa.sheetId);
  const listas = await getListasConfig(ctx.sheets, rifa.sheetId);
  const dashboard = await getDashboard(ctx.sheets, rifa.sheetId);
  return {
    version: 'netlify-sheets-api-v1',
    rifa: { codigo: rifa.codigo, nombre: rifa.nombre },
    dashboard,
    vendedores: listas.sellers,
    config: { precio1: listas.precio1, precio2: listas.precio2, promoActiva: listas.promoActiva, publicidad: listas.publicidad }
  };
}

async function apiConfirmSale(ctx, data) {
  const rifa = await findRifa(ctx.sheets, data.codigoRifa);
  const payload = data.payload || data;
  await ensureHeaders(ctx.sheets, rifa.sheetId);
  const cfg = await getListasConfig(ctx.sheets, rifa.sheetId);
  const tipoPrecio = String(payload.tipoPrecio || '1');
  const numero1 = Number(payload.numero);
  const numero2 = Number(payload.numero2);
  if (!Number.isFinite(numero1)) throw new Error('Número inválido.');
  if (numero1 < cfg.startNum || numero1 > cfg.endNum) throw new Error('Número fuera de rango.');
  const numeros = [numero1];
  if (tipoPrecio === '2') {
    if (!Number.isFinite(numero2)) throw new Error('Seleccioná el segundo número.');
    if (numero2 < cfg.startNum || numero2 > cfg.endNum) throw new Error('Segundo número fuera de rango.');
    if (numero1 === numero2) throw new Error('Los dos números no pueden ser iguales.');
    numeros.push(numero2);
  }
  const nombreApellido = String(payload.nombreApellido || '').trim();
  const dni = String(payload.dni || '').trim();
  const telefono = String(payload.telefono || '').trim();
  const vendedorRaw = String(payload.vendedor || '').trim();
  if (nombreApellido.length < 3) throw new Error('Ingresá nombre y apellido.');
  if (dni.length < 6) throw new Error('Ingresá DNI.');
  if (telefono.length < 6) throw new Error('Ingresá teléfono.');
  if (!vendedorRaw) throw new Error('Seleccioná un vendedor.');
  const vendedor = cfg.sellers.find(v => v.toLowerCase() === vendedorRaw.toLowerCase());
  if (!vendedor) throw new Error('El vendedor seleccionado no existe.');
  const ventas = await getVentas(ctx.sheets, rifa.sheetId);
  const sold = new Set(ventas.rows.map(r => Number(r[ventas.map.numero])).filter(Number.isFinite));
  numeros.forEach(n => { if (sold.has(n)) throw new Error(`El número ${n} ya está vendido.`); });
  const precioPorNumero = tipoPrecio === '2' ? (cfg.precio2 / 2) : cfg.precio1;
  const fecha = nowIso();
  const rows = numeros.map(n => [n, nombreApellido, dni, telefono, vendedor, fecha, precioPorNumero]);
  await appendValues(ctx.sheets, rifa.sheetId, `${SHEET_VENTAS}!A:G`, rows);
  return { msg: `OK - vendidos: ${numeros.join(', ')}`, vendidos: numeros, venta: { numeros, nombreApellido, dni, telefono, vendedor, fecha, precioPorNumero, rifa: rifa.nombre } };
}

async function apiGetSale(ctx, data) {
  const rifa = await findRifa(ctx.sheets, data.codigoRifa);
  const numero = Number(data.numero);
  const ventas = await getVentas(ctx.sheets, rifa.sheetId);
  const row = ventas.rows.find(r => Number(r[ventas.map.numero]) === numero);
  if (!row) throw new Error('Ese número no está vendido.');
  return { numero, nombreApellido: row[ventas.map.nombreApellido] || '', dni: row[ventas.map.dni] || '', telefono: row[ventas.map.telefono] || '', vendedor: row[ventas.map.vendedor] || '', fecha: row[ventas.map.fecha] || '' };
}

async function apiSellerSummary(ctx, data) {
  const rifa = await findRifa(ctx.sheets, data.codigoRifa);
  const ventas = await getVentas(ctx.sheets, rifa.sheetId);
  const resumen = {};
  ventas.rows.forEach(r => {
    const vendedor = String(r[ventas.map.vendedor] || 'Sin vendedor').trim();
    const precio = ventas.map.precioCobrado >= 0 ? Number(r[ventas.map.precioCobrado]) || 0 : 0;
    if (!resumen[vendedor]) resumen[vendedor] = { vendedor, cantidad: 0, totalEstimado: 0 };
    resumen[vendedor].cantidad++;
    resumen[vendedor].totalEstimado += precio;
  });
  return Object.values(resumen).sort((a,b) => b.cantidad-a.cantidad);
}

async function apiAdminLogin(ctx, data) {
  const rifa = await assertAdmin(ctx.sheets, data.codigoRifa, data.pin);
  return { msg: 'OK', codigo: rifa.codigo, nombre: rifa.nombre };
}

async function apiAdminSaveConfig(ctx, data) {
  const rifa = await assertAdmin(ctx.sheets, data.codigoRifa, data.pin);
  const config = data.config || {};
  const pub = config.publicidad || {};
  const precio1 = Number(config.precio1 || 0);
  const precio2 = Number(config.precio2 || 0);
  if (precio1 < 0 || precio2 < 0) throw new Error('Precio inválido.');
  const rows = [
    ['Vendedores','Ventas','','Total $','Precio unitario','', 'Publicidad',''],
    ['', '', '', '1 número', precio1, 'Desde', '', String(pub.titulo || '')],
    ['', '', '', '2 números', precio2, 'Hasta', '', String(pub.premios || '')],
    ['', '', '', 'Promo 2x activa', config.promoActiva ? 'SI' : 'NO', 'Contacto', '', String(pub.contacto || '')],
    ['', '', '', '', '', '', 'Imagen fondo ID', pub.imagenId || ''],
    ['', '', '', '', '', '', 'Texto extra', String(pub.textoExtra || '')],
    ['', '', '', '', '', '', 'Imagen fondo nombre', pub.imagenNombre || ''],
    ['', '', '', '', '', '', 'Mensaje inferior general', String(pub.mensajeInferior || '')]
  ];
  // No pisar F/G rango ni H5/H7 si no vienen.
  const current = await getValues(ctx.sheets, rifa.sheetId, `${SHEET_LISTAS}!A1:H8`);
  rows[1][6] = current[1]?.[6] || 1;
  rows[2][6] = current[2]?.[6] || 500;
  rows[4][7] = pub.imagenId || current[4]?.[7] || '';
  rows[6][7] = pub.imagenNombre || current[6]?.[7] || '';
  await updateValues(ctx.sheets, rifa.sheetId, `${SHEET_LISTAS}!A1:H8`, rows);
  return 'OK';
}

async function apiAdminSaveRange(ctx, data) {
  const rifa = await assertAdmin(ctx.sheets, data.codigoRifa, data.pin);
  const startNum = Number(data.startNum);
  const endNum = Number(data.endNum);
  if (!Number.isFinite(startNum) || !Number.isFinite(endNum) || startNum < 1 || endNum <= startNum) throw new Error('Rango inválido.');
  await updateValues(ctx.sheets, rifa.sheetId, `${SHEET_LISTAS}!F2:G3`, [['Desde', startNum], ['Hasta', endNum]]);
  return 'OK';
}

async function apiAdminAddSeller(ctx, data) {
  const rifa = await assertAdmin(ctx.sheets, data.codigoRifa, data.pin);
  const nombre = String(data.nombre || '').trim();
  if (!nombre) throw new Error('Ingresá el nombre del vendedor.');
  const cfg = await getListasConfig(ctx.sheets, rifa.sheetId);
  if (cfg.sellers.some(v => v.toLowerCase() === nombre.toLowerCase())) throw new Error('Ese vendedor ya existe.');
  await appendValues(ctx.sheets, rifa.sheetId, `${SHEET_LISTAS}!A:A`, [[nombre]]);
  return 'OK';
}

async function apiAdminDeleteSeller(ctx, data) {
  const rifa = await assertAdmin(ctx.sheets, data.codigoRifa, data.pin);
  const nombre = String(data.nombre || '').trim();
  if (!nombre) throw new Error('Seleccioná un vendedor.');
  const sheetId = await getSheetIdByName(ctx.sheets, rifa.sheetId, SHEET_LISTAS);
  const values = await getValues(ctx.sheets, rifa.sheetId, `${SHEET_LISTAS}!A2:A`);
  const idx = values.findIndex(r => String(r[0]||'').trim().toLowerCase() === nombre.toLowerCase());
  if (idx === -1) throw new Error('No encontré ese vendedor.');
  await ctx.sheets.spreadsheets.batchUpdate({ spreadsheetId: rifa.sheetId, requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: idx + 1, endIndex: idx + 2 } } }] } });
  return 'OK';
}

async function apiAdminDeleteSale(ctx, data) {
  const rifa = await assertAdmin(ctx.sheets, data.codigoRifa, data.pin);
  const numero = Number(data.numero);
  const sheetId = await getSheetIdByName(ctx.sheets, rifa.sheetId, SHEET_VENTAS);
  const ventas = await getVentas(ctx.sheets, rifa.sheetId);
  const idx = ventas.rows.findIndex(r => Number(r[ventas.map.numero]) === numero);
  if (idx === -1) throw new Error('Ese número no está vendido.');
  await ctx.sheets.spreadsheets.batchUpdate({ spreadsheetId: rifa.sheetId, requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: idx + 1, endIndex: idx + 2 } } }] } });
  return 'OK';
}

function shuffle(arr) { for (let i=arr.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); [arr[i],arr[j]]=[arr[j],arr[i]]; } return arr; }

async function apiAdminDrawWinners(ctx, data) {
  const rifa = await assertAdmin(ctx.sheets, data.codigoRifa, data.pin);
  const count = Math.floor(Number(data.count));
  if (!Number.isFinite(count) || count < 1) throw new Error('Cantidad inválida.');
  const dashboard = await getDashboard(ctx.sheets, rifa.sheetId);
  const existing = new Set((dashboard.winners || []).map(w => Number(w.numero)));
  const pool = dashboard.soldNumbers.filter(n => !existing.has(n));
  if (!pool.length) throw new Error('No hay números vendidos disponibles para sortear.');
  if (count > pool.length) throw new Error(`No se puede sortear ${count}. Disponibles: ${pool.length}.`);
  const winnersNums = shuffle(pool.slice()).slice(0,count);
  const ventas = await getVentas(ctx.sheets, rifa.sheetId);
  const rows = winnersNums.map(n => {
    const v = ventas.rows.find(r => Number(r[ventas.map.numero]) === n) || [];
    return [n, v[ventas.map.nombreApellido] || '', v[ventas.map.dni] || '', v[ventas.map.telefono] || '', v[ventas.map.vendedor] || '', nowIso()];
  });
  await appendValues(ctx.sheets, rifa.sheetId, `${SHEET_WINNERS}!A:F`, rows);
  return { winners: rows.map(r => ({ numero: r[0], nombreApellido: r[1], dni: r[2], telefono: r[3], vendedor: r[4], fechaSorteo: r[5] })) };
}

async function apiAdminResetWinners(ctx, data) {
  const rifa = await assertAdmin(ctx.sheets, data.codigoRifa, data.pin);
  await updateValues(ctx.sheets, rifa.sheetId, `${SHEET_WINNERS}!A1:F`, [WINNERS_HEADERS]);
  const sheetId = await getSheetIdByName(ctx.sheets, rifa.sheetId, SHEET_WINNERS);
  await ctx.sheets.spreadsheets.batchUpdate({ spreadsheetId: rifa.sheetId, requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: 1, endIndex: 10000 } } }] } }).catch(()=>{});
  return 'OK';
}

async function getOrCreateFolder(ctx, name, parentId) {
  const qParts = [`name='${name.replace(/'/g,"\\'")}'`, "mimeType='application/vnd.google-apps.folder'", 'trashed=false'];
  if (parentId) qParts.push(`'${parentId}' in parents`);
  const res = await ctx.drive.files.list({ q: qParts.join(' and '), fields: 'files(id,name,webViewLink)', spaces: 'drive' });
  if (res.data.files && res.data.files[0]) return res.data.files[0];
  const create = await ctx.drive.files.create({ requestBody: { name, mimeType: 'application/vnd.google-apps.folder', parents: parentId ? [parentId] : undefined }, fields: 'id,name,webViewLink' });
  return create.data;
}

async function apiUploadPublicityBackground(ctx, data) {
  const rifa = await assertAdmin(ctx.sheets, data.codigoRifa, data.pin);
  const dataUrl = String(data.dataUrl || '');
  const fileName = safeName(data.fileName || 'fondo_publicidad.png');
  const match = dataUrl.match(/^data:(image\/[\w.+-]+);base64,(.+)$/);
  if (!match) throw new Error('Formato de imagen inválido.');
  const mimeType = match[1];
  const buffer = Buffer.from(match[2], 'base64');
  if (buffer.length > 8*1024*1024) throw new Error('Imagen demasiado grande. Máximo 8 MB.');
  const pubFolder = await getOrCreateFolder(ctx, 'Publicidad', rifa.folderId || process.env.RIFITA_ROOT_FOLDER_ID);
  const file = await ctx.drive.files.create({ requestBody: { name: `fondo_${normalizeCode(rifa.codigo)}_${Date.now()}_${fileName}`, parents: pubFolder.id ? [pubFolder.id] : undefined }, media: { mimeType, body: require('stream').Readable.from(buffer) }, fields: 'id,name,webViewLink' });
  const cfg = await getListasConfig(ctx.sheets, rifa.sheetId);
  await updateValues(ctx.sheets, rifa.sheetId, `${SHEET_LISTAS}!G5:H7`, [
    ['Imagen fondo ID', file.data.id],
    ['Texto extra', cfg.publicidad.textoExtra || ''],
    ['Imagen fondo nombre', fileName]
  ]);
  return { fileId: file.data.id, name: fileName, folderId: pubFolder.id };
}

async function apiGetPublicityBackgroundData(ctx, data) {
  const rifa = await findRifa(ctx.sheets, data.codigoRifa);
  const cfg = await getListasConfig(ctx.sheets, rifa.sheetId);
  if (!cfg.publicidad.imagenId) return { dataUrl: '' };
  const file = await ctx.drive.files.get({ fileId: cfg.publicidad.imagenId, alt: 'media' }, { responseType: 'arraybuffer' });
  const meta = await ctx.drive.files.get({ fileId: cfg.publicidad.imagenId, fields: 'name,mimeType' }).catch(()=>({ data: {} }));
  const mime = meta.data.mimeType || 'image/png';
  const b64 = Buffer.from(file.data).toString('base64');
  return { dataUrl: `data:${mime};base64,${b64}`, name: meta.data.name || cfg.publicidad.imagenNombre, fileId: cfg.publicidad.imagenId };
}

async function apiListRifas(ctx, data) {
  assertMasterPin(data.adminPin);
  const rifas = await listRifasRaw(ctx.sheets);
  return rifas.map(r => ({ codigo: r.codigo, nombre: r.nombre, sheetId: r.sheetId, pinAdmin: r.pinAdmin, activa: r.activa, fecha: r.fecha, folderId: r.folderId, folderUrl: r.folderUrl }));
}
function assertMasterPin(pin) {
  const wanted = process.env.ADMIN_APP_PIN || '';
  if (wanted && String(pin || '') !== wanted) throw new Error('PIN maestro incorrecto.');
}

async function apiCreateRifa(ctx, data) {
  assertMasterPin(data.adminPin);
  const codigo = normalizeCode(data.codigo || data.rifa?.codigo);
  const nombre = String(data.nombre || data.rifa?.nombre || '').trim();
  const pinAdmin = String(data.pinAdmin || data.rifa?.pinAdmin || '').trim();
  if (!codigo || !nombre || !pinAdmin) throw new Error('Faltan datos para crear la rifa.');
  const existing = await listRifasRaw(ctx.sheets);
  if (existing.some(r => normalizeCode(r.codigo) === codigo)) throw new Error('Ya existe una rifa con ese código.');
  let parentId = process.env.RIFITA_ROOT_FOLDER_ID || '';
  let folder = null;
  if (parentId) folder = await getOrCreateFolder(ctx, safeName(`Rifa - ${nombre} (${codigo})`), parentId);
  const templateId = process.env.TEMPLATE_SS_ID || '';
  let fileId = '';
  if (templateId) {
    const copied = await ctx.drive.files.copy({ fileId: templateId, requestBody: { name: safeName(`Planilla - ${nombre}`), parents: folder?.id ? [folder.id] : undefined }, fields: 'id,webViewLink' });
    fileId = copied.data.id;
  } else {
    const created = await ctx.sheets.spreadsheets.create({ requestBody: { properties: { title: `Planilla - ${nombre}` } } });
    fileId = created.data.spreadsheetId;
    if (folder?.id) await ctx.drive.files.update({ fileId, addParents: folder.id, fields: 'id,parents' }).catch(()=>{});
  }
  await ensureHeaders(ctx.sheets, fileId);
  await appendValues(ctx.sheets, envRequired('MASTER_SS_ID'), `${SHEET_CONFIG_RIFAS}!A:H`, [[codigo, nombre, fileId, pinAdmin, 'SI', nowIso(), folder?.id || '', folder?.webViewLink || '']]);
  return { codigo, nombre, sheetId: fileId, folderId: folder?.id || '', folderUrl: folder?.webViewLink || '' };
}

async function apiToggleRifa(ctx, data) {
  assertMasterPin(data.adminPin);
  const codigo = normalizeCode(data.codigo);
  const rifas = await listRifasRaw(ctx.sheets);
  const r = rifas.find(x => normalizeCode(x.codigo) === codigo);
  if (!r) throw new Error('No encontré esa rifa.');
  await updateValues(ctx.sheets, envRequired('MASTER_SS_ID'), `${SHEET_CONFIG_RIFAS}!E${r.rowNumber}`, [[data.activa ? 'SI' : 'NO']]);
  return 'OK';
}

async function apiDeleteRifa(ctx, data) {
  assertMasterPin(data.adminPin);
  const codigo = normalizeCode(data.codigo);
  const rifas = await listRifasRaw(ctx.sheets);
  const r = rifas.find(x => normalizeCode(x.codigo) === codigo);
  if (!r) throw new Error('No encontré esa rifa.');
  const sheetId = await getSheetIdByName(ctx.sheets, envRequired('MASTER_SS_ID'), SHEET_CONFIG_RIFAS);
  await ctx.sheets.spreadsheets.batchUpdate({ spreadsheetId: envRequired('MASTER_SS_ID'), requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: r.rowNumber-1, endIndex: r.rowNumber } } }] } });
  if (r.folderId) await ctx.drive.files.update({ fileId: r.folderId, requestBody: { trashed: true } }).catch(()=>{});
  else if (r.sheetId) await ctx.drive.files.update({ fileId: r.sheetId, requestBody: { trashed: true } }).catch(()=>{});
  return 'OK';
}

const ACTIONS = {
  checkRifa: apiCheckRifa,
  init: apiInit,
  confirmSale: apiConfirmSale,
  getSale: apiGetSale,
  sellerSummary: apiSellerSummary,
  adminLogin: apiAdminLogin,
  saveConfig: apiAdminSaveConfig,
  saveRange: apiAdminSaveRange,
  addSeller: apiAdminAddSeller,
  deleteSeller: apiAdminDeleteSeller,
  deleteSale: apiAdminDeleteSale,
  drawWinners: apiAdminDrawWinners,
  resetWinners: apiAdminResetWinners,
  uploadPublicityBackground: apiUploadPublicityBackground,
  getPublicityBackgroundData: apiGetPublicityBackgroundData,
  listRifas: apiListRifas,
  createRifa: apiCreateRifa,
  toggleRifa: apiToggleRifa,
  deleteRifa: apiDeleteRifa
};

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return response(200, ok('OK'));
  try {
    const ctx = await getClients();
    if (event.httpMethod === 'GET') {
      return response(200, ok({
        service: 'Rifita Netlify API',
        alive: true,
        hasMaster: !!process.env.MASTER_SS_ID,
        hasRootFolder: !!process.env.RIFITA_ROOT_FOLDER_ID,
        actions: Object.keys(ACTIONS)
      }));
    }
    if (event.httpMethod !== 'POST') return response(405, fail('Método no permitido.'));
    const data = JSON.parse(event.body || '{}');
    const action = data.action;
    if (!action || !ACTIONS[action]) throw new Error('Acción inválida: ' + action);
    const result = await ACTIONS[action](ctx, data);
    return response(200, ok(result));
  } catch (e) {
    console.error(e);
    return response(200, fail(e));
  }
};
