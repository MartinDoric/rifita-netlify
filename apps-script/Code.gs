/** =========================
 *  Code.gs — ESTABLE (v8.1 sin Google)
 *  =========================
 *  - Ventas “smart”: VENTAS o fallback GANADORES si ahí están cargadas
 *  - Ganadores de sorteo en hoja separada: SORTEO_GANADORES (no pisa ventas)
 *  - apiInit devuelve dbg para ver qué hoja usa + conteos
 *  - Respuestas: {ok:true,data} / {ok:false,error}
 */

const VERSION = "v8.2-publicidad-fondo-2026-06-21";

const MASTER_SS_ID = "13LQLWXVjjvtTSVhyLHuvfAU1m6XNIpj0FAR7YS-yPVI";
const SHEET_CONFIG_RIFAS = "CONFIG_RIFAS";
const PUBLICITY_FOLDER_NAME = "RIFITA";
const RIFAS_ROOT_FOLDER_NAME = "RIFITA";

let CURRENT_RIFA_CODE = "";
// Hojas
const SHEET_BOARD   = "NUMEROS";
const SHEET_LISTAS  = "LISTAS";

// Ventas: preferimos VENTAS, fallback GANADORES (si ahí quedaron cargadas por error)
const SHEET_VENTAS_PRIMARY  = "VENTAS";
const SHEET_VENTAS_FALLBACK = "GANADORES";

// Ganadores del sorteo: SIEMPRE otra hoja
const SHEET_WINNERS = "SORTEO_GANADORES";

const VENTAS_HEADERS = ["Numero", "NombreApellido", "DNI", "Telefono", "Vendedor", "Fecha"];



// Slides
const SLIDES_ID = "1oaklvO8PeqhO3svkhPWmTjPARS9v_nW-tXseAQIMz74";
const SLIDES_TABLE_X = 40;
const SLIDES_TABLE_Y = 5;
const SLIDES_TABLE_W = 640;
const SLIDES_TABLE_H = 520;
const SLIDE_COLS = 10;
const SLIDE_ROWS = 10;
const PER_SLIDE  = 100;

// Colores tablero
const COLOR_AVAILABLE = "#C6EFCE";
const COLOR_SOLD      = "#FFC7CE";



// Cache key
const CACHE_DASH_KEY_PREFIX = "dash_v2_";

function getDashCacheKey_(){
  const code = normalizeCode_(CURRENT_RIFA_CODE);
  if(!code) throw new Error("Falta código de rifa para cache.");
  return CACHE_DASH_KEY_PREFIX + code;
}

/** ========= Respuestas ========= */
function ok_(data){ return { ok:true, data:data ?? null }; }
function fail_(err){
  const msg = (err && err.message) ? err.message : String(err);
  return { ok:false, error: msg };
}

/** ========= Spreadsheet helpers ========= */
function normalizeCode_(code) {
  return String(code || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

function setCurrentRifa_(codigoRifa) {
  const code = normalizeCode_(codigoRifa || CURRENT_RIFA_CODE);

  if (!code) {
    throw new Error("Falta el código de rifa.");
  }

  CURRENT_RIFA_CODE = code;
  return code;
}

function getRifaConfig_(codigoRifa) {
  const code = normalizeCode_(codigoRifa);

  if (!code) {
    throw new Error("Falta el código de rifa.");
  }

  const master = SpreadsheetApp.openById(MASTER_SS_ID);
  const sh = master.getSheetByName(SHEET_CONFIG_RIFAS);

  if (!sh) {
    throw new Error("No existe la hoja CONFIG_RIFAS.");
  }

  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error("No hay rifas configuradas.");
  }

  const headers = values[0].map(h => String(h).trim().toUpperCase());

  const idxCodigo = headers.indexOf("CODIGO");
  const idxNombre = headers.indexOf("NOMBRE");
  const idxSheetId = headers.indexOf("SHEET_ID");
  const idxPin = headers.indexOf("PIN_ADMIN");
  const idxActiva = headers.indexOf("ACTIVA");
  const idxFolderId = headers.indexOf("FOLDER_ID");
  const idxFolderUrl = headers.indexOf("FOLDER_URL");

  if (
    idxCodigo === -1 ||
    idxNombre === -1 ||
    idxSheetId === -1 ||
    idxPin === -1 ||
    idxActiva === -1
  ) {
    throw new Error("CONFIG_RIFAS debe tener CODIGO, NOMBRE, SHEET_ID, PIN_ADMIN y ACTIVA.");
  }

  for (let i = 1; i < values.length; i++) {
    const rowCode = normalizeCode_(values[i][idxCodigo]);

    if (rowCode === code) {
      const activa = String(values[i][idxActiva] || "").trim().toUpperCase();

      if (activa !== "SI") {
        throw new Error("Esta rifa no está activa.");
      }

      return {
        codigo: rowCode,
        nombre: String(values[i][idxNombre] || "").trim(),
        sheetId: String(values[i][idxSheetId] || "").trim(),
        pinAdmin: String(values[i][idxPin] || "").trim(),
        folderId: idxFolderId === -1 ? "" : String(values[i][idxFolderId] || "").trim(),
        folderUrl: idxFolderUrl === -1 ? "" : String(values[i][idxFolderUrl] || "").trim(),
        rowNumber: i + 1
      };
    }
  }

  throw new Error("Código de rifa inválido.");
}

/** ========= Carpetas por rifa ========= */
function safeDriveName_(name) {
  return String(name || "")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .substring(0, 120) || "sin_nombre";
}

function getOrCreateRifasRootFolder_() {
  const it = DriveApp.getFoldersByName(RIFAS_ROOT_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(RIFAS_ROOT_FOLDER_NAME);
}

function folderHasParent_(folder, parentId) {
  const parents = folder.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === parentId) return true;
  }
  return false;
}

function ensureFolderInsideRifitaRoot_(folder) {
  const root = getOrCreateRifasRootFolder_();

  if (!folder || folder.getId() === root.getId()) {
    return folder;
  }

  // Si la carpeta ya existía en otro lugar, la movemos dentro de RIFITA.
  try {
    if (!folderHasParent_(folder, root.getId())) {
      folder.moveTo(root);
    }
  } catch (_) {}

  return folder;
}

function ensureConfigFolderColumns_() {
  const master = SpreadsheetApp.openById(MASTER_SS_ID);
  const sh = master.getSheetByName(SHEET_CONFIG_RIFAS);
  if (!sh) throw new Error("No existe la hoja CONFIG_RIFAS.");

  const lastCol = Math.max(1, sh.getLastColumn());
  let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || "").trim().toUpperCase());

  const ensure = (name) => {
    let idx = headers.indexOf(name);
    if (idx !== -1) return idx + 1;

    const col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue(name);
    headers.push(name);
    return col;
  };

  return {
    folderIdCol: ensure("FOLDER_ID"),
    folderUrlCol: ensure("FOLDER_URL")
  };
}

function saveRifaFolderInMaster_(cfg, folder) {
  const cols = ensureConfigFolderColumns_();
  const master = SpreadsheetApp.openById(MASTER_SS_ID);
  const sh = master.getSheetByName(SHEET_CONFIG_RIFAS);

  sh.getRange(cfg.rowNumber, cols.folderIdCol).setValue(folder.getId());
  sh.getRange(cfg.rowNumber, cols.folderUrlCol).setValue(folder.getUrl());
}

function getOrCreateRifaFolder_(codigoRifa) {
  const code = setCurrentRifa_(codigoRifa || CURRENT_RIFA_CODE);
  const cfg = getRifaConfig_(code);

  if (cfg.folderId) {
    try {
      return ensureFolderInsideRifitaRoot_(DriveApp.getFolderById(cfg.folderId));
    } catch (_) {
      // Si la carpeta fue borrada, la recreamos abajo.
    }
  }

  const root = getOrCreateRifasRootFolder_();
  const folderName = safeDriveName_("Rifa - " + (cfg.nombre || code) + " (" + code + ")");

  const it = root.getFoldersByName(folderName);
  const folder = it.hasNext() ? it.next() : root.createFolder(folderName);

  saveRifaFolderInMaster_(cfg, folder);

  // Si la planilla de esa rifa estaba suelta en Drive, la metemos dentro de su carpeta.
  try {
    DriveApp.getFileById(cfg.sheetId).moveTo(folder);
  } catch (_) {}

  return folder;
}

function getOrCreateRifaSubFolder_(subFolderName) {
  const parent = getOrCreateRifaFolder_(CURRENT_RIFA_CODE);
  const safeName = safeDriveName_(subFolderName);
  const it = parent.getFoldersByName(safeName);
  if (it.hasNext()) return it.next();
  return parent.createFolder(safeName);
}

function getSS_(codigoRifa){
  const code = setCurrentRifa_(codigoRifa || CURRENT_RIFA_CODE);
  const cfg = getRifaConfig_(code);
  return SpreadsheetApp.openById(cfg.sheetId);
}

function getSheet_(name, codigoRifa){
  const ss = getSS_(codigoRifa || CURRENT_RIFA_CODE);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error("No existe la hoja: " + name);
  return sh;
}


/** ========= VENDEDORES ========= */
function getSellers_(codigoRifa) {
  setCurrentRifa_(codigoRifa || CURRENT_RIFA_CODE);
  const sh = getSheet_(SHEET_LISTAS);
  const last = sh.getLastRow();

  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, 1)
    .getValues()
    .flat()
    .map(v => String(v || "").trim())
    .filter(Boolean);
}

function findSellerName_(nombre, codigoRifa) {
  const wanted = String(nombre || "").trim();
  if (!wanted) return "";

  const sellers = getSellers_(codigoRifa || CURRENT_RIFA_CODE);
  const found = sellers.find(v => v.toLowerCase() === wanted.toLowerCase());

  return found || "";
}

/** ========= MENÚ ========= */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Rifa")
      .addItem("Actualizar tablero (Sheets)", "updateNumberBoard")
      .addItem("Actualizar grilla (Slides)", "updateFlyerBoardSlidesPaged_")
      .addSeparator()
      .addItem("DEBUG validar", "DEBUG_validate_")
      .addToUi();
  } catch (e) {
    console.log("Menú no disponible en este contexto: " + e.message);
  }
}

function getRaffleConfig_(){
  const sh = getSheet_(SHEET_LISTAS);

  let startNum = Number(sh.getRange("G2").getValue()) || 1;
  let endNum = Number(sh.getRange("G3").getValue()) || 500;

  if (!sh.getRange("F2").getValue()) sh.getRange("F2").setValue("Desde");
  if (!sh.getRange("F3").getValue()) sh.getRange("F3").setValue("Hasta");
  if (!sh.getRange("G2").getValue()) sh.getRange("G2").setValue(startNum);
  if (!sh.getRange("G3").getValue()) sh.getRange("G3").setValue(endNum);

  return { startNum, endNum };
}

function apiAdminSaveRange(startNum, endNum, pin, codigoRifa){
  try {
    setCurrentRifa_(codigoRifa);
    assertAdmin_(pin);

    startNum = Number(startNum);
    endNum = Number(endNum);

    if (!startNum || !endNum || startNum < 1 || endNum <= startNum) {
      throw new Error("Rango inválido");
    }

    const sh = getSheet_(SHEET_LISTAS);

    sh.getRange("F2").setValue("Desde");
    sh.getRange("G2").setValue(startNum);

    sh.getRange("F3").setValue("Hasta");
    sh.getRange("G3").setValue(endNum);

    invalidateDashCache_();

    return ok_("OK");

  } catch (e) {
    return fail_(e);
  }
}
function apiAdminAddSeller(nombre, pin, codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);
    assertAdmin_(pin);

    nombre = String(nombre || "").trim();

    if (!nombre) {
      throw new Error("Ingresá el nombre del vendedor.");
    }

    const sh = getSheet_(SHEET_LISTAS);

    if (!sh.getRange("A1").getValue()) {
      sh.getRange("A1:E1").setValues([["Vendedores","Ventas","","Total $","Precio unitario"]]);
    }

    const last = sh.getLastRow();

    const vendedores = last >= 2
      ? sh.getRange(2, 1, last - 1, 1)
          .getValues()
          .flat()
          .map(v => String(v || "").trim())
          .filter(Boolean)
      : [];

    const exists = vendedores.some(v => v.toLowerCase() === nombre.toLowerCase());

    if (exists) {
      throw new Error("Ese vendedor ya existe.");
    }

    sh.appendRow([nombre]);

    return ok_("OK");

  } catch (e) {
    return fail_(e);
  }
}

function apiAdminDeleteSeller(nombre, pin, codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);
    assertAdmin_(pin);

    nombre = String(nombre || "").trim();

    if (!nombre) {
      throw new Error("Seleccioná un vendedor.");
    }

    const sh = getSheet_(SHEET_LISTAS);
    const last = sh.getLastRow();

    if (last < 2) {
      throw new Error("No hay vendedores para eliminar.");
    }

    const values = sh.getRange(2, 1, last - 1, 1).getValues();

    for (let i = 0; i < values.length; i++) {
      const actual = String(values[i][0] || "").trim();

      if (actual.toLowerCase() === nombre.toLowerCase()) {
        sh.deleteRow(i + 2);
        return ok_("OK");
      }
    }

    throw new Error("No encontré ese vendedor.");

  } catch (e) {
    return fail_(e);
  }
}


function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

/** ========= WEB APP ========= */
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : "";
    const codigoRifa = e && e.parameter ? e.parameter.codigoRifa : "";

    if (action === "__call") {
      let args = [];
      try {
        args = JSON.parse(e.parameter.args || "[]");
      } catch (_) {
        args = [];
      }
      return json_(apiNetlifyCall_(e.parameter.method, args, codigoRifa));
    }

    if (action === "dashboard") {
      setCurrentRifa_(codigoRifa);
      return json_(apiGetDashboard());
    }

    if (action === "init") {
      return json_(apiInitFast(codigoRifa));
    }

    if (action === "sale") {
      const numero = e.parameter.numero;
      return json_(apiGetSale(numero, codigoRifa));
    }

    if (action === "sellerSummary") {
      return json_(apiGetSellerSummary(codigoRifa));
    }

    if (action === "config") {
      return json_(apiGetConfig(codigoRifa));
    }

    if (action === "winners") {
      return json_(apiGetWinners(codigoRifa));
    }

    if (action === "adminLogin") {
      return json_(apiAdminLogin(e.parameter.pin, codigoRifa));
    }

    if (action === "deleteSale") {
      return json_(apiAdminDeleteSale(e.parameter.numero, e.parameter.pin, codigoRifa));
    }

    if (action === "drawWinners") {
      return json_(apiAdminDrawWinners(e.parameter.count, e.parameter.pin, codigoRifa));
    }

    if (action === "resetWinners") {
      return json_(apiAdminResetWinners(e.parameter.pin, codigoRifa));
    }

    if (action === "soldPdf") {
      return json_(apiAdminCreateSoldNumbersPdf(e.parameter.pin, codigoRifa));
    }

    if (action === "sellerPdf") {
      return json_(apiAdminCreateSellerPdf(e.parameter.pin, codigoRifa));
    }

    const t = HtmlService.createTemplateFromFile("MobileApp");
    t.APP_VERSION = VERSION;

    return t.evaluate()
      .setTitle("Rifita")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    return json_(fail_(err));
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const action = body.action || "";
    const codigoRifa = body.codigoRifa || "";

    if (action === "__call") {
      return json_(apiNetlifyCall_(body.method, body.args || [], codigoRifa));
    }

    if (action === "confirmSale") {
      return json_(confirmSale(body.payload, codigoRifa));
    }

    if (action === "saveConfig") {
      return json_(apiAdminSaveConfig(body.config, body.pin, codigoRifa));
    }

    if (action === "saveRange") {
      return json_(apiAdminSaveRange(body.startNum, body.endNum, body.pin, codigoRifa));
    }


    return json_(fail_("Acción POST no válida: " + action));

  } catch (err) {
    return json_(fail_(err));
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** ========= INIT ========= */
function apiInit() {
  try {
    ensureSheetsAndHeaders_();
    ensureWinnersSheet_();

    const soldSet = getSoldSet_();
    const soldNumbers = Array.from(soldSet).sort((a, b) => a - b);

    const availableNumbers = [];
    for (let n = START_NUM; n <= END_NUM; n++) {
      if (!soldSet.has(n)) availableNumbers.push(n);
    }

    const shL = getSheet_(SHEET_LISTAS);
    const last = shL.getLastRow();

    const vendedores = last >= 2
      ? shL.getRange(2, 1, last - 1, 1)
          .getValues()
          .flat()
          .map(v => String(v).trim())
          .filter(Boolean)
      : [];

    const ventasSh = getVentasSheetSmart_();

    return ok_({
      version: VERSION,
      dashboard: {
        startNum: START_NUM,
        endNum: END_NUM,
        total: END_NUM - START_NUM + 1,
        soldCount: soldNumbers.length,
        availableCount: availableNumbers.length,
        soldNumbers,
        availableNumbers,
        winners: []
      },
      vendedores,
      dbg: {
        ventasSheet: ventasSh.getName(),
        ventasLastRow: ventasSh.getLastRow(),
        soldCount: soldNumbers.length,
        availableCount: availableNumbers.length,
        timestamp: new Date().toISOString()
      }
    });

  } catch (e) {
    return fail_(e);
  }
}
function apiInitFast(codigoRifa) {
  try {
    const code = setCurrentRifa_(codigoRifa);
    const rifaCfg = getRifaConfig_(code);
    const ss = SpreadsheetApp.openById(rifaCfg.sheetId);
    let shListas = ss.getSheetByName(SHEET_LISTAS);
    if (!shListas) shListas = ss.insertSheet(SHEET_LISTAS);

    let startNum = Number(shListas.getRange("G2").getValue());
    let endNum = Number(shListas.getRange("G3").getValue());

    if (!Number.isFinite(startNum) || startNum < 1) startNum = 1;
    if (!Number.isFinite(endNum) || endNum < startNum) endNum = 500;

    if (endNum - startNum + 1 > 1000) {
      throw new Error(
        `Rango demasiado grande: ${startNum}–${endNum}. Revisá LISTAS!G2 y LISTAS!G3.`
      );
    }

    shListas.getRange("F2").setValue("Desde");
    shListas.getRange("G2").setValue(startNum);
    shListas.getRange("F3").setValue("Hasta");
    shListas.getRange("G3").setValue(endNum);

    let shVentas = ss.getSheetByName(SHEET_VENTAS_PRIMARY);
    if (!shVentas) shVentas = ss.insertSheet(SHEET_VENTAS_PRIMARY);

    if (shVentas.getLastRow() === 0 || !shVentas.getRange("A1").getValue()) {
      shVentas.getRange(1, 1, 1, VENTAS_HEADERS.length).setValues([VENTAS_HEADERS]);
    }

    const map = headerMapVentas_(shVentas);
    const lastVentas = shVentas.getLastRow();
    const soldSet = new Set();

    if (lastVentas >= 2) {
      const nums = shVentas
        .getRange(2, map.Numero, lastVentas - 1, 1)
        .getValues()
        .flat();

      nums.forEach(v => {
        const raw = String(v || "").trim();
        if (!raw) return;

        const n = Number(raw);
        if (Number.isFinite(n) && n >= startNum && n <= endNum) {
          soldSet.add(n);
        }
      });
    }

    const soldNumbers = Array.from(soldSet).sort((a, b) => a - b);

    const availableNumbers = [];
    for (let n = startNum; n <= endNum; n++) {
      if (!soldSet.has(n)) availableNumbers.push(n);
    }

    const lastListas = shListas.getLastRow();

    const vendedores = lastListas >= 2
      ? shListas.getRange(2, 1, lastListas - 1, 1)
          .getValues()
          .flat()
          .map(v => String(v || "").trim())
          .filter(Boolean)
      : [];

    const winnersSafe = getWinners_().map(w => ({
      numero: Number(w.numero),
      nombreApellido: String(w.nombreApellido || ""),
      dni: String(w.dni || ""),
      telefono: String(w.telefono || ""),
      vendedor: String(w.vendedor || ""),
      fechaSorteo: w.fechaSorteo instanceof Date
        ? w.fechaSorteo.toISOString()
        : String(w.fechaSorteo || "")
    }));

    return ok_({
      version: VERSION,
      rifa: {
        codigo: rifaCfg.codigo,
        nombre: rifaCfg.nombre
      },
      dashboard: {
        startNum,
        endNum,
        total: endNum - startNum + 1,
        soldCount: soldNumbers.length,
        availableCount: availableNumbers.length,
        soldNumbers,
        availableNumbers,
        winners: winnersSafe
      },
      vendedores,
      dbg: {
        codigoRifa: CURRENT_RIFA_CODE,
        sheetId: rifaCfg.sheetId,
        ventasSheet: shVentas.getName(),
        ventasLastRow: shVentas.getLastRow(),
        soldCount: soldNumbers.length,
        availableCount: availableNumbers.length,
        vendedoresCount: vendedores.length,
        winnersCount: winnersSafe.length,
        time: new Date().toISOString()
      }
    });

  } catch (e) {
    return fail_(e);
  }
}

/** ========= Dashboard SIN CACHE ========= */
function apiGetDashboardNoCache_(codigoRifa){
  setCurrentRifa_(codigoRifa);

  ensureSheetsAndHeaders_();
  ensureWinnersSheet_();

  const rangeCfg = getRaffleConfig_();
  const START_NUM = rangeCfg.startNum;
  const END_NUM = rangeCfg.endNum;

  const soldSet = getSoldSet_();
  const sold = Array.from(soldSet).sort((a,b)=>a-b);

  const total = END_NUM - START_NUM + 1;
  const availableNumbers = [];

  for (let n = START_NUM; n <= END_NUM; n++){
    if (!soldSet.has(n)) availableNumbers.push(n);
  }

  return {
    startNum: START_NUM,
    endNum: END_NUM,
    total,
    soldCount: sold.length,
    availableCount: total - sold.length,
    soldNumbers: sold,
    availableNumbers,
    winners: getWinners_()
  };
}

/** ========= Dashboard con cache ========= */
function apiGetDashboard() {
  try {
    ensureSheetsAndHeaders_();
    ensureWinnersSheet_();

    const cache = CacheService.getScriptCache();
    const cached = cache.get(getDashCacheKey_());

    if (cached) {
      try {
        const obj = JSON.parse(cached);
        const ok =
          obj && typeof obj === "object" &&
          obj.startNum != null && obj.endNum != null &&
          typeof obj.soldCount === "number" &&
          typeof obj.availableCount === "number" &&
          Array.isArray(obj.availableNumbers) &&
          Array.isArray(obj.soldNumbers);

        if (ok) return ok_(obj);
        cache.remove(getDashCacheKey_());
      } catch (_) {
        cache.remove(getDashCacheKey_());
      }
    }

    const res = apiGetDashboardNoCache_();
    cache.put(getDashCacheKey_(), JSON.stringify(res), 30);
    return ok_(res);

  } catch (e){
    return fail_(e);
  }
}

function apiGetSale(numero, codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);
    return ok_(getSaleByNumber_(numero));
  } catch (e){
    return fail_(e);
  }
}

/** ========= VENDER ========= */
function confirmSale(payload, codigoRifa) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    setCurrentRifa_(codigoRifa);
    ensureSheetsAndHeaders_();
    const rangeCfg = getRaffleConfig_();
    const START_NUM = rangeCfg.startNum;
    const END_NUM = rangeCfg.endNum;
    const tipoPrecio = String(payload?.tipoPrecio || "1").trim();

    const numero1 = Number(payload?.numero);
    const numero2 = Number(payload?.numero2);

    if (!Number.isFinite(numero1)) throw new Error("Número inválido.");
    if (numero1 < START_NUM || numero1 > END_NUM) throw new Error("Número fuera de rango.");

    const numeros = [numero1];

    if (tipoPrecio === "2") {
      if (!Number.isFinite(numero2)) throw new Error("Seleccioná el segundo número.");
      if (numero2 < START_NUM || numero2 > END_NUM) throw new Error("Segundo número fuera de rango.");
      if (numero1 === numero2) throw new Error("Los dos números no pueden ser iguales.");
      numeros.push(numero2);
    }

    const nombreApellido = String(payload?.nombreApellido || "").trim();
    const dni = String(payload?.dni || "").trim();
    const telefono = String(payload?.telefono || "").trim();
    const vendedorRaw = String(payload?.vendedor || "").trim();
    const vendedor = findSellerName_(vendedorRaw, codigoRifa);

    if (nombreApellido.length < 3) throw new Error("Ingresá nombre y apellido.");
    if (dni.length < 6) throw new Error("Ingresá DNI.");
    if (telefono.length < 6) throw new Error("Ingresá teléfono.");
    if (!vendedorRaw) throw new Error("Seleccioná un vendedor.");
    if (!vendedor) throw new Error("El vendedor seleccionado no existe en LISTAS. Actualizá datos y probá de nuevo.");

    const soldSet = getSoldSet_();

    numeros.forEach(n => {
      if (soldSet.has(n)) throw new Error(`El número ${n} ya está vendido.`);
    });

    const sh = getVentasSheetSmart_();
    ensurePrecioColumn_();

    const map = headerMapVentas_(sh);

    const shL = getSheet_(SHEET_LISTAS);
    const precio1 = Number(shL.getRange("E2").getValue()) || 0;
    const precio2 = Number(shL.getRange("E3").getValue()) || 0;

    const precioPorNumero = tipoPrecio === "2" ? precio2 / 2 : precio1;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const precioCol = headers.indexOf("PrecioCobrado") + 1;

    if (precioCol < 1) throw new Error("No se pudo crear la columna PrecioCobrado.");

    numeros.forEach(n => {
      const row = sh.getLastRow() + 1;

      sh.getRange(row, map.Numero).setValue(n);
      sh.getRange(row, map.NombreApellido).setValue(nombreApellido);
      sh.getRange(row, map.DNI).setValue(dni);
      sh.getRange(row, map.Telefono).setValue(telefono);
      sh.getRange(row, map.Vendedor).setValue(vendedor);
      sh.getRange(row, map.Fecha).setValue(new Date());
      sh.getRange(row, precioCol).setValue(precioPorNumero);
    });

    invalidateDashCache_();

    return ok_({
      msg: `OK - vendidos: ${numeros.join(", ")}`,
      vendidos: numeros
    });

  } catch (e) {
    return fail_(e);

  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

/** ========= ADMIN ========= */
function apiAdminLogin(pass, codigoRifa){
  try {
    setCurrentRifa_(codigoRifa);
    const cfg = assertAdmin_(pass);

    return ok_({
      msg: "OK",
      codigo: cfg.codigo,
      nombre: cfg.nombre
    });

  } catch (e) {
    return fail_(e);
  }
}

function apiAdminDeleteSale(numero, pin, codigoRifa) {
  setCurrentRifa_(codigoRifa);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    assertAdmin_(pin);
    deleteSaleByNumber_(numero);
    invalidateDashCache_();
    return ok_("OK");
  } catch (e){
    return fail_(e);
  } finally {
    lock.releaseLock();
  }
}
function apiAdminCreateSoldNumbersPdf(pin, codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);
    assertAdmin_(pin);

    const sh = getVentasSheetSmart_();
    const last = sh.getLastRow();
    const map = headerMapVentas_(sh);

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const precioCol = headers.indexOf("PrecioCobrado") + 1;

    let html = `
      <h2>Listado completo de números vendidos</h2>
      <p>Fecha: ${new Date().toLocaleString()}</p>
      <table border="1" cellpadding="5" cellspacing="0" width="100%" style="border-collapse:collapse;font-size:11px;">
        <tr>
          <th>Número</th>
          <th>Nombre y apellido</th>
          <th>DNI</th>
          <th>Teléfono</th>
          <th>Vendedor</th>
          <th>Fecha</th>
          <th>Precio</th>
        </tr>
    `;

    if (last >= 2) {
      const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

      data
        .sort((a, b) => Number(a[map.Numero - 1]) - Number(b[map.Numero - 1]))
        .forEach(row => {
          html += `
            <tr>
              <td>${row[map.Numero - 1] || ""}</td>
              <td>${row[map.NombreApellido - 1] || ""}</td>
              <td>${row[map.DNI - 1] || ""}</td>
              <td>${row[map.Telefono - 1] || ""}</td>
              <td>${row[map.Vendedor - 1] || ""}</td>
              <td>${row[map.Fecha - 1] || ""}</td>
              <td>${precioCol > 0 ? "$" + (Number(row[precioCol - 1]) || 0) : ""}</td>
            </tr>
          `;
        });
    }

    html += `</table>`;

    const blob = Utilities.newBlob(html, "text/html", "numeros_vendidos.html")
      .getAs("application/pdf")
      .setName("Numeros vendidos.pdf");

    const reportsFolder = getOrCreateRifaSubFolder_("Reportes");
    const file = reportsFolder.createFile(blob);

    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    const fileId = file.getId();

    return ok_({
      url: "https://drive.google.com/file/d/" + fileId + "/view?usp=sharing",
      downloadUrl: "https://drive.google.com/uc?export=download&id=" + fileId
    });

  } catch (e) {
    return fail_(e);
  }
}

function apiAdminCreateSellerPdf(pin, codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);
    assertAdmin_(pin);

    const sh = getVentasSheetSmart_();
    const last = sh.getLastRow();
    const map = headerMapVentas_(sh);

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const precioCol = headers.indexOf("PrecioCobrado") + 1;

    const resumen = {};

    if (last >= 2) {
      const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

      data.forEach(row => {
        const vendedor = String(row[map.Vendedor - 1] || "Sin vendedor").trim();
        const precio = precioCol > 0 ? Number(row[precioCol - 1]) || 0 : 0;

        if (!resumen[vendedor]) resumen[vendedor] = { cantidad: 0, total: 0 };

        resumen[vendedor].cantidad++;
        resumen[vendedor].total += precio;
      });
    }

    let html = `
      <h2>Resumen de ventas por vendedor</h2>
      <p>Fecha: ${new Date().toLocaleString()}</p>
      <table border="1" cellpadding="6" cellspacing="0" width="100%">
        <tr>
          <th>Vendedor</th>
          <th>Cantidad</th>
          <th>Total generado</th>
        </tr>
    `;

    Object.keys(resumen).forEach(v => {
      html += `
        <tr>
          <td>${v}</td>
          <td>${resumen[v].cantidad}</td>
          <td>$${resumen[v].total}</td>
        </tr>
      `;
    });

    html += `</table>`;

    const blob = Utilities.newBlob(html, "text/html", "resumen.html")
      .getAs("application/pdf")
      .setName("Resumen vendedores.pdf");

    const reportsFolder = getOrCreateRifaSubFolder_("Reportes");
    const file = reportsFolder.createFile(blob);

    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    const fileId = file.getId();

    return ok_({
      url: "https://drive.google.com/file/d/" + fileId + "/view?usp=sharing",
      downloadUrl: "https://drive.google.com/uc?export=download&id=" + fileId
    });

  } catch (e) {
    return fail_(e);
  }
}

function apiAdminDrawWinners(count, pin, codigoRifa) {
  setCurrentRifa_(codigoRifa);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    assertAdmin_(pin);

    const winners = apiDrawWinners_(count);

    const winnersSafe = (winners || []).map(w => ({
      numero: Number(w.numero),
      nombreApellido: String(w.nombreApellido || ""),
      dni: String(w.dni || ""),
      telefono: String(w.telefono || ""),
      vendedor: String(w.vendedor || ""),
      fechaSorteo: (w.fechaSorteo instanceof Date)
        ? w.fechaSorteo.toISOString()
        : String(w.fechaSorteo || "")
    }));

    invalidateDashCache_();
    return ok_({ winners: winnersSafe });

  } catch (e){
    return fail_(e);
  } finally {
    lock.releaseLock();
  }
}


function apiAdminResetWinners(pin, codigoRifa) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    setCurrentRifa_(codigoRifa);
    assertAdmin_(pin);

    apiResetWinners_(codigoRifa);
    invalidateDashCache_();

    return ok_("OK");
  } catch (e){
    return fail_(e);
  } finally {
    lock.releaseLock();
  }
}

/** ========= SORTEO ========= */
function apiDrawWinners_(count, codigoRifa) {
  setCurrentRifa_(codigoRifa);
  ensureSheetsAndHeaders_();
  ensureWinnersSheet_();

  const nWinners = Number(count);
  if (!Number.isFinite(nWinners) || nWinners < 1) throw new Error("Cantidad inválida.");

  const sold = Array.from(getSoldSet_());
  if (sold.length === 0) throw new Error("No hay ventas detectadas. No se puede sortear.");

  const existing = new Set(getWinners_().map(w => Number(w.numero)));
  const pool = sold.filter(n => !existing.has(Number(n)));

  if (pool.length === 0) throw new Error("Todos los números vendidos ya ganaron.");
  if (nWinners > pool.length) throw new Error(`No se puede sortear ${nWinners}. Disponibles: ${pool.length}.`);

  shuffleInPlace_(pool);

  const winnersNums = pool.slice(0, nWinners);
  const winners = winnersNums.map(n => {
    const d = getSaleByNumber_(n);
    return {
      numero: n,
      nombreApellido: d.nombreApellido || "",
      dni: d.dni || "",
      telefono: d.telefono || "",
      vendedor: d.vendedor || "",
      fechaSorteo: new Date()
    };
  });

  saveWinners_(winners);
  return winners;
}

function apiResetWinners_(codigoRifa) {
  setCurrentRifa_(codigoRifa);
  ensureWinnersSheet_();
  const sh = getSheet_(SHEET_WINNERS);
  sh.clearContents();
  sh.getRange(1,1,1,6).setValues([["Numero","NombreApellido","DNI","Telefono","Vendedor","FechaSorteo"]]);
}

/** ========= TABLERO SHEETS ========= */
function updateNumberBoard() {
  const soldSet = getSoldSet_();
  const ss = getSS_();
  const rangeCfg = getRaffleConfig_();
  const START_NUM = rangeCfg.startNum;
  const END_NUM = rangeCfg.endNum;

  let sh = ss.getSheetByName(SHEET_BOARD);
  if (!sh) sh = ss.insertSheet(SHEET_BOARD);
  sh.clear();

  const cols = 10;
  const total = END_NUM - START_NUM + 1;
  const rows = Math.ceil(total / cols);

  const grid = Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => {
      const n = START_NUM + r * cols + c;
      return (n <= END_NUM) ? n : "";
    })
  );

  const range = sh.getRange(1, 1, rows, cols);
  range.setValues(grid);

  const backgrounds = grid.map(row =>
    row.map(v => v === "" ? null : (soldSet.has(Number(v)) ? COLOR_SOLD : COLOR_AVAILABLE))
  );
  range.setBackgrounds(backgrounds);
}

/** ========= TABLERO SLIDES ========= */
function updateFlyerBoardSlidesPaged_() {
  if (!SLIDES_ID) return;

  const soldSet = getSoldSet_();
  const rangeCfg = getRaffleConfig_();
  const START_NUM = rangeCfg.startNum;
  const END_NUM = rangeCfg.endNum;
  const totalNums = END_NUM - START_NUM + 1;
  const slidesNeeded = Math.ceil(totalNums / PER_SLIDE);

  const pres = SlidesApp.openById(SLIDES_ID);

  let slides = pres.getSlides();
  while (slides.length < slidesNeeded) {
    pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    slides = pres.getSlides();
  }

  for (let s = 0; s < slidesNeeded; s++) {
    const slide = slides[s];
    const base = START_NUM + s * PER_SLIDE;

    let table = slide.getTables()[0] || null;

    if (!table || table.getNumRows() !== SLIDE_ROWS || table.getNumColumns() !== SLIDE_COLS) {
      if (table) table.remove();
      table = slide.insertTable(SLIDE_ROWS, SLIDE_COLS, SLIDES_TABLE_X, SLIDES_TABLE_Y, SLIDES_TABLE_W, SLIDES_TABLE_H);
    }

    for (let r = 0; r < SLIDE_ROWS; r++) {
      for (let c = 0; c < SLIDE_COLS; c++) {
        const idx = r * SLIDE_COLS + c;
        const n = base + idx;
        const cell = table.getCell(r, c);

        if (n <= END_NUM) {
          cell.getText().setText(String(n));
          cell.getText().getTextStyle().setFontSize(9);
          cell.getFill().setSolidFill(soldSet.has(n) ? COLOR_SOLD : COLOR_AVAILABLE);
          cell.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        } else {
          cell.getText().setText("");
          cell.getFill().setSolidFill("#FFFFFF");
        }
      }
    }
  }

  pres.saveAndClose();
}

/** ========= ENSURE ========= */
function ensureSheetsAndHeaders_() {
  const ss = getSS_();

  // LISTAS
  let shListas = ss.getSheetByName(SHEET_LISTAS);
  if (!shListas) shListas = ss.insertSheet(SHEET_LISTAS);

  if (!shListas.getRange("A1").getValue()) {
    shListas.getRange("A1:E1").setValues([["Vendedores","Ventas","","Total $","Precio unitario"]]);
    shListas.getRange("E2").setValue(0);
  }

  // Crear primaria si no existe
  let shVentas = ss.getSheetByName(SHEET_VENTAS_PRIMARY);
  if (!shVentas) shVentas = ss.insertSheet(SHEET_VENTAS_PRIMARY);

  // Si está vacía, escribir headers
  const headerRange = shVentas.getRange(1, 1, 1, VENTAS_HEADERS.length);
  const current = headerRange.getValues()[0];
  if (current.every(v => !v)) headerRange.setValues([VENTAS_HEADERS]);

  ensureWinnersSheet_();
}

function ensureWinnersSheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(SHEET_WINNERS);
  if (!sh) sh = ss.insertSheet(SHEET_WINNERS);
  if (!sh.getRange("A1").getValue()) {
    sh.getRange(1,1,1,6).setValues([["Numero","NombreApellido","DNI","Telefono","Vendedor","FechaSorteo"]]);
  }
}

/** ========= SMART VENTAS ========= */
function ventasHeadersMatch_(sh){
  const want = ["numero","nombreapellido","dni","telefono","vendedor","fecha"];
  const row = sh.getRange(1,1,1,Math.max(6, sh.getLastColumn())).getValues()[0];

  const norm = (s)=>String(s||"").trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g,"");

  const got = row.slice(0,6).map(norm);
  return want.every((w,i)=>got[i] === w);
}

function getVentasSheetSmart_(){
  const ss = getSS_();
  const shV = ss.getSheetByName(SHEET_VENTAS_PRIMARY);
  const shF = ss.getSheetByName(SHEET_VENTAS_FALLBACK);

  // Preferimos VENTAS si tiene headers correctos y tiene data
  if (shV && ventasHeadersMatch_(shV)) {
    if (shV.getLastRow() >= 2) return shV;
  }

  // Fallback si GANADORES tiene headers y data
  if (shF && ventasHeadersMatch_(shF)) {
    if (shF.getLastRow() >= 2) return shF;
  }

  // Si VENTAS tiene headers, aunque esté vacía, la usamos
  if (shV && ventasHeadersMatch_(shV)) return shV;

  throw new Error(
    `No encuentro hoja de ventas válida. Necesito headers: ${VENTAS_HEADERS.join(",")}`
  );
}

function getWinners_() {
  const ss = getSS_();

  let sh = ss.getSheetByName(SHEET_WINNERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_WINNERS);
    sh.getRange(1, 1, 1, 6).setValues([
      ["Numero", "NombreApellido", "DNI", "Telefono", "Vendedor", "FechaSorteo"]
    ]);
    return [];
  }

  if (!sh.getRange("A1").getValue()) {
    sh.getRange(1, 1, 1, 6).setValues([
      ["Numero", "NombreApellido", "DNI", "Telefono", "Vendedor", "FechaSorteo"]
    ]);
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, 6).getValues();

  return data
    .filter(r => r[0] !== "" && r[0] !== null)
    .map(r => ({
      numero: Number(r[0]),
      nombreApellido: String(r[1] || ""),
      dni: String(r[2] || ""),
      telefono: String(r[3] || ""),
      vendedor: String(r[4] || ""),
      fechaSorteo: r[5]
    }));
}

function apiGetWinners(codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);

    const winners = getWinners_();

    const safe = winners.map(w => ({
      numero: Number(w.numero),
      nombreApellido: String(w.nombreApellido || ""),
      dni: String(w.dni || ""),
      telefono: String(w.telefono || ""),
      vendedor: String(w.vendedor || ""),
      fechaSorteo: w.fechaSorteo instanceof Date
        ? w.fechaSorteo.toISOString()
        : String(w.fechaSorteo || "")
    }));

    return ok_(safe);

  } catch (e) {
    return fail_(e);
  }
}

function saveWinners_(winners) {
  const sh = getSheet_(SHEET_WINNERS);
  const rows = winners.map(w => [w.numero, w.nombreApellido, w.dni, w.telefono, w.vendedor, w.fechaSorteo]);
  sh.getRange(sh.getLastRow()+1, 1, rows.length, 6).setValues(rows);
}

/** ========= MAP HEADERS ========= */
function headerMapVentas_(sh) {
  const norm = (s) =>
    String(s || "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, "");

  const headersRaw = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const headers = headersRaw.map(h => norm(h));

  const findCol = (name) => {
    const idx = headers.indexOf(norm(name));
    if (idx === -1) throw new Error(`Falta la columna '${name}' en hoja '${sh.getName()}'.`);
    return idx + 1;
  };

  return {
    Numero: findCol("Numero"),
    NombreApellido: findCol("NombreApellido"),
    DNI: findCol("DNI"),
    Telefono: findCol("Telefono"),
    Vendedor: findCol("Vendedor"),
    Fecha: findCol("Fecha")
  };
}

/** ========= SOLD SET ========= */
function getSoldSet_(){
  const sh = getVentasSheetSmart_();
  const last = sh.getLastRow();
  const sold = new Set();
  if (last < 2) return sold;

  const map = headerMapVentas_(sh);
  const data = sh.getRange(2, map.Numero, last - 1, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    const v = data[i][0];
    if (v === "" || v === null) continue;
    const n = Number(String(v).trim());
    if (Number.isFinite(n)) sold.add(n);
  }
  return sold;
}

function getSaleByNumber_(numero) {
  const n = Number(numero);
  if (!Number.isFinite(n)) throw new Error("Número inválido.");

  const sh = getVentasSheetSmart_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error("No hay ventas.");

  const map = headerMapVentas_(sh);
  const nums = sh.getRange(2, map.Numero, lastRow - 1, 1).getValues().flat()
    .map(x => Number(String(x).trim()));
  const idx = nums.findIndex(x => x === n);
  if (idx === -1) throw new Error("Ese número no está vendido.");

  const row = idx + 2;

  return {
    numero: n,
    nombreApellido: sh.getRange(row, map.NombreApellido).getValue(),
    dni: sh.getRange(row, map.DNI).getValue(),
    telefono: sh.getRange(row, map.Telefono).getValue(),
    vendedor: sh.getRange(row, map.Vendedor).getValue()
  };
}

function deleteSaleByNumber_(numero) {
  const n = Number(numero);
  if (!Number.isFinite(n)) throw new Error("Número inválido.");

  const sh = getVentasSheetSmart_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error("No hay ventas.");

  const map = headerMapVentas_(sh);
  const nums = sh.getRange(2, map.Numero, lastRow - 1, 1).getValues().flat()
    .map(x => Number(String(x).trim()));
  const idx = nums.findIndex(x => x === n);
  if (idx === -1) throw new Error("Ese número no está vendido.");

  sh.deleteRow(idx + 2);
}

/** ========= UTILS ========= */
function shuffleInPlace_(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}

function assertAdmin_(pin) {
  const cfg = getRifaConfig_(CURRENT_RIFA_CODE);

  if (String(pin || "") !== String(cfg.pinAdmin)) {
    throw new Error("PIN incorrecto.");
  }

  return cfg;
}

function apiAdminLogin(pass, codigoRifa){
  try {
    setCurrentRifa_(codigoRifa);
    const cfg = assertAdmin_(pass);

    return ok_({
      msg: "OK",
      codigo: cfg.codigo,
      nombre: cfg.nombre
    });

  } catch (e) {
    return fail_(e);
  }
}
function invalidateDashCache_() {
  CacheService.getScriptCache().remove(getDashCacheKey_());
}

/** ========= DEBUG ========= */
function DEBUG_validate_(){
  ensureSheetsAndHeaders_();
  ensureWinnersSheet_();
  const sh = getVentasSheetSmart_();
  const sold = Array.from(getSoldSet_());
  Logger.log("VENTAS SHEET USED: " + sh.getName());
  Logger.log("LAST ROW: " + sh.getLastRow());
  Logger.log("SOLD COUNT: " + sold.length);
  Logger.log("SOLD SAMPLE: " + sold.slice(0,20).join(","));
  return "OK";
}

function apiPing() {
  return ok_({
    msg: "Apps Script responde OK",
    version: VERSION,
    time: new Date().toISOString()
  });
}
function apiCheckRifa(codigoRifa) {
  try {
    const cfg = getRifaConfig_(codigoRifa);

    return ok_({
      codigo: cfg.codigo,
      nombre: cfg.nombre
    });

  } catch (e) {
    return fail_(e);
  }
}

/** Compatibilidad con HTML viejo */
function getFormData() {
  const dashboard = apiGetDashboardNoCache_();

  const shL = getSheet_(SHEET_LISTAS);
  const last = shL.getLastRow();

  const vendedores = last >= 2
    ? shL.getRange(2, 1, last - 1, 1)
        .getValues()
        .flat()
        .map(v => String(v).trim())
        .filter(Boolean)
    : [];

  const precioUnitario = Number(shL.getRange("E2").getValue()) || 0;

  return {
    numerosDisponibles: dashboard.availableNumbers,
    vendedores,
    precioUnitario
  };
}

function getSoldNumbers() {
  return Array.from(getSoldSet_()).sort((a, b) => a - b);
}

function getSaleByNumber(numero) {
  return getSaleByNumber_(numero);
}

function deleteSaleByNumber(numero) {
  deleteSaleByNumber_(numero);
  updateNumberBoard();
  invalidateDashCache_();
  return "OK";
}

function apiGetConfig(codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);

    const sh = getSheet_(SHEET_LISTAS);

    return ok_({
      precio1: Number(sh.getRange("E2").getValue()) || 0,
      precio2: Number(sh.getRange("E3").getValue()) || 0,
      promoActiva: String(sh.getRange("E4").getValue()).trim().toUpperCase() === "SI",
      pubTitulo: String(sh.getRange("H2").getValue() || ""),
      pubPremios: String(sh.getRange("H3").getValue() || ""),
      pubContacto: String(sh.getRange("H4").getValue() || ""),
      pubImagenId: String(sh.getRange("H5").getValue() || ""),
      pubExtra: String(sh.getRange("H6").getValue() || ""),
      pubImagenNombre: String(sh.getRange("H7").getValue() || ""),
      pubFooterMsg: String(sh.getRange("H8").getValue() || ""),
      publicidad: {
        titulo: String(sh.getRange("H2").getValue() || ""),
        premios: String(sh.getRange("H3").getValue() || ""),
        contacto: String(sh.getRange("H4").getValue() || ""),
        imagenId: String(sh.getRange("H5").getValue() || ""),
        textoExtra: String(sh.getRange("H6").getValue() || ""),
        imagenNombre: String(sh.getRange("H7").getValue() || ""),
        mensajeInferior: String(sh.getRange("H8").getValue() || "")
      }
    });

  } catch (e) {
    return fail_(e);
  }
}

function apiAdminSavePrices(precio1, precio2, pin) {
  try {
    assertAdmin_(pin);

    precio1 = Number(precio1);
    precio2 = Number(precio2);

    if (!Number.isFinite(precio1) || precio1 < 0) throw new Error("Precio 1 inválido.");
    if (!Number.isFinite(precio2) || precio2 < 0) throw new Error("Precio 2 inválido.");

    const sh = getSheet_(SHEET_LISTAS);

    sh.getRange("D2").setValue("1 número");
    sh.getRange("E2").setValue(precio1);

    sh.getRange("D3").setValue("2 números");
    sh.getRange("E3").setValue(precio2);

    return ok_("OK");

  } catch (e) {
    return fail_(e);
  }
}

function apiGetSellerSummary(codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);

    const sh = getVentasSheetSmart_();
    const last = sh.getLastRow();

    const shL = getSheet_(SHEET_LISTAS);
    const precio1 = Number(shL.getRange("E2").getValue()) || 0;
    const precio2 = Number(shL.getRange("E3").getValue()) || 0;

    if (last < 2) return ok_([]);

    const map = headerMapVentas_(sh);
    const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const precioCol = headers.indexOf("PrecioCobrado") + 1;

    const resumen = {};

    data.forEach(row => {
      const vendedor = String(row[map.Vendedor - 1] || "").trim() || "Sin vendedor";

      if (!resumen[vendedor]) {
        resumen[vendedor] = {
          vendedor,
          cantidad: 0,
          totalEstimado: 0
        };
      }

      resumen[vendedor].cantidad++;

      if (precioCol > 0) {
        resumen[vendedor].totalEstimado += Number(row[precioCol - 1]) || 0;
      }
    });

    // Si no existe PrecioCobrado, calcula aproximado con precios actuales
    Object.values(resumen).forEach(r => {
      if (r.totalEstimado > 0) return;

      const pares = Math.floor(r.cantidad / 2);
      const sueltos = r.cantidad % 2;

      r.totalEstimado = pares * precio2 + sueltos * precio1;
    });

    return ok_(
      Object.values(resumen)
        .sort((a, b) => b.cantidad - a.cantidad)
    );

  } catch (e) {
    return fail_(e);
  }
}



function apiAdminSaveConfig(config, pin, codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);
    assertAdmin_(pin);

    const sh = getSheet_(SHEET_LISTAS);

    const rawPrecio1 = String(config?.precio1 ?? "").trim();
    const rawPrecio2 = String(config?.precio2 ?? "").trim();
    const promoActiva = !!config?.promoActiva;

    sh.getRange("D2").setValue("1 número");
    sh.getRange("D3").setValue("2 números");
    sh.getRange("D4").setValue("Promo 2x activa");

    sh.getRange("G1").setValue("Publicidad");
    sh.getRange("G2").setValue("Título");
    sh.getRange("G3").setValue("Premios");
    sh.getRange("G4").setValue("Contacto");
    sh.getRange("G5").setValue("Imagen fondo ID");
    sh.getRange("G6").setValue("Texto extra");
    sh.getRange("G7").setValue("Imagen fondo nombre");
    sh.getRange("G8").setValue("Mensaje inferior general");

    // Si el campo está vacío, NO pisa el precio anterior.
    if (rawPrecio1 !== "") {
      const precio1 = Number(rawPrecio1);
      if (!Number.isFinite(precio1) || precio1 < 0) {
        throw new Error("Precio 1 inválido.");
      }
      sh.getRange("E2").setValue(precio1);
    }

    // Si el campo está vacío, NO pisa el precio anterior.
    if (rawPrecio2 !== "") {
      const precio2 = Number(rawPrecio2);
      if (!Number.isFinite(precio2) || precio2 < 0) {
        throw new Error("Precio promo inválido.");
      }
      sh.getRange("E3").setValue(precio2);
    }

    // La promo sí se guarda aunque no cambies precios.
    sh.getRange("E4").setValue(promoActiva ? "SI" : "NO");

    if (config) {
      const pub = config.publicidad || {
        titulo: config.pubTitulo || "",
        premios: config.pubPremios || "",
        contacto: config.pubContacto || "",
        textoExtra: config.pubExtra || "",
        mensajeInferior: config.pubFooterMsg || ""
      };

      sh.getRange("H2").setValue(String(pub.titulo || "").trim());
      sh.getRange("H3").setValue(String(pub.premios || "").trim());
      sh.getRange("H4").setValue(String(pub.contacto || "").trim());
      // H5/H7 se usan para la imagen de fondo fija, guardada por apiAdminUploadPublicityBackground.
      sh.getRange("H6").setValue(String(pub.textoExtra || "").trim());
      sh.getRange("H8").setValue(String(pub.mensajeInferior || pub.footerMsg || "").trim());
    }

    return ok_("OK");

  } catch (e) {
    return fail_(e);
  }
}



/** ========= PUBLICIDAD DRIVE ========= */
function getOrCreatePublicityRootFolder_() {
  // Compatibilidad con versiones anteriores.
  // Ahora los archivos se guardan dentro de la carpeta de cada rifa.
  return getOrCreateRifasRootFolder_();
}

function getOrCreatePublicityRifaFolder_() {
  return getOrCreateRifaSubFolder_("Publicidad");
}

function apiAdminUploadPublicityBackground(dataUrl, fileName, pin, codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);
    assertAdmin_(pin);

    dataUrl = String(dataUrl || "");
    fileName = String(fileName || "fondo_publicidad.png").trim() || "fondo_publicidad.png";

    const match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (!match) throw new Error("Formato de imagen inválido.");

    const mime = match[1];
    const b64 = match[2];
    const bytes = Utilities.base64Decode(b64);

    if (bytes.length > 8 * 1024 * 1024) {
      throw new Error("La imagen es demasiado grande. Probá con una imagen menor a 8 MB.");
    }

    const safeName = "fondo_publicidad_" + normalizeCode_(CURRENT_RIFA_CODE) + "_" + new Date().getTime() + "_" + fileName;
    const blob = Utilities.newBlob(bytes, mime, safeName);

    const sh = getSheet_(SHEET_LISTAS);
    const oldFileId = String(sh.getRange("H5").getValue() || "").trim();

    // Guardamos la imagen dentro de la carpeta de la rifa:
    // RIFITA / Rifa - Nombre (codigo) / Publicidad
    const folder = getOrCreatePublicityRifaFolder_();
    const file = folder.createFile(blob);

    sh.getRange("G5").setValue("Imagen fondo ID");
    sh.getRange("G7").setValue("Imagen fondo nombre");
    sh.getRange("H5").setValue(file.getId());
    sh.getRange("H7").setValue(fileName);

    // Para no llenar Drive de fondos viejos, mandamos a papelera el anterior.
    if (oldFileId && oldFileId !== file.getId()) {
      try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (_) {}
    }

    return ok_({
      fileId: file.getId(),
      name: fileName,
      folderName: folder.getName(),
      folderUrl: folder.getUrl()
    });

  } catch (e) {
    return fail_(e);
  }
}

function apiGetPublicityBackgroundData(codigoRifa) {
  try {
    setCurrentRifa_(codigoRifa);
    const sh = getSheet_(SHEET_LISTAS);
    const fileId = String(sh.getRange("H5").getValue() || "").trim();

    if (!fileId) {
      return ok_({ dataUrl: "" });
    }

    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const mime = blob.getContentType() || "image/png";
    const b64 = Utilities.base64Encode(blob.getBytes());

    return ok_({
      dataUrl: "data:" + mime + ";base64," + b64,
      name: file.getName(),
      fileId: fileId
    });

  } catch (e) {
    return fail_(e);
  }
}

function ensurePrecioColumn_() {
  const sh = getVentasSheetSmart_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  if (!headers.includes("PrecioCobrado")) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue("PrecioCobrado");
  }
}

function probarApiInitFastClub() {
  const res = apiInitFast("club2026");
  Logger.log(JSON.stringify(res, null, 2));
}

function probarLoginClub() {
  const res = apiAdminLogin("1234", "club2026");
  Logger.log(JSON.stringify(res, null, 2));
}

function DEBUG_bomberos(){
  const codigo = "bomberos";

  try {
    const cfg = getRifaConfig_(codigo);
    Logger.log("CONFIG OK");
    Logger.log(JSON.stringify(cfg));

    setCurrentRifa_(codigo);

    const ss = SpreadsheetApp.openById(cfg.sheetId);
    Logger.log("SHEET ABIERTO OK: " + ss.getName());

    Logger.log("Tiene LISTAS: " + !!ss.getSheetByName(SHEET_LISTAS));
    Logger.log("Tiene VENTAS: " + !!ss.getSheetByName(SHEET_VENTAS_PRIMARY));
    Logger.log("Tiene GANADORES: " + !!ss.getSheetByName(SHEET_VENTAS_FALLBACK));

    const init = apiInitFast(codigo);
    Logger.log("INIT:");
    Logger.log(JSON.stringify(init));

    return init;

  } catch(e) {
    Logger.log("ERROR BOMBEROS: " + e.message);
    return fail_(e);
  }
}


/** ========= NETLIFY BRIDGE =========
 * Permite que el frontend alojado en Netlify use estas funciones sin google.script.run.
 * Netlify llama por GET con action="__call", method y args. Esto evita el 401 que aparece en algunas cuentas al hacer POST anónimo a Apps Script.
 */
function apiNetlifyCall_(method, args, codigoRifa) {
  try {
    method = String(method || "").trim();
    args = Array.isArray(args) ? args : [];

    const allowed = {
      apiPing: apiPing,
      apiCheckRifa: apiCheckRifa,
      apiInitFast: apiInitFast,
      apiGetDashboard: apiGetDashboard,
      apiGetConfig: apiGetConfig,
      apiGetSale: apiGetSale,
      apiGetSellerSummary: apiGetSellerSummary,
      apiGetWinners: apiGetWinners,
      apiGetPublicityBackgroundData: apiGetPublicityBackgroundData,

      confirmSale: confirmSale,
      getFormData: getFormData,
      getSoldNumbers: getSoldNumbers,
      getSaleByNumber: getSaleByNumber,
      deleteSaleByNumber: deleteSaleByNumber,

      apiAdminLogin: apiAdminLogin,
      apiAdminDeleteSale: apiAdminDeleteSale,
      apiAdminDrawWinners: apiAdminDrawWinners,
      apiAdminResetWinners: apiAdminResetWinners,
      apiAdminCreateSoldNumbersPdf: apiAdminCreateSoldNumbersPdf,
      apiAdminCreateSellerPdf: apiAdminCreateSellerPdf,
      apiAdminSavePrices: apiAdminSavePrices,
      apiAdminSaveConfig: apiAdminSaveConfig,
      apiAdminSaveRange: apiAdminSaveRange,
      apiAdminAddSeller: apiAdminAddSeller,
      apiAdminDeleteSeller: apiAdminDeleteSeller,
      apiAdminUploadPublicityBackground: apiAdminUploadPublicityBackground
    };

    if (!allowed[method]) {
      throw new Error("Función no permitida desde Netlify: " + method);
    }

    // En Apps Script las variables globales no son confiables entre llamadas.
    // Por eso fijamos la rifa antes de ejecutar métodos que no reciben codigoRifa.
    if (codigoRifa) {
      try { setCurrentRifa_(codigoRifa); } catch (_) {}
    }

    return allowed[method].apply(null, args);

  } catch (e) {
    return fail_(e);
  }
}
