/********************* CONFIGURACIÓN *********************/
const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/198LrnmGHLFAaFQrhdc4SVJZKOPQ7m568LQPo1sPgo4E/edit#gid=914468775";
const SHEET_NAME      = "Sheet1";

const PROCESSED_LABEL    = "processed-inmuebles24";
const EXCLUDE_TEST_LABEL = "LEADS_TEST";

// 👇 Alias “Enviar como”
const REPLY_FROM   = "contacto@industrialestatemexico.com";
const REPLY_NAME   = "Industrial Estate Mexico";

// Sólo enviar si el alias existe
const ENFORCE_ALIAS = true;

const INCLUDE_NAME_IN_GREETING = true;
let DRY_RUN = false;  // producción

/********************* MENÚ EN LA HOJA *********************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Leads")
    .addItem("Procesar nuevos ahora", "processNewLeads")
    .addItem((DRY_RUN ? "Activar envío" : "Desactivar envío"), "toggleSendMode_")
    .addSeparator()
    .addItem("Procesar a Inactivos", "procesarMovimientos")
    .addItem("Aplicar fórmulas y referencias (Selección segura)", "aplicarFormulasYRefs_SeleccionSegura")
    .addToUi();
}

function toggleSendMode_(){
  DRY_RUN = !DRY_RUN;
  SpreadsheetApp.getActive().toast("Envío automático: " + (DRY_RUN ? "DESACTIVADO" : "ACTIVADO"));
}

/********************* FUNCIÓN DE VERIFICACIÓN DE ALIAS *********************/
function checkAliasReady_() {
  const aliases = GmailApp.getAliases();
  const hasAlias = aliases.includes(REPLY_FROM);

  Logger.log("Alias detectados: " + aliases.join(", "));

  if (!hasAlias) {
    const msg = `⚠️ El alias "${REPLY_FROM}" no está disponible actualmente.\n
    Ve a Gmail → Configuración → Cuentas e importación → “Enviar como” y verifica que esté configurado.`;
    SpreadsheetApp.getActive().toast(msg);
    throw new Error(msg);
  }
  return true;
}

/********************* PROCESO PRINCIPAL *********************/
function processNewLeads(){
  // ✅ Verifica alias de envío antes de empezar
  if (ENFORCE_ALIAS) checkAliasReady_();

  const labelProcessed = getOrCreateLabel_(PROCESSED_LABEL);

  const SEARCH_QUERY =
    'in:anywhere ' +
    '-label:' + PROCESSED_LABEL + ' ' +
    '-label:' + EXCLUDE_TEST_LABEL + ' ' +
    'newer_than:30d ' +
    '(' +
      'from:@usuarios.inmuebles24.com OR from:@usuarios.vivanuncios.com.mx OR from:notificaciones@propiedades.com OR from:@inbox.easybroker.com' +
    ')';

  const threads = GmailApp.search(SEARCH_QUERY, 0, 150);
  Logger.log("Encontrados: " + threads.length);
  if (!threads.length){ SpreadsheetApp.getActive().toast("No hay correos nuevos."); return; }

  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`No existe la hoja ${SHEET_NAME}`);

  const header = getHeader_(sh);
  const headerLen = header.length;
  const IDX = {
    RECEIVED: colIndexByName_(header, "Fecha inicial"),
    NEXT_FUP: colIndexByName_(header, "Next Fup"),
    PHONE:    colIndexByName_(header, "Teléfono"),
    NAME:     colIndexByName_(header, "Nombre"),
    EMAIL:    colIndexByName_(header, "Email"),
    SOURCE:   colIndexByName_(header, "Fuente"),
    TITLE:    colIndexByName_(header, "Propiedad inicial"),
    LINK:     colIndexByName_(header, "Publicación inicial")
  };

  const rows   = [];
  const toSend = [];

  threads.forEach(thread => {
    const msg = thread.getMessages().slice(-1)[0];
    const from      = (msg.getFrom() || "").toLowerCase();
    const subject   = (msg.getSubject() || "").toLowerCase();
    const bodyText  = msg.getPlainBody() || "";
    const bodyHtml  = msg.getBody() || "";
    const lastDate  = thread.getLastMessageDate();

    let prov = detectProviderQuick_(from, subject);
    if (!prov) prov = detectProviderFromText_(bodyText);
    if (!prov){ thread.addLabel(labelProcessed); return; }

    const fuente   = sourceLabel_(prov);
    const received = lastDate;
    const receivedPlus = addDays_(received, 3);

    const name     = extractName_(bodyText, bodyHtml) || "";
    const emailTo  = extractEmail_(bodyText, bodyHtml) || "";
    const phoneStr = extractPhonesSmart_(bodyText, bodyHtml).join(" / ");
    const link     = getListingLinkByProvider_(prov, bodyHtml, bodyText) || "";
    const title    = buildTitleGeneric_(bodyHtml || bodyText) || "";

    const row = new Array(headerLen).fill("");
    row[IDX.RECEIVED] = received;
    row[IDX.NEXT_FUP] = receivedPlus;
    row[IDX.PHONE]    = phoneStr;
    row[IDX.NAME]     = name;
    row[IDX.EMAIL]    = emailTo;
    row[IDX.SOURCE]   = fuente;
    row[IDX.TITLE]    = title;
    row[IDX.LINK]     = link;

    rows.push(row);
    if (emailTo) toSend.push({ emailTo, name });

    thread.addLabel(labelProcessed);
  });

  if (rows.length){
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, rows.length, headerLen).setValues(rows);

    const columnas = computeColumnasFormulas_(sh);
    if (columnas.length){
      copiarFormulasFormatosYValidaciones_(sh, startRow, rows.length, columnas);
    }
  }

  if (!DRY_RUN && toSend.length){
    toSend.forEach(({emailTo, name}) => sendAutoReply_(emailTo, name, true));
  }

  SpreadsheetApp.getActive().toast(`Procesados: ${rows.length}.`);
}

/********************* FUNCIONES AUXILIARES *********************/
function getHeader_(sh){
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
           .map(h => String(h || "").trim());
}
function colIndexByName_(header, name){
  return header.indexOf(String(name || "").trim());
}
function addDays_(d, days){ const x = new Date(d.getTime()); x.setDate(x.getDate() + days); return x; }
function getOrCreateLabel_(name){ return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name); }

/********************* DETECCIÓN DE PROVEEDOR *********************/
function detectProviderQuick_(from, subject){
  if (/@usuarios\.inmuebles24\.com/.test(from))     return "i24";
  if (/@usuarios\.vivanuncios\.com\.mx/.test(from)) return "vivanuncios";
  if (/notificaciones@propiedades\.com/.test(from)) return "propiedades";
  if (/@inbox\.easybroker\.com/.test(from)){
    if (/mercado libre/i.test(subject)) return "mercado_libre";
    if (/bolsa/i.test(subject)) return "easy_bolsa";
    if (/telefono/i.test(subject)) return "easy_telefono";
  }
  return null;
}
function detectProviderFromText_(content){
  const h = (content || "").toLowerCase();
  if (h.includes("inmuebles24")) return "i24";
  if (h.includes("vivanuncios")) return "vivanuncios";
  if (h.includes("propiedades.com")) return "propiedades";
  if (h.includes("easybroker")){
    if (h.includes("mercado libre")) return "mercado_libre";
    if (h.includes("bolsa")) return "easy_bolsa";
  }
  return null;
}
function sourceLabel_(prov){
  switch(prov){
    case "i24": return "Inmuebles24";
    case "vivanuncios": return "VivaAnuncios";
    case "propiedades": return "Propiedades.com";
    case "mercado_libre": return "Mercado Libre";
    case "easy_bolsa": return "Bolsa EasyBroker";
    case "easy_telefono": return "EasyBroker (Teléfono)";
    default: return "";
  }
}

/********************* LINKS Y TÍTULOS *********************/
function getListingLinkByProvider_(prov, html, text){
  if (prov === "i24" || prov === "vivanuncios") return getListingLink_I24_(html, text);
  if (prov === "propiedades") return getListingLink_Propiedades_(html);
  if (prov === "easy_bolsa" || prov === "mercado_libre" || prov === "easy_telefono") return getListingLink_EasyBroker_(html, text);
  return "";
}
function getListingLink_I24_(html, plainText){
  const h = html || "";
  const t = plainText || "";
  const direct = findEasyBrokerLinkInHtml_(h);
  if (direct) return direct;
  const ebCode = extractEasyBrokerCode_(h) || extractEasyBrokerCode_(t);
  if (ebCode) return "https://www.easybroker.com/property/" + ebCode;
  const i24Match = h.match(/<a[^>]+href="([^"]+inmuebles24\.com[^"]+)"[^>]*>/i);
  if (i24Match && i24Match[1]) return i24Match[1];
  return "";
}
function getListingLink_Propiedades_(html){
  const h = html || "";
  const reVerProp = /<a[^>]*href="([^"]+)"[^>]*>.*ver\s+propiedad.*<\/a>/i;
  let m = h.match(reVerProp);
  if (m && m[1] && !/^mailto:/i.test(m[1])) return m[1];
  const hrefRe = /href="([^"]+)"/gi; let r;
  while ((r = hrefRe.exec(h)) !== null){
    const url = (r[1] || "").trim();
    if (/^mailto:/i.test(url)) continue;
    if (/^https?:\/\//i.test(url) && url.includes("propiedades.com")) return url;
  }
  return "";
}
function getListingLink_EasyBroker_(html, text){
  const h = html || "";
  const t = text || "";
  const ebCode = extractEasyBrokerCode_(h) || extractEasyBrokerCode_(t);
  if (ebCode) return "https://www.easybroker.com/property/" + ebCode;
  const re = /href="([^"]+easybroker\.com[^"]+)"/gi;
  let r;
  while ((r = re.exec(h)) !== null){
    const u = r[1];
    if (/\/(property|properties|listing|listings)\//i.test(u)) return u;
  }
  return "";
}
function findEasyBrokerLinkInHtml_(html){
  const re = /href="([^"]+)"/gi;
  let m;
  while ((m = re.exec(html)) !== null){
    const url = (m[1] || "").trim();
    if (/^mailto:/i.test(url)) continue;
    if (/^https?:\/\/(?:www\.)?easybroker\.com\//i.test(url)) return url;
  }
  return "";
}
function extractEasyBrokerCode_(text){
  if (!text) return "";
  const m = text.match(/\bEB-[A-Z0-9]{4,}\b/i);
  return m ? m[0].toUpperCase() : "";
}
function buildTitleGeneric_(textOrHtml){
  const t = stripHtml_(textOrHtml);
  let mPrice = /(MN|US\$|USD|MXN)\s[\d.,]+/i.exec(t);
  if (!mPrice){
    const line = (t.match(/([A-ZÁÉÍÓÚÑ].{30,120})/i) || [])[1];
    return (line || t).slice(0, 140).trim();
  }
  const start = mPrice.index;
  const after = t.slice(start);
  const m2 = /(\d[\d.,]*)\s*(m2|m²)/i.exec(after);
  if (m2){
    const frag = after.slice(0, m2.index).trim().replace(/\s{2,}/g, " ");
    return `${frag} ${m2[0].replace(/\s+/g, " ")}`.trim();
  }
  return after.slice(0, 140).trim();
}

/********************* EXTRACCIÓN DE DATOS *********************/
function extractName_(bodyText, bodyHtml){
  const plain = (bodyText || "") + "\n" + stripHtml_(bodyHtml || "");
  let m = plain.match(/Nombre y apellido:\s*([^\n\r]+)/i) || plain.match(/Nombre:\s*([^\n\r]+)/i);
  if (m) return sanitizeName_(m[1]);
  return "";
}
function sanitizeName_(s){ return String(s || "").replace(/\s+/g," ").replace(/[<>]/g,"").trim(); }

function extractEmail_(bodyText, bodyHtml){
  const mail = getFirstMatch_(bodyHtml, /href="mailto:([^"]+)"/i);
  if (mail && !isSystemEmail_(mail)) return mail;
  const all = (bodyText + " " + stripHtml_(bodyHtml)).match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [];
  const good = all.find(e => !isSystemEmail_(e));
  return good || "";
}
function isSystemEmail_(e){
  return /@usuarios\.inmuebles24\.com$/i.test(e) ||
         /@usuarios\.vivanuncios\.com\.mx$/i.test(e) ||
         /@propiedades\.com$/i.test(e) ||
         /@easybroker\.com$/i.test(e);
}

function extractPhonesSmart_(bodyText, bodyHtml){
  const plain = (bodyText || "") + "\n" + stripHtml_(bodyHtml || "");
  const re  = /(\+?\d[\d\s().-]{7,}\d)/g;
  const set = new Set();
  (plain.match(re) || []).forEach(x => {
    let digits = x.replace(/\D/g, "");
    if (digits.startsWith("52") && digits.length > 10) digits = digits.slice(2);
    if (digits.length > 10) digits = digits.slice(-10);
    if (digits.length === 10) set.add(digits);
  });
  return Array.from(set);
}

function getFirstMatch_(text, regex){ const m = (text || "").match(regex); return m ? m[1] : ""; }
function stripHtml_(html){
  return (html || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}

/********************* CORREO AUTOMÁTICO *********************/
function sendAutoReply_(emailTo, name, hasAlias){
  const subject  = "Solicitud recibida para bodegas y terrenos industriales | Industrial Estate México";
  const nombre = (INCLUDE_NAME_IN_GREETING && name) ? " " + sanitizeName_(name) : "";
  const htmlBody =
    `<p>Hola${nombre},</p>
<p>Recibimos tu solicitud relacionada con una nave industrial publicada.</p>
<p>En <strong>Industrial Estate Mexico</strong> tenemos más de 1,000 propiedades industriales disponibles. Si puedes indicarme ubicación, superficie y tiempos estimados, te enviaré las opciones más relevantes.</p>
<p>Visita nuestros sitios:</p>
<ul>
<li><a href="https://industrialestatemexico.com">industrialestatemexico.com</a></li>
<li><a href="https://www.tierraindustrial.mx/">tierraindustrial.mx</a></li>
</ul>
<p>Saludos,<br><strong>Sebastián Dérieux</strong><br>Industrial Estate Mexico</p>`;

  const opts = {
    htmlBody: htmlBody,
    name: REPLY_NAME,
    replyTo: REPLY_FROM,
    noReply: false
  };
  if (hasAlias) opts.from = REPLY_FROM;

  GmailApp.sendEmail(emailTo, subject, "", opts);
}
function _norm_(s){
  return String(s||"")
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')  // quita acentos
    .trim();
}

/********************* FORMATO Y FÓRMULAS *********************/
function computeColumnasFormulas_(hoja){
  const header = hoja.getRange(1,1,1,hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
  const wanted = [
    "Titulo Lead","Activo","Días para F.Up","WA","wa.me","FUP WA","First name",
    "Priority","Antiguedad Score","Directo Score","Superficie Score",
    "Cercanía Cierre Score","Score Total"
  ];
  return wanted.map(name => header.indexOf(name)+1).filter(i => i>0);
}


function copiarFormulasFormatosYValidaciones_(hoja, filaInicio, numFilas, columnas){
  if (!columnas.length) return;
  for (let i = 0; i < numFilas; i++) {
    const r = filaInicio + i;
    const base = 2;  // ✅ siempre copia desde la fila 2
    hoja.getRange(base, 1, 1, hoja.getLastColumn())
        .copyTo(hoja.getRange(r, 1, 1, hoja.getLastColumn()), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    hoja.getRange(base, 1, 1, hoja.getLastColumn())
        .copyTo(hoja.getRange(r, 1, 1, hoja.getLastColumn()), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    columnas.forEach(c => {
      const f = hoja.getRange(base, c).getFormulaR1C1();
      if (f && f.startsWith("=")) hoja.getRange(r, c).setFormulaR1C1(f);
    });
  }
}

/*********** AUTO-MOVER CON MENÚ ***********/ 
const MASTER_SHEET_NAME   = "Sheet1";     
const INACTIVE_SHEET_NAME = "Inactivos";  

function procesarMovimientos() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh1 = ss.getSheetByName(MASTER_SHEET_NAME);
  const sh2 = ss.getSheetByName(INACTIVE_SHEET_NAME);

  mover_(sh1, sh2, ["I"]);
  mover_(sh2, sh1, ["A", "SC"]);

  try {
    SpreadsheetApp.getUi().alert("✅ Movimientos procesados correctamente");
  } catch (e) {
    Logger.log("✅ Movimientos procesados correctamente (sin UI)");
  }
}

function mover_(fromSh, toSh, valoresPermitidos) {
  const header = fromSh.getRange(1,1,1,fromSh.getLastColumn()).getValues()[0].map(String);
  const colActivo = header.map(h => h.trim().toLowerCase()).indexOf("activo") + 1;
  if (colActivo <= 0) return;

  const lastRow = fromSh.getLastRow();
  if (lastRow <= 1) return;

  const data = fromSh.getRange(2,1,lastRow-1,fromSh.getLastColumn()).getValues();
  const rowsToMove = [];

  data.forEach((rowData, idx) => {
    const valor = String(rowData[colActivo - 1] || "").trim().toUpperCase();
    if (valoresPermitidos.includes(valor)) {
      rowsToMove.push(idx + 2);
    }
  });

  if (!rowsToMove.length) return;

  rowsToMove.sort((a,b) => b - a);

  rowsToMove.forEach(r => {
    const lastColFrom = fromSh.getLastColumn();
    const lastColTo   = toSh.getLastColumn();
    const cols = Math.min(lastColFrom, lastColTo);
    const nextRow = toSh.getLastRow() + 1;

    fromSh.getRange(r, 1, 1, cols).copyTo(
      toSh.getRange(nextRow, 1, 1, cols),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

    toSh.setRowHeight(nextRow, fromSh.getRowHeight(r));
    fromSh.deleteRow(r);
  });
}

function aplicarFormulasYRefs_SeleccionSegura() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rango = hoja.getActiveRange();
  if (!rango) {
    SpreadsheetApp.getActive().toast("Selecciona las filas a procesar.");
    return;
  }

  const filas = rango.getNumRows();
  const filaInicio = rango.getRow();
  const hojaNombre = hoja.getName();

  const header = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0]
    .map(x => String(x || "").trim());

    const columnasFormulas = header.reduce((arr, h, i) => {
    if ([
      "Titulo Lead","Activo","WA","wa.me","FUP WA","Mail Fup","Priority","First name",
      "Antiguedad Score","Directo Score","Superficie Score","Cercanía Cierre Score","Score Total"
    ].includes(h)) {
      arr.push(i + 1);
    }
    return arr;
  }, []);

  for (let i = 0; i < filas; i++) {
    const r = filaInicio + i;
    if (r <= 1) continue;

    const filaBase = 2;  // ✅ siempre copiar desde fila 2

    hoja.getRange(filaBase, 1, 1, hoja.getLastColumn()).copyTo(
      hoja.getRange(r, 1, 1, hoja.getLastColumn()),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
    hoja.getRange(filaBase, 1, 1, hoja.getLastColumn()).copyTo(
      hoja.getRange(r, 1, 1, hoja.getLastColumn()),
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
      false
    );

    columnasFormulas.forEach(c => {
      const f = hoja.getRange(filaBase, c).getFormulaR1C1();
      if (f && f.startsWith("="))
        hoja.getRange(r, c).setFormulaR1C1(f);
    });
  }

  SpreadsheetApp.getActive().toast("✅ Fórmulas y formatos aplicados correctamente");
}
