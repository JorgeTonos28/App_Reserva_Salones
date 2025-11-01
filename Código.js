/**
 * *********************************
 * App – Reserva de Salones (v1)
 * *********************************
 * Proyecto: App – Reserva de Salones (Google Workspace)
 * Autoría: Jorge Tonos
 * Cargo: Enc. Centros de Incubación, Dirección de Innovación y Análisis de Datos
 * Institución/Año: INFOTEP, 2029
 * Mantenimiento: Equipo de TI interno / colaboradores designados
 * Descripción:
 *   Backend principal en Google Apps Script que gestiona el flujo de reservas
 *   para espacios y salones corporativos. Orquesta la interfaz web (HTML
 *   incluidos en el proyecto), la lógica de validaciones, la integración con
 *   Hojas de cálculo y el envío de notificaciones por correo.
 * Módulos clave:
 *   - Control de acceso por roles y estados (Usuarios, Admin, Solicitante).
 *   - Gestión de salones, horarios y disponibilidad mediante la hoja "Reservas".
 *   - Servicios auxiliares para logos almacenados en Drive y plantillas HTML.
 * Consideraciones de despliegue:
 *   - Publicar el Web App como aplicación ejecutándose como propietario y
 *     accesible solo para los usuarios del dominio deseado.
 *   - Revisar los permisos OAuth solicitados tras cada ajuste de alcance.
 * Historial:
 *   - Versión original Apps Script v1 migrada y extendida sucesivamente.
 *
 * NOTAS DE CONFIGURACIÓN DEL PROYECTO:
 * 1) Zona horaria del script: definir en File > Project properties > Script properties
 *    -> Time zone = America/Santo_Domingo (evita desfases de fecha/hora).
 * 2) Drive API avanzada (para miniaturas de logo):
 *    -> Services (puzzle icon) > Enable Advanced Google services...
 *    -> Activar "Drive API"
 *    -> En Google Cloud Console del proyecto, habilitar la API "Drive" y autorizar scopes.
 *    Si no se habilita, las funciones getDriveThumbnailBlob_/getLogoDataUrl retornarán vacío (fallback inofensivo).
 *    Recomendado para emails: usar una URL pública HTTPS en Config -> MAIL_LOGO_URL
 *    (evita data-URIs bloqueadas por algunos clientes de correo).
 */
const SS = SpreadsheetApp.getActive();
const SH_CFG = SS.getSheetByName('Config');
const SH_USR = SS.getSheetByName('Usuarios');
const SH_CON = SS.getSheetByName('Conserjes');
const SH_SAL = SS.getSheetByName('Salones');
const SH_RES = SS.getSheetByName('Reservas');
const APP_VERSION = 'salones-v10.1-2025-11-06';

// ========= Helpers de tiempo =========
function nowStr_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
function ymd_(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function toDate_(s){ const m=String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})/); return m? new Date(+m[1],+m[2]-1,+m[3]) : new Date(s); }
function tsCell_(v){
  // Devuelve 'YYYY-MM-DD HH:mm:ss' si es Date; si no, deja string plano
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)){
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(v || '');
}
function ymdCell_(v){
  try{
    // Convierte cualquier valor de celda (Date, string, número) a 'YYYY-MM-DD'
    return ymd_(new Date(v));
  }catch(e){
    // Fallback si fuese texto ya en formato YMD
    return String(v||'').slice(0,10);
  }
}
function hhmmFromCell_(v){
  // 1) Si es Date (incluye el caso "solo hora" anclado a 1899-12-30)
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)){
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  }
  // 2) Si es número (tiempo como fracción del día)
  if (typeof v === 'number' && isFinite(v)){
    var totalMin = Math.round((v % 1) * 24 * 60);
    var h = Math.floor(totalMin / 60), m = totalMin % 60;
    return Utilities.formatString('%02d:%02d', h, m);
  }
  // 3) Si es texto con horas tipo H:MM o HH:MM
  var s = String(v || '').trim();
  var m = s.match(/(\d{1,2}):(\d{1,2})/);
  if (m){
    var h = String(Math.max(0, Math.min(23, Number(m[1])))).padStart(2, '0');
    var mi = String(Math.max(0, Math.min(59, Number(m[2])))).padStart(2, '0');
    return h + ':' + mi;
  }
  return '';
}
function toMin_(hhmm){
  const m = String(hhmm || '').trim().match(/^(\d{1,2}):(\d{1,2})$/);
  if (!m) return NaN;
  const h = Number(m[1]), mi = Number(m[2]);
  if (!isFinite(h) || !isFinite(mi) || h<0 || h>23 || mi<0 || mi>59) return NaN;
  return h*60 + mi;
}
function normHHMM_(s){
  const m = String(s || '').trim().match(/^(\d{1,2}):(\d{1,2})$/);
  if (!m) return '';
  const h  = String(Math.max(0, Math.min(23, Number(m[1])))).padStart(2, '0');
  const mi = String(Math.max(0, Math.min(59, Number(m[2])))).padStart(2, '0');
  return h + ':' + mi;
}
function toHHMM_(mins){
  mins = Math.max(0, Math.floor(mins));
  var h = Math.floor(mins/60), m = mins%60;
  return Utilities.formatString('%02d:%02d', h, m);
}
function addMinutes_(hhmm, add){
  return toHHMM_(toMin_(hhmm) + Number(add||0));
}

// ========= Helpers de formato (para emails) =========
function fmtDMY_(ymdOrDate){
  // Entrada: 'YYYY-MM-DD' o Date; Salida: 'DD/MM/YYYY'
  try{
    var d = (Object.prototype.toString.call(ymdOrDate) === '[object Date]') ? ymdOrDate : toDate_(ymdOrDate);
    var dd = String(d.getDate()).padStart(2,'0');
    var mm = String(d.getMonth()+1).padStart(2,'0');
    var yy = d.getFullYear();
    return dd + '/' + mm + '/' + yy;
  }catch(e){
    var s = String(ymdOrDate||'').slice(0,10);
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? (m[3]+'/'+m[2]+'/'+m[1]) : s;
  }
}
function fmt12_(hhmm){
  // Entrada: 'HH:mm'; Salida: 'hh:mm a.m.' / 'hh:mm p.m.'
  var m = String(hhmm||'').match(/^(\d{2}):(\d{2})$/);
  if (!m) return String(hhmm||'');
  var H = Number(m[1]), Mi = Number(m[2]);
  var am = H < 12;
  var h12 = (H % 12) || 12;
  return Utilities.formatString('%02d:%02d %s', h12, Mi, am?'a.m.':'p.m.');
}

// ========= Config =========
function cfg_(key){
  if (!SH_CFG) return '';
  const lr = SH_CFG.getLastRow();
  if (lr<2) return '';
  const vals = SH_CFG.getRange(2,1,lr-1,2).getValues();
  const r = vals.find(v => String(v[0]).trim()===key);
  return r ? String(r[1]).trim() : '';
}

const ADMIN_ALT_CFG_KEYS = Object.freeze({
  ADMIN_EMAILS:true,
  HORARIO_INICIO:true,
  HORARIO_FIN:true,
  DURATION_MIN:true,
  DURATION_STEP:true,
  DURATION_MAX:true,
  MAIL_SENDER_NAME:true,
  MAIL_REPLY_TO:true,
  ADMIN_CONTACT_NAME:true,
  ADMIN_CONTACT_EMAIL:true,
  ADMIN_CONTACT_EXTENSION:true
});

const __ADMIN_CFG_CACHE__ = {};
const __ADMIN_RUNTIME_CACHE__ = {};

function normalizeAdminId_(id){
  const raw = String(id || '').trim();
  if (!raw) return '1';
  const num = Number(raw);
  if (isFinite(num) && num > 0){
    return String(Math.floor(num));
  }
  return raw;
}

function cfgSheetLookup_(sheetName, key){
  const cacheKey = sheetName + '::' + key;
  if (cacheKey in __ADMIN_CFG_CACHE__) return __ADMIN_CFG_CACHE__[cacheKey];
  const sh = SS.getSheetByName(sheetName);
  if (!sh){
    __ADMIN_CFG_CACHE__[cacheKey] = '';
    return '';
  }
  const lr = sh.getLastRow();
  if (lr < 2){
    __ADMIN_CFG_CACHE__[cacheKey] = '';
    return '';
  }
  const vals = sh.getRange(2,1,lr-1,2).getValues();
  const row = vals.find(v => String(v[0]).trim() === key);
  const val = row ? String(row[1]).trim() : '';
  __ADMIN_CFG_CACHE__[cacheKey] = val;
  return val;
}

function adminConfigSheetName_(adminId){
  const id = normalizeAdminId_(adminId);
  if (id === '1') return 'Config';
  return 'Config' + id;
}

function cfgForAdmin_(adminId, key){
  const normKey = String(key || '').trim();
  if (!normKey) return '';
  const id = normalizeAdminId_(adminId);
  if (!ADMIN_ALT_CFG_KEYS[normKey] || id === '1'){
    return cfg_(normKey);
  }
  const sheetName = adminConfigSheetName_(id);
  if (!sheetName) return cfg_(normKey);
  const raw = cfgSheetLookup_(sheetName, normKey);
  return raw !== '' ? raw : cfg_(normKey);
}

function getAdminRuntimeCfg_(adminId){
  const id = normalizeAdminId_(adminId);
  if (__ADMIN_RUNTIME_CACHE__[id]) return __ADMIN_RUNTIME_CACHE__[id];
  const horarioInicio = normHHMM_(cfgForAdmin_(id, 'HORARIO_INICIO') || '07:00') || '07:00';
  const horarioFin    = normHHMM_(cfgForAdmin_(id, 'HORARIO_FIN')    || '20:00') || '20:00';
  const durMin  = Number(cfgForAdmin_(id, 'DURATION_MIN')  || cfg_('DURATION_MIN')  || 30) || 30;
  const durMax  = Number(cfgForAdmin_(id, 'DURATION_MAX')  || cfg_('DURATION_MAX')  || 240) || 240;
  const durStep = Number(cfgForAdmin_(id, 'DURATION_STEP') || cfg_('DURATION_STEP') || 30) || 30;
  const cfgObj = {
    horarioInicio,
    horarioFin,
    duracionMin: Math.max(1, durMin),
    duracionMax: Math.max(Math.max(1, durMin), durMax),
    duracionStep: Math.max(1, durStep)
  };
  __ADMIN_RUNTIME_CACHE__[id] = cfgObj;
  return cfgObj;
}

function adminScopeForUser_(user){
  if (!user) return '1';
  const raw = user.administracion_id || user.admin_id || user.adminScope || '';
  return normalizeAdminId_(raw);
}

function isGeneralAdminUser_(user){
  if (!user) return false;
  if (String(user.rol||'').toUpperCase() !== 'ADMIN') return false;
  if (String(user.estado||'').toUpperCase() !== 'ACTIVO') return false;
  return adminScopeForUser_(user) === '1';
}

function isPublicoExternoOMixto_(tipo){
  const t = String(tipo||'').trim().toUpperCase();
  return t === 'EXTERNO' || t === 'MIXTO';
}

// ========= IDs / Hojas =========
function sh_(name){ const sh=SS.getSheetByName(name); if(!sh) throw new Error("No existe la pestaña '"+name+"'"); return sh; }
function nextId_(sh, prefix){
  const lr = sh.getLastRow();
  if (lr<2) return prefix+'00001';
  const last = String(sh.getRange(lr,1).getValue());
  const n = Number(last.replace(/\D/g,'')||0)+1;
  return prefix + String(n).padStart(5,'0');
}

// ========= Seguridad / Usuarios =========
function getUser_(){
  const email = (Session.getActiveUser().getEmail()||'').toLowerCase();
  const row = findUserByEmail_(email);
  if (row) {
    // Usuario encontrado en la hoja
    row._exists = true;
    if (!Array.isArray(row.prioridad_salones)){
      row.prioridad_salones = parsePrioridadSalones_(row.prioridad_salones_raw||'');
    }
    return row;
  }
  // Usuario NO existe en la hoja: devolvemos objeto “vacío” pero con bandera
  return {
    email,
    nombre:'',
    departamento:'',
    rol:'',
    prioridad:0,
    prioridad_salones:[],
    prioridad_salones_raw:'',
    estado:'',
    administracion_id:'1',
    _exists:false
  };
}
function findUserByEmail_(email){
  if (!email) return null;
  const lr = SH_USR.getLastRow();
  if (lr<2) return null;
  // lee columnas esperadas (incluye prioridad_salones, estado y extensión)
  const vals = SH_USR.getRange(2,1,lr-1,9).getValues();
  const r = vals.find(v => String(v[0]).trim().toLowerCase()===email.toLowerCase());
  if (!r) return null;
  const prioridadSalonesRaw = String(r[5]||'').trim();
  const adminId = normalizeAdminId_(r[8]||'1');
  return {
    email:r[0],
    nombre:r[1],
    departamento:r[2],
    rol:String(r[3]||'').toUpperCase(),
    prioridad:Number(r[4]||0),
    prioridad_salones: parsePrioridadSalones_(prioridadSalonesRaw),
    prioridad_salones_raw: prioridadSalonesRaw,
    estado:String(r[6]||'').toUpperCase(),       // NUEVO (PENDIENTE/ACTIVO/INACTIVO)
    extension:String(r[7]||''),             // NUEVO
    administracion_id: adminId
  };
}

function parsePrioridadSalones_(raw){
  const text = String(raw||'').trim();
  if (!text) return [];
  const seen = {};
  return text.split(';')
    .map(s => String(s||'').trim().toUpperCase())
    .filter(token => {
      if (!token) return false;
      if (seen[token]) return false;
      seen[token] = true;
      return true;
    });
}

function basePriorityForUser_(user){
  if (!user) return 0;
  const raw = Number(user && user.prioridad || 0);
  if (user && user.rol === 'ADMIN' && !Number(raw||0)){
    return 2;
  }
  return Number(raw||0);
}

function effectivePriorityForSalon_(user, salonId){
  const base = basePriorityForUser_(user);
  if (base <= 0) return base;
  if (!salonId) return base;
  const codes = Array.isArray(user && user.prioridad_salones)
    ? user.prioridad_salones
    : parsePrioridadSalones_(user && user.prioridad_salones_raw);
  if (!codes || !codes.length) return base;
  const target = String(salonId||'').trim().toUpperCase();
  return codes.includes(target) ? base : 0;
}

function isAdminEmail_(email){
  if (!email) return false;
  const u = findUserByEmail_(email);
  // Si el usuario existe y NO está ACTIVO → no es admin (aunque esté listado)
  if (u && String(u.estado||'').toUpperCase() !== 'ACTIVO') return false;
  if (u && u.rol==='ADMIN') return true;  // ADMIN + ACTIVO
  const cfgAdmins = (cfg_('ADMIN_EMAILS')||'').split(';').map(s=>s.trim().toLowerCase()).filter(Boolean);
  // Si viene por lista de config, también exigir ACTIVO en la hoja (si existe fila)
  if (cfgAdmins.includes(email.toLowerCase())){
    // Si no hay fila de usuario, lo tratamos como no activo
    return !!(u && String(u.estado||'').toUpperCase()==='ACTIVO');
  }
  return false;
}

// ========= Web (un solo enlace) =========
function doGet(e){
  const me = (Session.getActiveUser().getEmail()||'').toLowerCase();
  const u  = getUser_();

  // 1) Flujo de cancelación por link: ?cancel=TOKEN
  const cancelToken = (e && e.parameter && e.parameter.cancel) || '';
  if (cancelToken){
    const t = HtmlService.createTemplateFromFile('Cancel');
    t.logoUrl = getLogoDataUrl(128) || getLogoUrl();
    t.cancelToken = cancelToken;
    return t.evaluate()
      .setTitle('Cancelar reserva')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
      .addMetaTag('apple-mobile-web-app-capable','yes')
      .addMetaTag('mobile-web-app-capable','yes');
  }

  // ===== Gate por ESTADO (manda antes que ROL) =====
  const exists = !!(u && u._exists === true);       // 👈 existe en la hoja Usuarios
  const estado = String(u && u.estado || '').toUpperCase();

  // 2.a) Usuario NO existe en BD -> Denied con botón "Solicitar acceso"
  if (!exists){
    const t = HtmlService.createTemplateFromFile('Denied');
    t.logoUrl    = getLogoDataUrl(128) || getLogoUrl();
    t.appVersion = APP_VERSION;
    t.user = u || { email: me, nombre:'', departamento:'' };
    t.contact = {
      name: cfg_('ADMIN_CONTACT_NAME') || 'Administración',
      email: cfg_('ADMIN_CONTACT_EMAIL') || '',
      ext:   cfg_('ADMIN_CONTACT_EXTENSION') || ''
    };
    t.supportUrl = '';          // usamos modal interno
    t.isPending  = false;
    t.isInactive = false;
    t.noRole     = false;
    t.noState    = false;
    return t.evaluate()
      .setTitle('Acceso denegado – Reserva de Salones')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
      .addMetaTag('apple-mobile-web-app-capable','yes')
      .addMetaTag('mobile-web-app-capable','yes');
  }

  // 2.b) Existe en BD pero estado distinto de ACTIVO -> mensajes específicos
  //     (incluye PENDIENTE, INACTIVO y también "sin estado")
  if (estado !== 'ACTIVO'){
    const t = HtmlService.createTemplateFromFile('Denied');
    t.logoUrl    = getLogoDataUrl(128) || getLogoUrl();
    t.appVersion = APP_VERSION;
    t.user = u;
    t.contact = {
      name: cfg_('ADMIN_CONTACT_NAME') || 'Administración',
      email: cfg_('ADMIN_CONTACT_EMAIL') || '',
      ext:   cfg_('ADMIN_CONTACT_EXTENSION') || ''
    };
    t.supportUrl = '';
    t.isPending  = (estado === 'PENDIENTE');
    t.isInactive = (estado === 'INACTIVO');
    t.noRole     = false;
    t.noState    = (estado === '');   // << sin estado asignado
    return t.evaluate()
      .setTitle('Acceso denegado – Reserva de Salones')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
      .addMetaTag('apple-mobile-web-app-capable','yes')
      .addMetaTag('mobile-web-app-capable','yes');
  }

  // ===== A partir de aquí, el usuario está ACTIVO y existe en BD =====

  // 3) Admin
  if (isAdminEmail_(me)){
    const t = HtmlService.createTemplateFromFile('Admin');
    t.appVersion = APP_VERSION;
    t.logoUrl = getLogoDataUrl(128) || getLogoUrl();
    t.me = u;
    return t.evaluate()
      .setTitle('Reserva de Salones – Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
      .addMetaTag('apple-mobile-web-app-capable','yes')
      .addMetaTag('mobile-web-app-capable','yes');
  }

  // 4) Público (solicitante o admin)
  if (u && (u.rol==='SOLICITANTE' || u.rol==='ADMIN')){
    const t = HtmlService.createTemplateFromFile('Public');
    t.appVersion = APP_VERSION;
    t.logoUrl = getLogoDataUrl(128) || getLogoUrl();
    t.me = u;
    t.horarioInicio = cfg_('HORARIO_INICIO') || '07:00';
    t.horarioFin    = cfg_('HORARIO_FIN')    || '20:00';
    // Valores de duración para generar UI de manera dinámica
    t.cfg = {
      DUR_MIN:  Number(cfg_('DURATION_MIN')  || 30),
      DUR_MAX:  Number(cfg_('DURATION_MAX')  || 240),
      DUR_STEP: Number(cfg_('DURATION_STEP') || 30)
    };
    return t.evaluate()
      .setTitle('Reserva de Salones')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
      .addMetaTag('apple-mobile-web-app-capable','yes')
      .addMetaTag('mobile-web-app-capable','yes');
  }

  // 5) ACTIVO pero sin rol válido -> Denied con "Falta de permisos"
  const t = HtmlService.createTemplateFromFile('Denied');
  t.logoUrl    = getLogoDataUrl(128) || getLogoUrl();
  t.appVersion = APP_VERSION;
  t.user = u;
  t.contact = {
    name: cfg_('ADMIN_CONTACT_NAME') || 'Administración',
    email: cfg_('ADMIN_CONTACT_EMAIL') || '',
    ext:   cfg_('ADMIN_CONTACT_EXTENSION') || ''
  };
  t.supportUrl = '';
  t.isPending  = false;
  t.isInactive = false;
  t.noRole     = true;
  t.noState    = false;
  return t.evaluate()
    .setTitle('Acceso denegado – Reserva de Salones')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable','yes')
    .addMetaTag('mobile-web-app-capable','yes');
}

function include_(fn){ return HtmlService.createHtmlOutputFromFile(fn).getContent(); }

// ========= Salones =========
let __SALONES_CACHE__ = null;
let __SALONES_MAP__ = null;

function salonesCache_(){
  if (__SALONES_CACHE__) return __SALONES_CACHE__;
  const lr = SH_SAL.getLastRow();
  const rows = lr<2? [] : SH_SAL.getRange(2,1,lr-1,8).getValues();
  const data = rows.map(r=>{
    const adminId = normalizeAdminId_(r[6]||'1');
    const adminCfg = getAdminRuntimeCfg_(adminId);
    const requiereConserje = String(r[7]||'').toUpperCase()==='NO' ? 'NO' : 'SI';
    return {
      id: r[0],
      nombre: r[1],
      capacidad: Number(r[2]||0),
      habilitado: String(r[3]||'').toUpperCase()==='SI',
      sede: r[4]||'',
      restriccion: String(r[5]||'').trim(),
      administracion_id: adminId,
      requiere_conserje: requiereConserje,
      horario_inicio: adminCfg.horarioInicio,
      horario_fin: adminCfg.horarioFin,
      duracion_min: adminCfg.duracionMin,
      duracion_max: adminCfg.duracionMax,
      duracion_step: adminCfg.duracionStep
    };
  });
  __SALONES_CACHE__ = data;
  __SALONES_MAP__ = {};
  data.forEach(s => { __SALONES_MAP__[s.id] = s; });
  return data;
}

function salonesMap_(){
  if (!__SALONES_MAP__) salonesCache_();
  return __SALONES_MAP__ || {};
}

function invalidateSalonesCache_(){
  __SALONES_CACHE__ = null;
  __SALONES_MAP__ = null;
}

function apiListSalones(){
  const data = salonesCache_().map(s => ({ ...s }));
  return { ok:true, data };
}

function apiListSalonesAdmin(){
  const me = getUser_();
  if (!isAdminEmail_((me && me.email) || '')) return { ok:false, msg:'No autorizado' };
  const all = salonesCache_();
  const scope = adminScopeForUser_(me);
  const manage = scope === '1'
    ? all.slice()
    : all.filter(s => String(s.administracion_id||'1') === scope);
  return { ok:true, data:{ manage: manage.map(s => ({...s})), all: all.map(s => ({...s})) } };
}

function apiToggleSalon(salonId, habilitar){
  const me = getUser_();
  if (!me || String(me.rol||'').toUpperCase()!=='ADMIN' || String(me.estado||'').toUpperCase()!=='ACTIVO'){
    return {ok:false,msg:'No autorizado'};
  }
  const scope = adminScopeForUser_(me);
  const lr = SH_SAL.getLastRow(); if (lr<2) return {ok:false,msg:'No hay salones'};
  const rows = SH_SAL.getRange(2,1,lr-1,8).getValues();
  for (let i=0;i<rows.length;i++){
    if (rows[i][0]===salonId){
      const adminId = normalizeAdminId_(rows[i][6]||'1');
      if (scope !== '1' && scope !== adminId){
        return {ok:false, msg:'No autorizado para este salón'};
      }
      SH_SAL.getRange(i+2,4).setValue(habilitar?'SI':'NO');
      invalidateSalonesCache_();
      return {ok:true};
    }
  }
  return {ok:false,msg:'Salón no encontrado'};
}

// ========= Disponibilidad =========
function apiListDisponibilidad(fechaStr, salonId, duracionMin, userPrio){
  try {
    // --- Normalización de entrada ---
    fechaStr = String(fechaStr || '').slice(0, 10);
    salonId  = String(salonId || '').trim();

    const salon = getSalonById_(salonId);
    if (!salon) return { ok:true, data: [] };
    const adminId = salon.administracion_id || '1';
    const runtimeCfg = getAdminRuntimeCfg_(adminId);

    // Lee config y SANEAMOS todo (trim/valores por defecto)
    var cfgIni = String(runtimeCfg.horarioInicio||'07:00').trim();
    var cfgFin = String(runtimeCfg.horarioFin||'20:00').trim();
    var cfgStep= Number(runtimeCfg.duracionStep||30);
    var cfgMin = Number(runtimeCfg.duracionMin||30);
    var cfgMax = Number(runtimeCfg.duracionMax||240);

    var HINI = (/^\d{2}:\d{2}$/.test(cfgIni) ? cfgIni : '07:00');
    var HFIN = (/^\d{2}:\d{2}$/.test(cfgFin) ? cfgFin : '20:00');
    var STEP = (isFinite(cfgStep) && cfgStep >= 1) ? cfgStep : 30;

    var DUR  = Number(duracionMin || 0);
    if (!isFinite(DUR) || DUR < 1) DUR = cfgMin || 30;
    if (cfgMin && DUR < cfgMin) DUR = cfgMin;
    if (cfgMax && DUR > cfgMax) DUR = cfgMax;

    if (!fechaStr || !salonId) return { ok:true, data: [] };

    const me = getUser_();
    const prioAplicada = effectivePriorityForSalon_(me, salonId);
    userPrio = Number(prioAplicada||0);

    var startMin = toMin_(HINI);
    var endMin   = toMin_(HFIN);
    if (!(endMin > startMin)) { startMin = toMin_('07:00'); endMin = toMin_('20:00'); }

    if (DUR > (endMin - startMin)) return { ok:true, data: [] };

    // Genera slots candidatos
    var slots = [];
    for (var t = startMin; t + DUR <= endMin; t += STEP){
      slots.push({ inicio: toHHMM_(t), fin: toHHMM_(t + DUR) });
    }
    // Fallback por si STEP inválido en config
    if (!slots.length){
      for (var t2 = startMin; t2 + DUR <= endMin; t2 += 30){
        slots.push({ inicio: toHHMM_(t2), fin: toHHMM_(t2 + DUR) });
      }
    }
    if (!slots.length) return { ok:true, data: [] };

    // --- Restricciones del salón ---
    const restr = parseSalonRestriction_(salon && salon.restriccion);
    const restrictedIntervals = restr.intervals || [];
    const conflictStates = restr.requiresApproval ? ['APROBADA','PENDIENTE'] : ['APROBADA'];

    // --- Reservas del día/salón ---
    var ocupados = [];
    var sh = SS.getSheetByName('Reservas');
    if (sh && sh.getLastRow() >= 2) {
      var lr = sh.getLastRow(), lc = Math.max(22, sh.getLastColumn());
      var values = sh.getRange(2, 1, lr - 1, lc).getValues();
      for (var i=0; i<values.length; i++){
        var r = values[i];
        var estado = String(r[2] || '').toUpperCase();
        if (conflictStates.indexOf(estado) === -1) continue;
        if (ymdCell_(r[3]) !== String(fechaStr).slice(0,10)) continue;
        if (String(r[6] || '').trim() !== String(salonId).trim()) continue;

        var ini = hhmmFromCell_(r[4]);
        var fin = hhmmFromCell_(r[5]);
        var pr  = Number(r[15]||0);
        var email = String(r[9]||'');
        var nombre = String(r[10]||'');
        var evento = String(r[13]||'');
        var publico = String(r[14]||'');
        if (ini && fin){
          ocupados.push({ ini: ini, fin: fin, prio: pr, email: email, nombre: nombre, evento: evento, publico: publico, estado: estado });
        }
      }
    }

    userPrio = Number(userPrio||0);
    var out = slots.map(function(s){
      var sIni = toMin_(s.inicio), sFin = toMin_(s.fin);
      var maxPrio = -1; // -1 = sin ocupación
      var conflictos = [];
      var requiereConciliacion = false;
      var motivoConciliacion = '';
      for (var j=0; j<ocupados.length; j++){
        var occ = ocupados[j];
        var oIni = toMin_(occ.ini), oFin = toMin_(occ.fin);
        if (!(sFin <= oIni || sIni >= oFin)) {
          maxPrio = Math.max(maxPrio, Number(occ.prio||0));
          conflictos.push(occ);
        }
      }
      var esRestriccion = restrictedIntervals.some(function(it){ return !(sFin <= it.iniMin || sIni >= it.finMin); });
      var hayConflicto = conflictos.length > 0;
      if (hayConflicto){
        var hayEventoExterno = conflictos.some(function(c){ return isPublicoExternoOMixto_(c.publico); });
        if (hayEventoExterno){
          requiereConciliacion = true;
          motivoConciliacion = 'PUBLICO_EXTERNO';
        } else if (conflictos.some(function(c){ return Number(c.prio||0) === userPrio; })){ 
          requiereConciliacion = true;
          motivoConciliacion = 'MISMA_PRIORIDAD';
        }
      }
      var puedeReasignar = hayConflicto && !requiereConciliacion && (userPrio > maxPrio);
      var seleccionable = (!hayConflicto || puedeReasignar) && !esRestriccion;
      var disponible = !hayConflicto && !esRestriccion;
      var info = {
        inicio: s.inicio,
        fin: s.fin,
        disponible: disponible,
        maxPrio: hayConflicto ? maxPrio : null,
        seleccionable: seleccionable,
        requiereConciliacion: requiereConciliacion,
        motivoConciliacion: motivoConciliacion,
        ocupantes: conflictos.map(function(c){
          return {
            prioridad: Number(c.prio||0),
            email: c.email || '',
            nombre: c.nombre || '',
            evento: c.evento || '',
            publico_tipo: c.publico || '',
            inicio: c.ini,
            fin: c.fin,
            estado: c.estado || ''
          };
        })
      };
      if (esRestriccion){
        info.restringido = true;
        info.motivoRestriccion = 'RESTRICCION';
      }
      if (restr.requiresApproval){
        info.requiereAprobacion = true;
      }
      return info;
    });

    return { ok:true, data: out };
  } catch (e) {
    return { ok:false, error: String(e && e.message || e) };
  }
}

function apiCheckSlot(fechaStr, salonId, horaInicio, duracionMin, userPrio){
  try{
    fechaStr = String(fechaStr||'').slice(0,10);
    salonId  = String(salonId||'').trim();
    var dur  = Math.max(Number(duracionMin||0), 1);
    if (!fechaStr || !salonId || !/^\d{2}:\d{2}$/.test(horaInicio) || !dur){
      return { ok:true, disponible:false };
    }

    const salon = getSalonById_(salonId);
    if (!salon) return { ok:true, disponible:false };
    const adminId = salon.administracion_id || '1';
    const runtimeCfg = getAdminRuntimeCfg_(adminId);

    // Normaliza ventana para no permitir fuera de horario
    var cfgIni = String(runtimeCfg.horarioInicio||'07:00').trim();
    var cfgFin = String(runtimeCfg.horarioFin||'20:00').trim();
    var HINI   = (/^\d{2}:\d{2}$/.test(cfgIni)? cfgIni : '07:00');
    var HFIN   = (/^\d{2}:\d{2}$/.test(cfgFin)? cfgFin : '20:00');

    const minDur = Math.max(1, Number(runtimeCfg.duracionMin||1));
    const maxDur = Math.max(minDur, Number(runtimeCfg.duracionMax||minDur));
    if (dur < minDur) dur = minDur;
    if (dur > maxDur) dur = maxDur;

    var hIni = normHHMM_(horaInicio);
    if (!fechaStr || !salonId || !hIni || !dur){
      return { ok:true, disponible:false };
    }

    var horaFin = addMinutes_(hIni, dur);

    // fuera de ventana → no disponible
    if (toMin_(hIni) < toMin_(HINI) || toMin_(horaFin) > toMin_(HFIN)) {
      return { ok:true, disponible:false };
    }

    const restr = parseSalonRestriction_(salon && salon.restriccion);
    const restricted = (restr.intervals||[]).some(function(it){
      return !(toMin_(horaFin) <= it.iniMin || toMin_(hIni) >= it.finMin);
    });
    if (restricted){
      return { ok:true, disponible:false, motivo:'RESTRICCION', requiereAprobacion: restr.requiresApproval || false };
    }

    const me = getUser_();
    const prioAplicada = effectivePriorityForSalon_(me, salonId);

    var conflictStates = restr.requiresApproval ? ['APROBADA','PENDIENTE'] : ['APROBADA'];
    var conflicts = getConflicts_(fechaStr, salonId, hIni, horaFin, conflictStates);
    var maxPrio =  conflicts.reduce((m,c)=>Math.max(m, Number(c.prioridad||0)), -1);
    var hayEventoExterno = conflicts.some(function(c){ return isPublicoExternoOMixto_(c.publico_tipo); });
    var hayPendiente = conflicts.some(function(c){ return String(c.estado||'').toUpperCase()==='PENDIENTE'; });
    userPrio = Number(prioAplicada||0);
    var disponible = (conflicts.length===0);
    if (hayPendiente){
      disponible = false;
    }
    if (!disponible){
      if (hayEventoExterno){
        disponible = false;
      } else if (userPrio > maxPrio){
        disponible = true;
      }
    }
    return { ok:true, disponible, maxPrio: maxPrio>=0?maxPrio:null, requiereAprobacion: restr.requiresApproval || false };
  }catch(e){
    return { ok:false, msg:String(e && e.message || e) };
  }
}

// ========= Reservas =========
function apiCrearReserva(payload){
  try{
    const me = getUser_();
    if (!me || !me.email) return {ok:false, msg:'Usuario no autenticado.'};

    // ---- Salón / capacidad
    const sal = getSalonById_(payload.salon_id);
    if (!sal) return {ok:false, msg:'Salón inválido.'};
    if (!sal.habilitado) return {ok:false, msg:'Este salón está deshabilitado temporalmente.'};
    const restrSalon = parseSalonRestriction_(sal.restriccion);
    const adminId = sal.administracion_id || '1';
    const runtimeCfg = getAdminRuntimeCfg_(adminId);

    const cant = Number(payload.cant_personas || 0);
    if (!isFinite(cant) || cant < 1) return {ok:false, msg:'Indica la cantidad de personas (número válido).'};
    if (sal.capacidad && cant > sal.capacidad){
      return {ok:false, msg:'La cantidad de personas excede la capacidad máxima del salón ('+sal.capacidad+').'};
    }

    // ---- Normalización de inputs
    const fecha = String(payload.fecha||'').slice(0,10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return {ok:false, msg:'Fecha inválida.'};

    const horaIniIn = String(payload.hora_inicio||'');
    const horaIni   = normHHMM_((horaIniIn.length<=5?horaIniIn:horaIniIn.slice(0,5)));
    if (!horaIni) return {ok:false, msg:'Hora de inicio inválida.'};

    const cfgDurMin = Math.max(1, Number(runtimeCfg.duracionMin||30));
    const cfgDurMax = Math.max(cfgDurMin, Number(runtimeCfg.duracionMax||cfgDurMin));
    let DUR = Number(payload.duracion_min||0);
    if (!isFinite(DUR) || DUR < cfgDurMin) DUR = cfgDurMin;
    if (DUR > cfgDurMax) DUR = cfgDurMax;

    const horaFin = addMinutes_(horaIni, DUR);

    // ---- Ventana de horario (07:00–20:00 y última hora de inicio 19:00)
    const HINI = normHHMM_(runtimeCfg.horarioInicio || '07:00') || '07:00';
    const HFIN = normHHMM_(runtimeCfg.horarioFin    || '20:00') || '20:00';

    const iniMin = toMin_(horaIni);
    const finMin = toMin_(horaFin);
    if (isNaN(iniMin) || isNaN(finMin)) return {ok:false, msg:'Hora inválida.'};

    if (iniMin < toMin_(HINI)) return { ok:false, msg:'Hora fuera de horario permitido.' };
    if (finMin > toMin_(HFIN)) return { ok:false, msg:'La reserva debe terminar a más tardar a las '+HFIN+'.' };
    if (iniMin > toMin_('19:00')) return { ok:false, msg:'La última hora de inicio permitida es 19:00.' };

    // ---- Restricciones horarias específicas del salón
    const esRestriccion = (restrSalon.intervals||[]).some(it => !(finMin <= it.iniMin || iniMin >= it.finMin));
    if (esRestriccion){
      return { ok:false, msg:'Este salón no permite reservas en el horario seleccionado. Elige otro intervalo.' };
    }

    // ---- Conserje (>= 16:00 cualquiera de los extremos)
    const requiereConserjeSalon = String(sal.requiere_conserje||'SI').toUpperCase() !== 'NO';
    const conserjeReq = (requiereConserjeSalon && (iniMin >= toMin_('16:00') || finMin > toMin_('16:00'))) ? 'SI' : 'NO';

    // ---- Conflictos / prioridad
    // Prioridad del solicitante considerando salones permitidos (admins sin prioridad explícita => 2)
    const prioSolic = effectivePriorityForSalon_(me, sal.id);

    const conflictStates = restrSalon.requiresApproval ? ['APROBADA','PENDIENTE'] : ['APROBADA'];
    const conflicts = getConflicts_(fecha, sal.id, horaIni, horaFin, conflictStates);
    const hayPendiente = conflicts.some(c => String(c.estado||'').toUpperCase()==='PENDIENTE');
    if (hayPendiente){
      return { ok:false, msg:'Ya existe una solicitud pendiente para ese horario. Elige otro intervalo.' };
    }

    // 🔒 Cortafuegos explícito: prioridad 0 no puede solapar NUNCA
    if (conflicts.length && prioSolic <= 0){
      return { ok:false, msg:'Ese intervalo ya está reservado. Elige otra hora.' };
    }

    if (conflicts.length){
      const tieneEventoExterno = conflicts.some(c => isPublicoExternoOMixto_(c.publico_tipo));
      if (tieneEventoExterno){
        return { ok:false, msg:'Ese intervalo está reservado para un evento con público externo o mixto. Deben coordinar con el anfitrión para liberar el espacio.' };
      }
      const mismaPrioridad = conflicts.some(c => Number(c.prioridad||0) === prioSolic);
      let maxPrio = 0;
      for (let i=0;i<conflicts.length;i++){
        maxPrio = Math.max(maxPrio, Number(conflicts[i].prioridad||0));
      }
      if (mismaPrioridad){
        return { ok:false, msg:'Ese intervalo ya está reservado por otra persona con la misma prioridad. Deben coordinar con la otra parte para que cancele la reserva antes de agendar.' };
      }
      if (prioSolic <= maxPrio){
        return { ok:false, msg:'Ese intervalo ya está reservado. Elige otra hora.' };
      }
      // Mi prioridad es mayor: cancelar las inferiores (solo si no son eventos externos/mixtos)
      for (let j=0;j<conflicts.length;j++){
        if (String(conflicts[j].estado||'').toUpperCase()==='PENDIENTE') continue;
        if (isPublicoExternoOMixto_(conflicts[j].publico_tipo)) continue;
        if (prioSolic > Number(conflicts[j].prioridad||0)){
          cancelReservaById_(conflicts[j].id, 'Reasignada por usuario con prioridad', me.email);
        }
      }
    }

    // ---- Alta
    const id    = nextId_(SH_RES, 'R-');
    const token = Utilities.getUuid();
    const ts    = nowStr_();

    // Validación ligera de correo provisto (opcional pero útil)
    const emailForm = String(payload.email||'').trim().toLowerCase();
    const emailUso  = emailForm || me.email; // si no manda correo, caemos al de sesión

    const estadoInicial = restrSalon.requiresApproval ? 'PENDIENTE' : 'APROBADA';

    SH_RES.appendRow([
      id, token, estadoInicial, fecha, horaIni, horaFin, sal.id, sal.nombre,
      cant, emailUso, (payload.nombre||me.nombre||''), (payload.departamento||me.departamento||''),
      String(payload.extension||''), String(payload.evento_nombre||''), String(payload.publico_tipo||''),
      prioSolic, conserjeReq, 'NO', ts, ts, '', '', '', adminId
    ]);

    if (restrSalon.requiresApproval){
      try{ sendEmailPendiente_(id); }catch(e){}
      try{ notifyAdminReservaPendiente_(id); }catch(e){}
    } else {
      try{ sendEmailConfirmacion_(id); }catch(e){}
      // Si la reserva requiere conserje, notificar al área de Conserjería
      try{ notificarConserjeria_(id); }catch(e){}
    }

    return { ok:true, id, token, estado: estadoInicial, requiere_aprobacion: restrSalon.requiresApproval };
  }catch(e){
    return { ok:false, msg: String(e && e.message || e) };
  }
}

function getSalonById_(id){
  const map = salonesMap_();
  const rec = map[String(id||'')] || null;
  if (!rec) return null;
  // devolver copia para evitar mutaciones externas
  return { ...rec };
}

function parseSalonRestriction_(raw){
  const text = String(raw||'').trim();
  if (!text) return { raw:'', requiresApproval:false, intervals:[] };
  const parts = text.split(';').map(s=>String(s||'').trim()).filter(Boolean);
  const intervals = [];
  let requiresApproval = false;
  parts.forEach(part => {
    const up = part.toUpperCase();
    if (up === 'CONFIRM' || up === 'COFNIRM'){
      requiresApproval = true;
      return;
    }
    const tokens = part.split('-').map(s=>String(s||'').trim()).filter(Boolean);
    if (tokens.length !== 2) return;
    const iniStr = normHHMM_(tokens[0]);
    const finStr = normHHMM_(tokens[1]);
    if (!iniStr || !finStr) return;
    const iniMin = toMin_(iniStr);
    const finMin = toMin_(finStr);
    if (!isFinite(iniMin) || !isFinite(finMin) || finMin <= iniMin) return;
    intervals.push({ inicio:iniStr, fin:finStr, iniMin, finMin });
  });
  return { raw:text, requiresApproval, intervals };
}

function getConflicts_(fecha, salonId, ini, fin, estados){
  const lr = SH_RES.getLastRow(); if (lr<2) return [];
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();

  const iniMin = toMin_(normHHMM_(ini));
  const finMin = toMin_(normHHMM_(fin));
  if (isNaN(iniMin) || isNaN(finMin)) return [];

  const estadosCmp = (Array.isArray(estados) && estados.length ? estados : ['APROBADA'])
    .map(s => String(s||'').toUpperCase());

  return rows
    .filter(r => estadosCmp.indexOf(String(r[2]||'').toUpperCase()) !== -1
              && ymdCell_(r[3])===String(fecha).slice(0,10)
              && String(r[6]).trim()===String(salonId).trim())
    .filter(r => {
      const rIniMin = toMin_(hhmmFromCell_(r[4]));
      const rFinMin = toMin_(hhmmFromCell_(r[5]));
      if (isNaN(rIniMin) || isNaN(rFinMin)) return false;
      return !(finMin <= rIniMin || iniMin >= rFinMin); // hay solape si NO se separan
    })
    .map(r => ({
      id: r[0],
      estado: String(r[2]||''),
      prioridad: Number(r[15]||0),
      solicitante: String(r[9]||''),
      solicitante_email: String(r[9]||''),
      solicitante_nombre: String(r[10]||''),
      evento: String(r[13]||''),
      publico_tipo: String(r[14]||''),
      inicio: hhmmFromCell_(r[4]),
      fin: hhmmFromCell_(r[5])
    }));
}

function apiListMisReservas(fechaDesde, fechaHasta){
  const me = getUser_();
  const myEmail = String(me && me.email || '').toLowerCase();
  const lr = SH_RES.getLastRow(); if (lr<2) return {ok:true, data:[]};
  const lc = Math.max(24, SH_RES.getLastColumn());              // 👈 asegura leer col 24
  const rows = SH_RES.getRange(2,1,lr-1,lc).getValues();

  const d1 = fechaDesde ? toDate_(fechaDesde) : new Date(2000,0,1);
  const d2 = fechaHasta ? toDate_(fechaHasta) : new Date(2100,0,1);

  let data = rows
    .filter(r => String(r[9]||'').toLowerCase() === myEmail) // solicitante_email
    .filter(r => {
      const d = new Date(r[3]); // fecha (puede ser Date)
      if (isNaN(d)) return false;
      // Normalizamos a día (sin hora)
      const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      const d1n = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate());
      const d2n = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
      return dd >= d1n && dd <= d2n;
    })
    .map(r => toReservaObj_(r));

    // Enriquecer con nombre de conserje (si activo)
    const cmap = conserjeMap_();
    data = data.map(x => ({
      ...x,
      conserje_nombre: (x.conserje_codigo_asignado && cmap[x.conserje_codigo_asignado]?.nombre) || ''
    }));
  return {ok:true, data};
}

function apiListReservasAdmin(fechaDesde, fechaHasta){
  const me = getUser_();
  if (!isAdminEmail_((me && me.email) || '')) return {ok:false, msg:'No autorizado'};
  const scope = adminScopeForUser_(me);
  const lr = SH_RES.getLastRow(); if (lr<2) return {ok:true, data:[]};
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  const d1 = fechaDesde ? toDate_(fechaDesde) : new Date(2000,0,1);
  const d2 = fechaHasta ? toDate_(fechaHasta) : new Date(2100,0,1);
  let data = rows
    .filter(r => { const d=toDate_(r[3]); return d>=d1 && d<=d2; })
    .filter(r => scope==='1' || normalizeAdminId_(r[23]||'1') === scope)
    .map(r => toReservaObj_(r));
  const cmap = conserjeMap_();
  data = data.map(x => ({ ...x, conserje_nombre: (x.conserje_codigo_asignado && cmap[x.conserje_codigo_asignado]?.nombre) || '' }));
  return {ok:true, data};
}

function toReservaObj_(r){
  return {
    id:r[0], token:r[1], estado:r[2],
    fecha: ymdCell_(r[3]),                                // 'YYYY-MM-DD'
    hora_inicio: hhmmFromCell_(r[4]),                     // 'HH:mm'
    hora_fin:    hhmmFromCell_(r[5]),                     // 'HH:mm'
    salon_id:r[6], salon_nombre:r[7],
    cant_personas:Number(r[8]||0),
    solicitante_email:String(r[9]||''), 
    solicitante_nombre:String(r[10]||''),
    departamento:String(r[11]||''),
    extension:String(r[12]||''),
    evento_nombre:String(r[13]||''),
    publico_tipo:String(r[14]||''),
    prioridad:Number(r[15]||0),
    conserje_requerido:String(r[16]||''),
    conserje_notificado:String(r[17]||''),
    creado_en: tsCell_(r[18]),              // ✅ string seguro
    actualizado_en: tsCell_(r[19]),         // ✅ string seguro
    cancelado_por:String(r[20]||''),
    cancelado_motivo:String(r[21]||''),
    conserje_codigo_asignado:String(r[22]||''),
    administracion_id: normalizeAdminId_(r[23]||'1')
  };
}

function apiCancelarReservaByToken(token, motivo){
  if (!token) return { ok:false, msg:'Token inválido' };
  const lr = SH_RES.getLastRow(); if (lr<2) return {ok:false,msg:'No hay reservas'};
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  for (let i=0;i<rows.length;i++){
    if (rows[i][1]===token){
      if (String(rows[i][2]).toUpperCase()==='CANCELADA') return {ok:false,msg:'Ya estaba cancelada'};
      SH_RES.getRange(i+2,3).setValue('CANCELADA');
      SH_RES.getRange(i+2,22).setValue(String(motivo||'Cancelada por el solicitante vía enlace'));
      SH_RES.getRange(i+2,21).setValue((Session.getActiveUser().getEmail()||'')||'PUBLIC');
      SH_RES.getRange(i+2,20).setValue(nowStr_());
      SpreadsheetApp.flush();
      try{ sendEmailCancelacion_(rows[i][0]); }catch(e){}
      return {ok:true};
    }
  }
  return {ok:false,msg:'Reserva no encontrada'};
}

function apiCancelarReservaAdmin(id, motivo){
  const me = (Session.getActiveUser().getEmail()||'');
  const user = getUser_();
  if (!isAdminEmail_(me)) return {ok:false,msg:'No autorizado'};

  // Guardarrail: hasta 30 min antes del inicio
  const r = getReservaById_(id);
  if (!r) return {ok:false,msg:'Reserva no encontrada'};
  const scope = adminScopeForUser_(user);
  if (scope !== '1' && normalizeAdminId_(r.administracion_id||'1') !== scope){
    return {ok:false, msg:'No autorizado para este salón'};
  }
  const estado = String(r.estado).toUpperCase();
  if (estado==='CANCELADA') return {ok:false,msg:'La reserva ya está cancelada'};
  if (estado!=='PENDIENTE'){
    const start = toDate_(r.fecha);
    const [hh,mi] = (r.hora_inicio||'00:00').split(':').map(Number);
    start.setHours(hh||0, mi||0, 0, 0);
    const now = new Date();
    if (now.getTime() > (start.getTime() - 30*60*1000)) {
      return {ok:false, msg:'Fuera de ventana: solo se permite cancelar hasta 30 min antes del inicio.'};
    }
  }
  return cancelReservaById_(id, motivo||'Cancelada por administración', me);
}

function apiAprobarReservaAdmin(id){
  const me = (Session.getActiveUser().getEmail()||'');
  const user = getUser_();
  if (!isAdminEmail_(me)) return {ok:false,msg:'No autorizado'};
  const r = getReservaById_(id);
  if (!r) return {ok:false,msg:'Reserva no encontrada'};
  const scope = adminScopeForUser_(user);
  if (scope !== '1' && normalizeAdminId_(r.administracion_id||'1') !== scope){
    return {ok:false, msg:'No autorizado para este salón'};
  }
  const estado = String(r.estado||'').toUpperCase();
  if (estado === 'CANCELADA') return {ok:false,msg:'La reserva está cancelada'};
  if (estado === 'APROBADA') return {ok:true, already:true};
  setReservaField_(id, 'estado', 'APROBADA');
  try{ sendEmailAprobacion_(id); }catch(e){}
  try{ notificarConserjeria_(id); }catch(e){}
  return {ok:true};
}

function cancelReservaById_(id, motivo, who){
  const lr = SH_RES.getLastRow(); if (lr<2) return {ok:false,msg:'No hay reservas'};
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  for (let i=0;i<rows.length;i++){
    if (rows[i][0]===id){
      if (String(rows[i][2]).toUpperCase()==='CANCELADA') return {ok:false,msg:'Ya estaba cancelada'};
      SH_RES.getRange(i+2,3).setValue('CANCELADA');
      SH_RES.getRange(i+2,21).setValue(String(who||'ADMIN'));
      SH_RES.getRange(i+2,22).setValue(motivo||'');
      SH_RES.getRange(i+2,20).setValue(nowStr_());
      SpreadsheetApp.flush();
      try{ sendEmailCancelacion_(rows[i][0]); }catch(e){}
      return {ok:true};
    }
  }
  return {ok:false,msg:'Reserva no encontrada'};
}

// ========= Emails =========
function sendEmailConfirmacion_(reservaId){
  const r = getReservaById_(reservaId); if (!r) return;
  const adminId = normalizeAdminId_(r.administracion_id||'1');
  const base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();
  const cancelUrl = normalizeExecUrl_(base) + `?cancel=${encodeURIComponent(r.token)}`;
  const subj = `Reserva confirmada – ${r.salon_nombre} – ${fmtDMY_(r.fecha)} ${fmt12_(r.hora_inicio)}`;
  const html = emailLayout_({
    title: 'Reserva confirmada',
    preheader: `Tu reserva para ${r.salon_nombre} el ${r.fecha} fue registrada.`,
    htmlInner:
      emailParagraph_(`Hola <b>${escapeHtml_(r.solicitante_nombre||'')}</b>,`) +
      emailParagraph_('Tu reserva fue registrada con éxito. Estos son los detalles:') +
      emailDetailsRows_([
        ['Salón', r.salon_nombre],
        ['Fecha', fmtDMY_(r.fecha)],
        ['Hora', fmt12_(r.hora_inicio) + ' – ' + fmt12_(r.hora_fin)],
        ['Evento', r.evento_nombre||''],
        ['Asistentes', String(r.cant_personas||0)],
        ['Público', r.publico_tipo||'']
      ]) +
      emailParagraph_(`¿Ya no necesitas este espacio? Puedes cancelar la reserva con el botón.`),
    ctaUrl: cancelUrl,
    ctaLabel: 'Cancelar reserva',
    footer: 'Si tienes dudas, responde a este correo o contacta al equipo de coordinación.',
    adminId
  });

  MailApp.sendEmail({
    to: r.solicitante_email,
    subject: subj,
    htmlBody: html,
    name: cfgForAdmin_(adminId, 'MAIL_SENDER_NAME') || 'Reserva de Salones',
    replyTo: cfgForAdmin_(adminId, 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
  });
}

function sendEmailCancelacion_(reservaId){
  const r = getReservaById_(reservaId); if (!r) return;
  const adminId = normalizeAdminId_(r.administracion_id||'1');
  const subj = `Reserva cancelada – ${r.salon_nombre} – ${fmtDMY_(r.fecha)} ${fmt12_(r.hora_inicio)}`;
  const base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();
  const appUrl = normalizeExecUrl_(base);
  // Motivo “amable” si fue por prioridad
  var rawMotivo = String(r.cancelado_motivo||'No especificado');
  var softMotivo = /reasignada\s+por\s+usuario\s+con\s+prioridad/i.test(rawMotivo)
    ? 'Disponibilidad prioritaria en ese horario.'
    : rawMotivo;
  const html = emailLayout_({
    title: 'Reserva cancelada',
    preheader: `Se canceló la reserva de ${r.salon_nombre} para el ${fmtDMY_(r.fecha)}.`,
    htmlInner:
    emailParagraph_(`Se canceló la reserva del salón <b>${escapeHtml_(r.salon_nombre)}</b> para el <b>${escapeHtml_(fmtDMY_(r.fecha))}</b> a las <b>${escapeHtml_(fmt12_(r.hora_inicio))}</b> – lamentamos los inconvenientes.`) +
      emailParagraph_(`<b>Motivo:</b> ${escapeHtml_(softMotivo)}`) +
      emailParagraph_('Si aún necesitas el espacio, puedes realizar una nueva reserva desde el sistema.'),
    ctaUrl: appUrl,
    ctaLabel: 'Hacer nueva reserva',
    footer: 'Si esto fue un error, crea una nueva reserva o contacta a administración.',
    adminId
  });

  MailApp.sendEmail({
    to: r.solicitante_email,
    subject: subj,
    htmlBody: html,
    name: cfgForAdmin_(adminId, 'MAIL_SENDER_NAME') || 'Reserva de Salones',
    replyTo: cfgForAdmin_(adminId, 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
  });
}

function sendEmailPendiente_(reservaId){
  const r = getReservaById_(reservaId); if (!r) return;
  const adminId = normalizeAdminId_(r.administracion_id||'1');
  const base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();
  const cancelUrl = normalizeExecUrl_(base) + `?cancel=${encodeURIComponent(r.token)}`;
  const subj = `Reserva pendiente – ${r.salon_nombre} – ${fmtDMY_(r.fecha)} ${fmt12_(r.hora_inicio)}`;
  const disclaimer = '<tr><td style="padding:12px 2px 0 2px;"><div style="padding:12px 16px; border-radius:12px; border:1px solid #fcd34d; background:#fef3c7; color:#92400e; font-weight:600;">Este salón es restringido y requiere aprobación de la administración. Tu reserva aún no está confirmada. Recibirás una notificación cuando se haya evaluado.</div></td></tr>';
  const html = emailLayout_({
    title: 'Reserva pendiente de aprobación',
    preheader: `Tu solicitud para ${r.salon_nombre} está pendiente de aprobación.`,
    htmlInner:
      emailParagraph_(`Hola <b>${escapeHtml_(r.solicitante_nombre||'')}</b>,`) +
      emailParagraph_('Recibimos tu solicitud de reserva y se envió a la administración para aprobación. A continuación, el resumen del evento:') +
      emailDetailsRows_([
        ['Salón', r.salon_nombre],
        ['Fecha', fmtDMY_(r.fecha)],
        ['Hora', fmt12_(r.hora_inicio) + ' – ' + fmt12_(r.hora_fin)],
        ['Evento', r.evento_nombre||''],
        ['Asistentes', String(r.cant_personas||0)],
        ['Público', r.publico_tipo||'']
    ]) +
      emailParagraph_('Si ya no necesitas el espacio puedes cancelar la solicitud con el botón.') +
      disclaimer,
    ctaUrl: cancelUrl,
    ctaLabel: 'Cancelar solicitud',
    footer: 'La administración revisará tu solicitud y recibirás un correo con la decisión.',
    adminId
  });

  MailApp.sendEmail({
    to: r.solicitante_email,
    subject: subj,
    htmlBody: html,
    name: cfgForAdmin_(adminId, 'MAIL_SENDER_NAME') || 'Reserva de Salones',
    replyTo: cfgForAdmin_(adminId, 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
  });
}

function sendEmailAprobacion_(reservaId){
  const r = getReservaById_(reservaId); if (!r) return;
  const adminId = normalizeAdminId_(r.administracion_id||'1');
  const base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();
  const cancelUrl = normalizeExecUrl_(base) + `?cancel=${encodeURIComponent(r.token)}`;
  const subj = `Reserva aprobada – ${r.salon_nombre} – ${fmtDMY_(r.fecha)} ${fmt12_(r.hora_inicio)}`;
  const html = emailLayout_({
    title: 'Reserva aprobada',
    preheader: `La administración aprobó tu reserva para ${r.salon_nombre}.`,
    htmlInner:
      emailParagraph_(`Hola <b>${escapeHtml_(r.solicitante_nombre||'')}</b>,`) +
      emailParagraph_('La administración revisó tu solicitud y la reserva fue aprobada. Estos son los detalles confirmados:') +
      emailDetailsRows_([
        ['Salón', r.salon_nombre],
        ['Fecha', fmtDMY_(r.fecha)],
        ['Hora', fmt12_(r.hora_inicio) + ' – ' + fmt12_(r.hora_fin)],
        ['Evento', r.evento_nombre||''],
        ['Asistentes', String(r.cant_personas||0)],
        ['Público', r.publico_tipo||'']
    ]) +
      emailParagraph_('Si necesitas desistir de la reserva puedes cancelarla desde el sistema con el botón.'),
    ctaUrl: cancelUrl,
    ctaLabel: 'Cancelar reserva',
    footer: 'Gracias por utilizar el sistema de reservas.',
    adminId
  });

  MailApp.sendEmail({
    to: r.solicitante_email,
    subject: subj,
    htmlBody: html,
    name: cfgForAdmin_(adminId, 'MAIL_SENDER_NAME') || 'Reserva de Salones',
    replyTo: cfgForAdmin_(adminId, 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
  });
}

function notifyAdminReservaPendiente_(reservaId){
  const r = getReservaById_(reservaId); if (!r) return;
  const adminId = normalizeAdminId_(r.administracion_id||'1');
  const toList = (cfgForAdmin_(adminId, 'ADMIN_EMAILS') || cfgForAdmin_(adminId, 'ADMIN_CONTACT_EMAIL') || '')
    .split(';').map(s=>s.trim()).filter(Boolean);
  if (!toList.length) return;
  const subj = `Pendiente de aprobación – ${r.salon_nombre} – ${fmtDMY_(r.fecha)} ${fmt12_(r.hora_inicio)}`;
  const intro = emailParagraph_('Hola equipo de administración,')
    + emailParagraph_('Se registró una nueva reserva en un salón restringido y requiere aprobación.');
  const details = emailDetailsRows_([
    ['Salón', r.salon_nombre],
    ['Fecha', fmtDMY_(r.fecha)],
    ['Hora', `${fmt12_(r.hora_inicio)} – ${fmt12_(r.hora_fin)}`],
    ['Evento', r.evento_nombre || ''],
    ['Solicitante', `${r.solicitante_nombre||''} (${r.solicitante_email||''})`],
    ['Público', r.publico_tipo || '']
  ]);
  const reminder = '<tr><td style="padding:12px 2px 0 2px;"><div style="padding:12px 16px; border-radius:12px; border:1px solid #bfdbfe; background:#dbeafe; color:#1e3a8a; font-weight:600;">Revisa la solicitud, apruébala o cancélala desde el panel administrativo.</div></td></tr>';
  const html = emailLayout_({
    title: 'Nueva reserva pendiente de aprobación',
    preheader: `Solicitud registrada para ${r.salon_nombre}.`,
    htmlInner: intro + details + reminder,
    ctaUrl: adminPanelUrl_(),
    ctaLabel: 'Abrir panel administrativo',
    footer: 'Gracias por gestionar las solicitudes restringidas.',
    adminId
  });

  MailApp.sendEmail({
    to: toList.join(','),
    subject: subj,
    htmlBody: html,
    name: cfgForAdmin_(adminId, 'MAIL_SENDER_NAME') || 'Reserva de Salones',
    replyTo: cfgForAdmin_(adminId, 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
  });
}

function notificarConserjeria_(reservaId){
  const r = getReservaById_(reservaId); if (!r || String(r.conserje_requerido)!=='SI') return;
  const adminId = normalizeAdminId_(r.administracion_id||'1');
  // Correos de Conserjería desde Config (separados por ;)
  const to = (cfg_('CONSERJERIA_EMAILS')||'').split(';').map(s=>s.trim()).filter(Boolean).join(',');
  if (!to) return;
  const subj = `Asignación de conserje – ${r.salon_nombre} – ${fmtDMY_(r.fecha)} ${fmt12_(r.hora_inicio)}`;
  const base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();
  const appUrl = normalizeExecUrl_(base);

  // Contenido enriquecido
  const intro =
    emailParagraph_(
      `Hola equipo de <b>Conserjería</b>,`
    ) +
    emailParagraph_(
      `Se registró una nueva reserva que <b>requiere asignación de conserje</b>. A continuación, el resumen del evento.`
    );

  const details = emailDetailsRows_([
    ['Salón', r.salon_nombre],
    ['Fecha', fmtDMY_(r.fecha)],
    ['Horario', `${fmt12_(r.hora_inicio)} – ${fmt12_(r.hora_fin)}`],
    ['Evento', r.evento_nombre || ''],
    ['Solicitante', `${r.solicitante_nombre||''} (${r.solicitante_email||''})`]
  ]);

  // Checklist y nota importante
  const checklist = `
    <tr><td style="padding:8px 2px; color:#374151; font-size:14px;">
      <div style="margin:6px 0 4px; font-weight:600;">Pasos sugeridos</div>
      <ol style="margin:6px 0 0 18px; padding:0; color:#374151;">
        <li>Ingresar al panel y abrir la reserva.</li>
        <li>Seleccionar un conserje <b>activo</b> y <b>disponible</b> para el horario indicado.</li>
        <li>Confirmar la asignación (se notificará automáticamente por correo al conserje).</li>
      </ol>
      <div style="margin-top:10px; padding:10px; border:1px solid #e5e7eb; border-radius:10px; background:#f9fafb;">
        <div style="font-weight:600; color:#111827; margin-bottom:4px;">Importante</div>
        <div style="color:#4b5563;">
          Verifica conflictos de horario antes de asignar. Si no hay personal disponible, por favor comunica la incidencia a la administración.
        </div>
      </div>
    </td></tr>`;

  const htmlWithCta = emailLayout_({
    title: 'Se requiere asignación de conserje',
    preheader: `Reserva en ${r.salon_nombre} – ${fmtDMY_(r.fecha)} ${fmt12_(r.hora_inicio)}`,
    htmlInner: intro + details + checklist,
    ctaUrl: appUrl,
    ctaLabel: 'Asignar conserje ahora',
    footer: 'Gracias por su apoyo logístico.',
    adminId
  });

  MailApp.sendEmail({
    to,
    subject: subj,
    htmlBody: htmlWithCta,
    name: cfgForAdmin_(adminId, 'MAIL_SENDER_NAME') || 'Reserva de Salones',
    replyTo: cfgForAdmin_(adminId, 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
  });
  setReservaField_(r.id, 'conserje_notificado', 'SI');
}

function getReservaById_(id){
  const lr = SH_RES.getLastRow(); if (lr<2) return null;
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  const r = rows.find(v=>v[0]===id);
  return r ? toReservaObj_(r) : null;
}

//Obtener reserva por token (para la pantalla de cancelación)
function apiGetReservaByToken(token){
  token = String(token||'').trim();
  if (!token) return { ok:false, msg:'Token inválido' };
  const lr = SH_RES.getLastRow(); if (lr<2) return { ok:false, msg:'No hay reservas' };
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  for (let i=0;i<rows.length;i++){
    if (String(rows[i][1]||'') === token){
      return { ok:true, data: toReservaObj_(rows[i]) };
    }
  }
  return { ok:false, msg:'Reserva no encontrada' };
}

function setReservaField_(id, field, value){
  const colIdx = {
    estado:3, fecha:4, hora_inicio:5, hora_fin:6, salon_id:7, salon_nombre:8, cant_personas:9,
    solicitante_email:10, solicitante_nombre:11, departamento:12, extension:13, evento_nombre:14, publico_tipo:15,
    prioridad:16, conserje_requerido:17, conserje_notificado:18, creado_en:19, actualizado_en:20, cancelado_por:21, cancelado_motivo:22, conserje_codigo_asignado:23, administracion_id:24
  };
  const c = colIdx[field]; if (!c) return;
  const lr = SH_RES.getLastRow(); if (lr<2) return;
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  for (let i=0;i<rows.length;i++){ if (rows[i][0]===id){ SH_RES.getRange(i+2,c).setValue(value); SH_RES.getRange(i+2,20).setValue(nowStr_()); break; } }
}

function adminPanelUrl_(){
  const explicit = cfg_('ADMIN_PANEL_URL') || cfg_('ADMIN_WEBAPP_URL');
  if (explicit) return explicit;
  const base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();
  const exec = normalizeExecUrl_(base);
  return exec + (exec.indexOf('?')===-1 ? '?mode=admin' : '&mode=admin');
}

function normalizeExecUrl_(url){
  let base = String(url||'');
  if (!/\/exec(\?|$)/.test(base)) base = base.replace(/\/(dev|user|exec)?(\?.*)?$/, '')+'/exec';
  if (!/[?&]authuser=/.test(base)) base += (base.indexOf('?')===-1?'?':'&')+'authuser=1';
  return base;
}

/** ====== Email branding / layout helpers ====== **/
function mailCfg_(adminId){
  const scope = normalizeAdminId_(adminId||'1');
  return {
    brandName: cfgForAdmin_(scope, 'MAIL_SENDER_NAME') || 'Reserva de Salones',
    replyTo:   cfgForAdmin_(scope, 'MAIL_REPLY_TO')    || Session.getActiveUser().getEmail(),
    color:     cfg_('MAIL_BRAND_COLOR') || '#1d4ed8',
    bg:        cfg_('MAIL_BG_COLOR')    || '#f6f7fb',
    text:      cfg_('MAIL_TEXT_COLOR')  || '#111827',
    card:      cfg_('MAIL_CARD_BG')     || '#ffffff',
    border:    cfg_('MAIL_BORDER')      || '#e5e7eb',
    logoSize:  Number(cfg_('MAIL_LOGO_SIZE') || 100),
    logo:      (cfg_('MAIL_LOGO_URL') || getLogoUrl() || getLogoDataUrl(Number(cfg_('MAIL_LOGO_SIZE') || 64)) || '')
  };
}

function emailLayout_(opts){
  const scope = normalizeAdminId_(opts && opts.adminId || '1');
  var m = mailCfg_(scope);
  var title     = String(opts.title || m.brandName);
  var preheader = String(opts.preheader || 'Notificación del sistema de reservas');
  var inner     = String(opts.htmlInner || '');
  var ctaUrl    = String(opts.ctaUrl || '');
  var ctaLabel  = String(opts.ctaLabel || '');
  var footer    = String(opts.footer || 'Este es un mensaje automático. No respondas a este correo.');

  // Bloque de contacto de administración (línea adicional en letra pequeña)
  var adminName  = cfgForAdmin_(scope, 'ADMIN_CONTACT_NAME') || '';
  var adminEmail = cfgForAdmin_(scope, 'ADMIN_CONTACT_EMAIL') || '';
  var adminExt   = cfgForAdmin_(scope, 'ADMIN_CONTACT_EXTENSION') || '';
  var contactLine = '';
  if (adminName || adminEmail || adminExt){
    var parts = [];
    if (adminName)  parts.push(escapeHtml_(adminName));
    if (adminEmail) parts.push('<a href="mailto:'+escapeHtml_(adminEmail)+'">'+escapeHtml_(adminEmail)+'</a>');
    if (adminExt)   parts.push('Ext. '+escapeHtml_(adminExt));
    contactLine = '<div style="margin-top:4px;">Contacto administración: '+parts.join(' · ')+'</div>';
  }

  // Botón CTA (opcional)
  var ctaHtml = '';
  if (ctaUrl && ctaLabel){
    ctaHtml = '<tr><td align="center" style="padding: 18px 0 6px 0;">'
      + '<a href="'+ctaUrl+'" style="display:inline-block; padding:12px 18px; background:'+m.color+'; color:#fff; text-decoration:none; border-radius:10px; font-weight:600;">'
      + ctaLabel + '</a></td></tr>';
  }

  // Header: logo + title
  var logoHtml = m.logo
    ? ('<img src="'+m.logo+'" alt="Logo" width="'+m.logoSize+'" style="display:block; max-width:'+m.logoSize+'px; height:auto; border-radius:0"/>')
    : '';

  return ''
    + '<!doctype html><html><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1">'
    + '<title>'+escapeHtml_(title)+'</title>'
    + '</head><body style="margin:0; padding:0; background:'+m.bg+'; color:'+m.text+';">'
    + '<span style="display:none !important; visibility:hidden; opacity:0; color:transparent; height:0; width:0; overflow:hidden;">'+escapeHtml_(preheader)+'</span>'
    + '<table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="background:'+m.bg+';"><tr><td align="center" style="padding:24px 12px;">'
    +   '<table role="presentation" cellpadding="0" cellspacing="0" width="620" style="max-width:620px; width:100%;">'
    +     '<tr><td align="center" style="padding:8px 0 16px 0;">'+logoHtml+'</td></tr>'
    +     '<tr><td align="center" style="font-size:20px; font-weight:700; color:'+m.text+'; padding:0 0 8px 0;">'+escapeHtml_(title)+'</td></tr>'
    +     '<tr><td style="background:'+m.card+'; border:1px solid '+m.border+'; border-radius:14px; padding:18px;">'
    +       '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;">'
    +         inner
    +         ctaHtml
    +       '</table>'
    +     '</td></tr>'
    +     '<tr><td align="center" style="font-size:12px; color:#6b7280; padding:12px 6px;">'
    +       footer
    +       (contactLine ? ('<br/>' + contactLine) : '')
    +     '</td></tr>'
    +   '</table>'
    + '</td></tr></table>'
    + '</body></html>';
}

// Pequeña utilidad para filas con etiqueta:valor (tabla de detalles)
function emailDetailsRows_(pairs){
  // pairs: [ ['Salón', 'ECI - 4to'], ['Fecha', '2025-10-23'], ... ]
  var m = mailCfg_();
  var rows = pairs.map(function(p){
    return '<tr>'
      + '<td style="padding:8px 10px; border-bottom:1px solid '+m.border+'; color:#374151; width:38%; font-weight:600;">'+escapeHtml_(p[0]||'')+'</td>'
      + '<td style="padding:8px 10px; border-bottom:1px solid '+m.border+'; color:'+m.text+';">'+escapeHtml_(p[1]||'')+'</td>'
      + '</tr>';
  }).join('');
  return '<tr><td colspan="1" style="padding:4px 0;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0">'+rows+'</table></td></tr>';
}

// Utilidad para párrafos bloque
function emailParagraph_(html){
  return '<tr><td style="padding:6px 2px; line-height:1.6; color:#374151; font-size:14px;">'+html+'</td></tr>';
}

// Escapar texto simple
function escapeHtml_(s){
  s = String(s==null?'':s);
  return s.replace(/[&<>"]/g, function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c] || c;
  });
}

// ========= Tareas programadas =========
function recipientsJoin_(){
  // Admins + Conserjería (ambos desde Config), separados por ';'
  const A = (cfg_('ADMIN_EMAILS')||'').split(';').map(s=>s.trim()).filter(Boolean);
  const C = (cfg_('CONSERJERIA_EMAILS')||cfg_('CONSERJES_EMAILS')||'').split(';').map(s=>s.trim()).filter(Boolean);
  // Evita duplicados
  const set = {}; [...A, ...C].forEach(x=>{ if(x) set[x.toLowerCase()]=x; });
  return Object.values(set).join(',');
}

function job_Diario_5AM_ActividadesDeHoy(){
  const today = ymd_(new Date());               // YYYY-MM-DD
  const lr = SH_RES.getLastRow(); if (lr<2) return;
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  const cmap = conserjeMap_();
  const hoy = rows
    .filter(r => String(r[2]).toUpperCase()==='APROBADA' && ymdCell_(r[3])===today)
    .map(r => toReservaObj_(r))
    .map(x => ({ ...x, conserje_nombre: (x.conserje_codigo_asignado && cmap[x.conserje_codigo_asignado]?.nombre) || '' }));
  if (!hoy.length) return;

  // Agrupación por salón para lectura rápida
  const bySalon = {};
  hoy.forEach(r => {
    bySalon[r.salon_nombre] = bySalon[r.salon_nombre] || [];
    bySalon[r.salon_nombre].push(r);
  });
  Object.keys(bySalon).forEach(k => bySalon[k].sort((a,b)=> toMin_(a.hora_inicio)-toMin_(b.hora_inicio)));

  // Tabla rica
  const rowsHtml = (list)=> list.map(r=>{
    const host = `${escapeHtml_(r.solicitante_nombre||'')} (${escapeHtml_(r.solicitante_email||'')})`;
    const con  = (r.conserje_requerido==='SI')
                  ? (r.conserje_nombre? `Asignado: ${escapeHtml_(r.conserje_nombre)}` : 'Requerido')
                  : 'N/A';
    return `
      <tr>
        <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};white-space:nowrap">${fmt12_(r.hora_inicio)} – ${fmt12_(r.hora_fin)}</td>
        <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};">${escapeHtml_(r.evento_nombre||'')}</td>
        <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};">${host}</td>
        <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border}; text-align:center">${String(r.cant_personas||0)}</td>
        <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border}; text-align:center">${escapeHtml_(r.publico_tipo||'')}</td>
        <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};">${con}</td>
      </tr>`;
  }).join('');

  let blocks = '';
  Object.keys(bySalon).sort().forEach(salon=>{
    blocks += `
      <tr><td style="padding:12px 0 4px 0;">
        <div style="font-weight:700;color:${mailCfg_().text};font-size:14px;">${escapeHtml_(salon)}</div>
      </td></tr>
      <tr><td>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate">
          <tr style="background:#fafafa">
            <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};font-weight:600">Hora</td>
            <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};font-weight:600">Evento</td>
            <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};font-weight:600">Anfitrión</td>
            <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};font-weight:600;text-align:center">Asist.</td>
            <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};font-weight:600;text-align:center">Público</td>
            <td style="padding:8px 10px;border-bottom:1px solid ${mailCfg_().border};font-weight:600">Conserje</td>
          </tr>
          ${rowsHtml(bySalon[salon])}
        </table>
      </td></tr>`;
  });

  const base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();
  const ctaUrl = normalizeExecUrl_(base);
  const title  = `Agenda de hoy · ${fmtDMY_(today)}`;
  const pre    = 'Reservas aprobadas y datos logísticos para el día';

  const html = emailLayout_({
    title: title,
    preheader: pre,
    htmlInner:
      emailParagraph_(`Resumen del día <b>${fmtDMY_(today)}</b>. A continuación, el detalle por salón.`) +
      blocks,
    ctaUrl: ctaUrl,
    ctaLabel: 'Abrir sistema',
    footer: (function(){
      const admin = {
        name:  cfg_('ADMIN_CONTACT_NAME') || 'Administración',
        email: cfg_('ADMIN_CONTACT_EMAIL') || '',
        ext:   cfg_('ADMIN_CONTACT_EXTENSION') || ''
      };
      return `Reporte generado automáticamente a las 5:00 a. m.<br/>`;
    })()
  });

  const body = hoy.map(r => `• ${fmt12_(r.hora_inicio)}-${fmt12_(r.hora_fin)} · ${r.salon_nombre} · ${r.evento_nombre||''} · ${r.solicitante_nombre} (${r.solicitante_email}) · Asist:${r.cant_personas||0} · Cons:${r.conserje_requerido==='SI'?(r.conserje_nombre?('Asignado:'+r.conserje_nombre):'Requerido'):'N/A'}`).join('\n');

  const to = recipientsJoin_();
  if (!to) return;
  MailApp.sendEmail({
    to,
    subject: `Reserva de Salones | Agenda de hoy – ${fmtDMY_(today)}`,
    htmlBody: html,
    body,
    name: cfgForAdmin_('1', 'MAIL_SENDER_NAME') || 'INFOTEP - Reserva de Salones',
    replyTo: cfgForAdmin_('1', 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
  });
}

function job_Diario_7AM_Recordatorios(){
  // Eventos de "mañana"
  const d = new Date(); d.setDate(d.getDate()+1);
  const target = ymd_(d);
  const lr = SH_RES.getLastRow(); if (lr<2) return;
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  const list = rows
    .filter(r => String(r[2]).toUpperCase()==='APROBADA' && ymdCell_(r[3])===target)
    .map(r => toReservaObj_(r));
  if (!list.length) return;

  const base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();
  const appUrl = normalizeExecUrl_(base);

  list.forEach(r => {
    const cancelUrl = appUrl + `?cancel=${encodeURIComponent(r.token)}`;
    const subj = `Recordatorio – ${r.salon_nombre} – ${fmtDMY_(r.fecha)} ${fmt12_(r.hora_inicio)}`;
    const adminId = normalizeAdminId_(r.administracion_id||'1');
    const html = emailLayout_({
      title: 'Recordatorio de reserva',
      preheader: `Mañana tienes reservada ${r.salon_nombre}.`,
      htmlInner:
        emailParagraph_(`Hola <b>${escapeHtml_(r.solicitante_nombre||'')}</b>,`) +
        emailParagraph_(`Mañana tienes reservada la sala <b>${escapeHtml_(r.salon_nombre)}</b> entre <b>${fmt12_(r.hora_inicio)}–${fmt12_(r.hora_fin)}</b> para “${escapeHtml_(r.evento_nombre||'') }”.`) +
        emailDetailsRows_([
          ['Fecha', fmtDMY_(r.fecha)],
          ['Salón', r.salon_nombre],
          ['Horario', `${fmt12_(r.hora_inicio)} – ${fmt12_(r.hora_fin)}`]
        ]) +
        emailParagraph_(`¿Aún necesitas esta reserva? Si no es así, por favor <a href="${cancelUrl}">cancela la reserva</a> para liberar el espacio.`),
      ctaUrl: appUrl,
      ctaLabel: 'Abrir sistema',
      footer: (function(){
        const admin = {
          name:  cfgForAdmin_(adminId, 'ADMIN_CONTACT_NAME') || 'Administración',
          email: cfgForAdmin_(adminId, 'ADMIN_CONTACT_EMAIL') || '',
          ext:   cfgForAdmin_(adminId, 'ADMIN_CONTACT_EXTENSION') || ''
        };
        return `Este recordatorio se envía un día antes del evento.<br/>`;
      })(),
      adminId
    });

    MailApp.sendEmail({
      to: r.solicitante_email,
      subject: subj,
      htmlBody: html,
      name: cfgForAdmin_(adminId, 'MAIL_SENDER_NAME') || 'INFOTEP - Reserva de Salones',
      replyTo: cfgForAdmin_(adminId, 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
    });
  });
}

// ===== PREVISUALIZACIONES / PRUEBAS MANUALES =====
// Ejecuta estos helpers desde el editor para validar contenido sin esperar al trigger:
// - dbg_Send_5AM_Preview('2025-10-29', 'tu@correo');   // fecha opcional (por defecto hoy)
// - dbg_Send_7AM_Preview('2025-10-29', 'tu@correo');   // fecha base (envía recordatorios para el día siguiente)
function dbg_Send_5AM_Preview(fechaYMD, toOverride){
  const savedYmd = fechaYMD || ymd_(new Date());
  const lr = SH_RES.getLastRow(); if (lr<2){ Logger.log('No hay reservas'); return; }
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  const cmap = conserjeMap_();
  const list = rows.filter(r => String(r[2]).toUpperCase()==='APROBADA' && ymdCell_(r[3])===savedYmd)
                   .map(r => toReservaObj_(r))
                   .map(x => ({ ...x, conserje_nombre: (x.conserje_codigo_asignado && cmap[x.conserje_codigo_asignado]?.nombre) || '' }));
  if (!list.length){ Logger.log('Sin reservas para %s', savedYmd); return; }
  // Aprovecha la misma función armando correo real:
  job_Diario_5AM_ActividadesDeHoy();
  if (toOverride){
    // Reenvío de prueba al correo indicado
    MailApp.sendEmail({
      to: toOverride,
      subject: '[PREVIEW] Agenda de hoy – ' + fmtDMY_(savedYmd),
      htmlBody: 'Se ejecutó job_Diario_5AM_ActividadesDeHoy(). Revisa tu bandeja (destinatarios reales).',
    });
  }
}
function dbg_Send_7AM_Preview(baseYMD, toOverride){
  const d = baseYMD ? toDate_(baseYMD) : new Date();
  d.setDate(d.getDate()+1);
  const target = ymd_(d);
  const lr = SH_RES.getLastRow(); if (lr<2){ Logger.log('No hay reservas'); return; }
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues();
  const list = rows.filter(r => String(r[2]).toUpperCase()==='APROBADA' && ymdCell_(r[3])===target).map(r => toReservaObj_(r));
  if (!list.length){ Logger.log('Sin reservas para %s', target); return; }
  // Ejecuta job real
  job_Diario_7AM_Recordatorios();
  if (toOverride){
    MailApp.sendEmail({
      to: toOverride,
      subject: '[PREVIEW] Recordatorios – ' + fmtDMY_(target),
      htmlBody: 'Se ejecutó job_Diario_7AM_Recordatorios(). Revisa la bandeja de los destinatarios reales.',
    });
  }
}

// ========= Setup =========
function setupSheetsAndConfig(){
  const must = [
    {name:'Config', headers:['key','value']},
    {name:'Usuarios', headers:['email','nombre','departamento','rol','prioridad','prioridad_salones','estado','extension']},
    {name:'Conserjes', headers:['codigo','nombre','email','telefono','activo'] },
    {name:'Salones', headers:['id','salon_nombre','capacidad_max','habilitado','sede','restriccion']},
    {name:'Reservas', headers:[
      'id','token','estado','fecha','hora_inicio','hora_fin','salon_id','salon_nombre','cant_personas','solicitante_email','solicitante_nombre','departamento','extension','evento_nombre','publico_tipo','prioridad','conserje_requerido','conserje_notificado','creado_en','actualizado_en','cancelado_por','cancelado_motivo','conserje_codigo_asignado'
    ]},
  ];
  const ss = SpreadsheetApp.getActive();
  must.forEach(s=>{
    let sh = ss.getSheetByName(s.name);
    if (!sh){
      sh = ss.insertSheet(s.name);
      sh.getRange(1,1,1,s.headers.length).setValues([s.headers]);
    } else if (sh.getLastRow()<1){
      sh.getRange(1,1,1,s.headers.length).setValues([s.headers]);
    } else {
      const maxCols = Math.max(sh.getLastColumn(), s.headers.length);
      const current = sh.getRange(1,1,1,maxCols).getValues()[0];
      let needsUpdate = false;
      if (current.length < s.headers.length) {
        needsUpdate = true;
      } else {
        for (let i=0;i<s.headers.length;i++){
          if (String(current[i]||'').trim() !== s.headers[i]){ needsUpdate = true; break; }
        }
      }
      if (needsUpdate){
        sh.getRange(1,1,1,s.headers.length).setValues([s.headers]);
      }
    }
  });
  SpreadsheetApp.flush();
  return 'OK';
}

function installTriggers(){
  // borra existentes del proyecto
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // 5:00 a.m. santo domingo
  ScriptApp.newTrigger('job_Diario_5AM_ActividadesDeHoy').timeBased().atHour(5).everyDays(1).create();
  // 7:00 a.m. recordatorios del día siguiente
  ScriptApp.newTrigger('job_Diario_7AM_Recordatorios').timeBased().atHour(7).everyDays(1).create();
}

// ========= Logo helpers (reuso del otro proyecto) =========
function getLogoDataUrl(size){
  const S = Math.max(48, Number(size)||128);
  const cache = CacheService.getScriptCache();
  const key = 'LOGO_DATA_URL_'+S;
  const cached = cache.get(key);
  if (cached) return cached;

  const id = getLogoFileId_();
  if (!id) return '';

  const blob = getDriveThumbnailBlob_(id, S*3);
  if (!blob) return '';

  const mime = blob.getContentType() || 'image/png';
  const b64 = Utilities.base64Encode(blob.getBytes());
  const dataUrl = 'data:'+mime+';base64,'+b64;
  cache.put(key, dataUrl, 6*60*60);
  return dataUrl;
}
function getLogoFileId_(){
  let id = cfg_('LOGO_FILE_ID');
  if (!id){
    const u = getLogoUrl();
    const m = u && u.match(/id=([a-zA-Z0-9_-]+)/); id = m? m[1] : '';
  }
  return id||'';
}
function getDriveThumbnailBlob_(fileId, size){
  // Requiere Drive API avanzada habilitada. Si no está disponible, esta función
  // devolverá null y el sistema usará el fallback de getLogoUrl() (o sin logo).

  if (!fileId) return null;
  try{
    const file = Drive.Files.get(fileId, { fields:'thumbnailLink' });
    let url = file && file.thumbnailLink; if (!url) return null;
    url = url.replace(/=s\d+$/, '=s'+Math.max(32, Math.min(size||256, 1600)));
    const resp = UrlFetchApp.fetch(url, { headers:{ Authorization:'Bearer '+ScriptApp.getOAuthToken() }, muteHttpExceptions:true });
    if (resp.getResponseCode()!==200) return null;
    const blob = resp.getBlob();
    const ct = blob.getContentType();
    return /png|jpeg|jpg/i.test(ct) ? blob : blob.setContentType('image/png');
  }catch(e){ return null; }
}
function getLogoUrl(){
  const cfgId = cfg_('LOGO_FILE_ID');
  if (cfgId){ try{ const f=DriveApp.getFileById(cfgId); return 'https://drive.google.com/uc?export=view&id='+f.getId(); }catch(e){} }
  try{ const any = DriveApp.searchFiles('title = "logo.png" and trashed = false'); if (any.hasNext()){ const f = any.next(); return 'https://drive.google.com/uc?export=view&id='+f.getId(); }}catch(e){}
  return '';
}

// ========= Usuarios (Admin CRUD) =========
function apiListUsuarios(){
  const me = getUser_();
  if (!isGeneralAdminUser_(me)) return { ok:false, msg:'No autorizado' };
  const lr = SH_USR.getLastRow(); if (lr<2) return { ok:true, data:[] };
  const vals = SH_USR.getRange(2,1,lr-1,9).getValues(); // email,nombre,departamento,rol,prioridad,prioridad_salones,estado,extension,admin
  const data = vals.map(r=>{
    const prioSalonesRaw = String(r[5]||'').trim();
    return {
      email:String(r[0]||'').toLowerCase(),
      nombre:String(r[1]||''),
      departamento:String(r[2]||''),
      rol:String(r[3]||'').toUpperCase(),
      prioridad:Number(r[4]||0),
      prioridad_salones: prioSalonesRaw,
      prioridad_salones_list: parsePrioridadSalones_(prioSalonesRaw),
      estado:String(r[6]||'').toUpperCase(),
      extension:String(r[7]||''),
      administracion_id: normalizeAdminId_(r[8]||'1')
    };
  });
  return { ok:true, data };
}

function apiUpsertUsuario(u){
  const me = (Session.getActiveUser().getEmail()||'');
  const meUser = getUser_();
  if (!isAdminEmail_(me) || !isGeneralAdminUser_(meUser)) return { ok:false, msg:'No autorizado' };
  if (!u || !u.email) return { ok:false, msg:'Email requerido' };
  const email = String(u.email).toLowerCase().trim();
  const nombre = String(u.nombre||'').trim();
  const dep = String(u.departamento||'').trim();
  const rol = String(u.rol||'').toUpperCase().trim();
  const prio = Number(u.prioridad||0);
  const adminId = normalizeAdminId_(u.administracion_id || u.admin_id || '1');
  let prioSalonesInput = '';
  if (Array.isArray(u.prioridad_salones)){
    prioSalonesInput = u.prioridad_salones.join(';');
  } else if (Array.isArray(u.prioridadSalones)){
    prioSalonesInput = u.prioridadSalones.join(';');
  } else if (typeof u.prioridad_salones === 'string'){
    prioSalonesInput = u.prioridad_salones;
  } else if (typeof u.prioridadSalones === 'string'){
    prioSalonesInput = u.prioridadSalones;
  }
  const prioSalonesList = parsePrioridadSalones_(prioSalonesInput);
  const prioSalonesRaw = prioSalonesList.join(';');
  const est = String(u.estado||'').toUpperCase().trim();
  const ext = String(u.extension||'').trim();
  const lr = SH_USR.getLastRow();
  if (lr>=2){
    const vals = SH_USR.getRange(2,1,lr-1,9).getValues();
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][0]).toLowerCase().trim()===email){
        SH_USR.getRange(i+2,2,1,8).setValues([[nombre,dep,rol,prio,prioSalonesRaw,est,ext,adminId]]);
        return { ok:true, updated:true };
      }
    }
  }
  SH_USR.appendRow([email,nombre,dep,rol,prio,prioSalonesRaw,est,ext,adminId]);
  return { ok:true, created:true };
}

// ========= API conserjes =========

function apiListConserjes(){
  const me = getUser_();
  if (!isAdminEmail_((me && me.email)||'')) return { ok:false, msg:'No autorizado' };
  if (!SH_CON) return { ok:true, data:[] };
  const lr = SH_CON.getLastRow(); if (lr<2) return { ok:true, data:[] };
  const rows = SH_CON.getRange(2,1,lr-1,5).getValues(); // codigo,nombre,email,telefono,activo
  const data = rows.map(r => ({
    codigo: String(r[0]||''),
    nombre: String(r[1]||''),
    email:  String(r[2]||''),
    telefono: String(r[3]||''),
    activo: String(r[4]||'').toUpperCase()==='SI'
  }));
  return { ok:true, data };
}

function apiAddConserje(nombre, email, telefono){
  const me = (Session.getActiveUser().getEmail()||'');
  const user = getUser_();
  if (!isAdminEmail_(me) || !isGeneralAdminUser_(user)) return { ok:false, msg:'No autorizado' };
  if (!SH_CON) return { ok:false, msg:'Hoja Conserjes no existe' };
  const codigo = nextId_(SH_CON, 'C-');
  SH_CON.appendRow([codigo, nombre||'', (email||'').toLowerCase(), telefono||'', 'SI']);
  return { ok:true, codigo };
}

function apiDeleteConserje(codigo){
  const me = (Session.getActiveUser().getEmail()||'');
  const user = getUser_();
  if (!isAdminEmail_(me) || !isGeneralAdminUser_(user)) return { ok:false, msg:'No autorizado' };
  const lr = SH_CON.getLastRow(); if (lr<2) return { ok:false, msg:'No hay conserjes' };
  const vals = SH_CON.getRange(2,1,lr-1,5).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0])===String(codigo)){
      // Borrado físico o lógico. Propongo lógico:
      SH_CON.getRange(i+2,5).setValue('NO'); // activo = NO
      return { ok:true };
    }
  }
  return { ok:false, msg:'Conserje no encontrado' };
}

// === actualizar datos de un conserje existente ===
// fields = { nombre?, email?, telefono?, activo? }
function apiUpdateConserje(codigo, fields){
  const me = (Session.getActiveUser().getEmail()||'');
  const user = getUser_();
  if (!isAdminEmail_(me) || !isGeneralAdminUser_(user)) return { ok:false, msg:'No autorizado' };
  if (!SH_CON) return { ok:false, msg:'Hoja Conserjes no existe' };
  codigo = String(codigo||'').trim();
  if (!codigo) return { ok:false, msg:'Código inválido' };
  const lr = SH_CON.getLastRow(); if (lr<2) return { ok:false, msg:'No hay conserjes' };
  const vals = SH_CON.getRange(2,1,lr-1,5).getValues(); // codigo,nombre,email,telefono,activo
  const rowIdx = vals.findIndex(r => String(r[0]||'') === codigo);
  if (rowIdx === -1) return { ok:false, msg:'Conserje no encontrado' };

  // Estado actual
  const cur = {
    codigo: String(vals[rowIdx][0]||''),
    nombre: String(vals[rowIdx][1]||''),
    email:  String(vals[rowIdx][2]||''),
    telefono: String(vals[rowIdx][3]||''),
    activo: String(vals[rowIdx][4]||'').toUpperCase()==='SI'
  };
  const next = {
    nombre:   (fields && 'nombre'   in fields) ? String(fields.nombre||'') : cur.nombre,
    email:    (fields && 'email'    in fields) ? String(fields.email||'').toLowerCase() : cur.email,
    telefono: (fields && 'telefono' in fields) ? String(fields.telefono||'') : cur.telefono,
    activo:   (fields && 'activo'   in fields) ? !!fields.activo : cur.activo
  };

  // ¿Cambios?
  const changed = (
    cur.nombre   !== next.nombre ||
    cur.email    !== next.email  ||
    cur.telefono !== next.telefono ||
    cur.activo   !== next.activo
  );
  if (!changed) return { ok:true, changed:false };

  // Persistir solo columnas tocadas
  const r = rowIdx + 2; // compensar header
  if (cur.nombre !== next.nombre)     SH_CON.getRange(r, 2).setValue(next.nombre);
  if (cur.email  !== next.email)      SH_CON.getRange(r, 3).setValue(next.email);
  if (cur.telefono !== next.telefono) SH_CON.getRange(r, 4).setValue(next.telefono);
  if (cur.activo !== next.activo)     SH_CON.getRange(r, 5).setValue(next.activo ? 'SI' : 'NO');
  SpreadsheetApp.flush();
  return { ok:true, changed:true };
}

function getConserjeConflicts_(conserjeCodigo, fecha, ini, fin){
  const lr = SH_RES.getLastRow(); if (lr<2) return [];
  const rows = SH_RES.getRange(2,1,lr-1,24).getValues(); // incluye nueva col 24: conserje_codigo_asignado + admin
  const iniMin = toMin_(normHHMM_(ini));
  const finMin = toMin_(normHHMM_(fin));
  if (isNaN(iniMin) || isNaN(finMin)) return [];

  return rows
    .filter(r => String(r[2]).toUpperCase()==='APROBADA'
              && ymdCell_(r[3]) === String(fecha).slice(0,10)
              && String(r[22]||'').trim() === String(conserjeCodigo).trim()) // col 23 index 22
    .filter(r => {
      const rIniMin = toMin_(hhmmFromCell_(r[4]));
      const rFinMin = toMin_(hhmmFromCell_(r[5]));
      if (isNaN(rIniMin) || isNaN(rFinMin)) return false;
      return !(finMin <= rIniMin || iniMin >= rFinMin);
    })
    .map(r => ({ id:r[0], fecha: ymdCell_(r[3]), inicio: hhmmFromCell_(r[4]), fin: hhmmFromCell_(r[5]) }));
}

function apiAsignarConserje(reservaId, conserjeCodigo){
  const me = (Session.getActiveUser().getEmail()||'');
  const user = getUser_();
  if (!isAdminEmail_(me)) return { ok:false, msg:'No autorizado' };

  // Lee reserva
  const r = getReservaById_(reservaId);
  if (!r) return { ok:false, msg:'Reserva no encontrada' };
  const scope = adminScopeForUser_(user);
  if (scope !== '1' && normalizeAdminId_(r.administracion_id||'1') !== scope){
    return { ok:false, msg:'No autorizado para este salón' };
  }
  const adminId = normalizeAdminId_(r.administracion_id||'1');
  if (String(r.estado).toUpperCase()!=='APROBADA') return { ok:false, msg:'La reserva no está aprobada' };
  // No permitir asignar si ya pasó o faltan <30 min
  const start = toDate_(r.fecha);
  const [hh,mi] = (r.hora_inicio||'00:00').split(':').map(Number);
  start.setHours(hh||0, mi||0, 0, 0);
  const now = new Date();
  if (now.getTime() > (start.getTime() - 30*60*1000)) {
    return {ok:false, msg:'Fuera de ventana: asignación solo hasta 30 min antes del inicio.'};
  }
  if (String(r.conserje_requerido||'') !== 'SI') return { ok:false, msg:'Esta reserva no requiere conserje' };

  // Valida conserje
  const list = apiListConserjes();
  const c = (list.data||[]).find(x => x.codigo===String(conserjeCodigo) && x.activo);
  if (!c) return { ok:false, msg:'Conserje no válido o inactivo' };

  // Conflictos del conserje
  const conf = getConserjeConflicts_(c.codigo, r.fecha, r.hora_inicio, r.hora_fin);
  if (conf.length){
    return { ok:false, msg:'Ese conserje ya está asignado en ese horario' };
  }

  // Persistir asignación
  setReservaField_(r.id, 'conserje_codigo_asignado', c.codigo);

  // Email al conserje
  try{
    const subj = `Asignación de conserje – ${r.salon_nombre} – ${fmtDMY_(r.fecha)} ${fmt12_(r.hora_inicio)}`;
    const html = emailLayout_({
      title: 'Has sido asignado(a) a un evento',
      preheader: `Evento en ${r.salon_nombre} el ${r.fecha} ${r.hora_inicio}.`,
      htmlInner:
        emailParagraph_(`Hola <b>${escapeHtml_(c.nombre||'')}</b>,`) +
        emailParagraph_('Fuiste asignado(a) como conserje para el siguiente evento:') +
        emailDetailsRows_([
          ['Salón', r.salon_nombre],
          ['Fecha', fmtDMY_(r.fecha)],
          ['Hora', fmt12_(r.hora_inicio) + ' – ' + fmt12_(r.hora_fin)],
          ['Evento', r.evento_nombre||'']
        ]),
      footer: 'Gracias por tu apoyo logístico.',
      adminId
    });

    MailApp.sendEmail({
      to: c.email,
      subject: subj,
      htmlBody: html,
      name: cfgForAdmin_(adminId, 'MAIL_SENDER_NAME') || 'INFOTEP - Reserva de Salones',
      replyTo: cfgForAdmin_(adminId, 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
    });
  }catch(e){
    // si el email falla, no revertimos la asignación
  }

  return { ok:true, conserje: { codigo:c.codigo, nombre:c.nombre, email:c.email } };
}

// ========= Helpers generales =========
function apiSolicitarAcceso(nombre, departamento, extension){
  try{
    const email = (Session.getActiveUser().getEmail()||'').toLowerCase().trim();
    if (!email) return { ok:false, msg:'No se pudo identificar tu correo institucional.' };

    nombre = String(nombre||'').trim();
    departamento = String(departamento||'').trim();
    extension = String(extension||'').trim();

    if (!nombre || !departamento || !/^\d+$/.test(extension)){
      return { ok:false, msg:'Completa todos los campos. La extensión debe ser solo dígitos.' };
    }

    const lr = SH_USR.getLastRow();
    // buscamos fila existente
    if (lr >= 2){
      const vals = SH_USR.getRange(2,1,lr-1,9).getValues();
      for (let i=0;i<vals.length;i++){
        if (String(vals[i][0]).trim().toLowerCase() === email){
          // Actualiza nombre, departamento, extension y estado -> PENDIENTE
          SH_USR.getRange(i+2,2).setValue(nombre);
          SH_USR.getRange(i+2,3).setValue(departamento);
          // deja rol/prioridad como están (admin los define)
          SH_USR.getRange(i+2,7).setValue('PENDIENTE'); // estado (col 7 = G)
          SH_USR.getRange(i+2,8).setValue(extension);   // extension (col 8 = H)
          SpreadsheetApp.flush();
          try{ notifyAdminsNuevaSolicitud_(email, nombre, departamento, extension); }catch(e){}
          return { ok:true, updated:true };
        }
      }
    }

    // Si no existe, append con estado PENDIENTE
    // Si faltan columnas, Apps Script las crea al vuelo
    SH_USR.appendRow([email, nombre, departamento, '', 0, '', 'PENDIENTE', extension, '1']);
    SpreadsheetApp.flush();
    try{ notifyAdminsNuevaSolicitud_(email, nombre, departamento, extension); }catch(e){}
    return { ok:true, created:true };
  }catch(e){
    return { ok:false, msg:String(e && e.message || e) };
  }
}

function notifyAdminsNuevaSolicitud_(email, nombre, departamento, extension){
  const adminId = '1';
  const to = (cfgForAdmin_(adminId, 'ADMIN_EMAILS')||'').split(';').map(s=>s.trim()).filter(Boolean).join(',');
  if (!to) return;
  const subj = 'Nueva solicitud de acceso – Reserva de Salones';
  const base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();
  const appUrl = normalizeExecUrl_(base);
  const html = emailLayout_({
    title:'Nueva solicitud de acceso',
    preheader:`${nombre} solicita acceso`,
    htmlInner:
      emailDetailsRows_([
        ['Correo', email],
        ['Nombre', nombre],
        ['Departamento', departamento],
        ['Extensión', extension]
      ]) +
      emailParagraph_('Asigne <b>rol</b> y <b>prioridad</b> y cambie el <b>estado</b> a <b>ACTIVO</b> para habilitar al usuario.'),
    ctaUrl: appUrl,
    ctaLabel: 'Abrir sistema',
    footer:'Este mensaje se envía automáticamente al registrar una solicitud de acceso.',
    adminId
  });
  MailApp.sendEmail({
    to, subject:subj, htmlBody:html,
    name: cfgForAdmin_(adminId, 'MAIL_SENDER_NAME') || 'Reserva de Salones',
    replyTo: cfgForAdmin_(adminId, 'MAIL_REPLY_TO') || Session.getActiveUser().getEmail()
  });
}

function conserjeMap_(){
  const map = {};
  if (!SH_CON) return map;
  const lr = SH_CON.getLastRow(); if (lr<2) return map;
  const vals = SH_CON.getRange(2,1,lr-1,5).getValues();
  vals.forEach(r=>{
    const activo = String(r[4]||'').toUpperCase()==='SI';
    if (activo){
      map[String(r[0]||'')] = { nombre:String(r[1]||''), email:String(r[2]||'') };
    }
  });
  return map;
}

function isNumericStr_(s){ return /^\d+$/.test(String(s||'').trim()); }

function apiListMisReservasByEmail(fechaDesde, fechaHasta, email){
  const who = String(email||'').toLowerCase().trim();
  if (!who) return {ok:true, data:[]};

  const lr = SH_RES.getLastRow(); if (lr<2) return {ok:true, data:[]};
  const lc = Math.max(24, SH_RES.getLastColumn());              // 👈 asegura leer col 24
  const rows = SH_RES.getRange(2,1,lr-1,lc).getValues();

  // Parse local, sin desplazamiento UTC (evita off-by-one)
  const d1 = fechaDesde ? toDate_(fechaDesde) : new Date(2000,0,1);
  const d2 = fechaHasta ? toDate_(fechaHasta) : new Date(2100,0,1);
  const d1n = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate());
  const d2n = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());

  let data = rows
    .filter(r => String(r[9]||'').toLowerCase().trim() === who) // solicitante_email
    .filter(r => {
      const d = new Date(r[3]); // fecha
      if (isNaN(d)) return false;
      const dn = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      return dn >= d1n && dn <= d2n;
    })
    .map(r => toReservaObj_(r));

    // Enriquecer con nombre de conserje (si activo)
    const cmap = conserjeMap_();
    data = data.map(x => ({
      ...x,
      conserje_nombre: (x.conserje_codigo_asignado && cmap[x.conserje_codigo_asignado]?.nombre) || ''
    }));

  Logger.log('MIS: email=%s d1=%s d2=%s matches=%s', who, d1n, d2n, data.length);
  return {ok:true, data};
}

function maintenance_NormalizarHoras(){
  const sh = SH_RES; if (!sh) return 0;
  const lr = sh.getLastRow(); if (lr<2) return 0;
  const vals = sh.getRange(2,1,lr-1,24).getValues();
  let changes = 0;
  for (let i=0;i<vals.length;i++){
    const oldIni = vals[i][4], oldFin = vals[i][5];
    const newIni = normHHMM_(oldIni);
    const newFin = normHHMM_(oldFin);
    if (newIni && newIni !== oldIni){ sh.getRange(i+2,5).setValue(newIni); changes++; }
    if (newFin && newFin !== oldFin){ sh.getRange(i+2,6).setValue(newFin); changes++; }
  }
  SpreadsheetApp.flush();
  return changes;
}

function dbg_AvailabilityDump(fechaStr, salonId, duracionMin){
  // Ejecuta la API real para ver qué está devolviendo el servidor
  const res = apiListDisponibilidad(fechaStr, salonId, duracionMin);

  try{
    const sh = SpreadsheetApp.getActive().getSheetByName('Reservas');
    const lr = sh ? sh.getLastRow() : 0;
    const raw = (lr >= 2) ? sh.getRange(2, 1, lr - 1, 22).getValues() : [];

    Logger.log('DBG: fecha=%s salon=%s dur=%s', fechaStr, salonId, duracionMin);
    Logger.log('DBG: apiListDisponibilidad.ok=%s count=%s', res && res.ok, res && res.data && res.data.length);

    if (res && Array.isArray(res.data)){
      const sample = res.data.slice(0, 8).map(s => s.inicio + '-' + s.fin + ' ' + (s.disponible ? 'libre' : 'Ocup.'));
      Logger.log('DBG: slots(sample)=%s', JSON.stringify(sample));
      const occ = res.data.filter(s => !s.disponible).length;
      Logger.log('DBG: ocupados_contados=%s', occ);
    }

    // Reservas crudas del día/salón (usando ymdCell_ para celdas tipo Date)
    const listCruda = raw
      .filter(r =>
        String(r[2]).toUpperCase() === 'APROBADA' &&
        ymdCell_(r[3]) === String(fechaStr).slice(0, 10) &&
        String(r[6]).trim() === String(salonId).trim()
      )
      .map(r => ({
        id: String(r[0]),
        fecha: ymdCell_(r[3]),            // << muestra la fecha normalizada
        ini: hhmmFromCell_(r[4]),
        fin: hhmmFromCell_(r[5]),
        salon_id: String(r[6]),
        email: String(r[9]),
        prio: Number(r[15] || 0)
      }));

    Logger.log('DBG: crudo_reservas_dia_salon=%s', JSON.stringify(listCruda));

    return {
      ok: true,
      dump: {
        req:  { fechaStr, salonId, duracionMin },
        crudo: listCruda,
        out: res
      }
    };
  } catch (e){
    Logger.log('DBG ERROR: %s', e && e.message);
    return { ok:false, msg: String(e && e.message || e) };
  }
}

function dbg_CheckCreate(fechaStr, salonId, horaInicio, duracionMin){
  const me = getUser_();
  const hIni = normHHMM_(horaInicio);
  const DUR  = Math.max(Number(duracionMin||0), Number(cfg_('DURATION_MIN')||30));
  const hFin = addMinutes_(hIni, DUR);
  const conflicts = getConflicts_(fechaStr, salonId, hIni, hFin);
  const prio = effectivePriorityForSalon_(me, salonId);
  const maxPrio = conflicts.reduce((m,c)=>Math.max(m, Number(c.prioridad||0)),0);
  return {
    ok:true,
    me:{ email:me.email, prio },
    req:{ fechaStr, salonId, hIni, DUR, hFin },
    conflicts,
    wouldCreate: (conflicts.length===0) || (prio>maxPrio),
    reason: (conflicts.length===0)?'no_conflicts':(prio>maxPrio?'prio_override':'blocked_low_prio')
  };
}

