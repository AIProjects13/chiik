/* ==========================================================================
   Granjas · Frontend (ES Modules, sin frameworks)
   --------------------------------------------------------------------------
    1. Configuración           7. Navegación y estructura
    2. Estado                  8. Vistas: dashboard y registros
    3. Utilidades              9. Vistas: granjas, perfiles y usuarios
    4. API (backend)          10. Vistas: listas, bitácora y seguridad
    5. Datos y agregaciones   11. Modales y formularios
    6. Tablas y gráficos      12. Acciones, autenticación y arranque
   La interfaz solo oculta lo que cada rol no puede usar: los permisos reales
   (qué granjas ve cada quien) los aplica el servidor en cada solicitud.
   ========================================================================== */

import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import {
    initializeAuth, browserSessionPersistence, onAuthStateChanged, signInWithEmailAndPassword,
    signOut, sendPasswordResetEmail, updatePassword,
} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';
import { ico, hydrateIcons } from './icons.js';
import { mountChart, mountHeatmap, exportChartPng, downloadBlob, PALETTE } from './charts.js';
import { crearXlsx, leerXlsx } from './excel.js';
import { leerFormatoSemanal, agruparPorLote, numeroDeNave } from './formato-semanal.js';

// ==========================================================================
// 1. CONFIGURACIÓN
// ==========================================================================

const CONFIG = {
    // Analytics no se usa (no se envían datos a terceros).
    firebase: {
        apiKey: 'AIzaSyAMXTdCrgzvyyFxsByJ6Byl4Cooq1M6wRI',
        authDomain: 'granjas-c1fd1.firebaseapp.com',
        projectId: 'granjas-c1fd1',
        messagingSenderId: '946552384503',
        appId: '1:946552384503:web:c7f2e8dac34740b2631cf2',
    },
    // URL de la implementación "Aplicación web" de Apps Script (termina en /exec).
    API_URL: 'https://script.google.com/macros/s/AKfycbxWfN28IuaBKC4lZ0KedxJB55j1Ixr2uS-ExaEqEzozjfxbROMl3jbfK4_m6tahjo2n/exec',
    // Clave del sitio de la verificación "No soy un robot" (vacía = sin verificación). Ver docs/guias/publicar-fase-15.md.
    TURNSTILE_SITE_KEY: '',
    INACTIVIDAD_MIN: 30,        // cierre automático de sesión por inactividad
    MIN_CLAVE: 10,              // longitud mínima de la contraseña definitiva
    MIN_JUSTIFICACION: 10,
    RANGO_INICIAL_DIAS: 90,
    MAX_FILAS_TABLA: 1000,      // filas dibujadas por tabla (el Excel incluye todas)
};

const ROLES = { ADMIN: 'ADMIN', ADMIN_EMPRESA: 'ADMIN_EMPRESA', REPRESENTANTE: 'REPRESENTANTE' };
const ROLE_LABEL = { ADMIN: 'Plataforma', ADMIN_EMPRESA: 'Administrador de empresa', REPRESENTANTE: 'Representante' };
const LINEAS = ['Reproductora', 'Incubadora', 'Engorde', 'Postura'];
const LINEA_ICON = { Reproductora: 'bird', Incubadora: 'archive', Engorde: 'drumstick', Postura: 'egg' };
const SECCION_ICON = {
    'Información de origen': 'globe',
    'Información de incubadora': 'egg',
    'Programa de vacunación incubadora': 'syringe',
    'Información granja': 'warehouse',
};
const LISTA_LABEL = {
    Pais: 'Países', Raza: 'Razas', SistemaAlojamiento: 'Sistemas de alojamiento',
    EmpresaIncubadora: 'Empresas incubadoras', Vacuna: 'Vacunas',
    GeneticaPostura: 'Líneas genéticas · postura', GeneticaEngorde: 'Líneas genéticas · engorde', Sigla: 'Siglas (se conservan en mayúsculas en los nombres)',
};
const PRESETS = [[30, '30 días'], [90, '90 días'], [180, '6 meses'], [365, '1 año'], [0, 'Todo']];
const AGRUPAR = [['dia', 'Día'], ['semana', 'Semana'], ['mes', 'Mes']];
/** Pesos: el servidor guarda gramos; aquí solo se convierte para escribir y mostrar. */
const UNIDADES = { g: 1, kg: 1000, lb: 453.59237 };
const UNIDAD_NOMBRE = { g: 'gramos', kg: 'kilogramos', lb: 'libras' };
/** Colores del índice de estrés calórico (nivel → color). */
const IEC_COLOR = { verde: '#3FBF7F', amarillo: '#F2C94C', naranja: '#F2994A', rojo: '#C62828' };

window.__granjasIniciada = true;   // avisa a boot-check.js que la app sí arrancó
hydrateIcons();

const auth = initializeAuth(initializeApp(CONFIG.firebase), { persistence: browserSessionPersistence });

// ==========================================================================
// 2. ESTADO
// ==========================================================================

const emptyDb = () => ({
    granjas: [], perfiles: [], users: [], listas: [], sectores: [], naves: [], lotes: [], loteNaves: [],
    campos: { perfil: [], Reproductora: [], Incubadora: [], Engorde: [], Postura: [] }, tiposLista: [],
    geneticaDe: {}, iecRangos: [], diasRetroactivos: 7, maxMuestra: 1000, estandares: [], tiposEstandar: [],
    enfermedades: [], metodosSerologia: [], pruebasSerologia: {}, deventer: {}, produccion: {},
});
const filtrosIniciales = () => ({ desde: daysAgo(CONFIG.RANGO_INICIAL_DIAS), hasta: today(), granjas: [], agrupar: 'semana', eliminados: false, sector: '', nave: '' });

const state = {
    user: null,
    me: null,
    db: emptyDb(),
    busy: 0,
    granjaId: null,          // representante: granja activa (cada granja se trabaja por separado)
    empresaId: leerPref('granjas.empresa', ''),   // administrador con varias empresas: la elegida ('' = todas)
    linea: 'Postura',
    f: filtrosIniciales(),
    unidad: leerPref('granjas.unidad', 'kg'),   // unidad para ver pesos (g, kg, lb)
    cache: new Map(),        // respuestas de lectura (registros, pesajes, bitácora…) por solicitud
    loading: new Set(),
    sort: {},
    q: {},
    audit: null,
    ddOpen: false,
    registros: new Map(),    // RegistroID → registro visible (para las acciones de la tabla)
    pesajes: new Map(),      // PesajeID → pesaje visible
    lecturas: new Map(),     // LecturaID → lectura visible
    compras: new Map(),      // CompraID → compra de concentrado visible
    serologias: new Map(),   // SerologiaID → muestreo serológico visible
    cFase: '',               // Concentrado: fase de alimento elegida
    srEnf: '',               // Serología: enfermedad y método elegidos en los filtros (srEnf = la única marcada)
    srEnfs: [],              // Serología: enfermedades marcadas en el filtro (varias = se comparan)
    srDd: false,
    srMetodo: '',
    srComp: { prueba: 'HI', enfs: null },   // Serología: comparación entre enfermedades con la misma prueba
    cmp: { sel: [], estado: '', gen: '', q: '' },   // Comparativo de lotes
    estFilas: new Map(),     // EstandarID → fila de la curva visible
    estGen: '',              // Pesos estándar: línea genética y tabla elegidas
    estTipo: 'Comercial',
    bitGranja: '',
};

const tables = {};
const charts = {};
let chartUnmounts = [];
/** Captura de pesos en curso (modal de pesaje). */
let pz = null;
/** Cambia en cada cierre de sesión (ver limpiarSesion). */
let sesionN = 0;

function leerPref(k, def) {
    try { return localStorage.getItem(k) || def; } catch { return def; }
}
function guardarPref(k, v) {
    try { localStorage.setItem(k, v); } catch { /* sin almacenamiento */ }
}

// ==========================================================================
// 3. UTILIDADES
// ==========================================================================

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];

/** Escapa todo texto dinámico antes de insertarlo como HTML. */
function esc(v) {
    return String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
const nf = new Intl.NumberFormat('es-GT', { maximumFractionDigits: 2 });
function toNum(v) {
    if (v === '' || v === null || v === undefined) return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
}
function fmtNum(v) { const n = toNum(v); return n === null ? '—' : nf.format(n); }
const nfQ = new Intl.NumberFormat('es-GT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
/** Quetzales: "Q 6.50". */
function fmtQ(v) { const n = toNum(v); return n === null ? '—' : `Q ${nfQ.format(n)}`; }
function today() { return new Date().toLocaleDateString('en-CA'); }
function daysAgo(n) { const d = new Date(); d.setDate(d.getDate() - n); return d.toLocaleDateString('en-CA'); }
function fmtDate(s) {
    if (!s) return '—';
    const [y, m, d] = String(s).slice(0, 10).split('-');
    return d ? `${d}/${m}/${y}` : String(s);
}
function fmtDateTime(s) {
    if (!s) return '—';
    return String(s).length > 10 ? `${fmtDate(s)} ${String(s).slice(11, 16)}` : fmtDate(s);
}
function norm(s) { return String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim(); }
function slug(s) { return norm(s).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'datos'; }
function initials(name) { return String(name || '?').split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase(); }
/** Plataforma: ve y controla todo. */
const isAdmin = () => state.me?.Rol === ROLES.ADMIN;
/** Administrador de una empresa: todas las granjas de su empresa y sus cuentas. */
const esEmpresa = () => state.me?.Rol === ROLES.ADMIN_EMPRESA;
/** Trabaja con varias granjas a la vez (filtro de granjas en lugar de "Mi granja"). */
const varias = () => isAdmin() || esEmpresa();
const empresaById = id => (state.db.empresas || []).find(e => e.EmpresaID === id);
const empresaName = id => empresaById(id)?.Nombre || '—';
const duenoName = id => (state.db.duenos || []).find(d => d.DuenoID === id)?.Nombre || '—';
/** Plan vigente de una línea de una granja (del snapshot). */
const planDeLinea = (gid, linea) => {
    const lp = (state.db.lineasPlan || []).find(x => x.GranjaID === gid && x.Linea === linea);
    return lp ? (state.db.planes || []).find(p => p.PlanID === lp.PlanID) || null : null;
};
/**
 * ¿El plan incluye la herramienta en el contexto? (alguna de las granjas, en la línea elegida;
 * sin línea, en cualquiera de sus líneas). La plataforma no tiene límites. El servidor vuelve a
 * revisarlo; aquí solo se ocultan botones y pantallas.
 */
function incluye(clave, ids = ctxIds(), linea = state.linea) {
    if (isAdmin()) return true;
    const lineas = linea ? [linea] : LINEAS;
    return ids.some(id => lineas.some(l => (planDeLinea(id, l)?.Herramientas || []).includes(clave)));
}
const nombreHerramienta = clave => (state.db.herramientas || []).find(h => h.clave === clave)?.nombre || clave;
/** Aviso de pantalla bloqueada por el plan. */
function sinPlan(clave) {
    return emptyState(`El plan de ${varias() ? 'estas granjas' : 'su granja'} no incluye ${nombreHerramienta(clave).toLowerCase()}. ${esEmpresa() ? 'Puede cambiarlo en "Mi plan".' : 'Consulte con el administrador de su empresa.'}`, 'lock',
        esEmpresa() ? `<button type="button" class="btn btn-primary" data-action="nav" data-view="miplan">${ico('tag')}Ver mi plan</button>` : '');
}
/** Botón para cargar el formato semanal propio del cliente (registros de postura). */
function botonFormatoSemanal(linea = state.linea) {
    return linea === 'Postura' && incluye('cargarExcel', ctxIds(), linea)
        ? `<button type="button" class="btn" data-action="importar-semanal">${ico('file-spreadsheet')}Mi formato semanal</button>` : '';
}
/** Botón "Cargar desde Excel" (solo si el plan lo incluye). */
function botonImportar(tipo, linea = state.linea) {
    return incluye('cargarExcel', ctxIds(), linea) ? `<button type="button" class="btn" data-action="importar" data-tipo="${tipo}">${ico('upload')}Cargar desde Excel</button>` : '';
}

/** Empresa de quien está en sesión: la elegida arriba si administra varias (null para la plataforma). */
const miEmpresa = () => (isAdmin() ? null : empresaById(state.empresaId) || empresaById(state.me?.EmpresaID));
/** El administrador de empresa con más de una empresa elige con cuál trabaja (o todas). */
const variasEmpresas = () => esEmpresa() && (state.db.empresas || []).length > 1;
/** Granjas del contexto: con una empresa elegida, solo las suyas. */
const granjasVisibles = () => (variasEmpresas() && state.empresaId ? state.db.granjas.filter(g => g.EmpresaID === state.empresaId) : state.db.granjas);
const granjaById = id => state.db.granjas.find(g => g.GranjaID === id);
const granjaName = id => granjaById(id)?.NombreGranja || '—';
const userById = id => state.db.users.find(u => u.ID === id);

function badgeEstado(estado) {
    return estado === 'Activo' || estado === 'Activa'
        ? `<span class="badge badge-ok">${esc(estado)}</span>`
        : `<span class="badge badge-warn">${esc(estado || '—')}</span>`;
}
function emptyState(msg, icon = 'inbox', actionHtml = '') {
    return `<div class="card"><div class="empty"><span class="empty-icon">${ico(icon)}</span><p>${esc(msg)}</p>${actionHtml}</div></div>`;
}
function kpi(label, value, sub, icon) {
    return `<div class="kpi"><span class="kpi-icon">${ico(icon)}</span><span class="kpi-label">${esc(label)}</span><span class="kpi-value">${esc(value)}</span>${sub ? `<span class="kpi-sub">${esc(sub)}</span>` : ''}</div>`;
}
function alertBox(type, icon, html) { return `<div class="alert alert-${type}">${ico(icon)}<span>${html}</span></div>`; }

// ==========================================================================
// 4. API (backend)
// ==========================================================================

class ApiError extends Error {
    constructor(code, message) { super(message); this.code = code; }
}

// Mensajes para el usuario: nunca mencionan tecnologías. El detalle técnico va a la consola.
const MSG = {
    CONFIG: 'La plataforma no está disponible en este momento. Contacta al administrador.',
    NETWORK: 'No hay conexión. Revisa tu internet e intenta de nuevo.',
    BAD_RESPONSE: 'No pudimos confirmar si la operación se completó. Revisa la información actualizada antes de intentarlo de nuevo.',
};
const READ_ACTIONS = new Set(['getSnapshot', 'getRegistros', 'getResumen', 'getPesajes', 'getAmbiente', 'getBitacora', 'getAccesos', 'verifyAudit']);
const ACCESS_CODES = new Set(['NOT_REGISTERED', 'SUSPENDED', 'ACCOUNT_LOCKED']);

async function api(action, payload = {}, { forceRefresh = false, retried = false } = {}) {
    // Los formularios eligen la nave o el lote completo ("L:<LoteID>") en el mismo campo.
    if (typeof payload?.NaveID === 'string' && payload.NaveID.startsWith('L:')) {
        payload = { ...payload, LoteID: payload.NaveID.slice(2), NaveID: '' };
    }
    const idToken = await auth.currentUser?.getIdToken(forceRefresh);
    if (!idToken) throw new ApiError('UNAUTHENTICATED', 'Tu sesión no está iniciada.');

    let res;
    try {
        // text/plain evita el preflight CORS (el servidor no lo admite).
        res = await fetch(CONFIG.API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action, payload, idToken, humano: humano || undefined }),
            credentials: 'omit',
            referrerPolicy: 'no-referrer',
        });
    } catch (e) {
        console.error(`[Granjas] Falló la conexión (${action}):`, e);
        throw new ApiError('NETWORK', MSG.NETWORK);
    }
    const text = await res.text().catch(() => '');
    let json;
    try { json = JSON.parse(text); } catch {
        console.error(`[Granjas] Respuesta no válida (${action}) · HTTP ${res.status}\n`, text.slice(0, 1500));
        if (!retried && READ_ACTIONS.has(action)) return api(action, payload, { forceRefresh, retried: true });
        throw new ApiError('BAD_RESPONSE', MSG.BAD_RESPONSE);
    }
    if (json.ok && !Object.prototype.hasOwnProperty.call(json, 'data')) {
        if (!retried && READ_ACTIONS.has(action)) return api(action, payload, { forceRefresh, retried: true });
        throw new ApiError('BAD_RESPONSE', MSG.BAD_RESPONSE);
    }
    // La verificación de robot sirve una sola vez: el servidor ya la recordó para esta sesión.
    if (json.ok) humano = '';
    if (!json.ok && json.error?.code === 'HUMAN_CHECK' && CONFIG.TURNSTILE_SITE_KEY && !retried) {
        if (await pedirHumano()) return api(action, payload, { forceRefresh, retried: true });
    }
    if (!json.ok) {
        const err = new ApiError(json.error?.code || 'ERROR', json.error?.message || 'Error desconocido.');
        if (state.me) {
            if (err.code === 'PASSWORD_CHANGE_REQUIRED') showPasswordScreen();
            else if (err.code === 'UNAUTHENTICATED') setTimeout(() => signOut(auth), 1500);
            else if (ACCESS_CODES.has(err.code)) showAccessDenied(err);
        }
        throw err;
    }
    return json.data;
}

// ---------- Verificación "No soy un robot" ----------
// Con CONFIG.TURNSTILE_SITE_KEY, el inicio de sesión muestra la verificación; su respuesta viaja en la
// primera petición y el servidor la comprueba (y la recuerda para esa sesión). Si el servidor la vuelve
// a pedir (sesión muy larga), aparece en una ventana.
let humano = '';
let scriptHumano = null;
function cargarHumano() {
    if (!scriptHumano) {
        scriptHumano = new Promise((ok, mal) => {
            const s = document.createElement('script');
            s.src = 'https://challenges.cloudflare.com/turnstile/v0/api.js?render=explicit';
            s.async = true; s.onload = ok; s.onerror = () => { scriptHumano = null; mal(new Error('sin verificación')); };
            document.head.appendChild(s);
        });
    }
    return scriptHumano;
}
async function mostrarHumano(caja) {
    if (!CONFIG.TURNSTILE_SITE_KEY || !caja) return;
    humano = '';
    caja.hidden = false;
    try { await cargarHumano(); } catch { caja.textContent = 'No se pudo cargar la verificación de seguridad. Revise su conexión y recargue la página.'; return; }
    caja.innerHTML = '';
    window.turnstile.render(caja, {
        sitekey: CONFIG.TURNSTILE_SITE_KEY, language: 'es',
        callback: (t) => { humano = t; }, 'expired-callback': () => { humano = ''; }, 'error-callback': () => { humano = ''; },
    });
}
/** Pide la verificación en una ventana; true si se completó. */
async function pedirHumano() {
    let listo = false;
    const promesa = openModal({
        title: 'Verificación de seguridad', submitLabel: 'Continuar',
        body: '<p class="center">Por seguridad, confirme que no es un robot para continuar.</p><div class="humano" id="humano-ventana"></div>',
        onSubmit: async () => { if (!humano) throw new Error('Complete la verificación.'); listo = true; },
    });
    mostrarHumano($('#humano-ventana'));
    await promesa;
    return listo;
}

async function loadSnapshot(opts) {
    const data = await api('getSnapshot', {}, opts);
    state.me = data.me;
    state.mustChangePassword = !!data.mustChangePassword;
    state.db = { ...emptyDb(), ...data };
    if (!varias()) {
        const ids = state.db.granjas.map(g => g.GranjaID);
        if (!ids.includes(state.granjaId)) state.granjaId = ids[0] || null;
    } else {
        if (state.empresaId && !empresaById(state.empresaId)) state.empresaId = '';
        state.f.granjas = state.f.granjas.filter(id => granjaById(id));
    }
}

async function withBusy(fn) {
    state.busy++;
    $('#busy-bar').classList.add('is-active');
    try { return await fn(); } finally {
        state.busy--;
        if (!state.busy) $('#busy-bar').classList.remove('is-active');
    }
}

/** Recarga todo (datos base y consultas) y vuelve a dibujar. */
async function refreshAll() {
    state.cache.clear();
    await withBusy(() => loadSnapshot());
    renderChrome();
    render();
}

/**
 * Lectura con caché: devuelve { ok, data } o { ok:false, error }, o null mientras carga
 * (al terminar vuelve a dibujar la vista).
 */
function need(action, payload) {
    const key = action + JSON.stringify(payload);
    if (state.cache.has(key)) return state.cache.get(key);
    if (!state.loading.has(key)) {
        state.loading.add(key);
        const sesion = sesionN;
        const guardar = (v) => { if (sesion === sesionN) state.cache.set(key, v); };
        withBusy(() => api(action, payload))
            .then(data => guardar({ ok: true, data }))
            .catch(e => guardar({ ok: false, error: e.message, code: e.code }))
            .finally(() => { state.loading.delete(key); if (state.me) render(); });
    }
    return null;
}

function loadingBlock(n = 1) {
    return Array.from({ length: n }, () => '<div class="skeleton" aria-hidden="true"></div>').join('');
}
function errorBlock(r) {
    return `<div class="card"><div class="empty"><span class="empty-icon">${ico('alert-triangle')}</span><p>${esc(r.error)}</p>
        <button type="button" class="btn btn-primary" data-action="retry">${ico('rotate-cw')}Reintentar</button></div></div>`;
}

// ==========================================================================
// 5. DATOS Y AGREGACIONES
// ==========================================================================

function perfilDe(granjaId, linea) {
    return state.db.perfiles.find(p => p.GranjaID === granjaId && p.Linea === linea) || null;
}

/** Granjas del contexto actual. Representante: SOLO la granja activa. Admin: filtro o todas. */
function ctxIds() {
    if (!varias()) return state.granjaId ? [state.granjaId] : [];
    return state.f.granjas.length ? [...state.f.granjas] : granjasVisibles().map(g => g.GranjaID);
}
/** Una línea solo se muestra si alguna granja del contexto tiene su ficha activa. */
function lineasDisponibles(ids) {
    return LINEAS.filter(l => ids.some(id => perfilDe(id, l)?.Estado === 'Activo'));
}
function ensureLinea(ids) {
    const d = lineasDisponibles(ids);
    if (d.length && !d.includes(state.linea)) { state.linea = d[0]; state.f.sector = ''; state.f.nave = ''; }
    return d;
}
/** Granjas que se piden al servidor (el servidor vuelve a validar el acceso). */
function apiGranjas() {
    if (!varias()) return state.granjaId ? [state.granjaId] : [];
    if (state.f.granjas.length) return [...state.f.granjas];
    // Con una de varias empresas elegida, solo sus granjas (sin elegir, el servidor da todas).
    return variasEmpresas() && state.empresaId ? granjasVisibles().map(g => g.GranjaID) : undefined;
}

// ---------- Estructura: sectores → naves → lotes ----------
function byNombre(a, b) { return String(a.Nombre).localeCompare(String(b.Nombre), 'es', { numeric: true }); }
const sectorById = id => state.db.sectores.find(s => s.SectorID === id);
const naveById = id => state.db.naves.find(n => n.NaveID === id);
const loteById = id => state.db.lotes.find(l => l.LoteID === id);
const sectoresDe = (granjaId, linea) => state.db.sectores.filter(s => s.GranjaID === granjaId && s.Linea === linea).sort(byNombre);
const navesDe = sectorId => state.db.naves.filter(n => n.SectorID === sectorId).sort(byNombre);
/** Naves (galeras) de un lote, con sus aves y fechas. Un lote de antes de 4.1.0 sin filas: su nave. */
function navesDelLote(l) {
    const filas = (state.db.loteNaves || []).filter(x => x.LoteID === l.LoteID);
    return filas.length ? filas : [{ NaveID: l.NaveID, AvesIniciales: l.AvesIniciales, Desde: l.FechaInicio, Hasta: '', AvesSalida: '' }];
}
/** Naves del lote que siguen en él (sin fecha de salida). */
const navesVigentes = l => navesDelLote(l).filter(x => !x.Hasta);
const lotesDe = naveId => state.db.lotes.filter(l => navesDelLote(l).some(x => x.NaveID === naveId)).sort((a, b) => (a.FechaInicio < b.FechaInicio ? 1 : -1));
const loteActivo = naveId => state.db.lotes.find(l => l.Estado === 'Activo' && navesVigentes(l).some(x => x.NaveID === naveId)) || null;
/** Lote que ocupaba la nave en una fecha (igual que el servidor). */
function loteEn(naveId, fecha) {
    return lotesDe(naveId).find(l => l.FechaInicio <= fecha && (!l.FechaCierre || l.FechaCierre >= fecha) &&
        navesDelLote(l).some(x => x.NaveID === naveId && x.Desde <= fecha && (!x.Hasta || fecha < x.Hasta))) || null;
}
/** Forma de registro del lote en una fecha: 'Nave' o 'Lote' (todas sus naves juntas). */
function modoEn(l, fecha) {
    const forma = l.Registro || 'Nave';
    if (!l.RegistroDesde || fecha >= l.RegistroDesde) return forma;
    return forma === 'Lote' ? 'Nave' : 'Lote';
}
/** Ubicación elegida en un formulario: una nave, o "L:<LoteID>" (el lote completo). */
function loteDeUbicacion(valor, fecha) {
    return String(valor || '').startsWith('L:') ? loteById(valor.slice(2)) : loteEn(valor, fecha);
}

// ---------- Edad de las aves y peso esperado (mismas fórmulas que el servidor) ----------
/** Edad en días en una fecha: edad al iniciar (semanas) × 7 + días desde el inicio; null sin edad inicial. */
function edadDias(lote, fecha) {
    const semanas = toNum(lote?.EdadInicialSemanas);
    if (semanas === null || !lote.FechaInicio || !fecha) return null;
    const [a, b] = [lote.FechaInicio, String(fecha).slice(0, 10)].map(s => s.split('-').map(Number));
    return Math.round(semanas * 7) + Math.round((Date.UTC(b[0], b[1] - 1, b[2]) - Date.UTC(a[0], a[1] - 1, a[2])) / 86400000);
}
const tipoEstandar = linea => (linea === 'Incubadora' ? null : linea === 'Reproductora' ? 'Reproductora' : 'Comercial');
/** Peso esperado { minimoG, estandarG, maximoG, fuente } por interpolación, o null fuera de la curva. */
/** Valor "tabla" de producción de la guía (igual que produccionEsperada_ del servidor); null si no lo publica. */
function produccionEsperada(genetica, tipo, indicador, dias) {
    const f = state.db.produccion?.[`${genetica}|${tipo}`]?.indicadores?.[indicador];
    if (!f || dias === null || dias === undefined || dias < f[0][0] || dias > f[f.length - 1][0]) return null;
    let i = 0;
    while (i < f.length - 1 && f[i + 1][0] <= dias) i++;
    const a = f[i], b = f[Math.min(i + 1, f.length - 1)];
    if (b[0] === a[0] || (indicador === 'MortalidadIdeal' && b[1] < a[1])) return a[1];
    return a[1] + (b[1] - a[1]) * (dias - a[0]) / (b[0] - a[0]);
}
function esperado(genetica, tipo, dias) {
    if (dias === null || !tipo) return null;
    const filas = state.db.estandares.filter(x => x.LineaGenetica === genetica && x.Tipo === tipo)
        .map(x => ({ d: Number(x.EdadDias), min: toNum(x.MinimoG), std: toNum(x.EstandarG), max: toNum(x.MaximoG), fuente: x.Fuente }))
        .sort((p, q) => p.d - q.d);
    if (!filas.length || dias < filas[0].d || dias > filas[filas.length - 1].d) return null;
    let i = 0;
    while (i < filas.length - 1 && filas[i + 1].d <= dias) i++;
    const a = filas[i], b = filas[Math.min(i + 1, filas.length - 1)];
    const f = b.d === a.d ? 0 : (dias - a.d) / (b.d - a.d);
    const mezcla = k => (a[k] === null || b[k] === null ? null : a[k] + (b[k] - a[k]) * f);
    return { minimoG: mezcla('min'), estandarG: mezcla('std'), maximoG: mezcla('max'), fuente: f < 0.5 ? a.fuente : b.fuente };
}
/** "Semana 23 (día 163)". */
function textoEdad(dias) { return dias === null ? '' : `Semana ${Math.floor(dias / 7)} (día ${dias})`; }
/** Insignia del estado del pesaje contra la tabla (con texto, no solo color). */
function badgeEstadoPeso(estado) {
    const cls = { Bajo: 'badge-danger', Sobre: 'badge-warn', Dentro: 'badge-ok' }[estado] || 'badge-muted';
    return estado ? `<span class="badge ${cls}">${esc(estado)}</span>` : '—';
}
/** Naves que admiten datos nuevos: nave y sector activos. */
function navesActivas(granjaId, linea) {
    return sectoresDe(granjaId, linea).filter(s => s.Estado === 'Activo')
        .flatMap(s => navesDe(s.SectorID).filter(n => n.Estado === 'Activo').map(n => ({ s, n })));
}
/** Filtro de sector/nave: solo cuando el contexto es UNA granja. */
function filtroUbicacion() {
    const ids = ctxIds();
    if (ids.length !== 1) return {};
    const s = state.f.sector && sectorById(state.f.sector);
    const n = state.f.nave && naveById(state.f.nave);
    return {
        SectorID: s && s.GranjaID === ids[0] && s.Linea === state.linea ? s.SectorID : undefined,
        NaveID: n && n.GranjaID === ids[0] && n.Linea === state.linea ? n.NaveID : undefined,
    };
}
function registrosPayload(extra = {}) {
    return {
        Linea: state.linea, GranjaIDs: apiGranjas(), ...filtroUbicacion(),
        desde: state.f.desde || undefined, hasta: state.f.hasta || undefined, ...extra,
    };
}
function fechaMinima() { return isAdmin() ? '' : daysAgo(state.db.diasRetroactivos || 7); }
/** El representante edita o elimina solo datos de hasta N días atrás (el servidor lo vuelve a validar). */
function modificable(fecha) { return isAdmin() || String(fecha || '').slice(0, 10) >= fechaMinima(); }
function sinPermisoFecha() {
    const msg = `Solo la administración modifica datos de más de ${state.db.diasRetroactivos} días atrás`;
    return `<span class="icon-btn icon-btn-sm muted" title="${esc(msg)}" aria-label="${esc(msg)}">${ico('lock')}</span>`;
}

// ---------- Campos ----------
const camposDe = linea => state.db.campos[linea] || [];
const camposNumericos = linea => camposDe(linea).filter(c => c.agg);
const camposEntrada = linea => camposDe(linea).filter(c => !c.calc);
const campoDe = (linea, k) => camposDe(linea).find(c => c.k === k);

// ---------- Pesos (gramos ⇄ unidad elegida) ----------
const factorDe = u => UNIDADES[u] || 1;
function enUnidad(g, u = state.unidad) { const n = toNum(g); return n === null ? null : n / factorDe(u); }
function decimalesPeso(u = state.unidad) { return u === 'g' ? 1 : 3; }
function fmtPeso(g, u = state.unidad) {
    const v = enUnidad(g, u);
    return v === null ? '—' : `${new Intl.NumberFormat('es-GT', { maximumFractionDigits: decimalesPeso(u) }).format(v)} ${u}`;
}
function etiquetaCampo(c) { return c.peso ? `${c.l} (${state.unidad})` : c.l; }
/** Valor de un campo para mostrar o graficar (los pesos, en la unidad elegida). */
function valorCampo(c, v) { return c.peso ? enUnidad(v) : v; }

/** Clave de periodo: día AAAA-MM-DD · semana ISO AAAA-Sss · mes AAAA-MM (igual que el servidor). */
function periodo(fecha, agrupar) {
    const dia = String(fecha).slice(0, 10);
    if (agrupar === 'dia') return dia;
    if (agrupar === 'mes') return dia.slice(0, 7);
    const p = dia.split('-').map(Number);
    const d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
    const wd = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - wd);
    const y = d.getUTCFullYear();
    const w = Math.ceil(((d - Date.UTC(y, 0, 1)) / 86400000 + 1) / 7);
    return `${y}-S${String(w).padStart(2, '0')}`;
}
const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
function periodoLabel(key, agrupar) {
    if (agrupar === 'dia') return fmtDate(key).slice(0, 5) + '/' + key.slice(2, 4);
    if (agrupar === 'mes') return `${MESES[Number(key.slice(5, 7)) - 1]} ${key.slice(0, 4)}`;
    if (key.startsWith('E')) return `Semana ${Number(key.slice(1))}`;   // semana de edad del lote
    return `Sem. ${rangoSemanaIso(key)}`;
}
/** Semana de calendario ISO "AAAA-Sww" → "13–19 oct" (o "29 sep–5 oct"). */
function rangoSemanaIso(key) {
    const [y, w] = [Number(key.slice(0, 4)), Number(key.slice(6))];
    const ene4 = new Date(Date.UTC(y, 0, 4));
    const lunes = new Date(ene4.getTime() + ((w - 1) * 7 - ((ene4.getUTCDay() || 7) - 1)) * 86400000);
    const domingo = new Date(lunes.getTime() + 6 * 86400000);
    const d = x => `${x.getUTCDate()} ${MESES[x.getUTCMonth()]}`;
    return lunes.getUTCMonth() === domingo.getUTCMonth() ? `${lunes.getUTCDate()}–${d(domingo)}` : `${d(lunes)}–${d(domingo)}`;
}
/** Clave de semana de un registro: semana de edad de su lote ("E023"), o de calendario si el lote no tiene edad. */
function claveSemana(r) {
    const dias = edadDias(loteById(r.LoteID), r.Fecha);
    return dias === null ? periodo(r.Fecha, 'semana') : `E${String(Math.floor(dias / 7)).padStart(3, '0')}`;
}

/**
 * Agregación de registros. Los valores "por ave" se ponderan con las aves de cada día (un día de
 * 100 aves no pesa lo mismo que uno de 10,000); las aves y la mortalidad acumulada se toman del
 * último registro de cada lote del grupo y se suman entre lotes; la conversión es alimento total ÷
 * ganancia total.
 */
const POR_AVES = new Set(['MortalidadPorcentaje', 'ConsumoGramosAveDia', 'PosturaPorcentaje', 'PesoCorporalReal']);
const ULTIMO_LOTE = new Set(['NumeroAves', 'MortalidadAcumulada']);
const claveLote = r => r.LoteID || r.NaveID || r.GranjaID;
function aggregate(regs, campos, keyFn) {
    const map = new Map();
    regs.forEach((r) => {
        const k = keyFn(r);
        let g = map.get(k);
        if (!g) { g = { key: k, n: 0, s: {}, c: {}, w: {}, wd: {}, last: new Map(), alim: 0, gan: 0, qqDoc: 0, huevos: 0 }; map.set(k, g); }
        g.n++;
        const aves = toNum(r.NumeroAves);
        campos.forEach((c) => {
            const v = toNum(r[c.k]);
            if (v === null) return;
            g.s[c.k] = (g.s[c.k] || 0) + v;
            g.c[c.k] = (g.c[c.k] || 0) + 1;
            if (POR_AVES.has(c.k) && aves > 0) { g.w[c.k] = (g.w[c.k] || 0) + v * aves; g.wd[c.k] = (g.wd[c.k] || 0) + aves; }
        });
        const alim = toNum(r.AlimentoKg), gan = toNum(r.GananciaKg);
        if (alim !== null && gan > 0) { g.alim += alim; g.gan += gan; }
        // Postura: alimento por docena = Σ alimento ÷ (Σ huevos ÷ 12), de los días con ambos datos.
        const qq = toNum(r.ConsumoQuintalesTotal), hv = toNum(r.PosturaNominal);
        if (qq !== null && hv > 0) { g.qqDoc += qq; g.huevos += hv; }
        const prev = g.last.get(claveLote(r));
        if (!prev || String(r.Fecha) > String(prev.Fecha)) g.last.set(claveLote(r), r);
    });
    return [...map.values()].sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
}
/** Suma, entre los lotes del grupo, un valor del último registro de cada lote. */
function gUltimo(g, k) {
    let s = 0, hay = false;
    g?.last.forEach((r) => { const v = toNum(r[k]); if (v !== null) { s += v; hay = true; } });
    return hay ? s : null;
}
function gAvg(g, k) {
    if (!g) return null;
    if (k === 'ConversionAlimenticia' && g.gan > 0) return g.alim / g.gan;
    if (k === 'AlimentoPorDocena') return g.huevos > 0 ? g.qqDoc * KG_QUINTAL * 1000 / (g.huevos / 12) : null;
    if (k === 'MortalidadAcumuladaPorcentaje') {
        let m = 0, ini = 0;
        g.last.forEach((r) => { const a = toNum(r.MortalidadAcumulada), i = toNum(r.AvesIniciales); if (a !== null && i > 0) { m += a; ini += i; } });
        if (ini) return m / ini * 100;
    }
    if (POR_AVES.has(k) && g.wd[k]) return g.w[k] / g.wd[k];
    return g.c[k] ? g.s[k] / g.c[k] : null;
}
const gSum = (g, k) => (ULTIMO_LOTE.has(k) ? gUltimo(g, k) : g?.c[k] ? g.s[k] : null);
const gVal = (g, c) => (c.agg === 'suma' || ULTIMO_LOTE.has(c.k) ? gSum(g, c.k) : gAvg(g, c.k));
/** Sufijo de columna en tablas agregadas. */
const sufijoAgg = c => (c.agg === 'suma' ? ' (total)' : ULTIMO_LOTE.has(c.k) ? ' (último día)' : POR_AVES.has(c.k) ? ' (ponderado)' : ' (prom.)');
/** Valor agregado listo para mostrar (pesos convertidos, 2–3 decimales). */
function gShow(g, c) {
    const v = gVal(g, c);
    if (v === null) return null;
    return c.peso ? round(enUnidad(v), decimalesPeso()) : round(v);
}
const KG_QUINTAL = 45.359237;
function round(v, d = 2) { return v === null || v === undefined ? null : Math.round(v * 10 ** d) / 10 ** d; }

/** Gráficos por línea: [campo, nombre de la serie, color, punteada]. Los pesos usan la unidad elegida. */
const CHARTS_ZOO = [
    { id: 'peso', title: 'Peso corporal', icon: 'scale', type: 'line', peso: true, banda: true, series: [['PesoCorporalReal', 'Peso real (pesaje)', 'primary'], ['PesoCorporalTabla', 'Estándar de la línea genética', 'tabla', true]] },
    { id: 'conversion', title: 'Conversión alimenticia', icon: 'wheat', type: 'line', series: [['ConversionAlimenticia', 'Conversión (kg alimento / kg ganado)', 'primary']] },
    { id: 'mortalidad-acum', title: 'Mortalidad acumulada del lote (%)', icon: 'activity', unit: '%', type: 'line', series: [['MortalidadAcumuladaPorcentaje', 'Mortalidad acumulada', 'primary'], ['MortalidadIdeal', 'Ideal de la línea genética', 'tabla', true]] },
    { id: 'mortalidad', title: 'Mortalidad del día (%)', icon: 'activity', unit: '%', type: 'line', series: [['MortalidadPorcentaje', 'Mortalidad del día', 'primary']] },
    { id: 'consumo', title: 'Consumo de alimento (g/ave/día)', icon: 'wheat', unit: 'g', type: 'area', series: [['ConsumoGramosAveDia', 'Consumo real', 'primary'], ['ConsumoTabla', 'Tabla', 'tabla', true]] },
    { id: 'iec', title: 'Índice de estrés calórico (máx. del día)', icon: 'thermometer', type: 'line', series: [['IndiceEstresCalorico', 'IEC (°C + % HR)', 'primary']] },
    { id: 'muertes', title: 'Mortalidad (aves)', icon: 'bird', type: 'bar', series: [['MortalidadNominal', 'Mortalidad (aves)', 'primary']] },
];
/** Mínimo y máximo de la tabla de la línea genética (los calcula el servidor en cada registro). */
const CAMPOS_BANDA = [{ k: 'PesoTablaMinimo', agg: 'promedio', peso: true }, { k: 'PesoTablaMaximo', agg: 'promedio', peso: true }];
const CHARTS = {
    Reproductora: CHARTS_ZOO,
    Engorde: CHARTS_ZOO,
    // En postura, la conversión es huevo producido contra alimento (por docena y por kg de huevo).
    Postura: [
        { id: 'postura', title: 'Postura / huevo (%)', icon: 'egg', unit: '%', type: 'area', series: [['PosturaPorcentaje', 'Postura real', 'primary'], ['PosturaTabla', 'Tabla', 'tabla', true]] },
        { id: 'huevo', title: 'Clasificación de huevo', icon: 'chart-column', type: 'bar', stack: true, series: [['HuevoPiwi', 'Piwi', 'o1'], ['HuevoMediano', 'Mediano', 'o2'], ['HuevoGrande', 'Grande', 'o3'], ['HuevoExtragrande', 'Extragrande', 'o4']] },
        { id: 'docena', title: 'Alimento por docena de huevos', icon: 'wheat', type: 'line', peso: true, series: [['AlimentoPorDocena', 'Alimento por docena', 'primary']] },
        { id: 'conversion', title: 'Conversión (kg alimento / kg huevo)', icon: 'wheat', type: 'line', series: [['ConversionAlimenticia', 'kg de alimento por kg de huevo', 'primary']] },
    ].concat(CHARTS_ZOO.filter(c => c.id !== 'conversion')),
    Incubadora: [
        { id: 'nacimiento', title: '% de nacimiento', icon: 'egg', unit: '%', type: 'area', series: [['PorcentajeNacimiento', 'Nacimiento ÷ huevo incubable', 'primary']] },
        { id: 'huevos', title: 'Huevos y nacimientos', icon: 'chart-column', type: 'bar', series: [['HuevoIncubable', 'Huevo incubable', 'o1'], ['HuevoTransferencia', 'Huevo de transferencia', 'o3'], ['Nacimiento', 'Nacimiento', 'o4']] },
        { id: 'pesohuevo', title: 'Peso de huevo', icon: 'scale', type: 'line', peso: true, series: [['PesoHuevo', 'Peso promedio (pesaje)', 'primary']] },
        { id: 'mortalidad', title: 'Mortalidad (%)', icon: 'activity', unit: '%', type: 'line', series: [['MortalidadPorcentaje', 'Mortalidad', 'primary']] },
    ],
};
/** Indicador principal de cada línea para el comparativo (real y tabla). */
const COMPARATIVO = {
    Reproductora: ['PesoCorporalReal', 'PesoCorporalTabla', 'Peso corporal'],
    Engorde: ['PesoCorporalReal', 'PesoCorporalTabla', 'Peso corporal'],
    Postura: ['PosturaPorcentaje', 'PosturaTabla', 'Postura (%)'],
    Incubadora: ['PorcentajeNacimiento', null, '% de nacimiento'],
};

/** Nivel del índice de estrés calórico (rangos del servidor). */
function nivelIec(v) {
    const n = toNum(v);
    if (n === null) return null;
    return state.db.iecRangos.find(r => r.hasta === null || n < r.hasta) || state.db.iecRangos[state.db.iecRangos.length - 1];
}
function badgeIec(v, nivel) {
    const r = nivel ? state.db.iecRangos.find(x => x.nivel === nivel) : nivelIec(v);
    if (!r) return '—';
    return `<span class="badge badge-iec iec-${esc(r.color)}">${esc(r.nivel)}</span>`;
}

// ==========================================================================
// 6. TABLAS (ordenar A-Z / Z-A, buscar, Excel) Y GRÁFICOS (PNG)
// ==========================================================================

/**
 * columns: [{ k, l, t: 'text'|'num'|'date'|'datetime', get?(r), html?(r), csv?(r), sticky?, wrap?, nocsv? }]
 */
function table(id, { title, icon = 'table', columns, rows, actions, rowClass, empty = 'No hay datos para mostrar.', csv }) {
    tables[id] = { title, columns, rows, actions, rowClass, empty, csv };
    const n = visibleRows(id).length;
    return `<div class="card table-card">
        <div class="table-tools">
            <h3>${ico(icon)}${esc(title)} <span class="badge badge-muted" id="tcount-${id}">${nf.format(n)}</span></h3>
            <div class="tools">
                <div class="search">${ico('search', 'icon-sm')}<input class="input" type="search" placeholder="Buscar en la tabla" aria-label="Buscar en ${esc(title)}" data-input="table-q" data-table="${id}" value="${esc(state.q[id] || '')}"></div>
                ${incluye('descargarExcel', ctxIds(), null) ? `<button type="button" class="btn btn-sm" data-action="xlsx" data-table="${id}" title="Descargar en Excel">${ico('file-spreadsheet')}Excel</button>` : ''}
            </div>
        </div>
        <div class="table-wrap" id="tw-${id}">${tableInner(id)}</div>
    </div>`;
}

function cellValue(c, r) { return c.get ? c.get(r) : r[c.k]; }
function cellText(c, r) {
    const v = cellValue(c, r);
    if (c.t === 'datetime') return fmtDateTime(v);
    if (c.t === 'date') return fmtDate(v);
    return v ?? '';
}
function compare(c, a, b) {
    const va = cellValue(c, a), vb = cellValue(c, b);
    if (c.t === 'num') return toNum(va) - toNum(vb);
    if (c.t === 'date' || c.t === 'datetime') return String(va).localeCompare(String(vb));
    return String(va).localeCompare(String(vb), 'es', { sensitivity: 'base', numeric: true });
}
function isBlank(c, r) { const v = cellValue(c, r); return v === '' || v === null || v === undefined || (c.t === 'num' && toNum(v) === null); }

function visibleRows(id) {
    const t = tables[id];
    let rows = t.rows;
    const q = norm(state.q[id]);
    if (q) rows = rows.filter(r => t.columns.some(c => norm(cellText(c, r)).includes(q)));
    const s = state.sort[id];
    const c = s && t.columns.find(x => x.k === s.k);
    if (c) {
        const dir = s.dir === 'asc' ? 1 : -1;
        rows = [...rows].sort((a, b) => {
            const ba = isBlank(c, a), bb = isBlank(c, b);
            if (ba || bb) return ba === bb ? 0 : ba ? 1 : -1;   // vacíos siempre al final
            return compare(c, a, b) * dir;
        });
    }
    return rows;
}

function tableInner(id) {
    const t = tables[id];
    const rows = visibleRows(id);
    if (!rows.length) {
        return `<div class="empty"><span class="empty-icon">${ico('inbox')}</span><p>${esc(state.q[id] ? 'Ninguna fila coincide con la búsqueda.' : t.empty)}</p></div>`;
    }
    const s = state.sort[id];
    const head = t.columns.map((c) => {
        const sorted = s?.k === c.k;
        const iconName = !sorted ? 'arrow-up-down' : s.dir === 'asc' ? 'arrow-down-a-z' : 'arrow-down-z-a';
        const aria = sorted ? (s.dir === 'asc' ? 'ascending' : 'descending') : 'none';
        const next = sorted && s.dir === 'asc' ? 'Z-A' : 'A-Z';
        return `<th scope="col" class="${c.sticky ? 'sticky-col' : ''}" aria-sort="${aria}"><button type="button" class="th-sort${sorted ? ' is-sorted' : ''}" data-action="sort" data-table="${id}" data-k="${esc(c.k)}" title="Ordenar ${next}">${esc(c.l)}${ico(iconName)}</button></th>`;
    }).join('') + (t.actions ? '<th scope="col" class="actions">Acciones</th>' : '');

    const shown = rows.slice(0, CONFIG.MAX_FILAS_TABLA);
    const body = shown.map((r) => {
        const cells = t.columns.map((c) => {
            const cls = [c.sticky && 'sticky-col', c.t === 'num' && 'num', c.wrap && 'wrap'].filter(Boolean).join(' ');
            const html = c.html ? c.html(r)
                : c.t === 'num' ? fmtNum(cellValue(c, r))
                    : esc(cellText(c, r) === '' ? '—' : cellText(c, r));
            return `<td${cls ? ` class="${cls}"` : ''}>${html}</td>`;
        }).join('');
        const act = t.actions ? `<td class="actions"><div class="row">${t.actions(r)}</div></td>` : '';
        const rc = t.rowClass?.(r);
        return `<tr${rc ? ` class="${rc}"` : ''}>${cells}${act}</tr>`;
    }).join('');
    const foot = rows.length > shown.length
        ? `<div class="table-foot">Se muestran ${nf.format(shown.length)} de ${nf.format(rows.length)} filas. El Excel incluye todas.</div>` : '';
    return `<table class="table"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>${foot}`;
}

function refreshTable(id) {
    const wrap = $(`#tw-${id}`);
    if (wrap) wrap.innerHTML = tableInner(id);
    const count = $(`#tcount-${id}`);
    if (count) count.textContent = nf.format(visibleRows(id).length);
}

/** Qué datos se están viendo (para el encabezado de las descargas). */
function contextoFiltros(id) {
    const ids = ctxIds();
    const ub = filtroUbicacion();
    const partes = [
        varias() ? (state.f.granjas.length ? state.f.granjas.map(granjaName).join(', ') : 'Todas las granjas') : granjaName(state.granjaId),
        ub.SectorID && `Sector ${sectorById(ub.SectorID)?.Nombre}`, ub.NaveID && `Nave ${naveById(ub.NaveID)?.Nombre}`,
        state.f.desde || state.f.hasta ? `Del ${state.f.desde ? fmtDate(state.f.desde) : 'inicio'} al ${state.f.hasta ? fmtDate(state.f.hasta) : 'hoy'}` : 'Todo el historial',
        state.q[id] ? `Búsqueda: "${state.q[id]}"` : '',
    ];
    return ids.length ? partes.filter(Boolean).join(' · ') : '';
}

/**
 * Excel (.xlsx) legible para cualquier persona: título, filtros aplicados y fecha de descarga;
 * encabezados resaltados y fijos, con filtros; números como números, fechas como fechas.
 */
function downloadXlsx(id) {
    const t = tables[id];
    if (!t) return;
    const cols = t.columns.filter(c => !c.nocsv);
    const rows = visibleRows(id);
    const tipo = c => (c.t === 'num' ? 'num' : c.t === 'date' ? 'date' : c.t === 'datetime' ? 'datetime' : c.wrap ? 'wrap' : 'text');
    const valor = (c, r) => {
        const v = c.csv ? c.csv(r) : cellValue(c, r);
        return c.t === 'num' ? toNum(v) : v;
    };
    const blob = crearXlsx([{
        nombre: t.title, activa: true,
        titulo: [t.title, contextoFiltros(id), `Descargado el ${fmtDateTime(`${today()} ${new Date().toTimeString().slice(0, 5)}`)} por ${state.me?.Nombre || ''} · CHIIK`].filter(Boolean),
        columnas: cols.map(c => ({ l: c.l, t: tipo(c) })),
        filas: rows.map(r => cols.map(c => valor(c, r))),
    }]);
    downloadBlob(blob, `${slug(t.csv || t.title)}_${today()}.xlsx`);
    toast(`Excel descargado (${nf.format(rows.length)} filas).`, 'success');
}

function chartCard(id, spec) {
    charts[id] = spec;
    return `<div class="card chart-card">
        <div class="card-head"><h3>${ico(spec.icon || 'chart-line')}${esc(spec.title)}</h3>
            ${incluye('exportarImagen', ctxIds(), null) ? `<button type="button" class="btn btn-sm" data-action="png" data-chart="${id}" title="Descargar PNG">${ico('image-down')}PNG</button>` : ''}</div>
        <div class="chart-box"><canvas data-chart="${id}"></canvas></div>
    </div>`;
}

function mountCharts() {
    chartUnmounts.forEach(f => f());
    chartUnmounts = $$('canvas[data-chart]').map((cv) => {
        const spec = charts[cv.dataset.chart];
        return spec.type === 'heatmap' ? mountHeatmap(cv, spec) : mountChart(cv, spec);
    });
}

// ==========================================================================
// 7. NAVEGACIÓN Y ESTRUCTURA
// ==========================================================================

const NAV = {
    ADMIN: [['resumen', 'Resumen global', 'chart-line'], ['dashboard', 'Dashboard', 'layout-dashboard'], ['comparativo', 'Comparar lotes', 'chart-column'], ['registros', 'Registros', 'clipboard-list'],
        ['pesajes', 'Pesajes', 'scale'], ['serologia', 'Serología', 'syringe'], ['ambiente', 'Estrés calórico', 'thermometer'], ['concentrado', 'Concentrado', 'wheat'],
        ['empresas', 'Empresas', 'building'], ['planes', 'Planes y precios', 'tag'], ['cobros', 'Cobros y pagos', 'receipt'], ['granjas', 'Granjas', 'warehouse'], ['usuarios', 'Usuarios', 'users'], ['listas', 'Listas', 'list'], ['estandares', 'Pesos estándar', 'chart-line'],
        ['bitacora', 'Bitácora', 'history'], ['seguridad', 'Seguridad', 'shield-check'], ['uso', 'Uso', 'activity'], ['configuracion', 'Configuración', 'list-checks']],
    ADMIN_EMPRESA: [['resumen', 'Resumen de granjas', 'chart-line'], ['dashboard', 'Dashboard', 'layout-dashboard'], ['comparativo', 'Comparar lotes', 'chart-column'], ['registros', 'Registros', 'clipboard-list'],
        ['pesajes', 'Pesajes', 'scale'], ['serologia', 'Serología', 'syringe'], ['ambiente', 'Estrés calórico', 'thermometer'], ['concentrado', 'Concentrado', 'wheat'],
        ['granjas', 'Mis granjas', 'warehouse'], ['miplan', 'Mi plan y pagos', 'tag'], ['usuarios', 'Usuarios', 'users'], ['bitacora', 'Bitácora', 'history']],
    REPRESENTANTE: [['granja', 'Mi granja', 'warehouse'], ['dashboard', 'Dashboard', 'layout-dashboard'], ['comparativo', 'Comparar lotes', 'chart-column'],
        ['registros', 'Registros', 'clipboard-list'], ['pesajes', 'Pesajes', 'scale'], ['serologia', 'Serología', 'syringe'], ['ambiente', 'Estrés calórico', 'thermometer'],
        ['concentrado', 'Concentrado', 'wheat'], ['bitacora', 'Bitácora', 'history']],
};
const HOME = { ADMIN: 'dashboard', ADMIN_EMPRESA: 'dashboard', REPRESENTANTE: 'granja' };
const A = [ROLES.ADMIN], AE = [ROLES.ADMIN, ROLES.ADMIN_EMPRESA], ALL = [ROLES.ADMIN, ROLES.ADMIN_EMPRESA, ROLES.REPRESENTANTE];
const VIEWS = {
    dashboard: { roles: ALL, title: 'Dashboard', render: viewDashboard },
    comparativo: { roles: ALL, title: 'Comparativo de lotes', render: viewComparativo },
    registros: { roles: ALL, title: 'Registros diarios', render: viewRegistros },
    pesajes: { roles: ALL, title: 'Pesajes por muestra', render: viewPesajes },
    serologia: { roles: ALL, title: 'Serología por lote', render: viewSerologia },
    ambiente: { roles: ALL, title: 'Estrés calórico', render: viewAmbiente },
    concentrado: { roles: ALL, title: 'Compras de concentrado', render: viewConcentrado },
    estandares: { roles: A, title: 'Pesos estándar por línea genética', render: viewEstandares },
    granja: { roles: ALL, title: 'Granja', render: viewGranja },
    empresas: { roles: A, title: 'Empresas', render: viewEmpresas },
    planes: { roles: A, title: 'Planes y precios', render: viewPlanes },
    miplan: { roles: [ROLES.ADMIN_EMPRESA], title: 'Mi plan y pagos', render: viewPlanes },
    cobros: { roles: A, title: 'Cobros y pagos', render: viewCobros },
    resumen: { roles: AE, title: 'Resumen de granjas', render: viewResumen },
    uso: { roles: A, title: 'Uso de la plataforma', render: viewUso },
    configuracion: { roles: A, title: 'Configuración', render: viewConfiguracion },
    granjas: { roles: AE, title: 'Granjas', render: viewGranjas },
    usuarios: { roles: AE, title: 'Usuarios', render: viewUsuarios },
    listas: { roles: A, title: 'Listas desplegables', render: viewListas },
    bitacora: { roles: ALL, title: 'Bitácora de cambios', render: viewBitacora },
    seguridad: { roles: A, title: 'Seguridad', render: viewSeguridad },
};

function parseHash() {
    const [view, param] = location.hash.replace(/^#\/?/, '').split('/');
    return { view: view || '', param: param ? decodeURIComponent(param) : null };
}
function go(view, param = null) {
    const h = `#/${view}${param ? '/' + encodeURIComponent(param) : ''}`;
    if (location.hash === h) render(); else location.hash = h;
}

function render() {
    if (!state.me || $('#screen-app').hidden) return;
    const { view, param } = parseHash();
    const v = VIEWS[view];
    if (!v || !v.roles.includes(state.me.Rol)) { go(HOME[state.me.Rol]); return; }
    Object.keys(tables).forEach(k => delete tables[k]);
    Object.keys(charts).forEach(k => delete charts[k]);
    $('#page-title').textContent = view === 'granja' && !varias() ? 'Mi granja' : v.title;
    document.title = `${$('#page-title').textContent} · CHIIK`;
    $$('#nav a').forEach(a => a.classList.toggle('is-active', a.dataset.view === view || (view === 'granja' && a.dataset.view === 'granjas')));
    $('#view').innerHTML = avisoEmpresa() + v.render(param);
    mountCharts();
}

/** Empresa suspendida: su gente ve y consulta, pero no registra ni modifica (lo aplica el servidor). */
const AVISO_ICONO = { info: 'calendar-range', warn: 'alert-triangle', danger: 'lock' };
function avisoEmpresa() {
    const e = miEmpresa();
    if (!e) return '';
    const avisos = (state.db.cuenta?.avisos || []).map(a => alertBox(a.tipo, AVISO_ICONO[a.tipo] || 'info',
        `<strong>${esc(a.titulo)}.</strong> ${esc(a.mensaje)}${esEmpresa() ? ' <a href="#/miplan">Ver mi plan y pagos</a>' : ''}`)).join('');
    if (e.Estado === 'Activa') return avisos;
    return alertBox('warn', 'lock', '<strong>Ingreso de datos suspendido.</strong> Puede ver y consultar la información de su empresa, pero no registrar ni modificar datos. Contacte a la administración de CHIIK.') + avisos;
}

/** Al iniciar sesión, los avisos de cobro se muestran una vez en una ventana (luego quedan en la franja). */
function avisosAlEntrar() {
    const avisos = state.db.cuenta?.avisos || [];
    if (!avisos.length) return;
    setTimeout(() => openModal({
        title: avisos[0].titulo, cancelLabel: 'Entendido',
        body: avisos.map(a => alertBox(a.tipo, AVISO_ICONO[a.tipo] || 'info', esc(a.mensaje))).join('') +
            (esEmpresa() ? `<div class="row row-center"><button type="button" class="btn btn-primary" data-action="ir-miplan">${ico('tag')}Ver mi plan y pagos</button></div>`
                : '<p class="muted center">El administrador de su empresa puede ver el detalle.</p>'),
    }), 300);
}

function renderChrome() {
    const porPlan = { serologia: 'serologia', comparativo: 'comparativo', ambiente: 'estres' };
    $('#nav').innerHTML = NAV[state.me.Rol].filter(([view]) => !porPlan[view] || incluye(porPlan[view], state.db.granjas.map(g => g.GranjaID), null)).map(([view, label, icon]) =>
        `<a href="#/${view}" data-view="${view}">${ico(icon)}<span>${esc(label)}</span></a>`).join('');
    $('#user-chip').innerHTML = `<span class="avatar" aria-hidden="true">${esc(initials(state.me.Nombre))}</span>
        <span class="user-chip-text"><strong>${esc(state.me.Nombre)}</strong><span>${esc(ROLE_LABEL[state.me.Rol] || state.me.Rol)}</span></span>`;
}

function showScreen(name) {
    ['loading', 'login', 'denied', 'password', 'app'].forEach(s => { $(`#screen-${s}`).hidden = s !== name; });
    if (name !== 'app') closeModal();
}

// Menú: en computadora la hamburguesa oculta/muestra la barra lateral; en celular arranca oculta.
const mqMobile = matchMedia('(max-width: 959px)');
function sidebarVisible() {
    const shell = $('#screen-app');
    return mqMobile.matches ? shell.classList.contains('sidebar-open') : !shell.classList.contains('sidebar-collapsed');
}
function syncMenuButton() { $('#menu-btn').setAttribute('aria-expanded', String(sidebarVisible())); }
function toggleSidebar(force) {
    const shell = $('#screen-app');
    if (mqMobile.matches) {
        shell.classList.toggle('sidebar-open', force ?? !shell.classList.contains('sidebar-open'));
    } else {
        const collapsed = force === undefined ? !shell.classList.contains('sidebar-collapsed') : !force;
        shell.classList.toggle('sidebar-collapsed', collapsed);
        try { localStorage.setItem('granjas.menu', collapsed ? 'oculto' : 'visible'); } catch { /* sin almacenamiento */ }
    }
    syncMenuButton();
}
(function initSidebar() {
    let pref = 'visible';
    try { pref = localStorage.getItem('granjas.menu') || 'visible'; } catch { /* sin almacenamiento */ }
    $('#screen-app').classList.toggle('sidebar-collapsed', pref === 'oculto');
    mqMobile.addEventListener('change', () => { $('#screen-app').classList.remove('sidebar-open'); syncMenuButton(); });
    syncMenuButton();
})();

// ----------------------------------------------------------------------------
// Componentes de filtros
// ----------------------------------------------------------------------------

/** Representante: SOLO sus granjas, una a la vez (cada granja se trabaja por separado). */
function granjaSwitch() {
    const gs = state.db.granjas;
    if (gs.length <= 1) {
        return `<div class="filter-group"><span class="field-label">Granja</span><span class="badge">${ico('warehouse', 'icon-sm')}${esc(gs[0]?.NombreGranja || '—')}</span></div>`;
    }
    return `<label class="field"><span>Granja</span><select data-change="granja-activa">${gs.map(g =>
        `<option value="${esc(g.GranjaID)}"${g.GranjaID === state.granjaId ? ' selected' : ''}>${esc(g.NombreGranja)}</option>`).join('')}</select></label>`;
}

/** Administrador de varias empresas: todas o una (las granjas se filtran por la elegida). */
function empresaPicker() {
    if (!variasEmpresas()) return '';
    return `<label class="field"><span>Empresa</span><select data-change="empresa-activa"><option value="">Todas mis empresas</option>${state.db.empresas.map(e =>
        `<option value="${esc(e.EmpresaID)}"${e.EmpresaID === state.empresaId ? ' selected' : ''}>${esc(e.Nombre)}</option>`).join('')}</select></label>`;
}

/** Administrador: todas las granjas o las que elija. */
function granjaPicker() {
    const sel = state.f.granjas;
    const label = !sel.length ? 'Todas las granjas' : sel.length === 1 ? granjaName(sel[0]) : `${sel.length} granjas`;
    return `<div class="filter-group"><span class="field-label">Granjas</span>
        <details class="dropdown"${state.ddOpen ? ' open' : ''}>
            <summary>${ico('filter', 'icon-sm')}<span>${esc(label)}</span>${ico('chevron-down', 'icon-sm')}</summary>
            <div class="dropdown-panel">
                <label class="check"><input type="checkbox" data-change="f-granja-all"${!sel.length ? ' checked' : ''}> Todas las granjas</label>
                ${granjasVisibles().map(g => `<label class="check"><input type="checkbox" data-change="f-granja" value="${esc(g.GranjaID)}"${sel.includes(g.GranjaID) ? ' checked' : ''}> ${esc(g.NombreGranja)}${g.Estado !== 'Activo' ? ' <span class="badge badge-warn">Suspendida</span>' : ''}</label>`).join('')}
            </div>
        </details></div>`;
}

/** Solo aparecen las líneas que la granja tiene (las demás no se muestran). */
function lineaTabs(disp) {
    return `<div class="filter-group"><span class="field-label">Línea</span><div class="segmented${disp.length > 2 ? ' seg-lineas' : ''}" role="group" aria-label="Línea de producción">${disp.map(l =>
        `<button type="button" class="seg${l === state.linea ? ' is-active' : ''}" data-action="set-linea" data-linea="${l}" aria-pressed="${l === state.linea}">${ico(LINEA_ICON[l])}${l}</button>`).join('')}</div></div>`;
}

/** Sector y nave: solo con UNA granja en contexto. Incluye los inactivos (para ver su historial). */
function ubicacionFilters() {
    const ids = ctxIds();
    if (ids.length !== 1) return '';
    const secs = sectoresDe(ids[0], state.linea);
    if (!secs.length) return '';
    if (state.f.sector && !secs.some(s => s.SectorID === state.f.sector)) state.f.sector = '';
    const naves = (state.f.sector ? [sectorById(state.f.sector)] : secs).flatMap(s => navesDe(s.SectorID).map(n => ({ s, n })));
    if (state.f.nave && !naves.some(x => x.n.NaveID === state.f.nave)) state.f.nave = '';
    const inactivo = x => (x.Estado !== 'Activo' ? ' (inactiva)' : '');
    return `<label class="field"><span>Sector</span><select data-change="f-sector"><option value="">Todos</option>${secs.map(s =>
        `<option value="${esc(s.SectorID)}"${s.SectorID === state.f.sector ? ' selected' : ''}>${esc(s.Nombre)}${s.Estado !== 'Activo' ? ' (inactivo)' : ''}</option>`).join('')}</select></label>
        <label class="field"><span>Nave</span><select data-change="f-nave"><option value="">Todas</option>${naves.map(({ s, n }) =>
        `<option value="${esc(n.NaveID)}"${n.NaveID === state.f.nave ? ' selected' : ''}>${state.f.sector ? '' : esc(s.Nombre) + ' / '}${esc(n.Nombre)}${inactivo(n)}</option>`).join('')}</select></label>`;
}

function unidadToggle() {
    return `<div class="filter-group"><span class="field-label">Ver pesos en</span><div class="segmented" role="group" aria-label="Unidad de peso">${Object.keys(UNIDADES).map(u =>
        `<button type="button" class="seg${u === state.unidad ? ' is-active' : ''}" data-action="set-unidad" data-unidad="${u}" aria-pressed="${u === state.unidad}" title="${UNIDAD_NOMBRE[u]}">${u}</button>`).join('')}</div></div>`;
}

function dateFilters() {
    return `<div class="date-range"><label class="field"><span>Desde</span><input type="date" data-change="f-desde" value="${esc(state.f.desde)}"${state.f.hasta ? ` max="${esc(state.f.hasta)}"` : ''}></label>
        <label class="field"><span>Hasta</span><input type="date" data-change="f-hasta" value="${esc(state.f.hasta)}"${state.f.desde ? ` min="${esc(state.f.desde)}"` : ''}></label></div>
        <div class="filter-group"><span class="field-label">Rango rápido</span><div class="presets">${PRESETS.map(([d, l]) =>
        `<button type="button" class="btn btn-sm" data-action="preset" data-days="${d}">${l}</button>`).join('')}</div></div>`;
}

function filtersCard({ lineas = true, fechas = true, agrupar = false, eliminados = false, ubicacion = true, unidad = false, extra = '' } = {}) {
    const disp = lineas ? ensureLinea(ctxIds()) : [];
    return `<div class="card filters-card"><div class="filters">
        ${empresaPicker()}${varias() ? granjaPicker() : granjaSwitch()}
        ${lineas ? lineaTabs(disp) : ''}
        ${lineas && ubicacion ? ubicacionFilters() : ''}
        ${extra}
        ${fechas ? dateFilters() : ''}
        ${agrupar ? `<div class="filter-group"><span class="field-label">Agrupar por</span><div class="segmented" role="group" aria-label="Agrupar por">${AGRUPAR.map(([k, l]) =>
        `<button type="button" class="seg${state.f.agrupar === k ? ' is-active' : ''}" data-action="set-agrupar" data-agrupar="${k}" aria-pressed="${state.f.agrupar === k}">${l}</button>`).join('')}</div></div>` : ''}
        ${unidad ? unidadToggle() : ''}
        ${eliminados && isAdmin() ? `<label class="check"><input type="checkbox" data-change="f-eliminados"${state.f.eliminados ? ' checked' : ''}> Incluir eliminados</label>` : ''}
    </div></div>`;
}

function noGranjas() {
    if (esEmpresa()) return emptyState('Su empresa aún no tiene granjas. Cree la primera y agregue sus líneas de producción.', 'warehouse', `<button type="button" class="btn btn-primary" data-action="new-granja">${ico('plus')}Nueva granja</button>`);
    return varias()
        ? emptyState('Aún no hay granjas registradas.', 'warehouse', `<button type="button" class="btn btn-primary" data-action="new-granja">${ico('plus')}Nueva granja</button>`)
        : emptyState('No tienes granjas asignadas. Contacta a la administración.', 'warehouse');
}

function sinLineas() {
    return emptyState(varias()
        ? 'Las granjas seleccionadas aún no tienen ninguna línea de producción (ficha) activa.'
        : 'Tu granja aún no tiene líneas de producción activas. Contacta a la administración.', 'clipboard-list');
}

/** Aviso cuando no hay dónde ingresar datos (sin sectores, naves activas o lotes). */
function sinEstructura(linea) {
    const ids = ctxIds();
    const hay = ids.some(id => navesActivas(id, linea).some(({ n }) => loteActivo(n.NaveID)));
    if (hay) return '';
    const destino = ids.length === 1 ? `<button type="button" class="btn btn-sm" data-action="open-granja" data-id="${esc(ids[0])}">${ico('layout-grid')}Ir a sectores y naves</button>` : '';
    return `<div class="alert alert-info">${ico('info')}<span>Para ingresar datos de ${esc(linea.toLowerCase())} se necesita al menos un sector, una nave activa y un lote iniciado.</span>${destino}</div>`;
}

// ==========================================================================
// 8. VISTAS: DASHBOARD, REGISTROS, PESAJES Y AMBIENTE
// ==========================================================================

/** Columna de tabla para un campo de registro (los pesos, en la unidad elegida). */
function colCampo(c) {
    return {
        k: c.k, l: etiquetaCampo(c), t: c.agg ? 'num' : 'text', wrap: c.t === 'textoLargo',
        get: c.peso ? (r => enUnidad(r[c.k])) : undefined,
    };
}
const COLS_UBICACION = [{ k: 'Sector', l: 'Sector' }, { k: 'Nave', l: 'Nave' }, { k: 'NumeroLote', l: 'Lote' }];
function colsGranja() { return varias() ? [{ k: 'GranjaID', l: 'Granja', get: x => granjaName(x.GranjaID) }] : []; }

/**
 * Indicadores de las tarjetas que dependen de los lotes:
 *  aves vivas = aves al final del último día registrado de cada lote abierto;
 *  mortalidad del rango = muertas del rango ÷ aves al inicio del rango (sumando lotes);
 *  mortalidad acumulada = Σ muertas de los lotes abiertos ÷ Σ aves iniciales;
 *  peso = último pesaje de cada lote abierto, ponderado por sus aves.
 */
function indicadoresLotes(regs, linea) {
    const porLote = new Map();
    regs.forEach((r) => { const k = claveLote(r); if (!porLote.has(k)) porLote.set(k, []); porLote.get(k).push(r); });
    const abierto = (r) => {
        const l = r.LoteID && loteById(r.LoteID);
        return l ? l.Estado === 'Activo' : naveById(r.NaveID)?.Estado === 'Activo';
    };
    const out = { avesVivas: 0, lotes: 0, mortRango: null, mortAcum: null, muertasAcum: 0, iniciales: 0, peso: null, pesoTabla: null };
    let muertasRango = 0, avesInicioRango = 0, pp = 0, pt = 0, pa = 0;
    porLote.forEach((lista) => {
        lista.sort((a, b) => (String(a.Fecha) < String(b.Fecha) ? -1 : 1));
        const primero = lista[0], ultimo = lista[lista.length - 1];
        const avesIni = toNum(primero.NumeroAves);
        if (avesIni > 0) {
            avesInicioRango += avesIni;
            muertasRango += lista.reduce((a, r) => a + (toNum(r.MortalidadNominal) || 0), 0);
        }
        if (!abierto(ultimo)) return;
        out.lotes++;
        const aves = toNum(ultimo.NumeroAves) || 0;
        out.avesVivas += linea === 'Incubadora' ? aves : Math.max(0, aves - (toNum(ultimo.MortalidadNominal) || 0) - (toNum(ultimo.OtrasSalidas) || 0));
        const ini = toNum(ultimo.AvesIniciales), acum = toNum(ultimo.MortalidadAcumulada);
        if (ini > 0 && acum !== null) { out.iniciales += ini; out.muertasAcum += acum; }
        const pesado = [...lista].reverse().find(r => toNum(r.PesoCorporalReal) !== null);
        if (pesado) {
            const w = toNum(pesado.NumeroAves) || 1;
            pp += toNum(pesado.PesoCorporalReal) * w; pa += w;
            if (toNum(pesado.PesoCorporalTabla) !== null) pt += toNum(pesado.PesoCorporalTabla) * w;
        }
    });
    if (avesInicioRango) out.mortRango = muertasRango / avesInicioRango * 100;
    if (out.iniciales) out.mortAcum = out.muertasAcum / out.iniciales * 100;
    if (pa) { out.peso = pp / pa; out.pesoTabla = pt ? pt / pa : null; }
    return out;
}

function viewDashboard() {
    const ids = ctxIds();
    if (!ids.length) return noGranjas();
    const head = filtersCard({ agrupar: true, unidad: true });
    if (!lineasDisponibles(ids).length) return head + sinLineas();

    const r = need('getRegistros', registrosPayload());
    if (!r) return head + `<div class="kpis">${loadingBlock(1)}</div><div class="grid-2">${loadingBlock(2)}</div>`;
    if (!r.ok) return head + errorBlock(r);

    const linea = state.linea;
    const regs = r.data.registros;
    const campos = camposNumericos(linea);
    const agr = state.f.agrupar;
    const rango = state.f.desde || state.f.hasta ? `${fmtDate(state.f.desde) || 'inicio'} a ${fmtDate(state.f.hasta) || 'hoy'}` : 'Todo el historial';
    const ub = filtroUbicacion();
    const contexto = [varias() ? (state.f.granjas.length ? state.f.granjas.map(granjaName).join(', ') : 'Todas las granjas') : granjaName(state.granjaId),
        ub.SectorID && sectorById(ub.SectorID)?.Nombre, ub.NaveID && naveById(ub.NaveID)?.Nombre].filter(Boolean).join(' · ');

    let html = head;
    if (r.data.truncado) html += alertBox('warn', 'alert-triangle', 'Hay más registros de los que se pueden analizar a la vez. Reduce el rango de fechas para ver todo.');
    if (!regs.length) return html + sinEstructura(linea) + emptyState(`No hay registros de ${linea.toLowerCase()} en el rango seleccionado.`, 'chart-line');

    // KPIs del rango completo
    const tot = aggregate(regs, campos, () => 'total')[0];
    const ind = indicadoresLotes(regs, linea);
    const pc = k => campoDe(linea, k);
    let k = kpi('Registros', nf.format(regs.length), rango, 'clipboard-list')
        + kpi('Aves vivas', nf.format(ind.avesVivas), `${nf.format(ind.lotes)} ${ind.lotes === 1 ? 'lote abierto' : 'lotes abiertos'} · último día registrado`, 'bird');
    if (linea === 'Incubadora') {
        const inc = gSum(tot, 'HuevoIncubable'), nac = gSum(tot, 'Nacimiento');
        k += kpi('Huevo incubable', fmtNum(inc), `Transferencia ${fmtNum(gSum(tot, 'HuevoTransferencia'))}`, 'egg')
            + kpi('Nacimiento', fmtNum(nac), inc ? `${fmtNum(round(nac / inc * 100))} % del incubable` : '', 'trending-up')
            + kpi('Peso de huevo', fmtPeso(gAvg(tot, 'PesoHuevo')), 'Promedio de los pesajes', 'scale');
    } else {
        k += kpi('Mortalidad del rango', fmtNum(gSum(tot, 'MortalidadNominal')), ind.mortRango !== null ? `${fmtNum(round(ind.mortRango, 2))} % de las aves al inicio del rango` : '', 'activity')
            + kpi('Mortalidad acumulada', ind.mortAcum !== null ? `${fmtNum(round(ind.mortAcum, 2))} %` : '—', ind.mortAcum !== null ? `${nf.format(ind.muertasAcum)} de ${nf.format(ind.iniciales)} aves iniciales (lotes abiertos)` : 'Los lotes abiertos no tienen aves iniciales', 'activity')
            + kpi('Consumo', `${fmtNum(gSum(tot, 'ConsumoQuintalesTotal'))} qq`, `${fmtNum(round(gAvg(tot, 'ConsumoGramosAveDia'), 1))} g/ave/día (ponderado por aves)`, 'wheat')
            + kpi('Peso corporal', fmtPeso(ind.peso), ind.peso !== null ? `Tabla ${fmtPeso(ind.pesoTabla)} · último pesaje de cada lote abierto` : 'Sin pesajes de lotes abiertos', 'scale')
            + (linea === 'Postura'
                ? kpi('Alimento por docena', fmtPeso(gAvg(tot, 'AlimentoPorDocena')), 'Alimento ÷ docenas de huevo', 'wheat')
                    + kpi('Conversión', fmtNum(round(gAvg(tot, 'ConversionAlimenticia'), 3)), gAvg(tot, 'ConversionAlimenticia') !== null ? 'kg de alimento ÷ kg de huevo' : 'Anote el peso promedio del huevo para calcularla', 'wheat')
                : kpi('Conversión', fmtNum(round(gAvg(tot, 'ConversionAlimenticia'), 3)), 'kg de alimento ÷ kg ganados', 'wheat'));
        if (linea === 'Postura') k += kpi('Postura', `${fmtNum(round(gAvg(tot, 'PosturaPorcentaje'), 2))} %`, `Huevos ÷ aves del día · tabla ${fmtNum(round(gAvg(tot, 'PosturaTabla'), 2))}`, 'egg');
    }
    html += `<div class="kpis">${k}</div>`;

    // Panorama (como un tablero de Power BI): medidor, dona y comparativo horizontal.
    html += panorama(linea, regs, campos, tot, rango);

    // Tendencias por periodo
    // Por semana: semana de EDAD del lote ("Semana 23"); día y mes, de calendario.
    const grupos = aggregate(regs, campos.concat(CAMPOS_BANDA), x => (agr === 'semana' ? claveSemana(x) : periodo(x.Fecha, agr)));
    const labels = grupos.map(g => periodoLabel(g.key, agr));
    const sub = `${contexto} · ${rango} · por ${AGRUPAR.find(a => a[0] === agr)[1].toLowerCase()}`;
    html += '<div class="grid-2">' + CHARTS[linea].map((def) => {
        const spec = {
            title: def.peso ? `${def.title} (${state.unidad})` : def.title, subtitle: sub, icon: def.icon, type: def.type,
            unit: def.peso ? state.unidad : def.unit, labels,
            series: def.series.map(([key, name, color, dashed]) => ({
                name, color: PALETTE[color], dashed, stack: def.stack ? 'h' : undefined,
                data: grupos.map(g => { const c = pc(key); return c ? gShow(g, c) : null; }),
            })),
        };
        // Banda mín–máx de la tabla de la línea genética (solo si hay datos).
        if (def.banda) {
            const lo = grupos.map(g => gShow(g, CAMPOS_BANDA[0])), hi = grupos.map(g => gShow(g, CAMPOS_BANDA[1]));
            if (lo.some(v => v !== null)) spec.band = { name: 'Rango de la tabla', lo, hi, color: PALETTE.rango };
        }
        return chartCard(`${linea}-${def.id}`, spec);
    }).join('') + '</div>';

    // Comparativo: por granja (admin con varias) o por nave (una granja), en tabla.
    const variasGranjas = varias() && new Set(regs.map(x => x.GranjaID)).size > 1;
    const porNave = !variasGranjas && new Set(regs.map(x => x.NaveID).filter(Boolean)).size > 1;
    if (variasGranjas || porNave) {
        const clave = variasGranjas ? (x => x.GranjaID) : (x => x.NaveID || '—');
        const nombre = variasGranjas ? granjaName : nombreNave;
        const grupo = aggregate(regs, campos, clave);
        html += table('comparativo', {
            title: `Comparativo por ${variasGranjas ? 'granja' : 'nave'}`, icon: variasGranjas ? 'warehouse' : 'house', csv: `comparativo-${linea}`,
            columns: [{ k: 'nombre', l: variasGranjas ? 'Granja' : 'Nave', sticky: true }, { k: 'n', l: 'Registros', t: 'num' }]
                .concat(campos.map(c => ({ k: c.k, l: `${etiquetaCampo(c)}${sufijoAgg(c)}`, t: 'num' }))),
            rows: grupo.map(g => Object.assign({ nombre: nombre(g.key), n: g.n }, ...campos.map(c => ({ [c.k]: gShow(g, c) })))),
        });
    }

    // Resumen por periodo
    html += table('resumen', {
        title: `Resumen por ${AGRUPAR.find(a => a[0] === agr)[1].toLowerCase()}`, icon: 'calendar-range', csv: `resumen-${linea}-${agr}`,
        columns: [{ k: 'key', l: 'Periodo', sticky: true, html: x => esc(periodoLabel(x.key, agr)), csv: x => x.key }, { k: 'n', l: 'Registros', t: 'num' }]
            .concat(campos.map(c => ({ k: c.k, l: `${etiquetaCampo(c)}${sufijoAgg(c)}`, t: 'num' }))),
        rows: grupos.map(g => Object.assign({ key: g.key, n: g.n }, ...campos.map(c => ({ [c.k]: gShow(g, c) })))),
    });
    return html;
}

function nombreNave(id) {
    const n = naveById(id);
    return n ? `${sectorById(n.SectorID)?.Nombre || ''} / ${n.Nombre}` : '—';
}

/**
 * Fila de panorama: un medidor con el indicador principal de la línea, una dona (partes de un
 * total) y barras horizontales que comparan granjas (admin con varias) o naves (una granja).
 * Cada gráfico mide UNA sola cosa (sin mezclar unidades en un mismo eje).
 */
function panorama(linea, regs, campos, tot, rango) {
    const pc = k => campoDe(linea, k);
    const tarjetas = [];

    // 1. Medidor
    let g;
    if (linea === 'Postura') {
        const tabla = gAvg(tot, 'PosturaTabla');
        g = { title: 'Postura promedio', value: round(gAvg(tot, 'PosturaPorcentaje'), 1), max: 100, unit: '%', target: tabla !== null && tabla <= 100 ? round(tabla, 1) : null, targetLabel: 'Tabla', caption: rango };
    } else if (linea === 'Incubadora') {
        const inc = gSum(tot, 'HuevoIncubable'), nac = gSum(tot, 'Nacimiento');
        g = { title: 'Nacimiento sobre huevo incubable', value: inc ? round(nac / inc * 100, 1) : null, max: 100, unit: '%', caption: rango };
    } else {
        // Último pesaje de cada lote abierto (ponderado por aves) contra su tabla: no se mezclan edades.
        const ind = indicadoresLotes(regs, linea);
        const v = ind.peso !== null && ind.pesoTabla ? ind.peso / ind.pesoTabla * 100 : null;
        g = { title: 'Peso corporal frente a la tabla', value: round(v, 1), max: Math.max(120, Math.ceil(((v || 0) * 1.1) / 10) * 10), unit: '%', target: 100, targetLabel: 'Meta (100 % de la tabla)', caption: 'Último pesaje de cada lote abierto' };
    }
    tarjetas.push(chartCard(`${linea}-medidor`, { type: 'gauge', icon: 'activity', subtitle: rango, ...g }));

    // 2. Dona
    let d;
    if (linea === 'Postura') {
        const k = ['HuevoPiwi', 'HuevoMediano', 'HuevoGrande', 'HuevoExtragrande'];
        d = { title: 'Clasificación de huevo', labels: ['Piwi', 'Mediano', 'Grande', 'Extragrande'], values: k.map(x => gSum(tot, x) || 0), colors: [PALETTE.o1, PALETTE.o2, PALETTE.o3, PALETTE.o4], centerLabel: 'huevos' };
    } else if (linea === 'Incubadora') {
        const inc = gSum(tot, 'HuevoIncubable') || 0, nac = gSum(tot, 'Nacimiento') || 0;
        d = { title: 'Huevo incubable: nacidos y no nacidos', labels: ['Nacimiento', 'No nacidos'], values: [nac, Math.max(0, inc - nac)], colors: [PALETTE.primary, PALETTE.tabla], centerLabel: 'huevos incubables' };
    } else {
        // Aves vivas por nave (o por granja): al final del último día registrado de cada lote abierto.
        const porGranja = varias() && new Set(regs.map(x => x.GranjaID)).size > 1;
        const ult = new Map();
        regs.forEach((x) => { const k = claveLote(x); if (!ult.has(k) || String(x.Fecha) > String(ult.get(k).Fecha)) ult.set(k, x); });
        const suma = new Map();
        ult.forEach((x) => {
            const l = x.LoteID && loteById(x.LoteID);
            if (l ? l.Estado !== 'Activo' : naveById(x.NaveID)?.Estado !== 'Activo') return;
            const v = Math.max(0, (toNum(x.NumeroAves) || 0) - (toNum(x.MortalidadNominal) || 0) - (toNum(x.OtrasSalidas) || 0));
            const k = porGranja ? granjaName(x.GranjaID) : nombreNave(x.NaveID);
            suma.set(k, (suma.get(k) || 0) + v);
        });
        let filas = [...suma.entries()].sort((a, b) => b[1] - a[1]);
        if (filas.length > 6) filas = filas.slice(0, 5).concat([['Otras', filas.slice(5).reduce((a, f) => a + f[1], 0)]]);
        d = { title: `Aves vivas por ${porGranja ? 'granja' : 'nave'} (lotes abiertos)`, labels: filas.map(f => f[0]), values: filas.map(f => f[1]),
            colors: filas.map((f, i) => (f[0] === 'Otras' ? PALETTE.tabla : PALETTE.cat[i])), centerLabel: 'aves' };
    }
    tarjetas.push(chartCard(`${linea}-dona`, { type: 'donut', icon: 'chart-column', subtitle: rango, ...d }));

    // 3. Comparativo horizontal (solo si hay más de una granja o nave)
    const variasGranjas = varias() && new Set(regs.map(x => x.GranjaID)).size > 1;
    const porNave = !variasGranjas && new Set(regs.map(x => x.NaveID).filter(Boolean)).size > 1;
    if (variasGranjas || porNave) {
        const [real, tabla, titulo] = COMPARATIVO[linea];
        const cReal = pc(real), cTabla = tabla && pc(tabla);
        const grupo = aggregate(regs, campos, variasGranjas ? (x => x.GranjaID) : (x => x.NaveID || '—'))
            .map(gr => ({ nombre: variasGranjas ? granjaName(gr.key) : nombreNave(gr.key), v: gShow(gr, cReal), t: cTabla ? gShow(gr, cTabla) : null }))
            .sort((a, b) => (b.v ?? -Infinity) - (a.v ?? -Infinity));
        const unidad = cReal.peso ? state.unidad : linea === 'Postura' || linea === 'Incubadora' ? '%' : '';
        tarjetas.push(chartCard(`${linea}-comparativo`, {
            type: 'hbar', title: `${titulo} por ${variasGranjas ? 'granja' : 'nave'}`, subtitle: rango, icon: variasGranjas ? 'warehouse' : 'house',
            labels: grupo.map(x => x.nombre), values: grupo.map(x => x.v), unit: unidad,
            target: cTabla ? grupo.map(x => x.t) : null, valueName: 'Real (promedio)', targetName: 'Tabla (promedio)',
        }));
    }
    return `<div class="grid-3 panorama">${tarjetas.join('')}</div>`;
}

function viewRegistros() {
    const ids = ctxIds();
    if (!ids.length) return noGranjas();
    const head = filtersCard({ eliminados: true, unidad: true });
    if (!lineasDisponibles(ids).length) return head + sinLineas();

    const linea = state.linea;
    const aviso = sinEstructura(linea);
    const nuevo = aviso || `<div class="row row-center"><button type="button" class="btn btn-primary" data-action="new-registro">${ico('plus')}Nuevo registro de ${linea.toLowerCase()}</button>${botonImportar('registro')}${botonFormatoSemanal(linea)}</div>`;
    const r = need('getRegistros', registrosPayload({ incluirEliminados: isAdmin() && state.f.eliminados ? true : undefined }));
    if (!r) return head + nuevo + loadingBlock(1);
    if (!r.ok) return head + nuevo + errorBlock(r);

    state.registros = new Map(r.data.registros.map(x => [x.RegistroID, x]));
    const columns = [{ k: 'Fecha', l: 'Fecha', t: 'date', sticky: true }]
        .concat(colsGranja(), COLS_UBICACION, [{ k: 'LineaGenetica', l: 'Línea genética' }])
        .concat(camposDe(linea).map(c => (c.k !== 'PesoCorporalTabla' ? colCampo(c) : {
            // Peso tabla: de la línea genética; "manual" si es un valor escrito antes de esta versión.
            ...colCampo(c), csv: x => enUnidad(x.PesoCorporalTabla),
            html: x => `${fmtNum(round(enUnidad(x.PesoCorporalTabla), decimalesPeso()))}${x.PesoCorporalTablaOrigen === 'Manual' ? ' <span class="badge badge-muted" title="Escrito a mano antes de usar la tabla de la línea genética">manual</span>' : ''}`,
        })))
        .concat(isAdmin() && state.f.eliminados ? [{ k: 'Estado', l: 'Estado', html: x => badgeEstado(x.Estado === 'Activo' ? 'Activo' : 'Eliminado') }] : [])
        .concat([{ k: 'CreadoPor', l: 'Registrado por' }, { k: 'ActualizadoEn', l: 'Última edición', t: 'datetime' }]);

    let html = head + nuevo;
    if (r.data.truncado) html += alertBox('warn', 'alert-triangle', `Se muestran los ${nf.format(r.data.registros.length)} registros más recientes de ${nf.format(r.data.total)}. Reduce el rango de fechas para ver los demás.`);
    html += table('registros', {
        title: `Registros de ${linea.toLowerCase()}`, icon: LINEA_ICON[linea], csv: `registros-${linea}`,
        empty: `No hay registros de ${linea.toLowerCase()} en el rango seleccionado.`,
        columns, rows: r.data.registros,
        rowClass: x => (x.Estado !== 'Activo' ? 'is-deleted' : ''),
        actions: (x) => {
            const flag = x.Alterado ? `<span class="badge badge-danger" title="Modificado fuera de la plataforma">${ico('lock', 'icon-sm')}</span>` : '';
            if (x.Estado !== 'Activo') {
                return flag + (isAdmin() ? `<button type="button" class="icon-btn icon-btn-sm" data-action="restore-registro" data-id="${esc(x.RegistroID)}" title="Restaurar" aria-label="Restaurar registro">${ico('rotate-ccw')}</button>` : '');
            }
            if (!modificable(x.Fecha)) return flag + sinPermisoFecha();
            return flag + `<button type="button" class="icon-btn icon-btn-sm" data-action="edit-registro" data-id="${esc(x.RegistroID)}" title="Editar" aria-label="Editar registro">${ico('pencil')}</button>
                <button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-registro" data-id="${esc(x.RegistroID)}" title="Eliminar" aria-label="Eliminar registro">${ico('trash')}</button>`;
        },
    });
    return html;
}

function viewPesajes() {
    const ids = ctxIds();
    if (!ids.length) return noGranjas();
    const head = filtersCard({ unidad: true, eliminados: true });
    if (!lineasDisponibles(ids).length) return head + sinLineas();
    const linea = state.linea;
    const tipo = linea === 'Incubadora' ? 'huevo' : 'ave';
    const aviso = sinEstructura(linea);
    const nuevo = aviso || `<div class="section-head"><p class="muted">Se elige la nave, la unidad y el tamaño de la muestra, y se pesa ${tipo === 'huevo' ? 'cada huevo' : 'cada ave'} por separado. El promedio pasa al registro de ese día.</p>
        <div class="row">${botonImportar('pesaje')}<button type="button" class="btn btn-primary" data-action="new-pesaje">${ico('scale')}Nuevo pesaje</button></div></div>`;
    const r = need('getPesajes', registrosPayload({ incluirEliminados: isAdmin() && state.f.eliminados ? true : undefined }));
    if (!r) return head + nuevo + loadingBlock(1);
    if (!r.ok) return head + nuevo + errorBlock(r);

    state.pesajes = new Map(r.data.pesajes.map(x => [x.PesajeID, x]));
    const pesoCol = (k, l) => ({ k, l: `${l} (${state.unidad})`, t: 'num', get: x => round(enUnidad(x[k]), decimalesPeso()) });
    const columns = [{ k: 'Fecha', l: 'Fecha', t: 'date', sticky: true }]
        .concat(colsGranja(), COLS_UBICACION, [
            { k: 'Muestra', l: 'Muestra', t: 'num' },
            pesoCol('PromedioG', 'Promedio'), pesoCol('MinimoG', 'Mínimo'), pesoCol('MaximoG', 'Máximo'),
            { k: 'CV', l: 'CV (%)', t: 'num' }, { k: 'Uniformidad', l: 'Uniformidad (%)', t: 'num' },
        ])
        .concat(linea === 'Incubadora' ? [] : [
            // Comparación con la curva de la línea genética a la edad de las aves.
            { k: 'SemanaEdad', l: 'Semana de edad', t: 'num' },
            pesoCol('EsperadoG', 'Esperado'),
            { k: 'DiferenciaG', l: `Diferencia (${state.unidad})`, t: 'num', get: x => round(enUnidad(x.DiferenciaG), decimalesPeso()) },
            { k: 'DiferenciaPct', l: 'Diferencia (%)', t: 'num' },
            { k: 'EstadoPeso', l: 'Contra la tabla', html: x => badgeEstadoPeso(x.EstadoPeso) },
            { k: 'ConversionAlimenticia', l: 'Conversión alimenticia', t: 'num' },
        ])
        .concat([{ k: 'Unidad', l: 'Capturado en' }, { k: 'Pesos', l: 'Pesos individuales', nocsv: false, wrap: true, html: x => `<span class="muted">${esc(String(x.Pesos || '').split(',').length)} pesos</span>` },
            { k: 'CreadoPor', l: 'Registrado por' }])
        .concat(isAdmin() && state.f.eliminados ? [{ k: 'Estado', l: 'Estado', html: x => badgeEstado(x.Estado === 'Activo' ? 'Activo' : 'Eliminado') }] : []);

    return head + nuevo + table('pesajes', {
        title: `Pesajes de ${linea.toLowerCase()}`, icon: 'scale', csv: `pesajes-${linea}`,
        empty: 'No hay pesajes en el rango seleccionado.', columns, rows: r.data.pesajes,
        rowClass: x => (x.Estado !== 'Activo' ? 'is-deleted' : ''),
        actions: x => `<button type="button" class="icon-btn icon-btn-sm" data-action="ver-pesaje" data-id="${esc(x.PesajeID)}" title="Ver los pesos" aria-label="Ver los pesos">${ico('list-checks')}</button>`
            + (x.Estado === 'Activo' && !modificable(x.Fecha) ? sinPermisoFecha() : '')
            + (x.Estado === 'Activo' && modificable(x.Fecha) ? `<button type="button" class="icon-btn icon-btn-sm" data-action="edit-pesaje" data-id="${esc(x.PesajeID)}" title="Editar" aria-label="Editar pesaje">${ico('pencil')}</button>
               <button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-pesaje" data-id="${esc(x.PesajeID)}" title="Eliminar" aria-label="Eliminar pesaje">${ico('trash')}</button>` : ''),
    });
}

function viewAmbiente() {
    const ids = ctxIds();
    if (!ids.length) return noGranjas();
    const head = filtersCard();
    if (!lineasDisponibles(ids).length) return head + sinLineas();
    const linea = state.linea;
    if (!incluye('estres', ids, linea)) return head + sinPlan('estres');
    const hayNaves = ids.some(id => navesActivas(id, linea).length);
    const acciones = hayNaves
        ? `<div class="section-head"><p class="muted">Índice de estrés calórico (IEC) = temperatura ambiente en °C + humedad relativa en %. Cada lectura queda guardada con su fecha y hora.</p>
            <div class="row">${botonImportar('lectura')}<button type="button" class="btn btn-primary" data-action="new-lectura">${ico('thermometer')}Nueva lectura</button></div></div>`
        : `<div class="alert alert-info">${ico('info')}<span>Para registrar lecturas se necesita al menos una nave activa de ${esc(linea.toLowerCase())}.</span></div>`;
    const leyenda = `<div class="card"><div class="card-head card-title-center"><h3>${ico('thermometer')}Escala del índice de estrés calórico</h3></div>
        <div class="iec-scale">${state.db.iecRangos.map((x, i) => {
        const prev = state.db.iecRangos[i - 1];
        const rango = !prev ? `Menor a ${x.hasta}` : x.hasta === null ? `Mayor a ${Math.floor(prev.hasta)}` : `${prev.hasta} a ${Math.floor(x.hasta)}`;
        return `<div class="iec-item iec-${esc(x.color)}"><strong>${esc(rango)}</strong><span class="iec-desc">${esc(x.descripcion)}</span><span class="muted">${esc(x.comportamiento)}</span></div>`;
    }).join('')}</div></div>`;
    const payload = { Linea: linea, GranjaIDs: apiGranjas(), ...filtroUbicacion(), desde: state.f.desde || undefined, hasta: state.f.hasta || undefined };
    const r = need('getAmbiente', payload);
    if (!r) return head + acciones + leyenda + loadingBlock(1);
    if (!r.ok) return head + acciones + leyenda + errorBlock(r);

    const lecturas = r.data.lecturas;
    state.lecturas = new Map(lecturas.map(x => [x.LecturaID, x]));
    return head + acciones + chartCard(`iec-${linea}`, heatmapSpec(lecturas)) + leyenda + table('lecturas', {
        title: 'Lecturas de temperatura y humedad', icon: 'thermometer', csv: `ambiente-${linea}`,
        empty: 'No hay lecturas en el rango seleccionado.',
        columns: [{ k: 'FechaHora', l: 'Fecha y hora', t: 'datetime', sticky: true }]
            .concat(colsGranja(), [{ k: 'Sector', l: 'Sector' }, { k: 'Nave', l: 'Nave' },
                { k: 'Temperatura', l: 'Temperatura (°C)', t: 'num' }, { k: 'Humedad', l: 'Humedad relativa (%)', t: 'num' },
                { k: 'Indice', l: 'IEC', t: 'num' }, { k: 'Nivel', l: 'Nivel', html: x => badgeIec(x.Indice, x.Nivel) },
                { k: 'CreadoPor', l: 'Registrado por' }]),
        rows: lecturas,
        actions: x => (!modificable(x.FechaHora) ? sinPermisoFecha() : `<button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-lectura" data-id="${esc(x.LecturaID)}" title="Eliminar" aria-label="Eliminar lectura">${ico('trash')}</button>`),
    });
}

/** Mapa de calor: filas = naves, columnas = días (o semanas si el rango es largo); celda = IEC máximo. */
function heatmapSpec(lecturas) {
    const dias = lecturas.map(x => String(x.FechaHora).slice(0, 10));
    const desde = state.f.desde || (dias.length ? dias.reduce((a, b) => (a < b ? a : b)) : today());
    const hasta = state.f.hasta || today();
    const todos = [];
    for (let d = desde; d <= hasta && todos.length < 1000; d = sumarDias(d, 1)) todos.push(d);
    const semanal = todos.length > 92;
    const cols = [...new Set(todos.map(d => (semanal ? periodo(d, 'semana') : d)))];
    const naves = new Map();
    lecturas.forEach((x) => {
        if (!naves.has(x.NaveID)) naves.set(x.NaveID, `${varias() && ctxIds().length > 1 ? granjaName(x.GranjaID) + ' · ' : ''}${x.Sector} / ${x.Nave}`);
    });
    const filas = [...naves.entries()].sort((a, b) => a[1].localeCompare(b[1], 'es', { numeric: true }));
    const max = {};
    const cuenta = {};
    lecturas.forEach((x) => {
        const d = String(x.FechaHora).slice(0, 10);
        const k = `${x.NaveID}|${semanal ? periodo(d, 'semana') : d}`;
        const v = toNum(x.Indice);
        if (v === null) return;
        if (!(k in max) || v > max[k]) max[k] = v;
        cuenta[k] = (cuenta[k] || 0) + 1;
    });
    const values = filas.map(([id]) => cols.map(c => (`${id}|${c}` in max ? max[`${id}|${c}`] : null)));
    const etiqueta = c => (semanal ? periodoLabel(c, 'semana') : fmtDate(c).slice(0, 5));
    return {
        type: 'heatmap', title: 'Mapa de calor del estrés calórico', icon: 'thermometer',
        subtitle: `${semanal ? 'IEC máximo por semana' : 'IEC máximo por día'} · ${fmtDate(desde)} a ${fmtDate(hasta)}`,
        rows: filas.map(f => f[1]), cols: cols.map(etiqueta), values,
        color: v => IEC_COLOR[nivelIec(v)?.color] || '#ccc',
        legend: state.db.iecRangos.map(x => ({ color: IEC_COLOR[x.color], label: x.nivel })).concat([{ color: PALETTE.track, label: 'Sin lectura' }]),
        tip: (ri, ci) => {
            const v = values[ri][ci];
            const k = `${filas[ri][0]}|${cols[ci]}`;
            return [filas[ri][1], `${semanal ? 'Semana del ' + rangoSemanaIso(cols[ci]) : fmtDate(cols[ci])}`,
                v === null ? 'Sin lectura' : `IEC máximo ${fmtNum(v)} · ${nivelIec(v)?.nivel || ''}`, v === null ? '' : `${cuenta[k]} lectura(s)`].filter(Boolean);
        },
    };
}

function sumarDias(ymd, n) {
    const [y, m, d] = ymd.split('-').map(Number);
    const x = new Date(Date.UTC(y, m - 1, d + n));
    return x.toISOString().slice(0, 10);
}

// ==========================================================================
// 9. VISTAS: GRANJAS, PERFILES Y USUARIOS
// ==========================================================================

function representantesDe(granjaId) {
    return state.db.users.filter(u => u.Rol === ROLES.REPRESENTANTE && String(u.GranjaIDs || '').split(',').includes(granjaId));
}

function viewEmpresas() {
    const rows = (state.db.empresas || []).map(e => ({
        ...e,
        _granjas: state.db.granjas.filter(g => g.EmpresaID === e.EmpresaID).length,
        _cuentas: state.db.users.filter(u => u.EmpresaID === e.EmpresaID).length,
    }));
    return `<div class="section-head"><p class="muted">Cada empresa cliente tiene sus granjas (cada una en su propio archivo) y sus cuentas. Su administrador es el dueño: si administra varias empresas, recibe un solo cobro por todas. Una empresa suspendida puede consultar, pero no ingresar datos.</p>
        <button type="button" class="btn btn-primary" data-action="new-empresa">${ico('plus')}Nueva empresa</button></div>`
        + table('empresas', {
            title: 'Empresas', icon: 'building', empty: 'Aún no hay empresas registradas.',
            columns: [
                { k: 'Nombre', l: 'Empresa', sticky: true, html: e => `<strong>${esc(e.Nombre)}</strong>${e.Alterado ? ` <span class="badge badge-danger" title="Modificada fuera de la plataforma">${ico('lock', 'icon-sm')}</span>` : ''}` },
                { k: 'DuenoID', l: 'Dueño (quien paga)', get: e => duenoName(e.DuenoID) },
                { k: 'Contacto', l: 'Contacto' },
                { k: 'NIT', l: 'NIT' },
                { k: 'Telefono', l: 'Teléfono' },
                { k: 'Email', l: 'Correo' },
                { k: '_granjas', l: 'Granjas', t: 'num' },
                { k: '_cuentas', l: 'Cuentas', t: 'num' },
                { k: 'Estado', l: 'Estado', html: e => badgeEstado(e.Estado) },
                { k: 'FechaAlta', l: 'Alta', t: 'date' },
            ],
            rows,
            actions: e => `<button type="button" class="icon-btn icon-btn-sm" data-action="empresa-granjas" data-id="${esc(e.EmpresaID)}" title="Ver sus granjas" aria-label="Ver granjas de la empresa">${ico('warehouse')}</button>
                <button type="button" class="icon-btn icon-btn-sm" data-action="edit-empresa" data-id="${esc(e.EmpresaID)}" title="Editar" aria-label="Editar empresa">${ico('pencil')}</button>
                <button type="button" class="icon-btn icon-btn-sm" data-action="mover-empresa" data-id="${esc(e.EmpresaID)}" title="Cambiar dueño" aria-label="Cambiar el dueño de la empresa">${ico('users')}</button>
                <button type="button" class="icon-btn icon-btn-sm" data-action="estado-empresa" data-id="${esc(e.EmpresaID)}" title="${e.Estado === 'Activa' ? 'Suspender' : 'Activar'}" aria-label="${e.Estado === 'Activa' ? 'Suspender' : 'Activar'} empresa">${ico(e.Estado === 'Activa' ? 'circle-pause' : 'circle-play')}</button>
                <button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-empresa" data-id="${esc(e.EmpresaID)}" title="Eliminar" aria-label="Eliminar empresa"${e._granjas || e._cuentas ? ' disabled' : ''}>${ico('trash')}</button>`,
        });
}

function viewGranjas() {
    const admin = isAdmin();
    // La plataforma puede ver las granjas de una sola empresa (desde la vista de empresas).
    if (admin && state.empresaFiltro && !empresaById(state.empresaFiltro)) state.empresaFiltro = '';
    const filtro = admin ? state.empresaFiltro : variasEmpresas() ? state.empresaId : '';
    const rows = state.db.granjas.filter(g => !filtro || g.EmpresaID === filtro).map(g => ({
        ...g,
        _lineas: LINEAS.filter(l => perfilDe(g.GranjaID, l)),
        _naves: state.db.naves.filter(n => n.GranjaID === g.GranjaID && n.Estado === 'Activo').length,
        _reps: representantesDe(g.GranjaID).length,
    }));
    const lineasBadges = g => (g._lineas.length ? g._lineas.map((l) => {
        const p = perfilDe(g.GranjaID, l);
        return `<span class="badge ${p.Estado === 'Activo' ? '' : 'badge-warn'}" title="${p.Estado === 'Activo' ? '' : 'Suspendida'}">${ico(LINEA_ICON[l], 'icon-sm')}${esc(l)}</span>`;
    }).join(' ') : '<span class="badge badge-muted">Sin líneas</span>');
    const empresas = state.db.empresas || [];
    const selector = admin && empresas.length > 1
        ? `<label class="field"><span>Empresa</span><select data-change="empresa-filtro"><option value="">Todas las empresas</option>${empresas.map(e =>
            `<option value="${esc(e.EmpresaID)}"${e.EmpresaID === filtro ? ' selected' : ''}>${esc(e.Nombre)}</option>`).join('')}</select></label>` : empresaPicker();
    return `<div class="section-head"><p class="muted">${admin ? 'Cada granja tiene sus propias líneas (reproductora, incubadora, engorde y postura), cada una con su ficha, sus sectores y sus naves.'
        : `Granjas de ${esc(variasEmpresas() && !state.empresaId ? 'sus empresas' : miEmpresa()?.Nombre || 'su empresa')}. Al agregar una línea a una granja elige su plan y su tramo de aves, y ve cuánto sube su pago antes de confirmar.`}</p>
        ${selector}${admin || esEmpresa() ? `<button type="button" class="btn btn-primary" data-action="new-granja">${ico('plus')}Nueva granja</button>` : ''}</div>`
        + table('granjas', {
            title: 'Granjas', icon: 'warehouse', empty: 'Aún no hay granjas registradas.',
            columns: [
                { k: 'NombreGranja', l: 'Granja', sticky: true, html: g => `<button type="button" class="btn btn-link" data-action="open-granja" data-id="${esc(g.GranjaID)}">${esc(g.NombreGranja)}</button>${g.Alterado ? ` <span class="badge badge-danger" title="Modificada fuera de la plataforma">${ico('lock', 'icon-sm')}</span>` : ''}` },
            ].concat(admin && empresas.length > 1 ? [{ k: 'EmpresaID', l: 'Empresa', get: g => empresaName(g.EmpresaID) }] : [], [
                { k: 'Propietario', l: 'Propietario' },
                { k: 'Ubicacion', l: 'Ubicación' },
                { k: '_lineas', l: 'Líneas', get: g => g._lineas.join(', '), html: lineasBadges },
                { k: '_naves', l: 'Naves activas', t: 'num' },
                { k: '_reps', l: 'Representantes', t: 'num' },
                { k: 'Estado', l: 'Estado', html: g => badgeEstado(g.Estado) },
                { k: 'FechaAlta', l: 'Alta', t: 'date' },
            ]),
            rows,
            actions: g => `<button type="button" class="icon-btn icon-btn-sm" data-action="open-granja" data-id="${esc(g.GranjaID)}" title="Ver granja" aria-label="Ver granja">${ico('arrow-up-right')}</button>
                ${admin ? `<button type="button" class="icon-btn icon-btn-sm" data-action="edit-granja" data-id="${esc(g.GranjaID)}" title="Editar" aria-label="Editar granja">${ico('pencil')}</button>
                <button type="button" class="icon-btn icon-btn-sm" data-action="estado-granja" data-id="${esc(g.GranjaID)}" title="${g.Estado === 'Activo' ? 'Suspender' : 'Activar'}" aria-label="${g.Estado === 'Activo' ? 'Suspender' : 'Activar'} granja">${ico(g.Estado === 'Activo' ? 'circle-pause' : 'circle-play')}</button>
                <button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-granja" data-id="${esc(g.GranjaID)}" title="Eliminar" aria-label="Eliminar granja">${ico('trash')}</button>` : ''}`,
        });
}

/**
 * Granja: datos generales + cada línea que tiene (ficha, sectores, naves y lotes).
 * Las líneas que la granja no tiene no se muestran (el admin puede agregarlas).
 */
function viewGranja(param) {
    const id = varias() ? param : state.granjaId;
    if (!state.db.granjas.length && !varias()) return noGranjas();
    const g = id && granjaById(id);
    if (!g) {
        return emptyState('Granja no encontrada.', 'warehouse', varias() ? `<button type="button" class="btn" data-action="nav" data-view="granjas">${ico('arrow-left')}Volver a granjas</button>` : '');
    }
    const admin = isAdmin();
    let html = '';
    if (!varias() && state.db.granjas.length > 1) html += `<div class="card filters-card"><div class="filters">${granjaSwitch()}</div></div>`;

    const meta = [
        g.Propietario && `<span>${ico('user', 'icon-sm')}${esc(g.Propietario)}</span>`,
        g.Ubicacion && `<span>${ico('map-pin', 'icon-sm')}${esc(g.Ubicacion)}</span>`,
        g.Telefono && `<span>${ico('phone', 'icon-sm')}${esc(g.Telefono)}</span>`,
        g.EmailContacto && `<span>${ico('mail', 'icon-sm')}${esc(g.EmailContacto)}</span>`,
        g.FechaAlta && `<span>${ico('calendar-range', 'icon-sm')}Alta ${esc(fmtDate(g.FechaAlta))}</span>`,
        admin && (state.db.empresas || []).length > 1 && `<span>${ico('building', 'icon-sm')}${esc(empresaName(g.EmpresaID))}</span>`,
    ].filter(Boolean).join('');
    // En escritorio: ícono, nombre y estado en una fila, con los botones a la derecha (styles.css).
    html += `<div class="card farm-hero">
        <div class="farm-head">
            <div class="brand-mark" aria-hidden="true">${ico('warehouse')}</div>
            <div class="farm-title"><h2>${esc(g.NombreGranja)}</h2>
                <div class="row row-center">${badgeEstado(g.Estado)}${g.Alterado ? `<span class="badge badge-danger">${ico('lock', 'icon-sm')}Modificada fuera de la plataforma</span>` : ''}</div></div>
            <div class="row row-center farm-actions">
                <button type="button" class="btn btn-sm" data-action="goto-dashboard" data-id="${esc(g.GranjaID)}">${ico('layout-dashboard')}Dashboard</button>
                <button type="button" class="btn btn-sm" data-action="goto-registros" data-id="${esc(g.GranjaID)}">${ico('clipboard-list')}Registros</button>
                ${esEmpresa() ? `<button type="button" class="btn btn-sm" data-action="edit-granja" data-id="${esc(g.GranjaID)}">${ico('pencil')}Editar</button>` : ''}
                ${admin ? `<button type="button" class="btn btn-sm" data-action="edit-granja" data-id="${esc(g.GranjaID)}">${ico('pencil')}Editar</button>
                <button type="button" class="btn btn-sm" data-action="estado-granja" data-id="${esc(g.GranjaID)}">${ico(g.Estado === 'Activo' ? 'circle-pause' : 'circle-play')}${g.Estado === 'Activo' ? 'Suspender' : 'Activar'}</button>` : ''}
            </div>
        </div>
        ${meta ? `<div class="meta">${meta}</div>` : ''}
        ${g.Notas ? `<p class="muted">${esc(g.Notas)}</p>` : ''}
    </div>`;

    const lineas = LINEAS.filter(l => perfilDe(g.GranjaID, l));
    if (!lineas.length) {
        html += emptyState(admin || esEmpresa() ? 'Esta granja aún no tiene líneas de producción. Agregue la primera abajo.' : 'Esta granja aún no tiene líneas de producción. La administración de su empresa las agrega.', 'clipboard-list');
    }
    html += lineas.map(l => lineaCard(g, l)).join('');

    // Agregar líneas: solo el admin (el representante edita la ficha de las que ya tiene).
    const faltan = LINEAS.filter(l => !perfilDe(g.GranjaID, l));
    if (faltan.length && (admin || esEmpresa())) {
        html += `<div class="card"><div class="card-head card-title-center"><h3>${ico('plus')}Agregar línea de producción</h3></div>
            <p class="muted center">Cada línea tiene su propia ficha. Solo se muestran las líneas que la granja tiene.</p>
            <div class="row row-center">${faltan.map(l => `<button type="button" class="btn btn-sm" data-action="new-perfil" data-granja="${esc(g.GranjaID)}" data-linea="${l}">${ico(LINEA_ICON[l])}${l}</button>`).join('')}</div></div>`;
    }

    if (admin || esEmpresa()) {
        const reps = representantesDe(g.GranjaID);
        html += `<div class="card"><div class="card-head card-title-center"><h3>${ico('users')}Representantes</h3></div>
            ${reps.length ? `<div class="chips">${reps.map(u => `<span class="chip">${ico('user', 'icon-sm')}${esc(u.Nombre)} · ${esc(u.Email)} ${badgeEstado(u.Estado)}</span>`).join('')}</div>`
                : '<p class="muted center">Esta granja aún no tiene representantes.</p>'}
            <div class="row row-center"><button type="button" class="btn btn-sm" data-action="new-user" data-granja="${esc(g.GranjaID)}">${ico('user-plus')}Agregar representante</button></div></div>`;
    }
    return html;
}

function lineaCard(g, linea) {
    const p = perfilDe(g.GranjaID, linea);
    const admin = isAdmin();
    const activa = p.Estado === 'Activo';
    const secciones = [];
    state.db.campos.perfil.forEach((c) => {
        let s = secciones.find(x => x.s === c.s);
        if (!s) { s = { s: c.s, campos: [] }; secciones.push(s); }
        s.campos.push(c);
    });
    const valor = (c) => {
        const v = p[c.k];
        if (v === '' || v === undefined) return '<dd class="empty">—</dd>';
        return `<dd>${esc(c.t === 'fecha' ? fmtDate(v) : c.k === 'EdadReproductoraSemanas' ? `${v} semanas` : v)}</dd>`;
    };
    const sectores = sectoresDe(g.GranjaID, linea);
    return `<div class="card linea-card${activa ? '' : ' is-disabled'}">
        <div class="card-head card-title-center"><h2>${ico(LINEA_ICON[linea])}Línea de ${esc(linea.toLowerCase())}</h2>${badgeEstado(p.Estado)}${planDeLinea(g.GranjaID, linea) ? `<span class="badge badge-solid" title="Plan de la línea">${ico('tag', 'icon-sm')}${esc(planDeLinea(g.GranjaID, linea).Nombre)}</span>` : ''}
            ${admin || esEmpresa() ? `<button type="button" class="btn btn-sm" data-action="plan-linea" data-granja="${esc(g.GranjaID)}" data-linea="${esc(linea)}">Cambiar plan</button>` : ''}${p.Alterado ? `<span class="badge badge-danger">${ico('lock', 'icon-sm')}Modificada fuera de la plataforma</span>` : ''}</div>
        <details class="ficha"><summary>${ico('clipboard-list', 'icon-sm')}Ficha de la línea${ico('chevron-down', 'icon-sm')}</summary>
            <div class="stack">${secciones.map(s => `<section class="profile-section"><h3>${ico(SECCION_ICON[s.s] || 'list')}${esc(s.s)}</h3>
                <dl class="kv">${s.campos.map(c => `<div><dt>${esc(c.l)}</dt>${valor(c)}</div>`).join('')}</dl></section>`).join('')}
            ${p.ActualizadoEn ? `<p class="muted center">Última actualización ${esc(fmtDateTime(p.ActualizadoEn))}</p>` : ''}
            ${!admin && activa ? `<div class="row row-center"><button type="button" class="btn btn-sm" data-action="edit-perfil" data-id="${esc(p.PerfilID)}">${ico('pencil')}Editar ficha</button></div>` : ''}
            ${admin ? `<div class="row row-center">
                <button type="button" class="btn btn-sm" data-action="edit-perfil" data-id="${esc(p.PerfilID)}">${ico('pencil')}Editar ficha</button>
                <button type="button" class="btn btn-sm" data-action="estado-perfil" data-id="${esc(p.PerfilID)}">${ico(activa ? 'circle-pause' : 'circle-play')}${activa ? 'Suspender línea' : 'Activar línea'}</button>
                <button type="button" class="btn btn-sm btn-danger" data-action="delete-perfil" data-id="${esc(p.PerfilID)}">${ico('trash')}Eliminar línea</button></div>` : ''}</div>
        </details>
        <div class="estructura">
            <div class="row row-between"><h3 class="estructura-titulo">${ico('layout-grid')}Sectores y naves</h3>
                ${activa ? `<button type="button" class="btn btn-sm" data-action="new-sector" data-granja="${esc(g.GranjaID)}" data-linea="${linea}">${ico('plus')}Nuevo sector</button>` : ''}</div>
            ${sectores.length ? sectores.map(sectorBlock).join('') : `<p class="muted center">${activa ? 'Aún no hay sectores. Crea el primero para registrar datos.' : 'Línea suspendida.'}</p>`}
        </div>
    </div>`;
}

function sectorBlock(s) {
    const naves = navesDe(s.SectorID);
    const activo = s.Estado === 'Activo';
    return `<section class="sector${activo ? '' : ' is-inactive'}">
        <div class="sector-head">
            <strong>${ico('layout-grid', 'icon-sm')}${esc(s.Nombre)}</strong>${activo ? '' : badgeEstado('Inactivo')}
            <div class="row">
                ${activo ? `<button type="button" class="btn btn-sm" data-action="new-nave" data-id="${esc(s.SectorID)}">${ico('plus')}Nave</button>
                <button type="button" class="btn btn-sm" data-action="new-naves" data-id="${esc(s.SectorID)}">${ico('plus')}Varias naves</button>` : ''}
                <button type="button" class="icon-btn icon-btn-sm" data-action="edit-sector" data-id="${esc(s.SectorID)}" title="Renombrar" aria-label="Renombrar sector">${ico('pencil')}</button>
                <button type="button" class="icon-btn icon-btn-sm" data-action="estado-sector" data-id="${esc(s.SectorID)}" title="${activo ? 'Desactivar' : 'Activar'}" aria-label="${activo ? 'Desactivar' : 'Activar'} sector">${ico(activo ? 'circle-pause' : 'circle-play')}</button>
                ${naves.length ? '' : `<button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-sector" data-id="${esc(s.SectorID)}" title="Eliminar" aria-label="Eliminar sector">${ico('trash')}</button>`}
            </div>
        </div>
        ${naves.length ? `<div class="naves">${naves.map(naveCard).join('')}</div>` : '<p class="muted center">Sin naves.</p>'}
    </section>`;
}

function naveCard(n) {
    const activa = n.Estado === 'Activo';
    const lote = loteActivo(n.NaveID);
    const historial = lotesDe(n.NaveID).filter(l => l.Estado !== 'Activo');
    const tieneLotes = lotesDe(n.NaveID).length > 0;
    const dias = lote ? edadDias(lote, today()) : null;
    const costo = lote ? toNum(lote.CostoPorAve) : null;
    const total = costo !== null && toNum(lote.AvesIniciales) ? costo * toNum(lote.AvesIniciales) : null;
    const otras = lote ? navesVigentes(lote).filter(x => x.NaveID !== n.NaveID).map(x => naveById(x.NaveID)?.Nombre).filter(Boolean) : [];
    const propia = lote ? navesDelLote(lote).find(x => x.NaveID === n.NaveID && !x.Hasta) : null;
    const loteHtml = lote
        ? `<div class="lote">${ico('tag', 'icon-sm')}<span><strong>Lote ${esc(lote.NumeroLote)}</strong> · ${esc(lote.LineaGenetica)}<br><span class="muted">Desde ${esc(fmtDate(lote.FechaInicio))}${otras.length && propia?.AvesIniciales ? ` · ${esc(fmtNum(propia.AvesIniciales))} aves aquí de ${esc(fmtNum(lote.AvesIniciales))}` : lote.AvesIniciales ? ` · ${esc(fmtNum(lote.AvesIniciales))} aves` : ''}</span>
            ${otras.length ? `<br><span class="muted">También en ${esc(otras.join(', '))} · se registra ${modoEn(lote, today()) === 'Lote' ? 'por lote' : 'por nave'}</span>` : ''}
            ${dias !== null ? `<br><span class="muted">Hoy: ${esc(textoEdad(dias))}</span>` : `<br><span class="muted">${n.Linea === 'Incubadora' ? '' : 'Sin edad inicial: edita el lote para comparar con la tabla.'}</span>`}
            ${costo !== null ? `<br><span class="muted">${esc(fmtQ(costo))} por ave${total !== null ? ` · ${esc(fmtQ(total))} el lote` : ''}</span>` : ''}</span></div>`
        : '<div class="lote muted">Sin lote activo</div>';
    return `<div class="nave${activa ? '' : ' is-inactive'}">
        <div class="nave-head"><strong>${ico('house', 'icon-sm')}${esc(n.Nombre)}</strong>${activa ? '' : badgeEstado('Inactivo')}</div>
        ${loteHtml}
        ${activa ? referenciasNaveHtml(n) : ''}
        <div class="row row-center">
            ${activa && !lote ? `<button type="button" class="btn btn-sm btn-primary" data-action="new-lote" data-id="${esc(n.NaveID)}">${ico('tag')}Iniciar lote</button>` : ''}
            ${lote ? `<button type="button" class="btn btn-sm" data-action="edit-lote" data-id="${esc(lote.LoteID)}">${ico('tag')}Editar lote</button>
                <button type="button" class="btn btn-sm" data-action="lote-naves" data-id="${esc(lote.LoteID)}">${ico('house')}Naves del lote</button>
                <button type="button" class="btn btn-sm" data-action="cerrar-lote" data-id="${esc(lote.LoteID)}">Cerrar lote</button>` : ''}
            <button type="button" class="icon-btn icon-btn-sm" data-action="edit-nave" data-id="${esc(n.NaveID)}" title="Editar nave (nombre y medidas)" aria-label="Editar nave">${ico('pencil')}</button>
            <button type="button" class="icon-btn icon-btn-sm" data-action="estado-nave" data-id="${esc(n.NaveID)}" title="${activa ? 'Desactivar' : 'Activar'} nave" aria-label="${activa ? 'Desactivar' : 'Activar'} nave">${ico(activa ? 'circle-pause' : 'circle-play')}</button>
            ${tieneLotes ? '' : `<button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-nave" data-id="${esc(n.NaveID)}" title="Eliminar nave" aria-label="Eliminar nave">${ico('trash')}</button>`}
        </div>
        ${historial.length ? `<details class="historial"><summary>Lotes anteriores (${historial.length})</summary><ul>${historial.map((l, i) => `<li><span>${esc(l.NumeroLote)} · ${esc(l.LineaGenetica)} · ${esc(fmtDate(l.FechaInicio))} a ${esc(fmtDate(l.FechaCierre))}</span>
            ${i === 0 && !lote ? `<button type="button" class="btn btn-link" data-action="reabrir-lote" data-id="${esc(l.LoteID)}">Reabrir</button>` : ''}</li>`).join('')}</ul></details>` : ''}
    </div>`;
}

function viewUsuarios() {
    const admin = isAdmin();
    const conEmpresa = admin && (state.db.empresas || []).length > 1;
    const granjasDe = u => (u.Rol === ROLES.ADMIN ? 'Todas' : u.Rol === ROLES.ADMIN_EMPRESA ? 'Todas las de su empresa'
        : String(u.GranjaIDs || '').split(',').filter(Boolean).map(granjaName).join(', '));
    return `<div class="section-head"><p class="muted">${admin ? 'Solo la administración crea cuentas.' : `Cuentas de ${esc(miEmpresa()?.Nombre || 'su empresa')}.`} La contraseña temporal se muestra una sola vez.</p>
        <button type="button" class="btn btn-primary" data-action="new-user">${ico('user-plus')}Nuevo usuario</button></div>`
        + table('usuarios', {
            title: 'Usuarios', icon: 'users',
            columns: [
                { k: 'Nombre', l: 'Nombre', sticky: true, html: u => `${esc(u.Nombre)}${u.Alterado ? ` <span class="badge badge-danger" title="Modificado fuera de la plataforma">${ico('lock', 'icon-sm')}</span>` : ''}` },
                { k: 'Email', l: 'Correo' },
                { k: 'Rol', l: 'Nivel', get: u => ROLE_LABEL[u.Rol] || u.Rol, html: u => `<span class="badge ${u.Rol === ROLES.REPRESENTANTE ? '' : 'badge-solid'}">${esc(ROLE_LABEL[u.Rol] || u.Rol)}</span>` },
            ].concat(conEmpresa ? [{ k: 'EmpresaID', l: 'Empresa', get: u => (u.EmpresaID ? empresaName(u.EmpresaID) : 'CHIIK') }] : [], [
                { k: 'GranjaIDs', l: 'Granjas', get: granjasDe, wrap: true },
                { k: 'Telefono', l: 'Teléfono' },
                { k: 'Estado', l: 'Estado', html: u => badgeEstado(u.Estado) },
                { k: 'DebeCambiarClave', l: 'Contraseña', get: u => (u.DebeCambiarClave === 'SI' ? 'Temporal' : 'Definitiva'), html: u => (u.DebeCambiarClave === 'SI' ? '<span class="badge badge-warn">Temporal</span>' : '<span class="badge badge-ok">Definitiva</span>') },
                { k: 'FechaCreacion', l: 'Creado', t: 'datetime' },
            ]),
            rows: state.db.users,
            actions: (u) => {
                const self = u.ID === state.me.ID;
                return `<button type="button" class="icon-btn icon-btn-sm" data-action="edit-user" data-id="${esc(u.ID)}" title="Editar" aria-label="Editar usuario">${ico('pencil')}</button>
                    <button type="button" class="icon-btn icon-btn-sm" data-action="estado-user" data-id="${esc(u.ID)}" title="${u.Estado === 'Activo' ? 'Suspender' : 'Activar'}" aria-label="${u.Estado === 'Activo' ? 'Suspender' : 'Activar'} usuario"${self ? ' disabled' : ''}>${ico(u.Estado === 'Activo' ? 'circle-pause' : 'circle-play')}</button>
                    <button type="button" class="icon-btn icon-btn-sm" data-action="reset-user" data-id="${esc(u.ID)}" title="Nueva contraseña temporal" aria-label="Restablecer contraseña"${self ? ' disabled' : ''}>${ico('key-round')}</button>
                    <button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-user" data-id="${esc(u.ID)}" title="Eliminar" aria-label="Eliminar usuario"${self ? ' disabled' : ''}>${ico('trash')}</button>`;
            },
        });
}

// ==========================================================================
// 9b. PLANES (plataforma: herramientas y precios · empresa: "Mi plan")
// ==========================================================================

/** Dólares: enteros sin decimales ("US$ 15") y, si no, con dos ("US$ 227.50"). */
const fmtUSD = (v) => {
    if (v === null || v === undefined) return 'A la medida';
    const n = toNum(v);
    return n === null ? '—' : `US$ ${Number.isInteger(n) ? nf.format(n) : nfQ.format(n)}`;
};
const fmtTramo = t => (!t ? '—' : t.Hasta === null ? `Más de ${fmtNum(t.Desde - 1)} aves` : `${fmtNum(t.Desde)}–${fmtNum(t.Hasta)} aves`);

function viewPlanes() {
    const r = need('getPlanes', {});
    if (!r) return loadingBlock(3);
    if (!r.ok) return errorBlock(r);
    const d = r.data;
    const admin = isAdmin();
    const conEmpresa = admin && (state.db.empresas || []).length > 1;
    let html = admin
        ? '<div class="section-head"><p class="muted">Cada línea de cada granja tiene su plan y su tramo de aves: el precio es fijo por ese tramo y el tramo es el tope de aves vivas. Una línea sin tramo elegido paga según sus aves vivas. Los cambios de herramientas y precios se aplican de inmediato y quedan en la bitácora.</p></div>'
        : `<div class="section-head"><p class="muted">Cada línea de sus granjas tiene su plan y su tramo de aves: paga el precio de ese tramo cada mes y no puede tener más aves vivas que su tope. Al cambiar de plan o de tramo verá cuánto cambia su pago antes de confirmar.</p></div>
            <div class="kpis">${kpi('Total mensual de sus líneas', fmtUSD(d.total), `${d.lineas.length} líneas · ciclo de ${d.diasCiclo} días`, 'tag')}</div>` + seccionPagosEmpresa();

    html += table('planes-lineas', {
        title: admin ? 'Plan de cada línea' : 'Sus líneas', icon: 'tag', empty: 'Aún no hay líneas de producción.',
        columns: [{ k: 'NombreGranja', l: 'Granja', sticky: true }]
            .concat(conEmpresa ? [{ k: 'EmpresaID', l: 'Empresa', get: l => empresaName(l.EmpresaID) }] : [], [
                { k: 'Linea', l: 'Línea', html: l => `${ico(LINEA_ICON[l.Linea], 'icon-sm')} ${esc(l.Linea)}${l.EstadoLinea === 'Activo' ? '' : ' ' + badgeEstado(l.EstadoLinea)}` },
                { k: 'Plan', l: 'Plan', html: l => (l.Plan ? `<span class="badge badge-solid">${esc(l.Plan)}</span>` : '<span class="badge badge-warn">Sin plan</span>') + (l.Alterado ? ` <span class="badge badge-danger" title="Modificado fuera de la plataforma">${ico('lock', 'icon-sm')}</span>` : '') },
                { k: 'aves', l: 'Aves vivas', t: 'num' },
                { k: '_tramo', l: 'Tramo (tope)', get: l => fmtTramo(l.tramo), html: l => esc(fmtTramo(l.tramo)) + (l.tramoElegido ? '' : ' <span class="badge badge-muted" title="Sin tope: paga según sus aves vivas">Según aves</span>') },
                { k: 'precio', l: 'Precio al mes', t: 'num', get: l => l.precio, html: l => esc(fmtUSD(l.precio)) + (l.precioManual ? ' <span class="badge badge-muted">Manual</span>' : '') },
                { k: 'Desde', l: 'Plan desde', t: 'date' },
            ]),
        rows: d.lineas,
        actions: l => `<button type="button" class="btn btn-sm" data-action="plan-linea" data-granja="${esc(l.GranjaID)}" data-linea="${esc(l.Linea)}">${ico('tag')}Cambiar plan o tramo</button>`,
    });

    // Qué incluye cada plan (la plataforma lo edita; la empresa lo consulta para elegir).
    const herr = d.herramientas;
    const celda = (pl, h) => (h.siempre
        ? `<td class="center">${ico('check', 'icon-sm')}</td>`
        : `<td class="center"><input type="checkbox" class="check-plan" data-plan="${esc(pl.PlanID)}" value="${esc(h.clave)}"${pl.Herramientas.includes(h.clave) ? ' checked' : ''}${admin ? '' : ' disabled'} aria-label="${esc(h.nombre)} en ${esc(pl.Nombre)}"></td>`);
    html += `<div class="card table-card"><div class="table-tools"><h3>${ico('list-checks')}Qué incluye cada plan</h3>
            ${admin ? `<div class="tools"><button type="button" class="btn btn-sm btn-primary" data-action="guardar-herramientas">${ico('check')}Guardar herramientas</button></div>` : ''}</div>
        <div class="table-wrap" id="matriz-planes"><table><thead><tr><th>Herramienta</th>${d.planes.map(pl => `<th class="center">${esc(pl.Nombre)}</th>`).join('')}</tr></thead><tbody>
        ${herr.map(h => `<tr><td>${esc(h.nombre)}${h.proximamente ? ' <span class="badge badge-muted">Próximamente</span>' : ''}${h.control === 'pantalla' ? ' <span class="badge badge-muted" title="Se oculta el botón">En pantalla</span>' : ''}</td>${d.planes.map(pl => celda(pl, h)).join('')}</tr>`).join('')}
        </tbody></table></div></div>`;

    // Precios por tramo de aves.
    const precio = (t, pl) => (admin
        ? `<td class="center"><input class="input input-precio" type="number" min="0" step="0.01" inputmode="decimal" data-plan="${esc(pl.PlanID)}" data-tramo="${esc(t.TramoID)}" value="${t.precios[pl.PlanID] ?? ''}" placeholder="A la medida" aria-label="Precio de ${esc(pl.Nombre)}, ${esc(fmtTramo(t))}"></td>`
        : `<td class="center">${esc(fmtUSD(t.precios[pl.PlanID]))}</td>`);
    html += `<div class="card table-card"><div class="table-tools"><h3>${ico('tag')}Precio mensual por tramo de aves (USD)</h3>
            ${admin ? `<div class="tools"><button type="button" class="btn btn-sm btn-primary" data-action="guardar-precios">${ico('check')}Guardar precios</button></div>` : ''}</div>
        <div class="table-wrap" id="matriz-precios"><table><thead><tr><th>Aves de la línea (tramo)</th>${d.planes.map(pl => `<th class="center">${esc(pl.Nombre)}</th>`).join('')}</tr></thead><tbody>
        ${d.tramos.map(t => `<tr><td>${esc(fmtTramo(t))}</td>${d.planes.map(pl => precio(t, pl)).join('')}</tr>`).join('')}
        </tbody></table></div>
        <p class="muted center">Cada tramo incluye su número inicial: 1,000 aves exactas caen en el segundo.${admin ? ' Un precio vacío es "a la medida": se fija un precio manual en la línea.' : ''}</p></div>`;
    return html;
}

/** Cotización de una línea con el plan elegido (lo que cambia el cobro). */
function cotizacionHtml(q) {
    if (!q) return '';
    const n = q.nuevo;
    const filas = [
        ['Aves vivas hoy', fmtNum(n.aves)],
        [n.tramoElegido ? 'Tramo contratado (tope)' : 'Tramo según sus aves', fmtTramo(n.tramo)],
        ['Precio actual', q.actual ? `${fmtUSD(q.actual.precio)} al mes (${esc(q.actual.Plan)})` : 'Sin plan (línea nueva)'],
        ['Precio nuevo', `${fmtUSD(n.precio)} al mes (${esc(n.Plan)})`],
    ];
    if (q.diferencia !== null) {
        filas.push(['Cambio en su cobro mensual', `${q.diferencia >= 0 ? '+' : '−'} US$ ${fmtNum(Math.abs(q.diferencia))}`]);
        if (q.totalAntes !== null && q.totalDespues !== null) filas.push(['Su pago mensual (todas sus líneas)', `${fmtUSD(q.totalAntes)} → <strong>${fmtUSD(q.totalDespues)}</strong>`]);
        if (q.prorrateo) filas.push([`Proporcional de los ${q.ciclo.diasRestantes} días que quedan del ciclo`, `${q.prorrateo >= 0 ? '+' : '−'} US$ ${fmtNum(Math.abs(q.prorrateo))} en el próximo cobro`]);
    }
    return `<dl class="kv">${filas.map(([k, v]) => `<div><dt>${esc(k)}</dt><dd>${v}</dd></div>`).join('')}</dl>
        ${q.cabe === false ? alertBox('danger', 'alert-triangle', `La línea tiene hoy ${fmtNum(n.aves)} aves vivas: elija un tramo que las cubra.`) : ''}
        ${n.precio === null ? alertBox('info', 'info', 'Más de 300,000 aves: el precio lo acuerda con CHIIK.') : ''}
        <p class="muted center">El ciclo de cobro va del ${esc(fmtDate(q.ciclo.inicio))} al ${esc(fmtDate(q.ciclo.fin))}. El cargo queda anotado en "Mi plan y pagos".</p>`;
}

/** Pide la cotización al servidor y la muestra en la ventana abierta. */
async function cotizarEn(form, gid, linea) {
    const box = form?.querySelector('[data-cotizacion]');
    const planId = form?.elements.PlanID?.value;
    const tramoId = form?.elements.TramoID?.value || '';
    if (!box || !planId) return;
    box.innerHTML = loadingBlock(1);
    try {
        const q = await api('cotizarLinea', { GranjaID: gid, Linea: linea, PlanID: planId, TramoID: tramoId || undefined });
        if (form.elements.PlanID.value === planId && (form.elements.TramoID?.value || '') === tramoId) box.innerHTML = cotizacionHtml(q);
    } catch (e) {
        box.innerHTML = alertBox('danger', 'alert-triangle', esc(e.message));
    }
}

/**
 * Plan y tramo de aves de una línea. El tramo es el precio fijo y el tope de aves vivas. La plataforma
 * puede dejarlo "según sus aves" (sin tope); la empresa siempre elige uno.
 */
function selectorPlan(actualId, tramoId = '', nueva = false) {
    const tramos = state.db.tramos || [];
    const tramoSel = tramoId || (nueva && !isAdmin() ? tramos[0]?.TramoID : '');
    return `<label class="field"><span>Plan <span class="req">*</span></span><select name="PlanID" required data-change="cotizar">
        ${actualId ? '' : '<option value="">Elija el plan</option>'}${(state.db.planes || []).map(pl => `<option value="${esc(pl.PlanID)}"${pl.PlanID === actualId ? ' selected' : ''}>${esc(pl.Nombre)}</option>`).join('')}</select></label>
        <label class="field"><span>Aves de la línea (tramo) <span class="req">*</span></span><select name="TramoID" data-change="cotizar"${isAdmin() ? '' : ' required'}>
        ${isAdmin() ? `<option value=""${tramoSel ? '' : ' selected'}>Según sus aves vivas (sin tope)</option>` : ''}${tramos.map(t => `<option value="${esc(t.TramoID)}"${t.TramoID === tramoSel ? ' selected' : ''}>${esc(fmtTramo(t))}</option>`).join('')}</select>
        <small class="field-hint">Paga el precio de este tramo cada mes, tenga las aves que tenga, y no puede tener más aves vivas que su tope.</small></label>
        <div class="span-all" data-cotizacion></div>`;
}

async function openPlanLineaModal(gid, linea) {
    const g = granjaById(gid);
    const actual = planDeLinea(gid, linea);
    const admin = isAdmin();
    const promesa = openModal({
        title: `Plan de ${linea.toLowerCase()} · ${g?.NombreGranja || ''}`, wide: true, submitLabel: 'Confirmar plan',
        body: `<div class="form-grid" data-granja="${esc(gid)}" data-linea="${esc(linea)}">${selectorPlan(actual?.PlanID, (state.db.lineasPlan || []).find(x => x.GranjaID === gid && x.Linea === linea)?.TramoID || '')}
            ${admin ? '<label class="field"><span>Precio manual (USD al mes)</span><input name="PrecioManual" type="number" min="0" step="0.01" inputmode="decimal" placeholder="Precio de lista"><small class="field-hint">Solo para clientes con precio acordado.</small></label>' : ''}
            ${admin ? '<label class="field span-all"><span>Motivo (opcional)</span><textarea name="justificacion" maxlength="1000"></textarea><small class="field-hint">Queda registrado en la bitácora.</small></label>' : ''}</div>`,
        onSubmit: async (fd) => {
            const p = { GranjaID: gid, Linea: linea, PlanID: fd.get('PlanID'), TramoID: fd.get('TramoID') || undefined };
            if (admin) { p.PrecioManual = fd.get('PrecioManual'); p.justificacion = fd.get('justificacion'); }
            await save('setPlanLinea', p, r => (r?.sinCambios ? 'La línea ya tenía ese plan y tramo.' : 'Plan actualizado. El cargo queda anotado en "Mi plan y pagos".'));
        },
    });
    cotizarEn($('#modal form'), gid, linea);
    await promesa;
}

async function guardarHerramientas() {
    const cambios = (state.db.planes || []).map((pl) => {
        const marcadas = [...$$(`#matriz-planes input[data-plan="${pl.PlanID}"]:checked`)].map(i => i.value);
        const antes = (pl.Herramientas || []).slice().sort().join();
        return marcadas.slice().sort().join() === antes ? null : { PlanID: pl.PlanID, Herramientas: marcadas, Nombre: pl.Nombre };
    }).filter(Boolean);
    if (!cambios.length) { toast('No hay cambios en las herramientas.'); return; }
    await openModal({
        title: 'Guardar herramientas de los planes', submitLabel: 'Guardar',
        body: `<p class="center">Cambian: <strong>${cambios.map(c => esc(c.Nombre)).join(', ')}</strong>. Se aplica de inmediato a todas las líneas con ese plan; los datos nunca se borran.</p>${justField()}`,
        onSubmit: async (fd) => {
            for (const c of cambios) await withBusy(() => api('savePlan', { PlanID: c.PlanID, Herramientas: c.Herramientas, justificacion: fd.get('justificacion') }));
            state.cache.clear();
            await withBusy(() => loadSnapshot());
            toast('Herramientas actualizadas.', 'success');
            renderChrome();
            render();
        },
    });
}

async function guardarPrecios() {
    const r = state.cache.get('getPlanes{}');
    const tramos = r?.ok ? r.data.tramos : [];
    const precios = [...$$('#matriz-precios input[data-plan]')].map((i) => {
        const t = tramos.find(x => x.TramoID === i.dataset.tramo);
        const antes = t ? t.precios[i.dataset.plan] : null;
        const ahora = i.value.trim();
        return String(antes ?? '') === ahora ? null : { PlanID: i.dataset.plan, TramoID: i.dataset.tramo, PrecioUSD: ahora };
    }).filter(Boolean);
    if (!precios.length) { toast('No hay cambios en los precios.'); return; }
    await openModal({
        title: 'Guardar precios', submitLabel: 'Guardar',
        body: `<p class="center">Cambian <strong>${precios.length}</strong> precio${precios.length === 1 ? '' : 's'}. Los cobros nuevos usarán estos precios.</p>${justField()}`,
        onSubmit: async fd => save('savePrecios', { precios, justificacion: fd.get('justificacion') }, 'Precios actualizados.'),
    });
}

// ==========================================================================
// 9c. COBROS Y PAGOS (plataforma: todas las empresas · empresa: "Mi plan y pagos")
// ==========================================================================

const ESTADO_CUENTA = {
    'al-dia': ['Al día', 'badge-ok'], 'por-vencer': ['Por vencer', ''], gracia: ['En gracia', 'badge-warn'], suspendida: ['Suspendida', 'badge-danger'],
};
const badgeCuenta = c => { const [t, cls] = ESTADO_CUENTA[c?.estado] || ['—', 'badge-muted']; return `<span class="badge ${cls}">${t}</span>`; };
const ESTADO_COBRO = { Pendiente: 'badge-warn', Pagado: 'badge-ok', 'Sin cargo': 'badge-muted' };
const badgeCobro = c => `<span class="badge ${c.Estado === 'Pendiente' && c.Saldo > 0 && c.VenceEfectivo < today() ? 'badge-danger' : ESTADO_COBRO[c.Estado] || ''}">${esc(c.Estado === 'Pendiente' && c.VenceEfectivo < today() ? 'Vencido' : c.Estado)}</span>`;
const periodoCobro = c => `Ciclo ${c.Periodo} · ${fmtDate(c.Desde)} al ${fmtDate(addDays(c.Hasta, -1))}`;
function addDays(ymd, n) { const d = new Date(`${ymd}T12:00:00`); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); }

function columnasCobros(conDueno) {
    return [{ k: 'Periodo', l: 'Periodo', sticky: true, get: periodoCobro, csv: c => c.Periodo }]
        .concat(conDueno ? [{ k: 'DuenoNombre', l: 'Dueño' }, { k: '_empresas', l: 'Empresas', get: c => c.Detalle.map(x => x.Nombre).join(', '), wrap: true }] : [], [
            { k: 'VenceEfectivo', l: 'Vence', t: 'date' },
            { k: 'Total', l: 'Total', t: 'num', get: c => Number(c.Total), html: c => esc(fmtUSD(Number(c.Total))) },
            { k: 'Pagado', l: 'Pagado', t: 'num', html: c => esc(fmtUSD(c.Pagado)) },
            { k: 'Saldo', l: 'Saldo', t: 'num', html: c => esc(fmtUSD(c.Saldo)) },
            { k: 'Estado', l: 'Estado', html: c => badgeCobro(c) + (c.ProrrogaHasta ? ' <span class="badge badge-muted">Prórroga</span>' : '') },
        ]);
}

/** Sección de "Mi plan y pagos": estado de la cuenta, cobro en curso, cobros y pagos. */
function seccionPagosEmpresa() {
    const r = need('getCobros', {});
    if (!r) return loadingBlock(2);
    if (!r.ok) return errorBlock(r);
    const d = r.data;
    const dueno = d.duenos[0];
    const cuenta = dueno?.cuenta;
    const enCurso = d.enCurso;
    let html = (dueno?.empresas.length > 1 ? `<p class="muted center">Un solo cobro para sus ${dueno.empresas.length} empresas (${esc(dueno.empresas.map(e => e.Nombre).join(', '))}), con el detalle de cada una.</p>` : '') + `<div class="kpis">
        ${kpi('Estado de su cuenta', ESTADO_CUENTA[cuenta?.estado]?.[0] || '—', cuenta?.saldo ? `Saldo pendiente ${fmtUSD(cuenta.saldo)}` : 'Sin saldo pendiente', 'receipt')}
        ${enCurso ? kpi('Ciclo en curso (hasta hoy)', fmtUSD(enCurso.total), `${fmtDate(enCurso.desde)} al ${fmtDate(addDays(enCurso.hasta, -1))} · vence el ${fmtDate(enCurso.hasta)}`, 'calendar-range') : ''}
    </div>`;
    html += table('mis-cobros', {
        title: 'Sus cobros', icon: 'receipt', empty: 'Aún no hay cobros: el primero se emite al cumplirse 30 días desde su alta.',
        columns: columnasCobros(false), rows: d.cobros,
        actions: c => `<button type="button" class="btn btn-sm" data-action="detalle-cobro" data-id="${esc(c.CobroID)}">${ico('eye')}Detalle</button>`,
    });
    if (d.pagos.length) {
        html += table('mis-pagos', {
            title: 'Pagos registrados', icon: 'circle-check',
            columns: [{ k: 'Fecha', l: 'Fecha', t: 'date', sticky: true }, { k: 'Monto', l: 'Monto', t: 'num', html: x => esc(fmtUSD(Number(x.Monto))) },
                { k: 'Medio', l: 'Medio' }, { k: 'Referencia', l: 'Referencia' }, { k: 'Estado', l: 'Estado', html: x => badgeEstado(x.Estado) }],
            rows: d.pagos,
        });
    }
    return html + `<p class="muted center">Los pagos (enlace Visa o transferencia) los registra la administración de CHIIK. Un cobro vencido da ${d.diasGracia} días de gracia; después el ingreso de datos se suspende hasta registrar el pago.</p>`;
}

function viewCobros() {
    const r = need('getCobros', {});
    if (!r) return loadingBlock(3);
    if (!r.ok) return errorBlock(r);
    const d = r.data;
    let html = `<div class="section-head"><p class="muted">Cada dueño (quien paga) recibe un solo cobro al cumplir cada ciclo de 30 días desde su alta, con el detalle por empresa, granja y línea (prorrateado por días). Aquí se registran los pagos manuales; al cubrir el saldo, todas sus empresas se reactivan de inmediato.</p></div>`;
    html += table('cuentas', {
        title: 'Dueños', icon: 'building',
        columns: [
            { k: 'Nombre', l: 'Dueño', sticky: true },
            { k: '_empresas', l: 'Empresas', get: e => e.empresas.map(x => x.Nombre).join(', '), wrap: true },
            { k: '_estado', l: 'Cuenta', get: e => ESTADO_CUENTA[e.cuenta.estado]?.[0], html: e => badgeCuenta(e.cuenta) + (e.Estado === 'Activo' ? '' : ' ' + badgeEstado(e.Estado)) },
            { k: '_saldo', l: 'Saldo pendiente', t: 'num', get: e => e.cuenta.saldo, html: e => esc(fmtUSD(e.cuenta.saldo)) },
            { k: '_vence', l: 'Próximo vencimiento', t: 'date', get: e => e.cuenta.pendientes[0]?.vence || e.cuenta.ciclo.fin },
            { k: '_aviso', l: 'Aviso', get: e => e.cuenta.avisos[0]?.mensaje || '', wrap: true },
        ],
        rows: d.duenos,
        actions: e => `<button type="button" class="icon-btn icon-btn-sm" data-action="editar-dueno" data-id="${esc(e.DuenoID)}" title="Datos de cobro" aria-label="Editar los datos de cobro del dueño">${ico('pencil')}</button>
            <button type="button" class="icon-btn icon-btn-sm" data-action="estado-dueno" data-id="${esc(e.DuenoID)}" title="${e.Estado === 'Activo' ? 'Suspender' : 'Activar'}" aria-label="${e.Estado === 'Activo' ? 'Suspender' : 'Activar'} al dueño">${ico(e.Estado === 'Activo' ? 'circle-pause' : 'circle-play')}</button>`,
    });
    html += table('cobros', {
        title: 'Cobros', icon: 'receipt', empty: 'Aún no hay cobros emitidos.',
        columns: columnasCobros(true), rows: d.cobros,
        actions: c => `<button type="button" class="icon-btn icon-btn-sm" data-action="detalle-cobro" data-id="${esc(c.CobroID)}" title="Detalle" aria-label="Detalle del cobro">${ico('eye')}</button>
            ${c.Estado === 'Pendiente' ? `<button type="button" class="icon-btn icon-btn-sm" data-action="registrar-pago" data-id="${esc(c.CobroID)}" title="Registrar pago" aria-label="Registrar pago">${ico('circle-check')}</button>
            <button type="button" class="icon-btn icon-btn-sm" data-action="prorroga" data-id="${esc(c.CobroID)}" title="Prórroga" aria-label="Dar prórroga">${ico('calendar-range')}</button>` : ''}`,
    });
    html += table('pagos', {
        title: 'Pagos', icon: 'circle-check', empty: 'Aún no hay pagos registrados.',
        columns: [{ k: 'Fecha', l: 'Fecha', t: 'date', sticky: true }, { k: 'DuenoID', l: 'Dueño', get: x => d.duenos.find(o => o.DuenoID === x.DuenoID)?.Nombre || empresaName(x.EmpresaID) },
            { k: 'Monto', l: 'Monto', t: 'num', html: x => esc(fmtUSD(Number(x.Monto))) }, { k: 'Medio', l: 'Medio' }, { k: 'Referencia', l: 'Referencia' },
            { k: 'RegistradoPor', l: 'Registrado por' }, { k: 'Estado', l: 'Estado', html: x => badgeEstado(x.Estado) }],
        rows: d.pagos,
        actions: x => (x.Estado === 'Activo' ? `<button type="button" class="icon-btn icon-btn-sm danger" data-action="anular-pago" data-id="${esc(x.PagoID)}" title="Anular" aria-label="Anular pago">${ico('x')}</button>` : ''),
    });
    return html;
}

const cobroById = id => state.cache.get('getCobros{}')?.data?.cobros.find(c => c.CobroID === id);

async function detalleCobroModal(c) {
    if (!c) return;
    const lineas = ls => ls.flatMap(l => (l.activa ? l.segmentos.map((sg, i) => `<tr><td>${i ? '' : esc(l.NombreGranja)}</td><td>${i ? '' : esc(l.Linea)}</td>
        <td>${esc(sg.Plan)}${sg.precioManual ? ' <span class="badge badge-muted">Manual</span>' : ''}</td><td>${esc(fmtTramo(sg.tramo))}</td>
        <td>${esc(fmtDate(sg.desde))} al ${esc(fmtDate(addDays(sg.hasta, -1)))}</td><td class="num">${sg.dias}</td><td class="num">${esc(fmtUSD(sg.precio))}</td><td class="num">${esc(fmtUSD(sg.monto))}</td></tr>`)
        : [`<tr class="muted"><td>${esc(l.NombreGranja)}</td><td>${esc(l.Linea)}</td><td colspan="5">Sin lotes ni datos en el ciclo: no se cobra</td><td class="num">${esc(fmtUSD(0))}</td></tr>`]));
    // Un bloque por empresa del dueño, con su subtotal (y los días, si solo fue suya parte del ciclo).
    const filas = c.Detalle.flatMap(emp => [
        `<tr class="fila-grupo"><th colspan="8">${ico('building', 'icon-sm')} ${esc(emp.Nombre)}${emp.desde !== c.Desde || emp.hasta !== c.Hasta ? ` <span class="muted">(del ${esc(fmtDate(emp.desde))} al ${esc(fmtDate(addDays(emp.hasta, -1)))})</span>` : ''}</th></tr>`,
        ...lineas(emp.lineas),
        ...(c.Detalle.length > 1 ? [`<tr class="fila-subtotal"><td colspan="7">Subtotal ${esc(emp.Nombre)}</td><td class="num">${esc(fmtUSD(emp.subtotal))}</td></tr>`] : []),
    ]);
    await openModal({
        title: `${periodoCobro(c)}${c.DuenoNombre ? ' · ' + c.DuenoNombre : ''}`, wide: true, cancelLabel: 'Cerrar',
        body: `<div class="table-wrap"><table><thead><tr><th>Granja</th><th>Línea</th><th>Plan</th><th>Tramo</th><th>Periodo</th><th class="num">Días</th><th class="num">Precio al mes</th><th class="num">Monto</th></tr></thead>
            <tbody>${filas.join('')}</tbody><tfoot><tr><th colspan="7">Total</th><th class="num">${esc(fmtUSD(Number(c.Total)))}</th></tr></tfoot></table></div>
            <p class="muted center">Monto = precio al mes × días ÷ 30. Vence el ${esc(fmtDate(c.VenceEfectivo))}. Pagado ${esc(fmtUSD(c.Pagado))} · saldo ${esc(fmtUSD(c.Saldo))}.</p>`,
    });
}

async function registrarPagoModal(c) {
    if (!c) return;
    const medios = state.cache.get('getCobros{}')?.data?.medios || [];
    await openModal({
        title: `Registrar pago · ${c.DuenoNombre || empresaName(c.EmpresaID)}`, submitLabel: 'Registrar pago',
        body: `<p class="center">${esc(periodoCobro(c))} · saldo <strong>${esc(fmtUSD(c.Saldo))}</strong></p><div class="form-grid">
            <label class="field"><span>Monto (USD) <span class="req">*</span></span><input name="Monto" type="number" min="0.01" step="0.01" max="${c.Saldo}" value="${c.Saldo}" required inputmode="decimal"></label>
            <label class="field"><span>Fecha del pago <span class="req">*</span></span><input name="Fecha" type="date" max="${today()}" value="${today()}" required></label>
            <label class="field"><span>Medio <span class="req">*</span></span><select name="Medio" required>${medios.map(m => `<option>${esc(m)}</option>`).join('')}</select></label>
            <label class="field"><span>Referencia</span><input name="Referencia" maxlength="120" placeholder="Número de autorización o boleta"></label>
            <label class="field span-all"><span>Notas</span><textarea name="Notas" maxlength="500"></textarea></label></div>`,
        onSubmit: async fd => save('registrarPago', { CobroID: c.CobroID, ...formPayload(fd) }, r => (r?.cobroPagado ? 'Pago registrado: cobro pagado y cuenta al día.' : 'Pago parcial registrado.')),
    });
}

async function prorrogaModal(c) {
    if (!c) return;
    await openModal({
        title: `Prórroga · ${c.DuenoNombre || empresaName(c.EmpresaID)}`, submitLabel: 'Guardar prórroga',
        body: `<p class="center">${esc(periodoCobro(c))}. Vence el ${esc(fmtDate(c.Vence))}. Con prórroga, los avisos y la gracia se cuentan desde la fecha nueva. Déjela vacía para quitarla.</p>
            <div class="form-grid"><label class="field span-all"><span>Nueva fecha de vencimiento</span><input name="Hasta" type="date" min="${addDays(c.Vence, 1)}" value="${esc(c.ProrrogaHasta || '')}"></label>${justField()}</div>`,
        onSubmit: async fd => save('setProrroga', { CobroID: c.CobroID, Hasta: fd.get('Hasta'), justificacion: fd.get('justificacion') }, 'Prórroga guardada.'),
    });
}

async function openDuenoModal(d) {
    if (!d) return;
    const f = (k, l, attrs = '', req = false) => `<label class="field"><span>${l}${req ? ' <span class="req">*</span>' : ''}</span><input name="${k}" value="${esc(d[k] || '')}" ${attrs}${req ? ' required' : ''}></label>`;
    await openModal({
        title: `Datos de cobro · ${d.Nombre}`, wide: true, submitLabel: 'Guardar cambios',
        body: `<div class="form-grid">
            ${f('Nombre', 'Nombre (a quien se cobra)', 'maxlength="150" autocomplete="off"', true)}
            ${f('NIT', 'NIT', 'maxlength="30"')}
            ${f('Contacto', 'Persona de contacto', 'maxlength="150"')}
            ${f('Telefono', 'Teléfono', 'type="tel" inputmode="tel" maxlength="30" pattern="[0-9+()\\-\\s]{6,30}"')}
            ${f('Email', 'Correo para cobros', 'type="email" maxlength="254"')}
            <label class="field span-all"><span>Notas</span><textarea name="Notas" maxlength="1000">${esc(d.Notas || '')}</textarea></label>
            ${justField()}</div>`,
        onSubmit: async fd => save('saveDueno', { DuenoID: d.DuenoID, ...formPayload(fd) }, 'Datos de cobro actualizados.'),
    });
}

async function estadoDuenoFlow(d) {
    if (!d) return;
    const activar = d.Estado !== 'Activo';
    await openModal({
        title: `${activar ? 'Activar' : 'Suspender'} a ${d.Nombre}`, danger: !activar, submitLabel: activar ? 'Activar' : 'Suspender',
        body: `<p class="center">${activar ? 'Todas sus empresas vuelven a registrar y modificar datos de inmediato.'
            : 'Todas sus empresas podrán ver y consultar su información, pero no registrar, editar, cargar ni eliminar datos.'}</p>${justField()}`,
        onSubmit: async fd => save('setDuenoEstado', { DuenoID: d.DuenoID, Estado: activar ? 'Activo' : 'Suspendido', justificacion: fd.get('justificacion') }, activar ? 'Dueño activado.' : 'Dueño suspendido.'),
    });
}

async function anularPagoModal(id) {
    await openModal({
        title: 'Anular pago', danger: true, submitLabel: 'Anular pago',
        body: `<p class="center">El pago queda anulado (no se borra) y su cobro vuelve a quedar pendiente.</p>${justField()}`,
        onSubmit: async fd => save('anularPago', { PagoID: id, justificacion: fd.get('justificacion') }, 'Pago anulado.'),
    });
}

// ==========================================================================
// 9d. RESUMEN GLOBAL Y DE EMPRESA (lee los resúmenes consolidados del central)
// ==========================================================================

/** Suma de semanas (de una o varias granjas). Los indicadores salen de las sumas, nunca de promedios de promedios. */
function sumaSemanas(semanas) {
    const t = { n: 0, avd: 0, pob: 0, mort: 0, alim: 0, huevos: 0, inc: 0, nac: 0, pesoP: 0, pesoA: 0, tablaP: 0, tablaA: 0, alimC: 0, ganC: 0 };
    semanas.forEach((w) => {
        ['n', 'avd', 'mort', 'alim', 'huevos', 'inc', 'nac', 'pesoP', 'pesoA', 'tablaP', 'tablaA', 'alimC', 'ganC'].forEach((k) => { t[k] += Number(w[k]) || 0; });
        if (w.dias) t.pob += w.avd / w.dias;   // aves promedio de esa semana
    });
    return t;
}
const indicadores = t => ({
    mort: t.pob ? t.mort / t.pob * 100 : null,
    consumo: t.avd ? t.alim * 1000 / t.avd : null,
    pesoVsTabla: t.pesoA && t.tablaA && t.tablaP ? (t.pesoP / t.pesoA) / (t.tablaP / t.tablaA) * 100 : null,
    conversion: t.ganC ? t.alimC / t.ganC : null,
    postura: t.avd ? t.huevos / t.avd * 100 : null,
    nacimiento: t.inc ? t.nac / t.inc * 100 : null,
});
const fmt2 = v => (v === null || v === undefined || !isFinite(v) ? '—' : nf.format(Math.round(v * 100) / 100));
const ULTIMAS = 4;

function viewResumen() {
    const r = need('getResumenes', {});
    if (!r) return loadingBlock(3);
    if (!r.ok) return errorBlock(r);
    const d = r.data;
    const admin = isAdmin();
    const empresas = state.db.empresas || [];
    if (admin && state.resEmpresa && !empresaById(state.resEmpresa)) state.resEmpresa = '';
    const filtroE = admin ? state.resEmpresa || '' : '';
    const base = d.resumenes.filter(x => !filtroE || x.EmpresaID === filtroE);
    const lineas = LINEAS.filter(l => base.some(x => x.Linea === l));
    // Al entrar, la línea con más aves vivas.
    const avesDe = l => base.filter(x => x.Linea === l).reduce((t, x) => t + (Number(x.ultimo.aves) || 0), 0);
    if (!lineas.includes(state.resLinea)) state.resLinea = lineas.slice().sort((a, b) => avesDe(b) - avesDe(a))[0] || 'Engorde';
    const linea = state.resLinea;
    const filas = base.filter(x => x.Linea === linea && granjaById(x.GranjaID));
    const c = state.consol;
    const ultima = d.ultima;

    let html = `<div class="card filters-card"><div class="filters">
        ${admin && empresas.length > 1 ? `<label class="field"><span>Empresa</span><select data-change="res-empresa"><option value="">Todas las empresas</option>${empresas.map(e =>
            `<option value="${esc(e.EmpresaID)}"${e.EmpresaID === filtroE ? ' selected' : ''}>${esc(e.Nombre)}</option>`).join('')}</select></label>` : ''}
        ${lineas.length ? `<div class="filter-group"><span class="field-label">Línea</span><div class="segmented${lineas.length > 2 ? ' seg-lineas' : ''}" role="group" aria-label="Línea de producción">${lineas.map(l =>
            `<button type="button" class="seg${l === linea ? ' is-active' : ''}" data-action="res-linea" data-linea="${l}" aria-pressed="${l === linea}">${ico(LINEA_ICON[l])}${l}</button>`).join('')}</div></div>` : ''}
    </div></div>`;
    html += `<div class="section-head"><p class="muted">${ultima ? `Última actualización: ${esc(fmtDateTime(ultima.Fin || ultima.Inicio))} (${esc(ultima.Origen)}${ultima.Estado === 'Con errores' ? ', con errores' : ''}).` : 'Aún no se ha actualizado.'}
        ${d.pendientes ? `<strong>${d.pendientes} de ${d.granjas} granjas</strong> tienen datos nuevos.` : 'Todas las granjas están al día.'} Cada madrugada se actualiza sola.</p>
        <div class="row"><button type="button" class="btn btn-primary" data-action="res-actualizar"${c ? ' disabled' : ''}>${ico('rotate-cw')}${c ? 'Actualizando…' : 'Actualizar'}</button>
        ${admin ? `<button type="button" class="btn" data-action="res-actualizar" data-todas="1"${c ? ' disabled' : ''}>Recalcular todo</button>` : ''}</div></div>`;
    if (c) {
        const pct = c.total ? Math.round(c.hechas / c.total * 100) : 100;
        html += `<div class="card"><p class="center"><strong>Actualizando ${nf.format(c.hechas)} de ${nf.format(c.total)} granjas</strong>${c.errores ? ` · ${c.errores} con error` : ''}</p>
            <div class="progreso" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${pct}"><span style="width:${pct}%"></span></div></div>`;
    }
    if (!filas.length) return html + emptyState(d.resumenes.length ? 'No hay resúmenes de esta línea.' : 'Aún no hay resúmenes. Presione "Actualizar" para calcularlos.', 'chart-line');

    // Totales de las últimas semanas.
    const ult = x => x.semanas.slice(-ULTIMAS);
    const tot = indicadores(sumaSemanas(filas.flatMap(ult)));
    const aves = filas.reduce((t, x) => t + (Number(x.ultimo.aves) || 0), 0);
    const lotes = filas.reduce((t, x) => t + (Number(x.ultimo.lotes) || 0), 0);
    const principal = linea === 'Postura' ? ['Postura (%)', tot.postura, 'postura'] : linea === 'Incubadora' ? ['Nacimiento (%)', tot.nacimiento, 'nacimiento'] : ['Peso vs. tabla (%)', tot.pesoVsTabla, 'pesoVsTabla'];
    html += `<div class="kpis">
        ${kpi('Granjas', nf.format(filas.length), `${nf.format(lotes)} ${lotes === 1 ? 'lote abierto' : 'lotes abiertos'}`, 'warehouse')}
        ${kpi('Aves vivas', nf.format(aves), 'en lotes abiertos', 'bird')}
        ${linea === 'Incubadora' ? '' : kpi(`Mortalidad (${ULTIMAS} semanas)`, `${fmt2(tot.mort)} %`, 'sobre las aves promedio', 'activity')}
        ${linea === 'Incubadora' ? '' : kpi('Consumo', `${fmt2(tot.consumo)} g`, 'por ave por día', 'wheat')}
        ${kpi(principal[0], `${fmt2(principal[1])}${principal[2] === 'pesoVsTabla' ? ' %' : ' %'}`, `últimas ${ULTIMAS} semanas`, 'chart-line')}
        ${linea === 'Engorde' || linea === 'Reproductora' ? kpi('Conversión', fmt2(tot.conversion), 'kg alimento / kg ganado', 'scale') : ''}
    </div>`;

    // Tendencia semanal (últimas 12 semanas de las granjas filtradas).
    const claves = [...new Set(filas.flatMap(x => x.semanas.map(w => w.s)))].sort().slice(-12);
    const porSemana = claves.map(k => indicadores(sumaSemanas(filas.flatMap(x => x.semanas.filter(w => w.s === k)))));
    const labels = claves.map(k => periodoLabel(k, 'semana'));
    const graf = (id, titulo, campo, unidad, icono) => chartCard(`res-${linea}-${id}`, {
        type: 'line', title: titulo, subtitle: `${filas.length} granjas · por semana`, icon: icono, unit: unidad, labels,
        series: [{ name: titulo, color: PALETTE.primary, data: porSemana.map(w => (w[campo] === null ? null : Math.round(w[campo] * 100) / 100)) }],
    });
    html += '<div class="grid-2">' + (linea === 'Incubadora'
        ? graf('nac', '% de nacimiento', 'nacimiento', '%', 'egg')
        : graf('mort', 'Mortalidad semanal (%)', 'mort', '%', 'activity') + graf('principal', principal[0], principal[2], '%', 'chart-line')) + '</div>';

    // Ranking por granja.
    const conEmpresa = admin && empresas.length > 1 && !filtroE;
    const rows = filas.map((x) => {
        const i = indicadores(sumaSemanas(ult(x)));
        return { ...x, _aves: Number(x.ultimo.aves) || 0, _lotes: Number(x.ultimo.lotes) || 0, _mort: i.mort, _consumo: i.consumo, _principal: i[principal[2]], _conv: i.conversion };
    });
    const n2 = v => (v === null ? null : Math.round(v * 100) / 100);
    html += table(`res-${linea}`, {
        title: `Granjas · línea de ${linea.toLowerCase()} (últimas ${ULTIMAS} semanas)`, icon: 'warehouse',
        columns: [{ k: 'NombreGranja', l: 'Granja', sticky: true, html: x => `<button type="button" class="btn btn-link" data-action="open-granja" data-id="${esc(x.GranjaID)}">${esc(x.NombreGranja)}</button>` }]
            .concat(conEmpresa ? [{ k: 'EmpresaID', l: 'Empresa', get: x => empresaName(x.EmpresaID) }] : [], [
                { k: '_aves', l: 'Aves vivas', t: 'num' },
                { k: '_lotes', l: 'Lotes abiertos', t: 'num' },
            ], linea === 'Incubadora' ? [] : [
                { k: '_mort', l: 'Mortalidad (%)', t: 'num', get: x => n2(x._mort) },
                { k: '_consumo', l: 'Consumo (g/ave/día)', t: 'num', get: x => n2(x._consumo) },
            ], [{ k: '_principal', l: principal[0], t: 'num', get: x => n2(x._principal) }],
            linea === 'Engorde' || linea === 'Reproductora' ? [{ k: '_conv', l: 'Conversión', t: 'num', get: x => n2(x._conv) }] : [], [
                { k: '_fecha', l: 'Último dato', t: 'date', get: x => x.ultimo.fecha || '' },
                { k: 'ConsolidadoEn', l: 'Actualizado', t: 'datetime' },
            ]),
        rows,
    });
    return html;
}

/** Botón "Actualizar": trae las granjas con cambios por bloques y va llenando la vista. */
async function actualizarResumenes(todas) {
    if (state.consol) return;
    state.consol = { total: 0, hechas: 0, errores: 0 };
    render();
    let corrida = null;
    try {
        const ini = await api('consolidar', { fase: 'iniciar', todas: !!todas });
        corrida = ini.CorridaID;
        state.consol.total = ini.granjas.length;
        render();
        for (let i = 0; i < ini.granjas.length; i += ini.bloque) {
            const r = await api('consolidar', { fase: 'bloque', CorridaID: corrida, GranjaIDs: ini.granjas.slice(i, i + ini.bloque) });
            state.consol.hechas += r.resultados.length;
            state.consol.errores += r.resultados.filter(x => !x.ok).length;
            const cache = state.cache.get('getResumenes{}');
            if (cache?.ok) {
                const hechas = new Set(r.resultados.map(x => x.GranjaID));
                cache.data.resumenes = cache.data.resumenes.filter(x => !hechas.has(x.GranjaID)).concat(r.resumenes);
            }
            render();
        }
        await api('consolidar', { fase: 'fin', CorridaID: corrida });
        toast(state.consol.total ? `Resumen actualizado (${nf.format(state.consol.total)} granjas${state.consol.errores ? `, ${state.consol.errores} con error` : ''}).` : 'Todo estaba al día.', state.consol.errores ? 'error' : 'success');
    } catch (e) {
        toast(e.message, 'error');
        if (corrida) api('consolidar', { fase: 'fin', CorridaID: corrida }).catch(() => {});
    } finally {
        state.consol = null;
        state.cache.delete('getResumenes{}');
        render();
    }
}

// ==========================================================================
// 9e. PANEL: USO Y CONFIGURACIÓN (plataforma)
// ==========================================================================

function viewUso() {
    const r = need('getUso', {});
    if (!r) return loadingBlock(3);
    if (!r.ok) return errorBlock(r);
    const d = r.data;
    const pct = c => (c === null || c === undefined ? null : Math.round(c / d.limite * 1000) / 10);
    const filasTot = f => Object.values(f || {}).reduce((t, n) => t + n, 0);
    let html = `<div class="section-head"><p class="muted">Cada archivo admite hasta ${nf.format(d.limite)} celdas. El uso de cada granja se mide cada vez que se actualiza su resumen (botón "Actualizar" o la tarea de madrugada).</p></div>
        <div class="kpis">
            ${kpi('Archivo central', `${fmt2(pct(d.central.celdas))} %`, `${nf.format(d.central.celdas)} celdas`, 'archive')}
            ${kpi('Granjas', nf.format(d.granjas.length), `${d.granjas.filter(g => g.celdas !== null).length} medidas`, 'warehouse')}
            ${kpi('Granja más grande', `${fmt2(Math.max(0, ...d.granjas.map(g => pct(g.celdas) || 0)))} %`, 'del límite de su archivo', 'activity')}
        </div>`;
    html += table('uso-central', {
        title: 'Archivo central: filas por pestaña', icon: 'archive',
        columns: [{ k: 'tabla', l: 'Pestaña', sticky: true }, { k: 'filas', l: 'Filas', t: 'num' }],
        rows: Object.entries(d.central.filas).map(([tabla, filas]) => ({ tabla, filas })),
    });
    html += table('uso-granjas', {
        title: 'Archivo de cada granja', icon: 'warehouse',
        columns: [{ k: 'NombreGranja', l: 'Granja', sticky: true }, { k: 'EmpresaID', l: 'Empresa', get: g => empresaName(g.EmpresaID) },
            { k: 'celdas', l: 'Celdas', t: 'num' }, { k: '_pct', l: '% del límite', t: 'num', get: g => pct(g.celdas) },
            { k: '_filas', l: 'Filas de datos', t: 'num', get: g => filasTot(g.filas) }, { k: 'en', l: 'Medido', t: 'datetime' }],
        rows: d.granjas,
    });
    html += table('uso-tiempos', {
        title: 'Tiempo de respuesta de hoy (ms)', icon: 'activity', empty: 'Aún no hay solicitudes medidas hoy.',
        columns: [{ k: 'accion', l: 'Acción', sticky: true }, { k: 'n', l: 'Solicitudes', t: 'num' }, { k: 'promedio', l: 'Promedio', t: 'num' }, { k: 'max', l: 'Máximo', t: 'num' }],
        rows: d.tiempos,
    });
    html += table('uso-corridas', {
        title: 'Últimas actualizaciones de resúmenes', icon: 'rotate-cw', empty: 'Aún no hay actualizaciones.',
        columns: [{ k: 'Inicio', l: 'Inicio', t: 'datetime', sticky: true }, { k: 'Origen', l: 'Origen' }, { k: 'Por', l: 'Por' },
            { k: 'Granjas', l: 'Granjas', t: 'num' }, { k: 'Leidas', l: 'Leídas', t: 'num' },
            { k: 'Estado', l: 'Estado', html: c => `<span class="badge ${c.Estado === 'Terminada' ? 'badge-ok' : c.Estado === 'En curso' ? '' : 'badge-warn'}">${esc(c.Estado)}</span>` },
            { k: 'Errores', l: 'Errores', wrap: true, get: c => { try { return JSON.parse(c.Errores || '[]').map(e => `${granjaName(e.GranjaID)}: ${e.error}`).join(' · '); } catch (e) { return ''; } } }],
        rows: d.corridas,
    });
    return html;
}

function viewConfiguracion() {
    const r = need('getConfiguracion', {});
    if (!r) return loadingBlock(1);
    if (!r.ok) return errorBlock(r);
    return `<div class="card"><div class="card-head card-title-center"><h3>${ico('list-checks')}Cobros y tarea de madrugada</h3></div>
        <form class="form-grid" data-submit="configuracion" novalidate>
            ${r.data.valores.map(v => `<label class="field"><span>${esc(v.nombre)}</span><input name="${esc(v.clave)}" type="number" min="${v.min}" max="${v.max}" step="1" inputmode="numeric" value="${v.valor}" required>
                <small class="field-hint">Entre ${v.min} y ${v.max}. Por defecto, ${v.defecto}.${v.alterado ? ' ⚠️ Modificado fuera de la plataforma: se usa el valor por defecto.' : ''}</small></label>`).join('')}
            ${justField()}
            <div class="row row-center span-all"><button type="submit" class="btn btn-primary">${ico('check')}Guardar configuración</button></div>
        </form></div>
        <p class="muted center">Los avisos de cobro se muestran desde los días de aviso antes del vencimiento; después de los días de gracia, la empresa queda en solo lectura. La tarea de madrugada emite los cobros y actualiza los resúmenes.</p>`;
}

// ==========================================================================
// 10. VISTAS: LISTAS, BITÁCORA Y SEGURIDAD
// ==========================================================================

function viewListas() {
    const tipos = state.db.tiposLista.length ? state.db.tiposLista : Object.keys(LISTA_LABEL);
    return `<div class="card"><div class="card-head card-title-center"><h3>${ico('plus')}Agregar valor</h3></div>
        <form class="filters" data-submit="add-lista">
            <label class="field"><span>Lista</span><select name="Tipo" required>${tipos.map(t => `<option value="${esc(t)}">${esc(LISTA_LABEL[t] || t)}</option>`).join('')}</select></label>
            <label class="field"><span>Valor</span><input name="Valor" required maxlength="150" autocomplete="off"></label>
            <button type="submit" class="btn btn-primary">${ico('plus')}Agregar</button>
        </form>
        <p class="muted center">Estos valores aparecen como sugerencias al llenar las fichas; se puede escribir otro valor si hace falta.</p></div>
        <div class="grid-3">${tipos.map((t) => {
        const vals = state.db.listas.filter(l => l.Tipo === t).sort((a, b) => (toNum(a.Orden) ?? 1e9) - (toNum(b.Orden) ?? 1e9) || String(a.Valor).localeCompare(String(b.Valor), 'es'));
        return `<div class="card"><div class="card-head card-title-center"><h3>${ico('list')}${esc(LISTA_LABEL[t] || t)}</h3></div>
                ${vals.length ? `<div class="chips">${vals.map(v => `<span class="chip">${esc(v.Valor)}<button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-lista" data-id="${esc(v.ListaID)}" aria-label="Quitar ${esc(v.Valor)}" title="Quitar">${ico('x')}</button></span>`).join('')}</div>`
                : '<p class="muted center">Sin valores.</p>'}</div>`;
    }).join('')}</div>`;
}

function viewBitacora() {
    if (!varias() && !state.granjaId) return noGranjas();
    const granjaSel = varias()
        ? `<label class="field"><span>Granja</span><select data-change="bit-granja"><option value="">Todas</option>${state.db.granjas.map(g =>
            `<option value="${esc(g.GranjaID)}"${state.bitGranja === g.GranjaID ? ' selected' : ''}>${esc(g.NombreGranja)}</option>`).join('')}</select></label>`
        : granjaSwitch();
    const head = `<div class="card filters-card"><div class="filters">${granjaSel}${dateFilters()}</div></div>`;
    const payload = {
        GranjaID: varias() ? (state.bitGranja || undefined) : state.granjaId,
        desde: state.f.desde || undefined, hasta: state.f.hasta || undefined,
    };
    const r = need('getBitacora', payload);
    if (!r) return head + loadingBlock(1);
    if (!r.ok) return head + errorBlock(r);
    return head + (r.data.truncado ? alertBox('info', 'info', 'Se muestran los cambios más recientes. Reduce el rango de fechas para ver los anteriores.') : '')
        + table('bitacora', {
            title: 'Cambios registrados', icon: 'history', csv: 'bitacora',
            columns: [
                { k: 'Fecha', l: 'Fecha', t: 'datetime', sticky: true },
                { k: 'GranjaID', l: 'Granja', get: b => (b.GranjaID ? granjaName(b.GranjaID) : '') },
                { k: 'Entidad', l: 'Elemento' },
                { k: 'Accion', l: 'Acción' },
                { k: 'Detalle', l: 'Detalle', wrap: true },
                { k: 'Justificacion', l: 'Justificación', wrap: true },
                { k: 'ModificadoPor', l: 'Realizado por' },
            ],
            rows: r.data.bitacora,
        });
}

function viewSeguridad() {
    let html = `<div class="card"><div class="card-head card-title-center"><h2>${ico('shield-check')}Integridad de la base de datos</h2></div>
        <p class="muted center">Comprueba que nadie haya modificado, borrado o insertado datos directamente en la hoja, fuera de la plataforma.</p>
        <div class="row row-center"><button type="button" class="btn btn-primary" data-action="verify-audit">${ico('shield-check')}Verificar integridad</button></div>
        ${state.audit ? auditHtml(state.audit) : ''}</div>`;
    const head = `<div class="card filters-card"><div class="filters">${dateFilters()}</div></div>`;
    const r = need('getAccesos', { desde: state.f.desde || undefined, hasta: state.f.hasta || undefined });
    html += head;
    if (!r) return html + loadingBlock(1);
    if (!r.ok) return html + errorBlock(r);
    return html + table('accesos', {
        title: 'Ingresos e intentos rechazados', icon: 'lock', csv: 'accesos',
        columns: [
            { k: 'Fecha', l: 'Fecha', t: 'datetime', sticky: true },
            { k: 'Email', l: 'Correo' },
            { k: 'UsuarioID', l: 'Usuario', get: a => userById(a.UsuarioID)?.Nombre || (a.UsuarioID ? a.UsuarioID : 'Sin cuenta') },
            { k: 'Accion', l: 'Acción', get: a => ACCION_LABEL[a.Accion] || a.Accion },
            { k: 'Resultado', l: 'Resultado', get: a => RESULTADO_LABEL[a.Resultado] || a.Resultado, html: a => `<span class="badge ${a.Resultado === 'OK' ? 'badge-ok' : 'badge-danger'}">${esc(RESULTADO_LABEL[a.Resultado] || a.Resultado)}</span>` },
        ],
        rows: r.data.accesos,
    });
}

/**
 * Compras de concentrado: registro aparte (no toca lotes, naves ni cálculos). El representante ve
 * su granja; el admin una, varias o todas, con el costo promedio por quintal ponderado.
 */
function viewConcentrado() {
    const ids = ctxIds();
    if (!ids.length) return noGranjas();
    const head = filtersCard({ lineas: false });
    const r = need('getConcentrado', { GranjaIDs: apiGranjas(), desde: state.f.desde || undefined, hasta: state.f.hasta || undefined });
    const fasesTodas = r?.ok ? r.data.porFase.map(f => f.Fase) : [];
    if (state.cFase && !fasesTodas.includes(state.cFase)) state.cFase = '';
    const nueva = `<div class="section-head"><p class="muted">Cada compra guarda su fase de alimento (con el nombre que usa su granja), cuánto se compró, a cuánto el quintal y qué fechas cubre.</p>
        <div class="row">${fasesTodas.length ? `<label class="field field-inline"><span>Fase</span><select data-change="c-fase"><option value="">Todas</option>${fasesTodas.map(f => `<option${f === state.cFase ? ' selected' : ''}>${esc(f)}</option>`).join('')}</select></label>` : ''}
        ${botonImportar('compra', null)}<button type="button" class="btn btn-primary" data-action="new-compra">${ico('plus')}Nueva compra</button></div></div>`;
    if (!r) return head + nueva + loadingBlock(1);
    if (!r.ok) return head + nueva + errorBlock(r);
    const { porFase } = r.data;
    state.fasesPorGranja = r.data.fasesPorGranja || {};
    state.proveedoresAlimento = r.data.proveedores || [];
    const compras = r.data.compras.filter(c => !state.cFase || c.Fase === state.cFase);
    const resumir = (cs) => {
        const q = cs.reduce((a, c) => a + Number(c.CantidadQuintales), 0), costo = cs.reduce((a, c) => a + Number(c.CostoTotal), 0);
        return { quintales: round(q, 4), costo: round(costo, 2), costoPorQuintal: q ? round(costo / q, 2) : null };
    };
    const totales = resumir(compras);
    const porGranja = [...new Set(compras.map(c => c.GranjaID))].map(id => ({ GranjaID: id, ...resumir(compras.filter(c => c.GranjaID === id)) }));
    state.compras = new Map(compras.map(c => [c.CompraID, c]));
    let html = head + nueva + `<div class="kpis">
        ${kpi('Compras', nf.format(compras.length), [state.cFase, state.f.desde || state.f.hasta ? `${fmtDate(state.f.desde) || 'inicio'} a ${fmtDate(state.f.hasta) || 'hoy'}` : 'Todo el historial'].filter(Boolean).join(' · '), 'clipboard-list')}
        ${kpi('Quintales comprados', `${fmtNum(totales.quintales)} qq`, '', 'wheat')}
        ${kpi('Costo total', fmtQ(totales.costo), '', 'file-spreadsheet')}
        ${kpi('Costo promedio', totales.costoPorQuintal !== null ? `${fmtQ(totales.costoPorQuintal)} / qq` : '—', 'Ponderado por cantidad', 'scale')}</div>`;
    // En el tiempo: precio por quintal y quintales de cada compra (las fases elegidas: todas o una).
    if (compras.length) {
        const fases = state.cFase ? [state.cFase] : porFase.map(f => f.Fase);
        const fechas = [...new Set(compras.map(c => c.Desde))].sort();
        const serie = (fase, valor) => fechas.map((d) => {
            const cs = compras.filter(c => c.Fase === fase && c.Desde === d);
            return cs.length ? valor(cs) : null;
        });
        const color = i => PALETTE.cat[i % PALETTE.cat.length];
        html += `<div class="grid-2">${chartCard('concentrado-precio', {
            type: 'line', icon: 'scale', title: `Costo por quintal (Q) en el tiempo${state.cFase ? ` · ${state.cFase}` : ', por fase'}`, labels: fechas.map(fmtDate), unir: true,
            series: fases.map((f, i) => ({ name: f, color: color(i), data: serie(f, cs => resumir(cs).costoPorQuintal) })),
        })}${chartCard('concentrado-qq', {
            type: 'bar', icon: 'wheat', title: `Quintales comprados${state.cFase ? ` · ${state.cFase}` : ', por fase'}`, labels: fechas.map(fmtDate), unit: 'qq',
            series: fases.map((f, i) => ({ name: f, color: color(i), data: serie(f, cs => resumir(cs).quintales) })),
        })}</div>`;
    }
    // Por empresa de alimento (de la fase elegida o de todas).
    const conProveedor = compras.filter(c => c.Proveedor);
    if (conProveedor.length) {
        const provs = [...new Set(conProveedor.map(c => c.Proveedor))].sort((a, b) => a.localeCompare(b, 'es'));
        html += table('concentrado-proveedores', {
            title: `Por empresa de alimento${state.cFase ? ` · ${state.cFase}` : ''}`, icon: 'building', csv: 'concentrado-por-proveedor',
            columns: [{ k: 'Proveedor', l: 'Empresa de alimento', sticky: true }, { k: 'n', l: 'Compras', t: 'num' }, { k: 'quintales', l: 'Quintales', t: 'num' },
                { k: 'costo', l: 'Costo total (Q)', t: 'num' }, { k: 'costoPorQuintal', l: 'Costo promedio por quintal (Q)', t: 'num' }],
            rows: provs.map(p => { const cs = conProveedor.filter(c => c.Proveedor === p); return { Proveedor: p, n: cs.length, ...resumir(cs) }; }),
        });
    }
    // Para estudiar el alimento por fase: cuánto se compró de cada una y a qué costo promedio.
    if (!state.cFase && porFase.length) {
        html += `<div class="grid-2">${chartCard('concentrado-fase-qq', {
            type: 'bar', icon: 'wheat', title: 'Quintales por fase de alimento', labels: porFase.map(f => f.Fase), unit: 'qq',
            series: [{ name: 'Quintales', color: PALETTE.primary, data: porFase.map(f => Number(f.quintales)) }],
        })}${chartCard('concentrado-fase-q', {
            type: 'bar', icon: 'scale', title: 'Costo promedio por quintal (Q) por fase', labels: porFase.map(f => f.Fase),
            series: [{ name: 'Q por quintal (ponderado)', color: PALETTE.tabla, data: porFase.map(f => (f.costoPorQuintal === '' ? null : Number(f.costoPorQuintal))) }],
        })}</div>`;
        html += table('concentrado-fases', {
            title: 'Por fase de alimento', icon: 'wheat', csv: 'concentrado-por-fase',
            columns: [{ k: 'Fase', l: 'Fase', sticky: true }, { k: 'compras', l: 'Compras', t: 'num' }, { k: 'quintales', l: 'Quintales', t: 'num' },
                { k: 'costo', l: 'Costo total (Q)', t: 'num' }, { k: 'costoPorQuintal', l: 'Costo promedio por quintal (Q)', t: 'num' }],
            rows: porFase,
        });
    }
    if (varias() && porGranja.length > 1) {
        html += table('concentrado-granjas', {
            title: 'Por granja', icon: 'warehouse', csv: 'concentrado-por-granja',
            columns: [{ k: 'GranjaID', l: 'Granja', sticky: true, get: g => granjaName(g.GranjaID) }, { k: 'quintales', l: 'Quintales', t: 'num' },
                { k: 'costo', l: 'Costo total (Q)', t: 'num' }, { k: 'costoPorQuintal', l: 'Costo por quintal (Q)', t: 'num' }],
            rows: porGranja,
        });
    }
    return html + table('compras', {
        title: 'Compras de concentrado', icon: 'wheat', csv: 'compras-concentrado',
        empty: 'No hay compras en el rango seleccionado.',
        columns: [{ k: 'Desde', l: 'Desde', t: 'date', sticky: true }, { k: 'Hasta', l: 'Hasta', t: 'date' }].concat(colsGranja(), [
            { k: 'Fase', l: 'Fase de alimento' }, { k: 'Proveedor', l: 'Empresa de alimento' }, { k: 'CantidadQuintales', l: 'Cantidad (qq)', t: 'num' },
            { k: 'CostoPorQuintal', l: 'Costo por quintal (Q)', t: 'num' }, { k: 'CostoTotal', l: 'Costo total (Q)', t: 'num' }, { k: 'Comentarios', l: 'Comentarios', wrap: true },
            { k: 'CreadoPor', l: 'Registrado por' }, { k: 'CreadoEn', l: 'Registrado el', t: 'datetime' }]),
        rows: compras,
        actions: c => (!modificable(c.CreadoEn) ? sinPermisoFecha()
            : `<button type="button" class="icon-btn icon-btn-sm" data-action="edit-compra" data-id="${esc(c.CompraID)}" title="Editar" aria-label="Editar compra">${ico('pencil')}</button>
               <button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-compra" data-id="${esc(c.CompraID)}" title="Eliminar" aria-label="Eliminar compra">${ico('trash')}</button>`),
    });
}

/**
 * Serología: muestreos por lote (un título por ave) con su resumen, gráficos por edad y los
 * métodos o kits. Los gráficos solo comparan una enfermedad con un método (los kits no se mezclan).
 */
function viewSerologia() {
    const ids = ctxIds();
    if (!ids.length) return noGranjas();
    // Enfermedades con casillas: una = su detalle por método; varias = se comparan en el gráfico.
    state.srEnfs = state.srEnfs.filter(id => enfermedadById(id));
    state.srEnf = state.srEnfs.length === 1 ? state.srEnfs[0] : '';
    const metodos = state.srEnf ? metodosDe(state.srEnf, ids) : [];
    if (state.srMetodo && !metodos.some(m => m.MetodoID === state.srMetodo)) state.srMetodo = '';
    const metodo = metodoById(state.srMetodo);
    const etiquetaEnf = !state.srEnfs.length ? 'Todas' : state.srEnfs.length === 1 ? enfermedadById(state.srEnf).Nombre : `${state.srEnfs.length} enfermedades`;
    const selectores = `<div class="filter-group"><span class="field-label">Enfermedades</span>
        <details class="dropdown" data-dd="sr"${state.srDd ? ' open' : ''}><summary>${ico('filter', 'icon-sm')}<span>${esc(etiquetaEnf)}</span>${ico('chevron-down', 'icon-sm')}</summary>
            <div class="dropdown-panel"><label class="check"><input type="checkbox" data-change="sr-f-enf-todas"${state.srEnfs.length ? '' : ' checked'}> Todas</label>
            ${state.db.enfermedades.map(e => `<label class="check"><input type="checkbox" data-change="sr-f-enf" value="${esc(e.EnfermedadID)}"${state.srEnfs.includes(e.EnfermedadID) ? ' checked' : ''}> ${esc(e.Nombre)}</label>`).join('')}</div></details></div>
        <label class="field"><span>Método</span><select data-change="sr-f-metodo"${state.srEnf ? '' : ' disabled'}><option value="">${state.srEnf ? 'Todos' : 'Marque una sola enfermedad'}</option>${metodos.map(m =>
        `<option value="${esc(m.MetodoID)}"${m.MetodoID === state.srMetodo ? ' selected' : ''}>${esc(nombreMetodo(m))}</option>`).join('')}</select></label>`;
    const head = filtersCard({ eliminados: true, extra: selectores });
    if (!lineasDisponibles(ids).length) return head + sinLineas();
    const linea = state.linea;
    if (!incluye('serologia', ids, linea)) return head + sinPlan('serologia');
    const aviso = sinEstructura(linea);
    const nuevo = aviso || `<div class="section-head"><p class="muted">Se elige la nave, la fecha del muestreo, la enfermedad y el método, y se escribe el resultado de cada ave. La plataforma calcula el resumen.</p>
        <div class="row">${botonImportar('serologia')}<button type="button" class="btn btn-primary" data-action="new-serologia">${ico('syringe')}Nuevo muestreo</button></div></div>`;
    const r = need('getSerologias', registrosPayload({
        EnfermedadID: state.srEnf || undefined, MetodoID: state.srMetodo || undefined,
        incluirEliminados: isAdmin() && state.f.eliminados ? true : undefined,
    }));
    const cuerpo = (() => {
        if (!r) return loadingBlock(1);
        if (!r.ok) return errorBlock(r);
        const filas = state.srEnfs.length > 1 ? r.data.serologias.filter(x => state.srEnfs.includes(x.EnfermedadID)) : r.data.serologias;
        state.serologias = new Map(filas.map(x => [x.SerologiaID, x]));
        const graficos = metodo ? seroCharts(filas, metodo)
            : state.srEnfs.length > 1 ? seroComparar(filas, state.srEnfs) + seroPanorama(filas)
            : seroPanorama(filas) + (state.srEnf ? '' : seroComparar(filas));
        return graficos + table('serologias', {
            title: `Muestreos de ${linea.toLowerCase()}`, icon: 'syringe', csv: `serologia-${linea}`,
            empty: 'No hay muestreos en el rango seleccionado.', rows: filas,
            columns: [{ k: 'Fecha', l: 'Fecha', t: 'date', sticky: true }].concat(colsGranja(), COLS_UBICACION, [
                { k: 'EdadSemanas', l: 'Edad (semanas)', t: 'num' }, { k: 'EdadDias', l: 'Edad (días)', t: 'num' },
                { k: 'Enfermedad', l: 'Enfermedad' }, { k: 'Metodo', l: 'Método' }, { k: 'Prueba', l: 'Prueba' },
                { k: 'Muestra', l: 'Muestra', t: 'num' },
                { k: 'Resumen', l: 'Resumen', t: 'num', get: x => toNum(valorResumen(x)) },
                { k: 'Maximo', l: 'Dato mayor', t: 'num' }, { k: 'CV', l: 'CV (%)', t: 'num' },
                { k: 'PctPositivos', l: 'Positivas (%)', t: 'num' }, { k: 'PctProtegidas', l: 'Protegidas (%)', t: 'num' },
                { k: 'EdadVacunacion', l: 'Vacunar el día (Deventer)', t: 'num' },
            ], linea === 'Incubadora' ? [{ k: 'LoteOrigen', l: 'Lote de origen' }] : [], [
                { k: 'Laboratorio', l: 'Laboratorio' }, { k: 'Valores', l: 'Resultados', nocsv: false, wrap: true, html: x => `<span class="muted">${esc(String(x.Valores || '').split(',').filter(Boolean).length)} resultados</span>` },
                { k: 'CreadoPor', l: 'Registrado por' }, { k: 'CreadoEn', l: 'Registrado el', t: 'datetime' },
            ], isAdmin() && state.f.eliminados ? [{ k: 'Estado', l: 'Estado', html: x => badgeEstado(x.Estado === 'Activo' ? 'Activo' : 'Eliminado') }] : []),
            rowClass: x => (x.Estado !== 'Activo' ? 'is-deleted' : ''),
            actions: x => `<button type="button" class="icon-btn icon-btn-sm" data-action="ver-serologia" data-id="${esc(x.SerologiaID)}" title="Ver los resultados" aria-label="Ver los resultados">${ico('list-checks')}</button>`
                + (x.Estado === 'Activo' && !modificable(x.CreadoEn) ? sinPermisoFecha() : '')
                + (x.Estado === 'Activo' && modificable(x.CreadoEn) ? `<button type="button" class="icon-btn icon-btn-sm" data-action="edit-serologia" data-id="${esc(x.SerologiaID)}" title="Editar" aria-label="Editar muestreo">${ico('pencil')}</button>
                   <button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-serologia" data-id="${esc(x.SerologiaID)}" title="Eliminar" aria-label="Eliminar muestreo">${ico('trash')}</button>` : ''),
        });
    })();
    return head + nuevo + cuerpo + metodosTabla(ids);
}

/** Métodos o kits visibles en el contexto: generales y los propios de las granjas elegidas. */
function metodosTabla(ids) {
    const filas = state.db.metodosSerologia.filter(m => !m.GranjaID || ids.includes(m.GranjaID))
        .filter(m => !state.srEnfs.length || state.srEnfs.includes(m.EnfermedadID));
    const puedeEditar = m => isAdmin() || !!m.GranjaID;
    return `<div class="section-head"><p class="muted">Cada granja usa su propio método o kit. El corte de positivo y el nivel protector se anotan solo si están publicados (inserto del kit o documento del laboratorio), con su fuente.</p>
        <button type="button" class="btn" data-action="new-metodo">${ico('plus')}Nuevo método</button></div>` + table('metodos-serologia', {
        title: 'Métodos o kits', icon: 'list-checks', csv: 'metodos-serologia',
        empty: 'No hay métodos para la enfermedad elegida.', rows: filas,
        columns: [
            { k: 'EnfermedadID', l: 'Enfermedad', sticky: true, get: m => enfermedadById(m.EnfermedadID)?.Nombre || '—' },
            { k: 'Nombre', l: 'Método' }, { k: 'Prueba', l: 'Prueba' }, { k: 'Unidad', l: 'Unidad' },
            { k: 'GranjaID', l: 'De', get: m => (m.GranjaID ? granjaName(m.GranjaID) : 'General (todas)') },
            { k: 'Laboratorio', l: 'Laboratorio' }, { k: 'CortePositivo', l: 'Positivo desde', t: 'num' },
            { k: 'NivelProtector', l: 'Nivel protector', t: 'num' }, { k: 'TituloQuiebre', l: 'Quiebre de la vacuna (Deventer)', t: 'num' }, { k: 'Fuente', l: 'Fuente', wrap: true },
        ],
        actions: m => (puedeEditar(m)
            ? `<button type="button" class="icon-btn icon-btn-sm" data-action="edit-metodo" data-id="${esc(m.MetodoID)}" title="Editar" aria-label="Editar método">${ico('pencil')}</button>
               <button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-metodo" data-id="${esc(m.MetodoID)}" title="Eliminar" aria-label="Eliminar método">${ico('trash')}</button>`
            : `<span class="icon-btn icon-btn-sm muted" title="Método general: lo modifica la administración" aria-label="Método general: lo modifica la administración">${ico('lock')}</span>`),
    });
}

/**
 * Comparativo de lotes: se eligen hasta 6 lotes de una línea (abiertos o cerrados, de cualquier
 * nave; el admin, de cualquier granja) y se comparan por semana de edad, en gráficos y tablas.
 */
const MAX_COMPARAR = 6;
const CMP_INDICADORES = {
    zoo: [['peso', 'Peso corporal', 'PesoCorporalReal', 'scale'], ['mort', 'Mortalidad acumulada (%)', 'MortalidadAcumuladaPorcentaje', 'activity'],
        ['consumo', 'Consumo (g/ave/día)', 'ConsumoGramosAveDia', 'wheat'], ['conv', 'Conversión alimenticia', 'ConversionAlimenticia', 'wheat']],
    Postura: [['postura', 'Postura (%)', 'PosturaPorcentaje', 'egg']],
    Incubadora: [['nac', '% de nacimiento', 'PorcentajeNacimiento', 'egg'], ['pesohuevo', 'Peso de huevo', 'PesoHuevo', 'scale']],
};
function viewComparativo() {
    const ids = ctxIds();
    if (!ids.length) return noGranjas();
    const head = filtersCard({ fechas: false, ubicacion: false, unidad: true });
    if (!lineasDisponibles(ids).length) return head + sinLineas();
    const linea = state.linea;
    if (!incluye('comparativo', ids, linea)) return head + sinPlan('comparativo');
    const c = state.cmp;
    // Solo lotes de granjas cuyo plan de esta línea incluye el comparativo.
    const todos = state.db.lotes.filter(l => l.Linea === linea && ids.includes(l.GranjaID) && l.Estado !== 'Eliminado' && incluye('comparativo', [l.GranjaID], linea));
    c.sel = c.sel.filter(id => todos.some(l => l.LoteID === id));
    const geneticas = [...new Set(todos.map(l => l.LineaGenetica).filter(Boolean))].sort();
    const q = norm(c.q);
    const visibles = todos.filter(l => (!c.estado || (c.estado === 'abiertos' ? l.Estado === 'Activo' : l.Estado !== 'Activo'))
        && (!c.gen || l.LineaGenetica === c.gen)
        && (!q || norm([l.NumeroLote, granjaName(l.GranjaID), sectorById(l.SectorID)?.Nombre, naveById(l.NaveID)?.Nombre].join(' ')).includes(q)))
        .sort((a, b) => String(b.FechaInicio).localeCompare(String(a.FechaInicio)));
    const nombreLote = l => `${l.NumeroLote} · ${naveById(l.NaveID)?.Nombre || ''}${varias() && ids.length > 1 ? ` (${granjaName(l.GranjaID)})` : ''}`;
    const lista = `<div class="card"><div class="card-head card-title-center"><h3>${ico('list-checks')}Elija los lotes a comparar</h3></div>
        <p class="muted center">Abiertos y cerrados, de cualquier nave${varias() ? ' y granja' : ''}. Hasta ${MAX_COMPARAR} a la vez; se comparan por semana de edad.</p>
        <div class="filters">
            <label class="field"><span>Estado</span><select data-change="cmp-estado"><option value="">Todos</option><option value="abiertos"${c.estado === 'abiertos' ? ' selected' : ''}>Abiertos</option><option value="cerrados"${c.estado === 'cerrados' ? ' selected' : ''}>Cerrados</option></select></label>
            <label class="field"><span>Línea genética</span><select data-change="cmp-gen"><option value="">Todas</option>${geneticas.map(g => `<option${g === c.gen ? ' selected' : ''}>${esc(g)}</option>`).join('')}</select></label>
            <label class="field"><span>Buscar</span><input type="search" value="${esc(c.q)}" data-change="cmp-q" placeholder="Lote, nave o granja"></label>
        </div>
        ${visibles.length ? `<div class="table-wrap import-tabla"><table><thead><tr><th></th><th>Lote</th>${varias() ? '<th>Granja</th>' : ''}<th>Sector / nave</th><th>Línea genética</th><th>Inicio</th><th>Estado</th><th>Aves iniciales</th></tr></thead><tbody>
        ${visibles.map(l => `<tr><td><input type="checkbox" data-change="cmp-lote" value="${esc(l.LoteID)}" aria-label="Comparar el lote ${esc(l.NumeroLote)}"${c.sel.includes(l.LoteID) ? ' checked' : ''}${!c.sel.includes(l.LoteID) && c.sel.length >= MAX_COMPARAR ? ' disabled' : ''}></td>
            <td><strong>${esc(l.NumeroLote)}</strong></td>${varias() ? `<td>${esc(granjaName(l.GranjaID))}</td>` : ''}<td>${esc([sectorById(l.SectorID)?.Nombre, naveById(l.NaveID)?.Nombre].filter(Boolean).join(' / '))}</td>
            <td>${esc(l.LineaGenetica || '—')}</td><td>${esc(fmtDate(l.FechaInicio))}</td><td>${l.Estado === 'Activo' ? '<span class="badge badge-ok">Abierto</span>' : `<span class="badge badge-muted">Cerrado ${esc(fmtDate(l.FechaCierre))}</span>`}</td>
            <td>${l.AvesIniciales ? esc(fmtNum(l.AvesIniciales)) : '—'}</td></tr>`).join('')}</tbody></table></div>`
        : `<p class="muted center">No hay lotes de ${linea.toLowerCase()} con esos filtros.</p>`}
        ${c.sel.length ? `<div class="row row-center"><span class="muted">${c.sel.length} de ${MAX_COMPARAR} elegidos</span><button type="button" class="btn btn-sm" data-action="cmp-limpiar">${ico('x')}Quitar todos</button></div>` : ''}</div>`;
    if (!c.sel.length) return head + lista;
    const r = need('getRegistros', { Linea: linea, GranjaIDs: apiGranjas(), LoteIDs: c.sel });
    if (!r) return head + lista + loadingBlock(2);
    if (!r.ok) return head + lista + errorBlock(r);
    const regs = r.data.registros;
    const lotes = c.sel.map(loteById).filter(Boolean);
    const sinEdad = lotes.filter(l => edadDias(l, l.FechaInicio) === null);
    const conEdad = regs.filter(x => edadDias(loteById(x.LoteID), x.Fecha) !== null);
    const semanas = [...new Set(conEdad.map(x => Math.floor(edadDias(loteById(x.LoteID), x.Fecha) / 7)))].sort((a, b) => a - b);
    const grupos = new Map(lotes.map(l => [l.LoteID, new Map(aggregate(conEdad.filter(x => x.LoteID === l.LoteID), camposNumericos(linea),
        x => Math.floor(edadDias(l, x.Fecha) / 7)).map(g => [g.key, g]))]));
    const color = i => PALETTE.cat[i % PALETTE.cat.length];
    const indicadores = (linea === 'Incubadora' ? [] : CMP_INDICADORES.zoo).concat(CMP_INDICADORES[linea] || []);
    const misma = new Set(lotes.map(l => l.LineaGenetica)).size === 1;
    let html = head + lista;
    if (sinEdad.length) html += alertBox('warn', 'alert-triangle', `Sin edad al iniciar, no se pueden ubicar por semana: ${esc(sinEdad.map(l => l.NumeroLote).join(', '))}. Complete la edad en "Editar lote".`);
    html += `<div class="grid-2">${indicadores.map(([id, titulo, k, icon]) => {
        const campo = campoDe(linea, k) || { k, peso: k.startsWith('Peso') };
        const series = lotes.map((l, i) => ({ name: nombreLote(l), color: color(i),
            data: semanas.map(s => { const g = grupos.get(l.LoteID).get(s); const v = g ? gAvg(g, k) : null; return v === null ? null : round(valorCampo(campo, v), campo.peso ? decimalesPeso() : 2); }) }));
        // La tabla de la línea genética, si todos los lotes son de la misma.
        const tablaK = { PesoCorporalReal: 'PesoCorporalTabla', MortalidadAcumuladaPorcentaje: 'MortalidadIdeal', ConsumoGramosAveDia: 'ConsumoTabla', PosturaPorcentaje: 'PosturaTabla' }[k];
        if (misma && tablaK) {
            const g0 = grupos.get(lotes[0].LoteID);
            const datos = semanas.map(s => { const g = g0.get(s); const v = g ? gAvg(g, tablaK) : null; return v === null ? null : round(valorCampo(campo, v), campo.peso ? decimalesPeso() : 2); });
            if (datos.some(v => v !== null)) series.push({ name: `Tabla ${lotes[0].LineaGenetica}`, color: PALETTE.tabla, dashed: true, data: datos });
        }
        return chartCard(`cmp-${id}`, { type: 'line', icon, title: campo.peso ? `${titulo} (${state.unidad})` : titulo, labels: semanas.map(s => `Semana ${s}`), unir: true, series });
    }).join('')}</div>`;
    // Resumen por lote.
    const resumen = lotes.map((l) => {
        const propios = regs.filter(x => x.LoteID === l.LoteID).sort((a, b) => String(a.Fecha).localeCompare(String(b.Fecha)));
        const ult = propios[propios.length - 1];
        const tot = aggregate(propios, camposNumericos(linea), () => 't')[0];
        const pesado = [...propios].reverse().find(x => toNum(x.PesoCorporalReal ?? x.PesoHuevo) !== null);
        const fin = l.FechaCierre || ult?.Fecha;
        const avesFin = ult ? Math.max(0, (toNum(ult.NumeroAves) || 0) - (toNum(ult.MortalidadNominal) || 0) - (toNum(ult.OtrasSalidas) || 0)) : null;
        return {
            Lote: l.NumeroLote, Granja: granjaName(l.GranjaID), Nave: [sectorById(l.SectorID)?.Nombre, naveById(l.NaveID)?.Nombre].filter(Boolean).join(' / '),
            LineaGenetica: l.LineaGenetica, Estado: l.Estado === 'Activo' ? 'Abierto' : 'Cerrado', FechaInicio: l.FechaInicio, FechaFin: fin || '',
            Dias: fin && l.FechaInicio ? Math.round((new Date(String(fin).slice(0, 10)) - new Date(l.FechaInicio)) / 86400000) : null,
            AvesIniciales: toNum(l.AvesIniciales), AvesFinales: linea === 'Incubadora' ? null : avesFin,
            MortAcum: ult ? toNum(ult.MortalidadAcumuladaPorcentaje) : null,
            PesoFinal: pesado ? round(enUnidad(pesado.PesoCorporalReal ?? pesado.PesoHuevo), decimalesPeso()) : null,
            Conversion: tot ? round(gAvg(tot, 'ConversionAlimenticia'), 3) : null,
            Consumo: tot ? round(gSum(tot, 'ConsumoQuintalesTotal'), 2) : null,
            CostoPorAve: toNum(l.CostoPorAve),
        };
    });
    html += table('cmp-resumen', {
        title: 'Resumen por lote', icon: 'table', csv: `comparativo-lotes-${linea}`, rows: resumen,
        columns: [{ k: 'Lote', l: 'Lote', sticky: true }].concat(varias() ? [{ k: 'Granja', l: 'Granja' }] : [], [{ k: 'Nave', l: 'Sector / nave' },
            { k: 'LineaGenetica', l: 'Línea genética' }, { k: 'Estado', l: 'Estado' }, { k: 'FechaInicio', l: 'Inicio', t: 'date' }, { k: 'FechaFin', l: 'Cierre o último dato', t: 'date' },
            { k: 'Dias', l: 'Días', t: 'num' }, { k: 'AvesIniciales', l: 'Aves iniciales', t: 'num' }],
            linea === 'Incubadora' ? [] : [{ k: 'AvesFinales', l: 'Aves al final', t: 'num' }, { k: 'MortAcum', l: 'Mortalidad acumulada (%)', t: 'num' },
                { k: 'Consumo', l: 'Alimento (qq)', t: 'num' }, { k: 'Conversion', l: 'Conversión', t: 'num' }],
            [{ k: 'PesoFinal', l: `${linea === 'Incubadora' ? 'Peso de huevo' : 'Peso'} del último pesaje (${state.unidad})`, t: 'num' }, { k: 'CostoPorAve', l: 'Costo por ave (Q)', t: 'num' }]),
    });
    // Tabla semanal (una fila por lote y semana).
    const semanal = [];
    lotes.forEach(l => semanas.forEach((s) => {
        const g = grupos.get(l.LoteID).get(s);
        if (!g) return;
        const fila = { Lote: l.NumeroLote, Semana: s };
        indicadores.forEach(([, , k]) => { const campo = campoDe(linea, k) || { k, peso: k.startsWith('Peso') }; const v = gAvg(g, k); fila[k] = v === null ? null : round(valorCampo(campo, v), campo.peso ? decimalesPeso() : 2); });
        semanal.push(fila);
    }));
    html += table('cmp-semanal', {
        title: 'Por semana de edad', icon: 'calendar-range', csv: `comparativo-semanal-${linea}`, rows: semanal,
        columns: [{ k: 'Lote', l: 'Lote', sticky: true }, { k: 'Semana', l: 'Semana de edad', t: 'num' }]
            .concat(indicadores.map(([, titulo, k]) => ({ k, l: (campoDe(linea, k)?.peso || k.startsWith('Peso')) ? `${titulo} (${state.unidad})` : titulo, t: 'num' }))),
    });
    return html;
}

/**
 * Panorama: un gráfico por enfermedad (y por método, porque los kits no se mezclan) con todos los
 * muestreos del rango, por edad. Al elegir un método se ven sus gráficos detallados.
 */
function seroPanorama(filas) {
    const activos = filas.filter(x => x.Estado === 'Activo' && x.EdadDias !== '');
    const porMetodo = new Map();
    activos.forEach((x) => { if (!porMetodo.has(x.MetodoID)) porMetodo.set(x.MetodoID, []); porMetodo.get(x.MetodoID).push(x); });
    const metodos = [...porMetodo.keys()].map(metodoById).filter(Boolean)
        .sort((a, b) => String(enfermedadById(a.EnfermedadID)?.Nombre).localeCompare(String(enfermedadById(b.EnfermedadID)?.Nombre), 'es') || byNombre(a, b));
    if (!metodos.length) {
        return `<div class="alert alert-info">${ico('info')}<span>${filas.some(x => x.Estado === 'Activo')
            ? 'Los muestreos de lotes sin la edad de las aves al iniciar no se pueden ubicar por edad. Completa la edad en "Editar lote".'
            : 'No hay muestreos en el rango seleccionado.'}</span></div>`;
    }
    return `<div class="section-head"><p class="muted">Un gráfico por enfermedad y método, por edad de las aves. Elige un método en los filtros para ver su detalle (CV y % de positivas).</p></div>
        <div class="grid-2">${metodos.map(m => seroResumenCard(porMetodo.get(m.MetodoID), m, `sero-pan-${m.MetodoID}`, true)).join('')}</div>`;
}

/**
 * Comparación de enfermedades con la misma prueba. HI: promedio en log2 (misma escala). ELISA:
 * cada muestreo como % del nivel protector de su kit (o de su corte, si no hay nivel), porque los
 * títulos de kits distintos no se comparan en número.
 */
const COMPARAR_SUGERIDAS = { HI: ['IA', 'H5', 'NC'], ELISA: ['IBD', 'IBV'] };
/** Edad en el eje de los gráficos de serología: semanas con un decimal ("Sem. 5.3"). */
const etiquetaSemana = d => `Sem. ${nf.format(Math.round(d / 7 * 10) / 10)}`;
function seroComparar(filas, marcadas = null) {
    const activos = filas.filter(x => x.Estado === 'Activo' && x.EdadDias !== '' && x.Prueba !== 'PLACA');
    const conDatos = p => [...new Set(activos.filter(x => x.Prueba === p).map(x => x.EnfermedadID))].map(enfermedadById).filter(Boolean);
    const pruebas = ['HI', 'ELISA'].filter(p => conDatos(p).length >= 2);
    if (!pruebas.length) return '';
    const c = state.srComp;
    if (!pruebas.includes(c.prueba)) { c.prueba = pruebas[0]; c.enfs = null; }
    const disponibles = conDatos(c.prueba).sort(byNombre);
    // Con varias enfermedades marcadas en el filtro, se comparan esas (las de esta prueba).
    if (marcadas) c.enfs = marcadas.filter(id => disponibles.some(e => e.EnfermedadID === id));
    if (!c.enfs || !c.enfs.some(id => disponibles.some(e => e.EnfermedadID === id))) {
        const sug = disponibles.filter(e => COMPARAR_SUGERIDAS[c.prueba].includes(e.Codigo)).map(e => e.EnfermedadID);
        c.enfs = sug.length >= 2 ? sug : disponibles.slice(0, 2).map(e => e.EnfermedadID);
    }
    const elegidas = disponibles.filter(e => c.enfs.includes(e.EnfermedadID));
    const ref = (x) => {
        const m = metodoById(x.MetodoID);
        const r = toNum(m?.NivelProtector) || toNum(m?.CortePositivo);
        return r ? toNum(x.MediaGeometrica) / r * 100 : null;
    };
    const valor = x => (c.prueba === 'HI' ? toNum(x.Promedio) : ref(x));
    const usados = activos.filter(x => x.Prueba === c.prueba && c.enfs.includes(x.EnfermedadID) && valor(x) !== null);
    const dias = [...new Set(usados.map(x => Number(x.EdadDias)))].sort((a, b) => a - b);
    const series = elegidas.map((e, i) => {
        const porDia = new Map();
        usados.filter(x => x.EnfermedadID === e.EnfermedadID).forEach((x) => { const d = Number(x.EdadDias); if (!porDia.has(d)) porDia.set(d, []); porDia.get(d).push(valor(x)); });
        return { name: e.Nombre, color: PALETTE.cat[i % PALETTE.cat.length], data: dias.map(d => (porDia.has(d) ? round(porDia.get(d).reduce((a, v) => a + v, 0) / porDia.get(d).length, 2) : null)) };
    });
    if (c.prueba === 'ELISA') series.push({ name: '100 % = nivel protector (o corte) del kit', color: PALETTE.tabla, dashed: true, data: dias.map(() => 100) });
    const controles = `<div class="row row-center">
        <div class="segmented" role="group" aria-label="Prueba">${pruebas.map(p => `<button type="button" class="seg${p === c.prueba ? ' is-active' : ''}" data-action="sr-comp-prueba" data-prueba="${p}" aria-pressed="${p === c.prueba}">${p}</button>`).join('')}</div>
        ${marcadas ? '' : `<div class="chips">${disponibles.map(e => `<label class="chip"><input type="checkbox" value="${esc(e.EnfermedadID)}" data-change="sr-comp-enf"${c.enfs.includes(e.EnfermedadID) ? ' checked' : ''}> ${esc(e.Nombre)}</label>`).join('')}</div>`}</div>`;
    const nota = c.prueba === 'HI'
        ? 'Promedio de los títulos HI (log2) de todas las naves, por edad. Misma escala para todas las enfermedades.'
        : 'Cada muestreo como % del nivel protector de su kit (o de su corte, si el kit no publica nivel): así se comparan kits distintos. Los muestreos de kits sin corte ni nivel no se incluyen.';
    return `<div class="card"><div class="card-head card-title-center"><h3>${ico('chart-line')}Comparar enfermedades con la misma prueba</h3></div>
        ${controles}<p class="muted center">${esc(nota)}</p></div>
        ${elegidas.length >= 2 && dias.length ? chartCard('sero-comparar', {
        type: 'line', icon: 'syringe', title: c.prueba === 'HI' ? `HI: ${elegidas.map(e => e.Nombre).join(' vs ')} (log2)` : `ELISA: ${elegidas.map(e => e.Nombre).join(' vs ')} (% del nivel del kit)`,
        labels: dias.map(etiquetaSemana), unir: true, unit: c.prueba === 'HI' ? '' : '%', series,
    }) : `<div class="alert alert-info">${ico('info')}<span>Elige al menos dos enfermedades con muestreos en el rango.</span></div>`}`;
}

/** Valor que resume el muestreo según su prueba (HI: promedio log2; ELISA: media geométrica; placa: % positivas). */
function valorResumen(x) {
    return x.Prueba === 'HI' ? x.Promedio : x.Prueba === 'ELISA' ? x.MediaGeometrica : x.PctPositivos;
}

/**
 * Gráficos por edad (días) de una enfermedad y un método: una serie por granja (vista global con
 * varias granjas) o por nave. Varios muestreos de la misma serie y edad se combinan (ELISA con
 * media geométrica; el resto, promedio). El corte y el nivel protector van como líneas punteadas.
 */
function seroCharts(filas, metodo) {
    const activos = filas.filter(x => x.Estado === 'Activo' && x.EdadDias !== '');
    if (!activos.length) {
        return `<div class="alert alert-info">${ico('info')}<span>${filas.some(x => x.Estado === 'Activo')
            ? 'Los muestreos de lotes sin la edad de las aves al iniciar no se pueden ubicar por edad. Completa la edad en "Editar lote".'
            : 'No hay muestreos de este método en el rango seleccionado.'}</span></div>`;
    }
    const enf = enfermedadById(metodo.EnfermedadID)?.Nombre || '';
    const { labels, series } = seroSeries(activos);
    let html = seroResumenCard(activos, metodo, 'sero-resumen', false);
    if (metodo.Prueba === 'ELISA') {
        html += chartCard('sero-cv', { type: 'line', icon: 'activity', title: `${enf} · CV (%) por edad`, labels, unir: true, unit: '%', series: series('CV', false) });
    }
    if (metodo.Prueba !== 'PLACA' && metodo.CortePositivo !== '') {
        html += chartCard('sero-positivas', { type: 'line', icon: 'chart-line', title: `${enf} · Aves positivas (%) por edad`, labels, unir: true, unit: '%', series: series('PctPositivos', false) });
    }
    return `<div class="grid-2">${html}</div>`;
}

/**
 * Series por edad (días): una por granja (vista global con varias granjas) o por nave. Varios
 * muestreos de la misma serie y edad se combinan (ELISA con media geométrica; el resto, promedio).
 */
function seroSeries(activos) {
    const porGranja = varias() && new Set(activos.map(x => x.GranjaID)).size > 1;
    const clave = x => (porGranja ? granjaName(x.GranjaID) : [x.Sector, x.Nave].filter(Boolean).join(' / '));
    const dias = [...new Set(activos.map(x => Number(x.EdadDias)))].sort((a, b) => a - b);
    const combinar = (vals, geo) => (geo ? 10 ** (vals.reduce((a, v) => a + Math.log10(Math.max(v, 1)), 0) / vals.length) : vals.reduce((a, v) => a + v, 0) / vals.length);
    const series = (campo, geo) => {
        const grupos = new Map();
        activos.forEach((x) => {
            const v = toNum(x[campo]);
            if (v === null) return;
            if (!grupos.has(clave(x))) grupos.set(clave(x), new Map());
            const g = grupos.get(clave(x));
            const d = Number(x.EdadDias);
            if (!g.has(d)) g.set(d, []);
            g.get(d).push(v);
        });
        return [...grupos.entries()].sort((a, b) => a[0].localeCompare(b[0], 'es', { numeric: true })).map(([k, g], i) => ({
            name: k, color: PALETTE.cat[i % PALETTE.cat.length],
            data: dias.map(d => (g.has(d) ? round(combinar(g.get(d), geo), 2) : null)),
        }));
    };
    return { dias, labels: dias.map(etiquetaSemana), series };
}

/** Gráfico del valor que resume cada muestreo de un método, con el corte y el nivel protector punteados. */
function seroResumenCard(activos, metodo, id, conMetodo) {
    const { dias, labels, series } = seroSeries(activos);
    const referencias = [];
    if (metodo.Prueba !== 'PLACA' && metodo.CortePositivo !== '') {
        referencias.push({ name: `Positivo desde ${fmtNum(metodo.CortePositivo)}`, color: PALETTE.tabla, dashed: true, data: dias.map(() => Number(metodo.CortePositivo)) });
    }
    if (metodo.Prueba !== 'PLACA' && metodo.NivelProtector !== '') {
        referencias.push({ name: `Nivel protector ${fmtNum(metodo.NivelProtector)}`, color: PALETTE.cat[5], dashed: true, data: dias.map(() => Number(metodo.NivelProtector)) });
    }
    const enf = enfermedadById(metodo.EnfermedadID)?.Nombre || '';
    const titulo = { HI: 'Promedio de títulos HI (log2) por edad', ELISA: 'Media geométrica de títulos ELISA por edad', PLACA: 'Aves positivas (%) por edad' }[metodo.Prueba];
    const campo = { HI: 'Promedio', ELISA: 'MediaGeometrica', PLACA: 'PctPositivos' }[metodo.Prueba];
    return chartCard(id, {
        type: 'line', icon: 'syringe', title: conMetodo ? `${enf} · ${nombreMetodo(metodo)}` : `${enf} · ${titulo}`, labels, unir: true,
        unit: metodo.Prueba === 'PLACA' ? '%' : '', series: series(campo, metodo.Prueba === 'ELISA').concat(referencias),
    });
}

/** Pesos estándar (admin): curva de cada línea genética por edad, con su fuente; se puede corregir. */
function viewEstandares() {
    const r = need('getEstandares', {});
    if (!r) return loadingBlock(2);
    if (!r.ok) return errorBlock(r);
    const filas = r.data.estandares;
    const tiposGen = new Set(Object.values(state.db.geneticaDe || {}).flat());
    const geneticas = [...new Set(state.db.listas.filter(l => tiposGen.has(l.Tipo)).map(l => l.Valor))].sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));
    const tipos = state.db.tiposEstandar?.length ? state.db.tiposEstandar : ['Comercial', 'Reproductora'];
    const conTabla = new Set(filas.map(x => `${x.LineaGenetica}|${x.Tipo}`));
    if (!state.estGen || !geneticas.includes(state.estGen)) state.estGen = geneticas.find(g => conTabla.has(`${g}|Comercial`)) || geneticas[0] || '';
    if (!tipos.includes(state.estTipo)) state.estTipo = 'Comercial';
    const curva = filas.filter(x => x.LineaGenetica === state.estGen && x.Tipo === state.estTipo);
    state.estFilas = new Map(curva.map(x => [x.EstandarID, x]));
    const cobertura = geneticas.map(g => `<span class="chip">${esc(g)} ${tipos.map(t => `<span class="badge ${conTabla.has(`${g}|${t}`) ? 'badge-ok' : 'badge-muted'}" title="${esc(t)}">${esc(t === 'Comercial' ? 'C' : 'R')}${conTabla.has(`${g}|${t}`) ? '' : ' —'}</span>`).join(' ')}</span>`).join('');
    const filtros = `<div class="card filters-card"><div class="filters">
        <label class="field"><span>Línea genética</span><select data-change="est-gen">${geneticas.map(g => `<option${g === state.estGen ? ' selected' : ''}>${esc(g)}</option>`).join('')}</select></label>
        <div class="filter-group"><span class="field-label">Tabla</span><div class="segmented" role="group" aria-label="Tabla">${tipos.map(t =>
        `<button type="button" class="seg${t === state.estTipo ? ' is-active' : ''}" data-action="est-tipo" data-tipo="${esc(t)}" aria-pressed="${t === state.estTipo}">${esc(t === 'Comercial' ? 'Comercial (postura y engorde)' : 'Hembra reproductora')}</button>`).join('')}</div></div>
    </div></div>`;
    let html = filtros + `<div class="section-head"><p class="muted">Peso esperado por edad, tomado de la guía oficial de cada casa genética. Con esta tabla se juzga cada pesaje. Corrige una fila solo si la guía cambió (pide justificación y queda en la bitácora).</p>
        <button type="button" class="btn btn-primary" data-action="new-estandar">${ico('plus')}Agregar edad</button></div>`;
    if (!curva.length) {
        html += emptyState(`${state.estGen} no tiene tabla ${state.estTipo === 'Comercial' ? 'comercial' : 'de hembra reproductora'}. Sus pesajes quedan "Sin tabla" hasta que se cargue una guía oficial.`, 'chart-line');
    } else {
        const semanal = curva.length > 2 && Number(curva[1].EdadDias) - Number(curva[0].EdadDias) >= 7;
        const etq = x => (semanal ? `Sem. ${Math.round(Number(x.EdadDias) / 7)}` : `Día ${x.EdadDias}`);
        const spec = {
            title: `${state.estGen} · ${state.estTipo === 'Comercial' ? 'comercial' : 'hembra reproductora'} (${state.unidad})`, subtitle: curva[0].Fuente, icon: 'chart-line', type: 'line', unit: state.unidad,
            labels: curva.map(etq),
            series: [{ name: 'Estándar', color: PALETTE.primary, data: curva.map(x => round(enUnidad(x.EstandarG), decimalesPeso())) }],
        };
        if (curva.some(x => x.MinimoG !== '')) {
            spec.band = { name: 'Rango mín–máx', color: PALETTE.tabla, lo: curva.map(x => (x.MinimoG === '' ? null : round(enUnidad(x.MinimoG), decimalesPeso()))), hi: curva.map(x => (x.MaximoG === '' ? null : round(enUnidad(x.MaximoG), decimalesPeso()))) };
        }
        html += chartCard('estandar', spec) + table('estandares', {
            title: 'Tabla de pesos', icon: 'table', csv: `pesos-estandar-${state.estGen}-${state.estTipo}`,
            columns: [{ k: 'EdadDias', l: 'Edad (días)', t: 'num', sticky: true }, { k: 'semana', l: 'Semana', t: 'num', get: x => Math.floor(Number(x.EdadDias) / 7) },
                { k: 'MinimoG', l: 'Mínimo (g)', t: 'num' }, { k: 'EstandarG', l: 'Estándar (g)', t: 'num' }, { k: 'MaximoG', l: 'Máximo (g)', t: 'num' },
                { k: 'Fuente', l: 'Fuente', wrap: true, html: x => `${esc(x.Fuente)}${x.Alterado ? ` <span class="badge badge-danger" title="Modificada fuera de la plataforma">${ico('lock', 'icon-sm')}</span>` : ''}` }],
            rows: curva,
            actions: x => `<button type="button" class="icon-btn icon-btn-sm" data-action="edit-estandar" data-id="${esc(x.EstandarID)}" title="Corregir" aria-label="Corregir fila">${ico('pencil')}</button>
                <button type="button" class="icon-btn icon-btn-sm danger" data-action="delete-estandar" data-id="${esc(x.EstandarID)}" title="Borrar" aria-label="Borrar fila">${ico('trash')}</button>`,
        });
    }
    return html + `<div class="card"><div class="card-head card-title-center"><h3>${ico('list-checks')}Cobertura de las líneas genéticas</h3></div>
        <p class="muted center">C = tabla comercial · R = hembra reproductora · "—" = sin guía oficial cargada.</p><div class="chips">${cobertura}</div></div>`;
}

const TABLA_LABEL = {
    Usuarios: 'Usuarios', Granjas: 'Granjas', Perfiles: 'Fichas de línea', Listas: 'Listas',
    Sectores: 'Sectores', Naves: 'Naves', Lotes: 'Lotes', Pesajes: 'Pesajes', Ambiente: 'Lecturas de ambiente',
    Registros_Reproductora: 'Registros de reproductora', Registros_Incubadora: 'Registros de incubadora',
    Registros_Engorde: 'Registros de engorde', Registros_Postura: 'Registros de postura',
    PesosEstandar: 'Pesos estándar', Concentrado: 'Compras de concentrado',
    Enfermedades: 'Enfermedades', MetodosSerologia: 'Métodos serológicos', Serologias: 'Muestreos serológicos',
};
/** Una etiqueta por cada acción del servidor (los rechazos de cualquier acción quedan en Accesos). */
const ACCION_LABEL = {
    getSnapshot: 'Ingreso', passwordChanged: 'Cambio de contraseña',
    getRegistros: 'Consulta de registros', getResumen: 'Consulta de resumen', getBitacora: 'Consulta de bitácora',
    saveRegistro: 'Guardar registro', deleteRegistro: 'Eliminar registro', restoreRegistro: 'Restaurar registro',
    getPesajes: 'Consulta de pesajes', savePesaje: 'Guardar pesaje', deletePesaje: 'Eliminar pesaje',
    getAmbiente: 'Consulta de ambiente', saveLectura: 'Guardar lectura', deleteLectura: 'Eliminar lectura',
    saveSector: 'Guardar sector', setSectorEstado: 'Activar o desactivar sector', deleteSector: 'Eliminar sector',
    saveNave: 'Guardar nave', setNaveEstado: 'Activar o desactivar nave', deleteNave: 'Eliminar nave',
    saveLote: 'Guardar lote', cerrarLote: 'Cerrar lote', reabrirLote: 'Reabrir lote', deleteLote: 'Eliminar lote',
    saveGranja: 'Guardar granja', setGranjaEstado: 'Suspender o activar granja', deleteGranja: 'Eliminar granja',
    savePerfil: 'Guardar ficha', deletePerfil: 'Eliminar ficha',
    saveUser: 'Guardar usuario', setUserEstado: 'Suspender o activar usuario', resetPassword: 'Restablecer contraseña', deleteUser: 'Eliminar usuario',
    saveLista: 'Guardar valor de lista', deleteLista: 'Quitar valor de lista',
    getAccesos: 'Consulta de accesos', verifyAudit: 'Verificar integridad',
    getEstandares: 'Consulta de pesos estándar', saveEstandar: 'Guardar peso estándar', deleteEstandar: 'Borrar peso estándar',
    getConcentrado: 'Consulta de concentrado', saveCompra: 'Guardar compra de concentrado', deleteCompra: 'Eliminar compra de concentrado',
    getSerologias: 'Consulta de serología', saveSerologia: 'Guardar muestreo serológico', deleteSerologia: 'Eliminar muestreo serológico',
    saveMetodoSerologia: 'Guardar método serológico', deleteMetodoSerologia: 'Eliminar método serológico',
};
const RESULTADO_LABEL = {
    OK: 'Correcto', NOT_REGISTERED: 'Cuenta sin acceso', SUSPENDED: 'Acceso suspendido',
    ACCOUNT_LOCKED: 'Cuenta bloqueada por alteración', FORBIDDEN: 'Acción no permitida',
};

function auditHtml(r) {
    if (r.ok) return alertBox('ok', 'circle-check', `Todo íntegro. Verificado el ${esc(fmtDateTime(r.verificadoEn))}.`);
    const items = [];
    (r.bitacora?.errores || []).forEach(e => items.push(`Bitácora${e.fila ? `, fila ${e.fila}` : ''}: ${esc(e.motivo)}`));
    (r.alteradas || []).forEach(a => items.push(`${esc(TABLA_LABEL[a.tabla] || a.tabla)}, fila ${esc(a.fila)} (${esc(a.id)}): modificada fuera de la plataforma.`));
    (r.faltantes || []).forEach(f => items.push(`${esc(TABLA_LABEL[f.tabla] || f.tabla)} (${esc(f.id)}): ${esc(f.motivo)}`));
    return alertBox('danger', 'alert-triangle', 'Se detectaron cambios no autorizados. Las filas alteradas quedan bloqueadas hasta que el dueño de la base de datos las revise y apruebe.')
        + `<ul class="issue-list">${items.map(i => `<li>${i}</li>`).join('')}</ul>`;
}

// ==========================================================================
// 11. MODALES Y FORMULARIOS
// ==========================================================================

/**
 * Abre un modal con formulario. onSubmit(FormData, form) puede lanzar un error:
 * se muestra dentro del modal sin perder lo escrito.
 */
function openModal({ title, body, submitLabel = 'Guardar', danger = false, wide = false, onSubmit, cancelLabel = 'Cancelar' }) {
    const dlg = $('#modal');
    // El evento "close" del modal anterior puede llegar después de abrir este: se da por terminado
    // aquí, para que ese evento tardío no borre este contenido ni termine la promesa equivocada.
    terminarModal?.();
    dlg.className = `modal${wide ? ' modal-wide' : ''}`;
    dlg.innerHTML = `<form class="modal-card" novalidate>
        <header class="modal-head"><span></span><h2>${esc(title)}</h2><button type="button" class="icon-btn" data-close aria-label="Cerrar">${ico('x')}</button></header>
        <div class="modal-body">${body}</div>
        <p class="form-error" role="alert" hidden></p>
        <footer class="modal-foot">
            ${cancelLabel ? `<button type="button" class="btn" data-close>${esc(cancelLabel)}</button>` : ''}
            ${onSubmit ? `<button type="submit" class="btn ${danger ? 'btn-danger btn-solid' : 'btn-primary'}">${esc(submitLabel)}</button>` : ''}
        </footer></form>`;

    return new Promise((resolve) => {
        const form = dlg.querySelector('form');
        let result = null;
        dlg.querySelectorAll('[data-close]').forEach(b => b.addEventListener('click', () => dlg.close()));
        // Si al llegar el evento el diálogo ya está abierto otra vez, es el "close" tardío de un modal anterior.
        const alCerrar = () => { if (dlg.open) return; dlg.innerHTML = ''; terminar(); };
        const terminar = () => {
            dlg.removeEventListener('close', alCerrar);
            if (terminarModal === terminar) terminarModal = null;
            resolve(result);
        };
        terminarModal = terminar;
        dlg.addEventListener('close', alCerrar);
        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!onSubmit) return dlg.close();
            const err = form.querySelector('.form-error');
            if (!form.checkValidity()) {
                const bad = form.querySelector('.modal-body :invalid');
                const label = bad?.closest('.field')?.querySelector('span')?.textContent.replace('*', '').trim();
                err.textContent = label ? `${label}: ${validityMsg(bad)}` : 'Revisa los campos marcados.';
                err.hidden = false;
                bad?.focus();
                return;
            }
            const btn = form.querySelector('[type="submit"]');
            btn.disabled = true; btn.classList.add('is-loading'); err.hidden = true;
            try {
                const res = await onSubmit(new FormData(form), form);
                if (res === MANTENER_ABIERTO) return;   // pasos dentro del mismo modal (carga desde Excel)
                result = res ?? true;
                dlg.close();
            } catch (ex) {
                if (ex.code === 'BAD_RESPONSE') {
                    dlg.close();
                    toast(ex.message, 'info', 9000);
                    refreshAll().catch(() => {});
                    return;
                }
                err.textContent = ex.message; err.hidden = false;
            } finally {
                btn.disabled = false; btn.classList.remove('is-loading');
            }
        });
        dlg.showModal();
        form.querySelector('.modal-body input:not([disabled]):not([readonly]):not([type=hidden]), .modal-body select, .modal-body textarea')?.focus();
    });
}
const MANTENER_ABIERTO = Symbol('mantener abierto');
function closeModal() { const d = $('#modal'); if (d.open) d.close(); }
/** Termina el modal en curso (quita su escucha de "close" y resuelve su promesa); ver openModal. */
let terminarModal = null;

/** Mensajes de validación propios (en español, sin depender del idioma del navegador). */
function validityMsg(el) {
    const v = el.validity;
    if (v.valueMissing) return 'es obligatorio.';
    if (v.badInput) return 'escribe un número válido.';
    if (v.tooShort) return `debe tener al menos ${el.minLength} caracteres.`;
    if (v.rangeUnderflow || v.rangeOverflow) return `debe estar entre ${nf.format(Number(el.min))} y ${nf.format(Number(el.max))}.`;
    if (v.stepMismatch) return 'debe ser un número entero.';
    if (v.typeMismatch) return el.type === 'email' ? 'escribe un correo válido.' : 'formato no válido.';
    if (v.patternMismatch) return 'formato no válido.';
    return 'valor no válido.';
}

function toast(msg, type = 'info', ms = 4500) {
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.setAttribute('role', type === 'error' ? 'alert' : 'status');
    el.textContent = msg;
    $('#toasts').appendChild(el);
    setTimeout(() => { el.classList.add('is-leaving'); setTimeout(() => el.remove(), 220); }, ms);
}

function justField(help = 'Explica el motivo. Queda registrado en la bitácora.') {
    return `<label class="field span-all"><span>Justificación <span class="req">*</span></span>
        <textarea name="justificacion" required minlength="${CONFIG.MIN_JUSTIFICACION}" maxlength="1000"></textarea>
        <small class="field-hint">${esc(help)} Mínimo ${CONFIG.MIN_JUSTIFICACION} caracteres.</small></label>`;
}

function inputFor(c, value = '', { unidad = 'g' } = {}) {
    const req = c.req ? ' required' : '';
    const maxBase = c.peso && c.max !== undefined ? round(c.max / factorDe(unidad), 4) : c.max;
    const k = esc(c.k), min = esc(c.min ?? ''), max = esc(maxBase ?? '');
    const lbl = `<span>${esc(c.l)}${c.peso ? ` (<span data-unidad-label>${esc(unidad)}</span>)` : ''}${c.req ? ' <span class="req">*</span>' : ''}</span>`;
    const v = esc(value ?? '');
    if (c.t === 'textoLargo') return `<label class="field span-all">${lbl}<textarea name="${k}" maxlength="1000"${req}>${v}</textarea></label>`;
    if (c.t === 'entero' || c.t === 'numero') {
        return `<label class="field">${lbl}<input type="number" name="${k}"${c.peso ? ` data-peso data-max-g="${esc(c.max ?? '')}"` : ''} inputmode="${c.t === 'entero' ? 'numeric' : 'decimal'}" step="${c.t === 'entero' ? '1' : 'any'}"${min !== '' ? ` min="${min}"` : ''}${max !== '' ? ` max="${max}"` : ''} value="${v}"${req}></label>`;
    }
    if (c.t === 'fecha') return `<label class="field">${lbl}<input type="date" name="${k}" value="${v}"${req}></label>`;
    return `<label class="field">${lbl}<input type="text" name="${k}" maxlength="${esc(c.max || 150)}"${c.lista ? ` list="dl-${esc(c.lista)}"` : ''} value="${v}" autocomplete="off"${req}></label>`;
}

function datalists() {
    const tipos = [...new Set(state.db.listas.map(l => l.Tipo))];
    return tipos.map(t => `<datalist id="dl-${esc(t)}">${state.db.listas.filter(l => l.Tipo === t).map(l => `<option value="${esc(l.Valor)}"></option>`).join('')}</datalist>`).join('');
}

/** Guarda en el servidor, avisa y recarga. `snapshot` = también recargar granjas/usuarios/perfiles. */
async function save(action, payload, msg, { snapshot = true } = {}) {
    const r = await withBusy(() => api(action, payload));
    state.cache.clear();
    if (snapshot) { await withBusy(() => loadSnapshot()); renderChrome(); }
    const texto = typeof msg === 'function' ? msg(r) : msg;
    if (texto) toast(texto, 'success');
    render();
    return r ?? true;
}

function formPayload(fd) {
    const o = {};
    fd.forEach((v, k) => { o[k] = typeof v === 'string' ? v.trim() : v; });
    return o;
}

// ---------- Ubicación (nave + fecha) para registros y pesajes ----------
/** Selector de nave activa (agrupado por granja/sector) y fecha de trabajo. */
function ubicacionInputs(linea, { naveId = '', fecha = today(), conFecha = true, sinVentana = false } = {}) {
    const granjas = varias() ? ctxIds() : [state.granjaId];
    const todas = granjas.flatMap(gid => navesActivas(gid, linea).map(x => ({ gid, ...x })));
    // Un lote que se registra por lote (todas sus naves juntas) aparece una vez, en lugar de sus naves.
    const porLote = new Map();
    const opciones = todas.filter(({ gid, n }) => {
        const l = loteEn(n.NaveID, fecha);
        if (l && modoEn(l, fecha) === 'Lote') { porLote.set(l.LoteID, { gid, l }); return false; }
        return true;
    });
    if (!opciones.length && !porLote.size) return null;
    const grupos = new Map();
    opciones.forEach(({ gid, s, n }) => {
        const k = `${granjas.length > 1 ? granjaName(gid) + ' · ' : ''}${s.Nombre}`;
        if (!grupos.has(k)) grupos.set(k, []);
        grupos.get(k).push({ v: n.NaveID, t: n.Nombre });
    });
    porLote.forEach(({ gid, l }) => {
        const k = `${granjas.length > 1 ? granjaName(gid) + ' · ' : ''}Lotes completos (todas sus naves)`;
        if (!grupos.has(k)) grupos.set(k, []);
        grupos.get(k).push({ v: 'L:' + l.LoteID, t: `Lote ${l.NumeroLote} (${navesVigentes(l).map(x => naveById(x.NaveID)?.Nombre || '').join(', ')})` });
    });
    const valores = [...grupos.values()].flat().map(o => o.v);
    const elegida = naveId && valores.includes(naveId) ? naveId : (state.f.nave && valores.includes(state.f.nave) ? state.f.nave : valores[0]);
    const select = `<label class="field"><span>Nave o lote <span class="req">*</span></span><select name="NaveID" required data-change="ubic-info">${[...grupos.entries()].map(([k, ops]) =>
        `<optgroup label="${esc(k)}">${ops.map(o => `<option value="${esc(o.v)}"${o.v === elegida ? ' selected' : ''}>${esc(o.t)}</option>`).join('')}</optgroup>`).join('')}</select></label>`;
    const min = sinVentana ? '' : fechaMinima();
    const ayuda = sinVentana ? 'Día en que se tomaron las muestras (los resultados pueden llegar después).'
        : isAdmin() ? 'Hoy o cualquier día anterior.' : `Hoy o hasta ${state.db.diasRetroactivos} días atrás.`;
    const fechaHtml = conFecha ? `<label class="field"><span>${sinVentana ? 'Fecha del muestreo' : 'Fecha'} <span class="req">*</span></span><input type="date" name="Fecha" required value="${esc(fecha)}" max="${today()}"${min ? ` min="${min}"` : ''} data-change="ubic-info">
        <small class="field-hint">${ayuda}</small></label>` : '';
    return `${select}${fechaHtml}<div class="span-all" data-ubic-info>${infoUbicacion(elegida, fecha)}</div>`;
}
function infoUbicacion(naveId, fecha) {
    const delLote = String(naveId || '').startsWith('L:');
    const l = loteDeUbicacion(naveId, fecha);
    const n = delLote ? (l ? { Linea: l.Linea, Nombre: `Lote ${l.NumeroLote}` } : null) : naveById(naveId);
    if (!n) return '';
    if (!l) return `<div class="alert alert-warn">${ico('alert-triangle')}<span>${esc(n.Nombre)} no tenía un lote activo el ${esc(fmtDate(fecha))}. Inicia el lote en la vista de la granja.</span></div>`;
    if (modoEn(l, fecha) !== (delLote ? 'Lote' : 'Nave')) {
        return `<div class="alert alert-warn">${ico('alert-triangle')}<span>El lote ${esc(l.NumeroLote)} se registra ${modoEn(l, fecha) === 'Lote' ? 'por lote (todas sus naves juntas)' : 'por nave'} el ${esc(fmtDate(fecha))}: elija ${modoEn(l, fecha) === 'Lote' ? 'el lote' : 'la nave'}.</span></div>`;
    }
    const dias = edadDias(l, fecha);
    const tipo = tipoEstandar(n.Linea);
    const esp = esperado(l.LineaGenetica, tipo, dias);
    const costo = toNum(l.CostoPorAve);
    let extra = costo !== null ? ` · ${esc(fmtQ(costo))} por ave` : '';
    let boton = '';
    if (tipo && dias === null) {
        extra += '<br>El lote no tiene la edad de las aves al iniciar: sin ella no se compara con la tabla de su línea genética.';
        boton = `<button type="button" class="btn btn-sm" data-action="lote-desde-form" data-id="${esc(l.LoteID)}">${ico('tag')}Completar el lote</button>`;
    } else if (tipo) {
        extra += `<br>Edad: <strong>${esc(textoEdad(dias))}</strong>` + (esp
            ? ` · Peso esperado (${esc(l.LineaGenetica)}): <strong>${esc(fmtPeso(esp.estandarG))}</strong>${esp.minimoG !== null ? ` (rango ${esc(fmtPeso(esp.minimoG))} – ${esc(fmtPeso(esp.maximoG))})` : ''}`
            : ` · Sin tabla de peso de ${esc(l.LineaGenetica)} para esta edad.`);
    }
    return `<div class="alert alert-info">${ico('tag')}<span>Lote <strong>${esc(l.NumeroLote)}</strong> · Línea genética <strong>${esc(l.LineaGenetica)}</strong> · desde ${esc(fmtDate(l.FechaInicio))}${extra}</span>${boton}</div>`;
}
/** Peso esperado (g) de la nave en esa fecha, según la línea genética de su lote y la edad. */
function esperadoPesaje(naveId, fecha) {
    const l = loteDeUbicacion(naveId, fecha);
    return l ? esperado(l.LineaGenetica, tipoEstandar(l.Linea), edadDias(l, fecha)) : null;
}
/** Estado del promedio (g) contra el rango: igual que el servidor. */
function estadoPeso(promG, esp) {
    if (!esp) return 'Sin tabla';
    if (esp.minimoG === null || esp.maximoG === null) return 'Sin rango';
    return promG < esp.minimoG ? 'Bajo' : promG > esp.maximoG ? 'Sobre' : 'Dentro';
}
/**
 * Registro: si el lote tiene edad inicial, "Edad (semanas)" se calcula y no se escribe. Si tiene
 * aves iniciales, "Número de aves" también (iniciales − mortalidad y salidas de los días anteriores).
 */
function syncEdadRegistro(form, lote, fecha) {
    // Valores "tabla" que publica la guía de la línea genética del lote: se muestran y no se escriben.
    const diasGuia = edadDias(lote, fecha);
    camposDe(state.linea).filter(c => c.tablaGuia).forEach((c) => {
        const input = form?.elements[c.k];
        if (!input) return;
        const v = lote ? produccionEsperada(lote.LineaGenetica, tipoEstandar(state.linea), c.k, diasGuia) : null;
        input.readOnly = v !== null;
        if (v !== null) input.value = String(round(v, 2));
        else if (input.dataset.guia) input.value = '';
        input.dataset.guia = v !== null ? '1' : '';
        input.title = v !== null ? `De la guía oficial de ${lote.LineaGenetica} a esta edad` : '';
    });
    const aves = form?.elements.NumeroAves;
    if (aves && state.linea !== 'Incubadora') {
        const auto = toNum(lote?.AvesIniciales) > 0;
        aves.readOnly = auto;
        aves.required = !auto;
        if (auto && !aves.dataset.reg) aves.value = '';
        aves.placeholder = auto ? 'Se calcula con el lote' : '';
        aves.title = auto ? 'Aves iniciales del lote menos la mortalidad y las salidas de los días anteriores' : '';
        aves.closest('.field')?.querySelector('.req')?.toggleAttribute('hidden', auto);
    }
    const input = form?.elements.EdadSemanas;
    if (!input) return;
    const dias = edadDias(lote, fecha);
    input.readOnly = dias !== null;
    if (dias !== null) input.value = String(round(dias / 7, 1));
    else if (input.dataset.auto) input.value = '';
    input.dataset.auto = dias !== null ? '1' : '';
    input.title = dias !== null ? 'Se calcula con la edad del lote' : '';
}

// ---------- Registros ----------
async function openRegistroModal(reg = null) {
    const linea = state.linea;
    const campos = camposEntrada(linea);
    const calculados = camposDe(linea).filter(c => c.calc);
    const conPeso = campos.some(c => c.peso);
    const unidad = state.unidad;
    let ubic;
    if (reg) {
        ubic = `<label class="field"><span>Granja</span><input value="${esc(granjaName(reg.GranjaID))}" readonly></label>
            <label class="field"><span>Sector / nave</span><input value="${esc([reg.Sector, reg.Nave].filter(Boolean).join(' / ') || '—')}" readonly></label>
            <label class="field"><span>Lote</span><input value="${esc([reg.NumeroLote, reg.LineaGenetica].filter(Boolean).join(' · ') || '—')}" readonly></label>
            <label class="field"><span>Fecha</span><input value="${esc(fmtDate(reg.Fecha))}" readonly></label>`;
    } else {
        ubic = ubicacionInputs(linea);
        if (!ubic) return toast(`No hay naves activas de ${linea.toLowerCase()} donde registrar.`, 'error');
    }
    // Grupos de campos: generales, propios de la línea y observaciones.
    const zoo = new Set(camposEntrada('Engorde').map(c => c.k));
    const generales = campos.filter(c => (linea === 'Incubadora' ? ['NumeroAves', 'EdadSemanas', 'MortalidadPorcentaje'].includes(c.k) : zoo.has(c.k)) && c.k !== 'Comentarios');
    const propios = campos.filter(c => !generales.includes(c) && c.k !== 'Comentarios');
    const valor = c => (c.peso ? (reg ? round(enUnidad(reg[c.k], unidad), 4) : '') : reg?.[c.k]);
    const sec = (title, icon, list) => (list.length ? `<section class="form-section"><h3>${ico(icon)}${esc(title)}</h3><div class="form-grid">${list.map(c => inputFor(c, valor(c), { unidad })).join('')}</div></section>` : '');
    const body = `<section class="form-section"><div class="form-grid">${ubic}</div></section>
        ${conPeso ? `<section class="form-section"><div class="row row-center"><span class="field-label">Los pesos de este formulario están en</span>${unidadSelect('UnidadPeso', unidad)}</div></section>` : ''}
        ${sec('Datos zootécnicos', LINEA_ICON[linea], generales)}
        ${sec(linea === 'Postura' ? 'Producción de huevo' : linea === 'Incubadora' ? 'Incubación' : 'Otros datos', linea === 'Postura' ? 'egg' : 'archive', propios)}
        ${sec('Observaciones', 'clipboard-list', campos.filter(c => c.k === 'Comentarios'))}
        ${calculados.length ? alertBox('info', 'info', `Se calculan automáticamente: ${calculados.map(c => esc(c.l)).join(', ')}.`) : ''}
        ${reg ? `<section class="form-section">${justField('Explica qué se corrige y por qué.')}</section>` : ''}`;

    const promesa = openModal({
        title: reg ? `Editar registro de ${linea.toLowerCase()}` : `Nuevo registro de ${linea.toLowerCase()}`,
        wide: true, body, submitLabel: reg ? 'Guardar cambios' : 'Guardar registro',
        onSubmit: async (fd) => {
            const p = formPayload(fd);
            const payload = { ...p, Linea: linea };
            if (reg) payload.RegistroID = reg.RegistroID;
            const aves = toNum(p.NumeroAves), muertas = toNum(p.MortalidadNominal);
            if (aves !== null && muertas !== null && muertas > aves) throw new Error('La mortalidad no puede ser mayor que el número de aves.');
            const inc = toNum(p.HuevoIncubable), nac = toNum(p.Nacimiento);
            if (inc !== null && nac !== null && nac > inc) throw new Error('Los nacimientos no pueden ser más que los huevos incubables.');
            await save('saveRegistro', payload, reg ? 'Registro actualizado.' : 'Registro guardado.', { snapshot: false });
        },
    });
    const form = $('#modal form');
    if (reg && form?.elements.NumeroAves) form.elements.NumeroAves.dataset.reg = '1';
    if (reg) syncEdadRegistro(form, loteById(reg.LoteID), reg.Fecha);
    else if (form) syncEdadRegistro(form, loteDeUbicacion(form.NaveID.value, form.Fecha.value), form.Fecha.value);
    await promesa;
}

function unidadSelect(name, actual) {
    return `<select class="select select-inline" name="${esc(name)}" data-change="form-unidad" data-prev="${esc(actual)}" aria-label="Unidad de peso">${Object.keys(UNIDADES).map(u =>
        `<option value="${u}"${u === actual ? ' selected' : ''}>${UNIDAD_NOMBRE[u]} (${u})</option>`).join('')}</select>`;
}

async function deleteRegistroFlow(reg) {
    await openModal({
        title: 'Eliminar registro', danger: true, submitLabel: 'Eliminar',
        body: `<p class="center">Registro del <strong>${esc(fmtDate(reg.Fecha))}</strong> · ${esc([granjaName(reg.GranjaID), reg.Sector, reg.Nave].filter(Boolean).join(' / '))} · ${esc(fmtNum(reg.NumeroAves))} aves.</p>
            <p class="muted center">Dejará de verse y de contarse en el dashboard. Queda guardado en el historial${isAdmin() ? ' y puedes restaurarlo' : ' y la administración puede restaurarlo'}.</p>
            ${justField()}`,
        onSubmit: async fd => save('deleteRegistro', { Linea: state.linea, RegistroID: reg.RegistroID, justificacion: fd.get('justificacion') }, 'Registro eliminado.', { snapshot: false }),
    });
}

async function restoreRegistroFlow(reg) {
    await openModal({
        title: 'Restaurar registro', submitLabel: 'Restaurar',
        body: `<p class="center">Registro del <strong>${esc(fmtDate(reg.Fecha))}</strong> · ${esc([granjaName(reg.GranjaID), reg.Sector, reg.Nave].filter(Boolean).join(' / '))}.</p>${justField()}`,
        onSubmit: async fd => save('restoreRegistro', { Linea: state.linea, RegistroID: reg.RegistroID, justificacion: fd.get('justificacion') }, 'Registro restaurado.', { snapshot: false }),
    });
}

// ---------- Pesajes por muestra ----------
const borradorKey = (naveId, fecha) => `granjas.pesaje.${naveId}.${fecha}`;
function leerBorrador(naveId, fecha) {
    try { return JSON.parse(sessionStorage.getItem(borradorKey(naveId, fecha)) || 'null'); } catch { return null; }
}
function guardarBorrador() {
    if (!pz || pz.editando || !pz.naveId) return;
    try { sessionStorage.setItem(borradorKey(pz.naveId, pz.fecha), JSON.stringify({ unidad: pz.unidad, muestra: pz.muestra, pesos: pz.pesos })); } catch { /* sin almacenamiento */ }
}
function borrarBorrador(naveId, fecha) {
    try { sessionStorage.removeItem(borradorKey(naveId, fecha)); } catch { /* sin almacenamiento */ }
}

/** Estadísticas de la muestra en la unidad de captura (las mismas fórmulas que el servidor). */
function statsMuestra(pesos) {
    if (!pesos.length) return null;
    const prom = pesos.reduce((a, b) => a + b, 0) / pesos.length;
    const sd = pesos.length > 1 ? Math.sqrt(pesos.reduce((a, v) => a + (v - prom) ** 2, 0) / (pesos.length - 1)) : 0;
    const dentro = pesos.filter(v => Math.abs(v - prom) <= prom * 0.1).length;
    return { prom, min: Math.min(...pesos), max: Math.max(...pesos), cv: prom ? sd / prom * 100 : 0, unif: dentro / pesos.length * 100 };
}

function pzRender() {
    const box = $('[data-pz]');
    if (!box || !pz) return;
    const n = pz.pesos.length;
    const total = pz.muestra || 0;
    const quien = pz.tipo === 'huevo' ? 'huevo' : 'ave';
    box.querySelector('[data-pz-count]').textContent = `${nf.format(n)} de ${nf.format(total)}`;
    box.querySelector('[data-pz-bar]').style.width = `${total ? Math.min(100, n / total * 100) : 0}%`;
    box.querySelector('[data-pz-bar]').classList.toggle('is-full', n === total && total > 0);
    const input = box.querySelector('#pz-input');
    input.placeholder = n < total ? `Peso del ${quien} #${n + 1} (${pz.unidad})` : 'Muestra completa';
    input.disabled = n >= total;
    const st = statsMuestra(pz.pesos);
    const f = v => new Intl.NumberFormat('es-GT', { maximumFractionDigits: pz.unidad === 'g' ? 1 : 3 }).format(v);
    // Comparación en vivo con el peso esperado de la línea genética (en la unidad de captura).
    const esp = pz.tipo === 'huevo' ? null : pz.esperado;
    const promG = st ? st.prom * factorDe(pz.unidad) : null;
    const comparacion = st && esp ? (() => {
        const dif = promG - esp.estandarG;
        const estado = estadoPeso(promG, esp);
        return `<div><dt>Esperado</dt><dd>${esc(fmtPeso(esp.estandarG, pz.unidad))}</dd></div>
            <div><dt>Diferencia</dt><dd>${esc(`${dif >= 0 ? '+' : ''}${fmtPeso(dif, pz.unidad)} (${dif >= 0 ? '+' : ''}${fmtNum(dif / esp.estandarG * 100)} %)`)}</dd></div>
            <div><dt>Estado</dt><dd>${badgeEstadoPeso(estado)}</dd></div>`;
    })() : '';
    box.querySelector('[data-pz-stats]').innerHTML = st
        ? [['Promedio', `${f(st.prom)} ${pz.unidad}`], ['Mínimo', `${f(st.min)} ${pz.unidad}`], ['Máximo', `${f(st.max)} ${pz.unidad}`], ['CV', `${fmtNum(st.cv)} %`], ['Uniformidad', `${fmtNum(st.unif)} %`]]
            .map(([l, v]) => `<div><dt>${esc(l)}</dt><dd>${esc(v)}</dd></div>`).join('') + comparacion
        : `<p class="muted center">Ingresa el primer peso.${esp ? ` Peso esperado: ${esc(fmtPeso(esp.estandarG, pz.unidad))}${esp.minimoG !== null ? ` (rango ${esc(fmtPeso(esp.minimoG, pz.unidad))} – ${esc(fmtPeso(esp.maximoG, pz.unidad))})` : ''}.` : ''}</p>`;
    box.querySelector('[data-pz-list]').innerHTML = pz.pesos.map((v, i) =>
        `<li><span class="pz-num">${i + 1}</span>${esc(f(v))}<button type="button" class="icon-btn icon-btn-sm danger" data-action="pz-del" data-i="${i}" aria-label="Quitar el peso ${i + 1}" title="Quitar">${ico('x')}</button></li>`).join('');
    const excedente = n > total;
    box.querySelector('[data-pz-aviso]').innerHTML = excedente ? alertBox('warn', 'alert-triangle', `Hay ${n - total} peso(s) de más para una muestra de ${total}. Quita los sobrantes o aumenta la muestra.`) : '';
    guardarBorrador();
}

/** Agrega uno o varios pesos (separados por espacio, coma, punto y coma o salto de línea). */
function pzAgregar(texto) {
    if (!pz) return;
    const partes = String(texto).split(/[\s;,]+/).map(s => s.trim()).filter(Boolean);
    const malos = [];
    partes.forEach((s) => {
        const v = Number(s);
        if (!Number.isFinite(v) || v <= 0) { malos.push(s); return; }
        if (pz.pesos.length < pz.muestra) pz.pesos.push(v);
    });
    if (malos.length) toast(`No se agregaron valores no válidos: ${malos.slice(0, 5).join(' ')}`, 'error');
    pzRender();
}

async function openPesajeModal(pes = null) {
    const linea = pes ? pes.Linea : state.linea;
    const tipo = linea === 'Incubadora' ? 'huevo' : 'ave';
    let ubic;
    if (pes) {
        ubic = `<label class="field"><span>Sector / nave</span><input value="${esc([pes.Sector, pes.Nave].join(' / '))}" readonly></label>
            <label class="field"><span>Fecha</span><input value="${esc(fmtDate(pes.Fecha))}" readonly></label>`;
    } else {
        ubic = ubicacionInputs(linea);
        if (!ubic) return toast(`No hay naves activas de ${linea.toLowerCase()}.`, 'error');
    }
    pz = pes
        ? { editando: true, tipo, unidad: pes.Unidad || 'g', muestra: toNum(pes.Muestra) || 0, pesos: String(pes.Pesos || '').split(',').map(Number).filter(Number.isFinite) }
        : { editando: false, tipo, unidad: state.unidad === 'g' || tipo === 'huevo' ? 'g' : state.unidad, muestra: 100, pesos: [] };
    const body = `<section class="form-section"><div class="form-grid">${ubic}</div></section>
        <section class="form-section"><h3>${ico('list-checks')}Muestra</h3><div class="form-grid">
            <label class="field"><span>Unidad de pesaje <span class="req">*</span></span><select name="Unidad" data-change="pz-config">${Object.keys(UNIDADES).map(u => `<option value="${u}"${u === pz.unidad ? ' selected' : ''}>${UNIDAD_NOMBRE[u]} (${u})</option>`).join('')}</select>
                <small class="field-hint">Se usa para toda la muestra.</small></label>
            <label class="field"><span>Tamaño de la muestra <span class="req">*</span></span><input type="number" name="Muestra" min="1" max="${state.db.maxMuestra}" step="1" inputmode="numeric" required value="${pz.muestra}" data-change="pz-config">
                <small class="field-hint">Número de ${tipo === 'huevo' ? 'huevos' : 'aves'} que se van a pesar.</small></label>
        </div></section>
        <section class="form-section pz" data-pz>
            <h3>${ico('scale')}Peso de cada ${tipo}</h3>
            <div class="pz-progress"><div class="pz-bar" data-pz-bar></div></div>
            <p class="center"><strong data-pz-count></strong> pesados</p>
            <div class="pz-entrada">
                <input id="pz-input" class="input" type="text" inputmode="decimal" autocomplete="off" aria-label="Peso">
                <button type="button" class="btn btn-primary" data-action="pz-add">${ico('plus')}Agregar</button>
                <button type="button" class="btn" data-action="pz-undo" title="Quitar el último">${ico('undo')}Deshacer</button>
            </div>
            <small class="field-hint center">Presiona Enter después de cada peso. También puedes pegar varios separados por espacios o comas.</small>
            <div data-pz-aviso></div>
            <dl class="kv pz-stats" data-pz-stats></dl>
            <ol class="pz-list" data-pz-list></ol>
        </section>
        ${pes ? `<section class="form-section">${justField('Explica qué se corrige (por ejemplo, un peso mal digitado).')}</section>` : ''}`;

    // Si hay una muestra a medias guardada en este dispositivo, se recupera.
    if (!pes) {
        const sel = /name="NaveID"[^>]*>[\s\S]*?<option value="([^"]+)" selected/.exec(ubic);
        const b = sel && leerBorrador(sel[1], today());
        if (b && b.pesos?.length) { Object.assign(pz, { unidad: b.unidad, muestra: b.muestra, pesos: b.pesos }); setTimeout(() => toast(`Se recuperó una muestra sin guardar (${b.pesos.length} pesos).`, 'info'), 300); }
    }

    const promesa = openModal({
        title: pes ? 'Editar pesaje' : `Nuevo pesaje · ${linea.toLowerCase()}`, wide: true, body,
        submitLabel: pes ? 'Guardar cambios' : 'Guardar pesaje',
        onSubmit: async (fd) => {
            pz.unidad = fd.get('Unidad');
            pz.muestra = toNum(fd.get('Muestra'));
            if (pz.pesos.length !== pz.muestra) {
                throw new Error(`La muestra es de ${pz.muestra} y hay ${pz.pesos.length} pesos. Deben ser exactamente ${pz.muestra}.`);
            }
            const payload = { Unidad: pz.unidad, Muestra: pz.muestra, Pesos: pz.pesos };
            if (pes) Object.assign(payload, { PesajeID: pes.PesajeID, justificacion: fd.get('justificacion') });
            else Object.assign(payload, { Linea: linea, NaveID: fd.get('NaveID'), Fecha: fd.get('Fecha') });
            await save('savePesaje', payload, pes ? 'Pesaje actualizado.' : 'Pesaje guardado.', { snapshot: false });
            if (!pes) borrarBorrador(payload.NaveID, payload.Fecha);
        },
    });
    const form = $('#modal form');
    if (!pes && form) { pz.naveId = form.NaveID.value; pz.fecha = form.Fecha.value; pz.esperado = esperadoPesaje(pz.naveId, pz.fecha); }
    if (pes) {
        const l = loteById(pes.LoteID);
        pz.esperado = l ? esperado(l.LineaGenetica, tipoEstandar(pes.Linea), edadDias(l, pes.Fecha)) : null;
    }
    pzRender();
    $('#pz-input')?.focus();
    await promesa;
    pz = null;
}

async function verPesajeModal(pes) {
    const valores = String(pes.Pesos || '').split(',').filter(Boolean);
    await openModal({
        title: `Pesaje del ${fmtDate(pes.Fecha)}`, cancelLabel: 'Cerrar', wide: true,
        body: `<p class="center">${esc([granjaName(pes.GranjaID), pes.Sector, pes.Nave, pes.NumeroLote ? 'Lote ' + pes.NumeroLote : ''].filter(Boolean).join(' / '))}</p>
            <dl class="kv">${[['Muestra', pes.Muestra], ['Promedio', fmtPeso(pes.PromedioG)], ['Mínimo', fmtPeso(pes.MinimoG)], ['Máximo', fmtPeso(pes.MaximoG)],
            ['CV', `${fmtNum(pes.CV)} %`], ['Uniformidad', `${fmtNum(pes.Uniformidad)} %`], ['Conversión alimenticia', pes.ConversionAlimenticia ? fmtNum(pes.ConversionAlimenticia) : '—']]
            .map(([l, v]) => `<div><dt>${esc(l)}</dt><dd>${esc(v)}</dd></div>`).join('')}</dl>
            <section class="form-section"><h3>${ico('list-checks')}Pesos individuales (${esc(pes.Unidad)})</h3>
            <ol class="pz-list">${valores.map((v, i) => `<li><span class="pz-num">${i + 1}</span>${esc(v)}</li>`).join('')}</ol></section>`,
    });
}

async function deletePesajeFlow(pes) {
    await openModal({
        title: 'Eliminar pesaje', danger: true, submitLabel: 'Eliminar',
        body: `<p class="center">Pesaje del <strong>${esc(fmtDate(pes.Fecha))}</strong> · ${esc([pes.Sector, pes.Nave].join(' / '))} · muestra de ${esc(pes.Muestra)}.</p>${justField()}`,
        onSubmit: async fd => save('deletePesaje', { PesajeID: pes.PesajeID, justificacion: fd.get('justificacion') }, 'Pesaje eliminado.', { snapshot: false }),
    });
}

// ---------- Ambiente (estrés calórico) ----------
async function openLecturaModal() {
    const ubic = ubicacionInputs(state.linea, { conFecha: false });
    if (!ubic) return toast('No hay naves activas.', 'error');
    await openModal({
        title: 'Nueva lectura de ambiente',
        body: `<div class="form-grid">${ubic}
            <label class="field"><span>Temperatura ambiente en grados Celsius (°C) <span class="req">*</span></span><input type="number" name="Temperatura" step="any" min="-20" max="60" inputmode="decimal" required data-input="iec-calc">
                <small class="field-hint">Escríbela en °C (no en °F).</small></label>
            <label class="field"><span>Humedad relativa (%) <span class="req">*</span></span><input type="number" name="Humedad" step="any" min="0" max="100" inputmode="decimal" required data-input="iec-calc"></label>
            <div class="span-all" data-iec-preview><p class="muted center">IEC = temperatura (°C) + humedad (%).</p></div>
        </div>
        <p class="muted center">La fecha y la hora se registran automáticamente al guardar.</p>`,
        submitLabel: 'Guardar lectura',
        onSubmit: async fd => save('saveLectura', { NaveID: fd.get('NaveID'), Temperatura: fd.get('Temperatura'), Humedad: fd.get('Humedad') }, (r) => `Lectura guardada: IEC ${fmtNum(r.Indice)} (${r.Nivel}).`, { snapshot: false }),
    });
}

async function deleteLecturaFlow(x) {
    await openModal({
        title: 'Eliminar lectura', danger: true, submitLabel: 'Eliminar',
        body: `<p class="center">${esc(fmtDateTime(x.FechaHora))} · ${esc([x.Sector, x.Nave].join(' / '))} · IEC ${esc(fmtNum(x.Indice))}.</p>${justField()}`,
        onSubmit: async fd => save('deleteLectura', { LecturaID: x.LecturaID, justificacion: fd.get('justificacion') }, 'Lectura eliminada.', { snapshot: false }),
    });
}

// ---------- Estructura: sectores, naves y lotes ----------
async function openSectorModal(granjaId, linea, s = null) {
    await openModal({
        title: s ? `Renombrar ${s.Nombre}` : `Nuevo sector · ${linea.toLowerCase()}`,
        body: `<div class="form-grid"><label class="field span-all"><span>Nombre del sector <span class="req">*</span></span><input name="Nombre" required maxlength="80" value="${esc(s?.Nombre || '')}" autocomplete="off"></label>
            ${s ? justField('Los registros anteriores conservan el nombre que tenían.') : ''}</div>`,
        submitLabel: s ? 'Guardar' : 'Crear sector',
        onSubmit: async (fd) => {
            const p = { Nombre: fd.get('Nombre') };
            if (s) Object.assign(p, { SectorID: s.SectorID, justificacion: fd.get('justificacion') });
            else Object.assign(p, { GranjaID: granjaId, Linea: linea });
            await save('saveSector', p, s ? 'Sector renombrado.' : 'Sector creado.');
        },
    });
}

async function openNaveModal(sectorId, n = null) {
    const s = sectorById(sectorId || n?.SectorID);
    const linea = n?.Linea || s?.Linea;
    // Postura y reproductora: tipo de nave (piso o jaula), nidales y bebederos de campana.
    const conEquipo = linea === 'Postura' || linea === 'Reproductora';
    const num = (k, l, ayuda, paso = '1') => `<label class="field"><span>${l}</span><input name="${k}" type="number" min="0" step="${paso}" inputmode="decimal" value="${esc(n?.[k] ?? '')}">${ayuda ? `<small class="field-hint">${ayuda}</small>` : ''}</label>`;
    await openModal({
        title: n ? `Editar ${n.Nombre}` : `Nueva nave en ${s?.Nombre || ''}`,
        body: `<div class="form-grid"><label class="field span-all"><span>Nombre de la nave <span class="req">*</span></span><input name="Nombre" required maxlength="80" value="${esc(n?.Nombre || '')}" autocomplete="off"></label>
            ${num('AreaM2', 'Área de la nave (m²)', `Para calcular las aves por m² (recomendado ${linea === 'Engorde' ? '10–12' : '6–7 en piso'}).`, '0.01')}
            ${conEquipo ? `<label class="field"><span>Tipo de nave</span><select name="Tipo">${(state.db.tiposNave || ['Piso', 'Jaula']).map(t => `<option${(n?.Tipo || 'Piso') === t ? ' selected' : ''}>${t}</option>`).join('')}</select></label>
                ${num('Nidales', 'Nidales', 'Recomendado: 4–5 nidales por cada 20 gallinas (en piso).')}
                ${num('BebederosCampana', 'Bebederos de campana', 'Recomendado: un bebedero por cada 65–70 aves (en piso).')}` : ''}
            ${n ? justField('Los registros anteriores conservan el nombre que tenían.') : ''}</div>`,
        submitLabel: n ? 'Guardar' : 'Crear nave',
        onSubmit: async (fd) => {
            const p = { Nombre: fd.get('Nombre'), AreaM2: fd.get('AreaM2') };
            if (conEquipo) Object.assign(p, { Tipo: fd.get('Tipo'), Nidales: fd.get('Nidales'), BebederosCampana: fd.get('BebederosCampana') });
            if (n) Object.assign(p, { NaveID: n.NaveID, justificacion: fd.get('justificacion') });
            else p.SectorID = sectorId;
            await save('saveNave', p, n ? 'Nave actualizada.' : 'Nave creada.');
        },
    });
}

/**
 * Referencias de manejo de la nave (según su línea y tipo): aves por m², aves por nidal y aves por
 * bebedero de campana, con las aves vivas de hoy. En un lote que se registra por lote, se usan las aves
 * y las medidas de todas sus naves juntas.
 */
function referenciasNaveHtml(n) {
    const refLinea = state.db.referenciasNave?.[n.Linea];
    if (!refLinea || n.Linea === 'Incubadora') return '';
    const tipo = n.Linea === 'Engorde' ? '' : (n.Tipo || 'Piso');
    const ref = n.Linea === 'Engorde' ? refLinea : (refLinea[tipo] || {});
    const info = state.db.avesNave?.[n.NaveID];
    let naves = [n];
    if (info?.lote) naves = navesVigentes(loteById(info.LoteID)).map(x => naveById(x.NaveID)).filter(Boolean);
    const total = k => (naves.every(x => toNum(x[k]) !== null) ? naves.reduce((a, x) => a + toNum(x[k]), 0) : null);
    const aves = info ? info.aves : null;
    const fila = (nombre, divisor, rango, unidad, faltaTxt) => {
        const r = rango ? `${rango[0]}–${rango[1]} ${unidad}` : '';
        if (!divisor) return `<li><span>${nombre}</span><span class="muted">${esc(faltaTxt)}${r ? ` · recomendado ${r}` : ''}</span></li>`;
        if (aves === null) return `<li><span>${nombre}</span><span class="muted">Sin lote abierto${r ? ` · recomendado ${r}` : ''}</span></li>`;
        const v = aves / divisor;
        const estado = !rango ? '' : v > rango[1] ? '<span class="badge badge-warn">Alto</span>' : v < rango[0] ? '<span class="badge badge-muted">Bajo</span>' : '<span class="badge badge-ok">Dentro</span>';
        return `<li><span>${nombre}</span><span><strong>${esc(fmtNum(round(v, 1)))}</strong> ${unidad}${r ? ` <span class="muted">(recomendado ${r})</span>` : ''} ${estado}</span></li>`;
    };
    const filas = [fila('Densidad', total('AreaM2'), ref.densidad, 'aves/m²', 'Anote el área de la nave')];
    if (ref.avesPorNidal) filas.push(fila('Nidales', total('Nidales'), ref.avesPorNidal, 'aves por nidal', 'Anote los nidales'));
    if (ref.avesPorBebedero) filas.push(fila('Bebederos de campana', total('BebederosCampana'), ref.avesPorBebedero, 'aves por bebedero', 'Anote los bebederos'));
    return `<div class="nave-ref"><span class="nave-ref-titulo">${ico('list-checks', 'icon-sm')}Referencias${tipo ? ` · ${esc(tipo.toLowerCase())}` : ''}${info?.lote ? ' · lote completo' : ''}</span>
        ${tipo === 'Jaula' ? '<p class="muted">Sin referencias para jaula todavía.</p>' : ''}<ul>${filas.join('')}</ul></div>`;
}

/** Activar/desactivar un sector o una nave (lo pasado no cambia; lo nuevo se bloquea). */
async function estadoEstructuraFlow(tipo, row) {
    const activar = row.Estado !== 'Activo';
    const nombre = tipo === 'sector' ? 'el sector' : 'la nave';
    await openModal({
        title: `${activar ? 'Activar' : 'Desactivar'} ${row.Nombre}`, danger: !activar, submitLabel: activar ? 'Activar' : 'Desactivar',
        body: `<p class="center">${activar ? `Se podrán volver a ingresar datos en ${nombre}.` : `No se podrán ingresar datos nuevos en ${nombre}${tipo === 'sector' ? ' ni en sus naves' : ''}. Todo su historial se conserva y sigue visible.`}</p>${justField()}`,
        onSubmit: async (fd) => {
            const p = { Estado: activar ? 'Activo' : 'Inactivo', justificacion: fd.get('justificacion') };
            if (tipo === 'sector') p.SectorID = row.SectorID; else p.NaveID = row.NaveID;
            await save(tipo === 'sector' ? 'setSectorEstado' : 'setNaveEstado', p, activar ? 'Activado.' : 'Desactivado.');
        },
    });
}

async function deleteEstructuraFlow(tipo, row) {
    await openModal({
        title: `Eliminar ${row.Nombre}`, danger: true, submitLabel: 'Eliminar',
        body: `<p class="center">Solo es posible si ${tipo === 'sector' ? 'el sector no tiene naves' : 'la nave nunca tuvo lotes ni datos'}. Si ya se usó, desactívala para conservar el historial.</p>${justField()}`,
        onSubmit: async (fd) => {
            const p = { justificacion: fd.get('justificacion') };
            if (tipo === 'sector') p.SectorID = row.SectorID; else p.NaveID = row.NaveID;
            await save(tipo === 'sector' ? 'deleteSector' : 'deleteNave', p, 'Eliminado.');
        },
    });
}

/** Opciones de línea genética según la línea de producción (desplegable con grupos). */
function geneticaSelect(linea, actual = '') {
    const tipos = state.db.geneticaDe?.[linea] || [];
    const grupos = tipos.map(t => [t, state.db.listas.filter(l => l.Tipo === t).sort((a, b) => String(a.Valor).localeCompare(String(b.Valor), 'es', { numeric: true }))]);
    return `<select name="LineaGenetica" required><option value="">Elige la línea genética…</option>${grupos.map(([t, vals]) =>
        `<optgroup label="${esc(LISTA_LABEL[t] || t)}">${vals.map(v => `<option${v.Valor === actual ? ' selected' : ''}>${esc(v.Valor)}</option>`).join('')}</optgroup>`).join('')}</select>`;
}

async function openLoteModal(naveId, lote = null, previo = null) {
    const n = naveById(naveId || lote?.NaveID);
    // Al iniciar es obligatoria (salvo Incubadora, que pesa huevos); un lote anterior se edita sin ella.
    const edadObligatoria = !lote && n?.Linea !== 'Incubadora';
    const base = lote || previo || {};
    // Naves: al iniciar, las libres de la misma granja y línea (la elegida, marcada); al editar un lote
    // de varias naves, las aves de cada una (el total es su suma).
    const libres = lote ? [] : navesActivas(n.GranjaID, n.Linea).filter(x => !loteActivo(x.n.NaveID));
    const varias = lote ? navesDelLote(lote) : [];
    const avesHtml = lote && varias.length > 1
        ? `<div class="span-all"><span class="field-label">Aves que entraron a cada nave</span><div class="form-grid">${varias.map(x => `<label class="field"><span>${esc(naveById(x.NaveID)?.Nombre || '')}${x.Hasta ? ' (salió)' : ''}</span>
            <input type="number" name="aves-${esc(x.NaveID)}" min="0" step="1" inputmode="numeric" value="${esc(x.AvesIniciales || '')}"></label>`).join('')}</div></div>`
        : lote ? `<label class="field"><span>Aves iniciales</span><input type="number" name="AvesIniciales" min="0" step="1" inputmode="numeric" value="${esc(lote.AvesIniciales || '')}"></label>`
        : `<div class="span-all"><span class="field-label">Naves (galeras) del lote y aves que entraron a cada una <span class="req">*</span></span>
            <div class="lote-naves">${libres.map(({ s, n: nv }) => `<label class="lote-nave"><span class="check"><input type="checkbox" name="nave" value="${esc(nv.NaveID)}"${nv.NaveID === n.NaveID ? ' checked' : ''}> ${esc(s.Nombre)} / ${esc(nv.Nombre)}</span>
                <input type="number" name="aves-${esc(nv.NaveID)}" min="0" step="1" inputmode="numeric" placeholder="Aves" aria-label="Aves en ${esc(nv.Nombre)}"></label>`).join('')}</div>
            <small class="field-hint">Un lote puede ocupar varias naves (por ejemplo, el lote 52 en las galeras 5 y 6). Cada nave tiene un solo lote a la vez.</small></div>
            <label class="field span-all"><span>Forma de registro diario</span><select name="Registro"><option value="Nave">Por nave (cada galera por separado; la app suma el lote)</option><option value="Lote">Por lote (un solo total del lote cada día)</option></select>
                <small class="field-hint">Por nave guarda el detalle de cada galera. Por lote, se registra el total del día sin sumar a mano.</small></label>`;
    await openModal({
        title: lote ? `Editar lote ${lote.NumeroLote}` : `Iniciar lote en ${n?.Nombre || ''}`, wide: !lote,
        body: `<div class="form-grid">
            <label class="field"><span>Número de lote <span class="req">*</span></span><input name="NumeroLote" required maxlength="60" value="${esc(lote?.NumeroLote || '')}" autocomplete="off"></label>
            <label class="field"><span>Línea genética <span class="req">*</span></span>${geneticaSelect(n.Linea, base.LineaGenetica)}</label>
            <label class="field"><span>Fecha de inicio <span class="req">*</span></span><input type="date" name="FechaInicio" required max="${today()}" value="${esc(base.FechaInicio || today())}"></label>
            ${avesHtml}
            <label class="field"><span>Edad de las aves al iniciar (semanas)${edadObligatoria ? ' <span class="req">*</span>' : ''}</span><input type="number" name="EdadInicialSemanas" min="0" max="200" step="any" inputmode="decimal" value="${esc(base.EdadInicialSemanas ?? '')}"${edadObligatoria ? ' required' : ''}>
                <small class="field-hint">0 si son pollitos de un día; por ejemplo 17 si llegan pollas a postura. Con ella se calcula la edad de cada pesaje.</small></label>
            <label class="field"><span>Costo por ave (Q)${lote ? '' : ' <span class="req">*</span>'}</span><input type="number" name="CostoPorAve" min="0" max="100000" step="0.01" inputmode="decimal" value="${esc(base.CostoPorAve ?? '')}"${lote ? '' : ' required'}></label>
            ${lote ? justField('Si cambias el número de lote o la línea genética, se actualizan también sus registros y pesajes.') : `<label class="check span-all"><input type="checkbox" name="otro"> Al guardar, iniciar otro lote (con la misma línea genética, fecha, edad y costo)</label>`}
        </div>`,
        submitLabel: lote ? 'Guardar cambios' : 'Iniciar lote',
        onSubmit: async (fd) => {
            const p = formPayload(fd);
            const otro = !lote && fd.get('otro');
            delete p.otro;
            Object.keys(p).filter(k => k.startsWith('aves-') || k === 'nave').forEach(k => delete p[k]);
            if (lote) {
                p.LoteID = lote.LoteID;
                if (varias.length > 1) p.Naves = varias.map(x => ({ NaveID: x.NaveID, AvesIniciales: fd.get('aves-' + x.NaveID) }));
            } else {
                const elegidas = fd.getAll('nave');
                if (!elegidas.length) throw new Error('Elija al menos una nave para el lote.');
                p.Naves = elegidas.map(id => ({ NaveID: id, AvesIniciales: fd.get('aves-' + id) }));
                if (elegidas.length > 1 && p.Naves.some(x => !x.AvesIniciales)) throw new Error('Con varias naves, escriba las aves que entraron a cada una.');
            }
            await save('saveLote', p, lote ? 'Lote actualizado.' : `Lote ${p.NumeroLote} iniciado.`);
            if (otro) {
                const sigue = navesActivas(n.GranjaID, n.Linea).find(x => !loteActivo(x.n.NaveID));
                if (sigue) setTimeout(() => openLoteModal(sigue.n.NaveID, null, { LineaGenetica: p.LineaGenetica, FechaInicio: p.FechaInicio, EdadInicialSemanas: p.EdadInicialSemanas, CostoPorAve: p.CostoPorAve }), 50);
                else toast('Ya no quedan naves libres en esta línea.');
            }
        },
    });
}

/** Naves de un lote: entrar otra nave, sacar una antes del cierre y la forma de registro. */
async function lotesNavesModal(lote) {
    if (!lote) return;
    const filas = navesDelLote(lote);
    const vigentes = filas.filter(x => !x.Hasta);
    const libres = navesActivas(lote.GranjaID, lote.Linea).filter(x => !loteActivo(x.n.NaveID));
    const forma = modoEn(lote, today());
    await openModal({
        title: `Naves del lote ${lote.NumeroLote}`, wide: true, cancelLabel: 'Cerrar',
        body: `<div class="table-wrap"><table><thead><tr><th>Nave</th><th class="num">Aves al entrar</th><th>Desde</th><th>Salió</th><th></th></tr></thead><tbody>
            ${filas.map(x => `<tr><td>${esc(naveById(x.NaveID)?.Nombre || '—')}</td><td class="num">${esc(fmtNum(x.AvesIniciales) || '—')}</td><td>${esc(fmtDate(x.Desde))}</td>
                <td>${x.Hasta ? `${esc(fmtDate(x.Hasta))}${x.AvesSalida ? ` (${esc(fmtNum(x.AvesSalida))} aves)` : ''}` : '—'}</td>
                <td>${!x.Hasta && vigentes.length > 1 && lote.Estado === 'Activo' ? `<button type="button" class="btn btn-sm" data-action="sacar-nave" data-lote="${esc(lote.LoteID)}" data-id="${esc(x.NaveID)}">Sacar</button>` : ''}</td></tr>`).join('')}
            </tbody></table></div>
            <p class="center">Se registra <strong>${forma === 'Lote' ? 'por lote (todas sus naves juntas)' : 'por nave'}</strong>${lote.RegistroDesde ? ` desde el ${esc(fmtDate(lote.RegistroDesde))}` : ''}.
                ${lote.Estado === 'Activo' ? `<button type="button" class="btn btn-sm" data-action="forma-lote" data-id="${esc(lote.LoteID)}">Cambiar la forma de registro</button>` : ''}</p>
            ${lote.Estado === 'Activo' && libres.length ? `<section class="form-section"><h3>${ico('plus')}Agregar una nave al lote</h3><div class="form-grid" data-agregar-nave>
                <label class="field"><span>Nave</span><select name="NaveNueva">${libres.map(({ s, n }) => `<option value="${esc(n.NaveID)}">${esc(s.Nombre)} / ${esc(n.Nombre)}</option>`).join('')}</select></label>
                <label class="field"><span>Aves que entran</span><input type="number" name="AvesNueva" min="1" step="1" inputmode="numeric"></label>
                <label class="field"><span>Desde</span><input type="date" name="DesdeNueva" min="${esc(lote.FechaInicio)}" max="${today()}" value="${today()}"></label>
                <div class="row row-center span-all"><button type="button" class="btn btn-primary" data-action="agregar-nave-lote" data-id="${esc(lote.LoteID)}">${ico('plus')}Agregar nave</button></div></div></section>` : ''}`,
    });
}

async function agregarNaveLote(el) {
    const box = el.closest('[data-agregar-nave]');
    const v = k => box.querySelector(`[name="${k}"]`).value;
    if (!v('AvesNueva')) { toast('Escriba las aves que entran a la nave.', 'error'); return; }
    closeModal();
    await save('agregarNaveLote', { LoteID: el.dataset.id, NaveID: v('NaveNueva'), AvesIniciales: v('AvesNueva'), Desde: v('DesdeNueva') }, 'Nave agregada al lote.');
}

async function sacarNaveModal(loteId, naveId) {
    const lote = loteById(loteId);
    const porLote = modoEn(lote, today()) === 'Lote';
    await openModal({
        title: `Sacar ${naveById(naveId)?.Nombre || 'la nave'} del lote ${lote.NumeroLote}`, submitLabel: 'Sacar nave',
        body: `<div class="form-grid"><label class="field"><span>Libre desde <span class="req">*</span></span><input type="date" name="Hasta" required min="${esc(lote.FechaInicio)}" max="${today()}" value="${today()}"></label>
            ${porLote ? '<label class="field"><span>Aves que salen con la nave <span class="req">*</span></span><input type="number" name="AvesSalida" min="0" step="1" inputmode="numeric" required></label>'
                : '<p class="muted span-all">Salen las aves que le quedaban a la nave según sus registros.</p>'}
            ${justField()}</div>`,
        onSubmit: async fd => save('sacarNaveLote', { LoteID: loteId, NaveID: naveId, ...formPayload(fd) }, 'La nave quedó libre.'),
    });
}

async function formaLoteModal(lote) {
    const actual = modoEn(lote, today());
    await openModal({
        title: `Forma de registro del lote ${lote.NumeroLote}`, submitLabel: 'Guardar',
        body: `<div class="form-grid"><label class="field"><span>Registrar</span><select name="Registro">
                <option value="Lote"${actual === 'Lote' ? ' selected' : ''}>Por lote (un total diario)</option><option value="Nave"${actual === 'Nave' ? ' selected' : ''}>Por nave (cada galera)</option></select></label>
            <label class="field"><span>Desde</span><input type="date" name="Desde" min="${esc(lote.FechaInicio)}" max="${today()}" value="${today()}"></label>
            ${justField()}</div>
            <p class="muted center">Los días anteriores quedan como están. No se mezclan las dos formas el mismo día; de "por lote" no se puede volver a "por nave".</p>`,
        onSubmit: async fd => save('cambiarRegistroLote', { LoteID: lote.LoteID, ...formPayload(fd) }, 'Forma de registro actualizada.'),
    });
}

async function openNavesVariasModal(sectorId) {
    const s = sectorById(sectorId);
    await openModal({
        title: `Agregar varias naves · ${s?.Nombre || ''}`, submitLabel: 'Crear naves',
        body: `<div class="form-grid"><label class="field"><span>Nombre <span class="req">*</span></span><input name="Prefijo" required maxlength="60" value="Galera"></label>
            <label class="field"><span>Del número</span><input type="number" name="Desde" min="0" max="9999" value="1" required></label>
            <label class="field"><span>Al número</span><input type="number" name="Hasta" min="0" max="9999" value="8" required></label></div>
            <p class="muted center">Crea "Galera 1", "Galera 2"… (hasta 50 de una vez). Las que ya existen se saltan.</p>`,
        onSubmit: async fd => save('saveNavesVarias', { SectorID: sectorId, ...formPayload(fd) }, r => `${r?.creadas || 0} naves creadas.`),
    });
}

async function cerrarLoteFlow(lote) {
    await openModal({
        title: `Cerrar lote ${lote.NumeroLote}`, submitLabel: 'Cerrar lote',
        body: `<div class="form-grid"><label class="field span-all"><span>Fecha de cierre <span class="req">*</span></span><input type="date" name="FechaCierre" required min="${esc(lote.FechaInicio)}" max="${today()}" value="${today()}"></label>
            <label class="field span-all"><span>Comentario</span><textarea name="justificacion" maxlength="1000"></textarea></label></div>
            <p class="muted center">Después del cierre no se podrán ingresar datos de este lote. Sus datos se conservan.</p>`,
        onSubmit: async fd => save('cerrarLote', { LoteID: lote.LoteID, FechaCierre: fd.get('FechaCierre'), justificacion: fd.get('justificacion') }, 'Lote cerrado.'),
    });
}

async function reabrirLoteFlow(lote) {
    await openModal({
        title: `Reabrir lote ${lote.NumeroLote}`, submitLabel: 'Reabrir',
        body: `<p class="center">El lote volverá a quedar activo y se podrán ingresar datos.</p>${justField()}`,
        onSubmit: async fd => save('reabrirLote', { LoteID: lote.LoteID, justificacion: fd.get('justificacion') }, 'Lote reabierto.'),
    });
}

// ---------- Empresas (plataforma) ----------
async function openEmpresaModal(e = null) {
    const f = (k, l, attrs = '', req = false) => `<label class="field"><span>${l}${req ? ' <span class="req">*</span>' : ''}</span><input name="${k}" value="${esc(e?.[k] || '')}" ${attrs}${req ? ' required' : ''}></label>`;
    await openModal({
        title: e ? `Editar ${e.Nombre}` : 'Nueva empresa', wide: true,
        body: `<div class="form-grid">
            ${f('Nombre', 'Nombre de la empresa', 'maxlength="150" autocomplete="off"', true)}
            ${f('Contacto', 'Persona de contacto', 'maxlength="150"')}
            ${f('NIT', 'NIT', 'maxlength="30"')}
            ${f('Telefono', 'Teléfono', 'type="tel" inputmode="tel" maxlength="30" pattern="[0-9+()\\-\\s]{6,30}"')}
            ${f('Email', 'Correo de la empresa', 'type="email" maxlength="254"')}
            <label class="field span-all"><span>Notas</span><textarea name="Notas" maxlength="1000">${esc(e?.Notas || '')}</textarea></label>
            ${e ? justField() : ''}</div>
            ${e ? '' : `<section class="form-section"><h3>${ico('user-plus')}Administrador de la empresa (quien paga)</h3>
                <p class="muted center">Entrará a todas las granjas de la empresa, verá sus costos y creará sus cuentas. Si su correo ya administra otra empresa, esta se suma a su mismo cobro; si es nuevo, se le dará una contraseña temporal.</p>
                <div class="form-grid">
                    <label class="field"><span>Nombre</span><input name="AdminNombre" maxlength="120" autocomplete="off"></label>
                    <label class="field"><span>Correo electrónico</span><input name="AdminEmail" type="email" maxlength="254" autocomplete="off" data-change="admin-empresa"></label>
                    <div class="span-all" data-dueno-aviso></div>
                </div></section>
                <p class="muted center">Se crea su carpeta. Las granjas de la empresa se agregan después desde Granjas.</p>`}`,
        submitLabel: e ? 'Guardar cambios' : 'Crear empresa',
        onSubmit: async (fd) => {
            const p = { Nombre: fd.get('Nombre'), Contacto: fd.get('Contacto'), NIT: fd.get('NIT'), Telefono: fd.get('Telefono'), Email: fd.get('Email'), Notas: fd.get('Notas') };
            if (e) { p.EmpresaID = e.EmpresaID; p.justificacion = fd.get('justificacion'); }
            else if (fd.get('AdminEmail') || fd.get('AdminNombre')) {
                if (!fd.get('AdminEmail') || !fd.get('AdminNombre')) throw new Error('Para crear el administrador escriba su nombre y su correo.');
                p.Admin = { Nombre: fd.get('AdminNombre'), Email: fd.get('AdminEmail') };
            }
            const r = await save('saveEmpresa', p, e ? 'Empresa actualizada.' : r2 => (r2?.unida ? `Empresa creada y sumada al cobro de ${r2.dueno.Nombre}.` : p.Admin ? null : 'Empresa creada.'));
            if (r?.admin?.claveTemporal) setTimeout(() => showTempPassword(r.admin.Email, r.admin.claveTemporal, r.admin.cuentaExistente), 50);
        },
    });
}

/** Al escribir el correo del administrador: si ya administra empresas, la nueva se suma a su cobro. */
async function avisoDuenoEmpresa(input) {
    const box = input.form?.querySelector('[data-dueno-aviso]');
    const correo = input.value.trim();
    if (!box) return;
    if (!correo || !input.checkValidity()) { box.innerHTML = ''; return; }
    try {
        const r = await api('previewEmpresaAdmin', { AdminEmail: correo });
        if (input.value.trim() !== correo) return;
        const nombre = input.form.elements.AdminNombre;
        if (r.conflicto) box.innerHTML = alertBox('danger', 'alert-triangle', esc(r.mensaje));
        else if (r.existe) {
            box.innerHTML = alertBox('info', 'receipt', `<strong>Se sumará al cobro de ${esc(r.dueno.Nombre)}</strong>, que ya tiene ${r.dueno.empresas.length === 1 ? '1 empresa' : r.dueno.empresas.length + ' empresas'} (${esc(r.dueno.empresas.join(', '))}). ${esc(r.Nombre)} entrará también a esta empresa con su misma cuenta.`);
            if (nombre && !nombre.value) nombre.value = r.Nombre;
        } else box.innerHTML = '<p class="muted center">Correo nuevo: se creará su cuenta con una contraseña temporal y un cobro propio.</p>';
    } catch (e) {
        box.innerHTML = '';
    }
}

async function moverEmpresaModal(e) {
    const otros = (state.db.duenos || []).filter(d => d.DuenoID !== e.DuenoID);
    await openModal({
        title: `Cambiar el dueño de ${e.Nombre}`, submitLabel: 'Cambiar dueño',
        body: `<p class="center">Desde hoy la empresa pasa al cobro del dueño elegido. Lo de antes de hoy se cobra al dueño actual. Sus administradores dejan de entrar a las demás empresas del dueño actual.</p>
            <div class="form-grid"><label class="field span-all"><span>Nuevo dueño <span class="req">*</span></span><select name="DuenoID">
                <option value="">Un dueño nuevo (cobro propio)</option>${otros.map(d => `<option value="${esc(d.DuenoID)}">${esc(d.Nombre)}</option>`).join('')}</select></label>${justField()}</div>`,
        onSubmit: async fd => save('moverEmpresa', { EmpresaID: e.EmpresaID, DuenoID: fd.get('DuenoID'), justificacion: fd.get('justificacion') }, 'Dueño cambiado.'),
    });
}

async function estadoEmpresaFlow(e) {
    const activar = e.Estado !== 'Activa';
    await openModal({
        title: `${activar ? 'Activar' : 'Suspender'} ${e.Nombre}`, danger: !activar, submitLabel: activar ? 'Activar' : 'Suspender',
        body: `<p class="center">${activar ? 'Sus cuentas vuelven a registrar y modificar datos de inmediato.'
            : 'Sus cuentas podrán ver y consultar su información, pero no registrar, editar, cargar ni eliminar datos. Verán un aviso en todas las pantallas.'}</p>${justField()}`,
        onSubmit: async fd => save('setEmpresaEstado', { EmpresaID: e.EmpresaID, Estado: activar ? 'Activa' : 'Suspendida', justificacion: fd.get('justificacion') }, activar ? 'Empresa activada.' : 'Empresa suspendida.'),
    });
}

async function deleteEmpresaFlow(e) {
    await openModal({
        title: `Eliminar ${e.Nombre}`, danger: true, submitLabel: 'Eliminar empresa',
        body: '<p class="center">Solo es posible si la empresa no tiene granjas ni cuentas. Su carpeta se archiva (no se borra). Si tiene granjas, suspéndela.</p>' + justField(),
        onSubmit: async fd => save('deleteEmpresa', { EmpresaID: e.EmpresaID, justificacion: fd.get('justificacion') }, 'Empresa eliminada.'),
    });
}

// ---------- Granjas ----------
async function openGranjaModal(g = null) {
    const f = (k, l, attrs = '', req = false) => `<label class="field"><span>${l}${req ? ' <span class="req">*</span>' : ''}</span><input name="${k}" value="${esc(g?.[k] || '')}" ${attrs}${req ? ' required' : ''}></label>`;
    const empresas = (state.db.empresas || []).filter(e => e.Estado === 'Activa' || e.EmpresaID === g?.EmpresaID);
    if (!g && !empresas.length) { toast('Primero cree una empresa activa.', 'error'); return; }
    const pre = g?.EmpresaID || (isAdmin() ? state.empresaFiltro : state.empresaId) || (empresas.length === 1 ? empresas[0].EmpresaID : '');
    // La empresa se elige al crear la granja (su archivo nace en la carpeta de la empresa) y no cambia después.
    const empresaHtml = g
        ? `<label class="field"><span>Empresa</span><input value="${esc(empresaName(g.EmpresaID))}" readonly></label>`
        : `<label class="field"><span>Empresa <span class="req">*</span></span><select name="EmpresaID" required>${empresas.length > 1 ? '<option value="">Elija la empresa</option>' : ''}${empresas.map(e =>
            `<option value="${esc(e.EmpresaID)}"${e.EmpresaID === pre ? ' selected' : ''}>${esc(e.Nombre)}</option>`).join('')}</select></label>`;
    await openModal({
        title: g ? `Editar ${g.NombreGranja}` : 'Nueva granja', wide: true,
        body: `<div class="form-grid">
            ${empresaHtml}
            ${f('NombreGranja', 'Nombre de la granja', 'maxlength="150" autocomplete="off"', true)}
            ${f('Propietario', 'Propietario', 'maxlength="150"')}
            ${f('Ubicacion', 'Ubicación', 'maxlength="250"')}
            ${f('Telefono', 'Teléfono', 'type="tel" inputmode="tel" maxlength="30" pattern="[0-9+()\\-\\s]{6,30}"')}
            ${f('EmailContacto', 'Correo de contacto', 'type="email" maxlength="254"')}
            <label class="field span-all"><span>Notas</span><textarea name="Notas" maxlength="1000">${esc(g?.Notas || '')}</textarea></label>
            ${g ? justField() : ''}</div>
            ${g ? '' : `<p class="muted center">Después de crearla, agregue sus líneas de producción desde la vista de la granja.${esEmpresa() ? ' Cada línea suma a su pago mensual el precio de su plan y su tramo de aves (lo verá antes de confirmar).' : ''}</p>`}`,
        submitLabel: g ? 'Guardar cambios' : 'Crear granja',
        onSubmit: async (fd) => {
            const p = formPayload(fd);
            if (g) p.GranjaID = g.GranjaID;
            const r = await save('saveGranja', p, g ? 'Granja actualizada.' : 'Granja creada.');
            if (!g && r?.GranjaID) go('granja', r.GranjaID);
        },
    });
}

async function estadoGranjaFlow(g) {
    const activar = g.Estado !== 'Activo';
    await openModal({
        title: `${activar ? 'Activar' : 'Suspender'} ${g.NombreGranja}`, danger: !activar, submitLabel: activar ? 'Activar' : 'Suspender',
        body: `<p class="center">${activar ? 'Sus representantes volverán a tener acceso.' : 'Sus representantes perderán el acceso de inmediato. Los datos se conservan.'}</p>${justField()}`,
        onSubmit: async fd => save('setGranjaEstado', { GranjaID: g.GranjaID, Estado: activar ? 'Activo' : 'Suspendido', justificacion: fd.get('justificacion') }, activar ? 'Granja activada.' : 'Granja suspendida.'),
    });
}

async function deleteGranjaFlow(g) {
    await openModal({
        title: `Eliminar ${g.NombreGranja}`, danger: true, submitLabel: 'Eliminar granja',
        body: '<p class="center">Solo es posible si la granja no tiene registros de producción. Se eliminan también sus fichas. Si tiene registros, suspéndela.</p>' + justField(),
        onSubmit: async (fd) => {
            await save('deleteGranja', { GranjaID: g.GranjaID, justificacion: fd.get('justificacion') }, 'Granja eliminada.');
            go('granjas');
        },
    });
}

// ---------- Perfiles (ficha de cada línea) ----------
async function openPerfilModal(granjaId, linea, p = null) {
    const g = granjaById(granjaId);
    const fuentes = p ? [] : LINEAS.filter(l => l !== linea).map(l => perfilDe(granjaId, l)).filter(Boolean);
    const secciones = [];
    state.db.campos.perfil.forEach((c) => {
        let s = secciones.find(x => x.s === c.s);
        if (!s) { s = { s: c.s, campos: [] }; secciones.push(s); }
        s.campos.push(c);
    });
    // Valores sugeridos al crear: los datos de la granja (el resto se puede copiar de la otra línea).
    const inicial = p || { NombreGranja: g?.NombreGranja || '', Ubicacion: g?.Ubicacion || '' };
    const planHtml = p ? '' : `<section class="form-section"><h3>${ico('tag')}Plan de la línea</h3>
        <div class="form-grid" data-granja="${esc(granjaId)}" data-linea="${esc(linea)}">${selectorPlan('', '', true)}</div></section>`;
    const body = planHtml + (fuentes.length ? alertBox('info', 'copy', `Puedes partir de otra ficha de esta granja y luego ajustarla: ${fuentes.map(f => `<button type="button" class="btn btn-sm" data-action="copy-perfil" data-from="${esc(f.PerfilID)}">${ico(LINEA_ICON[f.Linea])}${esc(f.Linea)}</button>`).join(' ')}`) : '')
        + secciones.map(s => `<section class="form-section"><h3>${ico(SECCION_ICON[s.s] || 'list')}${esc(s.s)}</h3><div class="form-grid">${s.campos.map(c => inputFor(c, inicial[c.k])).join('')}</div></section>`).join('')
        + (p ? `<section class="form-section"><div class="form-grid">${justField('Explica qué se actualiza (por ejemplo, un lote nuevo).')}</div></section>` : '')
        + datalists();
    await openModal({
        title: `${p ? 'Editar' : 'Crear'} ficha de ${linea.toLowerCase()} · ${g?.NombreGranja || ''}`, wide: true, body,
        submitLabel: p ? 'Guardar cambios' : 'Crear perfil',
        onSubmit: async (fd) => {
            const payload = formPayload(fd);
            if (p) payload.PerfilID = p.PerfilID;
            else {
                if (!payload.PlanID) throw new Error('Elija el plan de la línea.');
                if (!payload.TramoID) delete payload.TramoID;
                Object.assign(payload, { GranjaID: granjaId, Linea: linea });
            }
            await save('savePerfil', payload, p ? 'Ficha actualizada.' : 'Perfil creado.');
        },
    });
}

/** Al editar el estado se reenvía la ficha completa (el servidor valida todos los campos). */
async function estadoPerfilFlow(p) {
    const activar = p.Estado !== 'Activo';
    await openModal({
        title: `${activar ? 'Activar' : 'Suspender'} línea de ${p.Linea.toLowerCase()}`, danger: !activar, submitLabel: activar ? 'Activar' : 'Suspender',
        body: `<p class="center">${activar ? 'Se podrán volver a ingresar registros de esta línea.' : 'No se podrán ingresar registros nuevos de esta línea. Los existentes se conservan.'}</p>${justField()}`,
        onSubmit: async (fd) => {
            const payload = { PerfilID: p.PerfilID, Estado: activar ? 'Activo' : 'Suspendido', justificacion: fd.get('justificacion') };
            state.db.campos.perfil.forEach((c) => { payload[c.k] = p[c.k] ?? ''; });
            await save('savePerfil', payload, activar ? 'Línea activada.' : 'Línea suspendida.');
        },
    });
}

async function deletePerfilFlow(p) {
    await openModal({
        title: `Eliminar línea de ${p.Linea.toLowerCase()}`, danger: true, submitLabel: 'Eliminar',
        body: '<p class="center">Solo es posible si la línea no tiene registros de producción. Si los tiene, suspéndela.</p>' + justField(),
        onSubmit: async fd => save('deletePerfil', { PerfilID: p.PerfilID, justificacion: fd.get('justificacion') }, 'Perfil eliminado.'),
    });
}

// ---------- Usuarios ----------
const ROL_OPCION = { REPRESENTANTE: 'Representante de granja', ADMIN_EMPRESA: 'Administrador de empresa', ADMIN: 'Plataforma (CHIIK)' };

async function openUserModal(u = null, granjaPre = null) {
    const rol = u?.Rol || ROLES.REPRESENTANTE;
    const asignadas = String(u?.GranjaIDs || granjaPre || '').split(',').filter(Boolean);
    const self = u?.ID === state.me.ID;
    const orden = Object.keys(ROL_OPCION);
    const roles = (state.db.roles || []).filter(r => ROL_OPCION[r]).sort((a, b) => orden.indexOf(a) - orden.indexOf(b));
    // Empresa de la cuenta: la plataforma la elige; el administrador de empresa solo crea cuentas de la suya.
    const empresas = isAdmin() || (esEmpresa() && !u) ? (state.db.empresas || []) : [];
    const empresaPre = u?.EmpresaID || (granjaPre && granjaById(granjaPre)?.EmpresaID) || (isAdmin() ? state.empresaFiltro : state.empresaId) || (empresas.length === 1 ? empresas[0].EmpresaID : '') || state.me.EmpresaID || '';
    const empresaHtml = empresas.length > 1
        ? `<label class="field" data-empresa${rol === ROLES.ADMIN ? ' hidden' : ''}><span>Empresa <span class="req">*</span></span><select name="EmpresaID" data-change="user-empresa"><option value="">Elija la empresa</option>${empresas.map(e =>
            `<option value="${esc(e.EmpresaID)}"${e.EmpresaID === empresaPre ? ' selected' : ''}>${esc(e.Nombre)}</option>`).join('')}</select></label>` : '';
    const chips = state.db.granjas.length ? state.db.granjas.map(g => `<label class="chip check" data-empresa-de="${esc(g.EmpresaID || '')}"${empresaPre && g.EmpresaID !== empresaPre ? ' hidden' : ''}><input type="checkbox" name="GranjaIDs" value="${esc(g.GranjaID)}"${asignadas.includes(g.GranjaID) ? ' checked' : ''}> ${esc(g.NombreGranja)}</label>`).join('')
        : `<p class="muted">${isAdmin() ? 'Primero crea una granja.' : 'Su empresa aún no tiene granjas.'}</p>`;
    await openModal({
        title: u ? `Editar ${u.Nombre}` : 'Nuevo usuario', wide: true,
        body: `<div class="form-grid">
            <label class="field"><span>Nombre <span class="req">*</span></span><input name="Nombre" required maxlength="120" value="${esc(u?.Nombre || '')}"></label>
            <label class="field"><span>Correo electrónico <span class="req">*</span></span><input name="Email" type="email" required maxlength="254" value="${esc(u?.Email || '')}"${u ? ' readonly' : ''} autocomplete="off"></label>
            <label class="field"><span>Nivel de acceso <span class="req">*</span></span><select name="Rol" data-change="user-rol"${self ? ' disabled' : ''}>
                ${roles.map(r => `<option value="${r}"${rol === r ? ' selected' : ''}>${ROL_OPCION[r]}</option>`).join('')}</select>
                ${self ? `<input type="hidden" name="Rol" value="${esc(rol)}"><small class="field-hint">No puedes cambiar tu propio nivel.</small>` : ''}</label>
            ${empresaHtml}
            <label class="field"><span>Teléfono</span><input name="Telefono" type="tel" inputmode="tel" maxlength="30" value="${esc(u?.Telefono || '')}"></label>
            <fieldset class="field span-all" data-granjas${rol === ROLES.REPRESENTANTE ? '' : ' hidden'}>
                <span class="field-label">Granjas a las que tendrá acceso <span class="req">*</span></span>
                <div class="chips">${chips}</div>
                <small class="field-hint center">Si tiene varias, verá cada granja por separado.</small>
            </fieldset>
            <p class="muted center span-all" data-nota-empresa${rol === ROLES.ADMIN_EMPRESA ? '' : ' hidden'}>El administrador de empresa entra a todas las granjas de su empresa y administra sus cuentas.</p>
            ${u ? justField() : ''}</div>
            ${u ? '' : '<p class="muted center">Se generará una contraseña temporal. En su primer ingreso deberá crear su propia contraseña.</p>'}`,
        submitLabel: u ? 'Guardar cambios' : 'Crear usuario',
        onSubmit: async (fd) => {
            const p = { Nombre: fd.get('Nombre'), Rol: fd.get('Rol'), Telefono: fd.get('Telefono'), GranjaIDs: fd.getAll('GranjaIDs') };
            if (empresas.length > 1 && p.Rol !== ROLES.ADMIN) {
                p.EmpresaID = fd.get('EmpresaID');
                if (!p.EmpresaID) throw new Error('Elija la empresa de la cuenta.');
                p.GranjaIDs = p.GranjaIDs.filter(id => granjaById(id)?.EmpresaID === p.EmpresaID);
            }
            if (p.Rol !== ROLES.REPRESENTANTE) p.GranjaIDs = [];
            else if (!p.GranjaIDs.length) throw new Error('Selecciona al menos una granja.');
            if (u) { p.ID = u.ID; p.justificacion = fd.get('justificacion'); } else p.Email = fd.get('Email');
            const r = await save('saveUser', p, u ? 'Usuario actualizado.' : null);
            if (!u && r?.claveTemporal) setTimeout(() => showTempPassword(p.Email, r.claveTemporal, r.cuentaExistente), 50);
        },
    });
}

async function showTempPassword(email, clave, existente) {
    await openModal({
        title: 'Contraseña temporal', cancelLabel: 'Listo',
        body: `<p class="center">Entrégale estos datos a la persona por un medio seguro. <strong>No se volverán a mostrar.</strong></p>
            <div class="stack"><span class="field-label center">Correo</span><div class="secret">${esc(email)}</div>
            <span class="field-label center">Contraseña temporal</span><div class="secret" id="temp-pass">${esc(clave)}</div></div>
            <div class="row row-center"><button type="button" class="btn btn-primary" data-action="copy-temp">${ico('copy')}Copiar contraseña</button></div>
            ${existente ? alertBox('info', 'info', 'Esta cuenta ya existía; se le asignó esta contraseña temporal.') : ''}
            <p class="muted center">En su primer ingreso deberá crear su propia contraseña.</p>`,
    });
}

async function estadoUserFlow(u) {
    const activar = u.Estado !== 'Activo';
    await openModal({
        title: `${activar ? 'Activar' : 'Suspender'} a ${u.Nombre}`, danger: !activar, submitLabel: activar ? 'Activar' : 'Suspender',
        body: `<p class="center">${activar ? 'Podrá volver a iniciar sesión.' : 'Su sesión se cerrará y no podrá volver a entrar hasta que la actives.'}</p>${justField()}`,
        onSubmit: async fd => save('setUserEstado', { ID: u.ID, Estado: activar ? 'Activo' : 'Suspendido', justificacion: fd.get('justificacion') }, activar ? 'Usuario activado.' : 'Usuario suspendido.'),
    });
}

async function resetUserFlow(u) {
    await openModal({
        title: `Nueva contraseña para ${u.Nombre}`, submitLabel: 'Generar contraseña temporal',
        body: `<p class="center">Se cerrarán sus sesiones abiertas y deberá crear una contraseña nueva al entrar.</p>${justField()}`,
        onSubmit: async (fd) => {
            const r = await save('resetPassword', { ID: u.ID, justificacion: fd.get('justificacion') }, null);
            if (r?.claveTemporal) setTimeout(() => showTempPassword(u.Email, r.claveTemporal, false), 50);
        },
    });
}

async function deleteUserFlow(u) {
    await openModal({
        title: `Eliminar a ${u.Nombre}`, danger: true, submitLabel: 'Eliminar usuario',
        body: `<p class="center">Se elimina su cuenta y su acceso. Los registros que ingresó se conservan.</p>${justField()}`,
        onSubmit: async fd => save('deleteUser', { ID: u.ID, justificacion: fd.get('justificacion') }, 'Usuario eliminado.'),
    });
}

// ---------- Compras de concentrado ----------
async function openCompraModal(c = null) {
    const granjas = varias() ? state.db.granjas.filter(g => g.Estado === 'Activo') : state.db.granjas.filter(g => g.GranjaID === state.granjaId);
    if (!c && !granjas.length) return toast('No hay granjas activas donde registrar la compra.', 'error');
    const pre = c?.GranjaID || (varias() && state.f.granjas.length === 1 ? state.f.granjas[0] : granjas[0]?.GranjaID);
    const granjaHtml = c
        ? `<label class="field"><span>Granja</span><input value="${esc(granjaName(c.GranjaID))}" readonly></label>`
        : `<label class="field"><span>Granja <span class="req">*</span></span><select name="GranjaID" required${granjas.length === 1 ? ' disabled' : ''}>${granjas.map(g => `<option value="${esc(g.GranjaID)}"${g.GranjaID === pre ? ' selected' : ''}>${esc(g.NombreGranja)}</option>`).join('')}</select>
            ${granjas.length === 1 ? `<input type="hidden" name="GranjaID" value="${esc(granjas[0].GranjaID)}">` : ''}</label>`;
    const fases = [...new Set(Object.values(state.fasesPorGranja || {}).flat())].sort((a, b) => a.localeCompare(b, 'es'));
    await openModal({
        title: c ? 'Editar compra de concentrado' : 'Nueva compra de concentrado', wide: true,
        body: `<div class="form-grid">${granjaHtml}
            <label class="field"><span>Fase de alimento <span class="req">*</span></span><input name="Fase" required maxlength="80" list="fases-alimento" autocomplete="off" placeholder="Por ejemplo: Inicio, Crecimiento 2" value="${esc(c && c.Fase !== 'Sin fase' ? c.Fase : '')}">
                <datalist id="fases-alimento">${fases.map(f => `<option value="${esc(f)}">`).join('')}</datalist><small class="muted">El nombre que usa su granja; se sugieren las fases ya registradas.</small></label>
            <label class="field"><span>Empresa de alimento</span><input name="Proveedor" maxlength="80" list="proveedores-alimento" autocomplete="off" placeholder="Quién vendió el concentrado" value="${esc(c?.Proveedor || '')}">
                <datalist id="proveedores-alimento">${(state.proveedoresAlimento || []).map(f => `<option value="${esc(f)}">`).join('')}</datalist></label>
            <label class="field"><span>Cubre desde <span class="req">*</span></span><input type="date" name="Desde" required max="${today()}" value="${esc(c?.Desde || today())}"></label>
            <label class="field"><span>Cubre hasta <span class="req">*</span></span><input type="date" name="Hasta" required value="${esc(c?.Hasta || '')}"></label>
            <label class="field"><span>Cantidad comprada (quintales) <span class="req">*</span></span><input type="number" name="CantidadQuintales" required min="0.01" step="any" inputmode="decimal" value="${esc(c?.CantidadQuintales || '')}" data-input="compra-calc"></label>
            <label class="field"><span>Costo por quintal (Q) <span class="req">*</span></span><input type="number" name="CostoPorQuintal" required min="0" step="0.01" inputmode="decimal" value="${esc(c?.CostoPorQuintal || '')}" data-input="compra-calc"></label>
            <div class="span-all" data-compra-calc>${costoQuintalHtml(c?.CantidadQuintales, c?.CostoPorQuintal)}</div>
            <label class="field span-all"><span>Comentarios</span><textarea name="Comentarios" maxlength="1000">${esc(c?.Comentarios || '')}</textarea></label>
            ${c ? justField() : ''}</div>
            <p class="muted center">La fecha y la hora de registro las pone la plataforma al guardar.</p>`,
        submitLabel: c ? 'Guardar cambios' : 'Guardar compra',
        onSubmit: async (fd) => {
            const p = formPayload(fd);
            if (p.Hasta < p.Desde) throw new Error('La fecha "hasta" no puede ser anterior a "desde".');
            if (c) Object.assign(p, { CompraID: c.CompraID, GranjaID: c.GranjaID });
            await save('saveCompra', p, c ? 'Compra actualizada.' : 'Compra guardada.', { snapshot: false });
        },
    });
}
/** Costo total = quintales × costo por quintal (el servidor lo calcula igual al guardar). */
function costoQuintalHtml(q, cpq) {
    const n = toNum(q), c = toNum(cpq);
    return n && c !== null ? `<div class="alert alert-info">${ico('wheat')}<span>Costo total: <strong>${esc(fmtQ(Math.round(n * c * 100) / 100))}</strong></span></div>` : '';
}
async function deleteCompraFlow(c) {
    await openModal({
        title: 'Eliminar compra de concentrado', danger: true, submitLabel: 'Eliminar',
        body: `<p class="center">${esc(c.Fase)}: ${esc(fmtNum(c.CantidadQuintales))} qq por ${esc(fmtQ(c.CostoTotal))}, del ${esc(fmtDate(c.Desde))} al ${esc(fmtDate(c.Hasta))}.</p>${justField()}`,
        onSubmit: async fd => save('deleteCompra', { CompraID: c.CompraID, justificacion: fd.get('justificacion') }, 'Compra eliminada.', { snapshot: false }),
    });
}

// ---------- Carga desde Excel ----------
// Plantilla por tipo de dato (y granja y línea): columnas con desplegables y validaciones; al
// subirla, cada fila se convierte en lo mismo que enviaría el formulario y el servidor la revisa
// con sus reglas antes de guardar (importarDatos).
const IMPORT_TIPOS = {
    registro: { titulo: l => `Registros diarios de ${l.toLowerCase()}`, porLinea: true },
    pesaje: { titulo: l => `Pesajes de ${l.toLowerCase()}`, porLinea: true },
    lectura: { titulo: l => `Lecturas de temperatura y humedad (${l.toLowerCase()})`, porLinea: true },
    serologia: { titulo: l => `Muestreos serológicos (${l.toLowerCase()})`, porLinea: true },
    compra: { titulo: () => 'Compras de concentrado', porLinea: false },
};
const AVES_PLANTILLA = { pesaje: 50, serologia: 30 };
const etiquetaNave = (s, n) => `${s.Nombre} / ${n.Nombre}`;
const limpiaEnc = s => norm(String(s).replace(/\*/g, ''));

/** Definición de columnas de la plantilla: [{ k, l, t, lista?, min?, max?, entero?, ayuda? }]. */
function columnasImport(tipo, granjaId, linea) {
    const naves = navesActivas(granjaId, linea).map(({ s, n }) => etiquetaNave(s, n));
    const ubic = [{ k: 'Nave', l: 'Nave (sector / nave) *', t: 'text', lista: naves, ayuda: 'Elija la nave de la lista.' },
        { k: 'Fecha', l: 'Fecha *', t: 'date', ayuda: 'Día del dato (dd/mm/aaaa).' }];
    const unidades = state.db.unidades || ['g', 'kg', 'lb'];
    if (tipo === 'registro') {
        const listaDe = t => state.db.listas.filter(x => x.Tipo === t).map(x => x.Valor);
        return ubic.concat(camposEntrada(linea).filter(c => !c.peso).map((c) => {
            const auto = c.k === 'NumeroAves' && linea !== 'Incubadora' ? 'Déjelo vacío si el lote tiene aves iniciales: se calcula solo.'
                : c.k === 'EdadSemanas' ? 'Déjelo vacío si el lote tiene edad inicial: se calcula sola.' : '';
            return { k: c.k, l: c.l, t: c.t === 'entero' || c.t === 'numero' ? 'num' : c.t === 'fecha' ? 'date' : c.t === 'textoLargo' ? 'wrap' : 'text',
                entero: c.t === 'entero', min: c.min, max: c.max, lista: c.lista ? listaDe(c.lista) : null, ayuda: auto };
        }));
    }
    if (tipo === 'pesaje') {
        return ubic.concat([{ k: 'Unidad', l: 'Unidad de los pesos *', t: 'text', lista: unidades, ayuda: 'g, kg o lb.' }],
            Array.from({ length: AVES_PLANTILLA.pesaje }, (_, i) => ({ k: `Ave${i + 1}`, l: `Ave ${i + 1}`, t: 'num', min: 0, max: 100000, ave: true })));
    }
    if (tipo === 'lectura') {
        return ubic.concat([{ k: 'Hora', l: 'Hora (hh:mm) *', t: 'text', ayuda: 'Hora de la lectura, por ejemplo 14:30.' },
            { k: 'Temperatura', l: 'Temperatura (°C) *', t: 'num', min: -20, max: 60 }, { k: 'Humedad', l: 'Humedad relativa (%) *', t: 'num', min: 0, max: 100 }]);
    }
    if (tipo === 'serologia') {
        const metodos = state.db.metodosSerologia.filter(m => m.Estado !== 'Eliminado' && (!m.GranjaID || m.GranjaID === granjaId))
            .map(m => `${enfermedadById(m.EnfermedadID)?.Nombre || ''} — ${nombreMetodo(m)}`).sort((a, b) => a.localeCompare(b, 'es'));
        return ubic.concat([{ k: 'Metodo', l: 'Enfermedad y método *', t: 'text', lista: metodos, ayuda: 'Elija de la lista (los métodos de su granja y los generales).' },
            { k: 'Laboratorio', l: 'Laboratorio', t: 'text' }].concat(linea === 'Incubadora' ? [{ k: 'LoteOrigen', l: 'Lote de origen (reproductora)', t: 'text' }] : [],
            [{ k: 'Comentarios', l: 'Comentarios', t: 'wrap' }],
            Array.from({ length: AVES_PLANTILLA.serologia }, (_, i) => ({ k: `Ave${i + 1}`, l: `Ave ${i + 1}`, t: 'text', ave: true }))));
    }
    // compra
    const granjas = varias() ? state.db.granjas.filter(g => g.Estado === 'Activo').map(g => g.NombreGranja) : [];
    return (varias() ? [{ k: 'Granja', l: 'Granja *', t: 'text', lista: granjas }] : []).concat([
        { k: 'Fase', l: 'Fase de alimento *', t: 'text', ayuda: 'El nombre que usa su granja (Inicio, Crecimiento…).' },
        { k: 'Proveedor', l: 'Empresa de alimento', t: 'text' },
        { k: 'Desde', l: 'Cubre desde *', t: 'date' }, { k: 'Hasta', l: 'Cubre hasta *', t: 'date' },
        { k: 'CantidadQuintales', l: 'Cantidad (quintales) *', t: 'num', min: 0, max: 10000000 },
        { k: 'CostoPorQuintal', l: 'Costo por quintal (Q) *', t: 'num', min: 0, max: 1000000 },
        { k: 'Comentarios', l: 'Comentarios', t: 'wrap' }]);
}

function descargarPlantilla(tipo, granjaId, linea) {
    const cols = columnasImport(tipo, granjaId, linea);
    const conLista = cols.filter(c => c.lista && c.lista.length);
    const largo = Math.max(1, ...conLista.map(c => c.lista.length));
    const listas = { nombre: 'Listas', simple: true, oculta: true, columnas: conLista.map(c => ({ l: c.l, t: 'text' })),
        filas: Array.from({ length: largo }, (_, i) => conLista.map(c => c.lista[i] ?? '')) };
    const hoy = today(), desde = '2000-01-01';
    const validaciones = [];
    cols.forEach((c, j) => {
        const ayuda = c.ayuda || '';
        if (c.lista && c.lista.length) {
            const col = String.fromCharCode(65 + conLista.indexOf(c));
            validaciones.push({ col: j, tipo: 'lista', lista: `Listas!$${col}$1:$${col}$${c.lista.length}`, ayuda, ayudaTitulo: c.l.replace(' *', ''), error: 'Elija un valor de la lista.' });
        } else if (c.t === 'date') {
            validaciones.push({ col: j, tipo: 'fecha', min: desde, max: hoy, ayuda, ayudaTitulo: 'Fecha',
                error: 'Escriba una fecha que no sea futura.' });
        } else if (c.t === 'num' && c.min !== undefined && c.max !== undefined) {
            validaciones.push({ col: j, tipo: c.entero ? 'entero' : 'decimal', min: c.min, max: c.max, ayuda, ayudaTitulo: c.l.replace(' *', ''), error: `Escriba un número entre ${c.min} y ${c.max}${c.entero ? ', sin decimales' : ''}.` });
        }
    });
    const vals = validaciones;
    const def = IMPORT_TIPOS[tipo];
    const g = granjaById(granjaId);
    const instrucciones = [
        `Plantilla de CHIIK: ${def.titulo(linea)}${tipo === 'compra' && varias() ? '' : ` · ${g?.NombreGranja || ''}`}.`,
        'Escriba un dato por fila, desde la fila de abajo de los encabezados. Las columnas con * son obligatorias.',
        'No cambie los encabezados ni el orden de las columnas. Puede dejar filas vacías: se ignoran.',
        tipo === 'pesaje' ? 'Escriba el peso de cada ave en una columna (Ave 1, Ave 2…). Si la muestra es de más aves, agregue columnas "Ave 51", "Ave 52"…' : '',
        tipo === 'serologia' ? 'Escriba el título de cada ave en una columna (Ave 1, Ave 2…): HI en log2 (0 a 16), ELISA el título final, placa P o N.' : '',
        'Puede cargar cualquier fecha pasada (por ejemplo, el historial de un lote); no se aceptan fechas futuras.',
        'Al subir el archivo, CHIIK revisa cada fila y le muestra qué se guardará antes de guardar. Los datos que ya existen se saltan.',
    ].filter(Boolean);
    const blob = crearXlsx([
        { nombre: 'Datos', activa: true, titulo: [def.titulo(linea), `${g && tipo !== 'compra' ? g.NombreGranja + ' · ' : ''}Plantilla para cargar en CHIIK · columnas con * obligatorias`],
            columnas: cols.map(c => ({ l: c.l, t: c.t === 'num' ? 'num' : c.t === 'date' ? 'date' : c.t === 'wrap' ? 'wrap' : 'text', ancho: c.ave ? 8 : undefined })),
            filas: [], filasVacias: 200, validaciones: vals },
        { nombre: 'Instrucciones', simple: true, columnas: [{ l: 'Instrucciones', t: 'text', ancho: 110 }], filas: instrucciones.map(x => [x]) },
        listas,
    ]);
    downloadBlob(blob, `plantilla-${slug(def.titulo(linea))}${g && tipo !== 'compra' ? '-' + slug(g.NombreGranja) : ''}.xlsx`);
}

/** Fecha escrita en Excel → 'AAAA-MM-DD' ('' si está vacía; null si no se entiende). */
function fechaImport(v) {
    if (v === '' || v === null || v === undefined) return '';
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
    if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    return null;
}

/**
 * Lee el archivo y convierte cada fila con datos al formato del formulario.
 * → { filas: [{ excel: n° de fila, payload | error, etiqueta }] }
 */
async function leerImport(archivo, tipo, granjaId, linea) {
    const { hojas } = await leerXlsx(archivo);
    const cols = columnasImport(tipo, granjaId, linea);
    const hoja = hojas.find(h => h.nombre === 'Datos') || hojas[0];
    if (!hoja) throw new Error('El archivo no tiene hojas.');
    const primera = limpiaEnc(cols[0].l);
    const iEnc = hoja.filas.findIndex(f => f && f.some(v => limpiaEnc(v) === primera));
    if (iEnc < 0) throw new Error(`No se encontró la fila de encabezados ("${cols[0].l.replace(' *', '')}"). Use la plantilla de este tipo de dato.`);
    const enc = hoja.filas[iEnc].map(limpiaEnc);
    const idx = {};
    cols.forEach((c) => { const j = enc.indexOf(limpiaEnc(c.l)); if (j >= 0) idx[c.k] = j; });
    const faltan = cols.filter(c => !c.ave && idx[c.k] === undefined && c.l.includes('*'));
    if (faltan.length) throw new Error(`Faltan columnas en el archivo: ${faltan.map(c => c.l.replace(' *', '')).join(', ')}. Use la plantilla de este tipo de dato.`);
    const colsAve = enc.map((h, j) => ({ h, j })).filter(x => /^ave \d+$/.test(x.h)).sort((a, b) => Number(a.h.slice(4)) - Number(b.h.slice(4)));
    const naves = new Map(navesActivas(granjaId, linea).map(({ s, n }) => [norm(etiquetaNave(s, n)), n]));
    const metodos = new Map(state.db.metodosSerologia.filter(m => !m.GranjaID || m.GranjaID === granjaId)
        .map(m => [norm(`${enfermedadById(m.EnfermedadID)?.Nombre || ''} — ${nombreMetodo(m)}`), m]));
    const granjasPorNombre = new Map(state.db.granjas.map(g => [norm(g.NombreGranja), g]));
    const out = [];
    hoja.filas.slice(iEnc + 1).forEach((f, i) => {
        if (!f || f.every(v => v === '' || v === null || v === undefined)) return;
        const excel = iEnc + i + 2;
        const val = k => (idx[k] === undefined ? '' : f[idx[k]] ?? '');
        const r = { excel, etiqueta: '' };
        try {
            const p = {};
            if (tipo !== 'compra') {
                const nave = naves.get(norm(val('Nave')));
                if (!nave) throw new Error(val('Nave') ? `La nave "${val('Nave')}" no existe o está desactivada.` : 'Falta la nave.');
                const fecha = fechaImport(val('Fecha'));
                if (fecha === null) throw new Error(`Fecha no válida: "${val('Fecha')}".`);
                Object.assign(p, { Linea: linea, NaveID: nave.NaveID, Fecha: fecha });
                r.etiqueta = `${val('Nave')} · ${fecha ? fmtDate(fecha) : 'sin fecha'}`;
            }
            if (tipo === 'registro') {
                cols.slice(2).forEach((c) => { const v = val(c.k); if (v !== '') p[c.k] = c.t === 'date' ? fechaImport(v) : v; });
            } else if (tipo === 'pesaje' || tipo === 'serologia') {
                const valores = colsAve.map(x => f[x.j]).filter(v => v !== '' && v !== null && v !== undefined);
                if (!valores.length) throw new Error('No trae los valores de las aves.');
                p.Muestra = valores.length;
                if (tipo === 'pesaje') {
                    Object.assign(p, { Unidad: String(val('Unidad')).trim().toLowerCase(), Pesos: valores.map(Number) });
                } else {
                    const m = metodos.get(norm(val('Metodo')));
                    if (!m) throw new Error(val('Metodo') ? `Método no encontrado: "${val('Metodo')}".` : 'Falta la enfermedad y el método.');
                    Object.assign(p, { EnfermedadID: m.EnfermedadID, MetodoID: m.MetodoID, Valores: valores.map(v => (m.Prueba === 'PLACA' ? String(v).trim().toUpperCase() : v)),
                        Laboratorio: val('Laboratorio'), Comentarios: val('Comentarios') });
                    if (linea === 'Incubadora') p.LoteOrigen = val('LoteOrigen');
                }
            } else if (tipo === 'lectura') {
                const h = String(val('Hora')).trim().match(/^(\d{1,2}):(\d{2})/);
                if (!h) throw new Error(`Hora no válida: "${val('Hora')}" (use hh:mm).`);
                Object.assign(p, { FechaHora: `${p.Fecha} ${h[1].padStart(2, '0')}:${h[2]}`, Temperatura: val('Temperatura'), Humedad: val('Humedad') });
                r.etiqueta += ` ${h[1].padStart(2, '0')}:${h[2]}`;
                delete p.Fecha; delete p.Linea;
            } else {
                const g = varias() ? granjasPorNombre.get(norm(val('Granja'))) : granjaById(granjaId);
                if (!g) throw new Error(val('Granja') ? `La granja "${val('Granja')}" no existe.` : 'Falta la granja.');
                const desde = fechaImport(val('Desde')), hasta = fechaImport(val('Hasta'));
                if (desde === null || hasta === null) throw new Error('Fechas "desde" o "hasta" no válidas.');
                Object.assign(p, { GranjaID: g.GranjaID, Fase: val('Fase'), Proveedor: val('Proveedor'), Desde: desde, Hasta: hasta,
                    CantidadQuintales: val('CantidadQuintales'), CostoPorQuintal: val('CostoPorQuintal'), Comentarios: val('Comentarios') });
                r.etiqueta = `${val('Fase')} · ${fmtDate(desde)} a ${fmtDate(hasta)}`;
            }
            r.payload = p;
        } catch (e) {
            r.error = e.message;
        }
        out.push(r);
    });
    // Repetidas dentro del mismo archivo (misma nave y fecha): solo cuenta la primera.
    if (tipo === 'registro' || tipo === 'pesaje') {
        const vistas = new Set();
        out.forEach((r) => {
            if (!r.payload) return;
            const k = `${r.payload.NaveID}|${r.payload.Fecha}`;
            if (vistas.has(k)) { r.error = 'Repetida en el archivo (misma nave y fecha que una fila anterior).'; r.repetida = true; delete r.payload; }
            vistas.add(k);
        });
    }
    return out;
}

/**
 * El servidor revisa o guarda hasta 300 filas por vez: con más, se mandan por tandas (una tras otra)
 * y se juntan los resultados. `avance(hechas, total)` permite mostrar el progreso.
 */
const TANDA_IMPORTAR = 300;
async function importarPorTandas(tipo, payloads, archivo, confirmar, avance) {
    const out = { resultados: [], listos: 0, guardados: 0, duplicados: 0, errores: 0 };
    for (let i = 0; i < payloads.length; i += TANDA_IMPORTAR) {
        const parte = payloads.slice(i, i + TANDA_IMPORTAR);
        const res = await withBusy(() => api('importarDatos', { tipo, filas: parte, archivo, confirmar }));
        res.resultados.forEach(r => out.resultados.push({ ...r, fila: r.fila + i }));
        ['listos', 'guardados', 'duplicados', 'errores'].forEach((k) => { out[k] += res[k] || 0; });
        avance?.(Math.min(i + TANDA_IMPORTAR, payloads.length), payloads.length);
    }
    return out;
}

async function openImportarModal(tipo) {
    const def = IMPORT_TIPOS[tipo];
    const linea = state.linea;
    const granjas = varias() ? state.db.granjas.filter(g => g.Estado === 'Activo') : state.db.granjas.filter(g => g.GranjaID === state.granjaId);
    if (!granjas.length) return toast('No hay granjas activas donde cargar datos.', 'error');
    const pre = varias() && state.f.granjas.length === 1 ? state.f.granjas[0] : (state.granjaId || granjas[0].GranjaID);
    let filas = null, archivoNombre = '', revision = null;
    const granjaSel = tipo === 'compra' && varias() ? '' : (granjas.length > 1
        ? `<label class="field"><span>Granja</span><select name="GranjaID">${granjas.map(g => `<option value="${esc(g.GranjaID)}"${g.GranjaID === pre ? ' selected' : ''}>${esc(g.NombreGranja)}</option>`).join('')}</select></label>`
        : `<input type="hidden" name="GranjaID" value="${esc(granjas[0].GranjaID)}">`);
    const body = `<div class="import-steps">
        <section class="form-section"><h3>${ico('download')}1. Descargue la plantilla</h3>
            <p class="muted">${esc(def.titulo(linea))}. Trae las naves de la granja en listas desplegables y revisa los valores mientras escribe.</p>
            <div class="form-grid">${granjaSel}</div>
            <div class="row row-center"><button type="button" class="btn" data-imp-plantilla>${ico('file-spreadsheet')}Descargar plantilla</button></div></section>
        <section class="form-section"><h3>${ico('upload')}2. Suba el archivo lleno</h3>
            <label class="field"><span>Archivo de Excel (.xlsx)</span><input type="file" name="archivo" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" data-imp-archivo></label>
            <p class="muted">Desde Excel puede cargar cualquier fecha pasada (el historial); no se aceptan fechas futuras.</p></section>
        <section class="form-section" data-imp-revision hidden></section></div>`;
    const granjaDe = form => form.GranjaID?.value || pre;
    const promesa = openModal({
        title: 'Cargar desde Excel', wide: true, body, submitLabel: 'Revisar archivo',
        onSubmit: async (fd, form) => {
            const btn = form.querySelector('[type="submit"]');
            if (revision) {
                // 2.º paso: guardar las filas listas.
                const listas = revision.filter(r => r.estado === 'listo');
                const res = await importarPorTandas(tipo, listas.map(r => r.payload), archivoNombre, true, (h, t) => { btn.textContent = `Guardando ${h} de ${t}…`; });
                state.cache.clear();
                render();
                toast(`Se guardaron ${nf.format(res.guardados)} filas${res.duplicados ? `; ${res.duplicados} ya existían y se saltaron` : ''}${res.errores ? `; ${res.errores} con error` : ''}.`, res.errores ? 'info' : 'success', 8000);
                return true;
            }
            if (!filas) throw new Error('Elija el archivo lleno.');
            const validas = filas.filter(r => r.payload);
            let res = { resultados: [] };
            if (validas.length) res = await importarPorTandas(tipo, validas.map(r => r.payload), archivoNombre, false, (h, t) => { btn.textContent = `Revisando ${h} de ${t}…`; });
            validas.forEach((r, i) => Object.assign(r, res.resultados[i]));
            revision = filas.map(r => (r.payload ? r : { ...r, estado: r.repetida ? 'duplicado' : 'error', mensaje: r.error }));
            const n = e => revision.filter(r => r.estado === e).length;
            const box = form.querySelector('[data-imp-revision]');
            box.hidden = false;
            box.innerHTML = `<h3>${ico('list-checks')}3. Revise antes de guardar</h3>
                <div class="chips"><span class="badge badge-ok">${n('listo')} listas</span><span class="badge badge-warn">${n('duplicado')} repetidas (se saltan)</span><span class="badge badge-danger">${n('error')} con error</span></div>
                <div class="table-wrap import-tabla"><table><thead><tr><th>Fila</th><th>Dato</th><th>Estado</th><th>Motivo</th></tr></thead><tbody>
                ${revision.map(r => `<tr><td>${r.excel}</td><td>${esc(r.etiqueta || '—')}</td><td>${r.estado === 'listo' ? '<span class="badge badge-ok">Lista</span>' : r.estado === 'duplicado' ? '<span class="badge badge-warn">Repetida</span>' : '<span class="badge badge-danger">Error</span>'}</td><td>${esc(r.mensaje || '')}</td></tr>`).join('')}
                </tbody></table></div>
                ${n('listo') ? '' : '<p class="muted">No hay filas para guardar. Corrija el archivo y vuelva a subirlo.</p>'}`;
            btn.textContent = n('listo') ? `Guardar ${n('listo')} ${n('listo') === 1 ? 'fila' : 'filas'}` : 'Revisar archivo';
            if (!n('listo')) revision = null;
            box.scrollIntoView({ behavior: 'smooth', block: 'start' });
            return MANTENER_ABIERTO;
        },
    });
    const form = $('#modal form');
    const reiniciar = () => {
        revision = null;
        const box = form.querySelector('[data-imp-revision]');
        box.hidden = true; box.innerHTML = '';
        form.querySelector('[type="submit"]').textContent = 'Revisar archivo';
    };
    form.querySelector('[data-imp-plantilla]').addEventListener('click', () => descargarPlantilla(tipo, granjaDe(form), linea));
    form.querySelector('[data-imp-archivo]').addEventListener('change', async (e) => {
        reiniciar();
        filas = null;
        const f = e.target.files[0];
        if (!f) return;
        archivoNombre = f.name;
        const err = form.querySelector('.form-error');
        try {
            filas = await leerImport(f, tipo, granjaDe(form), linea);
            if (!filas.length) throw new Error('El archivo no trae filas con datos.');
            err.hidden = true;
        } catch (ex) {
            filas = null;
            err.textContent = ex.message; err.hidden = false;
        }
    });
    form.GranjaID?.addEventListener?.('change', () => { reiniciar(); filas = null; form.archivo.value = ''; });
    return promesa;
}

// ---------- Carga del formato semanal del cliente (registros de postura) ----------
// Su libro de Excel con una hoja por semana ("SEM 01"…), un bloque por día y una fila por galera.
// La app lee los datos (web/formato-semanal.js), empareja cada galera con una nave, y los carga
// por nave (cada galera) o por lote (la suma de sus galeras), con la misma revisión que la plantilla.

/** Lote de la app con ese número que estaba activo en la fecha (en la granja, línea de postura). */
function loteDeNumero(granjaId, numero, fecha) {
    const delNumero = state.db.lotes.filter(l => l.GranjaID === granjaId && l.Linea === 'Postura' && String(l.NumeroLote).trim() === String(numero).trim());
    return delNumero.find(l => l.FechaInicio <= fecha && (!l.FechaCierre || fecha <= l.FechaCierre)) || null;
}

/**
 * Filas para revisar: { etiqueta, payload | error, aviso? }. Por nave, cada galera va a su nave; si el
 * lote de la app se registra por lote ese día, sus galeras se suman. Por lote, se suman siempre.
 */
function filasFormatoSemanal(leido, granjaId, mapa, forma) {
    const out = [];
    const delDia = x => ({ MortalidadNominal: x.muertas, OtrasSalidas: x.seleccion, ConsumoQuintalesTotal: x.consumoQq ?? '', PosturaNominal: x.huevos ?? '' });
    const avisoEdad = (l, x) => {
        const d = edadDias(l, x.fecha);
        return d !== null && x.edad !== null && Math.abs(Math.floor(d / 7) - x.edad) > 1 ? `Edad en el Excel: ${x.edad} semanas; en la app: ${Math.floor(d / 7)}.` : '';
    };
    const juntar = [];
    leido.dias.forEach((x) => {
        const et = `${x.hoja.trim()} · ${fmtDate(x.fecha)} · galera ${x.galera} · lote ${x.lote}`;
        const l = loteDeNumero(granjaId, x.lote, x.fecha);
        if (!l) { out.push({ etiqueta: et, error: `El lote ${x.lote} no existe en la granja o no estaba activo ese día: créelo antes.` }); return; }
        if (forma === 'lote' || modoEn(l, x.fecha) === 'Lote') { juntar.push(x); return; }
        if (!mapa[x.galera]) { out.push({ etiqueta: et, error: `La galera ${x.galera} no está emparejada con una nave.` }); return; }
        out.push({ etiqueta: et, payload: { Linea: 'Postura', NaveID: mapa[x.galera], Fecha: x.fecha, ...delDia(x) }, aviso: avisoEdad(l, x) });
    });
    agruparPorLote(juntar).forEach((g) => {
        const l = loteDeNumero(granjaId, g.lote, g.fecha);
        const et = `${g.hoja.trim()} · ${fmtDate(g.fecha)} · lote ${g.lote} (galeras ${g.galeras.join(', ')})`;
        if (modoEn(l, g.fecha) !== 'Lote') { out.push({ etiqueta: et, error: `El lote ${g.lote} se registra por nave ese día: cárguelo "por nave".` }); return; }
        out.push({ etiqueta: et, payload: { Linea: 'Postura', NaveID: '', LoteID: l.LoteID, Fecha: g.fecha, ...delDia(g) }, aviso: avisoEdad(l, g) });
    });
    return out;
}

/**
 * Galeras del Excel que su lote de la app aún no tiene (por ejemplo, un lote que en la app está en una
 * sola nave): "Agregar la galera 6 al lote 52 desde el 29/12/2025 con 6,502 aves".
 */
function propuestasNaves(leido, granjaId, mapa) {
    const vistas = new Set();
    const out = [];
    leido.dias.forEach((x) => {
        const k = `${x.lote}|${x.galera}`;
        if (vistas.has(k)) return;
        vistas.add(k);
        const l = loteDeNumero(granjaId, x.lote, x.fecha);
        const naveId = mapa[x.galera];
        if (!l || !naveId || modoEn(l, x.fecha) === 'Lote') return;
        if (navesDelLote(l).some(n => n.NaveID === naveId && n.Desde <= x.fecha && (!n.Hasta || x.fecha < n.Hasta))) return;
        const ocupada = loteEn(naveId, x.fecha);
        out.push({ LoteID: l.LoteID, NaveID: naveId, Desde: x.fecha, AvesIniciales: x.avesAyer, lote: x.lote, galera: x.galera, ocupada: ocupada && ocupada.LoteID !== l.LoteID ? ocupada.NumeroLote : '' });
    });
    return out;
}

async function openFormatoSemanalModal() {
    const linea = 'Postura';
    const granjas = (varias() ? state.db.granjas.filter(g => g.Estado === 'Activo') : state.db.granjas.filter(g => g.GranjaID === state.granjaId))
        .filter(g => perfilDe(g.GranjaID, linea)?.Estado === 'Activo');
    if (!granjas.length) return toast('No hay granjas con línea de postura donde cargar datos.', 'error');
    const pre = varias() && state.f.granjas.length === 1 && granjas.some(g => g.GranjaID === state.f.granjas[0]) ? state.f.granjas[0] : (state.granjaId || granjas[0].GranjaID);
    let leido = null, archivoNombre = '', revision = null, compras = [];
    const granjaSel = granjas.length > 1
        ? `<label class="field"><span>Granja</span><select name="GranjaID">${granjas.map(g => `<option value="${esc(g.GranjaID)}"${g.GranjaID === pre ? ' selected' : ''}>${esc(g.NombreGranja)}</option>`).join('')}</select></label>`
        : `<input type="hidden" name="GranjaID" value="${esc(granjas[0].GranjaID)}">`;
    const body = `<div class="import-steps">
        <section class="form-section"><h3>${ico('upload')}1. Suba su archivo semanal</h3>
            <p class="muted">Su libro de producción con una hoja por semana ("SEM 01", "SEM 02"…): la app lee cada día y cada galera (aves, muertas, selección, consumo en quintales y huevos). Docenas, porcentajes y conversión los calcula la app.</p>
            <div class="form-grid">${granjaSel}<label class="field"><span>Archivo de Excel (.xlsx)</span><input type="file" name="archivo" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" data-sem-archivo></label></div>
            <p class="muted">Puede cargar cualquier fecha pasada. Los días que ya están en la app se saltan.</p></section>
        <section class="form-section" data-sem-config hidden></section>
        <section class="form-section" data-sem-revision hidden></section></div>`;
    const promesa = openModal({
        title: 'Cargar mi formato semanal', wide: true, body, submitLabel: 'Revisar',
        onSubmit: async (fd, form) => {
            const btn = form.querySelector('[type="submit"]');
            const granjaId = form.GranjaID.value;
            if (revision) {
                const listas = revision.filter(r => r.estado === 'listo');
                const res = await importarPorTandas('registro', listas.map(r => r.payload), archivoNombre, true, (h, t) => { btn.textContent = `Guardando ${h} de ${t}…`; });
                let resC = null;
                if (form.querySelector('[name="compras"]')?.checked && compras.length) resC = await importarPorTandas('compra', compras, archivoNombre, true);
                state.cache.clear();
                await withBusy(() => loadSnapshot());
                render();
                toast(`Se guardaron ${nf.format(res.guardados)} registros${res.duplicados ? `; ${nf.format(res.duplicados)} ya existían y se saltaron` : ''}${res.errores ? `; ${res.errores} con error` : ''}` +
                    (resC ? `. Compras de concentrado: ${resC.guardados} nuevas${resC.duplicados ? `, ${resC.duplicados} ya estaban` : ''}.` : '.'), res.errores ? 'info' : 'success', 9000);
                return true;
            }
            if (!leido) throw new Error('Elija su archivo semanal.');
            const mapa = {};
            form.querySelectorAll('[data-galera]').forEach((s) => { if (s.value) mapa[s.dataset.galera] = s.value; });
            // Primero, las naves marcadas que se agregan a sus lotes (para que los días encuentren su lote).
            const marcadas = [...form.querySelectorAll('[name="propuesta"]:checked')].map(c => JSON.parse(c.value));
            for (const p of marcadas) {
                await withBusy(() => api('agregarNaveLote', { LoteID: p.LoteID, NaveID: p.NaveID, AvesIniciales: p.AvesIniciales, Desde: p.Desde, justificacion: `Formato semanal (${archivoNombre})` }));
            }
            if (marcadas.length) { await withBusy(() => loadSnapshot()); toast(`${marcadas.length} ${marcadas.length === 1 ? 'nave agregada' : 'naves agregadas'} a sus lotes.`, 'success'); }
            const forma = form.querySelector('[name="forma"]:checked')?.value || 'nave';
            const filas = filasFormatoSemanal(leido, granjaId, mapa, forma);
            const validas = filas.filter(r => r.payload);
            let res = { resultados: [] };
            if (validas.length) res = await importarPorTandas('registro', validas.map(r => r.payload), archivoNombre, false, (h, t) => { btn.textContent = `Revisando ${h} de ${t}…`; });
            validas.forEach((r, i) => Object.assign(r, res.resultados[i]));
            revision = filas.map(r => (r.payload ? r : { ...r, estado: 'error', mensaje: r.error }));
            compras = leido.concentrado.map(c => ({ GranjaID: granjaId, Fase: c.fase, Desde: c.desde, Hasta: c.hasta, CantidadQuintales: c.quintales, CostoPorQuintal: c.costoQq ?? '', Comentarios: `Formato semanal ${c.hoja.trim()}` }));
            const n = e => revision.filter(r => r.estado === e).length;
            const box = form.querySelector('[data-sem-revision]');
            box.hidden = false;
            // Se muestran los errores y repetidos completos; de las listas, las primeras (pueden ser miles).
            const mostrar = revision.filter(r => r.estado !== 'listo').concat(revision.filter(r => r.estado === 'listo').slice(0, 150));
            box.innerHTML = `<h3>${ico('list-checks')}3. Revise antes de guardar</h3>
                <div class="chips"><span class="badge badge-ok">${nf.format(n('listo'))} listos</span><span class="badge badge-warn">${nf.format(n('duplicado'))} ya estaban (se saltan)</span><span class="badge badge-danger">${nf.format(n('error'))} con error</span></div>
                ${revision.some(r => r.aviso) ? alertBox('info', 'info', `${revision.filter(r => r.aviso).length} filas con la edad distinta a la de la app (se guardan igual; revise la edad al iniciar del lote).`) : ''}
                <div class="table-wrap import-tabla"><table><thead><tr><th>Dato</th><th>Estado</th><th>Motivo</th></tr></thead><tbody>
                ${mostrar.map(r => `<tr><td>${esc(r.etiqueta)}</td><td>${r.estado === 'listo' ? '<span class="badge badge-ok">Listo</span>' : r.estado === 'duplicado' ? '<span class="badge badge-warn">Ya estaba</span>' : '<span class="badge badge-danger">Error</span>'}</td><td>${esc(r.mensaje || r.aviso || '')}</td></tr>`).join('')}
                </tbody></table></div>
                ${n('listo') > 150 ? `<p class="muted center">Se muestran 150 de los ${nf.format(n('listo'))} listos.</p>` : ''}
                ${compras.length ? `<label class="check"><input type="checkbox" name="compras" checked> Cargar también ${compras.length} compras de concentrado (cuadro "ALIMENTO" de cada semana; las que ya estén se saltan)</label>` : ''}`;
            btn.textContent = n('listo') ? `Guardar ${nf.format(n('listo'))} ${n('listo') === 1 ? 'registro' : 'registros'}` : 'Revisar';
            if (!n('listo')) revision = null;
            box.scrollIntoView({ behavior: 'smooth', block: 'start' });
            return MANTENER_ABIERTO;
        },
    });
    const form = $('#modal form');
    const config = form.querySelector('[data-sem-config]');
    const reiniciar = () => {
        revision = null;
        const box = form.querySelector('[data-sem-revision]');
        box.hidden = true; box.innerHTML = '';
        form.querySelector('[type="submit"]').textContent = 'Revisar';
    };
    // Paso 2: emparejar galeras con naves, forma de carga y naves que faltan en sus lotes.
    const pintarConfig = () => {
        const granjaId = form.GranjaID.value;
        const naves = navesActivas(granjaId, linea);
        const galeras = [...new Set(leido.dias.map(x => x.galera))].sort((a, b) => Number(a) - Number(b) || a.localeCompare(b));
        const adivina = g => naves.find(({ n }) => numeroDeNave(n.Nombre) === String(Number(g)))?.n.NaveID || '';
        const semanas = new Set(leido.dias.map(x => x.hoja.trim())).size;
        const fechas = leido.dias.map(x => x.fecha);
        config.hidden = false;
        config.innerHTML = `<h3>${ico('list-checks')}2. Revise cómo se carga</h3>
            <p class="center">${semanas} ${semanas === 1 ? 'semana' : 'semanas'} · ${nf.format(leido.dias.length)} galeras por día · del ${esc(fmtDate(fechas[0]))} al ${esc(fmtDate(fechas[fechas.length - 1]))} · lotes ${esc([...new Set(leido.dias.map(x => x.lote))].join(', '))}</p>
            ${leido.avisos.length ? alertBox('warn', 'alert-triangle', leido.avisos.slice(0, 5).map(a => `${esc(a.hoja.trim())}: ${esc(a.texto)}`).join('<br>')) : ''}
            <div class="form-grid">${galeras.map(g => `<label class="field"><span>Galera ${esc(g)} del Excel</span><select data-galera="${esc(g)}"><option value="">— Sin nave —</option>${naves.map(({ s, n }) =>
                `<option value="${esc(n.NaveID)}"${n.NaveID === adivina(g) ? ' selected' : ''}>${esc(s.Nombre)} / ${esc(n.Nombre)}</option>`).join('')}</select></label>`).join('')}</div>
            <div class="filter-group"><span class="field-label">Cargar</span>
                <label class="check"><input type="radio" name="forma" value="nave" checked> Por nave: cada galera por separado (si el lote se lleva por lote, la app suma sus galeras)</label>
                <label class="check"><input type="radio" name="forma" value="lote"> Por lote: la suma de las galeras de cada lote, sin tomar en cuenta la galera</label></div>
            <div data-sem-propuestas></div>`;
        pintarPropuestas();
    };
    const pintarPropuestas = () => {
        const mapa = {};
        form.querySelectorAll('[data-galera]').forEach((s) => { if (s.value) mapa[s.dataset.galera] = s.value; });
        const props = propuestasNaves(leido, form.GranjaID.value, mapa);
        const faltan = [...new Set(leido.dias.filter(x => !loteDeNumero(form.GranjaID.value, x.lote, x.fecha)).map(x => x.lote))];
        form.querySelector('[data-sem-propuestas]').innerHTML = (faltan.length ? alertBox('warn', 'alert-triangle', `Estos lotes no existen en la granja (o no estaban activos en esas fechas): ${esc(faltan.join(', '))}. Créelos antes (con su fecha de inicio) o esas filas darán error.`) : '')
            + (props.length ? `<p class="field-label">Galeras que su lote aún no tiene en la app</p>${props.map(p => `<label class="check"><input type="checkbox" name="propuesta" value="${esc(JSON.stringify(p))}"${p.ocupada ? ' disabled' : ' checked'}>
                Agregar la galera ${esc(p.galera)} (${esc(naveById(p.NaveID)?.Nombre || '')}) al lote ${esc(p.lote)} desde el ${esc(fmtDate(p.Desde))} con ${esc(fmtNum(p.AvesIniciales))} aves${p.ocupada ? ` — no se puede: esa nave tiene el lote ${esc(p.ocupada)}` : ''}</label>`).join('')}` : '');
    };
    form.addEventListener('change', (e) => {
        if (e.target.matches('[data-galera]')) { reiniciar(); pintarPropuestas(); }
        if (e.target.matches('[name="forma"]')) reiniciar();
    });
    form.querySelector('[data-sem-archivo]').addEventListener('change', async (e) => {
        reiniciar();
        leido = null;
        config.hidden = true; config.innerHTML = '';
        const f = e.target.files[0];
        if (!f) return;
        archivoNombre = f.name;
        const err = form.querySelector('.form-error');
        try {
            leido = leerFormatoSemanal(await leerXlsx(f));
            if (!leido.dias.length) throw new Error(leido.avisos[0]?.texto || 'No se encontraron hojas semanales ("SEM 01"…) con datos.');
            err.hidden = true;
            pintarConfig();
        } catch (ex) {
            leido = null;
            err.textContent = ex.message; err.hidden = false;
        }
    });
    form.GranjaID?.addEventListener?.('change', () => { reiniciar(); if (leido) pintarConfig(); });
    return promesa;
}

// ---------- Serología ----------
const PRUEBA_NOMBRE ={ HI: 'HI (inhibición de la hemaglutinación)', ELISA: 'ELISA', PLACA: 'Aglutinación en placa' };
const enfermedadById = id => state.db.enfermedades.find(e => e.EnfermedadID === id) || null;
const metodoById = id => state.db.metodosSerologia.find(m => m.MetodoID === id) || null;
const nombreMetodo = m => `${m.Nombre.includes(m.Prueba) ? m.Nombre : `${m.Nombre} · ${m.Prueba}`}${m.GranjaID ? '' : ' (general)'}`;
/** Métodos de una enfermedad que sirven en esas granjas: los generales y los propios de ellas. */
function metodosDe(enfId, granjaIds) {
    return state.db.metodosSerologia.filter(m => m.EnfermedadID === enfId && (!m.GranjaID || granjaIds.includes(m.GranjaID))).sort(byNombre);
}

/** Lee los resultados escritos o pegados (separados por espacios, comas, punto y coma o saltos de línea). */
function leerValoresSero(texto, prueba) {
    const partes = String(texto || '').split(/[\s;,]+/).map(s => s.trim()).filter(Boolean);
    const def = state.db.pruebasSerologia[prueba] || {};
    const valores = [], malos = [];
    partes.forEach((s) => {
        if (prueba === 'PLACA') {
            const v = s.toUpperCase();
            if (v === 'P' || v === 'N') valores.push(v); else malos.push(s);
            return;
        }
        const v = Number(s);
        const f = 10 ** (def.decimales || 0);
        if (!Number.isFinite(v) || v < def.min || v > def.max || (def.entero && !Number.isInteger(v)) || (!def.entero && Math.abs(Math.round(v * f) - v * f) > 1e-6)) malos.push(s); else valores.push(v);
    });
    return { valores, malos };
}

/** El mismo resumen que el servidor (ver resumenSerologia_ en el backend). */
function resumenSerologia(prueba, valores) {
    const n = valores.length;
    if (!n) return null;
    if (prueba === 'PLACA') {
        const pos = valores.filter(v => v === 'P').length;
        return { Positivos: pos, PctPositivos: pos / n * 100 };
    }
    const suma = valores.reduce((a, b) => a + b, 0);
    const o = { Maximo: Math.max(...valores), Minimo: Math.min(...valores) };
    if (prueba === 'HI') return { ...o, Promedio: suma / n };
    const media = suma / n;
    const sd = n > 1 ? Math.sqrt(valores.reduce((a, v) => a + (v - media) ** 2, 0) / (n - 1)) : 0;
    return { ...o, MediaGeometrica: suma === 0 ? 0 : 10 ** (valores.reduce((a, v) => a + Math.log10(Math.max(v, 1)), 0) / n), CV: media ? sd / media * 100 : 0 };
}

function resumenSeroHtml(prueba, valores, metodo) {
    const st = resumenSerologia(prueba, valores);
    if (!st) return '<p class="muted center">Escribe o pega los resultados para ver el resumen.</p>';
    const filas = prueba === 'PLACA'
        ? [['Positivas', `${st.Positivos} de ${valores.length}`], ['Positivas (%)', `${fmtNum(st.PctPositivos)} %`]]
        : [[prueba === 'HI' ? 'Promedio (log2)' : 'Media geométrica', fmtNum(prueba === 'HI' ? st.Promedio : st.MediaGeometrica)],
            ['Dato mayor', fmtNum(st.Maximo)], ['Dato menor', fmtNum(st.Minimo)]].concat(prueba === 'ELISA' ? [['CV', `${fmtNum(st.CV)} %`]] : []);
    if (prueba !== 'PLACA' && metodo?.CortePositivo !== '' && metodo?.CortePositivo !== undefined) {
        const pos = valores.filter(v => v >= Number(metodo.CortePositivo)).length;
        filas.push(['Positivas', `${pos} de ${valores.length} (${fmtNum(pos / valores.length * 100)} %)`]);
    }
    return filas.map(([l, v]) => `<div><dt>${esc(l)}</dt><dd>${esc(v)}</dd></div>`).join('');
}

/** Opciones de método para la enfermedad y la granja de la nave elegida en el formulario. */
function srMetodosHtml(enfId, granjaId, elegido) {
    const lista = enfId ? metodosDe(enfId, granjaId ? [granjaId] : []) : [];
    return lista.length
        ? lista.map(m => `<option value="${esc(m.MetodoID)}"${m.MetodoID === elegido ? ' selected' : ''}>${esc(nombreMetodo(m))}</option>`).join('')
        : '<option value="">Sin métodos para esta enfermedad</option>';
}

/** Recalcula el formulario de muestreo: métodos de la enfermedad, prueba, conteo y resumen. */
function srRender(form, { metodos = false } = {}) {
    if (!form?.querySelector('[data-sr]')) return;
    const naveId = form.NaveID?.value || form.dataset.nave;
    const granjaId = naveById(naveId)?.GranjaID || '';
    if (metodos && form.MetodoID?.tagName === 'SELECT') form.MetodoID.innerHTML = srMetodosHtml(form.EnfermedadID.value, granjaId, form.MetodoID.value);
    const metodo = metodoById(form.MetodoID?.value);
    const sin = form.querySelector('[data-sr-sin-metodo]');
    if (sin) sin.hidden = !!metodo;
    const prueba = metodo?.Prueba || '';
    const def = state.db.pruebasSerologia[prueba];
    form.querySelector('[data-sr-prueba]').innerHTML = metodo
        ? `Prueba: <strong>${esc(PRUEBA_NOMBRE[prueba])}</strong> · ${prueba === 'PLACA' ? 'escribe P (positivo) o N (negativo) por ave' : `título en ${esc(def.unidad)}${def.entero ? ', números enteros' : ''} de ${esc(fmtNum(def.min))} a ${esc(fmtNum(def.max))}`}`
            + (metodo.CortePositivo !== '' && prueba !== 'PLACA' ? ` · Positivo desde ${esc(fmtNum(metodo.CortePositivo))}` : '')
        : 'Elige un método para saber qué escribir.';
    const { valores, malos } = prueba ? leerValoresSero(form.Valores.value, prueba) : { valores: [], malos: [] };
    const total = Math.max(0, Math.floor(toNum(form.Muestra.value) || 0));
    form.querySelector('[data-sr-count]').textContent = `${nf.format(valores.length)} de ${nf.format(total)}`;
    form.querySelector('[data-sr-aviso]').innerHTML = malos.length
        ? alertBox('warn', 'alert-triangle', `No válidos: ${esc(malos.slice(0, 8).join(' '))}${malos.length > 8 ? '…' : ''}`)
        : valores.length > total ? alertBox('warn', 'alert-triangle', `Hay ${valores.length - total} resultado(s) de más para una muestra de ${total}.`) : '';
    form.querySelector('[data-sr-resumen]').innerHTML = prueba ? resumenSeroHtml(prueba, valores, metodo) : '';
}

/**
 * Nuevo muestreo o edición. Al crear, la casilla "agregar otra enfermedad" reabre el formulario
 * con la misma nave, fecha y muestra (el laboratorio analiza los mismos sueros para varias).
 */
async function openSerologiaModal(s = null, preset = {}) {
    const linea = s ? s.Linea : state.linea;
    let ubic;
    if (s) {
        ubic = `<label class="field"><span>Sector / nave</span><input value="${esc([s.Sector, s.Nave].join(' / '))}" readonly></label>
            <label class="field"><span>Fecha del muestreo</span><input value="${esc(fmtDate(s.Fecha))}" readonly></label>
            <label class="field"><span>Enfermedad</span><input value="${esc(s.Enfermedad)}" readonly></label>
            <label class="field"><span>Método</span><input value="${esc(s.Metodo)} · ${esc(s.Prueba)}" readonly></label>
            <input type="hidden" name="MetodoID" value="${esc(s.MetodoID)}">`;
    } else {
        ubic = ubicacionInputs(linea, { naveId: preset.naveId, fecha: preset.fecha || today(), sinVentana: true });
        if (!ubic) return toast(`No hay naves activas de ${linea.toLowerCase()}.`, 'error');
    }
    const enfInicial = preset.enfId || state.srEnf || state.db.enfermedades[0]?.EnfermedadID || '';
    const naveInicial = s ? s.NaveID : (/name="NaveID"[^>]*>[\s\S]*?<option value="([^"]+)" selected/.exec(ubic) || [])[1];
    const granjaInicial = naveById(naveInicial)?.GranjaID || '';
    const valores = s ? String(s.Valores || '').split(',').filter(Boolean).join(' ') : '';
    const body = `<section class="form-section"><div class="form-grid">${ubic}
        ${s ? '' : `<label class="field"><span>Enfermedad <span class="req">*</span></span><select name="EnfermedadID" required data-change="sr-enf">${state.db.enfermedades.map(e =>
        `<option value="${esc(e.EnfermedadID)}"${e.EnfermedadID === enfInicial ? ' selected' : ''}>${esc(e.Nombre)}</option>`).join('')}</select></label>
            <label class="field"><span>Método o kit <span class="req">*</span></span><select name="MetodoID" required data-change="sr-metodo">${srMetodosHtml(enfInicial, granjaInicial, preset.metodoId || state.srMetodo)}</select></label>
            <div class="span-all" data-sr-sin-metodo hidden><div class="alert alert-info">${ico('info')}<span>Esta enfermedad aún no tiene un método para esta granja.</span><button type="button" class="btn btn-sm" data-action="metodo-desde-form">${ico('plus')}Crear método</button></div></div>`}
    </div></section>
    <section class="form-section" data-sr><h3>${ico('syringe')}Resultados por ave</h3>
        <p class="field-hint center" data-sr-prueba></p>
        <div class="form-grid">
            <label class="field"><span>Tamaño de la muestra <span class="req">*</span></span><input type="number" name="Muestra" min="1" max="${state.db.maxMuestra}" step="1" inputmode="numeric" required value="${esc(s?.Muestra || preset.muestra || 20)}" data-input="sr-calc">
                <small class="field-hint">Número de aves (sueros) analizadas.</small></label>
            <label class="field span-all"><span>Resultado de cada ave <span class="req">*</span> <span class="muted" data-sr-count></span></span><textarea name="Valores" rows="4" required data-input="sr-calc" placeholder="Por ejemplo: 6 7 5 8 6 …">${esc(valores)}</textarea>
                <small class="field-hint">Cópialos del reporte del laboratorio, en orden, separados por espacios, comas o saltos de línea.</small></label>
        </div>
        <div data-sr-aviso></div>
        <dl class="kv pz-stats" data-sr-resumen></dl>
    </section>
    <section class="form-section"><div class="form-grid">
        ${linea === 'Incubadora' ? `<label class="field"><span>Lote reproductora de origen</span><input name="LoteOrigen" maxlength="80" value="${esc(s?.LoteOrigen || '')}">
            <small class="field-hint">Lote de las madres de estos pollitos (opcional).</small></label>` : ''}
        <label class="field"><span>Laboratorio</span><input name="Laboratorio" maxlength="120" value="${esc(s?.Laboratorio || preset.laboratorio || '')}"></label>
        <label class="field span-all"><span>Comentarios</span><textarea name="Comentarios" maxlength="1000">${esc(s?.Comentarios || '')}</textarea></label>
        ${s ? justField('Explica qué se corrige (por ejemplo, un resultado mal digitado).') : `<label class="check span-all"><input type="checkbox" name="otra"${preset.otra ? ' checked' : ''}> Al guardar, agregar otra enfermedad para esta misma nave y fecha</label>`}
    </div></section>`;
    let siguiente = null;
    const promesa = openModal({
        title: s ? 'Editar muestreo serológico' : `Nuevo muestreo serológico · ${linea.toLowerCase()}`, wide: true, body,
        submitLabel: s ? 'Guardar cambios' : 'Guardar muestreo',
        onSubmit: async (fd) => {
            const metodo = metodoById(fd.get('MetodoID')) || (s && { Prueba: s.Prueba });
            if (!metodo) throw new Error('Elige el método o kit (o crea uno para esta enfermedad).');
            const { valores: v, malos } = leerValoresSero(fd.get('Valores'), metodo.Prueba);
            const muestra = toNum(fd.get('Muestra'));
            if (malos.length) throw new Error(`Hay resultados no válidos: ${malos.slice(0, 8).join(' ')}.`);
            if (v.length !== muestra) throw new Error(`La muestra es de ${muestra} y hay ${v.length} resultados. Deben ser exactamente ${muestra}.`);
            const p = { Muestra: muestra, Valores: v, Laboratorio: fd.get('Laboratorio') || '', Comentarios: fd.get('Comentarios') || '', LoteOrigen: fd.get('LoteOrigen') || '' };
            if (s) Object.assign(p, { SerologiaID: s.SerologiaID, justificacion: fd.get('justificacion') });
            else Object.assign(p, { Linea: linea, NaveID: fd.get('NaveID'), Fecha: fd.get('Fecha'), EnfermedadID: fd.get('EnfermedadID'), MetodoID: fd.get('MetodoID') });
            await save('saveSerologia', p, s ? 'Muestreo actualizado.' : 'Muestreo guardado.', { snapshot: false });
            if (!s && fd.get('otra')) siguiente = { naveId: p.NaveID, fecha: p.Fecha, muestra, laboratorio: p.Laboratorio, otra: true };
        },
    });
    const form = $('#modal form');
    if (form && s) form.dataset.nave = s.NaveID;
    srRender(form);
    await promesa;
    if (siguiente) setTimeout(() => openSerologiaModal(null, siguiente), 60);
}

async function verSerologiaModal(s) {
    const valores = String(s.Valores || '').split(',').filter(Boolean);
    const filas = [['Enfermedad', s.Enfermedad], ['Método', `${s.Metodo} · ${s.Prueba}`], ['Edad', s.EdadDias !== '' ? textoEdad(Number(s.EdadDias)) : '—'], ['Muestra', s.Muestra]]
        .concat(s.Prueba === 'PLACA' ? [] : [[s.Prueba === 'HI' ? 'Promedio (log2)' : 'Media geométrica', fmtNum(valorResumen(s))], ['Dato mayor', fmtNum(s.Maximo)], ['Dato menor', fmtNum(s.Minimo)]])
        .concat(s.Prueba === 'ELISA' ? [['CV', `${fmtNum(s.CV)} %`]] : [])
        .concat(s.PctPositivos !== '' ? [['Positivas', `${s.Positivos} (${fmtNum(s.PctPositivos)} %)`]] : [])
        .concat(s.PctProtegidas !== '' ? [['Protegidas', `${fmtNum(s.PctProtegidas)} %`]] : [])
        .concat(s.LoteOrigen ? [['Lote de origen', s.LoteOrigen]] : []).concat(s.Laboratorio ? [['Laboratorio', s.Laboratorio]] : [])
        .concat(s.EdadVacunacion ? [['Vacunar el día (Deventer)', `Día ${s.EdadVacunacion} de vida`]] : [])
        .concat(s.Deventer ? [['Cálculo de Deventer', s.Deventer]] : []);
    await openModal({
        title: `Muestreo del ${fmtDate(s.Fecha)}`, cancelLabel: 'Cerrar', wide: true,
        body: `<p class="center">${esc([granjaName(s.GranjaID), s.Sector, s.Nave, s.NumeroLote ? 'Lote ' + s.NumeroLote : ''].filter(Boolean).join(' / '))}</p>
            <dl class="kv">${filas.map(([l, v]) => `<div><dt>${esc(l)}</dt><dd>${esc(v)}</dd></div>`).join('')}</dl>
            <section class="form-section"><h3>${ico('list-checks')}Resultado de cada ave</h3>
            <ol class="pz-list">${valores.map((v, i) => `<li><span class="pz-num">${i + 1}</span>${esc(v)}</li>`).join('')}</ol></section>`,
    });
}

async function deleteSerologiaFlow(s) {
    await openModal({
        title: 'Eliminar muestreo serológico', danger: true, submitLabel: 'Eliminar',
        body: `<p class="center">${esc(s.Enfermedad)} · ${esc(fmtDate(s.Fecha))} · ${esc([s.Sector, s.Nave].join(' / '))} · muestra de ${esc(s.Muestra)}.</p>${justField()}`,
        onSubmit: async fd => save('deleteSerologia', { SerologiaID: s.SerologiaID, justificacion: fd.get('justificacion') }, 'Muestreo eliminado.', { snapshot: false }),
    });
}

/** Método o kit: general (solo admin) o propio de una granja. */
async function openMetodoModal(m = null, preset = {}) {
    const granjas = varias() ? state.db.granjas.filter(g => g.Estado === 'Activo') : state.db.granjas.filter(g => g.GranjaID === state.granjaId);
    const granjaHtml = m
        ? `<label class="field"><span>De</span><input value="${esc(m.GranjaID ? granjaName(m.GranjaID) : 'General (todas las granjas)')}" readonly></label>`
        : isAdmin()
            ? `<label class="field"><span>De <span class="req">*</span></span><select name="GranjaID"><option value="">General (todas las granjas)</option>${granjas.map(g =>
                `<option value="${esc(g.GranjaID)}"${g.GranjaID === preset.granjaId ? ' selected' : ''}>${esc(g.NombreGranja)}</option>`).join('')}</select>
                <small class="field-hint">Un método general lo pueden usar todas las granjas.</small></label>`
            : esEmpresa()
                ? `<label class="field"><span>De <span class="req">*</span></span><select name="GranjaID" required>${granjas.map(g =>
                    `<option value="${esc(g.GranjaID)}"${g.GranjaID === (preset.granjaId || state.f.granjas[0]) ? ' selected' : ''}>${esc(g.NombreGranja)}</option>`).join('')}</select></label>`
            : `<label class="field"><span>De</span><input value="${esc(granjaName(state.granjaId))}" readonly><input type="hidden" name="GranjaID" value="${esc(state.granjaId)}"></label>`;
    const enfId = m?.EnfermedadID || preset.enfId || state.srEnf || state.db.enfermedades[0]?.EnfermedadID;
    const pruebasDe = id => String(enfermedadById(id)?.Pruebas || '').split(',').filter(Boolean);
    const opcionesPrueba = (id, elegida) => pruebasDe(id).map(p => `<option value="${p}"${p === elegida ? ' selected' : ''}>${esc(PRUEBA_NOMBRE[p])}</option>`).join('');
    const pruebaIni = m?.Prueba || pruebasDe(enfId)[0];
    const promesa = openModal({
        title: m ? 'Editar método o kit' : 'Nuevo método o kit', wide: true,
        body: `<div class="form-grid">${granjaHtml}
            <label class="field"><span>Enfermedad <span class="req">*</span></span><select name="EnfermedadID" required data-change="met-enf">${state.db.enfermedades.map(e =>
            `<option value="${esc(e.EnfermedadID)}"${e.EnfermedadID === enfId ? ' selected' : ''}>${esc(e.Nombre)}</option>`).join('')}</select></label>
            <label class="field"><span>Prueba <span class="req">*</span></span><select name="Prueba" required data-change="met-prueba">${opcionesPrueba(enfId, pruebaIni)}</select></label>
            <label class="field"><span>Nombre del método o kit <span class="req">*</span></span><input name="Nombre" required maxlength="80" value="${esc(m?.Nombre || '')}" placeholder="Por ejemplo: IDEXX IBD"></label>
            <label class="field"><span>Laboratorio</span><input name="Laboratorio" maxlength="120" value="${esc(m?.Laboratorio || '')}"></label>
            <label class="field" data-met-corte><span>Positivo desde</span><input type="number" name="CortePositivo" min="0" step="any" inputmode="decimal" value="${esc(m?.CortePositivo || '')}">
                <small class="field-hint">Título a partir del cual un ave cuenta como positiva.</small></label>
            <label class="field" data-met-corte><span>Nivel protector</span><input type="number" name="NivelProtector" min="0" step="any" inputmode="decimal" value="${esc(m?.NivelProtector || '')}">
                <small class="field-hint">Título a partir del cual un ave se considera protegida.</small></label>
            <label class="field span-all" data-met-corte><span>Fuente del corte o del nivel</span><input name="Fuente" maxlength="300" value="${esc(m?.Fuente || '')}" placeholder="Por ejemplo: inserto del kit, versión y página">
                <small class="field-hint">Escribe un corte o un nivel solo si está publicado. Si no hay fuente, déjalos vacíos: no se inventan.</small></label>
            <label class="field" data-met-deventer><span>Título de quiebre de la vacuna</span><input type="number" name="TituloQuiebre" min="1" step="any" inputmode="decimal" value="${esc(m?.TituloQuiebre || '')}">
                <small class="field-hint">El que indica el fabricante de la vacuna que usa la granja, para este kit. Con él se calcula el día de vacunación (fórmula de Deventer).</small></label>
            <label class="field" data-met-deventer><span>% de aves vacunables</span><input type="number" name="PorcentajeDeventer" min="1" max="100" step="any" inputmode="decimal" value="${esc(m?.PorcentajeDeventer || '')}" placeholder="5">
                <small class="field-hint">Vacío = 5 %, el valor por defecto de la fórmula.</small></label>
            ${m ? justField() : ''}</div>`,
        submitLabel: m ? 'Guardar cambios' : 'Guardar método',
        onSubmit: async (fd) => {
            const p = formPayload(fd);
            if ((p.CortePositivo || p.NivelProtector) && !p.Fuente) throw new Error('Indica la fuente del corte o del nivel protector.');
            if (m) Object.assign(p, { MetodoID: m.MetodoID });
            await save('saveMetodoSerologia', p, m ? 'Método actualizado.' : 'Método guardado.');
        },
    });
    const f = $('#modal form');
    if (f?.Prueba) CHANGES['met-prueba'](f.Prueba);
    return promesa;
}

async function deleteMetodoFlow(m) {
    await openModal({
        title: 'Eliminar método o kit', danger: true, submitLabel: 'Eliminar',
        body: `<p class="center">${esc(nombreMetodo(m))} · ${esc(enfermedadById(m.EnfermedadID)?.Nombre || '')}.</p>
            <p class="muted center">Solo se puede eliminar si no tiene muestreos.</p>${justField()}`,
        onSubmit: async fd => save('deleteMetodoSerologia', { MetodoID: m.MetodoID, justificacion: fd.get('justificacion') }, 'Método eliminado.'),
    });
}

// ---------- Pesos estándar ----------
async function openEstandarModal(x = null) {
    const gen = x?.LineaGenetica || state.estGen, tipo = x?.Tipo || state.estTipo;
    const ultima = [...(state.estFilas?.values() || [])].pop();
    await openModal({
        title: x ? `Corregir día ${x.EdadDias} · ${gen}` : `Agregar edad · ${gen}`, wide: true,
        body: `<p class="center">${esc(gen)} · tabla ${esc(tipo === 'Comercial' ? 'comercial' : 'de hembra reproductora')}. Los pesos van en gramos.</p>
            <div class="form-grid">
            <label class="field"><span>Edad (días) <span class="req">*</span></span><input type="number" name="EdadDias" required min="0" max="1000" step="1" value="${esc(x?.EdadDias ?? '')}">
                <small class="field-hint">Semana × 7 (por ejemplo, semana 17 = 119 días).</small></label>
            <label class="field"><span>Mínimo (g)</span><input type="number" name="MinimoG" min="0" step="any" value="${esc(x?.MinimoG ?? '')}"></label>
            <label class="field"><span>Estándar (g) <span class="req">*</span></span><input type="number" name="EstandarG" required min="0" step="any" value="${esc(x?.EstandarG ?? '')}"></label>
            <label class="field"><span>Máximo (g)</span><input type="number" name="MaximoG" min="0" step="any" value="${esc(x?.MaximoG ?? '')}"></label>
            <label class="field span-all"><span>Fuente (guía, edición y página) <span class="req">*</span></span><input name="Fuente" required maxlength="300" value="${esc(x?.Fuente || ultima?.Fuente || '')}"></label>
            ${x ? justField('Explica qué cambió en la guía.') : ''}</div>`,
        submitLabel: x ? 'Guardar corrección' : 'Agregar',
        onSubmit: async (fd) => {
            const p = { ...formPayload(fd), LineaGenetica: gen, Tipo: tipo };
            if (x) p.EstandarID = x.EstandarID;
            await save('saveEstandar', p, x ? 'Peso corregido.' : 'Edad agregada.');
        },
    });
}
async function deleteEstandarFlow(x) {
    await openModal({
        title: `Borrar día ${x.EdadDias} · ${x.LineaGenetica}`, danger: true, submitLabel: 'Borrar',
        body: `<p class="center">Estándar ${esc(fmtNum(x.EstandarG))} g. Los pesajes de esa edad se compararán con las edades vecinas.</p>${justField()}`,
        onSubmit: async fd => save('deleteEstandar', { EstandarID: x.EstandarID, justificacion: fd.get('justificacion') }, 'Fila borrada.'),
    });
}

// ---------- Listas ----------
async function deleteListaFlow(l) {
    if (!l) return;
    await openModal({
        title: `Quitar "${l.Valor}"`, danger: true, submitLabel: 'Quitar',
        body: `<p class="center">Se quita de la lista <strong>${esc(LISTA_LABEL[l.Tipo] || l.Tipo)}</strong>. Los datos ya guardados no cambian.</p>
            <p class="muted center">Una línea genética que ya usa algún lote no se puede quitar.</p>${justField()}`,
        onSubmit: async fd => save('deleteLista', { ListaID: l.ListaID, justificacion: fd.get('justificacion') }, 'Valor eliminado.'),
    });
}

// ==========================================================================
// 12. ACCIONES, AUTENTICACIÓN Y ARRANQUE
// ==========================================================================

function setFiltro(patch) {
    Object.assign(state.f, patch);
    render();
}

const ACTIONS = {
    'toggle-password': btn => setPasswordVisible(btn, btn.getAttribute('aria-pressed') !== 'true'),
    'toggle-sidebar': () => toggleSidebar(),
    'logout': async () => { closeModal(); await signOut(auth); },
    'sync': async () => {
        try { await refreshAll(); toast('Datos actualizados.'); } catch (e) { toast(e.message, 'error'); }
    },
    'retry': () => { state.cache.clear(); render(); },
    'retry-session': () => bootSession(true),
    'forgot-password': async () => {
        const email = $('#login-form [name=email]').value.trim();
        if (!email) return showLoginError('Escribe tu correo y vuelve a presionar "¿Olvidaste tu contraseña?".');
        try {
            await sendPasswordResetEmail(auth, email);
            toast('Si el correo está registrado, recibirás un enlace para restablecer tu contraseña.', 'success', 7000);
        } catch (e) { showLoginError(authErrorMsg(e)); }
    },
    'nav': el => go(el.dataset.view),
    'cmp-limpiar': () => { state.cmp.sel = []; render(); },
    'sr-comp-prueba': (el) => { state.srComp = { prueba: el.dataset.prueba, enfs: null }; render(); },
    'set-unidad': (el) => { state.unidad = el.dataset.unidad; guardarPref('granjas.unidad', state.unidad); render(); },

    // Pesajes
    'new-pesaje': () => openPesajeModal(),
    'edit-pesaje': el => openPesajeModal(state.pesajes.get(el.dataset.id)),
    'ver-pesaje': el => verPesajeModal(state.pesajes.get(el.dataset.id)),
    'delete-pesaje': el => deletePesajeFlow(state.pesajes.get(el.dataset.id)),
    'pz-add': () => { const i = $('#pz-input'); if (i && i.value.trim()) { pzAgregar(i.value); i.value = ''; i.focus(); } },
    'pz-undo': () => { if (pz?.pesos.length) { pz.pesos.pop(); pzRender(); $('#pz-input')?.focus(); } },
    'pz-del': (el) => { if (pz) { pz.pesos.splice(Number(el.dataset.i), 1); pzRender(); } },

    // Ambiente
    'new-lectura': () => openLecturaModal(),
    'delete-lectura': el => deleteLecturaFlow(state.lecturas.get(el.dataset.id)),

    // Estructura de la granja
    'new-sector': el => openSectorModal(el.dataset.granja, el.dataset.linea),
    'edit-sector': (el) => { const s = sectorById(el.dataset.id); openSectorModal(s.GranjaID, s.Linea, s); },
    'estado-sector': el => estadoEstructuraFlow('sector', sectorById(el.dataset.id)),
    'delete-sector': el => deleteEstructuraFlow('sector', sectorById(el.dataset.id)),
    'new-nave': el => openNaveModal(el.dataset.id),
    'edit-nave': el => openNaveModal(null, naveById(el.dataset.id)),
    'estado-nave': el => estadoEstructuraFlow('nave', naveById(el.dataset.id)),
    'delete-nave': el => deleteEstructuraFlow('nave', naveById(el.dataset.id)),
    'new-lote': el => openLoteModal(el.dataset.id),
    'new-naves': el => openNavesVariasModal(el.dataset.id),
    'lote-naves': el => lotesNavesModal(loteById(el.dataset.id)),
    'agregar-nave-lote': el => agregarNaveLote(el),
    'sacar-nave': (el) => { closeModal(); sacarNaveModal(el.dataset.lote, el.dataset.id); },
    'forma-lote': (el) => { closeModal(); formaLoteModal(loteById(el.dataset.id)); },
    'edit-lote': el => openLoteModal(null, loteById(el.dataset.id)),
    'cerrar-lote': el => cerrarLoteFlow(loteById(el.dataset.id)),
    'reabrir-lote': el => reabrirLoteFlow(loteById(el.dataset.id)),
    // Desde el formulario de registro o de pesaje: cierra ese formulario y abre el del lote.
    'lote-desde-form': (el) => {
        const lote = loteById(el.dataset.id);
        closeModal();
        setTimeout(() => openLoteModal(null, lote), 60);
    },

    // Filtros
    'set-linea': (el) => { state.linea = el.dataset.linea; state.f.sector = ''; state.f.nave = ''; render(); },
    'set-agrupar': el => setFiltro({ agrupar: el.dataset.agrupar }),
    'preset': (el) => {
        const d = Number(el.dataset.days);
        setFiltro(d ? { desde: daysAgo(d), hasta: today() } : { desde: '', hasta: '' });
    },
    'goto-dashboard': (el) => { focusGranja(el.dataset.id); go('dashboard'); },
    'goto-registros': (el) => { focusGranja(el.dataset.id); go('registros'); },

    // Tablas y gráficos
    'sort': (el) => {
        const { table: id, k } = el.dataset;
        const s = state.sort[id];
        state.sort[id] = { k, dir: s?.k === k && s.dir === 'asc' ? 'desc' : 'asc' };
        refreshTable(id);
    },
    'xlsx': el => downloadXlsx(el.dataset.table),
    'png': async (el) => {
        const spec = charts[el.dataset.chart];
        if (!spec) return;
        await exportChartPng(spec, `${slug(spec.title)}_${today()}.png`);
        toast('Gráfico descargado.', 'success');
    },

    // Registros
    'new-registro': () => openRegistroModal(),
    'edit-registro': el => openRegistroModal(state.registros.get(el.dataset.id)),
    'delete-registro': el => deleteRegistroFlow(state.registros.get(el.dataset.id)),
    'restore-registro': el => restoreRegistroFlow(state.registros.get(el.dataset.id)),

    // Granjas y perfiles
    'new-granja': () => openGranjaModal(),
    'new-empresa': () => openEmpresaModal(),
    'plan-linea': el => openPlanLineaModal(el.dataset.granja, el.dataset.linea),
    'ir-miplan': () => { closeModal(); go('miplan'); },
    'res-actualizar': el => actualizarResumenes(el.dataset.todas === '1'),
    'res-linea': (el) => { state.resLinea = el.dataset.linea; render(); },
    'detalle-cobro': el => detalleCobroModal(cobroById(el.dataset.id)),
    'registrar-pago': el => registrarPagoModal(cobroById(el.dataset.id)),
    'prorroga': el => prorrogaModal(cobroById(el.dataset.id)),
    'anular-pago': el => anularPagoModal(el.dataset.id),
    'guardar-herramientas': () => guardarHerramientas(),
    'guardar-precios': () => guardarPrecios(),
    'edit-empresa': el => openEmpresaModal(empresaById(el.dataset.id)),
    'mover-empresa': el => moverEmpresaModal(empresaById(el.dataset.id)),
    'editar-dueno': el => openDuenoModal((state.db.duenos || []).find(d => d.DuenoID === el.dataset.id)),
    'estado-dueno': el => estadoDuenoFlow((state.db.duenos || []).find(d => d.DuenoID === el.dataset.id)),
    'estado-empresa': el => estadoEmpresaFlow(empresaById(el.dataset.id)),
    'delete-empresa': el => deleteEmpresaFlow(empresaById(el.dataset.id)),
    'empresa-granjas': (el) => { state.empresaFiltro = el.dataset.id; go('granjas'); },
    'open-granja': el => go('granja', el.dataset.id),
    'edit-granja': el => openGranjaModal(granjaById(el.dataset.id)),
    'estado-granja': el => estadoGranjaFlow(granjaById(el.dataset.id)),
    'delete-granja': el => deleteGranjaFlow(granjaById(el.dataset.id)),
    'new-perfil': el => openPerfilModal(el.dataset.granja, el.dataset.linea),
    'edit-perfil': (el) => { const p = state.db.perfiles.find(x => x.PerfilID === el.dataset.id); openPerfilModal(p.GranjaID, p.Linea, p); },
    'estado-perfil': el => estadoPerfilFlow(state.db.perfiles.find(x => x.PerfilID === el.dataset.id)),
    'delete-perfil': el => deletePerfilFlow(state.db.perfiles.find(x => x.PerfilID === el.dataset.id)),
    'copy-perfil': (el) => {
        const src = state.db.perfiles.find(x => x.PerfilID === el.dataset.from);
        const form = el.closest('form');
        if (!src || !form) return;
        state.db.campos.perfil.forEach((c) => { if (form.elements[c.k]) form.elements[c.k].value = src[c.k] ?? ''; });
        toast('Datos copiados como sugerencia. Revísalos y ajusta lo que sea distinto.', 'success');
    },

    // Usuarios
    'new-user': el => openUserModal(null, el.dataset.granja || null),
    'edit-user': el => openUserModal(userById(el.dataset.id)),
    'estado-user': el => estadoUserFlow(userById(el.dataset.id)),
    'reset-user': el => resetUserFlow(userById(el.dataset.id)),
    'delete-user': el => deleteUserFlow(userById(el.dataset.id)),
    'copy-temp': async () => {
        const txt = $('#temp-pass')?.textContent || '';
        try { await navigator.clipboard.writeText(txt); toast('Contraseña copiada.', 'success'); } catch {
            const r = document.createRange(); r.selectNodeContents($('#temp-pass'));
            getSelection().removeAllRanges(); getSelection().addRange(r);
            toast('Selecciónala y cópiala manualmente.', 'info');
        }
    },

    // Listas
    'delete-lista': el => deleteListaFlow(state.db.listas.find(l => l.ListaID === el.dataset.id)),

    // Concentrado y pesos estándar
    // Serología
    'new-serologia': () => openSerologiaModal(),
    'edit-serologia': el => openSerologiaModal(state.serologias.get(el.dataset.id)),
    'ver-serologia': el => verSerologiaModal(state.serologias.get(el.dataset.id)),
    'delete-serologia': el => deleteSerologiaFlow(state.serologias.get(el.dataset.id)),
    'new-metodo': () => openMetodoModal(null, { enfId: state.srEnf }),
    'edit-metodo': el => openMetodoModal(metodoById(el.dataset.id)),
    'delete-metodo': el => deleteMetodoFlow(metodoById(el.dataset.id)),
    'metodo-desde-form': (el) => {
        const f = el.closest('form');
        const preset = { enfId: f?.EnfermedadID?.value, granjaId: naveById(f?.NaveID?.value)?.GranjaID };
        closeModal();
        setTimeout(() => openMetodoModal(null, preset), 60);
    },
    'new-compra': () => openCompraModal(),
    'importar': el => openImportarModal(el.dataset.tipo),
    'importar-semanal': () => openFormatoSemanalModal(),
    'edit-compra': el => openCompraModal(state.compras.get(el.dataset.id)),
    'delete-compra': el => deleteCompraFlow(state.compras.get(el.dataset.id)),
    'est-tipo': (el) => { state.estTipo = el.dataset.tipo; render(); },
    'new-estandar': () => openEstandarModal(),
    'edit-estandar': el => openEstandarModal(state.estFilas.get(el.dataset.id)),
    'delete-estandar': el => deleteEstandarFlow(state.estFilas.get(el.dataset.id)),

    // Seguridad
    'verify-audit': async (el) => {
        el.disabled = true; el.classList.add('is-loading');
        try { state.audit = await withBusy(() => api('verifyAudit')); render(); } catch (e) { toast(e.message, 'error'); } finally { el.disabled = false; el.classList.remove('is-loading'); }
    },
};

/** Al saltar a dashboard/registros desde una granja, se filtra por esa granja. */
function focusGranja(id) {
    if (varias()) state.f.granjas = [id];
    else if (state.db.granjas.some(g => g.GranjaID === id)) state.granjaId = id;
    state.f.sector = '';
    state.f.nave = '';
    ensureLinea(ctxIds());
}

const INPUTS = {
    'sr-calc': el => srRender(el.form),
    'compra-calc': (el) => {
        const box = el.form.querySelector('[data-compra-calc]');
        if (box) box.innerHTML = costoQuintalHtml(el.form.CantidadQuintales.value, el.form.CostoPorQuintal.value);
    },
    'table-q': (el) => { state.q[el.dataset.table] = el.value; refreshTable(el.dataset.table); },
    'iec-calc': (el) => {
        const t = toNum(el.form.Temperatura.value), h = toNum(el.form.Humedad.value);
        const box = el.form.querySelector('[data-iec-preview]');
        if (t === null || h === null) { box.innerHTML = '<p class="muted center">IEC = temperatura (°C) + humedad (%).</p>'; return; }
        const v = round(t + h);
        const r = nivelIec(v);
        box.innerHTML = `<div class="iec-preview iec-${esc(r?.color || '')}"><strong>IEC ${esc(fmtNum(v))}</strong><span>${esc(r?.nivel || '')} · ${esc(r?.descripcion || '')}</span></div>`;
    },
};

const CHANGES = {
    'granja-activa': (el) => { state.granjaId = el.value; state.f.sector = ''; state.f.nave = ''; ensureLinea(ctxIds()); render(); },
    'f-sector': el => setFiltro({ sector: el.value, nave: '' }),
    'f-nave': el => setFiltro({ nave: el.value }),
    'ubic-info': (el) => {
        const f = el.form;
        const box = f.querySelector('[data-ubic-info]');
        const fecha = f.Fecha ? f.Fecha.value : today();
        if (box) box.innerHTML = infoUbicacion(f.NaveID.value, fecha || today());
        syncEdadRegistro(f, loteDeUbicacion(f.NaveID.value, fecha || today()), fecha || today());
        if (pz) pz.esperado = esperadoPesaje(f.NaveID.value, fecha || today());
        srRender(f, { metodos: true });
        if (pz && !pz.editando) {
            pz.naveId = f.NaveID.value; pz.fecha = fecha;
            const b = leerBorrador(pz.naveId, pz.fecha);
            if (b?.pesos?.length && !pz.pesos.length) {
                Object.assign(pz, { unidad: b.unidad, muestra: b.muestra, pesos: b.pesos });
                f.Unidad.value = b.unidad; f.Muestra.value = b.muestra;
                toast(`Se recuperó una muestra sin guardar (${b.pesos.length} pesos).`, 'info');
            }
            pzRender();
        }
    },
    'pz-config': (el) => {
        if (!pz) return;
        const f = el.form;
        if (el.name === 'Unidad' && pz.pesos.length && el.value !== pz.unidad) {
            if (!confirm('Cambiar la unidad convertirá los pesos ya ingresados. ¿Continuar?')) { el.value = pz.unidad; return; }
            const fx = factorDe(pz.unidad) / factorDe(el.value);
            pz.pesos = pz.pesos.map(v => Math.round(v * fx * 1e6) / 1e6);
        }
        pz.unidad = f.Unidad.value;
        pz.muestra = Math.max(0, Math.floor(toNum(f.Muestra.value) || 0));
        pzRender();
    },
    'form-unidad': (el) => {
        // Convierte los pesos ya escritos a la nueva unidad y ajusta etiquetas y máximos.
        const antes = el.dataset.prev, ahora = el.value;
        const fx = factorDe(antes) / factorDe(ahora);
        el.form.querySelectorAll('input[data-peso]').forEach((i) => {
            if (i.value !== '') i.value = String(Math.round(Number(i.value) * fx * 1e4) / 1e4);
            if (i.dataset.maxG) i.max = String(round(Number(i.dataset.maxG) / factorDe(ahora), 4));
        });
        el.form.querySelectorAll('[data-unidad-label]').forEach((s) => { s.textContent = ahora; });
        el.dataset.prev = ahora;
    },
    'f-granja': (el) => {
        const set = new Set(state.f.granjas);
        if (el.checked) set.add(el.value); else set.delete(el.value);
        state.ddOpen = true;
        setFiltro({ granjas: [...set] });
    },
    'f-granja-all': () => { state.ddOpen = true; setFiltro({ granjas: [] }); },
    'f-desde': el => setFiltro({ desde: el.value }),
    'f-hasta': el => setFiltro({ hasta: el.value }),
    'f-eliminados': el => setFiltro({ eliminados: el.checked }),
    'bit-granja': (el) => { state.bitGranja = el.value; render(); },
    'est-gen': (el) => { state.estGen = el.value; render(); },
    'sr-f-enf': (el) => {
        const set = new Set(state.srEnfs);
        if (el.checked) set.add(el.value); else set.delete(el.value);
        state.srEnfs = [...set];
        state.srMetodo = '';
        state.srDd = true;
        render();
    },
    'sr-f-enf-todas': () => { state.srEnfs = []; state.srMetodo = ''; state.srDd = true; render(); },
    'sr-f-metodo': (el) => { state.srMetodo = el.value; render(); },
    'c-fase': (el) => { state.cFase = el.value; render(); },
    'cmp-estado': (el) => { state.cmp.estado = el.value; render(); },
    'cmp-gen': (el) => { state.cmp.gen = el.value; render(); },
    'cmp-q': (el) => { state.cmp.q = el.value; render(); },
    'cmp-lote': (el) => {
        const sel = state.cmp.sel.filter(id => id !== el.value);
        if (el.checked && sel.length < MAX_COMPARAR) sel.push(el.value);
        state.cmp.sel = sel;
        render();
    },
    'sr-comp-enf': (el) => {
        const set = new Set(state.srComp.enfs || []);
        if (el.checked) set.add(el.value); else set.delete(el.value);
        state.srComp.enfs = [...set];
        render();
    },
    'sr-enf': el => srRender(el.form, { metodos: true }),
    'sr-metodo': el => srRender(el.form),
    'met-enf': (el) => {
        const pruebas = String(enfermedadById(el.value)?.Pruebas || '').split(',').filter(Boolean);
        el.form.Prueba.innerHTML = pruebas.map(p => `<option value="${p}">${esc(PRUEBA_NOMBRE[p])}</option>`).join('');
        CHANGES['met-prueba'](el.form.Prueba);
    },
    'met-prueba': (el) => {
        // La aglutinación en placa da positivo o negativo: no lleva corte ni nivel.
        el.form.querySelectorAll('[data-met-corte]').forEach((x) => { x.hidden = el.value === 'PLACA'; });
        // Fórmula de Deventer: solo ELISA de enfermedades con vida media publicada (por ahora, Gumboro).
        const cod = enfermedadById(el.form.EnfermedadID?.value)?.Codigo;
        el.form.querySelectorAll('[data-met-deventer]').forEach((x) => { x.hidden = !(el.value === 'ELISA' && state.db.deventer?.[cod]); });
    },
    'user-rol': (el) => {
        const form = el.form;
        if (!form) return;
        const fs = form.querySelector('[data-granjas]');
        if (fs) fs.hidden = el.value !== ROLES.REPRESENTANTE;
        const em = form.querySelector('[data-empresa]');
        if (em) em.hidden = el.value === ROLES.ADMIN;
        const nota = form.querySelector('[data-nota-empresa]');
        if (nota) nota.hidden = el.value !== ROLES.ADMIN_EMPRESA;
    },
    // Solo se ofrecen las granjas de la empresa elegida (una cuenta no mezcla empresas).
    'user-empresa': (el) => {
        el.form?.querySelectorAll('[data-empresa-de]').forEach((c) => {
            c.hidden = !!el.value && c.dataset.empresaDe !== el.value;
            if (c.hidden) c.querySelector('input').checked = false;
        });
    },
    'empresa-filtro': (el) => { state.empresaFiltro = el.value; render(); },
    'admin-empresa': el => avisoDuenoEmpresa(el),
    'empresa-activa': (el) => {
        state.empresaId = el.value;
        try { localStorage.setItem('granjas.empresa', el.value); } catch { /* sin almacenamiento */ }
        state.f.granjas = [];
        state.f.sector = '';
        state.f.nave = '';
        ensureLinea(ctxIds());
        render();
    },
    'res-empresa': (el) => { state.resEmpresa = el.value; render(); },
    // Al elegir un plan se muestra cuánto cambia el cobro (granja y línea vienen del formulario).
    'cotizar': (el) => {
        const box = el.closest('[data-granja]');
        if (box) cotizarEn(el.form, box.dataset.granja, box.dataset.linea);
    },
};

const SUBMITS = {
    login: async (form) => {
        const email = form.email.value.trim();
        const password = form.password.value;
        if (!email || !password) return showLoginError('Escribe tu correo y tu contraseña.');
        if (CONFIG.TURNSTILE_SITE_KEY && !humano) return showLoginError('Completa la verificación de seguridad ("No soy un robot").');
        hideAllPasswords();
        const btn = form.querySelector('[type=submit]');
        btn.classList.add('is-loading'); btn.disabled = true;
        showLoginError('');
        try { await signInWithEmailAndPassword(auth, email, password); } catch (e) { showLoginError(authErrorMsg(e)); } finally { btn.classList.remove('is-loading'); btn.disabled = false; }
    },
    'change-password': async (form) => {
        const err = form.querySelector('.form-error');
        const fail = (msg) => { err.textContent = msg; err.hidden = false; };
        err.hidden = true;
        const p1 = form.p1.value;
        const problema = passwordProblem(p1);
        if (problema) return fail(problema);
        if (p1 !== form.p2.value) return fail('Las contraseñas no coinciden.');
        hideAllPasswords();
        const btn = form.querySelector('[type=submit]');
        btn.classList.add('is-loading'); btn.disabled = true;
        try {
            await updatePassword(auth.currentUser, p1);
            // El servidor confirma el cambio consultando a Firebase: no confía en el navegador.
            await api('passwordChanged', {}, { forceRefresh: true });
            form.reset();
            toast('Contraseña actualizada.', 'success');
            await bootSession(true);
        } catch (e) {
            if (e.code === 'auth/requires-recent-login') {
                fail('Por seguridad, vuelve a iniciar sesión con la contraseña temporal y cámbiala de inmediato.');
                setTimeout(() => signOut(auth), 3000);
            } else if (e.code === 'auth/weak-password' || e.code === 'auth/password-does-not-meet-requirements') {
                fail('Esa contraseña no cumple los requisitos de seguridad. Elige otra.');
            } else {
                fail(String(e.code || '').startsWith('auth/') ? authErrorMsg(e) : (e.message || authErrorMsg(e)));
            }
        } finally { btn.classList.remove('is-loading'); btn.disabled = false; }
    },
    'add-lista': async (form) => {
        if (!form.checkValidity()) return form.reportValidity();
        const btn = form.querySelector('[type=submit]');
        btn.disabled = true;
        try { await save('saveLista', { Tipo: form.Tipo.value, Valor: form.Valor.value.trim() }, 'Valor agregado.'); } catch (e) { toast(e.message, 'error'); } finally { btn.disabled = false; }
    },
    configuracion: async (form) => {
        if (!form.checkValidity()) return form.reportValidity();
        const btn = form.querySelector('[type=submit]');
        btn.disabled = true;
        const valores = {};
        ['DiasAviso', 'DiasGracia', 'HoraTarea'].forEach((k) => { if (form.elements[k]) valores[k] = Number(form.elements[k].value); });
        try {
            const r = await save('saveConfiguracion', { valores, justificacion: form.justificacion.value }, x => (x?.sinCambios ? 'No había cambios.' : 'Configuración guardada.'));
            if (r && !r.sinCambios) form.justificacion.value = '';
        } catch (e) { toast(e.message, 'error'); } finally { btn.disabled = false; }
    },
};

function setPasswordVisible(btn, visible) {
    const input = btn.closest('.password-field')?.querySelector('input');
    if (!input) return;
    input.type = visible ? 'text' : 'password';
    btn.innerHTML = ico(visible ? 'eye-off' : 'eye');
    btn.setAttribute('aria-pressed', String(visible));
    const label = visible ? 'Ocultar contraseña' : 'Mostrar contraseña';
    btn.setAttribute('aria-label', label);
    btn.title = label;
    input.focus({ preventScroll: true });
}
function hideAllPasswords() {
    $$('.password-toggle[aria-pressed="true"]').forEach(btn => setPasswordVisible(btn, false));
    document.activeElement?.blur?.();
}
function passwordProblem(pw) {
    if (pw.length < CONFIG.MIN_CLAVE) return `Debe tener al menos ${CONFIG.MIN_CLAVE} caracteres.`;
    if (!/[a-zñ]/.test(pw) || !/[A-ZÑ]/.test(pw) || !/\d/.test(pw)) return 'Debe combinar mayúsculas, minúsculas y números.';
    const local = String(auth.currentUser?.email || '').split('@')[0].toLowerCase();
    if (local.length >= 4 && pw.toLowerCase().includes(local)) return 'No uses tu correo dentro de la contraseña.';
    if (/^(.)\1+$/.test(pw) || /1234567|abcdefg|qwerty|password|contrase/i.test(pw)) return 'Esa contraseña es demasiado fácil de adivinar.';
    return '';
}

document.addEventListener('click', (ev) => {
    // Cierra el selector de granjas al hacer clic fuera de él.
    const dd = $('details.dropdown[open]');
    if (dd && !dd.contains(ev.target)) { dd.open = false; state.ddOpen = false; state.srDd = false; }
    const el = ev.target.closest('[data-action]');
    if (!el || el.disabled) return;
    const fn = ACTIONS[el.dataset.action];
    if (!fn) return;
    ev.preventDefault();
    Promise.resolve(fn(el, ev)).catch((e) => { console.error('[Granjas]', e); toast(e.message || 'No se pudo completar la acción.', 'error'); });
});
document.addEventListener('toggle', (ev) => {
    if (ev.target.matches?.('details.dropdown[data-dd="sr"]')) state.srDd = ev.target.open;
    else if (ev.target.matches?.('details.dropdown')) state.ddOpen = ev.target.open;
}, true);
document.addEventListener('paste', (ev) => {
    // Pegar varios pesos de una vez (separados por espacios, comas o saltos de línea).
    if (ev.target.id !== 'pz-input') return;
    const texto = ev.clipboardData?.getData('text') || '';
    if (/[\s,;]/.test(texto.trim())) { ev.preventDefault(); pzAgregar(texto); }
});
document.addEventListener('input', (ev) => {
    const el = ev.target.closest('[data-input]');
    if (el) INPUTS[el.dataset.input]?.(el, ev);
});
document.addEventListener('change', (ev) => {
    const el = ev.target.closest('[data-change]');
    if (el) CHANGES[el.dataset.change]?.(el, ev);
});
document.addEventListener('submit', (ev) => {
    const form = ev.target.closest('form[data-submit]');
    if (!form) return;
    ev.preventDefault();
    SUBMITS[form.dataset.submit]?.(form, ev);
});
document.addEventListener('keydown', (ev) => {
    // Pesaje: Enter agrega el peso y deja el cursor listo para el siguiente.
    if (ev.key === 'Enter' && ev.target.id === 'pz-input') {
        ev.preventDefault();
        if (ev.target.value.trim()) { pzAgregar(ev.target.value); ev.target.value = ''; }
        return;
    }
    if (ev.key === 'Escape' && mqMobile.matches && $('#screen-app').classList.contains('sidebar-open')) toggleSidebar(false);
});
window.addEventListener('hashchange', () => {
    if (mqMobile.matches) toggleSidebar(false);   // en celular el menú se cierra al navegar
    state.ddOpen = false;
    state.srDd = false;
    render();
    window.scrollTo({ top: 0 });
});

function authErrorMsg(e) {
    const map = {
        'auth/invalid-credential': 'Correo o contraseña incorrectos.',
        'auth/invalid-login-credentials': 'Correo o contraseña incorrectos.',
        'auth/wrong-password': 'Correo o contraseña incorrectos.',
        'auth/user-not-found': 'Correo o contraseña incorrectos.',
        'auth/invalid-email': 'El correo no es válido.',
        'auth/too-many-requests': 'Demasiados intentos. Espera unos minutos e intenta de nuevo.',
        'auth/network-request-failed': 'Sin conexión. Revisa tu internet.',
        'auth/user-disabled': 'Esta cuenta está suspendida. Contacta a la administración.',
    };
    if (!map[e?.code]) console.error('[Granjas] Error de autenticación:', e);
    return map[e?.code] || 'No se pudo completar la operación. Intenta de nuevo.';
}
function showLoginError(msg) {
    const el = $('#login-form .form-error');
    el.textContent = msg;
    el.hidden = !msg;
}
function showDenied(title, text, actionsHtml = '') {
    $('#denied-title').textContent = title;
    $('#denied-text').textContent = text;
    $('#denied-actions').innerHTML = actionsHtml;
    showScreen('denied');
}
function showAccessDenied(e) {
    const titles = { NOT_REGISTERED: 'Cuenta sin acceso', SUSPENDED: 'Acceso suspendido', ACCOUNT_LOCKED: 'Cuenta en revisión' };
    state.me = null;
    showDenied(titles[e.code] || 'Sin acceso', e.message);
}
function showPasswordScreen() {
    $('#password-email').textContent = auth.currentUser?.email || '';
    $('#password-form .form-error').hidden = true;
    showScreen('password');
}

// ---------- Pantalla de carga: barra con porcentaje y buenas prácticas ----------
// La carga real es una sola petición: la barra avanza sola hasta el 95 % (cada vez más despacio) y llega
// al 100 % cuando los datos están listos. Mientras tanto se muestra una buena práctica avícola al azar.
const CONSEJOS = [
    'Anote los datos el mismo día: los registros a tiempo permiten detectar problemas antes de que crezcan.',
    'Lave y desinfecte los bebederos con frecuencia; el agua limpia y fresca mejora el consumo.',
    'Respete la densidad: en engorde, 10 a 12 aves por m²; en postura en piso, 6 a 7 aves por m².',
    'En postura en piso, procure 4 a 5 nidales por cada 20 gallinas y recoja el huevo varias veces al día.',
    'Un bebedero de campana alcanza para 65 a 70 aves; con menos bebederos baja el consumo de agua.',
    'Mantenga un pediluvio con desinfectante a la entrada de cada nave y renuévelo seguido.',
    'Restrinja las visitas: cada persona que entra a la granja puede traer enfermedades.',
    'Retire la mortalidad todos los días y dispóngala lejos de las naves.',
    'Pese una muestra de aves cada semana y compárela con la tabla de su línea genética.',
    'En las horas de calor, ventile: la temperatura y la humedad altas causan estrés calórico.',
    'En días de calor, asegure agua fresca antes que alimento.',
    'Ajuste comederos y bebederos a la altura del lomo de las aves conforme crecen.',
    'Siga el calendario de vacunación que le indique su veterinario.',
    'Controle roedores y aves silvestres: transmiten enfermedades y se comen el alimento.',
    'Guarde el alimento en un lugar seco, sobre tarimas y separado de las paredes.',
    'Trabaje "todo dentro, todo fuera": un lote por nave y limpieza completa entre lotes.',
    'Después de lavar y desinfectar, deje descansar la nave antes del siguiente lote.',
    'La cama debe estar seca y suelta: la cama húmeda produce amoníaco y problemas de patas.',
    'Registre el consumo de alimento todos los días: es la base para calcular la conversión.',
    'Compare su postura con la tabla de la guía: una caída repentina es una señal temprana.',
    'En postura, mantenga un programa de luz constante; los cambios bruscos afectan la producción.',
    'Revise la uniformidad del lote: un lote parejo produce mejor.',
    'Separe las aves enfermas y consulte a su veterinario ante cualquier cambio.',
    'Anote cada vacuna y medicamento con su fecha, dosis y lote.',
    'Limpie y desinfecte el equipo antes de pasarlo de una nave a otra.',
    'Use ropa y botas exclusivas para cada nave siempre que pueda.',
    'Revise los resultados de serología para confirmar que las vacunas protegieron.',
    'Calcule el costo de alimento por docena o por kg de huevo: ayuda a decidir sus compras.',
];
const carga = { pct: 0, timer: null, consejos: null, usados: [] };
function pintarCarga() {
    const p = Math.round(carga.pct);
    const relleno = $('#carga-relleno');
    if (relleno) relleno.style.width = `${p}%`;
    $('#carga-pct') && ($('#carga-pct').textContent = `${p} %`);
    $('#carga-barra')?.setAttribute('aria-valuenow', String(p));
}
function siguienteConsejo() {
    const caja = $('#consejo');
    if (!caja) return;
    if (carga.usados.length >= CONSEJOS.length) carga.usados = [];
    const libres = CONSEJOS.map((_, i) => i).filter(i => !carga.usados.includes(i));
    const i = libres[Math.floor(Math.random() * libres.length)];
    carga.usados.push(i);
    $('#consejo-texto').textContent = CONSEJOS[i];
    caja.hidden = false;
}
function iniciarCarga(texto) {
    if (texto) $('#loading-text').textContent = texto;
    if (carga.timer) return;
    carga.pct = 0;
    pintarCarga();
    siguienteConsejo();
    carga.timer = setInterval(() => { carga.pct += (95 - carga.pct) * 0.025; pintarCarga(); }, 120);
    carga.consejos = setInterval(siguienteConsejo, 7000);
}
async function terminarCarga() {
    clearInterval(carga.timer); clearInterval(carga.consejos);
    carga.timer = null; carga.consejos = null;
    carga.pct = 100;
    pintarCarga();
    await new Promise(r => setTimeout(r, 250));
}
iniciarCarga();

async function bootSession(forceRefresh = false) {
    showScreen('loading');
    iniciarCarga('Cargando su información…');
    try {
        await loadSnapshot({ forceRefresh });
        await terminarCarga();
        if (state.mustChangePassword) return showPasswordScreen();
        renderChrome();
        showScreen('app');
        syncMenuButton();
        if (!location.hash || !VIEWS[parseHash().view]) history.replaceState(null, '', `#/${HOME[state.me.Rol]}`);
        render();
        avisosAlEntrar();
    } catch (e) {
        await terminarCarga();
        if (ACCESS_CODES.has(e.code) || e.code === 'FORBIDDEN') return showAccessDenied(e);
        if (!(e instanceof ApiError)) console.error('[Granjas] Error al cargar la sesión:', e);
        showDenied('No se pudo conectar', e instanceof ApiError ? e.message : MSG.BAD_RESPONSE,
            '<button type="button" class="btn btn-primary btn-block" data-action="retry-session">Reintentar</button>');
    }
}

/**
 * Al cerrar sesión no queda nada del usuario anterior: ni datos en memoria ni la vista
 * dibujada (oculta, pero legible en el código de la página en un equipo compartido).
 */
function limpiarSesion() {
    state.me = null;
    state.db = emptyDb();
    state.cache.clear();
    state.audit = null;
    state.granjaId = null;
    state.f = filtrosIniciales();
    state.q = {};
    state.sort = {};
    state.bitGranja = '';
    state.registros = new Map();
    state.pesajes = new Map();
    state.lecturas = new Map();
    state.compras = new Map();
    state.estFilas = new Map();
    pz = null;
    chartUnmounts.forEach(f => f());
    chartUnmounts = [];
    Object.keys(tables).forEach(k => delete tables[k]);
    Object.keys(charts).forEach(k => delete charts[k]);
    ['#view', '#nav', '#user-chip', '#page-title'].forEach(s => $(s).replaceChildren());
    sesionN++;   // una consulta que llegue tarde ya no entra a la caché de la sesión siguiente
    document.title = 'CHIIK · Registro avícola';
}

// Cierre de sesión automático por inactividad (equipos compartidos).
let lastActivity = Date.now();
['pointerdown', 'keydown', 'wheel', 'touchstart'].forEach(type =>
    window.addEventListener(type, () => { lastActivity = Date.now(); }, { passive: true, capture: true }));
setInterval(() => {
    if (!auth.currentUser || Date.now() - lastActivity < CONFIG.INACTIVIDAD_MIN * 60000) return;
    signOut(auth);
    toast(`Tu sesión se cerró tras ${CONFIG.INACTIVIDAD_MIN} minutos de inactividad.`, 'info', 10000);
}, 30000);

onAuthStateChanged(auth, (user) => {
    state.user = user;
    lastActivity = Date.now();
    if (!user) {
        limpiarSesion();
        $('#login-form').reset();
        terminarCarga();
        showScreen('login');
        mostrarHumano($('#humano'));
        return;
    }
    bootSession();
});
