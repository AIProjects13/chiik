// Lector del formato semanal de producción del cliente (un libro con una hoja por semana: "SEM 01",
// "SEM 02"…). Cada hoja trae un bloque por día (LUNES … DOMINGO) con una fila por galera (nave):
// GALERA, LOTE, EDAD, CANTIDAD AVES AYER, MUERTAS, SELEC., CANT AVES, CONSUMO ALIMENTO (quintales),
// POSTURA HUEVOS… y, al lado del domingo, el cuadro ALIMENTO con los quintales y el precio por fase.
// Módulo puro (sin pantalla): recibe lo que da leerXlsx (web/excel.js) y devuelve filas listas para
// convertir en registros. Lo usa la carga "Desde mi formato semanal" y lo prueba tests/test-formato-semanal.mjs.

const DIAS = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];
const MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

/** Texto comparable: sin tildes, en mayúsculas, sin puntos y con un solo espacio. */
const limpio = v => String(v ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase()
    .replace(/\./g, ' ').replace(/\s+/g, ' ').trim();
const num = (v) => {
    if (v === '' || v === null || v === undefined) return null;
    const n = typeof v === 'number' ? v : Number(String(v).replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
};
const pad = n => String(n).padStart(2, '0');

// Encabezado de una columna → campo (por su texto, no por su posición).
function campoDe(t) {
    if (t === 'GALERA' || t === 'NAVE') return 'galera';
    if (t === 'LOTE') return 'lote';
    if (t === 'EDAD') return 'edad';
    if (t.includes('AYER')) return 'avesAyer';
    if (t === 'MUERTAS' || t === 'MORTALIDAD') return 'muertas';
    if (t.startsWith('SELEC') || t.startsWith('DESCARTE')) return 'seleccion';
    if (t === 'CANT AVES' || t === 'CANTIDAD AVES') return 'avesHoy';
    if (t.startsWith('CONSUMO')) return 'consumoQq';
    if (t === 'POSTURA HUEVOS' || t === 'HUEVOS') return 'huevos';
    return null;
}
const OBLIGATORIOS = ['muertas', 'avesAyer', 'consumoQq', 'huevos'];

/** "LUNES 29 DE DICIEMBRE" → { dia: 1 (lunes), d: 29, mes: 12 } o null. */
function encabezadoDia(v) {
    const m = limpio(v).match(/^(DOMINGO|LUNES|MARTES|MIERCOLES|JUEVES|VIERNES|SABADO)\s+(\d{1,2})\s+DE\s+([A-Z]+)/);
    if (!m) return null;
    const mes = MESES.indexOf(m[3]) + 1;
    return mes ? { dia: DIAS.indexOf(m[1]), d: Number(m[2]), mes } : null;
}

/**
 * Columnas de un bloque que empieza en `c0`: busca cada encabezado en la fila `fe` desde `c0` hasta
 * antes de `c1`. GALERA, LOTE y EDAD se toman del bloque de la izquierda de la misma fila si el bloque
 * de la derecha no los repite (así viene el formato).
 */
function columnasBloque(filaEnc, c0, c1, izquierda) {
    const cols = {};
    for (let j = c0; j < c1; j++) {
        const k = campoDe(limpio(filaEnc[j]));
        if (k && cols[k] === undefined) cols[k] = j;
    }
    ['galera', 'lote', 'edad'].forEach((k) => { if (cols[k] === undefined && izquierda && izquierda[k] !== undefined) cols[k] = izquierda[k]; });
    return cols;
}

/**
 * libro = { hojas: [{ nombre, filas }] } → { dias, concentrado, avisos }
 *  dias: [{ hoja, fecha 'AAAA-MM-DD', galera, lote, edad, avesAyer, muertas, seleccion, avesHoy, consumoQq, huevos }]
 *  concentrado: [{ hoja, desde, hasta, fase, quintales, costoQq }]  (la semana de la hoja, por fase)
 *  avisos: [{ hoja, texto }]
 */
export function leerFormatoSemanal(libro) {
    const dias = [], concentrado = [], avisos = [];
    (libro?.hojas || []).forEach((h) => {
        const mSem = limpio(h.nombre).match(/^SEM\s*(\d+)/);
        if (!mSem) return;
        const semana = Number(mSem[1]);
        const filas = h.filas || [];
        // Año: del título ("SEMANA # 01 - 2026"); si no, del nombre de la hoja.
        const titulo = filas.slice(0, 4).flat().map(limpio).find(t => /\b(19|20)\d{2}\b/.test(t)) || limpio(h.nombre);
        const anio = Number((titulo.match(/\b((?:19|20)\d{2})\b/) || [])[1]);
        if (!anio) { avisos.push({ hoja: h.nombre, texto: 'No se encontró el año en el título de la hoja.' }); return; }

        // Bloques de día: celdas con "LUNES 29 DE DICIEMBRE"…
        const bloques = [];
        filas.forEach((f, i) => (f || []).forEach((v, j) => {
            const e = encabezadoDia(v);
            if (e) bloques.push({ i, j, ...e });
        }));
        if (!bloques.length) { avisos.push({ hoja: h.nombre, texto: 'La hoja no tiene bloques de días.' }); return; }
        const fechas = [];
        bloques.forEach((b) => {
            // La semana 1 empieza en diciembre del año anterior.
            const y = b.mes === 12 && semana <= 2 ? anio - 1 : b.mes === 1 && semana >= 52 ? anio + 1 : anio;
            const fecha = `${y}-${pad(b.mes)}-${pad(b.d)}`;
            const dt = new Date(Date.UTC(y, b.mes - 1, b.d));
            if (dt.getUTCMonth() !== b.mes - 1 || dt.getUTCDay() !== b.dia) {
                avisos.push({ hoja: h.nombre, texto: `"${DIAS[b.dia]} ${b.d}" no coincide con el calendario de ${y}: ese día se salta.` });
                return;
            }
            const filaEnc = filas[b.i + 1] || [];
            const siguiente = bloques.filter(x => x.i === b.i && x.j > b.j).map(x => x.j).sort((a, c) => a - c)[0] ?? filaEnc.length;
            const izq = b.j > 0 ? columnasBloque(filaEnc, 0, b.j, null) : null;
            const cols = columnasBloque(filaEnc, b.j, siguiente, izq);
            const faltan = OBLIGATORIOS.filter(k => cols[k] === undefined);
            if (cols.galera === undefined || cols.lote === undefined || faltan.length) {
                avisos.push({ hoja: h.nombre, texto: `${DIAS[b.dia]} ${b.d}: no se encontraron las columnas ${['galera', 'lote'].filter(k => cols[k] === undefined).concat(faltan).join(', ')}.` });
                return;
            }
            fechas.push(fecha);
            for (let i = b.i + 2; i < filas.length; i++) {
                const f = filas[i] || [];
                const galera = String(f[cols.galera] ?? '').trim();
                if (!galera || /^TOTAL/.test(limpio(galera)) || encabezadoDia(f[b.j])) break;
                const lote = String(f[cols.lote] ?? '').trim();
                const avesAyer = num(f[cols.avesAyer]);
                // Galera sin lote o sin aves (vacía, o días futuros en cero): se salta.
                if (!lote || !avesAyer) continue;
                const v = k => (cols[k] === undefined ? null : num(f[cols[k]]));
                dias.push({ hoja: h.nombre, fecha, galera, lote, edad: v('edad'), avesAyer, muertas: v('muertas') || 0, seleccion: v('seleccion') || 0,
                    avesHoy: v('avesHoy'), consumoQq: v('consumoQq'), huevos: v('huevos') });
            }
        });

        // Cuadro ALIMENTO (fase, quintales de la semana y precio por quintal).
        if (fechas.length) {
            const desde = fechas.slice().sort()[0], hasta = fechas.slice().sort().pop();
            filas.forEach((f, i) => (f || []).forEach((v, j) => {
                if (limpio(v) !== 'ALIMENTO') return;
                for (let k = i + 1; k < filas.length; k++) {
                    const r = filas[k] || [];
                    const fase = String(r[j] ?? '').trim();
                    if (!fase || /^TOTAL/.test(limpio(fase))) break;
                    const quintales = num(r[j + 1]);
                    // El precio por quintal es el primer número a la derecha de los quintales.
                    let costoQq = null;
                    for (let c = j + 2; c < r.length && costoQq === null; c++) costoQq = num(r[c]);
                    if (quintales) concentrado.push({ hoja: h.nombre, desde, hasta, fase, quintales, costoQq });
                }
            }));
        }
    });
    dias.sort((a, b) => (a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0));
    return { dias, concentrado, avisos };
}

/** Junta las galeras de un mismo lote y día (sumas; los porcentajes los calcula la app con las sumas). */
export function agruparPorLote(dias) {
    const grupos = new Map();
    dias.forEach((x) => {
        const k = `${x.fecha}|${x.lote}`;
        const g = grupos.get(k) || { fecha: x.fecha, lote: x.lote, hoja: x.hoja, galeras: [], avesAyer: 0, muertas: 0, seleccion: 0, consumoQq: 0, huevos: 0, edad: x.edad };
        g.galeras.push(x.galera);
        ['avesAyer', 'muertas', 'seleccion', 'consumoQq', 'huevos'].forEach((c) => { g[c] = Math.round((g[c] + (x[c] || 0)) * 1e4) / 1e4; });
        grupos.set(k, g);
    });
    return [...grupos.values()];
}

/** Número de la galera escrito en el Excel ("5") contra el nombre de una nave ("Galera 5", "Nave 05"). */
export function numeroDeNave(nombre) {
    const m = String(nombre || '').match(/(\d+)\s*$/);
    return m ? String(Number(m[1])) : null;
}
