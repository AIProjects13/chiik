/* ==========================================================================
   Granjas · Gráficos en canvas (sin librerías externas)
   - Líneas y áreas (series punteadas para valores "tabla"/ideales), columnas
     agrupadas o apiladas, barras horizontales, dona y medidor semicircular.
   - Responsive (se redibuja al cambiar de tamaño) y nítido en pantallas HiDPI.
   - Tooltip al pasar el cursor o tocar (móvil).
   - Exportación a PNG con título, leyenda y fondo blanco.
   spec = { title, type: 'line'|'area'|'bar', labels: [], unit?, series: [{ name, data: [], color, dashed?, stack? }] }
          | { type: 'hbar' | 'donut' | 'gauge', ... } (ver cada función)
   Paleta validada (dataviz): escala verde ordinal para categorías con orden, verde + acentos
   para categorías sin orden y gris neutro punteado para "tabla"/"ideal".
   ========================================================================== */

export const PALETTE = {
    // Marca CHIIK (manual de marca, 07 · Visualización de datos): planas, en bloques y honestas.
    primary: '#0B2545',       // el dato real va en azul noche
    tabla: '#0A6F78',         // referencia (tabla/estándar de la guía): verde profundo, siempre punteado
    rango: '#3DBCA5',         // rango de la guía: banda verde clara
    // Escala ordinal (claro → oscuro) para categorías con orden (tamaños de huevo, etapas): el degradado verde del logo.
    o1: '#3DBCA5', o2: '#00A388', o3: '#0A6F78', o4: '#0B2545',
    // Series sin orden (naves, granjas), en el orden del manual: azul noche, verde dato, maíz, verde profundo;
    // si hay más de 4 (el manual pide máximo 4), siguen pizarra y maíz oscuro.
    cat: ['#0B2545', '#00A388', '#E0A100', '#0A6F78', '#5B6B80', '#8A6500'],
    track: '#E3E8F0',         // fondo de barras y medidores
    grid: '#E3E8F0',          // rejilla muy suave, solo horizontal
    axis: '#0B2545',
    text: '#0B2545',
};

const FONT = 'Inter, system-ui, -apple-system, "Segoe UI", sans-serif';
const nf = new Intl.NumberFormat('es-GT', { maximumFractionDigits: 2 });
const nfCompact = new Intl.NumberFormat('es-GT', { notation: 'compact', maximumFractionDigits: 1 });

export function fmt(v) {
    return v === null || v === undefined || !isFinite(v) ? '—' : nf.format(v);
}

function axisFmt(v) {
    return Math.abs(v) >= 10000 ? nfCompact.format(v) : nf.format(v);
}

/** Escalas "redondas" para el eje Y. */
function niceScale(min, max, ticks = 5) {
    if (min === max) { min = min === 0 ? 0 : min * 0.9; max = max === 0 ? 1 : max * 1.1; }
    const range = max - min;
    const rough = range / ticks;
    const mag = Math.pow(10, Math.floor(Math.log10(rough)));
    const norm = rough / mag;
    const step = (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
    const lo = Math.floor(min / step) * step;
    const hi = Math.ceil(max / step) * step;
    const out = [];
    for (let v = lo; v <= hi + step / 2; v += step) out.push(Math.round(v / step) * step);
    return { lo, hi, ticks: out };
}

function valueRange(spec) {
    let min = Infinity, max = -Infinity;
    if (spec.type === 'bar') {
        const stacks = {};
        spec.series.forEach((s) => {
            s.data.forEach((v, i) => {
                if (v === null || v === undefined) return;
                if (s.stack) {
                    const k = s.stack + '|' + i;
                    stacks[k] = (stacks[k] || 0) + v;
                    max = Math.max(max, stacks[k]);
                } else max = Math.max(max, v);
                min = Math.min(min, v);
            });
        });
        min = Math.min(0, min);
    } else {
        spec.series.concat(spec.band ? [{ data: spec.band.lo }, { data: spec.band.hi }] : []).forEach(s => s.data.forEach((v) => {
            if (v === null || v === undefined) return;
            min = Math.min(min, v); max = Math.max(max, v);
        }));
        if (spec.type === 'area' || (min > 0 && min < (max - min) * 0.6)) min = Math.min(0, min);
    }
    if (!isFinite(min)) { min = 0; max = 1; }
    return niceScale(min, max);
}

/** Dibuja el gráfico completo en un contexto 2D de w×h (px CSS). */
function draw(ctx, w, h, spec, { hover = null, exporting = false } = {}) {
    ctx.save();
    ctx.clearRect(0, 0, w, h);
    if (exporting) { ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, w, h); }
    const especial = { hbar: drawHbar, donut: drawDonut, gauge: drawGauge }[spec.type];
    if (especial) { const out = especial(ctx, w, h, spec, { hover, exporting }); ctx.restore(); return out; }

    const n = spec.labels.length;
    const fs = exporting ? 15 : 12;
    let top = exporting ? 20 : 8;

    if (exporting && spec.title) {
        ctx.fillStyle = PALETTE.text;
        ctx.font = `800 ${fs + 7}px ${FONT}`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        ctx.fillText(spec.title, w / 2, top);
        top += fs + 18;
        if (spec.subtitle) {
            ctx.fillStyle = PALETTE.axis;
            ctx.font = `500 ${fs}px ${FONT}`;
            ctx.fillText(spec.subtitle, w / 2, top);
            top += fs + 10;
        }
    }

    // Leyenda centrada (con salto de línea si no cabe).
    ctx.font = `600 ${fs}px ${FONT}`;
    ctx.textBaseline = 'middle';
    const leyenda = spec.band ? spec.series.concat([{ name: spec.band.name, color: spec.band.color, banda: true }]) : spec.series;
    const items = leyenda.map(s => ({ s, w: 26 + ctx.measureText(s.name).width + 18 }));
    const rows = [[]];
    let rowW = 0;
    items.forEach((it) => {
        if (rowW + it.w > w - 24 && rows[rows.length - 1].length) { rows.push([]); rowW = 0; }
        rows[rows.length - 1].push(it); rowW += it.w;
    });
    const lh = fs + 10;
    rows.forEach((row, ri) => {
        const total = row.reduce((a, it) => a + it.w, 0) - 18;
        let x = (w - total) / 2;
        const y = top + ri * lh + lh / 2;
        row.forEach(({ s, w: iw }) => {
            ctx.strokeStyle = s.color; ctx.fillStyle = s.color; ctx.lineWidth = 3;
            if (s.banda) {
                ctx.fillStyle = s.color + '40'; roundRect(ctx, x, y - 6, 18, 12, 3); ctx.fill();
            } else if (spec.type === 'bar' && !s.line) {
                roundRect(ctx, x, y - 6, 18, 12, 3); ctx.fill();
            } else {
                ctx.setLineDash(s.dashed ? [5, 4] : []);
                ctx.beginPath(); ctx.moveTo(x, y); ctx.lineTo(x + 18, y); ctx.stroke();
                ctx.setLineDash([]);
            }
            ctx.fillStyle = PALETTE.text;
            ctx.textAlign = 'left';
            ctx.fillText(s.name, x + 26, y);
            x += iw;
        });
    });
    top += rows.length * lh + 8;

    // Área del gráfico
    const { lo, hi, ticks } = valueRange(spec);
    ctx.font = `500 ${fs - 1}px ${FONT}`;
    const yLabelW = Math.max(...ticks.map(t => ctx.measureText(axisFmt(t)).width)) + 12;
    const L = Math.max(36, yLabelW), R = 14, B = fs + 22, T = top;
    const pw = Math.max(10, w - L - R), ph = Math.max(10, h - T - B);
    const y = v => T + ph - ((v - lo) / (hi - lo || 1)) * ph;
    const band = pw / Math.max(n, 1);
    const xc = i => L + band * (i + 0.5);

    // Cuadrícula y eje Y
    ctx.strokeStyle = PALETTE.grid; ctx.lineWidth = 1;
    ctx.fillStyle = PALETTE.axis; ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
    ticks.forEach((t) => {
        const yy = Math.round(y(t)) + 0.5;
        ctx.beginPath(); ctx.moveTo(L, yy); ctx.lineTo(L + pw, yy); ctx.stroke();
        ctx.fillText(axisFmt(t), L - 8, yy);
    });

    // Etiquetas del eje X (se saltan las que no caben)
    ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    const maxLabelW = Math.max(1, ...spec.labels.map(l => ctx.measureText(String(l)).width)) + 10;
    const step = Math.max(1, Math.ceil(maxLabelW / band));
    spec.labels.forEach((l, i) => {
        if (i % step !== 0 && i !== n - 1) return;
        if (i === n - 1 && i % step !== 0 && (n - 1) % step < step / 2) return;
        ctx.fillText(String(l), xc(i), T + ph + 8);
    });

    if (!n) {
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillStyle = PALETTE.axis; ctx.font = `600 ${fs + 1}px ${FONT}`;
        ctx.fillText('Sin datos en el rango seleccionado', L + pw / 2, T + ph / 2);
        ctx.restore();
        return { L, T, pw, ph, band, n };
    }

    if (hover !== null) {
        ctx.fillStyle = 'rgba(59, 167, 242, 0.10)';
        ctx.fillRect(L + band * hover, T, band, ph);
    }

    // Banda mín–máx (rango de referencia): relleno suave detrás de las series, por tramos sin huecos.
    if (spec.band) {
        const { lo: bLo, hi: bHi } = spec.band;
        let tramo = [];
        const cerrar = () => {
            if (tramo.length) {
                ctx.beginPath();
                tramo.forEach((i, k) => (k ? ctx.lineTo(xc(i), y(bHi[i])) : ctx.moveTo(xc(i), y(bHi[i]))));
                [...tramo].reverse().forEach(i => ctx.lineTo(xc(i), y(bLo[i])));
                ctx.closePath();
                ctx.fillStyle = spec.band.color + '33'; ctx.fill();
                if (tramo.length === 1) { ctx.strokeStyle = spec.band.color + '80'; ctx.lineWidth = band * 0.5; ctx.stroke(); }
            }
            tramo = [];
        };
        bLo.forEach((v, i) => { if (v === null || bHi[i] === null) cerrar(); else tramo.push(i); });
        cerrar();
    }

    // Barras
    const bars = spec.series.filter(s => spec.type === 'bar' && !s.line);
    if (bars.length) {
        const groups = [];
        bars.forEach((s) => {
            const g = s.stack ? 'stack:' + s.stack : 'solo:' + s.name;
            if (!groups.includes(g)) groups.push(g);
        });
        const inner = band * 0.72;
        const gw = inner / groups.length;
        const base = {};
        bars.forEach((s) => {
            const gi = groups.indexOf(s.stack ? 'stack:' + s.stack : 'solo:' + s.name);
            ctx.fillStyle = s.color;
            s.data.forEach((v, i) => {
                if (v === null || v === undefined) return;
                const k = gi + '|' + i;
                const b0 = base[k] || 0;
                const x = xc(i) - inner / 2 + gi * gw + gw * 0.08;
                const y1 = y(b0 + v), y0 = y(b0);
                roundRect(ctx, x, Math.min(y0, y1), Math.max(1, gw * 0.84), Math.max(1, Math.abs(y0 - y1)), Math.min(4, gw / 4));
                ctx.fill();
                if (s.stack) base[k] = b0 + v;
            });
        });
    }

    // Líneas (y áreas: relleno suave bajo la serie principal)
    const lineas = spec.series.filter(s => spec.type === 'line' || spec.type === 'area' || s.line);
    if (spec.type === 'area') {
        lineas.filter(s => !s.dashed).forEach((s) => {
            let tramo = [];
            const cerrar = () => {
                if (tramo.length > 1) {
                    ctx.beginPath();
                    ctx.moveTo(tramo[0][0], y(lo));
                    tramo.forEach(([px, py]) => ctx.lineTo(px, py));
                    ctx.lineTo(tramo[tramo.length - 1][0], y(lo));
                    ctx.closePath();
                    const grad = ctx.createLinearGradient(0, T, 0, T + ph);
                    grad.addColorStop(0, s.color + '55'); grad.addColorStop(1, s.color + '08');
                    ctx.fillStyle = grad; ctx.fill();
                }
                tramo = [];
            };
            s.data.forEach((v, i) => { if (v === null || v === undefined) cerrar(); else tramo.push([xc(i), y(v)]); });
            cerrar();
        });
    }
    lineas.forEach((s) => {
        ctx.strokeStyle = s.color; ctx.lineWidth = exporting ? 3 : 2.5;
        ctx.lineJoin = 'round'; ctx.lineCap = 'round';
        ctx.setLineDash(s.dashed ? [6, 5] : []);
        ctx.beginPath();
        let open = false;
        s.data.forEach((v, i) => {
            // spec.unir: une los puntos por encima de los huecos (muestreos en edades distintas).
            if (v === null || v === undefined) { if (!spec.unir) open = false; return; }
            if (!open) { ctx.moveTo(xc(i), y(v)); open = true; } else ctx.lineTo(xc(i), y(v));
        });
        ctx.stroke();
        ctx.setLineDash([]);
        if (n <= 60) {
            s.data.forEach((v, i) => {
                if (v === null || v === undefined) return;
                const lado = hover === i ? 9 : 6;   // puntos cuadrados, como los bloques del logo
                ctx.beginPath(); ctx.rect(xc(i) - lado / 2, y(v) - lado / 2, lado, lado);
                ctx.fillStyle = s.dashed ? '#FFFFFF' : s.color; ctx.fill();
                ctx.lineWidth = 2; ctx.strokeStyle = s.color; ctx.stroke();
            });
        }
    });

    if (hover !== null) drawTooltip(ctx, spec, hover, xc(hover), T, w, fs);
    ctx.restore();
    return { L, T, pw, ph, band, n };
}

function drawTooltip(ctx, spec, i, x, T, w, fs) {
    const lines = spec.series.map(s => ({ s, text: `${s.name}: ${fmt(s.data[i])}${s.data[i] !== null && s.data[i] !== undefined && spec.unit ? ' ' + spec.unit : ''}` }));
    if (spec.band && spec.band.lo[i] !== null && spec.band.hi[i] !== null) {
        lines.push({ s: spec.band, text: `${spec.band.name}: ${fmt(spec.band.lo[i])} – ${fmt(spec.band.hi[i])}${spec.unit ? ' ' + spec.unit : ''}` });
    }
    ctx.font = `700 ${fs}px ${FONT}`;
    const title = String(spec.labels[i]);
    let bw = ctx.measureText(title).width;
    ctx.font = `500 ${fs}px ${FONT}`;
    lines.forEach(l => { bw = Math.max(bw, ctx.measureText(l.text).width + 16); });
    bw += 20;
    const lh = fs + 7;
    const bh = 14 + lh * (lines.length + 1);
    let bx = x + 14;
    if (bx + bw > w - 4) bx = x - 14 - bw;
    bx = Math.max(4, bx);
    const by = T + 4;
    ctx.fillStyle = 'rgba(255,255,255,0.97)';
    ctx.shadowColor = 'rgba(20, 60, 20, 0.22)'; ctx.shadowBlur = 16; ctx.shadowOffsetY = 4;
    roundRect(ctx, bx, by, bw, bh, 8); ctx.fill();
    ctx.shadowColor = 'transparent';
    ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
    ctx.fillStyle = PALETTE.text; ctx.font = `700 ${fs}px ${FONT}`;
    ctx.fillText(title, bx + 10, by + 7 + lh / 2);
    ctx.font = `500 ${fs}px ${FONT}`;
    lines.forEach((l, k) => {
        const yy = by + 7 + lh * (k + 1) + lh / 2;
        ctx.fillStyle = l.s.color;
        ctx.beginPath(); ctx.arc(bx + 14, yy, 4, 0, Math.PI * 2); ctx.fill();
        ctx.fillStyle = PALETTE.text;
        ctx.fillText(l.text, bx + 24, yy);
    });
}

function roundRect(ctx, x, y, w, h, r) {
    r = Math.max(0, Math.min(r, w / 2, h / 2));
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
}

// ----------------------------------------------------------------------------
// Gráficos de "una cifra" o de partes: barras horizontales, dona y medidor.
// ----------------------------------------------------------------------------

/** Encabezado común (título y subtítulo solo al exportar). Devuelve el alto usado. */
function exportHeader(ctx, w, spec, fs, exporting) {
    let top = exporting ? 20 : 8;
    if (exporting && spec.title) {
        ctx.fillStyle = PALETTE.text; ctx.font = `800 ${fs + 7}px ${FONT}`;
        ctx.textAlign = 'center'; ctx.textBaseline = 'top';
        ctx.fillText(spec.title, w / 2, top);
        top += fs + 18;
        if (spec.subtitle) { ctx.fillStyle = PALETTE.axis; ctx.font = `500 ${fs}px ${FONT}`; ctx.fillText(spec.subtitle, w / 2, top); top += fs + 10; }
    }
    return top;
}

const HBAR = { rowH: 34, minH: 140 };
/** Alto necesario de un gráfico de barras horizontales (crece con el número de filas). */
export function chartHeight(spec, exporting = false) {
    if (spec.type !== 'hbar') return null;
    return Math.max(HBAR.minH, (exporting ? 110 : 40) + spec.labels.length * HBAR.rowH + (spec.target ? 28 : 0));
}

/**
 * Barras horizontales con el valor al final y, si hay `target` (p. ej. tabla), una marca
 * vertical gris por fila. spec = { type:'hbar', labels, values, color?, unit?, target?: [n], targetName? }
 */
function drawHbar(ctx, w, h, spec, { exporting = false } = {}) {
    const fs = exporting ? 15 : 12;
    let top = exportHeader(ctx, w, spec, fs, exporting);
    if (spec.target) {
        ctx.font = `600 ${fs}px ${FONT}`; ctx.textBaseline = 'middle'; ctx.textAlign = 'left';
        const t1 = spec.valueName || 'Real', t2 = spec.targetName || 'Tabla';
        const total = 18 + 8 + ctx.measureText(t1).width + 24 + 4 + 8 + ctx.measureText(t2).width;
        let x = (w - total) / 2;
        const yy = top + 10;
        ctx.fillStyle = spec.color || PALETTE.primary; roundRect(ctx, x, yy - 6, 18, 12, 3); ctx.fill();
        ctx.fillStyle = PALETTE.text; ctx.fillText(t1, x + 26, yy); x += 26 + ctx.measureText(t1).width + 24;
        ctx.fillStyle = PALETTE.tabla; ctx.fillRect(x, yy - 8, 3, 16);
        ctx.fillStyle = PALETTE.text; ctx.fillText(t2, x + 11, yy);
        top += 28;
    }
    const n = spec.labels.length;
    if (!n) {
        ctx.textAlign = 'center'; ctx.fillStyle = PALETTE.axis; ctx.font = `600 ${fs + 1}px ${FONT}`;
        ctx.fillText('Sin datos en el rango seleccionado', w / 2, h / 2);
        return { n: 0 };
    }
    ctx.font = `600 ${fs - 1}px ${FONT}`;
    const L = Math.min(w * 0.38, Math.max(...spec.labels.map(l => ctx.measureText(String(l)).width)) + 16);
    ctx.font = `700 ${fs - 1}px ${FONT}`;
    const fmtV = v => (v === null || v === undefined ? '—' : `${fmt(v)}${spec.unit ? ' ' + spec.unit : ''}`);
    const R = Math.max(...spec.values.map(v => ctx.measureText(fmtV(v)).width)) + 14;
    const pw = Math.max(10, w - L - R - 8);
    const max = Math.max(1e-9, ...spec.values.map(v => v || 0), ...(spec.target || []).map(v => v || 0)) * 1.05;
    const rowH = Math.min(HBAR.rowH, (h - top - 8) / n);
    spec.labels.forEach((lab, i) => {
        const yy = top + i * rowH;
        const v = spec.values[i];
        ctx.fillStyle = PALETTE.text; ctx.font = `600 ${fs - 1}px ${FONT}`; ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
        ctx.fillText(String(lab), L - 10, yy + rowH / 2);
        ctx.fillStyle = PALETTE.track; roundRect(ctx, L, yy + rowH * 0.2, pw, rowH * 0.6, 4); ctx.fill();
        if (v !== null && v !== undefined) {
            ctx.fillStyle = spec.colors?.[i] || spec.color || PALETTE.primary;
            roundRect(ctx, L, yy + rowH * 0.2, Math.max(2, pw * (v / max)), rowH * 0.6, 4); ctx.fill();
        }
        const t = spec.target?.[i];
        if (t !== null && t !== undefined) {
            ctx.fillStyle = PALETTE.tabla;
            ctx.fillRect(L + pw * (t / max) - 1.5, yy + rowH * 0.1, 3, rowH * 0.8);
        }
        ctx.fillStyle = PALETTE.text; ctx.font = `700 ${fs - 1}px ${FONT}`; ctx.textAlign = 'left';
        ctx.fillText(fmtV(v), L + pw + 8, yy + rowH / 2);
    });
    return { n: 0 };
}

/**
 * Dona con total al centro y leyenda con valor y porcentaje.
 * spec = { type:'donut', labels, values, colors, centerLabel?, unit? }
 */
function drawDonut(ctx, w, h, spec, { exporting = false, hover = null } = {}) {
    const fs = exporting ? 15 : 12;
    const top = exportHeader(ctx, w, spec, fs, exporting);
    const vals = spec.values.map(v => Math.max(0, v || 0));
    const total = vals.reduce((a, b) => a + b, 0);
    if (!total) {
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillStyle = PALETTE.axis; ctx.font = `600 ${fs + 1}px ${FONT}`;
        ctx.fillText('Sin datos en el rango seleccionado', w / 2, (h + top) / 2);
        return { n: 0 };
    }
    // Leyenda a la derecha si hay espacio; si no, debajo.
    ctx.font = `600 ${fs}px ${FONT}`;
    const items = spec.labels.map((l, i) => ({ l: String(l), v: vals[i], c: spec.colors[i], p: vals[i] / total * 100 }));
    const lineTxt = it => `${it.l} · ${fmt(it.v)} (${fmt(Math.round(it.p * 10) / 10)} %)`;
    const legendW = Math.max(...items.map(it => ctx.measureText(lineTxt(it)).width)) + 30;
    const lado = w >= 420 && legendW < w * 0.55;
    const lh = fs + 10;
    const legendH = items.length * lh;
    const areaH = lado ? h - top - 8 : h - top - legendH - 16;
    const areaW = lado ? w - legendW - 16 : w;
    const r = Math.max(30, Math.min(areaW / 2 - 10, areaH / 2 - 4));
    const cx = lado ? Math.max(r + 10, (areaW) / 2) : w / 2;
    const cy = top + Math.max(r + 4, areaH / 2);
    const inner = r * 0.6;
    let a = -Math.PI / 2;
    vals.forEach((v, i) => {
        if (!v) return;
        const da = v / total * Math.PI * 2;
        ctx.beginPath();
        ctx.arc(cx, cy, hover === i ? r + 4 : r, a, a + da);
        ctx.arc(cx, cy, inner, a + da, a, true);
        ctx.closePath();
        ctx.fillStyle = spec.colors[i]; ctx.fill();
        ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = 2; ctx.stroke();
        a += da;
    });
    ctx.fillStyle = PALETTE.text; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    // El total se achica hasta caber dentro del hueco de la dona.
    let tam = fs + 8;
    ctx.font = `800 ${tam}px ${FONT}`;
    while (tam > 10 && ctx.measureText(fmt(total)).width > inner * 1.6) { tam -= 1; ctx.font = `800 ${tam}px ${FONT}`; }
    ctx.fillText(fmt(total), cx, cy - (spec.centerLabel ? 8 : 0));
    if (spec.centerLabel) { ctx.fillStyle = PALETTE.axis; ctx.font = `600 ${fs - 1}px ${FONT}`; ctx.fillText(spec.centerLabel, cx, cy + fs + 2); }
    const lx = lado ? areaW + 8 : (w - legendW) / 2;
    const ly = lado ? cy - legendH / 2 : top + areaH + 12;
    ctx.textAlign = 'left'; ctx.font = `600 ${fs}px ${FONT}`;
    items.forEach((it, i) => {
        const yy = ly + i * lh + lh / 2;
        ctx.fillStyle = it.c; roundRect(ctx, lx, yy - 6, 14, 12, 3); ctx.fill();
        ctx.fillStyle = PALETTE.text; ctx.font = `${hover === i ? 800 : 600} ${fs}px ${FONT}`;
        ctx.fillText(lineTxt(it), lx + 22, yy);
    });
    return { n: 0, donut: { cx, cy, r, inner, vals, total } };
}

/**
 * Medidor semicircular: valor contra un máximo, con marca opcional de referencia (tabla).
 * spec = { type:'gauge', value, max, min?, target?, targetLabel?, unit?, caption? }
 */
function drawGauge(ctx, w, h, spec, { exporting = false } = {}) {
    const fs = exporting ? 15 : 12;
    const top = exportHeader(ctx, w, spec, fs, exporting);
    const min = spec.min || 0;
    const max = spec.max || 100;
    const capH = spec.caption ? fs + 12 : 0;
    const r = Math.max(40, Math.min(w / 2 - 24, h - top - 46 - capH));
    const cx = w / 2;
    const cy = top + r + 10;
    const grosor = Math.max(14, r * 0.22);
    const frac = v => Math.max(0, Math.min(1, ((v ?? min) - min) / (max - min)));
    ctx.lineCap = 'butt';
    ctx.beginPath(); ctx.arc(cx, cy, r - grosor / 2, Math.PI, 2 * Math.PI);
    ctx.strokeStyle = PALETTE.track; ctx.lineWidth = grosor; ctx.stroke();
    if (spec.value !== null && spec.value !== undefined) {
        ctx.beginPath(); ctx.arc(cx, cy, r - grosor / 2, Math.PI, Math.PI + Math.PI * frac(spec.value));
        ctx.strokeStyle = spec.color || PALETTE.primary; ctx.stroke();
    }
    if (spec.target !== null && spec.target !== undefined) {
        const ang = Math.PI + Math.PI * frac(spec.target);
        ctx.beginPath();
        ctx.moveTo(cx + Math.cos(ang) * (r - grosor - 6), cy + Math.sin(ang) * (r - grosor - 6));
        ctx.lineTo(cx + Math.cos(ang) * (r + 6), cy + Math.sin(ang) * (r + 6));
        ctx.strokeStyle = PALETTE.text; ctx.lineWidth = 3; ctx.stroke();
    }
    ctx.fillStyle = PALETTE.text; ctx.textAlign = 'center'; ctx.textBaseline = 'alphabetic';
    ctx.font = `800 ${Math.round(r * 0.32)}px ${FONT}`;
    ctx.fillText(spec.value === null || spec.value === undefined ? '—' : `${fmt(spec.value)}${spec.unit || ''}`, cx, cy - 4);
    ctx.font = `600 ${fs - 1}px ${FONT}`; ctx.fillStyle = PALETTE.axis; ctx.textBaseline = 'top';
    ctx.fillText(fmt(min), cx - r + grosor / 2, cy + 6);
    ctx.fillText(fmt(max), cx + r - grosor / 2, cy + 6);
    let yy = cy + 6;
    if (spec.target !== null && spec.target !== undefined) {
        ctx.fillStyle = PALETTE.text; ctx.font = `600 ${fs}px ${FONT}`;
        ctx.fillText(`▍ ${spec.targetLabel || 'Referencia'}: ${fmt(spec.target)}${spec.unit || ''}`, cx, yy);
        yy += fs + 8;
    }
    if (spec.caption) { ctx.fillStyle = PALETTE.axis; ctx.font = `500 ${fs - 1}px ${FONT}`; ctx.fillText(spec.caption, cx, yy + 2); }
    return { n: 0 };
}

/** Monta un gráfico interactivo en un <canvas>. Devuelve una función para desmontarlo. */
export function mountChart(canvas, spec) {
    let hover = null;
    let layout = null;
    const box = canvas.parentElement;

    function paint() {
        const dpr = window.devicePixelRatio || 1;
        const w = box.clientWidth;
        const alto = chartHeight(spec);
        if (alto) canvas.style.height = alto + 'px';
        const h = alto || canvas.clientHeight || 280;
        if (!w) return;
        canvas.width = Math.round(w * dpr);
        canvas.height = Math.round(h * dpr);
        const ctx = canvas.getContext('2d');
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        layout = draw(ctx, w, h, spec, { hover });
    }

    function pick(ev) {
        if (layout?.donut) {
            // Dona: resalta el segmento bajo el cursor.
            const r = canvas.getBoundingClientRect();
            const d = layout.donut;
            const dx = ev.clientX - r.left - d.cx, dy = ev.clientY - r.top - d.cy;
            const dist = Math.hypot(dx, dy);
            let next = null;
            if (dist >= d.inner && dist <= d.r + 4) {
                let ang = Math.atan2(dy, dx) + Math.PI / 2;
                if (ang < 0) ang += Math.PI * 2;
                let acc = 0;
                next = d.vals.findIndex((v) => { acc += v / d.total * Math.PI * 2; return ang <= acc; });
            }
            if (next !== hover) { hover = next; paint(); }
            return;
        }
        if (!layout || !layout.n) return;
        const r = canvas.getBoundingClientRect();
        const x = ev.clientX - r.left;
        const i = Math.floor((x - layout.L) / layout.band);
        const next = i >= 0 && i < layout.n ? i : null;
        if (next !== hover) { hover = next; paint(); }
    }
    const leave = () => { if (hover !== null) { hover = null; paint(); } };
    canvas.addEventListener('pointermove', pick);
    canvas.addEventListener('pointerdown', pick);
    canvas.addEventListener('pointerleave', leave);

    const ro = new ResizeObserver(() => paint());
    ro.observe(box);
    paint();
    canvas.setAttribute('role', 'img');
    canvas.setAttribute('aria-label', spec.title || 'Gráfico');
    return () => { ro.disconnect(); };
}

// ----------------------------------------------------------------------------
// Mapa de calor (filas × columnas). spec = { type:'heatmap', title, subtitle?, rows:[nombres],
//   cols:[etiquetas], values:[[número|null]], color(v) → color, legend:[{color, label}],
//   tip(r, c) → [líneas de texto] }
// ----------------------------------------------------------------------------

const HM = { rowH: 30, minCell: 10 };

function heatmapLayout(ctx, w, spec, exporting) {
    const fs = exporting ? 15 : 12;
    let top = exporting ? 20 : 8;
    if (exporting && spec.title) top += fs + 18 + (spec.subtitle ? fs + 10 : 0);
    // Leyenda (puede ocupar varias filas)
    ctx.font = `600 ${fs}px ${FONT}`;
    const items = spec.legend.map(it => ({ it, w: 22 + ctx.measureText(it.label).width + 18 }));
    const rows = [[]];
    let rw = 0;
    items.forEach((x) => {
        if (rw + x.w > w - 24 && rows[rows.length - 1].length) { rows.push([]); rw = 0; }
        rows[rows.length - 1].push(x); rw += x.w;
    });
    const legendTop = top;
    const lh = fs + 10;
    top += rows.length * lh + 8;
    ctx.font = `600 ${fs - 1}px ${FONT}`;
    const L = Math.min(w * 0.38, Math.max(60, ...spec.rows.map(r => ctx.measureText(String(r)).width)) + 16);
    const R = 12;
    const pw = Math.max(10, w - L - R);
    const cw = pw / Math.max(spec.cols.length, 1);
    const rowH = HM.rowH;
    const B = fs + 22;
    const height = top + Math.max(spec.rows.length, 1) * rowH + B;
    return { fs, legendTop, legendRows: rows, lh, top, L, pw, cw, rowH, height };
}

/** Alto en px CSS que necesita el mapa de calor para un ancho dado. */
export function heatmapHeight(spec, w, exporting = false) {
    const c = document.createElement('canvas').getContext('2d');
    return heatmapLayout(c, w, spec, exporting).height;
}

function drawHeatmap(ctx, w, h, spec, { hover = null, exporting = false } = {}) {
    ctx.save();
    ctx.clearRect(0, 0, w, h);
    if (exporting) { ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, w, h); }
    const lay = heatmapLayout(ctx, w, spec, exporting);
    const { fs, L, pw, cw, rowH, top } = lay;

    if (exporting && spec.title) {
        ctx.fillStyle = PALETTE.text; ctx.font = `800 ${fs + 7}px ${FONT}`;
        ctx.textAlign = 'center'; ctx.textBaseline = 'top';
        ctx.fillText(spec.title, w / 2, 20);
        if (spec.subtitle) { ctx.fillStyle = PALETTE.axis; ctx.font = `500 ${fs}px ${FONT}`; ctx.fillText(spec.subtitle, w / 2, 20 + fs + 18); }
    }
    // Leyenda centrada
    ctx.font = `600 ${fs}px ${FONT}`; ctx.textBaseline = 'middle';
    lay.legendRows.forEach((row, ri) => {
        const total = row.reduce((a, x) => a + x.w, 0) - 18;
        let x = (w - total) / 2;
        const y = lay.legendTop + ri * lay.lh + lay.lh / 2;
        row.forEach(({ it, w: iw }) => {
            ctx.fillStyle = it.color; roundRect(ctx, x, y - 7, 14, 14, 3); ctx.fill();
            ctx.fillStyle = PALETTE.text; ctx.textAlign = 'left'; ctx.fillText(it.label, x + 22, y);
            x += iw;
        });
    });

    if (!spec.rows.length || !spec.cols.length) {
        ctx.textAlign = 'center'; ctx.fillStyle = PALETTE.axis; ctx.font = `600 ${fs + 1}px ${FONT}`;
        ctx.fillText('Sin lecturas en el rango seleccionado', w / 2, top + rowH);
        ctx.restore();
        return lay;
    }

    // Etiquetas de filas
    ctx.font = `600 ${fs - 1}px ${FONT}`; ctx.textAlign = 'right'; ctx.fillStyle = PALETTE.text;
    spec.rows.forEach((r, i) => ctx.fillText(String(r), L - 10, top + i * rowH + rowH / 2));

    // Celdas
    const gap = cw > 6 ? 2 : 0.5;
    spec.values.forEach((fila, i) => fila.forEach((v, j) => {
        ctx.fillStyle = v === null || v === undefined ? PALETTE.track : spec.color(v);
        roundRect(ctx, L + j * cw + gap / 2, top + i * rowH + 2, Math.max(1, cw - gap), rowH - 4, Math.min(5, cw / 3));
        ctx.fill();
    }));
    if (hover) {
        ctx.strokeStyle = PALETTE.text; ctx.lineWidth = 2;
        roundRect(ctx, L + hover.c * cw + 1, top + hover.r * rowH + 1, Math.max(2, cw - 2), rowH - 2, 4);
        ctx.stroke();
    }

    // Etiquetas de columnas (se saltan las que no caben)
    ctx.fillStyle = PALETTE.axis; ctx.font = `500 ${fs - 1}px ${FONT}`; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    const maxW = Math.max(1, ...spec.cols.map(c => ctx.measureText(String(c)).width)) + 10;
    const step = Math.max(1, Math.ceil(maxW / cw));
    const yLab = top + spec.rows.length * rowH + 8;
    spec.cols.forEach((c, j) => { if (j % step === 0) ctx.fillText(String(c), L + j * cw + cw / 2, yLab); });

    if (hover && spec.tip) {
        const lines = spec.tip(hover.r, hover.c);
        ctx.font = `600 ${fs}px ${FONT}`;
        const bw = Math.max(...lines.map(l => ctx.measureText(l).width)) + 24;
        const bh = 12 + lines.length * (fs + 7);
        let bx = L + hover.c * cw + cw + 10;
        if (bx + bw > w - 4) bx = L + hover.c * cw - bw - 10;
        bx = Math.max(4, bx);
        let by = top + hover.r * rowH - bh / 2 + rowH / 2;
        by = Math.max(4, Math.min(by, h - bh - 4));
        ctx.fillStyle = 'rgba(255,255,255,0.97)';
        ctx.shadowColor = 'rgba(11, 61, 145, 0.22)'; ctx.shadowBlur = 16; ctx.shadowOffsetY = 4;
        roundRect(ctx, bx, by, bw, bh, 10); ctx.fill();
        ctx.shadowColor = 'transparent';
        ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillStyle = PALETTE.text;
        lines.forEach((l, k) => { ctx.font = `${k === 0 ? 700 : 500} ${fs}px ${FONT}`; ctx.fillText(l, bx + 12, by + 6 + k * (fs + 7) + (fs + 7) / 2); });
    }
    ctx.restore();
    return lay;
}

/** Monta un mapa de calor interactivo. El alto del canvas se ajusta al número de filas. */
export function mountHeatmap(canvas, spec) {
    let hover = null;
    let lay = null;
    const box = canvas.parentElement;
    function paint() {
        const dpr = window.devicePixelRatio || 1;
        const w = box.clientWidth;
        if (!w) return;
        const h = heatmapHeight(spec, w);
        canvas.style.height = h + 'px';
        canvas.width = Math.round(w * dpr);
        canvas.height = Math.round(h * dpr);
        const ctx = canvas.getContext('2d');
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        lay = drawHeatmap(ctx, w, h, spec, { hover });
    }
    function pick(ev) {
        if (!lay) return;
        const r = canvas.getBoundingClientRect();
        const c = Math.floor((ev.clientX - r.left - lay.L) / lay.cw);
        const f = Math.floor((ev.clientY - r.top - lay.top) / lay.rowH);
        const next = c >= 0 && c < spec.cols.length && f >= 0 && f < spec.rows.length ? { r: f, c } : null;
        if (JSON.stringify(next) !== JSON.stringify(hover)) { hover = next; paint(); }
    }
    const leave = () => { if (hover) { hover = null; paint(); } };
    canvas.addEventListener('pointermove', pick);
    canvas.addEventListener('pointerdown', pick);
    canvas.addEventListener('pointerleave', leave);
    const ro = new ResizeObserver(() => paint());
    ro.observe(box);
    paint();
    canvas.setAttribute('role', 'img');
    canvas.setAttribute('aria-label', spec.title || 'Mapa de calor');
    return () => ro.disconnect();
}

/** Descarga el gráfico como PNG (1400×720, alta resolución, fondo blanco). */
export function exportChartPng(spec, filename) {
    const w = 1400, scale = 2;
    const h = spec.type === 'heatmap' ? heatmapHeight(spec, w, true)
        : spec.type === 'hbar' ? Math.max(520, chartHeight(spec, true))
            : spec.type === 'gauge' || spec.type === 'donut' ? 640 : 720;
    const canvas = document.createElement('canvas');
    canvas.width = w * scale; canvas.height = h * scale;
    const ctx = canvas.getContext('2d');
    ctx.scale(scale, scale);
    if (spec.type === 'heatmap') drawHeatmap(ctx, w, h, spec, { exporting: true });
    else draw(ctx, w, h, spec, { exporting: true });
    return new Promise((resolve) => {
        canvas.toBlob((blob) => {
            if (blob) downloadBlob(blob, filename);
            resolve(!!blob);
        }, 'image/png');
    });
}

export function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.rel = 'noopener';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
}
