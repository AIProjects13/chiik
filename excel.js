// ============================================================================
// Excel (.xlsx) sin librerías: escribir (descargas y plantillas) y leer (carga desde Excel).
// Un .xlsx es un ZIP con archivos XML. Para escribir se guardan sin comprimir; para leer se
// descomprime con DecompressionStream, que traen los navegadores.
// ============================================================================

// ---------- ZIP ----------
const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
})();
function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
}
/** archivos: [{ nombre, texto }] → Blob ZIP (método "store"). */
function zip(archivos) {
    const enc = new TextEncoder();
    const partes = [], central = [];
    let offset = 0;
    archivos.forEach(({ nombre, texto }) => {
        const nom = enc.encode(nombre), datos = enc.encode(texto), crc = crc32(datos);
        const loc = new DataView(new ArrayBuffer(30));
        loc.setUint32(0, 0x04034b50, true); loc.setUint16(4, 20, true); loc.setUint16(6, 0x0800, true);
        loc.setUint32(14, crc, true); loc.setUint32(18, datos.length, true); loc.setUint32(22, datos.length, true);
        loc.setUint16(26, nom.length, true);
        partes.push(loc, nom, datos);
        const cen = new DataView(new ArrayBuffer(46));
        cen.setUint32(0, 0x02014b50, true); cen.setUint16(4, 20, true); cen.setUint16(6, 20, true); cen.setUint16(8, 0x0800, true);
        cen.setUint32(16, crc, true); cen.setUint32(20, datos.length, true); cen.setUint32(24, datos.length, true);
        cen.setUint16(28, nom.length, true); cen.setUint32(42, offset, true);
        central.push(cen, nom);
        offset += 30 + nom.length + datos.length;
    });
    const tamCentral = central.reduce((a, p) => a + p.byteLength, 0);
    const fin = new DataView(new ArrayBuffer(22));
    fin.setUint32(0, 0x06054b50, true); fin.setUint16(8, archivos.length, true); fin.setUint16(10, archivos.length, true);
    fin.setUint32(12, tamCentral, true); fin.setUint32(16, offset, true);
    return new Blob([...partes, ...central, fin], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
async function unzip(buffer) {
    const v = new DataView(buffer);
    let fin = -1;
    for (let i = buffer.byteLength - 22; i >= Math.max(0, buffer.byteLength - 66000); i--) if (v.getUint32(i, true) === 0x06054b50) { fin = i; break; }
    if (fin < 0) throw new Error('El archivo no es un Excel (.xlsx) válido.');
    const n = v.getUint16(fin + 10, true);
    let p = v.getUint32(fin + 16, true);
    const dec = new TextDecoder();
    const out = {};
    for (let k = 0; k < n; k++) {
        const metodo = v.getUint16(p + 10, true), comp = v.getUint32(p + 20, true);
        const lnom = v.getUint16(p + 28, true), lext = v.getUint16(p + 30, true), lcom = v.getUint16(p + 32, true);
        const local = v.getUint32(p + 42, true);
        const nombre = dec.decode(new Uint8Array(buffer, p + 46, lnom));
        const ini = local + 30 + v.getUint16(local + 26, true) + v.getUint16(local + 28, true);
        const datos = new Uint8Array(buffer, ini, comp);
        out[nombre] = { metodo, datos };
        p += 46 + lnom + lext + lcom;
    }
    return {
        async texto(nombre) {
            const f = out[nombre];
            if (!f) return null;
            if (f.metodo === 0) return dec.decode(f.datos);
            if (f.metodo !== 8) throw new Error('El Excel usa una compresión que no se puede leer.');
            const s = new Blob([f.datos]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
            return dec.decode(await new Response(s).arrayBuffer());
        },
    };
}

// ---------- Escribir ----------
const xml = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
function colLetra(i) { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
/** 'AAAA-MM-DD' o 'AAAA-MM-DD HH:MM[:SS]' → número de serie de Excel (o null). */
function serieExcel(v) {
    const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
    if (!m) return null;
    const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
    return (ms - Date.UTC(1899, 11, 30)) / 86400000;
}
// Estilos (cellXfs): 0 normal · 1 encabezado · 2 título · 3 nota · 4 fecha · 5 fecha y hora · 6–10 número con 0–4 decimales · 11 texto largo · 12 celda para llenar
const ESTILOS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="7"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy hh:mm"/>
<numFmt numFmtId="166" formatCode="#,##0"/><numFmt numFmtId="167" formatCode="#,##0.0"/><numFmt numFmtId="168" formatCode="#,##0.00"/>
<numFmt numFmtId="169" formatCode="#,##0.000"/><numFmt numFmtId="170" formatCode="#,##0.0000"/></numFmts>
<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="14"/><color rgb="FF0B2545"/><name val="Calibri"/></font><font><i/><sz val="10"/><color rgb="FF5B6B80"/><name val="Calibri"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0B2545"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF3F5F9"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFE3E8F0"/></left><right style="thin"><color rgb="FFE3E8F0"/></right><top style="thin"><color rgb="FFE3E8F0"/></top><bottom style="thin"><color rgb="FFE3E8F0"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="13">
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" applyFont="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="167" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="168" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="169" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="170" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" applyFill="1" applyBorder="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

function decimalesDe(valores) {
    let d = 0;
    valores.forEach((v) => { const m = String(v).match(/\.(\d+)/); if (m) d = Math.max(d, Math.min(4, m[1].length)); });
    return d;
}
function celda(ref, v, t, estiloNum) {
    if (v === null || v === undefined || v === '') return '';
    if (t === 'num') {
        const n = typeof v === 'number' ? v : Number(String(v).replace(/,/g, ''));
        if (Number.isFinite(n)) return `<c r="${ref}" s="${estiloNum}"><v>${n}</v></c>`;
    }
    if (t === 'date' || t === 'datetime') {
        const s = serieExcel(v);
        if (s !== null) return `<c r="${ref}" s="${t === 'date' ? 4 : 5}"><v>${s}</v></c>`;
    }
    return `<c r="${ref}" s="${t === 'wrap' ? 11 : 0}" t="inlineStr"><is><t xml:space="preserve">${xml(v)}</t></is></c>`;
}

/**
 * hoja: { nombre, titulo?: [líneas], columnas: [{ l, t: 'text'|'num'|'date'|'datetime'|'wrap', ancho? }],
 *         filas: [[valor por columna]], filasVacias?: n (para llenar), validaciones?: [{ col, tipo: 'lista'|'entero'|'decimal'|'fecha', lista?: 'Hoja!$A$1:$A$9', min?, max?, ayuda? }],
 *         oculta?: bool, simple?: bool (sin título ni filtros: hojas de listas) }
 */
function hojaXml(h) {
    const titulo = h.simple ? [] : (h.titulo || []);
    const filaEnc = titulo.length ? titulo.length + 2 : 1;   // una fila en blanco entre el título y los encabezados
    const estilosNum = h.columnas.map((c, j) => (c.t === 'num' ? 6 + decimalesDe(h.filas.map(f => f[j]).filter(v => v !== '' && v !== null && v !== undefined)) : 0));
    const filas = [];
    titulo.forEach((l, i) => filas.push(`<row r="${i + 1}"><c r="A${i + 1}" s="${i ? 3 : 2}" t="inlineStr"><is><t xml:space="preserve">${xml(l)}</t></is></c></row>`));
    if (!h.simple) filas.push(`<row r="${filaEnc}" ht="32" customHeight="1">${h.columnas.map((c, j) => `<c r="${colLetra(j)}${filaEnc}" s="1" t="inlineStr"><is><t xml:space="preserve">${xml(c.l)}</t></is></c>`).join('')}</row>`);
    const primera = h.simple ? 1 : filaEnc + 1;
    h.filas.forEach((f, i) => {
        const r = primera + i;
        filas.push(`<row r="${r}">${f.map((v, j) => celda(`${colLetra(j)}${r}`, v, h.columnas[j]?.t, estilosNum[j])).join('')}</row>`);
    });
    // Filas para llenar (plantillas): no se escriben celdas; solo alargan validaciones y filtros.
    const vacias = h.filasVacias || 0;
    const ultima = primera + h.filas.length + vacias - 1;
    const anchos = h.columnas.map((c, j) => {
        if (c.ancho) return c.ancho;
        const largo = Math.max(String(c.l).length * 0.9, ...h.filas.slice(0, 300).map(f => String(f[j] ?? '').length));
        return Math.min(c.t === 'wrap' ? 60 : 40, Math.max(c.t === 'date' ? 12 : 9, largo + 2));
    });
    const cols = `<cols>${anchos.map((w, j) => `<col min="${j + 1}" max="${j + 1}" width="${w.toFixed(1)}" customWidth="1"/>`).join('')}</cols>`;
    const vista = h.simple ? '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        : `<sheetViews><sheetView workbookViewId="0"${h.activa ? ' tabSelected="1"' : ''}><pane ySplit="${filaEnc}" topLeftCell="A${filaEnc + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`;
    const ultCol = colLetra(h.columnas.length - 1);
    const filtro = !h.simple && ultima >= filaEnc ? `<autoFilter ref="A${filaEnc}:${ultCol}${Math.max(ultima, filaEnc)}"/>` : '';
    const vals = (h.validaciones || []).map((v) => {
        const sq = `${colLetra(v.col)}${primera}:${colLetra(v.col)}${Math.max(ultima, primera)}`;
        const ayuda = v.ayuda ? ` showInputMessage="1" promptTitle="${xml(String(v.ayudaTitulo || 'Dato').slice(0, 32))}" prompt="${xml(String(v.ayuda).slice(0, 250))}"` : '';
        const err = ` showErrorMessage="1" errorTitle="Dato no válido" error="${xml(v.error || 'Revise el valor de esta celda.')}"`;
        if (v.tipo === 'lista') return `<dataValidation type="list" allowBlank="1"${ayuda}${err} sqref="${sq}"><formula1>${xml(v.lista)}</formula1></dataValidation>`;
        const tipo = v.tipo === 'fecha' ? 'date' : v.tipo === 'entero' ? 'whole' : 'decimal';
        const f = x => (v.tipo === 'fecha' ? serieExcel(x) : x);
        return `<dataValidation type="${tipo}" operator="between" allowBlank="1"${ayuda}${err} sqref="${sq}"><formula1>${f(v.min)}</formula1><formula2>${f(v.max)}</formula2></dataValidation>`;
    });
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${vista}<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${filas.join('')}</sheetData>${filtro}${vals.length ? `<dataValidations count="${vals.length}">${vals.join('')}</dataValidations>` : ''}
<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/></worksheet>`;
}

/** hojas: [hoja] (ver hojaXml) → Blob .xlsx */
export function crearXlsx(hojas) {
    const nombres = hojas.map((h, i) => String(h.nombre || `Hoja ${i + 1}`).replace(/[\\/?*[\]:]/g, ' ').slice(0, 31));
    const definidos = hojas.map((h, i) => {
        if (h.simple) return '';
        const titulo = (h.titulo || []).length;
        const enc = titulo ? titulo + 2 : 1;
        const ult = enc + h.filas.length + (h.filasVacias || 0);
        return `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${nombres[i].replace(/'/g, "''")}'!$A$${enc}:$${colLetra(h.columnas.length - 1)}$${Math.max(ult, enc)}</definedName>`;
    }).join('');
    const archivos = [
        { nombre: '[Content_Types].xml', texto: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${hojas.map((h, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>` },
        { nombre: '_rels/.rels', texto: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
        { nombre: 'xl/workbook.xml', texto: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews>
<sheets>${hojas.map((h, i) => `<sheet name="${xml(nombres[i])}" sheetId="${i + 1}"${h.oculta ? ' state="hidden"' : ''} r:id="rId${i + 1}"/>`).join('')}</sheets>${definidos ? `<definedNames>${definidos}</definedNames>` : ''}</workbook>` },
        { nombre: 'xl/_rels/workbook.xml.rels', texto: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hojas.map((h, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}
<Relationship Id="rId${hojas.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
        { nombre: 'xl/styles.xml', texto: ESTILOS },
    ].concat(hojas.map((h, i) => ({ nombre: `xl/worksheets/sheet${i + 1}.xml`, texto: hojaXml(h) })));
    return zip(archivos);
}

// ---------- Leer ----------
/** Número de serie de Excel → 'AAAA-MM-DD'. */
export function fechaDeSerie(n) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(Number(n) * 86400000));
    return d.toISOString().slice(0, 10);
}
/** Serie con formato de fecha u hora → 'AAAA-MM-DD', 'HH:MM' (solo hora) o 'AAAA-MM-DD HH:MM'. */
function fechaHoraDeSerie(v) {
    const n = Number(v);
    const min = Math.round((n % 1) * 1440);
    const hora = `${String(Math.floor(min / 60) % 24).padStart(2, '0')}:${String(min % 60).padStart(2, '0')}`;
    if (n < 1) return hora;
    return min ? `${fechaDeSerie(Math.floor(n))} ${hora}` : fechaDeSerie(n);
}
function refCol(ref) { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }

/**
 * Lee un .xlsx → { hojas: [{ nombre, filas: [[valor]] }] }. Los números salen como número; el
 * texto como texto; las fechas con formato de fecha, como 'AAAA-MM-DD'.
 */
export async function leerXlsx(archivo) {
    const z = await unzip(await archivo.arrayBuffer());
    const dom = s => new DOMParser().parseFromString(s, 'application/xml');
    const wb = await z.texto('xl/workbook.xml');
    if (!wb) throw new Error('El archivo no es un Excel (.xlsx) válido.');
    const rels = dom(await z.texto('xl/_rels/workbook.xml.rels') || '<r/>');
    const destino = {};
    [...rels.getElementsByTagName('Relationship')].forEach(r => { destino[r.getAttribute('Id')] = r.getAttribute('Target'); });
    const compartidos = [];
    const ss = await z.texto('xl/sharedStrings.xml');
    if (ss) [...dom(ss).getElementsByTagName('si')].forEach(si => compartidos.push([...si.getElementsByTagName('t')].map(t => t.textContent).join('')));
    // Estilos con formato de fecha (para convertir números de serie a fecha).
    const esFecha = [];
    const st = await z.texto('xl/styles.xml');
    if (st) {
        const d = dom(st);
        const propios = {};
        [...d.getElementsByTagName('numFmt')].forEach(f => { propios[f.getAttribute('numFmtId')] = f.getAttribute('formatCode'); });
        const xfs = d.getElementsByTagName('cellXfs')[0];
        if (xfs) [...xfs.getElementsByTagName('xf')].forEach((xf, i) => {
            const id = Number(xf.getAttribute('numFmtId'));
            const cod = propios[id] || '';
            esFecha[i] = (id >= 14 && id <= 22) || (id >= 45 && id <= 47) || (/[dmy]/i.test(cod.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/0/.test(cod));
        });
    }
    const hojas = [];
    for (const s of [...dom(wb).getElementsByTagName('sheet')]) {
        const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
        let ruta = destino[rid] || '';
        ruta = ruta.startsWith('/') ? ruta.slice(1) : 'xl/' + ruta.replace(/^\.\//, '');
        const texto = await z.texto(ruta);
        if (!texto) continue;
        const filas = [];
        [...dom(texto).getElementsByTagName('row')].forEach((row) => {
            const r = Number(row.getAttribute('r')) - 1;
            const fila = [];
            [...row.getElementsByTagName('c')].forEach((c) => {
                const j = refCol(c.getAttribute('r'));
                const tipo = c.getAttribute('t');
                const v = c.getElementsByTagName('v')[0]?.textContent;
                let val;
                if (tipo === 's') val = compartidos[Number(v)] ?? '';
                else if (tipo === 'inlineStr') val = [...c.getElementsByTagName('t')].map(t => t.textContent).join('');
                else if (tipo === 'str' || tipo === 'e') val = v ?? '';
                else if (tipo === 'b') val = v === '1';
                else if (v === undefined) val = '';
                else val = esFecha[Number(c.getAttribute('s') || 0)] ? fechaHoraDeSerie(v) : Number(v);
                fila[j] = val;
            });
            filas[r] = fila;
        });
        hojas.push({ nombre: s.getAttribute('name'), filas: Array.from(filas, f => Array.from(f || [], x => (x === undefined ? '' : x))) });
    }
    return { hojas };
}
