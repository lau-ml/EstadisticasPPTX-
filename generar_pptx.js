'use strict';

const fs        = require('fs');
const path      = require('path');
const XLSX      = require('xlsx');
const PptxGenJS = require('pptxgenjs');

const BASE   = path.join(__dirname, '/');
const TURDIR = path.join(BASE, 'turnos_csv/');
const OUTPUT = path.join(BASE, 'Telesalud_HCANK_Estadisticas.pptx');

// ═══════════════════════════════════════════════════════════════════
// PALETA
// ═══════════════════════════════════════════════════════════════════
const C = {
  azul:    '2B6CB0',
  verde:   '276749',
  rojo:    'C53030',
  naranja: 'C05621',
  violeta: '553C9A',
  teal:    '2C7A7B',
  gris:    '718096',
  amarillo:'B7791F',
  oscuro:  '2D3748',
  navy:    '1A365D',
  celeste: '4299E1',
};
const PAL = [C.azul, C.verde, C.rojo, C.naranja, C.violeta, C.teal,
             C.gris, C.amarillo, C.oscuro];

// ═══════════════════════════════════════════════════════════════════
// PARSEO DE ARCHIVOS
// ═══════════════════════════════════════════════════════════════════

function parseCsvSemi(text) {
  const lines = text.replace(/\r/g, '').split('\n').filter(l => l.trim());
  const hdrs  = lines[0].split(';').map(h => h.trim().replace(/^"|"$/g, ''));
  return lines.slice(1).map(line => {
    const vals = line.split(';');
    return Object.fromEntries(hdrs.map((h, i) => [h, (vals[i] || '').trim().replace(/^"|"$/g, '')]));
  });
}

function parseCsvComma(text) {
  const parseRow = line => {
    const cells = [];
    let cur = '', inQ = false;
    for (const ch of line) {
      if (ch === '"') inQ = !inQ;
      else if (ch === ',' && !inQ) { cells.push(cur.trim()); cur = ''; }
      else cur += ch;
    }
    cells.push(cur.trim());
    return cells;
  };
  const lines = text.replace(/\r/g, '').split('\n').filter(l => l.trim());
  const hdrs  = parseRow(lines[0]);
  return lines.slice(1).map(line => {
    const vals = parseRow(line);
    return Object.fromEntries(hdrs.map((h, i) => [h, vals[i] || '']));
  });
}

function readXlsx(fp) {
  const wb = XLSX.readFile(fp);
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

// ═══════════════════════════════════════════════════════════════════
// HELPERS DE DOMINIO
// ═══════════════════════════════════════════════════════════════════

const ESTADOS_ATENDIDO = new Set(['Atendido', 'Asistencia', 'En Atención']);
const isAtendido = e => ESTADOS_ATENDIDO.has((e || '').trim());

const conOsTurnos = v => {
  const s = (v || '').toUpperCase();
  return !s.includes('00 - NO POSEE') && !s.includes('SIN COBERTURA') && s.trim() !== '';
};
const conOsAgendas = v => {
  const s = (v || '').toUpperCase();
  return !s.includes('00 - NO POSEE') && !s.includes('SIN COBERTURA') &&
         s.trim() !== '' && s !== 'NONE';
};

const RED_FNS = {
  'Cañuelas':        p => p.includes('cañuelas'),
  'San Vicente':     p => p.includes('san vicente'),
  'Gral. Las Heras': p => p.includes('las heras'),
  'Marcos Paz':      p => p.includes('marcos paz'),
  'Pte. Perón':      p => p.includes('perón') || p.includes('peron'),
};

function normMes(v) {
  if (!v) return null;
  const s = String(v).trim();
  let m;
  if ((m = s.match(/^(\d{4})[\/\-](\d{2})/))) return `${m[1]}-${m[2]}`;
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)))
    return `${m[3]}-${m[2].padStart(2, '0')}`;
  return null;
}

function counter(arr) {
  return arr.reduce((a, v) => { a[v] = (a[v] || 0) + 1; return a; }, {});
}
function topN(obj, n) {
  return Object.entries(obj)
    .filter(([k]) => k && k !== 'None' && k !== 'undefined' && k !== '')
    .sort((a, b) => b[1] - a[1])
    .slice(0, n);
}

// ── Helpers de tiempo ────────────────────────────────────────────────
function toMin(t) {
  if (!t || !t.trim()) return null;
  try {
    const parts = t.trim().split(':');
    return parseInt(parts[0]) * 60 + parseInt(parts[1]) +
           (parts[2] ? parseInt(parts[2]) / 60 : 0);
  } catch { return null; }
}

function diffMin(h1, h2) {
  const m1 = toMin(h1), m2 = toMin(h2);
  if (m1 === null || m2 === null) return null;
  const d = m2 - m1;
  return d >= 0 ? d : null;
}

// ── IQR + media ──────────────────────────────────────────────────────
function iqrFilter(vals) {
  if (vals.length < 4) return vals;
  const s = [...vals].sort((a, b) => a - b);
  const q1 = s[Math.floor(s.length / 4)];
  const q3 = s[Math.floor(3 * s.length / 4)];
  const iqr = q3 - q1;
  return vals.filter(v => v >= q1 - 1.5 * iqr && v <= q3 + 1.5 * iqr);
}
function mean(arr) { return arr.reduce((a, b) => a + b, 0) / arr.length; }

// ═══════════════════════════════════════════════════════════════════
// CARGA DE DATOS
// ═══════════════════════════════════════════════════════════════════

console.log('Cargando datos...');
const turnos      = parseCsvSemi(fs.readFileSync(path.join(TURDIR, 'Informe de Turnos.csv'), 'latin1'));
const agendas     = readXlsx(path.join(TURDIR, 'DETALLE DE AGENDAS.xlsx'));
const cancelados  = readXlsx(path.join(TURDIR, 'CANCELADOS.xlsx'));
const consolidado = parseCsvComma(fs.readFileSync(path.join(BASE, 'consolidado_pivoteado.csv'), 'utf8'));
console.log(`  Turnos: ${turnos.length}  |  Agendas: ${agendas.length}  |  Cancelados: ${cancelados.length}  |  Consolidado: ${consolidado.length}`);

// ═══════════════════════════════════════════════════════════════════
// INDICADORES 1–9  (sin cambios)
// ═══════════════════════════════════════════════════════════════════

function ind1() {
  const at  = turnos.filter(r => isAtendido(r.Estado));
  const con = at.filter(r => conOsTurnos(r.Cobertura)).length;
  return { labels: ['Con obra social', 'Sin obra social'], values: [con, at.length - con] };
}

function ind2() {
  const cnt = {};
  for (const r of agendas) {
    if (!isAtendido(r.Estado)) continue;
    const p = String(r.Partido || '').trim();
    if (!p || p === 'None') continue;
    const partido = p.includes('(') ? p.slice(p.lastIndexOf('(') + 1, p.lastIndexOf(')')) : p;
    cnt[partido] = (cnt[partido] || 0) + 1;
  }
  const top = topN(cnt, 12);
  return { labels: top.map(t => t[0]), values: top.map(t => t[1]) };
}

function ind3() {
  const munis  = Object.keys(RED_FNS);
  const conAt  = Object.fromEntries(munis.map(m => [m, 0]));
  const sinAt  = Object.fromEntries(munis.map(m => [m, 0]));
  const conNat = Object.fromEntries(munis.map(m => [m, 0]));
  for (const r of agendas) {
    const p = (r.Partido || '').toLowerCase();
    for (const [m, fn] of Object.entries(RED_FNS)) {
      if (!fn(p)) continue;
      const os = conOsAgendas(r.ObSoc), at = isAtendido(r.Estado);
      if (os && at)  conAt[m]++;
      if (!os && at) sinAt[m]++;
      if (os && !at) conNat[m]++;
      break;
    }
  }
  return { munis, conAt, sinAt, conNat };
}

function ind4() {
  const inas = agendas.filter(r => String(r.Estado).trim() === 'Inasistencia');
  const cnt  = counter(inas.map(r => String(r.Especialidad || '').trim()));
  delete cnt['']; delete cnt['None'];
  const top = topN(cnt, 12);
  return { labels: top.map(t => t[0]), values: top.map(t => t[1]), total: inas.length };
}

function ind5() {
  const cnt = counter(cancelados.map(r => String(r.Agenda || '').trim()));
  delete cnt['']; delete cnt['None'];
  const top = topN(cnt, 12);
  return { labels: top.map(t => t[0]), values: top.map(t => t[1]), total: cancelados.length };
}

// 6. Duración promedio por SERVICIO  (consolidado_pivoteado, IQR)
//    Columna SERVICIO disponible directamente; duración = Hora finalización - Hora creación
function ind6() {
  const durEsp = {};
  for (const row of consolidado) {
    const srv = (row['SERVICIO'] || '').trim();
    if (!srv) continue;
    const dur = diffMin(row['Hora creación'], row['Hora finalización']);
    if (dur === null || dur <= 0 || dur > 300) continue;
    if (!durEsp[srv]) durEsp[srv] = [];
    durEsp[srv].push(dur);
  }
  const result = Object.entries(durEsp)
    .map(([srv, vs]) => [srv, Math.round(mean(iqrFilter(vs)))])
    .sort((a, b) => b[1] - a[1]);
  return { labels: result.map(t => t[0]), values: result.map(t => t[1]) };
}

function ind7() {
  const eC = [], eS = [];
  for (const r of agendas) {
    if (!isAtendido(r.Estado)) continue;
    const edad = parseFloat(r.Edad);
    if (isNaN(edad) || edad <= 0 || edad > 120) continue;
    (conOsAgendas(r.ObSoc) ? eC : eS).push(edad);
  }
  const avg = arr => arr.length ? parseFloat((mean(arr)).toFixed(1)) : 0;
  return { labels: ['Con obra social', 'Sin obra social'],
           values: [avg(eC), avg(eS)], nCon: eC.length, nSin: eS.length };
}

function ind8() {
  const cnt = counter(
    agendas.filter(r => isAtendido(r.Estado)).map(r => String(r.Especialidad || '').trim())
  );
  delete cnt['']; delete cnt['None'];
  const top = topN(cnt, 12);
  return { labels: top.map(t => t[0]), values: top.map(t => t[1]) };
}

function ind9() {
  const mesEsp = {};
  for (const r of agendas) {
    if (!isAtendido(r.Estado)) continue;
    const mes = normMes(r.Fecha_Turno);
    const esp = String(r.Especialidad || '').trim();
    if (!mes || !esp || esp === 'None') continue;
    if (!mesEsp[esp]) mesEsp[esp] = {};
    mesEsp[esp][mes] = (mesEsp[esp][mes] || 0) + 1;
  }
  const top8 = Object.entries(mesEsp)
    .map(([e, d]) => [e, Object.values(d).reduce((a, b) => a + b, 0)])
    .sort((a, b) => b[1] - a[1]).slice(0, 8).map(t => t[0]);
  const meses = [...new Set(Object.values(mesEsp).flatMap(d => Object.keys(d)))].sort();
  const MESES_ES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const mesLabel = m => {
    const [y, mo] = m.split('-');
    return `${MESES_ES[parseInt(mo) - 1]} ${y.slice(2)}`;
  };
  return {
    labels: meses.map(mesLabel),
    series: top8.map(esp => ({ name: esp, values: meses.map(m => mesEsp[esp][m] || 0) })),
  };
}

// 10. Demora de atención por SERVICIO  (consolidado_pivoteado)
//     Columna PROFESIONAL (col 29) identifica al profesional;
//     SERVICIO (col 28) es la columna de servicio adyacente.
//     Espera = Hora primera aparición 2 (paciente entra) − Hora primera aparición 1 (profesional abre)
function ind10() {
  const delayEsp = {};
  for (const row of consolidado) {
    const srv = (row['SERVICIO'] || '').trim();
    if (!srv) continue;
    const h1 = row['Hora primera aparición 1'] || '';
    const h2 = row['Hora primera aparición 2'] || '';
    const d  = diffMin(h1, h2);
    if (d === null) continue;
    if (!delayEsp[srv]) delayEsp[srv] = [];
    delayEsp[srv].push(d);
  }
  const result = Object.entries(delayEsp)
    .filter(([, vs]) => vs.length >= 2)
    .map(([srv, vs]) => ({ srv, avg: parseFloat(mean(iqrFilter(vs)).toFixed(1)) }))
    .sort((a, b) => b.avg - a.avg);
  return {
    labels: result.map(r => r.srv),
    espera: result.map(r => r.avg),
  };
}

// ═══════════════════════════════════════════════════════════════════
// GENERACIÓN DE PPTX
// ═══════════════════════════════════════════════════════════════════

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.title  = 'Estadísticas Telesalud HCANK';
pptx.author = 'Telesalud HCANK';

const CX = 0.4, CY = 1.0, CW = 12.4, CH = 5.9;
const CW_N = 8, CX_N = 2.7;

function header(slide, title, subtitle) {
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:0.75, fill:{color:C.navy} });
  slide.addText(title, {
    x:0.35, y:0, w:12.6, h:0.75,
    color:'FFFFFF', bold:true, fontSize:18, fontFace:'Calibri', valign:'middle',
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x:0, y:0.76, w:'100%', h:0.28,
      color:'718096', fontSize:10, fontFace:'Calibri',
      align:'right', margin:[0, 0.25, 0, 0],
    });
  }
}

function barOpts(extra = {}) {
  return {
    x: CX, y: CY, w: CW, h: CH,
    showTitle: false, showLegend: false,
    dataLabelFontSize: 9,
    catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    ...extra,
  };
}

console.log('Generando PPTX...');

// ── Portada ──────────────────────────────────────────────────────────
const s0 = pptx.addSlide();
s0.background = { color: C.navy };
s0.addText('Telesalud HCANK', {
  x:1, y:1.5, w:11.33, h:1.2,
  color:'FFFFFF', bold:true, fontSize:40, fontFace:'Calibri', align:'center',
});
s0.addText('Estadísticas de Atención · 2026', {
  x:1, y:2.9, w:11.33, h:0.65,
  color:'90CDF4', fontSize:22, fontFace:'Calibri', align:'center',
});
s0.addText('Consolidado de indicadores clínicos y operativos', {
  x:1, y:3.75, w:11.33, h:0.4,
  color:'BEE3F8', fontSize:14, fontFace:'Calibri', align:'center',
});
s0.addShape(pptx.ShapeType.rect, { x:4.5, y:6.8, w:4.33, h:0.05, fill:{color:'4A90D9'} });

// ── 1. Con / Sin obra social ──────────────────────────────────────────
const d1 = ind1();
const s1  = pptx.addSlide();
header(s1, 'Pacientes atendidos: Con / Sin obra social',
  `Con OS: ${d1.values[0]}  ·  Sin OS: ${d1.values[1]}`);
s1.addChart(pptx.ChartType.pie,
  [{ name:'Cobertura', labels:d1.labels, values:d1.values }],
  { x:3, y:1.0, w:7.33, h:6.0,
    chartColors:[C.azul,C.rojo], showLegend:true, legendPos:'b', legendFontSize:14,
    showTitle:false, dataLabelFontSize:16, showPercent:true,
    dataLabelColor:'FFFFFF', dataLabelBold:true });

// ── 2. Atendidos por localidad ────────────────────────────────────────
const d2 = ind2();
const s2  = pptx.addSlide();
header(s2, 'Pacientes atendidos por localidad', 'Top 12 partidos · Detalle de Agendas');
s2.addChart(pptx.ChartType.bar,
  [{ name:'Atendidos', labels:d2.labels, values:d2.values }],
  barOpts({ barDir:'bar', chartColors:[C.azul], dataLabelPosition:'outEnd',
            valAxisMaxVal:Math.max(...d2.values)*1.18 }));

// ── 3. Red municipios ────────────────────────────────────────────────
const d3 = ind3();
const s3  = pptx.addSlide();
header(s3, 'Municipios de la red: cobertura y estado de atención',
  'Cañuelas · San Vicente · Gral. Las Heras · Marcos Paz · Pte. Perón');
s3.addChart(pptx.ChartType.bar, [
  { name:'Con OS · Atendidos',    labels:d3.munis, values:d3.munis.map(m=>d3.conAt[m])  },
  { name:'Sin OS · Atendidos',    labels:d3.munis, values:d3.munis.map(m=>d3.sinAt[m])  },
  { name:'Con OS · No atendidos', labels:d3.munis, values:d3.munis.map(m=>d3.conNat[m]) },
], barOpts({ barDir:'col', barGrouping:'clustered',
             chartColors:[C.verde,C.rojo,C.naranja],
             showLegend:true, legendPos:'b', legendFontSize:10,
             catAxisLabelFontSize:11, dataLabelPosition:'outEnd' }));

// ── 4. Inasistencias ─────────────────────────────────────────────────
const d4 = ind4();
const s4  = pptx.addSlide();
header(s4, 'Inasistencias por especialidad',
  `Pacientes que no asistieron sin cancelar el turno · Total: ${d4.total}`);
s4.addChart(pptx.ChartType.bar,
  [{ name:'Inasistencias', labels:d4.labels, values:d4.values }],
  barOpts({ barDir:'col', chartColors:[C.rojo], dataLabelPosition:'outEnd',
            catAxisLabelFontSize:8, valAxisMaxVal:Math.max(...d4.values)*1.2 }));

// ── 5. Turnos cancelados ──────────────────────────────────────────────
const d5 = ind5();
const s5  = pptx.addSlide();
header(s5, 'Turnos cancelados por agenda', `Total: ${d5.total} cancelaciones`);
s5.addChart(pptx.ChartType.bar,
  [{ name:'Cancelados', labels:d5.labels, values:d5.values }],
  barOpts({ barDir:'col', chartColors:[C.naranja], dataLabelPosition:'outEnd',
            catAxisLabelFontSize:8, valAxisMaxVal:Math.max(...d5.values)*1.2 }));

// ── 6. Duración por SERVICIO (consolidado_pivoteado, IQR corregido) ──
const d6 = ind6();
const s6  = pptx.addSlide();
header(s6, 'Duración promedio de llamada Telesalud por servicio',
  'Minutos · columna SERVICIO de consolidado_pivoteado · outliers removidos con IQR');
s6.addChart(pptx.ChartType.bar,
  [{ name:'Minutos promedio', labels:d6.labels, values:d6.values }],
  barOpts({ barDir:'col', chartColors:[C.violeta], dataLabelPosition:'outEnd',
            catAxisLabelFontSize:10, valAxisMaxVal:Math.max(...d6.values)*1.2 }));

// ── 7. Promedio de edad ───────────────────────────────────────────────
const d7 = ind7();
const s7  = pptx.addSlide();
header(s7, 'Promedio de edad de pacientes atendidos',
  `Con OS: n=${d7.nCon}  ·  Sin OS: n=${d7.nSin}`);
s7.addChart(pptx.ChartType.bar,
  [{ name:'Edad promedio', labels:d7.labels, values:d7.values }],
  { x:CX_N, y:CY, w:CW_N, h:CH, barDir:'col', chartColors:[C.azul,C.rojo],
    showTitle:false, showLegend:false, dataLabelFontSize:15, dataLabelBold:true,
    catAxisLabelFontSize:13, valAxisLabelFontSize:10, dataLabelPosition:'outEnd',
    valAxisMinVal:0, valAxisMaxVal:Math.max(...d7.values)*1.25 });

// ── 8. Atenciones por especialidad ────────────────────────────────────
const d8 = ind8();
const s8  = pptx.addSlide();
header(s8, 'Cantidad de atenciones por especialidad');
s8.addChart(pptx.ChartType.bar,
  [{ name:'Atenciones', labels:d8.labels, values:d8.values }],
  barOpts({ barDir:'bar', chartColors:[C.teal], dataLabelPosition:'outEnd',
            valAxisMaxVal:Math.max(...d8.values)*1.18 }));

// ── 9. Evolución mensual por especialidad ─────────────────────────────
const d9 = ind9();
const s9  = pptx.addSlide();
header(s9, 'Evolución mensual de atenciones por especialidad',
  'Top 8 especialidades · Detalle de Agendas');
s9.addChart(pptx.ChartType.line,
  d9.series.map(serie => ({ name:serie.name, labels:d9.labels, values:serie.values })),
  { x:CX, y:CY, w:CW, h:CH, chartColors:PAL.slice(0, d9.series.length),
    showTitle:false, showLegend:true, legendPos:'r', legendFontSize:9,
    lineSize:2, lineDataSymbol:'circle', lineDataSymbolSize:5,
    dataLabelFontSize:0, catAxisLabelFontSize:9, valAxisLabelFontSize:9,
    catAxisLabelRotate:30 });

// ── 10. Demora de atención por SERVICIO ───────────────────────────────
//     Fuente: consolidado_pivoteado  (columna PROFESIONAL → columna SERVICIO adyacente)
//     Espera paciente = Hora primera aparición 2 − Hora primera aparición 1 (IQR)
const d10 = ind10();
const s10  = pptx.addSlide();
header(s10, 'Demora promedio de atención por servicio',
  'Minutos de espera hasta que el paciente entra · consolidado_pivoteado · IQR');
s10.addChart(pptx.ChartType.bar,
  [{ name:'Espera del paciente (min)', labels:d10.labels, values:d10.espera }],
  barOpts({
    barDir:            'col',
    chartColors:       [C.azul],
    showLegend:        false,
    dataLabelPosition: 'outEnd',
    dataLabelFontSize: 11,
    dataLabelBold:     true,
    catAxisLabelFontSize: 11,
    valAxisLabelFontSize: 9,
    valAxisMaxVal:     Math.max(...d10.espera) * 1.25,
  }));

// ═══════════════════════════════════════════════════════════════════
// GUARDAR
// ═══════════════════════════════════════════════════════════════════

pptx.writeFile({ fileName: OUTPUT })
  .then(() => console.log(`\n✓ Archivo guardado: ${OUTPUT}`))
  .catch(err => { console.error('Error al guardar:', err); process.exit(1); });
