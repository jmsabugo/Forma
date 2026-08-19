'use strict';

// ===== Estado y persistencia local =====
const LS = {
  data: 'forma_data',         // datos parseados del Excel
  rev: 'forma_rev',           // rev de Dropbox de la última sincronización
  pending: 'forma_pending',   // hay cambios locales sin subir
  settings: 'forma_settings', // appKey, refreshToken, ruta del xlsx
  backup: 'forma_backup',     // copia local si hubo conflicto
  rutinaHoy: 'forma_rutina_hoy', // {fecha, nombre, extras} de la sesión en curso
};

const RUTA_DEFECTO = '/JnM Particular/1.1. Proyectos_Personales/App GYM - Forma - Encanta/Forma_Datos.xlsx';

let state = {
  data: null,        // {ejercicios:[], registro:[], config:{}}
  settings: { appKey: '', refreshToken: '', path: RUTA_DEFECTO, lastSync: '' },
  accessToken: null,
  tokenExpira: 0,
  tab: 'hoy',
  entrada: {},      // valores en curso de la entrada rápida: { [id]: {peso, reps} }
  rutinaHoy: null,  // rutina elegida para hoy (nombre, o '__libre__', o null)
  extras: [],       // ejercicios sueltos añadidos hoy fuera de la rutina
  hoyExpandido: null, // (sin uso desde la ficha enfocada; se conserva por si acaso)
  fichaAbierta: null, // id del ejercicio abierto en la ficha enfocada (pantalla completa)
  fichaLado: {},      // lado activo en la ficha de cada unilateral: {id: 'Izq'|'Der'}
  flashPR: null,    // clave de serie a animar al lograr un récord (+3 reps a ese peso)
  undoStack: [],    // instantáneas de datos para "deshacer" (solo en memoria)
  ejBuscar: '',     // texto del buscador del catálogo (pestaña Ejercicios)
  editEj: null,     // id del ejercicio en edición en el catálogo ('__nuevo__' al crear)
  histAbierto: null,    // fecha (ISO) del día desplegado en el Historial
  histEdit: null,       // índice en registro de la serie en edición
  histRutinaNueva: {},  // rutina elegida para registrar en un día vacío: {fecha: nombre}
  prog: { periodo: 'mes', kpi: 'entrenos', grupo: '', ejercicio: null, metrica: '1rm', _charts: [] }, // pestaña Progreso
};

// Valores admitidos en el catálogo (coinciden con los desplegables del Excel).
const EQUIPOS = ['Barra', 'Mancuernas', 'Polea', 'Lastre'];
const LATERALIDADES = ['Bilateral', 'Unilateral'];

// Stack de placas de polea por defecto (kg, conversión de 10..210 lb).
const DEFAULT_POLEA = [4.5, 9.1, 13.6, 18.1, 22.7, 27.2, 31.8, 36.3, 40.8, 45.4,
  50, 54.4, 59, 63.5, 68, 72.6, 77.1, 81.6, 86.2, 90.7, 95.3];

function loadLocal() {
  try {
    const d = localStorage.getItem(LS.data);
    if (d) { state.data = JSON.parse(d); lastGood = JSON.stringify(state.data); }
    const s = localStorage.getItem(LS.settings);
    if (s) state.settings = Object.assign(state.settings, JSON.parse(s));
    // Selección de rutina: solo vale si es de hoy.
    const r = localStorage.getItem(LS.rutinaHoy);
    if (r) {
      const o = JSON.parse(r);
      if (o.fecha === hoyISO()) { state.rutinaHoy = o.nombre; state.extras = o.extras || []; }
    }
  } catch (e) { console.error(e); }
}

function guardarRutinaHoy() {
  localStorage.setItem(LS.rutinaHoy, JSON.stringify({
    fecha: hoyISO(), nombre: state.rutinaHoy, extras: state.extras,
  }));
}

// Instantánea del último estado persistido y bandera para no apilar (config/sync).
let lastGood = null;
let suspenderUndo = false;

function deshacerOn() { return cfgBool('permitir_deshacer', true); }

function saveData() {
  const cur = JSON.stringify(state.data);
  if (deshacerOn() && !suspenderUndo && lastGood != null && lastGood !== cur) {
    state.undoStack.push(lastGood);
    if (state.undoStack.length > 20) state.undoStack.shift();
  }
  localStorage.setItem(LS.data, cur);
  lastGood = cur;
}

function deshacer() {
  if (!state.undoStack.length) return;
  const prev = state.undoStack.pop();
  state.data = JSON.parse(prev);
  lastGood = prev;                       // no volver a apilar este mismo estado
  localStorage.setItem(LS.data, prev);   // persistir sin pasar por la pila
  marcarPendiente(true);
  render();
}

function pintarUndo() {
  const b = document.getElementById('btn-undo');
  if (!b) return;
  b.hidden = !(deshacerOn() && state.undoStack.length);
}

// Conecta los buscadores: el del catálogo (oculta tarjetas) y los de los
// selectores de ejercicio (ocultan opciones del desplegable).
function wireBuscadores() {
  const cat = document.getElementById('ej-buscar');
  if (cat) {
    const aplicar = () => {
      state.ejBuscar = cat.value;
      const q = cat.value.trim().toLowerCase();
      document.querySelectorAll('[data-buscar]').forEach(el => {
        el.style.display = (!q || el.dataset.buscar.includes(q)) ? '' : 'none';
      });
    };
    cat.oninput = aplicar;
    aplicar(); // reaplica el filtro tras un re-render
  }
  document.querySelectorAll('.buscar-sel').forEach(inp => {
    const sel = document.getElementById(inp.dataset.target);
    if (!sel) return;
    inp.oninput = () => {
      const q = inp.value.trim().toLowerCase();
      [...sel.options].forEach(o => { if (o.value) o.hidden = !!q && !o.textContent.toLowerCase().includes(q); });
    };
  });
}

function saveSettings() {
  localStorage.setItem(LS.settings, JSON.stringify(state.settings));
}

function marcarPendiente(v) {
  if (v) localStorage.setItem(LS.pending, '1');
  else localStorage.removeItem(LS.pending);
  pintarBadge();
}

// ===== Excel <-> modelo de datos (SheetJS) =====
// El orden de estas columnas es el que se ESCRIBE; la lectura es por nombre de
// cabecera (parseHoja), así que reordenar columnas en el Excel no rompe nada.
const CABECERAS = {
  Ejercicios: ['ID', 'Ejercicio', 'Grupo muscular', 'Equipamiento', 'Lateralidad',
               'Descanso (min)', 'Series objetivo', 'Reps objetivo', 'Activo', 'Notas', 'Polea 1/2'],
  Registro: ['Fecha', 'ID', 'Ejercicio', 'Serie', 'Lado', 'Repeticiones', 'Peso (kg)', 'Rutina', 'Notas'],
  Config: ['Clave', 'Valor', 'Descripcion'],
  Rutinas: ['Rutina', 'Orden', 'ID', 'Ejercicio'],
};

// Lee una hoja como {idx, filas}: idx mapea nombre de cabecera → índice de columna,
// filas son los arrays de datos (sin la cabecera, sin filas vacías).
function parseHoja(wb, nombre, requerida) {
  const ws = wb.Sheets[nombre];
  if (!ws) {
    if (requerida) throw new Error(`Falta la hoja "${nombre}" en el Excel`);
    return { idx: {}, filas: [] };
  }
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const idx = {};
  (aoa[0] || []).forEach((h, i) => {
    const k = String(h).trim();
    if (k && !(k in idx)) idx[k] = i;
  });
  const filas = aoa.slice(1).filter(f => f.some(c => c !== '' && c != null));
  return { idx, filas };
}

// Valor de una celda por nombre de columna (''.si no existe la columna).
function celda(f, idx, nombre) {
  const i = idx[nombre];
  if (i == null) return '';
  return f[i] == null ? '' : f[i];
}

function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });

  const ej = parseHoja(wb, 'Ejercicios', true);
  const ejercicios = ej.filas.map(f => ({
    id: String(celda(f, ej.idx, 'ID')),
    nombre: celda(f, ej.idx, 'Ejercicio'),
    grupo: celda(f, ej.idx, 'Grupo muscular'),
    equipamiento: celda(f, ej.idx, 'Equipamiento'),
    lateralidad: celda(f, ej.idx, 'Lateralidad'),
    descanso: Number(celda(f, ej.idx, 'Descanso (min)')) || 0,
    seriesObj: Number(celda(f, ej.idx, 'Series objetivo')) || 0,
    repsObj: Number(celda(f, ej.idx, 'Reps objetivo')) || 0,
    activo: String(celda(f, ej.idx, 'Activo')).toUpperCase() !== 'NO',
    notas: celda(f, ej.idx, 'Notas') || '',
    polea12: String(celda(f, ej.idx, 'Polea 1/2')).trim().toUpperCase() === 'SI',
  }));

  const rg = parseHoja(wb, 'Registro', true);
  const registro = rg.filas.map(f => ({
    fecha: fechaISO(celda(f, rg.idx, 'Fecha')),
    id: String(celda(f, rg.idx, 'ID')),
    ejercicio: celda(f, rg.idx, 'Ejercicio'),
    serie: Number(celda(f, rg.idx, 'Serie')) || 0,
    lado: String(celda(f, rg.idx, 'Lado') || '').trim(),
    reps: Number(celda(f, rg.idx, 'Repeticiones')) || 0,
    peso: Number(celda(f, rg.idx, 'Peso (kg)')) || 0,
    rutina: String(celda(f, rg.idx, 'Rutina') || '').trim(),
    notas: celda(f, rg.idx, 'Notas') || '',
  }));

  const cf = parseHoja(wb, 'Config', true);
  const config = {};
  cf.filas.forEach(f => {
    const clave = celda(f, cf.idx, 'Clave');
    if (clave === '') return;
    config[clave] = { valor: celda(f, cf.idx, 'Valor'), descripcion: celda(f, cf.idx, 'Descripcion') || '' };
  });

  // La hoja Rutinas es opcional (Excels antiguos no la tienen).
  const rutinas = agruparRutinas(parseHoja(wb, 'Rutinas', false));
  return { ejercicios, registro, config, rutinas };
}

// Agrupa las filas planas de la hoja Rutinas en [{nombre, ids:[...]}],
// respetando la columna Orden.
function agruparRutinas(rt) {
  const m = new Map();
  rt.filas.forEach(f => {
    const nombre = String(celda(f, rt.idx, 'Rutina')).trim();
    const id = String(celda(f, rt.idx, 'ID')).trim();
    if (!nombre || !id) return;
    if (!m.has(nombre)) m.set(nombre, []);
    m.get(nombre).push({ orden: Number(celda(f, rt.idx, 'Orden')) || 0, id });
  });
  return [...m.entries()].map(([nombre, arr]) => ({
    nombre, ids: arr.sort((a, b) => a.orden - b.orden).map(x => x.id),
  }));
}

function buildWorkbook() {
  const d = state.data;
  const wb = XLSX.utils.book_new();
  const aoaEj = [CABECERAS.Ejercicios].concat(d.ejercicios.map(e => [
    e.id, e.nombre, e.grupo, e.equipamiento, e.lateralidad,
    e.descanso, e.seriesObj, e.repsObj, e.activo ? 'SI' : 'NO', e.notas,
    e.polea12 ? 'SI' : 'NO',
  ]));
  const aoaReg = [CABECERAS.Registro].concat(d.registro.map(r => [
    r.fecha, r.id, r.ejercicio, r.serie, r.lado || '', r.reps, r.peso, r.rutina || '', r.notas,
  ]));
  const aoaCfg = [CABECERAS.Config].concat(
    Object.entries(d.config).map(([k, v]) => [k, v.valor, v.descripcion]));
  const aoaRut = [CABECERAS.Rutinas];
  (d.rutinas || []).forEach(r => r.ids.forEach((id, i) => {
    const ej = d.ejercicios.find(e => e.id === id);
    aoaRut.push([r.nombre, i + 1, id, ej ? ej.nombre : '']);
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoaEj), 'Ejercicios');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoaReg), 'Registro');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoaCfg), 'Config');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoaRut), 'Rutinas');
  return wb;
}

function fechaISO(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'number') { // serial de Excel
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // DD/MM/YYYY
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return s.slice(0, 10);
}

// ===== Dropbox (OAuth PKCE, sin servidor) =====
function b64url(buf) {
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

async function conectarDropbox() {
  const appKey = document.getElementById('app-key').value.trim();
  if (!appKey) { alert('Pega primero la App Key de tu app de Dropbox.'); return; }
  state.settings.appKey = appKey;
  saveSettings();
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(48)));
  sessionStorage.setItem('pkce_verifier', verifier);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const url = 'https://www.dropbox.com/oauth2/authorize'
    + `?client_id=${encodeURIComponent(appKey)}`
    + '&response_type=code'
    + `&code_challenge=${challenge}`
    + '&code_challenge_method=S256'
    + '&token_access_type=offline'
    + `&redirect_uri=${encodeURIComponent(redirectUri())}`;
  location.href = url;
}

function redirectUri() {
  return location.origin + location.pathname;
}

async function canjearCodigo(code) {
  const verifier = sessionStorage.getItem('pkce_verifier');
  if (!verifier) return;
  const body = new URLSearchParams({
    code, grant_type: 'authorization_code',
    client_id: state.settings.appKey,
    code_verifier: verifier,
    redirect_uri: redirectUri(),
  });
  const r = await fetch('https://api.dropboxapi.com/oauth2/token', { method: 'POST', body });
  if (!r.ok) { alert('Error al conectar con Dropbox: ' + await r.text()); return; }
  const j = await r.json();
  state.settings.refreshToken = j.refresh_token;
  state.accessToken = j.access_token;
  state.tokenExpira = Date.now() + (j.expires_in - 60) * 1000;
  saveSettings();
  sessionStorage.removeItem('pkce_verifier');
  history.replaceState(null, '', redirectUri());
  await sincronizar();
}

async function token() {
  if (state.accessToken && Date.now() < state.tokenExpira) return state.accessToken;
  if (!state.settings.refreshToken) throw new Error('Dropbox no conectado');
  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: state.settings.refreshToken,
    client_id: state.settings.appKey,
  });
  const r = await fetch('https://api.dropboxapi.com/oauth2/token', { method: 'POST', body });
  if (!r.ok) throw new Error('No se pudo renovar el token de Dropbox');
  const j = await r.json();
  state.accessToken = j.access_token;
  state.tokenExpira = Date.now() + (j.expires_in - 60) * 1000;
  return state.accessToken;
}

async function descargarExcel() {
  const t = await token();
  const r = await fetch('https://content.dropboxapi.com/2/files/download', {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + t,
      'Dropbox-API-Arg': JSON.stringify({ path: state.settings.path }),
    },
  });
  if (!r.ok) throw new Error('Error al descargar el Excel: ' + await r.text());
  const meta = JSON.parse(r.headers.get('dropbox-api-result'));
  return { buf: await r.arrayBuffer(), rev: meta.rev };
}

async function subirExcel() {
  const t = await token();
  const wb = buildWorkbook();
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const revLocal = localStorage.getItem(LS.rev);
  const modo = revLocal ? { '.tag': 'update', update: revLocal } : 'overwrite';
  const r = await fetch('https://content.dropboxapi.com/2/files/upload', {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + t,
      'Dropbox-API-Arg': JSON.stringify({ path: state.settings.path, mode: modo, mute: true }),
      'Content-Type': 'application/octet-stream',
    },
    body: buf,
  });
  if (r.status === 409) { // conflicto: alguien cambió el Excel por fuera
    localStorage.setItem(LS.backup, JSON.stringify(state.data));
    const remoto = await descargarExcel();
    state.data = parseWorkbook(remoto.buf);
    saveData();
    localStorage.setItem(LS.rev, remoto.rev);
    marcarPendiente(false);
    alert('Conflicto: el Excel cambió fuera de la app. Se ha cargado la versión de Dropbox y tus cambios locales quedan en una copia de seguridad interna.');
    return;
  }
  if (!r.ok) throw new Error('Error al subir el Excel: ' + await r.text());
  const meta = await r.json();
  localStorage.setItem(LS.rev, meta.rev);
  marcarPendiente(false);
}

// Copia de seguridad del Excel: lo copia DENTRO de Dropbox (files/copy_v2) a la
// carpeta "Back up _ Forma_Datos" del proyecto, con nombre fechado. Si hay
// cambios sin subir, se sincronizan primero para que la copia incluya lo último.
const CARPETA_BACKUP = 'Back up _ Forma_Datos';

async function copiaSeguridad() {
  if (!state.settings.refreshToken) { alert('Conecta primero con Dropbox en Ajustes.'); return; }
  const btn = document.getElementById('btn-backup');
  if (btn) { btn.disabled = true; btn.textContent = 'Creando copia…'; }
  try {
    if (localStorage.getItem(LS.pending)) await subirExcel(); // la copia debe incluir lo último
    const t = await token();
    const origen = state.settings.path;
    const carpeta = origen.slice(0, origen.lastIndexOf('/'));
    const d = new Date();
    const p2 = (n) => String(n).padStart(2, '0');
    const sello = `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}${p2(d.getMinutes())}`;
    const destino = `${carpeta}/${CARPETA_BACKUP}/Forma_Datos backup ${sello}.xlsx`;
    const r = await fetch('https://api.dropboxapi.com/2/files/copy_v2', {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + t, 'Content-Type': 'application/json' },
      body: JSON.stringify({ from_path: origen, to_path: destino, autorename: true }),
    });
    if (!r.ok) throw new Error('Error al crear la copia: ' + await r.text());
    const j = await r.json();
    alert('Copia de seguridad creada:\n' + (j.metadata && j.metadata.name ? j.metadata.name : destino));
  } catch (e) {
    alert(e.message);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = '🗄️ Copia de seguridad'; }
    pintarBadge();
  }
}

async function sincronizar(silencioso) {
  if (!state.settings.refreshToken) { if (!silencioso) alert('Conecta primero con Dropbox en Ajustes.'); return false; }
  suspenderUndo = true;
  const btn = document.getElementById('btn-sync');
  if (btn) { btn.disabled = true; btn.textContent = 'Sincronizando…'; }
  let ok = true;
  try {
    const pendiente = !!localStorage.getItem(LS.pending);
    if (pendiente) {
      await subirExcel();
    } else {
      const { buf, rev } = await descargarExcel();
      state.data = parseWorkbook(buf);
      saveData();
      localStorage.setItem(LS.rev, rev);
    }
    state.settings.lastSync = new Date().toISOString();
    saveSettings();
    render();
  } catch (e) {
    ok = false;
    if (silencioso) console.warn('Sincronización automática falló:', e.message);
    else alert(e.message);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = 'Sincronizar ahora'; }
    suspenderUndo = false;
    state.undoStack = [];                                   // el histórico de sesión ya no aplica
    lastGood = state.data ? JSON.stringify(state.data) : null;
    pintarBadge();
  }
  return ok;
}

// ===== Importar / exportar manual (modo sin Dropbox) =====
function importarArchivo(ev) {
  const f = ev.target.files[0];
  if (!f) return;
  const lector = new FileReader();
  lector.onload = () => {
    try {
      suspenderUndo = true;
      state.data = parseWorkbook(lector.result);
      saveData();
      suspenderUndo = false;
      state.undoStack = [];
      marcarPendiente(false);
      render();
      alert('Excel importado correctamente.');
    } catch (e) { suspenderUndo = false; alert(e.message); }
  };
  lector.readAsArrayBuffer(f);
}

function exportarArchivo() {
  if (!state.data) { alert('No hay datos que exportar.'); return; }
  XLSX.writeFile(buildWorkbook(), 'Forma_Datos.xlsx');
}

// ===== Utilidades de datos =====
// Último registro de un ejercicio (fecha máx, serie máx). Si se pasa `lado`
// (Izq/Der), solo mira ese lado; sin lado o con '' mira todos los registros.
function ultimoPeso(idEjercicio, lado) {
  let mejor = null;
  for (const r of state.data.registro) {
    if (r.id !== idEjercicio) continue;
    if (lado && r.lado !== lado) continue;
    if (!mejor || r.fecha > mejor.fecha || (r.fecha === mejor.fecha && r.serie > mejor.serie)) mejor = r;
  }
  return mejor; // null si nunca se ha hecho
}

// Última serie registrada de un ejercicio (la del último entrenamiento).
// Sirve para precargar peso y repeticiones base de la entrada rápida.
function ultimaSerie(idEjercicio, lado) {
  return ultimoPeso(idEjercicio, lado); // misma lógica: fecha máx, serie máx
}

// Series del entrenamiento ANTERIOR de un ejercicio (excluye hoy), como
// referencia a superar en la entrada de Hoy. Con `lado` solo mira ese lado.
function ultimaSesionRef(id, lado) {
  const hoy = hoyISO();
  let fmax = '';
  for (const r of state.data.registro) {
    if (r.id !== id) continue;
    if (lado && r.lado !== lado) continue;
    if (r.fecha === hoy) continue;            // la referencia es el día anterior, no hoy
    if (r.fecha > fmax) fmax = r.fecha;
  }
  if (!fmax) return null;
  const series = state.data.registro
    .filter(r => r.id === id && r.fecha === fmax && (!lado || r.lado === lado))
    .sort((a, b) => a.serie - b.serie);
  return { fecha: fmax, series };
}

// Píldoras de la última vez: peso × reps de cada serie (referencia a superar).
function refUltimaHtml(id, lado) {
  const ref = ultimaSesionRef(id, lado);
  if (!ref || !ref.series.length) return '';
  const pills = ref.series.map(s =>
    `<span class="rep-pill">S${s.serie}: ${fmtPeso(s.peso)}×${s.reps}</span>`).join('');
  return `<div class="ref-ultima"><span class="ref-lbl">Última vez (${fmtFecha(ref.fecha)}):</span>${pills}</div>`;
}

// ----- Mejor set y récord por peso (señal de subir peso) -----
// Un "set" = todas las series de un ejercicio a un mismo peso un mismo día.
// La marca de un set es la SUMA de repeticiones de sus series. El récord salta
// cuando el total del set de hoy supera en 3 al mejor total previo a ese peso.
const PR_MARGEN = 3;

// Mejor set histórico a un peso concreto (el día con MÁS total de reps a ese peso).
// Con excluirHoy=true sirve de referencia a batir.
function mejorSetPeso(id, lado, peso, excluirHoy) {
  const pw = red2(peso);
  const hoy = hoyISO();
  const dias = {};
  for (const r of state.data.registro) {
    if (r.id !== id || (lado && r.lado !== lado) || red2(r.peso) !== pw) continue;
    if (excluirHoy && r.fecha === hoy) continue;
    (dias[r.fecha] = dias[r.fecha] || { series: [], total: 0 });
    dias[r.fecha].series.push(r);
    dias[r.fecha].total += r.reps;
  }
  let best = null;
  for (const f in dias) {
    const d = dias[f];
    if (!best || d.total > best.total || (d.total === best.total && f > best.fecha))
      best = { fecha: f, total: d.total, series: d.series };
  }
  if (best) best.series.sort((a, b) => a.serie - b.serie);
  return best;
}

// Total del set al que pertenece r y mejor total de días ANTERIORES a ese peso.
function statsSet(r) {
  const pw = red2(r.peso);
  let total = 0, mejorPrev = -1;
  const prev = {};
  for (const o of state.data.registro) {
    if (o.id !== r.id || (o.lado || '') !== (r.lado || '') || red2(o.peso) !== pw) continue;
    if (o.fecha === r.fecha) total += o.reps;
    else if (o.fecha < r.fecha) prev[o.fecha] = (prev[o.fecha] || 0) + o.reps;
  }
  for (const f in prev) if (prev[f] > mejorPrev) mejorPrev = prev[f];
  return { total, mejorPrev };
}

// ¿La serie r pertenece a un set récord? (total del set ≥ mejor previo + 3).
function esPRset(r) {
  if (!r) return false;
  const { total, mejorPrev } = statsSet(r);
  return mejorPrev >= 0 && total >= mejorPrev + PR_MARGEN;
}

// Clave de un set (ejercicio + lado + fecha + peso), para animar al lograrlo.
function setKey(r) { return `${r.fecha}|${r.id}|${r.lado || ''}|${red2(r.peso)}`; }

// Clases CSS de una serie: coral fuerte si es récord, con animación si se acaba de lograr.
function clasePR(r) {
  let c = '';
  if (esPRset(r)) c += ' pr';
  if (state.flashPR && setKey(r) === state.flashPR) c += ' pr-flash';
  return c;
}

// Referencia "Mejor a X kg": el mejor set previo a ese peso + objetivo para subir.
function refMejorHtml(id, lado, peso) {
  if (!peso) return '';
  const mj = mejorSetPeso(id, lado, peso, true);
  if (!mj) return '';
  const pills = mj.series.map(s => `<span class="rep-pill">${s.reps}</span>`).join('');
  return `<div class="ref-mejor"><span class="ref-lbl">Mejor a ${fmtPeso(peso)} kg (${fmtFecha(mj.fecha)}):</span>${pills}`
    + `<span class="ref-obj">objetivo ${mj.total + PR_MARGEN} (hoy)</span></div>`;
}

// Clave del estado de entrada en curso (por lado en unilaterales).
function entradaKey(id, lado) { return lado ? `${id}|${lado}` : id; }

// Nombre de rutina que se guarda en el Registro al apuntar una serie.
function rutinaParaGuardar() {
  if (!state.rutinaHoy || state.rutinaHoy === '__libre__') return 'Entreno Libre';
  return state.rutinaHoy;
}

function cfgNum(clave, defecto) {
  const c = state.data && state.data.config && state.data.config[clave];
  if (!c) return defecto;
  const n = Number(String(c.valor).replace(',', '.'));
  return Number.isFinite(n) && n > 0 ? n : defecto;
}

function incrementoPeso() { return cfgNum('incremento_peso', 1.25); }
function incrementoReps() { return cfgNum('incremento_reps', 1); }

// Lee una clave booleana de Config (cualquier valor distinto de "NO" = true).
function cfgBool(clave, defecto) {
  const c = state.data && state.data.config && state.data.config[clave];
  if (!c) return defecto;
  return String(c.valor).trim().toUpperCase() !== 'NO';
}

// Escribe/actualiza una clave de Config (se guarda en el Excel al sincronizar).
function setConfig(clave, valor, descripcion) {
  if (!state.data.config) state.data.config = {};
  const prev = state.data.config[clave];
  state.data.config[clave] = { valor, descripcion: (prev && prev.descripcion) || descripcion || '' };
  suspenderUndo = true; saveData(); suspenderUndo = false;  // los ajustes no entran en "deshacer"
  marcarPendiente(true);
}

// Confirmación de borrados, configurable en Ajustes (por defecto activada).
function confirmarBorradosOn() { return cfgBool('confirmar_borrados', true); }
function confirmar(msg) { return !confirmarBorradosOn() || confirm(msg); }

// Texto configurado del stack de polea (kg, decimales con punto, separados por comas).
function pesosPoleaTexto() {
  const c = state.data && state.data.config && state.data.config['pesos_polea'];
  if (c && String(c.valor).trim()) return String(c.valor).trim();
  return DEFAULT_POLEA.join(', ');
}
// Lista numérica y ordenada del stack de polea.
function stackPolea() {
  return pesosPoleaTexto().split(/[,;]/)
    .map(s => Number(s.trim()))
    .filter(n => Number.isFinite(n) && n > 0)
    .sort((a, b) => a - b);
}
// Siguiente (dir>0) o anterior (dir<0) placa del stack respecto a un peso actual.
function pasoPolea(actual, dir) {
  const st = stackPolea();
  if (!st.length) return Math.max(0, red2(actual + dir * incrementoPeso()));
  if (dir > 0) { const n = st.find(w => w > actual + 0.001); return n != null ? n : st[st.length - 1]; }
  const p = [...st].reverse().find(w => w < actual - 0.001); return p != null ? p : st[0];
}
function esPolea(e) { return !!e && String(e.equipamiento).toLowerCase() === 'polea'; }

// Paso de los botones +/− de peso según el ejercicio: con barra el peso va
// equilibrado (un disco a cada lado), así que el salto mínimo es el doble
// del incremento base (2×1,25 = 2,5 kg); el resto usa el incremento tal cual.
function pasoPeso(e) {
  const inc = incrementoPeso();
  return (e && String(e.equipamiento).toLowerCase() === 'barra') ? red2(inc * 2) : inc;
}

function hoyISO() { return new Date().toISOString().slice(0, 10); }

// Series ya registradas hoy para un ejercicio, ordenadas por nº de serie.
// Con `lado` (Izq/Der) cuenta solo ese lado.
function seriesDeHoy(idEjercicio, lado) {
  const h = hoyISO();
  return state.data.registro
    .filter(r => r.id === idEjercicio && r.fecha === h && (!lado || r.lado === lado))
    .sort((a, b) => a.serie - b.serie);
}

// Redondea a 2 decimales para evitar arrastres de coma flotante (1.25 + 1.25…).
function red2(n) { return Math.round(n * 100) / 100; }

function fmtPeso(n) {
  // Hasta 2 decimales, sin ceros sobrantes, con coma decimal (1.25 → "1,25"; 107.5 → "107,5").
  return red2(n).toFixed(2).replace(/\.?0+$/, '').replace('.', ',');
}

// Escapa texto para insertarlo con seguridad en HTML (nombres editables por el usuario).
function esc(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function ultimaFechaEntreno() {
  let max = '';
  for (const r of state.data.registro) if (r.fecha > max) max = r.fecha;
  return max || null;
}

function fmtFecha(iso) {
  if (!iso) return '—';
  const [a, m, d] = iso.split('-');
  return `${d}/${m}/${a}`;
}

function fmtFechaHora(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '—';
  const p = (n) => String(n).padStart(2, '0');
  return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

// ===== Vistas =====
function render() {
  const v = document.getElementById('view');
  document.querySelectorAll('nav button').forEach(b =>
    b.classList.toggle('active', b.dataset.tab === state.tab));
  timerVisible();   // el temporizador solo tiene sentido en Hoy
  if (!state.data && state.tab !== 'ajustes') {
    v.innerHTML = `
      <div class="card dato-grande">
        <div class="valor">👋</div>
        <div class="etiqueta">Aún no hay datos cargados.<br>
        Ve a <b>Ajustes</b> y conecta con Dropbox o importa el Excel.</div>
      </div>`;
    return;
  }
  if (state.tab === 'hoy') renderHoy(v);
  else if (state.tab === 'historial') renderHistorial(v);
  else if (state.tab === 'progreso') renderProgreso(v);
  else if (state.tab === 'ejercicios') renderEjercicios(v);
  else if (state.tab === 'rutinas') renderRutinas(v);
  else renderAjustes(v);
  pintarBadge();
  pintarUndo();
  wireBuscadores();
}

// Valores precargados para un ejercicio (y lado): si ya hay series hoy, se parte
// de la última de hoy; si no, del último entrenamiento registrado.
function valoresBase(e, lado) {
  const hoy = seriesDeHoy(e.id, lado);
  const ref = hoy.length ? hoy[hoy.length - 1] : ultimaSerie(e.id, lado);
  return {
    peso: ref ? red2(ref.peso) : 0,
    reps: ref ? ref.reps : (e.repsObj || 0),
  };
}

// Lista de ejercicios (objetos) a mostrar hoy según la rutina elegida.
function ejerciciosDeHoy() {
  const rutinas = state.data.rutinas || [];
  let lista;
  if (state.rutinaHoy === '__libre__') {
    lista = state.data.ejercicios.filter(e => e.activo);
  } else {
    const r = rutinas.find(x => x.nombre === state.rutinaHoy);
    lista = (r ? r.ids : [])
      .map(id => state.data.ejercicios.find(e => e.id === id))
      .filter(e => e && e.activo);
  }
  // Añadir ejercicios sueltos del día (sin duplicar)
  const extras = state.extras
    .map(id => state.data.ejercicios.find(e => e.id === id))
    .filter(e => e && e.activo && !lista.includes(e));
  return lista.concat(extras);
}

function elegirRutina(nombre) {
  state.rutinaHoy = nombre;
  state.extras = [];
  guardarRutinaHoy();
  render();
}

function anadirSuelto(id) {
  if (id && !state.extras.includes(id)) { state.extras.push(id); guardarRutinaHoy(); }
  render();
}

// IDs de ejercicios con al menos una serie registrada hoy, en orden de realización.
function idsHechosHoy() {
  const h = hoyISO();
  const ids = [];
  state.data.registro.forEach(r => { if (r.fecha === h && !ids.includes(r.id)) ids.push(r.id); });
  return ids;
}

function renderHoy(v) {
  const stepR = incrementoReps();
  const rutinas = state.data.rutinas || [];

  // Si la rutina elegida ya no existe (renombrada/borrada), volver al selector.
  if (state.rutinaHoy && state.rutinaHoy !== '__libre__'
      && !rutinas.some(r => r.nombre === state.rutinaHoy)) {
    state.rutinaHoy = null;
  }

  let html = `<h2>Hoy · ${fmtFecha(hoyISO())}</h2>`;

  // Sección "Hechos hoy" (compacta), encima del selector: no desaparecen al cambiar de rutina.
  const hechosIds = idsHechosHoy();
  const hechos = hechosIds.map(id => state.data.ejercicios.find(e => e.id === id)).filter(Boolean);
  if (hechos.length) {
    html += `<div class="hechos-hoy"><div class="hechos-tit">Hechos hoy</div>`
      + hechos.map(e => tarjetaHecha(e, pasoPeso(e), stepR)).join('') + `</div>`;
  }

  // Selector de rutina
  const chips = rutinas.map(r =>
    `<button class="chip-sel ${state.rutinaHoy === r.nombre ? 'sel' : ''}" data-rutina="${esc(r.nombre)}">${esc(r.nombre)}</button>`
  ).join('') +
    `<button class="chip-sel ${state.rutinaHoy === '__libre__' ? 'sel' : ''}" data-rutina="__libre__">Libre</button>`;
  html += `<div class="selector-rutina">${chips}</div>`;

  if (!state.rutinaHoy) {
    html += hechos.length
      ? '<p class="nota">Elige una rutina para seguir, o añade más series arriba.</p>'
      : '<p class="nota">Elige tu entrenamiento de hoy para empezar a registrar.</p>';
    v.innerHTML = html + fichaOverlay(stepR);
    bindHoy(v);
    return;
  }

  // Pendientes: ejercicios de la rutina/extras que aún NO se han hecho hoy.
  const hechosSet = new Set(hechosIds);
  const pendientes = ejerciciosDeHoy().filter(e => !hechosSet.has(e.id));
  html += pendientes.length
    ? pendientes.map(e => tarjetaEntrada(e)).join('')
    : (hechos.length ? '<p class="nota">¡Rutina completada! Puedes añadir otro ejercicio abajo.</p>'
                     : '<p class="nota">Esta rutina no tiene ejercicios activos. Edítala en la pestaña Rutinas.</p>');

  // Añadir un ejercicio suelto que no esté ni hecho ni pendiente.
  const presentes = new Set([...hechosSet, ...pendientes.map(e => e.id)]);
  const restantes = state.data.ejercicios.filter(e => e.activo && !presentes.has(e.id));
  if (restantes.length) {
    html += `<input class="buscador buscar-sel" type="search" placeholder="Buscar ejercicio…" data-target="add-suelto" autocapitalize="off">
    <select id="add-suelto" class="add-suelto">
      <option value="">+ Añadir otro ejercicio…</option>
      ${restantes.map(e => `<option value="${e.id}">${esc(e.nombre)}</option>`).join('')}
    </select>`;
  }

  v.innerHTML = html + fichaOverlay(stepR);
  bindHoy(v);
}

// HTML de la ficha enfocada si hay un ejercicio abierto ('' si no).
function fichaOverlay(stepR) {
  if (!state.fichaAbierta) return '';
  const e = ejPorId(state.fichaAbierta);
  if (!e) { state.fichaAbierta = null; return ''; }
  return fichaHtml(e, stepR);
}

// Tarjeta compacta de un ejercicio ya trabajado hoy: resumen de series;
// "+ Añadir serie" abre la ficha enfocada (pantalla completa).
function tarjetaHecha(e, stepP, stepR) {
  const unilateral = String(e.lateralidad).toLowerCase() === 'unilateral';
  // Resumen + referencia del mejor set al peso usado hoy (por lado).
  const bloque = (lado, etiqueta) => {
    const ss = seriesDeHoy(e.id, lado);
    const pills = ss.map(s => `<span class="serie-pill${clasePR(s)}">S${s.serie}: ${fmtPeso(s.peso)}kg × ${s.reps}</span>`).join('') || '—';
    const pesoHoy = ss.length ? ss[ss.length - 1].peso : 0;
    return `<div class="hh-resumen">${etiqueta ? `<span class="hh-lado">${etiqueta}</span> ` : ''}${pills}</div>`
      + refMejorHtml(e.id, lado, pesoHoy);
  };
  const cuerpo = unilateral ? bloque('Izq', 'izq') + bloque('Der', 'der') : bloque('', '');

  return `<div class="card entrada hecha" data-ficha="${e.id}">
    <div class="titulo">
      <span class="nombre">✓ ${esc(e.nombre)}</span>
      <button class="btn-mini-add" data-ficha="${e.id}">+ Añadir serie</button>
    </div>
    ${cuerpo}
  </div>`;
}

// ===== Ficha enfocada (pantalla completa) =====
// Se abre al guardar una serie (o con "+ Añadir serie"): el ejercicio en curso
// ocupa toda la pantalla, sin scroll. En unilaterales, pestañas Izq/Der con un
// solo juego de controles (el lado NO cambia solo al guardar).
function abrirFicha(id) {
  if (state.fichaAbierta !== id) {
    if (!state.fichaAbierta) history.pushState({ ficha: id }, '');  // "atrás" cierra
    state.fichaAbierta = id;
  }
  render();
}

function cerrarFicha(desdeAtras) {
  if (!state.fichaAbierta) return;
  const id = state.fichaAbierta;
  state.fichaAbierta = null;
  timerAlCerrarFicha(id);
  if (!desdeAtras && history.state && history.state.ficha) history.back();
  render();
}

// Al cerrar la ficha, ¿tiene sentido seguir descansando? Si ya has hecho las
// series objetivo del ejercicio, el descanso sobra y se detiene. Si te faltan
// series (o el ejercicio no tiene objetivo), sigue corriendo: así una superserie
// con otro ejercicio conserva la cuenta atrás. Nunca detiene un descanso que
// arrancó OTRO ejercicio.
function timerAlCerrarFicha(id) {
  if (!(timerFin > Date.now())) return;                    // no hay cuenta en marcha
  if (String(timerUltimo).split('|')[0] !== id) return;     // la arrancó otro ejercicio
  const e = ejPorId(id);
  if (!e || !(e.seriesObj > 0)) return;                    // sin objetivo, no decidimos
  const hechas = new Set(state.data.registro
    .filter(r => r.id === id && r.fecha === hoyISO())
    .map(r => r.serie)).size;                              // izq+der cuentan como una
  if (hechas >= e.seriesObj) timerParar();
}

window.addEventListener('popstate', () => { if (state.fichaAbierta) cerrarFicha(true); });

function fichaHtml(e, stepR) {
  const stepP = pasoPeso(e);
  const unilateral = String(e.lateralidad).toLowerCase() === 'unilateral';
  let cuerpo;
  if (unilateral) {
    const lado = state.fichaLado[e.id] || 'Izq';
    const tab = (l, txt) => {
      const n = seriesDeHoy(e.id, l).length;
      return `<button class="chip-sel lado-tab ${lado === l ? 'sel' : ''}" data-ficha-lado="${e.id}|${l}">
        ${txt}${n ? ` (${n})` : ''}</button>`;
    };
    cuerpo = `<div class="selector-rutina lado-tabs">${tab('Izq', 'Izquierda')}${tab('Der', 'Derecha')}</div>`
      + bloqueEntrada(e, lado, stepP, stepR);
  } else {
    cuerpo = bloqueEntrada(e, '', stepP, stepR);
  }
  return `
    <div class="ficha-sheet">
      <div class="ficha-cab">
        <span class="nombre">${esc(e.nombre)}</span>
        <button class="ficha-cerrar" data-ficha-cerrar="1" title="Cerrar">✕</button>
      </div>
      ${metaDescanso(e)}${metaNota(e)}
      ${cuerpo}
    </div>`;
}

function bindHoy(v) {
  v.querySelectorAll('[data-rutina]').forEach(btn =>
    btn.onclick = () => elegirRutina(btn.dataset.rutina));
  v.querySelectorAll('[data-paso]').forEach(btn =>
    btn.onclick = () => ajustar(btn.dataset.id, btn.dataset.campo, Number(btn.dataset.paso), btn.dataset.lado || ''));
  v.querySelectorAll('[data-edit]').forEach(inp => {
    inp.onchange = () => editarEntrada(inp.dataset.id, inp.dataset.edit, inp.value, inp.dataset.lado || '');
    inp.onfocus = () => inp.select();
  });
  v.querySelectorAll('[data-guardar]').forEach(btn =>
    btn.onclick = () => guardarSerie(btn.dataset.guardar, btn.dataset.guardarLado || ''));
  v.querySelectorAll('[data-ficha]').forEach(btn =>
    btn.onclick = () => abrirFicha(btn.dataset.ficha));
  v.querySelectorAll('[data-ficha-cerrar]').forEach(btn =>
    btn.onclick = () => cerrarFicha(false));
  v.querySelectorAll('[data-ficha-lado]').forEach(btn => btn.onclick = () => {
    const [id, lado] = btn.dataset.fichaLado.split('|');
    state.fichaLado[id] = lado;
    render();
  });
  const add = document.getElementById('add-suelto');
  if (add) add.onchange = () => anadirSuelto(add.value);
}

// Línea con el descanso entre series del ejercicio (del catálogo).
function metaDescanso(e) {
  return e.descanso ? `<div class="entrada-meta">⏱ Descanso ${fmtPeso(e.descanso)} min</div>` : '';
}

// Nota del ejercicio (del catálogo), visible al entrenar sin ir a la pestaña Ejercicios.
function metaNota(e) {
  return e.notas ? `<p class="ej-notas">${esc(e.notas)}</p>` : '';
}

// Tarjeta plegada de un ejercicio pendiente: información básica; tocarla abre
// la ficha enfocada, donde se introducen las series.
function tarjetaEntrada(e) {
  const unilateral = String(e.lateralidad).toLowerCase() === 'unilateral';
  const u = ultimoPeso(e.id, '');
  return `
    <div class="card entrada plegada" data-ficha="${e.id}">
      <div class="titulo">
        <span class="nombre">${esc(e.nombre)}</span>
        <span class="ultimo-peso">${u ? fmtPeso(u.peso) + ' kg' : '—'}
          <small>${u ? 'último (' + fmtFecha(u.fecha) + ')' : 'sin registros'}</small>
        </span>
      </div>
      <div class="chips">
        ${e.seriesObj || e.repsObj ? `<span class="chip">objetivo ${e.seriesObj}×${e.repsObj}</span>` : ''}
        ${e.descanso ? `<span class="chip">descanso ${fmtPeso(e.descanso)}′</span>` : ''}
        ${unilateral ? '<span class="chip chip-coral">Unilateral · por lado</span>' : ''}
      </div>
      <span class="plegada-chev">›</span>
    </div>`;
}

// Controles de peso/reps + guardar + pills de hoy, para un lado ('' = bilateral).
function bloqueEntrada(e, lado, stepP, stepR) {
  const key = entradaKey(e.id, lado);
  const base = state.entrada[key] || (state.entrada[key] = valoresBase(e, lado));
  const hoy = seriesDeHoy(e.id, lado);
  const proxSerie = hoy.length + 1;
  const etiqueta = lado === 'Izq' ? 'Izquierda' : lado === 'Der' ? 'Derecha' : '';
  const guardarTxt = lado ? `Guardar ${lado.toLowerCase()} ${proxSerie}` : `Guardar serie ${proxSerie}`;

  let cabLado = '';
  if (lado) {
    const u = ultimoPeso(e.id, lado);
    cabLado = `<div class="lado-cab"><span class="lado-tit">${etiqueta}</span>
      <span class="ultimo-peso">${u ? fmtPeso(u.peso) + ' kg' : '—'}
        <small>${u ? 'último (' + fmtFecha(u.fecha) + ')' : 'sin registros'}</small></span></div>`;
  }

  const seriesHoyHtml = hoy.length ? `
    <div class="series-hoy">
      ${hoy.map(s => `<span class="serie-pill${clasePR(s)}">S${s.serie}: ${fmtPeso(s.peso)}kg × ${s.reps}</span>`).join('')}
    </div>` : '';

  return `
    ${cabLado}
    ${refUltimaHtml(e.id, lado)}
    ${refMejorHtml(e.id, lado, base.peso)}
    <div class="control">
      <button class="paso" data-id="${e.id}" data-lado="${lado}" data-campo="peso" data-paso="${-stepP}">−</button>
      <div class="lectura"><input class="num num-input" type="text" inputmode="decimal" data-edit="peso" data-id="${e.id}" data-lado="${lado}" value="${fmtPeso(base.peso)}"><span class="ud">kg</span></div>
      <button class="paso" data-id="${e.id}" data-lado="${lado}" data-campo="peso" data-paso="${stepP}">+</button>
    </div>
    <div class="control">
      <button class="paso" data-id="${e.id}" data-lado="${lado}" data-campo="reps" data-paso="${-stepR}">−</button>
      <div class="lectura"><input class="num num-input" type="text" inputmode="numeric" data-edit="reps" data-id="${e.id}" data-lado="${lado}" value="${base.reps}"><span class="ud">reps${lado ? '/lado' : ''}</span></div>
      <button class="paso" data-id="${e.id}" data-lado="${lado}" data-campo="reps" data-paso="${stepR}">+</button>
    </div>
    <button class="btn" data-guardar="${e.id}" data-guardar-lado="${lado}">${guardarTxt}</button>
    ${seriesHoyHtml}`;
}

function ajustar(id, campo, paso, lado) {
  const key = entradaKey(id, lado);
  const v = state.entrada[key] || (state.entrada[key] = { peso: 0, reps: 0 });
  if (campo === 'peso') {
    // En polea, el +/- salta por las placas del stack; en el resto, incremento fijo.
    v.peso = esPolea(ejPorId(id)) ? pasoPolea(v.peso, paso > 0 ? 1 : -1)
                                  : Math.max(0, red2(v.peso + paso));
  } else {
    v.reps = Math.max(0, v.reps + paso);
  }
  render();
}

// Edición manual del peso/reps tocando el campo (cualquier valor, sin snap).
function editarEntrada(id, campo, valor, lado) {
  const key = entradaKey(id, lado);
  const v = state.entrada[key] || (state.entrada[key] = { peso: 0, reps: 0 });
  if (campo === 'peso') v.peso = Math.max(0, red2(Number(String(valor).replace(',', '.')) || 0));
  else v.reps = Math.max(0, Math.round(Number(valor) || 0));
  render();
}

function guardarSerie(id, lado) {
  const e = state.data.ejercicios.find(x => x.id === id);
  const key = entradaKey(id, lado);
  const val = state.entrada[key] || valoresBase(e, lado);
  const serie = seriesDeHoy(id, lado).length + 1;
  state.data.registro.push({
    fecha: hoyISO(), id, ejercicio: e.nombre, serie,
    lado: lado || '', reps: val.reps, peso: red2(val.peso),
    rutina: rutinaParaGuardar(), notas: '',
  });
  timerDesdeSerie(e, id, serie, lado || '');   // arranca el descanso
  if (lado) state.fichaLado[id] = lado;        // la ficha se queda en el lado usado
  // ¿Este set acaba de batir el récord a ese peso (+3 al total)? → animación.
  const nueva = state.data.registro[state.data.registro.length - 1];
  const { total, mejorPrev } = statsSet(nueva);
  const cruzaAhora = mejorPrev >= 0 && total >= mejorPrev + PR_MARGEN
    && (total - nueva.reps) < mejorPrev + PR_MARGEN;
  state.flashPR = cruzaAhora ? setKey(nueva) : null;
  saveData();
  marcarPendiente(true);
  abrirFicha(id);   // al guardar, el ejercicio pasa a la ficha enfocada (y render)
  if (state.flashPR) setTimeout(() => { state.flashPR = null; }, 1400);
}

function renderEjercicios(v) {
  // Orden manual (el del array, que se guarda en el Excel).
  const ejs = state.data.ejercicios;

  let html = '<h2>Catálogo</h2>';
  html += `<input id="ej-buscar" class="buscador" type="search" placeholder="Buscar ejercicio…" value="${esc(state.ejBuscar || '')}" autocapitalize="off">`;
  html += `<div class="orden-bar"><span class="orden-lbl">Ordenar por:</span>
    <button class="chip-sel" data-ordenar="nombre">Nombre</button>
    <button class="chip-sel" data-ordenar="grupo">Grupo</button>
    <button class="chip-sel" data-ordenar="uso">Último uso</button></div>`;
  html += ejs.map((e, i) =>
    state.editEj === e.id ? formEjercicio(e) : tarjetaEjercicio(e, i, ejs.length)).join('');

  html += state.editEj === '__nuevo__'
    ? formEjercicio(null)
    : '<button class="btn" id="nuevo-ej">+ Nuevo ejercicio</button>';

  html += `<p class="nota">El catálogo se guarda en la hoja <b>Ejercicios</b> del Excel.
    Al borrar un ejercicio, las series ya registradas se conservan en el historial.</p>`;
  v.innerHTML = html;
  bindEjercicios(v);
}

function tarjetaEjercicio(e, i, n) {
  const u = ultimoPeso(e.id);
  const unilateral = String(e.lateralidad).toLowerCase() === 'unilateral';
  return `
    <div class="card${e.activo ? '' : ' inactivo'}" data-buscar="${esc((e.nombre + ' ' + e.grupo).toLowerCase())}">
      <div class="titulo">
        <span class="nombre">${esc(e.nombre)}</span>
        <span class="ultimo-peso">${u ? fmtPeso(u.peso) + ' kg' : '—'}
          <small>${u ? 'último (' + fmtFecha(u.fecha) + ')' : 'sin registros'}</small>
        </span>
      </div>
      <div class="chips">
        <span class="chip">${esc(e.grupo)}</span>
        <span class="chip">${esc(e.equipamiento)}</span>
        ${esPolea(e) && e.polea12 ? '<span class="chip">Polea 1/2</span>' : ''}
        ${unilateral ? '<span class="chip chip-coral">Unilateral · reps por lado</span>' : ''}
        <span class="chip">${e.seriesObj}×${e.repsObj}</span>
        <span class="chip">descanso ${e.descanso}′</span>
        ${e.activo ? '' : '<span class="chip chip-coral">Inactivo</span>'}
      </div>
      ${e.notas ? `<p class="ej-notas">${esc(e.notas)}</p>` : ''}
      <div class="ej-fila-acc">
        <button class="mini" data-mover-ej="${i}:-1" ${i === 0 ? 'disabled' : ''}>▲</button>
        <button class="mini" data-mover-ej="${i}:1" ${i === n - 1 ? 'disabled' : ''}>▼</button>
        <button class="btn btn-sec" data-edit="${e.id}">✏️ Editar</button>
      </div>
    </div>`;
}

function moverEjercicioCatalogo(i, dir) {
  const a = state.data.ejercicios, j = i + dir;
  if (j < 0 || j >= a.length) return;
  [a[i], a[j]] = [a[j], a[i]];
  saveData(); marcarPendiente(true); render();
}

function ordenarCatalogo(criterio) {
  const a = state.data.ejercicios;
  const cmpNom = (x, y) => String(x.nombre).localeCompare(String(y.nombre), 'es');
  const usoFecha = (e) => { const u = ultimoPeso(e.id); return u ? u.fecha : ''; };
  const cmp = {
    nombre: cmpNom,
    grupo: (x, y) => String(x.grupo).localeCompare(String(y.grupo), 'es') || cmpNom(x, y),
    uso: (x, y) => usoFecha(y).localeCompare(usoFecha(x)) || cmpNom(x, y), // más reciente primero
  }[criterio];
  if (!cmp) return;
  // Inactivos al final; dentro de cada bloque, por el criterio elegido.
  a.sort((x, y) => (x.activo === y.activo ? cmp(x, y) : (x.activo ? -1 : 1)));
  saveData(); marcarPendiente(true); render();
}

// Formulario de alta/edición. e=null → ejercicio nuevo.
function formEjercicio(e) {
  const nuevo = !e;
  const val = e || { nombre: '', grupo: '', equipamiento: 'Barra', lateralidad: 'Bilateral',
    descanso: 2, seriesObj: 4, repsObj: 10, activo: true, notas: '' };
  const ops = (arr, sel) => arr.map(o =>
    `<option value="${o}" ${o === sel ? 'selected' : ''}>${o}</option>`).join('');
  return `
    <div class="card ej-form">
      <label for="f-nombre">Ejercicio</label>
      <input type="text" id="f-nombre" value="${esc(val.nombre)}" autocapitalize="sentences">
      <label for="f-grupo">Grupo muscular</label>
      <input type="text" id="f-grupo" list="grupos-lista" value="${esc(val.grupo)}" autocapitalize="sentences" autocomplete="off">
      <datalist id="grupos-lista">${gruposExistentes().map(g => `<option value="${esc(g)}"></option>`).join('')}</datalist>
      <div class="fila2">
        <div>
          <label for="f-equip">Equipamiento</label>
          <select id="f-equip" class="campo">${ops(EQUIPOS, val.equipamiento)}</select>
        </div>
        <div>
          <label for="f-lat">Lateralidad</label>
          <select id="f-lat" class="campo">${ops(LATERALIDADES, val.lateralidad)}</select>
        </div>
      </div>
      <label class="ej-activo${esPolea(val) ? '' : ' oculto'}" id="f-p12-wrap">
        <input type="checkbox" id="f-p12" ${val.polea12 ? 'checked' : ''}>
        Polea con proporción 1/2
      </label>
      <div class="fila2">
        <div>
          <label for="f-desc">Descanso (min)</label>
          <input type="number" id="f-desc" value="${val.descanso}" min="0" step="0.5" inputmode="decimal">
        </div>
        <div>
          <label for="f-series">Series obj.</label>
          <input type="number" id="f-series" value="${val.seriesObj}" min="0" step="1" inputmode="numeric">
        </div>
        <div>
          <label for="f-reps">Reps obj.</label>
          <input type="number" id="f-reps" value="${val.repsObj}" min="0" step="1" inputmode="numeric">
        </div>
      </div>
      <label for="f-notas">Notas</label>
      <input type="text" id="f-notas" value="${esc(val.notas)}" autocapitalize="sentences">
      <label class="ej-activo"><input type="checkbox" id="f-activo" ${val.activo ? 'checked' : ''}> Activo (aparece al entrenar)</label>
      <div class="ej-acciones">
        <button class="btn btn-sec" id="ej-cancelar">Cancelar</button>
        <button class="btn" id="ej-guardar">${nuevo ? 'Crear' : 'Guardar'}</button>
      </div>
      ${nuevo ? '' : '<button class="btn btn-coral" id="ej-borrar">Borrar ejercicio</button>'}
    </div>`;
}

function bindEjercicios(v) {
  const nuevo = document.getElementById('nuevo-ej');
  if (nuevo) nuevo.onclick = () => { state.editEj = '__nuevo__'; render(); };
  v.querySelectorAll('[data-edit]').forEach(btn =>
    btn.onclick = () => { state.editEj = btn.dataset.edit; render(); });
  v.querySelectorAll('[data-mover-ej]').forEach(btn => {
    const [i, dir] = btn.dataset.moverEj.split(':').map(Number);
    btn.onclick = () => moverEjercicioCatalogo(i, dir);
  });
  v.querySelectorAll('[data-ordenar]').forEach(btn =>
    btn.onclick = () => ordenarCatalogo(btn.dataset.ordenar));

  // La casilla "Polea 1/2" solo tiene sentido con equipamiento Polea.
  const equipSel = document.getElementById('f-equip');
  if (equipSel) equipSel.onchange = () => {
    document.getElementById('f-p12-wrap').classList.toggle('oculto',
      equipSel.value.toLowerCase() !== 'polea');
  };

  const cancelar = document.getElementById('ej-cancelar');
  if (cancelar) cancelar.onclick = () => { state.editEj = null; render(); };
  const guardar = document.getElementById('ej-guardar');
  if (guardar) guardar.onclick = guardarEjercicio;
  const borrar = document.getElementById('ej-borrar');
  if (borrar) borrar.onclick = () => borrarEjercicio(state.editEj);
}

// Grupos musculares ya usados en el catálogo (para sugerir, sin cerrar la lista).
function gruposExistentes() {
  return [...new Set(state.data.ejercicios.map(e => e.grupo).filter(g => g && String(g).trim()))]
    .sort((a, b) => String(a).localeCompare(String(b), 'es'));
}

// Genera el siguiente ID con el patrón Exx (E01, E02…), tolerando IDs ajenos.
function nuevoIdEjercicio() {
  let max = 0;
  for (const e of state.data.ejercicios) {
    const m = String(e.id).match(/^E(\d+)$/i);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return 'E' + String(max + 1).padStart(2, '0');
}

function guardarEjercicio() {
  const num = (id) => Math.max(0, Number(document.getElementById(id).value) || 0);
  const nombre = document.getElementById('f-nombre').value.trim();
  if (!nombre) { alert('El ejercicio necesita un nombre.'); return; }

  const datos = {
    nombre,
    grupo: document.getElementById('f-grupo').value.trim(),
    equipamiento: document.getElementById('f-equip').value,
    lateralidad: document.getElementById('f-lat').value,
    descanso: num('f-desc'),
    seriesObj: num('f-series'),
    repsObj: num('f-reps'),
    activo: document.getElementById('f-activo').checked,
    notas: document.getElementById('f-notas').value.trim(),
  };
  // Solo aplica a poleas: si se cambia el equipamiento, la marca 1/2 se limpia.
  datos.polea12 = datos.equipamiento.toLowerCase() === 'polea'
    && document.getElementById('f-p12').checked;

  if (state.editEj === '__nuevo__') {
    state.data.ejercicios.push(Object.assign({ id: nuevoIdEjercicio() }, datos));
  } else {
    const e = state.data.ejercicios.find(x => x.id === state.editEj);
    if (!e) { state.editEj = null; render(); return; }
    Object.assign(e, datos);
    // El nombre está desnormalizado en el Registro: lo mantenemos al día.
    state.data.registro.forEach(r => { if (r.id === e.id) r.ejercicio = nombre; });
  }
  state.editEj = null;
  saveData();
  marcarPendiente(true);
  render();
}

function borrarEjercicio(id) {
  const e = state.data.ejercicios.find(x => x.id === id);
  if (!e) return;
  const nSeries = state.data.registro.filter(r => r.id === id).length;
  const enRutinas = (state.data.rutinas || []).filter(r => r.ids.includes(id)).map(r => r.nombre);
  let msg = `¿Borrar "${e.nombre}" del catálogo?`;
  if (nSeries) msg += `\n\nTiene ${nSeries} serie(s) en el historial: se conservan en el Registro, pero el ejercicio dejará de aparecer. Si solo quieres dejar de verlo al entrenar, márcalo como inactivo en su lugar.`;
  if (enRutinas.length) msg += `\n\nSe quitará de las rutinas: ${enRutinas.join(', ')}.`;
  if (!confirmar(msg)) return;

  state.data.ejercicios = state.data.ejercicios.filter(x => x.id !== id);
  (state.data.rutinas || []).forEach(r => { r.ids = r.ids.filter(x => x !== id); });
  state.extras = state.extras.filter(x => x !== id);
  delete state.entrada[id];
  state.editEj = null;
  saveData();
  marcarPendiente(true);
  render();
}

// ===== Historial (agenda por mes / semana / día) =====
const DIAS_SEM = ['lun', 'mar', 'mié', 'jue', 'vie', 'sáb', 'dom'];
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
               'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

function isoADate(iso) { const [a, m, d] = iso.split('-').map(Number); return new Date(a, m - 1, d); }
function dateAIso(dt) {
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
}
function idxSemana(dt) { return (dt.getDay() + 6) % 7; }  // 0=lun … 6=dom
function lunesDe(dt) { const x = new Date(dt); x.setHours(0, 0, 0, 0); x.setDate(x.getDate() - idxSemana(x)); return x; }
function sumarDias(dt, n) { const x = new Date(dt); x.setDate(x.getDate() + n); return x; }

// Número de semana ISO 8601 (1–53); las semanas empiezan en lunes y la semana 1
// es la que contiene el primer jueves del año.
function numeroSemanaISO(dt) {
  const d = new Date(Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()));
  const dia = (d.getUTCDay() + 6) % 7;          // 0=lunes … 6=domingo
  d.setUTCDate(d.getUTCDate() - dia + 3);        // jueves de esta semana ISO
  const primerJueves = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  const diaPJ = (primerJueves.getUTCDay() + 6) % 7;
  primerJueves.setUTCDate(primerJueves.getUTCDate() - diaPJ + 3);
  return 1 + Math.round((d - primerJueves) / (7 * 24 * 3600 * 1000));
}

// Rutina asociada a un día: la elegida para un día nuevo, o la de sus filas.
function rutinaDelDia(iso) {
  if (state.histRutinaNueva[iso]) return state.histRutinaNueva[iso];
  const r = state.data.registro.find(x => x.fecha === iso && x.rutina);
  return r ? r.rutina : '';
}

function renderHistorial(v) {
  // Índice de filas por fecha, guardando el índice REAL dentro de registro.
  const porFecha = new Map();
  state.data.registro.forEach((r, i) => {
    if (!porFecha.has(r.fecha)) porFecha.set(r.fecha, []);
    porFecha.get(r.fecha).push(i);
  });

  if (!porFecha.size) {
    v.innerHTML = '<h2>Historial</h2><p class="nota">Aún no hay entrenamientos registrados. Empieza en la pestaña Hoy.</p>';
    return;
  }

  const hoy = isoADate(hoyISO());
  const primera = isoADate([...porFecha.keys()].sort()[0]);
  let html = '<h2>Historial</h2>';
  let mesActual = null;

  for (let lunes = lunesDe(hoy); lunes >= lunesDe(primera); lunes = sumarDias(lunes, -7)) {
    const dias = [];
    for (let k = 0; k < 7; k++) {
      const dt = sumarDias(lunes, k);
      if (dt <= hoy) dias.push(dt);   // no mostramos días futuros
    }
    if (!dias.length) continue;

    const claveMes = `${lunes.getFullYear()}-${lunes.getMonth()}`;
    if (claveMes !== mesActual) {
      mesActual = claveMes;
      html += `<div class="hist-mes">${MESES[lunes.getMonth()]} ${lunes.getFullYear()}</div>`;
    }

    const rotulo = `Semana ${numeroSemanaISO(lunes)}`;
    const conEntreno = dias.some(dt => porFecha.has(dateAIso(dt)));

    if (!conEntreno) {
      html += `<div class="hist-sem vacia">${rotulo} · sin entrenamientos</div>`;
      continue;
    }
    html += `<div class="hist-sem">${rotulo}</div>`;
    // Días de más reciente a más antiguo (hoy arriba del todo).
    for (let k = dias.length - 1; k >= 0; k--) html += filaDiaHist(dias[k], porFecha);
  }

  v.innerHTML = html;
  bindHistorial(v);
}

function filaDiaHist(dt, porFecha) {
  const iso = dateAIso(dt);
  const esHoy = iso === hoyISO();
  const esDomingo = idxSemana(dt) === 6;
  const fechaCls = esDomingo ? 'hist-fecha domingo' : 'hist-fecha';
  const dow = DIAS_SEM[idxSemana(dt)];
  const idxs = porFecha.get(iso);

  if (!idxs) {
    if (state.histAbierto === iso) return diaVacioAbierto(dt);
    return `<div class="hist-dia vacio${esHoy ? ' es-hoy' : ''}" data-abrir="${iso}">
      <div class="${fechaCls}"><span class="dow">${dow}</span><span class="dnum">${dt.getDate()}</span></div>
      <div class="hist-resumen">${esHoy ? 'hoy · ' : ''}descanso</div>
    </div>`;
  }

  const filas = idxs.map(i => state.data.registro[i]);
  const rutina = rutinaDelDia(iso) || 'Entreno Libre';
  const nEj = new Set(filas.map(f => f.id)).size;
  const nSeries = new Set(filas.map(f => f.id + '#' + f.serie)).size;
  const nReps = filas.reduce((s, f) => s + (f.reps || 0), 0);
  const abierto = state.histAbierto === iso;

  const cab = `<div class="hist-dia${esHoy ? ' es-hoy' : ''}${abierto ? ' abierto' : ''}" data-toggle="${iso}">
    <div class="${fechaCls}"><span class="dow">${dow}</span><span class="dnum">${dt.getDate()}</span></div>
    <div class="hist-info">
      <div class="hist-rutina">${esc(rutina)}${esHoy ? ' <span class="hoy-badge">hoy</span>' : ''}</div>
      <div class="hist-resumen">${nEj} ejercicios · ${nSeries} series · ${nReps} reps</div>
    </div>
    <span class="hist-chev">${abierto ? '▴' : '▸'}</span>
  </div>`;
  // Abierto: cabecera y detalle van dentro de una sola tarjeta unificada.
  if (abierto) return `<div class="hist-dia-card${esHoy ? ' es-hoy' : ''}">${cab}${diaExpandido(iso, idxs)}</div>`;
  return cab;
}

function diaExpandido(iso, idxs) {
  const orden = [];
  const porEj = new Map();
  idxs.forEach(i => {
    const r = state.data.registro[i];
    if (!porEj.has(r.id)) { porEj.set(r.id, []); orden.push(r.id); }
    porEj.get(r.id).push(i);
  });

  let html = '<div class="dia-detalle">';
  html += `<div class="rut-dia-edit"><label for="rd-${iso}">Rutina del día</label>
    <input type="text" id="rd-${iso}" class="rut-dia-input" data-rut-dia="${iso}" value="${esc(rutinaDelDia(iso) || 'Entreno Libre')}" autocapitalize="words"></div>`;
  orden.forEach(id => { html += bloqueEjercicioHist(iso, id, porEj.get(id)); });

  const presentes = new Set(orden);
  const restantes = state.data.ejercicios.filter(e => !presentes.has(e.id));
  if (restantes.length) {
    html += `<select class="add-suelto" data-add-ej-dia="${iso}">
      <option value="">+ Añadir ejercicio…</option>
      ${restantes.map(e => `<option value="${e.id}">${esc(e.nombre)}</option>`).join('')}
    </select>`;
  }
  html += `</div>`;
  return html;
}

function bloqueEjercicioHist(iso, id, idxs) {
  const e = state.data.ejercicios.find(x => x.id === id);
  const nombre = e ? e.nombre : (state.data.registro[idxs[0]].ejercicio || id);
  const unilateral = e && String(e.lateralidad).toLowerCase() === 'unilateral';
  let series = '';

  if (unilateral) {
    const porSerie = new Map();
    idxs.forEach(i => {
      const r = state.data.registro[i];
      if (!porSerie.has(r.serie)) porSerie.set(r.serie, {});
      porSerie.get(r.serie)[r.lado] = i;
    });
    [...porSerie.keys()].sort((a, b) => a - b).forEach(s => {
      const par = porSerie.get(s);
      series += `<div class="serie-uni">
        <span class="serie-n">S${s}</span>
        <span class="lado-vals">${ladoPillHist('Izq', par.Izq)}${ladoPillHist('Der', par.Der)}</span>
        <button class="mini mini-del" data-del-serie="${iso}|${id}|${s}" title="Borrar serie">🗑️</button>
      </div>`;
    });
  } else {
    series = '<div class="series-hist">' + idxs
      .sort((a, b) => state.data.registro[a].serie - state.data.registro[b].serie)
      .map(i => pillSerieHist(iso, id, i)).join('') + '</div>';
  }

  return `<div class="ej-hist">
    <div class="ej-hist-cab">
      <span class="ej-hist-nombre">${esc(nombre)}</span>
      <button class="mini mini-del" data-del-ej-dia="${iso}|${id}" title="Borrar ejercicio del día">🗑️</button>
    </div>
    ${series}
    <button class="add-serie" data-add-serie="${iso}|${id}">+ serie</button>
  </div>`;
}

function pillSerieHist(iso, id, i) {
  const r = state.data.registro[i];
  if (state.histEdit === i) {
    return `<span class="serie-edit">S${r.serie}
      <input class="he-peso" type="number" step="0.25" inputmode="decimal" value="${r.peso}"> kg ×
      <input class="he-reps" type="number" step="1" inputmode="numeric" value="${r.reps}">
      <button class="mini mini-ok" data-save-serie="${i}" title="Guardar">✓</button>
      <button class="mini mini-del" data-del-serie="${iso}|${id}|${r.serie}" title="Borrar serie">🗑️</button>
      <button class="mini mini-x" data-cancel-serie="1" title="Cancelar">✕</button>
    </span>`;
  }
  return `<span class="serie-pill editable${clasePR(r)}" data-edit-serie="${i}">S${r.serie}: ${fmtPeso(r.peso)}kg × ${r.reps}</span>`;
}

function ladoPillHist(lado, i) {
  if (i == null) return `<span class="lado-pill falta">${lado.toLowerCase()} —</span>`;
  const r = state.data.registro[i];
  if (state.histEdit === i) {
    return `<span class="serie-edit">${lado.toLowerCase()}
      <input class="he-peso" type="number" step="0.25" inputmode="decimal" value="${r.peso}"> kg ×
      <input class="he-reps" type="number" step="1" inputmode="numeric" value="${r.reps}">
      <button class="mini mini-ok" data-save-serie="${i}" title="Guardar">✓</button>
      <button class="mini mini-x" data-cancel-serie="1" title="Cancelar">✕</button>
    </span>`;
  }
  return `<span class="lado-pill editable${clasePR(r)}" data-edit-serie="${i}">${lado.toLowerCase()} ${fmtPeso(r.peso)}kg × ${r.reps}</span>`;
}

function diaVacioAbierto(dt) {
  const iso = dateAIso(dt);
  const esHoy = iso === hoyISO();
  const fechaCls = idxSemana(dt) === 6 ? 'hist-fecha domingo' : 'hist-fecha';
  const dow = DIAS_SEM[idxSemana(dt)];
  const rutinas = state.data.rutinas || [];
  const elegida = state.histRutinaNueva[iso];

  let panel;
  if (!elegida) {
    const chips = rutinas.map(r =>
      `<button class="chip-sel" data-rut-nueva="${iso}|${esc(r.nombre)}">${esc(r.nombre)}</button>`).join('')
      + `<button class="chip-sel" data-rut-nueva="${iso}|Entreno Libre">Libre</button>`;
    panel = `<p class="nota">Registrar un entreno este día. Elige la rutina:</p><div class="selector-rutina">${chips}</div>`;
  } else {
    panel = `<div class="hist-rutina">${esc(elegida)}</div>
      <select class="add-suelto" data-add-ej-dia="${iso}">
        <option value="">+ Añadir ejercicio…</option>
        ${state.data.ejercicios.filter(e => e.activo).map(e => `<option value="${e.id}">${esc(e.nombre)}</option>`).join('')}
      </select>`;
  }
  return `<div class="hist-dia-card${esHoy ? ' es-hoy' : ''}">
    <div class="hist-dia vacio abierto" data-toggle="${iso}">
      <div class="${fechaCls}"><span class="dow">${dow}</span><span class="dnum">${dt.getDate()}</span></div>
      <div class="hist-resumen">nuevo entreno</div>
      <span class="hist-chev">▴</span>
    </div><div class="dia-detalle">${panel}</div></div>`;
}

function bindHistorial(v) {
  v.querySelectorAll('[data-toggle]').forEach(el =>
    el.onclick = () => toggleDiaHist(el.dataset.toggle));
  v.querySelectorAll('[data-abrir]').forEach(el =>
    el.onclick = () => { state.histAbierto = el.dataset.abrir; state.histEdit = null; render(); });
  v.querySelectorAll('[data-rut-nueva]').forEach(el => el.onclick = () => {
    const [iso, nombre] = el.dataset.rutNueva.split('|');
    state.histRutinaNueva[iso] = nombre; render();
  });
  v.querySelectorAll('[data-edit-serie]').forEach(el => el.onclick = (ev) => {
    ev.stopPropagation(); state.histEdit = Number(el.dataset.editSerie); render();
  });
  v.querySelectorAll('[data-save-serie]').forEach(el => el.onclick = (ev) => {
    ev.stopPropagation(); guardarEdicionSerie(Number(el.dataset.saveSerie));
  });
  v.querySelectorAll('[data-cancel-serie]').forEach(el => el.onclick = (ev) => {
    ev.stopPropagation(); state.histEdit = null; render();   // descarta los cambios sin guardar
  });
  v.querySelectorAll('[data-del-serie]').forEach(el => el.onclick = (ev) => {
    ev.stopPropagation();
    const [iso, id, serie] = el.dataset.delSerie.split('|');
    borrarSerieHist(iso, id, Number(serie));
  });
  v.querySelectorAll('[data-del-ej-dia]').forEach(el => el.onclick = (ev) => {
    ev.stopPropagation();
    const [iso, id] = el.dataset.delEjDia.split('|');
    borrarEjercicioDia(iso, id);
  });
  v.querySelectorAll('[data-add-serie]').forEach(el => el.onclick = (ev) => {
    ev.stopPropagation();
    const [iso, id] = el.dataset.addSerie.split('|'); anadirSerieDia(iso, id);
  });
  v.querySelectorAll('[data-add-ej-dia]').forEach(el =>
    el.onchange = () => anadirSerieDia(el.dataset.addEjDia, el.value));
  v.querySelectorAll('[data-rut-dia]').forEach(inp => {
    inp.onclick = (ev) => ev.stopPropagation();
    inp.onchange = () => editarRutinaDia(inp.dataset.rutDia, inp.value);
  });
}

// Renombra la rutina de un día: actualiza todas las series de esa fecha.
function editarRutinaDia(iso, nombre) {
  nombre = (nombre || '').trim() || 'Entreno Libre';
  state.data.registro.forEach(r => { if (r.fecha === iso) r.rutina = nombre; });
  delete state.histRutinaNueva[iso];
  persistirRegistro();
}

function toggleDiaHist(iso) {
  state.histAbierto = (state.histAbierto === iso) ? null : iso;
  state.histEdit = null;
  render();
}

function guardarEdicionSerie(i) {
  const r = state.data.registro[i];
  if (!r) return;
  const peso = document.querySelector('.he-peso');
  const reps = document.querySelector('.he-reps');
  if (peso) r.peso = red2(Math.max(0, Number(peso.value) || 0));
  if (reps) r.reps = Math.max(0, Number(reps.value) || 0);
  state.histEdit = null;
  persistirRegistro();
}

function borrarSerieHist(iso, id, serie) {
  if (!confirmar(`¿Borrar la serie ${serie}?`)) return;
  state.data.registro = state.data.registro.filter(r => !(r.fecha === iso && r.id === id && r.serie === serie));
  renumerarSeries(iso, id);
  state.histEdit = null;
  persistirRegistro();
}

function borrarEjercicioDia(iso, id) {
  const e = state.data.ejercicios.find(x => x.id === id);
  const nombre = e ? e.nombre : id;
  if (!confirmar(`¿Borrar todas las series de "${nombre}" del ${fmtFecha(iso)}?`)) return;
  state.data.registro = state.data.registro.filter(r => !(r.fecha === iso && r.id === id));
  state.histEdit = null;
  persistirRegistro();
}

function anadirSerieDia(iso, id) {
  if (!id) return;
  const e = state.data.ejercicios.find(x => x.id === id);
  const unilateral = e && String(e.lateralidad).toLowerCase() === 'unilateral';
  const rutina = rutinaDelDia(iso) || 'Entreno Libre';
  const series = state.data.registro.filter(r => r.fecha === iso && r.id === id).map(r => r.serie);
  const next = (series.length ? Math.max(...series) : 0) + 1;
  (unilateral ? ['Izq', 'Der'] : ['']).forEach(lado => {
    const ref = ultimaSerie(id, lado);
    state.data.registro.push({
      fecha: iso, id, ejercicio: e ? e.nombre : id, serie: next, lado,
      reps: ref ? ref.reps : (e ? e.repsObj : 0), peso: ref ? red2(ref.peso) : 0,
      rutina, notas: '',
    });
  });
  persistirRegistro();
}

// Renumera las series de un ejercicio en un día a 1..n (izq/der comparten número).
function renumerarSeries(iso, id) {
  const series = [...new Set(state.data.registro
    .filter(r => r.fecha === iso && r.id === id).map(r => r.serie))].sort((a, b) => a - b);
  const mapa = new Map(series.map((s, i) => [s, i + 1]));
  state.data.registro.forEach(r => { if (r.fecha === iso && r.id === id) r.serie = mapa.get(r.serie); });
}

function persistirRegistro() { saveData(); marcarPendiente(true); render(); }

// ===== Progreso (estadísticas y gráficas) =====
const PAL = { azul: '#202A44', coral: '#C07A6B', lavanda: '#8388BA', neblina: '#E0DBE3' };

// Carga efectiva para el tonelaje: las mancuernas cuentan doble (un par); el
// resto, tal cual se apunta. (Decisión del proyecto.)
function cargaEfectiva(e, peso) {
  if (!e) return peso;
  if (String(e.equipamiento).toLowerCase() === 'mancuernas') return peso * 2;
  if (esPolea(e) && e.polea12) return peso / 2; // polea en proporción 1/2: carga real = mitad
  return peso;
}

// 1RM estimado (Epley): peso × (1 + reps/30).
function epley(peso, reps) { return peso * (1 + reps / 30); }

function ejPorId(id) { return state.data.ejercicios.find(e => e.id === id); }

// Rango [desde, hasta] (ISO) del periodo elegido, por calendario hasta hoy.
function rangoPeriodo(periodo) {
  const hoy = isoADate(hoyISO());
  let desde;
  if (periodo === 'semana') desde = lunesDe(hoy);
  else if (periodo === 'año') desde = new Date(hoy.getFullYear(), 0, 1);
  else desde = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  return { desde: dateAIso(desde), hasta: hoyISO() };
}

function filasEnRango(desde, hasta) {
  return state.data.registro.filter(r => r.fecha >= desde && r.fecha <= hasta);
}

// ¿La fila pertenece al grupo muscular filtrado en Progreso? ('' = todos).
function enGrupoProg(r) {
  if (!state.prog.grupo) return true;
  const e = ejPorId(r.id);
  return !!e && e.grupo === state.prog.grupo;
}
function filasPeriodoGrupo(desde, hasta) {
  return filasEnRango(desde, hasta).filter(enGrupoProg);
}

// Frecuencia de un ejercicio: nº de sesiones que lo incluyen, por periodo.
function frecuenciaEjercicio(id, escala) {
  const pers = listaPeriodos(escala);
  const valores = pers.map(p => new Set(
    state.data.registro.filter(r => r.id === id && r.fecha >= p.desde && r.fecha <= p.hasta).map(r => r.fecha)
  ).size);
  return { labels: pers.map(p => p.label), valores };
}

// Valor de una métrica de resumen (entrenos/series/reps) sobre unas filas.
function metricaValor(filas, metrica) {
  if (metrica === 'entrenos') return new Set(filas.map(f => f.fecha)).size;
  if (metrica === 'series') return new Set(filas.map(f => f.fecha + '|' + f.id + '|' + f.serie)).size;
  return filas.reduce((s, f) => s + (f.reps || 0), 0); // reps
}

// Fracción del periodo en curso ya transcurrida (para la proyección).
function fraccionPeriodo(escala) {
  const hoy = isoADate(hoyISO());
  if (escala === 'semana') return (idxSemana(hoy) + 1) / 7;
  if (escala === 'año') {
    const y = hoy.getFullYear();
    const dias = ((y % 4 === 0 && y % 100 !== 0) || y % 400 === 0) ? 366 : 365;
    return (Math.round((hoy - new Date(y, 0, 1)) / 86400000) + 1) / dias;
  }
  return hoy.getDate() / new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).getDate(); // mes
}

// Lista de periodos (semana/mes/año) desde el primer dato hasta hoy, en orden.
function listaPeriodos(escala) {
  const primera = isoADate(state.data.registro.map(r => r.fecha).sort()[0]);
  const hoy = isoADate(hoyISO());
  const out = [];
  if (escala === 'semana') {
    for (let l = lunesDe(primera), fin = lunesDe(hoy); l <= fin; l = sumarDias(l, 7))
      out.push({ desde: dateAIso(l), hasta: dateAIso(sumarDias(l, 6)), label: 'S' + numeroSemanaISO(l) });
  } else if (escala === 'año') {
    for (let y = primera.getFullYear(); y <= hoy.getFullYear(); y++)
      out.push({ desde: `${y}-01-01`, hasta: `${y}-12-31`, label: String(y) });
  } else {
    let y = primera.getFullYear(), m = primera.getMonth();
    while (y < hoy.getFullYear() || (y === hoy.getFullYear() && m <= hoy.getMonth())) {
      out.push({ desde: dateAIso(new Date(y, m, 1)), hasta: dateAIso(new Date(y, m + 1, 0)), label: MESES[m].slice(0, 3) });
      if (++m > 11) { m = 0; y++; }
    }
  }
  return out;
}

// Serie de barras de una métrica por periodo: valor real, proyección del periodo
// en curso (último) y media de los periodos completos (referencia).
function serieTendencia(metrica, escala) {
  const pers = listaPeriodos(escala);
  const real = pers.map(p => metricaValor(filasPeriodoGrupo(p.desde, p.hasta), metrica));
  const proy = pers.map(() => 0);
  const i = real.length - 1;
  if (i >= 0) {
    const frac = fraccionPeriodo(escala);
    const total = frac > 0 ? Math.round(real[i] / frac) : real[i];
    proy[i] = Math.max(0, total - real[i]);
  }
  const completos = real.slice(0, -1);
  const media = completos.length ? completos.reduce((a, b) => a + b, 0) / completos.length : (real[i] || 0);
  return { labels: pers.map(p => p.label), real, proy, media: red2(media), n: pers.length };
}

// Nº de sesiones y récord absoluto de un ejercicio según la métrica elegida.
function statsEjercicio(id, metrica, escala) {
  const rows = state.data.registro.filter(r => r.id === id);
  const sesiones = new Set(rows.map(r => r.fecha)).size;
  let record = null;
  if (rows.length) {
    if (metrica === 'peso') {
      const b = rows.reduce((m, r) => (r.peso > m.peso ? r : m));
      record = { valor: fmtPeso(b.peso) + ' kg', fecha: b.fecha };
    } else if (metrica === 'tonelaje') {
      const ev = serieEvolucion(id, 'tonelaje', escala);
      let bi = -1; ev.valores.forEach((v, k) => { if (v != null && (bi < 0 || v > ev.valores[bi])) bi = k; });
      const unidad = { semana: 'semana', mes: 'mes', año: 'año' }[escala];
      if (bi >= 0) record = { valor: `${ev.valores[bi]} kg/${unidad} (${ev.labels[bi]})`, fecha: null };
    } else if (metrica === 'frecuencia') {
      const fr = frecuenciaEjercicio(id, escala);
      let bi = 0; fr.valores.forEach((v, k) => { if (v > fr.valores[bi]) bi = k; });
      const unidad = { semana: 'semana', mes: 'mes', año: 'año' }[escala];
      record = { valor: `${fr.valores[bi]}/${unidad} (${fr.labels[bi]})`, fecha: null };
    } else {
      let bv = -1, bf = '';
      rows.forEach(r => { const v = epley(r.peso, r.reps || 0); if (v > bv) { bv = v; bf = r.fecha; } });
      record = { valor: fmtPeso(red2(bv)) + ' kg', fecha: bf };
    }
  }
  return { sesiones, record };
}

// Serie temporal de un ejercicio agregada por periodo (semana/mes/año) según la métrica.
// Peso top y 1RM: el mejor valor del periodo; Tonelaje: la suma del periodo.
// Cubre TODOS los periodos desde el primer dato (como Frecuencia), para que la
// gráfica siempre se pueda desplazar hasta el principio; los periodos sin datos
// del ejercicio van como null (la línea los salta con spanGaps).
function serieEvolucion(id, metrica, escala) {
  const e = ejPorId(id);
  const labels = [], valores = [];
  listaPeriodos(escala).forEach(p => {
    const rs = state.data.registro.filter(r => r.id === id && r.fecha >= p.desde && r.fecha <= p.hasta);
    let v = null;
    if (rs.length) {
      if (metrica === 'peso') v = red2(Math.max(...rs.map(r => r.peso)));
      else if (metrica === 'tonelaje') v = Math.round(rs.reduce((s, r) => s + cargaEfectiva(e, r.peso) * (r.reps || 0), 0));
      else v = red2(Math.max(...rs.map(r => epley(r.peso, r.reps || 0)))); // 1rm
    }
    labels.push(p.label);
    valores.push(v);
  });
  return { labels, valores };
}

// Sesiones de un ejercicio, de la más reciente a la primera, con sus series ordenadas.
function historialEjercicio(id) {
  const porFecha = new Map();
  state.data.registro.filter(r => r.id === id).forEach(r => {
    if (!porFecha.has(r.fecha)) porFecha.set(r.fecha, []);
    porFecha.get(r.fecha).push(r);
  });
  return [...porFecha.keys()].sort().reverse().map(f => ({
    fecha: f,
    series: porFecha.get(f).sort((a, b) => (a.serie - b.serie) || (a.lado || '').localeCompare(b.lado || '')),
  }));
}

// Rachas de constancia, contando por semanas (lunes a domingo) con ≥1 entreno.
// Respeta el filtro de grupo de Progreso (solo semanas con ese grupo entrenado).
function rachas() {
  const semanas = new Set();
  state.data.registro.filter(enGrupoProg)
    .forEach(r => semanas.add(dateAIso(lunesDe(isoADate(r.fecha)))));
  const dif7 = (a, b) => Math.round((isoADate(a) - isoADate(b)) / 86400000) === 7;

  const orden = [...semanas].sort();
  let mejor = 0, run = 0, prev = null;
  orden.forEach(l => { run = (prev && dif7(l, prev)) ? run + 1 : 1; mejor = Math.max(mejor, run); prev = l; });

  // Racha actual: hacia atrás desde esta semana (o la anterior si esta aún no tiene).
  let cursor = lunesDe(isoADate(hoyISO()));
  if (!semanas.has(dateAIso(cursor))) cursor = sumarDias(cursor, -7);
  let actual = 0;
  while (semanas.has(dateAIso(cursor))) { actual++; cursor = sumarDias(cursor, -7); }
  return { actual, mejor };
}

function renderProgreso(v) {
  // Destruir gráficas anteriores para no acumular instancias de Chart.js.
  (state.prog._charts || []).forEach(c => { try { c.destroy(); } catch (e) { /* noop */ } });
  state.prog._charts = [];

  if (!state.data.registro.length) {
    v.innerHTML = '<h2>Progreso</h2><p class="nota">Aún no hay entrenamientos registrados. Empieza en la pestaña Hoy.</p>';
    return;
  }

  const conReg = [...new Set(state.data.registro.map(r => r.id))];
  if (!state.prog.ejercicio || !conReg.includes(state.prog.ejercicio)) state.prog.ejercicio = conReg[0];
  const { periodo, kpi, metrica, grupo } = state.prog;

  // Si el grupo filtrado ya no existe (renombrado/borrado), volver a "todos".
  if (grupo && !gruposExistentes().includes(grupo)) state.prog.grupo = '';

  const grupoOpts = `<option value="" ${!state.prog.grupo ? 'selected' : ''}>Todos los grupos</option>` +
    gruposExistentes().map(g => `<option value="${esc(g)}" ${g === state.prog.grupo ? 'selected' : ''}>${esc(g)}</option>`).join('');

  const perChips = [['semana', 'Semana'], ['mes', 'Mes'], ['año', 'Año']].map(([k, t]) =>
    `<button class="chip-sel ${periodo === k ? 'sel' : ''}" data-periodo="${k}">${t}</button>`).join('');

  // Valores del periodo actual para los tres números (tocables), filtrados por grupo.
  const { desde, hasta } = rangoPeriodo(periodo);
  const fAct = filasPeriodoGrupo(desde, hasta);
  const valKpi = {
    entrenos: metricaValor(fAct, 'entrenos'),
    series: metricaValor(fAct, 'series'),
    reps: metricaValor(fAct, 'reps'),
  };
  const kpiCards = [['entrenos', 'entrenos'], ['series', 'series'], ['reps', 'reps']].map(([k, lbl]) =>
    `<div class="card kpi ${kpi === k ? 'activo' : ''}" data-kpi="${k}">
       <div class="kpi-num">${valKpi[k]}</div><div class="kpi-lbl">${lbl}</div>
     </div>`).join('');

  const tend = serieTendencia(kpi, periodo);
  const escalaLbl = { semana: 'por semana', mes: 'por mes', año: 'por año' }[periodo];
  const kpiLbl = { entrenos: 'Entrenos', series: 'Series', reps: 'Reps' }[kpi];

  const rch = rachas();

  const metChips = [['peso', 'Peso top'], ['1rm', '1RM est.'], ['tonelaje', 'Tonelaje'], ['frecuencia', 'Frecuencia']].map(([k, t]) =>
    `<button class="chip-sel ${metrica === k ? 'sel' : ''}" data-metrica="${k}">${t}</button>`).join('');
  const ejsReg = state.data.ejercicios.filter(e => conReg.includes(e.id));
  const ejActualNom = (ejsReg.find(e => e.id === state.prog.ejercicio) || {}).nombre || '';
  // Nº de puntos de la gráfica de evolución (para el ancho desplazable):
  // todas las métricas cubren todos los periodos desde el primer dato.
  const evN = listaPeriodos(periodo).length;
  const st = statsEjercicio(state.prog.ejercicio, metrica, periodo);
  const recordTxt = st.record
    ? esc(st.record.valor) + (st.record.fecha ? ` (${fmtFecha(st.record.fecha)})` : '')
    : '—';

  v.innerHTML = `
    <h2>Progreso</h2>
    <select id="prog-grupo" class="add-suelto">${grupoOpts}</select>
    <div class="selector-rutina">${perChips}</div>
    <div class="kpis kpis-3">${kpiCards}</div>
    <div class="card">
      <h3>Tendencia · ${kpiLbl} (${escalaLbl})</h3>
      <div class="chart-fy">
        <div class="chart-yaxis"><div class="chart-yaxis-inner"><canvas id="ch-tend-y"></canvas></div></div>
        <div class="chart-scroll"><div class="chart-inner" style="width: max(100%, ${tend.n * 36}px)"><canvas id="ch-tend"></canvas></div></div>
      </div>
      <p class="nota">Barra clara = proyección del periodo en curso · línea = tu media.</p>
    </div>
    <div class="card racha">
      <div class="racha-item"><span class="racha-num">${rch.actual}</span><span class="racha-lbl">semanas seguidas</span></div>
      <div class="racha-item"><span class="racha-num">${rch.mejor}</span><span class="racha-lbl">mejor racha</span></div>
    </div>
    <div class="card">
      <h3>Evolución por ejercicio</h3>
      <div class="combo" id="prog-ej-combo">
        <input id="prog-ej-input" class="buscador combo-input" type="text" autocapitalize="off"
               placeholder="Buscar ejercicio…" value="${esc(ejActualNom)}">
        <div class="combo-lista" id="prog-ej-lista" hidden>
          ${ejsReg.map(e => `<div class="combo-op${e.id === state.prog.ejercicio ? ' sel' : ''}" data-ej="${e.id}" data-nom="${esc(e.nombre.toLowerCase())}">${esc(e.nombre)}</div>`).join('')}
        </div>
      </div>
      <div class="ej-stats"><b>${st.sesiones}</b> sesiones · récord: <b>${recordTxt}</b></div>
      <div class="selector-rutina metrica">${metChips}</div>
      <div class="chart-fy">
        <div class="chart-yaxis"><div class="chart-yaxis-inner"><canvas id="ch-ev-y"></canvas></div></div>
        <div class="chart-scroll"><div class="chart-inner" style="width: max(100%, ${evN * 36}px)"><canvas id="ch-ev"></canvas></div></div>
      </div>
      <h4 class="ev-hist-tit">Todas las sesiones</h4>
      <div class="ev-hist">${historialEjercicio(state.prog.ejercicio).map(s => `
        <div class="ev-hist-dia">
          <span class="ev-hist-fecha">${fmtFecha(s.fecha)}</span>
          <span class="ev-hist-series">${s.series.map(r =>
            `<span class="serie-pill${clasePR(r)}">S${r.serie}${r.lado ? (r.lado === 'Izq' ? '·I' : '·D') : ''}: ${fmtPeso(r.peso)}kg × ${r.reps}</span>`).join('')}
          </span>
        </div>`).join('')}
      </div>
    </div>`;

  const selG = document.getElementById('prog-grupo');
  if (selG) selG.onchange = () => { state.prog.grupo = selG.value; render(); };
  v.querySelectorAll('[data-periodo]').forEach(b => b.onclick = () => { state.prog.periodo = b.dataset.periodo; render(); });
  v.querySelectorAll('[data-kpi]').forEach(b => b.onclick = () => { state.prog.kpi = b.dataset.kpi; render(); });
  v.querySelectorAll('[data-metrica]').forEach(b => b.onclick = () => { state.prog.metrica = b.dataset.metrica; render(); });

  // Buscador de ejercicio (combo propio: funciona en iOS, donde <option hidden> no).
  const inp = document.getElementById('prog-ej-input');
  const lista = document.getElementById('prog-ej-lista');
  if (inp && lista) {
    const ops = [...lista.querySelectorAll('.combo-op')];
    const filtrar = (q) => {
      q = q.trim().toLowerCase();
      ops.forEach(op => { op.style.display = (!q || op.dataset.nom.includes(q)) ? '' : 'none'; });
    };
    inp.onfocus = () => { inp.value = ''; lista.hidden = false; filtrar(''); };
    inp.oninput = () => { lista.hidden = false; filtrar(inp.value); };
    inp.onblur = () => setTimeout(() => { lista.hidden = true; }, 200);
    ops.forEach(op => {
      op.onmousedown = (e) => e.preventDefault();   // evita el blur antes del click
      op.onclick = () => { state.prog.ejercicio = op.dataset.ej; render(); };
    });
  }

  dibujarGraficas();
}

function opcionesGrafica(beginAtZero, stacked) {
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
      x: { stacked, grid: { display: false }, ticks: { color: PAL.lavanda, font: { size: 9 }, maxRotation: 0, autoSkip: true } },
      y: { stacked, grid: { color: PAL.neblina }, ticks: { color: PAL.lavanda, font: { size: 10 } }, beginAtZero },
    },
  };
}

function dibujarGraficas() {
  if (typeof Chart === 'undefined') return; // librería aún no disponible (1ª carga sin red)
  const { periodo, kpi, metrica } = state.prog;

  // Tendencia: barras (real + proyección apilada) + línea de media.
  const t = serieTendencia(kpi, periodo);
  graficaEjeFijo('ch-tend', 'ch-tend-y', () => ({
    data: {
      labels: t.labels,
      datasets: [
        { type: 'bar', data: t.real, backgroundColor: PAL.lavanda, stack: 's', borderRadius: 4 },
        { type: 'bar', data: t.proy, backgroundColor: 'rgba(131,136,186,0.35)', stack: 's', borderRadius: 4 },
        { type: 'line', data: t.labels.map(() => t.media), borderColor: PAL.coral, borderDash: [6, 4], borderWidth: 2, pointRadius: 0, fill: false, stack: 'm' },
      ],
    },
    options: opcionesGrafica(true, true),
  }));

  // Evolución por ejercicio: frecuencia = barras por periodo; el resto, línea por sesión.
  if (metrica === 'frecuencia') {
    const fr = frecuenciaEjercicio(state.prog.ejercicio, periodo);
    graficaEjeFijo('ch-ev', 'ch-ev-y', () => ({
      type: 'bar',
      data: { labels: fr.labels, datasets: [{ data: fr.valores, backgroundColor: PAL.coral, borderRadius: 4 }] },
      options: opcionesGrafica(true, false),
    }));
  } else {
    const ev = serieEvolucion(state.prog.ejercicio, metrica, periodo);
    graficaEjeFijo('ch-ev', 'ch-ev-y', () => ({
      type: 'line',
      data: {
        labels: ev.labels,
        datasets: [{ data: ev.valores, borderColor: PAL.coral, backgroundColor: PAL.coral, tension: 0.25, pointRadius: 3, fill: false, spanGaps: true }],
      },
      options: opcionesGrafica(false, false),
    }));
  }

  // Ambas gráficas arrancan mostrando lo más reciente (a la derecha).
  document.querySelectorAll('.chart-scroll').forEach(sc => { sc.scrollLeft = sc.scrollWidth; });
}

// Dibuja la gráfica desplazable (canvas real) y, superpuesto a la izquierda, un
// clon recortado a su eje Y para que las cantidades no se pierdan al desplazar.
// cfgFactory() devuelve una config nueva en cada llamada (Chart.js muta la config).
function graficaEjeFijo(idReal, idEje, cfgFactory) {
  const cReal = document.getElementById(idReal);
  if (!cReal) return;
  state.prog._charts.push(new Chart(cReal, cfgFactory()));

  const cEje = document.getElementById(idEje);
  if (!cEje) return;
  const cfg = cfgFactory();
  cfg.options.animation = false;
  // Datos invisibles: del clon solo queremos su eje Y (idéntico al real).
  (cfg.data.datasets || []).forEach(d => {
    d.backgroundColor = 'transparent'; d.borderColor = 'transparent';
    d.pointRadius = 0; d.borderWidth = 0;
  });
  const sc = cfg.options.scales;
  sc.x = Object.assign({}, sc.x, { grid: { display: false }, border: { display: false } });
  sc.x.ticks = Object.assign({}, sc.x.ticks, { color: 'transparent' }); // reserva el mismo alto que el real
  sc.y = Object.assign({}, sc.y, { grid: { display: false } });
  const chEje = new Chart(cEje, cfg);
  state.prog._charts.push(chEje);
  // Ajusta el ancho del recuadro fijo al ancho real del eje Y ya renderizado.
  requestAnimationFrame(() => {
    const box = cEje.closest('.chart-yaxis');
    if (box && chEje.scales && chEje.scales.y) box.style.width = (Math.ceil(chEje.scales.y.width) + 2) + 'px';
  });
}

// ===== Editor de rutinas =====
function renderRutinas(v) {
  const rutinas = state.data.rutinas || (state.data.rutinas = []);
  let html = '<h2>Rutinas</h2>';
  html += rutinas.length
    ? rutinas.map((r, ri) => tarjetaRutinaEditor(r, ri)).join('')
    : '<p class="nota">Aún no hay rutinas. Crea la primera con el botón de abajo.</p>';
  html += '<button class="btn" id="nueva-rutina">+ Nueva rutina</button>';
  html += `<p class="nota">Cada rutina es un grupo de ejercicios para elegir en la pestaña Hoy.
    Se guardan en la hoja <b>Rutinas</b> del Excel, así que también puedes editarlas a mano allí.</p>`;
  v.innerHTML = html;

  document.getElementById('nueva-rutina').onclick = nuevaRutina;
  v.querySelectorAll('[data-ren]').forEach(inp =>
    inp.onchange = () => renombrarRutina(Number(inp.dataset.ren), inp.value));
  v.querySelectorAll('[data-del-rut]').forEach(btn =>
    btn.onclick = () => borrarRutina(Number(btn.dataset.delRut)));
  v.querySelectorAll('[data-mover-rut]').forEach(btn => {
    const [ri, dir] = btn.dataset.moverRut.split(':').map(Number);
    btn.onclick = () => moverRutina(ri, dir);
  });
  v.querySelectorAll('[data-mover]').forEach(btn => {
    const [ri, ei, dir] = btn.dataset.mover.split(':').map(Number);
    btn.onclick = () => moverEjercicio(ri, ei, dir);
  });
  v.querySelectorAll('[data-quitar]').forEach(btn => {
    const [ri, ei] = btn.dataset.quitar.split(':').map(Number);
    btn.onclick = () => quitarEjercicio(ri, ei);
  });
  v.querySelectorAll('[data-add-rut]').forEach(sel =>
    sel.onchange = () => anadirEjercicioRutina(Number(sel.dataset.addRut), sel.value));
}

function tarjetaRutinaEditor(r, ri) {
  const ejs = r.ids.map(id => state.data.ejercicios.find(e => e.id === id)).filter(Boolean);
  const filas = ejs.map((e, ei) => `
    <li class="rut-item">
      <span class="rut-nombre">${esc(e.nombre)}</span>
      <span class="rut-acciones">
        <button class="mini" data-mover="${ri}:${ei}:-1" ${ei === 0 ? 'disabled' : ''}>▲</button>
        <button class="mini" data-mover="${ri}:${ei}:1" ${ei === ejs.length - 1 ? 'disabled' : ''}>▼</button>
        <button class="mini mini-x" data-quitar="${ri}:${ei}">✕</button>
      </span>
    </li>`).join('');

  const restantes = state.data.ejercicios.filter(e => !r.ids.includes(e.id));
  const addSel = restantes.length ? `
    <input class="buscador buscar-sel" type="search" placeholder="Buscar ejercicio…" data-target="add-rut-${ri}" autocapitalize="off">
    <select id="add-rut-${ri}" class="add-suelto" data-add-rut="${ri}">
      <option value="">+ Añadir ejercicio…</option>
      ${restantes.map(e => `<option value="${e.id}">${esc(e.nombre)}</option>`).join('')}
    </select>` : '';

  const total = state.data.rutinas.length;
  return `
    <div class="card rutina-edit">
      <div class="rut-cab">
        <button class="mini" data-mover-rut="${ri}:-1" ${ri === 0 ? 'disabled' : ''}>▲</button>
        <button class="mini" data-mover-rut="${ri}:1" ${ri === total - 1 ? 'disabled' : ''}>▼</button>
        <input type="text" class="rut-titulo" data-ren="${ri}" value="${esc(r.nombre)}" autocapitalize="words">
        <button class="mini mini-del" data-del-rut="${ri}" title="Borrar rutina">🗑️</button>
      </div>
      <ul class="rut-lista">${filas || '<li class="nota">Sin ejercicios todavía.</li>'}</ul>
      ${addSel}
    </div>`;
}

function nuevaRutina() {
  const nombre = (prompt('Nombre de la nueva rutina:') || '').trim();
  if (!nombre) return;
  if (state.data.rutinas.some(r => r.nombre.toLowerCase() === nombre.toLowerCase())) {
    alert('Ya existe una rutina con ese nombre.'); return;
  }
  state.data.rutinas.push({ nombre, ids: [] });
  persistirRutinas();
}

function renombrarRutina(ri, nombre) {
  nombre = nombre.trim();
  const r = state.data.rutinas[ri];
  if (!r) return;
  if (!nombre) { alert('El nombre no puede quedar vacío.'); render(); return; }
  if (state.data.rutinas.some((x, i) => i !== ri && x.nombre.toLowerCase() === nombre.toLowerCase())) {
    alert('Ya existe una rutina con ese nombre.'); render(); return;
  }
  if (state.rutinaHoy === r.nombre) state.rutinaHoy = nombre;
  r.nombre = nombre;
  persistirRutinas();
}

function borrarRutina(ri) {
  const r = state.data.rutinas[ri];
  if (!r) return;
  if (!confirmar(`¿Borrar la rutina "${r.nombre}"? Los entrenamientos registrados no se tocan.`)) return;
  state.data.rutinas.splice(ri, 1);
  persistirRutinas();
}

function moverEjercicio(ri, ei, dir) {
  const ids = state.data.rutinas[ri].ids;
  const j = ei + dir;
  if (j < 0 || j >= ids.length) return;
  [ids[ei], ids[j]] = [ids[j], ids[ei]];
  persistirRutinas();
}

// Reordena las rutinas (afecta al orden de los chips en la pestaña Hoy).
function moverRutina(ri, dir) {
  const a = state.data.rutinas, j = ri + dir;
  if (j < 0 || j >= a.length) return;
  [a[ri], a[j]] = [a[j], a[ri]];
  persistirRutinas();
}

function quitarEjercicio(ri, ei) {
  state.data.rutinas[ri].ids.splice(ei, 1);
  persistirRutinas();
}

function anadirEjercicioRutina(ri, id) {
  if (id && !state.data.rutinas[ri].ids.includes(id)) state.data.rutinas[ri].ids.push(id);
  persistirRutinas();
}

function persistirRutinas() {
  saveData();
  marcarPendiente(true);
  render();
}

function renderAjustes(v) {
  const s = state.settings;
  const conectado = !!s.refreshToken;
  v.innerHTML = `
    <h2>Dropbox</h2>
    <div class="card">
      <p>Estado: ${conectado ? '<span class="ok">conectado</span>' : '<span class="aviso">sin conectar</span>'}</p>
      ${conectado && s.lastSync ? `<p class="nota">Última sincronización: <b>${fmtFechaHora(s.lastSync)}</b></p>` : ''}
      <label for="app-key">App Key (consola de desarrolladores de Dropbox)</label>
      <input type="text" id="app-key" value="${s.appKey}" autocomplete="off" autocapitalize="off">
      <label for="ruta">Ruta del Excel en Dropbox</label>
      <input type="text" id="ruta" value="${s.path}" autocomplete="off" autocapitalize="off">
      ${conectado
        ? '<button class="btn" id="btn-sync">Sincronizar ahora</button>'
          + '<button class="btn btn-sec" id="btn-backup">🗄️ Copia de seguridad</button>'
          + '<button class="btn btn-coral" id="btn-desconectar">Desconectar</button>'
        : '<button class="btn" id="btn-conectar">Conectar con Dropbox</button>'}
      ${conectado ? `<p class="nota">La copia de seguridad guarda una copia fechada del Excel
      en la carpeta <b>${CARPETA_BACKUP}</b> del proyecto (en Dropbox), sincronizando antes
      si hay cambios pendientes.</p>` : ''}
      <p class="nota">La app de Dropbox debe ser tuya (gratuita), con permisos
      <b>files.content.read</b> y <b>files.content.write</b>, y esta URL registrada como
      Redirect URI: <b>${redirectUri()}</b>. La App Key no es secreta; aquí no se guarda
      ninguna contraseña.</p>
    </div>

    <h2>Preferencias</h2>
    <div class="card">
      <label class="ej-activo"><input type="checkbox" id="pref-confirmar" ${confirmarBorradosOn() ? 'checked' : ''}> Pedir confirmación al borrar</label>
      <label class="ej-activo"><input type="checkbox" id="pref-deshacer" ${deshacerOn() ? 'checked' : ''}> Permitir deshacer el último cambio</label>
      <label for="pref-polea">Pesos de polea (kg)</label>
      <input type="text" id="pref-polea" value="${esc(pesosPoleaTexto())}" inputmode="decimal" autocomplete="off" autocapitalize="off">
      <p class="nota">Las placas de tu máquina de poleas. En ejercicios de Polea, los botones +/− saltan por estos valores. Decimales con punto, separados por comas. (Siempre puedes editar el peso a mano tocándolo.)</p>
    </div>

    <h2>Datos en local</h2>
    <div class="card">
      <button class="btn btn-sec" id="btn-importar">Importar Excel…</button>
      <input type="file" id="file-import" accept=".xlsx" hidden>
      <button class="btn btn-sec" id="btn-exportar">Exportar Excel</button>
      <p class="nota">Para probar sin Dropbox: importa <b>Forma_Datos.xlsx</b> a mano y
      la app trabajará con la copia local del navegador.</p>
    </div>

    <h2>Aplicación</h2>
    <div class="card">
      <button class="btn btn-sec" id="btn-actualizar">🔄 Buscar actualización</button>
      <p class="nota">Descarga de GitHub la última versión del código y recarga la app.
      No toca tus datos ni la conexión con Dropbox. Úsalo tras publicar cambios.</p>
    </div>`;

  const on = (id, fn) => { const el = document.getElementById(id); if (el) el.onclick = fn; };
  on('btn-conectar', conectarDropbox);
  on('btn-sync', () => { guardarAjustesForm(); sincronizar(); });
  on('btn-backup', copiaSeguridad);
  on('btn-desconectar', () => {
    if (!confirm('¿Desconectar Dropbox? Los datos locales se conservan.')) return;
    state.settings.refreshToken = ''; state.accessToken = null;
    saveSettings(); render();
  });
  on('btn-importar', () => document.getElementById('file-import').click());
  on('btn-exportar', exportarArchivo);
  on('btn-actualizar', forzarActualizacion);
  const pc = document.getElementById('pref-confirmar');
  if (pc) pc.onchange = () => setConfig('confirmar_borrados', pc.checked ? 'SI' : 'NO', 'Pedir confirmación antes de borrar (SI/NO).');
  const pd = document.getElementById('pref-deshacer');
  if (pd) pd.onchange = () => { setConfig('permitir_deshacer', pd.checked ? 'SI' : 'NO', 'Permitir deshacer el último cambio (SI/NO).'); if (!pd.checked) state.undoStack = []; pintarUndo(); };
  const pp = document.getElementById('pref-polea');
  if (pp) pp.onchange = () => setConfig('pesos_polea', pp.value.trim(), 'Placas del stack de polea (kg), decimales con punto, separadas por comas.');
  document.getElementById('file-import').onchange = importarArchivo;
  document.getElementById('ruta').onchange = guardarAjustesForm;
  document.getElementById('app-key').onchange = guardarAjustesForm;
}

// Fuerza traer la última versión del código desde GitHub: borra las cachés del
// service worker, lo desregistra y recarga. NO toca localStorage (datos y token
// de Dropbox a salvo). Si hay cambios sin subir, avisa antes de recargar.
async function forzarActualizacion() {
  if (localStorage.getItem(LS.pending) &&
      !confirm('Tienes cambios sin sincronizar con Dropbox. Se conservan en este dispositivo, pero mejor sincroniza antes. ¿Actualizar la app igualmente?')) {
    return;
  }
  const btn = document.getElementById('btn-actualizar');
  if (btn) { btn.disabled = true; btn.textContent = 'Actualizando…'; }
  try {
    if ('caches' in window) {
      const keys = await caches.keys();
      await Promise.all(keys.map(k => caches.delete(k)));
    }
    if ('serviceWorker' in navigator) {
      const regs = await navigator.serviceWorker.getRegistrations();
      await Promise.all(regs.map(r => r.unregister()));
    }
  } catch (e) {
    console.warn('No se pudo limpiar la caché al actualizar', e);
  }
  // Recarga con parámetro anti-caché para saltarse también la caché HTTP del navegador.
  location.replace(location.pathname + '?v=' + Date.now());
}

function guardarAjustesForm() {
  const k = document.getElementById('app-key');
  const r = document.getElementById('ruta');
  if (k) state.settings.appKey = k.value.trim();
  if (r) state.settings.path = r.value.trim() || RUTA_DEFECTO;
  saveSettings();
}

function pintarBadge() {
  const b = document.getElementById('sync-badge');
  if (!b) return;
  if (localStorage.getItem(LS.pending)) { b.className = 'badge badge-pending'; b.title = 'Cambios sin subir'; }
  else if (state.settings.refreshToken) { b.className = 'badge badge-on'; b.title = 'Sincronizado con Dropbox'; }
  else { b.className = 'badge badge-off'; b.title = 'Sin conexión con Dropbox'; }
}

// ===== Temporizador de descanso (cabecera) =====
// Cuenta atrás que arranca al guardar una serie en Hoy, con la duración del
// "Descanso (min)" del ejercicio. Estado solo en memoria (no toca el Excel).
let timerFin = 0;         // timestamp (ms) en que termina la cuenta atrás
let timerDur = 120;       // última duración usada (seg), fallback si no hay descanso
let timerUltimo = '';     // "id|serie" que arrancó el timer (para no reiniciar con el 2º lado)
let timerInterval = null;
let timerAlarma = false;  // true = llegó a cero y está parpadeando
let audioCtx = null;      // WebAudio; se crea/reanuda en el tap de guardar (gesto, iOS)

function timerFmt(seg) {
  seg = Math.max(0, Math.round(seg));
  return `${Math.floor(seg / 60)}:${String(seg % 60).padStart(2, '0')}`;
}

// El chip (y su panel) solo se muestran en la pestaña Hoy; la cuenta atrás
// sigue corriendo por dentro aunque estés en otra pestaña.
function timerVisible() {
  const chip = document.getElementById('timer-chip');
  const panel = document.getElementById('timer-panel');
  if (!chip) return;
  const enHoy = state.tab === 'hoy';
  chip.hidden = !enHoy;
  if (!enHoy && panel) panel.hidden = true;
}

function timerPintar() {
  // Toda la barra azul parpadea al acabar: se ve aunque el chip esté oculto.
  const head = document.querySelector('header');
  if (head) head.classList.toggle('header-alarma', timerAlarma);
  const chip = document.getElementById('timer-chip');
  if (!chip) return;
  if (timerAlarma) { chip.textContent = '0:00'; chip.className = 'timer-chip timer-alarma'; return; }
  if (timerFin > Date.now()) {
    chip.textContent = timerFmt((timerFin - Date.now()) / 1000);
    chip.className = 'timer-chip timer-on';
  } else {
    chip.textContent = `⏱ ${timerFmt(timerDur)}`;
    chip.className = 'timer-chip';
  }
}

function timerTick() {
  if (timerFin && Date.now() >= timerFin) {   // llegó a cero
    timerFin = 0;
    timerAlarma = true;
    timerBeep();
  }
  timerPintar();
  if (!timerFin && !timerAlarma && timerInterval) { clearInterval(timerInterval); timerInterval = null; }
}

function timerArrancar(seg) {
  if (!(seg > 0)) return;
  timerDur = seg;
  timerFin = Date.now() + seg * 1000;
  timerAlarma = false;
  if (!timerInterval) timerInterval = setInterval(timerTick, 250);
  timerPintar();
}

function timerParar() {
  timerFin = 0;
  timerAlarma = false;
  if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
  timerPintar();
}

// Sonido de aviso. En iOS el audio se bloquea salvo que se "desbloquee" dentro
// de un gesto del usuario, y el contexto se SUSPENDE al apagar pantalla o cambiar
// de app: por eso se desbloquea en cualquier toque y se reanuda al volver.
// Aviso conocido: con el interruptor de silencio del iPhone puesto, iOS silencia
// también el audio web (WebAudio y <audio>); ahí solo queda el aviso visual.
const BEEP_WAV = 'data:audio/wav;base64,UklGRqQbAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YYAbAACAgIGCgX57ent/hIiIhH12c3Z+iI+PiHxxbHB7ipWWjXxsZWl4jJudkn1pXmJ0jaClmH9mV1twjaWsn4JkUVNqjKm0poZiS0tkiq27rotiRUNdh66/s49kRkJag6u+tZNoSEFXgKi+t5dsSkFUfKW9uZtwTEBReKK8u550T0BPdJ67vKJ4UUBMcJu5vaV8VEFKbJe3vqiAV0FIaJO1vquDWkJGZI+zv66HXUNEYYuwv7CLYURDXYeuv7OPZEZCWoOrvrWTaEhBV4CovreXbEpBVHylvbmbcExAUXiivLuedE9AT3Seu7yieFFATHCbub2lfFRBSmyXt76of1dBSGiTtb6rg1pCRmSPs7+uh11DRGGLsL+wi2FEQ12Hrr+zj2RGQlqDq761k2hIQVd/qL63l2xKQVR8pb25m3BMQFF4ory7nnRPQE90nru8onhRQExwm7m9pXxUQUpsl7e+qH9XQUhok7W+q4NaQkZkj7O/roddQ0Rhi7C/sIthRENdh66/s49kRkJag6u+tZNoSEFXgKi+t5dsSkFUfKW9uZtwTEBReKK8u550T0BPdJ67vKJ4UUBMcJu5vaV8VEFKbJe3vqh/V0FIaJO1vquDWkJGZI+zv66HXUNEYYuwv7CLYURDXYeuv7OPZEZCWoOrvrWTaEhBV4CovreXbEpBVHylvbmbcExAUXiivLuedE9AT3Seu7yieFFATHCbub2lfFRBSmyXt76of1dBSGiTtb6rg1pCRmSPs7+uh11DRGGLsL+wi2FEQ12Hrr+zj2RGQlqDq761k2hIQVd/qL63l2xKQVR8pb25m3BMQFF4ory7nnRPQE90nru8onhRQExwm7m9pXxUQUpsl7e+qIBXQUhok7W+q4NaQkZkj7O/roddQ0Rhi7C/sIthRENdh66/s49kRkJag6u+tZNoSEFXgKi+t5dsSkFUfKW9uZtwTEBReKK8u550T0BPdJ67vKJ4UUBMcJu5vaV8VEFKbJe3vqh/V0FIaJO1vquDWkJGZI+zv66HXUNEYYuwv7CLYURDXYeuv7OPZEZCWoOrvrWTaEhBV4CovreXbEpBVHylvbmbcExAUXiivLuedE9AT3Seu7yieFFATHCbub2lfFRBSmyXt76of1dBSGiTtb6rg1pCRmSPs7+uh11DRGGLsL+wi2FEQ12Hrr+zj2RGQlqDq761k2hIQVeAqL63l2xKQVR8pb25m3BMQFF4ory7nnRPQE90nru8onhRQExwm7m9pXxUQUpsl7e+qH9XQUhok7W+q4NaQkZkj7O/roddQ0Rhi7C/sIthRENdh66/s49kRkJag6u+tZNoSEFXgKi+t5dsSkFUfKW9uZtwTEBReKK8u550T0BPdJ67vKJ4UUBMcJu5vaV8VEFKbJe3vqh/V0FIaJO1vquDWkJGZI+zv66HXUNEYYuwv7CLYURDXYeuv7OPZEZCWoOrvrWTaEhBV4CovreXbEpBVHylvbmbcExAUXiivLuedE9AT3Seu7yieFFATHCbub2lfFRBSmyXt76of1dBSGiTtb6rg1pCRmSPs7+uh11DRGGLsL+wi2FEQ12Hrr+zj2RGQlqDq761k2hIQVd/qL63l2xKQVR8pb25m3BMQFF4ory7nnRPQE90nru8onhRQExwm7m9pXxUQUpsl7e+qH9XQUhok7W+q4NaQkZkj7O/roddQ0Rhi7C/sIthRENeh629sY9mSUVdg6i5sJFqTkhcf6O1r5NvU0tcfJ6xrZVzV05ceZmtq5Z3XFJdd5WoqZd6YVZedZGkppd9ZVlfdI2go5eAaV1hc4qcoJaCbWFjcoeYnZWDcGVmcoSUmZOEdGlpc4KQlpGFd21sdIGNko+FeXFvdYCJjoyFe3Rzd3+Gi4qEfXd2eX6Eh4eDfnt6e3+ChIOBf319fn+AgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGCgX57ent/hIiIhH12c3Z+iI+PiHxxbHB7ipWWjXxsZWl4jJudkn1pXmJ0jaClmH9mV1twjaWsn4JkUVNqjKm0poZiS0tkiq27rotiRUNdh66/s49kRkJag6u+tZNoSEFXgKi+t5dsSkFUfKW9uZtwTEBReKK8u550T0BPdJ67vKJ4UUBMcJu5vaV8VEFKbJe3vqiAV0FIaJO1vquDWkJGZI+zv66HXUNEYYuwv7CLYURDXYeuv7OPZEZCWoOrvrWTaEhBV4CovreXbEpBVHylvbmbcExAUXiivLuedE9AT3Seu7yieFFATHCbub2lfFRBSmyXt76of1dBSGiTtb6rg1pCRmSPs7+uh11DRGGLsL+wi2FEQ12Hrr+zj2RGQlqDq761k2hIQVd/qL63l2xKQVR8pb25m3BMQFF4ory7nnRPQE90nru8onhRQExwm7m9pXxUQUpsl7e+qH9XQUhok7W+q4NaQkZkj7O/roddQ0Rhi7C/sIthRENdh66/s49kRkJag6u+tZNoSEFXgKi+t5dsSkFUfKW9uZtwTEBReKK8u550T0BPdJ67vKJ4UUBMcJu5vaV8VEFKbJe3vqh/V0FIaJO1vquDWkJGZI+zv66HXUNEYYuwv7CLYURDXYeuv7OPZEZCWoOrvrWTaEhBV4CovreXbEpBVHylvbmbcExAUXiivLuedE9AT3Seu7yieFFATHCbub2lfFRBSmyXt76of1dBSGiTtb6rg1pCRmSPs7+uh11DRGGLsL+wi2FEQ12Hrr+zj2RGQlqDq761k2hIQVd/qL63l2xKQVR8pb25m3BMQFF4ory7nnRPQE90nru8onhRQExwm7m9pXxUQUpsl7e+qIBXQUhok7W+q4NaQkZkj7O/roddQ0Rhi7C/sIthRENdh66/s49kRkJag6u+tZNoSEFXgKi+t5dsSkFUfKW9uZtwTEBReKK8u550T0BPdJ67vKJ4UUBMcJu5vaV8VEFKbJe3vqh/V0FIaJO1vquDWkJGZI+zv66HXUNEYYuwv7CLYURDXYeuv7OPZEZCWoOrvrWTaEhBV4CovreXbEpBVHylvbmbcExAUXiivLuedE9AT3Seu7yieFFATHCbub2lfFRBSmyXt76of1dBSGiTtb6rg1pCRmSPs7+uh11DRGGLsL+wi2FEQ12Hrr+zj2RGQlqDq761k2hIQVeAqL63l2xKQVR8pb25m3BMQFF4ory7nnRPQE90nru8onhRQExwm7m9pXxUQUpsl7e+qH9XQUhok7W+q4NaQkZkj7O/roddQ0Rhi7C/sIthRENdh66/s49kRkJag6u+tZNoSEFXgKi+t5dsSkFUfKW9uZtwTEBReKK8u550T0BPdJ67vKJ4UUBMcJu5vaV8VEFKbJe3vqh/V0FIaJO1vquDWkJGZI+zv66HXUNEYYuwv7CLYURDXYeuv7OPZEZCWoOrvrWTaEhBV4CovreXbEpBVHylvbmbcExAUXiivLuedE9AT3Seu7yieFFATHCbub2lfFRBSmyXt76of1dBSGiTtb6rg1pCRmSPs7+uh11DRGGLsL+wi2FEQ12Hrr+zj2RGQlqDq761k2hIQVd/qL63l2xKQVR8pb25m3BMQFF4ory7nnRPQE90nru8onhRQExwm7m9pXxUQUpsl7e+qH9XQUhok7W+q4NaQkZkj7O/roddQ0Rhi7C/sIthRENeh629sY9mSUVdg6i5sJFqTkhcf6O1r5NvU0tcfJ6xrZVzV05ceZmtq5Z3XFJdd5WoqZd6YVZedZGkppd9ZVlfdI2go5eAaV1hc4qcoJaCbWFjcoeYnZWDcGVmcoSUmZOEdGlpc4KQlpGFd21sdIGNko+FeXFvdYCJjoyFe3Rzd3+Gi4qEfXd2eX6Eh4eDfnt6e3+ChIOBf319fn+AgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGBf3x7fYGGh4N7dnZ9ho2LgXVvc4COk4x7bWl0iJeYiXNkZnqSn5iBaF1og52llXZcWG6PqKeNaFJYeJ6xpYBZSl2Hrbedb0tIaJm6t49dQU15qb+sfU9AWou1vZ5rRkVqnL22jFtBTnyrv6p6TkFcjre8m2hERmyfvbSKWUBQf62/qHdMQV6RuLuYZkNHb6G+s4dWQFKBr7+mdEpCYZO5upZjQ0lypL6xhFRAVISxvqNySUNjlrq5k2FCSnWmv6+BUkBXh7O+oW9HRGaZu7iQXkFMeKi/rX5QQFmKtb2fbEZFaZu8to1cQU57q7+re05BW422vJxpRUZrnr21i1lAUH6tv6l4TEFekLe8mWdER26hvrOIV0BSgK+/p3VLQmCSubuXZENIcaO+sYVVQFSDsL6kc0lCY5W6uZRhQkp0pb+wglNAVoayvqJwSENlmLu4kV9BS3eov65/UUBYibS9n21GRGibvLeOXEFNeqq/rHxPQFqMtr2dakVFa529tYxaQE99rL+qeU1BXY+3vJpoREdtoL20iVhAUX+uv6d2S0Ffkri7mGVDSHCivrKGVkBTgrC/pXRKQmKUurqVYkJJc6W+sINTQFWFsr6jcUhDZJe7uZJgQkt2p7+ugFFAV4izvqBuR0Rnmry3j11BTXmpv6x9T0Bai7W9nmtGRWqcvbaNW0FOfKu/qnpOQVyOt7ybaERGbJ+9tIpZQFB+rb+od0xBXpG4u5lmQ0dvob6zh1ZAUoGvv6Z1SkJhk7m6lmNDSXKkvrGEVEBUhLG+pHJJQ2OWurmTYUJKdaa/r4FSQFaHs76hb0dEZpm7uJBeQUx4qL+tflBAWYq0vZ9sRkVpm7y2jlxBTnuqv6t7TkFbjba8nGlFRmuevbWLWUBQfa2/qXhMQV2Qt7yZZ0RHbqC+s4hXQFKAr7+ndktCYJK5u5dkQ0hxo76yhVVAVIOwvqRzSUJilbq5lGJCSnSlv7CCU0BWhrK+onBIQ2WYu7iRX0FLd6i/rn9RQFiJtL2fbUZEaJq8t49dQU16qr+sfE9AWoy1vZ1qRUVqnb21jFpAT3ysv6p5TUFdj7e8mmhER22gvbSJWEBRf66/p3dLQV+RuLuYZUNIcKK+soZWQFOCsL+ldEpCYpS6upViQklzpL6wg1RAVYWyvqNxSENkl7u5kmBCS3anv66AUUBXiLO+oG5HRGeavLePXUFNeam/rH1QQFmLtb2ea0ZFapy9to1bQU57q7+qek5BXI62vJtpRUZsn720illAUH6tv6h4TEFekLi7mWZER2+hvrOHVkBSga+/pnVKQmGTubqWY0NJcqS+sYRUQFSEsb6kcklDY5a6uZNhQkp1pr+vgVJAVoezvqFvR0Rmmbu4kF5BTHiov61+UEBZirS9n2xGRWmbvLaOXEFOeqq/q3tOQVuNtr2cakVGa569tYtZQFB9rL+peU1BXY+3vJpnREduoL6ziFdAUYCuv6d2S0Jgkrm7l2RDSHGjvrKFVUBUg7C+pHNJQmKVurqUYkJKdKW/sIJTQFaGsr6icEhDZZi7uJFfQUt3p7+uf1FAWIm0vaBtR0Romry3j11BTXmqv6x8T0BajLW9nWpFRWqdvbWMWkBPfKy/qnpNQV2Pt7yaaERGbZ+9tIlYQFF/rr+od0tBX5G4u5hlQ0hwor6yhlZAU4Kwv6V0SkJilLm6lWJCSXOkvrCDVEBVhbK+o3FIQ2SXu7mSYEJLdqe/r4BSQFeIs76gbkdEZ5m8t5BdQUx4qb+tfVBAWYu1vZ5rRkVpnLy2jVtBTnurv6p7TkFcjra8m2lFRmyfvbSKWUBQfq2/qHhMQV6QuLuZZkRHb6G+s4dWQFKBr7+mdUpCYZO5upZjQ0lypL6xhFRAVISxvqRySUNjlrq5k2FCSnWmv6+BUkBWh7O+oW9HQ2aZu7iRXkFMd6i/rX5QQFmKtL2fbEZEaJu8t45cQU56qr+rfE5BW422vZxqRUZrnr21i1pAT32sv6l5TUFdj7e8mmdER26gvrOIV0BRgK6/p3ZLQmCSubuXZENIcaO+soVVQFODsL6lc0lCYpW6upRiQkp0pb+wglNAVoayvqJwSENlmLu4kl9BS3anv65/UUBYibS9oG1HRGiavLePXUFNeaq/rH1PQFqMtb2da0VFap29toxaQE98rL+qek1BXI63vJtoREZtn720iVhAUX+uv6h3S0Ffkbi7mGVDSHCivrKGVkBTgrC/pXRKQmGUubqVY0JJc6S+sINUQFWFsb6jcUhDZJe7uZJgQkt1p7+vgFJAV4izvqFuR0Rnmby3kF5BTHipv61+UEBZi7W9nmtGRWmcvLaNW0FOe6u/q3tOQVyNtrybaUVGbJ+9tYpZQFB+rb+oeExBXpC4u5lmREdvob6zh1dAUoGvv6Z1SkJhk7m6lmNDSXKjvrGEVEBUhLG+pHJJQ2OWurmTYUJKdKa/r4FSQFaHs76hb0dDZpi7uJFeQUx3qL+tf1BAWYq0vZ9sRkRom7y3jlxBTnqqv6t8TkFbjLa9nGpFRmuevbWLWkBPfay/qXlNQV2Pt7yaZ0RHbqC+s4hXQFGArr+ndktCYJK5u5dkQ0hxo76yhVVAU4OwvqVzSUJilbq6lGJCSnSlv7CCU0BWhrK+onBIQ2WYu7iSX0FLdqe/roBRQFiJtL6gbUdEZ5q8t49dQU15qb+sfU9AWou1vZ1rRUVqnb22jFpBT3ysv6p6TUFcjre8m2hERm2fvbSJWEBRf66/qHdMQV+RuLuYZUNIcKK9sYZXQlWCrbujdU5HZJKzs5NmSlF0n7Wpg1tMXYSnsZt0VVFqkayqjmlTWnibrKCAYFVkhaGolXVcWnCPpKGJbFtie5ejmH5mXmuFm5+PdmNjdY2cmIZvZGp+kpqRfmxncoWUlol4a2x5ipOQg3Rtcn+MkYp+c3B4g4yNhXp0dXyGioiBeXZ5gIaHhH56eX2BhISBfnx9f4GBgYB/fw==';
let beepEl = null;   // <audio> de respaldo por si WebAudio está dormido

function audioDesbloquear() {
  try {
    if (!audioCtx && (window.AudioContext || window.webkitAudioContext))
      audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    if (audioCtx && audioCtx.state === 'suspended') audioCtx.resume();
    if (!beepEl) {
      beepEl = new Audio(BEEP_WAV);
      beepEl.preload = 'auto';
      beepEl.volume = 1;
      // Un play/pause silencioso dentro del gesto deja el elemento "autorizado".
      const v = beepEl.volume; beepEl.volume = 0;
      const p = beepEl.play();
      if (p && p.then) p.then(() => { beepEl.pause(); beepEl.currentTime = 0; beepEl.volume = v; })
                        .catch(() => { beepEl.volume = v; });
      else { beepEl.pause(); beepEl.currentTime = 0; beepEl.volume = v; }
    }
  } catch (e) { /* sin audio disponible */ }
}

// Cualquier toque en la app deja el audio listo; al volver a primer plano se reanuda.
document.addEventListener('pointerdown', audioDesbloquear, { passive: true });
document.addEventListener('touchend', audioDesbloquear, { passive: true });
document.addEventListener('visibilitychange', () => {
  if (!document.hidden && audioCtx && audioCtx.state === 'suspended') audioCtx.resume();
});

function timerBeep() {
  let sonoWebAudio = false;
  try {
    if (audioCtx) {
      if (audioCtx.state === 'suspended') audioCtx.resume();
      if (audioCtx.state === 'running') {
        [0, 0.3, 0.6].forEach((t0, i) => {
          const osc = audioCtx.createOscillator();
          const gain = audioCtx.createGain();
          osc.connect(gain); gain.connect(audioCtx.destination);
          osc.frequency.value = i === 2 ? 1046 : 880;
          const t = audioCtx.currentTime + t0;
          gain.gain.setValueAtTime(0.001, t);
          gain.gain.exponentialRampToValueAtTime(0.6, t + 0.02);
          gain.gain.exponentialRampToValueAtTime(0.001, t + 0.26);
          osc.start(t); osc.stop(t + 0.3);
        });
        sonoWebAudio = true;
      }
    }
  } catch (e) { console.warn('WebAudio no disponible:', e.message); }
  if (!sonoWebAudio && beepEl) {   // respaldo: elemento <audio>
    try { beepEl.currentTime = 0; const p = beepEl.play(); if (p && p.catch) p.catch(() => {}); }
    catch (e) { /* silencio */ }
  }
}

// Llamado desde guardarSerie: arranca el descanso del ejercicio guardado.
// En unilaterales, el segundo lado de la MISMA serie no reinicia la cuenta.
function timerDesdeSerie(e, id, serie, lado) {
  audioDesbloquear();   // el audio debe prepararse dentro de un gesto (iOS)
  const clave = `${id}|${serie}`;
  if (lado && timerUltimo === clave) return;   // 2º lado de la misma serie
  timerUltimo = clave;
  const seg = e && e.descanso > 0 ? Math.round(e.descanso * 60) : timerDur;
  timerArrancar(seg);
}

// Convierte lo escrito a segundos: "1:45" → 105; "90" → 90 (seg); "1,5" → 90 (min).
function timerParse(txt) {
  txt = String(txt).trim().replace(',', '.');
  if (!txt) return 0;
  const m = txt.match(/^(\d+):([0-5]?\d)$/);
  if (m) return Number(m[1]) * 60 + Number(m[2]);
  const n = Number(txt);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return n < 10 ? Math.round(n * 60) : Math.round(n);  // <10 se entiende como minutos
}

function setupTimer() {
  const chip = document.getElementById('timer-chip');
  const panel = document.getElementById('timer-panel');
  if (!chip || !panel) return;
  chip.onclick = () => {
    if (timerAlarma) { timerParar(); return; }   // tocar apaga la alarma
    panel.hidden = !panel.hidden;
    document.getElementById('timer-stop').hidden = !(timerFin > Date.now());
  };
  panel.querySelectorAll('[data-seg]').forEach(b => b.onclick = () => {
    audioDesbloquear();
    timerArrancar(Number(b.dataset.seg));
    panel.hidden = true;
  });
  const custom = document.getElementById('timer-custom');
  custom.onchange = () => {
    const seg = timerParse(custom.value);
    if (seg > 0) { timerArrancar(seg); custom.value = ''; panel.hidden = true; }
  };
  document.getElementById('timer-stop').onclick = () => { timerParar(); panel.hidden = true; };
  const head = document.querySelector('header');
  if (head) head.addEventListener('click', () => { if (timerAlarma) timerParar(); });
  timerPintar();
}
setupTimer();

// ===== Arranque =====
document.querySelectorAll('nav button').forEach(b =>
  b.onclick = () => { state.tab = b.dataset.tab; render(); });

const btnUndo = document.getElementById('btn-undo');
if (btnUndo) btnUndo.onclick = deshacer;

// Pull-to-refresh: tirar hacia abajo estando arriba del todo sincroniza (iOS).
function setupPullToRefresh() {
  const ptr = document.getElementById('ptr');
  if (!ptr) return;
  let startY = 0, pulling = false, dist = 0;
  const TH = 70;
  const pintar = (h, txt) => { ptr.style.height = h + 'px'; ptr.textContent = txt; };
  document.addEventListener('touchstart', (e) => {
    pulling = window.scrollY <= 0 && e.touches.length === 1;
    startY = pulling ? e.touches[0].clientY : 0; dist = 0;
  }, { passive: true });
  document.addEventListener('touchmove', (e) => {
    if (!pulling) return;
    dist = e.touches[0].clientY - startY;
    if (dist > 0 && window.scrollY <= 0) pintar(Math.min(dist, 110), dist > TH ? '↓ Suelta para sincronizar' : '↓ Tira para sincronizar');
    else { dist = 0; pintar(0, ''); }
  }, { passive: true });
  document.addEventListener('touchend', () => {
    if (!pulling) return;
    pulling = false;
    const go = dist > TH; dist = 0;
    if (!go) { pintar(0, ''); return; }
    if (!state.settings.refreshToken) { pintar(0, ''); alert('Conecta Dropbox en Ajustes para sincronizar.'); return; }
    ptr.classList.add('ptr-spin');
    pintar(60, '⟳ Sincronizando…');
    sincronizar().then((ok) => {
      pintar(60, ok ? '✓ Sincronizado' : '✕ Error al sincronizar');
      setTimeout(() => pintar(0, ''), 900);
    }).finally(() => ptr.classList.remove('ptr-spin'));
  }, { passive: true });
}
setupPullToRefresh();

if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');

loadLocal();
render();

const code = new URLSearchParams(location.search).get('code');
if (code) canjearCodigo(code);
else if (state.settings.refreshToken && navigator.onLine) sincronizar(true);
