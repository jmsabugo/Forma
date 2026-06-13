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
  settings: { appKey: '', refreshToken: '', path: RUTA_DEFECTO },
  accessToken: null,
  tokenExpira: 0,
  tab: 'hoy',
  entrada: {},      // valores en curso de la entrada rápida: { [id]: {peso, reps} }
  rutinaHoy: null,  // rutina elegida para hoy (nombre, o '__libre__', o null)
  extras: [],       // ejercicios sueltos añadidos hoy fuera de la rutina
  editEj: null,     // id del ejercicio en edición en el catálogo ('__nuevo__' al crear)
  histAbierto: null,    // fecha (ISO) del día desplegado en el Historial
  histEdit: null,       // índice en registro de la serie en edición
  histRutinaNueva: {},  // rutina elegida para registrar en un día vacío: {fecha: nombre}
  prog: { periodo: 'mes', kpi: 'entrenos', grupo: '', ejercicio: null, metrica: '1rm', _charts: [] }, // pestaña Progreso
};

// Valores admitidos en el catálogo (coinciden con los desplegables del Excel).
const EQUIPOS = ['Barra', 'Mancuernas', 'Polea', 'Lastre'];
const LATERALIDADES = ['Bilateral', 'Unilateral'];

function loadLocal() {
  try {
    const d = localStorage.getItem(LS.data);
    if (d) state.data = JSON.parse(d);
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

function saveData() {
  localStorage.setItem(LS.data, JSON.stringify(state.data));
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
               'Descanso (min)', 'Series objetivo', 'Reps objetivo', 'Activo', 'Notas'],
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

async function sincronizar() {
  if (!state.settings.refreshToken) { alert('Conecta primero con Dropbox en Ajustes.'); return; }
  const btn = document.getElementById('btn-sync');
  if (btn) { btn.disabled = true; btn.textContent = 'Sincronizando…'; }
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
    render();
  } catch (e) {
    alert(e.message);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = 'Sincronizar ahora'; }
    pintarBadge();
  }
}

// ===== Importar / exportar manual (modo sin Dropbox) =====
function importarArchivo(ev) {
  const f = ev.target.files[0];
  if (!f) return;
  const lector = new FileReader();
  lector.onload = () => {
    try {
      state.data = parseWorkbook(lector.result);
      saveData();
      marcarPendiente(false);
      render();
      alert('Excel importado correctamente.');
    } catch (e) { alert(e.message); }
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

// Series del último entrenamiento de un ejercicio (su fecha más reciente), como
// referencia en la entrada de Hoy. Con `lado` solo mira ese lado.
function ultimaSesionRef(id, lado) {
  let fmax = '';
  for (const r of state.data.registro) {
    if (r.id !== id) continue;
    if (lado && r.lado !== lado) continue;
    if (r.fecha > fmax) fmax = r.fecha;
  }
  if (!fmax) return null;
  const series = state.data.registro
    .filter(r => r.id === id && r.fecha === fmax && (!lado || r.lado === lado))
    .sort((a, b) => a.serie - b.serie);
  return { fecha: fmax, series };
}

// Píldoras con las reps de cada serie del último entrenamiento (a superar). El
// peso y la fecha ya salen en grande arriba; solo se muestra el peso de una serie
// si difiere del de referencia (el de la última serie, el que se ve arriba).
function refUltimaHtml(id, lado) {
  const ref = ultimaSesionRef(id, lado);
  if (!ref || !ref.series.length) return '';
  const pesoRef = red2(ref.series[ref.series.length - 1].peso);
  const pills = ref.series.map(s => {
    const dif = red2(s.peso) !== pesoRef ? ` (${fmtPeso(s.peso)}kg)` : '';
    return `<span class="rep-pill">S${s.serie}:${s.reps}${dif}</span>`;
  }).join('');
  return `<div class="ref-ultima"><span class="ref-lbl">Última vez:</span>${pills}</div>`;
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

// ===== Vistas =====
function render() {
  const v = document.getElementById('view');
  document.querySelectorAll('nav button').forEach(b =>
    b.classList.toggle('active', b.dataset.tab === state.tab));
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

function renderHoy(v) {
  const stepP = incrementoPeso();
  const stepR = incrementoReps();
  const rutinas = state.data.rutinas || [];

  // Si la rutina elegida ya no existe (renombrada/borrada), volver al selector.
  if (state.rutinaHoy && state.rutinaHoy !== '__libre__'
      && !rutinas.some(r => r.nombre === state.rutinaHoy)) {
    state.rutinaHoy = null;
  }

  // Selector de rutina
  const chips = rutinas.map(r =>
    `<button class="chip-sel ${state.rutinaHoy === r.nombre ? 'sel' : ''}" data-rutina="${esc(r.nombre)}">${esc(r.nombre)}</button>`
  ).join('') +
    `<button class="chip-sel ${state.rutinaHoy === '__libre__' ? 'sel' : ''}" data-rutina="__libre__">Libre</button>`;

  let html = `<h2>Hoy · ${fmtFecha(hoyISO())}</h2><div class="selector-rutina">${chips}</div>`;

  if (!state.rutinaHoy) {
    html += '<p class="nota">Elige tu entrenamiento de hoy para empezar a registrar.</p>';
    v.innerHTML = html;
    bindSelectorRutina(v);
    return;
  }

  const lista = ejerciciosDeHoy();
  html += lista.length
    ? lista.map(e => tarjetaEntrada(e, stepP, stepR)).join('')
    : '<p class="nota">Esta rutina no tiene ejercicios activos. Edítala en la pestaña Rutinas.</p>';

  // Añadir un ejercicio suelto que no estaba en la rutina
  const restantes = state.data.ejercicios.filter(e => e.activo && !lista.includes(e));
  if (restantes.length) {
    html += `<select id="add-suelto" class="add-suelto">
      <option value="">+ Añadir otro ejercicio…</option>
      ${restantes.map(e => `<option value="${e.id}">${esc(e.nombre)}</option>`).join('')}
    </select>`;
  }

  v.innerHTML = html;
  bindSelectorRutina(v);

  // Handlers de +/- y guardar (con lado en unilaterales)
  v.querySelectorAll('[data-paso]').forEach(btn => {
    btn.onclick = () => ajustar(btn.dataset.id, btn.dataset.campo, Number(btn.dataset.paso), btn.dataset.lado || '');
  });
  v.querySelectorAll('[data-guardar]').forEach(btn => {
    btn.onclick = () => guardarSerie(btn.dataset.guardar, btn.dataset.guardarLado || '');
  });
  const add = document.getElementById('add-suelto');
  if (add) add.onchange = () => anadirSuelto(add.value);
}

function bindSelectorRutina(v) {
  v.querySelectorAll('[data-rutina]').forEach(btn => {
    btn.onclick = () => elegirRutina(btn.dataset.rutina);
  });
}

function tarjetaEntrada(e, stepP, stepR) {
  const unilateral = String(e.lateralidad).toLowerCase() === 'unilateral';
  if (unilateral) {
    return `
      <div class="card entrada">
        <div class="titulo">
          <span class="nombre">${esc(e.nombre)}</span>
          <span class="chip chip-coral">Unilateral · por lado</span>
        </div>
        <div class="lado-bloque">${bloqueEntrada(e, 'Izq', stepP, stepR)}</div>
        <div class="lado-bloque">${bloqueEntrada(e, 'Der', stepP, stepR)}</div>
      </div>`;
  }
  const u = ultimoPeso(e.id, '');
  return `
    <div class="card entrada">
      <div class="titulo">
        <span class="nombre">${esc(e.nombre)}</span>
        <span class="ultimo-peso">${u ? fmtPeso(u.peso) + ' kg' : '—'}
          <small>${u ? 'último (' + fmtFecha(u.fecha) + ')' : 'sin registros'}</small>
        </span>
      </div>
      ${bloqueEntrada(e, '', stepP, stepR)}
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
      ${hoy.map(s => `<span class="serie-pill">S${s.serie}: ${fmtPeso(s.peso)}kg × ${s.reps}</span>`).join('')}
    </div>` : '';

  return `
    ${cabLado}
    ${refUltimaHtml(e.id, lado)}
    <div class="control">
      <button class="paso" data-id="${e.id}" data-lado="${lado}" data-campo="peso" data-paso="${-stepP}">−</button>
      <div class="lectura"><span class="num">${fmtPeso(base.peso)}</span><span class="ud">kg</span></div>
      <button class="paso" data-id="${e.id}" data-lado="${lado}" data-campo="peso" data-paso="${stepP}">+</button>
    </div>
    <div class="control">
      <button class="paso" data-id="${e.id}" data-lado="${lado}" data-campo="reps" data-paso="${-stepR}">−</button>
      <div class="lectura"><span class="num">${base.reps}</span><span class="ud">reps${lado ? '/lado' : ''}</span></div>
      <button class="paso" data-id="${e.id}" data-lado="${lado}" data-campo="reps" data-paso="${stepR}">+</button>
    </div>
    <button class="btn" data-guardar="${e.id}" data-guardar-lado="${lado}">${guardarTxt}</button>
    ${seriesHoyHtml}`;
}

function ajustar(id, campo, paso, lado) {
  const key = entradaKey(id, lado);
  const v = state.entrada[key] || (state.entrada[key] = { peso: 0, reps: 0 });
  if (campo === 'peso') v.peso = Math.max(0, red2(v.peso + paso));
  else v.reps = Math.max(0, v.reps + paso);
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
  saveData();
  marcarPendiente(true);
  render();
}

function renderEjercicios(v) {
  // Catálogo completo (activos primero, luego inactivos); cada uno editable.
  const ejs = state.data.ejercicios.slice().sort((a, b) =>
    (a.activo === b.activo) ? 0 : (a.activo ? -1 : 1));

  let html = '<h2>Catálogo</h2>';
  html += ejs.map(e =>
    state.editEj === e.id ? formEjercicio(e) : tarjetaEjercicio(e)).join('');

  html += state.editEj === '__nuevo__'
    ? formEjercicio(null)
    : '<button class="btn" id="nuevo-ej">+ Nuevo ejercicio</button>';

  html += `<p class="nota">El catálogo se guarda en la hoja <b>Ejercicios</b> del Excel.
    Al borrar un ejercicio, las series ya registradas se conservan en el historial.</p>`;
  v.innerHTML = html;
  bindEjercicios(v);
}

function tarjetaEjercicio(e) {
  const u = ultimoPeso(e.id);
  const unilateral = String(e.lateralidad).toLowerCase() === 'unilateral';
  return `
    <div class="card${e.activo ? '' : ' inactivo'}">
      <div class="titulo">
        <span class="nombre">${esc(e.nombre)}</span>
        <span class="ultimo-peso">${u ? fmtPeso(u.peso) + ' kg' : '—'}
          <small>${u ? 'último (' + fmtFecha(u.fecha) + ')' : 'sin registros'}</small>
        </span>
      </div>
      <div class="chips">
        <span class="chip">${esc(e.grupo)}</span>
        <span class="chip">${esc(e.equipamiento)}</span>
        ${unilateral ? '<span class="chip chip-coral">Unilateral · reps por lado</span>' : ''}
        <span class="chip">${e.seriesObj}×${e.repsObj}</span>
        <span class="chip">descanso ${e.descanso}′</span>
        ${e.activo ? '' : '<span class="chip chip-coral">Inactivo</span>'}
      </div>
      <button class="btn btn-sec" data-edit="${e.id}">✏️ Editar</button>
    </div>`;
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
  if (!confirm(msg)) return;

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
    dias.forEach(dt => { html += filaDiaHist(dt, porFecha); });
  }

  v.innerHTML = html;
  bindHistorial(v);
}

function filaDiaHist(dt, porFecha) {
  const iso = dateAIso(dt);
  const esHoy = iso === hoyISO();
  const dow = DIAS_SEM[idxSemana(dt)];
  const idxs = porFecha.get(iso);

  if (!idxs) {
    if (state.histAbierto === iso) return diaVacioAbierto(dt);
    return `<div class="hist-dia vacio" data-abrir="${iso}">
      <div class="hist-fecha"><span class="dow">${dow}</span><span class="dnum">${dt.getDate()}</span></div>
      <div class="hist-resumen">${esHoy ? 'hoy · ' : ''}descanso</div>
    </div>`;
  }

  const filas = idxs.map(i => state.data.registro[i]);
  const rutina = rutinaDelDia(iso) || 'Entreno Libre';
  const nEj = new Set(filas.map(f => f.id)).size;
  const nSeries = new Set(filas.map(f => f.id + '#' + f.serie)).size;
  const nReps = filas.reduce((s, f) => s + (f.reps || 0), 0);
  const abierto = state.histAbierto === iso;

  let html = `<div class="hist-dia${esHoy ? ' es-hoy' : ''}${abierto ? ' abierto' : ''}" data-toggle="${iso}">
    <div class="hist-fecha"><span class="dow">${dow}</span><span class="dnum">${dt.getDate()}</span></div>
    <div class="hist-info">
      <div class="hist-rutina">${esc(rutina)}${esHoy ? ' <span class="hoy-badge">hoy</span>' : ''}</div>
      <div class="hist-resumen">${nEj} ejercicios · ${nSeries} series · ${nReps} reps</div>
    </div>
    <span class="hist-chev">${abierto ? '▴' : '▸'}</span>
  </div>`;
  if (abierto) html += diaExpandido(iso, idxs);
  return html;
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
  orden.forEach(id => { html += bloqueEjercicioHist(iso, id, porEj.get(id)); });

  const presentes = new Set(orden);
  const restantes = state.data.ejercicios.filter(e => !presentes.has(e.id));
  if (restantes.length) {
    html += `<select class="add-suelto" data-add-ej-dia="${iso}">
      <option value="">+ Añadir ejercicio…</option>
      ${restantes.map(e => `<option value="${e.id}">${esc(e.nombre)}</option>`).join('')}
    </select>`;
  }
  html += `<button class="btn btn-coral" data-del-dia="${iso}">Borrar día</button></div>`;
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
    <div class="ej-hist-nombre">${esc(nombre)}</div>
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
  return `<span class="serie-pill editable" data-edit-serie="${i}">S${r.serie}: ${fmtPeso(r.peso)}kg × ${r.reps}</span>`;
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
  return `<span class="lado-pill editable" data-edit-serie="${i}">${lado.toLowerCase()} ${fmtPeso(r.peso)}kg × ${r.reps}</span>`;
}

function diaVacioAbierto(dt) {
  const iso = dateAIso(dt);
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
  return `<div class="hist-dia vacio abierto" data-toggle="${iso}">
      <div class="hist-fecha"><span class="dow">${dow}</span><span class="dnum">${dt.getDate()}</span></div>
      <div class="hist-resumen">nuevo entreno</div>
      <span class="hist-chev">▴</span>
    </div><div class="dia-detalle">${panel}</div>`;
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
  v.querySelectorAll('[data-add-serie]').forEach(el => el.onclick = (ev) => {
    ev.stopPropagation();
    const [iso, id] = el.dataset.addSerie.split('|'); anadirSerieDia(iso, id);
  });
  v.querySelectorAll('[data-add-ej-dia]').forEach(el =>
    el.onchange = () => anadirSerieDia(el.dataset.addEjDia, el.value));
  v.querySelectorAll('[data-del-dia]').forEach(el => el.onclick = (ev) => {
    ev.stopPropagation(); borrarDiaHist(el.dataset.delDia);
  });
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
  state.data.registro = state.data.registro.filter(r => !(r.fecha === iso && r.id === id && r.serie === serie));
  renumerarSeries(iso, id);
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

function borrarDiaHist(iso) {
  if (!confirm(`¿Borrar todo el entreno del ${fmtFecha(iso)}? Se eliminarán todas sus series.`)) return;
  state.data.registro = state.data.registro.filter(r => r.fecha !== iso);
  delete state.histRutinaNueva[iso];
  state.histAbierto = null;
  state.histEdit = null;
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
  return (e && String(e.equipamiento).toLowerCase() === 'mancuernas') ? peso * 2 : peso;
}

// 1RM estimado (Epley): peso × (1 + reps/30).
function epley(peso, reps) { return peso * (1 + reps / 30); }

function ejPorId(id) { return state.data.ejercicios.find(e => e.id === id); }

function fmtFechaCorta(iso) { const [, m, d] = iso.split('-'); return `${+d}/${+m}`; }

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
      const ev = serieEvolucion(id, 'tonelaje');
      let bi = 0; ev.valores.forEach((v, k) => { if (v > ev.valores[bi]) bi = k; });
      record = { valor: ev.valores[bi] + ' kg', fecha: ev.fechas[bi] };
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

// Serie temporal de un ejercicio (un punto por sesión) según la métrica.
function serieEvolucion(id, metrica) {
  const porFecha = new Map();
  state.data.registro.forEach(r => {
    if (r.id !== id) return;
    if (!porFecha.has(r.fecha)) porFecha.set(r.fecha, []);
    porFecha.get(r.fecha).push(r);
  });
  const e = ejPorId(id);
  const fechas = [...porFecha.keys()].sort();
  const valores = fechas.map(f => {
    const rs = porFecha.get(f);
    if (metrica === 'peso') return red2(Math.max(...rs.map(r => r.peso)));
    if (metrica === 'tonelaje') return Math.round(rs.reduce((s, r) => s + cargaEfectiva(e, r.peso) * (r.reps || 0), 0));
    return red2(Math.max(...rs.map(r => epley(r.peso, r.reps || 0)))); // 1rm
  });
  return { fechas, valores };
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
  const ejOpts = state.data.ejercicios.filter(e => conReg.includes(e.id))
    .map(e => `<option value="${e.id}" ${e.id === state.prog.ejercicio ? 'selected' : ''}>${esc(e.nombre)}</option>`).join('');
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
      <div class="chart-scroll"><div class="chart-inner" style="width: max(100%, ${tend.n * 36}px)"><canvas id="ch-tend"></canvas></div></div>
      <p class="nota">Barra clara = proyección del periodo en curso · línea = tu media.</p>
    </div>
    <div class="card racha">
      <div class="racha-item"><span class="racha-num">${rch.actual}</span><span class="racha-lbl">semanas seguidas</span></div>
      <div class="racha-item"><span class="racha-num">${rch.mejor}</span><span class="racha-lbl">mejor racha</span></div>
    </div>
    <div class="card">
      <h3>Evolución por ejercicio</h3>
      <select id="prog-ej" class="add-suelto">${ejOpts}</select>
      <div class="ej-stats"><b>${st.sesiones}</b> sesiones · récord: <b>${recordTxt}</b></div>
      <div class="selector-rutina metrica">${metChips}</div>
      <div class="chart-box"><canvas id="ch-ev"></canvas></div>
    </div>`;

  const selG = document.getElementById('prog-grupo');
  if (selG) selG.onchange = () => { state.prog.grupo = selG.value; render(); };
  v.querySelectorAll('[data-periodo]').forEach(b => b.onclick = () => { state.prog.periodo = b.dataset.periodo; render(); });
  v.querySelectorAll('[data-kpi]').forEach(b => b.onclick = () => { state.prog.kpi = b.dataset.kpi; render(); });
  v.querySelectorAll('[data-metrica]').forEach(b => b.onclick = () => { state.prog.metrica = b.dataset.metrica; render(); });
  const sel = document.getElementById('prog-ej');
  if (sel) sel.onchange = () => { state.prog.ejercicio = sel.value; render(); };

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
  const ct = document.getElementById('ch-tend');
  if (ct) state.prog._charts.push(new Chart(ct, {
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
  const ce = document.getElementById('ch-ev');
  if (ce) {
    if (metrica === 'frecuencia') {
      const fr = frecuenciaEjercicio(state.prog.ejercicio, periodo);
      state.prog._charts.push(new Chart(ce, {
        type: 'bar',
        data: { labels: fr.labels, datasets: [{ data: fr.valores, backgroundColor: PAL.coral, borderRadius: 4 }] },
        options: opcionesGrafica(true, false),
      }));
    } else {
      const ev = serieEvolucion(state.prog.ejercicio, metrica);
      state.prog._charts.push(new Chart(ce, {
        type: 'line',
        data: {
          labels: ev.fechas.map(fmtFechaCorta),
          datasets: [{ data: ev.valores, borderColor: PAL.coral, backgroundColor: PAL.coral, tension: 0.25, pointRadius: 3, fill: false }],
        },
        options: opcionesGrafica(false, false),
      }));
    }
  }

  // La tendencia arranca mostrando lo más reciente (a la derecha).
  const sc = document.querySelector('.chart-scroll');
  if (sc) sc.scrollLeft = sc.scrollWidth;
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
    <select class="add-suelto" data-add-rut="${ri}">
      <option value="">+ Añadir ejercicio…</option>
      ${restantes.map(e => `<option value="${e.id}">${esc(e.nombre)}</option>`).join('')}
    </select>` : '';

  return `
    <div class="card rutina-edit">
      <div class="rut-cab">
        <input type="text" class="rut-titulo" data-ren="${ri}" value="${esc(r.nombre)}" autocapitalize="words">
        <button class="mini mini-x" data-del-rut="${ri}" title="Borrar rutina">🗑️</button>
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
  if (!confirm(`¿Borrar la rutina "${r.nombre}"? Los entrenamientos registrados no se tocan.`)) return;
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
      <label for="app-key">App Key (consola de desarrolladores de Dropbox)</label>
      <input type="text" id="app-key" value="${s.appKey}" autocomplete="off" autocapitalize="off">
      <label for="ruta">Ruta del Excel en Dropbox</label>
      <input type="text" id="ruta" value="${s.path}" autocomplete="off" autocapitalize="off">
      ${conectado
        ? '<button class="btn" id="btn-sync">Sincronizar ahora</button>'
          + '<button class="btn btn-coral" id="btn-desconectar">Desconectar</button>'
        : '<button class="btn" id="btn-conectar">Conectar con Dropbox</button>'}
      <p class="nota">La app de Dropbox debe ser tuya (gratuita), con permisos
      <b>files.content.read</b> y <b>files.content.write</b>, y esta URL registrada como
      Redirect URI: <b>${redirectUri()}</b>. La App Key no es secreta; aquí no se guarda
      ninguna contraseña.</p>
    </div>

    <h2>Datos en local</h2>
    <div class="card">
      <button class="btn btn-sec" id="btn-importar">Importar Excel…</button>
      <input type="file" id="file-import" accept=".xlsx" hidden>
      <button class="btn btn-sec" id="btn-exportar">Exportar Excel</button>
      <p class="nota">Para probar sin Dropbox: importa <b>Forma_Datos.xlsx</b> a mano y
      la app trabajará con la copia local del navegador.</p>
    </div>`;

  const on = (id, fn) => { const el = document.getElementById(id); if (el) el.onclick = fn; };
  on('btn-conectar', conectarDropbox);
  on('btn-sync', () => { guardarAjustesForm(); sincronizar(); });
  on('btn-desconectar', () => {
    if (!confirm('¿Desconectar Dropbox? Los datos locales se conservan.')) return;
    state.settings.refreshToken = ''; state.accessToken = null;
    saveSettings(); render();
  });
  on('btn-importar', () => document.getElementById('file-import').click());
  on('btn-exportar', exportarArchivo);
  document.getElementById('file-import').onchange = importarArchivo;
  document.getElementById('ruta').onchange = guardarAjustesForm;
  document.getElementById('app-key').onchange = guardarAjustesForm;
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

// ===== Arranque =====
document.querySelectorAll('nav button').forEach(b =>
  b.onclick = () => { state.tab = b.dataset.tab; render(); });

if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');

loadLocal();
render();

const code = new URLSearchParams(location.search).get('code');
if (code) canjearCodigo(code);
else if (state.settings.refreshToken && navigator.onLine) sincronizar();
