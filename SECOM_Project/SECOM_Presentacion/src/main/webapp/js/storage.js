import { DEFAULT_INSUMO_CATALOG, DEFAULT_PACKAGE_PRESETS } from './catalogData.js';

function parseJsonSafe(text, fallback = null){
  try{
    return text ? JSON.parse(text) : fallback;
  } catch {
    return fallback;
  }
}

function request(method, url, body){
  const xhr = new XMLHttpRequest();
  xhr.open(method, url, false);
  xhr.setRequestHeader('Accept', 'application/json');

  if (body !== undefined && body !== null){
    xhr.setRequestHeader('Content-Type', 'application/json;charset=UTF-8');
  }

  xhr.send(body !== undefined && body !== null ? JSON.stringify(body) : null);

  const payload = parseJsonSafe(xhr.responseText, { ok:false, message:'Respuesta no válida del servidor.' });

  if (xhr.status >= 200 && xhr.status < 300){
    return payload?.data ?? payload;
  }

  throw new Error(payload?.message || `Error HTTP ${xhr.status}`);
}

const LS = {
  quotes: 'secom_fallback_quotes_v3',
  projects: 'secom_fallback_projects_v3',
  insumos: 'secom_fallback_insumos_v3',
  paquetes: 'secom_fallback_paquetes_v3',
  counters: 'secom_fallback_counters_v3',
};

function readLocal(key, fallback){
  try{
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : structuredCloneSafe(fallback);
  }catch{
    return structuredCloneSafe(fallback);
  }
}

function writeLocal(key, value){
  try{ localStorage.setItem(key, JSON.stringify(value)); }catch{}
  return value;
}

function structuredCloneSafe(value){
  try{ return structuredClone(value); }catch{ return JSON.parse(JSON.stringify(value)); }
}

function normalizeIdPrefix(prefix){
  const counters = readLocal(LS.counters, {});
  const next = Number(counters[prefix] || 0) + 1;
  counters[prefix] = next;
  writeLocal(LS.counters, counters);
  return `${prefix}-${String(next).padStart(4,'0')}`;
}

function todayIso(){
  return new Date().toISOString();
}

function seedInsumos(){
  const current = readLocal(LS.insumos, null);
  if (Array.isArray(current) && current.length) return current;
  const seeded = DEFAULT_INSUMO_CATALOG.map((it, idx) => ({
    id: idx + 1,
    ...structuredCloneSafe(it),
    activo: it.activo !== false,
    createdAt: todayIso(),
    updatedAt: todayIso(),
  }));
  return writeLocal(LS.insumos, seeded);
}

function seedPaquetes(){
  const current = readLocal(LS.paquetes, null);
  if (Array.isArray(current) && current.length) return current;
  const seeded = DEFAULT_PACKAGE_PRESETS.map((p, idx) => ({
    id: idx + 1,
    ...structuredCloneSafe(p),
    key: p.key || `paquete-${idx + 1}`,
    label: p.label || p.nombre || `Paquete ${idx + 1}`,
    nombre: p.nombre || p.label || `Paquete ${idx + 1}`,
    activo: p.activo !== false,
    createdAt: todayIso(),
    updatedAt: todayIso(),
  }));
  return writeLocal(LS.paquetes, seeded);
}

function activeItems(items){
  return (Array.isArray(items) ? items : []).filter(x => x && x.deletedAt == null && x.deleted_at == null);
}

function maybeBackend(method, url, body, fallbackFn){
  try{
    return request(method, url, body);
  }catch(err){
    console.warn('[SECOM] Backend no disponible, usando almacenamiento local:', err?.message || err);
    if (typeof fallbackFn === 'function') return fallbackFn(err);
    throw err;
  }
}

export function getQuotes(){
  return maybeBackend('GET', 'api/quotes', null, () => activeItems(readLocal(LS.quotes, [])).sort((a,b) => String(b.createdAt || b.fecha || '').localeCompare(String(a.createdAt || a.fecha || ''))));
}

export function saveQuote(quote){
  return maybeBackend('POST', 'api/quotes', quote, () => {
    const quotes = activeItems(readLocal(LS.quotes, []));
    const now = todayIso();
    const item = {
      ...structuredCloneSafe(quote || {}),
      id: quote?.id || normalizeIdPrefix('COT'),
      createdAt: quote?.createdAt || now,
      updatedAt: now,
      fecha: quote?.fecha || now,
    };
    quotes.unshift(item);
    writeLocal(LS.quotes, quotes);
    return item;
  });
}

export function updateQuote(id, patch){
  return maybeBackend('PUT', `api/quotes/${encodeURIComponent(id)}`, patch, () => {
    const quotes = activeItems(readLocal(LS.quotes, []));
    const idx = quotes.findIndex(q => String(q.id) === String(id));
    if (idx < 0) throw new Error('No se encontró la cotización local.');
    quotes[idx] = {...quotes[idx], ...structuredCloneSafe(patch || {}), id: quotes[idx].id, updatedAt: todayIso()};
    writeLocal(LS.quotes, quotes);
    return quotes[idx];
  });
}

export function removeQuote(id){
  return maybeBackend('DELETE', `api/quotes/${encodeURIComponent(id)}`, null, () => {
    const quotes = activeItems(readLocal(LS.quotes, [])).filter(q => String(q.id) !== String(id));
    writeLocal(LS.quotes, quotes);
    return {ok:true};
  });
}

export function getInsumos(){
  return maybeBackend('GET', 'api/insumos', null, () => activeItems(seedInsumos()));
}

export function saveInsumo(insumo){
  return maybeBackend('POST', 'api/insumos', insumo, () => {
    const items = activeItems(seedInsumos());
    const code = String(insumo?.codigo || '').trim().toUpperCase();
    if (items.some(x => String(x.codigo || '').toUpperCase() === code)) throw new Error('El código de insumo ya se encuentra registrado.');
    if (Number(insumo?.precio || 0) < 0) throw new Error('El precio unitario debe ser mayor o igual a cero.');
    const now = todayIso();
    const item = {id: nextNumericId(items), ...structuredCloneSafe(insumo || {}), codigo: code, createdAt: now, updatedAt: now};
    items.push(item);
    writeLocal(LS.insumos, items);
    return item;
  });
}

export function updateInsumo(id, patch){
  return maybeBackend('PUT', `api/insumos/${encodeURIComponent(id)}`, patch, () => {
    const items = activeItems(seedInsumos());
    const idx = items.findIndex(x => String(x.id ?? x.codigo) === String(id));
    if (idx < 0) throw new Error('No se encontró el insumo local.');
    if (patch?.precio != null && Number(patch.precio) < 0) throw new Error('El precio unitario debe ser mayor o igual a cero.');
    const code = String(patch?.codigo || items[idx].codigo || '').trim().toUpperCase();
    if (items.some((x,i) => i !== idx && String(x.codigo || '').toUpperCase() === code)) throw new Error('El código de insumo ya se encuentra registrado.');
    items[idx] = {...items[idx], ...structuredCloneSafe(patch || {}), id: items[idx].id, codigo: code, updatedAt: todayIso()};
    writeLocal(LS.insumos, items);
    return items[idx];
  });
}

export function removeInsumo(id){
  return maybeBackend('DELETE', `api/insumos/${encodeURIComponent(id)}`, null, () => {
    const items = activeItems(seedInsumos());
    const target = items.find(x => String(x.id ?? x.codigo) === String(id));
    const next = items.filter(x => String(x.id ?? x.codigo) !== String(id));
    writeLocal(LS.insumos, next);
    if (target){
      const code = String(target.codigo || '').toUpperCase();
      const paquetes = activeItems(seedPaquetes()).map(pkg => ({
        ...pkg,
        items: (pkg.items || pkg.insumos || []).filter(it => String(it.codigo || '').toUpperCase() !== code && String(it.insumoId ?? it.catalogId ?? '') !== String(id))
      }));
      writeLocal(LS.paquetes, paquetes);
    }
    return {ok:true};
  });
}

export function getPaquetes(){
  return maybeBackend('GET', 'api/paquetes', null, () => activeItems(seedPaquetes()));
}

export function savePaquete(paquete){
  return maybeBackend('POST', 'api/paquetes', paquete, () => {
    const paquetes = activeItems(seedPaquetes());
    const now = todayIso();
    const label = paquete?.label || paquete?.nombre || 'Paquete';
    const item = {
      id: nextNumericId(paquetes),
      ...structuredCloneSafe(paquete || {}),
      key: paquete?.key || slugify(label) || `paquete-${Date.now()}`,
      label,
      nombre: paquete?.nombre || label,
      activo: paquete?.activo !== false,
      createdAt: now,
      updatedAt: now,
    };
    paquetes.push(item);
    writeLocal(LS.paquetes, paquetes);
    return item;
  });
}

export function updatePaquete(id, patch){
  return maybeBackend('PUT', `api/paquetes/${encodeURIComponent(id)}`, patch, () => {
    const paquetes = activeItems(seedPaquetes());
    const idx = paquetes.findIndex(x => String(x.id ?? x.key) === String(id));
    if (idx < 0) throw new Error('No se encontró el paquete local.');
    paquetes[idx] = {...paquetes[idx], ...structuredCloneSafe(patch || {}), id: paquetes[idx].id, updatedAt: todayIso()};
    writeLocal(LS.paquetes, paquetes);
    return paquetes[idx];
  });
}

export function removePaquete(id){
  return maybeBackend('DELETE', `api/paquetes/${encodeURIComponent(id)}`, null, () => {
    const paquetes = activeItems(seedPaquetes()).filter(x => String(x.id ?? x.key) !== String(id));
    writeLocal(LS.paquetes, paquetes);
    return {ok:true};
  });
}

export function getProjects(){
  return maybeBackend('GET', 'api/projects', null, () => activeItems(readLocal(LS.projects, [])).sort((a,b) => String(b.createdAt || '').localeCompare(String(a.createdAt || ''))));
}

export function saveProjectFromQuote(quote){
  return maybeBackend('POST', `api/projects/from-quote/${encodeURIComponent(quote.id)}`, {}, () => {
    const projects = activeItems(readLocal(LS.projects, []));
    const now = todayIso();
    const item = {
      ...structuredCloneSafe(quote || {}),
      id: normalizeIdPrefix('PROY'),
      quoteId: quote?.id || '',
      status: 'En planeación',
      createdAt: now,
      updatedAt: now,
    };
    projects.unshift(item);
    writeLocal(LS.projects, projects);
    return item;
  });
}

export function saveProject(project){
  return maybeBackend('POST', 'api/projects', project, () => {
    const projects = activeItems(readLocal(LS.projects, []));
    const now = todayIso();
    const item = {...structuredCloneSafe(project || {}), id: project?.id || normalizeIdPrefix('PROY'), createdAt: now, updatedAt: now};
    projects.unshift(item);
    writeLocal(LS.projects, projects);
    return item;
  });
}

export function updateProject(id, patch){
  return maybeBackend('PUT', `api/projects/${encodeURIComponent(id)}`, patch, () => {
    const projects = activeItems(readLocal(LS.projects, []));
    const idx = projects.findIndex(p => String(p.id) === String(id));
    if (idx < 0) throw new Error('No se encontró el proyecto local.');
    projects[idx] = {...projects[idx], ...structuredCloneSafe(patch || {}), id: projects[idx].id, updatedAt: todayIso()};
    writeLocal(LS.projects, projects);
    return projects[idx];
  });
}

export function removeProject(id){
  return maybeBackend('DELETE', `api/projects/${encodeURIComponent(id)}`, null, () => {
    const projects = activeItems(readLocal(LS.projects, [])).filter(p => String(p.id) !== String(id));
    writeLocal(LS.projects, projects);
    return {ok:true};
  });
}

export function resetAllData(){
  return maybeBackend('POST', 'api/debug/reset', {}, () => {
    [LS.quotes, LS.projects, LS.insumos, LS.paquetes, LS.counters].forEach(k => localStorage.removeItem(k));
    seedInsumos();
    seedPaquetes();
    return {ok:true};
  });
}

export function getCotizacionesReport(filters = {}){
  const params = new URLSearchParams();
  if (filters.fechaInicio) params.set('fechaInicio', filters.fechaInicio);
  if (filters.fechaFin) params.set('fechaFin', filters.fechaFin);
  if (filters.status && filters.status !== 'todos') params.set('status', filters.status);
  if (filters.tarifa && filters.tarifa !== 'todas') params.set('tarifa', filters.tarifa);
  const qs = params.toString();
  return maybeBackend('GET', `api/reports/quotes${qs ? `?${qs}` : ''}`, null, () => buildLocalQuoteReport(filters));
}

function nextNumericId(items){
  return (items || []).reduce((max, it) => Math.max(max, Number(it.id || 0)), 0) + 1;
}

function slugify(value = ''){
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'paquete';
}

function buildLocalQuoteReport(filters = {}){
  const start = filters.fechaInicio || '1900-01-01';
  const end = filters.fechaFin || '2999-12-31';
  const statusFilter = filters.status || 'todos';
  const tarifaFilter = filters.tarifa || 'todas';
  const projects = activeItems(readLocal(LS.projects, []));
  const projectQuoteIds = new Set(projects.map(p => String(p.quoteId || p.id || '')));

  const rows = activeItems(readLocal(LS.quotes, []))
    .map(q => mapQuoteToReportRow(q, projectQuoteIds.has(String(q.id))))
    .filter(row => {
      const date = String(row.fecha || '').slice(0,10) || String(row.createdAt || '').slice(0,10);
      const statusOk = statusFilter === 'todos' || String(row.estatus || '') === String(statusFilter);
      const tariffOk = tarifaFilter === 'todas' || String(row.tarifa || '') === String(tarifaFilter);
      return date >= start && date <= end && statusOk && tariffOk;
    });

  return {
    filters: {fechaInicio:start, fechaFin:end, status:statusFilter, tarifa:tarifaFilter},
    summary: buildReportSummary(rows),
    rows,
  };
}

function mapQuoteToReportRow(q = {}, converted = false){
  const quote = q.quote || q;
  const receipt = q.receipt || {};
  const client = q.client || {};
  const fecha = q.fecha || q.createdAt || q.updatedAt || todayIso();
  return {
    ...structuredCloneSafe(q),
    folio: q.id || '',
    fecha: String(fecha).slice(0,10),
    fechaTexto: String(fecha).slice(0,10),
    cliente: client.nombre || receipt.nombre || '—',
    tarifa: receipt.tarifa || q.selectedTariff?.label || '—',
    consumoMensual: Number(quote.consumoMensual || 0),
    paneles: Number(quote.paneles || 0),
    potenciaKwp: Number(quote.kwp || 0),
    inversion: Number(quote.inversion || 0),
    ahorroMensual: Number(quote.ahorroMensual || 0),
    retornoAnios: Number(quote.retornoAnios || 0),
    estatus: q.status || q.estado || 'Guardada',
    proyectoGenerado: converted || q.proyectoGenerado === true || q.proyecto_generado === true,
  };
}

function buildReportSummary(rows = []){
  const total = rows.length;
  const monto = rows.reduce((a,r) => a + Number(r.inversion || 0), 0);
  const confirmadas = rows.filter(r => String(r.estatus || '').toLowerCase().includes('confirm')).length;
  const convertidas = rows.filter(r => r.proyectoGenerado).length;
  return {
    totalCotizaciones: total,
    montoTotal: Math.round(monto),
    promedioInversion: total ? Math.round(monto / total) : 0,
    confirmadas,
    convertidasProyecto: convertidas,
    pendientes: Math.max(0, total - confirmadas),
    potenciaTotalKwp: Math.round(rows.reduce((a,r) => a + Number(r.potenciaKwp || 0), 0) * 100) / 100,
    ahorroMensualTotal: Math.round(rows.reduce((a,r) => a + Number(r.ahorroMensual || 0), 0)),
  };
}
