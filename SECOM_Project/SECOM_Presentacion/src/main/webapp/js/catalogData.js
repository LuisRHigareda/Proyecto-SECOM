import { formatCurrencyMXN } from './utils.js';

export const DEFAULT_INSUMO_CATALOG = [
  { codigo:'PANEL-550', descripcion:'Panel solar monocristalino 550 W (paneles)', categoria:'Paneles', unidad:'PZA', precio: 3200, impuestoPct: 0.16 },
  { codigo:'PANEL-610', descripcion:'Panel solar monocristalino 610 W (paneles premium)', categoria:'Paneles', unidad:'PZA', precio: 3950, impuestoPct: 0.16 },
  { codigo:'INV-STR', descripcion:'Inversor interconectado string (inversor)', categoria:'Inversores', unidad:'PZA', precio: 18500, impuestoPct: 0.16 },
  { codigo:'INV-HIB', descripcion:'Inversor híbrido con monitoreo (inversor)', categoria:'Inversores', unidad:'PZA', precio: 36500, impuestoPct: 0.16 },
  { codigo:'EST-AL', descripcion:'Estructura de aluminio para azotea o techo (estructura)', categoria:'Estructura', unidad:'SERV', precio: 9800, impuestoPct: 0.16 },
  { codigo:'EST-LA', descripcion:'Estructura reforzada para lámina o teja (estructura)', categoria:'Estructura', unidad:'SERV', precio: 14500, impuestoPct: 0.16 },
  { codigo:'CAB-FV', descripcion:'Cable fotovoltaico y conectores MC4 (cableado)', categoria:'Cableado', unidad:'SERV', precio: 4200, impuestoPct: 0.16 },
  { codigo:'PROT-CC', descripcion:'Protecciones CC/CA y tablero de interconexión (protecciones)', categoria:'Protecciones', unidad:'SERV', precio: 7600, impuestoPct: 0.16 },
  { codigo:'MON-APP', descripcion:'Monitoreo remoto y puesta en marcha (monitoreo)', categoria:'Monitoreo', unidad:'SERV', precio: 3900, impuestoPct: 0.16 },
  { codigo:'TIERRA', descripcion:'Puesta a tierra y canalización (seguridad)', categoria:'Seguridad', unidad:'SERV', precio: 5600, impuestoPct: 0.16 },
  { codigo:'TRAM-CFE', descripcion:'Trámite de interconexión ante CFE (trámite)', categoria:'Trámite', unidad:'SERV', precio: 4800, impuestoPct: 0.16 },
  { codigo:'MO-INST', descripcion:'Mano de obra de instalación (instalación)', categoria:'Instalación', unidad:'SERV', precio: 12600, impuestoPct: 0.16 },
  { codigo:'FLETE', descripcion:'Flete, maniobras y logística (logística)', categoria:'Logística', unidad:'SERV', precio: 3400, impuestoPct: 0.16 },
  { codigo:'BAT-LFP', descripcion:'Banco de baterías LiFePO4 de respaldo (baterías)', categoria:'Baterías', unidad:'PZA', precio: 48500, impuestoPct: 0.16 },
  { codigo:'ING', descripcion:'Ingeniería, planos y memoria técnica (ingeniería)', categoria:'Ingeniería', unidad:'SERV', precio: 5200, impuestoPct: 0.16 },
];

export let INSUMO_CATALOG = DEFAULT_INSUMO_CATALOG.map(normalizeCatalogItem);

export function normalizeCatalogItem(item = {}) {
  return {
    id: item.id ?? item.dbId ?? item.insumoId ?? null,
    codigo: String(item.codigo ?? item.code ?? '').trim().toUpperCase(),
    descripcion: String(item.descripcion ?? item.nombre ?? item.description ?? '').trim(),
    categoria: String(item.categoria ?? 'General').trim() || 'General',
    unidad: String(item.unidad ?? 'UD').trim().toUpperCase() || 'UD',
    precio: Number(item.precio ?? item.precioUnitario ?? item.precio_unitario ?? 0) || 0,
    watts: extractWattsFromText(item),
    impuestoPct: normalizePct(item.impuestoPct ?? item.impuesto_pct ?? 0.16),
    activo: item.activo !== false,
    observaciones: String(item.observaciones ?? '').trim(),
    usoPaquetes: Number(item.usoPaquetes ?? 0) || 0,
  };
}

function normalizePct(value, fallback = 0.16) {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) return fallback;
  return n > 1 ? n / 100 : n;
}

function extractWattsFromText(item = {}) {
  const direct = Number(item.watts ?? item.capacidad ?? item.potencia ?? 0);
  if (Number.isFinite(direct) && direct > 0) return direct;
  const txt = `${item.codigo || ''} ${item.descripcion || item.nombre || ''}`;
  const match = txt.match(/(\d{3,4})\s*w/i);
  return match ? Number(match[1]) : 0;
}

export function setInsumoCatalog(items = []) {
  const normalized = Array.isArray(items)
    ? items.map(normalizeCatalogItem).filter(item => item.codigo && item.descripcion && item.activo !== false)
    : [];
  INSUMO_CATALOG = normalized.length ? normalized : DEFAULT_INSUMO_CATALOG.map(normalizeCatalogItem);
  return INSUMO_CATALOG;
}

export const DEFAULT_PACKAGE_PRESETS = [
  {
    key: 'basico',
    label: 'Paquete básico',
    nombre: 'Paquete básico',
    description: 'Interconexión esencial con componentes estándar y costo contenido.',
    descripcion: 'Interconexión esencial con componentes estándar y costo contenido.',
    badge: 'Recomendado para arranque',
    activo: true,
    items: [
      { codigo: 'PANEL-550', cantidadMode: 'paneles' },
      { codigo: 'INV-STR', cantidad: 1, precioOverride: 16500, precioOverrideAlt: 19800, threshold: 8 },
      { codigo: 'EST-AL', cantidad: 1 },
      { codigo: 'CAB-FV', cantidad: 1 },
      { codigo: 'PROT-CC', cantidad: 1, precioOverride: 6900 },
      { codigo: 'MO-INST', cantidad: 1, precioOverride: 10800 },
      { codigo: 'TRAM-CFE', cantidad: 1, precioOverride: 4200 },
    ],
  },
  {
    key: 'intermedio',
    label: 'Paquete intermedio',
    nombre: 'Paquete intermedio',
    description: 'Mejor balance entre rendimiento, protecciones y monitoreo.',
    descripcion: 'Mejor balance entre rendimiento, protecciones y monitoreo.',
    badge: 'Balanceado',
    activo: true,
    items: [
      { codigo: 'PANEL-550', cantidadMode: 'paneles' },
      { codigo: 'INV-STR', cantidad: 1, precioOverride: 19800, precioOverrideAlt: 23800, threshold: 10 },
      { codigo: 'EST-LA', cantidad: 1, precioOverride: 15800 },
      { codigo: 'CAB-FV', cantidad: 1, precioOverride: 5200 },
      { codigo: 'PROT-CC', cantidad: 1, precioOverride: 8400 },
      { codigo: 'MON-APP', cantidad: 1 },
      { codigo: 'TIERRA', cantidad: 1 },
      { codigo: 'MO-INST', cantidad: 1, precioOverride: 13600 },
      { codigo: 'TRAM-CFE', cantidad: 1 },
      { codigo: 'FLETE', cantidad: 1 },
    ],
  },
  {
    key: 'avanzado',
    label: 'Paquete avanzado',
    nombre: 'Paquete avanzado',
    description: 'Componentes premium, monitoreo extendido y preparación para respaldo.',
    descripcion: 'Componentes premium, monitoreo extendido y preparación para respaldo.',
    badge: 'Premium',
    activo: true,
    items: [
      { codigo: 'PANEL-610', cantidadMode: 'paneles' },
      { codigo: 'INV-HIB', cantidad: 1 },
      { codigo: 'EST-LA', cantidad: 1, precioOverride: 17500 },
      { codigo: 'CAB-FV', cantidad: 1, precioOverride: 6100 },
      { codigo: 'PROT-CC', cantidad: 1, precioOverride: 9800 },
      { codigo: 'MON-APP', cantidad: 1, precioOverride: 5200 },
      { codigo: 'TIERRA', cantidad: 1, precioOverride: 6400 },
      { codigo: 'ING', cantidad: 1 },
      { codigo: 'MO-INST', cantidad: 1, precioOverride: 15800 },
      { codigo: 'TRAM-CFE', cantidad: 1 },
      { codigo: 'FLETE', cantidad: 1, precioOverride: 4200 },
    ],
  },
];

export let PACKAGE_PRESETS = DEFAULT_PACKAGE_PRESETS.map(normalizePackagePreset);

function normalizePackagePreset(pkg = {}) {
  const label = String(pkg.label ?? pkg.nombre ?? 'Paquete').trim() || 'Paquete';
  const description = String(pkg.description ?? pkg.descripcion ?? '').trim();
  return {
    id: pkg.id ?? null,
    key: String(pkg.key ?? pkg.clave ?? pkg.id ?? slugify(label)).trim() || slugify(label),
    label,
    nombre: label,
    description,
    descripcion: description,
    badge: String(pkg.badge ?? pkg.etiqueta ?? 'Paquete').trim() || 'Paquete',
    activo: pkg.activo !== false,
    observaciones: String(pkg.observaciones ?? '').trim(),
    items: Array.isArray(pkg.items ?? pkg.insumos) ? (pkg.items ?? pkg.insumos).map(normalizePackagePresetItem) : [],
    subtotal: Number(pkg.subtotal ?? 0) || 0,
    impuestos: Number(pkg.impuestos ?? 0) || 0,
    total: Number(pkg.total ?? 0) || 0,
    createdAt: pkg.createdAt ?? null,
    updatedAt: pkg.updatedAt ?? null,
  };
}

function normalizePackagePresetItem(item = {}) {
  const codigo = String(item.codigo ?? '').trim().toUpperCase();
  const cantidad = Number(item.cantidad ?? 1) || 1;
  return {
    id: item.id ?? null,
    insumoId: item.insumoId ?? item.insumo_id ?? item.catalogId ?? null,
    catalogId: item.catalogId ?? item.insumoId ?? item.insumo_id ?? null,
    codigo,
    descripcion: String(item.descripcion ?? '').trim(),
    cantidad,
    cantidadMode: item.cantidadMode ?? item.cantidad_mode ?? (codigo.startsWith('PANEL') && cantidad <= 1 ? 'paneles' : null),
    unidad: String(item.unidad ?? '').trim().toUpperCase(),
    precio: Number(item.precio ?? item.precioUnitario ?? 0) || 0,
    precioOverride: item.precioOverride ?? null,
    precioOverrideAlt: item.precioOverrideAlt ?? null,
    threshold: item.threshold ?? null,
    impuestoPct: normalizePct(item.impuestoPct ?? item.impuesto_pct ?? 0.16),
    activo: item.activo !== false,
  };
}

export function setPackageCatalog(items = []) {
  const normalized = Array.isArray(items)
    ? items.map(normalizePackagePreset).filter(pkg => pkg.key && pkg.label && pkg.activo !== false)
    : [];
  PACKAGE_PRESETS = normalized.length ? normalized : DEFAULT_PACKAGE_PRESETS.map(normalizePackagePreset);
  return PACKAGE_PRESETS;
}

function slugify(value = '') {
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'paquete';
}

function estimatePanelCount(context = {}) {
  const paneles = Number(context?.paneles || context?.quote?.paneles || 0);
  if (Number.isFinite(paneles) && paneles > 0) return Math.round(paneles);
  const consumo = Number(context?.consumoMensual || context?.quote?.consumoMensual || context?.receipt?.consumoPeriodo || 0);
  if (consumo <= 0) return 6;
  if (consumo <= 350) return 6;
  if (consumo <= 600) return 8;
  if (consumo <= 900) return 12;
  return 16;
}

function findCatalogItem(ref = {}) {
  const id = ref.insumoId ?? ref.catalogId ?? null;
  const code = String(ref.codigo || '').toUpperCase();
  return INSUMO_CATALOG.find(it => (id != null && String(it.id) === String(id)) || (code && it.codigo === code));
}

function createItemFromCatalog(ref = {}, cantidad = 1, context = {}) {
  const base = findCatalogItem(ref);
  const panelCount = estimatePanelCount(context);
  const qty = ref.cantidadMode === 'paneles' ? panelCount : Number(cantidad || ref.cantidad || 1);
  const override = ref.precioOverrideAlt != null && ref.threshold != null && panelCount > Number(ref.threshold)
    ? Number(ref.precioOverrideAlt)
    : (ref.precioOverride != null ? Number(ref.precioOverride) : null);

  if (!base) {
    return {
      id: ref.id ?? null,
      catalogId: ref.insumoId ?? ref.catalogId ?? null,
      codigo: String(ref.codigo || '').toUpperCase(),
      descripcion: ref.descripcion || String(ref.codigo || 'Insumo no disponible'),
      cantidad: qty > 0 ? qty : 1,
      unidad: ref.unidad || 'SERV',
      precio: Number(override ?? ref.precio ?? 0) || 0,
      categoria: ref.categoria || 'General',
      watts: extractWattsFromText(ref),
      impuestoPct: normalizePct(ref.impuestoPct ?? 0.16),
    };
  }

  return {
    id: ref.id ?? null,
    catalogId: base.id ?? ref.catalogId ?? null,
    codigo: base.codigo,
    descripcion: ref.descripcion || base.descripcion,
    cantidad: qty > 0 ? qty : 1,
    unidad: ref.unidad || base.unidad,
    precio: Number(override ?? base.precio ?? ref.precio ?? 0) || 0,
    categoria: base.categoria || ref.categoria || 'General',
    watts: Number(base.watts || extractWattsFromText(base) || extractWattsFromText(ref) || 0),
    impuestoPct: normalizePct(base.impuestoPct ?? ref.impuestoPct ?? 0.16),
  };
}

export function buildPackageItems(packageKey, context = {}) {
  const preset = PACKAGE_PRESETS.find(p => String(p.key) === String(packageKey));
  if (!preset) return [];
  const sourceItems = Array.isArray(preset.items) ? preset.items : [];
  return sourceItems
    .filter(item => item.activo !== false)
    .map(item => createItemFromCatalog(item, item.cantidad || 1, context));
}

export function getPackageSummaryLabel(packageKey, context = {}) {
  const preset = PACKAGE_PRESETS.find(p => String(p.key) === String(packageKey));
  if (!preset) return 'Sin paquete';
  const items = buildPackageItems(packageKey, context);
  const total = items.reduce((acc, it) => acc + (Number(it.cantidad || 0) * Number(it.precio || 0)), 0);
  return `${preset.label} · ${items.length} insumos · ${formatCurrencyMXN(total)}`;
}
