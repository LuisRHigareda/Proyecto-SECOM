import { formatCurrencyMXN } from './utils.js';

export const DEFAULT_INSUMO_CATALOG = [
  { codigo:'PANEL-550', descripcion:'Panel solar monocristalino 550 W (paneles)', unidad:'PZA', precio: 3200 },
  { codigo:'PANEL-610', descripcion:'Panel solar monocristalino 610 W (paneles premium)', unidad:'PZA', precio: 3950 },
  { codigo:'INV-STR', descripcion:'Inversor interconectado string (inversor)', unidad:'PZA', precio: 18500 },
  { codigo:'INV-HIB', descripcion:'Inversor híbrido con monitoreo (inversor)', unidad:'PZA', precio: 36500 },
  { codigo:'EST-AL', descripcion:'Estructura de aluminio para azotea o techo (estructura)', unidad:'SERV', precio: 9800 },
  { codigo:'EST-LA', descripcion:'Estructura reforzada para lámina o teja (estructura)', unidad:'SERV', precio: 14500 },
  { codigo:'CAB-FV', descripcion:'Cable fotovoltaico y conectores MC4 (cableado)', unidad:'SERV', precio: 4200 },
  { codigo:'PROT-CC', descripcion:'Protecciones CC/CA y tablero de interconexión (protecciones)', unidad:'SERV', precio: 7600 },
  { codigo:'MON-APP', descripcion:'Monitoreo remoto y puesta en marcha (monitoreo)', unidad:'SERV', precio: 3900 },
  { codigo:'TIERRA', descripcion:'Puesta a tierra y canalización (seguridad)', unidad:'SERV', precio: 5600 },
  { codigo:'TRAM-CFE', descripcion:'Trámite de interconexión ante CFE (trámite)', unidad:'SERV', precio: 4800 },
  { codigo:'MO-INST', descripcion:'Mano de obra de instalación (instalación)', unidad:'SERV', precio: 12600 },
  { codigo:'FLETE', descripcion:'Flete, maniobras y logística (logística)', unidad:'SERV', precio: 3400 },
  { codigo:'BAT-LFP', descripcion:'Banco de baterías LiFePO4 de respaldo (baterías)', unidad:'PZA', precio: 48500 },
  { codigo:'ING', descripcion:'Ingeniería, planos y memoria técnica (ingeniería)', unidad:'SERV', precio: 5200 },
];

export let INSUMO_CATALOG = DEFAULT_INSUMO_CATALOG.map(normalizeCatalogItem);

export function normalizeCatalogItem(item = {}) {
  return {
    id: item.id ?? item.dbId ?? null,
    codigo: String(item.codigo ?? item.code ?? '').trim().toUpperCase(),
    descripcion: String(item.descripcion ?? item.nombre ?? item.description ?? '').trim(),
    categoria: String(item.categoria ?? 'General').trim() || 'General',
    unidad: String(item.unidad ?? 'UD').trim().toUpperCase() || 'UD',
    precio: Number(item.precio ?? item.precioUnitario ?? item.precio_unitario ?? 0) || 0,
    impuestoPct: Number(item.impuestoPct ?? item.impuesto_pct ?? 0.16),
    activo: item.activo !== false,
    observaciones: String(item.observaciones ?? '').trim(),
    usoPaquetes: Number(item.usoPaquetes ?? 0) || 0,
  };
}

export function setInsumoCatalog(items = []) {
  const normalized = Array.isArray(items)
    ? items.map(normalizeCatalogItem).filter(item => item.codigo && item.descripcion && item.activo !== false)
    : [];
  INSUMO_CATALOG = normalized.length ? normalized : DEFAULT_INSUMO_CATALOG.map(normalizeCatalogItem);
  return INSUMO_CATALOG;
}

export const PACKAGE_PRESETS = [
  {
    key: 'basico',
    label: 'Paquete básico',
    description: 'Interconexión esencial con componentes estándar y costo contenido.',
    badge: 'Recomendado para arranque',
  },
  {
    key: 'intermedio',
    label: 'Paquete intermedio',
    description: 'Mejor balance entre rendimiento, protecciones y monitoreo.',
    badge: 'Balanceado',
  },
  {
    key: 'avanzado',
    label: 'Paquete avanzado',
    description: 'Componentes premium, monitoreo extendido y preparación para respaldo.',
    badge: 'Premium',
  },
];

function estimatePanelCount(context = {}) {
  const paneles = Number(context?.paneles || context?.quote?.paneles || 0);
  if (Number.isFinite(paneles) && paneles > 0) return Math.round(paneles);
  const consumo = Number(context?.consumoMensual || context?.receipt?.consumoPeriodo || 0);
  if (consumo <= 0) return 6;
  if (consumo <= 350) return 6;
  if (consumo <= 600) return 8;
  if (consumo <= 900) return 12;
  return 16;
}

function panelCodeForPackage(packageKey) {
  return packageKey === 'avanzado' ? 'PANEL-610' : 'PANEL-550';
}

function createItemFromCatalog(code, cantidad = 1, overrides = {}) {
  const base = INSUMO_CATALOG.find(it => it.codigo === code && it.activo !== false);
  if (!base) {
    return { codigo: code, descripcion: code, cantidad, unidad: 'SERV', precio: 0, catalogId: null };
  }
  return {
    catalogId: base.id ?? null,
    codigo: base.codigo,
    descripcion: overrides.descripcion || base.descripcion,
    cantidad,
    unidad: overrides.unidad || base.unidad,
    precio: Number(overrides.precio ?? base.precio),
  };
}

export function buildPackageItems(packageKey, context = {}) {
  const panelCount = estimatePanelCount(context);

  if (packageKey === 'basico') {
    return [
      createItemFromCatalog(panelCodeForPackage(packageKey), panelCount),
      createItemFromCatalog('INV-STR', 1, { precio: panelCount <= 8 ? 16500 : 19800 }),
      createItemFromCatalog('EST-AL', 1),
      createItemFromCatalog('CAB-FV', 1),
      createItemFromCatalog('PROT-CC', 1, { precio: 6900 }),
      createItemFromCatalog('MO-INST', 1, { precio: 10800 }),
      createItemFromCatalog('TRAM-CFE', 1, { precio: 4200 }),
    ];
  }

  if (packageKey === 'intermedio') {
    return [
      createItemFromCatalog(panelCodeForPackage(packageKey), panelCount),
      createItemFromCatalog('INV-STR', 1, { precio: panelCount <= 10 ? 19800 : 23800 }),
      createItemFromCatalog('EST-LA', 1, { precio: 15800 }),
      createItemFromCatalog('CAB-FV', 1, { precio: 5200 }),
      createItemFromCatalog('PROT-CC', 1, { precio: 8400 }),
      createItemFromCatalog('MON-APP', 1),
      createItemFromCatalog('TIERRA', 1),
      createItemFromCatalog('MO-INST', 1, { precio: 13600 }),
      createItemFromCatalog('TRAM-CFE', 1),
      createItemFromCatalog('FLETE', 1),
    ];
  }

  return [
    createItemFromCatalog(panelCodeForPackage('avanzado'), panelCount),
    createItemFromCatalog('INV-HIB', 1),
    createItemFromCatalog('EST-LA', 1, { precio: 17500 }),
    createItemFromCatalog('CAB-FV', 1, { precio: 6100 }),
    createItemFromCatalog('PROT-CC', 1, { precio: 9800 }),
    createItemFromCatalog('MON-APP', 1, { precio: 5200 }),
    createItemFromCatalog('TIERRA', 1, { precio: 6400 }),
    createItemFromCatalog('ING', 1),
    createItemFromCatalog('MO-INST', 1, { precio: 15800 }),
    createItemFromCatalog('TRAM-CFE', 1),
    createItemFromCatalog('FLETE', 1, { precio: 4200 }),
  ];
}

export function getPackageSummaryLabel(packageKey, context = {}) {
  const preset = PACKAGE_PRESETS.find(p => p.key === packageKey);
  if (!preset) return 'Sin paquete';
  const items = buildPackageItems(packageKey, context);
  const total = items.reduce((acc, it) => acc + (Number(it.cantidad || 0) * Number(it.precio || 0)), 0);
  return `${preset.label} · ${items.length} insumos · ${formatCurrencyMXN(total)}`;
}
