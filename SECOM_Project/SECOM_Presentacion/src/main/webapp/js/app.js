import { $, $$, debounce, formatCurrencyMXN, formatDate, formatDateTime, formatNumber, openModal, closeModal, setPillStatus, toast } from './utils.js';
import { analyzeReceiptFile, createEmptyReceiptData } from './receiptParser.js';
import { computeQuote, buildExportHtml } from './quoteEngine.js';
import { getQuotes, getProjects, getInsumos, getPaquetes, getCotizacionesReport, resetAllData, saveInsumo, savePaquete, saveProjectFromQuote, saveProject, saveQuote, updateInsumo, updatePaquete, updateProject, updateQuote, removeInsumo, removePaquete, removeProject } from './storage.js';
import { INSUMO_CATALOG, PACKAGE_PRESETS, buildPackageItems, getPackageSummaryLabel, setInsumoCatalog, setPackageCatalog } from './catalogData.js';
import { applyTheme, getDefaultQuoteParams, loadUserPreferences, saveUserPreferences } from './preferences.js';

const state = {
    route: 'cotizador',
    selectedTariff: null,
    wizardStep: 1,
    receiptFile: null,
    receipt: null, // parsed
    receiptCanvas: null,
    client: {
        nombre: '',
        telefono: '',
        email: '',
        direccion: '',
    },
    // Información editable del sistema (no necesariamente derivada del recibo)
    quoteMeta: {
        panelModelo: 'Panel fotovoltaico',
        panelDimensiones: '',
        inversorModelo: '',
        tipoTecho: 'No especificado',
        sombras: 'No especificado',
        perdidasSombraPct: 0,
        notasFisicas: '',
    },
    // Permite sobrescribir resultados calculados (p. ej. paneles manuales)
    overrides: {
        paneles: null,
        consumoMensual: null,
    },
    params: getDefaultQuoteParams(),
    quote: null,
    savedQuote: null,
    chart: null,
    exportChart: null,
    selectedPackage: '',
    preferences: loadUserPreferences(),
    cashvolt: {
        file: null,
        data: null,
        status: 'Listo',
    },
    insumos: {
        items: [],
        search: '',
    },
    paquetes: {
        items: [],
        search: '',
    },
    reportes: {
        fechaInicio: '',
        fechaFin: '',
        status: 'todos',
        tarifa: 'todas',
        data: null,
    }
};

const TARIFFS = [
    {key: 'dom_m', label: 'Doméstica Mensual', periodo: 'Mensual', familia: 'Doméstica'},
    {key: 'dom_b', label: 'Doméstica Bimestral', periodo: 'Bimestral', familia: 'Doméstica'},
    {key: 'pdbt_m', label: 'PDBT Mensual', periodo: 'Mensual', familia: 'PDBT'},
    {key: 'pdbt_b', label: 'PDBT Bimestral', periodo: 'Bimestral', familia: 'PDBT'},
    {key: 'gdmth', label: 'GDMTH', periodo: 'Mensual', familia: 'GDMTH'},
    {key: 'gdmto', label: 'GDMTO', periodo: 'Mensual', familia: 'GDMTO'},
    {key: 'cashvolt', label: 'Subir Datos a CashVolt', kind: 'cashvolt'},
];

const TARIFF_CALC_RULES = {
    dom_m: {
        ahorroPct: 0.65,
        alcance: 'Recibo mensual doméstico. El consumo del recibo se usa completo como kWh/mes.',
        diferencia: 'Base de cálculo simple: consumo mensual igual al consumo capturado y ahorro estimado con 65% del pago mensual.'
    },
    dom_b: {
        ahorroPct: 0.65,
        alcance: 'Recibo bimestral doméstico. El consumo y el pago del recibo se convierten a equivalente mensual.',
        diferencia: 'Reduce la base mensual a la mitad: kWh/mes = consumo ÷ 2 y pago mensual = total ÷ 2.'
    },
    pdbt_m: {
        ahorroPct: 0.70,
        alcance: 'PDBT mensual. Se usa consumo mensual completo y una expectativa de ahorro comercial baja tensión.',
        diferencia: 'Mantiene kWh/mes completos, pero usa 70% del pago mensual como ahorro estimado para la precotización.'
    },
    pdbt_b: {
        ahorroPct: 0.70,
        alcance: 'PDBT bimestral. Convierte consumo y pago bimestral a equivalente mensual.',
        diferencia: 'Calcula kWh/mes = consumo ÷ 2 y usa 70% del pago mensual equivalente como ahorro estimado.'
    },
    gdmth: {
        ahorroPct: 0.75,
        alcance: 'GDMTH. Se conserva el consumo mensual y se usa una referencia de ahorro mayor por tarifa horaria.',
        diferencia: 'No modela demanda ni horarios; para mostrar impacto en esta versión usa 75% del pago mensual como ahorro estimado.'
    },
    gdmto: {
        ahorroPct: 0.72,
        alcance: 'GDMTO. Se conserva el consumo mensual y se usa una referencia intermedia de ahorro comercial.',
        diferencia: 'No modela demanda contratada; para mostrar impacto en esta versión usa 72% del pago mensual como ahorro estimado.'
    }
};

function getTariffRule(t) {
    return TARIFF_CALC_RULES[t?.key] || {
        ahorroPct: 0.65,
        alcance: 'Tarifa mensual por defecto.',
        diferencia: 'Usa el consumo capturado como kWh/mes y 65% del pago mensual como ahorro estimado.'
    };
}

function getTariffMonthlyConsumption(t, consumoPeriodo) {
    const consumo = Number(consumoPeriodo || 0);
    return t?.periodo === 'Bimestral' ? consumo / 2 : consumo;
}

function getTariffMonthlyBill(t, totalAPagar) {
    const total = Number(totalAPagar || 0);
    return t?.periodo === 'Bimestral' ? total / 2 : total;
}

function getTariffEstimatedSaving(t, totalAPagar) {
    const rule = getTariffRule(t);
    const pagoMensual = getTariffMonthlyBill(t, totalAPagar);
    return Math.round(Math.max(0, pagoMensual * rule.ahorroPct));
}

function applyTariffCalculationAssumptions(receipt, t) {
    if (!receipt || !t)
        return receipt;
    receipt.tipoPeriodo = t.periodo || receipt.tipoPeriodo || 'Mensual';
    receipt.periodo = receipt.periodo || {raw: '', start: null, end: null, days: 0};
    receipt.periodo.days = receipt.tipoPeriodo === 'Bimestral' ? 60 : 30;
    receipt.tarifaSeleccionada = t.label || receipt.tarifaSeleccionada || '';
    receipt.tarifaCalculo = {
        key: t.key,
        label: t.label,
        periodo: t.periodo || 'Mensual',
        familia: t.familia || 'Tarifa',
        ahorroPct: getTariffRule(t).ahorroPct,
        consumoMensualBase: Math.round(getTariffMonthlyConsumption(t, receipt.consumoPeriodo || 0)),
        pagoMensualBase: Math.round(getTariffMonthlyBill(t, receipt.totalAPagar || 0)),
    };
    if (Number(receipt.totalAPagar || 0) > 0) {
        receipt.ahorroEstimado = getTariffEstimatedSaving(t, receipt.totalAPagar || 0);
    }
    return receipt;
}

function getTariffImpact(t, quote = null, receipt = state.receipt) {
    const rule = getTariffRule(t);
    const isBimestral = t?.periodo === 'Bimestral';
    const consumoPeriodo = Number(receipt?.consumoPeriodo || 0);
    const totalAPagar = Number(receipt?.totalAPagar || 0);
    const consumoMensual = getTariffMonthlyConsumption(t, consumoPeriodo);
    const pagoMensual = getTariffMonthlyBill(t, totalAPagar);
    const ahorroEstimado = getTariffEstimatedSaving(t, totalAPagar);
    const formula = isBimestral
            ? 'Consumo mensual = consumo del recibo ÷ 2'
            : 'Consumo mensual = consumo del recibo';
    const short = isBimestral
            ? 'divide consumo y pago del recibo entre 2 para convertirlos a equivalente mensual'
            : 'usa consumo y pago del recibo directamente como valores mensuales';
    const family = t?.familia ? `Familia ${t.familia}` : 'Tarifa';
    const quoteLine = quote
            ? ` Con los datos actuales: ${formatNumber(quote.consumoMensual || 0)} kWh/mes, ${Number(quote.kwp || 0).toFixed(2)} kWp, ${formatNumber(quote.paneles || 0)} paneles, ahorro ${formatCurrencyMXN(quote.ahorroMensual || 0)}.`
            : '';

    return {
        formula,
        short,
        ahorroPct: rule.ahorroPct,
        message: `${t?.label || 'Tarifa'}: ${short}. Ahorro usado en precotización: ${Math.round(rule.ahorroPct * 100)}% del pago mensual.`,
        html: `
      <b>Impacto en cálculo:</b> ${escapeHtml(formula)}.<br/>
      <span>${escapeHtml(family)} · Periodo ${escapeHtml(t?.periodo || '—')} · Ahorro de referencia ${Math.round(rule.ahorroPct * 100)}%.</span><br/>
      <span>${escapeHtml(rule.diferencia)}${escapeHtml(quoteLine)}</span>
      ${consumoPeriodo || totalAPagar ? `<br/><span><b>Con los datos capturados:</b> ${formatNumber(consumoPeriodo)} kWh del recibo → ${formatNumber(consumoMensual)} kWh/mes; ${formatCurrencyMXN(totalAPagar)} del recibo → ${formatCurrencyMXN(pagoMensual)}/mes; ahorro estimado ${formatCurrencyMXN(ahorroEstimado)}.</span>` : ''}
    `
    };
}

function buildTariffTooltip(t) {
    const rule = getTariffRule(t);
    const periodo = t?.periodo || 'Mensual';
    const consumo = periodo === 'Bimestral' ? 'consumo ÷ 2' : 'consumo completo';
    const pago = periodo === 'Bimestral' ? 'pago ÷ 2' : 'pago completo';
    return `${t?.label || 'Tarifa'}\nPeriodo: ${periodo}\nConsumo mensual usado: ${consumo}\nPago mensual usado: ${pago}\nAhorro referencia: ${Math.round(rule.ahorroPct * 100)}%\n${rule.diferencia}`;
}

function renderTariffImpactBox(t, quote = null, id = '') {
    const info = getTariffImpact(t, quote);
    return `<div ${id ? `id="${id}"` : ''} class="review-ok" style="margin-top:12px">${info.html}</div>`;
}

function renderTariffHoverNote(t) {
    const info = getTariffImpact(t, null);
    return `
    <b>${escapeHtml(t?.label || 'Selecciona una tarifa')}</b><br/>
    <span>${escapeHtml(info.short)}.</span><br/>
    <span>Ahorro de referencia para precotización: <b>${Math.round(info.ahorroPct * 100)}%</b> del pago mensual equivalente.</span>
  `;
}

function buildReceiptForTariffScenario(t) {
    const base = structuredCloneSafe(state.receipt || createEmptyReceiptData(t));
    applyTariffCalculationAssumptions(base, t);
    return base;
}

function renderTariffComparisonTable(selected = state.selectedTariff) {
    const rows = TARIFFS.filter(t => t.kind !== 'cashvolt').map(t => {
        const scenarioReceipt = buildReceiptForTariffScenario(t);
        const scenarioQuote = computeQuote(scenarioReceipt, state.client || {}, state.params || {}, state.overrides || {});
        const isSelected = selected?.key === t.key;
        const rule = getTariffRule(t);
        return `
      <tr style="${isSelected ? 'font-weight:900' : ''}">
        <td>${escapeHtml(t.label)}${isSelected ? ' ✓' : ''}</td>
        <td>${escapeHtml(t.periodo || '—')}</td>
        <td>${t.periodo === 'Bimestral' ? 'kWh ÷ 2' : 'kWh directo'}</td>
        <td>${Math.round(rule.ahorroPct * 100)}%</td>
        <td>${formatNumber(scenarioQuote.consumoMensual || 0)} kWh/mes</td>
        <td>${formatNumber(scenarioQuote.paneles || 0)}</td>
        <td>${formatCurrencyMXN(scenarioQuote.ahorroMensual || 0)}</td>
        <td>${scenarioQuote.retornoAnios ? `${scenarioQuote.retornoAnios} años` : '—'}</td>
      </tr>
    `;
    }).join('');

    return `
    <div class="review-panel" id="tariffComparisonBox" style="margin-top:12px; overflow:auto">
      <div class="card__subtitle" style="margin-bottom:8px">Comparación rápida por tarifa</div>
      <table class="table table--tight" style="min-width:760px">
        <thead>
          <tr>
            <th>Tarifa</th>
            <th>Periodo</th>
            <th>Consumo mensual</th>
            <th>Ahorro ref.</th>
            <th>kWh/mes usados</th>
            <th>Paneles</th>
            <th>Ahorro mensual</th>
            <th>Retorno</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="help" style="margin-top:8px">Nota: GDMTH/GDMTO se muestran como aproximación de precotización; no se modelan demanda ni horarios en esta versión.</div>
    </div>
  `;
}

init();

function init() {
    // Theme y preferencias
    state.preferences = loadUserPreferences();
    applyTheme(state.preferences.theme || 'dark');
    loadInsumoCatalogSafe(true);
    loadPackageCatalogSafe(true);

    // Icons
    if (window.lucide)
        window.lucide.createIcons();
    else
        window.addEventListener('load', () => window.lucide?.createIcons(), {once: true});

    // Sidebar
    const sidebar = $('#sidebar');
    $('#btnSidebarToggle')?.addEventListener('click', () => {
        sidebar.classList.toggle('is-collapsed');
    });

    $('#btnOpenSidebar')?.addEventListener('click', () => sidebar.classList.add('is-open'));
    // Close sidebar when navigating on mobile
    window.addEventListener('resize', () => {
        if (window.innerWidth > 860)
            sidebar.classList.remove('is-open');
    });

    // Routing
    $$('.nav__item[data-route]').forEach(btn => {
        btn.addEventListener('click', () => {
            const r = btn.dataset.route;
            if (r === 'cotizador') {
                startNewQuoteFlow();
            } else if (r === 'cashvolt') {
                window.open('https://cashvolt.mx/public/login', '_blank');
            } else {
                setRoute(r);
            }
            sidebar.classList.remove('is-open');
        });
    });

    // Modal close
    $('#modal')?.addEventListener('click', (e) => {
        const t = e.target;
        if (t?.dataset?.close === 'true')
            closeModal();
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape')
            closeModal();
    });

    renderAllRoutes();
    // Pantalla principal: lista de cotizaciones
    setRoute('historial');
}

function setRoute(route) {
    state.route = route;

    $$('.nav__item').forEach(el => el.classList.toggle('is-active', el.dataset.route === route));
    $$('.route').forEach(el => el.classList.toggle('is-active', el.id === `route-${route}`));

    const titleMap = {
        cotizador: {t: 'Crear cotización', c: 'Cotizaciones / Crear'},
        cashvolt: {t: 'CashVolt', c: 'CashVolt / Carga'},
        proyectos: {t: 'Proyectos', c: 'Proyectos / Gestión'},
        paquetes: {t: 'Paquetes', c: 'Catálogos / Paquetes'},
        insumos: {t: 'Insumos', c: 'Catálogos / Insumos'},
        reportes: {t: 'Reportes', c: 'Reportes / Cotizaciones'},
        historial: {t: 'Lista de cotizaciones', c: 'Cotizaciones / Lista'},
        opciones: {t: 'Preferencias', c: 'Sistema / Preferencias'},
    };
    $('#pageTitle').textContent = titleMap[route]?.t || 'SECOM';
    $('#pageCrumb').textContent = titleMap[route]?.c || '';

    if (route === 'paquetes') {
        renderPaquetesRoute();
    }
    if (route === 'insumos') {
        renderInsumosRoute();
    }
    if (route === 'reportes') {
        renderReportesRoute();
    }
}

function startNewQuoteFlow() {
    state.selectedTariff = null;
    state.overrides = {paneles: null, consumoMensual: null};
    state.selectedPackage = '';
    state.preferences = loadUserPreferences();
    state.params = structuredCloneSafe(getDefaultQuoteParams());
    state.quoteMeta = {
        panelModelo: 'Panel fotovoltaico',
        panelDimensiones: '',
        inversorModelo: '',
        tipoTecho: 'No especificado',
        sombras: 'No especificado',
        perdidasSombraPct: 0,
        notasFisicas: '',
    };
    resetWizard(true);
    renderCotizadorRoute();
    setRoute('cotizador');
}

function renderAllRoutes() {
    renderCotizadorRoute();
    renderCashvoltRoute();
    renderHistorialRoute();
    renderProyectosRoute();
    renderPaquetesRoute();
    renderInsumosRoute();
    renderReportesRoute();
    renderOpcionesRoute();
}

// ------------------------------
// Cotizador (Wizard)
// ------------------------------

function renderCotizadorRoute() {
    const root = $('#route-cotizador');
    // Pantalla 1 (equivalente al menú de Tkinter)
    if (!state.selectedTariff) {
        root.innerHTML = renderTariffSelector();
        wireTariffSelector();
        window.lucide?.createIcons();
        return;
    }

    root.innerHTML = `
    <div class="row" style="justify-content:space-between; align-items:center; gap:12px">
      <div class="pill">
        <span class="pill__dot"></span>
        <span>Tarifa seleccionada:</span>&nbsp;<b>${escapeHtml(state.selectedTariff.label)}</b>
      </div>
      <button class="btn" id="btnChangeTarifa"><i data-lucide="repeat-2"></i>Cambiar</button>
    </div>

    ${renderTariffImpactBox(state.selectedTariff)}

    <div class="card card--flat" style="box-shadow:none; margin-top:12px">
      <div class="stepper" id="stepper"></div>
    </div>

    <div class="grid cols-2" style="margin-top:14px">
      <div class="card" id="wizardLeft"></div>
      <div class="card" id="wizardRight"></div>
    </div>
  `;

    $('#btnChangeTarifa')?.addEventListener('click', () => {
        state.selectedTariff = null;
        resetWizard(true);
        renderCotizadorRoute();
    });

    buildStepper();
    renderWizard();
}

function renderTariffSelector() {
    const buttons = TARIFFS.filter(t => t.kind !== 'cashvolt').map(t => {
        return `
      <button class="btn btn--big" data-tariff="${t.key}" title="${escapeAttr(buildTariffTooltip(t))}">
        <div style="display:flex; align-items:center; gap:10px">
          <span class="badge">${escapeHtml(t.periodo || '—')}</span>
          <div style="text-align:left">
            <div style="font-weight:900">${escapeHtml(t.label)}</div>
            <div style="color:var(--muted); font-size:12px">Selecciona para cargar el recibo</div>
            <div style="color:var(--text); font-size:12px; margin-top:4px; line-height:1.35">
              Impacto: ${escapeHtml(getTariffImpact(t).short)} · ahorro ref. ${Math.round(getTariffRule(t).ahorroPct * 100)}%.
            </div>
          </div>
        </div>
      </button>
    `;
    }).join('');

    return `
    <div class="card">
      <div class="card__title">Nueva cotización</div>
      <div class="card__subtitle">Selecciona el tipo de tarifa para iniciar el proceso.</div>
      <div class="review-ok" style="margin-top:12px">
        <b>Qué cambia al elegir tarifa:</b> el sistema fija si el recibo se tratará como mensual o bimestral y también aplica un porcentaje de ahorro de referencia para la precotización. Esto cambia el consumo mensual usado, la potencia requerida, paneles, producción, ahorro mensual y retorno estimado.
      </div>

      <div class="grid" style="grid-template-columns:repeat(2, minmax(0, 1fr)); gap:12px; margin-top:12px">
        ${buttons}
      </div>

      <div id="tariffHoverNote" class="review-panel" style="margin-top:12px">
        ${renderTariffHoverNote(TARIFFS.find(t => t.key === 'dom_m'))}
      </div>

      <div class="row" style="justify-content:flex-start; margin-top:14px">
        <button class="btn btn--success" id="btnGoCashvolt"><i data-lucide="cloud-upload"></i>Subir Datos a CashVolt</button>
      </div>
    </div>
  `;
}

function wireTariffSelector() {
    $('#btnGoCashvolt')?.addEventListener('click', () => setRoute('cashvolt'));
    const note = $('#tariffHoverNote');
    $$('#route-cotizador [data-tariff]').forEach(btn => {
        const showNote = () => {
            const t = TARIFFS.find(x => x.key === btn.dataset.tariff);
            if (t && note)
                note.innerHTML = renderTariffHoverNote(t);
        };
        btn.addEventListener('mouseenter', showNote);
        btn.addEventListener('focus', showNote);
        btn.addEventListener('click', () => {
            const key = btn.dataset.tariff;
            const t = TARIFFS.find(x => x.key === key);
            if (!t)
                return;

            state.selectedTariff = t;
            resetWizard(true);
            renderCotizadorRoute();
            toast({title: 'Tarifa seleccionada', message: getTariffImpact(t).message, icon: 'calculator'});

            // Abrir selector de archivo inmediatamente (flujo equivalente a “abrir PDF”)
            setTimeout(() => $('#fileInput')?.click(), 50);
        });
    });
}

// ------------------------------
// CashVolt
// ------------------------------

function renderCashvoltRoute() {
    const root = $('#route-cashvolt');
    if (!root)
        return;

    root.innerHTML = `
    <div class="card">
      <div class="card__title">CashVolt</div>
      <div class="card__subtitle">Acceso directo</div>
      <div class="help" style="margin-top:10px">Esta sección únicamente abre la plataforma oficial de CashVolt en una pestaña nueva.</div>
      <div class="wizard-actions" style="margin-top:14px">
        <button class="btn btn--primary" id="cvOpenOnly"><i data-lucide="external-link"></i>Ir a CashVolt</button>
      </div>
    </div>
  `;
    $('#cvOpenOnly')?.addEventListener('click', () => window.open('https://cashvolt.mx/public/login', '_blank'));
}

function wireCashvolt() {
    const fileInput = $('#cvFile');
    const dz = $('#cvDrop');

    const pick = () => fileInput.click();
    dz?.addEventListener('click', pick);
    dz?.addEventListener('keydown', (e) => (e.key === 'Enter' || e.key === ' ') && pick());

    dz?.addEventListener('dragover', (e) => {
        e.preventDefault();
        dz.classList.add('is-dragover');
    });
    dz?.addEventListener('dragleave', () => dz.classList.remove('is-dragover'));
    dz?.addEventListener('drop', (e) => {
        e.preventDefault();
        dz.classList.remove('is-dragover');
        const f = e.dataTransfer.files?.[0];
        if (f)
            setCashvoltFile(f);
    });

    fileInput?.addEventListener('change', () => {
        const f = fileInput.files?.[0];
        if (f)
            setCashvoltFile(f);
    });

    $('#cvUseLast')?.addEventListener('click', () => {
        const quotes = getQuotes();
        const last = quotes[0];
        if (!last) {
            toast({title: 'Sin cotizaciones', message: 'No hay cotizaciones guardadas para usar.', icon: 'alert-triangle'});
            return;
        }
        state.cashvolt.data = buildCashvoltDataFromQuote(last);
        $('#cvMsg').textContent = `Datos preparados desde ${last.id}.`;
        renderCashvoltTable();
    });

    $('#cvLoad')?.addEventListener('click', () => {
        // Si hay archivo, mostramos datos preparados (sin lectura profunda)
        if (state.cashvolt.file) {
            state.cashvolt.data = buildCashvoltDataFromQuote({
                client: {nombre: (state.client?.nombre || 'Cliente')},
                quote: state.quote || {},
                receipt: state.receipt || {},
            });
            $('#cvMsg').textContent = 'Datos preparados correctamente.';
            renderCashvoltTable();
            toast({title: 'CashVolt', message: 'Datos listos para capturar.', icon: 'check-circle'});
            return;
        }

        // Sin archivo: intenta usar la cotización actual
        if (state.quote && state.client?.nombre) {
            state.cashvolt.data = buildCashvoltDataFromQuote({client: state.client, quote: state.quote, receipt: state.receipt});
            $('#cvMsg').textContent = 'Datos preparados desde la cotización actual.';
            renderCashvoltTable();
            return;
        }
        toast({title: 'Falta información', message: 'Selecciona un archivo .xlsm o usa una cotización guardada.', icon: 'alert-triangle'});
    });

    $('#cvOpen')?.addEventListener('click', () => window.open('https://cashvolt.mx/public/login', '_blank'));
    $('#cvCopy')?.addEventListener('click', async () => {
        const rows = state.cashvolt.data || [];
        if (!rows.length) {
            toast({title: 'Nada para copiar', message: 'Primero prepara los datos.', icon: 'alert-triangle'});
            return;
        }
        const text = rows.map(r => `${r.label} ${r.value}`).join('\n');
        try {
            await navigator.clipboard.writeText(text);
            toast({title: 'Copiado', message: 'Se copió al portapapeles.', icon: 'copy'});
        } catch {
            toast({title: 'No se pudo copiar', message: 'Tu navegador bloqueó el portapapeles.', icon: 'x-circle'});
        }
    });

    renderCashvoltTable();
}

function setCashvoltFile(file) {
    state.cashvolt.file = file;
    $('#cvFileHint').textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
    $('#cvMsg').textContent = '';
}

function buildCashvoltDataFromQuote(q) {
    const quote = q.quote || q;
    const paneles = Number(quote.paneles || q.quote?.paneles || 0);
    const panelW = Number(q.quote?.params?.panelWatts || quote.params?.panelWatts || 550);
    const ahorro = Number(q.quote?.ahorroMensual || quote.ahorroMensual || 0);
    const costo = Number(q.quote?.inversion || quote.inversion || 0);
    const advisor = state.preferences?.company?.advisorName || 'Asesor SECOM';
    return [
        {label: 'Asesor de ventas:', value: advisor},
        {label: 'Nombre del cliente:', value: (q.client?.nombre || q.receipt?.nombre || '—')},
        {label: 'Ahorro mensual del proyecto:', value: String(Math.round(ahorro))},
        {label: 'Cantidad de Paneles:', value: String(paneles || '—')},
        {label: 'Capacidad del Panel:', value: String(panelW || '—')},
        {label: 'Costo del proyecto:', value: String(Math.round(costo))},
    ];
}

function renderCashvoltTable() {
    const tbody = $('#cvTable tbody');
    if (!tbody)
        return;
    const rows = state.cashvolt.data || [];
    tbody.innerHTML = rows.length ? rows.map((r, i) => `
    <tr>
      <td>${escapeHtml(r.label)}</td>
      <td>${escapeHtml(String(r.value))}</td>
      <td style="text-align:right">
        <button class="btn" data-cv-copy="${i}"><i data-lucide="copy"></i>Copiar</button>
      </td>
    </tr>
  `).join('') : `
    <tr><td colspan="3" style="color:var(--muted)">Aún no hay datos preparados.</td></tr>
  `;

    window.lucide?.createIcons();
    tbody.querySelectorAll('[data-cv-copy]').forEach(btn => {
        btn.addEventListener('click', async () => {
            const idx = Number(btn.dataset.cvCopy);
            const row = rows[idx];
            if (!row)
                return;
            try {
                await navigator.clipboard.writeText(String(row.value));
                toast({title: 'Copiado', message: row.label, icon: 'copy'});
            } catch {
                toast({title: 'No se pudo copiar', message: 'Tu navegador bloqueó el portapapeles.', icon: 'x-circle'});
            }
        });
    });
}

function buildStepper() {
    const el = $('#stepper');
    const steps = [
        {n: 1, label: 'Recibo'},
        {n: 2, label: 'Datos y sistema'},
        {n: 3, label: 'Resumen'},
        {n: 4, label: 'Exportar'},
    ];

    el.innerHTML = steps.map((s, i) => {
        const cls = s.n === state.wizardStep ? 'is-active' : (s.n < state.wizardStep ? 'is-done' : '');
        const line = i < steps.length - 1 ? '<div class="stepper__line"></div>' : '';
        return `
      <div class="stepper__item ${cls}">
        <div class="stepper__dot">${s.n}</div>
        <div class="stepper__label">${s.label}</div>
      </div>
      ${line}
    `;
    }).join('');
    window.lucide?.createIcons();
}

function gotoStep(n) {
    state.wizardStep = n;
    buildStepper();
    renderWizard();
}

function renderWizard() {
    const left = $('#wizardLeft');
    const right = $('#wizardRight');

    if (state.wizardStep === 1) {
        left.innerHTML = renderStep1Left();
        right.innerHTML = renderStep1Right();
        wireStep1();
    } else if (state.wizardStep === 2) {
        left.innerHTML = renderStep2Left();
        right.innerHTML = renderStep2Right();
        wireStep2();
    } else if (state.wizardStep === 3) {
        left.innerHTML = renderStep3Left();
        right.innerHTML = renderStep3Right();
        wireStep3();
    } else {
        left.innerHTML = renderStep4Left();
        right.innerHTML = renderStep4Right();
        wireStep4();
    }
    window.lucide?.createIcons();
}

function renderStep1Left() {
    return `
    <div class="card__title">Recibo de luz</div>
    <div class="help">Sube un archivo PDF o una imagen (JPG/PNG) del recibo.</div>

    <input id="fileInput" type="file" accept="application/pdf,image/png,image/jpeg" hidden />

    <div class="dropzone" id="dropzone" tabindex="0" role="button" aria-label="Cargar recibo">
      <div class="dropzone__icon"><i data-lucide="upload"></i></div>
      <div style="min-width:0">
        <div class="dropzone__title">Arrastre y suelte aquí, o seleccione un archivo</div>
        <div class="dropzone__sub" id="fileHint">PDF o imagen</div>
      </div>
    </div>

    <div class="wizard-actions">
      <button class="btn btn--primary" id="btnAnalyze"><i data-lucide="scan"></i>Analizar</button>
      <button class="btn" id="btnManualCapture"><i data-lucide="keyboard"></i>Captura manual</button>
      <button class="btn" id="btnClear"><i data-lucide="trash-2"></i>Limpiar</button>
    </div>

    <div class="kpis">
      <div class="kpi">
        <div class="kpi__label">Tarifa</div>
        <div class="kpi__value" id="kpiTarifa">—</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Consumo periodo</div>
        <div class="kpi__value" id="kpiConsumo">—</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Total a pagar</div>
        <div class="kpi__value" id="kpiTotal">—</div>
      </div>
    </div>

    <div class="wizard-actions" style="justify-content:space-between">
      <div class="help" id="analyzeMsg"> </div>
      <button class="btn btn--success" id="btnStep1Next" disabled><i data-lucide="arrow-right"></i>Continuar</button>
    </div>
  `;
}

function renderStep1Right() {
    return `
    <div class="card__title">Vista previa</div>
    <div class="preview" id="preview">
      <div class="preview__empty" id="previewEmpty">Sin archivo cargado</div>
    </div>

    <div class="row" style="margin-top:12px">
      <div class="kpi" style="flex:1">
        <div class="kpi__label">No. de servicio</div>
        <div class="kpi__value" id="kpiServicio">—</div>
      </div>
      <div class="kpi" style="flex:1">
        <div class="kpi__label">Periodo</div>
        <div class="kpi__value" id="kpiPeriodo">—</div>
      </div>
    </div>

    <div class="kpi" style="margin-top:10px">
      <div class="kpi__label">Tipo de periodo</div>
      <div class="kpi__value" id="kpiTipoPeriodo">—</div>
    </div>
  `;
}

function wireStep1() {
    const fileInput = $('#fileInput');
    const dz = $('#dropzone');

    const pickFile = () => fileInput.click();

    dz.addEventListener('click', pickFile);
    dz.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ')
            pickFile();
    });

    dz.addEventListener('dragover', (e) => {
        e.preventDefault();
        dz.classList.add('is-dragover');
    });
    dz.addEventListener('dragleave', () => dz.classList.remove('is-dragover'));
    dz.addEventListener('drop', (e) => {
        e.preventDefault();
        dz.classList.remove('is-dragover');
        const f = e.dataTransfer.files?.[0];
        if (f)
            handleFileSelected(f);
    });

    fileInput.addEventListener('change', () => {
        const f = fileInput.files?.[0];
        if (f)
            handleFileSelected(f);
    });

    $('#btnClear').addEventListener('click', () => resetWizard(true));

    $('#btnManualCapture').addEventListener('click', () => {
        state.receiptFile = null;
        state.receiptCanvas = null;
        state.receipt = createEmptyReceiptData(state.selectedTariff);
        state.receipt.instalacion = state.receipt.instalacion || {};
        state.receipt.insumos = Array.isArray(state.receipt.insumos) ? state.receipt.insumos : [];
        state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct)) ? Number(state.receipt.impuestosPct) : 0.16;
        state.receipt.instalacion.panelesGrupos = Array.isArray(state.receipt.instalacion.panelesGrupos) ? state.receipt.instalacion.panelesGrupos : [];

        updateStep1FileHint();
        updateStep1Preview();
        updateStep1FromReceipt();
        $('#analyzeMsg').textContent = 'Captura manual habilitada. Completa los datos en el siguiente paso.';
        $('#btnStep1Next').disabled = false;
        setPillStatus('Captura manual', 'busy');
        toast({title: 'Captura manual', message: 'Puedes continuar sin subir un archivo.', icon: 'keyboard'});
    });

    $('#btnAnalyze').addEventListener('click', async () => {
        if (!state.receiptFile) {
            toast({title: 'Falta el archivo', message: 'Selecciona un recibo para continuar.', icon: 'alert-triangle'});
            return;
        }

        setPillStatus('Analizando…', 'busy');
        $('#analyzeMsg').textContent = 'Procesando recibo…';
        $('#btnAnalyze').disabled = true;

        try {
            const result = await analyzeReceiptFile(state.receiptFile, {
                selectedTariff: state.selectedTariff,
                onProgress: (p) => {
                    if (p?.message)
                        $('#analyzeMsg').textContent = p.message;
                }
            });

            if (!result?.ok) {
                throw new Error(result?.message || 'No se pudo analizar el recibo.');
            }

            state.receipt = result.parsed;
            // Inicializa estructuras del flujo (si no existen)
            state.receipt.instalacion = state.receipt.instalacion || {};
            state.receipt.insumos = Array.isArray(state.receipt.insumos) ? state.receipt.insumos : [];
            state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct)) ? Number(state.receipt.impuestosPct) : 0.16;
            state.receipt.instalacion.panelesGrupos = Array.isArray(state.receipt.instalacion.panelesGrupos) ? state.receipt.instalacion.panelesGrupos : [];
            state.receiptCanvas = result.canvas;
            updateStep1FromReceipt();

            setPillStatus('Listo', 'ok');
            $('#analyzeMsg').textContent = 'Recibo analizado correctamente.';
            $('#btnStep1Next').disabled = false;

            // Prefill datos del cliente
            state.client.nombre = state.client.nombre || state.receipt.nombre || '';
            state.client.direccion = state.client.direccion || state.receipt.direccion || '';

            toast({title: 'Recibo listo', message: 'Datos detectados y listos para cotizar.', icon: 'check-circle'});
        } catch (err) {
            console.error(err);
            state.receipt = createEmptyReceiptData(state.selectedTariff);
            state.receipt.instalacion = state.receipt.instalacion || {};
            state.receipt.insumos = Array.isArray(state.receipt.insumos) ? state.receipt.insumos : [];
            state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct)) ? Number(state.receipt.impuestosPct) : 0.16;
            state.receipt.instalacion.panelesGrupos = Array.isArray(state.receipt.instalacion.panelesGrupos) ? state.receipt.instalacion.panelesGrupos : [];
            state.receipt.analisisEstado = 'fallido';
            state.receipt.analisisMensaje = err?.message || 'No se pudo analizar el recibo.';
            updateStep1FromReceipt();
            $('#btnStep1Next').disabled = false;
            setPillStatus('Análisis fallido · captura manual', 'error');
            $('#analyzeMsg').textContent = `${state.receipt.analisisMensaje} Captura los datos manualmente para continuar.`;
            toast({title: 'Análisis fallido', message: 'Se habilitó la captura manual para continuar.', icon: 'x-circle'});
        } finally {
            $('#btnAnalyze').disabled = false;
        }
    });

    $('#btnStep1Next').addEventListener('click', () => gotoStep(2));

    // Initial
    updateStep1FileHint();
    updateStep1Preview();
    if ($('#kpiTarifa') && state.selectedTariff?.label)
        $('#kpiTarifa').textContent = state.selectedTariff.label;
}

async function handleFileSelected(file) {
    state.receiptFile = file;
    state.receipt = null;
    state.receiptCanvas = null;
    state.quote = null;
    state.savedQuote = null;

    updateStep1FileHint();
    updateStep1Preview();
    $('#btnStep1Next').disabled = true;
    $('#analyzeMsg').textContent = '';

    // Also reset KPIs
    $('#kpiTarifa').textContent = state.selectedTariff?.label || '—';
    $('#kpiConsumo').textContent = '—';
    $('#kpiTotal').textContent = '—';
    $('#kpiServicio').textContent = '—';
    $('#kpiPeriodo').textContent = '—';
    $('#kpiTipoPeriodo').textContent = '—';
}

function updateStep1FileHint() {
    const f = state.receiptFile;
    if (f) {
        $('#fileHint').textContent = `${f.name} · ${(f.size / 1024 / 1024).toFixed(2)} MB`;
        return;
    }
    $('#fileHint').textContent = state.receipt ? 'Captura manual activa' : 'PDF o imagen';
}

function updateStep1Preview() {
    const preview = $('#preview');
    const empty = $('#previewEmpty');

    // Clear
    preview.querySelectorAll('canvas,img').forEach(x => x.remove());
    if (!state.receiptFile) {
        empty.style.display = 'block';
        empty.textContent = state.receipt ? 'Captura manual activa' : 'Sin archivo cargado';
        return;
    }
    empty.style.display = 'none';

    const isPdf = state.receiptFile.type === 'application/pdf' || state.receiptFile.name.toLowerCase().endsWith('.pdf');
    if (!isPdf) {
        const url = URL.createObjectURL(state.receiptFile);
        const img = document.createElement('img');
        img.src = url;
        img.onload = () => URL.revokeObjectURL(url);
        preview.appendChild(img);
        return;
    }

    // PDF preview (nativo del navegador)
    const url = URL.createObjectURL(state.receiptFile);
    const iframe = document.createElement('iframe');
    iframe.className = 'preview__frame';
    iframe.src = url;
    iframe.onload = () => {
        // liberamos después de cargar
        setTimeout(() => URL.revokeObjectURL(url), 2500);
    };
    preview.appendChild(iframe);
}

function updateStep1FromReceipt() {
    const r = state.receipt;
    if (!r)
        return;
    $('#kpiTarifa').textContent = r.tarifa || '—';
    $('#kpiConsumo').textContent = r.consumoPeriodo ? `${formatNumber(r.consumoPeriodo)} kWh` : '—';
    $('#kpiTotal').textContent = r.totalAPagar ? formatCurrencyMXN(r.totalAPagar) : '—';
    $('#kpiServicio').textContent = r.servicio || '—';
    $('#kpiPeriodo').textContent = r.periodo?.raw || '—';
    $('#kpiTipoPeriodo').textContent = r.tipoPeriodo || '—';
}


function currentStep2Quote() {
    const receiptForQuote = structuredCloneSafe(state.receipt || {});
    applyTariffCalculationAssumptions(receiptForQuote, state.selectedTariff);
    return computeQuote(receiptForQuote, state.client || {}, state.params || {}, state.overrides || {});
}

function renderPackageCards() {
    const q = currentStep2Quote();
    return PACKAGE_PRESETS.map(pkg => {
        const selected = state.selectedPackage === pkg.key;
        const items = buildPackageItems(pkg.key, {quote: q, receipt: state.receipt, paneles: q.paneles, consumoMensual: q.consumoMensual});
        const total = items.reduce((acc, it) => acc + (Number(it.cantidad || 0) * Number(it.precio || 0)), 0);
        return `
      <button type="button" class="package-card ${selected ? 'is-selected' : ''}" data-package="${pkg.key}">
        <div class="package-card__head">
          <div>
            <div class="package-card__title">${escapeHtml(pkg.label)}</div>
            <div class="package-card__badge">${escapeHtml(pkg.badge)}</div>
          </div>
          <div class="package-card__total">${formatCurrencyMXN(total)}</div>
        </div>
        <div class="package-card__desc">${escapeHtml(pkg.description)}</div>
        <div class="package-card__meta">${items.length} insumos base · ${q.paneles || 0} paneles estimados</div>
      </button>
    `;
    }).join('');
}

function renderPackagePreviewItems() {
    if (!state.selectedPackage) {
        return '<div class="help">Selecciona un paquete para cargar automáticamente insumos y precios por defecto.</div>';
    }
    const q = currentStep2Quote();
    const items = buildPackageItems(state.selectedPackage, {quote: q, receipt: state.receipt, paneles: q.paneles, consumoMensual: q.consumoMensual});
    return `
    <div class="package-preview-list">
      ${items.map(it => `
        <div class="package-preview-item">
          <span>${escapeHtml(it.descripcion)}</span>
          <b>${Number(it.cantidad || 0)} ${escapeHtml(it.unidad || 'UD')} · ${formatCurrencyMXN(it.precio || 0)}</b>
        </div>
      `).join('')}
    </div>
  `;
}

function syncStep2State() {
    if (!state.receipt)
        state.receipt = createEmptyReceiptData(state.selectedTariff);
    const toNum = (v, fallback = 0) => {
        const n = Number(String(v ?? '').replace(/,/g, '').trim());
        return Number.isFinite(n) ? n : fallback;
    };

    state.receipt.servicio = ($('#rServicio')?.value || '').replace(/\D/g, '').slice(0, 15);
    const selectedTariffKey = ($('#rSelectedTariff')?.value || '').trim();
    const selectedTariff = TARIFFS.find(t => t.key === selectedTariffKey) || null;
    state.selectedTariff = selectedTariff || state.selectedTariff || null;
    if (selectedTariff?.periodo && $('#rTipoPeriodo')) {
        $('#rTipoPeriodo').value = selectedTariff.periodo;
    }
    state.receipt.tarifa = ($('#rTarifa')?.value || '').trim().toUpperCase();
    state.receipt.nombre = ($('#rNombre')?.value || '').trim();
    state.receipt.direccion = ($('#rDireccion')?.value || '').trim();
    state.receipt.periodo = state.receipt.periodo || {raw: '', start: null, end: null, days: 0};
    state.receipt.periodo.raw = ($('#rPeriodo')?.value || '').trim();
    state.receipt.tipoPeriodo = selectedTariff?.periodo || ($('#rTipoPeriodo')?.value || '').trim();
    state.receipt.periodo.days = state.receipt.tipoPeriodo === 'Bimestral' ? 60 : 30;
    state.receipt.consumoPeriodo = toNum($('#rConsumo')?.value, 0);
    state.receipt.totalAPagar = toNum($('#rTotal')?.value, 0);
    applyTariffCalculationAssumptions(state.receipt, state.selectedTariff);
    if (state.selectedTariff?.periodo && $('#rTipoPeriodo')) {
        $('#rTipoPeriodo').value = state.selectedTariff.periodo;
    }
    state.receipt.hilos = ($('#rHilos')?.value || '').trim();
    state.receipt.estado = ($('#rEstado')?.value || '').trim().toUpperCase();
    state.receipt.ajusteConsumo = {
        kwhMes: toNum($('#rAjusteKwh')?.value, 0),
        nota: ($('#rAjusteNota')?.value || '').trim(),
    };

    state.client.nombre = ($('#cNombre')?.value || '').trim();
    state.client.telefono = ($('#cTel')?.value || '').trim();
    state.client.email = ($('#cEmail')?.value || '').trim();
    state.client.direccion = ($('#cDir')?.value || '').trim();

    state.params.yieldKwhPerKwpMonth = toNum($('#pYield')?.value, state.params.yieldKwhPerKwpMonth || 135);
    state.params.panelWatts = toNum($('#pPanel')?.value, state.params.panelWatts || 550);
    state.params.costPerKwp = toNum($('#pCost')?.value, state.params.costPerKwp || 22000);
    state.params.contingencyPct = Math.max(0, Math.min(0.30, toNum($('#pCont')?.value, state.params.contingencyPct || 0.06)));

    const consumoMensualManual = toNum($('#oConsumoMensual')?.value, 0);
    state.overrides.consumoMensual = consumoMensualManual > 0 ? consumoMensualManual : null;

    const pManual = toNum($('#oPaneles')?.value, 0);
    state.overrides.paneles = pManual > 0 ? Math.round(pManual) : null;

    state.quoteMeta.panelModelo = ($('#mPanelModelo')?.value || '').trim();
    state.quoteMeta.panelDimensiones = ($('#mPanelDim')?.value || '').trim();
    state.quoteMeta.inversorModelo = ($('#mInversor')?.value || '').trim();
    state.quoteMeta.tipoTecho = ($('#mTecho')?.value || 'No especificado');
    state.quoteMeta.perdidasSombraPct = Number($('#mSombras')?.value || 0);
    state.quoteMeta.sombras = $('#mSombras')?.selectedOptions?.[0]?.textContent?.split('(')?.[0]?.trim() || 'No especificado';
    state.quoteMeta.notasFisicas = ($('#mNotasFisicas')?.value || '').trim();

    state.receipt.instalacion = {
        ...(state.receipt.instalacion || {}),
        tipoTecho: state.quoteMeta.tipoTecho,
        perdidasSombraPct: state.quoteMeta.perdidasSombraPct,
        sombras: state.quoteMeta.sombras,
        notasFisicas: state.quoteMeta.notasFisicas,
        panelModelo: state.quoteMeta.panelModelo,
        panelDimensiones: state.quoteMeta.panelDimensiones,
        inversorModelo: state.quoteMeta.inversorModelo,
        paqueteSeleccionado: state.selectedPackage || '',
    };

    state.receipt.insumos = Array.isArray(state.receipt.insumos) ? state.receipt.insumos : [];
    state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct)) ? Number(state.receipt.impuestosPct) : (state.preferences?.quoteDefaults?.taxPct || 0.16);
    state.quote = currentStep2Quote();
    return state.quote;
}

function getStep2Alerts() {
    const alerts = [];
    if (!state.receipt?.servicio)
        alerts.push('Captura o corrige el número de servicio.');
    if (!state.receipt?.nombre)
        alerts.push('Revisa el nombre del titular del recibo.');
    if (!state.receipt?.direccion)
        alerts.push('Revisa la dirección del suministro.');
    if (!state.client?.nombre)
        alerts.push('Captura el nombre del cliente.');
    return alerts;
}

function refreshStep2Summary() {
    const q = syncStep2State();
    const alerts = getStep2Alerts();
    $('#step2Holder') && ($('#step2Holder').textContent = state.receipt?.nombre || '—');
    $('#step2Address') && ($('#step2Address').textContent = state.receipt?.direccion || '—');
    $('#step2Servicio') && ($('#step2Servicio').textContent = state.receipt?.servicio || '—');
    $('#step2Periodo') && ($('#step2Periodo').textContent = state.receipt?.periodo?.raw || '—');
    $('#step2Total') && ($('#step2Total').textContent = state.receipt?.totalAPagar ? formatCurrencyMXN(state.receipt.totalAPagar) : '—');
    $('#step2Kwp') && ($('#step2Kwp').textContent = `${Number(q.kwp || 0).toFixed(2)} kWp`);
    $('#step2Panels') && ($('#step2Panels').textContent = `${formatNumber(q.paneles || 0)} paneles`);
    $('#step2Saving') && ($('#step2Saving').textContent = formatCurrencyMXN(q.ahorroMensual || 0));
    
    const consumoTexto = state.overrides?.consumoMensual ? `${formatNumber(state.overrides.consumoMensual)} kWh/mes (manual)` : `${formatNumber(q.consumoMensual || 0)} kWh/mes`;
    $('#step2ConsumoMensual') && ($('#step2ConsumoMensual').textContent = consumoTexto);
    $('#tariffImpactData') && ($('#tariffImpactData').innerHTML = getTariffImpact(state.selectedTariff, q).html);
    $('#step2TariffFormula') && ($('#step2TariffFormula').textContent = getTariffImpact(state.selectedTariff, q).formula);
    
    const comparisonBox = $('#tariffComparisonBox');
    if (comparisonBox)
        comparisonBox.outerHTML = renderTariffComparisonTable(state.selectedTariff);
    

    let packageText = 'Configuración personalizada';
    if (state.selectedPackage) {
        // Buscamos el objeto del paquete para obtener solo el nombre (label)
        const pkgObj = PACKAGE_PRESETS.find(p => p.key === state.selectedPackage);
        const numInsumos = Array.isArray(state.receipt?.insumos) ? state.receipt.insumos.length : 0;
        
        if (pkgObj) {
            packageText = `${pkgObj.label} · ${numInsumos} insumos · ${formatCurrencyMXN(q.inversion)}`;
        }
    }


    $('#step2Package') && ($('#step2Package').textContent = packageText);
    $('#packageLabel') && ($('#packageLabel').textContent = packageText);
    $('#kpiQuoteTotal') && ($('#kpiQuoteTotal').textContent = formatCurrencyMXN(q.inversion || 0));
    $('#step2Alerts') && ($('#step2Alerts').innerHTML = alerts.length
            ? alerts.map(msg => `<div class="review-alert">${escapeHtml(msg)}</div>`).join('')
            : '<div class="review-ok">Los campos clave del recibo ya quedaron listos para cotizar.</div>');
    
    const preview = $('#packagePreview');
    if (preview)
        preview.innerHTML = renderPackagePreviewItems();
}

function applyPackagePreset(packageKey) {
    syncStep2State();
    state.selectedPackage = packageKey;
    const q = state.quote || currentStep2Quote();
    const manuales = (state.receipt?.insumos || []).filter(it => it.isManual);
    const insumosPaquete = buildPackageItems(packageKey, {quote: q, receipt: state.receipt, paneles: q.paneles, consumoMensual: q.consumoMensual});
    state.receipt.insumos = [...insumosPaquete, ...manuales];
    state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct)) ? Number(state.receipt.impuestosPct) : (state.preferences?.quoteDefaults?.taxPct || 0.16);
    state.quote = currentStep2Quote();
    renderWizard();
    toast({title: 'Paquete aplicado', message: `${PACKAGE_PRESETS.find(p => p.key === packageKey)?.label || 'Paquete'} cargado con precios por defecto.`, icon: 'package'});
}

// ------------------------------
// Step 2
// ------------------------------

function renderStep2Left() {
    const r = state.receipt || createEmptyReceiptData(state.selectedTariff);
    const ajusteKwh = Number(r?.ajusteConsumo?.kwhMes || 0);
    const ajusteNota = (r?.ajusteConsumo?.nota || '').trim();
    const insumos = Array.isArray(r?.insumos) ? r.insumos : [];
    const impuestosPct = Number.isFinite(Number(r?.impuestosPct)) ? Number(r.impuestosPct) : (state.preferences?.quoteDefaults?.taxPct || 0.16);
    const q = currentStep2Quote();
    const panelesShown = (state.overrides?.paneles && Number(state.overrides.paneles) > 0) ? Number(state.overrides.paneles) : q.paneles;
    const panelesManualValue = (state.overrides?.paneles && Number(state.overrides.paneles) > 0) ? String(state.overrides.paneles) : '';
    return `
    <div class="card__title">Datos del recibo y configuración del sistema</div>
    <div class="help">Aquí puedes corregir nombre, dirección y número de servicio detectados, además de definir el sistema y el paquete de insumos antes de generar la cotización.</div>

    <div class="card__subtitle" style="margin-top:12px">Campos críticos del recibo</div>

    <div class="row" style="margin-top:10px">
      <div class="field">
        <label>No. de servicio</label>
        <input id="rServicio" placeholder="###########" value="${escapeAttr(r?.servicio || '')}" />
      </div>
      <div class="field">
        <label>Tarifa</label>
        <input id="rTarifa" placeholder="1B / DAC / PDBT / ..." value="${escapeAttr(r?.tarifa || '')}" />
      </div>
    </div>

    <div class="row" style="margin-top:10px">
      <div class="field">
        <label>Titular del recibo</label>
        <input id="rNombre" placeholder="Nombre detectado del recibo" value="${escapeAttr(r?.nombre || '')}" />
      </div>
      <div class="field">
        <label>Estado (abreviatura)</label>
        <input id="rEstado" placeholder="SON" value="${escapeAttr(String(r?.estado || ''))}" />
      </div>
    </div>

    <div class="field" style="margin-top:10px">
      <label>Dirección del suministro</label>
      <textarea id="rDireccion" rows="3" placeholder="Dirección detectada del recibo">${escapeHtml(r?.direccion || '')}</textarea>
    </div>

    <div class="row" style="margin-top:10px">
      <div class="field" style="flex:1.4">
        <label>Periodo facturado</label>
        <input id="rPeriodo" placeholder="DD MMM AA - DD MMM AA" value="${escapeAttr(r?.periodo?.raw || '')}" />
      </div>
      <div class="field" style="flex:1">
        <label>Tipo de periodo</label>
        <select id="rTipoPeriodo">
          <option ${r?.tipoPeriodo === 'Mensual' ? 'selected' : ''}>Mensual</option>
          <option ${r?.tipoPeriodo === 'Bimestral' ? 'selected' : ''}>Bimestral</option>
        </select>
      </div>
    </div>

    <div class="row" style="margin-top:10px">
      <div class="field">
        <label>Consumo del periodo (kWh)</label>
        <input id="rConsumo" type="number" min="0" step="1" value="${escapeAttr(String(r?.consumoPeriodo || 0))}" />
      </div>
      <div class="field">
        <label>Total a pagar (MXN)</label>
        <input id="rTotal" type="number" min="0" step="1" value="${escapeAttr(String(r?.totalAPagar || 0))}" />
      </div>
    </div>

    <div class="row" style="margin-top:10px">
      <div class="field">
        <label>No. hilos</label>
        <input id="rHilos" placeholder="1" value="${escapeAttr(String(r?.hilos || ''))}" />
      </div>
      <div class="field">
        <label>Ajuste de consumo (kWh/mes)</label>
        <input id="rAjusteKwh" type="number" step="1" value="${escapeAttr(String(ajusteKwh))}" />
      </div>
    </div>

    <div class="field" style="margin-top:10px">
      <label>Nota del ajuste de consumo</label>
      <textarea id="rAjusteNota" rows="2" placeholder="Ejemplo: Se considera una ampliación de carga futura.">${escapeHtml(ajusteNota)}</textarea>
    </div>

    <div class="card__subtitle" style="margin-top:14px">Datos del cliente</div>

    <div class="row" style="margin-top:10px">
      <div class="field">
        <label>Nombre del cliente</label>
        <input id="cNombre" placeholder="Nombre del cliente" value="${escapeAttr(state.client.nombre || r?.nombre || '')}" />
      </div>
      <div class="field">
        <label>Teléfono</label>
        <input id="cTel" placeholder="(###) ### ####" value="${escapeAttr(state.client.telefono)}" />
      </div>
    </div>

    <div class="row" style="margin-top:10px">
      <div class="field">
        <label>Correo</label>
        <input id="cEmail" placeholder="cliente@correo.com" value="${escapeAttr(state.client.email)}" />
      </div>
      <div class="field">
        <label>Tipo seleccionado</label>
        <select id="rSelectedTariff">
          <option value="">Selecciona un tipo…</option>
          ${TARIFFS.filter(t => t.kind !== 'cashvolt').map(t => `<option value="${escapeAttr(t.key)}" ${state.selectedTariff?.key === t.key ? 'selected' : ''}>${escapeHtml(t.label)}</option>`).join('')}
        </select>
      </div>
    </div>

    ${renderTariffImpactBox(state.selectedTariff, q, 'tariffImpactData')}

    <div class="field" style="margin-top:10px">
      <label>Dirección de instalación / contacto</label>
      <textarea id="cDir" rows="3" placeholder="Dirección donde se instalará el sistema">${escapeHtml(state.client.direccion || r?.direccion || '')}</textarea>
    </div>

    <div class="card__subtitle" style="margin-top:14px">Edición del sistema</div>

    <div class="grid cols-3" style="margin-top:10px">
      <div class="field">
        <label>Producción promedio (kWh/kWp-mes)</label>
        <input id="pYield" type="number" min="60" max="220" step="1" value="${escapeAttr(String(state.params.yieldKwhPerKwpMonth))}" />
      </div>
      <div class="field">
        <label>Panel (W)</label>
        <input id="pPanel" type="number" min="350" max="700" step="10" value="${escapeAttr(String(state.params.panelWatts))}" />
      </div>
      <div class="field">
        <label>Costo por kWp (MXN)</label>
        <input id="pCost" type="number" min="12000" max="60000" step="500" value="${escapeAttr(String(state.params.costPerKwp))}" />
      </div>
    </div>

    <div class="grid cols-3" style="margin-top:10px">
      <div class="field">
        <label>Contingencia</label>
        <input id="pCont" type="number" min="0" max="0.30" step="0.01" value="${escapeAttr(String(state.params.contingencyPct))}" />
      </div>
      <div class="field">
        <label>Paneles (manual, opcional)</label>
        <input id="oPaneles" type="number" min="1" step="1" value="${escapeAttr(panelesManualValue)}" placeholder="Auto: ${escapeAttr(String(panelesShown))}" />
      </div>
      <div class="field">
        <label>Modelo de panel</label>
        <input id="mPanelModelo" placeholder="Ej. JA Solar 550 W" value="${escapeAttr(state.quoteMeta.panelModelo || '')}" />
      </div>
    </div>

    <div class="grid cols-3" style="margin-top:10px">
      <div class="field">
        <label>Dimensiones del panel</label>
        <input id="mPanelDim" placeholder="Ej. 2.27 × 1.13 m" value="${escapeAttr(state.quoteMeta.panelDimensiones || '')}" />
      </div>
      <div class="field">
        <label>Tipo de techo</label>
        <select id="mTecho">
          ${['No especificado', 'Losa', 'Lámina', 'Teja', 'Otro'].map(x => `<option ${x === state.quoteMeta.tipoTecho ? 'selected' : ''}>${x}</option>`).join('')}
        </select>
      </div>
      <div class="field">
        <label>Sombras</label>
        <select id="mSombras">
          <option value="0" ${Number(state.quoteMeta.perdidasSombraPct || 0) === 0 ? 'selected' : ''}>Ninguna (0%)</option>
          <option value="0.10" ${Number(state.quoteMeta.perdidasSombraPct || 0) === 0.10 ? 'selected' : ''}>Baja (10%)</option>
          <option value="0.20" ${Number(state.quoteMeta.perdidasSombraPct || 0) === 0.20 ? 'selected' : ''}>Media (20%)</option>
          <option value="0.35" ${Number(state.quoteMeta.perdidasSombraPct || 0) === 0.35 ? 'selected' : ''}>Alta (35%)</option>
        </select>
      </div>
    </div>

    <div class="row" style="margin-top:10px">
      <div class="field" style="flex:1.2">
        <label>Modelo de inversor</label>
        <input id="mInversor" placeholder="Ej. Huawei SUN2000" value="${escapeAttr(state.quoteMeta.inversorModelo || '')}" />
      </div>
      <div class="field" style="flex:1">
        <label>Consumo mensual manual (opcional)</label>
        <input id="oConsumoMensual" type="number" min="0" step="1" value="${escapeAttr(String(state.overrides?.consumoMensual || ''))}" placeholder="Auto: ${escapeAttr(String(q.consumoMensual || 0))} kWh/mes" />
      </div>
    </div>

    <div class="field" style="margin-top:10px">
      <label>Consideraciones físicas / notas</label>
      <textarea id="mNotasFisicas" rows="3" placeholder="Área disponible, orientación, sombras y observaciones del techo.">${escapeHtml(state.quoteMeta.notasFisicas || '')}</textarea>
    </div>

    <div class="card__subtitle" style="margin-top:14px">Paquetes de insumos</div>
    <div class="help">Al seleccionar un paquete se cargarán insumos sugeridos y precios por defecto. Después puedes editar cualquier partida.</div>

    <div class="package-grid" style="margin-top:10px">
      ${renderPackageCards()}
    </div>

    <div class="row" style="justify-content:space-between; margin-top:10px; align-items:flex-start">
      <div class="help" id="packageLabel">${state.selectedPackage ? getPackageSummaryLabel(state.selectedPackage, {quote: q, receipt: state.receipt, paneles: q.paneles, consumoMensual: q.consumoMensual}) : 'Sin paquete seleccionado.'}</div>
      <button class="btn" id="btnClearPackage"><i data-lucide="eraser"></i>Quitar paquete</button>
    </div>

    <div class="package-preview" id="packagePreview" style="margin-top:10px">${renderPackagePreviewItems()}</div>

    <div class="card__subtitle" style="margin-top:14px">Insumos editables</div>
    <div class="help">Puedes agregar más partidas, cambiar cantidades, precios o quitar insumos del paquete.</div>

    <div class="row" style="margin-top:10px; align-items:flex-end">
      <div class="field" style="flex:2">
        <label>Agregar desde catálogo</label>
        <select id="insCatalog">
          <option value="">Selecciona un insumo…</option>
          ${INSUMO_CATALOG.map((it, idx) => `<option value="${idx}">${escapeHtml(it.codigo)} · ${escapeHtml(it.descripcion)} · ${formatCurrencyMXN(it.precio)}</option>`).join('')}
        </select>
      </div>
      <div style="display:flex; gap:10px; flex-wrap:wrap">
        <button class="btn" id="btnAddCatalog"><i data-lucide="plus"></i>Agregar</button>
        <button class="btn" id="btnAddManual"><i data-lucide="edit-3"></i>Agregar manual</button>
      </div>
    </div>

    <div style="overflow:auto; margin-top:10px">
      <table class="table table--tight" id="insTable">
        <thead>
          <tr>
            <th style="min-width:140px">CÓDIGO</th>
            <th style="min-width:280px">DESCRIPCIÓN</th>
            <th style="min-width:120px">CANTIDAD</th>
            <th style="min-width:120px">UNIDAD</th>
            <th style="min-width:140px">PRECIO</th>
            <th style="min-width:140px">TOTAL</th>
          </tr>
        </thead>
        <tbody>
          ${insumos.length ? insumos.map((it, i) => renderInsumoRow(it, i)).join('') : `
            <tr><td colspan="6" style="color:var(--muted)">Aún no hay insumos agregados.</td></tr>
          `}
        </tbody>
      </table>
    </div>

    <div class="insumos-summary" style="margin-top:10px">
      <div class="insumos-summary__row"><span>Subtotal</span><b id="insSubtotal">—</b></div>
      <div class="insumos-summary__row">
        <span>Impuestos (IVA %)</span>
        <div style="display:flex; align-items:center; gap:10px">
          <input id="insTaxPct" class="insumos-summary__tax" type="number" min="0" max="30" step="0.5" value="${escapeAttr(String(Math.round(impuestosPct * 100 * 10) / 10))}" />
          <b id="insTaxes">—</b>
        </div>
      </div>
      <div class="insumos-summary__row insumos-summary__row--total"><span>Total</span><b id="insTotal">—</b></div>
    </div>

    <div class="wizard-actions" style="justify-content:space-between">
      <button class="btn" id="btnBack2"><i data-lucide="arrow-left"></i>Volver</button>
      <button class="btn btn--success" id="btnNext2"><i data-lucide="arrow-right"></i>Continuar</button>
    </div>

    <div class="help" id="msg2" style="margin-top:10px"></div>
  `;
}

function renderStep2Right() {
    const r = state.receipt || {};
    const val = validateReceiptAgainstSelection(r, state.selectedTariff);
    const q = currentStep2Quote();
    const alerts = getStep2Alerts();
    return `
    <div class="card__title">Resumen de revisión</div>
    <div class="help">Verifica que los campos clave se hayan leído correctamente antes de continuar.</div>

    <div class="row" style="margin-top:10px">
      <div class="kpi" style="flex:1">
        <div class="kpi__label">Tipo seleccionado</div>
        <div class="kpi__value" style="font-size:13px">${escapeHtml(state.selectedTariff?.label || '—')}</div>
      </div>
      <div class="kpi" style="flex:1">
        <div class="kpi__label">Validación</div>
        <div class="kpi__value" style="font-size:13px">${val.ok ? 'Correcta' : 'Revisar'}</div>
      </div>
    </div>

    <div class="help" style="margin-top:8px">${escapeHtml(val.message)}</div>

    <div class="review-panel" style="margin-top:12px">
      <div class="review-row"><span>Titular del recibo</span><b id="step2Holder">${escapeHtml(r?.nombre || '—')}</b></div>
      <div class="review-row"><span>No. de servicio</span><b id="step2Servicio">${escapeHtml(r?.servicio || '—')}</b></div>
      <div class="review-row"><span>Dirección del suministro</span><b id="step2Address">${escapeHtml(r?.direccion || '—')}</b></div>
      <div class="review-row"><span>Periodo</span><b id="step2Periodo">${escapeHtml(r?.periodo?.raw || '—')}</b></div>
      <div class="review-row"><span>Total a pagar</span><b id="step2Total">${r?.totalAPagar ? formatCurrencyMXN(r.totalAPagar) : '—'}</b></div>
      <div class="review-row"><span>Impacto de tarifa</span><b id="step2TariffFormula">${escapeHtml(getTariffImpact(state.selectedTariff, q).formula)}</b></div>
      <div class="review-row"><span>Paquete de insumos</span><b id="step2Package">${state.selectedPackage ? getPackageSummaryLabel(state.selectedPackage, {quote: q, receipt: state.receipt, paneles: q.paneles, consumoMensual: q.consumoMensual}) : 'Sin paquete seleccionado'}</b></div>
    </div>

    <div class="card__subtitle" style="margin-top:14px">Pre-cotización</div>
    <div class="grid cols-2" style="margin-top:10px">
      <div class="kpi">
        <div class="kpi__label">Potencia estimada</div>
        <div class="kpi__value" id="step2Kwp">${Number(q.kwp || 0).toFixed(2)} kWp</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Paneles</div>
        <div class="kpi__value" id="step2Panels">${formatNumber(q.paneles || 0)} paneles</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Ahorro mensual</div>
        <div class="kpi__value" id="step2Saving">${formatCurrencyMXN(q.ahorroMensual || 0)}</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Consumo mensual usado</div>
        <div class="kpi__value" id="step2ConsumoMensual">${state.overrides?.consumoMensual ? `${formatNumber(state.overrides.consumoMensual)} kWh/mes (manual)` : `${formatNumber(q.consumoMensual || 0)} kWh/mes`}</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Total de cotización</div>
        <div class="kpi__value" id="kpiQuoteTotal">${formatCurrencyMXN(q.inversion || 0)}</div>
      </div>
    </div>

    ${renderTariffComparisonTable(state.selectedTariff)}

    <div class="card__subtitle" style="margin-top:14px">Campos por revisar</div>
    <div id="step2Alerts" style="margin-top:8px">
      ${alerts.length ? alerts.map(msg => `<div class="review-alert">${escapeHtml(msg)}</div>`).join('') : '<div class="review-ok">Los campos clave del recibo ya quedaron listos para cotizar.</div>'}
    </div>
  `;
}

// ------------------------------
// Insumos (CRUD) - Paso 2
// ------------------------------

function normalizeInsumo(it) {
    const n = (v) => {
        const x = Number(String(v ?? '').replace(/,/g, '').trim());
        return Number.isFinite(x) ? x : 0;
    };
    return {
        codigo: String(it?.codigo ?? '').trim(),
        descripcion: String(it?.descripcion ?? '').trim(),
        cantidad: Math.max(0, n(it?.cantidad ?? 1)),
        unidad: String(it?.unidad ?? 'UD').trim() || 'UD',
        precio: Math.max(0, n(it?.precio ?? 0)),
        isManual: it?.isManual || false
    };
}

function calcInsumoTotal(it) {
    const x = normalizeInsumo(it);
    return Math.round((x.cantidad * x.precio) * 100) / 100;
}

function computeInsumosTotals() {
    const r = state.receipt || {};
    const ins = Array.isArray(r.insumos) ? r.insumos : [];
    const subtotal = ins.reduce((acc, it) => acc + calcInsumoTotal(it), 0);
    const pct = Number.isFinite(Number(r.impuestosPct)) ? Number(r.impuestosPct) : 0.16;
    const impuestos = subtotal * pct;
    const total = subtotal + impuestos;
    return {
        subtotal: Math.round(subtotal * 100) / 100,
        impuestos: Math.round(impuestos * 100) / 100,
        total: Math.round(total * 100) / 100,
        pct,
    };
}

function renderInsumoRow(it, i) {
    const x = normalizeInsumo(it);
    const total = calcInsumoTotal(x);
    const unidades = ['UD', 'PZA', 'M', 'W', 'SERV'];
    return `
    <tr data-i="${i}">
      <td>
        <div class="tbl-code">
          <input class="tbl-input" data-field="codigo" value="${escapeAttr(x.codigo)}" placeholder="P1" />
          <button class="icon-btn icon-btn--sm" type="button" data-del="${i}" title="Eliminar"><i data-lucide="trash-2"></i></button>
        </div>
      </td>
      <td><input class="tbl-input" data-field="descripcion" value="${escapeAttr(x.descripcion)}" placeholder="Descripción" /></td>
      <td><input class="tbl-input" data-field="cantidad" type="number" min="0" step="0.01" value="${escapeAttr(String(x.cantidad))}" /></td>
      <td>
        <select class="tbl-select" data-field="unidad">
          ${unidades.map(u => `<option ${u === x.unidad ? 'selected' : ''}>${u}</option>`).join('')}
        </select>
      </td>
      <td><input class="tbl-input" data-field="precio" type="number" min="0" step="0.01" value="${escapeAttr(String(x.precio))}" /></td>
      <td><div class="tbl-total" data-total>${formatCurrencyMXN(total)}</div></td>
    </tr>
  `;
}

function setupInsumosCrud() {
    const r = state.receipt;
    if (!r)
        return;
    r.insumos = Array.isArray(r.insumos) ? r.insumos : [];
    r.impuestosPct = Number.isFinite(Number(r.impuestosPct)) ? Number(r.impuestosPct) : 0.16;

    const tbody = $('#insTable tbody');
    if (!tbody)
        return;

    const renderBody = () => {
        const ins = r.insumos;
        tbody.innerHTML = ins.length
                ? ins.map((it, i) => renderInsumoRow(it, i)).join('')
                : `<tr><td colspan="6" style="color:var(--muted)">Aún no hay insumos agregados.</td></tr>`;
        window.lucide?.createIcons();
        updateTotalsUI();
    };

    const updateTotalsUI = () => {
        const t = computeInsumosTotals();
        $('#insSubtotal') && ($('#insSubtotal').textContent = formatCurrencyMXN(t.subtotal));
        $('#insTaxes') && ($('#insTaxes').textContent = formatCurrencyMXN(t.impuestos));
        $('#insTotal') && ($('#insTotal').textContent = formatCurrencyMXN(t.total));

        // Reflejar en el KPI de la derecha
        const q = computeQuote(r, state.client, state.params, state.overrides);
        const kpi = $('#kpiQuoteTotal');
        if (kpi)
            kpi.textContent = formatCurrencyMXN(q.inversion || 0);
        if ($('#step2Package'))
            refreshStep2Summary();
    };

    // Inicial
    updateTotalsUI();

    // Agregar desde catálogo
    $('#btnAddCatalog')?.addEventListener('click', () => {
        const sel = $('#insCatalog');
        if (!sel || sel.value === '') {
            toast({title: 'Selecciona un insumo', message: 'Elige un insumo del catálogo para agregar.', icon: 'alert-triangle'});
            return;
        }
        const idx = Number(sel.value);
        if (!Number.isFinite(idx)) {
            toast({title: 'Selecciona un insumo', message: 'Elige un insumo del catálogo para agregar.', icon: 'alert-triangle'});
            return;
        }
        const base = INSUMO_CATALOG[idx];
        if (!base)
            return;
        r.insumos.push({codigo: base.codigo, descripcion: base.descripcion, cantidad: 1, unidad: base.unidad, precio: base.precio, isManual: true});
        if (sel)
            sel.value = '';
        state.quote = null;
        renderBody();
    });

    // Agregar manual
    $('#btnAddManual')?.addEventListener('click', () => {
        r.insumos.push({codigo: '', descripcion: '', cantidad: 1, unidad: 'UD', precio: 0, isManual: true});
        state.quote = null;
        renderBody();
    });

    // Impuestos (IVA %)
    $('#insTaxPct')?.addEventListener('input', () => {
        const n = Number(String($('#insTaxPct').value || '').replace(/,/g, '').trim());
        const pct = Number.isFinite(n) ? Math.max(0, Math.min(30, n)) / 100 : 0.16;
        r.impuestosPct = pct;
        state.quote = null;
        updateTotalsUI();
    });

    // Delegación para inputs/selects y eliminación
    const syncRow = (tr) => {
        const idx = Number(tr?.dataset?.i);
        if (!Number.isFinite(idx))
            return;
        const current = r.insumos[idx] || {};
        const read = (field) => tr.querySelector(`[data-field="${field}"]`);
        const codigo = read('codigo')?.value ?? '';
        const descripcion = read('descripcion')?.value ?? '';
        const cantidad = Number(read('cantidad')?.value ?? 0);
        const unidad = read('unidad')?.value ?? 'UD';
        const precio = Number(read('precio')?.value ?? 0);
        r.insumos[idx] = normalizeInsumo({...current, codigo, descripcion, cantidad, unidad, precio});
        // Actualizar total de la fila
        const total = calcInsumoTotal(r.insumos[idx]);
        tr.querySelector('[data-total]') && (tr.querySelector('[data-total]').textContent = formatCurrencyMXN(total));
    };

    tbody.addEventListener('input', (e) => {
        const tr = e.target.closest('tr[data-i]');
        if (!tr)
            return;
        syncRow(tr);
        state.quote = null;
        updateTotalsUI();
    });

    tbody.addEventListener('change', (e) => {
        const tr = e.target.closest('tr[data-i]');
        if (!tr)
            return;
        syncRow(tr);
        state.quote = null;
        updateTotalsUI();
    });

    tbody.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-del]');
        if (!btn)
            return;
        const idx = Number(btn.dataset.del);
        if (!Number.isFinite(idx))
            return;
        r.insumos.splice(idx, 1);
        state.quote = null;
        renderBody();
    });

    // Asegurar render inicial de filas con íconos
    renderBody();
}

function wireStep2() {
    if (!state.receipt) {
        gotoStep(1);
        return;
    }

    $('#btnBack2').addEventListener('click', () => gotoStep(1));

    const syncAndRefresh = debounce(() => {
        syncStep2State();
        refreshStep2Summary();
    }, 120);

    [
        '#rServicio', '#rTarifa', '#rNombre', '#rDireccion', '#rPeriodo', '#rTipoPeriodo', '#rConsumo', '#rTotal', '#rHilos', '#rEstado',
        '#rAjusteKwh', '#rAjusteNota', '#cNombre', '#cTel', '#cEmail', '#cDir', '#rSelectedTariff',
        '#pYield', '#pPanel', '#pCost', '#pCont', '#oConsumoMensual', '#oPaneles', '#mPanelModelo', '#mPanelDim', '#mTecho', '#mSombras', '#mInversor', '#mNotasFisicas'
    ].forEach(id => {
        $(id)?.addEventListener('input', syncAndRefresh);
        $(id)?.addEventListener('change', syncAndRefresh);
    });

    $$('#route-cotizador [data-package]').forEach(btn => {
        btn.addEventListener('click', () => applyPackagePreset(btn.dataset.package));
    });

    $('#btnClearPackage')?.addEventListener('click', () => {
        state.selectedPackage = '';
        if (state.receipt?.instalacion)
            state.receipt.instalacion.paqueteSeleccionado = '';
        
        //Eliminar paquete de la lista de insumos
        if (state.receipt?.insumos) {
            const insumosManuales = state.receipt.insumos.filter(insumo => insumo.isManual);
            state.receipt.insumos.length = 0; 
            state.receipt.insumos.push(...insumosManuales);
        }
        state.quote = null;
        
        renderWizard();
        toast({title: 'Paquete retirado', message: 'Puedes mantener insumos manuales o cargar otro paquete.', icon: 'eraser'});
    });

    $('#btnNext2').addEventListener('click', () => {
        const q = syncStep2State();
        const msg = $('#msg2');

        if (!state.client.nombre) {
            msg.textContent = 'Por favor, captura el nombre del cliente.';
            toast({title: 'Falta información', message: 'El nombre del cliente es obligatorio.', icon: 'alert-triangle'});
            return;
        }

        if (!state.receipt.servicio) {
            msg.textContent = 'Por favor, captura el número de servicio del recibo.';
            toast({title: 'Dato crítico faltante', message: 'El número de servicio es obligatorio para continuar.', icon: 'alert-triangle'});
            return;
        }

        msg.textContent = '';
        state.quote = q;
        gotoStep(3);
    });

    setupInsumosCrud();
    refreshStep2Summary();
}

// ------------------------------
// Step 3
// ------------------------------

function renderStep3Left() {
    const q = state.quote || currentStep2Quote();
    state.quote = q;
    let packageLabel = 'Configuración personalizada';
    if (state.selectedPackage) {
        const pkgObj = PACKAGE_PRESETS.find(p => p.key === state.selectedPackage);
        const numInsumos = Array.isArray(state.receipt?.insumos) ? state.receipt.insumos.length : 0;
        
        if (pkgObj) {
            packageLabel = `${pkgObj.label} · ${numInsumos} insumos · ${formatCurrencyMXN(q.inversion)}`;
        }
    }
    
    return `
    <div class="card__title">Resumen técnico y económico</div>
    <div class="help">Esta vista consolida la información corregida del recibo, el sistema propuesto y los montos estimados. Si necesitas cambiar algo, regresa al paso anterior.</div>

    <div class="grid cols-3" style="margin-top:12px">
      <div class="kpi">
        <div class="kpi__label">Potencia instalada</div>
        <div class="kpi__value">${q.kwp.toFixed(2)} kWp</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Paneles</div>
        <div class="kpi__value">${q.paneles}</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Inversión</div>
        <div class="kpi__value">${formatCurrencyMXN(q.inversion)}</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Ahorro mensual</div>
        <div class="kpi__value">${formatCurrencyMXN(q.ahorroMensual)}</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Retorno</div>
        <div class="kpi__value">${q.retornoAnios ? `${q.retornoAnios} años` : '—'}</div>
      </div>
      <div class="kpi">
        <div class="kpi__label">Producción anual</div>
        <div class="kpi__value">${formatNumber(q.produccionAnual)} kWh</div>
      </div>
    </div>

    <div class="review-panel" style="margin-top:14px">
      <div class="review-row"><span>Titular</span><b>${escapeHtml(state.receipt?.nombre || '—')}</b></div>
      <div class="review-row"><span>Cliente</span><b>${escapeHtml(state.client?.nombre || '—')}</b></div>
      <div class="review-row"><span>No. de servicio</span><b>${escapeHtml(state.receipt?.servicio || '—')}</b></div>
      <div class="review-row"><span>Dirección</span><b>${escapeHtml(state.receipt?.direccion || state.client?.direccion || '—')}</b></div>
      <div class="review-row"><span>Panel</span><b>${escapeHtml(state.quoteMeta?.panelModelo || 'Panel fotovoltaico')}</b></div>
      <div class="review-row"><span>Inversor</span><b>${escapeHtml(state.quoteMeta?.inversorModelo || 'Por definir')}</b></div>
      <div class="review-row"><span>Techo / sombras</span><b>${escapeHtml(state.quoteMeta?.tipoTecho || 'No especificado')} · ${escapeHtml(state.quoteMeta?.sombras || 'Sin dato')}</b></div>
      <div class="review-row"><span>Paquete</span><b>${escapeHtml(packageLabel)}</b></div>
    </div>

    <div class="card__subtitle" style="margin-top:14px">Observaciones del sistema</div>
    <div class="help">${escapeHtml(state.quoteMeta?.notasFisicas || 'Sin observaciones adicionales.')}</div>

    <div class="wizard-actions" style="justify-content:space-between; margin-top:16px">
      <button class="btn" id="btnBack3"><i data-lucide="arrow-left"></i>Volver a editar</button>
      <button class="btn btn--success" id="btnNext3"><i data-lucide="arrow-right"></i>Preparar formato final</button>
    </div>
  `;
}

function renderStep3Right() {
    return `
    <div class="card__title">Consumo y pagos</div>
    <div class="help">Histórico detectado en el recibo.</div>

    <div style="margin-top:12px">
      <canvas id="chart" height="180"></canvas>
    </div>

    <div class="card__subtitle" style="margin-top:12px">Detalle</div>
    <div style="overflow:auto">
      <table class="table" id="histTable">
        <thead>
          <tr>
            <th>#</th>
            <th>Consumo (kWh)</th>
            <th>Pago (MXN)</th>
          </tr>
        </thead>
        <tbody></tbody>
      </table>
    </div>
  `;
}

function wireStep3() {
    if (!state.receipt) {
        gotoStep(1);
        return;
    }

    $('#btnBack3').addEventListener('click', () => gotoStep(2));
    $('#btnNext3').addEventListener('click', () => gotoStep(4));

    renderHistoryTable();
    renderChart();
}

function renderHistoryTable() {
    const tbody = $('#histTable tbody');
    if (!tbody)
        return;
    const hist = (state.receipt?.historial || []).slice(-12);
    const rows = hist.length ? hist.map((h, i) => `
    <tr>
      <td>${i + 1}</td>
      <td>${formatNumber(h.kwh)}</td>
      <td>${formatCurrencyMXN(h.pago)}</td>
    </tr>
  `).join('') : `
    <tr><td colspan="3" style="color:var(--muted)">Sin datos de consumo histórico.</td></tr>
  `;
    tbody.innerHTML = rows;
}

function renderChart() {
    const el = $('#chart');
    if (!el || !window.Chart)
        return;

    // Destroy previous
    if (state.chart) {
        state.chart.destroy();
        state.chart = null;
    }

    const hist = (state.receipt?.historial || []).slice(-12);
    const labels = hist.map((_, i) => `P${i + 1}`);
    const consumos = hist.map(h => h.kwh);
    const pagos = hist.map(h => h.pago);

    state.chart = new window.Chart(el, {
        type: 'line',
        data: {
            labels,
            datasets: [
                {label: 'Consumo (kWh)', data: consumos, tension: 0.25},
                {label: 'Pago (MXN)', data: pagos, tension: 0.25},
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {labels: {color: getComputedStyle(document.documentElement).getPropertyValue('--muted')}},
                tooltip: {enabled: true},
            },
            scales: {
                x: {ticks: {color: getComputedStyle(document.documentElement).getPropertyValue('--muted')}, grid: {color: 'rgba(255,255,255,0.06)'}},
                y: {ticks: {color: getComputedStyle(document.documentElement).getPropertyValue('--muted')}, grid: {color: 'rgba(255,255,255,0.06)'}},
            }
        }
    });
}

// ------------------------------
// Step 4
// ------------------------------

function renderStep4Left() {
    const q = state.quote || computeQuote(state.receipt, state.client, state.params, state.overrides);
    state.quote = q;
    return `
    <div class="card__title">Exportar</div>
    <div class="help">Revisa el documento y genera el PDF cuando esté listo.</div>

    <div class="wizard-actions" style="justify-content:space-between; margin-top:12px">
      <button class="btn" id="btnBack4"><i data-lucide="arrow-left"></i>Volver</button>
      <div style="display:flex; gap:10px; flex-wrap:wrap; justify-content:flex-end">
        <button class="btn" id="btnExport"><i data-lucide="download"></i>Descargar PDF</button>
        <button class="btn" id="btnGoCashvoltFromQuote"><i data-lucide="cloud-upload"></i>CashVolt</button>
        <button class="btn btn--primary" id="btnSaveQuote"><i data-lucide="save"></i>Guardar cotización</button>
        <button class="btn btn--success" id="btnConfirmProject"><i data-lucide="check"></i>Confirmar proyecto</button>
      </div>
    </div>

    <div class="help" id="msg4" style="margin-top:10px"></div>
  `;
}

function renderStep4Right() {
    const q = state.quote || computeQuote(state.receipt, state.client, state.params, state.overrides);
    state.quote = q;

    const exportHtml = buildExportHtml({
        ...q,
        receipt: state.receipt,
        client: state.client,
        params: state.params,
        id: state.savedQuote?.id || '',
        selectedTariff: state.selectedTariff,
    });

    return `
    <div class="card__title">Formato final</div>
    <div class="help">Vista previa del documento final.</div>

    <div class="card__subtitle" style="margin-top:12px">Vista previa</div>
    <div class="preview" style="min-height:420px; align-items:flex-start; justify-content:flex-start; overflow:auto; padding:12px">
      ${exportHtml}
    </div>
  `;
}

function wireStep4() {
    if (!state.receipt) {
        gotoStep(1);
        return;
    }

    $('#btnBack4').addEventListener('click', () => gotoStep(3));
    $('#btnExport').addEventListener('click', exportPdf);
    $('#btnGoCashvoltFromQuote')?.addEventListener('click', () => {
        window.open('https://cashvolt.mx/public/login', '_blank');
    });

    $('#btnSaveQuote').addEventListener('click', () => {
        try {
            const q = persistQuote('Guardada');
            toast({title: 'Cotización guardada', message: `Se agregó al historial (${q.id}).`, icon: 'save'});
            $('#msg4').textContent = `Guardada como ${q.id}.`;
            renderHistorialTable();
        } catch (e) {
            toast({title: 'No se pudo guardar', message: e?.message || 'Revisa los datos.', icon: 'x-circle'});
        }
    });

    $('#btnConfirmProject').addEventListener('click', () => {
        try {
            const q = persistQuote('Confirmada');
            const project = saveProjectFromQuote({...q, status: 'Confirmada'});
            updateQuote(q.id, {status: 'Confirmada'});
            toast({title: 'Proyecto confirmado', message: `Se agregó a Proyectos (${project.id}).`, icon: 'check'});
            renderHistorialTable();
            renderProyectosTable();
            setRoute('proyectos');
        } catch (e) {
            toast({title: 'No se pudo confirmar', message: e?.message || 'Revisa los datos.', icon: 'x-circle'});
        }
    });

    // Ensure export doc uses saved id if exists
    if (state.savedQuote?.id) {
        $('#msg4').textContent = `Cotización: ${state.savedQuote.id}`;
    }

    // Render chart inside export preview
    renderExportDocChart({root: document});
}

function renderExportDocChart( { root }){
    const canvas = (root || document).querySelector('#exportChart');
    if (!canvas || !window.Chart)
        return;

    // Destroy previous chart
    if (state.exportChart) {
        try {
            state.exportChart.destroy();
        } catch {
        }
        state.exportChart = null;
    }

    const hist = (state.receipt?.historial || []).slice(-12);
    const labels = (hist.length ? hist : Array.from({length: 12}, () => ({}))).map((_, i) => `P${i + 1}`);
    const consumos = hist.length ? hist.map(h => Number(h.kwh || 0)) : labels.map(() => 0);
    const prodMensual = Number(state.quote?.produccionMensual || 0);
    const produccion = labels.map(() => prodMensual);

    state.exportChart = new window.Chart(canvas, {
        type: 'line',
        data: {
            labels,
            datasets: [
                {label: 'Consumo (kWh)', data: consumos, tension: 0.25},
                {label: 'Producción estimada (kWh/mes)', data: produccion, tension: 0.15},
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {labels: {color: 'rgba(0,0,0,0.70)'}},
                tooltip: {enabled: true},
            },
            scales: {
                x: {ticks: {color: 'rgba(0,0,0,0.65)'}, grid: {color: 'rgba(0,0,0,0.08)'}},
                y: {ticks: {color: 'rgba(0,0,0,0.65)'}, grid: {color: 'rgba(0,0,0,0.08)'}},
            }
        }
    });
}

function persistQuote(status) {
    if (!state.client.nombre)
        throw new Error('Falta el nombre del cliente.');
    if (!state.receipt?.servicio)
        throw new Error('Falta el No. de servicio.');

    // If already saved, just update status
    if (state.savedQuote?.id) {
        const upd = updateQuote(state.savedQuote.id, {
            status,
            client: state.client,
            receipt: state.receipt,
            quote: state.quote,
            params: state.params,
            selectedTariff: state.selectedTariff,
            overrides: state.overrides,
            quoteMeta: state.quoteMeta,
        });
        state.savedQuote = upd;
        return upd;
    }

    const q = saveQuote({
        status,
        client: state.client,
        receipt: state.receipt,
        quote: state.quote,
        params: state.params,
        selectedTariff: state.selectedTariff,
        overrides: state.overrides,
        quoteMeta: state.quoteMeta,
    });
    state.savedQuote = q;
    return q;
}

async function exportPdf() {
    if (!window.html2canvas || !window.jspdf) {
        toast({title: 'Exportación no disponible', message: 'Faltan librerías de exportación en el navegador.', icon: 'alert-triangle'});
        return;
    }

    setPillStatus('Generando PDF…', 'busy');

    try {
        // Ensure we have an id
        if (!state.savedQuote?.id) {
            // Save silently to keep consistent naming
            state.savedQuote = persistQuote('Guardada');
        }

        const area = $('#exportDoc');
        if (!area)
            throw new Error('No se encontró el documento de exportación.');

        // Asegura que la gráfica del documento esté renderizada antes de capturar
        renderExportDocChart({root: document});

        // Pequeña espera para asegurar que el canvas tenga contenido
        await new Promise(r => setTimeout(r, 60));

        const canvas = await window.html2canvas(area, {scale: 2, useCORS: true, backgroundColor: '#ffffff'});
        const imgData = canvas.toDataURL('image/jpeg', 0.95);

        const {jsPDF} = window.jspdf;
        const pdf = new jsPDF({orientation: 'p', unit: 'pt', format: 'a4'});

        const pageW = pdf.internal.pageSize.getWidth();
        const pageH = pdf.internal.pageSize.getHeight();
        const margin = 28;

        const imgW = pageW - margin * 2;
        const imgH = canvas.height * (imgW / canvas.width);

        let heightLeft = imgH;
        let position = margin;

        pdf.addImage(imgData, 'JPEG', margin, position, imgW, imgH);
        heightLeft -= (pageH - margin * 2);

        while (heightLeft > 0) {
            pdf.addPage();
            position = margin - (imgH - heightLeft);
            pdf.addImage(imgData, 'JPEG', margin, position, imgW, imgH);
            heightLeft -= (pageH - margin * 2);
        }

        const name = (state.client.nombre || 'Cliente').trim().replace(/\s+/g, '_').slice(0, 36);
        const filename = `Cotizacion_SECOM_${name}_${state.savedQuote.id}.pdf`;
        pdf.save(filename);

        toast({title: 'PDF generado', message: 'Descarga iniciada.', icon: 'download'});
        setPillStatus('Listo', 'ok');
    } catch (e) {
        console.error(e);
        toast({title: 'No se pudo exportar', message: e?.message || 'Intenta de nuevo.', icon: 'x-circle'});
        setPillStatus('Error', 'error');
    }
}

// ------------------------------
// Exportación (Plantilla 5 páginas)
// ------------------------------

async function fetchAsDataUrl(url) {
    const res = await fetch(url);
    if (!res.ok)
        throw new Error(`No se pudo cargar: ${url}`);
    const blob = await res.blob();
    return await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error('No se pudo leer la imagen.'));
        reader.readAsDataURL(blob);
    });
}

async function exportTemplatePdfFromState() {
    // Asegura id (para nombre de archivo)
    if (!state.savedQuote?.id) {
        state.savedQuote = persistQuote('Guardada');
    }
    const id = state.savedQuote.id;
    const name = (state.client?.nombre || 'Cliente').trim().replace(/\s+/g, '_').slice(0, 36);
    await exportTemplatePdf({id, filename: `Cotizacion_SECOM_${name}_${id}.pdf`});
}

async function exportTemplatePdfFromStoredQuote(q) {
    const id = q.id;
    const name = (q.client?.nombre || q.receipt?.nombre || 'Cliente').trim().replace(/\s+/g, '_').slice(0, 36);
    await exportTemplatePdf({id, filename: `Cotizacion_SECOM_${name}_${id}.pdf`});
}

async function exportTemplatePdf( { id, filename }){
    if (!window.jspdf) {
        toast({title: 'Exportación no disponible', message: 'Falta jsPDF en el navegador.', icon: 'alert-triangle'});
        return;
    }

    setPillStatus('Generando PDF…', 'busy');

    try {
        const {jsPDF} = window.jspdf;
        const pdf = new jsPDF({orientation: 'p', unit: 'pt', format: 'a4'});
        const pageW = pdf.internal.pageSize.getWidth();
        const pageH = pdf.internal.pageSize.getHeight();

        const urls = [
            'assets/template/page1.jpg',
            'assets/template/page2.jpg',
            'assets/template/page3.jpg',
            'assets/template/page4.jpg',
            'assets/template/page5.jpg',
        ];

        for (let i = 0; i < urls.length; i++) {
            const imgData = await fetchAsDataUrl(urls[i]);
            if (i > 0)
                pdf.addPage();
            pdf.addImage(imgData, 'JPEG', 0, 0, pageW, pageH);

            // (Opcional) se puede agregar folio/fecha si se desea en el futuro.
        }

        pdf.save(filename || `Cotizacion_SECOM_${id || 'SIN_ID'}.pdf`);
        toast({title: 'PDF generado', message: 'Descarga iniciada.', icon: 'download'});
        setPillStatus('Listo', 'ok');
    } catch (e) {
        console.error(e);
        toast({title: 'No se pudo exportar', message: e?.message || 'Intenta de nuevo.', icon: 'x-circle'});
        setPillStatus('Error', 'error');
}
}

function resetWizard(full = false) {
    state.wizardStep = 1;
    if (full) {
        state.receiptFile = null;
        state.receipt = null;
        state.receiptCanvas = null;
        state.client = {nombre: '', telefono: '', email: '', direccion: ''};
        state.params = structuredCloneSafe(getDefaultQuoteParams());
        state.selectedPackage = '';
        state.quote = null;
        state.savedQuote = null;
    }
    // Si el wizard está montado, re-renderiza; si no, solo limpia estado.
    if ($('#stepper'))
        buildStepper();
    if ($('#wizardLeft'))
        renderWizard();
    setPillStatus('Listo', 'ok');
}

// ------------------------------
// Historial
// ------------------------------

function renderHistorialRoute() {
    const root = $('#route-historial');
    root.innerHTML = `
    <div class="grid" style="gap:14px">
      <div class="card">
        <div class="row" style="align-items:center; gap:12px">
          <div style="flex:1">
            <div class="card__title">Lista de cotizaciones</div>
            <div class="card__subtitle">Búsqueda y acciones rápidas</div>
          </div>
          <button class="btn btn--primary" id="btnNewQuoteFromList"><i data-lucide="plus"></i>Crear cotización</button>
          <div class="field" style="max-width:320px">
            <label>Buscar</label>
            <input id="histSearch" placeholder="Cliente, tarifa o No. de servicio" />
          </div>
        </div>

        <div style="overflow:auto; margin-top:10px">
          <table class="table" id="historialTable">
            <thead>
              <tr>
                <th>ID</th>
                <th>Cliente</th>
                <th>Tarifa</th>
                <th>Periodo</th>
                <th>Total</th>
                <th>Estatus</th>
                <th style="text-align:right">Acciones</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </div>
  `;

    $('#btnNewQuoteFromList')?.addEventListener('click', startNewQuoteFlow);
    $('#histSearch').addEventListener('input', debounce(renderHistorialTable, 120));
    renderHistorialTable();
}

function renderHistorialTable() {
    const tbody = $('#historialTable tbody');
    if (!tbody)
        return;

    const term = ($('#histSearch')?.value || '').toLowerCase().trim();
    const quotes = getQuotes();

    const filtered = term ? quotes.filter(q => {
        const c = (q.client?.nombre || q.receipt?.nombre || '').toLowerCase();
        const t = (q.receipt?.tarifa || '').toLowerCase();
        const s = (q.receipt?.servicio || '').toLowerCase();
        return c.includes(term) || t.includes(term) || s.includes(term) || (q.id || '').toLowerCase().includes(term);
    }) : quotes;

    tbody.innerHTML = filtered.length ? filtered.map(q => {
        const badge = q.status === 'Confirmada' ? 'badge--success' : 'badge--warn';
        const periodo = q.receipt?.periodo?.raw || '';
        const total = Number(q.quote?.inversion || q.quote?.totalInsumos || q.receipt?.totalAPagar || 0);

        return `
      <tr>
        <td>${escapeHtml(q.id)}</td>
        <td>${escapeHtml(q.client?.nombre || q.receipt?.nombre || '-')}</td>
        <td>${escapeHtml(q.receipt?.tarifa || '-')}</td>
        <td>${escapeHtml(periodo || '-')}</td>
        <td>${formatCurrencyMXN(total)}</td>
        <td><span class="badge ${badge}">${escapeHtml(q.status || '—')}</span></td>
        <td>
          <div class="actions">
            <button class="btn" data-action="view" data-id="${q.id}"><i data-lucide="eye"></i>Ver</button>
            <button class="btn" data-action="edit" data-id="${q.id}"><i data-lucide="pencil"></i>Editar</button>
            <button class="btn" data-action="export" data-id="${q.id}"><i data-lucide="download"></i>PDF</button>
            <button class="btn btn--success" data-action="confirm" data-id="${q.id}"><i data-lucide="check"></i>Proyecto</button>
          </div>
        </td>
      </tr>
    `;
    }).join('') : `
    <tr><td colspan="7" style="color:var(--muted)">Aún no hay cotizaciones guardadas.</td></tr>
  `;

    window.lucide?.createIcons();

    tbody.querySelectorAll('button[data-action]').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.dataset.id;
            const action = btn.dataset.action;
            const q = getQuotes().find(x => x.id === id);
            if (!q)
                return;

            if (action === 'view') {
                openModalForQuote(q);
            } else if (action === 'edit') {
                openQuoteEditor(q);
            } else if (action === 'export') {
                exportPdfFromStoredQuote(q);
            } else if (action === 'confirm') {
                const project = saveProjectFromQuote({...q, status: 'Confirmada'});
                updateQuote(q.id, {status: 'Confirmada'});
                toast({title: 'Proyecto confirmado', message: `Se agregó a Proyectos (${project.id}).`, icon: 'check'});
                renderHistorialTable();
                renderProyectosTable();
                setRoute('proyectos');
            }
        });
    });
}

function openModalForQuote(q) {
    const r = q.receipt || {};
    const c = q.client || {};

    // 1. Extraer el nombre/etiqueta del paquete
    const pkgKey = r.instalacion?.paqueteSeleccionado || '';
    let pkgText = 'Configuración personalizada';
    
    if (pkgKey) {
        const pkgObj = PACKAGE_PRESETS.find(p => p.key === pkgKey);
        const numInsumos = Array.isArray(r.insumos) ? r.insumos.length : 0;
        const inversionActual = q.quote?.inversion || 0;

        if (pkgObj) {
            pkgText = `${pkgObj.label} · ${numInsumos} insumos · ${formatCurrencyMXN(inversionActual)}`;
        }
    }
    
    // 2. Construir las filas de la tabla de insumos
    const insumos = Array.isArray(r.insumos) ? r.insumos : [];
    const insumosHtml = insumos.length ? insumos.map(it => `
    <tr>
      <td>${escapeHtml(it.codigo || '-')}</td>
      <td>${escapeHtml(it.descripcion || '-')}</td>
      <td>${Number(it.cantidad || 0)} ${escapeHtml(it.unidad || 'UD')}</td>
      <td style="text-align:right">${formatCurrencyMXN(Number(it.cantidad || 0) * Number(it.precio || 0))}</td>
    </tr>
  `).join('') : `<tr><td colspan="4" style="color:var(--muted)">No hay insumos agregados.</td></tr>`;

    // 3. Agregar la nueva tarjeta al diseño del modal
    const body = `
    <div class="grid cols-2">
      <div class="card" style="box-shadow:none">
        <div class="card__title">Cliente</div>
        <div class="help"><b>${escapeHtml(c.nombre || r.nombre || '-')}</b></div>
        <div class="help">${escapeHtml(c.telefono || '-')} · ${escapeHtml(c.email || '-')}</div>
        <div class="help" style="margin-top:8px">${escapeHtml(c.direccion || r.direccion || '-')}</div>
      </div>
      <div class="card" style="box-shadow:none">
        <div class="card__title">Recibo</div>
        <div class="help">No. de servicio: <b>${escapeHtml(r.servicio || '-')}</b></div>
        <div class="help">Tarifa: <b>${escapeHtml(r.tarifa || '-')}</b></div>
        <div class="help">Periodo: <b>${escapeHtml(r.periodo?.raw || '-')}</b></div>
        <div class="help">Total a pagar: <b>${formatCurrencyMXN(r.totalAPagar || 0)}</b></div>
      </div>
    </div>

    <div class="card" style="box-shadow:none; margin-top:12px">
      <div class="card__title">Resumen de cotización</div>
      <div class="row">
        <div class="kpi" style="flex:1">
          <div class="kpi__label">Potencia</div>
          <div class="kpi__value">${Number(q.quote?.kwp || 0).toFixed(2)} kWp</div>
        </div>
        <div class="kpi" style="flex:1">
          <div class="kpi__label">Paneles</div>
          <div class="kpi__value">${escapeHtml(q.quote?.paneles || '—')}</div>
        </div>
        <div class="kpi" style="flex:1">
          <div class="kpi__label">Inversión</div>
          <div class="kpi__value">${formatCurrencyMXN(q.quote?.inversion || 0)}</div>
        </div>
      </div>
    </div>

    <div class="card" style="box-shadow:none; margin-top:12px">
      <div class="card__title">Paquete e Insumos</div>
      <div class="help" style="margin-bottom:10px">Paquete base: <b>${escapeHtml(pkgText)}</b></div>
      <div style="max-height:180px; overflow-y:auto;">
        <table class="table table--tight">
          <thead>
            <tr>
              <th style="text-align:left">Código</th>
              <th style="text-align:left">Descripción</th>
              <th style="text-align:left">Cantidad</th>
              <th style="text-align:right">Total</th>
            </tr>
          </thead>
          <tbody>
            ${insumosHtml}
          </tbody>
        </table>
      </div>
    </div>
  `;

    openModal({
        title: q.id,
        subtitle: `Creada ${formatDateTime(q.createdAt)} · Estatus: ${q.status}`,
        bodyHtml: body,
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cerrar</button>
      <button class="btn" id="mEdit"><i data-lucide="pencil"></i>Editar</button>
      <button class="btn" id="mExport"><i data-lucide="download"></i>PDF</button>
      <button class="btn btn--success" id="mProject"><i data-lucide="check"></i>Confirmar proyecto</button>
    `
    });

    $('#mEdit')?.addEventListener('click', () => {
        closeModal();
        openQuoteEditor(q);
    });
    $('#mExport')?.addEventListener('click', () => exportPdfFromStoredQuote(q));
    $('#mProject')?.addEventListener('click', () => {
        const project = saveProjectFromQuote({...q, status: 'Confirmada'});
        updateQuote(q.id, {status: 'Confirmada'});
        toast({title: 'Proyecto confirmado', message: `Se agregó a Proyectos (${project.id}).`, icon: 'check'});
        closeModal();
        renderHistorialTable();
        renderProyectosTable();
        setRoute('proyectos');
    });
}

function openQuoteEditor(q) {
    // Cargar una cotización guardada en el flujo del cotizador
    state.selectedTariff = guessTariffFromQuote(q);
    // Abrir directamente en el paso de edición de datos y sistema
    state.wizardStep = 2;
    state.receiptFile = null; // no se conserva el archivo
    state.receiptCanvas = null;
    state.receipt = structuredCloneSafe(q.receipt || {});
    state.client = structuredCloneSafe(q.client || {nombre: '', telefono: '', email: '', direccion: ''});
    state.params = structuredCloneSafe(q.params || state.params);
    state.quote = structuredCloneSafe(q.quote || null);
    state.overrides = structuredCloneSafe(q.overrides || {paneles: null, consumoMensual: null});
    state.quoteMeta = structuredCloneSafe(q.quoteMeta || q.receipt?.instalacion || state.quoteMeta);
    state.selectedPackage = q.receipt?.instalacion?.paqueteSeleccionado || '';
    state.savedQuote = q;

    renderCotizadorRoute();
    setRoute('cotizador');
    gotoStep(2);
    toast({title: 'Edición', message: `Cotización ${q.id} cargada para editar.`, icon: 'pencil'});
}

function guessTariffFromQuote(q) {
    if (q?.selectedTariff?.key) {
        const found = TARIFFS.find(t => t.key === q.selectedTariff.key);
        if (found)
            return found;
    }
    const tp = q?.receipt?.tipoPeriodo || '';
    const isBim = String(tp).toLowerCase().includes('bim');
    // Prefer the family inferred from the menu (if saved) else default to Doméstica
    const candidates = TARIFFS.filter(t => t.kind !== 'cashvolt');
    const pref = candidates.find(t => String(t.label || '').toLowerCase().includes('dom')) || candidates[0];
    if (!pref)
        return candidates[0] || null;
    // Map by period
    const match = candidates.find(t => (isBim ? t.periodo === 'Bimestral' : t.periodo === 'Mensual') && String(t.label || '').toLowerCase().includes('doméstica'));
    return match || pref;
}

function structuredCloneSafe(obj) {
    try {
        return structuredClone(obj);
    } catch {
        return JSON.parse(JSON.stringify(obj || {}));
    }
}

async function exportPdfFromStoredQuote(q) {
    // Create a hidden export area in modal and reuse export pipeline
    const wrapper = document.createElement('div');
    wrapper.style.position = 'fixed';
    wrapper.style.left = '-9999px';
    wrapper.style.top = '0';
    wrapper.style.width = '800px';

    // Recalcular para reflejar insumos y/o cambios guardados en el recibo
    const recomputed = computeQuote(q.receipt || {}, q.client || {}, q.params || {}, q.overrides || {});
    const html = buildExportHtml({...recomputed, receipt: q.receipt, client: q.client, params: q.params, id: q.id, selectedTariff: q.selectedTariff});
    wrapper.innerHTML = html;
    document.body.appendChild(wrapper);

    try {
        setPillStatus('Generando PDF…', 'busy');
        const area = wrapper.querySelector('#exportDoc');

        // Render chart into the export doc
        const chartCanvas = wrapper.querySelector('#exportChart');
        let tmpChart = null;
        if (chartCanvas && window.Chart) {
            const hist = (q.receipt?.historial || []).slice(-12);
            const labels = (hist.length ? hist : Array.from({length: 12}, () => ({}))).map((_, i) => `P${i + 1}`);
            const consumos = hist.length ? hist.map(h => Number(h.kwh || 0)) : labels.map(() => 0);
            const prodMensual = Number(recomputed?.produccionMensual || 0);
            const produccion = labels.map(() => prodMensual);
            tmpChart = new window.Chart(chartCanvas, {
                type: 'line',
                data: {
                    labels,
                    datasets: [
                        {label: 'Consumo (kWh)', data: consumos, tension: 0.25},
                        {label: 'Producción estimada (kWh/mes)', data: produccion, tension: 0.15},
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {legend: {labels: {color: 'rgba(0,0,0,0.70)'}}},
                    scales: {
                        x: {ticks: {color: 'rgba(0,0,0,0.65)'}, grid: {color: 'rgba(0,0,0,0.08)'}},
                        y: {ticks: {color: 'rgba(0,0,0,0.65)'}, grid: {color: 'rgba(0,0,0,0.08)'}},
                    }
                }
            });
        }

        // Guarda referencia para limpieza
        wrapper.__tmpChart = tmpChart;

        await new Promise(r => setTimeout(r, 60));

        const canvas = await window.html2canvas(area, {scale: 2, useCORS: true, backgroundColor: '#ffffff'});
        const imgData = canvas.toDataURL('image/jpeg', 0.95);

        const {jsPDF} = window.jspdf;
        const pdf = new jsPDF({orientation: 'p', unit: 'pt', format: 'a4'});

        const pageW = pdf.internal.pageSize.getWidth();
        const pageH = pdf.internal.pageSize.getHeight();
        const margin = 28;

        const imgW = pageW - margin * 2;
        const imgH = canvas.height * (imgW / canvas.width);

        let heightLeft = imgH;
        let position = margin;

        pdf.addImage(imgData, 'JPEG', margin, position, imgW, imgH);
        heightLeft -= (pageH - margin * 2);
        while (heightLeft > 0) {
            pdf.addPage();
            position = margin - (imgH - heightLeft);
            pdf.addImage(imgData, 'JPEG', margin, position, imgW, imgH);
            heightLeft -= (pageH - margin * 2);
        }
        const name = (q.client?.nombre || 'Cliente').trim().replace(/\s+/g, '_').slice(0, 36);
        pdf.save(`Cotizacion_SECOM_${name}_${q.id}.pdf`);

        toast({title: 'PDF generado', message: 'Descarga iniciada.', icon: 'download'});
        setPillStatus('Listo', 'ok');
    } catch (e) {
        console.error(e);
        toast({title: 'No se pudo exportar', message: e?.message || 'Intenta de nuevo.', icon: 'x-circle'});
        setPillStatus('Error', 'error');
    } finally {
        try {
            wrapper.__tmpChart?.destroy();
        } catch {
        }
        wrapper.remove();
    }
}



// ------------------------------
// Paquetes (Catálogo global)
// ------------------------------

function loadPackageCatalogSafe(silent = false) {
    try {
        const items = getPaquetes();
        state.paquetes.items = Array.isArray(items) ? items : [];
        setPackageCatalog(state.paquetes.items);
        return state.paquetes.items;
    } catch (e) {
        console.warn('No se pudo cargar el catálogo de paquetes.', e);
        if (!silent) {
            toast({title: 'Paquetes no disponibles', message: e?.message || 'No se pudo consultar la base de datos.', icon: 'alert-triangle'});
        }
        setPackageCatalog(state.paquetes.items || []);
        return state.paquetes.items || [];
    }
}

function renderPaquetesRoute() {
    const root = $('#route-paquetes');
    if (!root) {
        return;
    }

    loadInsumoCatalogSafe(true);
    const items = loadPackageCatalogSafe(true);
    const active = items.filter(it => it.activo !== false).length;
    const total = items.reduce((acc, pkg) => acc + calcPackageTotal(pkg).total, 0);
    const avg = items.length ? total / items.length : 0;

    root.innerHTML = `
    <div class="grid" style="gap:14px">
      <div class="grid cols-3">
        <div class="card"><div class="card__title">${items.length}</div><div class="card__subtitle">Paquetes registrados</div></div>
        <div class="card"><div class="card__title">${active}</div><div class="card__subtitle">Activos</div></div>
        <div class="card"><div class="card__title">${formatCurrencyMXN(avg)}</div><div class="card__subtitle">Promedio por paquete</div></div>
      </div>

      <div class="card">
        <div class="row" style="align-items:center">
          <div style="flex:1">
            <div class="card__title">Gestión de paquetes</div>
            <div class="card__subtitle">Administra paquete básico, intermedio, avanzado y nuevos paquetes comerciales</div>
          </div>
          <button class="btn btn--primary" id="btnNewPaquete"><i data-lucide="plus"></i>Crear paquete</button>
          <div class="field" style="max-width:340px">
            <label>Buscar</label>
            <input id="paquetesSearch" placeholder="Nombre, descripción o estatus" value="${escapeAttr(state.paquetes.search || '')}" />
          </div>
        </div>

        <div class="review-ok" style="margin-top:12px">
          <b>Relación con insumos:</b> los paquetes toman el precio vigente del catálogo de insumos. Si un insumo cambia de precio o se elimina, el total del paquete se recalcula al consultar o editar el paquete. Las cotizaciones ya guardadas conservan su histórico.
        </div>

        <div style="overflow:auto; margin-top:12px">
          <table class="table" id="paquetesCatalogTable">
            <thead>
              <tr>
                <th>Paquete</th>
                <th>Descripción</th>
                <th>Insumos</th>
                <th>Subtotal</th>
                <th>IVA</th>
                <th>Total</th>
                <th>Estatus</th>
                <th style="text-align:right">Acciones</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </div>
  `;

    $('#paquetesSearch')?.addEventListener('input', debounce((e) => {
        state.paquetes.search = e.target.value || '';
        renderPaquetesTable();
    }, 120));

    $('#btnNewPaquete')?.addEventListener('click', () => openPaqueteEditor());
    renderPaquetesTable();
    window.lucide?.createIcons();
}

function renderPaquetesTable() {
    const tbody = $('#paquetesCatalogTable tbody');
    if (!tbody) {
        return;
    }

    const q = String(state.paquetes.search || '').trim().toLowerCase();
    const items = (state.paquetes.items || [])
            .slice()
            .sort((a, b) => Number(b.activo !== false) - Number(a.activo !== false) || String(a.nombre || a.label || '').localeCompare(String(b.nombre || b.label || ''), 'es'))
            .filter(pkg => {
                if (!q) return true;
                const haystack = [pkg.nombre, pkg.label, pkg.descripcion, pkg.description, pkg.badge, pkg.activo === false ? 'inactivo' : 'activo'].join(' ').toLowerCase();
                return haystack.includes(q);
            });

    tbody.innerHTML = items.length ? items.map(pkg => {
        const totals = calcPackageTotal(pkg);
        const count = getPackageItems(pkg).length;
        return `
      <tr>
        <td><b>${escapeHtml(pkg.nombre || pkg.label || 'Paquete')}</b><div class="help">${escapeHtml(pkg.badge || 'Paquete')}</div></td>
        <td>${escapeHtml(pkg.descripcion || pkg.description || '—')}</td>
        <td>${count}</td>
        <td>${formatCurrencyMXN(totals.subtotal)}</td>
        <td>${formatCurrencyMXN(totals.impuestos)}</td>
        <td><b>${formatCurrencyMXN(totals.total)}</b></td>
        <td><span class="badge ${pkg.activo === false ? 'badge--warn' : 'badge--success'}">${pkg.activo === false ? 'Inactivo' : 'Activo'}</span></td>
        <td>
          <div class="actions">
            <button class="btn" data-paq-ver="${escapeAttr(String(pkg.id))}"><i data-lucide="eye"></i>Ver</button>
            <button class="btn" data-paq-edit="${escapeAttr(String(pkg.id))}"><i data-lucide="edit-3"></i>Editar</button>
            <button class="btn btn--danger" data-paq-del="${escapeAttr(String(pkg.id))}"><i data-lucide="trash-2"></i>Eliminar</button>
          </div>
        </td>
      </tr>
    `;
    }).join('') : `<tr><td colspan="8" style="color:var(--muted)">No se encontraron paquetes.</td></tr>`;

    tbody.querySelectorAll('[data-paq-ver]').forEach(btn => {
        btn.addEventListener('click', () => openPaqueteViewer(findPaqueteById(btn.dataset.paqVer)));
    });
    tbody.querySelectorAll('[data-paq-edit]').forEach(btn => {
        btn.addEventListener('click', () => openPaqueteEditor(findPaqueteById(btn.dataset.paqEdit)));
    });
    tbody.querySelectorAll('[data-paq-del]').forEach(btn => {
        btn.addEventListener('click', () => openPaqueteDelete(findPaqueteById(btn.dataset.paqDel)));
    });

    window.lucide?.createIcons();
}

function findPaqueteById(id) {
    return (state.paquetes.items || []).find(it => String(it.id) === String(id));
}

function getPackageItems(pkg = {}) {
    return Array.isArray(pkg.items) ? pkg.items : (Array.isArray(pkg.insumos) ? pkg.insumos : []);
}

function findCatalogForPackageItem(item = {}) {
    const id = item.insumoId ?? item.catalogId ?? item.insumo_id ?? null;
    const code = String(item.codigo || '').toUpperCase();
    return INSUMO_CATALOG.find(it => (id != null && String(it.id) === String(id)) || (code && it.codigo === code));
}

function normalizePackageItemWithCatalog(item = {}) {
    const cat = findCatalogForPackageItem(item);
    const cantidad = Math.max(0.01, Number(item.cantidad || 1));
    return {
        id: item.id ?? null,
        insumoId: cat?.id ?? item.insumoId ?? item.catalogId ?? null,
        catalogId: cat?.id ?? item.catalogId ?? item.insumoId ?? null,
        codigo: cat?.codigo ?? String(item.codigo || '').toUpperCase(),
        descripcion: cat?.descripcion ?? item.descripcion ?? 'Insumo no disponible',
        unidad: cat?.unidad ?? item.unidad ?? 'UD',
        precio: Number(cat?.precio ?? item.precio ?? 0),
        impuestoPct: Number(cat?.impuestoPct ?? item.impuestoPct ?? 0.16),
        cantidad,
        activo: cat ? cat.activo !== false : item.activo !== false,
    };
}

function calcPackageTotal(pkg = {}) {
    const items = getPackageItems(pkg).map(normalizePackageItemWithCatalog).filter(it => it.activo !== false);
    const subtotal = items.reduce((acc, it) => acc + (Number(it.cantidad || 0) * Number(it.precio || 0)), 0);
    const impuestos = items.reduce((acc, it) => acc + (Number(it.cantidad || 0) * Number(it.precio || 0) * Number(it.impuestoPct ?? 0.16)), 0);
    return {subtotal, impuestos, total: subtotal + impuestos};
}

function openPaqueteViewer(pkg) {
    if (!pkg) {
        return;
    }

    const items = getPackageItems(pkg).map(normalizePackageItemWithCatalog);
    const totals = calcPackageTotal(pkg);
    openModal({
        title: 'Detalle de paquete',
        subtitle: pkg.nombre || pkg.label || 'Paquete',
        bodyHtml: `
      <div class="grid cols-3">
        <div class="kpi"><div class="kpi__label">Insumos</div><div class="kpi__value">${items.length}</div></div>
        <div class="kpi"><div class="kpi__label">Subtotal</div><div class="kpi__value">${formatCurrencyMXN(totals.subtotal)}</div></div>
        <div class="kpi"><div class="kpi__label">Total</div><div class="kpi__value">${formatCurrencyMXN(totals.total)}</div></div>
      </div>
      <div class="card card--flat" style="margin-top:12px; box-shadow:none">
        <div class="card__subtitle">Descripción</div>
        <div class="help">${escapeHtml(pkg.descripcion || pkg.description || 'Sin descripción.')}</div>
      </div>
      <div style="overflow:auto; margin-top:12px">
        <table class="table">
          <thead><tr><th>Código</th><th>Descripción</th><th>Cantidad</th><th>Unidad</th><th>Precio</th><th>Total</th></tr></thead>
          <tbody>
            ${items.length ? items.map(it => `
              <tr>
                <td><b>${escapeHtml(it.codigo || '—')}</b></td>
                <td>${escapeHtml(it.descripcion || '—')}</td>
                <td>${formatNumber(it.cantidad || 0)}</td>
                <td>${escapeHtml(it.unidad || 'UD')}</td>
                <td>${formatCurrencyMXN(it.precio || 0)}</td>
                <td>${formatCurrencyMXN(Number(it.cantidad || 0) * Number(it.precio || 0))}</td>
              </tr>
            `).join('') : '<tr><td colspan="6" style="color:var(--muted)">Sin insumos asociados.</td></tr>'}
          </tbody>
        </table>
      </div>
    `,
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cerrar</button>
      <button class="btn btn--primary" id="btnViewerEditPaquete"><i data-lucide="edit-3"></i>Editar</button>
    `
    });
    $('#btnViewerEditPaquete')?.addEventListener('click', () => openPaqueteEditor(pkg));
}

function normalizePaquetePayload(raw = {}) {
    const nombre = String(raw.nombre || raw.label || '').trim();
    const descripcion = String(raw.descripcion || raw.description || '').trim();
    const badge = String(raw.badge || 'Paquete').trim() || 'Paquete';
    const activo = raw.activo !== false && String(raw.activo || 'true') !== 'false';
    const observaciones = String(raw.observaciones || '').trim();
    const items = Array.isArray(raw.items || raw.insumos) ? (raw.items || raw.insumos).map(normalizePackageItemWithCatalog).filter(it => it.insumoId || it.codigo) : [];
    return {nombre, label: nombre, descripcion, description: descripcion, badge, activo, observaciones, items};
}

function validatePaquetePayload(payload) {
    if (!payload.nombre) {
        return 'El nombre del paquete es obligatorio.';
    }
    if (!payload.items.length) {
        return 'El paquete debe tener al menos un insumo asociado.';
    }
    const invalid = payload.items.find(it => !Number.isFinite(Number(it.cantidad)) || Number(it.cantidad) <= 0);
    if (invalid) {
        return 'Todas las cantidades deben ser mayores a cero.';
    }
    return '';
}

function openPaqueteEditor(pkg = null) {
    const isEdit = Boolean(pkg?.id);
    let draftItems = getPackageItems(pkg || {}).map(normalizePackageItemWithCatalog);

    const renderEditorBody = () => {
        const totals = calcPackageTotal({items: draftItems});
        return `
      <div class="grid cols-2">
        <div class="field">
          <label>Nombre del paquete *</label>
          <input id="paqFormNombre" value="${escapeAttr(pkg?.nombre || pkg?.label || '')}" placeholder="Ej. Paquete residencial estándar" />
        </div>
        <div class="field">
          <label>Etiqueta</label>
          <input id="paqFormBadge" value="${escapeAttr(pkg?.badge || 'Paquete')}" placeholder="Ej. Básico, Balanceado, Premium" />
        </div>
        <div class="field" style="grid-column:1 / -1">
          <label>Descripción</label>
          <input id="paqFormDescripcion" value="${escapeAttr(pkg?.descripcion || pkg?.description || '')}" placeholder="Descripción comercial del paquete" />
        </div>
        <div class="field">
          <label>Estatus</label>
          <select id="paqFormActivo">
            <option value="true" ${pkg?.activo === false ? '' : 'selected'}>Activo</option>
            <option value="false" ${pkg?.activo === false ? 'selected' : ''}>Inactivo</option>
          </select>
        </div>
        <div class="field">
          <label>Observaciones</label>
          <input id="paqFormObs" value="${escapeAttr(pkg?.observaciones || '')}" placeholder="Notas internas" />
        </div>
      </div>

      <div class="card card--flat" style="margin-top:12px; box-shadow:none">
        <div class="row" style="align-items:end">
          <div class="field" style="flex:1">
            <label>Agregar insumo</label>
            <select id="paqAddInsumo">
              <option value="">Selecciona un insumo activo</option>
              ${INSUMO_CATALOG.map(it => `<option value="${escapeAttr(String(it.id ?? it.codigo))}">${escapeHtml(it.codigo)} · ${escapeHtml(it.descripcion)} · ${formatCurrencyMXN(it.precio || 0)}</option>`).join('')}
            </select>
          </div>
          <button class="btn" id="btnPaqAddInsumo" type="button"><i data-lucide="plus"></i>Agregar</button>
        </div>
        <div style="overflow:auto; margin-top:12px">
          <table class="table" id="paqDraftTable">
            <thead><tr><th>Código</th><th>Descripción</th><th>Cantidad</th><th>Unidad</th><th>Precio vigente</th><th>Total</th><th></th></tr></thead>
            <tbody>
              ${draftItems.length ? draftItems.map((it, idx) => `
                <tr>
                  <td><b>${escapeHtml(it.codigo || '—')}</b></td>
                  <td>${escapeHtml(it.descripcion || '—')}</td>
                  <td><input data-paq-qty="${idx}" type="number" min="0.01" step="0.01" value="${escapeAttr(String(it.cantidad || 1))}" style="width:92px" /></td>
                  <td>${escapeHtml(it.unidad || 'UD')}</td>
                  <td>${formatCurrencyMXN(it.precio || 0)}</td>
                  <td>${formatCurrencyMXN(Number(it.cantidad || 0) * Number(it.precio || 0))}</td>
                  <td><button class="btn btn--danger" data-paq-remove="${idx}" type="button"><i data-lucide="trash-2"></i></button></td>
                </tr>
              `).join('') : '<tr><td colspan="7" style="color:var(--muted)">Agrega al menos un insumo para guardar el paquete.</td></tr>'}
            </tbody>
          </table>
        </div>
        <div class="quote-totals" style="margin-top:12px">
          <div><span>Subtotal</span><b>${formatCurrencyMXN(totals.subtotal)}</b></div>
          <div><span>IVA</span><b>${formatCurrencyMXN(totals.impuestos)}</b></div>
          <div><span>Total</span><b>${formatCurrencyMXN(totals.total)}</b></div>
        </div>
      </div>
      <div id="paqFormMsg" class="help" style="margin-top:10px"></div>
    `;
    };

    const bindEditorEvents = () => {
        $('#btnPaqAddInsumo')?.addEventListener('click', () => {
            const selected = $('#paqAddInsumo')?.value || '';
            if (!selected) {
                toast({title: 'Selecciona un insumo', message: 'Debe elegirse un insumo del catálogo.', icon: 'alert-triangle'});
                return;
            }
            const catalog = INSUMO_CATALOG.find(it => String(it.id ?? it.codigo) === String(selected));
            if (!catalog) {
                return;
            }
            const exists = draftItems.find(it => String(it.insumoId ?? it.codigo) === String(catalog.id ?? catalog.codigo));
            if (exists) {
                exists.cantidad = Number(exists.cantidad || 1) + 1;
            } else {
                draftItems.push(normalizePackageItemWithCatalog({insumoId: catalog.id, codigo: catalog.codigo, cantidad: 1}));
            }
            refreshEditor();
        });
        $$('#paqDraftTable [data-paq-qty]').forEach(input => {
            input.addEventListener('input', debounce((e) => {
                const idx = Number(e.target.dataset.paqQty);
                if (!draftItems[idx]) return;
                draftItems[idx].cantidad = Math.max(0.01, Number(e.target.value || 0));
                refreshEditor();
            }, 180));
        });
        $$('#paqDraftTable [data-paq-remove]').forEach(btn => {
            btn.addEventListener('click', () => {
                draftItems.splice(Number(btn.dataset.paqRemove), 1);
                refreshEditor();
            });
        });
        window.lucide?.createIcons();
    };

    const readPayload = () => normalizePaquetePayload({
        nombre: $('#paqFormNombre')?.value || '',
        descripcion: $('#paqFormDescripcion')?.value || '',
        badge: $('#paqFormBadge')?.value || '',
        activo: String($('#paqFormActivo')?.value || 'true') === 'true',
        observaciones: $('#paqFormObs')?.value || '',
        items: draftItems,
    });

    const refreshEditor = () => {
        const panel = $('#modalBody');
        if (!panel) return;
        const previous = readPayload();
        pkg = {...(pkg || {}), ...previous};
        draftItems = draftItems.map(normalizePackageItemWithCatalog);
        panel.innerHTML = renderEditorBody();
        bindEditorEvents();
    };

    openModal({
        title: isEdit ? 'Editar paquete' : 'Nuevo paquete',
        subtitle: isEdit ? 'Modifica datos e insumos del paquete' : 'Captura los datos del paquete',
        bodyHtml: renderEditorBody(),
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cancelar</button>
      <button class="btn btn--primary" id="btnSavePaqueteForm"><i data-lucide="save"></i>${isEdit ? 'Guardar cambios' : 'Guardar paquete'}</button>
    `
    });

    bindEditorEvents();
    $('#btnSavePaqueteForm')?.addEventListener('click', () => {
        const msg = $('#paqFormMsg');
        const payload = readPayload();
        const validation = validatePaquetePayload(payload);
        if (validation) {
            if (msg) msg.textContent = validation;
            toast({title: 'Validación', message: validation, icon: 'alert-triangle'});
            return;
        }

        try {
            setPillStatus('Guardando paquete', 'busy');
            if (isEdit) {
                updatePaquete(pkg.id, payload);
                toast({title: 'Paquete actualizado', message: `${payload.nombre} se actualizó correctamente.`, icon: 'save'});
            } else {
                savePaquete(payload);
                toast({title: 'Paquete creado', message: `${payload.nombre} se registró correctamente.`, icon: 'check-circle'});
            }
            closeModal();
            loadPackageCatalogSafe(false);
            renderPaquetesRoute();
            if (state.route === 'cotizador') {
                renderCotizadorRoute();
            }
            setPillStatus('Listo', 'ok');
        } catch (e) {
            console.error(e);
            const message = e?.message || 'No se pudo guardar el paquete.';
            if (msg) msg.textContent = message;
            toast({title: 'Error al guardar', message, icon: 'x-circle'});
            setPillStatus('Error', 'error');
        }
    });
}

function openPaqueteDelete(pkg) {
    if (!pkg) {
        return;
    }
    openModal({
        title: 'Eliminar paquete',
        subtitle: pkg.nombre || pkg.label || '',
        bodyHtml: `
      <div class="help">Se eliminará el paquete <b>${escapeHtml(pkg.nombre || pkg.label || 'Paquete')}</b> del catálogo visible.</div>
      <div class="review-ok" style="margin-top:12px">
        Si el paquete ya fue utilizado en cotizaciones, el sistema conservará el historial de esas cotizaciones y solo retirará el paquete de nuevas operaciones.
      </div>
    `,
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cancelar</button>
      <button class="btn btn--danger" id="btnConfirmDeletePaquete"><i data-lucide="trash-2"></i>Eliminar</button>
    `
    });

    $('#btnConfirmDeletePaquete')?.addEventListener('click', () => {
        try {
            setPillStatus('Eliminando paquete', 'busy');
            removePaquete(pkg.id);
            closeModal();
            loadPackageCatalogSafe(false);
            renderPaquetesRoute();
            if (state.route === 'cotizador') {
                renderCotizadorRoute();
            }
            toast({title: 'Paquete eliminado', message: `${pkg.nombre || pkg.label || 'Paquete'} se retiró del catálogo.`, icon: 'trash-2'});
            setPillStatus('Listo', 'ok');
        } catch (e) {
            console.error(e);
            toast({title: 'No se pudo eliminar', message: e?.message || 'Intenta de nuevo.', icon: 'x-circle'});
            setPillStatus('Error', 'error');
        }
    });
}

// ------------------------------
// Insumos (Catálogo global)
// ------------------------------

function loadInsumoCatalogSafe(silent = false) {
    try {
        const items = getInsumos();
        state.insumos.items = Array.isArray(items) ? items : [];
        setInsumoCatalog(state.insumos.items);
        syncCurrentQuoteInsumosWithCatalog();
        return state.insumos.items;
    } catch (e) {
        console.warn('No se pudo cargar el catálogo de insumos.', e);
        if (!state.insumos.items.length) {
            state.insumos.items = INSUMO_CATALOG.map(it => ({...it, activo: it.activo !== false}));
        }
        if (!silent) {
            toast({title: 'Catálogo no disponible', message: e?.message || 'No se pudo consultar la base de datos.', icon: 'alert-triangle'});
        }
        return state.insumos.items;
    }
}

function syncCurrentQuoteInsumosWithCatalog() {
    const insumos = state.receipt?.insumos;
    if (!Array.isArray(insumos) || !insumos.length) {
        return;
    }

    const byCode = new Map(INSUMO_CATALOG.map(item => [String(item.codigo || '').toUpperCase(), item]));
    let changed = false;

    state.receipt.insumos = insumos.map(item => {
        const code = String(item.codigo || '').toUpperCase();
        const catalog = byCode.get(code);
        if (!catalog) {
            return item;
        }
        changed = true;
        return {
            ...item,
            catalogId: catalog.id ?? item.catalogId ?? null,
            codigo: catalog.codigo,
            descripcion: catalog.descripcion,
            unidad: catalog.unidad,
            precio: Number(catalog.precio || 0),
        };
    });

    if (changed) {
        state.quote = null;
    }
}

function renderInsumosRoute() {
    const root = $('#route-insumos');
    if (!root) {
        return;
    }

    const items = loadInsumoCatalogSafe(true);
    const active = items.filter(it => it.activo !== false).length;
    const inactive = items.length - active;
    const totalCatalog = items.reduce((acc, it) => acc + Number(it.precio || 0), 0);

    root.innerHTML = `
    <div class="grid" style="gap:14px">
      <div class="grid cols-3">
        <div class="card"><div class="card__title">${items.length}</div><div class="card__subtitle">Insumos registrados</div></div>
        <div class="card"><div class="card__title">${active}</div><div class="card__subtitle">Activos</div></div>
        <div class="card"><div class="card__title">${formatCurrencyMXN(totalCatalog)}</div><div class="card__subtitle">Suma de precios unitarios</div></div>
      </div>

      <div class="card">
        <div class="row" style="align-items:center">
          <div style="flex:1">
            <div class="card__title">Gestión de insumos</div>
            <div class="card__subtitle">Catálogo central para paquetes y cotizaciones</div>
          </div>
          <button class="btn btn--primary" id="btnNewInsumo"><i data-lucide="plus"></i>Crear insumo</button>
          <div class="field" style="max-width:340px">
            <label>Buscar</label>
            <input id="insumosSearch" placeholder="Código, descripción, categoría o estatus" value="${escapeAttr(state.insumos.search || '')}" />
          </div>
        </div>

        <div class="review-ok" style="margin-top:12px">
          <b>Relación con paquetes:</b> los paquetes usarán este catálogo como fuente de precio. Si se cambia el precio de un insumo, el catálogo visible para paquetes y nuevas cotizaciones se actualiza de inmediato. Las cotizaciones ya guardadas conservan su información histórica.
        </div>

        <div style="overflow:auto; margin-top:12px">
          <table class="table" id="insumosCatalogTable">
            <thead>
              <tr>
                <th>Código</th>
                <th>Descripción</th>
                <th>Categoría</th>
                <th>Unidad</th>
                <th>Precio</th>
                <th>IVA</th>
                <th>Estatus</th>
                <th>Actualizado</th>
                <th style="text-align:right">Acciones</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </div>
  `;

    $('#insumosSearch')?.addEventListener('input', debounce((e) => {
        state.insumos.search = e.target.value || '';
        renderInsumosTable();
    }, 120));

    $('#btnNewInsumo')?.addEventListener('click', () => openInsumoEditor());
    renderInsumosTable();
    window.lucide?.createIcons();
}

function renderInsumosTable() {
    const tbody = $('#insumosCatalogTable tbody');
    if (!tbody) {
        return;
    }

    const q = String(state.insumos.search || '').trim().toLowerCase();
    const items = (state.insumos.items || [])
            .slice()
            .sort((a, b) => Number(b.activo !== false) - Number(a.activo !== false) || String(a.descripcion || '').localeCompare(String(b.descripcion || ''), 'es'))
            .filter(it => {
                if (!q) {
                    return true;
                }
                const haystack = [it.codigo, it.descripcion, it.categoria, it.unidad, it.activo === false ? 'inactivo' : 'activo'].join(' ').toLowerCase();
                return haystack.includes(q);
            });

    tbody.innerHTML = items.length ? items.map(it => {
        const ivaPct = Math.round(Number(it.impuestoPct ?? 0.16) * 1000) / 10;
        return `
      <tr>
        <td><b>${escapeHtml(it.codigo || '—')}</b></td>
        <td>${escapeHtml(it.descripcion || '—')}</td>
        <td>${escapeHtml(it.categoria || 'General')}</td>
        <td>${escapeHtml(it.unidad || 'UD')}</td>
        <td>${formatCurrencyMXN(it.precio || 0)}</td>
        <td>${ivaPct}%</td>
        <td><span class="badge ${it.activo === false ? 'badge--warn' : 'badge--success'}">${it.activo === false ? 'Inactivo' : 'Activo'}</span></td>
        <td>${it.updatedAt ? formatDateTime(it.updatedAt) : '—'}</td>
        <td>
          <div class="actions">
            <button class="btn" data-ins-ver="${escapeAttr(String(it.id))}"><i data-lucide="eye"></i>Ver</button>
            <button class="btn" data-ins-edit="${escapeAttr(String(it.id))}"><i data-lucide="edit-3"></i>Editar</button>
            <button class="btn btn--danger" data-ins-del="${escapeAttr(String(it.id))}"><i data-lucide="trash-2"></i>Eliminar</button>
          </div>
        </td>
      </tr>
    `;
    }).join('') : `<tr><td colspan="9" style="color:var(--muted)">No se encontraron insumos.</td></tr>`;

    tbody.querySelectorAll('[data-ins-ver]').forEach(btn => {
        btn.addEventListener('click', () => openInsumoViewer(findInsumoById(btn.dataset.insVer)));
    });
    tbody.querySelectorAll('[data-ins-edit]').forEach(btn => {
        btn.addEventListener('click', () => openInsumoEditor(findInsumoById(btn.dataset.insEdit)));
    });
    tbody.querySelectorAll('[data-ins-del]').forEach(btn => {
        btn.addEventListener('click', () => openInsumoDelete(findInsumoById(btn.dataset.insDel)));
    });

    window.lucide?.createIcons();
}

function findInsumoById(id) {
    return (state.insumos.items || []).find(it => String(it.id) === String(id));
}

function normalizeInsumoCatalogPayload(raw = {}) {
    const code = String(raw.codigo || '').trim().toUpperCase();
    const descripcion = String(raw.descripcion || '').trim();
    const categoria = String(raw.categoria || 'General').trim() || 'General';
    const unidad = String(raw.unidad || 'UD').trim().toUpperCase() || 'UD';
    const precio = Math.max(0, Number(raw.precio || 0));
    const impuestoInput = Number(raw.impuestoPct ?? 0.16);
    const impuestoPct = impuestoInput > 1 ? impuestoInput / 100 : impuestoInput;
    const activo = raw.activo !== false && String(raw.activo || 'true') !== 'false';
    const observaciones = String(raw.observaciones || '').trim();

    return {codigo: code, descripcion, categoria, unidad, precio, impuestoPct, activo, observaciones};
}

function readInsumoForm() {
    return normalizeInsumoCatalogPayload({
        codigo: $('#insFormCodigo')?.value || '',
        descripcion: $('#insFormDescripcion')?.value || '',
        categoria: $('#insFormCategoria')?.value || '',
        unidad: $('#insFormUnidad')?.value || '',
        precio: Number($('#insFormPrecio')?.value || 0),
        impuestoPct: Number($('#insFormIva')?.value || 16) / 100,
        activo: String($('#insFormActivo')?.value || 'true') === 'true',
        observaciones: $('#insFormObs')?.value || '',
    });
}

function validateInsumoPayload(payload) {
    if (!payload.codigo) {
        return 'El código del insumo es obligatorio.';
    }
    if (!payload.descripcion) {
        return 'La descripción del insumo es obligatoria.';
    }
    if (!payload.unidad) {
        return 'La unidad de medida es obligatoria.';
    }
    if (!Number.isFinite(Number(payload.precio)) || Number(payload.precio) < 0) {
        return 'El precio unitario debe ser mayor o igual a cero.';
    }
    if (!Number.isFinite(Number(payload.impuestoPct)) || Number(payload.impuestoPct) < 0 || Number(payload.impuestoPct) > 0.30) {
        return 'El IVA debe estar entre 0% y 30%.';
    }
    return '';
}

function openInsumoViewer(item) {
    if (!item) {
        return;
    }
    const ivaPct = Math.round(Number(item.impuestoPct ?? 0.16) * 1000) / 10;
    openModal({
        title: 'Detalle de insumo',
        subtitle: item.codigo || '',
        bodyHtml: `
      <div class="grid cols-2">
        <div class="kpi"><div class="kpi__label">Descripción</div><div class="kpi__value" style="font-size:13px">${escapeHtml(item.descripcion || '—')}</div></div>
        <div class="kpi"><div class="kpi__label">Precio unitario</div><div class="kpi__value">${formatCurrencyMXN(item.precio || 0)}</div></div>
        <div class="kpi"><div class="kpi__label">Categoría</div><div class="kpi__value" style="font-size:13px">${escapeHtml(item.categoria || 'General')}</div></div>
        <div class="kpi"><div class="kpi__label">Unidad</div><div class="kpi__value" style="font-size:13px">${escapeHtml(item.unidad || 'UD')}</div></div>
        <div class="kpi"><div class="kpi__label">IVA</div><div class="kpi__value" style="font-size:13px">${ivaPct}%</div></div>
        <div class="kpi"><div class="kpi__label">Estatus</div><div class="kpi__value" style="font-size:13px">${item.activo === false ? 'Inactivo' : 'Activo'}</div></div>
      </div>
      <div class="card card--flat" style="margin-top:12px; box-shadow:none">
        <div class="card__subtitle">Observaciones</div>
        <div class="help">${escapeHtml(item.observaciones || 'Sin observaciones.')}</div>
      </div>
    `,
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cerrar</button>
      <button class="btn btn--primary" id="btnViewerEditInsumo"><i data-lucide="edit-3"></i>Editar</button>
    `
    });
    $('#btnViewerEditInsumo')?.addEventListener('click', () => openInsumoEditor(item));
}

function openInsumoEditor(item = null) {
    const isEdit = Boolean(item?.id);
    const iva = Math.round(Number(item?.impuestoPct ?? 0.16) * 1000) / 10;
    openModal({
        title: isEdit ? 'Editar insumo' : 'Nuevo insumo',
        subtitle: isEdit ? 'Modifica los datos del catálogo' : 'Captura los datos del insumo',
        bodyHtml: `
      <div class="grid cols-2">
        <div class="field">
          <label>Código *</label>
          <input id="insFormCodigo" value="${escapeAttr(item?.codigo || '')}" placeholder="Ej. PANEL-550" />
        </div>
        <div class="field">
          <label>Categoría</label>
          <input id="insFormCategoria" value="${escapeAttr(item?.categoria || 'General')}" placeholder="Ej. Paneles, Inversores, Cableado" />
        </div>
        <div class="field" style="grid-column:1 / -1">
          <label>Descripción *</label>
          <input id="insFormDescripcion" value="${escapeAttr(item?.descripcion || '')}" placeholder="Nombre o descripción del insumo" />
        </div>
        <div class="field">
          <label>Unidad *</label>
          <select id="insFormUnidad">
            ${['UD', 'PZA', 'M', 'W', 'SERV', 'KIT'].map(u => `<option value="${u}" ${u === String(item?.unidad || 'UD').toUpperCase() ? 'selected' : ''}>${u}</option>`).join('')}
          </select>
        </div>
        <div class="field">
          <label>Precio unitario *</label>
          <input id="insFormPrecio" type="number" min="0" step="0.01" value="${escapeAttr(String(item?.precio ?? 0))}" />
        </div>
        <div class="field">
          <label>IVA (%)</label>
          <input id="insFormIva" type="number" min="0" max="30" step="0.5" value="${escapeAttr(String(iva))}" />
        </div>
        <div class="field">
          <label>Estatus</label>
          <select id="insFormActivo">
            <option value="true" ${item?.activo === false ? '' : 'selected'}>Activo</option>
            <option value="false" ${item?.activo === false ? 'selected' : ''}>Inactivo</option>
          </select>
        </div>
        <div class="field" style="grid-column:1 / -1">
          <label>Observaciones</label>
          <textarea id="insFormObs" rows="3" placeholder="Notas internas del insumo">${escapeHtml(item?.observaciones || '')}</textarea>
        </div>
      </div>
      <div id="insFormMsg" class="help" style="margin-top:10px"></div>
    `,
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cancelar</button>
      <button class="btn btn--primary" id="btnSaveInsumoForm"><i data-lucide="save"></i>${isEdit ? 'Guardar cambios' : 'Guardar insumo'}</button>
    `
    });

    $('#btnSaveInsumoForm')?.addEventListener('click', () => {
        const msg = $('#insFormMsg');
        const payload = readInsumoForm();
        const validation = validateInsumoPayload(payload);
        if (validation) {
            if (msg) msg.textContent = validation;
            toast({title: 'Validación', message: validation, icon: 'alert-triangle'});
            return;
        }

        try {
            setPillStatus('Guardando insumo', 'busy');
            if (isEdit) {
                updateInsumo(item.id, payload);
                toast({title: 'Insumo actualizado', message: `${payload.codigo} se actualizó correctamente.`, icon: 'save'});
            } else {
                saveInsumo(payload);
                toast({title: 'Insumo creado', message: `${payload.codigo} se registró correctamente.`, icon: 'check-circle'});
            }
            closeModal();
            loadInsumoCatalogSafe(false);
            loadPackageCatalogSafe(false);
            renderInsumosRoute();
            if (state.route === 'paquetes') {
                renderPaquetesRoute();
            }
            if (state.route === 'cotizador') {
                renderCotizadorRoute();
            }
            setPillStatus('Listo', 'ok');
        } catch (e) {
            console.error(e);
            const message = e?.message || 'No se pudo guardar el insumo.';
            if (msg) msg.textContent = message;
            toast({title: 'Error al guardar', message, icon: 'x-circle'});
            setPillStatus('Error', 'error');
        }
    });
}

function openInsumoDelete(item) {
    if (!item) {
        return;
    }
    openModal({
        title: 'Eliminar insumo',
        subtitle: item.codigo || '',
        bodyHtml: `
      <div class="help">Se eliminará el insumo <b>${escapeHtml(item.descripcion || item.codigo)}</b> del catálogo.</div>
      <div class="review-ok" style="margin-top:12px">
        Si posteriormente un paquete utiliza este insumo, el sistema deberá advertirlo y recalcular sus totales. Las cotizaciones ya guardadas no se modificarán.
      </div>
    `,
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cancelar</button>
      <button class="btn btn--danger" id="btnConfirmDeleteInsumo"><i data-lucide="trash-2"></i>Eliminar</button>
    `
    });

    $('#btnConfirmDeleteInsumo')?.addEventListener('click', () => {
        try {
            setPillStatus('Eliminando insumo', 'busy');
            removeInsumo(item.id);
            closeModal();
            loadInsumoCatalogSafe(false);
            loadPackageCatalogSafe(false);
            renderInsumosRoute();
            if (state.route === 'paquetes') {
                renderPaquetesRoute();
            }
            if (state.route === 'cotizador') {
                renderCotizadorRoute();
            }
            toast({title: 'Insumo eliminado', message: `${item.codigo} se retiró del catálogo.`, icon: 'trash-2'});
            setPillStatus('Listo', 'ok');
        } catch (e) {
            console.error(e);
            toast({title: 'No se pudo eliminar', message: e?.message || 'Intenta de nuevo.', icon: 'x-circle'});
            setPillStatus('Error', 'error');
        }
    });
}

// ------------------------------
// Proyectos
// ------------------------------

function renderProyectosRoute() {
    const root = $('#route-proyectos');
    const projects = getProjects();
    const completed = projects.filter(p => String(p.status || '').toLowerCase().includes('complet')).length;
    const inProgress = projects.filter(p => ['en planeación', 'en instalación', 'en trámite'].includes(String(p.status || '').toLowerCase())).length;

    root.innerHTML = `
    <div class="grid" style="gap:14px">
      <div class="grid cols-3">
        <div class="card"><div class="card__title">${projects.length}</div><div class="card__subtitle">Proyectos registrados</div></div>
        <div class="card"><div class="card__title">${inProgress}</div><div class="card__subtitle">En proceso</div></div>
        <div class="card"><div class="card__title">${completed}</div><div class="card__subtitle">Completados</div></div>
      </div>

      <div class="card">
        <div class="row" style="align-items:center">
          <div style="flex:1">
            <div class="card__title">Proyectos</div>
            <div class="card__subtitle">Seguimiento comercial y técnico de cotizaciones confirmadas</div>
          </div>
          <button class="btn btn--primary" id="btnNewProject"><i data-lucide="plus"></i>Agregar proyecto</button>
          <div class="field" style="max-width:320px">
            <label>Buscar</label>
            <input id="projSearch" placeholder="Cliente, estatus o No. de servicio" />
          </div>
        </div>

        <div style="overflow:auto; margin-top:10px">
          <table class="table" id="proyectosTable">
            <thead>
              <tr>
                <th>ID</th>
                <th>Cliente</th>
                <th>No. de servicio</th>
                <th>Potencia</th>
                <th>Inversión</th>
                <th>Estatus</th>
                <th style="text-align:right">Acciones</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </div>
  `;

    $('#projSearch').addEventListener('input', debounce(renderProyectosTable, 120));
    $('#btnNewProject')?.addEventListener('click', () => openProjectEditor());
    renderProyectosTable();
}

function renderProyectosTable() {
    const tbody = $('#proyectosTable tbody');
    if (!tbody)
        return;

    const term = ($('#projSearch')?.value || '').toLowerCase().trim();
    const projects = getProjects();
    const filtered = term ? projects.filter(p => {
        const c = (p.client?.nombre || p.receipt?.nombre || '').toLowerCase();
        const s = (p.receipt?.servicio || '').toLowerCase();
        const st = (p.status || '').toLowerCase();
        return c.includes(term) || s.includes(term) || st.includes(term) || (p.id || '').toLowerCase().includes(term);
    }) : projects;

    tbody.innerHTML = filtered.length ? filtered.map(p => {
        const badge = 'badge--success';
        return `
      <tr>
        <td>${escapeHtml(p.id)}</td>
        <td>${escapeHtml(p.client?.nombre || p.receipt?.nombre || '-')}</td>
        <td>${escapeHtml(p.receipt?.servicio || '-')}</td>
        <td>${Number(p.quote?.kwp || 0).toFixed(2)} kWp</td>
        <td>${formatCurrencyMXN(p.quote?.inversion || 0)}</td>
        <td><span class="badge ${badge}">${escapeHtml(p.status || '—')}</span></td>
        <td>
          <div class="actions">
            <button class="btn" data-action="view" data-id="${p.id}"><i data-lucide="eye"></i>Ver</button>
            <button class="btn" data-action="edit" data-id="${p.id}"><i data-lucide="pencil"></i>Editar</button>
            <button class="btn" data-action="status" data-id="${p.id}"><i data-lucide="refresh-cw"></i>Estatus</button>
            <button class="btn btn--danger" data-action="delete" data-id="${p.id}"><i data-lucide="trash-2"></i>Eliminar</button>
          </div>
        </td>
      </tr>
    `;
    }).join('') : `
    <tr><td colspan="7" style="color:var(--muted)">Aún no hay proyectos confirmados.</td></tr>
  `;

    window.lucide?.createIcons();

    tbody.querySelectorAll('button[data-action]').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.dataset.id;
            const action = btn.dataset.action;
            const p = getProjects().find(x => x.id === id);
            if (!p)
                return;

            if (action === 'view') {
                openModalForProject(p);
            } else if (action === 'edit') {
                openProjectEditor(p);
            } else if (action === 'status') {
                openStatusPicker(p);
            } else if (action === 'delete') {
                confirmDeleteProject(p);
            }
        });
    });
}

function confirmDeleteProject(p) {
    openModal({
        title: 'Eliminar proyecto',
        subtitle: p.id,
        bodyHtml: `<div class="help">Se eliminará el proyecto <b>${escapeHtml(p.id)}</b>. Esta acción no se puede deshacer.</div>`,
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cancelar</button>
      <button class="btn btn--danger" id="btnDelProj"><i data-lucide="trash-2"></i>Eliminar</button>
    `
    });

    $('#btnDelProj')?.addEventListener('click', () => {
        removeProject(p.id);
        toast({title: 'Proyecto eliminado', message: p.id, icon: 'trash-2'});
        closeModal();
        renderProyectosTable();
    });
}

function openProjectEditor(project = null) {
    const isEdit = Boolean(project);
    const quotes = getQuotes();
    const qOptions = quotes.map(q => {
        const cliente = q.client?.nombre || q.receipt?.nombre || 'Cliente';
        const total = formatCurrencyMXN(q.receipt?.totalAPagar || 0);
        return `<option value="${escapeAttr(q.id)}">${escapeHtml(q.id)} · ${escapeHtml(cliente)} · ${total}</option>`;
    }).join('');

    const p = project || {};
    const preQ = p.quoteId || '';

    const body = `
    <div class="grid cols-2">
      <div class="field" style="grid-column: span 2">
        <label>Cotización asociada (opcional)</label>
        <select id="pfQuote" ${isEdit ? 'disabled' : ''}>
          <option value="">Sin cotización</option>
          ${qOptions}
        </select>
        <div class="help" style="margin-top:6px">Si seleccionas una cotización, los datos se cargarán automáticamente.</div>
      </div>

      <div class="field">
        <label>Cliente</label>
        <input id="pfCliente" value="${escapeAttr(p.client?.nombre || p.receipt?.nombre || '')}" placeholder="Nombre del cliente" />
      </div>
      <div class="field">
        <label>No. de servicio</label>
        <input id="pfServicio" value="${escapeAttr(p.receipt?.servicio || '')}" placeholder="###########" />
      </div>

      <div class="field">
        <label>Potencia (kWp)</label>
        <input id="pfKwp" type="number" step="0.01" min="0" value="${escapeAttr(String(p.quote?.kwp || 0))}" />
      </div>
      <div class="field">
        <label>Inversión (MXN)</label>
        <input id="pfInv" type="number" step="1" min="0" value="${escapeAttr(String(p.quote?.inversion || 0))}" />
      </div>

      <div class="field" style="grid-column: span 2">
        <label>Estatus</label>
        <select id="pfStatus">
          ${['En planeación', 'En instalación', 'En trámite', 'Completado', 'Pausado'].map(s => `<option ${p.status === s ? 'selected' : ''}>${escapeHtml(s)}</option>`).join('')}
        </select>
      </div>

      <div class="field" style="grid-column: span 2">
        <label>Notas</label>
        <textarea id="pfNotes" rows="3" placeholder="Notas del proyecto">${escapeHtml(p.notes || '')}</textarea>
      </div>
    </div>
  `;

    openModal({
        title: isEdit ? 'Editar proyecto' : 'Agregar proyecto',
        subtitle: isEdit ? p.id : 'Completa la información',
        bodyHtml: body,
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cancelar</button>
      <button class="btn btn--primary" id="pfSave"><i data-lucide="save"></i>Guardar</button>
    `
    });

    // Prefill quote selection
    if (!isEdit && preQ) {
        const sel = $('#pfQuote');
        if (sel)
            sel.value = preQ;
    }

    const applyFromQuote = (qid) => {
        const q = quotes.find(x => x.id === qid);
        if (!q)
            return;
        $('#pfCliente').value = q.client?.nombre || q.receipt?.nombre || '';
        $('#pfServicio').value = q.receipt?.servicio || '';
        $('#pfKwp').value = Number(q.quote?.kwp || 0).toFixed(2);
        $('#pfInv').value = Math.round(Number(q.quote?.inversion || 0));
    };

    $('#pfQuote')?.addEventListener('change', (e) => {
        const qid = e.target.value;
        if (qid)
            applyFromQuote(qid);
    });

    // Save
    $('#pfSave')?.addEventListener('click', () => {
        const qid = $('#pfQuote')?.value || '';
        const status = $('#pfStatus').value;
        const notes = $('#pfNotes').value.trim();
        const cliente = $('#pfCliente').value.trim();
        const servicio = $('#pfServicio').value.trim();
        const kwp = Number($('#pfKwp').value || 0);
        const inv = Number($('#pfInv').value || 0);

        if (!cliente) {
            toast({title: 'Falta información', message: 'Captura el nombre del cliente.', icon: 'alert-triangle'});
            return;
        }

        if (isEdit) {
            updateProject(p.id, {
                status,
                notes,
                client: {...(p.client || {}), nombre: cliente},
                receipt: {...(p.receipt || {}), servicio},
                quote: {...(p.quote || {}), kwp, inversion: inv},
            });
            toast({title: 'Proyecto actualizado', message: p.id, icon: 'check-circle'});
            closeModal();
            renderProyectosTable();
            return;
        }

        // Create
        let created;
        if (qid) {
            const q = quotes.find(x => x.id === qid);
            if (!q) {
                toast({title: 'Cotización no encontrada', message: 'Selecciona una cotización válida.', icon: 'alert-triangle'});
                return;
            }
            created = saveProjectFromQuote({...q, status});
            updateProject(created.id, {status, notes});
            updateQuote(q.id, {status: q.status === 'Confirmada' ? 'Confirmada' : q.status});
        } else {
            created = saveProject({
                status,
                notes,
                client: {nombre: cliente},
                receipt: {servicio},
                quote: {kwp, inversion: inv},
            });
        }

        toast({title: 'Proyecto guardado', message: created.id, icon: 'briefcase'});
        closeModal();
        renderProyectosTable();
    });
}

function openModalForProject(p) {
    const r = p.receipt || {};
    const c = p.client || {};

    const body = `
    <div class="grid cols-2">
      <div class="card" style="box-shadow:none">
        <div class="card__title">Cliente</div>
        <div class="help"><b>${escapeHtml(c.nombre || r.nombre || '-')}</b></div>
        <div class="help">${escapeHtml(c.telefono || '-')} · ${escapeHtml(c.email || '-')}</div>
        <div class="help" style="margin-top:8px">${escapeHtml(c.direccion || r.direccion || '-')}</div>
      </div>
      <div class="card" style="box-shadow:none">
        <div class="card__title">Sistema</div>
        <div class="help">Potencia: <b>${Number(p.quote?.kwp || 0).toFixed(2)} kWp</b></div>
        <div class="help">Paneles: <b>${escapeHtml(p.quote?.paneles || '—')}</b></div>
        <div class="help">Inversión: <b>${formatCurrencyMXN(p.quote?.inversion || 0)}</b></div>
        <div class="help">Ahorro mensual: <b>${formatCurrencyMXN(p.quote?.ahorroMensual || 0)}</b></div>
      </div>
    </div>

    <div class="card" style="box-shadow:none; margin-top:12px">
      <div class="card__title">Recibo</div>
      <div class="row">
        <div class="kpi" style="flex:1">
          <div class="kpi__label">No. de servicio</div>
          <div class="kpi__value" style="font-size:13px">${escapeHtml(r.servicio || '-')}</div>
        </div>
        <div class="kpi" style="flex:1">
          <div class="kpi__label">Tarifa</div>
          <div class="kpi__value" style="font-size:13px">${escapeHtml(r.tarifa || '-')}</div>
        </div>
        <div class="kpi" style="flex:1">
          <div class="kpi__label">Total</div>
          <div class="kpi__value" style="font-size:13px">${formatCurrencyMXN(r.totalAPagar || 0)}</div>
        </div>
      </div>
    </div>
  `;

    openModal({
        title: p.id,
        subtitle: `Creado ${formatDateTime(p.createdAt)} · Estatus: ${p.status}`,
        bodyHtml: body,
        footHtml: `
      <button class="btn" data-close="true"><i data-lucide="x"></i>Cerrar</button>
      <button class="btn" id="pStatus"><i data-lucide="refresh-cw"></i>Cambiar estatus</button>
    `
    });

    $('#pStatus')?.addEventListener('click', () => openStatusPicker(p));
}

function openStatusPicker(p) {
    const options = ['En planeación', 'En instalación', 'En trámite', 'Completado', 'Pausado'];
    const body = `
    <div class="help">Selecciona el nuevo estatus para el proyecto <b>${escapeHtml(p.id)}</b>.</div>
    <div class="grid" style="margin-top:12px">
      ${options.map(o => `
        <button class="btn" data-status="${escapeAttr(o)}" style="justify-content:flex-start">
          <i data-lucide="dot"></i>${escapeHtml(o)}
        </button>
      `).join('')}
    </div>
  `;

    openModal({
        title: 'Cambiar estatus',
        subtitle: escapeHtml(p.client?.nombre || p.receipt?.nombre || ''),
        bodyHtml: body,
        footHtml: `<button class="btn" data-close="true"><i data-lucide="x"></i>Cancelar</button>`
    });

    $$('#modalBody button[data-status]').forEach(btn => {
        btn.addEventListener('click', () => {
            const status = btn.dataset.status;
            updateProject(p.id, {status});
            toast({title: 'Estatus actualizado', message: `${p.id} · ${status}`, icon: 'check-circle'});
            closeModal();
            renderProyectosTable();
        });
    });
}


// ------------------------------
// Reportes de cotización
// ------------------------------

function toInputDateLocal(date = new Date()) {
    const d = new Date(date);
    d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    return d.toISOString().slice(0, 10);
}

function firstDayOfCurrentMonthInput() {
    const now = new Date();
    return toInputDateLocal(new Date(now.getFullYear(), now.getMonth(), 1));
}

function ensureReportDefaultDates() {
    if (!state.reportes.fechaInicio) {
        state.reportes.fechaInicio = firstDayOfCurrentMonthInput();
    }
    if (!state.reportes.fechaFin) {
        state.reportes.fechaFin = toInputDateLocal();
    }
}

function reportStatusOptions(selected = 'todos') {
    const options = [
        {value: 'todos', label: 'Todos'},
        {value: 'Guardada', label: 'Guardadas'},
        {value: 'Confirmada', label: 'Confirmadas / proyecto'},
    ];
    return options.map(o => `<option value="${escapeAttr(o.value)}" ${o.value === selected ? 'selected' : ''}>${escapeHtml(o.label)}</option>`).join('');
}

function reportTariffOptions(selected = 'todas') {
    const tariffs = TARIFFS.filter(t => t.kind !== 'cashvolt');
    return `
      <option value="todas" ${selected === 'todas' ? 'selected' : ''}>Todas</option>
      ${tariffs.map(t => `<option value="${escapeAttr(t.label)}" ${t.label === selected ? 'selected' : ''}>${escapeHtml(t.label)}</option>`).join('')}
    `;
}

function renderReportesRoute() {
    const root = $('#route-reportes');
    if (!root) {
        return;
    }

    ensureReportDefaultDates();
    const data = state.reportes.data;

    root.innerHTML = `
    <div class="grid" style="gap:14px">
      <div class="card">
        <div class="row" style="align-items:center">
          <div style="flex:1">
            <div class="card__title">Generar reporte de cotizaciones</div>
            <div class="card__subtitle">Consulta cotizaciones por rango de fechas y genera una vista previa exportable.</div>
          </div>
        </div>

        <div class="grid cols-4" style="margin-top:14px">
          <div class="field">
            <label>Fecha inicial *</label>
            <input id="repFechaInicio" type="date" value="${escapeAttr(state.reportes.fechaInicio)}" />
          </div>
          <div class="field">
            <label>Fecha final *</label>
            <input id="repFechaFin" type="date" value="${escapeAttr(state.reportes.fechaFin)}" />
          </div>
          <div class="field">
            <label>Estatus</label>
            <select id="repStatus">${reportStatusOptions(state.reportes.status)}</select>
          </div>
          <div class="field">
            <label>Tarifa</label>
            <select id="repTarifa">${reportTariffOptions(state.reportes.tarifa)}</select>
          </div>
        </div>

        <div class="wizard-actions" style="justify-content:flex-end; margin-top:12px">
          <button class="btn" id="btnClearReport"><i data-lucide="rotate-ccw"></i>Limpiar</button>
          <button class="btn btn--primary" id="btnGenerateReport"><i data-lucide="file-bar-chart"></i>Generar reporte</button>
        </div>

        <div id="reportValidationMsg" class="help" style="margin-top:8px"></div>
      </div>

      <div id="reportPreviewWrap">
        ${data ? renderReportPreview(data) : renderEmptyReportState()}
      </div>
    </div>
  `;

    $('#repFechaInicio')?.addEventListener('change', (e) => state.reportes.fechaInicio = e.target.value || '');
    $('#repFechaFin')?.addEventListener('change', (e) => state.reportes.fechaFin = e.target.value || '');
    $('#repStatus')?.addEventListener('change', (e) => state.reportes.status = e.target.value || 'todos');
    $('#repTarifa')?.addEventListener('change', (e) => state.reportes.tarifa = e.target.value || 'todas');
    $('#btnGenerateReport')?.addEventListener('click', generateCotizacionesReport);
    $('#btnClearReport')?.addEventListener('click', () => {
        state.reportes = {
            fechaInicio: firstDayOfCurrentMonthInput(),
            fechaFin: toInputDateLocal(),
            status: 'todos',
            tarifa: 'todas',
            data: null,
        };
        renderReportesRoute();
    });

    bindReportPreviewActions();
    window.lucide?.createIcons();
}

function renderEmptyReportState() {
    return `
    <div class="card">
      <div class="empty">
        <div class="empty__icon"><i data-lucide="bar-chart-3"></i></div>
        <div class="card__title">Sin reporte generado</div>
        <div class="help">Selecciona fecha inicial y fecha final para consultar las cotizaciones registradas en ese periodo.</div>
      </div>
    </div>
  `;
}

function validateReportFilters(filters) {
    if (!filters.fechaInicio || !filters.fechaFin) {
        return 'La fecha inicial y la fecha final son obligatorias.';
    }
    if (filters.fechaInicio > filters.fechaFin) {
        return 'La fecha inicial no puede ser mayor que la fecha final.';
    }
    return '';
}

function currentReportFiltersFromForm() {
    return {
        fechaInicio: $('#repFechaInicio')?.value || state.reportes.fechaInicio || '',
        fechaFin: $('#repFechaFin')?.value || state.reportes.fechaFin || '',
        status: $('#repStatus')?.value || state.reportes.status || 'todos',
        tarifa: $('#repTarifa')?.value || state.reportes.tarifa || 'todas',
    };
}

function generateCotizacionesReport() {
    const msg = $('#reportValidationMsg');
    const filters = currentReportFiltersFromForm();
    const validation = validateReportFilters(filters);
    if (validation) {
        if (msg) msg.textContent = validation;
        toast({title: 'Validación', message: validation, icon: 'alert-triangle'});
        return;
    }

    try {
        setPillStatus('Generando reporte', 'busy');
        const data = getCotizacionesReport(filters);
        state.reportes = {...state.reportes, ...filters, data};
        if (msg) msg.textContent = '';
        renderReportesRoute();
        const count = Number(data?.summary?.totalCotizaciones || 0);
        toast({title: 'Reporte generado', message: `${formatNumber(count)} cotización(es) encontradas.`, icon: 'file-check'});
        setPillStatus('Listo', 'ok');
    } catch (e) {
        console.error(e);
        const message = e?.message || 'No se pudo generar el reporte.';
        if (msg) msg.textContent = message;
        toast({title: 'Error al generar reporte', message, icon: 'x-circle'});
        setPillStatus('Error', 'error');
    }
}

function renderReportPreview(data = {}) {
    const summary = data.summary || {};
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const filters = data.filters || state.reportes;
    const hasRows = rows.length > 0;

    return `
    <div class="card" id="reportPreview">
      <div class="row" style="align-items:center; margin-bottom:12px">
        <div style="flex:1">
          <div class="card__title">Reporte de cotizaciones</div>
          <div class="card__subtitle">Periodo: ${escapeHtml(filters.fechaInicio || '—')} al ${escapeHtml(filters.fechaFin || '—')} · Estatus: ${escapeHtml(filters.status && filters.status !== 'todos' ? filters.status : 'Todos')} · Tarifa: ${escapeHtml(filters.tarifa && filters.tarifa !== 'todas' ? filters.tarifa : 'Todas')}</div>
        </div>
        <div class="actions no-print">
          <button class="btn" id="btnReportPdf" ${hasRows ? '' : 'disabled'}><i data-lucide="download"></i>Descargar PDF</button>
          <button class="btn" id="btnReportCsv" ${hasRows ? '' : 'disabled'}><i data-lucide="sheet"></i>Exportar Excel</button>
        </div>
      </div>

      ${hasRows ? `
        <div class="grid cols-4">
          <div class="kpi"><div class="kpi__label">Cotizaciones</div><div class="kpi__value">${formatNumber(summary.totalCotizaciones || 0)}</div></div>
          <div class="kpi"><div class="kpi__label">Monto total cotizado</div><div class="kpi__value">${formatCurrencyMXN(summary.montoTotal || 0)}</div></div>
          <div class="kpi"><div class="kpi__label">Promedio de inversión</div><div class="kpi__value">${formatCurrencyMXN(summary.promedioInversion || 0)}</div></div>
          <div class="kpi"><div class="kpi__label">Cotizaciones confirmadas</div><div class="kpi__value">${formatNumber(summary.confirmadas || 0)}</div></div>
        </div>

        <div class="grid cols-4" style="margin-top:12px">
          <div class="kpi"><div class="kpi__label">Convertidas en proyecto</div><div class="kpi__value">${formatNumber(summary.convertidasProyecto || 0)}</div></div>
          <div class="kpi"><div class="kpi__label">Pendientes / guardadas</div><div class="kpi__value">${formatNumber(summary.pendientes || 0)}</div></div>
          <div class="kpi"><div class="kpi__label">Potencia total</div><div class="kpi__value">${Number(summary.potenciaTotalKwp || 0).toFixed(2)} kWp</div></div>
          <div class="kpi"><div class="kpi__label">Ahorro mensual estimado</div><div class="kpi__value">${formatCurrencyMXN(summary.ahorroMensualTotal || 0)}</div></div>
        </div>

        <div style="overflow:auto; margin-top:14px">
          <table class="table" id="reportQuotesTable">
            <thead>
              <tr>
                <th>Folio</th>
                <th>Fecha</th>
                <th>Cliente</th>
                <th>Tarifa</th>
                <th>Consumo</th>
                <th>Paneles</th>
                <th>Potencia</th>
                <th>Inversión</th>
                <th>Ahorro</th>
                <th>Retorno</th>
                <th>Estatus</th>
                <th style="text-align:right">Acciones</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map((row, idx) => renderReportRow(row, idx)).join('')}
            </tbody>
          </table>
        </div>
      ` : `
        <div class="empty">
          <div class="empty__icon"><i data-lucide="search-x"></i></div>
          <div class="card__title">No se encontraron cotizaciones</div>
          <div class="help">Modifica el rango de fechas o los filtros para generar nuevamente el reporte.</div>
        </div>
      `}
    </div>
  `;
}

function renderReportRow(row = {}, idx = 0) {
    const status = String(row.estatus || 'Guardada');
    const badge = status.toLowerCase().includes('confirm') ? 'badge--success' : 'badge--warn';
    return `
    <tr>
      <td>${escapeHtml(row.folio || row.id || '—')}</td>
      <td>${escapeHtml(row.fechaTexto || row.fecha || '—')}</td>
      <td>${escapeHtml(row.cliente || '—')}</td>
      <td>${escapeHtml(row.tarifa || '—')}</td>
      <td>${formatNumber(row.consumoMensual || 0)} kWh</td>
      <td>${formatNumber(row.paneles || 0)}</td>
      <td>${Number(row.potenciaKwp || 0).toFixed(2)} kWp</td>
      <td>${formatCurrencyMXN(row.inversion || 0)}</td>
      <td>${formatCurrencyMXN(row.ahorroMensual || 0)}</td>
      <td>${Number(row.retornoAnios || 0).toFixed(1)} años</td>
      <td><span class="badge ${badge}">${escapeHtml(status)}</span></td>
      <td><div class="actions"><button class="btn" data-report-detail="${idx}"><i data-lucide="eye"></i>Ver</button></div></td>
    </tr>
  `;
}

function bindReportPreviewActions() {
    $('#btnReportPdf')?.addEventListener('click', exportReportPdf);
    $('#btnReportCsv')?.addEventListener('click', exportReportCsv);
    $$('#reportQuotesTable [data-report-detail]').forEach(btn => {
        btn.addEventListener('click', () => {
            const idx = Number(btn.dataset.reportDetail || -1);
            const row = state.reportes.data?.rows?.[idx];
            if (row) {
                openReportQuoteDetail(row);
            }
        });
    });
}

function openReportQuoteDetail(row = {}) {
    openModal({
        title: row.folio || 'Cotización',
        subtitle: `${row.fechaTexto || row.fecha || 'Sin fecha'} · ${row.estatus || 'Guardada'}`,
        bodyHtml: `
      <div class="grid cols-2">
        <div class="kpi"><div class="kpi__label">Cliente</div><div class="kpi__value" style="font-size:13px">${escapeHtml(row.cliente || '—')}</div></div>
        <div class="kpi"><div class="kpi__label">Tarifa</div><div class="kpi__value" style="font-size:13px">${escapeHtml(row.tarifa || '—')}</div></div>
        <div class="kpi"><div class="kpi__label">Consumo mensual usado</div><div class="kpi__value">${formatNumber(row.consumoMensual || 0)} kWh</div></div>
        <div class="kpi"><div class="kpi__label">Paneles</div><div class="kpi__value">${formatNumber(row.paneles || 0)}</div></div>
        <div class="kpi"><div class="kpi__label">Potencia estimada</div><div class="kpi__value">${Number(row.potenciaKwp || 0).toFixed(2)} kWp</div></div>
        <div class="kpi"><div class="kpi__label">Inversión</div><div class="kpi__value">${formatCurrencyMXN(row.inversion || 0)}</div></div>
        <div class="kpi"><div class="kpi__label">Ahorro mensual</div><div class="kpi__value">${formatCurrencyMXN(row.ahorroMensual || 0)}</div></div>
        <div class="kpi"><div class="kpi__label">Retorno estimado</div><div class="kpi__value">${Number(row.retornoAnios || 0).toFixed(1)} años</div></div>
      </div>
      <div class="card card--flat" style="margin-top:12px; box-shadow:none">
        <div class="card__subtitle">Datos adicionales</div>
        <div class="review-row"><span>No. de servicio</span><b>${escapeHtml(row.servicio || '—')}</b></div>
        <div class="review-row"><span>Usuario que generó</span><b>${escapeHtml(row.usuario || 'Equipo SECOM')}</b></div>
        <div class="review-row"><span>Convertida en proyecto</span><b>${row.proyectoGenerado ? 'Sí' : 'No'}</b></div>
      </div>
    `,
        footHtml: `<button class="btn" data-close="true"><i data-lucide="x"></i>Cerrar</button>`
    });
}

async function exportReportPdf() {
    if (!window.jspdf || !window.html2canvas) {
        toast({title: 'Exportación no disponible', message: 'Falta jsPDF o html2canvas en el navegador.', icon: 'alert-triangle'});
        return;
    }

    const area = $('#reportPreview');
    if (!area || !state.reportes.data?.rows?.length) {
        toast({title: 'Sin datos', message: 'Genera primero un reporte con cotizaciones.', icon: 'alert-triangle'});
        return;
    }

    try {
        setPillStatus('Generando PDF…', 'busy');
        const canvas = await window.html2canvas(area, {scale: 2, useCORS: true, backgroundColor: '#ffffff', ignoreElements: el => el.classList?.contains('no-print')});
        const imgData = canvas.toDataURL('image/jpeg', 0.95);
        const {jsPDF} = window.jspdf;
        const pdf = new jsPDF({orientation: 'l', unit: 'pt', format: 'a4'});
        const pageW = pdf.internal.pageSize.getWidth();
        const pageH = pdf.internal.pageSize.getHeight();
        const margin = 24;
        const imgW = pageW - margin * 2;
        const imgH = canvas.height * (imgW / canvas.width);
        let heightLeft = imgH;
        let position = margin;

        pdf.addImage(imgData, 'JPEG', margin, position, imgW, imgH);
        heightLeft -= (pageH - margin * 2);
        while (heightLeft > 0) {
            pdf.addPage();
            position = margin - (imgH - heightLeft);
            pdf.addImage(imgData, 'JPEG', margin, position, imgW, imgH);
            heightLeft -= (pageH - margin * 2);
        }

        const f = state.reportes;
        pdf.save(`Reporte_Cotizaciones_SECOM_${f.fechaInicio}_a_${f.fechaFin}.pdf`);
        toast({title: 'PDF generado', message: 'Descarga iniciada.', icon: 'download'});
        setPillStatus('Listo', 'ok');
    } catch (e) {
        console.error(e);
        toast({title: 'No se pudo exportar', message: e?.message || 'Intenta de nuevo.', icon: 'x-circle'});
        setPillStatus('Error', 'error');
    }
}

function exportReportCsv() {
    const rows = state.reportes.data?.rows || [];
    if (!rows.length) {
        toast({title: 'Sin datos', message: 'Genera primero un reporte con cotizaciones.', icon: 'alert-triangle'});
        return;
    }

    const headers = ['Folio', 'Fecha', 'Cliente', 'Tarifa', 'Consumo mensual kWh', 'Paneles', 'Potencia kWp', 'Inversión MXN', 'Ahorro mensual MXN', 'Retorno años', 'Estatus', 'Usuario'];
    const lines = [headers, ...rows.map(r => [
        r.folio || r.id || '',
        r.fechaTexto || r.fecha || '',
        r.cliente || '',
        r.tarifa || '',
        Number(r.consumoMensual || 0),
        Number(r.paneles || 0),
        Number(r.potenciaKwp || 0),
        Number(r.inversion || 0),
        Number(r.ahorroMensual || 0),
        Number(r.retornoAnios || 0),
        r.estatus || '',
        r.usuario || 'Equipo SECOM',
    ])].map(cols => cols.map(csvEscape).join(',')).join('\n');

    const blob = new Blob(['\ufeff' + lines], {type: 'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const f = state.reportes;
    a.href = url;
    a.download = `Reporte_Cotizaciones_SECOM_${f.fechaInicio}_a_${f.fechaFin}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    toast({title: 'Reporte exportado', message: 'Archivo compatible con Excel generado.', icon: 'sheet'});
}

function csvEscape(value) {
    const text = String(value ?? '');
    return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

// ------------------------------
// Opciones
// ------------------------------

function renderOpcionesRoute() {
    const root = $('#route-opciones');
    state.preferences = loadUserPreferences();
    const prefs = state.preferences;
    const theme = prefs.theme || 'dark';

    root.innerHTML = `
    <div class="grid" style="gap:14px">
      <div class="card">
        <div class="card__title">Apariencia</div>
        <div class="card__subtitle">Tema del sistema</div>

        <div class="row" style="align-items:center">
          <div class="help" style="flex:1">Cambia entre modo claro y oscuro. La preferencia queda guardada para este equipo.</div>
          <div class="field" style="max-width:240px">
            <label>Tema</label>
            <select id="themeSelect">
              <option value="dark" ${theme === 'dark' ? 'selected' : ''}>Dark mode</option>
              <option value="light" ${theme === 'light' ? 'selected' : ''}>Light mode</option>
            </select>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card__title">Datos comerciales por defecto</div>
        <div class="card__subtitle">Se usarán en CashVolt y en la cotización final</div>
        <div class="grid cols-2" style="margin-top:10px">
          <div class="field">
            <label>Nombre del asesor</label>
            <input id="prefAdvisor" value="${escapeAttr(prefs.company?.advisorName || '')}" />
          </div>
          <div class="field">
            <label>Nombre de la empresa</label>
            <input id="prefCompany" value="${escapeAttr(prefs.company?.companyName || '')}" />
          </div>
          <div class="field">
            <label>Correo</label>
            <input id="prefEmail" value="${escapeAttr(prefs.company?.companyEmail || '')}" />
          </div>
          <div class="field">
            <label>Teléfono</label>
            <input id="prefPhone" value="${escapeAttr(prefs.company?.companyPhone || '')}" />
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card__title">Parámetros por defecto de cotización</div>
        <div class="card__subtitle">Sirven para nuevas cotizaciones y captura manual</div>
        <div class="grid cols-3" style="margin-top:10px">
          <div class="field">
            <label>Producción promedio</label>
            <input id="prefYield" type="number" min="60" max="220" step="1" value="${escapeAttr(String(prefs.quoteDefaults?.yieldKwhPerKwpMonth || 135))}" />
          </div>
          <div class="field">
            <label>Panel (W)</label>
            <input id="prefPanel" type="number" min="350" max="700" step="10" value="${escapeAttr(String(prefs.quoteDefaults?.panelWatts || 550))}" />
          </div>
          <div class="field">
            <label>Costo por kWp</label>
            <input id="prefCost" type="number" min="12000" max="60000" step="500" value="${escapeAttr(String(prefs.quoteDefaults?.costPerKwp || 22000))}" />
          </div>
          <div class="field">
            <label>Contingencia</label>
            <input id="prefCont" type="number" min="0" max="0.30" step="0.01" value="${escapeAttr(String(prefs.quoteDefaults?.contingencyPct || 0.06))}" />
          </div>
          <div class="field">
            <label>IVA por defecto</label>
            <input id="prefTax" type="number" min="0" max="0.30" step="0.01" value="${escapeAttr(String(prefs.quoteDefaults?.taxPct || 0.16))}" />
          </div>
          <div class="field">
            <label>OCR agresivo</label>
            <select id="prefOcrAggressive">
              <option value="true" ${prefs.ocr?.aggressiveMode ? 'selected' : ''}>Activado</option>
              <option value="false" ${!prefs.ocr?.aggressiveMode ? 'selected' : ''}>Desactivado</option>
            </select>
          </div>
        </div>

        <div class="wizard-actions" style="justify-content:flex-end; margin-top:12px">
          <button class="btn btn--primary" id="btnSavePrefs"><i data-lucide="save"></i>Guardar preferencias</button>
        </div>
      </div>

      <div class="card">
        <div class="card__title">Datos</div>
        <div class="card__subtitle">Administración de historial y proyectos</div>

        <div class="row" style="align-items:center">
          <div class="help" style="flex:1">Puedes limpiar el historial y proyectos almacenados en la base de datos visible desde esta interfaz.</div>
          <button class="btn btn--danger" id="btnReset"><i data-lucide="trash-2"></i>Limpiar todo</button>
        </div>
      </div>
    </div>
  `;

    $('#themeSelect').addEventListener('change', (e) => {
        const t = e.target.value;
        state.preferences.theme = t;
        saveUserPreferences(state.preferences);
        applyTheme(t);
        toast({title: 'Tema actualizado', message: `Modo ${t === 'dark' ? 'oscuro' : 'claro'} activado.`, icon: 'palette'});
        if (state.route === 'cotizador' && state.wizardStep === 3)
            renderChart();
    });

    $('#btnSavePrefs')?.addEventListener('click', () => {
        state.preferences = saveUserPreferences({
            ...state.preferences,
            company: {
                advisorName: $('#prefAdvisor')?.value || '',
                companyName: $('#prefCompany')?.value || '',
                companyEmail: $('#prefEmail')?.value || '',
                companyPhone: $('#prefPhone')?.value || '',
                companyWebsite: state.preferences.company?.companyWebsite || 'https://cashvolt.mx/public/login',
            },
            quoteDefaults: {
                yieldKwhPerKwpMonth: Number($('#prefYield')?.value || 135),
                panelWatts: Number($('#prefPanel')?.value || 550),
                costPerKwp: Number($('#prefCost')?.value || 22000),
                contingencyPct: Number($('#prefCont')?.value || 0.06),
                taxPct: Number($('#prefTax')?.value || 0.16),
            },
            ocr: {
                ...(state.preferences.ocr || {}),
                aggressiveMode: String($('#prefOcrAggressive')?.value || 'true') === 'true',
            }
        });
        toast({title: 'Preferencias guardadas', message: 'Los valores por defecto se actualizaron correctamente.', icon: 'save'});
    });

    $('#btnReset').addEventListener('click', () => {
        openModal({
            title: 'Limpiar datos',
            subtitle: 'Acción irreversible',
            bodyHtml: `
        <div class="help">Se eliminarán todas las cotizaciones del historial y todos los proyectos guardados en este equipo.</div>
        <div class="help" style="margin-top:10px">¿Deseas continuar?</div>
      `,
            footHtml: `
        <button class="btn" data-close="true"><i data-lucide="x"></i>Cancelar</button>
        <button class="btn btn--danger" id="confirmReset"><i data-lucide="trash-2"></i>Limpiar</button>
      `
        });

        $('#confirmReset')?.addEventListener('click', () => {
            resetAllData();
            toast({title: 'Datos eliminados', message: 'Historial y proyectos han sido limpiados.', icon: 'check-circle'});
            closeModal();
            renderHistorialTable();
            renderProyectosRoute();
            resetWizard(true);
            setRoute('cotizador');
        });
    });

    window.lucide?.createIcons();
}

// ------------------------------
// Small escaping helpers
// ------------------------------

function escapeHtml(s) {
    return String(s ?? '').replace(/[&<>"']/g, (ch) => ({
            '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
        }[ch]));
}

function escapeAttr(s) {
    return escapeHtml(s).replace(/\n/g, ' ');
}

function validateReceiptAgainstSelection(receipt, selected) {
    if (!receipt || !selected)
        return {ok: true, message: 'Listo para continuar.'};

    // Validación de periodo (mensual/bimestral)
    const selPeriodo = selected.periodo;
    const recPeriodo = receipt.tipoPeriodo;
    if (selPeriodo && recPeriodo && selPeriodo !== recPeriodo) {
        return {
            ok: false,
            message: `El tipo seleccionado es ${selPeriodo}, pero el recibo está marcado como ${recPeriodo}. Verifica el periodo facturado.`
        };
    }

    return {ok: true, message: 'Tarifa y periodo verificados.'};
}
