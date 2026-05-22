import { $, $$, debounce, formatCurrencyMXN, formatDate, formatDateTime, formatNumber, openModal, closeModal, setPillStatus, toast } from './utils.js';
import { analyzeReceiptFile, createEmptyReceiptData } from './receiptParser.js';
import { computeQuote, buildExportHtml } from './quoteEngine.js';
import { getQuotes, getProjects, getInsumos, getPaquetes, getCotizacionesReport, resetAllData, saveInsumo, savePaquete, saveProjectFromQuote, saveProject, saveQuote, updateInsumo, updatePaquete, updateProject, updateQuote, removeInsumo, removePaquete, removeProject, removeQuote } from './storage.js';
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

//init();

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
    if (route === 'opciones') {
        renderOpcionesRoute();
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

    if (!state.selectedTariff) {
        root.innerHTML = renderTariffSelector();
        wireTariffSelector();
        window.lucide?.createIcons();
        return;
    }

    root.innerHTML = `
        <div class="wizard-shell">

            <div class="wizard-top-clean">
                <div class="pill">
                    <span class="pill__dot"></span>
                    <span>Tarifa:</span>
                    <b>${escapeHtml(state.selectedTariff.label)}</b>
                </div>

                <button class="btn" id="btnChangeTarifa">
                    <i data-lucide="repeat-2"></i>
                    Cambiar
                </button>
            </div>

            <div class="card card--flat wizard-stepper-card">
                <div class="stepper" id="stepper"></div>
            </div>

            <div class="wizard-grid" id="wizardGrid">
                <div id="wizardLeft"></div>
                <div id="wizardRight"></div>
            </div>

            <div id="wizardSingle" class="wizard-single"></div>

        </div>
    `;

    $('#btnChangeTarifa')?.addEventListener('click', () => {
        state.selectedTariff = null;
        resetWizard(true);
        renderCotizadorRoute();
    });

    buildStepper();
    renderWizard();

    window.lucide?.createIcons();
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
    if (!el) return;

    const steps = [
        {n: 1, label: 'Recibo'},
        {n: 2, label: 'Consumo'},
        {n: 3, label: 'Pre-cálculo'},
        {n: 4, label: 'Productos'},
        {n: 5, label: 'Generar'},
    ];

    el.innerHTML = steps.map((s, i) => {

        const cls =
            s.n === state.wizardStep
                ? 'is-active'
                : (s.n < state.wizardStep ? 'is-done' : '');

        const line =
            i < steps.length - 1
                ? '<div class="stepper__line"></div>'
                : '';

        return `
        <div class="stepper__item ${cls}">
            <div class="stepper__dot">
                ${s.n}
            </div>

            <div class="stepper__label">
                ${s.label}
            </div>
        </div>

        ${line}
        `;
    }).join('');

    window.lucide?.createIcons();
}

function gotoStep(n) {
    state.wizardStep = Number(n);
    buildStepper();
    renderWizard();
}
function restoreWizardRight() {
    const grid = document.getElementById('wizardGrid');

    if (!grid) return;

    if (!document.getElementById('wizardRight')) {
        const right = document.createElement('div');
        right.id = 'wizardRight';
        grid.appendChild(right);
    }
}
function renderWizard() {
    const left = $('#wizardLeft');
    const right = $('#wizardRight');
    const grid = document.querySelector('.wizard-grid, .quote-flow-grid');

    if (!left || !right) return;

    grid?.classList.remove('wizard-grid--products', 'wizard-grid--consumption');

    if (state.wizardStep === 1) {
        left.innerHTML = renderStep1Left();
        right.innerHTML = renderStep1Right();
        wireStep1();

    } else if (state.wizardStep === 2) {
        grid?.classList.add('wizard-grid--consumption');

        left.innerHTML = secomV2RenderConsumptionLeft();
        right.innerHTML = secomV2RenderConsumptionRight();
        secomV2WireConsumption();

    } else if (state.wizardStep === 3) {
        left.innerHTML = secomV2RenderPrecalcLeft();
        right.innerHTML = secomV2RenderPrecalcRight();
        secomV2WirePrecalc();

    } else if (state.wizardStep === 4) {
        left.innerHTML = secomV2RenderPackageLeft();
        right.innerHTML = secomV2RenderPackageRight();
        secomV2WirePackage();

    } else {
        left.innerHTML = renderStep4Left();
        right.innerHTML = renderStep4Right();
        wireStep4();
    }

    window.lucide?.createIcons();
}
function renderStep1Left() {

    const hasReceipt = !!state.receipt;
    const r = state.receipt || createEmptyReceiptData(state.selectedTariff);

    return `
    <div class="card__title">Información del recibo</div>
    <div class="help">
        Valida la información detectada desde el recibo CFE.
    </div>

    <input id="fileInput" type="file" accept="application/pdf,image/png,image/jpeg" hidden />

    <div class="dropzone" id="dropzone" tabindex="0" role="button" aria-label="Cargar recibo">
        <div class="dropzone__icon">
            <i data-lucide="upload"></i>
        </div>

        <div style="min-width:0">
            <div class="dropzone__title">
                Arrastre y suelte aquí, o seleccione un archivo
            </div>

            <div class="dropzone__sub" id="fileHint">
                PDF o imagen
            </div>
        </div>
    </div>

    <div class="wizard-actions">
        <button class="btn btn--primary" id="btnAnalyze">
            <i data-lucide="scan"></i>
            Analizar
        </button>

        <button class="btn" id="btnManualCapture">
            <i data-lucide="keyboard"></i>
            Captura manual
        </button>

        <button class="btn" id="btnClear">
            <i data-lucide="trash-2"></i>
            Limpiar
        </button>
    </div>

    <div class="receipt-detected-panel">

        <div class="receipt-detected-header">

            <div class="receipt-status-icon">
                <i data-lucide="${hasReceipt ? 'check-circle' : 'file-text'}"></i>
            </div>

            <div>
                <div class="receipt-detected-title">
                    ${hasReceipt
                        ? 'Información del recibo'
                        : 'Información del recibo'}
                </div>

                <div class="receipt-detected-subtitle">
                    ${hasReceipt
                        ? 'Datos del recibo CFE'
                        : 'Completa o valida los datos antes de continuar'}
                </div>
            </div>

        </div>

        <div class="receipt-info-grid">

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="file-text"></i>
                    Número de servicio
                </div>

                <input
                    id="rServicio"
                    placeholder="###########"
                    value="${escapeAttr(r?.servicio || '')}"
                />
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="user-circle"></i>
                    Nombre del cliente
                </div>

                <input
                    id="rNombre"
                    placeholder="Titular del recibo"
                    value="${escapeAttr(r?.nombre || '')}"
                />
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="zap"></i>
                    Tarifa
                </div>

                <input
                    id="rTarifa"
                    placeholder="1B / DAC / PDBT"
                    value="${escapeAttr(r?.tarifa || state.selectedTariff?.label || '')}"
                />
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="map-pin"></i>
                    Estado
                </div>

                <input
                    id="rEstado"
                    placeholder="SON"
                    value="${escapeAttr(r?.estado || '')}"
                />
            </div>

            <div class="receipt-info-card receipt-info-card--wide">
                <div class="receipt-info-label">
                    <i data-lucide="map-pin"></i>
                    Direccion
                </div>

                <textarea
                    id="rDireccion"
                    rows="2"
                    placeholder="Dirección del suministro"
                >${escapeHtml(r?.direccion || '')}</textarea>
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="calendar"></i>
                     Periodo
                </div>

                <input
                    id="rPeriodo"
                    placeholder="DD MMM AA - DD MMM AA"
                    value="${escapeAttr(r?.periodo?.raw || '')}"
                />
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="calendar-days"></i>
                    Tipo de periodo
                </div>

                <select id="rTipoPeriodo">
                    <option ${r?.tipoPeriodo === 'Mensual' ? 'selected' : ''}>
                        Mensual
                    </option>

                    <option ${r?.tipoPeriodo === 'Bimestral' ? 'selected' : ''}>
                        Bimestral
                    </option>
                </select>
            </div>

        </div>
    </div>

    <div class="wizard-actions" style="justify-content:space-between">

        <div class="help" id="analyzeMsg"></div>

        <button
            class="btn btn--success"
            id="btnStep1Next"
            ${hasReceipt ? '' : 'disabled'}
        >
            <i data-lucide="arrow-right"></i>
            Continuar
        </button>

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
    $('#step2Dirección') && ($('#step2Dirección').textContent = state.receipt?.direccion || '—');
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
    const packageText = state.selectedPackage ? getPackageSummaryLabel(state.selectedPackage, {quote: q, receipt: state.receipt, paneles: q.paneles, consumoMensual: q.consumoMensual}) : 'Sin paquete seleccionado';
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
    state.receipt.insumos = buildPackageItems(packageKey, {quote: q, receipt: state.receipt, paneles: q.paneles, consumoMensual: q.consumoMensual});
    state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct)) ? Number(state.receipt.impuestosPct) : (state.preferences?.quoteDefaults?.taxPct || 0.16);
    state.quote = currentStep2Quote();
    renderWizard();
    toast({title: 'Paquete aplicado', message: `${PACKAGE_PRESETS.find(p => p.key === packageKey)?.label || 'Paquete'} cargado con precios por defecto.`, icon: 'package'});
}

// ------------------------------
// Step 2
// ------------------------------



function renderStep2Right() {
    return secomV2RenderConsumptionRight();
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
        r.insumos.push({codigo: base.codigo, descripcion: base.descripcion, cantidad: 1, unidad: base.unidad, precio: base.precio});
        if (sel)
            sel.value = '';
        state.quote = null;
        renderBody();
    });

    // Agregar manual
    $('#btnAddManual')?.addEventListener('click', () => {
        r.insumos.push({codigo: '', descripcion: '', cantidad: 1, unidad: 'UD', precio: 0});
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
    secomV2WireConsumption();
}
// ------------------------------
// Step 3
// ------------------------------
function renderStep3Left() {
    return secomV2RenderPrecalcLeft();
}

function renderStep3Right() {
    return secomV2RenderPrecalcRight();
}
function wireStep3() {
    secomV2WirePrecalc();
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
    const canvas = document.getElementById('chart');
    if (!canvas || typeof Chart === 'undefined') return;

    const r = state.receipt || {};
    const hist = (r.historial || []).slice(-12);

    const labels = hist.map((_, i) => `P${i + 1}`);
    const consumos = hist.map(h => Number(h.kwh || 0));
    const pagos = hist.map(h => Number(h.pago || 0));

    if (window.secomChart) {
        window.secomChart.destroy();
    }

    window.secomChart = new Chart(canvas, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    type: 'bar',
                    label: 'Consumo (kWh)',
                    data: consumos,
                    yAxisID: 'y',
                    borderWidth: 1
                },
                {
                    type: 'line',
                    label: 'Pago (MXN)',
                    data: pagos,
                    yAxisID: 'y1',
                    tension: 0.35,
                    borderWidth: 3,
                    pointRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    ticks: {
                        color: '#cbd5e1'
                    },
                    grid: {
                        color: 'rgba(255,255,255,.08)'
                    }
                },
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'kWh',
                        color: '#cbd5e1'
                    },
                    ticks: {
                        color: '#cbd5e1'
                    },
                    grid: {
                        color: 'rgba(255,255,255,.08)'
                    }
                },
                y1: {
                    beginAtZero: true,
                    position: 'right',
                    title: {
                        display: true,
                        text: 'MXN',
                        color: '#cbd5e1'
                    },
                    ticks: {
                        color: '#cbd5e1'
                    },
                    grid: {
                        drawOnChartArea: false
                    }
                }
            },
            plugins: {
                legend: {
                    labels: {
                        color: '#cbd5e1'
                    }
                }
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


function addCanvasFitToPage(pdf, canvas, options = {}) {
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const margin = Number(options.margin ?? 18);
    const usableW = pageW - margin * 2;
    const usableH = pageH - margin * 2;
    const ratio = Math.min(usableW / canvas.width, usableH / canvas.height);
    const imgW = canvas.width * ratio;
    const imgH = canvas.height * ratio;
    const x = (pageW - imgW) / 2;
    const y = (pageH - imgH) / 2;
    const imgData = canvas.toDataURL('image/png');
    pdf.addImage(imgData, 'PNG', x, y, imgW, imgH, undefined, 'FAST');
}

function addCanvasToPdfByPages(pdf, canvas, options = {}) {
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const margin = Number(options.margin ?? 24);
    const imgW = pageW - margin * 2;
    const usableH = pageH - margin * 2;
    const sliceHeightPx = Math.floor((usableH / imgW) * canvas.width);

    let y = 0;
    let pageIndex = 0;

    while (y < canvas.height) {
        const h = Math.min(sliceHeightPx, canvas.height - y);
        const pageCanvas = document.createElement('canvas');
        pageCanvas.width = canvas.width;
        pageCanvas.height = h;
        const ctx = pageCanvas.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
        ctx.drawImage(canvas, 0, y, canvas.width, h, 0, 0, canvas.width, h);

        if (pageIndex > 0) {
            pdf.addPage();
        }

        const imgData = pageCanvas.toDataURL('image/jpeg', 0.98);
        const imgH = h * (imgW / canvas.width);
        pdf.addImage(imgData, 'JPEG', margin, margin, imgW, imgH);

        y += h;
        pageIndex += 1;
    }
}

async function exportPdf() {
    if (!window.html2canvas || !window.jspdf) {
        toast({title: 'Exportación no disponible', message: 'Faltan librerías de exportación en el navegador.', icon: 'alert-triangle'});
        return;
    }

    setPillStatus('Generando PDF…', 'busy');

    try {
        // El PDF no debe depender de la base de datos. Si la cotización aún no se ha guardado,
        // se usa un folio temporal únicamente para el nombre del archivo y la vista previa.
        const exportFolio = state.savedQuote?.id || `TEMP-${new Date().toISOString().slice(0,10).replace(/-/g,'')}-${Date.now().toString().slice(-4)}`;

        const area = $('#exportDoc');
        if (!area)
            throw new Error('No se encontró el documento de exportación.');

        // Asegura que la gráfica del documento esté renderizada antes de capturar
        renderExportDocChart({root: document});

        // Pequeña espera para asegurar que el canvas tenga contenido
        await new Promise(r => setTimeout(r, 60));

        const canvas = await window.html2canvas(area, {
            scale: 3,
            useCORS: true,
            backgroundColor: '#ffffff',
            windowWidth: area.scrollWidth,
            windowHeight: area.scrollHeight,
            scrollX: 0,
            scrollY: 0,
            imageTimeout: 15000
        });

        const {jsPDF} = window.jspdf;
        const pdf = new jsPDF({orientation: 'p', unit: 'pt', format: 'letter'});
        addCanvasFitToPage(pdf, canvas, {margin: 0});

        const name = (state.client.nombre || 'Cliente').trim().replace(/\s+/g, '_').slice(0, 36);
        const filename = `Cotizacion_SECOM_${name}_${exportFolio}.pdf`;
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
            <button class="btn btn--danger" data-action="delete" data-id="${q.id}"><i data-lucide="trash-2"></i>Eliminar</button>
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
            } else if (action === 'delete') {
                const cliente = q.client?.nombre || q.receipt?.nombre || q.id;
                const ok = window.confirm(`¿Eliminar la cotización de ${cliente}? Esta acción quitará la cotización de la lista.`);
                if (!ok) return;

                try {
                    removeQuote(q.id);
                    toast({title: 'Cotización eliminada', message: 'La cotización se eliminó correctamente.', icon: 'trash-2'});
                    renderHistorialTable();
                } catch (err) {
                    toast({title: 'No se pudo eliminar', message: err?.message || 'Ocurrió un error al eliminar la cotización.', icon: 'alert-triangle'});
                }
            }
        });
    });
}

function openModalForQuote(q) {
    const r = q.receipt || {};
    const c = q.client || {};

    // 1. Extraer el nombre/etiqueta del paquete
    const pkgKey = r.instalacion?.paqueteSeleccionado || '';
    const pkgText = pkgKey ? getPackageSummaryLabel(pkgKey, {quote: q.quote, receipt: r, paneles: q.quote?.paneles, consumoMensual: q.quote?.consumoMensual}) : 'Sin paquete seleccionado';

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

        const canvas = await window.html2canvas(area, {
            scale: 3,
            useCORS: true,
            backgroundColor: '#ffffff',
            windowWidth: area.scrollWidth,
            windowHeight: area.scrollHeight,
            scrollX: 0,
            scrollY: 0,
            imageTimeout: 15000
        });

        const {jsPDF} = window.jspdf;
        const pdf = new jsPDF({orientation: 'p', unit: 'pt', format: 'letter'});
        addCanvasFitToPage(pdf, canvas, {margin: 0});
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
        const visibleCatalog = setPackageCatalog(state.paquetes.items);

        // La pantalla de cotización y el CRUD deben alimentarse del mismo catálogo.
        // Si el backend responde vacío, se usa el catálogo activo cargado en memoria para
        // evitar que aparezcan paquetes en la cotización pero no en el módulo Paquetes.
        if (!state.paquetes.items.length && Array.isArray(visibleCatalog) && visibleCatalog.length) {
            state.paquetes.items = visibleCatalog.map(pkg => ({...pkg, items: Array.isArray(pkg.items) ? pkg.items.map(it => ({...it})) : []}));
        }
        return state.paquetes.items;
    } catch (e) {
        console.warn('No se pudo cargar el catálogo de paquetes.', e);
        if (!silent) {
            toast({title: 'Paquetes no disponibles', message: e?.message || 'No se pudo consultar la base de datos.', icon: 'alert-triangle'});
        }
        const visibleCatalog = setPackageCatalog(state.paquetes.items || []);
        if (!state.paquetes.items?.length && Array.isArray(visibleCatalog) && visibleCatalog.length) {
            state.paquetes.items = visibleCatalog.map(pkg => ({...pkg, items: Array.isArray(pkg.items) ? pkg.items.map(it => ({...it})) : []}));
        }
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
        const rowKey = String(pkg.id ?? pkg.key ?? pkg.clave ?? '');
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
            <button class="btn" data-paq-ver="${escapeAttr(rowKey)}"><i data-lucide="eye"></i>Ver</button>
            <button class="btn" data-paq-edit="${escapeAttr(rowKey)}"><i data-lucide="edit-3"></i>Editar</button>
            <button class="btn btn--danger" data-paq-del="${escapeAttr(rowKey)}"><i data-lucide="trash-2"></i>Eliminar</button>
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
    const key = String(id ?? '');
    return (state.paquetes.items || []).find(it => String(it.id ?? '') === key || String(it.key ?? it.clave ?? '') === key);
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
    if (!pkg.id) {
        toast({title: 'Paquete de referencia', message: 'Este paquete todavía no está guardado en la base de datos. Editarlo y guardarlo lo registrará como paquete del catálogo.', icon: 'info'});
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
    // 1. Validar que la librería base exista
    if (!window.jspdf) {
        toast({title: 'Exportación no disponible', message: 'Falta jsPDF en el navegador.', icon: 'alert-triangle'});
        return;
    }

    // 2. Extraer la clase constructora jsPDF (¡Nota las mayúsculas!)
    const { jsPDF } = window.jspdf;

    // 3. Validar que el plugin autoTable se haya cargado
    if (!jsPDF.API.autoTable) {
        toast({title: 'Plugin faltante', message: 'Falta la librería jspdf-autotable en index.html', icon: 'alert-triangle'});
        return;
    }

    const rows = state.reportes.data?.rows || [];
    const summary = state.reportes.data?.summary || {};
    const filters = state.reportes;

    if (!rows.length) {
        toast({title: 'Sin datos', message: 'Genera primero un reporte con cotizaciones.', icon: 'alert-triangle'});
        return;
    }

    try {
        setPillStatus('Generando PDF…', 'busy');

        // Instanciamos el documento usando jsPDF
        const doc = new jsPDF({ orientation: 'l', unit: 'pt', format: 'a4' });

        // --- Encabezado ---
        doc.setFontSize(18);
        doc.setTextColor(33, 37, 41);
        doc.text('Reporte de Cotizaciones SECOM', 40, 50);

        doc.setFontSize(10);
        doc.setTextColor(108, 117, 125);
        doc.text(`Periodo: ${filters.fechaInicio || '—'} al ${filters.fechaFin || '—'}`, 40, 70);
        doc.text(`Estatus: ${filters.status === 'todos' ? 'Todos' : filters.status} | Tarifa: ${filters.tarifa === 'todas' ? 'Todas' : filters.tarifa}`, 40, 85);

        // --- KPIs (Resumen) ---
        doc.setFontSize(10);
        doc.setTextColor(33, 37, 41);
        doc.text(`Cotizaciones Totales: ${formatNumber(summary.totalCotizaciones || 0)}`, 40, 115);
        doc.text(`Monto Total: ${formatCurrencyMXN(summary.montoTotal || 0)}`, 240, 115);
        doc.text(`Confirmadas: ${formatNumber(summary.confirmadas || 0)}`, 440, 115);
        doc.text(`Potencia Total: ${Number(summary.potenciaTotalKwp || 0).toFixed(2)} kWp`, 620, 115);

        // --- Preparar datos para la tabla ---
        const tableBody = rows.map(r => [
            r.folio || r.id || '—',
            r.fechaTexto || r.fecha || '—',
            r.cliente || '—',
            r.tarifa || '—',
            `${formatNumber(r.consumoMensual || 0)} kWh`,
            formatNumber(r.paneles || 0),
            `${Number(r.potenciaKwp || 0).toFixed(2)} kWp`,
            formatCurrencyMXN(r.inversion || 0),
            formatCurrencyMXN(r.ahorroMensual || 0),
            `${Number(r.retornoAnios || 0).toFixed(1)} años`,
            r.estatus || '—'
        ]);

        // --- Dibujar tabla nativa ---
        doc.autoTable({
            startY: 140, // Iniciar debajo del resumen
            head: [['Folio', 'Fecha', 'Cliente', 'Tarifa', 'Consumo', 'Paneles', 'Potencia', 'Inversión', 'Ahorro', 'Retorno', 'Estatus']],
            body: tableBody,
            theme: 'striped',
            headStyles: { 
                fillColor: [15, 30, 45], // Color oscuro acorde a tu diseño
                textColor: 255, 
                fontSize: 9,
                fontStyle: 'bold'
            },
            bodyStyles: { 
                fontSize: 8, 
                textColor: 50 
            },
            alternateRowStyles: { 
                fillColor: [248, 249, 250] 
            },
            margin: { top: 40, left: 40, right: 40, bottom: 40 },
            // Agregar número de página en el pie
            didDrawPage: function (data) {
                doc.setFontSize(8);
                doc.setTextColor(150);
                doc.text(
                    `Página ${doc.internal.getNumberOfPages()}`, 
                    data.settings.margin.left, 
                    doc.internal.pageSize.height - 20
                );
            }
        });

        // Guardar el PDF con el nombre formateado
        const f = state.reportes;
        doc.save(`Reporte_Cotizaciones_SECOM_${f.fechaInicio}_a_${f.fechaFin}.pdf`);
        
        toast({title: 'PDF generado', message: 'Reporte nativo exportado correctamente.', icon: 'download'});
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
    if (!root) {
        return;
    }
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
              <option value="dark" ${theme === 'dark' ? 'selected' : ''}>Modo oscuro</option>
              <option value="light" ${theme === 'light' ? 'selected' : ''}>Modo claro</option>
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

    $('#themeSelect')?.addEventListener('change', (e) => {
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

    $('#btnReset')?.addEventListener('click', () => {
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

/* =========================================================
   SECOM - Wizard guiado de cotización V2
   Pegar este bloque AL FINAL de app.js
   ========================================================= */

function secomV2Number(v, fallback = 0) {
    const n = Number(String(v ?? '').replace(/,/g, '').trim());
    return Number.isFinite(n) ? n : fallback;
}

function secomV2GetCoberturaPct(q) {
    const consumo = Number(q?.consumoMensual || 0);
    const produccion = Number(q?.produccionMensual || 0);

    if (!state.selectedPackage && !(state.receipt?.insumos || []).length) {
        return null;
    }

    if (consumo <= 0) {
        return 0;
    }

    return Math.round((produccion / consumo) * 1000) / 10;
}

function secomV2SyncReceiptBasics() {
    if (!state.receipt) {
        state.receipt = createEmptyReceiptData(state.selectedTariff);
    }

    state.receipt.periodo = state.receipt.periodo || {raw: '', start: null, end: null, days: 0};

    if ($('#rServicio')) {
        state.receipt.servicio = ($('#rServicio')?.value || '').replace(/\D/g, '').slice(0, 15);
    }

    if ($('#rTarifa')) {
        state.receipt.tarifa = ($('#rTarifa')?.value || '').trim().toUpperCase();
    }

    if ($('#rNombre')) {
        state.receipt.nombre = ($('#rNombre')?.value || '').trim();
    }

    if ($('#rDireccion')) {
        state.receipt.direccion = ($('#rDireccion')?.value || '').trim();
    }

    if ($('#rPeriodo')) {
        state.receipt.periodo.raw = ($('#rPeriodo')?.value || '').trim();
    }

    if ($('#rTipoPeriodo')) {
        state.receipt.tipoPeriodo = ($('#rTipoPeriodo')?.value || '').trim();
        state.receipt.periodo.days = state.receipt.tipoPeriodo === 'Bimestral' ? 60 : 30;
    }

    if ($('#rEstado')) {
        state.receipt.estado = ($('#rEstado')?.value || '').trim().toUpperCase();
    }

    // En esta versión visual, el nombre y la dirección del cliente se capturan en el recibo.
    // Se sincronizan siempre para que el formato final muestre el nombre completo y la dirección completa.
    state.client.nombre = state.receipt.nombre || state.client.nombre || '';
    state.client.direccion = state.receipt.direccion || state.client.direccion || '';

    applyTariffCalculationAssumptions(state.receipt, state.selectedTariff);
    state.quote = currentStep2Quote();
    return state.quote;
}

function secomV2SyncConsumption() {
    if (!state.receipt) {
        state.receipt = createEmptyReceiptData(state.selectedTariff);
    }

    if ($('#rConsumo')) {
        state.receipt.consumoPeriodo = secomV2Number($('#rConsumo')?.value, state.receipt.consumoPeriodo || 0);
    }

    if ($('#rTotal')) {
        state.receipt.totalAPagar = secomV2Number($('#rTotal')?.value, state.receipt.totalAPagar || 0);
    }

    state.receipt.ajusteConsumo = {
        kwhMes: secomV2Number($('#rAjusteKwh')?.value, state.receipt?.ajusteConsumo?.kwhMes || 0),
        nota: ($('#rAjusteNota')?.value || state.receipt?.ajusteConsumo?.nota || '').trim(),
    };

    applyTariffCalculationAssumptions(state.receipt, state.selectedTariff);
    state.quote = currentStep2Quote();
    return state.quote;
}

function secomV2SyncPrecalc() {
    if (!state.receipt) {
        state.receipt = createEmptyReceiptData(state.selectedTariff);
    }

    if ($('#pYield')) {
        state.params.yieldKwhPerKwpMonth = secomV2Number($('#pYield')?.value, state.params.yieldKwhPerKwpMonth || 135);
    }

    if ($('#pPanel')) {
        state.params.panelWatts = secomV2Number($('#pPanel')?.value, state.params.panelWatts || 550);
    }

    if ($('#pCost')) {
        state.params.costPerKwp = secomV2Number($('#pCost')?.value, state.params.costPerKwp || 22000);
    }

    if ($('#pCont')) {
        state.params.contingencyPct = Math.max(0, Math.min(0.30, secomV2Number($('#pCont')?.value, state.params.contingencyPct || 0.06)));
    }

    if ($('#oPaneles')) {
        const pManual = secomV2Number($('#oPaneles')?.value, 0);
        state.overrides.paneles = pManual > 0 ? Math.round(pManual) : null;
    }

    if ($('#oConsumoMensual')) {
        const consumoManual = secomV2Number($('#oConsumoMensual')?.value, 0);
        state.overrides.consumoMensual = consumoManual > 0 ? consumoManual : null;
    }

    applyTariffCalculationAssumptions(state.receipt, state.selectedTariff);
    state.quote = currentStep2Quote();
    return state.quote;
}

function secomV2SyncPackage() {
    if (!state.receipt) {
        state.receipt = createEmptyReceiptData(state.selectedTariff);
    }

    state.quoteMeta.panelModelo = ($('#mPanelModelo')?.value || state.quoteMeta.panelModelo || '').trim();
    state.quoteMeta.panelDimensiones = ($('#mPanelDim')?.value || state.quoteMeta.panelDimensiones || '').trim();
    state.quoteMeta.inversorModelo = ($('#mInversor')?.value || state.quoteMeta.inversorModelo || '').trim();
    state.quoteMeta.tipoTecho = ($('#mTecho')?.value || state.quoteMeta.tipoTecho || 'No especificado');
    state.quoteMeta.perdidasSombraPct = Number($('#mSombras')?.value ?? state.quoteMeta.perdidasSombraPct ?? 0);
    state.quoteMeta.sombras = $('#mSombras')?.selectedOptions?.[0]?.textContent?.split('(')?.[0]?.trim() || state.quoteMeta.sombras || 'No especificado';
    state.quoteMeta.notasFisicas = ($('#mNotasFisicas')?.value || state.quoteMeta.notasFisicas || '').trim();

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
    state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct))
            ? Number(state.receipt.impuestosPct)
            : (state.preferences?.quoteDefaults?.taxPct || 0.16);

    state.quote = currentStep2Quote();
    return state.quote;
}

/* ------------------------------
   Sobrescribir stepper
------------------------------ */

buildStepper = function () {
    const el = $('#stepper');
    if (!el) return;
    const steps = [
        {n: 1, label: 'Recibo'},
        {n: 2, label: 'Consumo'},
        {n: 3, label: 'Pre-cálculo'},
        {n: 4, label: 'Paquete'},
        {n: 5, label: 'Generar'},
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
};

/* ------------------------------
   Sobrescribir render principal
------------------------------ */

renderWizard = function () {
    const grid = document.querySelector('.quote-flow-grid');

    if (!grid) return;

    if (state.wizardStep === 4) {
        grid.classList.add('quote-flow-grid--single');
        grid.innerHTML = `
            <div class="card quote-main-card quote-main-card--full" id="wizardLeft"></div>
        `;

        const left = $('#wizardLeft');
        left.innerHTML = secomV2RenderPackageLeft();

        secomV2WirePackage();
        window.lucide?.createIcons();
        return;
    }

    grid.classList.remove('quote-flow-grid--single');
    grid.innerHTML = `
        <div class="card quote-main-card" id="wizardLeft"></div>
        <div class="card quote-side-card" id="wizardRight"></div>
    `;

    const left = $('#wizardLeft');
    const right = $('#wizardRight');

    if (state.wizardStep === 1) {
        left.innerHTML = renderStep1Left();
        right.innerHTML = renderStep1Right();
        wireStep1();

    } else if (state.wizardStep === 2) {
        left.innerHTML = secomV2RenderConsumptionLeft();
        right.innerHTML = secomV2RenderConsumptionRight();
        secomV2WireConsumption();

    } else if (state.wizardStep === 3) {
        left.innerHTML = secomV2RenderPrecalcLeft();
        right.innerHTML = secomV2RenderPrecalcRight();
        secomV2WirePrecalc();

    } else {
        left.innerHTML = renderStep4Left();
        right.innerHTML = renderStep4Right();
        wireStep4();
    }

    window.lucide?.createIcons();
};

/* ------------------------------
   Paso 1: Recibo / cliente
------------------------------ */

renderStep1Left = function () {
    const r = state.receipt || createEmptyReceiptData(state.selectedTariff);

    return `
    <div class="card__title">Información del recibo</div>
    <div class="help">Valida la información detectada desde el recibo CFE.</div>

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

    <div class="receipt-detected-panel">

        <div class="receipt-detected-header">
            <div class="receipt-status-icon">
                <i data-lucide="check-circle"></i>
            </div>

            <div>
                <div class="receipt-detected-title">Información del recibo</div>
                <div class="receipt-detected-subtitle">Datos del recibo CFE</div>
            </div>
        </div>

        <div class="receipt-info-grid">

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="file-text"></i>
                    Número de servicio
                </div>
                <input id="rServicio" placeholder="###########" value="${escapeAttr(r?.servicio || '')}" />
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="user-circle"></i>
                    Nombre del cliente
                </div>
                <input id="rNombre" placeholder="Titular del recibo" value="${escapeAttr(r?.nombre || '')}" />
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="zap"></i>
                    Tarifa
                </div>
                <input id="rTarifa" placeholder="1B / DAC / PDBT / ..." value="${escapeAttr(r?.tarifa || state.selectedTariff?.label || '')}" />
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="map-pin"></i>
                    Estado
                </div>
                <input id="rEstado" placeholder="SON" value="${escapeAttr(String(r?.estado || ''))}" />
            </div>

            <div class="receipt-info-card receipt-info-card--wide">
                <div class="receipt-info-label">
                    <i data-lucide="map-pin"></i>
                    Dirección
                </div>
                <textarea id="rDireccion" rows="2" placeholder="Dirección del suministro">${escapeHtml(r?.direccion || '')}</textarea>
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="calendar"></i>
                     Periodo
                </div>
                <input id="rPeriodo" placeholder="DD MMM AA - DD MMM AA" value="${escapeAttr(r?.periodo?.raw || '')}" />
            </div>

            <div class="receipt-info-card">
                <div class="receipt-info-label">
                    <i data-lucide="calendar-days"></i>
                    Tipo de periodo
                </div>
                <select id="rTipoPeriodo">
                    <option ${r?.tipoPeriodo === 'Mensual' ? 'selected' : ''}>Mensual</option>
                    <option ${r?.tipoPeriodo === 'Bimestral' ? 'selected' : ''}>Bimestral</option>
                </select>
            </div>

        </div>
    </div>

    <div class="wizard-actions" style="justify-content:space-between">
      <div class="help" id="analyzeMsg"> </div>
      <button class="btn btn--success" id="btnStep1Next" ${state.receipt ? '' : 'disabled'}>
        <i data-lucide="arrow-right"></i>Continuar
      </button>
    </div>
  `;
};

renderStep1Right = function () {
    const r = state.receipt || {};

    return `
    <div class="card__title">Vista previa</div>

    <div class="preview" id="preview">
      <div class="preview__empty" id="previewEmpty">Sin archivo cargado</div>
    </div>

    <div class="quote-summary-panel" style="margin-top:12px">
      <div class="quote-summary-panel__title">Resumen de validación</div>
      <div class="review-row"><span>No. de servicio</span><b id="kpiServicio">${escapeHtml(r.servicio || '—')}</b></div>
      <div class="review-row"><span>Cliente</span><b id="kpiCliente">${escapeHtml(r.nombre || '—')}</b></div>
      <div class="review-row"><span>Tarifa</span><b id="kpiTarifa">${escapeHtml(r.tarifa || state.selectedTariff?.label || '—')}</b></div>
      <div class="review-row"><span>Dirección</span><b id="kpiDireccion">${escapeHtml(r.direccion || '—')}</b></div>
      <div class="review-row"><span>Periodo</span><b id="kpiPeriodo">${escapeHtml(r.periodo?.raw || '—')}</b></div>
      <div class="review-row"><span>Tipo</span><b id="kpiTipoPeriodo">${escapeHtml(r.tipoPeriodo || state.selectedTariff?.periodo || '—')}</b></div>
      <div class="review-row"><span>Estado</span><b id="kpiEstado">${escapeHtml(r.estado || '—')}</b></div>
    </div>
  `;
};

updateStep1FromReceipt = function () {
    const r = state.receipt;
    if (!r) {
        return;
    }

    const setText = (id, value) => {
        const el = $(`#${id}`);
        if (el) {
            el.textContent = value;
        }
    };

    setText('kpiTarifa', r.tarifa || state.selectedTariff?.label || '—');
    setText('kpiServicio', r.servicio || '—');
    setText('kpiPeriodo', r.periodo?.raw || '—');
    setText('kpiTipoPeriodo', r.tipoPeriodo || state.selectedTariff?.periodo || '—');
    setText('kpiCliente', r.nombre || '—');
    setText('kpiDireccion', r.direccion || '—');
    setText('kpiEstado', r.estado || '—');

    if ($('#rServicio')) $('#rServicio').value = r.servicio || '';
    if ($('#rNombre')) $('#rNombre').value = r.nombre || '';
    if ($('#rTarifa')) $('#rTarifa').value = r.tarifa || state.selectedTariff?.label || '';
    if ($('#rDireccion')) $('#rDireccion').value = r.direccion || '';
    if ($('#rPeriodo')) $('#rPeriodo').value = r.periodo?.raw || '';
    if ($('#rTipoPeriodo')) $('#rTipoPeriodo').value = r.tipoPeriodo || state.selectedTariff?.periodo || 'Mensual';
    if ($('#rEstado')) $('#rEstado').value = r.estado || '';
};

wireStep1 = function () {
    const fileInput = $('#fileInput');
    const dz = $('#dropzone');

    const pickFile = () => fileInput?.click();

    dz?.addEventListener('click', pickFile);
    dz?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') pickFile();
    });

    dz?.addEventListener('dragover', (e) => {
        e.preventDefault();
        dz.classList.add('is-dragover');
    });

    dz?.addEventListener('dragleave', () => dz.classList.remove('is-dragover'));

    dz?.addEventListener('drop', (e) => {
        e.preventDefault();
        dz.classList.remove('is-dragover');
        const f = e.dataTransfer.files?.[0];
        if (f) handleFileSelected(f);
    });

    fileInput?.addEventListener('change', () => {
        const f = fileInput.files?.[0];
        if (f) handleFileSelected(f);
    });

    $('#btnClear')?.addEventListener('click', () => resetWizard(true));

    $('#btnManualCapture')?.addEventListener('click', () => {
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

        $('#analyzeMsg').textContent = 'Captura manual habilitada. Valida los datos básicos y continúa.';
        $('#btnStep1Next').disabled = false;

        setPillStatus('Captura manual', 'busy');
        toast({title: 'Captura manual', message: 'Puedes continuar sin subir archivo.', icon: 'keyboard'});
    });

    $('#btnAnalyze')?.addEventListener('click', async () => {
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
                    if (p?.message) $('#analyzeMsg').textContent = p.message;
                }
            });

            if (!result?.ok) {
                throw new Error(result?.message || 'No se pudo analizar el recibo.');
            }

            state.receipt = result.parsed;
            state.receipt.instalacion = state.receipt.instalacion || {};
            state.receipt.insumos = Array.isArray(state.receipt.insumos) ? state.receipt.insumos : [];
            state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct)) ? Number(state.receipt.impuestosPct) : 0.16;
            state.receipt.instalacion.panelesGrupos = Array.isArray(state.receipt.instalacion.panelesGrupos) ? state.receipt.instalacion.panelesGrupos : [];
            state.receiptCanvas = result.canvas;

            applyTariffCalculationAssumptions(state.receipt, state.selectedTariff);

            state.client.nombre = state.client.nombre || state.receipt.nombre || '';
            state.client.direccion = state.client.direccion || state.receipt.direccion || '';

            updateStep1FromReceipt();

            setPillStatus('Listo', 'ok');
            $('#analyzeMsg').textContent = 'Recibo analizado correctamente.';
            $('#btnStep1Next').disabled = false;

            toast({title: 'Recibo listo', message: 'Datos detectados y listos para validar.', icon: 'check-circle'});
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

            toast({title: 'Análisis fallido', message: 'Se habilitó captura manual.', icon: 'x-circle'});
        } finally {
            $('#btnAnalyze').disabled = false;
        }
    });

    ['#rServicio', '#rTarifa', '#rNombre', '#rDireccion', '#rPeriodo', '#rTipoPeriodo', '#rEstado'].forEach(id => {
        $(id)?.addEventListener('input', () => {
            // No se reescribe el valor del input mientras el usuario escribe; esto permite capturar espacios
            // en nombre completo y dirección sin que el campo se recorte en cada pulsación.
            secomV2SyncReceiptBasics();
            if ($('#btnStep1Next')) $('#btnStep1Next').disabled = false;
            const r = state.receipt || {};
            $('#kpiTarifa') && ($('#kpiTarifa').textContent = r.tarifa || state.selectedTariff?.label || '—');
            $('#kpiServicio') && ($('#kpiServicio').textContent = r.servicio || '—');
            $('#kpiPeriodo') && ($('#kpiPeriodo').textContent = r.periodo?.raw || '—');
            $('#kpiTipoPeriodo') && ($('#kpiTipoPeriodo').textContent = r.tipoPeriodo || state.selectedTariff?.periodo || '—');
            $('#kpiCliente') && ($('#kpiCliente').textContent = r.nombre || '—');
            $('#kpiDireccion') && ($('#kpiDireccion').textContent = r.direccion || '—');
            $('#kpiEstado') && ($('#kpiEstado').textContent = r.estado || '—');
        });
        $(id)?.addEventListener('change', () => {
            secomV2SyncReceiptBasics();
            updateStep1FromReceipt();
            if ($('#btnStep1Next')) $('#btnStep1Next').disabled = false;
        });
    });

    $('#btnStep1Next')?.addEventListener('click', () => {
        secomV2SyncReceiptBasics();
        gotoStep(2);
    });

    updateStep1FileHint();
    updateStep1Preview();
    updateStep1FromReceipt();
};

/* ------------------------------
   Paso 2: Consumo
------------------------------ */
function secomV2GetConsumptionMetrics() {
    const r = state.receipt || createEmptyReceiptData(state.selectedTariff);
    const historial = Array.isArray(r?.historial) ? r.historial : [];

    const periodoEsBimestral = (periodo = '') => {
        const m = String(periodo || '').toUpperCase().match(/(?:DEL\s+)?(\d{1,2})\s+(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s+(\d{2})\s+(?:AL\s+)?(\d{1,2})\s+(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s+(\d{2})/);
        if (!m) return false;
        const months = {ENE:0,FEB:1,MAR:2,ABR:3,MAY:4,JUN:5,JUL:6,AGO:7,SEP:8,OCT:9,NOV:10,DIC:11};
        const a = new Date(2000 + Number(m[3]), months[m[2]] ?? 0, Number(m[1]));
        const b = new Date(2000 + Number(m[6]), months[m[5]] ?? 0, Number(m[4]));
        const days = Math.round((b - a) / 86400000);
        return Number.isFinite(days) && days >= 45;
    };

    const currentIsBimestral = r?.tipoPeriodo === 'Bimestral' || Number(r?.periodo?.days || 0) >= 45;
    const currentKwh = Number(r?.consumoPeriodo || 0);
    const currentPay = Number(r?.totalAPagar || 0);

    const consumosMensuales = [];
    const pagosMensuales = [];
    const consumosPeriodo = [];
    const pagosPeriodo = [];

    if (currentKwh > 0) {
        consumosPeriodo.push(currentKwh);
        consumosMensuales.push(currentIsBimestral ? currentKwh / 2 : currentKwh);
    }
    if (currentPay > 0) {
        pagosPeriodo.push(currentPay);
        pagosMensuales.push(currentIsBimestral ? currentPay / 2 : currentPay);
    }

    historial.forEach(h => {
        const kwh = Number(h?.kwh || 0);
        const pago = Number(h?.pago || h?.importe || 0);
        const bimestral = periodoEsBimestral(h?.periodo || h?.periodoRaw || '');
        if (kwh > 0) {
            consumosPeriodo.push(kwh);
            consumosMensuales.push(bimestral ? kwh / 2 : kwh);
        }
        if (pago > 0) {
            pagosPeriodo.push(pago);
            pagosMensuales.push(bimestral ? pago / 2 : pago);
        }
    });

    const sum = arr => arr.reduce((a, b) => a + b, 0);
    const avg = arr => arr.length ? sum(arr) / arr.length : 0;

    const promedioMensual = consumosMensuales.length
        ? avg(consumosMensuales)
        : (Number(r?.consumoPromedioMensual || 0) || (currentIsBimestral ? currentKwh / 2 : currentKwh));

    const promedioCostoMensual = pagosMensuales.length
        ? avg(pagosMensuales)
        : (Number(r?.pagoPromedioMensual || 0) || (currentIsBimestral ? currentPay / 2 : currentPay));

    const totalPagos = pagosPeriodo.length
        ? sum(pagosPeriodo)
        : Number(r?.totalPagoHistorico || r?.totalPagosHistorico || 0);

    return {
        promedioMensual,
        promedioDiarioWatts: promedioMensual ? (promedioMensual * 1000) / 30 : 0,
        promedioCostoMensual,
        consumoMasAlto: consumosPeriodo.length ? Math.max(...consumosPeriodo) : Number(r?.consumoMaximoHistorico || 0),
        totalPagos,
        tipoPeriodo: r?.tipoPeriodo || (currentIsBimestral ? 'Bimestral' : 'Mensual')
    };
}
function secomV2RenderConsumptionLeft() {
    const r = state.receipt || createEmptyReceiptData(state.selectedTariff);
    const m = secomV2GetConsumptionMetrics();

    const promedioMensual = Number(m.promedioMensual || 0);
    const promedioDiario = Number(m.promedioDiarioWatts || 0);
    const promedioCostoMensual = Number(m.promedioCostoMensual || 0);
    const consumoMasAlto = Number(m.consumoMasAlto || 0);
    const totalPagosRecibo = Number(m.totalPagos || 0);
    const nota = r?.ajusteConsumo?.nota || '';

    return `
    <div class="card__title">Información de consumo</div>
    <div class="help">Analiza el consumo energético detectado del recibo antes del pre-cálculo.</div>

    <div class="consumption-kpi-layout">
      <div class="consumption-kpi-card consumption-kpi-blue">
        <div class="consumption-kpi-icon"><i data-lucide="zap"></i></div>
        <div class="consumption-kpi-label">Promedio mensual</div>
        <div class="consumption-kpi-value">${formatNumber(promedioMensual)}</div>
        <div class="consumption-kpi-unit">kWh/mes</div>
      </div>

      <div class="consumption-kpi-card consumption-kpi-purple">
        <div class="consumption-kpi-icon"><i data-lucide="calendar-days"></i></div>
        <div class="consumption-kpi-label">Promedio diario</div>
        <div class="consumption-kpi-value">${formatNumber(promedioDiario)}</div>
        <div class="consumption-kpi-unit">W/día</div>
      </div>

      <div class="consumption-kpi-card consumption-kpi-green">
        <div class="consumption-kpi-icon"><i data-lucide="dollar-sign"></i></div>
        <div class="consumption-kpi-label">Promedio costo mensual</div>
        <div class="consumption-kpi-value">${formatCurrencyMXN(promedioCostoMensual)}</div>
        <div class="consumption-kpi-unit">MXN/mes</div>
      </div>

      <div class="consumption-kpi-card consumption-kpi-blue">
        <div class="consumption-kpi-icon"><i data-lucide="bar-chart-3"></i></div>
        <div class="consumption-kpi-label">Consumo más alto</div>
        <div class="consumption-kpi-value">${formatNumber(consumoMasAlto)}</div>
        <div class="consumption-kpi-unit">kWh</div>
      </div>

      <div class="consumption-kpi-card consumption-kpi-purple">
        <div class="consumption-kpi-icon"><i data-lucide="receipt"></i></div>
        <div class="consumption-kpi-label">Total pagado histórico</div>
        <div class="consumption-kpi-value">${formatCurrencyMXN(totalPagosRecibo)}</div>
        <div class="consumption-kpi-unit">MXN recibo CFE</div>
      </div>

      <div class="consumption-kpi-card consumption-kpi-green">
        <div class="consumption-kpi-icon"><i data-lucide="calendar"></i></div>
        <div class="consumption-kpi-label">Periodo</div>
        <div class="consumption-kpi-value">${escapeHtml(r?.tipoPeriodo || 'Mensual')}</div>
        <div class="consumption-kpi-unit">tipo de recibo</div>
      </div>
    </div>

    <div class="field" style="margin-top:18px">
      <label>Notas</label>
      <textarea id="rAjusteNota" rows="2" placeholder="Carga futura, ampliación, crecimiento del consumo, observaciones, etc.">${escapeHtml(nota)}</textarea>
    </div>

    <div class="wizard-actions" style="justify-content:space-between">
      <button class="btn" id="btnBackConsumption"><i data-lucide="arrow-left"></i>Volver</button>
      <button class="btn btn--success" id="btnNextConsumption"><i data-lucide="arrow-right"></i>Continuar</button>
    </div>
  `;
}
function secomV2GetHistoryPeriodLabel(h, i) {
    const raw =
        h?.periodo ||
        h?.periodoRaw ||
        h?.periodoFacturado ||
        h?.fecha ||
        h?.mes ||
        h?.label ||
        h?.nombrePeriodo ||
        h?.rango ||
        h?.descripcion ||
        '';

    return raw && String(raw).trim()
        ? String(raw).trim()
        : `P${i + 1}`;
}

function secomV2RenderConsumptionRight() {
    const r = state.receipt || createEmptyReceiptData(state.selectedTariff);
    const hist = (r?.historial || []).slice(-12);

    return `
    <div class="card__title">Tendencia de consumo</div>
    <div class="help">Gráfica y lista del historial detectado en el recibo.</div>

    <div class="consumption-chart-box">
      <canvas id="chart" height="260"></canvas>
    </div>

    <div class="card__subtitle" style="margin-top:16px">Lista de consumos</div>

    <div style="overflow:auto; margin-top:10px; max-height:360px">
      <table class="table table--tight consumption-history-table">
        <thead>
          <tr>
            <th>Periodo</th>
            <th>Consumo (kWh)</th>
            <th>Pago (MXN)</th>
          </tr>
        </thead>

        <tbody>
          ${hist.length ? hist.map((h, i) => `
            <tr>
              <td><b>${escapeHtml(secomV2GetHistoryPeriodLabel(h, i))}</b></td>
              <td><b>${formatNumber(h.kwh || 0)}</b></td>
              <td>${formatCurrencyMXN(h.pago || 0)}</td>
            </tr>
          `).join('') : `
            <tr>
              <td colspan="3" style="color:var(--muted)">
                Sin datos históricos detectados.
              </td>
            </tr>
          `}
        </tbody>
      </table>
    </div>
  `;
}

function secomV2WireConsumption() {
    if (!state.receipt) {
        gotoStep(1);
        return;
    }

    $('#btnBack2')?.addEventListener('click', () => {
        gotoStep(1);
    });

    $('#btnNext2')?.addEventListener('click', () => {
        secomV2SyncConsumption();
        gotoStep(3);
    });

    $('#rAjusteNota')?.addEventListener('input', debounce(() => {
        if (!state.receipt.ajusteConsumo) {
            state.receipt.ajusteConsumo = {};
        }

        state.receipt.ajusteConsumo.nota = $('#rAjusteNota').value.trim();
    }, 120));

    renderChart();
}
function secomV2RenderPrecalcLeft() {
    const q = state.quote || currentStep2Quote();
    state.quote = q;

    const consumoMensual = Number(q.consumoMensual || 0);
    const produccionMensual = Number(q.produccionMensual || 0);
    const produccionDiaria = produccionMensual > 0 ? produccionMensual / 30 : 0;
    const consumoDiario = consumoMensual > 0 ? consumoMensual / 30 : 0;

    const kwRequeridos = Number(q.kwp || 0);
    const kwRecomendados = kwRequeridos > 0 ? kwRequeridos * 1.1 : 0;

    return `
    <div class="card__title">Pre-cálculo energético</div>

    <div class="help">
      Esta sección muestra la potencia requerida y recomendada antes de seleccionar paquetes o productos.
    </div>

    <div class="energy-hero">

      <div class="energy-kpi energy-kpi--blue">
        <div class="energy-kpi__icon">
          <span class="energy-kpi__iconText">⚡</span>
        </div>

        <div class="energy-kpi__eyebrow">
          CAPACIDAD REQUERIDA
        </div>

        <div class="energy-kpi__title">
          Tamaño del sistema
        </div>

        <div class="energy-kpi__value">
          ${kwRequeridos.toFixed(2)}
          <span>kWp</span>
        </div>
      </div>

      <div class="energy-kpi energy-kpi--purple">
        <div class="energy-kpi__icon">
          <span class="energy-kpi__iconText">☀</span>
        </div>

        <div class="energy-kpi__eyebrow">
          CAPACIDAD RECOMENDADA
        </div>

        <div class="energy-kpi__title">
          Tamaño óptimo
        </div>

        <div class="energy-kpi__value">
          ${kwRecomendados.toFixed(2)}
          <span>kWp</span>
        </div>
      </div>

      <div class="energy-kpi energy-kpi--green">
        <div class="energy-kpi__icon">
          <span class="energy-kpi__iconText">↗</span>
        </div>

        <div class="energy-kpi__eyebrow">
          PRODUCCIÓN ESTIMADA
        </div>

        <div class="energy-kpi__title">
          Energía generada
        </div>

        <div class="energy-kpi__value">
          ${formatNumber(produccionMensual)}
          <span>kWh</span>
        </div>
      </div>

      <div class="energy-kpi energy-kpi--orange">
        <div class="energy-kpi__icon">
          <span class="energy-kpi__iconText">▭</span>
        </div>

        <div class="energy-kpi__eyebrow">
          CONSUMO DIARIO
        </div>

        <div class="energy-kpi__title">
          Patrón de uso
        </div>

        <div class="energy-kpi__value">
          ${formatNumber(consumoDiario)}
          <span>kWh/día</span>
        </div>
      </div>

    </div>

    <div class="review-panel" style="margin-top:16px">

      <div class="review-row">
        <span>Producción diaria estimada</span>
        <b>${formatNumber(produccionDiaria)} kWh/día</b>
      </div>

      <div class="review-row">
        <span>Consumo mensual usado</span>
        <b>${formatNumber(consumoMensual)} kWh/mes</b>
      </div>

      <div class="review-row">
        <span>Paneles sugeridos</span>
        <b>${formatNumber(q.paneles || 0)} paneles</b>
      </div>

      <div class="review-row">
        <span>Potencia instalada preliminar</span>
        <b>${Number(q.kwp || 0).toFixed(2)} kWp</b>
      </div>

    </div>

    <div class="wizard-actions" style="justify-content:space-between; margin-top:16px">

      <button type="button" class="btn" id="btnBackPrecalc">
        <i data-lucide="arrow-left"></i>
        Volver
      </button>

      <button type="button"
              class="btn btn--success"
              id="btnNextPrecalc">
        <i data-lucide="arrow-right"></i>
        Seleccionar Producto
      </button>

    </div>
  `;
}
function secomV2WirePrecalc() {

    if (!state.receipt) {
        gotoStep(1);
        return;
    }

    const btnBack = document.getElementById('btnBackPrecalc');
    const btnNext = document.getElementById('btnNextPrecalc');

    if (btnBack) {
        btnBack.onclick = () => {
            gotoStep(2);
        };
    }

    if (btnNext) {
        btnNext.onclick = () => {

            state.quote = state.quote || currentStep2Quote();

            gotoStep(4);
        };
    }

    window.lucide?.createIcons();
}



/* ------------------------------
   Paso 4: Paquete y productos
------------------------------ */
//
//    const r = state.receipt || createEmptyReceiptData(state.selectedTariff);
//    const q = currentStep2Quote();
//    const insumos = Array.isArray(r?.insumos) ? r.insumos : [];
//    const impuestosPct = Number.isFinite(Number(r?.impuestosPct))
//            ? Number(r.impuestosPct)
//            : (state.preferences?.quoteDefaults?.taxPct || 0.16);
//
//    return `
//    <div class="card__title">Paquete y productos</div>
//    <div class="help">Selecciona paquete, paneles, inversor, estructura, protecciones e insumos adicionales.</div>
//
//    <div class="card__subtitle" style="margin-top:14px">Paquete solar</div>
//    <div class="package-grid" style="margin-top:10px">
//      ${renderPackageCards()}
//    </div>
//
//    <div class="row" style="justify-content:space-between; margin-top:10px; align-items:flex-start">
//      <div class="help" id="packageLabel">
//        ${state.selectedPackage ? getPackageSummaryLabel(state.selectedPackage, {quote: q, receipt: state.receipt, paneles: q.paneles, consumoMensual: q.consumoMensual}) : 'Sin paquete seleccionado.'}
//      </div>
//      <button class="btn" id="btnClearPackage"><i data-lucide="eraser"></i>Quitar paquete</button>
//    </div>
//
//    <div class="package-preview" id="packagePreview" style="margin-top:10px">
//      ${renderPackagePreviewItems()}
//    </div>
//
//    <div class="card__subtitle" style="margin-top:14px">Componentes principales</div>
//
//    <div class="grid cols-3" style="margin-top:10px">
//      <div class="field">
//        <label>Paneles</label>
//        <input id="mPanelModelo" placeholder="Ej. JA Solar 550 W" value="${escapeAttr(state.quoteMeta.panelModelo || '')}" />
//      </div>
//      <div class="field">
//        <label>Dimensiones del panel</label>
//        <input id="mPanelDim" placeholder="Ej. 2.27 x 1.13 m" value="${escapeAttr(state.quoteMeta.panelDimensiones || '')}" />
//      </div>
//      <div class="field">
//        <label>Inversor</label>
//        <input id="mInversor" placeholder="Ej. Huawei SUN2000" value="${escapeAttr(state.quoteMeta.inversorModelo || '')}" />
//      </div>
//      <div class="field">
//        <label>Estructura / tipo de techo</label>
//        <select id="mTecho">
//          ${['No especificado', 'Losa', 'Lámina', 'Teja', 'Otro'].map(x => `<option ${x === state.quoteMeta.tipoTecho ? 'selected' : ''}>${x}</option>`).join('')}
//        </select>
//      </div>
//      <div class="field">
//        <label>Sombras / pérdidas</label>
//        <select id="mSombras">
//          <option value="0" ${Number(state.quoteMeta.perdidasSombraPct || 0) === 0 ? 'selected' : ''}>Ninguna (0%)</option>
//          <option value="0.10" ${Number(state.quoteMeta.perdidasSombraPct || 0) === 0.10 ? 'selected' : ''}>Baja (10%)</option>
//          <option value="0.20" ${Number(state.quoteMeta.perdidasSombraPct || 0) === 0.20 ? 'selected' : ''}>Media (20%)</option>
//          <option value="0.35" ${Number(state.quoteMeta.perdidasSombraPct || 0) === 0.35 ? 'selected' : ''}>Alta (35%)</option>
//        </select>
//      </div>
//      <div class="field">
//        <label>Protecciones</label>
//        <input value="Protecciones CC/CA" disabled />
//      </div>
//    </div>
//
//    <div class="field" style="margin-top:10px">
//      <label>Notas técnicas</label>
//      <textarea id="mNotasFisicas" rows="2" placeholder="Área disponible, orientación, estructura, protecciones u observaciones.">${escapeHtml(state.quoteMeta.notasFisicas || '')}</textarea>
//    </div>
//
//    <div class="card__subtitle" style="margin-top:14px">Productos / insumos adicionales</div>
//
//    <div class="row" style="margin-top:10px; align-items:flex-end">
//      <div class="field" style="flex:2">
//        <label>Agregar desde catálogo</label>
//        <select id="insCatalog">
//          <option value="">Selecciona un insumo...</option>
//          ${INSUMO_CATALOG.map((it, idx) => `<option value="${idx}">${escapeHtml(it.codigo)} · ${escapeHtml(it.descripcion)} · ${formatCurrencyMXN(it.precio)}</option>`).join('')}
//        </select>
//      </div>
//      <div style="display:flex; gap:10px; flex-wrap:wrap">
//        <button class="btn" id="btnAddCatalog"><i data-lucide="plus"></i>Agregar</button>
//        <button class="btn" id="btnAddManual"><i data-lucide="edit-3"></i>Agregar manual</button>
//      </div>
//    </div>
//
//    <div style="overflow:auto; margin-top:10px">
//      <table class="table table--tight" id="insTable">
//        <thead>
//          <tr>
//            <th style="min-width:140px">Código</th>
//            <th style="min-width:280px">Descripción</th>
//            <th style="min-width:120px">Cantidad</th>
//            <th style="min-width:120px">Unidad</th>
//            <th style="min-width:140px">Precio</th>
//            <th style="min-width:140px">Total</th>
//          </tr>
//        </thead>
//        <tbody>
//          ${insumos.length ? insumos.map((it, i) => renderInsumoRow(it, i)).join('') : `
//            <tr><td colspan="6" style="color:var(--muted)">Aún no hay insumos agregados.</td></tr>
//          `}
//        </tbody>
//      </table>
//    </div>
//
//    <div class="insumos-summary" style="margin-top:10px">
//      <div class="insumos-summary__row"><span>Subtotal</span><b id="insSubtotal">—</b></div>
//      <div class="insumos-summary__row">
//        <span>Impuestos (IVA %)</span>
//        <div style="display:flex; align-items:center; gap:10px">
//          <input id="insTaxPct" class="insumos-summary__tax" type="number" min="0" max="30" step="0.5" value="${escapeAttr(String(Math.round(impuestosPct * 100 * 10) / 10))}" />
//          <b id="insTaxes">—</b>
//        </div>
//      </div>
//      <div class="insumos-summary__row insumos-summary__row--total"><span>Total</span><b id="insTotal">—</b></div>
//    </div>
//
//    <div class="wizard-actions" style="justify-content:space-between">
//      <button class="btn" id="btnBackPackage"><i data-lucide="arrow-left"></i>Volver</button>
//      <button class="btn btn--success" id="btnNextPackage"><i data-lucide="arrow-right"></i>Generar cotización</button>
//    </div>
//
//    <div class="help" id="msg2" style="margin-top:10px"></div>
//  `;
//}
//
//    const q = currentStep2Quote();
//    const coverage = secomV2GetCoberturaPct(q);
//
//    return `
//    <div class="card__title">Resumen de cotización</div>
//
//    <div class="quote-summary-panel" style="margin-top:12px">
//      <div class="quote-summary-panel__title">${state.selectedPackage ? 'Paquete seleccionado' : 'Selecciona productos'}</div>
//      <div class="review-row"><span>Instalado</span><b id="step2Kwp">${Number(q.kwp || 0).toFixed(2)} kWp</b></div>
//      <div class="review-row"><span>Producción estimada</span><b>${formatNumber(q.produccionMensual || 0)} kWh/mes</b></div>
//      <div class="review-row"><span>Cobertura</span><b id="summaryCobertura">${coverage == null ? 'Pendiente' : `${coverage}%`}</b></div>
//      <div class="review-row"><span>Ahorro estimado</span><b id="step2Saving">${formatCurrencyMXN(q.ahorroMensual || 0)}</b></div>
//      <div class="review-row"><span>Total cotización</span><b id="kpiQuoteTotal">${formatCurrencyMXN(q.inversion || 0)}</b></div>
//    </div>
//
//    <div class="grid cols-2" style="margin-top:12px">
//      <div class="kpi quote-kpi-card">
//        <div class="kpi__label">Paneles</div>
//        <div class="kpi__value">${formatNumber(q.paneles || 0)}</div>
//      </div>
//      <div class="kpi quote-kpi-card">
//        <div class="kpi__label">Retorno</div>
//        <div class="kpi__value">${q.retornoAnios ? `${q.retornoAnios} años` : '—'}</div>
//      </div>
//    </div>
//
//    <div id="step2Alerts" style="margin-top:12px">
//      ${state.selectedPackage || (state.receipt?.insumos || []).length
//            ? '<div class="review-ok">Productos listos para generar la cotización.</div>'
//            : '<div class="review-alert">Selecciona un paquete o agrega productos para calcular cobertura.</div>'}
//    </div>
//  `;
//}

function secomV2RefreshPackageSummary() {
    const q = secomV2SyncPackage();
    const coverage = secomV2GetCoberturaPct(q);

    if ($('#step2Kwp')) {
        $('#step2Kwp').textContent = `${Number(q.kwp || 0).toFixed(2)} kWp`;
    }

    if ($('#step2Saving')) {
        $('#step2Saving').textContent = formatCurrencyMXN(q.ahorroMensual || 0);
    }

    if ($('#summaryCobertura')) {
        $('#summaryCobertura').textContent = coverage == null ? 'Pendiente' : `${coverage}%`;
    }

    if ($('#kpiQuoteTotal')) {
        $('#kpiQuoteTotal').textContent = formatCurrencyMXN(q.inversion || 0);
    }

    if ($('#packageLabel')) {
        $('#packageLabel').textContent = state.selectedPackage
                ? getPackageSummaryLabel(state.selectedPackage, {quote: q, receipt: state.receipt, paneles: q.paneles, consumoMensual: q.consumoMensual})
                : 'Sin paquete seleccionado.';
    }
}

applyPackagePreset = function (packageKey) {
    secomV2SyncPackage();
    state.selectedPackage = packageKey;

    const baseQuote = state.quote || currentStep2Quote();
    const pkg = (PACKAGE_PRESETS || []).find(p => String(p.key) === String(packageKey));

    state.receipt.insumos = buildPackageItems(packageKey, {
        quote: baseQuote,
        receipt: state.receipt,
        paneles: baseQuote.paneles,
        consumoMensual: baseQuote.consumoMensual
    });

    state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct))
            ? Number(state.receipt.impuestosPct)
            : (state.preferences?.quoteDefaults?.taxPct || 0.16);

    state.receipt.instalacion = {
        ...(state.receipt.instalacion || {}),
        paqueteSeleccionado: packageKey,
        paqueteLabel: pkg?.label || pkg?.nombre || 'Paquete completo'
    };

    state.quote = secomV3CurrentQuote();

    renderWizard();

    toast({
        title: 'Paquete aplicado',
        message: `${pkg?.label || 'Paquete'} cargado con insumos y precios vigentes.`,
        icon: 'package'
    });
};
//
//function secomV2WirePackage() {
//    if (!state.receipt) {
//        gotoStep(1);
//        return;
//    }
//
//    $('#btnBackPackage')?.addEventListener('click', () => gotoStep(3));
//
//    $('#btnNextPackage')?.addEventListener('click', () => {
//        const q = secomV2SyncPackage();
//        const msg = $('#msg2');
//
//        if (!state.client.nombre) {
//            state.client.nombre = state.receipt?.nombre || '';
//        }
//
//        if (!state.client.nombre) {
//            msg.textContent = 'Por favor, captura el nombre del cliente en la validación del recibo.';
//            toast({title: 'Falta información', message: 'El nombre del cliente es obligatorio.', icon: 'alert-triangle'});
//            return;
//        }
//
//        if (!state.receipt.servicio) {
//            msg.textContent = 'Por favor, captura el número de servicio del recibo.';
//            toast({title: 'Dato crítico faltante', message: 'El número de servicio es obligatorio para continuar.', icon: 'alert-triangle'});
//            return;
//        }
//
//        state.quote = q;
//        gotoStep(5);
//    });
//
//    $$('#route-cotizador [data-package]').forEach(btn => {
//        btn.addEventListener('click', () => applyPackagePreset(btn.dataset.package));
//    });
//
//    $('#btnClearPackage')?.addEventListener('click', () => {
//        state.selectedPackage = '';
//        if (state.receipt?.instalacion) {
//            state.receipt.instalacion.paqueteSeleccionado = '';
//        }
//        renderWizard();
//        toast({title: 'Paquete retirado', message: 'Puedes mantener insumos manuales o cargar otro paquete.', icon: 'eraser'});
//    });
//
//    ['#mPanelModelo', '#mPanelDim', '#mInversor', '#mTecho', '#mSombras', '#mNotasFisicas'].forEach(id => {
//        $(id)?.addEventListener('input', debounce(secomV2RefreshPackageSummary, 120));
//        $(id)?.addEventListener('change', debounce(secomV2RefreshPackageSummary, 120));
//    });
//
//    setupInsumosCrud();
//    secomV2RefreshPackageSummary();
//}

/* ------------------------------
   Paso 5: Ajustar nombre y regreso
------------------------------ */

renderStep4Left = function () {
    const q = state.quote || computeQuote(state.receipt, state.client, state.params, state.overrides);
    state.quote = q;

    return `
    <div class="card__title">Generar cotización</div>
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
};

wireStep4 = function () {
    if (!state.receipt) {
        gotoStep(1);
        return;
    }

    $('#btnBack4')?.addEventListener('click', () => gotoStep(4));
    $('#btnExport')?.addEventListener('click', exportPdf);

    $('#btnGoCashvoltFromQuote')?.addEventListener('click', () => {
        window.open('https://cashvolt.mx/public/login', '_blank');
    });

    $('#btnSaveQuote')?.addEventListener('click', () => {
        try {
            const q = persistQuote('Guardada');
            toast({title: 'Cotización guardada', message: `Se agregó al historial (${q.id}).`, icon: 'save'});
            $('#msg4').textContent = `Guardada como ${q.id}.`;
            renderHistorialTable();
        } catch (e) {
            toast({title: 'No se pudo guardar', message: e?.message || 'Revisa los datos.', icon: 'x-circle'});
        }
    });

    $('#btnConfirmProject')?.addEventListener('click', () => {
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

    if (state.savedQuote?.id) {
        $('#msg4').textContent = `Cotización: ${state.savedQuote.id}`;
    }

    renderExportDocChart({root: document});
};

/* ------------------------------
   Si el usuario está dentro del cotizador, refrescar pantalla
------------------------------ */

if (state.route === 'cotizador') {
    renderCotizadorRoute();
}
/* =========================================================
   SECOM - Mejora visual Wizard V2
   Pegar DESPUÉS del bloque V2 y ANTES del init() final
   ========================================================= */

renderCotizadorRoute = function () {
    const root = $('#route-cotizador');

    if (!state.selectedTariff) {
        root.innerHTML = renderTariffSelector();
        wireTariffSelector();
        window.lucide?.createIcons();
        return;
    }

    root.innerHTML = `
    <div class="quote-flow-shell">

     

        <div class="quote-flow-actions">
          <div class="pill quote-flow-pill">
            <span class="pill__dot"></span>
            <span>Tarifa:</span>&nbsp;<b>${escapeHtml(state.selectedTariff.label)}</b>
          </div>
          <button class="btn" id="btnChangeTarifa">
            <i data-lucide="repeat-2"></i>Cambiar
          </button>
        </div>
      </div>

      
      <div class="quote-flow-grid">
        <div class="card quote-main-card" id="wizardLeft"></div>
        <div class="card quote-side-card" id="wizardRight"></div>
      </div>

    </div>
  `;

    $('#btnChangeTarifa')?.addEventListener('click', () => {
        state.selectedTariff = null;
        resetWizard(true);
        renderCotizadorRoute();
    });

    buildStepper();
    renderWizard();
};

/* FIX WIZARD V2 - consumo dividido izquierda/derecha */

buildStepper = function () {
    const el = $('#stepper');
    if (!el) return;
    const steps = [
        {n: 1, label: 'Recibo'},
        {n: 2, label: 'Consumo'},
        {n: 3, label: 'Pre-cálculo'},
        {n: 4, label: 'Paquete'},
        {n: 5, label: 'Generar'},
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
};



secomV2RenderConsumptionRight = function () {
    const r = state.receipt || createEmptyReceiptData(state.selectedTariff);
    const hist = (r?.historial || []).slice(-12);

    return `
    <div class="card__title">Tendencia de consumo</div>
    <div class="help">Gráfica y lista del historial detectado en el recibo.</div>

    <div style="margin-top:12px; min-height:260px">
      <canvas id="chart" height="220"></canvas>
    </div>

    <div class="card__subtitle" style="margin-top:16px">Lista de consumos</div>

    <div style="overflow:auto; margin-top:10px; max-height:360px">
      <table class="table table--tight consumption-history-table">
        <thead>
          <tr>
            <th>Periodo</th>
            <th>Consumo (kWh)</th>
            <th>Pago (MXN)</th>
          </tr>
        </thead>
        <tbody>
          ${hist.length ? hist.map((h, i) => `
            <tr>
              <td>P${i + 1}</td>
              <td><b>${formatNumber(h.kwh || 0)}</b></td>
              <td>${formatCurrencyMXN(h.pago || 0)}</td>
            </tr>
          `).join('') : `
            <tr>
              <td colspan="3" style="color:var(--muted)">Sin datos históricos detectados.</td>
            </tr>
          `}
        </tbody>
      </table>
    </div>
  `;
};

secomV2WireConsumption = function () {
    if (!state.receipt) {
        gotoStep(1);
        return;
    }

    $('#btnBackConsumption')?.addEventListener('click', () => gotoStep(1));

    $('#btnNextConsumption')?.addEventListener('click', () => {
        secomV2SyncConsumption();
        gotoStep(3);
    });

    $('#rAjusteNota')?.addEventListener('input', debounce(() => {
        if (!state.receipt.ajusteConsumo) state.receipt.ajusteConsumo = {};
        state.receipt.ajusteConsumo.nota = $('#rAjusteNota').value.trim();
    }, 120));

    renderChart();
};
function secomV2RenderPrecalcRight() {
    const q = state.quote || currentStep2Quote();

    const consumoMensual = Number(q.consumoMensual || 0);
    const consumoDiario = consumoMensual > 0 ? consumoMensual / 30 : 0;

    return `
    <div class="card__title">Lectura técnica</div>
    <div class="help">Resumen técnico antes de seleccionar paquete o productos.</div>

    <div class="review-panel" style="margin-top:12px">
      <div class="review-row">
        <span>Consumo mensual</span>
        <b>${formatNumber(consumoMensual)} kWh</b>
      </div>

      <div class="review-row">
        <span>Consumo diario</span>
        <b>${formatNumber(consumoDiario)} kWh/día</b>
      </div>

      <div class="review-row">
        <span>Producción mensual estimada</span>
        <b>${formatNumber(q.produccionMensual || 0)} kWh</b>
      </div>

      <div class="review-row">
        <span>Producción anual estimada</span>
        <b>${formatNumber(q.produccionAnual || 0)} kWh</b>
      </div>

      <div class="review-row">
        <span>Paneles sugeridos</span>
        <b>${formatNumber(q.paneles || 0)}</b>
      </div>

      <div class="review-row">
        <span>Retorno estimado</span>
        <b>${q.retornoAnios ? `${q.retornoAnios} años` : '—'}</b>
      </div>
    </div>

    <div class="review-ok" style="margin-top:12px">
      La cobertura se calculará después de seleccionar paquete o productos.
    </div>
  `;
}
/* =========================================================
   SECOM V3 - PRODUCT CATALOG / CARRITO / PRECIO POR PANEL
   Pegar este bloque ANTES del init(); final
   ========================================================= */

function secomV3EnsureProductState() {
    if (!state.receipt) {
        state.receipt = createEmptyReceiptData(state.selectedTariff);
    }

    state.receipt.insumos = Array.isArray(state.receipt.insumos)
        ? state.receipt.insumos
        : [];

    state.receipt.impuestosPct = Number.isFinite(Number(state.receipt.impuestosPct))
        ? Number(state.receipt.impuestosPct)
        : (state.preferences?.quoteDefaults?.taxPct || 0.16);

    state.productCatalog = state.productCatalog || {
        activeCategory: 'paneles'
    };

    state.receipt.precioPanelInstalado = Number.isFinite(Number(state.receipt.precioPanelInstalado))
        ? Number(state.receipt.precioPanelInstalado)
        : 10000;

    if (typeof state.receipt.usarPrecioGlobalPanel !== 'boolean') {
        state.receipt.usarPrecioGlobalPanel = true;
    }
}

function secomV3UseGlobalPanelPrice() {
    secomV3EnsureProductState();
    return state.receipt.usarPrecioGlobalPanel !== false;
}

function secomV3FindCatalogProductForItem(item = {}) {
    const catalogId = item.catalogId != null ? String(item.catalogId) : '';
    const codigo = item.codigo != null ? String(item.codigo) : '';

    return secomV3GetCatalogProducts().find(product => {
        return (catalogId && String(product.id) === catalogId)
            || (codigo && String(product.codigo) === codigo);
    }) || null;
}

function secomV3ResolvePanelPrice(item = {}) {
    if (secomV3UseGlobalPanelPrice()) {
        return Number(state.receipt.precioPanelInstalado || 10000);
    }

    const catalogProduct = secomV3FindCatalogProductForItem(item);
    const catalogPrice = Number(catalogProduct?.precio || 0);

    if (Number.isFinite(catalogPrice) && catalogPrice > 0) {
        return catalogPrice;
    }

    const currentPrice = Number(item.precio || 0);
    return Number.isFinite(currentPrice) ? Math.max(0, currentPrice) : 0;
}

function secomV3SyncPanelPrices() {
    secomV3EnsureProductState();

    state.receipt.insumos = state.receipt.insumos.map(item => {
        if (!secomV3IsPanelItem(item)) {
            return item;
        }

        return {
            ...item,
            precio: secomV3ResolvePanelPrice(item)
        };
    });
}

function secomV3NormalizeCategory(value = '') {
    const txt = String(value || '').toLowerCase();

    if (txt.includes('panel')) return 'paneles';
    if (txt.includes('inversor')) return 'inversores';
    if (txt.includes('estructura') || txt.includes('montaje') || txt.includes('riel') || txt.includes('soporte')) return 'montaje';
    if (txt.includes('proteccion') || txt.includes('protección') || txt.includes('break') || txt.includes('interruptor') || txt.includes('fusible')) return 'proteccion';
    return 'otros';
}

function secomV3GetProductIcon(category) {
    const icons = {
        paneles: 'zap',
        inversores: 'cpu',
        montaje: 'wrench',
        proteccion: 'shield-check',
        sombras: 'cloud-sun',
        paquetes: 'boxes',
        otros: 'package'
    };

    return icons[category] || 'package';
}

function secomV3GetCategoryLabel(category) {
    const labels = {
        paneles: 'Paneles',
        inversores: 'Inversores',
        montaje: 'Montaje',
        proteccion: 'Protección',
        sombras: 'Sombras',
        paquetes: 'Paquetes completos',
        otros: 'Otros'
    };

    return labels[category] || 'Otros';
}

function secomV3ExtractWatts(item = {}) {
    const direct = Number(item.watts || item.capacidad || item.potencia || 0);

    if (Number.isFinite(direct) && direct > 0) {
        return direct;
    }

    const text = `${item.codigo || ''} ${item.descripcion || ''} ${item.nombre || ''}`;
    const match = text.match(/(\d{3,4})\s*w/i);

    return match ? Number(match[1]) : 0;
}

function secomV3GetCatalogProducts() {
    const source = Array.isArray(INSUMO_CATALOG) ? INSUMO_CATALOG : [];

    const products = source
        .filter(item => item && item.activo !== false)
        .map((item, index) => {
            const descripcion = item.descripcion || item.nombre || item.codigo || 'Producto';
            const category = secomV3NormalizeCategory(`${item.categoria || ''} ${descripcion} ${item.codigo || ''}`);
            const watts = secomV3ExtractWatts(item);

            return {
                id: String(item.id ?? item.codigo ?? index),
                codigo: item.codigo || `PROD-${index + 1}`,
                nombre: descripcion,
                marca: item.marca || item.proveedor || '',
                categoria: category,
                unidad: item.unidad || 'PZA',
                precio: Number(item.precio || 0),
                impuestoPct: Number(item.impuestoPct ?? 0.16),
                watts
            };
        });

    return products;
}

function secomV3GetProductsByCategory(category) {
    return secomV3GetCatalogProducts().filter(p => p.categoria === category);
}

function secomV3IsPanelItem(item = {}) {
    const txt = `${item.codigo || ''} ${item.descripcion || ''} ${item.categoria || ''}`.toLowerCase();

    return txt.includes('panel') || Number(item.watts || 0) > 0;
}

function secomV3PanelItems() {
    secomV3EnsureProductState();

    return state.receipt.insumos.filter(item => secomV3IsPanelItem(item));
}

function secomV3SelectedPanelCount() {
    return secomV3PanelItems().reduce((acc, item) => acc + Number(item.cantidad || 0), 0);
}

function secomV3InstalledWatts() {
    return secomV3PanelItems().reduce((acc, item) => {
        return acc + (Number(item.cantidad || 0) * Number(item.watts || 0));
    }, 0);
}

function secomV3ComputeTotals() {
    secomV3EnsureProductState();

    const subtotal = state.receipt.insumos.reduce((acc, item) => {
        return acc + (Number(item.cantidad || 0) * Number(item.precio || 0));
    }, 0);

    const ivaPct = Number(state.receipt.impuestosPct ?? 0.16);
    const iva = subtotal * ivaPct;
    const total = subtotal + iva;

    return {
        subtotal,
        iva,
        total,
        ivaPct
    };
}

function secomV3CurrentQuote() {
    secomV3EnsureProductState();

    const base = currentStep2Quote();

    const installedWatts = secomV3InstalledWatts();
    const selectedPanelCount = secomV3SelectedPanelCount();
    const selectedKwp = installedWatts > 0 ? installedWatts / 1000 : 0;

    const effectiveKwp = selectedKwp > 0 ? selectedKwp : Number(base.kwp || 0);
    const effectivePanels = selectedPanelCount > 0 ? selectedPanelCount : Number(base.paneles || 0);

    const yieldMonth = Number(base.yieldEfectivo || state.params?.yieldKwhPerKwpMonth || 135);
    const produccionMensual = effectiveKwp > 0 ? effectiveKwp * yieldMonth : 0;
    const produccionAnual = produccionMensual * 12;

    const consumoMensual = Number(base.consumoMensual || 0);
    const coverage = consumoMensual > 0 && produccionMensual > 0
        ? Math.round((produccionMensual / consumoMensual) * 1000) / 10
        : null;

    const totals = secomV3ComputeTotals();
    const hasSelectedInsumos = totals.total > 0;

    const pagoMensual = Number(base.pagoProm || 0);
    const coberturaSeleccionada = coverage == null ? 0 : Math.max(0, Math.min(1, coverage / 100));
    const coberturaBase = consumoMensual > 0
        ? Math.max(0, Math.min(1, Number(base.produccionMensual || 0) / consumoMensual))
        : 0;
    // Cuando se agregan insumos/productos, el retorno se estima contra el sistema completo recomendado,
    // no contra una partida aislada del catálogo. Así se evita que el retorno se dispare cuando se elige
    // un modelo de panel o un insumo sin capturar manualmente toda la cantidad sugerida de paneles.
    const coberturaFinanciera = hasSelectedInsumos
        ? Math.max(coberturaSeleccionada, coberturaBase)
        : (coverage == null ? coberturaBase || 1 : coberturaSeleccionada);
    const ahorroCalculado = pagoMensual > 0 ? pagoMensual * coberturaFinanciera : Number(base.ahorroMensual || 0);
    const ahorroMensual = Math.round(Math.max(0, Math.min(pagoMensual || ahorroCalculado, ahorroCalculado)));

    const inversion = hasSelectedInsumos ? totals.total : Number(base.inversion || 0);
    const retornoAnios = ahorroMensual > 0 && inversion > 0
        ? Math.round((inversion / (ahorroMensual * 12)) * 10) / 10
        : 0;

    return {
        ...base,
        kwp: effectiveKwp,
        paneles: effectivePanels,
        wattsInstalados: installedWatts,
        produccionMensual: Math.round(produccionMensual),
        produccionAnual: Math.round(produccionAnual),
        porcentajeCobertura: coverage,
        ahorroMensual,
        inversion,
        subtotalInsumos: totals.subtotal,
        impuestosInsumos: totals.iva,
        totalInsumos: totals.total,
        retornoAnios
    };
}

function secomV3RenderProductCard(product) {
    const icon = secomV3GetProductIcon(product.categoria);
    const isPanel = product.categoria === 'paneles';

    return `
        <div class="secom-product-card">
            <div class="secom-product-card__top">
                <div class="secom-product-card__icon">
                    <i data-lucide="${icon}"></i>
                </div>

                <span class="badge">
                    ${escapeHtml(secomV3GetCategoryLabel(product.categoria))}
                </span>
            </div>

            <div class="secom-product-card__name">
                ${escapeHtml(product.nombre)}
            </div>

            <div class="secom-product-card__meta">
                ${escapeHtml(product.codigo || '')}
                ${product.marca ? ` · ${escapeHtml(product.marca)}` : ''}
            </div>

            <div class="secom-product-card__specs">
                ${isPanel && product.watts
                    ? `<span>${formatNumber(product.watts)} W</span>`
                    : `<span>${escapeHtml(product.unidad || 'PZA')}</span>`
                }

                <span>
                    ${isPanel && secomV3UseGlobalPanelPrice()
                        ? `Global: ${formatCurrencyMXN(state.receipt.precioPanelInstalado || 10000)}`
                        : formatCurrencyMXN(product.precio || 0)
                    }
                </span>
            </div>

            <button class="btn btn--primary secom-product-card__btn" data-add-product="${escapeAttr(product.id)}">
                <i data-lucide="plus"></i>
                Agregar
            </button>
        </div>
    `;
}


function secomV3RenderShadowControls() {
    secomV3EnsureProductState();
    const pct = Number(state.quoteMeta?.perdidasSombraPct || state.receipt?.instalacion?.perdidasSombraPct || 0);
    const pctInt = Math.round(pct * 100);
    const options = [0, 5, 10, 15, 20, 25, 30];
    return `
        <div class="secom-shadow-panel">
            <div class="row" style="align-items:flex-start; gap:14px">
                <div class="secom-product-card__icon"><i data-lucide="cloud-sun"></i></div>
                <div style="flex:1; min-width:0">
                    <div class="card__title">Pérdidas por sombras</div>
                    <div class="help">Define el porcentaje estimado de sombras u obstrucciones. Este valor reduce la producción efectiva del sistema y ajusta cobertura, ahorro y retorno.</div>
                </div>
                <div class="badge">Actual: ${pctInt}%</div>
            </div>

            <div class="secom-shadow-options">
                ${options.map(value => `
                    <button type="button" class="secom-shadow-option ${value === pctInt ? 'is-active' : ''}" data-shadow-pct="${value}">
                        <span>${value === 0 ? 'Sin sombras' : value <= 10 ? 'Baja' : value <= 20 ? 'Media' : 'Alta'}</span>
                        <b>${value}%</b>
                    </button>
                `).join('')}
            </div>

            <div class="secom-shadow-range">
                <label>Porcentaje personalizado</label>
                <input id="shadowPctRange" type="range" min="0" max="60" step="1" value="${pctInt}" />
                <div class="help">Valor recomendado: 0% si no existen sombras relevantes. Usar 10%-20% para sombras parciales y mayor valor si hay obstrucciones significativas.</div>
            </div>
        </div>
    `;
}

function secomV3ApplyShadowPercent(value) {
    const pct = Math.max(0, Math.min(60, Number(value || 0))) / 100;
    state.quoteMeta.perdidasSombraPct = pct;
    state.quoteMeta.sombras = pct <= 0 ? 'Sin sombras' : `Sombras estimadas ${Math.round(pct * 100)}%`;
    if (!state.receipt) state.receipt = createEmptyReceiptData(state.selectedTariff);
    state.receipt.instalacion = {
        ...(state.receipt.instalacion || {}),
        perdidasSombraPct: pct,
        sombras: state.quoteMeta.sombras,
    };
    state.quote = secomV3CurrentQuote();
    renderWizard();
}

function secomV3GetPackageCards() {
    loadInsumoCatalogSafe(true);
    loadPackageCatalogSafe(true);
    const q = state.quote || currentStep2Quote();
    return (PACKAGE_PRESETS || [])
        .filter(pkg => pkg && pkg.activo !== false)
        .map(pkg => {
            const items = buildPackageItems(pkg.key, {
                quote: q,
                receipt: state.receipt,
                paneles: q.paneles,
                consumoMensual: q.consumoMensual
            });
            const subtotal = items.reduce((acc, it) => acc + (Number(it.cantidad || 0) * Number(it.precio || 0)), 0);
            const ivaPct = Number(state.receipt?.impuestosPct ?? 0.16);
            const total = subtotal * (1 + ivaPct);
            const selected = String(state.selectedPackage || '') === String(pkg.key);
            return `
                <div class="secom-product-card secom-package-product-card ${selected ? 'is-selected' : ''}">
                    <div class="secom-product-card__top">
                        <div class="secom-product-card__icon"><i data-lucide="boxes"></i></div>
                        <span class="badge">${escapeHtml(pkg.badge || 'Paquete')}</span>
                    </div>
                    <div class="secom-product-card__name">${escapeHtml(pkg.label || pkg.nombre || 'Paquete')}</div>
                    <div class="secom-product-card__meta">${escapeHtml(pkg.description || pkg.descripcion || '')}</div>
                    <div class="secom-product-card__specs">
                        <span>${items.length} insumos</span>
                        <span>${formatCurrencyMXN(total)}</span>
                    </div>
                    <button class="btn btn--primary secom-product-card__btn" data-add-package="${escapeAttr(pkg.key)}">
                        <i data-lucide="package-plus"></i>
                        Aplicar paquete
                    </button>
                </div>
            `;
        });
}

function secomV3RenderSelectedProduct(item, index) {
    const isPanel = secomV3IsPanelItem(item);

    return `
        <div class="secom-cart-item" data-cart-row="${index}">
            <div class="secom-cart-item__icon">
                <i data-lucide="${isPanel ? 'zap' : 'package'}"></i>
            </div>

            <div class="secom-cart-item__info">
                <div class="secom-cart-item__name">
                    ${escapeHtml(item.descripcion || 'Producto')}
                </div>

                <div class="secom-cart-item__meta">
                    ${escapeHtml(item.codigo || '')}
                    ${item.watts ? ` · ${formatNumber(item.watts)} W` : ''}
                    · ${formatCurrencyMXN(item.precio || 0)}
                </div>
            </div>

            <div class="secom-cart-item__qty">
                <button class="icon-btn icon-btn--sm" data-cart-minus="${index}" type="button">
                    <i data-lucide="minus"></i>
                </button>

                <input
                    class="tbl-input secom-cart-qty"
                    type="number"
                    min="1"
                    step="1"
                    value="${escapeAttr(String(item.cantidad || 1))}"
                    data-cart-qty="${index}"
                />

                <button class="icon-btn icon-btn--sm" data-cart-plus="${index}" type="button">
                    <i data-lucide="plus"></i>
                </button>
            </div>

            <div class="secom-cart-item__total">
                ${formatCurrencyMXN(Number(item.cantidad || 0) * Number(item.precio || 0))}
            </div>

            <button class="icon-btn icon-btn--sm" data-cart-remove="${index}" type="button">
                <i data-lucide="trash-2"></i>
            </button>
        </div>
    `;
}

function secomV2RenderPackageLeft() {
    secomV3EnsureProductState();

    const categories = ['paneles', 'inversores', 'montaje', 'proteccion', 'sombras', 'paquetes', 'otros'];
    const activeCategory = state.productCatalog?.activeCategory || 'paneles';
    const products = activeCategory === 'paquetes' || activeCategory === 'sombras' ? [] : (activeCategory === 'otros' ? secomV3GetCatalogProducts() : secomV3GetProductsByCategory(activeCategory));
    const packageCards = activeCategory === 'paquetes' ? secomV3GetPackageCards() : [];
    const shadowControls = activeCategory === 'sombras' ? secomV3RenderShadowControls() : '';
    const selected = state.receipt.insumos || [];
    const totals = secomV3ComputeTotals();
    const q = secomV3CurrentQuote();

    const coverage = q.porcentajeCobertura;
    const coverageText = coverage == null ? '0%' : `${coverage}%`;
    const coverageWidth = coverage == null ? 0 : Math.min(100, coverage);

    return `
        <div class="secom-products-full">

            <div class="secom-products-head">
                <div>
                    <div class="card__title">Catálogo de productos</div>
                    <div class="help">
                        Selecciona productos para construir el sistema solar personalizado.
                    </div>
                </div>
            </div>

            <div class="secom-category-tabs">
                ${categories.map(cat => `
                    <button
                        type="button"
                        class="secom-category-tab ${activeCategory === cat ? 'is-active' : ''}"
                        data-product-category="${cat}"
                    >
                        <i data-lucide="${secomV3GetProductIcon(cat)}"></i>
                        ${escapeHtml(secomV3GetCategoryLabel(cat))}
                    </button>
                `).join('')}
            </div>

            <div class="secom-product-catalog secom-product-catalog--wide">
                ${activeCategory === 'sombras'
                    ? shadowControls
                    : activeCategory === 'paquetes'
                        ? (packageCards.length ? packageCards.join('') : `<div class="review-alert" style="width:100%">No hay paquetes activos registrados.</div>`)
                        : (products.length
                            ? products.map(secomV3RenderProductCard).join('')
                            : `<div class="review-alert" style="width:100%">No hay productos registrados en esta categoría.</div>`)
                }
            </div>

            <div class="secom-commercial-card secom-commercial-card--below">
                <div class="secom-commercial-card__icon">
                    <i data-lucide="badge-dollar-sign"></i>
                </div>

                <div class="secom-commercial-card__body">
                    <div class="secom-commercial-card__title">
                        Precio global por panel instalado
                    </div>

                    <div class="help">
                        Si lo desactivas, cada panel usará el precio real registrado en el catálogo de insumos.
                    </div>

                    <label class="check-row" style="margin-top:10px; display:inline-flex; align-items:center; gap:8px; cursor:pointer;">
                        <input
                            id="usarPrecioGlobalPanel"
                            type="checkbox"
                            ${secomV3UseGlobalPanelPrice() ? 'checked' : ''}
                        />
                        <span>Usar precio global</span>
                    </label>
                </div>

                <div class="field secom-commercial-card__input">
                    <label>Precio por panel</label>
                    <input
                        id="precioPanelInstalado"
                        type="number"
                        min="0"
                        step="100"
                        value="${escapeAttr(String(state.receipt.precioPanelInstalado || 10000))}"
                        ${secomV3UseGlobalPanelPrice() ? '' : 'disabled'}
                    />
                </div>
            </div>

            <div class="secom-cart-panel">
                <div class="secom-cart-panel__head">
                    <div>
                        <div class="card__title">Productos seleccionados</div>
                        <div class="help">Productos e insumos seleccionados para la cotización actual.</div>
                    </div>

                    <button class="btn" id="btnAddManualProduct" type="button">
                        <i data-lucide="edit-3"></i>
                        Agregar manual
                    </button>
                </div>

                <div class="secom-cart-list">
                    ${selected.length
                        ? selected.map(secomV3RenderSelectedProduct).join('')
                        : `
                            <div class="empty" style="min-height:150px">
                                <div class="empty__icon"><i data-lucide="shopping-cart"></i></div>
                                <div class="card__title">No hay productos seleccionados</div>
                                <div class="help">Agrega productos o aplica un paquete completo desde el catálogo superior.</div>
                            </div>
                        `
                    }
                </div>
            </div>

            <div class="secom-summary-bottom">
                <div class="card__title">Resumen de cotización</div>

                <div class="secom-summary-bottom-grid">

                    <div class="secom-summary-figma secom-summary-figma--blue">
                        <div class="secom-summary-figma__label">
                            <i data-lucide="zap"></i>
                            Capacidad instalada
                        </div>
                        <div class="secom-summary-figma__value">
                            ${Number(q.kwp || 0).toFixed(2)}
                        </div>
                        <div class="secom-summary-figma__unit">kWp</div>
                    </div>

                    <div class="secom-summary-figma secom-summary-figma--green">
                        <div class="secom-summary-figma__label">
                            <i data-lucide="trending-up"></i>
                            Producción anual
                        </div>
                        <div class="secom-summary-figma__value">
                            ${formatNumber(q.produccionAnual || 0)}
                        </div>
                        <div class="secom-summary-figma__unit">kWh/año</div>
                    </div>

                    <div class="secom-summary-figma secom-summary-figma--purple">
                        <div class="secom-summary-figma__label">
                            <i data-lucide="battery"></i>
                            Cobertura
                        </div>
                        <div class="secom-summary-figma__value">
                            ${coverageText}
                        </div>
                        <div class="secom-summary-progress">
                            <span style="width:${coverageWidth}%"></span>
                        </div>
                    </div>

                    <div class="secom-summary-figma secom-summary-figma--yellow">
                        <div class="secom-summary-figma__label">
                            <i data-lucide="dollar-sign"></i>
                            Ahorro mensual
                        </div>
                        <div class="secom-summary-figma__value">
                            ${formatCurrencyMXN(q.ahorroMensual || 0)}
                        </div>
                        <div class="secom-summary-figma__unit">MXN/mes</div>
                    </div>

                    <div class="secom-summary-figma">
                        <div class="secom-summary-figma__label">
                            <i data-lucide="chart-no-axes-combined"></i>
                            Retorno estimado
                        </div>
                        <div class="secom-summary-figma__value">
                            ${q.retornoAnios || 0}
                        </div>
                        <div class="secom-summary-figma__unit">años</div>
                    </div>

                    <div class="secom-summary-figma secom-summary-figma--total">
                        <div class="secom-summary-figma__label">
                            Inversión total
                        </div>
                        <div class="secom-summary-figma__value">
                            ${formatCurrencyMXN(q.inversion || 0)}
                        </div>
                        <div class="secom-summary-figma__unit">MXN con IVA</div>
                    </div>

                </div>
            </div>

            <div class="insumos-summary" style="margin-top:16px">
                <div class="insumos-summary__row">
                    <span>Subtotal</span>
                    <b id="insSubtotal">${formatCurrencyMXN(totals.subtotal)}</b>
                </div>

                <div class="insumos-summary__row">
                    <span>IVA (%)</span>
                    <div style="display:flex; align-items:center; gap:10px">
                        <input
                            id="insTaxPct"
                            class="insumos-summary__tax"
                            type="number"
                            min="0"
                            max="30"
                            step="0.5"
                            value="${escapeAttr(String(Math.round(totals.ivaPct * 100 * 10) / 10))}"
                        />
                        <b id="insTaxes">${formatCurrencyMXN(totals.iva)}</b>
                    </div>
                </div>

                <div class="insumos-summary__row insumos-summary__row--total">
                    <span>Total</span>
                    <b id="insTotal">${formatCurrencyMXN(totals.total)}</b>
                </div>
            </div>

            <div class="wizard-actions" style="justify-content:space-between; margin-top:18px">
                <button class="btn" id="btnBackPackage">
                    <i data-lucide="arrow-left"></i>
                    Volver
                </button>

                <button class="btn btn--success" id="btnNextPackage">
                    <i data-lucide="arrow-right"></i>
                    Generar cotización
                </button>
            </div>

            <div class="help" id="msg2" style="margin-top:10px"></div>

        </div>
    `;
}

function secomV2RenderPackageRight() {
    return '';
}

function secomV3RefreshPackageUI() {
    const q = secomV3CurrentQuote();
    state.quote = q;

    const totals = secomV3ComputeTotals();

    $('#insSubtotal') && ($('#insSubtotal').textContent = formatCurrencyMXN(totals.subtotal));
    $('#insTaxes') && ($('#insTaxes').textContent = formatCurrencyMXN(totals.iva));
    $('#insTotal') && ($('#insTotal').textContent = formatCurrencyMXN(totals.total));
    $('#kpiQuoteTotal') && ($('#kpiQuoteTotal').textContent = formatCurrencyMXN(q.inversion || 0));
    $('#step2Kwp') && ($('#step2Kwp').textContent = `${Number(q.kwp || 0).toFixed(2)} kWp`);
    $('#step2Saving') && ($('#step2Saving').textContent = formatCurrencyMXN(q.ahorroMensual || 0));
    $('#summaryCobertura') && ($('#summaryCobertura').textContent = q.porcentajeCobertura == null ? 'Pendiente' : `${q.porcentajeCobertura}%`);
}

function secomV2WirePackage() {
    secomV3EnsureProductState();

    $('#btnBackPackage')?.addEventListener('click', () => {
        gotoStep(3);
    });

    $('#btnNextPackage')?.addEventListener('click', () => {
        const q = secomV3CurrentQuote();
        const msg = $('#msg2');

        if (!state.client.nombre) {
            state.client.nombre = state.receipt?.nombre || '';
        }

        if (!state.client.nombre) {
            if (msg) msg.textContent = 'Por favor, captura el nombre del cliente en la validación del recibo.';
            toast({
                title: 'Falta información',
                message: 'El nombre del cliente es obligatorio.',
                icon: 'alert-triangle'
            });
            return;
        }

        if (!state.receipt.servicio) {
            if (msg) msg.textContent = 'Por favor, captura el número de servicio del recibo.';
            toast({
                title: 'Dato crítico faltante',
                message: 'El número de servicio es obligatorio para continuar.',
                icon: 'alert-triangle'
            });
            return;
        }

        state.quote = q;
        gotoStep(5);
    });

    $$('#route-cotizador [data-product-category]').forEach(btn => {
        btn.addEventListener('click', () => {
            state.productCatalog = state.productCatalog || {};
            state.productCatalog.activeCategory = btn.dataset.productCategory || 'paneles';
            renderWizard();
        });
    });

    $$('#route-cotizador [data-shadow-pct]').forEach(btn => {
        btn.addEventListener('click', () => secomV3ApplyShadowPercent(btn.dataset.shadowPct));
    });

    $('#shadowPctRange')?.addEventListener('input', debounce((e) => {
        secomV3ApplyShadowPercent(e.target.value);
    }, 120));

    $('#usarPrecioGlobalPanel')?.addEventListener('change', () => {
        state.receipt.usarPrecioGlobalPanel = Boolean($('#usarPrecioGlobalPanel')?.checked);
        secomV3SyncPanelPrices();
        state.quote = secomV3CurrentQuote();
        renderWizard();
    });

    $('#precioPanelInstalado')?.addEventListener('input', debounce(() => {
        const precio = secomV2Number($('#precioPanelInstalado')?.value, 10000);
        state.receipt.precioPanelInstalado = Math.max(0, precio);

        if (secomV3UseGlobalPanelPrice()) {
            secomV3SyncPanelPrices();
        }

        state.quote = secomV3CurrentQuote();
        renderWizard();
    }, 180));

    $$('#route-cotizador [data-add-package]').forEach(btn => {
        btn.addEventListener('click', () => {
            const key = String(btn.dataset.addPackage || '');
            const pkg = (PACKAGE_PRESETS || []).find(p => String(p.key) === key);
            if (!pkg) {
                toast({title: 'Paquete no disponible', message: 'No se encontró el paquete seleccionado.', icon: 'alert-triangle'});
                return;
            }
            applyPackagePreset(key);
        });
    });

    $$('#route-cotizador [data-add-product]').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = String(btn.dataset.addProduct || '');
            const product = secomV3GetCatalogProducts().find(p => String(p.id) === id);

            if (!product) {
                return;
            }

            const isPanel = product.categoria === 'paneles';
            const precioFinal = isPanel
                ? secomV3ResolvePanelPrice(product)
                : Number(product.precio || 0);

            const existingIndex = state.receipt.insumos.findIndex(item => String(item.codigo) === String(product.codigo));

            if (existingIndex >= 0) {
                state.receipt.insumos[existingIndex].cantidad = Number(state.receipt.insumos[existingIndex].cantidad || 0) + 1;
                if (isPanel) {
                    state.receipt.insumos[existingIndex].precio = precioFinal;
                    state.receipt.insumos[existingIndex].watts = product.watts || state.receipt.insumos[existingIndex].watts || 0;
                }
            } else {
                state.receipt.insumos.push({
                    catalogId: product.id,
                    codigo: product.codigo,
                    descripcion: product.nombre,
                    cantidad: 1,
                    unidad: product.unidad || 'PZA',
                    precio: precioFinal,
                    impuestoPct: product.impuestoPct ?? 0.16,
                    categoria: product.categoria,
                    watts: product.watts || 0
                });
            }

            state.selectedPackage = '';
            if (state.receipt?.instalacion) {
                state.receipt.instalacion.paqueteSeleccionado = '';
                state.receipt.instalacion.paqueteLabel = '';
            }
            state.quote = secomV3CurrentQuote();

            renderWizard();

            toast({
                title: 'Producto agregado',
                message: product.nombre,
                icon: 'check-circle'
            });
        });
    });

    $('#btnAddManualProduct')?.addEventListener('click', () => {
        state.selectedPackage = '';
        if (state.receipt?.instalacion) {
            state.receipt.instalacion.paqueteSeleccionado = '';
            state.receipt.instalacion.paqueteLabel = '';
        }
        state.receipt.insumos.push({
            codigo: '',
            descripcion: 'Producto manual',
            cantidad: 1,
            unidad: 'PZA',
            precio: 0,
            impuestoPct: state.receipt.impuestosPct ?? 0.16,
            categoria: 'otros',
            watts: 0
        });

        state.quote = secomV3CurrentQuote();
        renderWizard();
    });

    $$('#route-cotizador [data-cart-plus]').forEach(btn => {
        btn.addEventListener('click', () => {
            const idx = Number(btn.dataset.cartPlus);
            if (!state.receipt.insumos[idx]) return;

            state.receipt.insumos[idx].cantidad = Number(state.receipt.insumos[idx].cantidad || 0) + 1;
            state.quote = secomV3CurrentQuote();
            renderWizard();
        });
    });

    $$('#route-cotizador [data-cart-minus]').forEach(btn => {
        btn.addEventListener('click', () => {
            const idx = Number(btn.dataset.cartMinus);
            if (!state.receipt.insumos[idx]) return;

            const next = Math.max(1, Number(state.receipt.insumos[idx].cantidad || 1) - 1);
            state.receipt.insumos[idx].cantidad = next;
            state.quote = secomV3CurrentQuote();
            renderWizard();
        });
    });

    $$('#route-cotizador [data-cart-qty]').forEach(input => {
        input.addEventListener('input', debounce(() => {
            const idx = Number(input.dataset.cartQty);
            if (!state.receipt.insumos[idx]) return;

            const qty = Math.max(1, secomV2Number(input.value, 1));
            state.receipt.insumos[idx].cantidad = qty;
            state.quote = secomV3CurrentQuote();
            renderWizard();
        }, 180));
    });

    $$('#route-cotizador [data-cart-remove]').forEach(btn => {
        btn.addEventListener('click', () => {
            const idx = Number(btn.dataset.cartRemove);
            if (!Number.isFinite(idx)) return;

            state.receipt.insumos.splice(idx, 1);
            state.quote = secomV3CurrentQuote();
            renderWizard();
        });
    });

    $('#insTaxPct')?.addEventListener('input', debounce(() => {
        const pctInput = secomV2Number($('#insTaxPct')?.value, 16);
        state.receipt.impuestosPct = Math.max(0, Math.min(30, pctInput)) / 100;
        state.quote = secomV3CurrentQuote();
        renderWizard();
    }, 180));

    window.lucide?.createIcons();
}
window.debugSECOM = () => {
    console.log('receipt:', state.receipt);
    console.log('historial:', state.receipt?.historial);
    console.log('consumoPeriodo:', state.receipt?.consumoPeriodo);
    console.log('metrics:', secomV2GetConsumptionMetrics());
};
init();