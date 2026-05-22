import { loadUserPreferences } from './preferences.js';
import { clamp, formatCurrencyMXN, formatNumber } from './utils.js';

export function computeQuote(receipt, client, params, overrides) {
    const defaults = {
        yieldKwhPerKwpMonth: 135, // producción promedio en México (aprox.)
        panelWatts: 550,
        costPerKwp: 22000, // MXN por kWp instalado
        contingencyPct: 0.06,
    };
    const p = {...defaults, ...(params || {})};
    const o = overrides || {};
    const consumoPeriodoReal = Number(receipt?.consumoPeriodo || 0);
    const totalReciboReal = Number(receipt?.totalAPagar || 0);
    const consumoPromedioReciboReal = Number(receipt?.consumoPromedioMensual || 0);
    const pagoPromedioReciboReal = Number(receipt?.pagoPromedioMensual || 0);
    const historialReal = Array.isArray(receipt?.historial)
            ? receipt.historial.some(x => Number(x?.kwh || 0) > 0 || Number(x?.pago || 0) > 0)
            : false;

    if ((consumoPeriodoReal <= 0 && consumoPromedioReciboReal <= 0) || (totalReciboReal <= 0 && pagoPromedioReciboReal <= 0 && !historialReal)) {
        return {
            client: client || {},
            receipt: receipt || {},
            params: p,
            consumoMensual: 0,
            consumoMensualBase: 0,
            ajusteKwhMes: Number(receipt?.ajusteConsumo?.kwhMes || 0),
            kwp: 0,
            panelesAuto: 0,
            paneles: 0,
            panelWatts: Number(o?.panelWatts ?? p.panelWatts),
            inversion: 0,
            inversionBase: 0,
            subtotalInsumos: 0,
            impuestosInsumos: 0,
            totalInsumos: 0,
            impuestosPct: Number(receipt?.impuestosPct ?? 0.16),
            pagoProm: 0,
            ahorroMensual: 0,
            retornoAnios: 0,
            produccionMensual: 0,
            produccionAnual: 0,
            yieldEfectivo: Math.round(p.yieldKwhPerKwpMonth),
            costPerKwpEff: p.costPerKwp,
            perdidasSombraPct: 0,
            incomplete: true,
        };
    }
    const consumoPeriodo = Number(receipt?.consumoPeriodo || 0);
    const periodoDays = Number(receipt?.periodo?.days || 0);
    const isBimestral = periodoDays >= 45 || receipt?.tipoPeriodo === 'Bimestral';

    const consumoPromedioRecibo = Number(receipt?.consumoPromedioMensual || 0);
    const consumoMensualBase = consumoPromedioRecibo > 0 ? consumoPromedioRecibo : (isBimestral ? (consumoPeriodo / 2) : consumoPeriodo);
    const ajusteKwhMes = Number(receipt?.ajusteConsumo?.kwhMes || 0);
    const consumoMensualManual = Number(o?.consumoMensual || 0);
    const consumoMensual = consumoMensualManual > 0 ? Math.max(0, consumoMensualManual) : Math.max(0, consumoMensualBase + ajusteKwhMes);

    // Consideraciones físicas (simulación): sombras reducen producción efectiva, techo puede afectar costo
    const perdidasSombraPct = clamp(Number(receipt?.instalacion?.perdidasSombraPct || 0), 0, 0.6);
    const yieldEfectivo = Math.max(1, p.yieldKwhPerKwpMonth * (1 - perdidasSombraPct));

    const techo = String(receipt?.instalacion?.tipoTecho || 'No especificado');
    const roofExtra = ({
        'Losa': 0,
        'Lámina': 0.03,
        'Teja': 0.06,
        'Otro': 0.04,
    })[techo] ?? 0;
    const costPerKwpEff = Math.round(p.costPerKwp * (1 + roofExtra));

    const kwp = clamp(consumoMensual / yieldEfectivo, 0.5, 60);
    const kwpRedondeado = Math.ceil(kwp * 2) / 2; // a 0.5kWp

    const panelWatts = Number(o?.panelWatts ?? p.panelWatts);
    const panelKw = panelWatts / 1000;
    const panelesAuto = Math.max(1, Math.ceil(kwpRedondeado / panelKw));
    const paneles = Number.isFinite(Number(o?.paneles)) && Number(o.paneles) > 0 ? Math.round(Number(o.paneles)) : panelesAuto;
    const kwpFinal = paneles * panelKw;

    const inversionBase = kwpFinal * costPerKwpEff;
    let inversion = Math.round(inversionBase * (1 + p.contingencyPct));

    // Ahorro estimado mensual. Para calcular retorno se usa un pago mensual realista:
    // 1) promedio mensual ya extraído del recibo, o 2) historial + recibo actual normalizados a mes.
    const pagoActualPeriodo = Number(receipt?.totalAPagar || 0);
    const normalizaPagoMensual = (value) => isBimestral ? (Number(value || 0) / 2) : Number(value || 0);
    const pagosHistoricos = (Array.isArray(receipt?.historial) ? receipt.historial : [])
        .map(x => Number(x?.pago || x?.importe || 0))
        .filter(n => Number.isFinite(n) && n > 0)
        .map(normalizaPagoMensual);
    const pagosMensuales = pagoActualPeriodo > 0
        ? [...pagosHistoricos, normalizaPagoMensual(pagoActualPeriodo)]
        : pagosHistoricos;
    const pagoPromedioHistorial = pagosMensuales.length
        ? Math.max(0, pagosMensuales.reduce((a, b) => a + b, 0) / pagosMensuales.length)
        : 0;
    const pagoPromedioRecibo = Math.max(0, Number(receipt?.pagoPromedioMensual || 0));
    const pagoProm = Math.max(pagoPromedioRecibo, pagoPromedioHistorial);

    const coberturaDisenio = consumoMensual > 0 ? clamp((kwpFinal * yieldEfectivo) / consumoMensual, 0, 1) : 0;
    const ahorroRecibo = Math.max(0, normalizaPagoMensual(Number(receipt?.ahorroEstimado || 0)));
    const ahorroPorCobertura = Math.max(0, pagoProm * coberturaDisenio);
    const ahorroMensualEstimado = Math.min(
        pagoProm || Math.max(ahorroPorCobertura, ahorroRecibo),
        Math.max(ahorroPorCobertura, ahorroRecibo)
    );

    // Si existe un desglose de insumos con precios, se usa como total de inversión
    const insumos = Array.isArray(receipt?.insumos) ? receipt.insumos : [];
    const impuestosPct = clamp(Number(receipt?.impuestosPct ?? 0.16), 0, 0.30);
    const subtotalInsumos = insumos.reduce((acc, it) => {
        const cant = Number(it?.cantidad || 0);
        const precio = Number(it?.precio || 0);
        if (!Number.isFinite(cant) || !Number.isFinite(precio))
            return acc;
        return acc + Math.max(0, cant) * Math.max(0, precio);
    }, 0);
    const hasInsumos = subtotalInsumos > 0;
    const impuestosInsumos = subtotalInsumos * impuestosPct;
    const totalInsumos = subtotalInsumos + impuestosInsumos;
    if (hasInsumos) {
        inversion = Math.round(totalInsumos);
    }

    const retornoAnios = ahorroMensualEstimado > 0 ? (inversion / (ahorroMensualEstimado * 12)) : 0;

    return {
        client: client || {},
        receipt: receipt || {},
        params: p,
        consumoMensual: Math.round(consumoMensual),
        consumoMensualBase: Math.round(consumoMensualBase),
        consumoMensualManual: consumoMensualManual > 0 ? Math.round(consumoMensualManual) : 0,
        ajusteKwhMes: Math.round(ajusteKwhMes),
        kwp: kwpFinal,
        panelesAuto,
        paneles,
        panelWatts,
        inversion,
        inversionBase: Math.round(inversionBase),
        subtotalInsumos: Math.round(subtotalInsumos * 100) / 100,
        impuestosInsumos: Math.round(impuestosInsumos * 100) / 100,
        totalInsumos: Math.round(totalInsumos * 100) / 100,
        impuestosPct,
        pagoProm,
        ahorroMensual: Math.round(ahorroMensualEstimado),
        retornoAnios: Math.round(retornoAnios * 10) / 10,
        produccionMensual: Math.round(kwpFinal * yieldEfectivo),
        produccionAnual: Math.round(kwpFinal * yieldEfectivo * 12),
        yieldEfectivo: Math.round(yieldEfectivo),
        costPerKwpEff,
        perdidasSombraPct,
    };
}

export function buildExportHtml(quote) {
    const r = quote.receipt || {};
    const c = quote.client || {};
    const prefs = loadUserPreferences();
    const fecha = new Date().toLocaleDateString('es-MX', {weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'});
    const esc = (value) => String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');

    const instal = r.instalacion || {};
    const insumos = Array.isArray(r.insumos) ? r.insumos : [];
    const subtotal = Number(quote.subtotalInsumos || insumos.reduce((acc, it) => acc + Number(it.cantidad || 0) * Number(it.precio || 0), 0));
    const ivaPct = Number(quote.impuestosPct ?? r.impuestosPct ?? 0.16);
    const iva = Number(quote.impuestosInsumos || (subtotal * ivaPct));
    const total = Number(quote.inversion || quote.totalInsumos || subtotal + iva);
    const anticipo = Math.round(total * 0.70);
    const finiquito = Math.round(total * 0.30);

    const produccionDiariaWh = (Number(quote.produccionMensual || 0) / 30) * 1000;
    const consumoDiarioWh = (Number(quote.consumoMensual || r.consumoPromedioMensual || 0) / 30) * 1000;
    const cobertura = consumoDiarioWh > 0 ? Math.round((produccionDiariaWh / consumoDiarioWh) * 100) : 0;
    const wattsInstalados = Number(quote.wattsInstalados || (Number(quote.kwp || 0) * 1000));
    const areaPanel = Math.round(Number(quote.paneles || 0) * 3.8);
    const packageLabel = instal.paqueteLabel || instal.paqueteSeleccionado || 'Configuración personalizada';
    const prefCompanyName = String(prefs?.company?.companyName || '').trim();
    const companyName = prefCompanyName && prefCompanyName !== 'SECOM Energía Solar' ? prefCompanyName : 'SOLUCIONES EN SEGURIDAD Y ENERGÍA';
    const advisor = prefs?.company?.advisorName || 'ING. JOSÉ RAMÓN DÍAZ GAXIOLA';
    const role = prefs?.company?.advisorRole || 'GERENTE OPERATIVO';
    const tarifa = r.tarifa || quote?.selectedTariff?.label || '—';
    const paneles = Number(quote.paneles || 0);
    const retorno = Number(quote.retornoAnios || 0);
    const ahorroMensual = Number(quote.ahorroMensual || 0);
    const direccionEmpresa = 'Calle Coahuila # 2008-Local 2, Cortinas 3ra. Sección, C. P. 85169, Ciudad Obregón, Sonora.';

    const lineTotal = (items) => items.reduce((acc, it) => acc + Math.max(0, Number(it.cantidad || 0)) * Math.max(0, Number(it.precio || 0)), 0);
    const normalizeText = value => String(value || '').toLowerCase();
    const by = (predicate) => insumos.filter(it => predicate({
        codigo: String(it.codigo || '').toUpperCase(),
        categoria: normalizeText(it.categoria),
        descripcion: normalizeText(it.descripcion),
        item: it,
    }));

    const panelItems = by(x => x.codigo.startsWith('PANEL') || x.categoria.includes('panel') || x.descripcion.includes('panel'));
    const inverterItems = by(x => x.codigo.startsWith('INV') || x.categoria.includes('inversor') || x.descripcion.includes('inversor'));
    const mountingItems = by(x => x.codigo.startsWith('EST') || x.categoria.includes('estructura') || x.categoria.includes('montaje') || x.descripcion.includes('estructura') || x.descripcion.includes('montaje'));
    const laborItems = by(x => x.codigo.includes('MO') || x.categoria.includes('instalación') || x.categoria.includes('instalacion') || x.descripcion.includes('mano de obra'));
    const usedCodes = new Set([...panelItems, ...inverterItems, ...mountingItems, ...laborItems].map(it => String(it.codigo || '').toUpperCase()));
    const electricalItems = insumos.filter(it => !usedCodes.has(String(it.codigo || '').toUpperCase()));

    const rows = [];
    const pushRow = (descripcion, cantidad, unidad, totalLinea) => {
        if (!totalLinea && !cantidad) return;
        const index = rows.length + 1;
        rows.push({
            parte: `1.${index}0`,
            descripcion,
            cantidad: cantidad || 1,
            unidad: unidad || 'LOTE',
            total: totalLinea || 0,
        });
    };

    if (panelItems.length) {
        const qty = panelItems.reduce((acc, it) => acc + Number(it.cantidad || 0), 0) || paneles || 1;
        const desc = panelItems[0]?.descripcion || `Módulo solar ${quote.panelWatts || 550} W monocristalino`;
        pushRow(desc, qty, panelItems[0]?.unidad || 'LOTE', lineTotal(panelItems));
    }
    if (inverterItems.length) {
        pushRow(inverterItems[0]?.descripcion || 'Inversor para interconexión a CFE con salida 220 Vca, módulo WiFi incluido', inverterItems.length, 'LOTE', lineTotal(inverterItems));
    }
    if (mountingItems.length) {
        pushRow('Montaje extra largo de aluminio anodizado para techo o piso de concreto, de alta resistencia a climas extremos y rápida instalación para arreglo de módulos fotovoltaicos.', 1, 'LOTE', lineTotal(mountingItems));
    }
    if (electricalItems.length) {
        pushRow('Suministro e instalación de material eléctrico de acuerdo con la NOM-001-SEDE vigente. Incluye canalización, protecciones, cableado, sistema de protecciones lado corriente alterna y lado corriente directa.', 1, 'LOTE', lineTotal(electricalItems));
    }
    if (laborItems.length) {
        pushRow('Servicio de mano de obra calificada. Técnicos con constancias para trabajos eléctricos y trabajos en alturas, certificación, ingeniería especializada, logística, fletes, maniobras, acarreos y limpieza en general.', 1, 'LOTE', lineTotal(laborItems));
    }
    if (!rows.length) {
        pushRow(`Sistema fotovoltaico de ${formatNumber(paneles || 0)} paneles solares, inversor, protecciones, montaje e instalación.`, 1, 'LOTE', subtotal || total);
    }
    const rowsSubtotal = rows.reduce((acc, row) => acc + Number(row.total || 0), 0);
    const diff = subtotal - rowsSubtotal;
    if (Math.abs(diff) >= 1 && rows.length < 5) {
        pushRow('Accesorios, materiales complementarios, ajustes de instalación y preparación para interconexión.', 1, 'LOTE', diff);
    }
    const tableRows = rows.slice(0, 5);

    const ahorroAnual = Math.max(0, ahorroMensual * 12);
    const chartW = 760;
    const chartH = 118;
    const left = 42;
    const right = 12;
    const top = 12;
    const bottom = 28;
    const plotW = chartW - left - right;
    const plotH = chartH - top - bottom;
    const years = Array.from({length: 25}, (_, i) => i + 1);
    const maxAhorro = Math.max(total, ahorroAnual * 25, 1);
    const scaleY = v => top + plotH - (Math.max(0, v) / maxAhorro) * plotH;
    const scaleX = i => left + (i / 24) * plotW;
    const ahorroPoints = years.map((y, i) => `${scaleX(i).toFixed(1)},${scaleY(ahorroAnual * y).toFixed(1)}`).join(' ');
    const inversionPoints = years.map((_, i) => `${scaleX(i).toFixed(1)},${scaleY(total).toFixed(1)}`).join(' ');
    const gridLines = Array.from({length: 7}, (_, i) => {
        const y = top + (plotH / 6) * i;
        const val = Math.round((maxAhorro * (6 - i)) / 6 / 1000);
        return `<line x1="${left}" y1="${y.toFixed(1)}" x2="${chartW - right}" y2="${y.toFixed(1)}"/><text x="${left - 7}" y="${(y + 3).toFixed(1)}" text-anchor="end">$${val}</text>`;
    }).join('');
    const xTicks = years.map((y, i) => `<text x="${scaleX(i).toFixed(1)}" y="${chartH - 8}" text-anchor="middle">${y}</text>`).join('');
    const chartSvg = `
        <svg class="secom-export-chart" viewBox="0 0 ${chartW} ${chartH}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Ahorro acumulado e inversión">
            <g class="grid">${gridLines}</g>
            <line class="axis" x1="${left}" y1="${top + plotH}" x2="${chartW - right}" y2="${top + plotH}"/>
            <line class="axis" x1="${left}" y1="${top}" x2="${left}" y2="${top + plotH}"/>
            <polyline class="saving" points="${ahorroPoints}"/>
            <polyline class="investment" points="${inversionPoints}"/>
            <g class="ticks">${xTicks}</g>
            <text class="axis-title" x="10" y="${top + 50}" transform="rotate(-90 10 ${top + 50})">MM de Pesos</text>
        </svg>`;

    const materialElectrico = 'Suministro e instalación de material eléctrico de acuerdo con la NOM-001-SEDE vigente. Incluye canalización, protecciones, cableado, sistema de protecciones lado corriente alterna y lado corriente directa.';
    const garantia = 'Garantía: 12 años en módulos, 12 años en inversores y 1 año en instalación eléctrica.';
    const nota = `Nota: Esta cotización es aproximada conforme a los datos proporcionados por el cliente y no constituye ningún compromiso hasta la firma del contrato. Su inversión total será de ${formatCurrencyMXN(total)}. La recuperación estimada de inversión es de ${retorno ? `${formatNumber(retorno)} años` : '—'}, considerando un ahorro mensual estimado de ${formatCurrencyMXN(ahorroMensual)}.`;

    return `
    <div class="export-doc export-doc--secom-template" id="exportDoc">
        <header class="secom-template-header">
            <div class="secom-template-logo"><img src="assets/logo.png" alt="SECOM" /></div>
            <div class="secom-template-title">
                <h1>${esc(companyName)}</h1>
            </div>
            <div class="secom-template-date">${esc(fecha)}</div>
        </header>

        <section class="secom-template-meta">
            <div class="secom-template-client">
                <div class="secom-template-client-title">Datos del cliente</div>
                <div class="secom-template-client-grid">
                    <div class="row"><span>Nombre:</span><b>${esc(c.nombre || r.nombre || '—')}</b></div>
                    <div class="row"><span>Dirección:</span><b>${esc(c.direccion || r.direccion || '—')}</b></div>
                    <div class="row"><span>Número de Servicio (RPU):</span><b>${esc(r.servicio || '—')}</b></div>
                    <div class="row"><span>Tarifa:</span><b>${esc(tarifa)}</b></div>
                </div>
            </div>
            <div class="secom-template-kpis">
                <div class="kpi"><b>Producción diaria de energía</b><span>${formatNumber(produccionDiariaWh)}</span></div>
                <div class="kpi"><b>% Producción vs Consumo</b><span>${formatNumber(cobertura)}%</span></div>
                <div class="kpi"><b>Consumo promedio diario de energía</b><span>${formatNumber(consumoDiarioWh)}</span></div>
                <div class="kpi"><b>Watts instalados</b><span>${formatNumber(wattsInstalados)}</span></div>
            </div>
        </section>

        <table class="secom-template-table">
            <colgroup><col class="col-part"/><col class="col-desc"/><col class="col-qty"/><col class="col-unit"/></colgroup>
            <thead><tr><th>Parte</th><th>Descripción</th><th>Cantidad</th><th>Unidad</th></tr></thead>
            <tbody>
                <tr class="system-row"><td></td><td>SISTEMA FOTOVOLTAICO</td><td></td><td></td></tr>
                ${tableRows.map(row => `<tr>
                    <td>${esc(row.parte)}</td>
                    <td>${esc(row.descripcion)}</td>
                    <td class="num">${formatNumber(row.cantidad)}</td>
                    <td class="unit">${esc(row.unidad || 'LOTE')}</td>
                </tr>`).join('')}
            </tbody>
        </table>

        <section class="secom-template-finance">
            <div class="payment-title">Condiciones de pago<br/>al contado:</div>
            <div class="payment-box">
                <div><b>70% DE ANTICIPO:</b><span>${formatCurrencyMXN(anticipo)}</span></div>
                <div><b>30% AL FINALIZAR:</b><span>${formatCurrencyMXN(finiquito)}</span></div>
                <div class="total"><b>Total:</b><span>${formatCurrencyMXN(total)}</span></div>
            </div>
            <div class="investment-title">Su inversión total<br/>será de:</div>
            <div class="investment-box">
                <div><b>Pesos:</b><span>${formatCurrencyMXN(subtotal)}</span></div>
                <div><b>IVA:</b><span>${formatCurrencyMXN(iva)}</span></div>
                <div class="total"><b>Total:</b><span>${formatCurrencyMXN(total)}</span></div>
            </div>
            <div class="roi-box"><b>Recuperación de la inversión</b><span>${retorno ? formatNumber(retorno) : '—'} años</span></div>
        </section>

        <section class="secom-template-chart-wrap">
            ${chartSvg}
            <div class="chart-legend"><span class="green"></span>Ahorro Acumulado <span class="blue"></span>Inversión</div>
            <div class="panel-warranty">Los paneles cuentan con 25 años de garantía de rendimiento lineal</div>
        </section>

        <section class="secom-template-install">
            <div class="bar"><b>Material de Instalación Paneles</b><span>Área Aproximada de instalación</span></div>
            <div class="row"><b>Accesorios en Aluminio Anodizado</b><span>${formatNumber(areaPanel)} m2</span></div>
            <div class="bar"><b>Instalación Eléctrica</b></div>
            <div class="row">De acuerdo a las normas eléctricas de CFE</div>
            <div class="bar"><b>Sistema de Fijación de Paneles</b></div>
            <div class="row">Fijación en techo o piso con protección de pasta epóxica a prueba de filtrado de agua</div>
        </section>

        <section class="secom-template-notes">
            <p><b>Tiempo estimado de entrega:</b> de 1 a 2 semanas</p>
            <p><b>Tiempo estimado de instalación:</b> Varía de acuerdo a dimensión del sistema</p>
            <p><b>${esc(garantia)}</b></p>
            <p><b>${esc(nota)}</b></p>
            <p>Tiene vigencia de 05 días a partir de la fecha establecida en esta cotización. Los trámites administrativos ante la CFE no tienen costo. Las cuotas extras de la dependencia de CFE no están incluidos en esta cotización. Los trabajos, materiales o viáticos no mencionados, serán cotizados de acuerdo a las instalaciones.</p>
            <p class="tax">LA LEY DE IMPUESTO SOBRE LA RENTA, ARTÍCULO 40 - 12 SECCIÓN 2, MARCA QUE LOS PORCENTAJES MÁXIMOS AUTORIZADOS TRATÁNDOSE DE ACTIVOS FIJOS POR TIPO DE BIEN, ENERGÍAS ALTERNATIVAS (SOLAR) PODRÁN SER DEDUCIDAS AL 100%</p>
        </section>

        <footer class="secom-template-footer">
            <div class="signature">
                <div>ATENTAMENTE</div>
                <span></span>
                <b>${esc(advisor)}</b>
                <small>${esc(role)}</small>
            </div>
            <div class="cert-badge"><img src="assets/sello-certificado.png" alt="Instalador certificado SECOM" /></div>
            <div class="footer-address">${esc(direccionEmpresa)}</div>
        </footer>
    </div>`;
}
