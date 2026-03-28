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
    const historialReal = Array.isArray(receipt?.historial)
            ? receipt.historial.some(x => Number(x?.kwh || 0) > 0 || Number(x?.pago || 0) > 0)
            : false;

    if (consumoPeriodoReal <= 0 || (totalReciboReal <= 0 && !historialReal)) {
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

    const consumoMensualBase = isBimestral ? (consumoPeriodo / 2) : consumoPeriodo;
    const ajusteKwhMes = Number(receipt?.ajusteConsumo?.kwhMes || 0);
    const consumoMensual = Math.max(0, consumoMensualBase + ajusteKwhMes);

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

    // Ahorro estimado: usando pago promedio y costo base del recibo cuando existe
    const pagoActual = Number(receipt?.totalAPagar || 0);
    const pagoProm = (() => {
        const pagos = (receipt?.historial || []).map(x => Number(x.pago || 0)).filter(n => n > 0);
        return pagos.length ? (pagos.reduce((a, b) => a + b, 0) / pagos.length) : pagoActual;
    })();

    const ahorroMensual = Math.max(0, Number(receipt?.ahorroEstimado || 0));
    const ahorroMensualEstimado = ahorroMensual > 0 ? ahorroMensual : Math.max(0, pagoProm * 0.65);

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
    const r = quote.receipt;
    const c = quote.client;
    const periodoRaw = r?.periodo?.raw || '';
    const fecha = new Date().toLocaleDateString('es-MX', {year: 'numeric', month: 'long', day: '2-digit'});

    const instal = r?.instalacion || {};
    const folio = (quote.id || '').replace(/</g, '') || '—';
    const tarifa = r?.tarifa || quote?.selectedTariff?.label || '-';

    const ajusteTxt = Number(quote.ajusteKwhMes || 0) !== 0
            ? `${formatNumber(quote.ajusteKwhMes)} kWh/mes`
            : '—';

    const perdidasPct = Math.round(Number(quote.perdidasSombraPct || instal.perdidasSombraPct || 0) * 100);
    const pagoConSolar = Math.max(0, Math.round(Number(quote.pagoProm || 0) - Number(quote.ahorroMensual || 0)));

    const esc = (s) => String(s ?? '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');

    const insumos = Array.isArray(r?.insumos) ? r.insumos : [];
    const showInsumos = insumos.length > 0;
    const subtotalIns = Number(quote.subtotalInsumos || 0);
    const impuestosIns = Number(quote.impuestosInsumos || 0);
    const totalIns = Number(quote.totalInsumos || 0);
    const ivaPct = Math.round((Number(quote.impuestosPct || r?.impuestosPct || 0.16) * 100) * 10) / 10;

    return `
    <div class="export-doc export-doc--formal" id="exportDoc">

      <div class="export-doc__band">
        <div class="export-doc__brand2">
          <img class="export-doc__logo2" src="assets/logo.png" alt="SECOM" />
          <div>
            <div class="export-doc__title">Cotización de Sistema Solar Fotovoltaico</div>
            <div class="export-doc__subtitle">Propuesta técnico-económica</div>
          </div>
        </div>
        <div class="export-doc__folio">
          <div class="export-doc__folioLabel">Folio</div>
          <div class="export-doc__folioValue">${folio}</div>
          <div class="export-doc__folioMeta">Fecha: ${fecha}</div>
        </div>
      </div>

      <div class="export-doc__section">
        <div class="export-doc__sectionTitle">Datos generales</div>

        <div class="export-doc__grid2">
          <div class="export-doc__box2">
            <div class="export-doc__label2">Cliente</div>
            <div class="export-doc__value2">${(c.nombre || r.nombre || '-')}</div>
            <div class="export-doc__meta2">${(c.telefono || '-')} · ${(c.email || '-')}</div>
            <div class="export-doc__meta2">${(c.direccion || r.direccion || '-')}</div>
          </div>
          <div class="export-doc__box2">
            <div class="export-doc__label2">Suministro</div>
            <div class="export-doc__row2"><span>No. de servicio</span><b>${r.servicio || '-'}</b></div>
            <div class="export-doc__row2"><span>Tarifa</span><b>${tarifa}</b></div>
            <div class="export-doc__row2"><span>Periodo</span><b>${periodoRaw || '-'}</b></div>
            <div class="export-doc__row2"><span>Titular del recibo</span><b>${(r.nombre || '-')}</b></div>
          </div>
        </div>
      </div>

      <div class="export-doc__section">
        <div class="export-doc__sectionTitle">Resumen ejecutivo</div>
        <div class="export-doc__kpis">
          <div class="export-doc__kpi">
            <div class="export-doc__kpiLabel">Potencia instalada</div>
            <div class="export-doc__kpiValue">${Number(quote.kwp || 0).toFixed(2)} kWp</div>
          </div>
          <div class="export-doc__kpi">
            <div class="export-doc__kpiLabel">Paneles</div>
            <div class="export-doc__kpiValue">${formatNumber(quote.paneles || 0)} × ${formatNumber(Number(quote.panelWatts || quote.params?.panelWatts || 0))} W</div>
          </div>
          <div class="export-doc__kpi">
            <div class="export-doc__kpiLabel">Generación anual estimada</div>
            <div class="export-doc__kpiValue">${formatNumber(quote.produccionAnual || 0)} kWh/año</div>
          </div>
          <div class="export-doc__kpi">
            <div class="export-doc__kpiLabel">Inversión total</div>
            <div class="export-doc__kpiValue">${formatCurrencyMXN(quote.inversion || 0)}</div>
          </div>
          <div class="export-doc__kpi">
            <div class="export-doc__kpiLabel">Ahorro mensual estimado</div>
            <div class="export-doc__kpiValue">${formatCurrencyMXN(quote.ahorroMensual || 0)}</div>
          </div>
          <div class="export-doc__kpi">
            <div class="export-doc__kpiLabel">Retorno estimado</div>
            <div class="export-doc__kpiValue">${quote.retornoAnios ? `${quote.retornoAnios} años` : '—'}</div>
          </div>
        </div>

        <div class="export-doc__mini">
          <div class="export-doc__miniRow">
            <span>Consumo mensual usado</span>
            <b>${formatNumber(quote.consumoMensual || 0)} kWh/mes</b>
          </div>
          <div class="export-doc__miniRow">
            <span>Ajuste de consumo</span>
            <b>${ajusteTxt}</b>
          </div>
          <div class="export-doc__miniRow">
            <span>Pago promedio CFE</span>
            <b>${formatCurrencyMXN(quote.pagoProm || 0)}</b>
          </div>
          <div class="export-doc__miniRow">
            <span>Pago estimado con solar</span>
            <b>${formatCurrencyMXN(pagoConSolar)}</b>
          </div>
        </div>
      </div>

      <div class="export-doc__section">
        <div class="export-doc__sectionTitle">Consumo vs. producción estimada</div>
        <div class="export-doc__sectionHelp">Histórico del recibo (últimos 12 periodos) comparado con la producción mensual estimada del sistema.</div>
        <div class="export-doc__chart">
          <canvas id="exportChart" height="160"></canvas>
        </div>
      </div>

      <div class="export-doc__section">
        <div class="export-doc__sectionTitle">Configuración del sistema propuesto</div>
        <div class="export-doc__grid2">
          <div class="export-doc__box2">
            <div class="export-doc__label2">Componentes</div>
            <div class="export-doc__row2"><span>Modelo de panel</span><b>${(instal.panelModelo || '—')}</b></div>
            <div class="export-doc__row2"><span>Dimensiones del panel</span><b>${(instal.panelDimensiones || '—')}</b></div>
            <div class="export-doc__row2"><span>Modelo de inversor</span><b>${(instal.inversorModelo || '—')}</b></div>
          </div>
          <div class="export-doc__box2">
            <div class="export-doc__label2">Condiciones de instalación</div>
            <div class="export-doc__row2"><span>Tipo de techo</span><b>${(instal.tipoTecho || '—')}</b></div>
            <div class="export-doc__row2"><span>Sombras</span><b>${(instal.sombras || '—')} ${Number.isFinite(perdidasPct) ? `(${perdidasPct}%)` : ''}</b></div>
            <div class="export-doc__row2"><span>Notas</span><b>${(instal.notasFisicas || '—')}</b></div>
          </div>
        </div>

        <div class="export-doc__assumptions">
          <div class="export-doc__assumpTitle">Supuestos de cálculo</div>
          <div class="export-doc__assumpGrid">
            <div><span>Producción promedio</span><b>${formatNumber(quote.yieldEfectivo || quote.params?.yieldKwhPerKwpMonth || 0)} kWh/kWp·mes</b></div>
            <div><span>Costo por kWp</span><b>${formatCurrencyMXN(quote.costPerKwpEff || quote.params?.costPerKwp || 0)}</b></div>
            <div><span>Contingencia</span><b>${Math.round((quote.params?.contingencyPct || 0) * 100)}%</b></div>
          </div>
        </div>
      </div>

      ${showInsumos ? `
      <div class="export-doc__section">
        <div class="export-doc__sectionTitle">Desglose de insumos</div>
        <div class="export-doc__sectionHelp">Detalle de materiales y servicios considerados en la propuesta.</div>

        <div style="border:1px solid rgba(0,0,0,0.14); border-radius:12px; overflow:hidden">
          <table style="width:100%; border-collapse:collapse; font-size:11px">
            <thead>
              <tr style="background:#f3f4f6">
                <th style="text-align:left; padding:10px; border-bottom:1px solid rgba(0,0,0,0.10)">CÓDIGO</th>
                <th style="text-align:left; padding:10px; border-bottom:1px solid rgba(0,0,0,0.10)">DESCRIPCIÓN</th>
                <th style="text-align:right; padding:10px; border-bottom:1px solid rgba(0,0,0,0.10)">CANTIDAD</th>
                <th style="text-align:left; padding:10px; border-bottom:1px solid rgba(0,0,0,0.10)">UNIDAD</th>
                <th style="text-align:right; padding:10px; border-bottom:1px solid rgba(0,0,0,0.10)">PRECIO</th>
                <th style="text-align:right; padding:10px; border-bottom:1px solid rgba(0,0,0,0.10)">TOTAL</th>
              </tr>
            </thead>
            <tbody>
              ${insumos.map(it => {
        const cant = Number(it?.cantidad || 0);
        const precio = Number(it?.precio || 0);
        const line = Math.round((Math.max(0, cant) * Math.max(0, precio)) * 100) / 100;
        return `
                  <tr>
                    <td style="padding:10px; border-bottom:1px solid rgba(0,0,0,0.06)">${esc(it?.codigo || '')}</td>
                    <td style="padding:10px; border-bottom:1px solid rgba(0,0,0,0.06)">${esc(it?.descripcion || '')}</td>
                    <td style="padding:10px; text-align:right; border-bottom:1px solid rgba(0,0,0,0.06)">${formatNumber(cant)}</td>
                    <td style="padding:10px; border-bottom:1px solid rgba(0,0,0,0.06)">${esc(it?.unidad || '')}</td>
                    <td style="padding:10px; text-align:right; border-bottom:1px solid rgba(0,0,0,0.06)">${formatCurrencyMXN(precio)}</td>
                    <td style="padding:10px; text-align:right; border-bottom:1px solid rgba(0,0,0,0.06)">${formatCurrencyMXN(line)}</td>
                  </tr>
                `;
    }).join('')}
            </tbody>
          </table>
        </div>

        <div style="display:flex; justify-content:flex-end; margin-top:10px">
          <div style="min-width:280px; border:1px solid rgba(0,0,0,0.14); border-radius:12px; padding:10px">
            <div style="display:flex; justify-content:space-between; gap:10px; font-size:12px"><span>Subtotal</span><b>${formatCurrencyMXN(subtotalIns)}</b></div>
            <div style="display:flex; justify-content:space-between; gap:10px; font-size:12px; margin-top:6px"><span>Impuestos (IVA ${ivaPct}%)</span><b>${formatCurrencyMXN(impuestosIns)}</b></div>
            <div style="display:flex; justify-content:space-between; gap:10px; font-size:13px; margin-top:8px"><span>Total</span><b>${formatCurrencyMXN(totalIns)}</b></div>
          </div>
        </div>
      </div>
      ` : ''}

      <div class="export-doc__foot2">
        <div><b>Vigencia:</b> 15 días naturales. <b>Observación:</b> La propuesta final está sujeta a verificación técnica en sitio y disponibilidad de equipo.</div>
        <div class="export-doc__footMeta">SECOM Energía Solar · Contacto: ${(c.telefono || '—')} · ${(c.email || '—')}</div>
      </div>

    </div>
  `;
}
