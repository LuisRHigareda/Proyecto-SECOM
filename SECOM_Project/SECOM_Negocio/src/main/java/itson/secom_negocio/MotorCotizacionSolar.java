/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_negocio;

import itson.secom_domain.DatosReciboCFE;
import itson.secom_domain.ParametrosSistema;
import itson.secom_domain.ProductoCantidad;
import itson.secom_domain.ResultadoCalculoCotizacion;
import itson.secom_domain.enumeradores.CategoriaProducto;
import itson.secom_domain.enumeradores.TipoTarifa;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author Serva
 */
public class MotorCotizacionSolar {
//
//    private static final double POTENCIA_PANEL_KWP = 0.550;
//    private static final double HSP_DIARIAS = 5.5;
//    private static final double FACTOR_RENDIMIENTO = 0.80;
//    private static final double FACTOR_CO2_KG_KWH = 0.423;
//    private static final double ABSORCION_ARBOL_KG = 20.0;
//    private static final int ANOS_PROYECCION = 25;
//
//    public ResultadoCalculoCotizacion calcular(DatosReciboCFE datos) throws Exception {
//        validar(datos);
//
//        ResultadoCalculoCotizacion r = new ResultadoCalculoCotizacion();
//        r.setNombreCliente(datos.getNombre());
//        r.setDireccion(datos.getDireccion());
//        r.setNoServicio(datos.getNoServicio());
//        r.setTarifa(datos.getTarifa());
//        r.setTipoTarifa(datos.getTipoTarifa());
//        r.setNoHilos(datos.getNoHilos());
//        r.setEsBimestral(datos.getTipoTarifa() != null
//                ? datos.getTipoTarifa().isEsBimestral()
//                : datos.getDuracionDias() >= 45);
//
//        double consumoMensual = calcularConsumoMensual(datos, r.isEsBimestral());
//        r.setConsumoPromedioMensualKwh(consumoMensual);
//
//        double pagoProm = calcularPagoProm(datos, r.isEsBimestral());
//        r.setPagoPromedioCFE(pagoProm);
//
//        double costoBase = calcularCostoBase(datos, r.isEsBimestral());
//        r.setCostoBaseConSolar(costoBase);
//
//        r.setAhorroMensualEstimado(Math.max(0, pagoProm - costoBase));
//        r.setPagoEstimadoConSolar(costoBase);
//
//        dimensionar(r, consumoMensual);
//        calcularImpacto(r);
//
//        return r;
//    }
//
//    private double calcularConsumoMensual(DatosReciboCFE d, boolean bimestral) {
//        List<Double> consumos = d.getConsumoHistoricos();
//        if (consumos == null || consumos.isEmpty()) {
//            double base = d.getConsumoActualKwh();
//            return bimestral ? base / 2.0 : base;
//        }
//        double suma = 0;
//        int n = 0;
//        for (Double c : consumos) {
//            if (c != null && c > 0) {
//                suma += c;
//                n++;
//            }
//        }
//        // FUERA del for
//        if (n == 0) {
//            return bimestral ? d.getConsumoActualKwh() / 2.0 : d.getConsumoActualKwh();
//        }
//        double prom = suma / n;
//        return bimestral ? prom / 2.0 : prom;
//    }
//
//    private double calcularPagoProm(DatosReciboCFE d, boolean bimestral) {
//        if (d.esIndustrial()) {
//            return calcularPagoPromIndustrial(d, bimestral);
//        }
//        return calcularPagoPromDomestica(d, bimestral);
//    }
//
//    private double calcularPagoPromDomestica(DatosReciboCFE d, boolean bimestral) {
//        List<Double> pagos = d.getPagosHistoricos();
//        if (pagos == null || pagos.isEmpty()) {
//            return d.getPagoActual();
//        }
//        double suma = 0;
//        int n = 0;
//        for (Double p : pagos) {
//            if (p != null && p > 0) {
//                suma += p;
//                n++;
//            }
//        }
//        // FUERA del for
//        if (n == 0) {
//            return d.getPagoActual();
//        }
//        double prom = suma / n;
//        return bimestral ? prom / 2.0 : prom;
//    }
//
//    private double calcularPagoPromIndustrial(DatosReciboCFE d, boolean bimestral) {
//        List<Double> precios = d.getPreciosMedios();
//        List<Double> consumos = d.getConsumoHistoricos();
//        if (precios == null || precios.isEmpty() || consumos == null || consumos.isEmpty()) {
//            return d.getPagoActual();
//        }
//
//        int n = Math.min(precios.size(), consumos.size());
//        double sumaKwh = 0, sumaPrecio = 0;
//        int count = 0;
//        for (int i = 0; i < n; i++) {
//            Double kwh = consumos.get(i), precio = precios.get(i);
//            if (kwh != null && kwh > 0 && precio != null && precio > 0) {
//                sumaKwh += kwh;
//                sumaPrecio += precio;
//                count++;
//            }
//        }
//        if (count == 0) {
//            return d.getPagoActual();
//        }
//        double consumoMensualProm = bimestral ? (sumaKwh / count) / 2.0 : sumaKwh / count;
//        return consumoMensualProm * (sumaPrecio / count);
//    }
//
//    private double calcularCostoBase(DatosReciboCFE d, boolean bimestral) {
//        double costoBase;
//        if (d.getCostoSuministro() > 0) {
//            costoBase = d.getCostoSuministro() * (1.0 + d.getIvaPorcentaje() / 100.0) + d.getCostoDAP();
//        } else {
//            costoBase = calcularPagoProm(d, bimestral) * 0.15;
//        }
//        return bimestral ? costoBase / 2.0 : costoBase;
//    }
//
//    private void dimensionar(ResultadoCalculoCotizacion r, double consumoMensual) {
//        if (consumoMensual <= 0) {
//            return;
//        }
//        double genPorPanelMes = POTENCIA_PANEL_KWP * HSP_DIARIAS * 30.0 * FACTOR_RENDIMIENTO;
//        int paneles = Math.max(1, (int) Math.ceil(consumoMensual / genPorPanelMes));
//
//        double wattsInstalados = paneles * POTENCIA_PANEL_KWP * 1000.0;
//        double kwp = wattsInstalados / 1000.0;
//        double genMensual = paneles * genPorPanelMes;
//        double genAnual = genMensual * 12.0;
//        double cobertura = consumoMensual > 0
//                ? Math.min(100.0, (genMensual / consumoMensual) * 100.0) : 0;
//        double retorno = r.getAhorroMensualEstimado() > 0
//                ? (kwp * 22000.0) / (r.getAhorroMensualEstimado() * 12.0) : 0;
//
//        r.setNumeroPaneles(paneles);
//        r.setPotenciaInstaladaKwp(kwp);
//        r.setWattsInstalados(wattsInstalados);
//        r.setGeneracionMensualEstimadaKwh(genMensual);
//        r.setGeneracioAnualEstimadaKwh(genAnual);
//        r.setProduccionDiariaEstimada(genMensual / 30.0);
//        r.setPorcentajCobertura(cobertura);
//        r.setRetornoInversion(retorno);
//    }
//
//    private void calcularImpacto(ResultadoCalculoCotizacion r) {
//        double genAnual = r.getGeneracioAnualEstimadaKwh();
//        if (genAnual <= 0) {
//            return;
//        }
//        double co2AnioKg = genAnual * FACTOR_CO2_KG_KWH;
//        r.setCo2EvitadoToneladas25años(
//                Math.round((co2AnioKg / 1000.0) * ANOS_PROYECCION * 10.0) / 10.0);
//        r.setArbolesEquivalentes25Años(
//                (int) (co2AnioKg / ABSORCION_ARBOL_KG * ANOS_PROYECCION));
//    }
//
//    private void validar(DatosReciboCFE d) throws Exception {
//        if (d == null) {
//            throw new Exception("Los datos del recibo no pueden ser nulos.");
//        }
//        if (d.getTipoTarifa() == null) {
//            throw new Exception("Se necesita especificar el tipo de tarifa.");
//        }
//        boolean tieneConsumo = d.getConsumoActualKwh() > 0
//                || (d.getConsumoHistoricos() != null && !d.getConsumoHistoricos().isEmpty());
//        if (!tieneConsumo) {
//            throw new Exception("Se necesita al menos un dato de consumo kWh.");
//        }
//    }
//
//    public static TipoTarifa detectarTipoTarifa(String codigoTarifa, int duracionDias) {
//        if (codigoTarifa == null) {
//            return null;
//        }
//        String t = codigoTarifa.trim().toUpperCase();
//        boolean bimestral = duracionDias >= 45;
//        if (t.equals("GDMTH")) {
//            return TipoTarifa.GDMTH;
//        }
//        if (t.equals("GDMTO")) {
//            return TipoTarifa.GDMTO;
//        }
//        if (t.equals("PDBT")) {
//            return bimestral ? TipoTarifa.PDBT_BIMESTRAL : TipoTarifa.PDBT_MENSUAL;
//        }
//        if (t.matches("^(1[A-F]?|DAC)$")) {
//            return bimestral ? TipoTarifa.DOMESTICA_BIMESTRAL : TipoTarifa.DOMESTICA_MENSUAL;
//        }
//        return null;
//    }
//}



/**
 * Motor encargado de realizar los cálculos de la cotización
 * de un sistema fotovoltaico siguiendo la lógica del Excel.
 */


    public ResultadoCalculoCotizacion calcular(
            DatosReciboCFE datos,
            ParametrosSistema params,
            List<ProductoCantidad> productos) throws Exception {

        validar(datos, params, productos);

        ResultadoCalculoCotizacion resultado = new ResultadoCalculoCotizacion();
        resultado.setNombreCliente(datos.getNombreCliente());
        resultado.setCiudad(datos.getCiudad());

        // =========================
        // CONSUMO
        // =========================
        double consumoMensual = calcularConsumoMensual(datos);
        double consumoDiario = consumoMensual / 30.0;

        resultado.setConsumoPromedioMensualKwh(consumoMensual);
        resultado.setConsumoPromedioDiarioKwh(consumoDiario);

        // =========================
        // kW REQUERIDOS
        // Fórmula del Excel:
        // kW = (ConsumoDiario * FactorConversion * FactorSistema) / HSP
        // =========================
        double kwpRequerido = (consumoDiario
                * params.getFactorConversion()
                * params.getFactorSistema())
                / params.getHsp();

        resultado.setKwpRequerido(kwpRequerido);

        // =========================
        // PANEL DEL PAQUETE
        // =========================
        ProductoCantidad panelPc = buscarPanel(productos);
        double potenciaPanelW = panelPc.getProducto().getCapacidad();
        int numeroPaneles = (int) Math.round(panelPc.getCantidad());

        // =========================
        // WATTS INSTALADOS
        // =========================
        double wattsInstalados = numeroPaneles * potenciaPanelW;

        resultado.setNumeroPaneles(numeroPaneles);
        resultado.setWattsInstalados(wattsInstalados);
        resultado.setPotenciaInstaladaKwp(wattsInstalados / 1000.0);

        // =========================
        // PRODUCCIÓN DE ENERGÍA
        // Producción diaria = (WattsInstalados * HSP * Eficiencia) / 1000
        // =========================
        double produccionDiaria = (wattsInstalados
                * params.getHsp()
                * params.getEficiencia()) / 1000.0;

        double produccionMensual = produccionDiaria * 30.0;
        double cobertura = (produccionDiaria / consumoDiario) * 100.0;

        resultado.setProduccionDiariaEstimada(produccionDiaria);
        resultado.setGeneracionMensualEstimadaKwh(produccionMensual);
        resultado.setPorcentajeCobertura(cobertura);

        // =========================
        // COSTOS
        // =========================
        double subtotal = calcularSubtotal(productos);
        double iva = subtotal * params.getIva();
        double total = subtotal + iva;

        resultado.setSubtotal(subtotal);
        resultado.setIva(iva);
        resultado.setTotal(total);

        // =========================
        // AHORRO Y RETORNO
        // =========================
        double costoMensualActual = consumoMensual * params.getPrecioKwhReferencia();
        double costoAnualActual = costoMensualActual * 12.0;

        resultado.setCostoPromedioMensual(costoMensualActual);
        resultado.setCostoPromedioAnual(costoAnualActual);

        double retorno = (costoAnualActual > 0)
                ? total / costoAnualActual
                : 0.0;

        resultado.setRetornoInversion(retorno);

        resultado.setProductosFinales(productos);

        return resultado;
    }

    private double calcularConsumoMensual(DatosReciboCFE datos) {
        List<Double> consumos = datos.getConsumosComoLista();
        double suma = consumos.stream().mapToDouble(Double::doubleValue).sum();
        double promedio = suma / consumos.size();

        if (datos.getTipoTarifa().isEsBimestral()) {
            promedio = promedio / 2.0;
        }

        return promedio;
    }

    private ProductoCantidad buscarPanel(List<ProductoCantidad> productos) {
        for (ProductoCantidad pc : productos) {
            if (pc.getProducto().getCategoria() == CategoriaProducto.PANEL) {
                return pc;
            }
        }
        throw new RuntimeException(
                "El paquete seleccionado no contiene un producto tipo PANEL.");
    }

    private double calcularSubtotal(List<ProductoCantidad> productos) {
        return productos.stream()
                .mapToDouble(pc ->
                        pc.getProducto().getPrecioUnitario() * pc.getCantidad())
                .sum();
    }

    private void validar(DatosReciboCFE datos,
                         ParametrosSistema params,
                         List<ProductoCantidad> productos) throws Exception {

        if (datos == null) {
            throw new Exception("Los datos del recibo son obligatorios.");
        }

        if (datos.getTipoTarifa() == null) {
            throw new Exception("El tipo de tarifa es obligatorio.");
        }

        if (params == null) {
            throw new Exception("No se pudieron cargar los parámetros del sistema.");
        }

        if (productos == null || productos.isEmpty()) {
            throw new Exception("El paquete seleccionado no tiene productos.");
        }
    }
}