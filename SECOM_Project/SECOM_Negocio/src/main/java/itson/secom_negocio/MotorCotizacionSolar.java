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



    public ResultadoCalculoCotizacion calcular(
            DatosReciboCFE datos,
            ParametrosSistema params,
            List<ProductoCantidad> productosBase) throws Exception {

        validar(datos, params, productosBase);

        ResultadoCalculoCotizacion r = new ResultadoCalculoCotizacion();
        r.setNombreCliente(datos.getNombreCliente());
        r.setCiudad(datos.getCiudad());

        double consumoMensual = calcularConsumoMensual(datos);
        double consumoDiario = consumoMensual / 30.0;

        r.setConsumoPromedioMensualKwh(consumoMensual);
        r.setConsumoPromedioDiarioKwh(consumoDiario);

        double kwpRequerido = consumoDiario / (params.getHsp() * params.getEficiencia());
        r.setKwpRequerido(kwpRequerido);

        List<ProductoCantidad> productosAjustados = clonarLista(productosBase);

        dimensionarSegunPaquete(productosAjustados, consumoMensual, params, r);

        double subtotal = calcularSubtotal(productosAjustados);
        double iva = subtotal * params.getIva();
        double total = subtotal + iva;

        r.setSubtotal(subtotal);
        r.setIva(iva);
        r.setTotal(total);

        double costoMensualActual = consumoMensual * params.getPrecioKwhReferencia();
        double costoAnualActual = costoMensualActual * 12.0;

        r.setCostoPromedioMensual(costoMensualActual);
        r.setCostoPromedioAnual(costoAnualActual);

        if (costoAnualActual > 0) {
            r.setRetornoInversion(total / costoAnualActual);
        } else {
            r.setRetornoInversion(0);
        }

        r.setProductosFinales(productosAjustados);

        return r;
    }

    private void dimensionarSegunPaquete(
            List<ProductoCantidad> productos,
            double consumoMensual,
            ParametrosSistema params,
            ResultadoCalculoCotizacion r) {

        ProductoCantidad panelBase = buscarPanel(productos);

        double panelesBase = panelBase.getCantidad();
        double potenciaPanelW = panelBase.getProducto().getCapacidad();

        int panelesFinales = (int) Math.ceil(panelesBase);
        double cobertura = 0;

        while (cobertura < 100.0) {
            double wattsTemporales = panelesFinales * potenciaPanelW;
            double produccionDiariaTemporal = (wattsTemporales / 1000.0) * params.getHsp() * params.getEficiencia();
            double produccionMensualTemporal = produccionDiariaTemporal * 30.0;

            cobertura = (produccionMensualTemporal / consumoMensual) * 100.0;

            if (cobertura < 100.0) {
                panelesFinales++;
            }
        }

        ajustarProductos(productos, panelesFinales, panelesBase);

        double wattsInstalados = calcularWattsDesdeProductos(productos);
        double kwpInstalados = wattsInstalados / 1000.0;
        double produccionDiaria = kwpInstalados * params.getHsp() * params.getEficiencia();
        double produccionMensual = produccionDiaria * 30.0;

        r.setNumeroPaneles(panelesFinales);
        r.setWattsInstalados(wattsInstalados);
        r.setPotenciaInstaladaKwp(kwpInstalados);
        r.setProduccionDiariaEstimada(produccionDiaria);
        r.setGeneracionMensualEstimadaKwh(produccionMensual);
        r.setPorcentajeCobertura((produccionMensual / consumoMensual) * 100.0);
    }

    private ProductoCantidad buscarPanel(List<ProductoCantidad> productos) {
        for (ProductoCantidad pc : productos) {
            if (pc.getProducto() != null
                    && pc.getProducto().getCategoria() != null
                    && pc.getProducto().getCategoria().equalsIgnoreCase("PANEL")) {
                return pc;
            }
        }
        throw new RuntimeException("El paquete seleccionado no contiene un producto tipo PANEL.");
    }

    private void ajustarProductos(List<ProductoCantidad> productos, int panelesFinales, double panelesBase) {
        double factor = panelesFinales / panelesBase;

        for (ProductoCantidad pc : productos) {
            CategoriaProducto categoria = pc.getProducto().getCategoria();

            if (categoria == CategoriaProducto.PANEL) {
                pc.setCantidad(panelesFinales);
            } else {
                pc.setCantidad(Math.ceil(pc.getCantidad() * factor));
            }
        }
    }

    private double calcularWattsDesdeProductos(List<ProductoCantidad> productos) {
        double totalWatts = 0;

        for (ProductoCantidad pc : productos) {
            if (pc.getProducto().getCategoria().equalsIgnoreCase("PANEL")) {
                totalWatts += pc.getProducto().getCapacidad() * pc.getCantidad();
            }
        }

        return totalWatts;
    }

    private double calcularSubtotal(List<ProductoCantidad> productos) {
        double subtotal = 0;

        for (ProductoCantidad pc : productos) {
            subtotal += pc.getProducto().getPrecioUnitario()* pc.getCantidad();
        }

        return subtotal;
    }

    private double calcularConsumoMensual(DatosReciboCFE datos) {
        List<Double> consumos = datos.getConsumosComoLista();

        if (consumos == null || consumos.isEmpty()) {
            return 0;
        }

        double suma = 0;
        int n = 0;

        for (Double c : consumos) {
            if (c != null && c > 0) {
                suma += c;
                n++;
            }
        }

        if (n == 0) {
            return 0;
        }

        double promedio = suma / n;

        return datos.getTipoTarifa().isEsBimestral() ? promedio / 2.0 : promedio;
    }

    private List<ProductoCantidad> clonarLista(List<ProductoCantidad> original) {
        List<ProductoCantidad> copia = new ArrayList<>();

        for (ProductoCantidad pc : original) {
            ProductoCantidad nuevo = new ProductoCantidad();
            nuevo.setProducto(pc.getProducto());
            nuevo.setCantidad(pc.getCantidad());
            copia.add(nuevo);
        }

        return copia;
    }

    private void validar(
            DatosReciboCFE datos,
            ParametrosSistema params,
            List<ProductoCantidad> productos) throws Exception {

        if (datos == null) {
            throw new Exception("Los datos del recibo son obligatorios.");
        }

        if (datos.getNombreCliente() == null || datos.getNombreCliente().isBlank()) {
            throw new Exception("El nombre del cliente es obligatorio.");
        }

        if (datos.getCiudad() == null || datos.getCiudad().isBlank()) {
            throw new Exception("La ciudad es obligatoria.");
        }

        if (datos.getTipoTarifa() == null) {
            throw new Exception("El tipo de tarifa es obligatorio.");
        }

        if (datos.getConsumos()== null || datos.getConsumos().isEmpty()) {
            throw new Exception("Debes capturar al menos un consumo.");
        }

        if (params == null) {
            throw new Exception("No se pudieron cargar los parámetros del sistema.");
        }

        if (params.getEficiencia() <= 0 || params.getHsp() <= 0) {
            throw new Exception("Los parámetros del sistema son inválidos.");
        }

        if (productos == null || productos.isEmpty()) {
            throw new Exception("El paquete seleccionado no tiene productos.");
        }
    }
}
