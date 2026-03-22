/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.TipoTarifa;

/**
 *
 * @author Serva
 */
public class ResultadoCalculoCotizacion {

    private String nombreCliente;
    private String direccion;
    private String noServicio;
    private String tarifa;
    private TipoTarifa tipoTarifa;
    private int noHilos;
    private boolean esBimestral;

    private double consumoPromedioMensualKwh;

    private double pagoPromedioCFE;
    private double costoBaseConSolar;
    private double ahorroMensualEstimado;
    private double pagoEstimadoConSolar;

    private int numeroPaneles;
    private double potenciaInstaladaKwp;
    private double wattsInstalados;
    private double generacionMensualEstimadaKwh;
    private double generacioAnualEstimadaKwh;
    private double produccionDiariaEstimada;
    private double porcentajCobertura;
    private double retornoInversion;

    //Impacto ambiental (25 años)
    private double co2EvitadoToneladas25años;
    private int arbolesEquivalentes25Años;

    public ResultadoCalculoCotizacion() {
    }

    public String getNombreCliente() {
        return nombreCliente;
    }

    public void setNombreCliente(String nombreCliente) {
        this.nombreCliente = nombreCliente;
    }

    public String getDireccion() {
        return direccion;
    }

    public void setDireccion(String direccion) {
        this.direccion = direccion;
    }

    public String getNoServicio() {
        return noServicio;
    }

    public void setNoServicio(String noServicio) {
        this.noServicio = noServicio;
    }

    public String getTarifa() {
        return tarifa;
    }

    public void setTarifa(String tarifa) {
        this.tarifa = tarifa;
    }

    public TipoTarifa getTipoTarifa() {
        return tipoTarifa;
    }

    public void setTipoTarifa(TipoTarifa tipoTarifa) {
        this.tipoTarifa = tipoTarifa;
    }

    public int getNoHilos() {
        return noHilos;
    }

    public void setNoHilos(int noHilos) {
        this.noHilos = noHilos;
    }

    public boolean isEsBimestral() {
        return esBimestral;
    }

    public void setEsBimestral(boolean esBimestral) {
        this.esBimestral = esBimestral;
    }

    public double getConsumoPromedioMensualKwh() {
        return consumoPromedioMensualKwh;
    }

    public void setConsumoPromedioMensualKwh(double consumoPromedioMensualKwh) {
        this.consumoPromedioMensualKwh = consumoPromedioMensualKwh;
    }

    public double getPagoPromedioCFE() {
        return pagoPromedioCFE;
    }

    public void setPagoPromedioCFE(double pagoPromedioCFE) {
        this.pagoPromedioCFE = pagoPromedioCFE;
    }

    public double getCostoBaseConSolar() {
        return costoBaseConSolar;
    }

    public void setCostoBaseConSolar(double costoBaseConSolar) {
        this.costoBaseConSolar = costoBaseConSolar;
    }

    public double getAhorroMensualEstimado() {
        return ahorroMensualEstimado;
    }

    public void setAhorroMensualEstimado(double ahorroMensualEstimado) {
        this.ahorroMensualEstimado = ahorroMensualEstimado;
    }

    public double getPagoEstimadoConSolar() {
        return pagoEstimadoConSolar;
    }

    public void setPagoEstimadoConSolar(double pagoEstimadoConSolar) {
        this.pagoEstimadoConSolar = pagoEstimadoConSolar;
    }

    public int getNumeroPaneles() {
        return numeroPaneles;
    }

    public void setNumeroPaneles(int numeroPaneles) {
        this.numeroPaneles = numeroPaneles;
    }

    public double getPotenciaInstaladaKwp() {
        return potenciaInstaladaKwp;
    }

    public void setPotenciaInstaladaKwp(double potenciaInstaladaKwp) {
        this.potenciaInstaladaKwp = potenciaInstaladaKwp;
    }

    public double getWattsInstalados() {
        return wattsInstalados;
    }

    public void setWattsInstalados(double wattsInstalados) {
        this.wattsInstalados = wattsInstalados;
    }

    public double getGeneracionMensualEstimadaKwh() {
        return generacionMensualEstimadaKwh;
    }

    public void setGeneracionMensualEstimadaKwh(double generacionMensualEstimadaKwh) {
        this.generacionMensualEstimadaKwh = generacionMensualEstimadaKwh;
    }

    public double getGeneracioAnualEstimadaKwh() {
        return generacioAnualEstimadaKwh;
    }

    public void setGeneracioAnualEstimadaKwh(double generacioAnualEstimadaKwh) {
        this.generacioAnualEstimadaKwh = generacioAnualEstimadaKwh;
    }

    public double getProduccionDiariaEstimada() {
        return produccionDiariaEstimada;
    }

    public void setProduccionDiariaEstimada(double produccionDiariaEstimada) {
        this.produccionDiariaEstimada = produccionDiariaEstimada;
    }

    public double getPorcentajCobertura() {
        return porcentajCobertura;
    }

    public void setPorcentajCobertura(double porcentajCobertura) {
        this.porcentajCobertura = porcentajCobertura;
    }

    public double getRetornoInversion() {
        return retornoInversion;
    }

    public void setRetornoInversion(double retornoInversion) {
        this.retornoInversion = retornoInversion;
    }

    public double getCo2EvitadoToneladas25años() {
        return co2EvitadoToneladas25años;
    }

    public void setCo2EvitadoToneladas25años(double co2EvitadoToneladas25años) {
        this.co2EvitadoToneladas25años = co2EvitadoToneladas25años;
    }

    public int getArbolesEquivalentes25Años() {
        return arbolesEquivalentes25Años;
    }

    public void setArbolesEquivalentes25Años(int arbolesEquivalentes25Años) {
        this.arbolesEquivalentes25Años = arbolesEquivalentes25Años;
    }

    public int getPotenciaPanelW() {
        return 550;
    }

    public long getCostoProyectoConIva() {
        double subtotal = potenciaInstaladaKwp * 22000.0;
        return Math.round(subtotal * 1.16);
    }

    @Override
    public String toString() {
        return String.format(
                "ResultadoCalculoCotizacion{\n"
                + "  cliente='%s', tarifa='%s' (%s)\n"
                + "  consumo=%.1f kWh/mes\n"
                + "  pagoCFE=$%.2f  ahorro=$%.2f/mes\n"
                + "  paneles=%d (%.2f kWp)\n"
                + "  generacion=%.1f kWh/año\n"
                + "  retorno=%.1f años\n"
                + "  CO2 25años=%.1f ton | árboles=%d\n"
                + "}",
                nombreCliente, tarifa, tipoTarifa,
                consumoPromedioMensualKwh,
                pagoPromedioCFE, ahorroMensualEstimado,
                numeroPaneles, potenciaInstaladaKwp,
                generacioAnualEstimadaKwh,
                retornoInversion,
                co2EvitadoToneladas25años, arbolesEquivalentes25Años
        );
    }

}
