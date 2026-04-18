/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class ResultadoCotizacion {
    // -------------------------------------------------------
    // Atributos
    // -------------------------------------------------------

    private int id;
    private double potenciaInstalada;
    private int numeroPaneles;
    private double generacionAnualEstimada;
    private double ahorroMensualEstimado;
    private double pagoPrimedioCFE;
    private double pagoEstimadoConSolar;
    private String modeloInversor;
    private String tipoTecho;
    private String sombras;

    private Cotizacion cotizacion;

    public ResultadoCotizacion() {
    }

    public ResultadoCotizacion(int id, double potenciaInstalada, int numeroPaneles,
            double generacionAnualEstimada, double ahorroMensualEstimado,
            double pagoPrimedioCFE, double pagoEstimadoConSolar,
            String modeloInversor, String tipoTecho, String sombras,
            Cotizacion cotizacion) {
        this.id = id;
        this.potenciaInstalada = potenciaInstalada;
        this.numeroPaneles = numeroPaneles;
        this.generacionAnualEstimada = generacionAnualEstimada;
        this.ahorroMensualEstimado = ahorroMensualEstimado;
        this.pagoPrimedioCFE = pagoPrimedioCFE;
        this.pagoEstimadoConSolar = pagoEstimadoConSolar;
        this.modeloInversor = modeloInversor;
        this.tipoTecho = tipoTecho;
        this.sombras = sombras;
        this.cotizacion = cotizacion;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public double getPotenciaInstalada() {
        return potenciaInstalada;
    }

    public void setPotenciaInstalada(double potenciaInstalada) {
        this.potenciaInstalada = potenciaInstalada;
    }

    public int getNumeroPaneles() {
        return numeroPaneles;
    }

    public void setNumeroPaneles(int numeroPaneles) {
        this.numeroPaneles = numeroPaneles;
    }

    public double getGeneracionAnualEstimada() {
        return generacionAnualEstimada;
    }

    public void setGeneracionAnualEstimada(double generacionAnualEstimada) {
        this.generacionAnualEstimada = generacionAnualEstimada;
    }

    public double getAhorroMensualEstimado() {
        return ahorroMensualEstimado;
    }

    public void setAhorroMensualEstimado(double ahorroMensualEstimado) {
        this.ahorroMensualEstimado = ahorroMensualEstimado;
    }

    public double getPagoPrimedioCFE() {
        return pagoPrimedioCFE;
    }

    public void setPagoPrimedioCFE(double pagoPrimedioCFE) {
        this.pagoPrimedioCFE = pagoPrimedioCFE;
    }

    public double getPagoEstimadoConSolar() {
        return pagoEstimadoConSolar;
    }

    public void setPagoEstimadoConSolar(double pagoEstimadoConSolar) {
        this.pagoEstimadoConSolar = pagoEstimadoConSolar;
    }

    public String getModeloInversor() {
        return modeloInversor;
    }

    public void setModeloInversor(String modeloInversor) {
        this.modeloInversor = modeloInversor;
    }

    public String getTipoTecho() {
        return tipoTecho;
    }

    public void setTipoTecho(String tipoTecho) {
        this.tipoTecho = tipoTecho;
    }

    public String getSombras() {
        return sombras;
    }

    public void setSombras(String sombras) {
        this.sombras = sombras;
    }

    public Cotizacion getCotizacion() {
        return cotizacion;
    }

    public void setCotizacion(Cotizacion cotizacion) {
        this.cotizacion = cotizacion;
    }

    @Override
    public String toString() {
        return "ResultadoCotizacion{"
                + "id=" + id
                + ", potenciaInstalada=" + potenciaInstalada
                + ", numeroPaneles=" + numeroPaneles
                + ", ahorroMensualEstimado=" + ahorroMensualEstimado
                + ", tipoTecho='" + tipoTecho + '\''
                + '}';
    }
}
