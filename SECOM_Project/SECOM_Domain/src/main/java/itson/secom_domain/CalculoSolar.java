/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import java.time.LocalDateTime;

/**
 *
 * @author Serva
 */
public class CalculoSolar {

    private int id;
    private String estadoMX;
    private double insolacionUsada;
    private double potencialPanel;
    private int numeroPaneles;
    private double wattsInstalados;
    private double capacidadInversor;
    private double produccionDiariaEstimada;
    private double produccionAnualEstimada;
    private double porcentajeGeneracion;
    private double factorConversionUsado;
    private double factorReflexionUsado;
    private LocalDateTime fechaCalculo;

    private Cotizacion cotizacion;

    public CalculoSolar() {
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getEstadoMX() {
        return estadoMX;
    }

    public void setEstadoMX(String estadoMX) {
        this.estadoMX = estadoMX;
    }

    public double getInsolacionUsada() {
        return insolacionUsada;
    }

    public void setInsolacionUsada(double insolacionUsada) {
        this.insolacionUsada = insolacionUsada;
    }

    public double getPotencialPanel() {
        return potencialPanel;
    }

    public void setPotencialPanel(double potencialPanel) {
        this.potencialPanel = potencialPanel;
    }

    public int getNumeroPaneles() {
        return numeroPaneles;
    }

    public void setNumeroPaneles(int numeroPaneles) {
        this.numeroPaneles = numeroPaneles;
    }

    public double getWattsInstalados() {
        return wattsInstalados;
    }

    public void setWattsInstalados(double wattsInstalados) {
        this.wattsInstalados = wattsInstalados;
    }

    public double getCapacidadInversor() {
        return capacidadInversor;
    }

    public void setCapacidadInversor(double capacidadInversor) {
        this.capacidadInversor = capacidadInversor;
    }

    public double getProduccionDiariaEstimada() {
        return produccionDiariaEstimada;
    }

    public void setProduccionDiariaEstimada(double produccionDiariaEstimada) {
        this.produccionDiariaEstimada = produccionDiariaEstimada;
    }

    public double getProduccionAnualEstimada() {
        return produccionAnualEstimada;
    }

    public void setProduccionAnualEstimada(double produccionAnualEstimada) {
        this.produccionAnualEstimada = produccionAnualEstimada;
    }

    public double getPorcentajeGeneracion() {
        return porcentajeGeneracion;
    }

    public void setPorcentajeGeneracion(double porcentajeGeneracion) {
        this.porcentajeGeneracion = porcentajeGeneracion;
    }

    public double getFactorConversionUsado() {
        return factorConversionUsado;
    }

    public void setFactorConversionUsado(double factorConversionUsado) {
        this.factorConversionUsado = factorConversionUsado;
    }

    public double getFactorReflexionUsado() {
        return factorReflexionUsado;
    }

    public void setFactorReflexionUsado(double factorReflexionUsado) {
        this.factorReflexionUsado = factorReflexionUsado;
    }

    public LocalDateTime getFechaCalculo() {
        return fechaCalculo;
    }

    public void setFechaCalculo(LocalDateTime fechaCalculo) {
        this.fechaCalculo = fechaCalculo;
    }

    public Cotizacion getCotizacion() {
        return cotizacion;
    }

    public void setCotizacion(Cotizacion cotizacion) {
        this.cotizacion = cotizacion;
    }

    @Override
    public String toString() {
        return "CalculoSolar{paneles=" + numeroPaneles
                + ", kWp=" + String.format("%.2f", wattsInstalados / 1000.0)
                + ", produccionAnual=" + produccionAnualEstimada + " kWh}";
    }

}
