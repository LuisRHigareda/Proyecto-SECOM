/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Arell
 */
public class ParametrosSistema {
    
    private double eficiencia;
    private double hsp;
    private double iva;
    private double precioKwhReferencia;

    public ParametrosSistema() {
    }

    public ParametrosSistema(double eficiencia, double hsp, double iva, double precioKwhReferencia) {
        this.eficiencia = eficiencia;
        this.hsp = hsp;
        this.iva = iva;
        this.precioKwhReferencia = precioKwhReferencia;
    }

    public double getEficiencia() {
        return eficiencia;
    }

    public void setEficiencia(double eficiencia) {
        this.eficiencia = eficiencia;
    }

    public double getHsp() {
        return hsp;
    }

    public void setHsp(double hsp) {
        this.hsp = hsp;
    }

    public double getIva() {
        return iva;
    }

    public void setIva(double iva) {
        this.iva = iva;
    }

    public double getPrecioKwhReferencia() {
        return precioKwhReferencia;
    }

    public void setPrecioKwhReferencia(double precioKwhReferencia) {
        this.precioKwhReferencia = precioKwhReferencia;
    }
}