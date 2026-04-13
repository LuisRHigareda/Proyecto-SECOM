/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Arell
 */


/**
 * Representa los parámetros del sistema necesarios para el cálculo
 * de una cotización fotovoltaica.
 */
public class ParametrosSistema {

    private final double eficiencia;          // Ej. 0.8
    private final double hsp;                 // Horas Sol Pico
    private final double iva;                 // Ej. 0.16
    private final double precioKwhReferencia; // Para cálculo de ahorro
    private final double factorConversion;    // Ej. 1.2 (Excel)
    private final double factorSistema;       // Ej. 1.1 (Excel)

    public ParametrosSistema(double eficiencia,
                             double hsp,
                             double iva,
                             double precioKwhReferencia,
                             double factorConversion,
                             double factorSistema) {
        this.eficiencia = eficiencia;
        this.hsp = hsp;
        this.iva = iva;
        this.precioKwhReferencia = precioKwhReferencia;
        this.factorConversion = factorConversion;
        this.factorSistema = factorSistema;
    }

    public double getEficiencia() {
        return eficiencia;
    }

    public double getHsp() {
        return hsp;
    }

    public double getIva() {
        return iva;
    }

    public double getPrecioKwhReferencia() {
        return precioKwhReferencia;
    }

    public double getFactorConversion() {
        return factorConversion;
    }

    public double getFactorSistema() {
        return factorSistema;
    }
}