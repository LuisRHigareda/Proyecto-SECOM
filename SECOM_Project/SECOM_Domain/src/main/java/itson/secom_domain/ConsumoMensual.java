/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class ConsumoMensual {

    private int id;
    private int mes;
    private int año;
    private double consumoKwh;

    private Cotizacion cotizacion;

    public ConsumoMensual() {
    }
    
    public ConsumoMensual(int mes, int año, double consumoKwh, Cotizacion cotizacion) {
    this.mes = mes;
    this.año = año;
    this.consumoKwh = consumoKwh;
    this.cotizacion = cotizacion;
}

    public ConsumoMensual(int id, int mes, int año, double consumoKwh, Cotizacion cotizacion) {
        this.id = id;
        this.mes = mes;
        this.año = año;
        this.consumoKwh = consumoKwh;
        this.cotizacion = cotizacion;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public int getMes() {
        return mes;
    }

    public void setMes(int mes) {
        this.mes = mes;
    }

    public int getAño() {
        return año;
    }

    public void setAño(int año) {
        this.año = año;
    }

    public double getConsumoKwh() {
        return consumoKwh;
    }

    public void setConsumoKwh(double consumoKwh) {
        this.consumoKwh = consumoKwh;
    }

    public Cotizacion getCotizacion() {
        return cotizacion;
    }

    public void setCotizacion(Cotizacion cotizacion) {
        this.cotizacion = cotizacion;
    }

    @Override
    public String toString() {
        return "ConsumoMensual{" + mes + "/" + año + "=" + consumoKwh + " kWh}";
    }

}
