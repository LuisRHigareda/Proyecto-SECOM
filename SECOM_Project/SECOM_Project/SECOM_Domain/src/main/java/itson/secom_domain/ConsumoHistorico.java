/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class ConsumoHistorico {

    private int id;
    private String perido;
    private double consumoKWH;
    private double pagoMXM;

    private Cotizacion cotizacion;

    public ConsumoHistorico() {
    }

    public ConsumoHistorico(int id, String perido, double consumoKWH,
            double pagoMXM, Cotizacion cotizacion) {
        this.id = id;
        this.perido = perido;
        this.consumoKWH = consumoKWH;
        this.pagoMXM = pagoMXM;
        this.cotizacion = cotizacion;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getPerido() {
        return perido;
    }

    public void setPerido(String perido) {
        this.perido = perido;
    }

    public double getConsumoKWH() {
        return consumoKWH;
    }

    public void setConsumoKWH(double consumoKWH) {
        this.consumoKWH = consumoKWH;
    }

    public double getPagoMXM() {
        return pagoMXM;
    }

    public void setPagoMXM(double pagoMXM) {
        this.pagoMXM = pagoMXM;
    }

    public Cotizacion getCotizacion() {
        return cotizacion;
    }

    public void setCotizacion(Cotizacion cotizacion) {
        this.cotizacion = cotizacion;
    }

    @Override
    public String toString() {
        return "ConsumoHistorico{"
                + "id=" + id
                + ", perido='" + perido + '\''
                + ", consumoKWH=" + consumoKWH
                + ", pagoMXM=" + pagoMXM
                + '}';
    }
}
