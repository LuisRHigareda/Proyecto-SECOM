/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class Recibo {

    private int id;
    private String noServicio;
    private String titularRecibido;
    private String tarifa;
    private String periodoFacturado;
    private double consumoPeriodo;
    private double totalAPagar;
    private double ajusteConsumo;

    private Cotizacion cotizacion;

    public Recibo() {
    }

    public Recibo(int id, String noServicio, String titularRecibido,
            String tarifa, String periodoFacturado, double consumoPeriodo,
            double totalAPagar, double ajusteConsumo, Cotizacion cotizacion) {
        this.id = id;
        this.noServicio = noServicio;
        this.titularRecibido = titularRecibido;
        this.tarifa = tarifa;
        this.periodoFacturado = periodoFacturado;
        this.consumoPeriodo = consumoPeriodo;
        this.totalAPagar = totalAPagar;
        this.ajusteConsumo = ajusteConsumo;
        this.cotizacion = cotizacion;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getNoServicio() {
        return noServicio;
    }

    public void setNoServicio(String noServicio) {
        this.noServicio = noServicio;
    }

    public String getTitularRecibido() {
        return titularRecibido;
    }

    public void setTitularRecibido(String titularRecibido) {
        this.titularRecibido = titularRecibido;
    }

    public String getTarifa() {
        return tarifa;
    }

    public void setTarifa(String tarifa) {
        this.tarifa = tarifa;
    }

    public String getPeriodoFacturado() {
        return periodoFacturado;
    }

    public void setPeriodoFacturado(String periodoFacturado) {
        this.periodoFacturado = periodoFacturado;
    }

    public double getConsumoPeriodo() {
        return consumoPeriodo;
    }

    public void setConsumoPeriodo(double consumoPeriodo) {
        this.consumoPeriodo = consumoPeriodo;
    }

    public double getTotalAPagar() {
        return totalAPagar;
    }

    public void setTotalAPagar(double totalAPagar) {
        this.totalAPagar = totalAPagar;
    }

    public double getAjusteConsumo() {
        return ajusteConsumo;
    }

    public void setAjusteConsumo(double ajusteConsumo) {
        this.ajusteConsumo = ajusteConsumo;
    }

    public Cotizacion getCotizacion() {
        return cotizacion;
    }

    public void setCotizacion(Cotizacion cotizacion) {
        this.cotizacion = cotizacion;
    }

    @Override
    public String toString() {
        return "Recibo{"
                + "id=" + id
                + ", noServicio='" + noServicio + '\''
                + ", titularRecibido='" + titularRecibido + '\''
                + ", tarifa='" + tarifa + '\''
                + ", consumoPeriodo=" + consumoPeriodo
                + ", totalAPagar=" + totalAPagar
                + '}';
    }
}
