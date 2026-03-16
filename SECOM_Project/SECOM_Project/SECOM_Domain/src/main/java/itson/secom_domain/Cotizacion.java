/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import java.time.LocalDate;

/**
 *
 * @author Serva
 */
public class Cotizacion {

    private int id;
    private String folio;
    private LocalDate fechaEmision;
    private double consumoEstimado;
    private double total;
    private String estado;
    private LocalDate vigencia;

    private Cliente cliente;

    public Cotizacion() {
    }

    public Cotizacion(int id, String folio, LocalDate fechaEmision,
            double consumoEstimado, double total,
            String estado, LocalDate vigencia, Cliente cliente) {
        this.id = id;
        this.folio = folio;
        this.fechaEmision = fechaEmision;
        this.consumoEstimado = consumoEstimado;
        this.total = total;
        this.estado = estado;
        this.vigencia = vigencia;
        this.cliente = cliente;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getFolio() {
        return folio;
    }

    public void setFolio(String folio) {
        this.folio = folio;
    }

    public LocalDate getFechaEmision() {
        return fechaEmision;
    }

    public void setFechaEmision(LocalDate fechaEmision) {
        this.fechaEmision = fechaEmision;
    }

    public double getConsumoEstimado() {
        return consumoEstimado;
    }

    public void setConsumoEstimado(double consumoEstimado) {
        this.consumoEstimado = consumoEstimado;
    }

    public double getTotal() {
        return total;
    }

    public void setTotal(double total) {
        this.total = total;
    }

    public String getEstado() {
        return estado;
    }

    public void setEstado(String estado) {
        this.estado = estado;
    }

    public LocalDate getVigencia() {
        return vigencia;
    }

    public void setVigencia(LocalDate vigencia) {
        this.vigencia = vigencia;
    }

    public Cliente getCliente() {
        return cliente;
    }

    public void setCliente(Cliente cliente) {
        this.cliente = cliente;
    }

    @Override
    public String toString() {
        return "Cotizacion{"
                + "id=" + id
                + ", folio='" + folio + '\''
                + ", fechaEmision=" + fechaEmision
                + ", total=" + total
                + ", estado='" + estado + '\''
                + ", vigencia=" + vigencia
                + '}';
    }
}
