/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.EstadoCotizacion;
import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 *
 * @author Serva
 */
public class Cotizacion {

    private int id;
    private LocalDateTime fecha;
    private double consumoPromedioMensualKwh; 
    private double total;
    private EstadoCotizacion estado;
    
    private Cliente cliente;

    public Cotizacion() {
    }

    public Cotizacion(int id, LocalDateTime fecha, double consumoPromedioMensualKwh, double total, EstadoCotizacion estado, Cliente cliente) {
        this.id = id;
        this.fecha = fecha;
        this.consumoPromedioMensualKwh = consumoPromedioMensualKwh;
        this.total = total;
        this.estado = estado;
        this.cliente = cliente;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public LocalDateTime getFecha() {
        return fecha;
    }

    public void setFecha(LocalDateTime fecha) {
        this.fecha = fecha;
    }

    public double getConsumoPromedioMensualKwh() {
        return consumoPromedioMensualKwh;
    }

    public void setConsumoPromedioMensualKwh(double consumoPromedioMensualKwh) {
        this.consumoPromedioMensualKwh = consumoPromedioMensualKwh;
    }

    public double getTotal() {
        return total;
    }

    public void setTotal(double total) {
        this.total = total;
    }

    public EstadoCotizacion getEstado() {
        return estado;
    }

    public void setEstado(EstadoCotizacion estado) {
        this.estado = estado;
    }

    public Cliente getCliente() {
        return cliente;
    }

    public void setCliente(Cliente cliente) {
        this.cliente = cliente;
    }

    @Override
    public String toString() {
        return "Cotizacion{" + "id=" + id + ", fecha=" + fecha + ", consumoPromedioMensualKwh=" + consumoPromedioMensualKwh + ", total=" + total + ", estado=" + estado + ", cliente=" + cliente + '}';
    }

    
}
