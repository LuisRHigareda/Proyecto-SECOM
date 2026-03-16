/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.EstadoProyecto;
import java.time.LocalDate;

/**
 *
 * @author Serva
 */
public class Proyecto {

    private int id;
    private LocalDate fechaInicio;
    private LocalDate fechaFin;
    private EstadoProyecto estado;

    private Cotizacion cotizacion;

    public Proyecto() {
    }

    public Proyecto(int id, LocalDate fechaInicio, LocalDate fechaFin,
            EstadoProyecto estado, Cotizacion cotizacion) {
        this.id = id;
        this.fechaInicio = fechaInicio;
        this.fechaFin = fechaFin;
        this.estado = estado;
        this.cotizacion = cotizacion;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public LocalDate getFechaInicio() {
        return fechaInicio;
    }

    public void setFechaInicio(LocalDate fechaInicio) {
        this.fechaInicio = fechaInicio;
    }

    public LocalDate getFechaFin() {
        return fechaFin;
    }

    public void setFechaFin(LocalDate fechaFin) {
        this.fechaFin = fechaFin;
    }

    public EstadoProyecto getEstado() {
        return estado;
    }

    public void setEstado(EstadoProyecto estado) {
        this.estado = estado;
    }

    public Cotizacion getCotizacion() {
        return cotizacion;
    }

    public void setCotizacion(Cotizacion cotizacion) {
        this.cotizacion = cotizacion;
    }

    @Override
    public String toString() {
        return "Proyecto{"
                + "id=" + id
                + ", fechaInicio=" + fechaInicio
                + ", fechaFin=" + fechaFin
                + ", estado='" + estado + '\''
                + '}';
    }
}
