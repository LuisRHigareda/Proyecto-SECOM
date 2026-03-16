/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.EstadoServicio;
import itson.secom_domain.enumeradores.TipoServicio;
import java.time.LocalDate;

/**
 *
 * @author Serva
 */
public class Servicio {

    private int id;
    private TipoServicio tipo;
    private LocalDate fecha;
    private String descripcion;
    private EstadoServicio estado;

    private Proyecto proyecto;

    public Servicio() {
    }

    public Servicio(int id, TipoServicio tipo, LocalDate fecha,
            String descripcion, EstadoServicio estado, Proyecto proyecto) {
        this.id = id;
        this.tipo = tipo;
        this.fecha = fecha;
        this.descripcion = descripcion;
        this.estado = estado;
        this.proyecto = proyecto;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public TipoServicio getTipo() {
        return tipo;
    }

    public void setTipo(TipoServicio tipo) {
        this.tipo = tipo;
    }

    public LocalDate getFecha() {
        return fecha;
    }

    public void setFecha(LocalDate fecha) {
        this.fecha = fecha;
    }

    public String getDescripcion() {
        return descripcion;
    }

    public void setDescripcion(String descripcion) {
        this.descripcion = descripcion;
    }

    public EstadoServicio getEstado() {
        return estado;
    }

    public void setEstado(EstadoServicio estado) {
        this.estado = estado;
    }

    public Proyecto getProyecto() {
        return proyecto;
    }

    public void setProyecto(Proyecto proyecto) {
        this.proyecto = proyecto;
    }

    @Override
    public String toString() {
        return "Servicio{"
                + "id=" + id
                + ", tipo='" + tipo + '\''
                + ", fecha=" + fecha
                + ", estado='" + estado + '\''
                + '}';
    }
}
