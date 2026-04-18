/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.EstatusPedido;
import java.time.LocalDate;

/**
 *
 * @author Serva
 */
public class Compra {

    private int id;
    private LocalDate fecha;
    private LocalDate fechaEntregaEstimada;
    private EstatusPedido estatusPedido;

    private Proyecto proyecto;

    public Compra() {
    }

    public Compra(int id, LocalDate fecha, LocalDate fechaEntregaEstimada,
            EstatusPedido estatusPedido, Proyecto proyecto) {
        this.id = id;
        this.fecha = fecha;
        this.fechaEntregaEstimada = fechaEntregaEstimada;
        this.estatusPedido = estatusPedido;
        this.proyecto = proyecto;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public LocalDate getFecha() {
        return fecha;
    }

    public void setFecha(LocalDate fecha) {
        this.fecha = fecha;
    }

    public LocalDate getFechaEntregaEstimada() {
        return fechaEntregaEstimada;
    }

    public void setFechaEntregaEstimada(LocalDate fechaEntregaEstimada) {
        this.fechaEntregaEstimada = fechaEntregaEstimada;
    }

    public EstatusPedido getEstatusPedido() {
        return estatusPedido;
    }

    public void setEstatusPedido(EstatusPedido estatusPedido) {
        this.estatusPedido = estatusPedido;
    }

    public Proyecto getProyecto() {
        return proyecto;
    }

    public void setProyecto(Proyecto proyecto) {
        this.proyecto = proyecto;
    }

    @Override
    public String toString() {
        return "Compra{"
                + "id=" + id
                + ", fecha=" + fecha
                + ", fechaEntregaEstimada=" + fechaEntregaEstimada
                + ", estatusPedido='" + estatusPedido + '\''
                + '}';
    }
}
