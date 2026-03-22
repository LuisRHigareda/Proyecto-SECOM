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

    private Vendedor vendedor;
    private Cliente cliente;
    private Paquete paquete;

    private double consumoPromedioMensualKwh;
    private double consumoPromedioDiarioKwh;
    private double consumoPromedioMensual;
    private double consumoPromedioAnual;

    private double wattsInstalados;
    private double produccionDiariaEstimada;
    private double porcentajeCobertura;
    private double retornoInversion;

    private double subtotal;
    private double iva;
    private double total;

    private EstadoCotizacion estado;
    private boolean financiamiento;
    private boolean proyectoGenerado;
    private String notas;

    private int createdBy;
    private int updatedBy;
    private LocalDateTime createdAt;
    private LocalDateTime updatedAt;

    public Cotizacion() {
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

    public Vendedor getVendedor() {
        return vendedor;
    }

    public void setVendedor(Vendedor vendedor) {
        this.vendedor = vendedor;
    }

    public Cliente getCliente() {
        return cliente;
    }

    public void setCliente(Cliente cliente) {
        this.cliente = cliente;
    }

    public Paquete getPaquete() {
        return paquete;
    }

    public void setPaquete(Paquete paquete) {
        this.paquete = paquete;
    }

    public double getConsumoPromedioMensualKwh() {
        return consumoPromedioMensualKwh;
    }

    public void setConsumoPromedioMensualKwh(double consumoPromedioMensualKwh) {
        this.consumoPromedioMensualKwh = consumoPromedioMensualKwh;
    }

    public double getConsumoPromedioDiarioKwh() {
        return consumoPromedioDiarioKwh;
    }

    public void setConsumoPromedioDiarioKwh(double consumoPromedioDiarioKwh) {
        this.consumoPromedioDiarioKwh = consumoPromedioDiarioKwh;
    }

    public double getConsumoPromedioMensual() {
        return consumoPromedioMensual;
    }

    public void setConsumoPromedioMensual(double consumoPromedioMensual) {
        this.consumoPromedioMensual = consumoPromedioMensual;
    }

    public double getConsumoPromedioAnual() {
        return consumoPromedioAnual;
    }

    public void setConsumoPromedioAnual(double consumoPromedioAnual) {
        this.consumoPromedioAnual = consumoPromedioAnual;
    }

    public double getWattsInstalados() {
        return wattsInstalados;
    }

    public void setWattsInstalados(double wattsInstalados) {
        this.wattsInstalados = wattsInstalados;
    }

    public double getProduccionDiariaEstimada() {
        return produccionDiariaEstimada;
    }

    public void setProduccionDiariaEstimada(double produccionDiariaEstimada) {
        this.produccionDiariaEstimada = produccionDiariaEstimada;
    }

    public double getPorcentajeCobertura() {
        return porcentajeCobertura;
    }

    public void setPorcentajeCobertura(double porcentajeCobertura) {
        this.porcentajeCobertura = porcentajeCobertura;
    }

    public double getRetornoInversion() {
        return retornoInversion;
    }

    public void setRetornoInversion(double retornoInversion) {
        this.retornoInversion = retornoInversion;
    }

    public double getSubtotal() {
        return subtotal;
    }

    public void setSubtotal(double subtotal) {
        this.subtotal = subtotal;
    }

    public double getIva() {
        return iva;
    }

    public void setIva(double iva) {
        this.iva = iva;
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

    public boolean isFinanciamiento() {
        return financiamiento;
    }

    public void setFinanciamiento(boolean financiamiento) {
        this.financiamiento = financiamiento;
    }

    public boolean isProyectoGenerado() {
        return proyectoGenerado;
    }

    public void setProyectoGenerado(boolean proyectoGenerado) {
        this.proyectoGenerado = proyectoGenerado;
    }

    public String getNotas() {
        return notas;
    }

    public void setNotas(String notas) {
        this.notas = notas;
    }

    public int getCreatedBy() {
        return createdBy;
    }

    public void setCreatedBy(int createdBy) {
        this.createdBy = createdBy;
    }

    public int getUpdatedBy() {
        return updatedBy;
    }

    public void setUpdatedBy(int updatedBy) {
        this.updatedBy = updatedBy;
    }

    public LocalDateTime getCreatedAt() {
        return createdAt;
    }

    public void setCreatedAt(LocalDateTime createdAt) {
        this.createdAt = createdAt;
    }

    public LocalDateTime getUpdatedAt() {
        return updatedAt;
    }

    public void setUpdatedAt(LocalDateTime updatedAt) {
        this.updatedAt = updatedAt;
    }

    @Override
    public String toString() {
        return "Cotizacion{id=" + id
                + ", cliente=" + (cliente != null ? cliente.getNombreComercial() : "—")
                + ", consumo=" + consumoPromedioMensualKwh + " kWh/mes"
                + ", total=$" + total
                + ", estado=" + estado + "}";
    }
}
