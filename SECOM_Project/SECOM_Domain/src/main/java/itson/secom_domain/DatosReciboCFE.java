/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.TipoTarifa;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 *
 * @author Serva
 */
public class DatosReciboCFE {
//
//    private TipoTarifa tipoTarifa;
//    private String nombre;
//    private String direccion;
//    private String noServicio;
//    private String tarifa;
//    private int noHilos;
//    private int numeroEstado;
//
//    private String periodoFacturado;
//    private int duracionDias;
//
//    private double consumoActualKwh;
//    private double pagoActual;
//
//    private List<Double> consumoHistoricos;
//    private List<Double> pagosHistoricos;
//
//    private List<Double> preciosMedios;
//
//    private double costoSuministro;
//    private double ivaPorcentaje;
//    private double costoDAP;
//
//    public DatosReciboCFE() {
//    }
//
//    public TipoTarifa getTipoTarifa() {
//        return tipoTarifa;
//    }
//
//    public void setTipoTarifa(TipoTarifa tipoTarifa) {
//        this.tipoTarifa = tipoTarifa;
//    }
//
//    public String getNombre() {
//        return nombre;
//    }
//
//    public void setNombre(String nombre) {
//        this.nombre = nombre;
//    }
//
//    public String getDireccion() {
//        return direccion;
//    }
//
//    public void setDireccion(String direccion) {
//        this.direccion = direccion;
//    }
//
//    public String getNoServicio() {
//        return noServicio;
//    }
//
//    public void setNoServicio(String noServicio) {
//        this.noServicio = noServicio;
//    }
//
//    public String getTarifa() {
//        return tarifa;
//    }
//
//    public void setTarifa(String tarifa) {
//        this.tarifa = tarifa;
//    }
//
//    public int getNoHilos() {
//        return noHilos;
//    }
//
//    public void setNoHilos(int noHilos) {
//        this.noHilos = noHilos;
//    }
//
//    public int getNumeroEstado() {
//        return numeroEstado;
//    }
//
//    public void setNumeroEstado(int numeroEstado) {
//        this.numeroEstado = numeroEstado;
//    }
//
//    public String getPeriodoFacturado() {
//        return periodoFacturado;
//    }
//
//    public void setPeriodoFacturado(String periodoFacturado) {
//        this.periodoFacturado = periodoFacturado;
//    }
//
//    public int getDuracionDias() {
//        return duracionDias;
//    }
//
//    public void setDuracionDias(int duracionDias) {
//        this.duracionDias = duracionDias;
//    }
//
//    public double getConsumoActualKwh() {
//        return consumoActualKwh;
//    }
//
//    public void setConsumoActualKwh(double consumoActualKwh) {
//        this.consumoActualKwh = consumoActualKwh;
//    }
//
//    public double getPagoActual() {
//        return pagoActual;
//    }
//
//    public void setPagoActual(double pagoActual) {
//        this.pagoActual = pagoActual;
//    }
//
//    public List<Double> getConsumoHistoricos() {
//        return consumoHistoricos;
//    }
//
//    public void setConsumoHistoricos(List<Double> consumoHistoricos) {
//        this.consumoHistoricos = consumoHistoricos;
//    }
//
//    public List<Double> getPagosHistoricos() {
//        return pagosHistoricos;
//    }
//
//    public void setPagosHistoricos(List<Double> pagosHistoricos) {
//        this.pagosHistoricos = pagosHistoricos;
//    }
//
//    public List<Double> getPreciosMedios() {
//        return preciosMedios;
//    }
//
//    public void setPreciosMedios(List<Double> preciosMedios) {
//        this.preciosMedios = preciosMedios;
//    }
//
//    public double getCostoSuministro() {
//        return costoSuministro;
//    }
//
//    public void setCostoSuministro(double costoSuministro) {
//        this.costoSuministro = costoSuministro;
//    }
//
//    public double getIvaPorcentaje() {
//        return ivaPorcentaje;
//    }
//
//    public void setIvaPorcentaje(double ivaPorcentaje) {
//        this.ivaPorcentaje = ivaPorcentaje;
//    }
//
//    public double getCostoDAP() {
//        return costoDAP;
//    }
//
//    public void setCostoDAP(double costoDAP) {
//        this.costoDAP = costoDAP;
//    }
//
//    public boolean esIndustrial() {
//        return tipoTarifa != null && tipoTarifa.esIndustrial();
//    }
//
//    @Override
//    public String toString() {
//        return "DatosReciboCFE{tarifa=" + tarifa
//                + ", tipo=" + tipoTarifa
//                + ", nombre=" + nombre
//                + ", consumo=" + consumoActualKwh + " kWh}";
//    }


    private String nombreCliente;
    private String ciudad;
    private List<Double> consumoHistoricos;
    private TipoTarifa tipoTarifa;
    private Double consumoDiarioDisenio; // 

public Double getConsumoDiarioDisenio() {
    return consumoDiarioDisenio;
}

public void setConsumoDiarioDisenio(Double consumoDiarioDisenio) {
    this.consumoDiarioDisenio = consumoDiarioDisenio;
}

    public DatosReciboCFE(String nombre, String consumos1, String tipo, String ciudad1) {
        this.consumoHistoricos = new ArrayList<>();
    }

    public TipoTarifa getTipoTarifa() {
        return tipoTarifa;
    }

    public void setTipoTarifa(TipoTarifa tipoTarifa) {
        this.tipoTarifa = tipoTarifa;
    }


    private String consumos;
    private String tipoPeriodo;


    public DatosReciboCFE(String nombreCliente, String consumos, String tipoPeriodo, String ciudad,TipoTarifa tipoTarifa) {
        this.nombreCliente = nombreCliente;
        this.consumos = consumos;
        this.tipoPeriodo = tipoPeriodo;
        this.ciudad = ciudad;
        this.tipoTarifa = tipoTarifa;
    }

    public String getNombreCliente() {
        return nombreCliente;
    }

    public String getConsumos() {
        return consumos;
    }

    public String getTipoPeriodo() {
        return tipoPeriodo;
    }

    public String getCiudad() {
        return ciudad;
    }
    
    public List<Double> getConsumosComoLista() {
    if (consumos == null || consumos.isEmpty()) {
        return List.of();
    }

    return Arrays.stream(consumos.split(","))
            .map(String::trim)
            .map(Double::parseDouble)
            .collect(Collectors.toList());
}
 public void setTipoTarifaDesdeString(String tarifa) {

    if (tarifa == null || tarifa.trim().isEmpty()) {
        throw new IllegalArgumentException("El tipo de tarifa es obligatorio");
    }

    try {
        this.tipoTarifa = TipoTarifa.valueOf(tarifa.trim().toUpperCase());
    } catch (IllegalArgumentException e) {
        throw new IllegalArgumentException(
            "Tarifa inválida. Usa: DOMESTICA_MENSUAL, DOMESTICA_BIMESTRAL, PDBT_MENSUAL..."
        );
    }
}
}


