/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class MediaCompras {

    private int id;
    private double promedioMonto;
    private int totalCompras;
    private String periodoCalculo;

    private Cliente cliente;

    public MediaCompras() {
    }

    public MediaCompras(int id, double promedioMonto, int totalCompras,
            String periodoCalculo, Cliente cliente) {
        this.id = id;
        this.promedioMonto = promedioMonto;
        this.totalCompras = totalCompras;
        this.periodoCalculo = periodoCalculo;
        this.cliente = cliente;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public double getPromedioMonto() {
        return promedioMonto;
    }

    public void setPromedioMonto(double promedioMonto) {
        this.promedioMonto = promedioMonto;
    }

    public int getTotalCompras() {
        return totalCompras;
    }

    public void setTotalCompras(int totalCompras) {
        this.totalCompras = totalCompras;
    }

    public String getPeriodoCalculo() {
        return periodoCalculo;
    }

    public void setPeriodoCalculo(String periodoCalculo) {
        this.periodoCalculo = periodoCalculo;
    }

    public Cliente getCliente() {
        return cliente;
    }

    public void setCliente(Cliente cliente) {
        this.cliente = cliente;
    }

    @Override
    public String toString() {
        return "MediaCompras{"
                + "id=" + id
                + ", promedioMonto=" + promedioMonto
                + ", totalCompras=" + totalCompras
                + ", periodoCalculo='" + periodoCalculo + '\''
                + '}';
    }
}
