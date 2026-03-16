/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class Financiamiento {

    private int id;
    private String folio;
    private String tipo;
    private String medioCobro;
    private double montoFinanciado;
    private String estatus;
    private String financiadoPor;

    private Proyecto proyecto;

    public Financiamiento() {
    }

    public Financiamiento(int id, String folio, String tipo, String medioCobro,
            double montoFinanciado, String estatus,
            String financiadoPor, Proyecto proyecto) {
        this.id = id;
        this.folio = folio;
        this.tipo = tipo;
        this.medioCobro = medioCobro;
        this.montoFinanciado = montoFinanciado;
        this.estatus = estatus;
        this.financiadoPor = financiadoPor;
        this.proyecto = proyecto;
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

    public String getTipo() {
        return tipo;
    }

    public void setTipo(String tipo) {
        this.tipo = tipo;
    }

    public String getMedioCobro() {
        return medioCobro;
    }

    public void setMedioCobro(String medioCobro) {
        this.medioCobro = medioCobro;
    }

    public double getMontoFinanciado() {
        return montoFinanciado;
    }

    public void setMontoFinanciado(double montoFinanciado) {
        this.montoFinanciado = montoFinanciado;
    }

    public String getEstatus() {
        return estatus;
    }

    public void setEstatus(String estatus) {
        this.estatus = estatus;
    }

    public String getFinanciadoPor() {
        return financiadoPor;
    }

    public void setFinanciadoPor(String financiadoPor) {
        this.financiadoPor = financiadoPor;
    }

    public Proyecto getProyecto() {
        return proyecto;
    }

    public void setProyecto(Proyecto proyecto) {
        this.proyecto = proyecto;
    }

    @Override
    public String toString() {
        return "Financiamiento{"
                + "id=" + id
                + ", folio='" + folio + '\''
                + ", tipo='" + tipo + '\''
                + ", montoFinanciado=" + montoFinanciado
                + ", estatus='" + estatus + '\''
                + '}';
    }
}
