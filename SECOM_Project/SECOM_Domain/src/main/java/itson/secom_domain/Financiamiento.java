/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.EstatusFinanciamiento;
import itson.secom_domain.enumeradores.MedioCobroFinanciamiento;
import itson.secom_domain.enumeradores.TipoFinanciamiento;

/**
 *
 * @author Serva
 */
public class Financiamiento {

    private int id;
    private String folio;
    private TipoFinanciamiento tipo;
    private MedioCobroFinanciamiento medioCobro;
    private double montoFinanciado;
    private EstatusFinanciamiento estatus;
    private String financiadoPor;

    private Proyecto proyecto;

    public Financiamiento() {
    }

    public Financiamiento(int id, String folio, TipoFinanciamiento tipo, MedioCobroFinanciamiento medioCobro,
            double montoFinanciado, EstatusFinanciamiento estatus,
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

    public TipoFinanciamiento getTipo() {
        return tipo;
    }

    public void setTipo(TipoFinanciamiento tipo) {
        this.tipo = tipo;
    }

    public MedioCobroFinanciamiento getMedioCobro() {
        return medioCobro;
    }

    public void setMedioCobro(MedioCobroFinanciamiento medioCobro) {
        this.medioCobro = medioCobro;
    }

    public double getMontoFinanciado() {
        return montoFinanciado;
    }

    public void setMontoFinanciado(double montoFinanciado) {
        this.montoFinanciado = montoFinanciado;
    }

    public EstatusFinanciamiento getEstatus() {
        return estatus;
    }

    public void setEstatus(EstatusFinanciamiento estatus) {
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
