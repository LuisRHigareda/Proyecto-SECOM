/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Enum.java to edit this template
 */
package itson.secom_domain.enumeradores;

/**
 *
 * @author Serva
 */
public enum TipoTarifa {
//    DOMESTICA_MENSUAL ("Domestica Mensual", false, true),
//    DOMESTICA_BIMESTRAL("Domestica Bimestral", true, true),
//    PDBT_MENSUAL("PDBT Mensual", false, false),
//    PDBT_BIMESTRAL("PDBT Bimestral", true, false),
//    GDMTH("GDMTH", false, false),
//    GDMTO("GDMTO", false, false);
//    
//    private final String etiqueta;
//    private final boolean esBimestral;
//    private final boolean esDomestica;
//
//    private TipoTarifa(String etiqueta, boolean esBimestral, boolean esDomestica) {
//        this.etiqueta = etiqueta;
//        this.esBimestral = esBimestral;
//        this.esDomestica = esDomestica;
//    }
//    
//    public String getEtiqueta(){
//        return etiqueta;
//    }
//    
//    public boolean isEsBimestral(){
//        return esBimestral;
//    }
//    
//    public boolean isEsDomestica(){
//        return esDomestica;
//    }
//    
//    public boolean esIndustrial(){
//        return !esDomestica;
//    }
//    
//    @Override
//    public String toString(){
//        return etiqueta;
//    }
 
    DOMESTICA_MENSUAL(false),
    DOMESTICA_BIMESTRAL(true),
    PDBT_MENSUAL(false),
    PDBT_BIMESTRAL(true),
    GDMTO(false),
    GDMTH(false);

    private final boolean esBimestral;

    TipoTarifa(boolean esBimestral) {
        this.esBimestral = esBimestral;
    }

    public boolean isEsBimestral() {
        return esBimestral;
    }
}

