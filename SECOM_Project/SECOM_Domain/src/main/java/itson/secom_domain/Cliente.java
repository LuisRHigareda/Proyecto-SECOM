/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.RolUsuario;
import java.time.LocalDateTime;
import java.util.List;

/**
 *
 * @author Serva
 */
public class Cliente extends Usuario {

    private int idCliente;
    private String rfc;
    private String razonSocial;
    private String nombreComercial;
    private String regimenFiscal;
    private String direccionFiscal;

    public Cliente() {
    }

    public Cliente(int idCliente, String rfc, String razonSocial, String nombreComercial, String regimenFiscal, String direccionFiscal) {
        this.idCliente = idCliente;
        this.rfc = rfc;
        this.razonSocial = razonSocial;
        this.nombreComercial = nombreComercial;
        this.regimenFiscal = regimenFiscal;
        this.direccionFiscal = direccionFiscal;
    }

    public Cliente(int idCliente, String rfc, String razonSocial, String nombreComercial, String regimenFiscal, String direccionFiscal, int idUsuario, String userName, String nombre, String email, String password, RolUsuario rolUsuario, List<String> telefono, String ciudad, boolean esActivo, LocalDateTime fechaRegistro) {
        super(idUsuario, userName, nombre, email, password, rolUsuario, telefono, ciudad, esActivo, fechaRegistro);
        this.idCliente = idCliente;
        this.rfc = rfc;
        this.razonSocial = razonSocial;
        this.nombreComercial = nombreComercial;
        this.regimenFiscal = regimenFiscal;
        this.direccionFiscal = direccionFiscal;
    }

    public Cliente(String rfc, String razonSocial, String nombreComercial, String regimenFiscal, String direccionFiscal, String userName, String nombre, String email, String password, RolUsuario rolUsuario, List<String> telefono, String ciudad, boolean esActivo, LocalDateTime fechaRegistro) {
        super(userName, nombre, email, password, rolUsuario, telefono, ciudad, esActivo, fechaRegistro);
        this.rfc = rfc;
        this.razonSocial = razonSocial;
        this.nombreComercial = nombreComercial;
        this.regimenFiscal = regimenFiscal;
        this.direccionFiscal = direccionFiscal;
    }

    public Cliente(String rfc, String razonSocial, String nombreComercial, String regimenFiscal, String direccionFiscal) {
        this.rfc = rfc;
        this.razonSocial = razonSocial;
        this.nombreComercial = nombreComercial;
        this.regimenFiscal = regimenFiscal;
        this.direccionFiscal = direccionFiscal;
    }

    public int getIdCliente() {
        return idCliente;
    }

    public void setIdCliente(int idCliente) {
        this.idCliente = idCliente;
    }

    public String getRfc() {
        return rfc;
    }

    public void setRfc(String rfc) {
        this.rfc = rfc;
    }

    public String getRazonSocial() {
        return razonSocial;
    }

    public void setRazonSocial(String razonSocial) {
        this.razonSocial = razonSocial;
    }

    public String getNombreComercial() {
        return nombreComercial;
    }

    public void setNombreComercial(String nombreComercial) {
        this.nombreComercial = nombreComercial;
    }

    public String getRegimenFiscal() {
        return regimenFiscal;
    }

    public void setRegimenFiscal(String regimenFiscal) {
        this.regimenFiscal = regimenFiscal;
    }

    public String getDireccionFiscal() {
        return direccionFiscal;
    }

    public void setDireccionFiscal(String direccionFiscal) {
        this.direccionFiscal = direccionFiscal;
    }

    @Override
    public String toString() {
        return "Cliente{" + "idCliente=" + idCliente + ", rfc=" + rfc + ", razonSocial=" + razonSocial + ", nombreComercial=" + nombreComercial + ", regimenFiscal=" + regimenFiscal + ", direccionFiscal=" + direccionFiscal + '}';
    }

}
