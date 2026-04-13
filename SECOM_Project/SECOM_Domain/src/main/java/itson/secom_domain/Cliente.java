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
    private String ciudad;
    private boolean clienteActivo;

    public Cliente() {
    }


    public Cliente(int idCliente, String rfc, String razonSocial,
            String nombreComercial, String regimenFiscal, String direccionFiscal,
            String ciudad,int idUsuario, String username, String nombre, String email,
            String password, RolUsuario rol, String telefono,
            boolean activo, LocalDateTime fechaRegistro) {
        
        super(idUsuario, username, nombre, email, password, rol,
                telefono, ciudad, activo, fechaRegistro);
        this.idCliente = idCliente;
        this.rfc = rfc;
        this.razonSocial = razonSocial;
        this.nombreComercial = nombreComercial;
        this.regimenFiscal = regimenFiscal;
        this.direccionFiscal = direccionFiscal;
        this.ciudad = ciudad;
        this.clienteActivo = activo;
    }


    public Cliente(int idCliente, String rfc, String razonSocial,
            String nombreComercial, String regimenFiscal, String direccionFiscal) {
        this.idCliente = idCliente;
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

    public boolean isClienteActivo() {
        return clienteActivo;
    }

    public void setClienteActivo(boolean clienteActivo) {
        this.clienteActivo = clienteActivo;
    }

    @Override
    public String toString() {
        return "Cliente{idCliente=" + idCliente + ", rfc=" + rfc
                + ", nombreComercial=" + nombreComercial + "}";
    }
}
