/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class Cliente extends Usuario {

    private int idCliente;
    private String nombreCompleto;
    private String telefono;
    private String direccion;

    public Cliente() {
    }

    public Cliente(int id, String nombreCompleto, String telefono,
            String direccion) {
        this.idCliente = id;
        this.nombreCompleto = nombreCompleto;
        this.telefono = telefono;
        this.direccion = direccion;
    }

    public int getIdCliente() {
        return idCliente;
    }

    public void setIdCliente(int id) {
        this.idCliente = id;
    }

    public String getNombreCompleto() {
        return nombreCompleto;
    }

    public void setNombreCompleto(String nombreCompleto) {
        this.nombreCompleto = nombreCompleto;
    }

    public String getTelefono() {
        return telefono;
    }

    public void setTelefono(String telefono) {
        this.telefono = telefono;
    }

    public String getDireccion() {
        return direccion;
    }

    public void setDireccion(String direccion) {
        this.direccion = direccion;
    }

    @Override
    public String toString() {
        return "Cliente{"
                + "id=" + idCliente
                + ", nombreCompleto='" + nombreCompleto + '\''
                + ", telefono='" + telefono + '\''
                + ", direccion='" + direccion + '\''
                + '}';
    }
}
