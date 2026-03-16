/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class Cliente {
    private int id;
    private String nombreCompleto;
    private String telefono;
    private String correoElectronico;
    private String direccion;
 
   public Cliente() {}
 
   public Cliente(int id, String nombreCompleto, String telefono,
                   String correoElectronico, String direccion) {
        this.id                = id;
        this.nombreCompleto    = nombreCompleto;
        this.telefono          = telefono;
        this.correoElectronico = correoElectronico;
        this.direccion         = direccion;
    }
 
  public int getId() { return id; }
    public void setId(int id) { this.id = id; }
 
    public String getNombreCompleto() { return nombreCompleto; }
    public void setNombreCompleto(String nombreCompleto) { this.nombreCompleto = nombreCompleto; }
 
    public String getTelefono() { return telefono; }
    public void setTelefono(String telefono) { this.telefono = telefono; }
 
    public String getCorreoElectronico() { return correoElectronico; }
    public void setCorreoElectronico(String correoElectronico) { this.correoElectronico = correoElectronico; }
 
    public String getDireccion() { return direccion; }
    public void setDireccion(String direccion) { this.direccion = direccion; }
 
     @Override
    public String toString() {
        return "Cliente{" +
                "id=" + id +
                ", nombreCompleto='" + nombreCompleto + '\'' +
                ", telefono='" + telefono + '\'' +
                ", correoElectronico='" + correoElectronico + '\'' +
                ", direccion='" + direccion + '\'' +
                '}';
    }
}