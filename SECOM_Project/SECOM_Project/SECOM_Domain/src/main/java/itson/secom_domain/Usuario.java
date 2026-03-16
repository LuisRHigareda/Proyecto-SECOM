/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class Usuario {

    private int id;
    private String userName;
    private String contrasena;
    private String correoElectronico;
    private String estado;
    private Rol tipoUsuario;

    public Usuario() {

    }

    public Usuario(int id, String userName, String contrasena,
            String correoElectronico, String estado, Rol tipoUsuario) {
        this.id = id;
        this.userName = userName;
        this.contrasena = contrasena;
        this.correoElectronico = correoElectronico;
        this.estado = estado;
        this.tipoUsuario = tipoUsuario;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getContrasena() {
        return contrasena;
    }

    public void setContrasena(String contrasena) {
        this.contrasena = contrasena;
    }

    public String getCorreoElectronico() {
        return correoElectronico;
    }

    public void setCorreoElectronico(String correoElectronico) {
        this.correoElectronico = correoElectronico;
    }

    public String getEstado() {
        return estado;
    }

    public void setEstado(String estado) {
        this.estado = estado;
    }

    public Rol getTipoUsuario() {
        return tipoUsuario;
    }

    public void setTipoUsuario(Rol tipoUsuario) {
        this.tipoUsuario = tipoUsuario;
    }

    @Override
    public String toString() {
        return "Usuario{"
                + "id=" + id
                + ", userName='" + userName + '\''
                + ", correoElectronico='" + correoElectronico + '\''
                + ", estado='" + estado + '\''
                + ", tipoUsuario=" + (tipoUsuario != null ? tipoUsuario.getTipoRol() : "null")
                + '}';
    }
}
