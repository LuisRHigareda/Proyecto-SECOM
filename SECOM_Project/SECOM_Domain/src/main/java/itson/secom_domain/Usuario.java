/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.RolUsuario;

/**
 *
 * @author Serva
 */
public class Usuario {

    private int idUsuario;
    private String userName;
    private String contrasena;
    private String correoElectronico;
    private boolean esActivo;
    private RolUsuario rolUsuario;

    public Usuario() {

    }

    public Usuario(int id, String userName, String contrasena,
            String correoElectronico, boolean esActivo, RolUsuario rolUsuario) {
        this.idUsuario = id;
        this.userName = userName;
        this.contrasena = contrasena;
        this.correoElectronico = correoElectronico;
        this.esActivo = esActivo;
        this.rolUsuario = rolUsuario;
    }

    public int getIdUsuario() {
        return idUsuario;
    }

    public void setIdUsuario(int id) {
        this.idUsuario = id;
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

    /**
     * Obtiene si el estado de un usuario es activo o no.
     *
     * @return True si esta activo. False en caso contrario.
     */
    public boolean esActivo() {
        return esActivo;
    }

    /**
     * Establece si un usuario esta activo o no.
     *
     * @param estado True si esta activo, Falso en caso contrario.
     */
    public void setEstado(boolean estado) {
        this.esActivo = estado;
    }

    public RolUsuario getRolUsuario() {
        return rolUsuario;
    }

    public void setRolUsuario(RolUsuario rolUsuario) {
        this.rolUsuario = rolUsuario;
    }

    @Override
    public String toString() {
        return "Usuario{"
                + "id=" + idUsuario
                + ", userName='" + userName + '\''
                + ", correoElectronico='" + correoElectronico + '\''
                + ", estado='" + esActivo + '\''
                + ", tipoUsuario=" + rolUsuario.name()
                + '}';
    }
}
