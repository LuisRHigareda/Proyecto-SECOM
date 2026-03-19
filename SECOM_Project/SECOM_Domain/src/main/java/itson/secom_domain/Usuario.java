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
public class Usuario {

    private int idUsuario;
    private String userName;
    private String nombre;
    private String email;
    private String password;
    private RolUsuario rolUsuario;
    private List<String> telefono;
    private String ciudad;
    private boolean esActivo;
    private LocalDateTime fechaRegistro;

    public Usuario() {
    }

    public Usuario(int idUsuario, String userName, String nombre, String email, String password, RolUsuario rolUsuario, List<String> telefono, String ciudad, boolean esActivo, LocalDateTime fechaRegistro) {
        this.idUsuario = idUsuario;
        this.userName = userName;
        this.nombre = nombre;
        this.email = email;
        this.password = password;
        this.rolUsuario = rolUsuario;
        this.telefono = telefono;
        this.ciudad = ciudad;
        this.esActivo = esActivo;
        this.fechaRegistro = fechaRegistro;
    }

    public Usuario(String userName, String nombre, String email, String password, RolUsuario rolUsuario, List<String> telefono, String ciudad, boolean esActivo, LocalDateTime fechaRegistro) {
        this.userName = userName;
        this.nombre = nombre;
        this.email = email;
        this.password = password;
        this.rolUsuario = rolUsuario;
        this.telefono = telefono;
        this.ciudad = ciudad;
        this.esActivo = esActivo;
        this.fechaRegistro = fechaRegistro;
    }

    public int getIdUsuario() {
        return idUsuario;
    }

    public void setIdUsuario(int idUsuario) {
        this.idUsuario = idUsuario;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public RolUsuario getRolUsuario() {
        return rolUsuario;
    }

    public void setRolUsuario(RolUsuario rolUsuario) {
        this.rolUsuario = rolUsuario;
    }

    public List<String> getTelefono() {
        return telefono;
    }

    public void setTelefono(List<String> telefono) {
        this.telefono = telefono;
    }

    public String getCiudad() {
        return ciudad;
    }

    public void setCiudad(String ciudad) {
        this.ciudad = ciudad;
    }

    public boolean isEsActivo() {
        return esActivo;
    }

    public void setEsActivo(boolean esActivo) {
        this.esActivo = esActivo;
    }

    public LocalDateTime getFechaRegistro() {
        return fechaRegistro;
    }

    public void setFechaRegistro(LocalDateTime fechaRegistro) {
        this.fechaRegistro = fechaRegistro;
    }

    @Override
    public String toString() {
        return "Usuario{" + "idUsuario=" + idUsuario + ", userName=" + userName + ", nombre=" + nombre + ", email=" + email + ", password=" + password + ", rolUsuario=" + rolUsuario + ", telefono=" + telefono + ", ciudad=" + ciudad + ", esActivo=" + esActivo + ", fechaRegistro=" + fechaRegistro + '}';
    }

}
