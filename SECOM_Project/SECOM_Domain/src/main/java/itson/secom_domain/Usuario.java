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

    private int id;
    private String username;
    private String nombre;
    private String email;
    private String password;
    private RolUsuario rol;
    private String telefono;
    private String ciudad;
    private boolean activo;
    private LocalDateTime fechaRegistro;

    public Usuario() {
    }

    public Usuario(int id, String username, String nombre, String email,
            String password, RolUsuario rol, String telefono,
            String ciudad, boolean activo, LocalDateTime fechaRegistro) {
        this.id = id;
        this.username = username;
        this.nombre = nombre;
        this.email = email;
        this.password = password;
        this.rol = rol;
        this.telefono = telefono;
        this.ciudad = ciudad;
        this.activo = activo;
        this.fechaRegistro = fechaRegistro;
    }

    // Constructor sin id (para insertar nuevo)
    public Usuario(String username, String nombre, String email,
            String password, RolUsuario rol, String telefono,
            String ciudad, boolean activo, LocalDateTime fechaRegistro) {
        this.username = username;
        this.nombre = nombre;
        this.email = email;
        this.password = password;
        this.rol = rol;
        this.telefono = telefono;
        this.ciudad = ciudad;
        this.activo = activo;
        this.fechaRegistro = fechaRegistro;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
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

    public RolUsuario getRol() {
        return rol;
    }

    public void setRol(RolUsuario rol) {
        this.rol = rol;
    }

    public String getTelefono() {
        return telefono;
    }

    public void setTelefono(String telefono) {
        this.telefono = telefono;
    }

    public String getCiudad() {
        return ciudad;
    }

    public void setCiudad(String ciudad) {
        this.ciudad = ciudad;
    }

    public boolean isActivo() {
        return activo;
    }

    public void setActivo(boolean activo) {
        this.activo = activo;
    }

    public LocalDateTime getFechaRegistro() {
        return fechaRegistro;
    }

    public void setFechaRegistro(LocalDateTime fechaRegistro) {
        this.fechaRegistro = fechaRegistro;
    }

    @Override
    public String toString() {
        return "Usuario{id=" + id + ", username=" + username
                + ", nombre=" + nombre + ", rol=" + rol + "}";
    }
}
