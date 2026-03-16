/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

/**
 *
 * @author Serva
 */
public class Rol {

    private int id;
    private String tipoRol;
    private String descripcion;    

    public Rol() {
    }

    public Rol(int id, String tipoRol, String descripcion) {
        this.id = id;
        this.tipoRol = tipoRol;
        this.descripcion = descripcion;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getTipoRol() {
        return tipoRol;
    }

    public void setTipoRol(String tipoRol) {
        this.tipoRol = tipoRol;
    }

    public String getDescripcion() {
        return descripcion;
    }

    public void setDescripcion(String descripcion) {
        this.descripcion = descripcion;
    }

    @Override
    public String toString() {
        return "Rol{"
                + "id=" + id
                + ", tipoRol='" + tipoRol + '\''
                + ", descripcion='" + descripcion + '\''
                + '}';
    }
}
