/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.RolUsuario;
import java.time.LocalDateTime;

/**
 *
 * @author Serva
 */
public class Vendedor extends Usuario {

    private double porcentajeComision;

    public Vendedor() {
    }

    public Vendedor(int idUsuario, String username, String nombre, String email,
            String password, RolUsuario rol, String telefono,
            String ciudad, boolean activo, LocalDateTime fechaRegistro,
            double porcentajeComision) {
        super(idUsuario, username, nombre, email, password, rol,
                telefono, ciudad, activo, fechaRegistro);
        this.porcentajeComision = porcentajeComision;
    }

    public Vendedor(int idUsuario, double porcentajeComision) {
        this.setId(idUsuario);
        this.porcentajeComision = porcentajeComision;
    }

    public double getPorcentajeComision() {
        return porcentajeComision;
    }

    public void setPorcentajeComision(double porcentajeComision) {
        this.porcentajeComision = porcentajeComision;
    }

    public int getUsuarioId() {
        return this.getId();
    }

    @Override
    public String toString() {
        return "Vendedor{id=" + getId() + ", nombre=" + getNombre()
                + ", comision=" + porcentajeComision + "%}";
    }

}
