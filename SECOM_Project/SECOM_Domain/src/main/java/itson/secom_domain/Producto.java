/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_domain;

import itson.secom_domain.enumeradores.CategoriaProducto;

/**
 *
 * @author Serva
 */
public class Producto {

    private int id;
    private String nombre;
    private CategoriaProducto categoria;
    private String espesificaciones;
    private double precioUnitario;
    private int stock;
    private int capacidad;

    public Producto() {
    }

    public Producto(int id, String nombre, CategoriaProducto categoria,
            String espesificaciones, double precioUnitario, int stock) {
        this.id = id;
        this.nombre = nombre;
        this.categoria = categoria;
        this.espesificaciones = espesificaciones;
        this.precioUnitario = precioUnitario;
        this.stock = stock;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public CategoriaProducto getCategoria() {
        return categoria;
    }

    public void setCategoria(CategoriaProducto categoria) {
        this.categoria = categoria;
    }

    public String getEspesificaciones() {
        return espesificaciones;
    }

    public void setEspesificaciones(String espesificaciones) {
        this.espesificaciones = espesificaciones;
    }

    public double getPrecioUnitario() {
        return precioUnitario;
    }

    public void setPrecioUnitario(double precioUnitario) {
        this.precioUnitario = precioUnitario;
    }

    public int getStock() {
        return stock;
    }

    public void setStock(int stock) {
        this.stock = stock;
    }

    public int getCapacidad() {
        return capacidad;
    }

    public void setCapacidad(int capacidad) {
        this.capacidad = capacidad;
    }
    

    @Override
    public String toString() {
        return "Producto{"
                + "id=" + id
                + ", nombre='" + nombre + '\''
                + ", categoria='" + categoria + '\''
                + ", precioUnitario=" + precioUnitario
                + ", stock=" + stock
                + '}';
    }
}
