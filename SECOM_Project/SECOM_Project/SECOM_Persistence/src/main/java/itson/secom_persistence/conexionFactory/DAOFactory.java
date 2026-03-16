/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.conexionFactory;

import itson.secom_persistence.IConnectionBD;
import itson.secom_persistence.connectionDB.ConnectionDB;

/**
 *
 * @author Sebas
 */
public class DAOFactory {

    private final IConnectionBD conexion;

    /**
     * Crea la conexion con una base de datos y utilizar los DAOs necesarios
     *
     * @param esPrueba TRUE si se utiliza la base de datos de prueba. FALSE en
     * caso contrario.
     */
    public DAOFactory(boolean esPrueba) {
        this.conexion = new ConnectionDB(esPrueba);
    }
    

    /**
     *
     * --- Ejemplo de Uso --- Cuando crezca el proyecto, se necesitara hacer lo
     * siguiente para crear una conexion para el resto de clases DAO
     *
     * Ejemplo con UsuariosDAO y ProductosDAO:
     *
     * public UsuariosDAO conexionUsuariosDAO() {
     *    return new UsuariosDAO(conexion);
     * }
     * 
     * public ProductosDAO conexionProductosDAO() {
     *    return new ProductosDAO(conexion);
     * }
     * 
     * 
     * Se llamaria de la siguiente manera para utilizarse:
     * (seria ture si fuera de prueba)
     * DAOFactory factory = new DAOFactory(false);
     *
     * UsuariosDAO usuariosDAO = factory.crearUsuariosDAO();
     *
     */
}
