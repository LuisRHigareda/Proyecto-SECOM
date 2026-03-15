/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Interface.java to edit this template
 */
package itson.secom_persistence;

import java.sql.Connection;

/**
 *
 * @author Sebas
 */
public interface IConnectionBD {
    
    Connection getConexion();
    
    void close();
}
