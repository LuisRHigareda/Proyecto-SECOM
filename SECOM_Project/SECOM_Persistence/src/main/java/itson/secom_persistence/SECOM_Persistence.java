/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package itson.secom_persistence;

import itson.secom_persistence.connectionDB.ConnectionDB;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.sql.Connection;

/**
 *
 * @author Sebas
 */
public class SECOM_Persistence {

    public static void main(String[] args) throws PersistenciaException {
        ConnectionDB conexion = new ConnectionDB(false);
        
        Connection conn = conexion.getConexion();
        
        System.out.println("Conectado a MySQL");
        
        
        conexion.close();
    }
}
