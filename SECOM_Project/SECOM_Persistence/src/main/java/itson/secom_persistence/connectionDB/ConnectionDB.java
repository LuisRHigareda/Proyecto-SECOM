/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.connectionDB;

import itson.secom_persistence.IConnectionBD;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;

/**
 *
 * @author Sebas
 */
public class ConnectionDB implements IConnectionBD{

    private Connection connection;
    private final String BD_REAL = "secom_pi";
    private final String BD_TEST = "secom_pi_test";

    public ConnectionDB(boolean esPrueba) {

        try (InputStream input = getClass().getClassLoader().getResourceAsStream("config.properties")) {

            Properties props = new Properties();

            if (input == null) {
                System.out.println("Error: No se pudo encontrar config.properties");
                throw new RuntimeException("Archivo config.properties no encontrado");
            }

            props.load(input);

            String usuario = props.getProperty("db.usuario");
            String contrasenia = props.getProperty("db.contrasenia");
            String ip = props.getProperty("db.ip");
            String puerto = props.getProperty("db.puerto");

            String baseDatos = esPrueba ? BD_TEST : BD_REAL;

            String url = String.format(
                    "jdbc:mysql://%s:%s/%s?useSSL=false&serverTimezone=UTC",
                    ip, puerto, baseDatos);
            
            connection = DriverManager.getConnection(url, usuario, contrasenia);

        } catch (Exception e) {
            throw new RuntimeException("Error al conectar con la base de datos", e);
        }

    }
    
    @Override
    public Connection getConexion() {
        return connection;
    }
    
    @Override
    public void close() {
        try {
            if (connection != null && !connection.isClosed()) {
                connection.close();
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

}
