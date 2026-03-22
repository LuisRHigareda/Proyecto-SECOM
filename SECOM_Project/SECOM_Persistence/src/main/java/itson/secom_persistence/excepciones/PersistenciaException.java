/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.excepciones;

/**
 *
 * @author PC
 */
public class PersistenciaException extends Exception {

    /**
     * Constructor vacio
     */
    public PersistenciaException() {
    }

    /**
     * Constructor que establece el mensaje de la excepcion.
     * @param message mensaje de la excepcion
     */
    public PersistenciaException(String message) {
        super(message);
    }

    /**
     * Constructor que establece el mensaje y la causa de la excepcion.
     * @param message mensaje de la excepcion
     * @param cause causa de la excepcion
     */
    public PersistenciaException(String message, Throwable cause) {
        super(message, cause);
    }

}
