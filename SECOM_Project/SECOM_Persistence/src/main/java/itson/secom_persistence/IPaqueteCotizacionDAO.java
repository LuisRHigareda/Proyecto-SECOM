/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Interface.java to edit this template
 */
package itson.secom_persistence;

/**
 *
 * @author Arell
 */
import itson.secom_domain.ProductoCantidad;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.util.List;

public interface IPaqueteCotizacionDAO extends AutoCloseable {

    List<ProductoCantidad> obtenerProductosPorPaquete(int paqueteId) throws PersistenciaException;

    @Override
    void close();
}
