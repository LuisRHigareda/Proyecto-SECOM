/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Interface.java to edit this template
 */
package itson.secom_persistence;

import itson.secom_domain.ParametrosSistema;
import itson.secom_persistence.excepciones.PersistenciaException;

/**
 *
 * @author Arell
 */
public interface IParametrosSistemaDAO extends AutoCloseable {
    

    ParametrosSistema obtenerParametros(String ciudad) throws PersistenciaException;

    @Override
    void close();
}
