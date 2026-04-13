/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Interface.java to edit this template
 */
package itson.secom_persistence;

import itson.secom_persistence.excepciones.PersistenciaException;

/**
 *
 * @author Arell
 */
public interface IInsolacionSolarDAO extends AutoCloseable {
    double obtenerHspPorCiudad(String ciudad) throws PersistenciaException;
}
