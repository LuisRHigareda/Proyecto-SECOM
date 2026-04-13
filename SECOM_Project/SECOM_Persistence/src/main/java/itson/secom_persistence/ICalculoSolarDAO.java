/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Interface.java to edit this template
 */
package itson.secom_persistence;

import itson.secom_domain.CalculoSolar;
import itson.secom_persistence.excepciones.PersistenciaException;

/**
 *
 * @author Serva
 */
public interface ICalculoSolarDAO {
    
    void guardar(CalculoSolar calculo) throws PersistenciaException;
    
    CalculoSolar obtenerPorCotizacion(int idCotizacion) throws PersistenciaException;
    
}
