/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Interface.java to edit this template
 */
package itson.secom_persistence;

import itson.secom_domain.ConsumoMensual;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.util.List;

/**
 *
 * @author Serva
 */
public interface IConsumoMensualDAO {
    
    void guardarTodos(List<ConsumoMensual> consumos) throws PersistenciaException;
    
    List<ConsumoMensual> obtenerPorCotizacion(int idCotizacion) throws PersistenciaException;
    
}
