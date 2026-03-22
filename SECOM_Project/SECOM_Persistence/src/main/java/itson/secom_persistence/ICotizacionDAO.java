/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Interface.java to edit this template
 */
package itson.secom_persistence;

import itson.secom_domain.Cotizacion;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.util.List;

/**
 *
 * @author Acer
 */
public interface ICotizacionDAO {
    void guardarCotizacion(Cotizacion cotizacion) throws PersistenciaException;
    
    List<Cotizacion> obtenerTodas()throws PersistenciaException;
    
    Cotizacion obtenerPorId(int id) throws PersistenciaException;
    
    void actualizarCotizacion(Cotizacion cotizacion) throws PersistenciaException;
    
    void eliminarCotizacion(int id) throws PersistenciaException;
}
