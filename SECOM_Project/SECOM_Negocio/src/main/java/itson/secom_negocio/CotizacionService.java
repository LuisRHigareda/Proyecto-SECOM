/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_negocio;

import itson.secom_domain.Cotizacion;
import itson.secom_persistence.ICotizacionDAO;
import itson.secom_persistence.conexionFactory.DAOFactory;
import itson.secom_persistence.excepciones.PersistenciaException;
import itson.secom_persistence.implementacion.CotizacionDAO;
import java.util.List;

/**
 *
 * @author Acer
 */
public class CotizacionService {
    private ICotizacionDAO cotizacionDAO;

    public CotizacionService() {
        DAOFactory factory = new DAOFactory(false);
        this.cotizacionDAO = factory.conexionCotizacionDAO();
    }

    public void guardarCotizacion(Cotizacion cotizacion) throws Exception {
        if (cotizacion.getConsumoPromedioMensualKwh() <= 0) {
            throw new Exception("El consumo promedio debe ser mayor a 0.");
        }
        if (cotizacion.getTotal() < 0) {
            throw new Exception("El total de la cotizacion no puede ser negativo.");
        }
        try {
            cotizacionDAO.guardarCotizacion(cotizacion);
        } catch (PersistenciaException ex) {
            throw new Exception("Error al guardar en la base de datos: " + ex.getMessage());
        }
    }

    public List<Cotizacion> obtenerTodas() throws Exception {
        try {
            return cotizacionDAO.obtenerTodas();
        } catch (PersistenciaException ex) {
            throw new Exception("Error al obtener las cotizaciones: " + ex.getMessage());
        }
    }

    public Cotizacion obtenerPorId(int id) throws Exception {
        if (id <= 0) {
            throw new Exception("El ID debe ser un número valido mayor a 0.");
        }
        try {
            return cotizacionDAO.obtenerPorId(id);
        } catch (PersistenciaException ex) {
            throw new Exception("Error al buscar la cotizacion: " + ex.getMessage());
        }
    }

    public void actualizarCotizacion(Cotizacion cotizacion) throws Exception {
        if (cotizacion.getId() <= 0) {
            throw new Exception("Se requiere un ID valido para actualizar.");
        }
        if (cotizacion.getTotal() < 0) {
            throw new Exception("El total no puede ser negativo.");
        }
        try {
            cotizacionDAO.actualizarCotizacion(cotizacion);
        } catch (PersistenciaException ex) {
            throw new Exception("Error al actualizar la cotizacion: " + ex.getMessage());
        }
    }

    public void eliminarCotizacion(int id) throws Exception {
        if (id <= 0) {
            throw new Exception("El ID a eliminar debe ser mayor a 0.");
        }
        try {
            cotizacionDAO.eliminarCotizacion(id);
        } catch (PersistenciaException ex) {
            throw new Exception("Error al eliminar la cotizacion: " + ex.getMessage());
        }
    }
}
