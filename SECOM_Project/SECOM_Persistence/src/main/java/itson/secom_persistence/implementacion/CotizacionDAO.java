/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.implementacion;

import itson.secom_domain.Cotizacion;
import itson.secom_domain.enumeradores.EstadoCotizacion;
import itson.secom_persistence.IConnectionBD;
import itson.secom_persistence.ICotizacionDAO;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author Acer
 */
public class CotizacionDAO implements ICotizacionDAO {

    private IConnectionBD connectionBD;

    public CotizacionDAO(IConnectionBD connectionBD) {
        this.connectionBD = connectionBD;
    }

    public void guardarCotizacion(Cotizacion cotizacion) throws PersistenciaException {

        String sql = "INSERT INTO cotizaciones (fecha, consumo_promedio_mensual_kwh, total, estado) "
                + "VALUES (?, ?, ?, ?)";

        try (Connection conexion = connectionBD.getConexion(); PreparedStatement comando = conexion.prepareStatement(sql)) {

            comando.setTimestamp(1, Timestamp.valueOf(cotizacion.getFecha()));
            comando.setDouble(2, cotizacion.getConsumoPromedioMensualKwh());
            comando.setDouble(3, cotizacion.getTotal());
            comando.setString(4, cotizacion.getEstado().name());

            comando.executeUpdate();

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al guardar la cotización: " + ex.getMessage());
        }
    }

    @Override
    public List<Cotizacion> obtenerTodas() throws PersistenciaException {
        List<Cotizacion> listaCotizaciones = new ArrayList<>();
        String sql = "SELECT id, fecha, consumo_promedio_mensual_kwh, total, estado FROM cotizaciones";

        try (Connection conexion = connectionBD.getConexion(); PreparedStatement comando = conexion.prepareStatement(sql); ResultSet resultados = comando.executeQuery()) {

            while (resultados.next()) {
                Cotizacion cotizacion = new Cotizacion();
                cotizacion.setId(resultados.getInt("id"));
                cotizacion.setFecha(resultados.getTimestamp("fecha").toLocalDateTime());
                cotizacion.setConsumoPromedioMensualKwh(resultados.getDouble("consumo_promedio_mensual_kwh"));
                cotizacion.setTotal(resultados.getDouble("total"));
                cotizacion.setEstado(EstadoCotizacion.valueOf(resultados.getString("estado")));

                listaCotizaciones.add(cotizacion);
            }

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al consultar las cotizaciones: " + ex.getMessage());
        }

        return listaCotizaciones;
    }

    @Override
    public Cotizacion obtenerPorId(int id) throws PersistenciaException {
        Cotizacion cotizacion = null;
        String sql = "SELECT id, fecha, consumo_promedio_mensual_kwh, total, estado "
                + "FROM cotizaciones WHERE id = ?";

        try (Connection conexion = connectionBD.getConexion(); PreparedStatement comando = conexion.prepareStatement(sql)) {

            comando.setInt(1, id);

            try (ResultSet resultados = comando.executeQuery()) {
                if (resultados.next()) {
                    cotizacion = new Cotizacion();
                    cotizacion.setId(resultados.getInt("id"));
                    cotizacion.setFecha(resultados.getTimestamp("fecha").toLocalDateTime());
                    cotizacion.setConsumoPromedioMensualKwh(resultados.getDouble("consumo_promedio_mensual_kwh"));
                    cotizacion.setTotal(resultados.getDouble("total"));
                    cotizacion.setEstado(EstadoCotizacion.valueOf(resultados.getString("estado")));
                }
            }

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al buscar la cotización por ID: " + ex.getMessage());
        }
        return cotizacion;
    }

    @Override
    public void actualizarCotizacion(Cotizacion cotizacion) throws PersistenciaException {
        String sql = "UPDATE cotizaciones SET fecha = ?, consumo_promedio_mensual_kwh = ?, "
                + "total = ?, estado = ? WHERE id = ?";

        try (Connection conexion = connectionBD.getConexion(); PreparedStatement comando = conexion.prepareStatement(sql)) {

            comando.setTimestamp(1, Timestamp.valueOf(cotizacion.getFecha()));
            comando.setDouble(2, cotizacion.getConsumoPromedioMensualKwh());
            comando.setDouble(3, cotizacion.getTotal());
            comando.setString(4, cotizacion.getEstado().name());
            comando.setInt(5, cotizacion.getId());

            int filasAfectadas = comando.executeUpdate();

            if (filasAfectadas == 0) {
                throw new PersistenciaException("No se encontro ninguna cotizacion con el ID especificado para actualizar.");
            }
        } catch (SQLException ex) {
            throw new PersistenciaException("Error al actualizar la cotizacion: " + ex.getMessage());
        }
    }

    @Override
    public void eliminarCotizacion(int id) throws PersistenciaException {
        String sql = "DELETE FROM cotizaciones WHERE id = ?";

        try (Connection conexion = connectionBD.getConexion(); PreparedStatement comando = conexion.prepareStatement(sql)) {

            comando.setInt(1, id);

            int filasAfectadas = comando.executeUpdate();

            if (filasAfectadas == 0) {
                throw new PersistenciaException("No se encontro ninguna cotizacion con el ID especificado para eliminar.");
            }

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al eliminar la cotizacion: " + ex.getMessage());
        }
    }
}
