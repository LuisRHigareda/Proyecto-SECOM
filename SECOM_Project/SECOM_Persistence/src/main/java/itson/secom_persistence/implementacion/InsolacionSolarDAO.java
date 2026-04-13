/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.implementacion;

import itson.secom_persistence.IConnectionBD;
import itson.secom_persistence.IInsolacionSolarDAO;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 *
 * @author Arell
 */
public class InsolacionSolarDAO implements IInsolacionSolarDAO {

    private final Connection conexion;

    public InsolacionSolarDAO(IConnectionBD connectionBD) throws PersistenciaException {
        this.conexion = connectionBD.getConexion();
    }

    @Override
    public double obtenerHspPorCiudad(String ciudad) throws PersistenciaException {
        String sql = "SELECT hsp FROM insolacion_solar WHERE UPPER(ciudad) = UPPER(?)";

        try (PreparedStatement ps = conexion.prepareStatement(sql)) {
            ps.setString(1, ciudad);

            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return rs.getDouble("hsp");
                } else {
                    throw new PersistenciaException(
                            "No se encontró la insolación para la ciudad: " + ciudad);
                }
            }
        } catch (SQLException ex) {
            throw new PersistenciaException(
                    "Error al obtener la insolación solar: " + ex.getMessage(), ex);
        }
    }

    @Override
    public void close() {
        // No cerrar la conexión aquí si es administrada por la fábrica
    }
}
