package itson.secom_persistence.implementacion;

import itson.secom_domain.ParametrosSistema;
import itson.secom_persistence.IConnectionBD;
import itson.secom_persistence.IParametrosSistemaDAO;
import itson.secom_persistence.excepciones.PersistenciaException;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * DAO encargado de obtener los parámetros del sistema necesarios para
 * el cálculo de la cotización solar.
 *
 * Obtiene:
 * - Eficiencia del sistema
 * - IVA
 * - Precio de referencia por kWh
 * - Horas Sol Pico (HSP) según la ciudad
 *
 * @author Arell
 */


/**
 * DAO encargado de obtener los parámetros del sistema necesarios
 * para el cálculo de la cotización solar.
 */
public class ParametrosSistemaDAO implements IParametrosSistemaDAO {

    private final IConnectionBD connectionBD;

    public ParametrosSistemaDAO(IConnectionBD connectionBD) {
        this.connectionBD = connectionBD;
    }

    @Override
    public ParametrosSistema obtenerParametros(String ciudad) throws PersistenciaException {

        String sqlParametro = "SELECT valor FROM parametros_sistema WHERE clave = ?";
        String sqlHsp = "SELECT hsp FROM insolacion_solar "
                + "WHERE UPPER(TRIM(ciudad)) = UPPER(TRIM(?)) LIMIT 1";

        try (Connection conexion = connectionBD.getConexion()) {

            double eficiencia = obtenerParametro(conexion, sqlParametro, "EFICIENCIA");
            double iva = obtenerParametro(conexion, sqlParametro, "IVA");
            double precioKwh = obtenerParametro(conexion, sqlParametro, "PRECIO_KWH_REFERENCIA");
            double factorConversion = obtenerParametro(conexion, sqlParametro, "FACTOR_CONVERSION");
            double factorSistema = obtenerParametro(conexion, sqlParametro, "FACTOR_SISTEMA");

            double hsp;
            try (PreparedStatement ps = conexion.prepareStatement(sqlHsp)) {
                ps.setString(1, ciudad);
                try (ResultSet rs = ps.executeQuery()) {
                    if (rs.next()) {
                        hsp = rs.getDouble("hsp");
                    } else {
                        throw new PersistenciaException(
                                "No se encontró la insolación solar para la ciudad: " + ciudad);
                    }
                }
            }

            return new ParametrosSistema(
                    eficiencia,
                    hsp,
                    iva,
                    precioKwh,
                    factorConversion,
                    factorSistema
            );

        } catch (SQLException ex) {
            throw new PersistenciaException(
                    "Error al obtener parámetros del sistema: " + ex.getMessage(), ex);
        }
    }

    private double obtenerParametro(Connection conexion, String sql, String clave)
            throws SQLException, PersistenciaException {

        try (PreparedStatement ps = conexion.prepareStatement(sql)) {
            ps.setString(1, clave);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return rs.getDouble("valor");
                } else {
                    throw new PersistenciaException(
                            "No se encontró el parámetro del sistema: " + clave);
                }
            }
        }
    }

    @Override
    public void close() {
        // No se requiere implementación.
    }
}