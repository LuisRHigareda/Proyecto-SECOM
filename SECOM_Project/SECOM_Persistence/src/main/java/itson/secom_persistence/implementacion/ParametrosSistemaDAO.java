/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
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
 *
 * @author Arell
 */
public class ParametrosSistemaDAO implements IParametrosSistemaDAO {
    




    private final IConnectionBD connectionBD;
    private Connection conexion;

    public ParametrosSistemaDAO(IConnectionBD connectionBD) {
        this.connectionBD = connectionBD;
    }

    @Override
    public ParametrosSistema obtenerParametros(String ciudad) throws PersistenciaException {

        String sqlEficiencia = "SELECT valor FROM parametros_sistema WHERE clave = 'EFICIENCIA' LIMIT 1";
        String sqlIva = "SELECT valor FROM parametros_sistema WHERE clave = 'IVA' LIMIT 1";
        String sqlPrecioKwh = "SELECT valor FROM parametros_sistema WHERE clave = 'PRECIO_KWH_REFERENCIA' LIMIT 1";
        String sqlHsp = "SELECT horas_sol_pico FROM insolacion_solar WHERE ciudad = ? LIMIT 1";

        try {
            conexion = connectionBD.getConexion();

            double eficiencia = 0;
            double iva = 0;
            double precioKwh = 0;
            double hsp = 0;

            try (PreparedStatement ps = conexion.prepareStatement(sqlEficiencia);
                 ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    eficiencia = rs.getDouble(1);
                }
            }

            try (PreparedStatement ps = conexion.prepareStatement(sqlIva);
                 ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    iva = rs.getDouble(1);
                }
            }

            try (PreparedStatement ps = conexion.prepareStatement(sqlPrecioKwh);
                 ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    precioKwh = rs.getDouble(1);
                }
            }

            try (PreparedStatement ps = conexion.prepareStatement(sqlHsp)) {
                ps.setString(1, ciudad);
                try (ResultSet rs = ps.executeQuery()) {
                    if (rs.next()) {
                        hsp = rs.getDouble(1);
                    }
                }
            }

            return new ParametrosSistema(eficiencia, hsp, iva, precioKwh);

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al obtener parámetros del sistema: " + ex.getMessage(), ex);
        }
    }

    @Override
    public void close() {
        try {
            if (conexion != null && !conexion.isClosed()) {
                conexion.close();
            }
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }
}
