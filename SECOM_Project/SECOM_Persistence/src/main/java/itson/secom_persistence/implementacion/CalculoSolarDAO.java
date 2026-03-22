/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.implementacion;

import itson.secom_domain.CalculoSolar;
import itson.secom_domain.Cotizacion;
import itson.secom_persistence.ICalculoSolarDAO;
import itson.secom_persistence.IConnectionBD;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;

/**
 *
 * @author Serva
 */
public class CalculoSolarDAO implements ICalculoSolarDAO {

    private final IConnectionBD connectionBD;

    public CalculoSolarDAO(IConnectionBD connectionBD) {
        this.connectionBD = connectionBD;
    }

    @Override
    public void guardar(CalculoSolar cs) throws PersistenciaException {
        String sql
                = "INSERT INTO Calculo_Solar_cotizacion "
                + "(cotizacion_id, estado, insolacion_usada, potencial_panel, numero_paneles "
                + " watts_instalados, capacidad_inversor, produccion_diaria_estimada, "
                + " produccion_anual_estimada, porcentaje_generacion, "
                + " factor_conversion_usada, factor_reflexion_usado) "
                + "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)";

        try (Connection conn = connectionBD.getConexion(); PreparedStatement cmd = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {

            cmd.setInt(1, cs.getCotizacion().getId());
            cmd.setString(2, cs.getEstadoMX() != null ? cs.getEstadoMX() : "");
            cmd.setDouble(3, cs.getInsolacionUsada());
            cmd.setDouble(4, cs.getPotencialPanel());
            cmd.setInt(5, cs.getNumeroPaneles());
            cmd.setDouble(6, cs.getWattsInstalados());
            cmd.setDouble(7, cs.getCapacidadInversor());
            cmd.setDouble(8, cs.getProduccionDiariaEstimada());
            cmd.setDouble(9, cs.getProduccionAnualEstimada());
            cmd.setDouble(10, cs.getPorcentajeGeneracion());
            cmd.setDouble(11, cs.getFactorConversionUsado());
            cmd.setDouble(12, cs.getFactorReflexionUsado());

            cmd.executeUpdate();

            try (ResultSet keys = cmd.getGeneratedKeys()) {
                if (keys.next()) {
                    cs.setId(keys.getInt(1));
                }
            }

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al guardar cálculo solar: " + ex.getMessage(), ex);
        }
    }

    @Override
    public CalculoSolar obtenerPorCotizacion(int idCotizacion) throws PersistenciaException {
        String sql
                = "SELECT id, cotizacion_id, estado, insolacion_usada, potencia_panel, "
                + "       numero_paneles, watts_instalados, capacidad_inversor, "
                + "       produccion_diaria_estimada, produccion_anual_estimada, "
                + "       porcentaje_generacion, factor_conversion_usado, "
                + "       factor_reflexion_usado, fecha_calculo "
                + "FROM calculo_solar_cotizacion "
                + "WHERE cotizacion_id = ? "
                + "ORDER BY fecha_calculo DESC LIMIT 1";

        try (Connection conn = connectionBD.getConexion(); PreparedStatement cmd = conn.prepareStatement(sql)) {
            cmd.setInt(1, idCotizacion);
            
            try(ResultSet rs = cmd.executeQuery()){
                if(rs.next()){
                    CalculoSolar cs = new CalculoSolar();
                    cs.setId(rs.getInt("id"));
                    cs.setEstadoMX(rs.getString("estado"));
                    cs.setInsolacionUsada(rs.getDouble("insolacion_usada"));
                    cs.setPotencialPanel(rs.getInt("potencia_panel"));
                    cs.setNumeroPaneles(rs.getInt("watts_instalados"));
                    cs.setWattsInstalados(rs.getDouble("watts_instalados"));
                    cs.setCapacidadInversor(rs.getDouble("capacidad_inversor"));
                    cs.setProduccionDiariaEstimada(rs.getDouble("produccion_diaria_estimada"));
                    cs.setProduccionAnualEstimada(rs.getDouble("produccion_anual_estimada"));
                    cs.setPorcentajeGeneracion(rs.getDouble("porcentaje_generacion"));
                    cs.setFactorConversionUsado(rs.getDouble("factor_conversion_usado"));
                    cs.setFactorReflexionUsado(rs.getDouble("factor_refleccion_usado"));
                    
                    Timestamp ts = rs.getTimestamp("fecha_calculo");
                    if (ts != null) cs.setFechaCalculo(ts.toLocalDateTime()); 
                        
                    Cotizacion ref = new Cotizacion();
                    ref.setId(rs.getInt("cotizacion_id"));
                    cs.setCotizacion(ref);
                    
                    return cs;
                }
            }
        } catch (SQLException ex){
            throw new PersistenciaException("Error al obtener calculo solar: " + ex.getMessage(), ex);
        }
        return null;
    }

}
