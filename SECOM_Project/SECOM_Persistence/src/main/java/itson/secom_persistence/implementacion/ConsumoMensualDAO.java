/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.implementacion;

import itson.secom_domain.ConsumoMensual;
import itson.secom_domain.Cotizacion;
import itson.secom_persistence.IConnectionBD;
import itson.secom_persistence.IConsumoMensualDAO;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author Serva
 */
public class ConsumoMensualDAO implements IConsumoMensualDAO{

    private final IConnectionBD connectionBD;

    public ConsumoMensualDAO(IConnectionBD connectionBD) {
        this.connectionBD = connectionBD;
    }
    
    @Override
    public void guardarTodos(List<ConsumoMensual> consumos) throws PersistenciaException {
        if(consumos == null || consumos.isEmpty()) return;
        
        String sql =
                "INSERT INTO consumos_mensuales (cotizacion_id, ano, consumo, consumo_kwh) " +
                "VALUES(?,?,?,?)";
        
        try(Connection conn = connectionBD.getConexion(); PreparedStatement cmd = conn.prepareStatement(sql)){
            for(ConsumoMensual cm : consumos){
                cmd.setInt(1, cm.getCotizacion().getId());
                cmd.setInt(2, cm.getMes());
                cmd.setInt(3, cm.getAño());
                cmd.setDouble(4, cm.getConsumoKwh());
                cmd.addBatch();
            }
            cmd.executeBatch();
        }catch(SQLException ex){
            throw new PersistenciaException("Error al guardar consumo mensuales: " + ex.getMessage(), ex);
        }
    }

    @Override
    public List<ConsumoMensual> obtenerPorCotizacion(int idCotizacion) throws PersistenciaException {
        List<ConsumoMensual> lista = new ArrayList<>();
        String sql =
                "SELECT id, cotizacion_id, mes, ano, consumo_kwh " +
                "FROM consumos_mensuales " +
                "WHERE cotizacion_id = ? " +
                "ORDER BY ano ASC, mes ASC";
        
        try(Connection conn = connectionBD.getConexion(); PreparedStatement cmd = conn.prepareStatement(sql)){
            cmd.setInt(1, idCotizacion);
            
            try(ResultSet rs = cmd.executeQuery()){
                while(rs.next()){
                    ConsumoMensual cm = new ConsumoMensual();
                    cm.setId(rs.getInt("id"));
                    cm.setMes(rs.getInt("Mes"));
                    cm.setAño(rs.getInt("ano"));
                    cm.setConsumoKwh(rs.getDouble("consumo_kwh"));
                    
                    Cotizacion ref = new Cotizacion();
                    ref.setId(rs.getInt("cotizacion_id"));
                    cm.setCotizacion(ref);
                    
                    lista.add(cm);
                }
            }
        }catch (SQLException ex){
            throw new PersistenciaException("Error al obtener consumos: " + ex.getMessage(), ex);
        }
        return lista;
    }
    
}
