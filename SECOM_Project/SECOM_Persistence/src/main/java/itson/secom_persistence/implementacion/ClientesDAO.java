/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.implementacion;

import itson.secom_domain.Cliente;
import itson.secom_domain.enumeradores.RolUsuario;
import itson.secom_persistence.IClientesDAO;
import itson.secom_persistence.IConnectionBD;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author PC
 */
public class ClientesDAO implements IClientesDAO {

    private IConnectionBD connectionBD;

    public ClientesDAO(IConnectionBD connectionBD) {
        this.connectionBD = connectionBD;
    }
    
    private static final String SQL_BASE =
            "SELECT c.id AS id_cliente, c.usuario_id, c.rfc, c.razon_social, " +
            "       c.nombre_comercial, c.regimen_fiscal, c.direccion_fiscal, " +
            "       c.ciudad AS ciudad_cliente, c.activo AS cliente_activo, " +
            "       u.id AS id_usuario, u.username, u.nombre, u.email, " +
            "       u.password, u.rol, u.telefono, u.ciudad, " +
            "       u.activo, u.fecha_registro " +
            "FROM clientes c " +
            "JOIN usuarios u ON c.usuario_id = u.id " +
            "WHERE c.deleted_at IS NULL AND u.deleted_At IS NULL ";

    @Override
    public List<Cliente> obtenerClientes() throws PersistenciaException {
        List<Cliente> lista = new ArrayList<>();
        String sql = SQL_BASE + "ORDER BY c.id";
        
        try(Connection conn = connectionBD.getConexion();
            PreparedStatement  cmd = conn.prepareStatement(sql);
            ResultSet rs = cmd.executeQuery()){
            
            while(rs.next()){
                lista.add(mapear(rs));
            }
            
        } catch (SQLException ex) {
            throw new PersistenciaException("Error al obtener clientes: " + ex.getMessage(), ex);
        }
        return lista;
    }

    @Override
    public Cliente obtenerCliente(int id) throws PersistenciaException {
        String sql = SQL_BASE + "AND c.id = ?";
        
        try(Connection conn = connectionBD.getConexion();
            PreparedStatement cmd = conn.prepareStatement(sql)){
            
            cmd.setInt(1, id);
            try(ResultSet rs = cmd.executeQuery()){
                if(rs.next()) return mapear(rs);
            }
            
        } catch (SQLException ex) {
            throw new PersistenciaException("Error al obtener cliente id=" + id + ": " + ex.getMessage(), ex);
        }
        return null;
    }
    
    
    private Cliente mapear(ResultSet rs) throws SQLException {
        Cliente c = new Cliente();
        c.setIdCliente(rs.getInt("id_cliente"));
        c.setId(rs.getInt("usuario_id"));
        c.setRfc(rs.getString("rfc"));
        c.setRazonSocial(rs.getString("razon_social"));
        c.setNombreComercial(rs.getString("nombre_comercial"));
        c.setRegimenFiscal(rs.getString("regimen_fiscal"));
        c.setDireccionFiscal(rs.getString("direccion_fiscal"));
        c.setClienteActivo(rs.getBoolean("cliente_activo"));
        
        c.setId(rs.getInt("id_usuario"));
        c.setUsername(rs.getString("username"));
        c.setNombre(rs.getString("nombre"));
        c.setEmail(rs.getString("email"));
        c.setPassword(rs.getString("password"));
        c.setTelefono(rs.getString("telefono"));
        c.setCiudad(rs.getString("ciudad"));
        c.setActivo(rs.getBoolean("activo"));
        
        String rolStr = rs.getString("rol");
        if (rolStr != null && !rolStr.isEmpty()) {
            String primerRol = rolStr.split(",")[0].trim().toUpperCase();
            try{
                c.setRol(RolUsuario.valueOf(primerRol));
            } catch(IllegalArgumentException e){
                c.setRol(RolUsuario.CLIENTE);
            }
        }
        
        java.sql.Timestamp ts = rs.getTimestamp("fecha_registro");
        if (ts != null) c.setFechaRegistro(ts.toLocalDateTime());
        
        return c;
    }
}
