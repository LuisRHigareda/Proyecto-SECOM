/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.implementacion;

import itson.secom_domain.Cliente;
import itson.secom_domain.Cotizacion;
import itson.secom_domain.Vendedor;
import itson.secom_domain.enumeradores.EstadoCotizacion;
import itson.secom_persistence.IConnectionBD;
import itson.secom_persistence.ICotizacionDAO;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author Acer
 */
public class CotizacionDAO implements ICotizacionDAO {

    private final IConnectionBD connectionBD;

    public CotizacionDAO(IConnectionBD connectionBD) {
        this.connectionBD = connectionBD;
    }

    private static final String SQL_SELECT
            = "SELECT co.id, co.vendedor_id, co.cliente_id, co.paquete_id, "
            + "       co.fecha, co.estado, "
            + "       co.consumo_promedio_mensual_kwh, co.consumo_promedio_diario_kwh, "
            + "       co.costo_promedio_mensual, co.costo_promedio_anual, "
            + "       co.watts_instalados, co.produccion_diaria_estimada, "
            + "       co.porcentaje_cobertura, co.retorno_inversion, "
            + "       co.subtotal, co.iva, co.total, "
            + "       co.financiamiento, co.proyecto_generado, co.notas, "
            + "       co.created_by, co.updated_by, "
            + "       cl.nombre_comercial, cl.rfc, "
            + "       u.nombre AS nombre_vendedor "
            + "FROM cotizaciones co "
            + "JOIN clientes cl ON co.cliente_id = cl.id "
            + "LEFT JOIN usuarios u ON co.vendedor_id = u.id "
            + "WHERE co.deleted_at IS NULL ";

    @Override
    public void guardarCotizacion(Cotizacion c) throws PersistenciaException {
        if (c.getCliente() == null || c.getCliente().getIdCliente() <= 0) {
            throw new PersistenciaException("La cotizacion necesita un cliente valido.");
        }

        String sql
                = "INSERT INTO cotizaciones "
                + "(vendedor_id, cliente_id, paquete_id, fecha, estado, "
                + " consumo_promedio_mensual_kwh, consumo_promedio_diario_kwh, "
                + " costo_promedio_mensual, costo_promedio_anual, "
                + " watts_instalados, produccion_diaria_estimada, "
                + " porcentaje_cobertura, retorno_inversion, "
                + " subtotal, iva, total, financiamiento, notas, created_by) "
                + "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

        Connection conn = connectionBD.getConexion();
try (PreparedStatement cmd = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {

            if (c.getVendedor() != null && c.getVendedor().getUsuarioId() > 0) {
                cmd.setInt(1, c.getVendedor().getUsuarioId());
            } else {
                cmd.setNull(1, Types.INTEGER);
            }

            cmd.setInt(2, c.getCliente().getIdCliente());

            if (c.getPaquete() != null && c.getPaquete().getId() > 0) {
                cmd.setInt(3, c.getPaquete().getId());
            } else {
                cmd.setNull(3, Types.INTEGER);
            }

            cmd.setTimestamp(4, c.getFecha() != null
                    ? Timestamp.valueOf(c.getFecha())
                    : new Timestamp(System.currentTimeMillis()));

            cmd.setString(5, c.getEstado() != null
                    ? c.getEstado().name().toLowerCase()
                    : EstadoCotizacion.BORRADOR.name().toLowerCase());

            cmd.setDouble(6, c.getConsumoPromedioMensualKwh());
            cmd.setDouble(7, c.getConsumoPromedioDiarioKwh());
            cmd.setDouble(8, c.getCostoPromedioMensual());
            cmd.setDouble(9, c.getCostoPromedioAnual());
            cmd.setDouble(10, c.getWattsInstalados());
            cmd.setDouble(11, c.getProduccionDiariaEstimada());
            cmd.setDouble(12, c.getPorcentajeCobertura());
            cmd.setDouble(13, c.getRetornoInversion());
            cmd.setDouble(14, c.getSubtotal());
            cmd.setDouble(15, c.getIva());
            cmd.setDouble(16, c.getTotal());
            cmd.setBoolean(17, c.isFinanciamiento());
            cmd.setString(18, c.getNotas());
            cmd.setInt(19, c.getCreatedBy() > 0 ? c.getCreatedBy() : 1);

            cmd.executeUpdate();

            try (ResultSet keys = cmd.getGeneratedKeys()) {
                if (keys.next()) {
                    c.setId(keys.getInt(1));
                }
            }

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al guardar cotizacion: " + ex.getMessage(), ex);
        }
    }

    @Override
    public List<Cotizacion> obtenerTodas() throws PersistenciaException {
        List<Cotizacion> lista = new ArrayList<>();
        String sql = SQL_SELECT + "ORDER BY co.fecha DESC";

        try (Connection conn = connectionBD.getConexion(); PreparedStatement cmd = conn.prepareStatement(sql); ResultSet rs = cmd.executeQuery()) {

            while (rs.next()) {
                lista.add(mapear(rs));
            }

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al listar cotizaciones: " + ex.getMessage(), ex);
        }
        return lista;
    }

    @Override
    public Cotizacion obtenerPorId(int id) throws PersistenciaException {
        String sql = SQL_SELECT + "AND co.id = ?";

        try (Connection conn = connectionBD.getConexion(); PreparedStatement cmd = conn.prepareStatement(sql)) {

            cmd.setInt(1, id);
            try (ResultSet rs = cmd.executeQuery()) {
                if (rs.next()) {
                    return mapear(rs);
                }
            }

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al buscar cotizacion id=" + id + ": " + ex.getMessage(), ex);
        }
        return null;
    }

    @Override
    public void actualizarCotizacion(Cotizacion c) throws PersistenciaException {
        String sql
                = "UPDATE cotizaciones SET "
                + "consumo_promedio_mensual_kwh=?, consumo_promedio_diario_kwh=?, "
                + "costo_promedio_mensual=?, costo_promedio_anual=?, "
                + "watts_instalados=?, produccion_diaria_estimada=?, "
                + "porcentaje_cobertura=?, retorno_inversion=?, "
                + "subtotal=?, iva=?, total=?, estado=?, notas=?, updated_by=? "
                + "WHERE id=? AND deleted_at IS NULL";

        try (Connection conn = connectionBD.getConexion(); PreparedStatement cmd = conn.prepareStatement(sql)) {

            cmd.setDouble(1, c.getConsumoPromedioMensualKwh());
            cmd.setDouble(2, c.getConsumoPromedioDiarioKwh());
            cmd.setDouble(3, c.getCostoPromedioMensual());
            cmd.setDouble(4, c.getCostoPromedioAnual());
            cmd.setDouble(5, c.getWattsInstalados());
            cmd.setDouble(6, c.getProduccionDiariaEstimada());
            cmd.setDouble(7, c.getPorcentajeCobertura());
            cmd.setDouble(8, c.getRetornoInversion());
            cmd.setDouble(9, c.getSubtotal());
            cmd.setDouble(10, c.getIva());
            cmd.setDouble(11, c.getTotal());
            cmd.setString(12, c.getEstado() != null
                    ? c.getEstado().name().toLowerCase()
                    : EstadoCotizacion.BORRADOR.name().toLowerCase());
            cmd.setString(13, c.getNotas());
            cmd.setInt(14, c.getUpdatedBy() > 0 ? c.getUpdatedBy() : 1);
            cmd.setInt(15, c.getId());

            int filas = cmd.executeUpdate();
            if (filas == 0) {
                throw new PersistenciaException("No existe cotizacion con id=" + c.getId());
            }

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al actualizar cotizacion: " + ex.getMessage(), ex);
        }
    }

    @Override
    public void eliminarCotizacion(int id) throws PersistenciaException {
        // Soft delete
        String sql = "UPDATE cotizaciones SET deleted_at = NOW() WHERE id = ? AND deleted_at IS NULL";

        try (Connection conn = connectionBD.getConexion(); PreparedStatement cmd = conn.prepareStatement(sql)) {

            cmd.setInt(1, id);
            int filas = cmd.executeUpdate();
            if (filas == 0) {
                throw new PersistenciaException("No existe cotizacion con id=" + id);
            }

        } catch (SQLException ex) {
            throw new PersistenciaException("Error al eliminar cotizacion: " + ex.getMessage(), ex);
        }
    }

    private Cotizacion mapear(ResultSet rs) throws SQLException {
        Cotizacion c = new Cotizacion();
        c.setId(rs.getInt("id"));
        c.setFecha(rs.getTimestamp("fecha").toLocalDateTime());

        String estadoStr = rs.getString("estado");
        if (estadoStr != null) {
            try {
                c.setEstado(EstadoCotizacion.valueOf(estadoStr.toUpperCase()));
            } catch (IllegalArgumentException e) {
                c.setEstado(EstadoCotizacion.BORRADOR);
            }
        }

        c.setConsumoPromedioMensualKwh(rs.getDouble("consumo_promedio_mensual_kwh"));
        c.setConsumoPromedioDiarioKwh(rs.getDouble("consumo_promedio_diario_kwh"));
        c.setCostoPromedioMensual(rs.getDouble("costo_promedio_mensual"));
        c.setCostoPromedioAnual(rs.getDouble("costo_promedio_anual"));
        c.setWattsInstalados(rs.getDouble("watts_instalados"));
        c.setProduccionDiariaEstimada(rs.getDouble("produccion_diaria_estimada"));
        c.setPorcentajeCobertura(rs.getDouble("porcentaje_cobertura"));
        c.setRetornoInversion(rs.getDouble("retorno_inversion"));
        c.setSubtotal(rs.getDouble("subtotal"));
        c.setIva(rs.getDouble("iva"));
        c.setTotal(rs.getDouble("total"));
        c.setFinanciamiento(rs.getBoolean("financiamiento"));
        c.setProyectoGenerado(rs.getBoolean("proyecto_generado"));
        c.setNotas(rs.getString("notas"));
        c.setCreatedBy(rs.getInt("created_by"));
        c.setUpdatedBy(rs.getInt("updated_by"));

        Cliente cliente = new Cliente();
        cliente.setIdCliente(rs.getInt("cliente_id"));
        cliente.setNombreComercial(rs.getString("nombre_comercial"));
        cliente.setRfc(rs.getString("rfc"));
        c.setCliente(cliente);

        int vendedorId = rs.getInt("vendedor_id");
        if (!rs.wasNull() && vendedorId > 0) {
            Vendedor v = new Vendedor(vendedorId, 0);
            v.setNombre(rs.getString("nombre_vendedor"));
            c.setVendedor(v);
        }

        return c;
    }
}
