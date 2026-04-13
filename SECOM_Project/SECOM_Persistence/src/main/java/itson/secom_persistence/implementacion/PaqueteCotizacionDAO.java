/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_persistence.implementacion;

/**
 *
 * @author Arell
 */

import itson.secom_domain.Producto;
import itson.secom_domain.ProductoCantidad;
import itson.secom_domain.enumeradores.CategoriaProducto;
import itson.secom_persistence.IConnectionBD;
import itson.secom_persistence.IPaqueteCotizacionDAO;
import itson.secom_persistence.excepciones.PersistenciaException;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;


public class PaqueteCotizacionDAO implements IPaqueteCotizacionDAO {

    private final IConnectionBD connectionBD;

    public PaqueteCotizacionDAO(IConnectionBD connectionBD) {
        this.connectionBD = connectionBD;
    }

    @Override
    public List<ProductoCantidad> obtenerProductosPorPaquete(int paqueteId)
            throws PersistenciaException {

        String sql = """
            SELECT 
                p.id,
                p.nombre,
                p.categoria,
                p.capacidad,
                p.precio_base,
                pp.cantidad_base
            FROM paquete_productos pp
            INNER JOIN productos p ON p.id = pp.producto_id
            WHERE pp.paquete_id = ? AND p.activo = TRUE
        """;

        List<ProductoCantidad> lista = new ArrayList<>();

        try (Connection conexion = connectionBD.getConexion();
             PreparedStatement ps = conexion.prepareStatement(sql)) {

            ps.setInt(1, paqueteId);

            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    Producto producto = new Producto();
                    producto.setId(rs.getInt("id"));
                    producto.setNombre(rs.getString("nombre"));
                    producto.setCategoria(
                            CategoriaProducto.valueOf(rs.getString("categoria").toUpperCase())
                    );
                    producto.setCapacidad((int) rs.getDouble("capacidad"));
                    producto.setPrecioUnitario(rs.getDouble("precio_base"));

                    ProductoCantidad pc = new ProductoCantidad();
                    pc.setProducto(producto);
                    pc.setCantidad(rs.getDouble("cantidad_base"));

                    lista.add(pc);
                }
            }

            if (lista.isEmpty()) {
                throw new PersistenciaException(
                        "El paquete con ID " + paqueteId + " no contiene productos.");
            }

            return lista;

        } catch (SQLException ex) {
            throw new PersistenciaException(
                    "Error al obtener productos del paquete: " + ex.getMessage(), ex);
        }
    }

    @Override
    public void close() {
        // No se requiere implementación.
    }
}