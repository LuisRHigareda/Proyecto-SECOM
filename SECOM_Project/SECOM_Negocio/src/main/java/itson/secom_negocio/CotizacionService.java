/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_negocio;

import itson.secom_domain.CalculoSolar;
import itson.secom_domain.Cliente;
import itson.secom_domain.ConsumoMensual;
import itson.secom_domain.Cotizacion;
import itson.secom_domain.DatosReciboCFE;
import itson.secom_domain.ParametrosSistema;
import itson.secom_domain.ProductoCantidad;
import itson.secom_domain.ResultadoCalculoCotizacion;
import itson.secom_domain.Vendedor;
import itson.secom_domain.enumeradores.EstadoCotizacion;
import itson.secom_persistence.ICalculoSolarDAO;
import itson.secom_persistence.IConsumoMensualDAO;
import itson.secom_persistence.ICotizacionDAO;
import itson.secom_persistence.IPaqueteCotizacionDAO;
import itson.secom_persistence.IParametrosSistemaDAO;
import itson.secom_persistence.conexionFactory.DAOFactory;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;


/**
 *
 * @author Acer
 */
public class CotizacionService {
//    private static final double IVA            = 0.16;
//    private static final double COSTO_KWP_MXN  = 22000.0;
// 
//    // Solo el motor y el factory — NO se guardan los DAOs como campos
//    private final MotorCotizacionSolar motor;
//    private final DAOFactory factory;
// 
//    public CotizacionService() {
//        this.factory = new DAOFactory(false);
//        this.motor   = new MotorCotizacionSolar();
//    }
// 
//    // -------------------------------------------------------
//    // Calcular + guardar completo
//    // -------------------------------------------------------
//    public Cotizacion calcularYGuardar(DatosReciboCFE datos, Cliente cliente, int vendedorId)
//            throws Exception {
//        validarEntrada(datos, cliente);
// 
//        ResultadoCalculoCotizacion resultado = motor.calcular(datos);
//        Cotizacion cotizacion = construirCotizacion(datos, resultado, cliente, vendedorId);
// 
//        // 1. Guardar cotizacion (conexion propia)
//        try (ICotizacionDAO dao = factory.conexionCotizacionDAO()) {
//            dao.guardarCotizacion(cotizacion);
//        } catch (PersistenciaException ex) {
//            throw new Exception("Error al guardar cotizacion: " + ex.getMessage());
//        }
// 
//        // 2. Guardar calculo solar (conexion propia)
//        guardarCalculoSolar(resultado, datos, cotizacion);
// 
//        // 3. Guardar consumos historicos (conexion propia)
//        guardarConsumosHistoricos(datos, cotizacion);
// 
//        return cotizacion;
//    }
// 
//    // -------------------------------------------------------
//    // Solo calcular sin guardar (preview)
//    // -------------------------------------------------------
//    public ResultadoCalculoCotizacion calcularSinGuardar(DatosReciboCFE datos) throws Exception {
//        return motor.calcular(datos);
//    }
// 
//    // -------------------------------------------------------
//    // Listar todas
//    // -------------------------------------------------------
//    public List<Cotizacion> listarTodas() throws Exception {
//        try (ICotizacionDAO dao = factory.conexionCotizacionDAO()) {
//            return dao.obtenerTodas();
//        } catch (PersistenciaException ex) {
//            throw new Exception("Error al listar cotizaciones: " + ex.getMessage());
//        }
//    }
// 
//    // -------------------------------------------------------
//    // Obtener por ID
//    // -------------------------------------------------------
//    public Cotizacion obtenerPorId(int id) throws Exception {
//        if (id <= 0) throw new Exception("ID invalido.");
//        try (ICotizacionDAO dao = factory.conexionCotizacionDAO()) {
//            return dao.obtenerPorId(id);
//        } catch (PersistenciaException ex) {
//            throw new Exception("Error al obtener cotizacion: " + ex.getMessage());
//        }
//    }
// 
//    // -------------------------------------------------------
//    // Obtener calculo solar
//    // -------------------------------------------------------
//    public CalculoSolar obtenerCalculoSolar(int idCotizacion) throws Exception {
//        try (ICalculoSolarDAO dao = factory.conexionCalculoSolarDAO()) {
//            return dao.obtenerPorCotizacion(idCotizacion);
//        } catch (PersistenciaException ex) {
//            throw new Exception("Error al obtener calculo solar: " + ex.getMessage());
//        }
//    }
// 
//    // -------------------------------------------------------
//    // Obtener consumos historicos
//    // -------------------------------------------------------
//    public List<ConsumoMensual> obtenerConsumos(int idCotizacion) throws Exception {
//        try (IConsumoMensualDAO dao = factory.conexionConsumoMensualDAO()) {
//            return dao.obtenerPorCotizacion(idCotizacion);
//        } catch (PersistenciaException ex) {
//            throw new Exception("Error al obtener consumos: " + ex.getMessage());
//        }
//    }
// 
//    // -------------------------------------------------------
//    // Cambiar estado
//    // -------------------------------------------------------
//    public void cambiarEstado(int idCotizacion, EstadoCotizacion nuevoEstado, int usuarioId)
//            throws Exception {
//        Cotizacion c = obtenerPorId(idCotizacion);
//        if (c == null) throw new Exception("No existe cotizacion con id=" + idCotizacion);
//        c.setEstado(nuevoEstado);
//        c.setUpdatedBy(usuarioId);
//        try (ICotizacionDAO dao = factory.conexionCotizacionDAO()) {
//            dao.actualizarCotizacion(c);
//        } catch (PersistenciaException ex) {
//            throw new Exception("Error al cambiar estado: " + ex.getMessage());
//        }
//    }
// 
//    // -------------------------------------------------------
//    // Helpers privados
//    // -------------------------------------------------------
// 
//    private Cotizacion construirCotizacion(DatosReciboCFE datos,
//                                            ResultadoCalculoCotizacion r,
//                                            Cliente cliente, int vendedorId) {
//        Cotizacion c = new Cotizacion();
//        c.setCliente(cliente);
//        c.setFecha(LocalDateTime.now());
//        c.setEstado(EstadoCotizacion.BORRADOR);
//        c.setConsumoPromedioMensualKwh(r.getConsumoPromedioMensualKwh());
//        c.setConsumoPromedioDiarioKwh(r.getConsumoPromedioMensualKwh() / 30.0);
//        c.setCostoPromedioMensual(r.getPagoPromedioCFE());
//        c.setCostoPromedioAnual(r.getPagoPromedioCFE() * 12.0);
//        c.setWattsInstalados(r.getWattsInstalados());
//        c.setProduccionDiariaEstimada(r.getProduccionDiariaEstimada());
//        c.setPorcentajeCobertura(r.getPorcentajCobertura());
//        c.setRetornoInversion(r.getRetornoInversion());
// 
//        double subtotal = r.getPotenciaInstaladaKwp() * COSTO_KWP_MXN;
//        c.setSubtotal(subtotal);
//        c.setIva(subtotal * IVA);
//        c.setTotal(subtotal * (1 + IVA));
// 
//        if (vendedorId > 0) {
//            c.setVendedor(new Vendedor(vendedorId, 0));
//        }
//        c.setCreatedBy(vendedorId > 0 ? vendedorId : cliente.getId());
//        return c;
//    }
// 
//    private void guardarCalculoSolar(ResultadoCalculoCotizacion r,
//                                      DatosReciboCFE datos,
//                                      Cotizacion cotizacion) throws Exception {
//        CalculoSolar cs = new CalculoSolar();
//        cs.setCotizacion(cotizacion);
//        cs.setEstadoMX(datos.getNumeroEstado() > 0
//                ? String.valueOf(datos.getNumeroEstado()) : "SON");
//        cs.setInsolacionUsada(5.5);
//        cs.setPotencialPanel(550.0);
//        cs.setNumeroPaneles(r.getNumeroPaneles());
//        cs.setWattsInstalados(r.getWattsInstalados());
//        cs.setCapacidadInversor(r.getPotenciaInstaladaKwp() * 1.25);
//        cs.setProduccionDiariaEstimada(r.getProduccionDiariaEstimada());
//        cs.setProduccionAnualEstimada(r.getGeneracioAnualEstimadaKwh());
//        cs.setPorcentajeGeneracion(r.getPorcentajCobertura());
//        cs.setFactorConversionUsado(0.80);
//        cs.setFactorReflexionUsado(1.0);
//        cs.setFechaCalculo(LocalDateTime.now());
// 
//        try (ICalculoSolarDAO dao = factory.conexionCalculoSolarDAO()) {
//            dao.guardar(cs);
//        } catch (PersistenciaException ex) {
//            throw new Exception("Error al guardar calculo solar: " + ex.getMessage());
//        }
//    }
// 
//    private void guardarConsumosHistoricos(DatosReciboCFE datos,
//                                            Cotizacion cotizacion) throws Exception {
//        List<Double> consumos = datos.getConsumoHistoricos();
//        if (consumos == null || consumos.isEmpty()) return;
// 
//        int mesActual  = LocalDateTime.now().getMonthValue();
//        int anioActual = LocalDateTime.now().getYear();
//        int total      = consumos.size();
//        List<ConsumoMensual> lista = new ArrayList<>();
// 
//        for (int i = 0; i < total; i++) {
//            Double kwh = consumos.get(i);
//            if (kwh == null || kwh <= 0) continue;
//            int offset = total - 1 - i;
//            int mes    = mesActual - offset;
//            int anio   = anioActual;
//            while (mes <= 0) { mes += 12; anio--; }
//            lista.add(new ConsumoMensual(mes, anio, kwh, cotizacion));
//        }
// 
//        try (IConsumoMensualDAO dao = factory.conexionConsumoMensualDAO()) {
//            dao.guardarTodos(lista);
//        } catch (PersistenciaException ex) {
//            throw new Exception("Error al guardar consumos: " + ex.getMessage());
//        }
//    }
// 
//    private void validarEntrada(DatosReciboCFE datos, Cliente cliente) throws Exception {
//        if (datos == null)
//            throw new Exception("Los datos del recibo son obligatorios.");
//        if (cliente == null || cliente.getIdCliente() <= 0)
//            throw new Exception("Se necesita un cliente valido.");
//        if (datos.getTipoTarifa() == null)
//            throw new Exception("Se necesita especificar el tipo de tarifa.");
//        boolean tieneConsumo = datos.getConsumoActualKwh() > 0
//                || (datos.getConsumoHistoricos() != null && !datos.getConsumoHistoricos().isEmpty());
//        if (!tieneConsumo)
//            throw new Exception("Se necesita al menos un dato de consumo kWh.");
//    }
//}


    private final DAOFactory factory;
    private final MotorCotizacionSolar motor;


    public CotizacionService() {
        this.factory = new DAOFactory(false);
        this.motor = new MotorCotizacionSolar();
    }

public CotizacionService(DAOFactory factory) {
    this.factory = factory;
    this.motor = new MotorCotizacionSolar(); // 🔥 FALTABA
}
    public ResultadoCalculoCotizacion calcularCotizacionConPaquete(
        DatosReciboCFE datos,
        int paqueteId) throws Exception {

    ParametrosSistema params;
    List<ProductoCantidad> productos;

    try (IParametrosSistemaDAO daoParams = factory.conexionParametrosSistemaDAO();
         IPaqueteCotizacionDAO daoPaquete = factory.conexionPaqueteCotizacionDAO()) {

        params = daoParams.obtenerParametros(datos.getCiudad());
        productos = daoPaquete.obtenerProductosPorPaquete(paqueteId);

    } catch (PersistenciaException ex) {
        throw new Exception("Error al cargar datos de cotización: " + ex.getMessage(), ex);
    }

    // 🔥 CALCULAR
    ResultadoCalculoCotizacion resultado = motor.calcular(datos, params, productos);

    // 🔥 AQUÍ ESTÁ LA CLAVE
    resultado.setProductos(productos);

    return resultado;
}
}