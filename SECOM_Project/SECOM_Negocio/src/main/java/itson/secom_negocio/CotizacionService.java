package itson.secom_negocio;

import itson.secom_domain.DatosReciboCFE;
import itson.secom_domain.ParametrosSistema;
import itson.secom_domain.ProductoCantidad;
import itson.secom_domain.ResultadoCalculoCotizacion;
import itson.secom_persistence.conexionFactory.DAOFactory;
import itson.secom_persistence.excepciones.PersistenciaException;
import itson.secom_persistence.IParametrosSistemaDAO;
import itson.secom_persistence.IPaqueteCotizacionDAO;
import java.util.List;


public class CotizacionService {

    private final DAOFactory factory;
    private final MotorCotizacionSolar motor;

    public CotizacionService() {
        this.factory = new DAOFactory(false);
        this.motor = new MotorCotizacionSolar();
    }

    public CotizacionService(DAOFactory factory) {
        this.factory = factory;
        this.motor = new MotorCotizacionSolar();
    }

    public ResultadoCalculoCotizacion calcularCotizacionConPaquete(
            DatosReciboCFE datos,
            int paqueteId) throws Exception {

        try {
            IParametrosSistemaDAO daoParams = factory.conexionParametrosSistemaDAO();
            IPaqueteCotizacionDAO daoPaquete = factory.conexionPaqueteCotizacionDAO();

            ParametrosSistema params = daoParams.obtenerParametros(datos.getCiudad());
            List<ProductoCantidad> productos = daoPaquete.obtenerProductosPorPaquete(paqueteId);

            ResultadoCalculoCotizacion resultado =
                    motor.calcular(datos, params, productos);

            resultado.setProductos(productos);

            return resultado;

        } catch (PersistenciaException ex) {
            throw new Exception(
                    "Error al cargar datos de cotización: " + ex.getMessage(), ex);
        }
    }
}