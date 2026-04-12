package itson.secom_presentacion.servlets;

import itson.secom_domain.Cliente;
import itson.secom_negocio.CotizacionService;
import itson.secom_persistence.connectionDB.ConnectionDB;
import itson.secom_presentacion.util.JsonResponse;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

@WebServlet(name = "HealthServlet", urlPatterns = {"/api/health"})
public class HealthServlet extends HttpServlet {

    private static final String[] REQUIRED_TABLES = {
        "usuarios",
        "vendedores",
        "clientes",
        "clientes_telefonos",
        "cotizaciones",
        "cotizacion_detalles",
        "proyectos",
        "proyecto_materiales"
    };

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws IOException {

        Map<String, Object> body = new LinkedHashMap<>();
        Map<String, Object> modulos = new LinkedHashMap<>();
        Map<String, Object> database = new LinkedHashMap<>();
        Map<String, Object> debug = new LinkedHashMap<>();

        body.put("app", "SECOM_Presentacion");

        modulos.put("presentacion", true);
        modulos.put("dominio", Cliente.class != null);
        modulos.put("negocio", CotizacionService.class != null);
        modulos.put("persistencia", ConnectionDB.class != null);
        body.put("modulos", modulos);

        try (InputStream input = Thread.currentThread()
                .getContextClassLoader()
                .getResourceAsStream("config.properties")) {

            Properties props = new Properties();
            if (input != null) {
                props.load(input);
            }

            String ip = firstNonBlank(System.getenv("SECOM_DB_HOST"), props.getProperty("db.ip"), "127.0.0.1");
            String puerto = firstNonBlank(System.getenv("SECOM_DB_PORT"), props.getProperty("db.puerto"), "3306");
            String nombreDb = firstNonBlank(System.getenv("SECOM_DB_NAME"), props.getProperty("db.nombre"), "secom_pi");

            debug.put("configEncontrado", input != null);
            debug.put("host", ip);
            debug.put("puerto", puerto);
            debug.put("baseDatosEsperada", nombreDb);

            ConnectionDB db = new ConnectionDB(false);
            try (Connection conn = db.getConexion();
                 PreparedStatement ps = conn.prepareStatement("SELECT DATABASE() AS nombre_bd, NOW() AS fecha_servidor");
                 ResultSet rs = ps.executeQuery()) {

                body.put("ok", true);
                body.put("message", "Presentación desplegada y conexión a base exitosa.");
                database.put("ok", true);

                if (rs.next()) {
                    database.put("nombre", rs.getString("nombre_bd"));
                    database.put("fechaServidor", String.valueOf(rs.getTimestamp("fecha_servidor")));
                }

                List<String> missingTables = findMissingTables(conn);
                database.put("requiredTables", REQUIRED_TABLES.length);
                database.put("missingTables", missingTables);
                database.put("schemaReady", missingTables.isEmpty());

                if (!missingTables.isEmpty()) {
                    body.put("ok", false);
                    body.put("message", "La conexión existe, pero faltan tablas requeridas para operar.");
                    database.put("ok", false);
                }
            } finally {
                db.close();
            }

        } catch (Exception ex) {
            body.put("ok", false);
            body.put("message", "La aplicación sí desplegó, pero falló la conexión con la base de datos.");
            database.put("ok", false);
            database.put("error", ex.getMessage());
            database.put("tipo", ex.getClass().getName());

            if (ex.getCause() != null) {
                database.put("causa", ex.getCause().getMessage());
                database.put("causaTipo", ex.getCause().getClass().getName());
            }
        }

        body.put("database", database);
        body.put("debug", debug);

        int status = Boolean.TRUE.equals(body.get("ok"))
                ? HttpServletResponse.SC_OK
                : HttpServletResponse.SC_SERVICE_UNAVAILABLE;

        JsonResponse.send(response, status, body);
    }

    private static List<String> findMissingTables(Connection conn) throws Exception {
        List<String> missing = new ArrayList<>();
        DatabaseMetaData meta = conn.getMetaData();

        for (String table : REQUIRED_TABLES) {
            boolean exists = false;
            try (ResultSet rs = meta.getTables(conn.getCatalog(), null, table, new String[]{"TABLE"})) {
                exists = rs.next();
            }
            if (!exists) {
                missing.add(table);
            }
        }

        return missing;
    }

    private static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }

    private static String firstNonBlank(String... values) {
        for (String value : values) {
            if (!isBlank(value)) {
                return value.trim();
            }
        }
        return "";
    }
}
