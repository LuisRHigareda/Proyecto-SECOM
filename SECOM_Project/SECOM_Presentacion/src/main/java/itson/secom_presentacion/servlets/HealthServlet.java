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
import java.sql.Driver;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

@WebServlet(name = "HealthServlet", urlPatterns = {"/api/health"})
public class HealthServlet extends HttpServlet {

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

            if (input == null) {
                throw new RuntimeException("No se encontró config.properties en el classpath");
            }

            Properties props = new Properties();
            props.load(input);

            String usuario = props.getProperty("db.usuario");
            String contrasenia = props.getProperty("db.contrasenia");
            String ip = props.getProperty("db.ip");
            String puerto = props.getProperty("db.puerto");

            String url = String.format(
                    "jdbc:mysql://%s:%s/%s?useSSL=false&allowPublicKeyRetrieval=true&serverTimezone=UTC",
                    ip, puerto, "secom_pi"
            );

            debug.put("configEncontrado", true);
            debug.put("ip", ip);
            debug.put("puerto", puerto);
            debug.put("usuario", usuario);
            debug.put("url", url);

            try {
                Class.forName("com.mysql.cj.jdbc.Driver");
                debug.put("driverClassLoaded", true);
            } catch (ClassNotFoundException ex) {
                debug.put("driverClassLoaded", false);
                debug.put("driverError", ex.toString());

                body.put("ok", false);
                body.put("message", "El WAR no está cargando mysql-connector-j.");
                database.put("ok", false);
                database.put("error", "No se pudo cargar com.mysql.cj.jdbc.Driver");
                body.put("database", database);
                body.put("debug", debug);

                JsonResponse.send(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR, body);
                return;
            }

            List<String> drivers = new ArrayList<>();
            Enumeration<Driver> enumeration = DriverManager.getDrivers();
            while (enumeration.hasMoreElements()) {
                drivers.add(enumeration.nextElement().getClass().getName());
            }
            debug.put("driversRegistrados", drivers);

            try (Connection conn = DriverManager.getConnection(url, usuario, contrasenia);
                 PreparedStatement ps = conn.prepareStatement("SELECT DATABASE() AS nombre_bd, NOW() AS fecha_servidor");
                 ResultSet rs = ps.executeQuery()) {

                body.put("ok", true);
                body.put("message", "Presentación desplegada y conexión a base exitosa.");

                database.put("ok", true);

                if (rs.next()) {
                    database.put("nombre", rs.getString("nombre_bd"));
                    database.put("fechaServidor", String.valueOf(rs.getTimestamp("fecha_servidor")));
                }
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
}