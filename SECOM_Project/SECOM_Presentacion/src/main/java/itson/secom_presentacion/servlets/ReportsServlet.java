package itson.secom_presentacion.servlets;

import itson.secom_presentacion.util.BackendStore;
import itson.secom_presentacion.util.JsonResponse;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

@WebServlet(name = "ReportsServlet", urlPatterns = {"/api/reports/quotes"})
public class ReportsServlet extends HttpServlet {

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws IOException {
        try {
            String fechaInicio = request.getParameter("fechaInicio");
            String fechaFin = request.getParameter("fechaFin");
            String status = request.getParameter("status");
            String tarifa = request.getParameter("tarifa");

            Map<String, Object> report = BackendStore.buildCotizacionesReport(fechaInicio, fechaFin, status, tarifa);
            Map<String, Object> body = new LinkedHashMap<>();
            body.put("ok", true);
            body.put("data", report);
            JsonResponse.send(response, HttpServletResponse.SC_OK, body);
        } catch (IllegalArgumentException ex) {
            JsonResponse.send(response, HttpServletResponse.SC_BAD_REQUEST, Map.of(
                    "ok", false,
                    "message", ex.getMessage(),
                    "type", ex.getClass().getName()
            ));
        } catch (Exception ex) {
            JsonResponse.send(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR, Map.of(
                    "ok", false,
                    "message", ex.getMessage(),
                    "type", ex.getClass().getName()
            ));
        }
    }
}
