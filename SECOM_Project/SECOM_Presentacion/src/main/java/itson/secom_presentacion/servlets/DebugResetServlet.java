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

@WebServlet(name = "DebugResetServlet", urlPatterns = {"/api/debug/reset"})
public class DebugResetServlet extends HttpServlet {

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws IOException {
        try {
            int affected = BackendStore.resetAllData();
            Map<String, Object> body = new LinkedHashMap<>();
            body.put("ok", true);
            body.put("affectedRows", affected);
            JsonResponse.send(response, HttpServletResponse.SC_OK, body);
        } catch (Exception ex) {
            JsonResponse.send(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR, Map.of(
                    "ok", false,
                    "message", ex.getMessage(),
                    "type", ex.getClass().getName()
            ));
        }
    }
}
