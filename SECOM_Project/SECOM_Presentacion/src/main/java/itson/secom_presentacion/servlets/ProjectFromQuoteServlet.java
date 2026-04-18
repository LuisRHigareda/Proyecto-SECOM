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

@WebServlet(name = "ProjectFromQuoteServlet", urlPatterns = {"/api/projects/from-quote/*"})
public class ProjectFromQuoteServlet extends HttpServlet {

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws IOException {
        String id = getPathId(request);
        if (id == null) {
            JsonResponse.send(response, HttpServletResponse.SC_BAD_REQUEST,
                    Map.of("ok", false, "message", "Falta el ID de la cotización."));
            return;
        }

        try {
            Map<String, Object> saved = BackendStore.createProjectFromQuote(id);

            Map<String, Object> body = new LinkedHashMap<>();
            body.put("ok", true);
            body.put("data", saved);

            JsonResponse.send(response, HttpServletResponse.SC_CREATED, body);
        } catch (Exception ex) {
            JsonResponse.send(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR, Map.of(
                    "ok", false,
                    "message", ex.getMessage(),
                    "type", ex.getClass().getName()
            ));
        }
    }

    private String getPathId(HttpServletRequest request) {
        String path = request.getPathInfo();
        if (path == null || path.isBlank() || "/".equals(path)) {
            return null;
        }
        return path.substring(1);
    }
}