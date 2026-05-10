package itson.secom_presentacion.servlets;

import itson.secom_presentacion.util.BackendStore;
import itson.secom_presentacion.util.JsonResponse;
import itson.secom_presentacion.util.RequestJson;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

@WebServlet(name = "InsumosServlet", urlPatterns = {"/api/insumos", "/api/insumos/*"})
public class InsumosServlet extends HttpServlet {

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws IOException {
        try {
            Map<String, Object> body = new LinkedHashMap<>();
            body.put("ok", true);
            body.put("data", BackendStore.listInsumos());
            JsonResponse.send(response, HttpServletResponse.SC_OK, body);
        } catch (Exception ex) {
            sendError(response, ex);
        }
    }

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws IOException {
        try {
            Map<String, Object> payload = RequestJson.readMap(request);
            Map<String, Object> saved = BackendStore.saveInsumo(payload);

            Map<String, Object> body = new LinkedHashMap<>();
            body.put("ok", true);
            body.put("data", saved);

            JsonResponse.send(response, HttpServletResponse.SC_CREATED, body);
        } catch (Exception ex) {
            sendError(response, ex);
        }
    }

    @Override
    protected void doPut(HttpServletRequest request, HttpServletResponse response) throws IOException {
        String id = getPathId(request);
        if (id == null) {
            JsonResponse.send(response, HttpServletResponse.SC_BAD_REQUEST,
                    Map.of("ok", false, "message", "Falta el ID del insumo."));
            return;
        }

        try {
            Map<String, Object> patch = RequestJson.readMap(request);
            Map<String, Object> saved = BackendStore.updateInsumo(id, patch);

            Map<String, Object> body = new LinkedHashMap<>();
            body.put("ok", true);
            body.put("data", saved);

            JsonResponse.send(response, HttpServletResponse.SC_OK, body);
        } catch (Exception ex) {
            sendError(response, ex);
        }
    }

    @Override
    protected void doDelete(HttpServletRequest request, HttpServletResponse response) throws IOException {
        String id = getPathId(request);
        if (id == null) {
            JsonResponse.send(response, HttpServletResponse.SC_BAD_REQUEST,
                    Map.of("ok", false, "message", "Falta el ID del insumo."));
            return;
        }

        try {
            BackendStore.deleteInsumo(id);
            JsonResponse.send(response, HttpServletResponse.SC_OK, Map.of("ok", true));
        } catch (Exception ex) {
            sendError(response, ex);
        }
    }

    private String getPathId(HttpServletRequest request) {
        String path = request.getPathInfo();
        if (path == null || path.isBlank() || "/".equals(path)) {
            return null;
        }
        return path.substring(1);
    }

    private void sendError(HttpServletResponse response, Exception ex) throws IOException {
        JsonResponse.send(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR, Map.of(
                "ok", false,
                "message", ex.getMessage(),
                "type", ex.getClass().getName()
        ));
    }
}
