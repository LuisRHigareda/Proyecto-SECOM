package itson.secom_presentacion.util;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;
import itson.secom_persistence.connectionDB.ConnectionDB;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public final class BackendStore {

    private static final Gson GSON = new GsonBuilder()
            .serializeNulls()
            .create();

    private static final String APP_SOURCE = "SECOM_UI_V1";

    private BackendStore() {
    }

    // =========================================================
    // QUOTES
    // =========================================================
    public static List<Map<String, Object>> listQuotes() throws Exception {
        ConnectionDB db = new ConnectionDB(false);
        try {
            Connection conn = db.getConexion();

            String sql = """
                SELECT q.id, q.fecha, q.estado, q.proyecto_generado, q.notas,
                       q.consumo_promedio_mensual_kwh, q.total,
                       q.created_at, q.updated_at,
                       c.nombre_comercial, c.ciudad, c.direccion_fiscal
                FROM cotizaciones q
                LEFT JOIN clientes c ON c.id = q.cliente_id
                WHERE q.deleted_at IS NULL
                ORDER BY q.fecha DESC, q.id DESC
            """;

            List<Map<String, Object>> out = new ArrayList<>();

            try (PreparedStatement ps = conn.prepareStatement(sql);
                 ResultSet rs = ps.executeQuery()) {

                while (rs.next()) {
                    out.add(mapQuoteRow(rs));
                }
            }

            return out;
        } finally {
            db.close();
        }
    }

    public static Map<String, Object> saveQuote(Map<String, Object> payload) throws Exception {
        Map<String, Object> state = normalizeState(payload);

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);

            Integer actorId = findFirstUserId(conn);
            Integer vendedorId = findFirstVendorId(conn);
            int clienteId = findOrCreateClient(conn, state);

            int quoteId = parseFrontId(asString(state.get("id")));
            boolean exists = quoteId > 0 && existsQuote(conn, quoteId);

            if (exists) {
                updateQuoteRow(conn, quoteId, state, clienteId, vendedorId, actorId);
            } else {
                quoteId = insertQuoteRow(conn, state, clienteId, vendedorId, actorId);
            }

            replaceQuoteDetails(conn, quoteId, state);
            conn.commit();

            return getQuoteById(conn, quoteId);
        } catch (Exception ex) {
            try {
                conn.rollback();
            } catch (SQLException ignore) {
            }
            throw ex;
        } finally {
            try {
                conn.setAutoCommit(true);
            } catch (SQLException ignore) {
            }
            db.close();
        }
    }

    public static Map<String, Object> updateQuote(String frontId, Map<String, Object> patch) throws Exception {
        int quoteId = parseFrontId(frontId);
        if (quoteId <= 0) {
            throw new IllegalArgumentException("ID de cotización inválido.");
        }

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);

            Map<String, Object> current = getStoredQuoteState(conn, quoteId);
            current.putAll(patch);
            current = normalizeState(current);

            Integer actorId = findFirstUserId(conn);
            Integer vendedorId = findFirstVendorId(conn);
            int clienteId = findOrCreateClient(conn, current);

            updateQuoteRow(conn, quoteId, current, clienteId, vendedorId, actorId);
            replaceQuoteDetails(conn, quoteId, current);

            conn.commit();
            return getQuoteById(conn, quoteId);
        } catch (Exception ex) {
            try {
                conn.rollback();
            } catch (SQLException ignore) {
            }
            throw ex;
        } finally {
            try {
                conn.setAutoCommit(true);
            } catch (SQLException ignore) {
            }
            db.close();
        }
    }

    // =========================================================
    // PROJECTS
    // =========================================================
    public static List<Map<String, Object>> listProjects() throws Exception {
        ConnectionDB db = new ConnectionDB(false);
        try {
            Connection conn = db.getConexion();

            String sql = """
                SELECT p.id, p.cotizacion_id, p.estado, p.notas,
                       p.total_vendido, p.potencia_total_instalada,
                       p.created_at, p.updated_at,
                       c.nombre_comercial, p.direccion_instalacion, p.ciudad, p.rpui
                FROM proyectos p
                LEFT JOIN clientes c ON c.id = p.cliente_id
                WHERE p.deleted_at IS NULL
                ORDER BY p.created_at DESC, p.id DESC
            """;

            List<Map<String, Object>> out = new ArrayList<>();

            try (PreparedStatement ps = conn.prepareStatement(sql);
                 ResultSet rs = ps.executeQuery()) {

                while (rs.next()) {
                    out.add(mapProjectRow(rs));
                }
            }

            return out;
        } finally {
            db.close();
        }
    }

    public static Map<String, Object> saveProject(Map<String, Object> payload) throws Exception {
        Map<String, Object> state = normalizeState(payload);

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);

            Integer actorId = findFirstUserId(conn);
            Integer vendedorId = findFirstVendorId(conn);
            int clienteId = findOrCreateClient(conn, state);

            int projectId = parseFrontId(asString(state.get("id")));
            boolean exists = projectId > 0 && existsProject(conn, projectId);

            if (exists) {
                updateProjectRow(conn, projectId, state, clienteId, vendedorId, actorId);
            } else {
                projectId = insertProjectRow(conn, state, clienteId, vendedorId, actorId, null);
            }

            conn.commit();
            return getProjectById(conn, projectId);
        } catch (Exception ex) {
            try {
                conn.rollback();
            } catch (SQLException ignore) {
            }
            throw ex;
        } finally {
            try {
                conn.setAutoCommit(true);
            } catch (SQLException ignore) {
            }
            db.close();
        }
    }

    public static Map<String, Object> updateProject(String frontId, Map<String, Object> patch) throws Exception {
        int projectId = parseFrontId(frontId);
        if (projectId <= 0) {
            throw new IllegalArgumentException("ID de proyecto inválido.");
        }

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);

            Map<String, Object> current = getStoredProjectState(conn, projectId);
            current.putAll(patch);
            current = normalizeState(current);

            Integer actorId = findFirstUserId(conn);
            Integer vendedorId = findFirstVendorId(conn);
            int clienteId = findOrCreateClient(conn, current);

            updateProjectRow(conn, projectId, current, clienteId, vendedorId, actorId);
            conn.commit();

            return getProjectById(conn, projectId);
        } catch (Exception ex) {
            try {
                conn.rollback();
            } catch (SQLException ignore) {
            }
            throw ex;
        } finally {
            try {
                conn.setAutoCommit(true);
            } catch (SQLException ignore) {
            }
            db.close();
        }
    }

    public static void deleteProject(String frontId) throws Exception {
        int projectId = parseFrontId(frontId);
        if (projectId <= 0) {
            throw new IllegalArgumentException("ID de proyecto inválido.");
        }

        ConnectionDB db = new ConnectionDB(false);
        try {
            Connection conn = db.getConexion();

            String sql = "UPDATE proyectos SET deleted_at = NOW() WHERE id = ? AND deleted_at IS NULL";
            try (PreparedStatement ps = conn.prepareStatement(sql)) {
                ps.setInt(1, projectId);
                ps.executeUpdate();
            }
        } finally {
            db.close();
        }
    }

    public static Map<String, Object> createProjectFromQuote(String quoteFrontId) throws Exception {
        int quoteId = parseFrontId(quoteFrontId);
        if (quoteId <= 0) {
            throw new IllegalArgumentException("ID de cotización inválido.");
        }

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);

            Integer actorId = findFirstUserId(conn);

            Integer existingProjectId = findProjectIdByQuote(conn, quoteId);
            if (existingProjectId != null) {
                conn.commit();
                return getProjectById(conn, existingProjectId);
            }

            Map<String, Object> quoteState = getStoredQuoteState(conn, quoteId);
            if (quoteState.isEmpty()) {
                throw new IllegalStateException("No se encontró la cotización solicitada.");
            }

            Integer clienteId = getQuoteClientId(conn, quoteId);
            Integer vendedorId = getQuoteVendorId(conn, quoteId);

            int projectId = insertProjectRow(conn, quoteState, clienteId, vendedorId, actorId, quoteId);

            String copyMaterials = """
                INSERT INTO proyecto_materiales (proyecto_id, producto_id, cantidad, precio_unitario, notas)
                SELECT ?, producto_id, cantidad, precio_unitario, NULL
                FROM cotizacion_detalles
                WHERE cotizacion_id = ?
            """;
            try (PreparedStatement ps = conn.prepareStatement(copyMaterials)) {
                ps.setInt(1, projectId);
                ps.setInt(2, quoteId);
                ps.executeUpdate();
            }

            String updateQuote = """
                UPDATE cotizaciones
                SET proyecto_generado = TRUE,
                    estado = 'FINALIZADA',
                    updated_by = ?
                WHERE id = ?
            """;
            try (PreparedStatement ps = conn.prepareStatement(updateQuote)) {
                setNullableInt(ps, 1, actorId);
                ps.setInt(2, quoteId);
                ps.executeUpdate();
            }

            conn.commit();
            return getProjectById(conn, projectId);
        } catch (Exception ex) {
            try {
                conn.rollback();
            } catch (SQLException ignore) {
            }
            throw ex;
        } finally {
            try {
                conn.setAutoCommit(true);
            } catch (SQLException ignore) {
            }
            db.close();
        }
    }

    public static void resetAppData() throws Exception {
        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);

            String marker = "%" + APP_SOURCE + "%";

            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE proyectos SET deleted_at = NOW() WHERE deleted_at IS NULL AND notas LIKE ?")) {
                ps.setString(1, marker);
                ps.executeUpdate();
            }

            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE cotizaciones SET deleted_at = NOW() WHERE deleted_at IS NULL AND notas LIKE ?")) {
                ps.setString(1, marker);
                ps.executeUpdate();
            }

            conn.commit();
        } catch (Exception ex) {
            try {
                conn.rollback();
            } catch (SQLException ignore) {
            }
            throw ex;
        } finally {
            try {
                conn.setAutoCommit(true);
            } catch (SQLException ignore) {
            }
            db.close();
        }
    }

    // =========================================================
    // INTERNALS: QUOTES
    // =========================================================
    private static int insertQuoteRow(Connection conn, Map<String, Object> state,
            int clienteId, Integer vendedorId, Integer actorId) throws Exception {

        Map<String, Object> receipt = asMap(state.get("receipt"));
        Map<String, Object> quote = asMap(state.get("quote"));

        double consumoMensual = positive(asDouble(quote.get("consumoMensual")));
        double consumoDiario = consumoMensual > 0 ? consumoMensual / 30.0 : 0.0;
        double costoMensual = firstPositive(asDouble(quote.get("pagoProm")), asDouble(receipt.get("totalAPagar")));
        double costoAnual = costoMensual * 12.0;
        double wattsInstalados = positive(asDouble(quote.get("kwp"))) * 1000.0;
        double produccionMensual = positive(asDouble(quote.get("produccionMensual")));
        double produccionDiaria = produccionMensual > 0 ? produccionMensual / 30.0 : 0.0;
        double cobertura = consumoMensual > 0 ? (produccionMensual / consumoMensual) * 100.0 : 0.0;
        double retorno = positive(asDouble(quote.get("retornoAnios")));

        double subtotal = positive(asDouble(quote.get("subtotalInsumos")));
        double iva = positive(asDouble(quote.get("impuestosInsumos")));
        double total = positive(asDouble(quote.get("totalInsumos")));

        if (total <= 0) {
            total = positive(asDouble(quote.get("inversion")));
            subtotal = total;
            iva = 0;
        }

        String sql = """
            INSERT INTO cotizaciones
            (vendedor_id, cliente_id, paquete_id, fecha, estado,
             consumo_promedio_mensual_kwh, consumo_promedio_diario_kwh,
             costo_promedio_mensual, costo_promedio_anual,
             watts_instalados, produccion_diaria_estimada,
             porcentaje_cobertura, retorno_inversion,
             subtotal, iva, total, financiamiento, proyecto_generado,
             notas, created_by, updated_by)
            VALUES (?, ?, NULL, NOW(), ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            setNullableInt(ps, 1, vendedorId);
            ps.setInt(2, clienteId);
            ps.setString(3, mapUiQuoteStatusToDb(asString(state.get("status")), false));
            ps.setDouble(4, consumoMensual);
            ps.setDouble(5, consumoDiario);
            ps.setDouble(6, costoMensual);
            ps.setDouble(7, costoAnual);
            ps.setDouble(8, wattsInstalados);
            ps.setDouble(9, produccionDiaria);
            ps.setDouble(10, cobertura);
            ps.setDouble(11, retorno);
            ps.setDouble(12, subtotal);
            ps.setDouble(13, iva);
            ps.setDouble(14, total);
            ps.setBoolean(15, false);
            ps.setBoolean(16, false);
            ps.setString(17, RequestJson.toJson(state));
            setNullableInt(ps, 18, actorId);
            setNullableInt(ps, 19, actorId);

            ps.executeUpdate();

            try (ResultSet rs = ps.getGeneratedKeys()) {
                if (rs.next()) {
                    return rs.getInt(1);
                }
            }
        }

        throw new IllegalStateException("No se pudo generar el ID de la cotización.");
    }

    private static void updateQuoteRow(Connection conn, int quoteId, Map<String, Object> state,
            int clienteId, Integer vendedorId, Integer actorId) throws Exception {

        Map<String, Object> receipt = asMap(state.get("receipt"));
        Map<String, Object> quote = asMap(state.get("quote"));

        boolean proyectoGenerado = isQuoteProjectGenerated(conn, quoteId);

        double consumoMensual = positive(asDouble(quote.get("consumoMensual")));
        double consumoDiario = consumoMensual > 0 ? consumoMensual / 30.0 : 0.0;
        double costoMensual = firstPositive(asDouble(quote.get("pagoProm")), asDouble(receipt.get("totalAPagar")));
        double costoAnual = costoMensual * 12.0;
        double wattsInstalados = positive(asDouble(quote.get("kwp"))) * 1000.0;
        double produccionMensual = positive(asDouble(quote.get("produccionMensual")));
        double produccionDiaria = produccionMensual > 0 ? produccionMensual / 30.0 : 0.0;
        double cobertura = consumoMensual > 0 ? (produccionMensual / consumoMensual) * 100.0 : 0.0;
        double retorno = positive(asDouble(quote.get("retornoAnios")));

        double subtotal = positive(asDouble(quote.get("subtotalInsumos")));
        double iva = positive(asDouble(quote.get("impuestosInsumos")));
        double total = positive(asDouble(quote.get("totalInsumos")));

        if (total <= 0) {
            total = positive(asDouble(quote.get("inversion")));
            subtotal = total;
            iva = 0;
        }

        String sql = """
            UPDATE cotizaciones
            SET vendedor_id = ?,
                cliente_id = ?,
                estado = ?,
                consumo_promedio_mensual_kwh = ?,
                consumo_promedio_diario_kwh = ?,
                costo_promedio_mensual = ?,
                costo_promedio_anual = ?,
                watts_instalados = ?,
                produccion_diaria_estimada = ?,
                porcentaje_cobertura = ?,
                retorno_inversion = ?,
                subtotal = ?,
                iva = ?,
                total = ?,
                notas = ?,
                updated_by = ?
            WHERE id = ? AND deleted_at IS NULL
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            setNullableInt(ps, 1, vendedorId);
            ps.setInt(2, clienteId);
            ps.setString(3, mapUiQuoteStatusToDb(asString(state.get("status")), proyectoGenerado));
            ps.setDouble(4, consumoMensual);
            ps.setDouble(5, consumoDiario);
            ps.setDouble(6, costoMensual);
            ps.setDouble(7, costoAnual);
            ps.setDouble(8, wattsInstalados);
            ps.setDouble(9, produccionDiaria);
            ps.setDouble(10, cobertura);
            ps.setDouble(11, retorno);
            ps.setDouble(12, subtotal);
            ps.setDouble(13, iva);
            ps.setDouble(14, total);
            ps.setString(15, RequestJson.toJson(state));
            setNullableInt(ps, 16, actorId);
            ps.setInt(17, quoteId);
            ps.executeUpdate();
        }
    }

    private static void replaceQuoteDetails(Connection conn, int quoteId, Map<String, Object> state) throws Exception {
        try (PreparedStatement del = conn.prepareStatement("DELETE FROM cotizacion_detalles WHERE cotizacion_id = ?")) {
            del.setInt(1, quoteId);
            del.executeUpdate();
        }

        Map<String, Object> receipt = asMap(state.get("receipt"));
        List<Map<String, Object>> insumos = asListOfMaps(receipt.get("insumos"));
        if (insumos.isEmpty()) {
            return;
        }

        Map<String, Object> quote = asMap(state.get("quote"));
        int paneles = (int) Math.round(positive(asDouble(quote.get("paneles"))));

        String sql = """
            INSERT INTO cotizacion_detalles
            (cotizacion_id, producto_id, cantidad, precio_unitario, numero_paneles, numero_inversores)
            VALUES (?, NULL, ?, ?, ?, 0)
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            for (Map<String, Object> item : insumos) {
                double cantidad = positive(asDouble(item.get("cantidad")));
                double precio = positive(asDouble(item.get("precio")));

                if (cantidad <= 0 && precio <= 0) {
                    continue;
                }

                String descripcion = asString(item.get("descripcion")).toLowerCase();
                int numeroPaneles = (descripcion.contains("panel") || descripcion.contains("módulo") || descripcion.contains("modulo"))
                        ? paneles : 0;

                ps.setInt(1, quoteId);
                ps.setDouble(2, cantidad);
                ps.setDouble(3, precio);
                ps.setInt(4, numeroPaneles);
                ps.addBatch();
            }
            ps.executeBatch();
        }
    }

    private static Map<String, Object> getQuoteById(Connection conn, int quoteId) throws Exception {
        String sql = """
            SELECT q.id, q.fecha, q.estado, q.proyecto_generado, q.notas,
                   q.consumo_promedio_mensual_kwh, q.total,
                   q.created_at, q.updated_at,
                   c.nombre_comercial, c.ciudad, c.direccion_fiscal
            FROM cotizaciones q
            LEFT JOIN clientes c ON c.id = q.cliente_id
            WHERE q.id = ? AND q.deleted_at IS NULL
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setInt(1, quoteId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return mapQuoteRow(rs);
                }
            }
        }

        throw new IllegalStateException("No se encontró la cotización guardada.");
    }

    private static Map<String, Object> getStoredQuoteState(Connection conn, int quoteId) throws Exception {
        String sql = "SELECT notas FROM cotizaciones WHERE id = ? AND deleted_at IS NULL";
        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setInt(1, quoteId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return normalizeState(parseJsonObject(rs.getString("notas")));
                }
            }
        }
        return new LinkedHashMap<>();
    }

    private static Map<String, Object> mapQuoteRow(ResultSet rs) throws Exception {
        Map<String, Object> state = normalizeState(parseJsonObject(rs.getString("notas")));

        state.put("id", formatQuoteId(rs.getInt("id")));
        state.put("dbId", rs.getInt("id"));
        state.put("createdAt", toMillis(rs.getTimestamp("created_at")));
        state.put("updatedAt", toMillis(rs.getTimestamp("updated_at")));
        state.put("status", mapDbQuoteStatusToUi(rs.getString("estado"), rs.getBoolean("proyecto_generado")));

        Map<String, Object> client = ensureMap(state, "client");
        if (isBlank(asString(client.get("nombre")))) {
            client.put("nombre", asString(rs.getString("nombre_comercial")));
        }
        if (isBlank(asString(client.get("direccion")))) {
            client.put("direccion", asString(rs.getString("direccion_fiscal")));
        }

        Map<String, Object> receipt = ensureMap(state, "receipt");
        if (isBlank(asString(receipt.get("nombre")))) {
            receipt.put("nombre", asString(rs.getString("nombre_comercial")));
        }
        if (isBlank(asString(receipt.get("direccion")))) {
            receipt.put("direccion", asString(rs.getString("direccion_fiscal")));
        }

        Map<String, Object> quote = ensureMap(state, "quote");
        if (positive(asDouble(quote.get("consumoMensual"))) <= 0) {
            quote.put("consumoMensual", rs.getDouble("consumo_promedio_mensual_kwh"));
        }
        if (positive(asDouble(quote.get("inversion"))) <= 0) {
            quote.put("inversion", rs.getDouble("total"));
        }

        return state;
    }

    // =========================================================
    // INTERNALS: PROJECTS
    // =========================================================
    private static int insertProjectRow(Connection conn, Map<String, Object> state,
            Integer clienteId, Integer vendedorId, Integer actorId, Integer linkedQuoteId) throws Exception {

        Map<String, Object> client = asMap(state.get("client"));
        Map<String, Object> receipt = asMap(state.get("receipt"));
        Map<String, Object> quote = asMap(state.get("quote"));

        String direccion = firstNonBlank(
                asString(client.get("direccion")),
                asString(receipt.get("direccion"))
        );

        String ciudad = firstNonBlank(
                asString(receipt.get("estado")),
                asString(receipt.get("ciudad")),
                asString(client.get("ciudad"))
        );

        String rpui = asString(receipt.get("servicio"));
        String titular = firstNonBlank(asString(receipt.get("nombre")), asString(client.get("nombre")));
        String tarifa = mapTarifaEnum(firstNonBlank(asString(receipt.get("tarifa")),
                asString(asMap(state.get("selectedTariff")).get("label"))));

        double totalVendido = firstPositive(asDouble(quote.get("inversion")), asDouble(quote.get("totalInsumos")));
        double potencia = positive(asDouble(quote.get("kwp")));

        String sql = """
            INSERT INTO proyectos
            (cotizacion_id, cliente_id, vendedor_id, tecnico_id,
             direccion_instalacion, ciudad, rpui, titular_servicio, tarifa_electrica,
             produccion_diaria_real, total_vendido, potencia_total_instalada,
             estado, notas, financiamiento, created_by, updated_by)
            VALUES (?, ?, ?, NULL, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?, FALSE, ?, ?)
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            setNullableInt(ps, 1, linkedQuoteId);
            setNullableInt(ps, 2, clienteId);
            setNullableInt(ps, 3, vendedorId);
            ps.setString(4, direccion);
            ps.setString(5, ciudad);
            ps.setString(6, rpui);
            ps.setString(7, titular);
            if (isBlank(tarifa)) {
                ps.setNull(8, java.sql.Types.VARCHAR);
            } else {
                ps.setString(8, tarifa);
            }
            ps.setDouble(9, totalVendido);
            ps.setDouble(10, potencia);
            ps.setString(11, mapUiProjectStatusToDb(asString(state.get("status"))));
            ps.setString(12, RequestJson.toJson(state));
            setNullableInt(ps, 13, actorId);
            setNullableInt(ps, 14, actorId);

            ps.executeUpdate();

            try (ResultSet rs = ps.getGeneratedKeys()) {
                if (rs.next()) {
                    return rs.getInt(1);
                }
            }
        }

        throw new IllegalStateException("No se pudo generar el ID del proyecto.");
    }

    private static void updateProjectRow(Connection conn, int projectId, Map<String, Object> state,
            Integer clienteId, Integer vendedorId, Integer actorId) throws Exception {

        Map<String, Object> client = asMap(state.get("client"));
        Map<String, Object> receipt = asMap(state.get("receipt"));
        Map<String, Object> quote = asMap(state.get("quote"));

        String direccion = firstNonBlank(
                asString(client.get("direccion")),
                asString(receipt.get("direccion"))
        );

        String ciudad = firstNonBlank(
                asString(receipt.get("estado")),
                asString(receipt.get("ciudad")),
                asString(client.get("ciudad"))
        );

        String rpui = asString(receipt.get("servicio"));
        String titular = firstNonBlank(asString(receipt.get("nombre")), asString(client.get("nombre")));
        String tarifa = mapTarifaEnum(firstNonBlank(asString(receipt.get("tarifa")),
                asString(asMap(state.get("selectedTariff")).get("label"))));

        double totalVendido = firstPositive(asDouble(quote.get("inversion")), asDouble(quote.get("totalInsumos")));
        double potencia = positive(asDouble(quote.get("kwp")));

        String sql = """
            UPDATE proyectos
            SET cliente_id = ?,
                vendedor_id = ?,
                direccion_instalacion = ?,
                ciudad = ?,
                rpui = ?,
                titular_servicio = ?,
                tarifa_electrica = ?,
                total_vendido = ?,
                potencia_total_instalada = ?,
                estado = ?,
                notas = ?,
                updated_by = ?
            WHERE id = ? AND deleted_at IS NULL
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            setNullableInt(ps, 1, clienteId);
            setNullableInt(ps, 2, vendedorId);
            ps.setString(3, direccion);
            ps.setString(4, ciudad);
            ps.setString(5, rpui);
            ps.setString(6, titular);
            if (isBlank(tarifa)) {
                ps.setNull(7, java.sql.Types.VARCHAR);
            } else {
                ps.setString(7, tarifa);
            }
            ps.setDouble(8, totalVendido);
            ps.setDouble(9, potencia);
            ps.setString(10, mapUiProjectStatusToDb(asString(state.get("status"))));
            ps.setString(11, RequestJson.toJson(state));
            setNullableInt(ps, 12, actorId);
            ps.setInt(13, projectId);
            ps.executeUpdate();
        }
    }

    private static Map<String, Object> getProjectById(Connection conn, int projectId) throws Exception {
        String sql = """
            SELECT p.id, p.cotizacion_id, p.estado, p.notas,
                   p.total_vendido, p.potencia_total_instalada,
                   p.created_at, p.updated_at,
                   c.nombre_comercial, p.direccion_instalacion, p.ciudad, p.rpui
            FROM proyectos p
            LEFT JOIN clientes c ON c.id = p.cliente_id
            WHERE p.id = ? AND p.deleted_at IS NULL
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setInt(1, projectId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return mapProjectRow(rs);
                }
            }
        }

        throw new IllegalStateException("No se encontró el proyecto guardado.");
    }

    private static Map<String, Object> getStoredProjectState(Connection conn, int projectId) throws Exception {
        String sql = "SELECT notas FROM proyectos WHERE id = ? AND deleted_at IS NULL";
        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setInt(1, projectId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return normalizeState(parseJsonObject(rs.getString("notas")));
                }
            }
        }
        return new LinkedHashMap<>();
    }

    private static Map<String, Object> mapProjectRow(ResultSet rs) throws Exception {
        Map<String, Object> state = normalizeState(parseJsonObject(rs.getString("notas")));

        state.put("id", formatProjectId(rs.getInt("id")));
        state.put("dbId", rs.getInt("id"));
        state.put("createdAt", toMillis(rs.getTimestamp("created_at")));
        state.put("updatedAt", toMillis(rs.getTimestamp("updated_at")));

        if (isBlank(asString(state.get("status")))) {
            state.put("status", mapDbProjectStatusToUi(rs.getString("estado")));
        }

        if (rs.getObject("cotizacion_id") != null && isBlank(asString(state.get("quoteId")))) {
            state.put("quoteId", formatQuoteId(rs.getInt("cotizacion_id")));
        }

        Map<String, Object> client = ensureMap(state, "client");
        if (isBlank(asString(client.get("nombre")))) {
            client.put("nombre", asString(rs.getString("nombre_comercial")));
        }
        if (isBlank(asString(client.get("direccion")))) {
            client.put("direccion", asString(rs.getString("direccion_instalacion")));
        }

        Map<String, Object> receipt = ensureMap(state, "receipt");
        if (isBlank(asString(receipt.get("servicio")))) {
            receipt.put("servicio", asString(rs.getString("rpui")));
        }

        Map<String, Object> quote = ensureMap(state, "quote");
        if (positive(asDouble(quote.get("kwp"))) <= 0) {
            quote.put("kwp", rs.getDouble("potencia_total_instalada"));
        }
        if (positive(asDouble(quote.get("inversion"))) <= 0) {
            quote.put("inversion", rs.getDouble("total_vendido"));
        }

        return state;
    }

    // =========================================================
    // INTERNALS: CLIENTS + LOOKUPS
    // =========================================================
    private static int findOrCreateClient(Connection conn, Map<String, Object> state) throws Exception {
        Map<String, Object> client = asMap(state.get("client"));
        Map<String, Object> receipt = asMap(state.get("receipt"));

        String nombre = firstNonBlank(
                asString(client.get("nombre")),
                asString(receipt.get("nombre")),
                "Cliente sin nombre"
        );
        String direccion = firstNonBlank(
                asString(client.get("direccion")),
                asString(receipt.get("direccion"))
        );
        String ciudad = firstNonBlank(
                asString(receipt.get("estado")),
                asString(receipt.get("ciudad")),
                asString(client.get("ciudad"))
        );
        String rfc = asString(client.get("rfc"));
        String telefono = asString(client.get("telefono"));

        if (!isBlank(rfc)) {
            String byRfc = "SELECT id FROM clientes WHERE rfc = ? AND deleted_at IS NULL LIMIT 1";
            try (PreparedStatement ps = conn.prepareStatement(byRfc)) {
                ps.setString(1, rfc);
                try (ResultSet rs = ps.executeQuery()) {
                    if (rs.next()) {
                        int id = rs.getInt(1);
                        upsertClientPhone(conn, id, telefono);
                        return id;
                    }
                }
            }
        }

        String byName = """
            SELECT id
            FROM clientes
            WHERE nombre_comercial = ?
              AND deleted_at IS NULL
            ORDER BY id DESC
            LIMIT 1
        """;
        try (PreparedStatement ps = conn.prepareStatement(byName)) {
            ps.setString(1, nombre);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    int id = rs.getInt(1);
                    upsertClientPhone(conn, id, telefono);
                    return id;
                }
            }
        }

        String insert = """
            INSERT INTO clientes
            (usuario_id, rfc, razon_social, nombre_comercial, regimen_fiscal, direccion_fiscal, ciudad, activo)
            VALUES (NULL, ?, ?, ?, NULL, ?, ?, TRUE)
        """;

        try (PreparedStatement ps = conn.prepareStatement(insert, Statement.RETURN_GENERATED_KEYS)) {
            if (isBlank(rfc)) {
                ps.setNull(1, java.sql.Types.VARCHAR);
            } else {
                ps.setString(1, rfc);
            }
            ps.setString(2, nombre);
            ps.setString(3, nombre);
            ps.setString(4, direccion);
            ps.setString(5, ciudad);
            ps.executeUpdate();

            try (ResultSet rs = ps.getGeneratedKeys()) {
                if (rs.next()) {
                    int id = rs.getInt(1);
                    upsertClientPhone(conn, id, telefono);
                    return id;
                }
            }
        }

        throw new IllegalStateException("No se pudo crear el cliente.");
    }

    private static void upsertClientPhone(Connection conn, int clientId, String telefono) throws Exception {
        if (isBlank(telefono)) {
            return;
        }

        String check = """
            SELECT id
            FROM clientes_telefonos
            WHERE cliente_id = ? AND principal = TRUE
            LIMIT 1
        """;
        Integer phoneId = null;

        try (PreparedStatement ps = conn.prepareStatement(check)) {
            ps.setInt(1, clientId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    phoneId = rs.getInt(1);
                }
            }
        }

        if (phoneId == null) {
            String insert = """
                INSERT INTO clientes_telefonos (cliente_id, telefono, tipo, principal)
                VALUES (?, ?, 'CELULAR', TRUE)
            """;
            try (PreparedStatement ps = conn.prepareStatement(insert)) {
                ps.setInt(1, clientId);
                ps.setString(2, telefono);
                ps.executeUpdate();
            }
        } else {
            String update = "UPDATE clientes_telefonos SET telefono = ? WHERE id = ?";
            try (PreparedStatement ps = conn.prepareStatement(update)) {
                ps.setString(1, telefono);
                ps.setInt(2, phoneId);
                ps.executeUpdate();
            }
        }
    }

    private static Integer findFirstUserId(Connection conn) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement("SELECT id FROM usuarios ORDER BY id ASC LIMIT 1");
             ResultSet rs = ps.executeQuery()) {
            if (rs.next()) {
                return rs.getInt(1);
            }
        }
        return null;
    }

    private static Integer findFirstVendorId(Connection conn) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement("SELECT usuario_id FROM vendedores ORDER BY usuario_id ASC LIMIT 1");
             ResultSet rs = ps.executeQuery()) {
            if (rs.next()) {
                return rs.getInt(1);
            }
        }
        return null;
    }

    private static Integer findProjectIdByQuote(Connection conn, int quoteId) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT id FROM proyectos WHERE cotizacion_id = ? AND deleted_at IS NULL LIMIT 1")) {
            ps.setInt(1, quoteId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return rs.getInt(1);
                }
            }
        }
        return null;
    }

    private static Integer getQuoteClientId(Connection conn, int quoteId) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement("SELECT cliente_id FROM cotizaciones WHERE id = ?")) {
            ps.setInt(1, quoteId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    Object v = rs.getObject(1);
                    return v != null ? rs.getInt(1) : null;
                }
            }
        }
        return null;
    }

    private static Integer getQuoteVendorId(Connection conn, int quoteId) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement("SELECT vendedor_id FROM cotizaciones WHERE id = ?")) {
            ps.setInt(1, quoteId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    Object v = rs.getObject(1);
                    return v != null ? rs.getInt(1) : null;
                }
            }
        }
        return null;
    }

    private static boolean existsQuote(Connection conn, int quoteId) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT 1 FROM cotizaciones WHERE id = ? AND deleted_at IS NULL")) {
            ps.setInt(1, quoteId);
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next();
            }
        }
    }

    private static boolean existsProject(Connection conn, int projectId) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT 1 FROM proyectos WHERE id = ? AND deleted_at IS NULL")) {
            ps.setInt(1, projectId);
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next();
            }
        }
    }

    private static boolean isQuoteProjectGenerated(Connection conn, int quoteId) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT proyecto_generado FROM cotizaciones WHERE id = ?")) {
            ps.setInt(1, quoteId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return rs.getBoolean(1);
                }
            }
        }
        return false;
    }

    // =========================================================
    // HELPERS
    // =========================================================
    private static Map<String, Object> normalizeState(Map<String, Object> state) {
        Map<String, Object> out = state != null ? new LinkedHashMap<>(state) : new LinkedHashMap<>();
        out.put("_appSource", APP_SOURCE);

        if (!(out.get("client") instanceof Map)) {
            out.put("client", new LinkedHashMap<String, Object>());
        }
        if (!(out.get("receipt") instanceof Map)) {
            out.put("receipt", new LinkedHashMap<String, Object>());
        }
        if (!(out.get("quote") instanceof Map)) {
            out.put("quote", new LinkedHashMap<String, Object>());
        }

        return out;
    }

    private static Map<String, Object> parseJsonObject(String raw) {
        if (raw == null || raw.isBlank()) {
            return new LinkedHashMap<>();
        }
        try {
            JsonElement element = JsonParser.parseString(raw);
            if (element != null && element.isJsonObject()) {
                @SuppressWarnings("unchecked")
                Map<String, Object> value = GSON.fromJson(element, Map.class);
                return value != null ? new LinkedHashMap<>(value) : new LinkedHashMap<>();
            }
        } catch (Exception ignore) {
        }
        return new LinkedHashMap<>();
    }

    @SuppressWarnings("unchecked")
    private static Map<String, Object> asMap(Object value) {
        if (value instanceof Map<?, ?> map) {
            return (Map<String, Object>) map;
        }
        return new LinkedHashMap<>();
    }

    private static Map<String, Object> ensureMap(Map<String, Object> root, String key) {
        Map<String, Object> map = asMap(root.get(key));
        if (map.isEmpty() && !(root.get(key) instanceof Map)) {
            map = new LinkedHashMap<>();
            root.put(key, map);
        }
        return map;
    }

    @SuppressWarnings("unchecked")
    private static List<Map<String, Object>> asListOfMaps(Object value) {
        List<Map<String, Object>> out = new ArrayList<>();
        if (value instanceof List<?> list) {
            for (Object item : list) {
                if (item instanceof Map<?, ?> map) {
                    out.add((Map<String, Object>) map);
                }
            }
        }
        return out;
    }

    private static String asString(Object value) {
        return value == null ? "" : String.valueOf(value).trim();
    }

    private static double asDouble(Object value) {
        if (value == null) {
            return 0.0;
        }
        if (value instanceof Number n) {
            return n.doubleValue();
        }
        try {
            String txt = String.valueOf(value).trim().replace(",", "");
            if (txt.isEmpty()) {
                return 0.0;
            }
            return Double.parseDouble(txt);
        } catch (Exception ex) {
            return 0.0;
        }
    }

    private static double positive(double value) {
        return value > 0 ? value : 0.0;
    }

    private static double firstPositive(double... values) {
        for (double v : values) {
            if (v > 0) {
                return v;
            }
        }
        return 0.0;
    }

    private static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }

    private static String firstNonBlank(String... values) {
        for (String v : values) {
            if (!isBlank(v)) {
                return v.trim();
            }
        }
        return "";
    }

    private static int parseFrontId(String value) {
        if (value == null || value.isBlank()) {
            return 0;
        }
        String digits = value.replaceAll("\\D+", "");
        if (digits.isBlank()) {
            return 0;
        }
        try {
            return Integer.parseInt(digits);
        } catch (Exception ex) {
            return 0;
        }
    }

    private static String formatQuoteId(int id) {
        return "COT-" + String.format("%03d", id);
    }

    private static String formatProjectId(int id) {
        return "PRO-" + String.format("%03d", id);
    }

    private static long toMillis(Timestamp ts) {
        return ts != null ? ts.getTime() : System.currentTimeMillis();
    }

    private static void setNullableInt(PreparedStatement ps, int index, Integer value) throws SQLException {
        if (value == null || value <= 0) {
            ps.setNull(index, java.sql.Types.INTEGER);
        } else {
            ps.setInt(index, value);
        }
    }

    private static String mapUiQuoteStatusToDb(String uiStatus, boolean projectGenerated) {
        if (projectGenerated) {
            return "FINALIZADA";
        }
        String s = asString(uiStatus).toLowerCase();
        if (s.contains("confirm")) {
            return "ACEPTADA";
        }
        return "COTIZADA";
    }

    private static String mapDbQuoteStatusToUi(String dbStatus, boolean projectGenerated) {
        if (projectGenerated) {
            return "Confirmada";
        }
        String s = asString(dbStatus).toUpperCase();
        return switch (s) {
            case "ACEPTADA", "FINALIZADA" -> "Confirmada";
            default -> "Guardada";
        };
    }

    private static String mapUiProjectStatusToDb(String uiStatus) {
        String s = asString(uiStatus).toLowerCase();
        if (s.contains("complet")) {
            return "COMPLETADO";
        }
        if (s.contains("instal") || s.contains("trámite") || s.contains("tramite") || s.contains("proceso")) {
            return "EN_PROCESO";
        }
        if (s.contains("paus")) {
            return "CANCELADO";
        }
        return "PENDIENTE";
    }

    private static String mapDbProjectStatusToUi(String dbStatus) {
        String s = asString(dbStatus).toUpperCase();
        return switch (s) {
            case "COMPLETADO", "INSTALADO" -> "Completado";
            case "EN_PROCESO" -> "En instalación";
            case "CANCELADO" -> "Pausado";
            default -> "En planeación";
        };
    }

    private static String mapTarifaEnum(String raw) {
        String s = asString(raw).toUpperCase();
        if (s.contains("GDMTH")) {
            return "GDMTH";
        }
        if (s.contains("GDMTO")) {
            return "GDMTO";
        }
        if (s.contains("PDBT")) {
            return "PDBT";
        }
        if (s.matches(".*\\b1[A-F]\\b.*")) {
            return s.replaceAll(".*\\b(1[A-F])\\b.*", "$1");
        }
        return "";
    }
}