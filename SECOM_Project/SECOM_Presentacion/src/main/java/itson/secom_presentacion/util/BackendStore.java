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
import java.time.LocalDate;
import java.time.format.DateTimeParseException;
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

    public static void deleteQuote(String frontId) throws Exception {
        int quoteId = parseFrontId(frontId);
        if (quoteId <= 0) {
            throw new IllegalArgumentException("ID de cotización inválido.");
        }

        ConnectionDB db = new ConnectionDB(false);
        try {
            Connection conn = db.getConexion();
            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE cotizaciones SET deleted_at = NOW(), updated_at = NOW() WHERE id = ? AND deleted_at IS NULL")) {
                ps.setInt(1, quoteId);
                int rows = ps.executeUpdate();
                if (rows == 0) {
                    throw new IllegalStateException("No se encontró la cotización indicada.");
                }
            }
        } finally {
            db.close();
        }
    }

    public static Map<String, Object> buildCotizacionesReport(String fechaInicioRaw, String fechaFinRaw,
            String statusFilter, String tarifaFilter) throws Exception {
        LocalDate fechaInicio = parseReportDate(fechaInicioRaw, "fecha inicial");
        LocalDate fechaFin = parseReportDate(fechaFinRaw, "fecha final");
        if (fechaInicio.isAfter(fechaFin)) {
            throw new IllegalArgumentException("La fecha inicial no puede ser mayor que la fecha final.");
        }

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
                  AND DATE(q.fecha) >= ?
                  AND DATE(q.fecha) <= ?
                ORDER BY q.fecha DESC, q.id DESC
            """;

            List<Map<String, Object>> rows = new ArrayList<>();
            try (PreparedStatement ps = conn.prepareStatement(sql)) {
                ps.setDate(1, java.sql.Date.valueOf(fechaInicio));
                ps.setDate(2, java.sql.Date.valueOf(fechaFin));
                try (ResultSet rs = ps.executeQuery()) {
                    while (rs.next()) {
                        Map<String, Object> quoteState = mapQuoteRow(rs);
                        Map<String, Object> row = mapQuoteReportRow(quoteState, rs.getTimestamp("fecha"), rs.getBoolean("proyecto_generado"));
                        if (!matchesReportStatus(row, statusFilter) || !matchesReportTariff(row, tarifaFilter)) {
                            continue;
                        }
                        rows.add(row);
                    }
                }
            }

            Map<String, Object> out = new LinkedHashMap<>();
            Map<String, Object> filters = new LinkedHashMap<>();
            filters.put("fechaInicio", fechaInicio.toString());
            filters.put("fechaFin", fechaFin.toString());
            filters.put("status", isBlank(asString(statusFilter)) ? "todos" : asString(statusFilter));
            filters.put("tarifa", isBlank(asString(tarifaFilter)) ? "todas" : asString(tarifaFilter));
            out.put("filters", filters);
            out.put("summary", buildQuoteReportSummary(rows));
            out.put("rows", rows);
            return out;
        } finally {
            db.close();
        }
    }

    public static int resetAllData() throws Exception {
        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);
            int affected = 0;

            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE proyecto_materiales pm JOIN proyectos p ON p.id = pm.proyecto_id SET pm.deleted_at = NOW() WHERE pm.deleted_at IS NULL AND p.deleted_at IS NULL")) {
                affected += ps.executeUpdate();
            } catch (SQLException ignore) {
                // La tabla puede no existir en instalaciones mínimas.
            }

            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE proyectos SET deleted_at = NOW(), updated_at = NOW() WHERE deleted_at IS NULL")) {
                affected += ps.executeUpdate();
            }

            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE cotizaciones SET deleted_at = NOW(), updated_at = NOW() WHERE deleted_at IS NULL")) {
                affected += ps.executeUpdate();
            }

            conn.commit();
            return affected;
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
    // INSUMOS
    // =========================================================
    public static List<Map<String, Object>> listInsumos() throws Exception {
        ConnectionDB db = new ConnectionDB(false);
        try {
            Connection conn = db.getConexion();
            ensureInsumosSchema(conn);
            seedDefaultInsumos(conn);

            String sql = """
                SELECT id, codigo, descripcion, categoria, unidad,
                       precio_unitario, impuesto_pct, activo, observaciones,
                       created_at, updated_at
                FROM insumos_catalogo
                WHERE deleted_at IS NULL
                ORDER BY activo DESC, descripcion ASC, id ASC
            """;

            List<Map<String, Object>> out = new ArrayList<>();
            try (PreparedStatement ps = conn.prepareStatement(sql);
                 ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    out.add(mapInsumoRow(rs));
                }
            }
            return out;
        } finally {
            db.close();
        }
    }

    public static Map<String, Object> saveInsumo(Map<String, Object> payload) throws Exception {
        Map<String, Object> insumo = normalizeInsumoPayload(payload);

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);
            ensureInsumosSchema(conn);
            validateUniqueInsumoCode(conn, asString(insumo.get("codigo")), null);

            String sql = """
                INSERT INTO insumos_catalogo
                (codigo, descripcion, categoria, unidad, precio_unitario,
                 impuesto_pct, activo, observaciones)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """;

            int id;
            try (PreparedStatement ps = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
                ps.setString(1, asString(insumo.get("codigo")));
                ps.setString(2, asString(insumo.get("descripcion")));
                ps.setString(3, asString(insumo.get("categoria")));
                ps.setString(4, asString(insumo.get("unidad")));
                ps.setDouble(5, positive(asDouble(insumo.get("precio"))));
                ps.setDouble(6, normalizePct(asDouble(insumo.get("impuestoPct")), 0.16));
                ps.setBoolean(7, asBoolean(insumo.get("activo"), true));
                ps.setString(8, asString(insumo.get("observaciones")));
                ps.executeUpdate();

                try (ResultSet rs = ps.getGeneratedKeys()) {
                    if (!rs.next()) {
                        throw new IllegalStateException("No se pudo generar el ID del insumo.");
                    }
                    id = rs.getInt(1);
                }
            }

            conn.commit();
            return getInsumoById(conn, id);
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

    public static Map<String, Object> updateInsumo(String frontId, Map<String, Object> patch) throws Exception {
        int id = parsePlainInt(frontId);
        if (id <= 0) {
            throw new IllegalArgumentException("ID de insumo inválido.");
        }

        Map<String, Object> insumo = normalizeInsumoPayload(patch);

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);
            ensureInsumosSchema(conn);
            if (!existsInsumo(conn, id)) {
                throw new IllegalStateException("No se encontró el insumo indicado.");
            }
            validateUniqueInsumoCode(conn, asString(insumo.get("codigo")), id);

            String sql = """
                UPDATE insumos_catalogo
                SET codigo = ?,
                    descripcion = ?,
                    categoria = ?,
                    unidad = ?,
                    precio_unitario = ?,
                    impuesto_pct = ?,
                    activo = ?,
                    observaciones = ?,
                    updated_at = NOW()
                WHERE id = ? AND deleted_at IS NULL
            """;

            try (PreparedStatement ps = conn.prepareStatement(sql)) {
                ps.setString(1, asString(insumo.get("codigo")));
                ps.setString(2, asString(insumo.get("descripcion")));
                ps.setString(3, asString(insumo.get("categoria")));
                ps.setString(4, asString(insumo.get("unidad")));
                ps.setDouble(5, positive(asDouble(insumo.get("precio"))));
                ps.setDouble(6, normalizePct(asDouble(insumo.get("impuestoPct")), 0.16));
                ps.setBoolean(7, asBoolean(insumo.get("activo"), true));
                ps.setString(8, asString(insumo.get("observaciones")));
                ps.setInt(9, id);
                ps.executeUpdate();
            }

            conn.commit();
            return getInsumoById(conn, id);
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

    public static void deleteInsumo(String frontId) throws Exception {
        int id = parsePlainInt(frontId);
        if (id <= 0) {
            throw new IllegalArgumentException("ID de insumo inválido.");
        }

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();
        try {
            conn.setAutoCommit(false);
            ensureInsumosSchema(conn);
            ensurePaquetesSchema(conn);

            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE paquete_insumos SET deleted_at = NOW(), updated_at = NOW() WHERE insumo_id = ? AND deleted_at IS NULL")) {
                ps.setInt(1, id);
                ps.executeUpdate();
            }

            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE insumos_catalogo SET deleted_at = NOW(), activo = FALSE, updated_at = NOW() WHERE id = ? AND deleted_at IS NULL")) {
                ps.setInt(1, id);
                int rows = ps.executeUpdate();
                if (rows == 0) {
                    throw new IllegalStateException("No se encontró el insumo indicado.");
                }
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

    public static List<Map<String, Object>> listPaquetesByInsumo(String frontId) throws Exception {
        int id = parsePlainInt(frontId);
        if (id <= 0) {
            throw new IllegalArgumentException("ID de insumo inválido.");
        }

        ConnectionDB db = new ConnectionDB(false);
        try {
            Connection conn = db.getConexion();
            ensureInsumosSchema(conn);
            ensurePaquetesSchema(conn);

            String sql = """
                SELECT DISTINCT p.id, p.clave, p.nombre, p.descripcion, p.badge,
                       p.activo, p.observaciones, p.created_at, p.updated_at
                FROM paquetes_catalogo p
                INNER JOIN paquete_insumos pi ON pi.paquete_id = p.id
                WHERE pi.insumo_id = ?
                  AND pi.deleted_at IS NULL
                  AND p.deleted_at IS NULL
                ORDER BY p.nombre ASC, p.id ASC
            """;

            List<Map<String, Object>> out = new ArrayList<>();
            try (PreparedStatement ps = conn.prepareStatement(sql)) {
                ps.setInt(1, id);
                try (ResultSet rs = ps.executeQuery()) {
                    while (rs.next()) {
                        out.add(mapPaqueteRow(conn, rs));
                    }
                }
            }
            return out;
        } finally {
            db.close();
        }
    }


    // =========================================================
    // PAQUETES
    // =========================================================
    public static List<Map<String, Object>> listPaquetes() throws Exception {
        ConnectionDB db = new ConnectionDB(false);
        try {
            Connection conn = db.getConexion();
            ensureInsumosSchema(conn);
            seedDefaultInsumos(conn);
            ensurePaquetesSchema(conn);
            seedDefaultPaquetes(conn);

            String sql = """
                SELECT id, clave, nombre, descripcion, badge, activo, observaciones,
                       created_at, updated_at
                FROM paquetes_catalogo
                WHERE deleted_at IS NULL
                ORDER BY activo DESC, nombre ASC, id ASC
            """;

            List<Map<String, Object>> out = new ArrayList<>();
            try (PreparedStatement ps = conn.prepareStatement(sql);
                 ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    out.add(mapPaqueteRow(conn, rs));
                }
            }
            return out;
        } finally {
            db.close();
        }
    }

    public static Map<String, Object> savePaquete(Map<String, Object> payload) throws Exception {
        Map<String, Object> paquete = normalizePaquetePayload(payload);

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);
            ensureInsumosSchema(conn);
            seedDefaultInsumos(conn);
            ensurePaquetesSchema(conn);

            String sql = """
                INSERT INTO paquetes_catalogo
                (clave, nombre, descripcion, badge, activo, observaciones)
                VALUES (?, ?, ?, ?, ?, ?)
            """;

            int id;
            try (PreparedStatement ps = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
                ps.setString(1, generatePackageKey(conn, asString(paquete.get("nombre")), null));
                ps.setString(2, asString(paquete.get("nombre")));
                ps.setString(3, asString(paquete.get("descripcion")));
                ps.setString(4, asString(paquete.get("badge")));
                ps.setBoolean(5, asBoolean(paquete.get("activo"), true));
                ps.setString(6, asString(paquete.get("observaciones")));
                ps.executeUpdate();
                try (ResultSet rs = ps.getGeneratedKeys()) {
                    if (!rs.next()) {
                        throw new IllegalStateException("No se pudo generar el ID del paquete.");
                    }
                    id = rs.getInt(1);
                }
            }

            replacePaqueteInsumos(conn, id, asListOfMaps(paquete.get("items")));
            conn.commit();
            return getPaqueteById(conn, id);
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

    public static Map<String, Object> updatePaquete(String frontId, Map<String, Object> patch) throws Exception {
        int id = parsePlainInt(frontId);
        if (id <= 0) {
            throw new IllegalArgumentException("ID de paquete inválido.");
        }

        Map<String, Object> paquete = normalizePaquetePayload(patch);

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();

        try {
            conn.setAutoCommit(false);
            ensureInsumosSchema(conn);
            seedDefaultInsumos(conn);
            ensurePaquetesSchema(conn);
            if (!existsPaquete(conn, id)) {
                throw new IllegalStateException("No se encontró el paquete indicado.");
            }

            String sql = """
                UPDATE paquetes_catalogo
                SET clave = ?,
                    nombre = ?,
                    descripcion = ?,
                    badge = ?,
                    activo = ?,
                    observaciones = ?,
                    updated_at = NOW()
                WHERE id = ? AND deleted_at IS NULL
            """;

            try (PreparedStatement ps = conn.prepareStatement(sql)) {
                ps.setString(1, generatePackageKey(conn, asString(paquete.get("nombre")), id));
                ps.setString(2, asString(paquete.get("nombre")));
                ps.setString(3, asString(paquete.get("descripcion")));
                ps.setString(4, asString(paquete.get("badge")));
                ps.setBoolean(5, asBoolean(paquete.get("activo"), true));
                ps.setString(6, asString(paquete.get("observaciones")));
                ps.setInt(7, id);
                ps.executeUpdate();
            }

            replacePaqueteInsumos(conn, id, asListOfMaps(paquete.get("items")));
            conn.commit();
            return getPaqueteById(conn, id);
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

    public static void deletePaquete(String frontId) throws Exception {
        int id = parsePlainInt(frontId);
        if (id <= 0) {
            throw new IllegalArgumentException("ID de paquete inválido.");
        }

        ConnectionDB db = new ConnectionDB(false);
        Connection conn = db.getConexion();
        try {
            conn.setAutoCommit(false);
            ensurePaquetesSchema(conn);
            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE paquete_insumos SET deleted_at = NOW(), updated_at = NOW() WHERE paquete_id = ? AND deleted_at IS NULL")) {
                ps.setInt(1, id);
                ps.executeUpdate();
            }
            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE paquetes_catalogo SET deleted_at = NOW(), activo = FALSE, updated_at = NOW() WHERE id = ? AND deleted_at IS NULL")) {
                ps.setInt(1, id);
                int rows = ps.executeUpdate();
                if (rows == 0) {
                    throw new IllegalStateException("No se encontró el paquete indicado.");
                }
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
    // INTERNALS: INSUMOS
    // =========================================================
    private static void ensureInsumosSchema(Connection conn) throws Exception {
        String sql = """
            CREATE TABLE IF NOT EXISTS insumos_catalogo (
                id INT AUTO_INCREMENT PRIMARY KEY,
                codigo VARCHAR(50) NOT NULL,
                descripcion VARCHAR(255) NOT NULL,
                categoria VARCHAR(80) NOT NULL DEFAULT 'General',
                unidad VARCHAR(20) NOT NULL DEFAULT 'UD',
                precio_unitario DECIMAL(12,2) NOT NULL DEFAULT 0,
                impuesto_pct DECIMAL(6,4) NOT NULL DEFAULT 0.1600,
                activo BOOLEAN NOT NULL DEFAULT TRUE,
                observaciones TEXT NULL,
                created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                deleted_at TIMESTAMP NULL DEFAULT NULL,
                INDEX idx_insumos_catalogo_codigo (codigo),
                INDEX idx_insumos_catalogo_activo (activo),
                INDEX idx_insumos_catalogo_deleted (deleted_at)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.executeUpdate();
        }
    }

    private static void seedDefaultInsumos(Connection conn) throws Exception {
        // Considera únicamente insumos visibles para evitar catálogos vacíos
        // cuando la tabla existe pero todos los registros fueron desactivados por eliminación lógica.
        String countSql = "SELECT COUNT(*) FROM insumos_catalogo WHERE deleted_at IS NULL";
        try (PreparedStatement ps = conn.prepareStatement(countSql);
             ResultSet rs = ps.executeQuery()) {
            if (rs.next() && rs.getInt(1) > 0) {
                return;
            }
        }

        insertDefaultInsumo(conn, "PANEL-550", "Panel solar monocristalino 550 W (paneles)", "Paneles", "PZA", 3200, 0.16);
        insertDefaultInsumo(conn, "PANEL-610", "Panel solar monocristalino 610 W (paneles premium)", "Paneles", "PZA", 3950, 0.16);
        insertDefaultInsumo(conn, "INV-STR", "Inversor interconectado string (inversor)", "Inversores", "PZA", 18500, 0.16);
        insertDefaultInsumo(conn, "INV-HIB", "Inversor híbrido con monitoreo (inversor)", "Inversores", "PZA", 36500, 0.16);
        insertDefaultInsumo(conn, "EST-AL", "Estructura de aluminio para azotea o techo (estructura)", "Estructura", "SERV", 9800, 0.16);
        insertDefaultInsumo(conn, "EST-LA", "Estructura reforzada para lámina o teja (estructura)", "Estructura", "SERV", 14500, 0.16);
        insertDefaultInsumo(conn, "CAB-FV", "Cable fotovoltaico y conectores MC4 (cableado)", "Cableado", "SERV", 4200, 0.16);
        insertDefaultInsumo(conn, "PROT-CC", "Protecciones CC/CA y tablero de interconexión (protecciones)", "Protecciones", "SERV", 7600, 0.16);
        insertDefaultInsumo(conn, "MON-APP", "Monitoreo remoto y puesta en marcha (monitoreo)", "Monitoreo", "SERV", 3900, 0.16);
        insertDefaultInsumo(conn, "TIERRA", "Puesta a tierra y canalización (seguridad)", "Seguridad", "SERV", 5600, 0.16);
        insertDefaultInsumo(conn, "TRAM-CFE", "Trámite de interconexión ante CFE (trámite)", "Trámite", "SERV", 4800, 0.16);
        insertDefaultInsumo(conn, "MO-INST", "Mano de obra de instalación (instalación)", "Instalación", "SERV", 12600, 0.16);
        insertDefaultInsumo(conn, "FLETE", "Flete, maniobras y logística (logística)", "Logística", "SERV", 3400, 0.16);
        insertDefaultInsumo(conn, "BAT-LFP", "Banco de baterías LiFePO4 de respaldo (baterías)", "Baterías", "PZA", 48500, 0.16);
        insertDefaultInsumo(conn, "ING", "Ingeniería, planos y memoria técnica (ingeniería)", "Ingeniería", "SERV", 5200, 0.16);
    }

    private static void insertDefaultInsumo(Connection conn, String codigo, String descripcion,
            String categoria, String unidad, double precio, double impuestoPct) throws Exception {
        String sql = """
            INSERT INTO insumos_catalogo
            (codigo, descripcion, categoria, unidad, precio_unitario, impuesto_pct, activo, observaciones)
            VALUES (?, ?, ?, ?, ?, ?, TRUE, 'Carga inicial del catálogo SECOM')
        """;
        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setString(1, codigo);
            ps.setString(2, descripcion);
            ps.setString(3, categoria);
            ps.setString(4, unidad);
            ps.setDouble(5, precio);
            ps.setDouble(6, impuestoPct);
            ps.executeUpdate();
        }
    }

    private static Map<String, Object> getInsumoById(Connection conn, int id) throws Exception {
        String sql = """
            SELECT id, codigo, descripcion, categoria, unidad,
                   precio_unitario, impuesto_pct, activo, observaciones,
                   created_at, updated_at
            FROM insumos_catalogo
            WHERE id = ? AND deleted_at IS NULL
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setInt(1, id);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return mapInsumoRow(rs);
                }
            }
        }

        throw new IllegalStateException("No se encontró el insumo solicitado.");
    }

    private static boolean existsInsumo(Connection conn, int id) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT 1 FROM insumos_catalogo WHERE id = ? AND deleted_at IS NULL")) {
            ps.setInt(1, id);
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next();
            }
        }
    }

    private static void validateUniqueInsumoCode(Connection conn, String codigo, Integer excludeId) throws Exception {
        String sql = excludeId == null
                ? "SELECT id FROM insumos_catalogo WHERE codigo = ? AND deleted_at IS NULL LIMIT 1"
                : "SELECT id FROM insumos_catalogo WHERE codigo = ? AND id <> ? AND deleted_at IS NULL LIMIT 1";

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setString(1, codigo);
            if (excludeId != null) {
                ps.setInt(2, excludeId);
            }
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    throw new IllegalArgumentException("El código de insumo ya se encuentra registrado.");
                }
            }
        }
    }

    private static Map<String, Object> normalizeInsumoPayload(Map<String, Object> payload) {
        Map<String, Object> in = payload != null ? payload : new LinkedHashMap<>();
        String codigo = asString(in.get("codigo")).toUpperCase();
        String descripcion = asString(in.get("descripcion"));
        String categoria = firstNonBlank(asString(in.get("categoria")), "General");
        String unidad = firstNonBlank(asString(in.get("unidad")).toUpperCase(), "UD");
        double precio = asDouble(in.get("precio"));
        double impuestoPct = normalizePct(asDouble(in.get("impuestoPct")), 0.16);
        boolean activo = asBoolean(in.get("activo"), true);
        String observaciones = asString(in.get("observaciones"));

        if (isBlank(codigo)) {
            throw new IllegalArgumentException("El código del insumo es obligatorio.");
        }
        if (isBlank(descripcion)) {
            throw new IllegalArgumentException("La descripción del insumo es obligatoria.");
        }
        if (isBlank(unidad)) {
            throw new IllegalArgumentException("La unidad de medida es obligatoria.");
        }
        if (precio < 0) {
            throw new IllegalArgumentException("El precio unitario debe ser mayor o igual a cero.");
        }

        Map<String, Object> out = new LinkedHashMap<>();
        out.put("codigo", codigo);
        out.put("descripcion", descripcion);
        out.put("categoria", categoria);
        out.put("unidad", unidad);
        out.put("precio", precio);
        out.put("impuestoPct", impuestoPct);
        out.put("activo", activo);
        out.put("observaciones", observaciones);
        return out;
    }

    private static Map<String, Object> mapInsumoRow(ResultSet rs) throws Exception {
        Map<String, Object> out = new LinkedHashMap<>();
        out.put("id", rs.getInt("id"));
        out.put("codigo", asString(rs.getString("codigo")));
        out.put("descripcion", asString(rs.getString("descripcion")));
        out.put("categoria", asString(rs.getString("categoria")));
        out.put("unidad", asString(rs.getString("unidad")));
        out.put("precio", rs.getDouble("precio_unitario"));
        out.put("precioUnitario", rs.getDouble("precio_unitario"));
        out.put("impuestoPct", rs.getDouble("impuesto_pct"));
        out.put("activo", rs.getBoolean("activo"));
        out.put("estatus", rs.getBoolean("activo") ? "Activo" : "Inactivo");
        out.put("observaciones", asString(rs.getString("observaciones")));
        out.put("usoPaquetes", 0);
        out.put("createdAt", toMillis(rs.getTimestamp("created_at")));
        out.put("updatedAt", toMillis(rs.getTimestamp("updated_at")));
        return out;
    }


    // =========================================================
    // INTERNALS: PAQUETES
    // =========================================================
    private static void ensurePaquetesSchema(Connection conn) throws Exception {
        String paquetesSql = """
            CREATE TABLE IF NOT EXISTS paquetes_catalogo (
                id INT AUTO_INCREMENT PRIMARY KEY,
                clave VARCHAR(120) NOT NULL,
                nombre VARCHAR(160) NOT NULL,
                descripcion TEXT NULL,
                badge VARCHAR(80) NOT NULL DEFAULT 'Paquete',
                activo BOOLEAN NOT NULL DEFAULT TRUE,
                observaciones TEXT NULL,
                created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                deleted_at TIMESTAMP NULL DEFAULT NULL,
                INDEX idx_paquetes_catalogo_clave (clave),
                INDEX idx_paquetes_catalogo_activo (activo),
                INDEX idx_paquetes_catalogo_deleted (deleted_at)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """;

        String insumosSql = """
            CREATE TABLE IF NOT EXISTS paquete_insumos (
                id INT AUTO_INCREMENT PRIMARY KEY,
                paquete_id INT NOT NULL,
                insumo_id INT NULL,
                codigo_snapshot VARCHAR(80) NOT NULL,
                descripcion_snapshot VARCHAR(255) NOT NULL,
                unidad_snapshot VARCHAR(20) NOT NULL DEFAULT 'UD',
                cantidad DECIMAL(12,4) NOT NULL DEFAULT 1,
                precio_snapshot DECIMAL(14,2) NOT NULL DEFAULT 0,
                impuesto_snapshot DECIMAL(6,4) NOT NULL DEFAULT 0.1600,
                created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                deleted_at TIMESTAMP NULL DEFAULT NULL,
                INDEX idx_paquete_insumos_paquete (paquete_id),
                INDEX idx_paquete_insumos_insumo (insumo_id),
                INDEX idx_paquete_insumos_deleted (deleted_at)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """;

        try (PreparedStatement ps = conn.prepareStatement(paquetesSql)) {
            ps.executeUpdate();
        }
        try (PreparedStatement ps = conn.prepareStatement(insumosSql)) {
            ps.executeUpdate();
        }
    }

    private static void seedDefaultPaquetes(Connection conn) throws Exception {
        // La carga inicial debe considerar solo paquetes visibles.
        // Si la tabla existe pero los paquetes fueron marcados como eliminados,
        // el catálogo quedaba vacío y el cotizador terminaba usando presets locales
        // que no aparecían en el CRUD de Paquetes.
        String countSql = "SELECT COUNT(*) FROM paquetes_catalogo WHERE deleted_at IS NULL";
        try (PreparedStatement ps = conn.prepareStatement(countSql);
             ResultSet rs = ps.executeQuery()) {
            if (rs.next() && rs.getInt(1) > 0) {
                return;
            }
        }

        int basico = insertDefaultPaquete(conn, "basico", "Paquete básico",
                "Interconexión esencial con componentes estándar y costo contenido.",
                "Recomendado para arranque");
        addDefaultPackageItem(conn, basico, "PANEL-550", 1);
        addDefaultPackageItem(conn, basico, "INV-STR", 1);
        addDefaultPackageItem(conn, basico, "EST-AL", 1);
        addDefaultPackageItem(conn, basico, "CAB-FV", 1);
        addDefaultPackageItem(conn, basico, "PROT-CC", 1);
        addDefaultPackageItem(conn, basico, "MO-INST", 1);
        addDefaultPackageItem(conn, basico, "TRAM-CFE", 1);

        int intermedio = insertDefaultPaquete(conn, "intermedio", "Paquete intermedio",
                "Mejor balance entre rendimiento, protecciones y monitoreo.",
                "Balanceado");
        addDefaultPackageItem(conn, intermedio, "PANEL-550", 1);
        addDefaultPackageItem(conn, intermedio, "INV-STR", 1);
        addDefaultPackageItem(conn, intermedio, "EST-LA", 1);
        addDefaultPackageItem(conn, intermedio, "CAB-FV", 1);
        addDefaultPackageItem(conn, intermedio, "PROT-CC", 1);
        addDefaultPackageItem(conn, intermedio, "MON-APP", 1);
        addDefaultPackageItem(conn, intermedio, "TIERRA", 1);
        addDefaultPackageItem(conn, intermedio, "MO-INST", 1);
        addDefaultPackageItem(conn, intermedio, "TRAM-CFE", 1);
        addDefaultPackageItem(conn, intermedio, "FLETE", 1);

        int avanzado = insertDefaultPaquete(conn, "avanzado", "Paquete avanzado",
                "Componentes premium, monitoreo extendido y preparación para respaldo.",
                "Premium");
        addDefaultPackageItem(conn, avanzado, "PANEL-610", 1);
        addDefaultPackageItem(conn, avanzado, "INV-HIB", 1);
        addDefaultPackageItem(conn, avanzado, "EST-LA", 1);
        addDefaultPackageItem(conn, avanzado, "CAB-FV", 1);
        addDefaultPackageItem(conn, avanzado, "PROT-CC", 1);
        addDefaultPackageItem(conn, avanzado, "MON-APP", 1);
        addDefaultPackageItem(conn, avanzado, "TIERRA", 1);
        addDefaultPackageItem(conn, avanzado, "ING", 1);
        addDefaultPackageItem(conn, avanzado, "MO-INST", 1);
        addDefaultPackageItem(conn, avanzado, "TRAM-CFE", 1);
        addDefaultPackageItem(conn, avanzado, "FLETE", 1);
    }

    private static int insertDefaultPaquete(Connection conn, String clave, String nombre, String descripcion, String badge) throws Exception {
        String sql = """
            INSERT INTO paquetes_catalogo
            (clave, nombre, descripcion, badge, activo, observaciones)
            VALUES (?, ?, ?, ?, TRUE, 'Carga inicial del catálogo SECOM')
        """;
        try (PreparedStatement ps = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            ps.setString(1, clave);
            ps.setString(2, nombre);
            ps.setString(3, descripcion);
            ps.setString(4, badge);
            ps.executeUpdate();
            try (ResultSet rs = ps.getGeneratedKeys()) {
                if (rs.next()) {
                    return rs.getInt(1);
                }
            }
        }
        throw new IllegalStateException("No se pudo crear el paquete predeterminado.");
    }

    private static void addDefaultPackageItem(Connection conn, int paqueteId, String codigo, double cantidad) throws Exception {
        Map<String, Object> insumo = getInsumoByCodigo(conn, codigo);
        if (insumo == null) {
            return;
        }
        insertPaqueteInsumo(conn, paqueteId, asMap(insumo), cantidad);
    }

    private static Map<String, Object> getInsumoByCodigo(Connection conn, String codigo) throws Exception {
        String sql = """
            SELECT id, codigo, descripcion, categoria, unidad,
                   precio_unitario, impuesto_pct, activo, observaciones,
                   created_at, updated_at
            FROM insumos_catalogo
            WHERE codigo = ? AND deleted_at IS NULL
            LIMIT 1
        """;
        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setString(1, asString(codigo).toUpperCase());
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return mapInsumoRow(rs);
                }
            }
        }
        return null;
    }

    private static Map<String, Object> getInsumoRef(Connection conn, Map<String, Object> item) throws Exception {
        int id = parsePlainInt(firstNonBlank(asString(item.get("insumoId")), asString(item.get("catalogId")), asString(item.get("insumo_id"))));
        if (id > 0) {
            try {
                return getInsumoById(conn, id);
            } catch (Exception ignore) {
            }
        }
        String codigo = asString(item.get("codigo")).toUpperCase();
        if (!isBlank(codigo)) {
            Map<String, Object> byCode = getInsumoByCodigo(conn, codigo);
            if (byCode != null) {
                return byCode;
            }
        }
        throw new IllegalArgumentException("Cada insumo del paquete debe existir en el catálogo de insumos.");
    }

    private static Map<String, Object> getPaqueteById(Connection conn, int id) throws Exception {
        String sql = """
            SELECT id, clave, nombre, descripcion, badge, activo, observaciones,
                   created_at, updated_at
            FROM paquetes_catalogo
            WHERE id = ? AND deleted_at IS NULL
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setInt(1, id);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    return mapPaqueteRow(conn, rs);
                }
            }
        }
        throw new IllegalStateException("No se encontró el paquete solicitado.");
    }

    private static boolean existsPaquete(Connection conn, int id) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT 1 FROM paquetes_catalogo WHERE id = ? AND deleted_at IS NULL")) {
            ps.setInt(1, id);
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next();
            }
        }
    }

    private static Map<String, Object> normalizePaquetePayload(Map<String, Object> payload) {
        Map<String, Object> in = payload != null ? payload : new LinkedHashMap<>();
        String nombre = firstNonBlank(asString(in.get("nombre")), asString(in.get("label")));
        String descripcion = firstNonBlank(asString(in.get("descripcion")), asString(in.get("description")));
        String badge = firstNonBlank(asString(in.get("badge")), "Paquete");
        boolean activo = asBoolean(in.get("activo"), true);
        String observaciones = asString(in.get("observaciones"));
        List<Map<String, Object>> items = asListOfMaps(firstNonNull(in.get("items"), in.get("insumos")));

        if (isBlank(nombre)) {
            throw new IllegalArgumentException("El nombre del paquete es obligatorio.");
        }
        if (items.isEmpty()) {
            throw new IllegalArgumentException("El paquete debe tener al menos un insumo asociado.");
        }

        for (Map<String, Object> item : items) {
            double cantidad = asDouble(item.get("cantidad"));
            if (cantidad <= 0) {
                throw new IllegalArgumentException("Todas las cantidades deben ser mayores a cero.");
            }
        }

        Map<String, Object> out = new LinkedHashMap<>();
        out.put("nombre", nombre);
        out.put("descripcion", descripcion);
        out.put("badge", badge);
        out.put("activo", activo);
        out.put("observaciones", observaciones);
        out.put("items", items);
        return out;
    }

    private static Object firstNonNull(Object... values) {
        for (Object v : values) {
            if (v != null) {
                return v;
            }
        }
        return null;
    }

    private static void replacePaqueteInsumos(Connection conn, int paqueteId, List<Map<String, Object>> items) throws Exception {
        try (PreparedStatement ps = conn.prepareStatement(
                "UPDATE paquete_insumos SET deleted_at = NOW(), updated_at = NOW() WHERE paquete_id = ? AND deleted_at IS NULL")) {
            ps.setInt(1, paqueteId);
            ps.executeUpdate();
        }

        for (Map<String, Object> item : items) {
            Map<String, Object> insumo = getInsumoRef(conn, item);
            insertPaqueteInsumo(conn, paqueteId, insumo, asDouble(item.get("cantidad")));
        }
    }

    private static void insertPaqueteInsumo(Connection conn, int paqueteId, Map<String, Object> insumo, double cantidad) throws Exception {
        String sql = """
            INSERT INTO paquete_insumos
            (paquete_id, insumo_id, codigo_snapshot, descripcion_snapshot,
             unidad_snapshot, cantidad, precio_snapshot, impuesto_snapshot)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """;

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setInt(1, paqueteId);
            ps.setInt(2, parsePlainInt(asString(insumo.get("id"))));
            ps.setString(3, asString(insumo.get("codigo")));
            ps.setString(4, asString(insumo.get("descripcion")));
            ps.setString(5, asString(insumo.get("unidad")));
            ps.setDouble(6, cantidad > 0 ? cantidad : 1.0);
            ps.setDouble(7, positive(asDouble(firstNonNull(insumo.get("precio"), insumo.get("precioUnitario")))));
            ps.setDouble(8, normalizePct(asDouble(insumo.get("impuestoPct")), 0.16));
            ps.executeUpdate();
        }
    }

    private static Map<String, Object> mapPaqueteRow(Connection conn, ResultSet rs) throws Exception {
        Map<String, Object> out = new LinkedHashMap<>();
        out.put("id", rs.getInt("id"));
        out.put("key", asString(rs.getString("clave")));
        out.put("clave", asString(rs.getString("clave")));
        out.put("nombre", asString(rs.getString("nombre")));
        out.put("label", asString(rs.getString("nombre")));
        out.put("descripcion", asString(rs.getString("descripcion")));
        out.put("description", asString(rs.getString("descripcion")));
        out.put("badge", asString(rs.getString("badge")));
        out.put("activo", rs.getBoolean("activo"));
        out.put("estatus", rs.getBoolean("activo") ? "Activo" : "Inactivo");
        out.put("observaciones", asString(rs.getString("observaciones")));
        out.put("createdAt", toMillis(rs.getTimestamp("created_at")));
        out.put("updatedAt", toMillis(rs.getTimestamp("updated_at")));

        List<Map<String, Object>> items = listPaqueteInsumos(conn, rs.getInt("id"));
        out.put("items", items);
        out.put("insumos", items);

        double subtotal = 0;
        double impuestos = 0;
        for (Map<String, Object> item : items) {
            double cantidad = asDouble(item.get("cantidad"));
            double precio = positive(asDouble(firstNonNull(item.get("precio"), item.get("precioUnitario"))));
            double impuestoPct = normalizePct(asDouble(item.get("impuestoPct")), 0.16);
            subtotal += cantidad * precio;
            impuestos += cantidad * precio * impuestoPct;
        }
        out.put("subtotal", subtotal);
        out.put("impuestos", impuestos);
        out.put("total", subtotal + impuestos);
        return out;
    }

    private static List<Map<String, Object>> listPaqueteInsumos(Connection conn, int paqueteId) throws Exception {
        String sql = """
            SELECT pi.id, pi.paquete_id, pi.insumo_id,
                   pi.codigo_snapshot, pi.descripcion_snapshot, pi.unidad_snapshot,
                   pi.cantidad, pi.precio_snapshot, pi.impuesto_snapshot,
                   i.codigo AS ins_codigo, i.descripcion AS ins_descripcion,
                   i.unidad AS ins_unidad, i.precio_unitario AS ins_precio,
                   i.impuesto_pct AS ins_impuesto, i.activo AS ins_activo,
                   i.deleted_at AS ins_deleted
            FROM paquete_insumos pi
            LEFT JOIN insumos_catalogo i ON i.id = pi.insumo_id
            WHERE pi.paquete_id = ? AND pi.deleted_at IS NULL
            ORDER BY pi.id ASC
        """;
        List<Map<String, Object>> out = new ArrayList<>();
        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setInt(1, paqueteId);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    boolean insumoActivo = rs.getTimestamp("ins_deleted") == null && rs.getBoolean("ins_activo");
                    Map<String, Object> item = new LinkedHashMap<>();
                    item.put("id", rs.getInt("id"));
                    item.put("paqueteId", paqueteId);
                    item.put("insumoId", rs.getObject("insumo_id") != null ? rs.getInt("insumo_id") : null);
                    item.put("catalogId", rs.getObject("insumo_id") != null ? rs.getInt("insumo_id") : null);
                    item.put("codigo", insumoActivo && !isBlank(rs.getString("ins_codigo")) ? asString(rs.getString("ins_codigo")) : asString(rs.getString("codigo_snapshot")));
                    item.put("descripcion", insumoActivo && !isBlank(rs.getString("ins_descripcion")) ? asString(rs.getString("ins_descripcion")) : asString(rs.getString("descripcion_snapshot")));
                    item.put("unidad", insumoActivo && !isBlank(rs.getString("ins_unidad")) ? asString(rs.getString("ins_unidad")) : asString(rs.getString("unidad_snapshot")));
                    item.put("cantidad", rs.getDouble("cantidad"));
                    item.put("precio", insumoActivo ? rs.getDouble("ins_precio") : rs.getDouble("precio_snapshot"));
                    item.put("precioUnitario", insumoActivo ? rs.getDouble("ins_precio") : rs.getDouble("precio_snapshot"));
                    item.put("impuestoPct", insumoActivo ? rs.getDouble("ins_impuesto") : rs.getDouble("impuesto_snapshot"));
                    item.put("activo", insumoActivo);
                    out.add(item);
                }
            }
        }
        return out;
    }

    private static String generatePackageKey(Connection conn, String nombre, Integer excludeId) throws Exception {
        String base = slugify(asString(nombre));
        if (isBlank(base)) {
            base = "paquete";
        }
        String candidate = base;
        int suffix = 2;
        while (existsPackageKey(conn, candidate, excludeId)) {
            candidate = base + "-" + suffix;
            suffix++;
        }
        return candidate;
    }

    private static boolean existsPackageKey(Connection conn, String clave, Integer excludeId) throws Exception {
        String sql = excludeId == null
                ? "SELECT id FROM paquetes_catalogo WHERE clave = ? AND deleted_at IS NULL LIMIT 1"
                : "SELECT id FROM paquetes_catalogo WHERE clave = ? AND id <> ? AND deleted_at IS NULL LIMIT 1";
        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setString(1, clave);
            if (excludeId != null) {
                ps.setInt(2, excludeId);
            }
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next();
            }
        }
    }

    private static String slugify(String value) {
        String s = asString(value).toLowerCase();
        s = java.text.Normalizer.normalize(s, java.text.Normalizer.Form.NFD)
                .replaceAll("\\p{M}", "")
                .replaceAll("[^a-z0-9]+", "-")
                .replaceAll("^-+|-+$", "");
        return isBlank(s) ? "paquete" : s;
    }

    // =========================================================
    // INTERNALS: REPORTS
    // =========================================================
    private static LocalDate parseReportDate(String raw, String fieldName) {
        if (isBlank(asString(raw))) {
            throw new IllegalArgumentException("La " + fieldName + " es obligatoria.");
        }
        try {
            return LocalDate.parse(asString(raw));
        } catch (DateTimeParseException ex) {
            throw new IllegalArgumentException("La " + fieldName + " no tiene un formato válido (AAAA-MM-DD).");
        }
    }

    private static Map<String, Object> mapQuoteReportRow(Map<String, Object> state, Timestamp fecha,
            boolean proyectoGenerado) {
        Map<String, Object> client = asMap(state.get("client"));
        Map<String, Object> receipt = asMap(state.get("receipt"));
        Map<String, Object> quote = asMap(state.get("quote"));
        Map<String, Object> selectedTariff = asMap(state.get("selectedTariff"));
        Map<String, Object> tarifaCalculo = asMap(receipt.get("tarifaCalculo"));

        double consumoMensual = firstPositive(
                asDouble(quote.get("consumoMensual")),
                asDouble(tarifaCalculo.get("consumoMensualBase")),
                asDouble(receipt.get("consumoPeriodo"))
        );
        double inversion = firstPositive(
                asDouble(quote.get("totalInsumos")),
                asDouble(quote.get("inversion")),
                asDouble(receipt.get("totalAPagar"))
        );

        Map<String, Object> row = new LinkedHashMap<>();
        row.put("folio", asString(state.get("id")));
        row.put("id", asString(state.get("id")));
        row.put("fecha", fecha != null ? fecha.toLocalDateTime().toLocalDate().toString() : "");
        row.put("fechaTexto", fecha != null ? fecha.toLocalDateTime().toLocalDate().toString() : "");
        row.put("fechaMillis", toMillis(fecha));
        row.put("cliente", firstNonBlank(asString(client.get("nombre")), asString(receipt.get("nombre")), "Sin cliente"));
        row.put("servicio", asString(receipt.get("servicio")));
        row.put("tarifa", firstNonBlank(
                asString(selectedTariff.get("label")),
                asString(receipt.get("tarifaSeleccionada")),
                asString(receipt.get("tarifa")),
                "Sin tarifa"));
        row.put("consumoMensual", consumoMensual);
        row.put("paneles", positive(asDouble(quote.get("paneles"))));
        row.put("potenciaKwp", firstPositive(asDouble(quote.get("kwp")), asDouble(quote.get("kwpFinal"))));
        row.put("inversion", inversion);
        row.put("ahorroMensual", firstPositive(asDouble(quote.get("ahorroMensual")), asDouble(receipt.get("ahorroEstimado"))));
        row.put("retornoAnios", positive(asDouble(quote.get("retornoAnios"))));
        row.put("estatus", firstNonBlank(asString(state.get("status")), "Guardada"));
        row.put("usuario", firstNonBlank(asString(state.get("usuario")), asString(state.get("createdBy")), "Equipo SECOM"));
        row.put("proyectoGenerado", proyectoGenerado);
        row.put("source", state);
        return row;
    }

    private static boolean matchesReportStatus(Map<String, Object> row, String statusFilter) {
        String filter = asString(statusFilter);
        if (isBlank(filter) || "todos".equalsIgnoreCase(filter)) {
            return true;
        }
        return asString(row.get("estatus")).equalsIgnoreCase(filter);
    }

    private static boolean matchesReportTariff(Map<String, Object> row, String tarifaFilter) {
        String filter = asString(tarifaFilter);
        if (isBlank(filter) || "todas".equalsIgnoreCase(filter)) {
            return true;
        }
        return asString(row.get("tarifa")).equalsIgnoreCase(filter);
    }

    private static Map<String, Object> buildQuoteReportSummary(List<Map<String, Object>> rows) {
        Map<String, Object> summary = new LinkedHashMap<>();
        int total = rows.size();
        int confirmadas = 0;
        int convertidasProyecto = 0;
        double montoTotal = 0.0;
        double consumoTotal = 0.0;
        double potenciaTotal = 0.0;
        double ahorroTotal = 0.0;
        double panelesTotal = 0.0;

        for (Map<String, Object> row : rows) {
            String status = asString(row.get("estatus")).toLowerCase();
            if (status.contains("confirm")) {
                confirmadas++;
            }
            if (asBoolean(row.get("proyectoGenerado"), false)) {
                convertidasProyecto++;
            }
            montoTotal += positive(asDouble(row.get("inversion")));
            consumoTotal += positive(asDouble(row.get("consumoMensual")));
            potenciaTotal += positive(asDouble(row.get("potenciaKwp")));
            ahorroTotal += positive(asDouble(row.get("ahorroMensual")));
            panelesTotal += positive(asDouble(row.get("paneles")));
        }

        summary.put("totalCotizaciones", total);
        summary.put("montoTotal", montoTotal);
        summary.put("promedioInversion", total > 0 ? montoTotal / total : 0.0);
        summary.put("confirmadas", confirmadas);
        summary.put("convertidasProyecto", convertidasProyecto);
        summary.put("pendientes", Math.max(0, total - confirmadas));
        summary.put("consumoMensualTotal", consumoTotal);
        summary.put("potenciaTotalKwp", potenciaTotal);
        summary.put("ahorroMensualTotal", ahorroTotal);
        summary.put("panelesTotal", panelesTotal);
        return summary;
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

    private static double normalizePct(double value, double fallback) {
        double pct = value > 1.0 ? value / 100.0 : value;
        if (pct < 0 || pct > 0.30) {
            return fallback;
        }
        return pct;
    }

    private static boolean asBoolean(Object value, boolean fallback) {
        if (value == null) {
            return fallback;
        }
        if (value instanceof Boolean b) {
            return b;
        }
        String s = String.valueOf(value).trim().toLowerCase();
        if (s.isBlank()) {
            return fallback;
        }
        return switch (s) {
            case "true", "1", "si", "sí", "activo", "activa" -> true;
            case "false", "0", "no", "inactivo", "inactiva" -> false;
            default -> fallback;
        };
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

    private static int parsePlainInt(String value) {
        if (value == null || value.isBlank()) {
            return 0;
        }
        try {
            String txt = value.trim();
            if (txt.matches("^-?\\d+(\\.0+)?$")) {
                return (int) Double.parseDouble(txt);
            }
            return Integer.parseInt(txt);
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