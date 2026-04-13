package itson.secom_presentacion.util;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.reflect.TypeToken;
import jakarta.servlet.http.HttpServletRequest;
import java.io.IOException;
import java.io.Reader;
import java.lang.reflect.Type;
import java.util.LinkedHashMap;
import java.util.Map;

public final class RequestJson {

    private static final Gson GSON = new GsonBuilder()
            .serializeNulls()
            .create();

    private static final Type MAP_TYPE = new TypeToken<Map<String, Object>>() {}.getType();

    private RequestJson() {
    }

    public static Map<String, Object> readMap(HttpServletRequest request) throws IOException {
        try (Reader reader = request.getReader()) {
            Map<String, Object> body = GSON.fromJson(reader, MAP_TYPE);
            return body != null ? body : new LinkedHashMap<>();
        }
    }

    public static String toJson(Object value) {
        return GSON.toJson(value);
    }
}