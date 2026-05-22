package itson.secom_presentacion.servlets;

import itson.secom_presentacion.util.JsonResponse;
import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.MultipartConfig;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.Part;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

@WebServlet(name = "ReceiptPdfTextServlet", urlPatterns = {"/api/receipts/pdf-text"})
@MultipartConfig(
        fileSizeThreshold = 1024 * 1024,
        maxFileSize = 20L * 1024L * 1024L,
        maxRequestSize = 25L * 1024L * 1024L
)
public class ReceiptPdfTextServlet extends HttpServlet {

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws IOException {
        try {
            Part filePart = request.getPart("file");

            if (filePart == null || filePart.getSize() <= 0) {
                JsonResponse.send(response, HttpServletResponse.SC_BAD_REQUEST, Map.of(
                        "ok", false,
                        "message", "No se recibió ningún archivo PDF."
                ));
                return;
            }

            String submittedName = String.valueOf(filePart.getSubmittedFileName() == null ? "" : filePart.getSubmittedFileName()).toLowerCase();
            String contentType = String.valueOf(filePart.getContentType() == null ? "" : filePart.getContentType()).toLowerCase();

            if (!submittedName.endsWith(".pdf") && !contentType.contains("pdf")) {
                JsonResponse.send(response, HttpServletResponse.SC_BAD_REQUEST, Map.of(
                        "ok", false,
                        "message", "El archivo enviado no parece ser un PDF."
                ));
                return;
            }

            byte[] pdfBytes = filePart.getInputStream().readAllBytes();

            try (PDDocument document = Loader.loadPDF(pdfBytes)) {
                if (document.isEncrypted()) {
                    JsonResponse.send(response, HttpServletResponse.SC_BAD_REQUEST, Map.of(
                            "ok", false,
                            "message", "El PDF está protegido o cifrado y no puede leerse automáticamente."
                    ));
                    return;
                }

                PDFTextStripper stripper = new PDFTextStripper();
                stripper.setSortByPosition(true);

                List<String> pageTexts = new ArrayList<>();

                for (int page = 1; page <= document.getNumberOfPages(); page++) {
                    stripper.setStartPage(page);
                    stripper.setEndPage(page);
                    pageTexts.add(normalizeText(stripper.getText(document)));
                }

                String fullText = normalizeText(String.join("\n\n", pageTexts));

                Map<String, Object> data = new LinkedHashMap<>();
                data.put("source", "PDFBox");
                data.put("pages", document.getNumberOfPages());
                data.put("text", fullText);
                data.put("pageTexts", pageTexts);
                data.put("textLength", fullText.length());
                data.put("hasText", fullText.replaceAll("\\s+", "").length() > 60);

                Map<String, Object> body = new LinkedHashMap<>();
                body.put("ok", true);
                body.put("data", data);

                JsonResponse.send(response, HttpServletResponse.SC_OK, body);
            }
        } catch (ServletException ex) {
            JsonResponse.send(response, HttpServletResponse.SC_BAD_REQUEST, Map.of(
                    "ok", false,
                    "message", "No se pudo procesar el archivo enviado.",
                    "type", ex.getClass().getName()
            ));
        } catch (Exception ex) {
            JsonResponse.send(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR, Map.of(
                    "ok", false,
                    "message", ex.getMessage() == null ? "Error al leer el PDF con PDFBox." : ex.getMessage(),
                    "type", ex.getClass().getName()
            ));
        }
    }

    private String normalizeText(String text) {
        return String.valueOf(text == null ? "" : text)
                .replace("\r", "\n")
                .replaceAll("[ \\t]+", " ")
                .replaceAll("\\n{3,}", "\n\n")
                .trim();
    }
}
