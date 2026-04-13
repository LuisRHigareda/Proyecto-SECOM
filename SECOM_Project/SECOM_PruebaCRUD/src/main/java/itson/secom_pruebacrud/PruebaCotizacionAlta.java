package itson.secom_pruebacrud;

import itson.secom_domain.DatosReciboCFE;
import itson.secom_domain.ResultadoCalculoCotizacion;
import itson.secom_domain.enumeradores.TipoTarifa;
import itson.secom_negocio.CotizacionService;
import itson.secom_persistence.conexionFactory.DAOFactory;
import java.text.NumberFormat;
import java.util.List;
import java.util.Locale;
import java.util.Scanner;

public class PruebaCotizacionAlta {

    public static void main(String[] args) {

        Scanner sc = new Scanner(System.in);

        DAOFactory factory = new DAOFactory(false);
        CotizacionService service = new CotizacionService(factory);

        int opcion;

        do {
            System.out.println("\n=================================");
            System.out.println("     SISTEMA DE COTIZACIONES");
            System.out.println("=================================");
            System.out.println("1. Crear cotización");
            System.out.println("0. Salir");
            System.out.print("Seleccione una opción: ");

            opcion = sc.nextInt();
            sc.nextLine(); // Limpiar buffer

            switch (opcion) {

                case 1:
                    try {

                        System.out.println("\n--- NUEVA COTIZACIÓN ---");
                        System.out.println("Formato:");
                        System.out.println("Nombre | consumos | tipoPeriodo | ciudad | tarifa");

                        String input = sc.nextLine();
                        String[] partes = input.split("\\|");

                        if (partes.length < 5) {
                            System.out.println("Formato incorrecto.");
                            break;
                        }

                        String nombre = partes[0].trim();
                        String consumos = partes[1].trim();
                        String tipo = partes[2].trim();
                        String ciudad = partes[3].trim();
                        String tarifa = partes[4].trim();

                        // Convertir String a Enum TipoTarifa
                        TipoTarifa tipoTarifa = TipoTarifa.valueOf(tarifa.toUpperCase());

                        // Crear objeto de dominio
                        DatosReciboCFE datos = new DatosReciboCFE(
                                nombre,
                                consumos,
                                tipo,
                                ciudad,
                                tipoTarifa
                        );

                        // =========================
                        // PRE-CÁLCULO
                        // =========================
                        ResultadoCalculoCotizacion previo =
                                service.calcularCotizacionConPaquete(datos, 1);

                        double consumoMensual = previo.getConsumoPromedioMensualKwh();

                        // Calcular consumo promedio diario en kWh y en Watts
                        double consumoDiarioKwh = consumoMensual / 30.0;
                        double consumoDiarioWatts = consumoDiarioKwh * 1000;

                        // Formateo para mostrar separador de miles
                        NumberFormat formato = NumberFormat.getInstance(new Locale("es", "MX"));
                        formato.setMaximumFractionDigits(0);
                        formato.setMinimumFractionDigits(0);

                        System.out.println("\n--- PRE-CÁLCULO ---");
                        System.out.println("Consumo mensual: "
                                + String.format("%.2f", consumoMensual) + " kWh");
                        System.out.println("Consumo promedio diario de energía: "
                                + formato.format(consumoDiarioWatts) + " W");
                        System.out.println("kW requeridos: "
                                + String.format("%.2f", previo.getKwpRequerido()));

                        // =========================
                        // SELECCIÓN DE PAQUETE
                        // =========================
                        System.out.print("\nSeleccione paquete (ID): ");
                        int paqueteId = sc.nextInt();
                        sc.nextLine();

                        ResultadoCalculoCotizacion r =
                                service.calcularCotizacionConPaquete(datos, paqueteId);

                        // =========================
                        // RESULTADO FINAL
                        // =========================
                        System.out.println("\n========== RESULTADO ==========");

                        System.out.println("Cliente: " + r.getNombreCliente());

                        System.out.println("\n--- CONSUMO ---");
                        System.out.println("Mensual: "
                                + String.format("%.2f", r.getConsumoPromedioMensualKwh()) + " kWh");
                        System.out.println("Diario: "
                                + String.format("%.2f", r.getConsumoPromedioDiarioKwh()) + " kWh");

                        System.out.println("\n--- SISTEMA ---");
                        System.out.println("kW requeridos: "
                                + String.format("%.2f", r.getKwpRequerido()));
                        System.out.println("Watts instalados: "
                                + String.format("%.2f", r.getWattsInstalados()));

                        System.out.println("\n--- PRODUCCIÓN ---");
                        System.out.println("Producción diaria: "
                                + String.format("%.2f", r.getProduccionDiariaEstimada()));
                        System.out.println("Cobertura: "
                                + String.format("%.2f", r.getPorcentajeCobertura()) + " %");

                        System.out.println("\n--- COSTOS ---");
                        System.out.println("Subtotal: $" + String.format("%.2f", r.getSubtotal()));
                        System.out.println("IVA: $" + String.format("%.2f", r.getIva()));
                        System.out.println("TOTAL: $" + String.format("%.2f", r.getTotal()));

                        System.out.println("=================================\n");

                    } catch (IllegalArgumentException e) {
                        System.out.println("Tarifa inválida. Verifique el valor ingresado.");
                    } catch (Exception e) {
                        System.out.println("Error en cotización: " + e.getMessage());
                        e.printStackTrace();
                    }
                    break;

                case 0:
                    System.out.println("Saliendo...");
                    break;

                default:
                    System.out.println("Opción inválida");
            }

        } while (opcion != 0);

        sc.close();
    }
}

/*
 * === CASO DE PRUEBA 1 - Paquete No. 1 ===
 *
 * Cliente: BORQUEZ MANZ JUAN HUMBERTO
 * Histórico de consumo (kWh): 309, 599, 1184, 1522, 1653, 1615, 1132, 763, 307, 305, 310, 302
 * Tipo: MENSUAL | Ciudad Obregón | DOMESTICA_MENSUAL
 * copia esto -->" BORQUEZ MANZ JUAN HUMBERTO | 309,599,1184,1522,1653,1615,1132,763,307,305,310,302 | MENSUAL | Ciudad Obregon | DOMESTICA_MENSUAL  "
 * RESULTADO ESPERADO:
 *   - Producción diaria de energía: 31,861 kWh
 *   - % Producción vs Consumo: 115%
 *   - Consumo promedio diario de energía: 27,598 kWh
 *   - Watts instalados: 6,200 W
 */

/*
 * === CASO DE PRUEBA 2 ===
 *  * copia esto -->"VALENZUELA CARRILLO LUIS CARLOS | 247,461,882,1021,1263,1318,920,621,310,245,238,256 | MENSUAL | Ciudad Obregon | DOMESTICA_MENSUAL"
*Cliente:VALENZUELA CARRILLO LUIS CARLOS
*Tipo: MENSUAL | Ciudad Obregón | DOMESTICA_MENSUAL
 * RESULTADO ESPERADO:
 *   - Producción diaria de energía: 63,722 kWh
 *   - % Producción vs Consumo: 105%
 *   - Consumo promedio diario de energía: 60,688 kWh
 *   - Watts instalados: 12,400 W
 */

