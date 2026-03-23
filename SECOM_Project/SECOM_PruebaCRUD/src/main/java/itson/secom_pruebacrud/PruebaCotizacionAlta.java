/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package itson.secom_pruebacrud;

/**
 *
 * @author Arell
 */

import itson.secom_domain.DatosReciboCFE;
import itson.secom_domain.ResultadoCalculoCotizacion;
import itson.secom_domain.ResultadoCotizacion;
import itson.secom_domain.enumeradores.TipoTarifa;
import itson.secom_negocio.CotizacionService;
import itson.secom_persistence.conexionFactory.DAOFactory;
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
        sc.nextLine(); // limpiar buffer

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

                    // 🔥 Crear objeto SIN tarifa en constructor
                    TipoTarifa tipoTarifa = TipoTarifa.valueOf(tarifa.trim().toUpperCase());

DatosReciboCFE datos = new DatosReciboCFE(
        nombre,
        consumos,
        tipo,
        ciudad,
        tipoTarifa
);

                    // 🔥 Setear tarifa desde string → enum
                   // datos.setTipoTarifaDesdeString(tarifa);

                    // 🔥 PRE-CÁLCULO (usa paquete 1 como base)
                    ResultadoCalculoCotizacion previo =
                            service.calcularCotizacionConPaquete(datos, 1);

                    System.out.println("\n--- PRE-CÁLCULO ---");
                    System.out.println("Consumo mensual: "
                            + previo.getConsumoPromedioMensualKwh() + " kWh");
                    System.out.println("kW requeridos: "
                            + previo.getKwpRequerido());

                    // 🔥 Selección de paquete
                    System.out.print("\nSeleccione paquete (ID): ");
                    int paqueteId = sc.nextInt();
                    sc.nextLine();

// 👇 AGREGA ESTO
                    CotizacionService service2 = new CotizacionService(new DAOFactory(false));

// 👇 CAMBIA A ESTO
                    ResultadoCalculoCotizacion r
                            = service2.calcularCotizacionConPaquete(datos, paqueteId);

                    // 🔥 RESULTADO FINAL
                    System.out.println("\n========== RESULTADO ==========");

                    System.out.println("Cliente: " + r.getNombreCliente());

                    System.out.println("\n--- CONSUMO ---");
                    System.out.println("Mensual: "
                            + r.getConsumoPromedioMensualKwh() + " kWh");
                    System.out.println("Diario: "
                            + r.getConsumoPromedioDiarioKwh() + " kWh");

                    System.out.println("\n--- SISTEMA ---");
                    System.out.println("kW requeridos: "
                            + r.getKwpRequerido());
                    System.out.println("Watts instalados: "
                            + r.getWattsInstalados());

                    System.out.println("\n--- PRODUCCIÓN ---");
                    System.out.println("Producción diaria: "
                            + r.getProduccionDiariaEstimada());
                    System.out.println("Cobertura: "
                            + r.getPorcentajeCobertura() + " %");

                    System.out.println("\n--- COSTOS ---");
                    System.out.println("Subtotal: $" + r.getSubtotal());
                    System.out.println("IVA: $" + r.getIva());
                    System.out.println("TOTAL: $" + r.getTotal());

                    System.out.println("=================================\n");

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

