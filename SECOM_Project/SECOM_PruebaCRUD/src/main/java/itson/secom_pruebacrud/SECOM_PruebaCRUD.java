/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package itson.secom_pruebacrud;
import itson.secom_domain.CalculoSolar;
import itson.secom_domain.Cliente;
import itson.secom_domain.ConsumoMensual;
import itson.secom_domain.Cotizacion;
import itson.secom_domain.DatosReciboCFE;
import itson.secom_domain.ResultadoCalculoCotizacion;
import itson.secom_domain.enumeradores.TipoTarifa;
import itson.secom_negocio.CotizacionService;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;
 
/**
 *
 * @author Acer
 */
public class SECOM_PruebaCRUD {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        CotizacionService servicio = new CotizacionService();
        int opcion = 0;
 
        System.out.println("=========================================");
        System.out.println("   SISTEMA SECOM - COTIZACIONES SOLAR   ");
        System.out.println("=========================================");
 
        do {
            System.out.println("\n--- MENU ---");
            System.out.println("1. Nueva cotizacion (captura manual)");
            System.out.println("2. Solo calcular (sin guardar)");
            System.out.println("3. Ver todas las cotizaciones");
            System.out.println("4. Ver detalle de una cotizacion");
            System.out.println("5. Cambiar estado de cotizacion");
            System.out.println("6. Salir");
            System.out.print("Opcion: ");
 
            try {
                opcion = Integer.parseInt(scanner.nextLine());
 
                switch (opcion) {
 
                    case 1 -> {
                        System.out.println("\n--- NUEVA COTIZACION ---");
 
                        // Datos del cliente (debe existir en BD)
                        System.out.print("ID del cliente (debe existir en BD): ");
                        int idCliente = Integer.parseInt(scanner.nextLine());
                        Cliente cliente = new Cliente();
                        cliente.setIdCliente(idCliente);
 
                        System.out.print("ID del vendedor (0 si no aplica): ");
                        int idVendedor = Integer.parseInt(scanner.nextLine());
 
                        // Datos del recibo
                        DatosReciboCFE datos = capturarRecibo(scanner);
 
                        // Calcular y guardar
                        Cotizacion cotizacion = servicio.calcularYGuardar(datos, cliente, idVendedor);
 
                        System.out.println("\n✓ Cotizacion guardada con ID: " + cotizacion.getId());
                        System.out.println("  Total: $" + String.format("%.2f", cotizacion.getTotal()));
                        System.out.println("  kWp instalados: " +
                            String.format("%.2f", cotizacion.getWattsInstalados() / 1000.0));
                    }
 
                    case 2 -> {
                        System.out.println("\n--- CALCULAR SIN GUARDAR ---");
                        DatosReciboCFE datos = capturarRecibo(scanner);
                        ResultadoCalculoCotizacion resultado = servicio.calcularSinGuardar(datos);
                        imprimirResultado(resultado);
                    }
 
                    case 3 -> {
                        System.out.println("\n--- COTIZACIONES REGISTRADAS ---");
                        List<Cotizacion> lista = servicio.listarTodas();
                        if (lista.isEmpty()) {
                            System.out.println("No hay cotizaciones registradas.");
                        } else {
                            System.out.printf("%-5s %-25s %-12s %-10s%n",
                                "ID", "Cliente", "Consumo kWh", "Estado");
                            System.out.println("-".repeat(60));
                            for (Cotizacion c : lista) {
                                System.out.printf("%-5d %-25s %-12.1f %-10s%n",
                                    c.getId(),
                                    c.getCliente() != null ? c.getCliente().getNombreComercial() : "-",
                                    c.getConsumoPromedioMensualKwh(),
                                    c.getEstado()
                                );
                            }
                        }
                    }
 
                    case 4 -> {
                        System.out.print("ID de la cotizacion: ");
                        int id = Integer.parseInt(scanner.nextLine());
 
                        Cotizacion c = servicio.obtenerPorId(id);
                        if (c == null) {
                            System.out.println("No se encontro la cotizacion.");
                        } else {
                            System.out.println("\n" + c.toString());
 
                            CalculoSolar cs = servicio.obtenerCalculoSolar(id);
                            if (cs != null) {
                                System.out.println("\nCalculo solar:");
                                System.out.println("  " + cs.toString());
                            }
 
                            List<ConsumoMensual> consumos = servicio.obtenerConsumos(id);
                            if (!consumos.isEmpty()) {
                                System.out.println("\nHistorial de consumo (" + consumos.size() + " registros):");
                                for (ConsumoMensual cm : consumos) {
                                    System.out.println("  " + cm.toString());
                                }
                            }
                        }
                    }
 
                    case 5 -> {
                        System.out.print("ID de la cotizacion: ");
                        int id = Integer.parseInt(scanner.nextLine());
                        System.out.println("Estados: 1=BORRADOR 2=COTIZADA 3=ACEPTADA 4=RECHAZADA 5=FINALIZADA");
                        System.out.print("Nuevo estado (1-5): ");
                        int est = Integer.parseInt(scanner.nextLine());
                        itson.secom_domain.enumeradores.EstadoCotizacion[] estados =
                            itson.secom_domain.enumeradores.EstadoCotizacion.values();
                        if (est < 1 || est > estados.length) {
                            System.out.println("Estado invalido.");
                        } else {
                            servicio.cambiarEstado(id, estados[est - 1], 1);
                            System.out.println("Estado actualizado a: " + estados[est - 1]);
                        }
                    }
 
                    case 6 -> System.out.println("Saliendo... hasta pronto.");
 
                    default -> System.out.println("Opcion no valida.");
                }
 
            } catch (NumberFormatException e) {
                System.out.println("Error: ingresa un numero valido.");
            } catch (Exception e) {
                System.out.println("Error: " + e.getMessage());
            }
 
        } while (opcion != 6);
 
        scanner.close();
    }
 
    // -------------------------------------------------------
    // Captura datos del recibo manualmente
    // -------------------------------------------------------
    private static DatosReciboCFE capturarRecibo(Scanner scanner) {
        DatosReciboCFE datos = new DatosReciboCFE();
 
        System.out.println("\nTipos de tarifa:");
        System.out.println("  1. Domestica Mensual");
        System.out.println("  2. Domestica Bimestral");
        System.out.println("  3. PDBT Mensual");
        System.out.println("  4. PDBT Bimestral");
        System.out.println("  5. GDMTH");
        System.out.println("  6. GDMTO");
        System.out.print("Selecciona (1-6): ");
        int tipoNum = Integer.parseInt(scanner.nextLine());
        TipoTarifa[] tipos = TipoTarifa.values();
        datos.setTipoTarifa(tipoNum >= 1 && tipoNum <= 6 ? tipos[tipoNum - 1] : TipoTarifa.DOMESTICA_MENSUAL);
 
        System.out.print("Nombre del titular: ");
        datos.setNombre(scanner.nextLine());
 
        System.out.print("No. de servicio (RPU): ");
        datos.setNoServicio(scanner.nextLine());
 
        System.out.print("Codigo de tarifa (ej. 1C, DAC, PDBT): ");
        datos.setTarifa(scanner.nextLine());
 
        System.out.print("Consumo del periodo actual (kWh): ");
        datos.setConsumoActualKwh(Double.parseDouble(scanner.nextLine()));
 
        System.out.print("Total a pagar del periodo ($): ");
        datos.setPagoActual(Double.parseDouble(scanner.nextLine()));
 
        System.out.print("Cuantos periodos historicos tienes? (0 para omitir): ");
        int numHist = Integer.parseInt(scanner.nextLine());
 
        if (numHist > 0) {
            double[] consumos = new double[numHist + 1];
            double[] pagos    = new double[numHist + 1];
            consumos[numHist] = datos.getConsumoActualKwh();
            pagos[numHist]    = datos.getPagoActual();
 
            for (int i = numHist - 1; i >= 0; i--) {
                System.out.print("  Periodo " + (numHist - i) + " atras - Consumo (kWh): ");
                consumos[i] = Double.parseDouble(scanner.nextLine());
                System.out.print("  Periodo " + (numHist - i) + " atras - Pago ($): ");
                pagos[i] = Double.parseDouble(scanner.nextLine());
            }
 
            List<Double> listaConsumos = new java.util.ArrayList<>();
            List<Double> listaPagos    = new java.util.ArrayList<>();
            for (double v : consumos) listaConsumos.add(v);
            for (double v : pagos)    listaPagos.add(v);
 
            datos.setConsumoHistoricos(listaConsumos);
            datos.setPagosHistoricos(listaPagos);
        } else {
            datos.setConsumoHistoricos(Arrays.asList(datos.getConsumoActualKwh()));
            datos.setPagosHistoricos(Arrays.asList(datos.getPagoActual()));
        }
 
        System.out.print("Costo de suministro/cargo fijo del recibo (0 si no lo tienes): ");
        datos.setCostoSuministro(Double.parseDouble(scanner.nextLine()));
        if (datos.getCostoSuministro() > 0) {
            System.out.print("IVA % (ej. 16): ");
            datos.setIvaPorcentaje(Double.parseDouble(scanner.nextLine()));
            System.out.print("DAP ($): ");
            datos.setCostoDAP(Double.parseDouble(scanner.nextLine()));
        }
 
        return datos;
    }
 
    // -------------------------------------------------------
    // Imprime el resultado del calculo
    // -------------------------------------------------------
    private static void imprimirResultado(ResultadoCalculoCotizacion r) {
        System.out.println("\n======= RESULTADO =======");
        System.out.printf("Cliente      : %s%n", r.getNombreCliente());
        System.out.printf("Tarifa       : %s (%s)%n", r.getTarifa(), r.getTipoTarifa());
        System.out.printf("Consumo/mes  : %.1f kWh%n", r.getConsumoPromedioMensualKwh());
        System.out.printf("Pago prom CFE: $%.2f%n", r.getPagoPromedioCFE());
        System.out.printf("Ahorro/mes   : $%.2f%n", r.getAhorroMensualEstimado());
        System.out.println("-------------------------");
        System.out.printf("Paneles      : %d x 550W%n", r.getNumeroPaneles());
        System.out.printf("Potencia     : %.2f kWp%n", r.getPotenciaInstaladaKwp());
        System.out.printf("Gen. anual   : %.1f kWh%n", r.getGeneracioAnualEstimadaKwh());
        System.out.printf("Retorno      : %.1f anios%n", r.getRetornoInversion());
        System.out.println("-------------------------");
        System.out.printf("CO2 (25 anios): %.1f ton%n", r.getCo2EvitadoToneladas25años());
        System.out.printf("Arboles equiv: %d%n", r.getArbolesEquivalentes25Años());
        System.out.printf("Costo c/IVA  : $%d%n", r.getCostoProyectoConIva());
        System.out.println("=========================");
    }
}