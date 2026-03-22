/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package itson.secom_pruebacrud;
import itson.secom_domain.Cotizacion;
import itson.secom_domain.enumeradores.EstadoCotizacion;
import itson.secom_negocio.CotizacionService;

import java.time.LocalDateTime;
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
        System.out.println("   SISTEMA SECOM - COTIZACIONES   ");
        System.out.println("=========================================");

        do {
            System.out.println("\n--- MENU PRINCIPAL ---");
            System.out.println("1. Crear nueva cotizacion");
            System.out.println("2. Ver todas las cotizaciones");
            System.out.println("3. Buscar cotizacion por ID");
            System.out.println("4. Actualizar una cotizacion");
            System.out.println("5. Eliminar una cotizacion");
            System.out.println("6. Salir");
            System.out.print("Elige una opcion: ");

            try {
                opcion = Integer.parseInt(scanner.nextLine());

                switch (opcion) {
                    case 1:
                        System.out.println("\n--- 1. CREAR COTIZACION ---");
                        Cotizacion nueva = new Cotizacion();
                        
                        System.out.print("Ingresa el consumo promedio mensual (kWh): ");
                        nueva.setConsumoPromedioMensualKwh(Double.parseDouble(scanner.nextLine()));

                        System.out.print("Ingresa el total de la cotizacion ($): ");
                        nueva.setTotal(Double.parseDouble(scanner.nextLine()));

                        nueva.setFecha(LocalDateTime.now());
                        nueva.setEstado(EstadoCotizacion.BORRADOR);

                        servicio.guardarCotizacion(nueva);
                        System.out.println(" ¡Cotizacion guardada exitosamente!");
                        break;

                    case 2:
                        System.out.println("\n--- 2. VER TODAS LAS COTIZACIONES ---");
                        List<Cotizacion> lista = servicio.obtenerTodas();
                        if (lista.isEmpty()) {
                            System.out.println("No hay cotizaciones registradas.");
                        } else {
                            for (Cotizacion c : lista) {
                                System.out.println(c.toString());
                            }
                        }
                        break;

                    case 3:
                        System.out.println("\n--- 3. BUSCAR POR ID ---");
                        System.out.print("Ingresa el ID de la cotizacion: ");
                        int idBuscar = Integer.parseInt(scanner.nextLine());
                        
                        Cotizacion encontrada = servicio.obtenerPorId(idBuscar);
                        if (encontrada != null) {
                            System.out.println(" Cotizacion encontrada: " + encontrada.toString());
                        } else {
                            System.out.println(" No se encontro ninguna cotizacion con el ID: " + idBuscar);
                        }
                        break;

                    case 4:
                        System.out.println("\n--- 4. ACTUALIZAR COTIZACION ---");
                        System.out.print("Ingresa el ID de la cotizacion a actualizar: ");
                        int idActualizar = Integer.parseInt(scanner.nextLine());
                        
                        Cotizacion cotizacionAEditar = servicio.obtenerPorId(idActualizar);
                        if (cotizacionAEditar != null) {
                            System.out.println("Datos actuales: " + cotizacionAEditar.toString());
                            
                            System.out.print("Nuevo consumo promedio mensual (kWh): ");
                            cotizacionAEditar.setConsumoPromedioMensualKwh(Double.parseDouble(scanner.nextLine()));
                            
                            System.out.print("Nuevo total ($): ");
                            cotizacionAEditar.setTotal(Double.parseDouble(scanner.nextLine()));
                            
                            servicio.actualizarCotizacion(cotizacionAEditar);
                            System.out.println(" ¡Cotizacion actualizada correctamente!");
                        } else {
                            System.out.println(" No existe una cotizacion con ese ID.");
                        }
                        break;

                    case 5: 
                        System.out.println("\n--- 5. ELIMINAR COTIZACION ---");
                        System.out.print("Ingresa el ID de la cotizacion a eliminar: ");
                        int idEliminar = Integer.parseInt(scanner.nextLine());
                        
                        servicio.eliminarCotizacion(idEliminar);
                        System.out.println("¡Cotizacion eliminada correctamente!");
                        break;

                    case 6:
                        System.out.println("Saliendo del sistema... ¡Hasta pronto!");
                        break;

                    default:
                        System.out.println("Opcion no valida. Intenta de nuevo.");
                }
            } catch (NumberFormatException e) {
                System.out.println("Error: Por favor, ingresa un numero valido.");
            } catch (Exception e) {
                System.out.println("Ocurrio un error: " + e.getMessage());
            }

        } while (opcion != 6);

        scanner.close();
    }
}
