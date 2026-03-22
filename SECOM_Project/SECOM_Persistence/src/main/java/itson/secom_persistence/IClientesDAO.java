/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Interface.java to edit this template
 */
package itson.secom_persistence;

import itson.secom_domain.Cliente;
import itson.secom_persistence.excepciones.PersistenciaException;
import java.util.List;

/**
 *
 * @author PC
 */
public interface IClientesDAO {

    /**
     * Obtiene una lista de todos los clientes de la base de datos.
     *
     * @return Lista con todos los clientes de la base de datos.
     * @throws PersistenciaException en caso de que exista algun error.
     *
     */
    public List<Cliente> obtenerClientes() throws PersistenciaException;

    /**
     * Obtiene un cliente por su id.
     *
     * @param id ID del cliente a obtener.
     * @return Datos del cliente.
     * @throws PersistenciaException en caso de que exista algun error
     */
    public Cliente obtenerCliente(int id) throws PersistenciaException;
}
