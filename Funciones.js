// ============================================================
// GOOGLE APPS SCRIPT - BACKEND MEJORADO
// Sistema de Inventario TODO SOBRE RUEDAS
// VERSIÓN MEJORADA: Preserva fórmulas al agregar datos
// ============================================================

// Configuración - ID de tu Google Sheet
// IMPORTANTE: Reemplazar esto con el ID real de tu archivo
const SPREADSHEET_ID = '1xMqZV5HlYMEdX0FDNi3GubAuw8vaXchV-AB0uQPf1xA'; // Se obtiene de la URL del Google Sheet

// ============================================================
// CONFIGURACIÓN DE SEGURIDAD - USUARIOS AUTORIZADOS
// ============================================================
// IMPORTANTE: Reemplaza estos emails con tus correos reales
const USUARIOS_AUTORIZADOS = [
  'tu-correo-principal@gmail.com',    // Tu correo principal
  'tu-segundo-correo@gmail.com'       // Tu segundo correo
];

/**
 * Verifica si el usuario actual está autorizado para modificar datos
 */
function verificarPermisoEdicion() {
  const userEmail = Session.getActiveUser().getEmail();
  
  // Si no hay email (modo de prueba en editor), denegar acceso
  if (!userEmail || userEmail === '') {
    return {
      autorizado: false,
      email: 'Usuario anónimo',
      mensaje: 'Debes iniciar sesión con una cuenta autorizada para realizar cambios'
    };
  }
  
  // Verificar si el email está en la lista de autorizados
  const autorizado = USUARIOS_AUTORIZADOS.some(function(emailAutorizado) {
    return userEmail.toLowerCase() === emailAutorizado.toLowerCase();
  });
  
  return {
    autorizado: autorizado,
    email: userEmail,
    mensaje: autorizado ? 'Autorizado' : 'Tu cuenta no tiene permisos para modificar datos'
  };
}

// Función para servir la interfaz web
function doGet() {
  // Todos pueden ver la interfaz, pero solo usuarios autorizados pueden guardar
  return HtmlService.createHtmlOutputFromFile('interfaz_web')
    .setTitle('TODO SOBRE RUEDAS - Sistema de Inventario')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// FUNCIONES PARA COMPRAS
// ============================================================

function guardarCompra(datos) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName('HISTORIAL COMPRAS');
    
    // Obtener la última fila con datos
    const ultimaFila = hoja.getLastRow();
    let nuevoID = 1;
    let filaDestino = 2; // Fila 2 es la primera después del encabezado
    
    // Si hay datos (más allá del encabezado)
    if (ultimaFila >= 2) {
      // Buscar el último ID válido
      for (let i = ultimaFila; i >= 2; i--) {
    // VERIFICAR PERMISOS DE EDICIÓN
    const permisos = verificarPermisoEdicion();
    if (!permisos.autorizado) {
      return {
        success: false,
        message: '🚫 ACCESO DENEGADO: ' + permisos.mensaje + '\n\nUsuario: ' + permisos.email + '\n\nContacta al administrador si necesitas acceso.'
      };
    }
    
        const idCelda = hoja.getRange(i, 1).getValue();
        if (idCelda && !isNaN(idCelda) && idCelda !== '' && idCelda !== 0) {
          nuevoID = parseInt(idCelda) + 1;
          filaDestino = i + 1; // La siguiente fila después del último registro
          break;
        }
      }
    }
    
    // NO insertar filas nuevas, simplemente escribir en la fila destino
    // Esto evita problemas con fórmulas y desplazamientos
    
    // Escribir todos los datos directamente
    hoja.getRange(filaDestino, 1).setValue(nuevoID);                    // A: ID
    hoja.getRange(filaDestino, 2).setValue(new Date(datos.fecha));      // B: FECHA COMPRA
    hoja.getRange(filaDestino, 3).setValue(datos.marca);                // C: MARCA
    hoja.getRange(filaDestino, 4).setValue(datos.medida);               // D: MEDIDA
    hoja.getRange(filaDestino, 5).setValue(datos.modelo);               // E: MODELO
    hoja.getRange(filaDestino, 6).setValue(parseInt(datos.aro));        // F: ARO
    hoja.getRange(filaDestino, 7).setValue(parseInt(datos.cantidad));   // G: CANTIDAD
    hoja.getRange(filaDestino, 8).setValue(parseFloat(datos.precio));   // H: PRECIO COMPRA
    
    // Fórmulas para PRECIO +35% y PRECIO VENTA
    // Calcular directamente sin fórmulas para evitar problemas de idioma
    const precioCompra = parseFloat(datos.precio);
    const precioMas35 = precioCompra * 1.35;
    hoja.getRange(filaDestino, 9).setValue(precioMas35);  // I: PRECIO +35%
    hoja.getRange(filaDestino, 10).setValue(precioMas35); // J: PRECIO VENTA
    
    hoja.getRange(filaDestino, 11).setValue(datos.proveedor || '');     // K: PROVEEDOR
    hoja.getRange(filaDestino, 12).setValue(datos.factura || '');       // L: N° FACTURA
    hoja.getRange(filaDestino, 13).setValue(datos.observaciones || ''); // M: OBSERVACIONES
    
    // Agregar al inventario si no existe
    actualizarInventario(datos);
    
    return {
      success: true,
      message: 'Compra registrada exitosamente',
      id: nuevoID
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error al guardar: ' + error.toString()
    };
  }
}

function actualizarInventario(datos) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaInventario = ss.getSheetByName('INVENTARIO COMPLETO');
  
  // Obtener todos los datos del inventario
  const ultimaFilaInv = hojaInventario.getLastRow();
  let productoExiste = false;
  
  // Buscar si el producto ya existe (solo si hay datos)
  if (ultimaFilaInv >= 2) {
    const datosInventario = hojaInventario.getRange(2, 1, ultimaFilaInv - 1, 6).getValues();
    
    for (let i = 0; i < datosInventario.length; i++) {
      const fila = datosInventario[i];
      // Verificar por ARO, marca, medida y modelo (columnas B, C, D, E = índices 1, 2, 3, 4)
      const aroFila = parseInt(fila[1]) || 0;
      const aroNuevo = parseInt(datos.aro) || 0;
      if (aroFila === aroNuevo && fila[2] === datos.marca && fila[3] === datos.medida && fila[4] === datos.modelo) {
        productoExiste = true;
        break;
      }
    }
  }
  
  // Si no existe, agregarlo
  if (!productoExiste) {
    // Determinar la fila destino
    let filaDestino = 2; // Primera fila después del encabezado
    
    if (ultimaFilaInv >= 2) {
      // Buscar la primera fila vacía o usar la siguiente después de la última con datos
      for (let i = 2; i <= ultimaFilaInv + 1; i++) {
        const codigo = hojaInventario.getRange(i, 1).getValue();
        if (!codigo || codigo === '' || codigo === 0) {
          filaDestino = i;
          break;
        }
        filaDestino = i + 1;
      }
    }
    
    // Generar código único
    let numeroSecuencial = 1;
    if (ultimaFilaInv >= 2) {
      const aros = hojaInventario.getRange(2, 2, ultimaFilaInv - 1, 1).getValues();
      for (let i = 0; i < aros.length; i++) {
        if (aros[i][0] == datos.aro) {
          numeroSecuencial++;
        }
      }
    }
    const codigo = `NEU-${datos.aro}-${String(numeroSecuencial).padStart(3, '0')}`;
    
    // Escribir datos directamente (sin insertar filas)
    hojaInventario.getRange(filaDestino, 1).setValue(codigo);              // A: CÓDIGO
    hojaInventario.getRange(filaDestino, 2).setValue(parseInt(datos.aro)); // B: ARO
    hojaInventario.getRange(filaDestino, 3).setValue(datos.marca);         // C: MARCA
    hojaInventario.getRange(filaDestino, 4).setValue(datos.medida);        // D: MEDIDA
    hojaInventario.getRange(filaDestino, 5).setValue(datos.modelo);        // E: MODELO
    
    // Calcular valores directamente sin fórmulas para evitar problemas de idioma
    // F: STOCK ACTUAL (inicialmente = cantidad comprada)
    hojaInventario.getRange(filaDestino, 6).setValue(parseInt(datos.cantidad));
    
    // G: PRECIO COMPRA PROMEDIO (inicialmente = precio de esta compra)
    const precioCompraInv = parseFloat(datos.precio);
    hojaInventario.getRange(filaDestino, 7).setValue(precioCompraInv);
    
    // H: PRECIO +35%
    const precioMas35Inv = precioCompraInv * 1.35;
    hojaInventario.getRange(filaDestino, 8).setValue(precioMas35Inv);
    
    // I: PRECIO VENTA
    hojaInventario.getRange(filaDestino, 9).setValue(precioMas35Inv);
    
    // J: ÚLTIMA COMPRA
    hojaInventario.getRange(filaDestino, 10).setValue(new Date(datos.fecha));
    
    // K: ÚLTIMA CANTIDAD
    hojaInventario.getRange(filaDestino, 11).setValue(parseInt(datos.cantidad));
  } else {
    // Si el producto ya existe, actualizar el inventario
    actualizarInventarioExistente(datos);
  }
}

/**
 * Actualiza el inventario cuando se agrega una compra de un producto existente
 */
function actualizarInventarioExistente(datos) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaInventario = ss.getSheetByName('INVENTARIO COMPLETO');
  const ultimaFilaInv = hojaInventario.getLastRow();
  
  if (ultimaFilaInv < 2) return;
  
  const datosInventario = hojaInventario.getRange(2, 1, ultimaFilaInv - 1, 11).getValues();
  
  for (let i = 0; i < datosInventario.length; i++) {
    const fila = datosInventario[i];
    // Verificar por ARO, marca, medida y modelo (columnas B, C, D, E = índices 1, 2, 3, 4)
    const aroFila = parseInt(fila[1]) || 0;
    const aroNuevo = parseInt(datos.aro) || 0;
    if (aroFila === aroNuevo && fila[2] === datos.marca && fila[3] === datos.medida && fila[4] === datos.modelo) {
      const filaReal = i + 2; // +2 porque empezamos en fila 2 y el índice es 0
      
      // F: Actualizar STOCK (sumar cantidad)
      const stockActual = fila[5] || 0;
      const nuevoStock = stockActual + parseInt(datos.cantidad);
      hojaInventario.getRange(filaReal, 6).setValue(nuevoStock);
      
      // G: Actualizar PRECIO COMPRA PROMEDIO (promedio ponderado)
      const precioActual = fila[6] || 0;
      const cantidadNueva = parseInt(datos.cantidad);
      const precioNuevo = parseFloat(datos.precio);
      let nuevoPrecioPromedio;
      
      if (stockActual > 0 && precioActual > 0) {
        // Promedio ponderado
        nuevoPrecioPromedio = ((precioActual * stockActual) + (precioNuevo * cantidadNueva)) / nuevoStock;
      } else {
        nuevoPrecioPromedio = precioNuevo;
      }
      hojaInventario.getRange(filaReal, 7).setValue(nuevoPrecioPromedio);
      
      // H: PRECIO +35%
      const precioMas35 = nuevoPrecioPromedio * 1.35;
      hojaInventario.getRange(filaReal, 8).setValue(precioMas35);
      
      // I: PRECIO VENTA
      hojaInventario.getRange(filaReal, 9).setValue(precioMas35);
      
      // J: ÚLTIMA COMPRA
      hojaInventario.getRange(filaReal, 10).setValue(new Date(datos.fecha));
      
      // K: ÚLTIMA CANTIDAD
      hojaInventario.getRange(filaReal, 11).setValue(cantidadNueva);
      
      break;
    }
  }// VERIFICAR PERMISOS DE EDICIÓN
    const permisos = verificarPermisoEdicion();
    if (!permisos.autorizado) {
      return {
        success: false,
        message: '🚫 ACCESO DENEGADO: ' + permisos.mensaje + '\n\nUsuario: ' + permisos.email + '\n\nContacta al administrador si necesitas acceso.'
      };
    }
    
    
}

// ============================================================
// FUNCIONES PARA VENTAS
// ============================================================

function guardarVenta(datos) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName('REGISTRO VENTAS');
    
    // Obtener la última fila con datos
    const ultimaFila = hoja.getLastRow();
    let nuevoID = 1;
    let filaDestino = 2; // Fila 2 es la primera después del encabezado
    
    // Si hay datos (más allá del encabezado)
    if (ultimaFila >= 2) {
      // Buscar el último ID válido
      for (let i = ultimaFila; i >= 2; i--) {
        const idCelda = hoja.getRange(i, 1).getValue();
        if (idCelda && !isNaN(idCelda) && idCelda !== '' && idCelda !== 0) {
          nuevoID = parseInt(idCelda) + 1;
          filaDestino = i + 1;
          break;
        }
      }
    }
    
    // Escribir todos los datos directamente
    hoja.getRange(filaDestino, 1).setValue(nuevoID);                    // A: ID
    hoja.getRange(filaDestino, 2).setValue(new Date(datos.fecha));      // B: FECHA VENTA
    hoja.getRange(filaDestino, 3).setValue(datos.marca);                // C: MARCA
    hoja.getRange(filaDestino, 4).setValue(datos.medida);               // D: MEDIDA
    hoja.getRange(filaDestino, 5).setValue(datos.modelo);               // E: MODELO
    hoja.getRange(filaDestino, 6).setValue(parseInt(datos.aro));        // F: ARO
    hoja.getRange(filaDestino, 7).setValue(parseInt(datos.cantidad));   // G: CANTIDAD
    hoja.getRange(filaDestino, 8).setValue(parseFloat(datos.precio));   // H: PRECIO VENTA UNIT
    
    // Calcular TOTAL VENTA directamente (sin fórmula para evitar problemas de idioma)
    const totalVenta = parseInt(datos.cantidad) * parseFloat(datos.precio);
    hoja.getRange(filaDestino, 9).setValue(totalVenta);                 // I: TOTAL VENTA
    
    hoja.getRange(filaDestino, 10).setValue(datos.cliente || '');       // J: CLIENTE
    hoja.getRange(filaDestino, 11).setValue(datos.servicio || '');      // K: TIPO SERVICIO
    hoja.getRange(filaDestino, 12).setValue(datos.observaciones || ''); // L: OBSERVACIONES
    
    // Actualizar inventario (restar stock)
    descontarInventario(datos);
    
    return {
      success: true,
      message: 'Venta registrada exitosamente',
      id: nuevoID
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error al guardar: ' + error.toString()
    };
  }
}

/**
 * Descuenta del inventario cuando se registra una venta
 */
function descontarInventario(datos) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaInventario = ss.getSheetByName('INVENTARIO COMPLETO');
  const ultimaFilaInv = hojaInventario.getLastRow();
  
  if (ultimaFilaInv < 2) return;
  
  const datosInventario = hojaInventario.getRange(2, 1, ultimaFilaInv - 1, 11).getValues();
  
  for (let i = 0; i < datosInventario.length; i++) {
    const fila = datosInventario[i];
    // Verificar por ARO, marca, medida y modelo
    const aroFila = parseInt(fila[1]) || 0;
    const aroNuevo = parseInt(datos.aro) || 0;
    if (aroFila === aroNuevo && fila[2] === datos.marca && fila[3] === datos.medida && fila[4] === datos.modelo) {
      const filaReal = i + 2;
      
      // F: Actualizar STOCK (restar cantidad vendida)
      const stockActual = fila[5] || 0;
      const nuevoStock = Math.max(0, stockActual - parseInt(datos.cantidad));
      hojaInventario.getRange(filaReal, 6).setValue(nuevoStock);
      
      break;
    }
  }
}

// ============================================================
// FUNCIONES PARA INVENTARIO
// ============================================================

function obtenerInventario() {
  try {
    Logger.log('Iniciando obtenerInventario...');
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('Spreadsheet abierto: ' + ss.getName());
    
    const hoja = ss.getSheetByName('INVENTARIO COMPLETO');
    
    if (!hoja) {
      Logger.log('ERROR: Hoja INVENTARIO COMPLETO no encontrada');
      return {
        success: false,
        message: 'Hoja INVENTARIO COMPLETO no encontrada'
      };
    }
    
    Logger.log('Hoja encontrada: ' + hoja.getName());
    
    const ultimaFila = hoja.getLastRow();
    Logger.log('Última fila: ' + ultimaFila);
    
    if (ultimaFila < 2) {
      Logger.log('No hay datos en la hoja');
      return {
        success: true,
        data: []
      };
    }
    
    // Leer solo las filas con datos (desde fila 2 hasta la última)
    const numFilas = ultimaFila - 1;
    const datos = hoja.getRange(2, 1, numFilas, 11).getValues();
    Logger.log('Datos leídos: ' + datos.length + ' filas');
    
    const inventario = [];
    
    for (let i = 0; i < datos.length; i++) {
      const fila = datos[i];
      
      // Solo agregar si tiene código (columna A)
      if (fila[0] && fila[0] !== '') {
        // Formatear fecha para enviar al cliente
        let fechaFormateada = '';
        if (fila[9]) {
          try {
            const fecha = new Date(fila[9]);
            if (!isNaN(fecha.getTime())) {
              const dia = String(fecha.getDate()).padStart(2, '0');
              const mes = String(fecha.getMonth() + 1).padStart(2, '0');
              const anio = fecha.getFullYear();
              fechaFormateada = dia + '/' + mes + '/' + anio;
            }
          } catch(e) {
            fechaFormateada = String(fila[9]);
          }
        }
        
        inventario.push({
          codigo: fila[0] || '',
          aro: fila[1] || 0,
          marca: fila[2] || '',
          medida: fila[3] || '',
          modelo: fila[4] || '',
          stock: fila[5] || 0,
          precioCompra: fila[6] || 0,
          precioVenta: fila[8] || 0,
          ultimaCompra: fechaFormateada
        });
      }
    }
    
    Logger.log('Inventario procesado: ' + inventario.length + ' productos');
    
    return {
      success: true,
      data: inventario
    };
    
  } catch (error) {
    Logger.log('ERROR en obtenerInventario: ' + error.toString());
    return {
      success: false,
      message: 'Error al cargar inventario: ' + error.toString()
    };
  }
}

// ============================================================
// FUNCIONES PARA BÚSQUEDA
// ============================================================

function buscarProducto(marca, medida) {
  try {
    Logger.log('Buscando producto - Marca: ' + marca + ', Medida: ' + medida);
    
    if (!marca || !medida) {
      return {
        success: false,
        message: 'Debe ingresar marca y medida'
      };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName('INVENTARIO COMPLETO');
    
    if (!hoja) {
      return {
        success: false,
        message: 'Hoja INVENTARIO COMPLETO no encontrada'
      };
    }
    
    const ultimaFila = hoja.getLastRow();
    
    if (ultimaFila < 2) {
      return {
        success: false,
        message: 'No hay productos en el inventario'
      };
    }
    
    const datos = hoja.getRange(2, 1, ultimaFila - 1, 11).getValues();
    
    // Buscar el producto (comparación sin distinción de mayúsculas)
    const marcaBuscar = marca.toLowerCase().trim();
    const medidaBuscar = medida.toLowerCase().trim();
    
    for (let i = 0; i < datos.length; i++) {
      const fila = datos[i];
      const marcaFila = String(fila[2] || '').toLowerCase().trim();
      const medidaFila = String(fila[3] || '').toLowerCase().trim();
      
      if (marcaFila === marcaBuscar && medidaFila === medidaBuscar) {
        Logger.log('Producto encontrado en fila ' + (i + 2));
        
        return {
          success: true,
          data: {
            codigo: fila[0] || '',
            aro: fila[1] || 0,
            marca: fila[2] || '',
            medida: fila[3] || '',
            modelo: fila[4] || '',
            stock: fila[5] || 0,
            precioCompra: fila[6] || 0,
            precioVenta: fila[8] || 0,
            ultimaCompra: fila[9] || ''
          }
        };
      }
    }
    
    Logger.log('Producto no encontrado');
    return {
      success: false,
      message: 'Producto no encontrado'
    };
    
  } catch (error) {
    Logger.log('ERROR en buscarProducto: ' + error.toString());
    return {
      success: false,
      message: 'Error en la búsqueda: ' + error.toString()
    };
  }
}

// ============================================================
// FUNCIONES PARA ESTADÍSTICAS
// ============================================================

function obtenerEstadisticas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConsultas = ss.getSheetByName('CONSULTAS');
    
    // Leer las celdas con las estadísticas
    const totalProductos = hojaConsultas.getRange('B14').getValue();
    const totalUnidades = hojaConsultas.getRange('B15').getValue();
    const valorInventario = hojaConsultas.getRange('B16').getValue();
    const totalCompras = hojaConsultas.getRange('B17').getValue();
    const totalVentas = hojaConsultas.getRange('B18').getValue();
    
    return {
      success: true,
      data: {
        totalProductos: totalProductos,
        totalUnidades: totalUnidades,
        valorInventario: valorInventario,
        totalCompras: totalCompras,
        totalVentas: totalVentas
      }
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error al cargar estadísticas: ' + error.toString()
    };
  }
}

// ============================================================
// FUNCIONES AUXILIARES
// ============================================================

function formatearFecha(fecha) {
  if (!fecha) return '';
  
  const d = new Date(fecha);
  const dia = String(d.getDate()).padStart(2, '0');
  const mes = String(d.getMonth() + 1).padStart(2, '0');
  const año = d.getFullYear();
  
  return `${dia}/${mes}/${año}`;
}

function formatearPrecio(precio) {
  if (!precio) return '$0';
  
  return '$' + Number(precio).toLocaleString('es-CL');
}

// ============================================================
// FUNCIÓN DE PRUEBA
// ============================================================

function testConexion() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const nombre = ss.getName();
    
    Logger.log('Conexión exitosa con: ' + nombre);
    return 'Conexión exitosa con: ' + nombre;
    
  } catch (error) {
    Logger.log('Error de conexión: ' + error.toString());
    return 'Error: ' + error.toString();
  }
}

/**
 * FUNCIÓN DE PRUEBA - Ejecuta esto para verificar que obtenerInventario funciona
 * Ve a Ejecutar > testObtenerInventario
 * Luego ve a Ver > Registros para ver el resultado
 */
function testObtenerInventario() {
  const resultado = obtenerInventario();
  Logger.log('=== RESULTADO DE obtenerInventario ===');
  Logger.log('Success: ' + resultado.success);
  
  if (resultado.success) {
    Logger.log('Cantidad de productos: ' + resultado.data.length);
    if (resultado.data.length > 0) {
      Logger.log('Primer producto: ' + JSON.stringify(resultado.data[0]));
    }
  } else {
    Logger.log('Error: ' + resultado.message);
  }
  
  return resultado;
}

/**
 * FUNCIÓN DE PRUEBA - Verifica que la hoja existe y tiene datos
 */
function testVerificarHoja() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('Spreadsheet: ' + ss.getName());
    
    // Listar todas las hojas
    const hojas = ss.getSheets();
    Logger.log('Hojas disponibles:');
    hojas.forEach(function(h) {
      Logger.log(' - ' + h.getName() + ' (filas: ' + h.getLastRow() + ')');
    });
    
    // Verificar hoja de inventario
    const hojaInv = ss.getSheetByName('INVENTARIO COMPLETO');
    if (hojaInv) {
      Logger.log('Hoja INVENTARIO COMPLETO encontrada');
      Logger.log('Última fila: ' + hojaInv.getLastRow());
      Logger.log('Última columna: ' + hojaInv.getLastColumn());
      
      // Leer encabezados
      if (hojaInv.getLastRow() >= 1) {
        const encabezados = hojaInv.getRange(1, 1, 1, hojaInv.getLastColumn()).getValues()[0];
        Logger.log('Encabezados: ' + encabezados.join(', '));
      }
      
      // Leer primera fila de datos
      if (hojaInv.getLastRow() >= 2) {
        const primeraFila = hojaInv.getRange(2, 1, 1, hojaInv.getLastColumn()).getValues()[0];
        Logger.log('Primera fila de datos: ' + primeraFila.join(', '));
      }
    } else {
      Logger.log('ERROR: Hoja INVENTARIO COMPLETO NO encontrada');
    }
    
    return 'Ver registros para más detalles';
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    return 'Error: ' + error.toString();
  }
}

// ============================================================
// SINCRONIZACIÓN DE INVENTARIO
// Recalcula todo el inventario basándose en HISTORIAL COMPRAS y REGISTRO VENTAS
// ============================================================

/**
 * FUNCIÓN PRINCIPAL DE SINCRONIZACIÓN
 * Ejecuta esta función para recalcular todo el inventario
 * basándose en las compras y ventas registradas en las hojas
 */
function sincronizarInventario() {
  try {
    Logger.log('=== INICIANDO SINCRONIZACIÓN DE INVENTARIO ===');
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaCompras = ss.getSheetByName('HISTORIAL COMPRAS');
    const hojaVentas = ss.getSheetByName('REGISTRO VENTAS');
    const hojaInventario = ss.getSheetByName('INVENTARIO COMPLETO');
    
    if (!hojaCompras || !hojaInventario) {
      Logger.log('ERROR: Hojas no encontradas');
      return { success: false, message: 'Hojas no encontradas' };
    }
    
    // Objeto para almacenar el inventario calculado
    // Clave: marca|medida|modelo
    const inventarioCalculado = {};
    
    // ========== PROCESAR COMPRAS ==========
    const ultimaFilaCompras = hojaCompras.getLastRow();
    Logger.log('Procesando compras... (filas: ' + ultimaFilaCompras + ')');
    
    if (ultimaFilaCompras >= 2) {
      const datosCompras = hojaCompras.getRange(2, 1, ultimaFilaCompras - 1, 13).getValues();
      
      for (let i = 0; i < datosCompras.length; i++) {
        const fila = datosCompras[i];
        
        // Verificar que la fila tenga datos válidos
        if (!fila[2] || !fila[3]) continue; // Necesita marca y medida
        
        const marca = String(fila[2]).trim();
        const medida = String(fila[3]).trim();
        const modelo = String(fila[4] || '').trim();
        const aro = parseInt(fila[5]) || 0;
        const cantidad = parseInt(fila[6]) || 0;
        const precioCompra = parseFloat(fila[7]) || 0;
        const fecha = fila[1];
        
        // Clave única incluye ARO para diferenciar el mismo neumático en distintos aros
        const clave = aro + '|' + marca.toLowerCase() + '|' + medida.toLowerCase() + '|' + modelo.toLowerCase();
        
        if (!inventarioCalculado[clave]) {
          inventarioCalculado[clave] = {
            marca: marca,
            medida: medida,
            modelo: modelo,
            aro: aro,
            stockCompras: 0,
            stockVentas: 0,
            totalGastado: 0,
            ultimaCompra: null,
            ultimaCantidad: 0
          };
        }
        
        inventarioCalculado[clave].stockCompras += cantidad;
        inventarioCalculado[clave].totalGastado += (cantidad * precioCompra);
        
        // Actualizar última compra
        if (!inventarioCalculado[clave].ultimaCompra || fecha > inventarioCalculado[clave].ultimaCompra) {
          inventarioCalculado[clave].ultimaCompra = fecha;
          inventarioCalculado[clave].ultimaCantidad = cantidad;
        }
      }
    }
    
    // ========== PROCESAR VENTAS ==========
    if (hojaVentas) {
      const ultimaFilaVentas = hojaVentas.getLastRow();
      Logger.log('Procesando ventas... (filas: ' + ultimaFilaVentas + ')');
      
      if (ultimaFilaVentas >= 2) {
        const datosVentas = hojaVentas.getRange(2, 1, ultimaFilaVentas - 1, 12).getValues();
        
        for (let i = 0; i < datosVentas.length; i++) {
          const fila = datosVentas[i];
          
          // Verificar que la fila tenga datos válidos
          if (!fila[2] || !fila[3]) continue; // Necesita marca y medida
          
          const marca = String(fila[2]).trim();
          const medida = String(fila[3]).trim();
          const modelo = String(fila[4] || '').trim();
          const aro = parseInt(fila[5]) || 0;
          const cantidad = parseInt(fila[6]) || 0;
          
          // Clave única incluye ARO
          const clave = aro + '|' + marca.toLowerCase() + '|' + medida.toLowerCase() + '|' + modelo.toLowerCase();
          
          if (inventarioCalculado[clave]) {
            inventarioCalculado[clave].stockVentas += cantidad;
          }
        }
      }
    }
    
    // ========== ACTUALIZAR HOJA DE INVENTARIO ==========
    Logger.log('Actualizando inventario...');
    
    // Limpiar inventario existente (excepto encabezados)
    const ultimaFilaInv = hojaInventario.getLastRow();
    if (ultimaFilaInv >= 2) {
      hojaInventario.getRange(2, 1, ultimaFilaInv - 1, 11).clearContent();
    }
    
    // Escribir nuevo inventario
    const claves = Object.keys(inventarioCalculado);
    Logger.log('Productos a escribir: ' + claves.length);
    
    // Contador para códigos por ARO
    const contadorPorAro = {};
    
    for (let i = 0; i < claves.length; i++) {
      const producto = inventarioCalculado[claves[i]];
      const filaDestino = i + 2;
      
      // Calcular stock actual
      const stockActual = producto.stockCompras - producto.stockVentas;
      
      // Calcular precio promedio
      const precioPromedio = producto.stockCompras > 0 ? producto.totalGastado / producto.stockCompras : 0;
      
      // Calcular precio +35%
      const precioVenta = precioPromedio * 1.35;
      
      // Generar código
      if (!contadorPorAro[producto.aro]) {
        contadorPorAro[producto.aro] = 1;
      }
      const codigo = 'NEU-' + producto.aro + '-' + String(contadorPorAro[producto.aro]).padStart(3, '0');
      contadorPorAro[producto.aro]++;
      
      // Escribir fila
      hojaInventario.getRange(filaDestino, 1).setValue(codigo);                    // A: CÓDIGO
      hojaInventario.getRange(filaDestino, 2).setValue(producto.aro);              // B: ARO
      hojaInventario.getRange(filaDestino, 3).setValue(producto.marca);            // C: MARCA
      hojaInventario.getRange(filaDestino, 4).setValue(producto.medida);           // D: MEDIDA
      hojaInventario.getRange(filaDestino, 5).setValue(producto.modelo);           // E: MODELO
      hojaInventario.getRange(filaDestino, 6).setValue(stockActual);               // F: STOCK ACTUAL
      hojaInventario.getRange(filaDestino, 7).setValue(precioPromedio);            // G: PRECIO COMPRA PROMEDIO
      hojaInventario.getRange(filaDestino, 8).setValue(precioVenta);               // H: PRECIO +35%
      hojaInventario.getRange(filaDestino, 9).setValue(precioVenta);               // I: PRECIO VENTA
      hojaInventario.getRange(filaDestino, 10).setValue(producto.ultimaCompra);    // J: ÚLTIMA COMPRA
      hojaInventario.getRange(filaDestino, 11).setValue(producto.ultimaCantidad);  // K: ÚLTIMA CANTIDAD
    }
    
    Logger.log('=== SINCRONIZACIÓN COMPLETADA ===');
    Logger.log('Productos sincronizados: ' + claves.length);
    
    return {
      success: true,
      message: 'Inventario sincronizado correctamente',
      productosActualizados: claves.length
    };
    
  } catch (error) {
    Logger.log('ERROR en sincronización: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

/**
 * TRIGGER AUTOMÁTICO
 * Se ejecuta cada vez que se edita la hoja de cálculo
 * Para activarlo: Ir a Triggers (Desencadenadores) > Agregar trigger
 * Seleccionar: onEdit, Al editar
 */
function onEdit(e) {
  try {
    const hoja = e.source.getActiveSheet();
    const nombreHoja = hoja.getName();
    
    // Solo sincronizar si se edita HISTORIAL COMPRAS o REGISTRO VENTAS
    if (nombreHoja === 'HISTORIAL COMPRAS' || nombreHoja === 'REGISTRO VENTAS') {
      // Esperar un momento para asegurar que los datos se guardaron
      Utilities.sleep(500);
      
      // Sincronizar inventario
      sincronizarInventario();
      
      Logger.log('Inventario sincronizado automáticamente después de editar ' + nombreHoja);
    }
  } catch (error) {
    Logger.log('Error en onEdit: ' + error.toString());
  }
}

/**
 * TRIGGER INSTALABLE (más confiable que onEdit simple)
 * Para instalarlo, ejecuta la función instalarTriggerEdicion()
 */
function instalarTriggerEdicion() {
  // Eliminar triggers anteriores de este tipo
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sincronizarAlEditar') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Crear nuevo trigger
  ScriptApp.newTrigger('sincronizarAlEditar')
    .forSpreadsheet(SPREADSHEET_ID)
    .onEdit()
    .create();
    
  Logger.log('Trigger de edición instalado correctamente');
  return 'Trigger instalado correctamente';
}

function sincronizarAlEditar(e) {
  try {
    const hoja = e.source.getActiveSheet();
    const nombreHoja = hoja.getName();
    
    // Solo sincronizar si se edita HISTORIAL COMPRAS o REGISTRO VENTAS
    if (nombreHoja === 'HISTORIAL COMPRAS' || nombreHoja === 'REGISTRO VENTAS') {
      sincronizarInventario();
    }
  } catch (error) {
    Logger.log('Error en sincronizarAlEditar: ' + error.toString());
  }
}

/**
 * Agregar botón en el menú de Google Sheets para sincronizar manualmente
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TSR Inventario')
    .addItem('Sincronizar Inventario Ahora', 'sincronizarInventarioConMensaje')
    .addSeparator()
    .addSubMenu(ui.createMenu('Importar Datos')
      .addItem('1. Validar Datos Importados', 'validarDatosImportados')
      .addItem('2. Limpiar y Normalizar Datos', 'limpiarDatosImportados')
      .addItem('3. Sincronizar Inventario', 'sincronizarInventarioConMensaje'))
    .addSeparator()
    .addItem('Ver Resumen', 'mostrarResumenInventario')
    .addItem('Instalar Sincronización Automática', 'instalarTriggerEdicion')
    .addToUi();
}

/**
 * Sincroniza el inventario y muestra un mensaje al usuario
 */
function sincronizarInventarioConMensaje() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert('Sincronizando...', 'Por favor espera mientras se sincroniza el inventario.', ui.ButtonSet.OK);
  
  const resultado = sincronizarInventario();
  
  if (resultado.success) {
    ui.alert('Sincronización Exitosa', 
      'El inventario se ha sincronizado correctamente.\n\n' +
      'Productos actualizados: ' + resultado.productosActualizados, 
      ui.ButtonSet.OK);
  } else {
    ui.alert('Error', 'Hubo un error: ' + resultado.message, ui.ButtonSet.OK);
  }
}

/**
 * Muestra un resumen del inventario actual
 */
function mostrarResumenInventario() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaCompras = ss.getSheetByName('HISTORIAL COMPRAS');
  const hojaVentas = ss.getSheetByName('REGISTRO VENTAS');
  const hojaInventario = ss.getSheetByName('INVENTARIO COMPLETO');
  
  const totalCompras = hojaCompras ? Math.max(0, hojaCompras.getLastRow() - 1) : 0;
  const totalVentas = hojaVentas ? Math.max(0, hojaVentas.getLastRow() - 1) : 0;
  const totalProductos = hojaInventario ? Math.max(0, hojaInventario.getLastRow() - 1) : 0;
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Resumen del Sistema', 
    'HISTORIAL COMPRAS: ' + totalCompras + ' registros\n' +
    'REGISTRO VENTAS: ' + totalVentas + ' registros\n' +
    'INVENTARIO COMPLETO: ' + totalProductos + ' productos\n\n' +
    'Si los números no coinciden con lo esperado,\nusa "Sincronizar Inventario Ahora".',
    ui.ButtonSet.OK);
}

/**
 * Valida los datos importados y muestra errores encontrados
 */
function validarDatosImportados() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let errores = [];
  let advertencias = [];
  
  // ===== VALIDAR HISTORIAL COMPRAS =====
  const hojaCompras = ss.getSheetByName('HISTORIAL COMPRAS');
  if (hojaCompras && hojaCompras.getLastRow() >= 2) {
    const datosCompras = hojaCompras.getRange(2, 1, hojaCompras.getLastRow() - 1, 13).getValues();
    
    for (let i = 0; i < datosCompras.length; i++) {
      const fila = datosCompras[i];
      const numFila = i + 2;
      
      // Verificar campos obligatorios
      if (!fila[2]) errores.push('Compras fila ' + numFila + ': Falta MARCA');
      if (!fila[3]) errores.push('Compras fila ' + numFila + ': Falta MEDIDA');
      if (!fila[5] || isNaN(fila[5])) advertencias.push('Compras fila ' + numFila + ': ARO inválido o vacío');
      if (!fila[6] || isNaN(fila[6])) errores.push('Compras fila ' + numFila + ': CANTIDAD inválida');
      if (!fila[7] || isNaN(fila[7])) errores.push('Compras fila ' + numFila + ': PRECIO inválido');
    }
  }
  
  // ===== VALIDAR REGISTRO VENTAS =====
  const hojaVentas = ss.getSheetByName('REGISTRO VENTAS');
  if (hojaVentas && hojaVentas.getLastRow() >= 2) {
    const datosVentas = hojaVentas.getRange(2, 1, hojaVentas.getLastRow() - 1, 12).getValues();
    
    for (let i = 0; i < datosVentas.length; i++) {
      const fila = datosVentas[i];
      const numFila = i + 2;
      
      // Verificar campos obligatorios
      if (!fila[2]) errores.push('Ventas fila ' + numFila + ': Falta MARCA');
      if (!fila[3]) errores.push('Ventas fila ' + numFila + ': Falta MEDIDA');
      if (!fila[6] || isNaN(fila[6])) errores.push('Ventas fila ' + numFila + ': CANTIDAD inválida');
    }
  }
  
  // Mostrar resultados
  let mensaje = '';
  
  if (errores.length === 0 && advertencias.length === 0) {
    mensaje = '¡Todos los datos están correctos!\n\nPuedes sincronizar el inventario.';
    ui.alert('Validación Exitosa', mensaje, ui.ButtonSet.OK);
  } else {
    if (errores.length > 0) {
      mensaje += 'ERRORES (deben corregirse):\n';
      mensaje += errores.slice(0, 10).join('\n');
      if (errores.length > 10) mensaje += '\n... y ' + (errores.length - 10) + ' errores más';
      mensaje += '\n\n';
    }
    if (advertencias.length > 0) {
      mensaje += 'ADVERTENCIAS:\n';
      mensaje += advertencias.slice(0, 5).join('\n');
      if (advertencias.length > 5) mensaje += '\n... y ' + (advertencias.length - 5) + ' advertencias más';
    }
    
    ui.alert('Problemas Encontrados', mensaje, ui.ButtonSet.OK);
  }
  
  return { errores: errores, advertencias: advertencias };
}

// ============================================================
// FUNCIONES PARA IMPORTAR DATOS DESDE EXCEL
// ============================================================

/**
 * INSTRUCCIONES PARA IMPORTAR DATOS:
 * 
 * 1. Asegúrate de que tu Excel tenga las siguientes hojas con estos nombres EXACTOS:
 *    - HISTORIAL COMPRAS
 *    - REGISTRO VENTAS
 *    - INVENTARIO COMPLETO (opcional, se recalculará)
 * 
 * 2. Las columnas deben estar en este orden:
 * 
 *    HISTORIAL COMPRAS:
 *    A: ID | B: FECHA | C: MARCA | D: MEDIDA | E: MODELO | F: ARO | G: CANTIDAD | H: PRECIO COMPRA | I: PRECIO +35% | J: PRECIO VENTA | K: PROVEEDOR | L: N° FACTURA | M: OBSERVACIONES
 * 
 *    REGISTRO VENTAS:
 *    A: ID | B: FECHA | C: MARCA | D: MEDIDA | E: MODELO | F: ARO | G: CANTIDAD | H: PRECIO VENTA UNIT | I: TOTAL VENTA | J: CLIENTE | K: TIPO SERVICIO | L: OBSERVACIONES
 * 
 * 3. Pasos para importar:
 *    a) Abre tu archivo Excel
 *    b) Selecciona todos los datos de HISTORIAL COMPRAS (sin encabezados)
 *    c) Copia (Ctrl+C)
 *    d) Ve a Google Sheets, hoja HISTORIAL COMPRAS, celda A2
 *    e) Pega (Ctrl+V)
 *    f) Repite para REGISTRO VENTAS
 *    g) Ve al menú TSR Inventario > Validar Datos Importados
 *    h) Si no hay errores, ve a TSR Inventario > Sincronizar Inventario Ahora
 */

/**
 * Limpia y normaliza los datos importados
 */
function limpiarDatosImportados() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let filasLimpiadas = 0;
  
  // Limpiar HISTORIAL COMPRAS
  const hojaCompras = ss.getSheetByName('HISTORIAL COMPRAS');
  if (hojaCompras && hojaCompras.getLastRow() >= 2) {
    const rango = hojaCompras.getRange(2, 1, hojaCompras.getLastRow() - 1, 13);
    const datos = rango.getValues();
    
    for (let i = 0; i < datos.length; i++) {
      // Limpiar espacios en blanco
      if (datos[i][2]) datos[i][2] = String(datos[i][2]).trim(); // Marca
      if (datos[i][3]) datos[i][3] = String(datos[i][3]).trim(); // Medida
      if (datos[i][4]) datos[i][4] = String(datos[i][4]).trim(); // Modelo
      
      // Asegurar que ARO sea número
      if (datos[i][5]) datos[i][5] = parseInt(datos[i][5]) || 0;
      
      // Asegurar que cantidad sea número
      if (datos[i][6]) datos[i][6] = parseInt(datos[i][6]) || 0;
      
      // Asegurar que precio sea número
      if (datos[i][7]) datos[i][7] = parseFloat(String(datos[i][7]).replace(/[^0-9.,]/g, '').replace(',', '.')) || 0;
      
      // Recalcular precio +35%
      if (datos[i][7] > 0) {
        datos[i][8] = datos[i][7] * 1.35;
        datos[i][9] = datos[i][7] * 1.35;
      }
      
      // Generar ID si no existe
      if (!datos[i][0] || datos[i][0] === '') {
        datos[i][0] = i + 1;
      }
      
      filasLimpiadas++;
    }
    
    rango.setValues(datos);
  }
  
  // Limpiar REGISTRO VENTAS
  const hojaVentas = ss.getSheetByName('REGISTRO VENTAS');
  if (hojaVentas && hojaVentas.getLastRow() >= 2) {
    const rango = hojaVentas.getRange(2, 1, hojaVentas.getLastRow() - 1, 12);
    const datos = rango.getValues();
    
    for (let i = 0; i < datos.length; i++) {
      // Limpiar espacios en blanco
      if (datos[i][2]) datos[i][2] = String(datos[i][2]).trim(); // Marca
      if (datos[i][3]) datos[i][3] = String(datos[i][3]).trim(); // Medida
      if (datos[i][4]) datos[i][4] = String(datos[i][4]).trim(); // Modelo
      
      // Asegurar que ARO sea número
      if (datos[i][5]) datos[i][5] = parseInt(datos[i][5]) || 0;
      
      // Asegurar que cantidad sea número
      if (datos[i][6]) datos[i][6] = parseInt(datos[i][6]) || 0;
      
      // Asegurar que precio sea número
      if (datos[i][7]) datos[i][7] = parseFloat(String(datos[i][7]).replace(/[^0-9.,]/g, '').replace(',', '.')) || 0;
      
      // Recalcular total
      if (datos[i][6] > 0 && datos[i][7] > 0) {
        datos[i][8] = datos[i][6] * datos[i][7];
      }
      
      // Generar ID si no existe
      if (!datos[i][0] || datos[i][0] === '') {
        datos[i][0] = i + 1;
      }
      
      filasLimpiadas++;
    }
    
    rango.setValues(datos);
  }
  
  ui.alert('Limpieza Completada', 
    'Se limpiaron y normalizaron ' + filasLimpiadas + ' filas.\n\n' +
    'Ahora puedes sincronizar el inventario.',
    ui.ButtonSet.OK);
  
  return filasLimpiadas;
}