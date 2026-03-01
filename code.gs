/**
 * Despliega la interfaz HTML.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Sistema de Inventario");
}

/**
 * Función auxiliar para incluir archivos HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Función de prueba para verificar que la búsqueda funcione
 */
function probarBusqueda() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inventario");
    
    if (!sheet) {
      return { error: "No se encontró la hoja 'Inventario'" };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Buscar índices de columnas importantes
    var idxSerie = headers.indexOf("Serie");
    var idxDispositivo = headers.indexOf("Dispositivo");
    var idxMarca = headers.indexOf("Marca");
    var idxModelo = headers.indexOf("Modelo");
    var idxCantidad = headers.indexOf("Cantidad");
    
    // Analizar datos de ejemplo para series
    var seriesEjemplo = [];
    var dispositivosEjemplo = [];
    var marcasEjemplo = [];
    
    for (var i = 1; i < Math.min(data.length, 6); i++) {
      var row = data[i];
      if (idxSerie !== -1 && row[idxSerie]) {
        seriesEjemplo.push(row[idxSerie].toString());
      }
      if (idxDispositivo !== -1 && row[idxDispositivo]) {
        dispositivosEjemplo.push(row[idxDispositivo].toString());
      }
      if (idxMarca !== -1 && row[idxMarca]) {
        marcasEjemplo.push(row[idxMarca].toString());
      }
    }
    
    return {
      success: true,
      totalRows: data.length - 1,
      headers: headers,
      sampleData: data.slice(1, 4), // Primeras 3 filas de datos
      columnIndices: {
        serie: idxSerie,
        dispositivo: idxDispositivo,
        marca: idxMarca,
        modelo: idxModelo,
        cantidad: idxCantidad
      },
      examples: {
        series: seriesEjemplo.slice(0, 3),
        dispositivos: dispositivosEjemplo.slice(0, 3),
        marcas: marcasEjemplo.slice(0, 3)
      }
    };
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Función para obtener todas las series disponibles en el inventario
 */
function obtenerTodasLasSeries() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inventario");
    
    if (!sheet) {
      return { error: "No se encontró la hoja 'Inventario'" };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxSerie = headers.indexOf("Serie");
    
    if (idxSerie === -1) {
      return { error: "No se encontró la columna 'Serie' en el inventario" };
    }
    
    var series = [];
    for (var i = 1; i < data.length; i++) {
      var serie = data[i][idxSerie];
      if (serie && serie.toString().trim() !== "") {
        series.push({
          serie: serie.toString(),
          fila: i + 1,
          dispositivo: data[i][headers.indexOf("Dispositivo")] || "",
          marca: data[i][headers.indexOf("Marca")] || "",
          modelo: data[i][headers.indexOf("Modelo")] || ""
        });
      }
    }
    
    return {
      success: true,
      totalSeries: series.length,
      series: series.slice(0, 10) // Solo las primeras 10 para no sobrecargar
    };
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Obtiene los datos de la hoja "Productos" para llenar las listas.
 * La hoja "Productos" debe tener los encabezados:
 * Productos (columna A), Marcas (B), Modelos (C) y Lugar (D).
 */
function getProductosData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Productos');
  if (!sheet) {
    return [];
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  // Lectura vertical real por columnas:
  // A: Productos, B: Marcas, C: Modelos, D: Lugar, E: Plataforma.
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  return data
    .map(function(row) {
      return {
        producto: (row[0] || "").toString().trim(),
        marca: (row[1] || "").toString().trim(),
        modelo: (row[2] || "").toString().trim(),
        lugar: (row[3] || "").toString().trim(),
        plataforma: (row[4] || "").toString().trim()
      };
    })
    .filter(function(item) {
      return item.producto || item.marca || item.modelo || item.lugar || item.plataforma;
    });
}

/**
 * Registra un movimiento de tipo "Venta".
 *
 * Se espera que formData contenga:
 *  - numeroVenta, cantidadProductos (número de líneas),
 *  - Para cada línea i: "serie"+i, "cantidadVendida"+i, "producto"+i, "marca"+i, "modelo"+i,
 *  - lugar, boletaFactura y tipoMovimiento (debe ser "Venta").
 *
 * Se construye la fila para Registro_Movimientos (26 columnas):
 *  A: Número de Venta  
 *  B: Fecha  
 *  C: Cantidad (suma total de unidades vendidas)  
 *  D-G: Producto1, Marca1, Modelo1, Serie1  
 *  H-K: Producto2, Marca2, Modelo2, Serie2  
 *  L-O: Producto3, Marca3, Modelo3, Serie3  
 *  P-S: Producto4, Marca4, Modelo4, Serie4  
 *  T-W: Producto5, Marca5, Modelo5, Serie5  
 *  X: Lugar  
 *  Y: Boleta/Factura  
 *  Z: Tipo de movimiento
 *
 * Además, se actualiza Inventario y se registra en Histórico:
 *  - Si la línea tiene serie (no vacía), se busca en Inventario por "Serie", se copia la fila (primeras 6 columnas)
 *    a Histórico y se elimina de Inventario.
 *  - Si la línea no tiene serie, se busca en Inventario por Dispositivo, Marca y Modelo y se descuenta la cantidad vendida,
 *    y se registra en Histórico una línea con la cantidad vendida.
 */
function registrarMovimiento(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var registroSheet = ss.getSheetByName("Registro_Movimientos");
  var fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  var numLines = parseInt(formData.cantidadProductos);
  if (isNaN(numLines) || numLines < 1) {
    throw new Error("Número de líneas inválido.");
  }
  
  var totalGlobal = 0;
  var productLines = [];
  for (var i = 1; i <= numLines; i++) {
    var prod = formData["producto" + i];
    var marca = formData["marca" + i];
    var modelo = formData["modelo" + i];
    var serie = (formData["serie" + i] || "").trim();
    var cantidadVendida = (serie !== "") ? 1 : (parseInt(formData["cantidadVendida" + i]) || 0);
    totalGlobal += cantidadVendida;
    productLines.push({ producto: prod, marca: marca, modelo: modelo, serie: serie, cantidad: cantidadVendida });
  }
  
  // Construir la fila para Registro_Movimientos (26 columnas)
  var row = [];
  row.push(formData.numeroVenta);   // Col A: Número de Venta
  row.push(fechaActual);              // Col B: Fecha
  row.push(totalGlobal);              // Col C: Cantidad total vendida
  for (var i = 0; i < 5; i++) {
    if (i < productLines.length) {
      row.push(productLines[i].producto); // ProductoX
      row.push(productLines[i].marca);      // MarcaX
      row.push(productLines[i].modelo);     // ModeloX
      row.push(productLines[i].serie);      // SerieX
    } else {
      row.push(""); row.push(""); row.push(""); row.push("");
    }
  }
  row.push(formData.lugar);           // Col X: Lugar
  row.push(formData.boletaFactura);   // Col Y: Boleta/Factura
  row.push(formData.tipoMovimiento);  // Col Z: Tipo de movimiento
  row.push(formData.plataforma);  // Col AA: Tipo de movimiento
  const idMovimiento = Utilities.getUuid(); // AB
  row.push(idMovimiento);
  
  registroSheet.appendRow(row);
  
  // Actualizar Inventario y registrar en Histórico
  var inventarioSheet = ss.getSheetByName("Inventario");
  var historicoSheet = ss.getSheetByName("Histórico");
  var inventarioData = inventarioSheet.getDataRange().getValues();
  var inventarioHeaders = inventarioData[0];
  
  var idxSerie = inventarioHeaders.indexOf("Serie");
  var idxDispositivo = inventarioHeaders.indexOf("Dispositivo");
  var idxMarca = inventarioHeaders.indexOf("Marca");
  var idxModelo = inventarioHeaders.indexOf("Modelo");
  var idxDescripcion = inventarioHeaders.indexOf("Descripción");
  var idxCantidad = inventarioHeaders.indexOf("Cantidad");
  
  for (var i = 0; i < productLines.length; i++) {
    var line = productLines[i];
    if (line.serie !== "") {
      // Producto con serie: buscar en Inventario por Serie
      for (var r = 1; r < inventarioData.length; r++) {
        if (inventarioData[r][idxSerie].toString() === line.serie) {
          // Copiar las primeras 6 columnas (Inventario: Serie, Dispositivo, Marca, Modelo, Descripción, Cantidad)
          var rowHistorico = inventarioData[r].slice(0, 6);
          historicoSheet.appendRow(rowHistorico);
          inventarioSheet.deleteRow(r + 1);
          break;
        }
      }
    } else {
      // Producto sin serie: buscar en Inventario por Dispositivo, Marca y Modelo
      for (var r = 1; r < inventarioData.length; r++) {
        if (inventarioData[r][idxDispositivo] == line.producto &&
            inventarioData[r][idxMarca] == line.marca &&
            inventarioData[r][idxModelo] == line.modelo) {
          var currentQty = parseInt(inventarioData[r][idxCantidad]);
          var newQty = currentQty - line.cantidad;
          if (newQty < 0) newQty = 0;
          inventarioSheet.getRange(r + 1, idxCantidad + 1).setValue(newQty);
          // Registrar en Histórico: se coloca serie vacía
          var rowHistorico = [
            "", // Serie
            inventarioData[r][idxDispositivo],
            inventarioData[r][idxMarca],
            inventarioData[r][idxModelo],
            inventarioData[r][idxDescripcion],
            line.cantidad

          ];
          historicoSheet.appendRow(rowHistorico);
          break;
        }
      }
    }
  }
  registroSheet.getRange(registroSheet.getLastRow(), 2).setNumberFormat("@");
  
  return "Movimiento de venta registrado correctamente.";
}

/**
 * Busca en Inventario un producto por su serie de manera optimizada.
 * Retorna un objeto con todos los datos de la fila o null.
 */
function buscarProductoPorSerie(serie) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inventario");
    
    if (!sheet) {
      throw new Error("No se encontró la hoja 'Inventario'");
    }
    
    // Validar entrada
    if (!serie || serie.toString().trim() === "") {
      throw new Error("Debe proporcionar un número de serie para buscar");
    }
    
    var serieBusqueda = serie.toString().trim();
    
    // Optimización: usar getRange específico en lugar de getDataRange completo
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return null; // Solo hay encabezados o está vacío
    }
    
    // Forzar formato de texto en la columna A para evitar problemas de formato
    sheet.getRange("A2:A" + lastRow).setNumberFormat("@");
    
    // Obtener solo las columnas necesarias (A, B, C, D, E, F, G)
    var data = sheet.getRange(1, 1, lastRow, 7).getValues();
    var headers = data[0];
    
    // Índices de columnas (optimizado para estructura fija)
    var idxSerie = 0;        // Columna A
    var idxDispositivo = 1;   // Columna B
    var idxMarca = 2;         // Columna C
    var idxModelo = 3;        // Columna D
    var idxDescripcion = 4;   // Columna E
    var idxCantidad = 5;      // Columna F
    var idxFecha = 6;         // Columna G
    
    // Búsqueda optimizada por serie (búsqueda exacta)
    for (var i = 1; i < data.length; i++) {
      var serieFila = data[i][idxSerie];
      var serieFilaStr = serieFila ? serieFila.toString().trim() : "";
      
      if (serieFilaStr === serieBusqueda && serieFilaStr !== "") {
        return {
          fila: i + 1,
          serie: serieFilaStr,
          dispositivo: data[i][idxDispositivo] || "",
          marca: data[i][idxMarca] || "",
          modelo: data[i][idxModelo] || "",
          descripcion: data[i][idxDescripcion] || "",
          cantidad: parseInt(data[i][idxCantidad]) || 0,
          fecha: data[i][idxFecha] || "",
          headers: headers
        };
      }
    }
    
    return null;
  } catch (error) {
    console.error("Error en buscarProductoPorSerie:", error);
    throw new Error("Error al buscar producto por serie: " + error.message);
  }
}

/**
 * Registra un movimiento de tipo "Registro" (nuevo producto).
 *
 * Se espera que formData contenga:
 *   - serie, dispositivo, marca, modelo, descripcion, cantidad.
 *
 * En Inventario:
 *   - Si la serie está vacía (producto que se vende por cantidad), se busca por Dispositivo, Marca y Modelo y se suma la cantidad.
 *   - Si no existe o si se proporciona una serie, se agrega una nueva fila con la fecha actual.
 * En Registro_Movimientos se registra una fila con 26 columnas:
 *   A: vacío, B: Fecha, C: Cantidad, D: Producto1, E: Marca1, F: Modelo1, G: Serie1,
 *   H a W: vacías, X: Lugar (vacío), Y: Boleta/Factura (vacío), Z: Tipo de movimiento ("Registro").
 */
function registrarProducto(formData) {
  if (formData.serie) {
    formData.serie = formData.serie.trim();
    if (!validarSerieUnica(formData.serie)) {
      throw new Error(`La serie ${formData.serie} ya existe en el inventario.`);
    }
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inventarioSheet = ss.getSheetByName("Inventario");
  var registroMovSheet = ss.getSheetByName("Registro_Movimientos");
  var fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  // Actualizar Inventario:
  if (!formData.serie || formData.serie.trim() === "") {
    var inventarioData = inventarioSheet.getDataRange().getValues();
    var inventarioHeaders = inventarioData[0];
    var idxDispositivo = inventarioHeaders.indexOf("Dispositivo");
    var idxMarca = inventarioHeaders.indexOf("Marca");
    var idxModelo = inventarioHeaders.indexOf("Modelo");
    var idxCantidad = inventarioHeaders.indexOf("Cantidad");
    var idxFecha = inventarioHeaders.indexOf("Fecha");
    var found = false;
    for (var r = 1; r < inventarioData.length; r++) {
      if (inventarioData[r][idxDispositivo] == formData.dispositivo &&
          inventarioData[r][idxMarca] == formData.marca &&
          inventarioData[r][idxModelo] == formData.modelo) {
        var currentQty = parseInt(inventarioData[r][idxCantidad]);
        var newQty = currentQty + (parseInt(formData.cantidad) || 1);
        inventarioSheet.getRange(r + 1, idxCantidad + 1).setValue(newQty);
        // Actualizar Fecha de ingreso
        inventarioSheet.getRange(r + 1, idxFecha + 1).setValue(fechaActual);
        found = true;
        break;
      }
    }
    if (!found) {
      var newRowInventario = [];
      newRowInventario.push(""); // Serie vacía
      newRowInventario.push(formData.dispositivo || "");
      newRowInventario.push(formData.marca || "");
      newRowInventario.push(formData.modelo || "");
      newRowInventario.push(formData.descripcion || "");
      newRowInventario.push(formData.cantidad ? parseInt(formData.cantidad) : 1);
      newRowInventario.push(fechaActual);
      inventarioSheet.appendRow(newRowInventario);
    }
  } else {
    // Producto con serie: agregar siempre nueva fila
    var newRowInventario = [];
    newRowInventario.push(formData.serie || "");
    newRowInventario.push(formData.dispositivo || "");
    newRowInventario.push(formData.marca || "");
    newRowInventario.push(formData.modelo || "");
    newRowInventario.push(formData.descripcion || "");
    newRowInventario.push(formData.cantidad ? parseInt(formData.cantidad) : 1);
    newRowInventario.push(fechaActual);
    inventarioSheet.appendRow(newRowInventario);
  }
  
  // Registrar en Registro_Movimientos para Registro.
  var row = [];
  row.push("");               // Col A: vacío
  row.push(fechaActual);      // Col B: Fecha
  row.push(formData.cantidad ? parseInt(formData.cantidad) : 1); // Col C: Cantidad
  row.push(formData.dispositivo || "");  // Col D: Producto1
  row.push(formData.marca || "");        // Col E: Marca1
  row.push(formData.modelo || "");       // Col F: Modelo1
  row.push(formData.serie || "");        // Col G: Serie1
  // Rellenar columnas H a W (16 columnas) con vacío
  for (var i = 0; i < 16; i++) {
    row.push("");
  }
  row.push("");               // Col X: Lugar
  row.push("");               // Col Y: Boleta/Factura
  row.push("Registro");       // Col Z: Tipo de movimiento
  row.push("");               // Col AA: Vacía
  row.push(Utilities.getUuid());
  
  registroMovSheet.appendRow(row);
  
  return "Producto registrado correctamente en Inventario y en Registro de Movimientos.";
}

/**
 * Obtiene el stock actual de la sede desde la hoja "Inventario".
 * Retorna un array de objetos con los datos de los productos.
 */
function getStockActual(dispositivoFiltro = "") {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Inventario");
  var data = sheet.getDataRange().getValues();
  
  var stockMap = {};
  
  // Usar índices fijos (mejor que headers.indexOf para evitar errores)
  const IDX_DISPOSITIVO = 1; // Columna B
  const IDX_MARCA = 2;       // Columna C
  const IDX_MODELO = 3;      // Columna D
  const IDX_CANTIDAD = 5;    // Columna F

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Aplicar filtro
    if (dispositivoFiltro && row[IDX_DISPOSITIVO] !== dispositivoFiltro) continue;
    
    var key = row[IDX_DISPOSITIVO] + "|" + row[IDX_MARCA] + "|" + row[IDX_MODELO];
    
    if (!stockMap[key]) {
      stockMap[key] = {
        dispositivo: row[IDX_DISPOSITIVO],
        marca: row[IDX_MARCA],
        modelo: row[IDX_MODELO],
        cantidad: 0
      };
    }
    stockMap[key].cantidad += parseInt(row[IDX_CANTIDAD]) || 0;
  }
  
  return Object.values(stockMap);
}

function getDispositivosUnicos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Inventario");
  var data = sheet.getRange("B2:B" + sheet.getLastRow()).getValues().flat();
  return [...new Set(data.filter(String))];
}

/**
 * Obtiene las marcas únicas del inventario
 */
function getMarcasUnicas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Inventario");
  var data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues().flat();
  return [...new Set(data.filter(String))].sort();
}

/**
 * Obtiene los modelos únicos del inventario
 */
function getModelosUnicos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Inventario");
  var data = sheet.getRange("D2:D" + sheet.getLastRow()).getValues().flat();
  return [...new Set(data.filter(String))].sort();
}

/**
 * Obtiene el historial simplificado de movimientos
 */
/**
 * Obtiene el historial de movimientos con las columnas correctas
 */
function getHistorialMovimientos(tipoFiltro, mesFiltro) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Registro_Movimientos");
  var data = sheet.getDataRange().getValues();
  
  var historial = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Obtener valores de las columnas por índice
    var fechaString = row[1]; // Columna B
    if (fechaString instanceof Date) {
      fechaString = Utilities.formatDate(fechaString, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");}
    var tipo = row[25];       // Columna Z
    var producto = row[3];    // Columna D
    var numeroVenta = row[0]; // Columna A
    var lugar = row[23];      // Columna X
    var plataforma = row[26];      // Columna X
    var id = row[27]; // <-- Nueva propiedad
    // Convertir fecha string a objeto Date
    var fecha = Utilities.parseDate(fechaString, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var mes = fecha.getMonth() + 1; // Mes numérico (1-12)
    
    // Aplicar filtros
    if (tipoFiltro && tipo !== tipoFiltro) continue;
    if (mesFiltro && mes !== parseInt(mesFiltro)) continue;
    
    historial.push({
      id: id,
      fecha: fechaString,
      tipo: tipo,
      producto: producto,
      cantidad: row[2], // Columna C
      numeroVenta: numeroVenta,
      lugar: lugar,
      plataforma: plataforma
    });
  }
  
  // Ordenar por fecha descendente
  historial.sort(function(a, b) {
    return new Date(b.fecha) - new Date(a.fecha);
  });
  
  return historial;
}

// ========== Carga Masiva ==========
function procesarCargaMasiva(csvData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Inventario");
    const data = Utilities.parseCsv(csvData, ";");
    
    // Validar encabezados
    if (data[0].join(";") !== "Serie;Dispositivo;Marca;Modelo;Descripción;Cantidad") {
        throw new Error("Formato de archivo inválido. Use la plantilla proporcionada.");
    }

    const seriesInvalidas = [];
    const registrosValidos = [];
    
    // Primera pasada: validaciones
    data.slice(1).forEach((row, i) => {
        const linea = i + 2;
        
        // Validar número de columnas
        if (row.length !== 6) {
            throw new Error(`Error en línea ${linea}: Debe tener 6 columnas`);
        }
        
        // Extraer valores renombrando la variable serie
        const [serieRow, dispositivo, marca, modelo, descripcion, cantidad] = row.map(c => c.trim());
        
        // Validar serie única
        if (serieRow && !validarSerieUnica(serieRow)) {
            seriesInvalidas.push(`Línea ${linea}: ${serieRow}`);
        }
        
        // Validar campos obligatorios
        if (!dispositivo || !marca || !modelo) {
            throw new Error(`Línea ${linea}: Dispositivo, Marca y Modelo son requeridos`);
        }
        
        // Validar cantidad
        if (isNaN(cantidad)) {
            throw new Error(`Línea ${linea}: Cantidad inválida ('${cantidad}')`);
        }
        
        // Almacenar registro válido
        registrosValidos.push({
            serie: serieRow,
            dispositivo: dispositivo,
            marca: marca,
            modelo: modelo,
            descripcion: descripcion,
            cantidad: parseInt(cantidad) || 1
        });
    });

    // Validar series duplicadas
    if (seriesInvalidas.length > 0) {
        throw new Error("Series duplicadas en inventario:\n" + seriesInvalidas.join("\n"));
    }

    // Segunda pasada: procesamiento
    registrosValidos.forEach(registro => {
        registrarProducto(registro);
    });

    return `Se procesaron ${registrosValidos.length} registros exitosamente`;
}


function exportarHistorialCSV() {
    const historial = getHistorialMovimientos();
    const csvRows = ["Fecha,Tipo,Producto,Cantidad,Número Venta,Lugar"];
    
    historial.forEach(mov => {
        csvRows.push([
            mov.fecha,
            mov.tipo,
            mov.producto,
            mov.cantidad,
            mov.numeroVenta,
            mov.lugar
        ].join(","));
    });
    
    return csvRows.join("\n");
}

function validarSerieUnica(serie) {
    const inventario = SpreadsheetApp.getActive().getSheetByName("Inventario");
    
    // Forzar formato de texto en la columna A
    const lastRow = inventario.getLastRow();
    if (lastRow > 1) {
        inventario.getRange("A2:A" + lastRow).setNumberFormat("@");
    }
    
    const series = inventario.getRange("A2:A" + lastRow)
                            .getValues()
                            .flat()
                            .map(s => s.toString().trim()); // Convertir a string y trim
    
    const serieBuscada = serie.trim();
    return !series.includes(serieBuscada);
}

/**
 * Agrega una nueva marca a la hoja Productos
 */
function agregarMarca(nombreMarca) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Productos");
    
    if (!sheet) {
      throw new Error("No se encontró la hoja 'Productos'");
    }
    
    if (!nombreMarca || nombreMarca.trim() === "") {
      throw new Error("El nombre de la marca no puede estar vacío");
    }
    
    nombreMarca = nombreMarca.trim();
    
    // Verificar si la marca ya existe
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toLowerCase() === nombreMarca.toLowerCase()) {
        throw new Error("La marca '" + nombreMarca + "' ya existe en el sistema");
      }
    }
    
    // Agregar la nueva marca (columna B)
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 2).setValue(nombreMarca);
    
    return {
      success: true,
      message: "Marca '" + nombreMarca + "' agregada correctamente"
    };
    
  } catch (error) {
    console.error("Error en agregarMarca:", error);
    return {
      error: error.message
    };
  }
}

/**
 * Agrega un nuevo modelo a la hoja Productos
 */
function agregarModelo(marca, nombreModelo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Productos");
    
    if (!sheet) {
      throw new Error("No se encontró la hoja 'Productos'");
    }
    
    if (!marca || marca.trim() === "") {
      throw new Error("Debe seleccionar una marca");
    }
    
    if (!nombreModelo || nombreModelo.trim() === "") {
      throw new Error("El nombre del modelo no puede estar vacío");
    }
    
    marca = marca.trim();
    nombreModelo = nombreModelo.trim();
    
    // Verificar si el modelo ya existe para esa marca
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toLowerCase() === marca.toLowerCase() &&
          data[i][2] && data[i][2].toString().trim().toLowerCase() === nombreModelo.toLowerCase()) {
        throw new Error("El modelo '" + nombreModelo + "' ya existe para la marca '" + marca + "'");
      }
    }
    
    // Agregar el nuevo modelo (columna C)
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue("Producto"); // Columna A: Producto
    sheet.getRange(lastRow + 1, 2).setValue(marca);      // Columna B: Marca
    sheet.getRange(lastRow + 1, 3).setValue(nombreModelo); // Columna C: Modelo
    sheet.getRange(lastRow + 1, 4).setValue("");         // Columna D: Lugar
    sheet.getRange(lastRow + 1, 5).setValue("");         // Columna E: Plataforma
    
    return {
      success: true,
      message: "Modelo '" + nombreModelo + "' agregado correctamente para la marca '" + marca + "'"
    };
    
  } catch (error) {
    console.error("Error en agregarModelo:", error);
    return {
      error: error.message
    };
  }
}

/**
 * Obtiene todas las marcas disponibles de la hoja Productos
 */
function obtenerMarcasDisponibles() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Productos");
    
    if (!sheet) {
      throw new Error("No se encontró la hoja 'Productos'");
    }
    
    var data = sheet.getDataRange().getValues();
    var marcas = new Set();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() !== "") {
        marcas.add(data[i][1].toString().trim());
      }
    }
    
    return Array.from(marcas).sort();
    
  } catch (error) {
    console.error("Error en obtenerMarcasDisponibles:", error);
    return {
      error: error.message
    };
  }
}

/**
 * Obtiene todos los modelos disponibles para una marca específica
 */
function obtenerModelosPorMarca(marca) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Productos");
    
    if (!sheet) {
      throw new Error("No se encontró la hoja 'Productos'");
    }
    
    if (!marca || marca.trim() === "") {
      throw new Error("Debe especificar una marca");
    }
    
    marca = marca.trim();
    var data = sheet.getDataRange().getValues();
    var modelos = new Set();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toLowerCase() === marca.toLowerCase() &&
          data[i][2] && data[i][2].toString().trim() !== "") {
        modelos.add(data[i][2].toString().trim());
      }
    }
    
    return Array.from(modelos).sort();
    
  } catch (error) {
    console.error("Error en obtenerModelosPorMarca:", error);
    return {
      error: error.message
    };
  }
}

/**
 * Actualiza los datos de un producto en el inventario por su serie
 */
function actualizarProductoPorSerie(serie, nuevosDatos) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inventario");
    
    if (!sheet) {
      throw new Error("No se encontró la hoja 'Inventario'");
    }
    
    // Forzar formato de texto en la columna A
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange("A2:A" + lastRow).setNumberFormat("@");
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Buscar la fila por serie en columna A
    var filaEncontrada = -1;
    for (var i = 1; i < data.length; i++) {
      var serieFila = data[i][0]; // Columna A
      // Convertir a string y limpiar espacios para comparación
      var serieFilaStr = serieFila ? serieFila.toString().trim() : "";
      var serieBusqueda = serie ? serie.toString().trim() : "";
      
      if (serieFilaStr === serieBusqueda && serieFilaStr !== "") {
        filaEncontrada = i + 1; // +1 porque las filas en Sheets empiezan en 1
        break;
      }
    }
    
    if (filaEncontrada === -1) {
      throw new Error("No se encontró la serie en el inventario");
    }
    
    // Obtener índices de columnas
    var idxMarca = headers.indexOf("Marca");
    var idxModelo = headers.indexOf("Modelo");
    var idxDescripcion = headers.indexOf("Descripción");
    
    if (idxMarca === -1 || idxModelo === -1 || idxDescripcion === -1) {
      throw new Error("No se encontraron las columnas necesarias (Marca, Modelo, Descripción)");
    }
    
    // Actualizar los campos
    if (nuevosDatos.marca !== undefined) {
      sheet.getRange(filaEncontrada, idxMarca + 1).setValue(nuevosDatos.marca);
    }
    
    if (nuevosDatos.modelo !== undefined) {
      sheet.getRange(filaEncontrada, idxModelo + 1).setValue(nuevosDatos.modelo);
    }
    
    if (nuevosDatos.descripcion !== undefined) {
      sheet.getRange(filaEncontrada, idxDescripcion + 1).setValue(nuevosDatos.descripcion);
    }
    
    return {
      success: true,
      message: "Producto actualizado correctamente",
      fila: filaEncontrada
    };
    
  } catch (error) {
    console.error("Error en actualizarProductoPorSerie:", error);
    throw new Error("Error al actualizar producto: " + error.message);
  }
}

function cancelarMovimiento(idMovimiento) {
  const ss = SpreadsheetApp.getActive();
  const registroSheet = ss.getSheetByName("Registro_Movimientos");
  const historicoSheet = ss.getSheetByName("Histórico");
  const inventarioSheet = ss.getSheetByName("Inventario");

  // Buscar el movimiento por ID
  const [header, ...movimientos] = registroSheet.getDataRange().getValues();
  const columnaId = header.indexOf("ID Movimiento");

  if (columnaId === -1) {
    throw new Error("No se encontró la columna 'ID Movimiento'.");
  }

  const idx = movimientos.findIndex(row => row[columnaId] === idMovimiento);

  if (idx === -1) {
    throw new Error("Movimiento no encontrado.");
  }

  const movimiento = movimientos[idx];
  const tipo = movimiento[header.indexOf("Tipo de movimiento")];

  // Revertir inventario y marcar como cancelado
  if (tipo === "Venta") {
    revertirVenta(movimiento, header, inventarioSheet, historicoSheet);
  } else if (tipo === "Registro") {
    revertirRegistro(movimiento, header, inventarioSheet, historicoSheet);
  }

  // Marcar como cancelado y mover a histórico
  movimiento[header.indexOf("Tipo de movimiento")] = "Cancelado";
  registroSheet.appendRow(movimiento);

  // Eliminar de Registro_Movimientos
  registroSheet.deleteRow(idx + 2); // +2 por header y índice base 1

  return "Movimiento cancelado exitosamente";
}

function revertirVenta(movimiento, header, inventarioSheet, historicoSheet) {
  const historicoData = historicoSheet.getDataRange().getValues();
  const historicoHeader = historicoData[0];

  // Índices de las columnas de series en el movimiento
  const idxSerie1 = header.indexOf("Serie1");
  const idxSerie2 = header.indexOf("Serie2");
  const idxSerie3 = header.indexOf("Serie3");
  const idxSerie4 = header.indexOf("Serie4");
  const idxSerie5 = header.indexOf("Serie5");

  // Array con todas las series del movimiento
  const series = [
    movimiento[idxSerie1],
    movimiento[idxSerie2],
    movimiento[idxSerie3],
    movimiento[idxSerie4],
    movimiento[idxSerie5],
  ].filter(serie => serie && serie.trim() !== ""); // Filtrar series vacías o con solo espacios

  console.log("Series a restaurar:", series); // Depuración

  // Restaurar cada serie en el inventario
  series.forEach(serie => {
    // Buscar en histórico
    const idxSerieHistorico = historicoHeader.indexOf("Serie");
    const historicoRow = historicoData.find(row => row[idxSerieHistorico].toString().trim() === serie.trim());

    if (historicoRow) {
      console.log("Restaurando serie:", serie); // Depuración

      // Copiar las primeras 7 columnas (Inventario: Serie, Dispositivo, Marca, Modelo, Descripción, Cantidad, Fecha)
      const rowToRestore = historicoRow.slice(0, 7);
      inventarioSheet.appendRow(rowToRestore);
    } else {
      console.log("Serie no encontrada en histórico:", serie); // Depuración
    }
  });

  // Eliminar filas del histórico en orden inverso
  const rowsToDelete = [];
  series.forEach(serie => {
    const idxSerieHistorico = historicoHeader.indexOf("Serie");
    const historicoRow = historicoData.find(row => row[idxSerieHistorico].toString().trim() === serie.trim());

    if (historicoRow) {
      const rowNum = historicoData.findIndex(row => row === historicoRow) + 1; // +1 porque las filas comienzan en 1
      rowsToDelete.push(rowNum);
      console.log("Fila a eliminar:", rowNum); // Depuración
    }
  });

  // Ordenar las filas a eliminar de mayor a menor
  rowsToDelete.sort((a, b) => b - a);

  // Eliminar filas en orden inverso
  rowsToDelete.forEach(rowNum => {
    console.log("Eliminando fila:", rowNum); // Depuración
    historicoSheet.deleteRow(rowNum);
  });
}

function revertirRegistro(movimiento, header, inventarioSheet, historicoSheet) {
  const serie = movimiento[6]?.toString().trim(); // Serie1 en índice 6 (col G)
  
  if (serie) {
    const inventarioData = inventarioSheet.getDataRange().getValues();
    const idx = inventarioData.findIndex(row => row[0]?.toString().trim() === serie);
    if (idx > 0) inventarioSheet.deleteRow(idx + 1);
  } else {
    // Lógica para productos sin serie
    const dispositivo = movimiento[3]; // Producto1 (col D)
    const marca = movimiento[4];       // Marca1 (col E)
    const modelo = movimiento[5];      // Modelo1 (col F)
    const cantidad = movimiento[2];    // Cantidad (col C)

    const inventarioData = inventarioSheet.getDataRange().getValues();
    const idxDispositivo = 1, idxMarca = 2, idxModelo = 3, idxCantidad = 5;
    const row = inventarioData.find(row => 
      row[idxDispositivo] === dispositivo && 
      row[idxMarca] === marca && 
      row[idxModelo] === modelo
    );
    
    if (row) {
      const newQty = row[idxCantidad] - cantidad;
      if (newQty <= 0) {
        inventarioSheet.deleteRow(inventarioData.indexOf(row) + 1);
      } else {
        inventarioSheet.getRange(inventarioData.indexOf(row) + 1, idxCantidad + 1).setValue(newQty);
      }
    }
  }
}

function inicializarColumnaID() {
    const sheet = SpreadsheetApp.getActive().getSheetByName("Registro_Movimientos");
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    if (!header.includes("ID Movimiento")) {
        sheet.getRange(1, 28).setValue("ID Movimiento"); // Columna AB = 28
    }
}

function agregarIDsFaltantes() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Registro_Movimientos");
  const data = sheet.getDataRange().getValues();
  
  data.forEach((row, index) => {
    if (index > 0 && !row[27]) { // Si AB está vacío
      sheet.getRange(index + 1, 28).setValue(Utilities.getUuid());
    }
  });
}

function verificarColumnas() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Registro_Movimientos");
  const headers = sheet.getRange(1, 1, 1, 28).getValues()[0];
  
  if (headers[27] !== "ID Movimiento") { // Columna AB = índice 27
    sheet.getRange(1, 28).setValue("ID Movimiento");
  }
}

function testValidarSerieUnica() {
    console.log(validarSerieUnica("  ABC123  ")); // Debe retornar false si existe "ABC123"
    console.log(validarSerieUnica("abc123")); // Retornará false si existe "ABC123" (con toLowerCase)
}

/**
 * Inicializa el formato de texto en la columna A de la hoja Inventario
 * Esta función debe ejecutarse una vez para asegurar que las series se lean correctamente
 */
function inicializarFormatoSeries() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inventario");
    
    if (!sheet) {
      throw new Error("No se encontró la hoja 'Inventario'");
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // Aplicar formato de texto a toda la columna A
      sheet.getRange("A2:A" + lastRow).setNumberFormat("@");
      
      // También aplicar formato de texto a la celda A1 (encabezado)
      sheet.getRange("A1").setNumberFormat("@");
      
      return {
        success: true,
        message: "Formato de texto aplicado correctamente a la columna A (Series)",
        filasProcesadas: lastRow
      };
    } else {
      return {
        success: true,
        message: "La hoja Inventario está vacía o solo tiene encabezados"
      };
    }
    
  } catch (error) {
    console.error("Error en inicializarFormatoSeries:", error);
    return {
      error: error.message
    };
  }
}

function loginUser(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Accesos");
  var data = sheet.getDataRange().getValues();
  
  // Asumimos que la primera fila es el encabezado
  data.shift();
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var usr = row[0].toString().trim();
    var name = row[1];
    var pass = row[2].toString().trim();
    var rol = row[3].toString().trim();
    
    if (usr === username && pass === password) {
      return { name: name, rol: rol };
    }
  }
  
  throw new Error("Credenciales inválidas");
}

/**
 * Obtiene métricas para el dashboard
 */
function obtenerMetricasDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventario = ss.getSheetByName("Inventario");
  
  // Obtener stock total y bajo stock
  const stockData = getStockActual();
  let totalStock = 0;
  let lowStock = 0;
  stockData.forEach(item => {
    totalStock += item.cantidad;
    if (item.cantidad < 10) lowStock++;
  });
  
  // Movimientos del mes actual
  const historial = getHistorialMovimientos("", new Date().getMonth() + 1);
  
  return {
    totalStock: totalStock,
    lowStock: lowStock,
    movimientosMes: historial.length,
    valorInventario: totalStock * 0 // Ajusta según tu cálculo real
  };
}

/**
 * Obtiene datos para gráfico de ventas por plataforma
 */
function obtenerVentasPorPlataforma() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registroSheet = ss.getSheetByName("Registro_Movimientos");
  const data = registroSheet.getDataRange().getValues();
  const conteoPlataformas = {};

  data.forEach((row, index) => {
    if (index === 0) return; // Saltar encabezado
    if (row[25] !== "Venta" || row[25] === "Cancelado") return; // Solo ventas activas
    
    const plataforma = row[26] || "Sin plataforma"; // Columna AA
    
    // Contar 1 por cada venta, sin importar la cantidad
    conteoPlataformas[plataforma] = (conteoPlataformas[plataforma] || 0) + 1;
  });

  return {
    labels: Object.keys(conteoPlataformas),
    values: Object.values(conteoPlataformas),
    total: Object.values(conteoPlataformas).reduce((a, b) => a + b, 0)
  };
}

// En Apps Script
function obtenerVentasPorMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Registro_Movimientos");
  const data = sheet.getDataRange().getValues();
  const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  
  const ventasPorMes = {};
  const mesActual = new Date().getMonth() + 1;
  
  // Inicializar todos los meses hasta el actual
  for (let i = 0; i < mesActual; i++) {
    ventasPorMes[i] = 0;
  }

  data.forEach((row, index) => {
    if (index === 0) return;
    if (row[25] !== "Venta" || row[25] === "Cancelado") return;
    
    const fecha = new Date(row[1]);
    const mesIdx = fecha.getMonth();
    
    if (mesIdx < mesActual) {
      ventasPorMes[mesIdx] += 1;
    }
  });

  const labels = [];
  const values = [];
  for (let i = 0; i < mesActual; i++) {
    labels.push(meses[i]);
    values.push(ventasPorMes[i]);
  }

  return { labels, values };
}

function obtenerVentasPorMesYPlataforma() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movimientosSheet = ss.getSheetByName('Registro_Movimientos');
  const data = movimientosSheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  
  // Configurar meses
  const meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 
               'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const ahora = new Date();
  const mesActual = ahora.getMonth();
  const mesesAMostrar = meses.slice(0, mesActual + 1);

  // Obtener plataformas únicas
  const plataformas = [...new Set(
    data.slice(1)
      .filter(row => row[25] === 'Venta' && row[26])
      .map(row => row[26])
  )];

  // Inicializar estructura de datos
  const ventas = {};
  mesesAMostrar.forEach(mes => {
    ventas[mes] = Object.fromEntries(plataformas.map(p => [p, 0]));
  });

  // Procesar datos
  data.slice(1).forEach(row => {
    try {
      const fecha = Utilities.parseDate(row[1], timeZone, "dd/MM/yyyy HH:mm:ss");
      const mesIdx = fecha.getMonth();
      const mesNombre = meses[mesIdx];
      const plataforma = row[26];
      
      if (row[25] === 'Venta' && plataforma && ventas[mesNombre]) {
        ventas[mesNombre][plataforma]++;
      }
    } catch (e) {
      console.error(`Error parseando fecha: ${row[1]}`, e);
    }
  });

  // Preparar datos para Chart.js
  const datasets = plataformas.map((plataforma, index) => {
    const colores = ['#2563eb', '#059669', '#d97706', '#dc2626'];
    return {
      label: plataforma,
      data: mesesAMostrar.map(mes => ventas[mes][plataforma] || 0),
      borderColor: colores[index % colores.length],
      backgroundColor: colores[index % colores.length] + '20',
      tension: 0.4,
      fill: true
    };
  });

  return {
    labels: mesesAMostrar,
    datasets: datasets
  };
}

/**
 * Devuelve los detalles completos de un movimiento de venta por su ID.
 * Retorna un array de objetos con producto, marca, modelo, descripción y cantidad.
 */
function getDetalleMovimiento(idMovimiento) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Registro_Movimientos');
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var idxId = header.indexOf('ID Movimiento');
  if (idxId === -1) throw new Error('No se encontró la columna ID Movimiento');
  var idxs = [
    {producto: 3, marca: 4, modelo: 5, serie: 6},   // Producto1
    {producto: 7, marca: 8, modelo: 9, serie: 10},  // Producto2
    {producto: 11, marca: 12, modelo: 13, serie: 14}, // Producto3
    {producto: 15, marca: 16, modelo: 17, serie: 18}, // Producto4
    {producto: 19, marca: 20, modelo: 21, serie: 22}  // Producto5
  ];
  var detalles = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] === idMovimiento) {
      for (var j = 0; j < idxs.length; j++) {
        var prod = data[i][idxs[j].producto];
        var marca = data[i][idxs[j].marca];
        var modelo = data[i][idxs[j].modelo];
        var serie = data[i][idxs[j].serie];
        if (prod || marca || modelo || serie) {
          // Buscar descripción en Inventario o Histórico
          var descripcion = '';
          var cantidad = 1;
          // Buscar en Histórico por serie si existe
          if (serie) {
            var histSheet = ss.getSheetByName('Histórico');
            var histData = histSheet.getDataRange().getValues();
            var idxSerie = histData[0].indexOf('Serie');
            var idxDesc = histData[0].indexOf('Descripción');
            var idxCant = histData[0].indexOf('Cantidad');
            for (var h = 1; h < histData.length; h++) {
              if (histData[h][idxSerie] == serie) {
                descripcion = histData[h][idxDesc];
                cantidad = histData[h][idxCant];
                break;
              }
            }
          }
          // Si no hay serie, buscar por producto, marca y modelo
          if (!descripcion) {
            var histSheet = ss.getSheetByName('Histórico');
            var histData = histSheet.getDataRange().getValues();
            var idxDisp = histData[0].indexOf('Dispositivo');
            var idxMarca = histData[0].indexOf('Marca');
            var idxModelo = histData[0].indexOf('Modelo');
            var idxDesc = histData[0].indexOf('Descripción');
            var idxCant = histData[0].indexOf('Cantidad');
            for (var h = 1; h < histData.length; h++) {
              if (
                histData[h][idxDisp] == prod &&
                histData[h][idxMarca] == marca &&
                histData[h][idxModelo] == modelo
              ) {
                descripcion = histData[h][idxDesc];
                cantidad = histData[h][idxCant];
                break;
              }
            }
          }
          detalles.push({
            producto: prod,
            marca: marca,
            modelo: modelo,
            descripcion: descripcion,
            serie: serie,
            cantidad: cantidad
          });
        }
      }
      break;
    }
  }
  return detalles;
}

/**
 * Devuelve la descripción de un equipo según dispositivo, marca y modelo.
 */
function getDescripcionPorModelo(dispositivo, marca, modelo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Inventario');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idxDisp = headers.indexOf('Dispositivo');
  var idxMarca = headers.indexOf('Marca');
  var idxModelo = headers.indexOf('Modelo');
  var idxDesc = headers.indexOf('Descripción');
  for (var i = 1; i < data.length; i++) {
    if (
      data[i][idxDisp] == dispositivo &&
      data[i][idxMarca] == marca &&
      data[i][idxModelo] == modelo
    ) {
      return data[i][idxDesc] || '';
    }
  }
  return '';
}


/**
 * Obtiene estadísticas detalladas del inventario
 */
function obtenerEstadisticasInventario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Inventario");
  var data = sheet.getDataRange().getValues();
  
  var stats = {
    totalProductos: 0,
    totalUnidades: 0,
    productosConSerie: 0,
    productosSinSerie: 0,
    stockBajo: 0,
    stockMedio: 0,
    stockAlto: 0,
    marcasUnicas: new Set(),
    dispositivosUnicos: new Set(),
    valorEstimado: 0
  };
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var cantidad = parseInt(row[5]) || 0; // Columna F (Cantidad)
    var serie = row[0] || ""; // Columna A (Serie)
    var dispositivo = row[1] || ""; // Columna B (Dispositivo)
    var marca = row[2] || ""; // Columna C (Marca)
    
    stats.totalProductos++;
    stats.totalUnidades += cantidad;
    
    if (serie.trim() !== "") {
      stats.productosConSerie++;
    } else {
      stats.productosSinSerie++;
    }
    
    if (cantidad < 10) {
      stats.stockBajo++;
    } else if (cantidad < 50) {
      stats.stockMedio++;
    } else {
      stats.stockAlto++;
    }
    
    if (marca) stats.marcasUnicas.add(marca);
    if (dispositivo) stats.dispositivosUnicos.add(dispositivo);
    
    // Valor estimado (puedes ajustar según tus precios)
    stats.valorEstimado += cantidad * 100; // Ejemplo: $100 por unidad
  }
  
  return {
    totalProductos: stats.totalProductos,
    totalUnidades: stats.totalUnidades,
    productosConSerie: stats.productosConSerie,
    productosSinSerie: stats.productosSinSerie,
    stockBajo: stats.stockBajo,
    stockMedio: stats.stockMedio,
    stockAlto: stats.stockAlto,
    marcasUnicas: stats.marcasUnicas.size,
    dispositivosUnicos: stats.dispositivosUnicos.size,
    valorEstimado: stats.valorEstimado
  };
}


/**
 * Obtiene productos con stock bajo (configurable)
 */
function obtenerProductosStockBajo(limite = 10) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Inventario");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var idxSerie = headers.indexOf("Serie");
  var idxDispositivo = headers.indexOf("Dispositivo");
  var idxMarca = headers.indexOf("Marca");
  var idxModelo = headers.indexOf("Modelo");
  var idxCantidad = headers.indexOf("Cantidad");
  var idxDescripcion = headers.indexOf("Descripción");
  
  var productosBajoStock = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var cantidad = parseInt(row[idxCantidad]) || 0;
    
    if (cantidad <= limite) {
      productosBajoStock.push({
        serie: row[idxSerie] || "",
        dispositivo: row[idxDispositivo] || "",
        marca: row[idxMarca] || "",
        modelo: row[idxModelo] || "",
        cantidad: cantidad,
        descripcion: row[idxDescripcion] || "",
        nivel: cantidad === 0 ? "Sin stock" : cantidad <= 5 ? "Crítico" : "Bajo"
      });
    }
  }
  
  return productosBajoStock;
}

/**
 * Obtiene movimientos recientes (últimos N días)
 */
function obtenerMovimientosRecientes(dias = 30) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Registro_Movimientos");
  var data = sheet.getDataRange().getValues();
  
  var fechaLimite = new Date();
  fechaLimite.setDate(fechaLimite.getDate() - dias);
  
  var movimientosRecientes = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var fechaMovimiento = new Date(row[1]); // Columna B (Fecha)
    
    if (fechaMovimiento >= fechaLimite) {
      movimientosRecientes.push({
        fecha: row[1],
        tipo: row[25], // Columna Z
        producto: row[3], // Columna D
        cantidad: row[2], // Columna C
        numeroVenta: row[0], // Columna A
        lugar: row[23], // Columna X
        plataforma: row[26], // Columna AA
        id: row[27] // Columna AB
      });
    }
  }
  
  return movimientosRecientes;
}

/**
 * Obtiene tendencias de ventas por período
 */
function obtenerTendenciasVentas(periodo = "mes") {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Registro_Movimientos");
  var data = sheet.getDataRange().getValues();
  
  var tendencias = {};
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[25] !== "Venta") continue; // Solo ventas
    
    var fecha = new Date(row[1]);
    var key = "";
    
    if (periodo === "mes") {
      key = fecha.getFullYear() + "-" + (fecha.getMonth() + 1);
    } else if (periodo === "semana") {
      var semana = Math.ceil((fecha.getDate() + fecha.getDay()) / 7);
      key = fecha.getFullYear() + "-W" + semana;
    } else if (periodo === "dia") {
      key = fecha.toDateString();
    }
    
    if (!tendencias[key]) {
      tendencias[key] = {
        ventas: 0,
        cantidad: 0,
        plataformas: {}
      };
    }
    
    tendencias[key].ventas++;
    tendencias[key].cantidad += parseInt(row[2]) || 0;
    
    var plataforma = row[26] || "Sin plataforma";
    if (!tendencias[key].plataformas[plataforma]) {
      tendencias[key].plataformas[plataforma] = 0;
    }
    tendencias[key].plataformas[plataforma]++;
  }
  
  return tendencias;
}

/**
 * Valida y limpia datos de entrada
 */
function validarDatosEntrada(datos) {
  var errores = [];
  
  // Validar campos obligatorios
  if (!datos.dispositivo || datos.dispositivo.trim() === "") {
    errores.push("El campo Dispositivo es obligatorio");
  }
  
  if (!datos.marca || datos.marca.trim() === "") {
    errores.push("El campo Marca es obligatorio");
  }
  
  if (!datos.modelo || datos.modelo.trim() === "") {
    errores.push("El campo Modelo es obligatorio");
  }
  
  // Validar cantidad
  if (datos.cantidad && (isNaN(datos.cantidad) || parseInt(datos.cantidad) < 0)) {
    errores.push("La cantidad debe ser un número positivo");
  }
  
  // Validar serie (si se proporciona)
  if (datos.serie && datos.serie.trim() !== "") {
    if (datos.serie.length < 3) {
      errores.push("La serie debe tener al menos 3 caracteres");
    }
    
    // Verificar si la serie ya existe
    if (!validarSerieUnica(datos.serie)) {
      errores.push("La serie ya existe en el inventario");
    }
  }
  
  return {
    valido: errores.length === 0,
    errores: errores
  };
}

/**
 * Obtiene sugerencias de autocompletado
 */
function obtenerSugerencias(campo, valor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Inventario");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var idxCampo = headers.indexOf(campo);
  if (idxCampo === -1) return [];
  
  var sugerencias = new Set();
  var valorLower = valor.toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    var valorCampo = data[i][idxCampo].toString();
    if (valorCampo.toLowerCase().indexOf(valorLower) !== -1) {
      sugerencias.add(valorCampo);
    }
  }
  
  return Array.from(sugerencias).slice(0, 10); // Máximo 10 sugerencias
}

/**
 * Función de diagnóstico completo para revisar todas las hojas del spreadsheet
 */
function diagnosticarHojasCompletas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = ss.getSheets();
    var diagnostico = {
      nombreSpreadsheet: ss.getName(),
      url: ss.getUrl(),
      totalHojas: hojas.length,
      hojas: [],
      errores: []
    };
    
    // Revisar cada hoja
    hojas.forEach(function(hoja, index) {
      try {
        var nombreHoja = hoja.getName();
        var data = hoja.getDataRange().getValues();
        var headers = data.length > 0 ? data[0] : [];
        var totalFilas = data.length - 1;
        
        var infoHoja = {
          nombre: nombreHoja,
          totalFilas: totalFilas,
          totalColumnas: headers.length,
          headers: headers,
          tieneDatos: totalFilas > 0,
          filasConDatos: 0,
          columnasConDatos: []
        };
        
        // Contar filas con datos reales
        for (var i = 1; i < data.length; i++) {
          var fila = data[i];
          var tieneDatos = false;
          for (var j = 0; j < fila.length; j++) {
            if (fila[j] && fila[j].toString().trim() !== "") {
              tieneDatos = true;
              if (!infoHoja.columnasConDatos.includes(j)) {
                infoHoja.columnasConDatos.push(j);
              }
            }
          }
          if (tieneDatos) {
            infoHoja.filasConDatos++;
          }
        }
        
        // Análisis específico para la hoja Inventario
        if (nombreHoja === "Inventario") {
          var idxSerie = headers.indexOf("Serie");
          var idxDispositivo = headers.indexOf("Dispositivo");
          var idxMarca = headers.indexOf("Marca");
          var idxModelo = headers.indexOf("Modelo");
          var idxCantidad = headers.indexOf("Cantidad");
          
          infoHoja.analisisInventario = {
            columnaSerie: {
              indice: idxSerie,
              existe: idxSerie !== -1,
              nombre: idxSerie !== -1 ? headers[idxSerie] : null
            },
            columnaDispositivo: {
              indice: idxDispositivo,
              existe: idxDispositivo !== -1,
              nombre: idxDispositivo !== -1 ? headers[idxDispositivo] : null
            },
            columnaMarca: {
              indice: idxMarca,
              existe: idxMarca !== -1,
              nombre: idxMarca !== -1 ? headers[idxMarca] : null
            },
            columnaModelo: {
              indice: idxModelo,
              existe: idxModelo !== -1,
              nombre: idxModelo !== -1 ? headers[idxModelo] : null
            },
            columnaCantidad: {
              indice: idxCantidad,
              existe: idxCantidad !== -1,
              nombre: idxCantidad !== -1 ? headers[idxCantidad] : null
            }
          };
          
          // Muestras de datos de series
          var seriesEjemplo = [];
          var dispositivosEjemplo = [];
          var marcasEjemplo = [];
          
          for (var i = 1; i < Math.min(data.length, 11); i++) {
            var fila = data[i];
            if (idxSerie !== -1 && fila[idxSerie] && fila[idxSerie].toString().trim() !== "") {
              seriesEjemplo.push({
                valor: fila[idxSerie].toString(),
                fila: i + 1,
                dispositivo: idxDispositivo !== -1 ? (fila[idxDispositivo] || "") : "",
                marca: idxMarca !== -1 ? (fila[idxMarca] || "") : "",
                modelo: idxModelo !== -1 ? (fila[idxModelo] || "") : ""
              });
            }
            if (idxDispositivo !== -1 && fila[idxDispositivo] && fila[idxDispositivo].toString().trim() !== "") {
              dispositivosEjemplo.push(fila[idxDispositivo].toString());
            }
            if (idxMarca !== -1 && fila[idxMarca] && fila[idxMarca].toString().trim() !== "") {
              marcasEjemplo.push(fila[idxMarca].toString());
            }
          }
          
          infoHoja.ejemplos = {
            series: seriesEjemplo.slice(0, 5),
            dispositivos: dispositivosEjemplo.slice(0, 5),
            marcas: marcasEjemplo.slice(0, 5)
          };
        }
        
        diagnostico.hojas.push(infoHoja);
        
      } catch (error) {
        diagnostico.errores.push({
          hoja: hoja.getName(),
          error: error.toString()
        });
      }
    });
    
    return diagnostico;
    
  } catch (error) {
    return {
      error: "Error en diagnóstico: " + error.toString()
    };
  }
}