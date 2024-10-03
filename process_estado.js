class ProcesarArchivosEstado {
  constructor(sheetId, folderIdOutput, sheetDiasHabilesId) {
    this.sheetId = sheetId;
    this.folderIdOutput = folderIdOutput;
    this.sheetDiasHabilesId = sheetDiasHabilesId;
  }

  procesarArchivoSubidoEstado() {
    try {
      Logger.log('Iniciando el procesamiento del archivo BANCO ESTADO con ID: ' + this.sheetId);

      // Abrir el archivo subido como Google Sheet
      var spreadsheet = SpreadsheetApp.openById(this.sheetId);

      // Seleccionar las hojas relevantes (Cartola y Libro Mayor)
      var cartolaSheet = spreadsheet.getSheetByName('Cartola');
      var libroMayorSheet = spreadsheet.getSheetByName('Libro Mayor');

      // Verificar si las hojas existen
      if (!cartolaSheet || !libroMayorSheet) {
        throw new Error('No se encontraron las hojas "Cartola" o "Libro Mayor".');
      }

      // Obtener la lista de días hábiles desde la hoja de días hábiles
      var diasHabiles = this.obtenerDiasHabiles();
      Logger.log('Días hábiles obtenidos: ' + diasHabiles);

      // Obtener el número de filas con datos en las hojas de Cartola y Libro Mayor
      var lastRowCartola = cartolaSheet.getLastRow();
      var lastRowLibroMayor = libroMayorSheet.getLastRow();

      // Obtener los datos de la Cartola
      var cartolaData = cartolaSheet.getRange('A3:J' + lastRowCartola).getValues();
      cartolaData = this.reemplazarNaNConCeroSoloNumerico(cartolaData, [7, 8]); // Reemplazar NaN por 0

      // Procesar Cartola
      var columnasNecesariasCartola = cartolaData.map((row) => {
        var descripcion = row[6];  // Columna G: Descripción
        var rutExtraido = this.extractRutCartola(descripcion) || 0;
        var nombreExtraido = this.extraerNombreCartola(descripcion) || 0;
        var chequesCargos = this.normalizeMonto(row[7]);  // Normalizar Cheques/Cargos
        var depositosAbonos = this.normalizeMonto(row[8]);  // Normalizar Depósitos/Abonos
        return [row[0], descripcion, chequesCargos, depositosAbonos, rutExtraido, nombreExtraido];  // Fecha, Descripción, Cheques/Cargos, Depósitos/Abonos, RUT, Nombre
      });

      // Obtener los datos del Libro Mayor
      var libroMayorData = libroMayorSheet.getRange('A10:M' + lastRowLibroMayor).getValues();
      libroMayorData = this.reemplazarNaNConCeroSoloNumerico(libroMayorData, [8, 9]);  // Reemplazar NaN por 0

      // Procesar Libro Mayor
      var columnasNecesariasLibroMayor = libroMayorData.map((row, index) => {
        Logger.log('Procesando fila: ' + index);
        
        var fecha = this.convertirFecha(row[2]);  // Columna C: Fecha
        if (!fecha) {
          return ['Fecha Inválida', row[5], row[7], this.normalizeMonto(row[8]), this.normalizeMonto(row[9]), 0, 0, 'Fecha Inválida'];
        }

        var glosa = row[5];  // Columna F: Glosa comprobante detalle
        var rutExtraido = this.extractRutLibroMayor(glosa) || 0;
        var nombreExtraido = this.extraerNombreLibroMayor(glosa) || 0;
        var debe = this.normalizeMonto(row[8]);  // Columna I: Debe
        var haber = this.normalizeMonto(row[9]);  // Columna J: Haber
        var numeroDocumento = row[7];  // Columna H: Número de Documento
        var siguienteDia = this.siguienteDiaHabil(fecha, diasHabiles);

        return [fecha, glosa, numeroDocumento, debe, haber, rutExtraido, nombreExtraido, siguienteDia];
      });

      // Crear una nueva hoja para guardar los resultados procesados
      var fechaActual = new Date();
      var sufijo = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'ddMMyyyy_HHmmss');
      var outputSS = SpreadsheetApp.create(spreadsheet.getName() + '_Procesado_' + sufijo);

      // Guardar el SheetId del archivo procesado
      var processedSheetId = outputSS.getId();

      // Crear hoja para Cartola y agregar encabezados + datos procesados
      var hojaCartola = outputSS.getSheetByName('Hoja 1');
      hojaCartola.setName('Cartola');
      hojaCartola.appendRow(['Fecha', 'Descripcion', 'ChequesCargos', 'DepositosAbonos', 'RutExtraido', 'NombreExtraido']);
      hojaCartola.getRange(2, 1, columnasNecesariasCartola.length, columnasNecesariasCartola[0].length).setValues(columnasNecesariasCartola);

      // Crear hoja para Libro Mayor y agregar encabezados + datos procesados
      var hojaLibroMayor = outputSS.insertSheet('Libro Mayor');
      hojaLibroMayor.appendRow(['Fecha', 'GlosaComprobanteDetalle', 'NumeroDocumento', 'Debe', 'Haber', 'RutExtraido', 'NombreExtraido', 'SiguienteDiaHabil']);
      hojaLibroMayor.getRange(2, 1, columnasNecesariasLibroMayor.length, columnasNecesariasLibroMayor[0].length).setValues(columnasNecesariasLibroMayor);

      // Mover el archivo procesado a la carpeta de procesados en Google Drive
      var file = DriveApp.getFileById(outputSS.getId());
      var folder = DriveApp.getFolderById(this.folderIdOutput);
      folder.addFile(file);
      Logger.log('Archivo movido a la carpeta de procesados correctamente.');

      // Obtener el enlace al archivo procesado
      var processedFileUrl = outputSS.getUrl();

      // Enviar notificación por correo
      var asunto = 'Archivo procesado correctamente';
      var mensaje = 'El archivo ha sido procesado y está listo para ser conciliado. Puede acceder a él en el siguiente enlace: <a href="' + processedFileUrl + '" target="_blank">Ver archivo procesado</a>';
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: asunto,
        htmlBody: mensaje
      });

      Logger.log('Correo enviado a: ' + Session.getActiveUser().getEmail());

      return {
        success: true,
        message: 'Archivo procesado correctamente. Se ha enviado una notificación al correo: ' + Session.getActiveUser().getEmail(),
        processedFileUrl: processedFileUrl,
        processedSheetId: processedSheetId
      };
    } catch (error) {
      Logger.log('Error al procesar el archivo: ' + error.message);
      return {
        success: false,
        message: 'Error al procesar el archivo: ' + error.message
      };
    }
  }

  obtenerDiasHabiles() {
    var sheetDiasHabiles = SpreadsheetApp.openById(this.sheetDiasHabilesId).getSheetByName('DiasHabiles2024');
    var lastRow = sheetDiasHabiles.getLastRow();
    var diasHabilesRange = sheetDiasHabiles.getRange('A2:A' + lastRow);
    var diasHabiles = diasHabilesRange.getValues().flat().map((fecha) => this.convertirFecha(fecha));
    return diasHabiles.filter(Boolean);  // Retornar solo fechas válidas
  }

  convertirFecha(fecha) {
    if (fecha instanceof Date) {
      return fecha;  // Si ya es un objeto Date, retornarlo
    }

    var nuevaFecha = new Date(fecha);
    if (isNaN(nuevaFecha.getTime())) {
      var partes = fecha.split('/');
      if (partes.length === 3) {
        nuevaFecha = new Date(partes[2], partes[1] - 1, partes[0]);  // Ajustar el formato
      }
    }
    return isNaN(nuevaFecha.getTime()) ? null : nuevaFecha;
  }

  siguienteDiaHabil(fecha, diasHabiles) {
    if (!fecha || isNaN(fecha.getTime())) {
      return 'Fecha Inválida';
    }
    diasHabiles.sort((a, b) => a - b);
    for (var i = 0; i < diasHabiles.length; i++) {
      var diaHabil = diasHabiles[i];
      if (diaHabil > fecha) {
        return Utilities.formatDate(diaHabil, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
    }
    return 'No Disponible';
  }

  extractRutCartola(text) {
    if (typeof text === 'string') {
      var match = text.match(/\b\d{7,8}-[kK\d]\b/);
      if (match) {
        return match[0].replace('-', '').toUpperCase();
      }
    }
    return null;
  }

  extractRutLibroMayor(glosa) {
    if (typeof glosa === 'string') {
      glosa = glosa.toUpperCase();
      var partes = glosa.split('/');
      if (partes.length > 1) {
        var rut = partes[1].trim();
        var rutValido = rut.match(/^\d{7,8}[kK\d]$/);
        if (rutValido) {
          return rutValido[0].toUpperCase();
        }
      }
    }
    return null;
  }

  extraerNombreCartola(descripcion) {
    if (typeof descripcion === 'string') {
      descripcion = descripcion.toUpperCase();
      var patronConRut = /\b\d{7,8}-[kK\d]\s+(.+)/;
      var coincidencia = descripcion.match(patronConRut);
      if (coincidencia) {
        return coincidencia[1].trim();
      }
      var patronesSinRut = [/DE\s+([A-ZÁÉÍÓÚÑ\s]+)/, /AL\s+([A-ZÁÉÍÓÚÑ\s]+)/];
      for (var i = 0; i < patronesSinRut.length; i++) {
        coincidencia = descripcion.match(patronesSinRut[i]);
        if (coincidencia && !coincidencia[1].includes('RUT')) {
          return coincidencia[1].trim();
        }
      }
    }
    return '';
  }

  extraerNombreLibroMayor(glosa) {
    if (typeof glosa === 'string') {
      glosa = glosa.toUpperCase();
      var patron = /TRANSFER\s+([A-Z\s]+)(?:\/\d+)?/;
      var coincidencia = glosa.match(patron);
      if (coincidencia) {
        return coincidencia[1].trim();
      }
    }
    return '';
  }

  normalizeMonto(monto) {
    if (typeof monto === 'string') {
      return parseInt(monto.replace(/[^\d-]/g, ''), 10) || 0;
    }
    return monto || 0;
  }

  reemplazarNaNConCeroSoloNumerico(data, columnasNumericas) {
    return data.map(fila => fila.map((celda, colIndex) => {
      if (columnasNumericas.includes(colIndex)) {
        return (isNaN(celda) || celda === "" || celda === null) ? 0 : celda;
      }
      return celda;
    }));
  }
}


function ejecutarProcesamientoEstado(sheetId) {
  try {
    // Define el ID de la carpeta de salida y la hoja de días hábiles (ajusta si es necesario)
    var folderIdOutput = '1WxA_WUIXKyiBxE6_BLCCBUB-87bqiwWO';  // ID de la carpeta de salida
    var sheetDiasHabilesId = '1ryzTtdCPEz-biB77Drn470D-nM6kCRSgl1MvYU4ud7o';  // ID de la hoja de días hábiles

    // Crear una instancia de la clase ProcesarArchivosEstado
    var procesador = new ProcesarArchivosEstado(sheetId, folderIdOutput, sheetDiasHabilesId);

    // Llamar al método para procesar el archivo subido
    var resultado = procesador.procesarArchivoSubidoEstado();

    // Retornar el resultado al frontend
    return resultado;
    
  } catch (error) {
    Logger.log('Error al procesar el archivo: ' + error.message);
    throw new Error('Error al procesar el archivo: ' + error.message);
  }
}

