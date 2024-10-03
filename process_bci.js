class ProcesarArchivosBCI {
  constructor(sheetId, folderIdOutput, sheetDiasHabilesId) {
    this.sheetId = sheetId;
    this.folderIdOutput = folderIdOutput;
    this.sheetDiasHabilesId = sheetDiasHabilesId;
  }

  procesarArchivoSubidoBCI() {
    try {
      Logger.log('Iniciando el procesamiento del archivo BANCO BCI con ID: ' + this.sheetId);

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

      // Procesar Cartola
      var columnasNecesariasCartola = this.procesarCartola(cartolaSheet);
      
      // Procesar Libro Mayor
      var columnasNecesariasLibroMayor = this.procesarLibroMayor(libroMayorSheet, diasHabiles);

      // Crear una nueva hoja para guardar los resultados procesados
      var resultado = this.guardarResultados(columnasNecesariasCartola, columnasNecesariasLibroMayor);
      
      return resultado;

    } catch (error) {
      Logger.log('Error al procesar el archivo: ' + error.message);
      return {
        success: false,
        message: 'Error al procesar el archivo: ' + error.message
      };
    }
  }

  procesarCartola(cartolaSheet) {
    var lastRowCartola = cartolaSheet.getLastRow();
    var cartolaData = cartolaSheet.getRange('A3:K' + lastRowCartola).getValues();
    cartolaData = this.reemplazarNaNConCeroSoloNumerico(cartolaData, [8, 9]);

    return cartolaData.map((row) => {
      var descripcion = row[5];
      var nDocumento = row[7];
      var chequesCargos = this.normalizeMonto(row[9]);
      var depositosAbonos = this.normalizeMonto(row[10]);
      return [row[0], descripcion, nDocumento, chequesCargos, depositosAbonos];
    });
  }

  procesarLibroMayor(libroMayorSheet, diasHabiles) {
    var lastRowLibroMayor = libroMayorSheet.getLastRow();
    var libroMayorData = libroMayorSheet.getRange('A10:K' + lastRowLibroMayor).getValues();
    libroMayorData = this.reemplazarNaNConCeroSoloNumerico(libroMayorData, [8, 9]);

    return libroMayorData.map((row) => {
      var fecha = this.convertirFecha(row[2]);
      if (!fecha) return ['Fecha Inválida', row[5], row[7], this.normalizeMonto(row[8]), this.normalizeMonto(row[9]), 0, 0, 'Fecha Inválida'];

      var glosa = row[5];
      var rutExtraido = this.extractRutLibroMayor(glosa) || 0;
      var nombreExtraido = this.extraerNombreLibroMayor(glosa) || 0;
      var debe = this.normalizeMonto(row[8]);
      var haber = this.normalizeMonto(row[9]);
      var numeroDocumento = this.extractDocumentoLibroMayor(glosa) || '0';
      var siguienteDia = this.siguienteDiaHabil(fecha, diasHabiles);

      return [fecha, glosa, numeroDocumento, debe, haber, rutExtraido, nombreExtraido, siguienteDia];
    });
  }

  guardarResultados(columnasNecesariasCartola, columnasNecesariasLibroMayor) {
    var fechaActual = new Date();
    var sufijo = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'ddMMyyyy_HHmmss');
    var outputSS = SpreadsheetApp.create(SpreadsheetApp.openById(this.sheetId).getName() + '_Procesado_BCI_' + sufijo);
    var processedSheetId = outputSS.getId();

    var hojaCartola = outputSS.getSheetByName('Hoja 1');
    hojaCartola.setName('Cartola');
    hojaCartola.appendRow(['Fecha', 'Descripcion', 'N° Documento','ChequesCargos','DepositosAbono']);
    hojaCartola.getRange(2, 1, columnasNecesariasCartola.length, columnasNecesariasCartola[0].length).setValues(columnasNecesariasCartola);

    var hojaLibroMayor = outputSS.insertSheet('Libro Mayor');
    hojaLibroMayor.appendRow(['Fecha', 'GlosaComprobanteDetalle', 'NumeroDocumento', 'Debe', 'Haber', 'RutExtraido', 'NombreExtraido', 'SiguienteDiaHabil']);
    hojaLibroMayor.getRange(2, 1, columnasNecesariasLibroMayor.length, columnasNecesariasLibroMayor[0].length).setValues(columnasNecesariasLibroMayor);

    var file = DriveApp.getFileById(outputSS.getId());
    var folder = DriveApp.getFolderById(this.folderIdOutput);
    folder.addFile(file);
    Logger.log('Archivo movido a la carpeta de procesados correctamente.');

    var processedFileUrl = outputSS.getUrl();
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: 'Archivo procesado correctamente',
      htmlBody: 'El archivo ha sido procesado y está listo para ser conciliado. Puede acceder a él en el siguiente enlace: <a href="' + processedFileUrl + '" target="_blank">Ver archivo procesado</a>'
    });

    return {
      success: true,
      message: 'Archivo procesado correctamente. Se ha enviado una notificación al correo: ' + Session.getActiveUser().getEmail(),
      processedFileUrl: processedFileUrl,
      processedSheetId: processedSheetId
    };
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
      return fecha;
    }
    var nuevaFecha = new Date(fecha);
    if (isNaN(nuevaFecha.getTime())) {
      var partes = fecha.split('/');
      if (partes.length === 3) {
        nuevaFecha = new Date(partes[2], partes[1] - 1, partes[0]);  // Ajustar el formato DD/MM/AAAA
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

  extractRutLibroMayor(glosa) {
    if (typeof glosa === 'string' && glosa.includes('/')) {
      var splitString = glosa.split('/');
      if (splitString.length > 2) {
        var rut = splitString[2].trim();
        if (/^\d{7,8}[kK0-9]$/.test(rut)) {
          return rut.replace('-', '').toUpperCase();
        }
      }
    }
    return null;
  }

  extraerNombreLibroMayor(glosa) {
    if (typeof glosa === 'string') {
      var patron = /TRANSFER\s+([A-Z\s]+)(?:\/\d+)?/;
      var coincidencia = glosa.match(patron);
      if (coincidencia) {
        return coincidencia[1].trim();
      }
    }
    return '';
  }

  extractDocumentoLibroMayor(glosa) {
    if (typeof glosa === 'string' && glosa.includes('/')) {
      var splitString = glosa.split('/');
      if (splitString.length > 1) {
        var nDocumento = splitString[1].trim();
        if (/^\d+$/.test(nDocumento)) {
          return nDocumento;
        }
      }
    }
    return null;
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

// Función para iniciar el procesamiento en Banco BCI
function ejecutarProcesamientoBCI(sheetId) {
  try {
    var folderIdOutput = '1r-6r3vhhxAWDKvSgARHxcNp2CuJS7SvX'; // Actualiza este valor
    var sheetDiasHabilesId = '1ryzTtdCPEz-biB77Drn470D-nM6kCRSgl1MvYU4ud7o'; 

    var procesador = new ProcesarArchivosBCI(sheetId, folderIdOutput, sheetDiasHabilesId);
    return procesador.procesarArchivoSubidoBCI();
  } catch (error) {
    Logger.log('Error al procesar el archivo: ' + error.message);
    throw new Error('Error al procesar el archivo: ' + error.message);
  }
}
