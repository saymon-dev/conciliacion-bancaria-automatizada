/**
 * Clase para procesar los archivos del Banco BCI, incluyendo la cartola y el libro mayor.
 * @class
 */
class ProcesarArchivosBCI {
  /**
   * Crea una instancia de ProcesarArchivosBCI.
   * @param {string} sheetId - ID de la hoja de cálculo cargada con los datos de la cartola y libro mayor.
   * @param {string} folderIdOutput - ID de la carpeta de destino donde se guardarán los resultados procesados.
   * @param {string} sheetDiasHabilesId - ID de la hoja de cálculo que contiene los días hábiles.
   */
  constructor(sheetId, folderIdOutput, sheetDiasHabilesId) {
    this.sheetId = sheetId;
    this.folderIdOutput = folderIdOutput;
    this.sheetDiasHabilesId = sheetDiasHabilesId;
  }

  /**
   * Procesa el archivo subido y genera los resultados.
   * @returns {Object} Resultado del procesamiento.
   */
  procesarArchivoSubidoBCI() {
    try {
      Logger.log('Iniciando el procesamiento del archivo BANCO BCI con ID: ' + this.sheetId);

      // Abrir el archivo subido como Google Sheet
      const spreadsheet = SpreadsheetApp.openById(this.sheetId);

      // Seleccionar las hojas relevantes (Cartola y Libro Mayor)
      const cartolaSheet = spreadsheet.getSheetByName('Cartola');
      const libroMayorSheet = spreadsheet.getSheetByName('Libro Mayor');

      // Verificar si las hojas existen
      if (!cartolaSheet || !libroMayorSheet) {
        throw new Error('No se encontraron las hojas "Cartola" o "Libro Mayor".');
      }

      // Obtener la lista de días hábiles desde la hoja de días hábiles
      const diasHabiles = this.obtenerDiasHabiles();
      Logger.log('Días hábiles obtenidos: ' + diasHabiles);

      // Procesar Cartola
      const columnasNecesariasCartola = this.procesarCartola(cartolaSheet);
      
      // Procesar Libro Mayor
      const columnasNecesariasLibroMayor = this.procesarLibroMayor(libroMayorSheet, diasHabiles);

      // Crear una nueva hoja para guardar los resultados procesados
      const resultado = this.guardarResultados(columnasNecesariasCartola, columnasNecesariasLibroMayor);
      
      return resultado;

    } catch (error) {
      Logger.log('Error al procesar el archivo: ' + error.message);
      return {
        success: false,
        message: 'Error al procesar el archivo: ' + error.message
      };
    }
  }

  /**
   * Procesa los datos de la cartola.
   * @param {Sheet} cartolaSheet - Hoja de cálculo de la cartola.
   * @returns {Array[]} Datos procesados de la cartola.
   */
  procesarCartola(cartolaSheet) {
    const lastRowCartola = cartolaSheet.getLastRow();
    const cartolaData = cartolaSheet.getRange('A3:K' + lastRowCartola).getValues();
    const cartolaCleanData = this.reemplazarNaNConCeroSoloNumerico(cartolaData, [9, 10]);

    return cartolaCleanData.map((row) => {
      const descripcion = row[5];  // Descripción
      const nDocumento = row[7];   // N° Documento
      const chequesCargos = this.normalizeMonto(row[9]);  // Cargo
      const depositosAbonos = this.normalizeMonto(row[10]); // Abono
      return [row[0], descripcion, nDocumento, chequesCargos, depositosAbonos];  // Mapeo correcto de columnas
    });
  }

  /**
   * Procesa los datos del libro mayor.
   * @param {Sheet} libroMayorSheet - Hoja de cálculo del libro mayor.
   * @param {Array<Date>} diasHabiles - Lista de días hábiles.
   * @returns {Array[]} Datos procesados del libro mayor.
   */
  procesarLibroMayor(libroMayorSheet, diasHabiles) {
    const lastRowLibroMayor = libroMayorSheet.getLastRow();
    const libroMayorData = libroMayorSheet.getRange('A10:K' + lastRowLibroMayor).getValues();
    const libroMayorCleanData = this.reemplazarNaNConCeroSoloNumerico(libroMayorData, [8, 9]);

    return libroMayorCleanData.map((row) => {
      const fecha = this.convertirFecha(row[2]);  // Fecha
      if (!fecha) return ['Fecha Inválida', row[5], '', 0, 0, 0, 0, 'Fecha Inválida'];

      const glosa = row[5];  // Glosa Comprobante Detalle
      const rutExtraido = this.extractRutLibroMayor(glosa) || 0;
      const nombreExtraido = this.extraerNombreLibroMayor(glosa) || 0;
      const debe = this.normalizeMonto(row[8]);  // Debe
      const haber = this.normalizeMonto(row[9]);  // Haber
      const numeroDocumento = this.extractDocumentoLibroMayor(glosa) || '0';  // Extraer número de documento
      const siguienteDia = this.siguienteDiaHabil(fecha, diasHabiles);

      return [fecha, glosa, numeroDocumento, debe, haber, rutExtraido, nombreExtraido, siguienteDia];
    });
  }

  /**
   * Guarda los resultados procesados en un nuevo archivo.
   * @param {Array[]} columnasNecesariasCartola - Datos procesados de la cartola.
   * @param {Array[]} columnasNecesariasLibroMayor - Datos procesados del libro mayor.
   * @returns {Object} Resultado del procesamiento y URL del archivo generado.
   */
  guardarResultados(columnasNecesariasCartola, columnasNecesariasLibroMayor) {
    const fechaActual = new Date();
    const sufijo = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'ddMMyyyy_HHmmss');
    const outputSS = SpreadsheetApp.create(SpreadsheetApp.openById(this.sheetId).getName() + '_Procesado_BCI_' + sufijo);
    const processedSheetId = outputSS.getId();

    const hojaCartola = outputSS.getSheetByName('Hoja 1');
    hojaCartola.setName('Cartola');
    hojaCartola.appendRow(['Fecha', 'Descripcion', 'N° Documento', 'ChequesCargos', 'DepositosAbono']);
    hojaCartola.getRange(2, 1, columnasNecesariasCartola.length, columnasNecesariasCartola[0].length).setValues(columnasNecesariasCartola);

    const hojaLibroMayor = outputSS.insertSheet('Libro Mayor');
    hojaLibroMayor.appendRow(['Fecha', 'GlosaComprobanteDetalle', 'NumeroDocumento', 'Debe', 'Haber', 'RutExtraido', 'NombreExtraido', 'SiguienteDiaHabil']);
    hojaLibroMayor.getRange(2, 1, columnasNecesariasLibroMayor.length, columnasNecesariasLibroMayor[0].length).setValues(columnasNecesariasLibroMayor);

    const file = DriveApp.getFileById(outputSS.getId());
    const folder = DriveApp.getFolderById(this.folderIdOutput);
    folder.addFile(file);
    Logger.log('Archivo movido a la carpeta de procesados correctamente.');

    const processedFileUrl = outputSS.getUrl();
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

  /**
   * Obtiene los días hábiles desde una hoja de cálculo.
   * @returns {Array<Date>} Lista de días hábiles.
   */
  obtenerDiasHabiles() {
    const sheetDiasHabiles = SpreadsheetApp.openById(this.sheetDiasHabilesId).getSheetByName('DiasHabiles2024');
    const lastRow = sheetDiasHabiles.getLastRow();
    const diasHabilesRange = sheetDiasHabiles.getRange('A2:A' + lastRow);
    const diasHabiles = diasHabilesRange.getValues().flat().map((fecha) => this.convertirFecha(fecha));
    return diasHabiles.filter(Boolean);  // Retornar solo fechas válidas
  }

  /**
   * Convierte una fecha en el formato correcto.
   * @param {string|Date} fecha - Fecha a convertir.
   * @returns {Date|null} Fecha convertida o null si es inválida.
   */
  convertirFecha(fecha) {
    if (fecha instanceof Date) {
      return fecha;
    }
    let nuevaFecha = new Date(fecha);
    if (isNaN(nuevaFecha.getTime())) {
      const partes = fecha.split('/');
      if (partes.length === 3) {
        nuevaFecha = new Date(partes[2], partes[1] - 1, partes[0]);  // Ajustar el formato DD/MM/AAAA
      }
    }
    return isNaN(nuevaFecha.getTime()) ? null : nuevaFecha;
  }

  /**
   * Encuentra el siguiente día hábil basado en una fecha y una lista de días hábiles.
   * @param {Date} fecha - Fecha actual.
   * @param {Array<Date>} diasHabiles - Lista de días hábiles.
   * @returns {string} Siguiente día hábil en formato 'dd/MM/yyyy' o 'No Disponible' si no hay.
   */
  siguienteDiaHabil(fecha, diasHabiles) {
    if (!fecha || isNaN(fecha.getTime())) {
      return 'Fecha Inválida';
    }
    diasHabiles.sort((a, b) => a - b);
    for (let i = 0; i < diasHabiles.length; i++) {
      const diaHabil = diasHabiles[i];
      if (diaHabil > fecha) {
        return Utilities.formatDate(diaHabil, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
    }
    return 'No Disponible';
  }

  /**
   * Extrae el RUT de la glosa del libro mayor.
   * @param {string} glosa - Glosa del libro mayor.
   * @returns {string|null} RUT extraído o null si no se encuentra.
   */
  extractRutLibroMayor(glosa) {
    if (typeof glosa === 'string' && glosa.includes('/')) {
      const splitString = glosa.split('/');
      if (splitString.length > 2) {
        const rut = splitString[2].trim();
        if (/^\d{7,8}[kK0-9]$/.test(rut)) {
          return rut.replace('-', '').toUpperCase();
        }
      }
    }
    return null;
  }

  /**
   * Extrae el nombre de la glosa del libro mayor.
   * @param {string} glosa - Glosa del libro mayor.
   * @returns {string} Nombre extraído o cadena vacía si no se encuentra.
   */
  extraerNombreLibroMayor(glosa) {
    if (typeof glosa === 'string') {
      const patron = /TRANSFER\s+([A-Z\s]+)(?:\/\d+)?/;
      const coincidencia = glosa.match(patron);
      if (coincidencia) {
        return coincidencia[1].trim();
      }
    }
    return '';
  }

  /**
   * Extrae el número de documento del libro mayor utilizando diferentes patrones.
   * @param {string} glosa - Glosa del libro mayor.
   * @returns {string|null} Número de documento extraído o null si no se encuentra.
   */
  extractDocumentoLibroMayor(glosa) {
    if (typeof glosa === 'string') {
      // Patrón estándar con '/'
      const splitString = glosa.split('/');
      if (splitString.length > 1) {
        const nDocumento = splitString[1].trim();
        if (/^\d+$/.test(nDocumento)) {
          return nDocumento;
        }
      }

      // Patrón para "CARGO POR PAGO NOMINA EN LINEA-XXXXXXX"
      const patronDocumento = /-\s*(\d{7,8})$/;
      const coincidencia = glosa.match(patronDocumento);
      if (coincidencia) {
        return coincidencia[1];
      }
    }
    return null;
  }

  /**
   * Normaliza un monto convirtiéndolo en un número.
   * @param {string|number} monto - Monto a normalizar.
   * @returns {number} Monto normalizado.
   */
  normalizeMonto(monto) {
    if (typeof monto === 'string') {
      return parseInt(monto.replace(/[^\d-]/g, ''), 10) || 0;
    }
    return monto || 0;
  }

  /**
   * Reemplaza NaN con cero en columnas numéricas.
   * @param {Array[]} data - Datos a procesar.
   * @param {Array<number>} columnasNumericas - Índices de las columnas numéricas.
   * @returns {Array[]} Datos procesados con NaN reemplazados por ceros.
   */
  reemplazarNaNConCeroSoloNumerico(data, columnasNumericas) {
    return data.map(fila => fila.map((celda, colIndex) => {
      if (columnasNumericas.includes(colIndex)) {
        return (isNaN(celda) || celda === "" || celda === null) ? 0 : celda;
      }
      return celda;
    }));
  }
}

/**
 * Función para iniciar el procesamiento de archivos del Banco BCI.
 * @param {string} sheetId - ID de la hoja de cálculo de entrada.
 * @returns {Object} Resultado del procesamiento.
 */
function ejecutarProcesamientoBCI(sheetId) {
  try {
    const folderIdOutput = '1r-6r3vhhxAWDKvSgARHxcNp2CuJS7SvX'; // Carpeta de destino (actualiza este valor)
    const sheetDiasHabilesId = '1ryzTtdCPEz-biB77Drn470D-nM6kCRSgl1MvYU4ud7o'; // Hoja de días hábiles

    const procesador = new ProcesarArchivosBCI(sheetId, folderIdOutput, sheetDiasHabilesId);
    return procesador.procesarArchivoSubidoBCI();
  } catch (error) {
    Logger.log('Error al procesar el archivo: ' + error.message);
    throw new Error('Error al procesar el archivo: ' + error.message);
  }
}
