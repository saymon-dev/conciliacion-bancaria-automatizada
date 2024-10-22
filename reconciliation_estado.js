/**
 * Clase para conciliar datos entre la cartola y el libro mayor del banco Estado.
 * @class
 */
class ConciliadorEstado {
  /**
   * Crea una instancia de ConciliadorEstado.
   * @param {Array[]} cartola - Datos de la cartola.
   * @param {Array[]} libroMayor - Datos del libro mayor.
   */
  constructor(cartola, libroMayor) {
    this.cartola = cartola;
    this.libroMayor = libroMayor;
    this.conciliados = [];
    this.pendientesCartola = [];
    this.pendientesLibroMayor = [];
    this.conciliadosLibroMayorIdx = new Set();
    this.conciliadosCartolaIdx = new Set();
    this.documentosConciliados = new Set();
    this.mapeoExacto = new Map();  // Mapeo para coincidencias exactas (RUT, fecha, cargo, abono)
    this.mapeoNombre = new Map();  // Mapeo para coincidencias por nombre (Nombre, cargo, abono)
    this.mapeoFechaCargoAbono = new Map(); // Mapeo para coincidencias por Fecha y Montos
  }

  /**
   * Formatea una fecha al formato 'dd/MM/yyyy'.
   * @param {string|Date} fecha - Fecha a formatear.
   * @returns {string} La fecha formateada.
   */
  formatearFecha(fecha) {
    return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  /**
   * Prepara los mapeos del libro mayor para realizar las conciliaciones.
   */
  prepararMapeoLibroMayor() {
    this.libroMayor.forEach((filaLibro, indexLibro) => {
      // Clave para coincidencia exacta: RUT, fecha, cargo, abono
      const claveExacta = `${filaLibro[5]}|${this.formatearFecha(filaLibro[0])}|${filaLibro[3]}|${filaLibro[4]}`;
      if (!this.mapeoExacto.has(claveExacta)) {
        this.mapeoExacto.set(claveExacta, []);
      }
      this.mapeoExacto.get(claveExacta).push({ indexLibro, filaLibro });

      // Clave para coincidencia por nombre: Nombre, cargo, abono
      const claveNombre = `${filaLibro[6]}|${filaLibro[3]}|${filaLibro[4]}`;
      if (!this.mapeoNombre.has(claveNombre)) {
        this.mapeoNombre.set(claveNombre, []);
      }
      this.mapeoNombre.get(claveNombre).push({ indexLibro, filaLibro });

      // Clave para coincidencia por fecha y montos: Fecha, cargo, abono
      const fechaLibroMayor = this.formatearFecha(filaLibro[0]);
      const siguienteDiaLibroMayor = this.formatearFecha(filaLibro[7]);

      const claveFechaCargoAbono = `${fechaLibroMayor}|${filaLibro[3]}|${filaLibro[4]}`;
      if (!this.mapeoFechaCargoAbono.has(claveFechaCargoAbono)) {
        this.mapeoFechaCargoAbono.set(claveFechaCargoAbono, []);
      }
      this.mapeoFechaCargoAbono.get(claveFechaCargoAbono).push({ indexLibro, filaLibro });

      const claveSiguienteDiaCargoAbono = `${siguienteDiaLibroMayor}|${filaLibro[3]}|${filaLibro[4]}`;
      if (!this.mapeoFechaCargoAbono.has(claveSiguienteDiaCargoAbono)) {
        this.mapeoFechaCargoAbono.set(claveSiguienteDiaCargoAbono, []);
      }
      this.mapeoFechaCargoAbono.get(claveSiguienteDiaCargoAbono).push({ indexLibro, filaLibro });
    });
  }

  /**
   * Realiza el proceso de conciliación entre los registros de la cartola y el libro mayor.
   */
  conciliarEstado() {
    try {
      this.prepararMapeoLibroMayor();
      Logger.log("Iniciando el proceso de conciliación...");

      this.cartola.forEach((filaCartola, indexCartola) => {
        if (this.conciliadosCartolaIdx.has(indexCartola)) {
          Logger.log(`Registro en la cartola en el índice ${indexCartola} ya conciliado, saltando...`);
          return;
        }

        const fechaCartola = this.formatearFecha(filaCartola[0]);
        const rutCartola = filaCartola[4];
        const cargoCartola = filaCartola[3];
        const abonoCartola = filaCartola[2];
        const nombreCartola = filaCartola[5];
        let conciliado = false;

        // Condición 1: Coincidencias exactas (RUT, fecha, cargo, abono)
        const claveExacta = `${rutCartola}|${fechaCartola}|${cargoCartola}|${abonoCartola}`;
        if (this.mapeoExacto.has(claveExacta)) {
          const coincidenciaExacta = this.mapeoExacto.get(claveExacta).find((item) => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
          if (coincidenciaExacta) {
            this.conciliarFilaEstado(filaCartola, coincidenciaExacta.filaLibro, coincidenciaExacta.indexLibro, indexCartola);
            conciliado = true;
          }
        }

        // Condición 2: Coincidencias por fecha y siguiente día hábil
        const claveSiguienteDia = `${rutCartola}|${this.formatearFecha(filaCartola[7])}|${cargoCartola}|${abonoCartola}`;
        if (!conciliado && this.mapeoExacto.has(claveSiguienteDia)) {
          const coincidenciaSiguienteDia = this.mapeoExacto.get(claveSiguienteDia).find((item) => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
          if (coincidenciaSiguienteDia) {
            this.conciliarFilaEstado(filaCartola, coincidenciaSiguienteDia.filaLibro, coincidenciaSiguienteDia.indexLibro, indexCartola);
            conciliado = true;
          }
        }

        // Condición 3: Coincidencias por nombre
        const claveNombre = `${nombreCartola}|${cargoCartola}|${abonoCartola}`;
        if (!conciliado && this.mapeoNombre.has(claveNombre)) {
          const coincidenciaPorNombre = this.mapeoNombre.get(claveNombre).find((item) => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
          if (coincidenciaPorNombre) {
            this.conciliarFilaEstado(filaCartola, coincidenciaPorNombre.filaLibro, coincidenciaPorNombre.indexLibro, indexCartola);
            conciliado = true;
          }
        }

        // Condición 4: Coincidencia por fecha (o siguiente día hábil), cargo y abono
        if (!conciliado) {
          const claveFechaCargoAbono = `${fechaCartola}|${cargoCartola}|${abonoCartola}`;
          const claveSiguienteDiaCargoAbono = `${this.formatearFecha(filaCartola[7])}|${cargoCartola}|${abonoCartola}`;

          const coincidenciasPorFecha = this.mapeoFechaCargoAbono.get(claveFechaCargoAbono) || [];
          const coincidenciasPorSiguienteDia = this.mapeoFechaCargoAbono.get(claveSiguienteDiaCargoAbono) || [];

          const coincidenciaPorFecha = [...coincidenciasPorFecha, ...coincidenciasPorSiguienteDia].find((item) => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
          if (coincidenciaPorFecha) {
            this.conciliarFilaEstado(filaCartola, coincidenciaPorFecha.filaLibro, coincidenciaPorFecha.indexLibro, indexCartola);
            conciliado = true;
          }
        }

        // Si no se ha conciliado, agregar a pendientes
        if (!conciliado) {
          Logger.log(`El registro de la Cartola en el índice ${indexCartola} no fue conciliado. Agregando a pendientes.`);
          this.pendientesCartola.push([
            this.formatearFecha(filaCartola[0]),
            filaCartola[1].toUpperCase(),
            filaCartola[4],
            filaCartola[3],
            filaCartola[2]
          ]);
        }
      });

      // Llenar la lista de pendientes del libro mayor (los que no fueron conciliados)
      this.libroMayor.forEach((filaLibro, indexLibro) => {
        if (!this.conciliadosLibroMayorIdx.has(indexLibro)) {
          Logger.log(`Registro en el Libro Mayor no conciliado en el índice ${indexLibro}. Agregando a pendientes del Libro Mayor.`);
          this.pendientesLibroMayor.push([
            this.formatearFecha(filaLibro[0]),
            filaLibro[1].toUpperCase(),
            filaLibro[5],
            filaLibro[3],
            filaLibro[4]
          ]);
        }
      });

      Logger.log("Conciliación completada para todos los registros de la cartola.");
    } catch (error) {
      Logger.log(`Error durante la conciliación: ${error.message}`);
      throw new Error('Error durante la conciliación');
    }
  }

  /**
   * Conciliar una fila entre la cartola y el libro mayor.
   * @param {Array} filaCartola - Fila de la cartola.
   * @param {Array} filaLibroMayor - Fila del libro mayor.
   * @param {number} idxLibroMayor - Índice de la fila en el libro mayor.
   * @param {number} idxCartola - Índice de la fila en la cartola.
   */
  conciliarFilaEstado(filaCartola, filaLibroMayor, idxLibroMayor, idxCartola) {
    Logger.log(`Conciliando fila de Cartola en el índice ${idxCartola} con el Libro Mayor en el índice ${idxLibroMayor}`);
    const registroConciliado = [
      this.formatearFecha(filaCartola[0]), 
      filaCartola[1].toUpperCase(),
      filaLibroMayor[2],
      filaCartola[4],
      filaCartola[3],
      filaCartola[2]
    ];

    this.conciliados.push(registroConciliado);
    this.conciliadosCartolaIdx.add(idxCartola);
    this.conciliadosLibroMayorIdx.add(idxLibroMayor);
    this.documentosConciliados.add(filaLibroMayor[2]);
  }

  /**
   * Genera el archivo de salida con los datos conciliados y pendientes.
   * @returns {string} URL del archivo generado.
   */
  generarArchivoSalidaEstado() {
    try {
      const fechaActual = new Date();
      const sufijoFechaHora = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      const nombreArchivo = `Conciliación_Bancaria_Estado_${sufijoFechaHora}`;
      
      const hojaSalida = SpreadsheetApp.create(nombreArchivo);
      const file = DriveApp.getFileById(hojaSalida.getId());
      const folder = DriveApp.getFolderById('1hvJkvXPHBORiNAtW0pWIk5n_hQmxdlE_');
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);  // Eliminar de la carpeta raíz de Drive
      
      const hojaConciliados = hojaSalida.getActiveSheet();
      hojaConciliados.setName('Conciliados');
      
      let datosConciliados = [['FECHA', 'DETALLE', 'N°DOCUMENTO', 'RUT', 'CARGO', 'ABONO']];
      this.conciliados.forEach(registro => datosConciliados.push(registro));
      hojaConciliados.getRange(1, 1, datosConciliados.length, datosConciliados[0].length).setValues(datosConciliados);

      const hojaPendientesCartola = hojaSalida.insertSheet('Pendientes Cartola');
      let datosPendientesCartola = [['FECHA', 'DETALLE', 'RUT', 'CARGO', 'ABONO']];
      this.pendientesCartola.forEach(registro => datosPendientesCartola.push(registro));
      hojaPendientesCartola.getRange(1, 1, datosPendientesCartola.length, datosPendientesCartola[0].length).setValues(datosPendientesCartola);

      const hojaPendientesLibroMayor = hojaSalida.insertSheet('Pendientes Libro Mayor');
      let datosPendientesLibroMayor = [['FECHA', 'DETALLE', 'RUT', 'CARGO', 'ABONO']];
      this.pendientesLibroMayor.forEach(registro => datosPendientesLibroMayor.push(registro));
      hojaPendientesLibroMayor.getRange(1, 1, datosPendientesLibroMayor.length, datosPendientesLibroMayor[0].length).setValues(datosPendientesLibroMayor);

      const urlArchivoSalida = hojaSalida.getUrl();
      Logger.log(`Archivo de salida generado: ${urlArchivoSalida}`);
      return urlArchivoSalida;
    } catch (error) {
      Logger.log('Error al generar el archivo de salida: ' + error.message);
      throw new Error('Error al generar el archivo de salida');
    }
  }
}

/**
 * Función para iniciar la conciliación desde el botón.
 * @param {string} sheetId - ID de la hoja de cálculo cargada.
 * @returns {Object} Resultado de la conciliación.
 */
function iniciarConciliacionDesdeBotonEstado(sheetId) {
  try {
    Logger.log('Iniciando conciliación con el archivo ID: ' + sheetId);

    const sheetProcesado = SpreadsheetApp.openById(sheetId);
    const cartolaSheet = sheetProcesado.getSheetByName('Cartola');
    const libroMayorSheet = sheetProcesado.getSheetByName('Libro Mayor');

    if (!cartolaSheet || !libroMayorSheet) {
      throw new Error("No se encontraron las hojas 'Cartola' o 'Libro Mayor' en el archivo.");
    }

    const cartolaData = cartolaSheet.getRange(2, 1, cartolaSheet.getLastRow() - 1, 6).getValues();
    const libroMayorData = libroMayorSheet.getRange(2, 1, libroMayorSheet.getLastRow() - 1, 8).getValues();

    if (!cartolaData.length || !libroMayorData.length) {
      throw new Error("Datos vacíos en 'Cartola' o 'Libro Mayor'.");
    }

    const conciliador = new ConciliadorEstado(cartolaData, libroMayorData);
    conciliador.conciliarEstado();

    const urlArchivoSalida = conciliador.generarArchivoSalidaEstado();
    Logger.log('URL Archivo Salida (Banco Estado): ' + urlArchivoSalida); // Se agrega para verificar la URL
    const emailUsuario = Session.getActiveUser().getEmail();
    const totalConciliados = conciliador.conciliados.length;
    const totalPendientesCartola = conciliador.pendientesCartola.length;
    const totalPendientesLibroMayor = conciliador.pendientesLibroMayor.length;
    const totalRegistrosCartola = cartolaData.length;
    const porcentajeConciliados = ((totalConciliados / totalRegistrosCartola) * 100).toFixed(2);

    const asunto = 'Conciliación BANCO ESTADO - Completada';
    const mensaje = `
      Estimado/a, <br><br>
      La conciliación bancaria para <strong>Banco Estado</strong> ha sido completada exitosamente. <br><br>
      <strong>Resultado:</strong> <br>
      Total de registros conciliados: ${totalConciliados} <br>
      Total de registros pendientes de la cartola: ${totalPendientesCartola} <br>
      Total de registros pendientes del libro mayor: ${totalPendientesLibroMayor} <br>
      <strong>Porcentaje de conciliación: ${porcentajeConciliados}%</strong> <br><br>
      Puede descargar el archivo de conciliación desde el siguiente enlace: <br>
      <a href="${urlArchivoSalida}" target="_blank">Ver archivo conciliado</a> <br><br>
      Saludos cordiales, <br>
      El equipo de Conciliación Bancaria.
    `;

    MailApp.sendEmail({
      to: emailUsuario,
      subject: asunto,
      htmlBody: mensaje
    });

    Logger.log('Correo enviado a: ' + emailUsuario + ' para Banco Estado');

    return {
      success: true,
      message: 'Conciliación completada exitosamente. Se ha enviado un correo a ' + emailUsuario,
      conciledFileUrl: urlArchivoSalida,
      totalConciliados: totalConciliados,
      totalPendientesCartola: totalPendientesCartola,
      totalPendientesLibroMayor: totalPendientesLibroMayor,
      porcentajeConciliados: porcentajeConciliados
    };
  } catch (error) {
    Logger.log('Error al iniciar la conciliación: ' + error.message);
    return {
      success: false,
      message: 'Error al iniciar la conciliación: ' + error.message
    };
  }
}
