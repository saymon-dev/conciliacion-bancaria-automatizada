class ConciliadorEstado {
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

    formatearFecha(fecha) {
        return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }

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

            // Clave para fecha exacta, cargo, abono
            const claveFechaCargoAbono = `${fechaLibroMayor}|${filaLibro[3]}|${filaLibro[4]}`;
            if (!this.mapeoFechaCargoAbono.has(claveFechaCargoAbono)) {
                this.mapeoFechaCargoAbono.set(claveFechaCargoAbono, []);
            }
            this.mapeoFechaCargoAbono.get(claveFechaCargoAbono).push({ indexLibro, filaLibro });

            // Clave para siguiente día hábil, cargo, abono
            const claveSiguienteDiaCargoAbono = `${siguienteDiaLibroMayor}|${filaLibro[3]}|${filaLibro[4]}`;
            if (!this.mapeoFechaCargoAbono.has(claveSiguienteDiaCargoAbono)) {
                this.mapeoFechaCargoAbono.set(claveSiguienteDiaCargoAbono, []);
            }
            this.mapeoFechaCargoAbono.get(claveSiguienteDiaCargoAbono).push({ indexLibro, filaLibro });
        });
    }


    conciliarEstado() {
        this.prepararMapeoLibroMayor(); // Preparar los mapeos antes de conciliar
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
                let coincidenciaExacta = this.mapeoExacto.get(claveExacta).find((item) => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
                if (coincidenciaExacta) {
                    this.conciliarFilaEstado(filaCartola, coincidenciaExacta.filaLibro, coincidenciaExacta.indexLibro, indexCartola);
                    conciliado = true;
                }
            }

            // Condición 2: Coincidencias por fecha y siguiente día hábil
            const claveSiguienteDia = `${rutCartola}|${this.formatearFecha(filaCartola[7])}|${cargoCartola}|${abonoCartola}`;
            if (!conciliado && this.mapeoExacto.has(claveSiguienteDia)) {
                let coincidenciaSiguienteDia = this.mapeoExacto.get(claveSiguienteDia).find((item) => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
                if (coincidenciaSiguienteDia) {
                    this.conciliarFilaEstado(filaCartola, coincidenciaSiguienteDia.filaLibro, coincidenciaSiguienteDia.indexLibro, indexCartola);
                    conciliado = true;
                }
            }

            // Condición 3: Coincidencias por nombre
            const claveNombre = `${nombreCartola}|${cargoCartola}|${abonoCartola}`;
            if (!conciliado && this.mapeoNombre.has(claveNombre)) {
                let coincidenciaPorNombre = this.mapeoNombre.get(claveNombre).find((item) => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
                if (coincidenciaPorNombre) {
                    this.conciliarFilaEstado(filaCartola, coincidenciaPorNombre.filaLibro, coincidenciaPorNombre.indexLibro, indexCartola);
                    conciliado = true;
                }
            }
            
            // Condición 4: Coincidencia por fecha (o siguiente día hábil), cargo y abono
            if (!conciliado) {
                const claveFechaCargoAbono = `${fechaCartola}|${cargoCartola}|${abonoCartola}`;
                const claveSiguienteDiaCargoAbono = `${this.formatearFecha(filaCartola[7])}|${cargoCartola}|${abonoCartola}`;

                let coincidenciasPorFecha = this.mapeoFechaCargoAbono.get(claveFechaCargoAbono) || [];
                let coincidenciasPorSiguienteDia = this.mapeoFechaCargoAbono.get(claveSiguienteDiaCargoAbono) || [];

                let coincidenciaPorFecha = [...coincidenciasPorFecha, ...coincidenciasPorSiguienteDia].find((item) => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
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
    }

    conciliarFilaEstado(filaCartola, filaLibroMayor, idxLibroMayor, idxCartola) {
        Logger.log(`Conciliando fila de Cartola en el índice ${idxCartola} con el Libro Mayor en el índice ${idxLibroMayor}`);
        let registroConciliado = [
            this.formatearFecha(filaCartola[0]),  
            filaCartola[1].toUpperCase(),  
            filaLibroMayor[2],  
            filaCartola[4],  
            filaCartola[3],  
            filaCartola[2]   
        ];

        // Agregar registro a conciliados
        this.conciliados.push(registroConciliado);
        this.conciliadosCartolaIdx.add(idxCartola);
        this.conciliadosLibroMayorIdx.add(idxLibroMayor);
        this.documentosConciliados.add(filaLibroMayor[2]);
    }

  generarArchivoSalidaEstado() {
      try {
          const fechaActual = new Date();
          const sufijoFechaHora = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
          const nombreArchivo = `Conciliación_Bancaria_Estado_${sufijoFechaHora}`;
          
          // Crear la hoja en blanco
          const hojaSalida = SpreadsheetApp.create(nombreArchivo);
          
          // Mover el archivo a la carpeta específica
          const file = DriveApp.getFileById(hojaSalida.getId());
          const folder = DriveApp.getFolderById('1JIwu5inEVZ9MAl4pdw5QdQcd2NJ2a-Xu');
          folder.addFile(file);
          DriveApp.getRootFolder().removeFile(file);  // Eliminar de la carpeta raíz de Drive
          
          const hojaConciliados = hojaSalida.getActiveSheet();
          hojaConciliados.setName('Conciliados');
          
          // Crear los encabezados y escribir los datos en un solo bloque
          let datosConciliados = [['FECHA', 'DETALLE', 'N°DOCUMENTO', 'RUT', 'CARGO', 'ABONO']];
          this.conciliados.forEach(registro => datosConciliados.push(registro));
          hojaConciliados.getRange(1, 1, datosConciliados.length, datosConciliados[0].length).setValues(datosConciliados);

          // Crear una nueva hoja para los pendientes de la cartola
          const hojaPendientesCartola = hojaSalida.insertSheet('Pendientes Cartola');
          let datosPendientesCartola = [['FECHA', 'DETALLE', 'RUT', 'CARGO', 'ABONO']];
          this.pendientesCartola.forEach(registro => datosPendientesCartola.push(registro));
          hojaPendientesCartola.getRange(1, 1, datosPendientesCartola.length, datosPendientesCartola[0].length).setValues(datosPendientesCartola);

          // Crear una nueva hoja para los pendientes del libro mayor
          const hojaPendientesLibroMayor = hojaSalida.insertSheet('Pendientes Libro Mayor');
          let datosPendientesLibroMayor = [['FECHA', 'DETALLE', 'RUT', 'CARGO', 'ABONO']];
          this.pendientesLibroMayor.forEach(registro => datosPendientesLibroMayor.push(registro));
          hojaPendientesLibroMayor.getRange(1, 1, datosPendientesLibroMayor.length, datosPendientesLibroMayor[0].length).setValues(datosPendientesLibroMayor);

          // Obtener el URL del archivo generado
          const urlArchivoSalida = hojaSalida.getUrl();
          Logger.log(`Archivo de salida generado: ${urlArchivoSalida}`);
          
          return urlArchivoSalida;
      } catch (error) {
          Logger.log('Error al generar el archivo de salida: ' + error.message);
          throw new Error('Error al generar el archivo de salida');
      }
  }

}

function iniciarConciliacionDesdeBotonEstado(sheetId) {
    try {
        Logger.log('Iniciando conciliación con el archivo ID: ' + sheetId);

        let sheetProcesado = SpreadsheetApp.openById(sheetId);

        // Obtener las hojas de Cartola y Libro Mayor
        let cartolaSheet = sheetProcesado.getSheetByName('Cartola');
        let libroMayorSheet = sheetProcesado.getSheetByName('Libro Mayor');

        if (!cartolaSheet || !libroMayorSheet) {
            throw new Error("No se encontraron las hojas 'Cartola' o 'Libro Mayor' en el archivo.");
        }

        // Obtener los datos de Cartola y Libro Mayor
        let cartolaData = cartolaSheet.getRange(2, 1, cartolaSheet.getLastRow() - 1, 6).getValues();
        let libroMayorData = libroMayorSheet.getRange(2, 1, libroMayorSheet.getLastRow() - 1, 8).getValues();

        if (!cartolaData.length || !libroMayorData.length) {
            throw new Error("Datos vacíos en 'Cartola' o 'Libro Mayor'.");
        }

        let conciliador = new ConciliadorEstado(cartolaData, libroMayorData);
        conciliador.conciliarEstado();

        // Generar el archivo de salida y obtener el enlace
        let urlArchivoSalida = conciliador.generarArchivoSalidaEstado();

        // Obtener correo del usuario activo
        let emailUsuario = Session.getActiveUser().getEmail();
        let totalConciliados = conciliador.conciliados.length;
        let totalPendientesCartola = conciliador.pendientesCartola.length;
        let totalPendientesLibroMayor = conciliador.pendientesLibroMayor.length;
        let totalRegistrosCartola = cartolaData.length;

        // Calcular el porcentaje de registros conciliados
        let porcentajeConciliados = ((totalConciliados / totalRegistrosCartola) * 100).toFixed(2);

        // Enviar correo con el resumen de la conciliación
        try {
            let asunto = 'Conciliación Bancaria Completada';
            let mensaje = `
                Estimado/a, <br><br>
                La conciliación bancaria ha sido completada exitosamente. <br><br>
                <strong>Resultado:</strong> <br>
                Total de registros conciliados: ${totalConciliados} <br>
                Total de registros pendientes de la cartola: ${totalPendientesCartola} <br>
                Total de registros pendientes del libro mayor: ${totalPendientesLibroMayor} <br>
                <strong>Porcentaje de conciliación: ${porcentajeConciliados}%</strong> <br><br>
                Puede descargar el archivo de conciliación desde el siguiente enlace: <br>
                <a href="${urlArchivoSalida}" target="_blank">Ver archivo conciliado</a> <br><br>
                Saludos, <br>
                El equipo de Conciliación Bancaria.
            `;

            MailApp.sendEmail({
                to: emailUsuario,
                subject: asunto,
                htmlBody: mensaje
            });
            Logger.log('Correo enviado a: ' + emailUsuario);

        } catch (error) {
            Logger.log('Error al enviar el correo: ' + error.message);
            throw new Error('Error al enviar el correo');
        }

        // Retornar el número de registros conciliados, pendientes y la URL del archivo conciliado
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

