class ConciliadorBCI {
    constructor(cartola, libroMayor) {
        this.cartola = cartola;
        this.libroMayor = libroMayor;
        this.conciliados = [];
        this.pendientesCartola = [];
        this.pendientesLibroMayor = [];
        this.conciliadosLibroMayorIdx = new Set();
        this.conciliadosCartolaIdx = new Set();
        this.mapeoExacto = new Map();  // Mapeo para coincidencias exactas (Documento, cargo, abono)
        this.mapeoNombre = new Map();  // Mapeo para coincidencias por nombre (Nombre, cargo, abono)
        this.mapeoFechaCargoAbono = new Map(); // Mapeo para coincidencias por Fecha y Montos
    }

    formatearFecha(fecha) {
        return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }

    prepararMapeoLibroMayor() {
        Logger.log("Iniciando preparación de mapeo del Libro Mayor...");
        this.libroMayor.forEach((filaLibro, indexLibro) => {
            Logger.log(`Procesando fila ${indexLibro + 1} del Libro Mayor`);

            // Clave para coincidencia exacta: N° Documento, cargo, abono
            const claveExacta = `${filaLibro[2]}|${filaLibro[3]}|${filaLibro[4]}`;
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
            const siguienteDiaLibroMayor = this.formatearFecha(filaLibro[7]); // Siguiente Día Hábil

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
        Logger.log("Mapeo del Libro Mayor completado.");
    }

    conciliarBCI() {
        Logger.log("Iniciando el proceso de conciliación para el Banco BCI...");
        this.prepararMapeoLibroMayor(); // Preparar los mapeos antes de conciliar

        this.cartola.forEach((filaCartola, indexCartola) => {
            Logger.log(`Procesando fila ${indexCartola + 1} de la Cartola`);

            if (this.conciliadosCartolaIdx.has(indexCartola)) {
                Logger.log(`La fila ${indexCartola + 1} de la Cartola ya ha sido conciliada. Saltando...`);
                return;
            }

            const documentoCartola = filaCartola[2]; // N° Documento
            const cargoCartola = filaCartola[3];
            const abonoCartola = filaCartola[4];
            const descripcionCartola = filaCartola[1];
            const fechaCartola = this.formatearFecha(filaCartola[0]);
            let conciliado = false;

            // Condición 1: Coincidencias exactas (Documento, cargo, abono)
            const claveExacta = `${documentoCartola}|${cargoCartola}|${abonoCartola}`;
            if (this.mapeoExacto.has(claveExacta)) {
                let coincidenciaExacta = this.mapeoExacto.get(claveExacta).find(item => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
                if (coincidenciaExacta) {
                    Logger.log(`Conciliación exacta encontrada en el índice ${indexCartola + 1}`);
                    this.conciliarFilaBCI(filaCartola, coincidenciaExacta.filaLibro, coincidenciaExacta.indexLibro, indexCartola);
                    conciliado = true;
                }
            }

            // Condición 2: Coincidencias por nombre
            const claveNombre = `${descripcionCartola}|${cargoCartola}|${abonoCartola}`;
            if (!conciliado && this.mapeoNombre.has(claveNombre)) {
                let coincidenciaPorNombre = this.mapeoNombre.get(claveNombre).find(item => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
                if (coincidenciaPorNombre) {
                    Logger.log(`Conciliación por nombre encontrada en el índice ${indexCartola + 1}`);
                    this.conciliarFilaBCI(filaCartola, coincidenciaPorNombre.filaLibro, coincidenciaPorNombre.indexLibro, indexCartola);
                    conciliado = true;
                }
            }

            // Condición 3: Coincidencias por fecha y montos
            const claveFechaCargoAbono = `${fechaCartola}|${cargoCartola}|${abonoCartola}`;
            if (!conciliado && this.mapeoFechaCargoAbono.has(claveFechaCargoAbono)) {
                let coincidenciaPorFecha = this.mapeoFechaCargoAbono.get(claveFechaCargoAbono).find(item => !this.conciliadosLibroMayorIdx.has(item.indexLibro));
                if (coincidenciaPorFecha) {
                    Logger.log(`Conciliación por fecha encontrada en el índice ${indexCartola + 1}`);
                    this.conciliarFilaBCI(filaCartola, coincidenciaPorFecha.filaLibro, coincidenciaPorFecha.indexLibro, indexCartola);
                    conciliado = true;
                }
            }

            // Si no se ha conciliado, agregar a pendientes
            if (!conciliado) {
                Logger.log(`La fila ${indexCartola + 1} de la Cartola no pudo ser conciliada. Agregando a pendientes.`);
                this.pendientesCartola.push([
                    this.formatearFecha(filaCartola[0]),
                    filaCartola[1].toUpperCase(),
                    filaCartola[2],
                    filaCartola[3],
                    filaCartola[4]
                ]);
            }
        });

        // Llenar la lista de pendientes del libro mayor
        this.libroMayor.forEach((filaLibro, indexLibro) => {
            if (!this.conciliadosLibroMayorIdx.has(indexLibro)) {
                Logger.log(`La fila ${indexLibro + 1} del Libro Mayor no pudo ser conciliada. Agregando a pendientes.`);
                this.pendientesLibroMayor.push([
                    this.formatearFecha(filaLibro[0]),
                    filaLibro[1].toUpperCase(),
                    filaLibro[2],
                    filaLibro[3],
                    filaLibro[4]
                ]);
            }
        });

        Logger.log("Conciliación completada para todos los registros de la Cartola.");
    }

    conciliarFilaBCI(filaCartola, filaLibroMayor, idxLibroMayor, idxCartola) {
        Logger.log(`Conciliando fila de Cartola en el índice ${idxCartola + 1} con el Libro Mayor en el índice ${idxLibroMayor + 1}`);
        let registroConciliado = [
            this.formatearFecha(filaCartola[0]),  // Fecha
            filaCartola[1].toUpperCase(),        // Descripción
            filaLibroMayor[2],                   // N° Documento
            filaLibroMayor[5],                   // RUT (Libro Mayor)
            filaCartola[3],                      // Cargo (Cartola)
            filaCartola[4]                       // Abono (Cartola)
        ];

        // Agregar a conciliados
        this.conciliados.push(registroConciliado);
        this.conciliadosCartolaIdx.add(idxCartola);
        this.conciliadosLibroMayorIdx.add(idxLibroMayor);
    }

    generarArchivoSalidaBCI() {
        try {
            Logger.log('Iniciando generación del archivo de salida...');
            const fechaActual = new Date();
            const sufijoFechaHora = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
            const nombreArchivo = `Conciliación_Bancaria_BCI_${sufijoFechaHora}`;
            const hojaSalida = SpreadsheetApp.create(nombreArchivo);

            // Mover archivo a la carpeta
            const file = DriveApp.getFileById(hojaSalida.getId());
            const folder = DriveApp.getFolderById('1YZeYpz9c79GFp8lOfrWgEUHIj0J2LTVt'); // Carpeta de Salida Conciliados BCI
            folder.addFile(file);
            DriveApp.getRootFolder().removeFile(file);

            const hojaConciliados = hojaSalida.getActiveSheet();
            hojaConciliados.setName('Conciliados');

            let datosConciliados = [['FECHA', 'DESCRIPCIÓN', 'N° DOCUMENTO', 'RUT', 'CARGO', 'ABONO']];
            this.conciliados.forEach(registro => datosConciliados.push(registro));
            hojaConciliados.getRange(1, 1, datosConciliados.length, datosConciliados[0].length).setValues(datosConciliados);

            const hojaPendientesCartola = hojaSalida.insertSheet('Pendientes Cartola');
            let datosPendientesCartola = [['FECHA', 'DESCRIPCIÓN', 'N° DOCUMENTO', 'CARGO', 'ABONO']];
            this.pendientesCartola.forEach(registro => datosPendientesCartola.push(registro));
            hojaPendientesCartola.getRange(1, 1, datosPendientesCartola.length, datosPendientesCartola[0].length).setValues(datosPendientesCartola);

            const hojaPendientesLibroMayor = hojaSalida.insertSheet('Pendientes Libro Mayor');
            let datosPendientesLibroMayor = [['FECHA', 'DESCRIPCIÓN', 'N° DOCUMENTO', 'CARGO', 'ABONO']];
            this.pendientesLibroMayor.forEach(registro => datosPendientesLibroMayor.push(registro));
            hojaPendientesLibroMayor.getRange(1, 1, datosPendientesLibroMayor.length, datosPendientesLibroMayor[0].length).setValues(datosPendientesLibroMayor);

            Logger.log("Archivo de salida generado exitosamente.");
            return hojaSalida.getUrl();
        } catch (error) {
            Logger.log('Error al generar el archivo de salida: ' + error.message);
            throw new Error('Error al generar el archivo de salida');
        }
    }
}

// Función para iniciar la conciliación para el banco BCI
function iniciarConciliacionDesdeBotonBCI(sheetId) {
    try {
        Logger.log(`Iniciando conciliación para el Banco BCI con el archivo ID: ${sheetId}`);
        let sheetProcesado = SpreadsheetApp.openById(sheetId);
        let cartolaSheet = sheetProcesado.getSheetByName('Cartola');
        let libroMayorSheet = sheetProcesado.getSheetByName('Libro Mayor');

        if (!cartolaSheet || !libroMayorSheet) {
            throw new Error("No se encontraron las hojas 'Cartola' o 'Libro Mayor'.");
        }

        let cartolaData = cartolaSheet.getRange(2, 1, cartolaSheet.getLastRow() - 1, 6).getValues();
        let libroMayorData = libroMayorSheet.getRange(2, 1, libroMayorSheet.getLastRow() - 1, 8).getValues();

        let conciliador = new ConciliadorBCI(cartolaData, libroMayorData);
        conciliador.conciliarBCI();

        let urlArchivoSalida = conciliador.generarArchivoSalidaBCI();

        let emailUsuario = Session.getActiveUser().getEmail();
        let totalConciliados = conciliador.conciliados.length;
        let totalPendientesCartola = conciliador.pendientesCartola.length;
        let totalPendientesLibroMayor = conciliador.pendientesLibroMayor.length;
        let totalRegistrosCartola = cartolaData.length;

        let porcentajeConciliados = ((totalConciliados / totalRegistrosCartola) * 100).toFixed(2);

        MailApp.sendEmail({
            to: emailUsuario,
            subject: 'Conciliación Bancaria BCI Completada',
            htmlBody: `
                La conciliación bancaria ha sido completada exitosamente. <br><br>
                Total de registros conciliados: ${totalConciliados} <br>
                Pendientes Cartola: ${totalPendientesCartola} <br>
                Pendientes Libro Mayor: ${totalPendientesLibroMayor} <br>
                Porcentaje Conciliados: ${porcentajeConciliados}% <br><br>
                Puede descargar el archivo desde el siguiente enlace: <a href="${urlArchivoSalida}" target="_blank">Ver archivo conciliado</a>
            `
        });

        return {
            success: true,
            message: `Conciliación completada. Correo enviado a ${emailUsuario}`,
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
