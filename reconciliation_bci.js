/**
 * Clase para realizar la conciliación bancaria del Banco BCI.
 * @class
 */
class ConciliadorBCI {
    /**
     * Crea una instancia de ConciliadorBCI.
     * @param {Array[]} cartola - Datos de la cartola bancaria procesada.
     * @param {Array[]} libroMayor - Datos del libro mayor contable procesado.
     */
    constructor(cartola, libroMayor) {
        this.cartola = cartola;
        this.libroMayor = libroMayor;
        this.conciliados = [];
        this.pendientesCartola = [];
        this.pendientesLibroMayor = [];
        this.conciliadosLibroMayorIdx = new Set();
        this.conciliadosCartolaIdx = new Set();
        this.mapeoExacto = new Map();  // Mapeo para coincidencias exactas (Documento, debe, haber)
        this.mapeoNDocumentoFechaCargoAbono = new Map(); // Mapeo para coincidencias por N° Documento, Fecha y Montos
    }

    /**
     * Formatea una fecha al formato 'dd/MM/yyyy'.
     * Si la fecha no es válida, retorna una cadena vacía.
     * @param {string|Date} fecha - La fecha a formatear.
     * @returns {string} La fecha formateada o una cadena vacía si es inválida.
     */
    formatearFecha(fecha) {
        try {
            return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        } catch (error) {
            Logger.log(`Error al formatear la fecha: ${fecha}. ${error.message}`);
            return '';
        }
    }

    /**
     * Prepara los mapeos del libro mayor para poder realizar las coincidencias durante la conciliación.
     */
    prepararMapeoLibroMayor() {
        this.libroMayor.forEach((filaLibro, indexLibro) => {
            const claveExacta = `${filaLibro[2]}|${filaLibro[3]}|${filaLibro[4]}`;
            if (!this.mapeoExacto.has(claveExacta)) {
                this.mapeoExacto.set(claveExacta, []);
            }
            this.mapeoExacto.get(claveExacta).push({ indexLibro, filaLibro });

            // Claves para coincidencias de N° Documento, Fecha y Montos, considerando Fecha o Siguiente Día Hábil
            const fechaLibroMayor = this.formatearFecha(filaLibro[0]);
            const siguienteDiaLibroMayor = this.formatearFecha(filaLibro[7]);

            const claveFecha = `${filaLibro[2]}|${fechaLibroMayor}|${filaLibro[3]}|${filaLibro[4]}`;
            const claveSiguienteDia = `${filaLibro[2]}|${siguienteDiaLibroMayor}|${filaLibro[3]}|${filaLibro[4]}`;
            
            if (!this.mapeoNDocumentoFechaCargoAbono.has(claveFecha)) {
                this.mapeoNDocumentoFechaCargoAbono.set(claveFecha, []);
            }
            this.mapeoNDocumentoFechaCargoAbono.get(claveFecha).push({ indexLibro, filaLibro });

            if (!this.mapeoNDocumentoFechaCargoAbono.has(claveSiguienteDia)) {
                this.mapeoNDocumentoFechaCargoAbono.set(claveSiguienteDia, []);
            }
            this.mapeoNDocumentoFechaCargoAbono.get(claveSiguienteDia).push({ indexLibro, filaLibro });
        });
    }

    /**
     * Realiza el proceso de conciliación entre la cartola y el libro mayor para el Banco BCI.
     */
    conciliarBCI() {
        this.prepararMapeoLibroMayor();
        this.cartola.forEach((filaCartola, indexCartola) => {
            if (this.conciliadosCartolaIdx.has(indexCartola)) return;

            const documentoCartola = filaCartola[2];
            const debeCartola = filaCartola[4];
            const haberCartola = filaCartola[3];
            const fechaCartola = this.formatearFecha(filaCartola[0]);
            let conciliado = false;

            // Condición 1: Coincidencia Exacta en Documento, Debe, Haber sin importar fecha
            const claveExacta = `${documentoCartola}|${debeCartola}|${haberCartola}`;
            const coincidenciaExacta = this.mapeoExacto.get(claveExacta)?.find(item => !this.conciliadosLibroMayorIdx.has(item.indexLibro));

            if (coincidenciaExacta) {
                this.conciliarFilaBCI(filaCartola, coincidenciaExacta.filaLibro, coincidenciaExacta.indexLibro, indexCartola);
                Logger.log(`Conciliación exacta encontrada: Cartola ${indexCartola + 1} con Libro Mayor ${coincidenciaExacta.indexLibro + 1}`);
                conciliado = true;
            }

            // Condición 2: Coincidencias en Documento, Fecha o Siguiente Día Hábil, Cargo y Abono
            const claveFecha = `${documentoCartola}|${fechaCartola}|${debeCartola}|${haberCartola}`;
            const claveSiguienteDia = `${documentoCartola}|${this.formatearFecha(filaCartola[7] || '')}|${debeCartola}|${haberCartola}`;
            const coincidenciaPorFecha = this.mapeoNDocumentoFechaCargoAbono.get(claveFecha)?.find(item => !this.conciliadosLibroMayorIdx.has(item.indexLibro)) ||
                                         this.mapeoNDocumentoFechaCargoAbono.get(claveSiguienteDia)?.find(item => !this.conciliadosLibroMayorIdx.has(item.indexLibro));

            if (!conciliado && coincidenciaPorFecha) {
                this.conciliarFilaBCI(filaCartola, coincidenciaPorFecha.filaLibro, coincidenciaPorFecha.indexLibro, indexCartola);
                Logger.log(`Conciliación por Documento y Fecha encontrada: Cartola ${indexCartola + 1} con Libro Mayor ${coincidenciaPorFecha.indexLibro + 1}`);
                conciliado = true;
            }

            if (!conciliado) {
                Logger.log(`Registro pendiente en Cartola índice ${indexCartola + 1}`);
                this.pendientesCartola.push([
                    this.formatearFecha(filaCartola[0]),  
                    filaCartola[1].toUpperCase(),         
                    documentoCartola,                     
                    debeCartola,                          
                    haberCartola                          
                ]);
            }
        });

        this.libroMayor.forEach((filaLibro, indexLibro) => {
            if (!this.conciliadosLibroMayorIdx.has(indexLibro)) {
                Logger.log(`Registro pendiente en Libro Mayor índice ${indexLibro + 1}`);
                this.pendientesLibroMayor.push([
                    this.formatearFecha(filaLibro[0]),  
                    filaLibro[1].toUpperCase(),         
                    filaLibro[2],                       
                    filaLibro[3],                       
                    filaLibro[4]                        
                ]);
            }
        });
    }

    /**
     * Conciliar una fila de la cartola con una fila del libro mayor.
     * @param {Array} filaCartola - Fila de la cartola.
     * @param {Array} filaLibroMayor - Fila del libro mayor.
     * @param {number} idxLibroMayor - Índice de la fila en el libro mayor.
     * @param {number} idxCartola - Índice de la fila en la cartola.
     */
    conciliarFilaBCI(filaCartola, filaLibroMayor, idxLibroMayor, idxCartola) {
        const registroConciliado = [
            this.formatearFecha(filaCartola[0]),  
            filaCartola[1].toUpperCase(),        
            filaLibroMayor[2],                   
            filaLibroMayor[5],                   
            filaCartola[4],                      
            filaCartola[3]                       
        ];

        this.conciliados.push(registroConciliado);
        this.conciliadosCartolaIdx.add(idxCartola);
        this.conciliadosLibroMayorIdx.add(idxLibroMayor);
    }

    /**
     * Genera el archivo de salida con los datos conciliados y pendientes.
     * @returns {string} URL del archivo generado.
     */
    generarArchivoSalidaBCI() {
        try {
            const fechaActual = new Date();
            const sufijoFechaHora = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
            const nombreArchivo = `Conciliación_Bancaria_BCI_${sufijoFechaHora}`;
            const hojaSalida = SpreadsheetApp.create(nombreArchivo);
            const file = DriveApp.getFileById(hojaSalida.getId());
            const folder = DriveApp.getFolderById('1YZeYpz9c79GFp8lOfrWgEUHIj0J2LTVt');
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

            return hojaSalida.getUrl();
        } catch (error) {
            Logger.log('Error al generar el archivo de salida: ' + error.message);
            throw new Error('Error al generar el archivo de salida');
        }
    }
}

/**
 * Función para iniciar la conciliación para el banco BCI desde un botón.
 * @param {string} sheetId - ID de la hoja de cálculo de entrada.
 * @returns {Object} Resultado de la conciliación.
 */
function iniciarConciliacionDesdeBotonBCI(sheetId) {
    try {
        Logger.log(`Iniciando conciliación para el Banco BCI con el archivo ID: ${sheetId}`);
        const sheetProcesado = SpreadsheetApp.openById(sheetId);
        const cartolaSheet = sheetProcesado.getSheetByName('Cartola');
        const libroMayorSheet = sheetProcesado.getSheetByName('Libro Mayor');

        if (!cartolaSheet || !libroMayorSheet) {
            throw new Error("No se encontraron las hojas 'Cartola' o 'Libro Mayor'.");
        }

        const cartolaData = cartolaSheet.getRange(2, 1, cartolaSheet.getLastRow() - 1, 5).getValues();
        const libroMayorData = libroMayorSheet.getRange(2, 1, libroMayorSheet.getLastRow() - 1, 8).getValues();

        const conciliador = new ConciliadorBCI(cartolaData, libroMayorData);
        conciliador.conciliarBCI();

        const urlArchivoSalida = conciliador.generarArchivoSalidaBCI();
        const emailUsuario = Session.getActiveUser().getEmail();
        const totalConciliados = conciliador.conciliados.length;
        const totalPendientesCartola = conciliador.pendientesCartola.length;
        const totalPendientesLibroMayor = conciliador.pendientesLibroMayor.length;
        const totalRegistrosCartola = cartolaData.length;
        const porcentajeConciliados = ((totalConciliados / totalRegistrosCartola) * 100).toFixed(2);

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
            totalConciliados,
            totalPendientesCartola,
            totalPendientesLibroMayor,
            porcentajeConciliados
        };
    } catch (error) {
        Logger.log('Error al iniciar la conciliación: ' + error.message);
        return {
            success: false,
            message: 'Error al iniciar la conciliación: ' + error.message
        };
    }
}
