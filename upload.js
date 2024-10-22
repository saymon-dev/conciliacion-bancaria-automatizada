// upload.gs

// ID de la carpeta de entrada de Google Drive
var folderIdInputEstado = '1hvJkvXPHBORiNAtW0pWIk5n_hQmxdlE_';  // Reemplaza con el ID de la carpeta de Banco Estado
var folderIdInputBCI = '1OxwizpiIep2FAiv35SKN8J54N4MfhJst';       // Reemplaza con el ID de la carpeta de Banco BCI


/**
 * Función para subir y convertir el archivo a Google Sheets
 * @param {string} fileName El nombre del archivo
 * @param {string} fileContent El contenido en base64 del archivo
 * @return {object} Resultado de la subida y conversión
 */
function uploadToDrive(fileName, base64Content, banco) {
  try {
    Logger.log('Iniciando el proceso de subida y conversión del archivo.');

    // Decodificar el contenido base64
    var decodedFile = Utilities.base64Decode(base64Content);
    var blob = Utilities.newBlob(decodedFile, MimeType.MICROSOFT_EXCEL, fileName);

    // Subir el archivo a Google Drive
    var folderIdInput = (banco === 'Estado') ? folderIdInputEstado : folderIdInputBCI;
    var folderInput = DriveApp.getFolderById(folderIdInput);
    var uploadedFile = folderInput.createFile(blob);

    Logger.log('Archivo subido a la carpeta de entrada con éxito: ' + uploadedFile.getId());

    // Convertir el archivo a Google Sheets
    var newFile = Drive.Files.copy({ title: fileName.replace(".xlsx", ""), mimeType: MimeType.GOOGLE_SHEETS }, uploadedFile.getId());
    
    // Eliminar el archivo XLSX original
    uploadedFile.setTrashed(true);

    return {
      success: true,
      message: 'Archivo subido correctamente.',
      sheetId: newFile.id,
      sheetUrl: 'https://docs.google.com/spreadsheets/d/' + newFile.id
    };
  } catch (error) {
    Logger.log('Error al subir el archivo: ' + error.message);
    return {
      success: false,
      message: 'Error al subir el archivo: ' + error.message
    };
  }
}


function obtenerFotoPerfilDesdeUsuarioActivo() {
  try {
    // Usamos la API de Google People para obtener la información del usuario activo ('people/me')
    const people = People.People.get('people/me', { personFields: 'photos' });

    if (people.photos && people.photos.length > 0) {
      // Devolver la URL de la primera foto de perfil del usuario
      return people.photos[0].url;
    }
  } catch (error) {
    Logger.log('Error al obtener la foto de perfil del usuario activo: ' + error.message);
  }

  // Si no se encuentra foto o ocurre un error, devolver una URL de imagen por defecto
  return 'https://www.researchgate.net/publication/315108532/figure/fig1/AS:472492935520261@1489662502634/Figura-2-Avatar-que-aparece-por-defecto-en-Facebook.png';  // Cambiar por una URL de imagen predeterminada
}


function obtenerFotoPerfilDesdeCorreo(email) {
  try {
    // Usamos la API de Google People para obtener la información del usuario a partir del correo
    const person = People.People.get('people/' + email, { personFields: 'photos' });

    if (person.photos && person.photos.length > 0) {
      // Asegurarse de que la primera foto tenga una URL válida
      return person.photos[0].url;
    }
  } catch (error) {
    Logger.log('Error al obtener la foto de perfil del usuario con correo: ' + email + '. ' + error.message);
  }

  // Si no se encuentra foto o ocurre un error, devolver una URL de imagen por defecto
  return 'https://www.researchgate.net/publication/315108532/figure/fig1/AS:472492935520261@1489662502634/Figura-2-Avatar-que-aparece-por-defecto-en-Facebook.png';  // Cambiar por una URL de imagen predeterminada
}



/**
 * Obtener el correo del usuario activo.
 * Esta función será llamada desde el HTML para mostrar el correo del usuario.
 * @return {string} Correo del usuario activo.
 */
function getActiveUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getUserInfo() {
  const email = getActiveUserEmail();

  // Primer enfoque: Obtener la foto desde el usuario activo
  let photoUrl = obtenerFotoPerfilDesdeUsuarioActivo();

  // Alternativamente, puedes probar el segundo enfoque usando el correo del usuario
  // let photoUrl = obtenerFotoPerfilDesdeCorreo(email);

  return {
    email: email,
    photoUrl: photoUrl
  };
}

/**
 *
 * Esta funcion decide que metodo ejecutar segun el banco selecionado 
 * 
 * **/
function ejecutarProcesamiento(sheetId, banco){
  try {
    if (banco === 'Estado') {
      return ejecutarProcesamientoEstado(sheetId);
    } else if (banco === 'BCI') {
      return ejecutarProcesamientoBCI(sheetId);
    } else {
      throw new Error('Banco no reconocido.');
    }
  } catch (error) {
    Logger.log('Error al procesar el archivo: ' + error.message);
    return {
      success: false,
      message: 'Error al procesar el archivo: ' + error.message
    };
  }
}

/**
 *
 * Esta función decide qué método de conciliación ejecutar según el banco seleccionado
 * 
 * @param {string} sheetId - El ID de la hoja de Google Sheets
 * @param {string} banco - El banco seleccionado ("Estado" o "BCI")
 * @return {object} - El resultado del proceso de conciliación
 * 
 **/
function iniciarConciliacion(sheetId, banco) {
  try {
    if (!sheetId) {
      throw new Error('El ID de la hoja de cálculo proporcionado no es válido.');
    }

    if (banco === 'Estado') {
      // Ejecuta el método de conciliación para Banco Estado
      return iniciarConciliacionDesdeBotonEstado(sheetId);
    } else if (banco === 'BCI') {
      // Ejecuta el método de conciliación para Banco BCI
      return iniciarConciliacionDesdeBotonBCI(sheetId);
    } else {
      // Si no se reconoce el banco, se lanza un error
      throw new Error('Banco no reconocido. Se proporcionó: ' + banco);
    }
  } catch (error) {
    // Captura y registra cualquier error que ocurra durante el proceso
    Logger.log('Error al conciliar los registros: ' + error.message);
    return {
      success: false,
      message: 'Error al conciliar los registros: ' + error.message
    };
  }
}



/**
 * Procesa un archivo y luego realiza la conciliación para el banco especificado.
 * 
 * @param {string} sheetId - ID de la hoja de Google Sheets del archivo subido.
 * @param {string} banco - Banco para el que se realizará el procesamiento y la conciliación ("Estado" o "BCI").
 * @returns {Object} - Resultado del proceso, incluyendo el éxito, mensaje descriptivo y URL del archivo generado.
 */
function procesarYConciliarGestionado(sheetId, banco) {
  try {
    // Verificar si el sheetId es válido
    if (!sheetId) {
      throw new Error('El ID de la hoja de cálculo es inválido.');
    }

    // Intentar abrir el archivo con el sheetId proporcionado
    let archivoProcesado;
    try {
      archivoProcesado = SpreadsheetApp.openById(sheetId);
      Logger.log('Archivo procesado abierto correctamente: ' + archivoProcesado.getId());
    } catch (error) {
      throw new Error('No se pudo abrir el archivo con el ID proporcionado.');
    }

    // Verificar si las hojas "Cartola" y "Libro Mayor" existen
    const hojaCartola = archivoProcesado.getSheetByName('Cartola');
    const hojaLibroMayor = archivoProcesado.getSheetByName('Libro Mayor');
    
    if (!hojaCartola || !hojaLibroMayor) {
      throw new Error('No se encontraron las hojas "Cartola" o "Libro Mayor" en el archivo.');
    }

    // 1. Procesar el archivo si las hojas existen
    Logger.log('Iniciando procesamiento del archivo para el banco: ' + banco);
    let resultadoProcesamiento = ejecutarProcesamiento(sheetId, banco);

    // Verificamos si hubo errores en el procesamiento
    if (!resultadoProcesamiento.success) {
      throw new Error('Error durante el procesamiento: ' + resultadoProcesamiento.message);
    }

    // 2. Iniciar la conciliación con el ID del archivo procesado
    Logger.log('Iniciando conciliación para el archivo procesado. Banco: ' + banco);
    let resultadoConciliacion = iniciarConciliacion(resultadoProcesamiento.processedSheetId, banco);

    // Verificamos si hubo errores en la conciliación
    if (!resultadoConciliacion.success) {
      throw new Error('Error durante la conciliación: ' + resultadoConciliacion.message);
    }

    // 3. Obtener el correo del usuario activo
    const emailUsuario = Session.getActiveUser().getEmail();
    
    // Retornar el resultado final al frontend con el URL del archivo conciliado
    return {
      success: true,
      message: '<strong>Procesamiento y conciliación completados exitosamente.</strong>',
      porcentajeConciliados: resultadoConciliacion.porcentajeConciliados,
      conciledFileUrl: resultadoConciliacion.conciledFileUrl,
      emailUsuario: emailUsuario
    };

  } catch (error) {
    Logger.log('Error en procesarYConciliarGestionado: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}


/** 
// Definir el ID de la hoja de cálculo para guardar conciliaciones
const hojaConciliaciones = '1vyVxU0Us-ETFewAH7Gu1YS2OIaB1lY3NU1XOXHfLCo4';  // Reemplaza con tu ID de la hoja de cálculo

// Función para abrir la hoja de cálculo
function abrirHojaDeCalculo() {
  try {
    Logger.log('Intentando abrir la hoja de cálculo con ID: ' + hojaConciliaciones);
    const ss = SpreadsheetApp.openById(hojaConciliaciones);
    return ss;
  } catch (error) {
    Logger.log('Error al abrir la hoja de cálculo: ' + error.message);
    throw new Error('No se pudo abrir la hoja de cálculo.');
  }
}

// Función para guardar conciliaciones en la hoja de cálculo
function guardarConciliacionEnSheet(datosConciliacion) {
  try {
    const sheet = abrirHojaDeCalculo().getSheetByName('Conciliaciones');
    
    if (!sheet) {
      throw new Error('No se encontró la hoja "Conciliaciones".');
    }

    // Añadir una nueva fila con los detalles
    sheet.appendRow([
      datosConciliacion.nombre,
      datosConciliacion.fechaHora,
      datosConciliacion.correo,
      datosConciliacion.urlArchivo,
      datosConciliacion.banco
    ]);

    Logger.log('Conciliación guardada exitosamente.');
  } catch (error) {
    Logger.log('Error al guardar la conciliación: ' + error.message);
    throw new Error('No se pudo guardar la conciliación.');
  }
}

// Función para cargar las conciliaciones guardadas del usuario
function cargarConciliacionesUsuario(email) {
  try {
    Logger.log('Conciliaciones cargadas exitosamente: ' + JSON.stringify(conciliaciones));
    const sheet = abrirHojaDeCalculo().getSheetByName('Conciliaciones');
    
    if (!sheet) {
      throw new Error('No se encontró la hoja "Conciliaciones".');
    }

    const data = sheet.getDataRange().getValues();
    const conciliaciones = [];

    // Iterar por las filas para buscar las conciliaciones que coincidan con el correo del usuario
    for (let i = 1; i < data.length; i++) {  // Empezar en 1 para saltar la fila de encabezado
      const row = data[i];
      if (row[2] === email) {  // Filtrar por correo del usuario
        conciliaciones.push({
          row: i + 1,  // Índice de la fila
          nombre: row[0],
          fechaHora: row[1],
          correo: row[2],
          urlArchivo: row[3],
          banco: row[4]  // Incluir el banco
        });
      }
    }

    return conciliaciones;  // Devolver las conciliaciones filtradas
  } catch (error) {
    Logger.log('Error al cargar las conciliaciones del usuario: ' + error.message);
    throw new Error('No se pudieron cargar las conciliaciones.');
  }
}

// Actualizar nombre de conciliación guardada desde frontend
function actualizarNombreConciliacion(row, nuevoNombre) {
  try {
    const sheet = abrirHojaDeCalculo().getSheetByName('Conciliaciones');
    
    if (!sheet) {
      throw new Error('No se encontró la hoja "Conciliaciones".');
    }

    sheet.getRange(row, 1).setValue(nuevoNombre);  // Actualizar nombre en la columna 1
    Logger.log('Nombre de la conciliación actualizado.');
  } catch (error) {
    Logger.log('Error al actualizar el nombre de la conciliación: ' + error.message);
    throw new Error('No se pudo actualizar el nombre de la conciliación.');
  }
}

// Eliminar conciliación guardada
function eliminarConciliacion(row) {
  try {
    const sheet = abrirHojaDeCalculo().getSheetByName('Conciliaciones');
    
    if (!sheet) {
      throw new Error('No se encontró la hoja "Conciliaciones".');
    }

    sheet.deleteRow(row);  // Eliminar la fila completa
    Logger.log('Conciliación eliminada.');
  } catch (error) {
    Logger.log('Error al eliminar la conciliación: ' + error.message);
    throw new Error('No se pudo eliminar la conciliación.');
  }
}
*/

/**
 * Función para cargar la página appPage.html
 * Esta función se llama cuando el usuario visita la URL de la aplicación web.
 * @return {HtmlOutput} La página HTML servida
 */

function doGet() {
  // Obtener información del usuario
  const userInfo = getUserInfo();

  // Crear la plantilla HTML
  const template = HtmlService.createTemplateFromFile('appPage');

  // Pasar la información del usuario a la plantilla
  template.email = userInfo.email;
  template.photoUrl = userInfo.photoUrl;

  // Devolver la página HTML evaluada
  return template.evaluate()
    .setTitle('Conciliación Bancaria')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
