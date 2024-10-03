// upload.gs

// ID de la carpeta de entrada de Google Drive
var folderIdInputEstado = '1I75CHO31YTgBXycx5eoBQIJeemtfXeEm';  // Reemplaza con el ID de la carpeta de Banco Estado
var folderIdInputBCI = '1TInd0WIU8I0LP5HT10gQB1MeDO8NZCNo';       // Reemplaza con el ID de la carpeta de Banco BCI


/**
 * Función para subir y convertir el archivo a Google Sheets
 * @param {string} fileName El nombre del archivo
 * @param {string} fileContent El contenido en base64 del archivo
 * @return {object} Resultado de la subida y conversión
 */
function uploadToDrive(fileName, fileContent, banco) {
  try {
    Logger.log('Iniciando el proceso de subida y conversión del archivo.');

    // Obtener el correo del usuario activo
    var emailUsuario = Session.getActiveUser().getEmail();
    Logger.log('Correo del usuario: ' + emailUsuario);

    // Restringir el acceso a usuarios de la organización 'hyl.cl'
    if (!emailUsuario.trim().toLowerCase().endsWith('@hyl.cl')) {
      throw new Error('Acceso denegado: Solo los usuarios de la organización "hyl.cl" pueden usar esta aplicación.');
    }

    // Decodificar el contenido base64 del archivo
    var decodedFile = Utilities.base64Decode(fileContent.split(',')[1]);
    var blob = Utilities.newBlob(decodedFile, MimeType.MICROSOFT_EXCEL, fileName);

    // Obtener la fecha y hora actual en el formato ddMMyyyy_HHmmss
    var fechaActual = new Date();
    var sufijo = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'ddMMyyyy_HHmmss');

    // Agregar el sufijo al nombre del archivo
    var newFileName = fileName.replace(/\.[^/.]+$/, "") + "_" + sufijo + ".xlsx";  // Suponiendo que el archivo es .xlsx

    // Determinar la carpeta de destino según el banco
    var folderIdInput;
    if (banco === 'Estado') {
      folderIdInput = folderIdInputEstado;
    } else if (banco === 'BCI') {
      folderIdInput = folderIdInputBCI;
    } else {
      throw new Error('Banco no válido: Debe ser "Estado" o "BCI".');
    }

    // Subir el archivo a la carpeta de entrada con el nombre actualizado
    var folderInput = DriveApp.getFolderById(folderIdInput);
    var uploadedFile = folderInput.createFile(blob.setName(newFileName));
    Logger.log('Archivo subido a la carpeta de entrada con éxito: ' + uploadedFile.getId());

    // Convertir el archivo subido a Google Sheets utilizando la API avanzada de Google Drive
    var newFile = Drive.Files.copy({ title: newFileName.replace(".xlsx", ""), mimeType: MimeType.GOOGLE_SHEETS }, uploadedFile.getId());
    Logger.log('Archivo convertido a Google Sheets con éxito. ID: ' + newFile.id);

    // Eliminar el archivo XLSX original
    uploadedFile.setTrashed(true);
    /*
    // Enviar correo al usuario
    MailApp.sendEmail({
      to: emailUsuario,
      subject: 'Archivo subido correctamente',
      htmlBody: 'El archivo ha sido subido correctamente. Puede acceder a él en el siguiente enlace: <a href="https://docs.google.com/spreadsheets/d/' + newFile.id + '">Ver archivo subido</a>'
    });
    Logger.log('Correo enviado a: ' + emailUsuario);
    */
    // Retornar el ID y URL del archivo convertido para el procesamiento
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
    if (banco === 'Estado') {
      // Ejecuta el método de conciliación para Banco Estado
      return iniciarConciliacionDesdeBotonEstado(sheetId);
    } else if (banco === 'BCI') {
      // Ejecuta el método de conciliación para Banco BCI
      return iniciarConciliacionDesdeBotonBCI(sheetId);
    } else {
      // Si no se reconoce el banco, se lanza un error
      throw new Error('Banco no reconocido.');
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
