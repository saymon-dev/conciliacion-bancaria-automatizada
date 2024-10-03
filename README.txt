# Conciliación Bancaria - Sistema de Automatización

Este proyecto proporciona un sistema automatizado para la conciliación bancaria para los bancos "Banco Estado" y "Banco BCI". El sistema permite a los usuarios subir archivos de conciliación, procesarlos, y realizar la conciliación entre la cartola bancaria y el libro mayor, generando reportes y archivos de salida.

## Estructura del Proyecto

### Archivos y Descripción:

1. **appPage.html**:
   - Proporciona la interfaz de usuario para la aplicación de conciliación bancaria. Permite la selección de bancos, la subida de archivos, el procesamiento, y la conciliación.
   - Incluye pestañas para cada banco con botones para subir archivos, procesarlos, y ejecutar la conciliación.
   - [Ver detalles](27).

2. **appsscript.json**:
   - Archivo de configuración para el proyecto de Google Apps Script. Define las dependencias necesarias (Google Sheets, Drive, Gmail, y People API) y los permisos OAuth requeridos.
   - Incluye la zona horaria para América/Santiago y permisos avanzados.
   - [Ver detalles](28).

3. **process_bci.js**:
   - Gestiona el procesamiento de archivos relacionados con el Banco BCI. Procesa las hojas de cartola y libro mayor, obteniendo los datos necesarios para la conciliación.
   - Mueve los archivos procesados a una carpeta de Google Drive.
   - [Ver detalles](29).

4. **process_estado.js**:
   - Similar a `process_bci.js`, pero para el Banco Estado. Procesa las hojas de cartola y libro mayor del Banco Estado y prepara los datos para la conciliación.
   - Incluye lógica para normalizar los datos y extraer campos clave.
   - [Ver detalles](30).

5. **reconciliation_bci.js**:
   - Ejecuta la conciliación automática para el Banco BCI. Compara los registros de la cartola con los del libro mayor utilizando coincidencias exactas y aproximadas (por documento, montos, o fechas).
   - Genera un archivo de conciliación con los resultados.
   - [Ver detalles](31).

6. **reconciliation_estado.js**:
   - Similar a `reconciliation_bci.js`, pero para el Banco Estado. Realiza la conciliación de registros, genera archivos de resultados y reportes sobre registros conciliados y pendientes.
   - [Ver detalles](32).

7. **upload.js**:
   - Gestiona la subida de archivos a Google Drive, su conversión a Google Sheets, y la verificación de acceso de los usuarios.
   - Los archivos subidos son enviados a las carpetas correspondientes de Banco Estado o Banco BCI, y se eliminan los archivos originales después de ser convertidos.
   - [Ver detalles](33).

## Funcionalidades Principales

- **Subida de Archivos**: Los usuarios pueden subir archivos en formato Excel, que se convierten a Google Sheets.
- **Procesamiento de Archivos**: El sistema procesa las hojas de cartola y libro mayor para cada banco.
- **Conciliación Automática**: El sistema compara los registros de la cartola con los del libro mayor, conciliando automáticamente donde sea posible.
- **Generación de Reportes**: Se generan archivos de conciliación con los registros procesados y los resultados de la conciliación.
- **Notificaciones por Correo**: Los usuarios reciben notificaciones por correo electrónico con los enlaces a los archivos procesados o conciliados.

## Instalación y Despliegue

1. **Configuración Inicial**:
   - Clona o descarga este repositorio.
   - Sube los archivos del proyecto a Google Apps Script.

2. **Habilitar APIs y Permisos**:
   - En Google Cloud Platform, habilita las APIs necesarias:
     - Google Drive API.
     - Google Sheets API.
     - Gmail API.
     - People API.
   - Asegúrate de configurar los permisos OAuth en `appsscript.json` para permitir el acceso a los servicios mencionados.

3. **Actualizar IDs de Carpetas**:
   - En `upload.js`, actualiza los IDs de las carpetas de entrada para almacenar los archivos subidos. Reemplaza los siguientes valores con los IDs de las carpetas de Google Drive correspondientes:
     ```javascript
     var folderIdInputEstado = 'ID_DE_LA_CARPETA_DE_ENTRADA_BANCO_ESTADO';
     var folderIdInputBCI = 'ID_DE_LA_CARPETA_DE_ENTRADA_BANCO_BCI';
     ```
   - En `process_bci.js` y `process_estado.js`, actualiza los IDs de las carpetas de salida para almacenar los archivos procesados:
     ```javascript
     var folderIdOutput = 'ID_DE_LA_CARPETA_DE_PROCESADOS';
     ```
   - En `reconciliation_bci.js` y `reconciliation_estado.js`, actualiza los IDs de las carpetas donde se almacenarán los archivos de conciliación:
     ```javascript
     var folderIdOutputConciliados = 'ID_DE_LA_CARPETA_DE_CONCILIADOS';
     ```

4. **Desplegar la Aplicación**:
   - Publica el proyecto en Google Apps Script como una aplicación web. Permite que "Cualquiera, incluso anónimos" pueda acceder a la web, pero asegúrate de restringir la subida de archivos a usuarios con dominio `@hyl.cl`.

5. **Notificaciones**:
   - Los usuarios recibirán correos electrónicos con los resultados de los procesos de subida, procesamiento, y conciliación.

## Uso

1. **Subir un Archivo**:
   - Selecciona el banco correspondiente (Estado o BCI).
   - Sube el archivo de conciliación utilizando el botón "Subir Archivo".
   
2. **Procesar Archivo**:
   - Una vez subido el archivo, presiona "Procesar Archivo" para extraer los datos de la cartola y el libro mayor.
   
3. **Conciliar Registros**:
   - Presiona "Conciliar Registros" para ejecutar la conciliación automática. Recibirás un correo con el archivo de resultados.

## API y Permisos

- **Google Drive API**: Para gestionar la subida, descarga y almacenamiento de archivos.
- **Google Sheets API**: Para la creación y modificación de hojas de cálculo.
- **Gmail API**: Para enviar notificaciones por correo.
- **People API**: Para obtener la información del usuario que sube los archivos.

---

Este archivo README proporciona una visión general del sistema y su funcionamiento. Si necesitas más información o ejemplos de código, consulta los comentarios en los archivos fuente.

--- 
