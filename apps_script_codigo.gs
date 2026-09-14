/**
 * APPS SCRIPT — Monitoreo de planes (Ley 26.331)
 *
 * Qué hace:
 *  - GET  → devuelve todos los registros de la hoja "Datos" en formato JSON.
 *  - POST → agrega un registro nuevo al final de la hoja "Datos".
 *
 * CÓMO USARLO:
 *  1. En tu Google Sheet: Extensiones → Apps Script.
 *  2. Borrá el contenido de Code.gs que aparece por defecto y pegá TODO este archivo.
 *  3. En la línea de CLAVE_SECRETA, cambiá el texto por una clave inventada por vos
 *     (no hace falta que sea complicada, es solo para que no cualquiera pueda escribir).
 *  4. Guardá (ícono de disquete).
 *  5. Implementar → Nueva implementación → tipo "Aplicación web".
 *       - Ejecutar como: Yo (tu cuenta)
 *       - Quién tiene acceso: Cualquier usuario
 *  6. Autorizá los permisos cuando te lo pida Google (es tu propio script, es seguro aceptarlo).
 *  7. Copiá la URL que te da ("URL de la aplicación web") y pasámela para conectarla con la página.
 *
 * Si en el futuro cambiás el código, tenés que hacer "Nueva implementación" de nuevo
 * (o "Gestionar implementaciones" → editar → nueva versión) para que los cambios se apliquen.
 */

const NOMBRE_HOJA = "Datos";
const CLAVE_SECRETA = "CAMBIAR_ESTA_CLAVE"; // <-- reemplazá esto por tu propia clave

const COLUMNAS = [
  "Año", "Provincia", "URE N°", "Trimestre", "ID PLAN", "ID POA",
  "Fecha de monitoreo en campo", "Alcanzado por SAT?",
  "Extensionistas ANA", "Acompañantes en campo", "IF",
  "Observaciones relevantes para una devolución al ALA"
];

function doGet(e) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  const valores = hoja.getDataRange().getValues();
  const encabezados = valores[0];
  const filas = valores.slice(1);

  const registros = filas
    .filter(fila => fila.some(celda => celda !== "" && celda !== null))
    .map(fila => {
      const obj = {};
      encabezados.forEach((col, i) => {
        let v = fila[i];
        if (Object.prototype.toString.call(v) === "[object Date]") {
          v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        obj[col] = v === "" || v === null || v === undefined ? null : v;
      });
      return obj;
    });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, registros: registros }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);

    if (datos.clave !== CLAVE_SECRETA) {
      return respuestaError("Clave incorrecta.");
    }

    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
    const fila = COLUMNAS.map(col => {
      const v = datos.registro ? datos.registro[col] : "";
      return v === undefined || v === null ? "" : v;
    });

    hoja.appendRow(fila);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return respuestaError("Error al procesar la solicitud: " + err.message);
  }
}

function respuestaError(mensaje) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: mensaje }))
    .setMimeType(ContentService.MimeType.JSON);
}
