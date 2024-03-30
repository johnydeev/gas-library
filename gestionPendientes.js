/** 
 * Crea un reporte de los pendientes en los que se encuentran los trabajos de un edificio especifico 
 * tomado de la hoja activa en donde se encuentre ubicado el usuario 
 **/
function crearReportePdfenHoja(FOLDER, COLUMRANGE) {

  let SS = SpreadsheetApp.getActiveSpreadsheet();
  let SHEET = SS.getActiveSheet()
  // Logger.log(SHEET.getName())
  ocultarColumnasReporte(SHEET, COLUMRANGE)
  SpreadsheetApp.flush()
  exportSheetAsPDF(FOLDER)

  mostrarColumnasReporte(SHEET, COLUMRANGE)
}
//--------------------------------------------------------------------------------------------------------------------
/** 
 * Crea un archivo PDF a partir de una hoja dada por el usuario en la que se encuentra activa 
 * la cual representa los pendientes y trabajos de un edificio especifico 
**/
function exportSheetAsPDF(FOLDER) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let sheetId = sheet.getSheetId();
  let spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  // Configurar parámetros para la exportación
  let url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?';
  let exportOptions =
    'exportFormat=pdf' +             // formato pdf
    '&format=pdf' +
    '&size=letter' +                 // tamaño papel (puede ser A4, letter, etc.)
    '&portrait=true' +              // orientación (puede ser false para horizontal)
    '&fitw=true' +                   // ajustar al ancho
    '&top_margin=0.5' +
    '&bottom_margin=0.5' +
    '&left_margin=0.22' +
    '&right_margin=0.22' +
    '&sheetnames=false&printtitle=false' +  // ocultar nombres y títulos de hojas
    '&pagenumbers=false&gridlines=false' +  // ocultar números de página y líneas de cuadrícula
    '&fzr=false' +                  // repetir filas congeladas
    '&gid=' + sheetId;              // ID de la hoja que quieres exportar

  let response = UrlFetchApp.fetch(url + exportOptions, {
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  });

  // Crear el PDF en Google Drive con el contenido del archivo PDF
  FOLDER.createFile(response.getBlob()).setName(sheet.getName());
}

//---------------------------------------------------------------------------------------------------
function ocultarColumnasReporte(SHEET) { 
  
  SHEET.hideColumn(SHEET.getRange("G:Z"));  
}

function mostrarColumnasReporte(SHEET) {
  
  SHEET.unhideColumn(SHEET.getRange("G:Z"))
}
