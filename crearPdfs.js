/**
 * Esta funcion recibe como parametro una URL y le se le setea valores 
 * para luego devolver un archivo blob como un PDF formateado 
 **/
function crearPdf(url) {

  let exportarUrl = url.replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf'
    + '&format=pdf'
    + '&size=A4'
    + '&portrait=true'
    + '&fitw=true'
    //+ '&scale=1' + //Ajustes {1=100%,2=Ancho,3=Alto,4=Página}       
    + '&top_margin=0.5'
    + '&bottom_margin=0.5'
    + '&left_margin=0.2'
    + '&right_margin=0.2'
    + '&sheetnames=false'
    + '&printtitle=false'
    // + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
    + '&gridlines=true'
    + '&fzr=FALSE'

  var respuesta = UrlFetchApp.fetch(exportarUrl, {
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  })
  return respuesta
}
//----------------------------------------------------------------------------------------------
/**Crea un reporte PDF con el nombre del archivo y la unidad funcional concatenada 
 * eliminando antes de crearlo el reporte personalizado
 **/
function crearReportePdf2(libro, carpeta, rangoCol) {

  let nombrelibro = libro.getName()

  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()

  let url = libro.getUrl()
  let blob = crearPdf(url)
  carpeta.createFile(blob).setName(nombrelibro)

  mostrarHojasyColumnasAH(libro, rangoCol)
}

//-----------------------------------------------------------------------------------------------
/** Crea un reporte PDF con el nombre del archivo y la unidad funcional concatenada **/
function crearReportePdf(libro, carpeta, celdaNombre, celdaUF, rangoCol) {

  let nombrelibro = libro.getName()
  let espacio = " "
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  let nombreUf = hojaDetalle.getRange(celdaNombre).getValue()
  let numeroUF = hojaDetalle.getRange(celdaUF).getValue()

  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()

  let url = libro.getUrl()
  let blob = crearPdf(url)
  carpeta.createFile(blob).setName("UF" + numeroUF + espacio + nombrelibro + espacio + nombreUf)

  mostrarHojasyColumnasAH(libro, rangoCol)
}
//---------------------------------------------------------------------------------------------------------------------