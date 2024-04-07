/**
 * Esta funcion recibe como parametro una URL y le se le setea valores 
 * para luego devolver un archivo blob como un PDF formateado 
 **/
function crearPdf(url) {
  try {
    let exportarUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf'
      + '&format=pdf'
      + '&size=A4'
      + '&portrait=true'
      + '&fitw=true'
      + '&top_margin=0.5'
      + '&bottom_margin=0.5'
      + '&left_margin=0.2'
      + '&right_margin=0.2'
      + '&sheetnames=false'
      + '&printtitle=false'
      + '&gridlines=true'
      + '&fzr=FALSE';
    
    const params = {      
      muteHttpExceptions: true,
      headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } 
    }
    let flag = 0
    while(flag == 0){
      let respuesta = UrlFetchApp.fetch(exportarUrl, params);
      let blob = respuesta.getBlob();
      let miblob = blob.getContentType().toString()
      console.log("miblob>>>",miblob)
      if (miblob == "application/pdf"){      
        return blob
        
      }else{
        
        continue
      }
    }
  } catch (error) {
    Logger.log("Error en crearPdf: " + error.message);
  }
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