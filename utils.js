/** Muestra un mensaje con la cuota disponible de mails diarios para el envio automatico **/
function verCuotaDeMails(){

  let cuota= MailApp.getRemainingDailyQuota()

  Logger.log(cuota)

  Browser.msgBox("Aun quedan "+cuota+" mails para enviar" ,Browser.Buttons.OK)
}

//----------------------------------------------------------------------------------------------------------------------------
/** oculta hojas y columnas para el seteo y creacion de el/los recibos **/
function ocultarHojasParaRecibo(libro){

  Logger.log("Ejecutando ocultarHojasParaRecibo")
  Logger.log("Ocultando hoja RECIBOS, MAILS, DETALLE DE GASTOS Y PRORRATEO")  
  libro.setActiveSheet(libro.getSheetByName("DEUDORES Y PRORRATEO"))
  libro.getActiveSheet().hideSheet() 
  libro.setActiveSheet(libro.getSheetByName("DETALLE DE GASTOS"))
  libro.getActiveSheet().hideSheet() 
  libro.setActiveSheet(libro.getSheetByName("RECIBOS"))
  libro.getActiveSheet().hideSheet()  
  libro.setActiveSheet(libro.getSheetByName("MAILS"))
  libro.getActiveSheet().hideSheet()  

}
//-----------------------------------------------------------------------------------------------------------------------------
/** muestra hojas y columnas luego del seteo y creacion de el/los recibos **/
function mostrarHojasParaRecibo(libro){

  Logger.log("Ejecutando mostrarHojasParaRecibo")
  Logger.log("Mostrando hoja RECIBOS, MAILS, DETALLE DE GASTOS Y PRORRATEO")  
  libro.setActiveSheet(libro.getSheetByName("DEUDORES Y PRORRATEO"))
  libro.getActiveSheet().showSheet()
  libro.setActiveSheet(libro.getSheetByName("DETALLE DE GASTOS"))
  libro.getActiveSheet().showSheet()
  libro.setActiveSheet(libro.getSheetByName("RECIBOS"))
  libro.getActiveSheet().showSheet()
  libro.setActiveSheet(libro.getSheetByName("MAILS"))
  libro.getActiveSheet().showSheet()
}
//-----------------------------------------------------------------------------------------------------------------------------
/** Oculta Hojas y Filas para generar reportes detallados masivamente**/
function ocultarHojasParaDetalle(libro){

  Logger.log("Ejecutando ocultarHojasParaRecibo")
  Logger.log("Ocultando hoja RECIBOS, MAILS, RECIBOS2 Y PRORRATEO")  
  libro.setActiveSheet(libro.getSheetByName("DEUDORES Y PRORRATEO"))
  libro.getActiveSheet().hideSheet() 
  libro.setActiveSheet(libro.getSheetByName("RECIBOS2"))
  libro.getActiveSheet().hideSheet() 
  libro.setActiveSheet(libro.getSheetByName("RECIBOS"))
  libro.getActiveSheet().hideSheet()  
  libro.setActiveSheet(libro.getSheetByName("MAILS"))
  libro.getActiveSheet().hideSheet()  
}
/** Muestra Hojas y Filas para generar reportes detallados masivamente**/
function mostrarHojasParaDetalle(libro){

  Logger.log("Ejecutando mostrarHojasParaRecibo")
  Logger.log("Mostrando hoja RECIBOS, MAILS, RECIBOS2 Y PRORRATEO")  
  libro.setActiveSheet(libro.getSheetByName("DEUDORES Y PRORRATEO"))
  libro.getActiveSheet().showSheet()
  libro.setActiveSheet(libro.getSheetByName("RECIBOS2"))
  libro.getActiveSheet().showSheet()
  libro.setActiveSheet(libro.getSheetByName("RECIBOS"))
  libro.getActiveSheet().showSheet()
  libro.setActiveSheet(libro.getSheetByName("MAILS"))
  libro.getActiveSheet().showSheet()
}

//----------------------------------------------------------------------------------------------------------------
/**
 *Ocuta y muestra un rango de columnas de la Hoja2 de boedo
 **/
function ocultarColumnasH2Boedo(libro,rangoCol){

  Logger.log("Ejecutando ocultarColumnasH2Boedo")
  Logger.log("Ocultando hoja RECIBOS2")
  Logger.log("Ocultando rango de columnas en Hoja PRORRATEO:")
  Logger.log(rangoCol)
  libro.setActiveSheet(libro.getSheetByName("DEUDORES Y PRORRATEO"))
  libro.getActiveSheet().hideColumn(libro.getRange(rangoCol))
  libro.setActiveSheet(libro.getSheetByName("RECIBOS2"))
  libro.getActiveSheet().hideSheet()

}
function mostrarColumnasH2Boedo(libro,rangoCol){

  Logger.log("Ejecutando mostrarColumnasH2Boedo")
  Logger.log("Mostrando hoja RECIBOS2")
  Logger.log("Mostrando rango de columnas en Hoja PRORRATEO:")
  Logger.log(rangoCol)
  libro.setActiveSheet(libro.getSheetByName("DEUDORES Y PRORRATEO"))
  libro.getActiveSheet().unhideColumn(libro.getRange(rangoCol))
  libro.setActiveSheet(libro.getSheetByName("RECIBOS2"))
  libro.getActiveSheet().showSheet()

}

//--------------------------------------------------------------------------------------------------------------------------------
/**
 *Oculta Hojas y columnas que no se desea que esten en el reporte PDF
 **/
function ocultarHojasyColumnasAH(libro,rangoCol){

  Logger.log("Ejecutando ocultarHojasyColumnasAH")
  Logger.log("Ocultando Hojas, RECIBOS y MAILS")
  Logger.log("Ocultando rango de columnas en DETALLE:")
  Logger.log(rangoCol)
  libro.setActiveSheet(libro.getSheetByName("RECIBOS"))
  libro.getActiveSheet().hideSheet()
  libro.setActiveSheet(libro.getSheetByName("MAILS"))
  libro.getActiveSheet().hideSheet()
  libro.setActiveSheet(libro.getSheetByName("DETALLE DE GASTOS"))
  libro.getActiveSheet().hideColumn(libro.getRange(rangoCol))
  
}

function mostrarHojasyColumnasAH(libro,rangoCol){

  Logger.log("Ejecutando mostrarHojasyColumnasAH")
  Logger.log("Mostrando Hojas, RECIBOS y MAILS")
  Logger.log("Mostrando rango de columnas en DETALLE:")
  Logger.log(rangoCol)
  libro.setActiveSheet(libro.getSheetByName("RECIBOS"))
  libro.getActiveSheet().showSheet()
  libro.setActiveSheet(libro.getSheetByName("MAILS"))
  libro.getActiveSheet().showSheet()
  libro.setActiveSheet(libro.getSheetByName("DETALLE DE GASTOS"))
  libro.getActiveSheet().unhideColumn(libro.getRange(rangoCol))
}









