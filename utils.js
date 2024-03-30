//Muestra por consola la cuota de mails disponibles
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
/** Oculta Hojas y Filas **/
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










