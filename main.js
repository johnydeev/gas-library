
/**
 *Menu personalizado para la ejecucion de script o funcionalidades requeridas en la/s hoja/s
 **/
function onOpen(){

  let menu = SpreadsheetApp.getUi().createMenu("Crear PDFs")

        menu.addItem("Crear Reporte PDF","crearReportePdf")
            .addItem("Crear Reporte PDF SIN Personalizar","crearPdfSinPersonalizar") 
            .addItem("Crear PDFs y Links Masivos","crearPdfsyLinksMasivos")
            .addItem("Reiniciar PDFs y Links Masivos","reiniciarPdfsyLinksMasivos")
            .addToUi()    
              
  let menu2 = SpreadsheetApp.getUi().createMenu("Enviar Mails")

      menu2.addItem("Enviar mail a una UF especifica", "enviarMail")
           .addItem("Reiniciar envios masivos desde UF especifica","reiniciarMailsMasivos")
           .addItem("Enviar Mails Masivos", "enviarMailsMasivos")
           .addItem("Consultar Cuota de Mails","verCuotaDeMails")
           .addToUi()   
   
  
  let menu3 = SpreadsheetApp.getUi().createMenu("Deudores")

      menu3.addItem("Crear tabla Deudores","crearListaDeudores")
      // menu3.addItem("Ocultar Hojas y Columnas","ocultarHojasyColumnas")
           .addToUi()

  let menu4 = SpreadsheetApp.getUi().createMenu("Recibos")

        menu4.addItem("Crear recibos masivos","crearRecibosMasivos")                        
            .addToUi()

}
//--------------------------------------------------------------------------------------------------------------------
/**
 *Menu personalizado para la ejecucion de script o funcionalidades requeridas en la/s hoja/s
 **/
function onOpen2(){

  let menu = SpreadsheetApp.getUi().createMenu("Crear PDFs")

        // menu.addItem("Crear Reporte PDF","crearReportePdf")
        menu.addItem("Crear Reporte PDF SIN Personalizar","crearReportePdf2") 
            // .addItem("Crear PDFs y Links Masivos","crearPdfsyLinksMasivos")
            .addToUi()
                          
  let menu2 = SpreadsheetApp.getUi().createMenu("Enviar Mails")

      menu2.addItem("Enviar mail a una UF especifica", "enviarMail")
           .addToUi()
  
  // let menu3 = SpreadsheetApp.getUi().createMenu("Mostrar Hojas y Columnas")

  //     menu3.addItem("Mostrar Hojas y Columnas","mostrarHojasyColumnasAH")
  //     // menu3.addItem("Ocultar Hojas y Columnas","ocultarHojasyColumnas")
  //          .addToUi()
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

function ocultarFilasDetalle(rangoCol){

}


