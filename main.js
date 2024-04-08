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
//--------------------------------------------------------------------------------------------------------------------
/**
 *Menu personalizado para la ejecucion de script o funcionalidades requeridas en la/s hoja/s
 **/
function onOpen3(){

  let menu = SpreadsheetApp.getUi().createMenu("Crear PDFs")

        menu.addItem("Crear Reporte PDF","crearReportePdf")
            .addItem("Crear Reporte PDF SIN Personalizar","crearPdfSinPersonalizar") 
            .addItem("Crear PDFs y Links Masivos","crearDetallePdfsyLinksMasivos")
            .addItem("Reiniciar PDFs y Links Masivos","reiniciarDetallePdfsyLinksMasivos")
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
             .addItem("Reiniciar recibos masivos","reiniciarRecibosMasivos")                       
             .addToUi()

}


