/**
 * Crea todos los recibos de cada inquilino
 **/
function crearRecibosMasivos(CARPETA,SPREADSHEET,HOJA_RECIBO,HOJA_MAILS,CANT_UF,CELDA_PROPIETARIO,CELDA_UF,CELDA_MES){
    
    //Se crean las variables necesarias para la creacion de los recibos
    let nombrePdf = "RECIBO DE"
    let espacio = " "
    let hojaProrrateo = SPREADSHEET.getSheetByName("DEUDORES Y PRORRATEO")
    let rangoUF = hojaProrrateo.getRange(6,1,CANT_UF).getValues()
    //Se crea una carpeta con el nombre de el mes actual extraido de una celda para luego guardar los PDFs generados
    let mesActual = hojaProrrateo.getRange(CELDA_MES).getValue()
    let carpetaMes = CARPETA.createFolder("RECIBOS "+mesActual)
    console.log("mesActual>>",mesActual)
    //-----------------------------------------------------------------------    
    ocultarHojasParaRecibo(SPREADSHEET)
    SpreadsheetApp.flush()
    console.log("en la puerta del for")

    for(let i = 0 ; i < CANT_UF ; i++){
      console.log("Entro al for..")
      Utilities.sleep(1500)
      HOJA_RECIBO.getRange(CELDA_UF).setValue(rangoUF[i])
      SpreadsheetApp.flush()
      Utilities.sleep(1000)
      let nombreUF = HOJA_RECIBO.getRange(CELDA_PROPIETARIO).getValue()
      let url = SPREADSHEET.getUrl()
      let blob = crearPdf(url)
      Utilities.sleep(1500)
      let archivo = carpetaMes.createFile(blob).setName("UF"+rangoUF[i]+espacio+nombrePdf + espacio + nombreUF)

      Logger.log(archivo.getName()+ rangoUF[i])
      Logger.log(archivo.getDownloadUrl())
      Logger.log(archivo.getId())

      HOJA_MAILS.getRange(i+2,8).setValue(archivo.getDownloadUrl())
      HOJA_MAILS.getRange(i+2,9).setValue(archivo.getId())

    }
    mostrarHojasParaRecibo(SPREADSHEET)
}
/**
 * Crea un recibo de una UF especifica
 **/

function crearRecibo(CARPETA,SPREADSHEET,HOJA_RECIBO,HOJA_MAILS,CANT_UF,CELDA_PROPIETARIO,CELDA_UF,UF){
    
    //Se crean las variables necesarias para la creacion de los recibos
    let nombrePdf = "RECIBO DE"
    let espacio = " "
    let hojaProrrateo = SPREADSHEET.getSheetByName("DEUDORES Y PRORRATEO")
    let rangoUF = hojaProrrateo.getRange(6,1,CANT_UF).getValues()
    let index = devolverIndiceUF(rangoUF,UF)
    //-----------------------------------------------------------------------    
    ocultarHojasParaRecibo(SPREADSHEET)
    SpreadsheetApp.flush()         
    
    HOJA_RECIBO.getRange(CELDA_UF).setValue(UF)
    SpreadsheetApp.flush()
     
    let nombreUF = HOJA_RECIBO.getRange(CELDA_PROPIETARIO).getValue()
    let url = SPREADSHEET.getUrl()
    let blob = crearPdf(url)
     
    let archivo = CARPETA.createFile(blob).setName("UF"+ UF +espacio+nombrePdf + espacio + nombreUF)

    Logger.log(archivo.getName()+ UF)
    Logger.log(archivo.getDownloadUrl())
    Logger.log(archivo.getId())

    HOJA_MAILS.getRange(index+2,8).setValue(archivo.getDownloadUrl())
    HOJA_MAILS.getRange(index+2,9).setValue(archivo.getId())

    
    mostrarHojasParaRecibo(SPREADSHEET)
}
/**
 * Reinicia la creacion de recibos desde una uf dada
 **/
function reiniciarRecibosMasivos(CARPETA,SPREADSHEET,HOJA_RECIBO,HOJA_MAILS,CANT_UF,CELDA_PROPIETARIO,CELDA_UF,UF){

    let nombrelibro = "RECIBO DE"
    let espacio = " "
    let hojaProrrateo = SPREADSHEET.getSheetByName("DEUDORES Y PRORRATEO")  
    let rangoUF = hojaProrrateo.getRange(6,1,CANT_UF).getValues()
    let index = devolverIndiceUF(rangoUF,UF)
    console.log("El indice de la UF:",UF," es: ",index)
    ocultarHojasParaRecibo(SPREADSHEET)
    SpreadsheetApp.flush()
    
    for(let i = index ; i< CANT_UF ; i++){
      
      Utilities.sleep(1500)     
      HOJA_RECIBO.getRange(CELDA_UF).setValue(rangoUF[i])
      SpreadsheetApp.flush()
      Utilities.sleep(1000)     
      let nombreUF = HOJA_RECIBO.getRange(CELDA_PROPIETARIO).getValue()           
      let url = SPREADSHEET.getUrl()
      let blob = crearPdf(url)
      Utilities.sleep(1500)     
      let archivo = CARPETA.createFile(blob).setName("UF"+rangoUF[i]+espacio+nombrelibro + espacio + nombreUF)       

      Logger.log(archivo.getName()+ rangoUF[i])
      Logger.log(archivo.getDownloadUrl())
      Logger.log(archivo.getId())

      HOJA_MAILS.getRange(i+2,8).setValue(archivo.getDownloadUrl())      
      HOJA_MAILS.getRange(i+2,9).setValue(archivo.getId())

    }
    mostrarHojasParaRecibo(SPREADSHEET)
}