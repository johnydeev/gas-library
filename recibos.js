//-----------------------------------------------------------------------------------------

function crearRecibosMasivos(FOLDER,SPREADSHEET,HOJA_RECIBO,HOJA_MAILS,cantUF,celdaNombre,celdaUF){

    let nombrelibro = "RECIBO DE"
    let espacio = " "

    let hojaProrrateo = SPREADSHEET.getSheetByName("DEUDORES Y PRORRATEO")  
    let rangoUF = hojaProrrateo.getRange(6,1,cantUF).getValues()

    ocultarHojasParaRecibo(SPREADSHEET)
    SpreadsheetApp.flush()
    
    for(let i = 0 ; i< cantUF ; i++){       
      
      Utilities.sleep(3000)     
      HOJA_RECIBO.getRange(celdaUF).setValue(rangoUF[i])
      SpreadsheetApp.flush()
      Utilities.sleep(2000)     
      let nombreUF = HOJA_RECIBO.getRange(celdaNombre).getValue()           
      let url = SPREADSHEET.getUrl()
      let blob = crearPdf(url)
      Utilities.sleep(3000)     
      let archivo = FOLDER.createFile(blob).setName("UF"+rangoUF[i]+espacio+nombrelibro + espacio + nombreUF)       

      Logger.log(archivo.getName()+ rangoUF[i])
      Logger.log(archivo.getDownloadUrl())
      Logger.log(archivo.getId())

      HOJA_MAILS.getRange(i+2,7).setValue(archivo.getDownloadUrl())      
      HOJA_MAILS.getRange(i+2,8).setValue(archivo.getId())

    }
    mostrarHojasParaRecibo(SPREADSHEET)
}