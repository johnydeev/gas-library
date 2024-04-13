/**
 * Funcion para enviar un mail a una Unidad Funcional especifica, 
 * recibe como parametro la UF y la Hoja de Mails (La UF se ingresa por prompt)
**/
//TODO --------------------------
// 1- LLAMAR A LA FUNCION devulveIndice para refactorizar el codigo
// 2- Extraer solo la columna UF para recorrer el rango y no toda la tabla ya que se utiliza la primera columna para comparar.
function enviarMail(UF,HOJA_MAILS){  

  let espacio = "] "
  
  let rangoUFyMails = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow() -1,6).getValues()
    
    for(let i=0; i< rangoUFyMails.length; i++){

      if(rangoUFyMails[i][0] == UF){

      let idPDF = rangoUFyMails[i][5]
      let archivo = DriveApp.getFileById(idPDF)
      let blob = archivo.getAs(MimeType.PDF)
      console.log("Nombre de Archivo: ",archivo.getName())            
      let nombreUF = rangoUFyMails[i][2]
      console.log("Nombre UF: ",rangoUFyMails[i][2])
      let mail = rangoUFyMails[i][3]
      console.log("Mail: ",rangoUFyMails[i][3])
        if(mail !== "vacio"){
        GmailApp.sendEmail(mail,"Administración Morinigo UF:"+UF,"Hola "+nombreUF+",le enviamos la liquidación del mes,solicitamos adjuntar el envío del comprobante de pago por favor..  Gracias y que tenga un excelente día.",{
          attachments:[blob]
          })
          return Browser.msgBox("Se envio un mail a UF: ["+rangoUFyMails[i][0]+ espacio + rangoUFyMails[i][2],Browser.Buttons.OK)      
        }else{
          return Browser.msgBox("La unidad funcional no tiene mail cargado",Browser.Buttons.OK)
        }        
      }     
    }    
    Browser.msgBox("No se encontro Unidad Funcional requerida",Browser.Buttons.OK) 

}
//-------------------------------------------------------------------------------------------------------------

function enviarRecibo(UF,HOJA_MAILS,SPREADSHEET){
  let espacio = "] "
    
  let rangoUFyMails = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow() -1,9).getValues()
  let hojaProrrateo = SPREADSHEET.getSheetByName("DEUDORES Y PRORRATEO") 
  let mesActual = hojaProrrateo.getRange("B3").getValue()
  console.log("Mes Actual >>",mesActual)
  
    for(let i=0; i< rangoUFyMails.length; i++){

      if(rangoUFyMails[i][0] == UF ){
        if(rangoUFyMails[i][6] == true ){
          console.log("¿PAGO?>>",rangoUFyMails[i][6])
          console.log("rangoUFyMails[i][8]>>", rangoUFyMails[i][8])
          let idRecibo = rangoUFyMails[i][8]
          console.log("ID recibo",idRecibo)
          let archivo = DriveApp.getFileById(idRecibo)
          // let blob = archivo.getAs(MimeType.PDF)
          console.log("Nombre de Archivo: ",archivo.getName())            
          let nombreUF = rangoUFyMails[i][2]
          console.log("Nombre UF: ",nombreUF)
          let mail = rangoUFyMails[i][3]
          console.log("Mail: ",mail)

         
          SpreadsheetApp.getUi().alert( `¡ATENCION! Se enviará el recibo a:

                          UF: ${UF}
                          Propietario: ${rangoUFyMails[i][2]}

          ACEPTAR SI ESTA SEGURO DE ENVIARLO.`, SpreadsheetApp.getUi().ButtonSet.OK);
          
            // if(mail !== "vacio"){
            // GmailApp.sendEmail(mail,"Administración Morinigo UF:"+UF,"Hola "+nombreUF+",Se le adjunta el recibo correspondiente al mes "+  +" Gracias y que tenga un excelente día.",{
            //   attachments:[blob]
            //   })
              

            //   return Browser.msgBox("Se envio el recibo a UF: ["+UF+ espacio + "Propietario: "+ rangoUFyMails[i][2],Browser.Buttons.OK)      
            // }else{
            //   return Browser.msgBox("La unidad funcional no tiene mail cargado",Browser.Buttons.OK)
            // }        
          HOJA_MAILS.getRange(i + 2, 10).setValue("ENVIADO")
          let fechaActual = new Date()
          console.log("fechaActual>>",fechaActual.toLocaleDateString())
          HOJA_MAILS.getRange(i + 2, 11).setValue(fechaActual)
          return Browser.msgBox("¡Mail Enviado!",Browser.Buttons.OK)
        }else{
          return Browser.msgBox("ATENCION!! Por favor revisar la siguiente informacion: La Unidad Funcional NO realizo el pago",Browser.Buttons.OK)
        }
      }
    }
    Browser.msgBox("No se encontro Unidad Funcional",Browser.Buttons.OK) 
}

//-----------------------------------------------------------------------------------------------------------------------------------
/**
 * Envia un mail a todas la unidades funcionales que se encuentren en la HOJA_MAILS 
 * y omite los envios en las celdas que se encuentre cargado la palabra "vacio" de lo contrario el codigo se rompera
**/
function enviarMailsMasivos(HOJA_MAILS){
             
    let rangoUFyMails = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow()-1,6).getValues()
    //Se utilizo para enviar un segundo archivo adjunto el cual recibe un parametro ID_PDF,agregar el archivo blob2 en el attachment si se utiliza
    // let general = DriveApp.getFileById(ID_PDF)
    // let blob2 = general.getAs(MimeType.PDF)
    //-----------------------------------------------------------
    Logger.log(rangoUFyMails.length)

    for(let i=0 ; i< rangoUFyMails.length ; i++){
      
      Logger.log(rangoUFyMails[i][0])//Unidad Funcional
      Logger.log(rangoUFyMails[i][2])// Nombre
      Logger.log(rangoUFyMails[i][3].toString()) // MAIL
      Logger.log(rangoUFyMails[i][5]) // ID pdf
      let idPDF = rangoUFyMails[i][5]
      let archivo = DriveApp.getFileById(idPDF)
      let blob = archivo.getAs(MimeType.PDF)
      let uf = rangoUFyMails[i][0]
      let nombreUF = rangoUFyMails[i][2]
      let mail = rangoUFyMails[i][3]

        if(mail !== "vacio"){
        GmailApp.sendEmail(mail,"Administración Morinigo UF:"+uf,"Hola "+nombreUF+", "+"le enviamos la liquidación del mes.\n Si usted figura moroso y realizó su pago de expensas, por favor solicitamos adjuntar el envío del comprobante de pago para poder imputar su cancelación. Gracias y que tenga un excelente dia",{
          attachments:[blob]
          })      
        }
        Utilities.sleep(10000)
    } 
  }

//---------------------------------------------------------------------------------
/** Reinicia el envio masivo de mails desde una UF especifica **/
function reiniciarMailsMasivos(UF,HOJA_MAILS){   
             
    let rangoUFyMails = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow()-1,6).getValues()
    let rangoUF = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow()-1,1).getValues()
    //----Se utilizo para enviar un segundo archivo adjunto el cual recibe un parametro ID_PDF
    // let general = DriveApp.getFileById(ID_PDF)
    // let blob2 = general.getAs(MimeType.PDF)
    //-----------------------------------------------------------
    console.log("RangoUfs>>", rangoUF.length)
    let index = devolverIndiceUF(rangoUF,UF)
    console.log("RangoUFyMails>>",rangoUFyMails.length)

    for(let i= index; i< rangoUFyMails.length ; i++){
      console.log("Mostrando Index en For: ",i)
      Logger.log(rangoUFyMails[i][0])//Unidad Funcional
      Logger.log(rangoUFyMails[i][2])// Nombre      
      Logger.log(rangoUFyMails[i][3].toString()) // MAIL
      Logger.log(rangoUFyMails[i][5]) // ID PDF  
      let idPDF = rangoUFyMails[i][5]
      let archivo = DriveApp.getFileById(idPDF)
      let blob = archivo.getAs(MimeType.PDF)
      let uf = rangoUFyMails[i][0]
      let nombreUF = rangoUFyMails[i][2]
      let mail = rangoUFyMails[i][3]

        if(mail !== "vacio"){
        GmailApp.sendEmail(mail,"Administración Morinigo UF:"+uf,"Hola "+nombreUF+", "+"le enviamos la liquidación del mes.Si usted figura moroso y realizó su pago de expensas, por favor solicitamos adjuntar el envío del comprobante de pago para poder imputar su cancelación. Gracias y que tenga un excelente dia",{
          attachments:[blob]
          })      
        }
        Utilities.sleep(10000)
       Logger.log("Mail Enviado")
    } 
  }