/**
 * Funcion para enviar un mail a una Unidad Funcional especifica, 
 * recibe como parametro la UF y la Hoja de Mails (La UF se ingresa por prompt)
**/
function enviarMail(UF,HOJA_MAILS){  

  let espacio = "] "
    
  let rangoUFyMails = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow() -1,6).getValues()
    
    for(let i=0; i< rangoUFyMails.length; i++){

      if(rangoUFyMails[i][0] == UF){

      let idPDF = rangoUFyMails[i][5]
      let archivo = DriveApp.getFileById(idPDF)
      let blob = archivo.getAs(MimeType.PDF)
      Logger.log("archivo")
      Logger.log(archivo)
      Logger.log("blob")
      Logger.log(blob)
      let nombreUF = rangoUFyMails[i][2]
      let mail = rangoUFyMails[i][3]
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

//-----------------------------------------------------------------------------------------------------------------------------------
/**
 * Envia un mail a todas la unidades funcionales que se encuentren en la HOJA_MAILS 
 * y omite los envios en las celdas que se encuentre cargado la palabra "vacio" de lo contrario el codigo se rompera
**/
function enviarMailsMasivos(HOJA_MAILS){   

             
    let rangoUFyMails = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow()-1,6).getValues()      
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
        GmailApp.sendEmail(mail,"Administración Morinigo UF:"+uf,"Hola "+nombreUF+" ,le enviamos la liquidacion del mes, solicitamos adjuntar el envio del comprobante de pago por favor.. Gracias y que tenga un excelente dia",{
          attachments:[blob]
          })      
        }
        Utilities.sleep(15000)
    } 
  }

//---------------------------------------------------------------------------------
/** Reinicia el envio masivo de mails desde una UF especifica **/
function reiniciarMailsMasivos(uf,hojaMails){   
             
    let rangoUFyMails = hojaMails.getRange(2,1,hojaMails.getLastRow()-1,6).getValues()      
    console.log("RangoUFs>>",rangoUFyMails.length)

    for(let i=uf-1 ; i< rangoUFyMails.length ; i++){
      
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
        GmailApp.sendEmail(mail,"Administración Morinigo UF:"+uf,"Hola "+nombreUF+" ,te dejo la liquidacion del mes, le solicitamos adjuntar el envio del comprobante de pago por favor.. Gracias y que tenga un excelente dia",{
          attachments:[blob]
          })      
        }
        Utilities.sleep(15000)
       Logger.log("Mail Enviado")
    } 
  }