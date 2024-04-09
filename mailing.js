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
function enviarMailsMasivos(HOJA_MAILS,ID_PDF){   

             
    let rangoUFyMails = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow()-1,6).getValues()
    let general = DriveApp.getFileById(ID_PDF)
    let blob2 = general.getAs(MimeType.PDF)
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
        GmailApp.sendEmail(mail,"Administración Morinigo UF:"+uf,"Hola "+nombreUF+" ,le enviamos la liquidación del mes. En ella encontrarán una cuota extraordinaria detallada en la misma para afrontar el cateo de la medianera de Santiago del Estero y a su vez, un incremento en las expensas ordinarias debido a los incrementos salariales y servicios que presenta el consorcio.\nSi usted figura moroso y realizó su pago de expensas, por favor solicitamos adjuntar el envío del comprobante de pago para poder imputar su cancelación, desde ya le pedimos las disculpas correspondientes por las molestias que le podamos ocasionar. Gracias y que tenga un excelente dia",{
          attachments:[blob, blob2]
          })      
        }
        Utilities.sleep(10000)
    } 
  }

//---------------------------------------------------------------------------------
/** Reinicia el envio masivo de mails desde una UF especifica **/
function reiniciarMailsMasivos(UF,HOJA_MAILS, ID_PDF){   
             
    let rangoUFyMails = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow()-1,6).getValues()
    let rangoUF = HOJA_MAILS.getRange(2,1,HOJA_MAILS.getLastRow()-1,1).getValues()
    let general = DriveApp.getFileById(ID_PDF)
    let blob2 = general.getAs(MimeType.PDF)
    console.log("RangoUfs>>", rangoUF.length)
    let index = devolverIndiceUF(rangoUF,UF)
    console.log("RangoUFyMails>>",rangoUFyMails.length)

    for(let i= index; i< rangoUFyMails.length ; i++){
      console.log("Mostrando Index en For: ",i)
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
        GmailApp.sendEmail(mail,"Administración Morinigo UF:"+uf,"Hola "+nombreUF+" ,le enviamos la liquidación del mes. En ella encontrarán una cuota extraordinaria detallada en la misma para afrontar el cateo de la medianera de Santiago del Estero y a su vez, un incremento en las expensas ordinarias debido a los incrementos salariales y servicios que presenta el consorcio.\nSi usted figura moroso y realizó su pago de expensas, por favor solicitamos adjuntar el envío del comprobante de pago para poder imputar su cancelación, desde ya le pedimos las disculpas correspondientes por las molestias que le podamos ocasionar. Gracias y que tenga un excelente dia",{
          attachments:[blob, blob2]
          })      
        }
        Utilities.sleep(10000)
       Logger.log("Mail Enviado")
    } 
  }