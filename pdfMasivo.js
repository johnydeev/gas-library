/**
 * crea reporte PDF de cada UF y links, luego extrae y pega el id del PDF y el link del mismo para ubicarlos 
 * en la columna y fila correspondiente en la HojaMails para su posterior uso.
 **/
function crearPdfsyLinksMasivos (LIBRO, CARPETA, HOJA_MAILS, CANT_UF, CELDA_NOMBRE, CELDA_UF, RANGO_COL, CELDA_MES) {
  
  let nombrelibro = LIBRO.getName()
  let espacio = " "
  let hojaDetalle = LIBRO.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = LIBRO.getSheetByName("DEUDORES Y PRORRATEO")
  let rangoUF = hojaProrrateo.getRange(6, 1, CANT_UF).getValues()
  let mesActual = hojaProrrateo.getRange(CELDA_MES).getValue()
  let carpetaMes = CARPETA.createFolder(mesActual)
  console.log("mesActual>>",mesActual)
  ocultarHojasyColumnasAH(LIBRO, RANGO_COL)
  SpreadsheetApp.flush()
  let url = LIBRO.getUrl()
  //--------------------- Variables preparadas

  for (let i = 0; i < CANT_UF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(1500)
    hojaDetalle.getRange(CELDA_UF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(1000)
    let nombreUF = hojaDetalle.getRange(CELDA_NOMBRE).getValue()
    // --------------------------------------------------------------------------------------
    let blob = crearPdf(url)
    Utilities.sleep(1500)
    let archivo = carpetaMes.createFile(blob).setName("UF" + rangoUF[i] + espacio + nombrelibro + espacio + nombreUF)

    // -------------------------------------------------------------------------------------------
    // Logger.log(archivo.getName()+ rangoUF[i])
    // Logger.log("archivo")
    // Logger.log(archivo)
    // Logger.log(archivo.getDownloadUrl())
    // Logger.log(archivo.getId())

    HOJA_MAILS.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // HOJA_MAILS.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    HOJA_MAILS.getRange(i + 2, 6).setValue(archivo.getId())

  }
  mostrarHojasyColumnasAH(LIBRO, RANGO_COL)
}

//-------------------------------------------------------------------------------------------------------------------------------------------
/**
 * Reinicia los reportes PDF desde una UF especifica justo con la creacion de links, luego extrae y pega el id del PDF y el link 
 * del mismo para ubicarlos en la columna y fila correspondiente en la HojaMails para su posterior uso
 **/
function reiniciarPdfsyLinksMasivos(LIBRO, CARPETA, HOJA_MAILS, CANT_UF, CELDA_NOMBRE, CELDA_UF, RANGO_COL, CELDA_MES, UF) {

  let nombrelibro = LIBRO.getName()
  let espacio = " "
  let hojaDetalle = LIBRO.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = LIBRO.getSheetByName("DEUDORES Y PRORRATEO")
  let rangoUF = hojaProrrateo.getRange(6, 1, CANT_UF).getValues()
  //obtengo el indice de la UF para reanudar la creacion de PDFs desde donde se interrumpio
  let index = devolverIndiceUF(rangoUF,UF)
  let mesActual = hojaProrrateo.getRange(CELDA_MES).getValue()
  ocultarHojasyColumnasAH(LIBRO, RANGO_COL)
  SpreadsheetApp.flush()
  let carpetaMes = CARPETA.createFolder(mesActual)
  //--------------------- Variables preparadas

  Logger.log("Mostrando uf:")
  Logger.log(UF)

  for (let i = index; i < CANT_UF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(1500)
    hojaDetalle.getRange(CELDA_UF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(1000)
    let nombreUF = hojaDetalle.getRange(CELDA_NOMBRE).getValue()
    let url = LIBRO.getUrl()
    let blob = crearPdf(url)
    Utilities.sleep(1500)

    let archivo = carpetaMes.createFile(blob).setName("UF" + rangoUF[i] + espacio + nombrelibro + espacio + nombreUF)

    Logger.log(archivo.getName() + rangoUF[i])
    Logger.log("archivo")
    Logger.log(archivo)
    Logger.log(archivo.getDownloadUrl())
    Logger.log(archivo.getId())

    HOJA_MAILS.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // HOJA_MAILS.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    HOJA_MAILS.getRange(i + 2, 6).setValue(archivo.getId())

  }
  mostrarHojasyColumnasAH(LIBRO, RANGO_COL)
}

//-------------------------------------------------------------------------------------------------------
/**
 * Crea un reporte PDF del detalle personalizado solamente
 **/
function crearDetallePdfsyLinksMasivos(LIBRO, CARPETA, HOJA_MAILS, CANT_UF, CELDA_NOMBRE, CELDA_UF, CELDA_MES) {

  let nombrelibro = LIBRO.getName()
  let espacio = " "
  let hojaDetalle = LIBRO.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = LIBRO.getSheetByName("DEUDORES Y PRORRATEO")
  let rangoUF = hojaProrrateo.getRange(6, 1, CANT_UF).getValues()
  let mesActual = hojaProrrateo.getRange(CELDA_MES).getValue()
  SpreadsheetApp.flush()
  let carpetaMes = CARPETA.createFolder(mesActual)
  let url = LIBRO.getUrl()
  hojaDetalle.hideRows(39, 825)
  ocultarHojasParaDetalle(LIBRO)
  //--------------------- Variables preparadas

  for (let i = 0; i < CANT_UF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])
    Utilities.sleep(1500)
    hojaDetalle.getRange(CELDA_UF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(1000)
    let nombreUF = hojaDetalle.getRange(CELDA_NOMBRE).getValue()
    // --------------------------------------------------------------------------------------
    let blob = crearPdf(url)
    Utilities.sleep(1500)
    let archivo = carpetaMes.createFile(blob).setName("UF" + rangoUF[i] + espacio + nombrelibro + espacio + nombreUF)

    // -------------------------------------------------------------------------------------------
    // Logger.log(archivo.getName()+ rangoUF[i])
    // Logger.log("archivo")
    // Logger.log(archivo)
    // Logger.log(archivo.getDownloadUrl())
    // Logger.log(archivo.getId())

    HOJA_MAILS.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // HOJA_MAILS.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    HOJA_MAILS.getRange(i + 2, 6).setValue(archivo.getId())

  }
  rango = hojaDetalle.getRange(39, 1, 825)
  hojaDetalle.unhideRow(rango)
  mostrarHojasParaDetalle(LIBRO)
}

//--------------------------------------------------------------------------------------------------------------------------
/**
 * Reinicia los reportes PDF desde una UF especifica justo con la creacion de links, luego extrae y pega el id del PDF y el link 
 * del mismo para ubicarlos en la columna y fila correspondiente en la HojaMails para su posterior uso
 **/
function reiniciarDetallePdfsyLinksMasivos(LIBRO, CARPETA, HOJA_MAILS, CANT_UF, CELDA_NOMBRE, CELDA_UF, CELDA_MES, UF) {

  let nombrelibro = LIBRO.getName()
  let espacio = " "
  let hojaDetalle = LIBRO.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = LIBRO.getSheetByName("DEUDORES Y PRORRATEO")
  let rangoUF = hojaProrrateo.getRange(6, 1, CANT_UF).getValues()
  //obtengo el indice de la UF para reanudar la creacion de PDFs desde donde se interrumpio
  let index = devolverIndiceUF(rangoUF,UF)
  console.log("El indice de la UF:",UF," es: ",index)
  let mesActual = hojaProrrateo.getRange(CELDA_MES).getValue()  
  let carpetaMes = CARPETA.createFolder(mesActual)
  let url = LIBRO.getUrl() 
  //--------------------- Variables preparadas
  hojaDetalle.hideRows(39, 825)
  ocultarHojasParaDetalle(LIBRO)
  //--------------------- Filas y Hojas ocultas.

  console.log("Mostrando Uf ingresada: ", UF)

  for (let i = index; i < CANT_UF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(1500)
    hojaDetalle.getRange(CELDA_UF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(1000)
    let nombreUF = hojaDetalle.getRange(CELDA_NOMBRE).getValue()
    let blob = crearPdf(url)
    Utilities.sleep(1500)

    let archivo = carpetaMes.createFile(blob).setName("UF" + rangoUF[i] + espacio + nombrelibro + espacio + nombreUF)

    Logger.log(archivo.getName() + rangoUF[i])
    Logger.log("archivo")
    Logger.log(archivo)
    Logger.log(archivo.getDownloadUrl())
    Logger.log(archivo.getId())

    HOJA_MAILS.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // HOJA_MAILS.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    HOJA_MAILS.getRange(i + 2, 6).setValue(archivo.getId())

  }
  rango = hojaDetalle.getRange(39, 1, 825)
  hojaDetalle.unhideRow(rango)
  mostrarHojasParaDetalle(LIBRO)
}


