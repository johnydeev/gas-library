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
  ocultarHojasyColumnasAH(LIBRO, RANGO_COL)
  SpreadsheetApp.flush()
  let carpetaMes = CARPETA.createFolder(mesActual)
  let url = LIBRO.getUrl()
  //--------------------- Variables preparadas

  for (let i = 0; i < CANT_UF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(3000)
    hojaDetalle.getRange(CELDA_UF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(2000)
    let nombreUF = hojaDetalle.getRange(CELDA_NOMBRE).getValue()
    // --------------------------------------------------------------------------------------
    let blob = crearPdf(url)
    Utilities.sleep(3000)
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
function reiniciarPdfsyLinksMasivos(LIBRO, CARPETA, HOJA_MAILS, CANT_UF, CELDA_NOMBRE, CELDA_UF, RANGO_COL, CELDA_MES, uf) {

  let nombrelibro = LIBRO.getName()
  let espacio = " "
  let hojaDetalle = LIBRO.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = LIBRO.getSheetByName("DEUDORES Y PRORRATEO")
  let rangoUF = hojaProrrateo.getRange(6, 1, CANT_UF).getValues()
  let mesActual = hojaProrrateo.getRange(CELDA_MES).getValue()
  ocultarHojasyColumnasAH(LIBRO, RANGO_COL)
  SpreadsheetApp.flush()
  let carpetaMes = CARPETA.createFolder(mesActual)
  //--------------------- Variables preparadas

  Logger.log("Mostrando uf:")
  Logger.log(uf)

  for (let i = uf; i < CANT_UF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(3000)
    hojaDetalle.getRange(CELDA_UF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(2000)
    let nombreUF = hojaDetalle.getRange(CELDA_NOMBRE).getValue()
    let url = LIBRO.getUrl()
    let blob = crearPdf(url)
    Utilities.sleep(3000)

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

    Utilities.sleep(3000)
    hojaDetalle.getRange(CELDA_UF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(2000)
    let nombreUF = hojaDetalle.getRange(CELDA_NOMBRE).getValue()
    // --------------------------------------------------------------------------------------
    let blob = crearPdf(url)
    Utilities.sleep(3000)
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



