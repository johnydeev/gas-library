/**
 * crea reporte PDF de cada UF y links, luego extrae y pega el id del PDF y el link del mismo para ubicarlos 
 * en la columna y fila correspondiente en la HojaMails para su posterior uso.
 **/
function crearPdfsyLinksMasivos (libro, carpeta, hojaMails, cantUF, celdaNombre, celdaUF, rangoCol, mes) {

  let nombrelibro = libro.getName()
  let espacio = " "
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = libro.getSheetByName("DEUDORES Y PRORRATEO")
  let rangoUF = hojaProrrateo.getRange(6, 1, cantUF).getValues()
  let mesActual = hojaProrrateo.getRange(mes).getValue()
  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()
  let carpetaMes = carpeta.createFolder(mesActual)
  let url = libro.getUrl()
  //--------------------- Variables preparadas

  for (let i = 0; i < cantUF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(3000)
    hojaDetalle.getRange(celdaUF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(2000)
    let nombreUF = hojaDetalle.getRange(celdaNombre).getValue()
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

    hojaMails.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // hojaMails.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    hojaMails.getRange(i + 2, 6).setValue(archivo.getId())

  }
  mostrarHojasyColumnasAH(libro, rangoCol)
}

//-------------------------------------------------------------------------------------------------------------------------------------------
/**
 * Reinicia los reportes PDF desde una UF especifica justo con la creacion de links, luego extrae y pega el id del PDF y el link 
 * del mismo para ubicarlos en la columna y fila correspondiente en la HojaMails para su posterior uso
 **/
function reiniciarPdfsyLinksMasivos(libro, carpeta, hojaMails, cantUF, celdaNombre, celdaUF, rangoCol, mes, uf) {

  let nombrelibro = libro.getName()
  let espacio = " "
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = libro.getSheetByName("DEUDORES Y PRORRATEO")
  let rangoUF = hojaProrrateo.getRange(6, 1, cantUF).getValues()
  let mesActual = hojaProrrateo.getRange(mes).getValue()
  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()
  let carpetaMes = carpeta.createFolder(mesActual)
  //--------------------- Variables preparadas

  Logger.log("Mostrando uf:")
  Logger.log(uf)

  for (let i = uf; i < cantUF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(3000)
    hojaDetalle.getRange(celdaUF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(2000)
    let nombreUF = hojaDetalle.getRange(celdaNombre).getValue()
    let url = libro.getUrl()
    let blob = crearPdf(url)
    Utilities.sleep(3000)

    let archivo = carpetaMes.createFile(blob).setName("UF" + rangoUF[i] + espacio + nombrelibro + espacio + nombreUF)

    Logger.log(archivo.getName() + rangoUF[i])
    Logger.log("archivo")
    Logger.log(archivo)
    Logger.log(archivo.getDownloadUrl())
    Logger.log(archivo.getId())

    hojaMails.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // hojaMails.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    hojaMails.getRange(i + 2, 6).setValue(archivo.getId())

  }
  mostrarHojasyColumnasAH(libro, rangoCol)
}

//-------------------------------------------------------------------------------------------------------
/**
 * Crea un reporte PDF del detalle personalizado solamente
 **/
function crearDetallePdfsyLinksMasivos(libro, carpeta, hojaMails, cantUF, celdaNombre, celdaUF, mes) {

  let nombrelibro = libro.getName()
  let espacio = " "
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = libro.getSheetByName("DEUDORES Y PRORRATEO")
  let rangoUF = hojaProrrateo.getRange(6, 1, cantUF).getValues()
  let mesActual = hojaProrrateo.getRange(mes).getValue()
  SpreadsheetApp.flush()
  let carpetaMes = carpeta.createFolder(mesActual)
  let url = libro.getUrl()
  hojaDetalle.hideRows(39, 825)
  ocultarHojasParaDetalle(libro)
  //--------------------- Variables preparadas

  for (let i = 0; i < cantUF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(3000)
    hojaDetalle.getRange(celdaUF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(2000)
    let nombreUF = hojaDetalle.getRange(celdaNombre).getValue()
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

    hojaMails.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // hojaMails.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    hojaMails.getRange(i + 2, 6).setValue(archivo.getId())

  }
  rango = hojaDetalle.getRange(39, 1, 825)
  hojaDetalle.unhideRow(rango)
  mostrarHojasParaDetalle(libro)
}

//--------------------------------------------------------------------------------------------------------------------------



