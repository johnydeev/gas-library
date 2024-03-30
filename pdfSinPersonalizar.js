/**
 * Crea un reporte PDF generico sin el apartado personalizado expresado en el intervalo 
 * que comienza en la fila r1 hasta r2 filas hacia abajo.
 **/
function crearPdfSinPersonalizar(libro, carpeta, rangoCol, r1, r2) {

  let nombrelibro = libro.getName()
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  hojaDetalle.hideRows(r1, r2)//---------------- Esta linea no me permite reutilizar la funcion para los edificios que no tienen personalizacion
  SpreadsheetApp.flush()

  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()

  let url = libro.getUrl()
  let blob = crearPdf(url)
  carpeta.createFile(blob).setName(nombrelibro)

  mostrarHojasyColumnasAH(libro, rangoCol)
  rango = hojaDetalle.getRange(r1, 1, r2)
  hojaDetalle.unhideRow(rango)
  Browser.msgBox("Se ah creado un PDF SIN personalizar ", Browser.Buttons.OK)

}
//---------------------------------------------------------------------------------------------------------------------------------