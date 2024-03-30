/** 
 * Crea un array de deudores/morosos a partir de los datos de la hoja DEUDORES Y PRORRATEO y luego retorna el array.
 * Ademas se crea una tabla y se setea algunos estilos para ubicar dentro el array de deudores
**/
function crearListaDeudores(SPREADSHEET,UFS,ROW){

  let hojaProrrateo = SPREADSHEET.getSheetByName("DEUDORES Y PRORRATEO")

  let rangoDeudores = hojaProrrateo.getRange(6,1,UFS,6).getValues()
  console.log("Deudores>>>", rangoDeudores)

  let listaDeudores = []
  console.log("RangoDeudores>>", rangoDeudores.length)
  for (let i = 0 ; i < rangoDeudores.length ; i++){
    
    let dpto = rangoDeudores[i][1]
    let nombreDeudor = rangoDeudores[i][2];
    let montoDeuda = rangoDeudores[i][5];

    // Agregar el objeto a la lista
    if (montoDeuda > 0) {
      listaDeudores.push([dpto, nombreDeudor, montoDeuda]);
    }    
  }
  console.log("LISTA DEUDORES>>", listaDeudores)

  /**  Buscar la hoja por nombre **/
  let detalleDeGastos = SPREADSHEET.getSheetByName("DETALLE DE GASTOS");
  detalleDeGastos.insertRows(ROW);
    
  // Crear un array con los encabezados de la tabla
  let encabezados = ["Departamento", "Nombre", "Monto"];

  // Obtener el rango donde se insertará la tabla
  let rangoDestino = detalleDeGastos.getRange(ROW, 2, listaDeudores.length + 1, encabezados.length);

  // Insertar nuevas filas para la tabla
  
  detalleDeGastos.insertRowsAfter(ROW, listaDeudores.length);

  // Combinar los encabezados y la lista de deudores en un solo array
  let datosTabla = [encabezados, ...listaDeudores];

  console.log("datosTabla>>", datosTabla)
  // Insertar la tabla en la hoja
  rangoDestino.setValues(datosTabla);

  rangoDestino.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  rangoDestino.offset(0, 0, 1, 3).setBackground("#cfd8dc").setFontWeight("bold");

}

