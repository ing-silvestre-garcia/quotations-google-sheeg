function myFunction() {
  recorrePartidas()
}

function recorrePartidas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaResumen = ss.getSheetByName('Resumen'); // Nombre de tu hoja de resumen
  var cantidadProducto = hojaResumen.getRange('B2').getValues(); // Obtiene los valores desde A5 hasta el final
  var data = hojaResumen.getRange('A5:A').getValues(); // Obtiene los valores desde A5 hasta el final

  // Filtra los valores vacíos o nulos al final de la columna
  var nuevasPartidas = data.filter(function(row) {
    return row[0] !== '' && row[0] !== null;
  });

  nuevasPartidas.forEach(function(partida) {
    var nombrePartida = partida[0];
    crearNuevasHojas(ss,nombrePartida,cantidadProducto)
  });

  nuevasPartidas.forEach(function(partida, index) {
      var nombrePartida = partida[0]; // Accede al valor de la celda que contiene el nombre de la partida
      var filaResumen = index + 5; // Utiliza el índice del bucle para obtener la fila actual

      var celdaResumen = hojaResumen.getRange('C' + filaResumen); // Obtiene la celda en la columna C
      celdaResumen.setFormula('=' + nombrePartida + '!B8'); // Referencia a la nueva hoja
      // Asegúrate de ajustar la fórmula según la celda donde quieras el resultado
  });
}

function crearNuevasHojas(ss,nombreHoja,cantidadProducto) {
  var nuevaHoja = ss.insertSheet(nombreHoja); // Crea una nueva hoja
  nuevaHoja.getRange('A1:B2').setValue(cantidadProducto); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('A1:B2').merge();// Fusiona las celdas después de establecer los valores
  var cuerpoCantidad = nuevaHoja.getRange('A1'); // Obtiene la celda en la columna C
  cuerpoCantidad.setNumberFormat('"CANT."\ 0');
  cuerpoCantidad.setBackground('#3c3c3c'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoCantidad.setFontColor('#fff'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoCantidad.setHorizontalAlignment('center'); // Alineación horizontal al centro
  cuerpoCantidad.setVerticalAlignment('middle'); // Alineación vertical en el medio
  
  nuevaHoja.getRange('C1:F2').setValue(nombreHoja); // Establece 'MUEBLE' en las celdas C1 a F1
  nuevaHoja.getRange('C1:F2').merge();
  var cuerpoTitulo = nuevaHoja.getRange('C1'); // Obtiene la celda en la columna C
  cuerpoTitulo.setBackground('#5aa5a5'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoTitulo.setFontColor('#fff'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoTitulo.setHorizontalAlignment('center'); // Alineación horizontal al centro
  cuerpoTitulo.setVerticalAlignment('middle'); // Alineación vertical en el medio

  nuevaHoja.getRange('G1:G2').setValue('0'); // Establece 'PN' en las celdas G1 y H1
  nuevaHoja.getRange('G1:G2').merge();
  var cuerpoPN = nuevaHoja.getRange('G1'); // Obtiene la celda en la columna C
  cuerpoPN.setNumberFormat('"PN"0');
  cuerpoPN.setBackground('#787878'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoPN.setFontColor('#fff'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoPN.setHorizontalAlignment('center'); // Alineación horizontal al centro
  cuerpoPN.setVerticalAlignment('middle'); // Alineación vertical en el medio

  nuevaHoja.getRange('A3:D3').setValue('CUERPO'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('A3:D3').merge();
  var cuerpoLabel = nuevaHoja.getRange('A3'); // Obtiene la celda en la columna C
  cuerpoLabel.setBackground('#787878'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoLabel.setFontColor('#fff'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoLabel.setHorizontalAlignment('center'); // Alineación horizontal al centro
  cuerpoLabel.setVerticalAlignment('middle'); // Alineación vertical en el medio

  nuevaHoja.getRange('E3:F3').setValue('0'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('E3:F3').merge();// Fusiona las celdas después de establecer los valores
  var cuerpoDivisiones = nuevaHoja.getRange('E3'); // Obtiene la celda en la columna C
  cuerpoDivisiones.setFormula('=SUM(E7:F8;E10:F11)'); // Referencia a la nueva hoja
  cuerpoDivisiones.setNumberFormat('0.00\ "m²"');
  cuerpoDivisiones.setBackground('#3c3c3c'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoDivisiones.setFontColor('#fff'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoDivisiones.setHorizontalAlignment('center'); // Alineación horizontal al centro
  cuerpoDivisiones.setVerticalAlignment('middle'); // Alineación vertical en el medio

  nuevaHoja.getRange('G3').setValue('0'); // Establece '0' en las celdas A1 y B1
  var cuerpoCubrecantos = nuevaHoja.getRange('G3'); // Obtiene la celda en la columna C
  cuerpoCubrecantos.setFormula('=SUM(G7:G11)'); // Referencia a la nueva hoja
  cuerpoCubrecantos.setNumberFormat('0.00\ "m"');
  cuerpoCubrecantos.setBackground('#3c3c3c'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoCubrecantos.setFontColor('#fff'); // Puedes utilizar códigos hexadecimales de colores o nombres de colores
  cuerpoCubrecantos.setHorizontalAlignment('center'); // Alineación horizontal al centro
  cuerpoCubrecantos.setVerticalAlignment('middle'); // Alineación vertical en el medio

  nuevaHoja.getRange('A4').setValue('Modulos'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('A4').setBackground('#b4d2d2');
  nuevaHoja.getRange('A4').setFontColor('#fff');
  nuevaHoja.getRange('B4').setValue('Ancho'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('B4').setBackground('#b4d2d2');
  nuevaHoja.getRange('B4').setFontColor('#fff');
  nuevaHoja.getRange('C4').setValue('Alturas'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('C4').setBackground('#b4d2d2');
  nuevaHoja.getRange('C4').setFontColor('#fff');
  nuevaHoja.getRange('D4').setValue('Profundidad'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('D4').setBackground('#b4d2d2');
  nuevaHoja.getRange('D4').setFontColor('#fff');
  nuevaHoja.getRange('E4:F4').setValue('Divisiones'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('E4:F4').merge();// Fusiona las celdas después de establecer los valores
  nuevaHoja.getRange('E4').setBackground('#b4d2d2');
  nuevaHoja.getRange('E4').setFontColor('#fff');
  nuevaHoja.getRange('G4').setValue('Entrepaños'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('G4').setBackground('#b4d2d2');
  nuevaHoja.getRange('G4').setFontColor('#fff');

  nuevaHoja.getRange('A5').setValue('0'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('B5').setValue('0'); // Establece '0' en las celdas A1 y B1
  var cuerpoAncho = nuevaHoja.getRange('B5'); // Obtiene la celda en la columna C
  cuerpoAncho.setNumberFormat('0\ "cm"');
  nuevaHoja.getRange('C5').setValue('0'); // Establece '0' en las celdas A1 y B1
  var cuerpoAltura = nuevaHoja.getRange('C5'); // Obtiene la celda en la columna C
  cuerpoAltura.setNumberFormat('0\ "cm"');
  nuevaHoja.getRange('D5').setValue('0'); // Establece '0' en las celdas A1 y B1
  var cuerpoProfundidad = nuevaHoja.getRange('D5'); // Obtiene la celda en la columna C
  cuerpoProfundidad.setNumberFormat('0\ "cm"');
  nuevaHoja.getRange('E5:F5').setValue('0'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('E5:F5').merge();// Fusiona las celdas después de establecer los valores
  nuevaHoja.getRange('G5').setValue('0'); // Establece '0' en las celdas A1 y B1

  nuevaHoja.getRange('A6').setValue('No. Pieza'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('A6').setBackground('#b4d2d2');
  nuevaHoja.getRange('A6').setFontColor('#fff');
  nuevaHoja.getRange('B6').setValue('Cantidad'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('B6').setBackground('#b4d2d2');
  nuevaHoja.getRange('B6').setFontColor('#fff');
  nuevaHoja.getRange('C6').setValue('Largo'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('C6').setBackground('#b4d2d2');
  nuevaHoja.getRange('C6').setFontColor('#fff');
  nuevaHoja.getRange('D6').setValue('Ancho'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('D6').setBackground('#b4d2d2');
  nuevaHoja.getRange('D6').setFontColor('#fff');
  nuevaHoja.getRange('E6:F6').setValue('Metros²'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('E6:F6').merge();// Fusiona las celdas después de establecer los valores
  nuevaHoja.getRange('E6').setBackground('#b4d2d2');
  nuevaHoja.getRange('E6').setFontColor('#fff');
  nuevaHoja.getRange('G6').setValue('Cubrecantos'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('G6').setBackground('#b4d2d2');
  nuevaHoja.getRange('G6').setFontColor('#fff');

  nuevaHoja.getRange('A7').setValue('HU-LR'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('A7').setFontColor('#aeabab');
  nuevaHoja.getRange('B7').setValue('0'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('B7').setFontColor('#aeabab');
  var hulrCantidad = nuevaHoja.getRange('B7'); // Obtiene la celda en la columna C
  hulrCantidad.setFormula('=(2*A5)*A1'); // Referencia a la nueva hoja
  nuevaHoja.getRange('C7').setValue('Largo'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('C7').setFontColor('#aeabab');
  var hulrLargo = nuevaHoja.getRange('C7'); // Obtiene la celda en la columna C
  hulrLargo.setFormula('=C5'); // Referencia a la nueva hoja
  nuevaHoja.getRange('D7').setValue('Ancho'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('D7').setFontColor('#aeabab');
  var hulrAncho = nuevaHoja.getRange('D7'); // Obtiene la celda en la columna C
  hulrAncho.setFormula('=D5'); // Referencia a la nueva hoja
  nuevaHoja.getRange('E7:F7').setValue('Metros²'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('E7:F7').merge();// Fusiona las celdas después de establecer los valores
  nuevaHoja.getRange('E7').setFontColor('#aeabab');
  var hulrDivisiones = nuevaHoja.getRange('E7'); // Obtiene la celda en la columna C
  hulrDivisiones.setFormula('=((C7*D7)*B7)/10000'); // Referencia a la nueva hoja
  nuevaHoja.getRange('G7').setValue('Cubrecantos'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('G7').setFontColor('#aeabab');
  var hulrCubrencantos = nuevaHoja.getRange('G7'); // Obtiene la celda en la columna C
  hulrCubrencantos.setFormula('=(C7*B7)/100'); // Referencia a la nueva hoja

  nuevaHoja.getRange('A8').setValue('PRECIO DE VENTA PIEZA	'); // Establece '0' en las celdas A1 y B1
  nuevaHoja.getRange('B8').setFormula('=E3*3/10'); // Referencia a la nueva hoja
  
} 
