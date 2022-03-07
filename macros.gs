function onOpen() {
  var nombre=Session.getEffectiveUser().getEmail().split(".")[0]
  var primeraLetra=nombre.charAt(0).toUpperCase()
  var nombreCompleto=primeraLetra+nombre.slice(1)
  var bienvenido="¡Bienvenido/a "
  SpreadsheetApp.getUi().alert(bienvenido.concat(nombreCompleto,"!"))  
}

function actualizarBD(){
  var html = HtmlService.createHtmlOutputFromFile('index').setWidth(200).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html,"Ejecutando programa, por favor espere");
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SS.getSheetByName('BD melis');
  var tiempoTranscurrido = Date.now()+(1000*60*60)*2
  var fecha = new Date(tiempoTranscurrido)
  var dia = fecha.getDate()
  
  var listaDeNumeros = [1,2,3,4,5,6,7,8,9]
  if (listaDeNumeros.includes(dia)){
    dia = "0" + dia
  }
  var mes = fecha.getMonth()+1
  if (listaDeNumeros.includes(mes)){
    mes = "0" + mes
  }
  var agno = fecha.getFullYear()
  var fechaOk = dia + "-" + mes + "-" + agno
  var nombreDelArchivo = fechaOk + ".csv"
  var nombreDelArchivo2 = "03-03-2022.csv"

  var files = DriveApp.searchFiles('title = "' + nombreDelArchivo2 + '"');
  var file = files.next();

  var csv = file.getBlob().getDataAsString('ISO-8859-1');
  csv = csv.replace(/[^a-zA-Z0-9\-\.\,\;\nÃÀÁÄÂÈÉËÊÌÍÏÎÒÓÖÔÙÚÜÛãàáäâèéëêìíïîòóöôùúüûÑñ ]/g, "");
  var values = Utilities.parseCsv(csv, ";");
  var parte1 = values.slice(1,50001) // dejo desde el 1 porque en el indice 0 se muestra inventory ID y valor de seguro
  var parte2 = values.slice(50001,100001) 
  var parte3 = values.slice(100001,150001) 
  var parte4 = values.slice(150001,200001) 
  var parte5 = values.slice(200001,250001) 
  var parte6 = values.slice(250001,300001) 
  var parte7 = values.slice(300001,350001) 
  var parte8 = values.slice(350001,400001) 
  var parte9 = values.slice(400001,450001) 
  var parte10 = values.slice(450001,500001) 
  var parte11 = values.slice(500001,550001) 
  var parte12 = values.slice(550001,600001) 
  var parte13 = values.slice(600001,650001) 
  var parte14 = values.slice(650001,700001) 
  var parte15 = values.slice(700001,750001) 
  var parte16 = values.slice(750001,800001) 
  var parte17 = values.slice(800001,850001) 
  var parte18 = values.slice(850001,900001) 
  var parte19 = values.slice(900001,950001) 
  var parte20 = values.slice(950001,1000001)
  var ultimoItemDelArray = values[values.length - 1]
  var parte21 = values.slice(1000001,values.indexOf(ultimoItemDelArray))
  parte21.push(ultimoItemDelArray)
  var rango1 = sheet.getRange('A13:B50012')
  var rango2 = sheet.getRange('A50013:B100012')
  var rango3 = sheet.getRange('A100013:B150012')
  var rango4 = sheet.getRange('A150013:B200012')
  var rango5 = sheet.getRange('A200013:B250012')
  var rango6 = sheet.getRange('A250013:B300012')
  var rango7 = sheet.getRange('A300013:B350012')
  var rango8 = sheet.getRange('A350013:B400012')
  var rango9 = sheet.getRange('A400013:B450012')
  var rango10 = sheet.getRange('A450013:B500012')
  var rango11 = sheet.getRange('A500013:B550012')
  var rango12 = sheet.getRange('A550013:B600012')
  var rango13 = sheet.getRange('A600013:B650012')
  var rango14 = sheet.getRange('A650013:B700012')
  var rango15 = sheet.getRange('A700013:B750012')
  var rango16 = sheet.getRange('A750013:B800012')
  var rango17 = sheet.getRange('A800013:B850012')
  var rango18 = sheet.getRange('A850013:B900012')
  var rango19 = sheet.getRange('A900013:B950012')
  var rango20 = sheet.getRange('A950013:B1000012')
  var contenidoParseado = []  

    // funcion para parsear todo
  function parseador(contenido){
    for (i in contenido){
      let tmp = contenido[i][0].split(",")
      contenidoParseado.push(tmp)
    }
    for (i in contenidoParseado){
      let elPunto = contenidoParseado[i][1].indexOf(".")
      if (elPunto != -1){
      contenidoParseado[i][1] = contenidoParseado[i][1].slice(0,elPunto)
      } 
    }
  }


// Bloque en el que aseguro que estén la cantidad de filas que corresponde

var qvalores = values.length + 11 // 11 
var qfilas = sheet.getMaxRows() // es -12 por las 12 primeras filas de la sheet
var ultimaFila = sheet.getLastRow()
var ultimoRango = "A1000013:B" + qvalores
var rango21 = sheet.getRange(ultimoRango)

if (qvalores > qfilas){
  sheet.insertRowsAfter(ultimaFila,qvalores - qfilas) // -1 porque no cuenta la 
}

if (qvalores < qfilas){
  let diferenciaDeFilas = qfilas - qvalores
  let primerFilaAEliminar = ultimaFila - diferenciaDeFilas
  sheet.deleteRows(primerFilaAEliminar,qfilas - qvalores)
}

  // Proceso de pegado de información por partes
  
  parseador(parte1)
  rango1.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []
  
  parseador(parte2)
  rango2.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []  

  parseador(parte3)
  rango3.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte4)
  rango4.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte5)
  rango5.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte6)
  rango6.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte7)
  rango7.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte8)
  rango8.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte9)
  rango9.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte10)
  rango10.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte11)
  rango11.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte12)
  rango12.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte13)
  rango13.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte14)
  rango14.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte15)
  rango15.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte16)
  rango16.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte17)
  rango17.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte18)
  rango18.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte19)
  rango19.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte20)
  rango20.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []

  parseador(parte21)
  rango21.setValues(contenidoParseado)
  Utilities.sleep(2000)
  contenidoParseado = []


  var fechaConBarras = dia + "/" + mes + "/" + agno
  sheet.getRange('A11').setValue(fechaConBarras)
  SpreadsheetApp.getUi().alert("Base de datos actualizada")
  sheet.hideSheet()
}
