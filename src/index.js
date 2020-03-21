const path = require('path');
const wordParser = require('word-text-parser');
const ExcelJS = require('exceljs');


/* La siguiente función se encarga de crear un documento de Excel
* a partir de una lista de datos recibida como parámetro
*
* @input dataList [Array]: Array con datos: { PU: <cadena>, Alta: <cadena>, Fecha: <cadena>}
*/
const createWorkBook = function(dataList) {
  if(!dataList || !Array.isArray(dataList) || !dataList.length) return;
  let wb = new ExcelJS.Workbook();
  wb.created = new Date();
  wb.creator = 'Someone';
  wb.views = [
    {
      x: 0, y: 0, width: 10000, height: 20000,
      firstSheet: 0, activeTab: 1, visibility: 'visible'
    }
  ]
  //creamos la hoja del libro
  let dataSheet = wb.addWorksheet('Data');

  //20 campos por registro, es lo que estableces en tu pregunta
  // los 3 datos van en las columnas 3, 5 y 7. con PU, alta y fecha respectivamente.

  // creamos la cabecera de datos
  let columnHeaders = [];
  for(let i = 0; i < 20; i++) {
    let header = {header: `Dato${i}`, key: `Dato${i}`, width: 30};
    if(i === 3) {
      header.header = 'PU';
      header.key = 'PU';
    }
    if(i === 5) {
      header.header = 'Alta';
      header.key = 'Alta';
    }
    if(i === 7) {
      header.header = 'Fecha';
      header.key = 'Fecha';
    }
    columnHeaders.push(header);
  }
  dataSheet.columns = columnHeaders;

  // Recorremos la lista y almacenamos los registros en la hoja
  dataList.forEach(data => {
    dataSheet.addRow({PU: data.PU, Alta: data.Alta, Fecha: data.Fecha});
  });

  // Escribimos el archivo resultante:
  wb.xlsx.writeFile(path.join(__dirname,'test.xlsx'))
  .then(() => {
    console.log('Done');
  })
  .catch(error => {
    console.error(`Algo salió mal: ${error.message}`);
  });
}

/*
* La siguiente función analiza los párrafos capturados del archivo de Word
* y devuelve una lista con los datos obtenidos de los mismos
* @input resultsList [Array]: Array con cadenas correspondientes a cada párrafo del archivo
* @output data [Array]: Array con datos: { PU: <cadena>, Alta: <cadena>, Fecha: <cadena>}
*/
const parseResult = function(resultsList) {
  if(!resultsList || !Array.isArray(resultsList) || !resultsList.length) return;

  // filtramos los párrafos que contengan "PU." y "Alta"
  const filteredParagraphs = resultsList.filter(p => p.includes('PU.') && p.includes('Alta'));

  // Array que almacenará los datos extraídos
  const data = [];

  // Expresión regular para evaluar la fecha.
  const regexFecha = /^\d\d\/\d\d$/i

  // recorremos cada párrafo
  filteredParagraphs.forEach(p => {
    //creamos el objeto para almacenar los datos
    let value = {
      PU: '',
      Alta: '',
      Fecha: ''
    }

    // separamos cada párrafo en palabras
    let splitedParagraph = p.split(' ');

    //analizamos cada palabra y capturamos los datos
    splitedParagraph.forEach((word, index) => {
      if(word.includes('PU.')) {
        value['PU'] = word.split('.')[1];
      }
      if(word === 'Alta') {
        value['Alta'] = splitedParagraph[index + 1];
      }
      if(regexFecha.test(word)) {
        value['Fecha'] = `${word}/2020`;
      }
    });

    // almacenamos el resultado en el array
    data.push(value);
  });

  // mostramos los datos por consola
  console.log(data);

  // llamamos al procedimiento para crear el libro de Excel a partir de los datos
  createWorkBook(data);
}

// nombre del archivo a analizar
const fileToParse = path.join(__dirname, 'demo.docx');

// llamamos al analizador del archivo
wordParser(fileToParse, parseResult);
