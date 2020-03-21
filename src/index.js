const path = require('path');
const wordParser = require('word-text-parser');
const ExcelJS = require('exceljs');

const createWorkBook = function(dataList) {
  let wb = new ExcelJS.Workbook();
  wb.created = new Date();
  wb.creator = 'Someone';
  wb.views = [
    {
      x: 0, y: 0, width: 10000, height: 20000,
      firstSheet: 0, activeTab: 1, visibility: 'visible'
    }
  ]
  let dataSheet = wb.addWorksheet('Data');
  //20 campos por registro, es lo que estableces en tu pregunta
  // los 3 datos van las columnas 3, 5 y 7. con PU, alta y fecha como orden.
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
  dataList.forEach(data => {
    dataSheet.addRow({PU: data.PU, Alta: data.Alta, Fecha: data.Fecha});
  });
  // verificar el nombre de salida del archivo
  wb.xlsx.writeFile(path.join(__dirname,'test.xlsx'))
  .then(() => {
    console.log('Done');
  })
}

const parseResult = function(resultsList) {
  const filteredParagraphs = resultsList.filter(p => p.includes('PU.') && p.includes('Alta'));
  const data = [];
  const regexFecha = /^\d\d\/\d\d$/i
  filteredParagraphs.forEach(p => {
    let value = {
      PU: '',
      Alta: '',
      Fecha: ''
    }
    let splitedParagraph = p.split(' ');
    splitedParagraph.forEach((word, index) => {
      if(word.includes('PU.')) {
        value['PU'] = word.split('.')[1];
      }
      if(word.includes('Alta')) {
        value['Alta'] = splitedParagraph[index + 1];
      }
      if(regexFecha.test(word)) {
        value['Fecha'] = `${word}/2020`;
      }
    });
    data.push(value);
  });
  console.log(data);
  createWorkBook(data);
}

// nombre del archivo de entrada
const fileToParse = path.join(__dirname, 'demo.docx');

wordParser(fileToParse, parseResult);
