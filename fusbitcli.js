const readline = require('readline');
const XLSX = require('xlsx');
const XlsxPopulate = require('xlsx-populate');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function questionAsync(question) {
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      resolve(answer);
    });
  });
}

const month = new Date().getMonth() + 1;
const year = new Date().getFullYear();

async function createExcel(responsablesCount, responsables, project) {
  const wb = XLSX.utils.book_new();

  const excelCols = [
    ['Fecha', 'Responsable', 'Incidente/Tarea', 'Horas', 'Descripcion']
  ];
  let excelSheet = XLSX.utils.aoa_to_sheet([]);

  let j = 1;
  const weekendCellLocation = [];

  for (let i = 0; i < responsablesCount; i++) {
    const daysInMonth = new Date(year, month, 0).getDate();
    for (; j <= daysInMonth; j++) {
      const currentDate = new Date(year, month - 1, j);
      const formattedDate = currentDate.toLocaleDateString('es-ES');

      const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;
      if (isWeekend) {
        const cellLocation = `${"A"}${j + 1}`;
        weekendCellLocation.push(cellLocation);
      }

      const rowData = [formattedDate, responsables[i], '', '', ''];
      excelCols.push(rowData);

    }
    j = 1;
  }

  excelSheet = XLSX.utils.aoa_to_sheet(excelCols);
  const excelName = `${project} ${year} Bitacora`;

  const date = new Date(year, month - 1);
  const monthNameInSpanish = date.toLocaleString('es-ES', { month: 'long' }).toUpperCase();
  const sheetName = `${monthNameInSpanish}`;
    XLSX.utils.book_append_sheet(wb, excelSheet, sheetName);
  const excelFileName = `${excelName}.xlsx`;
  XLSX.writeFile(wb, excelFileName);
  console.log(`Se genero el archivo: "${excelFileName}" satisfactoriamente.`);

  XlsxPopulate.fromFileAsync(`./${excelFileName}`)
    .then(workbook => {
      const sheet = workbook.sheet(sheetName);
      weekendCellLocation.forEach(cellLocation => {
        const cell = sheet.cell(cellLocation);
        cell.style({ fill: 'lightblue' });
      });
      return workbook.toFileAsync(`./${excelFileName}`);
    })
    .then(() => {
      rl.close();
    })
    .catch(error => {
      console.error('Error:', error);
      rl.close();
    });
}

async function main() {

  const responsablesCount = await questionAsync('Numero de participantes: ');

  const responsables = [];
  for (let i = 0; i < responsablesCount; i++) {
    const name = await questionAsync(`Nombre de participante numero ${i + 1}: `);
    responsables.push(name);
  }

  const project = await questionAsync('Proyecto: ');


  await createExcel(parseInt(responsablesCount), responsables, project);

}

main();
