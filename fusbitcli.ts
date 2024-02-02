const readline = require("readline");
const XLSX = require("xlsx");
const XlsxPopulate = require("xlsx-populate");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

function questionAsync<T>(question: string): Promise<T> {
  return new Promise((resolve) => {
    rl.question(question, (answer:any) => {
      resolve(answer as unknown as T);
    });
  });
}

let month = new Date().getMonth();
const year = new Date().getFullYear();

async function createExcel(responsablesCount: number, responsables: Array<string>, project: string) {
  const wb = XLSX.utils.book_new();

  const excelCols = [
    ["Fecha", "Responsable", "Incidente/Tarea", "Horas", "Descripcion"],
  ];
  let excelSheet = XLSX.utils.aoa_to_sheet([]);

  let j = 1;
  for (let i = 0; i < responsablesCount; i++) {
    const daysInMonth = new Date(year, month, 0).getDate();
    for (; j <= daysInMonth; j++) {
      const currentDate = new Date(year, month - 1, j);

      const rowData = [currentDate.toLocaleDateString("en-EN"), responsables[i], "", "", ""];
      excelCols.push(rowData);
    }
    j = 1;
  }

  excelSheet = XLSX.utils.aoa_to_sheet(excelCols);
  const excelName = `${project} ${year} Bitacora`;

  const date = new Date(year, month - 1);
  const monthNameInSpanish = date
    .toLocaleString("es-ES", { month: "long" })
    .toUpperCase();
  const sheetName = `${monthNameInSpanish}`;
  XLSX.utils.book_append_sheet(wb, excelSheet, sheetName);
  const excelFileName = `${excelName}.xlsx`;
  XLSX.writeFile(wb, excelFileName);
  console.log(`Se genero el archivo: "${excelFileName}" satisfactoriamente.`);

  XlsxPopulate.fromFileAsync(`./${excelFileName}`)
    .then((workbook:any) => {
      const sheet = workbook.sheet(sheetName);
      for (let row = 2; ; row++) {
        const cell = sheet.cell(`A${row}`);
        const dateValue = cell.value();
        if (!dateValue) break;

        const currentDate = new Date(dateValue);
        const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;

        if (isWeekend) {
          const weekendCell = sheet.cell(`A${row}`);
          weekendCell.style({ fill: "6840d6" });
        }
      }
      return workbook.toFileAsync(`./${excelFileName}`);
    })
    .then(() => {
      rl.close();
    })
    .catch((error:any) => {
      console.error("Error:", error);
      rl.close();
    });
}

async function main() {
  const responsablesCount: number = await questionAsync<number>('Numero de participantes: ');

  const responsables: Array<string> = [];
  for (let i = 0; i < responsablesCount; i++) {
    const name: string = await questionAsync<string>(`Nombre de participante numero ${i + 1}: `);
    responsables.push(name);
  }

  const project:string = await questionAsync<string>('Proyecto: ');

  await createExcel(responsablesCount, responsables, project);
}

main();
