const readline = require("readline");
const XLSX = require("xlsx");
const XlsxPopulate = require("xlsx-populate");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

function ask<T>(question: string): Promise<T> {
  return new Promise((resolve) => {
    rl.question(question, (answer: any) => {
      resolve(answer as unknown as T);
    });
  });
}

const year = new Date().getFullYear();

async function createExcel(
  responsablesCount: number,
  responsables: Array<string>,
  project: string
) {
  const wb = XLSX.utils.book_new();

  for (let month = 1; month <= 12; month++) {
    const excelCols = [
      [
        "Fecha",
        "Responsable",
        "Proyecto",
        "Incidente/Tarea",
        "Horas",
        "Descripcion",
      ],
    ];
    let excelSheet = XLSX.utils.aoa_to_sheet([]);
    let j = 1;
    for (let i = 0; i < responsablesCount; i++) {
      const daysInMonth = new Date(year, month, 0).getDate();
      for (; j <= daysInMonth; j++) {
        const currentDate = new Date(year, month - 1, j);

        const rowData = [
          currentDate.toLocaleDateString("en-EN"),
          responsables[i],
          "",
          "",
          "",
        ];
        excelCols.push(rowData);
      }
      j = 1; // Reseteamos los dias por cada participante para que, despues del ultimo dia, se itere de nuevo
    }
    excelSheet = XLSX.utils.aoa_to_sheet(excelCols);

    const date = new Date(year, month - 1);
    const monthNameInSpanish = date
      .toLocaleString("es-ES", { month: "long" })
      .toUpperCase();

    const sheetName = `${monthNameInSpanish}`;
    XLSX.utils.book_append_sheet(wb, excelSheet, sheetName);
  }

  const excelName = `${project} ${year} Bitacora`;
  const excelFileName = `${excelName}.xlsx`;
  XLSX.writeFile(wb, excelFileName);

  console.log(`Se genero el archivo: "${excelFileName}" satisfactoriamente.`);

  XlsxPopulate.fromFileAsync(`./${excelFileName}`)
    .then((workbook: any) => {
      workbook.sheets().map((sh: any) => {
        for (let row = 2; ; row++) {
          const cell = sh.cell(`A${row}`);
          sh.row(row).height(15);
          let dateValue = cell.value();
          if (!dateValue) break;

          // Esto para que este en formato dia/mes/anio en el excel
          const currentDate = new Date(dateValue);
          dateValue = cell.value(currentDate.toLocaleDateString("es-ES"));

          const isWeekend =
            currentDate.getDay() === 0 || currentDate.getDay() === 6;

          if (isWeekend) {
            const r = sh.range(`A${row}:F${row}`);
            r.style("fill", "4a74e8");
          }
        }
        const lastRow = sh.usedRange().endCell().rowNumber();


        const commonHexForCols = "8dde87";

        // Total de horas
        sh.cell(`G1`).value("Total de horas:");
        sh.column(`G`).width(15);
        sh.cell("G1").style("horizontalAlignment", "center");
        sh.cell("G1").style("fill", "d3db74");


        // Total de horas formula
        sh.cell(`G2`).formula(`SUM(E2:E${lastRow})`);
        sh.cell("G2").style("horizontalAlignment", "center");

        // Fecha
        sh.cell("A1").style("bold", true);
        sh.cell("A1").style("fill", commonHexForCols);
        sh.column("A").width(15);
        sh.column("A").style("horizontalAlignment", "center");
        sh.column("A").style("border", true);

        // Responsable
        sh.cell("B1").style("bold", true);
        sh.cell("B1").style("fill", commonHexForCols);
        sh.column("B").width(15);
        sh.column("B").style("border", true);

        // Proyecto
        sh.cell("C1").style("bold", true);
        sh.cell("C1").style("fill", commonHexForCols);
        sh.column("C").width(15);
        sh.column("C").style("border", true);

        // Incidente/Tarea
        sh.cell("D1").style("bold", true);
        sh.cell("D1").style("fill", commonHexForCols);
        sh.column("D").width(100);
        sh.column("D").style("border", true);

        // Horas
        sh.cell("E1").style("bold", true);
        sh.cell("E1").style("fill", commonHexForCols);
        sh.column("E").width(15);
        sh.column("E").style("border", true);

        // Descripcion
        sh.cell("F1").style("bold", true);
        sh.cell("F1").style("fill", commonHexForCols);
        sh.column("F").width(15);
        sh.column("F").style("border", true);
      });
      return workbook.toFileAsync(`./${excelFileName}`);
    })
    .then(() => {
      rl.close();
    })
    .catch((error: any) => {
      console.error("Error:", error);
      rl.close();
    });
}

async function main() {
  const responsablesCount: number = await ask<number>(
    "Numero de responsables: "
  );

  const responsables: Array<string> = [];
  for (let i = 0; i < responsablesCount; i++) {
    const name: string = await ask<string>(
      `Nombre de responsable numero ${i + 1}: `
    );
    responsables.push(name);
  }

  const project: string = await ask<string>("Proyecto: ");

  await createExcel(responsablesCount, responsables, project);
}

main();
