const readline = require("readline");
const XLSX = require("xlsx");
const XlsxPopulate = require("xlsx-populate");
const config = require('./config.json')

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
        "Año/Mes"
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
          "",
        ];
        excelCols.push(rowData);
      }
      j = 1; // Reseteamos los dias por cada participante para que, despues del ultimo dia, se itere de nuevo
    }
    excelSheet = XLSX.utils.aoa_to_sheet(excelCols);

    const date = new Date(year, month - 1);
    const monthNameInSpanish = date
      .toLocaleString(config.date.es, { month: config.date.lengthOfMonthName })
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
      console.log("Populando las columnas del excel con fechas y nombres...")
      workbook.sheets().map((sh: any) => {
        for (let row = 2; ; row++) {
          const cell = sh.cell(`A${row}`);
          sh.row(row).height(15);
          sh.row(row).style("fontFamily", config.font.defaultFont)
          let dateValue = cell.value();
          // La condicion para cortar el loop es si ya no hay fecha que completar por integrante.
          if (!dateValue) break;

          const currentDate = new Date(dateValue);
          // Esto para que este en formato dia/mes/anio en el excel
          dateValue = cell.value(currentDate.toLocaleDateString(config.date.es));

          const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;

          if (isWeekend) {
            const r = sh.range(`A${row}:F${row}`);
            r.style("fill", config.colors.weekendColor);
          }
        }
        const lastRow = sh.usedRange().endCell().rowNumber();

        const mainColumnsColor = config.colors.mainColumnsColor;

        sh.row(1).style("fontFamily", config.font.defaultFont)

        // Año/Mes
        sh.cell(`H1`).value("Año/Mes");
        sh.column(`H`).width(15);
        sh.cell("H1").style("horizontalAlignment", "center");
        sh.cell("H1").style("bold", true);
        sh.cell("H1").style("fill", config.colors.anioMesColumnColor);

        sh.cell(`H2`).value(`${year + " " + sh.name()}`);
        sh.column(`H`).width(15);
        sh.cell("H2").style("horizontalAlignment", "center");
        sh.cell("H2").style("bold", true);

        // Total de horas
        sh.cell(`G1`).value("Total de horas:");
        sh.column(`G`).width(15);
        sh.cell("G1").style("horizontalAlignment", "center");
        sh.cell("G1").style("fill", config.colors.totalHorasColumnColor);

        // Total de horas formula
        sh.cell(`G2`).formula(`SUM(E2:E${lastRow})`);
        sh.cell("G2").style("horizontalAlignment", "center");

        // Fecha
        sh.cell("A1").style("bold", true);
        sh.cell("A1").style("fill", mainColumnsColor);
        sh.column("A").width(15);
        sh.column("A").style("horizontalAlignment", "center");
        sh.column("A").style("border", true);

        // Responsable
        sh.cell("B1").style("bold", true);
        sh.cell("B1").style("fill", mainColumnsColor);
        sh.column("B").width(15);
        sh.column("B").style("border", true);

        // Proyecto
        sh.cell("C1").style("bold", true);
        sh.cell("C1").style("fill", mainColumnsColor);
        sh.column("C").width(15);
        sh.column("C").style("border", true);

        // Incidente/Tarea
        sh.cell("D1").style("bold", true);
        sh.cell("D1").style("fill", mainColumnsColor);
        sh.column("D").width(40);
        sh.column("D").style("border", true);

        // Horas
        sh.cell("E1").style("bold", true);
        sh.cell("E1").style("fill", mainColumnsColor);
        sh.column("E").width(15);
        sh.column("E").style("border", true);

        // Descripcion
        sh.cell("F1").style("bold", true);
        sh.cell("F1").style("fill", mainColumnsColor);
        sh.column("F").width(15);
        sh.column("F").style("border", true);
      });
      return workbook.toFileAsync(`./${excelFileName}`);
    })
    .then(() => {
      rl.close();
      console.log("Bitacora lista 👍")
    })
    .catch((error: any) => {
      console.error("Error:", error);
      rl.close();
    });
}

async function main() {
  let responsablesCount: number = await ask<number>(
    "Numero de responsables: "
  );

  while(!responsablesCount || isNaN(responsablesCount)) {
    if(isNaN(responsablesCount)) {
      console.log("No se acepta texto como parametro")
    }
    responsablesCount = await ask<number>(
      "Numero de responsables: "
    );
  }

  let responsables: Array<string> = [];
  for (let i = 0; i < responsablesCount; i++) {
    const name: string = await ask<string>(
      `Nombre de responsable numero ${i + 1}: `
    );
    responsables.push(name);
  }




  if(responsables.length == 0) {
    for (let i = 0; i < responsablesCount; i++) {
      const name: string = await ask<string>(
        `Nombre de responsable numero ${i + 1}: `
      );
      responsables.push(name);
    }
  }


  const project: string = await ask<string>("Proyecto: ");

  await createExcel(responsablesCount, responsables, project);
}

main();
