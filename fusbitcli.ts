import { Sheet } from "xlsx";

const XLSX = require("xlsx");
const XlsxPopulate = require("xlsx-populate");
const config = require("./config.json");
const readline = require("readline");

let STARTING_EXCEL_ROW = 3;
const year = new Date().getFullYear();
const feriadosArgentinos: Feriado[] = [];

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

export default function ask<T>(question: string): Promise<T> {
  return new Promise((resolve) => {
    rl.question(question, (answer: any) => {
      resolve(answer as unknown as T);
    });
  });
}

(async function populateFeriados() {
  const resp = await fetch("https://api.argentinadatos.com/v1/feriados");
  await resp
    .json()
    .then((resp) => {
      for (const date of resp) {
        feriadosArgentinos.push(date);
      }
    })
    .catch((e) => {
      console.log(e);
    });
})();

async function createExcel(responsablesCount: number, responsables: Array<string>, project: string, leader: string) {
  const wb = XLSX.utils.book_new();

  // O(n^3)
  // TODO optimizar esto para que ande en una tostadora
  for (let month = 1; month <= 12; month++) {
    const excelCols = [
      ["AÑO/MES", "", "", "TOTAL DE HORAS", "", "", ""],
      ["Fecha", "Responsable", "Proyecto", "Incidente/Tarea", "Horas", "Descripcion"],
    ];

    if (leader) {
      STARTING_EXCEL_ROW = 4;
      excelCols.push(["", leader, "", "", "", ""]);
    }

    // aoa es array of arrays
    let excelSheet = XLSX.utils.aoa_to_sheet([]);

    let j = 1;
    for (let i = 0; i < responsablesCount; i++) {
      const daysInMonth = new Date(year, month, 0).getDate();
      for (; j <= daysInMonth; j++) {
        const currentDate = new Date(year, month - 1, j);

        const rowData = [currentDate.toLocaleDateString("en-EN"), responsables[i], "", "", ""];
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
      workbook.sheets().map((sh: Sheet) => {
        for (let row = STARTING_EXCEL_ROW; ; row++) {
          const cell = sh.cell(`A${row}`);
          sh.row(row).height(15);
          sh.row(row).style("fontFamily", config.font.defaultFont);
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

          if (feriadosArgentinos.length !== 0) {
            const formattedCurrentDateForFeriadoAPI = currentDate.toISOString().split("T")[0];

            for (const feriado of feriadosArgentinos) {
              if (formattedCurrentDateForFeriadoAPI === feriado.fecha) {
                const r = sh.range(`A${row}:F${row}`);
                r.style("fill", config.colors.feriadosColor);
              }
            }
          }
        }
        const lastRow = sh.usedRange().endCell().rowNumber();

        const mainColumnsColor = config.colors.mainColumnsColor;

        applyStylesAndFormulasToRows(sh, lastRow, mainColumnsColor);
      });
      return workbook.toFileAsync(`./${excelFileName}`);
    })
    .then(() => {
      rl.close();
      console.log("Bitacora lista 👍");
    })
    .catch((error: any) => {
      console.error("Error:", error);
      rl.close();
    });
}

function applyStylesAndFormulasToRows(sh: any, lastRow: any, mainColumnsColor: any) {
  sh.row(1).style("fontFamily", config.font.defaultFont);
  sh.row(1).style("bold", true);
  sh.row(2).style("fontFamily", config.font.defaultFont);
  sh.row(2).style("bold", true);
  sh.row(1).height(20);

  sh.row(3).style("fontFamily", config.font.defaultFont);
  sh.row(3).style("fontFamily", config.font.defaultFont);
  sh.row(3).height(15);

  sh.cell("A2").style("fill", mainColumnsColor);
  sh.column("A").style("horizontalAlignment", "center");
  sh.column("A").style("border", true);
  sh.column("A").width(15);

  sh.cell(`B1`).value(`${year + " " + sh.name()}`);
  sh.cell("B1").style("horizontalAlignment", "center");
  sh.cell("B2").style("fill", mainColumnsColor);
  sh.column("B").style("border", true);
  sh.column(`B`).width(25);

  sh.cell("C2").style("fill", mainColumnsColor);
  sh.column("C").style("border", true);
  sh.column("C").width(15);

  sh.cell("D2").style("fill", mainColumnsColor);
  sh.column("D").style("border", true);
  sh.column("D").width(40);

  sh.cell(`E1`).formula(`SUM(E3:E${lastRow})`);
  sh.cell("E1").style("horizontalAlignment", "center");
  sh.cell("E2").style("fill", mainColumnsColor);
  sh.column("E").style("border", true);
  sh.column("E").width(15);

  sh.cell("F2").style("fill", mainColumnsColor);
  sh.column("F").style("border", true);
  sh.column("F").width(15);
}

async function main() {
  let responsablesAmount: number = await ask<number>(
    "Ingrese numero de responsables (no contar a quien da seguimiento al proyecto): "
  );

  while (!responsablesAmount || isNaN(responsablesAmount)) {
    if (isNaN(responsablesAmount)) {
      console.log("No se acepta texto como parametro");
    }
    responsablesAmount = await ask<number>(
      "Ingrese numero de responsables (no contar a quien da seguimiento al proyecto): "
    );
  }

  let responsables: Array<string> = [];
  for (let i = 0; i < responsablesAmount; i++) {
    const name: string = await ask<string>(`Ingrese nombre de responsable numero ${i + 1}: `);
    responsables.push(name);
  }

  const leader: string = await ask<string>(
    `Ingrese nombre de quien da seguimiento al proyecto (dejar vacio si no lo hay): `
  );

  const project: string = await ask<string>("Ingrese nombre del proyecto: ");

  await createExcel(responsablesAmount, responsables, project, leader);
}

main();
