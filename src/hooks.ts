import { read, writeXLSX } from "xlsx";
import Excel from "exceljs";
export const useExcelJS = () => {
  return {
    async getPendingEditedXls(file: File) {
      const wb = await loadFile(file);
      const sheets = getSheets(wb);
      const title = getRow(sheets[0], 1);
      console.log(title);
      const colsKey = getColumn(sheets[0], "K");
      // console.log(colsKey);
      const colsValue = getColumn(sheets[0], "B");
      const config = getMatchConfig(colsKey, colsValue);
      // console.log(config);
      const cell = sheets[0].getCell("K8");
      // console.log(cell.value);
      cell.value = 8;
      // console.log(getColumn(sheets[0], "K"));
      fillValues(sheets[0], config, colsKey, "D");
      setTimeout(() => {
        saveFile(wb, file.type);
      }, 1000);
    },
  };
};

export async function loadFile(file: File) {
  const blob = await file.arrayBuffer();
  const wb = await read(blob);

  const buffer = await writeXLSX(wb, { type: "buffer" });
  const workbook = new Excel.Workbook();
  const current = await workbook.xlsx.load(buffer);
  return current;
}

export async function saveFile(workbook: Excel.Workbook, type: string) {
  const buffer = await workbook.xlsx.writeBuffer();
  window.open(URL.createObjectURL(new Blob([buffer], { type })));
}

export function getSheets(workbook: Excel.Workbook) {
  return workbook.worksheets;
}

export function getRow(sheet: Excel.Worksheet, number: number) {
  const result: { label: string; value: string }[] = [];
  sheet.getRow(number).eachCell((it) => {
    result.push({
      label: `${it.address} ${it.value}`,
      value: it.address.replace(/[0-9]/g, ""),
    });
  });
  return result;
}

export function getColumn(sheet: Excel.Worksheet, col: string | number): [] {
  return (sheet.getColumn(col).values as []) || [];
}

export type ColsType = ReturnType<typeof getColumn>;
export type ConfigType = { [k: string]: string };
export function getMatchConfig(colsKey: ColsType, colsValue: ColsType) {
  const result: ConfigType = {};
  for (let i = 0; i < colsKey.length; i++) {
    const key = colsKey[i] || "";
    if (key) {
      result[key] = colsValue[i];
    }
  }
  return result;
}

export function fillValues(
  sheet: Excel.Worksheet,
  config: ReturnType<typeof getMatchConfig>,
  colsKey: ColsType,
  fillCol: string
) {
  const len = colsKey.length;
  for (let i = 1; i <= len; i++) {
    const key = colsKey[i];
    sheet.getCell(fillCol + i).value = config[key];
  }
}
