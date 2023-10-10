import { read, writeXLSX } from "xlsx";
import Excel from "exceljs";

export async function loadFile(file: File) {
  const blob = await file.arrayBuffer();
  const wb = await read(blob);

  const buffer = await writeXLSX(wb, { type: "buffer" });
  const workbook = new Excel.Workbook();
  const current = await workbook.xlsx.load(buffer);
  return current;
}

export async function saveFile(
  workbook: Excel.Workbook,
  type: string,
  fileName: string
) {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type });
  const link = window.URL.createObjectURL(blob);
  let a: HTMLAnchorElement | null = document.createElement("a");
  a.style.setProperty("display", "none");
  a.href = link;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  window.URL.revokeObjectURL(link);
  a.remove();
  a = null;
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
