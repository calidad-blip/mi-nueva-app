import * as XLSX from "xlsx";

const MAX_SHEETS = 20;
const MAX_ROWS = 50_000;
const MAX_COLUMNS = 100;
const MAX_CELL_LENGTH = 10_000;

export type PreasignacionRow = {
  cantidadSolicitada: number | null;
  cantidadPreasignada: number | null;
  preasignado: boolean;
  clienteNumero: string;
  clienteNombre: string;
  codigoProducto: string;
  nombreProducto: string;
};

export function parsePreassignments(buffer: ArrayBuffer): PreasignacionRow[] {
  let workbook: XLSX.WorkBook;

  try {
    workbook = XLSX.read(buffer, { type: "array" });
  } catch {
    throw new Error("El archivo XLSX está dañado o no se puede interpretar.");
  }

  if (workbook.SheetNames.length > MAX_SHEETS) {
    throw new Error(`El archivo supera el límite de ${MAX_SHEETS} hojas.`);
  }

  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  if (!sheet) {
    return [];
  }

  const sheetReference = sheet["!ref"];

  if (!sheetReference) {
    return [];
  }

  let range: XLSX.Range;

  try {
    range = XLSX.utils.decode_range(sheetReference);
  } catch {
    throw new Error("La hoja principal contiene un rango inválido.");
  }

  const rowCount = range.e.r - range.s.r + 1;
  const columnCount = range.e.c - range.s.c + 1;

  if (rowCount > MAX_ROWS) {
    throw new Error(`La hoja supera el límite de ${MAX_ROWS.toLocaleString("es-AR")} filas.`);
  }

  if (columnCount > MAX_COLUMNS) {
    throw new Error(`La hoja supera el límite de ${MAX_COLUMNS} columnas.`);
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
  }) as Array<Array<string | number | boolean | null | undefined>>;

  if (rows.length === 0) {
    return [];
  }

  if (
    rows.some((row) =>
      row.some((cell) => String(cell ?? "").length > MAX_CELL_LENGTH),
    )
  ) {
    throw new Error("El archivo contiene una celda excesivamente extensa.");
  }

  return rows
    .slice(1)
    .filter((row) => row.some((cell) => String(cell ?? "").trim() !== ""))
    .map((row) => ({
      cantidadSolicitada: normalizeNumber(row[1]),
      cantidadPreasignada: normalizeNumber(row[5]),
      preasignado: normalizeBoolean(row[6]),
      clienteNumero: normalizeText(row[10]),
      clienteNombre: normalizeText(row[11]),
      codigoProducto: normalizeText(row[13]),
      nombreProducto: normalizeText(row[14]),
    }))
    .filter(
      (row) =>
        row.clienteNumero ||
        row.clienteNombre ||
        row.codigoProducto ||
        row.nombreProducto ||
        row.cantidadPreasignada !== null,
    );
}

function normalizeNumber(value: string | number | boolean | null | undefined) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }

  const parsed = Number(String(value).trim().replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeBoolean(value: string | number | boolean | null | undefined) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "number") {
    return value === 1;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    return ["1", "si", "sí", "true", "yes", "y", "t"].includes(normalized);
  }

  return false;
}

function normalizeText(value: string | number | boolean | null | undefined) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}
