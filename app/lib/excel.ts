import JSZip, { type JSZipObject } from "jszip";
import * as XLSX from "xlsx";

const MAX_SHEETS = 20;
const MAX_ROWS = 50_000;
const MAX_COLUMNS = 100;
const MAX_CELL_LENGTH = 10_000;
type CellValue = string | number | boolean | null | undefined;

export type PreasignacionRow = {
  cantidadSolicitada: number | null;
  cantidadPreasignada: number | null;
  preasignado: boolean;
  clienteNumero: string;
  clienteNombre: string;
  codigoProducto: string;
  nombreProducto: string;
};

export async function parsePreassignments(buffer: ArrayBuffer): Promise<PreasignacionRow[]> {
  const format = detectExcelFormat(buffer);
  if (format === "unknown") {
    throw new Error("El archivo no es un Excel válido: no contiene una estructura XLSX (ZIP) ni XLS (OLE).");
  }
  if (format === "xls") return parseWithSheetJs(buffer);

  const archive = await loadAndValidateXlsx(buffer);
  try {
    return parseWithSheetJs(buffer);
  } catch (sheetJsError) {
    try {
      return rowsToPreassignments(await readFirstSheetFromOoxml(archive));
    } catch (fallbackError) {
      const reason = errorMessage(fallbackError) || errorMessage(sheetJsError);
      throw new Error(`El archivo XLSX está dañado o incompleto: ${reason}`);
    }
  }
}

function detectExcelFormat(buffer: ArrayBuffer): "xlsx" | "xls" | "unknown" {
  const bytes = new Uint8Array(buffer, 0, Math.min(buffer.byteLength, 8));
  const zip = bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b &&
    [0x03, 0x05, 0x07].includes(bytes[2]) && [0x04, 0x06, 0x08].includes(bytes[3]);
  if (zip) return "xlsx";
  const ole = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  return bytes.length === ole.length && ole.every((byte, index) => bytes[index] === byte)
    ? "xls" : "unknown";
}

async function loadAndValidateXlsx(buffer: ArrayBuffer): Promise<JSZip> {
  let archive: JSZip;
  try {
    archive = await JSZip.loadAsync(buffer, { checkCRC32: true });
  } catch (error) {
    const encrypted = errorMessage(error).toLowerCase().includes("encrypted");
    throw new Error(encrypted
      ? "El archivo XLSX está cifrado o protegido con contraseña."
      : "El archivo XLSX está dañado: el contenedor ZIP interno es inválido o tiene datos corruptos.");
  }
  const types = archive.file("[Content_Types].xml");
  if (!types || !archive.file("xl/workbook.xml") || !archive.file("xl/_rels/workbook.xml.rels")) {
    throw new Error("El archivo no es un XLSX válido: faltan componentes obligatorios del libro.");
  }
  if (!/spreadsheetml|sheet\.main\+xml/i.test(await types.async("string"))) {
    throw new Error("El archivo no es un XLSX válido: el contenedor no declara un libro de Excel.");
  }
  return archive;
}

function parseWithSheetJs(buffer: ArrayBuffer): PreasignacionRow[] {
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(buffer, {
      type: "array", dense: true, WTF: false, bookDeps: false, bookVBA: false,
      cellFormula: false, cellHTML: false, cellStyles: false, sheetRows: MAX_ROWS + 1,
    });
  } catch (error) {
    throw new Error(`SheetJS no pudo interpretar el libro: ${errorMessage(error) || "formato interno desconocido"}`);
  }
  if (workbook.SheetNames.length > MAX_SHEETS) {
    throw new Error(`El archivo supera el límite de ${MAX_SHEETS} hojas.`);
  }
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet || !sheet["!ref"]) return [];
  validateSheetRange(sheet["!ref"]);
  return rowsToPreassignments(XLSX.utils.sheet_to_json<CellValue[]>(sheet, {
    header: 1, raw: false, defval: "",
  }));
}

async function readFirstSheetFromOoxml(archive: JSZip): Promise<CellValue[][]> {
  const workbook = parseXml(await requiredFile(archive, "xl/workbook.xml").async("string"), "el libro principal no contiene XML válido");
  const sheets = Array.from(workbook.getElementsByTagName("sheet"));
  if (sheets.length > MAX_SHEETS) throw new Error(`el archivo supera el límite de ${MAX_SHEETS} hojas`);
  if (!sheets.length) return [];
  const relationshipId = sheets[0].getAttribute("r:id") ?? sheets[0].getAttributeNS(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
  if (!relationshipId) throw new Error("la primera hoja no tiene una relación válida");

  const rels = parseXml(await requiredFile(archive, "xl/_rels/workbook.xml.rels").async("string"), "las relaciones del libro no contienen XML válido");
  const relation = Array.from(rels.getElementsByTagName("Relationship"))
    .find((item) => item.getAttribute("Id") === relationshipId);
  const target = relation?.getAttribute("Target");
  if (!target) throw new Error("no se pudo localizar el contenido de la primera hoja");
  const sheet = parseXml(await requiredFile(archive, resolveTarget(target)).async("string"), "la primera hoja no contiene XML válido");
  const sharedStrings = await readSharedStrings(archive);
  const rows: CellValue[][] = [];

  for (const rowElement of Array.from(sheet.getElementsByTagName("row"))) {
    const declaredRow = Number(rowElement.getAttribute("r"));
    const rowIndex = Number.isInteger(declaredRow) && declaredRow > 0 ? declaredRow - 1 : rows.length;
    if (rowIndex >= MAX_ROWS) throw new Error(`la hoja supera el límite de ${MAX_ROWS.toLocaleString("es-AR")} filas`);
    const row: CellValue[] = [];
    for (const cell of Array.from(rowElement.getElementsByTagName("c"))) {
      const column = columnIndex(cell.getAttribute("r") ?? "");
      if (column >= MAX_COLUMNS) throw new Error(`la hoja supera el límite de ${MAX_COLUMNS} columnas`);
      row[column] = readOoxmlCell(cell, sharedStrings);
    }
    rows[rowIndex] = row;
  }
  return rows;
}

async function readSharedStrings(archive: JSZip): Promise<string[]> {
  const file = archive.file("xl/sharedStrings.xml");
  if (!file) return [];
  const document = parseXml(await file.async("string"), "la tabla de textos compartidos no contiene XML válido");
  return Array.from(document.getElementsByTagName("si")).map((item) =>
    Array.from(item.getElementsByTagName("t")).map((text) => text.textContent ?? "").join(""));
}

function readOoxmlCell(cell: Element, strings: string[]): CellValue {
  const type = cell.getAttribute("t");
  const value = cell.getElementsByTagName("v")[0]?.textContent ?? "";
  if (type === "inlineStr") return Array.from(cell.getElementsByTagName("t"))
    .map((text) => text.textContent ?? "").join("");
  if (type === "s") return strings[Number(value)] ?? "";
  if (type === "b") return value === "1";
  if (type === "str" || type === "e" || value === "") return value;
  const number = Number(value);
  return Number.isFinite(number) ? number : value;
}

function resolveTarget(target: string): string {
  const normalized = target.replace(/\\/g, "/").replace(/^\//, "");
  const parts: string[] = [];
  for (const part of (normalized.startsWith("xl/") ? normalized : `xl/${normalized}`).split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop(); else parts.push(part);
  }
  return parts.join("/");
}

function requiredFile(archive: JSZip, path: string): JSZipObject {
  const file = archive.file(path);
  if (!file) throw new Error(`falta el componente obligatorio “${path}”`);
  return file;
}

function parseXml(xml: string, reason: string): Document {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length) throw new Error(reason);
  return document;
}

function columnIndex(reference: string): number {
  const letters = reference.match(/^[A-Za-z]+/)?.[0];
  if (!letters) throw new Error("una celda tiene una referencia inválida");
  let index = 0;
  for (const letter of letters.toUpperCase()) index = index * 26 + letter.charCodeAt(0) - 64;
  return index - 1;
}

function validateSheetRange(reference: string) {
  let range: XLSX.Range;
  try { range = XLSX.utils.decode_range(reference); }
  catch { throw new Error("La hoja principal contiene un rango inválido."); }
  if (range.e.r - range.s.r + 1 > MAX_ROWS) {
    throw new Error(`La hoja supera el límite de ${MAX_ROWS.toLocaleString("es-AR")} filas.`);
  }
  if (range.e.c - range.s.c + 1 > MAX_COLUMNS) {
    throw new Error(`La hoja supera el límite de ${MAX_COLUMNS} columnas.`);
  }
}

function rowsToPreassignments(rows: CellValue[][]): PreasignacionRow[] {
  if (!rows.length) return [];
  if (rows.some((row) => row?.some((cell) => String(cell ?? "").length > MAX_CELL_LENGTH))) {
    throw new Error("El archivo contiene una celda excesivamente extensa.");
  }
  return rows.slice(1)
    .filter((row) => row?.some((cell) => String(cell ?? "").trim() !== ""))
    .map((row) => ({
      cantidadSolicitada: normalizeNumber(row[1]), cantidadPreasignada: normalizeNumber(row[5]),
      preasignado: normalizeBoolean(row[6]), clienteNumero: normalizeText(row[10]),
      clienteNombre: normalizeText(row[11]), codigoProducto: normalizeText(row[13]),
      nombreProducto: normalizeText(row[14]),
    }))
    .filter((row) => row.clienteNumero || row.clienteNombre || row.codigoProducto ||
      row.nombreProducto || row.cantidadPreasignada !== null);
}

function errorMessage(error: unknown) {
  return error instanceof Error ? error.message.replace(/^Error:\s*/i, "") : "";
}
function normalizeNumber(value: CellValue) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "boolean") return value ? 1 : 0;
  const parsed = Number(String(value).trim().replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}
function normalizeBoolean(value: CellValue) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  return typeof value === "string" && ["1", "si", "sí", "true", "yes", "y", "t"]
    .includes(value.trim().toLowerCase());
}
function normalizeText(value: CellValue) {
  return value === null || value === undefined ? "" : String(value).trim();
}
