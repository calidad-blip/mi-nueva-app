import * as XLSX from "xlsx";

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
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  if (!sheet) {
    return [];
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
  }) as Array<Array<string | number | boolean | null | undefined>>;

  if (rows.length === 0) {
    return [];
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
