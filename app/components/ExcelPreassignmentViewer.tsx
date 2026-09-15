"use client";

import { ChangeEvent, useMemo, useState } from "react";
import Image from "next/image";
import JSZip from "jszip";
import { parsePreassignments, type PreasignacionRow } from "../lib/excel";

const MAX_XLSX_SIZE_BYTES = 10 * 1024 * 1024;

function hasSupportedExcelSignature(buffer: ArrayBuffer, fileName: string) {
  const bytes = new Uint8Array(buffer, 0, Math.min(buffer.byteLength, 8));

  if (/\.xlsx$/i.test(fileName)) {
    return (
      bytes.length >= 4 &&
      bytes[0] === 0x50 &&
      bytes[1] === 0x4b &&
      bytes[2] === 0x03 &&
      bytes[3] === 0x04
    );
  }

  const xlsSignature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  return (
    bytes.length === xlsSignature.length &&
    xlsSignature.every((byte, index) => bytes[index] === byte)
  );
}

function formatNumber(value: number | null) {
  if (value === null) {
    return "—";
  }

  return value.toLocaleString("es-ES");
}

export default function ExcelPreassignmentViewer() {
  const [fileRows, setFileRows] = useState<[PreasignacionRow[], PreasignacionRow[]]>([[], []]);
  const [search, setSearch] = useState("");
  const [status, setStatus] = useState("REVISIÓN DE PREASIGNACIONES");
  const [loadedFileNames, setLoadedFileNames] = useState<[string | null, string | null]>([null, null]);
  const rows = useMemo(() => fileRows.flat(), [fileRows]);

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>, fileIndex: 0 | 1) {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    setLoadedFileNames((current) =>
      current.map((name, index) => index === fileIndex ? null : name) as [string | null, string | null],
    );

    if (!/\.(xls|xlsx)$/i.test(file.name)) {
      setStatus("El archivo debe tener formato .xls o .xlsx.");
      setFileRows((current) =>
        current.map((items, index) => index === fileIndex ? [] : items) as [PreasignacionRow[], PreasignacionRow[]],
      );
      event.target.value = "";
      return;
    }

    if (file.size > MAX_XLSX_SIZE_BYTES) {
      setStatus("El archivo supera el límite permitido de 10 MB.");
      setFileRows((current) =>
        current.map((items, index) => index === fileIndex ? [] : items) as [PreasignacionRow[], PreasignacionRow[]],
      );
      event.target.value = "";
      return;
    }

    try {
      setStatus("Leyendo archivo...");
      const buffer = await file.arrayBuffer();

      if (!hasSupportedExcelSignature(buffer, file.name)) {
        throw new Error("El archivo no tiene una estructura Excel válida.");
      }

      const parsedRows = await parsePreassignments(buffer);
      setFileRows((current) =>
        current.map((items, index) => index === fileIndex ? parsedRows : items) as [PreasignacionRow[], PreasignacionRow[]],
      );
      setLoadedFileNames((current) =>
        current.map((name, index) => index === fileIndex ? file.name : name) as [string | null, string | null],
      );
      setStatus(`Se procesaron ${parsedRows.length} filas del archivo.`);
    } catch (error) {
      console.error(error);
      setStatus(
        error instanceof Error
          ? error.message
          : "No se pudo leer el archivo. Intenta con otro archivo.",
      );
      setFileRows((current) =>
        current.map((items, index) => index === fileIndex ? [] : items) as [PreasignacionRow[], PreasignacionRow[]],
      );
      event.target.value = "";
    }
  }

  const filteredRows = useMemo(() => {
    const normalizedSearch = search.trim().toLowerCase();

    const candidateRows = rows.filter((row) => {
      if (!normalizedSearch) {
        return true;
      }

      const hayMatch =
        row.clienteNumero.toLowerCase() === normalizedSearch ||
        row.clienteNombre.toLowerCase().includes(normalizedSearch);

      return hayMatch;
    });

    return candidateRows;
  }, [rows, search]);

  async function handleDownloadExcel() {
    if (filteredRows.length === 0) {
      setStatus("No hay resultados para descargar.");
      return;
    }

    setStatus('Generando archivo (preservando plantilla)...');

    try {
      const resp = await fetch('/ET.xlsx');
      if (!resp.ok) throw new Error('No se pudo descargar la plantilla');
      const arrayBuffer = await resp.arrayBuffer();

      const zip = await JSZip.loadAsync(arrayBuffer);

      const workbookXml = await zip.file('xl/workbook.xml')!.async('string');
      const parser = new DOMParser();
      const workbookDoc = parser.parseFromString(workbookXml, 'application/xml');
      const firstSheet = workbookDoc.getElementsByTagName('sheet')[0];
      const rId = firstSheet?.getAttribute('r:id');
      if (!rId) throw new Error('No se pudo leer la hoja de la plantilla');

      const relsXml = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
      const relMatch = relsXml.match(new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*Target="([^"]+)"`));
      if (!relMatch) throw new Error('No se pudo resolver la relación de la hoja');
      const sheetPath = 'xl/' + relMatch[1].replace(/^\//, '');

      const headerRow = 8;
      const dataStartRow = headerRow + 1; // 9
      let sheetXml = await zip.file(sheetPath)!.async('string');
      let sharedStringsXml = await zip.file('xl/sharedStrings.xml')!.async('string');
      let stylesXml = await zip.file('xl/styles.xml')!.async('string');
      const cellStyleCount = Number(stylesXml.match(/<cellXfs\b[^>]*\bcount="(\d+)"/)?.[1]);
      if (!Number.isFinite(cellStyleCount)) {
        throw new Error('No se pudieron leer los estilos de la plantilla');
      }
      const totalCellStyle = cellStyleCount;
      stylesXml = stylesXml
        .replace(
          /(<cellXfs\b[^>]*\bcount=")\d+("[^>]*>)/,
          `$1${cellStyleCount + 1}$2`,
        )
        .replace(
          '</cellXfs>',
          '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs>',
        );
      const initialSharedStringCount = Number(sharedStringsXml.match(/\bcount="(\d+)"/)?.[1] ?? 0);
      const initialUniqueCount = Number(sharedStringsXml.match(/\buniqueCount="(\d+)"/)?.[1] ?? 0);
      const newSharedStrings = new Map<string, number>();
      let stringReferenceCount = 0;

      function escapeXml(value: string) {
        return value
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;')
          .replace(/'/g, '&apos;');
      }

      function sharedStringIndex(value: string) {
        stringReferenceCount += 1;
        const existingIndex = newSharedStrings.get(value);
        if (existingIndex !== undefined) return existingIndex;

        const index = initialUniqueCount + newSharedStrings.size;
        newSharedStrings.set(value, index);
        return index;
      }

      function cellXml(address: string, value: string | number, style?: number) {
        const styleAttribute = style === undefined ? '' : ` s="${style}"`;
        if (typeof value === 'number' && Number.isFinite(value)) {
          return `<c r="${address}"${styleAttribute} t="n"><v>${value}</v></c>`;
        }
        if (value === '') return `<c r="${address}"${styleAttribute}/>`;
        return `<c r="${address}"${styleAttribute} t="s"><v>${sharedStringIndex(String(value))}</v></c>`;
      }

      function setCell(address: string, value: string | number, style?: number) {
        const rowNumber = Number(address.match(/\d+/)?.[0]);
        const replacement = cellXml(address, value, style);
        const rowPattern = new RegExp(`<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`);
        const rowMatch = sheetXml.match(rowPattern);

        if (!rowMatch) {
          const newRow = `<row r="${rowNumber}">${replacement}</row>`;
          sheetXml = sheetXml.replace('</sheetData>', `${newRow}</sheetData>`);
          return;
        }

        const cellPattern = new RegExp(`<c\\b[^>]*\\br="${address}"(?:\\s[^>]*)?(?:\\/>|>[\\s\\S]*?<\\/c>)`);
        let updatedRow = rowMatch[0];
        if (cellPattern.test(updatedRow)) {
          updatedRow = updatedRow.replace(cellPattern, replacement);
        } else {
          const followingCellPattern = new RegExp(`(?=<c\\b[^>]*\\br="[E-Z]${rowNumber}")`);
          updatedRow = followingCellPattern.test(updatedRow)
            ? updatedRow.replace(followingCellPattern, replacement)
            : updatedRow.replace('</row>', `${replacement}</row>`);
        }
        sheetXml = sheetXml.replace(rowPattern, updatedRow);
      }

      // Write client info into B5 and B6
      const clienteVal = filteredRows[0]?.clienteNombre || '';
      const numeroVal = filteredRows[0]?.clienteNumero || '';
      setCell('B5', clienteVal, 4);
      setCell('B6', numeroVal, 5);


      const columnWidths = ['CODIGO', 'DETALLE', 'CANTIDAD', 'PREASIGNADO', 'ESTADO']
        .map((header) => header.length);
      const fixedColumnWidths: Array<number | null> = [10, null, 10, 14, 24];
      const dataColumnStyles = [2, 5, 2, 2, 2]; // B left; A and C-E centered

      // Fill data rows starting at dataStartRow
      filteredRows.forEach((row, idx) => {
        const r = dataStartRow + idx; // 9+
        const vals = [
          row.codigoProducto || '',
          row.nombreProducto || '',
          row.cantidadSolicitada ?? '',
          row.cantidadPreasignada ?? '',
          row.preasignado ? 'LISTO PARA DESPACHAR' : '',
        ];

        vals.forEach((val, colIdx) => {
          columnWidths[colIdx] = Math.max(columnWidths[colIdx], String(val).length);
          const colLetter = String.fromCharCode(65 + colIdx); // A,B,C,D,E
          const addr = `${colLetter}${r}`;
          setCell(addr, val, dataColumnStyles[colIdx]);
        });
      });


      columnWidths.forEach((contentLength, colIdx) => {
        const columnNumber = colIdx + 1;
        const width = fixedColumnWidths[colIdx] ?? Math.min(Math.max(contentLength + 2, 10), 255);
        const columnPattern = new RegExp(
          `<col\\b(?=[^>]*\\bmin="${columnNumber}")(?=[^>]*\\bmax="${columnNumber}")[^>]*/>`,
        );
        sheetXml = sheetXml.replace(columnPattern, (columnXml) =>
          columnXml.replace(/\bwidth="[^"]*"/, `width="${width}"`),
        );
      });
      const lastDataRow = dataStartRow + filteredRows.length - 1;
      const totalRow = lastDataRow + 1;
      const totalPreasignado = filteredRows.reduce(
        (total, row) => total + (row.cantidadPreasignada ?? 0),
        0,
      );
      const rowPattern = /<row\b[^>]*\br="(\d+)"[^>]*(?:\/>|>[\s\S]*?<\/row>)/g;
      sheetXml = sheetXml.replace(rowPattern, (rowXml, rowNumber) =>
        Number(rowNumber) > lastDataRow ? '' : rowXml
      );

      setCell(`A${totalRow}`, 'TOTAL ARTÍCULOS A DESPACHAR', totalCellStyle);
      setCell(`D${totalRow}`, totalPreasignado, totalCellStyle);

      const totalMergeRef = `A${totalRow}:C${totalRow}`;
      if (/<mergeCells\b[^>]*>/.test(sheetXml)) {
        sheetXml = sheetXml.replace(
          /<mergeCells\b([^>]*)>([\s\S]*?)<\/mergeCells>/,
          (_mergeCellsXml, attributes, mergeCells) => {
            const currentCount = Number(String(attributes).match(/\bcount="(\d+)"/)?.[1] ?? 0);
            const updatedAttributes = /\bcount="\d+"/.test(attributes)
              ? String(attributes).replace(/\bcount="\d+"/, `count="${currentCount + 1}"`)
              : `${attributes} count="${currentCount + 1}"`;
            return `<mergeCells${updatedAttributes}>${mergeCells}<mergeCell ref="${totalMergeRef}"/></mergeCells>`;
          },
        );
      } else {
        sheetXml = sheetXml.replace(
          /(?=<pageMargins\b)/,
          `<mergeCells count="1"><mergeCell ref="${totalMergeRef}"/></mergeCells>`,
        );
      }

      sheetXml = sheetXml
        .replace(/<dimension ref="[^"]+"/, `<dimension ref="A2:G${totalRow}"`)
        .replace(/<sortState ref="A9:[A-Z]+\d+"/, `<sortState ref="A9:E${lastDataRow}"`)
        .replace(/<sortCondition ref="B9:B\d+"/, `<sortCondition ref="B9:B${lastDataRow}"`);

      const addedSharedStrings = Array.from(newSharedStrings.keys())
        .map((value) => `<si><t xml:space="preserve">${escapeXml(value)}</t></si>`)
        .join('');
      sharedStringsXml = sharedStringsXml
        .replace(/\bcount="\d+"/, `count="${initialSharedStringCount + stringReferenceCount}"`)
        .replace(/\buniqueCount="\d+"/, `uniqueCount="${initialUniqueCount + newSharedStrings.size}"`)
        .replace('</sst>', `${addedSharedStrings}</sst>`);

      const tablePath = 'xl/tables/table1.xml';
      const tableFile = zip.file(tablePath);
      if (tableFile) {
        const tableXml = (await tableFile.async('string'))
          .replace(/\bref="A8:[A-Z]+\d+"/, `ref="A8:E${lastDataRow}"`);
        zip.file(tablePath, tableXml);
      }

      zip.file(sheetPath, sheetXml);
      zip.file('xl/sharedStrings.xml', sharedStringsXml);
      zip.file('xl/styles.xml', stylesXml);

      const outArray = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
      const blob = new Blob([outArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      const numeroCliente = filteredRows[0]?.clienteNumero || 'cliente';
      const nombreCliente = filteredRows[0]?.clienteNombre || 'sinnombre';
      const nombreArchivo = `${numeroCliente} ${nombreCliente}`.replace(/[^a-zA-Z0-9]+/g, ' ').trim().replace(/\s+/g, ' ').toUpperCase();

      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${nombreArchivo}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
      setStatus('Archivo Excel descargado (plantilla preservada).');
    } catch (err) {
      console.error(err);
      setStatus('Error generando el archivo con la plantilla.');
    }
  }

  return (
    <main className="min-h-screen bg-[#f5f5f5] px-4 py-10 text-slate-700 sm:px-6 lg:px-8">
      <div className="mx-auto flex max-w-6xl flex-col gap-6">
        <header className="space-y-2 rounded-[2rem] border border-slate-200 bg-white p-6 shadow-lg sm:p-8">
          <div className="flex items-start justify-between gap-4">
            <Image
              src="/LOGO.svg"
              alt="Logo"
              width={585}
              height={217}
              priority
              className="h-auto w-full max-w-[260px]"
            />
            <Image
              src="/MEC.svg"
              alt="MEC"
              width={981}
              height={1205}
              priority
              className="h-auto w-[70px] shrink-0 sm:w-[90px]"
            />
          </div>
          <p className="text-sm font-semibold uppercase tracking-[0.3em] text-[#d9534f]">
            Revisión de pendientes
          </p>
          <div className="flex items-center justify-between gap-4">
            <h1 className="text-3xl font-semibold sm:text-4xl text-[#333333]">
              ADMINISTRACIÓN Y VENTAS
            </h1>
            <Image
              src="/ISO.svg"
              alt="Certificación ISO"
              width={1434}
              height={725}
              className="h-auto w-[110px] shrink-0 sm:w-[150px]"
            />
          </div>
          <p className="max-w-3xl text-sm text-slate-600 sm:text-base">
            Sube archivo/s de Excel desde PRESEA con la extensión .xls o .xlsx para revisar los pendientes. Luego filtra por número o nombre del cliente para poder visualizarlo de una manera más fácil.
          </p>
        </header>

        <section className="grid gap-10 rounded-[2rem] border border-slate-200 bg-white p-6 shadow-lg sm:p-8">
          <div className="flex flex-col gap-2">
            <span className="text-lg font-extrabold uppercase tracking-[0.08em] text-[#EE4B25] sm:text-xl">CARGA DE ARCHIVOS EXCEL</span>
            <div className="grid gap-4 sm:grid-cols-2">
              {([0, 1] as const).map((fileIndex) => (
                <div key={fileIndex}>
                  <label className="block w-full cursor-pointer rounded-2xl border border-slate-300 bg-slate-50 px-4 py-3 text-center text-sm font-semibold text-slate-900 shadow-sm transition hover:border-orange-400">
                    {fileIndex === 0
                      ? "Excel CMF, CAE, CPE"
                      : "Excel CR"}
                    <input
                      type="file"
                      accept=".xls,.xlsx"
                      onChange={(event) => handleFileChange(event, fileIndex)}
                      className="sr-only"
                    />
                  </label>
                  {loadedFileNames[fileIndex] && (
                    <p className="mt-2 text-center text-sm text-green-600">
                      Archivo cargado correctamente: {loadedFileNames[fileIndex]}
                    </p>
                  )}
                </div>
              ))}
            </div>
          </div>

          <label className="flex flex-col gap-2">
            <span className="text-lg font-extrabold uppercase tracking-[0.08em] text-[#EE4B25] sm:text-xl">BUSCAR POR CLIENTE</span>
            <div className="w-full">
              <input
                type="text"
                value={search}
                onChange={(event) => setSearch(event.target.value)}
                placeholder="Número o nombre del cliente"
                className="w-full rounded-2xl border border-slate-300 bg-slate-50 px-4 py-3 text-center text-sm font-semibold text-slate-900 shadow-sm outline-none transition placeholder:text-slate-900 hover:border-orange-400 focus:border-orange-400 focus:ring-2 focus:ring-orange-100"
              />
            </div>
          </label>
        </section>

        <section className="rounded-[2rem] border border-slate-200 bg-white p-6 shadow-lg sm:p-8">
          <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
            <p className="text-lg font-extrabold uppercase tracking-[0.08em] text-[#EE4B25] sm:text-xl">{status}</p>
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={handleDownloadExcel}
                className="rounded-full bg-[#f28c28] px-3 py-1.5 text-xs font-semibold text-white transition hover:bg-[#e07b14]"
              >
                Descargar Excel
              </button>
              <span className="rounded-full border border-[#f28c28]/40 bg-[#f28c28]/10 px-3 py-1 text-xs font-semibold uppercase tracking-[0.2em] text-[#b85c00]">
                {filteredRows.length} resultados
              </span>
            </div>
          </div>

          {filteredRows.length === 0 ? (
            <div className="rounded-2xl border border-dashed border-[#d0d0d0] bg-white p-8 text-center text-sm text-slate-500">
              Aún no hay preasignaciones para mostrar. Prueba con otro cliente o sube un archivo distinto.
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="min-w-full border-collapse text-left text-sm">
                <thead>
                  <tr className="border-b border-[#e0e0e0] text-slate-600">
                    <th className="px-3 py-3 font-medium">Número</th>
                    <th className="px-3 py-3 font-medium">Cliente</th>
                    <th className="px-3 py-3 font-medium">Solicitado</th>
                    <th className="px-3 py-3 font-medium">Estado</th>
                    <th className="px-3 py-3 font-medium">Preasignado</th>
                    <th className="px-3 py-3 font-medium">Código</th>
                    <th className="px-3 py-3 font-medium">Producto</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.map((row, index) => (
                    <tr key={`${row.clienteNumero}-${row.codigoProducto}-${index}`} className="border-b border-[#ececec] text-slate-700">
                      <td className="px-3 py-3 font-semibold">{row.clienteNumero || "Sin número"}</td>
                      <td className="px-3 py-3">
                        <div className="font-semibold">{row.clienteNombre || "Sin nombre"}</div>
                      </td>
                      <td className="px-3 py-3">{formatNumber(row.cantidadSolicitada)}</td>
                      <td className="px-3 py-3">
                        <span
                          className={`rounded-full px-2.5 py-1 text-xs font-semibold ${
                            row.preasignado
                              ? "bg-[#e8f7ec] text-[#1f7a3d]"
                              : "bg-[#fff3e8] text-[#b85c00]"
                          }`}
                        >
                          {row.preasignado ? "Preasignado" : "Sin preasignación"}
                        </span>
                      </td>
                      <td className="px-3 py-3">{formatNumber(row.cantidadPreasignada)}</td>
                      <td className="px-3 py-3 font-semibold">{row.codigoProducto || "Sin código"}</td>
                      <td className="px-3 py-3">
                        <div className="font-semibold">{row.nombreProducto || "Sin producto"}</div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>
      </div>
    </main>
  );
}
