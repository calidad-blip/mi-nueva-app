"use client";

import { ChangeEvent, useMemo, useState } from "react";
import JSZip from "jszip";
import { parsePreassignments, type PreasignacionRow } from "../lib/excel";

function formatNumber(value: number | null) {
  if (value === null) {
    return "—";
  }

  return value.toLocaleString("es-ES");
}

export default function ExcelPreassignmentViewer() {
  const [rows, setRows] = useState<PreasignacionRow[]>([]);
  const [search, setSearch] = useState("");
  const [status, setStatus] = useState("Sube un archivo Excel para revisar las preasignaciones.");

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    if (!/\.(xls|xlsx)$/i.test(file.name)) {
      setStatus("El archivo debe ser un Excel .xls o .xlsx.");
      setRows([]);
      return;
    }

    try {
      setStatus("Leyendo archivo...");
      const buffer = await file.arrayBuffer();
      const parsedRows = parsePreassignments(buffer);
      setRows(parsedRows);
      setStatus(`Se procesaron ${parsedRows.length} filas del archivo.`);
    } catch (error) {
      console.error(error);
      setStatus("No se pudo leer el archivo. Intenta con otro archivo o formato.");
      setRows([]);
    }
  }

  const filteredRows = useMemo(() => {
    const normalizedSearch = search.trim().toLowerCase();

    const candidateRows = rows.filter((row) => {
      if (!normalizedSearch) {
        return true;
      }

      const hayMatch =
        row.clienteNumero.toLowerCase().includes(normalizedSearch) ||
        row.clienteNombre.toLowerCase().includes(normalizedSearch) ||
        row.codigoProducto.toLowerCase().includes(normalizedSearch) ||
        row.nombreProducto.toLowerCase().includes(normalizedSearch);

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
      const resp = await fetch('/375%20pendientes.xlsx');
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

      const sheetXml = await zip.file(sheetPath)!.async('string');
      const doc = parser.parseFromString(sheetXml, 'application/xml');

      const sheetData = doc.getElementsByTagName('sheetData')[0];
      if (!sheetData) throw new Error('sheetData no encontrado en la plantilla');

      const headerRow = 8;
      const dataStartRow = headerRow + 1; // 9

      function getCellNode(address: string) {
        const cells = doc.getElementsByTagName('c');
        for (let i = 0; i < cells.length; i++) {
          const c = cells[i];
          if (c.getAttribute('r') === address) return c;
        }
        return null;
      }

      function ensureRow(rowNum: number) {
        const existingRows = sheetData.getElementsByTagName('row');
        for (let i = 0; i < existingRows.length; i++) {
          if (existingRows[i].getAttribute('r') === String(rowNum)) return existingRows[i];
        }
        const newRow = doc.createElement('row');
        newRow.setAttribute('r', String(rowNum));
        sheetData.appendChild(newRow);
        return newRow;
      }

      function setCell(address: string, value: string | number, copyStyleFrom?: string | null) {
        const rowNum = parseInt(address.replace(/[^0-9]/g, ''), 10);
        let cell = getCellNode(address);
        let styleAttr: string | null = null;
        if (copyStyleFrom) {
          const tpl = getCellNode(copyStyleFrom);
          if (tpl && tpl.getAttribute('s')) styleAttr = tpl.getAttribute('s');
        }

        if (!cell) {
          const rowNode = ensureRow(rowNum);
          cell = doc.createElement('c');
          cell.setAttribute('r', address);
          if (styleAttr) cell.setAttribute('s', styleAttr);
          rowNode.appendChild(cell);
        }

        // remove existing children
        while (cell.firstChild) cell.removeChild(cell.firstChild);

        if (typeof value === 'number' && !Number.isNaN(value)) {
          cell.setAttribute('t', 'n');
          const v = doc.createElement('v');
          v.textContent = String(value);
          cell.appendChild(v);
        } else {
          cell.setAttribute('t', 'inlineStr');
          const is = doc.createElement('is');
          const t = doc.createElement('t');
          // preserve whitespace
          t.setAttribute('xml:space', 'preserve');
          t.textContent = String(value);
          is.appendChild(t);
          cell.appendChild(is);
        }
      }

      // Write client info into B5 and B6
      const clienteVal = filteredRows[0]?.clienteNombre || '';
      const numeroVal = filteredRows[0]?.clienteNumero || '';
      // choose template style cells if exist (B5/B6 or fallback A5/A6)
      const tplCliente = 'B5';
      const tplNumero = 'B6';
      const fallbackCliente = 'A5';
      const fallbackNumero = 'A6';
      const clienteStyleFrom = getCellNode(tplCliente) ? tplCliente : (getCellNode(fallbackCliente) ? fallbackCliente : null);
      const numeroStyleFrom = getCellNode(tplNumero) ? tplNumero : (getCellNode(fallbackNumero) ? fallbackNumero : null);
      setCell('B5', clienteVal, clienteStyleFrom);
      setCell('B6', numeroVal, numeroStyleFrom);

      // Fill data rows starting at dataStartRow
      filteredRows.forEach((row, idx) => {
        const r = dataStartRow + idx; // 9+
        const vals = [
          (row.codigoProducto || '').toUpperCase(),
          (row.nombreProducto || '').toUpperCase(),
          row.cantidadPreasignada ?? '',
          row.preasignado ? 'LISTO PARA DESPACHAR' : '',
        ];

        vals.forEach((val, colIdx) => {
          const colLetter = String.fromCharCode(65 + colIdx); // A,B,C,D
          const addr = `${colLetter}${r}`;
          // try copy style from template data row (row 9) else header row
          const styleFrom = getCellNode(`${colLetter}${dataStartRow}`) ? `${colLetter}${dataStartRow}` : `${colLetter}${headerRow}`;
          setCell(addr, val, styleFrom);
        });
      });

      const serializer = new XMLSerializer();
      const updatedSheetXml = serializer.serializeToString(doc);
      zip.file(sheetPath, updatedSheetXml);

      const outArray = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
      const blob = new Blob([outArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      const numeroCliente = filteredRows[0]?.clienteNumero || 'cliente';
      const nombreCliente = filteredRows[0]?.clienteNombre || 'sinnombre';
      const nombreArchivo = `${numeroCliente} ${nombreCliente}`.replace(/[^a-zA-Z0-9]+/g, ' ').trim().replace(/\s+/g, ' ').toLowerCase();

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
      <div className="mx-auto flex max-w-6xl flex-col gap-6 rounded-3xl border border-[#d9d9d9] bg-white p-6 shadow-[0_12px_40px_rgba(0,0,0,0.08)]">
        <header className="space-y-2">
          <p className="text-sm font-semibold uppercase tracking-[0.3em] text-[#d9534f]">
            Revisión de pendientes
          </p>
          <h1 className="text-3xl font-semibold sm:text-4xl text-[#333333]">
            ADMINISTRACIÓN Y VENTAS
          </h1>
          <p className="max-w-3xl text-sm text-slate-600 sm:text-base">
            Sube un archivo de excel desde PRESEA para comenzar
          </p>
        </header>

        <section className="grid gap-4 rounded-2xl border border-[#e5e5e5] bg-[#fafafa] p-4 md:grid-cols-[1.2fr_0.8fr]">
          <label className="flex flex-col gap-2">
            <span className="text-sm font-medium text-slate-700">Archivo Excel</span>
            <label className="flex cursor-pointer items-center justify-center rounded-xl border border-dashed border-[#c7c7c7] bg-white px-3 py-3 text-sm text-slate-700 hover:bg-[#fdf5eb]">
              <span className="rounded-full bg-[#f28c28] px-4 py-2 text-sm font-semibold text-white">
                Seleccione archivo de excel
              </span>
              <input
                type="file"
                accept=".xls,.xlsx"
                onChange={handleFileChange}
                className="sr-only"
              />
            </label>
          </label>

          <label className="flex flex-col gap-2">
            <span className="text-sm font-medium text-slate-700">Buscar por cliente</span>
            <div className="relative flex min-h-[52px] w-full items-center rounded-xl border border-dashed border-[#c7c7c7] bg-white px-3 py-3 text-sm text-slate-700 hover:bg-[#fdf5eb]">
              <input
                type="text"
                value={search}
                onChange={(event) => setSearch(event.target.value)}
                className="w-full bg-transparent text-sm text-slate-700 outline-none ring-0"
              />
              {!search && (
                <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 rounded-full bg-[#f28c28] px-3 py-1 text-xs font-semibold text-white">
                  Número o nombre del cliente
                </span>
              )}
            </div>
          </label>
        </section>

        <section className="rounded-2xl border border-[#e5e5e5] bg-[#fafafa] p-4">
          <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
            <p className="text-sm text-slate-600">{status}</p>
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
