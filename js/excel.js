import { SHEET_NAMES, LISTS } from "./schema.js";

/* ------------------------------------------------------------------ *
 * Utilidades
 * ------------------------------------------------------------------ */

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86400000;

export function toSerial(date) {
  return (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - EXCEL_EPOCH) / DAY;
}

export function parseISODate(value) {
  if (!value) return null;
  const [y, m, d] = String(value).split("-").map(Number);
  if (!y || !m || !d) return null;
  const date = new Date(y, m - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) return null;
  return date;
}

/** Converte um valor de data do Excel (número de série ou Date) em "AAAA-MM-DD". */
export function serialToISO(value) {
  if (value === null || value === undefined || value === "") return "";
  if (value instanceof Date) {
    const pad = (n) => String(n).padStart(2, "0");
    return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
  }
  if (typeof value === "number") {
    const d = new Date(EXCEL_EPOCH + value * DAY);
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  }
  const text = String(value).trim();
  const br = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (br) {
    const year = br[3].length === 2 ? `20${br[3]}` : br[3];
    return `${year}-${br[2].padStart(2, "0")}-${br[1].padStart(2, "0")}`;
  }
  return "";
}

export function formatDate(value) {
  if (value === null || value === undefined || value === "") return "";
  if (value instanceof Date) {
    return value.toLocaleDateString("pt-BR");
  }
  if (typeof value === "number") {
    const d = new Date(EXCEL_EPOCH + value * DAY);
    return d.toLocaleDateString("pt-BR", { timeZone: "UTC" });
  }
  return String(value);
}

/** Formato aplicado somente quando a célula de data está em "Geral". */
const DATE_FORMAT = "dd/mm/yyyy";

/** Converte Date, número de série, "AAAA-MM-DD" ou "DD/MM/AAAA" em número de série. */
function dateSerial(value) {
  if (value instanceof Date) return toSerial(value);
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const text = String(value === null || value === undefined ? "" : value).trim();
  if (!text) return null;

  const br = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(text);
  if (br) {
    const year = Number(br[3].length === 2 ? `20${br[3]}` : br[3]);
    const d = new Date(year, Number(br[2]) - 1, Number(br[1]));
    return Number.isNaN(d.getTime()) ? null : toSerial(d);
  }

  const iso = parseISODate(text);
  return iso ? toSerial(iso) : null;
}

/**
 * Compara um valor esperado com o valor atual da célula.
 * Em campos de data aceita Date, número de série e texto, comparando só o dia.
 */
function sameValue(expected, actual, field) {
  if (field && field.type === "date") {
    const a = dateSerial(expected);
    const b = dateSerial(actual);
    if (a === null || b === null) return a === null && b === null;
    return Math.floor(a) === Math.floor(b);
  }

  const a = expected === null || expected === undefined ? "" : String(expected).trim();
  const b = actual === null || actual === undefined ? "" : String(actual).trim();

  if (typeof expected === "number" || typeof actual === "number") {
    const na = Number(a);
    const nb = Number(b);
    if (a !== "" && b !== "" && Number.isFinite(na) && Number.isFinite(nb)) return na === nb;
  }
  return a === b;
}

function isPlainFormat(code) {
  const text = String(code === null || code === undefined ? "" : code).trim();
  return text === "" || /^general$/i.test(text);
}

function isEmpty(v) {
  return v === null || v === undefined || v === "";
}

function delay(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

/** Converte erro do Office.js em mensagem legível. */
export function describeError(err) {
  if (!err) return "Erro desconhecido.";
  const code = err.code || err.name || "";
  const msg = err.message || String(err);

  if (/AccessDenied|PermissionDenied/i.test(code) || /protected|proteç/i.test(msg)) {
    return "A planilha está protegida e a célula de destino não permite edição. " +
      "Siga o procedimento de liberação descrito na aba Ajuda.";
  }
  if (/ItemNotFound|NotFound/i.test(code)) {
    return "A aba esperada não foi encontrada neste arquivo. Confira se você abriu a planilha correta.";
  }
  if (/InvalidArgument|InvalidReference/i.test(code)) {
    return "Referência inválida na planilha. Avise o responsável técnico.";
  }
  if (/GeneralException|Unknown/i.test(code)) {
    return "O Excel recusou a operação: " + msg;
  }
  return msg;
}

/* ------------------------------------------------------------------ *
 * Leitura das listas
 * ------------------------------------------------------------------ */

let listsCache = null;

export async function loadLists() {
  if (listsCache) return listsCache;

  const result = await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(SHEET_NAMES.listas);
    const entries = Object.entries(LISTS);
    const ranges = entries.map(([, spec]) => {
      const r = ws.getRange(spec.range);
      r.load("values");
      return r;
    });
    await context.sync();

    const out = {};
    entries.forEach(([key], i) => {
      out[key] = ranges[i].values
        .flat()
        .map((v) => (v === null || v === undefined ? "" : String(v).trim()))
        .filter((v) => v.length > 0);
    });
    return out;
  });

  listsCache = result;
  return result;
}

export function invalidateLists() {
  listsCache = null;
}

/* ------------------------------------------------------------------ *
 * Colunas
 * ------------------------------------------------------------------ */

function colIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function colName(index) {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function buildCells(table, payload, includeEmpty = false) {
  const has = (key) => Object.prototype.hasOwnProperty.call(payload, key);
  const cells = [];
  const invalid = [];

  for (const field of table.fields) {
    const raw = payload[field.key];

    // Em edição, só limpa a célula quando o campo veio no formulário.
    const clear = includeEmpty && has(field.key);

    if (isEmpty(raw)) {
      if (clear) cells.push({ field, value: "", col: field.col });
      continue;
    }

    if (field.type === "date") {
      // Grava o número de série: evita deslocamento de fuso/DST na data.
      const serial = raw instanceof Date ? toSerial(raw) : dateSerial(raw);
      if (serial === null) {
        invalid.push(field);
        if (clear) cells.push({ field, value: "", col: field.col });
        continue;
      }
      cells.push({ field, value: serial, isDate: true, col: field.col });
      continue;
    }

    const value = String(raw).trim();
    if (!value) {
      if (clear) cells.push({ field, value: "", col: field.col });
      continue;
    }
    cells.push({ field, value, col: field.col });
  }

  return { cells, invalid };
}

/** A data de referência é o que define a próxima linha livre: não pode faltar. */
function missingReferenceDate(table, cells) {
  return !cells.some((cell) => cell.col === table.dateCol && !isEmpty(cell.value));
}

function invalidDateMessage(invalid) {
  return `Data inválida em: ${invalid.map((f) => f.label).join(", ")}. Use dd/mm/aaaa.`;
}

/* ------------------------------------------------------------------ *
 * Localização da próxima linha livre
 * ------------------------------------------------------------------ */

async function findNextRow(context, table) {
  const ws = context.workbook.worksheets.getItem(table.sheet);
  const addr = `${table.dateCol}${table.firstDataRow}:${table.dateCol}${table.lastRowGuard}`;
  const range = ws.getRange(addr);
  range.load("values");
  await context.sync();

  const values = range.values;
  let lastIndex = -1;
  for (let i = 0; i < values.length; i += 1) {
    if (!isEmpty(values[i][0]) && values[i][0] !== 0) lastIndex = i;
  }
  return table.firstDataRow + lastIndex + 1;
}

/** Mostra a próxima linha livre de uma aba (usado no diagnóstico). */
export async function nextFreeRow(table) {
  return Excel.run(async (context) => findNextRow(context, table));
}

/* ------------------------------------------------------------------ *
 * Gravação (novo lançamento)
 * ------------------------------------------------------------------ */

async function tryAppend(table, payload) {
  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(table.sheet);
    const row = await findNextRow(context, table);

    const { cells, invalid } = buildCells(table, payload);
    if (invalid.length > 0) {
      return { ok: false, reason: "invalid", message: invalidDateMessage(invalid) };
    }
    if (cells.length === 0) {
      return { ok: false, reason: "empty", message: "Preencha ao menos um campo." };
    }
    if (missingReferenceDate(table, cells)) {
      return {
        ok: false,
        reason: "invalid",
        message: "Informe a data do documento: é ela que define a próxima linha livre da aba.",
      };
    }

    if (table.editableUntil && row > table.editableUntil) {
      return {
        ok: false,
        reason: "blocked",
        row,
        message: `A próxima linha livre (${row}) está fora do trecho liberado para edição ` +
          `(até a linha ${table.editableUntil}). É preciso liberar mais linhas na aba "${table.sheet.trim()}".`,
      };
    }

    // Lê a linha de destino antes de escrever (detecta ocupação por outro usuário).
    const rowRange = ws.getRange(`A${row}:Y${row}`);
    rowRange.load("values");

    // Células de data no formato "Geral" recebem formatação de data.
    const dateCells = cells.filter((c) => c.isDate);
    const dateRanges = dateCells.map((cell) => {
      const r = ws.getRange(`${cell.col}${row}`);
      r.load("numberFormat");
      return r;
    });

    await context.sync();

    const before = rowRange.values[0];
    const dateIndex = colIndex(table.dateCol);
    if (!isEmpty(before[dateIndex])) {
      return { ok: false, reason: "collision", row, message: "Linha ocupada por outro lançamento." };
    }

    dateRanges.forEach((r) => {
      if (isPlainFormat(r.numberFormat[0][0])) r.numberFormat = [[DATE_FORMAT]];
    });

    for (const cell of cells) {
      ws.getRange(`${cell.col}${row}`).values = [[cell.value]];
    }
    await context.sync();

    // Conferência: relê o que foi gravado.
    const check = ws.getRange(`A${row}:Y${row}`);
    check.load("values");
    const protocolCell = ws.getRange(`${table.protocolCol}${row}`);
    protocolCell.load("values");
    await context.sync();

    const after = check.values[0];
    const mismatched = cells.filter((cell) => !sameValue(cell.value, after[colIndex(cell.col)], cell.field));

    if (mismatched.length > 0) {
      return {
        ok: false,
        reason: "collision",
        row,
        mixedRow: row,
        message: "Outro lançamento ocupou esta linha ao mesmo tempo.",
      };
    }

    const protocol = protocolCell.values[0][0];
    return {
      ok: true,
      row,
      protocol,
      cells: cells.map((c) => ({ label: c.field.label, col: c.col, row })),
    };
  });
}

/**
 * Grava um lançamento na próxima linha livre.
 * Faz até 3 tentativas quando há disputa pela mesma linha.
 */
export async function appendRow(table, payload, onAttempt) {
  const MAX = 3;
  let last = null;
  const warnings = [];

  for (let attempt = 1; attempt <= MAX; attempt += 1) {
    try {
      const result = await tryAppend(table, payload);
      if (result.ok) return { ...result, attempt, warnings };
      last = result;

      if (result.mixedRow) {
        warnings.push(
          `A linha ${result.mixedRow} recebeu dados de outra pessoa no mesmo instante. ` +
          `Confira-a na planilha e ajuste manualmente se necessário.`
        );
      }

      if (result.reason !== "collision") return { ...result, warnings };
      if (onAttempt) onAttempt(attempt, result.row);
      await delay(300 + Math.random() * 500);
    } catch (err) {
      return { ok: false, reason: "error", message: describeError(err), warnings };
    }
  }
  return last
    ? { ...last, warnings }
    : { ok: false, reason: "busy", message: "Não foi possível reservar uma linha livre.", warnings };
}

/* ------------------------------------------------------------------ *
 * Alteração de lançamento existente
 * ------------------------------------------------------------------ */

/**
 * Altera a linha indicada, gravando apenas as colunas de digitação.
 * `expected` (opcional) é o estado lido no início da edição: se a linha
 * tiver mudado nesse meio tempo, devolve `reason: "conflict"` para o
 * painel pedir confirmação antes de sobrescrever.
 */
export async function updateRow(table, row, payload, expected) {
  try {
    return await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(table.sheet);
      const addr = `A${row}:Y${row}`;

      const { cells, invalid } = buildCells(table, payload, true);
      if (invalid.length > 0) {
        return { ok: false, reason: "invalid", row, message: invalidDateMessage(invalid) };
      }
      if (missingReferenceDate(table, cells)) {
        return {
          ok: false,
          reason: "invalid",
          row,
          message: "A data do documento não pode ficar vazia: é ela que define a linha na planilha.",
        };
      }

      const range = ws.getRange(addr);
      range.load("values");

      const dateRanges = cells
        .filter((c) => c.isDate && !isEmpty(c.value))
        .map((cell) => {
          const r = ws.getRange(`${cell.col}${row}`);
          r.load("numberFormat");
          return r;
        });

      await context.sync();
      const current = range.values[0];

      if (expected) {
        const changed = table.fields.filter(
          (field) => !sameValue(expected[field.key] ?? "", current[colIndex(field.col)], field)
        );
        if (changed.length > 0) {
          return {
            ok: false,
            reason: "conflict",
            row,
            fields: changed.map((f) => f.label),
            message: `A linha ${row} foi alterada por outra pessoa ` +
              `(${changed.slice(0, 3).map((f) => f.label).join(", ")}).`,
          };
        }
      }

      dateRanges.forEach((r) => {
        if (isPlainFormat(r.numberFormat[0][0])) r.numberFormat = [[DATE_FORMAT]];
      });

      for (const cell of cells) {
        ws.getRange(`${cell.col}${row}`).values = [[cell.value]];
      }
      await context.sync();

      const check = ws.getRange(addr);
      check.load("values");
      await context.sync();
      const after = check.values[0];

      const mismatched = cells.filter((cell) => !sameValue(cell.value, after[colIndex(cell.col)], cell.field));
      if (mismatched.length > 0) {
        return {
          ok: false,
          reason: "error",
          row,
          message: "A gravação não foi confirmada: outra pessoa pode ter alterado esta linha ao mesmo tempo.",
        };
      }

      return { ok: true, row, protocol: after[colIndex(table.protocolCol)] };
    });
  } catch (err) {
    return { ok: false, reason: "error", row, message: describeError(err) };
  }
}

/* ------------------------------------------------------------------ *
 * Consulta
 * ------------------------------------------------------------------ */

export async function readLastRows(table, howMany = 300) {
  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(table.sheet);
    const last = await findNextRow(context, table);
    const endRow = Math.max(table.firstDataRow, last - 1);
    const startRow = Math.max(table.firstDataRow, endRow - howMany + 1);

    const range = ws.getRange(`A${startRow}:Y${endRow}`);
    range.load("values");
    await context.sync();

    return range.values
      .map((cells, i) => {
        const rowNumber = startRow + i;
        const obj = { __sheet: table.key, __row: rowNumber };
        for (let c = 0; c < 25; c += 1) {
          obj[colName(c)] = cells[c];
        }
        return obj;
      })
      .filter((r) => !isEmpty(r[table.dateCol]))
      .reverse();
  });
}

export async function focusRow(table, row) {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(table.sheet);
      ws.activate();
      const r = ws.getRange(`A${row}:Y${row}`);
      r.select();
      await context.sync();
    });
    return true;
  } catch {
    return false;
  }
}

/* ------------------------------------------------------------------ *
 * Diagnóstico
 * ------------------------------------------------------------------ */

export async function diagnostics(tables) {
  return Excel.run(async (context) => {
    const wb = context.workbook;
    const sheets = wb.worksheets;
    sheets.load("items/name");
    await context.sync();

    const names = sheets.items.map((s) => s.name);

    let workbookName = "";
    try {
      wb.load("name");
      await context.sync();
      workbookName = wb.name || "";
    } catch {
      workbookName = "";
    }

    const found = {};

    for (const [key, table] of Object.entries(tables)) {
      if (!names.includes(table.sheet)) {
        found[key] = { present: false };
        continue;
      }

      const info = {
        present: true,
        protected: null,
        rowCount: null,
        columnCount: null,
        nextRow: null,
        editableUntil: table.editableUntil,
      };

      try {
        const ws = wb.worksheets.getItem(table.sheet);
        const prot = ws.protection;
        prot.load("protected");
        await context.sync();
        info.protected = prot.protected;
      } catch {
        info.protected = null;
      }

      try {
        const ws = wb.worksheets.getItem(table.sheet);
        const used = ws.getUsedRange(true);
        used.load("rowCount,columnCount");
        await context.sync();
        info.rowCount = used.rowCount;
        info.columnCount = used.columnCount;
      } catch {
        info.rowCount = null;
      }

      try {
        info.nextRow = await findNextRow(context, table);
      } catch {
        info.nextRow = null;
      }

      found[key] = info;
    }

    return { workbook: workbookName, sheets: names, found };
  });
}
