const MONTH_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

const MONTH_FULL = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

const MONTH_INDEX: Record<string, number> = Object.fromEntries(
  MONTH_SHORT.flatMap((m, i) => [
    [m.toLowerCase(), i],
    [MONTH_FULL[i].toLowerCase(), i],
  ])
);

const ISO_DATE_RE =
  /^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/;

function excelSerialToDate(serial: number): Date | null {
  if (!Number.isFinite(serial) || serial < 1 || serial >= 2958466) return null;
  const utcDays = Math.floor(serial - 25569);
  const ms = utcDays * 86400 * 1000;
  const date = new Date(ms);
  return Number.isNaN(date.getTime()) ? null : date;
}

export function normalizeExcelNumFmt(numFmt: string): string {
  return numFmt.replace(/\\"/g, " ").replace(/"/g, "").replace(/\s+/g, " ").trim();
}

export function parseDateValue(value: unknown): Date | null {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === "number") {
    return excelSerialToDate(value);
  }

  if (typeof value !== "string") return null;

  const trimmed = value.trim();
  if (!trimmed) return null;

  if (ISO_DATE_RE.test(trimmed)) {
    const date = new Date(trimmed);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const textDate = trimmed.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{2,4})$/);
  if (textDate) {
    const day = parseInt(textDate[1], 10);
    const month = MONTH_INDEX[textDate[2].toLowerCase()];
    let year = parseInt(textDate[3], 10);
    if (month === undefined) return null;
    if (year < 100) year += 2000;
    const date = new Date(year, month, day);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const slashDate = trimmed.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (slashDate) {
    const day = parseInt(slashDate[1], 10);
    const month = parseInt(slashDate[2], 10) - 1;
    let year = parseInt(slashDate[3], 10);
    if (year < 100) year += 2000;
    const date = new Date(year, month, day);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  return null;
}

export function extractDateFromCellValue(cellValue: unknown): Date | null {
  if (cellValue === null || cellValue === undefined) return null;
  if (cellValue instanceof Date) {
    return Number.isNaN(cellValue.getTime()) ? null : cellValue;
  }
  if (typeof cellValue === "object" && "result" in cellValue) {
    return extractDateFromCellValue((cellValue as { result: unknown }).result);
  }
  if (typeof cellValue === "number") {
    return excelSerialToDate(cellValue);
  }
  if (typeof cellValue === "string") {
    return parseDateValue(cellValue);
  }
  return null;
}

export function formatDateValue(date: Date, pattern: string): string {
  const pad = (n: number, len = 2) => String(n).padStart(len, "0");
  const replacements: Array<[string, string]> = [
    ["yyyy", String(date.getFullYear())],
    ["yy", String(date.getFullYear()).slice(-2)],
    ["MMMM", MONTH_FULL[date.getMonth()]],
    ["mmmm", MONTH_FULL[date.getMonth()]],
    ["MMM", MONTH_SHORT[date.getMonth()]],
    ["mmm", MONTH_SHORT[date.getMonth()]],
    ["MM", pad(date.getMonth() + 1)],
    ["dd", pad(date.getDate())],
    ["HH", pad(date.getHours())],
    ["hh", pad(date.getHours() % 12 || 12)],
    ["mm", pad(date.getMinutes())],
    ["ss", pad(date.getSeconds())],
  ];

  let result = normalizeExcelNumFmt(pattern);
  for (const [token, value] of replacements) {
    result = result.replaceAll(token, value);
  }
  return result;
}

export function resolveDateFormat(
  dateFormat: string | undefined,
  cellNumFmt?: string
): string | undefined {
  if (!dateFormat) return undefined;
  if (dateFormat === "cell") {
    return cellNumFmt ? normalizeExcelNumFmt(cellNumFmt) : undefined;
  }
  return dateFormat;
}
