/**
 * lib/excel.ts
 * Read/write Excel files using ExcelJS.
 */
import ExcelJS from 'exceljs';

export interface LoginCredentials {
  username: string;
  password: string;
  url: string;
}

export interface ClaimRow {
  rowIndex: number;      // 1-based Excel row number (row 2 = first data row)
  subscriberNo: string;
  patientDOB: string;    // MM/DD/YYYY
  serviceDate: string;   // MM/DD/YYYY
  patientName?: string;
  patientFirstName?: string;
  patientLastName?: string;
  [key: string]: unknown;
}

export interface BotFields {
  BotClaimNumber?: string;
  BotClaimStatus?: string;
  BotPaidAmount?: string;
  BotBilledAmount?: string;
  BotCheckEFTNumber?: string;
  BotDenialReasonCode?: string;
  BotDenialDescription?: string;
  BotRemarkCodes?: string;
  BotProcessedDate?: string;
  BotClaimDetails?: string;
  BotClaimResult?: string;
  BotUpdateTime?: string;
  BotStatus: 'Success' | 'Error' | 'Skipped';
  BotStatusError?: string;
}

/**
 * Safely extract a string from a cell value, handling:
 *  - Plain strings (type 3)
 *  - Hyperlink objects: { text: string, hyperlink: string } (type 5)
 *  - Dates (type 4)
 *  - Numbers (type 2)
 */
function getCellText(cell: ExcelJS.Cell): string {
  const val = cell.value;
  if (val === null || val === undefined) return '';
  // Hyperlink object: { text, hyperlink }
  if (typeof val === 'object' && 'text' in (val as object)) {
    return String((val as { text: unknown }).text ?? '').trim();
  }
  // Date object — use UTC to avoid timezone offset shifting the date
  if (val instanceof Date) {
    const m = String(val.getUTCMonth() + 1).padStart(2, '0');
    const d = String(val.getUTCDate()).padStart(2, '0');
    const y = val.getUTCFullYear();
    return `${m}/${d}/${y}`;
  }
  return String(val).trim();
}

/** Normalize an Excel date cell value to MM/DD/YYYY string (UTC-safe) */
export function normalizeDate(val: ExcelJS.CellValue): string {
  if (!val) return '';
  if (val instanceof Date) {
    const m = String(val.getUTCMonth() + 1).padStart(2, '0');
    const d = String(val.getUTCDate()).padStart(2, '0');
    return `${m}/${d}/${val.getUTCFullYear()}`;
  }
  // Hyperlink object
  if (typeof val === 'object' && 'text' in (val as object)) {
    return String((val as { text: unknown }).text ?? '').trim();
  }
  return String(val).trim();
}

/** Find a column index by header name (case-insensitive, partial match) */
function findCol(headerRow: ExcelJS.Row, ...names: string[]): number {
  let found = -1;
  headerRow.eachCell((cell, colNum) => {
    // Use getCellText so hyperlink-type headers also match
    const cellText = getCellText(cell).toLowerCase();
    for (const name of names) {
      if (cellText.includes(name.toLowerCase())) {
        found = colNum;
        break;
      }
    }
  });
  return found;
}

/** Read login credentials from the UHC-Website Details.xlsx workbook */
export async function readLoginExcel(buffer: ArrayBuffer): Promise<LoginCredentials> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);
  const ws = wb.worksheets[0];

  const headerRow = ws.getRow(1);
  const userCol = findCol(headerRow, 'username', 'user name', 'user');
  const passCol = findCol(headerRow, 'password', 'pass');
  const urlCol  = findCol(headerRow, 'url', 'website', 'link');

  const dataRow = ws.getRow(2);
  // Use getCellText to safely handle hyperlink-type cells (type 5)
  const username = userCol > 0 ? getCellText(dataRow.getCell(userCol)) : '';
  const password = passCol > 0 ? getCellText(dataRow.getCell(passCol)) : '';
  let url = process.env.UHC_URL ?? 'https://secure.uhcprovider.com';
  if (urlCol > 0) {
    const u = getCellText(dataRow.getCell(urlCol));
    // Strip fragment — the hyperlink href may differ from display text
    if (u) url = u.split('#')[0].replace(/\/$/, '');
  }

  if (!username || !password) {
    throw new Error('Could not find Username/Password in the login Excel. Check column headers.');
  }
  return { username, password, url };
}

/** Read all claim rows from the claims workbook */
export async function readClaimsExcel(buffer: ArrayBuffer): Promise<ClaimRow[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);
  const ws = wb.worksheets[0];

  const headerRow  = ws.getRow(1);
  const subNoCol   = findCol(headerRow, 'subscriber no', 'subscriber number', 'member id', 'memberid');
  const dobCol     = findCol(headerRow, 'patient dob', 'dob', 'date of birth', 'birth');
  const svcDateCol = findCol(headerRow, 'service date', 'dos', 'date of service');
  const nameCol    = findCol(headerRow, 'patient name', 'subscriber name', 'name', 'member name');
  const firstCol   = findCol(headerRow, 'first name', 'patient first name', 'subscriber first name');
  const lastCol    = findCol(headerRow, 'last name', 'patient last name', 'subscriber last name');

  if (subNoCol < 0 || dobCol < 0 || svcDateCol < 0) {
    throw new Error(
      `Missing required columns. ` +
      `Expected "Subscriber No", "Patient DOB", "Service Date". ` +
      `Found subNo col=${subNoCol}, dob col=${dobCol}, svcDate col=${svcDateCol}`
    );
  }

  const rows: ClaimRow[] = [];
  ws.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    // getCellText handles hyperlinks, plain strings, and numbers
    const subscriberNo = getCellText(row.getCell(subNoCol));
    if (!subscriberNo) return;
    rows.push({
      rowIndex:    rowNum,
      subscriberNo,
      patientDOB:  getCellText(row.getCell(dobCol))  || normalizeDate(row.getCell(dobCol).value),
      serviceDate: getCellText(row.getCell(svcDateCol)) || normalizeDate(row.getCell(svcDateCol).value),
      patientName: nameCol > 0 ? getCellText(row.getCell(nameCol)) : undefined,
      patientFirstName: firstCol > 0 ? getCellText(row.getCell(firstCol)) : undefined,
      patientLastName: lastCol > 0 ? getCellText(row.getCell(lastCol)) : undefined,
    });
  });
  return rows;
}

/**
 * Ensure all Bot* column headers exist in the worksheet.
 * Adds them after the last existing column if missing.
 */
export async function ensureBotHeaders(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const BOT_COLUMNS: (keyof BotFields)[] = [
    'BotClaimNumber', 'BotClaimStatus', 'BotPaidAmount', 'BotBilledAmount',
    'BotCheckEFTNumber', 'BotDenialReasonCode', 'BotDenialDescription',
    'BotRemarkCodes', 'BotProcessedDate', 'BotClaimDetails', 'BotClaimResult',
    'BotUpdateTime', 'BotStatus', 'BotStatusError',
  ];

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);
  const ws = wb.worksheets[0];
  const headerRow = ws.getRow(1);

  // Build existing header map
  const existing = new Map<string, number>();
  headerRow.eachCell((cell, colNum) => {
    existing.set(String(cell.value ?? '').trim(), colNum);
  });

  // Add missing Bot headers
  let nextCol = (headerRow.cellCount ?? 0) + 1;
  for (const col of BOT_COLUMNS) {
    if (!existing.has(col)) {
      headerRow.getCell(nextCol).value = col;
      existing.set(col, nextCol);
      nextCol++;
    }
  }
  headerRow.commit();

  return await wb.xlsx.writeBuffer() as ArrayBuffer;
}

/**
 * Write Bot fields for a specific row (by 1-based rowIndex).
 * Reads the current buffer, applies changes, returns new buffer.
 */
export async function writeBotFields(
  buffer: ArrayBuffer,
  rowIndex: number,
  fields: BotFields
): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);
  const ws = wb.worksheets[0];

  const headerRow = ws.getRow(1);
  const colMap = new Map<string, number>();
  headerRow.eachCell((cell, colNum) => {
    colMap.set(String(cell.value ?? '').trim(), colNum);
  });

  const dataRow = ws.getRow(rowIndex);
  for (const [key, value] of Object.entries(fields)) {
    const colNum = colMap.get(key);
    if (colNum) dataRow.getCell(colNum).value = value ?? null;
  }
  dataRow.commit();

  return await wb.xlsx.writeBuffer() as ArrayBuffer;
}

/**
 * Post-processing: copy Bot columns from the first row in each
 * (subscriberNo, serviceDate) group to all duplicate rows.
 */
export async function postProcessWorksheet(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);
  const ws = wb.worksheets[0];

  const headerRow = ws.getRow(1);
  const subNoCol   = findCol(headerRow, 'subscriber no', 'subscriber number', 'member id');
  const svcDateCol = findCol(headerRow, 'service date', 'dos', 'date of service');

  const botColumnNumbers: number[] = [];
  headerRow.eachCell((cell, colNum) => {
    if (String(cell.value ?? '').startsWith('Bot')) botColumnNumbers.push(colNum);
  });

  if (subNoCol < 0 || svcDateCol < 0 || botColumnNumbers.length === 0) {
    return await wb.xlsx.writeBuffer() as ArrayBuffer;
  }

  const groups = new Map<string, number[]>();
  ws.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const subNo   = String(row.getCell(subNoCol).value ?? '').trim();
    const svcDate = normalizeDate(row.getCell(svcDateCol).value);
    if (!subNo) return;
    const key = `${subNo}|${svcDate}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(rowNum);
  });

  for (const [, rowNums] of groups) {
    if (rowNums.length < 2) continue;
    const sourceRow = ws.getRow(rowNums[0]);
    for (let i = 1; i < rowNums.length; i++) {
      const targetRow = ws.getRow(rowNums[i]);
      for (const botColNum of botColumnNumbers) {
        targetRow.getCell(botColNum).value = sourceRow.getCell(botColNum).value;
      }
      targetRow.commit();
    }
  }

  return await wb.xlsx.writeBuffer() as ArrayBuffer;
}
