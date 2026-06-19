/**
 * __tests__/excel.test.ts
 *
 * Tests for lib/excel.ts:
 * - getCellText / hyperlink cell reading (fixes [object Object] bug)
 * - normalizeDate (UTC correctness — no timezone shift)
 * - readLoginExcel (hyperlink cells, missing columns)
 * - readClaimsExcel (date normalization, filtering empty rows)
 * - postProcessWorksheet (duplicate row Bot column propagation)
 */
import ExcelJS from 'exceljs';
import {
  normalizeDate,
  readLoginExcel,
  readClaimsExcel,
  postProcessWorksheet,
} from '@/lib/excel';

// ── Helper: build an in-memory Excel buffer ────────────────────────────────
async function buildWorkbook(
  headers: string[],
  rows: Record<string, unknown>[]
): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');

  // Write headers (ExcelJS cell index is 1-based)
  const hdrRow = ws.getRow(1);
  headers.forEach((h, i) => {
    hdrRow.getCell(i + 1).value = h;
  });
  hdrRow.commit();

  // Write data rows
  rows.forEach((row, rowIdx) => {
    const wsRow = ws.getRow(rowIdx + 2);
    headers.forEach((h, colIdx) => {
      wsRow.getCell(colIdx + 1).value = (row[h] ?? null) as ExcelJS.CellValue;
    });
    wsRow.commit();
  });

  return wb.xlsx.writeBuffer() as Promise<ArrayBuffer>;
}

// ── normalizeDate ─────────────────────────────────────────────────────────
describe('normalizeDate', () => {
  it('handles a JS Date at UTC midnight without timezone shift', () => {
    // 2026-04-16T00:00:00.000Z — must return 04/16/2026, NOT 04/15/2026
    const d = new Date('2026-04-16T00:00:00.000Z');
    expect(normalizeDate(d)).toBe('04/16/2026');
  });

  it('handles a string date passthrough', () => {
    expect(normalizeDate('04/28/2026')).toBe('04/28/2026');
  });

  it('returns empty string for null/undefined', () => {
    expect(normalizeDate(null)).toBe('');
    expect(normalizeDate(undefined)).toBe('');
  });

  it('handles hyperlink object (type 5) with .text property', () => {
    // ExcelJS represents hyperlink cells as { text, hyperlink }
    const hyperlinkVal = { text: '06/01/1949', hyperlink: 'mailto:06/01/1949' };
    expect(normalizeDate(hyperlinkVal as unknown as ExcelJS.CellValue)).toBe('06/01/1949');
  });
});

// ── readLoginExcel ────────────────────────────────────────────────────────
describe('readLoginExcel', () => {
  it('reads plain string credentials correctly', async () => {
    const buf = await buildWorkbook(
      ['Payer', 'URL', 'User Name', 'Password'],
      [{ 'Payer': 'UHC', 'URL': 'https://secure.uhcprovider.com', 'User Name': 'Watt@123', 'Password': 'Nop@ssword$2026' }]
    );
    const creds = await readLoginExcel(buf);
    expect(creds.username).toBe('Watt@123');
    expect(creds.password).toBe('Nop@ssword$2026');
    expect(creds.url).toContain('secure.uhcprovider.com');
  });

  it('reads hyperlink-type credentials (type 5 cells) — fixes [object Object] bug', async () => {
    // Build workbook with hyperlink cells
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Details');
    ws.getRow(1).getCell(1).value = 'Payer';
    ws.getRow(1).getCell(2).value = 'URL';
    ws.getRow(1).getCell(3).value = 'User Name';
    ws.getRow(1).getCell(4).value = 'Password';
    ws.getRow(1).commit();

    // Simulate hyperlink cells (as stored in the real UHC-Website Details.xlsx)
    ws.getRow(2).getCell(1).value = 'UHC';
    ws.getRow(2).getCell(2).value = {
      text: 'https://secure.uhcprovider.com/#/',
      hyperlink: 'https://secure.uhcprovider.com/',
    };
    ws.getRow(2).getCell(3).value = {
      text: 'Watt@123',
      hyperlink: 'mailto:Watt@123',
    };
    ws.getRow(2).getCell(4).value = {
      text: 'Nop@ssword$2026',
      hyperlink: 'mailto:Nop@ssword$2026',
    };
    ws.getRow(2).commit();

    const buf = await wb.xlsx.writeBuffer() as ArrayBuffer;
    const creds = await readLoginExcel(buf);

    // Must NOT be '[object Object]'
    expect(creds.username).toBe('Watt@123');
    expect(creds.password).toBe('Nop@ssword$2026');
    expect(creds.url).toBe('https://secure.uhcprovider.com');
  });

  it('throws if username or password columns are missing', async () => {
    const buf = await buildWorkbook(
      ['Payer', 'URL'],
      [{ 'Payer': 'UHC', 'URL': 'https://example.com' }]
    );
    await expect(readLoginExcel(buf)).rejects.toThrow(/Username\/Password/);
  });

  it('throws if data row is empty / credentials are blank', async () => {
    const buf = await buildWorkbook(
      ['User Name', 'Password'],
      [{ 'User Name': '', 'Password': '' }]
    );
    await expect(readLoginExcel(buf)).rejects.toThrow(/Username\/Password/);
  });
});

// ── readClaimsExcel ───────────────────────────────────────────────────────
describe('readClaimsExcel', () => {
  it('reads claim rows with correct column mapping', async () => {
    const buf = await buildWorkbook(
      ['Subscriber No', 'Patient DOB', 'Service Date'],
      [
        { 'Subscriber No': '95940073000', 'Patient DOB': '06/01/1949', 'Service Date': '04/16/2026' },
        { 'Subscriber No': '843125056',   'Patient DOB': '07/12/1972', 'Service Date': '04/28/2026' },
      ]
    );
    const rows = await readClaimsExcel(buf);
    expect(rows).toHaveLength(2);
    expect(rows[0].subscriberNo).toBe('95940073000');
    expect(rows[0].patientDOB).toBe('06/01/1949');
    expect(rows[0].serviceDate).toBe('04/16/2026');
    expect(rows[0].rowIndex).toBe(2);
    expect(rows[1].rowIndex).toBe(3);
  });

  it('normalizes Date objects in DOB and Service Date columns (UTC-safe)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getRow(1).getCell(1).value = 'Subscriber No';
    ws.getRow(1).getCell(2).value = 'Patient DOB';
    ws.getRow(1).getCell(3).value = 'Service Date';
    ws.getRow(1).commit();
    ws.getRow(2).getCell(1).value = '95940073000';
    ws.getRow(2).getCell(2).value = new Date('1949-06-01T00:00:00.000Z');  // Date type
    ws.getRow(2).getCell(3).value = new Date('2026-04-16T00:00:00.000Z');  // Date type
    ws.getRow(2).commit();
    const buf = await wb.xlsx.writeBuffer() as ArrayBuffer;

    const rows = await readClaimsExcel(buf);
    expect(rows[0].patientDOB).toBe('06/01/1949');   // Must NOT be shifted by IST (+5:30)
    expect(rows[0].serviceDate).toBe('04/16/2026');
  });

  it('filters out rows with empty Subscriber No', async () => {
    const buf = await buildWorkbook(
      ['Subscriber No', 'Patient DOB', 'Service Date'],
      [
        { 'Subscriber No': '',            'Patient DOB': '06/01/1949', 'Service Date': '04/16/2026' },
        { 'Subscriber No': '843125056',   'Patient DOB': '07/12/1972', 'Service Date': '04/28/2026' },
      ]
    );
    const rows = await readClaimsExcel(buf);
    expect(rows).toHaveLength(1);
    expect(rows[0].subscriberNo).toBe('843125056');
  });

  it('throws if required columns are missing', async () => {
    const buf = await buildWorkbook(
      ['Patient Name', 'Charges'],
      [{ 'Patient Name': 'Test', 'Charges': 100 }]
    );
    await expect(readClaimsExcel(buf)).rejects.toThrow(/Missing required columns/);
  });
});

// ── postProcessWorksheet ──────────────────────────────────────────────────
describe('postProcessWorksheet', () => {
  it('copies Bot columns from the first row to all duplicate (SubNo + ServiceDate) rows', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');

    // Header row
    ['Subscriber No', 'Service Date', 'BotClaimNumber', 'BotClaimStatus', 'BotStatus'].forEach((h, i) => {
      ws.getRow(1).getCell(i + 1).value = h;
    });
    ws.getRow(1).commit();

    // Two rows for the same subscriber + service date (duplicate claim lines)
    ws.getRow(2).getCell(1).value = '843125056';
    ws.getRow(2).getCell(2).value = '04/28/2026';
    ws.getRow(2).getCell(3).value = 'FR33275204';  // BotClaimNumber (already written)
    ws.getRow(2).getCell(4).value = 'Finalized';   // BotClaimStatus
    ws.getRow(2).getCell(5).value = 'Success';     // BotStatus
    ws.getRow(2).commit();

    ws.getRow(3).getCell(1).value = '843125056';   // same subscriber
    ws.getRow(3).getCell(2).value = '04/28/2026';  // same date → duplicate
    ws.getRow(3).getCell(3).value = null;           // BotClaimNumber not yet set
    ws.getRow(3).getCell(4).value = null;
    ws.getRow(3).getCell(5).value = null;
    ws.getRow(3).commit();

    const buf = await wb.xlsx.writeBuffer() as ArrayBuffer;
    const processedBuf = await postProcessWorksheet(buf);

    // Read back and verify
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(Buffer.from(processedBuf as ArrayBuffer) as any);
    const ws2 = wb2.worksheets[0];

    expect(ws2.getRow(3).getCell(3).value).toBe('FR33275204');  // copied
    expect(ws2.getRow(3).getCell(4).value).toBe('Finalized');   // copied
    expect(ws2.getRow(3).getCell(5).value).toBe('Success');     // copied
  });

  it('does not modify rows when there are no duplicates', async () => {
    const buf = await buildWorkbook(
      ['Subscriber No', 'Service Date', 'BotStatus'],
      [
        { 'Subscriber No': '111', 'Service Date': '04/01/2026', 'BotStatus': 'Success' },
        { 'Subscriber No': '222', 'Service Date': '04/02/2026', 'BotStatus': 'Error' },
      ]
    );
    const processedBuf = await postProcessWorksheet(buf);
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(Buffer.from(processedBuf as ArrayBuffer) as any);
    const ws2 = wb2.worksheets[0];

    // Values should remain unchanged
    expect(ws2.getRow(2).getCell(3).value).toBe('Success');
    expect(ws2.getRow(3).getCell(3).value).toBe('Error');
  });
});
