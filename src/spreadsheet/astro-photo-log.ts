import { existsSync } from 'node:fs';
import { createSheetsClient } from './sheets.js';
import { SERVICE_ACCOUNT_FILE } from '../env.js';

export const LOG_ENTRY_HEADERS = [
  'Date (YYYY-MM-DD)',
  'Target/Object Name',
  'Object Type',
  'Catalog #',
  'Exposure Time per Sub (seconds)',
  'Number of Subs / Total Integration Time (seconds)',
  'Filter Used',
  'Gain / ISO',
  'RA',
  'DEC',
  'TELESCOP',
  'FITS File Path',
  'Notes / Issues',
];

export async function formatHeader(spreadsheetId: string, sheetName = 'Astro Photo Log') {
  try {
    const sheets = createSheetsClient();

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheet = spreadsheet.data.sheets?.find(s => s.properties?.title === sheetName);
    const sheetId = sheet?.properties?.sheetId ?? 0;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.95, green: 0.95, blue: 0.95 },
                  textFormat: { bold: true, fontSize: 11 },
                  horizontalAlignment: 'CENTER',
                },
              },
              fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
            },
          },
          {
            updateSheetProperties: {
              properties: { sheetId, gridProperties: { frozenRowCount: 1 } },
              fields: 'gridProperties.frozenRowCount',
            },
          },
          {
            autoResizeDimensions: {
              dimensions: { sheetId, dimension: 'COLUMNS' },
            },
          },
        ],
      },
    });
  } catch (err: any) {
    console.warn('⚠️ Header formatting skipped:', err.message);
  }
}

function getColumnLetter(columnIndex: number) {
  let dividend = columnIndex + 1;
  let columnName = '';
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return columnName;
}

export async function addObjectTypeColorRules(spreadsheetId: string, sheetName = 'Astro Photo Log') {
  try {
    const sheets = createSheetsClient();

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheet = spreadsheet.data.sheets?.find(s => s.properties?.title === sheetName);
    const sheetId = sheet?.properties?.sheetId ?? 0;

    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!1:1`,
    });
    const headers = headerRes.data.values?.[0] || [];
    const objectTypeColIndex = headers.findIndex(
      (header) => String(header).trim().toUpperCase() === 'OBJECT TYPE'
    );
    const finalObjectTypeColIndex = objectTypeColIndex >= 0 ? objectTypeColIndex : 2;
    const objectTypeColumn = getColumnLetter(finalObjectTypeColIndex);

    const existingRuleCount = sheet?.conditionalFormats?.length ?? 0;
    console.log(`🗑️  Found ${existingRuleCount} existing color rule(s) to replace`);
    if (existingRuleCount > 0) {
      const deleteRequests: any[] = [];
      for (let i = existingRuleCount - 1; i >= 0; i -= 1) {
        deleteRequests.push({ deleteConditionalFormatRule: { sheetId, index: i } });
      }
      try {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests: deleteRequests },
        });
      } catch (delErr: any) {
        console.warn(`⚠️ Could not delete ${existingRuleCount} old color rule(s): ${delErr.message}`);
      }
    }

    const objectTypeCell = `$${objectTypeColumn}2`;
    const cleanedObjectType = `UPPER(TRIM(REGEXEXTRACT(${objectTypeCell}, "^[^\\(]+")))`;

    const colorRules = [
      { pattern: 'Nebula',       color: { red: 0.85, green: 0.7,  blue: 0.95 } },
      { pattern: 'Galaxy',       color: { red: 0.7,  green: 0.8,  blue: 1.0  } },
      { pattern: 'Star Cluster', color: { red: 1.0,  green: 0.92, blue: 0.45 } },
      { pattern: 'Cluster',      color: { red: 1.0,  green: 0.85, blue: 0.7  } },
      { pattern: 'Planet',       color: { red: 1.0,  green: 0.95, blue: 0.7  } },
      { pattern: 'Moon',         color: { red: 1.0,  green: 0.95, blue: 0.7  } },
      { pattern: 'Sun',          color: { red: 1.0,  green: 0.95, blue: 0.7  } },
      { pattern: 'Milky Way',    color: { red: 0.7,  green: 0.95, blue: 0.9  } },
      { pattern: 'Landscape',    color: { red: 0.8,  green: 1.0,  blue: 0.8  } },
    ];

    const addRequests: any[] = [];

    colorRules.forEach((rule, index) => {
      const patternRegex = rule.pattern.toUpperCase().replace(/\s+/g, '\\s+');
      addRequests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{ sheetId, startRowIndex: 1, startColumnIndex: 0, endColumnIndex: 26 }],
            booleanRule: {
              condition: {
                type: 'CUSTOM_FORMULA',
                values: [{ userEnteredValue: `=REGEXMATCH(${cleanedObjectType}, "^${patternRegex}$")` }],
              },
              format: { backgroundColor: rule.color },
            },
          },
          index,
        },
      });
    });

    const masterPattern = colorRules
      .map(r => r.pattern.toUpperCase().replace(/\s+/g, '\\s+'))
      .join('|');

    addRequests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{ sheetId, startRowIndex: 1, startColumnIndex: 0, endColumnIndex: 26 }],
          booleanRule: {
            condition: {
              type: 'CUSTOM_FORMULA',
              values: [{ userEnteredValue: `=AND(NOT(ISBLANK(${objectTypeCell})), NOT(REGEXMATCH(${cleanedObjectType}, "^(${masterPattern})$")))` }],
            },
            format: { backgroundColor: { red: 0.95, green: 0.95, blue: 0.95 } },
          },
        },
        index: colorRules.length,
      },
    });

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: addRequests },
    });

    console.log('🎨 Applied color coding rules for Object Type column');
  } catch (err: any) {
    console.warn('⚠️ Color rules skipped:', err.message);
  }
}

/**
 * Ensures the Astro Photo Log has correct headers and color rules.
 * Called unconditionally on every run so checklists can always detect columns.
 */
/**
 * Writes correct headers and applies header formatting.
 * Does NOT apply color rules — call addObjectTypeColorRules separately,
 * AFTER any values.append calls, because INSERT_ROWS shifts conditional
 * format range start indices and leaves rows uncolored.
 */
export async function ensureSheetSetup(spreadsheetId: string, sheetName = 'Astro Photo Log') {
  if (!existsSync(SERVICE_ACCOUNT_FILE)) return;

  try {
    const sheets = createSheetsClient();

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: [LOG_ENTRY_HEADERS] },
    });

    await formatHeader(spreadsheetId, sheetName);

    console.log('🔧 Astro Photo Log headers verified');
  } catch (err: any) {
    console.warn('⚠️ Sheet setup failed:', err.message);
  }
}

export async function appendToGoogleSheet(spreadsheetId: string, rows: any[]) {
  if (!existsSync(SERVICE_ACCOUNT_FILE)) {
    console.warn('⚠️ service-account.json not found – skipping Google Sheets export.');
    return false;
  }

  try {
    const sheets = createSheetsClient();
    const sheetName = 'Astro Photo Log';

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetName}!A:A`,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        majorDimension: 'ROWS',
        values: rows.map(row => Object.values(row)),
      },
    });

    console.log(`✅ Successfully appended ${rows.length} new session(s) to Google Sheets`);
    return true;
  } catch (err: any) {
    console.error('❌ Google Sheets append failed:', err.message);
    if (err.message.includes('403')) {
      console.error('   → Make sure the service account email has **Editor** access to the spreadsheet.');
    }
    return false;
  }
}
