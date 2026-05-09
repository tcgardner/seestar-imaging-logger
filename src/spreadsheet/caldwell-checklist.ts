import { createSheetsClient } from './sheets.js';
import { ASTRONOMICAL_OBJECTS } from '../astronomical-objects.js';
import {
  findHeaderRow,
  detectColumns,
  debugChecklistDetection,
  buildCatalogLogs,
  normalizeCatalogId,
} from './checklist-utils.js';

// Maps each Caldwell key to its NGC/IC equivalents so log entries stored by NGC/IC
// catalog ID (as the Seestar FITS headers do) can be matched to Caldwell objects.
const CALDWELL_NGC_ALIASES: Record<string, string[]> = {
  C1: ['NGC188'], C2: ['NGC40'], C3: ['NGC4236'], C4: ['NGC7023'],
  C5: ['IC342'], C6: ['NGC6543'], C7: ['NGC2403'], C8: ['NGC559'],
  C10: ['NGC663'], C11: ['NGC7635'], C12: ['NGC6946'], C13: ['NGC457'],
  C14: ['NGC869', 'NGC884'], C15: ['NGC6826'], C16: ['NGC7243'],
  C17: ['NGC147'], C18: ['NGC185'], C19: ['IC5146'], C20: ['NGC7000'],
  C21: ['NGC4449'], C22: ['NGC7662'], C23: ['NGC891'], C24: ['NGC1275'],
  C25: ['NGC2419'], C26: ['NGC4244'], C27: ['NGC6888'], C28: ['NGC752'],
  C29: ['NGC5005'], C30: ['NGC7331'], C31: ['IC405'], C32: ['NGC4631'],
  C33: ['NGC6992', 'NGC6995'], C34: ['NGC6960'], C35: ['NGC4889'],
  C36: ['NGC4559'], C37: ['NGC6885'], C38: ['NGC4565'], C39: ['NGC2392'],
  C40: ['NGC3626'], C42: ['NGC7006'], C43: ['NGC7814'], C44: ['NGC7479'],
  C45: ['NGC5248'], C46: ['NGC2261'], C47: ['NGC6934'], C48: ['NGC2775'],
  C49: ['NGC2237', 'NGC2244'], C50: ['NGC2244'], C51: ['NGC1613'],
  C52: ['NGC4697'], C53: ['NGC3115'], C54: ['NGC2506'], C55: ['NGC7009'],
  C56: ['NGC246'], C57: ['NGC6822'], C58: ['NGC2360'], C59: ['NGC3242'],
  C60: ['NGC4038'], C61: ['NGC4039'], C62: ['NGC247'], C63: ['NGC7293'],
  C64: ['NGC2362'], C65: ['NGC253'], C66: ['NGC5694'], C67: ['NGC1097'],
  C68: ['NGC6729'], C69: ['NGC6302'], C70: ['NGC300'], C71: ['NGC2477'],
  C72: ['NGC55'], C73: ['NGC1851'], C74: ['NGC3132'], C75: ['NGC6124'],
  C76: ['NGC6231'], C77: ['NGC5128'], C78: ['NGC6541'], C79: ['NGC3201'],
  C80: ['NGC5139'], C81: ['NGC6352'], C82: ['NGC6193'], C83: ['NGC4945'],
  C84: ['NGC5286'], C85: ['IC2391'], C86: ['NGC6397'], C87: ['NGC1261'],
  C88: ['NGC5823'], C89: ['NGC6087'], C90: ['NGC2867'], C91: ['NGC3532'],
  C92: ['NGC3372'], C93: ['NGC6752'], C94: ['NGC4755'], C95: ['NGC6025'],
  C96: ['NGC2516'], C97: ['NGC3766'], C98: ['NGC4609'], C100: ['IC2944'],
  C101: ['NGC6744'], C102: ['IC2602'], C103: ['NGC2070'], C104: ['NGC362'],
  C105: ['NGC4833'], C106: ['NGC104'], C107: ['NGC6101'], C108: ['NGC4372'],
  C109: ['NGC3195'],
};

function getSheetIdByTitle(spreadsheet: any, title: string): number {
  return spreadsheet.data.sheets?.find((s: any) => s.properties?.title === title)
    ?.properties?.sheetId ?? 0;
}

export async function createCaldwellChecklistSheet(spreadsheetId: string) {
  try {
    const sheets = createSheetsClient();
    const checklistSheetName = 'Caldwell Checklist';
    const mainSheetName = 'Astro Photo Log';

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheet = spreadsheet.data.sheets?.find(
      (s: any) => s.properties?.title === checklistSheetName
    );

    let sheetId: number;

    if (!existingSheet) {
      const batchUpdateResponse = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: checklistSheetName,
                  gridProperties: { rowCount: 200, columnCount: 11 },
                },
              },
            },
          ],
        },
      });
      sheetId = batchUpdateResponse.data.replies?.[0]?.addSheet?.properties?.sheetId ?? 0;
      if (sheetId === 0) {
        const refetch = await sheets.spreadsheets.get({ spreadsheetId });
        sheetId = getSheetIdByTitle(refetch, checklistSheetName);
      }
    } else {
      sheetId = existingSheet.properties?.sheetId ?? 0;
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: `${checklistSheetName}!A:K`,
      });
    }

    let logValues: any[][] = [];
    try {
      const logRes = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${mainSheetName}!A:Z`,
      });
      logValues = logRes.data.values || [];
    } catch (logErr: any) {
      console.warn(`⚠️ Could not read ${mainSheetName}; checklist will still be created without session history.`);
    }

    const headerRowIndex = findHeaderRow(logValues, true);
    const headers = logValues[headerRowIndex] || [];
    const logs = logValues.slice(headerRowIndex + 1);

    const { catalogColumnIdx, dateColumnIdx } = detectColumns(headers, logs);

    debugChecklistDetection(mainSheetName, headerRowIndex, headers, catalogColumnIdx, dateColumnIdx, logs);

    const rawCatalogLogs = buildCatalogLogs(logs, catalogColumnIdx, dateColumnIdx);

    // Extend with NGC/IC aliases: entries stored by NGC/IC number map to their Caldwell key
    const catalogLogs: Record<string, string[]> = { ...rawCatalogLogs };
    for (const [caldwellKey, aliases] of Object.entries(CALDWELL_NGC_ALIASES)) {
      for (const alias of aliases) {
        const entries = rawCatalogLogs[alias] ?? [];
        if (entries.length > 0) {
          catalogLogs[caldwellKey] = [...(catalogLogs[caldwellKey] ?? []), ...entries];
        }
      }
    }

    const caldwellObjects: Array<{ catalog: string; name: string }> = [];
    for (let i = 1; i <= 110; i++) {
      const key = `C${i}`;
      const obj = ASTRONOMICAL_OBJECTS[key];
      if (obj) {
        caldwellObjects.push({ catalog: key, name: obj.name });
      }
    }

    const col1 = caldwellObjects.slice(0, 37);
    const col2 = caldwellObjects.slice(37, 74);
    const col3 = caldwellObjects.slice(74, 110);

    const data: any[] = [
      ['Caldwell #', 'Object Name', 'Logs', '', 'Caldwell #', 'Object Name', 'Logs', '', 'Caldwell #', 'Object Name', 'Logs'],
    ];

    const maxRows = Math.max(col1.length, col2.length, col3.length);
    for (let i = 0; i < maxRows; i++) {
      const obj1 = col1[i];
      const obj2 = col2[i];
      const obj3 = col3[i];
      const row: any[] = [];

      if (obj1) {
        const entries1 = catalogLogs[normalizeCatalogId(obj1.catalog)] || [];
        row.push(obj1.catalog, obj1.name, entries1.length > 0 ? `${entries1.length} (${entries1.join(', ')})` : '-');
      } else {
        row.push('', '', '');
      }
      row.push('');

      if (obj2) {
        const entries2 = catalogLogs[normalizeCatalogId(obj2.catalog)] || [];
        row.push(obj2.catalog, obj2.name, entries2.length > 0 ? `${entries2.length} (${entries2.join(', ')})` : '-');
      } else {
        row.push('', '', '');
      }
      row.push('');

      if (obj3) {
        const entries3 = catalogLogs[normalizeCatalogId(obj3.catalog)] || [];
        row.push(obj3.catalog, obj3.name, entries3.length > 0 ? `${entries3.length} (${entries3.join(', ')})` : '-');
      } else {
        row.push('', '', '');
      }

      data.push(row);
    }

    const totalLogged = caldwellObjects.filter(obj => (catalogLogs[normalizeCatalogId(obj.catalog)] || []).length > 0).length;
    const progressPercent = Math.round((totalLogged / 110) * 100);

    data.push([]);
    data.push(['Summary', '', '', '', '', '', '', '', '', '', '']);
    data.push(['Total Caldwell Objects', 110, '', '', 'Logged', totalLogged, '', '', '', '', '']);
    data.push(['Progress', `${progressPercent}%`, '', '', 'Remaining', 110 - totalLogged, '', '', '', '', '']);

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${checklistSheetName}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: data },
    });

    const summaryStartRow = data.length - 3;
    const maxRowIndex = Math.min(data.length + 2, 200);

    try {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              repeatCell: {
                range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 11 },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: { red: 0.8, green: 0.2, blue: 0.2 },
                    textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 12 },
                    horizontalAlignment: 'CENTER',
                  },
                },
                fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
              },
            },
            {
              addConditionalFormatRule: {
                rule: {
                  ranges: [{ sheetId, startRowIndex: 1, endRowIndex: summaryStartRow, startColumnIndex: 0, endColumnIndex: 3 }],
                  booleanRule: {
                    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($C2<>"", $C2<>"-")' }] },
                    format: { backgroundColor: { red: 0.85, green: 1, blue: 0.8 } },
                  },
                },
                index: 0,
              },
            },
            {
              addConditionalFormatRule: {
                rule: {
                  ranges: [{ sheetId, startRowIndex: 1, endRowIndex: summaryStartRow, startColumnIndex: 4, endColumnIndex: 7 }],
                  booleanRule: {
                    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($G2<>"", $G2<>"-")' }] },
                    format: { backgroundColor: { red: 0.85, green: 1, blue: 0.8 } },
                  },
                },
                index: 1,
              },
            },
            {
              addConditionalFormatRule: {
                rule: {
                  ranges: [{ sheetId, startRowIndex: 1, endRowIndex: summaryStartRow, startColumnIndex: 8, endColumnIndex: 11 }],
                  booleanRule: {
                    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($K2<>"", $K2<>"-")' }] },
                    format: { backgroundColor: { red: 0.85, green: 1, blue: 0.8 } },
                  },
                },
                index: 2,
              },
            },
            {
              repeatCell: {
                range: { sheetId, startRowIndex: summaryStartRow - 1, endRowIndex: maxRowIndex, startColumnIndex: 0, endColumnIndex: 11 },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: { red: 0.95, green: 0.8, blue: 0.8 },
                    textFormat: { bold: true, fontSize: 11 },
                  },
                },
                fields: 'userEnteredFormat(backgroundColor,textFormat)',
              },
            },
          ],
        },
      });

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                fields: 'gridProperties.frozenRowCount',
                properties: { sheetId, gridProperties: { frozenRowCount: 1 } },
              },
            },
          ],
        },
      });
    } catch (formatErr: any) {
      console.warn('⚠️ Formatting update failed (non-critical):', formatErr.message);
    }

    try {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 200 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 220 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 200 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 220 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 9, endIndex: 10 }, properties: { pixelSize: 200 }, fields: 'pixelSize' } },
            { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 220 }, fields: 'pixelSize' } },
          ],
        },
      });
    } catch (widthErr: any) {
      console.warn('⚠️ Column width update failed (non-critical):', widthErr.message);
    }

    console.log(`✅ Caldwell Checklist updated: ${totalLogged}/110 objects logged (${progressPercent}%)`);
    return true;
  } catch (err: any) {
    console.error('❌ Failed to create Caldwell Checklist:', err.message);
    return false;
  }
}
