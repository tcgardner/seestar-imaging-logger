import { createSheetsClient } from './sheets.js';
import { ASTRONOMICAL_OBJECTS } from '../astronomical-objects.js';
import {
  findHeaderRow,
  detectColumns,
  debugChecklistDetection,
  buildCatalogLogs,
  normalizeCatalogId,
} from './checklist-utils.js';

function getSheetIdByTitle(spreadsheet: any, title: string): number {
  return spreadsheet.data.sheets?.find((s: any) => s.properties?.title === title)
    ?.properties?.sheetId ?? 0;
}

export async function createMessierChecklistSheet(spreadsheetId: string) {
  try {
    const sheets = createSheetsClient();
    const checklistSheetName = 'Messier Checklist';
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

    const catalogLogs = buildCatalogLogs(logs, catalogColumnIdx, dateColumnIdx);

    const messierObjects: Array<{ catalog: string; name: string }> = [];
    for (let i = 1; i <= 110; i++) {
      const key = `M${i}`;
      const obj = ASTRONOMICAL_OBJECTS[key];
      if (obj) {
        messierObjects.push({ catalog: key, name: obj.name });
      }
    }

    const col1 = messierObjects.slice(0, 37);
    const col2 = messierObjects.slice(37, 74);
    const col3 = messierObjects.slice(74, 110);

    const data: any[] = [
      ['Messier #', 'Object Name', 'Logs', '', 'Messier #', 'Object Name', 'Logs', '', 'Messier #', 'Object Name', 'Logs'],
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

    const totalLogged = messierObjects.filter(obj => (catalogLogs[normalizeCatalogId(obj.catalog)] || []).length > 0).length;
    const progressPercent = Math.round((totalLogged / 110) * 100);

    data.push([]);
    data.push(['Summary', '', '', '', '', '', '', '', '', '', '']);
    data.push(['Total Messier Objects', 110, '', '', 'Logged', totalLogged, '', '', '', '', '']);
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
                    backgroundColor: { red: 0.2, green: 0.4, blue: 0.8 },
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
                    backgroundColor: { red: 0.95, green: 0.95, blue: 0.8 },
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

    console.log(`✅ Messier Checklist updated: ${totalLogged}/110 objects logged (${progressPercent}%)`);
    return true;
  } catch (err: any) {
    console.error('❌ Failed to create Messier Checklist:', err.message);
    return false;
  }
}
