import { createSheetsClient } from './sheets.js';
import { findHeaderRow } from './checklist-utils.js';

function parseIntegrationHours(value: string): number {
  const parts = value?.toString().split('/');
  return parts?.length === 2 ? (parseFloat(parts[1].trim()) || 0) / 3600 : 0;
}

function parseObjectCategory(rawType: string): string {
  return rawType?.replace(/\s*\([^)]*\)/g, '').trim() || 'Other';
}

export async function createAnalyticsSheet(spreadsheetId: string) {
  try {
    const sheets = createSheetsClient();
    const analyticsSheetName = 'Analytics';
    const mainSheetName = 'Astro Photo Log';

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheet = spreadsheet.data.sheets?.find((s: any) => s.properties?.title === analyticsSheetName);
    const chartDeletionRequests: any[] = [];
    let sheetId: number;

    if (!existingSheet) {
      const batchUpdateResponse = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{
            addSheet: {
              properties: {
                title: analyticsSheetName,
                gridProperties: { rowCount: 250, columnCount: 10 },
              },
            },
          }],
        },
      });
      sheetId = batchUpdateResponse.data.replies?.[0]?.addSheet?.properties?.sheetId ?? 0;
    } else {
      sheetId = existingSheet.properties?.sheetId ?? 0;
      if (Array.isArray(existingSheet.charts)) {
        existingSheet.charts.forEach((chart: any) => {
          if (chart.chartId) {
            chartDeletionRequests.push({ deleteEmbeddedObject: { objectId: chart.chartId } });
          }
        });
      }
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: `${analyticsSheetName}!A:J`,
      });
    }

    const logRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${mainSheetName}!A:Z`,
    });

    const logValues = logRes.data.values || [];
    const headerRowIndex = findHeaderRow(logValues);
    const headers = logValues[headerRowIndex] || [];
    const logs = logValues.slice(headerRowIndex + 1);

    // Detect all relevant columns in one pass
    let objectTypeColumnIdx = -1, dateColumnIdx = -1,
        integrationColumnIdx = -1, filterColumnIdx = -1, nameColumnIdx = -1;
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i]?.toString().toLowerCase() ?? '';
      if (h.includes('object type'))                          objectTypeColumnIdx  = i;
      if (h.includes('date'))                                 dateColumnIdx        = i;
      if (h.includes('integration'))                          integrationColumnIdx = i;
      if (h.includes('filter'))                               filterColumnIdx      = i;
      if (h.includes('target') || h.includes('object name')) nameColumnIdx        = i;
    }

    // Aggregations — single pass over all log rows
    const typeCounts:  Record<string, number> = {};
    const typeHours:   Record<string, number> = {};
    const monthCounts: Record<string, number> = {};
    const objectCounts: Record<string, number> = {};
    const filterCounts: Record<string, number> = {};
    let totalHours = 0;
    let firstDate = '', latestDate = '';

    for (const row of logs) {
      if (objectTypeColumnIdx >= 0) {
        const rawType = row[objectTypeColumnIdx]?.toString().trim() ?? '';
        if (rawType) {
          const cat = parseObjectCategory(rawType);
          typeCounts[cat] = (typeCounts[cat] || 0) + 1;

          if (integrationColumnIdx >= 0) {
            const hrs = parseIntegrationHours(row[integrationColumnIdx]?.toString() ?? '');
            typeHours[cat] = (typeHours[cat] || 0) + hrs;
            totalHours += hrs;
          }
        }
      }

      if (dateColumnIdx >= 0) {
        const date = row[dateColumnIdx]?.toString().trim() ?? '';
        const month = date.slice(0, 7);
        if (month) monthCounts[month] = (monthCounts[month] || 0) + 1;
        if (date && (!firstDate || date < firstDate)) firstDate = date;
        if (date && date > latestDate) latestDate = date;
      }

      if (nameColumnIdx >= 0) {
        const name = row[nameColumnIdx]?.toString().trim() ?? '';
        if (name) objectCounts[name] = (objectCounts[name] || 0) + 1;
      }

      if (filterColumnIdx >= 0) {
        const filter = row[filterColumnIdx]?.toString().trim() ?? '';
        if (filter) filterCounts[filter] = (filterCounts[filter] || 0) + 1;
      }
    }

    // Sort aggregations
    const typeEntries      = Object.entries(typeCounts) .sort((a, b) => b[1] - a[1]);
    const typeHoursEntries = Object.entries(typeHours)  .sort((a, b) => b[1] - a[1]);
    const monthEntries     = Object.entries(monthCounts).sort((a, b) => a[0].localeCompare(b[0]));
    const top15Objects     = Object.entries(objectCounts).sort((a, b) => b[1] - a[1]).slice(0, 15);
    const filterEntries    = Object.entries(filterCounts).sort((a, b) => b[1] - a[1]);

    const totalLogged    = typeEntries.reduce((sum, [, n]) => sum + n, 0);
    const distinctObjects = Object.keys(objectCounts).length;

    // ── Build values array, tracking 0-indexed row counter ──────────────────
    const values: any[][] = [];
    let r = 0;

    // Row 0: title
    values.push(['Imaging Log Analytics']);
    r++;

    // Row 1: summary stats across columns A–G
    values.push([
      `Total Sessions: ${totalLogged}`,
      '',
      `Total Integration: ${totalHours.toFixed(1)} hrs`,
      '',
      firstDate  ? `First: ${firstDate}`   : '',
      latestDate ? `Latest: ${latestDate}` : '',
      `Distinct objects: ${distinctObjects}`,
    ]);
    r++;

    // Row 2: spacer
    values.push([]);
    r++;

    // ── Block A: Object Type Distribution ──────────────────────────────────
    const blockA_section = r;
    values.push(['Object Type Distribution']);
    r++;
    const blockA_colHeader = r;
    values.push(['Object Type', 'Sessions']);
    r++;
    for (const [type, count] of typeEntries) { values.push([type, count]); r++; }
    const blockA_dataEnd = r;
    values.push([]);
    r++;

    // ── Block B: Integration Time by Object Type ───────────────────────────
    const blockB_section = r;
    values.push(['Total Integration Time (hrs)']);
    r++;
    const blockB_colHeader = r;
    values.push(['Object Type', 'Hours']);
    r++;
    for (const [type, hrs] of typeHoursEntries) { values.push([type, parseFloat(hrs.toFixed(2))]); r++; }
    const blockB_dataEnd = r;
    values.push([]);
    r++;

    // ── Block C: Sessions by Month ─────────────────────────────────────────
    const blockC_section = r;
    values.push(['Sessions by Month']);
    r++;
    const blockC_colHeader = r;
    values.push(['Month', 'Sessions']);
    r++;
    for (const [month, count] of monthEntries) { values.push([month, count]); r++; }
    const blockC_dataEnd = r;
    values.push([]);
    r++;

    // ── Block D: Top 15 Most-Imaged Objects (table only) ──────────────────
    const blockD_section = r;
    values.push(['Top 15 Most-Imaged Objects']);
    r++;
    values.push(['Object', 'Sessions']);
    r++;
    for (const [name, count] of top15Objects) { values.push([name, count]); r++; }
    values.push([]);
    r++;

    // ── Block E: Filter Usage ──────────────────────────────────────────────
    const blockE_section = r;
    values.push(['Filter Usage']);
    r++;
    const blockE_colHeader = r;
    values.push(['Filter', 'Sessions']);
    r++;
    for (const [filter, count] of filterEntries) { values.push([filter, count]); r++; }
    const blockE_dataEnd = r;
    values.push([]);
    r++;

    // Footer
    values.push(['Total Logged Sessions',  totalLogged]);
    values.push(['Total Integration Time', `${totalHours.toFixed(1)} hrs`]);

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${analyticsSheetName}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values },
    });

    // ── Formatting + charts ────────────────────────────────────────────────
    const requests: any[] = [];

    if (chartDeletionRequests.length) requests.push(...chartDeletionRequests);

    // Title row
    requests.push({
      repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 7 },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.15, green: 0.35, blue: 0.55 },
            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 14 },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });

    // Summary row
    requests.push({
      repeatCell: {
        range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 7 },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.85, green: 0.92, blue: 1.0 },
            textFormat: { bold: true, fontSize: 10 },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });

    // Section headers (gray bold)
    for (const sRow of [blockA_section, blockB_section, blockC_section, blockD_section, blockE_section]) {
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: sRow, endRowIndex: sRow + 1, startColumnIndex: 0, endColumnIndex: 3 },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.9, green: 0.9, blue: 0.9 },
              textFormat: { bold: true, fontSize: 11 },
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat)',
        },
      });
    }

    // Column header rows (light blue, centered)
    for (const chRow of [blockA_colHeader, blockB_colHeader, blockC_colHeader, blockE_colHeader]) {
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: chRow, endRowIndex: chRow + 1, startColumnIndex: 0, endColumnIndex: 3 },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.8, green: 0.9, blue: 1.0 },
              textFormat: { bold: true, fontSize: 11 },
              horizontalAlignment: 'CENTER',
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
        },
      });
    }

    requests.push(
      { updateSheetProperties: { properties: { sheetId, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } },
      { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 220 }, fields: 'pixelSize' } },
      { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },
      { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } },
    );

    const W = 460, H = 280;
    function anchorCell(rowIndex: number) {
      return { sheetId, rowIndex, columnIndex: 3 };
    }

    // Chart 1: Object Type Distribution (pie)
    if (typeEntries.length > 0) {
      requests.push({
        addChart: {
          chart: {
            spec: {
              title: 'Object Type Distribution',
              pieChart: {
                legendPosition: 'RIGHT_LEGEND',
                threeDimensional: false,
                domain: { sourceRange: { sources: [{ sheetId, startRowIndex: blockA_colHeader, endRowIndex: blockA_dataEnd, startColumnIndex: 0, endColumnIndex: 1 }] } },
                series: { sourceRange: { sources: [{ sheetId, startRowIndex: blockA_colHeader, endRowIndex: blockA_dataEnd, startColumnIndex: 1, endColumnIndex: 2 }] } },
              },
            },
            position: { overlayPosition: { anchorCell: anchorCell(blockA_colHeader), offsetXPixels: 16, offsetYPixels: 0, widthPixels: W, heightPixels: H } },
          },
        },
      });
    }

    // Chart 2: Integration Time by Object Type (bar)
    if (typeHoursEntries.length > 0) {
      requests.push({
        addChart: {
          chart: {
            spec: {
              title: 'Integration Time by Object Type',
              basicChart: {
                chartType: 'BAR',
                legendPosition: 'NO_LEGEND',
                axis: [
                  { position: 'BOTTOM_AXIS', title: 'Hours' },
                  { position: 'LEFT_AXIS',   title: 'Object Type' },
                ],
                domains: [{
                  domain: { sourceRange: { sources: [{ sheetId, startRowIndex: blockB_colHeader, endRowIndex: blockB_dataEnd, startColumnIndex: 0, endColumnIndex: 1 }] } },
                }],
                series: [{
                  series: { sourceRange: { sources: [{ sheetId, startRowIndex: blockB_colHeader, endRowIndex: blockB_dataEnd, startColumnIndex: 1, endColumnIndex: 2 }] } },
                  targetAxis: 'BOTTOM_AXIS',
                }],
                headerCount: 1,
              },
            },
            position: { overlayPosition: { anchorCell: anchorCell(blockB_colHeader), offsetXPixels: 16, offsetYPixels: 0, widthPixels: W, heightPixels: H } },
          },
        },
      });
    }

    // Chart 3: Sessions by Month (column)
    if (monthEntries.length > 0) {
      requests.push({
        addChart: {
          chart: {
            spec: {
              title: 'Sessions by Month',
              basicChart: {
                chartType: 'COLUMN',
                legendPosition: 'NO_LEGEND',
                axis: [
                  { position: 'BOTTOM_AXIS', title: 'Month' },
                  { position: 'LEFT_AXIS',   title: 'Sessions' },
                ],
                domains: [{
                  domain: { sourceRange: { sources: [{ sheetId, startRowIndex: blockC_colHeader, endRowIndex: blockC_dataEnd, startColumnIndex: 0, endColumnIndex: 1 }] } },
                }],
                series: [{
                  series: { sourceRange: { sources: [{ sheetId, startRowIndex: blockC_colHeader, endRowIndex: blockC_dataEnd, startColumnIndex: 1, endColumnIndex: 2 }] } },
                  targetAxis: 'LEFT_AXIS',
                }],
                headerCount: 1,
              },
            },
            position: { overlayPosition: { anchorCell: anchorCell(blockC_colHeader), offsetXPixels: 16, offsetYPixels: 0, widthPixels: W, heightPixels: H } },
          },
        },
      });
    }

    // Chart 4: Filter Usage (pie)
    if (filterEntries.length > 0) {
      requests.push({
        addChart: {
          chart: {
            spec: {
              title: 'Filter Usage',
              pieChart: {
                legendPosition: 'RIGHT_LEGEND',
                threeDimensional: false,
                domain: { sourceRange: { sources: [{ sheetId, startRowIndex: blockE_colHeader, endRowIndex: blockE_dataEnd, startColumnIndex: 0, endColumnIndex: 1 }] } },
                series: { sourceRange: { sources: [{ sheetId, startRowIndex: blockE_colHeader, endRowIndex: blockE_dataEnd, startColumnIndex: 1, endColumnIndex: 2 }] } },
              },
            },
            position: { overlayPosition: { anchorCell: anchorCell(blockE_colHeader), offsetXPixels: 16, offsetYPixels: 0, widthPixels: W, heightPixels: H } },
          },
        },
      });
    }

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });

    console.log('✅ Analytics tab created/updated');
    return true;
  } catch (err: any) {
    console.error('❌ Failed to create Analytics tab:', err.message);
    return false;
  }
}
