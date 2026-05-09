import { createSheetsClient } from './sheets.js';

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
          requests: [
            {
              addSheet: {
                properties: {
                  title: analyticsSheetName,
                  gridProperties: { rowCount: 80, columnCount: 8 },
                },
              },
            },
          ],
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
        range: `${analyticsSheetName}!A:K`,
      });
    }

    const logRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${mainSheetName}!A:Z`,
    });

    const logValues = logRes.data.values || [];
    const headers = logValues[0] || [];
    const logs = logValues.slice(1);

    let objectTypeColumnIdx = -1;
    for (let i = 0; i < headers.length; i++) {
      const headerLower = (headers[i] || '').toString().toLowerCase();
      if (headerLower.includes('object type')) {
        objectTypeColumnIdx = i;
        break;
      }
    }

    const typeCounts: Record<string, number> = {};
    if (objectTypeColumnIdx >= 0) {
      for (const row of logs) {
        const rawType = row[objectTypeColumnIdx]?.toString().trim();
        if (!rawType) continue;

        const type = rawType.replace(/\s*\([^)]*\)/g, '').trim();
        if (!type) continue;

        typeCounts[type] = (typeCounts[type] || 0) + 1;
      }
    }

    const typeEntries = Object.entries(typeCounts).sort((a, b) => b[1] - a[1]);
    const totalLogged = typeEntries.reduce((sum, [, count]) => sum + count, 0);

    const values: any[] = [
      ['Analytics Summary'],
      ['Updated', new Date().toISOString().slice(0, 10)],
      [],
      ['Object Type', 'Count'],
      ...typeEntries.map(([type, count]) => [type, count]),
      [],
      ['Total Logged Sessions', totalLogged],
      ['Distinct Object Types', typeEntries.length],
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${analyticsSheetName}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values },
    });

    const chartStartRow = 4;
    const chartEndRow = chartStartRow + typeEntries.length;

    const requests: any[] = [];
    if (chartDeletionRequests.length) {
      requests.push(...chartDeletionRequests);
    }

    requests.push(
      {
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: 2,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.15, green: 0.35, blue: 0.55 },
              textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 14 },
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat)',
        },
      },
      {
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: 3,
            endRowIndex: 4,
            startColumnIndex: 0,
            endColumnIndex: 2,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.8, green: 0.9, blue: 1 },
              textFormat: { bold: true, fontSize: 11 },
              horizontalAlignment: 'CENTER',
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
        },
      },
      {
        updateSheetProperties: {
          properties: { sheetId, gridProperties: { frozenRowCount: 4 } },
          fields: 'gridProperties.frozenRowCount',
        },
      },
      {
        updateDimensionProperties: {
          range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
          properties: { pixelSize: 220 },
          fields: 'pixelSize',
        },
      },
      {
        updateDimensionProperties: {
          range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
          properties: { pixelSize: 120 },
          fields: 'pixelSize',
        },
      }
    );

    if (typeEntries.length > 0) {
      requests.push({
        addChart: {
          chart: {
            spec: {
              title: 'Object Type Distribution',
              pieChart: {
                legendPosition: 'RIGHT_LEGEND',
                threeDimensional: false,
                domain: {
                  sourceRange: {
                    sources: [
                      {
                        sheetId,
                        startRowIndex: chartStartRow,
                        endRowIndex: chartEndRow,
                        startColumnIndex: 0,
                        endColumnIndex: 1,
                      },
                    ],
                  },
                },
                series: {
                  sourceRange: {
                    sources: [
                      {
                        sheetId,
                        startRowIndex: chartStartRow,
                        endRowIndex: chartEndRow,
                        startColumnIndex: 1,
                        endColumnIndex: 2,
                      },
                    ],
                  },
                },
              },
            },
            position: {
              overlayPosition: {
                anchorCell: {
                  sheetId,
                  rowIndex: 0,
                  columnIndex: 3,
                },
                offsetXPixels: 16,
                offsetYPixels: 16,
                widthPixels: 540,
                heightPixels: 340,
              },
            },
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
