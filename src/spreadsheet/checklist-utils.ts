export function normalizeCatalogId(value: string): string {
  return value
    .toString()
    .toUpperCase()
    .trim()
    .replace(/[^A-Z0-9]/g, '');
}

function isLikelyDateValue(value: any): boolean {
  const text = value?.toString().trim();
  if (!text) return false;
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return true;
  const parsed = new Date(text);
  return !Number.isNaN(parsed.getTime());
}

export function findHeaderRow(values: any[][], debug = false): number {
  const datePattern = /\bdate\b/i;
  const catalogPattern = /\bcat(alog)?\b/i;
  let bestRowIndex = 0;
  let bestScore = -1;
  const candidates: Array<{
    rowIndex: number;
    score: number;
    foundDate: boolean;
    foundCatalog: boolean;
    cells: string[];
  }> = [];

  for (let rowIndex = 0; rowIndex < Math.min(values.length, 10); rowIndex++) {
    const row = values[rowIndex];
    if (!Array.isArray(row)) continue;

    let score = 0;
    let foundDate = false;
    let foundCatalog = false;
    const cells: string[] = [];

    for (const cell of row) {
      const text = cell?.toString().trim() || '';
      cells.push(text || '<empty>');
      if (!text) continue;
      if (datePattern.test(text)) {
        score += 3;
        foundDate = true;
      }
      if (catalogPattern.test(text)) {
        score += 3;
        foundCatalog = true;
      }
      if (text.toLowerCase().includes('yyyy')) score += 1;
      if (text.includes('#')) score += 1;
    }

    candidates.push({ rowIndex, score, foundDate, foundCatalog, cells });

    if (score > bestScore) {
      bestScore = score;
      bestRowIndex = rowIndex;
    }

    if (foundDate && foundCatalog) break;
  }

  /*   if (debug) {
    console.log('🔎 findHeaderRow diagnostics for Astro Photo Log');
    console.log(`  scanned ${Math.min(values.length, 10)} rows`);
    candidates.forEach((c) => {
      console.log(`  row ${c.rowIndex}: score=${c.score} date=${c.foundDate} catalog=${c.foundCatalog} cells=[${c.cells.join(' | ')}]`);
    });
    console.log(`  selected header row ${bestRowIndex} (bestScore=${bestScore})`);
  }
 */
  return bestScore >= 1 ? bestRowIndex : 0;
}

export function detectColumns(
  headers: any[],
  sampleRows?: any[][]
): { catalogColumnIdx: number; dateColumnIdx: number } {
  let catalogColumnIdx = -1;
  let dateColumnIdx = -1;

  // Prefer exact match for catalog column first
  for (let i = 0; i < headers.length; i++) {
    const text = headers[i]?.toString().trim() ?? '';
    if (text === 'Catalog #') {
      catalogColumnIdx = i;
      break;
    }
  }

  // Prefer exact match for date column first
  for (let i = 0; i < headers.length; i++) {
    const text = headers[i]?.toString().trim() ?? '';
    if (text === 'Date (YYYY-MM-DD)') {
      dateColumnIdx = i;
      break;
    }
  }

  // Fall back to fuzzy matching for catalog if not found
  if (catalogColumnIdx < 0) {
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i]?.toString().toLowerCase() ?? '';
      if (h.includes('catalog') || h.includes('cat #') || h.includes('cat')) {
        catalogColumnIdx = i;
        break;
      }
    }
  }

  // Fall back to fuzzy matching for date if not found
  if (dateColumnIdx < 0) {
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i]?.toString().toLowerCase() ?? '';
      if (h.includes('date')) {
        dateColumnIdx = i;
        break;
      }
    }
  }

  // Last resort: scan data rows to find a column with date-like values and an empty header
  if (dateColumnIdx < 0 && sampleRows && sampleRows.length > 0) {
    const sampleCount = Math.min(5, sampleRows.length);
    for (let i = 0; i < headers.length; i++) {
      const headerText = headers[i]?.toString().trim() ?? '';
      if (headerText) continue; // only consider columns with empty/missing headers
      const dateMatches = sampleRows.slice(0, sampleCount).filter(row => isLikelyDateValue(row[i]));
      if (dateMatches.length >= Math.max(2, sampleCount - 1)) {
        dateColumnIdx = i;
        //console.log(`🔧 Fallback: detected date column at index ${i} from data values (header was empty)`);
        break;
      }
    }
  }

  return { catalogColumnIdx, dateColumnIdx };
}

export function debugChecklistDetection(
  mainSheetName: string,
  headerRowIndex: number,
  headers: any[],
  catalogColumnIdx: number,
  dateColumnIdx: number,
  sampleRows: any[][]
) {
  /*   console.log(`📋 ${mainSheetName} checklist debug:`);
  console.log(`   headerRowIndex=${headerRowIndex}`);
  console.log(`   headers=${headers.map((h) => h?.toString().trim() || '<empty>').join(' | ')}`);
  console.log(`   matched catalogColumnIdx=${catalogColumnIdx}, dateColumnIdx=${dateColumnIdx}`);
  if (catalogColumnIdx < 0 || dateColumnIdx < 0) {
    console.warn('   ❗ Could not match required Catalog/Date columns');
  }
  console.log('   first 5 rows after header:');
  sampleRows.slice(0, 5).forEach((row, idx) => {
    console.log(`     row ${idx}: ${row.map((cell: any) => cell?.toString().trim() || '<empty>').slice(0, 10).join(' | ')}`);
  }); */
}

export function buildCatalogLogs(
  logs: any[][],
  catalogColumnIdx: number,
  dateColumnIdx: number
): Record<string, string[]> {
  const catalogLogs: Record<string, string[]> = {};

  if (catalogColumnIdx < 0 || dateColumnIdx < 0) {
    console.warn(
      '⚠️ Could not detect Catalog or Date columns in Astro Photo Log; checklist will remain empty.'
    );
    return catalogLogs;
  }

  for (const log of logs) {
    const rawCatalog = log[catalogColumnIdx]?.toString() ?? '';
    // Fix: use ?. on trim() to avoid TypeError when value is undefined/null
    const date = log[dateColumnIdx]?.toString()?.trim() ?? '';
    const catalog = normalizeCatalogId(rawCatalog);
    if (catalog && date) {
      if (!catalogLogs[catalog]) {
        catalogLogs[catalog] = [];
      }
      catalogLogs[catalog].push(date);
    }
  }

  return catalogLogs;
}
