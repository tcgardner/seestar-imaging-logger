import fs, { existsSync, readFileSync, writeFileSync, appendFileSync } from 'node:fs';
import fs2 from 'fs/promises';
import path from 'path';
import { FITS } from 'fitsjs-ng';
import fg from 'fast-glob';
import { stringify } from 'csv-stringify/sync';
import { program } from 'commander';
import { getStandardHeader, parseFitsHeaders, readFitsHeadersFromFile } from './fits-header-parser.js';
import { ASTRONOMICAL_OBJECTS } from './astronomical-objects.js';
import { getObjectTypeFromCatalog, getObjectNameFromCatalog } from './astronomy-api.js';
import {
  createMessierChecklistSheet,
  createCaldwellChecklistSheet,
  createAnalyticsSheet,
  isSheetEmpty,
  appendToGoogleSheet,
} from './spreadsheet/index.js';
import { CONFIG_FILE, PROCESSED_FILE, DEFAULT_OUTPUT, DEFAULT_MAX_GAP_MINUTES, DEFAULT_MIN_FILES } from './env.js';

let processedFiles: Set<string> = new Set();

let config: {
  lastDirectory?: string;
  lastOutput: string;
  lastSpreadsheetId?: string;
  excludePatterns: string[];
  minFiles: number;
} = {
  lastOutput: DEFAULT_OUTPUT,
  excludePatterns: ['calibration', 'darks', 'flats', 'bias', 'thumbs'],
  minFiles: DEFAULT_MIN_FILES,
};

async function loadConfig() {
  if (existsSync(CONFIG_FILE)) {
    try {
      const data = readFileSync(CONFIG_FILE, 'utf-8');
      config = { ...config, ...JSON.parse(data) };
    } catch (e) {
      console.warn('⚠️ Could not load config, using defaults');
    }
  }
}

async function saveConfig(
  currentDir: string,
  currentOutput: string,
  spreadsheetId: string | undefined,
  excludes: string[],
  minFiles: number
) {
  config.lastDirectory = currentDir || config.lastDirectory;
  config.lastOutput = currentOutput;
  if (spreadsheetId) config.lastSpreadsheetId = spreadsheetId;
  config.excludePatterns = [...new Set(excludes)];
  config.minFiles = minFiles;
  writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2));
}

async function loadProcessedFiles() {
  if (existsSync(PROCESSED_FILE)) {
    const data = readFileSync(PROCESSED_FILE, 'utf-8');
    processedFiles = new Set(JSON.parse(data));
  }
}

async function saveProcessedFiles() {
  writeFileSync(PROCESSED_FILE, JSON.stringify(Array.from(processedFiles), null, 2));
}


export interface RawParsedFile {
  filePath: string;
  timestamp: Date;
  dateStr: string;
  objectName: string;
  objectType: string;
  catalog: string;
  expTimePerSub: string;
  numSubs: number;
  totalExp: number;
  filter: string;
  gain: string;
  ra: string;
  dec: string;
  telescop: string;
  notes: string;
}

export interface LogEntry {
  'Date (YYYY-MM-DD)': string;
  'Target/Object Name': string;
  'Object Type': string;
  'Catalog #': string;
  'Exposure Time per Sub (seconds)': string;
  'Number of Subs / Total Integration Time (seconds)': string;
  'Filter Used': string;
  'Gain / ISO': string;
  'RA': string;
  'DEC': string;
  'TELESCOP': string;
  'FITS File Path': string;
  'Notes / Issues': string;
}

async function loadFile(filename: string): Promise<ArrayBuffer> {
  const buffer = await fs2.readFile(filename);
  return buffer.buffer;
}

async function parseFITS(filePath: string): Promise<RawParsedFile | null> {
  try {   
    const hdus = parseFitsHeaders(await loadFile(filePath));    
    if (!hdus) {
      throw new Error('Failed to parse FITS header');
    }
    const primary = hdus[0];

    const std = getStandardHeader(primary);
    
    const dateObs = std.DATE_OBS || std.DATE || '';
    let timestamp = new Date();
    let dateStr = 'Unknown';

    if (dateObs) {
      // === ROBUST UTC → LOCAL CONVERSION ===
      // FITS DATE-OBS is always in UTC. We force UTC parsing, then extract
      // local date + local time using the machine's timezone.
      let utcStr = dateObs.toString().trim();

      // Normalize separator (some headers use space instead of T)
      if (!utcStr.includes('T')) {
        utcStr = utcStr.replace(' ', 'T');
      }

      // If no timezone/offset is present, treat as UTC
      if (!/Z|[+-]\d{2}/.test(utcStr)) {
        utcStr += 'Z';
      }

      timestamp = new Date(utcStr);

      // Fallback if parsing failed (very rare)
      if (isNaN(timestamp.getTime())) {
        timestamp = new Date(dateObs);
        console.warn(`⚠️  Date parsing fallback used for ${path.basename(filePath)}`);
      }

      // ⏰ EXPLICIT LOCAL TIMEZONE CONVERSION ⏰
      // Get the system's local timezone
      const userTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;

      // Use Intl to get local date in YYYY-MM-DD format
      const localDateFormatter = new Intl.DateTimeFormat('en-CA', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        timeZone: userTimeZone,
      });
      dateStr = localDateFormatter.format(timestamp);
    }

    const objectName = (std.OBJECT || path.basename(filePath, path.extname(filePath))).toString().trim();
    const catalogStr = (std.CATALOG || std.OBJECT || '').toString().trim();

    const numSubs = std.STACKCNT || 1;
    const totalExp = parseFloat(std.TOTALEXP?.toString() || std.EXPTIME?.toString() || '0') || 0;
    const expTimePerSub = String(std.EXPTIME || std.EXPOSURE || 'N/A');

    // Live lookup via Sesame CDS Name Resolver
    const objectType = await getObjectTypeFromCatalog(catalogStr);
    const canonicalName = getObjectNameFromCatalog(catalogStr);
    
    // Use canonical name if available, otherwise use the header name
    const finalObjectName = canonicalName || objectName;

    return {
      filePath,
      timestamp,
      dateStr,
      objectName: finalObjectName,
      objectType,
      catalog: catalogStr,
      expTimePerSub,
      numSubs,
      totalExp,
      filter: (std.FILTER || 'None').toString(),
      gain: std.GAIN ? String(std.GAIN) : 'Auto',
      ra: (std.RA || std.CRVAL1 || '').toString(),
      dec: (std.DEC || std.CRVAL2 || '').toString(),
      telescop: (std.TELESCOP || std.INSTRUME || 'Seestar S30 Pro').toString(),
      notes: `Header parsed | Size: ${std.NAXIS1 || '?'}x${std.NAXIS2 || '?'}`,
    };
  } catch (err: any) {
    console.error(`❌ Failed to parse ${path.basename(filePath)}:`, err.message);
    return null;
  }
}

export function groupIntoSessions(rawFiles: RawParsedFile[], minFiles: number): LogEntry[] {
  if (rawFiles.length === 0) return [];

  const sorted = [...rawFiles].sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());

  const sessions: LogEntry[] = [];
  let currentSession: RawParsedFile[] = [sorted[0]];

  for (let i = 1; i < sorted.length; i++) {
    const prev = currentSession[currentSession.length - 1];
    const curr = sorted[i];

    const sameDate = prev.dateStr === curr.dateStr;
    const sameObject = prev.objectName.toLowerCase() === curr.objectName.toLowerCase();
    const timeGapMs = curr.timestamp.getTime() - prev.timestamp.getTime();
    const timeGapMinutes = timeGapMs / (1000 * 60);

    if (sameDate && sameObject && timeGapMinutes <= DEFAULT_MAX_GAP_MINUTES) {
      currentSession.push(curr);
    } else {
      if (currentSession.length >= minFiles) {
        sessions.push(createSessionEntry(currentSession));
      } else {
        console.log(`⏭️  Skipped session with only ${currentSession.length} file(s): ${currentSession[0].objectName} (${currentSession[0].dateStr})`);
      }
      currentSession = [curr];
    }
  }

  if (currentSession.length >= minFiles) {
    sessions.push(createSessionEntry(currentSession));
  } else if (currentSession.length > 0) {
    console.log(`⏭️  Skipped session with only ${currentSession.length} file(s): ${currentSession[0].objectName} (${currentSession[0].dateStr})`);
  }

  return sessions;
}

export function createSessionEntry(group: RawParsedFile[]): LogEntry {
  const first = group[0];
  const last = group[group.length - 1];

  const totalSubs = ((group.reduce((sum, f) => sum + f.numSubs, 0)) / 2) + 1; // Adjusted to reflect actual subs (since some headers report total, some per sub)
  const totalIntegration = ((group.reduce((sum, f) => sum + f.totalExp, 0)) / 2) + (parseFloat(group[0].expTimePerSub) || 0); // Adjusted to reflect actual integration time

  const expPerSub = group.every(f => f.expTimePerSub === first.expTimePerSub) ? first.expTimePerSub : 'various';
  const filterUsed = group.every(f => f.filter === first.filter) ? first.filter : 'mixed';

  const notes = `${group.length} files grouped | Gap ≤ ${DEFAULT_MAX_GAP_MINUTES} min`;

  return {
    'Date (YYYY-MM-DD)': first.dateStr,
    'Target/Object Name': first.objectName,
    'Object Type': first.objectType,
    'Catalog #': first.catalog,
    'Exposure Time per Sub (seconds)': expPerSub,
    'Number of Subs / Total Integration Time (seconds)': `${totalSubs} / ${totalIntegration.toFixed(0)}`,
    'Filter Used': filterUsed,
    'Gain / ISO': first.gain,
    'RA': first.ra,
    'DEC': first.dec,
    'TELESCOP': first.telescop.split('_')[0], // Just the telescope name, without extra details
    'FITS File Path': first.filePath,
    'Notes / Issues': notes,
  };
}

async function main() {
  console.log('🔭 Seestar S30 Pro Imaging Log - Minimum files per session enabled');
  console.log('   🌐 Sesame CDS lookup active ');

  await loadConfig();
  await loadProcessedFiles();

  program
    .name('seestar-log')
    .description('Seestar S30 Pro FITS → Sessions → CSV + Google Sheets (Sesame CDS)')
    .argument('[directory]', 'Root folder (optional – uses saved config)', config.lastDirectory)
    .option('-o, --output <path>', 'CSV file', config.lastOutput)
    .option('-g, --google-sheet <id>', 'Google Spreadsheet ID', config.lastSpreadsheetId)
    .option('-e, --exclude <patterns>', 'Comma-separated folders to skip', config.excludePatterns.join(','))
    .option('--min-files <number>', 'Minimum number of FITS files to form a session', String(config.minFiles))
    .option('--reset', 'Reset saved config')
    .action(async (directory: string | undefined, options: any) => {
      if (options.reset) {
        config = { 
          lastOutput: DEFAULT_OUTPUT, 
          excludePatterns: ['calibration', 'darks', 'flats', 'bias', 'thumbs'],
          minFiles: DEFAULT_MIN_FILES 
        };
        await saveConfig('', DEFAULT_OUTPUT, undefined, config.excludePatterns, DEFAULT_MIN_FILES);
        console.log('🔄 Config reset to defaults.');
        return;
      }

      const rootDir = directory || config.lastDirectory;
      if (!rootDir) {
        console.error('❌ Please provide a directory.');
        process.exit(1);
      }

      const excludeList = options.exclude 
        ? options.exclude.split(',').map((s: string) => s.trim()).filter(Boolean) 
        : config.excludePatterns;

      const minFiles = parseInt(options.minFiles || config.minFiles.toString()) || DEFAULT_MIN_FILES;
      const spreadsheetId = options.googleSheet || config.lastSpreadsheetId;

      console.log(`🔍 Scanning ${rootDir}...`);
      console.log(`⛔ Skipping: ${excludeList.join(', ') || 'none'}`);
      console.log(`📏 Minimum files per session: ${minFiles}`);
      if (spreadsheetId) console.log(`📊 Google Sheets: ${spreadsheetId}`);

      const ignorePatterns = excludeList.map(folder => `**/${folder}/**`);
      const files = await fg('**/*.{fit,fits}', { 
        cwd: rootDir, 
        absolute: true, 
        onlyFiles: true, 
        ignore: ignorePatterns 
      });

      const newFiles = files.filter(f => !processedFiles.has(f));
      console.log(`📊 ${files.length} total FITS files, ${newFiles.length} new.\n`);

      if (newFiles.length === 0) {
        console.log('✅ No new files.');
        // Still update checklists even with no new files
        if (spreadsheetId) {
          await createMessierChecklistSheet(spreadsheetId);
          await createCaldwellChecklistSheet(spreadsheetId);
          await createAnalyticsSheet(spreadsheetId);
        }
        await saveConfig(rootDir, options.output, spreadsheetId, excludeList, minFiles);
        return;
      }

      const rawParsed: RawParsedFile[] = [];
      let processedCount = 0;

      for (const file of newFiles) {
        processedCount++;
        const percent = Math.round((processedCount / newFiles.length) * 100);
        const bar = '█'.repeat(Math.floor(percent / 5)) + '░'.repeat(20 - Math.floor(percent / 5));

        process.stdout.write(`\r\x1b[2K⏳ [${bar}] ${processedCount}/${newFiles.length} (${percent}%) – ${path.basename(file)}`);

        const parsed = await parseFITS(file);
        if (parsed) {
          rawParsed.push(parsed);
          processedFiles.add(file);
        }
      }
      console.log('');

      const sessionEntries = groupIntoSessions(rawParsed, minFiles);
      console.log(`📦 Created ${sessionEntries.length} valid imaging session(s)`);

      if (sessionEntries.length === 0) {
        console.log('⚠️ No sessions met the minimum file requirement.');
        // Still update checklists even with no sessions
        if (spreadsheetId) {
          await createMessierChecklistSheet(spreadsheetId);
          await createCaldwellChecklistSheet(spreadsheetId);
          await createAnalyticsSheet(spreadsheetId);
        }
        await saveProcessedFiles();
        await saveConfig(rootDir, options.output, spreadsheetId, excludeList, minFiles);
        return;
      }

      const csvString = stringify(sessionEntries, {
        header: !existsSync(options.output),
        columns: Object.keys(sessionEntries[0]) as any[],
        quoted: true,
      });

      if (existsSync(options.output)) {
        appendFileSync(options.output, csvString);
      } else {
        await fs2.writeFile(options.output, csvString);
      }

      if (spreadsheetId && sessionEntries.length > 0) {
        const isFirstTime = await isSheetEmpty(spreadsheetId);
        await appendToGoogleSheet(spreadsheetId, sessionEntries, isFirstTime);
      }

      // Always update checklists if spreadsheet is configured
      if (spreadsheetId) {
        await createMessierChecklistSheet(spreadsheetId);
        await createCaldwellChecklistSheet(spreadsheetId);
        await createAnalyticsSheet(spreadsheetId);
      }

      await saveProcessedFiles();
      await saveConfig(rootDir, options.output, spreadsheetId, excludeList, minFiles);

      console.log('\n🎉 Complete!');
      console.log(`   CSV → ${options.output}`);
      if (spreadsheetId) console.log(`   Google Sheet → ${spreadsheetId}`);
      console.log(`   New sessions added: ${sessionEntries.length}`);
    });

  await program.parseAsync(process.argv);
}

main().catch((err) => {
  console.error('💥 Error:', err.message);
  process.exit(1);
});