import fs from 'node:fs';
import path from 'path';

const ROOT = process.cwd();

function loadEnvFile(envPath: string): Record<string, string> {
  const env: Record<string, string> = {};
  if (!fs.existsSync(envPath)) return env;
  const contents = fs.readFileSync(envPath, 'utf-8');
  for (const line of contents.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const [key, ...rest] = trimmed.split('=');
    if (!key) continue;
    env[key.trim()] = rest.join('=').trim();
  }
  return env;
}

const env = loadEnvFile(path.join(ROOT, '.env'));

function resolvePath(value: string | undefined, fallback: string): string {
  if (!value) return fallback;
  return path.isAbsolute(value) ? value : path.join(ROOT, value);
}

export const CONFIG_FILE = resolvePath(
  process.env.CONFIG_FILE,
  path.join(ROOT, 'seestar-log-config.json')
);
export const PROCESSED_FILE = resolvePath(
  process.env.PROCESSED_FILE,
  path.join(ROOT, 'processed-files.json')
);
export const DEFAULT_OUTPUT = resolvePath(
  process.env.DEFAULT_OUTPUT,
  path.join(ROOT, 'seestar-imaging-log.csv')
);
export const SERVICE_ACCOUNT_FILE = resolvePath(
  process.env.SERVICE_ACCOUNT_FILE,
  path.join(ROOT, 'service-account.json')
);
export const DEFAULT_MAX_GAP_MINUTES = Number(process.env.DEFAULT_MAX_GAP_MINUTES) || 30;
export const DEFAULT_MIN_FILES = Number(process.env.DEFAULT_MIN_FILES) || 2;
