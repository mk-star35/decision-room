import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const OUT_DIR = join(dirname(fileURLToPath(import.meta.url)), '..', 'out');

export function writeJson(name, data) {
  mkdirSync(OUT_DIR, { recursive: true });
  const path = join(OUT_DIR, name);
  writeFileSync(path, JSON.stringify(data, null, 2), 'utf8');
  return path;
}

export function writeCsv(name, rows, columns) {
  mkdirSync(OUT_DIR, { recursive: true });
  const cols = columns ?? [...new Set(rows.flatMap((r) => Object.keys(r)))];
  const esc = (v) => {
    if (v === undefined || v === null) return '';
    const s = Array.isArray(v) ? v.join('|') : String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const csv = [cols.join(','), ...rows.map((r) => cols.map((c) => esc(r[c])).join(','))].join('\n');
  const path = join(OUT_DIR, name);
  // BOM so Excel opens the Korean text as UTF-8 instead of mojibake.
  writeFileSync(path, `﻿${csv}`, 'utf8');
  return path;
}

export function writeJsonl(name, rows) {
  mkdirSync(OUT_DIR, { recursive: true });
  const path = join(OUT_DIR, name);
  writeFileSync(path, rows.map((r) => JSON.stringify(r)).join('\n'), 'utf8');
  return path;
}

export function getKey() {
  const key = process.env.YOUTUBE_API_KEY || process.env.YT_API_KEY;
  if (!key) {
    console.error(
      'YOUTUBE_API_KEY is not set.\n' +
        'Get one at https://console.cloud.google.com/apis/credentials (enable "YouTube Data API v3"),\n' +
        'then run:  export YOUTUBE_API_KEY=...'
    );
    process.exit(1);
  }
  return key;
}

export function arg(flag, fallback) {
  const i = process.argv.indexOf(`--${flag}`);
  if (i === -1) return fallback;
  const v = process.argv[i + 1];
  return v && !v.startsWith('--') ? v : true;
}

export function banner(title) {
  console.log(`\n${'='.repeat(72)}\n${title}\n${'='.repeat(72)}`);
}
