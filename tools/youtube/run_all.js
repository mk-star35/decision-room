#!/usr/bin/env node
// Runs all four pulls in order. Q2 depends on Q1's output, so order matters.
//
//   export YOUTUBE_API_KEY=...
//   node tools/youtube/run_all.js
//
// Total quota estimate: ~3,600 of the default 10,000 units/day.

import { spawnSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const steps = [
  ['q1_solitude_videos.js', '~205 units'],
  ['q2_comment_sentiment.js', '~30 units'],
  ['q3_new_philosophy_channels.js', '~1,450 units'],
  ['q4_east_west_crossref.js', '~1,850 units'],
];

if (!process.env.YOUTUBE_API_KEY && !process.env.YT_API_KEY) {
  console.error('YOUTUBE_API_KEY is not set. See tools/youtube/README.md');
  process.exit(1);
}

const failures = [];
for (const [script, cost] of steps) {
  console.log(`\n\n### ${script}  (${cost})`);
  const r = spawnSync(process.execPath, [join(here, script), ...process.argv.slice(2)], { stdio: 'inherit' });
  if (r.status !== 0) failures.push(script);
}

console.log(`\n\n${'='.repeat(72)}`);
if (failures.length) {
  console.log(`FAILED: ${failures.join(', ')}`);
  console.log(`Completed: ${steps.length - failures.length}/${steps.length}. Outputs for the successful steps are in tools/youtube/out/.`);
  process.exit(1);
}
console.log(`All ${steps.length} pulls complete. Outputs in tools/youtube/out/.`);
