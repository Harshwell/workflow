import { spawnSync } from 'node:child_process';
import path from 'node:path';
import process from 'node:process';

const files = [
  'optional-project/Outstanding',
  'optional-project/salvage',
  'optional-project/SC-Meilani.js',
  'optional-project/Service Center Extractor.js'
];

for (const file of files) {
  const result = spawnSync(process.execPath, ['--check', path.resolve(file)], {
    encoding: 'utf8'
  });
  if (result.status !== 0) {
    process.stderr.write(result.stderr || result.stdout || `Syntax check failed: ${file}\n`);
    process.exit(result.status || 1);
  }
  process.stdout.write(`Syntax OK: ${file}\n`);
}
