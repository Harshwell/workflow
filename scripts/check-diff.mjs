import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import process from 'node:process';
import { scanAddedLines, validateDocumentationDrift } from './lib/diff-validation.mjs';

const base = process.env.DIFF_BASE?.trim();
const head = process.env.DIFF_HEAD?.trim() || 'HEAD';
const docsOnly = process.argv.includes('--docs-only');
const sensitiveOnly = process.argv.includes('--sensitive-only');
const rangeArgs = base && !/^0+$/.test(base) ? [`${base}..${head}`] : [];
const changed = git(['diff', '--name-only', '--diff-filter=ACDMRTUXB', ...rangeArgs]);
const diff = git(['diff', '--unified=0', '--no-ext-diff', ...rangeArgs]);

let changedFiles = lines(changed);
const records = parseAddedLines(diff);

if (!rangeArgs.length) {
  const untracked = lines(git(['ls-files', '--others', '--exclude-standard']));
  changedFiles = [...new Set([...changedFiles, ...untracked])];
  for (const file of untracked) {
    if (!fs.existsSync(file) || fs.statSync(file).isDirectory()) continue;
    records.push({
      file,
      lines: fs.readFileSync(file, 'utf8').split(/\r?\n/).map((text, index) => ({ line: index + 1, text }))
    });
  }
}

const errors = [];
if (!sensitiveOnly) errors.push(...validateDocumentationDrift(changedFiles, process.env.PR_BODY || ''));
if (!docsOnly) errors.push(...scanAddedLines(records));

if (errors.length) {
  process.stderr.write(`${errors.join('\n')}\n`);
  process.exit(1);
}
process.stdout.write(`Diff governance OK: ${changedFiles.length} changed file(s)\n`);

export function parseAddedLines(diffText) {
  const records = [];
  let current = null;
  let newLine = 0;
  for (const line of String(diffText || '').split(/\r?\n/)) {
    const fileMatch = /^\+\+\+ b\/(.+)$/.exec(line);
    if (fileMatch) {
      current = { file: fileMatch[1], lines: [] };
      records.push(current);
      continue;
    }
    const hunk = /^@@ -\d+(?:,\d+)? \+(\d+)(?:,\d+)? @@/.exec(line);
    if (hunk) {
      newLine = Number(hunk[1]);
      continue;
    }
    if (!current || line.startsWith('diff --git') || line.startsWith('--- ')) continue;
    if (line.startsWith('+')) {
      current.lines.push({ line: newLine, text: line.slice(1) });
      newLine += 1;
    } else if (!line.startsWith('-')) {
      newLine += 1;
    }
  }
  return records;
}

function git(args) {
  return execFileSync('git', args, { encoding: 'utf8' });
}

function lines(value) {
  return String(value || '').split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
}
