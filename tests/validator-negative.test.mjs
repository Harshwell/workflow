import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  validateLocalLinks,
  validateMarkdownPolicy
} from '../scripts/lib/docs-validation.mjs';
import {
  scanAddedLines,
  validateDocumentationDrift
} from '../scripts/lib/diff-validation.mjs';
import { validateCriticalMappings } from '../scripts/lib/mapping-contracts.mjs';
import { loadSources } from './mapping-contracts.test.mjs';

test('broken local links fail documentation validation', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'workflow-docs-'));
  const markdown = '# Fixture\n\n[Missing](missing.md)\n';
  const errors = validateLocalLinks(root, 'README.md', markdown, new Map([['README.md', markdown]]));
  assert.equal(errors.length, 1);
  assert.match(errors[0], /broken local link/);
});

test('timeline-style headings fail documentation policy', () => {
  const errors = validateMarkdownPolicy('README.md', '# Handbook\n\n## Timeline\n\n## Apps Script UAT\n');
  assert.ok(errors.some((error) => error.includes('prohibited historical')));
});

test('critical mapping drift fails the contract validator', () => {
  const sources = loadSources();
  sources.config = sources.config.replace(/('Carlcare',\r?\n\s*)'GSI'/, "$1'GSI-DRIFT'");
  const errors = validateCriticalMappings(sources);
  assert.ok(errors.some((error) => error.includes('root GSI -> Meilani')));
});

test('business code without changelog fails documentation drift', () => {
  const errors = validateDocumentationDrift(['05b_Pipeline_RoutingOperational.gs', 'README.md'], '');
  assert.ok(errors.some((error) => error.includes('CHANGELOG.md')));
});

test('empty or generic documentation escape hatch fails', () => {
  const changed = ['00_Config.gs', 'CHANGELOG.md'];
  assert.ok(validateDocumentationDrift(changed, 'Docs-Impact-README: none - no impact').length > 0);
  assert.deepEqual(
    validateDocumentationDrift(changed, 'Docs-Impact-README: none - Only comment punctuation changed; documented contracts are identical.'),
    []
  );
});

test('new secret and Google identifier fail sensitive diff validation', () => {
  const credential = ['api', 'key'].join('_') + ' = ' + '"' + 'live_' + 'x'.repeat(24) + '"';
  const googleId = '1' + 'A'.repeat(32) + '_' + 'B'.repeat(6);
  const errors = scanAddedLines([{ file: '00_Config.gs', lines: [credential, `const runtimeId = "${googleId}";`] }]);
  assert.equal(errors.length, 2);
});

test('raw customer metadata in logging fails sensitive diff validation', () => {
  const logCall = ['log', 'Line_', '(', '"FLOW", ', 'customer', 'Name, "", "", "INFO")'].join('');
  const errors = scanAddedLines([{ file: '06a_EntryPoints.gs', lines: [logCall] }]);
  assert.equal(errors.length, 1);
  assert.match(errors[0], /operational metadata/);
});
