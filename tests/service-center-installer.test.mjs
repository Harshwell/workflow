import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';

const source = fs.readFileSync('optional-project/Service Center Extractor.js', 'utf8');

function functionBody(name) {
  const start = source.indexOf(`function ${name}()`);
  assert.notEqual(start, -1, `${name} must exist`);

  const bodyStart = source.indexOf('{', start);
  let depth = 0;
  for (let index = bodyStart; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(bodyStart + 1, index);
  }
  assert.fail(`${name} must have a complete function body`);
}

test('manual installer defers menu UI creation until spreadsheet open', () => {
  const installer = functionBody('installServiceCenterTransferMenu');

  assert.match(installer, /newTrigger\('onOpenServiceCenterTransfer'\)/);
  assert.doesNotMatch(installer, /onOpenServiceCenterTransfer\s*\(/);
  assert.doesNotMatch(installer, /SpreadsheetApp\.getUi\s*\(/);
});
