import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { validateCriticalMappings } from '../scripts/lib/mapping-contracts.mjs';

export function loadSources() {
  return {
    config: fs.readFileSync('00_Config.gs', 'utf8'),
    utils: fs.readFileSync('01_Utils.gs', 'utf8'),
    sheets: fs.readFileSync('03_SheetsAndValidation.gs', 'utf8'),
    routing: fs.readFileSync('05b_Pipeline_RoutingOperational.gs', 'utf8'),
    entryPoints: fs.readFileSync('06a_EntryPoints.gs', 'utf8'),
    enrichment: fs.readFileSync('06b_PipelineAndEnrichment.gs', 'utf8'),
    postProcess: fs.readFileSync('06c_PostProcessAndUtils.gs', 'utf8'),
    extractor: fs.readFileSync('optional-project/Service Center Extractor.js', 'utf8'),
    scMeilani: fs.readFileSync('optional-project/SC-Meilani.js', 'utf8'),
    salvage: fs.readFileSync('optional-project/salvage', 'utf8'),
    samsungClaim: fs.readFileSync(
      'optional-project/Project Samsung/Samsung-Claim-Sync.js',
      'utf8'
    )
  };
}

test('critical mapping, status, header, and manual-field contracts remain aligned', () => {
  assert.deepEqual(validateCriticalMappings(loadSources()), []);
});

test('Samsung Claim Sync keeps its strict brand, cutoff, and target contracts', () => {
  const { samsungClaim } = loadSources();

  assert.match(samsungClaim, /TARGET_SPREADSHEET_ID:\s*'1eXrN6pFXHr1-Qj208tzsMa5ysXgEv84LG-XiV6HhcJI'/);
  assert.match(samsungClaim, /MIN_SUBMITTED_DATE:\s*new Date\(2026, 5, 1\)/);
  assert.match(samsungClaim, /if \(deviceBrand !== 'samsung'\) continue;/);
  assert.doesNotMatch(samsungClaim, /deviceBrand\.includes\('samsung'\)/);
});
