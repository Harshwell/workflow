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
    enrichment: fs.readFileSync('06b_PipelineAndEnrichment.gs', 'utf8'),
    postProcess: fs.readFileSync('06c_PostProcessAndUtils.gs', 'utf8'),
    extractor: fs.readFileSync('optional-project/Service Center Extractor.js', 'utf8'),
    salvage: fs.readFileSync('optional-project/salvage', 'utf8')
  };
}

test('critical mapping, status, header, and manual-field contracts remain aligned', () => {
  assert.deepEqual(validateCriticalMappings(loadSources()), []);
});
