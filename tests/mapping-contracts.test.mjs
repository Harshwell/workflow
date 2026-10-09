import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { validateCriticalMappings } from '../scripts/lib/mapping-contracts.mjs';
import { evaluateInitializer, loadFunctions } from '../scripts/lib/source-contracts.mjs';

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
    outstanding: fs.readFileSync('optional-project/Outstanding', 'utf8'),
    samsungClaim: fs.readFileSync(
      'optional-project/Project Samsung/Samsung-Claim-Sync.js',
      'utf8'
    )
  };
}

test('critical mapping, status, header, and manual-field contracts remain aligned', () => {
  assert.deepEqual(validateCriticalMappings(loadSources()), []);
});

test('Agung Cellular and Platinum Care resolve to Meindar across all mapping consumers', () => {
  const sources = loadSources();
  const FINISH_ONLY_REPLACEMENT_STATUSES = evaluateInitializer(sources.config, 'FINISH_ONLY_REPLACEMENT_STATUSES');
  const policy = evaluateInitializer(sources.config, 'OPS_ROUTING_POLICY', { FINISH_ONLY_REPLACEMENT_STATUSES });
  const branches = evaluateInitializer(sources.config, 'BRANCH_KEYWORDS');
  const extractorConfig = evaluateInitializer(sources.extractor, 'CONFIG');
  const routing = loadFunctions(sources.routing, ['filterScTargets05b_', 'normalizeScKeywordText05b_', 'scoreKeywords05b_']);
  const sub = loadFunctions(sources.entryPoints, ['__deriveServiceCenterPicSub06a_'], { OPS_ROUTING_POLICY: policy });
  const report = loadFunctions(sources.postProcess, [
    '__getBranchFromServiceCenter06_', '__getMiddlePicFromServiceCenter06_',
    '__normalizeHeaderText06_', '__normalizeScKeywordText06c_'
  ]);
  const extractor = loadFunctions(sources.extractor, ['_resolveSpecialDestination_']);
  const salvage = loadFunctions(sources.salvage, [
    'normalizeKey_', 'normalizeServiceCenterKey_', 'resolvePicByBranch_', 'resolveBranchByServiceCenter_'
  ]);
  const cfg = evaluateInitializer(sources.outstanding, 'CFG');
  const compiled = evaluateInitializer(sources.outstanding, 'COMPILED_CFG', { CFG: cfg });
  const outstandingHelpers = loadFunctions(sources.outstanding, ['containsAny_']);
  const outstanding = evaluateInitializer(sources.outstanding, 'PicResolver', {
    COMPILED_CFG: compiled, containsAny_: outstandingHelpers.containsAny_
  });
  const candidates = ['SC - Farhan', 'SC - Meilani', 'SC - Meindar', 'SC - Unmapped'];

  for (const [keyword, canonical] of [
    ['Agung Cellular', 'Agung Cellular Service Center'],
    ['Platinum Care', 'Platinum Care Service Centre']
  ]) {
    assert.ok(extractorConfig.DEST_SHEETS_TO_CLEAR.includes(canonical));
    assert.ok(extractorConfig.AUTO_MANAGED_DEST_SHEETS.has(canonical));
    assert.ok(extractorConfig.SERVICE_CENTER_MAPPING.some(row => row.name === canonical && row.pic === 'MEINDAR'));
    for (const name of [keyword, `${keyword} Service Center`, `${keyword} Service Centre`, `${keyword.toUpperCase()} Service Centre Jakarta`]) {
      const matchingBranches = Object.entries(branches).filter(([, tokens]) => tokens.some(token => name.toLowerCase().includes(token)));
      assert.deepEqual(matchingBranches.map(([branch]) => branch), [canonical], name);
      const targets = routing.filterScTargets05b_(candidates.slice(0, 3), name, ...candidates.slice(0, 3),
        ...candidates.slice(0, 3).map(sheet => policy.SC_NAME_KEYWORDS[sheet]), candidates[3]);
      assert.deepEqual(Array.from(targets), ['SC - Meindar'], name);
      assert.equal(sub.__deriveServiceCenterPicSub06a_(name), 'Meindar', name);
      assert.equal(report.__getBranchFromServiceCenter06_(name), canonical, name);
      assert.equal(report.__getMiddlePicFromServiceCenter06_(name), 'Meindar', name);
      const destination = extractor._resolveSpecialDestination_(name.toLowerCase().replace(/[^a-z0-9]/g, ''));
      assert.equal(destination.sheetName, canonical, name);
      assert.equal(destination.pic, 'MEINDAR', name);
      assert.ok(extractorConfig.ROUTE_RULES.find(rule => rule.sheet === canonical).tokens.some(token => name.toLowerCase().includes(token)));
      assert.equal(salvage.resolveBranchByServiceCenter_(name, ''), canonical, name);
      assert.equal(salvage.resolvePicByBranch_('', name, '', '', ''), 'Meindar', name);
      assert.equal(salvage.resolvePicByBranch_(name, '', '', '', ''), 'Meindar', name);
      assert.equal(outstanding.resolveMiddle_(name), 'Meindar', name);
    }
  }
  assert.equal(outstanding.resolveMiddle_('Klikcare'), 'Ivan');
  assert.equal(outstanding.resolveMiddle_('Unknown Service Center'), 'Unknown');
  assert.equal(report.__getMiddlePicFromServiceCenter06_('Unknown Service Center'), 'Unknown');
  assert.equal(extractor._resolveSpecialDestination_('unknownservicecenter'), null);
  assert.equal(salvage.resolveBranchByServiceCenter_('Unknown Service Center', 'Manual Branch'), 'Manual Branch');
  assert.equal(salvage.resolvePicByBranch_('Agung Cellular', 'EzCare', 'Apple', '', ''), 'Farhan');
});

test('Samsung Claim Sync keeps its strict brand, cutoff, and target contracts', () => {
  const { samsungClaim } = loadSources();

  assert.match(samsungClaim, /SOURCE_SPREADSHEET_ID:\s*'1zRlYrSRssv9LVcPKEq90CmmvTRsZoN_TqfIg2pNufbc'/);
  assert.match(samsungClaim, /TARGET_SPREADSHEET_ID:\s*'1eXrN6pFXHr1-Qj208tzsMa5ysXgEv84LG-XiV6HhcJI'/);
  assert.match(samsungClaim, /MIN_SUBMITTED_DATE:\s*new Date\(2026, 5, 1\)/);
  assert.match(samsungClaim, /if \(deviceBrand !== 'samsung'\) continue;/);
  assert.doesNotMatch(samsungClaim, /deviceBrand\.includes\('samsung'\)/);
  assert.doesNotMatch(samsungClaim, /SOURCE_SPREADSHEET_PROPERTY|PropertiesService/);
});

test('Salvage resolves full Service Center names independently of blank or stale Branch', () => {
  const { salvage: source } = loadSources();
  const salvage = loadFunctions(source, ['normalizeKey_', 'normalizeServiceCenterKey_', 'resolvePicByBranch_']);
  for (const [name, pic] of [
    ['J-Bros Computer Service Center Padang', 'Meindar'],
    ['B-Store Service Centre Jakarta', 'Meindar'],
    ['PT DELTASINDO SAGITA MANDIRI - Sorong Papua Barat', 'Meindar'],
    ['GH Store - Pontianak', 'Meindar'],
    ['CV Berkah Athallah Branch Store', 'Farhan']
  ]) {
    for (const branch of ['', 'Mitracare', 'GSI']) {
      assert.equal(salvage.resolvePicByBranch_(branch, name, '', '', ''), pic, name + ' with Branch ' + branch);
      assert.equal(salvage.resolvePicByBranch_(branch, name.toLowerCase(), '', '', ''), pic, name + ' lowercase');
    }
    assert.equal(salvage.resolvePicByBranch_(name, '', '', '', ''), pic, name + ' from Branch only');
  }
  assert.equal(salvage.resolvePicByBranch_('GH Store - Pontianak', 'EzCare', 'Apple', '', ''), 'Farhan');
  assert.equal(salvage.resolvePicByBranch_('Other Branch', 'Other Service Center', '', '', ''), 'Unknown');
});

test('Salvage canonicalizes requested Branch names without assigning a new Skylensindo PIC', () => {
  const { salvage: source } = loadSources();
  const salvage = loadFunctions(source, ['normalizeKey_', 'normalizeServiceCenterKey_', 'resolveBranchByServiceCenter_', 'resolvePicByBranch_']);
  for (const [name, branch, pic] of [
    ['CV Berkah Athallah Branch Store', 'CV Berkah', 'Farhan'],
    ['GH Store - Pontianak', 'GH Store', 'Meindar'],
    ['PT DELTASINDO SAGITA MANDIRI - Sorong Papua Barat', 'Deltafone', 'Meindar'],
    ['Skylensindo Service Center', 'Skylensindo', 'Unknown']
  ]) {
    for (const fallback of ['', 'Old Branch', branch]) {
      assert.equal(salvage.resolveBranchByServiceCenter_(name, fallback), branch, name);
      assert.equal(salvage.resolveBranchByServiceCenter_(name.toLowerCase(), fallback), branch, name + ' lowercase');
    }
    assert.equal(salvage.resolvePicByBranch_(branch, name, '', '', ''), pic, name + ' PIC');
  }
  assert.equal(salvage.resolveBranchByServiceCenter_('Unlisted Service Center', 'Manual Branch'), 'Manual Branch');
});
