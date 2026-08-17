import {
  evaluateInitializer,
  loadFunctions
} from './source-contracts.mjs';

export function validateCriticalMappings(sources) {
  const errors = [];
  const finishScMirrorStatuses = evaluateInitializer(sources.config, 'FINISH_SC_MIRROR_STATUSES');
  const rootPolicy = evaluateInitializer(sources.config, 'OPS_ROUTING_POLICY', { FINISH_SC_MIRROR_STATUSES: finishScMirrorStatuses });
  const statusTypes = evaluateInitializer(sources.config, 'STATUS_TYPE_BY_LAST_STATUS');
  const positions = evaluateInitializer(sources.config, 'POSITION_BY_LAST_STATUS');
  const rawTail = evaluateInitializer(sources.config, 'RAW_DATA_CUSTOM_TAIL_HEADERS');
  const columnTypes = evaluateInitializer(sources.config, 'COLUMN_TYPES');
  const specialPolicy = evaluateInitializer(sources.config, 'SPECIAL_CASE_WRITER_POLICY');
  const extractorConfig = evaluateInitializer(sources.extractor, 'CONFIG');

  expectIncludes(errors, rootPolicy.SC_NAME_KEYWORDS['SC - Meilani'], 'GSI', 'root GSI -> Meilani');
  expectIncludes(errors, rootPolicy.SC_NAME_KEYWORDS['SC - Farhan'], 'Rejeki Seluler', 'root Rejeki Seluler -> Farhan');
  expectIncludes(errors, rootPolicy.SC_NAME_KEYWORDS['SC - Farhan'], 'Rejeki Seluller', 'root Rejeki Seluller -> Farhan');
  expectIncludes(errors, rootPolicy.SC_NAME_KEYWORDS['SC - Farhan'], 'CV Berkah', 'root CV Berkah -> Farhan');
  expectIncludes(errors, rootPolicy.SC_NAME_KEYWORDS['SC - Meindar'], 'Deltasindo', 'root Deltasindo -> Meindar');
  expectIncludes(errors, rootPolicy.SC_NAME_KEYWORDS['SC - Meindar'], 'EzCare', 'root EzCare default -> Meindar');
  expectIncludes(errors, rootPolicy.SC_NAME_KEYWORDS['SC - Meilani'], 'Samsung Exclusive', 'root Samsung Exclusive -> Meilani');
  expectIncludes(errors, rootPolicy.SC_NAME_KEYWORDS['SC - Meilani'], 'Samsung Authorized Service Centre by Unicom', 'root Samsung Unicom variant -> Meilani');

  const statusKeys = Object.keys(statusTypes).sort();
  const positionKeys = Object.keys(positions).sort();
  if (statusKeys.length < 100) errors.push(`status registry unexpectedly small: ${statusKeys.length}`);
  if (JSON.stringify(statusKeys) !== JSON.stringify(positionKeys)) {
    errors.push('status type and position registries must cover the same statuses');
  }
  expectValue(errors, rootPolicy.LAST_STATUS_BY_SHEET.Submission, ['SUBMITTED', 'CLAIM_INITIATE'], 'Submission status routing');
  expectIncludes(errors, rootPolicy.LAST_STATUS_BY_SHEET['Expired Claim'], 'CLAIM_EXPIRE', 'Expired Claim routing');
  expectIncludes(errors, rootPolicy.LAST_STATUS_BY_SHEET.Exclusion, 'CLAIM_CANCELLED', 'Exclusion routing');

  expectValue(errors, columnTypes.RAW.claim_submitted_datetime, 'DATETIME', 'submitted datetime type');
  expectValue(errors, columnTypes.RAW.claim_submission_date, 'DATE', 'legacy submission date type');
  for (const header of ['Update Status', 'Timestamp', 'Status', 'Remarks', 'AWB', 'Timestamp AWB']) {
    expectIncludes(errors, rawTail, header, `raw manual field ${header}`);
  }
  for (const header of ['Update Status', 'Timestamp', 'Status']) {
    expectIncludes(errors, specialPolicy.MANUAL_HEADERS, header, `Special Case manual field ${header}`);
  }

  expectPattern(errors, sources.config, /claimSubmissionDate:\s*'claim_submitted_datetime'/, 'Submission Date canonical source');
  expectPattern(errors, sources.config, /serialNumber:\s*'imei_number'/, 'serial number alias');
  expectPattern(errors, sources.routing, /fmt\('IMEI\/SN',\s*'@'/, 'IMEI/SN plain-text format');
  expectPattern(errors, sources.routing, /normalizeImeiSnText_/, 'IMEI/SN normalization consumer');
  expectPattern(errors, sources.postProcess, /deprecatedEverywhere\s*=\s*\['DB',\s*'Status Type',\s*'Update Status Asso',\s*'Timestamp Asso',\s*'Update Status Admin',\s*'Timestamp Admin'\]/, 'deprecated operational fields');
  expectPattern(errors, sources.enrichment, /\['AWB',\s*'Timestamp AWB'\]/, 'AWB manual restore contract');

  const extractorSpecial = loadFunctions(
    sources.extractor,
    ['_resolveSpecialDestination_', '_norm_', '_compactNorm_', '_resolvePicOverride_'],
    { RUNTIME_CACHE: { normByValue: {}, normSize: 0, normMaxSize: 1000 } }
  );
  expectObject(errors, extractorSpecial._resolveSpecialDestination_('gsi'), { sheetName: 'GSI', pic: 'MEILANI' }, 'Extractor GSI');
  expectObject(errors, extractorSpecial._resolveSpecialDestination_('ptdeltasindosagitamandiri'), { sheetName: 'Deltafone', pic: 'MEINDAR' }, 'Extractor Deltasindo');
  expectObject(errors, extractorSpecial._resolveSpecialDestination_('cvberkahathallah'), { sheetName: 'CV Berkah Athallah', pic: 'FARHAN' }, 'Extractor CV Berkah');
  expectObject(errors, extractorSpecial._resolveSpecialDestination_('rejekiseluller'), { sheetName: 'Rejeki Seluler', pic: 'FARHAN' }, 'Extractor Rejeki Seluller');
  expectValue(errors, extractorSpecial._resolvePicOverride_({ serviceCenterName: 'EzCare', deviceBrand: 'Apple' }, 'MEINDAR'), 'FARHAN', 'Extractor EzCare Apple');
  expectValue(errors, extractorSpecial._resolvePicOverride_({ serviceCenterName: 'EzCare', deviceBrand: 'Samsung' }, 'FARHAN'), 'MEINDAR', 'Extractor EzCare non-Apple');

  const samsungRule = extractorConfig.ROUTE_RULES.find((rule) => rule.sheet === 'Samsung Exclusive');
  if (!samsungRule || !samsungRule.tokens.includes('samsung authorized service centre by unicom')) {
    errors.push('Extractor Samsung Unicom route is missing');
  }
  expectValue(
    errors,
    extractorConfig.MEILANI_SC_NAME_OVERRIDES['samsung authorized by unicom palangkaraya service center'],
    'Samsung Exclusive',
    'Extractor Samsung Unicom override'
  );

  const salvage = loadFunctions(sources.salvage, [
    'normalizeKey_',
    'normalizeServiceCenterKey_',
    'resolvePicByBranch_'
  ]);
  expectValue(errors, salvage.resolvePicByBranch_('GSI', '', '', '', ''), 'Meilani', 'Salvage GSI');
  expectValue(errors, salvage.resolvePicByBranch_('Deltafone', 'Deltasindo', '', '', ''), 'Meindar', 'Salvage Deltasindo');
  expectValue(errors, salvage.resolvePicByBranch_('', 'CV Berkah', '', '', ''), 'Farhan', 'Salvage CV Berkah');
  expectValue(errors, salvage.resolvePicByBranch_('', 'Rejeki Seluller', '', '', ''), 'Farhan', 'Salvage Rejeki Seluller');
  expectValue(errors, salvage.resolvePicByBranch_('EzCare', 'EzCare', 'Apple', '', '2026-07-15'), 'Farhan', 'Salvage EzCare Apple cutoff');
  expectValue(errors, salvage.resolvePicByBranch_('EzCare', 'EzCare', 'Samsung', '', '2026-07-15'), 'Meindar', 'Salvage EzCare non-Apple');

  const utils = loadFunctions(sources.utils, ['normalizeImeiSnText_'], {
    Utilities: { formatString: (_format, value) => String(Math.trunc(value)) }
  });
  expectValue(errors, utils.normalizeImeiSnText_('001,234,567'), '001234567', 'IMEI leading-zero preservation');
  expectValue(errors, utils.normalizeImeiSnText_(123456789012345), '123456789012345', 'numeric IMEI normalization');

  expectPattern(errors, sources.routing, /isEzCare\s*&&\s*isApple[\s\S]*new Date\(2026,\s*6,\s*15\)[\s\S]*scFarhanName/, 'root EzCare Apple split');
  expectPattern(errors, sources.routing, /all other EzCare claims retain the existing Meindar mapping/, 'root EzCare non-Apple contract');
  return errors;
}

function expectIncludes(errors, actual, expected, label) {
  if (!Array.isArray(actual) || !actual.includes(expected)) errors.push(`${label}: expected ${expected}`);
}

function expectValue(errors, actual, expected, label) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    errors.push(`${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
}

function expectObject(errors, actual, expected, label) {
  expectValue(errors, actual, expected, label);
}

function expectPattern(errors, source, pattern, label) {
  if (!pattern.test(source)) errors.push(`${label}: source contract missing`);
}
