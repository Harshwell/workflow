const GENERIC_REASON = /^(?:n\/?a|na|none|no impact|not needed|not applicable|docs? only|unchanged|irrelevant)[.!\s]*$/i;

export function validateDocumentationDrift(changedFiles, prBody = '') {
  const files = new Set(changedFiles.map(normalizePath));
  const errors = [];
  const material = [...files].some(isMaterialChange);
  const readmeRelevant = [...files].some(isReadmeRelevant);
  const agentsRelevant = [...files].some(isAgentsRelevant);

  if (material && !files.has('CHANGELOG.md')) {
    errors.push('Material repository changes require CHANGELOG.md; there is no escape hatch.');
  }
  if (readmeRelevant && !files.has('README.md')) {
    const reason = impactReason(prBody, 'README');
    if (!reason) errors.push('README.md evaluation required. Update README.md or add a specific Docs-Impact-README reason to the PR body.');
  }
  if (agentsRelevant && !files.has('AGENTS.md')) {
    const reason = impactReason(prBody, 'AGENTS');
    if (!reason) errors.push('AGENTS.md evaluation required. Update AGENTS.md or add a specific Docs-Impact-AGENTS reason to the PR body.');
  }
  return errors;
}

export function impactReason(prBody, target) {
  const pattern = new RegExp(`^Docs-Impact-${target}:\\s*none\\s*(?:—|-)\\s*(.+)$`, 'im');
  const match = pattern.exec(prBody || '');
  if (!match) return null;
  const reason = match[1].trim();
  if (reason.length < 12 || GENERIC_REASON.test(reason)) return null;
  return reason;
}

export function scanAddedLines(records) {
  const errors = [];
  for (const record of records) {
    const file = normalizePath(record.file);
    if (file === 'package-lock.json' || file.endsWith('.map')) continue;
    for (const item of record.lines) {
      const line = typeof item === 'string' ? item : item.text;
      const lineNumber = typeof item === 'string' ? '?' : item.line;
      const issue = sensitiveIssue(line, file);
      if (issue) errors.push(`${file}:${lineNumber}: ${issue}`);
    }
  }
  return errors;
}

export function sensitiveIssue(line, file = '') {
  const text = String(line || '');
  if (/-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/.test(text)) return 'private key material added';
  if (/\b(?:password|passwd|api[_-]?key|access[_-]?token|refresh[_-]?token|client[_-]?secret)\b\s*[:=]\s*["'][^"']{8,}["']/i.test(text)
      && !/(?:example|placeholder|redacted|dummy|test-only|process\.env|getProp)/i.test(text)) {
    return 'credential-like literal added';
  }
  const resourceIds = text.match(/\b1[A-Za-z0-9_-]{30,}\b/g) || [];
  if (resourceIds.some((value) => !(value.length === 40 && /^[0-9a-f]+$/i.test(value)))) {
    return 'Google resource identifier added';
  }

  const emails = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [];
  for (const email of emails) {
    if (!/(?:example\.(?:com|org|net)|users\.noreply\.github\.com)$/i.test(email)) return 'non-allowlisted email added';
  }

  const urls = text.match(/https?:\/\/[^\s"')>]+/gi) || [];
  for (const url of urls) {
    try {
      const host = new URL(url).hostname.toLowerCase();
      if (!PUBLIC_HOSTS.some((allowed) => host === allowed || host.endsWith(`.${allowed}`))) {
        return 'non-allowlisted URL added';
      }
    } catch {
      return 'malformed URL added';
    }
  }

  if (/\b(?:logLine_|console\.(?:log|info|warn|error)|Logger\.log)\s*\([^\n]*(?:customer|holder|emailSubject|email_subject|attachmentName|attachment_name|rawEmail|raw_customer)/i.test(text)) {
    return 'raw operational metadata added to logging';
  }
  if (/\b(?:subject|attachment(?:Name|_name)|customer(?:Name|_name)|holder(?:Name|_name))\s*[:=]\s*["'][^"']+["']/i.test(text)
      && /(?:fixture|testdata|sample|raw|log)/i.test(file)) {
    return 'raw-looking operational fixture added';
  }
  return null;
}

const PUBLIC_HOSTS = [
  'example.com',
  'github.com',
  'githubusercontent.com',
  'google.com',
  'mermaid.js.org',
  'nodejs.org',
  'npmjs.com',
  'registry.npmjs.org'
];

function isMaterialChange(file) {
  return /^(?:[^/]+\.gs|appsscript\.json|optional-project\/|scripts\/|tests\/|\.github\/|package(?:-lock)?\.json|\.node-version|\.markdownlint)/.test(file);
}

function isReadmeRelevant(file) {
  return /^(?:00_Config\.gs|03_SheetsAndValidation\.gs|05[bc]_[^/]+\.gs|06a_[^/]+\.gs|appsscript\.json|optional-project\/|tests\/mapping-contracts\.test\.mjs)/.test(file);
}

function isAgentsRelevant(file) {
  return /^(?:AGENTS\.md|scripts\/|tests\/validator-negative\.test\.mjs|\.github\/|package(?:-lock)?\.json|\.node-version|\.markdownlint)/.test(file);
}

function normalizePath(file) {
  return String(file || '').replaceAll('\\', '/').replace(/^\.\//, '');
}
