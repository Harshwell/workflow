import fs from 'node:fs';
import path from 'node:path';

export const KNOWLEDGE_FILES = ['AGENTS.md', 'CHANGELOG.md', 'README.md'];

export function githubSlug(value) {
  return value
    .trim()
    .toLowerCase()
    .replace(/<[^>]+>/g, '')
    .replace(/[^\p{L}\p{N}\s_-]/gu, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-');
}

export function collectAnchors(markdown) {
  const counts = new Map();
  const anchors = new Set();
  for (const line of markdown.split(/\r?\n/)) {
    const match = /^(#{1,6})\s+(.+?)\s*$/.exec(line);
    if (!match) continue;
    const base = githubSlug(match[2].replace(/\s+#+$/, ''));
    const count = counts.get(base) || 0;
    counts.set(base, count + 1);
    anchors.add(count === 0 ? base : `${base}-${count}`);
  }
  return anchors;
}

export function validateMarkdownPolicy(file, markdown) {
  const errors = [];
  const headings = markdown.split(/\r?\n/).filter((line) => /^#{1,6}\s+/.test(line));
  const banned = [
    /^(#{1,6})\s+timeline\b/i,
    /^(#{1,6})\s+update implementasi terbaru\b/i,
    /^(#{1,6})\s+audit (?:lama|selesai|terbaru)\b/i
  ];
  for (const heading of headings) {
    if (banned.some((pattern) => pattern.test(heading))) {
      errors.push(`${file}: prohibited historical/current-update heading: ${heading}`);
    }
  }
  if (file === 'README.md' && !/^## Apps Script UAT\s*$/m.test(markdown)) {
    errors.push('README.md: missing Apps Script UAT section');
  }
  if (file === 'CHANGELOG.md' && !/^## Unreleased\s*$/m.test(markdown)) {
    errors.push('CHANGELOG.md: missing Unreleased section');
  }
  return errors;
}

export function findMarkdownFiles(root) {
  const found = [];
  const visit = (directory) => {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      if (entry.name === '.git' || entry.name === 'node_modules') continue;
      const absolute = path.join(directory, entry.name);
      if (entry.isDirectory()) visit(absolute);
      else if (entry.isFile() && entry.name.toLowerCase().endsWith('.md')) {
        found.push(path.relative(root, absolute).replaceAll('\\', '/'));
      }
    }
  };
  visit(root);
  return found.sort();
}

export function validateKnowledgeArchitecture(root) {
  const actual = findMarkdownFiles(root);
  const expected = [...KNOWLEDGE_FILES].sort();
  if (JSON.stringify(actual) === JSON.stringify(expected)) return [];
  return [`Knowledge files must be exactly ${expected.join(', ')}; found ${actual.join(', ') || '(none)'}`];
}

export function validateLocalLinks(root, file, markdown, markdownByFile) {
  const errors = [];
  const linkPattern = /(?<!!)\[[^\]]*\]\(([^)\s]+)(?:\s+["'][^"']*["'])?\)/g;
  for (const match of markdown.matchAll(linkPattern)) {
    const target = match[1];
    if (/^(?:https?:|mailto:|tel:)/i.test(target)) continue;
    const [rawPath, rawAnchor = ''] = target.split('#', 2);
    const decodedPath = decodeURIComponent(rawPath || '');
    const resolvedRelative = decodedPath
      ? path.relative(root, path.resolve(root, path.dirname(file), decodedPath)).replaceAll('\\', '/')
      : file;
    const absolute = path.resolve(root, resolvedRelative);
    if (decodedPath && !fs.existsSync(absolute)) {
      errors.push(`${file}: broken local link ${target}`);
      continue;
    }
    if (rawAnchor) {
      const linkedMarkdown = markdownByFile.get(resolvedRelative)
        ?? (fs.existsSync(absolute) ? fs.readFileSync(absolute, 'utf8') : '');
      const anchors = collectAnchors(linkedMarkdown);
      if (!anchors.has(decodeURIComponent(rawAnchor).toLowerCase())) {
        errors.push(`${file}: missing anchor ${target}`);
      }
    }
  }
  return errors;
}

export function extractMermaidBlocks(markdown) {
  const blocks = [];
  const pattern = /```mermaid\s*\r?\n([\s\S]*?)```/gi;
  for (const match of markdown.matchAll(pattern)) blocks.push(match[1].trim());
  return blocks;
}

export function mermaidDiagramType(source) {
  const first = source.split(/\r?\n/, 1)[0].trim().split(/\s+/, 1)[0].toLowerCase();
  if (first === 'graph') return 'flowchart';
  if (first === 'architecture-beta') return 'architecture';
  return first;
}
