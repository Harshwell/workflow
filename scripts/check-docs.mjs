import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { parse } from '@mermaid-js/parser';
import {
  KNOWLEDGE_FILES,
  extractMermaidBlocks,
  mermaidDiagramType,
  validateKnowledgeArchitecture,
  validateLocalLinks,
  validateMarkdownPolicy
} from './lib/docs-validation.mjs';

const root = process.cwd();
const markdownByFile = new Map(
  KNOWLEDGE_FILES.map((file) => [file, fs.readFileSync(path.join(root, file), 'utf8')])
);
const errors = validateKnowledgeArchitecture(root);

for (const [file, markdown] of markdownByFile) {
  errors.push(...validateMarkdownPolicy(file, markdown));
  errors.push(...validateLocalLinks(root, file, markdown, markdownByFile));
  const blocks = extractMermaidBlocks(markdown);
  for (let index = 0; index < blocks.length; index += 1) {
    const source = blocks[index];
    try {
      await parse(mermaidDiagramType(source), source);
    } catch (error) {
      errors.push(`${file}: invalid Mermaid block ${index + 1}: ${error.message}`);
    }
  }
}

if (errors.length) {
  process.stderr.write(`${errors.join('\n')}\n`);
  process.exit(1);
}

process.stdout.write(`Documentation OK: ${KNOWLEDGE_FILES.join(', ')}\n`);
