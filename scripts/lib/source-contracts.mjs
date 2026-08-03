import vm from 'node:vm';

export function extractInitializer(source, name) {
  const marker = new RegExp(`\\b(?:const|let|var)\\s+${escapeRegExp(name)}\\s*=`);
  const match = marker.exec(source);
  if (!match) throw new Error(`Initializer not found: ${name}`);

  const start = match.index + match[0].length;
  let quote = null;
  let escaped = false;
  let lineComment = false;
  let blockComment = false;
  let round = 0;
  let square = 0;
  let curly = 0;

  for (let i = start; i < source.length; i += 1) {
    const char = source[i];
    const next = source[i + 1];

    if (lineComment) {
      if (char === '\n') lineComment = false;
      continue;
    }
    if (blockComment) {
      if (char === '*' && next === '/') {
        blockComment = false;
        i += 1;
      }
      continue;
    }
    if (quote) {
      if (escaped) escaped = false;
      else if (char === '\\') escaped = true;
      else if (char === quote) quote = null;
      continue;
    }
    if (char === '/' && next === '/') {
      lineComment = true;
      i += 1;
      continue;
    }
    if (char === '/' && next === '*') {
      blockComment = true;
      i += 1;
      continue;
    }
    if (char === "'" || char === '"' || char === '`') {
      quote = char;
      continue;
    }

    if (char === '(') round += 1;
    else if (char === ')') round -= 1;
    else if (char === '[') square += 1;
    else if (char === ']') square -= 1;
    else if (char === '{') curly += 1;
    else if (char === '}') curly -= 1;
    else if (char === ';' && round === 0 && square === 0 && curly === 0) {
      return source.slice(start, i).trim();
    }
  }

  throw new Error(`Unterminated initializer: ${name}`);
}

export function evaluateInitializer(source, name, context = {}) {
  const expression = extractInitializer(source, name);
  return vm.runInNewContext(`(${expression})`, { ...context }, {
    filename: `${name}.contract.js`,
    timeout: 1000
  });
}

export function extractFunction(source, name) {
  const marker = new RegExp(`\\bfunction\\s+${escapeRegExp(name)}\\s*\\(`);
  const match = marker.exec(source);
  if (!match) throw new Error(`Function not found: ${name}`);
  const open = source.indexOf('{', match.index);
  if (open < 0) throw new Error(`Function body not found: ${name}`);

  let quote = null;
  let escaped = false;
  let lineComment = false;
  let blockComment = false;
  let depth = 0;
  for (let i = open; i < source.length; i += 1) {
    const char = source[i];
    const next = source[i + 1];
    if (lineComment) {
      if (char === '\n') lineComment = false;
      continue;
    }
    if (blockComment) {
      if (char === '*' && next === '/') {
        blockComment = false;
        i += 1;
      }
      continue;
    }
    if (quote) {
      if (escaped) escaped = false;
      else if (char === '\\') escaped = true;
      else if (char === quote) quote = null;
      continue;
    }
    if (char === '/' && next === '/') {
      lineComment = true;
      i += 1;
      continue;
    }
    if (char === '/' && next === '*') {
      blockComment = true;
      i += 1;
      continue;
    }
    if (char === "'" || char === '"' || char === '`') {
      quote = char;
      continue;
    }
    if (char === '{') depth += 1;
    if (char === '}') {
      depth -= 1;
      if (depth === 0) return source.slice(match.index, i + 1);
    }
  }
  throw new Error(`Unterminated function: ${name}`);
}

export function loadFunctions(source, names, context = {}) {
  const definitions = names.map((name) => extractFunction(source, name)).join('\n');
  const exportsExpression = names.map((name) => `${name}: ${name}`).join(',');
  return vm.runInNewContext(`(() => { ${definitions}; return {${exportsExpression}}; })()`, {
    ...context
  }, { timeout: 1000 });
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
