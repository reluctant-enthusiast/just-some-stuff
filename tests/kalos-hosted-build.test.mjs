import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const index = await readFile(new URL('../kalos/index.html', import.meta.url), 'utf8');

assert.match(index, /<div\b[^>]*id=["']app["']>/,
  'the public route must contain the playable Kalos app directly');
assert.doesNotMatch(index, /\batob\s*\(/,
  'the public route must not depend on runtime Base64 decoding');
assert.doesNotMatch(index, /payload-\d+\.txt/,
  'the public route must not assemble split payload files');
assert.match(index, /enrichment\.js\?v=m01-context-20260804/,
  'the direct build must preserve the current narrative enrichment');

console.log('Kalos hosted-build integrity checks passed.');
