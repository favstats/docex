'use strict';

function countLiteral(text, needle) {
  if (needle === '') {
    throw new Error('replace-text find payload cannot be empty');
  }
  return text.split(needle).length - 1;
}

function replaceLiteral(text, needle, replacement) {
  return text.split(needle).join(replacement);
}

function clonePart(part) {
  return {
    ...part,
    buffer: Buffer.from(part.buffer),
  };
}

function applyTransforms(pkg) {
  const transforms = Array.isArray(pkg.transforms) ? pkg.transforms : [];
  if (transforms.length === 0) return pkg;

  const nextPkg = {
    ...pkg,
    parts: pkg.parts.map(clonePart),
  };

  for (const transform of transforms) {
    if (!transform || transform.type !== 'replace-text') {
      throw new Error(`Unsupported transform type: ${transform && transform.type}`);
    }

    const part = nextPkg.parts.find(candidate => candidate.path === transform.part);
    if (!part) {
      throw new Error(`Transform target part not found: ${transform.part}`);
    }
    if (part.encoding !== 'utf8') {
      throw new Error(`replace-text only supports utf8 parts: ${transform.part}`);
    }

    const currentText = part.buffer.toString('utf8');
    const matches = countLiteral(currentText, transform.find);
    const expected = transform.count === '' || transform.count == null
      ? null
      : Number(transform.count);

    if (expected !== null && (!Number.isInteger(expected) || expected < 0)) {
      throw new Error(`Invalid replace-text count for ${transform.part}: ${transform.count}`);
    }
    if (expected !== null && matches !== expected) {
      throw new Error(`replace-text expected ${expected} matches in ${transform.part} but found ${matches}`);
    }
    if (matches === 0) {
      throw new Error(`replace-text found no matches in ${transform.part}`);
    }

    const nextText = replaceLiteral(currentText, transform.find, transform.replace);
    part.buffer = Buffer.from(nextText, 'utf8');
    part.text = nextText;
  }

  return nextPkg;
}

module.exports = {
  applyTransforms,
};
