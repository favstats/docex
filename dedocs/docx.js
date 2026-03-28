'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const { DEDOCS_VERSION, parsePackage, serializePackage } = require('./format');
const { createGuides } = require('./guide');
const { applyTransforms } = require('./overlay');

function makeTempDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'dedocs-'));
}

function removeDir(dirPath) {
  if (!dirPath || !fs.existsSync(dirPath)) return;
  fs.rmSync(dirPath, { recursive: true, force: true });
}

function listZipEntries(docxPath) {
  const output = execFileSync('unzip', ['-Z1', docxPath], {
    encoding: 'utf8',
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  return output
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(line => line && !line.endsWith('/'));
}

function extractDocx(docxPath, targetDir) {
  execFileSync('unzip', ['-o', docxPath, '-d', targetDir], {
    stdio: ['pipe', 'pipe', 'pipe'],
  });
}

function inferMediaType(partPath) {
  const ext = path.extname(partPath).toLowerCase();
  if (ext === '.xml' || ext === '.rels') return 'application/xml';
  if (ext === '.png') return 'image/png';
  if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg';
  if (ext === '.gif') return 'image/gif';
  if (ext === '.svg') return 'image/svg+xml';
  if (ext === '.bin') return 'application/octet-stream';
  if (ext === '.txt') return 'text/plain';
  return 'application/octet-stream';
}

function isUtf8RoundTrip(buffer) {
  return Buffer.from(buffer.toString('utf8'), 'utf8').equals(buffer);
}

function inferEncoding(partPath, buffer) {
  const ext = path.extname(partPath).toLowerCase();
  if (ext === '.xml' || ext === '.rels' || ext === '.txt' || ext === '.json' || ext === '.md' || ext === '.yaml' || ext === '.yml') {
    return 'utf8';
  }
  return isUtf8RoundTrip(buffer) ? 'utf8' : 'base64';
}

function packageFromDocx(docxPath, opts = {}) {
  const absDocxPath = path.resolve(docxPath);
  const tmpDir = makeTempDir();

  try {
    extractDocx(absDocxPath, tmpDir);
    const parts = listZipEntries(absDocxPath).map(partPath => {
      const buffer = fs.readFileSync(path.join(tmpDir, partPath));
      return {
        path: partPath,
        mediaType: inferMediaType(partPath),
        encoding: inferEncoding(partPath, buffer),
        buffer,
      };
    });

    const pkg = {
      version: DEDOCS_VERSION,
      package: 'docx',
      fidelity: 'package-exact',
      source: opts.source || path.basename(absDocxPath),
      guides: [],
      transforms: [],
      parts,
    };

    if (opts.includeGuides !== false) {
      pkg.guides = createGuides(pkg);
    }

    return pkg;
  } finally {
    removeDir(tmpDir);
  }
}

function dedocsFromDocx(docxPath, opts = {}) {
  return serializePackage(packageFromDocx(docxPath, opts));
}

function normalizePackage(pkg, opts = {}) {
  const normalized = {
    version: pkg.version || DEDOCS_VERSION,
    package: pkg.package || 'docx',
    fidelity: pkg.fidelity || 'package-exact',
    source: pkg.source || '',
    transforms: Array.isArray(pkg.transforms) ? pkg.transforms.map(transform => ({ ...transform })) : [],
    parts: pkg.parts.map(part => ({
      ...part,
      buffer: Buffer.from(part.buffer),
    })),
    guides: [],
  };

  if (opts.regenerateGuides !== false) {
    const guideSource = opts.previewTransforms === false
      ? normalized
      : applyTransforms({
          ...normalized,
          parts: normalized.parts.map(part => ({
            ...part,
            buffer: Buffer.from(part.buffer),
          })),
        });
    normalized.guides = createGuides(guideSource);
  } else {
    normalized.guides = Array.isArray(pkg.guides) ? pkg.guides.map(guide => ({ ...guide })) : [];
  }

  return normalized;
}

function normalizeDedocsText(text, opts = {}) {
  const parsed = parsePackage(text, { strictMetadata: false });
  return serializePackage(normalizePackage(parsed, opts));
}

function normalizeDedocsFile(inputPath, outputPath, opts = {}) {
  const absInputPath = path.resolve(inputPath);
  const absOutputPath = path.resolve(outputPath || inputPath);
  const text = fs.readFileSync(absInputPath, 'utf8');
  const normalizedText = normalizeDedocsText(text, opts);
  fs.mkdirSync(path.dirname(absOutputPath), { recursive: true });
  fs.writeFileSync(absOutputPath, normalizedText, 'utf8');
  return absOutputPath;
}

function writeDedocsFile(docxPath, dedocsPath, opts = {}) {
  const absDedocsPath = path.resolve(dedocsPath);
  fs.mkdirSync(path.dirname(absDedocsPath), { recursive: true });
  const text = dedocsFromDocx(docxPath, opts);
  fs.writeFileSync(absDedocsPath, text, 'utf8');
  return absDedocsPath;
}

function ensureParentDir(filePath) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
}

function writePackageToDirectory(pkg, dirPath) {
  for (const part of pkg.parts) {
    const partPath = path.join(dirPath, part.path);
    ensureParentDir(partPath);
    fs.writeFileSync(partPath, part.buffer);
  }
}

function compilePackageToDocx(pkg, outputPath) {
  const absOutputPath = path.resolve(outputPath);
  const tmpDir = makeTempDir();

  try {
    writePackageToDirectory(pkg, tmpDir);
    fs.mkdirSync(path.dirname(absOutputPath), { recursive: true });
    if (fs.existsSync(absOutputPath)) fs.unlinkSync(absOutputPath);
    execFileSync('zip', ['-X', '-r', '-q', absOutputPath, '.'], {
      cwd: tmpDir,
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    return absOutputPath;
  } finally {
    removeDir(tmpDir);
  }
}

function compileDedocsText(text, outputPath, opts = {}) {
  const parsed = parsePackage(text, { strictMetadata: !!opts.strictMetadata });
  const transformed = applyTransforms(parsed);
  return compilePackageToDocx(transformed, outputPath);
}

function compileDedocsFile(dedocsPath, outputPath, opts = {}) {
  const text = fs.readFileSync(path.resolve(dedocsPath), 'utf8');
  return compileDedocsText(text, outputPath, opts);
}

function packageFromDedocsText(text, opts = {}) {
  return parsePackage(text, { strictMetadata: !!opts.strictMetadata });
}

function readDedocsFile(dedocsPath, opts = {}) {
  const text = fs.readFileSync(path.resolve(dedocsPath), 'utf8');
  return packageFromDedocsText(text, opts);
}

function listDirFiles(rootDir) {
  const results = [];

  function walk(dirPath) {
    const names = fs.readdirSync(dirPath).sort();
    for (const name of names) {
      const fullPath = path.join(dirPath, name);
      const relPath = path.relative(rootDir, fullPath).replace(/\\/g, '/');
      const stat = fs.statSync(fullPath);
      if (stat.isDirectory()) {
        walk(fullPath);
      } else {
        results.push(relPath);
      }
    }
  }

  walk(rootDir);
  return results;
}

function compareDocxPackages(leftPath, rightPath) {
  const leftTmp = makeTempDir();
  const rightTmp = makeTempDir();

  try {
    extractDocx(path.resolve(leftPath), leftTmp);
    extractDocx(path.resolve(rightPath), rightTmp);

    const leftFiles = listDirFiles(leftTmp);
    const rightFiles = listDirFiles(rightTmp);
    const fileSet = new Set([...leftFiles, ...rightFiles]);
    const diffs = [];

    for (const relPath of Array.from(fileSet).sort()) {
      const leftFile = path.join(leftTmp, relPath);
      const rightFile = path.join(rightTmp, relPath);
      const leftExists = fs.existsSync(leftFile);
      const rightExists = fs.existsSync(rightFile);

      if (!leftExists || !rightExists) {
        diffs.push({ path: relPath, type: 'missing' });
        continue;
      }

      const leftBuffer = fs.readFileSync(leftFile);
      const rightBuffer = fs.readFileSync(rightFile);
      if (!leftBuffer.equals(rightBuffer)) {
        diffs.push({ path: relPath, type: 'content' });
      }
    }

    return {
      equal: diffs.length === 0,
      diffs,
    };
  } finally {
    removeDir(leftTmp);
    removeDir(rightTmp);
  }
}

module.exports = {
  applyTransforms,
  compareDocxPackages,
  compileDedocsFile,
  compileDedocsText,
  compilePackageToDocx,
  dedocsFromDocx,
  normalizeDedocsFile,
  normalizeDedocsText,
  normalizePackage,
  packageFromDedocsText,
  packageFromDocx,
  readDedocsFile,
  writeDedocsFile,
};
