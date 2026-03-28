#!/usr/bin/env node
'use strict';

const path = require('path');
const {
  compareDocxPackages,
  compileDedocsFile,
  normalizeDedocsFile,
  writeDedocsFile,
} = require('./index');

function usage() {
  return [
    'Usage:',
    '  node dedocs/cli.js decompile <input.docx> <output.dedocs>',
    '  node dedocs/cli.js compile <input.dedocs> <output.docx>',
    '  node dedocs/cli.js normalize <input.dedocs> [output.dedocs]',
    '  node dedocs/cli.js verify <left.docx> <right.docx>',
  ].join('\n');
}

function main(argv) {
  const [command, arg1, arg2] = argv;
  if (!command) {
    process.stderr.write(usage() + '\n');
    process.exitCode = 1;
    return;
  }

  if (command === 'decompile') {
    if (!arg1 || !arg2) {
      throw new Error('decompile requires input.docx and output.dedocs');
    }
    const output = writeDedocsFile(arg1, arg2, { source: path.basename(arg1) });
    process.stdout.write(output + '\n');
    return;
  }

  if (command === 'compile') {
    if (!arg1 || !arg2) {
      throw new Error('compile requires input.dedocs and output.docx');
    }
    const output = compileDedocsFile(arg1, arg2);
    process.stdout.write(output + '\n');
    return;
  }

  if (command === 'verify') {
    if (!arg1 || !arg2) {
      throw new Error('verify requires left.docx and right.docx');
    }
    const result = compareDocxPackages(arg1, arg2);
    process.stdout.write(JSON.stringify(result, null, 2) + '\n');
    process.exitCode = result.equal ? 0 : 1;
    return;
  }

  if (command === 'normalize') {
    if (!arg1) {
      throw new Error('normalize requires input.dedocs');
    }
    const output = normalizeDedocsFile(arg1, arg2 || arg1);
    process.stdout.write(output + '\n');
    return;
  }

  throw new Error(`Unknown command: ${command}`);
}

try {
  main(process.argv.slice(2));
} catch (error) {
  process.stderr.write(String(error.message || error) + '\n');
  process.stderr.write(usage() + '\n');
  process.exitCode = 1;
}
