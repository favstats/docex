'use strict';

const {
  DEDOCS_VERSION,
  parsePackage,
  serializePackage,
  sha256,
} = require('./format');

const {
  compareDocxPackages,
  compileDedocsFile,
  compileDedocsText,
  compilePackageToDocx,
  dedocsFromDocx,
  packageFromDedocsText,
  packageFromDocx,
  readDedocsFile,
  writeDedocsFile,
} = require('./docx');

module.exports = {
  DEDOCS_VERSION,
  compareDocxPackages,
  compileDedocsFile,
  compileDedocsText,
  compilePackageToDocx,
  dedocsFromDocx,
  packageFromDedocsText,
  packageFromDocx,
  parsePackage,
  readDedocsFile,
  serializePackage,
  sha256,
  writeDedocsFile,
};
