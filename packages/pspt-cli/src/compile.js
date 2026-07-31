'use strict';

const fs = require('fs');
const path = require('path');
const { compile } = require('pspt-lang');

/**
 * Reads a .pspt file, compiles it, and writes <file>.gen.js next to it
 * (or to an explicit outPath).
 * @param {string} psptFile
 * @param {string} [outPath]
 * @returns {{ outPath: string, docType: string, title: string }}
 */
function compileFile(psptFile, outPath) {
  const source = fs.readFileSync(psptFile, 'utf8');
  const label = path.basename(psptFile);
  let result;
  try {
    result = compile(source, label);
  } catch (err) {
    if (err.name === 'PsptLexError' || err.name === 'PsptParseError') {
      throw new Error(`${label}: ${err.message}`, { cause: err });
    }
    throw err;
  }
  const genPath = outPath || psptFile.replace(/\.pspt$/, '') + '.gen.js';
  fs.writeFileSync(genPath, result.js);
  return { outPath: genPath, docType: result.docType, title: result.title };
}

module.exports = { compileFile };
