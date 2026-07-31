'use strict';

const path = require('path');

const { compileFile } = require('./compile');

/**
 * Compiles a .pspt file to a temporary .gen.js and executes its `build()`
 * export against the requested output path.
 * @param {string} psptFile
 * @param {string} outputPath - final .docx/.xlsx path.
 */
async function buildFile(psptFile, outputPath) {
  const { outPath: genPath, docType } = compileFile(psptFile);
  delete require.cache[require.resolve(path.resolve(genPath))];
  const generated = require(path.resolve(genPath));
  const resolvedOutput = outputPath || (docType === 'xlsx' ? 'output.xlsx' : 'output.docx');
  await generated.build(resolvedOutput);
  return { outputPath: resolvedOutput, docType, genPath };
}

module.exports = { buildFile };
