'use strict';

const { generate } = require('./codegen');
const { tokenize, LexError } = require('./lexer');
const { parse, ParseError } = require('./parser');

/**
 * Compiles .pspt source text into generated JS source text.
 * @param {string} source
 * @param {string} [sourceFileLabel] - used in the generated file's header comment.
 * @returns {{ js: string, docType: string, title: string }}
 */
function compile(source, sourceFileLabel = '<source>') {
  const tokens = tokenize(source);
  const program = parse(tokens);
  const js = generate(program, sourceFileLabel);
  return { js, docType: program.doc.docType, title: program.doc.title, program };
}

module.exports = { compile, tokenize, parse, generate, LexError, ParseError };
