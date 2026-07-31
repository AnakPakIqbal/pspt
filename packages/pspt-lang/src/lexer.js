'use strict';
/**
 * Hand-written lexer for the .pspt DSL.
 *
 * Token kinds: doc, calendar, group, task, callout, section, table, rows,
 * hardware, list, item, image, string, number, date, ident, symbol,
 * newline (insignificant — statements are newline/brace delimited by the
 * parser via explicit tokens, not indentation), eof.
 *
 * Every token carries {type, value, line, col} so the parser can produce
 * exact line-numbered error messages — the primary UX requirement, since
 * .pspt authors include non-developers and an AI agent, not just engineers.
 */

const KEYWORDS = new Set([
  'doc',
  'type',
  'calendar',
  'group',
  'task',
  'callout',
  'section',
  'table',
  'rows',
  'hardware',
  'list',
  'item',
  'color',
  'true',
  'false',
  'null',
]);

class LexError extends Error {
  constructor(message, line, col) {
    super(`Line ${line}, col ${col}: ${message}`);
    this.name = 'PsptLexError';
    this.line = line;
    this.col = col;
  }
}

function isDigit(ch) {
  return ch >= '0' && ch <= '9';
}
function isIdentStart(ch) {
  return /[A-Za-z_]/.test(ch);
}
function isIdentPart(ch) {
  return /\w/.test(ch);
}

/**
 * @param {string} source
 * @returns {Array<{type:string,value:*,line:number,col:number}>}
 */
function tokenize(source) {
  const tokens = [];
  let i = 0;
  let line = 1;
  let col = 1;
  const n = source.length;

  function advance(count = 1) {
    for (let k = 0; k < count; k++) {
      if (source[i] === '\n') {
        line++;
        col = 1;
      } else {
        col++;
      }
      i++;
    }
  }

  function peek(offset = 0) {
    return source[i + offset];
  }

  while (i < n) {
    const ch = peek();

    if (ch === '\n') {
      advance();
      continue;
    }
    if (ch === ' ' || ch === '\t' || ch === '\r') {
      advance();
      continue;
    }

    if (ch === '#' || (ch === '/' && peek(1) === '/')) {
      while (i < n && peek() !== '\n') advance();
      continue;
    }

    const startLine = line;
    const startCol = col;

    if (ch === '"') {
      advance();
      let value = '';
      while (i < n && peek() !== '"') {
        if (peek() === '\\' && (peek(1) === '"' || peek(1) === '\\' || peek(1) === 'n')) {
          const esc = peek(1);
          value += esc === 'n' ? '\n' : esc;
          advance(2);
        } else if (peek() === '\n') {
          throw new LexError('Unterminated string literal (hit end of line)', startLine, startCol);
        } else {
          value += peek();
          advance();
        }
      }
      if (i >= n)
        throw new LexError('Unterminated string literal (hit end of file)', startLine, startCol);
      advance(); // closing quote
      tokens.push({ type: 'string', value, line: startLine, col: startCol });
      continue;
    }

    // date literal: YYYY-MM-DD (must not be followed by more ident chars)
    if (
      isDigit(ch) &&
      /\d{4}-\d{2}-\d{2}/.test(source.slice(i, i + 10)) &&
      !isIdentPart(source[i + 10] || '')
    ) {
      const value = source.slice(i, i + 10);
      advance(10);
      tokens.push({ type: 'date', value, line: startLine, col: startCol });
      continue;
    }

    // Digit-leading hex color (e.g. `1B365D`, `00FF00`): a run of digits
    // immediately followed by ident chars (no whitespace/operator between)
    // is a bare hex token, not a number — colors are written unquoted
    // (`color=1B365D`) and may start with a digit.
    if (isDigit(ch)) {
      let lookahead = i;
      while (lookahead < n && isDigit(source[lookahead])) lookahead++;
      if (lookahead < n && isIdentPart(source[lookahead]) && !isDigit(source[lookahead])) {
        let ident = '';
        while (i < n && isIdentPart(peek())) {
          ident += peek();
          advance();
        }
        tokens.push({ type: 'ident', value: ident, line: startLine, col: startCol });
        continue;
      }
    }

    if (isDigit(ch) || (ch === '-' && isDigit(peek(1)))) {
      let numStr = '';
      if (ch === '-') {
        numStr += '-';
        advance();
      }
      while (i < n && isDigit(peek())) {
        numStr += peek();
        advance();
      }
      if (peek() === '.' && isDigit(peek(1))) {
        numStr += '.';
        advance();
        while (i < n && isDigit(peek())) {
          numStr += peek();
          advance();
        }
      }
      tokens.push({ type: 'number', value: Number(numStr), line: startLine, col: startCol });
      continue;
    }

    if (isIdentStart(ch)) {
      let ident = '';
      while (i < n && isIdentPart(peek())) {
        ident += peek();
        advance();
      }
      // `w2.5`-style weight suffixes: a decimal point immediately followed by
      // more ident chars extends the same token (e.g. `w2` + `.5` -> `w2.5`),
      // so `parseColumn`'s /^w\d+(\.\d+)?$/ check can match it as one ident.
      if (peek() === '.' && isDigit(peek(1))) {
        ident += '.';
        advance();
        while (i < n && isIdentPart(peek())) {
          ident += peek();
          advance();
        }
      }
      if (KEYWORDS.has(ident)) {
        tokens.push({ type: ident, value: ident, line: startLine, col: startCol });
      } else {
        tokens.push({ type: 'ident', value: ident, line: startLine, col: startCol });
      }
      continue;
    }

    if (ch === '.' && peek(1) === '.') {
      advance(2);
      tokens.push({ type: '..', value: '..', line: startLine, col: startCol });
      continue;
    }

    const singleCharSymbols = '{}[]:,=';
    if (singleCharSymbols.includes(ch)) {
      advance();
      tokens.push({ type: ch, value: ch, line: startLine, col: startCol });
      continue;
    }

    throw new LexError(`Unexpected character ${JSON.stringify(ch)}`, startLine, startCol);
  }

  tokens.push({ type: 'eof', value: null, line, col });
  return tokens;
}

module.exports = { tokenize, LexError };
