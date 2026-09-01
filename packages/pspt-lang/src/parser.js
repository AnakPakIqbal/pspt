'use strict';
/**
 * Hand-written recursive-descent parser for the .pspt DSL.
 *
 * Grammar (informal EBNF; no loops/conditionals/imports by design):
 *
 *   program       := doc-decl (calendar-decl)? (hardware-decl)? statement*
 *   doc-decl      := 'doc' string 'type' '=' ident
 *   calendar-decl := 'calendar' date '..' date
 *   hardware-decl := 'hardware' ':' bool          (legacy; warns at codegen)
 *
 *   statement     := section | part | group | callout | table-decl | rows-decl | field
 *
 *   part          := 'part' name '{' section-body* '}'      (type=master only)
 *   section       := 'section' name '{' section-body* '}'
 *   section-body  := field | table-decl | rows-decl | list-decl | object-decl | section
 *   field         := name ':' value
 *   list-decl     := 'list' name '{' item* '}'
 *   item          := 'item' string ('{' item* '}')?
 *   object-decl   := 'object' name '{' (name ':' value ','?)* '}'
 *
 *   table-decl    := 'table' name '{' column* '}'
 *   column        := name ':' string weight?
 *   weight        := ident            (e.g. 'w1', 'w1.5' — parsed from ident text)
 *   rows-decl     := 'rows' name '[' row (',' row)* ','? ']'
 *   row           := '{' kv (',' kv)* ','? '}'
 *   kv            := name ':' value
 *
 *   name          := ident | keyword
 *       Data keys come from the SDKs, which know nothing about these reserved
 *       words — `setNamingConventions` really does take an `item` key. So a
 *       keyword is treated as a plain name wherever a name is expected, and a
 *       `<keyword>: <value>` line is always read as a field rather than a block.
 *
 *   group         := 'group' string ('color' '=' ident)? '{' (task|callout)* '}'
 *   task          := 'task' string kvpair*
 *   callout       := 'callout' date string ('color' '=' ident)?
 *   kvpair        := ident '=' value
 *
 *   value         := string | number | date | bool | null | ident
 *                    | '[' (value (',' value)* ','?)? ']'      (list value)
 *                    | '{' (name ':' value ','?)* '}'          (inline object value)
 *
 * Every parse error thrown includes a 1-based line and column, since .pspt
 * authors include non-developers and an AI agent as well as engineers.
 */

const { KEYWORDS } = require('./lexer');
const { DOC_TYPES, suggest } = require('./parts');

class ParseError extends Error {
  constructor(message, line, col) {
    super(`Line ${line}, col ${col}: ${message}`);
    this.name = 'PsptParseError';
    this.line = line;
    this.col = col;
  }
}

class Parser {
  constructor(tokens) {
    this.tokens = tokens;
    this.pos = 0;
  }

  peek(offset = 0) {
    return this.tokens[this.pos + offset];
  }
  current() {
    return this.peek();
  }

  advance() {
    const t = this.tokens[this.pos];
    if (this.pos < this.tokens.length - 1) this.pos++;
    return t;
  }

  isTokenType(type) {
    return this.current().type === type;
  }

  expect(type, context) {
    const t = this.current();
    if (t.type !== type) {
      const contextSuffix = context ? ` (${context})` : '';
      throw new ParseError(
        `Expected '${type}'${contextSuffix} but found ${this.describe(t)}`,
        t.line,
        t.col,
      );
    }
    return this.advance();
  }

  /**
   * Accepts a *name* — an identifier, or a keyword used as one.
   *
   * Data keys come from the SDKs, which know nothing about this grammar's
   * reserved words: `setNamingConventions` really does take an `item` key, and
   * a table column really can be called `type` or `status`. Reserving those
   * names would make whole setters unreachable, so anywhere a name is
   * unambiguously expected, a keyword is just a name.
   */
  expectName(context) {
    const t = this.current();
    if (t.type === 'ident' || KEYWORDS.has(t.type)) {
      this.advance();
      return { value: t.value, line: t.line, col: t.col };
    }
    const contextSuffix = context ? ` (${context})` : '';
    throw new ParseError(
      `Expected a name${contextSuffix} but found ${this.describe(t)}`,
      t.line,
      t.col,
    );
  }

  /** True when the current token starts a `<name>: <value>` field. */
  isFieldAhead() {
    const t = this.current();
    return (
      (t.type === 'ident' || KEYWORDS.has(t.type)) && this.peek(1) && this.peek(1).type === ':'
    );
  }

  describe(t) {
    if (t.type === 'eof') return 'end of file';
    if (t.type === 'string') return `string "${t.value}"`;
    if (t.type === 'ident') return `identifier '${t.value}'`;
    return `'${t.value}'`;
  }

  parseValue() {
    const t = this.current();
    if (t.type === 'string') {
      this.advance();
      return { kind: 'string', value: t.value };
    }
    if (t.type === 'number') {
      this.advance();
      return { kind: 'number', value: t.value };
    }
    if (t.type === 'date') {
      this.advance();
      return { kind: 'date', value: t.value };
    }
    if (t.type === 'true') {
      this.advance();
      return { kind: 'bool', value: true };
    }
    if (t.type === 'false') {
      this.advance();
      return { kind: 'bool', value: false };
    }
    if (t.type === 'null') {
      this.advance();
      return { kind: 'null', value: null };
    }
    if (t.type === 'ident') {
      this.advance();
      return { kind: 'ident', value: t.value };
    }
    // A value may itself be a list, or an inline object. Some setters take rows
    // whose cells are lists of lines (an endpoint's request/response body) or
    // lists of objects (an entity's columns); without these two forms those
    // setters are simply unreachable from the DSL.
    if (t.type === '[') return this.parseArrayValue();
    if (t.type === '{') return this.parseObjectValue();
    throw new ParseError(
      `Expected a value (string/number/date/true/false/null, a [list], or an {object}) but found ${this.describe(t)}`,
      t.line,
      t.col,
    );
  }

  /** `[ value, value, ... ]` — a list value, trailing comma allowed. */
  parseArrayValue() {
    const start = this.expect('[');
    const items = [];
    while (!this.isTokenType(']')) {
      if (this.isTokenType('eof')) {
        throw new ParseError(
          "Unterminated list value — missing closing ']'",
          start.line,
          start.col,
        );
      }
      items.push(this.parseValue());
      if (this.isTokenType(',')) this.advance();
      else break;
    }
    this.expect(']', "a closing ']' to end the list value");
    return { kind: 'array', items, line: start.line };
  }

  /** `{ key: value, ... }` — an inline object value, trailing comma allowed. */
  parseObjectValue() {
    const start = this.expect('{');
    const fields = [];
    while (!this.isTokenType('}')) {
      if (this.isTokenType('eof')) {
        throw new ParseError(
          "Unterminated object value — missing closing '}'",
          start.line,
          start.col,
        );
      }
      const keyTok = this.expectName('a key inside the object value');
      this.expect(':', `a ':' after object key '${keyTok.value}'`);
      fields.push({ key: keyTok.value, value: this.parseValue() });
      if (this.isTokenType(',')) this.advance();
    }
    this.expect('}');
    return { kind: 'object', fields, line: start.line };
  }

  /** Parses `program := doc-decl statement*` and returns the AST root. */
  parseProgram() {
    const doc = this.parseDocDecl();
    let calendar = null;
    let hardware = null;
    const statements = [];

    if (this.isTokenType('calendar')) calendar = this.parseCalendarDecl();
    if (this.isTokenType('hardware')) hardware = this.parseHardwareDecl();

    while (!this.isTokenType('eof')) {
      statements.push(this.parseStatement());
    }

    return { type: 'Program', doc, calendar, hardware, statements };
  }

  parseDocDecl() {
    const start = this.expect('doc', 'every .pspt file must start with `doc "Title" type=<part>`');
    const title = this.expect('string', 'the document title').value;
    let docType = null;
    if (this.isTokenType('type')) {
      this.advance();
      this.expect('=', '`type=<part>`, e.g. `type=projectBrief`');
      const t = this.expect('ident', 'a document type');
      docType = t.value;
      if (!DOC_TYPES.includes(docType)) {
        // `docx` was the old single-document ProductSpecSDK. That SDK is gone,
        // so silently remapping it would produce a plausible-looking but wrong
        // document — name the replacement instead.
        if (docType === 'docx') {
          throw new ParseError(
            "Doc type 'docx' no longer exists — it referred to the single-document ProductSpecSDK, " +
              'which has been replaced by the 15-Part documentation suite. Pick the Part this file ' +
              `authors (${[...DOC_TYPES].filter((d) => d !== 'xlsx' && d !== 'master').join(', ')}), ` +
              'or `type=master` to author every Part in one file.',
            t.line,
            t.col,
          );
        }
        const hint = suggest(docType, DOC_TYPES);
        throw new ParseError(
          `Unknown doc type '${docType}'${hint ? ` — did you mean '${hint}'?` : '.'} Valid types: ${DOC_TYPES.join(', ')}`,
          t.line,
          t.col,
        );
      }
    }
    if (docType === null) {
      throw new ParseError(
        `Missing 'type=' on the doc declaration. There is no default any more — say which document this is, e.g. \`doc ${JSON.stringify(title)} type=projectBrief\`. Valid types: ${DOC_TYPES.join(', ')}`,
        start.line,
        start.col,
      );
    }
    return { type: 'DocDecl', title, docType, line: start.line };
  }

  parseCalendarDecl() {
    const start = this.expect('calendar');
    const from = this.expect('date', 'calendar start date, e.g. 2026-01-01').value;
    this.expect('..', "'..' between the two calendar dates");
    const until = this.expect('date', 'calendar end date, e.g. 2026-06-30').value;
    return { type: 'CalendarDecl', from, to: until, line: start.line };
  }

  parseHardwareDecl() {
    const start = this.expect('hardware');
    this.expect(':', "'hardware: true' or 'hardware: false'");
    const v = this.parseValue();
    if (v.kind !== 'bool') {
      throw new ParseError(
        `Expected true/false after 'hardware:' but found ${v.kind} value`,
        start.line,
        start.col,
      );
    }
    return { type: 'HardwareDecl', enabled: v.value, line: start.line };
  }

  parseStatement() {
    const t = this.current();
    if (this.isFieldAhead()) return this.parseField();
    switch (t.type) {
      case 'section':
        return this.parseSection();
      case 'part':
        return this.parsePartDecl();
      case 'group':
        return this.parseGroup();
      case 'callout':
        return this.parseCallout();
      case 'table':
        return this.parseTableDecl();
      case 'rows':
        return this.parseRowsDecl();
      case 'ident':
        return this.parseField();
      default:
        throw new ParseError(
          `Unexpected ${this.describe(t)} at top level — expected 'section', 'part', 'group', 'callout', 'table', 'rows', or a field name`,
          t.line,
          t.col,
        );
    }
  }

  parseSection() {
    const start = this.expect('section');
    const name = this.expect('ident', 'a section name, e.g. `section overview`').value;
    this.expect('{', `an opening '{' to start the body of section '${name}'`);
    const body = [];
    while (!this.isTokenType('}')) {
      if (this.isTokenType('eof')) {
        throw new ParseError(
          `Unterminated section '${name}' — missing closing '}'`,
          start.line,
          start.col,
        );
      }
      body.push(this.parseSectionBody());
    }
    this.expect('}');
    return { type: 'Section', name, body, line: start.line };
  }

  /**
   * `part <key> { ... }` — only meaningful in a `type=master` file, where each
   * block fills in one Part of the assembled document. The body is the same
   * grammar as a section body, so a Part can be lifted into its own
   * `type=<key>` file (or back) without rewriting its contents.
   */
  parsePartDecl() {
    const start = this.expect('part');
    const nameTok = this.expect('ident', 'a Part key, e.g. `part projectBrief`');
    this.expect('{', `an opening '{' to start the body of part '${nameTok.value}'`);
    const body = [];
    while (!this.isTokenType('}')) {
      if (this.isTokenType('eof')) {
        throw new ParseError(
          `Unterminated part '${nameTok.value}' — missing closing '}'`,
          start.line,
          start.col,
        );
      }
      body.push(this.parseSectionBody());
    }
    this.expect('}');
    return { type: 'Part', name: nameTok.value, body, line: start.line, col: nameTok.col };
  }

  parseSectionBody() {
    const t = this.current();
    // A keyword followed by ':' is a data key, not the start of a block.
    if (this.isFieldAhead()) return this.parseField();
    switch (t.type) {
      case 'table':
        return this.parseTableDecl();
      case 'rows':
        return this.parseRowsDecl();
      case 'list':
        return this.parseListDecl();
      case 'object':
        return this.parseObjectDecl();
      case 'section':
        return this.parseSection();
      case 'ident':
        return this.parseField();
      default:
        throw new ParseError(
          `Unexpected ${this.describe(t)} inside a section — expected a field, 'table', 'rows', 'list', or 'object'`,
          t.line,
          t.col,
        );
    }
  }

  /**
   * `object <name> { key: value ... }` — the payload shape for setters that
   * take a single object rather than a scalar or a row array (setMetadata,
   * setScope, setHeaderFooterLabels, ...). Commas between entries are optional.
   */
  parseObjectDecl() {
    const start = this.expect('object');
    const name = this.expect('ident', 'an object name, e.g. `object metadata`').value;
    this.expect('{', `an opening '{' to start object '${name}'`);
    const fields = [];
    while (!this.isTokenType('}')) {
      if (this.isTokenType('eof')) {
        throw new ParseError(
          `Unterminated object '${name}' — missing closing '}'`,
          start.line,
          start.col,
        );
      }
      const keyTok = this.expectName(`a key inside object '${name}'`);
      this.expect(':', `a ':' after object key '${keyTok.value}'`);
      const value = this.parseValue();
      fields.push({ key: keyTok.value, value, line: keyTok.line });
      if (this.isTokenType(',')) this.advance();
    }
    this.expect('}');
    return { type: 'ObjectBlock', name, fields, line: start.line };
  }

  parseField() {
    const nameTok = this.expectName('a field name');
    this.expect(':', `a ':' after field name '${nameTok.value}'`);
    // image field: `image key: "path"` handled uniformly — the codegen
    // decides docx image-vs-text handling based on key name convention
    // (keys ending in ImagePath / logoImagePath / diagramImagePath).
    const value = this.parseValue();
    return { type: 'Field', name: nameTok.value, value, line: nameTok.line };
  }

  parseListDecl() {
    const start = this.expect('list');
    const name = this.expect('ident', 'a list name, e.g. `list measures`').value;
    this.expect('{', `an opening '{' to start list '${name}'`);
    const items = [];
    while (!this.isTokenType('}')) {
      if (this.isTokenType('eof'))
        throw new ParseError(
          `Unterminated list '${name}' — missing closing '}'`,
          start.line,
          start.col,
        );
      items.push(this.parseItem());
    }
    this.expect('}');
    return { type: 'List', name, items, line: start.line };
  }

  parseItem() {
    const start = this.expect('item', 'an \'item "..."\' entry inside a list');
    const title = this.expect('string', 'the item text/title').value;
    const children = [];
    if (this.isTokenType('{')) {
      this.advance();
      while (!this.isTokenType('}')) {
        if (this.isTokenType('eof'))
          throw new ParseError(
            `Unterminated nested item block for '${title}'`,
            start.line,
            start.col,
          );
        const child = this.parseItem();
        children.push(child);
      }
      this.expect('}');
    }
    return { type: 'Item', title, children, line: start.line };
  }

  parseTableDecl() {
    const start = this.expect('table');
    const name = this.expect('ident', 'a table name, e.g. `table features`').value;
    this.expect('{', `an opening '{' to start table '${name}'`);
    const columns = [];
    while (!this.isTokenType('}')) {
      if (this.isTokenType('eof'))
        throw new ParseError(
          `Unterminated table '${name}' — missing closing '}'`,
          start.line,
          start.col,
        );
      columns.push(this.parseColumn());
    }
    this.expect('}');
    return { type: 'Table', name, columns, line: start.line };
  }

  parseColumn() {
    const keyTok = this.expectName('a column key, e.g. `name:`');
    this.expect(':', `a ':' after column key '${keyTok.value}'`);
    const header = this.expect('string', 'a column header string, e.g. "Feature Name"').value;
    let weight = 1;
    if (this.isTokenType('ident') && /^w\d+(\.\d+)?$/.test(this.current().value)) {
      const wTok = this.advance();
      weight = Number(wTok.value.slice(1));
    }
    return { key: keyTok.value, header, weight };
  }

  parseRowsDecl() {
    const start = this.expect('rows');
    const name = this.expect('ident', 'the table name these rows belong to').value;
    this.expect('[', `an opening '[' to start the row list for '${name}'`);
    const rows = [];
    while (!this.isTokenType(']')) {
      if (this.isTokenType('eof'))
        throw new ParseError(
          `Unterminated rows list for '${name}' — missing closing ']'`,
          start.line,
          start.col,
        );
      rows.push(this.parseRow());
      if (this.isTokenType(',')) {
        this.advance();
      } else break;
    }
    this.expect(']', `a closing ']' to end the row list for '${name}'`);
    return { type: 'Rows', name, rows, line: start.line };
  }

  parseRow() {
    const start = this.expect(
      '{',
      'an opening \'{\' for a table row, e.g. `{ name: "Auth", status: "Done" }`',
    );
    const fields = [];
    while (!this.isTokenType('}')) {
      if (this.isTokenType('eof'))
        throw new ParseError(
          "Unterminated row object — missing closing '}'",
          start.line,
          start.col,
        );
      const keyTok = this.expectName('a row field key');
      this.expect(':', `a ':' after row field key '${keyTok.value}'`);
      const value = this.parseValue();
      fields.push({ key: keyTok.value, value });
      if (this.isTokenType(',')) this.advance();
    }
    this.expect('}');
    return { type: 'Row', fields, line: start.line };
  }

  parseGroup() {
    const start = this.expect('group');
    const title = this.expect('string', 'the group/section title, e.g. `group "Backend"`').value;
    let bannerColor = null;
    let rowColor = null;
    while (this.isTokenType('color') || this.isTokenType('ident')) {
      const t = this.current();
      const label = t.type === 'color' ? 'color' : t.value;
      if (t.type !== 'color' && label !== 'rowColor' && label !== 'bannerColor') break;
      this.advance();
      this.expect('=', `a '=' after '${label}'`);
      const v = this.expect('ident', 'a color name (e.g. blue) or hex').value;
      if (label === 'color' || label === 'bannerColor') bannerColor = v;
      else rowColor = v;
    }
    this.expect('{', `an opening '{' to start the body of group '${title}'`);
    const tasks = [];
    while (!this.isTokenType('}')) {
      if (this.isTokenType('eof'))
        throw new ParseError(
          `Unterminated group '${title}' — missing closing '}'`,
          start.line,
          start.col,
        );
      if (this.isTokenType('task')) tasks.push(this.parseTask());
      else {
        const t = this.current();
        throw new ParseError(
          `Unexpected ${this.describe(t)} inside group '${title}' — expected 'task'`,
          t.line,
          t.col,
        );
      }
    }
    this.expect('}');
    return { type: 'Group', title, bannerColor, rowColor, tasks, line: start.line };
  }

  parseTask() {
    const start = this.expect('task');
    const name = this.expect('string', 'the task name, e.g. `task "Auth API"`').value;
    const attrs = {};
    while (this.isTokenType('ident')) {
      const keyTok = this.advance();
      this.expect('=', `a '=' after task attribute '${keyTok.value}'`);
      const value = this.parseValue();
      attrs[keyTok.value] = value;
    }
    return { type: 'Task', name, attrs, line: start.line };
  }

  parseCallout() {
    const start = this.expect('callout');
    const date = this.expect(
      'date',
      'a date for the callout, e.g. `callout 2026-07-04 "US Holiday"`',
    ).value;
    const text = this.expect('string', 'the callout text').value;
    let color = null;
    if (this.isTokenType('color')) {
      this.advance();
      this.expect('=', "a '=' after 'color'");
      color = this.expect('ident', 'a color name or hex').value;
    }
    return { type: 'Callout', date, text, color, line: start.line };
  }
}

function parse(tokens) {
  const parser = new Parser(tokens);
  return parser.parseProgram();
}

module.exports = { parse, Parser, ParseError };
