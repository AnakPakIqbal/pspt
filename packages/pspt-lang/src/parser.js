'use strict';
/**
 * Hand-written recursive-descent parser for the .pspt DSL.
 *
 * Grammar (informal EBNF; no loops/conditionals/imports by design):
 *
 *   program       := doc-decl (calendar-decl)? (hardware-decl)? statement*
 *   doc-decl      := 'doc' string ('type' '=' ident)?
 *   calendar-decl := 'calendar' date '..' date
 *   hardware-decl := 'hardware' ':' bool
 *
 *   statement     := section | group | callout | table-decl | rows-decl | field
 *
 *   section       := 'section' ident '{' section-body* '}'
 *   section-body  := field | table-decl | rows-decl | list-decl
 *   field         := ident ':' value
 *   list-decl     := 'list' ident '{' item* '}'
 *   item          := 'item' string ('{' item* '}')?
 *
 *   table-decl    := 'table' ident '{' column* '}'
 *   column        := ident ':' string weight?
 *   weight        := ident            (e.g. 'w1', 'w1.5' — parsed from ident text)
 *   rows-decl     := 'rows' ident '[' row (',' row)* ','? ']'
 *   row           := '{' kv (',' kv)* ','? '}'
 *   kv            := ident ':' value
 *
 *   group         := 'group' string ('color' '=' ident)? '{' (task|callout)* '}'
 *   task          := 'task' string kvpair*
 *   callout       := 'callout' date string ('color' '=' ident)?
 *   kvpair        := ident '=' value
 *
 *   value         := string | number | date | bool | null
 *
 * Every parse error thrown includes a 1-based line and column, since .pspt
 * authors include non-developers and an AI agent as well as engineers.
 */

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
    throw new ParseError(
      `Expected a value (string/number/date/true/false/null) but found ${this.describe(t)}`,
      t.line,
      t.col,
    );
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
    const start = this.expect(
      'doc',
      'every .pspt file must start with `doc "Title" type=docx|xlsx`',
    );
    const title = this.expect('string', 'the document title').value;
    let docType = 'docx';
    if (this.isTokenType('type')) {
      this.advance();
      this.expect('=', '`type=docx` or `type=xlsx`');
      const t = this.expect('ident', 'docx or xlsx');
      docType = t.value;
      if (docType !== 'docx' && docType !== 'xlsx') {
        throw new ParseError(
          `Unknown doc type '${docType}' — expected 'docx' or 'xlsx'`,
          t.line,
          t.col,
        );
      }
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
    switch (t.type) {
      case 'section':
        return this.parseSection();
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
          `Unexpected ${this.describe(t)} at top level — expected 'section', 'group', 'callout', 'table', 'rows', or a field name`,
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

  parseSectionBody() {
    const t = this.current();
    switch (t.type) {
      case 'table':
        return this.parseTableDecl();
      case 'rows':
        return this.parseRowsDecl();
      case 'list':
        return this.parseListDecl();
      case 'ident':
        return this.parseField();
      default:
        throw new ParseError(
          `Unexpected ${this.describe(t)} inside a section — expected a field, 'table', 'rows', or 'list'`,
          t.line,
          t.col,
        );
    }
  }

  parseField() {
    const nameTok = this.expect('ident');
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
    const keyTok = this.expect('ident', 'a column key, e.g. `name:`');
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
      const keyTok = this.expect('ident', 'a row field key');
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
