//! The formula tokenizer.

use super::FormulaError;
use super::reference::{
    Area, NameRef, Part, Reference, SheetPrefix, classify_part, join_parts,
};

/// One piece of a formula. `text` is exactly the source text it came from
/// (until a rewrite replaces it), so concatenating the texts gives the
/// formula back.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Token {
    pub kind: TokenKind,
    pub text: String,
}

/// What a [`Token`] is.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TokenKind {
    /// The leading `=`, or the `{=` opening an array-formula wrapper.
    Start,
    /// The `}` closing a `{=...}` wrapper.
    End,
    /// Spaces, tabs and line breaks.
    Whitespace,
    /// Whitespace acting as the intersection operator (`A1:C3 B2:D4`).
    Intersection,
    /// `1`, `1.5`, `.5`, `1E+3`.
    Number,
    /// A string literal with its quotes: `"say ""hi"""`.
    Text,
    /// `TRUE` / `FALSE` (any case).
    Bool,
    /// An error literal: `#REF!`, `#N/A`, `#DIV/0!`, `#GETTING_DATA`...
    Error,
    /// `+ - * / ^ & = <> < > <= >= % :` and the implicit intersection `@`.
    Operator,
    /// The spill operator `#` after a reference (`A1#`).
    Spill,
    /// A function name; the next token is its `(`. May look like a cell
    /// (`LOG10`) or carry a prefix (`_xlfn._xlws.FILTER`).
    Function,
    OpenParen,
    CloseParen,
    /// `,` between function arguments.
    ArgSeparator,
    /// `,` as the union operator (outside function arguments, e.g.
    /// `(A1:A2,A4:A5)` or `Sheet1!$A:$A,Sheet1!$1:$1`).
    Union,
    /// `{` starting an array constant.
    ArrayOpen,
    /// `}` ending an array constant.
    ArrayClose,
    /// `,` between array constant columns.
    ArrayColSep,
    /// `;` between array constant rows.
    ArrayRowSep,
    /// A cell, range, whole-column or whole-row reference with its optional
    /// sheet prefix, or `Sheet2!#REF!`.
    Reference(Reference),
    /// A defined name (with its optional sheet prefix).
    Name(NameRef),
    /// A structured table reference: `Table1[Col]`, `[@Col]`,
    /// `Table1[[#This Row],[A]:[B]]`.
    StructuredRef,
    /// Anything else (kept verbatim so the formula still round-trips).
    Unknown,
}

/// Split a formula into tokens. See the module docs for what is recognised;
/// only unterminated strings, quoted sheet names and `[` brackets are errors.
pub fn tokenize(formula: &str) -> Result<Vec<Token>, FormulaError> {
    let chars: Vec<(usize, char)> = formula.char_indices().collect();
    let n = chars.len();
    let mut lexer = Lexer {
        src: formula,
        chars,
        pos: 0,
        end: n,
        tokens: Vec::new(),
        parens: Vec::new(),
        array_depth: 0,
    };
    let wrapped = n >= 3 && formula.starts_with("{=") && formula.ends_with('}');
    if wrapped {
        lexer.push(TokenKind::Start, 0, 2);
        lexer.end = n - 1;
    } else if formula.starts_with('=') {
        lexer.push(TokenKind::Start, 0, 1);
    }
    lexer.run()?;
    if wrapped {
        lexer.end = n;
        lexer.push(TokenKind::End, n - 1, n);
    }
    mark_intersections(&mut lexer.tokens);
    Ok(lexer.tokens)
}

/// Join tokens back into formula text.
pub fn render(tokens: &[Token]) -> String {
    tokens.iter().map(|token| token.text.as_str()).collect()
}

/// Characters that can appear in names, unquoted sheet names and A1
/// references. Non-ASCII characters (accented letters, CJK, emoji) only ever
/// occur in names, so they all count.
fn is_word_char(c: char) -> bool {
    c.is_ascii_alphanumeric()
        || matches!(c, '_' | '.' | '\\' | '?' | '$')
        || (!c.is_ascii() && !c.is_whitespace())
}

/// Characters that can start a name or an unquoted sheet name.
fn is_name_start(c: char) -> bool {
    c.is_ascii_alphabetic() || c == '_' || c == '\\' || (!c.is_ascii() && !c.is_whitespace())
}

fn ends_operand(kind: &TokenKind) -> bool {
    matches!(
        kind,
        TokenKind::Reference(_)
            | TokenKind::Name(_)
            | TokenKind::StructuredRef
            | TokenKind::CloseParen
            | TokenKind::Spill
    )
}

fn starts_operand(kind: &TokenKind) -> bool {
    matches!(
        kind,
        TokenKind::Reference(_)
            | TokenKind::Name(_)
            | TokenKind::StructuredRef
            | TokenKind::Function
            | TokenKind::OpenParen
    )
}

/// Whitespace between two reference operands is the intersection operator.
fn mark_intersections(tokens: &mut [Token]) {
    for i in 1..tokens.len().saturating_sub(1) {
        if tokens[i].kind == TokenKind::Whitespace
            && ends_operand(&tokens[i - 1].kind)
            && starts_operand(&tokens[i + 1].kind)
        {
            tokens[i].kind = TokenKind::Intersection;
        }
    }
}

struct Lexer<'a> {
    src: &'a str,
    /// (byte offset, char) of every character.
    chars: Vec<(usize, char)>,
    /// Current char index.
    pos: usize,
    /// Char index where scanning stops (before the `}` of a `{=...}` wrapper).
    end: usize,
    tokens: Vec<Token>,
    /// One entry per open `(`: true if it opened a function call.
    parens: Vec<bool>,
    /// Nesting depth of `{...}` array constants.
    array_depth: usize,
}

impl<'a> Lexer<'a> {
    fn at(&self, i: usize) -> Option<char> {
        if i < self.end { Some(self.chars[i].1) } else { None }
    }

    fn byte(&self, i: usize) -> usize {
        self.chars.get(i).map_or(self.src.len(), |&(byte, _)| byte)
    }

    fn slice(&self, from: usize, to: usize) -> &'a str {
        &self.src[self.byte(from)..self.byte(to)]
    }

    fn error(&self, message: &str, at: usize) -> FormulaError {
        FormulaError::new(message, self.byte(at))
    }

    /// Emit chars `from..to` as one token and continue after it.
    fn push(&mut self, kind: TokenKind, from: usize, to: usize) {
        let text = self.slice(from, to).to_string();
        self.tokens.push(Token { kind, text });
        self.pos = to;
    }

    fn scan_while(&self, from: usize, keep: impl Fn(char) -> bool) -> usize {
        let mut i = from;
        while let Some(c) = self.at(i)
            && keep(c)
        {
            i += 1;
        }
        i
    }

    fn scan_word(&self, from: usize) -> usize {
        self.scan_while(from, is_word_char)
    }

    fn scan_digits(&self, from: usize) -> usize {
        self.scan_while(from, |c| c.is_ascii_digit())
    }

    fn run(&mut self) -> Result<(), FormulaError> {
        while let Some(c) = self.at(self.pos) {
            let start = self.pos;
            let in_array = self.array_depth > 0;
            match c {
                c if c.is_whitespace() => {
                    let end = self.scan_while(start, char::is_whitespace);
                    self.push(TokenKind::Whitespace, start, end);
                }
                '"' => self.string(start)?,
                '\'' if !in_array => self.quoted_prefix(start)?,
                '[' if !in_array => self.bracket(start)?,
                '{' => {
                    self.array_depth += 1;
                    self.push(TokenKind::ArrayOpen, start, start + 1);
                }
                '}' if in_array => {
                    self.array_depth -= 1;
                    self.push(TokenKind::ArrayClose, start, start + 1);
                }
                ',' => {
                    let kind = if in_array {
                        TokenKind::ArrayColSep
                    } else if self.parens.last() == Some(&true) {
                        TokenKind::ArgSeparator
                    } else {
                        TokenKind::Union
                    };
                    self.push(kind, start, start + 1);
                }
                ';' if in_array => self.push(TokenKind::ArrayRowSep, start, start + 1),
                '(' => {
                    let call = self.tokens.last().is_some_and(|t| t.kind == TokenKind::Function);
                    self.parens.push(call);
                    self.push(TokenKind::OpenParen, start, start + 1);
                }
                ')' => {
                    self.parens.pop();
                    self.push(TokenKind::CloseParen, start, start + 1);
                }
                '#' => self.hash(start, in_array)?,
                '0'..='9' => self.number_or_rows(start, in_array)?,
                '.' if self.at(start + 1).is_some_and(|c| c.is_ascii_digit()) => self.number(start),
                '$' if !in_array => self.word(start)?,
                c if is_name_start(c) => {
                    if in_array {
                        self.array_word(start);
                    } else {
                        self.word(start)?;
                    }
                }
                '+' | '-' | '*' | '/' | '^' | '&' | '=' | '%' | '@' | ':' => {
                    self.push(TokenKind::Operator, start, start + 1);
                }
                '<' => {
                    let len = if matches!(self.at(start + 1), Some('=' | '>')) { 2 } else { 1 };
                    self.push(TokenKind::Operator, start, start + len);
                }
                '>' => {
                    let len = if self.at(start + 1) == Some('=') { 2 } else { 1 };
                    self.push(TokenKind::Operator, start, start + len);
                }
                _ => self.push(TokenKind::Unknown, start, start + 1),
            }
        }
        Ok(())
    }

    /// `"..."` with `""` as an escaped quote.
    fn string(&mut self, start: usize) -> Result<(), FormulaError> {
        let mut i = start + 1;
        loop {
            match self.at(i) {
                None => return Err(self.error("unterminated string", start)),
                Some('"') if self.at(i + 1) == Some('"') => i += 2,
                Some('"') => break,
                Some(_) => i += 1,
            }
        }
        self.push(TokenKind::Text, start, i + 1);
        Ok(())
    }

    /// `'Sheet name'!...` with `''` as an escaped apostrophe.
    fn quoted_prefix(&mut self, start: usize) -> Result<(), FormulaError> {
        let mut i = start + 1;
        let mut content = String::new();
        loop {
            match self.at(i) {
                None => return Err(self.error("unterminated quoted sheet name", start)),
                Some('\'') if self.at(i + 1) == Some('\'') => {
                    content.push('\'');
                    i += 2;
                }
                Some('\'') => break,
                Some(c) => {
                    content.push(c);
                    i += 1;
                }
            }
        }
        let close = i + 1;
        if self.at(close) != Some('!') {
            self.push(TokenKind::Unknown, start, close);
            return Ok(());
        }
        let text = self.slice(start, close + 1).to_string();
        let prefix = SheetPrefix::from_parts(text, &content);
        self.pos = close + 1;
        self.local_part(start, Some(prefix))
    }

    /// Index just after the `]` matching the `[` at `start`. Inside
    /// brackets `'` escapes the next character (`[Col']]`).
    fn scan_bracket(&self, start: usize) -> Result<usize, FormulaError> {
        let mut depth = 0usize;
        let mut i = start;
        loop {
            match self.at(i) {
                None => return Err(self.error("unterminated '['", start)),
                Some('\'') => i += 2,
                Some('[') => {
                    depth += 1;
                    i += 1;
                }
                Some(']') => {
                    depth -= 1;
                    i += 1;
                    if depth == 0 {
                        return Ok(i);
                    }
                }
                Some(_) => i += 1,
            }
        }
    }

    /// If an unquoted sheet name or `First:Last` span starts at `from` and
    /// is followed by `!`, the index of that `!`.
    fn unquoted_sheet(&self, from: usize) -> Option<usize> {
        let first_end = self.scan_word(from);
        if first_end == from {
            return None;
        }
        match self.at(first_end) {
            Some('!') => Some(first_end),
            Some(':') => {
                let last_end = self.scan_word(first_end + 1);
                (last_end > first_end + 1 && self.at(last_end) == Some('!')).then_some(last_end)
            }
            _ => None,
        }
    }

    /// Emit what follows an unquoted prefix spanning `start..=bang`.
    fn prefixed(&mut self, start: usize, bang: usize) -> Result<(), FormulaError> {
        let content = self.slice(start, bang);
        let text = self.slice(start, bang + 1).to_string();
        let prefix = SheetPrefix::from_parts(text, content);
        self.pos = bang + 1;
        self.local_part(start, Some(prefix))
    }

    /// `[` outside an array: an external workbook prefix (`[1]Sheet1!`,
    /// `[1]!Name`) or a table-less structured reference (`[@Col]`).
    fn bracket(&mut self, start: usize) -> Result<(), FormulaError> {
        let close = self.scan_bracket(start)?;
        match self.at(close) {
            Some('!') => self.prefixed(start, close),
            Some(c) if c != '$' && is_word_char(c) => match self.unquoted_sheet(close) {
                Some(bang) => self.prefixed(start, bang),
                None => {
                    self.push(TokenKind::StructuredRef, start, close);
                    Ok(())
                }
            },
            _ => {
                self.push(TokenKind::StructuredRef, start, close);
                Ok(())
            }
        }
    }

    /// A word outside an array: sheet prefix, function, reference, boolean
    /// or name.
    fn word(&mut self, start: usize) -> Result<(), FormulaError> {
        if let Some(bang) = self.unquoted_sheet(start) {
            return self.prefixed(start, bang);
        }
        let end = self.scan_word(start);
        if self.at(end) == Some('(') {
            self.push(TokenKind::Function, start, end);
            return Ok(());
        }
        self.local_part(start, None)
    }

    /// Inside an array constant only booleans are words.
    fn array_word(&mut self, start: usize) {
        let end = self.scan_word(start);
        let word = self.slice(start, end);
        let kind = if word.eq_ignore_ascii_case("TRUE") || word.eq_ignore_ascii_case("FALSE") {
            TokenKind::Bool
        } else {
            TokenKind::Unknown
        };
        self.push(kind, start, end);
    }

    /// The part after an optional sheet prefix (which spans `start..pos`):
    /// `A1`, `A1:B2`, `A:C`, `1:3`, `#REF!`, a name, a structured reference
    /// or a (prefixed) function.
    fn local_part(
        &mut self,
        start: usize,
        prefix: Option<SheetPrefix>,
    ) -> Result<(), FormulaError> {
        let from = self.pos;
        let c = self.at(from);
        if c == Some('#') && prefix.is_some() {
            let end = self.error_literal_end(from);
            if end > from + 1 {
                let reference = Reference { sheet: prefix, area: Area::RefError };
                self.push(TokenKind::Reference(reference), start, end);
                return Ok(());
            }
        }
        if !c.is_some_and(is_word_char) {
            // A prefix with nothing usable after it.
            self.push(TokenKind::Unknown, start, from.max(start + 1));
            return Ok(());
        }
        let end = self.scan_word(from);
        match self.at(end) {
            Some('(') => {
                self.push(TokenKind::Function, start, end);
                return Ok(());
            }
            Some('[') => {
                let close = self.scan_bracket(end)?;
                self.push(TokenKind::StructuredRef, start, close);
                return Ok(());
            }
            _ => {}
        }
        let word = self.slice(from, end);
        if let Some(part) = classify_part(word) {
            if self.at(end) == Some(':') {
                let second_from = end + 1;
                let second_end = self.scan_word(second_from);
                if second_end > second_from
                    && !matches!(self.at(second_end), Some('(' | '[' | '!'))
                    && let Some(second) = classify_part(self.slice(second_from, second_end))
                    && let Some(area) = join_parts(part, second)
                {
                    let reference = Reference { sheet: prefix, area };
                    self.push(TokenKind::Reference(reference), start, second_end);
                    return Ok(());
                }
            }
            if let Part::Cell(cell) = part {
                let reference = Reference { sheet: prefix, area: Area::Cell(cell) };
                self.push(TokenKind::Reference(reference), start, end);
                return Ok(());
            }
        }
        let kind = if prefix.is_none()
            && (word.eq_ignore_ascii_case("TRUE") || word.eq_ignore_ascii_case("FALSE"))
        {
            TokenKind::Bool
        } else if word.contains('$') || !word.starts_with(is_name_start) {
            TokenKind::Unknown
        } else {
            TokenKind::Name(NameRef { sheet: prefix, name: word.to_string() })
        };
        self.push(kind, start, end);
        Ok(())
    }

    /// End of an error literal starting with the `#` at `start`: `#` plus
    /// letters/digits/`_`/`/` plus an optional final `!` or `?`.
    fn error_literal_end(&self, start: usize) -> usize {
        let end = self.scan_while(start + 1, |c| c.is_ascii_alphanumeric() || c == '_' || c == '/');
        if end > start + 1 && matches!(self.at(end), Some('!' | '?')) { end + 1 } else { end }
    }

    /// `#` is the spill operator right after a reference, else an error
    /// literal. `#REF!` directly followed by a reference is the prefix Excel
    /// leaves when the sheet was deleted (`#REF!A1`).
    fn hash(&mut self, start: usize, in_array: bool) -> Result<(), FormulaError> {
        let after_operand = !in_array
            && self.tokens.last().is_some_and(|t| match &t.kind {
                TokenKind::Reference(r) => !r.area.is_error(),
                TokenKind::Name(_) | TokenKind::CloseParen => true,
                _ => false,
            });
        if after_operand {
            self.push(TokenKind::Spill, start, start + 1);
            return Ok(());
        }
        let end = self.error_literal_end(start);
        if end == start + 1 {
            self.push(TokenKind::Unknown, start, end);
            return Ok(());
        }
        if !in_array
            && self.slice(start, end).eq_ignore_ascii_case("#REF!")
            && self.at(end).is_some_and(is_word_char)
        {
            let text = self.slice(start, end).to_string();
            let prefix = SheetPrefix {
                sheet: text.clone(),
                text,
                workbook: None,
                last_sheet: None,
            };
            self.pos = end;
            return self.local_part(start, Some(prefix));
        }
        self.push(TokenKind::Error, start, end);
        Ok(())
    }

    /// A number, or a whole-row reference such as `1:3` / `2:$5`.
    fn number_or_rows(&mut self, start: usize, in_array: bool) -> Result<(), FormulaError> {
        if !in_array {
            let digits_end = self.scan_digits(start);
            if self.at(digits_end) == Some(':') {
                let mut second = digits_end + 1;
                if self.at(second) == Some('$') {
                    second += 1;
                }
                let second_end = self.scan_digits(second);
                if second_end > second
                    && !self.at(second_end).is_some_and(|c| is_word_char(c) || "([!".contains(c))
                    && classify_part(self.slice(start, digits_end)).is_some()
                    && classify_part(self.slice(digits_end + 1, second_end)).is_some()
                {
                    return self.local_part(start, None);
                }
            }
        }
        self.number(start);
        Ok(())
    }

    /// `12`, `1.5`, `.5`, `5.`, `1E+3`, `2.5e-7`.
    fn number(&mut self, start: usize) {
        let mut end = self.scan_digits(start);
        if self.at(end) == Some('.') {
            end = self.scan_digits(end + 1);
        }
        if matches!(self.at(end), Some('e' | 'E')) {
            let mut exp = end + 1;
            if matches!(self.at(exp), Some('+' | '-')) {
                exp += 1;
            }
            let exp_end = self.scan_digits(exp);
            if exp_end > exp {
                end = exp_end;
            }
        }
        self.push(TokenKind::Number, start, end);
    }
}
