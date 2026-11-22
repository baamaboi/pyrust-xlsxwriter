//! Tests for the formula engine.

use super::*;

// ---------------------------------------------------------------------------
// Helpers.
// ---------------------------------------------------------------------------

fn label(kind: &TokenKind) -> &'static str {
    match kind {
        TokenKind::Start => "start",
        TokenKind::End => "end",
        TokenKind::Whitespace => "ws",
        TokenKind::Intersection => "isect",
        TokenKind::Number => "num",
        TokenKind::Text => "str",
        TokenKind::Bool => "bool",
        TokenKind::Error => "err",
        TokenKind::Operator => "op",
        TokenKind::Spill => "spill",
        TokenKind::Function => "func",
        TokenKind::OpenParen => "(",
        TokenKind::CloseParen => ")",
        TokenKind::ArgSeparator => "arg",
        TokenKind::Union => "union",
        TokenKind::ArrayOpen => "{",
        TokenKind::ArrayClose => "}",
        TokenKind::ArrayColSep => "acol",
        TokenKind::ArrayRowSep => "arow",
        TokenKind::Reference(_) => "ref",
        TokenKind::Name(_) => "name",
        TokenKind::StructuredRef => "sref",
        TokenKind::Unknown => "?",
    }
}

/// "kind:text" for every token.
fn toks(formula: &str) -> Vec<String> {
    let tokens = tokenize(formula).unwrap_or_else(|e| panic!("{formula:?}: {e}"));
    assert_eq!(render(&tokens), formula, "round trip");
    tokens.iter().map(|t| format!("{}:{}", label(&t.kind), t.text)).collect()
}

/// The single reference in a formula.
fn one_ref(formula: &str) -> Reference {
    let refs = references(formula).unwrap();
    assert_eq!(refs.len(), 1, "{formula:?}: {refs:?}");
    refs.into_iter().next().unwrap()
}

fn cell(row: u32, col: u16, row_abs: bool, col_abs: bool) -> CellRef {
    CellRef { row, col, row_abs, col_abs }
}

fn ins_rows(at: u32, count: u32) -> StructuralEdit {
    StructuralEdit::InsertRows { sheet: "Sheet1".into(), at, count }
}
fn del_rows(at: u32, count: u32) -> StructuralEdit {
    StructuralEdit::DeleteRows { sheet: "Sheet1".into(), at, count }
}
fn ins_cols(at: u16, count: u16) -> StructuralEdit {
    StructuralEdit::InsertCols { sheet: "Sheet1".into(), at, count }
}
fn del_cols(at: u16, count: u16) -> StructuralEdit {
    StructuralEdit::DeleteCols { sheet: "Sheet1".into(), at, count }
}

/// Shift a formula that lives on Sheet1.
fn shift(formula: &str, edit: &StructuralEdit) -> String {
    shift_formula(formula, "Sheet1", edit).unwrap()
}

fn check_shift(edit: &StructuralEdit, cases: &[(&str, &str)]) {
    for (input, expected) in cases {
        assert_eq!(shift(input, edit), *expected, "{input:?} with {edit:?}");
    }
}

// ---------------------------------------------------------------------------
// Tokenizer.
// ---------------------------------------------------------------------------

#[test]
fn leading_equals_is_optional_and_preserved() {
    assert_eq!(toks("=A1+1"), ["start:=", "ref:A1", "op:+", "num:1"]);
    assert_eq!(toks("A1+1"), ["ref:A1", "op:+", "num:1"]);
    assert_eq!(toks("=A1=B1"), ["start:=", "ref:A1", "op:=", "ref:B1"]);
    assert_eq!(toks(""), Vec::<String>::new());
    assert_eq!(toks("="), ["start:="]);
}

#[test]
fn strings_with_doubled_quotes() {
    assert_eq!(toks(r#""say ""hi""""#), [r#"str:"say ""hi""""#]);
    assert_eq!(toks(r#""""#), [r#"str:"""#]);
    assert_eq!(toks(r#""""""#), [r#"str:"""""#]);
    assert_eq!(
        toks(r#"LEN("A1,B2")&A1"#),
        ["func:LEN", "(:(", r#"str:"A1,B2""#, "):)", "op:&", "ref:A1"]
    );
    assert_eq!(toks(r#""it's""#), [r#"str:"it's""#]);
    let err = tokenize(r#"A1&"abc"#).unwrap_err();
    assert_eq!(err.position, 3);
    assert!(tokenize(r#""abc""#).is_ok());
    assert!(tokenize(r#""abc"""#).is_err());
}

#[test]
fn quoted_sheet_names() {
    assert_eq!(toks("'It''s'!A1"), ["ref:'It''s'!A1"]);
    let r = one_ref("'It''s'!A1");
    let prefix = r.sheet.unwrap();
    assert_eq!(prefix.sheet, "It's");
    assert_eq!(prefix.text, "'It''s'!");
    assert!(prefix.is_local_sheet());

    let r = one_ref("SUM('Sheet 1'!$A$1:$B$2)");
    assert_eq!(r.sheet.as_ref().unwrap().sheet, "Sheet 1");
    assert_eq!(r.area, Area::Range(cell(0, 0, true, true), cell(1, 1, true, true)));

    assert_eq!(one_ref("'Sheet!3'!$A$1:$A$5").sheet.unwrap().sheet, "Sheet!3");
    assert_eq!(one_ref("'S''heet''2'!$A$1").sheet.unwrap().sheet, "S'heet'2");
    assert_eq!(one_ref("'1Sheet'!$1:$2").sheet.unwrap().sheet, "1Sheet");
    assert!(tokenize("'Sheet 1!A1").is_err());
    // A quoted name without "!" is not a reference.
    assert_eq!(toks("'abc'"), ["?:'abc'"]);
}

#[test]
fn unquoted_sheet_names() {
    let r = one_ref("Sh.eet.1!$A$1:$A$5");
    assert_eq!(r.sheet.unwrap().sheet, "Sh.eet.1");
    assert_eq!(one_ref("Sheeté!$A$1").sheet.unwrap().sheet, "Sheeté");
    assert_eq!(one_ref("a_x005F_x0041_b!$A$1:$A$5").sheet.unwrap().sheet, "a_x005F_x0041_b");
    assert_eq!(one_ref("日本!A1").sheet.unwrap().sheet, "日本");
    assert_eq!(one_ref("'日本 語'!A1").sheet.unwrap().sheet, "日本 語");
}

#[test]
fn absolute_markers() {
    assert_eq!(one_ref("A1").area, Area::Cell(cell(0, 0, false, false)));
    assert_eq!(one_ref("$A$1").area, Area::Cell(cell(0, 0, true, true)));
    assert_eq!(one_ref("A$1").area, Area::Cell(cell(0, 0, true, false)));
    assert_eq!(one_ref("$A1").area, Area::Cell(cell(0, 0, false, true)));
    assert_eq!(one_ref("$XFD$1048576").area, Area::Cell(cell(1_048_575, 16_383, true, true)));
    assert_eq!(one_ref("a1").area, Area::Cell(cell(0, 0, false, false)));
}

#[test]
fn ranges_are_single_tokens() {
    assert_eq!(toks("A1:B2"), ["ref:A1:B2"]);
    assert_eq!(toks("$A$1:$B$2"), ["ref:$A$1:$B$2"]);
    assert_eq!(toks("SUM(A1:B2)"), ["func:SUM", "(:(", "ref:A1:B2", "):)"]);
    assert_eq!(one_ref("B2:A1").area, Area::Range(cell(1, 1, false, false), cell(0, 0, false, false)));
    // ":" between non-matching parts is an operator.
    assert_eq!(
        toks("A1:INDEX(B:B,2)"),
        ["ref:A1", "op::", "func:INDEX", "(:(", "ref:B:B", "arg:,", "num:2", "):)"]
    );
    assert_eq!(toks("Sheet1!A1:Sheet1!B2"), ["ref:Sheet1!A1", "op::", "ref:Sheet1!B2"]);
    assert_eq!(toks("A1:B2:C3"), ["ref:A1:B2", "op::", "ref:C3"]);
    assert_eq!(toks("Name1:Name2"), ["name:Name1", "op::", "name:Name2"]);
    assert_eq!(toks("A1:A"), ["ref:A1", "op::", "name:A"]);
}

#[test]
fn whole_rows_and_columns() {
    let col = |col, abs| ColRef { col, abs };
    let row = |row, abs| RowRef { row, abs };
    assert_eq!(one_ref("A:C").area, Area::Cols(col(0, false), col(2, false)));
    assert_eq!(one_ref("$A:$A").area, Area::Cols(col(0, true), col(0, true)));
    assert_eq!(one_ref("XFD:XFD").area, Area::Cols(col(16_383, false), col(16_383, false)));
    assert_eq!(one_ref("1:3").area, Area::Rows(row(0, false), row(2, false)));
    assert_eq!(one_ref("$1:$2").area, Area::Rows(row(0, true), row(1, true)));
    assert_eq!(one_ref("2:$5").area, Area::Rows(row(1, false), row(4, true)));
    assert_eq!(toks("Sheet1!$1:$2"), ["ref:Sheet1!$1:$2"]);
    assert_eq!(toks("Sheet1!A:A"), ["ref:Sheet1!A:A"]);
    assert_eq!(toks("SUM(1:1)"), ["func:SUM", "(:(", "ref:1:1", "):)"]);
    // Numbers are not rows.
    assert_eq!(toks("1+2"), ["num:1", "op:+", "num:2"]);
    assert_eq!(toks("1:1048577"), ["num:1", "op::", "num:1048577"]);
}

#[test]
fn functions_that_look_like_cells() {
    assert_eq!(toks("LOG10(100)"), ["func:LOG10", "(:(", "num:100", "):)"]);
    assert_eq!(toks("ATAN2(1,2)"), ["func:ATAN2", "(:(", "num:1", "arg:,", "num:2", "):)"]);
    assert_eq!(
        toks("DAYS360(A1,B1)"),
        ["func:DAYS360", "(:(", "ref:A1", "arg:,", "ref:B1", "):)"]
    );
    assert_eq!(toks("DEC2HEX(A1)")[0], "func:DEC2HEX");
    assert_eq!(toks("_xlfn.STDEV.S(A1:A5)")[0], "func:_xlfn.STDEV.S");
    assert_eq!(
        toks("_xlfn._xlws.FILTER(A1:A5,B1:B5>0)"),
        [
            "func:_xlfn._xlws.FILTER",
            "(:(",
            "ref:A1:A5",
            "arg:,",
            "ref:B1:B5",
            "op:>",
            "num:0",
            "):)"
        ]
    );
    assert_eq!(toks("TRUE()"), ["func:TRUE", "(:(", "):)"]);
    assert_eq!(toks("F.DIST(1,2,TRUE)")[0], "func:F.DIST");
    assert_eq!(
        toks("_xlfn.LET(_xlpm.x,1,_xlpm.x+1)"),
        [
            "func:_xlfn.LET",
            "(:(",
            "name:_xlpm.x",
            "arg:,",
            "num:1",
            "arg:,",
            "name:_xlpm.x",
            "op:+",
            "num:1",
            "):)"
        ]
    );
    assert_eq!(references("LOG10(A1)+ATAN2(B2,C3)").unwrap().len(), 3);
}

#[test]
fn defined_names() {
    for name in ["MyName", "_Name", "Name.With.Dots", "\\Name", "Name?", "R1C1", "R2C3", "RC", "R",
        "ABC", "A1B2", "XFE1", "A1048577", "数量", "AAAA1"]
    {
        assert_eq!(toks(name), [format!("name:{name}")], "{name}");
    }
    assert_eq!(toks("MyName*2"), ["name:MyName", "op:*", "num:2"]);
    let tokens = tokenize("Sheet1!LocalName").unwrap();
    let TokenKind::Name(name) = &tokens[0].kind else { panic!("{tokens:?}") };
    assert_eq!(name.name, "LocalName");
    assert_eq!(name.sheet.as_ref().unwrap().sheet, "Sheet1");
    let tokens = tokenize("'My Sheet'!LocalName").unwrap();
    assert!(matches!(&tokens[0].kind, TokenKind::Name(n) if n.sheet.as_ref().unwrap().sheet == "My Sheet"));
    let tokens = tokenize("[1]!ExtName").unwrap();
    let TokenKind::Name(name) = &tokens[0].kind else { panic!("{tokens:?}") };
    assert_eq!(name.sheet.as_ref().unwrap().workbook.as_deref(), Some("[1]"));
    assert_eq!(name.sheet.as_ref().unwrap().sheet, "");
    // Words with "$" that are not references are not names.
    assert_eq!(toks("$A"), ["?:$A"]);
    assert_eq!(toks("$1"), ["?:$1"]);
}

#[test]
fn structured_references() {
    assert_eq!(toks("Table1[Col]"), ["sref:Table1[Col]"]);
    assert_eq!(toks("Table1[[#This Row],[Col]]"), ["sref:Table1[[#This Row],[Col]]"]);
    assert_eq!(toks("Table1[#All]"), ["sref:Table1[#All]"]);
    assert_eq!(toks("Table1[]"), ["sref:Table1[]"]);
    assert_eq!(toks("[@Col]"), ["sref:[@Col]"]);
    assert_eq!(toks("Table1[@[Col 1]]"), ["sref:Table1[@[Col 1]]"]);
    assert_eq!(
        toks("SUM(Table1[[#This Row],[Column1]:[Column3]])"),
        ["func:SUM", "(:(", "sref:Table1[[#This Row],[Column1]:[Column3]]", "):)"]
    );
    assert_eq!(
        toks("SUBTOTAL(109,[Column']])"),
        ["func:SUBTOTAL", "(:(", "num:109", "arg:,", "sref:[Column']]", "):)"]
    );
    assert_eq!(toks("SUBTOTAL(109,[Column'[])")[4], "sref:[Column'[]");
    assert_eq!(toks("SUBTOTAL(109,[Column''])")[4], "sref:[Column'']");
    assert_eq!(toks("SUBTOTAL(109,[Column'#])")[4], "sref:[Column'#]");
    assert_eq!(toks("Table1[[#Headers],[A]:[B]]"), ["sref:Table1[[#Headers],[A]:[B]]"]);
    assert_eq!(toks("@Table1[Col]"), ["op:@", "sref:Table1[Col]"]);
    assert!(references("Table1[[#This Row],[A1]]").unwrap().is_empty());
    assert!(tokenize("Table1[Col").is_err());
    assert!(tokenize("[[#This Row],[Col]").is_err());
}

#[test]
fn array_constants() {
    assert_eq!(
        toks("{1,2;3,4}"),
        ["{:{", "num:1", "acol:,", "num:2", "arow:;", "num:3", "acol:,", "num:4", "}:}"]
    );
    assert_eq!(toks(r#"{"A1","B2"}"#), ["{:{", r#"str:"A1""#, "acol:,", r#"str:"B2""#, "}:}"]);
    assert_eq!(toks("{TRUE,false}"), ["{:{", "bool:TRUE", "acol:,", "bool:false", "}:}"]);
    assert_eq!(toks("{A1,$B$2}"), ["{:{", "?:A1", "acol:,", "?:$", "?:B$2", "}:}"]);
    assert_eq!(toks("{1,-2.5E+3}"), ["{:{", "num:1", "acol:,", "op:-", "num:2.5E+3", "}:}"]);
    assert_eq!(toks("{#N/A,1}"), ["{:{", "err:#N/A", "acol:,", "num:1", "}:}"]);
    assert_eq!(toks("{1:2}")[2], "op::");
    let refs = references("SUM({1,2},A1)*SUM({A1;B2})").unwrap();
    assert_eq!(refs.len(), 1);
    assert_eq!(
        toks("MATCH(1,{1,2},0)"),
        ["func:MATCH", "(:(", "num:1", "arg:,", "{:{", "num:1", "acol:,", "num:2", "}:}", "arg:,", "num:0", "):)"]
    );
}

#[test]
fn array_formula_wrapper() {
    assert_eq!(
        toks("{=SUM(A1:A2*B1:B2)}"),
        ["start:{=", "func:SUM", "(:(", "ref:A1:A2", "op:*", "ref:B1:B2", "):)", "end:}"]
    );
    assert_eq!(toks("{={1,2}}"), ["start:{=", "{:{", "num:1", "acol:,", "num:2", "}:}", "end:}"]);
    assert_eq!(shift("{=SUM(A5:A6)}", &ins_rows(0, 1)), "{=SUM(A6:A7)}");
}

#[test]
fn error_literals() {
    for error in ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
        "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!",
        "#UNKNOWN!"]
    {
        assert_eq!(toks(error), [format!("err:{error}")], "{error}");
    }
    assert_eq!(
        toks("IFERROR(A1,#N/A)"),
        ["func:IFERROR", "(:(", "ref:A1", "arg:,", "err:#N/A", "):)"]
    );
    assert_eq!(toks("#DIV/0!+1"), ["err:#DIV/0!", "op:+", "num:1"]);
    assert_eq!(toks("#"), ["?:#"]);
}

#[test]
fn ref_errors_with_prefixes() {
    // Excel's form after deleting referenced cells on another sheet.
    let r = one_ref("Sheet2!#REF!");
    assert_eq!(r.area, Area::RefError);
    assert_eq!(r.sheet.unwrap().sheet, "Sheet2");
    assert_eq!(toks("'My Sheet'!#REF!+1"), ["ref:'My Sheet'!#REF!", "op:+", "num:1"]);
    // Excel's form after deleting the referenced sheet.
    let r = one_ref("SUM(#REF!A1:B2)");
    let prefix = r.sheet.unwrap();
    assert!(prefix.is_deleted_sheet());
    assert!(!prefix.is_local_sheet());
    assert_eq!(r.area, Area::Range(cell(0, 0, false, false), cell(1, 1, false, false)));
    assert_eq!(toks("#REF!"), ["err:#REF!"]);
    assert_eq!(toks("#REF!+A1"), ["err:#REF!", "op:+", "ref:A1"]);
}

#[test]
fn booleans() {
    assert_eq!(toks("TRUE"), ["bool:TRUE"]);
    assert_eq!(toks("false"), ["bool:false"]);
    assert_eq!(
        toks("IF(B3,FALSE,TRUE)"),
        ["func:IF", "(:(", "ref:B3", "arg:,", "bool:FALSE", "arg:,", "bool:TRUE", "):)"]
    );
    // A sheet-qualified TRUE is a sheet-level name.
    assert!(matches!(tokenize("Sheet1!TRUE").unwrap()[0].kind, TokenKind::Name(_)));
}

#[test]
fn numbers() {
    for number in ["1", "1.5", ".5", "5.", "1E+3", "1e-3", "1.5E10", "0.98", "1E3", "123456789"] {
        assert_eq!(toks(number), [format!("num:{number}")], "{number}");
    }
    assert_eq!(toks("1E+3*E3"), ["num:1E+3", "op:*", "ref:E3"]);
    assert_eq!(toks("100%"), ["num:100", "op:%"]);
    assert_eq!(toks("A1%"), ["ref:A1", "op:%"]);
    assert_eq!(toks("1E"), ["num:1", "name:E"]);
    assert!(references("1E+3+2E-5").unwrap().is_empty());
}

#[test]
fn operators() {
    assert_eq!(
        toks("A1<>B1"),
        ["ref:A1", "op:<>", "ref:B1"]
    );
    assert_eq!(toks("A1<=1"), ["ref:A1", "op:<=", "num:1"]);
    assert_eq!(toks("A1>=1"), ["ref:A1", "op:>=", "num:1"]);
    assert_eq!(toks("A1<1"), ["ref:A1", "op:<", "num:1"]);
    assert_eq!(toks("2^3&\"x\""), ["num:2", "op:^", "num:3", "op:&", "str:\"x\""]);
    assert_eq!(toks("-A1"), ["op:-", "ref:A1"]);
    assert_eq!(toks("--A1"), ["op:-", "op:-", "ref:A1"]);
    assert_eq!(toks("+A1/-2"), ["op:+", "ref:A1", "op:/", "op:-", "num:2"]);
}

#[test]
fn intersection_and_whitespace() {
    assert_eq!(toks("A1:C3 B2:D4"), ["ref:A1:C3", "isect: ", "ref:B2:D4"]);
    assert_eq!(toks("Name1  Name2"), ["name:Name1", "isect:  ", "name:Name2"]);
    assert_eq!(toks("(A1:B2) (B1:C3)")[3], "isect: ");
    assert_eq!(
        toks("SUM( A1 , B1 )"),
        ["func:SUM", "(:(", "ws: ", "ref:A1", "ws: ", "arg:,", "ws: ", "ref:B1", "ws: ", "):)"]
    );
    assert_eq!(toks("1 + 2"), ["num:1", "ws: ", "op:+", "ws: ", "num:2"]);
    assert_eq!(toks("=A1\n+\r\nB1\t"), ["start:=", "ref:A1", "ws:\n", "op:+", "ws:\r\n", "ref:B1", "ws:\t"]);
    assert_eq!(toks(" A1"), ["ws: ", "ref:A1"]);
}

#[test]
fn union_versus_argument_separator() {
    assert_eq!(
        toks("(Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5)"),
        ["(:(", "ref:Sheet1!$A$1:$A$2", "union:,", "ref:Sheet1!$A$4:$A$5", "):)"]
    );
    assert_eq!(toks("SUM(A1,B1)")[3], "arg:,");
    assert_eq!(
        toks("SUM((A1,B1),C1)"),
        ["func:SUM", "(:(", "(:(", "ref:A1", "union:,", "ref:B1", "):)", "arg:,", "ref:C1", "):)"]
    );
    assert_eq!(toks("Sheet1!$A:$A,Sheet1!$1:$1"), ["ref:Sheet1!$A:$A", "union:,", "ref:Sheet1!$1:$1"]);
    assert_eq!(toks("INDEX((A1:B2,C1:D2),1,1,2)")[4], "union:,");
    assert_eq!(toks("INDEX((A1:B2,C1:D2),1,1,2)")[7], "arg:,");
}

#[test]
fn spill_and_implicit_intersection() {
    assert_eq!(toks("A1#"), ["ref:A1", "spill:#"]);
    assert_eq!(toks("Sheet1!A1#"), ["ref:Sheet1!A1", "spill:#"]);
    assert_eq!(toks("SUM(A1#)"), ["func:SUM", "(:(", "ref:A1", "spill:#", "):)"]);
    assert_eq!(toks("MyName#"), ["name:MyName", "spill:#"]);
    assert_eq!(toks("@A1:A10"), ["op:@", "ref:A1:A10"]);
    assert_eq!(toks("_xlfn.SINGLE(A1:A10)")[0], "func:_xlfn.SINGLE");
    // After an operator "#" starts an error literal.
    assert_eq!(toks("A1+#N/A"), ["ref:A1", "op:+", "err:#N/A"]);
}

#[test]
fn three_d_references() {
    let r = one_ref("SUM(Sheet1:Sheet3!A1)");
    let prefix = r.sheet.unwrap();
    assert_eq!(prefix.sheet, "Sheet1");
    assert_eq!(prefix.last_sheet.as_deref(), Some("Sheet3"));
    assert!(prefix.is_3d());
    assert!(!prefix.is_local_sheet());
    assert_eq!(toks("Sheet1:Sheet3!A1:B2"), ["ref:Sheet1:Sheet3!A1:B2"]);
    let r = one_ref("'Sheet 1:Sheet 3'!$A$1");
    let prefix = r.sheet.unwrap();
    assert_eq!((prefix.sheet.as_str(), prefix.last_sheet.as_deref()), ("Sheet 1", Some("Sheet 3")));
}

#[test]
fn external_references() {
    let r = one_ref("[1]Sheet1!A1");
    let prefix = r.sheet.unwrap();
    assert_eq!(prefix.workbook.as_deref(), Some("[1]"));
    assert_eq!(prefix.sheet, "Sheet1");
    assert!(prefix.is_external());
    let r = one_ref("'[1]Sheet 1'!A1");
    assert_eq!(r.sheet.as_ref().unwrap().workbook.as_deref(), Some("[1]"));
    assert_eq!(r.sheet.as_ref().unwrap().sheet, "Sheet 1");
    let r = one_ref(r"'C:\path\[Book.xlsx]Sheet1'!$A$1");
    assert_eq!(r.sheet.as_ref().unwrap().workbook.as_deref(), Some(r"C:\path\[Book.xlsx]"));
    assert_eq!(r.sheet.as_ref().unwrap().sheet, "Sheet1");
    assert_eq!(toks("[Book1.xlsx]Sheet1!$A$1"), ["ref:[Book1.xlsx]Sheet1!$A$1"]);
    let r = one_ref("SUM([1]Sheet1:Sheet3!A1)");
    assert!(r.sheet.as_ref().unwrap().is_external() && r.sheet.as_ref().unwrap().is_3d());
    // A workbook-level name in a workbook given by path.
    let tokens = tokenize(r"'C:\dir\Book.xlsx'!Rate*2").unwrap();
    let TokenKind::Name(name) = &tokens[0].kind else { panic!("{tokens:?}") };
    let prefix = name.sheet.as_ref().unwrap();
    assert_eq!(prefix.workbook.as_deref(), Some(r"C:\dir\Book.xlsx"));
    assert_eq!(prefix.sheet, "");
    assert!(!prefix.is_3d());
    assert_eq!(rename_sheet(r"'C:\dir\Book.xlsx'!Rate", "C", "D").unwrap(), r"'C:\dir\Book.xlsx'!Rate");
}

#[test]
fn malformed_input_still_round_trips() {
    for formula in ["SUM(A1", "A1)", "}", "]", "!", "Sheet1!", "Sheet1!+", "A1 ;B1", "~`?", "{1,2",
        "€", "'x'", "A1::B2", "Sheet1!#", "=+", "a😀!A1"]
    {
        toks(formula);
    }
    // Emoji in unquoted sheet names still bind to the reference.
    assert_eq!(one_ref("a😀!A1").sheet.unwrap().sheet, "a😀");
}

#[test]
fn hand_written_formulas_round_trip_without_unknown_tokens() {
    let formulas = [
        "=SUM(A1:A10)",
        "SUM(Sheet1!A1:B2,'My Sheet'!C3)",
        "IF(AND(A1>0,B1<>\"\"),A1*B1,\"n/a\")",
        "VLOOKUP($A2,Data!$A$1:$F$500,MATCH(B$1,Data!$1:$1,0),FALSE)",
        "INDEX(Sheet2!A:A,MATCH(MAX(Sheet2!B:B),Sheet2!B:B,0))",
        "_xlfn.XLOOKUP(A2,Table1[Key],Table1[Value],\"\",0)",
        "_xlfn._xlws.SORT(_xlfn.UNIQUE(A2:A100))",
        "_xlfn.LET(_xlpm.total,SUM(B2:B9),_xlpm.total/COUNT(B2:B9))",
        "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(3)",
        "SUMPRODUCT((A1:A10=\"x\")*(B1:B10))",
        "TEXT(TODAY(),\"yyyy-mm-dd\")&\" \"&B1",
        "IFERROR(1/0,#DIV/0!)",
        "SUM(Sheet1:Sheet3!B2)",
        "[1]Sheet1!$A$1*2",
        "'[2]Price List'!$B$4",
        "OFFSET($A$1,1,0,COUNTA($A:$A)-1,1)",
        "SUM(1:1)+SUM(A:A)",
        "A1:C3 B2:D4",
        "SUM((A1:A2,C1:C2))",
        "@Table1[[#This Row],[Amount]]*1.1",
        "Table1[@Price]*Table1[@Qty]",
        "{1,2,3;4,5,6}",
        "SUM(A1#)",
        "-1E-3+.5%",
        "R1C1+RC+R2C3",
        "LOG10(A1)+ATAN2(A1,B1)+DAYS360(A1,B1)",
        "COUNTIF(A:A,\">=\"&B1)",
        "\"say \"\"hi\"\"\"",
        "'It''s'!A1+'Sheet 1'!B2",
        "Sheet2!#REF!+#REF!",
        "HYPERLINK(\"#Sheet2!A1\",\"go\")",
        "GETPIVOTDATA(\"Sales\",$A$3)",
        "NOT(ISERROR(SEARCH(\"say \"\"hi\"\"\",A1)))",
        "SUBTOTAL(109,[Column'#])",
        "Sheet1!$A:$A,Sheet1!$1:$1",
        "_xlfn.STDEV.S(A1:A5)",
        "x14:sparkline",
    ];
    for formula in formulas {
        let tokens = tokenize(formula).unwrap();
        assert_eq!(render(&tokens), formula);
        if formula != "x14:sparkline" {
            assert!(
                !tokens.iter().any(|t| t.kind == TokenKind::Unknown),
                "{formula:?}: {tokens:?}"
            );
        }
    }
}

#[test]
fn random_input_never_panics_and_always_round_trips() {
    // Deterministic pseudo-random strings built from formula-ish pieces.
    const PIECES: &[&str] = &[
        "A1", "$B$2", ":", "!", "'", "''", "\"", "[", "]", "#", "#REF!", "{", "}", "(", ")", ",",
        ";", " ", "\n", "Sheet1", "My Sheet", "1", "1E+3", ".", "$", "@", "%", "+", "-", "=",
        "SUM", "_xlfn.", "TRUE", "XFD", "1048576", "é", "😀", "[1]", "Table1", "#This Row", "C",
        "R1C1", "<", ">", "&", "*", "/", "^", "\\",
    ];
    let mut seed: u64 = 0x2545_F491_4F6C_DD1D;
    let mut next = move || {
        seed = seed.wrapping_mul(6_364_136_223_846_793_005).wrapping_add(1_442_695_040_888_963_407);
        (seed >> 33) as usize
    };
    let edits = [ins_rows(0, 2), del_rows(0, 1), ins_cols(1, 1), del_cols(0, 3)];
    for _ in 0..20_000 {
        let len = next() % 12;
        let formula: String = (0..len).map(|_| PIECES[next() % PIECES.len()]).collect();
        match tokenize(&formula) {
            Ok(tokens) => {
                assert_eq!(render(&tokens), formula);
                for edit in &edits {
                    let shifted = shift(&formula, edit);
                    assert!(tokenize(&shifted).is_ok(), "{formula:?} -> {shifted:?}");
                }
                let moved = translate_formula(&formula, 1, -1).unwrap();
                assert!(tokenize(&moved).is_ok(), "{formula:?} -> {moved:?}");
                let renamed = rename_sheet(&formula, "Sheet1", "New Name").unwrap();
                assert!(tokenize(&renamed).is_ok(), "{formula:?} -> {renamed:?}");
            }
            Err(e) => {
                assert!(e.position <= formula.len());
                assert!(formula.is_char_boundary(e.position));
                assert!(shift_formula(&formula, "Sheet1", &edits[0]).is_err());
            }
        }
    }
}

// ---------------------------------------------------------------------------
// references(), Area, quoting.
// ---------------------------------------------------------------------------

#[test]
fn references_lists_only_cell_references() {
    let refs = references("SUM(Sheet2!A1:B2, 'My Sheet'!C3, MyName, Table1[Col], \"D4\", {5})+E5")
        .unwrap();
    let texts: Vec<String> = refs.iter().map(|r| r.to_string()).collect();
    assert_eq!(texts, ["Sheet2!A1:B2", "'My Sheet'!C3", "E5"]);
    assert_eq!(refs[0].sheet.as_ref().unwrap().sheet, "Sheet2");
    assert!(refs[2].sheet.is_none());
    assert!(references("\"unterminated").is_err());
}

#[test]
fn area_parse_and_display() {
    for text in ["A1", "$A$1", "A$1:$B2", "A:C", "$A:$A", "1:3", "$2:$2", "XFD1048576", "#REF!"] {
        assert_eq!(Area::parse(text).unwrap().to_string(), text, "{text}");
    }
    assert_eq!(Area::parse("a1:b2").unwrap().to_string(), "A1:B2");
    for bad in ["", "A", "1", "A0", "XFE1", "A1:", ":A1", "A1:B", "1:A", "Sheet1!A1", "A1:B2:C3"] {
        assert_eq!(Area::parse(bad), None, "{bad}");
    }
    assert_eq!(Area::parse("B3:A1").unwrap().bounds(), Some((0, 0, 2, 1)));
    assert_eq!(Area::parse("B:C").unwrap().bounds(), Some((0, 1, 1_048_575, 2)));
    assert_eq!(Area::parse("2:3").unwrap().bounds(), Some((1, 0, 2, 16_383)));
    assert_eq!(Area::RefError.bounds(), None);
    let edit = ins_rows(0, 1);
    assert_eq!(Area::parse("A1:B2").unwrap().shift(&edit), Area::parse("A2:B3"));
}

#[test]
fn sheet_name_quoting_rules() {
    for name in ["Sheet1", "Data", "Sh.eet.1", "Sheet_1", "Sheeté", "日本", "ABCD1", "Rabcd1"] {
        assert!(!sheet_name_needs_quotes(name), "{name}");
        assert_eq!(quote_sheet_name(name), name);
    }
    for name in ["My Sheet", "1Sheet", ".x", "A1", "xfd1048576", "R1C1", "R2", "C3", "RC1", "R",
        "c", "rc", "Ra1", "TRUE", "false", "Sheet-1", "It's", "a😀", ""]
    {
        assert!(sheet_name_needs_quotes(name), "{name}");
    }
    assert_eq!(quote_sheet_name("It's"), "'It''s'");
    assert_eq!(quote_sheet_name("My Sheet"), "'My Sheet'");
    assert_eq!(SheetPrefix::local("My Sheet").text, "'My Sheet'!");
    assert_eq!(SheetPrefix::local("Data").text, "Data!");
}

// ---------------------------------------------------------------------------
// Structural edits.
// ---------------------------------------------------------------------------

#[test]
fn insert_rows() {
    // Two rows inserted before row 5.
    check_shift(
        &ins_rows(4, 2),
        &[
            ("A5", "A7"),
            ("A4", "A4"),
            ("$A$5", "$A$7"),
            ("A$5", "A$7"),
            ("$A5+B6", "$A7+B8"),
            ("A1:A10", "A1:A12"),
            ("A5:A10", "A7:A12"),
            ("A1:A4", "A1:A4"),
            ("A1:A5", "A1:A7"),
            ("$B$2:$C$5", "$B$2:$C$7"),
            ("A:A", "A:A"),
            ("$B:$D", "$B:$D"),
            ("5:10", "7:12"),
            ("1:4", "1:4"),
            ("$4:$5", "$4:$7"),
            ("=SUM( A5 , \"x\" ) + B5", "=SUM( A7 , \"x\" ) + B7"),
            ("SUM(A5#)", "SUM(A7#)"),
            ("a5", "A7"),
            ("B2:A10", "B2:A12"),
            ("A10:A5", "A12:A7"),
        ],
    );
}

#[test]
fn insert_columns() {
    // One column inserted before column B.
    check_shift(
        &ins_cols(1, 1),
        &[
            ("A1", "A1"),
            ("B1", "C1"),
            ("$B$1", "$C$1"),
            ("A1:C1", "A1:D1"),
            ("B1:C1", "C1:D1"),
            ("B:D", "C:E"),
            ("A:A", "A:A"),
            ("A:B", "A:C"),
            ("1:1", "1:1"),
            ("$1:$3", "$1:$3"),
            ("SUM(A1:A9)+B2*$C$3", "SUM(A1:A9)+C2*$D$3"),
        ],
    );
}

#[test]
fn delete_rows() {
    // Rows 5 and 6 deleted.
    check_shift(
        &del_rows(4, 2),
        &[
            ("A5", "#REF!"),
            ("A6", "#REF!"),
            ("A7", "A5"),
            ("A4", "A4"),
            ("$A$7", "$A$5"),
            ("Sheet1!A5", "Sheet1!#REF!"),
            ("A1:A10", "A1:A8"),
            ("A5:A6", "#REF!"),
            ("A5:A10", "A5:A8"),
            ("A3:A5", "A3:A4"),
            ("A3:A6", "A3:A4"),
            ("A6:A8", "A5:A6"),
            ("A7:A9", "A5:A7"),
            ("A4:A7", "A4:A5"),
            ("SUM(A5:A6)", "SUM(#REF!)"),
            ("SUM(A5#)", "SUM(#REF!)"),
            ("A7#", "A5#"),
            ("5:6", "#REF!"),
            ("4:7", "4:5"),
            ("$8:$9", "$6:$7"),
            ("A:B", "A:B"),
            ("B10:A2", "B8:A2"),
            ("A5+A7", "#REF!+A5"),
            ("IF(A5>0,A5,A8)", "IF(#REF!>0,#REF!,A6)"),
        ],
    );
}

#[test]
fn delete_columns() {
    // Columns B and C deleted.
    check_shift(
        &del_cols(1, 2),
        &[
            ("A1", "A1"),
            ("B1", "#REF!"),
            ("C1", "#REF!"),
            ("D1", "B1"),
            ("$D$1", "$B$1"),
            ("A1:D1", "A1:B1"),
            ("B1:C5", "#REF!"),
            ("B1:D5", "B1:B5"),
            ("A1:B5", "A1:A5"),
            ("B:C", "#REF!"),
            ("A:D", "A:B"),
            ("C:E", "B:C"),
            ("D:D", "B:B"),
            ("1:3", "1:3"),
            ("A1:A1", "A1:A1"),
            ("A1:B1", "A1:A1"),
        ],
    );
}

#[test]
fn edits_only_touch_the_edited_sheet() {
    let edit = ins_rows(0, 1);
    assert_eq!(shift("Sheet2!A1", &edit), "Sheet2!A1");
    assert_eq!(shift("Sheet1!A1", &edit), "Sheet1!A2");
    assert_eq!(shift("sheet1!A1", &edit), "sheet1!A2");
    assert_eq!(shift("'Sheet1'!A1", &edit), "'Sheet1'!A2");
    assert_eq!(shift("A1+Sheet2!A1+Sheet1!B1", &edit), "A2+Sheet2!A1+Sheet1!B2");
    // A formula on another sheet: unqualified references are not moved.
    assert_eq!(shift_formula("A1+Sheet1!A1", "Sheet2", &edit).unwrap(), "A1+Sheet1!A2");
    // Quoted sheet names.
    let edit = StructuralEdit::DeleteRows { sheet: "My Sheet".into(), at: 4, count: 1 };
    assert_eq!(shift_formula("'My Sheet'!A5", "Sheet1", &edit).unwrap(), "'My Sheet'!#REF!");
    assert_eq!(shift_formula("'my sheet'!A6:B9", "Sheet1", &edit).unwrap(), "'my sheet'!A5:B8");
    assert_eq!(shift_formula("A6", "My Sheet", &edit).unwrap(), "A5");
    let edit = StructuralEdit::InsertCols { sheet: "It's".into(), at: 0, count: 2 };
    assert_eq!(shift_formula("'It''s'!A1", "x", &edit).unwrap(), "'It''s'!C1");
    // Sheet-scoped names and structured references never move.
    let edit = ins_rows(0, 1);
    assert_eq!(shift("Sheet1!MyName+Table1[Col]+MyName", &edit), "Sheet1!MyName+Table1[Col]+MyName");
}

#[test]
fn external_3d_and_deleted_sheet_references_are_not_shifted() {
    let edit = ins_rows(0, 1);
    for formula in ["[1]Sheet1!A1", "'[1]Sheet1'!A1", "Sheet1:Sheet3!A1", "SUM(Sheet1:Sheet1!A1)",
        "#REF!A1", "[1]!Name", "'C:\\x\\[Book.xlsx]Sheet1'!A1"]
    {
        assert_eq!(shift(formula, &edit), formula);
    }
}

#[test]
fn strings_array_constants_and_names_are_untouched() {
    let edit = ins_rows(0, 3);
    for formula in ["\"A1\"", "INDIRECT(\"Sheet1!A1\")", "{1,2;3,4}", "MyName", "R1C1",
        "LOG10(100)", "1E+3", "TRUE", "#N/A", "Table1[[#This Row],[A1]]", "SUBTOTAL(109,[A1])"]
    {
        assert_eq!(shift(formula, &edit), formula);
    }
    assert_eq!(shift("CONCAT(\"A1\",A1,{\"A1\"})", &edit), "CONCAT(\"A1\",A4,{\"A1\"})");
}

#[test]
fn sheet_limits() {
    check_shift(
        &ins_rows(0, 1),
        &[
            ("A1048576", "#REF!"),
            ("A1048575", "A1048576"),
            ("Sheet1!$A$1048576", "Sheet1!#REF!"),
            ("A5:A1048576", "A6:A1048576"),
            ("A1:A1048576", "A2:A1048576"),
            ("A1048576:A1048576", "#REF!"),
            ("1048576:1048576", "#REF!"),
            ("2:1048576", "3:1048576"),
            ("A1048576#", "#REF!"),
        ],
    );
    check_shift(&ins_rows(1_048_575, 5), &[("A1048576", "#REF!"), ("A1:A1048576", "A1:A1048576")]);
    check_shift(
        &ins_cols(0, 2),
        &[("XFD1", "#REF!"), ("XFC1", "#REF!"), ("XFB1", "XFD1"), ("B1:XFD1", "D1:XFD1"),
            ("A:XFD", "C:XFD"), ("XFD:XFD", "#REF!")],
    );
    check_shift(&ins_cols(5, 2), &[("A:XFD", "A:XFD"), ("A1:XFD1", "A1:XFD1")]);
    // Deleting rows past the end of a range or at the sheet end.
    check_shift(&del_rows(1_048_570, 100), &[("A1048576", "#REF!"), ("A1:A1048576", "A1:A1048570")]);
}

#[test]
fn zero_count_edits_change_nothing() {
    for edit in [ins_rows(0, 0), del_rows(0, 0), ins_cols(0, 0), del_cols(0, 0)] {
        assert_eq!(shift("SUM(A1:B2)+C3", &edit), "SUM(A1:B2)+C3");
    }
}

#[test]
fn unchanged_formulas_are_returned_byte_for_byte() {
    let edit = ins_rows(100, 1);
    for formula in ["sum( a1 :b2)", "SUM( A1 )", "='Sheet1'!a1", "a1+  b2"] {
        assert_eq!(shift(formula, &edit), formula);
    }
    // Only the changed token is re-rendered.
    assert_eq!(shift("sum(a1,a200)", &edit), "sum(a1,A201)");
}

#[test]
fn shift_mixed_real_world_formulas() {
    let edit = ins_rows(2, 1);
    assert_eq!(
        shift("IF(Sheet2!B5>0,SUM($B$2:$B$9),\"B5\")", &edit),
        "IF(Sheet2!B5>0,SUM($B$2:$B$10),\"B5\")"
    );
    assert_eq!(
        shift("VLOOKUP($A3,Data!$A:$F,2,FALSE)+SUM(B3:B9 B5:D5)", &edit),
        "VLOOKUP($A4,Data!$A:$F,2,FALSE)+SUM(B4:B10 B6:D6)"
    );
    assert_eq!(
        shift("(Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5)", &edit),
        "(Sheet1!$A$1:$A$2,Sheet1!$A$5:$A$6)"
    );
    assert_eq!(shift("Sheet1!$A:$A,Sheet1!$3:$3", &edit), "Sheet1!$A:$A,Sheet1!$4:$4");
    assert_eq!(shift("A1:INDEX(A:A,5)", &edit), "A1:INDEX(A:A,5)");
    assert_eq!(shift("A1:A3:A2", &edit), "A1:A4:A2");
    let edit = del_rows(2, 1);
    assert_eq!(
        shift("SUMIFS($C:$C,$A:$A,A3,$B:$B,\">0\")+A4", &edit),
        "SUMIFS($C:$C,$A:$A,#REF!,$B:$B,\">0\")+A3"
    );
}

#[test]
fn insert_then_delete_restores_the_formula() {
    let formulas = [
        "SUM(A1:A10)+B5*$C$7",
        "Sheet1!$B$2:$D$20",
        "A:C",
        "3:9",
        "INDEX($A$1:$Z$100,MATCH(F5,$A$1:$A$100,0),3)",
        "'Sheet1'!E5+Sheet2!E5",
    ];
    for formula in formulas {
        for at in [0u32, 1, 4, 6, 50] {
            for count in [1u32, 3] {
                let inserted = shift(formula, &ins_rows(at, count));
                assert_eq!(shift(&inserted, &del_rows(at, count)), formula, "{formula} rows {at}+{count}");
                let inserted = shift(formula, &ins_cols(at as u16, count as u16));
                assert_eq!(
                    shift(&inserted, &del_cols(at as u16, count as u16)),
                    formula,
                    "{formula} cols {at}+{count}"
                );
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Range lists.
// ---------------------------------------------------------------------------

#[test]
fn range_lists() {
    assert_eq!(shift_range_list("A1:B2 D4", true, &ins_rows(0, 1)).as_deref(), Some("A2:B3 D5"));
    assert_eq!(shift_range_list("A1:B2 D4", false, &ins_rows(0, 1)).as_deref(), Some("A1:B2 D4"));
    assert_eq!(shift_range_list("A1:B2 D4", true, &ins_rows(10, 1)).as_deref(), Some("A1:B2 D4"));
    assert_eq!(shift_range_list("A1:B2 D4", true, &ins_rows(1, 2)).as_deref(), Some("A1:B4 D6"));
    assert_eq!(shift_range_list("A5 B5:C5", true, &del_rows(4, 1)), None);
    assert_eq!(shift_range_list("A5 B5:C9", true, &del_rows(4, 1)).as_deref(), Some("B5:C8"));
    assert_eq!(shift_range_list("A1:A2", true, &del_rows(1, 1)).as_deref(), Some("A1"));
    assert_eq!(shift_range_list("A1:B1", true, &del_cols(1, 1)).as_deref(), Some("A1"));
    assert_eq!(shift_range_list("B1:C3 E1", true, &del_cols(1, 2)).as_deref(), Some("C1"));
    assert_eq!(shift_range_list("A:A 1:1", true, &ins_cols(0, 1)).as_deref(), Some("B:B 1:1"));
    assert_eq!(shift_range_list("A:A 2:2", true, &del_rows(1, 1)).as_deref(), Some("A:A"));
    assert_eq!(shift_range_list("A1:XFD1048576", true, &ins_rows(0, 1)).as_deref(), Some("A2:XFD1048576"));
    assert_eq!(shift_range_list("bogus A3", true, &ins_rows(0, 1)).as_deref(), Some("bogus A4"));
    assert_eq!(shift_range_list("", true, &ins_rows(0, 1)).as_deref(), Some(""));
    assert_eq!(shift_range_list("A1  B2", true, &ins_rows(9, 1)).as_deref(), Some("A1  B2"));
}

// ---------------------------------------------------------------------------
// Copy/paste translation.
// ---------------------------------------------------------------------------

#[test]
fn translate() {
    let cases: &[(&str, i64, i64, &str)] = &[
        ("A1", 1, 0, "A2"),
        ("A1", 0, 1, "B1"),
        ("$A$1", 5, 5, "$A$1"),
        ("A$1", 1, 1, "B$1"),
        ("$A1", 1, 1, "$A2"),
        ("SUM(A1:B2)", 2, 3, "SUM(D3:E4)"),
        ("SUM($A1:B$2)", 2, 3, "SUM($A3:E$2)"),
        ("B2", -1, -1, "A1"),
        ("A1", -1, 0, "#REF!"),
        ("A1+B2", -1, 0, "#REF!+B1"),
        ("Sheet2!A1", -1, 0, "Sheet2!#REF!"),
        ("SUM(A1:A5)", -1, 0, "SUM(#REF!)"),
        ("XFD1", 0, 1, "#REF!"),
        ("A1048576", 1, 0, "#REF!"),
        ("A:A", 5, 1, "B:B"),
        ("$A:B", 5, 1, "$A:C"),
        ("1:1", 5, 1, "6:6"),
        ("$1:2", 5, 1, "$1:7"),
        ("A:A", 0, -1, "#REF!"),
        ("'My Sheet'!B2*2", 3, 0, "'My Sheet'!B5*2"),
        ("[1]Sheet1!A1", 1, 0, "[1]Sheet1!A2"),
        ("Sheet1:Sheet3!A1", 1, 0, "Sheet1:Sheet3!A2"),
        ("\"A1\"&A1", 1, 0, "\"A1\"&A2"),
        ("{1,2}+Table1[Col]+MyName", 4, 4, "{1,2}+Table1[Col]+MyName"),
        ("A1#", 1, 1, "B2#"),
        ("A1#", -1, 0, "#REF!"),
        ("Sheet2!#REF!+A1", 1, 0, "Sheet2!#REF!+A2"),
        ("sum(a1)", 0, 0, "sum(a1)"),
        ("SUM(A1:A3)", 0, 0, "SUM(A1:A3)"),
    ];
    for (formula, rows, cols, expected) in cases {
        assert_eq!(translate_formula(formula, *rows, *cols).unwrap(), *expected, "{formula} {rows} {cols}");
    }
    assert!(translate_formula("\"x", 1, 1).is_err());
    assert!(translate_formula("\"x", 0, 0).is_err());
}

// ---------------------------------------------------------------------------
// Sheet renames.
// ---------------------------------------------------------------------------

#[test]
fn rename() {
    let cases: &[(&str, &str, &str, &str)] = &[
        ("Sheet1!A1", "Sheet1", "Data", "Data!A1"),
        ("Sheet1!A1", "Sheet1", "My Data", "'My Data'!A1"),
        ("'My Data'!A1:B2", "My Data", "Data", "Data!A1:B2"),
        ("Sheet1!A1", "Sheet1", "It's", "'It''s'!A1"),
        ("'It''s'!A1", "It's", "Plain", "Plain!A1"),
        ("SHEET1!A1", "Sheet1", "Data", "Data!A1"),
        ("'Sheet1'!A1", "Sheet1", "Data", "Data!A1"),
        ("Sheet1!A1", "Sheet1", "A1", "'A1'!A1"),
        ("Sheet1!A1", "Sheet1", "TRUE", "'TRUE'!A1"),
        ("Sheet1!A1", "Sheet1", "1abc", "'1abc'!A1"),
        ("Sheet1!A1", "Sheet1", "R1C1", "'R1C1'!A1"),
        ("Sheet1!A1", "Sheet1", "Sheeté", "Sheeté!A1"),
        ("Sheet1:Sheet3!A1", "Sheet3", "Last Sheet", "'Sheet1:Last Sheet'!A1"),
        ("Sheet1:Sheet3!A1", "Sheet1", "First", "First:Sheet3!A1"),
        ("'Sheet 1:Sheet 3'!A1", "Sheet 1", "S1", "'S1:Sheet 3'!A1"),
        ("'Sheet 1:Sheet 3'!A1", "Sheet 3", "S3", "'Sheet 1:S3'!A1"),
        ("Sheet1!Name+Sheet1!$A$1", "Sheet1", "Data", "Data!Name+Data!$A$1"),
        ("Sheet1!#REF!", "Sheet1", "Data", "Data!#REF!"),
        ("Sheet1!A1#", "Sheet1", "Data", "Data!A1#"),
        ("Sheet1!$1:$2", "Sheet1", "My Data", "'My Data'!$1:$2"),
        ("[1]Sheet1!A1", "Sheet1", "Data", "[1]Sheet1!A1"),
        ("'[1]Sheet1'!A1", "Sheet1", "Data", "'[1]Sheet1'!A1"),
        ("\"Sheet1!A1\"&Sheet1!a1", "Sheet1", "Data", "\"Sheet1!A1\"&Data!a1"),
        ("Sheet10!A1+Sheet1!A1", "Sheet1", "X", "Sheet10!A1+X!A1"),
        ("A1+Sheet2!A1", "Sheet1", "X", "A1+Sheet2!A1"),
        ("#REF!A1", "#REF!", "X", "#REF!A1"),
        ("SUM( Sheet1!A1 , 'Sheet1'!B2 )", "Sheet1", "New", "SUM( New!A1 , New!B2 )"),
    ];
    for (formula, old, new, expected) in cases {
        assert_eq!(rename_sheet(formula, old, new).unwrap(), *expected, "{formula} {old}->{new}");
    }
    assert!(rename_sheet("'abc", "a", "b").is_err());
}

// ---------------------------------------------------------------------------
// Real-world corpus: every formula in every reference file must round-trip.
// ---------------------------------------------------------------------------

mod corpus {
    use super::*;
    use quick_xml::events::Event;
    use std::io::Read;
    use std::path::{Path, PathBuf};

    /// Elements holding formulas: cell `<f>`, chart `<c:f>`, `<xm:f>`,
    /// conditional format `<formula>`, data validation `<formula1/2>`, and
    /// `<definedName>`.
    const FORMULA_ELEMENTS: &[&str] = &["f", "formula", "formula1", "formula2", "definedName"];

    /// `tests/corpus`, plus an optional extra directory of xlsx files given by
    /// `FORMULA_CORPUS_DIR` (e.g. rust_xlsxwriter's `tests/input`).
    fn corpus_dirs() -> Vec<PathBuf> {
        let mut dirs = vec![PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/corpus")];
        if let Ok(extra) = std::env::var("FORMULA_CORPUS_DIR") {
            dirs.push(PathBuf::from(extra));
        }
        dirs.into_iter().filter(|d| d.is_dir()).collect()
    }

    #[derive(Default)]
    struct Found {
        files: usize,
        /// (file, element, text)
        formulas: Vec<(String, String, String)>,
        /// (file, range list)
        range_lists: Vec<(String, String)>,
    }

    fn scan_xml(file: &str, xml: &str, found: &mut Found) {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut current: Option<(String, String)> = None;
        loop {
            let event = match reader.read_event() {
                Ok(event) => event,
                Err(e) => panic!("{file}: bad XML: {e}"),
            };
            let is_start = matches!(event, Event::Start(_));
            match event {
                Event::Start(e) | Event::Empty(e) => {
                    let local = e.local_name().as_ref().to_string();
                    for attr in e.attributes().flatten() {
                        let key = attr.key.local_name();
                        let key = key.as_ref();
                        if key == "sqref" || (key == "ref" && local == "mergeCell") {
                            // Range lists contain no XML escapes.
                            let value = attr.value.to_string();
                            found.range_lists.push((file.to_string(), value));
                        }
                    }
                    if is_start && (FORMULA_ELEMENTS.contains(&local.as_str()) || local == "sqref") {
                        current = Some((local, String::new()));
                    }
                }
                Event::Text(t) => {
                    if let Some((_, raw)) = &mut current {
                        raw.push_str(&t);
                    }
                }
                Event::GeneralRef(r) => {
                    if let Some((_, raw)) = &mut current {
                        raw.push('&');
                        raw.push_str(&r);
                        raw.push(';');
                    }
                }
                Event::CData(c) => {
                    if let Some((_, raw)) = &mut current {
                        raw.push_str(&quick_xml::escape::escape(&*c));
                    }
                }
                Event::End(e) => {
                    let local = e.local_name().as_ref().to_string();
                    if current.as_ref().is_some_and(|(name, _)| *name == local) {
                        let (name, raw) = current.take().unwrap();
                        let text = quick_xml::escape::unescape(&raw).unwrap().into_owned();
                        if name == "sqref" {
                            found.range_lists.push((file.to_string(), text));
                        } else {
                            found.formulas.push((file.to_string(), name, text));
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
    }

    fn scan_file(path: &Path, found: &mut Found) {
        let name = path.file_name().unwrap().to_string_lossy().into_owned();
        let Ok(file) = std::fs::File::open(path) else { return };
        let Ok(mut zip) = zip::ZipArchive::new(file) else {
            return; // not a zip (e.g. a deliberately broken file)
        };
        found.files += 1;
        for i in 0..zip.len() {
            let mut entry = zip.by_index(i).unwrap();
            let entry_name = entry.name().map(|n| n.into_owned()).unwrap_or_default();
            if !entry_name.ends_with(".xml") {
                continue;
            }
            let mut xml = String::new();
            if entry.read_to_string(&mut xml).is_err() {
                continue;
            }
            scan_xml(&format!("{name}:{entry_name}"), &xml, found);
        }
    }

    fn collect() -> Found {
        let mut found = Found::default();
        for dir in corpus_dirs() {
            let mut paths: Vec<PathBuf> = std::fs::read_dir(&dir)
                .unwrap()
                .flatten()
                .map(|e| e.path())
                .filter(|p| {
                    p.extension().is_some_and(|ext| ext == "xlsx" || ext == "xlsm")
                })
                .collect();
            paths.sort();
            for path in paths {
                scan_file(&path, &mut found);
            }
        }
        found
    }

    #[test]
    fn every_corpus_formula_round_trips() {
        let found = collect();
        assert!(found.files > 0, "no corpus files found");
        let mut unique = std::collections::BTreeSet::new();
        let no_edit = StructuralEdit::InsertRows { sheet: "\u{0}none".into(), at: 0, count: 1 };
        for (file, element, formula) in &found.formulas {
            let tokens = tokenize(formula).unwrap_or_else(|e| panic!("{file} <{element}> {formula:?}: {e}"));
            assert_eq!(render(&tokens), *formula, "{file} <{element}>");
            assert!(
                !tokens.iter().any(|t| t.kind == TokenKind::Unknown),
                "{file} <{element}> {formula:?}: {tokens:?}"
            );
            // Rewrites that affect nothing return the formula unchanged.
            // (The edit is on a sheet that no reference names.)
            assert_eq!(shift_formula(formula, "Sheet1", &no_edit).unwrap(), *formula);
            assert_eq!(translate_formula(formula, 0, 0).unwrap(), *formula);
            assert_eq!(rename_sheet(formula, "\u{0}none", "x").unwrap(), *formula);
            // Inserting then deleting the same rows/columns restores it.
            for (insert, delete) in [
                (ins_rows(1, 3), del_rows(1, 3)),
                (ins_cols(1, 2), del_cols(1, 2)),
            ] {
                let inserted = shift(formula, &insert);
                assert_eq!(shift(&inserted, &delete), *formula, "{file}: {formula}");
            }
            // Translating there and back restores it (unless it fell off).
            let moved = translate_formula(formula, 2, 3).unwrap();
            if moved.matches("#REF!").count() == formula.matches("#REF!").count() {
                assert_eq!(translate_formula(&moved, -2, -3).unwrap(), *formula);
            }
            unique.insert(formula.clone());
        }
        for (file, list) in &found.range_lists {
            for item in list.split_whitespace() {
                assert!(Area::parse(item).is_some(), "{file}: range list item {item:?}");
            }
            let edit = ins_rows(1_000_000, 1);
            assert_eq!(shift_range_list(list, true, &edit).as_deref(), Some(list.as_str()));
        }
        eprintln!(
            "formula corpus: {} files, {} formulas ({} unique), {} range lists",
            found.files,
            found.formulas.len(),
            unique.len(),
            found.range_lists.len()
        );
    }
}
