"""Data validation: DataValidation, DataValidationRule, DataValidationErrorStyle."""

import datetime as dt

import pytest

import pyrust_xlsxwriter as px

R = px.DataValidationRule


def validations(reopen, path, sheet="Sheet1"):
    sh = reopen(path)[sheet]
    return {str(dv.sqref): dv for dv in sh.data_validations.dataValidation}


def save(out, *pairs):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    for cell, dv in pairs:
        ws.add_data_validation(cell, dv)
    path = out()
    wb.save(path)
    return path


def test_number_rules(out, reopen):
    path = save(
        out,
        ("A1", px.DataValidation(whole_number=R.between(1, 10))),
        ("A2", px.DataValidation().allow_whole_number(R.not_between(-5, 5))),
        ("A3", px.DataValidation(decimal_number=R.greater_than(0.5))),
        ("A4", px.DataValidation().allow_decimal_number(R.less_than_or_equal_to(99))),
        ("A5", px.DataValidation(text_length=R.less_than(20))),
        ("A6", px.DataValidation(whole_number=R.equal_to(3))),
        ("A7", px.DataValidation(whole_number=R.not_equal_to(3))),
        ("A8", px.DataValidation(whole_number=R.greater_than_or_equal_to(3))),
    )
    dvs = validations(reopen, path)
    assert (dvs["A1"].type, dvs["A1"].operator, dvs["A1"].formula1, dvs["A1"].formula2) == ("whole", None, "1", "10")
    assert (dvs["A2"].operator, dvs["A2"].formula1, dvs["A2"].formula2) == ("notBetween", "-5", "5")
    assert (dvs["A3"].type, dvs["A3"].operator, dvs["A3"].formula1) == ("decimal", "greaterThan", "0.5")
    assert (dvs["A4"].operator, dvs["A4"].formula1) == ("lessThanOrEqual", "99")
    assert (dvs["A5"].type, dvs["A5"].operator, dvs["A5"].formula1) == ("textLength", "lessThan", "20")
    assert dvs["A6"].operator == "equal"
    assert dvs["A7"].operator == "notEqual"
    assert dvs["A8"].operator == "greaterThanOrEqual"


def test_date_time_rules(out, reopen):
    path = save(
        out,
        ("B1", px.DataValidation(date=R.between(dt.date(2024, 1, 1), dt.datetime(2024, 12, 31, 12, 0)))),
        ("B2", px.DataValidation().allow_time(R.greater_than(dt.time(12, 0)))),
    )
    dvs = validations(reopen, path)
    assert (dvs["B1"].type, dvs["B1"].formula1, dvs["B1"].formula2) == ("date", "45292", "45657.5")
    assert (dvs["B2"].type, dvs["B2"].operator, dvs["B2"].formula1) == ("time", "greaterThan", "0.5")


def test_formula_rules(out, reopen):
    path = save(
        out,
        ("C1", px.DataValidation().allow_whole_number_formula(R.between(1, "=$E$1"))),
        ("C2", px.DataValidation(decimal_number=R.less_than(px.Formula("=E2*2")))),
        ("C3", px.DataValidation().allow_date_formula(R.greater_than("=TODAY()"))),
        ("C4", px.DataValidation().allow_time_formula(R.less_than("=E4"))),
        ("C5", px.DataValidation().allow_text_length_formula(R.equal_to("=E5"))),
        ("C6", px.DataValidation().allow_date(R.between(dt.date(2024, 1, 1), "=E6"))),
        ("C7", px.DataValidation().allow_decimal_number_formula(R.between(1.5, "=E7"))),
    )
    dvs = validations(reopen, path)
    assert (dvs["C1"].type, dvs["C1"].formula1, dvs["C1"].formula2) == ("whole", "1", "$E$1")
    assert dvs["C2"].formula1 == "E2*2"
    assert (dvs["C3"].type, dvs["C3"].formula1) == ("date", "TODAY()")
    assert dvs["C4"].type == "time" and dvs["C5"].type == "textLength"
    assert (dvs["C6"].type, dvs["C6"].formula1, dvs["C6"].formula2) == ("date", "45292", "E6")
    assert (dvs["C7"].type, dvs["C7"].formula1, dvs["C7"].formula2) == ("decimal", "1.5", "E7")


def test_list_custom_and_any(out, reopen):
    path = save(
        out,
        ("D1", px.DataValidation(list_strings=["Open", 'Say "hi"', "Closed"])),
        ("D2", px.DataValidation().allow_list_formula("=$F$1:$F$3")),
        ("D3", px.DataValidation(custom="=ISTEXT(D3)")),
        ("D4", px.DataValidation(input_title="Note", input_message="Anything goes").allow_any_value()),
    )
    dvs = validations(reopen, path)
    assert dvs["D1"].type == "list" and dvs["D1"].formula1 == '"Open,Say ""hi"",Closed"'
    assert dvs["D2"].type == "list" and dvs["D2"].formula1 == "$F$1:$F$3"
    assert dvs["D3"].type == "custom" and dvs["D3"].formula1 == "ISTEXT(D3)"
    assert dvs["D4"].type is None and dvs["D4"].prompt == "Anything goes"


def test_messages_and_flags(out, reopen):
    dv = px.DataValidation(
        whole_number=R.between(1, 10),
        ignore_blank=False,
        show_dropdown=False,
        show_input_message=False,
        show_error_message=False,
        input_title="Enter",
        input_message="A number from 1 to 10",
        error_title="Oops",
        error_message="That is not 1 to 10",
        error_style=px.DataValidationErrorStyle.Warning,
        multi_range="A1:A3 C1:C3",
    )
    path = save(out, ("A1:A3", dv))
    (only,) = validations(reopen, path).values()
    assert str(only.sqref) == "A1:A3 C1:C3"
    assert not only.allow_blank
    assert only.showDropDown  # Excel's attribute is inverted: 1 hides the arrow
    assert not only.showInputMessage and not only.showErrorMessage
    assert (only.promptTitle, only.prompt) == ("Enter", "A number from 1 to 10")
    assert (only.errorTitle, only.error) == ("Oops", "That is not 1 to 10")
    assert only.errorStyle == "warning"


def test_error_style_information(out, reopen):
    dv = px.DataValidation(list_strings=["a"], error_style=px.DataValidationErrorStyle.Information)
    path = save(out, ("A1", dv))
    assert next(iter(validations(reopen, path).values())).errorStyle == "information"


def test_ints_and_a1(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.add_data_validation(1, 1, 3, 2, px.DataValidation(whole_number=R.greater_than(0)))
    ws.add_data_validation(0, 0, 0, 0, data_validation=px.DataValidation(list_strings=["x"]))
    path = out()
    wb.save(path)
    assert set(validations(reopen, path)) == {"B2:C4", "A1"}


def test_properties_round_trip():
    dv = px.DataValidation()
    assert dv.validation_type == px.DataValidationType.Any
    assert dv.ignore_blank and dv.show_dropdown and dv.show_input_message and dv.show_error_message
    assert dv.error_style == px.DataValidationErrorStyle.Stop
    assert dv.input_title is None and dv.multi_range is None
    dv.whole_number = R.between(1, 5)
    assert dv.validation_type == px.DataValidationType.WholeNumber
    assert dv.whole_number == R.between(1, 5) and dv.decimal_number is None
    dv.decimal_number = R.less_than(2.5)  # replaces the previous kind
    assert dv.whole_number is None and dv.decimal_number is not None and dv.decimal_number.values == (2.5,)
    dv.list_strings = ["a", "b"]
    assert dv.validation_type == px.DataValidationType.List and dv.list_strings == ["a", "b"]
    dv.list_formula = px.Formula("=$A$1:$A$3")
    assert dv.list_formula == "=$A$1:$A$3" and dv.list_strings is None
    dv.custom = "=A1>0"
    assert dv.validation_type == px.DataValidationType.Custom and dv.custom == "=A1>0"
    dv.custom = None
    assert dv.validation_type == px.DataValidationType.Any
    dv.date = R.equal_to(dt.date(2025, 1, 2))
    assert dv.date is not None and dv.date.values == (dt.date(2025, 1, 2),)
    dv.time = R.equal_to(dt.time(1, 2))
    assert dv.time is not None and dv.time.values == (dt.time(1, 2),) and dv.date is None
    dv.text_length = R.less_than(5)
    assert dv.validation_type == px.DataValidationType.TextLength
    copy = dv.copy()
    assert copy == dv
    copy.input_title = "x"
    assert copy != dv
    assert dv.set(error_title="E") is dv and dv.error_title == "E"


def test_rule_objects():
    rule = R.between(1, "=B1")
    assert rule.criteria == "between"
    assert rule.values == (1, "=B1")
    assert repr(rule) == "DataValidationRule.between(1, '=B1')"
    first = R.equal_to(px.Formula("=A1")).values[0]
    assert isinstance(first, px.Formula) and first.formula == "=A1"
    assert R.greater_than(1) == R.greater_than(1) and R.greater_than(1) != R.less_than(1)
    with pytest.raises(TypeError):
        R.equal_to(object())  # pyrefly: ignore[bad-argument-type]


def test_validation_errors():
    with pytest.raises(px.ParameterError, match="ints"):
        px.DataValidation(whole_number=R.between(1.5, 3))
    with pytest.raises(px.ParameterError, match="ints"):
        px.DataValidation(whole_number=R.equal_to(2**40))
    with pytest.raises(px.ParameterError, match="numbers"):
        px.DataValidation(decimal_number=R.equal_to("text"))
    with pytest.raises(px.ParameterError, match="dates"):
        px.DataValidation(date=R.equal_to(5))
    with pytest.raises(px.ParameterError, match="times"):
        px.DataValidation(time=R.equal_to(dt.date(2024, 1, 1)))
    with pytest.raises(px.ParameterError, match="non-negative"):
        px.DataValidation(text_length=R.greater_than(-1))
    with pytest.raises(px.ParameterError, match="32"):
        px.DataValidation(input_title="x" * 33)
    with pytest.raises(px.ParameterError, match="255"):
        px.DataValidation(error_message="x" * 256)
    with pytest.raises(px.ParameterError, match="255"):
        px.DataValidation(list_strings=["abcdefghij"] * 30)
    with pytest.raises(px.ParameterError, match="at least one"):
        px.DataValidation(list_strings=[])
    with pytest.raises(px.ParameterError, match="empty"):
        px.DataValidation(custom="=")
    with pytest.raises(px.ParameterError):
        px.DataValidation(date=R.equal_to(dt.datetime(2024, 1, 1, tzinfo=dt.UTC)))
    with pytest.raises(TypeError, match="unexpected keyword"):
        px.DataValidation(bogus=True)  # pyrefly: ignore[unexpected-keyword]
    ws = px.Workbook().add_worksheet()
    with pytest.raises(TypeError):
        ws.add_data_validation("A1", "not a validation")  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(px.RangeError):
        ws.add_data_validation("B2:A1", px.DataValidation())
