"""Unit tests for the PyExcel formula engine.  Run:  python -m unittest -v"""

import os
import tempfile
import unittest

import formula_engine as fe
from formula_engine import CellError, Workbook


def make(cells, sheet="Sheet1", wb=None):
    wb = wb or Workbook()
    for addr, raw in cells.items():
        r, c = fe.parse_address(addr)
        wb.set_raw(sheet, r, c, raw)
    return wb


def val(wb, addr, sheet="Sheet1"):
    r, c = fe.parse_address(addr)
    return wb.get_value(sheet, r, c)


class AddressTests(unittest.TestCase):
    def test_column_names(self):
        self.assertEqual(fe.col_name(0), "A")
        self.assertEqual(fe.col_name(25), "Z")
        self.assertEqual(fe.col_name(26), "AA")
        self.assertEqual(fe.col_index("AB"), 27)
        self.assertEqual(fe.parse_address("b3"), (2, 1))
        self.assertEqual(fe.parse_address("C4:A1"), (0, 0, 3, 2))
        self.assertIsNone(fe.parse_address("A0"))


class LiteralTests(unittest.TestCase):
    def test_literals(self):
        self.assertEqual(fe.parse_literal("42"), (42.0, None))
        self.assertEqual(fe.parse_literal("1,234.5"), (1234.5, "#,##0.00"))
        self.assertEqual(fe.parse_literal("50%"), (0.5, "0%"))
        self.assertEqual(fe.parse_literal("true"), (True, None))
        self.assertEqual(fe.parse_literal("'123"), ("123", None))
        self.assertEqual(fe.parse_literal("hello"), ("hello", None))
        self.assertEqual(fe.parse_literal("nan"), ("nan", None))
        value, fmt = fe.parse_literal("21/10/2026")
        self.assertEqual(fmt, "dd/mm/yyyy")
        self.assertEqual(fe.format_value(value, fmt), "21/10/2026")


class FormulaTests(unittest.TestCase):
    def setUp(self):
        self.wb = make({"A1": "10", "A2": "20", "A3": "30", "B1": "abc", "B2": "", "C1": "2"})

    def ev(self, formula):
        self.wb.set_raw("Sheet1", 99, 0, formula)
        return self.wb.get_value("Sheet1", 99, 0)

    def test_arithmetic_and_precedence(self):
        self.assertEqual(self.ev("=1+2*3"), 7)
        self.assertEqual(self.ev("=(1+2)*3"), 9)
        self.assertEqual(self.ev("=-2^2"), 4)  # Excel: negation binds tighter than ^
        self.assertEqual(self.ev("=2^3^2"), 64)
        self.assertEqual(self.ev("=50%*A1"), 5)
        self.assertEqual(self.ev("=A1/C1+A2"), 25)
        self.assertAlmostEqual(self.ev("=0.1+0.2"), 0.3)

    def test_text_and_comparison(self):
        self.assertEqual(self.ev('="x"&A1&"y"'), "x10y")
        self.assertIs(self.ev("=A1<A2"), True)
        self.assertIs(self.ev('=B1="ABC"'), True)
        self.assertIs(self.ev("=A1<>10"), False)
        self.assertIs(self.ev("=Z50=0"), True)  # blank compares as 0

    def test_errors(self):
        self.assertEqual(self.ev("=A1/0"), CellError(fe.DIV0))
        self.assertEqual(self.ev("=B1+1"), CellError(fe.VALUE))
        self.assertEqual(self.ev("=FOO(1)"), CellError(fe.NAME))
        self.assertEqual(self.ev("=1+"), CellError(fe.SYNTAX))
        self.assertEqual(self.ev("=Nope!A1"), CellError(fe.REF))
        self.assertEqual(self.ev("=SQRT(-1)"), CellError(fe.NUM))
        self.assertEqual(self.ev("=(1/0)+1"), CellError(fe.DIV0))

    def test_aggregates(self):
        self.assertEqual(self.ev("=SUM(A1:A3)"), 60)
        self.assertEqual(self.ev("=SUM(A1:B3, 5)"), 65)  # text in ranges is ignored
        self.assertEqual(self.ev("=AVERAGE(A1:A3)"), 20)
        self.assertEqual(self.ev("=MIN(A1:A3)"), 10)
        self.assertEqual(self.ev("=MAX(A1:A3)"), 30)
        self.assertEqual(self.ev("=COUNT(A1:C3)"), 4)
        self.assertEqual(self.ev("=COUNTA(A1:C3)"), 5)
        self.assertEqual(self.ev("=COUNTBLANK(A1:C3)"), 4)
        self.assertEqual(self.ev("=PRODUCT(A1,C1)"), 20)
        self.assertEqual(self.ev("=MEDIAN(A1:A3)"), 20)
        self.assertEqual(self.ev("=SUM(A1:A99,C1:C10000)"), 62)  # large sparse range
        self.assertEqual(self.ev("=SUM(A1:A10000)"), CellError(fe.CIRC))  # includes its own cell

    def test_math(self):
        self.assertEqual(self.ev("=ROUND(2.675,2)"), 2.68)
        self.assertEqual(self.ev("=ROUND(-2.5,0)"), -3)
        self.assertEqual(self.ev("=ROUND(1234,-2)"), 1200)
        self.assertEqual(self.ev("=ROUNDUP(1.21,1)"), 1.3)
        self.assertEqual(self.ev("=ROUNDDOWN(1.29,1)"), 1.2)
        self.assertEqual(self.ev("=INT(-1.5)"), -2)
        self.assertEqual(self.ev("=MOD(-7,3)"), 2)
        self.assertEqual(self.ev("=POWER(2,10)"), 1024)
        self.assertEqual(self.ev("=ABS(-3)"), 3)

    def test_logic(self):
        self.assertEqual(self.ev('=IF(A1>5,"lớn","nhỏ")'), "lớn")
        self.assertEqual(self.ev("=IF(A1>50,1)"), False)
        self.assertEqual(self.ev("=IF(TRUE,,1)"), 0)
        self.assertEqual(self.ev('=IFERROR(A1/0,"lỗi")'), "lỗi")
        self.assertIs(self.ev("=AND(A1>5,A2>5)"), True)
        self.assertIs(self.ev("=OR(A1>50,A2>50)"), False)
        self.assertIs(self.ev("=NOT(FALSE)"), True)
        self.assertIs(self.ev("=ISERROR(1/0)"), True)
        self.assertIs(self.ev("=ISBLANK(Z1)"), True)
        self.assertIs(self.ev("=ISNUMBER(A1)"), True)
        self.assertEqual(self.ev('=IFS(A1>20,"a",A1>5,"b")'), "b")
        self.assertEqual(self.ev("=IF(A1>5;1;2)"), 1)  # ';' separators work too

    def test_text_functions(self):
        self.assertEqual(self.ev('=CONCATENATE("a",A1,"b")'), "a10b")
        self.assertEqual(self.ev('=CONCAT(A1:A3)'), "102030")
        self.assertEqual(self.ev('=TEXTJOIN("-",TRUE,A1:B2)'), "10-abc-20")
        self.assertEqual(self.ev('=LEN("Xin chào")'), 8)
        self.assertEqual(self.ev('=UPPER(B1)'), "ABC")
        self.assertEqual(self.ev('=LEFT("Hello",2)&RIGHT("Hello",3)&MID("Hello",2,3)'), "Helloell")
        self.assertEqual(self.ev('=TRIM("  a   b ")'), "a b")
        self.assertEqual(self.ev('=SUBSTITUTE("a-b-c","-","+")'), "a+b+c")
        self.assertEqual(self.ev('=FIND("l","Hello")'), 3)
        self.assertEqual(self.ev('=TEXT(1234.5,"#,##0.00")'), "1,234.50")
        self.assertEqual(self.ev('=VALUE("12.5")'), 12.5)
        self.assertEqual(self.ev('="Total: "&1/4'), "Total: 0.25")

    def test_dates(self):
        self.assertEqual(self.ev("=DATE(2026,10,21)"), fe.parse_literal("21/10/2026")[0])
        self.assertEqual(self.ev("=YEAR(DATE(2026,10,21))"), 2026)
        self.assertEqual(self.ev("=MONTH(DATE(2026,14,1))"), 2)
        self.assertEqual(self.ev("=WEEKDAY(DATE(2026,10,18))"), 1)  # a Sunday
        self.wb.set_raw("Sheet1", 99, 0, "=DATE(2026,10,21)")
        self.assertEqual(fe.format_value(*self.wb.cell_display("Sheet1", 99, 0)), "21/10/2026")

    def test_lookup(self):
        wb = make({"A1": "apple", "B1": "1", "A2": "banana", "B2": "2", "A3": "cherry", "B3": "3",
                   "D1": "10", "D2": "20", "D3": "30"})
        wb.set_raw("Sheet1", 9, 0, '=VLOOKUP("banana",A1:B3,2,FALSE)')
        self.assertEqual(val(wb, "A10"), 2)
        wb.set_raw("Sheet1", 9, 0, "=VLOOKUP(25,D1:D3,1)")
        self.assertEqual(val(wb, "A10"), 20)
        wb.set_raw("Sheet1", 9, 0, '=VLOOKUP("kiwi",A1:B3,2,FALSE)')
        self.assertEqual(val(wb, "A10"), CellError(fe.NA))
        wb.set_raw("Sheet1", 9, 0, '=INDEX(A1:B3,3,2)')
        self.assertEqual(val(wb, "A10"), 3)
        wb.set_raw("Sheet1", 9, 0, '=MATCH("cherry",A1:A3,0)')
        self.assertEqual(val(wb, "A10"), 3)

    def test_conditional(self):
        wb = make({"A1": "x", "B1": "5", "A2": "y", "B2": "15", "A3": "x", "B3": "25", "A4": "xyz", "B4": "1"})
        cases = {
            '=COUNTIF(B1:B4,">10")': 2,
            '=SUMIF(A1:A4,"x",B1:B4)': 30,
            '=SUMIF(A1:A4,"x*",B1:B4)': 31,
            '=AVERAGEIF(A1:A4,"x",B1:B4)': 15,
            '=COUNTIFS(A1:A4,"x",B1:B4,">10")': 1,
            '=SUMIFS(B1:B4,A1:A4,"<>y")': 31,
            '=SUMPRODUCT(B1:B2,B3:B4)': 140,
        }
        for formula, expected in cases.items():
            wb.set_raw("Sheet1", 9, 0, formula)
            self.assertEqual(val(wb, "A10"), expected, formula)

    def test_more_math_and_statistics(self):
        wb = make({"A1": "3", "A2": "1", "A3": "4", "A4": "1", "A5": "5", "A6": "x"})
        cases = {
            "=LARGE(A1:A6,2)": 4, "=SMALL(A1:A6,1)": 1, "=LARGE(A1:A6,9)": CellError(fe.NUM),
            "=RANK(4,A1:A5)": 2, "=RANK(4,A1:A5,1)": 4, "=RANK.AVG(1,A1:A5,1)": 1.5, "=RANK(9,A1:A5)": CellError(fe.NA),
            "=MODE(A1:A5)": 1, "=PERCENTILE(A1:A5,0.5)": 3, "=QUARTILE(A1:A5,1)": 1, "=QUARTILE(A1:A5,4)": 5,
            "=VARP(A1:A5)": 2.56, "=SUMSQ(A1:A3)": 26, "=MROUND(7,3)": 6, "=MROUND(-7,3)": CellError(fe.NUM),
            "=QUOTIENT(-7,2)": -3, "=GCD(12,18)": 6, "=LCM(4,6)": 12, "=FACT(5)": 120, "=COMBIN(5,2)": 10,
            "=EVEN(1.5)": 2, "=EVEN(-1)": -2, "=ODD(2)": 3, "=ODD(0)": 1, "=DEGREES(PI())": 180,
            "=SIN(PI()/2)": 1, "=ATAN2(1,1)*4": 3.141592653589793,
        }
        for formula, expected in cases.items():
            wb.set_raw("Sheet1", 9, 0, formula)
            if isinstance(expected, CellError):
                self.assertEqual(val(wb, "A10"), expected, formula)
            else:
                self.assertAlmostEqual(val(wb, "A10"), expected, msg=formula)

    def test_conditional_ifs(self):
        wb = make({"A1": "x", "B1": "5", "A2": "y", "B2": "15", "A3": "x", "B3": "25", "A4": "x", "B4": "1"})
        cases = {
            '=AVERAGEIFS(B1:B4,A1:A4,"x",B1:B4,">2")': 15,
            '=MAXIFS(B1:B4,A1:A4,"x")': 25, '=MINIFS(B1:B4,A1:A4,"x")': 1, '=MAXIFS(B1:B4,A1:A4,"z")': 0,
            '=AVERAGEIFS(B1:B4,A1:A4,"z")': CellError(fe.DIV0),
        }
        for formula, expected in cases.items():
            wb.set_raw("Sheet1", 9, 0, formula)
            self.assertEqual(val(wb, "A10"), expected, formula)

    def test_subtotal_skips_hidden_rows_and_subtotals(self):
        wb = make({"A1": "1", "A2": "2", "A3": "4", "A4": "=SUBTOTAL(9,A1:A3)", "A5": "8"})
        wb.set_raw("Sheet1", 9, 0, "=SUBTOTAL(9,A1:A5)")
        self.assertEqual(val(wb, "A10"), 15)  # A4 is a subtotal, so it is not counted twice
        sh = wb.sheets["Sheet1"]
        sh.hidden_rows.add(1)
        wb.invalidate()
        self.assertEqual(val(wb, "A10"), 15)  # code 9 keeps rows hidden by hand
        wb.set_raw("Sheet1", 9, 0, "=SUBTOTAL(109,A1:A5)")
        self.assertEqual(val(wb, "A10"), 13)
        sh.filter = {"range": (0, 0, 4, 0), "criteria": {}, "hidden": {2}}
        wb.set_raw("Sheet1", 9, 0, "=SUBTOTAL(9,A1:A5)")
        self.assertEqual(val(wb, "A10"), 11)  # rows hidden by a filter are always skipped
        wb.set_raw("Sheet1", 9, 0, "=SUBTOTAL(2,A1:A5)")
        self.assertEqual(val(wb, "A10"), 3)

    def test_xlookup_and_friends(self):
        wb = make({"A1": "apple", "B1": "10", "C1": "a", "A2": "banana", "B2": "20", "C2": "b",
                   "A3": "cherry", "B3": "30", "C3": "c", "A4": "banana", "B4": "40", "C4": "d",
                   "E1": "1", "F1": "2", "G1": "3", "E2": "x", "F2": "y", "G2": "z"})
        cases = {
            '=XLOOKUP("banana",A1:A4,B1:B4)': 20,
            '=XLOOKUP("banana",A1:A4,B1:B4,,0,-1)': 40,
            '=XLOOKUP("kiwi",A1:A4,B1:B4)': CellError(fe.NA),
            '=XLOOKUP("kiwi",A1:A4,B1:B4,"không có")': "không có",
            '=XLOOKUP(25,B1:B4,A1:A4,,1)': "cherry",
            '=XLOOKUP(25,B1:B4,A1:A4,,-1)': "banana",
            '=XLOOKUP("ch*",A1:A4,C1:C4,,2)': "c",
            '=SUM(XLOOKUP("cherry",A1:A4,B1:C4))': 30,  # a whole row comes back
            "=XLOOKUP(2,E1:G1,E2:G2)": "y",
            "=XLOOKUP(2,E1:G1,E2:F2)": CellError(fe.VALUE),
            '=XMATCH("cherry",A1:A4)': 3, '=XMATCH("banana",A1:A4,0,-1)': 4,
            "=SUM(OFFSET(A1,1,1,2,1))": 50, "=OFFSET(B1,3,0)": 40, "=OFFSET(A1,-1,0)": CellError(fe.REF),
            '=INDIRECT("B"&2)': 20, '=SUM(INDIRECT("B1:B3"))': 60, '=INDIRECT("R3C2",FALSE)': 30,
            '=INDIRECT("nope")': CellError(fe.REF),
            '=SWITCH(C2,"a",1,"b",2)': 2, '=SWITCH(C4,"a",1,"b",2,"khác")': "khác", '=SWITCH(9,1,2)': CellError(fe.NA),
            '=ISFORMULA(A1)': False, '=ISEVEN(4)': True, '=ISODD(-3)': True, '=ISNONTEXT(B1)': True,
            '=IFNA(NA(),"ok")': "ok",
            '=TEXTBEFORE("a-b-c","-")': "a", '=TEXTAFTER("a-b-c","-")': "b-c", '=TEXTBEFORE("a-b-c","-",-1)': "a-b",
            '=TEXTAFTER("a-b-c","-",2)': "c", '=TEXTAFTER("a-b","+")': CellError(fe.NA),
        }
        for formula, expected in cases.items():
            wb.set_raw("Sheet1", 9, 0, formula)
            self.assertEqual(val(wb, "A10"), expected, formula)
        wb.set_raw("Sheet1", 9, 1, '=INDIRECT("A10")')  # follows the formula above
        self.assertEqual(val(wb, "B10"), CellError(fe.NA))

    def test_more_dates(self):
        d = fe.parse_literal

        def ev(formula):
            self.wb.set_raw("Sheet1", 99, 0, formula)
            return self.wb.get_value("Sheet1", 99, 0)

        def text(formula):
            ev(formula)
            return fe.format_value(*self.wb.cell_display("Sheet1", 99, 0))

        self.assertEqual(text("=EDATE(DATE(2026,1,31),1)"), "28/02/2026")
        self.assertEqual(text("=EDATE(DATE(2026,3,15),-3)"), "15/12/2025")
        self.assertEqual(text("=EOMONTH(DATE(2024,1,10),1)"), "29/02/2024")
        self.assertEqual(text("=WORKDAY(DATE(2026,10,23),1)"), "26/10/2026")  # Friday -> Monday
        self.assertEqual(ev("=WORKDAY(DATE(2026,10,26),-1)"), d("23/10/2026")[0])
        self.assertEqual(ev("=NETWORKDAYS(DATE(2026,10,1),DATE(2026,10,31))"), 22)
        self.assertEqual(ev("=NETWORKDAYS(DATE(2026,10,31),DATE(2026,10,1))"), -22)
        self.wb.set_raw("Sheet1", 0, 4, "02/10/2026")
        self.assertEqual(ev("=NETWORKDAYS(DATE(2026,10,1),DATE(2026,10,31),E1)"), 21)
        self.assertEqual(ev("=WORKDAY(DATE(2026,10,1),1,E1)"), d("05/10/2026")[0])
        self.assertEqual(ev('=DATEDIF(DATE(2000,5,20),DATE(2026,10,21),"Y")'), 26)
        self.assertEqual(ev('=DATEDIF(DATE(2000,5,20),DATE(2026,10,21),"M")'), 317)
        self.assertEqual(ev('=DATEDIF(DATE(2000,5,20),DATE(2026,10,21),"YM")'), 5)
        self.assertEqual(ev('=DATEDIF(DATE(2000,5,20),DATE(2026,10,21),"MD")'), 1)
        self.assertEqual(ev('=DATEDIF(DATE(2026,1,1),DATE(2026,3,1),"D")'), 59)
        self.assertEqual(ev('=DATEDIF(DATE(2026,3,1),DATE(2026,1,1),"D")'), CellError(fe.NUM))
        self.assertEqual(ev('=DATEDIF(DATE(2025,12,25),DATE(2026,1,5),"YD")'), 11)
        self.assertEqual(text("=TIME(14,30,0)"), "14:30")
        self.assertEqual(ev("=SECOND(TIME(1,2,3))"), 3)
        self.assertEqual(ev('=DATEVALUE("21/10/2026")'), d("21/10/2026")[0])
        self.assertEqual(ev('=DATEVALUE("abc")'), CellError(fe.VALUE))
        self.assertEqual(ev("=WEEKNUM(DATE(2026,1,3))"), 1)   # Saturday of the first week
        self.assertEqual(ev("=WEEKNUM(DATE(2026,1,4))"), 2)   # Sunday starts week 2
        self.assertEqual(ev("=WEEKNUM(DATE(2026,1,4),2)"), 1)  # weeks starting on Monday
        self.assertEqual(ev("=ISOWEEKNUM(DATE(2026,10,21))"), 43)

    def test_every_function_is_documented(self):
        from function_info import CATEGORIES, FUNCTION_INFO
        self.assertEqual(set(fe.function_names()) - set(FUNCTION_INFO), set())
        self.assertTrue(all(info[0] in CATEGORIES for info in FUNCTION_INFO.values()))

    def test_circular(self):
        wb = make({"A1": "=B1+1", "B1": "=A1+1", "C1": "=C1"})
        self.assertEqual(val(wb, "A1"), CellError(fe.CIRC))
        self.assertEqual(val(wb, "B1"), CellError(fe.CIRC))
        self.assertEqual(val(wb, "C1"), CellError(fe.CIRC))

    def test_recalculation_after_edit(self):
        wb = make({"A1": "1", "A2": "=A1*10", "A3": "=A2+1"})
        self.assertEqual(val(wb, "A3"), 11)
        wb.set_raw("Sheet1", 0, 0, "5")
        self.assertEqual(val(wb, "A3"), 51)

    def test_long_chain(self):
        wb = Workbook()
        wb.set_raw("Sheet1", 0, 0, "1")
        for r in range(1, 1500):
            wb.set_raw("Sheet1", r, 0, f"=A{r}+1")
        self.assertEqual(wb.get_value("Sheet1", 1499, 0), 1500)

    def test_cross_sheet(self):
        wb = Workbook(("Sheet1", "Data Sheet"))
        make({"A1": "7"}, sheet="Data Sheet", wb=wb)
        make({"A1": "='Data Sheet'!A1*2", "A2": "=SUM('data sheet'!A1:A5)"}, wb=wb)
        self.assertEqual(val(wb, "A1"), 14)
        self.assertEqual(val(wb, "A2"), 7)
        wb.rename_sheet("Data Sheet", "Data")
        self.assertEqual(wb.get_raw("Sheet1", 0, 0), "=Data!A1*2")
        self.assertEqual(val(wb, "A1"), 14)


class RewriteTests(unittest.TestCase):
    def test_normalize(self):
        self.assertEqual(fe.normalize_formula("=sum(a1:b2;3"), "=SUM(A1:B2,3)")
        self.assertEqual(fe.normalize_formula('="abc"&a1'), '="abc"&A1')
        self.assertEqual(fe.normalize_formula("hello"), "hello")

    def test_shift(self):
        self.assertEqual(fe.shift_formula("=A1+$B$1+B$2+$C3", 2, 1), "=B3+$B$1+C$2+$C5")
        self.assertEqual(fe.shift_formula("=SUM(A1:A3)", 1, 0), "=SUM(A2:A4)")
        self.assertEqual(fe.shift_formula("=A1", -1, 0), "=#REF!")
        self.assertEqual(fe.shift_formula('="A1"&A1', 0, 1), '="A1"&B1')
        self.assertEqual(fe.shift_formula("=Sheet2!A1", 1, 0), "=Sheet2!A2")

    def test_insert_delete_rows(self):
        wb = make({"A1": "1", "A2": "2", "A3": "3", "B1": "=SUM(A1:A3)", "B2": "=A3*2"})
        wb.insert_rows("Sheet1", 1)
        self.assertEqual(wb.get_raw("Sheet1", 0, 1), "=SUM(A1:A4)")
        self.assertEqual(wb.get_raw("Sheet1", 2, 1), "=A4*2")
        self.assertEqual(val(wb, "B1"), 6)
        wb.delete_rows("Sheet1", 3)  # removes the row holding 3
        self.assertEqual(wb.get_raw("Sheet1", 0, 1), "=SUM(A1:A3)")
        self.assertEqual(wb.get_raw("Sheet1", 2, 1), "=#REF!*2")
        self.assertEqual(val(wb, "B1"), 3)

    def test_insert_delete_cols(self):
        wb = make({"A1": "1", "B1": "2", "C1": "=A1+B1"})
        wb.insert_cols("Sheet1", 1)
        self.assertEqual(wb.get_raw("Sheet1", 0, 3), "=A1+C1")
        wb.delete_cols("Sheet1", 0)
        self.assertEqual(wb.get_raw("Sheet1", 0, 2), "=#REF!+B1")

    def test_fill_series(self):
        self.assertEqual(fe.fill_series(["1", "3"])(4), "9")
        self.assertIsNone(fe.fill_series(["5"]))
        self.assertEqual(fe.fill_series(["Item 1"])(2), "Item 3")
        self.assertEqual(fe.fill_series(["21/10/2026"])(11), "01/11/2026")
        self.assertIsNone(fe.fill_series(["=A1"]))

    def test_formats(self):
        self.assertEqual(fe.format_number(1234567.891, "#,##0.00"), "1,234,567.89")
        self.assertEqual(fe.format_number(0.256, "0.0%"), "25.6%")
        self.assertEqual(fe.format_number(-1500, "#,##0 ₫"), "-1,500 ₫")
        self.assertEqual(fe.format_number(1.5, "$#,##0.00"), "$1.50")
        self.assertEqual(fe.format_general(1 / 3), "0.3333333333")
        self.assertEqual(fe.format_general(1e20), "1E+20")
        self.assertEqual(fe.adjust_decimals(None, 1.5, 1), "0.00")
        self.assertEqual(fe.adjust_decimals("#,##0.00", 1, -1), "#,##0.0")


class FileTests(unittest.TestCase):
    def roundtrip(self, ext):
        wb = Workbook(("Sheet1", "Khác"))
        make({"A1": "Tên", "B1": "12.5", "A2": "=B1*2", "A3": "21/10/2026", "A4": "'007"}, wb=wb)
        wb.set_format("Sheet1", 0, 0, {"bold": True, "bg": "#ffff00", "align": "center"})
        wb.sheets["Sheet1"].col_widths[0] = 150
        make({"A1": "=Sheet1!B1+1"}, sheet="Khác", wb=wb)
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "test" + ext)
            fe.save_workbook(wb, path, active="Khác")
            return fe.load_workbook(path)

    def test_json(self):
        wb = self.roundtrip(".pyxl")
        self.assertEqual(list(wb.sheets), ["Sheet1", "Khác"])
        self.assertEqual(wb.active, "Khác")
        self.assertEqual(val(wb, "A2"), 25)
        self.assertEqual(val(wb, "A1", "Khác"), 13.5)
        self.assertEqual(wb.get_format("Sheet1", 0, 0)["bg"], "#ffff00")
        self.assertEqual(wb.sheets["Sheet1"].col_widths[0], 150)

    @unittest.skipUnless(fe.HAS_XLSX, "openpyxl not installed")
    def test_xlsx(self):
        wb = self.roundtrip(".xlsx")
        self.assertEqual(list(wb.sheets), ["Sheet1", "Khác"])
        self.assertEqual(wb.active, "Khác")
        self.assertEqual(wb.get_raw("Sheet1", 1, 0), "=B1*2")
        self.assertEqual(val(wb, "A2"), 25)
        self.assertEqual(val(wb, "A1", "Khác"), 13.5)
        self.assertEqual(val(wb, "A4"), "007")
        self.assertEqual(fe.format_value(*wb.cell_display("Sheet1", 2, 0)), "21/10/2026")
        fmt = wb.get_format("Sheet1", 0, 0)
        self.assertTrue(fmt["bold"])
        self.assertEqual(fmt["bg"], "#ffff00")
        self.assertEqual(fmt["align"], "center")
        self.assertEqual(wb.sheets["Sheet1"].col_widths[0], 150)

    def test_csv(self):
        wb = self.roundtrip(".csv")
        self.assertEqual(list(wb.sheets), ["Sheet1"])  # CSV holds only the active sheet's values
        self.assertEqual(val(wb, "A1"), 13.5)


class NumberFormatTests(unittest.TestCase):
    def test_sections_and_special_formats(self):
        acc = '#,##0 ₫;-#,##0 ₫;"-" ₫'
        self.assertEqual(fe.format_number(1500, acc), "1,500 ₫")
        self.assertEqual(fe.format_number(-1500, acc), "-1,500 ₫")
        self.assertEqual(fe.format_number(0, acc), "- ₫")
        self.assertEqual(fe.format_number(-3, "#,##0;(#,##0)"), "(3)")
        self.assertEqual(fe.format_number(12345.678, "0.00E+00"), "1.23E+04")
        self.assertEqual(fe.format_number(0.00012, "0.0E+00"), "1.2E-04")
        self.assertEqual(fe.format_number(1.25, "# ?/?"), "1 1/4")
        self.assertEqual(fe.format_number(0.5, "# ?/?"), "1/2")
        serial = fe.parse_literal("21/10/2026")[0]
        self.assertEqual(fe.format_number(serial, 'dddd, "ngày" dd "tháng" mm "năm" yyyy'),
                         "Thứ Tư, ngày 21 tháng 10 năm 2026")
        self.assertEqual(fe.format_number(serial + 0.5, "hh:mm:ss"), "12:00:00")
        self.assertEqual(fe.adjust_decimals('#,##0 ₫;-#,##0 ₫;"-" ₫', 1, 1), '#,##0.0 ₫;-#,##0.0 ₫;"-" ₫')

    def test_text_format_keeps_typed_text(self):
        wb = make({"A1": "007", "A2": "=1+1"})
        wb.set_format("Sheet1", 0, 0, {"num": "@"})
        wb.set_format("Sheet1", 1, 0, {"num": "@"})
        self.assertEqual(val(wb, "A1"), "007")
        self.assertEqual(val(wb, "A2"), "=1+1")
        wb.set_format("Sheet1", 0, 0, None)
        self.assertEqual(val(wb, "A1"), 7)

    def test_value_to_raw(self):
        self.assertEqual(fe.value_to_raw(3.0), "3")
        self.assertEqual(fe.value_to_raw("12"), "'12")
        self.assertEqual(fe.value_to_raw("abc"), "abc")
        self.assertEqual(fe.value_to_raw(True), "TRUE")
        self.assertIsNone(fe.value_to_raw(""))


class LayoutTests(unittest.TestCase):
    def test_merges_follow_row_and_column_changes(self):
        wb = make({"A1": "Tiêu đề"})
        sh = wb.sheets["Sheet1"]
        sh.set_merges([(0, 0, 0, 2), (4, 0, 5, 1)])
        self.assertEqual(sh.merge_at(0, 1), (0, 0, 0, 2))
        wb.insert_rows("Sheet1", 1)
        self.assertEqual(sh.merges, [(0, 0, 0, 2), (5, 0, 6, 1)])
        wb.insert_cols("Sheet1", 1)
        self.assertEqual(sh.merges, [(0, 0, 0, 3), (5, 0, 6, 2)])
        wb.delete_rows("Sheet1", 5, 2)
        self.assertEqual(sh.merges, [(0, 0, 0, 3)])

    def test_insert_and_delete_cells_shift_only_the_band(self):
        wb = make({"A1": "1", "A2": "2", "B1": "10", "B2": "20", "C1": "=SUM(A1:A2)", "C2": "=B2", "D1": "=A2"})
        wb.insert_cells("Sheet1", 0, 0, 0, 0, "down")  # only column A moves down
        self.assertEqual(wb.get_raw("Sheet1", 1, 0), "1")
        self.assertEqual(wb.get_raw("Sheet1", 0, 1), "10")  # column B untouched
        self.assertEqual(wb.get_raw("Sheet1", 0, 2), "=SUM(A2:A3)")
        self.assertEqual(wb.get_raw("Sheet1", 1, 2), "=B2")
        self.assertEqual(wb.get_raw("Sheet1", 0, 3), "=A3")
        wb.delete_cells("Sheet1", 0, 1, 0, 1, "left")  # B1 removed, C1.. shift left on row 1 only
        self.assertEqual(wb.get_raw("Sheet1", 0, 1), "=SUM(A2:A3)")
        self.assertEqual(wb.get_raw("Sheet1", 1, 1), "20")
        self.assertEqual(val(wb, "B1"), 3)

    def test_sizes_and_hidden_follow_rows(self):
        wb = Workbook()
        sh = wb.sheets["Sheet1"]
        sh.row_heights, sh.custom_rows, sh.hidden_rows = {3: 40}, {3}, {5}
        sh.hidden_cols, sh.col_widths = {2}, {2: 120}
        wb.insert_rows("Sheet1", 0, 2)
        wb.delete_cols("Sheet1", 0)
        self.assertEqual((sh.row_heights, sh.custom_rows, sh.hidden_rows), ({5: 40}, {5}, {7}))
        self.assertEqual((sh.hidden_cols, sh.col_widths), ({1}, {1: 120}))

    def test_snapshot_restore_and_copy_sheet(self):
        wb = make({"A1": "5", "B1": "=A1*2"})
        sh = wb.sheets["Sheet1"]
        sh.set_merges([(2, 0, 2, 1)])
        before = wb.snapshot()
        wb.delete_rows("Sheet1", 0, 5)
        self.assertNotEqual(wb.snapshot(), before)
        wb.restore(before)
        self.assertEqual(val(wb, "B1"), 10)
        self.assertEqual(wb.sheets["Sheet1"].merge_at(2, 1), (2, 0, 2, 1))
        new = wb.copy_sheet("Sheet1")
        self.assertEqual(new, "Sheet1 (2)")
        wb.set_raw(new, 0, 0, "7")
        self.assertEqual((val(wb, "B1"), val(wb, "B1", new)), (10, 14))


class ConditionalFormatTests(unittest.TestCase):
    def setUp(self):
        self.wb = make({f"A{i}": str(v) for i, v in enumerate([5, 15, 25, 35, 15, 0], 1)})
        self.wb.set_raw("Sheet1", 6, 0, "chuối")
        red = {"bg": "#ffc7ce", "fg": "#9c0006"}
        self.style = red

    def styles(self, rule):
        rule = dict(rule, range=(0, 0, 6, 0), style=self.style)
        self.wb.set_cond_formats("Sheet1", [rule])
        return [bool(self.wb.cond_format("Sheet1", r, 0)) for r in range(7)]

    def test_rules(self):
        T, F = True, False
        self.assertEqual(self.styles({"type": "cell", "op": ">", "v1": "20"}), [F, F, T, T, F, F, T])
        self.assertEqual(self.styles({"type": "cell", "op": "between", "v1": "10", "v2": "30"}),
                         [F, T, T, F, T, F, F])
        self.assertEqual(self.styles({"type": "text", "text": "UỐ"}), [F, F, F, F, F, F, T])
        self.assertEqual(self.styles({"type": "dup"}), [F, T, F, F, T, F, F])
        self.assertEqual(self.styles({"type": "top", "n": 2}), [F, F, T, T, F, F, F])
        self.assertEqual(self.styles({"type": "bottom", "n": 1}), [F, F, F, F, F, T, F])
        self.assertEqual(self.styles({"type": "above"}), [F, F, T, T, F, F, F])  # average 15.83
        bar = dict(type="databar", color="#638ec6", range=(0, 0, 6, 0))
        self.wb.set_cond_formats("Sheet1", [bar])
        self.assertEqual(self.wb.cond_format("Sheet1", 3, 0)["databar"], (1.0, "#638ec6"))
        self.assertAlmostEqual(self.wb.cond_format("Sheet1", 0, 0)["databar"][0], 5 / 35)
        scale = dict(type="scale", colors=["#f8696b", "#ffeb84", "#63be7b"], range=(0, 0, 6, 0))
        self.wb.set_cond_formats("Sheet1", [scale])
        self.assertEqual(self.wb.cond_format("Sheet1", 5, 0)["bg"], "#f8696b")
        self.assertEqual(self.wb.cond_format("Sheet1", 3, 0)["bg"], "#63be7b")

    def test_rules_follow_values(self):
        self.styles({"type": "cell", "op": ">", "v1": "=A1*4"})  # threshold 20 from a formula
        self.wb.set_raw("Sheet1", 0, 0, "1")                  # threshold now 4
        self.assertTrue(self.wb.cond_format("Sheet1", 1, 0))


class FilterTests(unittest.TestCase):
    def test_filter_hides_rows(self):
        wb = make({"A1": "Loại", "B1": "Số", "A2": "x", "B2": "1", "A3": "y", "B3": "2", "A4": "x", "B4": "3", "A5": ""})
        sh = wb.sheets["Sheet1"]
        sh.filter = {"range": (0, 0, 4, 1), "criteria": {0: {"x"}}, "hidden": set()}
        wb.apply_filter("Sheet1")
        self.assertEqual(sh.filter["hidden"], {2, 4})
        self.assertTrue(sh.row_hidden(2))
        self.assertEqual(wb.filter_values("Sheet1", 0), ["x", "y", ""])
        wb.insert_rows("Sheet1", 0)
        self.assertEqual((sh.filter["range"], sh.filter["hidden"]), ((1, 0, 5, 1), {3, 5}))


class LayoutFileTests(unittest.TestCase):
    def build(self):
        wb = make({"A1": "Báo cáo", "A2": "10", "A3": "20", "B2": "=A2*2", "A4": "x", "A5": "y"})
        sh = wb.sheets["Sheet1"]
        sh.set_merges([(0, 0, 0, 2)])
        wb.set_format("Sheet1", 0, 0, {"font": "Times New Roman", "size": 16, "bold": True, "align": "center",
                                       "valign": "center", "underline": "double"})
        wb.set_format("Sheet1", 1, 0, {"wrap": True, "rotate": -45, "indent": 2, "strike": True,
                                       "bt": ["thin", "#000000"], "bb": ["double", "#ff0000"],
                                       "num": '#,##0 ₫;-#,##0 ₫;"-" ₫'})
        sh.row_heights, sh.custom_rows = {0: 40}, {0}
        sh.hidden_cols = {5}
        sh.tab_color = "#ff0000"
        sh.cond_formats = [{"range": (1, 0, 2, 0), "type": "cell", "op": ">", "v1": "15",
                            "style": {"bg": "#ffc7ce", "fg": "#9c0006"}},
                           {"range": (1, 0, 2, 0), "type": "scale", "colors": ["#ffffff", "#63be7b"]}]
        sh.filter = {"range": (2, 0, 4, 0), "criteria": {0: {"x"}}, "hidden": set()}
        wb.apply_filter("Sheet1")
        return wb

    def check(self, wb):
        sh = wb.sheets["Sheet1"]
        self.assertEqual(sh.merges, [(0, 0, 0, 2)])
        f0, f1 = wb.get_format("Sheet1", 0, 0), wb.get_format("Sheet1", 1, 0)
        self.assertEqual((f0["font"], f0["size"], f0["underline"], f0["valign"]), ("Times New Roman", 16, "double", "center"))
        self.assertEqual((f1["wrap"], f1["rotate"], f1["indent"], f1["strike"]), (True, -45, 2, True))
        self.assertEqual((f1["bt"], f1["bb"]), (["thin", "#000000"], ["double", "#ff0000"]))
        self.assertEqual(fe.format_value(*wb.cell_display("Sheet1", 1, 0)), "10 ₫")
        self.assertEqual(sh.row_heights.get(0), 40)
        self.assertEqual((sh.hidden_cols, sh.tab_color), ({5}, "#ff0000"))
        self.assertEqual([r["type"] for r in sh.cond_formats], ["cell", "scale"])
        self.assertTrue(wb.cond_format("Sheet1", 2, 0)["fg"] == "#9c0006")
        self.assertEqual((sh.filter["range"], sh.filter["hidden"]), ((2, 0, 4, 0), {4}))
        self.assertNotIn(4, sh.hidden_rows)

    def roundtrip(self, ext):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "layout" + ext)
            fe.save_workbook(self.build(), path)
            return fe.load_workbook(path)

    def test_json_layout(self):
        self.check(self.roundtrip(".pyxl"))

    @unittest.skipUnless(fe.HAS_XLSX, "openpyxl not installed")
    def test_xlsx_layout(self):
        self.check(self.roundtrip(".xlsx"))


if __name__ == "__main__":
    unittest.main()
