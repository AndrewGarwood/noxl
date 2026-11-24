from __future__ import annotations

from collections import OrderedDict
from contextlib import nullcontext
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Callable
from unittest import TestCase
from unittest.mock import Mock, call, patch

from openpyxl import Workbook, load_workbook

from noxl.engines import pyxl
from noxl.engines.pyxl.session import PyxlSession
from noxl.engines.pyxl.types import PyxlWorkbook, PyxlWorkbookSource
from noxl.models import HeaderIndex, MappedData, MappedDataSearchOptions


LOADER = "noxl.engines.pyxl.session.load_workbook"
EXPECTED_ROWS = [("key", "value"), ("A", 10), ("B", 20)]
READERS: list[
    tuple[str, Callable[[PyxlWorkbookSource], object], object, str, bool]
] = [
    (
        "visible sheets",
        pyxl.get_visible_sheetnames,
        ["Data"],
        "noxl.engines.pyxl.utils.is_hidden_worksheet",
        False,
    ),
    (
        "sheet info",
        pyxl.get_sheetinfo,
        {
            "Data": pyxl.PyxlSheetInfo("visible", 0),
            "Hidden": pyxl.PyxlSheetInfo("hidden", 1),
        },
        "noxl.engines.pyxl.utils.PyxlSheetInfo",
        True,
    ),
    (
        "find sheet",
        lambda src: pyxl.find_sheetname(src, lambda name: name == "Data"),
        "Data",
        "noxl.engines.pyxl.utils.Array.find",
        False,
    ),
    (
        "find header",
        lambda src: pyxl.find_header(
            src, sheetname=lambda name: name == "Data",
            required_headers=["KEY", "VALUE"], unique_sheetname=True
        ),
        HeaderIndex(1, "Data", OrderedDict({"key": 1, "value": 2})),
        "noxl.engines.pyxl.utils.normalize_cell_value",
        False,
    ),
    (
        "rows",
        lambda src: pyxl.read_rows(src, sheetname="Data", origin="A1"),
        EXPECTED_ROWS,
        "noxl.engines.pyxl.worksheet.read.CellAddress.from_any",
        False,
    ),
]


class WorkbookFixture(TestCase):
    def setUp(self) -> None:
        temporary = TemporaryDirectory(prefix="noxl-ownership-")
        self.addCleanup(temporary.cleanup)
        self.path = Path(temporary.name) / "input.xlsx"
        wb = Workbook()
        try:
            ws = wb.active
            assert ws is not None
            ws.title = "Data"
            for row in EXPECTED_ROWS:
                ws.append(row)
            wb.create_sheet("Hidden").sheet_state = "hidden"
            wb.save(self.path)
        finally:
            wb.close()

    def borrowed(self) -> PyxlWorkbook:
        wb = load_workbook(self.path, read_only=True, data_only=False)
        self.addCleanup(wb.close)
        return wb

    def assert_closed(self, wb: PyxlWorkbook) -> None:
        with self.assertRaisesRegex(ValueError, "already closed"):
            list(wb["Data"].values)


class RowCountTests(WorkbookFixture):
    def test_counts_only_rows_with_values_in_the_selected_columns(self) -> None:
        wb = Workbook()
        self.addCleanup(wb.close)
        ws = wb["Sheet"]
        self.assertEqual(pyxl.row_count(ws), 0)
        ws["B1"] = "header"
        ws["A2"] = "left of origin"
        ws["B3"] = ""
        ws["C4"] = 0
        ws["B6"] = False
        ws["B7"] = " "
        ws["D8"] = "outside width"
        ws["C10"].number_format = "0.00"

        self.assertEqual(pyxl.row_count(ws, origin="B2", num_cols=2), 3)
        self.assertEqual(pyxl.row_count(ws, origin="B2"), 4)
        self.assertEqual(pyxl.row_count(ws, origin="B11"), 0)
        predicate = Mock()
        self.assertEqual(pyxl.row_count(ws, origin="E1", is_empty_row=predicate), 0)
        predicate.assert_not_called()

    def test_custom_emptiness_and_unknown_read_only_dimensions(self) -> None:
        wb = self.borrowed()
        ws = wb["Data"]
        ws.reset_dimensions()

        def ignore_low_values(row: pyxl.Row, /) -> bool:
            return all(not isinstance(value, int) or value < 20 for value in row)

        self.assertEqual(pyxl.row_count(ws, origin="B2", is_empty_row=ignore_low_values), 1)
        self.assertEqual(pyxl.row_count(wb, "Data", origin="A2"), 2)
        self.assertEqual(list(ws.values), EXPECTED_ROWS)

    def test_path_counts_in_read_only_mode_and_closes_on_predicate_failure(self) -> None:
        self.assertEqual(pyxl.row_count(self.path, "Data"), 3)
        failure = ValueError("predicate failed")
        for fail in (False, True):
            with self.subTest(failure=fail):
                wb = self.borrowed()
                with patch(LOADER, return_value=wb) as load:
                    if fail:
                        with self.assertRaises(ValueError) as caught:
                            pyxl.row_count(
                                str(self.path), "Data",
                                is_empty_row=Mock(side_effect=failure)
                            )
                        self.assertIs(caught.exception, failure)
                    else:
                        self.assertEqual(pyxl.row_count(str(self.path), "Data"), 3)
                    load.assert_called_once_with(str(self.path), read_only=True, data_only=True)
                self.assert_closed(wb)

    def test_invalid_width_and_origin_fail_before_loading(self) -> None:
        with patch(LOADER) as load:
            for width in (0, -1, True):
                with self.subTest(width=width), self.assertRaisesRegex(ValueError, "num_cols"):
                    pyxl.row_count(self.path, "Data", num_cols=width)
            with self.assertRaisesRegex(ValueError, "origin"):
                pyxl.row_count(self.path, "Data", origin="R0C1")
            load.assert_not_called()


class SessionTests(WorkbookFixture):
    def test_independent_opens_close_once_in_reverse_order_on_every_exit(self) -> None:
        for fail in (False, True):
            with self.subTest(body_failure=fail):
                session = pyxl.PyxlSession()
                error = ValueError("body failed")
                with self.assertRaises(ValueError) if fail else nullcontext():
                    with session:
                        first = session.open_workbook(
                            self.path, read_only=True, data_only=True
                        )
                        self.addCleanup(first.close)
                        second = session.open_workbook(
                            self.path, read_only=True, data_only=False
                        )
                        self.addCleanup(second.close)
                        first_close = self.enterContext(
                            patch.object(first, "close", wraps=first.close)
                        )
                        second_close = self.enterContext(
                            patch.object(second, "close", wraps=second.close)
                        )
                        order = Mock()
                        order.attach_mock(first_close, "first")
                        order.attach_mock(second_close, "second")
                        self.assertIsNot(first, second)
                        self.assertTrue(first.data_only)
                        self.assertFalse(second.data_only)
                        pyxl.get_sheetinfo(first)
                        header = pyxl.find_header(
                            first["Data"], required_headers=["key", "value"]
                        )
                        values = pyxl.read_rows(first["Data"])
                        self.assertEqual(list(first["Data"].values), EXPECTED_ROWS)
                        first_close.assert_not_called()
                        second_close.assert_not_called()
                        if fail:
                            raise error
                self.assertEqual(order.mock_calls, [call.second(), call.first()])
                self.assert_closed(first)
                self.assert_closed(second)
                self.assertEqual(header.headers, ("key", "value"))
                self.assertEqual(values, EXPECTED_ROWS)
                session.__exit__(None, None, None)
                first_close.assert_called_once()
                second_close.assert_called_once()

    def test_failed_later_load_still_closes_earlier_workbooks(self) -> None:
        wb = self.borrowed()
        with patch(LOADER, side_effect=[wb, FileNotFoundError("missing")]), patch.object(
            wb, "close", wraps=wb.close
        ) as closed:
            with self.assertRaises(FileNotFoundError):
                with PyxlSession() as session:
                    session.open_workbook(self.path, read_only=True)
                    session.open_workbook(self.path.with_name("missing.xlsx"))
            closed.assert_called_once()
            self.assert_closed(wb)

    def test_state_guards_and_nested_sessions(self) -> None:
        session = PyxlSession()
        with self.assertRaises(RuntimeError):
            session.open_workbook(self.path)
        with session:
            with self.assertRaises(RuntimeError):
                session.__enter__()
            outer = session.open_workbook(str(self.path), read_only=True)
            self.addCleanup(outer.close)
            with PyxlSession() as inner:
                inner_wb = inner.open_workbook(self.path, read_only=True)
                self.addCleanup(inner_wb.close)
            self.assert_closed(inner_wb)
            self.assertEqual(list(outer["Data"].values), EXPECTED_ROWS)
        with self.assertRaises(RuntimeError):
            session.open_workbook(self.path)
        with self.assertRaises(RuntimeError):
            session.__enter__()

    def test_session_does_not_adopt_external_workbooks(self) -> None:
        wb = self.borrowed()
        with PyxlSession() as session:
            self.assertRaises(TypeError, session.open_workbook, wb)
        self.assertEqual(list(wb["Data"].values), EXPECTED_ROWS)

    def test_cleanup_collects_errors_and_attempts_every_close(self) -> None:
        first, second, third = (Mock(spec=PyxlWorkbook) for _ in range(3))
        first_error, third_error = OSError("first close"), OSError("third close")
        first.close.side_effect = first_error
        third.close.side_effect = third_error
        order = Mock()
        for name, wb in (("first", first), ("second", second), ("third", third)):
            order.attach_mock(wb.close, name)
        session = PyxlSession()
        with patch(LOADER, side_effect=[first, second, third]):
            with self.assertRaises(ExceptionGroup) as caught:
                with session:
                    for _ in range(3):
                        session.open_workbook(self.path)
        self.assertEqual(order.mock_calls, [call.third(), call.second(), call.first()])
        self.assertEqual(caught.exception.exceptions, (third_error, first_error))
        self.assertIn(str(self.path), first_error.__notes__[0])
        session.__exit__(None, None, None)
        for wb in (first, second, third):
            wb.close.assert_called_once()

    def test_cleanup_preserves_body_exception_and_reports_close_failures(self) -> None:
        for error in (ValueError("body failed"), KeyboardInterrupt()):
            with self.subTest(exception=type(error).__name__):
                first, second = Mock(spec=PyxlWorkbook), Mock(spec=PyxlWorkbook)
                first.close.side_effect = OSError("close failed")
                with patch(LOADER, side_effect=[first, second]), patch(
                    "noxl.engines.pyxl.session.mlog.warning"
                ) as warning:
                    with self.assertRaises(type(error)) as caught:
                        with PyxlSession() as session:
                            session.open_workbook(self.path)
                            session.open_workbook(self.path)
                            raise error
                self.assertIs(caught.exception, error)
                self.assertIn("close failed", error.__notes__[0])
                self.assertIn(str(self.path), error.__notes__[0])
                warning.assert_called_once()
                first.close.assert_called_once()
                second.close.assert_called_once()

    def test_writable_session_never_saves_implicitly(self) -> None:
        before = self.path.read_bytes()
        with PyxlSession() as session:
            wb = session.open_workbook(self.path)
            self.addCleanup(wb.close)
            self.assertFalse(wb.read_only)
            self.assertTrue(wb.data_only)
            wb["Data"]["B2"] = 999
            saved = self.enterContext(patch.object(wb, "save"))
        saved.assert_not_called()
        self.assertEqual(self.path.read_bytes(), before)


class HelperOwnershipTests(WorkbookFixture):
    def test_helpers_only_close_owned_acquisitions_on_success_and_failure(self) -> None:
        for name, reader, expected, failure_target, read_only in READERS:
            for owned in (False, True):
                for fail in (False, True):
                    with self.subTest(helper=name, owned=owned, failure=fail):
                        wb = self.borrowed()
                        result: object = None
                        with patch(LOADER, return_value=wb) as load, patch.object(
                            wb, "close", wraps=wb.close
                        ) as closed:
                            failure = (
                                patch(failure_target, side_effect=ValueError("injected"))
                                if fail else nullcontext()
                            )
                            with failure:
                                source = self.path if owned else wb
                                if fail:
                                    with self.assertRaises((ValueError, RuntimeError)):
                                        reader(source)
                                else:
                                    result = reader(source)
                            if owned:
                                load.assert_called_once_with(
                                    self.path, read_only=read_only, data_only=True
                                )
                                closed.assert_called_once()
                                self.assert_closed(wb)
                            else:
                                load.assert_not_called()
                                closed.assert_not_called()
                                self.assertEqual(list(wb["Data"].values), EXPECTED_ROWS)
                                self.assertTrue(wb.read_only)
                                self.assertFalse(wb.data_only)
                            if not fail:
                                self.assertEqual(result, expected)

    def test_worksheet_helpers_borrow_parent_on_success_and_failure(self) -> None:
        wb = self.borrowed()
        with patch.object(wb, "close", wraps=wb.close) as closed:
            header = pyxl.find_header(wb["Data"], required_headers=["key", "value"])
            self.assertEqual(header.columns, {"key": 1, "value": 2})
            self.assertEqual(
                pyxl.read_rows(wb["Data"], origin="B2", max_cols=1),
                [(10,), (20,)]
            )
            with self.assertRaises(RuntimeError):
                pyxl.find_header(wb["Data"], required_headers=["missing"])
            with self.assertRaises(RuntimeError):
                pyxl.read_rows(wb["Data"], origin="not a cell")
            closed.assert_not_called()
            self.assertEqual(list(wb["Data"].values), EXPECTED_ROWS)

    def test_cell_results_require_an_explicit_owner(self) -> None:
        with patch(LOADER) as load:
            for source in (self.path, str(self.path)):
                with self.subTest(source=type(source).__name__):
                    self.assertRaisesRegex(
                        ValueError, "PyxlSession", pyxl.read_rows,
                        source, sheetname="Data", values_only=False
                    )
            load.assert_not_called()
        wb = self.borrowed()
        for cells in (
            pyxl.read_rows(wb, sheetname="Data", values_only=False),
            pyxl.read_rows(wb["Data"], values_only=False),
        ):
            self.assertIs(cells[0][0].parent, wb["Data"])
            self.assertEqual(cells[1][1].value, 10)
        self.assertEqual(list(wb["Data"].values), EXPECTED_ROWS)

    def test_mapped_data_closes_its_path_and_preserves_results(self) -> None:
        for fail in (False, True):
            with self.subTest(failure=fail):
                wb = self.borrowed()
                result: MappedData | None = None
                options = MappedDataSearchOptions(
                    self.path, "Data", ["missing"] if fail else ["key"], ["value"]
                )
                with patch(LOADER, return_value=wb) as load, patch.object(
                    wb, "close", wraps=wb.close
                ) as closed:
                    if fail:
                        with self.assertRaises(RuntimeError):
                            pyxl.read_mapped_data(options, data_only=False)
                    else:
                        result = pyxl.read_mapped_data(options, data_only=False)
                    load.assert_called_once_with(
                        self.path, read_only=True, data_only=False
                    )
                    closed.assert_called_once()
                    self.assert_closed(wb)
                if not fail:
                    assert result is not None
                    self.assertEqual(result.rows, [["A", 10, 2], ["B", 20, 3]])
                    self.assertEqual(result.columns, ["key", "value", "__row_index__"])
                    self.assertEqual(result.header.headers, ("key", "value"))

    def test_raw_resolvers_leave_cleanup_with_the_caller(self) -> None:
        for resolve in (pyxl.resolve_pyxl_wb, pyxl.workbook.resolve):
            with self.subTest(resolver=resolve.__module__):
                wb = resolve(self.path, read_only=True, data_only=False)
                self.addCleanup(wb.close)
                with patch.object(wb, "close", wraps=wb.close) as closed:
                    self.assertIs(resolve(wb), wb)
                    closed.assert_not_called()
                    self.assertEqual(list(wb["Data"].values), EXPECTED_ROWS)
                    self.assertTrue(wb.read_only)
                    self.assertFalse(wb.data_only)
