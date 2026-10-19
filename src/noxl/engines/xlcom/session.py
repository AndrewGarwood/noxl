"""
*`@file`* `engines/xlcom/session`
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import (
    Any, Literal, Optional,
)
from pathlib import Path
import win32process, pythoncom
import psutil, gc
import time
import threading
import win32com.client as win32
from win32com.client import gencache as win32gencache
from lute import (LogMsg, WarningManager, annotate, get_logger)
from pyqi import (fs, Array, safe, export)
from noxl.constants import vba
from noxl.utils.sheetname import get_resolved_sheetname
from .types import (ComWorkbookSource, SheetKey)
from .workbook._sheets import (add_worksheet, resolve_dst_sheet, get_sheetnames)
from noxl.types import (ExcelSaveTimeoutError, CalculationState)

mlog = get_logger(Path(__file__).parent.stem)

@dataclass
class ExcelSession():
    visible: bool = False
    display_alerts: bool = False
    screen_updating: bool = False
    enable_events: bool = False
    auto_recover: bool = False
    kill_on_exit: bool = True
    quit_timeout: float = 5.0

    _excel: Optional[win32.CDispatch] = field(default=None, init=False, repr=False)
    _pid: Optional[int] = field(default=None, init=False)
    _proc: Optional[psutil.Process] = field(default=None, init=False, repr=False)
    _workbooks: dict[str, win32.CDispatch] = field(default_factory=dict, init=False, repr=False)
    """ `{ [path: string]: WorkbookComObject }` """
    _coinitialized: bool = field(default=False, init=False)
    _warning_filters: set[str] = field(default_factory=set, init=False, repr=False)

    @property
    def excel(self) -> win32.CDispatch:
        if self._excel is None:
            raise RuntimeError(f"ExcelSession is not active (self._excel is None)")
        return self._excel

    @property
    def pid(self) -> int:
        if not isinstance(self._pid, int):
            raise RuntimeError(LogMsg.err( 
                error=ValueError(f"self._pid is not an int"),
                message=f"invalid usage of ExcelSession, must use in with statement"
            ))
        return self._pid
    def __enter__(self) -> ExcelSession:       
        try:
            pythoncom.CoInitialize()
            self._coinitialized = True

            self._excel = win32.DispatchEx("Excel.Application")
            # self._excel = win32gencache.EnsureDispatch("Excel.Application")
            if self._excel is None:
                raise RuntimeError(LogMsg(
                    f"win32.DispatchEx('Excel.Application') returned None"
                ))
            self._pid = _get_pid(self._excel)
            if self._pid is not None:
                self._proc = psutil.Process(self._pid)

            self._excel.Visible = self.visible
            self._excel.DisplayAlerts = self.display_alerts
            self._excel.ScreenUpdating = self.screen_updating
            self._excel.EnableEvents = self.enable_events
            self._excel.AutoRecover.Enabled = self.auto_recover
            self._warning_filters = set()
            self._add_warning_filter(r".*wmf image format is not supported.*")
            return self
        except Exception:
            self._cleanup()
            raise

    def _cleanup(self) -> None:
        try:
            if self._excel is not None:
                self._excel.Quit()
        except Exception:
            pass
        finally:
            self._excel = None
            self._workbooks.clear()
            gc.collect()
            time.sleep(0.25)
            if self.kill_on_exit:
                self.terminate_owned_process()
            if self._coinitialized:
                pythoncom.CoUninitialize()
                self._coinitialized = False

    def __exit__(self, exc_type, exc, tb) -> bool:
        try:
            self.close_all_workbooks(save_changes=False)
            time.sleep(1)
            try:
                if self._excel is not None:
                    self._excel.Quit()
            except Exception as e:
                mlog.warning(LogMsg.err(e,
                    f"Failed calling Excel.Application.Quit()"
                ))
        finally:
            self._excel = None
            self._workbooks.clear()
            gc.collect()
            time.sleep(0.25)
            if self.kill_on_exit:
                self.terminate_owned_process()
            if self._coinitialized:
                pythoncom.CoUninitialize()
                self._coinitialized = False
            if (self._warning_filters):
                WarningManager.remove_filter("ignore", *self._warning_filters)
        return False

    def _add_warning_filter(self, *substr: str) -> None:
        WarningManager.ignore(*substr)
        Array.forEach(substr, lambda x: self._warning_filters.add(x))

    def close_all_workbooks(self, /, save_changes: bool = False) -> None:
        for path, wb in reversed(self._workbooks.items()):
            try:
                wb.Close(SaveChanges=save_changes)
            except Exception as e:
                mlog.warning(LogMsg.err(e,
                    f"Failed closing workbook during ExcelSession cleanup.",
                    workbook=str(path)
                ))
        self._workbooks.clear()
        gc.collect()

    def _rm_wb_ref(self, wb: win32.CDispatch) -> None:
        """ `_remove_workbook_reference` """
        key = self.workbook_key(wb)
        if key in self._workbooks:
            del self._workbooks[key]
        
    def close_workbook(
        self,
        wb: win32.CDispatch,
        /,
        save_changes: bool = False,
    ) -> None:
        key: str | None = None
        try: key = self.workbook_key(wb)
        except Exception: key = None
        try:
            wb.Close(SaveChanges=save_changes)
        finally:
            if key is not None:
                self._workbooks.pop(key, None)
            gc.collect()
    
    def save_workbook(
        self,
        wb: win32.CDispatch,
        /,
        timeout: float=20.0,
    ) -> None:
        """
        Save workbook with a watchdog.

        Important:
            This does not magically interrupt wb.Save().
            Instead, if wb.Save() hangs, the timer terminates the owned Excel PID.
            The broken COM call should then return/raise, and this method raises
            ExcelSaveTimeoutError.
        """
        if self.is_readonly_workbook(wb):
            raise RuntimeError(LogMsg("Unable to save workbook because ReadOnly=True"))
        timed_out = threading.Event()
        started = time.perf_counter()
        def _on_timeout() -> None:
            timed_out.set()
            mlog.warning(LogMsg(
                f"wb.Save() exceeded timeout={timeout:.2f}s. "
                f"Terminating owned Excel PID={self._pid}."
            ))
            self.terminate_owned_process()
        timer = threading.Timer(timeout, _on_timeout)
        timer.daemon = True
        try:
            timer.start()
            wb.Save()
        except Exception as e:
            elapsed = time.perf_counter() - started
            if timed_out.is_set():
                raise ExcelSaveTimeoutError(LogMsg.err(e,
                    f"Excel save timed out after {elapsed:.2f}s "
                    f"Excel PID={self._pid} was terminated.",
                    wb=safe.readattr(wb, "FullName")
                )) from e
            raise
        finally:
            timer.cancel()
        elapsed = time.perf_counter() - started
        if timed_out.is_set():
            raise ExcelSaveTimeoutError(LogMsg(
                f"Excel save timed out after {elapsed:.2f}s.",
                wb=safe.readattr(wb, "FullName")
            ))
        if not bool(wb.Saved):
            raise RuntimeError(
                LogMsg(
                    "Workbook remains modified after wb.Save() returned",
                    wb=safe.readattr(wb, "FullName"),
                    elapsed=elapsed,
                )
            )

    def terminate_owned_process(self) -> bool:
        if self._proc is None or not self._proc.is_running():
            print(LogMsg(
                f"No owned Excel process to terminate.",
                pid=self._pid,
            ))
            return False
        return bool(fs.kill_procs([self._proc]))

    def wait_for_idle(self, timeout_seconds: int=900, poll_seconds: float=2.0):
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            try:
                calc_done = self.excel.CalculationState == CalculationState.DONE
                ready = bool(self.excel.Ready)
                if calc_done and ready:
                    return
            except Exception:
                pass
            time.sleep(poll_seconds)
        raise TimeoutError(LogMsg(
            f"Calculations DNF by deadline",
            f"{self.excel.ActiveWorkbook.Name=}",
            timeout_seconds=timeout_seconds,
        ))
    
    def wait_for_async_query_complete(self):
        """ `self.excel.CalculateUntilAsyncQueriesDone()` """
        self.excel.CalculateUntilAsyncQueriesDone()
        self.wait_for_idle()
    
    def open_workbook(
        self, 
        path: fs.StrPath,
        /, 
        update_links: int = 0, 
        read_only: bool = False,
        activate: bool = True
    ) -> win32.CDispatch:
        path = fs.to_path(path)
        wb = self.excel.Workbooks.Open(
            str(path),
            UpdateLinks=update_links,
            ReadOnly=read_only
        )
        if activate:
            wb.Activate()
        key = self.workbook_key(wb)
        self._workbooks[key] = wb
        return wb

    def add_workbook(self, path: fs.StrPath, /) -> win32.CDispatch:
        """
        wb = xl.excel.Workbooks.Add()
        wb.SaveAs(Filename=str(p), FileFormat=51)
        add entry to self._workbooks: dict[str, win32.CDispatch]
        """
        key = self.workbook_key(path)
        wb = self.excel.Workbooks.Add()
        wb.SaveAs(Filename=key, FileFormat=vba.xlOpenXMLWorkbook)
        self._workbooks[key] = wb
        return wb

    def resolve_wb_source(self, src: ComWorkbookSource,/) -> win32.CDispatch:
        """*`@returns`* `self.open_workbook(src) if isinstance(src, (str, Path)) else src` """
        return self.open_workbook(src) if isinstance(src, (str, Path)) else src

    def move_sheet(
        self, 
        src: ComWorkbookSource, 
        src_sheet: SheetKey, 
        dst: ComWorkbookSource,
        *,
        dst_anchor_key: SheetKey | None=None,
        dst_anchor_how: Literal["Before", "After"]="After",
        mk_copy: bool=True,
        result_sheetname: str | None=None
    ):
        """
        ***@not_implemented***

        - *`@param`* `src (ComWorkbookSource)` - source workbook
        - *`@param`* `src_sheet (SheetKey)` - identifier to get source sheet within `src`
        
        - for now, requires that source sheetname is not in destination workbook's sheetnames
        """
        try:
            src = self.resolve_wb_source(src)
            dst = self.resolve_wb_source(dst)
            if (self.workbook_key(src) not in self._workbooks 
                or self.workbook_key(dst) not in self._workbooks):
                raise RuntimeError(LogMsg(
                    f"Source and Destination workbooks must both be opened in same Excel instance"
                ))
            dst_sheetnames = get_sheetnames(dst)
            if (isinstance(dst_anchor_key, str) and dst_anchor_key not in dst_sheetnames):
                raise ValueError(LogMsg(
                    f"{dst_anchor_key=} is not an existing sheet in destination workbook",
                    dst_sheetnames=dst_sheetnames
                ))

            src_ws = src.Worksheets(src_sheet)
            src_sheetname = str(src_ws.Name)
            if (src_sheetname in dst_sheetnames):
                raise ValueError(LogMsg(
                    f"Destination Workbook already contains source Worksheet.Name",
                    src_sheetname=src_sheetname,
                    dst_sheetnames=str(dst_sheetnames)
                ))
            if (dst_anchor_key is None):
                dst_anchor_key = int(dst.Worksheets.Count)
                dst_anchor_how = "After"
            anchor = dst.Worksheets(dst_anchor_key)
            anchor_index = int(anchor.Index)
            # var definitions to use as keyword arguments.
            if dst_anchor_how == "Before":
                before = anchor
                after = None
                result_index = anchor_index
            else:
                before = None
                after = anchor
                result_index = anchor_index + 1

            dst_count_before = int(dst.Worksheets.Count)
            src_count_before = int(src.Worksheets.Count)
            workbook_count_before = int(self.excel.Workbooks.Count)

            src.Activate()
            src_ws = src.Worksheets(src_sheet)
            # Keep these as directly written keyword arguments.
            # Using a dynamically expanded kwargs mapping caused Excel to behave
            # as though Before and After were omitted.
            if mk_copy:
                src_ws.Copy(Before=before, After=after)
            else:
                src_ws.Move(Before=before, After=after)

            dst_count_after = int(dst.Worksheets.Count)
            src_count_after = int(src.Worksheets.Count)
            workbook_count_after = int(self.excel.Workbooks.Count)

            if workbook_count_after != workbook_count_before:
                raise RuntimeError(
                    LogMsg(
                        "Worksheet operation unexpectedly created a new workbook",
                        operation="Copy" if mk_copy else "Move",
                        workbook_count_before=workbook_count_before,
                        workbook_count_after=workbook_count_after,
                    )
                )

            if dst_count_after != dst_count_before + 1:
                raise RuntimeError(
                    LogMsg(
                        "Worksheet operation did not add a worksheet to destination",
                        operation="Copy" if mk_copy else "Move",
                        dst_count_before=dst_count_before,
                        dst_count_after=dst_count_after,
                    )
                )

            expected_src_count_after = (
                src_count_before
                if mk_copy
                else src_count_before - 1
            )

            if src_count_after != expected_src_count_after:
                raise RuntimeError(
                    LogMsg(
                        "Source worksheet count changed unexpectedly",
                        operation="Copy" if mk_copy else "Move",
                        src_count_before=src_count_before,
                        src_count_after=src_count_after,
                        expected_src_count_after=expected_src_count_after,
                    )
                )

            if not 1 <= result_index <= dst_count_after:
                raise RuntimeError(
                    LogMsg(
                        "Expected result index is outside destination worksheet bounds",
                        result_index=result_index,
                        dst_count=dst_count_after,
                    )
                )

            result_ws = dst.Worksheets(result_index)

            result_parent = str(result_ws.Parent.FullName)
            expected_parent = str(dst.FullName)

            if result_parent.casefold() != expected_parent.casefold():
                raise RuntimeError(
                    LogMsg(
                        "Resolved result worksheet belongs to an unexpected workbook",
                        result_sheet=safe.readattr(result_ws, "Name"),
                        result_parent=result_parent,
                        expected_parent=expected_parent,
                    )
                )
            if result_sheetname is not None:
                result_ws.Name = result_sheetname
            dst.Save()
            if not bool(dst.Saved):
                raise RuntimeError(
                    LogMsg(
                        "Destination workbook remains modified after Save returned",
                        dst=safe.readattr(dst, "FullName"),
                    )
                )
            return result_ws
        except Exception as e:
            raise RuntimeError(LogMsg.err(e, 
                "Failed to move sheet",
                src_wb=safe.readattr(src, "FullName"),
                src_sheet=src_sheet,
                dst=annotate(dst),
                mk_copy=mk_copy
            )) from e
    
    @staticmethod
    def workbook_key(wb: ComWorkbookSource, /) -> str:
        if (isinstance(wb, (str, Path))):
            return fs.normalized(str(wb))
        return fs.normalized(str(wb.FullName))
    
    def is_readonly_workbook(self, wb: win32.CDispatch, /) -> bool:
        return bool(wb.ReadOnly)
    # maybe see if can refactor to 
    # get_addin(addin_type: Literal["COMAddIns",...], name: str) -> self.excel[addin_type](name)
    def get_com_addin(self, name: str, /) -> win32.CDispatch:
        """
        *`@returns`* `self.excel.COMAddIns(name)`
        - `= Excel.Application.COMAddIns(name)`
        """
        try:
            addin = self.excel.COMAddIns(name)
            if addin is None:
                raise KeyError(LogMsg(
                    f"Application.COMAddIns({name}) returned None"
                ))
            return addin
        except Exception as e:
            raise Exception(LogMsg.err(e,
                f"Failed to get COM add in from excel instance",
            ))

    def get_addins(self) -> list[win32.CDispatch]:
        count = int(self.excel.COMAddIns.Count)
        addins: list[win32.CDispatch] = []
        for i in range(0, count+1):
            addins.append(self.excel.COMAddIns.Item(i))
        return addins

    def get_addin_snapshots(
        self, 
        attrs: list[str]=["description", "prog_id", "guid", "connect"], 
        /
    ) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        try:
            for addin in self.get_addins():
                rows.append({
                    k: safe.readattr(addin, k)
                    for k in attrs
                })
        except Exception as e:
            rows.append({"error": repr(e)})
        return rows
    
def _get_pid(excel: win32.CDispatch) -> Optional[int]:
    """
    Return the process id for an Excel Application COM object.

    Uses Application.Hwnd, then maps that HWND to a process id.
    """
    try:
        hwnd = excel.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid)
    except Exception as e:
        mlog.warning(LogMsg.err(e, f"Could not determine Excel PID"))
        return None


__all__ = export(ExcelSession)