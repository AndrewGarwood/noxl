"""
*`@file`* `engines/xlcom/session`
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import (
    Any, Optional,
)
from pathlib import Path
import win32process, pythoncom
import psutil, gc
import time
import threading
import win32com.client as win32
from lute import (LogMsg, WarningManager, annotate, get_logger)
from pyqi import (fs, Array)
from noxl.constants import vba
from .types import (ComWorkbookSource, SheetKey)
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
        key = str(wb.FullName)
        if key in self._workbooks:
            del self._workbooks[key]
        
    def close_workbook(
        self,
        wb: win32.CDispatch,
        /,
        save_changes: bool = False,
    ) -> None:
        key: str | None = None
        try: key = str(wb.FullName)
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
                    wb=str(getattr(wb, "FullName", "<name unretrievable>"))
                )) from e
            raise
        finally:
            timer.cancel()
        elapsed = time.perf_counter() - started
        if timed_out.is_set():
            raise ExcelSaveTimeoutError(LogMsg(
                f"Excel save timed out after {elapsed:.2f}s."
            ))

    def terminate_owned_process(self) -> bool:
        if self._proc is None or not self._proc.is_running():
            print(LogMsg(
                f"No owned Excel process to terminate.",
                pid=self._pid,
            ))
            return False
        return bool(fs.kill_procs([self._proc]))

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
                    k: _safe_str(getattr(addin, k, "<unknown>")) 
                    for k in attrs
                })
        except Exception as e:
            rows.append({"error": repr(e)})
        return rows
    
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
        key = wb.FullName
        if not isinstance(key, str):
            raise RuntimeError(LogMsg(
                f"wb.FullName is not a string.",
                actual_type=f"{type(wb.FullName)=}"
            ))
        self._workbooks[key] = wb
        return wb

    def add_workbook(self, path: fs.StrPath, /) -> win32.CDispatch:
        """
        wb = xl.excel.Workbooks.Add()
        wb.SaveAs(Filename=str(p), FileFormat=51)
        add entry to self._workbooks: dict[str, win32.CDispatch]
        """
        
        key = str(path)
        wb = self.excel.Workbooks.Add()
        wb.SaveAs(Filename=key, FileFormat=vba.xlOpenXMLWorkbook)
        self._workbooks[key] = wb
        return wb

    def resolve_wb_source(self, src: ComWorkbookSource,/) -> win32.CDispatch:
        """*`@returns`* `self.open_workbook(src) if isinstance(src, (str, Path)) else src` """
        return self.open_workbook(src) if isinstance(src, (str, Path)) else src

    def move_sheet(
        self, 
        src: win32.CDispatch, 
        src_sheet: SheetKey, 
        dst: win32.CDispatch | fs.StrPath, 
        after_dst_key: Optional[SheetKey]=None,
        dst_sheetname: Optional[str]=None,
        mk_copy: bool=True
    ) -> win32.CDispatch: 
        try:
            if (isinstance(dst, (str, Path))):
                dst = self.open_workbook(dst)
            if (after_dst_key is None):
                after_dst_key = len(tuple(dst.Sheets)) + 1
            dst_sheet: win32.CDispatch = dst.Sheets(after_dst_key)
            if (mk_copy):
                src.Sheets(src_sheet).Copy(After=dst_sheet)
            else: # move
                src.Sheets(src_sheet).Move(After=dst_sheet)
            if (dst_sheetname):
                dst_sheet.Name = dst_sheetname
            return dst_sheet
        except Exception as e:
            raise RuntimeError(LogMsg.err(e, 
                "Failed to move sheet",
                src_wb=getattr(src, "FullName", "<name unretrievable>"),
                src_sheet=src_sheet,
                dst=annotate(dst),
                after_dst_key=after_dst_key,
                dst_sheetname=dst_sheetname,
                mk_copy=mk_copy
            ))

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


def _safe_str(value: Any, /) -> str:
    try: return str(value)
    except Exception: return "<unprintable>"
