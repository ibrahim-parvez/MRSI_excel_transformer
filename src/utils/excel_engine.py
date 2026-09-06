"""
Shared background Excel engine.

The Combine processors keep one hidden Excel instance alive for a whole run and
reuse it to recalculate every workbook between steps.

On Windows, xlwings' App.quit() does more than quit that one instance: it also
runs Apps.cleanup(), which lists every EXCEL.exe with tasklist, subtracts the
PIDs it can find by walking top-level windows for an XLDESK/EXCEL7 child, and
runs "taskkill /F" on the rest as zombies. An instance holding no workbook has
no EXCEL7 window, so it cannot be found and gets killed. That is exactly the
state a shared instance sits in between refreshes, so any other short-lived
xw.App() quitting anywhere in the process would kill it, and the next call on
the dead COM object fails with:

    (-2147023174, 'The RPC server is unavailable.', None, None)

macOS is unaffected: its quit() is a plain AppleScript quit with no sweep.

Three things keep that from happening here:

    1. One shared Excel instance per run, so the step processors reuse it
       instead of starting and quitting their own.
    2. A keep-alive workbook (Windows only) so the instance always has an
       EXCEL7 window and is never mistaken for a zombie.
    3. disable_xlwings_zombie_cleanup(), plus an automatic restart and retry so
       a dead connection recovers instead of aborting the run.
"""

import os
import sys
import threading
import time

# Imported at module level (not lazily) so PyInstaller always bundles xlwings.
try:
    import xlwings as xw
except Exception:  # pragma: no cover - Excel/xlwings not installed
    xw = None


# HRESULTs that all mean "the Excel process behind this COM object is gone".
_DEAD_COM_HRESULTS = {
    -2147023174,  # 0x800706BA RPC_S_SERVER_UNAVAILABLE  <- the reported error
    -2147023170,  # 0x800706BE RPC_S_CALL_FAILED
    -2147417848,  # 0x80010108 RPC_E_DISCONNECTED
    -2147417851,  # 0x80010105 RPC_E_SERVERFAULT
    -2146959355,  # 0x80080005 CO_E_SERVER_EXEC_FAILURE
    -2147221021,  # 0x800401FD CO_E_OBJNOTCONNECTED
}


def is_dead_com_error(exc):
    """True if `exc` means the Excel COM server died or disconnected."""
    for arg in getattr(exc, "args", ()):
        if isinstance(arg, int) and arg in _DEAD_COM_HRESULTS:
            return True
    text = str(exc)
    return "RPC server is unavailable" in text or "RPC_E_DISCONNECTED" in text


_cleanup_patched = False


def disable_xlwings_zombie_cleanup():
    """
    Stop xlwings from force-killing Excel processes it cannot enumerate.

    Windows only; a no-op elsewhere and safe to call repeatedly. Without this,
    any ``App.quit()`` anywhere in the process can taskkill our background
    instance (and any hidden instance the user owns, unsaved work included).
    """
    global _cleanup_patched
    if _cleanup_patched or not sys.platform.startswith("win"):
        return
    try:
        from xlwings import _xlwindows
        _xlwindows.Apps.cleanup = staticmethod(lambda: None)
        _cleanup_patched = True
    except Exception:
        pass


class ExcelEngine:
    """One hidden Excel instance, restarted transparently if it ever dies."""

    def __init__(self, log=None, keep_alive=False):
        self._app = None
        self._keep_alive_book = None
        self._keep_alive = keep_alive
        self._lock = threading.RLock()
        self._log = log or (lambda msg: None)

    # ------------------------------------------------------------------
    # lifecycle
    # ------------------------------------------------------------------
    def _start(self):
        if xw is None:
            raise RuntimeError("xlwings is not available - cannot start Excel.")

        disable_xlwings_zombie_cleanup()

        # COM must be initialised on the calling thread. xlwings does this
        # itself, but a worker thread that never talked to COM before is a
        # common source of trouble, so make it explicit and harmless.
        try:
            import pythoncom  # pywin32, Windows only
            pythoncom.CoInitialize()
        except Exception:
            pass

        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        self._app = app

        # Excel opens a blank workbook of its own when it launches - on macOS
        # it does this even with add_book=False - which shows up as a "Book1"
        # window flashing open before every file. Close it so the only window
        # that ever appears is the file being processed.
        self._close_scratch_books(app)

        # The keep-alive workbook only earns its place for a long-lived
        # instance on Windows: an instance holding zero workbooks is invisible
        # to xlwings' process scan, which is what makes it a taskkill target.
        # It is pure noise everywhere else, so it is opt-in and Windows-only.
        if self._keep_alive and sys.platform.startswith("win"):
            try:
                self._keep_alive_book = app.books.add()
            except Exception:
                self._keep_alive_book = None
        return app

    @staticmethod
    def _close_scratch_books(app):
        """
        Close blank, never-saved workbooks in our own Excel instance.

        A saved workbook's `fullname` carries a directory; a scratch one is
        just "Book1". Safe against the user's own open files: xlwings gives us
        a private instance (newinstance=True on macOS, DispatchEx on Windows),
        so `app.books` never lists workbooks opened outside this process.
        """
        for book in list(app.books):
            try:
                if not os.path.dirname(str(book.fullname or "")):
                    book.close()
            except Exception:
                pass

    def _is_alive(self):
        if self._app is None:
            return False
        try:
            self._app.pid  # cheapest possible round-trip to the COM server
            return True
        except Exception:
            return False

    def _forget(self):
        self._app = None
        self._keep_alive_book = None

    @property
    def app(self):
        """The live Excel instance, started (or restarted) on demand."""
        with self._lock:
            if not self._is_alive():
                if self._app is not None:
                    self._log("Background Excel instance was closed - restarting it.")
                self._forget()
                self._start()
            return self._app

    def restart(self):
        """
        Drop the current instance (dead or alive). The next use of `.app`
        starts a fresh one -- deliberately lazy, so that a failure to launch
        Excel is reported by the retry loop rather than from inside it.
        """
        with self._lock:
            app = self._app
            self._forget()
            if app is not None:
                try:
                    app.quit()
                except Exception:
                    pass

    def shutdown(self):
        """Close every workbook and quit Excel. Never raises."""
        with self._lock:
            app = self._app
            self._forget()
            if app is None:
                return
            try:
                for book in list(app.books):
                    try:
                        book.close()
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                app.quit()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # work
    # ------------------------------------------------------------------
    def refresh(self, file_path, full=False, settle=1.0, attempts=3):
        """
        Open `file_path`, recalculate, save and close it.

        Retries on a dead COM connection by restarting Excel, so an instance
        killed by something else no longer aborts the whole run. Any other
        error (bad path, corrupt workbook, ...) is raised immediately.
        """
        path = os.path.abspath(file_path)

        for attempt in range(1, attempts + 1):
            book = None
            try:
                app = self.app
                book = app.books.open(path)
                if full:
                    try:
                        app.api.CalculateFull()
                    except Exception:
                        app.calculate()
                else:
                    app.calculate()
                if settle:
                    time.sleep(settle)
                book.save()
                return True
            except Exception as e:
                if not is_dead_com_error(e) or attempt == attempts:
                    raise
                self._log(
                    f"Lost the connection to Excel ({e}); "
                    f"restarting it and retrying ({attempt}/{attempts - 1})..."
                )
                try:
                    self.restart()
                except Exception:
                    pass
                time.sleep(2.0)
            finally:
                if book is not None:
                    try:
                        book.close()
                    except Exception:
                        pass
        return False

    def save_as(self, src_path, dest_path):
        """Open `src_path` in Excel and save it as `dest_path`."""
        book = self.app.books.open(os.path.abspath(src_path))
        try:
            book.save(os.path.abspath(dest_path))
        finally:
            try:
                book.close()
            except Exception:
                pass
        return dest_path


# ----------------------------------------------------------------------
# Process-wide shared engine
#
# A Combine run publishes its engine here so the step processors it calls
# reuse that instance instead of starting (and quitting) their own.
# ----------------------------------------------------------------------
_shared_engine = None
_shared_lock = threading.RLock()


def set_shared_engine(engine):
    global _shared_engine
    with _shared_lock:
        _shared_engine = engine


def get_shared_engine():
    with _shared_lock:
        return _shared_engine


def clear_shared_engine(engine=None):
    """Unpublish the shared engine (only if it is still `engine`, if given)."""
    global _shared_engine
    with _shared_lock:
        if engine is None or _shared_engine is engine:
            _shared_engine = None


def recalculate_workbook(file_path, full=False, settle=1.0):
    """
    Recalculate and save `file_path` in Excel. Returns True on success.

    Uses the shared engine when a Combine run owns one -- and never quits it.
    Otherwise starts a private instance and shuts it down afterwards.
    """
    engine = get_shared_engine()
    if engine is not None:
        try:
            return engine.refresh(file_path, full=full, settle=settle)
        except Exception as e:
            print(f"Excel recalculation failed: {e}")
            return False

    engine = ExcelEngine()
    try:
        return engine.refresh(file_path, full=full, settle=settle)
    except Exception as e:
        print(f"Excel recalculation failed: {e}")
        return False
    finally:
        engine.shutdown()


def save_workbook_as(src_path, dest_path):
    """Convert/save a workbook through Excel (e.g. .xls -> .xlsx)."""
    engine = get_shared_engine()
    private = engine is None
    if private:
        engine = ExcelEngine()
    try:
        return engine.save_as(src_path, dest_path)
    finally:
        if private:
            engine.shutdown()
