"""Entry point for the MRSI Data Normalization Tool.

Shows the splash screen first, then imports the heavy modules (pandas, the
main window) while the progress bar advances, so the app appears instantly
instead of after a multi-second cold import.
"""

import ctypes
import sys
import time

from PyQt6.QtWidgets import QApplication

from gui.splash import StartupSplashScreen
from utils.resources import app_icon

# Kept at module level so the window is not garbage collected once main() returns.
window = None

# The splash screen stays up for at least this many seconds.
MIN_DURATION = 1.5

# Windows groups taskbar buttons by this id; without it the app inherits
# Python's own icon instead of ours.
APP_USER_MODEL_ID = "mrsi.dnt.1.0"


def _configure_windows_taskbar() -> None:
    if sys.platform == "win32":
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)


def _close_bootloader_splash() -> None:
    """Dismiss the PyInstaller splash image used by the Windows one-file build.

    Only that build defines a splash, but ``pyi_splash`` is importable from
    any frozen build — elsewhere it complains on import and then raises on
    use, because the bootloader never initialised it. Hence the platform
    check, and the catch-all behind it: a decoration is never worth failing
    startup over.
    """
    if sys.platform != "win32":
        return

    try:
        import pyi_splash  # type: ignore

        pyi_splash.update_text("UI Loaded ...")
        pyi_splash.close()
    except Exception:
        pass


def _load_application(app: QApplication, splash: StartupSplashScreen) -> None:
    """Import and construct the main window, driving the splash progress bar.

    Runs only once the splash screen's update check has finished.
    """
    global window

    def smooth_progress(current, target, time_allocated, task_start_time):
        """Ease the bar from current to target over whatever time is left."""
        time_left = time_allocated - (time.time() - task_start_time)
        steps_to_move = target - current

        if time_left <= 0 or steps_to_move <= 0:
            splash.update_progress(target)
            app.processEvents()
            return

        delay_per_step = time_left / steps_to_move
        for value in range(current + 1, target + 1):
            splash.update_progress(value)
            app.processEvents()
            time.sleep(delay_per_step)

    chunk_time = MIN_DURATION / 3.0
    splash.update_progress(0)

    # Each blocking import runs while its own message is on screen.
    step_start = time.time()
    splash.loading_text.setText("Loading Data Engines (Pandas)...")
    app.processEvents()
    import pandas  # noqa: F401
    smooth_progress(0, 33, chunk_time, step_start)

    step_start = time.time()
    splash.loading_text.setText("Loading Interface Modules...")
    app.processEvents()
    from gui.main_window import DataToolApp
    smooth_progress(33, 66, chunk_time, step_start)

    step_start = time.time()
    splash.loading_text.setText("Constructing User Interface...")
    app.processEvents()
    window = DataToolApp()
    smooth_progress(66, 95, chunk_time, step_start)

    splash.update_progress(100)
    splash.loading_text.setText("Ready!")
    app.processEvents()
    time.sleep(0.3)

    splash.close()
    window.show()


def main() -> None:
    _configure_windows_taskbar()

    app = QApplication(sys.argv)
    app.setWindowIcon(app_icon())

    splash = StartupSplashScreen()
    splash.show()
    _close_bootloader_splash()

    splash.startup_ready.connect(lambda: _load_application(app, splash))

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
