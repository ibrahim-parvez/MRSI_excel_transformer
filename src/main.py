import sys
import time
import ctypes
import base64 # <-- Import base64
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QImage, QPixmap, QIcon
from gui.splash import StartupSplashScreen

logo_base64 = "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" 

# Define window at the module level so it doesn't get garbage collected
window = None

def main():
    global window
    if sys.platform == 'win32':
        myappid = 'mrsi.dnt.1.0' # Arbitrary unique identifier
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    app = QApplication(sys.argv)

    # ---------------------------------------------------------
    # Set the Windows Taskbar Icon using Base64 (Windows ONLY)
    # ---------------------------------------------------------
    if sys.platform == 'win32':
        try:
            icon_data = base64.b64decode(logo_base64)
            icon_image = QImage.fromData(icon_data)
            icon_pixmap = QPixmap.fromImage(icon_image)
            app.setWindowIcon(QIcon(icon_pixmap))
        except Exception as e:
            print(f"Failed to load taskbar icon: {e}")
    # ---------------------------------------------------------
    
    # 1. Show Splash immediately
    splash = StartupSplashScreen()
    splash.show()
    
    if sys.platform == 'win32':
        try:
            import pyi_splash  # type: ignore
            # Update the text on the bootloader splash just before closing (optional)
            pyi_splash.update_text('UI Loaded ...')
            pyi_splash.close()
        except ImportError:
            # If we are running from a standard Python environment, just pass
            pass
        # ---------------------------------------------------------

    # ---------------- SETTINGS ----------------
    # The splash screen will stay open for at LEAST this many seconds.
    MIN_DURATION = 1.5
    # ------------------------------------------

    def start_heavy_loading():
        """This runs ONLY after the update check finishes."""
        global window

        # Helper function to smooth out the jumps
        def smooth_progress(current_val, target_val, time_allocated, task_start_time):
            time_spent = time.time() - task_start_time
            time_left = time_allocated - time_spent
            steps_to_move = target_val - current_val
            
            if time_left > 0 and steps_to_move > 0:
                delay_per_step = time_left / steps_to_move
                for i in range(current_val + 1, target_val + 1):
                    splash.update_progress(i)
                    app.processEvents()
                    time.sleep(delay_per_step)
            else:
                splash.update_progress(target_val)
                app.processEvents()

        chunk_time = MIN_DURATION / 3.0 
        splash.update_progress(0)
        
        # ==========================================
        # STEP 1: LOAD DATA ENGINES (Pandas)
        # ==========================================
        step_start = time.time()
        splash.loading_text.setText("Loading Data Engines (Pandas)...")
        app.processEvents()
        
        import pandas  # Blocking import
        smooth_progress(0, 33, chunk_time, step_start)

        # ==========================================
        # STEP 2: LOAD GUI MODULES
        # ==========================================
        step_start = time.time()
        splash.loading_text.setText("Loading Interface Modules...")
        app.processEvents()
        
        from gui.main_window import DataToolApp # Blocking import
        smooth_progress(33, 66, chunk_time, step_start)

        # ==========================================
        # STEP 3: CONSTRUCT WINDOW
        # ==========================================
        step_start = time.time()
        splash.loading_text.setText("Constructing User Interface...")
        app.processEvents()
        
        window = DataToolApp() # Init the main window
        smooth_progress(66, 95, chunk_time, step_start)

        # ==========================================
        # FINISH
        # ==========================================
        splash.update_progress(100)
        splash.loading_text.setText("Ready!")
        app.processEvents()
        time.sleep(0.3)

        splash.close()
        window.show()

    # 2. Connect the signal from the splash screen to our loading function
    splash.startup_ready.connect(start_heavy_loading)

    sys.exit(app.exec())

if __name__ == "__main__":
    main()