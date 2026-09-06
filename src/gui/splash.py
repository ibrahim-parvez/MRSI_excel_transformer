from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QLabel, QProgressBar, QFrame,
                             QMessageBox, QDialog, QApplication)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont, QPalette

from utils.resources import logo_pixmap
from utils.updater import (AutoUpdater, CURRENT_VERSION, UpdateAvailableDialog,
                           apply_update_and_restart)
        
class StartupSplashScreen(QWidget):
    # Signal emitted when the update check is done
    startup_ready = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setFixedSize(500, 350)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        # --- AUTO DETECT SYSTEM THEME ---
        app = QApplication.instance()
        is_dark_mode = False
        if app:
            # If the default window background is dark, we are in dark mode
            bg_lightness = app.palette().color(QPalette.ColorRole.Window).lightness()
            is_dark_mode = bg_lightness < 128

        # Define dynamic colors based on the detected theme
        if is_dark_mode:
            frame_style = "background-color: #1E1F22; border-radius: 15px; border: 1px solid #333333;"
            title_color = "#FFFFFF"
            subtitle_color = "#DDDDDD"
            loading_color = "#AAAAAA"
            pbar_style = """
                QProgressBar { background-color: #2A2B2E; border-radius: 4px; border: 1px solid #444444; }
                QProgressBar::chunk { background-color: #2E7D32; border-radius: 4px; }
            """
        else:
            frame_style = "background-color: #ffffff; border-radius: 15px; border: 1px solid #ddd;"
            title_color = "#333333"
            subtitle_color = "#666666"
            loading_color = "#888888"
            pbar_style = """
                QProgressBar { background-color: #f0f0f0; border-radius: 4px; border: none; }
                QProgressBar::chunk { background-color: #7A003C; border-radius: 4px; }
            """

        self.main_frame = QFrame(self)
        self.main_frame.setGeometry(0, 0, 500, 350)
        self.main_frame.setStyleSheet(f"QFrame {{ {frame_style} }}")
        
        layout = QVBoxLayout(self.main_frame)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(40, 40, 40, 40)

        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.logo_label.setStyleSheet("border: none;")

        pixmap = logo_pixmap(120)
        if pixmap.isNull():
            self.logo_label.setText("MRSI")
            self.logo_label.setFont(QFont("Arial", 30, QFont.Weight.Bold))
            self.logo_label.setStyleSheet("color: #7A003C; border: none;")
        else:
            self.logo_label.setPixmap(pixmap)

        layout.addWidget(self.logo_label)
        layout.addSpacing(20)

        # 2. Title & Subtitle
        title = QLabel("Data Normalization Tool")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {title_color}; border: none;")
        layout.addWidget(title)

        subtitle = QLabel("McMaster Research Group for\nStable Isotopologues")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setFont(QFont("Arial", 12))
        subtitle.setStyleSheet(f"color: {subtitle_color}; border: none;")
        layout.addWidget(subtitle)
        
        layout.addSpacing(30)

        # 3. Loading Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet(pbar_style)
        layout.addWidget(self.progress_bar)

        # 4. Loading Text
        self.loading_text = QLabel("Checking for updates...")
        self.loading_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_text.setFont(QFont("Arial", 9))
        self.loading_text.setStyleSheet(f"color: {loading_color}; border: none; margin-top: 5px;")
        layout.addWidget(self.loading_text)

        # --- AUTO UPDATE TRIGGER ---
        self.start_update_check()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def start_update_check(self):
        self.updater_thread = AutoUpdater(mode="check")
        self.updater_thread.check_finished.connect(self.on_check_finished)
        self.updater_thread.error_occurred.connect(self.on_error)
        self.updater_thread.start()

    def on_check_finished(self, has_update, latest_version, download_url):
        if has_update:
            # Open our custom styled dialog
            dialog = UpdateAvailableDialog(CURRENT_VERSION, latest_version, self)
            
            # If they click "Update Now"
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.start_download(download_url)
                return
            # If they click "Update Later" (or close the window)
            else:
                self.proceed_with_startup()
                return
        
        # If no update is found
        self.proceed_with_startup()

    def start_download(self, url):
        self.loading_text.setText("Downloading update...")
        self.progress_bar.setValue(0)
        
        self.downloader_thread = AutoUpdater(mode="download", url=url)
        self.downloader_thread.progress_updated.connect(self.progress_bar.setValue)
        self.downloader_thread.download_finished.connect(self.on_download_finished)
        self.downloader_thread.error_occurred.connect(self.on_error)
        self.downloader_thread.start()

    def on_download_finished(self, download_path):
        if download_path:
            self.loading_text.setText("Installing update and restarting...")
            apply_update_and_restart(download_path)
        else:
            self.proceed_with_startup()

    def on_error(self, error_message):
        print(error_message) 
        self.proceed_with_startup()

    def proceed_with_startup(self):
        # Tell main.py to take over and start loading the heavy modules!
        self.startup_ready.emit()