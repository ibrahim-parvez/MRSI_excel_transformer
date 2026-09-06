"""Standalone downloader/installer for the MRSI Data Normalization Tool.

Built and shipped separately from the app itself; see packaging/installer.spec.
"""

import sys
import os
import json
import ssl
import certifi
import urllib.request
import urllib.error
import zipfile
import subprocess
import shutil
import time
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QGroupBox, QMessageBox, QProgressBar, QLineEdit,
    QTextEdit, QCheckBox, QComboBox, QFrame
)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QUrl, QTimer,
    QRect, QPropertyAnimation, QEasingCurve, QParallelAnimationGroup
)
from PyQt6.QtGui import QFont, QDesktopServices

# Running this file directly puts its own folder on sys.path, not src/,
# so make the shared utils package importable either way.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from utils.resources import app_icon, logo_pixmap

# ==========================================
# CONFIGURATION - CHANGE THESE!
# ==========================================
GITHUB_OWNER = "ibrahim-parvez"
GITHUB_REPO = "MRSI_Data_Normalization_Tool"

# ==========================================
# WORKER THREADS
# ==========================================
class FetchReleasesThread(QThread):
    releases_fetched = pyqtSignal(list)
    error_occurred = pyqtSignal(str)

    def run(self):
        api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases"
        req = urllib.request.Request(api_url, headers={'User-Agent': 'MRSI-Installer'})
        ssl_context = ssl.create_default_context(cafile=certifi.where())
        
        try:
            with urllib.request.urlopen(req, context=ssl_context) as response:
                data = json.loads(response.read().decode('utf-8'))
                valid_releases = [r for r in data if r.get('assets')]
                
                if not valid_releases and data:
                    self.error_occurred.emit("Found releases, but no files are attached to them.")
                else:
                    self.releases_fetched.emit(valid_releases)
                    
        except urllib.error.HTTPError as e:
            if e.code == 404:
                self.error_occurred.emit("Error 404: Repository not found. Make sure the repo is Public.")
            elif e.code == 403:
                self.error_occurred.emit("Error 403: GitHub API rate limit exceeded.")
            else:
                self.error_occurred.emit(f"GitHub Error {e.code}: {e.reason}")
        except Exception as e:
            self.error_occurred.emit(f"Network failed:\n{str(e)}")


class DownloadThread(QThread):
    progress = pyqtSignal(int)
    status_update = pyqtSignal(str)
    stats_update = pyqtSignal(str) 
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, download_url, save_path):
        super().__init__()
        self.download_url = download_url
        self.save_path = save_path
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def format_size(self, bytes_size):
        if bytes_size < 1024 * 1024:
            return f"{bytes_size / 1024:.1f} KB"
        return f"{bytes_size / (1024 * 1024):.1f} MB"

    def run(self):
        req = urllib.request.Request(self.download_url, headers={'User-Agent': 'MRSI-Installer'})
        ssl_context = ssl.create_default_context(cafile=certifi.where())
        
        try:
            with urllib.request.urlopen(req, context=ssl_context) as response:
                total_size = int(response.getheader('Content-Length', 0).strip())
                downloaded = 0
                chunk_size = 8192
                
                start_time = time.time()
                last_update_time = start_time
                last_downloaded = 0

                with open(self.save_path, 'wb') as file:
                    while not self._is_cancelled:
                        chunk = response.read(chunk_size)
                        if not chunk:
                            break
                        
                        file.write(chunk)
                        downloaded += len(chunk)
                        
                        current_time = time.time()
                        time_diff = current_time - last_update_time
                        
                        if time_diff >= 0.2 or downloaded == total_size:
                            speed_bps = (downloaded - last_downloaded) / time_diff if time_diff > 0 else 0
                            speed_str = f"{self.format_size(speed_bps)}/s"
                            
                            total_str = self.format_size(total_size)
                            dl_str = self.format_size(downloaded)
                            
                            self.stats_update.emit(f"{dl_str} / {total_str}  ({speed_str})")
                            
                            if total_size > 0:
                                percent = int((downloaded / total_size) * 100)
                                self.progress.emit(percent)
                                
                            last_update_time = current_time
                            last_downloaded = downloaded

            if self._is_cancelled:
                if os.path.exists(self.save_path):
                    os.remove(self.save_path)
                self.error.emit("Download cancelled by user.")
                return

            final_path = self.save_path

            # --- UNIVERSAL AUTO-EXTRACTION & CLEANUP ---
            if self.save_path.lower().endswith('.zip'):
                self.status_update.emit("Extracting and cleaning files...")
                self.stats_update.emit("") 
                
                target_install_dir = os.path.dirname(self.save_path)
                base_name = os.path.splitext(os.path.basename(self.save_path))[0]
                
                temp_extract_dir = os.path.join(target_install_dir, base_name + "_temp_extract")
                os.makedirs(temp_extract_dir, exist_ok=True)
                
                if sys.platform == "win32":
                    with zipfile.ZipFile(self.save_path, 'r') as zip_ref:
                        zip_ref.extractall(temp_extract_dir)
                else:
                    subprocess.run(["unzip", "-q", "-o", self.save_path, "-d", temp_extract_dir], check=True)
                    macosx_dir = os.path.join(temp_extract_dir, "__MACOSX")
                    if os.path.exists(macosx_dir):
                        shutil.rmtree(macosx_dir)

                os.remove(self.save_path)

                found_app = None
                
                # Windows Cleanup
                if sys.platform == "win32":
                    extracted_items = os.listdir(temp_extract_dir)
                    if len(extracted_items) == 1 and os.path.isdir(os.path.join(temp_extract_dir, extracted_items[0])):
                        app_folder_path = os.path.join(temp_extract_dir, extracted_items[0])
                    else:
                        app_folder_path = temp_extract_dir
                        
                    for root, dirs, files in os.walk(app_folder_path):
                        for f in files:
                            if f.endswith('.exe') and "MRSI" in f.upper():
                                old_exe_path = os.path.join(root, f)
                                name_no_ext, ext = os.path.splitext(f)
                                clean_exe_name = name_no_ext.replace('.', ' ').replace('_', ' ') + ext
                                new_exe_path = os.path.join(root, clean_exe_name)
                                if old_exe_path != new_exe_path:
                                    os.rename(old_exe_path, new_exe_path)
                                found_app = new_exe_path
                                break
                        if found_app: break
                    
                    clean_folder_name = "MRSI Data Normalization Tool"
                    final_folder_path = os.path.join(target_install_dir, clean_folder_name)
                    
                    if os.path.exists(final_folder_path):
                        shutil.rmtree(final_folder_path)
                        
                    shutil.move(app_folder_path, final_folder_path)
                    
                    if os.path.exists(temp_extract_dir):
                        shutil.rmtree(temp_extract_dir)
                        
                    if found_app:
                        rel_path = os.path.relpath(found_app, app_folder_path)
                        final_path = os.path.join(final_folder_path, rel_path)
                    else:
                        final_path = final_folder_path

                # macOS Cleanup
                else:
                    for root, dirs, files in os.walk(temp_extract_dir):
                        for d in dirs:
                            if d.endswith('.app') and "MRSI" in d.upper():
                                found_app = os.path.join(root, d)
                                subprocess.run(["xattr", "-cr", found_app], check=False)
                                break
                        if found_app:
                            break
                            
                    if found_app:
                        clean_name = "MRSI Data Normalization Tool.app"
                        final_app_path = os.path.join(target_install_dir, clean_name)
                        
                        if os.path.exists(final_app_path):
                            shutil.rmtree(final_app_path)
                            
                        shutil.move(found_app, final_app_path)
                        shutil.rmtree(temp_extract_dir)
                        final_path = final_app_path
                    else:
                        final_path = temp_extract_dir

            else:
                app_dir = os.path.dirname(final_path)
                app_name = os.path.basename(final_path)
                name_no_ext, ext = os.path.splitext(app_name)
                clean_name = name_no_ext.replace('.', ' ').replace('_', ' ') + ext
                clean_app_path = os.path.join(app_dir, clean_name)
                
                if final_path != clean_app_path:
                    if os.path.exists(clean_app_path):
                        os.remove(clean_app_path)
                    os.rename(final_path, clean_app_path)
                    final_path = clean_app_path

            self.finished.emit(final_path)

        except Exception as e:
            if os.path.exists(self.save_path):
                os.remove(self.save_path)
            self.error.emit(f"Download or Extraction failed:\n{str(e)}")

# ==========================================
# SPLASH SCREEN
# ==========================================
class SplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedSize(450, 250)
        self.setObjectName("SplashWindow")

        self.setWindowIcon(app_icon())

        self.setStyleSheet("""
            QWidget#SplashWindow {
                background-color: #FAFAFA; 
                border: 2px solid #DADCE0; 
                border-radius: 10px;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        logo_label = QLabel()
        logo_label.setPixmap(logo_pixmap(90))
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(logo_label)

        title = QLabel("McMaster Research Group\nfor Stable Isotopologues")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #111;")
        layout.addWidget(title)
        
        subtitle = QLabel("Data Normalization Tool Installer")
        subtitle.setFont(QFont("Arial", 11))
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #555;")
        layout.addWidget(subtitle)

        self.loading_text = QLabel("Connecting to GitHub...")
        self.loading_text.setFont(QFont("Arial", 9))
        self.loading_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_text.setStyleSheet("color: #888; margin-top: 10px;")
        layout.addWidget(self.loading_text)

        self.min_time_passed = False
        self.data_ready = False
        self.next_window = None

        self.setWindowOpacity(0.0)
        
        screen = QApplication.primaryScreen().geometry()
        center_x = screen.width() // 2
        center_y = screen.height() // 2
        
        self.geom_anim = QPropertyAnimation(self, b"geometry")
        self.geom_anim.setDuration(700)
        self.geom_anim.setStartValue(QRect(center_x, center_y, 0, 0))
        self.geom_anim.setEndValue(QRect(center_x - 225, center_y - 125, 450, 250))
        self.geom_anim.setEasingCurve(QEasingCurve.Type.OutBack)
        
        self.fade_anim = QPropertyAnimation(self, b"windowOpacity")
        self.fade_anim.setDuration(500)
        self.fade_anim.setStartValue(0.0)
        self.fade_anim.setEndValue(1.0)
        
        self.anim_group = QParallelAnimationGroup()
        self.anim_group.addAnimation(self.geom_anim)
        self.anim_group.addAnimation(self.fade_anim)
        self.anim_group.start()

        QTimer.singleShot(2200, self.mark_min_time_passed)

    def mark_min_time_passed(self):
        self.min_time_passed = True
        self.check_ready()

    def mark_data_ready(self):
        self.data_ready = True
        self.loading_text.setText("Starting Installer...")
        self.check_ready()

    def check_ready(self):
        if self.min_time_passed and self.data_ready and self.next_window:
            self.fade_out()

    def fade_out(self):
        self.fade_out_anim = QPropertyAnimation(self, b"windowOpacity")
        self.fade_out_anim.setDuration(400)
        self.fade_out_anim.setStartValue(1.0)
        self.fade_out_anim.setEndValue(0.0)
        self.fade_out_anim.finished.connect(self.close)
        self.fade_out_anim.finished.connect(self.next_window.show)
        self.fade_out_anim.start()

# ==========================================
# MAIN GUI
# ==========================================
class InstallerApp(QWidget):
    fetch_finished = pyqtSignal()
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MRSI DNT - Installer")
        self.setMinimumWidth(550)
        
        self.setWindowIcon(app_icon())

        self.layout = QVBoxLayout(self)
        self.layout.setSizeConstraint(QVBoxLayout.SizeConstraint.SetFixedSize)
        self.layout.setContentsMargins(20, 20, 20, 20)

        self.releases_data = []
        self.download_thread = None

        self.setStyleSheet("""
            QWidget { background-color: #FAFAFA; color: #202124; font-family: Arial; }
            QPushButton { background-color: #ECECEC; color: #333; border: 1px solid #CCC; border-radius: 6px; padding: 7px 12px; font-weight: bold; }
            QPushButton:hover { background-color: #E0E0E0; }
            QPushButton:disabled { background-color: #f5f5f5; color: #aaa; border: 1px solid #ddd; }
            
            QPushButton#primaryBtn { background-color: #4CAF50; color: white; border: none; }
            QPushButton#primaryBtn:hover { background-color: #45A049; }
            QPushButton#primaryBtn:disabled { background-color: #a5d6a7; color: #eee; }
            
            QPushButton#cancelBtn { background-color: #F44336; color: white; border: none; }
            QPushButton#cancelBtn:hover { background-color: #D32F2F; }
            
            QPushButton#toggleBtn { background-color: transparent; border: none; color: #2196F3; text-align: left; padding: 0px; font-size: 13px; font-weight: normal; }
            QPushButton#toggleBtn:hover { color: #0b7dda; text-decoration: underline; }
            
            QComboBox { background-color: white; border: 1px solid #CCC; color: black; border-radius: 4px; padding: 5px; font-size: 13px; }
            
            QGroupBox { border: 1px solid #DADCE0; background-color: #F7F7F7; border-radius: 6px; margin-top: 12px; padding: 12px; font-weight: bold; }
            QLineEdit { background-color: white; border: 1px solid #CCC; color: black; border-radius: 4px; padding: 5px; }
            QTextEdit { background-color: white; border: 1px solid #CCC; color: #333; border-radius: 6px; padding: 10px; }
            
            QProgressBar { border: 1px solid #AAA; background-color: white; border-radius: 6px; height: 18px; text-align: center; font-weight: bold;}
            QProgressBar::chunk { background-color: #4CAF50; border-radius: 6px; }
        """)

        self.init_ui()
        self.fetch_releases()

    def init_ui(self):
        # ---- Header ----
        header = QHBoxLayout()
        logo_label = QLabel()
        logo_label.setPixmap(logo_pixmap(70))
        logo_label.setFixedSize(70, 70)
        logo_label.setCursor(Qt.CursorShape.PointingHandCursor)

        def open_mcmaster_website(event):
            QDesktopServices.openUrl(QUrl("https://science.mcmaster.ca"))

        logo_label.mousePressEvent = open_mcmaster_website

        header.addWidget(logo_label)
        header.addStretch()

        title_layout = QVBoxLayout()
        title_layout.setSpacing(0)
        
        title = QLabel("McMaster Research Group for Stable Isotopologues")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        subtitle = QLabel("Data Normalization Tool Installer")
        subtitle.setFont(QFont("Arial", 11))
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        title_layout.addWidget(title)
        title_layout.addWidget(subtitle)
        header.addLayout(title_layout)
        header.addStretch()

        dummy = QLabel()
        dummy.setFixedSize(70, 70)
        header.addWidget(dummy)
        self.layout.addLayout(header)

        # Divider
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setStyleSheet("color: #DADCE0;")
        self.layout.addWidget(divider)

        # ---- Version Selection & Notes ----
        version_layout = QHBoxLayout()
        version_label = QLabel("Select Version:")
        version_label.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        version_layout.addWidget(version_label)
        
        self.version_combo = QComboBox()
        self.version_combo.addItem("Fetching releases...")
        self.version_combo.setEnabled(False)
        self.version_combo.setMinimumWidth(250)
        self.version_combo.currentIndexChanged.connect(self.on_version_changed)
        version_layout.addWidget(self.version_combo)
        version_layout.addStretch()
        self.layout.addLayout(version_layout)

        # Toggle Button
        self.toggle_notes_btn = QPushButton("▶ Show Release Notes")
        self.toggle_notes_btn.setObjectName("toggleBtn")
        self.toggle_notes_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.toggle_notes_btn.clicked.connect(self.toggle_notes)
        self.layout.addWidget(self.toggle_notes_btn)

        self.release_notes = QTextEdit()
        self.release_notes.setReadOnly(True)
        self.release_notes.setFixedSize(500, 150)
        self.release_notes.hide()
        self.layout.addWidget(self.release_notes)

        # ---- Clean Installation Settings ----
        settings_group = QGroupBox("Installation Settings")
        sg_layout = QVBoxLayout(settings_group)

        p_layout = QHBoxLayout()
        p_layout.addWidget(QLabel("Install To:"))
        self.path_input = QLineEdit()
        
        if sys.platform == "win32":
            default_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
        else:
            default_path = os.path.join(os.path.expanduser("~"), "Desktop")
            
        self.path_input.setText(default_path)
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.clicked.connect(self.browse_path)
        
        p_layout.addWidget(self.path_input, 1)
        p_layout.addWidget(self.browse_btn)
        sg_layout.addLayout(p_layout)

        self.open_checkbox = QCheckBox("Launch application immediately after installation")
        self.open_checkbox.setChecked(True)
        self.open_checkbox.setStyleSheet("font-weight: normal;")
        sg_layout.addWidget(self.open_checkbox)

        self.layout.addWidget(settings_group)

        # ---- Status & Progress ----
        self.status_label = QLabel("Ready.")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        self.layout.addWidget(self.progress_bar)

        self.stats_label = QLabel("")
        self.stats_label.setFont(QFont("Arial", 9))
        self.stats_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.stats_label.setStyleSheet("color: #666;")
        self.stats_label.hide()
        self.layout.addWidget(self.stats_label)

        # ---- Controls ----
        controls_layout = QHBoxLayout()
        
        self.cancel_btn = QPushButton("Cancel Download")
        self.cancel_btn.setObjectName("cancelBtn")
        self.cancel_btn.setFixedHeight(40)
        self.cancel_btn.hide()
        self.cancel_btn.clicked.connect(self.cancel_download)
        controls_layout.addWidget(self.cancel_btn)
        
        controls_layout.addStretch()

        self.install_btn = QPushButton("Download and Install")
        self.install_btn.setObjectName("primaryBtn")
        self.install_btn.setFixedHeight(40)
        self.install_btn.setFixedWidth(200)
        self.install_btn.setEnabled(False)
        self.install_btn.clicked.connect(self.start_installation)
        controls_layout.addWidget(self.install_btn)
        
        self.layout.addLayout(controls_layout)

    def toggle_notes(self):
        if self.release_notes.isHidden():
            self.release_notes.show()
            self.toggle_notes_btn.setText("▼ Hide Release Notes")
        else:
            self.release_notes.hide()
            self.toggle_notes_btn.setText("▶ Show Release Notes")

    def browse_path(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Installation Folder", self.path_input.text())
        if folder:
            self.path_input.setText(folder)

    def fetch_releases(self):
        self.fetch_thread = FetchReleasesThread()
        self.fetch_thread.releases_fetched.connect(self.on_releases_fetched)
        self.fetch_thread.error_occurred.connect(self.on_error)
        self.fetch_thread.start()

    def on_releases_fetched(self, releases):
        self.releases_data = releases
        self.version_combo.clear()
        
        if not releases:
            self.version_combo.addItem("No valid releases found.")
            self.fetch_finished.emit() 
            return

        for i, r in enumerate(releases):
            tag = r.get("tag_name", "Unknown")
            name = r.get("name", tag)
            if i == 0:
                self.version_combo.addItem(f"Latest Version ({name})", r)
            else:
                self.version_combo.addItem(f"Older: {name} ({tag})", r)

        self.version_combo.setEnabled(True)
        self.install_btn.setEnabled(True)
        self.fetch_finished.emit() 

    def on_version_changed(self, index):
        if index < 0 or not self.releases_data: return
        release_data = self.version_combo.itemData(index)
        body = release_data.get("body", "No release notes provided.")
        self.release_notes.setMarkdown(body)

    def on_error(self, msg):
        self.status_label.setText("Error fetching releases.")
        self.release_notes.setPlainText(msg)
        QMessageBox.critical(self, "Error", msg)
        self.fetch_finished.emit() 

    # --- OS Asset Selection ---
    def get_os_asset(self, release_data):
        assets = release_data.get("assets", [])
        is_windows = sys.platform == "win32"
        
        # Sort files to prevent Windows from grabbing the Mac .zip
        win_files = []
        mac_files = []
        
        for asset in assets:
            name = asset["name"].lower()
            if ".app" in name or "mac" in name or name.endswith(".dmg") or name.endswith(".pkg"):
                mac_files.append(asset)
            elif name.endswith(".exe") or "win" in name:
                win_files.append(asset)
            elif name.endswith(".zip"):
                # It's a zip file with no obvious OS indicators. Assume it's Windows fallback.
                win_files.append(asset)
                
        target_list = win_files if is_windows else mac_files
        
        if target_list:
            # If on Windows, prioritize the .exe if it exists, otherwise return the zip
            if is_windows:
                for a in target_list:
                    if a["name"].lower().endswith(".exe"):
                        return a
            return target_list[0]
            
        return assets[0] if assets else None

    def start_installation(self):
        index = self.version_combo.currentIndex()
        if index < 0: return
        release_data = self.version_combo.itemData(index)

        target_asset = self.get_os_asset(release_data)
        if not target_asset:
            QMessageBox.warning(self, "No Asset", "Could not find a valid file for your OS in this release.")
            return

        download_url = target_asset["browser_download_url"]
        filename = target_asset["name"]
        save_path = os.path.join(self.path_input.text(), filename)

        self.install_btn.hide()
        self.cancel_btn.show()
        self.version_combo.setEnabled(False)
        self.browse_btn.setEnabled(False)
        self.path_input.setEnabled(False)
        
        self.status_label.setText(f"Downloading {filename}...")
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        
        self.stats_label.setText("Starting download...") 
        self.stats_label.show()                           

        self.download_thread = DownloadThread(download_url, save_path)
        self.download_thread.progress.connect(self.progress_bar.setValue)
        self.download_thread.status_update.connect(self.status_label.setText)
        self.download_thread.stats_update.connect(self.stats_label.setText) 
        self.download_thread.finished.connect(self.on_download_finished)
        self.download_thread.error.connect(self.on_download_error)
        self.download_thread.start()

    def cancel_download(self):
        if self.download_thread and self.download_thread.isRunning():
            self.download_thread.cancel()
            self.cancel_btn.setEnabled(False)
            self.status_label.setText("Cancelling download...")

    def on_download_finished(self, final_filepath):
        self.progress_bar.setValue(100)
        self.stats_label.hide()
        
        if self.open_checkbox.isChecked() and os.path.exists(final_filepath):
            self.status_label.setText("Installation Complete! Launching application...")
            try:
                if sys.platform == "win32":
                    os.startfile(final_filepath)
                else:
                    subprocess.run(["open", final_filepath])
                    
                QTimer.singleShot(1500, QApplication.quit)
                return
            except Exception as e:
                QMessageBox.warning(self, "Launch Failed", f"Downloaded successfully but could not auto-launch:\n{str(e)}")
        else:
            self.status_label.setText("Installation Complete!")
            QMessageBox.information(self, "Success", f"Successfully installed to:\n{final_filepath}")

        self.cancel_btn.hide()
        self.cancel_btn.setEnabled(True)
        self.install_btn.show()
        self.install_btn.setEnabled(True)
        self.install_btn.setText("Reinstall")
        self.version_combo.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.path_input.setEnabled(True)

    def on_download_error(self, msg):
        self.status_label.setText("Download Failed/Cancelled.")
        self.stats_label.hide()
        
        self.cancel_btn.hide()
        self.cancel_btn.setEnabled(True)
        self.install_btn.show()
        self.install_btn.setEnabled(True)
        self.version_combo.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.path_input.setEnabled(True)
        self.progress_bar.hide()
        
        if "cancelled" not in msg.lower():
            QMessageBox.critical(self, "Error", msg)

if __name__ == "__main__":
    # --- Declare the App ID to Windows so the taskbar picks up the icon ---
    if sys.platform == "win32":
        try:
            import ctypes
            myappid = 'mrsi.dnt.installer.1.0' 
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

    app = QApplication(sys.argv)
    
    main_window = InstallerApp()
    splash = SplashScreen()
    splash.next_window = main_window
    main_window.fetch_finished.connect(splash.mark_data_ready)
    splash.show()
    
    sys.exit(app.exec())