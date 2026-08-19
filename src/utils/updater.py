import os
import sys
import time
import requests
import subprocess
import tempfile
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import (QVBoxLayout, QLabel, QDialog, QGridLayout, QHBoxLayout, QPushButton)
from PyQt6.QtCore import Qt, pyqtSignal

# Update this to match the tag name on your GitHub Release (e.g., "v1.0.1")
CURRENT_VERSION = "v2.0.0"
GITHUB_REPO = "ibrahim-parvez/MRSI_Data_Normalization_Tool"

class UpdateAvailableDialog(QDialog):
    """A custom, styled dialog for prompting the user to update during startup."""
    def __init__(self, current_version, new_version, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Update Available")
        self.setFixedSize(350, 200)
        
        self.dark_mode = getattr(parent, 'dark_mode', False)
        
        if self.dark_mode:
            self.setStyleSheet("QDialog { background-color: #2A2B2E; color: #E8EAED; } QLabel { color: #E8EAED; }")
            header_color = "#FF6B8B" # A lighter pink/maroon so it's readable in dark mode
            cancel_bg = "#424242"
            cancel_hover = "#616161"
        else:
            self.setStyleSheet("QDialog { background-color: #FAFAFA; color: #202124; } QLabel { color: #202124; }")
            header_color = "#7A003C" # McMaster Maroon for light mode
            cancel_bg = "#9E9E9E"
            cancel_hover = "#757575"

        layout = QVBoxLayout(self)
        
        # Header
        header_label = QLabel("🚀 A New Version is Available!")
        header_label.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {header_color};")
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        layout.addSpacing(15)
        
        # Versions Display (Grid for clean alignment)
        version_layout = QGridLayout()
        version_layout.setContentsMargins(20, 0, 20, 0)
        
        version_layout.addWidget(QLabel("<b>Current Version:</b>"), 0, 0)
        version_layout.addWidget(QLabel(current_version), 0, 1)
        
        version_layout.addWidget(QLabel("<b>New Version:</b>"), 1, 0)
        # Highlight the new version in green (works well in both themes)
        new_v_label = QLabel(f"<span style='color: #4CAF50; font-weight: bold;'>{new_version}</span>")
        version_layout.addWidget(new_v_label, 1, 1)
        
        layout.addLayout(version_layout)
        layout.addStretch()
        
        # Buttons
        btn_layout = QHBoxLayout()
        self.later_btn = QPushButton("Update Later")
        self.update_btn = QPushButton("Update Now")
        
        # Style the buttons
        self.update_btn.setStyleSheet("""
            QPushButton { background-color: #4CAF50; color: white; padding: 8px 15px; font-weight: bold; border-radius: 4px; border: none; }
            QPushButton:hover { background-color: #45A049; }
        """)
        
        # Use dynamic colors for the cancel button ---
        self.later_btn.setStyleSheet(f"""
            QPushButton {{ background-color: {cancel_bg}; color: white; padding: 8px 15px; border-radius: 4px; border: none; }}
            QPushButton:hover {{ background-color: {cancel_hover}; }}
        """)
        
        self.update_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.later_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.update_btn.clicked.connect(self.accept)
        self.later_btn.clicked.connect(self.reject)
        
        btn_layout.addWidget(self.later_btn)
        btn_layout.addWidget(self.update_btn)
        
        layout.addLayout(btn_layout)



class AutoUpdater(QThread):
    check_finished = pyqtSignal(bool, str, str) # has_update, latest_version, download_url
    progress_updated = pyqtSignal(int)
    download_finished = pyqtSignal(str) # path to downloaded file
    error_occurred = pyqtSignal(str)

    def __init__(self, mode="check", url=""):
        super().__init__()
        self.mode = mode
        self.url = url

    def run(self):
        if self.mode == "check":
            self.check_for_updates()
        elif self.mode == "download":
            self.download_update()

    def check_for_updates(self):
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            response = requests.get(url, timeout=5)
            response.raise_for_status()
            data = response.json()
            latest_version = data["tag_name"]
            
            if latest_version != CURRENT_VERSION:
                download_url = ""
                # Find the correct asset based on OS
                for asset in data["assets"]:
                    if sys.platform == "win32" and asset["name"].endswith(".exe"):
                        download_url = asset["browser_download_url"]
                    elif sys.platform == "darwin" and asset["name"].endswith(".zip"):
                        download_url = asset["browser_download_url"]
                
                if download_url:
                    self.check_finished.emit(True, latest_version, download_url)
                    return
            
            self.check_finished.emit(False, "", "")
        except Exception as e:
            self.error_occurred.emit(f"Failed to check for updates: {str(e)}")

    def download_update(self):
        try:
            response = requests.get(self.url, stream=True, timeout=10)
            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0))
            
            # FIX: Save the temp file to the system's temporary directory, outside the app
            temp_dir = tempfile.gettempdir()
            filename = "MRSI_Update_New.exe" if sys.platform == "win32" else "MRSI_Update_New.zip"
            download_path = os.path.join(temp_dir, filename)

            downloaded_size = 0
            with open(download_path, 'wb') as file:
                for data in response.iter_content(chunk_size=8192):
                    downloaded_size += len(data)
                    file.write(data)
                    if total_size > 0:
                        progress = int((downloaded_size / total_size) * 100)
                        self.progress_updated.emit(progress)

            self.download_finished.emit(download_path)
        except Exception as e:
            self.error_occurred.emit(f"Download failed: {str(e)}")

def apply_update_and_restart(download_path):
    """Generates and runs an OS-specific script to replace the running executable."""
    is_frozen = getattr(sys, 'frozen', False)
    if not is_frozen:
        print("Cannot auto-update while running from raw Python script. Please compile first.")
        return

    current_exe = sys.executable
    temp_dir = tempfile.gettempdir()

    if sys.platform == "win32":
        bat_path = os.path.join(temp_dir, "update_script.bat")
        
        # FIX 1: Added a WaitLoop so the script keeps trying until the file unlocks
        with open(bat_path, "w") as bat_file:
            bat_file.write(f"""@echo off
                            :WaitLoop
                            ping 127.0.0.1 -n 2 > NUL
                            move /y "{download_path}" "{current_exe}"
                            if errorlevel 1 goto WaitLoop
                            start "" "{current_exe}"
                            del "%~f0"
                            """)
        
        clean_env = os.environ.copy()
        keys_to_remove = [k for k in clean_env if k.startswith('_PYI_') or k.startswith('_MEI')]
        for k in keys_to_remove:
            clean_env.pop(k, None)

        subprocess.Popen(
            [bat_path], 
            creationflags=subprocess.CREATE_NO_WINDOW, 
            env=clean_env
        )
        # FIX 2: Use os._exit(0) to instantly kill the app and release the .exe lock
        os._exit(0)

    elif sys.platform == "darwin":
        app_bundle_path = os.path.dirname(os.path.dirname(os.path.dirname(current_exe)))
        parent_dir = os.path.dirname(app_bundle_path) 
        app_name = os.path.basename(app_bundle_path)  

        sh_path = os.path.join(temp_dir, "update_script.sh")
        with open(sh_path, "w") as sh_file:
            sh_file.write(f"""#!/bin/bash
                            sleep 2
                            rm -rf "{app_bundle_path}"
                            unzip -q "{download_path}" -d "{parent_dir}" -x "__MACOSX/*"
                            rm "{download_path}"
                            open "{parent_dir}/{app_name}"
                            rm "$0"
                            """)
        os.chmod(sh_path, 0o755) 
        
        clean_env = os.environ.copy()
        keys_to_remove = [k for k in clean_env if k.startswith('_PYI_') or k.startswith('_MEI')]
        for k in keys_to_remove:
            clean_env.pop(k, None)
            
        subprocess.Popen([sh_path], start_new_session=True, env=clean_env)
        # FIX 2: Apply to Mac as well for a clean instant kill
        os._exit(0)