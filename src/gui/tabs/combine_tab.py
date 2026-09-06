import os
import re
from PyQt6.QtGui import QPainter, QColor, QPen, QFont, QDesktopServices
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QGroupBox, QLineEdit, QAbstractItemView, 
    QMessageBox, QRadioButton, QButtonGroup, QTableWidget, QTableWidgetItem, 
    QHeaderView, QStyle, QCheckBox, QProgressBar, QListWidget, QScrollArea, QFrame
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl

import utils.settings as settings
from gui.tabs.combine_processors.combine_water import WaterCombineWorker
from gui.tabs.combine_processors.combine_carbonate import CarbonateCombineWorker

# =========================================================================
# BACKGROUND WORKER FOR FETCHING SHEET NAMES & MATERIALS
# =========================================================================

class FileScannerWorker(QThread):
    """Reads the sheet name and extracts unique identifiers in the background."""
    result_ready = pyqtSignal(str, str, list, list)  # Emits (file_path, sheet_name, references, samples)

    def __init__(self, file_path, default_fallback, mode, known_refs):
        super().__init__()
        self.file_path = file_path
        self.default_fallback = default_fallback
        self.mode = mode
        self.known_refs = known_refs

    def run(self):
        import pandas as pd 
        sheet_name = self.default_fallback
        refs, samples = [], []
        
        try:
            xl = pd.ExcelFile(self.file_path)
            if xl.sheet_names:
                sheet_name = xl.sheet_names[-1]  
                
            df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            
            # Locate Identifier 1 column safely
            id_col = None
            for col in df.columns:
                if str(col).strip().lower() == "identifier 1":
                    id_col = col
                    break
                    
            if id_col is not None:
                unique_materials = set()
                for val in df[id_col].dropna():
                    base = re.sub(r"\s*r\d+(\.\d+)?$", "", str(val), flags=re.IGNORECASE).strip()
                    if base:
                        unique_materials.add(base)
                
                # Heuristic Sorting
                for mat in sorted(unique_materials):
                    mat_upper = mat.upper().replace("STD", "").strip()
                    is_ref = False
                    
                    if self.mode == "water":
                        if "HECO2" in mat_upper or "CO2" in mat_upper or "C02" in mat_upper:
                            is_ref = True
                        else:
                            pat = [r'\bMRSI\b', r'\bMRSI[- ]?\d+\b', r'\bMRSI[- ]?STD', r'\bUSGS']
                            if any(re.search(p, mat_upper) for p in pat):
                                is_ref = True
                    else: # carbonate
                        if mat_upper.startswith("CO2") or mat_upper.startswith("HECO2") or mat_upper.startswith("C02"):
                            is_ref = True
                        else:
                            for known in self.known_refs:
                                if known and (known in mat_upper or (len(known) >= 4 and known in mat_upper)):
                                    is_ref = True
                                    break
                                    
                    if is_ref: refs.append(mat)
                    else: samples.append(mat)
                    
        except Exception:
            pass 
        
        self.result_ready.emit(self.file_path, sheet_name, refs, samples)

# =========================================================================
# CUSTOM UI WIDGETS
# =========================================================================

class OutputDragDropBox(QGroupBox):
    fileDropped = pyqtSignal(str) 

    def __init__(self, title, parent=None):
        super().__init__(title, parent)
        self.setAcceptDrops(True)
        self.drag_active = False
        self.drop_enabled = False 

    def dragEnterEvent(self, event):
        if not self.drop_enabled:
            event.ignore()
            return
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(('.xlsx', '.xls')):
                    event.accept()
                    self.drag_active = True
                    self.update() 
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.drag_active = False
        self.update()

    def dropEvent(self, event):
        self.drag_active = False
        self.update()
        
        if not self.drop_enabled:
            return
            
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(('.xlsx', '.xls')):
                self.fileDropped.emit(path)
                event.accept()
                return 

    def paintEvent(self, event):
        super().paintEvent(event)

        if self.drag_active and self.drop_enabled:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            rect = self.contentsRect()
            overlay_color = QColor("#E3F2FD") 
            overlay_color.setAlpha(200) 
            painter.setBrush(overlay_color)
            
            pen = QPen(QColor("#2196F3"))
            pen.setWidth(3)
            pen.setStyle(Qt.PenStyle.DashLine)
            painter.setPen(pen)
            
            painter.drawRoundedRect(rect.adjusted(5, 5, -5, -5), 10, 10)

class DragDropBox(QGroupBox):
    filesDropped = pyqtSignal(list) 

    def __init__(self, title, parent=None):
        super().__init__(title, parent)
        self.setAcceptDrops(True)
        self.drag_active = False

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(('.xlsx', '.xls')):
                    event.accept()
                    self.drag_active = True
                    self.update() 
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.drag_active = False
        self.update()

    def dropEvent(self, event):
        self.drag_active = False
        self.update()
        
        valid_files = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(('.xlsx', '.xls')):
                valid_files.append(path)
                
        if valid_files:
            self.filesDropped.emit(valid_files)
            event.accept()

    def paintEvent(self, event):
        super().paintEvent(event)

        if self.drag_active:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            rect = self.contentsRect()
            overlay_color = QColor("#E3F2FD") 
            overlay_color.setAlpha(200) 
            painter.setBrush(overlay_color)
            
            pen = QPen(QColor("#2196F3"))
            pen.setWidth(3)
            pen.setStyle(Qt.PenStyle.DashLine)
            painter.setPen(pen)
            
            painter.drawRoundedRect(rect.adjusted(5, 5, -5, -5), 10, 10)
            
            painter.setPen(QColor("#0D47A1"))
            font = QFont("Arial", 16, QFont.Weight.Bold)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, "📂 Drop Excel File(s) Here")

class FileTableWidget(QTableWidget):
    def paintEvent(self, event):
        super().paintEvent(event)
        
        if self.rowCount() == 0:
            painter = QPainter(self.viewport())
            rect = self.viewport().rect()
            painter.setPen(QColor("#888888"))
            font = self.font()
            font.setPointSize(14)
            font.setItalic(True)
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, "📥 Select mode & Drop File(s) Here")

# =========================================================================
# MAIN TAB CLASS
# =========================================================================
            
class CombineTab(QWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)  
        self._sheet_workers = []  
        self.init_ui()

    def init_ui(self):
        # 1. Main outer layout (flush with the tab borders)
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # 2. Scroll area to contain everything
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        
        # 3. Inner container widget where all the boxes actually go
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(10, 10, 10, 10) 
        layout.setSpacing(6)

        self.mode_warning_label = QLabel("Please select either water or carbonate")
        self.mode_warning_label.setStyleSheet("color: #d32f2f; font-size: 11px; font-weight: bold;")
        self.mode_warning_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.mode_warning_label)

        # --- 0. Mode Configuration ---
        mode_layout = QHBoxLayout()
        
        self.btn_water = QPushButton("Water")
        self.btn_carbonate = QPushButton("Carbonate")
        self.btn_water.setCheckable(True)
        self.btn_carbonate.setCheckable(True)
        
        self.btn_water.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_carbonate.setCursor(Qt.CursorShape.PointingHandCursor)
        
        toggle_style = """
            QPushButton {
                background-color: #f3f3f3;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 4px 12px;
                color: #333;
                min-width: 80px; 
            }
            QPushButton:hover {
                background-color: #e5e5e5;
            }
            QPushButton:checked {
                font-weight: bold;
                border: 2px solid #333;
                background-color: #e5e5e5; 
            }
        """
        self.btn_water.setStyleSheet(toggle_style)
        self.btn_carbonate.setStyleSheet(toggle_style)
        
        self.btn_water.clicked.connect(self._on_mode_clicked)
        self.btn_carbonate.clicked.connect(self._on_mode_clicked)

        mode_layout.addStretch()
        mode_layout.addWidget(self.btn_water)
        mode_layout.addWidget(self.btn_carbonate)
        mode_layout.addStretch()
        layout.addLayout(mode_layout)

        # --- 1. File Handling ---
        copy_group = QGroupBox("File Handling")
        copy_layout = QHBoxLayout()
        copy_layout.setContentsMargins(8, 16, 8, 8) 
        
        self.handling_group = QButtonGroup(self)
        self.radio_temp_copy = QRadioButton("Process data on temp files")
        self.radio_modify_orig = QRadioButton("Process data on original files")
        self.radio_temp_copy.setChecked(True) 
        
        self.handling_group.addButton(self.radio_temp_copy)
        self.handling_group.addButton(self.radio_modify_orig)
        
        copy_layout.addWidget(self.radio_temp_copy)
        copy_layout.addWidget(self.radio_modify_orig)
        copy_layout.addStretch() 
        
        copy_group.setLayout(copy_layout)
        layout.addWidget(copy_group)

        # --- Carbonate Specific Toggles ---
        self.carb_settings_group = QGroupBox("Additional Carbonate Settings")
        carb_layout = QHBoxLayout()
        carb_layout.setContentsMargins(8, 16, 8, 8)
        self.chk_yield = QCheckBox("Enable Yield Table Calculations")
        self.chk_co2 = QCheckBox("Enable CO2 Table Calculations")
        
        carb_layout.addWidget(self.chk_yield)
        carb_layout.addWidget(self.chk_co2)
        carb_layout.addStretch()
        self.carb_settings_group.setLayout(carb_layout)
        layout.addWidget(self.carb_settings_group)
        self.carb_settings_group.hide() # Hidden initially

        # --- 2. File List Section ---
        list_group = DragDropBox("Raw Files to Combine")
        list_group.filesDropped.connect(self._add_files_to_table)
        
        list_layout = QVBoxLayout()
        list_layout.setContentsMargins(8, 16, 8, 8)
        list_layout.setSpacing(2) 
        
        top_controls = QHBoxLayout()
        self.browse_files_btn = QPushButton(" Browse Files")
        folder_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)
        self.browse_files_btn.setIcon(folder_icon)
        self.browse_files_btn.clicked.connect(self.add_files)
        self.browse_files_btn.setStyleSheet("padding: 5px 15px;")
        
        self.clear_btn = QPushButton(" Clear All")
        trash_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon)
        self.clear_btn.setIcon(trash_icon)
        self.clear_btn.clicked.connect(self.clear_all)
        self.clear_btn.setStyleSheet("padding: 5px 15px;")

        top_controls.addWidget(self.browse_files_btn)
        top_controls.addStretch() 
        top_controls.addWidget(self.clear_btn)
        list_layout.addLayout(top_controls)
        
        self.file_table = FileTableWidget(0, 5)
        self.file_table.setHorizontalHeaderLabels(["File Name", "Sheet Name", "Reference Materials", "Samples", ""])
        
        self.file_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.file_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.file_table.setColumnWidth(4, 40) 
        
        self.file_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection) 
        self.file_table.setAlternatingRowColors(True)
        self.file_table.setMinimumHeight(250)
        list_layout.addWidget(self.file_table)
        
        self.footer_hint = QLabel("Drag & drop more files anywhere in this box...")
        self.footer_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.footer_hint.setStyleSheet("color: #999999; font-style: italic; font-size: 11px;")
        self.footer_hint.setContentsMargins(0, 0, 0, 0) 
        self.footer_hint.hide()
        list_layout.addWidget(self.footer_hint)

        list_group.setLayout(list_layout)
        layout.addWidget(list_group)
        
        # --- 3. Output Configuration ---
        self.output_group = OutputDragDropBox("Final Combined Output")
        self.output_group.fileDropped.connect(self._on_output_file_dropped)
        
        output_layout = QVBoxLayout()
        output_layout.setContentsMargins(8, 16, 8, 8)
        output_layout.setSpacing(5) 
        
        self.out_mode_group = QButtonGroup(self)
        self.radio_new_file = QRadioButton("Create New Combined File")
        self.radio_append_file = QRadioButton("Append to Existing Combined File")
        self.radio_new_file.setChecked(True)
        self.out_mode_group.addButton(self.radio_new_file)
        self.out_mode_group.addButton(self.radio_append_file)
        
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(self.radio_new_file)
        mode_layout.addWidget(self.radio_append_file)
        mode_layout.addStretch()
        output_layout.addLayout(mode_layout)

        self.radio_new_file.toggled.connect(self._toggle_output_mode)

        row_out = QHBoxLayout()
        self.output_path_input = QLineEdit()
        
        desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        default_out_path = os.path.join(desktop_dir, "Combined_Normalization_Data.xlsx")
        self.output_path_input.setText(default_out_path)
        
        self.browse_out_btn = QPushButton(" Browse Files")
        self.browse_out_btn.setFixedWidth(130) 
        self.browse_out_btn.setIcon(folder_icon)
        self.browse_out_btn.clicked.connect(self.browse_output)
        
        self.out_label = QLabel("Output File:")
        row_out.addWidget(self.out_label)
        row_out.addWidget(self.output_path_input)
        row_out.addWidget(self.browse_out_btn)
        
        output_layout.addLayout(row_out)

        action_row = QHBoxLayout()
        self.open_checkbox = QCheckBox("Open file upon completion of processing")
        self.open_checkbox.setChecked(True) 
        self.open_checkbox.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.btn_open_file = QPushButton(" Open File")
        self.btn_open_file.setFixedWidth(130) 
        self.btn_open_file.setCursor(Qt.CursorShape.PointingHandCursor)
        
        open_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)
        self.btn_open_file.setIcon(open_icon)
        self.btn_open_file.clicked.connect(self.open_combined_file)
        
        action_row.addWidget(self.open_checkbox)
        action_row.addStretch() 
        action_row.addWidget(self.btn_open_file)
        
        output_layout.addLayout(action_row)
        self.output_group.setLayout(output_layout)
        layout.addWidget(self.output_group)

        # 4. Finalize layouts
        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith(('.xlsx', '.xls')):
                files.append(file_path)
        if files:
            self._add_files_to_table(files)

    def _add_files_to_table(self, files):
        if not self.btn_water.isChecked() and not self.btn_carbonate.isChecked():
            self.mode_warning_label.show()
            QMessageBox.warning(self, "Selection Required", "Please select either Water or Carbonate before adding files.")
            return

        existing_files = []
        for i in range(self.file_table.rowCount()):
            item = self.file_table.item(i, 0)
            if item: existing_files.append(item.data(Qt.ItemDataRole.UserRole))
                
        is_water = self.btn_water.isChecked()
        mode_str = "water" if is_water else "carbonate"
        default_sheet = "ExportGB1.wke" if is_water else "ExportGB2.wke"
        
        known_refs = []
        if not is_water:
            mats = settings.get_setting("REFERENCE_MATERIALS", sub_key="Carbonate") or []
            known_refs = [str(m.get("col_c", "")).upper().replace("STD", "").strip() for m in mats]
        
        for f in files:
            if f not in existing_files:
                row = self.file_table.rowCount()
                self.file_table.insertRow(row)
                self.file_table.setRowHeight(row, 80) # Expand height to fit lists nicely
                
                # --- Set File Name ---
                filename = os.path.basename(f)
                path_item = QTableWidgetItem(filename)
                path_item.setData(Qt.ItemDataRole.UserRole, f) 
                path_item.setFlags(path_item.flags() & ~Qt.ItemFlag.ItemIsEditable) 
                self.file_table.setItem(row, 0, path_item)
                
                # --- Set Animated Loading UI in the Sheet Name Column ---
                loading_widget = QWidget()
                lw_layout = QHBoxLayout(loading_widget)
                lw_layout.setContentsMargins(10, 2, 10, 2)
                pbar = QProgressBar()
                pbar.setRange(0, 0) 
                pbar.setFixedHeight(12)
                pbar.setTextVisible(False)
                lbl = QLabel("Reading...")
                lbl.setStyleSheet("color: #666; font-size: 11px; font-style: italic;")
                lw_layout.addWidget(pbar)
                lw_layout.addWidget(lbl)
                self.file_table.setCellWidget(row, 1, loading_widget)
                
                # Setup Loading labels for lists
                for c in [2, 3]:
                    loading_lbl = QLabel("Fetching materials...")
                    loading_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    loading_lbl.setStyleSheet("color: gray; font-style: italic;")
                    self.file_table.setCellWidget(row, c, loading_lbl)
                
                # --- Setup Delete Button ---
                del_btn = QPushButton("−") 
                del_btn.setFixedSize(24, 24)
                del_btn.setToolTip("Remove this file")
                del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                del_btn.setStyleSheet("""
                    QPushButton { background-color: #ff4d4d; color: white; border: none; border-radius: 4px; font-weight: bold; font-size: 16px; padding: 0px; }
                    QPushButton:hover { background-color: #d32f2f; }
                """)
                del_btn.clicked.connect(lambda _, r=path_item: self._remove_specific_row(r))
                
                cell_widget = QWidget()
                cell_layout = QHBoxLayout(cell_widget)
                cell_layout.setContentsMargins(0, 0, 0, 0)
                cell_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
                cell_layout.addWidget(del_btn)
                self.file_table.setCellWidget(row, 4, cell_widget)
                
                # --- Spawn Background Thread ---
                worker = FileScannerWorker(f, default_sheet, mode_str, known_refs)
                worker.result_ready.connect(self._on_file_scanned)
                self._sheet_workers.append(worker)
                worker.finished.connect(lambda w=worker: self._cleanup_worker(w))
                worker.start()
            
        self._update_footer_visibility()

    def _cleanup_worker(self, worker):
        if worker in self._sheet_workers:
            self._sheet_workers.remove(worker)
        worker.deleteLater()

    def _on_file_scanned(self, file_path, sheet_name, refs, samples):
        for row in range(self.file_table.rowCount()):
            item = self.file_table.item(row, 0)
            if item and item.data(Qt.ItemDataRole.UserRole) == file_path:
                self.file_table.removeCellWidget(row, 1)
                sheet_item = QTableWidgetItem(sheet_name)
                self.file_table.setItem(row, 1, sheet_item)
                
                # Setup References Drag & Drop List
                self.file_table.removeCellWidget(row, 2)
                ref_list = QListWidget()
                ref_list.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
                ref_list.setDefaultDropAction(Qt.DropAction.MoveAction)
                ref_list.addItems(refs)
                self.file_table.setCellWidget(row, 2, ref_list)
                
                # Setup Samples Drag & Drop List
                self.file_table.removeCellWidget(row, 3)
                samp_list = QListWidget()
                samp_list.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
                samp_list.setDefaultDropAction(Qt.DropAction.MoveAction)
                samp_list.addItems(samples)
                self.file_table.setCellWidget(row, 3, samp_list)
                break

    def _remove_specific_row(self, item):
        row = self.file_table.row(item)
        if row >= 0:
            self.file_table.removeRow(row)
        self._update_footer_visibility()

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Raw Excel Files", "", "Excel Files (*.xlsx *.xls)"
        )
        if files:
            self._add_files_to_table(files)

    def remove_selected(self):
        rows = sorted(set(index.row() for index in self.file_table.selectedIndexes()), reverse=True)
        for row in rows:
            self.file_table.removeRow(row)

    def clear_all(self):
        self.file_table.setRowCount(0)
        self._update_footer_visibility()

    def _toggle_output_mode(self):
        if self.radio_new_file.isChecked():
            self.output_group.drop_enabled = False 
            self.out_label.setText("Output File:")
            desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
            self.output_path_input.setText(os.path.join(desktop_dir, "Combined_Normalization_Data.xlsx"))
        else:
            self.output_group.drop_enabled = True 
            self.out_label.setText("Existing File:")
            self.output_path_input.clear()
            self.output_path_input.setPlaceholderText("Select or drag-and-drop the existing file anywhere in this box...")

    def _on_output_file_dropped(self, path):
        self.output_path_input.setText(path)

    def browse_output(self):
        if self.radio_new_file.isChecked():
            path, _ = QFileDialog.getSaveFileName(
                self, "Save Combined File", "Combined_Normalization_Data.xlsx", "Excel Files (*.xlsx)"
            )
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "Select Existing Combined File", "", "Excel Files (*.xlsx *.xls)"
            )
        
        if path:
            self.output_path_input.setText(path)

    def _on_mode_clicked(self, checked):
        sender = self.sender()
        if not checked:
            sender.setChecked(True)
            return

        if sender == self.btn_water:
            self.btn_carbonate.setChecked(False)
            self.carb_settings_group.hide()
        else:
            self.btn_water.setChecked(False)
            self.carb_settings_group.show()
            
        self.mode_warning_label.hide()
        
        if self.file_table.rowCount() > 0:
            reply = QMessageBox.question(self, "Change Mode", "Changing modes requires clearing the current file list to reset processing rules. Continue?", 
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.clear_all()
            else:
                sender.setChecked(False)
                if sender == self.btn_water:
                    self.btn_carbonate.setChecked(True)
                    self.carb_settings_group.show()
                else:
                    self.btn_water.setChecked(True)
                    self.carb_settings_group.hide()

    def get_run_parameters(self):
        if not self.btn_water.isChecked() and not self.btn_carbonate.isChecked():
            self.mode_warning_label.show()
            QMessageBox.warning(self, "Selection Required", "Please select either Water or Carbonate before running.")
            return None

        if self.file_table.rowCount() == 0:
            QMessageBox.warning(self, "No Files", "Please add at least one raw Excel file to combine.")
            return None
            
        for row in range(self.file_table.rowCount()):
            if self.file_table.cellWidget(row, 1) is not None:
                QMessageBox.warning(self, "Still Loading", "Please wait for all files to finish loading their sheet and material data.")
                return None
            
        if self.radio_append_file.isChecked():
            out_path = self.output_path_input.text().strip()
            if not out_path or not os.path.exists(out_path):
                QMessageBox.warning(self, "Invalid Target File", "Please select a valid existing combined file to append to.")
                return None

        file_list = []
        for row in range(self.file_table.rowCount()):
            f_path = self.file_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
            sheet_n = self.file_table.item(row, 1).text().strip()
            
            ref_widget = self.file_table.cellWidget(row, 2)
            row_refs = []
            if isinstance(ref_widget, QListWidget):
                row_refs = [ref_widget.item(i).text() for i in range(ref_widget.count())]
                
            samp_widget = self.file_table.cellWidget(row, 3)
            row_samps = []
            if isinstance(samp_widget, QListWidget):
                row_samps = [samp_widget.item(i).text() for i in range(samp_widget.count())]
                
            file_list.append({
                "path": f_path,
                "sheet": sheet_n,
                "references": row_refs,
                "samples": row_samps
            })
            
        out_path = self.output_path_input.text().strip()

        if not out_path:
            QMessageBox.warning(self, "No Output File",
                                "Please choose where to save the combined file.")
            return None

        # The combined workbook is written to out_path at the very END of the
        # run. If that path is also one of the inputs, the raw file would be
        # replaced by the combined output.
        clashes = [os.path.basename(f["path"]) for f in file_list
                   if f.get("path") and os.path.abspath(f["path"]) == os.path.abspath(out_path)]
        if clashes:
            QMessageBox.warning(self, "Output Would Overwrite an Input",
                                f"The output file is also one of the files being processed "
                                f"({clashes[0]}).\n\nChoose a different output file so the raw "
                                f"data is not overwritten.")
            return None

        # Check the destination folder now rather than discovering it is missing
        # after every file has already been processed.
        out_dir = os.path.dirname(os.path.abspath(out_path))
        if not os.path.isdir(out_dir):
            QMessageBox.warning(self, "Output Folder Not Found",
                                f"The folder for the output file does not exist:\n{out_dir}")
            return None

        return {
            "mode": "water" if self.btn_water.isChecked() else "carbonate",
            "protect_originals": self.radio_temp_copy.isChecked(),
            "file_list": file_list,
            "output_path": self.output_path_input.text().strip(),
            "open_on_complete": self.open_checkbox.isChecked(),
            "append_mode": self.radio_append_file.isChecked(),
            "calc_yield": self.chk_yield.isChecked(),
            "calc_co2": self.chk_co2.isChecked()
        }

    def open_combined_file(self):
        path = self.output_path_input.text().strip()
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "File Not Found", "The combined file has not been created yet or the path is invalid.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))
    
    def _update_footer_visibility(self):
        if self.file_table.rowCount() == 0:
            self.footer_hint.hide()
        else:
            self.footer_hint.show()
    
# =========================================================================
# ALL-IN-ONE COMBINE WORKER PROXY
# =========================================================================
class CombineWorker(QThread):
    log = pyqtSignal(str, str)
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    stopped_early = pyqtSignal()

    def __init__(self, params):
        super().__init__()
        self.params = params
        self._worker = None

    def stop(self):
        if self._worker:
            self._worker.stop()

    def run(self):
        mode = self.params.get("mode")
        
        if mode == "water":
            self._worker = WaterCombineWorker(self.params)
        else:
            self._worker = CarbonateCombineWorker(self.params)

        self._worker.log.connect(self.log.emit)
        self._worker.progress.connect(self.progress.emit)
        self._worker.finished.connect(self.finished.emit)
        self._worker.error.connect(self.error.emit)
        self._worker.stopped_early.connect(self.stopped_early.emit)

        self._worker.run()