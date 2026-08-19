from PyQt6.QtGui import QFont, QIcon, QCursor, QPainter, QColor, QPen, QAction, QKeySequence, QPixmap, QImage, QDesktopServices, QDoubleValidator
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QTabWidget, QTextEdit, QCheckBox,
    QLineEdit, QComboBox, QGroupBox, QMessageBox, QMenu, QProgressBar, QFrame,
    QSizePolicy, QSpacerItem, QGridLayout, QTabBar, QDialog, QScrollArea, QButtonGroup, 
    QRadioButton, QListWidget, QAbstractItemView, QTableWidget, QTableWidgetItem, QHeaderView, QLayout,
    QToolTip, QStyleOptionGroupBox, QProgressDialog, QLabel, QStyle, QDoubleSpinBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QPoint, QRect, QSize, QPropertyAnimation, QEasingCurve, QByteArray, QUrl

import utils.settings as settings
import gui.main_window as main_window

class ManageSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Settings")
        self.setFixedSize(400, 300)
        
        layout = QVBoxLayout(self)
        
        # --- Auto-Save Table Section ---
        layout.addWidget(QLabel("<b>Recent Auto-Saves:</b><br><small>(Created automatically when leaving the settings tab)</small>"))
        
        self.table = QTableWidget(0, 1)
        self.table.setHorizontalHeaderLabels(["Timestamp"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        
        self.btn_restore = QPushButton("Restore Selected Save")
        self.btn_restore.setEnabled(False) 
        self.btn_restore.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.table.itemSelectionChanged.connect(lambda: self.btn_restore.setEnabled(bool(self.table.selectedItems())))
        self.btn_restore.clicked.connect(self._do_restore)
        
        self._load_table()
        
        layout.addWidget(self.table)
        layout.addWidget(self.btn_restore)
        
        # Divider
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(divider)
        
        # --- Import / Export Section ---
        layout.addWidget(QLabel("<b>File Operations:</b>"))
        
        file_layout = QHBoxLayout()
        self.btn_export = QPushButton("Export Settings...")
        self.btn_import = QPushButton("Import Settings...")
        
        self.btn_export.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_import.setCursor(Qt.CursorShape.PointingHandCursor)
        
        file_layout.addWidget(self.btn_export)
        file_layout.addWidget(self.btn_import)
        
        self.btn_export.clicked.connect(self._do_export)
        self.btn_import.clicked.connect(self._do_import)
        
        layout.addLayout(file_layout)
        
    def _load_table(self):
        """Populates the table with available auto-saves and adds hover tooltips."""
        saves = settings.get_auto_saves()
        self.table.setRowCount(len(saves))
        for row, timestamp in enumerate(saves):
            item = QTableWidgetItem(timestamp)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            
            try:
                save_data = settings.get_auto_save_data(row) 
                
                if save_data and isinstance(save_data, dict):
                    tooltip_lines = [f"<b>Preview for {timestamp}:</b><hr>"]
                    for key, value in save_data.items():
                        
                        if key == "REFERENCE_MATERIALS" and isinstance(value, dict):
                            tooltip_lines.append(f"<b>{key}:</b>")
                            for mat_type, mat_list in value.items():
                                names = [m.get("col_c", "Unknown") for m in mat_list if m.get("col_c")]
                                tooltip_lines.append(f"&nbsp;&nbsp;• <i>{mat_type}:</i> {', '.join(names)}")
                        
                        elif key == "SLOPE_INTERCEPT_GROUPS" and isinstance(value, dict):
                            tooltip_lines.append(f"<b>{key}:</b>")
                            for mat_type, group_list in value.items():
                                group_strs = [f"({', '.join(g)})" for g in group_list]
                                tooltip_lines.append(f"&nbsp;&nbsp;• <i>{mat_type}:</i> {', '.join(group_strs)}")
                        
                        elif isinstance(value, list):
                            tooltip_lines.append(f"<b>{key}:</b> <i>[{len(value)} items]</i>")
                        elif isinstance(value, dict):
                            tooltip_lines.append(f"<b>{key}:</b> <i>[{len(value)} properties]</i>")
                        else:
                            tooltip_lines.append(f"<b>{key}:</b> {value}")
                    
                    item.setToolTip("<br>".join(tooltip_lines))
                else:
                    item.setToolTip("<i>No preview data available for this save.</i>")
            except AttributeError:
                item.setToolTip("<i>Preview unavailable.<br>Add get_auto_save_data(row) to settings.py.</i>")
            
            self.table.setItem(row, 0, item)
            
    def _do_restore(self):
        selected = self.table.selectedItems()
        if not selected:
            return
            
        row = selected[0].row()
        if settings.restore_auto_save(row):
            QMessageBox.information(self, "Restored", "Settings successfully restored from auto-save!")
            self.accept()
        
    def _do_export(self):
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Export Settings", "MRSI_Settings_Backup.json", "JSON Files (*.json)"
        )
        if filepath:
            success, msg = settings.export_to_file(filepath)
            if success:
                QMessageBox.information(self, "Export", msg)
            else:
                QMessageBox.critical(self, "Export Error", msg)
                
    def _do_import(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Import Settings", "", "JSON Files (*.json)"
        )
        if filepath:
            success, msg = settings.import_from_file(filepath)
            if success:
                QMessageBox.information(self, "Import", msg)
                self.accept()
            else:
                QMessageBox.critical(self, "Import Error", msg)


class YieldTabWidget(QWidget):
    """
    A widget that mirrors the layout of MaterialTypeWidget but is 
    customized for Yield (Read-only Atomic Weights table + Dynamic Slope Groups).
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.slope_widgets = []
        self._loading = False 
        
        self.layout = QVBoxLayout(self)
        self._create_table_section()
        self.layout.addSpacing(10)
        self._create_slope_section()
        self.load_data()

    def _create_table_section(self):
        grp = QGroupBox("Yield Compound Selection")
        l = QVBoxLayout()
        grp.setLayout(l)
        
        lbl_info = QLabel("<small>Select which compound is used for the Reference Material and which for Samples.</small>")
        lbl_info.setStyleSheet("color: gray;")
        l.addWidget(lbl_info)
        
        # Create a visually appealing grid layout
        table_layout = QGridLayout()
        table_layout.setSpacing(10)
        
        # Headers
        table_layout.addWidget(QLabel("<b>Compound</b>"), 0, 0)
        table_layout.addWidget(QLabel("<b>Reference Material</b>"), 0, 1, alignment=Qt.AlignmentFlag.AlignCenter)
        table_layout.addWidget(QLabel("<b>Sample</b>"), 0, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # Add a subtle divider line under headers
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        table_layout.addWidget(line, 1, 0, 1, 3)
        
        self.compounds_list = ["CaCO3", "MgCO3", "CaMg(CO3)2", "Li2CO3", "MnCO3", "FeCO3"]
        
        # Group the radio buttons so only one can be checked per column
        self.ref_bg = QButtonGroup(self)
        self.samp_bg = QButtonGroup(self)
        
        for i, comp in enumerate(self.compounds_list):
            row_idx = i + 2
            
            # Label
            comp_lbl = QLabel(comp)
            comp_lbl.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            table_layout.addWidget(comp_lbl, row_idx, 0)
            
            # Reference Radio
            rb_ref = QRadioButton()
            rb_ref.setCursor(Qt.CursorShape.PointingHandCursor)
            self.ref_bg.addButton(rb_ref, i)
            table_layout.addWidget(rb_ref, row_idx, 1, alignment=Qt.AlignmentFlag.AlignCenter)
            
            # Sample Radio
            rb_samp = QRadioButton()
            rb_samp.setCursor(Qt.CursorShape.PointingHandCursor)
            self.samp_bg.addButton(rb_samp, i)
            table_layout.addWidget(rb_samp, row_idx, 2, alignment=Qt.AlignmentFlag.AlignCenter)

        self.ref_bg.idClicked.connect(self._save_compounds)
        self.samp_bg.idClicked.connect(self._save_compounds)
        
        l.addLayout(table_layout)
        l.addStretch()
        self.layout.addWidget(grp)

    def load_data(self):
        self._loading = True
        
        # Fetch existing setting or set to defaults
        yield_compounds = settings.get_setting("YIELD_COMPOUNDS") or {"ref": "CaCO3", "samp": "MnCO3"}
        
        ref_comp = yield_compounds.get("ref", "CaCO3")
        samp_comp = yield_compounds.get("samp", "MnCO3")
        
        if ref_comp in self.compounds_list:
            self.ref_bg.button(self.compounds_list.index(ref_comp)).setChecked(True)
        if samp_comp in self.compounds_list:
            self.samp_bg.button(self.compounds_list.index(samp_comp)).setChecked(True)
            
        self.refresh_slope_ui()
        self._loading = False
        
    def _save_compounds(self, btn_id=None):
        """Saves the selected radio buttons to the settings memory."""
        if self._loading: return
        
        ref_idx = self.ref_bg.checkedId()
        samp_idx = self.samp_bg.checkedId()
        
        if ref_idx != -1 and samp_idx != -1:
            compounds = {
                "ref": self.compounds_list[ref_idx],
                "samp": self.compounds_list[samp_idx]
            }
            settings.set_setting("YIELD_COMPOUNDS", compounds)

    def _create_slope_section(self):
        grp = QGroupBox("Yield Slope and Intercept Groups")
        l = QVBoxLayout()
        grp.setLayout(l)
        
        lbl_slopes = QLabel("<small>Reference Materials detected from Carbonate tab</small>")
        lbl_slopes.setStyleSheet("color: gray;")
        l.addWidget(lbl_slopes)
        
        self.slope_container = QWidget()
        self.slope_layout = QVBoxLayout(self.slope_container)
        self.slope_layout.setContentsMargins(0, 0, 0, 0)
        l.addWidget(self.slope_container)
        
        self.add_slope_btn = QPushButton("Add Normalization Group")
        self.add_slope_btn.clicked.connect(self.add_slope_group)
        l.addWidget(self.add_slope_btn)
        
        self.layout.addWidget(grp)

    def refresh_slope_ui(self):
        for i in reversed(range(self.slope_layout.count())): 
            w = self.slope_layout.itemAt(i).widget()
            if w: w.setParent(None)
        self.slope_widgets.clear()
        
        slope_groups = settings.get_setting("SLOPE_INTERCEPT_GROUPS", sub_key="Yield") or []
        available = settings.get_reference_names("Carbonate")
        
        if slope_groups:
            for i, group in enumerate(slope_groups):
                # Using the exact SlopeGroupWidget class you provided
                w = main_window.SlopeGroupWidget(i, available, group, parent_widget=self)
                self.slope_layout.addWidget(w)
                self.slope_widgets.append(w)

    def add_slope_group(self):
        available = settings.get_reference_names("Carbonate")
        idx = len(self.slope_widgets)
        w = main_window.SlopeGroupWidget(idx, available, [], parent_widget=self)
        self.slope_layout.addWidget(w)
        self.slope_widgets.append(w)
        self.save_slope_config()

    def remove_slope_group_widget(self, widget):
        self.slope_layout.removeWidget(widget)
        widget.deleteLater()
        if widget in self.slope_widgets:
            self.slope_widgets.remove(widget)
        self.save_slope_config()

    def save_slope_config(self):
        if self._loading: return
        new_config = []
        for w in self.slope_widgets:
            sel = w.get_selected_materials()
            if sel: new_config.append(sel)
        settings.set_setting("SLOPE_INTERCEPT_GROUPS", new_config, sub_key="Yield")


class AdvancedSettingsTab(QWidget):
    def __init__(self):
        super().__init__()
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        
        self.scroll_area = QScrollArea()
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff) 
        
        self.scroll_content = QWidget()
        self.layout = QVBoxLayout(self.scroll_content)
        self.layout.setContentsMargins(10, 10, 10, 10)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        self._create_ui()
        
        self.scroll_area.setWidget(self.scroll_content)
        self.main_layout.addWidget(self.scroll_area)
        
    def hideEvent(self, event):
        """Automatically saves settings to memory when the user navigates away from this tab."""
        settings.auto_save()
        super().hideEvent(event)

    def _create_ui(self):
        # 1. Reset Button
        self.btn_reset = QPushButton("Reset to Default", self) 
        self.btn_reset.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_reset.setStyleSheet("""
            QPushButton {
                background-color: #f3f3f3;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 3px 5px;
                color: #333;
                font-weight: bold;
                margin-right: 10px;
            }
            QPushButton:hover {
                background-color: #e5e5e5;
            }
        """)
        self.btn_reset.clicked.connect(self._reset_to_default)
        
        # 2. Manage Settings Button
        self.btn_manage = QPushButton("Manage Settings", self)
        self.btn_manage.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_manage.setStyleSheet(self.btn_reset.styleSheet())
        self.btn_manage.clicked.connect(self._open_manage_settings)
        
        # Explicitly show the buttons
        self.btn_reset.show()
        self.btn_manage.show()

        # Construct the rest of the UI Layouts
        self._create_general_config()
        self._create_outlier_settings()
        self._create_calc_logic_section()
        self._create_material_tabs()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'btn_reset') and self.btn_reset.isVisible():
            self.btn_reset.adjustSize()
            
            x_reset = self.width() - self.btn_reset.width() - 25
            y = 15
            
            self.btn_reset.move(x_reset, y)
            self.btn_reset.raise_()
            
            if hasattr(self, 'btn_manage') and self.btn_manage.isVisible():
                self.btn_manage.adjustSize()
                x_manage = x_reset - self.btn_manage.width() - 10
                self.btn_manage.move(x_manage, y)
                self.btn_manage.raise_()
        
    def _create_general_config(self):
        group = QGroupBox("1. Conditional Formatting for Excel")
        layout = QVBoxLayout() 
        group.setLayout(layout)
        
        def create_info_label(tooltip_text):
            lbl = main_window.InstantTooltipLabel("ⓘ") 
            lbl.setCursor(Qt.CursorShape.WhatsThisCursor)
            lbl.setToolTip(tooltip_text)
            lbl.setStyleSheet("""
                QLabel {
                    color: #555;
                    font-size: 14px;
                    font-weight: bold;
                    margin-left: 2px;
                    margin-right: 5px;
                }
                QLabel:hover {
                    color: #0078d7; 
                }
            """)
            return lbl

        visual_layout = QHBoxLayout()
        visual_layout.setContentsMargins(0, 0, 0, 0)
        
        self.lbl_visual_good = QLabel()
        self.lbl_visual_good.setFixedSize(45, 25)
        self.lbl_visual_good.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_visual_good.setStyleSheet("""
            background-color: #FFFFFF; 
            color: #000000; 
            border: 1px solid #D4D4D4; 
            font-family: 'Segoe UI', sans-serif;
            font-size: 11px;
        """)
        
        lbl_arrow = QLabel("➔")
        lbl_arrow.setStyleSheet("color: #888; font-size: 16px; font-weight: bold;")
        lbl_arrow.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.lbl_visual_bad = QLabel()
        self.lbl_visual_bad.setFixedSize(45, 25)
        self.lbl_visual_bad.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_visual_bad.setStyleSheet("""
            background-color: #FFC7CE; 
            color: #9C0006; 
            border: 1px solid #FFC7CE;
            font-family: 'Segoe UI', sans-serif;
            font-size: 11px;
            font-weight: bold;
        """)
        
        visual_layout.addWidget(QLabel("<small style='color: gray;'><i>Example:</i></small>"))
        visual_layout.addWidget(self.lbl_visual_good)
        visual_layout.addWidget(lbl_arrow)
        visual_layout.addWidget(self.lbl_visual_bad)
        visual_layout.addStretch()
        layout.addLayout(visual_layout)

        toggle_layout = QHBoxLayout()
        self.chk_stdev = QCheckBox()
        
        is_enabled = settings.get_setting("STDEV_THRESHOLD_ENABLED")
        is_enabled_bool = is_enabled if is_enabled is not None else False
        self.chk_stdev.setChecked(is_enabled_bool)
        self.chk_stdev.setText("Enabled" if is_enabled_bool else "Disabled")
        self.chk_stdev.setStyleSheet("font-weight: normal;" if is_enabled_bool else "font-weight: bold;")
        self.chk_stdev.stateChanged.connect(self._on_stdev_toggled)
        
        toggle_layout.addWidget(self.chk_stdev)
        toggle_layout.addStretch()
        layout.addLayout(toggle_layout)

        row1 = QHBoxLayout()
        lbl_layout = QVBoxLayout()
        lbl_layout.setSpacing(0)
        lbl_layout.addWidget(QLabel("<b>Stdev Threshold</b>"))
        lbl_layout.addWidget(QLabel("<small style='color: gray;'>(All Steps)</small>"))
        
        row1.addLayout(lbl_layout)
        row1.addWidget(create_info_label(
            "<b>Standard Deviation Limit</b><br>"
            "Defines the cutoff value for the standard deviation.<br>"
            "Any value above this limit will be highlighted <span style='color:red;'>red</span> in the stdev columns."
        ))
        row1.addWidget(QLabel(":"))

        self.input_stdev = QLineEdit()
        self.input_stdev.setFixedWidth(55) 
        self.input_stdev.setStyleSheet("""
            QLineEdit:disabled {
                background-color: #EAEAEA;
                color: #B0B0B0;
                border: 1px solid #D3D3D3;
            }
        """)
        
        validator = QDoubleValidator(0.0, 100.0, 99, self)
        validator.setNotation(QDoubleValidator.Notation.StandardNotation)
        self.input_stdev.setValidator(validator)
        
        current_stdev = float(settings.get_setting("STDEV_THRESHOLD") or 0.8)
        self.input_stdev.setText(f"{current_stdev:g}")
        self.input_stdev.editingFinished.connect(self._on_stdev_changed)
        self.input_stdev.textChanged.connect(self._on_text_changed_for_visual)
        row1.addWidget(self.input_stdev)
        
        self.btn_up = QPushButton("▲")
        self.btn_down = QPushButton("▼")
        
        for btn in [self.btn_up, self.btn_down]:
            btn.setFixedSize(20, 13)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet("""
                QPushButton { 
                    background-color: transparent; 
                    color: #888; 
                    border: 1px solid #888; 
                    border-radius: 2px; 
                    font-size: 8px; 
                    padding: 0px;
                }
                QPushButton:hover { background-color: #888; color: white; }
                QPushButton:disabled { border: 1px solid #D3D3D3; color: #D3D3D3; }
            """)
            
        self.btn_up.clicked.connect(self._step_up_stdev)
        self.btn_down.clicked.connect(self._step_down_stdev)
        
        spin_btn_layout = QVBoxLayout()
        spin_btn_layout.setSpacing(2)
        spin_btn_layout.setContentsMargins(0, 0, 0, 0)
        spin_btn_layout.addWidget(self.btn_up)
        spin_btn_layout.addWidget(self.btn_down)
        
        row1.addLayout(spin_btn_layout)
        row1.addStretch()
        
        self._update_stdev_state(self.chk_stdev.isChecked())
        layout.addLayout(row1)
        self.layout.addWidget(group)
        self._update_visual_example(current_stdev)

    def _update_visual_example(self, current_limit):
        self.lbl_visual_good.setText(f"{current_limit:.3f}")
        self.lbl_visual_bad.setText(f"{current_limit:.3f}")

    def _on_stdev_toggled(self):
        is_enabled = self.chk_stdev.isChecked()
        self.chk_stdev.setText("Enabled" if is_enabled else "Disabled")
        self.chk_stdev.setStyleSheet("font-weight: normal;" if is_enabled else "font-weight: bold;")
        settings.set_setting("STDEV_THRESHOLD_ENABLED", is_enabled)
        self._update_stdev_state(is_enabled)
        
    def _update_stdev_state(self, is_enabled):
        self.input_stdev.setEnabled(is_enabled)
        self.btn_up.setEnabled(is_enabled)
        self.btn_down.setEnabled(is_enabled)

    def _create_outlier_settings(self):
        group = QGroupBox("2. Outlier Settings")
        layout = QVBoxLayout() 
        group.setLayout(layout)
        
        example_layout = QHBoxLayout()
        example_layout.setContentsMargins(0, 0, 0, 5)
        
        lbl_example_text = QLabel("<small style='color: gray;'><i>Example:</i></small>")
        lbl_outlier_cell = QLabel("<s>4.020</s>")
        lbl_outlier_cell.setFixedSize(45, 22)
        lbl_outlier_cell.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_outlier_cell.setStyleSheet("""
            background-color: #FFFFFF; 
            color: #FF0000; 
            border: 1px solid #D4D4D4; 
            font-family: 'Segoe UI', sans-serif;
            font-size: 11px;
        """)
        
        example_layout.addWidget(lbl_example_text)
        example_layout.addWidget(lbl_outlier_cell)
        example_layout.addStretch()
        layout.addLayout(example_layout)
        
        def create_info_label(tooltip_text):
            lbl = main_window.InstantTooltipLabel("ⓘ") 
            lbl.setCursor(Qt.CursorShape.WhatsThisCursor)
            lbl.setToolTip(tooltip_text)
            lbl.setStyleSheet("""
                QLabel { color: #555; font-size: 14px; font-weight: bold; margin-left: 2px; margin-right: 5px; }
                QLabel:hover { color: #0078d7; }
            """)
            return lbl

        row1 = QHBoxLayout()
        lbl_layout1 = QVBoxLayout()
        lbl_layout1.setSpacing(0)
        lbl_layout1.addWidget(QLabel("<b>Outlier Calculation</b>"))
        lbl_layout1.addWidget(QLabel("<small style='color: gray;'>(Steps: Data, Group, Normalization)</small>"))
        
        row1.addLayout(lbl_layout1)
        row1.addWidget(create_info_label(
            "<b>Sigma Threshold (Standard Deviations)</b><br>"
            "Determines how strict the outlier detection is.<br>"
            "<ul><li><b>1σ:</b> avg +- std</li><li><b>2σ:</b> avg +- 2*std</li><li><b>3σ:</b> avg +- 3*std</li></ul>"
        ))
        row1.addWidget(QLabel(":"))

        self.bg_sigma = QButtonGroup(self)
        self.rb_1sigma = QRadioButton("1σ")
        self.rb_2sigma = QRadioButton("2σ")
        self.rb_2sigma.setStyleSheet("font-weight: bold;")
        self.rb_3sigma = QRadioButton("3σ")
        
        self.bg_sigma.addButton(self.rb_1sigma, 1)
        self.bg_sigma.addButton(self.rb_2sigma, 2)
        self.bg_sigma.addButton(self.rb_3sigma, 3)
        
        row1.addWidget(self.rb_1sigma)
        row1.addWidget(self.rb_2sigma)
        row1.addWidget(self.rb_3sigma)
        row1.addStretch()
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        lbl_layout2 = QVBoxLayout()
        lbl_layout2.setSpacing(0)
        lbl_layout2.addWidget(QLabel("<b>Exclusion Logic</b>"))
        lbl_layout2.addWidget(QLabel("<small style='color: gray;'>(Steps: Data, Group, Normalization)</small>"))
        
        row2.addLayout(lbl_layout2)
        row2.addWidget(create_info_label(
            "<b>How to Handle Outliers</b><br>"
            "<ul><li><b>Individual:</b> If Carbon (δ13C) is an outlier but Oxygen (δ18O) is good, keep the Oxygen value.</li>"
            "<li><b>Exclude Row:</b> If <i>either</i> value is an outlier, discard the entire measurement row.</li></ul>"
        ))
        row2.addWidget(QLabel(":"))

        self.bg_excl = QButtonGroup(self)
        self.rb_excl_row = QRadioButton("Exclude Entire Row")
        self.rb_excl_ind = QRadioButton("Individual (Keep Valid C or O)")
        self.rb_excl_ind.setStyleSheet("font-weight: bold;") 
        
        self.bg_excl.addButton(self.rb_excl_ind)
        self.bg_excl.addButton(self.rb_excl_row)
        
        row2.addWidget(self.rb_excl_ind)
        row2.addWidget(self.rb_excl_row)
        row2.addStretch()
        layout.addLayout(row2)

        curr_sigma = settings.get_setting("OUTLIER_SIGMA") or 2
        if curr_sigma == 1: self.rb_1sigma.setChecked(True)
        elif curr_sigma == 3: self.rb_3sigma.setChecked(True)
        else: self.rb_2sigma.setChecked(True)
        
        curr_excl = settings.get_setting("OUTLIER_EXCLUSION_MODE") or "Individual"
        if curr_excl == "Exclude Row": self.rb_excl_row.setChecked(True)
        else: self.rb_excl_ind.setChecked(True)

        self.bg_sigma.idClicked.connect(self._on_sigma_changed)
        self.bg_excl.buttonToggled.connect(self._on_excl_mode_changed)
        self.layout.addWidget(group)

    def _create_calc_logic_section(self):
        group = QGroupBox("3. Data Selection")
        layout = QVBoxLayout()
        group.setLayout(layout)
        
        def create_info_label(tooltip_text):
            lbl = main_window.InstantTooltipLabel("ⓘ") 
            lbl.setCursor(Qt.CursorShape.WhatsThisCursor)
            lbl.setToolTip(tooltip_text)
            lbl.setStyleSheet("""
                QLabel { color: #555; font-size: 14px; font-weight: bold; margin-left: 2px; margin-right: 5px; }
                QLabel:hover { color: #0078d7; }
            """)
            return lbl

        row1 = QHBoxLayout()
        lbl_layout1 = QVBoxLayout()
        lbl_layout1.setSpacing(0)
        lbl_layout1.addWidget(QLabel("<b>Measured 𝛅 values</b>"))
        lbl_layout1.addWidget(QLabel("<small style='color: gray;'>(Step 3: Last 6)</small>"))
        
        row1.addLayout(lbl_layout1)
        row1.addWidget(create_info_label(
            "<b>Calculation Mode for Step 3</b><br>"
            "Decides which data is used to calculate the 'Last 6' Averages.<br>"
            "<ul><li><b>Last 6:</b> Takes the raw average of the last 6 measurements.</li>"
            "<li><b>Last 6 Outliers Excluded:</b> Removes statistical outliers <i>before</i> calculating the average.</li></ul>"
        ))
        row1.addWidget(QLabel(":"))

        self.bg_step3 = QButtonGroup(self)
        self.rb_s3_last6 = QRadioButton("Last 6")
        self.rb_s3_last6.setStyleSheet("font-weight: bold;") 
        self.rb_s3_last6_excl = QRadioButton("Last 6 Outliers Excluded (See Section 2)")
        self.bg_step3.addButton(self.rb_s3_last6)
        self.bg_step3.addButton(self.rb_s3_last6_excl)
        row1.addWidget(self.rb_s3_last6)
        row1.addWidget(self.rb_s3_last6_excl)
        row1.addStretch()
        layout.addLayout(row1)
        
        row2 = QHBoxLayout()
        lbl_layout2 = QVBoxLayout()
        lbl_layout2.setSpacing(0)
        lbl_layout2.addWidget(QLabel("<b>Average for RM</b>"))
        lbl_layout2.addWidget(QLabel("<small style='color: gray;'>(Step 7: Normalization)</small>"))
        
        row2.addLayout(lbl_layout2)
        row2.addWidget(create_info_label(
            "<b>Normalization Calculation</b><br>"
            "Determines which data points are used to calculate the Average and Standard Deviation for the Reference Materials (RMs) during normalization.<br>"
            "<ul><li><b>All Values:</b> Computes the metrics using every measurement, including those flagged as outliers.</li>"
            "<li><b>Outliers Excluded:</b> Computes the metrics using only valid data points, ignoring any measurements flagged as outliers.</li></ul>"
        ))
        row2.addWidget(QLabel(":"))

        self.bg_step7 = QButtonGroup(self)
        self.rb_s7_all = QRadioButton("All Values")
        self.rb_s7_all.setStyleSheet("font-weight: bold;")
        self.rb_s7_outlier = QRadioButton("Outliers Excluded (See Section 2)")
        self.bg_step7.addButton(self.rb_s7_all)
        self.bg_step7.addButton(self.rb_s7_outlier)
        row2.addWidget(self.rb_s7_all)
        row2.addWidget(self.rb_s7_outlier)
        row2.addStretch()
        layout.addLayout(row2)
        
        if settings.get_setting("CALC_MODE_STEP3") == "Last 6 Outliers Excl.": self.rb_s3_last6_excl.setChecked(True)
        else: self.rb_s3_last6.setChecked(True)
            
        if settings.get_setting("CALC_MODE_STEP7") == "Outliers Excluded": self.rb_s7_outlier.setChecked(True)
        else: self.rb_s7_all.setChecked(True)

        self.bg_step3.buttonToggled.connect(self._on_calc_mode_changed)
        self.bg_step7.buttonToggled.connect(self._on_calc_mode_changed)
        self.layout.addWidget(group)

    def _create_material_tabs(self):
        self.tabs = QTabWidget()
        
        self.water_widget = main_window.MaterialTypeWidget("Water",
                                               ["Water Standards", "Col D", "Col E", "Col F (δ²H)", "Col G (δ¹⁸O SMOW)", "Col H", "Color"])
        self.tabs.addTab(self.water_widget, "Water")

        self.carb_widget = main_window.MaterialTypeWidget("Carbonate", 
                                              ["Col C (Name)", "Col D", "Col E", "Col F (d13C)", "Col G (d18O)", "Col H", "Color"])
        self.tabs.addTab(self.carb_widget, "Carbonate")

        # --- UPDATED CO2 Tab Setup ---
        self.co2_widget = main_window.MaterialTypeWidget("CO2", 
                                              ["Name", "", "", "d13C", "d18O CO2", "", "Color"])
        
        co2_radio_layout = QHBoxLayout()
        co2_radio_layout.setContentsMargins(0, 5, 0, 5)
        co2_radio_layout.addWidget(QLabel("<b>CO2 Quick Switch:</b>"))
        
        self.co2_bg = QButtonGroup(self)
        self.rb_25c = QRadioButton("25 degrees C")
        self.rb_72c = QRadioButton("72 degrees C")
        self.rb_custom = QRadioButton("Custom")
        
        for rb in [self.rb_25c, self.rb_72c, self.rb_custom]:
            rb.setCursor(Qt.CursorShape.PointingHandCursor)
            
        self.co2_bg.addButton(self.rb_25c, 1)
        self.co2_bg.addButton(self.rb_72c, 2)
        self.co2_bg.addButton(self.rb_custom, 3)
        
        co2_radio_layout.addWidget(self.rb_25c)
        co2_radio_layout.addWidget(self.rb_72c)
        co2_radio_layout.addWidget(self.rb_custom)
        co2_radio_layout.addStretch()
        
        self.co2_bg.idClicked.connect(self._apply_co2_profile)
        
        # Insert exactly above the Add/Remove buttons using the placeholder we created
        self.co2_widget.extra_controls_layout.addLayout(co2_radio_layout)
        
        # Listen for manual table edits so we can auto-switch to "Custom"
        self.co2_widget.table.itemChanged.connect(self._check_co2_profile_state)
        
        self.tabs.addTab(self.co2_widget, "CO2")
        
        # --- NEW Yield Tab Setup ---
        self.yield_widget = YieldTabWidget()
        self.tabs.addTab(self.yield_widget, "Yield")
        
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.layout.addWidget(self.tabs)

        # Run the detector once on startup so the correct radio button is highlighted
        self._check_co2_profile_state()
        
    def _on_tab_changed(self, index):
        if self.tabs.tabText(index) == "Yield":
            self.yield_widget.load_data()

    def _step_up_stdev(self):
        try: val = float(self.input_stdev.text() or 0.0)
        except ValueError: val = 0.0
        new_val = val + 0.01
        self.input_stdev.setText(f"{new_val:g}")
        self._on_stdev_changed()

    def _step_down_stdev(self):
        try: val = float(self.input_stdev.text() or 0.0)
        except ValueError: val = 0.0
        new_val = max(0.0, val - 0.01)
        self.input_stdev.setText(f"{new_val:g}")
        self._on_stdev_changed()

    def _on_text_changed_for_visual(self, text):
        try: val = float(text) if text else 0.0
        except ValueError: val = 0.0
        self._update_visual_example(val)

    def _on_stdev_changed(self):
        try: val = float(self.input_stdev.text() or 0.0)
        except ValueError: val = 0.0
        settings.set_setting("STDEV_THRESHOLD", val)

    def _on_calc_mode_changed(self, btn, checked):
        if not checked: return
        val3 = "Last 6 Outliers Excl." if self.rb_s3_last6_excl.isChecked() else "Last 6"
        settings.set_setting("CALC_MODE_STEP3", val3)
        val7 = "Outliers Excluded" if self.rb_s7_outlier.isChecked() else "All Values"
        settings.set_setting("CALC_MODE_STEP7", val7)
    
    def _on_sigma_changed(self, btn_id):
        settings.set_setting("OUTLIER_SIGMA", btn_id)
        
    def _on_excl_mode_changed(self, btn, checked):
        if not checked: return
        mode = "Individual" if self.rb_excl_ind.isChecked() else "Exclude Row"
        settings.set_setting("OUTLIER_EXCLUSION_MODE", mode)

    def _reset_to_default(self):
        self.chk_stdev.setChecked(False)
        self.input_stdev.setText("0.08") 
        self._on_stdev_changed()         
        
        self.rb_2sigma.setChecked(True)
        self.rb_excl_ind.setChecked(True)
        
        self.rb_s3_last6.setChecked(True)
        self.rb_s7_all.setChecked(True)
        
        default_carb_mats = [
            {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "-2.37", "col_h": "", "color": "green"},
            {"col_c": "LSVEC",    "col_d": "", "col_e": "", "col_f": "-46.6", "col_g": "-26.7", "col_h": "", "color": "lightblue"},
            {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-23.01", "col_h": "", "color": "red"},
            {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "-2.20",  "col_h": "", "color": "darkblue"}
        ]
        default_carb_slopes = [
            ["NBS 18", "NBS 19"],
            ["NBS 18", "NBS 19", "IAEA 603"]
        ]
        settings.set_setting("REFERENCE_MATERIALS", default_carb_mats, sub_key="Carbonate")
        settings.set_setting("SLOPE_INTERCEPT_GROUPS", default_carb_slopes, sub_key="Carbonate")

        default_water_mats = [
            {"col_c": "MRSI-STD-W1", "col_d": "", "col_e": "", "col_f": "-3.52", "col_g": "-0.58", "col_h": "", "color": "red"},
            {"col_c": "MRSI-STD-W2",  "col_d": "", "col_e": "", "col_f": "-214.79", "col_g": "-28.08", "col_h": "", "color": "darkblue"},
            {"col_c": "USGS W-67400",  "col_d": "", "col_e": "", "col_f": "1.25", "col_g": "-1.97", "col_h": "", "color": "orange"},
            {"col_c": "USGS W-64444",  "col_d": "", "col_e": "", "col_f": "-399.1", "col_g": "-51.14", "col_h": "", "color": "green"}
        ]
        default_water_slopes = [
            ["MRSI-STD-W1", "MRSI-STD-W2"],
            ["USGS W-67400", "USGS W-64444"]
        ]
        settings.set_setting("REFERENCE_MATERIALS", default_water_mats, sub_key="Water")
        settings.set_setting("SLOPE_INTERCEPT_GROUPS", default_water_slopes, sub_key="Water")

        default_co2_mats = [
            {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "7.86", "col_h": "", "color": "green"},
            {"col_c": "LSVEC",    "col_d": "", "col_e": "", "col_f": "-46.6", "col_g": "", "col_h": "", "color": "lightblue"},
            {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-13.00", "col_h": "", "color": "red"},
            {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "8.03",  "col_h": "", "color": "darkblue"}
        ]
        settings.set_setting("REFERENCE_MATERIALS", default_co2_mats, sub_key="CO2")
        settings.set_setting("SLOPE_INTERCEPT_GROUPS", default_carb_slopes, sub_key="CO2")

        # Yield default configuration matches Carbonate initially
        default_yield_slopes = [
            ["NBS 18", "NBS 19"],
            ["Carrara"]
        ]
        settings.set_setting("SLOPE_INTERCEPT_GROUPS", default_yield_slopes, sub_key="Yield")
        
        settings.set_setting("YIELD_COMPOUNDS", {"ref": "CaCO3", "samp": "MnCO3"})

        self.carb_widget.load_data()
        self.water_widget.load_data()
        self.co2_widget.load_data() 
        self.yield_widget.load_data()
        
        self._check_co2_profile_state()

    def _open_manage_settings(self):
        """Opens the dialog and refreshes the UI if settings were imported/loaded."""
        dialog = ManageSettingsDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._refresh_ui_from_settings()

    def _refresh_ui_from_settings(self):
        """Pulls the current states from settings.py and forces the UI to match."""
        # 1. Stdev Configuration
        is_enabled = settings.get_setting("STDEV_THRESHOLD_ENABLED")
        self.chk_stdev.setChecked(bool(is_enabled))
        
        stdev_val = float(settings.get_setting("STDEV_THRESHOLD") or 0.08)
        self.input_stdev.setText(f"{stdev_val:g}")
        self._update_visual_example(stdev_val)

        # 2. Outlier Configuration
        sigma = settings.get_setting("OUTLIER_SIGMA")
        if sigma == 1: self.rb_1sigma.setChecked(True)
        elif sigma == 3: self.rb_3sigma.setChecked(True)
        else: self.rb_2sigma.setChecked(True)

        excl = settings.get_setting("OUTLIER_EXCLUSION_MODE")
        if excl == "Exclude Row": self.rb_excl_row.setChecked(True)
        else: self.rb_excl_ind.setChecked(True)

        # 3. Calculation Logic
        step3 = settings.get_setting("CALC_MODE_STEP3")
        if step3 == "Last 6 Outliers Excl.": self.rb_s3_last6_excl.setChecked(True)
        else: self.rb_s3_last6.setChecked(True)

        step7 = settings.get_setting("CALC_MODE_STEP7")
        if step7 == "Outliers Excluded": self.rb_s7_outlier.setChecked(True)
        else: self.rb_s7_all.setChecked(True)

        # 4. Refresh Material Tables
        self.carb_widget.load_data()
        self.water_widget.load_data()
        self.co2_widget.load_data()
        self.yield_widget.load_data()

        # Just call our new validation function to set the right radio button
        self._check_co2_profile_state()
    
    def _apply_co2_profile(self, btn_id):
        """Applies the selected predefined CO2 profile, or clears d18O if Custom is clicked."""
        self.co2_widget.table.blockSignals(True)
        
        if btn_id == 3:
            # Custom selected: clear all d18O CO2 (col_g) values
            settings.set_setting("CO2_TEMP_MODE", "Custom") # <--- SAVES TO SETTINGS
            current_co2 = settings.get_setting("REFERENCE_MATERIALS", sub_key="CO2") or []
            for mat in current_co2:
                mat["col_g"] = ""
            settings.set_setting("REFERENCE_MATERIALS", current_co2, sub_key="CO2")
            self.co2_widget.load_data()
            
        elif btn_id == 1:
            # 25 degrees C Profile
            settings.set_setting("CO2_TEMP_MODE", "25 °C") # <--- SAVES TO SETTINGS
            co2_mats = [
                {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "7.86", "col_h": "", "color": "green"},
                {"col_c": "LSVEC",    "col_d": "", "col_e": "", "col_f": "-46.6", "col_g": "", "col_h": "", "color": "lightblue"},
                {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-13.00", "col_h": "", "color": "red"},
                {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "8.03",  "col_h": "", "color": "darkblue"}
            ]
            settings.set_setting("REFERENCE_MATERIALS", co2_mats, sub_key="CO2")
            self.co2_widget.load_data()
            
        elif btn_id == 2:
            # 72 degrees C Profile
            settings.set_setting("CO2_TEMP_MODE", "72 °C") # <--- SAVES TO SETTINGS
            co2_mats = [
                {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "6.26", "col_h": "", "color": "green"},
                {"col_c": "LSVEC",    "col_d": "", "col_e": "", "col_f": "-46.6", "col_g": "", "col_h": "", "color": "lightblue"},
                {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-14.56", "col_h": "", "color": "red"},
                {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "6.43",  "col_h": "", "color": "darkblue"}
            ]
            settings.set_setting("REFERENCE_MATERIALS", co2_mats, sub_key="CO2")
            self.co2_widget.load_data()
            
        self.co2_widget.table.blockSignals(False)

    def _check_co2_profile_state(self, item=None):
        """Monitors the table directly. If values match 72°C or 25°C exactly, select them. Otherwise, Custom."""
        if self.co2_widget._loading: return
        
        # Read directly from the live table cells for instant detection
        vals = {}
        for r in range(self.co2_widget.table.rowCount()):
            name_item = self.co2_widget.table.item(r, 0)
            if name_item:
                mat_name = name_item.text().strip()
                if mat_name in ["IAEA 603", "NBS 18", "NBS 19"]:
                    # col_g (d18O CO2) is at column index 4
                    d18o_item = self.co2_widget.table.item(r, 4)
                    vals[mat_name] = d18o_item.text().strip() if d18o_item else ""
        
        self.co2_bg.blockSignals(True)
        
        # Check for 25°C Exact Match
        if vals.get("IAEA 603") == "7.86" and vals.get("NBS 18") == "-13.00" and vals.get("NBS 19") == "8.03":
            self.rb_25c.setChecked(True)
            settings.set_setting("CO2_TEMP_MODE", "25 °C") # <--- SAVES TO SETTINGS
        # Check for 72°C Exact Match
        elif vals.get("IAEA 603") == "6.26" and vals.get("NBS 18") == "-14.56" and vals.get("NBS 19") == "6.43":
            self.rb_72c.setChecked(True)
            settings.set_setting("CO2_TEMP_MODE", "72 °C") # <--- SAVES TO SETTINGS
        # Any deviation triggers Custom
        else:
            self.rb_custom.setChecked(True)
            settings.set_setting("CO2_TEMP_MODE", "Custom") # <--- SAVES TO SETTINGS
            
        self.co2_bg.blockSignals(False)