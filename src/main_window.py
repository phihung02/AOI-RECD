import os
from PyQt6.QtWidgets import (QMainWindow, QTableView, QVBoxLayout, 
                             QWidget, QHeaderView, QPushButton, QFileDialog, 
                             QLabel, QHBoxLayout, QSplitter, 
                             QScrollArea, QFormLayout, QFrame, QDateTimeEdit, 
                             QDialog, QMessageBox, QGraphicsDropShadowEffect,
                             QListWidget, QListWidgetItem,QComboBox)
from PyQt6.QtCore import Qt, QDateTime, QEvent, QSize
from PyQt6.QtGui import QIcon, QColor

from src.config import HEADERS
from src.widgets import (TagWidget, CheckableComboBox, ScanOptionDialog, 
                         ImageDelegate, RowPreviewDialog, ExportConfigDialog)
from src.data import (FastTableModel, DataLoaderThread, ExcelExportThread, 
                      ImageExportThread)

# ================= GIAO DIỆN CHÍNH =================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AOI Tech MI RECD v0.1.c - Retrieve electronic component data")
        self.setWindowIcon(QIcon("appicon.ico"))
        self.resize(1400, 700)

        self.image_folder_path = os.path.abspath(os.path.join(os.getcwd(), "./IMAGE/EXPORTEDIMAGES"))
        self.folder_path = None
        
        # Biến trạng thái quản lý cột (Tags)
        self.active_cols = HEADERS.copy()
        self.hidden_cols = []

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        top_bar = QHBoxLayout()
        self.btn_load = QPushButton("📁 Open")
        self.btn_load.clicked.connect(self.start_loading)
        
        self.btn_export = QPushButton("📊 Export")
        self.btn_export.setStyleSheet("background-color: #107c41; color: white; font-weight: bold;")
        self.btn_export.clicked.connect(self.show_export_dialog)

        self.btn_toggle_nav = QPushButton("Filter ◀")
        self.btn_toggle_nav.setCheckable(True)
        self.btn_toggle_nav.setChecked(True)
        self.btn_toggle_nav.clicked.connect(self.toggle_navigator)

        self.lbl_status = QLabel(f"Ready.")

        top_bar.addWidget(self.btn_load)
        top_bar.addWidget(self.btn_export)
        top_bar.addWidget(self.btn_toggle_nav)
        top_bar.addWidget(self.lbl_status, stretch=1)
        main_layout.addLayout(top_bar)

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.body = QVBoxLayout()
        self.footer = QLabel()
        self.footer.setText("I am Optimus Prime, and I send this message: our planet will always be remembered, for in our memories, we live on. ©2026 - Developer by Ex Tech Hung")
        self.body.addWidget(self.splitter, stretch=1)
        self.body.addWidget(self.footer, alignment=Qt.AlignmentFlag.AlignRight)
        
        main_layout.addLayout(self.body, stretch=1)

        self.nav_frame = QFrame()
        self.nav_frame.setStyleSheet(" border-right: 1px solid #ccc;")
        nav_layout = QVBoxLayout(self.nav_frame)
        
        # 1. KHU VỰC TAG QUẢN LÝ CỘT (MỚI)
        lbl_col = QLabel("SHOW DATA COLUMNS")
        lbl_col.setStyleSheet("font-weight: bold;")
        nav_layout.addWidget(lbl_col)

        self.combo_add_col = QComboBox()
        self.combo_add_col.addItem("-- Add Hidden Column --")
        self.combo_add_col.activated.connect(self.show_column)
        nav_layout.addWidget(self.combo_add_col)

        self.tag_list = QListWidget()
        self.tag_list.setFlow(QListWidget.Flow.LeftToRight)
        self.tag_list.setWrapping(True)
        self.tag_list.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.tag_list.setSpacing(2)
        self.tag_list.setStyleSheet("QListWidget { border: none;}")
        self.tag_list.setMaximumHeight(120)
        nav_layout.addWidget(self.tag_list)

        # 2. KHU VỰC LỌC THỜI GIAN
        lbl_date = QLabel("Filter by date")
        lbl_date.setStyleSheet("font-weight: bold; margin-top: 10px;")
        nav_layout.addWidget(lbl_date)
        
        self.dt_from = QDateTimeEdit()
        self.dt_from.setCalendarPopup(True)
        self.dt_from.setDisplayFormat("MM/dd/yyyy HH:mm:ss")
        self.dt_to = QDateTimeEdit()
        self.dt_to.setCalendarPopup(True)
        self.dt_to.setDisplayFormat("MM/dd/yyyy HH:mm:ss")
        
        nav_layout.addWidget(QLabel("From (A):"))
        nav_layout.addWidget(self.dt_from)
        nav_layout.addWidget(QLabel("To (B):"))
        nav_layout.addWidget(self.dt_to)
        
        self.btn_apply_date = QPushButton("✅ Apply")
        self.btn_apply_date.clicked.connect(self.apply_filters)
        nav_layout.addWidget(self.btn_apply_date)
        
        # 3. KHU VỰC LỌC THUỘC TÍNH
        nav_layout.addWidget(QLabel("\nFilter by attribute"), alignment=Qt.AlignmentFlag.AlignBottom)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        filter_widget = QWidget()
        self.filter_layout = QFormLayout(filter_widget)
        scroll_area.setWidget(filter_widget)
        nav_layout.addWidget(scroll_area)
        
        self.btn_clear_filter = QPushButton("Clear")
        self.btn_clear_filter.clicked.connect(self.clear_all_filters)
        nav_layout.addWidget(self.btn_clear_filter)

        self.table_view = QTableView()
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSortingEnabled(True)
        self.table_view.verticalHeader().setDefaultSectionSize(70) 
        
        self.table_view.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table_view.setSelectionMode(QTableView.SelectionMode.ExtendedSelection)

        self.model = FastTableModel()
        self.table_view.setModel(self.model)

        self.splitter.addWidget(self.nav_frame)
        self.splitter.addWidget(self.table_view)
        self.splitter.setSizes([300, 1100])
        self.filter_combos = {}

        self.image_delegate = ImageDelegate(self.image_folder_path, self.table_view)
        self.table_view.setItemDelegateForColumn(HEADERS.index("Review Image"), self.image_delegate)
        self.table_view.setItemDelegateForColumn(HEADERS.index("Target Image"), self.image_delegate)
        self.table_view.installEventFilter(self)
        self.table_view.doubleClicked.connect(self.show_row_preview)

        self.refresh_tags()

    def eventFilter(self, source, event):
        if source == self.table_view and event.type() == QEvent.Type.KeyPress:
            if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                self.show_row_preview()
                return True
        return super().eventFilter(source, event)

    def show_row_preview(self):
        indexes = self.table_view.selectionModel().selectedRows()
        if indexes:
            row_idx = indexes[0].row()
        else:
            curr = self.table_view.currentIndex()
            if not curr.isValid(): return
            row_idx = curr.row()

        if 0 <= row_idx < len(self.model._display_data):
            row_data = self.model._display_data[row_idx]
            dialog = RowPreviewDialog(row_data, self.image_folder_path, self)
            dialog.setGraphicsEffect(QGraphicsDropShadowEffect(blurRadius=12, xOffset=0, yOffset=0, color=QColor(0, 0, 0, 160)))
            dialog.exec()

    # ================= LOGIC TAG QUẢN LÝ CỘT =================
    def refresh_tags(self):
        self.tag_list.clear()
        self.combo_add_col.clear()
        self.combo_add_col.addItem("-- Add Hidden Column --")
        
        if self.hidden_cols:
            self.combo_add_col.addItems(self.hidden_cols)
            
        for col_name in self.active_cols:
            tag_widget = TagWidget(col_name)
            tag_widget.removed.connect(self.hide_column)
            
            item = QListWidgetItem()
            item.setSizeHint(QSize(100, 24)) 
            self.tag_list.addItem(item)
            self.tag_list.setItemWidget(item, tag_widget)

    def hide_column(self, col_name):
        self.active_cols.remove(col_name)
        self.hidden_cols.append(col_name)
        col_idx = HEADERS.index(col_name)
        self.table_view.horizontalHeader().setSectionHidden(col_idx, True)
        self.refresh_tags()

    def show_column(self, index):
        if index == 0: return 
        col_name = self.combo_add_col.currentText()
        
        self.hidden_cols.remove(col_name)
        self.active_cols.append(col_name)
        self.active_cols.sort(key=lambda x: HEADERS.index(x))
        
        col_idx = HEADERS.index(col_name)
        self.table_view.horizontalHeader().setSectionHidden(col_idx, False)
        self.refresh_tags()

    # ================= CÁC HÀM XỬ LÝ KHÁC =================
    def select_image_folder(self):
        if self.folder_path:
            folder = f"{self.folder_path}/IMAGE/EXPORTEDIMAGES"
            if os.path.exists(folder):
                self.image_folder_path = folder
                self.image_delegate.image_folder = folder
                self.image_delegate.pixmap_cache.clear()
                self.table_view.viewport().update()

    def toggle_navigator(self, checked):
        self.nav_frame.setVisible(checked)
        self.btn_toggle_nav.setText("Filter ◀" if checked else "Filter ▶")

    def start_loading(self):
        options = (QFileDialog.Option.ShowDirsOnly | QFileDialog.Option.DontResolveSymlinks | QFileDialog.Option.DontUseNativeDialog)
        self.folder_path = QFileDialog.getExistingDirectory(self, "Select the folder containing the data.", "", options)
        if not self.folder_path: return
        
        xml_folder_path = f"{self.folder_path}/XML/XML_PM"
        
        dialog = ScanOptionDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            scan_all = dialog.radio_all.isChecked()
            start_dt = dialog.dt_from.dateTime().toPyDateTime()
            end_dt = dialog.dt_to.dateTime().toPyDateTime()

            self.btn_load.setEnabled(False)
            self.thread = DataLoaderThread(xml_folder_path, scan_all, start_dt, end_dt)
            self.thread.progress.connect(lambda c, t, n: self.lbl_status.setText(f"⏳ Loading ({c}/{t}): {n}"))
            self.thread.finished.connect(self.on_loading_finished)
            self.thread.start()

    def on_loading_finished(self, raw_data, unique_values):
        self.model.load_data(raw_data)
        self.build_filter_navigator(unique_values)
        self.btn_load.setEnabled(True)
        self.select_image_folder()
        self.lbl_status.setText(f"✅ Ready! Total lines: {len(raw_data)}")

    def build_filter_navigator(self, unique_values):
        if not unique_values or len(unique_values) < len(HEADERS): return

        for i in reversed(range(self.filter_layout.count())): 
            widget = self.filter_layout.itemAt(i).widget()
            if widget: widget.deleteLater()
        self.filter_combos.clear()

        date_idx = HEADERS.index("Test Date")
        dates = unique_values[date_idx]
        if dates:
            min_date, max_date = min(dates), max(dates)
            self.dt_from.setDateTime(QDateTime.fromString(min_date, "MM/dd/yyyy HH:mm:ss"))
            self.dt_to.setDateTime(QDateTime.fromString(max_date, "MM/dd/yyyy HH:mm:ss"))

        for col_idx, col_name in enumerate(HEADERS):
            if col_name in ["Test Date", "X(mm)", "Y(mm)", "Loc Insp Info", "Review Image", "Target Image"]:
                continue
            vals = unique_values[col_idx]
            if not vals: continue

            combo = CheckableComboBox()
            combo.addItems(vals)
            combo.selectionChanged.connect(self.apply_filters)
            combo.setFixedWidth(300)
            
            self.filter_layout.addRow(col_name + ":", combo)
            self.filter_combos[col_idx] = combo

    def apply_filters(self):
        active_filters = {}
        for col_idx, combo in self.filter_combos.items():
            checked_vals = combo.get_checked_items()
            if "-- All --" not in checked_vals:
                active_filters[col_idx] = checked_vals

        dt_start = self.dt_from.dateTime().toPyDateTime()
        dt_end = self.dt_to.dateTime().toPyDateTime()

        self.model.filter_data(active_filters, date_range=(dt_start, dt_end))
        self.lbl_status.setText(f"🔍 Filtered: {self.model.rowCount()} rows remaining.")

    def clear_all_filters(self):
        for combo in self.filter_combos.values():
            combo.blockSignals(True)
            combo.reset_checks()
            combo.blockSignals(False)
        self.apply_filters()

    def show_export_dialog(self):
        selected_indexes = self.table_view.selectionModel().selectedRows()
        export_data = []

        if selected_indexes:
            export_data = [self.model._display_data[index.row()] for index in selected_indexes]
        else:
            if self.model.rowCount() == 0:
                QMessageBox.warning(self, "Warning", "Data table is empty!")
                return
            reply = QMessageBox.question(self, "Confirm", "You haven't selected any specific rows.\nExport ALL currently displayed data?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes)
            if reply == QMessageBox.StandardButton.Yes:
                export_data = self.model._display_data.copy()
            else:
                return 

        dialog = ExportConfigDialog(self.hidden_cols, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            mode, target_img_col, selected_cols = dialog.get_configuration()
            
            if mode == "excel":
                if not selected_cols:
                    QMessageBox.warning(self, "Error", "You have not selected any columns to export to Excel.")
                    return
                save_path, _ = QFileDialog.getSaveFileName(self, "Save Excel file", "", "Excel Files (*.xlsx)")
                if not save_path: return

                self.btn_export.setEnabled(False)
                self.lbl_status.setText("⏳ Extracting to Excel...")
                self.export_thread = ExcelExportThread(export_data, save_path, self.image_folder_path, selected_cols)
                self.export_thread.progress.connect(lambda c, t, txt: self.lbl_status.setText(f"⏳ {txt} ({c}/{t})"))
                self.export_thread.finished.connect(self.on_export_finished)
                self.export_thread.start()

            elif mode == "image":
                save_dir = QFileDialog.getExistingDirectory(self, "Select folder to save images")
                if not save_dir: return

                target_col_idx = HEADERS.index(target_img_col)
                self.btn_export.setEnabled(False)
                self.lbl_status.setText("⏳ Copying images...")
                
                self.img_export_thread = ImageExportThread(export_data, save_dir, self.image_folder_path, target_col_idx)
                self.img_export_thread.progress.connect(lambda c, t, txt: self.lbl_status.setText(f"⏳ {txt} ({c}/{t})"))
                self.img_export_thread.finished.connect(self.on_export_finished)
                self.img_export_thread.start()

    def on_export_finished(self, success, message):
        self.btn_export.setEnabled(True)
        self.lbl_status.setText(message.replace("\n", " "))
        if success: QMessageBox.information(self, "Success", message)
        else: QMessageBox.critical(self, "Error", message)
