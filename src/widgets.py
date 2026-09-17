import sys
import os
import glob
import re
import shutil
import io
from datetime import datetime
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image
Image.MAX_IMAGE_PIXELS = None
from src.config import HEADERS
from PyQt6.QtWidgets import (  QVBoxLayout, 
                             QWidget,  QPushButton, 
                             QLabel, QHBoxLayout, QComboBox,
                             QScrollArea,  QDateTimeEdit, 
                             QStyledItemDelegate, QDialog, QRadioButton, QDialogButtonBox, 
                             QCheckBox, QGroupBox, 
                              QGridLayout)
from PyQt6.QtCore import Qt,  pyqtSignal, QDateTime, QEvent
from PyQt6.QtGui import QPixmap, QStandardItemModel, QStandardItem, QIcon

class TagWidget(QWidget):
    removed = pyqtSignal(str)

    def __init__(self, text, parent=None):
        super().__init__(parent)
        self.text = text
        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 2, 6, 2)
        layout.setSpacing(4)
        self.setStyleSheet("""
            QWidget { border: 1px solid #cce8ff; border-radius: 10px; }
        """)
        
        lbl = QLabel(text)
        lbl.setStyleSheet("border: none; font-size: 11px;")
        
        btn = QPushButton("✕")
        btn.setFixedSize(16, 16)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet("""
            QPushButton { border: none; background: transparent;  font-weight: bold; font-size: 10px; }
            QPushButton:hover { color: red; background-color: #fde7e9; border-radius: 8px;}
        """)
        btn.clicked.connect(lambda: self.removed.emit(self.text))
        
        layout.addWidget(lbl)
        layout.addWidget(btn)

class CheckableComboBox(QComboBox):
    selectionChanged = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setModel(QStandardItemModel(self))
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self.lineEdit().setText("-- All --")
        self._updating = False
        self.model().dataChanged.connect(self.on_data_changed)
        self.view().viewport().installEventFilter(self)

    def eventFilter(self, widget, event):
        if widget == self.view().viewport() and event.type() == QEvent.Type.MouseButtonRelease:
            index = self.view().indexAt(event.pos())
            if index.isValid():
                item = self.model().item(index.row())
                new_state = Qt.CheckState.Unchecked if item.checkState() == Qt.CheckState.Checked else Qt.CheckState.Checked
                item.setCheckState(new_state)
            return True 
        return super().eventFilter(widget, event)

    def addItems(self, texts):
        item_all = QStandardItem("-- All --")
        item_all.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
        item_all.setData(Qt.CheckState.Checked, Qt.ItemDataRole.CheckStateRole)
        self.model().appendRow(item_all)
        
        for text in texts:
            item = QStandardItem(text)
            item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            item.setData(Qt.CheckState.Unchecked, Qt.ItemDataRole.CheckStateRole)
            self.model().appendRow(item)

    def on_data_changed(self, top_left, bottom_right, roles):
        if self._updating: return
        self._updating = True
        first_item = self.model().item(0)
        changed_row = top_left.row()
        
        if changed_row == 0:
            if first_item.checkState() == Qt.CheckState.Checked:
                for i in range(1, self.model().rowCount()):
                    self.model().item(i).setCheckState(Qt.CheckState.Unchecked)
        else:
            if self.model().item(changed_row).checkState() == Qt.CheckState.Checked:
                first_item.setCheckState(Qt.CheckState.Unchecked)
        
        checked_items = [self.model().item(i).text() for i in range(1, self.model().rowCount()) if self.model().item(i).checkState() == Qt.CheckState.Checked]
        
        if not checked_items:
            first_item.setCheckState(Qt.CheckState.Checked)
            self.lineEdit().setText("-- All --")
        else:
            self.lineEdit().setText(", ".join(checked_items))
            
        self._updating = False
        self.selectionChanged.emit()

    def get_checked_items(self):
        if self.model().item(0).checkState() == Qt.CheckState.Checked:
            return ["-- All --"]
        return [self.model().item(i).text() for i in range(1, self.model().rowCount()) if self.model().item(i).checkState() == Qt.CheckState.Checked]

    def reset_checks(self):
        self.model().item(0).setCheckState(Qt.CheckState.Checked)

class ScanOptionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Open")
        self.setMinimumWidth(350)
        layout = QVBoxLayout(self)
        
        self.radio_all = QRadioButton("Scan entire folder (Load All)")
        self.radio_all.setChecked(True)
        self.radio_range = QRadioButton("Select the time you want:")
        
        self.dt_from = QDateTimeEdit(QDateTime.currentDateTime().addDays(-7))
        self.dt_from.setDisplayFormat("MM/dd/yyyy HH:mm:ss")
        self.dt_from.setCalendarPopup(True)
        self.dt_from.setEnabled(False)
        
        self.dt_to = QDateTimeEdit(QDateTime.currentDateTime())
        self.dt_to.setDisplayFormat("MM/dd/yyyy HH:mm:ss")
        self.dt_to.setCalendarPopup(True)
        self.dt_to.setEnabled(False)

        self.radio_all.toggled.connect(self.toggle_dates)
        self.radio_range.toggled.connect(self.toggle_dates)

        layout.addWidget(self.radio_all)
        layout.addWidget(self.radio_range)
        layout.addWidget(QLabel("From (A):"))
        layout.addWidget(self.dt_from)
        layout.addWidget(QLabel("To (B):"))
        layout.addWidget(self.dt_to)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def toggle_dates(self):
        is_range = self.radio_range.isChecked()
        self.dt_from.setEnabled(is_range)
        self.dt_to.setEnabled(is_range)

class ImageDelegate(QStyledItemDelegate):
    def __init__(self, image_folder, parent=None):
        super().__init__(parent)
        self.image_folder = image_folder
        self.pixmap_cache = {}

    def paint(self, painter, option, index):
        image_name = index.data(Qt.ItemDataRole.DisplayRole)
        if not image_name or image_name.strip() == "":
            super().paint(painter, option, index)
            return

        if image_name not in self.pixmap_cache:
            file_name = image_name if image_name.lower().endswith(".jpg") else f"{image_name}.jpg"
            img_path = os.path.join(self.image_folder, file_name)
            
            if os.path.exists(img_path):
                pixmap = QPixmap(img_path)
                scaled_pixmap = pixmap.scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                self.pixmap_cache[image_name] = scaled_pixmap
            else:
                self.pixmap_cache[image_name] = None

        pixmap = self.pixmap_cache.get(image_name)
        if pixmap:
            rect = option.rect
            x = rect.x() + (rect.width() - pixmap.width()) // 2
            y = rect.y() + (rect.height() - pixmap.height()) // 2
            painter.drawPixmap(x, y, pixmap)
        else:
            super().paint(painter, option, index)

class RowPreviewDialog(QDialog):
    def __init__(self, row_data, image_folder, parent=None):
        super().__init__(parent)
        loc = row_data[HEADERS.index("Location")] if "Location" in HEADERS else ""
        self.setWindowTitle(f"Detect - {loc or 'Row Detail'}")
        self.setWindowIcon(QIcon("detecticon.ico"))
        self.setMinimumWidth(560)
        self.setStyleSheet("""
            titlebar { background-color: #f0f0f0; font-weight: bold; }
            QDialog { background-color: white; }  
            QGroupBox { font-weight: bold; border: 1px solid #d0d7de; border-radius: 6px; margin-top: 8px; padding-top: 10px;  }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            QLabel { font-size: 12px; }
            QPushButton { padding: 5px 15px; border-radius: 4px; border: 1px solid #ccc; }
            QPushButton:hover {  }
        """)
        
      
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Khung hiển thị hình ảnh
        img_box = QGroupBox("Images Overview")
        img_layout = QHBoxLayout(img_box)
        img_layout.setSpacing(20)

        for title, col_name in [("Review Image", "Review Image"), ("Target Image", "Target Image")]:
            col_idx = HEADERS.index(col_name)
            img_name = row_data[col_idx] if col_idx < len(row_data) else ""
            
            sub_layout = QVBoxLayout()
            lbl_title = QLabel(f"<b>{title}</b>")
            lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
            sub_layout.addWidget(lbl_title)

            lbl_img = QLabel()
            lbl_img.setFixedSize(160, 160)
            lbl_img.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl_img.setStyleSheet("border: 1px dashed #bcc3ce;  border-radius: 4px;")

            if img_name:
                fname = img_name if img_name.lower().endswith(".jpg") else f"{img_name}.jpg"
                path = os.path.join(image_folder, fname)
                if os.path.exists(path):
                    pix = QPixmap(path).scaled(156, 156, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                    lbl_img.setPixmap(pix)
                else:
                    lbl_img.setText(f"Not Found\n({fname})")
            else:
                lbl_img.setText("No Image")

            sub_layout.addWidget(lbl_img, alignment=Qt.AlignmentFlag.AlignCenter)
            img_layout.addLayout(sub_layout)

        layout.addWidget(img_box)

        # Khung thông tin chi tiết
        info_box = QGroupBox("Row Details")
        info_grid = QGridLayout(info_box)
        info_grid.setHorizontalSpacing(16)
        info_grid.setVerticalSpacing(6)

        display_headers = [h for h in HEADERS if h not in ("Review Image", "Target Image")]
        for idx, h in enumerate(display_headers):
            h_idx = HEADERS.index(h)
            val = str(row_data[h_idx]) if h_idx < len(row_data) else ""
            r, c = divmod(idx, 2)
            lbl_k = QLabel(f"<b>{h}:</b>")
            lbl_k.setStyleSheet("font-weight: bold;")
            lbl_v = QLabel(val if val else "-")
            lbl_v.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            if h == "Review":
                lbl_v.setStyleSheet("color: #1a7f37; font-weight: bold;" if str(val).lower() == "accept" else "color: #cf222e; font-weight: bold;")
            elif h == "Inspection" and str(val).lower() not in ("pass", "accept"):
                lbl_v.setStyleSheet("color: #cf222e; font-weight: bold;")

            info_grid.addWidget(lbl_k, r, c * 2)
            info_grid.addWidget(lbl_v, r, c * 2 + 1)
        layout.addWidget(info_box)

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)


class ExportConfigDialog(QDialog):
    def __init__(self, hidden_cols, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Exports")
        self.setMinimumWidth(400)
        layout = QVBoxLayout(self)

        mode_group = QGroupBox("Select export mode:")
        mode_layout = QVBoxLayout()
        self.radio_excel = QRadioButton("Export selected columns to Excel.")
        self.radio_excel.setChecked(True)
        self.radio_image = QRadioButton("Export only the image column.")
        
        mode_layout.addWidget(self.radio_excel)
        mode_layout.addWidget(self.radio_image)
        mode_group.setLayout(mode_layout)
        layout.addWidget(mode_group)

        img_layout = QHBoxLayout()
        img_layout.addWidget(QLabel("Selection column"))
        self.combo_target_img = QComboBox()
        self.combo_target_img.addItems(["Review Image", "Target Image"])
        self.combo_target_img.setEnabled(False)
        img_layout.addWidget(self.combo_target_img)
        layout.addLayout(img_layout)

        col_group = QGroupBox("Select columns to export to Excel:")
        self.col_layout = QVBoxLayout()
        
        btn_layout = QHBoxLayout()
        btn_select_all = QPushButton("Select All")
        btn_select_all.clicked.connect(lambda: self.set_all_checkboxes(True))
        btn_clear_all = QPushButton("Deselect All")
        btn_clear_all.clicked.connect(lambda: self.set_all_checkboxes(False))
        btn_layout.addWidget(btn_select_all)
        btn_layout.addWidget(btn_clear_all)
        self.col_layout.addLayout(btn_layout)

        self.checkboxes = []
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        for idx, header in enumerate(HEADERS):
            cb = QCheckBox(header)
            # Tự động gỡ chọn nếu cột đó đang bị ẩn trên bảng Tag Manager
            cb.setChecked(header not in hidden_cols) 
            cb.setProperty("col_idx", idx)
            self.checkboxes.append(cb)
            scroll_layout.addWidget(cb)
            
        scroll.setWidget(scroll_content)
        self.col_layout.addWidget(scroll)
        col_group.setLayout(self.col_layout)
        layout.addWidget(col_group)

        self.radio_excel.toggled.connect(self.toggle_ui)
        self.radio_image.toggled.connect(self.toggle_ui)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def set_all_checkboxes(self, state):
        for cb in self.checkboxes:
            cb.setChecked(state)

    def toggle_ui(self):
        is_img_mode = self.radio_image.isChecked()
        self.combo_target_img.setEnabled(is_img_mode)
        for i in range(self.col_layout.count()):
            widget = self.col_layout.itemAt(i).widget()
            if widget: widget.setEnabled(not is_img_mode)

    def get_configuration(self):
        mode = "excel" if self.radio_excel.isChecked() else "image"
        target_img_col = self.combo_target_img.currentText()
        selected_cols = [cb.property("col_idx") for cb in self.checkboxes if cb.isChecked()]
        return mode, target_img_col, selected_cols