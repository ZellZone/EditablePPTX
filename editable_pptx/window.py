from __future__ import annotations

import ctypes
import os
import queue
import subprocess
import sys
import threading
import traceback
from dataclasses import dataclass

from PIL import Image
from PySide6.QtCore import QEvent, QObject, QSize, Qt, QThread, Signal
from PySide6.QtGui import QColor, QFont, QIcon, QImage, QKeySequence, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from editable_pptx.convert import analyze_image
from editable_pptx.export_pptx import build_presentation
from editable_pptx.imaging import (
    app_dir,
    collect_image_paths,
    compute_slide_size,
    load_image,
    make_demo_image,
)
from editable_pptx.model import ConvertOptions, SlideSize, SlideState, TextBlock
from editable_pptx.ocr_engine import ensure_mta, query_languages
from editable_pptx.preview import render_preview

STYLE = """
QMainWindow, QWidget#AppRoot, QWidget#Center {
    background: #f3efe8;
    color: #1f2a2e;
    font-family: "Microsoft JhengHei UI", "Segoe UI";
    font-size: 13px;
}
QFrame#Header {
    background: #fffdf9;
    border-bottom: 1px solid #e4ddd2;
}
QLabel#Wordmark {
    font-size: 18px;
    font-weight: 700;
    color: #1f2a2e;
}
QLabel#Tagline { color: #6d7874; padding-left: 8px; }
QFrame#SideBar { background: #1f2a2e; }
QLabel#SideTitle { color: #f6f3ee; font-weight: 600; }
QLabel#CountLabel { color: #b7c2bc; }
QListWidget {
    background: #1f2a2e;
    color: #f4f1ea;
    border: none;
    outline: none;
    padding: 4px;
}
QListWidget::item { border-radius: 8px; padding: 4px; margin: 2px 4px; }
QListWidget::item:selected { background: #c45c26; color: white; }
QListWidget::item:hover:!selected { background: #2c3a40; }
QFrame#SideBar QPushButton {
    background: #2a383e;
    color: #f6f3ee;
    border: 1px solid #3c4c52;
    border-radius: 8px;
    padding: 7px 10px;
}
QFrame#SideBar QPushButton:hover { background: #35464d; }
QFrame#SideBar QPushButton:disabled { color: #7d8a86; background: #243034; }
QScrollArea#Settings, QWidget#Settings {
    background: #fffdf9;
    border: none;
}
QWidget#Settings { border-left: 1px solid #e4ddd2; }
QLabel#Section { font-weight: 700; color: #1f2a2e; }
QLabel#Help { color: #5c6b66; font-size: 12px; }
QLabel#Warn { color: #8a3d16; font-size: 12px; }
QLabel#State { color: #5c6b66; }
QPushButton {
    background: #f3ece3;
    color: #1f2a2e;
    border: 1px solid #e0d5c8;
    border-radius: 8px;
    padding: 7px 12px;
}
QPushButton:hover { background: #e8dfd3; }
QPushButton:disabled { color: #a09890; background: #f6f3ee; }
QPushButton#Primary {
    background: #c45c26;
    color: white;
    border: none;
    font-weight: 600;
    padding: 9px 16px;
}
QPushButton#Primary:hover { background: #a84c1d; }
QPushButton#Primary:disabled { background: #e3c2b2; color: white; }
QComboBox {
    background: white;
    color: #1f2a2e;
    border: 1px solid #cfc4b6;
    border-radius: 6px;
    padding: 4px 8px;
}
QComboBox::drop-down {
    border: none;
    width: 22px;
}
QComboBox::down-arrow {
    width: 0px;
    height: 0px;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #5c6b66;
}
QComboBox QAbstractItemView {
    background: white;
    color: #1f2a2e;
    selection-background-color: #c45c26;
    selection-color: white;
}
QCheckBox {
    color: #1f2a2e;
    spacing: 8px;
}
QLabel#Page { color: #1f2a2e; font-weight: 600; }
QTableWidget {
    background: white;
    color: #1f2a2e;
    border: 1px solid #e4ddd2;
    border-radius: 8px;
    gridline-color: #f0ebe4;
    selection-background-color: #f3d7c8;
    selection-color: #1f2a2e;
}
QTableWidget::item {
    color: #1f2a2e;
    padding: 4px;
}
QHeaderView::section {
    background: #f7f4ef;
    color: #1f2a2e;
    border: none;
    padding: 6px;
    font-weight: 600;
}
QProgressBar {
    border: none;
    background: #e7e0d6;
    border-radius: 6px;
    max-height: 14px;
    text-align: center;
}
QProgressBar::chunk { background: #c45c26; border-radius: 6px; }
QSplitter::handle { background: #e4ddd2; }
"""


def pil_to_qpixmap(image: Image.Image) -> QPixmap:
    rgb = image.convert("RGB")
    width, height = rgb.size
    data = rgb.tobytes("raw", "RGB")
    qimage = QImage(data, width, height, width * 3, QImage.Format_RGB888)
    return QPixmap.fromImage(qimage.copy())


def make_thumb(image: Image.Image, width: int = 160) -> Image.Image:
    height = max(1, int(width * image.height / image.width))
    return image.resize((width, height), Image.Resampling.BILINEAR)


def build_app_icon() -> QIcon:
    pix = QPixmap(64, 64)
    pix.fill(Qt.transparent)
    painter = QPainter(pix)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setPen(Qt.NoPen)
    painter.setBrush(QColor("#c45c26"))
    painter.drawRoundedRect(4, 4, 56, 56, 14, 14)
    painter.setBrush(QColor("#fffdf9"))
    painter.drawRoundedRect(14, 16, 36, 28, 4, 4)
    painter.setBrush(QColor("#1f2a2e"))
    painter.drawRoundedRect(18, 22, 20, 3, 1, 1)
    painter.drawRoundedRect(18, 28, 28, 3, 1, 1)
    painter.drawRoundedRect(18, 34, 14, 3, 1, 1)
    painter.end()
    return QIcon(pix)


def prefer_language(languages: list[tuple[str, str]]) -> int:
    tags = [tag.lower() for tag, _name in languages]
    for pref in ("zh-hant-tw", "zh-hant", "zh-tw", "zh-hans", "zh-cn", "en-us", "en"):
        for index, tag in enumerate(tags):
            if tag == pref or tag.startswith(pref):
                return index
    return 0


@dataclass
class BatchJob:
    job_id: int
    images: list[Image.Image]
    options: ConvertOptions
    slide_size: SlideSize
    cancel: threading.Event


class EngineThread(QThread):
    languages_ready = Signal(object, str)
    slide_ready = Signal(int, int, object, str)
    progressed = Signal(int, int)
    batch_finished = Signal(int, bool)
    batch_failed = Signal(int, str)

    def __init__(self) -> None:
        super().__init__()
        self._queue: queue.Queue = queue.Queue()

    def submit(self, job: BatchJob) -> None:
        self._queue.put(job)

    def shutdown(self) -> None:
        self.requestInterruption()
        self._queue.put(None)

    def run(self) -> None:
        ensure_mta()
        try:
            languages = query_languages()
            error = "" if languages else "這台電腦沒有可用的 Windows 光學字元辨識語言。"
        except Exception as exc:
            languages = []
            error = str(exc)
        self.languages_ready.emit(languages, error)
        while True:
            job = self._queue.get()
            if job is None or self.isInterruptionRequested():
                return
            self._process(job)

    def _process(self, job: BatchJob) -> None:
        total = len(job.images)
        try:
            for index, image in enumerate(job.images):
                if job.cancel.is_set() or self.isInterruptionRequested():
                    self.batch_finished.emit(job.job_id, True)
                    return
                try:
                    if job.options.mode == "text":
                        blocks = analyze_image(image, job.options, job.slide_size)
                        error = "" if blocks else "這張圖沒有辨識到文字，會以圖片方式放入。"
                    else:
                        blocks = []
                        error = ""
                except Exception as exc:
                    blocks = []
                    error = str(exc)
                self.slide_ready.emit(job.job_id, index, blocks, error)
                self.progressed.emit(index + 1, total)
        except Exception as exc:
            self.batch_failed.emit(job.job_id, str(exc))
            self.batch_finished.emit(job.job_id, True)
            return
        self.batch_finished.emit(job.job_id, False)


class PreviewCanvas(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._pix = QPixmap()
        self.setMinimumHeight(320)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    def set_pixmap(self, pixmap: QPixmap) -> None:
        self._pix = pixmap
        self.update()

    def paintEvent(self, _event) -> None:
        painter = QPainter(self)
        try:
            painter.setRenderHint(QPainter.Antialiasing, True)
            painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
            painter.fillRect(self.rect(), QColor("#efeae2"))
            if self._pix.isNull():
                pen = QPen(QColor("#c8bfb2"))
                pen.setStyle(Qt.DashLine)
                pen.setWidth(2)
                painter.setPen(pen)
                painter.setBrush(Qt.NoBrush)
                painter.drawRoundedRect(self.rect().adjusted(18, 18, -18, -18), 16, 16)
                painter.setPen(QColor("#6d7874"))
                painter.drawText(
                    self.rect().adjusted(32, 32, -32, -32),
                    Qt.AlignCenter,
                    "將圖片拖放到這裡\n或按「新增圖片」",
                )
                return
            margin = 28
            avail = self.rect().adjusted(margin, margin, -margin, -margin)
            scaled = self._pix.scaled(avail.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            x = avail.x() + (avail.width() - scaled.width()) // 2
            y = avail.y() + (avail.height() - scaled.height()) // 2
            painter.fillRect(x + 3, y + 6, scaled.width(), scaled.height(), QColor(31, 42, 46, 40))
            painter.drawPixmap(x, y, scaled)
        finally:
            painter.end()


class ImageDropFilter(QObject):
    def __init__(self, window: QMainWindow, callback, busy_check) -> None:
        super().__init__(window)
        self._window = window
        self._callback = callback
        self._busy_check = busy_check

    def eventFilter(self, obj, event) -> bool:
        if event.type() not in (QEvent.DragEnter, QEvent.DragMove, QEvent.Drop):
            return False
        if not isinstance(obj, QWidget) or obj.window() is not self._window:
            return False
        mime = event.mimeData() if hasattr(event, "mimeData") else None
        if mime is None or not mime.hasUrls():
            return False
        paths = collect_image_paths([url.toLocalFile() for url in mime.urls()])
        if not paths:
            return False
        event.acceptProposedAction()
        if event.type() == QEvent.Drop and not self._busy_check():
            self._callback(paths)
        return True


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.slides: list[SlideState] = []
        self.busy = False
        self.pending_export = False
        self.job_id = 0
        self.job_keys: dict[int, tuple] = {}
        self.current_cancel: threading.Event | None = None
        self.languages_loaded = False
        self.language_error = ""
        self.setWindowTitle("EditablePPTX")
        self.setWindowIcon(build_app_icon())
        self.resize(1360, 860)
        self.setMinimumSize(1100, 700)
        self._build_ui()
        self._connect()
        self.engine = EngineThread()
        self.engine.languages_ready.connect(self.on_languages)
        self.engine.slide_ready.connect(self.on_slide_ready)
        self.engine.progressed.connect(self.on_progress)
        self.engine.batch_finished.connect(self.on_batch_finished)
        self.engine.batch_failed.connect(self.on_batch_failed)
        self.engine.start()
        app = QApplication.instance()
        if app is not None:
            self._drop_filter = ImageDropFilter(self, self.add_paths, lambda: self.busy)
            app.installEventFilter(self._drop_filter)
        self.sync_buttons()
        self.statusBar().showMessage("加入圖片後，按「轉換並預覽」。")

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("AppRoot")
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header = QFrame()
        header.setObjectName("Header")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(18, 12, 18, 12)
        wordmark = QLabel("EditablePPTX")
        wordmark.setObjectName("Wordmark")
        tagline = QLabel("圖片轉成可編輯簡報")
        tagline.setObjectName("Tagline")
        self.export_btn = QPushButton("匯出 PPTX")
        self.export_btn.setObjectName("Primary")
        self.cancel_btn = QPushButton("取消轉換")
        self.cancel_btn.hide()
        header_layout.addWidget(wordmark)
        header_layout.addWidget(tagline)
        header_layout.addStretch(1)
        header_layout.addWidget(self.cancel_btn)
        header_layout.addWidget(self.export_btn)
        outer.addWidget(header)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_sidebar())
        splitter.addWidget(self._build_center())
        splitter.addWidget(self._build_settings())
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([250, 780, 320])
        outer.addWidget(splitter, 1)

        self.progress = QProgressBar()
        self.progress.setFixedWidth(180)
        self.progress.hide()
        status = QStatusBar()
        status.addPermanentWidget(self.progress)
        self.setStatusBar(status)

        file_menu = self.menuBar().addMenu("檔案")
        file_menu.addAction("新增圖片...", self.add_images, QKeySequence.Open)
        file_menu.addAction("加入範例圖片", self.add_demo)
        file_menu.addSeparator()
        file_menu.addAction("匯出 PPTX...", self.export_pptx, QKeySequence.Save)
        file_menu.addSeparator()
        file_menu.addAction("結束", self.close, QKeySequence.Quit)
        view_menu = self.menuBar().addMenu("檢視")
        view_menu.addAction("上一頁", self.prev_slide, QKeySequence(Qt.Key_PageUp))
        view_menu.addAction("下一頁", self.next_slide, QKeySequence(Qt.Key_PageDown))

    def _build_sidebar(self) -> QWidget:
        side = QFrame()
        side.setObjectName("SideBar")
        side.setMinimumWidth(220)
        layout = QVBoxLayout(side)
        layout.setContentsMargins(12, 14, 12, 14)
        layout.setSpacing(8)
        title = QLabel("投影片")
        title.setObjectName("SideTitle")
        self.count_label = QLabel("尚未加入圖片")
        self.count_label.setObjectName("CountLabel")
        self.list = QListWidget()
        self.list.setIconSize(QSize(112, 63))
        self.list.setSpacing(2)
        self.list.setTextElideMode(Qt.ElideRight)
        self.list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.add_btn = QPushButton("新增圖片")
        self.demo_btn = QPushButton("加入範例")
        self.remove_btn = QPushButton("移除")
        self.clear_btn = QPushButton("清空")
        self.up_btn = QPushButton("上移")
        self.down_btn = QPushButton("下移")
        row_a = QHBoxLayout()
        row_a.addWidget(self.remove_btn)
        row_a.addWidget(self.clear_btn)
        row_b = QHBoxLayout()
        row_b.addWidget(self.up_btn)
        row_b.addWidget(self.down_btn)
        layout.addWidget(title)
        layout.addWidget(self.count_label)
        layout.addWidget(self.list, 1)
        layout.addWidget(self.add_btn)
        layout.addWidget(self.demo_btn)
        layout.addLayout(row_a)
        layout.addLayout(row_b)
        return side

    def _build_center(self) -> QWidget:
        center = QWidget()
        center.setObjectName("Center")
        layout = QVBoxLayout(center)
        layout.setContentsMargins(16, 14, 16, 12)
        layout.setSpacing(8)
        self.state_label = QLabel("將圖片拖放到這裡")
        self.state_label.setObjectName("State")
        self.canvas = PreviewCanvas()
        nav = QHBoxLayout()
        self.prev_btn = QPushButton("上一頁")
        self.page_label = QLabel("尚未加入圖片")
        self.page_label.setObjectName("Page")
        self.page_label.setAlignment(Qt.AlignCenter)
        self.next_btn = QPushButton("下一頁")
        nav.addWidget(self.prev_btn)
        nav.addStretch(1)
        nav.addWidget(self.page_label)
        nav.addStretch(1)
        nav.addWidget(self.next_btn)
        self.table_hint = QLabel("轉換完成後，辨識出的文字會列在下面，雙擊即可修改。")
        self.table_hint.setObjectName("Help")
        self.table_hint.setWordWrap(True)
        self.table = QTableWidget(0, 1)
        self.table.setHorizontalHeaderLabels(["文字內容（雙擊可修改）"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setMinimumHeight(120)
        self.table.setMaximumHeight(190)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.delete_text_btn = QPushButton("刪除選取的文字")
        layout.addWidget(self.state_label)
        layout.addWidget(self.canvas, 1)
        layout.addLayout(nav)
        layout.addWidget(self.table_hint)
        layout.addWidget(self.table)
        layout.addWidget(self.delete_text_btn, 0, Qt.AlignLeft)
        return center

    def _build_settings(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setObjectName("Settings")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setMinimumWidth(280)
        panel = QWidget()
        panel.setObjectName("Settings")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(8)

        def section(text: str) -> QLabel:
            label = QLabel(text)
            label.setObjectName("Section")
            return label

        layout.addWidget(section("轉換模式"))
        self.mode = QComboBox()
        self.mode.addItem("抽出文字（可編輯文字方塊）", "text")
        self.mode.addItem("整頁圖片（可移動的圖片）", "image")
        layout.addWidget(self.mode)

        layout.addWidget(section("投影片比例"))
        self.ratio = QComboBox()
        self.ratio.addItem("16:9 寬螢幕", "16:9")
        self.ratio.addItem("4:3 標準", "4:3")
        self.ratio.addItem("依第一張圖片比例", "first")
        layout.addWidget(self.ratio)

        layout.addWidget(section("圖片配置"))
        self.fit = QComboBox()
        self.fit.addItem("完整顯示（留白）", "contain")
        self.fit.addItem("填滿頁面（裁切）", "cover")
        layout.addWidget(self.fit)

        layout.addWidget(section("辨識語言"))
        self.lang = QComboBox()
        self.lang.addItem("正在檢查語言...", "")
        layout.addWidget(self.lang)
        self.ocr_warning = QLabel(
            "若沒有辨識語言，請到 設定 → 時間與語言 → 語言與地區，"
            "開啟語言選項並安裝光學字元辨識。安裝前仍可把圖片做成可移動的圖片投影片。"
        )
        self.ocr_warning.setObjectName("Warn")
        self.ocr_warning.setWordWrap(True)
        self.ocr_warning.hide()
        layout.addWidget(self.ocr_warning)

        self.cover = QCheckBox("用背景色蓋住原圖文字")
        self.cover.setChecked(True)
        self.boxes = QCheckBox("預覽時框出文字方塊")
        self.boxes.setChecked(True)
        layout.addWidget(self.cover)
        layout.addWidget(self.boxes)

        help_label = QLabel(
            "抽出文字會把圖中的字變成 PowerPoint 文字方塊，匯出後可以直接修改。"
            "純色背景的簡報最清楚。底圖很複雜時，可以取消覆蓋，文字方塊會疊在原圖上。"
        )
        help_label.setObjectName("Help")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)
        layout.addStretch(1)
        self.convert_btn = QPushButton("轉換並預覽")
        self.convert_btn.setObjectName("Primary")
        layout.addWidget(self.convert_btn)
        scroll.setWidget(panel)
        return scroll

    def _connect(self) -> None:
        self.add_btn.clicked.connect(self.add_images)
        self.demo_btn.clicked.connect(self.add_demo)
        self.remove_btn.clicked.connect(self.remove_slide)
        self.clear_btn.clicked.connect(self.clear_slides)
        self.up_btn.clicked.connect(lambda: self.move_slide(-1))
        self.down_btn.clicked.connect(lambda: self.move_slide(1))
        self.prev_btn.clicked.connect(self.prev_slide)
        self.next_btn.clicked.connect(self.next_slide)
        self.list.currentRowChanged.connect(self.on_row_changed)
        self.mode.currentIndexChanged.connect(self.on_options_changed)
        self.ratio.currentIndexChanged.connect(self.on_options_changed)
        self.fit.currentIndexChanged.connect(self.on_options_changed)
        self.lang.currentIndexChanged.connect(self.on_options_changed)
        self.cover.toggled.connect(lambda _checked: self.update_preview(True))
        self.boxes.toggled.connect(lambda _checked: self.update_preview(True))
        self.convert_btn.clicked.connect(self.convert_all)
        self.export_btn.clicked.connect(self.export_pptx)
        self.cancel_btn.clicked.connect(self.cancel_convert)
        self.table.itemChanged.connect(self.on_text_edited)
        self.delete_text_btn.clicked.connect(self.delete_text_row)

    def collect_options(self) -> ConvertOptions:
        return ConvertOptions(
            mode=self.mode.currentData() or "text",
            ratio=self.ratio.currentData() or "16:9",
            fit=self.fit.currentData() or "contain",
            lang=self.lang.currentData() or "",
            cover_text=self.cover.isChecked(),
            show_boxes=self.boxes.isChecked(),
        )

    def current_size(self) -> SlideSize:
        image = self.slides[0].image if self.slides else None
        return compute_slide_size(self.ratio.currentData() or "16:9", image)

    def active_blocks(self, slide: SlideState, options: ConvertOptions, size: SlideSize) -> list[TextBlock] | None:
        if options.mode != "text":
            return []
        if slide.blocks is None or slide.key != options.cache_key(size):
            return None
        return slide.blocks

    def on_languages(self, languages: list, error: str) -> None:
        self.languages_loaded = True
        self.language_error = error
        self.lang.blockSignals(True)
        self.lang.clear()
        if languages:
            for tag, name in languages:
                self.lang.addItem(f"{name}  ({tag})", tag)
            self.lang.setCurrentIndex(prefer_language(languages))
            self.ocr_warning.hide()
        else:
            self.lang.addItem("沒有可用語言", "")
            self.ocr_warning.show()
            if error:
                self.statusBar().showMessage(error)
        self.lang.blockSignals(False)
        self.apply_mode_widgets()

    def apply_mode_widgets(self) -> None:
        text_mode = (self.mode.currentData() or "text") == "text"
        enabled = text_mode and not self.busy
        self.cover.setEnabled(enabled)
        self.boxes.setEnabled(enabled)
        self.lang.setEnabled(enabled)

    def on_options_changed(self) -> None:
        self.apply_mode_widgets()
        self.update_preview(True)
        options = self.collect_options()
        size = self.current_size()
        if self.slides and options.mode == "text":
            if any(slide.blocks is not None and slide.key != options.cache_key(size) for slide in self.slides):
                self.statusBar().showMessage("設定已變更，請再按一次「轉換並預覽」。")

    def add_images(self) -> None:
        pictures = os.path.join(os.path.expanduser("~"), "Pictures")
        paths, _selected = QFileDialog.getOpenFileNames(
            self,
            "選擇圖片",
            pictures if os.path.isdir(pictures) else os.path.expanduser("~"),
            "圖片 (*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff *.gif)",
        )
        if paths:
            self.add_paths(paths)

    def add_demo(self) -> None:
        image = make_demo_image()
        folder = os.path.join(app_dir(), "output")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "demo-slide.png")
        image.save(path, "PNG")
        self._append_slide(path, image, "範例投影片.png")
        self.refresh_labels()
        self.list.setCurrentRow(self.list.count() - 1)
        self.statusBar().showMessage("已加入範例投影片。按「轉換並預覽」即可看結果。")

    def add_paths(self, paths: list[str]) -> None:
        if self.busy:
            self.statusBar().showMessage("轉換中，請稍候再加入圖片。")
            return
        files = collect_image_paths(paths)
        if not files:
            self.statusBar().showMessage("沒有找到支援的圖片。")
            return
        existing = {os.path.normcase(slide.path) for slide in self.slides}
        added = 0
        skipped = 0
        errors: list[str] = []
        self.list.blockSignals(True)
        for path in files:
            key = os.path.normcase(path)
            if key in existing:
                skipped += 1
                continue
            try:
                image = load_image(path)
            except Exception as exc:
                errors.append(f"{os.path.basename(path)}：{exc}")
                continue
            self._append_slide(path, image, "")
            existing.add(key)
            added += 1
        self.refresh_labels()
        self.list.blockSignals(False)
        if added:
            self.list.setCurrentRow(self.list.count() - 1)
        else:
            self.sync_buttons()
        if errors:
            QMessageBox.warning(self, "部分圖片無法讀取", "\n".join(errors[:8]))
        if added:
            extra = f"，略過 {skipped} 張重複圖片" if skipped else ""
            self.statusBar().showMessage(f"已加入 {added} 張圖片{extra}。")
        elif skipped:
            self.statusBar().showMessage("這些圖片已經在清單裡。")

    def _append_slide(self, path: str, image: Image.Image, name: str) -> None:
        slide = SlideState(path=path, image=image, name=name)
        self.slides.append(slide)
        item = QListWidgetItem(slide.display_name)
        item.setIcon(QIcon(pil_to_qpixmap(make_thumb(image))))
        item.setSizeHint(QSize(210, 78))
        self.list.addItem(item)

    def refresh_labels(self) -> None:
        for index, slide in enumerate(self.slides):
            item = self.list.item(index)
            if item is None:
                continue
            item.setText(f"{index + 1:02d}  {slide.display_name}")
            item.setToolTip(slide.path)
        self.count_label.setText(f"{len(self.slides)} 張圖片" if self.slides else "尚未加入圖片")

    def remove_slide(self) -> None:
        row = self.list.currentRow()
        if row < 0 or self.busy:
            return
        del self.slides[row]
        self.list.blockSignals(True)
        self.list.takeItem(row)
        self.refresh_labels()
        self.list.blockSignals(False)
        if self.slides:
            self.list.setCurrentRow(min(row, len(self.slides) - 1))
        else:
            self.update_preview(True)
            self.sync_buttons()

    def clear_slides(self) -> None:
        if self.busy or not self.slides:
            return
        self.slides.clear()
        self.list.clear()
        self.refresh_labels()
        self.update_preview(True)
        self.sync_buttons()
        self.statusBar().showMessage("已清空圖片。")

    def move_slide(self, delta: int) -> None:
        row = self.list.currentRow()
        new_row = row + delta
        if row < 0 or self.busy or not (0 <= new_row < len(self.slides)):
            return
        self.slides[row], self.slides[new_row] = self.slides[new_row], self.slides[row]
        self.list.blockSignals(True)
        item = self.list.takeItem(row)
        self.list.insertItem(new_row, item)
        self.refresh_labels()
        self.list.blockSignals(False)
        self.list.setCurrentRow(new_row)

    def prev_slide(self) -> None:
        row = self.list.currentRow()
        if row > 0:
            self.list.setCurrentRow(row - 1)

    def next_slide(self) -> None:
        row = self.list.currentRow()
        if 0 <= row < len(self.slides) - 1:
            self.list.setCurrentRow(row + 1)

    def on_row_changed(self, _row: int) -> None:
        self.update_preview(True)
        self.sync_buttons()

    def preview_for(self, index: int) -> Image.Image:
        slide = self.slides[index]
        options = self.collect_options()
        size = self.current_size()
        blocks = self.active_blocks(slide, options, size)
        token = (
            None if blocks is None else slide.key,
            slide.revision,
            options.cover_text,
            options.show_boxes,
            options.fit,
            options.mode,
            size.emu_w,
            size.emu_h,
        )
        if slide.preview is not None and slide.preview_token == token:
            return slide.preview
        image = render_preview(
            slide.image,
            blocks,
            size,
            options.fit,
            options.cover_text,
            bool(options.show_boxes and blocks),
            width=1600,
        )
        slide.preview = image
        slide.preview_token = token
        return image

    def update_preview(self, refill_table: bool) -> None:
        row = self.list.currentRow()
        options = self.collect_options()
        if row < 0 or row >= len(self.slides):
            self.canvas.set_pixmap(QPixmap())
            self.page_label.setText("尚未加入圖片")
            self.state_label.setText("將圖片拖放到這裡")
            self.table_hint.setText("轉換完成後，辨識出的文字會列在下面，雙擊即可修改。")
            if refill_table:
                self.fill_table(None)
            self.sync_buttons()
            return
        slide = self.slides[row]
        preview = self.preview_for(row)
        self.canvas.set_pixmap(pil_to_qpixmap(preview))
        self._set_icon(row, preview)
        blocks = self.active_blocks(slide, options, self.current_size())
        self.page_label.setText(f"第 {row + 1} / {len(self.slides)} 頁")
        if options.mode != "text":
            state = "整頁圖片 · 可在 PowerPoint 移動與縮放"
            hint = "此模式保留整張圖片，不抽出文字。"
        elif blocks is None:
            state = "原圖配置 · 按「轉換並預覽」抽出文字"
            hint = "轉換完成後，辨識出的文字會列在下面，雙擊即可修改。"
        elif not blocks:
            state = slide.error or "沒有辨識到文字"
            hint = "這張圖會以圖片方式放入投影片。"
        else:
            count = len([block for block in blocks if block.text.strip()])
            state = f"可編輯預覽 · {count} 個文字方塊"
            hint = "雙擊文字即可修改，匯出時會寫入 PPTX。"
        self.state_label.setText(state)
        self.table_hint.setText(hint)
        if refill_table:
            self.fill_table(blocks)
        self.sync_buttons()

    def _set_icon(self, index: int, preview: Image.Image) -> None:
        item = self.list.item(index)
        if item is not None:
            item.setIcon(QIcon(pil_to_qpixmap(make_thumb(preview))))

    def fill_table(self, blocks: list[TextBlock] | None) -> None:
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        rows = blocks or []
        self.table.setRowCount(len(rows))
        for index, block in enumerate(rows):
            self.table.setItem(index, 0, QTableWidgetItem(block.text))
        self.table.blockSignals(False)
        self.delete_text_btn.setEnabled(bool(rows) and not self.busy)

    def on_text_edited(self, item: QTableWidgetItem) -> None:
        row = self.list.currentRow()
        if row < 0:
            return
        slide = self.slides[row]
        blocks = self.active_blocks(slide, self.collect_options(), self.current_size())
        if not blocks or item.row() >= len(blocks):
            return
        blocks[item.row()].text = item.text().strip()
        from editable_pptx.imaging import fit_block_font

        fit_block_font(blocks[item.row()], self.current_size())
        slide.revision += 1
        slide.preview = None
        self.update_preview(False)

    def delete_text_row(self) -> None:
        row = self.list.currentRow()
        text_row = self.table.currentRow()
        if row < 0 or text_row < 0:
            return
        slide = self.slides[row]
        blocks = self.active_blocks(slide, self.collect_options(), self.current_size())
        if not blocks or text_row >= len(blocks):
            return
        del blocks[text_row]
        slide.revision += 1
        slide.preview = None
        self.update_preview(True)

    def convert_all(self) -> None:
        if self.busy or not self.slides:
            return
        options = self.collect_options()
        if options.mode == "text" and not options.lang:
            QMessageBox.warning(
                self,
                "無法辨識文字",
                self.language_error or "尚未取得文字辨識語言。可以改用「整頁圖片」模式匯出。",
            )
            return
        size = self.current_size()
        self.job_id += 1
        job_id = self.job_id
        self.job_keys[job_id] = options.cache_key(size)
        cancel = threading.Event()
        self.current_cancel = cancel
        self.engine.submit(
            BatchJob(
                job_id=job_id,
                images=[slide.image for slide in self.slides],
                options=options,
                slide_size=size,
                cancel=cancel,
            )
        )
        self.set_busy(True)
        self.progress.setRange(0, len(self.slides))
        self.progress.setValue(0)
        self.progress.show()
        self.statusBar().showMessage("正在轉換...")

    def cancel_convert(self) -> None:
        if self.current_cancel is not None:
            self.current_cancel.set()
        self.statusBar().showMessage("正在取消...")

    def on_progress(self, done: int, total: int) -> None:
        self.progress.setMaximum(max(1, total))
        self.progress.setValue(done)

    def on_slide_ready(self, job_id: int, index: int, blocks: list, error: str) -> None:
        if job_id != self.job_id or index >= len(self.slides):
            return
        slide = self.slides[index]
        slide.blocks = list(blocks)
        slide.key = self.job_keys.get(job_id)
        slide.error = error
        slide.revision += 1
        slide.preview = None
        if index == self.list.currentRow():
            self.update_preview(True)
        else:
            self._set_icon(index, self.preview_for(index))

    def on_batch_failed(self, job_id: int, message: str) -> None:
        if job_id != self.job_id:
            return
        self.pending_export = False
        QMessageBox.warning(self, "轉換失敗", message)

    def on_batch_finished(self, job_id: int, cancelled: bool) -> None:
        if job_id != self.job_id:
            return
        self.set_busy(False)
        self.progress.hide()
        if cancelled:
            self.pending_export = False
            self.statusBar().showMessage("已取消轉換。")
            self.update_preview(True)
            return
        options = self.collect_options()
        size = self.current_size()
        count = 0
        for slide in self.slides:
            blocks = self.active_blocks(slide, options, size) or []
            count += len([block for block in blocks if block.text.strip()])
        if options.mode == "text":
            self.statusBar().showMessage(f"轉換完成，共 {len(self.slides)} 頁、{count} 個可編輯文字方塊。")
        else:
            self.statusBar().showMessage(f"轉換完成，共 {len(self.slides)} 頁。每頁是一張可移動的圖片。")
        self.update_preview(True)
        if self.pending_export:
            self.pending_export = False
            self.write_pptx()

    def export_pptx(self) -> None:
        if self.busy or not self.slides:
            return
        options = self.collect_options()
        size = self.current_size()
        if options.mode == "text" and any(self.active_blocks(slide, options, size) is None for slide in self.slides):
            self.pending_export = True
            self.convert_all()
            return
        self.write_pptx()

    def write_pptx(self) -> None:
        options = self.collect_options()
        size = self.current_size()
        documents = os.path.join(os.path.expanduser("~"), "Documents")
        default = os.path.join(documents if os.path.isdir(documents) else os.path.expanduser("~"), "可編輯簡報.pptx")
        path, _selected = QFileDialog.getSaveFileName(self, "匯出 PPTX", default, "PowerPoint (*.pptx)")
        if not path:
            return
        if not path.lower().endswith(".pptx"):
            path += ".pptx"
        pairs = []
        for slide in self.slides:
            blocks = self.active_blocks(slide, options, size) or []
            pairs.append((slide.image, blocks))
        try:
            data = build_presentation(pairs, size, options)
            with open(path, "wb") as handle:
                handle.write(data)
        except PermissionError:
            QMessageBox.warning(self, "無法儲存", "檔案可能正被 PowerPoint 開啟，請先關閉同名簡報再匯出。")
            return
        except Exception as exc:
            QMessageBox.warning(self, "匯出失敗", str(exc))
            return
        self.statusBar().showMessage(f"已儲存 {path}")
        box = QMessageBox(self)
        box.setWindowTitle("匯出完成")
        box.setText("已儲存可編輯的 PPTX。")
        box.setInformativeText(path)
        open_btn = box.addButton("用 PowerPoint 開啟", QMessageBox.AcceptRole)
        folder_btn = box.addButton("開啟資料夾", QMessageBox.ActionRole)
        box.addButton("關閉", QMessageBox.RejectRole)
        box.exec()
        clicked = box.clickedButton()
        if clicked is open_btn:
            os.startfile(path)
        elif clicked is folder_btn:
            subprocess.run(["explorer", "/select,", os.path.normpath(path)], check=False)

    def set_busy(self, busy: bool) -> None:
        self.busy = busy
        self.cancel_btn.setVisible(busy)
        for widget in (
            self.add_btn,
            self.demo_btn,
            self.remove_btn,
            self.clear_btn,
            self.up_btn,
            self.down_btn,
            self.mode,
            self.ratio,
            self.fit,
            self.convert_btn,
            self.export_btn,
        ):
            widget.setEnabled(not busy)
        self.apply_mode_widgets()
        self.sync_buttons()

    def sync_buttons(self) -> None:
        has = bool(self.slides)
        row = self.list.currentRow()
        if not self.busy:
            self.convert_btn.setEnabled(has)
            self.export_btn.setEnabled(has)
            self.remove_btn.setEnabled(has)
            self.clear_btn.setEnabled(has)
            self.up_btn.setEnabled(has and row > 0)
            self.down_btn.setEnabled(has and 0 <= row < len(self.slides) - 1)
        self.prev_btn.setEnabled(has and row > 0)
        self.next_btn.setEnabled(has and 0 <= row < len(self.slides) - 1)
        blocks = None
        if 0 <= row < len(self.slides):
            blocks = self.active_blocks(self.slides[row], self.collect_options(), self.current_size())
        self.delete_text_btn.setEnabled(bool(blocks) and not self.busy)

    def closeEvent(self, event) -> None:
        if self.current_cancel is not None:
            self.current_cancel.set()
        self.engine.shutdown()
        self.engine.wait(15000)
        event.accept()


def _install_excepthook() -> None:
    def hook(exc_type, exc, tb) -> None:
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        log_path = os.path.join(os.environ.get("TEMP", "."), "EditablePPTX-error.log")
        try:
            with open(log_path, "w", encoding="utf-8") as handle:
                handle.write(text)
        except OSError:
            log_path = ""
        ctypes.windll.user32.MessageBoxW(0, f"{exc}\n\n{log_path}", "EditablePPTX", 0x10)

    sys.excepthook = hook


def main() -> None:
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("EditablePPTX.App")
    except Exception:
        pass
    _install_excepthook()
    app = QApplication(sys.argv)
    app.setApplicationName("EditablePPTX")
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft JhengHei UI", 10))
    app.setStyleSheet(STYLE)
    window = MainWindow()
    if len(sys.argv) > 1:
        window.add_paths(sys.argv[1:])
    window.show()
    sys.exit(app.exec())
