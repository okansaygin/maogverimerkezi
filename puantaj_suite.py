# -*- coding: utf-8 -*-
"""
PUANTAJ & RAPORLAMA MERKEZİ (puantaj_suite.py)
=================================================
Puantaj aktarımı ve ekip fazla mesai raporlamasını TEK bir pencerede,
tek bir profesyonel arayüzde birleştiren uygulama.

Klasör yapısı (bu dosyanın bulunduğu klasörde):
    SOURCE/   -> aylık kaynak puantaj exceli
    TARGET/   -> puantaj aktarımı için günlük hedef excel(ler)
    OUTPUT/   -> her iki modülün de çıktıları buraya kaydedilir

Gerekli dosyalar (hepsi aynı klasörde olmalı):
    puantaj_engine.py
    team_report.py
    puantaj_suite.py   (bu dosya - tek çalıştırılacak dosya budur)

Kurulum (bir defalık):
    pip install openpyxl PySide6

Çalıştırma:
    python puantaj_suite.py

Not: Eskiden ayrı ayrı çalışan puantaj_gui.py ve team_report_gui.py'nin
yerini bu dosya alır. Ayar dosyaları (puantaj_settings.json,
team_report_settings.json, team_report_history.json) aynı isimlerle
kullanılmaya devam eder - eski ayarların kaybolmasına gerek yok.
"""

import datetime
import html as html_lib
import json
import os
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QDate, Signal
from PySide6.QtGui import QTextCursor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QLabel, QLineEdit, QPushButton, QListWidget, QAbstractItemView,
    QListWidgetItem, QDateEdit, QCheckBox, QComboBox, QTextEdit, QFileDialog,
    QMessageBox, QDoubleSpinBox, QTreeWidget, QTreeWidgetItem, QStackedWidget,
    QFrame, QProgressBar, QDialog, QSplitter, QButtonGroup,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractScrollArea, QRadioButton,
)

import puantaj_engine as engine
import team_report as report
import cost_report as cost
import personnel_db as db

SCRIPT_DIR = Path(__file__).resolve().parent
SOURCE_DIR = SCRIPT_DIR / "SOURCE"
TARGET_DIR = SCRIPT_DIR / "TARGET"
OUTPUT_DIR = SCRIPT_DIR / "OUTPUT"

PUANTAJ_SETTINGS_PATH = SCRIPT_DIR / "puantaj_settings.json"
TEAM_SETTINGS_PATH = SCRIPT_DIR / "team_report_settings.json"
TEAM_HISTORY_PATH = SCRIPT_DIR / "team_report_history.json"
SUITE_SETTINGS_PATH = SCRIPT_DIR / "suite_settings.json"

TAG_COLORS = {
    "green": "#1b5e20", "red": "#b71c1c", "orange": "#e65100",
    "blue": "#0d47a1", "purple": "#6a1b9a", "black": "#212121",
}

PUANTAJ_ADVANCED_FIELDS = [
    ("source_sheet", "Kaynak sayfa adı", str),
    ("source_tc_col", "Kaynak TC sütunu", str),
    ("source_name_col", "Kaynak isim sütunu", str),
    ("source_type_col", "Kaynak NORMAL/FAZLA sütunu", str),
    ("source_date_col_start", "Kaynak gün sütunları başlangıç", str),
    ("source_date_col_end", "Kaynak gün sütunları bitiş", str),
    ("source_date_header_row", "Kaynak tarih satırı", int),
    ("source_data_start_row", "Kaynak veri başlangıç satırı", int),
    ("source_data_end_row", "Kaynak veri bitiş satırı", int),
    ("target_sheet", "Hedef sayfa adı", str),
    ("target_tc_col", "Hedef TC sütunu", str),
    ("target_name_col", "Hedef isim sütunu", str),
    ("target_write_col", "Hedef yazma sütunu (saat)", str),
    ("target_note_col", "Hedef not sütunu", str),
    ("target_status_col", "Hedef durum sütunu (ÇALIŞTI yazan)", str),
    ("target_data_start_row", "Hedef veri başlangıç satırı", int),
    ("max_reasonable_hours", "Anormal değer eşiği (saat)", int),
]

PERIOD_OPTIONS = [
    ("Haftalık (7 gün)", 7),
    ("15 Günlük", 15),
    ("Tek Blok (Aralığın Tamamı)", None),
]

TEAM_NAV_PAGES = [
    ("📁", "Kaynak"),
    ("🔍", "Filtreler"),
    ("📅", "Tarih && Dönem"),
    ("⚙", "Seçenekler"),
    ("📊", "Oluştur && Sonuç"),
]


# ============================================================ ortak yardımcılar

def load_json(path, default):
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def add_to_history(entry, keep=25):
    hist = load_json(TEAM_HISTORY_PATH, [])
    hist.insert(0, entry)
    save_json(TEAM_HISTORY_PATH, hist[:keep])


def normalize_display(s):
    return (s or "").upper().replace("İ", "I").strip()


def open_path(path):
    path = str(path)
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


# ============================================================ hızlı tarih aralıkları (rapor modülü için)

def _week_range(today, weeks_back=0):
    start = today - datetime.timedelta(days=today.weekday() + 7 * weeks_back)
    end = start + datetime.timedelta(days=6)
    return start, end


def _month_range(today, months_back=0):
    year, month = today.year, today.month - months_back
    while month < 1:
        month += 12
        year -= 1
    start = datetime.date(year, month, 1)
    if month == 12:
        end = datetime.date(year, 12, 31)
    else:
        end = datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)
    return start, end


def quick_ranges():
    today = datetime.date.today()
    return [
        ("Bugün", (today, today)),
        ("Bu Hafta", _week_range(today, 0)),
        ("Geçen Hafta", _week_range(today, 1)),
        ("Son 15 Gün", (today - datetime.timedelta(days=14), today)),
        ("Son 30 Gün", (today - datetime.timedelta(days=29), today)),
        ("Bu Ay", _month_range(today, 0)),
        ("Geçen Ay", _month_range(today, 1)),
    ]


# ============================================================ stil (paylaşımlı, karanlık mod destekli)

def build_stylesheet(dark: bool) -> str:
    if dark:
        bg, card, border, text, subtext, field = "#1b1f24", "#242a31", "#33393f", "#eceff1", "#90a4ae", "#2b3238"
        accent_hover = "#33393f"
    else:
        bg, card, border, text, subtext, field = "#f3f5f8", "#ffffff", "#dbe0e6", "#263238", "#607d8b", "#fafbfc"
        accent_hover = "#dbe1e5"
    return f"""
    QMainWindow, QWidget {{ background-color: {bg}; font-family: 'Segoe UI'; font-size: 10pt; color: {text}; }}
    #titleLabel {{ font-size: 19px; font-weight: 700; color: #1a237e; }}
    #suiteHeader {{ background-color: {card}; border-bottom: 1px solid {border}; }}
    #sidebar {{ background-color: {card}; border: none; border-right: 1px solid {border}; }}
    #sidebar::item {{ padding: 14px 10px; border-radius: 8px; margin: 3px 8px; font-weight: 600; }}
    #sidebar::item:selected {{ background-color: #2e7d32; color: white; }}
    #sidebar::item:hover:!selected {{ background-color: {accent_hover}; }}
    #statusBar {{ background-color: {card}; border-top: 1px solid {border}; padding: 6px 14px; }}
    #statusText {{ color: {subtext}; font-size: 9pt; }}
    QGroupBox {{
        background-color: {card}; border: 1px solid {border}; border-radius: 12px;
        margin-top: 16px; padding: 14px; font-weight: 600; color: {text};
    }}
    QGroupBox::title {{ subcontrol-origin: margin; left: 14px; padding: 0 8px; }}
    QLineEdit, QDateEdit, QListWidget, QTextEdit, QComboBox, QTreeWidget {{
        border: 1px solid {border}; border-radius: 8px; padding: 6px; background: {field};
        selection-background-color: #90caf9; color: {text};
    }}
    QTreeWidget::item {{ padding: 3px; }}
    QPushButton {{
        border: none; border-radius: 8px; padding: 9px 16px; background-color: {field};
        color: {text}; font-weight: 600;
    }}
    QPushButton:hover {{ background-color: {accent_hover}; }}
    QPushButton:disabled {{ background-color: {field}; color: #90a4ae; }}
    #primaryButton {{ background-color: #2e7d32; color: white; font-size: 12pt; padding: 12px; }}
    #primaryButton:hover {{ background-color: #276829; }}
    #primaryButton:disabled {{ background-color: #7fa982; color: #eeeeee; }}
    #secondaryButton {{ background-color: #37474f; color: white; padding: 10px; }}
    #secondaryButton:hover {{ background-color: #455a64; }}
    #moduleTab {{
        background-color: {field}; color: {subtext}; font-weight: 700; font-size: 11pt;
        padding: 12px 22px; border-radius: 10px; border: 1px solid {border};
    }}
    #moduleTab:checked {{ background-color: #1a237e; color: white; border: 1px solid #1a237e; }}
    #moduleTab:hover:!checked {{ background-color: {accent_hover}; }}
    QCheckBox {{ font-weight: 500; }}
    QProgressBar {{ border: none; border-radius: 6px; background: {field}; height: 12px; text-align: center; color: {text}; }}
    QProgressBar::chunk {{ background-color: #2e7d32; border-radius: 6px; }}
    """


# ============================================================ sürükle-bırak kutusu (tek/çoklu dosya)

class DropZone(QFrame):
    filesDropped = Signal(list)

    def __init__(self, hint, multiple=False, parent=None):
        super().__init__(parent)
        self.multiple = multiple
        self.setAcceptDrops(True)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet(
            "QFrame { border: 2px dashed #90a4ae; border-radius: 10px; padding: 16px; }"
        )
        layout = QVBoxLayout(self)
        self.label = QLabel(hint)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("border: none; color: #607d8b; font-weight: 600;")
        layout.addWidget(self.label)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(".xlsx"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        paths = [url.toLocalFile() for url in event.mimeData().urls()
                 if url.toLocalFile().lower().endswith(".xlsx")]
        if not paths:
            return
        if not self.multiple:
            paths = paths[:1]
        self.filesDropped.emit(paths)


# ============================================================ hiyerarşik filtre ağacı (rapor modülü)

class HierarchyTree(QTreeWidget):
    """Grup > Bağlı Olduğu Ekip > Meslek Grubu şeklinde 3 seviyeli, onay kutulu,
    üst-alt kademeli (parent işaretlenince çocukları da işaretlenir) ağaç."""

    changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self._updating = False
        self.itemChanged.connect(self._on_item_changed)

    def set_hierarchy(self, hierarchy: dict):
        self._updating = True
        self.clear()
        for team, subteams in hierarchy.items():
            team_item = QTreeWidgetItem([team])
            team_item.setFlags(team_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            team_item.setCheckState(0, Qt.CheckState.Unchecked)
            self.addTopLevelItem(team_item)
            for subteam, roles in subteams.items():
                sub_item = QTreeWidgetItem([subteam])
                sub_item.setFlags(sub_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                sub_item.setCheckState(0, Qt.CheckState.Unchecked)
                team_item.addChild(sub_item)
                for role in roles:
                    role_item = QTreeWidgetItem([role])
                    role_item.setFlags(role_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    role_item.setCheckState(0, Qt.CheckState.Unchecked)
                    sub_item.addChild(role_item)
        self._updating = False
        self.collapseAll()

    def _on_item_changed(self, item, _col):
        if self._updating:
            return
        self._updating = True
        state = item.checkState(0)
        if state != Qt.CheckState.PartiallyChecked:
            self._set_children_state(item, state)
        self._update_ancestors(item.parent())
        self._updating = False
        self.changed.emit()

    def _set_children_state(self, item, state):
        for i in range(item.childCount()):
            child = item.child(i)
            child.setCheckState(0, state)
            self._set_children_state(child, state)

    def _update_ancestors(self, item):
        while item is not None:
            states = {item.child(i).checkState(0) for i in range(item.childCount())}
            if states == {Qt.CheckState.Checked}:
                item.setCheckState(0, Qt.CheckState.Checked)
            elif states == {Qt.CheckState.Unchecked} or not states:
                item.setCheckState(0, Qt.CheckState.Unchecked)
            else:
                item.setCheckState(0, Qt.CheckState.PartiallyChecked)
            item = item.parent()

    def _checked_at_depth(self, depth):
        result = []

        def walk(item, level):
            state = item.checkState(0)
            if level == depth and state != Qt.CheckState.Unchecked:
                result.append((item.text(0), state))
            for i in range(item.childCount()):
                walk(item.child(i), level + 1)

        for i in range(self.topLevelItemCount()):
            walk(self.topLevelItem(i), 0)
        return result

    def checked_teams(self):
        return [t for t, _s in self._checked_at_depth(0)]

    def checked_subteams(self):
        return [t for t, _s in self._checked_at_depth(1)]

    def checked_roles(self):
        return [t for t, s in self._checked_at_depth(2) if s == Qt.CheckState.Checked]

    def set_all_checked(self, checked):
        self._updating = True
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for i in range(self.topLevelItemCount()):
            top = self.topLevelItem(i)
            top.setCheckState(0, state)
            self._set_children_state(top, state)
        self._updating = False
        self.changed.emit()

    def filter_text(self, query):
        query = (query or "").strip().lower()

        def walk(item):
            child_visible = False
            for i in range(item.childCount()):
                if walk(item.child(i)):
                    child_visible = True
            self_match = query in item.text(0).lower()
            visible = self_match or child_visible
            item.setHidden(not visible)
            if child_visible and query:
                item.setExpanded(True)
            return visible

        for i in range(self.topLevelItemCount()):
            walk(self.topLevelItem(i))


# ============================================================ manuel eşleştirme (puantaj modülü)

class ManualMatchDialog(QDialog):
    """Eşleşmeyen hedef satırları için isimle öneri + manuel arama ile
    kaynak TC ataması yapılmasını sağlayan pencere."""

    def __init__(self, parent, unmatched, name_suggestions, source_names, file_label):
        super().__init__(parent)
        self.setWindowTitle(f"Manuel Eşleştirme - {file_label}")
        self.resize(860, 560)

        self.unmatched = unmatched
        self.name_suggestions = name_suggestions
        self.all_source = sorted(source_names.items(), key=lambda kv: (kv[1] or ""))
        self.mapping = {}
        self.current_row = None
        self.result_mapping = {}

        self._build()
        self._refresh_list()

    def _build(self):
        outer = QVBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        outer.addWidget(splitter, stretch=1)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_label = QLabel("Eşleşmeyen satırlar:")
        left_label.setStyleSheet("font-weight:700;")
        left_layout.addWidget(left_label)
        self.listbox = QListWidget()
        self.listbox.currentRowChanged.connect(self._on_select)
        left_layout.addWidget(self.listbox)
        splitter.addWidget(left)

        right = QWidget()
        right_layout = QVBoxLayout(right)

        self.selected_label = QLabel("Bir satır seç →")
        self.selected_label.setStyleSheet("font-weight:700;")
        self.selected_label.setWordWrap(True)
        right_layout.addWidget(self.selected_label)

        self.suggestion_box = QGroupBox("Önerilen eşleşmeler")
        self.suggestion_layout = QVBoxLayout()
        self.suggestion_box.setLayout(self.suggestion_layout)
        right_layout.addWidget(self.suggestion_box)

        search_box = QGroupBox("Manuel ara (tüm kaynak personeli)")
        search_layout = QVBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("İsim yazmaya başla...")
        self.search_edit.textChanged.connect(self._filter_search)
        search_layout.addWidget(self.search_edit)
        self.search_list = QListWidget()
        search_layout.addWidget(self.search_list, stretch=1)
        assign_btn = QPushButton("Seçileni bu satıra ata")
        assign_btn.clicked.connect(self._assign_from_search)
        row = QHBoxLayout()
        row.addStretch()
        row.addWidget(assign_btn)
        search_layout.addLayout(row)
        search_box.setLayout(search_layout)
        right_layout.addWidget(search_box, stretch=1)

        splitter.addWidget(right)
        splitter.setSizes([320, 540])

        bottom = QHBoxLayout()
        bottom.addStretch()
        cancel_btn = QPushButton("İptal")
        cancel_btn.clicked.connect(self._cancel)
        finish_btn = QPushButton("Tamamla ve Devam Et")
        finish_btn.setObjectName("primaryButton")
        finish_btn.clicked.connect(self._finish)
        bottom.addWidget(cancel_btn)
        bottom.addWidget(finish_btn)
        outer.addLayout(bottom)

        self._search_pool = [(f"{name}  (TC: {tc})", tc) for tc, name in self.all_source]
        for label, _tc in self._search_pool:
            self.search_list.addItem(label)

    def _refresh_list(self):
        self.listbox.blockSignals(True)
        self.listbox.clear()
        self._row_by_index = []
        for row, tc, name in self.unmatched:
            mark = "✓ " if row in self.mapping else "    "
            self.listbox.addItem(f"{mark}satır {row}: {name or '(isim yok)'}  (TC hedef: {tc})")
            self._row_by_index.append(row)
        self.listbox.blockSignals(False)

    def _on_select(self, idx):
        if idx < 0 or idx >= len(self._row_by_index):
            return
        row = self._row_by_index[idx]
        self.current_row = row
        _, tc, name = next(u for u in self.unmatched if u[0] == row)
        current = f"  →  atanan TC: {self.mapping[row]}" if row in self.mapping else ""
        self.selected_label.setText(f"Satır {row}: {name or '(isim yok)'}  (hedef TC: {tc}){current}")

        while self.suggestion_layout.count():
            item = self.suggestion_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        suggestions = self.name_suggestions.get(row, [])
        if not suggestions:
            self.suggestion_layout.addWidget(QLabel("Uygun bir isim önerisi bulunamadı."))
        for s_tc, s_name, score in suggestions:
            line = QWidget()
            h = QHBoxLayout(line)
            h.setContentsMargins(0, 0, 0, 0)
            h.addWidget(QLabel(f"{s_name}  (TC: {s_tc})  — benzerlik %{int(score * 100)}"))
            h.addStretch()
            btn = QPushButton("Ata")
            btn.clicked.connect(lambda _checked=False, t=s_tc: self._assign(row, t))
            h.addWidget(btn)
            self.suggestion_layout.addWidget(line)

    def _assign(self, row, tc):
        self.mapping[row] = tc
        self._refresh_list()
        if row in self._row_by_index:
            self.listbox.setCurrentRow(self._row_by_index.index(row))

    def _assign_from_search(self):
        items = self.search_list.selectedItems()
        if not items or self.current_row is None:
            return
        label = items[0].text()
        tc = dict(self._search_pool).get(label)
        if tc:
            self._assign(self.current_row, tc)

    def _filter_search(self, text):
        query = normalize_display(text)
        self.search_list.clear()
        for label, _tc in self._search_pool:
            if query in normalize_display(label):
                self.search_list.addItem(label)

    def _finish(self):
        self.result_mapping = dict(self.mapping)
        self.accept()

    def _cancel(self):
        self.result_mapping = {}
        self.reject()


def page_title(text):
    lbl = QLabel(text)
    lbl.setStyleSheet("font-size:15px; font-weight:700; color:#1a237e;")
    return lbl


# ============================================================ MODÜL 1: PUANTAJ AKTARIMI

class PuantajModule(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.target_paths = []
        self.output_paths = []
        self._log_buffer = []

        saved = load_json(PUANTAJ_SETTINGS_PATH, {})
        self.cfg = dict(engine.DEFAULT_CONFIG)
        self.cfg.update({k: v for k, v in saved.items() if k in engine.DEFAULT_CONFIG})

        self._build_ui()
        self._autodetect_files()

    # ------------------------------------------------------------ UI iskeleti
    def _build_ui(self):
        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.nav = QListWidget()
        self.nav.setObjectName("sidebar")
        self.nav.setFixedWidth(190)
        self.nav.setFrameShape(QFrame.Shape.NoFrame)
        for icon, label in [("📁", "Kaynak && Hedef"), ("📅", "Tarih"),
                            ("⚙", "Ayarlar"), ("📊", "Aktar && Sonuç")]:
            self.nav.addItem(QListWidgetItem(f"{icon}   {label.replace('&&', '&')}"))
        self.nav.currentRowChanged.connect(lambda row: self.pages.setCurrentIndex(max(row, 0)))
        outer.addWidget(self.nav)

        right = QVBoxLayout()
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(0)

        self.pages = QStackedWidget()
        self.pages.addWidget(self._build_source_target_page())
        self.pages.addWidget(self._build_date_page())
        self.pages.addWidget(self._build_settings_page())
        self.pages.addWidget(self._build_result_page())
        right.addWidget(self.pages, stretch=1)

        self.status_bar_widget = QFrame()
        self.status_bar_widget.setObjectName("statusBar")
        sb = QHBoxLayout(self.status_bar_widget)
        sb.setContentsMargins(14, 6, 14, 6)
        self.status_label = QLabel("Kaynak yüklenmedi.")
        self.status_label.setObjectName("statusText")
        sb.addWidget(self.status_label)
        sb.addStretch()
        right.addWidget(self.status_bar_widget)

        right_widget = QWidget()
        right_widget.setLayout(right)
        outer.addWidget(right_widget, stretch=1)

        self.nav.setCurrentRow(0)

    # ------------------------------------------------------------ Sayfa 1: Kaynak & Hedef
    def _build_source_target_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)
        layout.addWidget(page_title("Kaynak ve Hedef Dosyalar"))

        source_group = QGroupBox("Kaynak (aylık puantaj)")
        source_layout = QVBoxLayout()
        self.source_drop = DropZone("📄  Kaynak excel'i buraya sürükle-bırak", multiple=False)
        self.source_drop.filesDropped.connect(lambda paths: self._set_source(paths[0]))
        source_layout.addWidget(self.source_drop)
        row = QHBoxLayout()
        self.source_edit = QLineEdit()
        self.source_edit.setReadOnly(True)
        row.addWidget(self.source_edit, stretch=1)
        browse_source_btn = QPushButton("Gözat...")
        browse_source_btn.clicked.connect(self.browse_source)
        row.addWidget(browse_source_btn)
        source_layout.addLayout(row)
        source_group.setLayout(source_layout)
        layout.addWidget(source_group)

        target_group = QGroupBox("Hedef dosya(lar) (günlük liste)")
        target_layout = QVBoxLayout()
        self.target_drop = DropZone("📄  Hedef excel(ler)i buraya sürükle-bırak", multiple=True)
        self.target_drop.filesDropped.connect(self._add_targets)
        target_layout.addWidget(self.target_drop)

        list_row = QHBoxLayout()
        self.target_list = QListWidget()
        self.target_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        list_row.addWidget(self.target_list, stretch=1)
        target_btns = QVBoxLayout()
        add_btn = QPushButton("Dosya Ekle...")
        add_btn.clicked.connect(self.browse_target_add)
        remove_btn = QPushButton("Seçileni Çıkar")
        remove_btn.clicked.connect(self.remove_selected_target)
        detect_btn = QPushButton("TARGET Klasöründen Algıla")
        detect_btn.clicked.connect(self._autodetect_files)
        for b in (add_btn, remove_btn, detect_btn):
            target_btns.addWidget(b)
        target_btns.addStretch()
        list_row.addLayout(target_btns)
        target_layout.addLayout(list_row)
        target_group.setLayout(target_layout)
        layout.addWidget(target_group, stretch=1)

        return page

    # ------------------------------------------------------------ Sayfa 2: Tarih
    def _build_date_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)
        layout.addWidget(page_title("Aktarım Tarihi"))

        date_group = QGroupBox("Tarih")
        date_layout = QHBoxLayout()
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("dd.MM.yyyy")
        self.date_edit.dateChanged.connect(self._update_status_bar)
        date_layout.addWidget(self.date_edit)

        dun_btn = QPushButton("Dün")
        dun_btn.clicked.connect(lambda: self.date_edit.setDate(QDate.currentDate().addDays(-1)))
        date_layout.addWidget(dun_btn)
        today_btn = QPushButton("Bugün")
        today_btn.clicked.connect(lambda: self.date_edit.setDate(QDate.currentDate()))
        date_layout.addWidget(today_btn)
        yarin_btn = QPushButton("Yarın")
        yarin_btn.clicked.connect(lambda: self.date_edit.setDate(QDate.currentDate().addDays(1)))
        date_layout.addWidget(yarin_btn)
        date_layout.addStretch()
        date_group.setLayout(date_layout)
        layout.addWidget(date_group)

        layout.addStretch()
        return page

    # ------------------------------------------------------------ Sayfa 3: Ayarlar
    def _build_settings_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)
        layout.addWidget(page_title("Ayarlar"))

        opt_group = QGroupBox("Rapor İçeriği")
        opt_layout = QHBoxLayout()
        self.summary_checkbox = QCheckBox("Çıktıya Özet sayfası ekle")
        self.summary_checkbox.setChecked(self.cfg.get("write_summary_sheet", True))
        opt_layout.addWidget(self.summary_checkbox)
        opt_layout.addStretch()
        opt_group.setLayout(opt_layout)
        layout.addWidget(opt_group)

        adv_group = QGroupBox("Gelişmiş Ayarlar (kaydedilir)")
        adv_layout = QGridLayout()
        self.adv_edits = {}
        for i, (key, label, _type) in enumerate(PUANTAJ_ADVANCED_FIELDS):
            r, c = divmod(i, 2)
            adv_layout.addWidget(QLabel(label + ":"), r, c * 2)
            edit = QLineEdit(str(self.cfg[key]))
            adv_layout.addWidget(edit, r, c * 2 + 1)
            self.adv_edits[key] = edit
        adv_group.setLayout(adv_layout)
        layout.addWidget(adv_group)

        layout.addStretch()
        return page

    # ------------------------------------------------------------ Sayfa 4: Aktar & Sonuç
    def _build_result_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)
        layout.addWidget(page_title("Aktar ve Sonuç"))

        action_layout = QHBoxLayout()
        self.preview_btn = QPushButton("ÖNİZLE")
        self.preview_btn.setObjectName("secondaryButton")
        self.preview_btn.clicked.connect(self.on_preview)
        self.run_btn = QPushButton("AKTAR")
        self.run_btn.setObjectName("primaryButton")
        self.run_btn.clicked.connect(self.on_run)
        self.open_output_btn = QPushButton("Çıktı Klasörünü Aç")
        self.open_output_btn.setEnabled(False)
        self.open_output_btn.clicked.connect(self.open_output_folder)
        action_layout.addWidget(self.preview_btn, stretch=1)
        action_layout.addWidget(self.run_btn, stretch=1)
        action_layout.addWidget(self.open_output_btn)
        layout.addLayout(action_layout)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        result_group = QGroupBox("Sonuç")
        result_layout = QVBoxLayout()
        self.badges_widget = QWidget()
        self.badges_layout = QHBoxLayout(self.badges_widget)
        self.badges_layout.setContentsMargins(0, 0, 0, 6)
        result_layout.addWidget(self.badges_widget)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        result_layout.addWidget(self.log, stretch=1)
        result_group.setLayout(result_layout)
        layout.addWidget(result_group, stretch=1)

        return page

    # ------------------------------------------------------------ durum çubuğu
    def _update_status_bar(self):
        src_name = Path(self.source_edit.text()).name if self.source_edit.text() else "kaynak yok"
        n_targets = len(self.target_paths)
        date_txt = self.date_edit.date().toString("dd.MM.yyyy")
        self.status_label.setText(
            f"📁 {src_name}   ·   🎯 {n_targets} hedef dosya   ·   📅 {date_txt}"
        )

    # ------------------------------------------------------------ dosyalar
    def _autodetect_files(self):
        try:
            src = engine.find_single_xlsx(SOURCE_DIR, "SOURCE")
            self.source_edit.setText(str(src))
        except engine.TransferError:
            self.source_edit.setText("")

        self.target_paths = list(engine.find_all_xlsx(TARGET_DIR, "TARGET", allow_empty=True))
        self._refresh_target_list()
        self._update_status_bar()

    def _refresh_target_list(self):
        self.target_list.clear()
        for p in self.target_paths:
            self.target_list.addItem(p.name)

    def _set_source(self, path):
        self.source_edit.setText(str(path))
        self._update_status_bar()

    def browse_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "Kaynak Excel Seç", "", "Excel dosyaları (*.xlsx)")
        if path:
            self._set_source(path)

    def _add_targets(self, paths):
        for p in paths:
            pp = Path(p)
            if pp not in self.target_paths:
                self.target_paths.append(pp)
        self._refresh_target_list()
        self._update_status_bar()

    def browse_target_add(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Hedef Excel(ler) Seç", "", "Excel dosyaları (*.xlsx)")
        if paths:
            self._add_targets(paths)

    def remove_selected_target(self):
        rows = sorted((idx.row() for idx in self.target_list.selectedIndexes()), reverse=True)
        for r in rows:
            del self.target_paths[r]
        self._refresh_target_list()
        self._update_status_bar()

    # ------------------------------------------------------------ ayarlar
    def _current_cfg(self):
        cfg = dict(engine.DEFAULT_CONFIG)
        for key, _label, typ in PUANTAJ_ADVANCED_FIELDS:
            raw = self.adv_edits[key].text().strip()
            cfg[key] = typ(raw)
        cfg["write_summary_sheet"] = self.summary_checkbox.isChecked()
        return cfg

    def _persist_settings(self, cfg):
        data = {k: cfg[k] for k, _l, _t in PUANTAJ_ADVANCED_FIELDS}
        data["write_summary_sheet"] = cfg.get("write_summary_sheet", True)
        save_json(PUANTAJ_SETTINGS_PATH, data)

    def save_state(self):
        try:
            self._persist_settings(self._current_cfg())
        except Exception:
            pass

    def _current_date_str(self):
        return self.date_edit.date().toString("yyyy-MM-dd")

    def _validate_inputs(self):
        source_path = self.source_edit.text()
        if not source_path or not Path(source_path).exists():
            QMessageBox.critical(self, "Eksik dosya", "Geçerli bir kaynak excel dosyası seçmelisin.")
            return None, None, None
        if not self.target_paths:
            QMessageBox.critical(self, "Eksik dosya", "En az bir hedef excel dosyası eklemelisin.")
            return None, None, None
        try:
            cfg = self._current_cfg()
            target_date_str = self._current_date_str()
        except ValueError as e:
            QMessageBox.critical(self, "Geçersiz ayar", f"Bir alan sayısal olmalı: {e}")
            return None, None, None
        return source_path, cfg, target_date_str

    # ------------------------------------------------------------ log / rozet
    def _clear_log(self):
        self._log_buffer = []
        self._render_log()

    def _log(self, text, color=None, bold=False):
        self._log_buffer.append((text, color, bold))
        self._render_log()

    def _render_log(self):
        parts = []
        for text, color, bold in self._log_buffer:
            escaped = html_lib.escape(text).replace("\n", "<br>")
            styles = []
            if color:
                styles.append(f"color:{TAG_COLORS.get(color, color)}")
            if bold:
                styles.append("font-weight:700")
            style_attr = f' style="{";".join(styles)}"' if styles else ""
            parts.append(f"<span{style_attr}>{escaped}</span>")
        body = "".join(parts)
        self.log.setHtml(f"<div style='font-family:Consolas,monospace; font-size:10pt;'>{body}</div>")
        self.log.moveCursor(QTextCursor.MoveOperation.End)

    def _update_badges(self, items):
        while self.badges_layout.count():
            item = self.badges_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        for label_text, value, color in items:
            lbl = QLabel(f"{label_text}: {value}")
            lbl.setStyleSheet(
                f"background-color:{color}; color:white; border-radius:11px; "
                f"padding:5px 14px; font-weight:700; font-size:10pt;"
            )
            self.badges_layout.addWidget(lbl)
        self.badges_layout.addStretch()

    def _log_result(self, result, title):
        self._log(f"\n=== {title} ===\n", bold=True)
        for w in result["warnings"]:
            self._log(f"UYARI: {w}\n", "orange")
        self._log("Eşleşen: ")
        self._log(f"{len(result['matched'])}", "green", bold=True)
        if result.get("manual_matched"):
            self._log(f"  (manuel: {len(result['manual_matched'])})", "purple")
        self._log("\n")
        self._log(f"Eşleşmeyen / kırmızı işaretli: {len(result['flagged'])}\n", "red")

        if result["unmatched"]:
            self._log("Eşleşmeyen satırlar: ", "red")
            self._log(", ".join(f"satır {r} ({name or tc})" for r, tc, name in result["unmatched"][:15]) +
                      (" ..." if len(result["unmatched"]) > 15 else "") + "\n")

        if result["on_leave"]:
            self._log(f"İzinli/raporlu görünen ama eşleşen: {len(result['on_leave'])}\n", "orange")

        if result["notes_written"]:
            self._log("Açıklama notu aktarılan:\n", "blue")
            for r, tc, name, note in result["notes_written"]:
                self._log(f"  satır {r}: {name or tc} -> \"{note}\"\n")

        if result.get("leave_marked"):
            self._log(f"İzinli işaretlenip SARIYA boyanan: {len(result['leave_marked'])} kişi\n", "orange")
            for r, tc, name, label in result["leave_marked"]:
                self._log(f"  satır {r}: {name or tc} -> {label}\n")

        if result["abnormal"]:
            self._log(f"ANORMAL DEĞER ({len(result['abnormal'])} kişi):\n", "red")
            for r, tc, name, val in result["abnormal"]:
                self._log(f"  satır {r}: {name or tc} -> {val} saat\n")

        if result["tc_issues"]:
            self._log(f"TC format/checksum sorunu ({len(result['tc_issues'])} kayıt):\n", "orange")
            for issue in result["tc_issues"][:15]:
                self._log(f"  [{issue['side']}] satır {issue['row']}: TC {issue['tc']} "
                          f"({issue['name']}) -> {issue['reason']}\n")

        if result["source_not_in_target"]:
            self._log(f"Kaynakta var, hedefte yok: {len(result['source_not_in_target'])} kişi\n", "purple")
            for tc, name, val, leave in result["source_not_in_target"][:15]:
                extra = f", kod: {leave[0][1]}" if leave else ""
                self._log(f"  {name} (TC {tc}) -> {val} saat{extra}\n")
            if len(result["source_not_in_target"]) > 15:
                self._log(f"  ... ve {len(result['source_not_in_target']) - 15} kişi daha\n")

        if result["odd_row_counts"]:
            self._log(f"UYARI: beklenmeyen satır sayısı olan TC'ler: {len(result['odd_row_counts'])}\n", "orange")

        if result.get("output_path"):
            self._log(f"Çıktı: {result['output_path']}\n", bold=True)

    # ------------------------------------------------------------ önizleme
    def on_preview(self):
        source_path, cfg, target_date_str = self._validate_inputs()
        if not source_path:
            return
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            self._clear_log()
            self._log(f"ÖNİZLEME - {target_date_str}\n", bold=True)
            total_matched = total_unmatched = total_flagged = 0
            for target_path in self.target_paths:
                try:
                    result = engine.run_transfer(cfg, Path(source_path), target_path,
                                                  output_path=None, target_date_str=target_date_str, dry_run=True)
                    self._log_result(result, f"Önizleme: {target_path.name}")
                    total_matched += len(result["matched"])
                    total_unmatched += len(result["unmatched"])
                    total_flagged += len(result["flagged"])
                except engine.TransferError as e:
                    self._log(f"\nHATA ({target_path.name}): {e}\n", "red")
                QApplication.processEvents()
            self._log("\nBu bir önizlemedir, hiçbir dosya değiştirilmedi.\n", bold=True)
            self._update_badges([
                ("Eşleşen", total_matched, "#2e7d32"),
                ("Eşleşmeyen", total_unmatched, "#b71c1c"),
                ("İşaretli", total_flagged, "#e65100"),
            ])
        finally:
            QApplication.restoreOverrideCursor()

    # ------------------------------------------------------------ gerçek çalıştırma
    def on_run(self):
        source_path, cfg, target_date_str = self._validate_inputs()
        if not source_path:
            return

        self._persist_settings(cfg)
        self.run_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        self.open_output_btn.setEnabled(False)
        self.progress.setVisible(True)
        self._clear_log()
        self.output_paths = []

        totals = {"matched": 0, "flagged": 0, "leave": 0, "abnormal": 0, "issues": 0}
        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            for target_path in self.target_paths:
                QApplication.restoreOverrideCursor()
                result = self._process_one_file(cfg, Path(source_path), target_path, target_date_str)
                QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
                if result:
                    totals["matched"] += len(result["matched"])
                    totals["flagged"] += len(result["flagged"])
                    totals["leave"] += len(result.get("leave_marked", []))
                    totals["abnormal"] += len(result["abnormal"])
                    totals["issues"] += len(result["tc_issues"])
                QApplication.processEvents()
        finally:
            QApplication.restoreOverrideCursor()
            self.progress.setVisible(False)
            self.run_btn.setEnabled(True)
            self.preview_btn.setEnabled(True)
            if self.output_paths:
                self.open_output_btn.setEnabled(True)

        self._update_badges([
            ("Eşleşen", totals["matched"], "#2e7d32"),
            ("Kırmızı", totals["flagged"], "#b71c1c"),
            ("İzinli (sarı)", totals["leave"], "#f9a825"),
            ("Anormal", totals["abnormal"], "#e65100"),
            ("TC sorunu", totals["issues"], "#6a1b9a"),
        ])

        QMessageBox.information(self, "Tamamlandı",
                                 f"{len(self.output_paths)} dosya işlendi.\nDetaylar sonuç ekranında.")

    def _process_one_file(self, cfg, source_path, target_path, target_date_str):
        try:
            preview = engine.run_transfer(cfg, source_path, target_path,
                                           output_path=None, target_date_str=target_date_str, dry_run=True)
        except engine.TransferError as e:
            self._log(f"\nHATA ({target_path.name}): {e}\n", "red")
            return None

        manual_mapping = {}
        if preview["unmatched"]:
            answer = QMessageBox.question(
                self, "Eşleşmeyen personel bulundu",
                f"{target_path.name} dosyasında {len(preview['unmatched'])} eşleşmeyen satır var.\n\n"
                f"İsimle manuel eşleştirme yapmak ister misin?\n"
                f"(Hayır dersen bu satırlar kırmızı işaretlenip geçilir.)",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if answer == QMessageBox.StandardButton.Yes:
                dialog = ManualMatchDialog(self, preview["unmatched"], preview["name_suggestions"],
                                            preview["source_names"], target_path.name)
                dialog.exec()
                manual_mapping = dialog.result_mapping or {}

        output_path = OUTPUT_DIR / target_path.name
        result = engine.run_transfer(cfg, source_path, target_path, output_path,
                                      target_date_str, manual_mapping=manual_mapping, dry_run=False)
        self.output_paths.append(result["output_path"])
        self._log_result(result, target_path.name)
        return result

    # ------------------------------------------------------------ diğer
    def open_output_folder(self):
        open_path(OUTPUT_DIR)


# ============================================================ MODÜL 2: FAZLA MESAİ RAPORU

class TeamReportModule(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.source_path = None
        self.output_path = None
        self._hierarchy = {}
        self._log_buffer = []
        self.settings = load_json(TEAM_SETTINGS_PATH, {})

        self._build_ui()
        self._restore_settings()
        self.autodetect_source()

    # ================================================================== UI iskeleti
    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # --- gövde: sol menü + sayfalar ---
        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)

        self.nav = QListWidget()
        self.nav.setObjectName("sidebar")
        self.nav.setFixedWidth(210)
        self.nav.setFrameShape(QFrame.Shape.NoFrame)
        for icon, label in TEAM_NAV_PAGES:
            self.nav.addItem(QListWidgetItem(f"{icon}   {label.replace('&&', '&')}"))
        self.nav.currentRowChanged.connect(self._on_nav_changed)
        body.addWidget(self.nav)

        self.pages = QStackedWidget()
        self.pages.addWidget(self._build_source_page())
        self.pages.addWidget(self._build_filter_page())
        self.pages.addWidget(self._build_date_page())
        self.pages.addWidget(self._build_options_page())
        self.pages.addWidget(self._build_result_page())
        body.addWidget(self.pages, stretch=1)

        body_widget = QWidget()
        body_widget.setLayout(body)
        outer.addWidget(body_widget, stretch=1)

        # --- alt canlı özet çubuğu (her zaman görünür) ---
        self.status_bar_widget = QFrame()
        self.status_bar_widget.setObjectName("statusBar")
        sb_layout = QHBoxLayout(self.status_bar_widget)
        sb_layout.setContentsMargins(14, 6, 14, 6)
        self.status_label = QLabel("Kaynak yüklenmedi.")
        self.status_label.setObjectName("statusText")
        sb_layout.addWidget(self.status_label)
        sb_layout.addStretch()
        outer.addWidget(self.status_bar_widget)

        self.nav.setCurrentRow(0)

    def _on_nav_changed(self, row):
        if row >= 0:
            self.pages.setCurrentIndex(row)

    # ================================================================== SAYFA 1: Kaynak
    def _build_source_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        layout.addWidget(self._page_title("Kaynak Excel"))

        self.drop_zone = DropZone("📄  Kaynak excel'i buraya sürükle-bırak", multiple=False)
        self.drop_zone.filesDropped.connect(lambda paths: self._load_source(Path(paths[0])) if paths else None)
        layout.addWidget(self.drop_zone)

        row = QHBoxLayout()
        self.source_edit = QLineEdit()
        self.source_edit.setReadOnly(True)
        row.addWidget(self.source_edit, stretch=1)
        browse_btn = QPushButton("Gözat...")
        browse_btn.clicked.connect(self.browse_source)
        row.addWidget(browse_btn)
        detect_btn = QPushButton("SOURCE Klasöründen Algıla")
        detect_btn.clicked.connect(self.autodetect_source)
        row.addWidget(detect_btn)
        layout.addLayout(row)

        info_group = QGroupBox("Kaynak Bilgisi")
        info_layout = QVBoxLayout()
        self.source_info_label = QLabel("Henüz bir kaynak dosya yüklenmedi.")
        self.source_info_label.setWordWrap(True)
        info_layout.addWidget(self.source_info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        layout.addStretch()
        return page

    # ================================================================== SAYFA 2: Filtreler
    def _build_filter_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(10)

        layout.addWidget(self._page_title("Filtreler"))

        source_group = QGroupBox("Personel Kaynağı")
        source_layout = QHBoxLayout()
        self.person_source_group = QButtonGroup(self)
        self.radio_source_excel = QRadioButton("📄 Kaynak Excel (bu ayın grup/ekip/meslek bilgisi)")
        self.radio_source_excel.setChecked(True)
        self.radio_source_dataset = QRadioButton("🗂️ Veri Seti (projenin kalıcı personel veritabanı)")
        self.person_source_group.addButton(self.radio_source_excel, 0)
        self.person_source_group.addButton(self.radio_source_dataset, 1)
        self.radio_source_excel.toggled.connect(self._on_person_source_changed)
        source_layout.addWidget(self.radio_source_excel)
        source_layout.addWidget(self.radio_source_dataset)
        source_layout.addStretch()
        source_group.setLayout(source_layout)
        layout.addWidget(source_group)

        hint = QLabel("Bir grubu işaretlemek altındaki tüm alt ekip ve meslekleri de işaretler. "
                      "Hiçbir şey işaretlenmezse tümü dahil edilir.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#78909c; font-size:9pt;")
        layout.addWidget(hint)

        search_row = QHBoxLayout()
        self.filter_search = QLineEdit()
        self.filter_search.setPlaceholderText("🔎  Grup / ekip / meslek ara...")
        self.filter_search.textChanged.connect(lambda t: self.tree.filter_text(t))
        search_row.addWidget(self.filter_search)
        all_btn = QPushButton("Tümünü Seç")
        all_btn.clicked.connect(lambda: self.tree.set_all_checked(True))
        none_btn = QPushButton("Temizle")
        none_btn.clicked.connect(lambda: self.tree.set_all_checked(False))
        search_row.addWidget(all_btn)
        search_row.addWidget(none_btn)
        layout.addLayout(search_row)

        self.tree = HierarchyTree()
        self.tree.changed.connect(self._update_status_bar)
        layout.addWidget(self.tree, stretch=1)

        self.filter_summary = QLabel()
        self.filter_summary.setWordWrap(True)
        self.filter_summary.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(self.filter_summary)

        return page

    def _on_person_source_changed(self, _checked=None):
        if self.radio_source_dataset.isChecked():
            self.tree.set_hierarchy(db.get_hierarchy())
            self.filter_search.setPlaceholderText("🔎  Veri setinde grup / ekip / meslek ara...")
        else:
            self.tree.set_hierarchy(self._hierarchy)
            self.filter_search.setPlaceholderText("🔎  Grup / ekip / meslek ara...")
        self.filter_search.clear()
        self._update_status_bar()

    # ================================================================== SAYFA 3: Tarih & Dönem
    def _build_date_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        layout.addWidget(self._page_title("Tarih Aralığı ve Dönem"))

        quick_group = QGroupBox("Hızlı Seçim")
        quick_layout = QHBoxLayout()
        for label, (start, end) in quick_ranges():
            btn = QPushButton(label)
            btn.clicked.connect(lambda _checked=False, s=start, e=end: self._set_date_range(s, e))
            quick_layout.addWidget(btn)
        quick_group.setLayout(quick_layout)
        layout.addWidget(quick_group)

        range_group = QGroupBox("Tarih Aralığı")
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel("Başlangıç:"))
        self.date_start_edit = QDateEdit(QDate.currentDate())
        self.date_start_edit.setCalendarPopup(True)
        self.date_start_edit.setDisplayFormat("dd.MM.yyyy")
        self.date_start_edit.dateChanged.connect(self._update_status_bar)
        range_layout.addWidget(self.date_start_edit)

        range_layout.addWidget(QLabel("Bitiş:"))
        self.date_end_edit = QDateEdit(QDate.currentDate())
        self.date_end_edit.setCalendarPopup(True)
        self.date_end_edit.setDisplayFormat("dd.MM.yyyy")
        self.date_end_edit.dateChanged.connect(self._update_status_bar)
        range_layout.addWidget(self.date_end_edit)
        range_layout.addStretch()
        range_group.setLayout(range_layout)
        layout.addWidget(range_group)

        period_group = QGroupBox("Dönem Uzunluğu")
        period_layout = QHBoxLayout()
        period_layout.addWidget(QLabel("Dönem:"))
        self.period_combo = QComboBox()
        for label, _val in PERIOD_OPTIONS:
            self.period_combo.addItem(label)
        self.period_combo.currentIndexChanged.connect(self._update_status_bar)
        period_layout.addWidget(self.period_combo)
        period_layout.addStretch()
        period_group.setLayout(period_layout)
        layout.addWidget(period_group)

        layout.addStretch()
        return page

    # ================================================================== SAYFA 4: Seçenekler
    def _build_options_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        layout.addWidget(self._page_title("Ek Seçenekler"))

        opt_group = QGroupBox("Rapor İçeriği")
        opt_layout = QVBoxLayout()
        self.employee_list_checkbox = QCheckBox("Personel listesi ekle (isim + görev + günlük kırılım)")
        self.employee_list_checkbox.setChecked(True)
        opt_layout.addWidget(self.employee_list_checkbox)
        self.compare_checkbox = QCheckBox("Önceki eşit uzunluktaki dönemle karşılaştır")
        self.compare_checkbox.setChecked(True)
        opt_layout.addWidget(self.compare_checkbox)
        opt_group.setLayout(opt_layout)
        layout.addWidget(opt_group)

        quality_group = QGroupBox("Geçmiş Ay / Veri Kalitesi (düzensiz kaynak dosyalar için)")
        quality_layout = QVBoxLayout()
        self.dataset_correct_checkbox = QCheckBox(
            "Grup/Alt Ekip/Görev bilgisini Veri Seti'nden düzelt (kaynaktaki eksik/hatalı etiketleri onarır)")
        quality_layout.addWidget(self.dataset_correct_checkbox)
        self.history_checkbox = QCheckBox(
            "Rapora 'Personel Değişiklikleri' sayfası ekle (Veri Seti geçmişinden)")
        self.history_checkbox.setChecked(True)
        quality_layout.addWidget(self.history_checkbox)
        quality_hint = QLabel("İşaretlenirse rapora ayrıca bir 'Veri Kalitesi Kontrolü' sayfası eklenir: "
                              "eksik alt ekip/görev, hatalı TC ve Veri Seti ile çelişen kayıtları gösterir.")
        quality_hint.setWordWrap(True)
        quality_hint.setStyleSheet("color:#78909c; font-size:9pt;")
        quality_layout.addWidget(quality_hint)
        quality_group.setLayout(quality_layout)
        layout.addWidget(quality_group)

        threshold_group = QGroupBox("Kritik Fazla Mesai Eşiği (Aylık)")
        threshold_layout = QVBoxLayout()
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Aylık bu saati aşan personelin fazla mesaisi kritik sayılır (kırmızı işaretlenir):"))
        self.threshold_spin = QDoubleSpinBox()
        self.threshold_spin.setRange(0, 500)
        self.threshold_spin.setDecimals(1)
        self.threshold_spin.setSingleStep(1.0)
        self.threshold_spin.setValue(float(report.CONFIG.get("overtime_threshold", 52.0)))
        self.threshold_spin.setSuffix(" saat / ay")
        row1.addWidget(self.threshold_spin)
        row1.addStretch()
        threshold_layout.addLayout(row1)
        self.threshold_effective_label = QLabel()
        self.threshold_effective_label.setStyleSheet("color:#78909c; font-size:9pt;")
        threshold_layout.addWidget(self.threshold_effective_label)
        self.threshold_spin.valueChanged.connect(self._update_threshold_effective_label)
        threshold_group.setLayout(threshold_layout)
        layout.addWidget(threshold_group)

        layout.addStretch()
        return page

    # ================================================================== SAYFA 5: Oluştur & Sonuç
    def _build_result_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        layout.addWidget(self._page_title("Oluştur ve Sonuç"))

        action_layout = QHBoxLayout()
        self.preview_btn = QPushButton("ÖNİZLE  (Ctrl+P)")
        self.preview_btn.setObjectName("secondaryButton")
        self.preview_btn.clicked.connect(self.on_preview)
        self.generate_btn = QPushButton("RAPOR OLUŞTUR  (Ctrl+G)")
        self.generate_btn.setObjectName("primaryButton")
        self.generate_btn.clicked.connect(self.on_generate)
        action_layout.addWidget(self.preview_btn, stretch=1)
        action_layout.addWidget(self.generate_btn, stretch=1)
        layout.addLayout(action_layout)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFormat("Hazır")
        layout.addWidget(self.progress)

        self.badges_widget = QWidget()
        self.badges_layout = QHBoxLayout(self.badges_widget)
        self.badges_layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.badges_widget)

        out_row = QHBoxLayout()
        self.open_output_btn = QPushButton("Çıktı Klasörünü Aç")
        self.open_output_btn.setEnabled(False)
        self.open_output_btn.clicked.connect(self.open_output_folder)
        self.open_file_btn = QPushButton("Raporu Excel'de Aç")
        self.open_file_btn.setEnabled(False)
        self.open_file_btn.clicked.connect(self.open_output_file)
        out_row.addWidget(self.open_output_btn)
        out_row.addWidget(self.open_file_btn)
        out_row.addStretch()
        layout.addLayout(out_row)

        split = QHBoxLayout()

        log_group = QGroupBox("Ayrıntılı Günlük")
        log_layout = QVBoxLayout()
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        log_layout.addWidget(self.log)
        log_group.setLayout(log_layout)
        split.addWidget(log_group, stretch=2)

        hist_group = QGroupBox("Geçmiş Raporlar (çift tıkla aç)")
        hist_layout = QVBoxLayout()
        self.history_list = QListWidget()
        self.history_list.itemDoubleClicked.connect(self._open_history_item)
        hist_layout.addWidget(self.history_list)
        hist_group.setLayout(hist_layout)
        split.addWidget(hist_group, stretch=1)

        layout.addLayout(split, stretch=1)

        self._refresh_history_list()
        return page

    @staticmethod
    def _page_title(text):
        lbl = QLabel(text)
        lbl.setStyleSheet("font-size:15px; font-weight:700; color:#1a237e;")
        return lbl

    # ================================================================== kaynak yükleme
    def autodetect_source(self):
        try:
            src = engine.find_single_xlsx(report.SOURCE_DIR, "SOURCE")
            self._load_source(src)
        except engine.TransferError as e:
            self._log_line(str(e), "orange")

    def browse_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "Kaynak Excel Seç", "", "Excel dosyaları (*.xlsx)")
        if path:
            self._load_source(Path(path))

    def _load_source(self, path):
        self.source_path = Path(path)
        self.source_edit.setText(str(self.source_path))
        try:
            opts = report.scan_filter_options(report.CONFIG, self.source_path)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kaynak dosya okunamadı:\n{e}")
            return

        self._hierarchy = opts.get("hierarchy", {})
        if not (hasattr(self, "radio_source_dataset") and self.radio_source_dataset.isChecked()):
            self.tree.set_hierarchy(self._hierarchy)

        if opts["date_min"] and opts["date_max"]:
            self.date_start_edit.setDate(QDate(opts["date_min"].year, opts["date_min"].month, opts["date_min"].day))
            self.date_end_edit.setDate(QDate(opts["date_max"].year, opts["date_max"].month, opts["date_max"].day))

        self.source_info_label.setText(
            f"<b>{self.source_path.name}</b><br>"
            f"{len(opts['teams'])} grup · {len(opts['subteams'])} alt ekip · "
            f"{len(opts.get('roles', []))} meslek grubu<br>"
            f"Tarih aralığı: {opts['date_min']:%d.%m.%Y} - {opts['date_max']:%d.%m.%Y}"
            if opts["date_min"] else f"<b>{self.source_path.name}</b>"
        )
        if opts["warning"]:
            self._log_line(f"UYARI: {opts['warning']}", "orange")

        self.settings["last_source"] = str(self.source_path)
        self._update_status_bar()

    # ================================================================== durum çubuğu / özet
    def _scope_description(self):
        teams = self.tree.checked_teams()
        subteams = self.tree.checked_subteams()
        roles = self.tree.checked_roles()
        if roles:
            return " / ".join(roles) if len(roles) <= 3 else f"{len(roles)} meslek grubu"
        if subteams:
            return " / ".join(subteams) if len(subteams) <= 3 else f"{len(subteams)} alt ekip"
        if teams:
            return " / ".join(teams) if len(teams) <= 3 else f"{len(teams)} grup"
        return "Tüm Gruplar"

    def _update_status_bar(self):
        scope = self._scope_description()
        start = self.date_start_edit.date().toPython()
        end = self.date_end_edit.date().toPython()
        span = (end - start).days + 1 if end >= start else 0
        period_label = PERIOD_OPTIONS[self.period_combo.currentIndex()][0]
        auto_daily = span and span <= report.CONFIG.get("auto_daily_max_days", 10)
        period_txt = f"{period_label} → günlük kırılıma geçer" if auto_daily else period_label

        self.status_label.setText(
            f"📁 {self.source_path.name if self.source_path else 'kaynak yok'}   ·   "
            f"🔍 Kapsam: {scope}   ·   "
            f"📅 {start:%d.%m.%Y} - {end:%d.%m.%Y} ({span} gün)   ·   ⏱ {period_txt}"
        )

        n_teams = len(self.tree.checked_teams())
        n_sub = len(self.tree.checked_subteams())
        n_roles = len(self.tree.checked_roles())
        if n_teams or n_sub or n_roles:
            self.filter_summary.setText(
                f"<b>Seçili:</b> {n_teams} grup, {n_sub} alt ekip, {n_roles} meslek grubu &nbsp;"
                f"<span style='color:#78909c;'>— {scope}</span>"
            )
        else:
            self.filter_summary.setText("<span style='color:#78909c;'>Hiçbir filtre seçilmedi — tüm veriler dahil edilecek.</span>")

        self._update_threshold_effective_label()

    def _update_threshold_effective_label(self, _value=None):
        start = self.date_start_edit.date().toPython()
        end = self.date_end_edit.date().toPython()
        span = (end - start).days + 1 if end >= start else 0
        monthly = self.threshold_spin.value()
        effective = monthly * span / 30.0 if span else 0.0
        self.threshold_effective_label.setText(
            f"Seçili dönem {span} gün → bu rapor için etkin eşik: {effective:.1f} saat"
        )

    def _set_date_range(self, start, end):
        self.date_start_edit.setDate(QDate(start.year, start.month, start.day))
        self.date_end_edit.setDate(QDate(end.year, end.month, end.day))
        self._update_status_bar()

    # ================================================================== ayarlar
    def _restore_settings(self):
        s = self.settings
        self.threshold_spin.setValue(float(s.get("threshold", report.CONFIG.get("overtime_threshold", 52.0))))
        self.employee_list_checkbox.setChecked(s.get("include_employee_list", True))
        self.compare_checkbox.setChecked(s.get("compare_previous", True))
        period_idx = s.get("period_index", 0)
        if 0 <= period_idx < self.period_combo.count():
            self.period_combo.setCurrentIndex(period_idx)

    def _persist_settings(self):
        self.settings.update({
            "threshold": self.threshold_spin.value(),
            "include_employee_list": self.employee_list_checkbox.isChecked(),
            "compare_previous": self.compare_checkbox.isChecked(),
            "period_index": self.period_combo.currentIndex(),
        })
        save_json(TEAM_SETTINGS_PATH, self.settings)

    def save_state(self):
        try:
            self._persist_settings()
        except Exception:
            pass

    # ================================================================== log / rozet
    def _clear_log(self):
        self._log_buffer = []
        self._render_log()

    def _log_line(self, text, color=None, bold=False):
        self._log_buffer.append((text, color, bold))
        self._render_log()

    def _render_log(self):
        import html as html_lib
        parts = []
        for text, color, bold in self._log_buffer:
            escaped = html_lib.escape(text).replace("\n", "<br>")
            styles = []
            if color:
                styles.append(f"color:{TAG_COLORS.get(color, color)}")
            if bold:
                styles.append("font-weight:700")
            style_attr = f' style="{";".join(styles)}"' if styles else ""
            parts.append(f"<span{style_attr}>{escaped}</span>")
        self.log.setHtml(f"<div style='font-family:Consolas,monospace; font-size:10pt;'>{''.join(parts)}</div>")
        self.log.moveCursor(QTextCursor.MoveOperation.End)

    def _update_badges(self, items):
        while self.badges_layout.count():
            item = self.badges_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        for label_text, value, color in items:
            lbl = QLabel(f"{label_text}: {value}")
            lbl.setStyleSheet(
                f"background-color:{color}; color:white; border-radius:11px; "
                f"padding:5px 14px; font-weight:700; font-size:10pt;"
            )
            self.badges_layout.addWidget(lbl)
        self.badges_layout.addStretch()

    # ================================================================== parametre toplama
    def _validate_and_collect(self):
        if not self.source_path:
            QMessageBox.critical(self, "Eksik dosya", "Önce bir kaynak excel seçmelisin.")
            return None

        date_start = self.date_start_edit.date().toPython()
        date_end = self.date_end_edit.date().toPython()
        if date_start > date_end:
            QMessageBox.critical(self, "Geçersiz aralık", "Başlangıç tarihi bitiş tarihinden sonra olamaz.")
            return None

        cfg = dict(report.CONFIG)
        cfg["source_path"] = str(self.source_path)
        cfg["overtime_threshold"] = self.threshold_spin.value()

        using_dataset = self.radio_source_dataset.isChecked()
        tc_filter = None
        team_filter = subteam_filter = role_filter = None
        if using_dataset:
            checked_grup = self.tree.checked_teams() or None
            checked_ekip = self.tree.checked_subteams() or None
            checked_gorev = self.tree.checked_roles() or None
            tc_filter = db.get_tc_set(grup_filter=checked_grup, ekip_filter=checked_ekip,
                                       gorev_filter=checked_gorev)
            if not tc_filter:
                QMessageBox.critical(self, "Eşleşen personel yok",
                                      "Veri setinde bu filtrelerle eşleşen personel bulunamadı.")
                return None
        else:
            team_filter = self.tree.checked_teams() or None
            subteam_filter = self.tree.checked_subteams() or None
            role_filter = self.tree.checked_roles() or None

        # Kademe 1: Veri Seti ile etiket düzeltme
        dataset_lookup = None
        if using_dataset:
            dataset_lookup = db.get_lookup_dict(tc_list=list(tc_filter))
        elif self.dataset_correct_checkbox.isChecked():
            dataset_lookup = db.get_lookup_dict()

        return {
            "cfg": cfg,
            "source_path": self.source_path,
            "date_start": date_start,
            "date_end": date_end,
            "period_days": PERIOD_OPTIONS[self.period_combo.currentIndex()][1],
            "team_filter": team_filter,
            "subteam_filter": subteam_filter,
            "role_filter": role_filter,
            "tc_filter": tc_filter,
            "using_dataset": using_dataset,
            "dataset_lookup": dataset_lookup,
            "want_history": self.history_checkbox.isChecked(),
            "include_employee_list": self.employee_list_checkbox.isChecked(),
            "compare_previous": self.compare_checkbox.isChecked(),
            "threshold": self.threshold_spin.value(),
        }

    def _preview_badges_from_data(self, data):
        grand_normal = sum(sum(t.values()) for t in data["team_period_normal"].values())
        grand_fazla = sum(sum(t.values()) for t in data["team_period_fazla"].values())
        grand_total = grand_normal + grand_fazla
        grand_yevmiye = sum(data["team_yevmiye_days"].values())
        grand_employees = sum(len(s) for s in data["team_employees"].values())
        over = sum(1 for h in data["employee_fazla_totals"].values() if h > data.get("threshold", 20.0))
        badges = [
            ("Toplam Mesai", f"{report.format_tr(grand_total)} sa", "#2e7d32"),
            ("Normal", f"{report.format_tr(grand_normal)} sa", "#1565c0"),
            ("Fazla", f"{report.format_tr(grand_fazla)} sa", "#1565c0"),
            ("Yevmiye Günü", f"{grand_yevmiye}", "#00695c"),
            ("Personel", f"{grand_employees}", "#1565c0"),
            ("Eşik Aşan (Fazla)", f"{over} kişi", "#e65100" if over else "#607d8b"),
        ]
        if data.get("previous_total"):
            change = (grand_total - data["previous_total"]) / data["previous_total"] * 100
            arrow = "▲" if change > 0 else ("▼" if change < 0 else "=")
            badges.append(("Önceki Döneme Göre", f"{arrow} %{report.format_tr(abs(change))}",
                          "#b71c1c" if change > 0 else "#2e7d32"))
        return badges

    # ================================================================== ÖNİZLE
    def on_preview(self):
        params = self._validate_and_collect()
        if not params:
            return
        self.nav.setCurrentRow(4)
        self._clear_log()
        self.progress.setFormat("Önizleme hesaplanıyor...")
        self.progress.setValue(30)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            data = report.build_report_data(
                params["cfg"], params["source_path"], params["date_start"], params["date_end"],
                params["period_days"], params["team_filter"], params["subteam_filter"],
                params["role_filter"], params["include_employee_list"], params["compare_previous"],
                tc_filter=params["tc_filter"], dataset_lookup=params["dataset_lookup"],
            )
            data["threshold"] = params["threshold"]
            self.progress.setValue(100)
            self.progress.setFormat("Önizleme hazır (dosya oluşturulmadı)")
            self._update_badges(self._preview_badges_from_data(data))
            self._log_line("Bu bir önizlemedir — hiçbir dosya oluşturulmadı.\n", bold=True)
            self._log_line(f"Kapsam: {self._scope_description()}\n")
            self._log_line(f"Tarih: {params['date_start']:%d.%m.%Y} - {params['date_end']:%d.%m.%Y}\n")
            self._log_line(f"Dönem sayısı: {len(data['period_blocks'])}\n")
            if data.get("missing_subteam_count") or data.get("missing_role_count") or data.get("tc_issues"):
                self._log_line(
                    f"Veri kalitesi uyarısı: {data.get('missing_subteam_count', 0)} eksik alt ekip, "
                    f"{data.get('missing_role_count', 0)} eksik görev, "
                    f"{len(data.get('tc_issues', []))} TC sorunu\n", "orange")
            if data.get("dataset_mismatches"):
                self._log_line(f"Veri Seti ile {len(data['dataset_mismatches'])} alan farklı bulundu (düzeltilecek).\n", "blue")
            self._log_line("\n'RAPOR OLUŞTUR' ile bu ayarlarla excel dosyasını oluşturabilirsin.", "blue")
        except engine.TransferError as e:
            self.progress.setValue(0)
            self.progress.setFormat("Hazır")
            self._log_line(f"HATA: {e}\n", "red")
        except Exception as e:
            self.progress.setValue(0)
            self.progress.setFormat("Hazır")
            self._log_line(f"Beklenmeyen hata: {e}\n", "red")
        finally:
            QApplication.restoreOverrideCursor()

    # ================================================================== RAPOR OLUŞTUR
    def on_generate(self):
        params = self._validate_and_collect()
        if not params:
            return

        self.nav.setCurrentRow(4)
        self._persist_settings()
        self.generate_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        self._clear_log()
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            self.progress.setValue(15)
            self.progress.setFormat("Kaynak okunuyor ve hesaplanıyor...")
            QApplication.processEvents()

            data = report.build_report_data(
                params["cfg"], params["source_path"], params["date_start"], params["date_end"],
                params["period_days"], params["team_filter"], params["subteam_filter"],
                params["role_filter"], params["include_employee_list"], params["compare_previous"],
                tc_filter=params["tc_filter"], dataset_lookup=params["dataset_lookup"],
            )

            self.progress.setValue(60)
            self.progress.setFormat("Excel raporu yazılıyor...")
            QApplication.processEvents()

            report.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_path = report.OUTPUT_DIR / params["cfg"].get("output_name", "Ekip_Mesai_Raporu.xlsx")

            history_changes = None
            if params.get("want_history"):
                all_tcs = {tc for s in data["team_employees"].values() for tc in s}
                history_changes = db.get_history_for_tcs(list(all_tcs)) if all_tcs else []

            locked_warning = None
            try:
                report.write_report(params["cfg"], data, output_path, history_changes=history_changes)
            except PermissionError:
                # Muhtemelen dosya Excel'de açık - farklı bir isimle kaydetmeyi dene
                stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                fallback_path = output_path.with_name(f"{output_path.stem}_{stamp}{output_path.suffix}")
                report.write_report(params["cfg"], data, fallback_path, history_changes=history_changes)
                locked_warning = (f"'{output_path.name}' dosyası açık olduğu için üzerine yazılamadı "
                                  f"(muhtemelen Excel'de açık). Bunun yerine '{fallback_path.name}' "
                                  f"olarak kaydedildi.")
                output_path = fallback_path

            self.progress.setValue(100)
            self.progress.setFormat("Tamamlandı")
            self.output_path = output_path
            self.open_output_btn.setEnabled(True)
            self.open_file_btn.setEnabled(True)

            span = (params["date_end"] - params["date_start"]).days + 1
            auto_daily = span <= report.CONFIG.get("auto_daily_max_days", 10)

            self._log_line("Rapor başarıyla oluşturuldu.\n", "green", bold=True)
            if locked_warning:
                self._log_line(f"UYARI: {locked_warning}\n", "orange", bold=True)
            self._log_line(f"Tarih aralığı: {params['date_start']:%d.%m.%Y} - {params['date_end']:%d.%m.%Y}  ({span} gün)\n")
            period_label = PERIOD_OPTIONS[self.period_combo.currentIndex()][0]
            self._log_line(f"Dönem: {period_label}" + (" → günlük kırılıma geçildi\n" if auto_daily else "\n"),
                          "orange" if auto_daily else None)
            self._log_line(f"Personel Kaynağı: {'🗂️ Veri Seti' if params.get('using_dataset') else '📄 Kaynak Excel'}\n", bold=True)
            if params.get("using_dataset"):
                self._log_line(f"Kapsam: {self._scope_description()}\n")
            else:
                self._log_line(f"Grup: {', '.join(params['team_filter']) if params['team_filter'] else 'Tümü'}\n")
                self._log_line(f"Bağlı olduğu ekip: {', '.join(params['subteam_filter']) if params['subteam_filter'] else 'Tümü'}\n")
                self._log_line(f"Meslek grubu: {', '.join(params['role_filter']) if params['role_filter'] else 'Tümü'}\n")
            self._log_line(f"Fazla mesai eşiği: {params['threshold']:g} saat\n")
            self._log_line(f"Personel listesi: {'Eklendi' if params['include_employee_list'] else 'Eklenmedi'}\n")
            self._log_line(f"Önceki dönemle karşılaştırma: {'Açık' if params['compare_previous'] else 'Kapalı'}\n")

            if data.get("missing_subteam_count") or data.get("missing_role_count") or data.get("tc_issues"):
                self._log_line(
                    f"\nVeri kalitesi: {data.get('missing_subteam_count', 0)} eksik alt ekip, "
                    f"{data.get('missing_role_count', 0)} eksik görev, "
                    f"{len(data.get('tc_issues', []))} TC sorunu — 'Veri Kalitesi Kontrolü' sayfasına bak.\n",
                    "orange")
            if data.get("dataset_mismatches"):
                self._log_line(f"Veri Seti ile {len(data['dataset_mismatches'])} alan düzeltildi.\n", "blue")
            kismi = sum(1 for d in data.get("employee_durum", {}).values() if d != "Tam Dönem")
            if kismi:
                self._log_line(f"{kismi} kişi bu dönemde kısmi çalıştı (ay içi giriş/çıkış).\n", "orange")

            self._log_line(f"\nÇıktı: {output_path}\n", bold=True)

            data["threshold"] = params["threshold"]
            self._update_badges(self._preview_badges_from_data(data))

            add_to_history({
                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                "path": str(output_path),
                "scope": self._scope_description(),
                "date_range": f"{params['date_start']:%d.%m.%Y} - {params['date_end']:%d.%m.%Y}",
            })
            self._refresh_history_list()

        except engine.TransferError as e:
            self.progress.setValue(0)
            self.progress.setFormat("Hata")
            self._log_line(f"HATA: {e}\n", "red")
            QMessageBox.critical(self, "Hata", str(e))
        except Exception as e:
            self.progress.setValue(0)
            self.progress.setFormat("Hata")
            self._log_line(f"Beklenmeyen hata: {e}\n", "red")
            QMessageBox.critical(self, "Hata", str(e))
        finally:
            QApplication.restoreOverrideCursor()
            self.generate_btn.setEnabled(True)
            self.preview_btn.setEnabled(True)

    # ================================================================== geçmiş / dosya açma
    def _refresh_history_list(self):
        self.history_list.clear()
        for entry in load_json(TEAM_HISTORY_PATH, []):
            label = f"{entry.get('timestamp', '?')}  ·  {entry.get('scope', '?')}\n{entry.get('date_range', '')}"
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, entry.get("path"))
            self.history_list.addItem(item)

    def _open_history_item(self, item):
        path = item.data(Qt.ItemDataRole.UserRole)
        if path and Path(path).exists():
            self._open_path(path)
        else:
            QMessageBox.warning(self, "Bulunamadı", "Bu rapor dosyası artık mevcut değil.")

    def open_output_folder(self):
        self._open_path(report.OUTPUT_DIR)

    def open_output_file(self):
        if self.output_path:
            self._open_path(self.output_path)

    @staticmethod
    def _open_path(path):
        path = str(path)
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])


# ============================================================ MODÜL 3: EKİP LİSTESİ

class RosterModule(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.source_path = None
        self.output_path = None
        self._hierarchy = {}
        self._log_buffer = []

        self._build_ui()
        self.autodetect_source()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.nav = QListWidget()
        self.nav.setObjectName("sidebar")
        self.nav.setFixedWidth(190)
        self.nav.setFrameShape(QFrame.Shape.NoFrame)
        for icon, label in [("📁", "Kaynak"), ("🔍", "Filtreler"), ("📊", "Oluştur && Sonuç")]:
            self.nav.addItem(QListWidgetItem(f"{icon}   {label.replace('&&', '&')}"))
        self.nav.currentRowChanged.connect(lambda row: self.pages.setCurrentIndex(max(row, 0)))
        outer.addWidget(self.nav)

        right = QVBoxLayout()
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(0)

        self.pages = QStackedWidget()
        self.pages.addWidget(self._build_source_page())
        self.pages.addWidget(self._build_filter_page())
        self.pages.addWidget(self._build_result_page())
        right.addWidget(self.pages, stretch=1)

        self.status_bar_widget = QFrame()
        self.status_bar_widget.setObjectName("statusBar")
        sb = QHBoxLayout(self.status_bar_widget)
        sb.setContentsMargins(14, 6, 14, 6)
        self.status_label = QLabel("Kaynak yüklenmedi.")
        self.status_label.setObjectName("statusText")
        sb.addWidget(self.status_label)
        sb.addStretch()
        right.addWidget(self.status_bar_widget)

        right_widget = QWidget()
        right_widget.setLayout(right)
        outer.addWidget(right_widget, stretch=1)

        self.nav.setCurrentRow(0)

    # ------------------------------------------------------------ Sayfa 1: Kaynak
    def _build_source_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)
        layout.addWidget(page_title("Kaynak Excel"))

        self.drop_zone = DropZone("📄  Kaynak excel'i buraya sürükle-bırak", multiple=False)
        self.drop_zone.filesDropped.connect(lambda paths: self._load_source(Path(paths[0])) if paths else None)
        layout.addWidget(self.drop_zone)

        row = QHBoxLayout()
        self.source_edit = QLineEdit()
        self.source_edit.setReadOnly(True)
        row.addWidget(self.source_edit, stretch=1)
        browse_btn = QPushButton("Gözat...")
        browse_btn.clicked.connect(self.browse_source)
        row.addWidget(browse_btn)
        detect_btn = QPushButton("SOURCE Klasöründen Algıla")
        detect_btn.clicked.connect(self.autodetect_source)
        row.addWidget(detect_btn)
        layout.addLayout(row)

        info_group = QGroupBox("Kaynak Bilgisi")
        info_layout = QVBoxLayout()
        self.source_info_label = QLabel("Henüz bir kaynak dosya yüklenmedi.")
        self.source_info_label.setWordWrap(True)
        info_layout.addWidget(self.source_info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        layout.addStretch()
        return page

    # ------------------------------------------------------------ Sayfa 2: Filtreler
    def _build_filter_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(10)
        layout.addWidget(page_title("Filtreler"))

        hint = QLabel("Bir grubu işaretlemek altındaki tüm alt ekip ve meslekleri de işaretler. "
                      "Hiçbir şey işaretlenmezse tüm personel listelenir.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#78909c; font-size:9pt;")
        layout.addWidget(hint)

        search_row = QHBoxLayout()
        self.filter_search = QLineEdit()
        self.filter_search.setPlaceholderText("🔎  Grup / ekip / meslek ara...")
        self.filter_search.textChanged.connect(lambda t: self.tree.filter_text(t))
        search_row.addWidget(self.filter_search)
        all_btn = QPushButton("Tümünü Seç")
        all_btn.clicked.connect(lambda: self.tree.set_all_checked(True))
        none_btn = QPushButton("Temizle")
        none_btn.clicked.connect(lambda: self.tree.set_all_checked(False))
        search_row.addWidget(all_btn)
        search_row.addWidget(none_btn)
        layout.addLayout(search_row)

        self.tree = HierarchyTree()
        self.tree.changed.connect(self._update_status_bar)
        layout.addWidget(self.tree, stretch=1)

        self.filter_summary = QLabel()
        self.filter_summary.setWordWrap(True)
        self.filter_summary.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(self.filter_summary)

        return page

    # ------------------------------------------------------------ Sayfa 3: Oluştur & Sonuç
    def _build_result_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)
        layout.addWidget(page_title("Ekip Listesi Oluştur"))

        action_layout = QHBoxLayout()
        self.generate_btn = QPushButton("EKİP LİSTESİ OLUŞTUR")
        self.generate_btn.setObjectName("primaryButton")
        self.generate_btn.clicked.connect(self.on_generate)
        self.open_output_btn = QPushButton("Çıktı Klasörünü Aç")
        self.open_output_btn.setEnabled(False)
        self.open_output_btn.clicked.connect(self.open_output_folder)
        self.open_file_btn = QPushButton("Listeyi Excel'de Aç")
        self.open_file_btn.setEnabled(False)
        self.open_file_btn.clicked.connect(self.open_output_file)
        action_layout.addWidget(self.generate_btn, stretch=1)
        action_layout.addWidget(self.open_output_btn)
        action_layout.addWidget(self.open_file_btn)
        layout.addLayout(action_layout)

        self.badges_widget = QWidget()
        self.badges_layout = QHBoxLayout(self.badges_widget)
        self.badges_layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.badges_widget)

        result_group = QGroupBox("Sonuç")
        result_layout = QVBoxLayout()
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        result_layout.addWidget(self.log)
        result_group.setLayout(result_layout)
        layout.addWidget(result_group, stretch=1)

        return page

    # ------------------------------------------------------------ kaynak yükleme
    def autodetect_source(self):
        try:
            src = engine.find_single_xlsx(SOURCE_DIR, "SOURCE")
            self._load_source(src)
        except engine.TransferError as e:
            self._log_line(str(e), "orange")

    def browse_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "Kaynak Excel Seç", "", "Excel dosyaları (*.xlsx)")
        if path:
            self._load_source(Path(path))

    def _load_source(self, path):
        self.source_path = Path(path)
        self.source_edit.setText(str(self.source_path))
        try:
            opts = report.scan_filter_options(report.CONFIG, self.source_path)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kaynak dosya okunamadı:\n{e}")
            return

        self._hierarchy = opts.get("hierarchy", {})
        self.tree.set_hierarchy(self._hierarchy)

        self.source_info_label.setText(
            f"<b>{self.source_path.name}</b><br>"
            f"{len(opts['teams'])} grup · {len(opts['subteams'])} alt ekip · "
            f"{len(opts.get('roles', []))} meslek grubu"
        )
        if opts["warning"]:
            self._log_line(f"UYARI: {opts['warning']}", "orange")

        self._update_status_bar()

    # ------------------------------------------------------------ durum çubuğu
    def _scope_description(self):
        teams = self.tree.checked_teams()
        subteams = self.tree.checked_subteams()
        roles = self.tree.checked_roles()
        if roles:
            return " / ".join(roles) if len(roles) <= 3 else f"{len(roles)} meslek grubu"
        if subteams:
            return " / ".join(subteams) if len(subteams) <= 3 else f"{len(subteams)} alt ekip"
        if teams:
            return " / ".join(teams) if len(teams) <= 3 else f"{len(teams)} grup"
        return "Tüm Gruplar"

    def _update_status_bar(self):
        scope = self._scope_description()
        self.status_label.setText(
            f"📁 {self.source_path.name if self.source_path else 'kaynak yok'}   ·   🔍 Kapsam: {scope}"
        )
        n_teams = len(self.tree.checked_teams())
        n_sub = len(self.tree.checked_subteams())
        n_roles = len(self.tree.checked_roles())
        if n_teams or n_sub or n_roles:
            self.filter_summary.setText(
                f"<b>Seçili:</b> {n_teams} grup, {n_sub} alt ekip, {n_roles} meslek grubu &nbsp;"
                f"<span style='color:#78909c;'>— {scope}</span>"
            )
        else:
            self.filter_summary.setText("<span style='color:#78909c;'>Hiçbir filtre seçilmedi — tüm personel listelenecek.</span>")

    # ------------------------------------------------------------ log
    def _clear_log(self):
        self._log_buffer = []
        self._render_log()

    def _log_line(self, text, color=None, bold=False):
        self._log_buffer.append((text, color, bold))
        self._render_log()

    def _render_log(self):
        parts = []
        for text, color, bold in self._log_buffer:
            escaped = html_lib.escape(text).replace("\n", "<br>")
            styles = []
            if color:
                styles.append(f"color:{TAG_COLORS.get(color, color)}")
            if bold:
                styles.append("font-weight:700")
            style_attr = f' style="{";".join(styles)}"' if styles else ""
            parts.append(f"<span{style_attr}>{escaped}</span>")
        self.log.setHtml(f"<div style='font-family:Consolas,monospace; font-size:10pt;'>{''.join(parts)}</div>")
        self.log.moveCursor(QTextCursor.MoveOperation.End)

    def _update_badges(self, items):
        while self.badges_layout.count():
            item = self.badges_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        for label_text, value, color in items:
            lbl = QLabel(f"{label_text}: {value}")
            lbl.setStyleSheet(
                f"background-color:{color}; color:white; border-radius:11px; "
                f"padding:5px 14px; font-weight:700; font-size:10pt;"
            )
            self.badges_layout.addWidget(lbl)
        self.badges_layout.addStretch()

    # ------------------------------------------------------------ oluşturma
    def on_generate(self):
        if not self.source_path:
            QMessageBox.critical(self, "Eksik dosya", "Önce bir kaynak excel seçmelisin.")
            return

        cfg = dict(report.CONFIG)
        cfg["source_path"] = str(self.source_path)
        team_filter = self.tree.checked_teams() or None
        subteam_filter = self.tree.checked_subteams() or None
        role_filter = self.tree.checked_roles() or None

        self.generate_btn.setEnabled(False)
        self._clear_log()
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            output_path, roster_data, scope_str = report.generate_roster(
                cfg, source_path=self.source_path, team_filter=team_filter,
                subteam_filter=subteam_filter, role_filter=role_filter,
            )
            self.output_path = output_path
            self.open_output_btn.setEnabled(True)
            self.open_file_btn.setEnabled(True)

            sorted_records, role_counts = report.sort_roster(roster_data["records"])

            self._log_line("Ekip listesi başarıyla oluşturuldu.\n", "green", bold=True)
            self._log_line(f"Kapsam: {scope_str}\n")
            self._log_line(f"Toplam personel: {len(sorted_records)}\n\n", bold=True)
            self._log_line("Göreve göre dağılım:\n", "blue")
            seen = []
            for rec in sorted_records:
                if rec["role"] not in seen:
                    seen.append(rec["role"])
            for role in seen:
                self._log_line(f"  {role}: {role_counts[role]} kişi\n")
            self._log_line(f"\nÇıktı: {output_path}\n", bold=True)

            self._update_badges([
                ("Toplam Personel", len(sorted_records), "#2e7d32"),
                ("Görev Sayısı", len(seen), "#1565c0"),
            ])
        except engine.TransferError as e:
            self._log_line(f"HATA: {e}\n", "red")
            QMessageBox.critical(self, "Hata", str(e))
        except Exception as e:
            self._log_line(f"Beklenmeyen hata: {e}\n", "red")
            QMessageBox.critical(self, "Hata", str(e))
        finally:
            QApplication.restoreOverrideCursor()
            self.generate_btn.setEnabled(True)

    # ------------------------------------------------------------ diğer
    def save_state(self):
        pass

    def open_output_folder(self):
        open_path(OUTPUT_DIR)

    def open_output_file(self):
        if self.output_path:
            open_path(self.output_path)


# ============================================================ MODÜL 4: PERSONEL MALİYET RAPORU

class CostReportModule(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.source_path = None
        self.hakedis_path = None
        self._hierarchy = {}
        self._log_buffer = []
        self.output_path = None

        self._build_ui()
        self.autodetect_source()
        self.autodetect_hakedis()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.nav = QListWidget()
        self.nav.setObjectName("sidebar")
        self.nav.setFixedWidth(200)
        self.nav.setFrameShape(QFrame.Shape.NoFrame)
        for icon, label in [("📁", "Kaynak"), ("💰", "Hakediş Dosyası"),
                            ("🔍", "Filtreler"), ("📊", "Oluştur && Sonuç")]:
            self.nav.addItem(QListWidgetItem(f"{icon}   {label.replace('&&', '&')}"))
        self.nav.currentRowChanged.connect(lambda row: self.pages.setCurrentIndex(max(row, 0)))
        outer.addWidget(self.nav)

        right = QVBoxLayout()
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(0)

        self.pages = QStackedWidget()
        self.pages.addWidget(self._build_source_page())
        self.pages.addWidget(self._build_hakedis_page())
        self.pages.addWidget(self._build_filter_page())
        self.pages.addWidget(self._build_result_page())
        right.addWidget(self.pages, stretch=1)

        self.status_bar_widget = QFrame()
        self.status_bar_widget.setObjectName("statusBar")
        sb = QHBoxLayout(self.status_bar_widget)
        sb.setContentsMargins(14, 6, 14, 6)
        self.status_label = QLabel("Kaynak yüklenmedi.")
        self.status_label.setObjectName("statusText")
        sb.addWidget(self.status_label)
        sb.addStretch()
        right.addWidget(self.status_bar_widget)

        right_widget = QWidget()
        right_widget.setLayout(right)
        outer.addWidget(right_widget, stretch=1)

        self.nav.setCurrentRow(0)

    # ------------------------------------------------------------ Sayfa 1: Kaynak
    def _build_source_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)
        layout.addWidget(page_title("Kaynak Puantaj Excel (Personel Seçimi İçin)"))

        self.source_drop = DropZone("📄  Kaynak puantaj excel'ini buraya sürükle-bırak", multiple=False)
        self.source_drop.filesDropped.connect(lambda paths: self._load_source(Path(paths[0])) if paths else None)
        layout.addWidget(self.source_drop)

        row = QHBoxLayout()
        self.source_edit = QLineEdit()
        self.source_edit.setReadOnly(True)
        row.addWidget(self.source_edit, stretch=1)
        browse_btn = QPushButton("Gözat...")
        browse_btn.clicked.connect(self.browse_source)
        row.addWidget(browse_btn)
        detect_btn = QPushButton("SOURCE Klasöründen Algıla")
        detect_btn.clicked.connect(self.autodetect_source)
        row.addWidget(detect_btn)
        layout.addLayout(row)

        info_group = QGroupBox("Kaynak Bilgisi")
        info_layout = QVBoxLayout()
        self.source_info_label = QLabel("Henüz bir kaynak dosya yüklenmedi.")
        self.source_info_label.setWordWrap(True)
        info_layout.addWidget(self.source_info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        layout.addStretch()
        return page

    # ------------------------------------------------------------ Sayfa 2: Hakediş Dosyası
    def _build_hakedis_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)
        layout.addWidget(page_title("Hakediş (Maliyet) Excel Dosyası"))

        self.hakedis_drop = DropZone("💰  Hakediş excel'ini buraya sürükle-bırak", multiple=False)
        self.hakedis_drop.filesDropped.connect(lambda paths: self._load_hakedis(Path(paths[0])) if paths else None)
        layout.addWidget(self.hakedis_drop)

        row = QHBoxLayout()
        self.hakedis_edit = QLineEdit()
        self.hakedis_edit.setReadOnly(True)
        row.addWidget(self.hakedis_edit, stretch=1)
        browse_btn = QPushButton("Gözat...")
        browse_btn.clicked.connect(self.browse_hakedis)
        row.addWidget(browse_btn)
        detect_btn = QPushButton("HAKEDIS Klasöründen Algıla")
        detect_btn.clicked.connect(self.autodetect_hakedis)
        row.addWidget(detect_btn)
        layout.addLayout(row)

        sheet_group = QGroupBox("Kullanılacak Hakediş Sayfaları (bu ayki personel bordro sayfalarını işaretle)")
        sheet_layout = QVBoxLayout()
        hint = QLabel("Dosyada T.C. / Adı Soyadı başlığı bulunan sayfalar otomatik listelenir "
                      "(ör. bir ay için hem Türk hem yabancı personel sayfası olabilir).")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#78909c; font-size:9pt;")
        sheet_layout.addWidget(hint)
        self.hakedis_sheet_list = QListWidget()
        sheet_layout.addWidget(self.hakedis_sheet_list)
        btn_row = QHBoxLayout()
        all_btn = QPushButton("Tümünü Seç")
        all_btn.clicked.connect(lambda: self._set_all_sheet_checks(True))
        none_btn = QPushButton("Temizle")
        none_btn.clicked.connect(lambda: self._set_all_sheet_checks(False))
        btn_row.addWidget(all_btn)
        btn_row.addWidget(none_btn)
        sheet_layout.addLayout(btn_row)
        sheet_group.setLayout(sheet_layout)
        layout.addWidget(sheet_group, stretch=1)

        return page

    # ------------------------------------------------------------ Sayfa 3: Filtreler
    def _build_filter_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(10)
        layout.addWidget(page_title("Filtreler"))

        hint = QLabel("Bir grubu işaretlemek altındaki tüm alt ekip ve meslekleri de işaretler. "
                      "Hiçbir şey işaretlenmezse tüm personel dahil edilir.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#78909c; font-size:9pt;")
        layout.addWidget(hint)

        search_row = QHBoxLayout()
        self.filter_search = QLineEdit()
        self.filter_search.setPlaceholderText("🔎  Grup / ekip / meslek ara...")
        self.filter_search.textChanged.connect(lambda t: self.tree.filter_text(t))
        search_row.addWidget(self.filter_search)
        all_btn = QPushButton("Tümünü Seç")
        all_btn.clicked.connect(lambda: self.tree.set_all_checked(True))
        none_btn = QPushButton("Temizle")
        none_btn.clicked.connect(lambda: self.tree.set_all_checked(False))
        search_row.addWidget(all_btn)
        search_row.addWidget(none_btn)
        layout.addLayout(search_row)

        self.tree = HierarchyTree()
        self.tree.changed.connect(self._update_status_bar)
        layout.addWidget(self.tree, stretch=1)

        self.filter_summary = QLabel()
        self.filter_summary.setWordWrap(True)
        self.filter_summary.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(self.filter_summary)

        return page

    # ------------------------------------------------------------ Sayfa 4: Oluştur & Sonuç
    def _build_result_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)
        layout.addWidget(page_title("Maliyet Raporu Oluştur"))

        action_layout = QHBoxLayout()
        self.generate_btn = QPushButton("MALİYET RAPORU OLUŞTUR")
        self.generate_btn.setObjectName("primaryButton")
        self.generate_btn.clicked.connect(self.on_generate)
        self.open_output_btn = QPushButton("Çıktı Klasörünü Aç")
        self.open_output_btn.setEnabled(False)
        self.open_output_btn.clicked.connect(self.open_output_folder)
        self.open_file_btn = QPushButton("Raporu Excel'de Aç")
        self.open_file_btn.setEnabled(False)
        self.open_file_btn.clicked.connect(self.open_output_file)
        action_layout.addWidget(self.generate_btn, stretch=1)
        action_layout.addWidget(self.open_output_btn)
        action_layout.addWidget(self.open_file_btn)
        layout.addLayout(action_layout)

        self.badges_widget = QWidget()
        self.badges_layout = QHBoxLayout(self.badges_widget)
        self.badges_layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.badges_widget)

        result_group = QGroupBox("Sonuç")
        result_layout = QVBoxLayout()
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        result_layout.addWidget(self.log)
        result_group.setLayout(result_layout)
        layout.addWidget(result_group, stretch=1)

        return page

    # ------------------------------------------------------------ kaynak (puantaj) yükleme
    def autodetect_source(self):
        try:
            src = engine.find_single_xlsx(report.SOURCE_DIR, "SOURCE")
            self._load_source(src)
        except engine.TransferError as e:
            self._log_line(str(e), "orange")

    def browse_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "Kaynak Excel Seç", "", "Excel dosyaları (*.xlsx)")
        if path:
            self._load_source(Path(path))

    def _load_source(self, path):
        self.source_path = Path(path)
        self.source_edit.setText(str(self.source_path))
        try:
            opts = report.scan_filter_options(cost.CONFIG, self.source_path)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kaynak dosya okunamadı:\n{e}")
            return

        self._hierarchy = opts.get("hierarchy", {})
        self.tree.set_hierarchy(self._hierarchy)

        self.source_info_label.setText(
            f"<b>{self.source_path.name}</b><br>"
            f"{len(opts['teams'])} grup · {len(opts['subteams'])} alt ekip · "
            f"{len(opts.get('roles', []))} meslek grubu"
        )
        if opts["warning"]:
            self._log_line(f"UYARI: {opts['warning']}", "orange")
        self._update_status_bar()

    # ------------------------------------------------------------ hakediş dosyası yükleme
    def autodetect_hakedis(self):
        try:
            hk = engine.find_single_xlsx(cost.HAKEDIS_DIR, "HAKEDIS")
            self._load_hakedis(hk)
        except engine.TransferError as e:
            self._log_line(str(e), "orange")

    def browse_hakedis(self):
        path, _ = QFileDialog.getOpenFileName(self, "Hakediş Excel Seç", "", "Excel dosyaları (*.xlsx)")
        if path:
            self._load_hakedis(Path(path))

    def _load_hakedis(self, path):
        self.hakedis_path = Path(path)
        self.hakedis_edit.setText(str(self.hakedis_path))
        try:
            import openpyxl
            wb = openpyxl.load_workbook(self.hakedis_path, data_only=True)
            sheets = cost.find_payroll_sheets(wb)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Hakediş dosyası okunamadı:\n{e}")
            return

        self.hakedis_sheet_list.clear()
        for sn in sheets:
            item = QListWidgetItem(sn)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.hakedis_sheet_list.addItem(item)

        if not sheets:
            self._log_line("UYARI: Hakediş dosyasında T.C./Adı Soyadı başlıklı bir sayfa bulunamadı.", "orange")
        else:
            self._log_line(f"Hakediş dosyası yüklendi: {len(sheets)} personel sayfası bulundu. "
                           f"Kullanmak istediklerini işaretle.", bold=True)
        self._update_status_bar()

    def _set_all_sheet_checks(self, checked):
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for i in range(self.hakedis_sheet_list.count()):
            self.hakedis_sheet_list.item(i).setCheckState(state)

    def _checked_sheets(self):
        return [self.hakedis_sheet_list.item(i).text() for i in range(self.hakedis_sheet_list.count())
                if self.hakedis_sheet_list.item(i).checkState() == Qt.CheckState.Checked]

    # ------------------------------------------------------------ durum çubuğu
    def _scope_description(self):
        teams = self.tree.checked_teams()
        subteams = self.tree.checked_subteams()
        roles = self.tree.checked_roles()
        if roles:
            return " / ".join(roles) if len(roles) <= 3 else f"{len(roles)} meslek grubu"
        if subteams:
            return " / ".join(subteams) if len(subteams) <= 3 else f"{len(subteams)} alt ekip"
        if teams:
            return " / ".join(teams) if len(teams) <= 3 else f"{len(teams)} grup"
        return "Tüm Gruplar"

    def _update_status_bar(self):
        scope = self._scope_description()
        src_name = self.source_path.name if self.source_path else "kaynak yok"
        hk_name = self.hakedis_path.name if self.hakedis_path else "hakediş yok"
        n_sheets = len(self._checked_sheets())
        self.status_label.setText(
            f"📁 {src_name}   ·   💰 {hk_name} ({n_sheets} sayfa)   ·   🔍 Kapsam: {scope}"
        )
        n_teams = len(self.tree.checked_teams())
        n_sub = len(self.tree.checked_subteams())
        n_roles = len(self.tree.checked_roles())
        if n_teams or n_sub or n_roles:
            self.filter_summary.setText(
                f"<b>Seçili:</b> {n_teams} grup, {n_sub} alt ekip, {n_roles} meslek grubu &nbsp;"
                f"<span style='color:#78909c;'>— {scope}</span>"
            )
        else:
            self.filter_summary.setText("<span style='color:#78909c;'>Hiçbir filtre seçilmedi — tüm personel dahil edilecek.</span>")

    # ------------------------------------------------------------ log / rozet
    def _clear_log(self):
        self._log_buffer = []
        self._render_log()

    def _log_line(self, text, color=None, bold=False):
        self._log_buffer.append((text, color, bold))
        self._render_log()

    def _render_log(self):
        parts = []
        for text, color, bold in self._log_buffer:
            escaped = html_lib.escape(text).replace("\n", "<br>")
            styles = []
            if color:
                styles.append(f"color:{TAG_COLORS.get(color, color)}")
            if bold:
                styles.append("font-weight:700")
            style_attr = f' style="{";".join(styles)}"' if styles else ""
            parts.append(f"<span{style_attr}>{escaped}</span>")
        self.log.setHtml(f"<div style='font-family:Consolas,monospace; font-size:10pt;'>{''.join(parts)}</div>")
        self.log.moveCursor(QTextCursor.MoveOperation.End)

    def _update_badges(self, items):
        while self.badges_layout.count():
            item = self.badges_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        for label_text, value, color in items:
            lbl = QLabel(f"{label_text}: {value}")
            lbl.setStyleSheet(
                f"background-color:{color}; color:white; border-radius:11px; "
                f"padding:5px 14px; font-weight:700; font-size:10pt;"
            )
            self.badges_layout.addWidget(lbl)
        self.badges_layout.addStretch()

    # ------------------------------------------------------------ oluşturma
    def on_generate(self):
        if not self.source_path:
            QMessageBox.critical(self, "Eksik dosya", "Önce bir kaynak puantaj excel seçmelisin.")
            return
        if not self.hakedis_path:
            QMessageBox.critical(self, "Eksik dosya", "Önce bir hakediş excel seçmelisin.")
            return
        sheets = self._checked_sheets()
        if not sheets:
            QMessageBox.critical(self, "Eksik seçim", "Hakediş dosyasından en az bir sayfa işaretlemelisin.")
            return

        team_filter = self.tree.checked_teams() or None
        subteam_filter = self.tree.checked_subteams() or None
        role_filter = self.tree.checked_roles() or None

        self.generate_btn.setEnabled(False)
        self._clear_log()
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            output_path, data, scope_str = cost.generate_cost_report(
                cost.CONFIG, source_path=self.source_path, hakedis_path=self.hakedis_path,
                hakedis_sheets=sheets, team_filter=team_filter,
                subteam_filter=subteam_filter, role_filter=role_filter,
            )
            self.output_path = output_path
            self.open_output_btn.setEnabled(True)
            self.open_file_btn.setEnabled(True)

            merged = data["merged"]
            grand_toplam = cost._sum(merged, "toplam_odenecek")
            grand_personele = cost._sum(merged, "personele_odenecek_toplam")

            self._log_line("Personel maliyet raporu başarıyla oluşturuldu.\n", "green", bold=True)
            self._log_line(f"Kapsam: {scope_str}\n")
            self._log_line(f"Kullanılan hakediş sayfaları: {', '.join(sheets)}\n")
            self._log_line(f"Eşleşen personel: {len(merged)}\n", bold=True)
            if data["not_found"]:
                self._log_line(f"Hakedişte bulunamayan: {len(data['not_found'])} kişi "
                               f"(ayrı sayfada listelendi)\n", "orange")
            if data["quality_flags"]:
                self._log_line(f"\nVeri kalitesi uyarısı olan personel ({len(data['quality_flags'])} kişi):\n", "red")
                for tc, name, flags in data["quality_flags"][:15]:
                    self._log_line(f"  {name} (TC {tc}): {', '.join(flags)}\n")
                if len(data["quality_flags"]) > 15:
                    self._log_line(f"  ... ve {len(data['quality_flags']) - 15} kişi daha\n")
            self._log_line(f"\nPersonele Ödenen Toplam: {report.format_tr(grand_personele)} TL\n", bold=True)
            self._log_line(f"TOPLAM MALİYET: {report.format_tr(grand_toplam)} TL\n", "green", bold=True)
            self._log_line(f"\nÇıktı: {output_path}\n", bold=True)

            self._update_badges([
                ("Eşleşen Personel", len(merged), "#2e7d32"),
                ("Eşleşmeyen", len(data["not_found"]), "#b71c1c" if data["not_found"] else "#607d8b"),
                ("Uyarı", len(data["quality_flags"]), "#e65100" if data["quality_flags"] else "#607d8b"),
                ("Toplam Maliyet (TL)", report.format_tr(grand_toplam, 0), "#1565c0"),
            ])
        except cost.CostReportError as e:
            self._log_line(f"HATA: {e}\n", "red")
            QMessageBox.critical(self, "Hata", str(e))
        except Exception as e:
            self._log_line(f"Beklenmeyen hata: {e}\n", "red")
            QMessageBox.critical(self, "Hata", str(e))
        finally:
            QApplication.restoreOverrideCursor()
            self.generate_btn.setEnabled(True)

    # ------------------------------------------------------------ diğer
    def save_state(self):
        pass

    def open_output_folder(self):
        open_path(cost.OUTPUT_DIR)

    def open_output_file(self):
        if self.output_path:
            open_path(self.output_path)


# ============================================================ MODÜL 5: VERİ SETİ

class DataSetModule(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.source_path = None
        self._hierarchy = {}
        self._log_buffer = []
        self._editing_tc = None
        self._current_records = []

        self._build_ui()
        self._refresh_all()
        self.autodetect_source()

    # ------------------------------------------------------------------ UI iskeleti
    def _build_ui(self):
        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.nav = QListWidget()
        self.nav.setObjectName("sidebar")
        self.nav.setFixedWidth(200)
        self.nav.setFrameShape(QFrame.Shape.NoFrame)
        for icon, label in [("📊", "Genel Bakış"), ("🔎", "Kayıtlar"),
                            ("➕", "Ekle && Düzenle"), ("📥", "Excel'den İçe Aktar")]:
            self.nav.addItem(QListWidgetItem(f"{icon}   {label.replace('&&', '&')}"))
        self.nav.currentRowChanged.connect(self._on_nav_changed)
        outer.addWidget(self.nav)

        right = QVBoxLayout()
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(0)

        self.pages = QStackedWidget()
        self.pages.addWidget(self._build_dashboard_page())
        self.pages.addWidget(self._build_records_page())
        self.pages.addWidget(self._build_form_page())
        self.pages.addWidget(self._build_import_page())
        right.addWidget(self.pages, stretch=1)

        self.status_bar_widget = QFrame()
        self.status_bar_widget.setObjectName("statusBar")
        sb = QHBoxLayout(self.status_bar_widget)
        sb.setContentsMargins(14, 6, 14, 6)
        self.status_label = QLabel(f"Veritabanı: {db.DB_PATH.name}")
        self.status_label.setObjectName("statusText")
        sb.addWidget(self.status_label)
        sb.addStretch()
        right.addWidget(self.status_bar_widget)

        right_widget = QWidget()
        right_widget.setLayout(right)
        outer.addWidget(right_widget, stretch=1)

        self.nav.setCurrentRow(0)

    def _on_nav_changed(self, row):
        self.pages.setCurrentIndex(max(row, 0))
        if row == 0:
            self._refresh_dashboard()
        elif row == 1:
            self._refresh_filter_combos()
            self._refresh_table()

    # ------------------------------------------------------------ Sayfa 1: Genel Bakış
    def _build_dashboard_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)
        layout.addWidget(page_title("Veri Seti — Genel Bakış"))

        self.dash_badges_widget = QWidget()
        self.dash_badges_layout = QHBoxLayout(self.dash_badges_widget)
        self.dash_badges_layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.dash_badges_widget)

        tables_row = QHBoxLayout()

        group_box = QGroupBox("Grup Bazında Dağılım")
        group_layout = QVBoxLayout()
        self.dash_group_table = QTableWidget(0, 2)
        self.dash_group_table.setHorizontalHeaderLabels(["Grup", "Personel Sayısı"])
        self.dash_group_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.dash_group_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        group_layout.addWidget(self.dash_group_table)
        group_box.setLayout(group_layout)
        tables_row.addWidget(group_box)

        role_box = QGroupBox("Görev Bazında Dağılım (ilk 15)")
        role_layout = QVBoxLayout()
        self.dash_role_table = QTableWidget(0, 2)
        self.dash_role_table.setHorizontalHeaderLabels(["Görevi", "Personel Sayısı"])
        self.dash_role_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.dash_role_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        role_layout.addWidget(self.dash_role_table)
        role_box.setLayout(role_layout)
        tables_row.addWidget(role_box)

        layout.addLayout(tables_row, stretch=1)

        refresh_btn = QPushButton("Yenile")
        refresh_btn.clicked.connect(self._refresh_dashboard)
        layout.addWidget(refresh_btn, alignment=Qt.AlignmentFlag.AlignLeft)

        return page

    def _refresh_dashboard(self):
        stats = db.get_stats()
        while self.dash_badges_layout.count():
            item = self.dash_badges_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        for label, value, color in [
            ("Toplam Personel", stats["total"], "#1565c0"),
            ("Aktif", stats["aktif"], "#2e7d32"),
            ("Ayrılmış", stats["ayrilmis"], "#b71c1c" if stats["ayrilmis"] else "#607d8b"),
        ]:
            lbl = QLabel(f"{label}: {value}")
            lbl.setStyleSheet(f"background-color:{color}; color:white; border-radius:11px; "
                              f"padding:6px 16px; font-weight:700; font-size:11pt;")
            self.dash_badges_layout.addWidget(lbl)
        self.dash_badges_layout.addStretch()

        self.dash_group_table.setRowCount(len(stats["by_group"]))
        for i, row in enumerate(stats["by_group"]):
            self.dash_group_table.setItem(i, 0, QTableWidgetItem(row["grup"] or "-"))
            self.dash_group_table.setItem(i, 1, QTableWidgetItem(str(row["adet"])))

        top_roles = stats["by_role"][:15]
        self.dash_role_table.setRowCount(len(top_roles))
        for i, row in enumerate(top_roles):
            self.dash_role_table.setItem(i, 0, QTableWidgetItem(row["gorevi"] or "-"))
            self.dash_role_table.setItem(i, 1, QTableWidgetItem(str(row["adet"])))

    # ------------------------------------------------------------ Sayfa 2: Kayıtlar
    def _build_records_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(10)
        layout.addWidget(page_title("Kayıtları Görüntüle / Filtrele"))

        filter_row = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔎  İsim veya TC ile ara...")
        self.search_edit.textChanged.connect(self._refresh_table)
        filter_row.addWidget(self.search_edit, stretch=2)

        self.filter_grup = QComboBox()
        self.filter_grup.currentIndexChanged.connect(self._on_filter_grup_changed)
        filter_row.addWidget(self.filter_grup, stretch=1)

        self.filter_ekip = QComboBox()
        self.filter_ekip.currentIndexChanged.connect(self._refresh_table)
        filter_row.addWidget(self.filter_ekip, stretch=1)

        self.filter_gorev = QComboBox()
        self.filter_gorev.currentIndexChanged.connect(self._refresh_table)
        filter_row.addWidget(self.filter_gorev, stretch=1)

        self.filter_durum = QComboBox()
        self.filter_durum.addItems(["Tümü", "Aktif", "Ayrılmış"])
        self.filter_durum.currentIndexChanged.connect(self._refresh_table)
        filter_row.addWidget(self.filter_durum, stretch=1)

        layout.addLayout(filter_row)

        self.table = QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels(
            ["Ad Soyad", "TC Kimlik No", "Grup", "Bağlı Olduğu Ekip", "Görevi",
             "İşe Giriş", "İşten Çıkış", "Durum", "Kaynak", "Son Güncelleme"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.itemDoubleClicked.connect(lambda _item: self._start_edit_selected())
        layout.addWidget(self.table, stretch=1)

        btn_row = QHBoxLayout()
        self.result_count_label = QLabel("0 kayıt")
        btn_row.addWidget(self.result_count_label)
        btn_row.addStretch()
        edit_btn = QPushButton("Düzenle")
        edit_btn.clicked.connect(self._start_edit_selected)
        del_btn = QPushButton("Sil")
        del_btn.clicked.connect(self._delete_selected)
        export_btn = QPushButton("Dışa Aktar (Excel)")
        export_btn.clicked.connect(self._export_current_view)
        refresh_btn = QPushButton("Yenile")
        refresh_btn.clicked.connect(lambda: (self._refresh_filter_combos(), self._refresh_table()))
        for b in (edit_btn, del_btn, export_btn, refresh_btn):
            btn_row.addWidget(b)
        layout.addLayout(btn_row)

        return page

    def _refresh_filter_combos(self):
        for combo, key in [(self.filter_grup, "grup"), (self.filter_ekip, "bagli_oldugu_ekip"),
                            (self.filter_gorev, "gorevi")]:
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("Tüm " + ("Gruplar" if key == "grup" else "Ekipler" if key == "bagli_oldugu_ekip" else "Görevler"))
            combo.addItems(db.distinct_values(key))
            idx = combo.findText(current)
            combo.setCurrentIndex(idx if idx >= 0 else 0)
            combo.blockSignals(False)

    def _on_filter_grup_changed(self):
        self._refresh_table()

    def _refresh_table(self):
        search = self.search_edit.text().strip() or None
        grup = self.filter_grup.currentText()
        grup = None if grup.startswith("Tüm") else grup
        ekip = self.filter_ekip.currentText()
        ekip = None if ekip.startswith("Tüm") else ekip
        gorev = self.filter_gorev.currentText()
        gorev = None if gorev.startswith("Tüm") else gorev
        durum_txt = self.filter_durum.currentText()
        durum = {"Aktif": "aktif", "Ayrılmış": "ayrilmis"}.get(durum_txt)

        records = db.query_personnel(search=search, grup=grup, subteam=ekip, gorevi=gorev, durum=durum)
        self._current_records = records
        self.result_count_label.setText(f"{len(records)} kayıt")

        self.table.setRowCount(len(records))
        for i, r in enumerate(records):
            durum_label = "Ayrılmış" if r.get("isten_cikis_tarihi") else "Aktif"
            vals = [r["ad_soyad"], r["tc"], r.get("grup") or "", r.get("bagli_oldugu_ekip") or "",
                    r.get("gorevi") or "", r.get("ise_giris_tarihi") or "", r.get("isten_cikis_tarihi") or "",
                    durum_label, r.get("kaynak") or "", r.get("updated_at") or ""]
            for c, v in enumerate(vals):
                item = QTableWidgetItem(str(v))
                if c == 7 and durum_label == "Ayrılmış":
                    item.setForeground(Qt.GlobalColor.red)
                self.table.setItem(i, c, item)

    def _selected_tc(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self._current_records):
            return None
        return self._current_records[row]["tc"]

    def _start_edit_selected(self):
        tc = self._selected_tc()
        if not tc:
            QMessageBox.information(self, "Seçim yok", "Önce tablodan bir personel seç.")
            return
        rec = db.get_personnel(tc)
        if not rec:
            return
        self._load_record_into_form(rec)
        self.nav.setCurrentRow(2)

    def _delete_selected(self):
        tc = self._selected_tc()
        if not tc:
            QMessageBox.information(self, "Seçim yok", "Önce tablodan bir personel seç.")
            return
        rec = db.get_personnel(tc)
        name = rec["ad_soyad"] if rec else tc
        answer = QMessageBox.question(
            self, "Silme Onayı",
            f"'{name}' (TC: {tc}) veri setinden kalıcı olarak silinecek. Emin misin?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if answer == QMessageBox.StandardButton.Yes:
            db.delete_personnel(tc)
            self._refresh_all()

    def _export_current_view(self):
        if not self._current_records:
            QMessageBox.information(self, "Boş liste", "Dışa aktarılacak kayıt yok.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Excel Olarak Kaydet",
                                               str(db.SCRIPT_DIR / "Personel_Veri_Seti.xlsx"),
                                               "Excel dosyaları (*.xlsx)")
        if not path:
            return
        out = db.export_to_excel(self._current_records, path)
        QMessageBox.information(self, "Tamamlandı", f"{len(self._current_records)} kayıt dışa aktarıldı:\n{out}")

    # ------------------------------------------------------------ Sayfa 3: Ekle / Düzenle
    def _build_form_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        self.form_title = page_title("Yeni Personel Ekle")
        layout.addWidget(self.form_title)

        form_group = QGroupBox("Personel Bilgileri")
        form_layout = QGridLayout()

        form_layout.addWidget(QLabel("Ad Soyad *:"), 0, 0)
        self.f_ad_soyad = QLineEdit()
        form_layout.addWidget(self.f_ad_soyad, 0, 1, 1, 3)

        form_layout.addWidget(QLabel("TC Kimlik No *:"), 1, 0)
        self.f_tc = QLineEdit()
        self.f_tc.setPlaceholderText("11 haneli TC Kimlik No")
        form_layout.addWidget(self.f_tc, 1, 1)
        self.f_tc_status = QLabel("")
        form_layout.addWidget(self.f_tc_status, 1, 2, 1, 2)
        self.f_tc.textChanged.connect(self._validate_tc_live)

        form_layout.addWidget(QLabel("Grup:"), 2, 0)
        self.f_grup = QComboBox()
        self.f_grup.setEditable(True)
        form_layout.addWidget(self.f_grup, 2, 1)

        form_layout.addWidget(QLabel("Bağlı Olduğu Ekip:"), 2, 2)
        self.f_ekip = QComboBox()
        self.f_ekip.setEditable(True)
        form_layout.addWidget(self.f_ekip, 2, 3)

        form_layout.addWidget(QLabel("Görevi:"), 3, 0)
        self.f_gorev = QComboBox()
        self.f_gorev.setEditable(True)
        form_layout.addWidget(self.f_gorev, 3, 1, 1, 3)

        form_layout.addWidget(QLabel("İşe Giriş Tarihi:"), 4, 0)
        self.f_giris_bilinmiyor = QCheckBox("Tarih bilinmiyor")
        self.f_giris_date = QDateEdit(QDate.currentDate())
        self.f_giris_date.setCalendarPopup(True)
        self.f_giris_date.setDisplayFormat("dd.MM.yyyy")
        self.f_giris_bilinmiyor.toggled.connect(self.f_giris_date.setDisabled)
        giris_row = QHBoxLayout()
        giris_row.addWidget(self.f_giris_date)
        giris_row.addWidget(self.f_giris_bilinmiyor)
        form_layout.addLayout(giris_row, 4, 1, 1, 3)

        form_layout.addWidget(QLabel("İşten Çıkış Tarihi:"), 5, 0)
        self.f_hala_calisiyor = QCheckBox("Hâlâ çalışıyor (çıkış tarihi yok)")
        self.f_hala_calisiyor.setChecked(True)
        self.f_cikis_date = QDateEdit(QDate.currentDate())
        self.f_cikis_date.setCalendarPopup(True)
        self.f_cikis_date.setDisplayFormat("dd.MM.yyyy")
        self.f_cikis_date.setEnabled(False)
        self.f_hala_calisiyor.toggled.connect(lambda checked: self.f_cikis_date.setDisabled(checked))
        cikis_row = QHBoxLayout()
        cikis_row.addWidget(self.f_cikis_date)
        cikis_row.addWidget(self.f_hala_calisiyor)
        form_layout.addLayout(cikis_row, 5, 1, 1, 3)

        form_layout.addWidget(QLabel("Notlar:"), 6, 0)
        self.f_notlar = QTextEdit()
        self.f_notlar.setFixedHeight(70)
        form_layout.addWidget(self.f_notlar, 6, 1, 1, 3)

        form_group.setLayout(form_layout)
        layout.addWidget(form_group)

        btn_row = QHBoxLayout()
        self.save_btn = QPushButton("KAYDET")
        self.save_btn.setObjectName("primaryButton")
        self.save_btn.clicked.connect(self._save_form)
        clear_btn = QPushButton("Temizle / Yeni Kayıt")
        clear_btn.clicked.connect(self._clear_form)
        btn_row.addWidget(self.save_btn, stretch=1)
        btn_row.addWidget(clear_btn)
        layout.addLayout(btn_row)

        self.form_feedback = QLabel("")
        self.form_feedback.setWordWrap(True)
        layout.addWidget(self.form_feedback)

        layout.addStretch()
        return page

    def _validate_tc_live(self):
        tc = self.f_tc.text().strip()
        if not tc:
            self.f_tc_status.setText("")
            return
        ok, msg = db.validate_tc(tc)
        self.f_tc_status.setText("✓ Geçerli" if ok else f"⚠ {msg}")
        self.f_tc_status.setStyleSheet(f"color: {'#2e7d32' if ok else '#e65100'};")

    def _load_record_into_form(self, rec):
        self._editing_tc = rec["tc"]
        self.form_title.setText(f"Personel Düzenle — {rec['ad_soyad']}")
        self.f_ad_soyad.setText(rec["ad_soyad"] or "")
        self.f_tc.setText(rec["tc"] or "")
        self.f_tc.setEnabled(False)
        self.f_grup.setCurrentText(rec.get("grup") or "")
        self.f_ekip.setCurrentText(rec.get("bagli_oldugu_ekip") or "")
        self.f_gorev.setCurrentText(rec.get("gorevi") or "")

        giris = rec.get("ise_giris_tarihi")
        if giris:
            try:
                y, m, d = map(int, giris.split("-")[:3])
                self.f_giris_date.setDate(QDate(y, m, d))
                self.f_giris_bilinmiyor.setChecked(False)
            except Exception:
                self.f_giris_bilinmiyor.setChecked(True)
        else:
            self.f_giris_bilinmiyor.setChecked(True)

        cikis = rec.get("isten_cikis_tarihi")
        if cikis:
            try:
                y, m, d = map(int, cikis.split("-")[:3])
                self.f_cikis_date.setDate(QDate(y, m, d))
                self.f_hala_calisiyor.setChecked(False)
            except Exception:
                self.f_hala_calisiyor.setChecked(True)
        else:
            self.f_hala_calisiyor.setChecked(True)

        self.f_notlar.setPlainText(rec.get("notlar") or "")
        self.form_feedback.setText("")

    def _clear_form(self):
        self._editing_tc = None
        self.form_title.setText("Yeni Personel Ekle")
        self.f_ad_soyad.clear()
        self.f_tc.clear()
        self.f_tc.setEnabled(True)
        self.f_grup.setCurrentText("")
        self.f_ekip.setCurrentText("")
        self.f_gorev.setCurrentText("")
        self.f_giris_bilinmiyor.setChecked(True)
        self.f_hala_calisiyor.setChecked(True)
        self.f_notlar.clear()
        self.f_tc_status.setText("")
        self.form_feedback.setText("")

    def _save_form(self):
        ad_soyad = self.f_ad_soyad.text().strip()
        tc = self.f_tc.text().strip()

        if not ad_soyad:
            self.form_feedback.setText("⚠ Ad Soyad boş olamaz.")
            self.form_feedback.setStyleSheet("color:#b71c1c;")
            return

        ok, msg = db.validate_tc(tc)
        if not ok:
            self.form_feedback.setText(f"⚠ {msg}")
            self.form_feedback.setStyleSheet("color:#b71c1c;")
            return

        record = {
            "tc": tc,
            "ad_soyad": ad_soyad,
            "grup": self.f_grup.currentText().strip(),
            "bagli_oldugu_ekip": self.f_ekip.currentText().strip(),
            "gorevi": self.f_gorev.currentText().strip(),
            "ise_giris_tarihi": None if self.f_giris_bilinmiyor.isChecked()
                                 else self.f_giris_date.date().toString("yyyy-MM-dd"),
            "isten_cikis_tarihi": None if self.f_hala_calisiyor.isChecked()
                                   else self.f_cikis_date.date().toString("yyyy-MM-dd"),
            "notlar": self.f_notlar.toPlainText().strip(),
        }

        try:
            result = db.upsert_personnel(record, source="manuel", overwrite_blanks=True)
        except db.DatasetError as e:
            self.form_feedback.setText(f"⚠ {e}")
            self.form_feedback.setStyleSheet("color:#b71c1c;")
            return

        labels = {"inserted": "Yeni personel eklendi.", "updated": "Personel güncellendi.",
                   "unchanged": "Değişiklik yok (bilgiler zaten aynıydı)."}
        self.form_feedback.setText(f"✓ {labels.get(result, result)}")
        self.form_feedback.setStyleSheet("color:#2e7d32;")
        self._refresh_all()
        self._clear_form()
        self.nav.setCurrentRow(1)

    # ------------------------------------------------------------ Sayfa 4: Excel'den İçe Aktar
    def _build_import_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)
        layout.addWidget(page_title("Kaynak Excel'den Toplu İçe Aktarma"))

        self.import_drop = DropZone("📄  Kaynak puantaj excel'ini buraya sürükle-bırak", multiple=False)
        self.import_drop.filesDropped.connect(lambda paths: self._load_source(Path(paths[0])) if paths else None)
        layout.addWidget(self.import_drop)

        row = QHBoxLayout()
        self.import_source_edit = QLineEdit()
        self.import_source_edit.setReadOnly(True)
        row.addWidget(self.import_source_edit, stretch=1)
        browse_btn = QPushButton("Gözat...")
        browse_btn.clicked.connect(self.browse_source)
        row.addWidget(browse_btn)
        detect_btn = QPushButton("SOURCE Klasöründen Algıla")
        detect_btn.clicked.connect(self.autodetect_source)
        row.addWidget(detect_btn)
        layout.addLayout(row)

        filter_group = QGroupBox("Filtreler (hiçbiri işaretlenmezse tüm personel aktarılır)")
        filter_layout = QVBoxLayout()
        search_row = QHBoxLayout()
        self.import_search = QLineEdit()
        self.import_search.setPlaceholderText("🔎  Grup / ekip / meslek ara...")
        self.import_search.textChanged.connect(lambda t: self.import_tree.filter_text(t))
        search_row.addWidget(self.import_search)
        all_btn = QPushButton("Tümünü Seç")
        all_btn.clicked.connect(lambda: self.import_tree.set_all_checked(True))
        none_btn = QPushButton("Temizle")
        none_btn.clicked.connect(lambda: self.import_tree.set_all_checked(False))
        search_row.addWidget(all_btn)
        search_row.addWidget(none_btn)
        filter_layout.addLayout(search_row)
        self.import_tree = HierarchyTree()
        filter_layout.addWidget(self.import_tree, stretch=1)
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group, stretch=1)

        action_row = QHBoxLayout()
        self.import_btn = QPushButton("İÇE AKTAR")
        self.import_btn.setObjectName("primaryButton")
        self.import_btn.clicked.connect(self._run_import)
        action_row.addWidget(self.import_btn, stretch=1)
        layout.addLayout(action_row)

        self.import_log = QTextEdit()
        self.import_log.setReadOnly(True)
        self.import_log.setFixedHeight(160)
        layout.addWidget(self.import_log)

        return page

    def autodetect_source(self):
        try:
            src = engine.find_single_xlsx(report.SOURCE_DIR, "SOURCE")
            self._load_source(src)
        except engine.TransferError:
            pass

    def browse_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "Kaynak Excel Seç", "", "Excel dosyaları (*.xlsx)")
        if path:
            self._load_source(Path(path))

    def _load_source(self, path):
        self.source_path = Path(path)
        self.import_source_edit.setText(str(self.source_path))
        try:
            opts = report.scan_filter_options(cost.CONFIG, self.source_path)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kaynak dosya okunamadı:\n{e}")
            return
        self._hierarchy = opts.get("hierarchy", {})
        self.import_tree.set_hierarchy(self._hierarchy)

    def _run_import(self):
        if not self.source_path:
            QMessageBox.critical(self, "Eksik dosya", "Önce bir kaynak excel seçmelisin.")
            return

        team_filter = self.import_tree.checked_teams() or None
        subteam_filter = self.import_tree.checked_subteams() or None
        role_filter = self.import_tree.checked_roles() or None

        self.import_btn.setEnabled(False)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            cfg = dict(cost.CONFIG)
            cfg["source_giris_col"] = "O"
            cfg["source_cikis_col"] = "P"
            result = db.import_from_source(cfg, self.source_path, team_filter, subteam_filter, role_filter)

            lines = [
                f"Taranan personel: {result['total_scanned']}",
                f"Yeni eklenen: {result['inserted']}",
                f"Güncellenen: {result['updated']}",
                f"Değişmeyen (zaten kayıtlı): {result['unchanged']}",
            ]
            if result["errors"]:
                lines.append(f"Hata: {result['errors']}")
                for tc, name, err in result["error_list"][:10]:
                    lines.append(f"  {name} (TC {tc}): {err}")
            if result.get("warning"):
                lines.append(f"UYARI: {result['warning']}")
            self.import_log.setPlainText("\n".join(lines))
            self._refresh_all()
            QMessageBox.information(self, "Tamamlandı",
                                     f"{result['inserted']} yeni, {result['updated']} güncellendi, "
                                     f"{result['unchanged']} değişmedi.")
        except Exception as e:
            self.import_log.setPlainText(f"HATA: {e}")
            QMessageBox.critical(self, "Hata", str(e))
        finally:
            QApplication.restoreOverrideCursor()
            self.import_btn.setEnabled(True)

    # ------------------------------------------------------------ ortak
    def _refresh_all(self):
        db.init_db()
        self._refresh_dashboard()
        self._refresh_filter_combos()
        self._refresh_table()

    def save_state(self):
        pass


# ============================================================ DIŞ KABUK: BİRLEŞİK PENCERE

class UnifiedWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Puantaj & Raporlama Merkezi")
        self.resize(1280, 880)

        self.suite_settings = load_json(SUITE_SETTINGS_PATH, {})
        self.dark_mode = self.suite_settings.get("dark_mode", False)

        self._build_ui()
        self._apply_style()
        self._install_shortcuts()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # --- üst başlık + modül sekmeleri + karanlık mod ---
        header = QFrame()
        header.setObjectName("suiteHeader")
        h = QHBoxLayout(header)
        h.setContentsMargins(20, 14, 20, 14)
        h.setSpacing(14)

        title = QLabel("Puantaj & Raporlama Merkezi")
        title.setObjectName("titleLabel")
        h.addWidget(title)
        h.addSpacing(24)

        self.tab_group = QButtonGroup(self)
        self.tab_group.setExclusive(True)
        self.btn_puantaj = QPushButton("📋  Puantaj Aktarımı   (Ctrl+1)")
        self.btn_puantaj.setObjectName("moduleTab")
        self.btn_puantaj.setCheckable(True)
        self.btn_puantaj.setChecked(True)
        self.btn_report = QPushButton("📊  Fazla Mesai Raporu   (Ctrl+2)")
        self.btn_report.setObjectName("moduleTab")
        self.btn_report.setCheckable(True)
        self.btn_roster = QPushButton("🧾  Ekip Listesi   (Ctrl+3)")
        self.btn_roster.setObjectName("moduleTab")
        self.btn_roster.setCheckable(True)
        self.btn_cost = QPushButton("💰  Personel Maliyet Raporu   (Ctrl+4)")
        self.btn_cost.setObjectName("moduleTab")
        self.btn_cost.setCheckable(True)
        self.btn_dataset = QPushButton("🗂️  Veri Seti   (Ctrl+5)")
        self.btn_dataset.setObjectName("moduleTab")
        self.btn_dataset.setCheckable(True)
        self.tab_group.addButton(self.btn_puantaj, 0)
        self.tab_group.addButton(self.btn_report, 1)
        self.tab_group.addButton(self.btn_roster, 2)
        self.tab_group.addButton(self.btn_cost, 3)
        self.tab_group.addButton(self.btn_dataset, 4)
        self.btn_puantaj.clicked.connect(lambda: self.pages.setCurrentIndex(0))
        self.btn_report.clicked.connect(lambda: self.pages.setCurrentIndex(1))
        self.btn_roster.clicked.connect(lambda: self.pages.setCurrentIndex(2))
        self.btn_cost.clicked.connect(lambda: self.pages.setCurrentIndex(3))
        self.btn_dataset.clicked.connect(lambda: self.pages.setCurrentIndex(4))
        h.addWidget(self.btn_puantaj)
        h.addWidget(self.btn_report)
        h.addWidget(self.btn_roster)
        h.addWidget(self.btn_cost)
        h.addWidget(self.btn_dataset)
        h.addStretch()

        self.dark_toggle = QCheckBox("🌙 Karanlık Mod")
        self.dark_toggle.setChecked(self.dark_mode)
        self.dark_toggle.toggled.connect(self._on_dark_toggled)
        h.addWidget(self.dark_toggle)

        outer.addWidget(header)

        # --- dört modül, yığın halinde ---
        self.pages = QStackedWidget()
        self.puantaj_module = PuantajModule()
        self.team_module = TeamReportModule()
        self.roster_module = RosterModule()
        self.cost_module = CostReportModule()
        self.dataset_module = DataSetModule()
        self.pages.addWidget(self.puantaj_module)
        self.pages.addWidget(self.team_module)
        self.pages.addWidget(self.roster_module)
        self.pages.addWidget(self.cost_module)
        self.pages.addWidget(self.dataset_module)
        outer.addWidget(self.pages, stretch=1)

    def _apply_style(self):
        self.setStyleSheet(build_stylesheet(self.dark_mode))

    def _on_dark_toggled(self, checked):
        self.dark_mode = checked
        self._apply_style()

    def _install_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+G"), self, activated=self._shortcut_generate)
        QShortcut(QKeySequence("Ctrl+O"), self, activated=self._shortcut_open_source)
        QShortcut(QKeySequence("Ctrl+P"), self, activated=self._shortcut_preview)
        QShortcut(QKeySequence("Ctrl+1"), self, activated=lambda: self.btn_puantaj.click())
        QShortcut(QKeySequence("Ctrl+2"), self, activated=lambda: self.btn_report.click())
        QShortcut(QKeySequence("Ctrl+3"), self, activated=lambda: self.btn_roster.click())
        QShortcut(QKeySequence("Ctrl+4"), self, activated=lambda: self.btn_cost.click())
        QShortcut(QKeySequence("Ctrl+5"), self, activated=lambda: self.btn_dataset.click())

    def _current_module(self):
        return [self.puantaj_module, self.team_module, self.roster_module,
                self.cost_module, self.dataset_module][self.pages.currentIndex()]

    def _shortcut_generate(self):
        module = self._current_module()
        if module is self.puantaj_module:
            module.on_run()
        elif hasattr(module, "on_generate"):
            module.on_generate()

    def _shortcut_open_source(self):
        self._current_module().browse_source()

    def _shortcut_preview(self):
        module = self._current_module()
        if hasattr(module, "on_preview"):
            module.on_preview()

    def closeEvent(self, event):
        self.suite_settings["dark_mode"] = self.dark_mode
        save_json(SUITE_SETTINGS_PATH, self.suite_settings)
        self.puantaj_module.save_state()
        self.team_module.save_state()
        self.roster_module.save_state()
        self.cost_module.save_state()
        self.dataset_module.save_state()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = UnifiedWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
