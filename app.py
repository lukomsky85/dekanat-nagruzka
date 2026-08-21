"""
GUI-приложение на PyQt6.
"""

import os
import sys
import tempfile
from datetime import datetime

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QTabWidget, QLabel, QLineEdit, QPushButton, QComboBox,
    QCheckBox, QRadioButton, QGroupBox, QTextEdit,  # <-- ДОБАВЛЕНО QTextEdit
    QTableWidget, QTableWidgetItem, QHeaderView,
    QProgressBar, QFileDialog, QMessageBox, QFrame,
    QAbstractItemView, QApplication, QDialog, QDialogButtonBox,
    QListWidget, QListWidgetItem, QInputDialog,  # <-- ДОБАВЛЕНО для мульти-фильтров и пресетов
    QProgressDialog,  # <-- ДОБАВЛЕНО для прогресса скачивания обновления
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QIcon, QAction

from config import load_config, save_config, APP_DIR
from database import (
    get_connection, test_connection,
    fetch_kafedry, fetch_prepodavateli, fetch_gody,
    fetch_vidy_zanyatiy, fetch_vidy_kontrolya, fetch_nagruzka,
)
from excel_export import build_workbook, transliterate_to_latin
from history import (
    load_report_history, save_report_history,
    add_report_to_history, get_last_report_for_teacher,
)
from constants import (
    ALL_KAFEDRY, ALL_PREPOD, ALL_GODY,
    VERSION, APP_NAME, APP_YEAR,
    DEVELOPER_NAME, DEVELOPER_TELEGRAM, DEVELOPER_GITHUB, TARGET_SYSTEM,
    TABLE_HEADERS_ROW1, TABLE_HEADERS_ROW2,
    UPDATE_MANIFEST_URL,
)
import updater


# =========================================================
#  ФОНОВЫЕ ЗАДАЧИ
# =========================================================

class Worker(QThread):
    """Универсальный воркер для фоновых задач."""
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            result = self.func(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class DownloadWorker(QThread):
    """Воркер для скачивания файла обновления с прогрессом (в процентах)."""
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, url, dest_path):
        super().__init__()
        self.url = url
        self.dest_path = dest_path

    def run(self):
        try:
            def on_progress(downloaded, total):
                pct = int(downloaded * 100 / total) if total else 0
                self.progress.emit(pct)

            updater.download_update(self.url, self.dest_path, progress_callback=on_progress)
            self.finished.emit(self.dest_path)
        except Exception as e:
            self.error.emit(str(e))


# =========================================================
#  ДИАЛОГ "О ПРОГРАММЕ"
# =========================================================

class AboutDialog(QDialog):
    """Пользовательский диалог 'О программе' с поддержкой кликабельных ссылок."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"О программе {APP_NAME}")
        self.resize(480, 500)
        
        layout = QVBoxLayout(self)
        
        about_text = f"""
        <h2>{APP_NAME}</h2>
        <p><b>Версия:</b> {VERSION}</p>
        <p><b>Разработано для:</b> {TARGET_SYSTEM}</p>
        <hr>
        <p>Программа для формирования отчётов по учебной нагрузке преподавателей.</p>
        <p><b>Возможности:</b></p>
        <ul>
            <li>Подключение к базе данных «Деканат» (MS SQL Server)</li>
            <li>Фильтрация по кафедрам, преподавателям и учебным годам</li>
            <li>Гибкие режимы вывода (по семестрам / общий)</li>
            <li>Автоматическая нумерация строк</li>
            <li>Экспорт в Excel с фирменной шапкой</li>
            <li>Сохранение в локальную или сетевую папку</li>
            <li>Автоматическое имя файла (дата_ФИО_латиницей)</li>
            <li>История сгенерированных отчётов</li>
            <li>Автоматическое подставление заведующего кафедрой и декана</li>
        </ul>
        <hr>
        <p><b>Разработчик:</b> {DEVELOPER_NAME}</p>
        <p><b>Telegram:</b> <a href="https://t.me/{DEVELOPER_TELEGRAM}">@{DEVELOPER_TELEGRAM}</a></p>
        <p><b>GitHub:</b> <a href="{DEVELOPER_GITHUB}">{DEVELOPER_GITHUB}</a></p>
        <hr>
        <p><b>Технологии:</b> Python, PyQt6, pyodbc, openpyxl</p>
        <p><b>База данных:</b> Microsoft SQL Server</p>
        <hr>
        <p style="color: #666; font-size: 11px;">
        © {APP_YEAR}. Все права защищены.
        </p>
        """
        
        # QLabel поддерживает setOpenExternalLinks, в отличие от QMessageBox
        label = QLabel(about_text)
        label.setTextFormat(Qt.TextFormat.RichText)
        label.setOpenExternalLinks(True)  # <-- Делает ссылки кликабельными!
        label.setWordWrap(True)
        label.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.addWidget(label)
        
        # Кнопка "OK" для закрытия
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


# =========================================================
#  ГЛАВНОЕ ОКНО
# =========================================================

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} — Деканат v{VERSION}")

        self.config = load_config()
        self.resize(820, 680)

        # Установка иконки, если есть
        icon_path = APP_DIR / "icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self._conn_cache = None
        self._prepod_map = {}
        self._worker = None  # ссылка на активного воркера
        self._last_preview_columns = None  # кэш последней выборки для превью (чтобы смена
        self._last_preview_rows = None     # нумерации не требовала повторного запроса к БД

        self._build_ui()
        self._create_menu_bar()
        self._load_initial_state()

    # ---------------------------------------------------------
    #  МЕНЮ
    # ---------------------------------------------------------

    def _create_menu_bar(self):
        """Создаёт меню бар приложения."""
        menubar = self.menuBar()

        # Меню "Файл"
        file_menu = menubar.addMenu("Файл")
        exit_action = QAction("Выход", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Меню "Справка"
        help_menu = menubar.addMenu("Справка")

        update_action = QAction("🔄  Проверить обновления...", self)
        update_action.triggered.connect(lambda: self.check_for_updates(silent=False))
        help_menu.addAction(update_action)
        help_menu.addSeparator()

        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        help_menu.addSeparator()
        about_qt_action = QAction("О Qt", self)
        about_qt_action.triggered.connect(lambda: QMessageBox.aboutQt(self))
        help_menu.addAction(about_qt_action)

    def show_about(self):
        """Показывает окно 'О программе'."""
        dialog = AboutDialog(self)
        dialog.exec()

    # =========================================================
    #  ПРОВЕРКА ОБНОВЛЕНИЙ
    # =========================================================

    def check_for_updates(self, silent: bool = True):
        """Запускает фоновую проверку версии по манифесту.
        silent=True (автозапуск при старте) - никаких окон, если обновлений
        нет или сервер недоступен. silent=False (кнопка в меню) - показывает
        результат в любом случае."""
        manifest_url = self.config.get("update_manifest_url") or UPDATE_MANIFEST_URL
        self._update_check_worker = Worker(updater.check_for_update, manifest_url, VERSION)
        self._update_check_worker.finished.connect(lambda result: self._on_update_check_finished(result, silent))
        self._update_check_worker.error.connect(lambda msg: self._on_update_check_error(msg, silent))
        self._update_check_worker.start()

    def _on_update_check_finished(self, result, silent: bool):
        if result is None:
            if not silent:
                QMessageBox.information(self, "Обновления", f"У вас установлена последняя версия ({VERSION}).")
            return

        notes = result.get("notes") or "(без описания изменений)"
        msg = (
            f"Доступна новая версия {result['version']} (у вас установлена {VERSION}).\n\n"
            f"Что нового:\n{notes}\n\n"
            f"Скачать и установить сейчас?"
        )
        reply = QMessageBox.question(
            self, "Доступно обновление", msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._start_update_download(result.get("url", ""), result["version"])

    def _on_update_check_error(self, message: str, silent: bool):
        if not silent:
            QMessageBox.warning(self, "Проверка обновлений", f"Не удалось проверить обновления:\n{message}")
        # silent=True - молча игнорируем (например, на этом ПК нет доступа в интернет)

    def _start_update_download(self, url: str, version: str):
        if not updater.is_frozen():
            QMessageBox.information(
                self, "Обновление",
                "Самообновление доступно только для собранного .exe.\n"
                "Вы сейчас запустили программу из исходников python - "
                "просто обновите файлы вручную из репозитория.\n\n"
                f"Ссылка на новую версию:\n{url or '(не указана в манифесте)'}"
            )
            return
        if not url:
            QMessageBox.warning(self, "Обновление", "В манифесте обновлений не указана ссылка на файл.")
            return

        self._update_progress_dialog = QProgressDialog("Скачивание обновления...", "Отмена", 0, 100, self)
        self._update_progress_dialog.setWindowTitle("Обновление")
        self._update_progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self._update_progress_dialog.setMinimumDuration(0)
        self._update_progress_dialog.setValue(0)

        dest_path = os.path.join(tempfile.gettempdir(), f"UchebnayaNagruzka_update_{version}.exe")

        self._download_worker = DownloadWorker(url, dest_path)
        self._download_worker.progress.connect(self._update_progress_dialog.setValue)
        self._download_worker.finished.connect(self._on_download_finished)
        self._download_worker.error.connect(self._on_download_error)
        self._update_progress_dialog.canceled.connect(self._download_worker.terminate)
        self._download_worker.start()
        self._update_progress_dialog.show()

    def _on_download_finished(self, dest_path: str):
        if hasattr(self, "_update_progress_dialog"):
            self._update_progress_dialog.close()

        reply = QMessageBox.question(
            self, "Обновление готово",
            "Новая версия скачана. Для завершения установки приложение "
            "закроется и перезапустится автоматически.\n\nПродолжить?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            updater.apply_update_and_restart(dest_path)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка обновления", str(e))
            return

        if self._conn_cache:
            try:
                self._conn_cache.close()
            except Exception:
                pass

        # ВАЖНО: sys.exit() внутри слота Qt не гарантирует завершение процесса -
        # SystemExit может быть "проглочен" циклом обработки событий QApplication.exec()
        # и процесс продолжает висеть в памяти. Из-за этого .bat-скрипт обновления
        # (который ждёт через tasklist, что процесс реально закрылся) зависал в
        # бесконечном ожидании и никогда не подменял .exe.
        # os._exit() завершает процесс немедленно на уровне ОС, в обход Qt - это
        # безопасно здесь, т.к. соединение с БД уже закрыто вручную выше.
        os._exit(0)

    def _on_download_error(self, message: str):
        if hasattr(self, "_update_progress_dialog"):
            self._update_progress_dialog.close()
        QMessageBox.warning(self, "Ошибка скачивания", f"Не удалось скачать обновление:\n{message}")

    # ---------------------------------------------------------
    #  ПОСТРОЕНИЕ ИНТЕРФЕЙСА
    # ---------------------------------------------------------

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(8, 8, 8, 8)

        # Вкладки
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self._build_connection_tab()
        self._build_parameters_tab()
        self._build_save_tab()
        self._build_history_tab()
        self._build_template_tab()
        self._build_query_presets_tab()

        # Нижняя панель
        bottom_frame = QFrame()
        bottom_layout = QVBoxLayout(bottom_frame)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        self.run_btn = QPushButton("  Сформировать отчёт")
        self.run_btn.setEnabled(False)
        self.run_btn.setMinimumHeight(40)
        self.run_btn.clicked.connect(self.run_report)
        bottom_layout.addWidget(self.run_btn)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        bottom_layout.addWidget(self.progress)

        self.status_label = QLabel(
            'Настройте подключение и нажмите "Подключиться и загрузить списки".'
        )
        self.status_label.setStyleSheet("color: #444;")
        self.status_label.setWordWrap(True)
        bottom_layout.addWidget(self.status_label)

        main_layout.addWidget(bottom_frame)

    # ---------- Вкладка 1: Подключение ----------

    def _build_connection_tab(self):
        tab = QWidget()
        layout = QGridLayout(tab)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        layout.addWidget(QLabel("Сервер SQL:"), 0, 0)
        self.server_edit = QLineEdit(self.config.get("server", r"DEKANAT100\SQLEXPRESS"))
        layout.addWidget(self.server_edit, 0, 1)

        layout.addWidget(QLabel("База данных:"), 1, 0)
        self.db_edit = QLineEdit(self.config.get("database", "Деканат"))
        layout.addWidget(self.db_edit, 1, 1)

        layout.addWidget(QLabel("Аутентификация:"), 2, 0)
        auth_layout = QHBoxLayout()
        self.auth_windows = QRadioButton("Windows")
        self.auth_sql = QRadioButton("SQL Server")
        auth_layout.addWidget(self.auth_windows)
        auth_layout.addWidget(self.auth_sql)
        auth_layout.addStretch()
        auth_group = QWidget()
        auth_group.setLayout(auth_layout)
        layout.addWidget(auth_group, 2, 1)

        self.auth_windows.setChecked(self.config.get("auth_type", "windows") == "windows")
        self.auth_sql.setChecked(self.config.get("auth_type") == "sql")
        self.auth_windows.toggled.connect(self.on_auth_change)

        layout.addWidget(QLabel("Логин:"), 3, 0)
        self.login_edit = QLineEdit(self.config.get("username", ""))
        layout.addWidget(self.login_edit, 3, 1)

        layout.addWidget(QLabel("Пароль:"), 4, 0)
        self.password_edit = QLineEdit(self.config.get("password", ""))
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password_edit, 4, 1)

        self.save_password_cb = QCheckBox("Запомнить пароль")
        self.save_password_cb.setChecked(self.config.get("save_password", False))
        self.save_password_cb.toggled.connect(self.on_auth_change)
        layout.addWidget(self.save_password_cb, 5, 1)

        self.warn_label = QLabel("⚠ Пароль сохраняется в открытом виде")
        self.warn_label.setStyleSheet("color: #cc6600; font-size: 11px;")
        layout.addWidget(self.warn_label, 6, 1)

        btn_frame = QWidget()
        btn_layout = QHBoxLayout(btn_frame)
        btn_layout.setContentsMargins(0, 0, 0, 0)

        self.test_btn = QPushButton("  Проверить подключение")
        self.test_btn.clicked.connect(self.test_connection_ui)
        btn_layout.addWidget(self.test_btn)

        self.connect_btn = QPushButton("⚡  Подключиться и загрузить списки")
        self.connect_btn.clicked.connect(self.load_lists)
        btn_layout.addWidget(self.connect_btn)

        self.save_conn_btn = QPushButton("💾  Сохранить параметры")
        self.save_conn_btn.clicked.connect(self.save_connection_params)
        btn_layout.addWidget(self.save_conn_btn)

        btn_layout.addStretch()
        layout.addWidget(btn_frame, 7, 0, 1, 2)

        hint = QLabel(
            "💡  После успешного подключения станут доступны фильтры на вкладке «Параметры»."
        )
        hint.setStyleSheet("color: #666; font-size: 11px;")
        hint.setWordWrap(True)
        layout.addWidget(hint, 8, 0, 1, 2)

        layout.setColumnStretch(1, 1)
        self.tabs.addTab(tab, "  🔌  Подключение  ")
        self.on_auth_change()

    # ---------- Вкладка 2: Параметры ----------

    def _build_parameters_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # Фильтры
        filters_group = QGroupBox("Фильтры")
        filters_layout = QGridLayout(filters_group)
        filters_layout.setSpacing(8)

        filters_layout.addWidget(QLabel("Кафедра:"), 0, 0)
        self.kafedra_combo = QComboBox()
        self.kafedra_combo.addItem(ALL_KAFEDRY)
        self.kafedra_combo.currentTextChanged.connect(self.on_kafedra_change)
        filters_layout.addWidget(self.kafedra_combo, 0, 1)

        filters_layout.addWidget(QLabel("Преподаватель:"), 1, 0)
        self.prepod_combo = QComboBox()
        self.prepod_combo.addItem(ALL_PREPOD)
        self.prepod_combo.currentTextChanged.connect(self.on_prepod_change)
        filters_layout.addWidget(self.prepod_combo, 1, 1)

        self.last_report_label = QLabel("")
        self.last_report_label.setStyleSheet("color: #0066cc;")
        self.last_report_label.setWordWrap(True)
        filters_layout.addWidget(self.last_report_label, 2, 0, 1, 2)

        filters_layout.addWidget(QLabel("Учебный год:"), 3, 0)
        self.god_combo = QComboBox()
        self.god_combo.addItem(ALL_GODY)
        filters_layout.addWidget(self.god_combo, 3, 1)

        note = QLabel('  "Все преподаватели" — по отдельному листу на каждого.')
        note.setStyleSheet("color: #777; font-size: 11px;")
        note.setWordWrap(True)
        filters_layout.addWidget(note, 4, 0, 1, 2)

        filters_layout.setColumnStretch(1, 1)
        layout.addWidget(filters_group)

        # Режим вывода
        mode_group = QGroupBox("Режим вывода")
        mode_layout = QHBoxLayout(mode_group)

        mode_layout.addWidget(QLabel("Показывать:"))
        self._mode_map = {
            "По семестрам (осенний + весенний)": "semester",
            "Только осенний семестр":            "autumn",
            "Только весенний семестр":           "spring",
            "Общий (одной таблицей)":            "all",
        }
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(list(self._mode_map.keys()))
        saved_mode = self.config.get("mode", "По семестрам (осенний + весенний)")
        idx = list(self._mode_map.keys()).index(saved_mode) if saved_mode in self._mode_map else 0
        self.mode_combo.setCurrentIndex(idx)
        mode_layout.addWidget(self.mode_combo)

        mode_layout.addSpacing(20)
        mode_layout.addWidget(QLabel("Нумерация:"))
        self._numbering_map = {
            "Как в базе (оригинальный № стр.)": "original",
            "Последовательно (1, 2, 3...)":    "sequential",
        }
        self.numbering_combo = QComboBox()
        self.numbering_combo.addItems(list(self._numbering_map.keys()))
        saved_num = self.config.get("numbering", "original")
        num_text = {v: k for k, v in self._numbering_map.items()}.get(
            saved_num, "Как в базе (оригинальный № стр.)"
        )
        self.numbering_combo.setCurrentText(num_text)
        mode_layout.addWidget(self.numbering_combo)

        mode_layout.addStretch()
        layout.addWidget(mode_group)

        # Дополнительные фильтры выборки данных
        extra_group = QGroupBox("Дополнительные фильтры выборки")
        extra_layout = QGridLayout(extra_group)
        extra_layout.setSpacing(8)

        extra_layout.addWidget(QLabel("Вид занятий:"), 0, 0)
        self.vid_zanyatiy_list = QListWidget()
        self.vid_zanyatiy_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.vid_zanyatiy_list.setMaximumHeight(90)
        extra_layout.addWidget(self.vid_zanyatiy_list, 0, 1)

        extra_layout.addWidget(QLabel("Вид контроля:"), 1, 0)
        self.vid_kontrolya_list = QListWidget()
        self.vid_kontrolya_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.vid_kontrolya_list.setMaximumHeight(90)
        extra_layout.addWidget(self.vid_kontrolya_list, 1, 1)

        extra_layout.addWidget(QLabel("Финансирование:"), 2, 0)
        self.finance_combo = QComboBox()
        self._finance_map = {
            "Всё": None,
            "Только бюджет": "budget",
            "Только внебюджет": "vneb",
        }
        self.finance_combo.addItems(list(self._finance_map.keys()))
        extra_layout.addWidget(self.finance_combo, 2, 1)

        extra_hint = QLabel("  Ничего не выбрано в списке — фильтр по этому полю не применяется.")
        extra_hint.setStyleSheet("color: #777; font-size: 11px;")
        extra_hint.setWordWrap(True)
        extra_layout.addWidget(extra_hint, 3, 0, 1, 2)

        extra_layout.setColumnStretch(1, 1)
        layout.addWidget(extra_group)

        layout.addStretch()

        self.tabs.addTab(tab, "  📋  Параметры  ")

    # ---------- Вкладка 3: Сохранение ----------

    def _build_save_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        save_group = QGroupBox("Место сохранения")
        save_layout = QVBoxLayout(save_group)

        self.use_network_cb = QCheckBox("Сохранять в сетевой папке")
        self.use_network_cb.setChecked(self.config.get("use_network_folder", False))
        save_layout.addWidget(self.use_network_cb)

        net_layout = QHBoxLayout()
        net_layout.addWidget(QLabel("Сетевая папка:"))
        self.network_path_edit = QLineEdit(self.config.get("network_folder_path", r"\\server\reports"))
        net_layout.addWidget(self.network_path_edit)
        net_browse_btn = QPushButton("Обзор...")
        net_browse_btn.clicked.connect(self.choose_network_folder)
        net_layout.addWidget(net_browse_btn)
        save_layout.addLayout(net_layout)

        local_layout = QHBoxLayout()
        local_layout.addWidget(QLabel("Локальная папка:"))
        self.local_path_edit = QLineEdit(self.config.get("local_folder_path", str(APP_DIR)))
        local_layout.addWidget(self.local_path_edit)
        local_browse_btn = QPushButton("Обзор...")
        local_browse_btn.clicked.connect(self.choose_local_folder)
        local_layout.addWidget(local_browse_btn)
        save_layout.addLayout(local_layout)

        layout.addWidget(save_group)

        name_group = QGroupBox("Имя файла")
        name_layout = QVBoxLayout(name_group)
        self.output_edit = QLineEdit(self._default_output_name())
        name_layout.addWidget(self.output_edit)
        hint = QLabel("📝  Формат: ГГГГММДД_Фамилия_Имя_Отчество.xlsx  (автоматически)")
        hint.setStyleSheet("color: #777; font-size: 11px;")
        name_layout.addWidget(hint)
        layout.addWidget(name_group)

        layout.addStretch()
        self.tabs.addTab(tab, "  💾  Сохранение  ")

    # ---------- Вкладка 4: История ----------

    def _build_history_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)

        btn_frame = QWidget()
        btn_layout = QHBoxLayout(btn_frame)
        btn_layout.setContentsMargins(0, 0, 0, 0)

        refresh_btn = QPushButton("🔄  Обновить")
        refresh_btn.clicked.connect(self.refresh_history)
        btn_layout.addWidget(refresh_btn)

        clear_btn = QPushButton("🗑  Очистить историю")
        clear_btn.clicked.connect(self.clear_history)
        btn_layout.addWidget(clear_btn)

        open_btn = QPushButton("📂  Открыть файл")
        open_btn.clicked.connect(self.open_history_file)
        btn_layout.addWidget(open_btn)

        btn_layout.addStretch()
        layout.addWidget(btn_frame)

        self.history_table = QTableWidget()
        self.history_table.setColumnCount(5)
        self.history_table.setHorizontalHeaderLabels(
            ["Дата", "Преподаватель", "Год", "Время", "Путь к файлу"]
        )
        self.history_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.history_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.doubleClicked.connect(self.open_history_file)

        header = self.history_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

        layout.addWidget(self.history_table)
        self.tabs.addTab(tab, "  📜  История  ")

        self.refresh_history()
        
    # ---------- Вкладка: Шаблон отчёта ----------

    def _build_template_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        group = QGroupBox("Настройки печатной формы отчёта")
        group_layout = QVBoxLayout(group)

        group_layout.addWidget(QLabel("Строки шапки учреждения (каждая с новой строки):"))
        self.inst_lines_edit = QTextEdit()
        self.inst_lines_edit.setPlainText(self.config.get("template_institution_lines", ""))
        self.inst_lines_edit.setMaximumHeight(80)
        group_layout.addWidget(self.inst_lines_edit)

        group_layout.addWidget(QLabel("Должность утверждающего:"))
        self.approver_title_edit = QLineEdit(self.config.get("template_approver_title", ""))
        group_layout.addWidget(self.approver_title_edit)

        group_layout.addWidget(QLabel("ФИО утверждающего:"))
        self.approver_name_edit = QLineEdit(self.config.get("template_approver_name", ""))
        group_layout.addWidget(self.approver_name_edit)

        group_layout.addWidget(QLabel("ФИО Начальника УМО:"))
        self.umo_name_edit = QLineEdit(self.config.get("template_umo_name", ""))
        group_layout.addWidget(self.umo_name_edit)

        layout.addWidget(group)

        # ---------- Предпросмотр шапки отчёта ----------
        preview_group = QGroupBox("Предпросмотр шапки отчёта")
        preview_layout = QVBoxLayout(preview_group)

        self.template_preview_label = QLabel()
        self.template_preview_label.setTextFormat(Qt.TextFormat.RichText)
        self.template_preview_label.setWordWrap(True)
        self.template_preview_label.setStyleSheet(
            "background: white; border: 1px solid #ccc; padding: 14px;"
        )
        self.template_preview_label.setMinimumHeight(160)
        preview_layout.addWidget(self.template_preview_label)

        layout.addWidget(preview_group)

        # Живое обновление предпросмотра при вводе шаблона
        self.inst_lines_edit.textChanged.connect(self._update_template_preview)
        self.approver_title_edit.textChanged.connect(self._update_template_preview)
        self.approver_name_edit.textChanged.connect(self._update_template_preview)
        self.umo_name_edit.textChanged.connect(self._update_template_preview)

        # Живое обновление предпросмотра также при изменении фильтров на вкладке "Параметры"
        # - режим не требует нового запроса к БД, только перерисовку
        self.mode_combo.currentTextChanged.connect(self._update_template_preview)
        # - нумерация не требует нового запроса, но меняет номера строк в примере таблицы
        self.numbering_combo.currentTextChanged.connect(self._on_numbering_changed)
        # - кафедра/преподаватель/год/доп.фильтры влияют на сами данные -> дозапрашиваем пример строк
        self.kafedra_combo.currentTextChanged.connect(self._refresh_data_preview)
        self.prepod_combo.currentTextChanged.connect(self._refresh_data_preview)
        self.god_combo.currentTextChanged.connect(self._refresh_data_preview)
        self.finance_combo.currentTextChanged.connect(self._refresh_data_preview)
        self.vid_zanyatiy_list.itemSelectionChanged.connect(self._refresh_data_preview)
        self.vid_kontrolya_list.itemSelectionChanged.connect(self._refresh_data_preview)

        # Кнопки
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("💾 Сохранить шаблон")
        save_btn.clicked.connect(self.save_template)
        btn_layout.addWidget(save_btn)

        reset_btn = QPushButton("🔄 Сбросить по умолчанию")
        reset_btn.clicked.connect(self.reset_template)
        btn_layout.addWidget(reset_btn)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)
        layout.addStretch()

        self.tabs.addTab(tab, "  📝  Шаблон отчёта  ")
        self._data_preview_html = ""
        self._refresh_data_preview()

    def _update_template_preview(self):
        """Собирает HTML-превью шапки отчёта из текущих (ещё не сохранённых) полей формы
        и из текущих значений фильтров на вкладке «Параметры»."""
        inst_raw = self.inst_lines_edit.toPlainText().strip()
        inst_lines = [line.strip() for line in inst_raw.split("\n") if line.strip()] or ["(строки шапки не заданы)"]
        approver_title = self.approver_title_edit.text().strip() or "(должность не задана)"
        approver_name = self.approver_name_edit.text().strip() or "(ФИО не задано)"
        umo_name = self.umo_name_edit.text().strip() or "(ФИО не задано)"

        inst_html = "<br>".join(f"<b>{line}</b>" for line in inst_lines)
        approver_title_html = approver_title.replace("\n", "<br>")
        year = datetime.now().year

        # ---- Текущие значения фильтров с вкладки "Параметры" ----
        kafedra = self.kafedra_combo.currentText().strip() if hasattr(self, "kafedra_combo") else ALL_KAFEDRY
        prepod = self.prepod_combo.currentText().strip() if hasattr(self, "prepod_combo") else ALL_PREPOD
        god = self.god_combo.currentText().strip() if hasattr(self, "god_combo") else ALL_GODY
        mode_text = self.mode_combo.currentText() if hasattr(self, "mode_combo") else ""
        numbering_text = self.numbering_combo.currentText() if hasattr(self, "numbering_combo") else ""

        kafedra_display = kafedra if kafedra and kafedra != ALL_KAFEDRY else "(все кафедры — отдельный лист на каждую)"
        prepod_display = prepod if prepod and prepod != ALL_PREPOD else "(все преподаватели — отдельный лист на каждого)"
        god_display = god if god and god != ALL_GODY else "20__/20__"

        # ---- Доп. фильтры (вид занятий/контроля, финансирование) ----
        note_parts = []
        if hasattr(self, "vid_zanyatiy_list"):
            selected = [i.text() for i in self.vid_zanyatiy_list.selectedItems()]
            if selected:
                note_parts.append("вид занятий: " + ", ".join(selected))
        if hasattr(self, "vid_kontrolya_list"):
            selected = [i.text() for i in self.vid_kontrolya_list.selectedItems()]
            if selected:
                note_parts.append("вид контроля: " + ", ".join(selected))
        if hasattr(self, "finance_combo") and self.finance_combo.currentText() not in ("", "Всё"):
            note_parts.append(self.finance_combo.currentText().lower())
        filter_note_html = ""
        if note_parts:
            filter_note_html = (
                f'<p style="font-size:9px; color:#888; text-align:center;">'
                f'Применены фильтры: {"; ".join(note_parts)}</p>'
            )

        mode_line = ""
        if mode_text:
            mode_line = f'<p style="font-size:10px; color:#555;"><b>Режим вывода:</b> {mode_text} &nbsp;&nbsp; <b>Нумерация:</b> {numbering_text}</p>'

        html = f"""
        <table width="100%" style="font-family: sans-serif; font-size: 11px;">
        <tr>
          <td width="70%" style="text-align:center; vertical-align:top;">{inst_html}</td>
          <td width="30%" style="text-align:left; vertical-align:top; font-size:10px;">
            УТВЕРЖДАЮ:<br>
            {approver_title_html}<br>
            _____________ {approver_name}<br>
            «___»___________ {year} г.
          </td>
        </tr>
        </table>
        <p style="text-align:center; font-weight:bold; font-size:14px; margin-top:10px;">
          УЧЕБНАЯ НАГРУЗКА НА {god_display} УЧЕБНЫЙ ГОД
        </p>
        {filter_note_html}
        <p style="font-size:11px; margin:2px 0;"><b>Кафедра:</b> {kafedra_display}</p>
        <p style="font-size:11px; margin:2px 0;"><b>ФИО преподавателя:</b> {prepod_display}</p>
        {mode_line}
        <hr>
        {self._data_preview_html or '<p style="font-size:10px; color:#888;">... таблица нагрузки ...</p>'}
        <p style="font-size:11px; margin:2px 0;">
          Заведующий кафедрой _____________&nbsp;&nbsp;&nbsp;
          Декан факультета _____________&nbsp;&nbsp;&nbsp;
          Начальник УМО _____________ {umo_name}
        </p>
        """
        self.template_preview_label.setText(html)

    def _on_numbering_changed(self):
        """Смена режима нумерации не требует повторного запроса к БД -
        просто перерисовываем таблицу-превью из уже загруженных строк, если они есть."""
        if self._last_preview_rows is not None:
            numbering = self._numbering_map.get(self.numbering_combo.currentText(), "original")
            self._data_preview_html = self._build_data_preview_table(
                self._last_preview_columns, self._last_preview_rows, numbering
            )
        self._update_template_preview()

    def _refresh_data_preview(self):
        """Дозапрашивает пример строк реальной нагрузки под текущие фильтры
        (только если выбран конкретный преподаватель и есть подключение к БД),
        затем перерисовывает предпросмотр."""
        if not hasattr(self, "template_preview_label"):
            return  # вкладка "Шаблон отчёта" ещё не построена

        prepod = self.prepod_combo.currentText().strip() if hasattr(self, "prepod_combo") else ""
        kod = self._prepod_map.get(prepod) if prepod else None

        if not self._conn_cache:
            self._last_preview_columns = None
            self._last_preview_rows = None
            self._data_preview_html = self._build_header_only_table(
                "(подключитесь к базе на вкладке «Подключение», чтобы видеть пример строк)"
            )
            self._update_template_preview()
            return

        if not prepod or prepod == ALL_PREPOD or kod is None:
            self._last_preview_columns = None
            self._last_preview_rows = None
            self._data_preview_html = self._build_header_only_table(
                "Выберите конкретного преподавателя, чтобы увидеть пример строк его нагрузки."
            )
            self._update_template_preview()
            return

        god = self.god_combo.currentText().strip() if hasattr(self, "god_combo") else ALL_GODY
        vidy_zanyatiy = [i.text() for i in self.vid_zanyatiy_list.selectedItems()] or None
        vidy_kontrolya = [i.text() for i in self.vid_kontrolya_list.selectedItems()] or None
        finance = self._finance_map.get(self.finance_combo.currentText())

        try:
            columns, rows = fetch_nagruzka(self._conn_cache, kod, god, vidy_zanyatiy, vidy_kontrolya, finance)
        except Exception as e:
            self._last_preview_columns = None
            self._last_preview_rows = None
            self._data_preview_html = f'<p style="color:#c00; font-size:10px;">Не удалось загрузить пример данных: {e}</p>'
            self._update_template_preview()
            return

        self._last_preview_columns = columns
        self._last_preview_rows = rows
        numbering = self._numbering_map.get(self.numbering_combo.currentText(), "original")
        self._data_preview_html = self._build_data_preview_table(columns, rows, numbering)
        self._update_template_preview()

    def _table_header_html(self):
        """HTML-шапка таблицы нагрузки с теми же объединениями колонок, что и в Excel
        (Поток / Нагрузка / Финансирование)."""
        return f"""
        <tr style="background:#eef;">
          <td rowspan="2" style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW1[0]}</td>
          <td rowspan="2" style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW1[1]}</td>
          <td rowspan="2" style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW1[2]}</td>
          <td rowspan="2" style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW1[3]}</td>
          <td rowspan="2" style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW1[4]}</td>
          <td rowspan="2" style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW1[5]}</td>
          <td rowspan="2" style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW1[6]}</td>
          <td rowspan="2" style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW1[7]}</td>
          <td colspan="2" style="border:1px solid #bbb; padding:2px; text-align:center;">{TABLE_HEADERS_ROW1[8]}</td>
          <td colspan="3" style="border:1px solid #bbb; padding:2px; text-align:center;">{TABLE_HEADERS_ROW1[10]}</td>
          <td colspan="2" style="border:1px solid #bbb; padding:2px; text-align:center;">{TABLE_HEADERS_ROW1[13]}</td>
        </tr>
        <tr style="background:#eef;">
          <td style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW2[8]}</td>
          <td style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW2[9]}</td>
          <td style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW2[10]}</td>
          <td style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW2[11]}</td>
          <td style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW2[12]}</td>
          <td style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW2[13]}</td>
          <td style="border:1px solid #bbb; padding:2px;">{TABLE_HEADERS_ROW2[14]}</td>
        </tr>
        """

    def _build_header_only_table(self, note_text: str) -> str:
        return f"""
        <table style="border-collapse:collapse; font-size:9px; width:100%;">
          {self._table_header_html()}
        </table>
        <p style="font-size:10px; color:#888; margin-top:4px;">{note_text}</p>
        """

    def _build_data_preview_table(self, columns, rows, numbering="original") -> str:
        col_map = {name: i for i, name in enumerate(columns)}

        def val(row, key):
            return row[col_map[key]] if key in col_map else ""

        if not rows:
            return self._build_header_only_table(
                "Нагрузка по текущим фильтрам не найдена (0 строк)."
            )

        max_rows = 5
        sample = rows[:max_rows]
        rows_html = ""
        for row_counter, row in enumerate(sample, start=1):
            kurs = val(row, "Курс")
            sem = val(row, "Семестр")
            kurs_sem = "/".join(str(x) for x in (kurs, sem) if x is not None)
            line_number = row_counter if numbering == "sequential" else val(row, "НомерСтроки")
            cells = [
                line_number, val(row, "Группа"), val(row, "Блок"),
                val(row, "Дисциплина"), kurs_sem, val(row, "ВидЗанятий"),
                val(row, "ВидКонтроля"), val(row, "Студентов"), val(row, "НомерПотока"),
                val(row, "ИндикаторПотока"), val(row, "НагрузкаАуд"), val(row, "НагрузкаДр"),
                val(row, "Итого"), val(row, "Бюджет"), val(row, "Внебюджет"),
            ]
            tds = "".join(
                f'<td style="border:1px solid #ddd; padding:2px;">{"" if c is None else c}</td>'
                for c in cells
            )
            rows_html += f"<tr>{tds}</tr>"

        more_note = ""
        if len(rows) > max_rows:
            more_note = (
                f'<p style="font-size:10px; color:#888; margin-top:2px;">'
                f'... и ещё {len(rows) - max_rows} строк(и) — показаны первые {max_rows}.</p>'
            )

        return f"""
        <table style="border-collapse:collapse; font-size:9px; width:100%;">
          {self._table_header_html()}
          {rows_html}
        </table>
        {more_note}
        """

    def save_template(self):
        """Сохраняет настройки шаблона в config.json"""
        self.config["template_institution_lines"] = self.inst_lines_edit.toPlainText()
        self.config["template_approver_title"] = self.approver_title_edit.text()
        self.config["template_approver_name"] = self.approver_name_edit.text()
        self.config["template_umo_name"] = self.umo_name_edit.text()
        
        save_config(self.config)
        QMessageBox.information(self, "Успех", "Шаблон отчёта успешно сохранён!\nНовые значения будут использованы при следующем формировании отчёта.")

    def reset_template(self):
        """Возвращает значения по умолчанию из constants.py"""
        from constants import INSTITUTION_LINES, APPROVER_TITLE, APPROVER_NAME, UMO_NAME
        
        self.inst_lines_edit.setPlainText("\n".join(INSTITUTION_LINES))
        self.approver_title_edit.setText(APPROVER_TITLE)
        self.approver_name_edit.setText(APPROVER_NAME)
        self.umo_name_edit.setText(UMO_NAME)      

    # ---------- Вкладка: Шаблоны выборки (пресеты фильтров) ----------

    def _build_query_presets_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        info = QLabel(
            "Сохраняйте комбинации фильтров (кафедра, преподаватель, год, режим, "
            "виды занятий/контроля, финансирование) под именем и вызывайте одним кликом."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: #666;")
        layout.addWidget(info)

        group = QGroupBox("Сохранённые шаблоны выборки")
        group_layout = QVBoxLayout(group)

        self.preset_combo = QComboBox()
        self._refresh_preset_combo()
        group_layout.addWidget(self.preset_combo)

        btn_layout = QHBoxLayout()

        apply_btn = QPushButton("📥  Применить")
        apply_btn.clicked.connect(self.apply_preset)
        btn_layout.addWidget(apply_btn)

        save_btn = QPushButton("💾  Сохранить текущие фильтры как...")
        save_btn.clicked.connect(self.save_preset_as)
        btn_layout.addWidget(save_btn)

        delete_btn = QPushButton("🗑  Удалить")
        delete_btn.clicked.connect(self.delete_preset)
        btn_layout.addWidget(delete_btn)

        btn_layout.addStretch()
        group_layout.addLayout(btn_layout)
        layout.addWidget(group)
        layout.addStretch()

        self.tabs.addTab(tab, "  💾  Шаблоны выборки  ")

    def _refresh_preset_combo(self):
        self.preset_combo.clear()
        templates = self.config.get("query_templates", {})
        if templates:
            self.preset_combo.addItems(sorted(templates.keys()))
        else:
            self.preset_combo.addItem("(нет сохранённых шаблонов)")

    def _collect_current_filters(self) -> dict:
        """Снимок текущего состояния всех фильтров выборки данных."""
        return {
            "kafedra": self.kafedra_combo.currentText(),
            "prepod": self.prepod_combo.currentText(),
            "god": self.god_combo.currentText(),
            "mode": self.mode_combo.currentText(),
            "numbering": self.numbering_combo.currentText(),
            "vid_zanyatiy": [i.text() for i in self.vid_zanyatiy_list.selectedItems()],
            "vid_kontrolya": [i.text() for i in self.vid_kontrolya_list.selectedItems()],
            "finance": self.finance_combo.currentText(),
        }

    def _apply_filters_dict(self, data: dict):
        """Восстанавливает состояние виджетов фильтров из сохранённого пресета."""
        def set_combo(combo, value):
            idx = combo.findText(value)
            if idx >= 0:
                combo.setCurrentIndex(idx)

        set_combo(self.kafedra_combo, data.get("kafedra", ALL_KAFEDRY))
        # Список преподавателей зависит от кафедры - обновится по сигналу currentTextChanged,
        # затем пробуем выставить нужного преподавателя (если он есть в новом списке)
        set_combo(self.prepod_combo, data.get("prepod", ALL_PREPOD))
        set_combo(self.god_combo, data.get("god", ALL_GODY))
        set_combo(self.mode_combo, data.get("mode", self.mode_combo.currentText()))
        set_combo(self.numbering_combo, data.get("numbering", self.numbering_combo.currentText()))
        set_combo(self.finance_combo, data.get("finance", "Всё"))

        def select_items(list_widget, values):
            list_widget.clearSelection()
            values = set(values or [])
            for i in range(list_widget.count()):
                item = list_widget.item(i)
                item.setSelected(item.text() in values)

        select_items(self.vid_zanyatiy_list, data.get("vid_zanyatiy"))
        select_items(self.vid_kontrolya_list, data.get("vid_kontrolya"))

    def apply_preset(self):
        name = self.preset_combo.currentText()
        templates = self.config.get("query_templates", {})
        if name not in templates:
            QMessageBox.information(self, "Информация", "Выберите сохранённый шаблон из списка.")
            return
        self._apply_filters_dict(templates[name])
        self.set_status(f"✅  Применён шаблон выборки «{name}».")

    def save_preset_as(self):
        if not self._conn_cache:
            QMessageBox.warning(self, "Внимание",
                                 "Сначала подключитесь и загрузите списки на вкладке «Подключение».")
            return
        name, ok = QInputDialog.getText(self, "Сохранить шаблон", "Название шаблона:")
        if not ok or not name.strip():
            return
        name = name.strip()
        templates = self.config.get("query_templates", {})
        templates[name] = self._collect_current_filters()
        self.config["query_templates"] = templates
        save_config(self.config)
        self._refresh_preset_combo()
        idx = self.preset_combo.findText(name)
        if idx >= 0:
            self.preset_combo.setCurrentIndex(idx)
        self.set_status(f"✅  Шаблон выборки «{name}» сохранён.")

    def delete_preset(self):
        name = self.preset_combo.currentText()
        templates = self.config.get("query_templates", {})
        if name not in templates:
            return
        confirm = QMessageBox.question(
            self, "Подтверждение", f"Удалить шаблон «{name}»?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm == QMessageBox.StandardButton.Yes:
            del templates[name]
            self.config["query_templates"] = templates
            save_config(self.config)
            self._refresh_preset_combo()
            self.set_status(f"🗑  Шаблон «{name}» удалён.")

    # ---------- Начальное состояние ----------

    def _load_initial_state(self):
        self.on_auth_change()
        if self.config.get("auto_check_updates", True):
            # Тихая проверка через 2с после старта - не мешает загрузке окна
            # и не показывает ничего, если обновлений нет или сервер недоступен
            QTimer.singleShot(2000, lambda: self.check_for_updates(silent=True))

    # =========================================================
    #  СЛОТЫ
    # =========================================================

    def on_auth_change(self):
        is_sql = self.auth_sql.isChecked()
        self.login_edit.setEnabled(is_sql)
        self.password_edit.setEnabled(is_sql)
        self.warn_label.setVisible(is_sql and self.save_password_cb.isChecked())

    def save_connection_params(self):
        self.config["server"] = self.server_edit.text().strip()
        self.config["database"] = self.db_edit.text().strip()
        self.config["auth_type"] = "sql" if self.auth_sql.isChecked() else "windows"
        self.config["username"] = self.login_edit.text().strip()
        self.config["save_password"] = self.save_password_cb.isChecked()

        if self.save_password_cb.isChecked():
            self.config["password"] = self.password_edit.text()
        else:
            self.config["password"] = ""

        save_config(self.config)
        self.set_status("✅  Параметры подключения сохранены.")
        QMessageBox.information(
            self, "Сохранено",
            "Параметры подключения сохранены в config.json\n\n"
            "Теперь при следующем запуске они будут загружены автоматически."
        )

    def test_connection_ui(self):
        self.test_btn.setEnabled(False)
        self.set_status("  Проверка подключения...")

        server = self.server_edit.text().strip()
        database = self.db_edit.text().strip()
        auth_type = "sql" if self.auth_sql.isChecked() else "windows"
        username = self.login_edit.text().strip() if auth_type == "sql" else ""
        password = self.password_edit.text() if auth_type == "sql" else ""

        self._worker = Worker(test_connection, server, database, auth_type, username, password)
        self._worker.finished.connect(self._on_test_finished)
        self._worker.error.connect(self._on_test_error)
        self._worker.start()

    def _on_test_finished(self, result):
        self.test_btn.setEnabled(True)
        success, message = result
        if success:
            QMessageBox.information(self, "✅ Успех", message)
            self.set_status("✅  Подключение успешно!")
        else:
            QMessageBox.critical(self, "❌ Ошибка", message)
            self.set_status("❌  Ошибка подключения.")

    def _on_test_error(self, error):
        self.test_btn.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", error)
        self.set_status("❌  Ошибка.")

    def load_lists(self):
        self.config["server"] = self.server_edit.text().strip()
        self.config["database"] = self.db_edit.text().strip()
        self.config["auth_type"] = "sql" if self.auth_sql.isChecked() else "windows"
        self.config["username"] = self.login_edit.text().strip()
        self.config["save_password"] = self.save_password_cb.isChecked()
        if self.save_password_cb.isChecked():
            self.config["password"] = self.password_edit.text()
        else:
            self.config["password"] = ""

        self.config["use_network_folder"] = self.use_network_cb.isChecked()
        self.config["network_folder_path"] = self.network_path_edit.text().strip()
        self.config["local_folder_path"] = self.local_path_edit.text().strip()
        save_config(self.config)

        self.connect_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.set_status("  Подключаюсь к серверу...")

        server = self.server_edit.text().strip()
        database = self.db_edit.text().strip()
        auth_type = self.config["auth_type"]
        username = self.config["username"]
        password = self.config["password"]

        self._worker = Worker(self._do_load_lists, server, database, auth_type, username, password)
        self._worker.progress.connect(self.set_status)
        self._worker.finished.connect(self._on_load_finished)
        self._worker.error.connect(self._on_load_error)
        self._worker.start()

    def _do_load_lists(self, server, database, auth_type, username, password):
        if auth_type == "sql" and not password and self.config.get("save_password"):
            password = self.config.get("password", "")

        conn = get_connection(server, database, auth_type, username, password)

        self._worker.progress.emit("📚  Загружаю список кафедр...")
        kafedry = fetch_kafedry(conn)

        self._worker.progress.emit("🏫  Загружаю список преподавателей...")
        prepodavateli = fetch_prepodavateli(conn, None)

        self._worker.progress.emit("📅  Загружаю список учебных годов...")
        gody = fetch_gody(conn)

        self._worker.progress.emit("🗂  Загружаю виды занятий и контроля...")
        vidy_zanyatiy = fetch_vidy_zanyatiy(conn)
        vidy_kontrolya = fetch_vidy_kontrolya(conn)

        return conn, kafedry, prepodavateli, gody, vidy_zanyatiy, vidy_kontrolya

    def _on_load_finished(self, result):
        conn, kafedry, prepodavateli, gody, vidy_zanyatiy, vidy_kontrolya = result
        self._conn_cache = conn

        self.kafedra_combo.clear()
        self.kafedra_combo.addItem(ALL_KAFEDRY)
        self.kafedra_combo.addItems(kafedry)

        self._prepod_map = {fio: kod for kod, fio in prepodavateli}
        self.prepod_combo.clear()
        self.prepod_combo.addItem(ALL_PREPOD)
        self.prepod_combo.addItems(list(self._prepod_map.keys()))

        self.god_combo.clear()
        self.god_combo.addItem(ALL_GODY)
        self.god_combo.addItems(gody)

        self.vid_zanyatiy_list.clear()
        self.vid_zanyatiy_list.addItems(vidy_zanyatiy)

        self.vid_kontrolya_list.clear()
        self.vid_kontrolya_list.addItems(vidy_kontrolya)

        self.run_btn.setEnabled(True)
        self.connect_btn.setEnabled(True)
        self.progress.setVisible(False)
        self.set_status(
            f"✅  Загружено: кафедр — {len(kafedry)}, преподавателей — {len(prepodavateli)}, годов — {len(gody)}."
        )
        self._refresh_data_preview()

    def _on_load_error(self, error):
        self.progress.setVisible(False)
        self.connect_btn.setEnabled(True)
        self.set_status("❌  Ошибка подключения.")
        QMessageBox.warning(self, "Ошибка подключения", f"Не удалось подключиться:\n{error}")

    def on_kafedra_change(self, kafedra):
        if not self._conn_cache:
            return
        if kafedra == ALL_KAFEDRY:
            self.prepod_combo.clear()
            self.prepod_combo.addItem(ALL_PREPOD)
            self.prepod_combo.addItems(list(self._prepod_map.keys()))
            return

        self.set_status("  Обновляю список преподавателей...")
        self._worker = Worker(fetch_prepodavateli, self._conn_cache, kafedra)
        self._worker.finished.connect(self._on_kafedra_filtered)
        self._worker.error.connect(lambda e: self.set_status(f"  Ошибка: {e}"))
        self._worker.start()

    def _on_kafedra_filtered(self, prepodavateli):
        self._prepod_map = {fio: kod for kod, fio in prepodavateli}
        self.prepod_combo.clear()
        self.prepod_combo.addItem(ALL_PREPOD)
        self.prepod_combo.addItems(list(self._prepod_map.keys()))
        self.set_status(f"✅  Преподавателей по кафедре: {len(prepodavateli)}.")

    def on_prepod_change(self, prepod):
        god = self.god_combo.currentText()
        if prepod and prepod != ALL_PREPOD:
            last_report = get_last_report_for_teacher(prepod, god if god != ALL_GODY else None)
            if last_report:
                gen_time = last_report.get("generation_time", "")
                try:
                    dt = datetime.fromisoformat(gen_time)
                    formatted_time = dt.strftime("%d.%m.%Y %H:%M")
                except Exception:
                    formatted_time = gen_time
                self.last_report_label.setText(f"  Последний отчёт: {formatted_time}")
            else:
                self.last_report_label.setText("  Отчёты ещё не генерировались")
        else:
            self.last_report_label.setText("")

    def choose_network_folder(self):
        path = QFileDialog.getExistingDirectory(
            self, "Выберите сетевую папку", self.network_path_edit.text()
        )
        if path:
            self.network_path_edit.setText(path)

    def choose_local_folder(self):
        path = QFileDialog.getExistingDirectory(
            self, "Выберите локальную папку", self.local_path_edit.text()
        )
        if path:
            self.local_path_edit.setText(path)

    def _default_output_name(self):
        return f"{datetime.now():%Y%m%d}_otchet.xlsx"

    def set_status(self, text):
        self.status_label.setText(text)
        QApplication.processEvents()

    # ---------- История ----------

    def refresh_history(self):
        self.history_table.setRowCount(0)
        history = load_report_history()
        history.sort(key=lambda x: x.get("generation_time", ""), reverse=True)

        self.history_table.setRowCount(len(history))
        for row_idx, record in enumerate(history):
            gen_time = record.get("generation_time", "")
            try:
                dt = datetime.fromisoformat(gen_time)
                date_str = dt.strftime("%d.%m.%Y")
                time_str = dt.strftime("%H:%M:%S")
            except Exception:
                date_str = record.get("date", "")
                time_str = ""

            self.history_table.setItem(row_idx, 0, QTableWidgetItem(date_str))
            self.history_table.setItem(row_idx, 1, QTableWidgetItem(record.get("teacher_fio", "")))
            self.history_table.setItem(row_idx, 2, QTableWidgetItem(record.get("year", "")))
            self.history_table.setItem(row_idx, 3, QTableWidgetItem(time_str))
            self.history_table.setItem(row_idx, 4, QTableWidgetItem(record.get("file_path", "")))

    def clear_history(self):
        reply = QMessageBox.question(
            self, "Подтверждение", "Удалить всю историю отчётов?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            save_report_history([])
            self.refresh_history()
            self.set_status("История очищена.")

    def open_history_file(self):
        selected = self.history_table.selectedItems()
        if not selected:
            QMessageBox.information(self, "Информация", "Выберите запись в таблице.")
            return

        row = selected[0].row()
        file_path = self.history_table.item(row, 4).text()

        if file_path and os.path.exists(file_path):
            try:
                os.startfile(file_path)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл:\n{e}")
        else:
            QMessageBox.warning(self, "Файл не найден", f"Файл не существует:\n{file_path}")

    # ---------- Формирование отчёта ----------

    def run_report(self):
        self.run_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.set_status("📊  Формирую отчёт...")

        kafedra = self.kafedra_combo.currentText().strip()
        prepod = self.prepod_combo.currentText().strip()
        god = self.god_combo.currentText().strip()
        mode = self._mode_map.get(self.mode_combo.currentText(), "semester")
        numbering = self._numbering_map.get(self.numbering_combo.currentText(), "original")
        vidy_zanyatiy = [i.text() for i in self.vid_zanyatiy_list.selectedItems()] or None
        vidy_kontrolya = [i.text() for i in self.vid_kontrolya_list.selectedItems()] or None
        finance = self._finance_map.get(self.finance_combo.currentText())

        self.config["mode"] = self.mode_combo.currentText()
        self.config["numbering"] = numbering
        save_config(self.config)

        save_folder = (
            self.network_path_edit.text() if self.use_network_cb.isChecked()
            else self.local_path_edit.text()
        ).strip()

        self._worker = Worker(
            self._do_build_report,
            kafedra, prepod, god, mode, numbering, save_folder,
            vidy_zanyatiy, vidy_kontrolya, finance
        )
        self._worker.progress.connect(self.set_status)
        self._worker.finished.connect(self._on_report_finished)
        self._worker.error.connect(self._on_report_error)
        self._worker.start()

    def _do_build_report(self, kafedra, prepod, god, mode, numbering, save_folder,
                          vidy_zanyatiy=None, vidy_kontrolya=None, finance=None):
        conn = self._conn_cache
        if conn is None:
            server = self.server_edit.text().strip()
            database = self.db_edit.text().strip()
            auth_type = self.config["auth_type"]
            username = self.config["username"]
            password = self.config["password"]
            conn = get_connection(server, database, auth_type, username, password)
            self._conn_cache = conn

        if prepod and prepod != ALL_PREPOD:
            kod = self._prepod_map.get(prepod)
            if kod is None:
                raise Exception("Не удалось определить код выбранного преподавателя.")
            teachers = [(kod, prepod)]
            single_teacher = True
            teacher_fio_for_name = prepod
        else:
            teachers = fetch_prepodavateli(conn, kafedra)
            single_teacher = False
            teacher_fio_for_name = "all_teachers"

        if not teachers:
            raise Exception("По заданным фильтрам преподавателей не найдено.")

        self._worker.progress.emit(f"📊  Собираю данные по {len(teachers)} преподавател(ю/ям)...")
        wb = build_workbook(conn, teachers, kafedra, god, mode, numbering,
                             vidy_zanyatiy, vidy_kontrolya, finance)

        date_str = datetime.now().strftime("%Y%m%d")
        # base-имя всегда определено (и для одного преподавателя, и для пакетного режима) -
        # раньше в пакетном режиме teacher_latin оставался None и ломал имя файла при коллизии
        teacher_latin = transliterate_to_latin(teacher_fio_for_name) if single_teacher else "all_teachers"
        filename = f"{date_str}_{teacher_latin}.xlsx"

        if not os.path.exists(save_folder):
            os.makedirs(save_folder, exist_ok=True)

        full_path = os.path.join(save_folder, filename)

        if os.path.exists(full_path):
            time_str = datetime.now().strftime("%H%M%S")
            filename = f"{date_str}_{teacher_latin}_{time_str}.xlsx"
            full_path = os.path.join(save_folder, filename)

        wb.save(full_path)

        generation_time = datetime.now().isoformat()
        network_path = save_folder if self.use_network_cb.isChecked() else ""

        if single_teacher:
            add_report_to_history(
                teacher_fio=teacher_fio_for_name,
                year=god if god != ALL_GODY else "",
                file_path=full_path,
                generation_time=generation_time,
                network_path=network_path
            )

        return full_path

    def _on_report_finished(self, full_path):
        self.progress.setVisible(False)
        self.run_btn.setEnabled(True)
        self.set_status(f"✅  Готово! Файл сохранён: {full_path}")
        self.refresh_history()

        reply = QMessageBox.question(
            self, "✅ Готово", f"Отчёт сохранён:\n{full_path}\n\nОткрыть файл?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            try:
                os.startfile(full_path)
            except Exception:
                pass

    def _on_report_error(self, error):
        self.progress.setVisible(False)
        self.run_btn.setEnabled(True)
        self.set_status("❌  Произошла ошибка.")
        QMessageBox.warning(self, "Ошибка", f"Ошибка:\n{error}")