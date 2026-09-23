# -*- coding: utf-8 -*-
"""
Модуль интерактивного диалога помощи AgiAnalytics
Современный интерфейс в премиальном минималистичном стиле без использования эмодзи.
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QScrollArea, QWidget,
                             QFrame, QGridLayout, QListWidget,
                             QListWidgetItem, QStackedWidget, QLineEdit,
                             QProgressBar)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class HelpDialog(QDialog):
    """Диалог интерактивной помощи в современном стиле"""

    def __init__(self, role, db, parent=None):
        super().__init__(parent)
        self.role = role
        self.db = db
        self.current_step = 0
        self.setModal(True)
        self.setWindowTitle('Справка и руководство пользователя — AgiAnalytics')
        self.setMinimumSize(960, 680)
        self.resize(1020, 720)

        # Главная палитра стилей и сброс конфликтующих стилей
        self.setStyleSheet("""
            QDialog {
                background-color: #f8fafc;
                font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                border: none;
                background-color: #f1f5f9;
                width: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background-color: #cbd5e1;
                border-radius: 4px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #94a3b8;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ================= 1. ВЕРХНЯЯ ПАНЕЛЬ (HEADER) =================
        header_widget = QWidget()
        header_widget.setFixedHeight(72)
        header_widget.setStyleSheet("""
            QWidget {
                background-color: #0f172a;
                border-bottom: 1px solid #1e293b;
            }
        """)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(24, 0, 24, 0)
        header_layout.setSpacing(16)

        # Блок заголовка
        titles_text_layout = QVBoxLayout()
        titles_text_layout.setSpacing(2)
        titles_text_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        title_text = QLabel("Интерактивное руководство пользователя")
        title_text.setStyleSheet("color: #ffffff; font-size: 17px; font-weight: 700; font-family: 'Segoe UI';")

        subtitle_text = QLabel("AgiAnalytics • Система учета и аналитики абитуриентов")
        subtitle_text.setStyleSheet("color: #94a3b8; font-size: 12px; font-weight: 400;")

        titles_text_layout.addWidget(title_text)
        titles_text_layout.addWidget(subtitle_text)
        header_layout.addLayout(titles_text_layout)

        header_layout.addStretch()

        # Прогресс прохождения
        progress_box = QVBoxLayout()
        progress_box.setSpacing(4)
        progress_box.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        self.progress_label = QLabel("Шаг 1 из 11")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.progress_label.setStyleSheet("color: #cbd5e1; font-size: 12px; font-weight: 600;")

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(5)
        self.progress_bar.setFixedWidth(150)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #334155;
                border: none;
                border-radius: 2px;
            }
            QProgressBar::chunk {
                background-color: #2563eb;
                border-radius: 2px;
            }
        """)

        progress_box.addWidget(self.progress_label)
        progress_box.addWidget(self.progress_bar)
        header_layout.addLayout(progress_box)

        # Кнопка закрытия
        close_btn = QPushButton("Закрыть")
        close_btn.setFixedSize(80, 32)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #1e293b;
                color: #e2e8f0;
                border: 1px solid #334155;
                border-radius: 6px;
                font-size: 12px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #334155;
                color: #ffffff;
            }
        """)
        close_btn.clicked.connect(self.accept)
        header_layout.addWidget(close_btn)

        main_layout.addWidget(header_widget)

        # ================= 2. ОСНОВНОЙ КОНТЕНТ =================
        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        # Левая панель — Навигация
        nav_widget = QWidget()
        nav_widget.setFixedWidth(270)
        nav_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border-right: 1px solid #e2e8f0;
            }
        """)
        nav_layout = QVBoxLayout(nav_widget)
        nav_layout.setContentsMargins(16, 16, 16, 16)
        nav_layout.setSpacing(12)

        # Поиск по разделам
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по справке...")
        self.search_input.setStyleSheet("""
            QLineEdit {
                background-color: #f8fafc;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 8px 12px;
                font-size: 13px;
                color: #0f172a;
            }
            QLineEdit:focus {
                border-color: #2563eb;
                background-color: #ffffff;
            }
        """)
        self.search_input.textChanged.connect(self.filter_navigation)
        nav_layout.addWidget(self.search_input)

        # Список тем
        self.nav_list = QListWidget()
        self.nav_list.setStyleSheet("""
            QListWidget {
                border: none;
                background-color: transparent;
                outline: none;
            }
            QListWidget::item {
                padding: 9px 12px;
                color: #475569;
                border-radius: 6px;
                font-size: 13px;
                font-weight: 500;
                margin-bottom: 2px;
            }
            QListWidget::item:hover {
                background-color: #f1f5f9;
                color: #0f172a;
            }
            QListWidget::item:selected {
                background-color: #eff6ff;
                color: #1d4ed8;
                font-weight: 600;
            }
        """)
        self.nav_list.itemClicked.connect(self.on_nav_clicked)
        nav_layout.addWidget(self.nav_list)

        # Кнопки Перехода
        nav_buttons_layout = QHBoxLayout()
        nav_buttons_layout.setSpacing(8)

        self.prev_btn = QPushButton("← Назад")
        self.prev_btn.setFixedHeight(36)
        self.prev_btn.setCursor(Qt.PointingHandCursor)
        self.prev_btn.setStyleSheet("""
            QPushButton {
                background-color: #f1f5f9;
                color: #334155;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                font-weight: 600;
                font-size: 12px;
            }
            QPushButton:hover:!disabled {
                background-color: #e2e8f0;
            }
            QPushButton:disabled {
                opacity: 0.4;
                color: #94a3b8;
                background-color: #f8fafc;
                border-color: #e2e8f0;
            }
        """)
        self.prev_btn.clicked.connect(self.prev_step)

        self.next_btn = QPushButton("Далее →")
        self.next_btn.setFixedHeight(36)
        self.next_btn.setCursor(Qt.PointingHandCursor)
        self.next_btn.setStyleSheet("""
            QPushButton {
                background-color: #2563eb;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                font-weight: 600;
                font-size: 12px;
            }
            QPushButton:hover:!disabled {
                background-color: #1d4ed8;
            }
            QPushButton:disabled {
                background-color: #93c5fd;
                color: #ffffff;
            }
        """)
        self.next_btn.clicked.connect(self.next_step)

        nav_buttons_layout.addWidget(self.prev_btn)
        nav_buttons_layout.addWidget(self.next_btn)
        nav_layout.addLayout(nav_buttons_layout)

        content_layout.addWidget(nav_widget)

        # Правая панель — Стек с контентом
        self.content_stack = QStackedWidget()
        self.content_stack.setStyleSheet("QStackedWidget { background-color: #f8fafc; }")
        content_layout.addWidget(self.content_stack, 1)

        main_layout.addWidget(content_widget, 1)

        # Инициализация всех шагов
        self._init_steps()

        # Загрузка первого шага
        self.go_to_step(0)

    def _init_steps(self):
        """Инициализация разделов и генераторов элементов"""
        self.steps = []

        steps_data = [
            ("01. Обзор программы", self._create_main_window),
            ("02. Добавление абитуриента", self._create_add_applicant),
            ("03. Форма: Данные абитуриента", self._create_add_form_part1),
            ("04. Форма: Данные агитатора", self._create_add_form_part2),
            ("05. Завершение формы", self._create_add_form_part3),
            ("06. Динамические поля", self._create_hidden_elements),
            ("07. Права и редактирование", self._create_edit_delete),
            ("08. Поиск и фильтрация", self._create_search),
            ("09. Модуль статистики", self._create_statistics),
            ("10. Импорт и экспорт Excel", self._create_import_export),
            ("11. План набора", self._create_plan_management),
        ]

        if self.role == 'admin':
            steps_data.append(("12. Настройки администратора", self._create_admin_settings))

        for title, func in steps_data:
            step_widget = QWidget()
            self.content_stack.addWidget(step_widget)

            self.steps.append({
                'widget': step_widget,
                'title': title,
                'func': func
            })

            item = QListWidgetItem(title)
            self.nav_list.addItem(item)

    def filter_navigation(self, text):
        """Фильтрация разделов по поисковому запросу"""
        text = text.lower().strip()
        for i in range(self.nav_list.count()):
            item = self.nav_list.item(i)
            item.setHidden(text not in item.text().lower())

    def on_nav_clicked(self, item):
        index = self.nav_list.row(item)
        self.go_to_step(index)

    def go_to_step(self, index):
        if index < 0 or index >= len(self.steps):
            return

        self.current_step = index
        self.nav_list.setCurrentRow(index)

        total_steps = len(self.steps)
        self.progress_label.setText(f"Раздел {index + 1} из {total_steps}")
        self.progress_bar.setMaximum(total_steps)
        self.progress_bar.setValue(index + 1)

        self.prev_btn.setEnabled(index > 0)
        self.next_btn.setEnabled(index < total_steps - 1)

        step_data = self.steps[index]
        widget = step_data['widget']

        # Очистка старого содержимого перед перерисовкой
        self._clear_widget(widget)

        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Вызов функции сборки макета
        step_data['func'](layout)

        self.content_stack.setCurrentWidget(widget)

    def _clear_widget(self, widget):
        if widget.layout():
            old_layout = widget.layout()
            while old_layout.count():
                item = old_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            QWidget().setLayout(old_layout)

    def next_step(self):
        if self.current_step < len(self.steps) - 1:
            self.go_to_step(self.current_step + 1)

    def prev_step(self):
        if self.current_step > 0:
            self.go_to_step(self.current_step - 1)

    # ================= СТРУКТУРНЫЕ КОМПОНЕНТЫ РАЗМЕТКИ =================

    def _create_scroll_content(self, layout, title, description):
        """Базовый контейнер прокрутки с заголовком страницы"""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(32, 28, 32, 28)
        container_layout.setSpacing(16)

        # Шапка карточки шага
        header_box = QHBoxLayout()
        header_box.setSpacing(12)

        title_label = QLabel(title)
        title_font = QFont()
        title_font.setPointSize(15)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #0f172a;")
        header_box.addWidget(title_label)
        header_box.addStretch()

        badge = QLabel(f"РАЗДЕЛ {self.current_step + 1:02d}")
        badge.setStyleSheet("""
            QLabel {
                background-color: #eff6ff;
                color: #2563eb;
                border: 1px solid #bfdbfe;
                border-radius: 4px;
                padding: 4px 10px;
                font-size: 11px;
                font-weight: 700;
                letter-spacing: 0.5px;
            }
        """)
        header_box.addWidget(badge)
        container_layout.addLayout(header_box)

        if description:
            desc_label = QLabel(description)
            desc_label.setStyleSheet("color: #475569; font-size: 13px; line-height: 1.5;")
            desc_label.setWordWrap(True)
            container_layout.addWidget(desc_label)

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setStyleSheet("background-color: #e2e8f0; max-height: 1px; border: none;")
        container_layout.addWidget(divider)

        items_container = QWidget()
        items_layout = QVBoxLayout(items_container)
        items_layout.setContentsMargins(0, 4, 0, 0)
        items_layout.setSpacing(12)
        container_layout.addWidget(items_container)

        container_layout.addStretch()
        scroll.setWidget(container)
        layout.addWidget(scroll)

        return items_layout

    def _add_action_step(self, parent_layout, number, title, action, location):
        """Интерактивная карточка последовательного шага"""
        widget = QFrame()
        widget.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border-radius: 8px;
                border: 1px solid #e2e8f0;
            }
            QFrame:hover {
                border-color: #cbd5e1;
            }
        """)

        layout = QHBoxLayout(widget)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(16)

        # Четкий цифровой индикатор
        num_label = QLabel(str(number))
        num_label.setFixedSize(30, 30)
        num_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        num_label.setStyleSheet("""
            QLabel {
                background-color: #2563eb;
                color: #ffffff;
                border-radius: 15px;
                font-size: 13px;
                font-weight: 700;
            }
        """)
        layout.addWidget(num_label)

        # Текстовое описание
        text_widget = QWidget()
        text_layout = QVBoxLayout(text_widget)
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(4)

        title_label = QLabel(title)
        title_label.setStyleSheet("font-weight: 700; color: #0f172a; font-size: 13px;")
        text_layout.addWidget(title_label)

        if action:
            action_label = QLabel(f"Действие: {action}")
            action_label.setStyleSheet("color: #2563eb; font-size: 12px; font-weight: 600;")
            action_label.setWordWrap(True)
            text_layout.addWidget(action_label)

        if location:
            location_label = QLabel(f"Расположение: {location}")
            location_label.setStyleSheet("color: #64748b; font-size: 11px;")
            text_layout.addWidget(location_label)

        layout.addWidget(text_widget, 1)
        parent_layout.addWidget(widget)

    def _add_highlight(self, parent_layout, title, content, level="info"):
        """Информационный блок без значков и эмодзи"""
        styles = {
            "tip": {"bg": "#f0fdf4", "border": "#16a34a", "title": "#15803d", "label": "РЕКОМЕНДАЦИЯ"},
            "info": {"bg": "#eff6ff", "border": "#2563eb", "title": "#1d4ed8", "label": "ИНФОРМАЦИЯ"},
            "warning": {"bg": "#fffbeb", "border": "#d97706", "title": "#b45309", "label": "ВНИМАНИЕ"},
            "danger": {"bg": "#fef2f2", "border": "#dc2626", "title": "#b91c1c", "label": "ОГРАНИЧЕНИЕ"},
        }
        cfg = styles.get(level, styles["info"])

        widget = QFrame()
        widget.setStyleSheet(f"""
            QFrame {{
                background-color: {cfg['bg']};
                border-left: 4px solid {cfg['border']};
                border-radius: 6px;
            }}
        """)

        layout = QVBoxLayout(widget)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(4)

        header_layout = QHBoxLayout()
        header_layout.setSpacing(8)

        badge_label = QLabel(cfg['label'])
        badge_label.setStyleSheet(f"""
            color: {cfg['title']};
            font-size: 10px;
            font-weight: 800;
            letter-spacing: 0.5px;
        """)
        header_layout.addWidget(badge_label)

        title_label = QLabel(title)
        title_label.setStyleSheet(f"color: {cfg['title']}; font-weight: 700; font-size: 12px;")
        header_layout.addWidget(title_label, 1)

        layout.addLayout(header_layout)

        content_label = QLabel(content)
        content_label.setStyleSheet("color: #334155; font-size: 12px; line-height: 1.4;")
        content_label.setWordWrap(True)
        layout.addWidget(content_label)

        parent_layout.addWidget(widget)

    def _add_info_grid(self, parent_layout, items, cols=2):
        """Плиточная сетка справочной информации"""
        grid_widget = QWidget()
        grid_layout = QGridLayout(grid_widget)
        grid_layout.setContentsMargins(0, 0, 0, 0)
        grid_layout.setSpacing(10)

        for i, (title, value) in enumerate(items):
            row = i // cols
            col = i % cols

            item_widget = QFrame()
            item_widget.setStyleSheet("""
                QFrame {
                    background-color: #ffffff;
                    border-radius: 6px;
                    border: 1px solid #e2e8f0;
                }
            """)

            item_layout = QVBoxLayout(item_widget)
            item_layout.setContentsMargins(12, 10, 12, 10)
            item_layout.setSpacing(3)

            title_label = QLabel(title)
            title_label.setStyleSheet("color: #64748b; font-size: 11px; font-weight: 700; text-transform: uppercase;")
            item_layout.addWidget(title_label)

            value_label = QLabel(value)
            value_label.setStyleSheet("color: #0f172a; font-size: 12px; font-weight: 600;")
            value_label.setWordWrap(True)
            item_layout.addWidget(value_label)

            grid_layout.addWidget(item_widget, row, col)

        parent_layout.addWidget(grid_widget)

    def _add_section_header(self, parent_layout, title):
        """Чистый подзаголовок раздела"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 10, 0, 4)

        label = QLabel(title)
        label.setStyleSheet("color: #0f172a; font-weight: 700; font-size: 13px;")
        layout.addWidget(label)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background-color: #cbd5e1; max-height: 1px; border: none;")
        layout.addWidget(line, 1)

        parent_layout.addWidget(widget)

    # ================= КОНТЕНТ РАЗДЕЛОВ СПРАВКИ =================

    def _create_main_window(self, layout):
        items = self._create_scroll_content(layout, "Обзор интерфейса программы",
                                            "AgiAnalytics предназначен для централизованного учета, анализа и контроля результатов агитационной работы.")

        self._add_action_step(items, 1, "Панель инструментов",
                              "Основные функции: Добавление, Редактирование, Удаление, Поиск, Экспорт и Импорт",
                              "Верхняя область основного окна")

        self._add_action_step(items, 2, "Вкладки разделов",
                              "Переключение между Реестром абитуриентов, Сводной статистикой и Настройками",
                              "Панель навигации под верхней строкой")

        self._add_action_step(items, 3, "Реестр абитуриентов",
                              "Просмотр записей в виде таблицы. Поддерживает сортировку по столбцам и выбор строки",
                              "Центральная часть рабочей области")

        self._add_highlight(items, "Статус подключения и учетная запись",
                            "В нижней строке состояния отображаются текущий пользователь, назначенное подразделение и системные уведомления.", "info")

    def _create_add_applicant(self, layout):
        items = self._create_scroll_content(layout, "Добавление абитуриента",
                                            "Порядок внесения новых персональных данных в базу.")

        self._add_action_step(items, 1, "Открытие диалога ввода",
                              "Нажмите кнопку 'Добавить' на панели инструментов",
                              "Панель инструментов")

        self._add_action_step(items, 2, "Заполнение обязательных полей",
                              "Заполните поля, отмеченные символом звёздочки (*)",
                              "Окно формы ввода")

        self._add_action_step(items, 3, "Сохранение записи",
                              "Нажмите кнопку 'Сохранить' для валидации и записи данных в БД",
                              "Нижняя панель формы ввода")

        self._add_highlight(items, "Автоматический контроль дубликатов",
                            "Система сверяет ФИО и дату рождения. Внесение идентичных записей блокируется во избежание дублирования.", "warning")

    def _create_add_form_part1(self, layout):
        items = self._create_scroll_content(layout, "Форма ввода: Данные абитуриента",
                                            "Описание параметров абитуриента.")

        self._add_action_step(items, 1, "ФИО абитуриента (*)", "Укажите полностью Фамилию, Имя и Отчество", "Первая строка формы")
        self._add_action_step(items, 2, "Регион РФ (*)", "Выберите субъект из выпадающего списка", "Выпадающий список")
        self._add_action_step(items, 3, "Контактный телефон (*)", "Введите номер с кодом региона", "Поле с автоформатом")

        self._add_info_grid(items, [
            ("Категория (*)", "Мужчина / Женщина / Военнослужащий"),
            ("Населенный пункт", "Наименование города, села или района"),
            ("Уровень образования", "Среднее общее, СПО, Высшее, СВУ"),
            ("Планирование поступления", "Поступает или Отказывается"),
            ("Статус документов", "Состояние дела (ВК, ОК, ВА ВКО)")
        ])

    def _create_add_form_part2(self, layout):
        items = self._create_scroll_content(layout, "Форма ввода: Данные агитатора",
                                            "Фиксация сведений о лице, проводившем агитацию.")

        self._add_action_step(items, 1, "Выбор категории агитатора",
                              "Установите переключатель 'Курсант' или 'Офицер / Военнослужащий'",
                              "Блок сведений об агитаторе")

        self._add_action_step(items, 2, "Подразделение (*)",
                              "Выберите соответствующий факультет или кафедру",
                              "Выпадающий список 'Подразделение'")

        self._add_action_step(items, 3, "ФИО агитатора (*)",
                              "Укажите фамилию и инициалы должностного лица",
                              "Поле ввода ФИО")

        self._add_highlight(items, "Агитатор — Курсант", "При выборе курсанта обязательно указываются Номер курса (1–5) и Номер учебной группы.", "tip")
        self._add_highlight(items, "Агитатор — Офицер", "При выборе офицера указывается воинское звание из утвержденного справочника.", "tip")

    def _create_add_form_part3(self, layout):
        items = self._create_scroll_content(layout, "Проверка и фиксация данных",
                                            "Завершение работы с формой.")

        self._add_action_step(items, 1, "Дополнительные примечания", "Внесите индивидуальные пометки при необходимости", "Текстовое поле примечаний")
        self._add_action_step(items, 2, "Сохранить запись", "Нажмите кнопку 'Сохранить' для фиксации в базе данных", "Нижняя панель")

        self._add_highlight(items, "Алгоритм валидации",
                            "При сохранении проверяются: полнота заполнения обязательных полей, корректность структуры телефона и отсутствие совпадений в БД.", "info")

    def _create_hidden_elements(self, layout):
        items = self._create_scroll_content(layout, "Динамическое поведение формы",
                                            "Автоматическая перестройка полей в зависимости от условий.")

        self._add_action_step(items, 1, "Адаптация полей агитатора",
                              "Состав полей меняется мгновенно при переключении статуса (Курсант / Офицер)",
                              "Форма ввода")

        self._add_action_step(items, 2, "Защищенные справочники",
                              "Выбор субъектов РФ и подразделений ограничен нормативным перечнем",
                              "Выпадающие списки")

        self._add_highlight(items, "Точность аналитики", "Корректный выбор категории агитатора критически важен для формирования отчетов по подразделениям.", "warning")

    def _create_edit_delete(self, layout):
        items = self._create_scroll_content(layout, "Управление записями",
                                            "Редактирование и удаление сведений.")

        self._add_action_step(items, 1, "Выбор строки", "Кликните по нужной записи в таблице", "Основная таблица")
        self._add_action_step(items, 2, "Редактирование", "Нажмите кнопку 'Редактировать' на верхней панели", "Панель инструментов")
        self._add_action_step(items, 3, "Удаление", "Нажмите кнопку 'Удалить' (требуется подтверждение)", "Панель инструментов")

        self._add_highlight(items, "Разграничение прав доступа",
                            "Пользователи с базовыми правами могут изменять только созданные ими записи. Администраторы обладают полным доступом ко всей базе данных.", "danger")

    def _create_search(self, layout):
        items = self._create_scroll_content(layout, "Поиск и фильтрация",
                                            "Быстрый поиск информации в базе данных.")

        self._add_action_step(items, 1, "Быстрый поиск", "Введите ФИО, город или телефон в поисковую строку", "Верхняя правая область")
        self._add_action_step(items, 2, "Фильтрация по курсу", "Выберите курс в выпадающем фильтре", "Панель фильтрации")
        self._add_action_step(items, 3, "Расширенный фильтр", "Используйте кнопку 'Расширенный поиск' для комбинированного отбора", "Рядом со строкой поиска")

        self._add_highlight(items, "Мгновенный отклик", "Фильтрация таблицы выполняется автоматически при вводе символов.", "tip")

    def _create_statistics(self, layout):
        items = self._create_scroll_content(layout, "Модуль аналитики и отчетов", "Анализ выполнения целевых показателей.")

        self._add_action_step(items, 1, "Переход в модуль", "Откройте вкладку 'Статистика'", "Главная панель навигации")
        self._add_action_step(items, 2, "Параметры фильтрации", "Задайте курс или категорию кандидатов", "Верхняя панель статистики")
        self._add_action_step(items, 3, "Детализация подразделений", "Раскройте карточку подразделения для детализации", "Сводная таблица")

        self._add_info_grid(items, [
            ("Метрики карточки", "Установленный план, Фактически привлечено, Выполнение (%)"),
            ("Анализ регионов", "Распределение кандидатов по закрепленным субъектам РФ")
        ])

    def _create_import_export(self, layout):
        items = self._create_scroll_content(layout, "Импорт и Экспорт файлов Excel", "Пакетная обработка данных.")

        self._add_action_step(items, 1, "Импорт реестра", "Нажмите 'Импорт' -> Выберите файл .xlsx -> Настройте сопоставление столбцов", "Панель инструментов")
        self._add_action_step(items, 2, "Экспорт отчета", "Нажмите 'Экспорт' -> Укажите каталог для сохранения файла", "Панель инструментов")

        self._add_highlight(items, "Требования к файлам",
                            "Мастер импорта поддерживает файлы с защитой и выбором листов. Обязательным является сопоставление колонок ФИО.", "info")

    def _create_plan_management(self, layout):
        items = self._create_scroll_content(layout, "Управление плановыми показателями", "Корректировка планов для подразделений.")

        self._add_action_step(items, 1, "Выбор подразделения", "Найдите требуемое подразделение на вкладке 'Статистика'", "Раздел статистики")
        self._add_action_step(items, 2, "Редактирование плана", "Нажмите 'Изменить план' внутри карточки подразделения", "Карточка подразделения")

        self._add_info_grid(items, [
            ("План: Категория М", "Целевой показатель для кандидатов-мужчин"),
            ("План: Категория Ж", "Целевой показатель для кандидатов-женщин"),
            ("План: Военнослужащие", "Целевой показатель для лиц, проходящих службу")
        ])

    def _create_admin_settings(self, layout):
        items = self._create_scroll_content(layout, "Администрирование системы",
                                            "Раздел настроек для пользователей с ролью 'Admin'.")

        self._add_action_step(items, 1, "Переход в настройки", "Откройте вкладку 'Настройки (Администратор)'", "Панель навигации")

        self._add_section_header(items, "1. Управление пользователями")
        self._add_action_step(items, "1.1", "Учетные записи", "Ведение списка пользователей, сброс паролей и назначение ролей", "Вкладка 'Пользователи'")
        self._add_action_step(items, "1.2", "Права доступа", "Настройка доступа пользователей к конкретным подразделениям", "Вкладка 'Права доступа'")

        self._add_section_header(items, "2. Ведение системных справочников")
        self._add_highlight(items, "Управляемые справочники",
                            "• Структура подразделений (Факультеты, Кафедры, Учебные группы)\n"
                            "• Закрепление регионов за подразделениями\n"
                            "• Справочник уровней образования и статусов документов\n"
                            "• График разрешенных периодов внесения данных", "tip")