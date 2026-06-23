# -*- coding: utf-8 -*-
"""
Модуль интерактивного диалога помощи
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QScrollArea, QWidget,
                             QFrame, QGridLayout, QListWidget,
                             QListWidgetItem, QStackedWidget)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon
from resource_helper import resource_path


class HelpDialog(QDialog):
    """Диалог интерактивной помощи"""

    def __init__(self, role, db, parent=None):
        super().__init__(parent)
        self.role = role
        self.db = db
        self.current_step = 0
        self.setModal(True)
        self.setWindowTitle('Интерактивная помощь')
        self.setMinimumSize(850, 650)
        self.setMaximumSize(1100, 750)
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f6fa;
            }
        """)
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Верхний бар
        header_widget = QWidget()
        header_widget.setFixedHeight(70)
        header_widget.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #6c5ce7, stop:1 #a29bfe);
            }
        """)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(30, 0, 30, 0)

        # Заголовок
        title_widget = QWidget()
        title_layout = QHBoxLayout(title_widget)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(12)

        icon_label = QLabel()
        icon_path = resource_path("icons/help.png")
        if icon_path:
            icon_label.setPixmap(QIcon(icon_path).pixmap(32, 32))
        else:
            icon_label.setText("?")
            icon_label.setStyleSheet("""
                QLabel {
                    background-color: rgba(255,255,255,0.2);
                    color: white;
                    border-radius: 18px;
                    font-size: 18px;
                    font-weight: bold;
                }
            """)
        icon_label.setFixedSize(36, 36)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_layout.addWidget(icon_label)

        title_text = QLabel("Интерактивная помощь")
        title_text.setStyleSheet("color: white; font-size: 20px; font-weight: bold;")
        title_layout.addWidget(title_text)

        header_layout.addWidget(title_widget)
        header_layout.addStretch()

        # Индикатор прогресса
        self.progress_label = QLabel("1 из 12")
        self.progress_label.setStyleSheet("color: rgba(255,255,255,0.8); font-size: 13px;")
        header_layout.addWidget(self.progress_label)

        # Кнопка закрытия
        close_btn = QPushButton("X")
        close_btn.setFixedSize(32, 32)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: rgba(255,255,255,0.1);
                color: white;
                border: none;
                border-radius: 16px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: rgba(255,255,255,0.2);
            }
        """)
        close_btn.clicked.connect(self.accept)
        header_layout.addWidget(close_btn)

        main_layout.addWidget(header_widget)

        # Основной контент
        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        # Левая панель - навигация
        nav_widget = QWidget()
        nav_widget.setFixedWidth(240)
        nav_widget.setStyleSheet("""
            QWidget {
                background-color: white;
                border-right: 1px solid #e9ecef;
            }
        """)
        nav_layout = QVBoxLayout(nav_widget)
        nav_layout.setContentsMargins(0, 20, 0, 20)
        nav_layout.setSpacing(0)

        nav_title = QLabel("Содержание")
        nav_title.setStyleSheet("""
            QLabel {
                color: #2d3436;
                font-size: 14px;
                font-weight: bold;
                padding: 0 20px 15px 20px;
            }
        """)
        nav_layout.addWidget(nav_title)

        self.nav_list = QListWidget()
        self.nav_list.setStyleSheet("""
            QListWidget {
                border: none;
                background-color: transparent;
                outline: none;
            }
            QListWidget::item {
                padding: 10px 20px;
                color: #636e72;
                border: none;
                border-left: 3px solid transparent;
            }
            QListWidget::item:hover {
                background-color: #f8f9fa;
            }
            QListWidget::item:selected {
                background-color: #f0edff;
                color: #6c5ce7;
                border-left: 3px solid #6c5ce7;
            }
        """)
        self.nav_list.itemClicked.connect(self.on_nav_clicked)
        nav_layout.addWidget(self.nav_list)

        nav_buttons_widget = QWidget()
        nav_buttons_layout = QHBoxLayout(nav_buttons_widget)
        nav_buttons_layout.setContentsMargins(20, 15, 20, 15)

        self.prev_btn = QPushButton("< Назад")
        self.prev_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfe6e9;
                color: #2d3436;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover:!disabled {
                background-color: #b2bec3;
            }
            QPushButton:disabled {
                opacity: 0.5;
            }
        """)
        self.prev_btn.clicked.connect(self.prev_step)

        self.next_btn = QPushButton("Далее >")
        self.next_btn.setStyleSheet("""
            QPushButton {
                background-color: #6c5ce7;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover:!disabled {
                background-color: #5f3dc4;
            }
            QPushButton:disabled {
                opacity: 0.5;
            }
        """)
        self.next_btn.clicked.connect(self.next_step)

        nav_buttons_layout.addWidget(self.prev_btn)
        nav_buttons_layout.addWidget(self.next_btn)
        nav_layout.addWidget(nav_buttons_widget)

        content_layout.addWidget(nav_widget)

        # Правая панель - контент
        self.content_stack = QStackedWidget()
        self.content_stack.setStyleSheet("""
            QStackedWidget {
                background-color: white;
            }
        """)
        content_layout.addWidget(self.content_stack, 1)

        main_layout.addWidget(content_widget, 1)

        # Инициализация шагов
        self._init_steps()

        # Переход на первый шаг
        self.go_to_step(0)

    def _init_steps(self):
        """Инициализация всех шагов"""
        self.steps = []

        steps_data = [
            ("Главное окно", self._create_main_window),
            ("Добавление абитуриента", self._create_add_applicant),
            ("Форма - Часть 1", self._create_add_form_part1),
            ("Форма - Часть 2", self._create_add_form_part2),
            ("Форма - Часть 3", self._create_add_form_part3),
            ("Скрытые элементы", self._create_hidden_elements),
            ("Редактирование", self._create_edit_delete),
            ("Поиск", self._create_search),
            ("Статистика", self._create_statistics),
            ("Импорт/Экспорт", self._create_import_export),
            ("Работа с планом", self._create_plan_management),
        ]

        if self.role == 'admin':
            steps_data.append(("Настройки админа", self._create_admin_settings))

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

    def on_nav_clicked(self, item):
        """Обработка клика по навигации"""
        index = self.nav_list.row(item)
        self.go_to_step(index)

    def go_to_step(self, index):
        """Переход к шагу"""
        if index < 0 or index >= len(self.steps):
            return

        self.current_step = index

        self.nav_list.setCurrentRow(index)
        self.progress_label.setText(f"{index + 1} из {len(self.steps)}")
        self.prev_btn.setEnabled(index > 0)
        self.next_btn.setEnabled(index < len(self.steps) - 1)

        step_data = self.steps[index]
        widget = step_data['widget']

        # Очищаем виджет полностью
        self._clear_widget(widget)

        # Создаем новый layout
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Создаем контент
        step_data['func'](layout)

        self.content_stack.setCurrentWidget(widget)

    def _clear_widget(self, widget):
        """Полная очистка виджета"""
        # Удаляем все дочерние виджеты
        for child in widget.findChildren(QWidget):
            child.deleteLater()

        # Удаляем layout
        if widget.layout():
            old_layout = widget.layout()
            # Удаляем все элементы из layout
            while old_layout.count():
                item = old_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            # Удаляем layout
            old_layout.deleteLater()

    def next_step(self):
        if self.current_step < len(self.steps) - 1:
            self.go_to_step(self.current_step + 1)

    def prev_step(self):
        if self.current_step > 0:
            self.go_to_step(self.current_step - 1)

    # ============ МЕТОДЫ СОЗДАНИЯ КОНТЕНТА ============

    def _create_scroll_content(self, layout, title, description):
        """Создание скролл-контента"""
        # Создаем скролл область
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
        """)

        # Контейнер для контента
        container = QWidget()
        container.setStyleSheet("background-color: transparent;")
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(30, 30, 30, 30)
        container_layout.setSpacing(20)

        # Заголовок
        title_widget = QWidget()
        title_widget_layout = QHBoxLayout(title_widget)
        title_widget_layout.setContentsMargins(0, 0, 0, 0)

        title_label = QLabel(title)
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #1a1a2e;")
        title_widget_layout.addWidget(title_label)
        title_widget_layout.addStretch()

        badge = QLabel("Шаг")
        badge.setStyleSheet("""
            QLabel {
                background-color: #6c5ce7;
                color: white;
                border-radius: 12px;
                padding: 4px 12px;
                font-size: 11px;
                font-weight: bold;
            }
        """)
        title_widget_layout.addWidget(badge)

        container_layout.addWidget(title_widget)

        if description:
            desc_label = QLabel(description)
            desc_label.setStyleSheet("color: #636e72; font-size: 14px;")
            desc_label.setWordWrap(True)
            container_layout.addWidget(desc_label)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background-color: #dfe6e9; max-height: 1px;")
        container_layout.addWidget(line)

        # Контейнер для элементов
        items_container = QWidget()
        items_layout = QVBoxLayout(items_container)
        items_layout.setSpacing(12)
        container_layout.addWidget(items_container)

        container_layout.addStretch()

        scroll.setWidget(container)
        layout.addWidget(scroll)

        return items_layout

    def _add_action_step(self, parent_layout, number, title, action, location, color="#2ecc71"):
        """Добавление шага действия"""
        widget = QWidget()
        widget.setStyleSheet(f"""
            QWidget {{
                background-color: #f0fff4;
                border-radius: 10px;
                border: 1px solid #b2dfdb;
            }}
        """)

        layout = QHBoxLayout(widget)
        layout.setContentsMargins(15, 12, 15, 12)
        layout.setSpacing(15)

        num_label = QLabel(str(number))
        num_label.setFixedSize(30, 30)
        num_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        num_label.setStyleSheet(f"""
            QLabel {{
                background-color: {color};
                color: white;
                border-radius: 15px;
                font-size: 14px;
                font-weight: bold;
            }}
        """)
        layout.addWidget(num_label)

        text_widget = QWidget()
        text_layout = QVBoxLayout(text_widget)
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(4)

        title_label = QLabel(title)
        title_label.setStyleSheet("font-weight: bold; color: #2d3436; font-size: 13px;")
        text_layout.addWidget(title_label)

        if action:
            action_label = QLabel(f"Действие: {action}")
            action_label.setStyleSheet("color: #6c5ce7; font-size: 12px; font-weight: bold;")
            text_layout.addWidget(action_label)

        if location:
            location_label = QLabel(f"Где: {location}")
            location_label.setStyleSheet("color: #636e72; font-size: 12px;")
            text_layout.addWidget(location_label)

        layout.addWidget(text_widget, 1)
        parent_layout.addWidget(widget)

    def _add_highlight(self, parent_layout, title, content, color="#fdcb6e"):
        """Добавление выделенного блока"""
        widget = QFrame()
        widget.setStyleSheet(f"""
            QFrame {{
                background-color: #fff8e7;
                border-left: 4px solid {color};
                border-radius: 8px;
                padding: 12px;
            }}
        """)

        layout = QVBoxLayout(widget)
        layout.setSpacing(6)

        title_label = QLabel(title)
        title_label.setStyleSheet(f"color: #2d3436; font-weight: bold; font-size: 13px;")
        layout.addWidget(title_label)

        content_label = QLabel(content)
        content_label.setStyleSheet("color: #636e72; font-size: 12px;")
        content_label.setWordWrap(True)
        layout.addWidget(content_label)

        parent_layout.addWidget(widget)

    def _add_info_grid(self, parent_layout, items, cols=2):
        """Добавление сетки информации"""
        grid_widget = QWidget()
        grid_layout = QGridLayout(grid_widget)
        grid_layout.setSpacing(10)

        for i, (title, value) in enumerate(items):
            row = i // cols
            col = i % cols

            item_widget = QWidget()
            item_widget.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border-radius: 8px;
                    border: 1px solid #e9ecef;
                }
            """)

            item_layout = QVBoxLayout(item_widget)
            item_layout.setContentsMargins(12, 10, 12, 10)
            item_layout.setSpacing(4)

            title_label = QLabel(title)
            title_label.setStyleSheet(
                "color: #b2bec3; font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px;")
            item_layout.addWidget(title_label)

            value_label = QLabel(value)
            value_label.setStyleSheet("color: #2d3436; font-size: 13px; font-weight: 500;")
            value_label.setWordWrap(True)
            item_layout.addWidget(value_label)

            grid_layout.addWidget(item_widget, row, col)

        parent_layout.addWidget(grid_widget)

    # ============ КОНТЕНТ ШАГОВ ============

    def _create_main_window(self, layout):
        items = self._create_scroll_content(layout, "Добро пожаловать в AgiAnalytics",
                                            "Ваш помощник в работе с абитуриентами")

        self._add_action_step(items, 1, "Панель инструментов вверху",
                              "Кнопки: Добавить, Редактировать, Удалить, Импорт, Экспорт, Помощь, Выход",
                              "Верхняя часть окна")

        self._add_action_step(items, 2, "Вкладки",
                              "Кликните по названию вкладки: 'Данные абитуриентов', 'Статистика'",
                              "Под панелью инструментов")

        self._add_action_step(items, 3, "Таблица с данными",
                              "Кликните по строке для выделения",
                              "Центральная область вкладки 'Данные абитуриентов'")

        self._add_highlight(items, "Совет",
                            "Панель инструментов содержит все основные действия. Статусная строка внизу показывает информацию о текущем пользователе.")

    def _create_add_applicant(self, layout):
        items = self._create_scroll_content(layout, "Добавление абитуриента", "Пошаговая инструкция")

        self._add_action_step(items, 1, "Нажмите кнопку 'Добавить'",
                              "Кликните по кнопке с иконкой документа на панели инструментов",
                              "Панель инструментов (первая кнопка слева)")

        self._add_action_step(items, 2, "Заполните форму",
                              "Введите данные абитуриента и агитатора",
                              "Диалоговое окно 'Добавить абитуриента'")

        self._add_action_step(items, 3, "Нажмите 'Сохранить'",
                              "Кликните по зеленой кнопке 'Сохранить' внизу диалога",
                              "Нижняя часть диалогового окна")

        self._add_highlight(items, "Важно",
                            "Все поля с пометкой * обязательны для заполнения. Нельзя добавить абитуриента с уже существующим ФИО.")

    def _create_add_form_part1(self, layout):
        items = self._create_scroll_content(layout, "Форма добавления - Блок 1", "Информация об абитуриенте")

        self._add_action_step(items, 1, "ФИО абитуриента *",
                              "Введите полное имя в текстовое поле",
                              "Первое поле в блоке 'Информация об абитуриенте'")

        self._add_action_step(items, 2, "Субъект РФ *",
                              "Нажмите на стрелку и выберите регион из списка",
                              "Второе поле - выпадающий список")

        self._add_action_step(items, 3, "Категория *",
                              "Нажмите на стрелку и выберите категорию",
                              "Поле 'Категория'")

        self._add_action_step(items, 4, "Телефон *",
                              "Введите номер (маска подставится автоматически)",
                              "Поле 'Телефон' с маской ввода")

        self._add_info_grid(items, [
            ("Населенный пункт", "Введите город или село"),
            ("Образование", "Выберите из выпадающего списка"),
            ("Статус *", "Выберите 'Поступает' или 'Отказывается'"),
            ("Документы", "Выберите статус документов")
        ])

    def _create_add_form_part2(self, layout):
        items = self._create_scroll_content(layout, "Форма добавления - Блок 2", "Информация об агитаторе")

        self._add_action_step(items, 1, "Тип агитатора",
                              "Нажмите на чекбокс 'Агитатор - курсант' или 'Агитатор - офицер/военнослужащий'",
                              "Раздел 'Информация об агитаторе'")

        self._add_action_step(items, 2, "Подразделение *",
                              "Нажмите на стрелку и выберите подразделение из списка",
                              "Поле 'Подразделение' - выпадающий список")

        self._add_action_step(items, 3, "ФИО агитатора *",
                              "Введите полное имя агитатора",
                              "Поле 'ФИО агитатора'")

        self._add_highlight(items, "Для курсанта",
                            "После выбора чекбокса 'Агитатор - курсант' появятся поля: Курс (выберите из списка) и Группа (введите номер)")
        self._add_highlight(items, "Для офицера",
                            "После выбора чекбокса 'Агитатор - офицер/военнослужащий' появится поле: Звание (выберите из списка)")

    def _create_add_form_part3(self, layout):
        items = self._create_scroll_content(layout, "Форма добавления - Блок 3",
                                            "Дополнительная информация и сохранение")

        self._add_action_step(items, 1, "Примечания",
                              "Введите комментарии в текстовое поле",
                              "Блок 'Дополнительная информация'")

        self._add_action_step(items, 2, "Сохранение",
                              "Нажмите зеленую кнопку 'Сохранить'",
                              "Нижняя часть диалогового окна")

        self._add_action_step(items, 3, "Отмена",
                              "Нажмите серую кнопку 'Отмена' для закрытия без сохранения",
                              "Нижняя часть диалогового окна")

        self._add_highlight(items, "Что происходит после сохранения",
                            "1. Проверка всех обязательных полей\n2. Проверка дубликатов по ФИО\n3. Добавление в базу данных\n4. Обновление таблицы и статистики")

    def _create_hidden_elements(self, layout):
        items = self._create_scroll_content(layout, "Скрытые элементы формы",
                                            "Обратите внимание на дополнительные возможности")

        self._add_action_step(items, 1, "Чекбокс 'Агитатор - курсант'",
                              "Нажмите на чекбокс для активации полей 'Курс' и 'Группа'",
                              "Раздел 'Информация об агитаторе'")

        self._add_action_step(items, 2, "Чекбокс 'Агитатор - офицер/военнослужащий'",
                              "Нажмите на чекбокс для активации поля 'Звание'",
                              "Раздел 'Информация об агитаторе'")

        self._add_action_step(items, 3, "Выбор подразделения",
                              "Нажмите на стрелку и выберите из списка. Вручную ввести нельзя!",
                              "Поле 'Подразделение'")

        self._add_action_step(items, 4, "Выбор субъекта РФ",
                              "Нажмите на стрелку и выберите регион из списка. Вручную ввести нельзя!",
                              "Поле 'Субъект РФ'")

        self._add_highlight(items, "Подсказка",
                            "Курсант - студент, ведущий агитацию. Офицер/военнослужащий - действующий военнослужащий. Правильный выбор важен для статистики.")

    def _create_edit_delete(self, layout):
        items = self._create_scroll_content(layout, "Редактирование и удаление", "Управление существующими записями")

        self._add_action_step(items, 1, "Выделите запись",
                              "Кликните по строке в таблице",
                              "Таблица на вкладке 'Данные абитуриентов'")

        self._add_action_step(items, 2, "Редактирование",
                              "Нажмите кнопку 'Редактировать' (иконка карандаша)",
                              "Панель инструментов (вторая кнопка)")

        self._add_action_step(items, 3, "Удаление",
                              "Нажмите кнопку 'Удалить' (иконка корзины)",
                              "Панель инструментов (третья кнопка)")

        self._add_highlight(items, "Права доступа",
                            "Вы можете редактировать и удалять только свои записи. Администратор имеет доступ ко всем записям.")

    def _create_search(self, layout):
        items = self._create_scroll_content(layout, "Поиск абитуриентов", "Как быстро найти нужные записи")

        self._add_action_step(items, 1, "Обычный поиск",
                              "Введите текст в поле 'Поиск'",
                              "Правая часть панели инструментов")

        self._add_action_step(items, 2, "Расширенный поиск",
                              "Нажмите кнопку 'Расширенный поиск' (фиолетовая)",
                              "Рядом с полем поиска")

        self._add_action_step(items, 3, "Сброс фильтров",
                              "Нажмите кнопку 'Сбросить фильтры' (оранжевая)",
                              "Рядом с кнопкой 'Расширенный поиск'")

        self._add_highlight(items, "Возможности расширенного поиска",
                            "Поиск по ФИО, региону, категории, статусу, ФИО агитатора, подразделению, курсу")

    def _create_statistics(self, layout):
        items = self._create_scroll_content(layout, "Статистика", "Анализ данных по подразделениям")

        self._add_action_step(items, 1, "Перейдите на вкладку 'Статистика'",
                              "Кликните по названию вкладки",
                              "Верхняя часть окна (вторая вкладка)")

        self._add_action_step(items, 2, "Просмотрите общую сводку",
                              "Вверху показаны общие показатели",
                              "Верхняя часть вкладки 'Статистика'")

        self._add_action_step(items, 3, "Раскройте карточку подразделения",
                              "Нажмите на стрелку (▶) в карточке",
                              "Карточка подразделения в центральной части")

        self._add_action_step(items, 4, "Статистика по регионам",
                              "Нажмите кнопку 'Статистика по регионам'",
                              "Внутри карточки подразделения (фиолетовая кнопка)")

        self._add_info_grid(items, [
            ("В карточке", "План, Всего, Отобраны, % выполнения"),
            ("Внутри карточки", "План по категориям, Отобраны, Статусы документов")
        ])

    def _create_import_export(self, layout):
        items = self._create_scroll_content(layout, "Импорт и экспорт данных", "Работа с Excel файлами")

        self._add_action_step(items, 1, "Импорт из Excel",
                              "Нажмите кнопку 'Импорт из Excel'",
                              "Панель инструментов (кнопка с иконкой импорта)")

        self._add_action_step(items, 2, "Экспорт данных",
                              "Нажмите кнопку 'Экспорт'",
                              "Панель инструментов (кнопка с иконкой экспорта)")

        self._add_highlight(items, "Шаги импорта",
                            "1. Выберите файл\n2. Укажите пароль\n3. Выберите листы\n4. Настройте сопоставление\n5. Нажмите 'Начать импорт'")
        self._add_highlight(items, "Совет",
                            "Обязательные поля: ФИО абитуриента и ФИО агитатора. Для каждого листа своё сопоставление.")

    def _create_plan_management(self, layout):
        items = self._create_scroll_content(layout, "Управление планом набора",
                                            "Редактирование плана для подразделений")

        self._add_action_step(items, 1, "Найдите подразделение",
                              "Перейдите на вкладку 'Статистика'",
                              "Вкладка 'Статистика' - карточки подразделений")

        self._add_action_step(items, 2, "Раскройте карточку",
                              "Нажмите на стрелку (▶)",
                              "В карточке подразделения")

        self._add_action_step(items, 3, "Нажмите 'Редактировать план'",
                              "Кликните по оранжевой кнопке",
                              "Внутри раскрытой карточки подразделения")

        self._add_highlight(items, "Кто может редактировать",
                            "Администратор - любой план. Начальник подразделения - только свой план. Пользователь - не может редактировать.")

        self._add_info_grid(items, [
            ("План по мужчинам (М)", "Количество"),
            ("План по женщинам (Ж)", "Количество"),
            ("План по военнослужащим (в/сл)", "Количество")
        ])

    def _create_admin_settings(self, layout):
        items = self._create_scroll_content(layout, "Настройки администратора", "Управление системой - полный контроль")

        # Переход к настройкам
        self._add_action_step(items, 1, "Перейдите на вкладку 'Настройки (админ)'",
                              "Кликните по названию вкладки (доступна только администраторам)",
                              "Верхняя часть окна (третья вкладка)")

        self._add_highlight(items, "Доступные разделы",
                            "Внутри вкладки 'Настройки (админ)' доступны следующие разделы:\n"
                            "1. Пользователи - управление учетными записями\n"
                            "2. Подразделения - структура организации\n"
                            "3. Регионы - справочник субъектов РФ\n"
                            "4. Образование - типы образования\n"
                            "5. Статусы документов - статусы документов\n"
                            "6. Расписание - рабочие дни\n"
                            "7. Ответственные за регионы - назначение регионов подразделениям")

        # ===== РАЗДЕЛ 1: ПОЛЬЗОВАТЕЛИ =====
        self._add_section_header(items, "1. Управление пользователями")

        self._add_action_step(items, "1.1", "Добавить пользователя",
                              "Нажмите кнопку 'Добавить пользователя' (зеленая кнопка с плюсом)",
                              "Вкладка 'Пользователи' в настройках")

        self._add_action_step(items, "1.2", "Редактировать пользователя",
                              "Выделите строку с пользователем и нажмите 'Редактировать' (синяя кнопка)",
                              "Вкладка 'Пользователи' - таблица пользователей")

        self._add_action_step(items, "1.3", "Удалить пользователя",
                              "Выделите строку с пользователем и нажмите 'Удалить' (красная кнопка)",
                              "Вкладка 'Пользователи' - таблица пользователей")

        self._add_highlight(items, "Поля при создании пользователя",
                            "• Логин * - уникальное имя для входа\n"
                            "• Пароль * - пароль для входа\n"
                            "• ФИО * - полное имя пользователя\n"
                            "• Роль - 'Администратор' или 'Пользователь'\n"
                            "• Подразделение - к какому подразделению привязан\n"
                            "• Должность - должность пользователя\n"
                            "• Звание - воинское звание (если есть)\n"
                            "• Начальник подразделения - чекбокс для назначения начальником\n"
                            "• Права доступа - выбор подразделений для просмотра")

        # ===== РАЗДЕЛ 2: ПОДРАЗДЕЛЕНИЯ =====
        self._add_section_header(items, "2. Управление подразделениями")

        self._add_action_step(items, "2.1", "Добавить подразделение",
                              "Нажмите кнопку 'Добавить подразделение' (зеленая кнопка)",
                              "Вкладка 'Подразделения' в настройках")

        self._add_action_step(items, "2.2", "Редактировать подразделение",
                              "Выделите строку и нажмите 'Редактировать' (синяя кнопка)",
                              "Вкладка 'Подразделения' - таблица подразделений")

        self._add_action_step(items, "2.3", "Удалить подразделение",
                              "Выделите строку и нажмите 'Удалить' (красная кнопка)",
                              "Вкладка 'Подразделения' - таблица подразделений")

        self._add_highlight(items, "Типы подразделений",
                            "• Факультет - основное подразделение (например, 'Факультет 1')\n"
                            "• Кафедра - подразделение внутри факультета\n"
                            "• Группа - группа внутри кафедры (можно выбрать родительское подразделение)\n"
                            "• Начальник подразделения - назначается из списка пользователей")

        # ===== РАЗДЕЛ 3: РЕГИОНЫ =====
        self._add_section_header(items, "3. Управление регионами")

        self._add_action_step(items, "3.1", "Добавить регион",
                              "Нажмите кнопку 'Добавить регион' (зеленая кнопка)",
                              "Вкладка 'Регионы' в настройках")

        self._add_action_step(items, "3.2", "Удалить регион",
                              "Выделите регион в списке и нажмите 'Удалить регион' (красная кнопка)",
                              "Вкладка 'Регионы' - список регионов")

        self._add_highlight(items, "Назначение регионов подразделениям",
                            "Перейдите на вкладку 'Ответственные за регионы'.\n"
                            "Выберите подразделение и нажмите 'Добавить регион' для назначения.\n"
                            "Регион будет закреплен за подразделением для сбора статистики.")

        # ===== РАЗДЕЛ 4: ОБРАЗОВАНИЕ =====
        self._add_section_header(items, "4. Управление образованием")

        self._add_action_step(items, "4.1", "Добавить тип образования",
                              "Нажмите кнопку 'Добавить' (зеленая кнопка)",
                              "Вкладка 'Образование' в настройках")

        self._add_action_step(items, "4.2", "Удалить тип образования",
                              "Выделите тип в списке и нажмите 'Удалить' (красная кнопка)",
                              "Вкладка 'Образование' - список типов")

        self._add_highlight(items, "Типы образования по умолчанию",
                            "• СОШ - средняя общеобразовательная школа\n"
                            "• СПО - среднее профессиональное образование\n"
                            "• СВУ - суворовское военное училище\n"
                            "• ПКУ - президентское кадетское училище\n"
                            "• КК - кадетский корпус")

        # ===== РАЗДЕЛ 5: СТАТУСЫ ДОКУМЕНТОВ =====
        self._add_section_header(items, "5. Управление статусами документов")

        self._add_action_step(items, "5.1", "Добавить статус документов",
                              "Нажмите кнопку 'Добавить' (зеленая кнопка)",
                              "Вкладка 'Статусы документов' в настройках")

        self._add_action_step(items, "5.2", "Удалить статус документов",
                              "Выделите статус в списке и нажмите 'Удалить' (красная кнопка)",
                              "Вкладка 'Статусы документов' - список статусов")

        self._add_highlight(items, "Статусы документов по умолчанию",
                            "• ВК - документы в военкомате\n"
                            "• ОК - документы в отделе кадров\n"
                            "• ВА ВКО - документы в военной академии")

        # ===== РАЗДЕЛ 6: РАСПИСАНИЕ =====
        self._add_section_header(items, "6. Настройка расписания")

        self._add_action_step(items, "6.1", "Выбрать рабочие дни",
                              "Отметьте чекбоксы для дней, когда разрешено добавлять записи",
                              "Вкладка 'Расписание' в настройках")

        self._add_action_step(items, "6.2", "Сохранить настройки",
                              "Нажмите кнопку 'Сохранить настройки' (зеленая кнопка)",
                              "Вкладка 'Расписание' - внизу")

        self._add_highlight(items, "Назначение расписания",
                            "В выбранные рабочие дни пользователи могут добавлять новых абитуриентов.\n"
                            "В выходные дни добавление записей будет запрещено.\n"
                            "По умолчанию: понедельник - пятница (рабочие дни).")

        # ===== РАЗДЕЛ 7: ОТВЕТСТВЕННЫЕ ЗА РЕГИОНЫ =====
        self._add_section_header(items, "7. Ответственные за регионы")

        self._add_action_step(items, "7.1", "Выбрать подразделение",
                              "Выберите подразделение из выпадающего списка",
                              "Вкладка 'Ответственные за регионы'")

        self._add_action_step(items, "7.2", "Добавить регион подразделению",
                              "Нажмите кнопку 'Добавить регион' и выберите регион из списка",
                              "Вкладка 'Ответственные за регионы'")

        self._add_action_step(items, "7.3", "Удалить регион у подразделения",
                              "Выберите регион в списке и нажмите 'Удалить регион'",
                              "Вкладка 'Ответственные за регионы' - список регионов")

        self._add_highlight(items, "Назначение",
                            "Закрепление регионов за подразделениями позволяет:\n"
                            "• Отслеживать эффективность работы в конкретных регионах\n"
                            "• Анализировать статистику по регионам для каждого подразделения\n"
                            "• Выявлять наиболее успешные регионы для агитации")

        # Итоговая информация
        self._add_section_header(items, "Важная информация")
        self._add_highlight(items, "Права доступа",
                            "• Администратор - полный доступ ко всем функциям системы\n"
                            "• Начальник подразделения - управление своим подразделением\n"
                            "• Пользователь - работа с абитуриентами (свои записи)")

        self._add_highlight(items, "Советы администратору",
                            "1. Регулярно проверяйте список пользователей и их права\n"
                            "2. Назначайте начальников подразделений для распределения ответственности\n"
                            "3. Поддерживайте справочники в актуальном состоянии\n"
                            "4. Настраивайте расписание в соответствии с рабочими днями\n"
                            "5. Закрепляйте регионы за подразделениями для точной статистики")

    def _add_section_header(self, parent_layout, title):
        """Добавление заголовка раздела"""
        widget = QWidget()
        widget.setStyleSheet("""
            QWidget {
                background-color: #f0edff;
                border-radius: 8px;
                padding: 8px;
                margin-top: 10px;
            }
        """)

        layout = QHBoxLayout(widget)
        layout.setContentsMargins(10, 8, 10, 8)

        label = QLabel(title)
        label.setStyleSheet("color: #6c5ce7; font-weight: bold; font-size: 14px;")
        layout.addWidget(label)
        layout.addStretch()

        parent_layout.addWidget(widget)