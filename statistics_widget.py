# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QComboBox, QPushButton, QScrollArea, QFrame,
                             QGridLayout, QSizePolicy)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtWidgets import QDialog, QDialogButtonBox, QFormLayout, QMessageBox


class CourseSection(QFrame):
    """Секция для отображения статистики по курсу"""

    def __init__(self, course_name, stats, parent=None):
        super().__init__(parent)
        self.course_name = course_name
        self.stats = stats
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        self.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 10px;
                border: 1px solid #e0e0e0;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 20)
        layout.setSpacing(15)

        # Заголовок курса
        header_widget = QWidget()
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 0)

        # Иконка и название курса
        course_container = QWidget()
        course_layout = QHBoxLayout(course_container)
        course_layout.setContentsMargins(0, 0, 0, 0)
        course_layout.setSpacing(10)

        icon_label = QLabel("🎓")
        icon_label.setStyleSheet("font-size: 24px;")
        course_layout.addWidget(icon_label)

        course_label = QLabel(self.course_name)
        course_font = QFont()
        course_font.setPointSize(16)
        course_font.setBold(True)
        course_label.setFont(course_font)
        course_label.setStyleSheet("color: #2c3e50;")
        course_layout.addWidget(course_label)

        # Если есть факультет, добавляем его
        if self.stats.get('faculty'):
            faculty_label = QLabel(f"({self.stats.get('faculty')})")
            faculty_label.setStyleSheet("color: #7f8c8d; font-size: 14px;")
            course_layout.addWidget(faculty_label)

        header_layout.addWidget(course_container)
        header_layout.addStretch()

        # Бейдж с общим количеством
        total_badge = QLabel(f"Всего: {self.stats.get('total', 0)}")
        total_badge.setStyleSheet("""
            QLabel {
                background-color: #3498db;
                color: white;
                border-radius: 12px;
                padding: 6px 12px;
                font-weight: bold;
            }
        """)
        header_layout.addWidget(total_badge)

        layout.addWidget(header_widget)

        # Сетка карточек статистики
        cards_widget = QWidget()
        cards_layout = QGridLayout(cards_widget)
        cards_layout.setSpacing(15)
        cards_layout.setContentsMargins(0, 0, 0, 0)

        # Основные карточки
        main_cards = [
            ('👥 Всего', 'total', '#3498db'),
            ('✅ Поступают', 'applying', '#2ecc71'),
            ('❌ Отказались', 'refused', '#e74c3c'),
            ('👨 Мужчины', 'male', '#9b59b6'),
            ('👩 Женщины', 'female', '#e67e22'),
            ('🎖️ Военнослужащие', 'military', '#1abc9c'),
        ]

        for i, (title, key, color) in enumerate(main_cards):
            row = i // 3
            col = i % 3
            value = self.stats.get(key, 0)
            card = StatisticsCard(title, value, color)
            cards_layout.addWidget(card, row, col)

        layout.addWidget(cards_widget)

        # Статус документов (если есть данные)
        if any(key in self.stats for key in ['doc1', 'doc2', 'doc3']):
            docs_header = QLabel("📄 Статус документов:")
            docs_header.setStyleSheet("""
                QLabel {
                    color: #2c3e50;
                    font-weight: bold;
                    font-size: 14px;
                    padding-top: 10px;
                    border-top: 1px solid #eee;
                    margin-top: 5px;
                }
            """)
            layout.addWidget(docs_header)

            docs_widget = QWidget()
            docs_layout = QHBoxLayout(docs_widget)
            docs_layout.setSpacing(15)
            docs_layout.setContentsMargins(0, 10, 0, 0)

            doc_cards = [
                ('📋 Формируется', 'doc1', '#f39c12'),
                ('📤 Отправлено', 'doc2', '#8e44ad'),
                ('📥 В ВА ВКО', 'doc3', '#16a085'),
            ]

            for title, key, color in doc_cards:
                value = self.stats.get(key, 0)
                card = StatisticsCard(title, value, color)
                card.setMinimumWidth(160)
                docs_layout.addWidget(card)

            docs_layout.addStretch()
            layout.addWidget(docs_widget)

        # Визуализация (простая круговая диаграмма)
        if self.stats.get('total', 0) > 0:
            self.add_chart_section(layout)

    def add_chart_section(self, layout):
        """Добавление секции с диаграммами"""
        chart_widget = QWidget()
        chart_layout = QHBoxLayout(chart_widget)
        chart_layout.setSpacing(20)

        # Создаем простые круговые диаграммы
        charts = [
            ('Статус поступления', ['Поступают', 'Отказались'],
             [self.stats.get('applying', 0), self.stats.get('refused', 0)],
             ['#2ecc71', '#e74c3c']),

            ('Распределение по категориям', ['Мужчины', 'Женщины', 'Военнослужащие'],
             [self.stats.get('male', 0), self.stats.get('female', 0), self.stats.get('military', 0)],
             ['#3498db', '#e67e22', '#1abc9c'])
        ]

        for title, labels, data, colors in charts:
            if sum(data) > 0:
                chart_container = QFrame()
                chart_container.setStyleSheet("""
                    QFrame {
                        background-color: #f8f9fa;
                        border-radius: 8px;
                        border: 1px solid #e9ecef;
                    }
                """)
                chart_container.setFixedSize(280, 220)

                chart_inner = QVBoxLayout(chart_container)
                chart_inner.setContentsMargins(10, 10, 10, 10)

                chart_title = QLabel(title)
                chart_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
                chart_title.setStyleSheet("""
                    QLabel {
                        font-weight: bold;
                        color: #495057;
                        margin-bottom: 10px;
                    }
                """)
                chart_inner.addWidget(chart_title)

                # Создаем простую текстовую визуализацию
                text_widget = QWidget()
                text_layout = QVBoxLayout(text_widget)
                text_layout.setSpacing(5)

                for label, value, color in zip(labels, data, colors):
                    if value > 0:
                        item_widget = QWidget()
                        item_layout = QHBoxLayout(item_widget)
                        item_layout.setContentsMargins(5, 2, 5, 2)

                        color_indicator = QLabel("⬤")
                        color_indicator.setStyleSheet(f"color: {color}; font-size: 10px;")

                        label_text = QLabel(f"{label}: {value}")
                        label_text.setStyleSheet("color: #6c757d; font-size: 12px;")

                        item_layout.addWidget(color_indicator)
                        item_layout.addWidget(label_text)
                        item_layout.addStretch()

                        text_layout.addWidget(item_widget)

                chart_inner.addWidget(text_widget)
                chart_layout.addWidget(chart_container)

        if chart_layout.count() > 0:
            chart_layout.addStretch()
            layout.addWidget(chart_widget)


class TotalStatisticsSection(QFrame):
    """Секция с общей статистикой (только для админа)"""

    def __init__(self, stats, parent=None):
        super().__init__(parent)
        self.stats = stats
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        self.setStyleSheet("""
            QFrame {
                background-color: #2c3e50;
                border-radius: 10px;
                border: none;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 25)
        layout.setSpacing(20)

        # Заголовок
        header_widget = QWidget()
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 0)

        title_label = QLabel("📊 Общая статистика (все курсы)")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: white;")

        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Бейдж с общим количеством
        total_badge = QLabel(f"Всего абитуриентов: {self.stats.get('total', 0)}")
        total_badge.setStyleSheet("""
            QLabel {
                background-color: #e74c3c;
                color: white;
                border-radius: 15px;
                padding: 8px 16px;
                font-weight: bold;
                font-size: 14px;
            }
        """)
        header_layout.addWidget(total_badge)

        layout.addWidget(header_widget)

        # Сетка карточек статистики
        cards_widget = QWidget()
        cards_layout = QGridLayout(cards_widget)
        cards_layout.setSpacing(15)
        cards_layout.setContentsMargins(0, 0, 0, 0)

        # Основные карточки общей статистики
        main_cards = [
            ('👥 Всего абитуриентов', 'total', '#3498db'),
            ('✅ Поступают', 'applying', '#2ecc71'),
            ('❌ Отказались', 'refused', '#e74c3c'),
            ('👨 Мужчины', 'male', '#9b59b6'),
            ('👩 Женщины', 'female', '#e67e22'),
            ('🎖️ Военнослужащие', 'military', '#1abc9c'),
        ]

        for i, (title, key, color) in enumerate(main_cards):
            row = i // 3
            col = i % 3
            value = self.stats.get(key, 0)
            card = StatisticsCard(title, value, color)
            cards_layout.addWidget(card, row, col)

        layout.addWidget(cards_widget)

        # Процентное соотношение
        if self.stats.get('total', 0) > 0:
            self.add_percentage_section(layout)

    def add_percentage_section(self, layout):
        """Добавление секции с процентным соотношением"""
        total = self.stats.get('total', 0)
        applying = self.stats.get('applying', 0)
        refused = self.stats.get('refused', 0)

        if total > 0:
            percent_widget = QWidget()
            percent_layout = QHBoxLayout(percent_widget)
            percent_layout.setSpacing(30)
            percent_layout.setContentsMargins(0, 15, 0, 0)

            # Поступают
            if applying > 0:
                percent_applying = (applying / total) * 100
                applying_widget = QWidget()
                applying_layout = QVBoxLayout(applying_widget)
                applying_layout.setContentsMargins(0, 0, 0, 0)

                applying_label = QLabel("✅ Поступают")
                applying_label.setStyleSheet("color: #bdc3c7; font-weight: bold; font-size: 14px;")
                applying_layout.addWidget(applying_label)

                applying_percent = QLabel(f"{percent_applying:.1f}%")
                applying_percent.setStyleSheet("color: #2ecc71; font-size: 24px; font-weight: bold;")
                applying_layout.addWidget(applying_percent)

                applying_count = QLabel(f"({applying} чел.)")
                applying_count.setStyleSheet("color: #95a5a6; font-size: 12px;")
                applying_layout.addWidget(applying_count)

                percent_layout.addWidget(applying_widget)

            # Отказались
            if refused > 0:
                percent_refused = (refused / total) * 100
                refused_widget = QWidget()
                refused_layout = QVBoxLayout(refused_widget)
                refused_layout.setContentsMargins(0, 0, 0, 0)

                refused_label = QLabel("❌ Отказались")
                refused_label.setStyleSheet("color: #bdc3c7; font-weight: bold; font-size: 14px;")
                refused_layout.addWidget(refused_label)

                refused_percent = QLabel(f"{percent_refused:.1f}%")
                refused_percent.setStyleSheet("color: #e74c3c; font-size: 24px; font-weight: bold;")
                refused_layout.addWidget(refused_percent)

                refused_count = QLabel(f"({refused} чел.)")
                refused_count.setStyleSheet("color: #95a5a6; font-size: 12px;")
                refused_layout.addWidget(refused_count)

                percent_layout.addWidget(refused_widget)

            percent_layout.addStretch()
            layout.addWidget(percent_widget)


class EmptyStateWidget(QFrame):
    """Виджет для состояния без данных"""

    def __init__(self, message, parent=None):
        super().__init__(parent)
        self.message = message
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.NoFrame)
        self.setStyleSheet("background-color: transparent;")

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        icon_label = QLabel("📊")
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_font = QFont()
        icon_font.setPointSize(48)
        icon_label.setFont(icon_font)
        icon_label.setStyleSheet("color: #bdc3c7; margin-bottom: 20px;")

        message_label = QLabel(self.message)
        message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        message_label.setWordWrap(True)
        message_label.setStyleSheet("""
            QLabel {
                color: #7f8c8d;
                font-size: 16px;
                font-weight: medium;
            }
        """)

        sub_label = QLabel("Добавьте данные через вкладку '📋 Данные абитуриентов'")
        sub_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub_label.setStyleSheet("color: #95a5a6; font-size: 13px; margin-top: 10px;")

        layout.addStretch()
        layout.addWidget(icon_label)
        layout.addWidget(message_label)
        layout.addWidget(sub_label)
        layout.addStretch()

        self.setMinimumHeight(300)



class StatisticsCard(QFrame):
    """Карточка статистики"""

    def __init__(self, title, values, colors, parent=None):
        super().__init__(parent)
        self.title = title
        self.values = values  # {'М': 10, 'Ж': 5, 'в/сл': 3}
        self.colors = colors  # {'М': '#3498db', 'Ж': '#e67e22', 'в/сл': '#2ecc71'}
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        self.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 12px;
                border: 1px solid #e0e0e0;
            }
            QFrame:hover {
                background-color: #f8f9fa;
                border-color: #3498db;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Заголовок
        title_label = QLabel(self.title)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 5px;")
        layout.addWidget(title_label)

        # Разделитель
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        line.setStyleSheet("background-color: #ecf0f1;")
        layout.addWidget(line)

        # Значения
        for key, value in self.values.items():
            value_widget = QWidget()
            value_layout = QHBoxLayout(value_widget)
            value_layout.setContentsMargins(5, 2, 5, 2)

            # Цветной индикатор
            indicator = QLabel("●")
            indicator.setStyleSheet(f"color: {self.colors.get(key, '#95a5a6')}; font-size: 14px;")

            # Название
            name_label = QLabel(key)
            name_label.setStyleSheet("color: #7f8c8d; font-size: 13px;")

            # Значение
            value_label = QLabel(str(value))
            value_label.setAlignment(Qt.AlignmentFlag.AlignRight)
            value_label.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")

            value_layout.addWidget(indicator)
            value_layout.addWidget(name_label)
            value_layout.addStretch()
            value_layout.addWidget(value_label)

            layout.addWidget(value_widget)

        # Итого
        total = sum(self.values.values())
        total_widget = QWidget()
        total_layout = QHBoxLayout(total_widget)
        total_layout.setContentsMargins(5, 8, 5, 5)

        total_label = QLabel("ИТОГО:")
        total_label.setStyleSheet("color: #2c3e50; font-weight: bold; font-size: 13px;")

        total_value = QLabel(str(total))
        total_value.setAlignment(Qt.AlignmentFlag.AlignRight)
        total_value.setStyleSheet("color: #3498db; font-size: 16px; font-weight: bold;")

        total_layout.addWidget(total_label)
        total_layout.addStretch()
        total_layout.addWidget(total_value)

        layout.addWidget(total_widget)

        self.setMinimumWidth(180)


class PlanDialog(QDialog):
    """Диалог для редактирования плана"""

    def __init__(self, department_id, department_name, current_plan, year, db, parent=None):
        super().__init__(parent)
        self.department_id = department_id
        self.department_name = department_name
        self.current_plan = current_plan
        self.year = year
        self.db = db
        self.setModal(True)
        self.setWindowTitle(f"Редактирование плана - {department_name}")
        self.setFixedSize(400, 300)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Заголовок
        title = QLabel(f"📊 План набора на {self.year} год")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # Форма
        form_layout = QFormLayout()
        form_layout.setSpacing(15)

        # План по мужчинам
        self.plan_m = QLabel()
        self.plan_m_input = QLabel()
        # Создаем спинбокс
        from PyQt5.QtWidgets import QSpinBox
        self.plan_m_spin = QSpinBox()
        self.plan_m_spin.setRange(0, 1000)
        self.plan_m_spin.setValue(self.current_plan.get('plan_m', 0))
        form_layout.addRow("План по мужчинам (М):", self.plan_m_spin)

        # План по женщинам
        self.plan_f_spin = QSpinBox()
        self.plan_f_spin.setRange(0, 1000)
        self.plan_f_spin.setValue(self.current_plan.get('plan_f', 0))
        form_layout.addRow("План по женщинам (Ж):", self.plan_f_spin)

        # План по военнослужащим
        self.plan_military_spin = QSpinBox()
        self.plan_military_spin.setRange(0, 1000)
        self.plan_military_spin.setValue(self.current_plan.get('plan_military', 0))
        form_layout.addRow("План по военнослужащим (в/сл):", self.plan_military_spin)

        layout.addLayout(form_layout)

        # Кнопки
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.save_plan)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)

    def save_plan(self):
        """Сохранение плана"""
        plan_data = {
            'plan_m': self.plan_m_spin.value(),
            'plan_f': self.plan_f_spin.value(),
            'plan_military': self.plan_military_spin.value()
        }

        success = self.db.set_plan(
            self.department_id,
            self.year,
            plan_data['plan_m'],
            plan_data['plan_f'],
            plan_data['plan_military']
        )

        if success:
            QMessageBox.information(self, "Успех", "План успешно сохранен!")
            self.accept()
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось сохранить план!")


class StatisticsWidget(QWidget):
    """Главный виджет статистики"""

    def __init__(self, user_id, role, db):
        super().__init__()
        self.user_id = user_id
        self.role = role
        self.db = db
        self.current_year = 2026  # Можно сделать выбор года
        self.init_ui()

    def init_ui(self):
        # Основной layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 15, 20, 15)

        # Заголовок
        title_container = QWidget()
        title_layout = QHBoxLayout(title_container)
        title_layout.setContentsMargins(0, 0, 0, 0)

        title_label = QLabel("📊 Статистика по подразделениям")
        title_font = QFont()
        title_font.setPointSize(22)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50;")

        title_layout.addWidget(title_label)
        title_layout.addStretch()

        # Выбор года
        year_label = QLabel("Год:")
        self.year_combo = QComboBox()
        self.year_combo.addItems(["2024", "2025", "2026", "2027", "2028"])
        self.year_combo.setCurrentText(str(self.current_year))
        self.year_combo.currentTextChanged.connect(self.on_year_changed)

        title_layout.addWidget(year_label)
        title_layout.addWidget(self.year_combo)

        main_layout.addWidget(title_container)

        # Панель управления
        controls_widget = QFrame()
        controls_widget.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 8px;
                border: 1px solid #e0e0e0;
                padding: 15px;
            }
        """)

        controls_layout = QHBoxLayout(controls_widget)
        controls_layout.setSpacing(15)

        # Подразделение
        dept_label = QLabel("Подразделение:")
        dept_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.department_combo = QComboBox()
        self.department_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background-color: white;
                min-width: 200px;
            }
        """)
        self.load_departments()

        # Кнопка редактирования плана (только для админа или начальника)
        self.edit_plan_btn = QPushButton("✏️ Редактировать план")
        self.edit_plan_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
        """)
        self.edit_plan_btn.clicked.connect(self.edit_plan)

        # Кнопка обновления
        self.refresh_btn = QPushButton("🔄 Обновить")
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.refresh_btn.clicked.connect(self.update_statistics)

        controls_layout.addWidget(dept_label)
        controls_layout.addWidget(self.department_combo)
        controls_layout.addStretch()
        controls_layout.addWidget(self.edit_plan_btn)
        controls_layout.addWidget(self.refresh_btn)

        main_layout.addWidget(controls_widget)

        # Область с прокруткой для статистики
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
        """)

        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setSpacing(20)
        self.scroll_layout.setContentsMargins(5, 5, 5, 5)

        self.scroll_area.setWidget(self.scroll_content)
        main_layout.addWidget(self.scroll_area)

        # Информационная панель
        self.info_panel = QFrame()
        self.info_panel.setStyleSheet("""
            QFrame {
                background-color: #e8f4fc;
                border-radius: 8px;
                border: 1px solid #b3e0ff;
                padding: 12px;
            }
        """)
        self.info_panel.setVisible(False)

        info_layout = QHBoxLayout(self.info_panel)
        self.info_label = QLabel("")
        self.info_label.setStyleSheet("color: #2c3e50; font-size: 13px;")
        info_layout.addWidget(self.info_label)

        main_layout.addWidget(self.info_panel)

        # Инициализация данных
        self.update_statistics()

    def load_departments(self):
        """Загрузка подразделений"""
        self.department_combo.clear()

        if self.role == 'admin':
            # Админ видит все подразделения
            departments = self.db.get_departments()
            self.department_combo.addItem("Все подразделения")
            for dept in departments:
                self.department_combo.addItem(dept['name'])
        else:
            # Обычный пользователь видит только свое подразделение
            user_info = self.db.get_user_by_id(self.user_id)
            if user_info and user_info['department_id']:
                cursor = self.db.conn.cursor()
                cursor.execute('SELECT name FROM departments WHERE id = ?', (user_info['department_id'],))
                dept = cursor.fetchone()
                if dept:
                    self.department_combo.addItem(dept['name'])
            else:
                self.department_combo.addItem("Нет подразделения")
            self.department_combo.setEnabled(False)
            self.edit_plan_btn.setVisible(True)  # Начальник может редактировать план

    def on_year_changed(self, year):
        """Изменение года"""
        self.current_year = int(year)
        self.update_statistics()

    def edit_plan(self):
        """Редактирование плана"""
        department_name = self.department_combo.currentText()
        if department_name == "Все подразделения":
            QMessageBox.warning(self, "Внимание", "Выберите конкретное подразделение для редактирования плана!")
            return

        # Получаем ID подразделения
        cursor = self.db.conn.cursor()
        cursor.execute('SELECT id FROM departments WHERE name = ?', (department_name,))
        result = cursor.fetchone()

        if not result:
            QMessageBox.warning(self, "Ошибка", "Подразделение не найдено!")
            return

        department_id = result['id']

        # Получаем текущий план
        current_plan = self.db.get_plan(department_id, self.current_year)

        # Открываем диалог редактирования
        dialog = PlanDialog(department_id, department_name, current_plan, self.current_year, self.db, self)
        if dialog.exec():
            self.update_statistics()

    def update_statistics(self):
        """Обновление статистики"""
        # Очистка предыдущих данных
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        department_name = self.department_combo.currentText()

        if department_name == "Все подразделения":
            self.display_all_departments_stats()
        else:
            self.display_department_stats(department_name)

        # Добавляем растягивающийся элемент
        self.scroll_layout.addStretch()

    def display_all_departments_stats(self):
        """Отображение статистики по всем подразделениям"""
        departments = self.db.get_departments()

        for dept in departments:
            self.display_department_stats(dept['name'])

        if not departments:
            empty_widget = self.create_empty_widget("Нет данных по подразделениям")
            self.scroll_layout.addWidget(empty_widget)

    def display_department_stats(self, department_name):
        """Отображение статистики по одному подразделению"""
        # Получаем статистику из БД
        stats = self.db.get_statistics_by_department(department_name)

        # Получаем план
        cursor = self.db.conn.cursor()
        cursor.execute('SELECT id FROM departments WHERE name = ?', (department_name,))
        result = cursor.fetchone()

        plan = {'plan_m': 0, 'plan_f': 0, 'plan_military': 0}
        if result:
            plan = self.db.get_plan(result['id'], self.current_year)

        # Создаем карточки
        cards_widget = QWidget()
        cards_layout = QGridLayout(cards_widget)
        cards_layout.setSpacing(15)
        cards_layout.setContentsMargins(0, 0, 0, 0)

        # Цвета для категорий
        colors = {
            'М': '#3498db',
            'Ж': '#e67e22',
            'в/сл': '#2ecc71'
        }

        # 1. Блок "План"
        plan_values = {
            'М': plan['plan_m'],
            'Ж': plan['plan_f'],
            'в/сл': plan['plan_military']
        }
        plan_card = StatisticsCard("📋 ПЛАН", plan_values, colors)
        cards_layout.addWidget(plan_card, 0, 0)

        # 2. Блок "Поступают"
        applying_values = {
            'М': stats['applying_vk'] + stats['applying_ok'] + stats['applying_vavko'],
            'Ж': 0,  # Нужно добавить в БД разделение по полу для поступающих
            'в/сл': 0
        }
        # TODO: Добавить в БД разделение по полу для статусов документов
        # Пока используем общее количество
        total_applying = stats['applying_vk'] + stats['applying_ok'] + stats['applying_vavko']
        applying_values = {
            'М': total_applying,
            'Ж': 0,
            'в/сл': 0
        }
        applying_card = StatisticsCard("✅ ПОСТУПАЮТ", applying_values, colors)
        cards_layout.addWidget(applying_card, 0, 1)

        # 3. Блок "Дело в ВК"
        vk_values = {
            'М': stats['applying_vk'],
            'Ж': 0,
            'в/сл': 0
        }
        vk_card = StatisticsCard("📄 ДЕЛО В ВК", vk_values, colors)
        cards_layout.addWidget(vk_card, 0, 2)

        # 4. Блок "Дело в ОК"
        ok_values = {
            'М': stats['applying_ok'],
            'Ж': 0,
            'в/сл': 0
        }
        ok_card = StatisticsCard("📋 ДЕЛО В ОК", ok_values, colors)
        cards_layout.addWidget(ok_card, 0, 3)

        # Добавляем название подразделения
        dept_header = QLabel(f"🏢 {department_name}")
        dept_header.setStyleSheet("""
            QLabel {
                font-size: 16px;
                font-weight: bold;
                color: #2c3e50;
                padding: 10px;
                background-color: #ecf0f1;
                border-radius: 8px;
            }
        """)

        # Собираем все в контейнер
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setSpacing(10)
        container_layout.addWidget(dept_header)
        container_layout.addWidget(cards_widget)

        self.scroll_layout.addWidget(container)

    def create_empty_widget(self, message):
        """Создание виджета для пустого состояния"""
        from PyQt5.QtWidgets import QFrame, QVBoxLayout, QLabel
        widget = QFrame()
        widget.setFrameStyle(QFrame.Shape.NoFrame)

        layout = QVBoxLayout(widget)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        icon_label = QLabel("📊")
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_font = QFont()
        icon_font.setPointSize(48)
        icon_label.setFont(icon_font)
        icon_label.setStyleSheet("color: #bdc3c7; margin-bottom: 20px;")

        message_label = QLabel(message)
        message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        message_label.setStyleSheet("color: #7f8c8d; font-size: 16px;")

        layout.addStretch()
        layout.addWidget(icon_label)
        layout.addWidget(message_label)
        layout.addStretch()

        widget.setMinimumHeight(300)
        return widget