# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QComboBox, QPushButton, QScrollArea, QFrame,
                             QGridLayout, QSizePolicy, QToolButton, QTableWidget, QTableWidgetItem, QFileDialog,
                             QProgressBar)
from PyQt5.QtCore import Qt, pyqtSignal, QPropertyAnimation, QEasingCurve
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
        # ИСПРАВЛЕНО: Применяем стиль строго к CourseSection, чтобы дочерние элементы не наследовались
        self.setStyleSheet("""
            CourseSection {
                background-color: white;
                border-radius: 10px;
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

        icon_label = QLabel()
        icon_label.setStyleSheet("font-size: 24px; background: transparent; border: none;")
        course_layout.addWidget(icon_label)

        course_label = QLabel(self.course_name)
        course_font = QFont()
        course_font.setPointSize(16)
        course_font.setBold(True)
        course_label.setFont(course_font)
        course_label.setStyleSheet("color: #2c3e50; background: transparent; border: none;")
        course_layout.addWidget(course_label)

        # Если есть факультет, добавляем его
        if self.stats.get('faculty'):
            faculty_label = QLabel(f"({self.stats.get('faculty')})")
            faculty_label.setStyleSheet("color: #7f8c8d; font-size: 14px; background: transparent; border: none;")
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
                border: none;
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
            ('Всего', 'total', '#3498db'),
            ('Отобраны', 'applying', '#2ecc71'),
            ('Отказались', 'refused', '#e74c3c'),
            ('Мужчины', 'male', '#9b59b6'),
            ('Женщины', 'female', '#e67e22'),
            ('Военнослужащие', 'military', '#1abc9c'),
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
            docs_header = QLabel("Статус документов:")
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
                ('Формируется', 'doc1', '#f39c12'),
                ('Отправлено', 'doc2', '#8e44ad'),
                ('В ВА ВКО', 'doc3', '#16a085'),
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
                # ИСПРАВЛЕНО: Убираем влияние на внутренние QLabel (добавили ID или имя объекта)
                chart_container.setObjectName("ChartContainer")
                chart_container.setStyleSheet("""
                    QFrame#ChartContainer {
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
                        border: none;
                        background: transparent;
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

                        color_indicator = QLabel()
                        color_indicator.setStyleSheet(f"color: {color}; font-size: 10px; border: none; background: transparent;")

                        label_text = QLabel(f"{label}: {value}")
                        label_text.setStyleSheet("color: #6c757d; font-size: 12px; border: none; background: transparent;")

                        item_layout.addWidget(color_indicator)
                        item_layout.addWidget(label_text)
                        item_layout.addStretch()

                        text_layout.addWidget(item_widget)

                chart_inner.addWidget(text_widget)
                chart_layout.addWidget(chart_container)

        if chart_layout.count() > 0:
            chart_layout.addStretch()
            layout.addWidget(chart_widget)


class StatisticsCard(QFrame):
    """Карточка статистики - упрощенная версия"""

    def __init__(self, title, values, colors, parent=None):
        super().__init__(parent)
        self.title = title
        if isinstance(values, dict):
            self.is_dict = True
            self.values_dict = values
            self.value = None
        else:
            self.is_dict = False
            self.values_dict = None
            self.value = values
        self.colors = colors
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        # ИСПРАВЛЕНО: Применяем стиль конкретно к классу StatisticsCard
        self.setStyleSheet("""
            StatisticsCard {
                background-color: white;
                border-radius: 10px;
            }
            StatisticsCard:hover {
                background-color: #f8f9fa;
                border-color: #3498db;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # Заголовок
        title_label = QLabel(self.title)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50;")
        layout.addWidget(title_label)

        # Разделитель
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background-color: #ecf0f1; max-height: 1px;")
        layout.addWidget(line)

        if self.is_dict:
            # Отображаем словарь значений
            for key, value in self.values_dict.items():
                value_widget = QWidget()
                value_layout = QHBoxLayout(value_widget)
                value_layout.setContentsMargins(5, 2, 5, 2)

                indicator = QLabel()
                indicator.setStyleSheet(f"color: {self.colors.get(key, '#95a5a6')}; font-size: 12px;")

                name_label = QLabel(key)
                name_label.setStyleSheet("color: #7f8c8d; font-size: 11px;")

                value_label = QLabel(str(value))
                value_label.setAlignment(Qt.AlignmentFlag.AlignRight)
                value_label.setStyleSheet("color: #2c3e50; font-size: 14px; font-weight: bold;")

                value_layout.addWidget(indicator)
                value_layout.addWidget(name_label)
                value_layout.addStretch()
                value_layout.addWidget(value_label)

                layout.addWidget(value_widget)

            # Итого
            total = sum(self.values_dict.values())
            total_widget = QWidget()
            total_layout = QHBoxLayout(total_widget)
            total_layout.setContentsMargins(5, 5, 5, 0)

            total_label = QLabel("ИТОГО:")
            total_label.setStyleSheet("color: #2c3e50; font-weight: bold; font-size: 10px;")

            total_value = QLabel(str(total))
            total_value.setAlignment(Qt.AlignmentFlag.AlignRight)
            total_value.setStyleSheet("color: #3498db; font-size: 14px; font-weight: bold;")

            total_layout.addWidget(total_label)
            total_layout.addStretch()
            total_layout.addWidget(total_value)

            layout.addWidget(total_widget)
        else:
            # Простое значение
            value_label = QLabel(str(self.value))
            value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            value_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #3498db;")
            layout.addWidget(value_label)

        self.setMinimumWidth(140)


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

        title_label = QLabel("Общая статистика (все курсы)")
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
            ('Всего абитуриентов', 'total', '#3498db'),
            ('Отобраны', 'applying', '#2ecc71'),
            ('Отказались', 'refused', '#e74c3c'),
            ('Мужчины', 'male', '#9b59b6'),
            ('Женщины', 'female', '#e67e22'),
            ('Военнослужащие', 'military', '#1abc9c'),
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

                applying_label = QLabel("Поступают")
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

                refused_label = QLabel("Отказались")
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

        icon_label = QLabel("")
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

        sub_label = QLabel("Добавьте данные через вкладку 'Данные абитуриентов'")
        sub_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub_label.setStyleSheet("color: #95a5a6; font-size: 13px; margin-top: 10px;")

        layout.addStretch()
        layout.addWidget(icon_label)
        layout.addWidget(message_label)
        layout.addWidget(sub_label)
        layout.addStretch()

        self.setMinimumHeight(300)


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
        title = QLabel(f"План набора на {self.year} год")
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


class ExpandableDepartmentCard(QFrame):
    """Раскрывающаяся карточка подразделения"""

    def __init__(self, department_name, stats, plan, parent=None):
        super().__init__(parent)
        self.department_name = department_name
        self.stats = stats
        self.plan = plan
        self.is_expanded = False
        self.animation_duration = 300
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        self.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 12px;
                margin: 5px;
            }
            QFrame:hover {
                border-color: #3498db;
            }
        """)

        # Основной layout
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # Заголовок карточки (всегда виден)
        self.header_widget = QWidget()
        self.header_widget.setCursor(Qt.CursorShape.PointingHandCursor)
        self.header_widget.setStyleSheet("""
            QWidget {
                background-color: transparent;
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
            }
            QWidget:hover {
                background-color: #f8f9fa;
            }
        """)

        header_layout = QHBoxLayout(self.header_widget)
        header_layout.setContentsMargins(20, 15, 20, 15)
        header_layout.setSpacing(15)

        # Кнопка раскрытия
        self.expand_btn = QToolButton()
        self.expand_btn.setArrowType(Qt.ArrowType.RightArrow)
        self.expand_btn.setStyleSheet("""
            QToolButton {
                border: none;
                background-color: #3498db;
                border-radius: 4px;
                color: white;
                font-weight: bold;
                padding: 4px;
            }
            QToolButton:hover {
                background-color: #2980b9;
            }
        """)
        self.expand_btn.setFixedSize(24, 24)
        self.expand_btn.clicked.connect(self.toggle_expand)
        header_layout.addWidget(self.expand_btn)

        # Иконка подразделения
        icon_label = QLabel()
        icon_label.setStyleSheet("font-size: 28px;")
        header_layout.addWidget(icon_label)

        # Название подразделения
        name_label = QLabel(self.department_name)
        name_font = QFont()
        name_font.setPointSize(14)
        name_font.setBold(True)
        name_label.setFont(name_font)
        name_label.setStyleSheet("color: #2c3e50;")
        header_layout.addWidget(name_label)

        header_layout.addStretch()

        # Краткая статистика (всегда видна)
        quick_stats_widget = QWidget()
        quick_stats_layout = QHBoxLayout(quick_stats_widget)
        quick_stats_layout.setSpacing(20)
        quick_stats_layout.setContentsMargins(0, 0, 0, 0)

        # План
        total_plan = self.plan.get('plan_m', 0) + self.plan.get('plan_f', 0) + self.plan.get('plan_military', 0)
        if total_plan > 0:
            quick_stats_layout.addWidget(self._create_stat_badge("План", total_plan, "#f39c12"))
        # Процент выполнения = (Дела в ОК / План) * 100%
        ok_total = self.stats.get('ok_m', 0) + self.stats.get('ok_f', 0) + self.stats.get('ok_mil', 0)
        if total_plan > 0:
            percent = int((ok_total / total_plan) * 100)
            quick_stats_layout.addWidget(self._create_percent_badge(percent))

        header_layout.addWidget(quick_stats_widget)

        self.main_layout.addWidget(self.header_widget)

        # Контентная часть (скрыта по умолчанию)
        self.content_widget = QWidget()
        self.content_widget.setVisible(False)
        self.content_widget.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                border-bottom-left-radius: 12px;
                border-bottom-right-radius: 12px;
            }
        """)

        content_layout = QVBoxLayout(self.content_widget)
        content_layout.setContentsMargins(20, 15, 20, 20)
        content_layout.setSpacing(15)

        # Добавляем детальную статистику
        self._add_detailed_stats(content_layout)

        self.main_layout.addWidget(self.content_widget)

    def _create_stat_badge(self, label, value, color):
        """Создание бейджа со статистикой"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)

        label_widget = QLabel(label)
        label_widget.setStyleSheet(f"""
            QLabel {{
                color: #7f8c8d;
                font-size: 11px;
                font-weight: normal;
            }}
        """)

        value_widget = QLabel(str(value))
        value_widget.setStyleSheet(f"""
            QLabel {{
                color: {color};
                font-size: 16px;
                font-weight: bold;
            }}
        """)

        layout.addWidget(label_widget)
        layout.addWidget(value_widget)

        return widget

    def _create_percent_badge(self, percent):
        """Создание бейджа с процентом выполнения"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        # Цвет в зависимости от процента
        if percent >= 80:
            color = "#2ecc71"
        elif percent >= 50:
            color = "#f39c12"
        else:
            color = "#e74c3c"

        value_widget = QLabel(f"{percent}%")
        value_widget.setStyleSheet(f"""
            QLabel {{
                background-color: {color};
                color: white;
                font-size: 12px;
                font-weight: bold;
                padding: 4px 8px;
                border-radius: 12px;
            }}
        """)

        layout.addWidget(value_widget)
        return widget

    def _add_detailed_stats(self, layout):
        """Добавление детальной статистики"""

        # Блок с карточками
        cards_widget = QWidget()
        cards_layout = QGridLayout(cards_widget)
        cards_layout.setSpacing(15)
        cards_layout.setContentsMargins(0, 0, 0, 0)

        colors = {
            'М': '#3498db',
            'Ж': '#e67e22',
            'в/сл': '#2ecc71'
        }

        # План
        plan_values = {
            'М': self.plan.get('plan_m', 0),
            'Ж': self.plan.get('plan_f', 0),
            'в/сл': self.plan.get('plan_military', 0)
        }
        plan_card = StatisticsCard("ПЛАН", plan_values, colors)
        cards_layout.addWidget(plan_card, 0, 0)

        # Отобраны
        applying_values = {
            'М': self.stats.get('applying_m', 0),
            'Ж': self.stats.get('applying_f', 0),
            'в/сл': self.stats.get('applying_mil', 0)
        }
        applying_card = StatisticsCard("ОТОБРАНЫ", applying_values, colors)
        cards_layout.addWidget(applying_card, 0, 1)

        # Дело в ВК
        vk_values = {
            'М': self.stats.get('vk_m', 0),
            'Ж': self.stats.get('vk_f', 0),
            'в/сл': self.stats.get('vk_mil', 0)
        }
        vk_card = StatisticsCard("ДЕЛО В ВК", vk_values, colors)
        cards_layout.addWidget(vk_card, 0, 2)

        # Дело в ОК
        ok_values = {
            'М': self.stats.get('ok_m', 0),
            'Ж': self.stats.get('ok_f', 0),
            'в/сл': self.stats.get('ok_mil', 0)
        }
        ok_card = StatisticsCard("ДЕЛО В ОК", ok_values, colors)
        cards_layout.addWidget(ok_card, 0, 3)

        layout.addWidget(cards_widget)

        # Кнопки управления
        buttons_widget = QWidget()
        buttons_layout = QHBoxLayout(buttons_widget)
        buttons_layout.setContentsMargins(0, 10, 0, 0)
        buttons_layout.setSpacing(10)

        # Кнопка "Редактировать план"
        edit_plan_btn = QPushButton("Редактировать план")
        edit_plan_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
        """)
        edit_plan_btn.clicked.connect(self.edit_plan)
        buttons_layout.addWidget(edit_plan_btn)

        # Кнопка "Статистика по регионам"
        region_btn = QPushButton("Статистика по регионам")
        region_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        region_btn.clicked.connect(self.show_region_stats)
        buttons_layout.addWidget(region_btn)

        buttons_layout.addStretch()
        layout.addWidget(buttons_widget)

    def edit_plan(self):
        """Редактирование плана для этого подразделения"""
        # Находим родительский StatisticsWidget
        parent_widget = self.parent()
        while parent_widget and not isinstance(parent_widget, StatisticsWidget):
            parent_widget = parent_widget.parent()

        if not parent_widget:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить родительский виджет!")
            return

        # Проверяем права
        if not hasattr(parent_widget, 'role'):
            QMessageBox.warning(self, "Ошибка", "Не удалось определить права доступа!")
            return

        # Проверяем, есть ли у пользователя права на редактирование плана
        can_edit = False

        # Админ может редактировать всё
        if parent_widget.role == 'admin':
            can_edit = True
        else:
            # Проверяем, является ли пользователь начальником этого подразделения
            user_info = parent_widget.db.get_user_by_id(parent_widget.user_id)
            user_dict = dict(user_info) if user_info else {}

            if user_dict.get('is_head'):
                # Проверяем, что это его подразделение
                cursor = parent_widget.db.conn.cursor()
                cursor.execute('SELECT name FROM departments WHERE id = ?', (user_dict.get('department_id'),))
                dept = cursor.fetchone()
                if dept and dept['name'] == self.department_name:
                    can_edit = True

        if not can_edit:
            QMessageBox.warning(self, "Внимание", "У вас нет прав на редактирование плана этого подразделения!")
            return

        # Получаем ID подразделения
        cursor = parent_widget.db.conn.cursor()
        cursor.execute('SELECT id FROM departments WHERE name = ?', (self.department_name,))
        result = cursor.fetchone()

        if not result:
            QMessageBox.warning(self, "Ошибка", "Подразделение не найдено!")
            return

        department_id = result['id']
        current_year = parent_widget.current_year if hasattr(parent_widget, 'current_year') else 2026

        # Получаем текущий план
        current_plan = parent_widget.db.get_plan(department_id, current_year)

        # Открываем диалог редактирования
        dialog = PlanDialog(department_id, self.department_name, current_plan, current_year, parent_widget.db, self)
        if dialog.exec():
            # Обновляем статистику
            parent_widget.update_statistics()
            QMessageBox.information(self, "Успех", "План успешно обновлен!")

    def show_region_stats(self):
        """Показать статистику по регионам для этого подразделения"""
        # Получаем ID подразделения
        cursor = self.db.conn.cursor()
        cursor.execute('SELECT id FROM departments WHERE name = ?', (self.department_name,))
        result = cursor.fetchone()

        if result:
            # Получаем роль родительского виджета
            parent_widget = self.window()
            if hasattr(parent_widget, 'role'):
                role = parent_widget.role
            else:
                role = 'admin'
            dialog = RegionStatsDialog(self.department_name, result['id'], self.db, role, self.window())
            dialog.exec()

    def toggle_expand(self):
        """Переключение раскрытия карточки"""
        self.is_expanded = not self.is_expanded

        if self.is_expanded:
            self.expand_btn.setArrowType(Qt.ArrowType.DownArrow)
            self.content_widget.setVisible(True)
            # Анимация появления
            self.content_widget.setMaximumHeight(0)
            self.animation = QPropertyAnimation(self.content_widget, b"maximumHeight")
            self.animation.setDuration(300)
            self.animation.setStartValue(0)
            self.animation.setEndValue(self.content_widget.sizeHint().height())
            self.animation.setEasingCurve(QEasingCurve.Type.OutCubic)
            self.animation.start()
        else:
            self.expand_btn.setArrowType(Qt.ArrowType.RightArrow)
            # Анимация скрытия
            self.animation = QPropertyAnimation(self.content_widget, b"maximumHeight")
            self.animation.setDuration(300)
            self.animation.setStartValue(self.content_widget.height())
            self.animation.setEndValue(0)
            self.animation.setEasingCurve(QEasingCurve.Type.InCubic)
            self.animation.finished.connect(lambda: self.content_widget.setVisible(False))
            self.animation.start()

    # Этот метод нужно добавить, чтобы передать db в карточку
    def set_db(self, db):
        self.db = db


class StatisticsWidget(QWidget):
    """Главный виджет статистики"""

    def __init__(self, user_id, role, db):
        super().__init__()
        self.user_id = user_id
        self.role = role
        self.db = db
        self.current_year = 2026
        self.cards = []  # Список карточек
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

        title_label = QLabel("Статистика по подразделениям")
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
        self.year_combo.setStyleSheet("""
            QComboBox {
                padding: 6px 12px;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background-color: white;
                min-width: 100px;
            }
        """)

        title_layout.addWidget(year_label)
        title_layout.addWidget(self.year_combo)

        # Кнопка обновления
        self.refresh_btn = QPushButton("Обновить")
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
        title_layout.addWidget(self.refresh_btn)

        main_layout.addWidget(title_container)

        # Краткая сводка (общая статистика)
        self.summary_widget = QFrame()
        self.summary_widget.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #2c3e50, stop:1 #34495e);
                border-radius: 12px;
                padding: 15px;
            }
        """)
        summary_layout = QHBoxLayout(self.summary_widget)
        summary_layout.setSpacing(30)
        main_layout.addWidget(self.summary_widget)

        # Область с прокруткой для карточек
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
        self.scroll_layout.setSpacing(10)
        self.scroll_layout.setContentsMargins(5, 5, 5, 5)

        self.scroll_area.setWidget(self.scroll_content)
        main_layout.addWidget(self.scroll_area)

        # Инициализация данных
        self.update_statistics()

    def load_departments(self):
        """Загрузка подразделений"""
        self.department_combo.clear()

        if self.role == 'admin':
            departments = self.db.get_departments()
            self.department_combo.addItem("Все подразделения")
            for dept in departments:
                self.department_combo.addItem(dept['name'])
            self.edit_plan_btn.setVisible(True)  # Админ может редактировать
        else:
            user_info = self.db.get_user_by_id(self.user_id)
            user_dict = dict(user_info) if user_info else {}

            if user_dict.get('is_head') and user_dict.get('department_id'):
                # Начальник может редактировать план своего подразделения
                cursor = self.db.conn.cursor()
                cursor.execute('SELECT name FROM departments WHERE id = ?', (user_dict['department_id'],))
                dept = cursor.fetchone()
                if dept:
                    self.department_combo.addItem(dept['name'])
                    self.edit_plan_btn.setVisible(True)  # Начальник может редактировать
                else:
                    self.department_combo.addItem("Нет подразделения")
                    self.edit_plan_btn.setVisible(False)
            else:
                # Обычный пользователь
                self.department_combo.addItem("Только мои записи")
                self.department_combo.setEnabled(False)
                self.edit_plan_btn.setVisible(False)

    def update_summary(self, departments_stats, visible_department_ids):
        """Обновление общей сводки"""
        # Очищаем старую сводку
        while self.summary_widget.layout().count():
            item = self.summary_widget.layout().takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        summary_layout = self.summary_widget.layout()

        # Считаем план ТОЛЬКО для видимых подразделений
        total_plans = 0
        for dept_id in visible_department_ids:
            plan = self.db.get_plan(dept_id, self.current_year)
            total_plans += plan.get('plan_m', 0) + plan.get('plan_f', 0) + plan.get('plan_military', 0)

        # Отобрано (сумма по всем категориям)
        total_applying = sum(
            s.get('applying_m', 0) + s.get('applying_f', 0) + s.get('applying_mil', 0) for s in departments_stats)

        # Дела в ВК (сумма по всем категориям)
        total_vk = sum(
            s.get('vk_m', 0) + s.get('vk_f', 0) + s.get('vk_mil', 0) for s in departments_stats)

        # Дела в ОК (сумма по всем категориям)
        total_ok = sum(
            s.get('ok_m', 0) + s.get('ok_f', 0) + s.get('ok_mil', 0) for s in departments_stats)

        # Создаем виджеты сводки
        summary_items = [
            ("План", total_plans, "#f39c12"),
            ("Отобрано", total_applying, "#2ecc71"),
            ("Дела в ВК", total_vk, "#f39c12"),
            ("Дела в ОК", total_ok, "#8e44ad"),
        ]

        # Выполнение плана
        if total_plans > 0:
            percent = int((total_ok / total_plans) * 100) if total_plans > 0 else 0
            percent_color = "#2ecc71" if percent >= 80 else "#f39c12" if percent >= 50 else "#e74c3c"
            summary_items.append(("Выполнение плана", f"{percent}%", percent_color))

        for label, value, color in summary_items:
            item_widget = QWidget()
            item_layout = QVBoxLayout(item_widget)
            item_layout.setContentsMargins(10, 5, 10, 5)

            label_widget = QLabel(label)
            label_widget.setStyleSheet("color: #bdc3c7; font-size: 12px;")

            value_widget = QLabel(str(value))
            value_widget.setStyleSheet(f"""
                QLabel {{
                    color: {color};
                    font-size: 24px;
                    font-weight: bold;
                }}
            """)

            item_layout.addWidget(label_widget)
            item_layout.addWidget(value_widget)
            summary_layout.addWidget(item_widget)

        summary_layout.addStretch()

    def update_statistics(self):
        """Обновление статистики"""
        # Очистка предыдущих карточек
        for card in self.cards:
            card.deleteLater()
        self.cards.clear()

        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        departments = self.db.get_departments()
        departments_stats = []
        visible_department_ids = []  # <-- Добавляем список ID видимых подразделений

        if self.role == 'admin':
            # Админ видит все подразделения
            for dept in departments:
                stats = self.db.get_statistics_by_department(dept['name'])
                plan = self.db.get_plan(dept['id'], self.current_year)
                departments_stats.append(stats)
                visible_department_ids.append(dept['id'])  # <-- Добавляем ID

                card = ExpandableDepartmentCard(dept['name'], stats, plan)
                card.set_db(self.db)
                self.cards.append(card)
                self.scroll_layout.addWidget(card)
        else:
            # Не-админ: определяем, какие подразделения может видеть пользователь
            user_info = self.db.get_user_by_id(self.user_id)
            user_dict = dict(user_info) if user_info else {}

            allowed_departments = []

            if user_dict.get('is_head') and user_dict.get('department_id'):
                for dept in departments:
                    if dept['id'] == user_dict['department_id']:
                        allowed_departments.append(dept)
                        break

            permissions = self.db.get_user_department_permissions(self.user_id)
            for perm in permissions:
                if perm['can_view']:
                    for dept in departments:
                        if dept['id'] == perm['department_id'] and dept not in allowed_departments:
                            allowed_departments.append(dept)

            if not allowed_departments:
                empty_widget = self._create_empty_widget("У вас нет доступа к подразделениям")
                self.scroll_layout.addWidget(empty_widget)
                self.scroll_layout.addStretch()
                return

            for dept in allowed_departments:
                stats = self.db.get_statistics_by_department_for_user(dept['name'], self.user_id, self.role)
                plan = self.db.get_plan(dept['id'], self.current_year)
                departments_stats.append(stats)
                visible_department_ids.append(dept['id'])  # <-- Добавляем ID

                card = ExpandableDepartmentCard(dept['name'], stats, plan)
                card.set_db(self.db)
                self.cards.append(card)
                self.scroll_layout.addWidget(card)

        # Обновляем сводку - передаем ID видимых подразделений
        self.update_summary(departments_stats, visible_department_ids)  # <-- Передаем ID

        if not departments_stats and self.role != 'admin':
            empty_widget = self._create_empty_widget("У вас нет доступа к подразделениям")
            self.scroll_layout.addWidget(empty_widget)

        self.scroll_layout.addStretch()

    def _create_empty_widget(self, message):
        """Создание виджета для пустого состояния"""
        widget = QFrame()
        widget.setFrameStyle(QFrame.Shape.NoFrame)

        layout = QVBoxLayout(widget)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        icon_label = QLabel()
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_font = QFont()
        icon_font.setPointSize(48)
        icon_label.setFont(icon_font)
        icon_label.setStyleSheet("color: #bdc3c7; margin-bottom: 20px;")

        message_label = QLabel(message)
        message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        message_label.setStyleSheet("color: #7f8c8d; font-size: 16px;")

        sub_label = QLabel("Добавьте подразделения в настройках администратора")
        sub_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub_label.setStyleSheet("color: #95a5a6; font-size: 13px; margin-top: 10px;")

        layout.addStretch()
        layout.addWidget(icon_label)
        layout.addWidget(message_label)
        layout.addWidget(sub_label)
        layout.addStretch()

        widget.setMinimumHeight(300)
        return widget

    def on_year_changed(self, year):
        """Изменение года"""
        self.current_year = int(year)
        self.update_statistics()


class RegionCard(QFrame):
    """Карточка региона"""

    def __init__(self, region_name, stats, parent=None):
        super().__init__(parent)
        self.region_name = region_name if region_name and region_name != "Не указан" else "Вне плана"
        self.stats = stats
        self.is_expanded = False
        self.animation = None
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        # Исправлено: применяем стиль только к конкретному классу карточки, чтобы дочерние QLabel не ломались
        self.setStyleSheet("""
            RegionCard {
                background-color: white;
                border-radius: 12px;
            }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ========== ЗАГОЛОВОК ==========
        self.header_widget = QWidget()
        self.header_widget.setCursor(Qt.CursorShape.PointingHandCursor)
        self.header_widget.setMinimumHeight(70)
        # Исправлено: селектор только для непосредственного виджета хедера
        self.header_widget.setStyleSheet("""
            QWidget#HeaderWidget {
                background-color: #ffffff;
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
            }
            QWidget#HeaderWidget:hover {
                background-color: #f8f9fa;
            }
        """)
        self.header_widget.setObjectName("HeaderWidget")

        def mousePressEvent(event):
            self.toggle_expand()

        self.header_widget.mousePressEvent = mousePressEvent

        header_layout = QHBoxLayout(self.header_widget)
        header_layout.setContentsMargins(20, 15, 20, 15)
        header_layout.setSpacing(15)

        # Стрелка
        self.arrow_label = QLabel()
        self.arrow_label.setFixedSize(24, 24)
        self.arrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.arrow_label.setStyleSheet("font-size: 14px; color: #9b59b6; font-weight: bold;")
        header_layout.addWidget(self.arrow_label)

        # Иконка
        icon_label = QLabel("")
        icon_label.setStyleSheet("font-size: 24px;")
        header_layout.addWidget(icon_label)

        # Название региона
        name_label = QLabel(self.region_name)
        name_font = QFont()
        name_font.setPointSize(13)
        name_font.setBold(True)
        name_label.setFont(name_font)
        name_label.setStyleSheet("color: #2c3e50; border: none; background: transparent;")
        name_label.setMinimumWidth(200)
        header_layout.addWidget(name_label, 2)

        # Данные для заголовка
        total = self.stats.get('total', 0)
        selected = self.stats.get('selected', 0)
        plan = self.stats.get('plan', 0)

        percent = int((selected / plan) * 100) if plan > 0 else 0

        # if percent >= 80:
        #     percent_color = "#2ecc71"
        #     percent_bg = "#d5f5e3"
        # elif percent >= 50:
        #     percent_color = "#f39c12"
        #     percent_bg = "#fdebd0"
        # else:
        #     percent_color = "#e74c3c"
        #     percent_bg = "#fadbd8"
        #
        # percent_widget = QLabel(f"{percent}%")
        # percent_widget.setFixedSize(65, 32)
        # percent_widget.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # percent_widget.setStyleSheet(f"""
        #     background-color: {percent_bg};
        #     color: {percent_color};
        #     font-size: 14px;
        #     font-weight: bold;
        #     border-radius: 16px;
        # """)
        # header_layout.addWidget(percent_widget)

        # Краткая статистика
        stats_widget = QWidget()
        stats_layout = QHBoxLayout(stats_widget)
        stats_layout.setSpacing(15)
        stats_layout.setContentsMargins(0, 0, 0, 0)

        # selected_widget = QLabel(f"{selected}")
        # selected_widget.setStyleSheet("color: #2ecc71; font-size: 14px; font-weight: bold;")
        # stats_layout.addWidget(selected_widget)
        #
        # total_widget = QLabel(f"{total}")
        # total_widget.setStyleSheet("color: #3498db; font-size: 14px; font-weight: bold;")
        # stats_layout.addWidget(total_widget)

        header_layout.addWidget(stats_widget)
        main_layout.addWidget(self.header_widget)

        # ========== ДЕТАЛЬНАЯ ЧАСТЬ ==========
        self.content_widget = QWidget()
        self.content_widget.setVisible(False)
        self.content_widget.setStyleSheet("""
                    QWidget#ContentWidget {
                        background-color: #f8f9fa;
                        border-bottom-left-radius: 12px;
                        border-bottom-right-radius: 12px;
                    }
                """)
        self.content_widget.setObjectName("ContentWidget")

        content_layout = QVBoxLayout(self.content_widget)
        content_layout.setContentsMargins(20, 15, 20, 15)
        content_layout.setSpacing(10)

        # Прогресс-бар
        # if plan > 0:
            # progress_widget = QWidget()
            # progress_layout = QVBoxLayout(progress_widget)
            # progress_layout.setContentsMargins(0, 0, 0, 0)
            # progress_layout.setSpacing(6)
            #
            # progress_header = QWidget()
            # progress_header_layout = QHBoxLayout(progress_header)
            # progress_header_layout.setContentsMargins(0, 0, 0, 0)
            #
            # progress_label = QLabel("Выполнение плана")
            # progress_label.setStyleSheet(
            #     "color: #2c3e50; font-size: 13px; font-weight: bold; background: transparent; border: none;")
            # progress_header_layout.addWidget(progress_label)
            # progress_header_layout.addStretch()
            #
            # progress_value = QLabel(f"{selected} из {plan} ({percent}%)")
            # progress_value.setStyleSheet(
            #     f"color: {percent_color}; font-size: 13px; font-weight: bold; background: transparent; border: none;")
            # progress_header_layout.addWidget(progress_value)
            #
            # progress_layout.addWidget(progress_header)
            #
            # progress_bar = QProgressBar()
            # progress_bar.setMaximum(plan)
            # progress_bar.setValue(selected)
            # progress_bar.setFixedHeight(12)
            # progress_bar.setTextVisible(False)
            # progress_bar.setStyleSheet(f"""
            #             QProgressBar {{
            #                 border: none;
            #                 border-radius: 6px;
            #                 background-color: #e0e0e0;
            #             }}
            #             QProgressBar::chunk {{
            #                 background-color: {percent_color};
            #                 border-radius: 6px;
            #             }}
            #         """)
            # progress_layout.addWidget(progress_bar)
            # content_layout.addWidget(progress_widget)

        # ===== СТАТУС ДОКУМЕНТОВ =====
        docs_label = QLabel("Статус документов:")
        docs_label.setStyleSheet(
            "color: #2c3e50; font-size: 13px; font-weight: bold; margin-top: 5px; background: transparent; border: none;")
        content_layout.addWidget(docs_label)

        docs_widget = QWidget()
        # ИСПРАВЛЕНО: Жестко говорим контейнеру не расти больше, чем высота внутренних плашек (85px)
        # docs_widget.setFixedHeight(85)

        docs_layout = QHBoxLayout(docs_widget)
        docs_layout.setSpacing(12)
        docs_layout.setContentsMargins(0, 0, 0, 0)

        vk = self.stats.get('vk', 0)
        ok = self.stats.get('ok', 0)
        vavko = self.stats.get('vavko', 0)

        docs_layout.addWidget(self._create_stat_block("ВК", vk, "#f39c12"))
        docs_layout.addWidget(self._create_stat_block("ОК", ok, "#8e44ad"))
        # docs_layout.addWidget(self._create_stat_block("Нет", vavko, "#16a085"))
        docs_layout.addStretch()
        content_layout.addWidget(docs_widget)

        # ===== КАТЕГОРИИ =====
        categories_label = QLabel("Распределение по категориям:")
        categories_label.setStyleSheet(
            "color: #2c3e50; font-size: 13px; font-weight: bold; margin-top: 5px; background: transparent; border: none;")
        content_layout.addWidget(categories_label)

        cats_widget = QWidget()
        # ИСПРАВЛЕНО: Задаем фиксированную высоту и для этого контейнера
        # cats_widget.setFixedHeight(85)

        cats_layout = QHBoxLayout(cats_widget)
        cats_layout.setSpacing(12)
        cats_layout.setContentsMargins(0, 0, 0, 0)

        male = self.stats.get('male', 0)
        female = self.stats.get('female', 0)
        military = self.stats.get('military', 0)
        not_selected = self.stats.get('not_selected', 0)

        cats_layout.addWidget(self._create_stat_block("Мужчины", male, "#3498db"))
        # cats_layout.addWidget(self._create_stat_block("Женщины", female, "#e67e22"))
        cats_layout.addWidget(self._create_stat_block("Военнослужащие", military, "#1abc9c"))
        cats_layout.addWidget(self._create_stat_block("Не отобраны", not_selected, "#e74c3c"))
        cats_layout.addStretch()
        content_layout.addWidget(cats_widget)

        # ИСПРАВЛЕНО: Добавляем stretch в самый конец детальной панели,
        # чтобы он забрал на себя все лишнее пространство и не ломал блоки
        content_layout.addStretch(1)

        main_layout.addWidget(self.content_widget)

    def _create_stat_block(self, title, value, color):
        """Создание блока статистики: цифра сверху, текст снизу"""
        widget = QWidget()
        widget.setMinimumWidth(130)
        widget.setMinimumHeight(80)  # Даем минимальную комфортную высоту вместо жесткой фиксации

        widget.setObjectName("StatBlock")
        widget.setStyleSheet("""
            QWidget#StatBlock {
                background-color: white;
                border-radius: 10px;
                border: 1px solid #e8e8e8;
            }
        """)

        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Крупная цифра (сверху)
        value_label = QLabel(str(value))
        value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        value_label.setStyleSheet(f"""
            QLabel {{
                color: {color};
                font-size: 22px;
                font-weight: bold;
                border: none;
                background: transparent;
            }}
        """)
        layout.addWidget(value_label)

        # Подпись (снизу)
        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("""
            QLabel {
                color: #555555;
                font-size: 12px;
                font-weight: 500;
                border: none;
                background: transparent;
            }
        """)
        title_label.setWordWrap(True)
        layout.addWidget(title_label)

        return widget

    def toggle_expand(self):
        """Переключение раскрытия с динамическим расчетом высоты"""
        self.is_expanded = not self.is_expanded

        if self.is_expanded:
            self.arrow_label.setText('')
            self.content_widget.setVisible(True)

            # Вычисляем реальную идеальную высоту всего контента внутри лэйаута
            ideal_height = self.content_widget.layout().sizeHint().height()

            if self.animation:
                self.animation.stop()

            self.animation = QPropertyAnimation(self.content_widget, b"maximumHeight")
            self.animation.setDuration(300)
            self.animation.setStartValue(0)
            self.animation.setEndValue(ideal_height)  # Открываем на реальную высоту (обычно ~360-400px)
            self.animation.setEasingCurve(QEasingCurve.Type.OutCubic)

            # После окончания анимации убираем лимит maximumHeight,
            # чтобы виджет вел себя естественно при изменении размеров окна
            self.animation.finished.connect(lambda: self.content_widget.setMaximumHeight(16777215))
            self.animation.start()
        else:
            self.arrow_label.setText('')

            if self.animation:
                self.animation.stop()

            self.animation = QPropertyAnimation(self.content_widget, b"maximumHeight")
            self.animation.setDuration(250)
            self.animation.setStartValue(self.content_widget.height())
            self.animation.setEndValue(0)
            self.animation.setEasingCurve(QEasingCurve.Type.InCubic)
            self.animation.finished.connect(lambda: self.content_widget.setVisible(False))
            self.animation.start()


class RegionStatsDialog(QDialog):
    """Диалог статистики по регионам"""

    def __init__(self, department_name, department_id, db, role='admin', parent=None):
        super().__init__(parent)
        self.department_name = department_name
        self.department_id = department_id
        self.db = db
        self.role = role  # Добавляем роль
        self.setModal(True)
        self.setWindowTitle(f'Статистика по регионам - {department_name}')
        self.setMinimumSize(950, 700)
        self.resize(1100, 800)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # Заголовок
        header_widget = QWidget()
        header_widget.setFixedHeight(80)
        header_widget.setStyleSheet("""
            QWidget {
                background-color: #9b59b6;
                border-radius: 12px;
            }
        """)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(20, 15, 20, 15)

        icon_label = QLabel()
        icon_label.setStyleSheet("font-size: 32px;")
        header_layout.addWidget(icon_label)

        title_widget = QWidget()
        title_layout = QVBoxLayout(title_widget)
        title_layout.setSpacing(5)
        title_layout.setContentsMargins(0, 0, 0, 0)

        title_label = QLabel("Статистика по регионам")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: white;")
        title_layout.addWidget(title_label)

        subtitle_label = QLabel(f"Подразделение: {self.department_name}")
        subtitle_label.setStyleSheet("color: #d5b8e8; font-size: 13px;")
        title_layout.addWidget(subtitle_label)

        header_layout.addWidget(title_widget)
        header_layout.addStretch()
        layout.addWidget(header_widget)

        # Панель фильтров
        filter_widget = QWidget()
        filter_layout = QHBoxLayout(filter_widget)
        filter_layout.setContentsMargins(0, 5, 0, 5)

        filter_label = QLabel("Фильтр по региону:")
        filter_label.setStyleSheet("font-weight: bold; color: #2c3e50; font-size: 13px;")
        filter_layout.addWidget(filter_label)

        self.region_combo = QComboBox()
        self.region_combo.addItem("Все регионы")
        self.load_regions()
        self.region_combo.currentTextChanged.connect(self.load_stats)
        self.region_combo.setMinimumWidth(250)
        self.region_combo.setStyleSheet("""
            QComboBox {
                padding: 8px 12px;
                border: 1px solid #ddd;
                border-radius: 8px;
                background-color: white;
                font-size: 13px;
            }
        """)
        filter_layout.addWidget(self.region_combo)
        filter_layout.addStretch()

        self.refresh_btn = QPushButton("Обновить")
        self.refresh_btn.setFixedWidth(120)
        self.refresh_btn.setFixedHeight(35)
        self.refresh_btn.clicked.connect(self.load_stats)
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        filter_layout.addWidget(self.refresh_btn)
        layout.addWidget(filter_widget)

        # Сводка (Summary Panel)
        self.summary_widget = QFrame()
        self.summary_widget.setFixedHeight(95)
        self.summary_widget.setStyleSheet("""
            QFrame#SummaryWidget {
                background-color: #f8f9fa;
                border-radius: 12px;
            }
        """)
        self.summary_widget.setObjectName("SummaryWidget")

        # Исправлено: Сразу инициализируем пустой лэйаут для сводки, чтобы потом наполнять его
        summary_layout = QHBoxLayout(self.summary_widget)
        summary_layout.setContentsMargins(20, 15, 20, 15)
        summary_layout.setSpacing(20)
        self.summary_widget.setLayout(summary_layout)

        layout.addWidget(self.summary_widget)

        # Область с прокруткой
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
        self.scroll_layout.setSpacing(10)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)

        self.scroll_area.setWidget(self.scroll_content)
        layout.addWidget(self.scroll_area, 1)

        # Кнопка экспорта
        export_btn = QPushButton("📥 Экспортировать в CSV")
        export_btn.setFixedHeight(45)
        export_btn.clicked.connect(self.export_to_csv)
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        layout.addWidget(export_btn)

        self.load_stats()

    def load_regions(self):
        """Загрузка регионов"""
        regions = self.db.get_regions_for_department(self.department_id)
        self.region_combo.addItem("Вне плана")
        for region in regions:
            self.region_combo.addItem(region['name'])

    def update_summary(self, all_stats):
        """Обновление сводки"""
        summary_layout = self.summary_widget.layout()

        while summary_layout.count():
            item = summary_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # План уже есть в каждом регионе, суммируем
        total_plan = sum(s.get('plan', 0) for s in all_stats)
        total_selected = sum(s.get('selected', 0) for s in all_stats)
        total_applicants = sum(s.get('total', 0) for s in all_stats)
        total_regions = len(all_stats)

        percent = int((total_selected / total_plan) * 100) if total_plan > 0 else 0

        summary_items = [
            ("Регионов", total_regions, "#9b59b6"),
            # ("Абитуриентов", total_applicants, "#3498db"),
            # ("Отобрано", total_selected, "#2ecc71"),
            # ("План", total_plan, "#f39c12"),
            # ("Выполнение", f"{percent}%", "#2ecc71" if percent >= 80 else "#f39c12" if percent >= 50 else "#e74c3c"),
        ]

        for label, value, color in summary_items:
            item_widget = QWidget()
            item_layout = QVBoxLayout(item_widget)
            item_layout.setContentsMargins(0, 0, 0, 0)
            item_layout.setSpacing(4)

            value_widget = QLabel(str(value))
            value_widget.setAlignment(Qt.AlignmentFlag.AlignCenter)
            value_widget.setMinimumWidth(80)
            value_widget.setFixedHeight(35)
            value_widget.setStyleSheet(f"""
                color: {color}; 
                font-size: 16px; 
                font-weight: bold; 
                background-color: white; 
                border-radius: 15px;
            """)

            label_widget = QLabel(label)
            label_widget.setAlignment(Qt.AlignmentFlag.AlignCenter)
            label_widget.setStyleSheet("color: #7f8c8d; font-size: 12px; background: transparent;")

            item_layout.addWidget(value_widget)
            item_layout.addWidget(label_widget)
            summary_layout.addWidget(item_widget)

        summary_layout.addStretch()

    def load_stats(self):
        """Загрузка статистики"""
        region_name = self.region_combo.currentText()
        region_id = None

        if region_name == "Вне плана":
            # Особая логика — показать только внеплановые
            stats = self.db.get_detailed_region_stats(self.department_id, None)
            # Отфильтровать только "Вне плана"
            stats = [s for s in stats if s.get('region_name') == 'Вне плана']
        elif region_name != "Все регионы":
            cursor = self.db.conn.cursor()
            cursor.execute('SELECT id FROM regions WHERE name = ?', (region_name,))
            result = cursor.fetchone()
            if result:
                region_id = result['id']
            stats = self.db.get_detailed_region_stats(self.department_id, region_id)
        else:
            stats = self.db.get_detailed_region_stats(self.department_id, None)

        self.update_summary(stats)

        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for stat in stats:
            stat_dict = dict(stat) if stat else {}
            region_name_display = stat_dict.get('region_name', 'Вне плана')
            if not region_name_display or region_name_display == "Не указан":
                region_name_display = "Вне плана"

            card = RegionCard(region_name_display, stat_dict)
            self.scroll_layout.addWidget(card)

        if not stats:
            empty_widget = self._create_empty_widget()
            self.scroll_layout.addWidget(empty_widget)

        self.scroll_layout.addStretch()

    def _create_empty_widget(self):
        """Пустое состояние"""
        widget = QFrame()
        layout = QVBoxLayout(widget)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(0, 80, 0, 80)

        icon_label = QLabel()
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_label.setStyleSheet("font-size: 64px; color: #bdc3c7;")
        layout.addWidget(icon_label)

        message_label = QLabel("Нет данных по регионам")
        message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        message_label.setStyleSheet("color: #7f8c8d; font-size: 16px; margin-top: 20px;")
        layout.addWidget(message_label)

        return widget

    def export_to_csv(self):
        """Экспорт в CSV"""
        import pandas as pd
        from datetime import datetime

        file_path, _ = QFileDialog.getSaveFileName(
            self, 'Сохранить статистику',
            f'region_stats_{self.department_name}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
            'CSV Files (*.csv);;All Files (*)'
        )

        if not file_path:
            return

        data = []
        for i in range(self.scroll_layout.count()):
            widget = self.scroll_layout.itemAt(i).widget()
            if isinstance(widget, RegionCard):
                stats = widget.stats
                data.append({
                    'Регион': widget.region_name,
                    'План': stats.get('plan', 0),
                    'Отобраны': stats.get('selected', 0),
                    'ВК': stats.get('vk', 0),
                    'ОК': stats.get('ok', 0),
                    'Нет': stats.get('vavko', 0),
                    'Мужчины': stats.get('male', 0),
                    'Женщины': stats.get('female', 0),
                    'Военнослужащие': stats.get('military', 0),
                    'Не отобраны': stats.get('not_selected', 0),
                    'Всего': stats.get('total', 0),
                })

        if data:
            df = pd.DataFrame(data)
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
            QMessageBox.information(self, "Успех", "Статистика экспортирована в файл")
        else:
            QMessageBox.warning(self, "Внимание", "Нет данных для экспорта")