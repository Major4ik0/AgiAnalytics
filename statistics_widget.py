# -*- coding: utf-8 -*-
# statistics_widget.py
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QComboBox, QPushButton, QScrollArea, QFrame,
                             QGridLayout, QSizePolicy)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont, QColor
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


class StatisticsCard(QFrame):
    """Карточка статистики"""

    def __init__(self, title, value, color, icon='', parent=None):
        super().__init__(parent)
        self.title = title
        self.value = value
        self.color = color
        self.icon = icon
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        self.setStyleSheet(f"""
            QFrame {{
                background-color: {self.color};
                border-radius: 12px;
                border: none;
            }}
            QFrame:hover {{
                background-color: {self.lighten_color(self.color)};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        # Заголовок с иконкой
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)

        if self.icon:
            icon_label = QLabel(self.icon)
            icon_label.setStyleSheet("font-size: 20px; color: white;")
            header_layout.addWidget(icon_label)

        title_label = QLabel(self.title)
        title_label.setStyleSheet("""
            color: white; 
            font-size: 14px; 
            font-weight: bold;
        """)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        layout.addLayout(header_layout)

        # Значение
        value_label = QLabel(str(self.value))
        value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        value_font = QFont()
        value_font.setPointSize(32)
        value_font.setBold(True)
        value_label.setFont(value_font)
        value_label.setStyleSheet("""
            color: white;
            padding: 10px 0;
        """)

        layout.addWidget(value_label)

        # Процентная строка (если применимо)
        if self.title in ['Поступают', 'Отказались'] and self.value != 0:
            total = 100  # Временное значение, будет заменено
            if total > 0:
                percentage = (self.value / total) * 100
                percent_label = QLabel(f"{percentage:.1f}%")
                percent_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                percent_label.setStyleSheet("""
                    color: rgba(255, 255, 255, 0.9); 
                    font-size: 13px;
                    font-weight: bold;
                """)
                layout.addWidget(percent_label)

        # Минимальный размер для карточки
        self.setMinimumSize(180, 140)

    def lighten_color(self, color):
        """Осветление цвета"""
        if color.startswith('#'):
            r = min(255, int(color[1:3], 16) + 40)
            g = min(255, int(color[3:5], 16) + 40)
            b = min(255, int(color[5:7], 16) + 40)
            return f'#{r:02x}{g:02x}{b:02x}'
        return color


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


class StatisticsWidget(QWidget):
    """Главный виджет статистики"""

    def __init__(self, user_id, role, db):
        super().__init__()
        self.user_id = user_id
        self.role = role
        self.db = db
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

        title_label = QLabel("📊 Статистика")
        title_font = QFont()
        title_font.setPointSize(22)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50;")

        title_layout.addWidget(title_label)
        title_layout.addStretch()

        # Бейдж роли пользователя
        role_badge = QLabel(f"👤 {self.role.capitalize()}")
        role_badge.setStyleSheet("""
            QLabel {
                background-color: #ecf0f1;
                color: #34495e;
                border-radius: 15px;
                padding: 8px 16px;
                font-weight: bold;
                font-size: 13px;
            }
        """)
        title_layout.addWidget(role_badge)

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

        # Курс
        course_label = QLabel("Курс:")
        course_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.course_combo = QComboBox()
        self.course_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background-color: white;
                min-width: 150px;
            }
            QComboBox:hover {
                border-color: #3498db;
            }
            QComboBox:focus {
                border-color: #2980b9;
            }
        """)

        if self.role == 'admin':
            self.course_combo.addItems(['Все курсы', '1 курс', '2 курс', '3 курс', '4 курс', '5 курс'])
        else:
            # Для обычного пользователя только его курс
            user_info = self.db.conn.execute(
                'SELECT course FROM users WHERE id = ?',
                (self.user_id,)
            ).fetchone()
            user_course = user_info['course'] if user_info else '1 курс'
            self.course_combo.addItems([user_course])
            self.course_combo.setEnabled(False)

        # Категория
        category_label = QLabel("Категория:")
        category_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.category_combo = QComboBox()
        self.category_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background-color: white;
                min-width: 150px;
            }
            QComboBox:hover {
                border-color: #3498db;
            }
            QComboBox:focus {
                border-color: #2980b9;
            }
        """)
        self.category_combo.addItems(['Все категории', 'муж', 'жен', 'в/сл'])

        # Кнопка обновления
        self.refresh_btn = QPushButton("🔄 Обновить")
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #1c6ea4;
            }
        """)
        self.refresh_btn.clicked.connect(self.update_statistics)

        # Добавляем элементы
        controls_layout.addWidget(course_label)
        controls_layout.addWidget(self.course_combo)
        controls_layout.addWidget(category_label)
        controls_layout.addWidget(self.category_combo)
        controls_layout.addStretch()
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
            QScrollBar:vertical {
                background-color: #ecf0f1;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background-color: #bdc3c7;
                border-radius: 6px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #95a5a6;
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

    def update_statistics(self):
        """Обновление статистики"""
        # Очистка предыдущих данных
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        selected_course = self.course_combo.currentText()
        selected_category = self.category_combo.currentText()

        # Получение данных из БД
        if self.role == 'admin':
            if selected_course != 'Все курсы':
                stats = self.db.get_statistics(self.user_id, self.role, selected_course)
                self.display_course_stats(selected_course, stats, selected_category)
            else:
                # Для всех курсов
                self.display_all_courses_stats(selected_category)
        else:
            # Для обычного пользователя
            user_info = self.db.conn.execute(
                'SELECT course FROM users WHERE id = ?',
                (self.user_id,)
            ).fetchone()
            user_course = user_info['course'] if user_info else '1 курс'
            stats = self.db.get_statistics(self.user_id, self.role, user_course)
            self.display_course_stats(user_course, stats, selected_category)

        # Добавляем растягивающийся элемент
        self.scroll_layout.addStretch()

    def display_course_stats(self, course_name, stats, category_filter):
        """Отображение статистики по одному курсу"""
        if stats and stats[0]['total'] > 0:
            stats_dict = self.filter_by_category(dict(stats[0]), category_filter)
            section = CourseSection(course_name, stats_dict)
            self.scroll_layout.addWidget(section)

            # Обновляем информационную панель
            self.info_label.setText(f"📊 Отображается статистика по курсу: {course_name}")
            self.info_panel.setVisible(True)
        else:
            empty_state = EmptyStateWidget(f"Нет данных для курса '{course_name}'")
            self.scroll_layout.addWidget(empty_state)
            self.info_panel.setVisible(False)

    def display_all_courses_stats(self, category_filter):
        """Отображение статистики по всем курсам"""
        courses = ['1 курс', '2 курс', '3 курс', '4 курс', '5 курс']
        has_data = False

        for course in courses:
            stats = self.db.get_statistics(self.user_id, 'admin', course)
            if stats and stats[0]['total'] > 0:
                has_data = True
                stats_dict = self.filter_by_category(dict(stats[0]), category_filter)
                section = CourseSection(course, stats_dict)
                self.scroll_layout.addWidget(section)

        if has_data:
            self.info_label.setText("📊 Отображается статистика по всем курсам")
            self.info_panel.setVisible(True)
        else:
            empty_state = EmptyStateWidget("Нет данных ни по одному курсу")
            self.scroll_layout.addWidget(empty_state)
            self.info_panel.setVisible(False)

    def filter_by_category(self, stats_dict, category):
        """Фильтрация статистики по категории"""
        if category == 'Все категории':
            return stats_dict

        filtered = stats_dict.copy()

        if category == 'муж':
            filtered['male'] = filtered.get('male', 0)
            filtered['female'] = 0
            filtered['military'] = 0
        elif category == 'жен':
            filtered['male'] = 0
            filtered['female'] = filtered.get('female', 0)
            filtered['military'] = 0
        elif category == 'в/сл':
            filtered['male'] = 0
            filtered['female'] = 0
            filtered['military'] = filtered.get('military', 0)

        filtered['total'] = filtered['male'] + filtered['female'] + filtered['military']

        return filtered