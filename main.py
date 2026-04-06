# -*- coding: utf-8 -*-
import sys

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QGridLayout, QStackedWidget, QLabel,
                             QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
                             QHeaderView, QMessageBox, QComboBox, QFrame,
                             QGroupBox, QTabWidget, QDialog, QDialogButtonBox,
                             QFormLayout, QTextEdit, QDateEdit, QSpinBox,
                             QFileDialog, QToolBar, QStatusBar, QMenuBar, QMenu,
                             QScrollArea, QAction, QInputDialog, QCheckBox, QProgressDialog)  # Добавлен QProgressDialog
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor
from datetime import datetime
from database import Database
from resource_helper import get_icon_path, resource_path
from statistics_widget import StatisticsWidget
import pandas as pd
import os

os.environ['QT_MAC_WANTS_LAYER'] = '1'

ICONS = get_icon_path("icon.ico") or "icons/icon.ico"


class LoginWindow(QWidget):
    """Окно авторизации"""
    login_success = pyqtSignal(dict)  # Сигнал успешной авторизации

    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon(ICONS))
        self.setWindowTitle('Авторизация')
        self.setFixedSize(400, 300)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)

        # Заголовок
        title = QLabel('AgiAnalytics')
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # Поля ввода
        form_widget = QWidget()
        form_layout = QFormLayout()

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText('Введите логин')
        self.username_input.setMinimumHeight(40)
        self.username_input.returnPressed.connect(self.authenticate)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText('Введите пароль')
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setMinimumHeight(40)
        self.password_input.returnPressed.connect(self.authenticate)

        form_layout.addRow('Логин:', self.username_input)
        form_layout.addRow('Пароль:', self.password_input)
        form_widget.setLayout(form_layout)
        layout.addWidget(form_widget)

        # Кнопка входа
        self.login_btn = QPushButton('Войти')
        self.login_btn.setMinimumHeight(45)
        self.login_btn.clicked.connect(self.authenticate)
        layout.addWidget(self.login_btn)

        # Информация
        info_label = QLabel()
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info_label.setStyleSheet('color: #666; font-style: italic;')
        layout.addWidget(info_label)
        self.setLayout(layout)

    def keyPressEvent(self, event):
        """Обработка нажатия клавиш"""
        if event.key() == Qt.Key.Key_Return or event.key() == Qt.Key.Key_Enter:
            self.authenticate()
        else:
            super().keyPressEvent(event)

    def authenticate(self):
        username = self.username_input.text().strip()
        password = self.password_input.text().strip()

        if not username or not password:
            QMessageBox.warning(self, 'Ошибка', 'Заполните все поля!')
            return

        db = Database()
        user = db.get_user_by_credentials(username, password)

        if user:
            user_dict = dict(user)
            self.login_success.emit(user_dict)
        else:
            QMessageBox.critical(self, 'Ошибка', 'Неверный логин или пароль!')


class ApplicantDialog(QDialog):
    """Диалог добавления/редактирования абитуриента"""

    def __init__(self, applicant_data=None, user_role='user', db=None, parent=None):
        super().__init__(parent)
        self.applicant_data = applicant_data
        self.user_role = user_role
        self.db = db
        self.setModal(True)

        if applicant_data:
            self.setWindowTitle('✏️ Редактировать абитуриента')
        else:
            self.setWindowTitle('➕ Добавить абитуриента')

        self.setMinimumSize(700, 750)
        self.setMaximumSize(900, 800)
        self.init_ui()

    def init_ui(self):
        """Инициализация интерфейса"""
        # Основной layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)

        # Создаем скролл область для длинной формы
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
        """)

        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setSpacing(20)

        # ========== БЛОК 1: ИНФОРМАЦИЯ ОБ АБИТУРИЕНТЕ ==========
        applicant_group = QGroupBox("📋 Информация об абитуриенте")
        applicant_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 14px;
                border: 2px solid #3498db;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
                color: #3498db;
            }
        """)

        applicant_layout = QFormLayout()
        applicant_layout.setSpacing(12)
        applicant_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        # ФИО абитуриента (обязательное)
        self.applicant_name = QLineEdit()
        self.applicant_name.setPlaceholderText("Иванов Иван Иванович")
        self.applicant_name.setMinimumHeight(35)
        applicant_layout.addRow("ФИО абитуриента *:", self.applicant_name)

        # Субъект РФ
        self.region = QComboBox()
        self.region.setEditable(True)
        self.region.setPlaceholderText("Выберите или введите субъект РФ")
        self.region.addItems([
            "Москва", "Санкт-Петербург", "Московская область", "Ленинградская область",
            "Краснодарский край", "Красноярский край", "Ставропольский край",
            "Республика Татарстан", "Республика Башкортостан", "Свердловская область",
            "Ростовская область", "Новосибирская область", "Челябинская область",
            "Нижегородская область", "Самарская область", "Омская область",
            "Воронежская область", "Пермский край", "Волгоградская область"
        ])
        self.region.setMinimumHeight(35)
        applicant_layout.addRow("Субъект РФ:", self.region)

        # Населенный пункт
        self.city = QLineEdit()
        self.city.setPlaceholderText("Населенный пункт")
        self.city.setMinimumHeight(35)
        applicant_layout.addRow("Населенный пункт:", self.city)

        # Категория
        self.category = QComboBox()
        self.category.addItems(["м", "ж", "всл"])
        self.category.setMinimumHeight(35)
        # Добавляем подсказки
        self.category.setItemText(0, "м (Мужчина)")
        self.category.setItemText(1, "ж (Женщина)")
        self.category.setItemText(2, "всл (Военнослужащий)")
        applicant_layout.addRow("Категория *:", self.category)

        # Телефон с маской
        self.phone = QLineEdit()
        self.phone.setPlaceholderText("+7 (999) 999-99-99")
        self.phone.setInputMask("+7 (999) 999-99-99;_")
        self.phone.setMinimumHeight(35)
        self.phone.textChanged.connect(self.validate_phone)
        applicant_layout.addRow("Телефон *:", self.phone)

        # Образование
        self.education = QComboBox()
        self.education.setMinimumHeight(35)
        self.load_education_types()
        applicant_layout.addRow("Образование:", self.education)

        # Статус
        self.status = QComboBox()
        self.status.addItems(["поступает", "отказывается"])
        self.status.setMinimumHeight(35)
        # Добавляем иконки в текст
        self.status.setItemText(0, "✅ Поступает")
        self.status.setItemText(1, "❌ Отказывается")
        applicant_layout.addRow("Статус *:", self.status)

        # Статус документов
        self.document_status = QComboBox()
        self.document_status.setMinimumHeight(35)
        self.load_document_statuses()
        applicant_layout.addRow("Документы:", self.document_status)

        applicant_group.setLayout(applicant_layout)
        scroll_layout.addWidget(applicant_group)

        # ========== БЛОК 2: ИНФОРМАЦИЯ ОБ АГИТАТОРЕ ==========
        agitator_group = QGroupBox("👤 Информация об агитаторе")
        agitator_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 14px;
                border: 2px solid #2ecc71;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
                color: #2ecc71;
            }
        """)

        agitator_layout = QFormLayout()
        agitator_layout.setSpacing(12)
        agitator_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        # Тип агитатора
        self.agitator_type_widget = QWidget()
        agitator_type_layout = QHBoxLayout(self.agitator_type_widget)
        agitator_type_layout.setContentsMargins(0, 0, 0, 0)
        agitator_type_layout.setSpacing(15)

        self.agitator_is_cadet = QCheckBox("Агитатор - курсант")
        self.agitator_is_cadet.setStyleSheet("""
            QCheckBox {
                spacing: 8px;
                font-weight: bold;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
        """)
        self.agitator_is_cadet.stateChanged.connect(self.on_agitator_type_changed)

        self.agitator_is_officer = QCheckBox("Агитатор - офицер/военнослужащий")
        self.agitator_is_officer.setStyleSheet("""
            QCheckBox {
                spacing: 8px;
                font-weight: bold;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
        """)
        self.agitator_is_officer.stateChanged.connect(self.on_agitator_type_changed)

        agitator_type_layout.addWidget(self.agitator_is_cadet)
        agitator_type_layout.addWidget(self.agitator_is_officer)
        agitator_type_layout.addStretch()
        agitator_layout.addRow("Тип агитатора:", self.agitator_type_widget)

        # Подразделение
        self.agitator_department = QComboBox()
        self.agitator_department.currentTextChanged.connect(self.on_department_changed)
        self.agitator_department.setEditable(True)
        self.agitator_department.setMinimumHeight(35)
        self.load_departments()
        agitator_layout.addRow("Подразделение:", self.agitator_department)

        # ФИО агитатора
        self.agitator_name = QLineEdit()
        self.agitator_name.setPlaceholderText("Фамилия И.О.")
        self.agitator_name.setMinimumHeight(35)
        agitator_layout.addRow("ФИО агитатора *:", self.agitator_name)

        # Контейнер для полей курсанта (по умолчанию скрыт)
        self.cadet_widget = QWidget()
        cadet_layout = QGridLayout(self.cadet_widget)
        cadet_layout.setContentsMargins(0, 0, 0, 0)
        cadet_layout.setSpacing(10)

        # Курс (для курсанта)
        self.agitator_course = QComboBox()
        self.agitator_course.addItems(["1 курс", "2 курс", "3 курс", "4 курс", "5 курс"])
        self.agitator_course.setMinimumHeight(35)
        cadet_layout.addWidget(QLabel("Курс:"), 0, 0)
        cadet_layout.addWidget(self.agitator_course, 0, 1)

        # Группа (для курсанта)
        self.agitator_group = QLineEdit()
        self.agitator_group.setPlaceholderText("Номер группы")
        self.agitator_group.setMinimumHeight(35)
        cadet_layout.addWidget(QLabel("Группа:"), 1, 0)
        cadet_layout.addWidget(self.agitator_group, 1, 1)

        agitator_layout.addRow("", self.cadet_widget)

        # Контейнер для полей офицера/военнослужащего (по умолчанию скрыт)
        self.officer_widget = QWidget()
        officer_layout = QHBoxLayout(self.officer_widget)
        officer_layout.setContentsMargins(0, 0, 0, 0)

        # Звание (для офицера/военнослужащего)
        self.agitator_rank = QComboBox()
        self.agitator_rank.addItems([
            "ряд.", "ефр.", "мл. серж.", "серж.", "ст. серж.", "старшина",
            "прапорщик", "ст. прапорщик", "мл. лейт.", "лейтенант", "ст. лейтенант",
            "капитан", "майор", "подполковник", "полковник", "генерал-майор",
            "генерал-лейтенант", "генерал-полковник"
        ])
        self.agitator_rank.setMinimumHeight(35)
        officer_layout.addWidget(QLabel("Звание:"))
        officer_layout.addWidget(self.agitator_rank)
        officer_layout.addStretch()

        agitator_layout.addRow("", self.officer_widget)

        agitator_group.setLayout(agitator_layout)
        scroll_layout.addWidget(agitator_group)

        # ========== БЛОК 3: ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ ==========
        extra_group = QGroupBox("📝 Дополнительная информация")
        extra_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 14px;
                border: 2px solid #e67e22;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
                color: #e67e22;
            }
        """)

        extra_layout = QFormLayout()
        extra_layout.setSpacing(12)

        # Примечания
        self.notes = QLineEdit()
        self.notes.setPlaceholderText("Дополнительные замечания, комментарии...")
        self.notes.setMinimumHeight(35)
        extra_layout.addRow("Примечания:", self.notes)

        extra_group.setLayout(extra_layout)
        scroll_layout.addWidget(extra_group)

        # Информация об обязательных полях
        info_label = QLabel("* - поля, обязательные для заполнения")
        info_label.setStyleSheet("color: #e74c3c; font-size: 11px; margin-top: 5px;")
        scroll_layout.addWidget(info_label)

        scroll_layout.addStretch()
        scroll_area.setWidget(scroll_widget)
        main_layout.addWidget(scroll_area)

        # Кнопки
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.validate_and_accept)
        button_box.rejected.connect(self.reject)

        # Стилизация кнопок
        ok_button = button_box.button(QDialogButtonBox.StandardButton.Ok)
        ok_button.setText("💾 Сохранить")
        ok_button.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)

        cancel_button = button_box.button(QDialogButtonBox.StandardButton.Cancel)
        cancel_button.setText("❌ Отмена")
        cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
        """)

        main_layout.addWidget(button_box)

        # Заполнение данных если редактирование
        if self.applicant_data:
            self.load_applicant_data()

        # Устанавливаем начальное состояние полей
        self.on_agitator_type_changed()

    def load_education_types(self):
        """Загрузка типов образования из БД"""
        if self.db:
            education_list = self.db.get_education_types()
            self.education.clear()
            self.education.addItems(education_list)
        else:
            # Данные по умолчанию
            self.education.addItems(["СОШ", "СПО", "СВУ", "ПКУ", "КК"])

    def load_document_statuses(self):
        """Загрузка статусов документов из БД"""
        if self.db:
            statuses = self.db.get_document_statuses()
            self.document_status.clear()
            self.document_status.addItems(statuses)
        else:
            # Данные по умолчанию
            self.document_status.addItems(["ВК", "ОК", "ВА ВКО"])

    def load_departments(self):
        """Загрузка подразделений из БД (без дублей)"""
        if self.db:
            departments = self.db.get_departments()
            self.agitator_department.clear()

            # Используем set для уникальности
            unique_departments = set()
            for dept in departments:
                unique_departments.add(dept['name'])

            # Добавляем "Выберите подразделение" первым
            self.agitator_department.addItem("")

            # Добавляем уникальные подразделения
            for dept_name in sorted(unique_departments):
                self.agitator_department.addItem(dept_name)
        else:
            # Данные по умолчанию
            self.agitator_department.addItems([
                "",
                "Факультет 1",
                "Факультет 2",
                "Кафедра 1",
                "Кафедра 2"
            ])

    def on_agitator_type_changed(self):
        """Обработка изменения типа агитатора"""
        is_cadet = self.agitator_is_cadet.isChecked()
        is_officer = self.agitator_is_officer.isChecked()

        # Получаем выбранное подразделение
        department = self.agitator_department.currentText()

        # Проверяем, может ли быть курсант в этом подразделении
        can_be_cadet = 'Факультет 1' in department or 'Факультет 2' in department

        # Обновляем состояние чекбокса курсанта
        self.agitator_is_cadet.setEnabled(can_be_cadet)

        if not can_be_cadet and is_cadet:
            # Если выбрано неподходящее подразделение, переключаем на офицера
            self.agitator_is_cadet.setChecked(False)
            self.agitator_is_officer.setChecked(True)
            is_cadet = False
            is_officer = True

        # Показываем/скрываем соответствующие поля
        self.cadet_widget.setVisible(is_cadet and can_be_cadet)
        self.officer_widget.setVisible(is_officer)

        # Если ни один не выбран, показываем оба (по умолчанию курсант если можно)
        if not is_cadet and not is_officer:
            if can_be_cadet:
                self.agitator_is_cadet.setChecked(True)
            else:
                self.agitator_is_officer.setChecked(True)

    def on_department_changed(self, text):
        """Обработка изменения подразделения"""
        self.on_agitator_type_changed()

    def validate_phone(self, text):
        """Валидация номера телефона"""
        # Удаляем маску для проверки
        clean_text = text.replace('(', '').replace(')', '').replace('-', '').replace(' ', '').replace('_', '').replace(
            '+', '')

        if len(clean_text) < 11:
            self.phone.setStyleSheet("border: 1px solid #e74c3c;")
        else:
            self.phone.setStyleSheet("border: 1px solid #2ecc71;")

    def validate_and_accept(self):
        """Валидация обязательных полей (только те, что в таблице)"""
        errors = []

        # Обязательные поля из таблицы:
        # 1. ФИО абитуриента
        if not self.applicant_name.text().strip():
            errors.append("ФИО абитуриента")
            self.applicant_name.setStyleSheet("border: 2px solid #e74c3c;")
        else:
            self.applicant_name.setStyleSheet("")

        # 2. Телефон (должен быть заполнен)
        phone_clean = self.clean_phone_number(self.phone.text())
        if not phone_clean or len(phone_clean) < 10:
            errors.append("Телефон")
            self.phone.setStyleSheet("border: 2px solid #e74c3c;")
        else:
            self.phone.setStyleSheet("")

        # 3. ФИО агитатора (обязательно)
        if not self.agitator_name.text().strip():
            errors.append("ФИО агитатора")
            self.agitator_name.setStyleSheet("border: 2px solid #e74c3c;")
        else:
            self.agitator_name.setStyleSheet("")

        # 4. Подразделение (обязательно)
        if not self.agitator_department.currentText().strip():
            errors.append("Подразделение")
            self.agitator_department.setStyleSheet("border: 2px solid #e74c3c;")
        else:
            self.agitator_department.setStyleSheet("")

        # Категория - всегда выбрана по умолчанию, не проверяем
        # Статус - всегда выбран по умолчанию, не проверяем

        # Для курсанта: если выбран тип "курсант", то группа обязательна
        if self.agitator_is_cadet.isChecked():
            if not self.agitator_group.text().strip():
                errors.append("Группа (для курсанта)")
                self.agitator_group.setStyleSheet("border: 2px solid #e74c3c;")
            else:
                self.agitator_group.setStyleSheet("")

        # Если есть ошибки
        if errors:
            error_msg = "Пожалуйста, заполните следующие обязательные поля:\n• " + "\n• ".join(errors)
            QMessageBox.warning(self, "Ошибка валидации", error_msg)
            return

        self.accept()

    def clean_phone_number(self, phone):
        """Очистка номера телефона от форматирования"""
        if not phone:
            return ""

        # Удаляем все нецифровые символы
        digits = ''.join(filter(str.isdigit, phone))

        if not digits:
            return ""

        # Приводим к формату 7XXXXXXXXXX
        if digits.startswith('8') and len(digits) == 11:
            return '7' + digits[1:]
        elif len(digits) == 10:
            return '7' + digits
        elif digits.startswith('7') and len(digits) == 11:
            return digits

        return digits

    def load_applicant_data(self):
        """Загрузка данных абитуриента для редактирования"""
        # Абитуриент
        self.applicant_name.setText(self.applicant_data.get('applicant_name', ''))

        # Регион
        region = self.applicant_data.get('region', '')
        if region:
            index = self.region.findText(region)
            if index >= 0:
                self.region.setCurrentIndex(index)
            else:
                self.region.setEditText(region)

        self.city.setText(self.applicant_data.get('city', ''))

        # Категория
        category = self.applicant_data.get('category', 'м')
        if category == 'м':
            self.category.setCurrentIndex(0)
        elif category == 'ж':
            self.category.setCurrentIndex(1)
        else:
            self.category.setCurrentIndex(2)

        # Телефон
        phone = self.applicant_data.get('phone', '')
        if phone:
            formatted = self.format_phone_for_display(phone)
            self.phone.setText(formatted)

        # Образование
        edu = self.applicant_data.get('education', '')
        if edu:
            index = self.education.findText(edu)
            if index >= 0:
                self.education.setCurrentIndex(index)

        # Статус
        status = self.applicant_data.get('status', 'поступает')
        self.status.setCurrentIndex(0 if status == 'поступает' else 1)

        # Документы
        doc_status = self.applicant_data.get('document_status', '')
        if doc_status:
            index = self.document_status.findText(doc_status)
            if index >= 0:
                self.document_status.setCurrentIndex(index)

        # Агитатор
        department = self.applicant_data.get('agitator_department', '')
        if department:
            index = self.agitator_department.findText(department)
            if index >= 0:
                self.agitator_department.setCurrentIndex(index)
            else:
                self.agitator_department.setEditText(department)

        self.agitator_name.setText(self.applicant_data.get('agitator_name', ''))

        # Тип агитатора
        is_cadet = self.applicant_data.get('agitator_is_cadet', False)
        if is_cadet:
            self.agitator_is_cadet.setChecked(True)
            course = self.applicant_data.get('agitator_course', '1 курс')
            index = self.agitator_course.findText(course)
            if index >= 0:
                self.agitator_course.setCurrentIndex(index)
            self.agitator_group.setText(self.applicant_data.get('agitator_group', ''))
        else:
            self.agitator_is_officer.setChecked(True)
            rank = self.applicant_data.get('agitator_rank', '')
            if rank:
                index = self.agitator_rank.findText(rank)
                if index >= 0:
                    self.agitator_rank.setCurrentIndex(index)

        # Примечания
        self.notes.setText(self.applicant_data.get('notes', ''))

        # Обновляем состояние полей
        self.on_agitator_type_changed()

    @staticmethod
    def format_phone_for_display(phone):
        """Форматирование телефона для отображения в маске"""
        digits = ''.join(filter(str.isdigit, str(phone)))

        if len(digits) >= 11:
            # Формат +7 (XXX) XXX-XX-XX
            return f"+7 ({digits[1:4]}) {digits[4:7]}-{digits[7:9]}-{digits[9:11]}"
        return phone

    def get_data(self):
        """Получение данных из формы"""
        # Получаем чистый телефон
        phone = self.clean_phone_number(self.phone.text())

        # Получаем категорию
        category_map = {0: 'м', 1: 'ж', 2: 'всл'}
        category = category_map.get(self.category.currentIndex(), 'м')

        # Получаем статус
        status = 'поступает' if self.status.currentIndex() == 0 else 'отказывается'

        # Данные абитуриента
        data = {
            'applicant_name': self.applicant_name.text().strip(),
            'region': self.region.currentText().strip() if self.region.currentText() != "Выберите или введите субъект РФ" else "",
            'city': self.city.text().strip(),
            'category': category,
            'phone': phone,
            'education': self.education.currentText(),
            'status': status,
            'document_status': self.document_status.currentText(),
            'agitator_department': self.agitator_department.currentText().strip(),
            'agitator_name': self.agitator_name.text().strip(),
            'agitator_is_cadet': self.agitator_is_cadet.isChecked(),
            'notes': self.notes.text().strip(),
        }

        # Данные в зависимости от типа агитатора
        if self.agitator_is_cadet.isChecked():
            data['agitator_course'] = self.agitator_course.currentText()
            data['agitator_group'] = self.agitator_group.text().strip()
            data['agitator_rank'] = ''
        else:
            data['agitator_course'] = ''
            data['agitator_group'] = ''
            data['agitator_rank'] = self.agitator_rank.currentText()

        return data


# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QLineEdit, QComboBox, QPushButton, QLabel,
                             QDialogButtonBox, QGroupBox, QCheckBox, QWidget,
                             QScrollArea, QGridLayout)
from PyQt5.QtCore import Qt


class AdvancedSearchDialog(QDialog):
    """Диалог расширенного поиска"""

    def __init__(self, db=None, parent=None):
        super().__init__(parent)
        self.db = db
        self.filters = {}
        self.setModal(True)
        self.setWindowTitle('🔍 Расширенный поиск')
        self.setMinimumSize(500, 450)
        self.setMaximumSize(700, 600)
        self.init_ui()

    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Заголовок
        title_label = QLabel("Расширенный поиск абитуриентов")
        title_font = title_label.font()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 10px;")
        layout.addWidget(title_label)

        # Создаем скролл область
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; }")

        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setSpacing(15)

        # ========== БЛОК 1: ИНФОРМАЦИЯ ОБ АБИТУРИЕНТЕ ==========
        applicant_group = QGroupBox("📋 Информация об абитуриенте")
        applicant_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)

        applicant_layout = QFormLayout()
        applicant_layout.setSpacing(10)

        # ФИО абитуриента
        self.applicant_name = QLineEdit()
        self.applicant_name.setPlaceholderText("Введите ФИО полностью или частично")
        self.applicant_name.setClearButtonEnabled(True)
        applicant_layout.addRow("ФИО абитуриента:", self.applicant_name)

        # Субъект РФ
        self.region = QComboBox()
        self.region.setEditable(True)
        self.region.setPlaceholderText("Выберите или введите субъект РФ")
        self.region.setClearButtonEnabled(True)
        self.load_regions()
        applicant_layout.addRow("Субъект РФ:", self.region)

        # Населенный пункт
        self.city = QLineEdit()
        self.city.setPlaceholderText("Введите населенный пункт")
        self.city.setClearButtonEnabled(True)
        applicant_layout.addRow("Населенный пункт:", self.city)

        # Категория
        self.category = QComboBox()
        self.category.addItems(["все", "м", "ж", "всл"])
        self.category.setItemText(0, "Все категории")
        self.category.setItemText(1, "Мужчина")
        self.category.setItemText(2, "Женщина")
        self.category.setItemText(3, "Военнослужащий")
        applicant_layout.addRow("Категория:", self.category)

        # Образование (множественный выбор)
        self.education_group = QGroupBox("Образование")
        self.education_group.setStyleSheet("QGroupBox { border: none; margin-top: 5px; }")
        education_layout = QHBoxLayout(self.education_group)

        self.education_checkboxes = []
        education_types = self.get_education_types()
        for edu in education_types:
            checkbox = QCheckBox(edu)
            self.education_checkboxes.append(checkbox)
            education_layout.addWidget(checkbox)
        education_layout.addStretch()

        applicant_layout.addRow("", self.education_group)

        # Статус
        self.status = QComboBox()
        self.status.addItems(["все", "поступает", "отказывается"])
        self.status.setItemText(0, "Все статусы")
        applicant_layout.addRow("Статус:", self.status)

        applicant_group.setLayout(applicant_layout)
        scroll_layout.addWidget(applicant_group)

        # ========== БЛОК 2: ИНФОРМАЦИЯ ОБ АГИТАТОРЕ ==========
        agitator_group = QGroupBox("👤 Информация об агитаторе")
        agitator_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)

        agitator_layout = QFormLayout()
        agitator_layout.setSpacing(10)

        # ФИО агитатора
        self.agitator_name = QLineEdit()
        self.agitator_name.setPlaceholderText("Введите ФИО агитатора")
        self.agitator_name.setClearButtonEnabled(True)
        agitator_layout.addRow("ФИО агитатора:", self.agitator_name)

        # Подразделение
        self.agitator_department = QComboBox()
        self.agitator_department.setEditable(True)
        self.agitator_department.setPlaceholderText("Выберите или введите подразделение")
        self.agitator_department.setClearButtonEnabled(True)
        self.load_departments()
        agitator_layout.addRow("Подразделение:", self.agitator_department)

        # Тип агитатора
        self.agitator_type = QComboBox()
        self.agitator_type.addItems(["все", "курсант", "офицер/военнослужащий"])
        self.agitator_type.setItemText(0, "Все")
        agitator_layout.addRow("Тип агитатора:", self.agitator_type)

        agitator_group.setLayout(agitator_layout)
        scroll_layout.addWidget(agitator_group)

        # ========== БЛОК 3: ДОПОЛНИТЕЛЬНЫЕ ПАРАМЕТРЫ ==========
        extra_group = QGroupBox("📅 Дополнительные параметры")
        extra_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)

        extra_layout = QGridLayout()
        extra_layout.setSpacing(10)

        # Статус документов
        extra_layout.addWidget(QLabel("Документы:"), 0, 0)
        self.document_status = QComboBox()
        self.document_status.setEditable(True)
        self.document_status.setClearButtonEnabled(True)
        self.load_document_statuses()
        extra_layout.addWidget(self.document_status, 0, 1)

        # Курс (для курсанта)
        extra_layout.addWidget(QLabel("Курс:"), 1, 0)
        self.course = QComboBox()
        self.course.addItems(["все", "1 курс", "2 курс", "3 курс", "4 курс", "5 курс"])
        self.course.setItemText(0, "Все курсы")
        extra_layout.addWidget(self.course, 1, 1)

        # Группа
        extra_layout.addWidget(QLabel("Группа:"), 2, 0)
        self.group = QLineEdit()
        self.group.setPlaceholderText("Номер группы")
        self.group.setClearButtonEnabled(True)
        extra_layout.addWidget(self.group, 2, 1)

        extra_group.setLayout(extra_layout)
        scroll_layout.addWidget(extra_group)

        # Кнопки быстрого выбора
        quick_buttons_widget = QWidget()
        quick_buttons_layout = QHBoxLayout(quick_buttons_widget)

        self.clear_all_btn = QPushButton("🗑️ Очистить все")
        self.clear_all_btn.clicked.connect(self.clear_all_filters)

        self.select_none_btn = QPushButton("☐ Снять все выделения")
        self.select_none_btn.clicked.connect(self.clear_education_selection)

        quick_buttons_layout.addWidget(self.clear_all_btn)
        quick_buttons_layout.addWidget(self.select_none_btn)
        quick_buttons_layout.addStretch()

        scroll_layout.addWidget(quick_buttons_widget)
        scroll_layout.addStretch()

        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)

        # Кнопки диалога
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )

        ok_button = button_box.button(QDialogButtonBox.StandardButton.Ok)
        ok_button.setText("🔍 Найти")
        ok_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)

        cancel_button = button_box.button(QDialogButtonBox.StandardButton.Cancel)
        cancel_button.setText("❌ Отмена")

        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        layout.addWidget(button_box)

    def load_regions(self):
        """Загрузка субъектов РФ (уникальные из БД)"""
        if self.db:
            cursor = self.db.conn.cursor()
            cursor.execute(
                'SELECT DISTINCT region FROM applicants WHERE region IS NOT NULL AND region != "" ORDER BY region')
            regions = [row['region'] for row in cursor.fetchall()]
            self.region.addItems(regions)
        else:
            # Данные по умолчанию
            default_regions = [
                "Москва", "Санкт-Петербург", "Московская область", "Ленинградская область",
                "Краснодарский край", "Красноярский край", "Ставропольский край"
            ]
            self.region.addItems(default_regions)

    def load_departments(self):
        """Загрузка подразделений (уникальные из БД)"""
        if self.db:
            cursor = self.db.conn.cursor()
            cursor.execute(
                'SELECT DISTINCT agitator_department FROM applicants WHERE agitator_department IS NOT NULL AND agitator_department != "" ORDER BY agitator_department')
            departments = [row['agitator_department'] for row in cursor.fetchall()]
            self.agitator_department.addItems(departments)
        else:
            self.agitator_department.addItems(["Факультет 1", "Факультет 2", "Кафедра 1", "Кафедра 2"])

    def get_education_types(self):
        """Получение типов образования"""
        if self.db:
            return self.db.get_education_types()
        return ["СОШ", "СПО", "СВУ", "ПКУ", "КК"]

    def load_document_statuses(self):
        """Загрузка статусов документов"""
        if self.db:
            statuses = self.db.get_document_statuses()
            self.document_status.addItems([""] + statuses)
        else:
            self.document_status.addItems(["", "ВК", "ОК", "ВА ВКО"])

    def clear_all_filters(self):
        """Очистка всех фильтров"""
        self.applicant_name.clear()
        self.region.setCurrentIndex(0)
        self.city.clear()
        self.category.setCurrentIndex(0)
        self.status.setCurrentIndex(0)
        self.agitator_name.clear()
        self.agitator_department.setCurrentIndex(0)
        self.agitator_type.setCurrentIndex(0)
        self.document_status.setCurrentIndex(0)
        self.course.setCurrentIndex(0)
        self.group.clear()
        self.clear_education_selection()

    def clear_education_selection(self):
        """Снятие всех выделений образования"""
        for checkbox in self.education_checkboxes:
            checkbox.setChecked(False)

    def get_filters(self):
        """Получение выбранных фильтров"""
        filters = {}

        # Абитуриент
        if self.applicant_name.text().strip():
            filters['applicant_name'] = self.applicant_name.text().strip()

        if self.region.currentText().strip() and self.region.currentText() != "Выберите или введите субъект РФ":
            filters['region'] = self.region.currentText().strip()

        if self.city.text().strip():
            filters['city'] = self.city.text().strip()

        category = self.category.currentText()
        if category != "все":
            filters['category'] = 'м' if category == "Мужчина" else 'ж' if category == "Женщина" else 'всл'

        # Образование (множественный выбор)
        selected_education = [cb.text() for cb in self.education_checkboxes if cb.isChecked()]
        if selected_education:
            filters['education'] = selected_education

        status = self.status.currentText()
        if status != "все":
            filters['status'] = status

        # Агитатор
        if self.agitator_name.text().strip():
            filters['agitator_name'] = self.agitator_name.text().strip()

        if self.agitator_department.currentText().strip():
            filters['agitator_department'] = self.agitator_department.currentText().strip()

        agitator_type = self.agitator_type.currentText()
        if agitator_type != "все":
            filters['agitator_is_cadet'] = (agitator_type == "курсант")

        # Дополнительные
        if self.document_status.currentText().strip():
            filters['document_status'] = self.document_status.currentText().strip()

        course = self.course.currentText()
        if course != "все":
            filters['agitator_course'] = course

        if self.group.text().strip():
            filters['agitator_group'] = self.group.text().strip()

        return filters


class SelectionStatsDialog(QDialog):
    """Диалог со статистикой по выделенным строкам"""

    def __init__(self, selected_data, parent=None):
        super().__init__(parent)
        self.selected_data = selected_data
        self.setModal(False)  # Не модальный, чтобы можно было продолжать работу
        self.setWindowTitle('Статистика по выделенным записям')
        self.setFixedSize(500, 600)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Заголовок
        title_label = QLabel(f"📊 Статистика по выделенным записям")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 15px;")
        layout.addWidget(title_label)

        # Информация о количестве
        count_label = QLabel(f"Выделено записей: {len(self.selected_data)}")
        count_label.setStyleSheet("font-weight: bold; font-size: 13px; margin-bottom: 10px;")
        layout.addWidget(count_label)

        # Статистика
        stats_group = QGroupBox("Статистика по выделенным записям")
        stats_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)

        stats_layout = QFormLayout()
        stats_layout.setSpacing(12)

        # Вычисляем статистику
        stats = self.calculate_stats()

        # Добавляем статистику
        stats_layout.addRow("👥 Всего:", QLabel(str(stats['total'])))
        stats_layout.addRow("✅ Поступают:", QLabel(f"{stats['applying']} ({stats['applying_percent']:.1f}%)"))
        stats_layout.addRow("❌ Отказались:", QLabel(f"{stats['refused']} ({stats['refused_percent']:.1f}%)"))

        # Разделитель
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        stats_layout.addRow(line)

        stats_layout.addRow("👨 Мужчины:", QLabel(f"{stats['male']} ({stats['male_percent']:.1f}%)"))
        stats_layout.addRow("👩 Женщины:", QLabel(f"{stats['female']} ({stats['female_percent']:.1f}%)"))
        stats_layout.addRow("🎖️ Военнослужащие:", QLabel(f"{stats['military']} ({stats['military_percent']:.1f}%)"))

        # Разделитель
        line2 = QFrame()
        line2.setFrameShape(QFrame.Shape.HLine)
        line2.setFrameShadow(QFrame.Shadow.Sunken)
        stats_layout.addRow(line2)

        stats_layout.addRow("📋 Формируется в военкомате:", QLabel(str(stats['doc1'])))
        stats_layout.addRow("📤 Отправлено в ВА ВКО:", QLabel(str(stats['doc2'])))
        stats_layout.addRow("📥 В ВА ВКО:", QLabel(str(stats['doc3'])))

        # Статистика по курсам (если есть разные курсы)
        if len(stats['courses']) > 1:
            line3 = QFrame()
            line3.setFrameShape(QFrame.Shape.HLine)
            line3.setFrameShadow(QFrame.Shadow.Sunken)
            stats_layout.addRow(line3)

            courses_text = ", ".join([f"{k}: {v}" for k, v in stats['courses'].items()])
            stats_layout.addRow("📚 По курсам:", QLabel(courses_text))

        stats_group.setLayout(stats_layout)
        layout.addWidget(stats_group)

        # Кнопка копирования
        copy_btn = QPushButton("📋 Копировать статистику")
        copy_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
                font-weight: bold;
                margin-top: 15px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        copy_btn.clicked.connect(self.copy_stats)
        layout.addWidget(copy_btn)

        # Кнопка закрытия
        close_btn = QPushButton("Закрыть")
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px;
                margin-top: 5px;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
        """)
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)

        self.setLayout(layout)

    def calculate_stats(self):
        """Вычисление статистики по выделенным данным"""
        stats = {
            'total': len(self.selected_data),
            'applying': 0,
            'refused': 0,
            'male': 0,
            'female': 0,
            'military': 0,
            'doc1': 0,
            'doc2': 0,
            'doc3': 0,
            'courses': {},
            'applying_percent': 0,
            'refused_percent': 0,
            'male_percent': 0,
            'female_percent': 0,
            'military_percent': 0
        }
        for row_data in self.selected_data:
            # Статус поступления
            status = row_data.get('status', '')
            if '1)поступает' in status or status == 'поступает':

                stats['applying'] += 1
            elif '2)не поступает' in status or status == 'не поступает':
                stats['refused'] += 1

            # Категория
            category = row_data.get('category', '')
            if category == 'муж':
                stats['male'] += 1
            elif category == 'жен':
                stats['female'] += 1
            elif category == 'в/сл':
                stats['military'] += 1

            # Статус документов
            doc_status = row_data.get('document_status', '')
            if 'Формируется' in doc_status:
                stats['doc1'] += 1
            elif 'Отправлено' in doc_status:
                stats['doc2'] += 1
            elif 'В ВА ВКО' in doc_status:
                stats['doc3'] += 1

            # Курс
            course = row_data.get('course', '')
            if course:
                stats['courses'][course] = stats['courses'].get(course, 0) + 1

        # Вычисляем проценты
        total = stats['total']
        if total > 0:
            stats['applying_percent'] = (stats['applying'] / total) * 100
            stats['refused_percent'] = (stats['refused'] / total) * 100
            stats['male_percent'] = (stats['male'] / total) * 100
            stats['female_percent'] = (stats['female'] / total) * 100
            stats['military_percent'] = (stats['military'] / total) * 100

        return stats

    def copy_stats(self):
        """Копирование статистики в буфер обмена"""
        stats = self.calculate_stats()

        clipboard_text = f"""
📊 СТАТИСТИКА ПО ВЫДЕЛЕННЫМ ЗАПИСЯМ
{'=' * 40}

Выделено записей: {stats['total']}

📈 СТАТУС ПОСТУПЛЕНИЯ:
  • Поступают: {stats['applying']} ({stats['applying_percent']:.1f}%)
  • Отказались: {stats['refused']} ({stats['refused_percent']:.1f}%)

👥 КАТЕГОРИИ:
  • Мужчины: {stats['male']} ({stats['male_percent']:.1f}%)
  • Женщины: {stats['female']} ({stats['female_percent']:.1f}%)
  • Военнослужащие: {stats['military']} ({stats['military_percent']:.1f}%)

📄 СТАТУС ДОКУМЕНТОВ:
  • Формируется в военкомате: {stats['doc1']}
  • Отправлено в ВА ВКО: {stats['doc2']}
  • В ВА ВКО: {stats['doc3']}

📚 ПО КУРСАМ:
"""
        for course, count in stats['courses'].items():
            percent = (count / stats['total']) * 100 if stats['total'] > 0 else 0
            clipboard_text += f"  • {course}: {count} ({percent:.1f}%)\n"

        clipboard_text += f"\n{'=' * 40}\n"
        clipboard_text += f"📅 {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}"

        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_text)

        QMessageBox.information(self, "Успех", "Статистика скопирована в буфер обмена!")


class UserDialog(QDialog):
    """Диалог добавления/редактирования пользователя"""

    def __init__(self, user_data=None, parent=None):
        super().__init__(parent)
        self.user_data = user_data
        self.setModal(True)

        if user_data:
            self.setWindowTitle('Редактировать пользователя')
        else:
            self.setWindowTitle('Добавить пользователя')

        self.setFixedSize(500, 400)
        self.set_application_icon()
        self.init_ui()

    def set_application_icon(self):
        """Установка иконки приложения в зависимости от платформы"""
        try:
            if sys.platform == "darwin":  # macOS
                # Для macOS можно попробовать .png или .icns
                icon_path = resource_path('icons/icon.png')
            elif sys.platform == "win32":  # Windows
                icon_path = resource_path('icons/icon.ico')
            else:  # Linux
                icon_path = resource_path('icons/icon.png')

            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
                # Также установить иконку для всего приложения
                app = QApplication.instance()
                if app:
                    app.setWindowIcon(QIcon(icon_path))
            else:
                print(f"Иконка не найдена: {icon_path}")

        except Exception as e:
            print(f"Ошибка при установке иконки: {e}")

    def init_ui(self):
        layout = QVBoxLayout()

        form_layout = QFormLayout()

        # Поля формы
        self.username = QLineEdit()
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.EchoMode.Password)
        self.full_name = QLineEdit()

        self.role = QComboBox()
        self.role.addItems(['admin', 'user'])

        self.course = QComboBox()
        self.course.addItems(['', '1 курс', '2 курс', '3 курс', '4 курс', '5 курс'])

        # Добавление полей в форму
        form_layout.addRow('Логин:', self.username)
        form_layout.addRow('Пароль:', self.password)
        form_layout.addRow('ФИО:', self.full_name)
        form_layout.addRow('Роль:', self.role)
        form_layout.addRow('Курс (опционально):', self.course)
        # form_layout.addRow('Факультет (опционально):', self.faculty)

        # Заполнение данных если редактирование
        if self.user_data:
            self.username.setText(self.user_data.get('username', ''))
            self.password.setText(self.user_data.get('password', ''))
            self.full_name.setText(self.user_data.get('full_name', ''))
            self.role.setCurrentText(self.user_data.get('role', 'user'))
            self.course.setCurrentText(self.user_data.get('course', ''))
            # self.faculty.setText(self.user_data.get('faculty', ''))

        layout.addLayout(form_layout)

        # Кнопки
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)

    def get_data(self):
        """Получение данных из формы"""
        return {
            'username': self.username.text().strip(),
            'password': self.password.text().strip(),
            'full_name': self.full_name.text().strip(),
            'role': self.role.currentText(),
            'course': self.course.currentText() if self.course.currentText() != '' else None,
            # 'faculty': self.faculty.text().strip() if self.faculty.text().strip() != '' else None
        }


class PermissionDialog(QDialog):
    """Диалог добавления права доступа"""

    def __init__(self, user_id, user_name, parent=None):
        super().__init__(parent)
        self.user_id = user_id
        self.user_name = user_name
        self.setModal(True)
        self.setWindowTitle(f'Добавить права для {user_name}')
        self.setFixedSize(400, 200)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        form_layout = QFormLayout()

        # Поля формы
        self.permission_type = QComboBox()
        self.permission_type.addItems(['all', 'course'])
        self.permission_type.currentTextChanged.connect(self.update_fields)

        self.course = QComboBox()
        self.course.addItems(['1 курс', '2 курс', '3 курс', '4 курс', '5 курс'])

        # self.faculty = QLineEdit()
        # self.faculty.setPlaceholderText('Введите факультет')

        # Скрываем по умолчанию
        self.course.setVisible(False)
        # self.faculty.setVisible(False)

        # Добавление полей в форму
        form_layout.addRow('Тип права:', self.permission_type)
        form_layout.addRow('Курс:', self.course)
        # form_layout.addRow('Факультет:', self.faculty)

        layout.addLayout(form_layout)

        # Кнопки
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)

    def update_fields(self, permission_type):
        """Обновление видимости полей в зависимости от типа права"""
        self.course.setVisible(permission_type == 'course')

    def get_data(self):
        """Получение данных из формы"""
        permission_type = self.permission_type.currentText()

        data = {
            'user_id': self.user_id,
            'permission_type': permission_type,
            'course': None,
            'faculty': None
        }

        if permission_type == 'course':
            data['course'] = self.course.currentText()
        elif permission_type == 'faculty':
            data['faculty'] = self.faculty.text().strip()

        return data


class ImportDialog(QDialog):
    """Диалог для импорта данных с паролем и выбором листов"""

    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.setModal(True)
        self.setWindowTitle('Импорт данных из Excel')
        self.setFixedSize(500, 400)
        self.sheet_checkboxes = []
        self.available_sheets = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Информация о файле
        file_info = QLabel(f"Файл: {os.path.basename(self.file_path)}")
        file_info.setStyleSheet("font-weight: bold; color: #2c3e50;")
        layout.addWidget(file_info)

        # Пароль для Excel файла (если требуется)
        excel_password_layout = QHBoxLayout()
        excel_password_label = QLabel("Пароль Excel файла (если требуется):")
        self.excel_password_input = QLineEdit()
        self.excel_password_input.setPlaceholderText('Оставьте пустым, если файл не защищен')
        self.excel_password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.excel_password_input.textChanged.connect(self.on_password_changed)
        excel_password_layout.addWidget(excel_password_label)
        excel_password_layout.addWidget(self.excel_password_input)
        layout.addLayout(excel_password_layout)

        # Кнопка для загрузки листов
        self.load_sheets_btn = QPushButton("📄 Загрузить список листов")
        self.load_sheets_btn.clicked.connect(self.load_sheets)
        self.load_sheets_btn.setEnabled(True)
        layout.addWidget(self.load_sheets_btn)

        # Область для отображения листов
        sheets_label = QLabel("Выберите листы для импорта:")
        sheets_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        self.sheets_label = sheets_label
        self.sheets_label.setVisible(False)
        layout.addWidget(self.sheets_label)

        self.sheet_container = QWidget()
        self.sheet_layout = QVBoxLayout(self.sheet_container)

        # Добавляем прокрутку
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.sheet_container)
        self.scroll_area.setMaximumHeight(150)
        self.scroll_area.setVisible(False)
        layout.addWidget(self.scroll_area)

        # Виджет с кнопками выбора всех/очистки
        self.buttons_widget = QWidget()
        self.buttons_layout = QHBoxLayout(self.buttons_widget)
        self.select_all_btn = QPushButton('Выбрать все')
        self.select_all_btn.clicked.connect(self.select_all_sheets)
        self.clear_all_btn = QPushButton('Очистить все')
        self.clear_all_btn.clicked.connect(self.clear_all_sheets)
        self.buttons_layout.addWidget(self.select_all_btn)
        self.buttons_layout.addWidget(self.clear_all_btn)
        self.buttons_layout.addStretch()
        self.buttons_widget.setVisible(False)
        layout.addWidget(self.buttons_widget)

        # Информация
        info_label = QLabel("""
        ⚠️ При импорте:
        • Проверяются дубликаты (по ФИО и телефону)
        • Существующие данные не будут перезаписаны
        • Пустые строки игнорируются
        """)
        info_label.setStyleSheet("color: #e67e22; font-size: 12px; margin-top: 10px;")
        layout.addWidget(info_label)

        layout.addStretch()

        # Кнопки
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        self.ok_button = button_box.button(QDialogButtonBox.StandardButton.Ok)
        self.ok_button.setEnabled(False)
        layout.addWidget(button_box)

        self.setLayout(layout)

    def on_password_changed(self, text):
        """При изменении пароля сбрасываем выбор листов"""
        self.reset_sheet_selection()

    def reset_sheet_selection(self):
        """Сброс выбора листов"""
        while self.sheet_layout.count():
            item = self.sheet_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        self.sheet_checkboxes = []
        self.available_sheets = []
        self.sheets_label.setVisible(False)
        self.scroll_area.setVisible(False)
        self.buttons_widget.setVisible(False)
        self.ok_button.setEnabled(False)

    def load_sheets(self):
        """Загрузка списка листов из файла с использованием пароля"""
        # Очищаем предыдущие чекбоксы
        self.reset_sheet_selection()

        password = self.excel_password_input.text().strip()
        file_path = self.file_path

        xls = None  # Объявляем переменную заранее
        temp_file_path = None

        try:
            # Если указан пароль, пробуем расшифровать
            if password:
                try:
                    import msoffcrypto
                    import tempfile

                    with open(file_path, "rb") as f:
                        office_file = msoffcrypto.OfficeFile(f)
                        office_file.load_key(password=password)

                        temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
                        temp_file_path = temp_file.name
                        office_file.decrypt(temp_file)
                        temp_file.close()

                        file_path = temp_file_path

                except Exception as e:
                    QMessageBox.warning(self, 'Ошибка пароля',
                                        'Неверный пароль или файл не защищен паролем. '
                                        'Попробуйте снова или оставьте поле пустым.')
                    return

            try:
                # Пробуем разные движки для чтения Excel
                try:
                    xls = pd.ExcelFile(file_path, engine='openpyxl')
                except Exception as e1:
                    try:
                        xls = pd.ExcelFile(file_path, engine='xlrd')
                    except Exception as e2:
                        # Пробуем без указания движка
                        try:
                            xls = pd.ExcelFile(file_path)
                        except Exception as e3:
                            QMessageBox.critical(self, 'Ошибка чтения файла',
                                                 f'Не удалось прочитать файл Excel.\n'
                                                 f'Ошибки:\n'
                                                 f'Openpyxl: {str(e1)}\n'
                                                 f'Xlrd: {str(e2)}\n'
                                                 f'Automatic: {str(e3)}')
                            return

                self.available_sheets = xls.sheet_names

                # Создаем чекбоксы для каждого листа
                for sheet in self.available_sheets:
                    checkbox = QCheckBox(sheet)
                    # Автоматически выбираем листы с "курс"
                    if 'курс' in sheet.lower():
                        checkbox.setChecked(True)
                    self.sheet_checkboxes.append(checkbox)
                    self.sheet_layout.addWidget(checkbox)

                # Показываем элементы
                self.sheets_label.setVisible(True)
                self.scroll_area.setVisible(True)
                self.buttons_widget.setVisible(True)
                self.ok_button.setEnabled(any(cb.isChecked() for cb in self.sheet_checkboxes))

                # Обновляем состояние OK кнопки при изменении чекбоксов
                for checkbox in self.sheet_checkboxes:
                    checkbox.stateChanged.connect(self.update_ok_button)

            except Exception as e:
                QMessageBox.critical(self, 'Ошибка', f'Ошибка загрузки листов: {str(e)}')

            finally:
                # ВАЖНО: закрываем файл
                if xls is not None:
                    try:
                        xls.close()
                    except:
                        pass

                # Удаляем временный файл
                if temp_file_path and os.path.exists(temp_file_path):
                    try:
                        import time
                        time.sleep(0.1)  # Даем время системе освободить файл

                        for attempt in range(3):
                            try:
                                os.unlink(temp_file_path)
                                break
                            except PermissionError:
                                time.sleep(0.1)
                                continue
                    except Exception as e:
                        print(f"Ошибка удаления временного файла в диалоге: {e}")

        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', f'Ошибка загрузки файла: {str(e)}')

    def update_ok_button(self):
        """Обновление состояния кнопки OK в зависимости от выбора листов"""
        has_selected = any(cb.isChecked() for cb in self.sheet_checkboxes)
        self.ok_button.setEnabled(has_selected)

    def select_all_sheets(self):
        """Выбрать все листы"""
        for checkbox in self.sheet_checkboxes:
            checkbox.setChecked(True)
        self.update_ok_button()

    def clear_all_sheets(self):
        """Очистить все выборы"""
        for checkbox in self.sheet_checkboxes:
            checkbox.setChecked(False)
        self.update_ok_button()

    def get_excel_password(self):
        return self.excel_password_input.text().strip()

    def get_selected_sheets(self):
        """Получить список выбранных листов"""
        return [checkbox.text() for checkbox in self.sheet_checkboxes if checkbox.isChecked()]


class MainWindow(QMainWindow):
    """Главное окно приложения"""

    # Сигнал для выхода из системы
    logout_requested = pyqtSignal()

    def __init__(self, user_data):
        super().__init__()
        icon_path = get_icon_path("icon.ico")
        self.advanced_filters = {}
        if icon_path and os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.user_data = user_data
        self.db = Database()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(f'Агитация факультета 2026 - {self.user_data["full_name"]}')
        self.setGeometry(100, 100, 1200, 700)

        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Панель инструментов
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        # Действия
        add_action = QAction(QIcon(resource_path("icons/add_document.png")), 'Добавить', self)
        add_action.triggered.connect(self.add_applicant)
        toolbar.addAction(add_action)

        edit_action = QAction(QIcon(resource_path("icons/pencel.png")), 'Редактировать', self)
        edit_action.triggered.connect(self.edit_applicant)
        toolbar.addAction(edit_action)

        delete_action = QAction(QIcon(resource_path("icons/delete_document.png")), '️Удалить', self)
        delete_action.triggered.connect(self.delete_applicant)
        toolbar.addAction(delete_action)

        toolbar.addSeparator()

        # Кнопка выхода
        logout_action = QAction(QIcon(resource_path("icons/logout.png")), 'Выход', self)
        logout_action.triggered.connect(self.logout)
        toolbar.addAction(logout_action)

        import_action = QAction(QIcon(resource_path("icons/import.png")), 'Импорт из Excel', self)
        import_action.triggered.connect(self.import_from_excel)
        toolbar.addAction(import_action)

        export_action = QAction(QIcon(resource_path("icons/export.png")), 'Экспорт', self)
        export_action.triggered.connect(self.export_data)
        toolbar.addAction(export_action)

        # Вкладки
        self.tab_widget = QTabWidget()

        # Вкладка с данными
        self.data_tab = QWidget()
        self.init_data_tab()
        self.tab_widget.addTab(self.data_tab, QIcon(resource_path("icons/information.png")), 'Данные абитуриентов')

        # Вкладка со статистикой
        self.stats_tab = StatisticsWidget(
            self.user_data['id'],
            self.user_data['role'],
            self.db
        )
        self.tab_widget.addTab(self.stats_tab, QIcon(resource_path("icons/stata.png")), 'Статистика')

        # Вкладка настроек (только для админа)
        if self.user_data['role'] == 'admin':
            self.settings_tab = QWidget()
            self.init_settings_tab()
            self.tab_widget.addTab(self.settings_tab, QIcon(resource_path("icons/settings.png")), 'Настройки (админ)')

        main_layout.addWidget(self.tab_widget)

        # Статус бар
        status_bar = QStatusBar()
        self.setStatusBar(status_bar)
        status_bar.showMessage(f'Пользователь: {self.user_data["full_name"]} | Роль: {self.user_data["role"]}')

        # Обновление данных
        self.refresh_data()

    def open_advanced_search(self):
        """Открытие диалога расширенного поиска"""
        dialog = AdvancedSearchDialog(self.db, self)

        # Если есть сохраненные фильтры, загружаем их (опционально)
        if self.advanced_filters:
            # Здесь можно восстановить предыдущие фильтры
            pass

        if dialog.exec():
            self.advanced_filters = dialog.get_filters()
            self.refresh_data()

            # Показываем информацию о примененных фильтрах
            if self.advanced_filters:
                filter_count = len(self.advanced_filters)
                self.statusBar().showMessage(f"🔍 Применено расширенных фильтров: {filter_count}", 3000)
                # Подсвечиваем кнопку
                self.advanced_search_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #8e44ad;
                        color: white;
                        border: 2px solid #f39c12;
                        border-radius: 5px;
                        padding: 8px 16px;
                        font-weight: bold;
                    }
                    QPushButton:hover {
                        background-color: #7d3c98;
                    }
                """)
            else:
                self.reset_filters_style()

    def reset_all_filters(self):
        """Сброс всех фильтров"""
        # Сбрасываем обычный поиск
        self.search_input.clear()

        # Сбрасываем расширенные фильтры
        self.advanced_filters = {}

        # Сбрасываем фильтр по курсу (если есть)
        if hasattr(self, 'course_filter'):
            self.course_filter.setCurrentIndex(0)

        # Обновляем таблицу
        self.refresh_data()

        # Сбрасываем стиль кнопки
        self.reset_filters_style()

        self.statusBar().showMessage("Все фильтры сброшены", 2000)

    def reset_filters_style(self):
        """Сброс стиля кнопки расширенного поиска"""
        self.advanced_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)

    def logout(self):
        """Выход из системы"""
        reply = QMessageBox.question(
            self, 'Подтверждение выхода',
            'Вы уверены, что хотите выйти из системы?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.db.close()
            self.logout_requested.emit()  # Отправляем сигнал
            self.close()

    def init_data_tab(self):
        """Инициализация вкладки с данными (обновленная версия с панелью выделения)"""
        layout = QVBoxLayout()

        # Панель фильтров
        filter_widget = QWidget()
        filter_layout = QHBoxLayout()

        # Фильтр по курсу (только для админа)
        if self.user_data['role'] == 'admin':
            filter_label = QLabel('Фильтр по курсу:')
            self.course_filter = QComboBox()
            self.course_filter.addItems(['Все курсы', '1 курс', '2 курс', '3 курс', '4 курс', '5 курс'])
            self.course_filter.currentTextChanged.connect(self.refresh_data)

            filter_layout.addWidget(filter_label)
            filter_layout.addWidget(self.course_filter)

        filter_layout.addStretch()

        # Обычный поиск
        search_label = QLabel('Поиск:')
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Введите текст для поиска...')
        self.search_input.textChanged.connect(self.refresh_data)
        self.search_input.setMinimumWidth(250)

        # Кнопка расширенного поиска
        self.advanced_search_btn = QPushButton("🔍 Расширенный поиск")
        self.advanced_search_btn.setStyleSheet("""
                QPushButton {
                    background-color: #9b59b6;
                    color: white;
                    border: none;
                    border-radius: 5px;
                    padding: 8px 16px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #8e44ad;
                }
            """)
        self.advanced_search_btn.clicked.connect(self.open_advanced_search)

        # Кнопка сброса фильтров
        self.reset_filters_btn = QPushButton("✖️ Сбросить фильтры")
        self.reset_filters_btn.setStyleSheet("""
                QPushButton {
                    background-color: #e67e22;
                    color: white;
                    border: none;
                    border-radius: 5px;
                    padding: 8px 16px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #d35400;
                }
            """)
        self.reset_filters_btn.clicked.connect(self.reset_all_filters)

        filter_layout.addWidget(search_label)
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.advanced_search_btn)
        filter_layout.addWidget(self.reset_filters_btn)

        filter_widget.setLayout(filter_layout)
        layout.addWidget(filter_widget)

        self.table = QTableWidget()
        self.table.setColumnCount(11)  # 10 колонок + ID скрытая
        self.table.setHorizontalHeaderLabels([
            "ID",  # скрытая
            "ФИО абитуриента",
            "Субъект РФ",
            "Населенный пункт",
            "Категория",
            "Телефон",
            "Образование",
            "Статус",
            "Документы",
            "Подразделение агитатора",
            "ФИО агитатора"
        ])

        # Настройка таблицы
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)

        # Скрываем колонку ID
        self.table.setColumnHidden(0, True)

        # Устанавливаем ширину колонок
        self.table.setColumnWidth(1, 250)  # ФИО абитуриента
        self.table.setColumnWidth(2, 150)  # Субъект РФ
        self.table.setColumnWidth(3, 150)  # Населенный пункт
        self.table.setColumnWidth(4, 100)  # Категория
        self.table.setColumnWidth(5, 120)  # Телефон
        self.table.setColumnWidth(6, 100)  # Образование
        self.table.setColumnWidth(7, 100)  # Статус
        self.table.setColumnWidth(8, 100)  # Документы
        self.table.setColumnWidth(9, 150)  # Подразделение агитатора
        self.table.setColumnWidth(10, 200)  # ФИО агитатора

        layout.addWidget(self.table)

        # Панель статуса выделения
        self.selection_status_widget = QWidget()
        self.selection_status_widget.setStyleSheet("""
            QWidget {
                background-color: #e8f4fc;
                border-radius: 6px;
                margin-top: 5px;
            }
        """)
        selection_layout = QHBoxLayout(self.selection_status_widget)
        selection_layout.setContentsMargins(15, 10, 15, 10)

        self.selection_info_label = QLabel("Выделено: 0 записей")
        self.selection_info_label.setStyleSheet("font-weight: bold; color: #2c3e50;")

        self.show_stats_btn = QPushButton("📊 Показать статистику по выделенным")
        self.show_stats_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 6px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        self.show_stats_btn.clicked.connect(self.show_selection_stats)
        self.show_stats_btn.setEnabled(False)

        self.clear_selection_btn = QPushButton("✖️ Снять выделение")
        self.clear_selection_btn.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 6px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
        """)
        self.clear_selection_btn.clicked.connect(self.clear_selection)

        selection_layout.addWidget(self.selection_info_label)
        selection_layout.addStretch()
        selection_layout.addWidget(self.show_stats_btn)
        selection_layout.addWidget(self.clear_selection_btn)

        layout.addWidget(self.selection_status_widget)

        self.data_tab.setLayout(layout)

    def on_selection_changed(self):
        """Обработка изменения выделения в таблице"""
        selected_rows = set()
        for item in self.table.selectedItems():
            selected_rows.add(item.row())

        count = len(selected_rows)
        self.selection_info_label.setText(f"📌 Выделено: {count} записей")
        self.show_stats_btn.setEnabled(count > 0)

        # Обновляем статусную строку
        if count > 0:
            self.statusBar().showMessage(f"Выделено {count} записей. Нажмите 'Показать статистику' для анализа.")
        else:
            self.statusBar().showMessage(
                f"Пользователь: {self.user_data['full_name']} | Роль: {self.user_data['role']}")

    def clear_selection(self):
        """Снять все выделения"""
        self.table.clearSelection()
        self.on_selection_changed()

    def show_selection_stats(self):
        """Показать статистику по выделенным строкам"""
        selected_rows = set()
        for item in self.table.selectedItems():
            selected_rows.add(item.row())

        if not selected_rows:
            QMessageBox.warning(self, "Внимание", "Нет выделенных записей!")
            return

        # Собираем данные выделенных строк
        selected_data = []
        for row in selected_rows:
            row_data = {}

            # Получаем ID записи (первая колонка)
            id_item = self.table.item(row, 0)
            if id_item and id_item.text().strip():
                try:
                    row_data['id'] = int(id_item.text())
                except ValueError:
                    row_data['id'] = None
            else:
                row_data['id'] = None

            # Получаем остальные данные
            columns = ['study_group', 'rank', 'student_name', 'region', 'city',
                       'category', 'applicant_name', 'phone', 'status', 'document_status']

            # Фактические индексы колонок в таблице (пропускаем ID)
            for i, col_name in enumerate(columns, start=1):
                item = self.table.item(row, i)
                if item and item.text().strip():
                    row_data[col_name] = item.text()
                else:
                    row_data[col_name] = ''

            # Добавляем курс (нужно получить из БД или из другой колонки)
            # В текущей таблице нет курса, поэтому получаем из БД
            if row_data.get('id'):
                try:
                    cursor = self.db.conn.cursor()
                    cursor.execute('SELECT course FROM applicants WHERE id = ?', (row_data['id'],))
                    result = cursor.fetchone()
                    if result:
                        row_data['course'] = result['course']
                    else:
                        row_data['course'] = ''
                except Exception:
                    row_data['course'] = ''
            else:
                row_data['course'] = ''

            selected_data.append(row_data)

        if not selected_data:
            QMessageBox.warning(self, "Внимание", "Не удалось получить данные для выделенных записей!")
            return

        # Показываем диалог со статистикой
        stats_dialog = SelectionStatsDialog(selected_data, self)
        stats_dialog.show()

    def refresh_data(self):
        """Обновление данных в таблице"""
        # Сохраняем выделенные ID
        selected_ids = set()
        for item in self.table.selectedItems():
            row = item.row()
            id_item = self.table.item(row, 0)
            if id_item:
                selected_ids.add(int(id_item.text()))

        # Получение данных из БД с учетом расширенных фильтров
        if self.user_data['role'] == 'admin':
            course = self.course_filter.currentText() if hasattr(self, 'course_filter') else None
            applicants = self.db.get_applicants(
                self.user_data['id'],
                'admin',
                course if course != 'Все курсы' else None,
                self.advanced_filters if self.advanced_filters else None
            )
        else:
            applicants = self.db.get_applicants(
                self.user_data['id'],
                self.user_data['role'],
                None,
                self.advanced_filters if self.advanced_filters else None
            )

        # Обычный поиск (дополнительно к расширенному)
        search_text = self.search_input.text().lower().strip()
        if search_text and not self.advanced_filters:
            # Если нет расширенных фильтров, применяем обычный поиск
            filtered_applicants = []
            for applicant in applicants:
                applicant_dict = dict(applicant)
                text_fields = [
                    str(applicant_dict.get('applicant_name', '')),
                    str(applicant_dict.get('region', '')),
                    str(applicant_dict.get('city', '')),
                    str(applicant_dict.get('phone', '')),
                    str(applicant_dict.get('agitator_name', '')),
                    str(applicant_dict.get('agitator_department', '')),
                ]
                if any(search_text in field.lower() for field in text_fields):
                    filtered_applicants.append(applicant)
            applicants = filtered_applicants
        elif search_text and self.advanced_filters:
            # Если есть расширенные фильтры, обычный поиск работает поверх них
            filtered_applicants = []
            for applicant in applicants:
                applicant_dict = dict(applicant)
                text_fields = [
                    str(applicant_dict.get('applicant_name', '')),
                    str(applicant_dict.get('region', '')),
                    str(applicant_dict.get('city', '')),
                    str(applicant_dict.get('phone', '')),
                    str(applicant_dict.get('agitator_name', '')),
                ]
                if any(search_text in field.lower() for field in text_fields):
                    filtered_applicants.append(applicant)
            applicants = filtered_applicants

        self.table.blockSignals(True)
        self.table.setRowCount(len(applicants))

        id_to_row = {}

        # Категории для отображения
        category_map = {'м': 'Мужчина', 'ж': 'Женщина', 'всл': 'Военнослужащий'}

        for row, applicant in enumerate(applicants):
            applicant_dict = dict(applicant)
            applicant_id = applicant_dict.get('id')
            id_to_row[applicant_id] = row

            phone = applicant_dict.get('phone', '')
            formatted_phone = self.format_phone_number(phone)

            # Отображение категории
            category = applicant_dict.get('category', '')
            category_display = category_map.get(category, category)

            # Статус
            status = applicant_dict.get('status', '')
            status_display = '✅ Поступает' if status == 'поступает' else '❌ Отказывается'

            items = [
                QTableWidgetItem(str(applicant_id)),  # ID скрытый
                QTableWidgetItem(applicant_dict.get('applicant_name', '')),
                QTableWidgetItem(applicant_dict.get('region', '')),
                QTableWidgetItem(applicant_dict.get('city', '')),
                QTableWidgetItem(category_display),
                QTableWidgetItem(formatted_phone),
                QTableWidgetItem(applicant_dict.get('education', '')),
                QTableWidgetItem(status_display),
                QTableWidgetItem(applicant_dict.get('document_status', '')),
                QTableWidgetItem(applicant_dict.get('agitator_department', '')),
                QTableWidgetItem(applicant_dict.get('agitator_name', '')),
            ]

            # Цветовая индикация
            if status == 'поступает':
                items[7].setBackground(QColor(230, 255, 230))
                items[7].setForeground(QColor(0, 100, 0))
            else:
                items[7].setBackground(QColor(255, 230, 230))
                items[7].setForeground(QColor(150, 0, 0))

            # Цвет для категории
            if category == 'м':
                items[4].setBackground(QColor(230, 240, 255))
            elif category == 'ж':
                items[4].setBackground(QColor(255, 230, 240))
            elif category == 'всл':
                items[4].setBackground(QColor(230, 255, 230))

            # Подсветка пустых полей
            for col, item in enumerate(items):
                if not item.text().strip() and col not in [0, 4]:  # ID и категория могут быть пустыми
                    item.setBackground(QColor(255, 255, 200))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, col, item)

        # Восстанавливаем выделение
        for applicant_id in selected_ids:
            if applicant_id in id_to_row:
                self.table.selectRow(id_to_row[applicant_id])

        self.table.resizeColumnsToContents()
        self.table.blockSignals(False)
        self.on_selection_changed()

    def init_settings_tab(self):
        """Инициализация вкладки настроек (только для админа)"""
        layout = QVBoxLayout()

        # Вкладки внутри настроек
        self.admin_tabs = QTabWidget()

        # Вкладка пользователей
        self.users_tab = QWidget()
        self.init_users_tab()
        self.admin_tabs.addTab(self.users_tab, QIcon(resource_path("icons/users.png")), 'Пользователи')

        # Вкладка прав доступа
        self.permissions_tab = QWidget()
        self.init_permissions_tab()
        self.admin_tabs.addTab(self.permissions_tab, QIcon('icons/rules.png'), 'Права доступа')

        layout.addWidget(self.admin_tabs)
        self.settings_tab.setLayout(layout)

    def init_users_tab(self):
        """Инициализация вкладки пользователей"""
        layout = QVBoxLayout()

        # Панель управления
        controls_widget = QWidget()
        controls_layout = QHBoxLayout()

        # Кнопки
        self.add_user_btn = QPushButton(QIcon(resource_path("icons/add_user.png")), 'Добавить пользователя')
        self.add_user_btn.clicked.connect(self.add_user)

        self.edit_user_btn = QPushButton(QIcon(resource_path("icons/edit_user.png")), 'Редактировать пользователя')
        self.edit_user_btn.clicked.connect(self.edit_user)

        self.delete_user_btn = QPushButton(QIcon(resource_path("icons/delete_user.png")), 'Удалить пользователя')
        self.delete_user_btn.clicked.connect(self.delete_user)

        # Поиск
        self.search_user_input = QLineEdit()
        self.search_user_input.setPlaceholderText('Поиск пользователей...')
        self.search_user_input.textChanged.connect(self.refresh_users)
        self.search_user_input.setMinimumWidth(250)

        # Добавление виджетов
        controls_layout.addWidget(self.add_user_btn)
        controls_layout.addWidget(self.edit_user_btn)
        controls_layout.addWidget(self.delete_user_btn)
        controls_layout.addStretch()
        controls_layout.addWidget(self.search_user_input)

        controls_widget.setLayout(controls_layout)
        layout.addWidget(controls_widget)

        # Таблица пользователей
        self.users_table = QTableWidget()
        self.users_table.setColumnCount(4)
        self.users_table.setHorizontalHeaderLabels([
            'Логин', 'ФИО', 'Роль', 'Курс',
        ])

        # Настройка таблица
        self.users_table.horizontalHeader().setStretchLastSection(True)
        self.users_table.setAlternatingRowColors(True)
        self.users_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.users_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        layout.addWidget(self.users_table)

        self.users_tab.setLayout(layout)

        # Загрузка пользователей
        self.refresh_users()

    # Добавьте этот метод в класс MainWindow или как статический метод:

    def format_phone_number(self, phone):
        """Форматирование номера телефона в стандартный формат"""
        if not phone:
            return ""

        # Удаляем все нецифровые символы
        digits = ''.join(filter(str.isdigit, str(phone)))

        # Проверяем длину
        if len(digits) < 10:
            return phone  # Возвращаем как есть, если номер слишком короткий

        # Определяем код страны
        if digits.startswith('8') and len(digits) == 11:
            # Российский номер в формате 8XXXXXXXXXX
            digits = '7' + digits[1:]
        elif digits.startswith('7') and len(digits) == 11:
            pass  # Уже в правильном формате
        elif len(digits) == 10:
            # Номер без кода страны
            digits = '7' + digits

        # Форматируем в стандартный вид
        if len(digits) >= 11:
            return f"+7 ({digits[1:4]}) {digits[4:7]}-{digits[7:9]}-{digits[9:]}"

        return phone

    def init_permissions_tab(self):
        """Инициализация вкладки прав доступа"""
        layout = QVBoxLayout()

        # Панель управления
        controls_widget = QWidget()
        controls_layout = QHBoxLayout()

        # Выбор пользователя
        user_label = QLabel('Пользователь:')
        self.user_combo = QComboBox()
        self.user_combo.currentIndexChanged.connect(self.refresh_permissions)

        self.add_permission_btn = QPushButton('➕ Добавить права')
        self.add_permission_btn.clicked.connect(self.add_permission)

        # Добавление виджетов
        controls_layout.addWidget(user_label)
        controls_layout.addWidget(self.user_combo)
        controls_layout.addWidget(self.add_permission_btn)
        controls_layout.addStretch()

        controls_widget.setLayout(controls_layout)
        layout.addWidget(controls_widget)

        # Таблица прав доступа
        self.permissions_table = QTableWidget()
        self.permissions_table.setColumnCount(4)
        self.permissions_table.setHorizontalHeaderLabels([
            'Пользователь', 'Тип права', 'Курс'
        ])

        # Настройка таблицы
        self.permissions_table.horizontalHeader().setStretchLastSection(True)
        self.permissions_table.setAlternatingRowColors(True)
        self.permissions_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.permissions_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        layout.addWidget(self.permissions_table)

        self.permissions_tab.setLayout(layout)

        # Загрузка пользователей для комбобокса
        self.load_users_for_combo()

    def load_users_for_combo(self):
        """Загрузка пользователей в комбобокс"""
        self.user_combo.clear()
        self.user_combo.addItem('-- Выберите пользователя --', None)

        users = self.db.get_all_users()
        for user in users:
            user_dict = dict(user)
            if user_dict['role'] != 'admin':  # Не показываем администраторов
                self.user_combo.addItem(f"{user_dict['username']} ({user_dict['full_name']})", user_dict['id'])

    def refresh_users(self):
        """Обновление списка пользователей"""
        users = self.db.get_all_users()
        search_text = self.search_user_input.text().lower().strip()

        # Фильтрация по поиску
        if search_text:
            filtered_users = []
            for user in users:
                user_dict = dict(user)
                if (search_text in user_dict['username'].lower() or
                        search_text in user_dict['full_name'].lower()):
                    filtered_users.append(user)
            users = filtered_users

        # Заполнение таблицы
        self.users_table.setRowCount(len(users))

        for row, user in enumerate(users):
            user_dict = dict(user)

            items = [
                QTableWidgetItem(user_dict.get('username', '')),
                QTableWidgetItem(user_dict.get('full_name', '')),
                QTableWidgetItem(user_dict.get('role', '')),
                QTableWidgetItem(user_dict.get('course', '')),
                # QTableWidgetItem(user_dict.get('faculty', ''))
            ]

            for col, item in enumerate(items):
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.users_table.setItem(row, col, item)

        self.users_table.resizeColumnsToContents()

    def refresh_permissions(self):
        """Обновление списка прав доступа"""
        user_id = self.user_combo.currentData()

        if not user_id:
            self.permissions_table.setRowCount(0)
            return

        permissions = self.db.get_user_permissions(user_id)

        # Заполнение таблицы
        self.permissions_table.setRowCount(len(permissions))

        for row, permission in enumerate(permissions):
            perm_dict = dict(permission)

            items = [
                QTableWidgetItem(f"{perm_dict.get('full_name', '')}"),
                QTableWidgetItem(perm_dict.get('permission_type', '')),
                QTableWidgetItem(perm_dict.get('course', '')),
                # QTableWidgetItem(perm_dict.get('faculty', ''))
            ]

            for col, item in enumerate(items):
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.permissions_table.setItem(row, col, item)

        self.permissions_table.resizeColumnsToContents()

    def add_user(self):
        """Добавление нового пользователя"""
        dialog = UserDialog()
        if dialog.exec():
            data = dialog.get_data()

            # Проверка обязательных полей
            if not data['username'] or not data['password'] or not data['full_name']:
                QMessageBox.warning(self, 'Ошибка', 'Заполните все обязательные поля!')
                return

            user_id = self.db.add_user(
                data['username'], data['password'], data['full_name'],
                data['role'], data['course']
            )

            if user_id:
                QMessageBox.information(self, 'Успех', 'Пользователь успешно добавлен!')
                self.refresh_users()
                self.load_users_for_combo()
            else:
                QMessageBox.critical(self, 'Ошибка',
                                     'Не удалось добавить пользователя. Возможно, логин уже существует.')

    def edit_user(self):
        """Редактирование выбранного пользователя"""
        selected_rows = self.users_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Внимание', 'Выберите пользователя для редактирования!')
            return

        row = selected_rows[0].row()
        user_id = int(self.users_table.item(row, 0).text())

        # Получение данных пользователя из БД
        user_data = self.db.get_user_by_id(user_id)
        if not user_data:
            QMessageBox.critical(self, 'Ошибка', 'Пользователь не найден!')
            return

        user_dict = dict(user_data)
        dialog = UserDialog(user_dict, self)
        if dialog.exec():
            data = dialog.get_data()

            # Если пароль не изменен, оставляем старый
            if not data['password']:
                data['password'] = user_dict['password']

            success = self.db.update_user(user_id, data)

            if success:
                QMessageBox.information(self, 'Успех', 'Данные пользователя обновлены!')
                self.refresh_users()
                self.load_users_for_combo()
            else:
                QMessageBox.critical(self, 'Ошибка', 'Не удалось обновить данные пользователя.')

    def delete_user(self):
        """Удаление выбранного пользователя"""
        selected_rows = self.users_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Внимание', 'Выберите пользователя для удаления!')
            return

        row = selected_rows[0].row()
        user_id = int(self.users_table.item(row, 0).text())
        username = self.users_table.item(row, 1).text()

        # Нельзя удалить самого себя
        if user_id == self.user_data['id']:
            QMessageBox.warning(self, 'Ошибка', 'Вы не можете удалить самого себя!')
            return

        reply = QMessageBox.question(
            self, 'Подтверждение',
            f'Вы уверены, что хотите удалить пользователя "{username}"?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            success = self.db.delete_user(user_id)

            if success:
                QMessageBox.information(self, 'Успех', 'Пользователь успешно удален!')
                self.refresh_users()
                self.load_users_for_combo()
            else:
                QMessageBox.critical(self, 'Ошибка', 'Не удалось удалить пользователя.')

    def add_permission(self):
        """Добавление нового права доступа"""
        user_id = self.user_combo.currentData()

        if not user_id:
            QMessageBox.warning(self, 'Внимание', 'Выберите пользователя!')
            return

        user_name = self.user_combo.currentText()
        dialog = PermissionDialog(user_id, user_name, self)

        if dialog.exec():
            data = dialog.get_data()
            permission_id = self.db.add_permission(
                data['user_id'], data['permission_type'], None,
                data['course']
            )

            if permission_id:
                QMessageBox.information(self, 'Успех', 'Права доступа успешно добавлено!')
                self.refresh_permissions()
            else:
                QMessageBox.critical(self, 'Ошибка', 'Не удалось добавить права доступа.')

    # В классе MainWindow, обновляем методы add_applicant и edit_applicant:

    def add_applicant(self):
        """Добавление нового абитуриента"""
        dialog = ApplicantDialog(
            applicant_data=None,
            user_role=self.user_data['role'],
            db=self.db,
            parent=self
        )
        if dialog.exec():
            data = dialog.get_data()

            # Проверка обязательных полей (дублируем для безопасности)
            if not data['applicant_name']:
                QMessageBox.warning(self, 'Ошибка', 'ФИО абитуриента обязательно!')
                return

            if not data['agitator_name']:
                QMessageBox.warning(self, 'Ошибка', 'ФИО агитатора обязательно!')
                return

            # Добавляем в БД
            self.db.add_applicant(self.user_data['id'], data)

            # Обновляем таблицу и статистику
            self.refresh_data()
            self.stats_tab.update_statistics()

            QMessageBox.information(self, 'Успех', 'Абитуриент успешно добавлен!')

    def edit_applicant(self):
        """Редактирование выбранного абитуриента"""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Внимание', 'Выберите запись для редактирования!')
            return

        row = selected_rows[0].row()
        applicant_id = int(self.table.item(row, 0).text())

        # Получение данных абитуриента из БД
        cursor = self.db.conn.cursor()
        cursor.execute('SELECT * FROM applicants WHERE id = ?', (applicant_id,))
        applicant_data = dict(cursor.fetchone())

        dialog = ApplicantDialog(
            applicant_data=applicant_data,
            user_role=self.user_data['role'],
            db=self.db,
            parent=self
        )
        if dialog.exec():
            data = dialog.get_data()
            self.db.update_applicant(applicant_id, data)
            self.refresh_data()
            self.stats_tab.update_statistics()
            QMessageBox.information(self, 'Успех', 'Данные абитуриента обновлены!')

    def delete_applicant(self):
        """Удаление выбранного абитуриента"""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Внимание', 'Выберите запись для удаления!')
            return

        reply = QMessageBox.question(
            self, 'Подтверждение',
            'Вы уверены, что хотите удалить выбранную запись?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            for index in selected_rows:
                row = index.row()
                applicant_id = int(self.table.item(row, 0).text())
                self.db.delete_applicant(applicant_id)

            self.refresh_data()
            self.stats_tab.update_statistics()

    def import_from_excel(self):
        """Импорт данных из Excel файла"""
        try:
            # Выбор файла
            file_path, _ = QFileDialog.getOpenFileName(
                self, 'Выберите Excel файл', '',
                'Excel Files (*.xlsx *.xls *.xlsm);;All Files (*)'
            )

            if not file_path:
                return

            # Проверка расширения файла
            if not file_path.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                QMessageBox.warning(self, 'Ошибка', 'Пожалуйста, выберите файл Excel (.xlsx, .xls, .xlsm)')
                return

            # Показываем диалог
            dialog = ImportDialog(file_path, self)
            if dialog.exec():
                selected_sheets = dialog.get_selected_sheets()
                excel_password = dialog.get_excel_password()

                if not selected_sheets:
                    QMessageBox.warning(self, 'Ошибка', 'Выберите хотя бы один лист для импорта!')
                    return

                # Показываем прогресс
                progress = QProgressDialog("Импорт данных...", "Отмена", 0, 0, self)
                progress.setWindowTitle("Импорт")
                progress.setWindowModality(Qt.WindowModality.WindowModal)
                progress.show()

                # Обновляем интерфейс
                QApplication.processEvents()

                try:
                    # Импортируем данные
                    success, message = self.import_excel_data(
                        file_path,
                        excel_password,
                        selected_sheets
                    )

                    progress.close()

                    if success:
                        QMessageBox.information(self, 'Успех', message)
                        # Обновляем данные
                        self.refresh_data()
                        self.stats_tab.update_statistics()
                    else:
                        QMessageBox.critical(self, 'Ошибка импорта', message)

                except Exception as e:
                    progress.close()
                    QMessageBox.critical(self, 'Ошибка', f'Ошибка при импорте: {str(e)}')
                    import traceback
                    traceback.print_exc()

        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', f'Ошибка при импорте: {str(e)}')
            import traceback
            traceback.print_exc()

    @staticmethod
    def check_file_access(file_path):
        """Проверка доступности файла"""
        try:
            import time
            # Пробуем открыть файл несколько раз
            for attempt in range(5):
                try:
                    with open(file_path, 'rb') as f:
                        f.read(1)
                    return True
                except PermissionError:
                    time.sleep(0.1)
                    continue
            return False
        except Exception:
            return False

    def import_excel_data(self, file_path, excel_password, selected_sheets):
        """Импорт данных из Excel файла"""
        temp_file_path = None
        try:
            # Если указан пароль Excel, пробуем расшифровать
            if excel_password:
                try:
                    import msoffcrypto
                    import tempfile

                    with open(file_path, "rb") as f:
                        office_file = msoffcrypto.OfficeFile(f)
                        office_file.load_key(password=excel_password)

                        temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
                        temp_file_path = temp_file.name
                        office_file.decrypt(temp_file)
                        temp_file.close()

                        file_path = temp_file_path

                except Exception as e:
                    return False, f"Неверный пароль Excel файла или ошибка расшифровки: {str(e)}"

            try:
                # Определяем движок для чтения Excel
                engine = None
                if file_path.endswith('.xlsx'):
                    engine = 'openpyxl'
                elif file_path.endswith('.xls'):
                    engine = 'xlrd'  # Для старых .xls файлов
                else:
                    engine = 'openpyxl'  # по умолчанию

                # Пробуем прочитать файл
                try:
                    xls = pd.ExcelFile(file_path, engine=engine)
                except Exception as e:
                    # Если не получилось с выбранным движком, пробуем другой
                    if engine == 'openpyxl':
                        engine = 'xlrd'
                    else:
                        engine = 'openpyxl'

                    try:
                        xls = pd.ExcelFile(file_path, engine=engine)
                    except Exception as e2:
                        return False, f"Не удалось прочитать файл. Возможно, файл поврежден или имеет неподдерживаемый формат. Ошибки: {str(e)}, {str(e2)}"

                imported_count = 0
                duplicate_count = 0
                skipped_count = 0

                # Проверяем доступные листы
                available_sheets = xls.sheet_names

                for sheet in selected_sheets:
                    if sheet not in available_sheets:
                        return False, f"Лист '{sheet}' не найден в файле"

                for sheet_name in selected_sheets:
                    try:
                        # Читаем лист
                        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)

                        # Определяем курс из названия листа
                        course_from_sheet = '1 курс'
                        for course_num in ['1', '2', '3', '4', '5']:
                            if f'{course_num} курс' in sheet_name.lower():
                                course_from_sheet = f'{course_num} курс'
                                break
                            elif f'курс {course_num}' in sheet_name.lower():
                                course_from_sheet = f'{course_num} курс'
                                break

                        # Импортируем данные
                        for index, row in df.iterrows():
                            # Пропускаем пустые строки
                            if row.isnull().all():
                                continue

                            # Получаем данные
                            applicant_data = self.extract_applicant_data(row, df, course_from_sheet)

                            if applicant_data:
                                # Проверяем дубликат
                                if self.check_duplicate(applicant_data):
                                    duplicate_count += 1
                                    continue

                                # Добавляем в базу
                                try:
                                    self.db.add_applicant(self.user_data['id'], applicant_data)
                                    imported_count += 1
                                except Exception as e:
                                    skipped_count += 1
                                    print(f"Ошибка добавления абитуриента: {e}")
                            else:
                                skipped_count += 1

                    except Exception as e:
                        error_msg = f"Ошибка импорта листа {sheet_name}: {e}"
                        print(error_msg)
                        import traceback
                        traceback.print_exc()
                        # Продолжаем с другими листами
                        continue

                result_message = f"""
                Импорт завершен!
                Всего импортировано: {imported_count}
                Пропущено дубликатов: {duplicate_count}
                Пропущено других ошибок: {skipped_count}
                Листы: {', '.join(selected_sheets)}
                """

                return True, result_message

            except Exception as e:
                error_msg = f"Ошибка чтения Excel файла: {str(e)}"
                print(error_msg)
                import traceback
                traceback.print_exc()
                return False, error_msg

            finally:
                # ЗАКРЫВАЕМ ExcelFile если он был открыт
                if 'xls' in locals():
                    xls.close()  # Явно закрываем файл

                # Удаляем временный файл если он был создан
                if temp_file_path and os.path.exists(temp_file_path):
                    try:
                        # Даем системе время освободить файл
                        import time
                        time.sleep(0.1)

                        # Пробуем удалить файл несколько раз
                        for attempt in range(3):
                            try:
                                os.unlink(temp_file_path)
                                break
                            except PermissionError:
                                time.sleep(0.1)
                                continue
                            except Exception as e:
                                print(f"Ошибка удаления временного файла: {e}")
                                break
                    except Exception as e:
                        print(f"Ошибка при удалении временного файла: {e}")

        except Exception as e:
            error_msg = f"Общая ошибка импорта: {str(e)}"
            print(error_msg)
            import traceback
            traceback.print_exc()
            return False, error_msg

    def extract_applicant_data(self, row, df, course):
        """Извлечение данных абитуриента из строки Excel"""
        try:
            # Ищем ФИО абитуриента
            applicant_name = ''
            for col in df.columns:
                col_str = str(col)
                if 'абитуриент' in col_str.lower():
                    value = row[col]
                    if pd.notna(value):
                        applicant_name = str(value).strip()
                        break

            if not applicant_name:
                return None

            # Собираем данные
            applicant_data = {
                'study_group': self._get_cell_value(row, ['уч.гр.', 'Уч. группа', 'учебная группа', 'Учебная группа',
                                                          'Группа']),
                'rank': self._get_cell_value(row, ['В.зв', 'Звание', 'воинское звание', 'Воинское звание', 'звание'],
                                             'ряд.'),
                'student_name': self._get_cell_value(row, ['ФИО', 'ФИО студента', 'Студент', 'фио', 'Студент ФИО']),
                'region': self._get_cell_value(row,
                                               ['Субъект РФ', 'Регион', 'область', 'Субъект_РФ', 'Регион (область)']),
                'city': self._get_cell_value(row, ['Населённый пункт', 'Город', 'город', 'Населенный пункт',
                                                   'Город/населенный пункт']),
                'category': self._get_cell_value(row, ['категория', 'Категория', 'пол', 'Пол', 'Кат.'], 'муж'),
                'applicant_name': applicant_name,
                'phone': self.format_imported_phone(
                self._get_cell_value(row, ['Телефон', 'телефон', 'контактный телефон', 'Телефон абитуриента', 'Тел.'])),
                'status': self._get_cell_value(row, ['Статус', 'статус', 'Статус поступления', 'Статус абитуриента'],
                                               'поступает'),
                'document_status': self._get_cell_value(row, ['Состояние личного дела на поступление', 'Документы',
                                                              'документы', 'Статус документов']),
                'notes': self._get_cell_value(row, ['Примечание', 'примечание', 'комментарий', 'Комментарий', 'Прим.']),
                'course': course,
                # 'faculty': self._get_cell_value(row, ['Факультет', 'факультет', 'Факультет/отделение', 'Отделение'])
            }

            # Нормализация статуса
            status = applicant_data['status'].lower()
            if 'поступает' in status or status == '1' or status == '1)поступает' or '1)' in status:
                applicant_data['status'] = '1)поступает'
            elif 'не поступает' in status or status == '0' or status == '2)не поступает' or '2)' in status or 'не' in status:
                applicant_data['status'] = '2)не поступает'
            else:
                applicant_data['status'] = '1)поступает'

            # Нормализация статуса документов
            doc_status = applicant_data['document_status'].lower()
            if 'формируется' in doc_status or doc_status == '1' or doc_status == '1)формируется' or '1)' in doc_status:
                applicant_data['document_status'] = '1)Формируется в военкомате'
            elif 'отправлено' in doc_status or doc_status == '2' or doc_status == '2)отправлено' or '2)' in doc_status:
                applicant_data['document_status'] = '2)Отправлено в ВА ВКО'
            elif 'ва вко' in doc_status or doc_status == '3' or doc_status == '3)в ва вко' or '3)' in doc_status:
                applicant_data['document_status'] = '3)В ВА ВКО'
            else:
                applicant_data['document_status'] = ''

            return applicant_data

        except Exception as e:
            print(f"Ошибка извлечения данных: {e}")
            import traceback
            traceback.print_exc()
            return None

    @staticmethod
    def _get_cell_value(row, possible_columns, default=''):
        """Получение значения ячейки из возможных колонок"""
        for col in possible_columns:
            if col in row and pd.notna(row[col]):
                return str(row[col]).strip()
        return default

    @staticmethod
    def format_imported_phone(phone):
        """Форматирование номера телефона при импорте"""
        if not phone:
            return ""

        # Удаляем все нецифровые символы
        digits = ''.join(filter(str.isdigit, str(phone)))

        if not digits:
            return phone

        # Приводим к стандартному формату
        if digits.startswith('8') and len(digits) == 11:
            # 8XXXXXXXXXX -> 7XXXXXXXXXX
            return '7' + digits[1:]
        elif len(digits) == 10:
            # XXXXXXXXXX -> 7XXXXXXXXXX
            return '7' + digits
        elif digits.startswith('7') and len(digits) == 11:
            return digits

        return phone

    def check_duplicate(self, applicant_data):
        """Проверка на дубликат (все пользователи)"""
        cursor = self.db.conn.cursor()
        cursor.execute('''
            SELECT COUNT(*) FROM applicants 
            WHERE applicant_name = ?
            AND phone = ?
            AND course = ?
        ''', (applicant_data['applicant_name'],
              applicant_data['phone'],
              applicant_data['course']))
        count = cursor.fetchone()[0]
        return count > 0

    def export_data(self):
        """Экспорт данных"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, 'Сохранить данные',
            f'export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            'Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*)'
        )

        if file_path:
            # Здесь можно реализовать экспорт данных
            QMessageBox.information(self, 'Информация', 'Экспорт данных будет реализован в следующей версии')

    def closeEvent(self, event):
        """Обработка закрытия окна"""
        self.db.close()
        event.accept()


class Application:
    """Главный класс приложения"""

    def __init__(self):
        self.app = QApplication(sys.argv)
        self.set_styles()

        # Показываем окно авторизации
        self.show_login()

    def set_styles(self):
        """Установка стилей приложения"""
        self.app.setStyle('Fusion')

        # Кастомная палитра
        palette = self.app.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(50, 50, 50))
        palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(245, 245, 245))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor(50, 50, 50))
        palette.setColor(QPalette.ColorRole.Text, QColor(50, 50, 50))
        palette.setColor(QPalette.ColorRole.Button, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(50, 50, 50))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(41, 128, 185))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
        self.app.setPalette(palette)

        # Стили для QPushButton
        self.app.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #1c6ea4;
            }
            QLineEdit, QComboBox, QTextEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
                background-color: white;
            }
            QLineEdit:focus, QComboBox:focus, QTextEdit:focus {
                border: 1px solid #3498db;
            }
            QTableWidget {
                background-color: white;
                alternate-background-color: #f9f9f9;
                gridline-color: #ddd;
                border: 1px solid #ddd;
            }
            QHeaderView::section {
                background-color: #3498db;
                color: white;
                padding: 5px;
                border: 1px solid #2980b9;
                font-weight: bold;
            }
            QTabWidget::pane {
                border: 1px solid #ddd;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #ecf0f1;
                padding: 8px 16px;
                margin-right: 2px;
                border: 1px solid #ddd;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #3498db;
                color: white;
            }
            QTabBar::tab:hover {
                background-color: #bdc3c7;
            }
        """)

    def show_login(self):
        """Показать окно авторизации"""
        self.login_window = LoginWindow()
        self.login_window.login_success.connect(self.on_login_success)
        self.login_window.show()

    def on_login_success(self, user_data):
        """Обработка успешной авторизации"""
        self.login_window.hide()
        self.main_window = MainWindow(user_data)
        self.main_window.logout_requested.connect(self.on_logout)
        self.main_window.show()

    def on_logout(self):
        """Обработка выхода из системы"""
        if hasattr(self, 'main_window'):
            self.main_window.close()
            self.main_window = None

        self.show_login()

    def run(self):
        """Запуск приложения"""
        return self.app.exec()


if __name__ == '__main__':
    application = Application()
    sys.exit(application.run())