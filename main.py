# -*- coding: utf-8 -*-
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QGridLayout, QStackedWidget, QLabel,
                             QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
                             QHeaderView, QMessageBox, QComboBox, QFrame,
                             QGroupBox, QTabWidget, QDialog, QDialogButtonBox,
                             QFormLayout, QTextEdit, QDateEdit, QSpinBox,
                             QFileDialog, QToolBar, QStatusBar, QMenuBar, QMenu,
                             QScrollArea, QAction, QInputDialog, QCheckBox)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor
from datetime import datetime
from database import Database
from statistics_widget import StatisticsWidget
import os
os.environ['QT_MAC_WANTS_LAYER'] = '1'



class LoginWindow(QWidget):
    """Окно авторизации"""
    login_success = pyqtSignal(dict)  # Сигнал успешной авторизации

    def __init__(self):
        super().__init__()
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

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText('Введите пароль')
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setMinimumHeight(40)

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

    def __init__(self, applicant_data=None, parent=None):
        super().__init__(parent)
        self.applicant_data = applicant_data
        self.setModal(True)

        if applicant_data:
            self.setWindowTitle('Редактировать абитуриента')
        else:
            self.setWindowTitle('Добавить абитуриента')

        self.setFixedSize(500, 600)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        form_layout = QFormLayout()

        # Поля формы
        self.study_group = QLineEdit()
        self.rank = QComboBox()
        self.rank.addItems(['ряд.', 'ефр.', 'мл. серж.', 'серж.', 'ст. серж.', "пр-к"])

        self.student_name = QLineEdit()
        self.region = QLineEdit()
        self.city = QLineEdit()

        self.category = QComboBox()
        self.category.addItems(['муж', 'жен', 'в/сл'])

        self.applicant_name = QLineEdit()
        self.phone = QLineEdit()

        self.status = QComboBox()
        self.status.addItems(['поступает', 'не поступает'])

        self.document_status = QComboBox()
        self.document_status.addItems([
            '',
            'Формируется в военкомате',
            'Отправлено в ВА ВКО',
            'В ВА ВКО'
        ])

        self.notes = QTextEdit()
        self.notes.setMaximumHeight(100)

        self.course = QComboBox()
        self.course.addItems(['1 курс', '2 курс', '3 курс', '4 курс', '5 курс'])

        self.faculty = QLineEdit()

        # Добавление полей в форму
        form_layout.addRow('Учебная группа:', self.study_group)
        form_layout.addRow('Звание:', self.rank)
        form_layout.addRow('ФИО студента:', self.student_name)
        form_layout.addRow('Субъект РФ:', self.region)
        form_layout.addRow('Населенный пункт:', self.city)
        form_layout.addRow('Категория:', self.category)
        form_layout.addRow('ФИО абитуриента:', self.applicant_name)
        form_layout.addRow('Телефон:', self.phone)
        form_layout.addRow('Статус:', self.status)
        form_layout.addRow('Статус документов:', self.document_status)
        # form_layout.addRow('Факультет:', self.faculty)
        form_layout.addRow('Примечания:', self.notes)

        # Заполнение данных если редактирование
        if self.applicant_data:
            self.study_group.setText(self.applicant_data.get('study_group', ''))
            self.rank.setCurrentText(self.applicant_data.get('rank', 'ряд.'))
            self.student_name.setText(self.applicant_data.get('student_name', ''))
            self.region.setText(self.applicant_data.get('region', ''))
            self.city.setText(self.applicant_data.get('city', ''))
            self.category.setCurrentText(self.applicant_data.get('category', 'муж'))
            self.applicant_name.setText(self.applicant_data.get('applicant_name', ''))
            self.phone.setText(self.applicant_data.get('phone', ''))
            self.status.setCurrentText(self.applicant_data.get('status', 'поступает'))
            self.document_status.setCurrentText(self.applicant_data.get('document_status', ''))
            self.course.setCurrentText(self.applicant_data.get('course', '1 курс'))
            # self.faculty.setText(self.applicant_data.get('faculty', ''))
            self.notes.setText(self.applicant_data.get('notes', ''))

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
            'study_group': self.study_group.text().strip(),
            'rank': self.rank.currentText(),
            'student_name': self.student_name.text().strip(),
            'region': self.region.text().strip(),
            'city': self.city.text().strip(),
            'category': self.category.currentText(),
            'applicant_name': self.applicant_name.text().strip(),
            'phone': self.phone.text().strip(),
            'status': self.status.currentText(),
            'document_status': self.document_status.currentText(),
            'notes': self.notes.toPlainText().strip(),
            'course': self.course.currentText(),
            'faculty': self.faculty.text().strip()
        }


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
        self.init_ui()

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

        # self.faculty = QLineEdit()
        # self.faculty.setPlaceholderText('Введите факультет')

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
        self.permission_type.addItems(['all', 'course', 'faculty'])
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


class MainWindow(QMainWindow):
    """Главное окно приложения"""

    # Сигнал для выхода из системы
    logout_requested = pyqtSignal()

    def __init__(self, user_data):
        super().__init__()
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
        add_action = QAction('➕ Добавить', self)
        add_action.triggered.connect(self.add_applicant)
        toolbar.addAction(add_action)

        edit_action = QAction('✏ Редактировать', self)
        edit_action.triggered.connect(self.edit_applicant)
        toolbar.addAction(edit_action)

        delete_action = QAction('🗑️ Удалить', self)
        delete_action.triggered.connect(self.delete_applicant)
        toolbar.addAction(delete_action)

        toolbar.addSeparator()

        # Кнопка выхода
        logout_action = QAction('🚪 Выход', self)
        logout_action.triggered.connect(self.logout)
        toolbar.addAction(logout_action)

        import_action = QAction('📁 Импорт из Excel', self)
        import_action.triggered.connect(self.import_from_excel)
        toolbar.addAction(import_action)

        export_action = QAction('💾 Экспорт', self)
        export_action.triggered.connect(self.export_data)
        toolbar.addAction(export_action)

        # Вкладки
        self.tab_widget = QTabWidget()

        # Вкладка с данными
        self.data_tab = QWidget()
        self.init_data_tab()
        self.tab_widget.addTab(self.data_tab, '📋 Данные абитуриентов')

        # Вкладка со статистикой
        self.stats_tab = StatisticsWidget(
            self.user_data['id'],
            self.user_data['role'],
            self.db
        )
        self.tab_widget.addTab(self.stats_tab, '📊 Статистика')

        # Вкладка настроек (только для админа)
        if self.user_data['role'] == 'admin':
            self.settings_tab = QWidget()
            self.init_settings_tab()
            self.tab_widget.addTab(self.settings_tab, '⚙ Настройки (админ)')

        main_layout.addWidget(self.tab_widget)

        # Статус бар
        status_bar = QStatusBar()
        self.setStatusBar(status_bar)
        status_bar.showMessage(f'Пользователь: {self.user_data["full_name"]} | Роль: {self.user_data["role"]}')

        # Обновление данных
        self.refresh_data()

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
        """Инициализация вкладки с данными"""
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

        # Поиск
        search_label = QLabel('Поиск:')
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Введите текст для поиска...')
        self.search_input.textChanged.connect(self.refresh_data)
        self.search_input.setMinimumWidth(300)

        filter_layout.addWidget(search_label)
        filter_layout.addWidget(self.search_input)

        filter_widget.setLayout(filter_layout)
        layout.addWidget(filter_widget)

        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            'ID', 'Уч. группа', 'Звание', 'ФИО студента',
            'Регион', 'Город', 'Категория', 'ФИО абитуриента',
            'Телефон', 'Статус', 'Документы'
        ])

        # Настройка таблицы
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        layout.addWidget(self.table)

        self.data_tab.setLayout(layout)

    def init_settings_tab(self):
        """Инициализация вкладки настроек (только для админа)"""
        layout = QVBoxLayout()

        # Вкладки внутри настроек
        self.admin_tabs = QTabWidget()

        # Вкладка пользователей
        self.users_tab = QWidget()
        self.init_users_tab()
        self.admin_tabs.addTab(self.users_tab, '👥 Пользователи')

        # Вкладка прав доступа
        self.permissions_tab = QWidget()
        self.init_permissions_tab()
        self.admin_tabs.addTab(self.permissions_tab, '🔒 Права доступа')

        layout.addWidget(self.admin_tabs)
        self.settings_tab.setLayout(layout)

    def init_users_tab(self):
        """Инициализация вкладки пользователей"""
        layout = QVBoxLayout()

        # Панель управления
        controls_widget = QWidget()
        controls_layout = QHBoxLayout()

        # Кнопки
        self.add_user_btn = QPushButton('➕ Добавить пользователя')
        self.add_user_btn.clicked.connect(self.add_user)

        self.edit_user_btn = QPushButton('✏ Редактировать')
        self.edit_user_btn.clicked.connect(self.edit_user)

        self.delete_user_btn = QPushButton('🗑️ Удалить')
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
        self.users_table.setColumnCount(5)
        self.users_table.setHorizontalHeaderLabels([
            'ID', 'Логин', 'ФИО', 'Роль', 'Курс',
        ])

        # Настройка таблицы
        self.users_table.horizontalHeader().setStretchLastSection(True)
        self.users_table.setAlternatingRowColors(True)
        self.users_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.users_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        layout.addWidget(self.users_table)

        self.users_tab.setLayout(layout)

        # Загрузка пользователей
        self.refresh_users()

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
            'ID', 'Пользователь', 'Тип права', 'Курс'
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
                QTableWidgetItem(str(user_dict.get('id', ''))),
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
                QTableWidgetItem(str(perm_dict.get('id', ''))),
                QTableWidgetItem(f"{perm_dict.get('full_name', '')}"),
                QTableWidgetItem(perm_dict.get('permission_type', '')),
                QTableWidgetItem(perm_dict.get('course', '')),
                # QTableWidgetItem(perm_dict.get('faculty', ''))
            ]

            for col, item in enumerate(items):
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.permissions_table.setItem(row, col, item)

        self.permissions_table.resizeColumnsToContents()

    def refresh_data(self):
        """Обновление данных в таблице"""
        # Получение данных из БД
        if self.user_data['role'] == 'admin':
            if hasattr(self, 'course_filter') and self.course_filter.currentText() != 'Все курсы':
                applicants = self.db.get_applicants(self.user_data['id'], 'admin', self.course_filter.currentText())
            else:
                applicants = self.db.get_applicants(self.user_data['id'], 'admin')
        else:
            applicants = self.db.get_applicants(self.user_data['id'])

        # Применение поиска
        search_text = self.search_input.text().lower().strip()
        if search_text:
            filtered_applicants = []
            for applicant in applicants:
                applicant_dict = dict(applicant)
                # Поиск по всем текстовым полям
                text_fields = [
                    str(applicant_dict.get('study_group', '')),
                    str(applicant_dict.get('student_name', '')),
                    str(applicant_dict.get('region', '')),
                    str(applicant_dict.get('city', '')),
                    str(applicant_dict.get('applicant_name', '')),
                    str(applicant_dict.get('phone', '')),
                    str(applicant_dict.get('course', '')),
                    # str(applicant_dict.get('faculty', ''))
                ]
                if any(search_text in field.lower() for field in text_fields):
                    filtered_applicants.append(applicant)
            applicants = filtered_applicants

        # Заполнение таблицы
        self.table.setRowCount(len(applicants))

        for row, applicant in enumerate(applicants):
            applicant_dict = dict(applicant)

            items = [
                QTableWidgetItem(str(applicant_dict.get('id', ''))),
                QTableWidgetItem(applicant_dict.get('study_group', '')),
                QTableWidgetItem(applicant_dict.get('rank', '')),
                QTableWidgetItem(applicant_dict.get('student_name', '')),
                QTableWidgetItem(applicant_dict.get('region', '')),
                QTableWidgetItem(applicant_dict.get('city', '')),
                QTableWidgetItem(applicant_dict.get('category', '')),
                QTableWidgetItem(applicant_dict.get('applicant_name', '')),
                QTableWidgetItem(applicant_dict.get('phone', '')),
                QTableWidgetItem(applicant_dict.get('status', '')),
                QTableWidgetItem(applicant_dict.get('document_status', '')),
                QTableWidgetItem(applicant_dict.get('course', '')),
                # QTableWidgetItem(applicant_dict.get('faculty', ''))
            ]

            for col, item in enumerate(items):
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, col, item)

        self.table.resizeColumnsToContents()

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

    def add_applicant(self):
        """Добавление нового абитуриента"""
        dialog = ApplicantDialog()
        if dialog.exec():
            data = dialog.get_data()

            # Проверка обязательных полей
            if not data['applicant_name']:
                QMessageBox.warning(self, 'Ошибка', 'ФИО абитуриента обязательно!')
                return

            self.db.add_applicant(self.user_data['id'], data)
            self.refresh_data()
            self.stats_tab.update_statistics()

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

        dialog = ApplicantDialog(applicant_data, self)
        if dialog.exec():
            data = dialog.get_data()
            self.db.update_applicant(applicant_id, data)
            self.refresh_data()
            self.stats_tab.update_statistics()

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
        file_path, _ = QFileDialog.getOpenFileName(
            self, 'Выберите Excel файл', '',
            'Excel Files (*.xlsx *.xls);;All Files (*)'
        )

        if file_path:
            if self.db.import_from_excel(file_path, self.user_data['id']):
                QMessageBox.information(self, 'Успех', 'Данные успешно импортированы!')
                self.refresh_data()
                self.stats_tab.update_statistics()
            else:
                QMessageBox.critical(self, 'Ошибка', 'Ошибка при импорте данных!')

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