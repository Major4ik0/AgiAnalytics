# -*- coding: utf-8 -*-
import sqlite3
import pandas as pd
from datetime import datetime


class Database:
    def __init__(self):
        self.conn = None
        self.connect()
        self.create_tables()

    def connect(self):
        """Подключение к базе данных"""
        self.conn = sqlite3.connect('agitation.db', check_same_thread=False)
        self.conn.row_factory = sqlite3.Row

    def create_tables(self):
        """Создание таблиц в базе данных"""
        cursor = self.conn.cursor()

        # Таблица пользователей
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                full_name TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                department_id INTEGER,
                position TEXT,
                rank TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Таблица подразделений (кафедры, факультеты)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS departments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                parent_id INTEGER,
                type TEXT DEFAULT 'department',
                head_user_id INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (parent_id) REFERENCES departments(id),
                FOREIGN KEY (head_user_id) REFERENCES users(id)
            )
        ''')

        # Таблица вариантов образования
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS education_types (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                sort_order INTEGER DEFAULT 0
            )
        ''')

        # Таблица статусов документов
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS document_statuses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                sort_order INTEGER DEFAULT 0
            )
        ''')

        # Таблица планов по подразделениям
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS plans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                department_id INTEGER NOT NULL,
                year INTEGER NOT NULL,
                plan_m INTEGER DEFAULT 0,
                plan_f INTEGER DEFAULT 0,
                plan_military INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (department_id) REFERENCES departments(id),
                UNIQUE(department_id, year)
            )
        ''')

        # Таблица прав доступа пользователей к подразделениям
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_department_permissions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                department_id INTEGER NOT NULL,
                can_view BOOLEAN DEFAULT 1,
                can_edit_plan BOOLEAN DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users(id),
                FOREIGN KEY (department_id) REFERENCES departments(id),
                UNIQUE(user_id, department_id)
            )
        ''')

        # Основная таблица абитуриентов (обновленная структура)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS applicants (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                -- Поля абитуриента
                applicant_name TEXT NOT NULL,
                region TEXT,
                city TEXT,
                category TEXT CHECK(category IN ('м', 'ж', 'всл')),
                phone TEXT,
                education TEXT,
                status TEXT CHECK(status IN ('поступает', 'отказывается')),
                document_status TEXT,

                -- Поля агитатора
                agitator_department TEXT,
                agitator_name TEXT,
                agitator_course TEXT,
                agitator_group TEXT,
                agitator_rank TEXT,
                agitator_is_cadet BOOLEAN DEFAULT 0,

                -- Служебные поля
                created_by INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

                FOREIGN KEY (created_by) REFERENCES users(id)
            )
        ''')

        # Добавляем начальные данные
        self.init_default_data()
        self.clean_duplicate_departments()

        self.conn.commit()

    def init_default_data(self):
        """Инициализация начальных данных (только если таблицы пустые)"""
        cursor = self.conn.cursor()

        # Добавляем администратора по умолчанию (только если нет ни одного пользователя)
        cursor.execute('SELECT COUNT(*) as count FROM users')
        user_count = cursor.fetchone()['count']

        if user_count == 0:
            cursor.execute(
                'INSERT INTO users (username, password, full_name, role) VALUES (?, ?, ?, ?)',
                ('admin', 'admin', 'Администратор системы', 'admin')
            )
            admin_id = cursor.lastrowid

            # Создаем главное подразделение для админа
            cursor.execute(
                'INSERT OR IGNORE INTO departments (name, type, head_user_id) VALUES (?, ?, ?)',
                ('Все подразделения', 'root', admin_id)
            )

        # Добавляем варианты образования (только если таблица пустая)
        cursor.execute('SELECT COUNT(*) as count FROM education_types')
        edu_count = cursor.fetchone()['count']

        if edu_count == 0:
            education_defaults = ['СОШ', 'СПО', 'СВУ', 'ПКУ', 'КК']
            for i, edu in enumerate(education_defaults):
                cursor.execute(
                    'INSERT INTO education_types (name, sort_order) VALUES (?, ?)',
                    (edu, i)
                )

        # Добавляем статусы документов (только если таблица пустая)
        cursor.execute('SELECT COUNT(*) as count FROM document_statuses')
        doc_count = cursor.fetchone()['count']

        if doc_count == 0:
            document_defaults = ['ВК', 'ОК', 'ВА ВКО']
            for i, doc in enumerate(document_defaults):
                cursor.execute(
                    'INSERT INTO document_statuses (name, sort_order) VALUES (?, ?)',
                    (doc, i)
                )

        # Добавляем примерные подразделения (только если таблица пустая)
        cursor.execute('SELECT COUNT(*) as count FROM departments WHERE type != "root"')
        dept_count = cursor.fetchone()['count']

        if dept_count == 0:
            departments = [
                ('Факультет 1', 'faculty'),
                ('Факультет 2', 'faculty'),
                ('Кафедра 1', 'department'),
                ('Кафедра 2', 'department'),
            ]
            for name, dept_type in departments:
                cursor.execute(
                    'INSERT INTO departments (name, type) VALUES (?, ?)',
                    (name, dept_type)
                )

    def clean_duplicate_departments(self):
        """Очистка дублирующихся подразделений"""
        cursor = self.conn.cursor()

        # Находим дубликаты
        cursor.execute('''
            SELECT name, MIN(id) as keep_id, COUNT(*) as cnt
            FROM departments
            WHERE type != 'root'
            GROUP BY name
            HAVING cnt > 1
        ''')

        duplicates = cursor.fetchall()

        for dup in duplicates:
            name = dup['name']
            keep_id = dup['keep_id']

            # Удаляем дубликаты, оставляя один
            cursor.execute('DELETE FROM departments WHERE name = ? AND id != ?', (name, keep_id))

        self.conn.commit()
        print(f"Очищено дубликатов подразделений: {len(duplicates)}")

    # ==================== РАБОТА С АБИТУРИЕНТАМИ ====================

    def add_applicant(self, user_id, data):
        """Добавление абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO applicants (
                applicant_name, region, city, category, phone, education,
                status, document_status, agitator_department, agitator_name,
                agitator_course, agitator_group, agitator_rank, agitator_is_cadet,
                created_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data.get('applicant_name', ''),
            data.get('region', ''),
            data.get('city', ''),
            data.get('category', ''),
            data.get('phone', ''),
            data.get('education', ''),
            data.get('status', 'поступает'),
            data.get('document_status', ''),
            data.get('agitator_department', ''),
            data.get('agitator_name', ''),
            data.get('agitator_course', ''),
            data.get('agitator_group', ''),
            data.get('agitator_rank', ''),
            1 if data.get('agitator_is_cadet') else 0,
            user_id
        ))
        self.conn.commit()
        return cursor.lastrowid

    def get_applicants(self, user_id=None, role=None, department=None, filters=None):
        """Получение списка абитуриентов с учетом прав доступа и фильтров"""
        cursor = self.conn.cursor()

        query = '''
            SELECT 
                a.id,
                a.applicant_name,
                a.region,
                a.city,
                a.category,
                a.phone,
                a.education,
                a.status,
                a.document_status,
                a.agitator_department,
                a.agitator_name,
                a.agitator_course,
                a.agitator_group,
                a.agitator_rank,
                a.agitator_is_cadet,
                a.created_at
            FROM applicants a
            WHERE 1=1
        '''
        params = []

        if role == 'admin':
            # Админ видит всех
            pass
        else:
            # Обычный пользователь видит только своих абитуриентов
            query += ' AND a.created_by = ?'
            params.append(user_id)

        # Фильтр по подразделению
        if department and department != 'Все курсы' and department != 'Все подразделения':
            query += ' AND a.agitator_course = ?'
            params.append(department)

        # Расширенные фильтры
        if filters:
            if filters.get('applicant_name'):
                query += ' AND a.applicant_name LIKE ?'
                params.append(f"%{filters['applicant_name']}%")

            if filters.get('region'):
                query += ' AND a.region LIKE ?'
                params.append(f"%{filters['region']}%")

            if filters.get('city'):
                query += ' AND a.city LIKE ?'
                params.append(f"%{filters['city']}%")

            if filters.get('category'):
                query += ' AND a.category = ?'
                params.append(filters['category'])

            if filters.get('education'):
                placeholders = ','.join(['?'] * len(filters['education']))
                query += f' AND a.education IN ({placeholders})'
                params.extend(filters['education'])

            if filters.get('status'):
                query += ' AND a.status = ?'
                params.append(filters['status'])

            if filters.get('agitator_name'):
                query += ' AND a.agitator_name LIKE ?'
                params.append(f"%{filters['agitator_name']}%")

            if filters.get('agitator_department'):
                query += ' AND a.agitator_department = ?'
                params.append(filters['agitator_department'])

            if 'agitator_is_cadet' in filters:
                query += ' AND a.agitator_is_cadet = ?'
                params.append(1 if filters['agitator_is_cadet'] else 0)

            if filters.get('document_status'):
                query += ' AND a.document_status = ?'
                params.append(filters['document_status'])

            if filters.get('agitator_course'):
                query += ' AND a.agitator_course = ?'
                params.append(filters['agitator_course'])

            if filters.get('agitator_group'):
                query += ' AND a.agitator_group LIKE ?'
                params.append(f"%{filters['agitator_group']}%")

        query += ' ORDER BY a.id DESC'
        cursor.execute(query, params)
        return cursor.fetchall()

    def update_applicant(self, applicant_id, data):
        """Обновление данных абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('''
            UPDATE applicants SET
                applicant_name = ?, region = ?, city = ?, category = ?,
                phone = ?, education = ?, status = ?, document_status = ?,
                agitator_department = ?, agitator_name = ?, agitator_course = ?,
                agitator_group = ?, agitator_rank = ?, agitator_is_cadet = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ''', (
            data.get('applicant_name', ''),
            data.get('region', ''),
            data.get('city', ''),
            data.get('category', ''),
            data.get('phone', ''),
            data.get('education', ''),
            data.get('status', 'поступает'),
            data.get('document_status', ''),
            data.get('agitator_department', ''),
            data.get('agitator_name', ''),
            data.get('agitator_course', ''),
            data.get('agitator_group', ''),
            data.get('agitator_rank', ''),
            1 if data.get('agitator_is_cadet') else 0,
            applicant_id
        ))
        self.conn.commit()
        return cursor.rowcount > 0

    def delete_applicant(self, applicant_id):
        """Удаление абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM applicants WHERE id = ?', (applicant_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    # ==================== РАБОТА СО СПРАВОЧНИКАМИ ====================

    def get_education_types(self):
        """Получение списка типов образования"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT name FROM education_types ORDER BY sort_order')
        return [row['name'] for row in cursor.fetchall()]

    def add_education_type(self, name):
        """Добавление типа образования"""
        cursor = self.conn.cursor()
        try:
            cursor.execute('INSERT INTO education_types (name) VALUES (?)', (name,))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False

    def delete_education_type(self, name):
        """Удаление типа образования"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM education_types WHERE name = ?', (name,))
        self.conn.commit()
        return cursor.rowcount > 0

    def get_document_statuses(self):
        """Получение списка статусов документов"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT name FROM document_statuses ORDER BY sort_order')
        return [row['name'] for row in cursor.fetchall()]

    def add_document_status(self, name):
        """Добавление статуса документов"""
        cursor = self.conn.cursor()
        try:
            cursor.execute('INSERT INTO document_statuses (name) VALUES (?)', (name,))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False

    def delete_document_status(self, name):
        """Удаление статуса документов"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM document_statuses WHERE name = ?', (name,))
        self.conn.commit()
        return cursor.rowcount > 0

    def get_departments(self):
        """Получение списка подразделений (уникальные)"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT DISTINCT id, name, type FROM departments WHERE type != "root" ORDER BY name')
        return cursor.fetchall()

    def add_department(self, name, dept_type='department', parent_id=None):
        """Добавление подразделения"""
        cursor = self.conn.cursor()
        cursor.execute(
            'INSERT INTO departments (name, type, parent_id) VALUES (?, ?, ?)',
            (name, dept_type, parent_id)
        )
        self.conn.commit()
        return cursor.lastrowid

    def delete_department(self, dept_id):
        """Удаление подразделения"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM departments WHERE id = ?', (dept_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    # ==================== ПЛАНЫ ====================

    def get_plan(self, department_id, year):
        """Получение плана для подразделения на год"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT plan_m, plan_f, plan_military 
            FROM plans 
            WHERE department_id = ? AND year = ?
        ''', (department_id, year))
        result = cursor.fetchone()
        if result:
            return dict(result)
        return {'plan_m': 0, 'plan_f': 0, 'plan_military': 0}

    def set_plan(self, department_id, year, plan_m, plan_f, plan_military):
        """Установка плана для подразделения на год"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO plans (department_id, year, plan_m, plan_f, plan_military, updated_at)
            VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(department_id, year) DO UPDATE SET
                plan_m = excluded.plan_m,
                plan_f = excluded.plan_f,
                plan_military = excluded.plan_military,
                updated_at = CURRENT_TIMESTAMP
        ''', (department_id, year, plan_m, plan_f, plan_military))
        self.conn.commit()
        return True

    # ==================== ПРАВА ДОСТУПА ====================

    def get_user_department_permissions(self, user_id):
        """Получение прав пользователя на подразделения"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT p.*, d.name as department_name
            FROM user_department_permissions p
            JOIN departments d ON d.id = p.department_id
            WHERE p.user_id = ?
        ''', (user_id,))
        return cursor.fetchall()

    def add_user_department_permission(self, user_id, department_id, can_view=True, can_edit_plan=False):
        """Добавление права на подразделение"""
        cursor = self.conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO user_department_permissions (user_id, department_id, can_view, can_edit_plan)
                VALUES (?, ?, ?, ?)
            ''', (user_id, department_id, can_view, can_edit_plan))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False

    # ==================== СТАТИСТИКА ====================

    def get_statistics_by_department(self, department_name=None):
        """Получение статистики по подразделению с разделением по полу"""
        cursor = self.conn.cursor()

        query = '''
            SELECT 
                -- Поступают по статусам документов
                COUNT(CASE WHEN status = 'поступает' AND document_status = 'ВК' THEN 1 END) as applying_vk,
                COUNT(CASE WHEN status = 'поступает' AND document_status = 'ОК' THEN 1 END) as applying_ok,
                COUNT(CASE WHEN status = 'поступает' AND document_status = 'ВА ВКО' THEN 1 END) as applying_vavko,

                -- Поступают по полу
                COUNT(CASE WHEN status = 'поступает' AND category = 'м' THEN 1 END) as applying_male,
                COUNT(CASE WHEN status = 'поступает' AND category = 'ж' THEN 1 END) as applying_female,
                COUNT(CASE WHEN status = 'поступает' AND category = 'всл' THEN 1 END) as applying_military,

                -- Отказались по полу
                COUNT(CASE WHEN status = 'отказывается' AND category = 'м' THEN 1 END) as refused_male,
                COUNT(CASE WHEN status = 'отказывается' AND category = 'ж' THEN 1 END) as refused_female,
                COUNT(CASE WHEN status = 'отказывается' AND category = 'всл' THEN 1 END) as refused_military,

                -- Документы по полу
                COUNT(CASE WHEN document_status = 'ВК' AND category = 'м' THEN 1 END) as vk_male,
                COUNT(CASE WHEN document_status = 'ВК' AND category = 'ж' THEN 1 END) as vk_female,
                COUNT(CASE WHEN document_status = 'ВК' AND category = 'всл' THEN 1 END) as vk_military,

                COUNT(CASE WHEN document_status = 'ОК' AND category = 'м' THEN 1 END) as ok_male,
                COUNT(CASE WHEN document_status = 'ОК' AND category = 'ж' THEN 1 END) as ok_female,
                COUNT(CASE WHEN document_status = 'ОК' AND category = 'всл' THEN 1 END) as ok_military,

                COUNT(CASE WHEN document_status = 'ВА ВКО' AND category = 'м' THEN 1 END) as vavko_male,
                COUNT(CASE WHEN document_status = 'ВА ВКО' AND category = 'ж' THEN 1 END) as vavko_female,
                COUNT(CASE WHEN document_status = 'ВА ВКО' AND category = 'всл' THEN 1 END) as vavko_military,

                -- Общее количество
                COUNT(*) as total
            FROM applicants
            WHERE 1=1
        '''
        params = []

        if department_name and department_name != 'Все подразделения':
            query += ' AND agitator_department = ?'
            params.append(department_name)

        cursor.execute(query, params)
        result = cursor.fetchone()

        if result:
            return dict(result)
        return {
            'applying_vk': 0,
            'applying_ok': 0,
            'applying_vavko': 0,
            'applying_male': 0,
            'applying_female': 0,
            'applying_military': 0,
            'refused_male': 0,
            'refused_female': 0,
            'refused_military': 0,
            'vk_male': 0,
            'vk_female': 0,
            'vk_military': 0,
            'ok_male': 0,
            'ok_female': 0,
            'ok_military': 0,
            'vavko_male': 0,
            'vavko_female': 0,
            'vavko_military': 0,
            'total': 0
        }

    # ==================== ПОЛЬЗОВАТЕЛИ ====================

    def get_user_by_credentials(self, username, password):
        """Получение пользователя по логину и паролю"""
        cursor = self.conn.cursor()
        cursor.execute(
            'SELECT * FROM users WHERE username = ? AND password = ?',
            (username, password)
        )
        return cursor.fetchone()

    def get_user_by_id(self, user_id):
        """Получение пользователя по ID"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM users WHERE id = ?', (user_id,))
        return cursor.fetchone()

    def get_all_users(self):
        """Получение всех пользователей"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT u.*, d.name as department_name 
            FROM users u
            LEFT JOIN departments d ON d.id = u.department_id
            ORDER BY u.id DESC
        ''')
        return cursor.fetchall()

    def add_user(self, username, password, full_name, role='user', department_id=None, position=None, rank=None):
        """Добавление нового пользователя"""
        cursor = self.conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO users (username, password, full_name, role, department_id, position, rank) 
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (username, password, full_name, role, department_id, position, rank))
            self.conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError as e:
            print(f"Ошибка добавления пользователя: {e}")
            return None

    def update_user(self, user_id, data):
        """Обновление данных пользователя"""
        cursor = self.conn.cursor()
        try:
            cursor.execute('''
                UPDATE users SET
                    username = ?, password = ?, full_name = ?, 
                    role = ?, department_id = ?, position = ?, rank = ?
                WHERE id = ?
            ''', (
                data['username'], data['password'], data['full_name'],
                data['role'], data.get('department_id'), data.get('position'), data.get('rank'),
                user_id
            ))
            self.conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Ошибка обновления пользователя: {e}")
            return False

    def delete_user(self, user_id):
        """Удаление пользователя"""
        cursor = self.conn.cursor()
        try:
            cursor.execute('DELETE FROM user_department_permissions WHERE user_id = ?', (user_id,))
            cursor.execute('DELETE FROM users WHERE id = ?', (user_id,))
            self.conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Ошибка удаления пользователя: {e}")
            return False

    # Добавьте этот метод в класс Database в файл database.py

    def get_statistics(self, user_id=None, role=None, course=None, faculty=None):
        """Получение статистики для StatisticsWidget (совместимость со старым кодом)"""
        cursor = self.conn.cursor()

        if role == 'admin':
            if course and course != 'Все курсы':
                # Статистика по конкретному курсу
                cursor.execute('''
                    SELECT 
                        ? as course,
                        COUNT(*) as total,
                        COUNT(CASE WHEN status = 'поступает' THEN 1 END) as applying,
                        COUNT(CASE WHEN status = 'отказывается' THEN 1 END) as refused,
                        COUNT(CASE WHEN category = 'м' THEN 1 END) as male,
                        COUNT(CASE WHEN category = 'ж' THEN 1 END) as female,
                        COUNT(CASE WHEN category = 'всл' THEN 1 END) as military,
                        COUNT(CASE WHEN document_status = 'ВК' THEN 1 END) as doc1,
                        COUNT(CASE WHEN document_status = 'ОК' THEN 1 END) as doc2,
                        COUNT(CASE WHEN document_status = 'ВА ВКО' THEN 1 END) as doc3
                    FROM applicants 
                    WHERE agitator_course = ?
                ''', (course, course))
            else:
                # Общая статистика по всем курсам
                cursor.execute('''
                    SELECT 
                        agitator_course as course,
                        COUNT(*) as total,
                        COUNT(CASE WHEN status = 'поступает' THEN 1 END) as applying,
                        COUNT(CASE WHEN status = 'отказывается' THEN 1 END) as refused,
                        COUNT(CASE WHEN category = 'м' THEN 1 END) as male,
                        COUNT(CASE WHEN category = 'ж' THEN 1 END) as female,
                        COUNT(CASE WHEN category = 'всл' THEN 1 END) as military,
                        COUNT(CASE WHEN document_status = 'ВК' THEN 1 END) as doc1,
                        COUNT(CASE WHEN document_status = 'ОК' THEN 1 END) as doc2,
                        COUNT(CASE WHEN document_status = 'ВА ВКО' THEN 1 END) as doc3
                    FROM applicants 
                    WHERE agitator_course IS NOT NULL AND agitator_course != ''
                    GROUP BY agitator_course
                ''')
        else:
            # Для обычного пользователя
            cursor.execute('''
                SELECT 
                    agitator_course as course,
                    COUNT(*) as total,
                    COUNT(CASE WHEN status = 'поступает' THEN 1 END) as applying,
                    COUNT(CASE WHEN status = 'отказывается' THEN 1 END) as refused,
                    COUNT(CASE WHEN category = 'м' THEN 1 END) as male,
                    COUNT(CASE WHEN category = 'ж' THEN 1 END) as female,
                    COUNT(CASE WHEN category = 'всл' THEN 1 END) as military,
                    COUNT(CASE WHEN document_status = 'ВК' THEN 1 END) as doc1,
                    COUNT(CASE WHEN document_status = 'ОК' THEN 1 END) as doc2,
                    COUNT(CASE WHEN document_status = 'ВА ВКО' THEN 1 END) as doc3
                FROM applicants 
                WHERE created_by = ?
                GROUP BY agitator_course
            ''', (user_id,))

        result = cursor.fetchall()

        # Преобразуем в список словарей для совместимости
        stats_list = []
        for row in result:
            stats_dict = dict(row)
            stats_list.append(stats_dict)

        if not stats_list:
            # Возвращаем пустую статистику
            return [{
                'course': course or '1 курс',
                'total': 0,
                'applying': 0,
                'refused': 0,
                'male': 0,
                'female': 0,
                'military': 0,
                'doc1': 0,
                'doc2': 0,
                'doc3': 0
            }]

        return stats_list

    def close(self):
        """Закрытие соединения с БД"""
        if self.conn:
            self.conn.close()