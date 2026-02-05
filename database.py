# database.py
import sqlite3
import pandas as pd
from datetime import datetime
import json


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
                course TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Таблица абитуриентов
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS applicants (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_by INTEGER,
                study_group TEXT,
                rank TEXT,
                student_name TEXT,
                region TEXT,
                city TEXT,
                category TEXT,
                applicant_name TEXT,
                phone TEXT,
                status TEXT,
                document_status TEXT,
                notes TEXT,
                course TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (created_by) REFERENCES users (id)
            )
        ''')

        # Добавляем администратора по умолчанию
        cursor.execute('SELECT * FROM users WHERE username = ?', ('admin',))
        if not cursor.fetchone():
            cursor.execute(
                'INSERT INTO users (username, password, full_name, role) VALUES (?, ?, ?, ?)',
                ('admin', 'admin123', 'Администратор системы', 'admin')
            )

        self.conn.commit()

    def get_user_by_credentials(self, username, password):
        """Получение пользователя по логину и паролю"""
        cursor = self.conn.cursor()
        cursor.execute(
            'SELECT * FROM users WHERE username = ? AND password = ?',
            (username, password)
        )
        return cursor.fetchone()

    def add_user(self, username, password, full_name, role='user', course=None):
        """Добавление нового пользователя"""
        cursor = self.conn.cursor()
        try:
            cursor.execute(
                '''INSERT INTO users (username, password, full_name, role, course) 
                   VALUES (?, ?, ?, ?, ?)''',
                (username, password, full_name, role, course)
            )
            self.conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            return None

    def add_applicant(self, user_id, data):
        """Добавление абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO applicants (
                created_by, study_group, rank, student_name, region, city,
                category, applicant_name, phone, status, document_status, notes, course
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            user_id, data['study_group'], data['rank'], data['student_name'],
            data['region'], data['city'], data['category'], data['applicant_name'],
            data['phone'], data['status'], data['document_status'],
            data.get('notes', ''), data.get('course', '')
        ))
        self.conn.commit()
        return cursor.lastrowid

    def get_applicants(self, user_id=None, role=None, course=None):
        """Получение списка абитуриентов"""
        cursor = self.conn.cursor()
        if role == 'admin':
            if course:
                cursor.execute('SELECT * FROM applicants WHERE course = ? ORDER BY id DESC', (course,))
            else:
                cursor.execute('SELECT * FROM applicants ORDER BY id DESC')
        else:
            cursor.execute(
                'SELECT * FROM applicants WHERE created_by = ? ORDER BY id DESC',
                (user_id,)
            )

        return cursor.fetchall()

    def update_applicant(self, applicant_id, data):
        """Обновление данных абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('''
            UPDATE applicants SET
                study_group = ?, rank = ?, student_name = ?, region = ?,
                city = ?, category = ?, applicant_name = ?, phone = ?,
                status = ?, document_status = ?, notes = ?, course = ?
            WHERE id = ?
        ''', (
            data['study_group'], data['rank'], data['student_name'],
            data['region'], data['city'], data['category'], data['applicant_name'],
            data['phone'], data['status'], data['document_status'],
            data.get('notes', ''), data.get('course', ''), applicant_id
        ))
        self.conn.commit()
        return cursor.rowcount > 0

    def delete_applicant(self, applicant_id):
        """Удаление абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM applicants WHERE id = ?', (applicant_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    def get_statistics(self, user_id=None, role=None, course=None):
        """Получение статистики"""
        cursor = self.conn.cursor()

        if role == 'admin':
            if course and course != 'Все курсы':
                # Статистика по конкретному курсу для админа
                cursor.execute('''
                    SELECT 
                        ? as course,
                        COUNT(*) as total,
                        COUNT(CASE WHEN status = 'поступает' THEN 1 END) as applying,
                        COUNT(CASE WHEN status = 'не поступает' THEN 1 END) as refused,
                        COUNT(CASE WHEN category = 'муж' THEN 1 END) as male,
                        COUNT(CASE WHEN category = 'жен' THEN 1 END) as female,
                        COUNT(CASE WHEN category = 'в/сл' THEN 1 END) as military,
                        COUNT(CASE WHEN document_status = 'Формируется в военкомате' THEN 1 END) as doc1,
                        COUNT(CASE WHEN document_status = 'Отправлено в ВА ВКО' THEN 1 END) as doc2,
                        COUNT(CASE WHEN document_status = 'В ВА ВКО' THEN 1 END) as doc3
                    FROM applicants 
                    WHERE course = ?
                ''', (course, course))
            else:
                # Общая статистика по всем курсам для админа - группируем по курсам
                cursor.execute('''
                    SELECT 
                        course,
                        COUNT(*) as total,
                        COUNT(CASE WHEN status = 'поступает' THEN 1 END) as applying,
                        COUNT(CASE WHEN status = 'не поступает' THEN 1 END) as refused,
                        COUNT(CASE WHEN category = 'муж' THEN 1 END) as male,
                        COUNT(CASE WHEN category = 'жен' THEN 1 END) as female,
                        COUNT(CASE WHEN category = 'в/сл' THEN 1 END) as military,
                        COUNT(CASE WHEN document_status = 'Формируется в военкомате' THEN 1 END) as doc1,
                        COUNT(CASE WHEN document_status = 'Отправлено в ВА ВКО' THEN 1 END) as doc2,
                        COUNT(CASE WHEN document_status = 'В ВА ВКО' THEN 1 END) as doc3
                    FROM applicants 
                    GROUP BY course
                    ORDER BY 
                        CASE 
                            WHEN course = '1 курс' THEN 1
                            WHEN course = '2 курс' THEN 2
                            WHEN course = '3 курс' THEN 3
                            WHEN course = '4 курс' THEN 4
                            WHEN course = '5 курс' THEN 5
                            ELSE 6
                        END
                ''')
        else:
            # Статистика для обычного пользователя (только его данные)
            user_course = self.conn.execute(
                'SELECT course FROM users WHERE id = ?',
                (user_id,)
            ).fetchone()

            user_course = user_course['course'] if user_course else '1 курс'

            cursor.execute('''
                SELECT 
                    ? as course,
                    COUNT(*) as total,
                    COUNT(CASE WHEN status = 'поступает' THEN 1 END) as applying,
                    COUNT(CASE WHEN status = 'не поступает' THEN 1 END) as refused,
                    COUNT(CASE WHEN category = 'муж' THEN 1 END) as male,
                    COUNT(CASE WHEN category = 'жен' THEN 1 END) as female,
                    COUNT(CASE WHEN category = 'в/сл' THEN 1 END) as military,
                    COUNT(CASE WHEN document_status = 'Формируется в военкомате' THEN 1 END) as doc1,
                    COUNT(CASE WHEN document_status = 'Отправлено в ВА ВКО' THEN 1 END) as doc2,
                    COUNT(CASE WHEN document_status = 'В ВА ВКО' THEN 1 END) as doc3
                FROM applicants 
                WHERE created_by = ? AND course = ?
            ''', (user_course, user_id, user_course))

        stats = cursor.fetchall()

        # Если нет данных, возвращаем пустой результат
        if not stats:
            if course and course != 'Все курсы':
                return [{
                    'course': course,
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

        return stats

    def import_from_excel(self, file_path, user_id):
        """Импорт данных из Excel файла"""
        try:
            # Чтение всех листов
            xls = pd.ExcelFile(file_path)

            for sheet_name in xls.sheet_names:
                if 'курс' in sheet_name.lower():
                    df = pd.read_excel(xls, sheet_name=sheet_name)

                    # Определение курса из названия листа


                    # Преобразование данных
                    for _, row in df.iterrows():
                        if pd.notna(row.get('ФИО абитуриента', '')):
                            applicant_data = {
                                'study_group': str(row.get('уч.гр.', '')),
                                'rank': str(row.get('В.зв', 'ряд.')),
                                'student_name': str(row.get('ФИО', '')),
                                'region': str(row.get('Субъект РФ', '')),
                                'city': str(row.get('Населённый пункт', '')),
                                'category': str(row.get('категория', 'муж')),
                                'applicant_name': str(row.get('ФИО абитуриента', '')),
                                'phone': str(row.get('Телефон', '')),
                                'status': str(row.get('Статус', '1)поступает')),
                                'document_status': str(row.get('Состояние личного дела на поступление', '')),
                                'notes': str(row.get('Примечание', '')),
                                'course': sheet_name.lower().strip()
                            }

                            self.add_applicant(user_id, applicant_data)

            return True
        except Exception as e:
            print(f"Ошибка импорта: {e}")
            return False

    def close(self):
        """Закрытие соединения с БД"""
        if self.conn:
            self.conn.close()