# -*- coding: utf-8 -*-
import sqlite3
import pandas as pd


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
                faculty TEXT,
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
                faculty TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (created_by) REFERENCES users (id)
            )
        ''')

        # Таблица прав доступа
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS permissions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                permission_type TEXT NOT NULL,  -- 'all', 'faculty', 'course'
                faculty TEXT,
                course TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')

        # Добавляем администратора по умолчанию
        cursor.execute('SELECT * FROM users WHERE username = ?', ('admin',))
        if not cursor.fetchone():
            cursor.execute(
                'INSERT INTO users (username, password, full_name, role) VALUES (?, ?, ?, ?)',
                ('admin', 'admin', 'Администратор системы', 'admin')
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

    def get_all_users(self):
        """Получение всех пользователей"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM users ORDER BY id DESC')
        return cursor.fetchall()

    def get_user_by_id(self, user_id):
        """Получение пользователя по ID"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM users WHERE id = ?', (user_id,))
        return cursor.fetchone()

    def add_user(self, username, password, full_name, role='user', course=None, faculty=None):
        """Добавление нового пользователя"""
        cursor = self.conn.cursor()
        try:
            cursor.execute(
                '''INSERT INTO users (username, password, full_name, role, course, faculty) 
                   VALUES (?, ?, ?, ?, ?, ?)''',
                (username, password, full_name, role, course, faculty)
            )
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
                    role = ?, course = ?, faculty = ?
                WHERE id = ?
            ''', (
                data['username'], data['password'], data['full_name'],
                data['role'], data.get('course'), data.get('faculty'), user_id
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
            # Сначала удаляем права пользователя
            cursor.execute('DELETE FROM permissions WHERE user_id = ?', (user_id,))
            # Затем удаляем пользователя
            cursor.execute('DELETE FROM users WHERE id = ?', (user_id,))
            self.conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Ошибка удаления пользователя: {e}")
            return False

    def add_applicant(self, user_id, data):
        """Добавление абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO applicants (
                created_by, study_group, rank, student_name, region, city,
                category, applicant_name, phone, status, document_status, 
                notes, course, faculty
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            user_id, data['study_group'], data['rank'], data['student_name'],
            data['region'], data['city'], data['category'], data['applicant_name'],
            data['phone'], data['status'], data['document_status'],
            data.get('notes', ''), data.get('course', ''), data.get('faculty', '')
        ))
        self.conn.commit()
        return cursor.lastrowid

    def get_applicants(self, user_id=None, role=None, course=None, faculty=None):
        """Получение списка абитуриентов с учетом прав доступа"""
        cursor = self.conn.cursor()
        if role == 'admin':
            query = 'SELECT * FROM applicants WHERE 1=1'
            params = []

            if course and course != 'Все курсы':
                query += ' AND course = ?'
                params.append(course)

            query += ' ORDER BY id DESC'
            cursor.execute(query, params)
        else:
            # Для обычного пользователя - только его данные и данные по его правам
            cursor.execute('''
                SELECT DISTINCT a.* 
                FROM applicants a
                LEFT JOIN users u 
                LEFT JOIN permissions p ON u.id = p.user_id
                WHERE a.created_by = ? OR (
                    p.permission_type = 'all' OR
                    (p.permission_type = 'course' AND p.course = a.course)
                )
                ORDER BY a.id DESC
            ''', (user_id,))
        return cursor.fetchall()

    def import_from_excel(self, file_path, user_id, password=None, selected_sheets=None):
        """Импорт данных из Excel файла с паролем и выбором листов"""
        try:
            # Проверка пароля (если предоставлен)
            if password:
                cursor = self.conn.cursor()
                cursor.execute('SELECT password FROM users WHERE id = ?', (user_id,))
                user_record = cursor.fetchone()

                if not user_record or user_record['password'] != password:
                    return False, "Неверный пароль"

            # Чтение всех листов
            try:
                # Пробуем разные движки для чтения Excel
                try:
                    xls = pd.ExcelFile(file_path, engine='openpyxl')
                except:
                    try:
                        xls = pd.ExcelFile(file_path, engine='xlrd')
                    except:
                        xls = pd.ExcelFile(file_path)  # Пусть pandas сам выберет движок

                imported_count = 0
                duplicate_count = 0
                skipped_count = 0
                available_sheets = xls.sheet_names

                # Если не указаны листы для импорта, используем все листы с "курс"
                if not selected_sheets:
                    selected_sheets = [sheet for sheet in available_sheets if 'курс' in sheet.lower()]
                else:
                    # Фильтруем только те листы, которые есть в файле
                    selected_sheets = [sheet for sheet in selected_sheets if sheet in available_sheets]

                if not selected_sheets:
                    return False, "Нет подходящих листов для импорта. Ожидаются листы с названиями, содержащими 'курс'"

                for sheet_name in selected_sheets:
                    try:
                        # Читаем лист
                        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)

                        # Определяем курс из названия листа
                        course_from_sheet = '1 курс'  # по умолчанию
                        for course_num in ['1', '2', '3', '4', '5']:
                            if f'{course_num} курс' in sheet_name.lower():
                                course_from_sheet = f'{course_num} курс'
                                break
                            elif f'курс {course_num}' in sheet_name.lower():
                                course_from_sheet = f'{course_num} курс'
                                break

                        # Преобразование данных
                        for index, row in df.iterrows():
                            try:
                                # Пропускаем пустые строки
                                if row.isnull().all():
                                    continue

                                # Ищем ФИО абитуриента
                                applicant_name = ''
                                for col in df.columns:
                                    col_str = str(col)
                                    if 'абитуриент' in col_str.lower() and pd.notna(row[col]):
                                        applicant_name = str(row[col])
                                        break

                                if not applicant_name.strip():
                                    skipped_count += 1
                                    continue

                                # Собираем данные
                                applicant_data = {
                                    'study_group': self._get_value_from_row(row, df,
                                                                            ['уч.гр.', 'Уч. группа', 'учебная группа']),
                                    'rank': self._get_value_from_row(row, df, ['В.зв', 'Звание', 'воинское звание'],
                                                                     'ряд.'),
                                    'student_name': self._get_value_from_row(row, df,
                                                                             ['ФИО', 'ФИО студента', 'Студент']),
                                    'region': self._get_value_from_row(row, df, ['Субъект РФ', 'Регион', 'область']),
                                    'city': self._get_value_from_row(row, df, ['Населённый пункт', 'Город', 'город']),
                                    'category': self._get_value_from_row(row, df, ['категория', 'Категория', 'пол'],
                                                                         'муж'),
                                    'applicant_name': applicant_name,
                                    'phone': self._get_value_from_row(row, df,
                                                                      ['Телефон', 'телефон', 'контактный телефон']),
                                    'status': self._get_value_from_row(row, df,
                                                                       ['Статус', 'статус', 'Статус поступления'],
                                                                       '1)поступает'),
                                    'document_status': self._get_value_from_row(row, df,
                                                                                ['Состояние личного дела на поступление',
                                                                                 'Документы', 'документы']),
                                    'notes': self._get_value_from_row(row, df,
                                                                      ['Примечание', 'примечание', 'комментарий']),
                                    'course': course_from_sheet,
                                    'faculty': self._get_value_from_row(row, df, ['Факультет', 'факультет',
                                                                                  'Факультет/отделение'])
                                }

                                # Проверка на дубликат
                                if self.check_duplicate_applicant(user_id, applicant_data):
                                    duplicate_count += 1
                                    continue

                                self.add_applicant(user_id, applicant_data)
                                imported_count += 1

                            except Exception as e:
                                print(f"Ошибка в строке {index + 2} листа {sheet_name}: {e}")
                                skipped_count += 1
                                continue

                    except Exception as e:
                        print(f"Ошибка импорта листа {sheet_name}: {e}")
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
                return False, f"Ошибка чтения файла Excel: {str(e)}"

        except Exception as e:
            print(f"Общая ошибка импорта: {e}")
            return False, f"Ошибка импорта: {str(e)}"

    def check_duplicate_applicant(self, user_id, applicant_data):
        """Проверка на дубликат абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT COUNT(*) FROM applicants 
            WHERE created_by = ? 
            AND applicant_name = ?
            AND phone = ?
            AND course = ?
        ''', (user_id, applicant_data['applicant_name'],
              applicant_data['phone'], applicant_data['course']))
        count = cursor.fetchone()[0]
        return count > 0

    def _get_value_from_row(self, row, df, possible_columns, default=''):
        """Получение значения из строки по возможным названиям колонок"""
        for col in possible_columns:
            if col in df.columns:
                value = row[col]
                if pd.notna(value):
                    return str(value).strip()

        # Если не нашли в точных совпадениях, ищем частичные
        for df_col in df.columns:
            df_col_str = str(df_col)
            for possible in possible_columns:
                if possible.lower() in df_col_str.lower():
                    value = row[df_col]
                    if pd.notna(value):
                        return str(value).strip()

        return default

    def update_applicant(self, applicant_id, data):
        """Обновление данных абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('''
            UPDATE applicants SET
                study_group = ?, rank = ?, student_name = ?, region = ?,
                city = ?, category = ?, applicant_name = ?, phone = ?,
                status = ?, document_status = ?, notes = ?, course = ?, faculty = ?
            WHERE id = ?
        ''', (
            data['study_group'], data['rank'], data['student_name'],
            data['region'], data['city'], data['category'], data['applicant_name'],
            data['phone'], data['status'], data['document_status'],
            data.get('notes', ''), data.get('course', ''), data.get('faculty', ''), applicant_id
        ))
        self.conn.commit()
        return cursor.rowcount > 0

    def delete_applicant(self, applicant_id):
        """Удаление абитуриента"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM applicants WHERE id = ?', (applicant_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    # Методы для работы с правами доступа
    def get_user_permissions(self, user_id):
        """Получение прав доступа пользователя"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT p.*, u.* FROM permissions p
            left join users u on u.id = p.user_id
            WHERE user_id = ? 
            ORDER BY id DESC
        ''', (user_id,))
        return cursor.fetchall()

    def add_permission(self, user_id, permission_type, faculty=None, course=None):
        """Добавление права доступа"""
        cursor = self.conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO permissions (user_id, permission_type, faculty, course)
                VALUES (?, ?, ?, ?)
            ''', (user_id, permission_type, faculty, course))
            self.conn.commit()
            return cursor.lastrowid
        except Exception as e:
            print(f"Ошибка добавления права: {e}")
            return None

    def delete_permission(self, permission_id):
        """Удаление права доступа"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM permissions WHERE id = ?', (permission_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    def get_all_permissions(self):
        """Получение всех прав доступа"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT p.*, u.username, u.full_name 
            FROM permissions p
            JOIN users u ON p.user_id = u.id
            ORDER BY p.id DESC
        ''')
        return cursor.fetchall()

    def get_statistics(self, user_id=None, role=None, course=None, faculty=None):
        """Получение статистики с учетом прав доступа"""
        cursor = self.conn.cursor()

        if role == 'admin':
            if course and course != 'Все курсы':
                # Статистика по конкретному курсу для админа
                query = '''
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
                '''
                params = [course, course]

                if faculty and faculty != 'Все факультеты':
                    query = query.replace('WHERE course = ?', 'WHERE course = ? AND faculty = ?')
                    params.append(faculty)

                cursor.execute(query, params)
            else:
                # Общая статистика по всем курсам для админа - группируем по курсам
                query = '''
                    SELECT 
                        course,
                        faculty,
                        COUNT(*) as total,
                        COUNT(CASE WHEN status = '1)поступает' THEN 1 END) as applying,
                        COUNT(CASE WHEN status = '2)не поступает' THEN 1 END) as refused,
                        COUNT(CASE WHEN category = 'муж' THEN 1 END) as male,
                        COUNT(CASE WHEN category = 'жен' THEN 1 END) as female,
                        COUNT(CASE WHEN category = 'в/сл' THEN 1 END) as military,
                        COUNT(CASE WHEN document_status = '1)Формируется в военкомате' THEN 1 END) as doc1,
                        COUNT(CASE WHEN document_status = '2)Отправлено в ВА ВКО' THEN 1 END) as doc2,
                        COUNT(CASE WHEN document_status = '3)В ВА ВКО' THEN 1 END) as doc3
                    FROM applicants 
                '''

                params = []
                if faculty and faculty != 'Все факультеты':
                    query += ' WHERE faculty = ?'
                    params.append(faculty)

                query += '''
                    GROUP BY course, faculty
                    ORDER BY 
                        CASE 
                            WHEN course = '1 курс' THEN 1
                            WHEN course = '2 курс' THEN 2
                            WHEN course = '3 курс' THEN 3
                            WHEN course = '4 курс' THEN 4
                            WHEN course = '5 курс' THEN 5
                            ELSE 6
                        END
                '''

                cursor.execute(query, params)
        else:
            # Статистика для обычного пользователя (только его данные + права)
            user_info = self.get_user_by_id(user_id)
            user_course = user_info['course'] if user_info else '1 курс'
            user_faculty = user_info['faculty'] if user_info else None

            # Получаем права пользователя
            permissions = self.get_user_permissions(user_id)

            # Формируем запрос на основе прав
            if any(p['permission_type'] == 'all' for p in permissions):
                # Если есть право "все"
                query = '''
                    SELECT 
                        course,
                        COUNT(*) as total,
                        COUNT(CASE WHEN status = '1)поступает' THEN 1 END) as applying,
                        COUNT(CASE WHEN status = '2)не поступает' THEN 1 END) as refused,
                        COUNT(CASE WHEN category = 'муж' THEN 1 END) as male,
                        COUNT(CASE WHEN category = 'жен' THEN 1 END) as female,
                        COUNT(CASE WHEN category = 'в/сл' THEN 1 END) as military
                    FROM applicants 
                    GROUP BY course
                '''
                cursor.execute(query)
            else:
                # Собираем доступные курсы и факультеты из прав
                available_courses = set()
                available_faculties = set()

                for perm in permissions:
                    if perm['permission_type'] == 'course' and perm['course']:
                        available_courses.add(perm['course'])
                    elif perm['permission_type'] == 'faculty' and perm['faculty']:
                        available_faculties.add(perm['faculty'])

                # Добавляем собственные данные пользователя
                if user_course:
                    available_courses.add(user_course)
                if user_faculty:
                    available_faculties.add(user_faculty)

                # Формируем запрос
                if available_courses or available_faculties:
                    conditions = []
                    params = []

                    if available_courses:
                        courses_list = list(available_courses)
                        placeholders = ','.join(['?'] * len(courses_list))
                        conditions.append(f'course IN ({placeholders})')
                        params.extend(courses_list)

                    if available_faculties:
                        faculties_list = list(available_faculties)
                        placeholders = ','.join(['?'] * len(faculties_list))
                        conditions.append(f'faculty IN ({placeholders})')
                        params.extend(faculties_list)

                    where_clause = ' OR '.join(conditions)

                    cursor.execute(f'''
                        SELECT 
                            course,
                            COUNT(*) as total,
                            COUNT(CASE WHEN status = '1)поступает' THEN 1 END) as applying,
                            COUNT(CASE WHEN status = '2)не поступает' THEN 1 END) as refused,
                            COUNT(CASE WHEN category = 'муж' THEN 1 END) as male,
                            COUNT(CASE WHEN category = 'жен' THEN 1 END) as female,
                            COUNT(CASE WHEN category = 'в/сл' THEN 1 END) as military
                        FROM applicants 
                        WHERE {where_clause}
                        GROUP BY course
                    ''', params)
                else:
                    # Только свои данные
                    cursor.execute('''
                        SELECT 
                            ? as course,
                            COUNT(*) as total,
                            COUNT(CASE WHEN status = '1)поступает' THEN 1 END) as applying,
                            COUNT(CASE WHEN status = '2)не поступает' THEN 1 END) as refused,
                            COUNT(CASE WHEN category = 'муж' THEN 1 END) as male,
                            COUNT(CASE WHEN category = 'жен' THEN 1 END) as female,
                            COUNT(CASE WHEN category = 'в/сл' THEN 1 END) as military
                        FROM applicants 
                        WHERE created_by = ?
                    ''', (user_course, user_id))

        stats = cursor.fetchall()

        # Если нет данных, возвращаем пустой результат
        if not stats:
            if course and course != 'Все курсы':
                return [{
                    'course': course,
                    'faculty': faculty or '',
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
                                'course': sheet_name.lower().strip(),
                                'faculty': str(row.get('Факультет', ''))
                            }

                            self.add_applicant(user_id, applicant_data)

            return True
        except Exception as e:
            print(f"Ошибка импорта: {e}")
            return False

    def get_all_faculties(self):
        """Получение списка всех факультетов"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT DISTINCT faculty FROM applicants WHERE faculty IS NOT NULL AND faculty != ""')
        return [row['faculty'] for row in cursor.fetchall()]

    def get_all_courses(self):
        """Получение списка всех курсов"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT DISTINCT course FROM applicants WHERE course IS NOT NULL AND course != ""')
        return [row['course'] for row in cursor.fetchall()]

    def close(self):
        """Закрытие соединения с БД"""
        if self.conn:
            self.conn.close()