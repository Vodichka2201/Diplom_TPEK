import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import mysql.connector
import openpyxl
from datetime import datetime, timedelta
from dotenv import load_dotenv
import os

load_dotenv()
password = os.getenv("bd_pass")

# ==================== НАСТРОЙКИ ПОДКЛЮЧЕНИЯ ====================
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': password,
    'database': 'smartstudentdb_test'
}

# ==================== КЛАСС ПРИЛОЖЕНИЯ ====================

class MigrationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Миграция данных студентов колледжа")
        self.root.geometry("1100x650")
        self.root.resizable(True, True)

        self.excel_data = []
        self.db_connection = None

        self.create_widgets()

    def create_widgets(self):
        title = tk.Label(self.root, text="Миграция данных студентов колледжа",
                         font=("Arial", 16, "bold"))
        title.pack(pady=10)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=5)

        self.btn_load = tk.Button(btn_frame, text="1. Выбрать Excel-файл",
                                  command=self.load_excel, width=22, bg="#e0e0e0")
        self.btn_load.pack(side=tk.LEFT, padx=5)

        self.btn_save = tk.Button(btn_frame, text="2. Сохранить в БД",
                                  command=self.save_to_db, width=18, bg="#c8e6c9",
                                  state=tk.DISABLED)
        self.btn_save.pack(side=tk.LEFT, padx=5)

        self.btn_match = tk.Button(btn_frame, text="3. Привязать учебные заведения",
                           command=self.open_match_window, width=26, bg="#fff9c4")
        self.btn_match.pack(side=tk.LEFT, padx=5)

        # Таблица
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("№", "Фамилия", "Имя", "Отчество", "Дата рождения",
                   "Пол", "Телефон", "Специальность", "Курс", "Бюджет/Платно")

        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)

        widths = [35, 100, 90, 100, 90, 70, 110, 160, 50, 90]
        for col, w in zip(columns, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w)

        scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.status = tk.Label(self.root, text="Готов к работе. Выберите Excel-файл.",
                               bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 10))
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    def parse_date(self, date_str):
        """Пробует разные форматы даты, включая числовой формат Excel"""
        if date_str is None:
            return None
        
        # Если это уже datetime
        if isinstance(date_str, datetime):
            return date_str.strftime('%Y-%m-%d')
        
        date_str = str(date_str).strip()
        if not date_str:
            return None
        
        # Пробуем стандартные форматы
        for fmt in ['%d.%m.%Y', '%Y-%m-%d', '%d.%m.%y', '%Y.%m.%d', '%d/%m/%Y']:
            try:
                return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
            except ValueError:
                continue
        
        # Пробуем числовой формат Excel (дни от 1899-12-30)
        try:
            serial = int(date_str)
            if serial > 0:
                base_date = datetime(1899, 12, 30)
                return (base_date + timedelta(days=serial)).strftime('%Y-%m-%d')
        except (ValueError, OverflowError):
            pass
        
        return None

    def load_excel(self):
        filepath = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not filepath:
            return

        try:
            wb = openpyxl.load_workbook(filepath)
            ws = wb.active

            self.excel_data = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] is None:
                    continue

                self.excel_data.append({
                    'surname': str(row[0] or ''),
                    'name': str(row[1] or ''),
                    'patronymic': str(row[2] or ''),
                    'birth_date': self.parse_date(row[3]),
                    'gender': str(row[4] or 'мужской'),
                    'social_status': str(row[5] or ''),
                    'phone_number': str(row[6] or ''),
                    'adress': str(row[7] or ''),
                    'insurance_number': str(row[8] or ''),
                    'date_of_insurance': self.parse_date(row[9]),
                    'medical_policy': str(row[10] or ''),
                    'date_of_medical_policy': self.parse_date(row[11]),
                    'insurance_companie': str(row[12] or ''),
                    'native_language': str(row[13] or ''),
                    'nationality': str(row[14] or ''),
                    'passport_series': str(row[15] or ''),
                    'passport_number': str(row[16] or ''),
                    'issuer': str(row[17] or ''),
                    'issue_date': self.parse_date(row[18]),
                    'subdivision_code': str(row[19] or ''),
                    'place_of_birth': str(row[20] or ''),
                    'place_adress': str(row[21] or ''),
                    'citizenship': str(row[22] or ''),
                    'type_of_education': str(row[23] or ''),
                    'average_mark': float(str(row[24]).replace(',', '.')) if row[24] else None,
                    'reception_id': int(row[25]) if row[25] else None,
                    'reg_number': str(row[26] or ''),
                    'is_adopted': int(row[27]) if row[27] else 0,
                    'specialization_id': int(row[29]) if len(row) > 29 and row[29] else None,
                    'grade': int(row[30]) if len(row) > 30 and row[30] else 9,
                    'is_budget': int(row[31]) if len(row) > 31 and row[31] is not None else 1,
                    'is_full_time': int(row[32]) if len(row) > 32 and row[32] is not None else 1,
                    'specialization_code': str(row[33] or '') if len(row) > 33 else '',
                    'specialization_title': str(row[34] or '') if len(row) > 34 else '',
                })
                

            # Заполняем таблицу
            for item in self.tree.get_children():
                self.tree.delete(item)
            for i, data in enumerate(self.excel_data, 1):
                self.tree.insert("", tk.END, values=(
                    i, data['surname'], data['name'], data['patronymic'],
                    data['birth_date'], data['gender'], data['phone_number'],
                    data['specialization_code'] + ' ' + data['specialization_title'],
                    data['grade'],
                    'Бюджет' if data['is_budget'] == 1 else 'Платно'
                ))

            self.status.config(text=f"Загружено {len(self.excel_data)} записей из файла: {filepath}")
            self.btn_save.config(state=tk.NORMAL)
            

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")

    def save_to_db(self):
        if not self.excel_data:
            messagebox.showwarning("Внимание", "Сначала загрузите Excel-файл.")
            return

        ok = messagebox.askyesno(
            "Подтверждение",
            f"Сохранить {len(self.excel_data)} записей в базу данных?\n\n"
            "Данные будут записаны в таблицы:\n"
            "- people (личные данные)\n"
            "- person_extra_data (доп. сведения)\n"
            "- passport_data (паспорт)\n"
            "- person_insurance (СНИЛС, полис)\n"
            "- person_certificates (образование)\n"
            "- enrollments (поступление)"
        )
        if not ok:
            return

        try:
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()

            saved = 0
            skipped = 0
            errors = 0

            for data in self.excel_data:
                try:
                    person_id = self.insert_or_get_person(cursor, data)
                    if not person_id:
                        skipped += 1
                        continue

                    self.insert_person_extra_data(cursor, person_id, data)
                    self.insert_passport_data(cursor, person_id, data)
                    self.insert_person_insurance(cursor, person_id, data)
                    self.insert_person_certificate(cursor, person_id, data)
                    self.insert_enrollment(cursor, person_id, data)

                    saved += 1
                except Exception as e:
                    errors += 1
                    print(f"Ошибка в строке {data['surname']} {data['name']}: {e}")

            conn.commit()
            cursor.close()
            conn.close()

            self.status.config(
                text=f"Сохранение завершено. Сохранено: {saved}, пропущено: {skipped}, ошибок: {errors}."
            )
            messagebox.showinfo(
                "Готово",
                f"Результаты миграции:\n\n"
                f"Успешно сохранено: {saved}\n"
                f"Пропущено: {skipped}\n"
                f"Ошибок: {errors}"
            )

        except mysql.connector.Error as e:
            messagebox.showerror("Ошибка БД", f"Ошибка подключения:\n{e}")
            conn.rollback()

    def insert_or_get_person(self, cursor, data):
        birth_date = data['birth_date'] if data['birth_date'] else '2000-01-01'
        
        cursor.execute("""
            SELECT id FROM people
            WHERE surname = %s AND name = %s
            AND (patronymic = %s OR (patronymic IS NULL AND %s IS NULL))
            AND birth_date = %s
        """, (data['surname'], data['name'], data['patronymic'], data['patronymic'], birth_date))

        row = cursor.fetchone()
        if row:
            return row[0]

        cursor.execute("""
            INSERT INTO people (surname, name, patronymic, birth_date, gender, phone)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (data['surname'], data['name'], data['patronymic'],
            birth_date, data['gender'],
            data['phone_number'] if data['phone_number'] else None))
        return cursor.lastrowid

    def insert_person_extra_data(self, cursor, person_id, data):
        cursor.execute("SELECT id FROM person_extra_data WHERE person_id = %s", (person_id,))
        if cursor.fetchone():
            return
        cursor.execute("""
            INSERT INTO person_extra_data (person_id, social_status, address, native_language, nationality)
            VALUES (%s, %s, %s, %s, %s)
        """, (person_id,
              data['social_status'] if data['social_status'] else None,
              data['adress'] if data['adress'] else None,
              data['native_language'] if data['native_language'] else None,
              data['nationality'] if data['nationality'] else None))

    def insert_passport_data(self, cursor, person_id, data):
        cursor.execute("SELECT id FROM passport_data WHERE person_id = %s", (person_id,))
        if cursor.fetchone():
            return
        cursor.execute("""
            INSERT INTO passport_data (person_id, passport_series, passport_number, issuer,
                                       issue_date, subdivision_code, place_of_birth, address, citizenship)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (person_id,
              data['passport_series'] if data['passport_series'] else None,
              data['passport_number'] if data['passport_number'] else None,
              data['issuer'] if data['issuer'] else None,
              data['issue_date'],
              data['subdivision_code'] if data['subdivision_code'] else None,
              data['place_of_birth'] if data['place_of_birth'] else None,
              data['place_adress'] if data['place_adress'] else None,
              data['citizenship'] if data['citizenship'] else None))

    def insert_person_insurance(self, cursor, person_id, data):
        cursor.execute("SELECT id FROM person_insurance WHERE person_id = %s", (person_id,))
        if cursor.fetchone():
            return
        # Ищем страховую компанию в справочнике guides
        insurance_company_id = None
        if data['insurance_companie']:
            cursor.execute("""
                SELECT id FROM guides
                WHERE text LIKE %s AND category_id IN (SELECT id FROM guide_categories WHERE name = 'страховая компания')
                LIMIT 1
            """, (f"%{data['insurance_companie']}%",))
            row = cursor.fetchone()
            if row:
                insurance_company_id = row[0]

        cursor.execute("""
            INSERT INTO person_insurance (person_id, insurance_number, date_of_insurance,
                                          medical_policy, date_of_medical_policy, insurance_companies_id)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (person_id,
              data['insurance_number'] if data['insurance_number'] else None,
              data['date_of_insurance'],
              data['medical_policy'] if data['medical_policy'] else None,
              data['date_of_medical_policy'],
              insurance_company_id))

    def insert_person_certificate(self, cursor, person_id, data):
        cursor.execute("""
            SELECT id FROM person_certificates
            WHERE person_id = %s AND certificate_type_of_education = %s
        """, (person_id, data['type_of_education'] if data['type_of_education'] else None))
        if cursor.fetchone():
            return
        cursor.execute("""
            INSERT INTO person_certificates (person_id, certificate_type_of_education, average_score)
            VALUES (%s, %s, %s)
        """, (person_id,
              data['type_of_education'] if data['type_of_education'] else None,
              data['average_mark']))

    def insert_enrollment(self, cursor, person_id, data):
        # Ищем specialization_id по коду из Excel
        code = data.get('specialization_code', '')
        cursor.execute("SELECT id FROM specializations WHERE code = %s", (code,))
        row = cursor.fetchone()
        if not row:
            print(f"  Специальность с кодом '{code}' не найдена в БД")
            return
        spec_id = row[0]

        grade = data.get('grade', 9)
        is_budget = data.get('is_budget', 1)
        is_full_time = data.get('is_full_time', 1)

        # Ищем приёмную кампанию 2025 года
        cursor.execute("""
            SELECT id FROM receptions
            WHERE specialization_id = %s 
              AND grade = %s 
              AND is_budget = %s
              AND is_full_time = %s 
              AND year = 2025
            LIMIT 1
        """, (spec_id, grade, is_budget, is_full_time))
        row = cursor.fetchone()
        if not row:
            print(f"  Кампания не найдена: code={code}, spec_id={spec_id}, grade={grade}, is_budget={is_budget}, is_full_time={is_full_time}")
            return
        reception_id = row[0]

        # Статус
        status = 'зачислен' if data.get('is_adopted') == 1 else 'на рассмотрении'

        # Проверка дубликата
        cursor.execute("""
            SELECT id FROM enrollments
            WHERE reception_id = %s AND person_id = %s
        """, (reception_id, person_id))
        if cursor.fetchone():
            return

        cursor.execute("""
            INSERT INTO enrollments (reception_id, person_id, reg_number, is_priority, status)
            VALUES (%s, %s, %s, 0, %s)
        """, (reception_id, person_id,
              data.get('reg_number') if data.get('reg_number') else None,
              status))
        
    def open_match_window(self):
        """Открывает окно привязки учебных заведений"""
        match_win = tk.Toplevel(self.root)
        match_win.title("Привязка учебных заведений")
        match_win.geometry("1200x700")
        
        # ===== ЛЕВАЯ ЧАСТЬ =====
        left_frame = tk.Frame(match_win)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        tk.Label(left_frame, text="Студенты", font=("Arial", 12, "bold")).pack()
        
        # Фильтр по тексту
        tk.Label(left_frame, text="Поиск по фамилии/имени:").pack()
        self.filter_entry = tk.Entry(left_frame, width=30)
        self.filter_entry.pack(pady=2)
        self.filter_entry.bind("<KeyRelease>", lambda e: self.refresh_students())
        
        # Список студентов
        self.student_list = ttk.Treeview(left_frame, columns=("Фамилия", "Имя", "Отчество", "Рег.номер"),
                                         show="headings", height=25)
        self.student_list.heading("Фамилия", text="Фамилия")
        self.student_list.heading("Имя", text="Имя")
        self.student_list.heading("Отчество", text="Отчество")
        self.student_list.heading("Рег.номер", text="Рег.номер")
        self.student_list.column("Фамилия", width=120)
        self.student_list.column("Имя", width=100)
        self.student_list.column("Отчество", width=120)
        self.student_list.column("Рег.номер", width=100)
        self.student_list.pack(fill=tk.BOTH, expand=True)
        
        # ===== ПРАВАЯ ЧАСТЬ =====
        right_frame = tk.Frame(match_win)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        tk.Label(right_frame, text="Учебные заведения", font=("Arial", 12, "bold")).pack()
        
        # Фильтр по региону
        tk.Label(right_frame, text="Код региона:").pack()
        self.region_filter = tk.Entry(right_frame, width=30)
        self.region_filter.pack(pady=2)
        self.region_filter.bind("<KeyRelease>", lambda e: self.refresh_schools())
        
        # Фильтр по названию школы
        tk.Label(right_frame, text="Название школы:").pack()
        self.school_filter = tk.Entry(right_frame, width=30)
        self.school_filter.pack(pady=2)
        self.school_filter.bind("<KeyRelease>", lambda e: self.refresh_schools())
        
        # Список школ
        self.school_list = ttk.Treeview(right_frame, columns=("ID", "Название", "Регион", "Адрес"),
                                        show="headings", height=25)
        self.school_list.heading("ID", text="ID")
        self.school_list.heading("Название", text="Краткое название")
        self.school_list.heading("Регион", text="Регион")
        self.school_list.heading("Адрес", text="Адрес")
        self.school_list.column("ID", width=50)
        self.school_list.column("Название", width=200)
        self.school_list.column("Регион", width=120)
        self.school_list.column("Адрес", width=150)
        self.school_list.pack(fill=tk.BOTH, expand=True)
        
        # ===== КНОПКА ПРИВЯЗКИ =====
        btn_frame = tk.Frame(match_win)
        btn_frame.pack(side=tk.BOTTOM, pady=10)
        
        tk.Button(btn_frame, text="Привязать выделенное заведение к выделенному студенту",
                  command=self.bind_school_to_student, bg="#c8e6c9", width=50, height=2).pack()
        
        self.status_match = tk.Label(match_win, text="", font=("Arial", 10))
        self.status_match.pack(side=tk.BOTTOM)
        
        # Загружаем данные
        self.refresh_students()
        self.refresh_schools()
    
    def refresh_students(self):
        """Обновляет список студентов"""
        for item in self.student_list.get_children():
            self.student_list.delete(item)
        
        filter_text = self.filter_entry.get().strip()
        
        try:
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()
            
            query = """
                SELECT p.surname, p.name, p.patronymic, e.reg_number, p.id
                FROM people p
                LEFT JOIN enrollments e ON p.id = e.person_id
                WHERE p.id > 10
            """
            params = []
            if filter_text:
                query += " AND (p.surname LIKE %s OR p.name LIKE %s)"
                params.extend([f"%{filter_text}%", f"%{filter_text}%"])
            
            query += " ORDER BY p.surname LIMIT 500"
            
            cursor.execute(query, params)
            for row in cursor.fetchall():
                self.student_list.insert("", tk.END, values=(row[0], row[1], row[2], row[3]),
                                         iid=str(row[4]))  # iid = person_id
            
            cursor.close()
            conn.close()
        except Exception:
            pass
    
    def refresh_schools(self):
        """Обновляет список школ с фильтрами и нечётким поиском"""
        for item in self.school_list.get_children():
            self.school_list.delete(item)
        
        region_text = self.region_filter.get().strip()
        school_text = self.school_filter.get().strip()
        
        try:
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()
            
            query = """
                SELECT ec.id, ec.edu_org_short_name, r.name, ec.edu_org_address
                FROM educational_certificate ec
                JOIN regions r ON ec.region_id = r.id
                WHERE ec.status_name = 'Действующее'
            """
            params = []
            
            if region_text:
                query += " AND r.code LIKE %s"
                params.append(f"%{region_text}%")
            
            cursor.execute(query, params)
            all_schools = cursor.fetchall()
            
            cursor.close()
            conn.close()
            
            # Если есть текст для поиска — применяем нечёткое сопоставление
            if school_text:
                from difflib import SequenceMatcher
                import re
                
                def clean(s):
                    s = s.lower().strip()
                    s = re.sub(r'[«»""]', '', s)
                    s = re.sub(r'\s+', ' ', s)
                    return s
                
                scored = []
                clean_input = clean(school_text)
                for school in all_schools:
                    clean_short = clean(school[1] or '')
                    clean_full = clean(school[3] or '')
                    score = max(
                        SequenceMatcher(None, clean_input, clean_short).ratio(),
                        SequenceMatcher(None, clean_input, clean_full).ratio()
                    )
                    if score > 0.2:
                        scored.append((score, school))
                
                scored.sort(key=lambda x: x[0], reverse=True)
                all_schools = [s[1] for s in scored[:50]]
            
            for school in all_schools:
                self.school_list.insert("", tk.END, values=(school[0], school[1], school[2], school[3]),
                                        iid=str(school[0]))
            
        except Exception:
            pass
    
    def bind_school_to_student(self):
        """Привязывает студента к учебному заведению"""
        student_sel = self.student_list.selection()
        school_sel = self.school_list.selection()
        
        if not student_sel:
            self.status_match.config(text="Выберите студента из левого списка", fg="red")
            return
        if not school_sel:
            self.status_match.config(text="Выберите учебное заведение из правого списка", fg="red")
            return
        
        person_id = student_sel[0]
        school_id = school_sel[0]
        
        try:
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()
            
            # Получаем данные студента для заполнения полей
            cursor.execute("""
                SELECT birth_date FROM people WHERE id = %s
            """, (person_id,))
            person_row = cursor.fetchone()
            birth_date = person_row[0] if person_row else None
            
            # Проверяем, есть ли уже связь
            cursor.execute("""
                SELECT id FROM person_certificates
                WHERE person_id = %s AND edu_certificate_id = %s
            """, (person_id, school_id))
            
            if cursor.fetchone():
                # Обновляем существующую связь
                cursor.execute("""
                    UPDATE person_certificates
                    SET certificate_type_of_education = 'Аттестат о среднем общем образовании',
                        certificate_date = %s,
                        is_original = 0,
                        updated_at = NOW()
                    WHERE person_id = %s AND edu_certificate_id = %s
                """, (birth_date, person_id, school_id))
                msg = "Связь обновлена"
            else:
                # Создаём новую связь с заполнением всех полей
                cursor.execute("""
                    INSERT INTO person_certificates
                    (person_id, edu_certificate_id, certificate_type_of_education,
                     certificate_date, is_original, created_at, updated_at)
                    VALUES (%s, %s, 'Аттестат о среднем общем образовании',
                            %s, 0, NOW(), NOW())
                """, (person_id, school_id, birth_date))
                msg = "Связь создана"
            
            conn.commit()
            cursor.close()
            conn.close()
            
            student_name = f"{self.student_list.item(student_sel[0], 'values')[0]} {self.student_list.item(student_sel[0], 'values')[1]}"
            school_name = self.school_list.item(school_sel[0], "values")[1]
            
            self.status_match.config(
                text=f"{msg}: {student_name} → {school_name}", fg="green")
            
        except Exception as e:
            self.status_match.config(text=f"Ошибка: {e}", fg="red")

# ==================== ЗАПУСК ====================

if __name__ == "__main__":
    root = tk.Tk()
    app = MigrationApp(root)
    root.mainloop()