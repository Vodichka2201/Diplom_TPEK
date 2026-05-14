import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import mysql.connector
import re
from difflib import SequenceMatcher
import openpyxl
from datetime import datetime
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

# ==================== ФУНКЦИИ СОПОСТАВЛЕНИЯ ====================

def extract_school_number(name):
    match = re.search(r'(?:№|#|No\.?|школа)\s*(\d+)', name, re.IGNORECASE)
    return match.group(1) if match else None

def extract_city(name):
    match = re.search(r'г\.?\s*([А-Яа-яЁё\-]+)', name, re.IGNORECASE)
    return match.group(1) if match else None

def clean_name(name):
    name = name.lower().strip()
    name = re.sub(r'[«»""]', '', name)
    name = re.sub(r'\s+', ' ', name)
    return name

def find_best_match(cursor, dirty_name):
    number = extract_school_number(dirty_name)
    city = extract_city(dirty_name)
    clean_dirty = clean_name(dirty_name)
    
    query = """
        SELECT ec.id, ec.edu_org_full_name, ec.edu_org_short_name, r.name AS region
        FROM educational_certificate ec
        JOIN regions r ON ec.region_id = r.id
        WHERE ec.status_name = 'Действующее'
    """
    params = []
    
    if number:
        query += " AND (ec.edu_org_full_name LIKE %s OR ec.edu_org_short_name LIKE %s)"
        params.extend([f'%{number}%', f'%{number}%'])
    
    if city:
        query += " AND (ec.edu_org_address LIKE %s OR r.name LIKE %s)"
        params.extend([f'%{city}%', f'%{city}%'])
    
    cursor.execute(query, params)
    candidates = cursor.fetchall()
    
    if not candidates:
        return None
    
    best_score = 0
    best_match = None
    
    for school_id, full_name, short_name, region in candidates:
        score_full = SequenceMatcher(None, clean_dirty, clean_name(full_name or '')).ratio()
        score_short = SequenceMatcher(None, clean_dirty, clean_name(short_name or '')).ratio()
        score = max(score_full, score_short)
        
        if city and city.lower() in (full_name + ' ' + (short_name or '') + ' ' + region).lower():
            score += 0.1
        
        if score > best_score:
            best_score = score
            best_match = (school_id, full_name, short_name, region, score)
    
    return best_match


# ==================== КЛАСС ПРИЛОЖЕНИЯ ====================

class MigrationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Миграция данных студентов колледжа")
        self.root.geometry("1000x650")
        self.root.resizable(True, True)
        
        self.excel_data = []          # загруженные строки из Excel
        self.match_results = []       # результаты сопоставления
        self.db_connection = None
        
        self.create_widgets()
    
    def create_widgets(self):
        # Заголовок
        title = tk.Label(self.root, text="Миграция данных студентов",
                         font=("Arial", 16, "bold"))
        title.pack(pady=10)
        
        # Фрейм с кнопками
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=5)
        
        self.btn_load = tk.Button(btn_frame, text="1. Выбрать Excel-файл",
                                  command=self.load_excel, width=22, bg="#e0e0e0")
        self.btn_load.pack(side=tk.LEFT, padx=5)
        
        self.btn_match = tk.Button(btn_frame, text="2. Сопоставить учебные заведения",
                                   command=self.match_schools, width=26, bg="#e0e0e0",
                                   state=tk.DISABLED)
        self.btn_match.pack(side=tk.LEFT, padx=5)
        
        self.btn_save = tk.Button(btn_frame, text="3. Сохранить в БД",
                                  command=self.save_to_db, width=18, bg="#c8e6c9",
                                  state=tk.DISABLED)
        self.btn_save.pack(side=tk.LEFT, padx=5)
        
        # Таблица с данными
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("№", "Фамилия", "Имя", "Отчество", "Дата рождения",
                   "Школа (из Excel)", "Сопоставленное заведение", "Совпадение, %")
        
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)
        
        # Настройка колонок
        widths = [40, 120, 100, 120, 100, 200, 250, 100]
        for col, w in zip(columns, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w)
        
        # Прокрутка
        scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Статус-бар
        self.status = tk.Label(self.root, text="Готов к работе. Выберите Excel-файл.",
                               bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 10))
        self.status.pack(side=tk.BOTTOM, fill=tk.X)
    
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
            for row in ws.iter_rows(min_row=2, values_only=True):  # пропускаем заголовок
                if row[0] is None:
                    continue
                # Предполагаем: Фамилия, Имя, Отчество, Дата рождения, Школа
                surname = str(row[0]) if row[0] else ""
                name = str(row[1]) if row[1] else ""
                patronymic = str(row[2]) if len(row) > 2 and row[2] else ""
                birth_date = str(row[3]) if len(row) > 3 and row[3] else ""
                school = str(row[4]) if len(row) > 4 and row[4] else ""
                
                self.excel_data.append({
                    'surname': surname,
                    'name': name,
                    'patronymic': patronymic,
                    'birth_date': birth_date,
                    'school': school
                })
            
            # Очищаем таблицу
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Заполняем таблицу
            for i, data in enumerate(self.excel_data, 1):
                self.tree.insert("", tk.END, values=(
                    i, data['surname'], data['name'], data['patronymic'],
                    data['birth_date'], data['school'], "", ""
                ))
            
            self.status.config(text=f"Загружено {len(self.excel_data)} записей из файла: {filepath}")
            self.btn_match.config(state=tk.NORMAL)
            self.match_results = []
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")
    
    def match_schools(self):
        if not self.excel_data:
            messagebox.showwarning("Внимание", "Сначала загрузите Excel-файл.")
            return
        
        try:
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()
            
            self.match_results = []
            
            for data in self.excel_data:
                school_name = data['school']
                if not school_name:
                    self.match_results.append(None)
                    continue
                
                match = find_best_match(cursor, school_name)
                self.match_results.append(match)
            
            cursor.close()
            conn.close()
            
            # Обновляем таблицу
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            matched_count = 0
            for i, (data, match) in enumerate(zip(self.excel_data, self.match_results), 1):
                if match:
                    school_id, full_name, short_name, region, confidence = match
                    matched_name = short_name or full_name
                    confidence_str = f"{confidence * 100:.0f}%"
                    matched_count += 1
                else:
                    matched_name = "НЕ НАЙДЕНО"
                    confidence_str = "0%"
                
                self.tree.insert("", tk.END, values=(
                    i, data['surname'], data['name'], data['patronymic'],
                    data['birth_date'], data['school'], matched_name, confidence_str
                ))
            
            self.status.config(
                text=f"Сопоставление завершено. Найдено: {matched_count} из {len(self.excel_data)}. "
                     f"Проверьте результаты перед сохранением."
            )
            self.btn_save.config(state=tk.NORMAL)
            
        except mysql.connector.Error as e:
            messagebox.showerror("Ошибка БД", f"Не удалось подключиться к базе данных:\n{e}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сопоставлении:\n{e}")
    
    def save_to_db(self):
        if not self.match_results:
            messagebox.showwarning("Внимание", "Сначала выполните сопоставление.")
            return
        
        # Подтверждение
        ok = messagebox.askyesno(
            "Подтверждение",
            "Сохранить результаты миграции в базу данных?\n\n"
            "Будут созданы записи в таблицах people и person_certificates."
        )
        if not ok:
            return
        
        try:
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()
            
            saved = 0
            skipped = 0
            
            for data, match in zip(self.excel_data, self.match_results):
                if match is None:
                    skipped += 1
                    continue
                
                school_id = match[0]
                
                # Вставка в people
                cursor.execute("""
                    INSERT IGNORE INTO people (surname, name, patronymic, birth_date, gender)
                    VALUES (%s, %s, %s, %s, 'мужской')
                """, (data['surname'], data['name'], data['patronymic'],
                      data['birth_date'] if data['birth_date'] else None))
                
                # Получаем id человека
                cursor.execute("""
                    SELECT id FROM people
                    WHERE surname = %s AND name = %s AND patronymic = %s AND birth_date = %s
                """, (data['surname'], data['name'], data['patronymic'],
                      data['birth_date'] if data['birth_date'] else None))
                
                person_row = cursor.fetchone()
                if person_row:
                    person_id = person_row[0]
                    
                    # Вставка в person_certificates
                    cursor.execute("""
                        INSERT IGNORE INTO person_certificates
                        (person_id, edu_certificate_id, certificate_type_of_education)
                        VALUES (%s, %s, 'Аттестат')
                    """, (person_id, school_id))
                    
                    saved += 1
            
            conn.commit()
            cursor.close()
            conn.close()
            
            self.status.config(
                text=f"Сохранение завершено. Сохранено: {saved}, пропущено: {skipped}."
            )
            messagebox.showinfo(
                "Готово",
                f"Результаты миграции сохранены в базу данных.\n\n"
                f"Успешно сохранено: {saved}\n"
                f"Пропущено (не найдено совпадений): {skipped}"
            )
            
        except mysql.connector.Error as e:
            messagebox.showerror("Ошибка БД", f"Ошибка при сохранении:\n{e}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка:\n{e}")


# ==================== ЗАПУСК ====================

if __name__ == "__main__":
    root = tk.Tk()
    app = MigrationApp(root)
    root.mainloop()