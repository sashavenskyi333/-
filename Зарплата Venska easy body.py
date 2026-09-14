# -*- coding: utf-8 -*-
"""
Created on Mon Sep 14 10:41:19 2026

@author: olve
"""

import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

# Налаштування теми
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class SalaryApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Система розрахунку зарплати салону")
        self.geometry("1000x600")
        self.minsize(950, 550)

        # Сховища даних
        self.employees_data = []
        self.individual_records = {}

        # Налаштування сітки головного вікна
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Ліва панель (Введення даних з вкладками) ---
        self.sidebar = ctk.CTkFrame(self, width=320, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(1, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar, text="Введення даних", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=15, pady=(15, 5))

        # Створення вкладок для вибору ролі
        self.tabview = ctk.CTkTabview(self.sidebar, width=290)
        self.tabview.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        
        self.tab_emp = self.tabview.add("Працівник")
        self.tab_admin = self.tabview.add("Адміністратор")

        # --- Форма: Працівник ---
        self.entry_emp_name = ctk.CTkEntry(self.tab_emp, placeholder_text="Ім'я працівника", width=250)
        self.entry_emp_name.pack(pady=10)

        self.entry_emp_percent = ctk.CTkEntry(self.tab_emp, placeholder_text="Відсоток (%) (напр. 30)", width=250)
        self.entry_emp_percent.pack(pady=10)

        self.entry_emp_procedures = ctk.CTkEntry(self.tab_emp, placeholder_text="Суми процедур через кому", width=250)
        self.entry_emp_procedures.pack(pady=10)

        self.btn_add_emp = ctk.CTkButton(self.tab_emp, text="Додати працівника", command=self.add_employee)
        self.btn_add_emp.pack(pady=20)

        # --- Форма: Адміністратор ---
        self.entry_adm_name = ctk.CTkEntry(self.tab_admin, placeholder_text="Ім'я адміністратора", width=250)
        self.entry_adm_name.pack(pady=8)

        self.entry_adm_base = ctk.CTkEntry(self.tab_admin, placeholder_text="Ставка (напр. 5000)", width=250)
        self.entry_adm_base.pack(pady=8)

        self.entry_adm_percent = ctk.CTkEntry(self.tab_admin, placeholder_text="Відсоток від оберту (%)", width=250)
        self.entry_adm_percent.pack(pady=8)

        self.entry_adm_turnover = ctk.CTkEntry(self.tab_admin, placeholder_text="Загальний оберт салону", width=250)
        self.entry_adm_turnover.pack(pady=8)

        self.btn_add_admin = ctk.CTkButton(self.tab_admin, text="Додати адміністратора", fg_color="#1f6aa5", command=self.add_admin)
        self.btn_add_admin.pack(pady=15)

        # --- Права панель (Таблиця та кнопки дій) ---
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=15, pady=15)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        self.label_title = ctk.CTkLabel(self.main_frame, text="Зведений список на виплату", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.grid(row=0, column=0, padx=15, pady=10, sticky="w")

        # Налаштування стилю таблиці
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", rowheight=28, fieldbackground="#2b2b2b", font=('Segoe UI', 10))
        style.map('Treeview', background=[('selected', '#1f538d')])
        style.configure("Treeview.Heading", background="#414547", foreground="white", font=('Segoe UI', 10, 'bold'))

        # Таблиця (без стовпчика "База")
        columns = ("Роль", "Ім'я", "Ставка / %", "Зарплата")
        self.tree = ttk.Treeview(self.main_frame, columns=columns, show="headings")
        
        # Налаштування ширини колонок
        col_widths = {"Роль": 150, "Ім'я": 220, "Ставка / %": 200, "Зарплата": 150}
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths[col], anchor="center")
            
        self.tree.grid(row=1, column=0, sticky="nsew", padx=15, pady=10)

        # Нижній блок кнопок
        self.bottom_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.bottom_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=10)

        self.btn_clear = ctk.CTkButton(self.bottom_frame, text="Очистити все", fg_color="#c93434", hover_color="#9e2a2a", command=self.clear_all)
        self.btn_clear.pack(side="left", padx=5)

        self.btn_export = ctk.CTkButton(self.bottom_frame, text="Зберегти в Excel", fg_color="#2ba84a", hover_color="#21823a", command=self.export_excel)
        self.btn_export.pack(side="right", padx=5)

    def add_employee(self):
        name = self.entry_emp_name.get().strip()
        percent_str = self.entry_emp_percent.get().replace(',', '.').strip()
        procedures_str = self.entry_emp_procedures.get().strip()

        if not name:
            messagebox.showwarning("Помилка", "Введіть ім'я працівника.")
            return

        try:
            percent = float(percent_str)
        except ValueError:
            messagebox.showwarning("Помилка", "Неправильний формат відсотка.")
            return

        try:
            procedures_list = [float(p.strip()) for p in procedures_str.split(',') if p.strip()]
            if not procedures_list:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Помилка", "Введіть суми процедур через кому (наприклад, 500, 250.5, 1000).")
            return

        total_sum = sum(procedures_list)
        salary = total_sum * (percent / 100)

        self.employees_data.append({
            "Роль": "Працівник",
            "Ім'я": name,
            "Ставка / %": f"{percent}%",
            "Зарплата до виплати": round(salary, 2),
            "Тип": "employee"
        })
        self.individual_records[name] = {"type": "employee", "procedures": procedures_list, "percent": percent}

        self.tree.insert("", "end", values=("Працівник", name, f"{percent}%", round(salary, 2)))

        # Очищення полів
        self.entry_emp_name.delete(0, 'end')
        self.entry_emp_percent.delete(0, 'end')
        self.entry_emp_procedures.delete(0, 'end')

    def add_admin(self):
        name = self.entry_adm_name.get().strip()
        base_str = self.entry_adm_base.get().replace(',', '.').strip()
        percent_str = self.entry_adm_percent.get().replace(',', '.').strip()
        turnover_str = self.entry_adm_turnover.get().replace(',', '.').strip()

        if not name:
            messagebox.showwarning("Помилка", "Введіть ім'я адміністратора.")
            return

        try:
            base_rate = float(base_str) if base_str else 0.0
            percent = float(percent_str) if percent_str else 0.0
            turnover = float(turnover_str) if turnover_str else 0.0
        except ValueError:
            messagebox.showwarning("Помилка", "Перевірте некоректні значення чисел (ставка, відсоток або оберт).")
            return

        bonus = turnover * (percent / 100)
        salary = base_rate + bonus
        details_str = f"{base_rate} + {percent}%"

        self.employees_data.append({
            "Роль": "Адміністратор",
            "Ім'я": name,
            "Ставка / %": details_str,
            "Зарплата до виплати": round(salary, 2),
            "Тип": "admin"
        })
        
        self.individual_records[name] = {
            "type": "admin",
            "base_rate": base_rate,
            "turnover": turnover,
            "percent": percent,
            "bonus": bonus
        }

        self.tree.insert("", "end", values=("Адміністратор", name, details_str, round(salary, 2)))

        # Очищення полів
        self.entry_adm_name.delete(0, 'end')
        self.entry_adm_base.delete(0, 'end')
        self.entry_adm_percent.delete(0, 'end')
        self.entry_adm_turnover.delete(0, 'end')

    def clear_all(self):
        if messagebox.askyesno("Підтвердження", "Ви дійсно хочете очистити всі дані?"):
            self.employees_data.clear()
            self.individual_records.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)

    def export_excel(self):
        if not self.employees_data:
            messagebox.showinfo("Увага", "Немає даних для експорту.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Зберегти звіт",
            initialfile="Зарплати_Звіт.xlsx"
        )

        if not file_path:
            return

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 1. Загальний лист (без службового поля "Тип")
                df_general = pd.DataFrame(self.employees_data)
                df_export = df_general.drop(columns=["Тип"])
                df_export.to_excel(writer, sheet_name="Загальна таблиця", index=False)

                # 2. Детальні листи по кожному співробітнику
                for name, info in self.individual_records.items():
                    safe_sheet_name = name[:31]
                    
                    if info["type"] == "employee":
                        procedures = info["procedures"]
                        df_ind = pd.DataFrame({
                            "Номер процедури": range(1, len(procedures) + 1),
                            "Сума за процедуру": procedures
                        })
                        df_ind.loc[len(df_ind)] = ["УСЬОГО (Сума):", sum(procedures)]
                        df_ind.loc[len(df_ind)] = ["ВІДСОТОК (%):", info["percent"]]
                        df_ind.loc[len(df_ind)] = ["ЗАРПЛАТА:", sum(procedures) * (info["percent"] / 100)]
                        df_ind.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                    
                    elif info["type"] == "admin":
                        df_ind = pd.DataFrame({
                            "Показник": ["Фіксована ставка", f"Бонус від оберту ({info['percent']}%)", "Загальний оберт салону", "РАЗОМ ДО ВИПЛАТИ"],
                            "Значення": [info["base_rate"], info["bonus"], info["turnover"], info["base_rate"] + info["bonus"]]
                        })
                        df_ind.to_excel(writer, sheet_name=safe_sheet_name, index=False)

            messagebox.showinfo("Успіх", f"Файл успішно збережено:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося зберегти файл:\n{e}")

if __name__ == "__main__":
    app = SalaryApp()
    app.mainloop()
