import json
import re
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk

import pandas as pd


class MyApp:

    def __init__(self, root):
        self.root = root
        self.root.geometry("900x500")
        self.root.title('Остатки')
        self.root.resizable(True, True)
        # Инициализация переменных экземпляра
        self.mol_list = ["--- Выберите МОЛ ---", ]
        self.sections_full = {}
        self.sections = {}
        self.current_section = None
        self.section_name = None
        self.df = None  # Добавляем инициализацию df

        self.create_menu()
        self.create_widgets()

    def create_widgets(self):
        frame_top = tk.Frame(self.root)
        frame_top.pack(padx=10, pady=10, fill='x')

        tk.Label(frame_top, text="МОЛ:", font=('Arial', 10)).pack(side='left')

        # Поле выбора МОЛ, с возможностью поиска
        self.mol_var = tk.StringVar()
        self.mol_combobox = ttk.Combobox(
            frame_top,
            textvariable=self.mol_var,
            state="normal",
            font=('Arial', 10),
            width=50
        )
        self.mol_combobox.pack(side='left', padx=5)
        self.mol_combobox.bind("<KeyRelease>", self.filter_mol_list)
        self.mol_combobox.bind("<<ComboboxSelected>>", self.on_select)

        # Создаём Treeview и Scrollbar
        self.tree = ttk.Treeview(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)

        # Позиционируем виджеты
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=5, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Связываем Treeview с Scrollbar
        self.tree.config(yscrollcommand=scrollbar.set)

        self.setup_columns()

        self.copy_button = tk.Button(frame_top, text="Копировать выделенные строки в буфер обмена", command=self.copy_selected)
        self.copy_button.pack(pady=5)

    def setup_columns(self):
        self.tree.delete(*self.tree.get_children())

        self.tree["columns"] = ("name", "price", "quantity", "total")

        self.tree.heading("#0", text="№п/п", anchor=tk.W)
        self.tree.column('#0', width=50, minwidth=50, stretch=tk.NO)

        self.tree.heading("name", text="Наименование", anchor=tk.W)
        self.tree.column("name", width=400, minwidth=200, stretch=tk.YES)

        self.tree.heading("price", text="Цена", anchor=tk.W)
        self.tree.column("price", width=100, minwidth=80, stretch=tk.NO)

        self.tree.heading("quantity", text="Количество", anchor=tk.W)
        self.tree.column("quantity", width=100, minwidth=80, stretch=tk.NO)

        self.tree.heading("total", text="Сумма", anchor=tk.W)
        self.tree.column("total", width=100, minwidth=80, stretch=tk.NO)

    def on_select(self, event ):
        selected_item = self.mol_combobox.get()
        if selected_item == "--- Выберите МОЛ ---":
            for item in self.tree.get_children():
                self.tree.delete(item)
            return
        self.section_name = selected_item

        self.create_file_json(self.sections_full, self.section_name)
        self.load_json('data.json')

    def update_combobox(self):
        """Обновление значений combobox"""
        self.mol_combobox['values'] = self.mol_list
        self.mol_combobox.set("")


    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Открыть", command=self.open_file_xls)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.quit)
        menubar.add_cascade(label="Файл", menu=file_menu)
        self.root.config(menu=menubar)

    def open_file_xls(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path, sheet_name='Лист_1', header=None)
            self.mol_list = ["--- Выберите МОЛ ---"] + self.create_mol_list(self.df)
            self.sections_full = self.get_sections_full(self.df, self.mol_list[1:])  # Исключаем первый элемент
            self.update_combobox()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def create_mol_list(self, df):
        """
        Функция получения списка фамилий, МОЛ
        :param df: данные
        :return: список
        """
        lst_ = []
        for index, row in df.iterrows():
            if pd.notna(row[0]) and isinstance(row[0], str):
                cell_value = row[0].strip()
                if len(cell_value.split()) == 3 and not re.search(r'\d', cell_value) and not re.search(r'[^\w\s]', cell_value):
                    lst_.append(cell_value)
        return lst_

    def filter_mol_list(self, event=None):
        typed = self.mol_var.get().lower()

        if typed == "":
            self.mol_combobox['values'] = self.mol_list
        else:
            filtered_values = [name for name in self.mol_list if typed in name.lower()]
            self.mol_combobox['values'] = filtered_values

    def copy_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Предупреждение", "Нет выделенных строк.")
            return

        # Получаем заголовки столбцов (для формирования строки с разделителями)
        columns = [col for col in self.tree["columns"]]

        # Формируем данные: каждая строка — это список значений по столбцам
        data_to_copy = []
        for item_id in selected_items:
            values = self.tree.item(item_id, "values")
            data_to_copy.append("\t\t".join(str(val) for val in values))

        # Добавляем заголовки и данные в текст для копирования
        # header = "\t".join(columns)
        # clipboard_text = f"{header}\n{'\n'.join(data_to_copy)}"
        clipboard_text = '\n'.join(data_to_copy)

        # Копируем в буфер обмена
        self.root.clipboard_clear()
        self.root.clipboard_append(clipboard_text)

        messagebox.showinfo("Успех", "Данные скопированы в буфер обмена.")

    def get_sections_full(self, df, mol_list):
        """
        Получение полного словаря словарей
        :param df:
        :param mol_list:
        :return:
        """
        sections = {}
        current_section = None
        for index, row in df.iterrows():
            if pd.notna(row[0]) and isinstance(row[0], str) and row[0].strip() in mol_list:
                current_section = row[0].strip()
                sections[current_section] = []
                sections[current_section].append(row)
                continue

            if current_section is not None and pd.notna(row[0]) and (
                    str(row[0]).strip() in ['105.31', '105.33', '105.35', '105.36', '1', '2.04']):
                continue
            if current_section is not None:
                sections[current_section].append(row)

        return sections

    def create_file_json(self, s_f, s_n):
        """
        Создание JSON-файла
        :param s_f: словарь с разделами
        :param s_n: название раздела
        :return: None
        """
        if s_n not in s_f:
            return

        dict_a = {}
        for index, value in enumerate(s_f[s_n]):
            # Пропускаем заголовки и пустые строки
            if index >= 2 and pd.notna(value[0]):
                dict_a[index - 1] = {
                    "0": str(value[0]) if pd.notna(value[0]) else "",
                    "1": float(value[1]) if pd.notna(value[1]) else 0.0,
                    "2": str(value[2]) if pd.notna(value[2]) else "",
                    "3": float(value[3]) if pd.notna(value[3]) else 0.0
                }

        with open('data.json', 'w', encoding='utf-8') as f:
            json.dump(dict_a, f, indent=4, ensure_ascii=False)


    def load_json(self, file_path):
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding="utf-8") as file:
                data = json.load(file)
            self.display_data(data)
        except Exception as e:
            messagebox.showerror("Error", f"Не удалось загрузить файл:\n{str(e)}")

    def display_data(self, data):
        self.tree.delete(*self.tree.get_children())

        if not isinstance(data, dict):
            tk.messagebox.showerror("Error", "JSON должен быть объектом")
            return

        for item_id, item_data in data.items():
            values = (
                item_data.get("0", ""),
                f"{float(item_data.get('1', 0)):.2f}",
                item_data.get("2", ""),
                f"{float(item_data.get('3', 0)):.2f}"
            )
            self.tree.insert("", "end", text=item_id, values=values)

