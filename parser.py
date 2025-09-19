import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class ParserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Парсер TXT/Excel → Excel")
        self.root.geometry("900x650")
        self.root.resizable(True, True)

        self.manual_entries = []        # Поля ручного ввода
        self.file_values_vars = []      # (значение, tk.BooleanVar)
        self.search_files_vars = []     # (путь_к_файлу, tk.BooleanVar)

        # ==== ВВОД ЗНАЧЕНИЙ ====
        tk.Label(root, text="Введите значения для поиска:", font=("Arial", 12, "bold")).pack(pady=5)
        self.frame_inputs = tk.Frame(root)
        self.frame_inputs.pack()
        self.add_entry()
        tk.Button(root, text="Добавить поле", command=self.add_entry).pack(pady=3)

        tk.Button(root, text="Загрузить TXT со значениями",
                  command=self.load_value_file).pack(pady=8)

        # === Прокрутка для значений ===
        tk.Label(root, text="Значения (отметьте нужные):", font=("Arial", 11, "italic")).pack()
        self.values_canvas, self.check_frame = self.make_scrollable(height=150)

        # === Прокрутка для файлов ===
        tk.Label(root, text="Выбранные файлы (отметьте нужные):", font=("Arial", 11, "italic")).pack(pady=5)
        self.files_canvas, self.files_frame = self.make_scrollable(height=150)

        # === Кнопки выбора файлов и запуска ===
        bottom = tk.Frame(root)
        bottom.pack(fill="x", pady=10)
        tk.Button(bottom, text="Добавить файлы", command=self.choose_files).pack(side="left", padx=10)
        tk.Button(bottom, text="Начать поиск", command=self.process).pack(side="right", padx=10)

    # ---------- Вспомогательные ----------
    def make_scrollable(self, height=180):
        container = tk.Frame(self.root)
        container.pack(fill="x", padx=10)
        canvas = tk.Canvas(container, height=height)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return canvas, frame

    def add_entry(self):
        entry = tk.Entry(self.frame_inputs, width=60)
        entry.pack(pady=2)
        self.manual_entries.append(entry)

    def load_value_file(self):
        path = filedialog.askopenfilename(title="Выберите TXT со значениями",
                                          filetypes=[("Text files", "*.txt")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                values = [line.strip() for line in f if line.strip()]
        except Exception as e:
            messagebox.showerror("Ошибка чтения", str(e))
            return

        # Очищаем старые чекбоксы
        for w in self.check_frame.winfo_children():
            w.destroy()
        self.file_values_vars.clear()

        for val in values:
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(self.check_frame, text=val, variable=var, anchor="w", justify="left")
            cb.pack(fill="x", padx=5)
            self.file_values_vars.append((val, var))

    def choose_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите TXT/XLSX файлы",
            filetypes=[("TXT or Excel", "*.txt *.xlsx")])
        if not files:
            return
        # добавляем только новые
        existing = {p for p, _ in self.search_files_vars}
        for f in files:
            if f not in existing:
                var = tk.BooleanVar(value=True)
                row = tk.Frame(self.files_frame)
                cb = tk.Checkbutton(row, text=os.path.basename(f), variable=var, anchor="w", width=60, justify="left")
                cb.pack(side="left", fill="x")
                rm = tk.Button(row, text="X", width=3,
                               command=lambda p=f, r=row: self.remove_file(p, r))
                rm.pack(side="right", padx=3)
                row.pack(fill="x", pady=1)
                self.search_files_vars.append((f, var))

    def remove_file(self, path, row_widget):
        self.search_files_vars = [(p, v) for p, v in self.search_files_vars if p != path]
        row_widget.destroy()

    # ---------- Основной процесс ----------
    def process(self):
        manual_values = [e.get().strip() for e in self.manual_entries if e.get().strip()]
        checked_values = [val for val, var in self.file_values_vars if var.get()]
        search_values = list(set(manual_values + checked_values))
        if not search_values:
            messagebox.showwarning("Нет значений", "Выберите или введите хотя бы одно значение!")
            return

        selected_files = [p for p, v in self.search_files_vars if v.get()]
        if not selected_files:
            messagebox.showwarning("Нет файлов", "Отметьте хотя бы один файл!")
            return

        all_results = []
        for file_path in selected_files:
            ext = os.path.splitext(file_path)[1].lower()
            try:
                if ext == ".txt":
                    with open(file_path, "r", encoding="utf-8") as f:
                        text = f.read()
                else:
                    excel = pd.read_excel(file_path, sheet_name=None, dtype=str)
                    text = "\n".join(df.to_string() for df in excel.values())
            except Exception as e:
                messagebox.showerror("Ошибка чтения", f"{file_path}\n{e}")
                continue

            for val in search_values:
                count = len(re.findall(re.escape(val), text, flags=re.IGNORECASE))
                all_results.append({"Файл": os.path.basename(file_path),
                                    "Значение": val,
                                    "Количество": count})

        if not all_results:
            messagebox.showwarning("Пусто", "Нет данных для записи.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Сохранить результат как",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="result.xlsx"
        )
        if not save_path:
            return

        try:
            pd.DataFrame(all_results).to_excel(save_path, index=False)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))
            return

        messagebox.showinfo("Готово!", f"Файл сохранён:\n{save_path}")

if __name__ == "__main__":
    root = tk.Tk()
    ParserApp(root)
    root.mainloop()
