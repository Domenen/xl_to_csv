<<<<<<< HEAD
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar
import threading

class ExcelToCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер Excel в CSV")
        self.root.geometry("600x500")
        
        # Переменные
        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.progress = tk.DoubleVar()
        self.progress.set(0)
        self.delimiter = StringVar(value=",")
        self.encoding = StringVar(value="utf-8")
        
        # Создание интерфейса
        self.create_widgets()
    
    def create_widgets(self):
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=5)
        
        # Основной контейнер
        main_frame = ttk.Frame(self.root)
        main_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Фрейм для выбора файла
        file_frame = ttk.LabelFrame(main_frame, text="1. Выберите Excel файл", padding=10)
        file_frame.pack(pady=5, fill="x")
        
        ttk.Label(file_frame, textvariable=self.input_file, relief="sunken", width=60).pack(pady=5)
        ttk.Button(file_frame, text="Выбрать файл", command=self.select_file).pack(pady=5)
        
        # Фрейм для выбора директории
        dir_frame = ttk.LabelFrame(main_frame, text="2. Выберите место сохранения", padding=10)
        dir_frame.pack(pady=5, fill="x")
        
        ttk.Label(dir_frame, textvariable=self.output_dir, relief="sunken", width=60).pack(pady=5)
        ttk.Button(dir_frame, text="Выбрать папку", command=self.select_output_dir).pack(pady=5)
        
        # Фрейм настроек CSV
        settings_frame = ttk.LabelFrame(main_frame, text="3. Настройки CSV", padding=10)
        settings_frame.pack(pady=5, fill="x")
        
        # Выбор разделителя
        ttk.Label(settings_frame, text="Разделитель:").grid(row=0, column=0, sticky="w", padx=5)
        delimiter_frame = ttk.Frame(settings_frame)
        delimiter_frame.grid(row=0, column=1, sticky="w")
        
        delimiters = [
            ("Запятая (,)", ","),
            ("Точка с запятой (;)", ";"),
            ("Табуляция (\\t)", "\t"),
            ("Вертикальная черта (|)", "|"),
            ("Другой", "other")
        ]
        
        for i, (text, val) in enumerate(delimiters):
            rb = ttk.Radiobutton(
                delimiter_frame, 
                text=text, 
                variable=self.delimiter, 
                value=val
            )
            rb.pack(side="left", padx=5)
        
        self.custom_delimiter = ttk.Entry(delimiter_frame, width=3)
        self.custom_delimiter.pack(side="left", padx=5)
        self.custom_delimiter.insert(0, ",")
        
        # Кодировка
        ttk.Label(settings_frame, text="Кодировка:").grid(row=1, column=0, sticky="w", padx=5)
        encodings = ["utf-8", "windows-1251", "cp1251", "ascii"]
        encoding_menu = ttk.Combobox(settings_frame, textvariable=self.encoding, values=encodings, width=15)
        encoding_menu.grid(row=1, column=1, sticky="w", pady=5)
        
        # Прогресс-бар
        progress_frame = ttk.LabelFrame(main_frame, text="Прогресс", padding=10)
        progress_frame.pack(pady=10, fill="x")
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress, maximum=100)
        self.progress_bar.pack(fill="x")
        
        # Кнопка конвертации
        ttk.Button(main_frame, text="Начать конвертацию", command=self.start_conversion).pack(pady=10)
        
        # Отслеживаем выбор "Другой" разделитель
        self.delimiter.trace_add("write", self.on_delimiter_change)
    
    def on_delimiter_change(self, *args):
        if self.delimiter.get() == "other":
            self.custom_delimiter.config(state="normal")
        else:
            self.custom_delimiter.config(state="disabled")
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.input_file.set(file_path)
    
    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title="Выберите папку для сохранения")
        if dir_path:
            self.output_dir.set(dir_path)
    
    def start_conversion(self):
        if not self.input_file.get():
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл Excel")
            return
        
        delimiter = self.custom_delimiter.get() if self.delimiter.get() == "other" else self.delimiter.get()
        
        output_dir = self.output_dir.get() or str(Path(self.input_file.get()).parent)
        output_file = Path(output_dir) / (Path(self.input_file.get()).stem + ".csv")
        
        threading.Thread(
            target=self.convert_file,
            args=(self.input_file.get(), str(output_file), delimiter, self.encoding.get()),
            daemon=True
        ).start()
    
    def convert_file(self, input_path, output_path, delimiter, encoding):
        try:
            self.progress.set(10)
            
            df = pd.read_excel(input_path)
            self.progress.set(50)
            
            Path(output_path).parent.mkdir(parents=True, exist_ok=True)
            df.to_csv(
                output_path, 
                index=False, 
                sep=delimiter,
                encoding=encoding
            )
            
            self.progress.set(100)
            messagebox.showinfo("Успех", f"Файл успешно сохранен:\n{output_path}")
        
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
        finally:
            self.progress.set(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToCSVConverter(root)
=======
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar
import threading

class ExcelToCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер Excel в CSV")
        self.root.geometry("600x500")
        
        # Переменные
        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.progress = tk.DoubleVar()
        self.progress.set(0)
        self.delimiter = StringVar(value=",")
        self.encoding = StringVar(value="utf-8")
        
        # Создание интерфейса
        self.create_widgets()
    
    def create_widgets(self):
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=5)
        
        # Основной контейнер
        main_frame = ttk.Frame(self.root)
        main_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Фрейм для выбора файла
        file_frame = ttk.LabelFrame(main_frame, text="1. Выберите Excel файл", padding=10)
        file_frame.pack(pady=5, fill="x")
        
        ttk.Label(file_frame, textvariable=self.input_file, relief="sunken", width=60).pack(pady=5)
        ttk.Button(file_frame, text="Выбрать файл", command=self.select_file).pack(pady=5)
        
        # Фрейм для выбора директории
        dir_frame = ttk.LabelFrame(main_frame, text="2. Выберите место сохранения", padding=10)
        dir_frame.pack(pady=5, fill="x")
        
        ttk.Label(dir_frame, textvariable=self.output_dir, relief="sunken", width=60).pack(pady=5)
        ttk.Button(dir_frame, text="Выбрать папку", command=self.select_output_dir).pack(pady=5)
        
        # Фрейм настроек CSV
        settings_frame = ttk.LabelFrame(main_frame, text="3. Настройки CSV", padding=10)
        settings_frame.pack(pady=5, fill="x")
        
        # Выбор разделителя
        ttk.Label(settings_frame, text="Разделитель:").grid(row=0, column=0, sticky="w", padx=5)
        delimiter_frame = ttk.Frame(settings_frame)
        delimiter_frame.grid(row=0, column=1, sticky="w")
        
        delimiters = [
            ("Запятая (,)", ","),
            ("Точка с запятой (;)", ";"),
            ("Табуляция (\\t)", "\t"),
            ("Вертикальная черта (|)", "|"),
            ("Другой", "other")
        ]
        
        for i, (text, val) in enumerate(delimiters):
            rb = ttk.Radiobutton(
                delimiter_frame, 
                text=text, 
                variable=self.delimiter, 
                value=val
            )
            rb.pack(side="left", padx=5)
        
        self.custom_delimiter = ttk.Entry(delimiter_frame, width=3)
        self.custom_delimiter.pack(side="left", padx=5)
        self.custom_delimiter.insert(0, ",")
        
        # Кодировка
        ttk.Label(settings_frame, text="Кодировка:").grid(row=1, column=0, sticky="w", padx=5)
        encodings = ["utf-8", "windows-1251", "cp1251", "ascii"]
        encoding_menu = ttk.Combobox(settings_frame, textvariable=self.encoding, values=encodings, width=15)
        encoding_menu.grid(row=1, column=1, sticky="w", pady=5)
        
        # Прогресс-бар
        progress_frame = ttk.LabelFrame(main_frame, text="Прогресс", padding=10)
        progress_frame.pack(pady=10, fill="x")
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress, maximum=100)
        self.progress_bar.pack(fill="x")
        
        # Кнопка конвертации
        ttk.Button(main_frame, text="Начать конвертацию", command=self.start_conversion).pack(pady=10)
        
        # Отслеживаем выбор "Другой" разделитель
        self.delimiter.trace_add("write", self.on_delimiter_change)
    
    def on_delimiter_change(self, *args):
        if self.delimiter.get() == "other":
            self.custom_delimiter.config(state="normal")
        else:
            self.custom_delimiter.config(state="disabled")
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.input_file.set(file_path)
    
    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title="Выберите папку для сохранения")
        if dir_path:
            self.output_dir.set(dir_path)
    
    def start_conversion(self):
        if not self.input_file.get():
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл Excel")
            return
        
        delimiter = self.custom_delimiter.get() if self.delimiter.get() == "other" else self.delimiter.get()
        
        output_dir = self.output_dir.get() or str(Path(self.input_file.get()).parent)
        output_file = Path(output_dir) / (Path(self.input_file.get()).stem + ".csv")
        
        threading.Thread(
            target=self.convert_file,
            args=(self.input_file.get(), str(output_file), delimiter, self.encoding.get()),
            daemon=True
        ).start()
    
    def convert_file(self, input_path, output_path, delimiter, encoding):
        try:
            self.progress.set(10)
            
            df = pd.read_excel(input_path)
            self.progress.set(50)
            
            Path(output_path).parent.mkdir(parents=True, exist_ok=True)
            df.to_csv(
                output_path, 
                index=False, 
                sep=delimiter,
                encoding=encoding
            )
            
            self.progress.set(100)
            messagebox.showinfo("Успех", f"Файл успешно сохранен:\n{output_path}")
        
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
        finally:
            self.progress.set(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToCSVConverter(root)
>>>>>>> 91313f4f66b9012cd2ea091a542ccbcc6836901b
    root.mainloop()