"""
Модуль основной программы: Графический интерфейс для обработки файлов Excel, Word, DWG и SHA.

Это приложение предоставляет GUI на базе Tkinter для выбора проекта, блока, входной/выходной папок
и запуска обработки файлов с заменой текста по правилам из config.json. Поддерживает логирование,
переименование файлов и интеграцию с парсерами для различных форматов.

Зависимости: os, sys, re, glob, tkinter, datetime, PIL, json, shutil, а также импорт парсеров
(excel_parser, word_parser, dwg_parser, sha_parser).

Программа предназначена для автоматизации задач в контексте АЭС (атомных электростанций).
"""
import os, sys
import re
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from datetime import datetime
from excel_parser import ExcelProcessor
from word_parser import WordProcessor
from dwg_parser import AutoCADProcessor
#from pdf_parser import PdfProcessor
from PIL import Image, ImageTk
import json
import shutil

class FileProcessorGUI:
    """Класс для графического интерфейса обработки файлов.

        Создаёт окно Tkinter с элементами для выбора проекта, блока, папок и запуска обработки.
        Загружает конфигурацию из JSON, обрабатывает файлы через специализированные парсеры,
        логирует процесс в файл и GUI.

        Attributes:
            root (tk.Tk): Корневое окно Tkinter.
            replacement_digit (tk.StringVar): Переменная для выбранной цифры замены.
            project (tk.StringVar): Переменная для выбранного проекта.
            input_dir (tk.StringVar): Путь к входной папке.
            output_dir (tk.StringVar): Путь к выходной папке.
            debug_logging (tk.BooleanVar): Флаг отладочного логирования.
            log_file (file): Открытый файл для логирования (или None).
            config_data (dict): Загруженная конфигурация из JSON.
            lbl_project (tk.Label): Метка для отображения текущего проекта.
        """
    def __init__(self, root):
        """Инициализирует GUI и атрибуты.

            Устанавливает заголовок, размер окна, переменные Tkinter, загружает конфиг
            и вызывает создание виджетов.
            Args:
                root (tk.Tk): Корневое окно приложения.
        """
        self.root = root
        self.root.title("Обработчик Excel, Word, DWG и SHA для АЭС")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        self.replacement_digit = tk.StringVar()
        self.project = tk.StringVar()
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.debug_logging = tk.BooleanVar(value=False)
        self.log_file = None
        self.config_data = self._load_config()
        self.lbl_project = None
        self.create_widgets()

    def resource_path(self, relative_path):
        """Возвращает абсолютный путь к ресурсу, учитывая упаковку PyInstaller.

        Если приложение упаковано (--onefile), использует _MEIPASS; иначе — текущую директорию.

        Args:
            relative_path (str): Относительный путь к ресурсу (например, 'config.json').

        Returns:
            str: Абсолютный путь к ресурсу.
        """
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    def _load_config(self):
        config_path = self.resource_path('config.json')
        if not os.path.exists(config_path):
            messagebox.showerror("Ошибка", "Не удалось найти config.json в ресурсах")
            sys.exit(1)
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить config.json: {e}")
            sys.exit(1)

    def create_widgets(self):
        tk.Label(self.root, text="Выберите проект:").pack(anchor="w", padx=10, pady=5)
        projects = list(self.config_data.keys())
        self.project_combobox = ttk.Combobox(self.root, textvariable=self.project, values=projects, state="readonly")
        self.project_combobox.pack(anchor="w", padx=10, ipadx=50)
        self.project_combobox.bind("<<ComboboxSelected>>", self.update_digits)
        if projects:
            self.project.set(projects[0])

        tk.Label(self.root, text="Выберите Блок:").pack(anchor="w", padx=10, pady=5)

        frame_right = tk.Frame(self.root, width=100, height=120, bg="SystemButtonFace")
        frame_right.place(x=400, y=30)

        self.lbl_project = tk.Label(frame_right, text=self.project.get(), font=("Arial", 14, "italic"))
        self.lbl_project.pack(expand=True)

        self.frame_digits = tk.Frame(self.root)
        self.frame_digits.pack(anchor="w", padx=10)
        self.update_digits()

        tk.Label(self.root, text="Папка с исходными файлами:").pack(anchor="w", padx=10, pady=5)
        frame_in = tk.Frame(self.root)
        frame_in.pack(anchor="w", padx=10)
        tk.Entry(frame_in, textvariable=self.input_dir, width=50).pack(side="left")
        tk.Button(frame_in, text="Выбрать...", command=self.choose_input_dir).pack(side="left", padx=5)

        tk.Label(self.root, text="Папка для сохранения новых файлов:").pack(anchor="w", padx=10, pady=5)
        frame_out = tk.Frame(self.root)
        frame_out.pack(anchor="w", padx=10)
        tk.Entry(frame_out, textvariable=self.output_dir, width=50).pack(side="left")
        tk.Button(frame_out, text="Выбрать...", command=self.choose_output_dir).pack(side="left", padx=5)

        tk.Checkbutton(self.root, text="Отладочные логи", variable=self.debug_logging).pack(anchor="w", padx=10, pady=5)
        btn_run = tk.Button(frame_right, text="Запустить обработку",
                            command=self.run_processing,
                            bg="green", fg="white", font=("Arial", 11), padx=50, pady=5)
        btn_run.pack(pady=50)

        tk.Label(self.root, text="Процесс обработки:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self.root, width=80, height=5)
        self.log_text.pack(padx=10, pady=5, fill="both", expand=True)
        self.log_text.tag_configure("error", foreground="red", font=("Arial", 10, "bold"))
        self.log_text.tag_configure("skip", foreground="red", font=("Arial", 10, "bold"))

        tk.Label(self.root, text="by Артем Баюшкин", font=("Arial", 9, "italic")).pack(anchor="e", padx=10, pady=5)
        tk.Button(self.root, text="О программе", command=self.show_about).pack(anchor="e", padx=10, pady=5)
        tk.Button(self.root, text="Редактировать config", command=self.edit_config).pack(anchor="e", padx=10, pady=5)

    def update_digits(self, event=None):
        project = self.project.get()
        if self.lbl_project:
            self.lbl_project.config(text=project)
        digits = self.config_data.get(project, {}).get("digits", [])
        for widget in self.frame_digits.winfo_children():
            widget.destroy()
        if digits:
            self.replacement_digit.set(digits[0][1])
            for text, value in digits:
                tk.Radiobutton(self.frame_digits, text=text, variable=self.replacement_digit,
                               value=value).pack(side="left", padx=5)

    def choose_input_dir(self):
        folder = filedialog.askdirectory(title="Выберите папку с исходными файлами")
        if folder:
            self.input_dir.set(folder)

    def choose_output_dir(self):
        folder = filedialog.askdirectory(title="Выберите папку для сохранения")
        if folder:
            self.output_dir.set(folder)

    def edit_config(self):
        config_path = self.resource_path('config.json')
        if not os.path.exists(config_path):
            src = self.resource_path('config.json')
            if os.path.exists(src):
                shutil.copy(src, config_path)
            else:
                messagebox.showerror("Ошибка", "Не удалось найти config.json")
                return
        os.startfile(config_path)
        print('Файл config обновлен')

    def log_to_file(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        if self.log_file:
            try:
                self.log_file.write(log_message + "\n")
                self.log_file.flush()
            except Exception as e:
                self.log_to_gui(f"Ошибка записи в лог-файл: {str(e)}")

    def log_to_gui(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        if (message.startswith("=== Запуск обработки ===") or
            message.startswith("Успешно: ") or
            message.startswith("Ошибка обработки: ") or
            message.startswith("Обработка завершена. ")):
            tag = "error" if message.startswith("Ошибка обработки: ") else None
            self.log_text.insert(tk.END, log_message + "\n", tag)
            self.log_text.see(tk.END)
            self.root.update()

    def log(self, message):
        """
        Данная функция предназначена для задания параметров записи логов в файл log.txt
        """
        always_log = (
            message.startswith("=== Запуск обработки ===") or
            message.startswith("Успешно: ") or
            message.startswith("Ошибка обработки: ") or
            message.startswith("Обработка завершена. ") or
            message.startswith("Результаты сохранены в: ")
        )
        if self.debug_logging.get() or always_log:
            self.log_to_file(message)
        self.log_to_gui(message)

    def select_files(self, input_dir):
        '''
        Метод для загрузки файлов различного расширения
        :param input_dir:
        :return:
        '''
        excel_files = glob.glob(os.path.join(input_dir, '*.xls*'))
        word_files = glob.glob(os.path.join(input_dir, '*.do*'))
        dwg_files = glob.glob(os.path.join(input_dir, '*.dwg'))
        sha_files = glob.glob(os.path.join(input_dir, '*.sha'))
        #pdf_files = glob.glob(os.path.join(input_dir, '*.pdf'))
        files = excel_files + word_files + dwg_files + sha_files # + pdf_files
        return files

    def process_files(self, input_files, output_dir, replacement_digit):
        os.makedirs(output_dir, exist_ok=True)
        processed = 0
        project = self.project.get()
        sha_processor = None
        sha_app_started = False

        try:
            for input_path in input_files:
                try:
                    filename = os.path.basename(input_path)
                    name, ext = os.path.splitext(filename)

                    new_name = name
                    file_rename_rules = self.config_data.get(project, {}).get("file_rename", {})
                    for rule_name, rule in file_rename_rules.items():
                        try:
                            pattern = eval(rule["pattern"], {"re": re})
                            repl_str = rule["replacement"]
                            repl = eval(repl_str, {"replacement_digit": replacement_digit})
                            new_name = pattern.sub(repl, new_name)
                        except Exception as e:
                            self.log(
                                f"Ошибка при применении правила переименования {rule_name} для {filename}: {str(e)}")
                    if new_name == name:
                        new_name = f"processed_{name}"

                    if ext == ".xls":
                        output_path = os.path.join(output_dir, new_name + ".xlsm")
                    else:
                        output_path = os.path.join(output_dir, new_name + ext)
                    extension = ext.lower()

                    if extension in ('.doc', '.docx', '.dotx'):
                        rules = self.config_data.get(project, {}).get("word_parser", {})
                        processor = WordProcessor(replacement_digit, project, rules, log_callback=self.log,
                                                  debug=self.debug_logging.get())
                        success = processor.process_file(input_path, output_path)
                        if success:
                            self.log(f"Успешно: {filename}")
                            processed += 1
                        else:
                            self.log(f"Ошибка обработки: {filename}")

                    elif extension in ('.xls', '.xlsx', '.xlsm'):
                        rules = self.config_data.get(project, {}).get("excel_parser", {})
                        processor = ExcelProcessor(replacement_digit, project, rules, log_callback=self.log,
                                                   debug=self.debug_logging.get())
                        success = processor.process_file(input_path, output_path)
                        if success:
                            self.log(f"Успешно: {filename}")
                            processed += 1
                        else:
                            self.log(f"Ошибка обработки: {filename}")

                    elif extension == '.dwg':
                        rules = self.config_data.get(project, {}).get("dwg_parser", {})
                        processor = AutoCADProcessor(replacement_digit, project, rules, log_callback=self.log,
                                                     debug=self.debug_logging.get())
                        output_path = os.path.join(output_dir, new_name + ext)  # Используем new_name!
                        success = processor.process_file(input_path, output_path)
                        if success:
                            self.log(f"Успешно: {filename}")
                            processed += 1
                        else:
                            self.log(f"Ошибка обработки: {filename}")

                    #elif extension == '.pdf':
                    #    rules = self.config_data.get(project, {}).get("pdf_parser", {})
                    #    processor = PdfProcessor(replacement_digit, project, rules, log_callback=self.log,
                    #                             debug=self.debug_logging.get())
                    #    success = processor.process_file(input_path, output_path)
                    #    if success:
                    #        self.log(f"Успешно: {filename}")
                    #        processed += 1
                    #    else:
                    #        self.log(f"Ошибка обработки: {filename}")

                    elif extension == '.sha':
                        rules = self.config_data.get(project, {}).get("sha_parser", {})
                        if rules:  # Проверяем, есть ли правила
                            if not sha_processor:
                                from sha_parser import ShaProcessorWinAPI
                                sha_processor = ShaProcessorWinAPI(replacement_digit, project, rules,
                                                                   log_callback=self.log,
                                                                   debug=self.debug_logging.get())
                            if not sha_app_started:
                                try:
                                    sha_processor.start_app()
                                    sha_app_started = True
                                except Exception as e:
                                    self.log(f"Ошибка запуска SmartSketch для {filename}: {str(e)}")
                                    continue
                            success = sha_processor.process_file(input_path, output_path)
                            if success:
                                self.log(f"Успешно: {filename}")
                                processed += 1
                            else:
                                self.log(f"Ошибка обработки: {filename}")
                        else:
                            self.log(f"Пропуск {filename} (нет правил для sha_parser в config)")

                    else:
                        self.log(f"Пропуск {filename} (неподдерживаемый формат: {extension})")

                except Exception as e:
                    self.log(f"Критическая ошибка {filename}: {str(e)}")

        finally:
            if sha_app_started and sha_processor:
                sha_processor.stop_app()

        return processed

    def run_processing(self):
        project = self.project.get()
        if "digits" in self.config_data.get(project, {}):
            repl_digit = self.replacement_digit.get().strip()
            if not repl_digit.isdigit():
                messagebox.showerror("Ошибка", "Введите корректную цифру для замены!")
                return
        else:
            repl_digit = ""

        input_dir = self.input_dir.get().strip()
        output_dir = self.output_dir.get().strip()

        if not os.path.isdir(input_dir):
            messagebox.showerror("Ошибка", "Выберите существующую папку с исходными файлами!")
            return
        if not output_dir:
            messagebox.showerror("Ошибка", "Выберите папку для сохранения файлов!")
            return

        try:
            log_file_path = os.path.join(output_dir, "log.txt")
            self.log_file = open(log_file_path, 'a', encoding='utf-8')
            self.log("=== Запуск обработки ===")
        except Exception as e:
            self.log(f"Ошибка открытия лог-файла: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось открыть лог-файл: {str(e)}")
            return

        try:
            input_files = self.select_files(input_dir)
            if not input_files:
                self.log("Файлы не найдены.")
                return

            processed_count = self.process_files(input_files, output_dir, repl_digit)
            self.log(f"Обработка завершена. Успешно обработано: {processed_count}/{len(input_files)}")
            self.log(f"Результаты сохранены в: {output_dir}")

            messagebox.showinfo(
                "Готово",
                f"Обработка завершена.\nУспешно обработано: {processed_count}/{len(input_files)}"
            )
        finally:
            if self.log_file:
                self.log_file.close()
                self.log_file = None
                self.log("Лог-файл закрыт")

    def show_about(self):
        about_win = tk.Toplevel(self.root)
        about_win.title("О программе")
        about_win.geometry("400x400")
        about_win.resizable(False, False)

        text = (
            "WESA_Parser\n"
            "Обработчик Excel, Word, DWG и SHA\n\n"
            "--------------------------------------\n\n"
            "Принцип работы программы:\n\n"
            "1. Выберите проект и блок, НА который\n"
            "необходимо произвести замену.\n\n"
            "2. Выберите папку с исходными файлами.\n\n"
            "3. Укажите папку для сохранения.\n\n"
            "4. Программа переименует файлы, заменит\n"
            "   цифру Блока в содержимом, а также\n"
            "   ревизию на C01.\n\n"
            "Поддерживаемые форматы:\n"
            ".doc, .docx, .dotx, .xls, .xlsx, .xlsm, .dwg, .sha.\n\n"
            "Автор: Артем Баюшкин\n"
            "Версия: 1.02 "
            "   MIT License"
        )

        lbl = tk.Label(about_win, text=text, justify="left", padx=5, pady=5)
        lbl.pack(fill="both", expand=True)

def set_icon(root, icon_path):
    img = Image.open(icon_path)
    icon = ImageTk.PhotoImage(img)
    root.iconphoto(False, icon)

if __name__ == "__main__":
    root = tk.Tk()
    app = FileProcessorGUI(root)
    set_icon(root, app.resource_path("icon.png"))
    root.mainloop()