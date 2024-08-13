import os
import shutil
from datetime import datetime, timedelta
from PIL import Image
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import ctypes
import json

# Основные категории
categories = ["Выговоры", "Объезды", "Обзвон", "Оплаченные выговоры", "Таблетки", "ПМП", "ПМП ОБ", "Мед. Карты", "Уколы", "МП-ГМП-АКЦИИ"]

# Подкатегории актуальности
subcategories = ["Повышение + Недельный отчет", "Недельный отчет", "Мусор"]

# Настройка Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Путь к файлу настроек
SETTINGS_FILE = "settings.json"

# Получение текущего времени
now = datetime.now()
start_of_week = (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)

# Функция для определения категории на основе распознавания текста
def determine_category(screenshot):
    try:
        resolution = screenshot.size  # (width, height)

        # Новые области для распознавания текста
        areas = [
            (327, 972, 994, 1030),    # Новая область 1
            (312, 387, 905, 931),     # Новая область 2
            (313, 257, 729, 908),     # Новая область 3
            (298, 5, 1273, 308),      # Новая область 4
            (311, 17, 701, 62),       # Новая область 5
            (356, 24, 556, 58),       # Новая область 6
        ]

        # Старые области для распознавания текста
        old_areas = [
            (resolution[0] / 2 - 270, resolution[1] - 290, resolution[0] / 2 + 290, resolution[1] - 30),  # area1
            (resolution[0] / 2 - 585, resolution[1] / 2 - 225, resolution[0] / 2 + 115, resolution[1] / 2 - 185),  # area2
            (resolution[0] / 2 - 250, resolution[1] / 2 - 95, resolution[0] / 2 - 30, resolution[1] / 2 - 75),  # area3
            (int(430 * resolution[0] / 2560), resolution[1] - 75, int(730 * resolution[0] / 2560), resolution[1] - 25)  # area4
        ]

        # Извлечение текста из каждой области
        texts = []
        for area in areas + old_areas:
            region = screenshot.crop(area)
            text = str(pytesseract.image_to_string(region, lang='rus+eng')).strip().lower()
            texts.append(text)

        # Объединение текста из всех областей
        combined_text = ' '.join(texts)

        # Определение категории по ключевым словам
        if 'выговор' in combined_text:
            return "Выговоры"
        elif 'обзвон' in combined_text:
            return "Обзвоны"
        elif 'объезд' in combined_text:
            return "Объезды"
        elif 'отработка' in combined_text:
            return "Оплаченные выговоры"
        elif 'ваши' in combined_text and 'предметы' in combined_text and 'склад' in combined_text:
            return "ПМП"
        elif all(keyword in texts[0] for keyword in ['вы', 'успешно', 'оказали', 'первую', 'помощь']):
            if any(place in texts[3] for place in
                   ['занкудо', 'палето', 'чилиад', 'гора', 'джосайя', 'гордо', 'пустыня', 'гранд', 'сенора',
                    'сан-шаньский', 'шорс', 'чумаш', 'грейпсид', 'брэддока', 'хармони', 'ратон', 'аламо', 'кэтфиш',
                    'дигнити', 'дейвис', 'кварц', 'тюрьма']):
                return "ПМП ОБ"
            else:
                return "ПМП"
        elif 'гражданин' in combined_text and 'принял' in combined_text and 'предложение' in combined_text:
            if '1500' in combined_text:
                return "Уколы"
            else:
                return "Таблетки"
        elif 'medical' in combined_text and 'card' in combined_text:
            return "Мед. Карты"
        else:
            return None
    except Exception as e:
        print(f"Ошибка при распознавании текста: {e}")
        return None

# Функция для определения категории на основе имени файла
def determine_category_by_filename(filename):
    filename_lower = filename.lower()
    if 'выговор' in filename_lower:
        return "Выговоры"
    elif 'объезд' in filename_lower:
        return "Объезды"
    elif 'обзвон' in filename_lower:
        return "Обзвон"
    elif 'отработка' in filename_lower:
        return "Оплаченные выговоры"
    else:
        return None

# Функция для перемещения скриншотов в нужные подкатегории
def move_to_subcategory(category_path, screenshot_path, screenshot_name, last_promotion_time):
    try:
        creation_time = datetime.fromtimestamp(os.path.getmtime(screenshot_path))
        if creation_time > last_promotion_time:
            new_subcategory = subcategories[0]  # Повышение + Недельный отчет
        elif creation_time > start_of_week:
            new_subcategory = subcategories[1]  # Недельный отчет
        else:
            new_subcategory = subcategories[2]  # Мусор

        new_subcategory_path = os.path.join(category_path, new_subcategory)
        if not os.path.exists(new_subcategory_path):
            os.makedirs(new_subcategory_path)
        shutil.move(screenshot_path, os.path.join(new_subcategory_path, screenshot_name))
    except Exception as e:
        print(f"Ошибка при перемещении файла: {e}")

# Функция для распределения скриншотов по категориям и подкатегориям на основе распознавания текста
def distribute_screenshots_by_text(base_dir, last_promotion_time, log):
    try:
        for filename in os.listdir(base_dir):
            filepath = os.path.join(base_dir, filename)
            if os.path.isfile(filepath) and filepath.lower().endswith(('.png', '.jpg', '.jpeg')):
                try:
                    screenshot = Image.open(filepath)
                    category = determine_category(screenshot)
                    if category:
                        log.insert(tk.END, f"Файл: {filename} -> Категория: {category}\n")
                        category_path = os.path.join(base_dir, category)
                        if not os.path.exists(category_path):
                            os.makedirs(category_path)
                        move_to_subcategory(category_path, filepath, filename, last_promotion_time)
                except Exception as e:
                    log.insert(tk.END, f"Ошибка при обработке файла {filename}: {e}\n")
    except Exception as e:
        log.insert(tk.END, f"Ошибка при распределении скриншотов по тексту: {e}\n")

# Функция для распределения скриншотов по категориям и подкатегориям на основе имени файла
def distribute_screenshots_by_filename(base_dir, last_promotion_time, log):
    try:
        for filename in os.listdir(base_dir):
            filepath = os.path.join(base_dir, filename)
            if os.path.isfile(filepath) and filepath.lower().endswith(('.png', '.jpg', '.jpeg')):
                try:
                    category = determine_category_by_filename(filename)
                    if category:
                        log.insert(tk.END, f"Файл: {filename} -> Категория: {category}\n")
                        category_path = os.path.join(base_dir, category)
                        if not os.path.exists(category_path):
                            os.makedirs(category_path)
                        move_to_subcategory(category_path, filepath, filename, last_promotion_time)
                except Exception as e:
                    log.insert(tk.END, f"Ошибка при обработке файла {filename}: {e}\n")
    except Exception as e:
        log.insert(tk.END, f"Ошибка при распределении скриншотов по имени файла: {e}\n")

# Функция для сохранения настроек
def save_settings(settings):
    try:
        with open(SETTINGS_FILE, "w", encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Ошибка при сохранении настроек: {e}")

# Функция для загрузки настроек
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Ошибка при загрузке настроек: {e}")
    return {}

# Функция для GUI
def run_gui():
    # Загрузка настроек
    settings = load_settings()
    base_dir = settings.get("base_dir", "")
    last_promotion_time_str = settings.get("last_promotion_time", "")

    # Конвертация строки времени последнего повышения в объект datetime
    if last_promotion_time_str:
        last_promotion_time = datetime.fromisoformat(last_promotion_time_str)
    else:
        last_promotion_time = datetime.now()

    # Создание главного окна
    root = tk.Tk()
    root.title("Screenshot Sorter")
    root.geometry("800x600")
    root.resizable(False, False)

    # Логирование
    log = scrolledtext.ScrolledText(root, width=95, height=25)
    log.pack(pady=10)

    # Фрейм для кнопок
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # Функция для выбора директории
    def select_directory():
        nonlocal base_dir
        selected_dir = filedialog.askdirectory(initialdir=base_dir)
        if selected_dir:
            base_dir = selected_dir
            settings["base_dir"] = base_dir
            save_settings(settings)
            log.insert(tk.END, f"Выбранный каталог: {base_dir}\n")

    # Функция для установки времени последнего повышения
    def set_last_promotion_time():
        nonlocal last_promotion_time
        selected_time = tk.simpledialog.askstring("Время последнего повышения", "Введите дату и время в формате YYYY-MM-DD HH:MM:SS")
        if selected_time:
            try:
                last_promotion_time = datetime.strptime(selected_time, "%Y-%m-%d %H:%M:%S")
                settings["last_promotion_time"] = last_promotion_time.isoformat()
                save_settings(settings)
                log.insert(tk.END, f"Время последнего повышения установлено на: {selected_time}\n")
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты и времени.")

    # Функция для начала сортировки по тексту
    def start_sorting_by_text():
        if not base_dir:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите каталог с скриншотами.")
            return
        distribute_screenshots_by_text(base_dir, last_promotion_time, log)
        log.insert(tk.END, "Сортировка по тексту завершена.\n")

    # Функция для начала сортировки по имени файла
    def start_sorting_by_filename():
        if not base_dir:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите каталог с скриншотами.")
            return
        distribute_screenshots_by_filename(base_dir, last_promotion_time, log)
        log.insert(tk.END, "Сортировка по имени файла завершена.\n")

    # Кнопки
    btn_select_dir = tk.Button(button_frame, text="Выбрать каталог", command=select_directory, width=20)
    btn_set_promotion_time = tk.Button(button_frame, text="Установить время повышения", command=set_last_promotion_time, width=25)
    btn_sort_by_text = tk.Button(button_frame, text="Сортировать по тексту", command=start_sorting_by_text, width=20)
    btn_sort_by_filename = tk.Button(button_frame, text="Сортировать по имени файла", command=start_sorting_by_filename, width=25)

    # Расположение кнопок
    btn_select_dir.grid(row=0, column=0, padx=10, pady=5)
    btn_set_promotion_time.grid(row=0, column=1, padx=10, pady=5)
    btn_sort_by_text.grid(row=1, column=0, padx=10, pady=5)
    btn_sort_by_filename.grid(row=1, column=1, padx=10, pady=5)

    # Запуск GUI
    root.mainloop()

# Запуск программы
if __name__ == "__main__":
    run_gui()
