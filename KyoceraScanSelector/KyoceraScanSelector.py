import os
import re
import shutil
import time
import threading
import configparser
import tkinter as tk
from tkinter import ttk, messagebox
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ------------------ ПУТИ ------------------
KYOCERA_PATH_RAW = r"C:\Users\%username%\AppData\Roaming\Kyocera\KM_TWAIN"
KYOCERA_PATH = os.path.expandvars(KYOCERA_PATH_RAW)

# Файл пресетов на сетевом ресурсе
REMOTE_PRESETS_PATH = r"\\storage\Instal\printers\presets.ini"

# Кэш на случай, если сеть недоступна (используем LOCALAPPDATA вместо ProgramData)
LOCAL_CACHE_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "KyoceraPresetCache")
LOCAL_CACHE_FILE = os.path.join(LOCAL_CACHE_DIR, "presets.cache.ini")

# ------------------ УТИЛИТЫ ------------------
IP_RE = re.compile(r"^(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)$")

def is_valid_ip(ip: str) -> bool:
    """Проверка валидности IP адреса"""
    return bool(IP_RE.match(ip.strip()))

def ensure_directory(path: str) -> bool:
    """Безопасное создание директории с обработкой ошибок"""
    try:
        os.makedirs(path, exist_ok=True)
        # Проверяем, что директория действительно создана и доступна для записи
        test_file = os.path.join(path, ".write_test")
        try:
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)
            logger.info(f"Директория создана и доступна для записи: {path}")
            return True
        except (OSError, IOError) as e:
            logger.warning(f"Директория создана, но недоступна для записи: {path} - {e}")
            return False
    except (OSError, IOError, PermissionError) as e:
        logger.error(f"Не удалось создать директорию: {path} - {e}")
        return False

def check_file_writable(file_path: str) -> bool:
    """Проверка возможности записи в файл"""
    try:
        # Если файл существует, проверяем права на запись
        if os.path.exists(file_path):
            return os.access(file_path, os.W_OK)
        # Если файла нет, проверяем права на запись в директорию
        else:
            directory = os.path.dirname(file_path)
            if not directory:
                directory = "."
            return os.access(directory, os.W_OK)
    except Exception as e:
        logger.error(f"Ошибка проверки прав доступа для {file_path}: {e}")
        return False

def resolve_kyocera_path():
    """Определение и создание пути к файлу конфигурации Kyocera с обработкой ошибок"""
    base = KYOCERA_PATH

    # Проверяем существующие варианты файла
    if os.path.isfile(base):
        logger.info(f"Найден файл конфигурации: {base}")
        return base
    if os.path.isfile(base + ".ini"):
        logger.info(f"Найден файл конфигурации: {base}.ini")
        return base + ".ini"

    # Создаем директорию и файл, если их нет
    directory = os.path.dirname(base)
    if directory and not ensure_directory(directory):
        logger.error(f"Не удалось создать директорию для конфигурации: {directory}")
        raise PermissionError(f"Нет прав для создания директории: {directory}")

    # Создаем файл с настройками по умолчанию
    try:
        default_config = "[Contents]\nUnit=0\nCompression=0\nCompressionGray=0\nScannerAddress=10.0.0.1\n\n[Authentication]\nUnit=0\nUserName=\nPassword=\n"
        with open(base, "w", encoding="utf-8") as f:
            f.write(default_config)
        logger.info(f"Создан файл конфигурации по умолчанию: {base}")
        return base
    except (OSError, IOError, PermissionError) as e:
        logger.error(f"Не удалось создать файл конфигурации: {base} - {e}")
        raise PermissionError(f"Нет прав для создания файла конфигурации: {base}")

def try_copy_remote_to_cache(remote_path: str) -> bool:
    """Попытка скопировать файл пресетов из сети в локальный кэш"""
    try:
        # Проверяем доступность файла
        if not os.path.isfile(remote_path):
            logger.warning(f"Файл пресетов не найден: {remote_path}")
            return False

        # Создаем директорию для кэша, если её нет
        if not ensure_directory(LOCAL_CACHE_DIR):
            logger.error(f"Не удалось создать директорию кэша: {LOCAL_CACHE_DIR}")
            return False

        # Проверяем права на запись в файл кэша
        if not check_file_writable(LOCAL_CACHE_FILE):
            logger.error(f"Нет прав на запись в файл кэша: {LOCAL_CACHE_FILE}")
            return False

        # Копируем файл
        shutil.copyfile(remote_path, LOCAL_CACHE_FILE)
        logger.info(f"Пресеты скопированы в кэш: {LOCAL_CACHE_FILE}")
        return True

    except PermissionError as e:
        logger.error(f"Нет прав для копирования в кэш: {e}")
        return False
    except (OSError, IOError) as e:
        logger.error(f"Ошибка копирования в кэш: {e}")
        return False
    except Exception as e:
        logger.error(f"Неожиданная ошибка при копировании в кэш: {e}")
        return False

def load_presets(remote_path: str) -> dict:
    """Загрузка пресетов из сетевого файла или локального кэша"""
    cfg = configparser.ConfigParser()
    used = None

    # Сначала пытаемся загрузить из сетевого источника
    if os.path.isfile(remote_path):
        try:
            cfg.read(remote_path, encoding="utf-8")
            logger.info(f"Пресеты загружены из сети: {remote_path}")
            # Пытаемся обновить кэш
            if try_copy_remote_to_cache(remote_path):
                logger.info("Кэш успешно обновлен")
            used = remote_path
        except PermissionError as e:
            logger.error(f"Нет прав для чтения файла пресетов: {remote_path} - {e}")
            used = None
        except Exception as e:
            logger.error(f"Ошибка чтения файла пресетов: {remote_path} - {e}")
            used = None
    else:
        logger.warning(f"Сетевой файл пресетов недоступен: {remote_path}")

    # Если не удалось загрузить из сети, используем кэш
    if used is None and os.path.isfile(LOCAL_CACHE_FILE):
        try:
            cfg.read(LOCAL_CACHE_FILE, encoding="utf-8")
            logger.info(f"Пресеты загружены из кэша: {LOCAL_CACHE_FILE}")
            used = LOCAL_CACHE_FILE
        except Exception as e:
            logger.error(f"Ошибка чтения кэша: {LOCAL_CACHE_FILE} - {e}")
            used = None

    # Парсим конфигурацию
    presets = {}
    for section in cfg.sections():
        try:
            ip = cfg.get(section, "ScannerAddress", fallback="").strip()
            if is_valid_ip(ip):
                presets[section] = ip
            else:
                logger.warning(f"Некорректный IP для пресета '{section}': {ip}")
        except Exception as e:
            logger.error(f"Ошибка обработки секции '{section}': {e}")

    if used:
        logger.info(f"Загружено {len(presets)} пресетов из {used}")
    else:
        logger.warning("Не удалось загрузить пресеты ни из сети, ни из кэша")

    return presets

def read_scanner_ip(ini_path: str) -> str:
    """Чтение IP адреса сканера из INI файла"""
    try:
        cfg = configparser.ConfigParser()
        cfg.read(ini_path, encoding="utf-8")
        if "Contents" not in cfg:
            logger.warning(f"Секция [Contents] не найдена в {ini_path}")
            return ""
        ip = cfg["Contents"].get("ScannerAddress", "").strip()
        logger.info(f"Прочитан IP из конфигурации: {ip}")
        return ip
    except PermissionError as e:
        logger.error(f"Нет прав для чтения файла: {ini_path} - {e}")
        return ""
    except Exception as e:
        logger.error(f"Ошибка чтения IP из файла: {ini_path} - {e}")
        return ""

def write_scanner_ip(ini_path: str, ip: str):
    """Запись IP адреса сканера в INI файл"""
    # Проверяем права на запись
    if not check_file_writable(ini_path):
        error_msg = f"Нет прав для записи в файл: {ini_path}"
        logger.error(error_msg)
        raise PermissionError(error_msg)

    try:
        # Читаем существующую конфигурацию
        cfg = configparser.ConfigParser()
        if os.path.exists(ini_path):
            cfg.read(ini_path, encoding="utf-8")

        # Обновляем IP
        if "Contents" not in cfg:
            cfg["Contents"] = {}
        cfg["Contents"]["ScannerAddress"] = ip

        # Записываем в файл
        with open(ini_path, "w", encoding="utf-8") as f:
            cfg.write(f)

        logger.info(f"IP успешно записан в {ini_path}: {ip}")

    except PermissionError as e:
        error_msg = f"Нет прав для записи в файл: {ini_path}"
        logger.error(f"{error_msg} - {e}")
        raise PermissionError(error_msg)
    except (OSError, IOError) as e:
        error_msg = f"Ошибка записи в файл: {ini_path}"
        logger.error(f"{error_msg} - {e}")
        raise IOError(error_msg)
    except Exception as e:
        error_msg = f"Неожиданная ошибка при записи в файл: {ini_path}"
        logger.error(f"{error_msg} - {e}")
        raise

# ------------------ GUI ------------------
class KyoceraGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Kyocera Scan Selector")
        self.geometry("520x400")
        self.resizable(False, False)

        # Настройка цветовой схемы
        self.configure(bg="#f0f0f0")

        # Попытка установить иконку
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'printer.ico')
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
                logger.info(f"Иконка загружена: {icon_path}")
        except Exception as e:
            logger.debug(f"Не удалось загрузить иконку: {e}")

        # Настройка стилей
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('Segoe UI', 12, 'bold'), background="#f0f0f0", foreground="#333")
        style.configure('Status.TLabel', font=('Segoe UI', 9), background="#f0f0f0", foreground="#666")
        style.configure('Success.TLabel', font=('Segoe UI', 9), background="#f0f0f0", foreground="#2d8659")
        style.configure('Custom.TLabelframe', background="#f0f0f0")
        style.configure('Custom.TLabelframe.Label', font=('Segoe UI', 10, 'bold'), background="#f0f0f0", foreground="#444")
        style.configure('Action.TButton', font=('Segoe UI', 9, 'bold'), padding=6)

        # Инициализация пути к конфигурации с обработкой ошибок
        try:
            self.kyocera_ini_path = resolve_kyocera_path()
            logger.info(f"Путь к конфигурации: {self.kyocera_ini_path}")
        except PermissionError as e:
            messagebox.showerror(
                "Ошибка прав доступа",
                f"Нет прав для создания файла конфигурации.\n\n{e}\n\nПопробуйте запустить программу от имени администратора."
            )
            logger.critical(f"Критическая ошибка: {e}")
            self.destroy()
            return

        # Создание меню
        self._create_menu()

        # Заголовок приложения
        header_frame = tk.Frame(self, bg="#2d5f8d", height=60)
        header_frame.pack(fill="x", pady=(0, 15))
        header_frame.pack_propagate(False)

        title_label = tk.Label(header_frame, text="🖨  Kyocera Scan Selector",
                               font=('Segoe UI', 16, 'bold'), bg="#2d5f8d", fg="white")
        title_label.pack(pady=15)

        # Основной контейнер
        main_container = tk.Frame(self, bg="#f0f0f0")
        main_container.pack(fill="both", expand=True, padx=20, pady=10)

        # Фрейм текущего IP
        frame_cur = ttk.LabelFrame(main_container, text="Адрес сканера",
                                   padding=15, style='Custom.TLabelframe')
        frame_cur.pack(fill="x", pady=(0, 15))

        current_ip = read_scanner_ip(self.kyocera_ini_path)
        self.var_ip = tk.StringVar(value=current_ip if current_ip else "10.0.0.1")

        ip_frame = tk.Frame(frame_cur, bg="#f0f0f0")
        ip_frame.pack(fill="x")

        tk.Label(ip_frame, text="IP адрес:", font=('Segoe UI', 10),
                bg="#f0f0f0", fg="#444").pack(side="left", padx=(0, 10))

        ip_entry = ttk.Entry(ip_frame, textvariable=self.var_ip, width=20, font=('Consolas', 11))
        ip_entry.pack(side="left", padx=(0, 10))

        ttk.Button(ip_frame, text="💾 Сохранить", command=self.save_ip,
                  style='Action.TButton').pack(side="left")

        # Фрейм пресетов
        frame_pre = ttk.LabelFrame(main_container, text="Быстрый выбор",
                                   padding=15, style='Custom.TLabelframe')
        frame_pre.pack(fill="both", expand=True)

        tk.Label(frame_pre, text="Выберите сканер из списка:",
                font=('Segoe UI', 10), bg="#f0f0f0", fg="#444").pack(anchor="w", pady=(0, 8))

        self.var_preset = tk.StringVar()
        self.combo = ttk.Combobox(frame_pre, textvariable=self.var_preset,
                                 state="readonly", width=40, font=('Segoe UI', 10))
        self.combo.pack(fill="x", pady=(0, 12))

        ttk.Button(frame_pre, text="✓ Применить", command=self.apply_preset,
                  style='Action.TButton').pack()

        # Статус
        status_frame = tk.Frame(frame_pre, bg="#f0f0f0", height=30)
        status_frame.pack(fill="x", pady=(15, 5))
        status_frame.pack_propagate(False)

        self.var_status = tk.StringVar(value="Загрузка...")
        self.status_label = tk.Label(status_frame, textvariable=self.var_status,
                                     font=('Segoe UI', 9), bg="#f0f0f0", fg="#666", anchor="w")
        self.status_label.pack(fill="both")

        # Нижняя панель
        bottom_frame = tk.Frame(self, bg="#f0f0f0", height=40)
        bottom_frame.pack(fill="x", side="bottom", padx=20, pady=(0, 10))
        bottom_frame.pack_propagate(False)

        # Чекбокс автообновления
        self.var_auto = tk.BooleanVar(value=True)
        auto_check = ttk.Checkbutton(bottom_frame, text="Автоматически обновлять список",
                                     variable=self.var_auto)
        auto_check.pack(side="left")

        # первичная загрузка
        self.presets = {}
        self.refresh_presets()
        self.stop_flag = threading.Event()
        threading.Thread(target=self.watcher, daemon=True).start()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _create_menu(self):
        """Создание меню приложения"""
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        # Меню "Справка"
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Справка", menu=help_menu)
        help_menu.add_command(label="О программе", command=self._show_about)
        help_menu.add_separator()
        help_menu.add_command(label="Техническая информация", command=self._show_tech_info)

    def _show_about(self):
        """Показать окно 'О программе'"""
        about_window = tk.Toplevel(self)
        about_window.title("О программе")
        about_window.geometry("420x320")
        about_window.resizable(False, False)
        about_window.configure(bg="#f0f0f0")
        about_window.transient(self)
        about_window.grab_set()

        # Заголовок
        header = tk.Frame(about_window, bg="#2d5f8d", height=80)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="🖨", font=('Segoe UI', 32), bg="#2d5f8d", fg="white").pack(pady=(10, 0))
        tk.Label(header, text="Kyocera Scan Selector", font=('Segoe UI', 14, 'bold'),
                bg="#2d5f8d", fg="white").pack()

        # Основная информация
        content = tk.Frame(about_window, bg="#f0f0f0")
        content.pack(fill="both", expand=True, padx=30, pady=20)

        tk.Label(content, text="Версия 2.0", font=('Segoe UI', 10),
                bg="#f0f0f0", fg="#666").pack(pady=(0, 15))

        tk.Label(content, text="Утилита для быстрого переключения\nмежду сканерами Kyocera",
                font=('Segoe UI', 10), bg="#f0f0f0", fg="#444",
                justify="center").pack(pady=(0, 20))

        # Разработчик
        dev_frame = tk.Frame(content, bg="#e8e8e8", relief="ridge", bd=1)
        dev_frame.pack(fill="x", pady=10)

        tk.Label(dev_frame, text="Разработчик", font=('Segoe UI', 9, 'bold'),
                bg="#e8e8e8", fg="#333").pack(pady=(10, 5))

        email_label = tk.Label(dev_frame, text="bigus400@gmail.com",
                              font=('Segoe UI', 10), bg="#e8e8e8", fg="#2d5f8d",
                              cursor="hand2")
        email_label.pack(pady=(0, 10))
        email_label.bind("<Button-1>", lambda e: self._copy_to_clipboard("bigus400@gmail.com"))

        tk.Label(content, text="© 2025 Все права защищены",
                font=('Segoe UI', 8), bg="#f0f0f0", fg="#999").pack(side="bottom", pady=(15, 0))

    def _show_tech_info(self):
        """Показать окно с технической информацией"""
        tech_window = tk.Toplevel(self)
        tech_window.title("Техническая информация")
        tech_window.geometry("600x400")
        tech_window.resizable(True, True)
        tech_window.configure(bg="#f0f0f0")
        tech_window.transient(self)

        # Заголовок
        header = tk.Frame(tech_window, bg="#2d5f8d", height=50)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="⚙ Техническая информация", font=('Segoe UI', 12, 'bold'),
                bg="#2d5f8d", fg="white").pack(pady=12)

        # Текстовая область
        text_frame = tk.Frame(tech_window, bg="#f0f0f0")
        text_frame.pack(fill="both", expand=True, padx=15, pady=15)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        text_widget = tk.Text(text_frame, wrap="word", font=('Consolas', 9),
                             bg="#ffffff", fg="#333", yscrollcommand=scrollbar.set,
                             relief="solid", bd=1)
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=text_widget.yview)

        # Формирование технической информации
        tech_info = f"""ПУТИ К ФАЙЛАМ
{'='*60}

Файл конфигурации Kyocera:
  {self.kyocera_ini_path}

Директория кэша:
  {LOCAL_CACHE_DIR}

Файл кэша пресетов:
  {LOCAL_CACHE_FILE}

Сетевой файл пресетов:
  {REMOTE_PRESETS_PATH}


СТАТУС ФАЙЛОВ
{'='*60}

Конфигурация существует: {'Да' if os.path.exists(self.kyocera_ini_path) else 'Нет'}
Кэш существует: {'Да' if os.path.exists(LOCAL_CACHE_FILE) else 'Нет'}
Сетевой файл доступен: {'Да' if os.path.exists(REMOTE_PRESETS_PATH) else 'Нет'}


ТЕКУЩИЕ НАСТРОЙКИ
{'='*60}

IP адрес сканера: {self.var_ip.get()}
Количество пресетов: {len(self.presets)}
Автообновление: {'Включено' if self.var_auto.get() else 'Выключено'}


ЗАГРУЖЕННЫЕ ПРЕСЕТЫ
{'='*60}
"""

        for name, ip in sorted(self.presets.items()):
            tech_info += f"\n{name}: {ip}"

        if not self.presets:
            tech_info += "\nПресеты не загружены"

        text_widget.insert("1.0", tech_info)
        text_widget.config(state="disabled")

    def _copy_to_clipboard(self, text):
        """Копировать текст в буфер обмена"""
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("Скопировано", f"'{text}' скопирован в буфер обмена")

    def refresh_presets(self):
        """Обновление списка пресетов из сети или кэша"""
        try:
            self.presets = load_presets(REMOTE_PRESETS_PATH)

            if not self.presets:
                self.combo["values"] = []
                # Проверяем, есть ли доступ к кэшу
                if os.path.exists(LOCAL_CACHE_FILE):
                    self.var_status.set("⚠ Пресеты недоступны (файл пустой или поврежден)")
                else:
                    self.var_status.set("⚠ Нет доступа к сети или кэшу")
                logger.warning("Не удалось загрузить пресеты")
            else:
                preset_names = sorted(self.presets.keys())
                self.combo["values"] = preset_names

                # Устанавливаем первый пресет, если ничего не выбрано
                if not self.var_preset.get() and preset_names:
                    self.var_preset.set(preset_names[0])

                # Показываем источник данных
                if os.path.exists(REMOTE_PRESETS_PATH):
                    source = "сеть"
                else:
                    source = "кэш"

                self.var_status.set(f"✓ Загружено {len(self.presets)} пресетов ({source})")
                logger.info(f"Обновлено {len(self.presets)} пресетов из {source}")

        except Exception as e:
            logger.error(f"Ошибка обновления пресетов: {e}")
            self.var_status.set(f"✗ Ошибка загрузки пресетов")

    def watcher(self):
        """Фоновый поток для автоматического обновления пресетов"""
        last_ts = 0
        consecutive_errors = 0
        max_consecutive_errors = 5

        while not self.stop_flag.is_set():
            if self.var_auto.get():
                try:
                    # Проверяем время модификации файла пресетов
                    ts = os.path.getmtime(REMOTE_PRESETS_PATH)
                    if ts != last_ts:
                        last_ts = ts
                        self.after(0, self.refresh_presets)
                        logger.info("Обнаружено обновление файла пресетов")
                    consecutive_errors = 0  # Сбрасываем счетчик ошибок при успехе

                except PermissionError as e:
                    consecutive_errors += 1
                    logger.warning(f"Нет доступа к файлу пресетов: {e}")
                    if consecutive_errors >= max_consecutive_errors:
                        logger.error(f"Слишком много ошибок доступа ({consecutive_errors}). Проверьте права доступа к сетевому ресурсу.")
                        consecutive_errors = 0  # Сбрасываем, чтобы не спамить логами

                except (OSError, IOError) as e:
                    consecutive_errors += 1
                    logger.debug(f"Сетевой файл недоступен (попытка {consecutive_errors}): {e}")
                    if consecutive_errors >= max_consecutive_errors:
                        logger.warning(f"Сетевой файл недоступен после {consecutive_errors} попыток. Используется кэш.")
                        consecutive_errors = 0

                except Exception as e:
                    logger.error(f"Неожиданная ошибка в watcher: {e}")

            time.sleep(30)

    def apply_preset(self):
        name = self.var_preset.get().strip()
        if name and name in self.presets:
            ip = self.presets[name]
            self.var_ip.set(ip)
            self.save_ip()
        else:
            messagebox.showinfo("Инфо", "Выберите пресет из списка")

    def save_ip(self):
        """Сохранение IP адреса с проверкой прав доступа"""
        ip = self.var_ip.get().strip()

        # Проверка валидности IP
        if not is_valid_ip(ip):
            messagebox.showwarning(
                "Некорректный адрес",
                "Введите правильный IP адрес.\n\nПример: 192.168.1.100"
            )
            logger.warning(f"Попытка сохранить некорректный IP: {ip}")
            return

        try:
            write_scanner_ip(self.kyocera_ini_path, ip)
            self.var_status.set(f"✓ IP адрес успешно сохранен: {ip}")
            self.status_label.config(fg="#2d8659")
            logger.info(f"IP успешно сохранен: {ip}")
            # Показываем успешное сообщение
            messagebox.showinfo("Успешно", f"IP адрес сканера изменен на:\n{ip}")
        except PermissionError as e:
            error_msg = "Нет прав для сохранения настроек.\n\nПопробуйте запустить программу от имени администратора."
            messagebox.showerror("Ошибка доступа", error_msg)
            self.var_status.set("✗ Ошибка: нет прав доступа")
            self.status_label.config(fg="#c93838")
            logger.error(f"Ошибка прав при сохранении IP: {e}")
        except IOError as e:
            error_msg = "Не удалось сохранить настройки.\n\nПроверьте, что файл конфигурации не открыт в другой программе."
            messagebox.showerror("Ошибка записи", error_msg)
            self.var_status.set("✗ Ошибка записи в файл")
            self.status_label.config(fg="#c93838")
            logger.error(f"Ошибка I/O при сохранении IP: {e}")
        except Exception as e:
            error_msg = f"Произошла неожиданная ошибка.\n\n{str(e)}"
            messagebox.showerror("Ошибка", error_msg)
            self.var_status.set("✗ Неожиданная ошибка")
            self.status_label.config(fg="#c93838")
            logger.error(f"Неожиданная ошибка при сохранении IP: {e}")

    def on_close(self):
        self.stop_flag.set()
        self.destroy()

if __name__ == "__main__":
    KyoceraGUI().mainloop()
