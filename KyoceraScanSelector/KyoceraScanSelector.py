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
        self.title("Kyocera Scanner Presets")
        self.geometry("500x380")
        self.resizable(False, False)

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

        ttk.Label(self, text=f"INI файл: {self.kyocera_ini_path}", foreground="#555").pack(anchor="w", padx=10, pady=5)

        # Информация о директории кэша
        ttk.Label(self, text=f"Кэш: {LOCAL_CACHE_DIR}", foreground="#888", font=("Arial", 8)).pack(anchor="w", padx=10, pady=2)

        # текущий IP
        frame_cur = ttk.LabelFrame(self, text="Текущий IP", padding=10)
        frame_cur.pack(fill="x", padx=10, pady=5)

        current_ip = read_scanner_ip(self.kyocera_ini_path)
        self.var_ip = tk.StringVar(value=current_ip if current_ip else "10.0.0.1")

        ttk.Entry(frame_cur, textvariable=self.var_ip, width=30).pack(side="left", padx=5)
        ttk.Button(frame_cur, text="Сохранить", command=self.save_ip).pack(side="left", padx=5)

        # пресеты
        frame_pre = ttk.LabelFrame(self, text="Преднастройки (сети / кэш)", padding=10)
        frame_pre.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(frame_pre, text=f"Источник: {REMOTE_PRESETS_PATH}").pack(anchor="w")

        self.var_preset = tk.StringVar()
        self.combo = ttk.Combobox(frame_pre, textvariable=self.var_preset, state="readonly", width=40)
        self.combo.pack(pady=8)
        ttk.Button(frame_pre, text="Применить пресет", command=self.apply_preset).pack()

        self.var_status = tk.StringVar(value="Готово")
        ttk.Label(frame_pre, textvariable=self.var_status, foreground="#007700").pack(pady=6)

        # автообновление
        self.var_auto = tk.BooleanVar(value=True)
        ttk.Checkbutton(self, text="Автообновление пресетов каждые 30 сек", variable=self.var_auto).pack(anchor="w", padx=10)

        # первичная загрузка
        self.presets = {}
        self.refresh_presets()
        self.stop_flag = threading.Event()
        threading.Thread(target=self.watcher, daemon=True).start()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

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
            messagebox.showwarning("Ошибка валидации", "Некорректный IP адрес.\n\nПример правильного IP: 192.168.1.1")
            logger.warning(f"Попытка сохранить некорректный IP: {ip}")
            return

        try:
            write_scanner_ip(self.kyocera_ini_path, ip)
            self.var_status.set(f"✓ Сохранено: {ip}")
            logger.info(f"IP успешно сохранен: {ip}")
        except PermissionError as e:
            error_msg = f"Нет прав для записи в файл конфигурации.\n\n{self.kyocera_ini_path}\n\nПопробуйте запустить программу от имени администратора."
            messagebox.showerror("Ошибка прав доступа", error_msg)
            self.var_status.set("✗ Ошибка: нет прав доступа")
            logger.error(f"Ошибка прав при сохранении IP: {e}")
        except IOError as e:
            error_msg = f"Ошибка записи в файл.\n\n{self.kyocera_ini_path}\n\nПроверьте, что файл не открыт в другой программе."
            messagebox.showerror("Ошибка записи", error_msg)
            self.var_status.set("✗ Ошибка записи")
            logger.error(f"Ошибка I/O при сохранении IP: {e}")
        except Exception as e:
            error_msg = f"Неожиданная ошибка при сохранении.\n\n{str(e)}"
            messagebox.showerror("Ошибка", error_msg)
            self.var_status.set("✗ Ошибка сохранения")
            logger.error(f"Неожиданная ошибка при сохранении IP: {e}")

    def on_close(self):
        self.stop_flag.set()
        self.destroy()

if __name__ == "__main__":
    KyoceraGUI().mainloop()
