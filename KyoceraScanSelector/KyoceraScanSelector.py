import os
import re
import shutil
import time
import threading
import configparser
import tkinter as tk
from tkinter import ttk, messagebox

# ------------------ ПУТИ ------------------
KYOCERA_PATH_RAW = r"C:\Users\%username%\AppData\Roaming\Kyocera\KM_TWAIN"
KYOCERA_PATH = os.path.expandvars(KYOCERA_PATH_RAW)

# Файл пресетов на сетевом ресурсе
REMOTE_PRESETS_PATH = r"\\storage\Instal\printers\presets.ini"

# Кэш на случай, если сеть недоступна
LOCAL_CACHE_DIR = os.path.join(os.environ.get("ProgramData", "."), "KyoceraPresetCache")
os.makedirs(LOCAL_CACHE_DIR, exist_ok=True)
LOCAL_CACHE_FILE = os.path.join(LOCAL_CACHE_DIR, "presets.cache.ini")

# ------------------ УТИЛИТЫ ------------------
IP_RE = re.compile(r"^(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)$")

def is_valid_ip(ip: str) -> bool:
    return bool(IP_RE.match(ip.strip()))

def resolve_kyocera_path():
    base = KYOCERA_PATH
    if os.path.isfile(base):
        return base
    if os.path.isfile(base + ".ini"):
        return base + ".ini"
    os.makedirs(os.path.dirname(base), exist_ok=True)
    with open(base, "w", encoding="utf-8") as f:
        f.write("[Contents]\nUnit=0\nCompression=0\nCompressionGray=0\nScannerAddress=10.0.0.1\n\n[Authentication]\nUnit=0\nUserName=\nPassword=\n")
    return base

def try_copy_remote_to_cache(remote_path: str) -> bool:
    try:
        if os.path.isfile(remote_path):
            shutil.copyfile(remote_path, LOCAL_CACHE_FILE)
            return True
    except Exception:
        pass
    return False

def load_presets(remote_path: str) -> dict:
    cfg = configparser.ConfigParser()
    used = None
    if os.path.isfile(remote_path):
        try:
            cfg.read(remote_path, encoding="utf-8")
            try_copy_remote_to_cache(remote_path)
            used = remote_path
        except Exception:
            used = None
    if used is None and os.path.isfile(LOCAL_CACHE_FILE):
        try:
            cfg.read(LOCAL_CACHE_FILE, encoding="utf-8")
            used = LOCAL_CACHE_FILE
        except Exception:
            used = None
    presets = {}
    for section in cfg.sections():
        ip = cfg.get(section, "ScannerAddress", fallback="").strip()
        if is_valid_ip(ip):
            presets[section] = ip
    return presets

def read_scanner_ip(ini_path: str) -> str:
    cfg = configparser.ConfigParser()
    cfg.read(ini_path, encoding="utf-8")
    if "Contents" not in cfg:
        cfg["Contents"] = {}
    return cfg["Contents"].get("ScannerAddress", "").strip()

def write_scanner_ip(ini_path: str, ip: str):
    cfg = configparser.ConfigParser()
    cfg.read(ini_path, encoding="utf-8")
    if "Contents" not in cfg:
        cfg["Contents"] = {}
    cfg["Contents"]["ScannerAddress"] = ip
    with open(ini_path, "w", encoding="utf-8") as f:
        cfg.write(f)

# ------------------ GUI ------------------
class KyoceraGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Kyocera Scanner Presets")
        self.geometry("500x340")
        self.resizable(False, False)

        ttk.Label(self, text=f"INI файл: {resolve_kyocera_path()}", foreground="#555").pack(anchor="w", padx=10, pady=5)

        # текущий IP
        frame_cur = ttk.LabelFrame(self, text="Текущий IP", padding=10)
        frame_cur.pack(fill="x", padx=10, pady=5)
        self.var_ip = tk.StringVar(value=read_scanner_ip(resolve_kyocera_path()))
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
        self.presets = load_presets(REMOTE_PRESETS_PATH)
        if not self.presets:
            self.combo["values"] = []
            self.var_status.set("Нет доступа к сети или кэшу")
        else:
            self.combo["values"] = list(self.presets.keys())
            if not self.var_preset.get() and self.presets:
                self.var_preset.set(list(self.presets.keys())[0])
            self.var_status.set(f"Загружено {len(self.presets)} пресетов")

    def watcher(self):
        last_ts = 0
        while not self.stop_flag.is_set():
            if self.var_auto.get():
                try:
                    ts = os.path.getmtime(REMOTE_PRESETS_PATH)
                    if ts != last_ts:
                        last_ts = ts
                        self.after(0, self.refresh_presets)
                except Exception:
                    pass
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
        ip = self.var_ip.get().strip()
        if not is_valid_ip(ip):
            messagebox.showwarning("Ошибка", "Некорректный IP адрес")
            return
        try:
            write_scanner_ip(resolve_kyocera_path(), ip)
            self.var_status.set(f"Сохранено: {ip}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            self.var_status.set("Ошибка сохранения")

    def on_close(self):
        self.stop_flag.set()
        self.destroy()

if __name__ == "__main__":
    KyoceraGUI().mainloop()
