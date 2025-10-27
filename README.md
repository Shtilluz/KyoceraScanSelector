# Kyocera Scan Selector

Простая утилита с графическим интерфейсом для смены IP-адреса сканера Kyocera без необходимости ручного редактирования INI-файла.

## Возможности
- Читает и изменяет файл `KM_TWAIN` в профиле пользователя.
- Поддержка пресетов, хранящихся на сетевом ресурсе (`\\storage\Instal\printers\presets.ini`).
- Автоматическое обновление списка пресетов каждые 30 секунд.
- Работа из локального кэша, если сеть недоступна.

## Установка
```bash
git clone https://github.com/Shtilluz/KyoceraScanSelector.git
cd KyoceraScanSelector
pip install -r requirements.txt
python KyoceraScanSelector.py
