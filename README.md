# Timesheet Timer

Приложение для учёта рабочего времени: выбираете проект и вид работ, запускаете таймер и результат записывается в Excel.

## Требования к Excel‑файлу

- Лист `Справочник` с двумя столбцами: Проект, Вид работ.
- Лист `Учет времени` с колонками: Дата, Проект, Вид работ, Длительность (формат Excel — Время).
- В меню «Помощь → Требования к Excel‑файлу…» можно создать шаблон книги и сразу открыть его для заполнения.

## Быстрый старт (из исходников)

1. Установите Python 3.11+ и создайте окружение:
   ```bash
   python -m venv .venv
   .venv\\Scripts\\activate  # Windows
   ```
2. Установите зависимости:
   ```bash
   pip install -r requirements.txt
   ```
3. Запуск приложения:
   ```bash
   python run_timesheet.py
   ```

## Сборка EXE (PyInstaller)

1. Установите PyInstaller:
   ```bash
   python -m pip install --upgrade pip
   python -m pip install pyinstaller
   ```
2. Выполните команду сборки. Важно: параметр `--add-data` должен поместить папку иконок в корень как `assets` (именно так ищет приложение):
   - Windows/PowerShell (разделитель `;`):
     ```bash
     python -m PyInstaller --noconfirm --onefile --name TimesheetTimer --windowed --add-data src/timesheet_app/assets;assets src/timesheet_app/app.py
     ```
   - Linux/macOS/WSL (разделитель `:`):
     ```bash
     python -m PyInstaller --noconfirm --onefile --name TimesheetTimer --windowed --add-data src/timesheet_app/assets:assets src/timesheet_app/app.py
     ```
   Если при запуске EXE появляется ошибка вида «assets/play.png: no such file or directory», значит ресурсы были упакованы не в `assets`. Пересоберите командой выше.

## Установка (опционально)

Готовый `TimesheetTimer.exe` находится в папке `dist`. Для инсталлятора можно использовать Inno Setup (скрипт в папке `installer`).

## Структура проекта

```
Timesheet/
├── README.md
├── run_timesheet.py
├── requirements.txt
├── src/
│   └── timesheet_app/
│       ├── app.py
│       ├── config.py
│       ├── excel_manager.py
│       ├── version.py
│       └── assets/
│           ├── play.png
│           ├── play_hover.png
│           ├── pause.png
│           ├── pause_hover.png
│           ├── stop.png
│           └── stop_hover.png
└── installer/
    └── TimesheetTimer.iss
```

## Подсказки

- Если файл Excel ещё не выбран, используйте «Файл → Выбрать файл Excel» или создайте шаблон через «Помощь → Требования к Excel‑файлу → Создать шаблон».
- В строке состояния всегда отображается выбранный файл: «Файл: …».

