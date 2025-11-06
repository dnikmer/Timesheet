"""Графический интерфейс (Tkinter) для таймера учёта времени.

Кратко о возможностях:
- выбор проекта и вида работ из Excel-справочника;
- таймер с кнопками Старт/Пауза/Стоп;
- запись результата в книгу Excel (лист "Учет времени");
- меню Файл/Помощь; в Помощи есть окно с требованиями и кнопкой "Создать шаблон";
- строка состояния внизу окна с путём к выбранному файлу.
"""

from __future__ import annotations

import math
import os
import subprocess
import sys
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font, messagebox, ttk
from typing import Callable, Optional


# Импорты одинаково работают и при запуске из исходников, и при запуске из пакета
if __package__ in {None, ""}:  # pragma: no cover - запуск как скрипт
    try:
        from timesheet_app.config import AppConfig
        from timesheet_app.excel_manager import (
            ExcelStructureError,
            REFERENCE_SHEET,
            TIMESHEET_SHEET,
            append_time_entry,
            load_reference_data,
            create_template,
        )
        from timesheet_app.version import VERSION
    except ModuleNotFoundError:  # скрипт рядом с файлами
        from config import AppConfig  # type: ignore
        from excel_manager import (  # type: ignore
            ExcelStructureError,
            REFERENCE_SHEET,
            TIMESHEET_SHEET,
            append_time_entry,
            load_reference_data,
            create_template,
        )
        from version import VERSION  # type: ignore
else:  # стандартный путь импорта пакета
    from .config import AppConfig
    from .excel_manager import (
        ExcelStructureError,
        REFERENCE_SHEET,
        TIMESHEET_SHEET,
        append_time_entry,
        load_reference_data,
        create_template,
    )
    from .version import VERSION


def _asset_path(filename: str) -> str:
    """Вернуть абсолютный путь к ресурсу (иконке).

    - при запуске из исходников ресурсы лежат в `assets` рядом с `app.py`;
    - при запуске из EXE (PyInstaller) ресурсы распакованы во временный каталог
      `_MEIPASS`. Поддерживаем две схемы: `assets/<file>` и
      `timesheet_app/assets/<file>`.
    """

    base_path = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)  # type: ignore[attr-defined]
    primary = Path(base_path) / "assets" / filename
    if primary.exists():
        return str(primary)
    alt = Path(base_path) / "timesheet_app" / "assets" / filename
    return str(alt if alt.exists() else primary)


class DropdownField(ttk.Frame):
    """Поле с подписью и выпадающим списком (Combobox)."""

    _image_cache: dict[str, tk.PhotoImage] = {}

    def __init__(self, parent: tk.Widget, label_text: str, variable: tk.StringVar) -> None:
        super().__init__(parent)
        self.variable = variable
        self._choices: list[str] = []

        # Шрифт для вычисления ширины списков
        try:
            family = "Segoe UI" if sys.platform.startswith("win") else "Arial"
            self._menu_font = font.Font(family=family, size=9)
        except tk.TclError:
            self._menu_font = font.nametofont("TkMenuFont")

        self.label = ttk.Label(self, text=label_text, style="Timesheet.Label")
        self.label.pack(anchor=tk.W, pady=(0, 4))

        # Используем Combobox: выпадающий список открывается под полем
        self.combobox = ttk.Combobox(
            self,
            textvariable=variable,
            state="readonly",
            width=20,
            style="Timesheet.TCombobox",
        )
        self.combobox.pack(fill=tk.X)
        # После выбора пункта убираем выделение текста
        self.combobox.bind("<<ComboboxSelected>>", self._on_combo_selected)
        # Даже если пользователь просто открыл/закрыл список без изменения,
        # Windows оставляет выделение. Уберём его отложенно и снимем фокус.
        self.combobox.bind("<FocusIn>", self._on_focus_in)
        self.combobox.bind("<ButtonRelease-1>", self._on_mouse_release)

    def set_options(self, options: list[str], *, selected: Optional[str] = None) -> None:
        """Задать список значений и выбрать начальное."""

        self._choices = options[:]

        if not options:
            self.combobox.configure(state="disabled", values=[])
            self.variable.set("")
            self.refresh_width()
            return

        self.combobox.configure(state="readonly", values=options)
        if selected in options:
            self.variable.set(selected)
        elif self.variable.get() in options:
            pass  # оставляем предыдущее значение
        else:
            self.variable.set(options[0])

        self.refresh_width()

    def refresh_width(self) -> None:
        """Подобрать ширину виджета по самому длинному варианту."""

        if not self._choices:
            self.combobox.configure(width=20)
            return

        max_pixels = max(self._menu_font.measure(item) for item in self._choices)
        average_char = max(self._menu_font.measure("0"), 1)
        width_chars = max(20, min(int(math.ceil((max_pixels + 24) / average_char)), 64))
        self.combobox.configure(width=width_chars)

    def measure_longest_option(self) -> int:
        """Вернуть ширину (px) самого длинного значения."""

        if not self._choices:
            return 0
        return max(self._menu_font.measure(item) for item in self._choices)

    def _on_combo_selected(self, _event: tk.Event) -> None:  # type: ignore[override]
        """Убрать выделение текста после выбора."""

        try:
            self.combobox.selection_clear()
            self.combobox.icursor("end")
            # Переводим фокус на контейнер, чтобы убрать синий бэкграунд Windows
            # и визуально не подсвечивать поле после выбора значения.
            self.focus_set()
        except Exception:  # pragma: no cover - защита от платформенных мелочей
            pass

    def _on_focus_in(self, _event: tk.Event) -> None:  # type: ignore[override]
        """Снять выделение, если фокус попал в комбобокс без изменения значения."""

        def _defocus() -> None:
            try:
                self.combobox.selection_clear()
                self.combobox.icursor("end")
                self.focus_set()
            except Exception:
                pass

        # Отложим на следующий тик цикла событий, чтобы перебить штатное выделение.
        self.after_idle(_defocus)

    def _on_mouse_release(self, _event: tk.Event) -> None:  # type: ignore[override]
        """После закрытия списка мышью убираем выделение и фокус."""

        self.after(10, lambda: (self.combobox.selection_clear(), self.focus_set()))


class IconButton(ttk.Frame):
    """Кнопка-иконка: меняет изображение при наведении мыши."""

    _image_cache: dict[str, tk.PhotoImage] = {}

    def __init__(self, parent: tk.Widget, icon: str, command: Optional[Callable[[], None]]) -> None:
        super().__init__(parent)
        self.command = command
        self._image_normal = self._load_image(icon)
        self._image_hover = self._load_image(f"{icon}_hover")

        self._button = ttk.Button(
            self,
            image=self._image_normal,
            command=self._on_click,
            style="Timesheet.IconButton.TButton",
            takefocus=False,
        )
        self._button.pack()
        self._button.configure(cursor="hand2")
        self._button.bind("<Enter>", self._on_enter)
        self._button.bind("<Leave>", self._on_leave)

    @classmethod
    def _load_image(cls, name: str) -> tk.PhotoImage:
        if name in cls._image_cache:
            return cls._image_cache[name]
        path = _asset_path(f"{name}.png")
        image = tk.PhotoImage(file=path)
        cls._image_cache[name] = image
        return image

    def _on_click(self) -> None:
        if callable(self.command):
            self.command()

    def _on_enter(self, _event: tk.Event) -> None:  # type: ignore[override]
        self._button.configure(image=self._image_hover)

    def _on_leave(self, _event: tk.Event) -> None:  # type: ignore[override]
        self._button.configure(image=self._image_normal)


class TimeTrackerApp(tk.Tk):
    """Главное окно приложения: меню, форма, таймер и строка состояния."""

    def __init__(self) -> None:
        super().__init__()
        # Окно
        self.title("Учёт рабочего времени")
        self.geometry("440x340")
        self.minsize(420, 320)
        self.resizable(True, True)
        self.configure(background="#f5f5f5")

        # Конфиг и состояние
        self.config_manager = AppConfig.load()
        self.projects: list[str] = []
        self.work_types: list[str] = []

        self._timer_job: Optional[str] = None
        self._timer_running = False
        self._start_reference = 0.0
        self._elapsed_seconds = 0.0

        self.project_var = tk.StringVar()
        self.work_type_var = tk.StringVar()
        self.timer_var = tk.StringVar(value="00:00:00")
        self.status_var = tk.StringVar()

        self._configure_styles()
        self._build_menu()
        self._build_layout()
        self._refresh_status()

        # Если файл уже выбран — пробуем загрузить справочники
        if self.config_manager.excel_path:
            try:
                self._load_reference(self.config_manager.excel_path)
            except Exception as exc:  # pylint: disable=broad-except
                messagebox.showerror("Ошибка", f"Не удалось загрузить Excel файл:\n{exc}")
                self.config_manager.excel_path = None
                self.config_manager.save()
                self._refresh_status()
        # Если данных нет — предложим выбрать файл после старта
        if not self.projects or not self.work_types:
            self.after(100, self._prompt_for_excel)
        else:
            # Если данные подгружены — разрешим выбор
            self._set_inputs_enabled(True)

    # ------------------------- Построение UI -------------------------
    def _configure_styles(self) -> None:
        """Настроить тему и стили виджетов ttk."""

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # Шрифты
        family = "Segoe UI" if sys.platform.startswith("win") else "Arial"
        try:
            default_font = font.nametofont("TkDefaultFont")
            default_font.configure(family=family, size=9)
        except tk.TclError:
            pass
        try:
            menu_font = font.nametofont("TkMenuFont")
            menu_font.configure(family=family, size=9)
        except tk.TclError:
            pass

        # Стили
        style.configure("TFrame", background="#f5f5f5")
        style.configure("Timesheet.Label", font=(family, 9), foreground="#1f1f1f", background="#f5f5f5")
        style.configure("Timesheet.Timer.TLabel", font=(family, 32, "bold"), foreground="#1f1f1f", background="#f5f5f5")
        style.configure(
            "Timesheet.TCombobox",
            padding=(6, 2),
            relief="flat",
            borderwidth=1,
            foreground="#1f1f1f",
            fieldbackground="#ffffff",
            background="#ffffff",
        )
        style.configure(
            "Timesheet.IconButton.TButton",
            background="#ffffff",
            relief="flat",
            padding=0,
            borderwidth=0,
        )
        style.configure("Timesheet.Status.TLabel", font=(family, 9), foreground="#555555", background="#f5f5f5")

    def _build_menu(self) -> None:
        """Создать меню приложения (Файл/Помощь)."""

        menu_bar = tk.Menu(self)

        # Файл
        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Выбрать файл Excel", command=self._prompt_for_excel)
        file_menu.add_command(label="Открыть текущий файл", command=self._open_current_file)
        # Подменю «Обновить»: перечитать лист «Справочник» из выбранного файла
        file_menu.add_command(label="Обновить", command=self._reload_reference)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.destroy)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

        # Помощь
        help_menu = tk.Menu(menu_bar, tearoff=False)
        help_menu.add_command(label="Требования к Excel-файлу...", command=self._show_excel_requirements)
        help_menu.add_separator()
        help_menu.add_command(label="О приложении", command=self._show_about)
        menu_bar.add_cascade(label="Помощь", menu=help_menu)

        self.config(menu=menu_bar)

    def _show_about(self) -> None:
        """Показать информацию о версии приложения."""

        messagebox.showinfo("О приложении", f"Timesheet\nВерсия: {VERSION}")

    def _open_current_file(self) -> None:
        """Открыть текущий выбранный Excel-файл средствами ОС."""

        path = self.config_manager.excel_path
        if not path:
            messagebox.showinfo("Текущий файл", "Файл Excel не выбран")
            return
        try:
            if not Path(path).exists():
                messagebox.showwarning("Нет файла", "Указанный файл не существует. Выберите файл Excel заново.")
                return
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{exc}")

    def _reload_reference(self) -> None:
        """Перечитать лист «Справочник» из выбранного файла.

        Нужна, когда пользователь правит справочник (проекты/виды работ) и хочет
        увидеть изменения без перезапуска приложения.
        """

        if not self.config_manager.excel_path:
            messagebox.showwarning("Нет файла", "Сначала выберите Excel файл через меню 'Файл'.")
            return
        try:
            self._load_reference(self.config_manager.excel_path)
            # Всплывающее сообщение об успешном обновлении
            messagebox.showinfo("Готово", "Справочник обновлён.")
            # В строке состояния оставляем текущий файл
            self._refresh_status()
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("Ошибка", f"Не удалось обновить справочник:\n{exc}")

    def _show_excel_requirements(self) -> None:
        """Показать модальное окно с требованиями к Excel и кнопкой "Создать шаблон"."""

        win = tk.Toplevel(self)
        win.title("Требования к Excel-файлу")
        win.transient(self)
        win.resizable(False, False)
        win.configure(background="#f5f5f5")
        win.grab_set()

        body = ttk.Frame(win, padding=16)
        body.pack(fill=tk.BOTH, expand=True)

        msg = (
            f"Файл Excel должен содержать лист '{REFERENCE_SHEET}'.\n"
            "В нём два столбца: Проект и Вид работ.\n\n"
            f"Также нужен лист '{TIMESHEET_SHEET}', куда добавляются записи:\n"
            "- Дата\n- Проект\n- Вид работ\n- Длительность (формат Время)."
        )
        label = ttk.Label(body, text=msg, justify=tk.LEFT, anchor=tk.W, style="Timesheet.Status.TLabel")
        try:
            label.configure(font=font.nametofont("TkMenuFont"))
        except Exception:
            pass
        label.pack(fill=tk.BOTH, expand=True)

        buttons = ttk.Frame(body)
        buttons.pack(fill=tk.X, pady=(12, 0))

        def on_create_template() -> None:
            if not messagebox.askokcancel("Создать файл?", "Создать файл?"):
                return
            save_path = filedialog.asksaveasfilename(
                parent=win,
                title="Сохранить как",
                defaultextension=".xlsx",
                filetypes=(("Excel", "*.xlsx"), ("All files", "*.*")),
                initialfile="timesheet_template.xlsx",
            )
            if not save_path:
                return
            try:
                # 1) создаём книгу
                create_template(save_path)
                # 2) открываем для заполнения
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(save_path)  # type: ignore[attr-defined]
                    elif sys.platform == "darwin":
                        subprocess.Popen(["open", save_path])
                    else:
                        subprocess.Popen(["xdg-open", save_path])
                except Exception:
                    pass
                # 3) просим вернуться, когда Excel сохранён и закрыт
                messagebox.showinfo(
                    "Продолжите",
                    "После внесения данных в Excel и сохранения файла,\nзакройте Excel и нажмите OK для выбора файла в приложении.",
                )
                # 4) выбираем файл в приложении
                self.config_manager.excel_path = save_path
                self.config_manager.save()
                try:
                    self._load_reference(save_path)
                except Exception:
                    # Если пользователь закрыл шаблон без заполнения справочника —
                    # очищаем текущие списки и блокируем поля, чтобы не остались
                    # данные от предыдущего файла.
                    self.projects = []
                    self.work_types = []
                    self.project_field.set_options([])
                    self.work_field.set_options([])
                    self._set_inputs_enabled(False)
                self._refresh_status()
                messagebox.showinfo("Готово", "Файл выбран в приложении.")
                win.destroy()
            except Exception as exc:  # pylint: disable=broad-except
                messagebox.showerror("Ошибка", f"Не удалось создать файл:\n{exc}")

        create_btn = ttk.Button(buttons, text="Создать шаблон", command=on_create_template)
        ok_btn = ttk.Button(buttons, text="OK", command=win.destroy)
        create_btn.pack(side=tk.LEFT)
        ok_btn.pack(side=tk.RIGHT)

        # Центрируем диалог
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (win.winfo_width() // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (win.winfo_height() // 2)
        win.geometry(f"+{x}+{y}")

    def _build_layout(self) -> None:
        """Построить основную разметку окна (поля, кнопки, статус)."""

        padding = {"padx": 16, "pady": 8}

        container = ttk.Frame(self, style="TFrame")
        container.pack(fill=tk.BOTH, expand=True, **padding)

        self.project_field = DropdownField(container, "Проект", self.project_var)
        self.project_field.grid(row=0, column=0, columnspan=2, sticky=(tk.W + tk.E), pady=(0, 12))

        self.work_field = DropdownField(container, "Вид работы", self.work_type_var)
        self.work_field.grid(row=1, column=0, columnspan=2, sticky=(tk.W + tk.E), pady=(0, 16))

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        timer_label = ttk.Label(container, textvariable=self.timer_var, style="Timesheet.Timer.TLabel")
        timer_label.grid(row=2, column=0, columnspan=2, pady=(8, 12))

        buttons_frame = ttk.Frame(container, style="TFrame")
        buttons_frame.grid(row=3, column=0, columnspan=2, pady=4)

        self._start_button = IconButton(buttons_frame, "play", command=self.start_timer)
        self._start_button.grid(row=0, column=0, padx=6)
        self._pause_button = IconButton(buttons_frame, "pause", command=self.pause_timer)
        self._pause_button.grid(row=0, column=1, padx=6)
        self._stop_button = IconButton(buttons_frame, "stop", command=self.stop_timer)
        self._stop_button.grid(row=0, column=2, padx=6)

        container.rowconfigure(4, weight=1)
        container.rowconfigure(5, minsize=24)  # строка состояния всегда видима

        status_label = ttk.Label(
            container,
            textvariable=self.status_var,
            anchor=tk.W,
            wraplength=600,
            style="Timesheet.Status.TLabel",
        )
        status_label.grid(row=5, column=0, columnspan=2, sticky=(tk.W + tk.E + tk.S), pady=(12, 0))
        self._status_label = status_label

    # ------------------------- Работа с Excel -------------------------
    def _prompt_for_excel(self) -> None:
        """Показать диалог выбора Excel-файла и загрузить справочники."""

        filename = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=(("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")),
        )
        if not filename:
            if not self.config_manager.excel_path:
                messagebox.showinfo("Файл не выбран", "Без Excel файла приложение не сможет работать.")
            return
        try:
            self._load_reference(filename)
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("Ошибка", f"Не удалось загрузить Excel файл:\n{exc}")
            # Очищаем текущие списки и блокируем выбор, чтобы не остались старые данные
            self.projects = []
            self.work_types = []
            self.project_field.set_options([])
            self.work_field.set_options([])
            self._set_inputs_enabled(False)
            return

        self.config_manager.excel_path = filename
        self.config_manager.save()
        self._refresh_status()
        self.status_var.set(f"Файл: {self.config_manager.excel_path}")
        # После удачной загрузки разрешим выбор значений
        self._set_inputs_enabled(True)

    def _load_reference(self, path: str) -> None:
        """Загрузить данные листа 'Справочник' и обновить выпадающие списки."""

        projects, work_types = load_reference_data(path)
        if not projects or not work_types:
            raise ExcelStructureError(
                "В листе 'Справочник' должны быть заполнены столбцы с проектами и видами работ."
            )
        current_project = self.project_var.get()
        current_work_type = self.work_type_var.get()

        self.projects = projects
        self.work_types = work_types
        self.project_field.set_options(self.projects, selected=current_project)
        self.work_field.set_options(self.work_types, selected=current_work_type)

        if current_project in self.projects:
            self.project_var.set(current_project)
        elif self.projects:
            self.project_var.set(self.projects[0])

        if current_work_type in self.work_types:
            self.work_type_var.set(current_work_type)
        elif self.work_types:
            self.work_type_var.set(self.work_types[0])

        self._refresh_status()
        self._adjust_layout_for_content()
        # Списки подгружены — поля доступны
        self._set_inputs_enabled(True)

    def _set_inputs_enabled(self, enabled: bool) -> None:
        """Включить/выключить поля выбора проекта и вида работ."""

        state = "readonly" if enabled else "disabled"
        try:
            self.project_field.combobox.configure(state=state)
            self.work_field.combobox.configure(state=state)
        except Exception:
            pass

    def _refresh_status(self) -> None:
        """Обновить строку состояния: показываем путь к файлу (или отсутствие)."""

        if self.config_manager.excel_path:
            self.status_var.set(f"Файл: {self.config_manager.excel_path}")
        else:
            self.status_var.set("Файл Excel не выбран")

    def _adjust_layout_for_content(self) -> None:
        """Подогнать ширину окна под самые длинные пункты выпадающих списков."""

        longest_width = max(self.project_field.measure_longest_option(), self.work_field.measure_longest_option())
        if longest_width <= 0:
            return
        desired_width = max(440, min(int(longest_width + 260), 1000))
        self.update_idletasks()
        current_height = max(self.winfo_height(), 340)
        self.geometry(f"{desired_width}x{current_height}")
        self.minsize(desired_width, 320)
        self.project_field.refresh_width()
        self.work_field.refresh_width()
        self._status_label.configure(wraplength=max(desired_width - 40, 200))

    # --------------------------- Логика таймера ---------------------------
    def start_timer(self) -> None:
        """Запустить таймер и начать считать время работы."""

        if not self.config_manager.excel_path:
            messagebox.showwarning("Нет файла", "Сначала выберите Excel файл через меню 'Файл'.")
            return
        if not self.project_var.get() or not self.work_type_var.get():
            messagebox.showwarning("Нет данных", "Выберите проект и вид работы.")
            return
        if not self.projects or not self.work_types:
            messagebox.showwarning("Нет данных", "Не удалось загрузить данные из Excel файла.")
            return

        if not self._timer_running:
            self._start_reference = time.perf_counter() - self._elapsed_seconds
            self._timer_running = True
            self._schedule_timer_update()
            # На время отсчёта блокируем изменение полей
            self._set_inputs_enabled(False)

    def pause_timer(self) -> None:
        """Поставить таймер на паузу (не записывает в Excel)."""

        if not self._timer_running:
            return
        self._elapsed_seconds = time.perf_counter() - self._start_reference
        self._timer_running = False
        if self._timer_job is not None:
            self.after_cancel(self._timer_job)
            self._timer_job = None

    def stop_timer(self) -> None:
        """Остановить таймер и записать результат в Excel."""

        if self._timer_running:
            self._elapsed_seconds = time.perf_counter() - self._start_reference
            self._timer_running = False
        if self._timer_job is not None:
            self.after_cancel(self._timer_job)
            self._timer_job = None

        if self._elapsed_seconds <= 0:
            return

        elapsed = self._elapsed_seconds
        self._elapsed_seconds = 0
        self.timer_var.set("00:00:00")

        try:
            append_time_entry(
                self.config_manager.excel_path,
                project=self.project_var.get(),
                work_type=self.work_type_var.get(),
                elapsed_seconds=elapsed,
                finished_at=datetime.now(),
            )
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("Ошибка", f"Не удалось записать данные в Excel:\n{exc}")
            # Разблокируем поля, чтобы пользователь мог скорректировать выбор
            self._set_inputs_enabled(True)
            return

        messagebox.showinfo("Запись добавлена", "Строка успешно записана на лист 'Учет времени'.")
        # После успешной записи — снова разрешаем менять значения
        self._set_inputs_enabled(True)

    def _schedule_timer_update(self) -> None:
        """Планировать регулярное обновление отображения счётчика."""

        self._update_timer_display()
        self._timer_job = self.after(200, self._schedule_timer_update)

    def _update_timer_display(self) -> None:
        """Обновить текст таймера на экране."""

        if self._timer_running:
            self._elapsed_seconds = time.perf_counter() - self._start_reference
        self.timer_var.set(self._format_time(self._elapsed_seconds))

    @staticmethod
    def _format_time(seconds: float) -> str:
        """Форматировать секунды в строку HH:MM:SS."""

        total_seconds = int(seconds)
        hours, remainder = divmod(total_seconds, 3600)
        minutes, secs = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def main() -> None:
    """Точка входа: создать и запустить приложение."""

    app = TimeTrackerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
