"""Tkinter user interface for the Timesheet timer."""

from __future__ import annotations

import math
import sys
import time
import os
import subprocess
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font, messagebox, ttk
from typing import Callable, Optional

if __package__ in {None, ""}:  # pragma: no cover - runtime shim for bundled execution
    try:
        from timesheet_app.config import AppConfig
        from timesheet_app.excel_manager import (
            ExcelStructureError,
            REFERENCE_SHEET,
            TIMESHEET_SHEET,
            append_time_entry,
            load_reference_data,
        )
    except ModuleNotFoundError:  # Running as a loose script without installation
        from config import AppConfig
        from excel_manager import ExcelStructureError, append_time_entry, load_reference_data
else:  # Standard package import path
    from .config import AppConfig
    from .excel_manager import (
        ExcelStructureError,
        REFERENCE_SHEET,
        TIMESHEET_SHEET,
        append_time_entry,
        load_reference_data,
    )


def _asset_path(filename: str) -> str:
    """Return the absolute path to an asset bundled with the app."""

    base_path = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)  # type: ignore[attr-defined]
    return str(Path(base_path) / "assets" / filename)


class DropdownField(ttk.Frame):
    """Labeled dropdown field that mimics a modern select control."""

    def __init__(self, parent: tk.Widget, label_text: str, variable: tk.StringVar) -> None:
        super().__init__(parent)
        self.variable = variable
        self._choices: list[str] = []
        try:
            _family = "Segoe UI" if sys.platform.startswith("win") else "Arial"
            self._menu_font = font.Font(family=_family, size=9)
        except tk.TclError:
            self._menu_font = font.nametofont("TkMenuFont")
            try:
                _family = "Segoe UI" if sys.platform.startswith("win") else "Arial"
                self._menu_font.configure(family=_family, size=9)
            except tk.TclError:
                # Оставляем системный шрифт по умолчанию
                pass

        self.label = ttk.Label(self, text=label_text, style="Timesheet.Label")
        self.label.pack(anchor=tk.W, pady=(0, 4))

        self.combobox = ttk.Combobox(
            self,
            textvariable=variable,
            state="readonly",
            width=20,
            style="Timesheet.TCombobox",
        )
        self.combobox.pack(fill=tk.X)

    def set_options(self, options: list[str], *, selected: Optional[str] = None) -> None:
        """Populate the dropdown with the provided options."""

        self._choices = options[:]
        # menu removed (OptionMenu -> Combobox)
        # no explicit clearing needed for Combobox values

        if not options:
            self.combobox.configure(state="disabled", values=[])
            self.variable.set("")
            self.refresh_width()
            return

        if options:
            self.combobox.configure(state="readonly", values=options)
            # values are already set on Combobox; no per-item commands needed

            if selected in options:
                self.variable.set(selected)
            elif self.variable.get() in options:
                # Keep the previously selected value.
                pass
            else:
                self.variable.set(options[0])
        else:
            placeholder = "Нет данных"
            menu.add_command(label=placeholder, state="disabled")
            self.variable.set("")

        self.refresh_width()

    def refresh_width(self) -> None:
        """Refresh the visible width based on the longest option."""

        if not self._choices:
            self.combobox.configure(width=20)
            return

        max_pixels = max(self._menu_font.measure(item) for item in self._choices)
        average_char = max(self._menu_font.measure("0"), 1)
        width_chars = max(20, min(int(math.ceil((max_pixels + 24) / average_char)), 64))
        self.combobox.configure(width=width_chars)

    def measure_longest_option(self) -> int:
        """Return the pixel width of the longest option."""

        if not self._choices:
            return 0
        return max(self._menu_font.measure(item) for item in self._choices)


class IconButton(ttk.Frame):
    """Image-based icon button that swaps graphics on hover."""

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
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Учет рабочего времени")
        self.geometry("440x300")
        self.minsize(420, 280)
        self.resizable(True, True)
        self.configure(background="#f5f5f5")
        # Переустановим заголовок окна корректной Unicode-строкой
        self.title("\u0423\u0447\u0435\u0442 \u0440\u0430\u0431\u043e\u0447\u0435\u0433\u043e \u0432\u0440\u0435\u043c\u0435\u043d\u0438")

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
        # Исправляем подписи меню на корректные русские после создания меню
        # Переименуем/перестроим пункты меню уже после инициализации UI,
        # чтобы корректно заменить метки каскадов на Unicode
        self.after(0, self._fix_menu_labels_for_locale)
        self._build_layout()
        self._refresh_status()
        # Отобразим выбранный файл в строке состояния
        if self.config_manager.excel_path:
            self.status_var.set(f"\u0424\u0430\u0439\u043b: {self.config_manager.excel_path}")
        else:
            self.status_var.set("\u0424\u0430\u0439\u043b Excel \u043d\u0435 \u0432\u044b\u0431\u0440\u0430\u043d")

        if self.config_manager.excel_path:
            try:
                self._load_reference(self.config_manager.excel_path)
            except Exception as exc:  # pylint: disable=broad-except
                messagebox.showerror("Ошибка", f"Не удалось загрузить Excel файл:\n{exc}")
                self.config_manager.excel_path = None
                self.config_manager.save()
                self._refresh_status()

        if not self.projects or not self.work_types:
            self.after(100, self._prompt_for_excel)

    # ------------------------------------------------------------------
    # UI construction helpers
    # ------------------------------------------------------------------
    def _configure_styles(self) -> None:
        """Configure ttk styles and default fonts for the app."""

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        try:
            _family = "Segoe UI" if sys.platform.startswith("win") else "Arial"
            default_font = font.nametofont("TkDefaultFont")
            default_font.configure(family=_family, size=9)
        except tk.TclError:
            pass

        try:
            _family = "Segoe UI" if sys.platform.startswith("win") else "Arial"
            menu_font = font.nametofont("TkMenuFont")
            menu_font.configure(family=_family, size=9)
        except tk.TclError:
            pass

        style.configure("TFrame", background="#f5f5f5")
        style.configure("Timesheet.Label", font=("Proxima Nova", 9), foreground="#1f1f1f", background="#f5f5f5")
        style.configure("Timesheet.Timer.TLabel", font=("Proxima Nova", 32, "bold"), foreground="#1f1f1f", background="#f5f5f5")
        # Стиль для Combobox
        style.configure(
            "Timesheet.TCombobox",
            padding=(6, 2),
            relief="flat",
            borderwidth=1,
            foreground="#1f1f1f",
            fieldbackground="#ffffff",
            background="#ffffff",
        )
        style.map(
            "Timesheet.TCombobox",
            fieldbackground=[("readonly", "#ffffff"), ("active", "#f4f0ff")],
            background=[("active", "#f4f0ff")],
            foreground=[("disabled", "#9f9f9f")],
        )
        style.configure(
            "Timesheet.IconButton.TButton",
            background="#ffffff",
            relief="flat",
            padding=0,
            borderwidth=0,
        )
        style.map(
            "Timesheet.IconButton.TButton",
            background=[("active", "#f4f0ff")],
            relief=[("pressed", "flat")],
        )
        style.configure(
            "Timesheet.Status.TLabel",
            font=("Proxima Nova", 7),
            foreground="#555555",
            background="#f5f5f5",
        )

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self)

        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Выбрать файл Excel", command=self._prompt_for_excel)
        file_menu.add_command(label="Текущий файл", command=self._show_current_file)
        file_menu.insert_command(1, label="Открыть текущий файл", command=self._open_current_file)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.destroy)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

        # Меню "Помощь" справа от "Файл"
        help_menu = tk.Menu(menu_bar, tearoff=False)
        help_menu.add_command(label="�����⨢�� � Excel...", command=self._show_excel_requirements)
        menu_bar.add_cascade(label="�������", menu=help_menu)

        self.config(menu=menu_bar)

    def _insert_open_current_file_menu(self) -> None:
        """Insert 'Открыть текущий файл' into the File menu after the first item."""
        try:
            menubar = self.nametowidget(self["menu"])  # type: ignore[assignment]
            submenu_name = menubar.entrycget(0, "menu")  # first cascade (Файл)
            file_menu = self.nametowidget(submenu_name)
            file_menu.insert_command(1, label="Открыть текущий файл", command=self._open_current_file)
        except Exception:
            # If menu structure differs, silently ignore.
            pass

    def _open_current_file(self) -> None:
        path = self.config_manager.excel_path
        if not path:
            messagebox.showinfo("Текущий файл", "Файл Excel не выбран")
            return
        try:
            if not Path(path).exists():
                messagebox.showwarning(
                    "Файл не найден",
                    "Указанный файл не существует. Выберите файл Excel заново.",
                )
                return
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{exc}")

    def _show_excel_requirements(self) -> None:
        """Показать требования к структуре Excel-файла."""
        msg = (
            f"���� Excel ������ ��������� ��� '{REFERENCE_SHEET}'.\n"
            "� ��� ���� ���᪮��: ������ � ��� ����� (�� 2-�� ������).\n\n"
            f"����� ����� ��� '{TIMESHEET_SHEET}', ��� ���������� ���������� ������:\n"
            "- ����\n- ������\n- ��� �����\n- ����������� (������ �����)."
        )
        messagebox.showinfo("���������� � Excel", msg)

    def _build_layout(self) -> None:
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
        # Гарантируем, что строка состояния всегда видна
        container.rowconfigure(5, minsize=24)

        status_label = ttk.Label(
            container,
            textvariable=self.status_var,
            anchor=tk.W,
            wraplength=600,
            style="Timesheet.Status.TLabel",
        )
        status_label.grid(row=5, column=0, columnspan=2, sticky=(tk.W + tk.E + tk.S), pady=(12, 0))
        self._status_label = status_label

    # ------------------------------------------------------------------
    # Excel helpers
    # ------------------------------------------------------------------
    def _prompt_for_excel(self) -> None:
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
            return

        self.config_manager.excel_path = filename
        self.config_manager.save()
        self._refresh_status()
        # Немедленно отобразим путь к выбранному файлу в строке состояния
        self.status_var.set(f"\u0424\u0430\u0439\u043b: {self.config_manager.excel_path}")

    def _load_reference(self, path: str) -> None:
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

    def _refresh_status(self) -> None:
        if self.config_manager.excel_path:
            self.status_var.set("Файл Excel готов к использованию")
        else:
            self.status_var.set("Файл Excel не выбран")

    def _show_current_file(self) -> None:
        if self.config_manager.excel_path:
            messagebox.showinfo("Текущий файл", self.config_manager.excel_path)
        else:
            messagebox.showinfo("Текущий файл", "Файл Excel не выбран")

    def _adjust_layout_for_content(self) -> None:
        longest_width = max(
            self.project_field.measure_longest_option(),
            self.work_field.measure_longest_option(),
        )
        if longest_width <= 0:
            return
        desired_width = max(440, min(int(longest_width + 260), 1000))
        self.update_idletasks()
        current_height = max(self.winfo_height(), 300)
        self.geometry(f"{desired_width}x{current_height}")
        self.minsize(desired_width, 280)
        self.project_field.refresh_width()
        self.work_field.refresh_width()
        self._status_label.configure(wraplength=max(desired_width - 40, 200))

    def _fix_menu_labels_for_locale(self) -> None:
        """Полностью перестроить меню с корректными Unicode‑подписями.

        Убираем подпункт «Текущий файл», как просили, оставляя:
        Файл → Выбрать файл Excel, Открыть текущий файл, Выход
        Помощь → Требования к Excel‑файлу
        """
        try:
            menu_bar = tk.Menu(self)

            # Файл
            file_menu = tk.Menu(menu_bar, tearoff=False)
            file_menu.add_command(
                label="\u0412\u044b\u0431\u0440\u0430\u0442\u044c \u0444\u0430\u0439\u043b Excel",
                command=self._prompt_for_excel,
            )
            file_menu.add_command(
                label="\u041e\u0442\u043a\u0440\u044b\u0442\u044c \u0442\u0435\u043a\u0443\u0449\u0438\u0439 \u0444\u0430\u0439\u043b",
                command=self._open_current_file,
            )
            file_menu.add_separator()
            file_menu.add_command(label="\u0412\u044b\u0445\u043e\u0434", command=self.destroy)
            menu_bar.add_cascade(label="\u0424\u0430\u0439\u043b", menu=file_menu)

            # Помощь
            help_menu = tk.Menu(menu_bar, tearoff=False)
            help_menu.add_command(
                label="\u0422\u0440\u0435\u0431\u043e\u0432\u0430\u043d\u0438\u044f \u043a Excel-\u0444\u0430\u0439\u043b\u0443...",
                command=self._show_excel_requirements,
            )
            menu_bar.add_cascade(label="\u041f\u043e\u043c\u043e\u0449\u044c", menu=help_menu)

            self.config(menu=menu_bar)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Timer logic
    # ------------------------------------------------------------------

    # Дублирующее определение для корректного текста диалога (перекроет предыдущий метод)
    def _show_excel_requirements(self) -> None:  # type: ignore[override]
        msg = (
            f"\u0424\u0430\u0439\u043b Excel \u0434\u043e\u043b\u0436\u0435\u043d \u0441\u043e\u0434\u0435\u0440\u0436\u0430\u0442\u044c \u043b\u0438\u0441\u0442 '{REFERENCE_SHEET}'.\n"
            "\u0412 \u043d\u0451\u043c \u0434\u0432\u0430 \u0441\u0442\u043e\u043b\u0431\u0446\u0430: \u041f\u0440\u043e\u0435\u043a\u0442 \u0438 \u0412\u0438\u0434 \u0440\u0430\u0431\u043e\u0442.\n\n"
            f"\u0422\u0430\u043a\u0436\u0435 \u043d\u0443\u0436\u0435\u043d \u043b\u0438\u0441\u0442 '{TIMESHEET_SHEET}', \u043a\u0443\u0434\u0430 \u0434\u043e\u0431\u0430\u0432\u043b\u044f\u044e\u0442\u0441\u044f \u0437\u0430\u043f\u0438\u0441\u0438:\n"
            "- \u0414\u0430\u0442\u0430\n- \u041f\u0440\u043e\u0435\u043a\u0442\n- \u0412\u0438\u0434 \u0440\u0430\u0431\u043e\u0442\n- \u0414\u043b\u0438\u0442\u0435\u043b\u044c\u043d\u043e\u0441\u0442\u044c (\u0444\u043e\u0440\u043c\u0430\u0442 \u0412\u0440\u0435\u043c\u044f)."
        )
        messagebox.showinfo("\u0422\u0440\u0435\u0431\u043e\u0432\u0430\u043d\u0438\u044f \u043a Excel-\u0444\u0430\u0439\u043b\u0443", msg)
    def start_timer(self) -> None:
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

    def pause_timer(self) -> None:
        if not self._timer_running:
            return
        self._elapsed_seconds = time.perf_counter() - self._start_reference
        self._timer_running = False
        if self._timer_job is not None:
            self.after_cancel(self._timer_job)
            self._timer_job = None

    def stop_timer(self) -> None:
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
            return

        messagebox.showinfo(
            "Время сохранено",
            "Запись успешно добавлена в лист 'Учет времени'.",
        )

    def _schedule_timer_update(self) -> None:
        self._update_timer_display()
        self._timer_job = self.after(200, self._schedule_timer_update)

    def _update_timer_display(self) -> None:
        if self._timer_running:
            self._elapsed_seconds = time.perf_counter() - self._start_reference
        self.timer_var.set(self._format_time(self._elapsed_seconds))

    @staticmethod
    def _format_time(seconds: float) -> str:
        total_seconds = int(seconds)
        hours, remainder = divmod(total_seconds, 3600)
        minutes, secs = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def main() -> None:
    app = TimeTrackerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
