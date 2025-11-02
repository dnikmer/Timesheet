"""Tkinter user interface for the Timesheet timer."""

from __future__ import annotations

import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from typing import Optional

if __package__ in {None, ""}:  # pragma: no cover - runtime shim for bundled execution
    try:
        from timesheet_app.config import AppConfig
        from timesheet_app.excel_manager import (
            ExcelStructureError,
            append_time_entry,
            load_reference_data,
        )
    except ModuleNotFoundError:  # Running as a loose script without installation
        from config import AppConfig
        from excel_manager import ExcelStructureError, append_time_entry, load_reference_data
else:  # Standard package import path
    from .config import AppConfig
    from .excel_manager import ExcelStructureError, append_time_entry, load_reference_data


class TimeTrackerApp(tk.Tk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Учет рабочего времени")
        self.geometry("480x300")
        self.minsize(420, 280)
        self.resizable(True, True)

        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.config_manager = AppConfig.load()
        self.projects: list[str] = []
        self.work_types: list[str] = []

        self.theme_definitions = {
            "apple_light": {
                "label": "Светлый (Apple)",
                "window_bg": "#f5f5f7",
                "accent": "#007aff",
                "text": "#1c1c1e",
                "muted": "#6c6c70",
                "button_bg": "#ffffff",
                "button_hover": "#e5e5ea",
                "button_active": "#d1d1d6",
                "entry_bg": "#ffffff",
            },
            "windows11": {
                "label": "Windows 11",
                "window_bg": "#f3f3f3",
                "accent": "#2563eb",
                "text": "#1f2937",
                "muted": "#4b5563",
                "button_bg": "#ffffff",
                "button_hover": "#e0e7ff",
                "button_active": "#c7d2fe",
                "entry_bg": "#ffffff",
            },
            "aero_glass": {
                "label": "Aero",
                "window_bg": "#edf5ff",
                "accent": "#0ea5e9",
                "text": "#1f2933",
                "muted": "#52606d",
                "button_bg": "#ffffff",
                "button_hover": "#dbeafe",
                "button_active": "#bfdbfe",
                "entry_bg": "#ffffff",
            },
            "material_ocean": {
                "label": "Material Ocean",
                "window_bg": "#1f2933",
                "accent": "#38bdf8",
                "text": "#f8fafc",
                "muted": "#94a3b8",
                "button_bg": "#334155",
                "button_hover": "#3e4c61",
                "button_active": "#4c566a",
                "entry_bg": "#0f172a",
            },
            "solarized": {
                "label": "Solarized",
                "window_bg": "#fdf6e3",
                "accent": "#268bd2",
                "text": "#073642",
                "muted": "#586e75",
                "button_bg": "#eee8d5",
                "button_hover": "#e4ddc9",
                "button_active": "#d6cdb6",
                "entry_bg": "#ffffff",
            },
            "midnight": {
                "label": "Midnight Blue",
                "window_bg": "#111827",
                "accent": "#60a5fa",
                "text": "#f9fafb",
                "muted": "#9ca3af",
                "button_bg": "#1f2937",
                "button_hover": "#273549",
                "button_active": "#32425b",
                "entry_bg": "#0f172a",
            },
        }

        self._timer_job: Optional[str] = None
        self._timer_running = False
        self._start_reference = 0.0
        self._elapsed_seconds = 0.0

        self.project_var = tk.StringVar()
        self.work_type_var = tk.StringVar()
        self.timer_var = tk.StringVar(value="00:00:00")
        self.status_var = tk.StringVar()

        self.theme_var = tk.StringVar(value=self.config_manager.theme)
        if self.theme_var.get() not in self.theme_definitions:
            self.theme_var.set("apple_light")

        self._build_menu()
        self._build_layout()
        self._apply_theme(self.theme_var.get())
        self.bind("<Configure>", self._on_configure)
        self._refresh_status()

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
    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self)

        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Выбрать файл Excel", command=self._prompt_for_excel)
        file_menu.add_command(label="Инфо", command=self._show_info)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.destroy)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

        view_menu = tk.Menu(menu_bar, tearoff=False)
        for theme_key, data in self.theme_definitions.items():
            view_menu.add_radiobutton(
                label=data["label"],
                variable=self.theme_var,
                value=theme_key,
                command=lambda key=theme_key: self._apply_theme(key),
            )
        menu_bar.add_cascade(label="Вид", menu=view_menu)

        self.config(menu=menu_bar)

    def _build_layout(self) -> None:
        padding = {"padx": 12, "pady": 6}

        container = ttk.Frame(self, style="Main.TFrame")
        container.pack(fill=tk.BOTH, expand=True, **padding)

        ttk.Label(container, text="Проект:", style="Main.TLabel").grid(row=0, column=0, sticky=tk.W)
        self.project_combo = ttk.Combobox(
            container, textvariable=self.project_var, state="readonly", style="Main.TCombobox"
        )
        self.project_combo.grid(row=0, column=1, sticky=(tk.W + tk.E))

        ttk.Label(container, text="Вид работы:", style="Main.TLabel").grid(row=1, column=0, sticky=tk.W)
        self.work_combo = ttk.Combobox(
            container, textvariable=self.work_type_var, state="readonly", style="Main.TCombobox"
        )
        self.work_combo.grid(row=1, column=1, sticky=(tk.W + tk.E))

        container.columnconfigure(1, weight=1)

        self.timer_label = ttk.Label(
            container,
            textvariable=self.timer_var,
            font=("Segoe UI", 32, "bold"),
            anchor=tk.CENTER,
            style="Timer.TLabel",
        )
        self.timer_label.grid(row=2, column=0, columnspan=2, pady=(16, 12), sticky=(tk.W + tk.E))

        buttons_frame = ttk.Frame(container, style="Main.TFrame")
        buttons_frame.grid(row=3, column=0, columnspan=2, pady=4)

        ttk.Button(buttons_frame, text="▶", command=self.start_timer, width=4, style="Timer.TButton").grid(
            row=0, column=0, padx=6
        )
        ttk.Button(buttons_frame, text="⏸", command=self.pause_timer, width=4, style="Timer.TButton").grid(
            row=0, column=1, padx=6
        )
        ttk.Button(buttons_frame, text="⏹", command=self.stop_timer, width=4, style="Timer.TButton").grid(
            row=0, column=2, padx=6
        )

        self.status_label = ttk.Label(
            container,
            textvariable=self.status_var,
            wraplength=360,
            style="Status.TLabel",
            anchor=tk.W,
        )
        self.status_label.grid(row=4, column=0, columnspan=2, sticky=(tk.W + tk.E), pady=(12, 0))

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

    def _load_reference(self, path: str) -> None:
        projects, work_types = load_reference_data(path)
        if not projects or not work_types:
            raise ExcelStructureError(
                "В листе 'Справочник' должны быть заполнены столбцы с проектами и видами работ."
            )
        self.projects = projects
        self.work_types = work_types
        self.project_combo.configure(values=self.projects)
        self.work_combo.configure(values=self.work_types)

        current_project = self.project_var.get()
        current_work_type = self.work_type_var.get()

        if current_project in self.projects:
            self.project_var.set(current_project)
        else:
            self.project_var.set(self.projects[0])

        if current_work_type in self.work_types:
            self.work_type_var.set(current_work_type)
        else:
            self.work_type_var.set(self.work_types[0])
        self._adjust_window_width()
        self._refresh_status()

    def _refresh_status(self) -> None:
        if self.config_manager.excel_path:
            self.status_var.set("Файл Excel готов к использованию")
        else:
            self.status_var.set("Выберите Excel файл через меню 'Файл'.")

    def _show_info(self) -> None:
        if self.config_manager.excel_path:
            messagebox.showinfo("Информация", f"Текущий файл Excel:\n{self.config_manager.excel_path}")
        else:
            messagebox.showinfo("Информация", "Файл Excel не выбран.")

    def _apply_theme(self, theme_key: str) -> None:
        theme = self.theme_definitions.get(theme_key, self.theme_definitions["apple_light"])
        self.style.theme_use("clam")

        window_bg = theme["window_bg"]
        text_color = theme["text"]
        muted = theme["muted"]
        accent = theme["accent"]
        button_bg = theme["button_bg"]
        button_hover = theme["button_hover"]
        button_active = theme["button_active"]
        entry_bg = theme["entry_bg"]

        self.configure(background=window_bg)

        self.style.configure("Main.TFrame", background=window_bg)
        self.style.configure("Main.TLabel", background=window_bg, foreground=text_color)
        self.style.configure("Status.TLabel", background=window_bg, foreground=muted)
        self.style.configure(
            "Timer.TLabel",
            background=window_bg,
            foreground=accent,
            font=("Segoe UI", 32, "bold"),
        )

        self.style.configure(
            "Main.TCombobox",
            fieldbackground=entry_bg,
            background=entry_bg,
            foreground=text_color,
            arrowcolor=accent,
        )
        self.style.map(
            "Main.TCombobox",
            fieldbackground=[("readonly", entry_bg), ("focus", entry_bg)],
            foreground=[("disabled", muted)],
        )

        self.style.configure(
            "Timer.TButton",
            background=button_bg,
            foreground=text_color,
            borderwidth=1,
            focusthickness=2,
            focuscolor=accent,
            padding=(12, 6),
        )
        self.style.map(
            "Timer.TButton",
            background=[("pressed", button_active), ("active", button_hover)],
            foreground=[("disabled", muted)],
            relief=[("pressed", "sunken"), ("!pressed", "flat")],
        )

        self.theme_var.set(theme_key)
        self.config_manager.theme = theme_key
        self.config_manager.save()

        self.update_idletasks()
        wrap_width = max(360, self.winfo_width() - 60)
        self.status_label.configure(wraplength=wrap_width)

    def _adjust_window_width(self) -> None:
        all_items = [*self.projects, *self.work_types]
        if not all_items:
            return

        font_obj = tkfont.nametofont("TkDefaultFont")
        max_width = max(font_obj.measure(item) for item in all_items)
        desired_width = max(480, min(960, max_width + 220))

        self.update_idletasks()
        current_height = max(self.winfo_height(), 300)
        self.geometry(f"{desired_width}x{current_height}")
        self.minsize(desired_width, 280)
        self.status_label.configure(wraplength=desired_width - 60)

    def _on_configure(self, event: tk.Event[tk.Misc]) -> None:  # type: ignore[name-defined]
        if event.widget is self:
            wrap_width = max(360, event.width - 60)
            self.status_label.configure(wraplength=wrap_width)

    # ------------------------------------------------------------------
    # Timer logic
    # ------------------------------------------------------------------
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
