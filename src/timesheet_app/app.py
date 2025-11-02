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
        self.geometry("520x320")
        self.minsize(480, 300)
        self.resizable(True, True)

        self.style = ttk.Style(self)

        self._setup_fonts()
        self._configure_styles()

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

        self._build_menu()
        self._build_layout()
        self.project_var.trace_add("write", self._on_selection_change)
        self.work_type_var.trace_add("write", self._on_selection_change)
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
        file_menu.add_command(label="Текущий файл", command=self._show_info)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.destroy)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

        self.config(menu=menu_bar)

    def _build_layout(self) -> None:
        padding = {"padx": 18, "pady": 12}

        content = ttk.Frame(self, style="Mac.Content.TFrame")
        content.pack(fill=tk.BOTH, expand=True, padx=24, pady=(20, 12))

        ttk.Label(content, text="Проект:", style="Mac.TLabel").grid(row=0, column=0, sticky=tk.W)
        self.project_var.set("Выберите проект")
        self.project_menu = ttk.OptionMenu(
            content,
            self.project_var,
            self.project_var.get(),
        )
        self.project_menu.configure(style="Mac.OptionMenu.TMenubutton")
        self.project_menu.configure(width=24)
        self.project_menu.grid(row=0, column=1, sticky="ew", **padding)
        self._initialise_menu(self.project_menu)

        ttk.Label(content, text="Вид работы:", style="Mac.TLabel").grid(row=1, column=0, sticky=tk.W)
        self.work_type_var.set("Выберите вид работы")
        self.work_menu = ttk.OptionMenu(
            content,
            self.work_type_var,
            self.work_type_var.get(),
        )
        self.work_menu.configure(style="Mac.OptionMenu.TMenubutton")
        self.work_menu.configure(width=24)
        self.work_menu.grid(row=1, column=1, sticky="ew", **padding)
        self._initialise_menu(self.work_menu)

        content.columnconfigure(0, weight=0)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(2, weight=1)

        self.timer_label = ttk.Label(
            content,
            textvariable=self.timer_var,
            anchor=tk.CENTER,
            style="Mac.Timer.TLabel",
        )
        self.timer_label.grid(row=2, column=0, columnspan=2, pady=(24, 18), sticky="nsew")

        buttons_frame = ttk.Frame(content, style="Mac.Content.TFrame")
        buttons_frame.grid(row=3, column=0, columnspan=2, pady=(0, 8))

        ttk.Button(
            buttons_frame,
            text="▶",
            command=self.start_timer,
            width=3,
            style="Mac.Toolbar.TButton",
        ).grid(row=0, column=0, padx=8)
        ttk.Button(
            buttons_frame,
            text="⏸",
            command=self.pause_timer,
            width=3,
            style="Mac.Toolbar.TButton",
        ).grid(row=0, column=1, padx=8)
        ttk.Button(
            buttons_frame,
            text="⏹",
            command=self.stop_timer,
            width=3,
            style="Mac.Toolbar.TButton",
        ).grid(row=0, column=2, padx=8)

        for column_index in range(3):
            buttons_frame.columnconfigure(column_index, weight=1)

        status_frame = ttk.Frame(self, style="Mac.Status.TFrame")
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            style="Mac.Status.TLabel",
            anchor=tk.W,
        )
        self.status_label.pack(side=tk.LEFT, padx=18, pady=6)

    def _setup_fonts(self) -> None:
        available = {family for family in tkfont.families()}

        def pick_font(candidates: tuple[str, ...], fallback: str = "Helvetica") -> str:
            for name in candidates:
                if name in available:
                    return name
            return fallback

        primary_family = pick_font(("SF Pro Text", "San Francisco", "Helvetica Neue", "Helvetica", "Arial"))
        display_family = pick_font(("SF Pro Display", "SF Pro Text", "Helvetica Neue", "Helvetica", "Arial"))

        self.base_font = tkfont.Font(family=primary_family, size=13)
        self.small_font = tkfont.Font(family=primary_family, size=11)
        self.timer_font = tkfont.Font(family=display_family, size=44, weight="bold")

        self.option_add("*Font", self.base_font)
        self.option_add("*Menu.font", self.base_font)

    def _configure_styles(self) -> None:
        self.style.theme_use("clam")

        mac_bg = "#e8e8ed"
        panel_bg = "#f5f5f7"
        accent = "#0a84ff"
        text = "#1c1c1e"
        muted = "#636366"
        border = "#c7c7cc"
        status_bg = "#d1d1d6"

        self.configure(background=mac_bg)

        self.style.configure("Mac.TFrame", background=mac_bg)
        self.style.configure("Mac.Content.TFrame", background=mac_bg)
        self.style.configure("Mac.TLabel", background=mac_bg, foreground=text, font=self.base_font)
        self.style.configure("Mac.Timer.TLabel", background=mac_bg, foreground=accent, font=self.timer_font)

        option_style = "Mac.OptionMenu.TMenubutton"
        self.style.configure(
            option_style,
            background=panel_bg,
            foreground=text,
            font=self.base_font,
            borderwidth=1,
            bordercolor=border,
            relief="flat",
            padding=(16, 6),
        )
        self.style.map(
            option_style,
            background=[("active", "#ebeaf0")],
            foreground=[("disabled", muted)],
            bordercolor=[("focus", accent)],
        )

        self.style.configure(
            "Mac.Toolbar.TButton",
            background=panel_bg,
            foreground=text,
            font=self.base_font,
            borderwidth=1,
            bordercolor=border,
            padding=(12, 6),
            focusthickness=2,
            focuscolor=accent,
        )
        self.style.map(
            "Mac.Toolbar.TButton",
            background=[("pressed", "#dcdcde"), ("active", "#ebeaf0")],
            foreground=[("disabled", muted)],
            relief=[("pressed", "sunken"), ("!pressed", "flat")],
        )

        self.style.configure(
            "Mac.Status.TFrame",
            background=status_bg,
            borderwidth=1,
            relief="flat",
            bordercolor=border,
        )
        self.style.configure("Mac.Status.TLabel", background=status_bg, foreground=text, font=self.small_font)

    def _initialise_menu(self, option: ttk.OptionMenu) -> None:
        menu: tk.Menu = option["menu"]
        menu.delete(0, "end")
        menu.configure(tearoff=False, font=self.base_font)
        menu.add_command(label="Нет данных", state="disabled")
        option.state(["disabled"])

    def _update_option_menu(
        self, option: ttk.OptionMenu, variable: tk.StringVar, values: list[str]
    ) -> None:
        menu: tk.Menu = option["menu"]
        menu.delete(0, "end")
        menu.configure(tearoff=False, font=self.base_font)

        if not values:
            variable.set("Нет данных")
            menu.add_command(label="Нет данных", state="disabled")
            option.state(["disabled"])
            return

        option.state(["!disabled"])
        current = variable.get()
        if current not in values:
            variable.set(values[0])
        for value in values:
            menu.add_radiobutton(label=value, value=value, variable=variable)

    def _on_selection_change(self, *_: object) -> None:
        self.after_idle(self._adjust_window_width)

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
        self._update_option_menu(self.project_menu, self.project_var, self.projects)
        self._update_option_menu(self.work_menu, self.work_type_var, self.work_types)
        self._adjust_window_width()
        self._refresh_status()

    def _refresh_status(self) -> None:
        if self.config_manager.excel_path:
            self.status_var.set("Файл Excel готов к использованию")
        else:
            self.status_var.set("Выберите Excel файл через меню 'Файл'.")

    def _show_info(self) -> None:
        if self.config_manager.excel_path:
            messagebox.showinfo("Текущий файл", f"Текущий файл Excel:\n{self.config_manager.excel_path}")
        else:
            messagebox.showinfo("Текущий файл", "Файл Excel не выбран.")

    def _adjust_window_width(self) -> None:
        all_items = [*self.projects, *self.work_types]
        if not all_items:
            return

        font_obj = self.base_font
        max_width = max(font_obj.measure(item) for item in all_items)
        desired_width = max(520, min(1024, max_width + 280))

        char_width = max(font_obj.measure("0"), 1)
        menu_chars = min(48, max(18, (max_width + 60) // char_width))
        self.project_menu.configure(width=menu_chars)
        self.work_menu.configure(width=menu_chars)

        self.update_idletasks()
        current_height = max(self.winfo_height(), 320)
        self.geometry(f"{desired_width}x{current_height}")
        self.minsize(desired_width, 300)

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
