"""Tkinter user interface for the Timesheet timer."""

from __future__ import annotations

import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
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
        self.geometry("420x280")
        self.resizable(False, False)

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
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.destroy)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

        self.config(menu=menu_bar)

    def _build_layout(self) -> None:
        padding = {"padx": 12, "pady": 6}

        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, **padding)

        ttk.Label(container, text="Проект:").grid(row=0, column=0, sticky=tk.W)
        self.project_combo = ttk.Combobox(container, textvariable=self.project_var, state="readonly")
        self.project_combo.grid(row=0, column=1, sticky=(tk.W + tk.E))

        ttk.Label(container, text="Вид работы:").grid(row=1, column=0, sticky=tk.W)
        self.work_combo = ttk.Combobox(container, textvariable=self.work_type_var, state="readonly")
        self.work_combo.grid(row=1, column=1, sticky=(tk.W + tk.E))

        container.columnconfigure(1, weight=1)

        timer_label = ttk.Label(container, textvariable=self.timer_var, font=("Segoe UI", 32, "bold"))
        timer_label.grid(row=2, column=0, columnspan=2, pady=(16, 12))

        buttons_frame = ttk.Frame(container)
        buttons_frame.grid(row=3, column=0, columnspan=2, pady=4)

        ttk.Button(buttons_frame, text="Запуск", command=self.start_timer, width=12).grid(
            row=0, column=0, padx=4
        )
        ttk.Button(buttons_frame, text="Пауза", command=self.pause_timer, width=12).grid(row=0, column=1, padx=4)
        ttk.Button(buttons_frame, text="Стоп", command=self.stop_timer, width=12).grid(row=0, column=2, padx=4)

        status_label = ttk.Label(container, textvariable=self.status_var, wraplength=360, foreground="#555")
        status_label.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(12, 0))

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
        self._refresh_status()

    def _refresh_status(self) -> None:
        if self.config_manager.excel_path:
            self.status_var.set(f"Файл Excel: {self.config_manager.excel_path}")
        else:
            self.status_var.set("Файл Excel не выбран")

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
