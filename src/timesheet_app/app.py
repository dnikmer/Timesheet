"""Tkinter user interface for the Timesheet timer."""

from __future__ import annotations

import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, font, messagebox, ttk
from typing import Callable, Optional

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


class IconButton(ttk.Frame):
    """Canvas-based icon button with square outline and simple glyph."""

    def __init__(self, parent: tk.Widget, icon: str, command: Optional[Callable[[], None]]) -> None:
        super().__init__(parent)
        self.command = command
        self.icon = icon
        background = "#f4f4f4"
        if hasattr(parent, "cget"):
            try:
                background = parent.cget("background")
            except tk.TclError:
                style_name = ""
                try:
                    style_name = parent.cget("style")
                except tk.TclError:
                    style_name = ""
                style = ttk.Style()
                if style_name:
                    background = style.lookup(style_name, "background", default=background)
                else:
                    background = style.lookup(parent.winfo_class(), "background", default=background)
        self._canvas = tk.Canvas(
            self,
            width=54,
            height=54,
            highlightthickness=0,
            borderwidth=0,
            background=background,
        )
        self._canvas.pack(fill=tk.BOTH, expand=True)
        self._canvas.configure(cursor="hand2")
        self._draw_icon()
        self._canvas.bind("<Button-1>", self._on_click)
        self._canvas.bind("<Enter>", self._on_enter)
        self._canvas.bind("<Leave>", self._on_leave)

    def _draw_icon(self, hover: bool = False) -> None:
        self._canvas.delete("all")
        base_outline = "#222" if not hover else "#111"
        fill_color = "#fafafa" if not hover else "#f0f0f0"
        self._canvas.create_rectangle(8, 8, 46, 46, outline=base_outline, width=2, fill=fill_color)
        glyph_color = "#222"
        if self.icon == "play":
            self._canvas.create_polygon(23, 18, 23, 38, 38, 28, fill=glyph_color, outline=glyph_color)
        elif self.icon == "pause":
            self._canvas.create_rectangle(20, 18, 26, 38, fill=glyph_color, outline=glyph_color)
            self._canvas.create_rectangle(30, 18, 36, 38, fill=glyph_color, outline=glyph_color)
        elif self.icon == "stop":
            self._canvas.create_rectangle(20, 18, 36, 34, fill=glyph_color, outline=glyph_color)

    def _on_click(self, _event: tk.Event) -> None:  # type: ignore[override]
        if callable(self.command):
            self.command()

    def _on_enter(self, _event: tk.Event) -> None:  # type: ignore[override]
        self._draw_icon(hover=True)

    def _on_leave(self, _event: tk.Event) -> None:  # type: ignore[override]
        self._draw_icon(hover=False)


class TimeTrackerApp(tk.Tk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Учет рабочего времени")
        self.geometry("440x300")
        self.minsize(420, 280)
        self.resizable(True, True)

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
        file_menu.add_command(label="Текущий файл", command=self._show_current_file)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.destroy)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

        self.config(menu=menu_bar)

    def _build_layout(self) -> None:
        padding = {"padx": 16, "pady": 8}

        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, **padding)

        ttk.Label(container, text="Проект:").grid(row=0, column=0, columnspan=2, sticky=tk.W)
        self.project_combo = ttk.Combobox(container, textvariable=self.project_var, state="readonly")
        self.project_combo.grid(row=1, column=0, columnspan=2, sticky=(tk.W + tk.E), pady=(0, 8))

        ttk.Label(container, text="Вид работы:").grid(row=2, column=0, columnspan=2, sticky=tk.W)
        self.work_combo = ttk.Combobox(container, textvariable=self.work_type_var, state="readonly")
        self.work_combo.grid(row=3, column=0, columnspan=2, sticky=(tk.W + tk.E), pady=(0, 12))

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        timer_label = ttk.Label(container, textvariable=self.timer_var, font=("Segoe UI", 32, "bold"))
        timer_label.grid(row=4, column=0, columnspan=2, pady=(8, 12))

        buttons_frame = ttk.Frame(container)
        buttons_frame.grid(row=5, column=0, columnspan=2, pady=4)

        self._start_button = IconButton(buttons_frame, "play", command=self.start_timer)
        self._start_button.grid(row=0, column=0, padx=6)
        self._pause_button = IconButton(buttons_frame, "pause", command=self.pause_timer)
        self._pause_button.grid(row=0, column=1, padx=6)
        self._stop_button = IconButton(buttons_frame, "stop", command=self.stop_timer)
        self._stop_button.grid(row=0, column=2, padx=6)

        container.rowconfigure(6, weight=1)

        status_label = ttk.Label(
            container,
            textvariable=self.status_var,
            foreground="#555",
            anchor=tk.W,
            wraplength=600,
        )
        status_label.grid(row=7, column=0, columnspan=2, sticky=(tk.W + tk.E + tk.S), pady=(12, 0))

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
        longest_items = self.projects + self.work_types
        if not longest_items:
            return
        font_name = self.project_combo.cget("font") or "TkDefaultFont"
        try:
            default_font = font.nametofont(font_name)
        except tk.TclError:
            try:
                default_font = font.nametofont("TkDefaultFont")
            except tk.TclError:
                default_font = font.Font(self, font=("Segoe UI", 10))
        max_width = max(default_font.measure(item) for item in longest_items)
        # Provide extra padding for dropdown arrow and margins
        desired_width = max(420, min(max_width + 160, 900))
        self.update_idletasks()
        current_height = max(self.winfo_height(), 300)
        self.geometry(f"{desired_width}x{current_height}")
        self.minsize(desired_width, 280)
        combo_width_chars = max(len(max(longest_items, key=len)) + 2, 20)
        self.project_combo.configure(width=combo_width_chars)
        self.work_combo.configure(width=combo_width_chars)

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
