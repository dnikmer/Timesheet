"""Tkinter user interface for the Timesheet timer."""

from __future__ import annotations

import math
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


class DropdownField(ttk.Frame):
    """Labeled dropdown field that mimics a modern select control."""

    def __init__(self, parent: tk.Widget, label_text: str, variable: tk.StringVar) -> None:
        super().__init__(parent)
        self.variable = variable
        self._choices: list[str] = []
        try:
            self._menu_font = font.Font(family="Proxima Nova", size=9)
        except tk.TclError:
            self._menu_font = font.nametofont("TkMenuFont")
            self._menu_font.configure(family="Proxima Nova", size=9)

        self.label = ttk.Label(self, text=label_text, style="Timesheet.Label")
        self.label.pack(anchor=tk.W, pady=(0, 4))

        self.option_menu = ttk.OptionMenu(self, variable, variable.get())
        self.option_menu.configure(style="Timesheet.OptionMenu.TMenubutton", width=20)
        self.option_menu.pack(fill=tk.X)
        self.option_menu["menu"].configure(font=self._menu_font)

    def set_options(self, options: list[str], *, selected: Optional[str] = None) -> None:
        """Populate the dropdown with the provided options."""

        self._choices = options[:]
        menu = self.option_menu["menu"]
        menu.delete(0, "end")

        if options:
            for option in options:
                menu.add_command(label=option, command=lambda value=option: self.variable.set(value))

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
            self.option_menu.configure(width=20)
            return

        max_pixels = max(self._menu_font.measure(item) for item in self._choices)
        average_char = max(self._menu_font.measure("0"), 1)
        width_chars = max(20, min(int(math.ceil((max_pixels + 24) / average_char)), 64))
        self.option_menu.configure(width=width_chars)

    def measure_longest_option(self) -> int:
        """Return the pixel width of the longest option."""

        if not self._choices:
            return 0
        return max(self._menu_font.measure(item) for item in self._choices)


class IconButton(ttk.Frame):
    """Canvas-based icon button with a crisp square outline and glyph."""

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
            width=60,
            height=60,
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
        outline = "#1c1c1c" if not hover else "#000000"
        fill_color = "#ffffff" if not hover else "#f2f2f2"
        self._canvas.create_rectangle(10, 10, 50, 50, outline=outline, width=3, fill=fill_color)
        glyph_color = "#1c1c1c"
        if self.icon == "play":
            self._canvas.create_polygon(28, 22, 28, 38, 42, 30, fill=glyph_color, outline=glyph_color)
        elif self.icon == "pause":
            self._canvas.create_rectangle(24, 22, 30, 38, fill=glyph_color, outline=glyph_color)
            self._canvas.create_rectangle(32, 22, 38, 38, fill=glyph_color, outline=glyph_color)
        elif self.icon == "stop":
            self._canvas.create_rectangle(24, 24, 38, 38, fill=glyph_color, outline=glyph_color)

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
        self.configure(background="#f5f5f5")

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
            default_font = font.nametofont("TkDefaultFont")
            default_font.configure(family="Proxima Nova", size=9)
        except tk.TclError:
            pass

        try:
            menu_font = font.nametofont("TkMenuFont")
            menu_font.configure(family="Proxima Nova", size=9)
        except tk.TclError:
            pass

        style.configure("TFrame", background="#f5f5f5")
        style.configure("Timesheet.Label", font=("Proxima Nova", 9), foreground="#1f1f1f", background="#f5f5f5")
        style.configure("Timesheet.Timer.TLabel", font=("Proxima Nova", 32, "bold"), foreground="#1f1f1f", background="#f5f5f5")
        style.configure(
            "Timesheet.OptionMenu.TMenubutton",
            font=("Proxima Nova", 9),
            padding=(14, 8),
            relief="flat",
            borderwidth=1,
            background="#ffffff",
            foreground="#1f1f1f",
            bordercolor="#7c3aed",
        )
        style.map(
            "Timesheet.OptionMenu.TMenubutton",
            background=[("active", "#f4f0ff"), ("pressed", "#ede7ff")],
            bordercolor=[("focus", "#7c3aed"), ("active", "#7c3aed")],
            foreground=[("disabled", "#9f9f9f")],
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
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.destroy)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

        self.config(menu=menu_bar)

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
