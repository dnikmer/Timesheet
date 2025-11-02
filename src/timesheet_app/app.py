"""Tkinter user interface for the Timesheet timer."""

from __future__ import annotations

import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
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


class MacIconButton(tk.Canvas):
    """Custom circular button mimicking macOS toolbar controls."""

    def __init__(
        self,
        master: tk.Widget,
        *,
        command: Callable[[], None],
        icon: str,
        palette: dict[str, str],
        diameter: int = 48,
    ) -> None:
        super().__init__(
            master,
            width=diameter,
            height=diameter,
            highlightthickness=0,
            bd=0,
            background=palette["panel_bg"],
            cursor="hand2",
        )
        self._command = command
        self._icon = icon
        self._palette = palette
        self._diameter = diameter
        self._hover = False
        self._pressed = False

        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

        self._redraw()

    # ------------------------------------------------------------------
    def _on_enter(self, _event: tk.Event) -> None:
        self._hover = True
        self._redraw()

    def _on_leave(self, _event: tk.Event) -> None:
        self._hover = False
        self._pressed = False
        self._redraw()

    def _on_press(self, _event: tk.Event) -> None:
        self._pressed = True
        self._redraw()

    def _on_release(self, _event: tk.Event) -> None:
        was_pressed = self._pressed
        self._pressed = False
        self._redraw()
        if was_pressed and callable(self._command):
            self.after_idle(self._command)

    # ------------------------------------------------------------------
    def _redraw(self) -> None:
        self.delete("all")
        self.configure(background=self._palette["panel_bg"])
        radius = self._diameter - 8
        offset = (self._diameter - radius) // 2

        base = self._palette["toolbar_base"]
        hover = self._palette["toolbar_hover"]
        active = self._palette["toolbar_active"]
        stroke = self._palette["toolbar_outline"]
        highlight = self._palette["toolbar_highlight"]

        fill = base
        if self._pressed:
            fill = active
        elif self._hover:
            fill = hover

        self.create_oval(
            offset,
            offset,
            offset + radius,
            offset + radius,
            fill=fill,
            outline=stroke,
            width=1,
        )

        # Subtle highlight
        self.create_oval(
            offset + 2,
            offset + 2,
            offset + radius - 2,
            offset + radius // 2,
            fill=highlight,
            outline="",
        )

        icon_color = self._palette["accent"]
        center = self._diameter // 2
        glyph_size = max(10, radius - 20)

        if self._icon == "play":
            self.create_polygon(
                center - glyph_size // 2,
                center - glyph_size,
                center - glyph_size // 2,
                center + glyph_size,
                center + glyph_size,
                center,
                fill=icon_color,
                outline="",
            )
        elif self._icon == "pause":
            bar_width = max(4, glyph_size // 2)
            spacing = bar_width // 2
            self._create_round_rect(
                center - spacing - bar_width,
                center - glyph_size,
                center - spacing,
                center + glyph_size,
                radius=4,
                fill=icon_color,
            )
            self._create_round_rect(
                center + spacing,
                center - glyph_size,
                center + spacing + bar_width,
                center + glyph_size,
                radius=4,
                fill=icon_color,
            )
        else:  # stop
            side = glyph_size * 1.4
            self._create_round_rect(
                center - side / 2,
                center - side / 2,
                center + side / 2,
                center + side / 2,
                radius=6,
                fill=icon_color,
            )

    # ------------------------------------------------------------------
    def _create_round_rect(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        *,
        radius: float,
        fill: str,
    ) -> None:
        """Draw a rounded rectangle on the canvas."""

        self.create_arc(x1, y1, x1 + 2 * radius, y1 + 2 * radius, start=90, extent=90, fill=fill, outline="")
        self.create_arc(x2 - 2 * radius, y1, x2, y1 + 2 * radius, start=0, extent=90, fill=fill, outline="")
        self.create_arc(x2 - 2 * radius, y2 - 2 * radius, x2, y2, start=270, extent=90, fill=fill, outline="")
        self.create_arc(x1, y2 - 2 * radius, x1 + 2 * radius, y2, start=180, extent=90, fill=fill, outline="")
        self.create_rectangle(x1 + radius, y1, x2 - radius, y2, fill=fill, outline="")
        self.create_rectangle(x1, y1 + radius, x2, y2 - radius, fill=fill, outline="")


class TimeTrackerApp(tk.Tk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Учет рабочего времени")
        self.geometry("520x320")
        self.minsize(480, 300)
        self.resizable(True, True)

        self.style = ttk.Style(self)
        self._palette: dict[str, str] = {}
        self.button_diameter = 48

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
        chrome = ttk.Frame(self, style="Mac.Chrome.TFrame")
        chrome.pack(fill=tk.X, side=tk.TOP, padx=26, pady=(18, 8))

        traffic = ttk.Frame(chrome, style="Mac.Chrome.TFrame")
        traffic.pack(side=tk.LEFT, padx=(0, 12))
        for color in ("#ff5f57", "#febb2e", "#28c840"):
            dot = tk.Canvas(
                traffic,
                width=14,
                height=14,
                highlightthickness=0,
                bd=0,
                background=self._palette["mac_bg"],
            )
            dot.create_oval(2, 2, 12, 12, fill=color, outline=color)
            dot.pack(side=tk.LEFT, padx=4)

        ttk.Label(
            chrome,
            text="Timesheet Timer",
            style="Mac.WindowTitle.TLabel",
        ).pack(side=tk.LEFT)

        background = ttk.Frame(self, style="Mac.Background.TFrame")
        background.pack(fill=tk.BOTH, expand=True)

        panel = ttk.Frame(background, style="Mac.Panel.TFrame", padding=(0, 12, 0, 20))
        panel.pack(fill=tk.BOTH, expand=True, padx=40, pady=(0, 32))
        panel.columnconfigure(0, weight=1)
        self._content_panel = panel

        ttk.Label(panel, text="Учет рабочего времени", style="Mac.Header.TLabel").grid(
            row=0, column=0, sticky="w", padx=32, pady=(28, 6)
        )
        ttk.Separator(panel, orient=tk.HORIZONTAL, style="Mac.Separator.TSeparator").grid(
            row=1, column=0, sticky="ew", padx=32, pady=(0, 18)
        )

        self.project_var.set("Выберите проект")
        ttk.Label(panel, text="Проект", style="Mac.FieldLabel.TLabel").grid(
            row=2, column=0, sticky="w", padx=32, pady=(4, 4)
        )
        self.project_menu = ttk.OptionMenu(
            panel,
            self.project_var,
            self.project_var.get(),
        )
        self.project_menu.configure(style="Mac.OptionMenu.TMenubutton", width=24)
        self.project_menu.grid(row=3, column=0, sticky="ew", padx=32, pady=(0, 12))
        self._initialise_menu(self.project_menu)

        self.work_type_var.set("Выберите вид работы")
        ttk.Label(panel, text="Вид работы", style="Mac.FieldLabel.TLabel").grid(
            row=4, column=0, sticky="w", padx=32, pady=(0, 4)
        )
        self.work_menu = ttk.OptionMenu(
            panel,
            self.work_type_var,
            self.work_type_var.get(),
        )
        self.work_menu.configure(style="Mac.OptionMenu.TMenubutton", width=24)
        self.work_menu.grid(row=5, column=0, sticky="ew", padx=32, pady=(0, 18))
        self._initialise_menu(self.work_menu)

        panel.rowconfigure(6, weight=1)

        self.timer_label = ttk.Label(
            panel,
            textvariable=self.timer_var,
            anchor=tk.CENTER,
            style="Mac.Timer.TLabel",
        )
        self.timer_label.grid(row=6, column=0, sticky="nsew", padx=32, pady=(6, 18))

        buttons_frame = ttk.Frame(panel, style="Mac.Section.TFrame")
        buttons_frame.grid(row=7, column=0, pady=(0, 28))

        self.play_button = MacIconButton(
            buttons_frame,
            command=self.start_timer,
            icon="play",
            palette=self._palette,
            diameter=self.button_diameter,
        )
        self.play_button.grid(row=0, column=0, padx=10)

        self.pause_button = MacIconButton(
            buttons_frame,
            command=self.pause_timer,
            icon="pause",
            palette=self._palette,
            diameter=self.button_diameter,
        )
        self.pause_button.grid(row=0, column=1, padx=10)

        self.stop_button = MacIconButton(
            buttons_frame,
            command=self.stop_timer,
            icon="stop",
            palette=self._palette,
            diameter=self.button_diameter,
        )
        self.stop_button.grid(row=0, column=2, padx=10)

        for column_index in range(3):
            buttons_frame.grid_columnconfigure(column_index, weight=1)

        status_frame = ttk.Frame(self, style="Mac.Status.TFrame")
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            style="Mac.Status.TLabel",
            anchor=tk.W,
        )
        self.status_label.pack(side=tk.LEFT, padx=18, pady=6)
        self.bind("<Configure>", self._update_status_wrap)
        self.after_idle(self._update_status_wrap)

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
        self.timer_font = tkfont.Font(
            family=display_family,
            size=max(28, int(self.button_diameter * 0.9)),
            weight="bold",
        )

        self.option_add("*Font", self.base_font)
        self.option_add("*Menu.font", self.base_font)

    def _configure_styles(self) -> None:
        self.style.theme_use("clam")

        palette = {
            "mac_bg": "#e8e8ed",
            "panel_bg": "#f5f5f7",
            "accent": "#0a84ff",
            "text": "#1c1c1e",
            "muted": "#636366",
            "border": "#c7c7cc",
            "status_bg": "#d1d1d6",
            "field_bg": "#ffffff",
            "field_border": "#d0d0d5",
            "field_focus": "#5e5ce6",
            "toolbar_base": "#ffffff",
            "toolbar_hover": "#f7f7f9",
            "toolbar_active": "#e2e2e8",
            "toolbar_outline": "#c7c7cc",
            "toolbar_highlight": "#ffffff",
        }
        self._palette = palette

        self.configure(background=palette["mac_bg"])

        self.style.configure("Mac.Background.TFrame", background=palette["mac_bg"])
        self.style.configure("Mac.Chrome.TFrame", background=palette["mac_bg"])
        self.style.configure(
            "Mac.Panel.TFrame",
            background=palette["panel_bg"],
            borderwidth=1,
            relief="solid",
            bordercolor=palette["border"],
        )
        self.style.configure("Mac.Section.TFrame", background=palette["panel_bg"])
        self.style.configure("Mac.TLabel", background=palette["panel_bg"], foreground=palette["text"], font=self.base_font)
        self.style.configure("Mac.FieldLabel.TLabel", background=palette["panel_bg"], foreground=palette["muted"], font=self.base_font)
        self.style.configure(
            "Mac.WindowTitle.TLabel",
            background=palette["mac_bg"],
            foreground=palette["muted"],
            font=self.base_font,
        )
        self.style.configure(
            "Mac.Header.TLabel",
            background=palette["panel_bg"],
            foreground=palette["text"],
            font=self.base_font,
        )
        self.style.configure("Mac.Timer.TLabel", background=palette["panel_bg"], foreground=palette["accent"], font=self.timer_font)
        self.style.layout("Mac.Separator.TSeparator", [("Separator.separator", {"sticky": "we"})])
        self.style.configure("Mac.Separator.TSeparator", background=palette["border"])

        option_style = "Mac.OptionMenu.TMenubutton"
        self.style.layout(
            option_style,
            [
                (
                    "Menubutton.padding",
                    {
                        "sticky": "nswe",
                        "children": [
                            (
                                "Menubutton.background",
                                {
                                    "sticky": "nswe",
                                    "children": [
                                        ("Menubutton.label", {"sticky": "w"}),
                                        ("Menubutton.indicator", {"side": "right", "sticky": ""}),
                                    ],
                                },
                            ),
                        ],
                    },
                ),
            ],
        )
        self.style.configure(
            option_style,
            background=palette["field_bg"],
            foreground=palette["text"],
            font=self.base_font,
            borderwidth=1,
            bordercolor=palette["field_border"],
            relief="flat",
            padding=(18, 10, 30, 10),
            arrowcolor=palette["muted"],
            arrowsize=12,
        )
        self.style.map(
            option_style,
            background=[("active", "#f2f2f7"), ("pressed", "#e9e9ef")],
            foreground=[("disabled", palette["muted"])],
            arrowcolor=[("active", palette["text"])],
            bordercolor=[("focus", palette["field_focus"]), ("active", palette["field_border"])],
        )

        self.style.configure(
            "Mac.Status.TFrame",
            background=palette["status_bg"],
            borderwidth=1,
            relief="flat",
        )
        self.style.configure(
            "Mac.Status.TLabel",
            background=palette["status_bg"],
            foreground=palette["text"],
            font=self.base_font,
        )

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

    def _update_status_wrap(self, _event: Optional[tk.Event] = None) -> None:
        available_width = max(120, self.winfo_width() - 96)
        self.status_label.configure(wraplength=available_width)

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
        desired_width = max(640, min(1200, max_width + 360))

        char_width = max(font_obj.measure("0"), 1)
        menu_chars = min(52, max(20, (max_width + 80) // char_width))
        self.project_menu.configure(width=menu_chars)
        self.work_menu.configure(width=menu_chars)

        self.update_idletasks()
        current_height = max(self.winfo_height(), 400)
        self.geometry(f"{desired_width}x{current_height}")
        self.minsize(desired_width, 380)

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
