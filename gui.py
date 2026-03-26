from __future__ import annotations

import contextlib
import io
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from matplotlib import colormaps

from config import AppConfig, DEFAULT_COLORMAP, DEFAULT_OUTPUT_FORMAT, DEFAULT_PLOT_BACKEND, DEFAULT_THEME
from data_loader import DataLoaderError
from main import run, sync_output_settings
from plot_catalog import DEFAULT_PLOT_TYPE, available_plot_type_labels
from plotter import PlottingError
from utils import (
    UserInputError,
    build_output_path,
    parse_axis_limits,
    parse_cycle_expression,
    parse_figure_size,
    parse_float_text,
    parse_font_family,
    parse_int_text,
    parse_mode_overrides,
    parse_optional_color,
    parse_output_format,
    validate_input_file,
)

COMMON_COLORMAPS = (
    "tab10",
    "tab20",
    "Set1",
    "Set2",
    "Set3",
    "viridis",
    "plasma",
    "inferno",
    "magma",
    "cividis",
    "turbo",
    "coolwarm",
    "Spectral",
    "RdYlBu",
    "rainbow",
)


class TextRedirector(io.TextIOBase):
    def __init__(self, widget: tk.Text, root: tk.Tk) -> None:
        self.widget = widget
        self.root = root

    def write(self, text: str) -> int:
        if not text:
            return 0

        def append() -> None:
            self.widget.configure(state="normal")
            self.widget.insert("end", text)
            self.widget.see("end")
            self.widget.configure(state="disabled")

        self.root.after(0, append)
        return len(text)

    def flush(self) -> None:
        return None


class BatteryPlotterGUI:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("电池充放电曲线绘图工具 GUI")
        self.root.geometry("980x760")
        self.root.minsize(900, 700)

        self.input_path_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar()
        self.cycles_var = tk.StringVar()
        self.output_dir_var = tk.StringVar(value=str(Path.cwd()))
        self.base_filename_var = tk.StringVar(value="battery_curve")
        self.backend_var = tk.StringVar(value=DEFAULT_PLOT_BACKEND)
        self.output_format_var = tk.StringVar(value=DEFAULT_OUTPUT_FORMAT)
        self.title_var = tk.StringVar(value="Battery Charge-Discharge Curves")
        self.x_label_var = tk.StringVar(value="Specific Capacity (mAh/g)")
        self.y_label_var = tk.StringVar(value="Voltage (V)")
        self.dpi_var = tk.StringVar(value="300")
        self.theme_var = tk.StringVar(value=DEFAULT_THEME)
        self.colormap_var = tk.StringVar(value=DEFAULT_COLORMAP)
        self.legend_loc_var = tk.StringVar(value="best")
        self.line_width_var = tk.StringVar(value="2.2")
        self.figure_size_var = tk.StringVar(value="8,6")
        self.x_lim_var = tk.StringVar()
        self.y_lim_var = tk.StringVar()
        self.font_family_var = tk.StringVar(value="Microsoft YaHei,SimHei,DejaVu Sans")
        self.mode_map_var = tk.StringVar()
        self.line_color_var = tk.StringVar(value="#1f77b4")
        self.plot_type_vars = {
            key: tk.BooleanVar(value=(key == DEFAULT_PLOT_TYPE)) for key, _ in available_plot_type_labels()
        }

        self.show_legend_var = tk.BooleanVar(value=True)
        self.grid_var = tk.BooleanVar(value=False)
        self.color_by_cycle_var = tk.BooleanVar(value=True)
        self.auto_sort_var = tk.BooleanVar(value=True)
        self.absolute_specific_capacity_var = tk.BooleanVar(value=True)
        self.show_after_save_var = tk.BooleanVar(value=False)
        self.transparent_background_var = tk.BooleanVar(value=False)
        self.use_demo_var = tk.BooleanVar(value=False)
        self.save_origin_project_var = tk.BooleanVar(value=True)

        self.start_button: ttk.Button | None = None
        self.log_text: tk.Text | None = None
        self.colormap_preview_canvas: tk.Canvas | None = None

        self._build_layout()
        self.colormap_var.trace_add("write", self._on_colormap_change)

    def _build_layout(self) -> None:
        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        note = (
            "说明：如果 Excel 中包含多个子表格，程序会自动依次绘图；"
            "若信息表中含原始子表路径/文件名，会优先按原子表名字命名输出。"
        )
        ttk.Label(main_frame, text=note, foreground="#1f4e79", wraplength=920).grid(
            row=0, column=0, sticky="ew", pady=(0, 10)
        )

        content_pane = ttk.Panedwindow(main_frame, orient="vertical")
        content_pane.grid(row=1, column=0, sticky="nsew")

        form_container = ttk.Frame(content_pane, padding=(0, 0, 0, 10))
        form_container.columnconfigure(0, weight=1)
        form_container.rowconfigure(0, weight=1)
        content_pane.add(form_container, weight=3)

        notebook = ttk.Notebook(form_container)
        notebook.grid(row=0, column=0, sticky="nsew")

        basic_tab = self._create_scrollable_tab(notebook, "基础设置")
        advanced_tab = self._create_scrollable_tab(notebook, "高级设置")

        self._build_basic_tab(basic_tab)
        self._build_advanced_tab(advanced_tab)

        log_frame = ttk.LabelFrame(content_pane, text="运行日志", padding=8)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        content_pane.add(log_frame, weight=2)

        self.log_text = tk.Text(log_frame, height=16, wrap="word", state="disabled")
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scrollbar.grid(row=0, column=1, sticky="ns")

        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))

        self.start_button = ttk.Button(action_frame, text="开始绘图", command=self.start_run)
        self.start_button.pack(side="left")
        ttk.Button(action_frame, text="清空日志", command=self.clear_log).pack(side="left", padx=(8, 0))
        ttk.Button(action_frame, text="恢复默认", command=self.reset_defaults).pack(side="left", padx=(8, 0))
        ttk.Button(action_frame, text="退出", command=self.root.destroy).pack(side="right")

    def _create_scrollable_tab(self, notebook: ttk.Notebook, title: str) -> ttk.Frame:
        outer = ttk.Frame(notebook)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        inner = ttk.Frame(canvas, padding=12)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def on_inner_configure(_: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(_: tk.Event) -> None:
            canvas.itemconfigure(window_id, width=canvas.winfo_width())

        def on_mousewheel(event: tk.Event) -> None:
            if canvas.winfo_height() < inner.winfo_reqheight():
                canvas.yview_scroll(int(-event.delta / 120), "units")

        inner.bind("<Configure>", on_inner_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        canvas.bind("<Enter>", lambda _event: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda _event: canvas.unbind_all("<MouseWheel>"))

        notebook.add(outer, text=title)
        return inner

    def _build_basic_tab(self, frame: ttk.Frame) -> None:
        for col in range(3):
            frame.columnconfigure(col, weight=1 if col == 1 else 0)

        row = 0
        self._add_entry_row(frame, row, "输入 Excel：", self.input_path_var, self.choose_input_file)
        row += 1
        self._add_entry_row(frame, row, "Sheet 名：", self.sheet_name_var)
        row += 1
        self._add_entry_row(frame, row, "循环编号：", self.cycles_var)
        row += 1
        self._add_entry_row(frame, row, "输出目录：", self.output_dir_var, self.choose_output_dir, directory_mode=True)
        row += 1
        self._add_entry_row(frame, row, "基础文件名：", self.base_filename_var)
        row += 1

        ttk.Label(frame, text="输出格式：").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Combobox(
            frame,
            textvariable=self.output_format_var,
            values=("png", "svg", "pdf"),
            state="readonly",
            width=12,
        ).grid(row=row, column=1, sticky="w", pady=6)
        row += 1

        ttk.Label(frame, text="绘图后端：").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Combobox(
            frame,
            textvariable=self.backend_var,
            values=("matplotlib", "origin"),
            state="readonly",
            width=12,
        ).grid(row=row, column=1, sticky="w", pady=6)
        row += 1

        plot_type_frame = ttk.LabelFrame(frame, text="Plot Types", padding=10)
        plot_type_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        for col in range(2):
            plot_type_frame.columnconfigure(col, weight=1)
        for index, (plot_type, label) in enumerate(available_plot_type_labels()):
            ttk.Checkbutton(
                plot_type_frame,
                text=label,
                variable=self.plot_type_vars[plot_type],
            ).grid(row=index // 2, column=index % 2, sticky="w", padx=(0, 10), pady=2)
        row += 1

        self._add_entry_row(frame, row, "图标题：", self.title_var)
        row += 1
        self._add_entry_row(frame, row, "X 轴标题：", self.x_label_var)
        row += 1
        self._add_entry_row(frame, row, "Y 轴标题：", self.y_label_var)
        row += 1

        option_frame = ttk.LabelFrame(frame, text="常用选项", padding=10)
        option_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        for col in range(3):
            option_frame.columnconfigure(col, weight=1)

        ttk.Checkbutton(option_frame, text="显示图例", variable=self.show_legend_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(option_frame, text="显示网格", variable=self.grid_var).grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(option_frame, text="按循环着色", variable=self.color_by_cycle_var).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(option_frame, text="自动排序数据", variable=self.auto_sort_var).grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Checkbutton(
            option_frame,
            text="比容量取绝对值",
            variable=self.absolute_specific_capacity_var,
        ).grid(row=1, column=1, sticky="w", pady=(8, 0))
        ttk.Checkbutton(option_frame, text="保存后显示图像", variable=self.show_after_save_var).grid(
            row=1, column=2, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(
            option_frame,
            text="透明背景",
            variable=self.transparent_background_var,
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Checkbutton(option_frame, text="使用内置 Demo", variable=self.use_demo_var).grid(
            row=2, column=1, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(
            option_frame,
            text="保存 Origin 工程",
            variable=self.save_origin_project_var,
        ).grid(row=2, column=2, sticky="w", pady=(8, 0))

    def _build_advanced_tab(self, frame: ttk.Frame) -> None:
        for col in range(3):
            frame.columnconfigure(col, weight=1 if col == 1 else 0)

        row = 0
        self._add_entry_row(frame, row, "DPI：", self.dpi_var)
        row += 1
        self._add_entry_row(frame, row, "图例位置：", self.legend_loc_var)
        row += 1

        ttk.Label(frame, text="主题：").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Combobox(
            frame,
            textvariable=self.theme_var,
            values=("paper", "default"),
            state="readonly",
            width=12,
        ).grid(row=row, column=1, sticky="w", pady=6)
        row += 1

        self._add_colormap_row(frame, row)
        row += 1
        self._add_entry_row(frame, row, "统一颜色：", self.line_color_var)
        row += 1
        self._add_entry_row(frame, row, "线宽：", self.line_width_var)
        row += 1
        self._add_entry_row(frame, row, "图尺寸：", self.figure_size_var)
        row += 1
        self._add_entry_row(frame, row, "X 轴范围：", self.x_lim_var)
        row += 1
        self._add_entry_row(frame, row, "Y 轴范围：", self.y_lim_var)
        row += 1
        self._add_entry_row(frame, row, "字体列表：", self.font_family_var)
        row += 1
        self._add_entry_row(frame, row, "模式映射：", self.mode_map_var)
        row += 1

        help_text = (
            "填写示例：\n"
            "- 循环编号：1,2,5-8\n"
            "- 图尺寸：8,6\n"
            "- 坐标范围：0,200\n"
            "- 字体列表：Microsoft YaHei,SimHei,DejaVu Sans\n"
            "- 模式映射：恒流充电=charge,恒流放电=discharge"
        )
        ttk.Label(frame, text=help_text, foreground="#555555", justify="left").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(12, 0)
        )

    def _add_colormap_row(self, frame: ttk.Frame, row: int) -> None:
        ttk.Label(frame, text="Colormap：").grid(row=row, column=0, sticky="nw", pady=6)

        inner = ttk.Frame(frame)
        inner.grid(row=row, column=1, columnspan=2, sticky="ew", pady=6)
        inner.columnconfigure(0, weight=1)

        combobox = ttk.Combobox(
            inner,
            textvariable=self.colormap_var,
            values=COMMON_COLORMAPS,
            state="normal",
        )
        combobox.grid(row=0, column=0, sticky="ew")
        combobox.bind("<<ComboboxSelected>>", self._on_colormap_selected)

        self.colormap_preview_canvas = tk.Canvas(
            inner,
            height=24,
            highlightthickness=1,
            highlightbackground="#BFBFBF",
            bg="white",
        )
        self.colormap_preview_canvas.grid(row=1, column=0, sticky="ew", pady=(6, 2))
        self.colormap_preview_canvas.bind("<Configure>", lambda _event: self._update_colormap_preview())

        ttk.Label(
            inner,
            text="上方可直接选择常用色卡，也可手动输入 matplotlib 支持的 colormap 名称。",
            foreground="#666666",
        ).grid(row=2, column=0, sticky="w")

        self._update_colormap_preview()

    def _add_entry_row(
        self,
        frame: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        button_command: Callable[[], None] | None = None,
        directory_mode: bool = False,
    ) -> None:
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=6, padx=(8, 8))
        if button_command is not None:
            ttk.Button(frame, text="选择...", command=button_command).grid(row=row, column=2, sticky="ew", pady=6)
        elif directory_mode:
            ttk.Label(frame, text="").grid(row=row, column=2, sticky="ew", pady=6)

    def _on_colormap_selected(self, _event: tk.Event) -> None:
        self._update_colormap_preview()

    def _on_colormap_change(self, *_args: object) -> None:
        self._update_colormap_preview()

    def _update_colormap_preview(self) -> None:
        canvas = self.colormap_preview_canvas
        if canvas is None:
            return

        canvas.delete("all")
        width = max(canvas.winfo_width(), 240)
        height = max(canvas.winfo_height(), 24)
        name = self.colormap_var.get().strip()
        if not name:
            canvas.create_text(width / 2, height / 2, text="未设置 colormap", fill="#888888")
            return

        try:
            cmap = colormaps[name]
        except Exception:
            canvas.create_rectangle(0, 0, width, height, fill="#F8D7DA", outline="")
            canvas.create_text(width / 2, height / 2, text=f"未找到 colormap: {name}", fill="#A94442")
            return

        sample_count = max(min(width // 18, 24), 12)
        segment_width = width / sample_count
        for index in range(sample_count):
            rgba = cmap(index / max(sample_count - 1, 1))
            red = int(rgba[0] * 255)
            green = int(rgba[1] * 255)
            blue = int(rgba[2] * 255)
            color = f"#{red:02x}{green:02x}{blue:02x}"
            x0 = index * segment_width
            x1 = (index + 1) * segment_width
            canvas.create_rectangle(x0, 0, x1, height, fill=color, outline=color)

    def choose_input_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not selected:
            return
        self.input_path_var.set(selected)
        input_path = Path(selected)
        current_output_dir = self.output_dir_var.get().strip()
        if not current_output_dir or Path(current_output_dir) == Path.cwd():
            self.output_dir_var.set(str(input_path.parent))
        if self.base_filename_var.get().strip() == "battery_curve":
            self.base_filename_var.set(input_path.stem)

    def choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择输出目录")
        if selected:
            self.output_dir_var.set(selected)

    def append_log(self, message: str) -> None:
        if self.log_text is None:
            return
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_log(self) -> None:
        if self.log_text is None:
            return
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def reset_defaults(self) -> None:
        self.sheet_name_var.set("")
        self.cycles_var.set("")
        self.output_format_var.set(DEFAULT_OUTPUT_FORMAT)
        self.backend_var.set(DEFAULT_PLOT_BACKEND)
        for key, var in self.plot_type_vars.items():
            var.set(key == DEFAULT_PLOT_TYPE)
        self.title_var.set("Battery Charge-Discharge Curves")
        self.x_label_var.set("Specific Capacity (mAh/g)")
        self.y_label_var.set("Voltage (V)")
        self.dpi_var.set("300")
        self.theme_var.set(DEFAULT_THEME)
        self.colormap_var.set(DEFAULT_COLORMAP)
        self.legend_loc_var.set("best")
        self.line_width_var.set("2.2")
        self.figure_size_var.set("8,6")
        self.x_lim_var.set("")
        self.y_lim_var.set("")
        self.font_family_var.set("Microsoft YaHei,SimHei,DejaVu Sans")
        self.mode_map_var.set("")
        self.line_color_var.set("#1f77b4")
        self.show_legend_var.set(True)
        self.grid_var.set(False)
        self.color_by_cycle_var.set(True)
        self.auto_sort_var.set(True)
        self.absolute_specific_capacity_var.set(True)
        self.show_after_save_var.set(False)
        self.transparent_background_var.set(False)
        self.use_demo_var.set(False)
        self.save_origin_project_var.set(True)

    def build_config(self) -> tuple[AppConfig, bool]:
        config = AppConfig()
        use_demo = self.use_demo_var.get()

        if not use_demo:
            input_text = self.input_path_var.get().strip()
            if not input_text:
                raise UserInputError("请选择输入 Excel 文件。")
            config.input_path = validate_input_file(input_text)

        sheet_text = self.sheet_name_var.get().strip()
        config.sheet_name = sheet_text or None

        output_dir_text = self.output_dir_var.get().strip()
        if not output_dir_text:
            raise UserInputError("请选择输出目录。")

        output_format = parse_output_format(self.output_format_var.get(), DEFAULT_OUTPUT_FORMAT)
        base_filename = self.base_filename_var.get().strip() or "battery_curve"
        output_target = Path(output_dir_text) / f"{base_filename}.{output_format}"
        config.output_path = build_output_path(str(output_target), output_format)
        config.output_format = output_format
        config.plot_backend = (self.backend_var.get().strip() or DEFAULT_PLOT_BACKEND).lower()
        config.plot_types = [key for key, var in self.plot_type_vars.items() if var.get()]
        if not config.plot_types:
            raise UserInputError("请至少选择一种绘图类型。")

        cycles_text = self.cycles_var.get().strip()
        config.cycles = parse_cycle_expression(cycles_text) if cycles_text else None
        config.title = self.title_var.get().strip() or config.title
        config.x_label = self.x_label_var.get().strip() or config.x_label
        config.y_label = self.y_label_var.get().strip() or config.y_label
        config.dpi = parse_int_text(self.dpi_var.get().strip(), config.dpi, minimum=50)
        config.legend_loc = self.legend_loc_var.get().strip() or config.legend_loc
        config.theme = self.theme_var.get().strip() or config.theme
        config.colormap = self.colormap_var.get().strip() or config.colormap
        if config.colormap:
            try:
                colormaps[config.colormap]
            except Exception as exc:
                raise UserInputError(f"无效的 colormap 名称: {config.colormap}") from exc
        config.line_width = parse_float_text(self.line_width_var.get().strip(), config.line_width, minimum=0.1)
        config.line_color = parse_optional_color(self.line_color_var.get().strip(), config.line_color)
        width, height = parse_figure_size(self.figure_size_var.get().strip(), config.figure_size())
        config.figure_width = width
        config.figure_height = height
        config.x_lim = parse_axis_limits(self.x_lim_var.get().strip()) if self.x_lim_var.get().strip() else None
        config.y_lim = parse_axis_limits(self.y_lim_var.get().strip()) if self.y_lim_var.get().strip() else None
        config.font_family = parse_font_family(self.font_family_var.get().strip())
        config.mode_overrides = parse_mode_overrides(self.mode_map_var.get().strip())
        config.show_legend = self.show_legend_var.get()
        config.grid = self.grid_var.get()
        config.color_by_cycle = self.color_by_cycle_var.get()
        config.auto_sort = self.auto_sort_var.get()
        config.absolute_specific_capacity = self.absolute_specific_capacity_var.get()
        config.show_after_save = self.show_after_save_var.get()
        config.transparent_background = self.transparent_background_var.get()
        config.save_origin_project = self.save_origin_project_var.get()
        sync_output_settings(config)
        return config, use_demo

    def start_run(self) -> None:
        try:
            config, use_demo = self.build_config()
        except (UserInputError, ValueError) as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        if self.start_button is not None:
            self.start_button.configure(state="disabled")
        self.append_log("\n=== 开始执行 ===\n")
        if use_demo:
            self.append_log("输入数据：内置 Demo\n")
        else:
            self.append_log(f"输入文件：{config.input_path}\n")
        self.append_log(f"输出路径：{config.output_path}\n")
        self.append_log(f"绘图后端：{config.plot_backend}\n")

        worker = threading.Thread(
            target=self._run_worker,
            args=(config, use_demo),
            daemon=True,
        )
        worker.start()

    def _run_worker(self, config: AppConfig, use_demo: bool) -> None:
        redirector = TextRedirector(self.log_text, self.root) if self.log_text is not None else io.StringIO()
        try:
            with contextlib.redirect_stdout(redirector), contextlib.redirect_stderr(redirector):
                saved_paths = run(config, use_demo=use_demo)
                if len(saved_paths) == 1:
                    print(f"图片已保存到 {saved_paths[0]}")
                else:
                    print("图片已全部保存到：")
                    for saved_path in saved_paths:
                        print(f"- {saved_path}")

            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "完成",
                    f"绘图完成，共生成 {len(saved_paths)} 个文件。",
                ),
            )
        except (UserInputError, DataLoaderError, PlottingError, ValueError) as exc:
            self.root.after(0, lambda: messagebox.showerror("运行失败", str(exc)))
            self.root.after(0, lambda: self.append_log(f"错误: {exc}\n"))
        finally:
            self.root.after(0, self._finish_run)

    def _finish_run(self) -> None:
        self.append_log("=== 执行结束 ===\n")
        if self.start_button is not None:
            self.start_button.configure(state="normal")

    def run_forever(self) -> None:
        self.root.mainloop()


def launch_gui() -> None:
    app = BatteryPlotterGUI()
    app.run_forever()


if __name__ == "__main__":
    launch_gui()
