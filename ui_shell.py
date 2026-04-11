from __future__ import annotations

import json
import os
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Callable

import pandas as pd
import sv_ttk

from common_utils import describe_filter_conditions, open_with_default_app
from feature_support import (
    ensure_advanced_state,
    merge_multiple_files_into_app,
    open_area_conversion_dialog,
    open_date_grouping_dialog,
    open_field_recommendations,
    open_filter_preset_manager,
    open_group_mapping_manager,
    open_operation_history,
    open_task_center,
    publish_output_bundle,
    save_workspace_as,
    load_workspace_from_file,
    undo_last_action,
)


PREFERENCES_FILE = "app_preferences.json"
DEFAULT_PREFERENCES = {
    "default_mode": "horizontal",
    "theme": "dark",
    "reopen_last_file": False,
    "output_location": "source_dir",
    "append_timestamp_to_export": False,
}
MODE_LABELS = {
    "horizontal": "横版",
    "vertical": "竖版",
}
MODE_NAMES = {label: value for value, label in MODE_LABELS.items()}
THEME_LABELS = {
    "dark": "深色",
    "light": "浅色",
}
THEME_NAMES = {label: value for value, label in THEME_LABELS.items()}
OUTPUT_LOCATION_LABELS = {
    "source_dir": "源文件同目录",
    "app_dir": "程序目录",
}
OUTPUT_LOCATION_NAMES = {label: value for value, label in OUTPUT_LOCATION_LABELS.items()}


def _preferences_path() -> Path:
    return Path(__file__).resolve().with_name(PREFERENCES_FILE)


def load_app_preferences() -> dict[str, object]:
    preferences = DEFAULT_PREFERENCES.copy()
    path = _preferences_path()
    if not path.exists():
        return preferences

    try:
        with path.open("r", encoding="utf-8") as file:
            loaded = json.load(file)
    except Exception:
        return preferences

    if isinstance(loaded, dict):
        for key in DEFAULT_PREFERENCES:
            if key in loaded:
                preferences[key] = loaded[key]
    return preferences


def save_app_preferences(preferences: dict[str, object]) -> None:
    payload = DEFAULT_PREFERENCES.copy()
    for key in DEFAULT_PREFERENCES:
        payload[key] = preferences.get(key, DEFAULT_PREFERENCES[key])

    with _preferences_path().open("w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def apply_app_theme(theme: str) -> str:
    normalized = theme if theme in THEME_LABELS else str(DEFAULT_PREFERENCES["theme"])
    try:
        sv_ttk.set_theme(normalized)
    except Exception:
        normalized = "light"
        try:
            sv_ttk.set_theme(normalized)
        except Exception:
            pass
    return normalized


def center_window(window: tk.Misc, width: int, height: int) -> None:
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = max((screen_width - width) // 2, 0)
    y = max((screen_height - height) // 2, 0)
    window.geometry(f"{width}x{height}+{x}+{y}")


def initialize_shell(app, *, mode_name: str, title: str) -> None:
    app.mode_name = mode_name
    app.mode_label = MODE_LABELS.get(mode_name, mode_name)
    app.app_preferences = load_app_preferences()
    app.last_output_path = None
    app.current_sheet_name = ""
    theme = str(app.app_preferences.get("theme", DEFAULT_PREFERENCES["theme"]))
    app.current_theme = apply_app_theme(theme)
    ensure_advanced_state(app)
    app.root.title(f"{title} · {app.mode_label}")


def set_status(app, message: str) -> None:
    if hasattr(app, "status_var"):
        app.status_var.set(message)


def begin_busy(app, message: str) -> None:
    set_status(app, message)
    if hasattr(app, "progress"):
        app.progress.pack(fill="x", padx=12, pady=(0, 4))
        app.progress.start()
    app.root.update_idletasks()


def end_busy(app, message: str | None = None) -> None:
    if hasattr(app, "progress"):
        app.progress.stop()
        app.progress.pack_forget()
    if message:
        set_status(app, message)


def clear_recent_files(app) -> None:
    app.recent_files = []
    app.save_config(show_msg=False)
    if hasattr(app, "create_recent_menu"):
        app.create_recent_menu()
    set_status(app, "最近打开已清空")


def populate_recent_menus(app) -> None:
    menus = list(getattr(app, "_recent_menus", []))
    if hasattr(app, "recent_menu") and app.recent_menu not in menus:
        menus.append(app.recent_menu)

    for menu in menus:
        menu.delete(0, "end")
        recent_files = getattr(app, "recent_files", [])
        for path in recent_files:
            menu.add_command(label=os.path.basename(path), command=lambda p=path: app.open_recent(p))
        if not recent_files:
            menu.add_command(label="无记录", state="disabled")


def maybe_restore_recent_file(app) -> None:
    if not bool(app.app_preferences.get("reopen_last_file", False)):
        return

    recent_files = getattr(app, "recent_files", [])
    if recent_files and os.path.exists(recent_files[0]):
        app.root.after(150, lambda: app.open_recent(recent_files[0]))


def resolve_output_directory(app) -> Path:
    output_location = str(app.app_preferences.get("output_location", DEFAULT_PREFERENCES["output_location"]))
    if output_location == "source_dir" and getattr(app, "excel_path", ""):
        source_dir = Path(str(app.excel_path)).expanduser().resolve().parent
        if source_dir.exists():
            return source_dir
    return Path(__file__).resolve().parent


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    index = 2
    while True:
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def build_output_path(app, basename: str) -> str:
    export_dir = resolve_output_directory(app)
    stem = basename
    if bool(app.app_preferences.get("append_timestamp_to_export", False)):
        stem = f"{stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return str(_unique_path(export_dir / f"{stem}.xlsx"))


def copy_text(app, text: str, status_message: str = "内容已复制到剪贴板") -> None:
    app.root.clipboard_clear()
    app.root.clipboard_append(text)
    set_status(app, status_message)


def open_output_folder(app) -> None:
    path = getattr(app, "last_output_path", None)
    if path and os.path.exists(path):
        directory = os.path.dirname(path)
    else:
        directory = str(resolve_output_directory(app))
    open_with_default_app(directory)


def preview_last_output(app) -> None:
    path = getattr(app, "last_output_path", None)
    if not path or not os.path.exists(path):
        messagebox.showinfo("提示", "还没有可预览的导出结果。", parent=app.root)
        return
    app.open_preview(path)


def open_last_output(app) -> None:
    path = getattr(app, "last_output_path", None)
    if not path or not os.path.exists(path):
        messagebox.showinfo("提示", "还没有可打开的导出结果。", parent=app.root)
        return
    app.open_file(path)


def mark_output(app, filepath: str) -> None:
    app.last_output_path = os.path.abspath(filepath)
    set_status(app, f"导出完成：{os.path.basename(filepath)}")


def _loaded_dataframe(app) -> pd.DataFrame | None:
    if hasattr(app, "get_active_data"):
        try:
            active = app.get_active_data()
        except Exception:
            active = None
        if isinstance(active, pd.DataFrame) and len(active.columns) > 0:
            return active

    data = getattr(app, "data", None)
    if isinstance(data, pd.DataFrame) and len(data.columns) > 0:
        return data
    return None


def _sample_values(series: pd.Series, limit: int = 3) -> str:
    values = []
    for value in series.dropna().astype(str).unique()[:limit]:
        values.append(value if len(value) <= 18 else f"{value[:15]}...")
    return " / ".join(values) if values else "-"


def build_data_report_text(app) -> str:
    df = _loaded_dataframe(app)
    if df is None:
        return "当前还没有已加载的数据。"

    rows, cols = df.shape
    raw_data = getattr(app, "data", None)
    raw_rows = int(len(raw_data)) if isinstance(raw_data, pd.DataFrame) else rows
    file_path = getattr(app, "excel_path", "") or "未加载文件"
    sheet_name = getattr(app, "current_sheet_name", "") or "当前数据"
    numeric_count = len(df.select_dtypes(include="number").columns)
    duplicate_rows = int(df.duplicated().sum()) if rows else 0
    empty_columns = int((df.count() == 0).sum())
    missing_counts = df.isna().sum().sort_values(ascending=False)
    missing_counts = missing_counts[missing_counts > 0]

    lines = [
        f"文件：{file_path}",
        f"工作表/数据源：{sheet_name}",
        f"记录数：{rows}",
        f"字段数：{cols}",
        f"数值字段：{numeric_count}",
        f"非数值字段：{cols - numeric_count}",
        f"重复行：{duplicate_rows}",
        f"全空字段：{empty_columns}",
    ]

    active_filters = getattr(app, "filter_conditions", [])
    if active_filters:
        lines.append(f"原始记录数：{raw_rows}")
        lines.append(f"筛选条件：{describe_filter_conditions(active_filters)}")

    active_template_name = getattr(app, "active_group_template_name", "")
    if active_template_name:
        lines.append(f"当前模板：{active_template_name}")

    if not missing_counts.empty:
        lines.append("")
        lines.append("缺失值最多的字段：")
        for column, count in missing_counts.head(8).items():
            percent = (count / rows * 100) if rows else 0
            lines.append(f"- {column}: {int(count)} ({percent:.1f}%)")

    numeric_targets: list[tuple[str, str]] = []
    categorical_targets: list[tuple[str, str]] = []

    if getattr(app, "val_var", None) is not None and app.val_var.get():
        numeric_targets.append(("值字段", app.val_var.get()))
    if getattr(app, "area_cb", None) is not None and bool(app.area_cb.get()):
        ratio_var = getattr(app, "ratio_var", None)
        if ratio_var is not None and ratio_var.get():
            numeric_targets.append(("面积字段", ratio_var.get()))
    for field in getattr(app, "batch_fields", []):
        numeric_targets.append(("批量值字段", field))
    for field in getattr(app, "group_batch_fields", []):
        categorical_targets.append(("批量分组字段", field))
    for level in getattr(app, "levels", []):
        field = level["var"].get()
        if field:
            categorical_targets.append(("分类级别", field))

    seen: set[tuple[str, str]] = set()
    filtered_numeric_targets = []
    for item in numeric_targets:
        if item not in seen and item[1] in df.columns:
            filtered_numeric_targets.append(item)
            seen.add(item)

    seen.clear()
    filtered_categorical_targets = []
    for item in categorical_targets:
        if item not in seen and item[1] in df.columns:
            filtered_categorical_targets.append(item)
            seen.add(item)

    if filtered_numeric_targets or filtered_categorical_targets:
        lines.append("")
        lines.append("当前选择检查：")

    for label, field in filtered_numeric_targets:
        series = df[field]
        non_null = int(series.notna().sum())
        numeric_valid = int(pd.to_numeric(series, errors="coerce").notna().sum())
        invalid = max(non_null - numeric_valid, 0)
        lines.append(f"- {label} {field}: 非空 {non_null}，可转数值 {numeric_valid}，异常 {invalid}")

    for label, field in filtered_categorical_targets:
        series = df[field]
        unique_count = int(series.nunique(dropna=True))
        missing = int(series.isna().sum())
        lines.append(f"- {label} {field}: 唯一值 {unique_count}，缺失 {missing}")

    return "\n".join(lines)


def open_data_inspector(app, initial_tab: str = "preview") -> None:
    df = _loaded_dataframe(app)
    if df is None:
        messagebox.showinfo("提示", "请先加载数据后再查看。", parent=app.root)
        return

    dialog = tk.Toplevel(app.root)
    dialog.title("数据检查器")
    dialog.geometry("1120x680")
    dialog.minsize(920, 560)
    dialog.transient(app.root)

    container = ttk.Frame(dialog, padding=12)
    container.pack(fill="both", expand=True)

    rows, cols = df.shape
    source_label = getattr(app, "current_sheet_name", "") or "当前数据"
    source_file = os.path.basename(getattr(app, "excel_path", "") or "未命名数据")
    ttk.Label(container, text=f"{source_file} | {source_label} | {rows} 行 × {cols} 列").pack(anchor="w", pady=(0, 8))

    notebook = ttk.Notebook(container)
    notebook.pack(fill="both", expand=True)

    preview_tab = ttk.Frame(notebook)
    fields_tab = ttk.Frame(notebook)
    report_tab = ttk.Frame(notebook)
    notebook.add(preview_tab, text="预览")
    notebook.add(fields_tab, text="字段概览")
    notebook.add(report_tab, text="检查报告")

    preview_controls = ttk.Frame(preview_tab)
    preview_controls.pack(fill="x", padx=8, pady=(8, 4))
    limit_var = tk.StringVar(value="300")
    mode_var = tk.StringVar(value="前N行")
    search_var = tk.StringVar()
    info_var = tk.StringVar(value="")

    ttk.Label(preview_controls, text="预览方式").pack(side="left")
    ttk.Combobox(preview_controls, textvariable=mode_var, values=["前N行", "随机抽样", "筛选命中"], state="readonly", width=10).pack(side="left", padx=(6, 10))
    ttk.Label(preview_controls, text="预览行数").pack(side="left")
    ttk.Combobox(preview_controls, textvariable=limit_var, values=["100", "300", "1000", "3000"], state="readonly", width=8).pack(side="left", padx=(6, 10))
    ttk.Label(preview_controls, text="关键词").pack(side="left")
    ttk.Entry(preview_controls, textvariable=search_var, width=18).pack(side="left", padx=(6, 10))

    preview_container = ttk.Frame(preview_tab)
    preview_container.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    preview_tree = ttk.Treeview(preview_container, show="headings")
    preview_vsb = ttk.Scrollbar(preview_container, orient="vertical", command=preview_tree.yview)
    preview_hsb = ttk.Scrollbar(preview_container, orient="horizontal", command=preview_tree.xview)
    preview_tree.configure(yscroll=preview_vsb.set, xscroll=preview_hsb.set)
    preview_tree.grid(row=0, column=0, sticky="nsew")
    preview_vsb.grid(row=0, column=1, sticky="ns")
    preview_hsb.grid(row=1, column=0, sticky="ew")
    preview_container.rowconfigure(0, weight=1)
    preview_container.columnconfigure(0, weight=1)

    def load_preview() -> None:
        preview_tree.delete(*preview_tree.get_children())
        try:
            limit = int(limit_var.get())
        except ValueError:
            limit = 300
            limit_var.set("300")

        keyword = search_var.get().strip().lower()
        mode = mode_var.get().strip()
        preview_df = df.copy()
        if keyword:
            mask = preview_df.astype(str).apply(lambda col: col.str.lower().str.contains(keyword, na=False))
            preview_df = preview_df[mask.any(axis=1)]
        if mode == "随机抽样":
            preview_df = preview_df.sample(min(limit, len(preview_df)), random_state=42) if len(preview_df) > 0 else preview_df
        else:
            preview_df = preview_df.head(limit).copy()
        columns = ["#"] + [str(column) for column in preview_df.columns]
        preview_tree["columns"] = columns
        for column in columns:
            preview_tree.heading(column, text=column)
            preview_tree.column(column, width=60 if column == "#" else 140, anchor="center")
        for index, row in enumerate(preview_df.itertuples(index=False), start=1):
            preview_tree.insert("", "end", values=(index, *row))
        info_var.set(f"已显示 {len(preview_df)} / {len(df)} 行")

    ttk.Button(preview_controls, text="刷新", command=load_preview).pack(side="left")
    search_var.trace_add("write", lambda *_: load_preview())
    ttk.Button(preview_controls, text="复制字段名", command=lambda: copy_text(app, "\n".join(map(str, df.columns)), "字段名已复制")).pack(side="left", padx=(8, 0))
    ttk.Label(preview_controls, textvariable=info_var).pack(side="right")

    fields_controls = ttk.Frame(fields_tab)
    fields_controls.pack(fill="x", padx=8, pady=(8, 4))
    ttk.Button(fields_controls, text="复制字段名", command=lambda: copy_text(app, "\n".join(map(str, df.columns)), "字段名已复制")).pack(side="left")

    fields_container = ttk.Frame(fields_tab)
    fields_container.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    fields_tree = ttk.Treeview(fields_container, columns=("field", "dtype", "non_null", "missing", "unique", "sample"), show="headings")
    fields_vsb = ttk.Scrollbar(fields_container, orient="vertical", command=fields_tree.yview)
    fields_hsb = ttk.Scrollbar(fields_container, orient="horizontal", command=fields_tree.xview)
    fields_tree.configure(yscroll=fields_vsb.set, xscroll=fields_hsb.set)
    fields_tree.grid(row=0, column=0, sticky="nsew")
    fields_vsb.grid(row=0, column=1, sticky="ns")
    fields_hsb.grid(row=1, column=0, sticky="ew")
    fields_container.rowconfigure(0, weight=1)
    fields_container.columnconfigure(0, weight=1)

    headings = {
        "field": "字段",
        "dtype": "类型",
        "non_null": "非空",
        "missing": "缺失",
        "unique": "唯一值",
        "sample": "示例值",
    }
    widths = {
        "field": 220,
        "dtype": 120,
        "non_null": 90,
        "missing": 90,
        "unique": 90,
        "sample": 360,
    }
    for column, title in headings.items():
        fields_tree.heading(column, text=title)
        fields_tree.column(column, width=widths[column], anchor="center" if column != "sample" else "w")

    for column in df.columns:
        series = df[column]
        fields_tree.insert(
            "",
            "end",
            values=(
                column,
                str(series.dtype),
                int(series.notna().sum()),
                int(series.isna().sum()),
                int(series.nunique(dropna=True)),
                _sample_values(series),
            ),
        )

    report_container = ttk.Frame(report_tab)
    report_container.pack(fill="both", expand=True, padx=8, pady=8)
    report_text = tk.Text(report_container, wrap="word")
    report_text.insert("1.0", build_data_report_text(app))
    report_text.configure(state="disabled")
    report_text.pack(fill="both", expand=True)
    report_buttons = ttk.Frame(report_tab)
    report_buttons.pack(fill="x", padx=8, pady=(0, 8))
    ttk.Button(report_buttons, text="复制报告", command=lambda: copy_text(app, build_data_report_text(app), "检查报告已复制")).pack(side="right")

    load_preview()
    if initial_tab == "report":
        notebook.select(report_tab)
    elif initial_tab == "fields":
        notebook.select(fields_tab)
    else:
        notebook.select(preview_tab)

    dialog.grab_set()
    app.root.wait_window(dialog)


def _theme_menu_items(menu: tk.Menu, app) -> None:
    theme_var = tk.StringVar(value=app.current_theme)

    def change_theme(value: str) -> None:
        app.current_theme = apply_app_theme(value)
        app.app_preferences["theme"] = app.current_theme
        save_app_preferences(app.app_preferences)
        theme_var.set(app.current_theme)
        set_status(app, f"主题已切换为{THEME_LABELS[app.current_theme]}")

    for value, label in THEME_LABELS.items():
        menu.add_radiobutton(label=label, value=value, variable=theme_var, command=lambda v=value: change_theme(v))


def build_app_menu(app) -> None:
    menubar = tk.Menu(app.root)
    app._recent_menus = []

    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="打开文件...\tCtrl+O", command=app.select_file)
    recent_menu = tk.Menu(file_menu, tearoff=0)
    app._recent_menus.append(recent_menu)
    file_menu.add_cascade(label="最近打开", menu=recent_menu)
    file_menu.add_separator()
    file_menu.add_command(label="打开导出目录", command=lambda: open_output_folder(app))
    file_menu.add_command(label="保存当前配置", command=lambda: app.save_config())
    file_menu.add_command(label="清空最近打开", command=lambda: clear_recent_files(app))
    file_menu.add_separator()
    file_menu.add_command(label="首选项...\tCtrl+,", command=lambda: open_preferences_dialog(app))
    file_menu.add_separator()
    file_menu.add_command(label="退出", command=app.on_closing)
    menubar.add_cascade(label="文件", menu=file_menu)

    edit_menu = tk.Menu(menubar, tearoff=0)
    edit_menu.add_command(label="自定义统计量顺序", command=app.open_stats_order)
    edit_menu.add_command(label="自定义排序", command=app.open_custom_sort)
    if hasattr(app, "open_filter_builder"):
        edit_menu.add_command(label="条件筛选...", command=app.open_filter_builder)
    if hasattr(app, "open_group_template_manager"):
        edit_menu.add_command(label="分组模板...", command=app.open_group_template_manager)
    edit_menu.add_separator()
    edit_menu.add_command(label="命令面板...\tCtrl+Shift+P", command=lambda: open_command_palette(app))
    menubar.add_cascade(label="编辑", menu=edit_menu)

    data_menu = tk.Menu(menubar, tearoff=0)
    data_menu.add_command(label="字段智能推荐", command=lambda: open_field_recommendations(app))
    data_menu.add_command(label="日期类统计...", command=lambda: open_date_grouping_dialog(app))
    data_menu.add_command(label="分组映射 / 区间分组...", command=lambda: open_group_mapping_manager(app))
    data_menu.add_command(label="面积单位换算...", command=lambda: open_area_conversion_dialog(app))
    data_menu.add_command(label="多文件合并统计...", command=lambda: merge_multiple_files_into_app(app))
    menubar.add_cascade(label="数据增强", menu=data_menu)

    selection_menu = tk.Menu(menubar, tearoff=0)
    selection_menu.add_command(label="选择统计量", command=app.choose_stats)
    if hasattr(app, "choose_batch_fields"):
        selection_menu.add_command(label="选择批量值字段", command=app.choose_batch_fields)
    if hasattr(app, "choose_group_batch_fields"):
        selection_menu.add_command(label="选择批量分组字段", command=app.choose_group_batch_fields)
    if hasattr(app, "open_filter_builder"):
        selection_menu.add_command(label="管理筛选条件", command=app.open_filter_builder)
    if hasattr(app, "open_group_template_manager"):
        selection_menu.add_command(label="分组模板", command=app.open_group_template_manager)
    menubar.add_cascade(label="选择", menu=selection_menu)

    view_menu = tk.Menu(menubar, tearoff=0)
    theme_menu = tk.Menu(view_menu, tearoff=0)
    _theme_menu_items(theme_menu, app)
    view_menu.add_cascade(label="主题", menu=theme_menu)
    view_menu.add_command(label="当前数据预览\tCtrl+Shift+I", command=lambda: open_data_inspector(app, "preview"))
    view_menu.add_command(label="数据检查报告\tCtrl+Shift+D", command=lambda: open_data_inspector(app, "report"))
    view_menu.add_command(label="预览上次导出结果", command=lambda: preview_last_output(app))
    view_menu.add_command(label="打开命令面板", command=lambda: open_command_palette(app))
    menubar.add_cascade(label="查看", menu=view_menu)

    go_menu = tk.Menu(menubar, tearoff=0)
    go_menu.add_command(label="切换到横版\tCtrl+1", command=lambda: restart_in_mode(app, "horizontal"))
    go_menu.add_command(label="切换到竖版\tCtrl+2", command=lambda: restart_in_mode(app, "vertical"))
    menubar.add_cascade(label="转到", menu=go_menu)

    run_menu = tk.Menu(menubar, tearoff=0)
    run_menu.add_command(label="计算并导出\tF5", command=app.calculate)
    run_menu.add_command(label="打开上次导出文件", command=lambda: open_last_output(app))
    run_menu.add_command(label="打开导出目录", command=lambda: open_output_folder(app))
    if hasattr(app, "easter_egg"):
        run_menu.add_separator()
        run_menu.add_command(label="彩蛋", command=app.easter_egg)
    menubar.add_cascade(label="运行", menu=run_menu)

    batch_menu = tk.Menu(menubar, tearoff=0)
    batch_menu.add_command(label="筛选方案", command=lambda: open_filter_preset_manager(app))
    batch_menu.add_command(label="批量任务中心", command=lambda: open_task_center(app))
    batch_menu.add_command(label="保存工作区", command=lambda: save_workspace_as(app))
    batch_menu.add_command(label="加载工作区", command=lambda: load_workspace_from_file(app))
    batch_menu.add_separator()
    batch_menu.add_command(label="撤销最近一步", command=lambda: undo_last_action(app))
    batch_menu.add_command(label="操作历史", command=lambda: open_operation_history(app))
    menubar.add_cascade(label="流程", menu=batch_menu)

    publish_menu = tk.Menu(menubar, tearoff=0)
    publish_menu.add_command(label="生成发布包", command=lambda: publish_output_bundle(app))
    menubar.add_cascade(label="发布", menu=publish_menu)

    help_menu = tk.Menu(menubar, tearoff=0)
    help_menu.add_command(label="更新历史", command=app.show_update_history)
    help_menu.add_command(label="首选项说明", command=lambda: _show_preferences_help(app))
    menubar.add_cascade(label="帮助", menu=help_menu)

    app.root.config(menu=menubar)
    populate_recent_menus(app)


def bind_shortcuts(app) -> None:
    bindings: list[tuple[str, Callable[[], None]]] = [
        ("<Control-o>", app.select_file),
        ("<F5>", app.calculate),
        ("<Control-comma>", lambda: open_preferences_dialog(app)),
        ("<Control-Shift-P>", lambda: open_command_palette(app)),
        ("<Control-Shift-I>", lambda: open_data_inspector(app, "preview")),
        ("<Control-Shift-D>", lambda: open_data_inspector(app, "report")),
        ("<Control-Key-1>", lambda: restart_in_mode(app, "horizontal")),
        ("<Control-Key-2>", lambda: restart_in_mode(app, "vertical")),
    ]
    for sequence, callback in bindings:
        app.root.bind_all(sequence, lambda _event, fn=callback: _run_binding(fn))


def _run_binding(callback: Callable[[], None]) -> str:
    callback()
    return "break"


def _show_preferences_help(app) -> None:
    messagebox.showinfo(
        "首选项说明",
        "可以设置默认启动模式、界面主题、导出目录、导出文件名时间戳，以及启动时是否恢复最近文件。",
        parent=app.root,
    )


def open_preferences_dialog(app) -> None:
    dialog = tk.Toplevel(app.root)
    dialog.title("首选项")
    dialog.resizable(False, False)
    center_window(dialog, 480, 340)
    dialog.transient(app.root)

    container = ttk.Frame(dialog, padding=14)
    container.pack(fill="both", expand=True)

    current_mode = str(app.app_preferences.get("default_mode", DEFAULT_PREFERENCES["default_mode"]))
    current_output = str(app.app_preferences.get("output_location", DEFAULT_PREFERENCES["output_location"]))
    mode_var = tk.StringVar(value=MODE_LABELS.get(current_mode, "横版"))
    theme_var = tk.StringVar(value=THEME_LABELS.get(app.current_theme, "深色"))
    output_var = tk.StringVar(value=OUTPUT_LOCATION_LABELS.get(current_output, "源文件同目录"))
    reopen_var = tk.BooleanVar(value=bool(app.app_preferences.get("reopen_last_file", False)))
    timestamp_var = tk.BooleanVar(value=bool(app.app_preferences.get("append_timestamp_to_export", False)))

    ttk.Label(container, text="默认启动模式").grid(row=0, column=0, sticky="w", pady=(0, 6))
    ttk.Combobox(container, textvariable=mode_var, values=list(MODE_NAMES.keys()), state="readonly", width=18).grid(row=0, column=1, sticky="w", pady=(0, 6))

    ttk.Label(container, text="界面主题").grid(row=1, column=0, sticky="w", pady=6)
    ttk.Combobox(container, textvariable=theme_var, values=list(THEME_NAMES.keys()), state="readonly", width=18).grid(row=1, column=1, sticky="w", pady=6)

    ttk.Label(container, text="导出目录").grid(row=2, column=0, sticky="w", pady=6)
    ttk.Combobox(container, textvariable=output_var, values=list(OUTPUT_LOCATION_NAMES.keys()), state="readonly", width=18).grid(row=2, column=1, sticky="w", pady=6)

    ttk.Checkbutton(container, text="启动时恢复最近打开的文件", variable=reopen_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 4))
    ttk.Checkbutton(container, text="导出文件名追加时间戳", variable=timestamp_var).grid(row=4, column=0, columnspan=2, sticky="w", pady=4)
    ttk.Label(container, text="提示：默认会自动避让重名导出文件。", justify="left").grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 8))

    buttons = ttk.Frame(container)
    buttons.grid(row=6, column=0, columnspan=2, sticky="e", pady=(14, 0))

    def save_preferences(and_switch: bool = False) -> None:
        selected_mode = MODE_NAMES[mode_var.get()]
        selected_theme = THEME_NAMES[theme_var.get()]
        selected_output = OUTPUT_LOCATION_NAMES[output_var.get()]
        app.app_preferences["default_mode"] = selected_mode
        app.app_preferences["theme"] = selected_theme
        app.app_preferences["output_location"] = selected_output
        app.app_preferences["reopen_last_file"] = reopen_var.get()
        app.app_preferences["append_timestamp_to_export"] = timestamp_var.get()
        save_app_preferences(app.app_preferences)
        app.current_theme = apply_app_theme(selected_theme)
        set_status(app, "首选项已保存")
        dialog.destroy()
        if and_switch and selected_mode != app.mode_name:
            restart_in_mode(app, selected_mode)

    ttk.Button(buttons, text="保存", command=save_preferences).pack(side="right", padx=(8, 0))
    ttk.Button(buttons, text="保存并切换", command=lambda: save_preferences(and_switch=True)).pack(side="right")
    ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="right", padx=(0, 8))

    dialog.grab_set()
    app.root.wait_window(dialog)


def restart_in_mode(app, mode: str) -> None:
    if mode not in MODE_LABELS:
        return

    app.app_preferences["default_mode"] = mode
    save_app_preferences(app.app_preferences)
    if hasattr(app, "save_config"):
        app.save_config(show_msg=False)

    if getattr(sys, "frozen", False):
        os.execl(sys.executable, sys.executable, "--mode", mode)
    else:
        main_path = Path(__file__).resolve().with_name("main.py")
        os.execl(sys.executable, sys.executable, str(main_path), "--mode", mode)


def build_command_actions(app) -> list[tuple[str, Callable[[], None]]]:
    actions: list[tuple[str, Callable[[], None]]] = [
        ("文件: 打开文件", app.select_file),
        ("文件: 打开导出目录", lambda: open_output_folder(app)),
        ("文件: 首选项", lambda: open_preferences_dialog(app)),
        ("编辑: 选择统计量", app.choose_stats),
        ("编辑: 自定义统计量顺序", app.open_stats_order),
        ("编辑: 自定义排序", app.open_custom_sort),
        ("查看: 当前数据预览", lambda: open_data_inspector(app, "preview")),
        ("查看: 数据检查报告", lambda: open_data_inspector(app, "report")),
        ("查看: 切换深色主题", lambda: _apply_theme_from_palette(app, "dark")),
        ("查看: 切换浅色主题", lambda: _apply_theme_from_palette(app, "light")),
        ("转到: 切换到横版", lambda: restart_in_mode(app, "horizontal")),
        ("转到: 切换到竖版", lambda: restart_in_mode(app, "vertical")),
        ("运行: 计算并导出", app.calculate),
        ("运行: 打开上次导出文件", lambda: open_last_output(app)),
        ("运行: 打开导出目录", lambda: open_output_folder(app)),
        ("查看: 预览上次导出结果", lambda: preview_last_output(app)),
        ("帮助: 更新历史", app.show_update_history),
    ]

    if hasattr(app, "open_filter_builder"):
        actions.append(("编辑: 条件筛选", app.open_filter_builder))
    if hasattr(app, "open_group_template_manager"):
        actions.append(("编辑: 分组模板", app.open_group_template_manager))
    actions.extend([
        ("数据增强: 字段智能推荐", lambda: open_field_recommendations(app)),
        ("数据增强: 日期类统计", lambda: open_date_grouping_dialog(app)),
        ("数据增强: 分组映射/区间分组", lambda: open_group_mapping_manager(app)),
        ("数据增强: 面积单位换算", lambda: open_area_conversion_dialog(app)),
        ("数据增强: 多文件合并统计", lambda: merge_multiple_files_into_app(app)),
        ("流程: 筛选方案", lambda: open_filter_preset_manager(app)),
        ("流程: 批量任务中心", lambda: open_task_center(app)),
        ("流程: 保存工作区", lambda: save_workspace_as(app)),
        ("流程: 加载工作区", lambda: load_workspace_from_file(app)),
        ("流程: 撤销最近一步", lambda: undo_last_action(app)),
        ("流程: 操作历史", lambda: open_operation_history(app)),
        ("发布: 生成发布包", lambda: publish_output_bundle(app)),
    ])

    for index, path in enumerate(getattr(app, "recent_files", [])[:8], start=1):
        actions.append((f"文件: 最近打开 {index}. {os.path.basename(path)}", lambda p=path: app.open_recent(p)))
    return actions


def _apply_theme_from_palette(app, theme: str) -> None:
    app.current_theme = apply_app_theme(theme)
    app.app_preferences["theme"] = app.current_theme
    save_app_preferences(app.app_preferences)
    set_status(app, f"主题已切换为{THEME_LABELS[app.current_theme]}")


def open_command_palette(app) -> None:
    dialog = tk.Toplevel(app.root)
    dialog.title("命令面板")
    dialog.resizable(False, False)
    center_window(dialog, 560, 460)
    dialog.transient(app.root)

    container = ttk.Frame(dialog, padding=14)
    container.pack(fill="both", expand=True)

    query_var = tk.StringVar()
    ttk.Label(container, text="输入关键词快速执行命令。").pack(anchor="w")
    entry = ttk.Entry(container, textvariable=query_var)
    entry.pack(fill="x", pady=(6, 10))

    listbox = tk.Listbox(container, height=16)
    listbox.pack(fill="both", expand=True)

    actions = build_command_actions(app)
    visible_actions = actions.copy()

    def refresh(*_args) -> None:
        nonlocal visible_actions
        keyword = query_var.get().strip().lower()
        visible_actions = [item for item in actions if keyword in item[0].lower()] if keyword else actions.copy()

        listbox.delete(0, "end")
        for title, _ in visible_actions:
            listbox.insert("end", title)
        if visible_actions:
            listbox.selection_clear(0, "end")
            listbox.selection_set(0)

    def execute_selected(_event=None) -> None:
        selection = listbox.curselection()
        if not selection:
            return
        _, callback = visible_actions[selection[0]]
        dialog.destroy()
        callback()

    query_var.trace_add("write", refresh)
    listbox.bind("<Double-Button-1>", execute_selected)
    listbox.bind("<Return>", execute_selected)
    entry.bind("<Return>", execute_selected)
    ttk.Button(container, text="执行", command=execute_selected).pack(anchor="e", pady=(10, 0))

    refresh()
    entry.focus_set()
    dialog.grab_set()
    app.root.wait_window(dialog)
