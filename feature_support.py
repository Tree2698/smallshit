from __future__ import annotations

import copy
import json
import os
import sys
import traceback
import zipfile
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
import tkinter as tk
from typing import Any

import pandas as pd

from common_utils import (
    AREA_UNITS,
    DATE_GROUPINGS,
    apply_filter_conditions,
    apply_mapping_rules,
    area_column_name,
    build_date_group_series,
    build_publish_readme,
    friendly_error_message,
    merge_data_files,
    recommend_fields,
)


def ensure_advanced_state(app) -> None:
    defaults = {
        "filter_presets": {},
        "task_queue": [],
        "mapping_rules": [],
        "date_grouping": {"enabled": False, "source": "", "granularity": "月", "output": ""},
        "area_source_unit": "平方米",
        "area_target_unit": "亩",
        "export_charts": True,
        "field_recommendations": {},
        "workspace_path": "",
        "operation_history": [],
        "plugin_folder": str(Path(__file__).resolve().with_name("plugins")),
    }
    for key, value in defaults.items():
        if not hasattr(app, key):
            setattr(app, key, copy.deepcopy(value))


def collect_advanced_state(app) -> dict[str, Any]:
    ensure_advanced_state(app)
    return {
        "filter_presets": copy.deepcopy(getattr(app, "filter_presets", {})),
        "task_queue": copy.deepcopy(getattr(app, "task_queue", [])),
        "mapping_rules": copy.deepcopy(getattr(app, "mapping_rules", [])),
        "date_grouping": copy.deepcopy(getattr(app, "date_grouping", {})),
        "area_source_unit": getattr(app, "area_source_unit", "平方米"),
        "area_target_unit": getattr(app, "area_target_unit", "亩"),
        "export_charts": bool(getattr(app, "export_charts", True)),
        "workspace_path": getattr(app, "workspace_path", ""),
        "last_output_path": getattr(app, "last_output_path", None),
    }


def apply_advanced_state(app, payload: dict[str, Any]) -> None:
    ensure_advanced_state(app)
    if not isinstance(payload, dict):
        return
    app.filter_presets = payload.get("filter_presets", app.filter_presets)
    app.task_queue = payload.get("task_queue", app.task_queue)
    app.mapping_rules = payload.get("mapping_rules", app.mapping_rules)
    app.date_grouping = payload.get("date_grouping", app.date_grouping)
    app.area_source_unit = payload.get("area_source_unit", app.area_source_unit)
    app.area_target_unit = payload.get("area_target_unit", app.area_target_unit)
    app.export_charts = bool(payload.get("export_charts", app.export_charts))
    app.workspace_path = payload.get("workspace_path", app.workspace_path)
    app.last_output_path = payload.get("last_output_path", getattr(app, "last_output_path", None))


def push_history(app, action_name: str) -> None:
    ensure_advanced_state(app)
    snapshot = {
        "action": action_name,
        "advanced": collect_advanced_state(app),
        "ui_state": app._collect_ui_state() if hasattr(app, "_collect_ui_state") else {},
        "last_stats": copy.deepcopy(getattr(app, "last_stats", [])),
        "custom_orders": copy.deepcopy(getattr(app, "custom_orders", {})),
        "filter_conditions": copy.deepcopy(getattr(app, "filter_conditions", [])),
        "active_group_template_name": getattr(app, "active_group_template_name", ""),
    }
    history = getattr(app, "operation_history", [])
    history.append(snapshot)
    app.operation_history = history[-20:]


def undo_last_action(app) -> None:
    history = getattr(app, "operation_history", [])
    if not history:
        messagebox.showinfo("提示", "没有可撤销的操作。", parent=app.root)
        return
    snapshot = history.pop()
    app.operation_history = history
    apply_advanced_state(app, snapshot.get("advanced", {}))
    app.last_stats = snapshot.get("last_stats", getattr(app, "last_stats", []))
    app.custom_orders = snapshot.get("custom_orders", getattr(app, "custom_orders", {}))
    app.filter_conditions = snapshot.get("filter_conditions", getattr(app, "filter_conditions", []))
    app.active_group_template_name = snapshot.get("active_group_template_name", "")
    app.saved_ui_state = snapshot.get("ui_state", {})
    if hasattr(app, "_restore_ui_state") and hasattr(app, "data"):
        app._restore_ui_state(list(app.data.columns))
    if hasattr(app, "_update_loaded_status"):
        app._update_loaded_status()
    messagebox.showinfo("撤销完成", f"已撤销：{snapshot.get('action', '最近操作')}", parent=app.root)


def open_operation_history(app) -> None:
    dialog = tk.Toplevel(app.root)
    dialog.title("操作历史")
    dialog.geometry("560x360")
    dialog.transient(app.root)
    text = tk.Text(dialog, wrap="word")
    text.pack(fill="both", expand=True, padx=12, pady=12)
    history = getattr(app, "operation_history", [])
    if not history:
        text.insert("1.0", "暂无可显示的操作历史。")
    else:
        for index, item in enumerate(reversed(history), start=1):
            text.insert("end", f"{index}. {item.get('action', '未命名操作')}\n")
    text.configure(state="disabled")
    ttk.Button(dialog, text="撤销最近一步", command=lambda: (dialog.destroy(), undo_last_action(app))).pack(pady=(0, 12))
    dialog.grab_set()


def get_filtered_data_generic(app, show_error: bool = True):
    if getattr(app, "data", pd.DataFrame()).empty:
        return app.data.copy()
    try:
        return apply_filter_conditions(app.data, getattr(app, "filter_conditions", []))
    except Exception as exc:
        if show_error:
            messagebox.showerror("筛选条件有误", friendly_error_message(exc), parent=app.root)
        return None


def get_active_data_generic(app):
    filtered = get_filtered_data_generic(app, show_error=False)
    if isinstance(filtered, pd.DataFrame):
        return filtered
    return getattr(app, "data", pd.DataFrame()).copy()


def update_loaded_status_generic(app) -> None:
    data = getattr(app, "data", pd.DataFrame())
    if data.empty:
        if hasattr(app, "status_var"):
            app.status_var.set(f"{getattr(app, 'mode_label', '程序')} 就绪")
        return
    total_rows = len(data)
    total_cols = len(data.columns)
    filtered = get_active_data_generic(app)
    message = f"已加载 {total_rows} 行 / {total_cols} 列"
    if len(filtered) != total_rows:
        message += f"，筛选后 {len(filtered)} 行"
    if hasattr(app, "status_var"):
        app.status_var.set(message)


def prepare_loaded_dataframe(app, df: pd.DataFrame) -> pd.DataFrame:
    ensure_advanced_state(app)
    app.raw_data = df.copy()
    enhanced = rebuild_dataframe(app, df.copy())
    app.field_recommendations = recommend_fields(enhanced)
    return enhanced


def rebuild_dataframe(app, df: pd.DataFrame | None = None) -> pd.DataFrame:
    ensure_advanced_state(app)
    base = df.copy() if isinstance(df, pd.DataFrame) else getattr(app, "raw_data", getattr(app, "data", pd.DataFrame())).copy()
    if base.empty:
        return base
    enhanced = apply_mapping_rules(base, getattr(app, "mapping_rules", []))
    date_grouping = getattr(app, "date_grouping", {})
    if isinstance(date_grouping, dict) and date_grouping.get("enabled"):
        source = str(date_grouping.get("source", "")).strip()
        output = str(date_grouping.get("output", "")).strip() or f"日期（{date_grouping.get('granularity', '月')}）"
        granularity = str(date_grouping.get("granularity", "月"))
        if source in enhanced.columns:
            enhanced[output] = build_date_group_series(enhanced[source], granularity)
            app.date_grouping["output"] = output
    app.data = enhanced.copy()
    app.field_recommendations = recommend_fields(enhanced)
    return enhanced


def refresh_controls_with_dataframe(app) -> None:
    cols = list(getattr(app, "data", pd.DataFrame()).columns)
    for attr in ["val_menu", "ratio_menu", "sheet_menu"]:
        if hasattr(app, attr) and attr != "sheet_menu":
            try:
                getattr(app, attr)["values"] = cols
            except Exception:
                pass
    if hasattr(app, "_refresh_levels"):
        try:
            app._refresh_levels(cols)
        except Exception:
            pass
    elif hasattr(app, "levels"):
        for lvl in app.levels:
            lvl["combo"]["values"] = cols
    if hasattr(app, "_restore_ui_state"):
        try:
            app._restore_ui_state(cols)
        except Exception:
            pass
    if hasattr(app, "_update_loaded_status"):
        app._update_loaded_status()


def open_filter_builder_generic(app) -> None:
    if getattr(app, "data", pd.DataFrame()).empty:
        messagebox.showwarning("提示", "请先加载数据后再设置筛选条件", parent=app.root)
        return

    dialog = tk.Toplevel(app.root)
    dialog.title("条件筛选")
    dialog.geometry("760x420")
    dialog.transient(app.root)

    container = ttk.Frame(dialog, padding=12)
    container.pack(fill="both", expand=True)
    container.rowconfigure(1, weight=1)
    container.columnconfigure(0, weight=1)

    ttk.Label(container, text=f"当前数据共 {len(app.data)} 行，可添加多个筛选条件。", justify="left").grid(row=0, column=0, sticky="w")
    rows_frame = ttk.Frame(container)
    rows_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 8))
    summary_var = tk.StringVar()
    detail_var = tk.StringVar()
    cols = list(app.data.columns)
    condition_rows: list[dict[str, Any]] = []

    def collect_conditions() -> list[dict[str, Any]]:
        output = []
        for item in condition_rows:
            field = item["field"].get().strip()
            op = item["op"].get().strip()
            value = item["value"].get().strip()
            enabled = bool(item["enabled"].get())
            if field and op:
                output.append({"field": field, "op": op, "value": value, "enabled": enabled})
        return output

    def refresh_summary() -> None:
        conditions = collect_conditions()
        try:
            filtered = apply_filter_conditions(app.data, conditions)
            detail_parts = []
            for c in conditions:
                if c.get("enabled", True):
                    if c["op"] in {"为空", "不为空"}:
                        detail_parts.append(f"{c['field']} {c['op']}")
                    else:
                        detail_parts.append(f"{c['field']} {c['op']} {c['value']}")
            detail_var.set("；".join(detail_parts) if detail_parts else "未启用筛选")
            summary_var.set(f"筛选后 {len(filtered)} / {len(app.data)} 行")
        except Exception as exc:
            summary_var.set(f"条件有误：{friendly_error_message(exc)}")

    def remove_condition(item: dict[str, Any]) -> None:
        item["frame"].destroy()
        condition_rows.remove(item)
        if not condition_rows:
            add_condition_row()
        refresh_summary()

    def add_condition_row(condition: dict[str, Any] | None = None) -> None:
        condition = condition or {}
        frame = ttk.Frame(rows_frame)
        frame.pack(fill="x", pady=4)
        enabled_var = tk.BooleanVar(value=bool(condition.get("enabled", True)))
        field_var = tk.StringVar(value=str(condition.get("field", cols[0] if cols else "")))
        op_var = tk.StringVar(value=str(condition.get("op", "等于")))
        value_var = tk.StringVar(value=str(condition.get("value", "")))
        ttk.Checkbutton(frame, text="启用", variable=enabled_var, command=refresh_summary).pack(side="left")
        field_box = ttk.Combobox(frame, values=cols, textvariable=field_var, state="readonly", width=20)
        field_box.pack(side="left", padx=(8, 6))
        op_box = ttk.Combobox(frame, values=["等于", "不等于", "包含", "不包含", "大于", "大于等于", "小于", "小于等于", "为空", "不为空"], textvariable=op_var, state="readonly", width=12)
        op_box.pack(side="left", padx=6)
        value_entry = ttk.Entry(frame, textvariable=value_var, width=22)
        value_entry.pack(side="left", padx=6)
        item = {"frame": frame, "enabled": enabled_var, "field": field_var, "op": op_var, "value": value_var}

        def sync_state(*_args) -> None:
            disabled = op_var.get() in {"为空", "不为空"}
            value_entry.configure(state="disabled" if disabled else "normal")
            if disabled:
                value_var.set("")
            refresh_summary()

        op_var.trace_add("write", sync_state)
        field_box.bind("<<ComboboxSelected>>", lambda _e: refresh_summary())
        value_entry.bind("<KeyRelease>", lambda _e: refresh_summary())
        ttk.Button(frame, text="删除", command=lambda: remove_condition(item)).pack(side="left", padx=(6, 0))
        condition_rows.append(item)
        sync_state()

    controls = ttk.Frame(container)
    controls.grid(row=2, column=0, sticky="ew")
    ttk.Button(controls, text="新增条件", command=add_condition_row).pack(side="left")
    ttk.Button(controls, text="应用当前筛选为方案", command=lambda: open_filter_preset_manager(app, preset_from_dialog=collect_conditions())).pack(side="left", padx=6)
    ttk.Button(controls, text="清空筛选", command=lambda: (condition_rows.clear(), [child.destroy() for child in rows_frame.winfo_children()], add_condition_row(), refresh_summary())).pack(side="left", padx=6)
    ttk.Label(container, textvariable=summary_var).grid(row=3, column=0, sticky="w", pady=(10, 2))
    ttk.Label(container, textvariable=detail_var, wraplength=700).grid(row=4, column=0, sticky="w")

    buttons = ttk.Frame(container)
    buttons.grid(row=5, column=0, sticky="e", pady=(12, 0))

    def apply_conditions() -> None:
        push_history(app, "修改筛选条件")
        conditions = collect_conditions()
        try:
            filtered = apply_filter_conditions(app.data, conditions)
        except Exception as exc:
            messagebox.showerror("筛选条件有误", friendly_error_message(exc), parent=dialog)
            return
        app.filter_conditions = [c for c in conditions if c.get("enabled", True)]
        update_loaded_status_generic(app)
        if hasattr(app, "save_config"):
            app.save_config(show_msg=False)
        dialog.destroy()
        if hasattr(app, "status_var"):
            app.status_var.set(f"筛选已更新，当前保留 {len(filtered)} 行")

    ttk.Button(buttons, text="应用", command=apply_conditions).pack(side="left", padx=(0, 6))
    ttk.Button(buttons, text="关闭", command=dialog.destroy).pack(side="left")

    existing = getattr(app, "filter_conditions", []) or [{}]
    for condition in existing:
        add_condition_row(condition)
    refresh_summary()
    dialog.grab_set()
    app.root.wait_window(dialog)


def open_filter_preset_manager(app, preset_from_dialog: list[dict[str, Any]] | None = None) -> None:
    ensure_advanced_state(app)
    if preset_from_dialog is not None:
        initial_name = simpledialog.askstring("保存筛选方案", "请输入方案名称：", parent=app.root, initialvalue="常用筛选")
        if not initial_name:
            return
        app.filter_presets[initial_name.strip()] = copy.deepcopy(preset_from_dialog)
        if hasattr(app, "save_config"):
            app.save_config(show_msg=False)
        return

    dialog = tk.Toplevel(app.root)
    dialog.title("筛选方案")
    dialog.geometry("620x360")
    dialog.transient(app.root)
    left = ttk.Frame(dialog, padding=12)
    left.pack(side="left", fill="y")
    right = ttk.Frame(dialog, padding=12)
    right.pack(side="left", fill="both", expand=True)
    listbox = tk.Listbox(left, width=24, height=14)
    listbox.pack(fill="y", expand=True)
    preview = tk.Text(right, wrap="word", state="disabled")
    preview.pack(fill="both", expand=True)

    def refresh(target: str = "") -> None:
        names = sorted(app.filter_presets)
        listbox.delete(0, "end")
        for name in names:
            listbox.insert("end", name)
        if names:
            idx = names.index(target) if target in names else 0
            listbox.selection_set(idx)
            update_preview()

    def selected_name() -> str:
        sel = listbox.curselection()
        return listbox.get(sel[0]) if sel else ""

    def update_preview(_e=None) -> None:
        name = selected_name()
        content = "暂无方案"
        if name:
            items = app.filter_presets.get(name, [])
            content = "\n".join([f"- {row['field']} {row['op']} {row.get('value', '')}" for row in items]) or "空方案"
        preview.configure(state="normal")
        preview.delete("1.0", "end")
        preview.insert("1.0", content)
        preview.configure(state="disabled")

    def save_current() -> None:
        name = simpledialog.askstring("保存方案", "方案名称：", parent=dialog, initialvalue=selected_name() or "常用筛选")
        if not name:
            return
        app.filter_presets[name.strip()] = copy.deepcopy(getattr(app, "filter_conditions", []))
        refresh(name.strip())

    def apply_selected() -> None:
        name = selected_name()
        if not name:
            return
        push_history(app, f"应用筛选方案：{name}")
        app.filter_conditions = copy.deepcopy(app.filter_presets.get(name, []))
        update_loaded_status_generic(app)
        dialog.destroy()

    def delete_selected() -> None:
        name = selected_name()
        if not name:
            return
        app.filter_presets.pop(name, None)
        refresh()

    btns = ttk.Frame(right)
    btns.pack(fill="x", pady=(8, 0))
    ttk.Button(btns, text="保存当前筛选", command=save_current).pack(side="left")
    ttk.Button(btns, text="应用", command=apply_selected).pack(side="left", padx=6)
    ttk.Button(btns, text="删除", command=delete_selected).pack(side="left", padx=6)
    ttk.Button(btns, text="关闭", command=dialog.destroy).pack(side="right")
    listbox.bind("<<ListboxSelect>>", update_preview)
    refresh()
    dialog.grab_set()


def capture_task_snapshot(app) -> dict[str, Any]:
    return {
        "name": f"任务{len(getattr(app, 'task_queue', [])) + 1}",
        "mode": getattr(app, "mode_name", "horizontal"),
        "file_path": getattr(app, "excel_path", ""),
        "sheet_name": getattr(app, "current_sheet_name", getattr(app, "sheet_var", tk.StringVar(value="")).get() if hasattr(app, "sheet_var") else ""),
        "ui_state": app._collect_ui_state() if hasattr(app, "_collect_ui_state") else {},
        "last_stats": copy.deepcopy(getattr(app, "last_stats", [])),
        "filter_conditions": copy.deepcopy(getattr(app, "filter_conditions", [])),
        "custom_orders": copy.deepcopy(getattr(app, "custom_orders", {})),
        "advanced": collect_advanced_state(app),
    }


def load_dataframe_for_snapshot(snapshot: dict[str, Any]) -> pd.DataFrame:
    path = str(snapshot.get("file_path", ""))
    if not path or not os.path.exists(path):
        raise FileNotFoundError(f"任务源文件不存在：{path}")
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        from common_utils import read_csv_safely
        return read_csv_safely(path)
    sheet_name = snapshot.get("sheet_name") or None
    excel = pd.ExcelFile(path)
    chosen = sheet_name if sheet_name in excel.sheet_names else excel.sheet_names[0]
    return pd.read_excel(path, sheet_name=chosen)


def apply_snapshot(app, snapshot: dict[str, Any]) -> None:
    ensure_advanced_state(app)
    push_history(app, f"应用任务/工作区：{snapshot.get('name', '未命名')}")
    apply_advanced_state(app, snapshot.get("advanced", {}))
    df = load_dataframe_for_snapshot(snapshot)
    app.excel_path = snapshot.get("file_path", "")
    app.current_sheet_name = snapshot.get("sheet_name", "") or ("CSV" if str(app.excel_path).lower().endswith(".csv") else "")
    app.saved_ui_state = snapshot.get("ui_state", {})
    app.last_stats = snapshot.get("last_stats", getattr(app, "last_stats", []))
    app.filter_conditions = snapshot.get("filter_conditions", [])
    app.custom_orders = snapshot.get("custom_orders", {})
    app.data = prepare_loaded_dataframe(app, df)
    if hasattr(app, "lbl_file"):
        app.lbl_file.config(text=os.path.basename(app.excel_path))
    refresh_controls_with_dataframe(app)


def run_task_snapshot(app, snapshot: dict[str, Any], silent: bool = True) -> tuple[bool, str]:
    original_dialog = getattr(app, "show_result_dialog", None)
    original_error = messagebox.showerror
    captured_output = {"path": ""}
    errors: list[str] = []

    def fake_dialog(filepath, elapsed):
        captured_output["path"] = filepath
        if not silent and original_dialog:
            original_dialog(filepath, elapsed)

    def fake_error(title, msg, **kwargs):
        errors.append(f"{title}: {msg}")

    try:
        apply_snapshot(app, snapshot)
        if original_dialog:
            app.show_result_dialog = fake_dialog
        messagebox.showerror = fake_error
        app._calculate()
        return (len(errors) == 0, captured_output["path"] or "\n".join(errors))
    except Exception as exc:
        return (False, friendly_error_message(exc))
    finally:
        if original_dialog:
            app.show_result_dialog = original_dialog
        messagebox.showerror = original_error


def open_task_center(app) -> None:
    ensure_advanced_state(app)
    dialog = tk.Toplevel(app.root)
    dialog.title("批量任务中心")
    dialog.geometry("760x420")
    dialog.transient(app.root)
    left = ttk.Frame(dialog, padding=12)
    left.pack(side="left", fill="y")
    right = ttk.Frame(dialog, padding=12)
    right.pack(side="left", fill="both", expand=True)
    listbox = tk.Listbox(left, width=28, height=16)
    listbox.pack(fill="y", expand=True)
    preview = tk.Text(right, wrap="word")
    preview.pack(fill="both", expand=True)

    def refresh(target: str = ""):
        listbox.delete(0, "end")
        for task in app.task_queue:
            listbox.insert("end", task.get("name", "未命名任务"))
        if app.task_queue:
            idx = next((i for i, x in enumerate(app.task_queue) if x.get("name") == target), 0)
            listbox.selection_set(idx)
            update_preview()

    def selected_task():
        sel = listbox.curselection()
        return app.task_queue[sel[0]] if sel else None

    def update_preview(_e=None):
        task = selected_task()
        preview.delete("1.0", "end")
        if not task:
            preview.insert("1.0", "暂无任务")
            return
        preview.insert("1.0", json.dumps(task, ensure_ascii=False, indent=2))

    def add_current():
        snapshot = capture_task_snapshot(app)
        name = simpledialog.askstring("任务名称", "请输入任务名称：", parent=dialog, initialvalue=snapshot["name"])
        if not name:
            return
        snapshot["name"] = name.strip()
        app.task_queue.append(snapshot)
        refresh(name.strip())

    def apply_selected():
        task = selected_task()
        if not task:
            return
        apply_snapshot(app, task)
        dialog.destroy()

    def run_all():
        results = []
        for task in app.task_queue:
            ok, msg = run_task_snapshot(app, task, silent=True)
            results.append(f"{'成功' if ok else '失败'} | {task.get('name')} | {msg}")
        messagebox.showinfo("批量任务完成", "\n".join(results) or "没有可执行任务", parent=dialog)
        update_preview()

    def remove_selected():
        task = selected_task()
        if not task:
            return
        app.task_queue.remove(task)
        refresh()

    def export_tasks():
        path = filedialog.asksaveasfilename(title="导出任务", defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(app.task_queue, f, ensure_ascii=False, indent=2)

    def import_tasks():
        path = filedialog.askopenfilename(title="导入任务", filetypes=[("JSON", "*.json")])
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        if isinstance(payload, list):
            app.task_queue = payload
            refresh()

    btns = ttk.Frame(right)
    btns.pack(fill="x", pady=(8, 0))
    ttk.Button(btns, text="添加当前配置", command=add_current).pack(side="left")
    ttk.Button(btns, text="应用所选", command=apply_selected).pack(side="left", padx=6)
    ttk.Button(btns, text="运行全部", command=run_all).pack(side="left", padx=6)
    ttk.Button(btns, text="删除", command=remove_selected).pack(side="left", padx=6)
    ttk.Button(btns, text="导入", command=import_tasks).pack(side="right")
    ttk.Button(btns, text="导出", command=export_tasks).pack(side="right", padx=6)
    listbox.bind("<<ListboxSelect>>", update_preview)
    refresh()
    dialog.grab_set()


def open_area_conversion_dialog(app) -> None:
    ensure_advanced_state(app)
    dialog = tk.Toplevel(app.root)
    dialog.title("面积单位换算")
    dialog.geometry("360x220")
    dialog.transient(app.root)
    source_var = tk.StringVar(value=getattr(app, "area_source_unit", "平方米"))
    target_var = tk.StringVar(value=getattr(app, "area_target_unit", "亩"))
    chart_var = tk.BooleanVar(value=bool(getattr(app, "export_charts", True)))
    frm = ttk.Frame(dialog, padding=14)
    frm.pack(fill="both", expand=True)
    ttk.Label(frm, text="当前面积字段单位").grid(row=0, column=0, sticky="w", pady=6)
    ttk.Combobox(frm, textvariable=source_var, values=AREA_UNITS, state="readonly", width=16).grid(row=0, column=1, sticky="w", pady=6)
    ttk.Label(frm, text="导出换算为").grid(row=1, column=0, sticky="w", pady=6)
    ttk.Combobox(frm, textvariable=target_var, values=AREA_UNITS, state="readonly", width=16).grid(row=1, column=1, sticky="w", pady=6)
    ttk.Checkbutton(frm, text="导出图表工作表", variable=chart_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=10)
    ttk.Label(frm, text=f"导出列名示例：{area_column_name(target_var.get())}").grid(row=3, column=0, columnspan=2, sticky="w")

    def save() -> None:
        push_history(app, "修改面积单位设置")
        app.area_source_unit = source_var.get()
        app.area_target_unit = target_var.get()
        app.export_charts = bool(chart_var.get())
        if hasattr(app, "save_config"):
            app.save_config(show_msg=False)
        dialog.destroy()

    ttk.Button(frm, text="保存", command=save).grid(row=4, column=1, sticky="e", pady=(16, 0))
    dialog.grab_set()


def open_field_recommendations(app) -> None:
    ensure_advanced_state(app)
    rec = recommend_fields(getattr(app, "data", pd.DataFrame()))
    app.field_recommendations = rec
    dialog = tk.Toplevel(app.root)
    dialog.title("字段智能推荐")
    dialog.geometry("620x360")
    dialog.transient(app.root)
    text = tk.Text(dialog, wrap="word")
    text.pack(fill="both", expand=True, padx=12, pady=12)
    lines = [
        f"推荐值字段：{rec.get('recommended_value') or '无'}",
        f"推荐面积字段：{rec.get('recommended_area') or '无'}",
        f"推荐日期字段：{rec.get('recommended_date') or '无'}",
        "",
        f"数值字段：{'、'.join(rec.get('numeric_fields', [])) or '无'}",
        f"日期字段：{'、'.join(rec.get('date_fields', [])) or '无'}",
        f"分类字段：{'、'.join(rec.get('category_fields', [])) or '无'}",
    ]
    text.insert("1.0", "\n".join(lines))
    text.configure(state="disabled")
    btns = ttk.Frame(dialog)
    btns.pack(fill="x", padx=12, pady=(0, 12))

    def apply_rec() -> None:
        push_history(app, "应用字段推荐")
        value_field = rec.get("recommended_value")
        area_field = rec.get("recommended_area")
        if value_field and hasattr(app, "val_var"):
            app.val_var.set(str(value_field))
        if area_field and hasattr(app, "ratio_var"):
            app.ratio_var.set(str(area_field))
        dialog.destroy()

    def apply_date_rec() -> None:
        date_field = rec.get("recommended_date")
        if not date_field:
            return
        app.date_grouping = {"enabled": True, "source": str(date_field), "granularity": "月", "output": f"{date_field}_月"}
        rebuild_dataframe(app)
        refresh_controls_with_dataframe(app)
        dialog.destroy()

    ttk.Button(btns, text="应用值/面积推荐", command=apply_rec).pack(side="left")
    ttk.Button(btns, text="启用日期分组推荐", command=apply_date_rec).pack(side="left", padx=6)
    ttk.Button(btns, text="关闭", command=dialog.destroy).pack(side="right")
    dialog.grab_set()


def open_date_grouping_dialog(app) -> None:
    ensure_advanced_state(app)
    dialog = tk.Toplevel(app.root)
    dialog.title("日期类统计")
    dialog.geometry("420x240")
    dialog.transient(app.root)
    cols = list(getattr(app, "raw_data", getattr(app, "data", pd.DataFrame())).columns)
    state = getattr(app, "date_grouping", {"enabled": False, "source": "", "granularity": "月", "output": ""})
    enabled_var = tk.BooleanVar(value=bool(state.get("enabled", False)))
    field_var = tk.StringVar(value=str(state.get("source", cols[0] if cols else "")))
    granularity_var = tk.StringVar(value=str(state.get("granularity", "月")))
    output_var = tk.StringVar(value=str(state.get("output", "")))
    frm = ttk.Frame(dialog, padding=14)
    frm.pack(fill="both", expand=True)
    ttk.Checkbutton(frm, text="启用日期分组", variable=enabled_var).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
    ttk.Label(frm, text="日期字段").grid(row=1, column=0, sticky="w", pady=6)
    ttk.Combobox(frm, textvariable=field_var, values=cols, state="readonly", width=20).grid(row=1, column=1, sticky="w", pady=6)
    ttk.Label(frm, text="统计粒度").grid(row=2, column=0, sticky="w", pady=6)
    ttk.Combobox(frm, textvariable=granularity_var, values=DATE_GROUPINGS, state="readonly", width=20).grid(row=2, column=1, sticky="w", pady=6)
    ttk.Label(frm, text="输出字段名").grid(row=3, column=0, sticky="w", pady=6)
    ttk.Entry(frm, textvariable=output_var, width=24).grid(row=3, column=1, sticky="w", pady=6)

    def sync_output(*_args):
        if not output_var.get().strip():
            output_var.set(f"{field_var.get()}_{granularity_var.get()}")
    field_var.trace_add("write", sync_output)
    granularity_var.trace_add("write", sync_output)
    sync_output()

    def save() -> None:
        push_history(app, "修改日期统计设置")
        app.date_grouping = {
            "enabled": bool(enabled_var.get()),
            "source": field_var.get().strip(),
            "granularity": granularity_var.get().strip(),
            "output": output_var.get().strip() or f"{field_var.get()}_{granularity_var.get()}",
        }
        rebuild_dataframe(app)
        refresh_controls_with_dataframe(app)
        if hasattr(app, "save_config"):
            app.save_config(show_msg=False)
        dialog.destroy()

    ttk.Button(frm, text="保存", command=save).grid(row=4, column=1, sticky="e", pady=(16, 0))
    dialog.grab_set()


def open_group_mapping_manager(app) -> None:
    ensure_advanced_state(app)
    dialog = tk.Toplevel(app.root)
    dialog.title("分组映射 / 区间分组")
    dialog.geometry("760x460")
    dialog.transient(app.root)
    left = ttk.Frame(dialog, padding=12)
    left.pack(side="left", fill="y")
    right = ttk.Frame(dialog, padding=12)
    right.pack(side="left", fill="both", expand=True)
    listbox = tk.Listbox(left, width=24, height=18)
    listbox.pack(fill="y", expand=True)
    preview = tk.Text(right, wrap="word")
    preview.pack(fill="both", expand=True)

    def refresh(target: str = ""):
        listbox.delete(0, "end")
        for item in app.mapping_rules:
            listbox.insert("end", item.get("output", "未命名规则"))
        if app.mapping_rules:
            idx = next((i for i, x in enumerate(app.mapping_rules) if x.get("output") == target), 0)
            listbox.selection_set(idx)
            update_preview()

    def selected_rule():
        sel = listbox.curselection()
        return app.mapping_rules[sel[0]] if sel else None

    def update_preview(_e=None):
        preview.delete("1.0", "end")
        rule = selected_rule()
        if not rule:
            preview.insert("1.0", "暂无规则")
            return
        preview.insert("1.0", json.dumps(rule, ensure_ascii=False, indent=2))

    def add_exact_rule():
        cols = list(getattr(app, "raw_data", getattr(app, "data", pd.DataFrame())).columns)
        source = simpledialog.askstring("源字段", f"请输入源字段名，候选：{', '.join(cols[:12])}", parent=dialog)
        if not source:
            return
        output = simpledialog.askstring("输出字段", "请输入新字段名：", parent=dialog, initialvalue=f"{source}_映射")
        if not output:
            return
        lines = simpledialog.askstring("映射规则", "按“原值=新值”逐行输入，例如：\nA=甲\nB=乙", parent=dialog)
        if not lines:
            return
        mapping = {}
        for line in lines.splitlines():
            if "=" in line:
                left_value, right_value = line.split("=", 1)
                mapping[left_value.strip()] = right_value.strip()
        app.mapping_rules.append({"type": "exact", "source": source.strip(), "output": output.strip(), "mapping": mapping, "enabled": True})
        rebuild_dataframe(app)
        refresh_controls_with_dataframe(app)
        refresh(output.strip())

    def add_bin_rule():
        cols = list(getattr(app, "raw_data", getattr(app, "data", pd.DataFrame())).columns)
        source = simpledialog.askstring("源字段", f"请输入数值字段名，候选：{', '.join(cols[:12])}", parent=dialog)
        if not source:
            return
        output = simpledialog.askstring("输出字段", "请输入新字段名：", parent=dialog, initialvalue=f"{source}_分组")
        if not output:
            return
        breaks = simpledialog.askstring("区间断点", "请输入逗号分隔断点，例如：0,10,20,50,100", parent=dialog)
        if not breaks:
            return
        break_values = [item.strip() for item in breaks.split(",") if item.strip()]
        app.mapping_rules.append({"type": "bins", "source": source.strip(), "output": output.strip(), "breaks": break_values, "enabled": True})
        rebuild_dataframe(app)
        refresh_controls_with_dataframe(app)
        refresh(output.strip())

    def add_date_rule():
        cols = list(getattr(app, "raw_data", getattr(app, "data", pd.DataFrame())).columns)
        source = simpledialog.askstring("源字段", f"请输入日期字段名，候选：{', '.join(cols[:12])}", parent=dialog)
        if not source:
            return
        output = simpledialog.askstring("输出字段", "请输入新字段名：", parent=dialog, initialvalue=f"{source}_月")
        if not output:
            return
        granularity = simpledialog.askstring("日期粒度", "请输入 年 / 季度 / 月 / 周 / 日", parent=dialog, initialvalue="月")
        if not granularity:
            return
        app.mapping_rules.append({"type": "date", "source": source.strip(), "output": output.strip(), "granularity": granularity.strip(), "enabled": True})
        rebuild_dataframe(app)
        refresh_controls_with_dataframe(app)
        refresh(output.strip())

    def delete_selected():
        rule = selected_rule()
        if not rule:
            return
        app.mapping_rules.remove(rule)
        rebuild_dataframe(app)
        refresh_controls_with_dataframe(app)
        refresh()

    btns = ttk.Frame(right)
    btns.pack(fill="x", pady=(8, 0))
    ttk.Button(btns, text="添加映射规则", command=add_exact_rule).pack(side="left")
    ttk.Button(btns, text="添加区间规则", command=add_bin_rule).pack(side="left", padx=6)
    ttk.Button(btns, text="添加日期规则", command=add_date_rule).pack(side="left", padx=6)
    ttk.Button(btns, text="删除", command=delete_selected).pack(side="left", padx=6)
    ttk.Button(btns, text="关闭", command=dialog.destroy).pack(side="right")
    listbox.bind("<<ListboxSelect>>", update_preview)
    refresh()
    dialog.grab_set()


def save_workspace_as(app) -> None:
    ensure_advanced_state(app)
    path = filedialog.asksaveasfilename(title="保存工作区", defaultextension=".json", filetypes=[("JSON", "*.json")])
    if not path:
        return
    payload = capture_task_snapshot(app)
    payload["app_preferences"] = getattr(app, "app_preferences", {})
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    app.workspace_path = path
    if hasattr(app, "save_config"):
        app.save_config(show_msg=False)
    messagebox.showinfo("保存完成", f"工作区已保存到：\n{path}", parent=app.root)


def load_workspace_from_file(app) -> None:
    path = filedialog.askopenfilename(title="加载工作区", filetypes=[("JSON", "*.json")])
    if not path:
        return
    with open(path, "r", encoding="utf-8") as f:
        payload = json.load(f)
    apply_snapshot(app, payload)
    app.workspace_path = path
    messagebox.showinfo("加载完成", f"已加载工作区：\n{path}", parent=app.root)


def publish_output_bundle(app, filepath: str | None = None) -> str | None:
    target = filepath or getattr(app, "last_output_path", None)
    if not target or not os.path.exists(target):
        messagebox.showinfo("提示", "还没有可发布的导出结果。", parent=app.root)
        return None
    metadata = {
        "导出文件": os.path.basename(target),
        "源文件": os.path.basename(getattr(app, "excel_path", "") or "未加载文件"),
        "工作表": getattr(app, "current_sheet_name", "当前数据"),
        "模式": getattr(app, "mode_label", getattr(app, "mode_name", "程序")),
        "统计量": "、".join(getattr(app, "last_stats", [])),
        "筛选条件": json.dumps(getattr(app, "filter_conditions", []), ensure_ascii=False),
        "面积换算": f"{getattr(app, 'area_source_unit', '平方米')} -> {getattr(app, 'area_target_unit', '亩')}",
    }
    publish_dir = os.path.dirname(target)
    readme_path = os.path.join(publish_dir, Path(target).stem + "_发布说明.txt")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(build_publish_readme(metadata))
    zip_path = os.path.join(publish_dir, Path(target).stem + "_发布包.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(target, arcname=os.path.basename(target))
        zf.write(readme_path, arcname=os.path.basename(readme_path))
    return zip_path


def merge_multiple_files_into_app(app) -> None:
    paths = filedialog.askopenfilenames(title="选择多个文件进行合并", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")])
    if not paths:
        return
    push_history(app, "合并多个文件")
    merged, summary = merge_data_files(list(paths), getattr(app, "current_sheet_name", None), include_source=True)
    if merged.empty:
        messagebox.showwarning("提示", "没有可合并的数据。", parent=app.root)
        return
    app.excel_path = os.path.join(os.path.dirname(paths[0]), "合并数据.xlsx")
    app.current_sheet_name = "合并数据"
    app.merge_summary = summary
    app.data = prepare_loaded_dataframe(app, merged)
    if hasattr(app, "lbl_file"):
        app.lbl_file.config(text=f"已合并 {len(paths)} 个文件")
    refresh_controls_with_dataframe(app)
    messagebox.showinfo("合并完成", f"已合并 {len(paths)} 个文件，共 {len(merged)} 行。", parent=app.root)


def show_exception_dialog(app, title: str, exc: Exception) -> None:
    dialog = tk.Toplevel(app.root)
    dialog.title(title)
    dialog.geometry("760x420")
    dialog.transient(app.root)
    text = tk.Text(dialog, wrap="word")
    text.pack(fill="both", expand=True, padx=12, pady=12)
    text.insert("1.0", friendly_error_message(exc) + "\n\n详细堆栈：\n" + traceback.format_exc())
    btns = ttk.Frame(dialog)
    btns.pack(fill="x", padx=12, pady=(0, 12))
    ttk.Button(btns, text="复制详情", command=lambda: app.root.clipboard_append(text.get("1.0", "end-1c"))).pack(side="left")
    ttk.Button(btns, text="关闭", command=dialog.destroy).pack(side="right")
    dialog.grab_set()


def collect_group_template_payload(app) -> dict[str, Any]:
    return {
        "ui_state": app._collect_ui_state() if hasattr(app, "_collect_ui_state") else {},
        "stats": copy.deepcopy(getattr(app, "last_stats", [])),
        "filter_conditions": copy.deepcopy(getattr(app, "filter_conditions", [])),
        "advanced": collect_advanced_state(app),
    }


def apply_group_template_payload(app, payload: dict[str, Any], template_name: str = "") -> bool:
    if getattr(app, "data", pd.DataFrame()).empty:
        messagebox.showwarning("提示", "请先加载数据后再应用模板", parent=app.root)
        return False
    if not isinstance(payload, dict):
        messagebox.showerror("错误", "模板内容无效", parent=app.root)
        return False
    push_history(app, f"应用模板：{template_name or '未命名'}")
    apply_advanced_state(app, payload.get("advanced", {}))
    app.saved_ui_state = payload.get("ui_state", {})
    if hasattr(app, "_restore_ui_state"):
        app._restore_ui_state(list(app.data.columns))
    app.last_stats = [stat for stat in payload.get("stats", []) if stat in getattr(app, "all_stats", [])] or getattr(app, "last_stats", [])
    app.filter_conditions = payload.get("filter_conditions", getattr(app, "filter_conditions", []))
    app.active_group_template_name = template_name
    update_loaded_status_generic(app)
    return True


def open_group_template_manager_generic(app) -> None:
    dialog = tk.Toplevel(app.root)
    dialog.title("分组模板")
    dialog.geometry("720x420")
    dialog.transient(app.root)
    container = ttk.Frame(dialog, padding=12)
    container.pack(fill="both", expand=True)
    container.columnconfigure(1, weight=1)
    container.rowconfigure(0, weight=1)
    left = ttk.Frame(container)
    left.grid(row=0, column=0, sticky="ns", padx=(0, 12))
    right = ttk.Frame(container)
    right.grid(row=0, column=1, sticky="nsew")
    right.rowconfigure(1, weight=1)
    ttk.Label(left, text="已保存模板").pack(anchor="w")
    listbox = tk.Listbox(left, width=24, height=16)
    listbox.pack(fill="y", expand=True, pady=(6, 8))
    ttk.Label(right, text="模板内容").grid(row=0, column=0, sticky="w")
    preview = tk.Text(right, wrap="word", state="disabled", height=16)
    preview.grid(row=1, column=0, sticky="nsew", pady=(6, 8))
    buttons = ttk.Frame(right)
    buttons.grid(row=2, column=0, sticky="e")

    def selected_name():
        sel = listbox.curselection()
        return listbox.get(sel[0]) if sel else ""

    def refresh(target_name: str = ""):
        names = sorted(getattr(app, "group_templates", {}))
        listbox.delete(0, "end")
        for name in names:
            listbox.insert("end", name)
        if names:
            idx = names.index(target_name) if target_name in names else 0
            listbox.selection_set(idx)
            update_preview()

    def update_preview(_event=None):
        name = selected_name()
        content = "暂无模板"
        if name:
            content = json.dumps(app.group_templates.get(name, {}), ensure_ascii=False, indent=2)
        preview.configure(state="normal")
        preview.delete("1.0", "end")
        preview.insert("1.0", content)
        preview.configure(state="disabled")

    def save_current():
        name = simpledialog.askstring("保存模板", "请输入模板名称：", parent=dialog, initialvalue=selected_name() or "常用模板")
        if not name:
            return
        app.group_templates[name.strip()] = collect_group_template_payload(app)
        app.active_group_template_name = name.strip()
        refresh(name.strip())

    def apply_selected():
        name = selected_name()
        if not name:
            return
        if apply_group_template_payload(app, app.group_templates.get(name, {}), name):
            if hasattr(app, "save_config"):
                app.save_config(show_msg=False)
            dialog.destroy()

    def delete_selected():
        name = selected_name()
        if not name:
            return
        app.group_templates.pop(name, None)
        if getattr(app, "active_group_template_name", "") == name:
            app.active_group_template_name = ""
        refresh()

    ttk.Button(buttons, text="保存当前", command=save_current).pack(side="left", padx=(0, 6))
    ttk.Button(buttons, text="应用", command=apply_selected).pack(side="left", padx=6)
    ttk.Button(buttons, text="删除", command=delete_selected).pack(side="left", padx=6)
    ttk.Button(buttons, text="关闭", command=dialog.destroy).pack(side="left", padx=(6, 0))
    listbox.bind("<<ListboxSelect>>", update_preview)
    refresh(getattr(app, "active_group_template_name", ""))
    dialog.grab_set()
