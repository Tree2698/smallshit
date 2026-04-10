# app_vertical.py

import json
import os
import threading
import time
from decimal import Decimal, ROUND_HALF_UP
from functools import partial

import numpy as np
import pandas as pd
import tkinter as tk
from openpyxl.styles import Alignment, Border, Side
from tkinter import filedialog, messagebox, ttk

from common_utils import open_with_default_app, read_csv_safely, round2, sort_key

# 当前程序版本号——每次发布时请手动更新
APP_VERSION = "1.0.8"

UPDATE_CONTENT = """
小捞翔·至尊版 v1.0.8 更新日志：
- 新增：竖版自定义统计量排序
- 新增：竖版最近打开
- 新增：竖版自定义排序
- 新增：竖版快速打开文件
- 新增：竖版浏览文件
- 新增：竖版面积自定义小数位数
- 修复：在选择横竖版时点击关闭，无法正常关闭
- 修复：竖版删除被限制行
"""

# 默认配置文件名（可自定义目录）
CONFIG_FILE = "small_shit_vertical.json"

MAX_LEVELS = 3  # 最多三级分类




class VerticalApp:
    def __init__(self, root):
        self.root = root
        # -------- 新增：自定义排序状态 --------
        self.custom_orders = {}  # 用户自定义的每列排序列表
        self.original_orders = {}  # 首次记录每列原始唯一值顺序
        # -------- 新增 recent_files 存储 最近打开 列表 --------
        self.recent_files = []
        root.title("小捞翔 PRO MAX PLUS")
        root.geometry("500x560")  # 增加高度以适应新控件
        root.resizable(False, False)

        # 在窗口关闭前保存配置
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 初始化属性
        self.excel_path = ""
        self.data = pd.DataFrame()
        self.batch_fields = []  # 批量值字段列表
        self.group_batch_fields = []  # 批量分组字段列表
        self.update_history = []  # 更新历史记录
        self.shown_version = ""  # 已显示的版本

        # 保存上次选择的统计量，初始全选
        self.all_stats = [
            "数量", "平均值", "中位值", "最大值",
            "最小值", "标准差", "变异系数", "数量占比"
        ]
        self.last_stats = self.all_stats.copy()

        # 构建菜单栏
        self._build_menu()
        self._build_file_sheet_selector()
        self._build_levels_frame()
        self._build_field_controls()
        self._build_batch_controls()  # 添加批量控制组件
        self._build_buttons()

        # 载入上次保存的配置
        self.load_config()
        # 检查更新
        self.check_for_update()

    def on_closing(self):
        # 退出前保存配置
        self.save_config(show_msg=False)
        self.root.destroy()

    def save_config(self, path=None, show_msg=True):
        """保存配置到JSON文件"""
        cfg = {
            "shown_version": getattr(self, "shown_version", ""),
            "update_history": getattr(self, "update_history", []),
            "last_stats": self.last_stats,
            "recent_files": self.recent_files
            , "custom_orders": self.custom_orders,
            "area_decimals": self.area_decimals_var.get()
        }
        cfg_path = path or CONFIG_FILE
        try:
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            if show_msg:
                messagebox.showinfo("配置已保存", f"配置已保存到：\n{cfg_path}")
        except Exception as e:
            if show_msg:
                messagebox.showerror("保存失败", str(e))

    def load_config(self, path=None):
        """从JSON文件加载配置"""
        cfg_path = path or CONFIG_FILE
        if not os.path.exists(cfg_path):
            return

        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)

            # 已展示过更新提示的版本号
            self.shown_version = cfg.get("shown_version", "")

            # 已累计的更新日志历史（多条）
            self.update_history = cfg.get("update_history", [])

            # 上次选择的统计量列表
            self.last_stats = cfg.get("last_stats", self.all_stats.copy())
            # 恢复最近打开列表 & 自定义排序
            self.recent_files = cfg.get("recent_files", [])
            self.custom_orders = cfg.get("custom_orders", {})
            dec = cfg.get("area_decimals", 2)
            self.area_decimals_var.set(dec)
        # 配置载入后刷新最近打开菜单（如果已经创建）
            if hasattr(self, "create_recent_menu"):
                self.create_recent_menu()


        except Exception as e:
            messagebox.showerror("加载失败", str(e))

    def check_for_update(self):
        """检查并显示更新日志"""
        if self.shown_version != APP_VERSION:
            dlg = tk.Toplevel(self.root)
            dlg.title(f"更新日志 v{APP_VERSION}")
            dlg.geometry("400x300")
            dlg.resizable(False, False)

            # — 文本区 —
            frm_txt = ttk.Frame(dlg, padding=8)
            frm_txt.pack(fill="both", expand=True)
            txt = tk.Text(frm_txt, wrap="word", height=10)
            txt.insert("1.0", UPDATE_CONTENT.strip())
            txt.configure(state="disabled")  # 只读
            txt.pack(fill="both", expand=True)

            # — 按钮区 —
            frm_btn = ttk.Frame(dlg, padding=(8, 8))
            frm_btn.pack(fill="x")

            def _on_close():
                self.shown_version = APP_VERSION
                # 把本次更新内容加入历史（避免重复）
                if UPDATE_CONTENT.strip() not in self.update_history:
                    self.update_history.append(UPDATE_CONTENT.strip())
                self.save_config(show_msg=False)
                dlg.destroy()

            btn = ttk.Button(frm_btn, text="知道了", command=_on_close)
            btn.pack(side="right", padx=4, pady=4)

            dlg.transient(self.root)
            dlg.grab_set()
            self.root.wait_window(dlg)

    def show_update_history(self):
        """显示更新历史记录"""
        history = getattr(self, "update_history", [])
        if not history:
            messagebox.showinfo("更新历史", "暂无更新记录。")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("更新历史")
        dlg.geometry("400x300")
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=8)
        frm.pack(fill="both", expand=True)

        txt = tk.Text(frm, wrap="word")
        txt.insert("1.0", "\n\n".join(history))
        txt.configure(state="disabled")
        txt.pack(fill="both", expand=True)

        btn = ttk.Button(frm, text="关闭", command=dlg.destroy)
        btn.pack(pady=6)

        dlg.transient(self.root)
        dlg.grab_set()
        self.root.wait_window(dlg)

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        # 添加关于菜单
        menubar.add_command(label="关于", command=self.show_update_history)
        # -------- 新增：最近打开 菜单 --------
        self.recent_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="最近打开", menu=self.recent_menu)
        self.create_recent_menu()
        self.root.config(menu=menubar)
    def create_recent_menu(self):
        """刷新 最近打开 下拉列表"""
        self.recent_menu.delete(0, "end")
        for p in self.recent_files:
            name = os.path.basename(p)
            self.recent_menu.add_command(
                label=name,
                command=lambda _p=p: self.open_recent(_p)
            )
        if not self.recent_files:
            self.recent_menu.add_command(label="— 无记录 —", state="disabled")

    def add_recent(self, path):
        """把 path 插入到 最近打开（去重 + 限长 20）"""
        if path in self.recent_files:
            self.recent_files.remove(path)
        self.recent_files.insert(0, path)
        self.recent_files = self.recent_files[:20]
        # 静默保存并刷新菜单
        self.save_config(show_msg=False)
        self.create_recent_menu()

    def open_recent(self, path):
        """通过 最近打开 重新载入文件"""
        if not os.path.exists(path):
            messagebox.showerror("文件不存在", path)
            self.recent_files.remove(path)
            self.save_config(show_msg=False)
            self.create_recent_menu()
            return

        # 跟 select_file 一致的加载逻辑
        self.excel_path = path
        self.lbl_file.config(text=os.path.basename(path))
        if path.lower().endswith(".csv"):
            self.data = read_csv_safely(path)
            self.sheet_menu["values"] = []
            self.sheet_menu.state(["disabled"])
            self._on_sheet_loaded(self.data, "CSV")
        else:
            sheets = pd.ExcelFile(path).sheet_names
            self.sheet_menu["values"] = sheets
            self.sheet_menu.state(["!disabled"])
            self.sheet_var.set(sheets[0])
            self.load_sheet(sheets[0])

        self.add_recent(path)

    def _build_file_sheet_selector(self):
        frm = ttk.Frame(self.root)
        frm.pack(fill="x", padx=10, pady=8)

        ttk.Button(frm, text="选择 Excel", command=self.select_file) \
            .grid(row=0, column=0, sticky="w")

        self.lbl_file = ttk.Label(frm, text="未选择文件")
        self.lbl_file.grid(row=0, column=1, padx=8, sticky="w")

        ttk.Label(frm, text="子表:") \
            .grid(row=1, column=0, pady=6, sticky="w")

        self.sheet_var = tk.StringVar()
        self.sheet_menu = ttk.Combobox(
            frm, textvariable=self.sheet_var,
            state="readonly", width=28
        )
        self.sheet_menu.grid(row=1, column=1, sticky="w")
        self.sheet_menu.bind(
            "<<ComboboxSelected>>",
            lambda e: self.load_sheet(self.sheet_var.get())
        )

    def _build_levels_frame(self):
        self.frame_levels = ttk.LabelFrame(
            self.root, text="分类级别 (1～3 级)"
        )
        self.frame_levels.pack(fill="x", padx=10, pady=6)

        self.levels = []
        self.add_btn = ttk.Button(
            self.frame_levels, text="+ 添加级别",
            command=self.add_level
        )
        self.add_level()
        self.add_level()

    def _build_field_controls(self):
        frm = ttk.Frame(self.root)
        frm.pack(fill="x", padx=10, pady=6)

        # 值字段
        ttk.Label(frm, text="值字段:") \
            .grid(row=0, column=0, sticky="w")
        self.val_var = tk.StringVar()
        self.val_menu = ttk.Combobox(
            frm, textvariable=self.val_var,
            state="readonly", width=28
        )
        self.val_menu.grid(row=0, column=1, sticky="w")

        # 面积占比字段
        ttk.Label(frm, text="面积占比字段:") \
            .grid(row=1, column=0, sticky="w")
        self.ratio_var = tk.StringVar()
        self.ratio_menu = ttk.Combobox(
            frm, textvariable=self.ratio_var,
            state="readonly", width=28
        )
        self.ratio_menu.grid(row=1, column=1, sticky="w")
        # 面积小数位
        ttk.Label(frm, text="面积小数位:") \
            .grid(row=2, column=0, sticky="w")

        self.area_decimals_var = tk.IntVar(value=2)
        self.area_decimals_spin = ttk.Spinbox(
            frm, from_=0, to=10, textvariable=self.area_decimals_var,
            width=5, state="readonly"
        )
        self.area_decimals_spin.grid(row=2, column=1, sticky="w")
        self.area_cb = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm, text="计算面积占比", variable=self.area_cb
        ).grid(row=1, column=2, padx=5)

    def _build_batch_controls(self):
        """添加批量功能控件"""
        frm = ttk.LabelFrame(self.root, text="批量功能")
        frm.pack(fill="x", padx=10, pady=6)

        # 批量值字段
        self.batch_var = tk.BooleanVar(value=False)
        batch_frame = ttk.Frame(frm)
        batch_frame.pack(fill="x", padx=5, pady=5)
        ttk.Checkbutton(
            batch_frame, text="批量值字段",
            variable=self.batch_var,
            command=self.toggle_batch
        ).pack(side="left", padx=5)
        self.batch_btn = ttk.Button(
            batch_frame, text="选择值字段",
            command=self.choose_batch_fields,
            state="disabled"
        )
        self.batch_btn.pack(side="left", padx=5)

        # 批量分组字段
        self.group_batch_var = tk.BooleanVar(value=False)
        group_frame = ttk.Frame(frm)
        group_frame.pack(fill="x", padx=5, pady=5)
        # -------- 新增：自定义统计量顺序 按钮 --------
        btn = ttk.Button(frm, text="自定义统计量顺序", command=self.open_stats_order)
        btn.pack(side="right", padx=5)

        ttk.Checkbutton(
            group_frame, text="批量分组字段",
            variable=self.group_batch_var,
            command=self.toggle_group_batch
        ).pack(side="left", padx=5)
        self.group_batch_btn = ttk.Button(
            group_frame, text="选择分组字段",
            command=self.choose_group_batch_fields,
            state="disabled"
        )
        # -------- 新增：选择统计量 按钮，与横版保持一致 --------
        stats_btn = ttk.Button(frm, text="选择统计量", command=self.choose_stats)
        stats_btn.pack(side="right", padx=5)
        sort_btn = ttk.Button(frm, text="自定义排序", command=self.open_custom_sort)
        sort_btn.pack(side="right", padx=5)
        self.group_batch_btn.pack(side="left", padx=5)




    def toggle_batch(self):
        """切换批量值字段按钮状态"""
        if self.batch_var.get():
            self.batch_btn.state(["!disabled"])
        else:
            self.batch_btn.state(["disabled"])

    def toggle_group_batch(self):
        """切换批量分组字段按钮状态"""
        if self.group_batch_var.get():
            self.group_batch_btn.state(["!disabled"])
        else:
            self.group_batch_btn.state(["disabled"])

    def choose_batch_fields(self):
        """选择多个值字段"""
        if self.data.empty:
            messagebox.showwarning("提示", "请先加载数据")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("选择值字段(多选)")
        cols = list(self.data.columns)
        lb = tk.Listbox(dlg, selectmode="multiple", height=10)
        for c in cols:
            lb.insert("end", c)
        lb.pack(padx=12, pady=6)

        def on_ok():
            self.batch_fields = [cols[i] for i in lb.curselection()]
            dlg.destroy()

        ttk.Button(dlg, text="确定", command=on_ok, width=10).pack(pady=6)
        dlg.grab_set()
        self.root.wait_window(dlg)

    def choose_group_batch_fields(self):
        """选择多个分组字段"""
        if self.data.empty:
            messagebox.showwarning("提示", "请先加载数据")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("选择分组字段(多选)")
        cols = list(self.data.columns)
        lb = tk.Listbox(dlg, selectmode="multiple", height=10)
        for c in cols:
            lb.insert("end", c)
        lb.pack(padx=12, pady=6)

        def on_ok():
            self.group_batch_fields = [cols[i] for i in lb.curselection()]
            dlg.destroy()

        ttk.Button(dlg, text="确定", command=on_ok, width=10).pack(pady=6)
        dlg.grab_set()
        self.root.wait_window(dlg)

    def _build_buttons(self):
        btnf = ttk.Frame(self.root)
        btnf.pack(pady=12)

        ttk.Button(
            btnf, text="计算捞翔并导出",
            command=self.calculate, width=16
        ).grid(row=0, column=0, padx=10)

        ttk.Button(
            btnf, text="小捞翔彩蛋",
            command=self.easter_egg, width=10
        ).grid(row=0, column=1)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="选择 xlsx 文件",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("所有文件", "*.*")]
        )
        if not path:
            return
        self.add_recent(path)
        self.excel_path = path
        self.lbl_file.config(text=os.path.basename(path))

        try:
            if path.lower().endswith(".csv"):
                # 如果是CSV文件，直接读取
                self.data = read_csv_safely(path)
                self.sheet_menu["values"] = []
                self.sheet_menu.state(["disabled"])
                self._on_sheet_loaded(self.data, "CSV")
            else:
                # Excel文件
                sheets = pd.ExcelFile(path).sheet_names
                self.sheet_menu["values"] = sheets
                self.sheet_menu.state(["!disabled"])
                self.sheet_var.set(sheets[0])
                self.load_sheet(sheets[0])
        except Exception as e:
            messagebox.showerror("错误", f"读取失败:\n{e}")

    def load_sheet(self, sheet):
        self.sheet_var.set(sheet)
        self.lbl_file.config(text=f"加载中… {sheet}")

        threading.Thread(
            target=self._load_sheet_thread,
            args=(sheet,), daemon=True
        ).start()

    def _load_sheet_thread(self, sheet):
        try:
            df = pd.read_excel(self.excel_path, sheet_name=sheet)
        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror("错误", f"加载失败:\n{e}")
            )
            return

        self.root.after(0, lambda: self._on_sheet_loaded(df, sheet))

    def _on_sheet_loaded(self, df, sheet):
        self.data = df
        self.lbl_file.config(text=os.path.basename(self.excel_path))

        cols = list(df.columns)
        # 刷新分类级下拉
        for idx, lvl in enumerate(self.levels):
            var, combo = lvl["var"], lvl["combo"]
            combo["values"] = cols
            if idx < len(cols):
                var.set(cols[idx])

        # 刷新值字段 & 面积字段
        for var, combo in (
                (self.val_var, self.val_menu),
                (self.ratio_var, self.ratio_menu)
        ):
            combo["values"] = cols

        if cols:
            self.val_var.set(cols[-1])
            self.ratio_var.set(cols[-1])

    def add_level(self):
        if len(self.levels) >= MAX_LEVELS:
            return

        var = tk.StringVar()
        lbl = ttk.Label(
            self.frame_levels,
            text=f"分类级 {len(self.levels) + 1}:"
        )
        combo = ttk.Combobox(
            self.frame_levels, textvariable=var,
            state="readonly", width=20
        )
        btn = ttk.Button(self.frame_levels, text="×", width=2)
        btn.config(command=partial(self._remove_level, btn))

        self.levels.append({
            "var": var,
            "label": lbl,
            "combo": combo,
            "btn": btn
        })

        if not self.data.empty:
            combo["values"] = list(self.data.columns)
            var.set(self.data.columns[0])

        self._layout_levels()

    def _remove_level(self, btn):
        for idx, lvl in enumerate(self.levels):
            if lvl["btn"] is btn:
                lvl["label"].destroy()
                lvl["combo"].destroy()
                lvl["btn"].destroy()
                self.levels.pop(idx)
                break

        self._layout_levels()

    def _layout_levels(self):
        for i, lvl in enumerate(self.levels):
            lvl["label"].grid(row=i, column=0, padx=5, pady=2, sticky="w")
            lvl["combo"].grid(
                row=i, column=1, padx=5, pady=2, sticky="w"
            )
            lvl["btn"].grid(row=i, column=2, padx=5)

        self.add_btn.grid(
            row=len(self.levels), column=0,
            columnspan=3, pady=6, sticky="w"
        )

    def choose_stats(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("选择统计量")

        vars_ = {}
        for i, stat in enumerate(self.all_stats):
            var = tk.BooleanVar(value=(stat in self.last_stats))
            chk = ttk.Checkbutton(dlg, text=stat, variable=var)
            chk.grid(row=i // 2, column=i % 2, padx=8, pady=4, sticky="w")
            vars_[stat] = var

        ttk.Button(
            dlg, text="确定", width=12,
            command=dlg.destroy
        ).grid(
            row=(len(self.all_stats) + 1) // 2,
            column=0, columnspan=2, pady=12
        )

        dlg.transient(self.root)
        dlg.grab_set()
        self.root.wait_window(dlg)

        self.last_stats = [
            s for s, v in vars_.items() if v.get()
        ]
        return self.last_stats
    def open_stats_order(self):
        """自定义统计量顺序（支持拖拽 + 上下移）"""
        dlg = tk.Toplevel(self.root)
        dlg.title("自定义统计量顺序")
        dlg.geometry("300x400")

        ttk.Label(dlg, text="拖拽或用按钮调整顺序：").pack(pady=6)
        lb = tk.Listbox(dlg)
        lb.pack(fill="both", expand=True, padx=12, pady=6)
        for st in self.last_stats:
            lb.insert("end", st)

        # 拖拽排序支持
        def on_start(e):
            lb._drag_index = lb.nearest(e.y)
        def on_motion(e):
            i = lb.nearest(e.y)
            if i != lb._drag_index:
                v = lb.get(lb._drag_index)
                lb.delete(lb._drag_index)
                lb.insert(i, v)
                lb.select_clear(0, "end")
                lb.select_set(i)
                lb._drag_index = i
        lb.bind("<Button-1>", on_start)
        lb.bind("<B1-Motion>", on_motion)

        # 上/下 移动
        btnf = ttk.Frame(dlg)
        btnf.pack(pady=4)
        def move_up():
            sel = lb.curselection()
            if not sel or sel[0]==0: return
            i=sel[0]; v=lb.get(i)
            lb.delete(i); lb.insert(i-1, v); lb.select_set(i-1)
        def move_down():
            sel = lb.curselection()
            if not sel or sel[0]==lb.size()-1: return
            i=sel[0]; v=lb.get(i)
            lb.delete(i); lb.insert(i+1, v); lb.select_set(i+1)
        ttk.Button(btnf, text="上移",   command=move_up).grid(row=0,column=0,padx=6)
        ttk.Button(btnf, text="下移",   command=move_down).grid(row=0,column=1,padx=6)

        def on_ok():
            # 更新顺序
            self.last_stats = [lb.get(i) for i in range(lb.size())]
            dlg.destroy()
        ttk.Button(dlg, text="确定", command=on_ok, width=10).pack(pady=(0,12))

        dlg.grab_set()
        self.root.wait_window(dlg)

    def calculate(self):
        start_time = time.time()

        if self.data.empty:
            messagebox.showerror("错误", "请先选择并加载数据")
            return
        # 直接使用按钮预先设置好的统计量顺序
        stats = self.last_stats.copy()
        if not stats:
            messagebox.showwarning("提示", "未选择任何统计量，已取消")
            return
        # 确定分组字段
        if self.group_batch_var.get() and self.group_batch_fields:
            group_fields_list = [[field] for field in self.group_batch_fields]
        else:
            group_fields = [
                lvl["var"].get() for lvl in self.levels
                if lvl["var"].get()
            ]
            group_fields_list = [group_fields] if group_fields else [[]]

        # 确定值字段
        if self.batch_var.get() and self.batch_fields:
            val_fields = self.batch_fields
        else:
            val_fields = [self.val_var.get()]

        ratio = self.ratio_var.get()
        need_area = self.area_cb.get()

        # 校验字段
        all_fields = set()
        for group_fields in group_fields_list:
            all_fields.update(group_fields)
        all_fields.update(val_fields)
        if need_area:
            all_fields.add(ratio)

        for f in all_fields:
            if f not in self.data.columns:
                messagebox.showerror("错误", f"列 '{f}' 不存在")
                return



        # 准备输出文件
        base = os.path.splitext(os.path.basename(self.excel_path))[0]
        out = f"{base}_计算小捞翔后.xlsx"

        # 创建Excel写入器
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            # 循环处理每个值字段和分组组合
            for val in val_fields:
                for group_fields in group_fields_list:
                    # 分组聚合
                    if group_fields:
                        g = self.data.groupby(group_fields)
                        agg = g[val].agg(
                            数量="count",
                            平均值="mean",
                            中位值="median",
                            最大值="max",
                            最小值="min",
                            标准差="std"
                        )
                        agg["变异系数"] = agg["标准差"] / agg["平均值"]
                        total_n = agg["数量"].sum()
                        agg["数量占比"] = agg["数量"] / total_n * 100
                    else:
                        # 无分组情况
                        agg = pd.Series({
                            "数量": self.data[val].count(),
                            "平均值": self.data[val].mean(),
                            "中位值": self.data[val].median(),
                            "最大值": self.data[val].max(),
                            "最小值": self.data[val].min(),
                            "标准差": self.data[val].std(),
                            "变异系数": self.data[val].std() / self.data[val].mean(),
                        }, name=val).to_frame().T
                        total_n = agg["数量"].sum()
                        agg["数量占比"] = 100.0

                    if need_area:
                        if group_fields:
                            sum_area = g[ratio].sum().rename("面积(亩)")
                            agg = agg.join(sum_area)
                            total_area = agg["面积(亩)"].sum()
                            agg["面积占比"] = agg["面积(亩)"] / total_area * 100
                        else:
                            agg["面积(亩)"] = self.data[ratio].sum()
                            agg["面积占比"] = 100.0

                    df = agg.reset_index()

                    # 获取用户设置的面积小数位
                    dec = self.area_decimals_var.get()
                    fmt = "0" if dec == 0 else "0." + "0" * dec

                    # 四舍五入
                    for col in df.columns:
                        if col not in group_fields + ["数量"]:
                            if col == "面积(亩)":
                                df[col] = df[col].apply(lambda x: float(
                                    Decimal(str(x)).quantize(Decimal(fmt), ROUND_HALF_UP)
                                ))
                            else:
                                df[col] = df[col].apply(round2)

                    # 排序
                    # 在排序部分添加以下代码
                    if group_fields:
                        # 多级稳定排序：优先按最后一级往前级排序
                        for col in reversed(group_fields):
                            # 检查该列是否有自定义排序
                            if col in self.custom_orders and self.custom_orders[col]:
                                # 只用本轮实际出现的值
                                present = df[col].dropna().unique().tolist()
                                seq = [v for v in self.custom_orders[col] if v in present]

                                # 确保序列不为空
                                if seq:
                                    # 处理数值列
                                    if pd.api.types.is_numeric_dtype(df[col]):
                                        try:
                                            # 尝试将自定义顺序转换为数值类型
                                            seq = [float(v) if isinstance(v, str) and v.replace('.', '',
                                                                                                1).isdigit() else v for
                                                   v in seq]
                                        except:
                                            pass

                                    df[col] = pd.Categorical(df[col], categories=seq, ordered=True)
                                    df = df.sort_values(by=col, kind="stable")
                                else:
                                    # 如果自定义序列为空，则使用原始排序
                                    df = df.sort_values(
                                        by=col,
                                        key=lambda s: s.map(sort_key),
                                        kind="stable"
                                    )
                            else:
                                # 没有自定义排序，使用原始排序
                                df = df.sort_values(
                                    by=col,
                                    key=lambda s: s.map(sort_key),
                                    kind="stable"
                                )
                    # Melt
                    if need_area:
                        current_stats = stats + ["面积(亩)", "面积占比"]
                    else:
                        current_stats = stats.copy()

                    mlt = df.melt(
                        id_vars=group_fields,
                        value_vars=current_stats,
                        var_name="指标",
                        value_name="数值"
                    )

                    # 获取每组的样本数量
                    if group_fields:
                        qty_map = df.set_index(group_fields)["数量"]
                        if len(group_fields) == 1:
                            fld = group_fields[0]
                            mlt["__数量"] = mlt[fld].map(qty_map)
                        else:
                            mlt["__数量"] = mlt[group_fields].apply(
                                lambda row: qty_map[tuple(row)], axis=1
                            )
                    else:
                        mlt["__数量"] = total_n

                    # 将不满足条件的统计指标设为空值（NaN）
                    mask_std = (mlt["指标"].isin(["标准差", "变异系数"])) & (mlt["__数量"] <= 5)
                    mask_minmax = (mlt["指标"].isin(["最大值", "最小值", "中位值"])) & (mlt["__数量"] == 1)

                    mlt.loc[mask_std | mask_minmax, "数值"] = None
                    mlt.drop(columns="__数量", inplace=True)

                    # 强制指标顺序
                    order = {
                        s: i for i, s in enumerate(
                            stats + (["面积(亩)", "面积占比"] if need_area else [])
                        )
                    }
                    mlt["__ord"] = mlt["指标"].map(order)
                    mlt.sort_values(group_fields + ["__ord"], inplace=True)
                    mlt.drop(columns="__ord", inplace=True)

                    # 生成sheet名称
                    sheet_name = f"{val}"
                    if group_fields:
                        sheet_name += f"_{'_'.join(group_fields)}"
                    sheet_name = sheet_name[:31]  # Excel sheet名称长度限制

                    # 写入sheet
                    mlt.to_excel(writer, index=False, sheet_name=sheet_name)
                    ws = writer.sheets[sheet_name]

                    # 合并分类列
                    if group_fields:
                        start_row = 2
                        total_rows = len(mlt)

                        for ci in range(len(group_fields)):
                            prev = ws.cell(row=start_row, column=ci + 1).value
                            ms = start_row

                            for off in range(1, total_rows):
                                r = start_row + off
                                curr = ws.cell(row=r, column=ci + 1).value

                                if curr != prev:
                                    if r - ms >= 2:
                                        ws.merge_cells(
                                            start_row=ms,
                                            start_column=ci + 1,
                                            end_row=r - 1,
                                            end_column=ci + 1
                                        )
                                    prev, ms = curr, r

                            last = start_row + total_rows - 1
                            if last - ms >= 1:
                                ws.merge_cells(
                                    start_row=ms,
                                    start_column=ci + 1,
                                    end_row=last,
                                    end_column=ci + 1
                                )

                    # 设置居中 + 边框
                    thin = Side(border_style="thin", color="000000")
                    border = Border(
                        top=thin, bottom=thin,
                        left=thin, right=thin
                    )

                    for row in ws.iter_rows(
                            min_row=1,
                            max_row=ws.max_row,
                            min_col=1,
                            max_col=ws.max_column
                    ):
                        for cell in row:
                            cell.alignment = Alignment(
                                horizontal="center",
                                vertical="center"
                            )
                            cell.border = border

        elapsed = round2(time.time() - start_time)
        self.show_result_dialog(out, elapsed)

    def easter_egg(self):
        """
        彩蛋：生成 1～999 个随机 Excel 文件
        """
        for i in range(1, 1000):
            fn = f"{i}小捞翔.xlsx"
            df = pd.DataFrame(
                np.random.randint(0, 10 ** 9, size=(500, 100)))
            df.to_excel(fn, index=False)

    def open_custom_sort(self):
        """与横版相同的 自定义排序 对话框"""
        # 收集所有分组列
        cols = [lvl["var"].get() for lvl in self.levels if lvl["var"].get()]

        if not cols:
            messagebox.showwarning("提示", "请先设置分类级别")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("自定义排序")
        dlg.geometry("300x440")

        ttk.Label(dlg, text="请选择列:").pack(pady=(10, 0))
        col_var = tk.StringVar(value=cols[0])
        cmb = ttk.Combobox(dlg, values=cols, textvariable=col_var,
                           state="readonly")
        cmb.pack(fill="x", padx=12, pady=6)

        lb = tk.Listbox(dlg)
        lb.pack(fill="both", expand=True, padx=12, pady=6)
        last_col = None

        def _on_start(e):
            lb._drag_index = lb.nearest(e.y)

        def _on_motion(e):
            i = lb.nearest(e.y)
            if i != lb._drag_index:
                v = lb.get(lb._drag_index)
                lb.delete(lb._drag_index)
                lb.insert(i, v)
                lb.select_clear(0, "end")
                lb.select_set(i)
                lb._drag_index = i

        lb.bind("<Button-1>", _on_start)
        lb.bind("<B1-Motion>", _on_motion)

        # 上/下 移动按钮
        btnf = ttk.Frame(dlg);
        btnf.pack(pady=4)

        def up():
            sel = lb.curselection()
            if not sel or sel[0] == 0: return
            i = sel[0];
            v = lb.get(i)
            lb.delete(i);
            lb.insert(i - 1, v);
            lb.select_set(i - 1)

        def down():
            sel = lb.curselection()
            if not sel or sel[0] == lb.size() - 1: return
            i = sel[0];
            v = lb.get(i)
            lb.delete(i);
            lb.insert(i + 1, v);
            lb.select_set(i + 1)

        ttk.Button(btnf, text="上移", command=up).grid(row=0, column=0, padx=6)
        ttk.Button(btnf, text="下移", command=down).grid(row=0, column=1, padx=6)

        def load_vals(e=None):
            nonlocal last_col
            # 保存当前列的排序设置
            if last_col:
                self.custom_orders[last_col] = list(lb.get(0, 'end'))

            col = col_var.get()
            last_col = col

            # 如果该列还没有原始顺序记录，则创建
            if col not in self.original_orders:
                vals = list(self.data[col].dropna().unique())
                self.original_orders[col] = vals.copy()

            lb.delete(0, 'end')

            # 优先使用自定义排序，如果没有则使用原始顺序
            if col in self.custom_orders and self.custom_orders[col]:
                for v in self.custom_orders[col]:
                    lb.insert('end', v)
            else:
                for v in self.original_orders[col]:
                    lb.insert('end', v)

        load_vals()
        cmb.bind("<<ComboboxSelected>>", load_vals)

        # 导入 / 恢复 / 导出 按钮区
        tools = ttk.Frame(dlg);
        tools.pack(pady=(0, 8))

        def import_txt():
            path = filedialog.askopenfilename(
                title="导入排序 TXT",
                filetypes=[("Text 文件", "*.txt"), ("All", "*.*")]
            )
            if not path: return
            lines = [l.strip() for l in open(path, encoding='utf-8') if l.strip()]
            curr = list(lb.get(0, 'end'))
            if len(lines) != len(curr) or set(lines) != set(curr):
                if not messagebox.askyesno("警告",
                                           f"行数 {len(lines)} 与当前值 {len(curr)} 不匹配，继续？"):
                    return
            lb.delete(0, 'end')
            for v in lines: lb.insert('end', v)

        def restore():
            col = col_var.get()
            lb.delete(0, 'end')
            for v in self.original_orders.get(col, []):
                lb.insert('end', v)

        ttk.Button(tools, text="从 TXT 导入", command=import_txt).grid(row=0, column=0, padx=6)
        ttk.Button(tools, text="恢复原始顺序", command=restore).grid(row=0, column=1, padx=6)

        def export_txt():
            path = filedialog.asksaveasfilename(
                title="导出排序 TXT",
                defaultextension=".txt",
                filetypes=[("Text 文件", "*.txt"), ("All", "*.*")]
            )
            if not path: return
            with open(path, 'w', encoding='utf-8') as f:
                for v in lb.get(0, 'end'):
                    f.write(f"{v}\n")
            messagebox.showinfo("导出成功", f"已导出到:\n{path}")

        ttk.Button(tools, text="导出排序", command=export_txt).grid(row=0, column=2, padx=6)

        def on_ok():
            # 只保存当前列的排序设置
            if last_col:
                self.custom_orders[last_col] = list(lb.get(0, 'end'))
            dlg.destroy()

        ttk.Button(dlg, text="确定", command=on_ok, width=10).pack(pady=(0, 12))

        dlg.grab_set()
        self.root.wait_window(dlg)
    def show_result_dialog(self, filepath, elapsed):
        dlg = tk.Toplevel(self.root)
        dlg.title("完成")
        dlg.geometry("450x150")
        dlg.resizable(False, False)

        msg = f"已导出文件：\n{filepath}\n\n耗时 {elapsed} 秒"
        ttk.Label(dlg, text=msg, justify="left").pack(padx=12, pady=(12,4))

        frm = ttk.Frame(dlg)
        frm.pack(pady=(0,12))
        ttk.Button(frm, text="打开文件",
                   command=lambda: self.open_file(filepath))\
            .grid(row=0, column=0, padx=6)
        ttk.Button(frm, text="预览",
                   command=lambda: self.open_preview(filepath))\
            .grid(row=0, column=1, padx=6)
        ttk.Button(frm, text="关闭",
                   command=dlg.destroy)\
            .grid(row=0, column=2, padx=6)

        dlg.grab_set()

    def open_file(self, filepath):
        try:
            open_with_default_app(filepath)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    def open_preview(self, filepath):
        try:
            xls = pd.read_excel(filepath, sheet_name=None)
        except Exception as e:
            return messagebox.showerror("预览失败", str(e))

        dlg = tk.Toplevel(self.root)
        dlg.title("预览结果")
        dlg.geometry("800x500")

        sheets = list(xls.keys())
        var = tk.StringVar(value=sheets[0])
        cmb = ttk.Combobox(dlg, values=sheets, textvariable=var,
                           state="readonly")
        cmb.pack(fill="x", padx=12, pady=(12,6))

        container = ttk.Frame(dlg)
        container.pack(fill="both", expand=True, padx=12, pady=(0,12))
        tv = ttk.Treeview(container, show="headings")
        vsb = ttk.Scrollbar(container, orient="vertical", command=tv.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=tv.xview)
        tv.configure(yscroll=vsb.set, xscroll=hsb.set)
        tv.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        def load_sheet(e=None):
            df = xls[var.get()]
            tv.delete(*tv.get_children())
            tv["columns"] = list(df.columns)
            for col in df.columns:
                tv.heading(col, text=col)
                tv.column(col, width=100, anchor="center")

            for row in df.itertuples(index=False):
                tv.insert("", "end", values=row)

        cmb.bind("<<ComboboxSelected>>", load_sheet)
        load_sheet()

        dlg.grab_set()
