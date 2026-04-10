import itertools
import json
import os
import threading
import time
import traceback
from decimal import Decimal, ROUND_HALF_UP

import numpy as np
import pandas as pd
import sv_ttk
import tkinter as tk
from openpyxl.styles import Alignment, Border, Side
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES

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


MAX_LEVELS = 3
# 默认配置文件名（可自定义目录）
CONFIG_FILE = "small_shit.json"

# 1) 定义一个可折叠面板类

class HorizontalApp:
    def __init__(self, root):
        self.root = root
        root.title("小捞翔·至尊版")
        root.geometry("570x550")
        root.resizable(True, True)

        # 在窗口关闭前保存配置
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)


        sv_ttk.set_theme("light")
        menubar = tk.Menu(root)
        tm = tk.Menu(menubar, tearoff=0)
        tm.add_command(label="Light", command=lambda: sv_ttk.set_theme("light"))
        tm.add_command(label="Dark",  command=lambda: sv_ttk.set_theme("dark"))
        menubar.add_cascade(label="Theme", menu=tm)
        # 关于 → 弹出“更新历史内容”
        menubar.add_command(label="关于", command=self.show_update_history)

        root.config(menu=menubar)

        root.drop_target_register(DND_FILES)
        root.dnd_bind('<<DragEnter>>',  lambda e: root.configure(bg='#e3f2fd'))
        root.dnd_bind('<<DragLeave>>',  lambda e: root.configure(bg='SystemButtonFace'))
        root.dnd_bind('<<Drop>>',       self.handle_drop)

        def sep(): ttk.Separator(root, orient='horizontal').pack(fill='x', pady=6)

        self.excel_path = ""
        # … 其他属性初始化 …
        self.recent_files = [] 
        self.custom_orders = {}    # 用户确认后的自定义顺序
        self.original_orders = {}
        self.data = pd.DataFrame()
        self.all_stats = ["数量","平均值","中位值","最大值","最小值",
                          "标准差","变异系数","数量占比","合计"]
        self.last_stats = self.all_stats.copy()
        self.batch_fields = []
        self.group_batch_fields = []

        # 文件 & Sheet 选择
        frm = ttk.Frame(root); frm.pack(fill='x', padx=12, pady=6)
        ttk.Button(frm, text="选择 Excel", command=self.select_file)\
            .grid(row=0, column=0, padx=(0,8))
        self.lbl_file = ttk.Label(frm, text="— 未选择文件 —")
        self.lbl_file.grid(row=0, column=1, sticky='w')
        ttk.Label(frm, text="子表:").grid(row=1, column=0, pady=6)
        self.sheet_var = tk.StringVar()
        self.sheet_menu = ttk.Combobox(frm, textvariable=self.sheet_var,
                                       state="readonly", width=38)
        self.sheet_menu.grid(row=1, column=1, sticky='w')
        self.sheet_menu.bind("<<ComboboxSelected>>",
                             lambda e: self.load_sheet(self.sheet_var.get()))
        sep()
        # “选择 Excel”旁边的“最近打开”下拉
        self.recent_btn = ttk.Menubutton(frm, text="最近打开")
        self.recent_menu = tk.Menu(self.recent_btn, tearoff=0)
        self.recent_btn["menu"] = self.recent_menu
        self.recent_btn.grid(row=0, column=2, padx=(8, 0))

        # 分类级别(1~3级)
        self.frame_levels = ttk.LabelFrame(root, text="分类级别(1~3级)")
        self.frame_levels.pack(fill='x', padx=12, pady=6)

        self.levels = []
        self.add_btn = ttk.Button(self.frame_levels, text="+ 添加级别",
                                  command=self.add_level)
        self.add_level(); self.add_level()
        sep()


        # 值字段 & 面积占比
        frm2 = ttk.Frame(root)
        frm2.pack(fill='x', padx=12, pady=6)

        ttk.Label(frm2, text="值字段:").grid(row=0, column=0)
        self.val_var = tk.StringVar()
        self.val_menu = ttk.Combobox(frm2, textvariable=self.val_var,
                                     state="readonly", width=36)
        self.val_menu.grid(row=0, column=1, sticky='w')
        self.val_enable_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm2, text="启用值字段",
                        variable=self.val_enable_var,
                        command=self.toggle_value_field) \
            .grid(row=0, column=2, padx=6)
        ttk.Label(frm2, text="面积字段:").grid(row=1, column=0, pady=6)
        self.ratio_var = tk.StringVar()
        self.ratio_menu = ttk.Combobox(frm2, textvariable=self.ratio_var,
                                       state="readonly", width=36)
        self.ratio_menu.grid(row=1, column=1, sticky='w')
        self.area_cb = tk.BooleanVar()
        # … 已有的面积字段 & 复选框 …
        ttk.Checkbutton(frm2, text="计算面积占比", variable=self.area_cb,
                        command=self.toggle_area_fields) \
            .grid(row=1, column=2, padx=6)

        # 新增：面积小数位选择
        ttk.Label(frm2, text="面积小数位:").grid(row=2, column=0, pady=(4, 0))
        self.area_decimals_var = tk.IntVar(value=2)
        self.area_decimals_spin = ttk.Spinbox(
            frm2, from_=0, to=10, textvariable=self.area_decimals_var,
            width=5, state="readonly"
        )
        self.area_decimals_spin.grid(row=2, column=1, sticky='w')

        sep()

        # 统计量 & 批量 & 自定义排序

        # 一定要先定义这两个属性
        self.batch_var       = tk.BooleanVar()
        self.group_batch_var = tk.BooleanVar()

        # 统计量 & 批量 & 自定义排序 分两行布局
        stats_frm = ttk.Frame(root)
        stats_frm.pack(fill='x', padx=12, pady=6)

        # 第一行：统计量
        row1 = ttk.Frame(stats_frm)
        row1.pack(fill='x', pady=2)
        ttk.Button(row1, text="选择统计量", command=self.choose_stats, width=16)\
            .pack(side='left', padx=6)
        ttk.Button(row1, text="自定义统计量顺序", command=self.open_stats_order, width=16)\
            .pack(side='left', padx=6)
        ttk.Button(row1, text="自定义排序", command=self.open_custom_sort, width=10)\
            .pack(side='left', padx=6)
        # 第二行：批量 & 分组 & 自定义排序
        row2 = ttk.Frame(stats_frm)
        row2.pack(fill='x', pady=2)
        ttk.Checkbutton(
            row2, text="批量值字段", variable=self.batch_var,
            command=self.toggle_batch
        ).pack(side='left', padx=6)
        self.batch_btn = ttk.Button(
            row2, text="选择值字段",
            command=self.choose_batch_fields,
            state="disabled", width=10
        )
        self.batch_btn.pack(side='left', padx=6)

        ttk.Checkbutton(
            row2, text="批量分组(级别1)",
            variable=self.group_batch_var,
            command=self.toggle_group_batch
        ).pack(side='left', padx=6)
        self.group_batch_btn = ttk.Button(
            row2, text="选择分组字段",
            command=self.choose_group_batch_fields,
            state="disabled", width=12
        )
        self.group_batch_btn.pack(side='left', padx=6)
        sep()

        # 操作按钮
        btnf = ttk.Frame(root); btnf.pack(pady=8)
        self.calculate_btn = ttk.Button(btnf, text="计算并导出",
                                        command=self.calculate, width=16)
        self.calculate_btn.grid(row=0, column=0, padx=6)
        ttk.Button(btnf, text="彩蛋", command=self.easter_egg,
                   width=8).grid(row=0, column=1)
        sep()

        # 进度条 & 状态栏
        self.progress = ttk.Progressbar(root, mode="indeterminate")
        self.progress.pack(fill='x', padx=12); self.progress.pack_forget()
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(root, textvariable=self.status_var, relief='sunken',
                  anchor='w').pack(side='bottom', fill='x')

        self.toggle_area_fields()
        # 载入上次保存的配置
        self.load_config()
        # 确保存在默认结构
        self.advanced_order = getattr(self, "advanced_order",
                                      {"groups": [], "stats": [], "exports": []})

        self.create_recent_menu()
        self.check_for_update()
    def show_update_history(self):
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
        # 一条一条字体换行
        txt.insert("1.0", "\n\n".join(history))
        txt.configure(state="disabled")
        txt.pack(fill="both", expand=True)

        btn = ttk.Button(frm, text="关闭", command=dlg.destroy)
        btn.pack(pady=6)

        dlg.transient(self.root)
        dlg.grab_set()
        self.root.wait_window(dlg)

    def check_for_update(self):
        if getattr(self, "shown_version", "") != APP_VERSION:
            dlg = tk.Toplevel(self.root)
            dlg.title(f"更新日志  v{APP_VERSION}")
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
            # 把按钮靠右放，有内边距
            btn.pack(side="right", padx=4, pady=4)

            dlg.transient(self.root)
            dlg.grab_set()
            self.root.wait_window(dlg)

    # 批量值字段
    # 自定义排序对话框：选择分组列，拖拽/上下移动唯一值
    def save_config(self, path=None, show_msg=True):

        cfg = {
            "custom_orders": self.custom_orders,
            "last_stats": self.last_stats,
            "recent_files": self.recent_files,
            "shown_version": getattr(self, "shown_version", ""),
            "update_history": getattr(self, "update_history", []),
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
        # 确保 update_history 属性总是存在
        self.update_history = []
        """
        从 JSON 配置文件加载：custom_orders, last_stats, recent_files,
        shown_version, update_history 以及 area_decimals（面积小数位）。
        """
        cfg_path = path or CONFIG_FILE
        # 如果配置文件不存在，直接返回（保持各属性的默认值）
        if not os.path.exists(cfg_path):
            return

        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)

            # 用户在 自定义排序 中保存的列顺序
            self.custom_orders = cfg.get("custom_orders", {})

            # 上次选择的统计量列表
            self.last_stats = cfg.get("last_stats", self.all_stats.copy())

            # 最近打开文件列表
            self.recent_files = cfg.get("recent_files", [])

            # 已展示过更新提示的版本号
            self.shown_version = cfg.get("shown_version", "")

            # 已累计的更新日志历史（多条）
            self.update_history = cfg.get("update_history", [])

            # 面积小数位：默认为已有控件的当前值
            dec = cfg.get("area_decimals", self.area_decimals_var.get())
            self.area_decimals_var.set(dec)

        except Exception as e:
            messagebox.showerror("加载失败", str(e))

    def on_closing(self):
        # 退出前保存配置
        self.save_config(show_msg=False)
        self.root.destroy()

    def open_stats_order(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("自定义统计量顺序")
        dlg.geometry("300x400")

        ttk.Label(dlg, text="拖拽或用按钮调整顺序：").pack(pady=6)

        lb = tk.Listbox(dlg)
        lb.pack(fill='both', expand=True, padx=12, pady=6)
        for stat in self.last_stats:
            lb.insert('end', stat)

        # 拖拽排序支持
        def _on_start(e):
            lb._drag_index = lb.nearest(e.y)

        def _on_motion(e):
            i = lb.nearest(e.y)
            if i != lb._drag_index:
                v = lb.get(lb._drag_index)
                lb.delete(lb._drag_index)
                lb.insert(i, v)
                lb.select_clear(0, 'end')
                lb.select_set(i)
                lb._drag_index = i

        lb.bind("<Button-1>", _on_start)
        lb.bind("<B1-Motion>", _on_motion)

        # 上下移动按钮
        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=4)


        def move_up():
            sel = lb.curselection()
            if not sel or sel[0] == 0: return
            i = sel[0];
            v = lb.get(i)
            lb.delete(i);
            lb.insert(i - 1, v);
            lb.select_set(i - 1)

        def move_down():
            sel = lb.curselection()
            if not sel or sel[0] == lb.size() - 1: return
            i = sel[0];
            v = lb.get(i)
            lb.delete(i);
            lb.insert(i + 1, v);
            lb.select_set(i + 1)

        ttk.Button(btn_frame, text="上移", command=move_up).grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="下移", command=move_down).grid(row=0, column=1, padx=6)

        # 确定后更新 self.last_stats
        def on_ok():
            self.last_stats = [lb.get(i) for i in range(lb.size())]
            dlg.destroy()

        ttk.Button(dlg, text="确定", command=on_ok, width=10) \
            .pack(pady=(0, 12))

        dlg.grab_set()
        self.root.wait_window(dlg)

    def open_custom_sort(self):
        # 先收集所有已选的分类级别列
        cols = [l["var"].get() for l in self.levels if l["var"].get()]

        # 如果批量分组开启，把它们也加入可选列
        if getattr(self, "group_batch_fields", None):
            for g in self.group_batch_fields:
                if g not in cols:
                    cols.append(g)

        if not cols:
            messagebox.showwarning("提示", "请先设置分类级别")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("自定义排序")
        dlg.geometry("300x440")

        ttk.Label(dlg, text="请选择列:").pack(pady=(10,0))
        col_var = tk.StringVar(value=cols[0])
        col_cmb = ttk.Combobox(dlg, values=cols, textvariable=col_var,
                               state="readonly")
        col_cmb.pack(fill='x', padx=12, pady=6)

        lb = tk.Listbox(dlg)
        lb.pack(fill='both', expand=True, padx=12, pady=6)
        last_col = None
        # 拖拽排序
        def _on_start_drag(e):
            lb._drag_index = lb.nearest(e.y)
        def _on_drag_motion(e):
            i = lb.nearest(e.y)
            if i != lb._drag_index:
                v = lb.get(lb._drag_index)
                lb.delete(lb._drag_index)
                lb.insert(i, v)
                lb.select_clear(0, 'end')
                lb.select_set(i)
                lb._drag_index = i

        lb.bind("<Button-1>", _on_start_drag)
        lb.bind("<B1-Motion>", _on_drag_motion)

        # 上/下移
        btnf = ttk.Frame(dlg); btnf.pack(pady=4)
        def up():
            sel = lb.curselection()
            if not sel or sel[0]==0: return
            i=sel[0];v=lb.get(i)
            lb.delete(i); lb.insert(i-1,v); lb.select_set(i-1)
        def down():
            sel = lb.curselection()
            if not sel or sel[0]==lb.size()-1: return
            i=sel[0];v=lb.get(i)
            lb.delete(i); lb.insert(i+1,v); lb.select_set(i+1)
        ttk.Button(btnf, text="上移",   command=up).grid(row=0,column=0,padx=6)
        ttk.Button(btnf, text="下移",   command=down).grid(row=0,column=1,padx=6)

        # 加载列表初始值
        def load_vals(e=None):
            nonlocal last_col
            if last_col:
                self.custom_orders[last_col] = list(lb.get(0, 'end'))
            col = col_var.get()
            # 记录原始顺序（只第一次记录）
            if col not in self.original_orders:
                vals = list(self.data[col].dropna().unique())
                self.original_orders[col] = vals.copy()
            # 先清空
            lb.delete(0, 'end')

            seq = self.custom_orders.get(col, self.original_orders[col])
            for v in seq:
                lb.insert('end', v)
                # 切换完成后，将本次列名记为 last_col
            last_col = col

        load_vals()
        col_cmb.bind("<<ComboboxSelected>>", load_vals)

        # 导入 TXT（保持你已有逻辑）
        def import_txt():
            path = filedialog.askopenfilename(
                title="导入排序 TXT",
                filetypes=[("Text 文件","*.txt"),("All","*.*")]
            )
            if not path: return
            with open(path, encoding='utf-8') as f:
                lines = [l.strip() for l in f if l.strip()]
            curr = list(lb.get(0,'end'))
            if len(lines)!=len(curr) or set(lines)!=set(curr):
                cont = messagebox.askyesno(
                    "导入警告",
                    f"TXT 行数 {len(lines)} 与当前唯一值 {len(curr)} 不匹配，是否继续导入？"
                )
                if not cont: return
            lb.delete(0,'end')
            for v in lines: lb.insert('end', v)

        # 恢复原始顺序
        def restore_original():
            col = col_var.get()
            if col in self.original_orders:
                lb.delete(0,'end')
                for v in self.original_orders[col]:
                    lb.insert('end', v)

        # 确定
        def on_ok():
            col = col_var.get()
            raw = list(lb.get(0, 'end'))
            # —— 如果原始列是数值型，就把字符串转成对应数字
            if pd.api.types.is_numeric_dtype(self.data[col]):
                dtype = self.data[col].dtype
                def to_num(x):
                    try:
                        # 整数列用 int，浮点列用 float
                        return int(x) if np.issubdtype(dtype, np.integer) else float(x)
                    except:
                        return float(x)
                self.custom_orders[col] = [to_num(v) for v in raw]
            else:
                # 非数值型，保留字符串排序
                self.custom_orders[col] = raw
            dlg.destroy()

        btn_tools = ttk.Frame(dlg); btn_tools.pack(pady=(0,8))
        ttk.Button(btn_tools, text="从 TXT 导入", command=import_txt)\
            .grid(row=0, column=0, padx=6)
        ttk.Button(btn_tools, text="恢复原始顺序",
                   command=restore_original)\
            .grid(row=0, column=1, padx=6)

        # 定义导出函数
        def export_txt():
            path = filedialog.asksaveasfilename(
                title="导出排序 TXT",
                defaultextension=".txt",
                filetypes=[("Text 文件", "*.txt"), ("All", "*.*")]
            )
            if not path:
                return

            # 读取 Listbox 中的当前顺序
            lines = [lb.get(i) for i in range(lb.size())]

            try:
                with open(path, 'w', encoding='utf-8') as f:
                    for v in lines:
                        f.write(f"{v}\n")
                messagebox.showinfo("导出成功", f"已将当前排序导出到：\n{path}")
            except Exception as e:
                messagebox.showerror("导出失败", str(e))

        # 在工具按钮区添加 “导出排序” 按钮
        ttk.Button(btn_tools, text="导出排序", command=export_txt) \
            .grid(row=0, column=2, padx=6)

        ttk.Button(dlg, text="确定", command=on_ok, width=10)\
            .pack(pady=(0,12))

        dlg.grab_set()
        self.root.wait_window(dlg)

    def toggle_batch(self):
        if self.batch_var.get():
            self.batch_btn.state(["!disabled"])
        else:
            self.batch_btn.state(["disabled"])
            self.batch_fields = []

    def choose_batch_fields(self):
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

    # 批量分组字段(分类级别1)
    def toggle_group_batch(self):
        if self.group_batch_var.get():
            self.group_batch_btn.state(["!disabled"])
        else:
            self.group_batch_btn.state(["disabled"])
            self.group_batch_fields = []

    def choose_group_batch_fields(self):
        if self.data.empty:
            messagebox.showwarning("提示", "请先加载数据")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("选择分组字段(多选) —— 作为 分类级别1")
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

    # 整表总计行
    def _append_overall_total(self, final, df, gf, val, ratio, need_area):
        stats_map = {
            "数量": df[val].count(),
            "平均值": round2(df[val].mean()),
            "中位值": round2(df[val].median()),
            "最大值": round2(df[val].max()),
            "最小值": round2(df[val].min()),
            "标准差": round2(df[val].std()),
            "变异系数": round2(df[val].std() / df[val].mean() if df[val].mean() else np.nan),
            "数量占比": round2(100.0),
        }
        if need_area:
            total_area = df[ratio].sum()

            dec = self.area_decimals_var.get()
            fmt = "0" if dec == 0 else "0." + "0" * dec
            stats_map["面积(亩)"] = float(Decimal(str(total_area))
                                          .quantize(Decimal(fmt), ROUND_HALF_UP))
            stats_map["面积占比"] = round2(100.0)

        total_row = {}
        for col in final.columns:
            if col in gf:
                total_row[col] = "总计" if col == gf[0] else ""
            else:
                total_row[col] = stats_map.get(col, "")

        overall_df = pd.DataFrame([total_row], columns=final.columns)
        return pd.concat([final, overall_df], ignore_index=True)

    # 文件 & 数据加载
    def toggle_area_fields(self):
        if self.area_cb.get():
            self.ratio_menu.state(["!disabled"])
            self.area_decimals_spin.state(["!disabled"])
        else:
            self.ratio_menu.state(["disabled"])
            self.ratio_var.set("")
            self.area_decimals_spin.state(["disabled"])

    def toggle_value_field(self):
        """控制值字段下拉和批量值字段复选框的启用/禁用"""
        if self.val_enable_var.get():
                self.val_menu.state(["!disabled"])
                # 保持 batch 的 checkbox 可操作
        else:
                # 禁用下拉，并清空已选
                self.val_menu.state(["disabled"])
                self.val_var.set("")

                self.batch_var.set(False)
                self.toggle_batch()
    def select_file(self):
        self._show_progress("选择文件中...")
        path = filedialog.askopenfilename(
            filetypes=[
                ("Excel 文件", ("*.xlsx", "*.xls")),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        if not path:
            return self._hide_progress("就绪")
        self.add_recent(path)
        self.excel_path = path
        self.lbl_file.config(text=os.path.basename(path))
        ext = os.path.splitext(path)[1].lower()

        if ext == ".csv":
            threading.Thread(target=self._read_csv, daemon=True).start()
        else:
            threading.Thread(target=self._get_sheets, daemon=True).start()

    def handle_drop(self, e):
        path = e.data.strip("{}")
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".xlsx", ".xls", ".csv"):
            return messagebox.showerror("错误", "只支持拖拽 .xlsx、.xls 或 .csv")
        self._show_progress("读取文件中...")
        self.add_recent(path)
        self.excel_path = path
        self.lbl_file.config(text=os.path.basename(path))

        if ext == ".csv":
            threading.Thread(target=self._read_csv, daemon=True).start()
        else:
            threading.Thread(target=self._get_sheets, daemon=True).start()

    def _get_sheets(self):
        try:
            sheets = pd.ExcelFile(self.excel_path).sheet_names
        except Exception as ex:
            self.root.after(0, lambda: messagebox.showerror("错误", f"{ex}"))
            return self._hide_progress("就绪")
        self.root.after(0, lambda: self._on_sheets(sheets))

    def _on_sheets(self, sheets):
        self._hide_progress()
        self.sheet_menu["values"] = sheets
        if sheets:
            self.sheet_var.set(sheets[0])
            self.load_sheet(sheets[0])

    def load_sheet(self, sheet):
        self.calculate_btn.state(["disabled"])
        self._show_progress(f"加载子表: {sheet}")
        threading.Thread(target=self._read_sheet, args=(sheet,), daemon=True).start()

    def _read_sheet(self, sheet):
        try:
            df = pd.read_excel(self.excel_path, sheet_name=sheet)
        except Exception as ex:
            self.root.after(0, lambda: messagebox.showerror("错误", f"{ex}"))
            return self._hide_progress("就绪")
        self.root.after(0, lambda: self._on_data(df))

    def _read_csv(self):
        try:
            df = read_csv_safely(self.excel_path)
        except Exception as ex:
            self.root.after(0, lambda: messagebox.showerror("错误", str(ex)))
            return self._hide_progress("就绪")
        self.root.after(0, lambda: self._on_data(df))

    def _on_data(self, df):
        self._hide_progress()
        self.data = df

        cols = list(df.columns)
        for cmb in (self.val_menu, self.ratio_menu):
            cmb["values"] = cols
        if cols:
            self.val_var.set(cols[-1])
            self.ratio_var.set(cols[-1])
        self._refresh_levels(cols)

        ext = os.path.splitext(self.excel_path)[1].lower()
        if ext == ".csv":
            self.sheet_menu.set("")
            self.sheet_menu["values"] = []
            self.sheet_menu.state(["disabled"])
        else:
            if "disabled" in self.sheet_menu.state():
                self.sheet_menu.state(["!disabled"])

        self.toggle_area_fields()
        self.calculate_btn.state(["!disabled"])

    def _refresh_levels(self, cols):
        for lvl in self.levels:
            c = lvl["combo"]
            c["values"] = cols
            if not c.get() or c.get() not in cols:
                c.set(cols[0] if cols else "")

    def add_level(self):
        if len(self.levels) >= MAX_LEVELS:
            return
        var = tk.StringVar()
        lbl = ttk.Label(self.frame_levels, text=f"分类级 {len(self.levels)+1}:")
        cmb = ttk.Combobox(
            self.frame_levels, textvariable=var,
            state="readonly", width=22
        )
        btn = ttk.Button(self.frame_levels, text="×", width=2)
        btn.config(command=lambda b=btn: self._remove_level(b))
        self.levels.append({"var":var, "label":lbl, "combo":cmb, "btn":btn})
        if not self.data.empty:
            cs = list(self.data.columns)
            cmb["values"] = cs
            cmb.set(cs[0])
        self._layout()

    def _remove_level(self, btn):
        for i, l in enumerate(self.levels):
            if l["btn"] is btn:
                l["label"].destroy()
                l["combo"].destroy()
                l["btn"].destroy()
                self.levels.pop(i)
                break
        self._layout()

    def _layout(self):
        for i, l in enumerate(self.levels):
            l["label"].grid(row=i, column=0, padx=6, pady=2, sticky='w')
            l["combo"].grid(row=i, column=1, padx=6, pady=2, sticky='w')
            l["btn"].grid(row=i, column=2, padx=6)
        self.add_btn.grid(row=len(self.levels), column=0, columnspan=3, pady=6)

    def choose_stats(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("选择统计量")
        dlg.geometry("300x260")
        frm = ttk.Frame(dlg)
        frm.pack(padx=12, pady=12)
        ttk.Label(frm, text="请选择统计量:").pack(pady=(0,8))
        chkf = ttk.Frame(frm)
        chkf.pack()
        # 用两个 dict 分别保存 BooleanVar 和对应的 Checkbutton
        tmp = {}  # {stat_name: BooleanVar}
        cbs = {}  # {stat_name: Checkbutton}
        for i, st in enumerate(self.all_stats):
            v = tk.BooleanVar(value=(st in self.last_stats))
            cb = ttk.Checkbutton(chkf, text=st, variable=v)
            cb.grid(row=i//3, column=i%3, padx=6, pady=4, sticky='w')
            tmp[st] = v
            cbs[st] = cb
            # ——— 如果只有一级分类，则自动取消“合计”的勾选（并可选禁用它） ———
        active_levels = sum(1 for lvl in self.levels if lvl['var'].get())
        if active_levels <= 1 and "合计" in tmp:
            # 取消勾选
            tmp["合计"].set(False)
            # 可选：禁用这个 Checkbutton，防止用户再勾上
            cbs["合计"].state(["disabled"])
        ttk.Button(frm, text="确定", command=dlg.destroy, width=10).pack(pady=8)
        dlg.grab_set()
        self.root.wait_window(dlg)
        self.last_stats = [s for s, v in tmp.items() if v.get()]

    def calculate(self):
        if not self.val_enable_var.get() and not self.area_cb.get() and not self.batch_var.get():
            messagebox.showwarning("提示",
                                    "请勾选“启用值字段”、“计算面积占比”或“批量值字段”后再导出"
            )
            return
        try:
            self._calculate()
        except Exception as ex:
            self.status_var.set(f"计算错误: {ex}")
            messagebox.showerror("错误", traceback.format_exc())


    def create_recent_menu(self):
        """根据 self.recent_files 构建下拉菜单。"""
        self.recent_menu.delete(0, 'end')
        for path in self.recent_files:
            label = os.path.basename(path)
            # 菜单项回调：打开并更新最近记录
            self.recent_menu.add_command(
                label=label,
                command=lambda p=path: self.open_recent(p)
            )
        if not self.recent_files:
            self.recent_menu.add_command(label="— 无记录 —", state="disabled")

    def add_recent(self, path):
        """将 path 插入到最近打开列表顶端，去重并限长20，静默保存。"""
        if path in self.recent_files:
            self.recent_files.remove(path)
        self.recent_files.insert(0, path)
        self.recent_files = self.recent_files[:20]
        self.save_config(show_msg=False)
        self.create_recent_menu()

    def open_recent(self, path):
        """通过菜单快速打开已有文件。"""
        if not os.path.exists(path):
            messagebox.showerror("文件不存在", path)
            self.recent_files.remove(path)
            self.save_config(show_msg=False)
            self.create_recent_menu()
            return

        self.excel_path = path
        self.lbl_file.config(text=os.path.basename(path))

        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            threading.Thread(target=self._read_csv, daemon=True).start()
        else:
            threading.Thread(target=self._get_sheets, daemon=True).start()

        # 更新最近打开顺序
        self.add_recent(path)

    def _calculate(self):
        start = time.time()
        self._show_progress("正在计算...")
        if self.data.empty:
            messagebox.showerror("错误", "请先加载数据")
            return self._hide_progress("就绪")

        # 默认分类级别
        gf = [l["var"].get() for l in self.levels if l["var"].get()]

        # 批量值字段
        fields = [self.val_var.get()]
        invalid = [f for f in fields
                   if f and not pd.api.types.is_numeric_dtype(self.data[f])]

        if invalid:
            # 弹警告，提示哪些字段有问题，然后中断
            messagebox.showwarning(
                "字段类型错误",
                f"以下值字段包含非数值或全部字符串，无法计算：\n{invalid}"
            )
            return self._hide_progress("就绪")

        # 正常开始计时和进度条
        start = time.time()
        self._show_progress("正在计算...")
        if self.batch_var.get() and getattr(self, "batch_fields", None):
            fields = self.batch_fields or fields

        # 批量分组字段
        groups = []
        if self.group_batch_var.get() and getattr(self, "group_batch_fields", None):
            groups = self.group_batch_fields

        # 先取到当前是否开启批量值字段 / 批量分组
        batch_on = self.batch_var.get()
        group_on = self.group_batch_var.get()

        # 取出 fields, groups
        fields = [self.val_var.get()]
        if batch_on and getattr(self, "batch_fields", None):
            fields = self.batch_fields or fields

        groups = []
        if group_on and getattr(self, "group_batch_fields", None):
            groups = self.group_batch_fields or []

        # --------- 这里替换原来的 loops 逻辑 ------------
        loops = []
        if batch_on and group_on:
            # 如果一一对应
            if len(fields) == len(groups):
                loops = [(f, [g]) for f, g in zip(fields, groups)]
            else:
                # 弹窗确认
                cont = messagebox.askyesno(
                    "提示",
                    "所选值字段和分组字段数量不一致，\n"
                    "是否按“所有值字段 × 所有分组字段”计算？"
                )
                if not cont:
                    # 用户点“否”，就直接退出
                    return self._hide_progress("就绪")
                # 真正的笛卡尔积：所有 (f,g) 对
                loops = [(f, [g]) for f, g in itertools.product(fields, groups)]

        elif batch_on:
            # 只有批量值字段
            loops = [(f, self.levels and [l["var"].get() for l in self.levels if l["var"].get()] or []) for f in fields]

        elif group_on:
            # 只有批量分组字段
            loops = [(self.val_var.get(), [g]) for g in groups]

        else:
            # 都没批量
            loops = [(self.val_var.get(),
                      [l["var"].get() for l in self.levels if l["var"].get()])]

        # 打个日志，确认 loops 里到底有哪些组合
        print("DEBUG loops:", loops)
        ratio = self.ratio_var.get()
        need_area = self.area_cb.get()
        need_sub = ("合计" in self.last_stats)

        basename = os.path.splitext(os.path.basename(self.excel_path))[0]

        out = f"{basename}_结果.xlsx"

        # 如果没有任何可统计的(field, group)组合，提前退出
        if not loops:
            messagebox.showerror("错误", "没有可统计的值/分组组合，无法导出")
            return self._hide_progress("就绪")

        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            for field, gf_current in loops:
                self.status_var.set(f"计算: 值字段={field}, 分组={gf_current}")

                df0 = self.data.copy()
                df0[field] = pd.to_numeric(df0[field], errors="coerce")

                # 仅在需要面积占比时才处理 ratio 列
                if need_area:
                    df0[ratio] = pd.to_numeric(df0[ratio], errors="coerce")
                    area_total = df0[ratio].sum()
                else:
                    area_total = None

                # 聚合
                if gf_current:
                    g = df0.groupby(gf_current)
                    agg = g[field].agg(
                        数量="count", 平均值="mean", 中位值="median",
                        最大值="max", 最小值="min", 标准差="std"
                    )
                else:
                    n = df0[field].count()
                    agg = pd.DataFrame({
                        "数量": [n],
                        "平均值": [df0[field].mean()],
                        "中位值": [df0[field].median()],
                        "最大值": [df0[field].max()],
                        "最小值": [df0[field].min()],
                        "标准差": [df0[field].std()],
                    })

                agg["变异系数"] = agg["标准差"] / agg["平均值"]
                agg["数量占比"] = agg["数量"] / agg["数量"].sum() * 100

                if need_area:
                    if gf_current:
                        agg["面积(亩)"] = g[ratio].sum()
                    else:
                        agg["面积(亩)"] = area_total
                    agg["面积占比"] = agg["面积(亩)"] / area_total * 100

                res = agg.reset_index()

                # 四舍五入 & 基本排序
                # 用户选的面积小数位
                dec = self.area_decimals_var.get()

                def round_dec(x, d):
                    try:
                        fmt = "0" if d == 0 else "0." + "0" * d
                        return float(Decimal(str(x)).quantize(Decimal(fmt), ROUND_HALF_UP))
                    except:
                        return x

                for c in res.select_dtypes(include="number").columns:
                    # 跳过分类列和 数量 列
                    if c in gf_current + ["数量"]:
                        continue

                    if c == "面积(亩)":
                        # 面积按用户设置的小数位
                        res[c] = res[c].apply(lambda x: round_dec(x, dec))
                    elif c == "面积占比":
                        # 面积占比固定两位
                        res[c] = res[c].apply(round2)
                    else:
                        # 其它列仍然两位
                        res[c] = res[c].apply(round2)

                # 阈值过滤
                filtered = res.copy()
                cnt_ser = filtered["数量"]
                if "标准差" in filtered.columns and "变异系数" in filtered.columns:
                    filtered.loc[cnt_ser <= 5, ["标准差", "变异系数"]] = np.nan
                for stat in ("最大值", "最小值", "中位值"):
                    if stat in filtered.columns:
                        filtered.loc[cnt_ser <= 1, stat] = np.nan

                # 列切片
                display_cols = gf_current + [s for s in self.last_stats if s != "合计"]
                if need_area:
                    display_cols += ["面积(亩)", "面积占比"]
                tmp = filtered[display_cols].copy()
                  # —— 1）先对 tmp 排序（仅对数据行，不含小计/总计） ——

                # —— 多级稳定排序 ——
                if gf_current:
                    for col in reversed(gf_current):
                        if col in self.custom_orders:
                            # 只保留本轮实际出现的顺序 —— 先转成 Python list 再判断
                            present_vals = tmp[col].dropna().unique().tolist()
                            uniq = [v for v in self.custom_orders[col]
                                    if v in present_vals]
                            tmp[col] = pd.Categorical(
                                tmp[col],
                                categories=uniq,
                                ordered=True
                            )
                            tmp = tmp.sort_values(by=col, kind="stable")
                        else:
                            tmp = tmp.sort_values(
                                by=col,
                                key=lambda s: s.map(sort_key),
                                kind="stable"
                            )

                # 小计逻辑：仅当 need_sub 且有分组时，并且该组行数 >1 才输出小计
                if need_sub and gf_current:
                    # 统计每个一级分组的行数
                    grp_counts = tmp[gf_current[0]].value_counts().to_dict()

                    out_rows = []
                    prev = None
                    for _, r in tmp.iterrows():
                        cur = r[gf_current[0]]
                        # 切组边界时，只有上一个组大小>1，才插小计
                        if prev is not None and cur != prev and grp_counts.get(prev, 0) > 1:
                            out_rows.append(
                                self._subtotal(
                                    prev, gf_current, field, ratio,
                                    df0[field].count(),
                                    area_total, need_area
                                )
                            )
                        out_rows.append(r.to_dict())
                        prev = cur
                    # 最后一组结束后，若该组大小>1，则插小计
                    if prev is not None and grp_counts.get(prev, 0) > 1:
                        out_rows.append(
                            self._subtotal(
                                prev, gf_current, field, ratio,
                                df0[field].count(),
                                area_total, need_area
                            )
                        )
                    final_df = pd.DataFrame(out_rows, columns=tmp.columns)
                else:
                    final_df = tmp.copy()

                # 整表总计行
                final_df = self._append_overall_total(
                    final_df, self.data, gf_current, field,
                    ratio, need_area
                )


                # 写入 Sheet
                # 新的 sheet 名：field + 所有分组，防止同名覆盖
                # gf_current 是一个 list，如果只有一级就只有一个元素
                name_parts = [field] + gf_current
                sheet_name = "_".join(name_parts)[:31]
                final_df.to_excel(writer, index=False, sheet_name=sheet_name)
                ws = writer.sheets[sheet_name]


                if gf_current:
                    start_row = 2
                    row_count = len(final_df)
                    thin = Side(border_style='thin', color='000000')
                    bd = Border(thin, thin, thin, thin)

                    # 1) 垂直合并：每一级分类列不跨“小计” nor “总计”
                    for ci in range(len(gf_current)):
                        merge_start = start_row
                        prev = ws.cell(merge_start, ci + 1).value

                        for r in range(start_row + 1, start_row + row_count):
                            curr = ws.cell(r, ci + 1).value

                            # 标记：如果 prev 是小计或总计，就一定切断
                            is_sub = isinstance(prev, str) and prev.endswith(" 合计")
                            is_total = (prev == "总计")

                            # 值变了 或者 刚好是小计/总计，都要切断上一段
                            if curr != prev or is_sub or is_total:
                                # 如果上一段有多行，并且上一段不是合计，就合并
                                if r - merge_start > 1 and not (is_sub or is_total):
                                    ws.merge_cells(
                                        start_row=merge_start,
                                        start_column=ci + 1,
                                        end_row=r - 1,
                                        end_column=ci + 1
                                    )
                                merge_start = r
                                prev = curr

                        # 合并最后一段，排除合计
                        last_span = (start_row + row_count - 1) - merge_start + 1
                        if last_span > 1:
                            if not (isinstance(prev, str) and prev.endswith(" 合计")) \
                                    and prev != "总计":
                                ws.merge_cells(
                                    start_row=merge_start,
                                    start_column=ci + 1,
                                    end_row=start_row + row_count - 1,
                                    end_column=ci + 1
                                )

                    # 2) 横向合并：所有“XXX 合计” 和 最后的“总计”
                    for r in range(start_row, start_row + row_count):
                        v = ws.cell(r, 1).value
                        if (isinstance(v, str) and v.endswith(" 合计")) or v == "总计":
                            ws.merge_cells(
                                start_row=r,
                                start_column=1,
                                end_row=r,
                                end_column=len(gf_current)
                            )

                # 3) 整表样式 & 数字格式
                # 3) 整表样式 & 数字格式 —— 按列名动态应用格式
                dec = self.area_decimals_var.get()
                area_fmt = "0" if dec == 0 else "0." + "0" * dec
                percent_fmt = "0.00"

                # 把 final_df.columns 映射到 Excel 列号
                col_idx_to_name = {i + 1: name for i, name in enumerate(final_df.columns)}

                for row in ws.iter_rows(
                        min_row=1, max_row=ws.max_row,
                        min_col=1, max_col=ws.max_column
                ):
                    for col_idx, cell in enumerate(row, start=1):
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = bd if "bd" in locals() else Border()

                        # 跳过第一行表头
                        if cell.row == 1:
                            continue

                        col_name = col_idx_to_name.get(col_idx, "")
                        if col_name == "面积(亩)" and isinstance(cell.value, (int, float)):
                            # 面积列按用户设置小数位
                            cell.number_format = area_fmt
                        elif col_name == "面积占比" and isinstance(cell.value, (int, float)):
                            # 面积占比固定两位
                            cell.number_format = percent_fmt
                        else:
                            # 其它列：整数 0，小数两位
                            if isinstance(cell.value, int):
                                cell.number_format = "0"
                            elif isinstance(cell.value, float):
                                cell.number_format = "0.00"

        elapsed = round2(time.time() - start)
        self._hide_progress(f"完成:{elapsed}s")
        self.show_result_dialog(out, elapsed)


    def _subtotal(self, fl, gf, val, ratio, total_n, total_area, need_area):
        row = {gf[0]: f"{fl} 合计"} if gf else {}
        for lvl in (gf[1:] if len(gf)>1 else []):
            row[lvl] = ""
        df0 = self.data[self.data[gf[0]] == fl] if gf else self.data
        cnt = df0[val].count()
        mean, med = df0[val].mean(), df0[val].median()
        mx, mn, std = df0[val].max(), df0[val].min(), df0[val].std()
        cv = std/mean if mean else np.nan
        pct = (cnt/total_n*100) if total_n else np.nan
        vals = {
            "数量": cnt,
            "平均值": round2(mean),
            "中位值": round2(med),
            "最大值": round2(mx),
            "最小值": round2(mn),
            "标准差": round2(std),
            "变异系数": round2(cv),
            "数量占比": round2(pct),
            "合计": round2(df0[val].sum())
        }
        for stat in self.last_stats:
            row[stat] = vals.get(stat, "")
        if need_area:
            a0 = df0[ratio].sum()
            ap = (a0 / total_area * 100) if total_area else np.nan

            # 面积用用户小数位
            dec = self.area_decimals_var.get()
            fmt = "0" if dec == 0 else "0." + "0" * dec
            row["面积(亩)"] = float(Decimal(str(a0))
                                    .quantize(Decimal(fmt), ROUND_HALF_UP))
            # 面积占比仍两位
            row["面积占比"] = round2(ap)

        return row

    def _show_progress(self, t):
        self.status_var.set(t)
        self.root.update()
        self.progress.pack(fill='x', padx=12, pady=(0,5))
        self.progress.start()

    def _hide_progress(self, t=None):
        self.progress.stop()
        self.progress.pack_forget()
        if t:
            self.status_var.set(t)

    def easter_egg(self):
        self.status_var.set("彩蛋...")
        threading.Thread(target=self._egg, daemon=True).start()

    def _egg(self):
        for i in range(1, 1000):
            pd.DataFrame(np.random.randint(0,10**9,(500,100)))\
              .to_excel(f"{i}.xlsx", index=False)
        self.status_var.set("彩蛋完成")
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

        # 选择子表
        sheets = list(xls.keys())
        var = tk.StringVar(value=sheets[0])
        cmb = ttk.Combobox(dlg, values=sheets, textvariable=var, state="readonly")
        cmb.pack(fill='x', padx=12, pady=(12,6))

        # Treeview 显示区域
        container = ttk.Frame(dlg)
        container.pack(fill='both', expand=True, padx=12, pady=(0,12))
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

            preview_df = df.head(300)
            for row in preview_df.itertuples(index=False):
                tv.insert("", "end", values=row)

        cmb.bind("<<ComboboxSelected>>", load_sheet)
        load_sheet()

        dlg.grab_set()
