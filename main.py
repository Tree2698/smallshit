from __future__ import annotations

import importlib
import tkinter as tk
from tkinter import ttk

from tkinterdnd2 import TkinterDnD


APP_OPTIONS = {
    "horizontal": ("横版", "app_horizontal", "HorizontalApp"),
    "vertical": ("竖版", "app_vertical", "VerticalApp"),
}


def choose_orientation(root: tk.Misc) -> str | None:
    """使用模态对话框选择启动模式，避免重复创建 Tk 根窗口。"""
    dialog = tk.Toplevel(root)
    dialog.title("选择应用模式")
    dialog.geometry("260x140")
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.grab_set()

    choice = tk.StringVar(value="horizontal")
    result = {"mode": None}

    ttk.Label(dialog, text="请选择启动模式：").pack(pady=(14, 6))
    for mode, (label, _, _) in APP_OPTIONS.items():
        ttk.Radiobutton(dialog, text=label, variable=choice, value=mode).pack(anchor="w", padx=36, pady=2)

    def on_confirm() -> None:
        result["mode"] = choice.get()
        dialog.destroy()

    def on_cancel() -> None:
        result["mode"] = None
        dialog.destroy()

    ttk.Button(dialog, text="确定", command=on_confirm).pack(pady=12)
    dialog.protocol("WM_DELETE_WINDOW", on_cancel)
    root.wait_window(dialog)
    return result["mode"]



def load_app_class(mode: str):
    _, module_name, class_name = APP_OPTIONS[mode]
    module = importlib.import_module(module_name)
    return getattr(module, class_name)



def main() -> None:
    root = TkinterDnD.Tk()
    root.withdraw()

    mode = choose_orientation(root)
    if mode is None:
        root.destroy()
        return

    app_class = load_app_class(mode)
    root.deiconify()
    app_class(root)
    root.mainloop()


if __name__ == "__main__":
    main()
