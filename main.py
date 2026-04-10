from __future__ import annotations

import importlib
import tkinter as tk
from tkinter import ttk

from tkinterdnd2 import TkinterDnD


APP_OPTIONS = {
    "horizontal": ("横版", "app_horizontal", "HorizontalApp"),
    "vertical": ("竖版", "app_vertical", "VerticalApp"),
}


def center_window(window: tk.Misc, width: int, height: int) -> None:
    """让窗口尽量出现在屏幕中间。"""
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = max((screen_width - width) // 2, 0)
    y = max((screen_height - height) // 2, 0)
    window.geometry(f"{width}x{height}+{x}+{y}")


def choose_orientation() -> str | None:
    """先弹出独立选择窗，再创建主程序窗口。"""
    dialog = tk.Tk()
    dialog.title("选择应用模式")
    width, height = 280, 160
    dialog.geometry(f"{width}x{height}")
    dialog.resizable(False, False)
    center_window(dialog, width, height)
    dialog.lift()
    dialog.focus_force()

    try:
        dialog.attributes("-topmost", True)
        dialog.after(400, lambda: dialog.attributes("-topmost", False))
    except tk.TclError:
        pass

    choice = tk.StringVar(value="horizontal")
    result = {"mode": None}

    ttk.Label(dialog, text="请选择启动模式：").pack(pady=(18, 8))
    for mode, (label, _, _) in APP_OPTIONS.items():
        ttk.Radiobutton(dialog, text=label, variable=choice, value=mode).pack(anchor="w", padx=48, pady=2)

    def on_confirm() -> None:
        result["mode"] = choice.get()
        dialog.destroy()

    def on_cancel() -> None:
        result["mode"] = None
        dialog.destroy()

    ttk.Button(dialog, text="确定", command=on_confirm).pack(pady=14)
    dialog.protocol("WM_DELETE_WINDOW", on_cancel)
    dialog.mainloop()
    return result["mode"]


def load_app_class(mode: str):
    _, module_name, class_name = APP_OPTIONS[mode]
    module = importlib.import_module(module_name)
    return getattr(module, class_name)


def main() -> None:
    print("程序已启动，请查看弹出的窗口并选择启动模式。")
    mode = choose_orientation()
    if mode is None:
        return

    app_class = load_app_class(mode)
    root = TkinterDnD.Tk()
    app_class(root)
    root.mainloop()


if __name__ == "__main__":
    main()
