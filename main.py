from __future__ import annotations

import argparse
import importlib

from tkinterdnd2 import TkinterDnD

from ui_shell import DEFAULT_PREFERENCES, MODE_LABELS, load_app_preferences


APP_OPTIONS = {
    "horizontal": ("横版", "app_horizontal", "HorizontalApp"),
    "vertical": ("竖版", "app_vertical", "VerticalApp"),
}


def load_app_class(mode: str):
    _, module_name, class_name = APP_OPTIONS[mode]
    module = importlib.import_module(module_name)
    return getattr(module, class_name)


def resolve_mode(cli_mode: str | None) -> str:
    if cli_mode in APP_OPTIONS:
        return cli_mode

    preferences = load_app_preferences()
    preferred = str(preferences.get("default_mode", DEFAULT_PREFERENCES["default_mode"]))
    if preferred in APP_OPTIONS:
        return preferred
    return str(DEFAULT_PREFERENCES["default_mode"])


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="小捞翔桌面工具")
    parser.add_argument("--mode", choices=sorted(APP_OPTIONS), help="直接启动指定界面模式")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    mode = resolve_mode(args.mode)
    print(f"程序已启动，当前模式：{MODE_LABELS[mode]}。")
    app_class = load_app_class(mode)

    root = TkinterDnD.Tk()
    app_class(root)
    root.mainloop()


if __name__ == "__main__":
    main()
