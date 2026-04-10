"""项目公共工具函数。"""

from __future__ import annotations

import os
import platform
import subprocess
from decimal import Decimal, ROUND_HALF_UP
from typing import Any

import pandas as pd
from pypinyin import Style, lazy_pinyin, pinyin


CSV_ENCODINGS = ("utf-8", "utf-8-sig", "gb18030", "gbk")


def sort_key(value: Any) -> tuple[str, int | float]:
    """生成适合中文排序的 key（拼音 + 笔画）。"""
    text = str(value).strip()
    if not text:
        return ("", float("inf"))

    first_char = text[0]
    pinyin_value = lazy_pinyin(first_char)[0]
    try:
        strokes = int(pinyin(first_char, style=Style.STROKES)[0][0])
    except Exception:
        strokes = float("inf")
    return (pinyin_value, strokes)



def round2(value: Any) -> Any:
    """四舍五入保留 2 位小数。"""
    try:
        return float(Decimal(str(value)).quantize(Decimal("0.00"), ROUND_HALF_UP))
    except Exception:
        return value



def open_with_default_app(filepath: str) -> None:
    """使用系统默认程序打开文件。"""
    system = platform.system()
    if system == "Windows":
        os.startfile(filepath)  # type: ignore[attr-defined]
    elif system == "Darwin":
        subprocess.call(["open", filepath])
    else:
        subprocess.call(["xdg-open", filepath])



def read_csv_safely(path: str, **kwargs: Any) -> pd.DataFrame:
    """按常见中文环境编码依次尝试读取 CSV。"""
    errors: list[str] = []
    for encoding in CSV_ENCODINGS:
        try:
            return pd.read_csv(path, encoding=encoding, **kwargs)
        except UnicodeDecodeError as exc:
            errors.append(f"{encoding}: {exc}")
        except Exception:
            raise

    detail = "\\n".join(errors) if errors else "未捕获到具体编码错误。"
    raise ValueError(f"无法识别 CSV 编码，可尝试另存为 UTF-8 后再导入。\\n{detail}")
