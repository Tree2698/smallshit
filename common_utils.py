"""项目公共工具函数。"""

from __future__ import annotations

import os
import platform
import subprocess
from decimal import Decimal, ROUND_HALF_UP
from typing import Any

import numpy as np
import pandas as pd
from pypinyin import Style, lazy_pinyin, pinyin


CSV_ENCODINGS = ("utf-8", "utf-8-sig", "gb18030", "gbk")
FILTER_OPERATORS = [
    "等于",
    "不等于",
    "包含",
    "不包含",
    "大于",
    "大于等于",
    "小于",
    "小于等于",
    "为空",
    "不为空",
]
NUMERIC_FILTER_OPERATORS = {"大于", "大于等于", "小于", "小于等于"}
EXTENDED_STATS = [
    "数量",
    "缺失数",
    "缺失率",
    "平均值",
    "中位值",
    "最小值",
    "下四分位数",
    "上四分位数",
    "最大值",
    "极差",
    "标准差",
    "方差",
    "变异系数",
    "偏度",
    "峰度",
    "合计",
    "数量占比",
]
COUNT_MASK_STATS = {"最大值", "最小值", "中位值", "下四分位数", "上四分位数", "极差"}
STD_MASK_STATS = {"标准差", "方差", "变异系数"}
SHAPE_MASK_STATS = {"偏度", "峰度"}


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


def calculate_series_stats(series: pd.Series, total_valid_count: int | None = None) -> dict[str, Any]:
    """基于数值序列计算扩展统计指标。"""
    numeric = pd.to_numeric(series, errors="coerce")
    valid_count = int(numeric.count())
    total_count = int(len(numeric))
    missing_count = total_count - valid_count

    if valid_count:
        q25 = numeric.quantile(0.25)
        q75 = numeric.quantile(0.75)
        minimum = numeric.min()
        maximum = numeric.max()
        std = numeric.std()
        mean = numeric.mean()
        variance = numeric.var()
        total_sum = numeric.sum()
        skew = numeric.skew()
        kurt = numeric.kurt()
    else:
        q25 = np.nan
        q75 = np.nan
        minimum = np.nan
        maximum = np.nan
        std = np.nan
        mean = np.nan
        variance = np.nan
        total_sum = np.nan
        skew = np.nan
        kurt = np.nan

    count_pct = (valid_count / total_valid_count * 100) if total_valid_count else np.nan
    missing_rate = (missing_count / total_count * 100) if total_count else np.nan
    value_range = maximum - minimum if valid_count else np.nan
    cv = (std / mean) if valid_count and mean not in (0, 0.0) else np.nan

    return {
        "数量": valid_count,
        "缺失数": missing_count,
        "缺失率": round2(missing_rate),
        "平均值": round2(mean),
        "中位值": round2(numeric.median()) if valid_count else np.nan,
        "最小值": round2(minimum),
        "下四分位数": round2(q25),
        "上四分位数": round2(q75),
        "最大值": round2(maximum),
        "极差": round2(value_range),
        "标准差": round2(std),
        "方差": round2(variance),
        "变异系数": round2(cv),
        "偏度": round2(skew),
        "峰度": round2(kurt),
        "合计": round2(total_sum),
        "数量占比": round2(count_pct),
    }


def apply_count_masks(df: pd.DataFrame) -> pd.DataFrame:
    """按样本量对不稳定指标做空值屏蔽。"""
    if "数量" not in df.columns:
        return df

    result = df.copy()
    counts = pd.to_numeric(result["数量"], errors="coerce").fillna(0)

    for name in COUNT_MASK_STATS:
        if name in result.columns:
            result.loc[counts <= 1, name] = np.nan

    for name in STD_MASK_STATS:
        if name in result.columns:
            result.loc[counts <= 5, name] = np.nan

    for name in SHAPE_MASK_STATS:
        if name in result.columns:
            result.loc[counts <= 7, name] = np.nan

    return result


def build_grouped_stats_frame(df: pd.DataFrame, field: str, group_fields: list[str]) -> pd.DataFrame:
    """把 DataFrame 中某个值字段按分组计算成统计结果表。"""
    data = df.copy()
    data[field] = pd.to_numeric(data[field], errors="coerce")
    total_valid_count = int(data[field].count())

    if group_fields:
        rows: list[dict[str, Any]] = []
        for keys, group in data.groupby(group_fields, dropna=False):
            key_tuple = keys if isinstance(keys, tuple) else (keys,)
            row = {column: value for column, value in zip(group_fields, key_tuple)}
            row.update(calculate_series_stats(group[field], total_valid_count))
            rows.append(row)
        return pd.DataFrame(rows)

    return pd.DataFrame([calculate_series_stats(data[field], total_valid_count)])


def describe_filter_conditions(conditions: list[dict[str, Any]]) -> str:
    """把筛选条件列表转成便于显示的摘要。"""
    parts = []
    for condition in conditions:
        if not condition.get("field") or not condition.get("op"):
            continue
        field = str(condition["field"])
        op = str(condition["op"])
        value = str(condition.get("value", "")).strip()
        if op in {"为空", "不为空"}:
            parts.append(f"{field} {op}")
        else:
            parts.append(f"{field} {op} {value}")
    return "；".join(parts) if parts else "未启用"


def apply_filter_conditions(df: pd.DataFrame, conditions: list[dict[str, Any]]) -> pd.DataFrame:
    """按条件列表过滤 DataFrame。"""
    filtered = df.copy()
    for condition in conditions:
        field = str(condition.get("field", "")).strip()
        op = str(condition.get("op", "")).strip()
        value = str(condition.get("value", "")).strip()
        enabled = condition.get("enabled", True)

        if not enabled or not field or not op:
            continue
        if field not in filtered.columns:
            raise KeyError(f"筛选字段不存在：{field}")

        series = filtered[field]
        text_series = series.fillna("").astype(str).str.strip()

        if op == "为空":
            mask = series.isna() | (text_series == "")
        elif op == "不为空":
            mask = ~(series.isna() | (text_series == ""))
        elif op == "包含":
            mask = text_series.str.contains(value, case=False, na=False, regex=False)
        elif op == "不包含":
            mask = ~text_series.str.contains(value, case=False, na=False, regex=False)
        elif op in NUMERIC_FILTER_OPERATORS:
            try:
                target = float(value)
            except ValueError as exc:
                raise ValueError(f"筛选条件“{field} {op} {value}”不是有效数字。") from exc
            numeric_series = pd.to_numeric(series, errors="coerce")
            if op == "大于":
                mask = numeric_series > target
            elif op == "大于等于":
                mask = numeric_series >= target
            elif op == "小于":
                mask = numeric_series < target
            else:
                mask = numeric_series <= target
        else:
            numeric_series = pd.to_numeric(series, errors="coerce")
            numeric_target = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
            if pd.notna(numeric_target) and numeric_series.notna().any():
                compare_series = numeric_series
                compare_value = float(numeric_target)
            else:
                compare_series = text_series
                compare_value = value

            if op == "等于":
                mask = compare_series == compare_value
            elif op == "不等于":
                mask = compare_series != compare_value
            else:
                raise ValueError(f"不支持的筛选操作：{op}")

        filtered = filtered.loc[mask.fillna(False)].copy()

    return filtered
