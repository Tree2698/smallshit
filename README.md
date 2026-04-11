# 小捞翔 · 统计增强版

一个基于 **Tkinter + Pandas + OpenPyXL** 的桌面统计工具，支持 **横版** / **竖版** 两种导出模式，适合 Excel / CSV 数据的分组统计、面积换算、批量任务和报表发布。

## 这次升级包含什么

这版在原来的基础上，一次性补上了三梯队的大部分核心能力，并统一整理了打包与仓库说明。

## v2.0.3 修复

- 修复源码运行或打包后切换到竖版时报 `No module named app_vertical`
- 修复竖版统计说明页面积换算 f-string 语法错误
- 修复横版条件筛选里 `apply_filter_conditions` / `describe_filter_conditions` 未定义
- 优化打包配置，显式包含横版、竖版和高级功能模块

### 第一梯队

- 统计图表导出
  - 导出结果时可自动生成图表工作表
  - 默认会为结果表生成柱状图，分组较少时附带饼图
- 筛选方案保存
  - 可把当前筛选条件保存为命名方案
  - 支持再次应用、删除
- 批量任务中心
  - 可把当前配置保存为任务
  - 支持批量执行、导入、导出
- 面积单位换算增强版
  - 支持平方米、亩、公顷、平方千米
  - 导出列名自动写成 `面积（单位）`
- 多文件合并统计
  - 支持一次导入多个 Excel / CSV 合并后统计
  - 自动追加 `来源文件` 字段

### 第二梯队

- 字段智能推荐
  - 自动识别数值字段、日期字段、分类字段
  - 可一键应用推荐的值字段 / 面积字段 / 日期字段
- 日期类统计
  - 支持按年、季度、月、周、日生成日期分组字段
- 分组映射 / 区间分组
  - 支持值映射
  - 支持数值断点区间分组
  - 支持日期字段派生分组
- 更强的预览模式
  - 支持前 N 行、随机抽样、关键词命中预览
- 错误提示优化
  - 关键错误改为更易懂的业务提示
  - 计算失败会弹出可复制详细信息的错误窗口

### 第三梯队

- 工作区
  - 支持保存当前工作状态为工作区 JSON
  - 支持重新加载工作区
- 导出发布链路
  - 可为导出结果一键生成“发布包” ZIP
  - 自动附带发布说明文本
- 插件化指标
  - 支持通过 `plugins/` 目录加载自定义统计指标
  - 已附带示例插件 `plugins/example_plugin.py`
- 操作历史 / 撤销
  - 支持记录关键操作历史
  - 支持撤销最近一步高级配置操作

## 当前主要菜单结构

- **文件**：打开文件、最近打开、首选项、导出目录
- **编辑**：统计量、自定义排序、筛选、模板、命令面板
- **数据增强**：字段推荐、日期统计、分组映射、面积换算、多文件合并
- **流程**：筛选方案、批量任务中心、保存/加载工作区、撤销、操作历史
- **发布**：生成发布包
- **查看**：数据检查器、检查报告、预览上次导出结果、主题切换
- **运行**：计算并导出、打开上次导出结果、打开导出目录

## 项目结构

```text
.
├── app_horizontal.py
├── app_vertical.py
├── common_utils.py
├── feature_support.py
├── ui_shell.py
├── main.py
├── main.spec
├── requirements.txt
├── build_windows.bat
├── plugins/
│   └── example_plugin.py
└── README.md
```

## 运行环境

- Python 3.10+
- Windows 优先

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行

```bash
python main.py
```

程序会按首选项里的默认模式启动，也可以手动指定：

```bash
python main.py --mode horizontal
python main.py --mode vertical
```

## 打包

### 方式 1：直接用 spec

```bash
pyinstaller --noconfirm --clean main.spec
```

### 方式 2：运行批处理

```bash
build_windows.bat
```

> 如果你要带图标，请确保项目根目录存在 `x.ico`。

## 插件统计指标

把自定义指标放到 `plugins/` 目录即可。

### 简单插件格式

```python
STAT_NAME = "我的指标"

def compute(series):
    values = series.dropna()
    if len(values) == 0:
        return None
    return float(values.mean())
```

也支持：

```python
def register_stats():
    return {
        "指标A": func_a,
        "指标B": func_b,
    }
```

## 发布包

导出完成后可以生成发布包，内容包括：

- 导出的 Excel 结果
- 自动生成的发布说明文本
- 打包后的 ZIP 文件

## 本地配置文件

运行过程中会生成：

- `small_shit.json`
- `small_shit_vertical.json`
- `app_preferences.json`

这些文件主要保存：

- 最近打开
- 首选项
- 自定义排序
- 统计量选择
- 工作区相关状态
- 高级增强配置

## 已知说明

- Windows 可执行文件需要在 Windows 环境中打包
- 某些高级功能依赖当前已加载数据字段，如果更换源数据结构差异很大，建议重新检查模板 / 任务 / 工作区
- 图表会优先选择当前结果表中的主要统计列，复杂多指标图表仍建议后续按业务再定制

## License

当前仓库未附正式许可证。对外开源前建议补充 MIT License。
