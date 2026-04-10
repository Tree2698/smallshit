# 小捞翔·至尊版

一个基于 **Tkinter + Pandas + OpenPyXL** 的桌面统计工具，支持 **横版** / **竖版** 两种统计导出模式，适合对 Excel / CSV 数据做分组统计、面积占比计算、批量导出、自定义排序与条件筛选。

## 这次修复了什么

这份版本重点把之前“看起来有功能、但没有完全闭环”的问题一次性补齐了：

- 修复：**条件筛选现在会真实参与导出统计**
- 修复：**竖版的条件筛选 / 分组模板功能已完整可用**
- 修复：**打包后模式切换路径不稳定**，现在兼容源码运行和 PyInstaller 打包运行
- 修复：`main.py` 改为**显式导入**，避免 PyInstaller 漏打 `app_horizontal` / `app_vertical`
- 修复：导出结果预览改为**限量预览**，降低大表卡顿概率
- 优化：导出时自动写入 **统计说明** 与 **导出清单**
- 优化：会在配置里记录最近一次导出文件路径，方便继续打开
- 优化：打包配置改为更适合 Windows 分发的 `main.spec`

## 功能特性

- 支持 Excel / CSV 文件导入
- 支持横版、竖版两种统计导出模式
- 支持 1～3 级分类分组
- 支持批量值字段、批量分组字段
- 支持扩展统计指标：
  - 数量
  - 缺失数 / 缺失率
  - 平均值 / 中位值
  - 最小值 / 最大值 / 极差
  - 下四分位数 / 上四分位数
  - 标准差 / 方差 / 变异系数
  - 偏度 / 峰度
  - 合计 / 数量占比
- 支持面积字段汇总和面积占比计算
- 支持条件筛选
- 支持分组模板保存 / 导入 / 导出
- 支持最近打开记录
- 支持数据检查器
- 支持导出结果预览
- 支持拖拽导入
- 支持首选项：
  - 默认启动模式
  - 主题
  - 导出目录
  - 导出文件名时间戳
  - 启动时恢复最近文件

## 项目结构

```text
.
├── app_horizontal.py      # 横版主界面
├── app_vertical.py        # 竖版主界面
├── common_utils.py        # 公共工具函数
├── ui_shell.py            # 公共界面壳层 / 菜单 / 首选项 / 数据检查器
├── main.py                # 启动入口
├── main.spec              # PyInstaller 打包配置
├── build_windows.bat      # Windows 一键打包脚本
├── app_preferences.json   # 默认首选项
├── requirements.txt       # 运行依赖
├── .gitignore             # Git 忽略规则
└── README.md              # 项目说明
```

## 运行环境

- Python 3.10+
- Windows 优先

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动项目

```bash
python main.py
```

也可以直接指定模式：

```bash
python main.py --mode horizontal
python main.py --mode vertical
```

## 打包 Windows 可执行文件

### 方式一：直接用 spec
推荐：

```bash
pyinstaller --noconfirm --clean main.spec
```

### 方式二：直接命令行打包
如果不想用 spec，也可以：

```bash
pyinstaller --onedir --windowed main.py --hidden-import app_horizontal --hidden-import app_vertical
```

如果项目目录里有 `x.ico`，可以再加：

```bash
--icon=x.ico
```

### 一键打包
Windows 下直接双击或执行：

```bash
build_windows.bat
```

## 输出文件说明

程序会在你设置的导出目录中生成结果文件，例如：

- `原文件名_结果.xlsx`
- `原文件名_计算小捞翔后.xlsx`

导出结果中会附带：

- `统计说明`
- `导出清单`

## 配置文件说明

程序运行后会生成或更新以下本地配置文件：

- `small_shit.json`
- `small_shit_vertical.json`

用于保存：

- 最近打开记录
- 自定义排序
- 统计量选择
- 界面状态
- 筛选条件
- 分组模板
- 最近一次导出文件路径

这些文件属于本地使用状态，**不建议提交到 GitHub**。

## 说明

- 当前版本已经修复“筛选只显示、不参与导出”的问题。
- 当前版本已经补齐竖版筛选与模板功能。
- 为了兼容 PyInstaller，入口文件已避免使用动态模块导入。
- 结果预览默认只展示部分行，用于快速检查，避免大文件卡死。

## License

当前仓库未附带正式许可证。如需公开开源，建议补充 `MIT` 或 `Apache-2.0`。
