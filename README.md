# 小捞翔·至尊版

一个基于 **Tkinter + Pandas + OpenPyXL** 的桌面统计工具，支持 **横版** / **竖版** 两种结果导出方式，适合对 Excel / CSV 数据做分组统计、面积占比计算、批量导出和自定义排序。

## 功能特性

- 支持 Excel / CSV 文件导入
- 支持横版、竖版两种统计导出模式
- 支持 1～3 级分类分组
- 支持批量值字段、批量分组字段
- 支持统计量勾选与自定义排序
- 支持面积字段汇总和面积占比计算
- 支持最近打开记录
- 支持导出结果预览
- 支持拖拽导入（横版）

## 这次整理做了什么

这份仓库版代码在尽量不改变原有功能的前提下，做了几项更适合 GitHub 托管的整理：

1. **优化启动流程**：避免重复创建 Tk 根窗口，启动逻辑更稳。
2. **抽出公共工具模块**：把 CSV 读取、文件打开、中文排序、四舍五入等逻辑集中到 `common_utils.py`。
3. **修复 CSV 加载体验**：横版在读取 CSV 后也会正确刷新值字段、面积字段和分类级下拉。
4. **清理重复导入与冗余代码**：减少重复 import，提升可维护性。
5. **补齐仓库文件**：增加 `README.md`、`requirements.txt`、`.gitignore`，更方便直接上传 GitHub。

## 项目结构

```text
.
├── app_horizontal.py      # 横版主界面
├── app_vertical.py        # 竖版主界面
├── common_utils.py        # 公共工具函数
├── main.py                # 启动入口
├── requirements.txt       # Python 依赖
├── .gitignore             # Git 忽略规则
└── README.md              # 项目说明
```

## 运行环境

- Python 3.10+
- Windows 优先（当前交互逻辑更偏桌面端 Windows 使用习惯）

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动项目

```bash
python main.py
```

启动后会先弹出模式选择窗口：

- **横版**：适合导出横向展开的统计表
- **竖版**：适合导出纵向指标表

## 依赖说明

| 依赖 | 作用 |
|---|---|
| pandas | 数据读取与统计 |
| numpy | 数值处理 |
| openpyxl | Excel 导出 |
| pypinyin | 中文排序 |
| sv-ttk | 横版界面主题 |
| tkinterdnd2 | 拖拽导入 |

## 输出文件说明

程序会在当前目录生成结果文件，例如：

- `原文件名_结果.xlsx`
- `原文件名_计算小捞翔后.xlsx`

建议把导出结果放到单独目录，方便管理。

## 配置文件说明

程序运行过程中会生成以下本地配置文件：

- `small_shit.json`
- `small_shit_vertical.json`

这两个文件用于保存：

- 最近打开记录
- 自定义排序
- 上次选择的统计量
- 更新历史
- 面积小数位设置

这些文件已加入 `.gitignore`，一般 **不建议提交到 GitHub**。

## 打包建议

如果你后续要发给别人直接使用，可以考虑使用 PyInstaller：

```bash
pip install pyinstaller
pyinstaller -F -w main.py
```

如需把依赖资源一起打包，建议后续单独补一个 `build.spec` 文件再细化。

## 已知说明

- 横版支持拖拽导入，竖版当前仍以按钮选文件为主。
- CSV 会自动尝试 `utf-8`、`utf-8-sig`、`gb18030`、`gbk` 编码读取。
- 预览窗口为了避免界面卡顿，只展示部分行用于快速检查。

## License

当前仓库未附带正式许可证。如需开源发布，建议补充 `MIT` 或 `Apache-2.0` 许可证文件。
