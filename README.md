# Battery Voltage vs Specific Capacity Plotter

一个面向电池/固态电池循环测试数据的小型 Python 项目，用于从设备导出的 Excel 明细数据中自动识别逐点曲线，并绘制常见的充放电曲线图（Voltage vs Specific Capacity）。

项目默认以“交互式运行”为主：

```bash
python main.py
```

运行后，程序会依次询问输入文件路径、sheet 名、输出路径、循环范围以及关键绘图参数；也支持命令行参数方式进行批处理。

现在也支持图形界面（GUI）运行：

```bash
python gui.py
```

或：

```bash
python main.py --gui
```

---

## 最近更新

- 新增桌面 GUI（`gui.py` / `python main.py --gui`）；
- 支持一个 Excel 内含多个子表格时自动逐个绘图；
- 支持同一个 sheet 内多个块状子表格自动拆分；
- 支持根据信息表中的原始子表路径/文件名自动命名输出；
- GUI 新增常用 `colormap` 下拉选择与色卡预览。

---

## 1. 项目功能介绍

本项目适合处理电池循环测试中的逐点数据，并输出科研绘图风格的充放电曲线图。

主要功能：

- 优先支持 `.xlsx` / `.xlsm` 文件；
- 自动识别最可能包含逐点曲线数据的 sheet；
- 自动识别逐点数据表头所在行，避免误读前面的循环汇总区；
- 支持一个 Excel 内存在多个子表格时依次读取并分别绘图（包括多个 sheet，或同一 sheet 内多个块状子表格）；
- 若 Excel 的信息表中包含原始子表路径/文件名，导出图片会优先按原子表名字自动命名；
- 自动兼容中英文列名，优先识别：
  - `循环序号`
  - `工作模式`
  - `电压/V`
  - `比容量/mAh/g`
- 根据工作模式区分充电 / 放电 / 静置；
- 静置数据自动跳过；
- 支持全部循环、指定循环、区间循环绘图；
- 同一循环的充放电默认同色不同线型；
- 支持 `png` / `svg` / `pdf` 高质量导出；
- 支持中文标题、中文字体、论文风格主题、透明背景、colormap、自定义线宽/线型/坐标范围；
- 提供内置 demo 数据，方便快速试跑；
- 提供桌面 GUI，可通过按钮选择文件、设置参数并直接执行绘图；
- 对 `.ccs` 原始文件做了清晰提示：当前不直接解析，建议先从设备软件导出为 `.xlsx` 再处理。

---

## 2. 安装方法

建议先创建虚拟环境，再安装依赖。

### Windows PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

### 依赖安装后运行

```powershell
python main.py
```

如需图形界面：

```powershell
python gui.py
```

---

## 3. 依赖说明

`requirements.txt` 中包含以下核心依赖：

- `pandas`：读取和整理 Excel 数据
- `matplotlib`：绘图和导出图片
- `openpyxl`：读取 `.xlsx` / `.xlsm`
- `numpy`：数值处理与 demo 数据生成

---

## 4. 输入数据格式说明

### 推荐输入格式

优先使用设备软件导出的 `.xlsx` 或 `.xlsm` 文件。

### 关于 `.ccs`

当前版本**没有直接实现稳定的 `.ccs` 原始文件解析**。如果你的原始数据是 `.ccs`，建议先在设备软件中导出为 `.xlsx`，再使用本工具处理。

### Excel 数据特点兼容

本工具针对以下常见情况做了处理：

1. 工作表前几行是循环汇总区，而不是真正的逐点数据；
2. 真正的逐点数据表头可能出现在靠后几行；
3. 表头可能包含中文列名；
4. 文件可能包含多个 sheet；
5. 同一个 sheet 中也可能包含多个前后排列的块状子表格；
6. 不同设备导出的工作模式文字可能不完全一致。

### 优先识别的关键列

程序会优先围绕以下列进行识别：

- `电压/V`
- `比容量/mAh/g`

同时建议表中尽量包含：

- `循环序号`
- `工作模式`
- `工步序号`（可选）
- `记录序号`（可选）

### 常见可兼容列名示例

- 循环列：`循环序号`、`循环号`、`cycle`
- 模式列：`工作模式`、`模式`、`mode`
- 电压列：`电压/V`、`电压`、`Voltage/V`
- 比容量列：`比容量/mAh/g`、`比容量`、`Specific Capacity`

---

## 5. 默认交互式运行方法

这是本项目**默认推荐的使用方式**。

```powershell
python main.py
```

程序会在终端中依次提示：

- 请输入数据文件路径
- 请输入 sheet 名（直接回车则自动识别）
- 请输入输出图片保存路径
- 请输入要绘制的循环编号，例如 `1,2,3` 或 `1-5`
- 是否显示图例（y/n）
- 请输入标题（直接回车使用默认标题）
- 是否进入高级参数设置（可继续设置 dpi、字体、线宽、坐标范围、colormap 等）

### 交互输入的设计细节

- 支持直接粘贴 Windows 路径；
- 自动去掉路径两端的引号；
- 输入文件不存在时会重新提示，而不是崩溃；
- 输出目录不存在时会自动创建；
- 输出文件缺少扩展名时，会根据输出格式自动补全；
- 如果自动识别到多个子表格，会在你给定的输出目录下按各子表名字分别保存图片；
- 直接回车即可接受默认值；
- 非法输入会给出提示并允许重新输入。

---

## 5.1 图形界面（GUI）运行方法

如果你不想输入命令参数，也不想逐行回答终端问题，可以直接使用 GUI：

```powershell
python gui.py
```

或者：

```powershell
python main.py --gui
```

GUI 中可直接完成：

- 选择输入 Excel；
- 选择输出目录；
- 设置基础文件名、输出格式、标题、循环编号；
- 通过下拉框选择常用 colormap，并查看色卡预览；
- 勾选是否显示图例、网格、按循环着色等选项；
- 设置 DPI、图尺寸、坐标范围、字体列表、模式映射等高级参数；
- 查看运行日志；
- 一键开始绘图。

对于包含多个子表格的 Excel，GUI 版也会按现有规则自动逐个输出图片。

---

## 5.2 GUI 中几个容易混淆的参数

### colormap 是什么？

当你开启“按循环着色”时，不同循环会从一个颜色方案中自动取色。

常见例子：

- `tab10`：适合离散循环，颜色区分明显；
- `viridis`：渐变自然，科研绘图常用；
- `plasma`：对比更强；
- `coolwarm`：冷暖渐变明显。

GUI 中可以直接下拉选择，也能看到色卡预览。

### 统一颜色是什么？

如果你**不按循环着色**，那么所有曲线都会优先使用“统一颜色”。

例如填入：

```text
#1f77b4
```

则所有线条会统一用这一种颜色。

### 两者如何配合？

- 开启“按循环着色” → 主要由 `colormap` 控制；
- 关闭“按循环着色” → 主要由“统一颜色”控制。

---

## 6. 命令行运行方法

除了默认的交互式运行，本工具也支持命令行参数方式，便于批处理。

### 示例 1：最常见写法

```powershell
python main.py --input "D:\data\cell_01.xlsx" --output "D:\figures\curve.png"
```

### 示例 2：指定 sheet 和循环编号

```powershell
python main.py --input "D:\data\cell_01.xlsx" --output "D:\figures\curve.png" --sheet Sheet1 --cycles 1 2 5 10
```

### 示例 3：用区间选择循环

```powershell
python main.py --input "D:\data\cell_01.xlsx" --output "D:\figures\curve.svg" --cycles 1-5
```

### 示例 4：自定义绘图参数

```powershell
python main.py \
  --input "D:\data\cell_01.xlsx" \
  --output "D:\figures\curve.pdf" \
  --sheet Sheet1 \
  --cycles 1 2 3 \
  --title "Solid-state battery cycling curves" \
  --x-label "Specific Capacity (mAh/g)" \
  --y-label "Voltage (V)" \
  --dpi 600 \
  --colormap viridis \
  --line-width 2.0 \
  --legend-loc "upper right" \
  --theme paper
```

### 示例 5：关闭图例和网格

```powershell
python main.py --input data.xlsx --output figure.png --no-show-legend --no-grid
```

### 示例 6：内置 demo

```powershell
python main.py --demo --output demo_curve.png
```

### 示例 7：一个 Excel 中有多个子表格

```powershell
python main.py --input "D:\data\merged_export.xlsx" --output "D:\figures\result.png"
```

说明：

- 如果文件中只识别到一个可绘制子表格，仍按你指定的 `result.png` 输出；
- 如果识别到多个子表格，程序会忽略 `result` 这个文件名部分，仅使用其所在目录 `D:\figures\`；
- 每张图会按原子表名字自动命名，例如：
  - `cyclingtest_2026.3.14-Gr_only_NMC0.788mg-CritCRate_20260314155506_DefaultGroup_37_1.png`
  - `cyclingtest_2026.3.14-Gr_only_NMC0.799mg-CritCRate_20260314160000_DefaultGroup_37_2.png`

---

## 7. 运行时会提示输入哪些内容

默认交互式模式下，程序会提示输入：

1. 数据文件完整路径
2. sheet 名（可选）
3. 输出图片完整路径
4. 循环范围或循环编号（可选）
5. 是否显示图例
6. 图标题
7. 是否进入高级参数设置

如果进入高级设置，还可以继续配置：

- 输出格式
- dpi
- 图例位置
- 坐标轴标题
- 标题字号 / 坐标轴字号 / 刻度字号
- 线宽
- 统一颜色
- 充电线型 / 放电线型
- 网格开关与网格线型
- 图尺寸
- 是否按循环着色
- colormap
- 主题（paper/default）
- 字体列表
- 坐标轴范围
- 是否自动排序数据点
- 是否对比容量取绝对值
- 是否保存后同时显示图像
- 是否透明背景
- 工作模式映射补充规则

补充说明：

- 若同一个 Excel 中有多个可绘制子表格，程序会逐个输出；
- 若你提供的信息表里含有类似 `路径` / `文件名` 的原始子表信息，会优先据此命名导出图片；
- 若没有这些信息，则会回退为 `sheet名_part1`、`sheet名_part2` 这类命名。

---

## 8. 如何输入数据文件路径与输出图片路径

### 输入文件路径示例

```text
D:\battery_data\sample.xlsx
```

也可以直接粘贴带引号的路径：

```text
"D:\battery_data\sample.xlsx"
```

程序会自动去掉首尾引号。

### 输出图片路径示例

```text
D:\battery_figures\sample_curve.png
```

如果当前 Excel 中实际包含多个子表格，也可以仍然这样输入。

程序会把 `D:\battery_figures` 当作输出目录，再按识别到的子表名字自动生成多个文件名。

如果你只输入：

```text
D:\battery_figures\sample_curve
```

程序会根据当前输出格式自动补成：

```text
D:\battery_figures\sample_curve.png
```

### 多子表格时的输出命名规则

如果 Excel 中存在多个可绘制子表格，程序会优先读取信息表中的原始路径/文件名，并按下面规则命名：

原始子表路径例如：

```text
E:/HKUST-GZ/cyclingtest/2026.3.14-Gr_only_NMC0.788mg-CritCRate_20260314155506_DefaultGroup_37_1.ccs
```

则导出文件名会优先变成：

```text
cyclingtest_2026.3.14-Gr_only_NMC0.788mg-CritCRate_20260314155506_DefaultGroup_37_1.png
```

如果信息表中没有这些信息，则会回退为：

```text
sheet名_part1.png
sheet名_part2.png
```

---

## 9. 如何选择指定循环进行绘图

支持以下输入方式：

- 全部循环：直接回车
- 指定多个循环：`1,2,5,10`
- 指定区间：`1-5`
- 混合写法：`1,2,5-8,10`

命令行参数中也支持类似写法：

```powershell
python main.py --input data.xlsx --output figure.png --cycles 1 2 5-8 10
```

---

## 10. 如何修改线条颜色、粗细、标题、字体、图例

### 交互式方式

运行：

```powershell
python main.py
```

当程序问到“是否进入高级参数设置”时输入 `y`，即可继续设置：

- `线宽`
- `统一线条颜色`
- `colormap`
- `标题`
- `字体列表`
- `图例位置`
- `是否按循环分别着色`
- `充电/放电线型`

### 命令行方式

```powershell
python main.py \
  --input data.xlsx \
  --output figure.svg \
  --title "Cycle Curves" \
  --line-width 2.2 \
  --line-color "#d62728" \
  --font-family "Microsoft YaHei,SimHei,DejaVu Sans" \
  --legend-loc "upper right" \
  --colormap plasma
```

---

## 11. 如何处理中文字体问题

本项目已经默认设置了一组常见中文字体候选：

- `Microsoft YaHei`
- `SimHei`
- `Noto Sans CJK SC`
- `Source Han Sans SC`
- `Arial Unicode MS`
- `DejaVu Sans`

如果你的系统没有合适字体，可能会出现：

- 中文标题显示成方框
- 中文图例乱码

### 解决方法 1：在交互式高级设置中指定字体

例如输入：

```text
Microsoft YaHei,SimHei,DejaVu Sans
```

### 解决方法 2：命令行指定字体

```powershell
python main.py --input data.xlsx --output figure.png --font-family "Microsoft YaHei,SimHei,DejaVu Sans"
```

---

## 12. 常见报错及解决办法

### 报错：文件不存在

原因：输入路径错误，或文件没有放在该位置。

解决：重新检查路径，建议直接复制资源管理器里的完整路径。

### 报错：未找到 sheet

原因：指定的 sheet 名拼写错误。

解决：检查 Excel 中真实的 sheet 名，或直接回车让程序自动识别。

### 报错：未能在任何 sheet 中找到包含 电压/V 和 比容量/mAh/g 的逐点数据表头

原因：

- 当前文件不是逐点测试数据导出表；
- 表头命名与设备导出差异较大；
- 数据被手工修改后丢失关键列。

解决：

1. 检查原始导出文件是否包含 `电压/V` 与 `比容量/mAh/g`；
2. 检查逐点数据是否在其他 sheet；
3. 尽量直接使用设备软件导出的原始明细表；
4. 如果列名差异较大，可自行在 `parser.py` 中补充 `HEADER_ALIASES`。

### 报错：没有找到可绘制的充放电逐点数据

原因：

- 所选循环范围没有数据；
- 工作模式列没有被正确识别为 charge/discharge；
- 该 sheet 主要是汇总数据，不是逐点数据。

解决：

1. 先不筛选循环，尝试全部循环；
2. 在高级设置中补充工作模式映射；
3. 检查原始模式列里是否使用了特殊命名。

### 报错：暂未稳定支持 .ccs 原始文件

原因：当前版本没有直接解析 `.ccs`。

解决：请先在设备软件中导出为 `.xlsx`，再运行本工具。

### 报错：输出路径必须包含有效的文件名

原因：输出路径只写了目录，没有写文件名。

解决：请写成类似：

```text
D:\figures\curve.png
```

---

## 13. 输出示例说明

输出图像默认具有以下特征：

- 横坐标：`Specific Capacity (mAh/g)`
- 纵坐标：`Voltage (V)`
- 同一循环的充电/放电曲线使用相同颜色、不同线型
- 不同循环默认按 colormap 着色
- 适合常规科研汇报和论文初稿制图

支持导出格式：

- `png`：适合汇报、Word、PPT
- `svg`：适合矢量编辑
- `pdf`：适合论文与高质量排版

---

## 14. 项目结构

```text
main.py          # 程序入口，支持交互式 + 命令行 + --gui
gui.py           # tkinter 图形界面
config.py        # 默认配置
utils.py         # 路径处理、循环解析、交互辅助
parser.py        # 表头识别、字段映射、模式分类、数据清洗
data_loader.py   # Excel 读取、多 sheet / 多块子表格识别、命名匹配
plotter.py       # matplotlib 绘图与导出
requirements.txt # 依赖列表
README.md        # 使用说明
USER_MANUAL.md   # 面向新手的详细手册
```

---

## 15. 代码可扩展点

如果你后续想继续扩展，本项目比较容易修改的地方包括：

- 在 `parser.py` 的 `HEADER_ALIASES` 中添加新的列名别名；
- 在 `config.py` 的 `DEFAULT_MODE_RULES` 中增加新的工作模式识别规则；
- 在 `plotter.py` 中修改论文风格、颜色方案、图例样式；
- 在 `data_loader.py` 中继续扩展对其他文件类型的支持。

---

## 16. 快速建议

如果你的设备导出格式经常比较固定，建议先用一次交互式模式跑通；如果后续要批量处理，再把参数整理成命令行命令或批处理脚本。
