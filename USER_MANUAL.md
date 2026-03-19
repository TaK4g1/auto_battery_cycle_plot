# 用户手册

这是一份面向日常使用者的说明。  
如果你不想看太多技术细节，只想知道“怎么跑、怎么出图、怎么改参数”，看这份就够了。

---

## 1. 这个项目是干什么的？

它可以帮你做下面几件事：

- 读取电池测试 Excel 数据
- 自动识别逐点充放电数据
- 画出 **Voltage vs Specific Capacity** 曲线
- 支持筛选指定循环
- 支持一个 Excel 中有多个子表格时自动分别出图
- 支持用 Origin 输出**可编辑工程**

---

## 2. 我应该怎么启动？

### 方法 1：最推荐

```powershell
python main.py
```

这是默认交互式模式。  
程序会一步一步问你要什么参数，适合大多数情况。

### 方法 2：图形界面

```powershell
python gui.py
```

或者：

```powershell
python main.py --gui
```

如果你更喜欢点按钮，而不是在终端里逐项输入，建议直接用 GUI。

---

## 3. 第一次使用前要做什么？

在项目目录里执行：

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

以后再次使用时，一般只需要：

```powershell
.\.venv\Scripts\activate
python main.py
```

---

## 4. 输入文件支持什么格式？

推荐：

- `.xlsx`
- `.xlsm`

### 关于 `.ccs`

当前版本不建议直接拿 `.ccs` 给程序处理。  
请先在设备软件中导出为 Excel，再导入本项目。

---

## 5. 我最常见的使用流程

### 方式 A：交互式

运行：

```powershell
python main.py
```

然后按提示输入：

1. Excel 文件路径
2. sheet 名（不会就直接回车）
3. 输出图片路径
4. 循环编号（不会就直接回车表示全部）
5. 是否显示图例
6. 标题
7. 是否进入高级设置

### 方式 B：GUI

在 GUI 中通常这样操作：

1. 选择输入 Excel
2. 选择输出目录
3. 设置基础文件名
4. 选择输出格式
5. 选择后端：`matplotlib` 或 `origin`
6. 设置循环编号
7. 点击开始绘图

---

## 6. 多子表格是怎么处理的？

如果一个 Excel 里有多个子表格，程序会自动：

- 按顺序识别每个子表格
- 每个子表格各画一张图
- 每张图单独保存

### 命名规则

如果信息表中识别到原始子表路径，比如：

```text
E:/HKUST-GZ/cyclingtest/2026.3.14-Gr_only_NMC0.788mg-CritCRate_20260314155506_DefaultGroup_37_1.ccs
```

导出结果会优先命名为：

```text
cyclingtest_2026.3.14-Gr_only_NMC0.788mg-CritCRate_20260314155506_DefaultGroup_37_1.png
```

如果识别不到原始子表名，则回退为：

```text
sheet名_part1.png
sheet名_part2.png
```

---

## 7. 现在默认图长什么样？

当前版本已经做过一轮美化，默认风格是：

- 同一个循环的**充电/放电同色**
- 不同循环之间按 `colormap` **渐变**
- charge 和 discharge 用不同线型区分
- 图例改成更紧凑的形式
- 线条比之前更清楚
- 更适合论文初稿、组会、汇报

---

## 7.1 现在可以选哪些图？

当前已经支持下面几类常见图：

- `voltage_capacity`：最常用的 GCD 充放电曲线
- `long_cycling`：长循环图（容量 + 库仑效率）
- `rate_capability`：倍率性能图
- `dqdv`：dQ/dV 曲线
- `dvdq`：dV/dQ 曲线

如果你用 GUI：

- 直接在主界面的 **Plot Types** 区域勾选即可

如果你用命令行：

```powershell
python main.py --input "D:\data\cell.xlsx" --output "D:\figures\curve.png" --plot-types voltage_capacity long_cycling dqdv
```

如果想一次性全部输出：

```powershell
python main.py --input "D:\data\cell.xlsx" --output "D:\figures\curve.png" --plot-types all
```

程序会自动在文件名后面加后缀，例如：

- `_gcd`
- `_cycling`
- `_rate`
- `_dqdv`
- `_dvdq`

---

## 8. colormap 和统一颜色是什么意思？

这两个最容易混淆。

### colormap

当你开启“按循环着色”时：

- 每个循环会从一个颜色方案里自动取色
- 不同循环会呈现渐变或分组配色

当前默认比较推荐：

- `viridis`
- `plasma`
- `cividis`
- `tab10`

### 统一颜色

当你关闭“按循环着色”时：

- 所有曲线都用同一个颜色

### 怎么理解最简单？

- 想让不同循环颜色不一样 → 看 `colormap`
- 想让所有线都一个颜色 → 看“统一颜色”

---

## 9. GUI 里我最常用的参数有哪些？

### 基础参数

- 输入 Excel
- 输出目录
- 基础文件名
- 输出格式
- 绘图后端
- 循环编号

### 高级参数

- DPI
- 图例位置
- colormap
- 统一颜色
- 线宽
- 图尺寸
- X/Y 范围
- 字体
- 是否显示图例
- 是否显示网格
- 是否按循环着色
- 是否保存 Origin 工程

---

## 10. Matplotlib 和 Origin 有什么区别？

### Matplotlib

优点：

- 运行快
- 安装简单
- 适合批量出图
- 能直接导出 `png/svg/pdf`

适合：

- 快速看结果
- 做脚本批处理
- 一般科研绘图

### Origin

优点：

- 可以生成 `.opju`
- 之后能在 Origin 里继续编辑
- 适合需要手动微调图的情况

适合：

- 最终排版前还想在 Origin 里改
- 想保存可编辑工程

---

## 11. 如何输出可编辑的 Origin 工程？

如果你电脑上已经安装 Origin，并且 Python 端配置好了 `originpro`：

命令行可以这样：

```powershell
python main.py --input "D:\data\cell.xlsx" --output "D:\figures\curve.png" --backend origin --save-origin-project
```

GUI 里则：

- 后端选择 `origin`
- 勾选“保存 Origin 工程”

运行后你会得到：

```text
curve.png
curve.opju
```

其中：

- `png` 是导出的图片
- `opju` 是后续在 Origin 中继续编辑的工程文件

---

## 12. 为什么我打开 `.opju` 以后应该能看到数据？

当前版本已经专门修复了：

- 工作簿可见性
- 图页可见性
- 保存后打开看起来像空白的问题

现在正常情况下，你打开 `.opju` 应该能看到：

- 含数据的 workbook
- 含曲线的 graph page

---

## 13. 为什么有时候最后一圈没画全？

先不要急着认为是程序漏画。

很多时候真正原因是：

- 源 Excel 中该循环只有 charge
- 没有对应的 discharge 数据

当前版本遇到这种情况时，会给出提示。  
也就是说：**程序会画出它实际读到的数据**。

---

## 14. 推荐的日常设置

如果你只是正常看图，推荐：

- 后端：`matplotlib`
- colormap：`viridis`
- 线宽：`2.2`
- theme：`paper`
- grid：关闭
- 按循环着色：开启

如果你后面还想去 Origin 里手动改：

- 后端：`origin`
- 保存 Origin 工程：开启

---

## 15. 常见命令模板

### 最简单

```powershell
python main.py --input "D:\data\cell.xlsx" --output "D:\figures\curve.png"
```

### 指定循环

```powershell
python main.py --input "D:\data\cell.xlsx" --output "D:\figures\curve.png" --cycles 1 2 5-8
```

### 使用 Origin 并保存工程

```powershell
python main.py --input "D:\data\cell.xlsx" --output "D:\figures\curve.png" --backend origin --save-origin-project
```

### 启动 GUI

```powershell
python main.py --gui
```

---

## 16. 常见问题

### Q1：提示文件不存在

检查：

- 路径是不是写错了
- 文件后缀是不是对的

### Q2：提示没找到 sheet

解决：

- 如果不确定 sheet 名，直接留空让程序自动识别

### Q3：提示没有可绘制数据

可能原因：

- 你筛选的循环范围没有数据
- 这个 sheet 主要是汇总，不是逐点数据
- 工作模式列命名和常见写法差异太大

### Q4：中文显示异常

可以在高级参数中设置字体，例如：

```text
Microsoft YaHei,SimHei,DejaVu Sans
```

### Q5：GUI 中为什么没有显示所有“常用选项”？

当前 GUI 已做过布局修复。  
如果仍显示不全，通常和：

- Windows 缩放比例
- 窗口大小
- 字体缩放

有关。建议优先：

- 拉大窗口
- 切换到高级设置页查看

---

## 17. 如果你只想记住一句话

最稳妥的用法就是：

```powershell
python main.py
```

或者：

```powershell
python gui.py
```

然后把 Excel 选进去，输出路径设好，直接跑。
