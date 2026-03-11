# 电池充放电曲线绘图工具用户手册

这是一份给初学者准备的“傻瓜式”操作说明。

如果你以前几乎没怎么用过 Python，也可以照着一步一步操作。

---

## 1. 这个工具是干什么的？

这个工具用来做下面这件事：

- 读取你的电池测试 Excel 数据文件
- 自动找到真正的逐点曲线数据
- 自动识别充电和放电数据
- 画出“电压 - 比容量”曲线图
- 把图片保存成 `png`、`svg` 或 `pdf`

你最常见的用途就是：

- 输入一个 `.xlsx` 文件
- 输出一张充放电曲线图

---

## 2. 你需要准备什么？

在使用前，你只需要准备两样东西：

### 2.1 项目文件夹

也就是这个目录：

```text
D:\PythonProject\cycle
```

### 2.2 你的测试数据文件

例如：

```text
D:\PythonProject\cycle\Export20260310203118.xlsx
```

注意：

- 推荐使用 `.xlsx`
- 如果你的原始文件是 `.ccs`，请先在测试设备软件里导出成 `.xlsx`

---

## 3. 第一次使用前，要做什么？

如果这是你第一次在这台电脑上运行这个项目，请按下面步骤做一次准备。

### 第 1 步：打开 PowerShell

你可以这样打开：

- 按键盘 `Win`
- 输入 `PowerShell`
- 打开“Windows PowerShell”

### 第 2 步：进入项目目录

在 PowerShell 里输入：

```powershell
cd D:\PythonProject\cycle
```

然后按回车。

### 第 3 步：激活虚拟环境

输入：

```powershell
.\.venv\Scripts\activate
```

然后按回车。

如果成功，你会看到命令行前面通常多了一个类似：

```text
(.venv)
```

这表示虚拟环境已经激活。

### 第 4 步：安装依赖

输入：

```powershell
pip install -r requirements.txt
```

然后按回车，等安装完成。

如果你之前已经安装过，就不用每次都安装。

---

## 4. 以后每次最简单的使用方法

以后你最常用的方式就是这 3 步：

### 第 1 步

打开 PowerShell

### 第 2 步

进入项目目录并激活环境：

```powershell
cd D:\PythonProject\cycle
.\.venv\Scripts\activate
```

### 第 3 步

运行程序：

```powershell
python main.py
```

这就是默认推荐方式。

程序启动后，会一步一步问你要输入什么，你照着填就行。

---

## 5. 最推荐的运行方式：交互式运行

### 5.1 什么叫交互式运行？

就是你只需要输入：

```powershell
python main.py
```

程序就会自己问你：

- 数据文件在哪里
- 输出图片保存到哪里
- 要画哪些循环
- 要不要图例
- 标题写什么

这种方式最适合新手。

### 5.2 怎么运行？

在 PowerShell 中输入：

```powershell
cd D:\PythonProject\cycle
.\.venv\Scripts\activate
python main.py
```

然后按回车。

---

## 6. 交互式运行时，每一步该怎么填？

下面我用最常见的情况举例。

假设你的数据文件是：

```text
D:\PythonProject\cycle\Export20260310203118.xlsx
```

你想把图片保存成：

```text
D:\PythonProject\cycle\my_curve.png
```

那么运行 `python main.py` 后，你会看到类似提示。

### 提示 1：请输入数据文件路径

你输入：

```text
D:\PythonProject\cycle\Export20260310203118.xlsx
```

说明：

- 可以直接复制粘贴
- 带引号也没关系，例如：

```text
"D:\PythonProject\cycle\Export20260310203118.xlsx"
```

程序会自动处理。

### 提示 2：请输入 sheet 名（直接回车则自动识别）

如果你不知道 sheet 名，**直接按回车** 就行。

程序会自动帮你找最可能包含逐点数据的 sheet。

如果你明确知道 sheet 名，也可以输入，例如：

```text
DefaultGroup_33_4
```

### 提示 3：请输入输出图片保存路径

你输入：

```text
D:\PythonProject\cycle\my_curve.png
```

说明：

- 这是图片保存的位置
- 你可以改成别的文件名
- 也可以改成 `svg` 或 `pdf`

例如：

```text
D:\PythonProject\cycle\my_curve.svg
```

如果你只写：

```text
D:\PythonProject\cycle\my_curve
```

程序会自动补后缀。

### 提示 4：请输入要绘制的循环编号，例如 1,2,3 或 1-5（直接回车表示全部循环）

这里有几种常见写法：

- 直接回车：画全部循环
- 输入 `1,2,3`：只画第 1、2、3 圈
- 输入 `1-5`：画第 1 到第 5 圈
- 输入 `1,2,5-8`：混合选择

如果你暂时不确定，建议先直接回车，先看全部循环效果。

### 提示 5：是否显示图例（y/n）

常用输入：

- 输入 `y`：显示图例
- 输入 `n`：不显示图例

一般建议输入：

```text
y
```

### 提示 6：请输入标题（直接回车使用默认标题）

你可以输入你自己的标题，例如：

```text
Solid-state battery charge-discharge curves
```

如果你懒得改，直接回车也可以。

### 提示 7：是否进入高级参数设置（y/n，直接回车默认 n）

如果你只是想快速出图，输入：

```text
n
```

如果你想改：

- 线宽
- 颜色
- 图尺寸
- 字体
- 坐标范围
- 透明背景

那就输入：

```text
y
```

---

## 7. 给你一个可以直接照抄的完整示例

### 7.1 运行命令

```powershell
cd D:\PythonProject\cycle
.\.venv\Scripts\activate
python main.py
```

### 7.2 假设你按下面这样输入

```text
请输入数据文件路径：
D:\PythonProject\cycle\Export20260310203118.xlsx

请输入 sheet 名（直接回车则自动识别）：

请输入输出图片保存路径：
D:\PythonProject\cycle\result.png

请输入要绘制的循环编号，例如 1,2,3 或 1-5（直接回车表示全部循环）：

是否显示图例（y/n）：
y

请输入标题（直接回车使用默认标题）：
我的电池充放电曲线

是否进入高级参数设置（y/n，直接回车默认 n）：
n
```

### 7.3 运行结束后

程序会提示：

```text
图片已保存到 D:\PythonProject\cycle\result.png
```

这就表示成功了。

然后你去这个路径下，就能找到图片。

---

## 8. 如果你不想一步一步输入，也可以直接一条命令运行

这叫“命令行参数方式”。

适合你以后熟悉了之后使用。

### 最简单示例

```powershell
python main.py --input "D:\PythonProject\cycle\Export20260310203118.xlsx" --output "D:\PythonProject\cycle\curve.png"
```

### 指定循环

只画第 1 到第 5 圈：

```powershell
python main.py --input "D:\PythonProject\cycle\Export20260310203118.xlsx" --output "D:\PythonProject\cycle\curve.png" --cycles 1-5
```

只画第 1、2、10 圈：

```powershell
python main.py --input "D:\PythonProject\cycle\Export20260310203118.xlsx" --output "D:\PythonProject\cycle\curve.png" --cycles 1 2 10
```

### 指定输出为 svg

```powershell
python main.py --input "D:\PythonProject\cycle\Export20260310203118.xlsx" --output "D:\PythonProject\cycle\curve.svg"
```

### 指定输出为 pdf

```powershell
python main.py --input "D:\PythonProject\cycle\Export20260310203118.xlsx" --output "D:\PythonProject\cycle\curve.pdf"
```

---

## 9. 高级参数设置怎么理解？

如果你在交互式运行时选择进入高级参数设置，程序会继续问你更多内容。

下面是最常见的几个：

### 9.1 输出格式

可选：

- `png`
- `svg`
- `pdf`

建议：

- 日常查看：`png`
- 后期编辑：`svg`
- 论文排版：`pdf`

### 9.2 DPI

就是清晰度。

常用建议：

- `300`：一般够用
- `600`：更清晰，适合论文

### 9.3 图例位置

常见写法：

- `best`
- `upper right`
- `upper left`
- `lower right`
- `lower left`

如果你不懂，保持默认 `best` 就行。

### 9.4 线宽

如果曲线看起来太细，可以调大一点，比如：

```text
2.0
```

### 9.5 颜色

可以输入颜色代码，例如：

```text
#1f77b4
```

或者你让程序按循环自动分色。

### 9.6 图尺寸

例如输入：

```text
8,6
```

表示宽 8 英寸，高 6 英寸。

### 9.7 坐标范围

如果你想手动设置横轴或纵轴范围，可以输入：

```text
0,200
```

或者：

```text
2.5,4.5
```

如果不需要，直接回车即可。

### 9.8 透明背景

如果你想把图放到 PPT、论文或者其他图中，有时透明背景会更方便。

可选：

- `y`：透明背景
- `n`：普通白色背景

---

## 10. 输出图片保存在哪里？

图片保存在哪里，完全取决于你输入的“输出图片保存路径”。

例如你输入：

```text
D:\PythonProject\cycle\result.png
```

那图片就会保存到：

```text
D:\PythonProject\cycle\result.png
```

如果你输入：

```text
D:\MyFigures\battery\fig1.svg
```

那图片就会保存到：

```text
D:\MyFigures\battery\fig1.svg
```

如果文件夹不存在，程序会自动创建。

---

## 11. 如何只画我想要的几个循环？

你可以在交互输入时或者命令行里指定循环。

### 写法 1：画全部循环

直接回车，不输入任何东西。

### 写法 2：画几个指定循环

输入：

```text
1,2,5,10
```

表示只画第 1、2、5、10 圈。

### 写法 3：画一个连续区间

输入：

```text
1-5
```

表示画第 1 到第 5 圈。

### 写法 4：混合写法

输入：

```text
1,2,5-8,10
```

表示画：

- 第 1 圈
- 第 2 圈
- 第 5 到第 8 圈
- 第 10 圈

---

## 12. 程序会自动帮你做哪些事？

这个工具不是让你自己手工处理 Excel，而是会自动帮你做很多事：

- 自动识别最合适的 sheet
- 自动寻找真正的逐点数据表头
- 自动跳过前面的汇总区
- 自动识别关键列
- 自动识别充电/放电/静置
- 自动忽略静置数据
- 自动清理空值和异常值
- 自动创建输出目录
- 自动补全输出文件后缀

所以通常你不需要自己先整理 Excel。

---

## 13. 常见问题与解决办法

### 问题 1：提示“文件不存在”

原因：

- 你输入的路径不对
- 文件名打错了
- 文件不在这个位置

解决：

- 最好直接去资源管理器里复制完整路径
- 注意文件后缀是不是 `.xlsx`

### 问题 2：提示没找到 sheet

原因：

- 你手动输入的 sheet 名不对

解决：

- 不知道 sheet 名时，直接回车，让程序自动识别

### 问题 3：程序说没有找到可绘制数据

原因可能是：

- 你选的循环没有数据
- 这个 sheet 主要是汇总，不是逐点数据
- 工作模式命名比较特殊，程序没识别出来

解决办法：

1. 先不要筛选循环，直接画全部
2. 换一个 sheet 试试
3. 确认 Excel 中确实有 `电压/V` 和 `比容量/mAh/g`

### 问题 4：中文乱码或显示成方框

原因：

- 系统没有合适的中文字体

解决：

- 在高级设置里指定字体，例如：

```text
Microsoft YaHei,SimHei,DejaVu Sans
```

### 问题 5：原始文件是 `.ccs` 怎么办？

解决：

- 先去设备软件中把 `.ccs` 导出成 `.xlsx`
- 再用本工具读取 `.xlsx`

---

## 14. 最适合新手记住的最短流程

如果你不想看太多内容，只记住下面这几步就够了。

### 每次使用：

```powershell
cd D:\PythonProject\cycle
.\.venv\Scripts\activate
python main.py
```

然后按提示输入：

1. Excel 文件路径
2. sheet 名（不会就直接回车）
3. 输出图片路径
4. 循环编号（不会就直接回车）
5. 图例（一般输入 `y`）
6. 标题（不会就回车）
7. 高级设置（新手一般输入 `n`）

最后看到：

```text
图片已保存到 xxx
```

就表示成功。

---

## 15. 给你的建议

如果你是第一次用，建议你这样做：

### 第一次

- 用 `python main.py`
- 按默认方式一步一步输入
- 先成功生成一张图

### 第二次以后

- 如果你经常处理相同类型文件
- 再开始尝试命令行参数方式

这样最不容易出错。

---

## 16. 一个最适合直接照抄的版本

以后你完全可以直接复制下面这段：

```powershell
cd D:\PythonProject\cycle
.\.venv\Scripts\activate
python main.py
```

然后按下面思路填写：

```text
数据文件路径：你的 xlsx 完整路径
sheet 名：不会就直接回车
输出路径：你想保存图片的位置，例如 D:\PythonProject\cycle\result.png
循环编号：不会就直接回车
图例：y
标题：不会就直接回车
高级设置：n
```

这就是最稳妥的用法。
