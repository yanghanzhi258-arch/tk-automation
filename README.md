# Desktop Excel Cleaner

一个面向 Windows 11 的 Python 自动化项目，用于读取桌面上的 Excel 文件、自动整理数据，并将整理后的 CSV 文件保存回桌面。

## 功能

- 自动定位当前用户桌面目录（兼容 OneDrive 桌面与普通桌面）。
- 读取 `.xlsx`、`.xlsm`、`.xltx`、`.xltm` Excel 文件。
- 自动整理数据：
  - 去除空行、空列。
  - 清理文本字段首尾空格。
  - 规范列名并为空列名生成默认名称。
  - 删除完全重复的数据行。
  - 可选按指定列排序。
- 导出 UTF-8 BOM CSV，方便 Excel 在 Windows 11 上直接打开中文不乱码。
- 支持命令行参数，也支持双击运行的交互式提示。

## 项目结构

```text
.
├── README.md
├── requirements.txt
├── src/
│   └── desktop_excel_cleaner/
│       ├── __init__.py
│       ├── __main__.py
│       ├── cli.py
│       ├── config.py
│       ├── cleaner.py
│       └── desktop.py
└── tests/
    ├── test_cleaner.py
    └── test_desktop.py
```

## 环境要求

- Windows 11
- Python 3.10 或更高版本
- Microsoft Excel 不是必需的；程序直接读取 Excel 文件内容。

## 安装

在项目目录打开 PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## 快速使用

### 方式一：指定文件名

如果 Excel 文件在桌面，例如 `销售数据.xlsx`：

```powershell
python -m desktop_excel_cleaner --input "销售数据.xlsx"
```

程序会在桌面生成类似下面的文件：

```text
销售数据_cleaned.csv
```

### 方式二：指定完整路径

```powershell
python -m desktop_excel_cleaner --input "C:\Users\你的用户名\Desktop\销售数据.xlsx"
```

### 方式三：指定工作表、输出名与排序列

```powershell
python -m desktop_excel_cleaner --input "销售数据.xlsx" --sheet "Sheet1" --output "整理后数据.csv" --sort-by "日期"
```

### 方式四：双击或无参数运行

```powershell
python -m desktop_excel_cleaner
```

程序会列出桌面上的 Excel 文件，并提示输入文件编号或路径。

## 常用参数

| 参数 | 说明 |
| --- | --- |
| `--input` | Excel 文件名或完整路径。只写文件名时默认从桌面读取。 |
| `--sheet` | 工作表名称或索引，默认读取第一个工作表。 |
| `--output` | CSV 输出文件名或完整路径。只写文件名时默认保存到桌面。 |
| `--sort-by` | 按指定列排序。 |
| `--keep-duplicates` | 保留完全重复行，默认会删除重复行。 |

## 开发与测试

```powershell
python -m pip install -r requirements.txt
pytest
```
