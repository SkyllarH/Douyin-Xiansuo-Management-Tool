# 抖音线索整理&合并工具

一个用于批量处理和整理抖音线索Excel文件的Python应用程序，支持自动识别格式、提取数据、查重去重等功能。

## 功能特性

- **批量整理线索**：支持处理多个Excel文件，自动识别格式并提取数据
- **合并整理**：将多个Excel文件合并为一个，并自动去重
- **智能查重**：基于手机号和微信号的智能查重去重
- **本地线索判断**：可自定义本地线索判断逻辑（默认为"广东"）
- **Excel元素居中**：导出的Excel文件中所有元素自动居中
- **现代简约界面**：美观的现代简约风格界面
- **自动时间戳**：输出文件名自动添加时间戳
- **打开文件提示**：处理完成后可选择直接打开生成的文件

## 安装依赖

1. 确保已安装Python 3.6或更高版本
2. 安装所需依赖：

```bash
pip install -r requirements.txt
```

## 运行应用

```bash
python qt_excel_tool.py
```

## 打包为可执行文件

### Windows (exe)

1. 安装PyInstaller：

```bash
pip install pyinstaller
```

2. 打包应用：

```bash
pyinstaller --onefile --windowed --name 抖音线索工具 qt_excel_tool.py
```

3. 可执行文件将在`dist`目录中生成

### macOS (dmg)

1. 安装PyInstaller：

```bash
pip install pyinstaller
```

2. 打包应用：

```bash
pyinstaller --onefile --windowed --name 抖音线索工具 qt_excel_tool.py
```

3. 使用`create-dmg`工具创建dmg文件（需要先安装）：

```bash
npm install -g create-dmg
create-dmg dist/抖音线索工具.app
```

## 使用指南

1. **选择文件**：点击"选择文件"按钮选择单个或多个Excel文件，或点击"选择文件夹"按钮选择包含Excel文件的文件夹
2. **设置选项**：
   - 输出目录：选择处理后文件的保存位置
   - 主播名称：设置默认主播名称（默认为"晁天娇"）
   - 本地线索：设置本地线索判断关键词（默认为"广东"）
   - 删除空行：选择是否删除空行
   - 线索去重：选择是否对线索进行去重
3. **执行操作**：
   - 点击"整理线索"按钮处理选中的文件
   - 点击"合并整理"按钮将多个文件合并为一个
4. **查看结果**：处理结果将显示在"处理结果"区域
5. **打开文件**：处理完成后会提示是否打开生成的文件

## 技术栈

- Python 3.6+
- PyQt5 (GUI)
- pandas (数据处理)
- openpyxl, xlrd (Excel文件处理)
- PyInstaller (打包工具)

## 作者信息

**作者**：MirrorLab
**版权**：Copyright © 2026 MirrorLab. All rights reserved.

## 注意事项

- 支持的Excel格式：.xlsx, .xls
- 自动识别两种格式：包含"用户信息"和"城市"的格式，以及包含"抖音昵称"和"询价城市"的格式
- 查重规则：手机号或微信号重复时，仅保留第一条数据
