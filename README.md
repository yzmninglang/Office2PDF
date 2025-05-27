### **README.md**

---

## 📑 Office2PDF - Word/PPT 转 PDF & PDF 合并工具

**Office2PDF** 是一个多功能的桌面应用程序，支持将 Word 和 PowerPoint 文件批量转换为 PDF，并提供 PDF 合并功能。此外，它还支持书签保留和自定义排序，非常适合日常办公文档处理。

---

### 🚀 功能概述

1. **Word/PPT 转 PDF**
   - 支持 `.doc`, `.docx`, `.ppt`, `.pptx` 文件格式。
   - 批量转换文件夹中的所有文档。
   - 实时进度条显示转换进度。

2. **PDF 合并器**
   - 支持拖动排序 PDF 文件。
   - 右键菜单移除选中文件。
   - 按名称或创建时间排序。
   - 合并时保留原有书签结构。
   - 添加文件名作为主书签。
   - 合并完成后可直接打开文件所在位置。

3. **多选操作**
   - 使用 `Ctrl + A` 全选 PDF 文件。
   - 使用 `Shift + 左键` 或 `Ctrl + 左键` 多选文件。

4. **界面友好**
   - 简洁直观的 GUI 设计。
   - ~~  实时更新状态信息。 ~~
   - 支持右键菜单操作。

---

### 🛠 技术栈

- **Python**: 用于核心逻辑实现。
- **PyQt5**: 构建图形用户界面（GUI）。
- **PyPDF2**: 处理 PDF 合并与书签操作。
- **win32com.client**: 用于 Word 和 PowerPoint 的 COM 接口调用。

---

### 🏁 快速开始

#### 1️⃣ 安装依赖库

确保你已经安装了以下依赖库：

```bash
pip install pywin32 PyQt5 PyPDF2
```

#### 2️⃣ 运行程序

将代码保存为 `pdf-ppt.py`，然后运行以下命令启动程序：

```bash
python pdf-ppt.py
```

#### 3️⃣ 打包为独立 `.exe` 文件

使用 `PyInstaller` 将项目打包成一个独立的 `.exe` 文件：

```bash
pyinstaller --name=office2pdf --onefile --windowed --icon=icon.ico .\pdf-ppt.py
```

- `--name=office2pdf`: 设置生成的 `.exe` 文件名为 `office2pdf.exe`。
- `--onefile`: 打包成单个文件。
- `--windowed`: 隐藏控制台窗口。
- `--icon=icon.ico`: 设置程序图标（确保 `icon.ico` 文件存在）。

打包完成后，会在 `dist/` 目录下生成 `office2pdf.exe`。

---

### 📂 项目目录结构

```
.
├── .gitignore
├── main.py
├── pdf-ppt-jingdui.py
├── pdf-ppt.py          # 主程序文件
├── pdf1jinduiao.py
└── pdfmerge.py
```

- `pdf-ppt.py`: 主程序文件，包含 Word/PPT 转换和 PDF 合并逻辑。
- `pdfmerge.py`: PDF 合并相关功能模块。
- `pdf1jinduiao.py`: 进度条相关功能模块。
- `main.py`: 示例入口文件（可选）。

---

### 🎯 使用说明

#### 1️⃣ Word/PPT 转 PDF

1. 点击 **“选择文件夹（Word/PPT）”**，选择包含 Word 和 PPT 文件的文件夹。
2. 点击 **“转换为 PDF”**，程序会自动将文件夹中的所有 Word 和 PPT 文件转换为 PDF。
3. 转换过程中会显示实时进度条和当前处理的文件名。

#### 2️⃣ PDF 合并

1. 点击 **“添加 PDF 文件”**，选择需要合并的 PDF 文件。
2. 在列表中拖动文件调整顺序，或使用右键菜单移除不需要的文件。
3. 选择排序方式（按名称或创建时间）。
4. 点击 **“合并 PDF”**，程序会将 PDF 文件合并为一个，并保留原有书签结构。
5. 合并完成后，点击弹窗中的 **“打开文件所在位置”**，可以直接跳转到合并后的 PDF 文件。

---

### 🛠 注意事项

1. **依赖 Microsoft Office**：
   - Word/PPT 转换功能依赖于 Microsoft Office 的 COM 接口，因此需要在目标电脑上安装 Microsoft Office。

2. **图标文件**：
   - 如果你需要自定义图标，请确保 `icon.ico` 文件存在于项目根目录，并且符合 Windows 图标规范。

3. **跨平台限制**：
   - 当前程序仅支持 Windows 平台，因为依赖于 `win32com.client`。

4. **权限问题**：
   - 如果遇到权限问题（如无法写入文件），请以管理员身份运行程序。

---

### 🤝 贡献与反馈

如果你有任何建议、发现 bug 或希望添加新功能，请提交 Issue 或 Pull Request。

---

### 📜 版权声明

本项目遵循 [MIT License](LICENSE) 开源协议，欢迎 fork 和贡献。

---


---

感谢使用 **Office2PDF**！希望这个工具能帮助你更高效地处理办公文档 😊
