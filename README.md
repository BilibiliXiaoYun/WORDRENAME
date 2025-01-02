# Word Rename Tool

![Icon](./path_to_your_icon.png) <!-- 替换为图标路径 -->

## 概述

**Word Rename Tool** 是一个基于 Python 的工具，用于批量重命名 Word 文档文件。该工具结合了现代化的 GUI 界面和强大的正则表达式功能，能够快速高效地提取文档内容中的特定信息并自动重命名文件。

---

## 特性

- **批量处理**：支持一次性处理多个 Word 文档。
- **正则提取**：根据用户输入的前缀和后缀，智能提取文档中的关键内容作为文件名。
- **界面友好**：基于 `Tkinter` 和 `ttkbootstrap` 的现代化 GUI，易于操作。
- **完全可定制**：用户可以选择目录、正则表达式规则，以及自定义文件名格式。
- **格式支持**：支持 `.doc` 和 `.docx` 文件。

---

## 使用说明

### 环境要求

1. Python 3.8 或以上版本。
2. 已安装以下依赖库：
   - `ttkbootstrap`
   - `pywin32`

   **安装依赖库**：
   ```bash
   pip install ttkbootstrap pywin32
