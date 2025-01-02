# -*- coding: utf-8 -*-
# 作者: MEI
# 版权所有 (c) 2025, MEI. 保留所有权利.
# 本软件遵循 [许可证名称]，可以在 [许可证链接] 上查看详细信息。
# 请勿未经许可使用、复制或修改本代码。import win32com.client
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.constants import *
from tkinter import scrolledtext
from tkinter import ttk  # 导入 ttk 模块
import sys
import traceback
# 可用主题列表
themes = [
    'superhero', 'flatly', 'cyborg', 'slate', 'darkly', 
    'lumen', 'lux', 'pulse', 'sandstone', 'united', 
    'paper', 'yeti', 'minty', 'solar', 'cosmo'
]

def change_theme(event):
    # 获取用户选择的主题
    selected_theme = theme_combobox.get()
    
    # 切换主题
    style.theme_use(selected_theme)

# 创建主窗口
root = tk.Tk()
root.title("选择主题")
root.geometry("400x300")

# 创建 ttkbootstrap 样式
style = Style(theme='superhero')  # 默认主题

# 添加主题选择标签
theme_label = ttk.Label(root, text="选择一个主题", font=('Helvetica', 14))
theme_label.pack(pady=20)

# 创建下拉菜单供用户选择主题
theme_combobox = ttk.Combobox(root, values=themes, state="readonly", font=('Helvetica', 12))
theme_combobox.set('superhero')  # 默认选择的主题
theme_combobox.pack(pady=10)

# 当用户选择不同的主题时调用 change_theme 函数
theme_combobox.bind("<<ComboboxSelected>>", change_theme)
def extract_student_name(doc, pattern):
    """
    从 Word 文档中提取名称。
    
    :param doc: Word 文档对象
    :param pattern: 正则表达式模式
    :return: 提取到的名称或None
    """
    try:
        content = doc.Content.Text
        match = re.search(pattern, content)
        if match:
            return match.group(1).strip()
        else:
            return None
    except Exception as e:
        return None

def make_safe_filename(name):
    """
    将字符串转换为合法的文件名。
    
    :param name: 原始名字
    :return: 清理后的名字
    """
    return re.sub(r'[\\/*?:"<>|]', "_", name)

def generate_regex(prefix, suffix):
    """
    根据用户输入的前缀和后缀生成正则表达式。
    
    :param prefix: 前缀文本
    :param suffix: 后缀文本
    :return: 生成的正则表达式模式
    """
    # 转义前缀和后缀中的正则特殊字符
    escaped_prefix = re.escape(prefix)
    escaped_suffix = re.escape(suffix)
    
    # 在前缀和后缀与名称之间添加\s*，匹配零个或多个空白字符
    regex_pattern = rf"{escaped_prefix}\s*([\u4e00-\u9fa5]+)\s*{escaped_suffix}"
    
    return regex_pattern

def rename_documents(input_dir, pattern, output_log):
    """
    遍历指定目录中的所有 Word 文档，提取名称并重命名文件。
    
    :param input_dir: 输入文件夹路径
    :param pattern: 用于提取名称的正则表达式模式
    :param output_log: 用于显示日志信息的文本框
    """
    # 初始化 Word 应用
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 设置为 False 以在后台运行
        word.DisplayAlerts = 0  # wdAlertsNone
    except Exception as e:
        messagebox.showerror("错误", f"无法初始化 Word 应用程序: {e}")
        return
    
    try:
        # 支持的文件扩展名
        supported_extensions = [".doc", ".docx"]
        
        # 获取所有支持的文件
        files = [f for f in os.listdir(input_dir) if os.path.splitext(f)[1].lower() in supported_extensions]
        
        if not files:
            messagebox.showinfo("信息", f"在目录 '{input_dir}' 中未找到任何支持的Word文档（.doc, .docx）。")
            return
        
        # 跟踪已使用的文件名以避免重复
        used_names = {}
        
        for filename in files:
            file_path = os.path.join(input_dir, filename)
            output_log.config(state='normal')
            output_log.insert(tk.END, f"正在处理文件: {file_path}\n")
            output_log.see(tk.END)
            output_log.config(state='disabled')
            try:
                # 打开文档
                doc = word.Documents.Open(FileName=file_path, ReadOnly=True, AddToRecentFiles=False)
                
                # 提取名称
                name = extract_student_name(doc, pattern)
                if name:
                    safe_name = make_safe_filename(name)
                    if not safe_name:
                        safe_name = os.path.splitext(filename)[0]
                        output_log.config(state='normal')
                        output_log.insert(tk.END, f"提取到的名称为空，使用原文件名。\n")
                        output_log.see(tk.END)
                        output_log.config(state='disabled')
                else:
                    safe_name = os.path.splitext(filename)[0]
                    output_log.config(state='normal')
                    output_log.insert(tk.END, f"未在文件 '{filename}' 中找到匹配的名称，使用原文件名。\n")
                    output_log.see(tk.END)
                    output_log.config(state='disabled')
                
                # 处理重复名称
                original_safe_name = safe_name
                count = 1
                while safe_name in used_names:
                    safe_name = f"{original_safe_name}_{count}"
                    count += 1
                used_names[safe_name] = True
                
                # 新文件名
                new_extension = os.path.splitext(filename)[1]
                new_filename = f"{safe_name}{new_extension}"
                new_file_path = os.path.join(input_dir, new_filename)
                
                # 关闭文档
                doc.Close(False)
                
                # 重命名文件
                if os.path.exists(new_file_path):
                    output_log.config(state='normal')
                    output_log.insert(tk.END, f"目标文件名 '{new_filename}' 已存在，跳过重命名。\n")
                    output_log.see(tk.END)
                    output_log.config(state='disabled')
                else:
                    os.rename(file_path, new_file_path)
                    output_log.config(state='normal')
                    output_log.insert(tk.END, f"已重命名为: {new_filename}\n")
                    output_log.see(tk.END)
                    output_log.config(state='disabled')
            except Exception as e:
                output_log.config(state='normal')
                output_log.insert(tk.END, f"处理文件 '{filename}' 时发生错误: {e}\n")
                output_log.see(tk.END)
                output_log.config(state='disabled')
                continue
                
        messagebox.showinfo("完成", "所有文件处理完成。")
                
    finally:
        # 退出 Word 应用
        word.Quit()

def browse_directory():
    """
    打开目录选择对话框并设置路径。
    """
    directory = filedialog.askdirectory()
    if directory:
        dir_path.set(directory)

def validate_regex(pattern):
    """
    验证用户输入的正则表达式是否有效。
    
    :param pattern: 用户输入的正则表达式
    :return: True 如果有效，False 否则
    """
    try:
        re.compile(pattern)
        return True
    except re.error:
        return False

def show_help():
    """
    显示正则表达式帮助信息。
    """
    help_text = (
        "正则表达式帮助:\n\n"
        "1. 工具会根据您输入的前缀和后缀自动生成正则表达式，用于提取中间的名称部分。\n"
        "2. 例如，假设文件内容中有“下面将（张三）同学”，您可以输入：\n"
        "   - 前缀文本：下面将（\n"
        "   - 后缀文本：）同学\n"
        "   生成的正则表达式将是：下面将\\s*([\\u4e00-\\u9fa5]+)\\s*）同学\n\n"
        "   解释：\n"
        "   - `下面将（`：匹配文本“下面将（”。\n"
        "   - `\\s*`：匹配零个或多个空白字符（如空格）。\n"
        "   - `([\\u4e00-\\u9fa5]+)`：捕获组，匹配一个或多个中文字符。\n"
        "   - `）同学`：匹配文本“）同学”。\n\n"
        "3. `\\s*` 用于匹配前缀和名称、名称和后缀之间可能存在的零个或多个空白字符（如空格）。\n"
        "4. 确保前缀和后缀文本与文档中的实际文本一致，包括使用的括号类型（全角或半角）。\n"
        "5. 如果前后文本中包含特殊正则字符（如 `.`、`*`、`?` 等），工具会自动对其进行转义，确保正则表达式的正确性。\n"
        "6. 您可以使用在线工具（如 https://regex101.com/ ）来测试和调试生成的正则表达式。"
    )
    messagebox.showinfo("正则表达式帮助", help_text)

def generate_regex_pattern():
    """
    根据前缀和后缀文本生成正则表达式模式。
    """
    try:
        prefix = prefix_entry.get()
        suffix = suffix_entry.get()
        
        if not prefix:
            messagebox.showwarning("警告", "请先输入前缀文本。")
            return
        if not suffix:
            messagebox.showwarning("警告", "请先输入后缀文本。")
            return
        
        pattern = generate_regex(prefix, suffix)
        regex_var.set(pattern)
        
        # 在日志中显示生成的正则表达式
        output_log.config(state='normal')
        output_log.insert(tk.END, f"生成的正则表达式: {pattern}\n")
        output_log.see(tk.END)
        output_log.config(state='disabled')
        
        messagebox.showinfo("转换成功", "已根据前缀和后缀文本生成正则表达式。")
    except Exception as e:
        messagebox.showerror("错误", f"生成正则表达式时发生错误: {e}")

def start_renaming():
    """
    开始重命名过程。
    """
    try:
        directory = dir_path.get()
        prefix = prefix_entry.get()
        suffix = suffix_entry.get()
        pattern = regex_var.get()
        
        if not directory:
            messagebox.showwarning("警告", "请先选择要处理的目录。")
            return
        if not prefix:
            messagebox.showwarning("警告", "请先输入前缀文本。")
            return
        if not suffix:
            messagebox.showwarning("警告", "请先输入后缀文本。")
            return
        if not pattern:
            messagebox.showwarning("警告", "请先生成正则表达式模式。")
            return
        if not validate_regex(pattern):
            messagebox.showerror("错误", "生成的正则表达式无效。请检查前缀和后缀文本是否正确。")
            return
        
        # 清空日志
        output_log.config(state='normal')
        output_log.delete(1.0, tk.END)
        output_log.config(state='disabled')
        
        rename_documents(directory, pattern, output_log)
    except Exception as e:
        messagebox.showerror("错误", f"开始重命名时发生错误: {e}")
def main():
    try:
        # 创建主窗口并应用现代主题
        style = Style(theme='flatly')  # 选择一个现代主题，可根据需要更改
        
        root = style.master
        root.title("Word 文档批量重命名工具 By：MEI")
        root.geometry("1000x1000")
        root.resizable(False, False)
        
        # 设置全局变量
        global dir_path, regex_var, prefix_entry, suffix_entry, output_log
        dir_path = tk.StringVar()
        regex_var = tk.StringVar()
        
        # 创建并放置组件
        # 使用 ttk.Frame 而不是 style.Frame
        frame = ttk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 目录选择
        dir_label = ttk.Label(frame, text="选择要重命名的目录：", font=('Helvetica', 12))
        dir_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        dir_entry = ttk.Entry(frame, textvariable=dir_path, width=60, font=('Helvetica', 12))
        dir_entry.grid(row=0, column=1, padx=5, pady=(0, 10))
        
        browse_button = ttk.Button(frame, text="浏览", command=browse_directory, style='primary.TButton')
        browse_button.grid(row=0, column=2, padx=5, pady=(0, 10))
        
        # 前缀文本输入
        prefix_label = ttk.Label(frame, text="前缀文本：", font=('Helvetica', 12))
        prefix_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 10))
        
        prefix_entry = ttk.Entry(frame, width=60, font=('Helvetica', 12))
        prefix_entry.grid(row=1, column=1, padx=5, pady=(0, 10), columnspan=2, sticky=tk.W)
        
        # 后缀文本输入
        suffix_label = ttk.Label(frame, text="后缀文本：", font=('Helvetica', 12))
        suffix_label.grid(row=2, column=0, sticky=tk.W, pady=(0, 10))
        
        suffix_entry = ttk.Entry(frame, width=60, font=('Helvetica', 12))
        suffix_entry.grid(row=2, column=1, padx=5, pady=(0, 10), columnspan=2, sticky=tk.W)
        
        # 生成正则表达式按钮
        generate_button = ttk.Button(frame, text="生成正则表达式", command=generate_regex_pattern, style='info.TButton')
        generate_button.grid(row=3, column=1, padx=5, pady=(0, 10), sticky=tk.W)
        
        # 正则表达式显示
        regex_label = ttk.Label(frame, text="生成的正则表达式模式：", font=('Helvetica', 12))
        regex_label.grid(row=4, column=0, sticky=tk.W, pady=(0, 10))
        
        regex_entry = ttk.Entry(frame, textvariable=regex_var, width=60, font=('Helvetica', 12))
        regex_entry.grid(row=4, column=1, padx=5, pady=(0, 10), columnspan=2, sticky=tk.W)
        
        # 帮助按钮
        help_button = ttk.Button(frame, text="帮助", command=show_help, style='secondary.TButton')
        help_button.grid(row=5, column=2, padx=5, pady=(0, 10), sticky=tk.E)
        
        # 示例文本
        example_label = ttk.Label(frame, text='示例：\n前缀文本: "下面将"\n后缀文本: "同学"', font=('Helvetica', 10), foreground='gray')
        example_label.grid(row=6, column=1, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # 开始重命名按钮
        start_button = ttk.Button(frame, text="开始重命名", command=start_renaming, style='success.TButton', width=20)
        start_button.grid(row=7, column=1, pady=20)
        
        # 日志显示
        log_label = ttk.Label(frame, text="处理日志：", font=('Helvetica', 12))
        log_label.grid(row=8, column=0, sticky=tk.W, pady=(0, 10))
        
        output_log = scrolledtext.ScrolledText(frame, width=80, height=25, state='normal', font=('Helvetica', 10))
        output_log.grid(row=9, column=0, columnspan=3, pady=10)
        
        # 运行主循环
        root.mainloop()
    except Exception as e:
        error_message = ''.join(traceback.format_exception(None, e, e.__traceback__))
        messagebox.showerror("程序错误", f"程序发生未捕获的错误:\n{error_message}")
        sys.exit(1)

if __name__ == '__main__':
    main()
