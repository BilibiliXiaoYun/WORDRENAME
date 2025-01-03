# -*- coding: utf-8 -*-
# 作者: MEI
# 版权所有 (c) 2025, MEI. 保留所有权利.
# 本软件遵循 MIT，可以在 GITHUB 上查看详细信息。
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
import win32com.client
import json
import os
import win32com.client
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.constants import *
from tkinter import scrolledtext
from tkinter import ttk
import sys
import traceback
import tkinter as tk
from tkinter import ttk, messagebox
from ttkbootstrap import Style

def show_about():
    """
    显示关于框
    """
    messagebox.showinfo(
        "关于本软件",
        "Word 文档重命名工具\n\n"
        "版本：1.2\n"
        "作者：MEI\n"
        "许可证：MIT License\n\n"
        "感谢您使用本工具！"
    )

def main():
    style = Style(theme='flatly')
    root = style.master
    root.title("Word 文档重命名工具")
    root.geometry("1000x1000")

    # 创建菜单栏
    menu_bar = tk.Menu(root)
    root.config(menu=menu_bar)

    # 创建“帮助”菜单
    help_menu = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="帮助", menu=help_menu)

    # 在帮助菜单中添加“关于”选项
    help_menu.add_command(label="关于", command=show_about)

    # 主界面内容
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 示例主界面组件
    label = ttk.Label(frame, text="欢迎使用 Word 文档重命名工具", font=("Helvetica", 14))
    label.pack(pady=20)


if __name__ == '__main__':
    main()

history_file = "rename_history.json"  # 用于存储重命名历史的文件

# 使用脚本所在目录作为基础路径
script_dir = os.path.dirname(os.path.abspath(__file__))  # 获取当前脚本所在的目录
history_file = os.path.join(script_dir, "rename_history.json")  # 在脚本目录下创建 rename_history.json

# 初始化重命名历史文件
if not os.path.exists(history_file):
    with open(history_file, "w") as f:
        json.dump([], f)
# 读取历史记录文件，如果文件为空或损坏，则初始化为 []
try:
    with open(history_file, "r") as f:
        history = json.load(f)  # 尝试读取历史记录
except (json.JSONDecodeError, FileNotFoundError):
    # 如果文件为空或损坏，重新初始化为 [] 
    history = []
    with open(history_file, "w") as f:
        json.dump(history, f)  # 写入一个空的 JSON 列表
def undo_rename(output_log):
    """
    撤销最近的一次重命名操作。
    """
    if not os.path.exists(history_file):
        messagebox.showinfo("提示", "没有可撤销的操作！")
        return

    try:
        with open(history_file, "r+") as f:
            history = json.load(f)
            if not history:
                messagebox.showinfo("提示", "没有可撤销的操作！")
                return

            # 撤销操作
            for record in reversed(history):
                old_path = record["old"]
                new_path = record["new"]

                if os.path.exists(new_path):
                    os.rename(new_path, old_path)  # 恢复文件名
                    output_log.config(state='normal')
                    output_log.insert(tk.END, f"已撤销文件: {new_path} -> {old_path}\n")
                    output_log.see(tk.END)
                    output_log.config(state='disabled')
                else:
                    output_log.config(state='normal')
                    output_log.insert(tk.END, f"文件不存在，无法撤销: {new_path}\n")
                    output_log.see(tk.END)
                    output_log.config(state='disabled')

            # 清空历史记录
            f.seek(0)
            json.dump([], f, indent=4)  # 清空历史文件
            f.truncate()

        messagebox.showinfo("完成", "撤销操作完成！")
    except Exception as e:
        messagebox.showerror("错误", f"撤销失败：{e}")

try:
    word = win32com.client.Dispatch("Word.Application")
    print("Win32com 已成功运行！")
except Exception as e:
    print(f"错误: {e}")

# 可用主题列表
themes = [
    'superhero', 'flatly', 'cyborg', 'slate', 'darkly', 
    'lumen', 'lux', 'pulse', 'sandstone', 'united', 
    'paper', 'yeti', 'minty', 'solar', 'cosmo'
]
icon_path = os.path.join(os.path.dirname(__file__), "my_icon.ico")

# 检查自定义图标是否存在，如果不存在则使用 Python 默认图标
if os.path.exists(icon_path):
    root.iconbitmap(icon_path)  # 使用自定义图标
else:
    print("未找到自定义图标，将使用 Python 默认图标")  # 打印提示信息

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

# 当用户选择不同的主题时调用 change_theme 函数d

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
    """
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
    except Exception as e:
        messagebox.showerror("错误", f"无法初始化 Word 应用程序: {e}")
        return

    try:
        supported_extensions = [".doc", ".docx"]
        files = [f for f in os.listdir(input_dir) if os.path.splitext(f)[1].lower() in supported_extensions]
        
        if not files:
            messagebox.showinfo("信息", f"在目录 '{input_dir}' 中未找到任何支持的Word文档（.doc, .docx）。")
            return
        
        rename_history = []  # 用于记录重命名历史
        
        for filename in files:
            file_path = os.path.abspath(os.path.join(input_dir, filename))  # 使用完整路径
            output_log.config(state='normal')
            output_log.insert(tk.END, f"正在处理文件: {file_path}\n")
            output_log.see(tk.END)
            output_log.config(state='disabled')

            # 检查文件是否存在
            if not os.path.exists(file_path):
                output_log.config(state='normal')
                output_log.insert(tk.END, f"文件不存在，跳过: {file_path}\n")
                output_log.see(tk.END)
                output_log.config(state='disabled')
                continue

            try:
                # 打开文档
                doc = word.Documents.Open(FileName=file_path, ReadOnly=True, AddToRecentFiles=False)

                # 提取名称
                name = extract_student_name(doc, pattern)
                if name:
                    safe_name = make_safe_filename(name)
                else:
                    safe_name = os.path.splitext(filename)[0]

                new_extension = os.path.splitext(filename)[1]
                new_filename = f"{safe_name}{new_extension}"
                new_file_path = os.path.join(input_dir, new_filename)
                
                doc.Close(False)

# 检查文件名冲突，添加序号
                count = 1
                while os.path.exists(new_file_path):
                    new_filename = f"{safe_name}_{count}{new_extension}"
                    new_file_path = os.path.join(input_dir, new_filename)
                    count += 1

                # 重命名文件
                os.rename(file_path, new_file_path)
                rename_history.append({"old": file_path, "new": new_file_path})  # 记录历史
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
        
        # 将历史记录保存到文件
        with open(history_file, "r+") as f:
            history = json.load(f)
            history.extend(rename_history)
            f.seek(0)
            json.dump(history, f, indent=4)
        
        messagebox.showinfo("完成", "所有文件处理完成。")
    finally:
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
        root.title("Word 文档批量重命名工具 By：Bilibili XiaoYun")
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
            # 显示软件协议提示框

        # 开始重命名按钮
        start_button = ttk.Button(frame, text="开始重命名", command=start_renaming, style='success.TButton', width=20)
        start_button.grid(row=7, column=1, pady=20)
        
        # 撤销按钮
        undo_button = ttk.Button(frame, text="撤销重命名", command=lambda: undo_rename(output_log), style='warning.TButton', width=20)
        undo_button.grid(row=7, column=2, pady=20)

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
