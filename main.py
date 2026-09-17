# -*- coding: utf-8 -*-
"""
一键处理工具（单文件 exe，模板已内置，无需外部目录）
"""
import os
import sys
import glob
import shutil
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import yaml

# ==================== 配置区 ====================
# 三个模板的名字（对应打包时 --add-data 里 resource 目录下的子目录）
TEMPLATE_NAMES = ["SHUI", "冷冻水", "冷源", "IEAC"]  # ← 改成实际名字

SCRIPT_NAME = "excel_cleaner.py"
CONFIG_NAME = "config.yaml"
PROJECT_NAME_KEY = "project_name"


def get_resource_path(relative_path):
    """获取打包后资源的真实路径（兼容开发环境和 exe 环境）"""
    # 打包后资源在 sys._MEIPASS 下
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def update_project_name(config_path, project_name):
    """修改 config.yaml 里的 project_name"""
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    cfg[PROJECT_NAME_KEY] = project_name
    with open(config_path, "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, allow_unicode=True, default_flow_style=False)


def run_cleaner_script(script_path, cwd):
    """在当前进程执行 excel_cleaner.py"""
    old_cwd = os.getcwd()
    os.chdir(cwd)
    try:
        with open(script_path, "r", encoding="utf-8") as f:
            code = f.read()
        exec(compile(code, script_path, "exec"), {"__name__": "__main__"})
    finally:
        os.chdir(old_cwd)


def main():
    root = tk.Tk()
    root.withdraw()

    # 1. 选择源数据文件
    messagebox.showinfo("步骤 1/3", "请选择源数据文件")
    src_file = filedialog.askopenfilename(
        title="选择源数据文件",
        filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
    )
    if not src_file:
        messagebox.showwarning("取消", "未选择文件，程序退出")
        return

    # 2. 自动提取 project_name（去掉扩展名）
    project_name = os.path.splitext(os.path.basename(src_file))[0].strip()

    # 3. 选择模板
    choice = simpledialog.askstring(
        "选择模板",
        f"请输入模板名（{'/'.join(TEMPLATE_NAMES)}）："
    )
    if choice not in TEMPLATE_NAMES:
        messagebox.showerror("错误", f"模板名无效，应为: {'/'.join(TEMPLATE_NAMES)}")
        return

    # 4. 选择输出目录
    messagebox.showinfo("提示", "请选择结果输出目录")
    out_root = filedialog.askdirectory(title="选择输出目录")
    if not out_root:
        messagebox.showwarning("取消", "未选择输出目录，程序退出")
        return

    # 5. 解压模板到临时目录
    messagebox.showinfo("步骤 3/3", "开始处理，请稍候...")
    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            # 把内置模板复制到临时目录
            src_template = get_resource_path(os.path.join("resource", choice))
            work_dir = os.path.join(tmp_dir, choice)
            shutil.copytree(src_template, work_dir)

            input_dir = os.path.join(work_dir, "input")
            config_path = os.path.join(work_dir, CONFIG_NAME)
            script_path = os.path.join(work_dir, SCRIPT_NAME)

            # 复制源文件到 input
            os.makedirs(input_dir, exist_ok=True)
            shutil.copy2(src_file, os.path.join(input_dir, os.path.basename(src_file)))

            # 修改 project_name
            update_project_name(config_path, project_name)

            # 执行模板脚本
            run_cleaner_script(script_path, work_dir)

            # 把 output 目录复制到用户选择的输出目录
            output_dir = os.path.join(work_dir, "output")
            if os.path.isdir(output_dir):
                final_out = os.path.join(out_root, project_name)
                shutil.copytree(output_dir, final_out, dirs_exist_ok=True)
                messagebox.showinfo("完成", f"处理完成！\n\n结果已保存到:\n{final_out}")
            else:
                messagebox.showwarning("提示", "处理完成，但未生成 output 目录，请检查模板脚本逻辑。")
    except Exception as e:
        messagebox.showerror("出错", f"处理失败:\n{e}")


if __name__ == "__main__":
    main()
