# -*- coding: utf-8 -*-
"""
题库录入服务 - 可视化启动器
提供 GUI 界面，可启动后端服务、下载插件压缩包
"""
import sys
import os
import shutil
import threading
import webbrowser
from pathlib import Path

# 打包成 exe 时，确保工作目录为 exe 所在目录
if getattr(sys, "frozen", False):
    exe_dir = Path(sys.executable).parent
    if exe_dir != Path.cwd():
        os.chdir(exe_dir)
    # 若未存在 config.yaml，从打包内嵌资源复制 config.example.yaml
    config_path = exe_dir / "config.yaml"
    if not config_path.exists():
        src = Path(sys._MEIPASS) / "config.example.yaml"
        if src.exists():
            shutil.copy(src, config_path)

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess


class ServerGUI:
    VERSION = "7.2.2"
    
    def __init__(self, root):
        self.root = root
        self.root.title(f"题库录入服务 v{self.VERSION}")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        self.server_process = None
        self.server_running = False
        
        self.setup_ui()
        self.center_window()
        
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text=f"题库录入 AI 服务 v{self.VERSION}",
            font=("Microsoft YaHei", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 状态显示
        self.status_frame = ttk.LabelFrame(main_frame, text="服务状态", padding="10")
        self.status_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.status_label = ttk.Label(
            self.status_frame,
            text="● 服务未启动",
            font=("Microsoft YaHei", 12),
            foreground="gray"
        )
        self.status_label.pack()
        
        self.url_label = ttk.Label(
            self.status_frame,
            text="",
            font=("Microsoft YaHei", 10),
            foreground="blue",
            cursor="hand2"
        )
        self.url_label.pack()
        self.url_label.bind("<Button-1>", lambda e: self.open_url())
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 启动/停止服务按钮
        self.start_btn = ttk.Button(
            btn_frame,
            text="启动服务",
            command=self.toggle_server,
            width=20
        )
        self.start_btn.pack(pady=5)
        
        # 分隔线
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # 插件下载区域
        plugin_frame = ttk.LabelFrame(main_frame, text="Chrome 插件", padding="10")
        plugin_frame.pack(fill=tk.X, pady=(0, 15))
        
        plugin_info = ttk.Label(
            plugin_frame,
            text="下载插件压缩包后，在 Chrome 扩展程序页面\n开启「开发者模式」，将 zip 解压后加载",
            font=("Microsoft YaHei", 9),
            justify=tk.CENTER
        )
        plugin_info.pack(pady=(0, 10))
        
        self.download_btn = ttk.Button(
            plugin_frame,
            text="下载插件 (extension.zip)",
            command=self.download_extension,
            width=25
        )
        self.download_btn.pack()
        
        # 配置文件区域
        config_frame = ttk.LabelFrame(main_frame, text="配置文件", padding="10")
        config_frame.pack(fill=tk.X)
        
        config_btn_frame = ttk.Frame(config_frame)
        config_btn_frame.pack()
        
        self.open_config_btn = ttk.Button(
            config_btn_frame,
            text="打开配置文件",
            command=self.open_config,
            width=15
        )
        self.open_config_btn.pack(side=tk.LEFT, padx=5)
        
        self.open_folder_btn = ttk.Button(
            config_btn_frame,
            text="打开所在目录",
            command=self.open_folder,
            width=15
        )
        self.open_folder_btn.pack(side=tk.LEFT, padx=5)
        
        # 关闭时的处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def toggle_server(self):
        if self.server_running:
            self.stop_server()
        else:
            self.start_server()
            
    def start_server(self):
        def run_server():
            try:
                import uvicorn
                self.server_running = True
                self.root.after(0, self.update_ui_running)
                uvicorn.run(
                    "app:app",
                    host="0.0.0.0",
                    port=8766,
                    log_level="info",
                )
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"服务启动失败:\n{e}"))
            finally:
                self.server_running = False
                self.root.after(0, self.update_ui_stopped)
        
        self.server_thread = threading.Thread(target=run_server, daemon=True)
        self.server_thread.start()
        
    def stop_server(self):
        messagebox.showinfo("提示", "请直接关闭此窗口来停止服务")
        
    def update_ui_running(self):
        self.status_label.config(text="● 服务运行中", foreground="green")
        self.url_label.config(text="http://127.0.0.1:8766 (点击打开)")
        self.start_btn.config(text="服务运行中...", state=tk.DISABLED)
        
    def update_ui_stopped(self):
        self.status_label.config(text="● 服务已停止", foreground="gray")
        self.url_label.config(text="")
        self.start_btn.config(text="启动服务", state=tk.NORMAL)
        
    def open_url(self):
        if self.server_running:
            webbrowser.open("http://127.0.0.1:8766")
            
    def download_extension(self):
        # 查找 extension.zip 的位置
        if getattr(sys, "frozen", False):
            # 打包后，从内嵌资源中复制
            src_path = Path(sys._MEIPASS) / "extension.zip"
        else:
            # 开发模式，从当前目录
            src_path = Path(__file__).parent / "extension.zip"
            
        if not src_path.exists():
            messagebox.showerror("错误", "未找到插件文件 extension.zip")
            return
            
        # 选择保存位置
        save_path = filedialog.asksaveasfilename(
            title="保存插件压缩包",
            defaultextension=".zip",
            filetypes=[("ZIP 文件", "*.zip")],
            initialfile=f"题库录入插件_v{self.VERSION}.zip"
        )
        
        if save_path:
            try:
                shutil.copy(src_path, save_path)
                messagebox.showinfo("成功", f"插件已保存到:\n{save_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败:\n{e}")
                
    def open_config(self):
        config_path = Path.cwd() / "config.yaml"
        if config_path.exists():
            if sys.platform == "win32":
                os.startfile(config_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", config_path])
            else:
                subprocess.run(["xdg-open", config_path])
        else:
            messagebox.showwarning("提示", "config.yaml 不存在，请先启动一次服务生成配置文件")
            
    def open_folder(self):
        folder = Path.cwd()
        if sys.platform == "win32":
            os.startfile(folder)
        elif sys.platform == "darwin":
            subprocess.run(["open", folder])
        else:
            subprocess.run(["xdg-open", folder])
            
    def on_closing(self):
        if self.server_running:
            if messagebox.askokcancel("退出", "服务正在运行，确定要退出吗？"):
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    app = ServerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
