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
import time
import queue
import io
from pathlib import Path
from datetime import datetime

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
from tkinter import ttk, messagebox, filedialog, scrolledtext


class LogRedirector:
    """重定向 stdout/stderr 到队列"""
    def __init__(self, log_queue, original_stream):
        self.log_queue = log_queue
        self.original_stream = original_stream
        
    def write(self, text):
        if text.strip():
            self.log_queue.put(text)
        if self.original_stream:
            self.original_stream.write(text)
            
    def flush(self):
        if self.original_stream:
            self.original_stream.flush()


class ServerGUI:
    VERSION = "7.3.5"
    
    def __init__(self, root):
        self.root = root
        self.root.title(f"题库录入服务 v{self.VERSION}")
        self.root.geometry("600x550")
        self.root.resizable(True, True)
        self.root.minsize(500, 450)
        
        self.server_thread = None
        self.server_running = False
        self.start_time = None
        self.log_queue = queue.Queue()
        
        self.setup_ui()
        self.center_window()
        self.update_timer()
        self.process_log_queue()
        
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text=f"题库录入 AI 服务 v{self.VERSION}",
            font=("Microsoft YaHei", 16, "bold")
        )
        title_label.pack(pady=(0, 15))
        
        # 状态显示区域
        status_frame = ttk.LabelFrame(main_frame, text="服务状态", padding="10")
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 状态行
        status_row = ttk.Frame(status_frame)
        status_row.pack(fill=tk.X)
        
        self.status_indicator = tk.Canvas(status_row, width=16, height=16, highlightthickness=0)
        self.status_indicator.pack(side=tk.LEFT, padx=(0, 8))
        self.draw_indicator("gray")
        
        self.status_label = ttk.Label(
            status_row,
            text="服务未启动",
            font=("Microsoft YaHei", 11)
        )
        self.status_label.pack(side=tk.LEFT)
        
        # 运行时间
        self.time_label = ttk.Label(
            status_row,
            text="",
            font=("Microsoft YaHei", 10),
            foreground="gray"
        )
        self.time_label.pack(side=tk.RIGHT)
        
        # URL 行
        url_row = ttk.Frame(status_frame)
        url_row.pack(fill=tk.X, pady=(8, 0))
        
        ttk.Label(url_row, text="服务地址：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        
        self.url_label = ttk.Label(
            url_row,
            text="http://127.0.0.1:8766",
            font=("Microsoft YaHei", 10),
            foreground="blue",
            cursor="hand2"
        )
        self.url_label.pack(side=tk.LEFT)
        self.url_label.bind("<Button-1>", lambda e: self.open_url())
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.start_btn = ttk.Button(
            btn_frame,
            text="▶ 启动服务",
            command=self.toggle_server,
            width=15
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_log_btn = ttk.Button(
            btn_frame,
            text="清空日志",
            command=self.clear_log,
            width=10
        )
        self.clear_log_btn.pack(side=tk.LEFT)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=10,
            font=("Consolas", 9),
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 底部按钮区域
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X)
        
        # 左侧：插件下载
        left_btns = ttk.Frame(bottom_frame)
        left_btns.pack(side=tk.LEFT)
        
        self.download_btn = ttk.Button(
            left_btns,
            text="📦 下载插件",
            command=self.download_extension,
            width=12
        )
        self.download_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 右侧：配置相关
        right_btns = ttk.Frame(bottom_frame)
        right_btns.pack(side=tk.RIGHT)
        
        self.open_config_btn = ttk.Button(
            right_btns,
            text="配置文件",
            command=self.open_config,
            width=10
        )
        self.open_config_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.open_folder_btn = ttk.Button(
            right_btns,
            text="打开目录",
            command=self.open_folder,
            width=10
        )
        self.open_folder_btn.pack(side=tk.LEFT)
        
        # 关闭时的处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def draw_indicator(self, color):
        """绘制状态指示灯"""
        self.status_indicator.delete("all")
        self.status_indicator.create_oval(2, 2, 14, 14, fill=color, outline="")
        
    def log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{timestamp}] {message}")
        
    def process_log_queue(self):
        """处理日志队列"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.root.after(100, self.process_log_queue)
        
    def clear_log(self):
        """清空日志"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def update_timer(self):
        """更新运行时间"""
        if self.server_running and self.start_time:
            elapsed = time.time() - self.start_time
            hours = int(elapsed // 3600)
            minutes = int((elapsed % 3600) // 60)
            seconds = int(elapsed % 60)
            if hours > 0:
                time_str = f"运行时间: {hours}:{minutes:02d}:{seconds:02d}"
            else:
                time_str = f"运行时间: {minutes}:{seconds:02d}"
            self.time_label.config(text=time_str)
        self.root.after(1000, self.update_timer)
        
    def toggle_server(self):
        if self.server_running:
            self.stop_server()
        else:
            self.start_server()
            
    def start_server(self):
        """启动服务"""
        self.log("正在启动服务...")
        self.start_btn.config(state=tk.DISABLED, text="启动中...")
        
        def run_server():
            try:
                import uvicorn
                import logging
                from app import app
                
                self.server_running = True
                self.start_time = time.time()
                self.root.after(0, self.update_ui_running)
                self.log("服务启动成功！")
                self.log(f"监听地址: http://0.0.0.0:8766")
                
                # 配置 uvicorn，禁用默认日志配置避免 GUI 模式下的问题
                config = uvicorn.Config(
                    app,
                    host="0.0.0.0",
                    port=8766,
                    log_level="info",
                    log_config=None,  # 禁用默认日志配置
                    access_log=False,  # 禁用访问日志避免输出问题
                )
                
                # 设置基本日志
                logging.basicConfig(level=logging.INFO)
                
                server = uvicorn.Server(config)
                server.run()
                
            except Exception as e:
                import traceback
                err_msg = f"服务启动失败: {e}"
                self.log(err_msg)
                self.log(traceback.format_exc())
                self.root.after(0, lambda: messagebox.showerror("错误", err_msg))
            finally:
                self.server_running = False
                self.start_time = None
                self.root.after(0, self.update_ui_stopped)
                self.log("服务已停止")
        
        self.server_thread = threading.Thread(target=run_server, daemon=True)
        self.server_thread.start()
        
        # 延迟检查服务是否启动成功
        self.root.after(2000, self.check_server_status)
        
    def check_server_status(self):
        """检查服务是否启动成功"""
        if not self.server_running:
            return
        try:
            import urllib.request
            urllib.request.urlopen("http://127.0.0.1:8766/", timeout=2)
            self.log("✓ 服务健康检查通过")
        except Exception as e:
            self.log(f"⚠ 健康检查: {e}")
        
    def stop_server(self):
        """停止服务提示"""
        messagebox.showinfo("提示", "请直接关闭此窗口来停止服务\n\n服务将随窗口一起关闭")
        
    def update_ui_running(self):
        """更新 UI 为运行状态"""
        self.draw_indicator("#22c55e")  # 绿色
        self.status_label.config(text="服务运行中")
        self.start_btn.config(text="■ 停止服务", state=tk.NORMAL)
        
    def update_ui_stopped(self):
        """更新 UI 为停止状态"""
        self.draw_indicator("gray")
        self.status_label.config(text="服务未启动")
        self.time_label.config(text="")
        self.start_btn.config(text="▶ 启动服务", state=tk.NORMAL)
        
    def open_url(self):
        """打开服务地址"""
        webbrowser.open("http://127.0.0.1:8766")
            
    def download_extension(self):
        """下载插件"""
        if getattr(sys, "frozen", False):
            src_path = Path(sys._MEIPASS) / "extension.zip"
        else:
            src_path = Path(__file__).parent / "extension.zip"
            
        if not src_path.exists():
            messagebox.showerror("错误", "未找到插件文件 extension.zip")
            return
            
        save_path = filedialog.asksaveasfilename(
            title="保存插件压缩包",
            defaultextension=".zip",
            filetypes=[("ZIP 文件", "*.zip")],
            initialfile=f"题库录入插件_v{self.VERSION}.zip"
        )
        
        if save_path:
            try:
                shutil.copy(src_path, save_path)
                self.log(f"插件已保存: {save_path}")
                messagebox.showinfo("成功", f"插件已保存到:\n{save_path}\n\n请解压后在 Chrome 扩展程序页面加载")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败:\n{e}")
                
    def open_config(self):
        """打开配置文件"""
        config_path = Path.cwd() / "config.yaml"
        if config_path.exists():
            if sys.platform == "win32":
                os.startfile(config_path)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.run(["open", config_path])
            else:
                import subprocess
                subprocess.run(["xdg-open", config_path])
            self.log(f"已打开配置文件: {config_path}")
        else:
            messagebox.showwarning("提示", "config.yaml 不存在\n请先启动一次服务生成配置文件")
            
    def open_folder(self):
        """打开所在目录"""
        folder = Path.cwd()
        if sys.platform == "win32":
            os.startfile(folder)
        elif sys.platform == "darwin":
            import subprocess
            subprocess.run(["open", folder])
        else:
            import subprocess
            subprocess.run(["xdg-open", folder])
        self.log(f"已打开目录: {folder}")
            
    def on_closing(self):
        """关闭窗口"""
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
