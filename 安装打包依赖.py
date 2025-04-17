import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from datetime import datetime

# 需要安装的依赖包
REQUIRED_PACKAGES = [
    "pyinstaller",  # 打包工具
    "pandas",      # 数据处理
    "openpyxl",    # Excel处理
    "xlrd",        # Excel读取
    "pywin32"      # Windows API接口
]

class LogHandler:
    """自定义日志处理器，将日志输出到GUI"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        
    def write(self, message):
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, message)
            self.text_widget.configure(state='disabled')
            self.text_widget.yview(tk.END)
        # 在主线程中更新UI
        self.text_widget.after(0, append)
    
    def flush(self):
        pass

class DependencyInstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("考勤自动化处理系统依赖安装工具")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        # 设置窗口图标
        try:
            self.root.iconbitmap("favicon.ico")
        except:
            pass
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题
        title_label = ttk.Label(self.main_frame, text="考勤自动化处理系统依赖安装工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 创建说明文本
        desc_text = "此工具将安装考勤自动化处理系统所需的所有Python依赖包。\n请确保您的计算机已连接到互联网。"
        desc_label = ttk.Label(self.main_frame, text=desc_text, font=("Arial", 10))
        desc_label.pack(pady=5)
        
        # 创建依赖列表框架
        deps_frame = ttk.LabelFrame(self.main_frame, text="需要安装的依赖包", padding="10")
        deps_frame.pack(fill=tk.X, pady=5)
        
        # 显示依赖包列表
        for package in REQUIRED_PACKAGES:
            package_frame = ttk.Frame(deps_frame)
            package_frame.pack(fill=tk.X, pady=2)
            
            package_label = ttk.Label(package_frame, text=package, width=15)
            package_label.pack(side=tk.LEFT, padx=5)
            
            status_var = tk.StringVar(value="等待安装")
            setattr(self, f"{package}_status", status_var)
            
            status_label = ttk.Label(package_frame, textvariable=status_var, width=15)
            status_label.pack(side=tk.LEFT, padx=5)
        
        # 创建进度条
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(self.main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        # 创建日志显示区域
        log_frame = ttk.LabelFrame(self.main_frame, text="安装日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 创建日志处理器
        self.log_handler = LogHandler(self.log_text)
        
        # 创建按钮框架
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 创建安装按钮
        self.install_button = ttk.Button(
            button_frame, 
            text="开始安装", 
            command=self.start_installation,
            style="Big.TButton"
        )
        self.install_button.pack(side=tk.LEFT, padx=5)
        
        # 创建退出按钮
        self.exit_button = ttk.Button(button_frame, text="退出", command=self.root.destroy)
        self.exit_button.pack(side=tk.RIGHT, padx=5)
        
        # 创建进度百分比标签
        self.progress_percent = tk.StringVar(value="0%")
        progress_label = ttk.Label(self.main_frame, textvariable=self.progress_percent, font=("Arial", 10, "bold"))
        progress_label.pack(pady=(0, 5))
        
        # 初始化
        self.is_running = False
    
    def log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_handler.write(f"{timestamp} - {message}\n")
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_var.set(value)
        self.progress_percent.set(f"{int(value)}%")
    
    def update_package_status(self, package, status):
        """更新包安装状态"""
        status_var = getattr(self, f"{package}_status", None)
        if status_var:
            status_var.set(status)
    
    def install_package(self, package):
        """安装单个包"""
        self.log(f"正在安装 {package}...")
        self.update_package_status(package, "安装中")
        
        try:
            # 检查包是否已安装
            result = subprocess.run(
                [sys.executable, "-m", "pip", "show", package], 
                capture_output=True, 
                text=True
            )
            
            if result.returncode == 0:
                self.log(f"{package} 已安装")
                self.update_package_status(package, "已安装")
                return True
            
            # 安装包
            process = subprocess.Popen(
                [sys.executable, "-m", "pip", "install", package],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8'
            )
            
            # 实时读取输出
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    self.log(output.strip())
            
            # 获取剩余输出和错误
            stdout, stderr = process.communicate()
            
            if stdout:
                self.log(stdout.strip())
            
            if process.returncode != 0:
                self.log(f"安装 {package} 失败: {stderr.strip()}")
                self.update_package_status(package, "安装失败")
                return False
            
            self.log(f"{package} 安装成功")
            self.update_package_status(package, "安装成功")
            return True
            
        except Exception as e:
            self.log(f"安装 {package} 时出错: {str(e)}")
            self.update_package_status(package, "安装失败")
            return False
    
    def install_all_packages(self):
        """安装所有依赖包"""
        try:
            self.log("开始安装依赖包...")
            
            # 更新pip
            self.log("正在更新pip...")
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", "--upgrade", "pip"],
                    check=True,
                    capture_output=True,
                    text=True
                )
                self.log("pip更新成功")
            except Exception as e:
                self.log(f"更新pip时出错: {str(e)}")
            
            # 安装所有包
            total_packages = len(REQUIRED_PACKAGES)
            for i, package in enumerate(REQUIRED_PACKAGES):
                success = self.install_package(package)
                progress = ((i + 1) / total_packages) * 100
                self.update_progress(progress)
            
            self.log("所有依赖包安装完成")
            messagebox.showinfo("安装完成", "所有依赖包安装完成！\n现在您可以运行打包工具了。")
            
        except Exception as e:
            self.log(f"安装过程中发生错误: {str(e)}")
            messagebox.showerror("错误", f"安装过程中发生错误: {str(e)}")
        finally:
            self.is_running = False
            self.install_button.config(state="normal")
    
    def start_installation(self):
        """开始安装流程"""
        if self.is_running:
            return
        
        # 重置状态
        self.progress_var.set(0)
        self.progress_percent.set("0%")
        
        # 重置包状态
        for package in REQUIRED_PACKAGES:
            self.update_package_status(package, "等待安装")
        
        # 清空日志
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        
        # 禁用安装按钮
        self.install_button.config(state="disabled")
        
        # 设置运行标志
        self.is_running = True
        
        # 启动安装线程
        import threading
        install_thread = threading.Thread(target=self.install_all_packages)
        install_thread.daemon = True
        install_thread.start()

# 主函数
def main():
    root = tk.Tk()
    app = DependencyInstallerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()