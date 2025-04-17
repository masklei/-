import subprocess
import sys
import os
import tkinter as tk
from tkinter import messagebox

def install_package(package):
    """安装Python包"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except subprocess.CalledProcessError:
        return False

def main():
    # 创建简单的GUI窗口
    root = tk.Tk()
    root.title("依赖安装")
    root.geometry("400x300")
    
    # 设置窗口居中
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - 400) // 2
    y = (screen_height - 300) // 2
    root.geometry(f"400x300+{x}+{y}")
    
    # 创建标签
    label = tk.Label(root, text="正在安装必要的依赖包...", font=("Arial", 12))
    label.pack(pady=20)
    
    # 创建文本框显示安装进度
    text = tk.Text(root, height=10, width=45)
    text.pack(pady=10, padx=10)
    
    # 要安装的包列表
    packages = [
        "pandas",
        "openpyxl",
        "xlrd",
        "pyxlsb",
        "pywin32"
    ]
    
    # 安装进度
    progress = 0
    total = len(packages)
    
    def update_progress(package, success):
        nonlocal progress
        progress += 1
        status = "成功" if success else "失败"
        text.insert(tk.END, f"[{progress}/{total}] 安装 {package}: {status}\n")
        text.see(tk.END)
        
        # 更新标签
        percent = int((progress / total) * 100)
        label.config(text=f"正在安装必要的依赖包... {percent}%")
        
        # 如果所有包都已处理，显示完成消息
        if progress >= total:
            if all_success:
                label.config(text="所有依赖安装完成！")
                messagebox.showinfo("安装完成", "所有依赖包已成功安装，现在可以运行自动化流程了。")
            else:
                label.config(text="部分依赖安装失败，请查看详情")
                messagebox.showwarning("安装警告", "部分依赖包安装失败，可能会影响程序运行。")
            
            # 添加关闭按钮
            close_button = tk.Button(root, text="关闭", command=root.destroy, width=10)
            close_button.pack(pady=10)
    
    # 安装包
    all_success = True
    
    def install_all_packages():
        nonlocal all_success
        for package in packages:
            text.insert(tk.END, f"正在安装 {package}...\n")
            text.see(tk.END)
            success = install_package(package)
            if not success:
                all_success = False
            root.after(100, lambda p=package, s=success: update_progress(p, s))
    
    # 在单独的线程中安装包
    root.after(500, install_all_packages)
    
    root.mainloop()

if __name__ == "__main__":
    main()