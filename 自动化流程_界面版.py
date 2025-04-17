import os
import sys
import glob
import time
import subprocess
import logging
import threading
import tempfile  # 添加这行导入语句
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from datetime import datetime

# 工作目录
WORK_DIR = os.path.dirname(os.path.abspath(__file__))

# 配置日志
log_file = os.path.join(WORK_DIR, "processing.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

class LogHandler(logging.Handler):
    """自定义日志处理器，将日志输出到GUI"""
    def __init__(self, text_widget):
        logging.Handler.__init__(self)
        self.text_widget = text_widget
        
    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.yview(tk.END)
        # 在主线程中更新UI
        self.text_widget.after(0, append)

class AutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("考勤自动化处理系统@2025")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 设置窗口图标
        try:
            self.root.iconbitmap("favicon.ico")
        except:
            pass
            
        # 创建自定义样式
        style = ttk.Style()
        style.configure("Big.TButton", font=("Arial", 12, "bold"))
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题
        title_label = ttk.Label(self.main_frame, text="考勤自动化处理流程", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 创建步骤框架
        self.steps_frame = ttk.LabelFrame(self.main_frame, text="处理步骤", padding="10")
        self.steps_frame.pack(fill=tk.X, pady=5)
        
        # 创建步骤标签和状态
        self.steps = [
            {"name": "步骤1: 文件修复", "status": "等待中", "var": tk.StringVar(value="等待中")},
            {"name": "步骤2: 班别分类", "status": "等待中", "var": tk.StringVar(value="等待中")},
            {"name": "步骤3: 白班稽核", "status": "等待中", "var": tk.StringVar(value="等待中")},
            {"name": "步骤4: 夜班稽核", "status": "等待中", "var": tk.StringVar(value="等待中")},
            {"name": "步骤5: 合并文件", "status": "等待中", "var": tk.StringVar(value="等待中")},
            {"name": "步骤6: 异常数据稽核", "status": "等待中", "var": tk.StringVar(value="等待中")},
            {"name": "步骤7: 内容优化", "status": "等待中", "var": tk.StringVar(value="等待中")}
        ]
        
        # 创建步骤UI
        for i, step in enumerate(self.steps):
            step_frame = ttk.Frame(self.steps_frame)
            step_frame.pack(fill=tk.X, pady=2)
            
            step_label = ttk.Label(step_frame, text=step["name"], width=20)
            step_label.pack(side=tk.LEFT, padx=5)
            
            status_label = ttk.Label(step_frame, textvariable=step["var"], width=10)
            status_label.pack(side=tk.LEFT, padx=5)
        
        # 创建进度条
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(self.main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        # 创建日志显示区域
        log_frame = ttk.LabelFrame(self.main_frame, text="处理日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 减小日志区域高度，确保按钮可见
        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', height=8)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 添加自定义日志处理器
        self.log_handler = LogHandler(self.log_text)
        self.log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(self.log_handler)
        
        # 创建一个更美观的开始按钮
        self.main_start_button = tk.Button(
            self.main_frame, 
            text="开始处理", 
            command=self.start_process,
            font=("微软雅黑", 14, "bold"),
            bg="#3498db",  # 蓝色背景
            fg="white",    # 白色文字
            height=2,      # 增加按钮高度
            relief=tk.RAISED,
            borderwidth=3,
            cursor="hand2"  # 鼠标悬停时显示手型光标
        )
        self.main_start_button.pack(fill=tk.X, pady=10)
        
        # 创建进度百分比标签
        self.progress_percent = tk.StringVar(value="0%")
        progress_label = ttk.Label(self.main_frame, textvariable=self.progress_percent, font=("Arial", 10, "bold"))
        progress_label.pack(pady=(0, 5))
        
        # 创建按钮框架
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        # 创建重新开始按钮
        self.start_button = ttk.Button(button_frame, text="重新开始", command=self.start_process)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        # 创建查看日志按钮
        self.log_button = ttk.Button(button_frame, text="查看完整日志", command=self.view_log)
        self.log_button.pack(side=tk.LEFT, padx=5)
        
        # 创建退出按钮
        self.exit_button = ttk.Button(button_frame, text="退出", command=self.root.destroy)
        self.exit_button.pack(side=tk.RIGHT, padx=5)
        
        # 处理窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 初始化处理线程
        self.process_thread = None
        self.is_running = False
    
    def update_step_status(self, step_index, status):
        """更新步骤状态"""
        self.steps[step_index]["status"] = status
        self.steps[step_index]["var"].set(status)
        
        # 更新进度条
        completed_steps = sum(1 for step in self.steps if step["status"] in ["完成", "跳过"])
        progress = (completed_steps / len(self.steps)) * 100
        self.progress_var.set(progress)
        
        # 更新百分比显示
        self.progress_percent.set(f"{int(progress)}%")
    
    def run_script(self, script_name):
        """运行Python脚本"""
        script_path = os.path.join(WORK_DIR, script_name)
        if not os.path.exists(script_path):
            logging.error(f"脚本不存在: {script_path}")
            return False
        
        # 获取打包后的可执行文件所在目录
        if getattr(sys, 'frozen', False):
            # 打包后的运行模式
            application_path = os.path.dirname(sys.executable)
        else:
            # 普通Python运行模式
            application_path = WORK_DIR
        
        cmd = [sys.executable, script_path]
        
        logging.info(f"开始执行: {script_name}")
        start_time = time.time()
        try:
            # 修改环境变量，确保使用UTF-8编码
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            
            process = subprocess.Popen(
                cmd,
                cwd=application_path,  # 设置工作目录
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='gbk',
                env=env
            )
            stdout, stderr = process.communicate()
            
            end_time = time.time()
            run_time = end_time - start_time
            
            if process.returncode != 0:
                logging.error(f"执行失败: {script_name}")
                logging.error(f"错误信息: {stderr}")
                return False
            
            logging.info(f"执行成功: {script_name}, 耗时: {run_time:.2f}秒")
            return True
        except Exception as e:
            end_time = time.time()
            run_time = end_time - start_time
            logging.error(f"执行异常: {script_name}, 错误: {str(e)}, 耗时: {run_time:.2f}秒")
            return False
    
    def find_latest_file(self, pattern, exclude_pattern=None, in_data_dir=False):
        """查找最新的匹配文件
        
        Args:
            pattern: 文件匹配模式
            exclude_pattern: 排除的文件模式
            in_data_dir: 是否在考勤数据目录中查找
        """
        # 确定查找目录
        search_dir = WORK_DIR
        if in_data_dir:
            search_dir = os.path.join(WORK_DIR, "考勤数据")
            
        # 构建完整的查找路径
        search_path = os.path.join(search_dir, pattern)
        files = glob.glob(search_path)
        
        # 排除不需要的文件
        if exclude_pattern:
            files = [f for f in files if exclude_pattern not in f]
        
        if not files:
            logging.warning(f"未找到匹配 '{pattern}' 的文件 (在{'考勤数据目录' if in_data_dir else '当前目录'})")
            return None
        
        # 按修改时间排序，返回最新的文件
        latest_file = max(files, key=os.path.getmtime)
        logging.info(f"找到最新文件: {os.path.basename(latest_file)}")
        return latest_file
    
    def process_automation(self):
        """执行自动化处理流程"""
        # 确保工作目录正确
        if getattr(sys, 'frozen', False):
            os.chdir(os.path.dirname(sys.executable))
        else:
            os.chdir(WORK_DIR)
        
        start_time = time.time()
        logging.info("===== 开始自动化考勤处理流程 =====")
        
        try:
            # 步骤1: 运行EXCEL修复.py
            self.update_step_status(0, "执行中")
            logging.info("步骤1: 运行Excel修复工具")
            if not self.run_script("EXCEL修复.py"):
                logging.error("Excel修复失败")
                self.update_step_status(0, "失败")
            else:
                self.update_step_status(0, "完成")
            
            # 等待文件系统更新
            time.sleep(2)
            
            # 步骤2: 运行班别分类.py
            self.update_step_status(1, "执行中")
            logging.info("步骤2: 运行班别分类工具")
            if not self.run_script("班别分类.py"):
                logging.error("班别分类失败")
                self.update_step_status(1, "失败")
                return
            else:
                self.update_step_status(1, "完成")
            
            # 等待文件系统更新
            time.sleep(2)
            
            # 查找班别分类生成的文件 - 在考勤数据目录中查找
            班别分类结果 = self.find_latest_file("*班别*结果*.xlsx", in_data_dir=True)
            if not 班别分类结果:
                # 尝试其他可能的文件名模式
                班别分类结果 = self.find_latest_file("班别*.xlsx", in_data_dir=True)
                
            if not 班别分类结果:
                logging.error("未找到班别分类结果文件，流程终止")
                messagebox.showerror("错误", "未找到班别分类结果文件，请确认班别分类步骤是否正确完成")
                return
            
            logging.info(f"找到班别分类结果文件: {os.path.basename(班别分类结果)}")
            
            # 步骤3: 运行白班稽核1_1.py
            self.update_step_status(2, "执行中")
            logging.info("步骤3: 运行白班稽核工具")
            if not self.run_script("白班稽核1_1.py"):
                logging.error("白班稽核失败")
                self.update_step_status(2, "失败")
                return
            else:
                self.update_step_status(2, "完成")
            
            # 等待文件系统更新
            time.sleep(2)
            
            # 查找白班稽核生成的文件 - 使用正确的文件名模式
            白班异常文件 = self.find_latest_file("*白班稽查结果*.xlsx", in_data_dir=True)
            if not 白班异常文件:
                # 尝试其他可能的文件名模式
                白班异常文件 = self.find_latest_file("*白班*.xlsx", in_data_dir=True)
                
            if not 白班异常文件:
                logging.warning("未找到白班稽核结果文件，可能没有白班异常")
            else:
                logging.info(f"找到白班稽核结果文件: {os.path.basename(白班异常文件)}")
            
            # 步骤4: 运行夜班稽核.py
            self.update_step_status(3, "执行中")
            logging.info("步骤4: 运行夜班稽核工具")
            if not self.run_script("夜班稽核.py"):
                logging.error("夜班稽核失败")
                self.update_step_status(3, "失败")
                return
            else:
                self.update_step_status(3, "完成")
            
            # 等待文件系统更新
            time.sleep(2)
            
            # 查找夜班稽核生成的文件 - 使用正确的文件名模式
            夜班异常文件 = self.find_latest_file("*夜班稽查结果*.xlsx", in_data_dir=True)
            if not 夜班异常文件:
                # 尝试其他可能的文件名模式
                夜班异常文件 = self.find_latest_file("*夜班*.xlsx", in_data_dir=True)
                
            if not 夜班异常文件:
                logging.warning("未找到夜班稽核结果文件，可能没有夜班异常")
            else:
                logging.info(f"找到夜班稽核结果文件: {os.path.basename(夜班异常文件)}")
            
            # 步骤5: 运行合并Excel文件.py (如果有两个异常文件)
            if 白班异常文件 and 夜班异常文件:
                self.update_step_status(4, "执行中")
                logging.info("步骤5: 运行合并Excel文件工具")
                if not self.run_script("合并Excel文件.py"):
                    logging.error("合并Excel文件失败")
                    self.update_step_status(4, "失败")
                    return
                else:
                    self.update_step_status(4, "完成")
                
                # 等待文件系统更新
                time.sleep(2)
                
                # 查找合并结果 - 修正文件查找模式
                合并结果文件 = self.find_latest_file("*合并结果*.xlsx", in_data_dir=True)
                if not 合并结果文件:
                    # 尝试其他可能的文件名模式
                    合并结果文件 = self.find_latest_file("合并*.xlsx", in_data_dir=True)
                    
                if not 合并结果文件:
                    logging.error("未找到合并结果文件，流程终止")
                    return
                
                logging.info(f"找到合并结果文件: {os.path.basename(合并结果文件)}")
            else:
                logging.info("跳过合并步骤，因为没有足够的异常文件")
                self.update_step_status(4, "跳过")
                # 使用可用的异常文件作为最终结果
                合并结果文件 = 白班异常文件 or 夜班异常文件
            
            # 步骤6: 运行异常数据稽核.py (原考勤核查.py)
            if 合并结果文件:
                self.update_step_status(5, "执行中")
                logging.info("步骤6: 运行异常数据稽核工具")
                if not self.run_script("异常数据稽核.py"):
                    logging.error("异常数据稽核失败")
                    self.update_step_status(5, "失败")
                    return
                else:
                    self.update_step_status(5, "完成")
                
                # 等待文件系统更新
                time.sleep(2)
                
                # 查找核对版数据文件 - 确保在考勤数据目录中查找最新生成的文件
                核对版数据文件 = self.find_latest_file("*核对版数据*.xlsx", in_data_dir=True)
                if not 核对版数据文件:
                    # 尝试其他可能的文件名模式
                    核对版数据文件 = self.find_latest_file("*核对版*.xlsx", in_data_dir=True)
                
                if not 核对版数据文件:
                    logging.warning("未找到核对版数据文件，尝试继续执行内容优化步骤")
                    # 尝试查找可能的输入文件
                    可能的输入文件 = self.find_latest_file("*.xlsx", in_data_dir=True)
                    if 可能的输入文件:
                        logging.info(f"将使用找到的最新Excel文件作为内容优化的输入: {os.path.basename(可能的输入文件)}")
                        核对版数据文件 = 可能的输入文件
                    else:
                        logging.error("未找到任何可用的Excel文件，无法继续执行内容优化")
                        self.update_step_status(6, "跳过")
                        return
                else:
                    logging.info(f"找到核对版数据文件: {os.path.basename(核对版数据文件)}")
                
                # 步骤7: 运行内容优化.py
                self.update_step_status(6, "执行中")
                logging.info("步骤7: 运行内容优化工具")
                
                # 确保考勤数据目录存在
                data_dir = os.path.join(WORK_DIR, "考勤数据")
                if not os.path.exists(data_dir):
                    os.makedirs(data_dir)
                    logging.info(f"创建考勤数据目录: {data_dir}")
                
                # 运行内容优化脚本
                if not self.run_script("内容优化.py"):
                    logging.error("内容优化失败")
                    self.update_step_status(6, "失败")
                else:
                    self.update_step_status(6, "完成")
                    
                # 等待文件系统更新
                time.sleep(2)
                
                # 查找优化后的文件
                优化结果文件 = self.find_latest_file("*稽核数据核对版.xlsx", in_data_dir=True)
                if not 优化结果文件:
                    # 如果在考勤数据目录中找不到，尝试在当前目录查找
                    优化结果文件 = self.find_latest_file("*稽核数据核对版.xlsx")
                    
                    # 如果找到了，将其移动到考勤数据目录
                    if 优化结果文件:
                        import shutil
                        dest_file = os.path.join(data_dir, os.path.basename(优化结果文件))
                        try:
                            shutil.move(优化结果文件, dest_file)
                            优化结果文件 = dest_file
                            logging.info(f"已将优化结果文件移动到考勤数据目录: {os.path.basename(dest_file)}")
                        except Exception as e:
                            logging.error(f"移动优化结果文件失败: {str(e)}")
                
                if 优化结果文件:
                    logging.info(f"找到优化结果文件: {os.path.basename(优化结果文件)}")
                else:
                    logging.warning("未找到优化结果文件")
            else:
                logging.info("跳过异常数据稽核和内容优化步骤，因为没有异常文件")
                self.update_step_status(5, "跳过")
                self.update_step_status(6, "跳过")
            
            # 计算总运行时间
            end_time = time.time()
            run_time = end_time - start_time
            logging.info(f"===== 自动化考勤处理流程完成 =====")
            logging.info(f"总运行时间: {run_time:.2f}秒")
            
            # 复制最终结果文件到脚本目录并清理临时文件
            final_result_file = None
            
            # 查找最终的优化结果文件 - 修正查找模式
            final_result_file = self.find_latest_file("*稽核数据核对版.xlsx", in_data_dir=True)
            if not final_result_file:
                final_result_file = self.find_latest_file("*考勤稽核数据核对版.xlsx", in_data_dir=True)
            if not final_result_file:
                final_result_file = self.find_latest_file("*核对版*.xlsx", in_data_dir=True)
            
            if final_result_file:
                # 复制文件到脚本目录
                import shutil
                dest_file = os.path.join(WORK_DIR, os.path.basename(final_result_file))
                try:
                    shutil.copy2(final_result_file, dest_file)
                    logging.info(f"已将最终结果文件复制到脚本目录: {os.path.basename(dest_file)}")
                    
                    # 确保文件复制成功后再清理临时文件夹
                    if os.path.exists(dest_file):
                        # 清理临时文件夹
                        data_dir = os.path.join(WORK_DIR, "考勤数据")
                        if os.path.exists(data_dir) and os.path.isdir(data_dir):
                            try:
                                # 删除临时文件夹及其内容
                                shutil.rmtree(data_dir)
                                logging.info("已清理考勤数据临时文件夹")
                                
                                # 重新创建空文件夹以备下次使用
                                os.makedirs(data_dir)
                                logging.info("已重新创建考勤数据目录")
                            except Exception as e:
                                logging.warning(f"清理临时文件夹失败: {str(e)}")
                    else:
                        logging.warning("最终结果文件复制可能不成功，跳过清理临时文件夹")
                except Exception as e:
                    logging.error(f"复制最终结果文件失败: {str(e)}")
            else:
                logging.warning("未找到最终结果文件，无法复制到脚本目录")
            
            # 显示完成消息
            messagebox.showinfo("处理完成", f"自动化考勤处理流程已完成！\n总运行时间: {run_time:.2f}秒")
            
        except Exception as e:
            logging.error(f"处理过程中发生错误: {str(e)}")
            messagebox.showerror("错误", f"处理过程中发生错误: {str(e)}")
        finally:
            self.is_running = False
            # 启用两个开始按钮
            self.start_button.config(state="normal")
            self.main_start_button.config(state="normal")
    
    def start_process(self):
        """开始处理流程"""
        if self.is_running:
            return
        
        # 重置步骤状态
        for i in range(len(self.steps)):
            self.update_step_status(i, "等待中")
        
        self.progress_var.set(0)
        self.is_running = True
        # 禁用两个开始按钮
        self.start_button.config(state="disabled")
        self.main_start_button.config(state="disabled")
        
        # 在新线程中运行处理流程
        self.process_thread = threading.Thread(target=self.process_automation)
        self.process_thread.daemon = True
        self.process_thread.start()
    
    def view_log(self):
        """查看完整日志文件"""
        try:
            os.startfile(log_file)
        except:
            try:
                subprocess.Popen(['notepad', log_file])
            except Exception as e:
                messagebox.showerror("错误", f"无法打开日志文件: {str(e)}")
    
    def on_closing(self):
        """窗口关闭事件处理"""
        if self.is_running:
            if messagebox.askyesno("确认", "处理流程正在运行中，确定要退出吗？"):
                self.root.destroy()
        else:
            self.root.destroy()

# 在文件末尾的if __name__ == "__main__"部分添加以下代码
if __name__ == "__main__":
    # 确定工作目录
    if getattr(sys, 'frozen', False):
        WORK_DIR = os.path.dirname(sys.executable)
    else:
        WORK_DIR = os.path.dirname(os.path.abspath(__file__))
    
    # 修改锁文件处理逻辑
    lock_file = os.path.join(tempfile.gettempdir(), "考勤自动化处理系统.lock")
    
    try:
        # 检查并删除旧的锁文件
        if os.path.exists(lock_file):
            try:
                os.remove(lock_file)
            except:
                pass
        
        # 创建新的锁文件
        with open(lock_file, 'w') as f:
            f.write(str(time.time()))
            
        # 注册退出时删除锁文件
        def cleanup():
            try:
                if os.path.exists(lock_file):
                    os.remove(lock_file)
            except:
                pass
        
        import atexit
        atexit.register(cleanup)
    except Exception as e:
        print(f"锁文件处理异常: {e}")

    # 确保工作目录正确
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    # 正常启动程序
    root = tk.Tk()
    app = AutomationApp(root)
    root.mainloop()