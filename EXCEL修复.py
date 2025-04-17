import os
import pandas as pd
from openpyxl import load_workbook
import xlrd
import pyxlsb
import tkinter as tk
from tkinter import filedialog
import tempfile
import shutil
import time
import concurrent.futures
from functools import partial
import re

def select_excel_files():
    """选择多个要测试的Excel文件"""
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="选择要修复的Excel文件(可多选)",
        filetypes=[("Excel文件", "*.xlsx *.xls *.xlsb"), ("所有文件", "*.*")]
    )
    return file_paths

def process_single_file(file_path, i, total_files):
    """处理单个文件的线程函数"""
    print(f"\n正在处理文件 {i}/{total_files}: {file_path}")
    if repair_excel_file(file_path):
        print(f"✅ 文件修复成功: {file_path}")
        return True
    else:
        print(f"❌ 文件修复失败: {file_path}")
        return False

def main():
    print("=== Excel文件批量修复工具 ===")
    file_paths = select_excel_files()
    
    if not file_paths:
        print("未选择文件，程序退出")
        return
    
    # 使用线程池处理文件
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # 创建处理函数的部分应用
        process_func = partial(process_single_file, total_files=len(file_paths))
        # 提交所有任务
        futures = [executor.submit(process_func, path, i+1) 
                  for i, path in enumerate(file_paths)]
        
        # 统计结果
        success_count = sum(future.result() for future in concurrent.futures.as_completed(futures) 
                          if future.result() is not None)
        fail_count = len(file_paths) - success_count
    
    print("\n修复结果统计:")
    print(f"成功修复: {success_count} 个文件")
    print(f"修复失败: {fail_count} 个文件")

def test_pandas_read(file_path):
    """测试用pandas读取Excel文件"""
    print("\n尝试使用pandas读取...")
    try:
        # 尝试读取Excel文件
        df = pd.read_excel(file_path)
        print("✅ pandas读取成功!")
        print(f"读取到 {len(df)} 行数据")
        print("前5行数据:")
        print(df.head())
        return True
    except Exception as e:
        print(f"❌ pandas读取失败: {str(e)}")
        return False

def test_openpyxl_read(file_path):
    """测试用openpyxl读取Excel文件"""
    print("\n尝试使用openpyxl读取...")
    try:
        wb = load_workbook(filename=file_path)
        sheet = wb.active
        print("✅ openpyxl读取成功!")
        print(f"工作表名称: {sheet.title}")
        print(f"第一行数据: {[cell.value for cell in sheet[1]]}")
        return True
    except Exception as e:
        print(f"❌ openpyxl读取失败: {str(e)}")
        return False

def test_xlrd_read(file_path):
    """测试用xlrd读取旧版Excel文件"""
    print("\n尝试使用xlrd读取...")
    try:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_index(0)
        print("✅ xlrd读取成功!")
        print(f"工作表名称: {sheet.name}")
        print(f"第一行数据: {sheet.row_values(0)}")
        return True
    except Exception as e:
        print(f"❌ xlrd读取失败: {str(e)}")
        return False

def test_pyxlsb_read(file_path):
    """测试用pyxlsb读取二进制Excel文件"""
    if not file_path.lower().endswith('.xlsb'):
        print("\n跳过pyxlsb测试(非.xlsb文件)")
        return False
    
    print("\n尝试使用pyxlsb读取...")
    try:
        with pyxlsb.open_workbook(file_path) as wb:
            with wb.get_sheet(1) as sheet:
                print("✅ pyxlsb读取成功!")
                for row in sheet.rows():
                    print(f"第一行数据: {row}")
                    break
        return True
    except Exception as e:
        print(f"❌ pyxlsb读取失败: {str(e)}")
        return False

def check_file_properties(file_path):
    """检查文件基本属性"""
    print("\n检查文件属性...")
    try:
        file_size = os.path.getsize(file_path)
        print(f"文件大小: {file_size/1024:.2f} KB")
        
        if file_size == 0:
            print("⚠️ 文件大小为0，可能是损坏的文件")
        
        with open(file_path, 'rb') as f:
            header = f.read(8)
            print(f"文件头: {header}")
            
            # 检查是否是加密文件
            if header.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'):
                print("⚠️ 文件可能是加密的或受密码保护")
            
        return True
    except Exception as e:
        print(f"❌ 无法检查文件属性: {str(e)}")
        return False

def extract_chinese_from_filename(filename):
    """从文件名中提取中文部分"""
    # 提取文件名中的中文部分
    chinese_part = ''.join(re.findall('[\u4e00-\u9fa5]', filename))
    return chinese_part if chinese_part else os.path.splitext(filename)[0]

def repair_excel_file(file_path):
    """使用Excel COM修复Excel文件"""
    print("\n尝试修复Excel文件...")
    try:
        # 创建考勤数据目录（在脚本所在目录下）
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, "考勤数据")
        os.makedirs(output_dir, exist_ok=True)
        
        # 创建临时修复目录
        temp_dir = tempfile.mkdtemp()
        temp_file = os.path.join(temp_dir, "repaired.xlsx")
        
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
            
        # 尝试打开文件
        wb = excel.Workbooks.Open(file_path)
        # 强制重新计算
        excel.CalculateFull()
        # 另存为新文件
        wb.SaveAs(temp_file, FileFormat=51)
        wb.Close()
        excel.Quit()
            
        # 验证并保存到考勤数据目录
        if os.path.exists(temp_file):
            print("✅ 使用Excel COM修复成功!")
            # 从原文件名提取中文部分
            original_filename = os.path.basename(file_path)
            chinese_name = extract_chinese_from_filename(original_filename)
            # 生成输出文件名（使用中文命名）
            output_file = os.path.join(output_dir, f"{chinese_name}.xlsx")
            # 处理文件名重复情况
            counter = 1
            while os.path.exists(output_file):
                output_file = os.path.join(output_dir, f"{chinese_name}_{counter}.xlsx")
                counter += 1
                
            # 先关闭原文件句柄
            time.sleep(1)
            try:
                shutil.copy2(temp_file, output_file)
                return True
            except PermissionError:
                print("⚠️ 文件权限不足，已保存修复文件到:", temp_file)
                return True
    except Exception as e:
        print(f"❌ Excel COM修复失败: {e}")
        return False
    finally:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except:
            pass

if __name__ == "__main__":
    main()