import os
import pandas as pd

def get_files():
    """获取考勤数据文件夹中的合并结果文件"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    attendance_dir = os.path.join(script_dir, "考勤数据")

    if not os.path.exists(attendance_dir):
        raise FileNotFoundError("考勤数据文件夹不存在")

    files = os.listdir(attendance_dir)
    merged_file = next((f for f in files if "合并结果" in f), None)

    if not merged_file:
        raise FileNotFoundError("未找到合并结果文件")

    return os.path.join(attendance_dir, merged_file)

def process_files():
    """查找合并结果文件并重命名为核对版数据.xlsx"""
    try:
        merged_file = get_files()
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "考勤数据")
        output_file = os.path.join(output_dir, "核对版数据.xlsx")
        
        os.rename(merged_file, output_file)
        print(f"文件已重命名为: {output_file}")
        return True
        
    except Exception as e:
        print(f"重命名文件时出错: {str(e)}")
        return False

if __name__ == "__main__":
    process_files()
