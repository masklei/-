import pandas as pd
import os
import time
from datetime import datetime

def process_data(card_detail_file, attendance_file):
    """处理数据并生成新的Excel文件"""
    try:
        # 读取刷卡明细表，从第7行开始（索引为6）
        card_detail = pd.read_excel(card_detail_file, header=6)

        # 读取上下班打卡明细，从第7行开始（索引为6）
        attendance = pd.read_excel(attendance_file, header=6)

        # 确保日期列是日期类型
        card_detail['刷卡日期'] = pd.to_datetime(card_detail['刷卡日期']).dt.date
        attendance['出勤日期'] = pd.to_datetime(attendance['出勤日期']).dt.date

        # 创建一个字典，键为(姓名, 出勤日期)，值为班别
        shift_dict = dict(zip(zip(attendance['姓名'], attendance['出勤日期']), attendance['班别']))

        # 为刷卡明细表添加班别列
        card_detail['班别'] = None

        # 按姓名和刷卡日期排序
        card_detail = card_detail.sort_values(by=['姓名', '刷卡日期'])

        # 为每个人填充班别信息
        for name in card_detail['姓名'].unique():
            person_data = card_detail[card_detail['姓名'] == name].copy()

            # 为每一行填充班别
            for idx, row in person_data.iterrows():
                key = (row['姓名'], row['刷卡日期'])
                if key in shift_dict:
                    card_detail.at[idx, '班别'] = shift_dict[key]

        # 前向填充空的班别值（使用前一天的班别）
        card_detail['班别'] = card_detail.groupby('姓名')['班别'].ffill()

        # 获取当前时间作为文件名的一部分
        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        
        # 确保考勤数据目录存在
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "考勤数据")
        os.makedirs(output_dir, exist_ok=True)
        
        # 输出文件路径
        output_file = os.path.join(output_dir, "班别匹配结果.xlsx")

        # 保存结果
        card_detail.to_excel(output_file, index=False)

        return True
    except Exception as e:
        print(f"处理数据时出错：{str(e)}")
        return False

def get_files_from_attendance_folder():
    """从考勤数据文件夹获取刷卡明细和打卡明细文件"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    attendance_dir = os.path.join(script_dir, "考勤数据")
    
    if not os.path.exists(attendance_dir):
        raise FileNotFoundError("考勤数据文件夹不存在")
    
    files = os.listdir(attendance_dir)
    card_detail_files = [f for f in files if "刷卡明细" in f]
    attendance_files = [f for f in files if "打卡明细" in f]
    
    if not card_detail_files:
        raise FileNotFoundError("未找到刷卡明细文件")
    if not attendance_files:
        raise FileNotFoundError("未找到打卡明细文件")
    
    # 取最新的文件
    card_detail_file = max(card_detail_files, key=lambda f: os.path.getmtime(os.path.join(attendance_dir, f)))
    attendance_file = max(attendance_files, key=lambda f: os.path.getmtime(os.path.join(attendance_dir, f)))
    
    return (
        os.path.join(attendance_dir, card_detail_file),
        os.path.join(attendance_dir, attendance_file)
    )

if __name__ == "__main__":
    try:
        card_detail_file, attendance_file = get_files_from_attendance_folder()
        if process_data(card_detail_file, attendance_file):
            print("处理完成")
        else:
            print("处理失败")
    except Exception as e:
        print(f"程序运行出错：{str(e)}")