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
        
        # 新增：从考勤数据文件夹获取考勤报表文件
        script_dir = os.path.dirname(os.path.abspath(__file__))
        attendance_dir = os.path.join(script_dir, "考勤数据")
        report_files = [f for f in os.listdir(attendance_dir) if "考勤报表" in f]
        if report_files:
            report_file = max(report_files, key=lambda f: os.path.getmtime(os.path.join(attendance_dir, f)))
            report = pd.read_excel(os.path.join(attendance_dir, report_file), header=6)  # 修改为header=6
            # 获取员工职务性质映射
            job_nature = dict(zip(report['姓名'], report['职务性质']))
        else:
            job_nature = {}

        # 新增：从考勤数据文件夹获取加班流程表文件
        overtime_files = [f for f in os.listdir(attendance_dir) if "加班流程表" in f]
        if overtime_files:
            overtime_file = max(overtime_files, key=lambda f: os.path.getmtime(os.path.join(attendance_dir, f)))
            overtime = pd.read_excel(os.path.join(attendance_dir, overtime_file), header=6)
            # 确保日期列是日期类型
            overtime['出勤日期'] = pd.to_datetime(overtime['出勤日期']).dt.date
            # 创建加班信息字典
            overtime_dict = {}
            for _, row in overtime.iterrows():
                key = (row['姓名'], row['出勤日期'])
                overtime_dict[key] = {
                    '加班单开始日期': row.get('加班单开始日期', ''),
                    '加班单开始时间': row.get('加班单开始时间', ''),
                    '加班单结束日期': row.get('加班单结束日期', ''),
                    '加班单结束时间': row.get('加班单结束时间', ''),
                    '加班单时数': row.get('加班单时数', '')
                }
        else:
            overtime_dict = {}
            
        # 新增：从考勤数据文件夹获取请假流程表文件
        leave_files = [f for f in os.listdir(attendance_dir) if "请假流程表" in f]
        if leave_files:
            leave_file = max(leave_files, key=lambda f: os.path.getmtime(os.path.join(attendance_dir, f)))
            leave = pd.read_excel(os.path.join(attendance_dir, leave_file), header=6)
            # 确保日期列是日期类型
            if '请假开始日期' in leave.columns:
                leave['请假开始日期'] = pd.to_datetime(leave['请假开始日期']).dt.date
            if '请假结束日期' in leave.columns:
                leave['请假结束日期'] = pd.to_datetime(leave['请假结束日期']).dt.date
            
            # 创建请假信息字典
            leave_dict = {}
            for _, row in leave.iterrows():
                # 使用姓名和请假开始日期作为键
                if '姓名' in row and '请假开始日期' in row:
                    key = (row['姓名'], row['请假开始日期'])
                    leave_dict[key] = {
                        '请假开始时间': row.get('请假开始时间', ''),
                        '请假结束时间': row.get('请假结束时间', ''),
                        '请假时数': row.get('请假时数', '')
                    }
        else:
            leave_dict = {}

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

            # 新增：检查职务性质，如果是白领则跳过该员工
            if name in job_nature and "白领" in str(job_nature[name]):
                card_detail = card_detail[card_detail['姓名'] != name]
                continue

            # 为每一行填充班别
            for idx, row in person_data.iterrows():
                key = (row['姓名'], row['刷卡日期'])
                if key in shift_dict:
                    card_detail.at[idx, '班别'] = shift_dict[key]

        # 前向填充空的班别值（使用前一天的班别）
        card_detail['班别'] = card_detail.groupby('姓名')['班别'].ffill()

        # 为刷卡明细表添加加班信息列
        for col in ['加班单开始日期', '加班单开始时间', '加班单结束日期', '加班单结束时间', '加班单时数']:
            if col not in card_detail.columns:
                card_detail[col] = ''
                
        # 为刷卡明细表添加请假信息列
        for col in ['请假开始时间', '请假结束时间', '请假时数']:
            if col not in card_detail.columns:
                card_detail[col] = ''

        # 填充加班信息
        for idx, row in card_detail.iterrows():
            key = (row['姓名'], row['刷卡日期'])
            if key in overtime_dict:
                for col in ['加班单开始日期', '加班单开始时间', '加班单结束日期', '加班单结束时间', '加班单时数']:
                    card_detail.at[idx, col] = overtime_dict[key][col]
            
            # 填充请假信息 - 通过姓名和刷卡日期匹配请假开始日期
            # 查找该员工在该日期是否有请假记录
            for leave_key, leave_info in leave_dict.items():
                leave_name, leave_date = leave_key
                if leave_name == row['姓名'] and leave_date == row['刷卡日期']:
                    for col in ['请假开始时间', '请假结束时间', '请假时数']:
                        card_detail.at[idx, col] = leave_info[col]
                    break

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
