import pandas as pd
import tkinter as tk
from tkinter import filedialog
import datetime
import re
from collections import defaultdict
import os
import concurrent.futures
from functools import partial


def select_file():
    """选择输入文件"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="选择考勤数据文件",
                                           filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"),
                                                      ("All files", "*.*")])
    root.destroy()
    return file_path


def is_night_shift(shift):
    """判断是否为夜班"""
    return '夜班' in str(shift)


def parse_datetime(date_str, time_str):
    """解析日期和时间字符串为datetime对象"""
    try:
        # 确保输入不是空值或数字
        if pd.isna(date_str) or pd.isna(time_str) or isinstance(date_str, (int, float)) or isinstance(time_str, (int, float)):
            return None
            
        # 转换为datetime对象
        date_obj = pd.to_datetime(date_str, errors='coerce')
        time_obj = pd.to_datetime(time_str, errors='coerce')
        
        # 检查转换是否成功
        if pd.isna(date_obj) or pd.isna(time_obj):
            return None
            
        return datetime.datetime.combine(date_obj.date(), time_obj.time())
    except Exception as e:
        print(f"解析日期时间错误: {e}, date_str={date_str}, time_str={time_str}")
        return None


def get_time_diff_minutes(time1, time2):
    """计算两个时间之间的分钟差"""
    diff = (time2 - time1).total_seconds() / 60
    return abs(diff)


def check_night_shift_anomalies(df):
    """检查夜班考勤异常"""
    # 初始化结果 DataFrame
    result_df = pd.DataFrame()  # 显式初始化为空 DataFrame

    # 按员工分组
    employee_groups = df.groupby('姓名')

    for name, group in employee_groups:
        # 筛选夜班记录
        night_shifts = group[group['班别'].apply(is_night_shift)]
        if night_shifts.empty:
            continue

        # 按日期和时间排序
        night_shifts = night_shifts.sort_values(['刷卡日期', '刷卡时间'])

        # 重新组织夜班记录（以12点为分界）
        shift_records = defaultdict(list)
        for _, row in night_shifts.iterrows():
            dt = parse_datetime(row['刷卡日期'], row['刷卡时间'])
            if not dt:
                continue

            # 忽略12:00-18:00的打卡记录
            if datetime.time(12, 0) <= dt.time() <= datetime.time(18, 0):
                continue

            # 确定记录属于哪个夜班班次（以12点为分界）
            if dt.time() >= datetime.time(12, 0):
                shift_date = dt.date()
            else:
                shift_date = dt.date() - datetime.timedelta(days=1)

            shift_records[shift_date].append({
                'datetime': dt,
                'type': row['来源'],
                'direction': row['刷卡机'],
                'row': row
            })

        # 检查每个夜班班次的异常
        for shift_date, records in shift_records.items():
            # 按时间排序
            records = sorted(records, key=lambda x: x['datetime'])

            # 处理两分钟内多次打卡的情况
            filtered_records = []
            anomalies_in_group = False
            i = 0
            while i < len(records):
                current = records[i]
                j = i + 1
                duplicates = [current]

                # 查找2分钟内的所有记录
                while j < len(records) and (records[j]['datetime'] - current['datetime']).total_seconds() <= 120:
                    duplicates.append(records[j])
                    j += 1

                # 检查是否有进出方向不一致的记录
                directions = set(r['direction'] for r in duplicates)
                if len(directions) > 1:
                    # 如果进出方向不一致，标记为异常
                    anomalies_in_group = True
                    for r in duplicates:
                        filtered_records.append(r)
                else:
                    # 如果方向一致，只保留最后一条
                    filtered_records.append(duplicates[-1])

                i = j

            records = filtered_records

            # 检查其他异常
            anomalies = []
            descriptions = []

            if records:
                # 获取夜班的上班时间(20:00)和下班时间(04:00)
                work_start_time = datetime.time(20, 0)
                work_end_time = datetime.time(4, 0)

                # 加班开始时间(5:10)和结束时间(8:10)
                overtime_start_time = datetime.time(5, 10)
                overtime_end_time = datetime.time(8, 10)

                # 检查首次进入是否在20:01前
                first_in = None
                for record in records:
                    if record['direction'] == '进':
                        first_in = record
                        break

                if first_in and first_in['datetime'].time() > datetime.time(20, 1):
                    anomalies.append("首次进入超时")
                    descriptions.append(f"首次进入时间为{first_in['datetime'].strftime('%H:%M:%S')}，超过20:01")

                # 检查最后一次出是否在隔天的4:00后
                last_out = None
                for record in reversed(records):
                    if record['direction'] == '出':
                        last_out = record
                        break
                
                # 添加对提前下班的检查
                if last_out and last_out['datetime'].time() < work_end_time:
                    anomalies.append("提前下班")
                    descriptions.append(f"最后一次出卡时间为{last_out['datetime'].strftime('%H:%M:%S')}，早于正常下班时间04:00")

                # 检查异常1：工作期间出入时间差大于15分钟
                i = 0
                while i < len(records) - 1:
                    if records[i]['direction'] == '出' and records[i + 1]['direction'] == '进':
                        time_diff = get_time_diff_minutes(records[i]['datetime'], records[i + 1]['datetime'])
                        if time_diff > 15:
                            anomalies.append("工作时间外出超过15分钟")
                            descriptions.append(f"外出时间为{records[i]['datetime'].strftime('%H:%M:%S')}，"
                                                f"再次进入时间为{records[i + 1]['datetime'].strftime('%H:%M:%S')}，"
                                                f"外出时长{int(time_diff)}分钟")
                    i += 1

                # 检查异常2：有进无出
                i = 0
                while i < len(records) - 1:
                    if records[i]['direction'] == '进' and records[i + 1]['direction'] == '进':
                        anomalies.append("有进无出")
                        descriptions.append(f"在{records[i]['datetime'].strftime('%H:%M:%S')}进入后，"
                                            f"在{records[i + 1]['datetime'].strftime('%H:%M:%S')}再次进入，无出记录")
                    i += 1

                # 检查异常3：有出无进
                i = 0
                while i < len(records) - 1:
                    if records[i]['direction'] == '出' and records[i + 1]['direction'] == '出':
                        anomalies.append("有出无进")
                        descriptions.append(f"在{records[i]['datetime'].strftime('%H:%M:%S')}外出后，"
                                            f"在{records[i + 1]['datetime'].strftime('%H:%M:%S')}再次外出，无进入记录")
                    i += 1

                # 检查异常4：首次打卡为出
                if records[0]['direction'] == '出':
                    anomalies.append("无进有出")
                    descriptions.append(f"首次打卡为出，无进入记录")

                # 检查异常5：最后一次打卡为进
                if records[-1]['direction'] == '进':
                    anomalies.append("有进无出")
                    descriptions.append(f"最后一次打卡为进，无外出记录")

                # 检查加班时长
                overtime_records = [r for r in records if r['datetime'].time() >= overtime_start_time
                                    and r['datetime'].time() <= overtime_end_time]

                if overtime_records:
                    # 计算加班时长
                    overtime_in = None
                    overtime_out = None

                    for record in overtime_records:
                        if record['direction'] == '进' and (
                                overtime_in is None or record['datetime'] < overtime_in['datetime']):
                            overtime_in = record

                    for record in reversed(overtime_records):
                        if record['direction'] == '出':
                            overtime_out = record
                            break

                    if overtime_in and overtime_out:
                        # 如果加班开始时间早于5:10，则按5:10计算
                        start_time = max(overtime_in['datetime'],
                                         datetime.datetime.combine(overtime_in['datetime'].date(), overtime_start_time))

                        overtime_minutes = get_time_diff_minutes(start_time, overtime_out['datetime'])
                        if overtime_minutes < 180:  # 3小时 = 180分钟
                            anomalies.append("加班时长不足3小时")
                            descriptions.append(f"加班时长为{int(overtime_minutes)}分钟，不足3小时")

            # 如果有异常，将所有记录添加到结果中
            if anomalies or anomalies_in_group:
                for record in records:
                    new_row = record['row'].copy()
                    new_row['异常'] = '是'

                    # 合并异常描述
                    if anomalies_in_group:
                        desc = "两分钟内进出方向不一致"
                        if anomalies:
                            desc += "；" + '；'.join(set(descriptions))
                        new_row['异常描述'] = desc
                    else:
                        new_row['异常描述'] = '；'.join(set(descriptions))

                    if result_df.empty:
                        result_df = pd.DataFrame([new_row])
                    else:
                        result_df = pd.concat([result_df, pd.DataFrame([new_row])], ignore_index=True)

    return result_df


def get_matched_file():
    """从考勤数据文件夹获取班别匹配结果文件"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    attendance_dir = os.path.join(script_dir, "考勤数据")
    
    if not os.path.exists(attendance_dir):
        raise FileNotFoundError("考勤数据文件夹不存在")
    
    files = os.listdir(attendance_dir)
    matched_files = [f for f in files if "班别匹配结果" in f]
    
    if not matched_files:
        raise FileNotFoundError("未找到班别匹配结果文件")
    
    # 取最新的文件
    matched_file = max(matched_files, key=lambda f: os.path.getmtime(os.path.join(attendance_dir, f)))
    return os.path.join(attendance_dir, matched_file)

def process_in_thread(file_path):
    """线程处理函数"""
    try:
        # 读取文件
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)

        # 处理夜班考勤异常
        result_df = check_night_shift_anomalies(df)

        # 保存结果
        output_path = os.path.join(os.path.dirname(file_path), "夜班稽查结果.xlsx")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='夜班异常')
            # 获取工作簿和工作表
            workbook = writer.book
            worksheet = writer.sheets['夜班异常']

            # 合并相同班次的异常描述单元格
            from openpyxl.utils import get_column_letter

            # 获取异常描述列的索引
            desc_col_idx = result_df.columns.get_loc('异常描述') + 1  # Excel列从1开始

            # 按姓名和班次日期分组，合并单元格
            current_name = None
            current_shift_date = None
            start_row = 2  # Excel数据从第2行开始（第1行是表头）

            for i, row in enumerate(result_df.itertuples(), start=2):
                name = getattr(row, '姓名')
                dt = parse_datetime(getattr(row, '刷卡日期'), getattr(row, '刷卡时间'))

                # 确定班次日期（以12点为分界）
                if dt and dt.time() >= datetime.time(12, 0):
                    shift_date = dt.date()
                else:
                    shift_date = dt.date() - datetime.timedelta(days=1) if dt else None

                if name != current_name or shift_date != current_shift_date:
                    # 如果是新的员工或班次，结束上一个合并区域
                    if current_name is not None and i > start_row:
                        cell_range = f"{get_column_letter(desc_col_idx)}{start_row}:{get_column_letter(desc_col_idx)}{i - 1}"
                        worksheet.merge_cells(cell_range)

                    # 开始新的合并区域
                    current_name = name
                    current_shift_date = shift_date
                    start_row = i

            # 处理最后一组
            if current_name is not None and start_row < i:
                cell_range = f"{get_column_letter(desc_col_idx)}{start_row}:{get_column_letter(desc_col_idx)}{i}"
                worksheet.merge_cells(cell_range)

        print(f"处理完成，结果已保存至: {output_path}")
        print(f"共发现 {len(result_df)} 条异常记录")
    except Exception as e:
        print(f"处理文件时出错: {str(e)}")

if __name__ == "__main__":
    try:
        # 自动获取文件
        file_path = get_matched_file()
        
        # 使用线程池处理
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future = executor.submit(process_in_thread, file_path)
            future.result()  # 等待线程完成
            
    except Exception as e:
        print(f"程序运行出错：{str(e)}")