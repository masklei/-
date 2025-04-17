import pandas as pd
import re
import os  # 添加os模块导入
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import concurrent.futures
from functools import partial

def is_baiban(shift):
    """判断班次是否为白班"""
    # 修改白班判断逻辑，增加对日期格式的处理
    shift_str = str(shift)
    if '白班' in shift_str:
        return True
    # 尝试判断是否为日期格式，如果是日期可能是白班
    try:
        # 如果能转换为日期，则认为可能是白班
        datetime.strptime(shift_str, '%Y-%m-%d')
        return True
    except ValueError:
        pass
    return False

def parse_time(time_str):
    """解析时间字符串为datetime对象"""
    if isinstance(time_str, str):
        try:
            return datetime.strptime(time_str, '%H:%M:%S')
        except ValueError:
            try:
                return datetime.strptime(time_str, '%H:%M')
            except ValueError:
                return None
    return None

def time_diff_minutes(time1, time2):
    """计算两个时间之间的分钟差"""
    if time1 and time2:
        diff = (time2 - time1).total_seconds() / 60
        return abs(diff)
    return 0

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

def process_attendance_data(file_path, output_file):
    """处理考勤数据并检测异常"""
    # 读取数据
    try:
        df = pd.read_excel(file_path)
        
        # 检查必要的列是否存在
        required_columns = ['姓名', '刷卡日期', '班别', '刷卡时间', '来源']
        for col in required_columns:
            if col not in df.columns:
                return None
        
        # 添加进出列（如果不存在）
        if '进出' not in df.columns:
            if '刷卡机' in df.columns:
                df['进出'] = df['刷卡机'].apply(lambda x: '进' if '进' in str(x) else ('出' if '出' in str(x) else '未知'))
            else:
                return None
        
        # 确保所需列存在
        for col in ['单位', '部门', '部门CXO-2', '工号']:
            if col not in df.columns:
                df[col] = ''
        
        # 转换日期和时间格式
        df['刷卡时间'] = pd.to_datetime(df['刷卡时间'], errors='coerce')
        df['时间'] = df['刷卡时间'].dt.time
        df['刷卡日期'] = pd.to_datetime(df['刷卡日期'], errors='coerce')
        
    except Exception as e:
        return None
    
    # 定义白班的时间界限
    work_start_time = datetime.strptime('08:00', '%H:%M').time()
    work_end_time = datetime.strptime('16:40', '%H:%M').time()
    overtime_start_time = datetime.strptime('17:10', '%H:%M').time()
    overtime_end_time = datetime.strptime('20:10', '%H:%M').time()
    
    # 按姓名和日期分组处理数据
    result_records = []
    grouped = df.groupby(['姓名', '刷卡日期'])
    
    for (name, date), group in grouped:
        # 检查是否为白班
        if not any(is_baiban(shift) for shift in group['班别'].unique()):
            continue
        
        # 按时间排序并处理重复打卡
        group = group.sort_values('刷卡时间')
        filtered_records = []
        i = 0
        while i < len(group):
            current_record = group.iloc[i]
            current_in_out = current_record['进出']
            j = i + 1
            
            while (j < len(group) and 
                   (group.iloc[j]['刷卡时间'] - current_record['刷卡时间']).total_seconds() <= 120 and
                   group.iloc[j]['进出'] == current_in_out):
                current_record = group.iloc[j]
                j += 1
            
            filtered_records.append(current_record)
            i = j if j > i else i + 1
        
        filtered_group = pd.DataFrame(filtered_records)
        
        # 初始化异常标记
        has_anomaly = False
        anomaly_desc = []
        anomaly_cells = []  # 存储需要标黄的单元格位置
        
        # 新增异常检测：首次打卡应该是进入
        if not filtered_group.empty:
            first_record = filtered_group.iloc[0]
            if first_record['进出'] == '出':
                has_anomaly = True
                anomaly_desc.append("首次打卡为出")
                anomaly_cells.append(('进出', 0))  # 标记首行的进出列
        
        # 检查首次进入是否在8:01前
        first_in = filtered_group[filtered_group['进出'] == '进'].iloc[0] if not filtered_group[filtered_group['进出'] == '进'].empty else None
        if first_in is not None and first_in['时间'] > datetime.strptime('08:01', '%H:%M').time():
            has_anomaly = True
            anomaly_desc.append(f"首次进入迟到({first_in['时间']})")
            # 找到对应行索引
            first_in_idx = filtered_group[filtered_group['进出'] == '进'].index[0]
            anomaly_cells.append(('时间', first_in_idx))
        
        # 检查最后一次出是否在16:40后
        last_out = filtered_group[filtered_group['进出'] == '出'].iloc[-1] if not filtered_group[filtered_group['进出'] == '出'].empty else None
        if last_out is not None and last_out['时间'] < work_end_time:
            has_anomaly = True
            anomaly_desc.append(f"最后一次出早退({last_out['时间']})")
            last_out_idx = filtered_group[filtered_group['进出'] == '出'].index[-1]
            anomaly_cells.append(('时间', last_out_idx))
        
        # 异常1：工作时间外出超过15分钟
        in_out_pairs = []
        current_in = None
        
        for idx, record in filtered_group.iterrows():
            time = record['时间']
            in_out = record['进出']
            
            # 工作时间内的记录
            if work_start_time <= time <= work_end_time:
                if in_out == '进':
                    current_in = (idx, record)
                elif in_out == '出' and current_in is not None:
                    in_out_pairs.append((current_in, (idx, record)))
                    current_in = None
        
        for (in_idx, in_record), (out_idx, out_record) in in_out_pairs:
            time_diff = time_diff_minutes(
                datetime.combine(date, in_record['时间']),
                datetime.combine(date, out_record['时间'])
            )
            if time_diff > 15:
                has_anomaly = True
                anomaly_desc.append(f"工作时间外出时间超过15分钟({in_record['时间']}-{out_record['时间']}, {time_diff:.0f}分钟)")
                anomaly_cells.append(('时间', in_idx))
                anomaly_cells.append(('时间', out_idx))
        
        # 异常2：有进无出卡记录
        in_records = filtered_group[filtered_group['进出'] == '进']
        out_records = filtered_group[filtered_group['进出'] == '出']
        
        if len(in_records) > len(out_records):
            has_anomaly = True
            anomaly_desc.append("有进无出卡记录")
            for idx in in_records.index:
                anomaly_cells.append(('进出', idx))
        
        # 异常3：有出无进记录
        in_times = in_records['刷卡时间'].tolist()
        out_times = out_records['刷卡时间'].tolist()
        in_indices = in_records.index.tolist()
        out_indices = out_records.index.tolist()
        
        has_out_no_in = False
        for i in range(len(out_times)):
            if i == 0 and (len(in_times) == 0 or out_times[i] < in_times[0]):
                has_anomaly = True
                has_out_no_in = True
                anomaly_cells.append(('进出', out_indices[i]))
                break
            if i > 0 and i < len(in_times) and out_times[i] < in_times[i]:
                has_anomaly = True
                has_out_no_in = True
                anomaly_cells.append(('进出', out_indices[i]))
                break
        
        if has_out_no_in:
            anomaly_desc.append("有出无进记录")
        
        # 异常4：最后一次打卡是出，但之前没有对应的进
        if not out_records.empty and not in_records.empty:
            last_out_time = out_records.iloc[-1]['刷卡时间']
            last_in_time = in_records.iloc[-1]['刷卡时间']
            last_out_idx = out_records.index[-1]
            
            if last_out_time > last_in_time:
                # 检查是否有对应的进入记录
                has_corresponding_in = False
                for in_time, in_idx in zip(in_times, in_indices):
                    if in_time < last_out_time and (len([t for t in out_times if in_time < t < last_out_time]) == 0):
                        has_corresponding_in = True
                        break
                
                if not has_corresponding_in:
                    has_anomaly = True
                    anomaly_desc.append("有出无进记录")
                    anomaly_cells.append(('进出', last_out_idx))
        
        # 检查加班时长
        overtime_records = filtered_group[filtered_group['时间'] >= overtime_start_time]
        if not overtime_records.empty:
            # 找到加班开始时间（不早于17:10）
            overtime_start = None
            overtime_start_idx = None
            for idx, record in overtime_records.iterrows():
                if record['进出'] == '进':
                    overtime_start = max(record['时间'], overtime_start_time)
                    overtime_start_idx = idx
                    break
            
            if overtime_start is None:
                overtime_start = overtime_start_time
            
            # 找到加班结束时间
            overtime_end = None
            overtime_end_idx = None
            for idx, record in overtime_records.sort_values('刷卡时间', ascending=False).iterrows():
                if record['进出'] == '出':
                    overtime_end = record['时间']
                    overtime_end_idx = idx
                    break
            
            if overtime_end is not None:
                # 计算加班时长（小时）
                overtime_start_dt = datetime.combine(date, overtime_start)
                overtime_end_dt = datetime.combine(date, overtime_end)
                
                # 如果结束时间小于开始时间，说明跨天了
                if overtime_end_dt < overtime_start_dt:
                    overtime_end_dt += timedelta(days=1)
                
                overtime_hours = (overtime_end_dt - overtime_start_dt).total_seconds() / 3600
                
                if overtime_hours < 3:
                    has_anomaly = True
                    anomaly_desc.append(f"加班时长不足3小时({overtime_hours:.2f}小时)")
                    if overtime_start_idx is not None:
                        anomaly_cells.append(('时间', overtime_start_idx))
                    if overtime_end_idx is not None:
                        anomaly_cells.append(('时间', overtime_end_idx))
        
        # 如果有异常，将所有记录添加到结果中
        if has_anomaly:
            # 为每条记录添加异常标记和描述
            group_records = []
            for idx, record in group.iterrows():
                record_dict = record.to_dict()
                record_dict['异常'] = '是'
                record_dict['异常描述'] = '；'.join(anomaly_desc)
                record_dict['异常单元格'] = [(col, idx) for col, i in anomaly_cells if i == idx]
                group_records.append(record_dict)
            
            # 添加分组信息，用于后续合并单元格
            for record in group_records:
                record['分组'] = f"{name}_{date.strftime('%Y-%m-%d')}"
                result_records.append(record)
    
    # 创建结果DataFrame
    if result_records:
        # 使用原始打卡时间替换处理后的打卡时间
        for record in result_records:
            if '原始打卡时间' in record:
                record['刷卡时间'] = record['原始打卡时间']
            if '异常单元格' in record:
                del record['异常单元格']
            if '分组' in record:
                del record['分组']
        
        result_df = pd.DataFrame(result_records)
        
        # 添加异常和异常描述列
        if '异常' not in result_df.columns:
            result_df['异常'] = '是'
        if '异常描述' not in result_df.columns:
            result_df['异常描述'] = ''
        
        # 删除不需要的列
        if '时间' in result_df.columns:
            result_df = result_df.drop(columns=['时间'])
        if '原始打卡时间' in result_df.columns:
            result_df = result_df.drop(columns=['原始打卡时间'])
        
        # 确保输出列名按要求排序
        output_columns = ['单位', '部门', '部门CXO-2', '工号', '姓名', '刷卡日期', '刷卡时间', '来源', '刷卡机', '班别', '异常', '异常描述']
        
        # 确保所有需要的列都存在
        for col in output_columns:
            if col not in result_df.columns:
                result_df[col] = ''
        
        # 按指定顺序排列列
        result_df = result_df[output_columns]
        
        # 保存结果
        output_file = os.path.join(os.path.dirname(file_path), "白班稽查结果.xlsx")
        result_df.to_excel(output_file, index=False)
        
        # 使用openpyxl标记异常单元格并合并异常描述单元格
        wb = load_workbook(output_file)
        ws = wb.active
        
        # 定义黄色填充
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        
        # 获取列名到列索引的映射
        column_indices = {}
        for i, column in enumerate(ws[1]):
            column_indices[column.value] = i + 1  # Excel列从1开始
        
        # 按分组合并异常描述单元格
        groups = {}
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            name = row[column_indices.get('姓名', 0) - 1]
            date_str = row[column_indices.get('刷卡日期', 0) - 1]
            if isinstance(date_str, datetime):
                date_str = date_str.strftime('%Y-%m-%d')
            group_key = f"{name}_{date_str}"
            
            if group_key not in groups:
                groups[group_key] = {'start': i, 'end': i}
            else:
                groups[group_key]['end'] = i
        
        # 合并单元格并标记异常
        for group_key, range_info in groups.items():
            start_row = range_info['start']
            end_row = range_info['end']
            
            # 如果有多行，合并异常描述单元格
            if start_row != end_row and '异常描述' in column_indices:
                desc_col = column_indices['异常描述']
                ws.merge_cells(start_row=start_row, start_column=desc_col, 
                              end_row=end_row, end_column=desc_col)
        
        # 标记异常单元格
        for i, row in enumerate(result_records, start=2):
            if '异常单元格' in row:
                for col_name, _ in row['异常单元格']:
                    if col_name in column_indices:
                        col_idx = column_indices[col_name]
                        cell = ws.cell(row=i, column=col_idx)
                        cell.fill = yellow_fill
        
        # 保存修改后的Excel文件
        wb.save(output_file)
        print(f"异常数据已保存到: {output_file}")
        return result_df
    else:
        print("未发现异常数据")
        return None

def process_in_thread(file_path):
    """线程处理函数"""
    try:
        output_file = os.path.join(os.path.dirname(file_path), "白班稽查结果.xlsx")
        process_attendance_data(file_path, output_file)
        print(f"处理完成，结果已保存到: {output_file}")
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {str(e)}")

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