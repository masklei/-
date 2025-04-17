import os
import pandas as pd
from datetime import datetime


def get_files():
    """获取考勤数据文件夹中的文件"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    attendance_dir = os.path.join(script_dir, "考勤数据")

    if not os.path.exists(attendance_dir):
        raise FileNotFoundError("考勤数据文件夹不存在")

    files = os.listdir(attendance_dir)
    merged_file = next((f for f in files if "合并结果" in f), None)
    overtime_file = next((f for f in files if "加班流程表" in f), None)
    leave_file = next((f for f in files if "请假流程表" in f), None)  # 新增请假流程表

    if not merged_file or not overtime_file or not leave_file:  # 新增检查
        raise FileNotFoundError("未找到必要的文件")

    return (
        os.path.join(attendance_dir, merged_file),
        os.path.join(attendance_dir, overtime_file),
        os.path.join(attendance_dir, leave_file)  # 新增返回值
    )


def process_files():
    """处理文件并进行匹配"""
    try:
        # 获取文件路径
        merged_file, overtime_file, leave_file = get_files()

        # 读取异常考勤表
        abnormal_df = pd.read_excel(merged_file)

        # 读取加班流程表，跳过前6行，第7行为列名
        overtime_df = pd.read_excel(overtime_file, header=6)

        # 读取请假流程表，跳过前6行，第7行为列名
        leave_df = pd.read_excel(leave_file, header=6)

        # 确保异常考勤表中有姓名和刷卡日期列
        required_columns_abnormal = ['姓名', '刷卡日期']
        for col in required_columns_abnormal:
            if col not in abnormal_df.columns:
                print(f"异常考勤表中缺少必要的列: {col}")
                return False

        # 确保加班流程表中有姓名、出勤日期和加班单时数列
        required_columns_overtime = ['姓名', '出勤日期', '加班单时数']
        for col in required_columns_overtime:
            if col not in overtime_df.columns:
                print(f"加班流程表中缺少必要的列: {col}")
                return False

        # 确保请假流程表中有姓名、请假开始日期、请假开始时间、请假结束时间和请假时数列
        required_columns_leave = ['姓名', '请假开始日期', '请假开始时间', '请假结束时间', '请假时数']
        for col in required_columns_leave:
            if col not in leave_df.columns:
                print(f"请假流程表中缺少必要的列: {col}")
                return False

        # 转换异常考勤表中的日期格式
        if pd.api.types.is_datetime64_any_dtype(abnormal_df['刷卡日期']):
            abnormal_df['刷卡日期'] = abnormal_df['刷卡日期'].dt.date
        else:
            try:
                abnormal_df['刷卡日期'] = pd.to_datetime(abnormal_df['刷卡日期']).dt.date
            except:
                print("错误: 无法转换异常考勤表中的日期格式")
                return False

        # 转换加班流程表中的日期格式
        overtime_df['出勤日期'] = pd.to_datetime(overtime_df['出勤日期']).dt.date

        # 转换请假流程表中的日期格式
        leave_df['请假开始日期'] = pd.to_datetime(leave_df['请假开始日期']).dt.date

        # 添加加班单时数列到异常考勤表
        abnormal_df['加班单时数'] = '无'

        # 遍历异常考勤表中的每一行，匹配加班信息
        for index, row in abnormal_df.iterrows():
            # 在加班流程表中查找匹配的记录
            matching_records = overtime_df[
                (overtime_df['姓名'] == row['姓名']) &
                (overtime_df['出勤日期'] == row['刷卡日期'])
                ]

            # 如果找到匹配的记录，更新加班单时数
            if not matching_records.empty:
                abnormal_df.at[index, '加班单时数'] = matching_records.iloc[0]['加班单时数']

        # 新增请假信息列
        abnormal_df['请假开始时间'] = None
        abnormal_df['请假结束时间'] = None
        abnormal_df['请假时数'] = None

        # 确保请假流程表中有必要的列
        required_columns_leave = ['姓名', '请假开始日期', '请假开始时间', '请假结束时间', '请假时数']
        for col in required_columns_leave:
            if col not in leave_df.columns:
                print(f"请假流程表中缺少必要的列: {col}")
                return False

        # 转换请假流程表中的日期格式
        leave_df['请假开始日期'] = pd.to_datetime(leave_df['请假开始日期']).dt.date

        # 匹配请假信息
        for index, row in abnormal_df.iterrows():
            # 在请假流程表中查找匹配的记录
            matching_records = leave_df[
                (leave_df['姓名'] == row['姓名']) &
                (leave_df['请假开始日期'] == row['刷卡日期'])
                ]

            # 如果找到匹配的记录，更新请假信息
            if not matching_records.empty:
                abnormal_df.at[index, '请假开始时间'] = matching_records.iloc[0]['请假开始时间']
                abnormal_df.at[index, '请假结束时间'] = matching_records.iloc[0]['请假结束时间']
                abnormal_df.at[index, '请假时数'] = matching_records.iloc[0]['请假时数']

        # 新增功能：提取加班时长信息
        abnormal_df['加班时长比对'] = None
        for index, row in abnormal_df.iterrows():
            if row['加班单时数'] != '无' and '加班时长' in str(row['异常描述']):
                # 提取加班时长信息
                overtime_info = [s for s in row['异常描述'].split('；') if '加班时长' in s]
                if overtime_info:
                    abnormal_df.at[index, '加班时长比对'] = overtime_info[0]

        # 生成输出文件名
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "考勤数据")
        output_file = os.path.join(output_dir, "核对版数据.xlsx")

        # 保存结果
        abnormal_df.to_excel(output_file, index=False)
        print(f"核对完成！结果已保存至: {output_file}")
        return True

    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        return False


if __name__ == "__main__":
    process_files()