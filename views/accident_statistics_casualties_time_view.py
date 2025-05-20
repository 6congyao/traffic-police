import pandas as pd
from dateutil import parser


def analyze_casualties_time_view():
    """
    分析轻伤事故的时间分布
    """
    # 读取Excel文件
    df = pd.read_excel('testcases/test.xlsx')
    
    # 只保留轻伤的数据
    light_injuries = df[df['伤害程度'] == '轻伤']
    
    # 对身份证明号码去重
    light_injuries = light_injuries.drop_duplicates(subset=['身份证明号码'])
    
    print("\n=== 解析后的时间点 ===")
    parsed_times = []
    
    # 初始化24小时的统计字典
    hour_counts = {hour: 0 for hour in range(24)}
    
    for time_str in light_injuries['事故发生时间'].dropna():
        try:
            # 使用parser解析时间
            time_obj = parser.parse(str(time_str))
            print(f"原始时间：{time_str} -> 解析后：{time_obj.strftime('%Y-%m-%d %H:%M:%S')}")
            hour = time_obj.hour
            hour_counts[hour] = hour_counts.get(hour, 0) + 1
            parsed_times.append(time_obj)
        except Exception:
            continue
    
    # 统计所有时间分布
    print("\n=== 时间分布统计（按1小时间隔） ===")
    total_casualties = sum(hour_counts.values())
    if total_casualties == 0:
        print("没有找到符合条件的数据")
        return
        
    print(f"总受伤人数：{total_casualties}人")
    print("\n时间分布详情：")
    for hour in range(24):
        count = hour_counts[hour]
        percentage = (count / total_casualties) * 100 if count > 0 else 0.0
        print(f"{hour:02d} - {(hour+1):02d}: {count:3d}人 ({percentage:5.1f}%)")
    
    # 获取并输出top3时间区间
    print("\n=== Top3高发时段统计（从高到低） ===")
    top3_hours = sorted(hour_counts.items(), key=lambda x: x[1], reverse=True)[:3]
    
    for rank, (hour, count) in enumerate(top3_hours, 1):
        percentage = (count / total_casualties) * 100
        print(f"第{rank}名: {hour:02d} - {(hour+1):02d}   {count:3d}人 ({percentage:5.1f}%)")
    
    # 创建时段分布DataFrame
    print("\n=== 时段分布统计 ===")
    
    # 准备24小时的数据，每6小时一组
    data = []
    for i in range(6):  # 每组6行数据
        row = []
        for period in range(4):  # 4个时段
            base_hour = period * 6 + i
            time_str = f"{base_hour:02d}-{(base_hour+1):02d}"
            count = hour_counts[base_hour]
            row.extend([time_str, count])
        data.append(row)
    
    # 创建DataFrame
    df = pd.DataFrame(data, columns=[
        "时段1", "伤人数1", 
        "时段2", "伤人数2", 
        "时段3", "伤人数3", 
        "时段4", "伤人数4"
    ])
    
    print("\n时间分布DataFrame (按1小时间隔):")
    print(df)

def get_casualties_time_view(df):
    # 只保留轻伤的数据
    light_injuries = df[df['伤害程度'] == '轻伤']
    
    # 对身份证明号码去重
    light_injuries = light_injuries.drop_duplicates(subset=['身份证明号码'])

    parsed_times = []

    # 初始化24小时的统计字典
    hour_counts = {hour: 0 for hour in range(24)}
    
    for time_str in light_injuries['事故发生时间'].dropna():
        try:
            # 使用parser解析时间
            time_obj = parser.parse(str(time_str))
            hour = time_obj.hour
            hour_counts[hour] = hour_counts.get(hour, 0) + 1
            parsed_times.append(time_obj)
        except Exception:
            continue

    # 准备24小时的数据，每6小时一组
    data6 = []
    for i in range(6):  # 每组6行数据
        row = []
        for period in range(4):  # 4个时段
            base_hour = period * 6 + i
            time_str = f"{base_hour:02d}-{(base_hour+1):02d}"
            count = hour_counts[base_hour]
            row.extend([time_str, count])
        data6.append(row)
    
    # 创建DataFrame
    res_table_df = pd.DataFrame(data6, columns=[
        "时段1", "伤人数1", 
        "时段2", "伤人数2", 
        "时段3", "伤人数3", 
        "时段4", "伤人数4"
    ])

    # 准备24小时的数据
    data24 = []
    
    for hour in range(24):
        count = hour_counts[hour]
        data24.append({
            '时段': f"{hour:02d}-{(hour+1):02d}",
            '受伤人数': count
        })
    
    # 创建DataFrame
    res_chart_df = pd.DataFrame(data24)

    # 获取top3时间区间
    total_casualties = sum(hour_counts.values())
    top3_hours = sorted(hour_counts.items(), key=lambda x: x[1], reverse=True)[:3]
    res_top3 = []
    for rank, (hour, count) in enumerate(top3_hours, 1):
        percentage = (count / total_casualties) * 100
        res_top3.append({
            '排名': rank,
            '时段': f"{hour:02d}-{(hour+1):02d}",
            '受伤人数': count,
            '占比': f"{percentage:.1f}"
        })
    
    return res_table_df, res_chart_df, res_top3


if __name__ == "__main__":
    analyze_casualties_time_view()
