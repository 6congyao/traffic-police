import pandas as pd
from dateutil import parser


def analyze_decedents_time_view():
    """
    分析轻伤事故的时间分布
    """
    # 读取Excel文件
    df = pd.read_excel('testcases/test.xlsx')
    
    # 只保留轻伤的数据
    decedents = df[df['伤害程度'] == '死亡']
    
    # 对身份证明号码去重
    decedents = decedents.drop_duplicates(subset=['身份证明号码'])
    
    print("\n=== 解析后的时间点 ===")
    parsed_times = []
    
    # 初始化24小时的统计字典
    hour_counts = {hour: 0 for hour in range(24)}
    
    for time_str in decedents['事故发生时间'].dropna():
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
    total_decedents = sum(hour_counts.values())
    if total_decedents == 0:
        print("没有找到符合条件的数据")
        return
        
    print(f"总死亡人数：{total_decedents}人")
    print("\n时间分布详情：")
    for hour in range(24):
        count = hour_counts[hour]
        percentage = (count / total_decedents) * 100 if count > 0 else 0.0
        print(f"{hour:02d}:00 - {hour:02d}:59: {count:3d}人 ({percentage:5.1f}%)")
    
    # 获取并输出top1时间区间
    print("\n=== Top1高发时段统计（从高到低） ===")
    top1_hours = sorted(hour_counts.items(), key=lambda x: x[1], reverse=True)[:1]
    
    for rank, (hour, count) in enumerate(top1_hours, 1):
        percentage = (count / total_decedents) * 100
        print(f"第{rank}名: {hour:02d}:00 - {hour:02d}:59   {count:3d}人 ({percentage:5.1f}%)")

if __name__ == "__main__":
    analyze_decedents_time_view()
