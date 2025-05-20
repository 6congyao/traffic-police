import pandas as pd

def analyze_casualties_cause_view():
    # 读取Excel文件
    df = pd.read_excel('test.xlsx')
    
    total_count = df.shape[0]
    total_violation_series = df['违法行为'].value_counts()
    top3_total_violations = [
        {"违法行为": index, "数量": count, "占比": f"{((count / total_count) * 100):.1f}"} 
        for index, count in total_violation_series.head(3).items()
    ]

     # 过滤出“轻伤”和“死亡”的数据
    filtered_df = df[df['伤害程度'].isin(['轻伤', '死亡'])]

    # 统计过滤后的数据数量
    filtered_count = filtered_df.shape[0]

    # 针对“违法行为”列，统计相同值的数量
    filtered_violation_series = filtered_df['违法行为'].value_counts()

    top3_filtered_violations = [
        {"违法行为": index, "数量": count} 
        for index, count in filtered_violation_series.head(3).items()
    ]
    
    print(f"违法行为总数量: {total_count}")
    print("违法行为前三名:")
    print(top3_total_violations)
    print(f"伤亡违法行为的总数量: {filtered_count}")
    print("伤亡违法行为前三名:")
    print(top3_filtered_violations)

def get_casualties_cause_view(df):
    total_count = df.shape[0]
    total_violation_series = df['违法行为'].value_counts()
    top3_total_violations = [
        {"违法行为": index, "数量": count, "占比": f"{((count / total_count) * 100):.1f}"} 
        for index, count in total_violation_series.head(3).items()
    ]

     # 过滤出“轻伤”和“死亡”的数据
    filtered_df = df[df['伤害程度'].isin(['轻伤', '死亡'])]

    # 统计过滤后的数据数量
    filtered_count = filtered_df.shape[0]

    # 针对“违法行为”列，统计相同值的数量
    filtered_violation_series = filtered_df['违法行为'].value_counts()

    top3_filtered_violations = [
        {"违法行为": index, "数量": count} 
        for index, count in filtered_violation_series.head(3).items()
    ]

    return top3_total_violations, top3_filtered_violations


if __name__ == "__main__":
    analyze_casualties_cause_view()
