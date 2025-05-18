import pandas as pd

def analyze_casualties_cause_view():
    # 读取Excel文件
    df = pd.read_excel('test.xlsx')
    
    # 筛选轻伤数据
    minor_injuries = df[df['伤害程度'] == '轻伤']
    
    # 对违法行为进行分布统计，忽略NaN值
    violations = minor_injuries['违法行为'].value_counts()
    total_count = violations.sum()
    
    # 计算占比
    percentages = (violations / total_count * 100).round(1)
    
    # 输出所有分布和占比
    print("\n所有轻伤事故的违法行为分布统计：")
    for violation, count in violations.items():
        percentage = percentages[violation]
        print(f"{violation}: {count}次 ({percentage}%)")
    
    # 输出top3统计
    print("\nTop 3违法行为统计：")
    for violation, count in violations.head(3).items():
        percentage = percentages[violation]
        print(f"{violation}: {count}次 ({percentage}%)")

if __name__ == "__main__":
    analyze_casualties_cause_view()
