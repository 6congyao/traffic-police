import pandas as pd

def analyze_casualties_team_view():
    # 读取Excel文件
    df = pd.read_excel('test.xlsx')
    
    # 筛选出伤害程度为"轻伤"和"死亡"的数据
    df = df[df['伤害程度'].isin(['轻伤', '死亡'])]
    
    # 以身份证明号码为key进行去重，保留第一条记录
    df = df.drop_duplicates(subset=['身份证明号码'], keep='first')
    
    # 按照来源和伤害程度分组
    grouped = df.groupby(['来源', '伤害程度'])
    
    print("\n=== 数据统计结果 ===")
    # 打印每个分组的数据
    for (source, injury_type), group in grouped:
        print(f"\n来源：{source} - 伤害程度：{injury_type}（共{len(group)}条）:")
        print("-" * 50)
        print(group)
        print("-" * 50)
        
        # 统计所属中队出现的次数
        team_counts = group['所属中队'].value_counts()
        print(f"\n当前类别所属中队出现次数统计：")
        print("-" * 30)
        for team, count in team_counts.items():
            if pd.notna(team):  # 排除空值
                # 将科学计数法转换为普通数字字符串
                team_str = f"{team:.0f}"
                print(f"中队编号 {team_str}: {count}次")
        print("-" * 30)

if __name__ == '__main__':
    try:
        analyze_casualties_team_view()
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
