import pandas as pd
import jionlp as jio
from datetime import datetime
from collections import defaultdict

def analyze_casualties_age_view():
    # Read the Excel file
    df = pd.read_excel('test.xlsx')
    
    # Filter for light injuries (轻伤)
    df = df[df['伤害程度'] == '轻伤']

    # 获取去重后的数据
    c_set = df.drop_duplicates(subset=['身份证明号码'])
    
    # Get current year
    current_year = datetime.now().year
    
    # 初始化年龄分布字典，包含所有可能的区间
    age_distribution = defaultdict(int)
    # 初始化所有5年间隔的区间（0-64岁）
    for i in range(0, 65, 5):
        age_distribution[i] = 0
    # 添加65岁及以上的区间
    age_distribution[65] = 0
    
    # Process each valid ID number
    for idx, row in c_set.iterrows():
        id_number = row['身份证明号码']
        
        # Skip NaN values
        if pd.isna(id_number):
            continue
            
        # Convert to string and extract birth year using jionlp
        try:
            id_str = str(int(id_number))  # Convert to int first to remove any decimal points
            id_info = jio.parse_id_card(id_str)
            if id_info and 'birth_year' in id_info:
                birth_year = id_info['birth_year']
                age = current_year - int(birth_year)
                print(f"伤者身份证：{id_str}  伤者年龄：{age}")
                # 对65岁以上的单独处理
                if age >= 65:
                    age_distribution[65] += 1
                else:
                    # 其他年龄按5年间隔分组
                    age_group = (age // 5) * 5
                    age_distribution[age_group] += 1
        except (ValueError, TypeError):
            continue
    
    # Calculate total number of valid records
    total_count = sum(age_distribution.values())
    
    if total_count == 0:
        print("没有找到有效的数据记录")
        return

    # Print age distribution with percentages
    print("\n年龄分布统计 (5年为间隔):")
    print("-" * 50)
    print(f"{'年龄区间':<15}{'人数':<10}{'占比':<10}")
    print("-" * 50)
    
    # Sort age groups for display
    for age_group in sorted(age_distribution.keys()):
        count = age_distribution[age_group]
        if count >= 0:  # 只显示有数据的年龄段
            percentage = (count / total_count) * 100
            if age_group == 65:
                print(f"65岁及以上{count:>8}人{percentage:>10.1f}%")
            else:
                print(f"{age_group}-{age_group+4}岁{count:>8}人{percentage:>10.1f}%")

    # Print top 3 age groups by count
    print("\nTop 3 年龄区间统计:")
    print("-" * 50)
    print(f"{'排名':<6}{'年龄区间':<15}{'人数':<10}{'占比':<10}")
    print("-" * 50)
    
    # Sort by count in descending order
    sorted_groups = sorted(
        [(age, count) for age, count in age_distribution.items() if count > 0],
        key=lambda x: x[1],
        reverse=True
    )
    
    for i, (age_group, count) in enumerate(sorted_groups[:3], 1):
        percentage = (count / total_count) * 100
        if age_group == 65:
            print(f"第{i}位  {'65岁及以上':<13}{count:>8}人{percentage:>10.1f}%")
        else:
            print(f"第{i}位  {f'{age_group}-{age_group+4}岁':<13}{count:>8}人{percentage:>10.1f}%")

    # Print total
    print("-" * 50)
    print(f"总计:{total_count:>11}人{100:>10.1f}%")

if __name__ == '__main__':
    analyze_casualties_age_view()
