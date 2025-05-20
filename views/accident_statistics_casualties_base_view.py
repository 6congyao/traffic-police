import pandas as pd

def analyze_casualties_base_view(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)
    
    df = df[df['伤害程度'] == '轻伤']

    get_casualties_count_view(df)
    get_accident_count_view(df)


def get_casualties_count_view(df):
    # 1. 按身份证明号码去重得到新数据集
    c_set = df.drop_duplicates(subset=['身份证明号码'])
    
    # 2. 对新数据集按来源分类统计
    source_stats = c_set['来源'].value_counts()
    # 检查并添加缺失的来源
    for source in ['一般事故', '简易事故']:
        if source not in source_stats:
            source_stats[source] = 0
    print("\n按来源分类的轻伤人数统计:")
    for source, count in source_stats.items():
        print(f"{source}: {count}人")
    
    return source_stats

def get_accident_count_view(df):
    # 1. 按事故编号去重得到新数据集
    a_set = df.drop_duplicates(subset=['事故编号'])
    
    # 2. 对新数据集按来源分类统计
    source_stats = a_set['来源'].value_counts()
    # 检查并添加缺失的来源
    for source in ['一般事故', '简易事故']:
        if source not in source_stats:
            source_stats[source] = 0
    print("\n按来源分类的轻伤事故统计:")
    for source, count in source_stats.items():
        print(f"{source}: {count}起")

    return source_stats

def get_casualties_base_view(df):
    casualties_df = df[df['伤害程度'] == '轻伤']
    return get_accident_count_view(casualties_df), get_casualties_count_view(casualties_df)


if __name__ == "__main__":
    file_path = "testcases/3月样例.xlsx"
    analyze_casualties_base_view(file_path)
