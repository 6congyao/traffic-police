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
    print("\n按来源分类的轻伤人数统计:")
    for source, count in source_stats.items():
        print(f"{source}: {count}人")

def get_accident_count_view(df):
    # 1. 按事故编号去重得到新数据集
    a_set = df.drop_duplicates(subset=['事故编号'])
    
    # 2. 对新数据集按来源分类统计
    source_stats = a_set['来源'].value_counts()
    print("\n按来源分类的轻伤事故统计:")
    for source, count in source_stats.items():
        print(f"{source}: {count}起")

if __name__ == "__main__":
    file_path = "test.xlsx"
    analyze_casualties_base_view(file_path)
