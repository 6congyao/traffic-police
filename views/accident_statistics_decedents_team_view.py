import pandas as pd
import re

def analyze_accidents_decedents_team_view(file_path):
    """
    分析死亡案例数据，按来源和伤害程度分类统计所属中队的出现次数
    
    Args:
        file_path: Excel文件路径
    Returns:
        dict: 包含分类统计结果的字典
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 过滤出死亡案例
        decedents_df = df[df['伤害程度'] == '死亡']
        
        # 根据事故编号去重，保留第一次出现的记录
        unique_decedents = decedents_df.drop_duplicates(subset=['事故编号'], keep='first')
        
        # 统计所属中队的出现次数，在统计时将数值转换为字符串
        team_counts = unique_decedents['所属中队'].apply(
            lambda x: re.search(r'大队(.+)', x).group(1) if pd.notnull(x) and re.search(r'大队(.+)', x) else None
        ).value_counts().to_dict()
            
        return team_counts
    except Exception as e:
        print(f"处理数据时出错: {str(e)}")
        return {}
    
def get_accidents_decedents_team_view(df):
    # 过滤出死亡案例
    decedents_df = df[df['伤害程度'] == '死亡']
    
    # 根据事故编号去重，保留第一次出现的记录
    unique_decedents = decedents_df.drop_duplicates(subset=['事故编号'], keep='first')
    
    # 统计所属中队的出现次数，在统计时将数值转换为字符串
    team_counts = unique_decedents['所属中队'].apply(
        lambda x: re.search(r'大队(.+)', x).group(1) if pd.notnull(x) and re.search(r'大队(.+)', x) else None
    ).value_counts().to_dict()
    
    df_results = pd.DataFrame(list(team_counts.items()), columns=['所属中队', '亡人事故次数'])
    df_results = df_results.sort_values(by='亡人事故次数', ascending=False)
    # 重置索引并隐藏索引列
    df_results.index = range(1, len(df_results) + 1)

    return df_results
    

if __name__ == '__main__':
    # 示例使用
    file_path = 'testcases/test.xlsx'
    results = analyze_accidents_decedents_team_view(file_path)
    # 将结果转换为DataFrame并按次数从高到低排序
    print("\n所属中队出现次数统计（按次数从高到低排序）：")
    print("-" * 50)
    # 创建DataFrame并排序
    df_results = pd.DataFrame(list(results.items()), columns=['所属中队', '亡人事故次数'])
    df_results = df_results.sort_values(by='亡人事故次数', ascending=False)
    # 重置索引并隐藏索引列
    df_results.index = range(1, len(df_results) + 1)
    # 打印表格
    print(df_results.to_string(index=True))
