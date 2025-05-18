import pandas as pd
import sys
from datetime import datetime

def analyze_accidents_overall_view(file_path):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        print(f"正在读取文件: {file_path}")
        print(f"文件读取成功，总行数: {len(df)}")
            
        # 检查必要的列是否存在
        required_columns = ['事故编号', '来源', '身份证明号码', '伤害程度']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Excel文件中缺少必要的列: {col}")
        
        get_accidents_overall_view(df)
        get_accidents_casualties_overall_view(df)
        get_statistics_comments(df)
                
    except FileNotFoundError:
        print(f"错误: 找不到文件 '{file_path}'", flush=True)
    except pd.errors.EmptyDataError:
        print(f"错误: Excel文件 '{file_path}' 是空的", flush=True)
    except Exception as e:
        print(f"处理数据时发生错误: {str(e)}", flush=True)
        import traceback
        print("详细错误信息:", flush=True)
        print(traceback.format_exc(), flush=True)

def get_accidents_overall_view(df):
    try:
        # 1. 事故编号统计分析
        print("\n=== 1. 事故编号统计分析 ===")
        # 按来源分组统计不同事故编号的数量
        accident_stats = df.groupby('来源')['事故编号'].nunique()
        print("\n各来源下不同事故编号的数量统计：")
        print(str(accident_stats))
        return accident_stats.to_dict()

    except Exception as e:
        print(f"处理数据时发生错误: {str(e)}", flush=True)
        import traceback
        print("详细错误信息:", flush=True)
        print(traceback.format_exc(), flush=True)

def get_accidents_casualties_overall_view(df):
    try:
        # 2. 身份证明号码统计分析
        print("\n=== 2. 身份证明号码统计分析 ===")
        # 筛选出伤害程度为"轻伤"和"死亡"的数据
        df = df[df['伤害程度'].isin(['轻伤', '死亡'])]
        # 按来源和伤害程度分组统计不同身份证明号码的数量
        id_stats = df.groupby(['来源', '伤害程度'])['身份证明号码'].nunique()
        print("\n各来源和伤害程度下不同身份证明号码的数量统计：")
        print(str(id_stats))
        return id_stats.to_dict()

    except Exception as e:
        print(f"处理数据时发生错误: {str(e)}", flush=True)
        import traceback
        print("详细错误信息:", flush=True)
        print(traceback.format_exc(), flush=True)

def get_statistics_comments(df):
    try:
        # 补充详细统计信息
        print("\n=== 补充统计信息 ===")
        # 统计来源分布
        source_counts = df['来源'].value_counts(dropna=True)
        print("\n各来源的总记录数量:")
        print(str(source_counts))
        
        # 统计伤害程度分布
        injury_counts = df['伤害程度'].value_counts(dropna=True)
        print("\n各伤害程度的总记录数量:")
        print(str(injury_counts))

    except Exception as e:
        print(f"处理数据时发生错误: {str(e)}", flush=True)
        import traceback
        print("详细错误信息:", flush=True)
        print(traceback.format_exc(), flush=True)

if __name__ == '__main__':
    analyze_accidents_overall_view('test.xlsx')
