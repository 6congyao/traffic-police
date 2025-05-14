import pandas as pd
import sys
from datetime import datetime

def analyze_accidents(file_path):
    # 准备输出文件
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f'accident_statistics_{timestamp}.txt'
    
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"正在读取文件: {file_path}\n")
            f.write(f"文件读取成功，总行数: {len(df)}\n")
            
            # 检查必要的列是否存在
            required_columns = ['事故编号', '来源', '身份证明号码', '伤害程度']
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"Excel文件中缺少必要的列: {col}")
            
            # 1. 事故编号统计分析
            f.write("\n=== 1. 事故编号统计分析 ===\n")
            # 按来源分组统计不同事故编号的数量
            accident_stats = df.groupby('来源')['事故编号'].nunique()
            f.write("\n各来源下不同事故编号的数量统计：\n")
            f.write(str(accident_stats) + "\n")
            
            # 2. 身份证明号码统计分析
            f.write("\n=== 2. 身份证明号码统计分析 ===\n")
            # 按来源和伤害程度分组统计不同身份证明号码的数量
            id_stats = df.groupby(['来源', '伤害程度'])['身份证明号码'].nunique()
            f.write("\n各来源和伤害程度下不同身份证明号码的数量统计：\n")
            f.write(str(id_stats) + "\n")
            
            # 补充详细统计信息
            f.write("\n=== 补充统计信息 ===\n")
            # 统计来源分布
            source_counts = df['来源'].value_counts(dropna=True)
            f.write("\n各来源的总记录数量:\n")
            f.write(str(source_counts) + "\n")
            
            # 统计伤害程度分布
            injury_counts = df['伤害程度'].value_counts(dropna=True)
            f.write("\n各伤害程度的总记录数量:\n")
            f.write(str(injury_counts) + "\n")
        
        print(f"统计结果已保存到文件: {output_file}", flush=True)
                
    except FileNotFoundError:
        print(f"错误: 找不到文件 '{file_path}'", flush=True)
    except pd.errors.EmptyDataError:
        print(f"错误: Excel文件 '{file_path}' 是空的", flush=True)
    except Exception as e:
        print(f"处理数据时发生错误: {str(e)}", flush=True)
        import traceback
        print("详细错误信息:", flush=True)
        print(traceback.format_exc(), flush=True)

if __name__ == '__main__':
    analyze_accidents('test.xlsx')
