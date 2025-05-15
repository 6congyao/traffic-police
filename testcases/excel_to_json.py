import pandas as pd
import json
import os

def process_excel_data(file_path):
    """
    处理Excel文件，将数据按来源分类并保存为json文件
    
    Args:
        file_path (str): Excel文件路径
    
    Returns:
        tuple: 包含两个字典，分别对应来源1和来源2的数据
    """
    print(f"\n开始处理文件: {file_path}", flush=True)
    print(f"文件是否存在: {os.path.exists(file_path)}", flush=True)
    
    try:
        prefix = None
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 打印Excel文件的基本信息
        print(f"\nExcel文件信息:", flush=True)
        print(f"总行数: {len(df)}", flush=True)
        print(f"列名: {df.columns.tolist()}", flush=True)
        print("\n开始处理数据...", flush=True)
        
        # 创建两个空字典用于存储不同来源的数据
        general_accidents_all = {}
        simple_accidents_all = {}
        
        # 遍历DataFrame的每一行
        for index, row in df.iterrows():
            accident_id = row['事故编号']
            source = row['来源']

            # 将行数据转换为字典并进行数据清理
            row_data = {}
            for column, value in row.items():
                # 确保所有值都是JSON可序列化的
                if pd.isna(value):  # 处理NaN和None
                    cleaned_value = None
                else:
                    cleaned_value = str(value)  # 将所有其他值转换为字符串
                row_data[column] = cleaned_value
            
            
            # 根据来源分类存储数据（按照一般事故和简易事故分类）
            if str(source).strip() == '一般事故':
                if accident_id not in general_accidents_all:
                    general_accidents_all[accident_id] = []
                general_accidents_all[accident_id].append(row_data)
            if str(source).strip() == '简易事故':
                if accident_id not in simple_accidents_all:
                    simple_accidents_all[accident_id] = []
                simple_accidents_all[accident_id].append(row_data)
        
            if prefix is None:
                py, pm, _ = row_data['事故发生时间'].split('-')
                prefix = f"{py}_{pm}_"

        return general_accidents_all, simple_accidents_all, prefix
        
    except FileNotFoundError:
        print(f"\n错误:找不到文件 '{file_path}'", flush=True)
        return {}, {}, {}
    except pd.errors.EmptyDataError:
        print(f"\n错误:Excel文件为空 '{file_path}'", flush=True)
        return {}, {}, {}
    except Exception as e:
        print(f"\n处理数据时发生错误: {str(e)}", flush=True)
        import traceback
        print("详细错误信息:", flush=True)
        print(traceback.format_exc(), flush=True)
        return {}, {}, {}

def write_to_file(general_accidents, simple_accidents, prefix):
    try:
        # 将数据写入json文件
        with open('./outputs/' + prefix + 'general_accidents_all_datasets.json', 'w', encoding='utf-8') as f:
            json.dump(general_accidents, f, ensure_ascii=False, indent=2)
            
        with open('./outputs/' + prefix + 'simple_accidents_all_datasets.json', 'w', encoding='utf-8') as f:
            json.dump(simple_accidents, f, ensure_ascii=False, indent=2)
    except Exception as json_error:
        print(f"\nJSON序列化错误: {str(json_error)}", flush=True)
        raise
        
    print(f"\n处理完成:", flush=True)
    print(f"一般事故计数: {len(general_accidents)}", flush=True)
    print(f"简易事故计数: {len(simple_accidents)}", flush=True)
    print(f"一般事故数据条数: {sum(len(v) for v in general_accidents.values())}", flush=True)
    print(f"简易事故数据条数: {sum(len(v) for v in simple_accidents.values())}", flush=True)
    
if __name__ == '__main__':
    # 使用函数处理Excel文件
    general_accidents, simple_accidents, prefix = process_excel_data('test.xlsx')

    write_to_file(general_accidents, simple_accidents, prefix)
