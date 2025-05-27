import pandas as pd
import jionlp as jio
import re

def analyze_casualties_road_view(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 过滤出“伤害程度”为“轻伤”的数据
    related_df = df[df['伤害程度'] == '轻伤']

    # 对“身份证明号码”列去重
    related_df = related_df.drop_duplicates(subset=['身份证明号码'], keep='first')

    total_count = related_df.shape[0]
    # 解析“事故地点”并提取“道路”信息
    def extract_road(location):
        # 使用jionlp解析地址
        parsed_location = jio.parse_location(location)
        # 使用正则表达式提取道路的信息
        match = re.search(r'^(.*?[路道口])', parsed_location['detail'])
        if match:
            road = match.group(0).replace(parsed_location['city'], '')
        return road if match else None

    related_df['道路信息'] = related_df['事故地点'].apply(extract_road)

    # 统计“道路信息”出现的次数
    road_df = related_df['道路信息'].value_counts().reset_index()
    road_df.columns = ['道路信息', '受伤人数']

    # 计算占比
    def calculate_percentage(count):
        return f"{(count / total_count * 100):5.1f}%"
    
    road_df['占比'] = road_df['受伤人数'].apply(calculate_percentage)

    # 按“受伤人数”排序
    road_df = road_df.sort_values(by='受伤人数', ascending=False).reset_index(drop=True)

    # 输出结果
    print(road_df)

def get_casualties_road_view(df):
    # 过滤出“伤害程度”为“轻伤”的数据
    related_df = df[df['伤害程度'] == '轻伤']

    # 对“身份证明号码”列去重
    related_df = related_df.drop_duplicates(subset=['身份证明号码'], keep='first')

    total_count = related_df.shape[0]
    # 解析“事故地点”并提取“道路”信息
    def extract_road(location):
        # 使用jionlp解析地址
        parsed_location = jio.parse_location(location)
        # 使用正则表达式提取道路的信息
        match = re.search(r'^(.*?[路道口])', parsed_location['detail'])
        if match:
            road = match.group(0).replace(parsed_location['city'], '')
        return road if match else None

    related_df['道路信息'] = related_df['事故地点'].apply(extract_road)

    # 统计“道路信息”出现的次数
    road_df = related_df['道路信息'].value_counts().reset_index()
    road_df.columns = ['道路信息', '受伤人数']

    # 计算占比
    def calculate_percentage(count):
        return f"{(count / total_count * 100):5.1f}%"
    
    road_df['占比'] = road_df['受伤人数'].apply(calculate_percentage)
    # 按“受伤人数”排序
    road_df = road_df.sort_values(by='受伤人数', ascending=False).reset_index(drop=True)

    return road_df

if __name__ == '__main__':
    analyze_casualties_road_view('testcases/test.xlsx')
