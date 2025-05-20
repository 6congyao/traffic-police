import pandas as pd

def analyze_casualties_team_view():
    """
    分析轻伤数据，统计各中队的一般事故和简易事故伤人情况
    """
    # 设置pandas显示选项
    pd.set_option('display.float_format', lambda x: '%.0f' % x)
    try:
        # 读取Excel文件
        df = pd.read_excel('testcases/test.xlsx')
        
        # 筛选轻伤数据
        light_injuries = df[df['伤害程度'] == '轻伤']
        
        # 去除所属中队为空的记录
        light_injuries = light_injuries.dropna(subset=['所属中队'])
        
        # 分别统计一般事故和简易事故的伤人情况
        def count_injuries(data, source):
            result = data[data['来源'] == source].groupby('所属中队')['身份证明号码'].nunique()
            return result
            
        # 获取一般事故和简易事故的统计结果
        general_injuries = count_injuries(light_injuries, '一般事故')
        simple_injuries = count_injuries(light_injuries, '简易事故')
        
        # 合并数据
        result = pd.DataFrame({
            '本月伤人情况': 0,
            '总环比': 0,
            '总同比': 0,
            '简易事故伤人情况': simple_injuries,
            '简易环比': 0,
            '简易同比': 0,
            '一般事故伤人情况': general_injuries
            
        }).fillna(0)
        
        # 计算总计并排序
        result['本月伤人情况'] = result['一般事故伤人情况'] + result['简易事故伤人情况']
        result = result.sort_values('本月伤人情况', ascending=False)
        
        # 重置索引，将所属中队变为列
        result = result.reset_index()
        
        # 格式化所属中队编号
        result['所属中队'] = result['所属中队'].apply(lambda x: str(x).split('.')[0])
        
        # 确保所有数值列为整数
        for col in ['本月伤人情况','一般事故伤人情况','简易事故伤人情况']:
            result[col] = result[col].astype(int)
        
        # 输出结果
        print('-' * 80)
        print(result.to_string(index=False, justify='center'))
        print('-' * 80)
                
    except Exception as e:
        print(f"处理数据时发生错误: {str(e)}")

def get_casualties_team_view(df):
    # 筛选轻伤数据
    light_injuries = df[df['伤害程度'] == '轻伤']
    
    # 去除所属中队为空的记录
    light_injuries = light_injuries.dropna(subset=['所属中队'])
    
    # 分别统计一般事故和简易事故的伤人情况
    def count_injuries(data, source):
        result = data[data['来源'] == source].groupby('所属中队')['身份证明号码'].nunique()
        return result
        
    # 获取一般事故和简易事故的统计结果
    general_injuries = count_injuries(light_injuries, '一般事故')
    simple_injuries = count_injuries(light_injuries, '简易事故')
    
    # 合并数据
    result = pd.DataFrame({
        '本月伤人情况': 0,
        '总环比': 0,
        '总同比': 0,
        '简易事故伤人情况': simple_injuries,
        '简易环比': 0,
        '简易同比': 0,
        '一般事故伤人情况': general_injuries
        
    }).fillna(0)
    
    # 计算总计并排序
    result['本月伤人情况'] = result['一般事故伤人情况'] + result['简易事故伤人情况']
    result = result.sort_values('本月伤人情况', ascending=False)
    
    # 重置索引，将所属中队变为列
    result = result.reset_index()
    
    # 格式化所属中队编号
    result['所属中队'] = result['所属中队'].apply(lambda x: str(x).split('.')[0])
    # 确保所有数值列为整数
    for col in ['本月伤人情况','一般事故伤人情况','简易事故伤人情况']:
        result[col] = result[col].astype(int)
    return result

if __name__ == '__main__':
    analyze_casualties_team_view()
